param()

# 更新確認・提出用PDF出力の高速化で、結果が変わっていないことを server.ps1 の実物の関数で確かめる。
# Java・Office は使わない。
#
# 守ること:
# - 提出用PDFの指紋（「最新」かどうかの判定値）と、manifest・物理ページ・ブロッカー・指紋の入力が、
#   変更前の実装（下の Get-FinalBuildInputSnapshotBaseline。2026.10.30.29 の写し）と同じ入力から
#   1文字も違わない（0件・1件・複数件、同じ order の並び、変換PDFが無い・範囲外・隠し属性のページを含む）
# - 要求内のキャッシュの照合（Test-FinalSnapshotContentUnchanged）は、新しいスナップショットでも
#   変わっていなければ一致し、変換PDFが差し替わると外れる
# - 画面へ返す readiness の snapshot は itemCount・outputFileName・physicalPages だけで、
#   元の readiness・スナップショットは書き換えない（/api/state・/api/v2/state も同じ）
# - New-Sha256 は 0 バイト・1 バイト・バッファの境目・境目をまたぐ大きさで .NET の SHA-256 と同じ値を返し、
#   書き込みのために開かれているファイルも読める。Copy-FileSharedRead は同じバイト列を写す
# - Scan-Updates に一式を渡すと、その一式の原稿だけを確かめ、ほかの一式の原稿には触れない
$ErrorActionPreference = 'Stop'
$testsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $testsRoot
$appRoot = Join-Path $repoRoot 'app'
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-speed-' + [Guid]::NewGuid().ToString('N'))
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_OUTPUT_TEST_APPROOT')
    $testBody = @'
function Assert-SpeedTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalConfigRoot 'data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
Save-AppConfig ([pscustomobject][ordered]@{
    schemaVersion = 2; lastSubmissionDir = $submissionDir; lastDataDir = $dataDir; lastOutputDir = $outputDir
})
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')

# ---- 変更前の入力スナップショット（2026.10.30.29 の Get-FinalBuildInputSnapshot をそのまま写した） ----
# 指紋の互換を確かめる基準。server.ps1 側を直しても、ここは直さない。
function Get-FinalBuildInputSnapshotBaseline($Structure,[string]$Language,[string]$Volume,[string]$Category) {
    $scope=Resolve-DocumentPackScope $Structure $Category $false;$packId=[string]$scope.packId;$cat=[string]$scope.category
    if(@(Get-PackVolumeList $Language $scope.pack $false) -notcontains $Volume){throw [ArgumentException]::new("不正な成果物です: $Volume")}
    $workspace=Get-WorkspacePath $Language;$wbMap=@{};$targets=@(Get-Array $Structure.workbooks|Where-Object{Test-WorkbookPack $_ $packId});foreach($wb in $targets){$wbMap[[string]$wb.workbookId]=$wb}
    $sourceMap=@{};foreach($source in @(Get-Array (Get-DataProperty $Structure 'sources' @()))){$sourceMap[[string](Get-DataProperty $source 'sourceId' '')]=$source}
    $itemMap=@{};foreach($item in @(Get-Array (Get-DataProperty $Structure 'items' @()))){$itemMap[[string](Get-DataProperty $item 'itemId' '')]=$item}
    $unitMap=@{};foreach($unit in @(Get-Array (Get-DataProperty $Structure 'units' @()))){$unitMap[[string](Get-DataProperty $unit 'unitId' '')]=$unit}
    $artifactMap=@{};foreach($artifact in @(Get-Array (Get-DataProperty $Structure 'artifacts' @()))){$artifactMap[[string](Get-DataProperty $artifact 'artifactId' '')]=$artifact}
    $packObject=$scope.pack
    $effectiveTemplate=Get-PackEffectiveTemplate $Language $packObject
    $naming=Get-PackTargetOutputNaming $Structure $Language $scope $Volume;$settings=$naming.settings;$targetName=[string]$naming.targetName;$projectId=[string]$naming.projectId;$outputFileName=[string]$naming.outputFileName
    $document=[ordered]@{title=[string]$settings.documentTitle;subtitle=[string]$settings.documentSubtitle;targetName=$targetName;outputFileNamePattern=[string]$settings.outputFileNamePattern;outputFileName=$outputFileName}
    $pages=@(Get-Array $Structure.pages|Where-Object{$_.enabled -eq $true -and [string]$_.volume -eq $Volume -and $wbMap.ContainsKey([string]$_.workbookId)}|Sort-Object {[double]$_.order},{Resolve-PageId $_})
    $blockers=@();$manifest=@();$fpPages=@()
    $requiredRequirementIds=@{}
    foreach($requirement in @(Get-Array (Get-DataProperty $effectiveTemplate 'sourceRequirements' @()))){
        if(-not [bool](Get-DataProperty $requirement 'required' $true)){continue}
        $requirementId=[string](Get-DataProperty $requirement 'requirementId' '')
        if([string]::IsNullOrWhiteSpace($requirementId)){continue}
        $requiredRequirementIds[$requirementId]=$true
        $assigned=@($targets|Where-Object{[string](Get-DataProperty $_ 'requirementId' '') -eq $requirementId})
        if($assigned.Count -eq 0){
            $blockers+=[ordered]@{code='required-source-unregistered';requirementId=$requirementId;sourceId='';pageTitle='';workbookName=[string](Get-DataProperty $requirement 'displayName' '');ownerDepartment=[string](Get-DataProperty $requirement 'ownerDepartment' '');dueDate=[string](Get-DataProperty $requirement 'dueDate' '');message='必須原稿がまだ登録されていません。必要原稿リストから原稿を割り当ててください。'}
        }
    }
    # A source marked "required" is a safety contract, independent of template
    # defaults. Never publish a pack while that source is missing, failed, or
    # stale; otherwise a partial PDF could be labelled as current.
    $submissionRoot=[string](Get-DataProperty (Get-Paths) 'submissionDir' '')
    foreach($requiredSource in @($targets|Where-Object{[bool](Get-DataProperty $_ 'required' $false) -or $requiredRequirementIds.ContainsKey([string](Get-DataProperty $_ 'requirementId' ''))})){
            $requiredName=[string](Get-DataProperty $requiredSource 'displayName' (Get-DataProperty $requiredSource 'fileName' ''))
            $relativePath=[string](Get-DataProperty $requiredSource 'relativePath' '')
            $sourceExists=$false
            if(-not [string]::IsNullOrWhiteSpace($relativePath)){
                try{$sourceExists=Test-Path -LiteralPath (Join-Safe $submissionRoot $relativePath) -PathType Leaf}catch{$sourceExists=$false}
            }
            if(-not $sourceExists){
                $blockers+=[ordered]@{code='required-source-missing';sourceId=[string](Get-DataProperty $requiredSource 'workbookId' '');pageTitle='';workbookName=$requiredName;message='必須原稿のファイルが見つかりません。提出フォルダを確認してください。'}
                continue
            }
            $lastError=[string](Get-DataProperty $requiredSource 'lastError' '')
            $sourceStatus=([string](Get-DataProperty $requiredSource 'status' '')).ToLowerInvariant()
            if(-not [string]::IsNullOrWhiteSpace($lastError) -or $sourceStatus -match 'fail|error'){
                $blockers+=[ordered]@{code='required-source-failed';sourceId=[string](Get-DataProperty $requiredSource 'workbookId' '');pageTitle='';workbookName=$requiredName;message='必須原稿のPDF変換に失敗しています。エラーを解消して再作成してください。'}
                continue
            }
            if(-not (Test-WorkbookRenderIsCurrent $requiredSource)){
                $hasRendered=-not [string]::IsNullOrWhiteSpace([string](Get-DataProperty $requiredSource 'lastRenderedVersionId' ''))
                $blockers+=[ordered]@{code=$(if($hasRendered){'required-source-stale'}else{'required-source-not-rendered'});sourceId=[string](Get-DataProperty $requiredSource 'workbookId' '');pageTitle='';workbookName=$requiredName;message=$(if($hasRendered){'必須原稿が更新されています。先に変換PDFを再作成してください。'}else{'必須原稿の変換PDFが未作成です。先にPDFを作成してください。'})}
            }
    }
    # 変換PDFが見つからないページは manifest に入れない（$full を空に戻す）。入れると頁数を1と
    # 仮定した物理ページができ、試算の unknownPageIds から漏れて、後ろの頁・左右がずれる。
    if($pages.Count -eq 0){$blockers+= [ordered]@{code='no-pages';pageTitle='';workbookName='';message='対象ページがありません。ページ構成を確認してください。'}}
    foreach($p in $pages){$pageId=Resolve-PageId $p;$wb=$wbMap[[string]$p.workbookId];$title=[string]$p.title;if([string]::IsNullOrWhiteSpace($title)){$title=[string]$p.sheetName};$wbName=[string]$wb.fileName;$rel=[string]$p.contentPdf;$full='';$size=0L;$ticks=0L;$sourcePageCount=1
        # 記録済みの頁数（artifact.pageCount／unit.physicalPageCount）があれば変換PDFを開いて数えない。
        # 出力・readiness のたびに全ページの変換PDFを丸ごと読み込むと、それだけで数秒かかる。
        $item=if($itemMap.ContainsKey($pageId)){$itemMap[$pageId]}else{$null};$unitId=[string](Get-DataProperty $item 'unitId' (New-WorksheetUnitId ([string]$p.workbookId) ([string]$p.sheetName)));$unit=if($unitMap.ContainsKey($unitId)){$unitMap[$unitId]}else{$null};$artifactId=[string](Get-DataProperty $item 'artifactId' '');$artifact=if($artifactMap.ContainsKey($artifactId)){$artifactMap[$artifactId]}else{$null};$source=if($sourceMap.ContainsKey([string]$p.workbookId)){$sourceMap[[string]$p.workbookId]}else{$null};$recordedPageCount=Get-IntDataProperty $artifact 'pageCount' (Get-IntDataProperty $unit 'physicalPageCount' 0);if($recordedPageCount -gt 0){$sourcePageCount=$recordedPageCount}
        if(-not $rel){$blockers+=[ordered]@{code='content-missing';pageTitle=$title;workbookName=$wbName;message='PDF未作成のページがあります。先にPDF作成してください。'}}
        else{try{$full=[IO.Path]::GetFullPath((Join-Path $workspace $rel));$root=[IO.Path]::GetFullPath($workspace);if(-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)){$root+=[IO.Path]::DirectorySeparatorChar};if(-not $full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'outside'};if(-not(Test-Path -LiteralPath $full)){throw 'missing'};$it=Get-Item -LiteralPath $full;$size=$it.Length;$ticks=$it.LastWriteTimeUtc.Ticks;if($recordedPageCount -le 0){$sourcePageCount=Get-PdfPageCount $full}}catch{$full='';$blockers+=[ordered]@{code='content-file-missing';pageTitle=$title;workbookName=$wbName;message='レンダリング済みPDFが見つかりません。先にPDF作成してください。'}}}
        if(-not(Test-WorkbookRenderIsCurrent $wb)){$blockers+=[ordered]@{code='stale-content';pageTitle=$title;workbookName=$wbName;message='元原稿が更新されています。先に変換PDFを作成してください。'}}
        elseif(-not(Test-WorkbookRenderedSheetContains $wb ([string]$p.sheetName))){$blockers+=[ordered]@{code='sheet-changed';pageTitle=$title;workbookName=$wbName;message='原稿のページ構成が変わっています。先に変換PDFを作成してください。'}}
        elseif(-not(Test-PageContentMatchesWorkbookVersion $p $wb)){$blockers+=[ordered]@{code='old-content';pageTitle=$title;workbookName=$wbName;message='古いPDF参照が残っています。先にPDF作成してください。'}}
        $rangeValue=Get-DataProperty $p 'pageRange' (Get-DataProperty $item 'pageRange' $null);$range=$null;try{$range=ConvertTo-NormalizedPageRange $rangeValue}catch{$blockers+=[ordered]@{code='invalid-page-range';pageTitle=$title;workbookName=$wbName;message=$_.Exception.Message}}
        $rangeStart=if($null -ne $range){[int]$range.start}else{1};$rangeEnd=if($null -ne $range){[int]$range.end}else{$sourcePageCount}
        if($rangeStart -gt $sourcePageCount -or $rangeEnd -gt $sourcePageCount){$blockers+=[ordered]@{code='page-range-out-of-bounds';pageTitle=$title;workbookName=$wbName;message="ページ範囲 $rangeStart-$rangeEnd は変換PDF（$sourcePageCount ページ）の範囲外です。"}}
        $sourceType=[string](Get-DataProperty $source 'sourceType' (Get-DataProperty $wb 'sourceType' 'excel'));$sourceKey=[string](Get-DataProperty $unit 'sourceKey' ([string]$p.sheetName));$sectionId="source:$([string]$p.workbookId)";$sectionTitle=[string](Get-DataProperty $source 'displayName' (Get-DataProperty $wb 'displayName' $wbName));$bookmarkTitle=if([string]::IsNullOrWhiteSpace($sectionTitle)){$title}else{"$sectionTitle / $title"}
        $normalized=($rel -replace '\\','/').ToLowerInvariant();$fpPages+=[ordered]@{pageId=$pageId;itemId=[string](Get-DataProperty $item 'itemId' $pageId);order=[double]$p.order;enabled=[bool]$p.enabled;volume=[string]$p.volume;title=$title;pageRange=$range;numberingMode=[string]$p.numberingMode;contentPdf=$normalized;contentPdfSize=$size;contentPdfLastWriteUtcTicks=$ticks;lastRenderedVersionId=[string]$wb.lastRenderedVersionId}
        if($full){$manifest+=[ordered]@{pageId=$pageId;itemId=[string](Get-DataProperty $item 'itemId' $pageId);sourceId=[string]$p.workbookId;unitId=$unitId;artifactId=[string](Get-DataProperty $artifact 'artifactId' $artifactId);sourceType=$sourceType;sourceKey=$sourceKey;title=$title;bookmarkTitle=$bookmarkTitle;sectionId=$sectionId;sectionTitle=$sectionTitle;sourcePdf=$full;sourcePageStart=$rangeStart;sourcePageEnd=$rangeEnd;sourcePageCount=$sourcePageCount;numberingMode=[string]$p.numberingMode;punchShiftPt=(Convert-CmToPt 0.2)}}
    }
    # 左右は出力先ごとの物理ページ番号の奇偶で決める（奇数＝右）。ReportPdfComposer.java の
    # パンチ穴シフトと同じ規則。1頁は台紙0の右、2・3頁は台紙1。指紋（fpPages）には入れない。
    $physicalPages=@();$outputPageNumber=0
    foreach($entry in $manifest){
        $pagePartCount=[Math]::Max(0,[int]$entry.sourcePageEnd-[int]$entry.sourcePageStart+1)
        for($sourcePage=[int]$entry.sourcePageStart;$sourcePage -le [int]$entry.sourcePageEnd;$sourcePage++){$outputPageNumber++;$physicalPages+=[ordered]@{outputPageNumber=$outputPageNumber;side=$(if($outputPageNumber%2 -eq 1){'right'}else{'left'});spreadIndex=[int][Math]::Floor($outputPageNumber/2);pagePartIndex=$sourcePage-[int]$entry.sourcePageStart+1;pagePartCount=$pagePartCount;kind='content';itemId=[string]$entry.itemId;pageId=[string]$entry.pageId;sourceId=[string]$entry.sourceId;unitId=[string]$entry.unitId;artifactId=[string]$entry.artifactId;sourceType=[string]$entry.sourceType;sourceKey=[string]$entry.sourceKey;sourcePdfPageNumber=$sourcePage;title=[string]$entry.title;sectionId=[string]$entry.sectionId;numberingMode=[string]$entry.numberingMode}}
    }
    $fpDocument=[ordered]@{};foreach($fpKey in @($document.Keys)){$fpDocument[$fpKey]=$document[$fpKey]};$fpDocument.outputFileName=[string]$naming.outputFileNameKey;$input=[ordered]@{composerProfileVersion=$Script:FinalPdfComposerProfileVersion;language=$Language;packId=$packId;category=$cat;volume=$Volume;document=$fpDocument;pages=$fpPages};$json=ConvertTo-Json $input -Depth 20 -Compress;$fingerprint=Get-Sha256Text $json
    return [ordered]@{language=$Language;packId=$packId;category=$cat;volume=$Volume;pages=$pages;manifestPages=$manifest;physicalPages=$physicalPages;document=$document;blockers=@($blockers);itemCount=$pages.Count;pageCount=$outputPageNumber;projectId=$projectId;outputFileName=$outputFileName;fingerprint=$fingerprint;fingerprintInput=$input}
}

# ---- 入力スナップショット: 変更前と1文字も違わない ------------------------------------------
function New-SpeedWorkbook([string]$PackId) {
    $workbookId = 'wb_' + [Guid]::NewGuid().ToString('N').Substring(0,12)
    [void](Update-StructureLocked 'ja' {
        param($st)
        $wb = [pscustomobject][ordered]@{
            workbookId=$workbookId; packId=$PackId; sourceType='excel'; fileName="$workbookId.xlsx"; relativePath="$workbookId.xlsx"; displayName="原稿 $workbookId"
            status='rendered'; excelSheetSelectionMode='all-visible'; lastRenderedSheetSelectionMode='all-visible'; lastRenderedSheets=@(); lastRenderedVersionId='v1'
        }
        Set-NoteProperty $st 'workbooks' (@(Get-Array $st.workbooks) + @($wb))
        return $null
    })
    return $workbookId
}
# $Spec は @{ sheet; volume; enabled; order; range; pdfPages }。pdfPages が 0 なら変換PDFを置かない、
# -1 なら contentPdf 自体を持たない。
function Add-SpeedPages([string]$WorkbookId, [object[]]$Spec) {
    foreach ($s in $Spec) {
        $rel = $null
        if ([int]$s.pdfPages -ge 0) { $rel = "content-pdf\$WorkbookId\v1\" + [string]$s.sheet + '.pdf' }
        if ([int]$s.pdfPages -gt 0) {
            $full = [IO.Path]::GetFullPath((Join-Path (Get-WorkspacePath 'ja') $rel))
            New-Item -ItemType Directory -Path (Split-Path -Parent $full) -Force | Out-Null
            $body = '%PDF-1.4' + "`n" + ((1..([int]$s.pdfPages) | ForEach-Object { "$_ 0 obj << /Type /Page >> endobj" }) -join "`n") + "`n%%EOF`n"
            [IO.File]::WriteAllText($full, $body, [Text.Encoding]::ASCII)
            # 隠し属性の変換PDF。以前の Get-Item（-Force なし）は見つからない扱いにしていた。
            if ([bool]$s.hidden) { [IO.File]::SetAttributes($full, ([IO.File]::GetAttributes($full) -bor [IO.FileAttributes]::Hidden)) }
        }
        $s['rel'] = $rel
    }
    [void](Update-StructureLocked 'ja' {
        param($st)
        $pages = @(Get-Array $st.pages)
        foreach ($s in $Spec) {
            $pages += [pscustomobject][ordered]@{
                pageId=("{0}-{1}" -f $WorkbookId, [string]$s.sheet); workbookId=$WorkbookId; sheetName=[string]$s.sheet; title=[string]$s.title
                volume=[string]$s.volume; enabled=[bool]$s.enabled; order=$s.order; pageRange=$s.range; numberingMode='visible'
                contentPdf=$s.rel; status='rendered'
            }
        }
        Set-NoteProperty $st 'pages' @($pages)
        return $null
    })
}
function ConvertTo-SpeedJson($Value) { return (ConvertTo-Json -InputObject $Value -Depth 30 -Compress) }
function Assert-SnapshotSame([string]$PackId, [string]$Volume, [string]$Label) {
    $structure = Get-Structure 'ja'
    # 出力ファイル名は日付を含みうる。2回の呼び出しの間に日付が変わったら作り直す（指紋は日付を含まない）。
    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        $old = Get-FinalBuildInputSnapshotBaseline $structure 'ja' $Volume $PackId
        $new = Get-FinalBuildInputSnapshot $structure 'ja' $Volume $PackId
        if ([string]$new.outputFileName -eq [string]$old.outputFileName) { break }
    }
    Assert-SpeedTest ([string]$new.fingerprint -eq [string]$old.fingerprint) "$Label : the fingerprint changed ($($old.fingerprint) -> $($new.fingerprint))."
    Assert-SpeedTest ((ConvertTo-SpeedJson $new.fingerprintInput) -eq (ConvertTo-SpeedJson $old.fingerprintInput)) "$Label : the fingerprint input changed."
    foreach ($key in @('manifestPages','physicalPages','blockers','document')) {
        Assert-SpeedTest ((ConvertTo-SpeedJson $new[$key]) -eq (ConvertTo-SpeedJson $old[$key])) "$Label : $key changed."
    }
    Assert-SpeedTest ((@($new.pages | ForEach-Object { Resolve-PageId $_ }) -join '|') -eq (@($old.pages | ForEach-Object { Resolve-PageId $_ }) -join '|')) "$Label : the page order changed."
    foreach ($key in @('itemCount','pageCount','outputFileName','projectId','packId','category','volume')) {
        Assert-SpeedTest ([string]$new[$key] -eq [string]$old[$key]) "$Label : $key changed ($($old[$key]) -> $($new[$key]))."
    }
    # 配列のまま返す（List を返すと、呼び出し側の .Count や JSON の形が変わりうる）。
    foreach ($key in @('manifestPages','physicalPages','blockers')) {
        Assert-SpeedTest ($new[$key] -is [object[]]) "$Label : $key is not an array."
    }
    Assert-SpeedTest ($new.fingerprintInput.pages -is [object[]]) "$Label : the fingerprint pages are not an array."
    return $new
}

$pack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='速度の一式' })
$packId = [string]$pack.packId
$otherPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='別の一式' })
$otherPackId = [string]$otherPack.packId
$emptyPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='ページの無い一式' })
$emptyPackId = [string]$emptyPack.packId
$wbA = New-SpeedWorkbook $packId
$wbB = New-SpeedWorkbook $packId
$wbOther = New-SpeedWorkbook $otherPackId
Add-SpeedPages $wbA @(
    @{ sheet='表紙'; title=''; volume='ja-main'; enabled=$true; order=10; range=$null; pdfPages=1 },
    @{ sheet='損益'; title='損益計算書'; volume='ja-main'; enabled=$true; order=20; range=$null; pdfPages=3 },
    @{ sheet='営利'; title=''; volume='ja-main'; enabled=$true; order=20; range='2-3'; pdfPages=3 },
    @{ sheet='消えた'; title=''; volume='ja-main'; enabled=$true; order=25; range=$null; pdfPages=0 },
    @{ sheet='未作成'; title=''; volume='ja-main'; enabled=$true; order=26; range=$null; pdfPages=-1 },
    @{ sheet='範囲外'; title=''; volume='ja-main'; enabled=$true; order=27; range='5'; pdfPages=1 },
    @{ sheet='隠し'; title=''; volume='ja-main'; enabled=$true; order=28; range=$null; pdfPages=2; hidden=$true },
    @{ sheet='外した'; title=''; volume='none'; enabled=$false; order=5; range=$null; pdfPages=1 },
    @{ sheet='無効'; title=''; volume='ja-main'; enabled=$false; order=6; range=$null; pdfPages=1 },
    @{ sheet='付録'; title='付録'; volume='ja-appendix'; enabled=$true; order=10; range=$null; pdfPages=2 }
)
Add-SpeedPages $wbB @(
    @{ sheet='注記'; title=''; volume='ja-main'; enabled=$true; order=15.5; range=$null; pdfPages=2 }
)
Add-SpeedPages $wbOther @(
    @{ sheet='他'; title=''; volume='ja-main'; enabled=$true; order=1; range=$null; pdfPages=1 }
)

$mainSnap = Assert-SnapshotSame $packId 'ja-main' '本体（複数ページ）'
# 空振りでないこと: 複数ページの変換PDF（物理ページ＞manifest）と、変換PDFの無いページ（指紋には入るが
# manifest には入らない）が実際に含まれ、ほかの一式のページが入っていない。
Assert-SpeedTest (@($mainSnap.manifestPages).Count -ge 3) "本体 should have several manifest entries, got $(@($mainSnap.manifestPages).Count)."
Assert-SpeedTest (@($mainSnap.physicalPages).Count -gt @($mainSnap.manifestPages).Count) '本体 should include a multi-page converted PDF.'
Assert-SpeedTest (@($mainSnap.fingerprintInput.pages).Count -gt @($mainSnap.manifestPages).Count) '本体 should include pages without a converted PDF.'
$mainCodes = @($mainSnap.blockers | ForEach-Object { [string]$_.code })
foreach ($code in @('content-file-missing','content-missing','page-range-out-of-bounds')) {
    Assert-SpeedTest ($mainCodes -contains $code) "本体 should report the blocker '$code' (got: $($mainCodes -join ', '))."
}
Assert-SpeedTest (-not ((ConvertTo-SpeedJson $mainSnap.manifestPages).Contains($wbOther))) 'A page of another pack entered the snapshot.'
$appendixSnap = Assert-SnapshotSame $packId 'ja-appendix' '補足（1ページ）'
Assert-SpeedTest (@($appendixSnap.manifestPages).Count -eq 1 -and @($appendixSnap.physicalPages).Count -eq 2) '補足 should have one manifest entry with two physical pages.'
$emptySnap = Assert-SnapshotSame $emptyPackId 'ja-main' 'ページの無い一式'
Assert-SpeedTest (@($emptySnap.blockers | Where-Object { [string]$_.code -eq 'no-pages' }).Count -eq 1) 'An empty target did not report no-pages.'

# 要求内キャッシュの照合は、変換PDFのサイズ・更新時刻を記録どおりに読めていれば一致する。
Assert-SpeedTest (Test-FinalSnapshotContentUnchanged 'ja' $mainSnap) 'The cache check rejected a snapshot whose converted PDFs did not change.'
$touched = [IO.Path]::GetFullPath((Join-Path (Get-WorkspacePath 'ja') "content-pdf\$wbA\v1\表紙.pdf"))
[IO.File]::AppendAllText($touched, "% changed`n")
Assert-SpeedTest (-not (Test-FinalSnapshotContentUnchanged 'ja' $mainSnap)) 'The cache check accepted a snapshot whose converted PDF was replaced.'
$afterTouch = Assert-SnapshotSame $packId 'ja-main' '本体（変換PDFの差し替え後）'
Assert-SpeedTest ([string]$afterTouch.fingerprint -ne [string]$mainSnap.fingerprint) 'Replacing a converted PDF did not change the fingerprint.'

# ---- readiness の応答: snapshot は画面が読む3つだけ。元は書き換えない ---------------------------
$ready = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $packId
$readyJsonBefore = ConvertTo-Json -InputObject $ready -Depth 50 -Compress
$view = ConvertTo-FinalReadinessResponse $ready
Assert-SpeedTest ((@($view.snapshot.Keys) -join ',') -eq 'itemCount,outputFileName,physicalPages') "The response snapshot keys are $(@($view.snapshot.Keys) -join ',')."
Assert-SpeedTest ((ConvertTo-SpeedJson $view.snapshot.physicalPages) -eq (ConvertTo-SpeedJson $ready.snapshot.physicalPages)) 'The response physical pages differ from the snapshot.'
Assert-SpeedTest ([int]$view.snapshot.itemCount -eq [int]$ready.snapshot.itemCount -and [string]$view.snapshot.outputFileName -eq [string]$ready.snapshot.outputFileName) 'The response item count or file name differs from the snapshot.'
foreach ($key in @($ready.Keys | Where-Object { $_ -ne 'snapshot' })) {
    Assert-SpeedTest ((ConvertTo-SpeedJson $view[$key]) -eq (ConvertTo-SpeedJson $ready[$key])) "The response changed the readiness field '$key'."
}
$viewJson = ConvertTo-Json -InputObject $view -Depth 50 -Compress
foreach ($field in @('"manifestPages"','"fingerprintInput"','"sourcePdf"','"pages"','"document"')) {
    Assert-SpeedTest (-not $viewJson.Contains($field)) "The readiness response still contains $field."
}
Assert-SpeedTest ((ConvertTo-Json -InputObject $ready -Depth 50 -Compress) -eq $readyJsonBefore) 'Making the response changed the original readiness.'
Assert-SpeedTest ($ready.snapshot.Contains('manifestPages') -and $ready.snapshot.Contains('fingerprintInput')) 'The server-side readiness lost its full snapshot.'
# 1ページだけでも配列のまま返す。
$appendixView = ConvertTo-FinalReadinessResponse (Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-appendix' $packId)
Assert-SpeedTest ((ConvertTo-Json -InputObject $appendixView -Depth 50 -Compress).Contains('"physicalPages":[{')) 'The physical pages of a small target are not an array.'
$emptyView = ConvertTo-FinalReadinessResponse (Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $emptyPackId)
Assert-SpeedTest ((ConvertTo-Json -InputObject $emptyView -Depth 50 -Compress).Contains('"physicalPages":[]')) 'The physical pages of an empty target are not an empty array.'
# 確認に失敗した出力先（snapshot を持たない）はそのまま返す。
$failedReady = [ordered]@{ packId=$packId; targetId='main'; readinessError=$true; canBuild=$false; blockers=@() }
$failedView = ConvertTo-FinalReadinessResponse $failedReady
Assert-SpeedTest ((@($failedView.Keys) -join ',') -eq (@($failedReady.Keys) -join ',')) 'A readiness without a snapshot was changed.'

# 状態の応答（/api/state・/api/v2/state）も同じ形。
$v2 = Get-V2StatePayload 'ja'
$v2Ready = $v2.finalReadiness[$packId].volumes['ja-main']
Assert-SpeedTest ($null -ne $v2Ready -and (@($v2Ready.snapshot.Keys) -join ',') -eq 'itemCount,outputFileName,physicalPages') '/api/v2/state still returns the full input snapshot.'
Assert-SpeedTest ([string]$v2Ready.currentFingerprint -eq [string]$ready.currentFingerprint) '/api/v2/state returned a different fingerprint.'
$legacyState = Get-StatePayload 'ja'
foreach ($cat in @($legacyState.finalReadiness.Keys)) {
    foreach ($legacyReady in @($legacyState.finalReadiness[$cat].volumes.Values)) {
        Assert-SpeedTest ((@($legacyReady.snapshot.Keys) -join ',') -eq 'itemCount,outputFileName,physicalPages') "/api/state still returns the full input snapshot for $cat."
    }
}
$v2Json = ConvertTo-Json -InputObject $v2.finalReadiness -Depth 50 -Compress
Assert-SpeedTest (-not $v2Json.Contains('"manifestPages"') -and -not $v2Json.Contains('"fingerprintInput"')) 'The state readiness still contains the full input snapshot.'

# ---- ハッシュとコピー: 読み込みバッファを大きくしても値・中身・共有モードは同じ -----------------
Assert-SpeedTest ([int]$Script:FileReadBufferSize -ge 65536) 'The file read buffer is still small.'
$hashDir = Join-Path $Script:LocalConfigRoot 'hash'
New-Item -ItemType Directory -Path $hashDir -Force | Out-Null
$buffer = [int]$Script:FileReadBufferSize
$rng = New-Object Random 20261102
$reference = [Security.Cryptography.SHA256]::Create()
try {
    foreach ($size in @(0, 1, 4095, 4096, 4097, ($buffer - 1), $buffer, ($buffer + 1), (3 * $buffer + 7))) {
        $bytes = New-Object byte[] $size
        $rng.NextBytes($bytes)
        $path = Join-Path $hashDir ("{0}.bin" -f $size)
        [IO.File]::WriteAllBytes($path, $bytes)
        $expected = [BitConverter]::ToString($reference.ComputeHash($bytes)).Replace('-', '')
        Assert-SpeedTest ((New-Sha256 $path) -eq $expected) "New-Sha256 returned a different hash for $size bytes."
        $copy = Join-Path $hashDir ("{0}.copy" -f $size)
        # 長い既存ファイルへ写しても、後ろに古い中身が残らない。
        [IO.File]::WriteAllBytes($copy, (New-Object byte[] ($size + 10)))
        Copy-FileSharedRead $path $copy
        $copied = [IO.File]::ReadAllBytes($copy)
        Assert-SpeedTest ($copied.Length -eq $size -and [BitConverter]::ToString($reference.ComputeHash($copied)).Replace('-', '') -eq $expected) "Copy-FileSharedRead did not copy $size bytes exactly."
    }
    # Excel が書き込みのために開いているファイルも読める（FileShare.ReadWrite のまま）。
    $openPath = Join-Path $hashDir ("{0}.bin" -f ($buffer + 1))
    $openBytes = [IO.File]::ReadAllBytes($openPath)
    $openExpected = [BitConverter]::ToString($reference.ComputeHash($openBytes)).Replace('-', '')
    $writer = [IO.File]::Open($openPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::ReadWrite)
    try {
        Assert-SpeedTest ((New-Sha256 $openPath) -eq $openExpected) 'New-Sha256 could not read a file opened for writing.'
        $openCopy = Join-Path $hashDir 'open.copy'
        Copy-FileSharedRead $openPath $openCopy
        Assert-SpeedTest ([BitConverter]::ToString($reference.ComputeHash([IO.File]::ReadAllBytes($openCopy))).Replace('-', '') -eq $openExpected) 'Copy-FileSharedRead could not copy a file opened for writing.'
    } finally { $writer.Dispose() }
    Assert-SpeedTest ((New-Sha256 (Join-Path $hashDir 'missing.bin')) -eq '') 'New-Sha256 of a missing file is not empty.'
} finally { $reference.Dispose() }

# ---- Scan-Updates: 一式を渡すと、その一式の原稿だけを確かめる -------------------------------
function Set-SpeedSource([string]$WorkbookId, [string]$Text) {
    [IO.File]::WriteAllText((Join-Path $submissionDir "$WorkbookId.xlsx"), $Text, [Text.Encoding]::ASCII)
}
function Get-SpeedWorkbook([string]$WorkbookId) {
    return @(Get-Array (Get-Structure 'ja').workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId })[0]
}
foreach ($id in @($wbA, $wbB, $wbOther)) { Set-SpeedSource $id "first $id" }
[void](Scan-Updates 'ja')
foreach ($id in @($wbA, $wbB, $wbOther)) {
    Assert-SpeedTest ([int64](Get-DataProperty (Get-SpeedWorkbook $id) 'currentExcelSize' -1) -eq ("first $id").Length) "The full scan did not record $id."
}
foreach ($id in @($wbA, $wbB, $wbOther)) { Set-SpeedSource $id "second version of $id" }
$packScan = Scan-Updates 'ja' $null $false $packId
Assert-SpeedTest ([int]$packScan.scanned -eq 2) "A pack scan looked at $($packScan.scanned) sources instead of 2."
foreach ($id in @($wbA, $wbB)) {
    Assert-SpeedTest ([int64](Get-DataProperty (Get-SpeedWorkbook $id) 'currentExcelSize' -1) -eq ("second version of $id").Length) "The pack scan did not detect the update of $id."
}
Assert-SpeedTest ([int64](Get-DataProperty (Get-SpeedWorkbook $wbOther) 'currentExcelSize' -1) -eq ("first $wbOther").Length) 'The pack scan touched a source of another pack.'
$fullScan = Scan-Updates 'ja'
Assert-SpeedTest ([int]$fullScan.scanned -eq 3) "A full scan looked at $($fullScan.scanned) sources instead of 3."
Assert-SpeedTest ([int64](Get-DataProperty (Get-SpeedWorkbook $wbOther) 'currentExcelSize' -1) -eq ("second version of $wbOther").Length) 'A full scan without a pack did not check every source.'

Write-Output 'speed regression ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
