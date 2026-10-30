param()

# 提出用PDFの出力先パスの重なり（H-1）と、共有用フォルダーへのコピー
# （/api/v2/outputs/publish-all の本体 Publish-DocumentPackAllPdfsToShared）を、
# server.ps1 の実物の関数で確かめる。Java・Office は使わない。組版の代わりに、
# 出力済みのPDFと出力状態を直接置き、入力スナップショットだけを差し替える。
# あわせて、ページ構成の保存（Reorder-Pages）が不完全な並びを拒むこと（H-2）も見る。
$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-output-publish-' + [Guid]::NewGuid().ToString('N'))
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
function Assert-OutputTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Assert-Throws([scriptblock]$Action, [string]$ExpectedFragment, [string]$Message) {
    $caught = $null
    try { & $Action } catch { $caught = $_.Exception.Message }
    if ($null -eq $caught) { throw "$Message (no error was raised)" }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedFragment) -and -not $caught.Contains($ExpectedFragment)) { throw "$Message (unexpected error: $caught)" }
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

# 一式に原稿1件と、指定した出力先へページを置く。PDFの中身は使わない。
function Add-TestPages([string]$PackId, [string[]]$Volumes) {
    [void](Update-StructureLocked 'ja' {
        param($st)
        $workbookId = 'wb_' + [Guid]::NewGuid().ToString('N').Substring(0,12)
        $wb = [pscustomobject][ordered]@{ workbookId=$workbookId; packId=$PackId; fileName="$workbookId.xlsx"; relativePath="$workbookId.xlsx"; displayName=$workbookId; status='rendered' }
        Set-NoteProperty $st 'workbooks' (@(Get-Array $st.workbooks) + @($wb))
        $pages = @(Get-Array $st.pages)
        $order = 10
        foreach ($volume in $Volumes) {
            $pages += [pscustomobject][ordered]@{ pageId="${workbookId}_$order"; workbookId=$workbookId; sheetName="Sheet$order"; title="Sheet$order"; volume=$volume; enabled=$true; order=$order; numberingMode='visible' }
            $order += 10
        }
        Set-NoteProperty $st 'pages' @($pages)
        return $null
    })
}
# 保存時の検査を通らない古い型を、保存済みのデータとして直接置く。
function Set-LegacyPattern([string]$PackId, [string]$Pattern) {
    [void](Update-StructureLocked 'ja' {
        param($st)
        $pack = Get-PackRecord $st $PackId
        $settings = Get-ResolvedPackSettings $pack 'ja'
        $settings.outputFileNamePattern = $Pattern
        Set-NoteProperty $pack 'settings' $settings
        return $null
    })
}
function Get-PlannedOutputPath([string]$PackId, [string]$Volume) {
    $st = Get-Structure 'ja'
    $scope = Resolve-DocumentPackScope $st $PackId $true
    return [IO.Path]::GetFullPath((Join-Path $outputDir ([string](Get-PackTargetOutputNaming $st 'ja' $scope $Volume).outputFileName)))
}

# ---- H-1: 出力先パスの重なり -------------------------------------------------
$sharedName = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='共有名の一式' })
$sharedId = [string]$sharedName.packId
Add-TestPages $sharedId @('ja-main','ja-appendix')
Assert-OutputTest (Test-PackVolumeHasEnabledPages (Get-Structure 'ja') $sharedId 'ja-appendix') 'Test pages were not stored on the appendix target.'
# 既定の型（{targetName} 入り）なら本体と補足は別の名前になる。
Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' $sharedId @('ja-main','ja-appendix')
Set-LegacyPattern $sharedId '{packName}_{yyyyMMdd}.pdf'
Assert-Throws { Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' $sharedId @('ja-main') } '{targetName}' 'Building 本体 to the same path as 補足 was not rejected.'
Assert-Throws { Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' $sharedId @('ja-appendix') } '{targetName}' 'Building 補足 to the same path as 本体 was not rejected.'

# 補足にページが無ければ補足は書き出されないので、古い型でも本体だけは出力できる。
$mainOnly = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='本体だけの一式' })
$mainOnlyId = [string]$mainOnly.packId
Add-TestPages $mainOnlyId @('ja-main')
Set-LegacyPattern $mainOnlyId '{packName}_{yyyyMMdd}.pdf'
Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' $mainOnlyId @('ja-main')

# 他の一式が記録している出力PDFと同じパスにも書かない（{packName} を含まない型どうし）。
$packA = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='一式A'; settings=[pscustomobject]@{ outputFileNamePattern='{targetName}.pdf' } })
$packB = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='一式B'; settings=[pscustomobject]@{ outputFileNamePattern='{targetName}.pdf' } })
$packAPath = Get-PlannedOutputPath ([string]$packA.packId) 'ja-main'
[void](Update-StructureLocked 'ja' {
    param($st)
    $state = Get-PackOutputState $st 'ja' ([string]$packA.packId) 'ja-main' $true
    Set-NoteProperty $state 'outputPdf' $packAPath
    Set-NoteProperty $state 'builtFingerprint' 'fp'
    return $null
})
Assert-Throws { Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' ([string]$packB.packId) @('ja-main') } '一式「一式A」' 'Overwriting the output PDF recorded by another pack was not rejected.'
# 自分が前回書いたPDFへの再出力は重なりではない。
Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' ([string]$packA.packId) @('ja-main')
# 大文字小文字だけが違うパスも同じファイルとして扱う。
[void](Update-StructureLocked 'ja' {
    param($st)
    $state = Get-PackOutputState $st 'ja' ([string]$packA.packId) 'ja-main' $true
    Set-NoteProperty $state 'outputPdf' $packAPath.ToUpperInvariant()
    return $null
})
Assert-Throws { Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' ([string]$packB.packId) @('ja-main') } '一式「一式A」' 'A recorded output path differing only by case was not treated as the same file.'

# 組み込み一式の状態は volumes と outputs の両方に写しがある。自分の写しとは重ならない。
$ecmMainPath = Get-PlannedOutputPath 'pack_ecm' 'ja-main'
[void](Update-StructureLocked 'ja' {
    param($st)
    $state = Get-PackOutputState $st 'ja' 'pack_ecm' 'ja-main' $true
    Set-NoteProperty $state 'outputPdf' $ecmMainPath
    return $null
})
$ecmStructure = Get-Structure 'ja'
$ecmMirror = Get-DataProperty (Get-DataProperty $ecmStructure 'outputs' $null) 'pack_ecm|main' $null
Assert-OutputTest ($null -ne $ecmMirror -and [string](Get-DataProperty $ecmMirror 'outputPdf' '') -eq $ecmMainPath) 'The built-in output mirror was not populated; the self-exclusion case is not exercised.'
Assert-PackOutputPathsDistinct $ecmStructure 'ja' 'pack_ecm' @('ja-main')

# ---- H-2: ページ構成の保存は、送られたレーンのページを欠けなく1回ずつ含む --------
$layoutPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='並べ替えの一式' })
$layoutId = [string]$layoutPack.packId
Add-TestPages $layoutId @('ja-main','ja-main','ja-appendix','none')
function Get-LayoutIds([string]$Volume) {
    $st = Get-Structure 'ja'
    $ids = @{}
    foreach ($wb in @(Get-Array $st.workbooks | Where-Object { Test-WorkbookPack $_ $layoutId })) { $ids[[string]$wb.workbookId] = $true }
    return @(Get-Array $st.pages | Where-Object { $ids.ContainsKey([string]$_.workbookId) -and [string]$_.volume -eq $Volume } | Sort-Object {[double]$_.order} | ForEach-Object { Resolve-PageId $_ })
}
function Get-LayoutSignature { return (@('ja-main','ja-appendix','none') | ForEach-Object { $_ + '=' + ((Get-LayoutIds $_) -join ',') }) -join ';' }
$main = @(Get-LayoutIds 'ja-main'); $appendix = @(Get-LayoutIds 'ja-appendix'); $none = @(Get-LayoutIds 'none')
Assert-OutputTest ($main.Count -eq 2 -and $appendix.Count -eq 1 -and $none.Count -eq 1) 'Layout test pages were not placed as expected.'
$before = Get-LayoutSignature
function New-Layout($Volumes) { return [pscustomobject][ordered]@{ packId=$layoutId; volumes=[pscustomobject]$Volumes } }
Assert-Throws { Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@(); 'ja-appendix'=$appendix; none=$none })) } '不完全' 'An empty lane that still has pages was accepted.'
Assert-Throws { Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@($main[1]); 'ja-appendix'=$appendix; none=$none })) } '不完全' 'A lane missing one of its pages was accepted.'
Assert-Throws { Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@($main + @('page_unknown')); 'ja-appendix'=$appendix; none=$none })) } 'この一式に無いページ' 'An unknown page id was silently ignored.'
Assert-Throws { Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=$main; 'ja-appendix'=@($appendix + @($main[0])); none=$none })) } '2回' 'A page sent in two lanes was accepted.'
Assert-OutputTest ((Get-LayoutSignature) -eq $before) 'A rejected layout save still changed the page composition.'
# 正しい全体の並べ替えと、送ったレーンの間だけの移動（補足から本体へ）は通る。
[void](Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@($main[1],$main[0]); 'ja-appendix'=$appendix; none=$none })))
Assert-OutputTest (((Get-LayoutIds 'ja-main') -join ',') -eq (@($main[1],$main[0]) -join ',')) 'A complete reorder was not applied.'
[void](Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@($main[1],$main[0],$appendix[0]) })))
Assert-OutputTest (((Get-LayoutIds 'ja-main') -join ',') -eq (@($main[1],$main[0],$appendix[0]) -join ',') -and @(Get-LayoutIds 'ja-appendix').Count -eq 0) 'Moving a page from an unsent lane into a sent lane was rejected or misapplied.'

# ---- M-1: 日付が変わっただけでは「再出力が必要」にしない ----------------------------
# ここではまだ実物の Get-FinalBuildInputSnapshot を使う（下で差し替える）。
$dated = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='日付の一式' })
Add-TestPages ([string]$dated.packId) @('ja-main')
$today = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' ([string]$dated.packId)
$realGetDate = Get-Command Get-Date -CommandType Cmdlet
function Get-Date { param([string]$Format) $tomorrow = (& $realGetDate).AddDays(1); if ($Format) { return $tomorrow.ToString($Format) } return $tomorrow }
try { $tomorrow = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' ([string]$dated.packId) }
finally { Remove-Item Function:\Get-Date }
Assert-OutputTest ([string]$today.outputFileName -ne [string]$tomorrow.outputFileName) 'The date override did not reach the output file name; the M-1 case is not exercised.'
Assert-OutputTest ([string]$today.fingerprint -eq [string]$tomorrow.fingerprint) 'A date change alone changed the output input fingerprint.'

# ---- 性能: 記録済みの頁数があれば変換PDFを開かない --------------------------------
# ここもまだ実物の Get-FinalBuildInputSnapshot を使う。頁数の数え方だけ差し替えて呼ばれた回数を見る。
$counted = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='頁数記録の一式' })
$countedId = [string]$counted.packId
Add-TestPages $countedId @('ja-main')
$countedPdfRel = 'content-pdf\selfcheck\counted.pdf'
$countedPdf = Join-Path (Get-WorkspacePath 'ja') $countedPdfRel
New-Item -ItemType Directory -Path (Split-Path -Parent $countedPdf) -Force | Out-Null
[IO.File]::WriteAllText($countedPdf, '%PDF-counted')
[void](Update-StructureLocked 'ja' {
    param($st)
    $ids = @{}; foreach ($wb in @(Get-Array $st.workbooks | Where-Object { Test-WorkbookPack $_ $countedId })) { $ids[[string]$wb.workbookId] = $true }
    foreach ($page in @(Get-Array $st.pages | Where-Object { $ids.ContainsKey([string]$_.workbookId) })) { Set-NoteProperty $page 'contentPdf' $countedPdfRel }
    return $null
})
$realPageCount = ${function:Get-PdfPageCount}
$Script:PageCountCalls = 0
function Get-PdfPageCount([string]$PdfPath) { $Script:PageCountCalls++; return 2 }
try {
    $unrecorded = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' $countedId
    Assert-OutputTest ($Script:PageCountCalls -eq 1 -and [int]$unrecorded.pageCount -eq 2) "Without a recorded page count the converted PDF was not counted (calls=$($Script:PageCountCalls), pages=$($unrecorded.pageCount))."
    [void](Update-StructureLocked 'ja' {
        param($st)
        foreach ($artifact in @(Get-Array (Get-DataProperty $st 'artifacts' @()) | Where-Object { [string](Get-DataProperty $_ 'relativePdfPath' '') -eq $countedPdfRel })) { Set-NoteProperty $artifact 'pageCount' 3 }
        return $null
    })
    $recordedArtifacts = @(Get-Array (Get-DataProperty (Get-Structure 'ja') 'artifacts' @()) | Where-Object { [string](Get-DataProperty $_ 'relativePdfPath' '') -eq $countedPdfRel -and (Get-IntDataProperty $_ 'pageCount' 0) -eq 3 })
    Assert-OutputTest ($recordedArtifacts.Count -ge 1) 'The recorded artifact page count was not stored; the recorded-count case is not exercised.'
    $Script:PageCountCalls = 0
    $recorded = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' $countedId
    Assert-OutputTest ($Script:PageCountCalls -eq 0) 'Get-PdfPageCount was called although the artifact page count was recorded.'
    Assert-OutputTest ([int]$recorded.pageCount -eq 3) "The recorded page count was not used (pages=$($recorded.pageCount))."
} finally {
    Set-Item Function:\Get-PdfPageCount $realPageCount
}

# ---- 性能: 入力スナップショットは要求（ジョブ）の中で structure の版ごとに使い回す ----
$realSnapshot = ${function:Get-FinalBuildInputSnapshot}
$Script:SnapshotCalls = 0
function Get-FinalBuildInputSnapshot { $Script:SnapshotCalls++; return (& $realSnapshot @args) }
$Script:FinalSnapshotCache = @{}
try {
    $view = Get-StructureWithStamp 'ja'
    Assert-OutputTest (-not [string]::IsNullOrWhiteSpace([string]$view.stamp)) 'Get-StructureWithStamp returned no stamp; the cache case is not exercised.'
    $first = Get-FinalBuildReadiness $view.structure 'ja' 'ja-main' $countedId $true ([string]$view.stamp)
    $second = Get-FinalBuildReadiness $view.structure 'ja' 'ja-main' $countedId $false ([string]$view.stamp)
    Assert-OutputTest ($Script:SnapshotCalls -eq 1) "The readiness of the same structure version rebuilt the input snapshot (calls=$($Script:SnapshotCalls))."
    Assert-OutputTest ([string]$first.currentFingerprint -eq [string]$second.currentFingerprint) 'A cached snapshot returned a different fingerprint.'
    # 版を渡さない呼び出し（mutation の中など）は使い回さない。
    [void](Get-FinalBuildReadiness $view.structure 'ja' 'ja-main' $countedId)
    Assert-OutputTest ($Script:SnapshotCalls -eq 2) 'A readiness call without a structure version used the snapshot cache.'
    # 返したスナップショットに書き足しても、キャッシュの実体は汚れない。
    $copy = Get-CachedFinalBuildInputSnapshot $view.structure 'ja' 'ja-main' $countedId ([string]$view.stamp)
    Set-NoteProperty $copy 'sourceWorkbooks' @('x')
    $again = Get-CachedFinalBuildInputSnapshot $view.structure 'ja' 'ja-main' $countedId ([string]$view.stamp)
    Assert-OutputTest (-not $again.Contains('sourceWorkbooks')) 'A caller-side addition leaked into the cached snapshot.'
    # 変換PDFが差し替わったら（structure はそのまま）引き直す。
    $calls = $Script:SnapshotCalls
    [IO.File]::AppendAllText($countedPdf, '-changed')
    $changedPdf = Get-FinalBuildReadiness $view.structure 'ja' 'ja-main' $countedId $true ([string]$view.stamp)
    Assert-OutputTest ($Script:SnapshotCalls -eq $calls + 1) 'A replaced converted PDF did not invalidate the cached snapshot.'
    Assert-OutputTest ([string]$changedPdf.currentFingerprint -ne [string]$first.currentFingerprint) 'The fingerprint did not follow the replaced converted PDF.'
    # structure を保存したら古い結果は使わない。
    $calls = $Script:SnapshotCalls
    [void](Update-StructureLocked 'ja' { param($st) return $null })
    $view2 = Get-StructureWithStamp 'ja'
    [void](Get-FinalBuildReadiness $view2.structure 'ja' 'ja-main' $countedId $true ([string]$view2.stamp))
    Assert-OutputTest ($Script:SnapshotCalls -eq $calls + 1) 'A saved structure still served the cached snapshot.'
} finally {
    $Script:FinalSnapshotCache = $null
    Set-Item Function:\Get-FinalBuildInputSnapshot $realSnapshot
}

# ---- 性能: 出力前の既定番号で何も変わらなければ structure.json を書き直さない ---------
$numberingFirst = Update-StructureLocked 'ja' { param($st) return [ordered]@{ changed = (Update-DefaultNumberingPerVolume 'ja' $st $countedId) } }
$stampBefore = Get-StructureFileStamp 'ja'
$numberingAgain = Update-StructureLocked 'ja' {
    param($st)
    $changed = Update-DefaultNumberingPerVolume 'ja' $st $countedId
    return [ordered]@{ changed = $changed; __noStructureWrite = (-not $changed) }
}
Assert-OutputTest ([bool]$numberingFirst.changed) 'The first default numbering pass reported no change; the no-write case is not exercised.'
Assert-OutputTest (-not [bool]$numberingAgain.changed) 'Re-applying the default numbering reported a change.'
Assert-OutputTest ((Get-StructureFileStamp 'ja') -eq $stampBefore) 'An unchanged default numbering pass still rewrote structure.json.'
Assert-OutputTest (-not $numberingAgain.Contains('__noStructureWrite')) 'The __noStructureWrite marker leaked to the caller.'

# ---- 作成中に原稿・入力が変わっても、作成開始時のスナップショットで最後まで作る ----------
# 組版（Java）と入力スナップショットだけを差し替え、実物の Build-DocumentPackPdf と
# Invoke-FinalBuildTransaction を通す。組版の最中に「変換し直し」を起こす（元の変換PDFを
# 世代整理で消し、入力の指紋を変える）。以前はここで「PDF作成中に…変更されました」と失敗した。
$midWorkspace = Get-WorkspacePath 'ja'
$midPdfRel = 'content-pdf\selfcheck-mid\v1\Sheet10.pdf'
$midPdf = Join-Path $midWorkspace $midPdfRel
New-Item -ItemType Directory -Path (Split-Path -Parent $midPdf) -Force | Out-Null
$Script:MidInputVersion = 'v1'
$Script:MidOutputName = 'mid-build.pdf'
$Script:MidComposeHook = $null
$Script:MidComposedSources = @()
$realMidSnapshot = ${function:Get-FinalBuildInputSnapshot}
$realMidScan = ${function:Scan-Updates}
$realMidJava = ${function:Resolve-JavaExe}
$realMidNative = ${function:Invoke-NativeCapture}
$realMidAppRoot = $Script:AppRoot
# 呼ばれた時点の変換PDFのサイズ・更新時刻を記録する。指紋は $Script:MidInputVersion で決める。
function Get-FinalBuildInputSnapshot($Structure, [string]$Language, [string]$Volume, [string]$PackId) {
    $info = New-Object IO.FileInfo $midPdf
    $size = if ($info.Exists) { [int64]$info.Length } else { 0L }
    $ticks = if ($info.Exists) { [int64]$info.LastWriteTimeUtc.Ticks } else { 0L }
    $fpPage = [ordered]@{ pageId='mid-p1'; contentPdf=$midPdfRel; contentPdfSize=$size; contentPdfLastWriteUtcTicks=$ticks }
    $entry = [ordered]@{ pageId='mid-p1'; itemId='mid-p1'; sourceId='selfcheck-mid'; title='Sheet10'; sourcePdf=$midPdf; sourcePageStart=1; sourcePageEnd=1; sourcePageCount=1; numberingMode='visible' }
    return [ordered]@{ language=$Language; packId=$PackId; volume=$Volume; pages=@(); manifestPages=@($entry); physicalPages=@(); document=[ordered]@{}; blockers=@(); itemCount=1; pageCount=1; projectId=''; outputFileName=$Script:MidOutputName; fingerprint=('fp|mid|' + $Script:MidInputVersion); fingerprintInput=[ordered]@{ pages=@($fpPage) } }
}
function Scan-Updates { return $null }
function Resolve-JavaExe { return 'java' }
# 組版の代わり: manifest が指す変換PDFを読んでつなげる。読む前に $Script:MidComposeHook を走らせる。
function Invoke-NativeCapture([string]$FilePath, [string[]]$ArgumentList, [int]$TimeoutSeconds = 120) {
    $manifestArg = [string]$ArgumentList[[Array]::IndexOf($ArgumentList, '--manifest') + 1]
    $m = Get-Content -LiteralPath $manifestArg -Raw -Encoding UTF8 | ConvertFrom-Json
    $Script:MidComposedSources = @($m.pages | ForEach-Object { [string]$_.sourcePdf })
    if ($null -ne $Script:MidComposeHook) { & $Script:MidComposeHook }
    $text = (@($Script:MidComposedSources | ForEach-Object { [IO.File]::ReadAllText($_) }) -join '|')
    [IO.File]::WriteAllText([string]$m.outputPdf, $text)
    return [pscustomobject]@{ exitCode = 0; text = 'composed'; timedOut = $false }
}
$fakeAppRoot = Join-Path $Script:LocalConfigRoot 'fake-app'
New-Item -ItemType Directory -Path (Join-Path $fakeAppRoot 'lib\pdfbox') -Force | Out-Null
foreach ($jar in @('ReportPdfComposer.jar','pdfbox-app.jar')) { [IO.File]::WriteAllText((Join-Path (Join-Path $fakeAppRoot 'lib\pdfbox') $jar), 'jar') }
$Script:AppRoot = $fakeAppRoot
$changedDuringBuildHook = { Remove-Item -LiteralPath $midPdf -Force; $Script:MidInputVersion = ('changed-' + [Guid]::NewGuid().ToString('N')) }
function Get-InputCopyDirs { return @(Get-ChildItem -LiteralPath (Join-Path $midWorkspace 'state\tmp') -Directory -Filter 'final-inputs-*' -ErrorAction SilentlyContinue) }
function Assert-BuiltFromStartSnapshot($Result, [string]$PackId, [string]$Expected, [string]$Label) {
    Assert-OutputTest ([bool](Get-DataProperty $Result 'inputChangedDuringBuild' $false)) "$Label did not report that the inputs changed during the build."
    $outPath = [string](Get-DataProperty $Result 'outputPdf' '')
    Assert-OutputTest ((Test-Path -LiteralPath $outPath) -and [IO.File]::ReadAllText($outPath) -eq $Expected) "$Label did not keep the PDF composed from the start-of-build inputs."
    Assert-OutputTest ($Script:MidComposedSources.Count -eq 1 -and $Script:MidComposedSources[0] -ne $midPdf -and $Script:MidComposedSources[0] -like '*final-inputs-*') "$Label composed from the live converted PDF instead of the start-of-build copy: $($Script:MidComposedSources -join ', ')"
    Assert-OutputTest (-not (Test-Path -LiteralPath $Script:MidComposedSources[0])) "$Label left the start-of-build copy behind."
    $state = Get-PackOutputState (Get-Structure 'ja') 'ja' $PackId 'ja-main' $false
    Assert-OutputTest ([string](Get-DataProperty $state 'builtFingerprint' '') -eq 'fp|mid|v1') "$Label did not record the start-of-build fingerprint."
    Assert-OutputTest ([string](Get-DataProperty $state 'status' '') -eq 'needs-rebuild') "$Label did not mark the output as needs-rebuild."
    $ready = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $PackId
    Assert-OutputTest ([string]$ready.displayState -eq 'needs-rebuild') "$Label is not shown as needs-rebuild (displayState=$($ready.displayState))."
    $firstReason = @(Get-Array $ready.staleReasons) | Select-Object -First 1
    Assert-OutputTest ([string](Get-DataProperty $firstReason 'type' '') -eq 'changed-during-build') "$Label did not explain that the inputs changed during the build."
}
try {
    # 1. カスタム一式（Build-DocumentPackPdf）。作成中に変わっても失敗せず、開始時の内容で確定する。
    $midPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='作成中に変わる一式' })
    $midId = [string]$midPack.packId
    Add-TestPages $midId @('ja-main')
    # 作成前からあった理由（前回の出力に対するもの）は、今回の出力で解消済みなので残さない。
    [void](Update-StructureLocked 'ja' {
        param($st)
        $state = Get-PackOutputState $st 'ja' $midId 'ja-main' $true
        Set-NoteProperty $state 'builtFingerprint' 'fp|older'
        Set-NoteProperty $state 'staleReasons' @([ordered]@{ type='unregister'; at='2000-01-01T00:00:00+09:00'; detail='前回の出力に対する理由'; count=1 })
        return $null
    })
    [IO.File]::WriteAllText($midPdf, '%PDF-start-of-build')
    $Script:MidInputVersion = 'v1'
    $Script:MidComposeHook = $changedDuringBuildHook
    $midResult = Build-DocumentPackPdf 'ja' $midId 'main'
    Assert-BuiltFromStartSnapshot $midResult $midId '%PDF-start-of-build' 'A custom pack build'
    $midState = Get-PackOutputState (Get-Structure 'ja') 'ja' $midId 'ja-main' $false
    Assert-OutputTest (@(Get-Array (Get-DataProperty $midState 'staleReasons' @()) | Where-Object { [string](Get-DataProperty $_ 'type' '') -eq 'unregister' }).Count -eq 0) 'A stale reason from before the build survived the new output.'
    $savedManifest = Read-JsonFile (Join-Path $midWorkspace ("exports\manifest_{0}_main.json" -f $midId)) $null
    Assert-OutputTest ([string]@($savedManifest.pages)[0].sourcePdf -eq $midPdf) 'The saved manifest points at the temporary copy instead of the converted PDF.'
    Assert-OutputTest (@(Get-InputCopyDirs).Count -eq 0) 'The start-of-build copy folder was left behind.'

    # 変わらなければ今までどおり built。前回の changed-during-build は残さない。
    [IO.File]::WriteAllText($midPdf, '%PDF-second')
    $Script:MidInputVersion = 'v1'
    $Script:MidComposeHook = $null
    $steady = Build-DocumentPackPdf 'ja' $midId 'main'
    Assert-OutputTest (-not [bool](Get-DataProperty $steady 'inputChangedDuringBuild' $true)) 'An unchanged build reported changed inputs.'
    $steadyReady = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $midId
    Assert-OutputTest ([string]$steadyReady.displayState -eq 'built' -and @($steadyReady.staleReasons).Count -eq 0) "An unchanged build is not shown as built (displayState=$($steadyReady.displayState))."

    # 2. 組み込み一式の一括出力（Invoke-FinalBuildTransaction）でも同じ。
    $Script:MidOutputName = 'mid-build-ecm.pdf'
    [IO.File]::WriteAllText($midPdf, '%PDF-start-of-transaction')
    $Script:MidInputVersion = 'v1'
    $Script:MidComposeHook = $changedDuringBuildHook
    $txResult = Invoke-FinalBuildTransaction 'ja' 'ecm' @('ja-main')
    Assert-OutputTest (@($txResult.built).Count -eq 1) 'The final build transaction did not complete when the inputs changed during the build.'
    Assert-BuiltFromStartSnapshot @($txResult.built)[0] (Get-BuiltinPackId 'ecm') '%PDF-start-of-transaction' 'A built-in pack build'
    Assert-OutputTest (@(Get-InputCopyDirs).Count -eq 0) 'The start-of-build copy folder of the transaction was left behind.'

    # 3. 合成する変換PDFそのものが開始時点で消えた・差し替わったときだけ失敗する。
    $Script:MidOutputName = 'mid-build.pdf'
    $Script:MidComposeHook = $null
    [IO.File]::WriteAllText($midPdf, '%PDF-probe')
    $probeSnapshot = Get-FinalBuildInputSnapshot $null 'ja' 'ja-main' $midId
    [IO.File]::AppendAllText($midPdf, '-replaced')
    Assert-Throws { Copy-FinalBuildInputPdfs 'ja' $probeSnapshot (Join-Path $Script:LocalConfigRoot 'copy-probe-replaced') } '変換PDF' 'A converted PDF replaced after the snapshot was copied as if it were the start-of-build version.'
    $probeSnapshot = Get-FinalBuildInputSnapshot $null 'ja' 'ja-main' $midId
    Remove-Item -LiteralPath $midPdf -Force
    Assert-Throws { Copy-FinalBuildInputPdfs 'ja' $probeSnapshot (Join-Path $Script:LocalConfigRoot 'copy-probe-missing') } '変換PDF' 'A missing converted PDF did not stop the build.'
    $beforeMissing = [string](Get-DataProperty (Get-PackOutputState (Get-Structure 'ja') 'ja' $midId 'ja-main' $false) 'builtFingerprint' '')
    Assert-Throws { Build-DocumentPackPdf 'ja' $midId 'main' } '変換PDF' 'A build whose converted PDF is missing did not fail.'
    Assert-OutputTest ([string](Get-DataProperty (Get-PackOutputState (Get-Structure 'ja') 'ja' $midId 'ja-main' $false) 'builtFingerprint' '') -eq $beforeMissing) 'A failed build changed the recorded output.'
    Assert-OutputTest (@(Get-InputCopyDirs).Count -eq 0) 'A failed build left the start-of-build copy folder behind.'
} finally {
    $Script:AppRoot = $realMidAppRoot
    Set-Item Function:\Get-FinalBuildInputSnapshot $realMidSnapshot
    Set-Item Function:\Scan-Updates $realMidScan
    Set-Item Function:\Resolve-JavaExe $realMidJava
    Set-Item Function:\Invoke-NativeCapture $realMidNative
}

# ---- 共有用フォルダーへのまとめてコピー（publish-all）と出力PDFの照合（M-2） ----------
# 原稿のPDFを置かずに済むよう、入力スナップショットだけを差し替える。指紋は出力先ごとに固定。
function Get-FinalBuildInputSnapshot($Structure, [string]$Language, [string]$Volume, [string]$PackId) {
    $ids = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks | Where-Object { Test-WorkbookPack $_ $PackId })) { $ids[[string]$wb.workbookId] = $true }
    $count = @(Get-Array $Structure.pages | Where-Object { $ids.ContainsKey([string]$_.workbookId) -and [string]$_.volume -eq $Volume -and $_.enabled -ne $false }).Count
    return [pscustomobject]@{ fingerprint="fp|$PackId|$Volume"; blockers=@(); pageCount=$count }
}
# 出力ジョブが確定したときと同じ記録を置く。
function Set-BuiltOutput([string]$PackId, [string]$Volume, [string]$Content, [bool]$RecordIntegrity = $true) {
    $path = Get-PlannedOutputPath $PackId $Volume
    [IO.File]::WriteAllText($path, $Content)
    [void](Update-StructureLocked 'ja' {
        param($st)
        $state = Get-PackOutputState $st 'ja' $PackId $Volume $true
        Set-NoteProperty $state 'outputPdf' $path
        Set-NoteProperty $state 'builtFingerprint' ([string](Get-FinalBuildInputSnapshot $st 'ja' $Volume $PackId).fingerprint)
        Set-NoteProperty $state 'outputIntegrity' $(if ($RecordIntegrity) { Get-OutputPdfIntegrity $path } else { $null })
        return $null
    })
    return $path
}
function Get-SharedEntries { return @(Get-ChildItem -LiteralPath $submissionDir -Force | ForEach-Object { $_.Name }) }
$publishPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='発行の一式' })
$publishId = [string]$publishPack.packId
Add-TestPages $publishId @('ja-main','ja-appendix')
$mainPdf = Set-BuiltOutput $publishId 'ja-main' '%PDF-main'
$appendixPdf = Set-BuiltOutput $publishId 'ja-appendix' '%PDF-appendix'
$published = Publish-DocumentPackAllPdfsToShared 'ja' $publishId
Assert-OutputTest (@($published.files).Count -eq 2) 'publish-all did not copy both targets with pages.'
foreach ($file in @($published.files)) {
    $local = if ([string]$file.targetId -eq 'main') { $mainPdf } else { $appendixPdf }
    Assert-OutputTest ((New-Sha256 ([string]$file.sharedPath)) -eq (New-Sha256 $local)) "publish-all copied a different file for $($file.targetId)."
    Assert-OutputTest ([IO.Path]::GetDirectoryName([string]$file.sharedPath) -eq [string]$published.sharedFolder) 'publish-all split the targets across folders.'
}
Assert-OutputTest (@(Get-SharedEntries | Where-Object { $_ -like '.publishing-*' }).Count -eq 0) 'A staging folder was left behind after publishing.'
# 発行フォルダー名は「日時_一式名」だけ。誰が発行したかは publishedBy に残す。
$publisher = Get-SafePublishUserName
Assert-OutputTest ([string]$published.folderName -match ('^\d{4}_\d{6}_' + [regex]::Escape((Get-SafePublishPackName '発行の一式')) + '$')) "The published folder name is not 'MMdd_HHmmss_<pack>': $($published.folderName)"
Assert-OutputTest (-not ([string]$published.folderName).EndsWith('_' + $publisher)) 'The published folder name still contains the Windows user name.'
Assert-OutputTest ([IO.Path]::GetFileName([string]$published.sharedFolder) -eq [string]$published.folderName) 'folderName does not match the published folder.'
Assert-OutputTest ([string]$published.publishedBy -eq $publisher) 'publishedBy no longer records the Windows user name.'
$sharedBefore = @(Get-SharedEntries)

# 1冊でも最新でなければ何もコピーしない。
[void](Update-StructureLocked 'ja' { param($st) Set-NoteProperty (Get-PackOutputState $st 'ja' $publishId 'ja-appendix' $true) 'builtFingerprint' 'old'; return $null })
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '最新ではありません' 'publish-all copied while one target was stale.'
Assert-OutputTest ((@(Get-SharedEntries) -join '|') -eq ($sharedBefore -join '|')) 'A rejected publish-all still created something in the shared folder.'
$appendixPdf = Set-BuiltOutput $publishId 'ja-appendix' '%PDF-appendix'

# 出力した後に書き換えられたPDFは「最新」と表示せず、発行もしない。
[IO.File]::AppendAllText($mainPdf, 'edited')
$changed = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $publishId
Assert-OutputTest ([string]$changed.displayState -eq 'needs-rebuild' -and @($changed.staleReasons | Where-Object { [string]$_.type -eq 'output-changed' }).Count -eq 1) 'A PDF changed after output was still shown as latest.'
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '変更または上書き' 'A PDF changed after output was published.'
# サイズも更新日時も同じまま中身だけ違うものは、発行前のハッシュで止める。
$mainPdf = Set-BuiltOutput $publishId 'ja-main' '%PDF-main'
$stamp = (Get-Item -LiteralPath $mainPdf).LastWriteTimeUtc
[IO.File]::WriteAllText($mainPdf, '%PDF-MAIN')
(Get-Item -LiteralPath $mainPdf).LastWriteTimeUtc = $stamp
Assert-OutputTest ([string](Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $publishId).displayState -eq 'built') 'The same-size case did not reach the hash check.'
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '中身' 'A PDF whose content differs from the recorded hash was published.'
# 照合の記録が無い（以前の版で出力した）PDFは発行しない。
$mainPdf = Set-BuiltOutput $publishId 'ja-main' '%PDF-main' $false
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '記録がありません' 'A PDF without an output record was published.'
Assert-OutputTest ((@(Get-SharedEntries) -join '|') -eq ($sharedBefore -join '|')) 'A rejected publish still created something in the shared folder.'
$mainPdf = Set-BuiltOutput $publishId 'ja-main' '%PDF-main'

# ページの無い出力先は対象外（補足が空なら本体だけを発行する）。
$mainOnlyPublish = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='本体だけの発行' })
Add-TestPages ([string]$mainOnlyPublish.packId) @('ja-main')
[void](Set-BuiltOutput ([string]$mainOnlyPublish.packId) 'ja-main' '%PDF-only')
$onlyMain = Publish-DocumentPackAllPdfsToShared 'ja' ([string]$mainOnlyPublish.packId)
Assert-OutputTest (@($onlyMain.files).Count -eq 1 -and [string]@($onlyMain.files)[0].targetId -eq 'main') 'publish-all did not skip the target without pages.'

# ---- M-3: 出力中は共有用フォルダーへコピーしない ------------------------------------
$sharedBefore = @(Get-SharedEntries)
# 出力が同じ出力先のPDFを置き換えている間（出力のロックを持っている間）はコピーしない。
$outputLock = Join-Path (Get-WorkspacePath 'ja') ("locks\pack-output_{0}_main.lock" -f $publishId)
New-Item -ItemType Directory -Path (Split-Path -Parent $outputLock) -Force | Out-Null
$held = [IO.File]::Open($outputLock, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
try { Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } 'ロック中' 'Publishing ran while the output lock was held.' }
finally { $held.Dispose() }
# 出力ジョブが動いている間はコピーしない。
$jobId = 'final_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '_' + [Guid]::NewGuid().ToString('N').Substring(0,8)
$jobStatus = Join-Path (Get-RenderJobDir 'ja') "$jobId.status.json"
Write-JsonFile $jobStatus ([ordered]@{ ok=$true; jobId=$jobId; kind='final-build'; status='running'; processId=$PID; startedAt=(New-NowIso); updatedAt=(New-NowIso) })
try { Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '出力中' 'Publishing ran while an output job was active.' }
finally { Remove-Item -LiteralPath $jobStatus -Force }
Assert-OutputTest ((@(Get-SharedEntries) -join '|') -eq ($sharedBefore -join '|')) 'A publish refused during output still created something in the shared folder.'
$afterJob = Publish-DocumentPackAllPdfsToShared 'ja' $publishId
Assert-OutputTest (@($afterJob.files).Count -eq 2) 'Publishing did not resume after the output finished.'

# ---- M-7: 発行の作業用フォルダーは隠し、残骸は次の発行の前に消す -----------------
$me = Get-SafePublishUserName
function New-StagingLeftover([string]$Name, [double]$HoursOld) {
    $dir = New-Item -ItemType Directory -Path (Join-Path $submissionDir $Name) -Force
    [IO.File]::WriteAllText((Join-Path $dir.FullName 'half.pdf'), '%PDF-half')
    $dir.LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddHours(-$HoursOld)
    return $Name
}
$guid = { [Guid]::NewGuid().ToString('N') }
$ownOld = New-StagingLeftover (".publishing-$me-" + (& $guid)) 2
$ownRecent = New-StagingLeftover (".publishing-$me-" + (& $guid)) 0
$otherOld = New-StagingLeftover (".publishing-$me-x-" + (& $guid)) 48
$legacyOld = New-StagingLeftover ('.publishing-' + (& $guid)) 48
$legacyRecent = New-StagingLeftover ('.publishing-' + (& $guid)) 2
# コピーの瞬間に作業用フォルダーの属性を記録する。
$Script:StagingAttributes = @()
function Copy-Item {
    param([string]$LiteralPath, [string]$Destination, [switch]$Force)
    $Script:StagingAttributes += (Get-Item -LiteralPath (Split-Path -Parent $Destination) -Force).Attributes
    Microsoft.PowerShell.Management\Copy-Item -LiteralPath $LiteralPath -Destination $Destination -Force:$Force
}
[void](Publish-DocumentPackAllPdfsToShared 'ja' $publishId)
Remove-Item Function:\Copy-Item
$left = @(Get-SharedEntries)
Assert-OutputTest (-not ($left -contains $ownOld)) 'An old staging folder left by this user was not removed.'
Assert-OutputTest (-not ($left -contains $legacyOld)) 'An old staging folder in the previous naming was not removed.'
Assert-OutputTest ($left -contains $ownRecent -and $left -contains $legacyRecent) 'A staging folder that may still be in use was removed.'
Assert-OutputTest ($left -contains $otherOld) "Another user's staging folder was removed."
Assert-OutputTest ($Script:StagingAttributes.Count -eq 2) 'The staging copy was not observed.'
if ($env:OS -eq 'Windows_NT') {
    Assert-OutputTest (@($Script:StagingAttributes | Where-Object { ($_ -band [IO.FileAttributes]::Hidden) -eq 0 }).Count -eq 0) 'The staging folder was visible in the shared folder while copying.'
}

# ---- 隠し属性は作業用フォルダーだけ。発行フォルダーとPDFには残さない --------------------
# 改名しても属性はフォルダーに残るため、以前は発行フォルダーごと見えなくなっていた。
# 属性そのものは Windows でしか立たないので、どの環境でも「改名後の発行フォルダーと
# 中のPDFを1つずつ見えるようにしたか」を呼び出しで確かめ、Windows では実際の属性も見る。
$realClearHidden = ${function:Clear-PublishHiddenAttribute}
$Script:ClearedPaths = @()
$Script:FailClearOnPublishedFolder = $false
$submissionFull = [IO.Path]::GetFullPath($submissionDir)
function Clear-PublishHiddenAttribute([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    $Script:ClearedPaths += $full
    # 改名した後の発行フォルダー（共有ルート直下で「.」で始まらないもの）だけ失敗させる。
    if ($Script:FailClearOnPublishedFolder -and [IO.Path]::GetDirectoryName($full) -eq $submissionFull -and [IO.Path]::GetFileName($full) -notlike '.*') { return $false }
    return (& $realClearHidden $Path)
}
$onWindows = $env:OS -eq 'Windows_NT'
# Get-Item は -Force が無いと隠しファイルを返さない。照合が隠されたPDFを「見つかりません」に
# しないこと。Windows では隠し属性で、それ以外では「.」で始まる名前で隠す。
$hiddenProbe = Join-Path (Split-Path -Parent $mainPdf) '.hidden-integrity-probe.pdf'
[IO.File]::WriteAllText($hiddenProbe, '%PDF-hidden')
if ($onWindows) { (Get-Item -LiteralPath $hiddenProbe -Force).Attributes = [IO.FileAttributes]::Hidden }
try {
    $hiddenState = [pscustomobject]@{ outputIntegrity = (Get-OutputPdfIntegrity $hiddenProbe) }
    $hiddenProblem = Get-OutputPdfIntegrityProblem $hiddenState $hiddenProbe $true
    Assert-OutputTest ($hiddenProblem -eq '') "A hidden output PDF failed the integrity check: $hiddenProblem"
} finally {
    if ($onWindows) { (Get-Item -LiteralPath $hiddenProbe -Force).Attributes = [IO.FileAttributes]::Normal }
    Remove-Item -LiteralPath $hiddenProbe -Force
}
# 出力したPDFが隠されていても、Copy-Item で写った属性を発行先に残さない。
if ($onWindows) { (Get-Item -LiteralPath $mainPdf -Force).Attributes = (Get-Item -LiteralPath $mainPdf -Force).Attributes -bor [IO.FileAttributes]::Hidden }
try { $visible = Publish-DocumentPackAllPdfsToShared 'ja' $publishId }
finally { if ($onWindows) { (Get-Item -LiteralPath $mainPdf -Force).Attributes = [IO.FileAttributes]::Normal } }
$visibleDir = [IO.Path]::GetFullPath([string]$visible.sharedFolder)
Assert-OutputTest ([IO.Path]::GetFileName($visibleDir) -notlike '.*') 'The published folder kept a staging-style dot name.'
Assert-OutputTest ($Script:ClearedPaths -contains $visibleDir) 'The hidden attribute was not cleared on the published folder after the rename.'
foreach ($file in @($visible.files)) {
    Assert-OutputTest ($Script:ClearedPaths -contains [IO.Path]::GetFullPath([string]$file.sharedPath)) "The hidden attribute was not checked on the published PDF $($file.fileName)."
    Assert-OutputTest (Test-Path -LiteralPath ([string]$file.sharedPath) -PathType Leaf) "The published PDF $($file.fileName) is missing."
}
if ($onWindows) {
    $hiddenItems = @(@(Get-Item -LiteralPath $visibleDir -Force) + @(Get-ChildItem -LiteralPath $visibleDir -Force -Recurse) | Where-Object { ($_.Attributes -band [IO.FileAttributes]::Hidden) -ne 0 })
    Assert-OutputTest ($hiddenItems.Count -eq 0) ('The published folder or its PDFs are hidden: ' + (($hiddenItems | ForEach-Object { $_.FullName }) -join ', '))
    Assert-OutputTest (@(Get-ChildItem -LiteralPath $submissionDir -Directory | Where-Object { $_.FullName -eq $visibleDir }).Count -eq 1) 'The published folder is not listed without -Force.'
}
# 隠し属性を外せなければ、見えない発行フォルダーを残さずに失敗にする。
$sharedBefore = @(Get-SharedEntries)
$Script:FailClearOnPublishedFolder = $true
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '隠し属性を外せません' 'A published folder whose hidden attribute could not be cleared was reported as success.'
Assert-OutputTest ((@(Get-SharedEntries) -join '|') -eq ($sharedBefore -join '|')) 'A published folder that stayed hidden was left in the shared folder.'
$Script:FailClearOnPublishedFolder = $false
Set-Item Function:\Clear-PublishHiddenAttribute $realClearHidden

Write-Output 'output-publish selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
