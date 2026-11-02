param()

# 提出用PDF一式の変更点（出力先単位）の計算を、server.ps1 の実物の関数で確かめる。
# Java・Office は使わない。ページの対応付け（Compare-PackArchivePageMaps）は副作用の無い
# 関数なので、manifest・metadata を直接組んで渡す。既定の組み合わせ（Resolve-PackChangePair）と
# GET /api/v2/outputs/changes の本体（Get-PackChangeSummary）は、出力済みのPDFとアーカイブを
# 直接置いて確かめる（output-publish-selfcheck.ps1 と同じやり方）。
$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-pack-change-' + [Guid]::NewGuid().ToString('N'))
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
function Assert-ChangeTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Assert-Throws([scriptblock]$Action, [string]$ExpectedFragment, [string]$Message) {
    $caught = $null
    try { & $Action } catch { $caught = $_.Exception.Message }
    if ($null -eq $caught) { throw "$Message (no error was raised)" }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedFragment) -and -not $caught.Contains($ExpectedFragment)) { throw "$Message (unexpected error: $caught)" }
}

# ---- 組み立て用 ---------------------------------------------------------------
# Get-FinalBuildInputSnapshot と同じ形の physicalPages を作る（奇数＝右、台紙＝floor(頁/2)）。
# ここはテストの入力を作るだけで、変更点の計算は左右を決めない（記録された値を比べるだけ）。
function New-TestItem([string]$Item, [int]$Parts = 1, [string]$Source = 'wb_a', [string]$Title = '', [string]$Key = '', [int]$Start = 1) {
    $resolvedTitle = $(if ($Title) { $Title } else { "シート$Item" })
    $resolvedKey = $(if ($Key) { $Key } else { "worksheet:シート$Item" })
    return [ordered]@{ item = $Item; parts = $Parts; source = $Source; key = $resolvedKey; title = $resolvedTitle; start = $Start }
}
function New-TestManifest([object[]]$Items, [bool]$WithPageMap = $true) {
    $physical = @()
    $pages = @()
    $n = 0
    foreach ($it in $Items) {
        $parts = [int]$it.parts
        $start = [int]$it.start
        $pages += [ordered]@{ pageId = "p-$($it.item)"; itemId = [string]$it.item; sourceId = [string]$it.source; sourceKey = [string]$it.key; title = [string]$it.title; sectionTitle = "原稿 $($it.source)"; sourcePageStart = $start; sourcePageEnd = ($start + $parts - 1) }
        for ($i = 1; $i -le $parts; $i++) {
            $n++
            $physical += [ordered]@{ outputPageNumber = $n; side = $(if ($n % 2 -eq 1) { 'right' } else { 'left' }); spreadIndex = [int][Math]::Floor($n / 2); pagePartIndex = $i; pagePartCount = $parts; kind = 'content'; itemId = [string]$it.item; pageId = "p-$($it.item)"; sourceId = [string]$it.source; sourceKey = [string]$it.key; sourcePdfPageNumber = ($start + $i - 1); title = [string]$it.title }
        }
    }
    $manifest = [ordered]@{ schemaVersion = 3; pages = @($pages) }
    if ($WithPageMap) { $manifest.physicalPages = @($physical) }
    return $manifest
}
# 原稿ID → 'snapshotId|versionId'
function New-TestMetadata([hashtable]$Versions) {
    $sourceWorkbooks = @()
    foreach ($k in @($Versions.Keys | Sort-Object)) {
        $parts = ([string]$Versions[$k]).Split('|')
        $sourceWorkbooks += [ordered]@{ workbookId = $k; fileName = "$k.xlsx"; snapshotId = $parts[0]; versionId = $parts[1] }
    }
    return [ordered]@{ schemaVersion = 2; sourceWorkbooks = @($sourceWorkbooks) }
}
# visual-hashes.json と同じ形。シート名 → ページごとのハッシュ。
function New-TestHashes([hashtable]$Sheets) {
    $list = @()
    foreach ($name in @($Sheets.Keys)) { $list += [ordered]@{ sheetName = [string]$name; pageCount = @($Sheets[$name]).Count; pageHashes = @($Sheets[$name]) } }
    return [ordered]@{ schemaVersion = 1; sheets = @($list) }
}
function Get-RowAt($Comparison, [int]$ToPage) {
    $rows = @($Comparison.pages | Where-Object { $null -ne $_.toPage -and [int]$_.toPage -eq $ToPage })
    Assert-ChangeTest ($rows.Count -eq 1) "Expected exactly one row for output page $ToPage (rows=$($rows.Count))."
    return $rows[0]
}
function Get-Kinds($Row) { return (@($Row.kinds) -join ',') }

$v1 = @{ wb_a = 's1|v1'; wb_b = 's1|v1' }

# ---- 同じ版・同じ位置は「変わっていない」 -----------------------------------------
# 版の照合を外すと、ハッシュの無いこのテストでは全ページが「内容が変わった可能性あり」になって落ちる。
$base = @((New-TestItem 'A'), (New-TestItem 'B' 2), (New-TestItem 'C'))
$same = Compare-PackArchivePageMaps (New-TestManifest $base) (New-TestMetadata $v1) (New-TestManifest $base) (New-TestMetadata $v1) @{}
Assert-ChangeTest ([int]$same.counts.total -eq 4 -and [int]$same.counts.unchanged -eq 4 -and [int]$same.counts.changed -eq 0) "Identical archives were not all unchanged: $($same.counts | ConvertTo-Json -Compress)"
Assert-ChangeTest ([int]$same.counts.maybeContent -eq 0 -and [bool]$same.complete) 'Same-version pages were sent to the hash comparison.'
Assert-ChangeTest ([bool]$same.hasSides -and [string](Get-RowAt $same 2).toSide -eq 'left' -and [string](Get-RowAt $same 3).fromSide -eq 'right') 'The recorded sides were not carried through.'
Assert-ChangeTest ([string](Get-RowAt $same 1).sourceName -eq '原稿 wb_a' -and [string](Get-RowAt $same 1).sheetName -eq 'シートA') 'The source name or hash sheet name was not resolved.'

# ---- 1ページ増えたとき、後ろは「場所が変わった」で「内容が変わった」にならない ------------
# キーを outputPageNumber に変えると、後ろのページが内容の変化・増減として出て落ちる。
$inserted = @((New-TestItem 'A'), (New-TestItem 'X' 1 'wb_b'), (New-TestItem 'B' 2), (New-TestItem 'C'))
$ins = Compare-PackArchivePageMaps (New-TestManifest $base) (New-TestMetadata $v1) (New-TestManifest $inserted) (New-TestMetadata $v1) @{}
Assert-ChangeTest ([int]$ins.counts.added -eq 1 -and [int]$ins.counts.removed -eq 0) "An inserted page was not counted as one added page: $($ins.counts | ConvertTo-Json -Compress)"
Assert-ChangeTest ([int]$ins.counts.content -eq 0 -and [int]$ins.counts.maybeContent -eq 0) 'Pages pushed back by an insertion were reported as changed content.'
Assert-ChangeTest ([int]$ins.counts.moved -eq 3 -and [int]$ins.counts.shifted -eq 3) "Pages pushed back by an insertion were not reported as shifted: $($ins.counts | ConvertTo-Json -Compress)"
Assert-ChangeTest ((Get-Kinds (Get-RowAt $ins 2)) -eq 'added' -and $null -eq (Get-RowAt $ins 2).fromPage) 'The inserted page row is wrong.'
$pushed = Get-RowAt $ins 3
Assert-ChangeTest ([int]$pushed.fromPage -eq 2 -and [bool]$pushed.shiftOnly -and (Get-Kinds $pushed) -eq 'moved,side') "A pushed-back page was not 'moved' and 'side' only: $(Get-Kinds $pushed)"
Assert-ChangeTest ((@($ins.pages | ForEach-Object { [int]$_.toPage }) -join ',') -eq '1,2,3,4,5') 'The rows are not in the order of the current output pages.'

# ---- 左右の入れ替わりを数える ---------------------------------------------------
# side の比較を消すと落ちる。入れ替えた2ページはずれではない（shiftOnly にならない）。
$four = @((New-TestItem 'A'), (New-TestItem 'B'), (New-TestItem 'C'), (New-TestItem 'D'))
$swappedOrder = @((New-TestItem 'B'), (New-TestItem 'A'), (New-TestItem 'C'), (New-TestItem 'D'))
$swap = Compare-PackArchivePageMaps (New-TestManifest $four) (New-TestMetadata $v1) (New-TestManifest $swappedOrder) (New-TestMetadata $v1) @{}
Assert-ChangeTest ([int]$swap.counts.side -eq 2 -and [int]$swap.counts.moved -eq 2 -and [int]$swap.counts.shifted -eq 0) "Swapped pages were not counted as side swaps: $($swap.counts | ConvertTo-Json -Compress)"
$movedA = Get-RowAt $swap 2
Assert-ChangeTest ([string]$movedA.itemId -eq 'A' -and [string]$movedA.fromSide -eq 'right' -and [string]$movedA.toSide -eq 'left' -and -not [bool]$movedA.shiftOnly) 'A reordered page was reported as a plain shift.'
Assert-ChangeTest ((Get-Kinds (Get-RowAt $swap 3)) -eq 'unchanged') 'A page that kept its place was reported as changed.'

# ---- なくなったページ ------------------------------------------------------------
$removedItems = @((New-TestItem 'A'), (New-TestItem 'C'), (New-TestItem 'D'))
$rem = Compare-PackArchivePageMaps (New-TestManifest $four) (New-TestMetadata $v1) (New-TestManifest $removedItems) (New-TestMetadata $v1) @{}
Assert-ChangeTest ([int]$rem.counts.removed -eq 1 -and [int]$rem.counts.shifted -eq 2 -and [int]$rem.counts.content -eq 0) "A removed page was not counted correctly: $($rem.counts | ConvertTo-Json -Compress)"
# なくなったページは、前回その直前にあったページのすぐ後ろに並ぶ。
Assert-ChangeTest ([string]@($rem.pages)[1].itemId -eq 'B' -and $null -eq @($rem.pages)[1].toPage -and [int]@($rem.pages)[1].fromPage -eq 2) 'The removed page row is not placed where it used to be.'

# ---- 原稿のページ数が 3→4、版の違いは見た目のハッシュで分ける ---------------------
# 余りの扱いを消すと4ページ目が出ずに落ちる。ハッシュの照合を消すと same-look / content が出ずに落ちる。
$hashes = @{}
$hashes[(Get-PackChangeVersionKey 'wb_a' 's1' 'v1')] = New-TestHashes @{ 'シートA' = @('H1','H2','H3') }
$hashes[(Get-PackChangeVersionKey 'wb_a' 's2' 'v2')] = New-TestHashes @{ 'シートA' = @('h1','CHANGED') }
$grow = Compare-PackArchivePageMaps (New-TestManifest @((New-TestItem 'A' 3))) (New-TestMetadata @{ wb_a = 's1|v1' }) (New-TestManifest @((New-TestItem 'A' 4))) (New-TestMetadata @{ wb_a = 's2|v2' }) $hashes
Assert-ChangeTest ((Get-Kinds (Get-RowAt $grow 1)) -eq 'same-look') "A page with the same picture in a new version was not 'same-look': $(Get-Kinds (Get-RowAt $grow 1))"
Assert-ChangeTest ((Get-Kinds (Get-RowAt $grow 2)) -eq 'content') 'A page with a different picture was not reported as changed content.'
Assert-ChangeTest ((Get-Kinds (Get-RowAt $grow 3)) -eq 'maybe-content' -and -not [bool]$grow.complete) 'A page without a hash in the new version was not "maybe changed".'
Assert-ChangeTest ((Get-Kinds (Get-RowAt $grow 4)) -eq 'added' -and [int]$grow.counts.added -eq 1 -and [int]$grow.counts.partCountChanged -eq 1) 'The extra page of a longer source was not reported as added.'
Assert-ChangeTest ([string](Get-RowAt $grow 1).message -like '*前回 3 → 今回 4*' -and [bool](Get-RowAt $grow 4).partCountChanged) 'The page-count change was not explained on the rows.'
# 「くわしく見る」へは前回→今回の版をこの向きで渡す（入れ替えると原稿単位の比較が逆になる）。
$detailRow = Get-RowAt $grow 2
Assert-ChangeTest ([bool]$detailRow.detailAvailable -and [string]$detailRow.fromSnapshotId -eq 's1' -and [string]$detailRow.toSnapshotId -eq 's2' -and [string]$detailRow.toVersionId -eq 'v2') 'The source-level comparison versions are wrong or swapped.'
Assert-ChangeTest (-not [bool](Get-RowAt $grow 4).detailAvailable) 'An added page offered a source-level comparison.'
Assert-ChangeTest (-not [bool](Get-RowAt $same 1).detailAvailable) 'A page whose source did not change offered a source-level comparison.'

# ---- 比べる元の版の記録が消えている ---------------------------------------------
$gone = @{}
$gone[(Get-PackChangeVersionKey 'wb_a' 's1' 'v1')] = 'missing'
$gone[(Get-PackChangeVersionKey 'wb_a' 's2' 'v2')] = New-TestHashes @{ 'シートA' = @('X') }
$lost = Compare-PackArchivePageMaps (New-TestManifest @((New-TestItem 'A'))) (New-TestMetadata @{ wb_a = 's1|v1' }) (New-TestManifest @((New-TestItem 'A'))) (New-TestMetadata @{ wb_a = 's2|v2' }) $gone
Assert-ChangeTest ((Get-Kinds (Get-RowAt $lost 1)) -eq 'unavailable' -and [int]$lost.counts.unavailable -eq 1 -and -not [bool](Get-RowAt $lost 1).detailAvailable) 'A page whose previous version is gone was not reported as not comparable.'

# ---- 名前だけ変わった ----------------------------------------------------------
$renamed = Compare-PackArchivePageMaps (New-TestManifest @((New-TestItem 'A' 1 'wb_a' '売上' 'worksheet:S1'))) (New-TestMetadata $v1) (New-TestManifest @((New-TestItem 'A' 1 'wb_a' '売上（改）' 'worksheet:S1'))) (New-TestMetadata $v1) @{}
Assert-ChangeTest ((Get-Kinds (Get-RowAt $renamed 1)) -eq 'renamed' -and [string](Get-RowAt $renamed 1).fromTitle -eq '売上') 'A renamed page was not reported as renamed only.'

# ---- physicalPages の無い古いアーカイブ ------------------------------------------
# 左右を推定して埋めるようにすると、fromSide が入る・side が数えられるので落ちる（見開きの契約）。
$legacy = Compare-PackArchivePageMaps (New-TestManifest $base $false) (New-TestMetadata $v1) (New-TestManifest $inserted) (New-TestMetadata $v1) @{}
Assert-ChangeTest (-not [bool]$legacy.hasSides -and -not [bool]$legacy.fromHasPageMap -and [bool]$legacy.toHasPageMap) 'An archive without physicalPages was treated as having sides.'
Assert-ChangeTest ([int]$legacy.counts.side -eq 0) 'Side swaps were reported against an archive that never recorded sides.'
Assert-ChangeTest (@($legacy.pages | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.fromSide) -or $null -ne $_.fromSpreadIndex }).Count -eq 0) 'Sides or spreads were filled in for an archive that never recorded them.'
Assert-ChangeTest ([int]$legacy.counts.added -eq 1 -and [int]$legacy.counts.shifted -eq 3 -and [int](Get-RowAt $legacy 3).fromPage -eq 2) 'Page numbers were not counted from pages[] for an old archive.'

# ---- 既定の組み合わせと /api/v2/outputs/changes の本体 ------------------------------
$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalConfigRoot 'data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
Save-AppConfig ([pscustomobject][ordered]@{
    schemaVersion = 2; lastSubmissionDir = $submissionDir; lastDataDir = $dataDir; lastOutputDir = $outputDir
})
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')

# 原稿のPDFを置かずに済むよう、入力スナップショットだけを差し替える（指紋は出力先ごとに固定）。
function Get-FinalBuildInputSnapshot($Structure, [string]$Language, [string]$Volume, [string]$PackId) {
    return [pscustomobject]@{ fingerprint="fp|$PackId|$Volume"; blockers=@(); pageCount=1 }
}
# 原稿の版の保護（pin）はこの検査の対象ではない。架空の版に pin を作らない。
function Set-FinalPdfSnapshotPins { }

$pack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='変更点の一式' })
$packId = [string]$pack.packId
[void](Update-StructureLocked 'ja' {
    param($st)
    $wb = [pscustomobject][ordered]@{ workbookId='wb_a'; packId=$packId; fileName='wb_a.xlsx'; relativePath='wb_a.xlsx'; displayName='wb_a'; status='rendered' }
    Set-NoteProperty $st 'workbooks' (@(Get-Array $st.workbooks) + @($wb))
    $pages = @(Get-Array $st.pages)
    $pages += [pscustomobject][ordered]@{ pageId='wb_a_10'; workbookId='wb_a'; sheetName='シートA'; title='シートA'; volume='ja-main'; enabled=$true; order=10; numberingMode='visible' }
    $pages += [pscustomobject][ordered]@{ pageId='wb_a_20'; workbookId='wb_a'; sheetName='シートZ'; title='シートZ'; volume='ja-appendix'; enabled=$true; order=20; numberingMode='visible' }
    Set-NoteProperty $st 'pages' @($pages)
    return $null
})
function Get-PlannedOutputPath([string]$Volume) {
    $st = Get-Structure 'ja'
    $scope = Resolve-DocumentPackScope $st $packId $true
    return [IO.Path]::GetFullPath((Join-Path $outputDir ([string](Get-PackTargetOutputNaming $st 'ja' $scope $Volume).outputFileName)))
}
# 出力ジョブ（Build-DocumentPackPdf）が確定したときと同じく、出力状態とアーカイブを置く。
function Set-ArchivedOutput([string]$Volume, [string]$Content, $Manifest, $Metadata) {
    $path = Get-PlannedOutputPath $Volume
    [IO.File]::WriteAllText($path, $Content)
    $buildId = New-RbId
    [void](Update-StructureLocked 'ja' {
        param($st)
        $state = Get-PackOutputState $st 'ja' $packId $Volume $true
        Set-NoteProperty $state 'outputPdf' $path
        Set-NoteProperty $state 'builtFingerprint' ([string](Get-FinalBuildInputSnapshot $st 'ja' $Volume $packId).fingerprint)
        Set-NoteProperty $state 'outputIntegrity' (Get-OutputPdfIntegrity $path)
        Set-NoteProperty $state 'buildId' $buildId
        return $null
    })
    [void](New-FinalArchive 'ja' $packId $Volume $buildId $path $Manifest ([ordered]@{ sourceWorkbooks=@($Metadata.sourceWorkbooks); pages=@(1) }))
    # builtAt は秒までなので、同じ秒の出力は buildId（ミリ秒）で並ぶ。名前が重ならないよう少し待つ。
    Start-Sleep -Milliseconds 20
    return $buildId
}
function Write-TestVisualHashes([string]$Snapshot, [string]$Version, [hashtable]$Sheets) {
    Write-JsonFile (Join-Path (Get-RenderRecordDir 'ja' 'wb_a' $Snapshot $Version) 'visual-hashes.json') (New-TestHashes $Sheets)
}

$oneItem = @((New-TestItem 'A' 2))
$noOutput = Get-PackChangeSummary 'ja' $packId 'main'
Assert-ChangeTest ([string]$noOutput.status -eq 'no-output' -and $null -eq $noOutput.from -and $null -eq $noOutput.to) "A target without any output was not reported as 'no-output': $($noOutput.status)"

$build1 = Set-ArchivedOutput 'ja-main' '%PDF-change-1' (New-TestManifest $oneItem) (New-TestMetadata @{ wb_a = 's1|v1' })
$first = Get-PackChangeSummary 'ja' $packId 'main'
Assert-ChangeTest ([string]$first.status -eq 'no-baseline' -and [string]$first.to.buildId -eq $build1 -and $null -eq $first.from) "The first output was not reported as having nothing to compare: $($first.status)"
Assert-ChangeTest ([bool]$first.to.isCurrent -and [bool]$first.to.isLatest -and [int]$first.to.pageCount -eq 2) 'The current output was not described as current and latest with its physical page count.'

# 2回目（まだコピーしていない）→ 1つ前の出力と比べる。原稿の版が変わり、見た目のハッシュはまだ無い。
# 版の記録（スナップショット）はある。無ければ「比べられません」になる（下で確かめる）。
foreach ($snapshotId in @('s1','s2','s3')) { New-Item -ItemType Directory -Path (Get-SnapshotDir 'ja' 'wb_a' $snapshotId) -Force | Out-Null }
$build2 = Set-ArchivedOutput 'ja-main' '%PDF-change-2' (New-TestManifest $oneItem) (New-TestMetadata @{ wb_a = 's2|v2' })
$unpublished = Get-PackChangeSummary 'ja' $packId 'main'
Assert-ChangeTest ([string]$unpublished.status -eq 'ready' -and [string]$unpublished.from.buildId -eq $build1 -and [string]$unpublished.to.buildId -eq $build2 -and [string]$unpublished.fromReason -eq 'previous') 'Without a copy the default pair did not compare with the previous output.'
Assert-ChangeTest (@($unpublished.notes | Where-Object { [string]$_.code -eq 'not-published' }).Count -eq 1) 'The "not copied yet" note is missing.'
Assert-ChangeTest ([int]$unpublished.counts.maybeContent -eq 2 -and -not [bool]$unpublished.complete) 'Pages of a new source version without hashes were not "maybe changed".'
# ハッシュの揃っていない結果は覚えた値を使い回さない。解析が済めば結果が変わる。
Write-TestVisualHashes 's1' 'v1' @{ 'シートA' = @('P1','P2') }
Write-TestVisualHashes 's2' 'v2' @{ 'シートA' = @('P1','Q2') }
$analysed = Get-PackChangeSummary 'ja' $packId 'main'
Assert-ChangeTest ([int]$analysed.counts.sameLook -eq 1 -and [int]$analysed.counts.content -eq 1 -and [bool]$analysed.complete) "An incomplete cached comparison was reused after the hashes appeared: $($analysed.counts | ConvertTo-Json -Compress)"
$cacheDir = Get-PackChangeReportDir 'ja' $packId 'main' $build1 $build2
$cachePath = Join-Path $cacheDir 'summary.json'
Assert-ChangeTest ((Test-Path -LiteralPath $cachePath -PathType Leaf) -and $cacheDir.StartsWith((Join-Path (Get-WorkspacePath 'ja') 'exports\change-reports'))) 'The comparison was not kept under exports\change-reports.'
# 揃った結果は覚えた値を使う（アーカイブは書き換わらない）。
$cachedJson = Read-JsonFile $cachePath $null
$cachedJson.comparison.counts.content = 99
Write-JsonFile $cachePath $cachedJson
Assert-ChangeTest ([int](Get-PackChangeSummary 'ja' $packId 'main').counts.content -eq 99) 'A complete cached comparison was not reused.'
Remove-Item -LiteralPath $cachePath -Force

# コピーすると、その版が「前回コピーした版」。今の版と同じなら1つ前と比べ、そう伝える。
$publish = Publish-DocumentPackPdfToShared 'ja' @('ja-main') $packId
Assert-ChangeTest ([string]@($publish.files)[0].buildId -eq $build2) 'The publish did not copy the second output.'
$samePublished = Get-PackChangeSummary 'ja' $packId 'main'
Assert-ChangeTest ([string]$samePublished.from.buildId -eq $build1 -and @($samePublished.notes | Where-Object { [string]$_.code -eq 'same-as-published' }).Count -eq 1) 'Comparing right after a copy did not fall back to the previous output with a note.'
Assert-ChangeTest ([bool]$samePublished.to.isLastPublished -and [string]$samePublished.to.folderName -eq [string]$publish.folderName) 'The copied version was not described as the last copy.'

# 3回目・4回目（コピーせず）→ 比べる元は前回コピーした2回目のまま（1つ前の3回目ではない）。
# Get-LastPublishedBuildId の結果を無視すると3回目と比べて落ちる。
$build3 = Set-ArchivedOutput 'ja-main' '%PDF-change-3' (New-TestManifest @((New-TestItem 'A' 2), (New-TestItem 'B'))) (New-TestMetadata @{ wb_a = 's2|v2' })
$build4 = Set-ArchivedOutput 'ja-main' '%PDF-change-4' (New-TestManifest @((New-TestItem 'B'), (New-TestItem 'A' 2))) (New-TestMetadata @{ wb_a = 's2|v2' })
$sincePublish = Get-PackChangeSummary 'ja' $packId 'main'
Assert-ChangeTest ([string]$sincePublish.from.buildId -eq $build2 -and [string]$sincePublish.to.buildId -eq $build4 -and [string]$sincePublish.fromReason -eq 'published' -and [bool]$sincePublish.defaultPair) "The default pair did not compare with the last copied version: from=$($sincePublish.from.buildId)"
Assert-ChangeTest ([int]$sincePublish.counts.added -eq 1 -and [int]$sincePublish.counts.side -eq 2 -and [int]$sincePublish.counts.content -eq 0) "The changes since the last copy are wrong: $($sincePublish.counts | ConvertTo-Json -Compress)"
Assert-ChangeTest ([string]$sincePublish.from.publishedAt -eq [string]$publish.publishedAt -and -not [bool]$sincePublish.from.isCurrent -and [bool]$sincePublish.to.isCurrent) 'The from/to bands do not say which was copied and which is current.'
Assert-ChangeTest ([string]$sincePublish.report.reason -eq 'not-available' -and -not [bool]$sincePublish.report.publishable) 'The change report was offered before it can be made.'

# 好きな2回: 新しい→古いの順に選んでも、古い→新しいに入れ替えて比べる。
$picked = Get-PackChangeSummary 'ja' $packId 'main' $build4 $build1
Assert-ChangeTest ([string]$picked.from.buildId -eq $build1 -and [string]$picked.to.buildId -eq $build4 -and [bool]$picked.swapped -and -not [bool]$picked.defaultPair) 'A reversed pick was not swapped to old → new.'
Assert-ChangeTest (@($picked.notes | Where-Object { [string]$_.code -eq 'swapped' }).Count -eq 1) 'The swap was not explained.'
# 比べる元だけを指定したら、比べる先は今の提出用PDF（保存済み出力の「今の版と比べる」）。
$fromOnly = Get-PackChangeSummary 'ja' $packId 'main' $build3
Assert-ChangeTest ([string]$fromOnly.from.buildId -eq $build3 -and [string]$fromOnly.to.buildId -eq $build4 -and [string]$fromOnly.fromReason -eq 'selected') 'Picking only the older version did not compare it with the current output.'
Assert-Throws { Get-PackChangeSummary 'ja' $packId 'main' $build2 $build2 } '同じ版' 'Comparing a version with itself was accepted.'
# 別の出力先の版・無い版・パスを抜ける識別子は 400（ArgumentException）。
$appendixBuild = Set-ArchivedOutput 'ja-appendix' '%PDF-appendix-1' (New-TestManifest @((New-TestItem 'Z'))) (New-TestMetadata @{ wb_a = 's2|v2' })
Assert-Throws { Get-PackChangeSummary 'ja' $packId 'main' $appendixBuild $build4 } '保存済み出力にありません' 'A version of another target was accepted.'
Assert-Throws { Get-PackChangeSummary 'ja' $packId 'main' 'missing-build' } '保存済み出力にありません' 'A missing version was accepted.'
Assert-Throws { Get-PackChangeSummary 'ja' $packId 'main' '..\other' } '不正' 'A buildId with a path separator was accepted.'
Assert-Throws { Get-PackChangeSummary 'ja' $packId 'nowhere' } '存在しない出力先' 'An unknown target was accepted.'
$caughtType = ''
try { [void](Get-PackChangeSummary 'ja' $packId 'main' $appendixBuild) } catch { $caughtType = $_.Exception.GetType().FullName }
Assert-ChangeTest ($caughtType -eq 'System.ArgumentException') "A bad version pick must be a 400 (ArgumentException), got $caughtType."

# 比べる元の版の記録（スナップショット）が消えていたら、その行だけ「比べられません」。
$build5 = Set-ArchivedOutput 'ja-main' '%PDF-change-5' (New-TestManifest @((New-TestItem 'B'), (New-TestItem 'A' 2))) (New-TestMetadata @{ wb_a = 's3|v3' })
Write-TestVisualHashes 's3' 'v3' @{ 'シートA' = @('R1','R2'); 'シートB' = @('R3') }
Remove-Item -LiteralPath (Get-SnapshotDir 'ja' 'wb_a' 's2') -Recurse -Force
$partial = Get-PackChangeSummary 'ja' $packId 'main' $build4 $build5
Assert-ChangeTest ([string]$partial.status -eq 'ready' -and [int]$partial.counts.unavailable -eq 3 -and @($partial.notes | Where-Object { [string]$_.code -eq 'some-unavailable' }).Count -eq 1) "Rows whose previous version is gone were not reported as not comparable: $($partial.counts | ConvertTo-Json -Compress)"

# アーカイブのPDFがすり替えられていたら、比べる先にしない。
$archivedPdf = Join-Path (Join-Path (Join-Path (Join-Path (Get-ArchiveRoot 'ja') $packId) 'ja-main') $build5) 'final.pdf'
[IO.File]::WriteAllText($archivedPdf, '%PDF-swapped')
$tampered = Get-PackChangeSummary 'ja' $packId 'main' $build4 $build5
Assert-ChangeTest ([string]$tampered.status -eq 'unavailable' -and $null -eq $tampered.from) 'A swapped archive PDF was still compared.'

# 古い形式（physicalPages の無い）版と比べると、左右は比べないと伝える。
$build6 = Set-ArchivedOutput 'ja-main' '%PDF-change-6' (New-TestManifest @((New-TestItem 'A' 2)) $false) (New-TestMetadata @{ wb_a = 's3|v3' })
$old = Get-PackChangeSummary 'ja' $packId 'main' $build6 $build4
Assert-ChangeTest ([string]$old.from.buildId -eq $build4 -and [string]$old.to.buildId -eq $build6 -and [bool]$old.swapped) 'The legacy pick was not ordered old → new.'
Assert-ChangeTest (-not [bool]$old.hasSides -and @($old.notes | Where-Object { [string]$_.code -eq 'no-page-map' }).Count -eq 1 -and [int]$old.counts.side -eq 0) 'Comparing with an archive without sides did not say so.'

# 組み込み一式は出力の版を保存しないので、比べられないと返す（落ちない）。
$builtin = Get-PackChangeSummary 'ja' 'ecm' 'main'
Assert-ChangeTest ([string]$builtin.status -eq 'unsupported') "A built-in pack was not reported as unsupported: $($builtin.status)"

# 応答は JSON にしても形が崩れない（kinds が1件でも配列のまま）。
$json = ConvertTo-Json -InputObject ([ordered]@{ ok=$true; changes=$sincePublish }) -Depth 50 | ConvertFrom-Json
Assert-ChangeTest (@($json.changes.pages).Count -eq 3 -and @($json.changes.pages)[0].kinds -is [array]) 'The changes response lost its arrays when serialized.'

Write-Output 'pack-change selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
