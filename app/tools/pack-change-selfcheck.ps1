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
Assert-ChangeTest ([string]$sincePublish.report.reason -eq 'not-created' -and -not [bool]$sincePublish.report.exists -and -not [bool]$sincePublish.report.publishable) 'A change report was reported as made (or copyable) before it was made.'
# 「前回コピーした版 → 今の提出用PDF」の組み合わせは、作れば共有コピーへ同梱できる。
Assert-ChangeTest ([bool]$sincePublish.report.attachable) 'The last-copy → current pair was not offered for the shared copy.'

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

# ==== 変更点PDF（表紙・変更一覧。Edge のヘッドレス印刷） ===================================
function New-TestReportSummary([object[]]$Pages, [bool]$HasSides = $true, [hashtable]$Counts = @{}) {
    $c = New-PackChangeCounts
    foreach ($k in @($Counts.Keys)) { $c[$k] = $Counts[$k] }
    return [ordered]@{
        status = 'ready'; hasSides = $HasSides; complete = $true; counts = $c; notes = @(); pages = @($Pages)
        from = [ordered]@{ buildId = 'b1'; builtAt = '2026-10-28T14:05:00+09:00'; pageCount = 32; outputFileName = 'old.pdf'; publishedAt = '2026-10-28T15:00:00+09:00'; folderName = '20261028_1500_user'; isLastPublished = $true; isCurrent = $false }
        to = [ordered]@{ buildId = 'b2'; builtAt = '2026-11-02T09:40:00+09:00'; pageCount = 34; outputFileName = '一式_本体.pdf'; isCurrent = $true; isLatest = $true }
    }
}
function New-TestReportRow([int]$ToPage, [int]$FromPage, [string[]]$Kinds, [bool]$ShiftOnly = $false, [string]$ToSide = 'right', [string]$FromSide = 'right', [string]$SourceName = '売上.xlsx', [string]$SheetName = '売上') {
    return [ordered]@{ toPage = $(if ($ToPage -gt 0) { $ToPage } else { $null }); toSide = $ToSide; fromPage = $(if ($FromPage -gt 0) { $FromPage } else { $null }); fromSide = $FromSide; sourceName = $SourceName; sheetName = $SheetName; title = $SheetName; fromTitle = $SheetName; kinds = @($Kinds); shiftOnly = $ShiftOnly; changedRatio = $null; regionCount = $null; message = '' }
}
function Get-TextCount([string]$Text, [string]$Needle) {
    return ([regex]::Matches($Text, [regex]::Escape($Needle))).Count
}

# 利用者が付けた原稿名・シート名はエスケープする（<script> を含む名前で試す）。
# ずれただけのページの連続は1行にまとめ、左右が入れ替わったページは別の行に残す。変わっていないページは載せない。
$reportRows = New-Object System.Collections.ArrayList
[void]$reportRows.Add((New-TestReportRow 12 10 @('content') $false 'right' 'right' '<script>alert(1)</script>' 'S&1'))
for ($p = 15; $p -le 34; $p++) { [void]$reportRows.Add((New-TestReportRow $p ($p - 2) @('moved') $true ($(if ($p % 2 -eq 1) { 'right' } else { 'left' })) ($(if ($p % 2 -eq 1) { 'right' } else { 'left' })))) }
[void]$reportRows.Add((New-TestReportRow 35 34 @('moved','side') $true 'right' 'left'))
[void]$reportRows.Add((New-TestReportRow 36 36 @('unchanged')))
[void]$reportRows.Add((New-TestReportRow 37 37 @('same-look')))
# 左右の記録が無い古い版の行。頁番号だけを書き、左右を補わない（見開きの契約）。
[void]$reportRows.Add((New-TestReportRow 40 39 @('moved') $false '' ''))
$entries = @(Get-PackChangeReportEntries @($reportRows))
Assert-ChangeTest ($entries.Count -eq 4) "The report list did not fold shifted pages into one line or did not drop unchanged pages: $($entries.Count) entries"
Assert-ChangeTest ([string]$entries[1].type -eq 'shift' -and [int]$entries[1].startPage -eq 15 -and [int]$entries[1].endPage -eq 34 -and [int]$entries[1].count -eq 20) 'The shifted run is wrong.'
$html = New-PackChangeReportHtml (New-TestReportSummary @($reportRows) $true @{ changed = 23; content = 1; moved = 22; shifted = 21; side = 1 }) '<b>一式</b>' '本体' '2026-11-02T10:00:00+09:00'
Assert-ChangeTest (-not $html.Contains('<script>alert') -and $html.Contains('&lt;script&gt;alert(1)&lt;/script&gt;') -and $html.Contains('S&amp;1')) 'Source and sheet names were not escaped in the report HTML.'
Assert-ChangeTest (-not $html.Contains('<b>一式</b>') -and $html.Contains('&lt;b&gt;一式&lt;/b&gt;')) 'The pack name was not escaped in the report HTML.'
Assert-ChangeTest ((Get-TextCount $html 'p.15〜p.34 は前のページが増えたため2ページ後ろにずれました') -eq 1) 'The shifted pages were not summarised in one line.'
Assert-ChangeTest ((Get-TextCount $html '<tr class="r') -eq 4 -and -not $html.Contains('>p.16（')) 'Folded or unchanged pages were listed one by one.'
Assert-ChangeTest ($html.Contains('p.35（右）') -and $html.Contains('p.34（左）') -and $html.Contains('左右が入れ替わった')) 'A page whose side changed after a shift was hidden.'
Assert-ChangeTest ($html.Contains('<td>p.40</td>') -and $html.Contains('<td>p.39</td>') -and -not $html.Contains('p.40（')) 'A side was filled in for a page without a recorded side.'
# 全ページの上に「確認用」、下に「変更点 n / N」。総ページ数はサーバーで決める。
Assert-ChangeTest ((Get-TextCount $html '<section class="sheet') -eq 2 -and (Get-TextCount $html 'class="mark-badge">確認用') -eq 2) 'Every report page must carry the 確認用 mark.'
Assert-ChangeTest ($html.Contains('変更点 1 / 2') -and $html.Contains('変更点 2 / 2') -and (Get-TextCount $html '提出物ではありません') -ge 3) 'The page footer or the "not a submission" notice is missing.'
Assert-ChangeTest ($html.Contains('前回（共有用フォルダーにコピーした版）') -and $html.Contains('今回（今の提出用PDF）') -and $html.Contains('20261028_1500_user') -and $html.Contains('左右が入れ替わったページ')) 'The cover does not describe both versions and the side count.'
Assert-ChangeTest ($html.Contains('@page{size:A4 landscape') -and $html.Contains('<meta charset="utf-8">')) 'The report HTML is not an A4 landscape UTF-8 page.'
# 左右の記録が無い版との比較では、左右の数を 0 と書かない。
$noSides = New-PackChangeReportHtml (New-TestReportSummary @($reportRows) $false) '一式' '本体'
Assert-ChangeTest (-not $noSides.Contains('左右が入れ替わったページ') -and $noSides.Contains('左右までは比べていません')) 'The side count was printed for archives without recorded sides.'
# 1ページに載せる行数を超えたらページを分け、総ページ数も合わせる。
$many = @(for ($p = 1; $p -le 30; $p++) { New-TestReportRow $p $p @('content') })
$manyHtml = New-PackChangeReportHtml (New-TestReportSummary $many) '一式' '本体'
$expectedSheets = 1 + [int][Math]::Ceiling(30 / [double]$Script:PackChangeReportRowsPerPage)
Assert-ChangeTest ((Get-TextCount $manyHtml '<section class="sheet') -eq $expectedSheets -and $manyHtml.Contains(('変更点 {0} / {0}' -f $expectedSheets)) -and (Get-TextCount $manyHtml '<tr class="r') -eq 30) 'Long lists were not split across pages with a matching total.'
# 変更が1つも無いときは、「変更なし」と書いた1枚だけ。
$noneHtml = New-PackChangeReportHtml (New-TestReportSummary @((New-TestReportRow 1 1 @('unchanged'))) $true @{ unchanged = 1 }) '一式' '本体'
Assert-ChangeTest ((Get-TextCount $noneHtml '<section class="sheet') -eq 1 -and $noneHtml.Contains('変更点 1 / 1') -and $noneHtml.Contains('前回コピーしたときから、変わったページはありません。')) 'A report without changes must be a single "no changes" page.'

# ---- Edge のヘッドレス印刷 ---------------------------------------------------------
# 専用の --user-data-dir を必ず渡す（起動中の Edge にのっとられると PDF ができない）。
$edgeArgs = @(Get-EdgePrintToPdfArguments 'C:\t\profile' 'C:\t\report.pdf' 'file:///C:/t/report.html')
Assert-ChangeTest ($edgeArgs -contains '--headless' -and $edgeArgs -contains '--user-data-dir=C:\t\profile' -and $edgeArgs -contains '--no-pdf-header-footer' -and $edgeArgs -contains '--print-to-pdf=C:\t\report.pdf' -and $edgeArgs[-1] -eq 'file:///C:/t/report.html') "The Edge print arguments are wrong: $($edgeArgs -join ' ')"
Assert-Throws { Get-EdgePrintToPdfArguments '' 'C:\t\report.pdf' 'file:///x' } 'プロファイル' 'Edge was allowed to print with the default profile.'
# 待ち時間の上限は件数に比例させ、最大10分。
Assert-ChangeTest ((Get-PackChangeReportTimeoutSeconds 0) -eq 60 -and (Get-PackChangeReportTimeoutSeconds 100) -eq 110 -and (Get-PackChangeReportTimeoutSeconds 2000) -eq 600) 'The report timeout is not proportional to the item count.'

# Edge の代わりに、印刷を真似る偽物を使う（PDF を書く／時間切れ／PDF を書かずに正常終了）。
$Script:EdgeCalls = New-Object System.Collections.ArrayList
$Script:EdgeMode = 'ok'
function Get-EdgeExecutablePath { return 'C:\fake\Edge\msedge.exe' }
function Invoke-NativeCapture([string]$FilePath, [string[]]$ArgumentList, [int]$TimeoutSeconds = 120) {
    $profileArg = @($ArgumentList | Where-Object { $_ -like '--user-data-dir=*' } | Select-Object -First 1)
    $pdfArg = @($ArgumentList | Where-Object { $_ -like '--print-to-pdf=*' } | Select-Object -First 1)
    $edgeProfileDir = $(if ($profileArg.Count -gt 0) { $profileArg[0].Substring('--user-data-dir='.Length) } else { '' })
    $edgeProfileExisted = (-not [string]::IsNullOrWhiteSpace($edgeProfileDir) -and (Test-Path -LiteralPath $edgeProfileDir))
    [void]$Script:EdgeCalls.Add([pscustomobject]@{ file = $FilePath; args = @($ArgumentList); timeout = $TimeoutSeconds; profile = $edgeProfileDir; profileExisted = $edgeProfileExisted })
    if ($Script:EdgeMode -eq 'timeout') { return [ordered]@{ exitCode = -1; output = @(); text = 'NATIVE_TIMEOUT: fake'; timedOut = $true; cancelled = $false } }
    if ($Script:EdgeMode -eq 'blocked') { return [ordered]@{ exitCode = 0; output = @('HeadlessModeEnabled=false'); text = 'HeadlessModeEnabled=false'; timedOut = $false; cancelled = $false } }
    [IO.File]::WriteAllText($pdfArg[0].Substring('--print-to-pdf='.Length), '%PDF-1.7 fake change report')
    return [ordered]@{ exitCode = 0; output = @(); text = ''; timedOut = $false; cancelled = $false }
}
$printDir = Join-Path $Script:LocalConfigRoot 'print-test'
New-Item -ItemType Directory -Path $printDir -Force | Out-Null
$printHtml = Join-Path $printDir 'report.html'
[IO.File]::WriteAllText($printHtml, '<!DOCTYPE html><p>x</p>')
$printed = Invoke-EdgePrintToPdf $printHtml (Join-Path $printDir 'report.pdf') 90
Assert-ChangeTest ([bool]$printed.ok -and (Test-PdfMagic (Join-Path $printDir 'report.pdf'))) 'A successful Edge print was not accepted.'
$firstCall = $Script:EdgeCalls[$Script:EdgeCalls.Count - 1]
Assert-ChangeTest ([bool]$firstCall.profileExisted -and ([IO.Path]::GetFileName($firstCall.profile)) -like 'rb-chg-edge-*' -and -not (Test-Path -LiteralPath $firstCall.profile) -and [int]$firstCall.timeout -eq 90) 'Edge was not given its own temporary profile (or it was not cleaned up).'
[void](Invoke-EdgePrintToPdf $printHtml (Join-Path $printDir 'report.pdf') 90)
Assert-ChangeTest ([string]$Script:EdgeCalls[$Script:EdgeCalls.Count - 1].profile -ne [string]$firstCall.profile) 'Two prints shared the same Edge profile.'
# timedOut を読まずに「PDF が無い」と言うと、何をすればよいか分からない（DEVELOPMENT_RULES §1）。
$Script:EdgeMode = 'timeout'
$timedOut = Invoke-EdgePrintToPdf $printHtml (Join-Path $printDir 'report.pdf') 90
Assert-ChangeTest (-not [bool]$timedOut.ok -and [string]$timedOut.code -eq 'timeout' -and [bool]$timedOut.timedOut -and [string]$timedOut.message -like '*時間内*') "A timed-out Edge print was not reported as a timeout: $($timedOut.code)"
Assert-ChangeTest (-not (Test-Path -LiteralPath (Join-Path $printDir 'report.pdf'))) 'An old PDF survived a failed print and could be mistaken for the new one.'
# 管理ポリシーで止められていると、Edge は正常終了しても PDF を書かない。
$Script:EdgeMode = 'blocked'
$blocked = Invoke-EdgePrintToPdf $printHtml (Join-Path $printDir 'report.pdf') 90
Assert-ChangeTest (-not [bool]$blocked.ok -and [string]$blocked.code -eq 'no-pdf' -and [string]$blocked.message -like '*ヘッドレス印刷*' -and [string]$blocked.log -like '*HeadlessModeEnabled*') 'A print that wrote no PDF was accepted or not explained.'

# ---- 作成ジョブの本体・作成済みの判定・開く ------------------------------------------
$Script:EdgeMode = 'ok'
$reportJobId = 'job_20261102_000000_aaaaaaaa'
$reportStatusPath = Join-Path (Get-RenderJobDir 'ja') "$reportJobId.status.json"
Invoke-PackChangeReportJobCore ([pscustomobject]@{ jobId = $reportJobId; mode = 'ja'; packId = $packId; targetId = 'main'; fromBuildId = $build3; toBuildId = $build4; statusPath = $reportStatusPath })
$doneJob = Read-JsonFile $reportStatusPath $null
Assert-ChangeTest ([string]$doneJob.status -eq 'completed' -and [string]$doneJob.jobType -eq 'change-report' -and [bool]$doneJob.report.exists) "The report job did not complete: $($doneJob.status) $($doneJob.message)"
$reportDir = Get-PackChangeReportDir 'ja' $packId 'main' $build3 $build4
foreach ($name in @('report.pdf','report.html','meta.json','summary.json')) { Assert-ChangeTest (Test-Path -LiteralPath (Join-Path $reportDir $name) -PathType Leaf) "The report folder is missing $name." }
Assert-ChangeTest ($reportDir.StartsWith((Join-Path (Get-WorkspacePath 'ja') 'exports\change-reports'))) 'The report was not kept under exports\change-reports.'
# 一式の名前（変更点の一式）は提出用PDFの名前にも入るので、変更点PDFの名前（…_変更点.pdf）で探す。
Assert-ChangeTest (@(Get-ChildItem -LiteralPath $outputDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*_変更点.pdf' }).Count -eq 0) 'The change report was copied into the local output folder.'
$reportSummary = Get-PackChangeSummary 'ja' $packId 'main' $build3 $build4
Assert-ChangeTest ([int]$Script:EdgeCalls[$Script:EdgeCalls.Count - 1].timeout -eq (Get-PackChangeReportTimeoutSeconds @($reportSummary.pages).Count)) 'The report job did not use the proportional timeout.'
# 好きな2回（比べる元が前回コピーした版でない）で作ったものは、作成済みでも共有コピーへは付けない。
Assert-ChangeTest ([bool]$reportSummary.report.exists -and [string]$reportSummary.report.reason -eq 'not-attachable' -and -not [bool]$reportSummary.report.attachable -and -not [bool]$reportSummary.report.publishable -and [string]$reportSummary.report.reportId -eq ($build3 + '__' + $build4)) "The changes response does not describe the made report: $($reportSummary.report | ConvertTo-Json -Compress)"
Assert-ChangeTest ([string]$reportSummary.report.fileName -eq ([IO.Path]::GetFileNameWithoutExtension([string]$reportSummary.to.outputFileName) + '_変更点.pdf')) 'The report file name does not follow the submission PDF.'
# 作成済みなら作り直さない（子プロセスを起動しない）。2つの版の指定は必須。
$again = Start-PackChangeReportJob 'ja' $packId 'main' $build3 $build4
Assert-ChangeTest ([string]$again.status -eq 'completed' -and [string]$again.jobId -eq '' -and [bool]$again.reportReady) 'A made report was made again.'
Assert-Throws { Start-PackChangeReportJob 'ja' $packId 'main' '' $build4 } 'fromBuildId' 'A report was started without the versions shown on screen.'
Assert-Throws { Start-PackChangeReportJob 'ja' 'ecm' 'main' 'b1' 'b2' } '組み込み' 'A report was started for a built-in pack.'
Assert-ChangeTest ((Resolve-PackChangeReportFile 'ja' $packId 'main' $build3 $build4) -eq (Join-Path $reportDir 'report.pdf')) 'The made report cannot be opened.'
Assert-Throws { Resolve-PackChangeReportFile 'ja' $packId 'main' '..\x' $build4 } '不正' 'A report path outside exports\change-reports was accepted.'
# 見た目のハッシュがそろう前に作ったものは、そろった後には古い（一覧の中身が変わる）。
$metaPath = Join-Path $reportDir 'meta.json'
$meta = Read-JsonFile $metaPath $null
$meta.complete = $false
Write-JsonFile $metaPath $meta
$stateArgs = @('ja', $packId, 'main', $build3, $build4, [string]$reportSummary.from.outputPdfSha256, [string]$reportSummary.to.outputPdfSha256)
Assert-ChangeTest (-not [bool](Get-PackChangeReportState @stateArgs $true).exists -and [string](Get-PackChangeReportState @stateArgs $true).reason -eq 'outdated') 'A report made before the hashes were complete was still offered.'
Assert-ChangeTest ([bool](Get-PackChangeReportState @stateArgs $false).exists) 'A report made from the same incomplete comparison was not offered.'
$meta.complete = $true
Write-JsonFile $metaPath $meta
# 別の提出用PDFについての記録なら作成済みにしない。
Assert-ChangeTest (-not [bool](Get-PackChangeReportState 'ja' $packId 'main' $build3 $build4 ([string]$reportSummary.from.outputPdfSha256) ('0' * 64) $true).exists) 'A report about another submission PDF was offered.'
# 中身が作ったときから変わっていたら、作成済みにしないし開かない。
[IO.File]::WriteAllText((Join-Path $reportDir 'report.pdf'), '%PDF-tampered')
$tamperedReport = Get-PackChangeSummary 'ja' $packId 'main' $build3 $build4
Assert-ChangeTest (-not [bool]$tamperedReport.report.exists -and [string]$tamperedReport.report.reason -eq 'outdated') 'A tampered change report was still offered.'
Assert-Throws { Resolve-PackChangeReportFile 'ja' $packId 'main' $build3 $build4 } '変わっています' 'A tampered change report was opened.'
# 作り直すと、置き換えた report.pdf と meta.json がそろう。
Invoke-PackChangeReportJobCore ([pscustomobject]@{ jobId = $reportJobId; mode = 'ja'; packId = $packId; targetId = 'main'; fromBuildId = $build3; toBuildId = $build4; statusPath = $reportStatusPath })
$remadeJob = Read-JsonFile $reportStatusPath $null
Assert-ChangeTest ([string]$remadeJob.status -eq 'completed') "Re-making a tampered report failed: $($remadeJob.status) $($remadeJob.message) $(@($remadeJob.errors | ForEach-Object { [string]$_.detail }) -join ' | ')"
Assert-ChangeTest ([bool](Get-PackChangeSummary 'ja' $packId 'main' $build3 $build4).report.exists) 'Re-making a tampered report did not replace it.'
Assert-ChangeTest (@(Get-ChildItem -LiteralPath $reportDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*.tmp' }).Count -eq 0) 'Re-making a report left a temporary file behind.'

# 共有コピーへの同梱: 既定の組み合わせ（前回コピーした2回目 → 今の6回目）で作ったものだけ publishable。
# 画面の「いっしょにコピーできる」は、コピーのときと同じ Resolve-PublishableChangeReport で決める。
Invoke-PackChangeReportJobCore ([pscustomobject]@{ jobId = $reportJobId; mode = 'ja'; packId = $packId; targetId = 'main'; fromBuildId = $build2; toBuildId = $build6; statusPath = $reportStatusPath })
$attachJob = Read-JsonFile $reportStatusPath $null
Assert-ChangeTest ([string]$attachJob.status -eq 'completed') "The default pair report job failed: $($attachJob.status) $($attachJob.message)"
$attachSummary = Get-PackChangeSummary 'ja' $packId 'main'
Assert-ChangeTest ([string]$attachSummary.from.buildId -eq $build2 -and [string]$attachSummary.to.buildId -eq $build6 -and [bool]$attachSummary.report.exists) "The default pair report was not made: $($attachSummary.report | ConvertTo-Json -Compress)"
Assert-ChangeTest ([bool]$attachSummary.report.attachable -and [bool]$attachSummary.report.publishable -and [string]$attachSummary.report.reason -eq 'publishable') "A made last-copy → current report was not offered for the shared copy: $($attachSummary.report | ConvertTo-Json -Compress)"
$attachCheck = Resolve-PublishableChangeReport 'ja' $packId 'main' $build6 ([string]$attachSummary.to.outputPdfSha256) ([string]$attachSummary.to.outputFileName)
Assert-ChangeTest ([bool]$attachCheck.ok -and [string]$attachCheck.fromBuildId -eq $build2 -and [string]$attachCheck.sourceFull -eq (Join-Path (Get-PackChangeReportDir 'ja' $packId 'main' $build2 $build6) 'report.pdf')) 'The publish check disagrees with the changes response.'
# 今の提出用PDFが出力の後に変わったら、同じ組み合わせでも同梱できると言わない。
$currentOutput = Get-PlannedOutputPath 'ja-main'
$currentOriginal = [IO.File]::ReadAllText($currentOutput)
$currentStamp = (Get-Item -LiteralPath $currentOutput).LastWriteTimeUtc
[IO.File]::AppendAllText($currentOutput, 'edited')
$editedSummary = Get-PackChangeSummary 'ja' $packId 'main' $build2 $build6
Assert-ChangeTest (-not [bool]$editedSummary.report.attachable -and -not [bool]$editedSummary.report.publishable) 'A report was offered for the shared copy after the submission PDF was edited.'
[IO.File]::WriteAllText($currentOutput, $currentOriginal)
(Get-Item -LiteralPath $currentOutput).LastWriteTimeUtc = $currentStamp

# Edge が時間切れ・PDF を書かないときは、ジョブを失敗にして理由を残す。作成済みの記録は書かない。
foreach ($case in @(@('timeout','時間内'), @('blocked','ヘッドレス印刷'))) {
    $Script:EdgeMode = $case[0]
    Invoke-PackChangeReportJobCore ([pscustomobject]@{ jobId = $reportJobId; mode = 'ja'; packId = $packId; targetId = 'main'; fromBuildId = $build1; toBuildId = $build4; statusPath = $reportStatusPath })
    $failedJob = Read-JsonFile $reportStatusPath $null
    Assert-ChangeTest ([string]$failedJob.status -eq 'failed' -and [string]$failedJob.message -like ('*' + $case[1] + '*') -and @($failedJob.errors).Count -eq 1) "A failed Edge print ($($case[0])) was not reported: $($failedJob.status) $($failedJob.message)"
    Assert-ChangeTest (-not (Test-Path -LiteralPath (Join-Path (Get-PackChangeReportDir 'ja' $packId 'main' $build1 $build4) 'meta.json'))) 'A failed print left a "made" record behind.'
}
Assert-ChangeTest ([string](Read-JsonFile $reportStatusPath $null).code -eq 'no-pdf') 'The failure code of the Edge print was not kept on the job.'
$Script:EdgeMode = 'ok'

Write-Output 'pack-change selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
