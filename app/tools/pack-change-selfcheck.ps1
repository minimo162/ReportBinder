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

# ---- 行の安定したキーと、原稿の変換PDFの中のページ番号 --------------------------------
# 変更点PDF用に「くわしく見る」の結果を行へ結び付ける。キーに出力ページ番号を使うと、前に1ページ
# 増えただけで後ろの行のキーが変わって落ちる。原稿内のページ番号に pagePartIndex を使うと、
# 原稿の途中のページから載せた項目（Start=3）で落ちる。
Assert-ChangeTest ([string](Get-RowAt $grow 2).rowKey -eq 'item:A#2' -and [int](Get-RowAt $grow 2).toSourcePageNumber -eq 2 -and [int](Get-RowAt $grow 2).fromSourcePageNumber -eq 2) 'The row key or the source page numbers of a changed page are wrong.'
Assert-ChangeTest ([string](Get-RowAt $grow 4).rowKey -eq 'item:A#4' -and [int](Get-RowAt $grow 4).toSourcePageNumber -eq 4 -and $null -eq (Get-RowAt $grow 4).fromSourcePageNumber) 'An added page must have no previous source page number.'
Assert-ChangeTest ([string](Get-RowAt $same 3).rowKey -eq 'item:B#2' -and [string](Get-RowAt $ins 4).rowKey -eq 'item:B#2') 'A page pushed back by an insertion must keep its row key.'
$removedRow = @($rem.pages | Where-Object { $null -eq $_.toPage })[0]
Assert-ChangeTest ([string]$removedRow.rowKey -eq 'item:B#1' -and $null -eq $removedRow.toSourcePageNumber -and [int]$removedRow.fromSourcePageNumber -eq 1) 'A removed page must keep its previous source page number only.'
Assert-ChangeTest (@($ins.pages | ForEach-Object { [string]$_.rowKey } | Select-Object -Unique).Count -eq @($ins.pages).Count) 'Row keys must be unique.'
$offset = Compare-PackArchivePageMaps (New-TestManifest @((New-TestItem 'A' 2 'wb_a' '' '' 3))) (New-TestMetadata @{ wb_a = 's1|v1' }) (New-TestManifest @((New-TestItem 'A' 2 'wb_a' '' '' 3))) (New-TestMetadata @{ wb_a = 's2|v2' }) @{}
Assert-ChangeTest ([int](Get-RowAt $offset 2).toSourcePageNumber -eq 4 -and [int](Get-RowAt $offset 2).fromSourcePageNumber -eq 4 -and [string](Get-RowAt $offset 2).rowKey -eq 'item:A#2') 'The source page number must be the page inside the source PDF, not the part index.'
$legacyOffset = Compare-PackArchivePageMaps (New-TestManifest @((New-TestItem 'A' 2 'wb_a' '' '' 3)) $false) (New-TestMetadata $v1) (New-TestManifest @((New-TestItem 'A' 2 'wb_a' '' '' 3)) $false) (New-TestMetadata $v1) @{}
Assert-ChangeTest ([int](Get-RowAt $legacyOffset 1).toSourcePageNumber -eq 3 -and [string](Get-RowAt $legacyOffset 1).rowKey -eq 'item:A#1') 'An old archive must count source page numbers from pages[].'
Assert-ChangeTest ([int]$grow.algorithmVersion -eq 2) 'The page comparison version was not raised, so old summary.json files without row keys would be reused.'

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

# ==== ページごとの詳細（PR-B2） =====================================================
function New-TestDetailRow([string]$RowKey, [int]$ToPage, [int]$FromPage, [string[]]$Kinds, [int]$ToSourcePage = 2, [int]$FromSourcePage = 2) {
    $row = New-TestReportRow $ToPage $FromPage $Kinds
    $row.rowKey = $RowKey
    $row.sourceId = 'wb_a'; $row.fromSourceId = $(if ($FromPage -gt 0) { 'wb_a' } else { '' })
    $row.fromSheetName = $(if ($FromPage -gt 0) { '売上' } else { '' })
    $row.fromSnapshotId = $(if ($FromPage -gt 0) { 's1' } else { '' }); $row.fromVersionId = $(if ($FromPage -gt 0) { 'v1' } else { '' })
    $row.toSnapshotId = $(if ($ToPage -gt 0) { 's2' } else { '' }); $row.toVersionId = $(if ($ToPage -gt 0) { 'v2' } else { '' })
    $row.toSourcePageNumber = $(if ($ToPage -gt 0) { $ToSourcePage } else { $null })
    $row.fromSourcePageNumber = $(if ($FromPage -gt 0) { $FromSourcePage } else { $null })
    $row.pagePartIndex = 2; $row.toPartCount = 3; $row.fromPartCount = 3
    $row.detailAvailable = ($ToPage -gt 0 -and $FromPage -gt 0)
    return $row
}

# 「アプリの『くわしく見る』で確認できます」は、アプリを持たない受け取り手には実行できないので書かない。
$plainContent = New-TestDetailRow 'item:A#2' 12 10 @('content')
Assert-ChangeTest ((Get-PackChangeReportVisualText $plainContent) -eq '変更箇所を特定できませんでした（ページ全体を確認してください）') 'A changed row without a detail page must ask to check the whole page.'
Assert-ChangeTest ((Get-PackChangeReportVisualText $plainContent $null 0 $true) -like '*最初の200ページ分*') 'A row beyond the detail limit must say the details were left out.'

# ---- details の検証（中身は信用しない） ----
$detailSummary = [ordered]@{ pages = @(
    (New-TestDetailRow 'item:A#2' 12 10 @('content')),
    (New-TestDetailRow 'item:B#1' 13 0 @('added') 1 0),
    (New-TestDetailRow 'item:C#1' 14 11 @('moved'))
) }
$longText = ('長' * 5000)
$manyRegions = @(for ($n = 0; $n -lt 40; $n++) { [ordered]@{ kind = 'modified'; before = [ordered]@{ x = 0.01; y = 0.01; width = 0.01; height = 0.01 }; after = [ordered]@{ x = 0.01; y = 0.01; width = 0.01; height = 0.01 }; beforeText = ''; afterText = '' } })
$rawDetails = @(
    [ordered]@{ rowKey = 'item:A#2'; fromVersionId = 'v1'; toVersionId = 'v2'; beforePageNumber = 2; afterPageNumber = 2; regionCount = 6; regions = @(
        [ordered]@{ kind = 'modified'; before = [ordered]@{ x = 0.1; y = 0.1; width = 0.2; height = 0.1 }; after = [ordered]@{ x = 0.1; y = 0.1; width = 0.2; height = 0.1 }; beforeText = '1,234'; afterText = '1,240' },
        [ordered]@{ kind = 'added'; before = [ordered]@{ x = -0.5; y = 0.5; width = 5; height = 0.1 }; after = [ordered]@{ x = 0.3; y = 0.3; width = 0.9; height = 0.1 }; beforeText = ''; afterText = ('<script>alert(2)</script>' + $longText) },
        [ordered]@{ kind = 'MODIFIED'; before = [ordered]@{ x = 0.1; y = 0.1; width = 0.1; height = 0.1 }; after = $null; beforeText = 'x'; afterText = 'y' },
        [ordered]@{ kind = 'evil'; before = [ordered]@{ x = 0.1; y = 0.1; width = 0.1; height = 0.1 }; after = $null; beforeText = 'x'; afterText = 'y' },
        [ordered]@{ kind = 'removed'; before = [ordered]@{ x = '0.1'; y = 0.1; width = 0.1; height = 0.1 }; after = [ordered]@{ x = 1.5; y = 0.1; width = 0.1; height = 0.1 }; beforeText = 'z'; afterText = '' }
    ) },
    # 同じ rowKey の2件目・実在しない rowKey・内容の変わっていない行・増えた行・版やページ番号の違うものは捨てる。
    [ordered]@{ rowKey = 'item:A#2'; fromVersionId = 'v1'; toVersionId = 'v2'; beforePageNumber = 2; afterPageNumber = 2; regions = @() },
    [ordered]@{ rowKey = 'item:Z#9'; fromVersionId = 'v1'; toVersionId = 'v2'; beforePageNumber = 2; afterPageNumber = 2; regions = @() },
    [ordered]@{ rowKey = 'item:C#1'; fromVersionId = 'v1'; toVersionId = 'v2'; beforePageNumber = 2; afterPageNumber = 2; regions = @() },
    [ordered]@{ rowKey = 'item:B#1'; fromVersionId = ''; toVersionId = 'v2'; beforePageNumber = 0; afterPageNumber = 1; regions = @() },
    'not an object', 42
)
# 画面から届くのと同じく、JSON を通した形で渡す。
$parsedDetails = (ConvertTo-Json -InputObject $rawDetails -Depth 10) | ConvertFrom-Json
$validated = ConvertTo-PackChangeReportDetails $parsedDetails $detailSummary  # ps51-count-ok（Hashtable を返す）
Assert-ChangeTest ($validated.Count -eq 1 -and $validated.ContainsKey('item:A#2')) "details were not filtered to existing changed rows: $(@($validated.Keys) -join ',')"
$validA = $validated['item:A#2']
# 種類が4つ以外（大文字違いも）・使える枠が無い（文字列の座標・ページの外）の枠は捨てる。regionCount は送られた数。
Assert-ChangeTest (@($validA.regions).Count -eq 2 -and [int]$validA.regionCount -eq 6) "Regions with an unknown kind or without a usable box were not dropped: $(@($validA.regions).Count)"
$addedRegion = @($validA.regions)[1]
Assert-ChangeTest ([double]$addedRegion.before.x -eq 0 -and [double]$addedRegion.before.width -eq 1 -and [double]$addedRegion.after.x -eq 0.3 -and [Math]::Abs([double]$addedRegion.after.width - 0.7) -lt 0.00001) 'Out-of-range coordinates were not clamped to the page.'
Assert-ChangeTest (([string]$addedRegion.afterText).Length -eq 200 -and ([string]$addedRegion.afterText).StartsWith('<script>alert(2)</script>')) 'Region text was not cut to 200 characters.'
$wrongVersion = ConvertTo-PackChangeReportDetails ((ConvertTo-Json -InputObject @([ordered]@{ rowKey = 'item:A#2'; fromVersionId = 'v1'; toVersionId = 'v9'; beforePageNumber = 2; afterPageNumber = 2; regions = @() }) -Depth 6) | ConvertFrom-Json) $detailSummary  # ps51-count-ok（Hashtable を返す）
$wrongPage = ConvertTo-PackChangeReportDetails ((ConvertTo-Json -InputObject @([ordered]@{ rowKey = 'item:A#2'; fromVersionId = 'v1'; toVersionId = 'v2'; beforePageNumber = 3; afterPageNumber = 2; regions = @() }) -Depth 6) | ConvertFrom-Json) $detailSummary  # ps51-count-ok（Hashtable を返す）
$wrongCase = ConvertTo-PackChangeReportDetails ((ConvertTo-Json -InputObject @([ordered]@{ rowKey = 'ITEM:A#2'; fromVersionId = 'v1'; toVersionId = 'v2'; beforePageNumber = 2; afterPageNumber = 2; regions = @() }) -Depth 6) | ConvertFrom-Json) $detailSummary  # ps51-count-ok（Hashtable を返す）
Assert-ChangeTest ($wrongVersion.Count -eq 0 -and $wrongPage.Count -eq 0 -and $wrongCase.Count -eq 0) 'details of another version, another page, or a differently cased row key were accepted.'
$tooMany = ConvertTo-PackChangeReportDetails ((ConvertTo-Json -InputObject @([ordered]@{ rowKey = 'item:A#2'; fromVersionId = 'v1'; toVersionId = 'v2'; beforePageNumber = 2; afterPageNumber = 2; regions = @($manyRegions) }) -Depth 8) | ConvertFrom-Json) $detailSummary  # ps51-count-ok（Hashtable を返す）
Assert-ChangeTest (@($tooMany['item:A#2'].regions).Count -eq 30 -and [int]$tooMany['item:A#2'].regionCount -eq 30) 'More than 30 regions were kept for one row.'
Assert-ChangeTest ((ConvertTo-PackChangeReportDetails 'x' $detailSummary).Count -eq 0 -and (ConvertTo-PackChangeReportDetails $null $detailSummary).Count -eq 0) 'Malformed details were not ignored.'  # ps51-count-ok（Hashtable を返す）
# 全体の大きさの上限を超える分は捨てる。
$savedMax = $Script:PackChangeReportDetailsMaxBytes
$Script:PackChangeReportDetailsMaxBytes = 100
Assert-ChangeTest ((ConvertTo-PackChangeReportDetails $parsedDetails $detailSummary).Count -eq 0) 'details over the size limit were accepted.'  # ps51-count-ok（Hashtable を返す）
$Script:PackChangeReportDetailsMaxBytes = $savedMax

# ---- 座標の換算（前後でページの大きさが違う） ----
# ブラウザ・サーバーの枠は前後のうち大きい方の寸法に左上を合わせた割合。半分の大きさの画像では倍の割合になる。
$halfBox = ConvertTo-PackChangeReportImageBox ([pscustomobject]@{ x = 0.1; y = 0.1; width = 0.2; height = 0.1 }) 1000 1400 500 700
Assert-ChangeTest ([double]$halfBox.x -eq 0.2 -and [double]$halfBox.y -eq 0.2 -and [double]$halfBox.width -eq 0.4 -and [double]$halfBox.height -eq 0.2) "A region was not converted to the smaller page: $($halfBox | ConvertTo-Json -Compress)"
Assert-ChangeTest ($null -eq (ConvertTo-PackChangeReportImageBox ([pscustomobject]@{ x = 0.6; y = 0.1; width = 0.2; height = 0.1 }) 1000 1400 500 700)) 'A region outside the smaller page was drawn on it.'
$sameBox = ConvertTo-PackChangeReportImageBox ([pscustomobject]@{ x = 0.1; y = 0.1; width = 0.2; height = 0.1 }) 1000 1400 1000 1400
Assert-ChangeTest ([double]$sameBox.x -eq 0.1 -and [double]$sameBox.width -eq 0.2) 'A region on a page of the full size was moved.'

# ---- 詳細ページの中身 ----
$imgBefore = [pscustomobject]@{ src = 'img/i0000/0002-b.png'; width = 1000; height = 1400 }
$imgAfterHalf = [pscustomobject]@{ src = 'img/i0000/0002-a.png'; width = 500; height = 700 }
$targetA = [pscustomobject]@{ rowKey = 'item:A#2'; mode = 'diff'; row = $detailSummary.pages[0] }
$pageA = New-PackChangeReportDetailPage $targetA $validA $null $imgBefore $imgAfterHalf
Assert-ChangeTest ([string]$pageA.source -eq 'browser' -and @($pageA.regions).Count -eq 2 -and [int]$pageA.regionCount -eq 6) "The browser detail page is wrong: $($pageA.source) $(@($pageA.regions).Count)"
# 追加の枠は今回の画像だけ、変更の枠は両方（アプリの差分画面と同じ）。
Assert-ChangeTest (@($pageA.before.boxes).Count -eq 1 -and [string]@($pageA.before.boxes)[0].kind -eq 'modified' -and @($pageA.after.boxes).Count -eq 2) 'Added regions must be drawn on the current page only.'
Assert-ChangeTest ([double]@($pageA.after.boxes)[0].x -eq 0.2 -and [double]@($pageA.after.boxes)[0].width -eq 0.4) 'The detail page did not convert regions to the smaller current page.'
Assert-ChangeTest (@($pageA.notes | Where-Object { $_ -like '*最初の2か所だけ*全部で 6 か所*' }).Count -eq 1) 'The detail page did not say that only some regions are shown.'
$fallbackPage = [pscustomobject]@{ width = 1000; height = 1400; regions = @([pscustomobject]@{ kind = 'modified'; x = 0.1; y = 0.2; width = 0.3; height = 0.05 }) }
$pageServer = New-PackChangeReportDetailPage $targetA $null $fallbackPage $imgBefore $imgBefore
Assert-ChangeTest ([string]$pageServer.source -eq 'server' -and @($pageServer.notes | Where-Object { $_ -like '*自動で検出した範囲*' }).Count -eq 1 -and @($pageServer.before.boxes).Count -eq 1) 'The server fallback page was not marked as automatically detected.'
$pageNone = New-PackChangeReportDetailPage $targetA $null $null $imgBefore $imgBefore
Assert-ChangeTest (@($pageNone.regions).Count -eq 0 -and @($pageNone.notes | Where-Object { $_ -like '*特定できませんでした*' }).Count -eq 1) 'A detail page without regions did not ask to check the whole page.'
Assert-ChangeTest ($null -eq (New-PackChangeReportDetailPage $targetA $validA $null $null $null)) 'A detail page was made without any picture.'
$targetB = [pscustomobject]@{ rowKey = 'item:B#1'; mode = 'added'; row = $detailSummary.pages[1] }
$pageB = New-PackChangeReportDetailPage $targetB $null $null $null $imgBefore
Assert-ChangeTest ([string]$pageB.source -eq 'page' -and $null -eq $pageB.before -and @($pageB.after.boxes).Count -eq 1 -and [string]@($pageB.after.boxes)[0].kind -eq 'added' -and [double]@($pageB.after.boxes)[0].width -gt 0.9) 'An added page did not get one picture with a whole-page frame.'
Assert-ChangeTest ($null -eq (New-PackChangeReportDetailPage $targetB $null $null $imgBefore $null)) 'An added page was shown without its current picture.'
Assert-ChangeTest ((Get-PackChangeReportDetailSource @($pageA)) -eq 'browser' -and (Get-PackChangeReportDetailSource @($pageA, $pageServer, $pageB)) -eq 'mixed' -and (Get-PackChangeReportDetailSource @($pageServer, $pageB)) -eq 'server' -and (Get-PackChangeReportDetailSource @($pageB)) -eq 'none') 'detailSource is wrong.'

# ---- 詳細ページの HTML（枠の座標・番号・文字・エスケープ・詳細ページの番号・総ページ数） ----
$detailRows = @(
    (New-TestDetailRow 'item:A#2' 12 10 @('content')),
    (New-TestDetailRow 'item:B#1' 13 0 @('added') 1 0),
    (New-TestDetailRow 'item:D#1' 15 13 @('content'))
)
$detailHtml = New-PackChangeReportHtml (New-TestReportSummary $detailRows $true @{ changed = 3; content = 2; added = 1 }) '一式' '本体' '' @($pageA, $pageB)
Assert-ChangeTest ((Get-TextCount $detailHtml '<section class="sheet') -eq 4 -and $detailHtml.Contains('変更点 4 / 4') -and (Get-TextCount $detailHtml 'class="mark-badge">確認用') -eq 4) 'The detail pages were not counted in the page total.'
Assert-ChangeTest ($detailHtml.Contains('変更 1か所・追加 1か所・ほか 4か所（詳細 p.3）') -and $detailHtml.Contains('今回だけにあるページ（詳細 p.4）')) 'The list does not point to the detail page numbers.'
Assert-ChangeTest ($detailHtml.Contains('<td>変更箇所を特定できませんでした（ページ全体を確認してください）</td>') -and $detailHtml.Contains('1 ページは画像を用意できなかったため')) 'A changed row without a detail page is not explained.'
Assert-ChangeTest ($detailHtml.Contains('今回 p.12（右） ← 前回 p.10（右）／売上.xlsx / 売上（シート内 2/3 ページ）')) 'The detail heading is wrong.'
Assert-ChangeTest ($detailHtml.Contains('<img src="img/i0000/0002-b.png"') -and $detailHtml.Contains('<img src="img/i0000/0002-a.png"')) 'The detail pictures are missing.'
Assert-ChangeTest ($detailHtml.Contains('class="box modified" style="left:10%;top:10%;width:20%;height:10%"><span class="tag">①変更</span>') -and $detailHtml.Contains('class="box modified" style="left:20%;top:20%;width:40%;height:20%"><span class="tag">①変更</span>')) 'The numbered frames are not placed at the converted coordinates.'
Assert-ChangeTest ($detailHtml.Contains('<span class="tag">②追加</span>') -and $detailHtml.Contains('<li class="item">① 変更: 1,234 → 1,240</li>')) 'The numbered list of changed text is missing.'
Assert-ChangeTest (-not $detailHtml.Contains('<script>alert(2)') -and $detailHtml.Contains('② 追加: &lt;script&gt;alert(2)&lt;/script&gt;')) 'Changed text was not escaped in the report HTML.'
Assert-ChangeTest (-not $detailHtml.Contains('くわしく見る') -and -not $detailHtml.Contains('line-clamp') -and -not $detailHtml.Contains('class="clip"')) 'The report still refers to the app or clips list cells.'
Assert-ChangeTest ($detailHtml.Contains('p.3 から 2 ページ分')) 'The cover does not point to the detail pages.'

# ---- 詳細ページの寸法（PR-A: 凡例は見出しの行、画像の枠は文字の一覧に合わせて 110〜158mm） ----
Assert-ChangeTest ([int]$Script:PackChangeReportImageDpi -eq 120 -and [int]$Script:PackChangeReportVersion -eq 2) 'The detail pictures must be 120dpi without making existing reports outdated.'
Assert-ChangeTest ([int]$pageA.before.width -eq 1000 -and [int]$pageA.before.height -eq 1400 -and [int]$pageA.after.width -eq 500 -and [int]$pageA.after.height -eq 700) 'The detail page must keep the picture size for the HTML.'
# 見積もりの合計（印＋見出し＋左右の見出し＋画像の枠＋下の余白＋文字の一覧＋余裕）が本文 186mm を超えない。
function Get-TestDetailTotal($Size) {
    return [double]$Script:PackChangeReportMarkMm + [double]$Size.headingHeight + [double]$Script:PackChangeReportSideHeadMm + [double]$Size.picHeight + [double]$Script:PackChangeReportPairMarginMm + [double]$Size.firstBudget + [double]$Script:PackChangeReportDetailSlackMm
}
$sizeShort = Get-PackChangeReportDetailLayout 'x' @(10)
Assert-ChangeTest ([Math]::Abs([double]$sizeShort.picHeight - 147.7) -lt 0.001 -and [Math]::Abs([double]$sizeShort.firstBudget - 10) -lt 0.001) "A short list must leave the rest of the page to the pictures: $($sizeShort | ConvertTo-Json -Compress)"
$sizeEmpty = Get-PackChangeReportDetailLayout 'x' @()
Assert-ChangeTest ([double]$sizeEmpty.picHeight -le [double]$Script:PackChangeReportPicMaxMm -and [double]$sizeEmpty.picHeight -gt 150) "Without a list the pictures must be as tall as the page allows: $($sizeEmpty.picHeight)"
$sizeLong = Get-PackChangeReportDetailLayout ('長' * 400) @(1..40 | ForEach-Object { 5.3 })
Assert-ChangeTest ([double]$sizeLong.picHeight -eq [double]$Script:PackChangeReportPicMinMm) "A long list must shrink the pictures to the minimum: $($sizeLong.picHeight)"
foreach ($size in @($sizeShort, $sizeEmpty, $sizeLong, (Get-PackChangeReportDetailLayout ('長' * 120) @(30, 30)))) {
    Assert-ChangeTest ((Get-TestDetailTotal $size) -le [double]$Script:PackChangeReportBodyHeightMm + 0.001) "A detail page is taller than the sheet: $(Get-TestDetailTotal $size)"
    Assert-ChangeTest (([double]$Script:PackChangeReportMarkMm + [double]$size.headingHeight + [double]$size.moreBudget + [double]$Script:PackChangeReportDetailSlackMm) -le [double]$Script:PackChangeReportBodyHeightMm + 0.001) 'A continued detail page is taller than the sheet.'
}
# 縦長は高さで、横長は幅で決まる（左右1枚の枠は (273−5)/2＝134mm、線のぶん 1mm 引く）。
$portraitMm = Get-PackChangeReportImageMm 1000 1400 157.7
Assert-ChangeTest ([Math]::Abs([double]$portraitMm.height - 156.7) -lt 0.001 -and [Math]::Abs([double]$portraitMm.width - 111.9) -lt 0.001) "A portrait picture must fill the height: $($portraitMm | ConvertTo-Json -Compress)"
$landscapeMm = Get-PackChangeReportImageMm 1400 1000 157.7
Assert-ChangeTest ([Math]::Abs([double]$landscapeMm.width - 133) -lt 0.001 -and [Math]::Abs([double]$landscapeMm.height - 95) -lt 0.001) "A landscape picture must fill the width: $($landscapeMm | ConvertTo-Json -Compress)"
Assert-ChangeTest ($null -eq (Get-PackChangeReportImageMm 0 1400 157.7)) 'A picture of unknown size must fall back to the CSS limits.'
# HTML: 凡例は見出しの行、枠と画像の高さは見積もりと同じ値、手で書いた 96mm／95mm は無い。
$layoutA = Get-PackChangeReportDetailLayout (Get-PackChangeReportDetailHeading $detailRows[0]) @(Get-PackChangeReportDetailLines $pageA | ForEach-Object { (Get-PackChangeReportLineCount ([string]$_.text) 78.0) * 4.7 + 0.6 })
$imageAmm = Get-PackChangeReportImageMm 1000 1400 $layoutA.picHeight
Assert-ChangeTest ($detailHtml.Contains('<div class="dhr"><h2 class="dh">今回 p.12（右）') -and $detailHtml.Contains('</h2><p class="legend">')) 'The legend must sit on the heading line.'
Assert-ChangeTest ($detailHtml.Contains('<div class="pic" style="height:' + (Format-PackChangeReportMm $layoutA.picHeight) + '">')) "The picture frame height does not match the estimate: $($layoutA.picHeight)"
Assert-ChangeTest ($detailHtml.Contains('<img src="img/i0000/0002-b.png" alt="" style="width:' + (Format-PackChangeReportMm $imageAmm.width) + ';height:' + (Format-PackChangeReportMm $imageAmm.height) + '">')) 'The picture size does not match the frame.'
Assert-ChangeTest (-not $detailHtml.Contains('height:96mm') -and -not $detailHtml.Contains('max-height:95mm') -and $detailHtml.Contains('max-height:' + (Format-PackChangeReportMm ($Script:PackChangeReportPicMaxMm - 1.0)))) 'The CSS must use the shared sizes.'
Assert-ChangeTest ($detailHtml.Contains('<div class="pic empty" style="height:')) 'A missing picture must keep the same frame height.'

# 一覧のセルは3行で切らない。長い名前の行は高さを見積もってページを分け、名前はそのまま載せる。
$longName = ('とても長い原稿名' * 25)
$tallRows = @(for ($p = 1; $p -le 12; $p++) { New-TestReportRow $p $p @('content') $false 'right' 'right' $(if ($p -eq 1) { $longName } else { '売上.xlsx' }) })
$tallHtml = New-PackChangeReportHtml (New-TestReportSummary $tallRows) '一式' '本体'
Assert-ChangeTest ((Get-TextCount $tallHtml '<section class="sheet') -eq 3 -and $tallHtml.Contains('変更点 3 / 3') -and $tallHtml.Contains($longName)) 'A tall row was clipped or did not push the list onto another page.'
$hugeName = ('長' * 3000)
$hugeHtml = New-PackChangeReportHtml (New-TestReportSummary @((New-TestReportRow 1 1 @('content') $false 'right' 'right' $hugeName))) '一式' '本体'
Assert-ChangeTest ($hugeHtml.Contains('…（長いため省略しました）') -and -not $hugeHtml.Contains($hugeName)) 'A cell taller than a page was not shortened visibly.'
Assert-ChangeTest ([Math]::Abs((Get-PackChangeReportTextUnits 'ab全') - 2.2) -lt 0.0001) 'Half-width characters must count as narrower than full-width ones.'
$split = Split-PackChangeReportPages @(10, 10, 10, 30) 15 25 0
Assert-ChangeTest ($split.pages.Count -eq 3 -and @($split.pages[0]).Count -eq 1 -and @($split.pages[1]).Count -eq 2 -and @($split.pages[2])[0] -eq 3) 'The height-based page split is wrong.'
$splitFirst = Split-PackChangeReportPages @(20) 15 25 0
Assert-ChangeTest ($splitFirst.pages.Count -eq 2 -and @($splitFirst.pages[0]).Count -eq 0) 'An item too tall for the first page must move to the next page.'

# 文字の一覧が1ページに入りきらない詳細は「続き」のページへ回し、総ページ数も合わせる。
$wordy = [pscustomobject]@{ rowKey = 'item:A#2'; beforePageNumber = 2; afterPageNumber = 2; regionCount = 30; regions = @(for ($n = 0; $n -lt 30; $n++) { [pscustomobject]@{ kind = 'modified'; before = [pscustomobject]@{ x = 0.1; y = 0.1; width = 0.1; height = 0.1 }; after = [pscustomobject]@{ x = 0.1; y = 0.1; width = 0.1; height = 0.1 }; beforeText = ('前' * 200); afterText = ('後' * 200) } }) }
$wordyPage = New-PackChangeReportDetailPage $targetA $wordy $null $imgBefore $imgBefore
$wordyHtml = New-PackChangeReportHtml (New-TestReportSummary @($detailRows[0])) '一式' '本体' '' @($wordyPage)
$wordySheets = Get-TextCount $wordyHtml '<section class="sheet'
Assert-ChangeTest ($wordySheets -gt 3 -and $wordyHtml.Contains(('変更点 {0} / {0}' -f $wordySheets)) -and $wordyHtml.Contains('（続き）') -and (Get-TextCount $wordyHtml '<li class="item">') -eq 30) "A long list of changed text was not continued on extra pages: $wordySheets sheets"
Assert-ChangeTest ($wordyHtml.Contains('㉚変更') -and $wordyHtml.Contains('⑳変更')) 'Frames beyond 20 must keep readable numbers.'

# 200件の上限: 201件目からは詳細を載せず、表紙と一覧でそう伝える。
$limitRows = @(for ($p = 1; $p -le 205; $p++) { New-TestDetailRow ('item:L' + $p + '#1') $p $p @('content') })
$limitSelection = Get-PackChangeReportDetailTargets @(Get-PackChangeReportEntries $limitRows)
Assert-ChangeTest (@($limitSelection.targets).Count -eq 200 -and @($limitSelection.omitted).Count -eq 5 -and [string]@($limitSelection.omitted)[0] -eq 'item:L201#1') 'The detail pages were not limited to the first 200 rows.'
$limitHtml = New-PackChangeReportHtml (New-TestReportSummary $limitRows) '一式' '本体'
Assert-ChangeTest ($limitHtml.Contains('ページごとの詳細は最初の200ページ分だけ載せています（ほかの 5 ページは一覧だけです）') -and (Get-TextCount $limitHtml '詳細は省略しました') -eq 5) 'The 200-page limit is not explained.'
# ずれただけ・左右の入れ替わりだけの行には詳細ページを作らない。増えた・なくなった行は作る。
$kindSelection = Get-PackChangeReportDetailTargets @(Get-PackChangeReportEntries @((New-TestDetailRow 'item:M#1' 3 3 @('moved','side')), (New-TestDetailRow 'item:R#1' 0 4 @('removed') 0 1), (New-TestDetailRow 'item:N#1' 5 0 @('added') 1 0)))
Assert-ChangeTest ((@($kindSelection.targets | ForEach-Object { [string]$_.mode }) -join ',') -eq 'removed,added') 'Detail targets must be changed content, added and removed pages only.'
# 前回の画像は前回の原稿・シートで引く（原稿の差し替え・シート名の変更）。古い summary.json は同じ原稿と分かるときだけ。
$renamedRow = New-TestDetailRow 'item:A#2' 12 10 @('content')
$renamedRow.fromSheetName = '旧売上'; $renamedRow.fromSourceId = 'wb_old'
$renamedSides = Get-PackChangeReportImageSides $renamedRow 'diff'
Assert-ChangeTest ([string]$renamedSides.before.sheetName -eq '旧売上' -and [string]$renamedSides.before.workbookId -eq 'wb_old' -and [string]$renamedSides.after.sheetName -eq '売上' -and [int]$renamedSides.after.pageNumber -eq 2 -and [string]$renamedSides.before.versionId -eq 'v1') 'The previous picture must come from the previous source and sheet.'
$legacyRow = New-TestDetailRow 'item:A#2' 12 10 @('content')
$legacyRow.Remove('fromSourceId'); $legacyRow.Remove('fromSheetName'); $legacyRow.detailAvailable = $false
Assert-ChangeTest ($null -eq (Get-PackChangeReportImageSides $legacyRow 'diff').before) 'An old summary row of another source must not guess the previous source.'
Assert-ChangeTest ($null -eq (Get-PackChangeReportImageSides (New-TestDetailRow 'item:N#1' 5 0 @('added') 1 0) 'added').before) 'An added page has no previous picture.'

# ---- Edge のヘッドレス印刷 ---------------------------------------------------------
# 専用の --user-data-dir を必ず渡す（起動中の Edge にのっとられると PDF ができない）。
$edgeArgs = @(Get-EdgePrintToPdfArguments 'C:\t\profile' 'C:\t\report.pdf' 'file:///C:/t/report.html')
Assert-ChangeTest ($edgeArgs -contains '--headless' -and $edgeArgs -contains '--user-data-dir=C:\t\profile' -and $edgeArgs -contains '--no-pdf-header-footer' -and $edgeArgs -contains '--print-to-pdf=C:\t\report.pdf' -and $edgeArgs[-1] -eq 'file:///C:/t/report.html') "The Edge print arguments are wrong: $($edgeArgs -join ' ')"
Assert-Throws { Get-EdgePrintToPdfArguments '' 'C:\t\report.pdf' 'file:///x' } 'プロファイル' 'Edge was allowed to print with the default profile.'
# 待ち時間の上限は件数と詳細ページの画像の枚数（120dpi で1枚 2.5 秒）に比例させ、最大15分。
Assert-ChangeTest ((Get-PackChangeReportTimeoutSeconds 0) -eq 60 -and (Get-PackChangeReportTimeoutSeconds 100) -eq 110 -and (Get-PackChangeReportTimeoutSeconds 2000) -eq 900) 'The report timeout is not proportional to the item count.'
Assert-ChangeTest ((Get-PackChangeReportTimeoutSeconds 100 40) -eq 210 -and (Get-PackChangeReportTimeoutSeconds 100 300) -eq 860 -and (Get-PackChangeReportTimeoutSeconds 100 400) -eq 900) 'The report timeout does not count the detail pictures.'

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
# 実装は GetFullPath で返すので、比べる側もそろえる（一時フォルダーの 8.3 形式 RUNNER~1 などの違いを比べない）。
Assert-ChangeTest ((Resolve-PackChangeReportFile 'ja' $packId 'main' $build3 $build4) -eq [IO.Path]::GetFullPath((Join-Path $reportDir 'report.pdf'))) 'The made report cannot be opened.'
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
Assert-ChangeTest ([bool]$attachCheck.ok -and [string]$attachCheck.fromBuildId -eq $build2 -and [string]$attachCheck.sourceFull -eq [IO.Path]::GetFullPath((Join-Path (Get-PackChangeReportDir 'ja' $packId 'main' $build2 $build6) 'report.pdf'))) 'The publish check disagrees with the changes response.'
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

# ---- ページごとの詳細を載せる作成ジョブ（偽の画像化・lease で、実際に呼ばれる経路を通す） -------------
# 新しい2つの版: シートA の2ページ目だけ内容が変わり、シートN が1ページ増えた。
foreach ($snapshotId in @('s7','s8')) { New-Item -ItemType Directory -Path (Get-SnapshotDir 'ja' 'wb_a' $snapshotId) -Force | Out-Null }
Write-TestVisualHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') }
Write-TestVisualHashes 's8' 'v8' @{ 'シートA' = @('K1','K2X'); 'シートN' = @('N1') }
$build7 = Set-ArchivedOutput 'ja-main' '%PDF-change-7' (New-TestManifest @((New-TestItem 'A' 2))) (New-TestMetadata @{ wb_a = 's7|v7' })
$build8 = Set-ArchivedOutput 'ja-main' '%PDF-change-8' (New-TestManifest @((New-TestItem 'A' 2), (New-TestItem 'N'))) (New-TestMetadata @{ wb_a = 's8|v8' })
$detailJobSummary = Get-PackChangeSummary 'ja' $packId 'main' $build7 $build8
$changedRow = @($detailJobSummary.pages | Where-Object { [string]$_.rowKey -eq 'item:A#2' })[0]
Assert-ChangeTest ((Get-Kinds $changedRow) -eq 'content' -and [string]$changedRow.fromSourceId -eq 'wb_a' -and [string]$changedRow.fromSheetName -eq 'シートA') "The detail job fixture is wrong: $(Get-Kinds $changedRow)"

function Write-FakePng([string]$Path, [int]$Width, [int]$Height) {
    $bytes = New-Object byte[] 33
    $head = @(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52)
    for ($i = 0; $i -lt $head.Count; $i++) { $bytes[$i] = [byte]$head[$i] }
    $bytes[18] = [byte][Math]::Floor($Width / 256); $bytes[19] = [byte]($Width % 256)
    $bytes[22] = [byte][Math]::Floor($Height / 256); $bytes[23] = [byte]($Height % 256)
    [IO.File]::WriteAllBytes($Path, $bytes)
}
$Script:FakeLeaseVersions = New-Object System.Collections.ArrayList
$Script:FakeLeaseReleased = 0
$Script:FakePdfCalls = New-Object System.Collections.ArrayList
$Script:FakeBatchDpi = 0
$Script:FakeBatchMode = 'ok'
function New-PackChangeReportLeases([string]$Language, $Versions, [string]$JobId) {
    $out = @()
    foreach ($v in @($Versions)) {
        [void]$Script:FakeLeaseVersions.Add(([string]$v.snapshotId + '|' + [string]$v.versionId))
        $out += [pscustomobject]@{ workbookId = [string]$v.workbookId; snapshotId = [string]$v.snapshotId; versionId = [string]$v.versionId; snapshotLeaseName = 'fake'; contentLeaseName = 'fake' }
    }
    return $out
}
function Remove-PackChangeReportLeases([string]$Language, $Leases) { $Script:FakeLeaseReleased += @($Leases).Count }
function Resolve-DiffContentPdfPath([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId, [string]$SheetName) {
    [void]$Script:FakePdfCalls.Add(($WorkbookId + '|' + $SnapshotId + '|' + $VersionId + '|' + $SheetName))
    return ('C:\fake\' + $VersionId + '-' + $SheetName + '.pdf')
}
# diff-image-batch.ps1 の代わり。pagePairs（行の組）で頼まれたページだけを、頼まれた順に 1, 2, … と番号を振って返す
# （本物と同じ契約）。シートA の2ページ目の今回は半分の大きさ（用紙サイズ違い）。頼まれた中身は FakeBatchItems に残す。
$Script:FakeBatchItems = New-Object System.Collections.ArrayList
function Invoke-DiffImageBatchGeneration($Items, [int]$Dpi = 0) {
    $Script:FakeBatchDpi = $Dpi
    foreach ($item in @($Items)) { [void]$Script:FakeBatchItems.Add(((ConvertTo-Json -InputObject $item -Depth 10) | ConvertFrom-Json)) }
    if ($Script:FakeBatchMode -eq 'throw') { throw '差分画像の一括生成に失敗しました。' }
    $results = @()
    foreach ($item in @($Items)) {
        $dir = [string]$item.outputDirectory
        $pages = @()
        $number = 0
        foreach ($pair in @($item.pagePairs)) {
            $number++
            $stem = $number.ToString('0000')
            $b = [int]$pair.beforePageNumber; $a = [int]$pair.afterPageNumber
            $beforeFile = ''; $afterFile = ''
            if ($b -gt 0) { $beforeFile = $stem + '-b.png'; Write-FakePng (Join-Path $dir $beforeFile) 1000 1400 }
            if ($a -gt 0) {
                $afterFile = $stem + '-a.png'
                if ([string]$item.kind -eq 'modified' -and $a -eq 2) { Write-FakePng (Join-Path $dir $afterFile) 500 700 } else { Write-FakePng (Join-Path $dir $afterFile) 1000 1400 }
            }
            $regions = @()
            if ($b -eq 2 -and $a -eq 2) { $regions = @([pscustomobject]@{ kind = 'modified'; x = 0.1; y = 0.2; width = 0.3; height = 0.05 }) }
            $pages += [pscustomobject]@{ pageNumber = $number; width = 1000; height = 1400; beforePageNumber = $b; afterPageNumber = $a; beforeFile = $beforeFile; afterFile = $afterFile; status = 'ready'; regions = @($regions) }
        }
        $results += [pscustomobject]@{ id = [string]$item.id; ok = $true; pages = @($pages) }
    }
    return [pscustomobject]@{ ok = $true; items = @($results) }
}
$detailReportDir = Get-PackChangeReportDir 'ja' $packId 'main' $build7 $build8
function Invoke-DetailReportJob($Details) {
    $job = [pscustomobject]@{ jobId = $reportJobId; mode = 'ja'; packId = $packId; targetId = 'main'; fromBuildId = $build7; toBuildId = $build8; statusPath = $reportStatusPath; details = $Details }
    Invoke-PackChangeReportJobCore $job
    $state = Read-JsonFile $reportStatusPath $null
    Assert-ChangeTest ([string]$state.status -eq 'completed') "The detail report job failed: $($state.status) $($state.message) $(@($state.errors | ForEach-Object { [string]$_.detail }) -join ' | ')"
    return [pscustomobject]@{
        status = $state
        meta = (Read-JsonFile (Join-Path $detailReportDir 'meta.json') $null)
        html = [IO.File]::ReadAllText((Join-Path $detailReportDir 'report.html'), [Text.Encoding]::UTF8)
    }
}

# 1) 画面の計算（details）がある行は、その枠と文字をそのまま使う（ジョブの入力は JSON を通った形）。
$jobDetails = (ConvertTo-Json -InputObject @(
    [ordered]@{ rowKey = 'item:A#2'; fromVersionId = 'v7'; toVersionId = 'v8'; beforePageNumber = 2; afterPageNumber = 2; regionCount = 1; regions = @(
        [ordered]@{ kind = 'modified'; before = [ordered]@{ x = 0.1; y = 0.1; width = 0.2; height = 0.1 }; after = [ordered]@{ x = 0.1; y = 0.1; width = 0.2; height = 0.1 }; beforeText = '1,234'; afterText = '<b>1,240</b>' }) },
    [ordered]@{ rowKey = 'item:N#1'; fromVersionId = ''; toVersionId = 'v8'; beforePageNumber = 0; afterPageNumber = 1; regions = @() }
) -Depth 10) | ConvertFrom-Json
$browserRun = Invoke-DetailReportJob $jobDetails
Assert-ChangeTest ([int]$browserRun.meta.reportVersion -eq 2 -and [string]$browserRun.meta.detailSource -eq 'browser' -and [int]$browserRun.meta.detailPageCount -eq 2) "meta.json does not describe the detail pages: $($browserRun.meta | ConvertTo-Json -Compress)"
Assert-ChangeTest ([string]$browserRun.status.report.detailSource -eq 'browser') 'The made report does not report where its frames came from.'
Assert-ChangeTest ($browserRun.html.Contains('<li class="item">① 変更: 1,234 → &lt;b&gt;1,240&lt;/b&gt;</li>') -and -not $browserRun.html.Contains('<b>1,240</b>')) 'The detail page does not list the changed text (escaped).'
Assert-ChangeTest ($browserRun.html.Contains('style="left:10%;top:10%;width:20%;height:10%"') -and $browserRun.html.Contains('style="left:20%;top:20%;width:40%;height:20%"')) 'Browser regions were not converted to the picture of each side (paper size changed).'
Assert-ChangeTest ($browserRun.html.Contains('<img src="img/i0000/0001-b.png"') -and $browserRun.html.Contains('<img src="img/i0000/0001-a.png"') -and $browserRun.html.Contains('今回だけにあるページ（詳細 p.4）') -and $browserRun.html.Contains('（詳細 p.3）')) 'The detail pages or their page numbers are missing.'
Assert-ChangeTest (-not $browserRun.html.Contains('自動で検出した範囲')) 'Browser regions were labelled as automatically detected.'
# 画像は詳細ページ用の解像度で、記録された版（v7・v8）の変換PDFから作り、lease は作り終えたら外す。
Assert-ChangeTest ([int]$Script:FakeBatchDpi -eq [int]$Script:PackChangeReportImageDpi) 'The detail pictures were not rendered at the report resolution.'
Assert-ChangeTest ($Script:FakePdfCalls -contains 'wb_a|s7|v7|シートA' -and $Script:FakePdfCalls -contains 'wb_a|s8|v8|シートA' -and $Script:FakePdfCalls -contains 'wb_a|s8|v8|シートN') "The pictures were not made from the archived versions: $(@($Script:FakePdfCalls) -join ', ')"
Assert-ChangeTest ((@($Script:FakeLeaseVersions | Sort-Object -Unique) -join ',') -eq 's7|v7,s8|v8' -and [int]$Script:FakeLeaseReleased -eq 2) "The versions were not leased while the pictures were made: $(@($Script:FakeLeaseVersions) -join ',') released=$($Script:FakeLeaseReleased)"
# 画像化・比較は行の組だけ（シートA は前回 p.2 ↔ 今回 p.2、シートN は今回 p.1 だけ）。シートの全ページは頼まない。
# 版の解析の画像の記録（visualHashProfile）が無い版は、保存画像を使わずに変換PDFから画像化する。
$batchA = @($Script:FakeBatchItems | Where-Object { [string]$_.afterPdf -like '*v8-シートA.pdf' } | Select-Object -Last 1)
$batchN = @($Script:FakeBatchItems | Where-Object { [string]$_.afterPdf -like '*v8-シートN.pdf' } | Select-Object -Last 1)
Assert-ChangeTest ($batchA.Count -eq 1 -and (@($batchA[0].pagePairs | ForEach-Object { '{0}-{1}' -f $_.beforePageNumber, $_.afterPageNumber }) -join ',') -eq '2-2') "The batch was not asked for the row's page pair only: $(ConvertTo-Json -InputObject $batchA -Depth 6 -Compress)"
Assert-ChangeTest ($batchN.Count -eq 1 -and [string]$batchN[0].kind -eq 'added' -and (@($batchN[0].pagePairs | ForEach-Object { '{0}-{1}' -f $_.beforePageNumber, $_.afterPageNumber }) -join ',') -eq '0-1') "An added page was not asked as an after-only pair: $(ConvertTo-Json -InputObject $batchN -Depth 6 -Compress)"
Assert-ChangeTest ([string]$batchA[0].beforeRasterDirectory -eq '' -and [string]$batchA[0].afterRasterDirectory -eq '' -and [int]$batchA[0].beforePageCount -eq 0) 'Saved page pictures were used for versions whose analysis did not record how they were drawn.'
# 待ち時間の上限に画像の枚数（前回・今回の2枚＋増えたページの1枚）を数える。
Assert-ChangeTest ([int]$Script:EdgeCalls[$Script:EdgeCalls.Count - 1].timeout -eq (Get-PackChangeReportTimeoutSeconds @($detailJobSummary.pages).Count 3)) 'The report timeout does not count the detail pictures.'

# 2) details が無い（画面で計算できなかった）行は、サーバーの画素比較の枠で埋めて印を付ける。
$serverRun = Invoke-DetailReportJob @()
Assert-ChangeTest ([string]$serverRun.meta.detailSource -eq 'server' -and $serverRun.html.Contains('枠は、前回と今回のページの画像の違いから自動で検出した範囲です。')) 'Rows without browser details were not filled with the server frames.'
Assert-ChangeTest ($serverRun.html.Contains('style="left:10%;top:20%;width:30%;height:5%"') -and $serverRun.html.Contains('style="left:20%;top:40%;width:60%;height:10%"')) 'The server frames were not converted to the picture of each side.'
# 不正な details（版の違う結果）は捨てて、予備の枠を使う。
$forged = (ConvertTo-Json -InputObject @([ordered]@{ rowKey = 'item:A#2'; fromVersionId = 'v1'; toVersionId = 'v8'; beforePageNumber = 2; afterPageNumber = 2; regions = @([ordered]@{ kind = 'modified'; before = [ordered]@{ x = 0.5; y = 0.5; width = 0.1; height = 0.1 }; after = [ordered]@{ x = 0.5; y = 0.5; width = 0.1; height = 0.1 }; beforeText = 'X'; afterText = 'Y' }) }) -Depth 10) | ConvertFrom-Json
$forgedRun = Invoke-DetailReportJob $forged
Assert-ChangeTest ([string]$forgedRun.meta.detailSource -eq 'server' -and -not $forgedRun.html.Contains('X → Y')) 'details of another version were used for the report.'

# 3) 画像を作れなくても変更点PDFは作り、詳細の無いことを伝える（作成は失敗にしない）。
$Script:FakeBatchMode = 'throw'
$failedImages = Invoke-DetailReportJob $jobDetails
Assert-ChangeTest ([int]$failedImages.meta.detailPageCount -eq 0 -and [string]$failedImages.meta.detailSource -eq 'none' -and $failedImages.html.Contains('2 ページは画像を用意できなかったため')) 'A report without pictures did not explain the missing details.'
# @($null).Count も 1 なので、件数だけでなく中身（code）も見る。
Assert-ChangeTest (@($failedImages.status.warnings).Count -eq 1 -and [string]@($failedImages.status.warnings)[0].code -eq 'detail-images' -and [int]$Script:FakeLeaseReleased -eq 8) "The picture failure was not kept as a warning or the leases were not released: released=$($Script:FakeLeaseReleased)"
$Script:FakeBatchMode = 'ok'

# 版の解析が保存したページ画像（raster-v1）を使うのは、記録が同じ解像度・色・PDFBox で、シートが解析でき、
# ページ数がわかり、フォルダーがあるときだけ。どれか欠ければ $null（変換PDFから画像化する）。
function Write-TestRasterHashes([string]$Snapshot, [string]$Version, [hashtable]$Sheets, $HashProfile, [string]$Status = 'ok') {
    $list = @()
    foreach ($name in @($Sheets.Keys)) { $list += [ordered]@{ sheetName = [string]$name; status = $Status; pageCount = @($Sheets[$name]).Count; pageHashes = @($Sheets[$name]) } }
    $value = [ordered]@{ schemaVersion = 1; sheets = @($list) }
    if ($null -ne $HashProfile) { $value.visualHashProfile = $HashProfile }
    Write-JsonFile (Join-Path (Get-RenderRecordDir 'ja' 'wb_a' $Snapshot $Version) 'visual-hashes.json') $value
    $Script:VisualHashCache.Clear()
}
function Copy-TestProfile($HashProfile, [string]$Name, $Value) {
    $copy = [ordered]@{}
    foreach ($k in @($HashProfile.Keys)) { $copy[$k] = $HashProfile[$k] }
    $copy[$Name] = $Value
    return $copy
}
$rasterSide = [pscustomobject]@{ workbookId = 'wb_a'; snapshotId = 's7'; versionId = 'v7'; sheetName = 'シートA'; pageNumber = 2 }
$rasterDir = Get-RenderRasterSheetDir 'ja' 'wb_a' 's7' 'v7' 'シートA'
$goodProfile = Get-VisualHashProfile
Write-TestRasterHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') } $goodProfile
Assert-ChangeTest ($null -eq (Get-PackChangeReportRasterCache 'ja' $rasterSide)) 'Saved page pictures were used although the picture folder does not exist.'
New-Item -ItemType Directory -Path $rasterDir -Force | Out-Null
Write-FakePng (Join-Path $rasterDir 'page-0001.png') 1000 1400; Write-FakePng (Join-Path $rasterDir 'page-0002.png') 1000 1400
$rasterCache = Get-PackChangeReportRasterCache 'ja' $rasterSide
Assert-ChangeTest ($null -ne $rasterCache -and [string]$rasterCache.directory -eq $rasterDir -and [int]$rasterCache.pageCount -eq 2) "Saved page pictures of an analysed version were not offered: $(ConvertTo-Json -InputObject $rasterCache -Compress)"
Write-TestRasterHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') } (Copy-TestProfile $goodProfile 'dpi' 150)
Assert-ChangeTest ($null -eq (Get-PackChangeReportRasterCache 'ja' $rasterSide)) 'Saved page pictures of another resolution were used.'
Write-TestRasterHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') } (Copy-TestProfile $goodProfile 'pdfBoxVersion' 'older')
Assert-ChangeTest ($null -eq (Get-PackChangeReportRasterCache 'ja' $rasterSide)) 'Saved page pictures drawn by another PDFBox were used.'
Write-TestRasterHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') } (Copy-TestProfile $goodProfile 'colorMode' 'GRAY')
Assert-ChangeTest ($null -eq (Get-PackChangeReportRasterCache 'ja' $rasterSide)) 'Saved page pictures of another color mode were used.'
Write-TestRasterHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') } $goodProfile 'error'
Assert-ChangeTest ($null -eq (Get-PackChangeReportRasterCache 'ja' $rasterSide)) 'Saved page pictures of a sheet that failed analysis were used.'
Write-TestRasterHashes 's7' 'v7' @{ 'シートB' = @('K1','K2') } $goodProfile
Assert-ChangeTest ($null -eq (Get-PackChangeReportRasterCache 'ja' $rasterSide)) 'Saved page pictures of another sheet were used.'
# 作成ジョブは、使える側だけ保存画像のフォルダーとページ数をバッチに渡す（今回の版 v8 は記録が無いので渡さない）。
Write-TestRasterHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') } $goodProfile
$null = Invoke-DetailReportJob $jobDetails
$cachedBatch = @($Script:FakeBatchItems | Where-Object { [string]$_.afterPdf -like '*v8-シートA.pdf' } | Select-Object -Last 1)
Assert-ChangeTest ([string]$cachedBatch[0].beforeRasterDirectory -eq $rasterDir -and [int]$cachedBatch[0].beforePageCount -eq 2 -and [string]$cachedBatch[0].afterRasterDirectory -eq '' -and [int]$cachedBatch[0].afterPageCount -eq 0) "The job did not pass the saved page pictures of the analysed version: $(ConvertTo-Json -InputObject $cachedBatch -Depth 6 -Compress)"
Write-TestVisualHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') }
$Script:VisualHashCache.Clear()

# 版 1 の作成済みPDF（ページごとの詳細の無いもの）は、作り直すまで作成済みにしない。
$detailMetaPath = Join-Path $detailReportDir 'meta.json'
$detailMeta = Read-JsonFile $detailMetaPath $null
$detailStateArgs = @('ja', $packId, 'main', $build7, $build8, [string]$detailJobSummary.from.outputPdfSha256, [string]$detailJobSummary.to.outputPdfSha256)
Assert-ChangeTest ([bool](Get-PackChangeReportState @detailStateArgs $true).exists) 'A report of the current version was not offered.'
$detailMeta.reportVersion = 1
Write-JsonFile $detailMetaPath $detailMeta
Assert-ChangeTest (-not [bool](Get-PackChangeReportState @detailStateArgs $true).exists -and [string](Get-PackChangeReportState @detailStateArgs $true).reason -eq 'outdated') 'A report made before the detail pages (version 1) was still offered.'

# 原稿のページ数が変わった行（シートの途中でページが増えた）は、行の組が中身のずれた2枚になるので、
# サーバーの予備の枠を使わない（画面の結果が無ければ「特定できませんでした」）。
New-Item -ItemType Directory -Path (Get-SnapshotDir 'ja' 'wb_a' 's9') -Force | Out-Null
Write-TestVisualHashes 's9' 'v9' @{ 'シートA' = @('K1','K2X','K3') }
$build9 = Set-ArchivedOutput 'ja-main' '%PDF-change-9' (New-TestManifest @((New-TestItem 'A' 3))) (New-TestMetadata @{ wb_a = 's9|v9' })
$grownSummary = Get-PackChangeSummary 'ja' $packId 'main' $build7 $build9
$grownRow = @($grownSummary.pages | Where-Object { [string]$_.rowKey -eq 'item:A#2' })[0]
Assert-ChangeTest ((Get-Kinds $grownRow) -eq 'content' -and [bool]$grownRow.partCountChanged) "The grown-sheet fixture is wrong: $(Get-Kinds $grownRow)"
Invoke-PackChangeReportJobCore ([pscustomobject]@{ jobId = $reportJobId; mode = 'ja'; packId = $packId; targetId = 'main'; fromBuildId = $build7; toBuildId = $build9; statusPath = $reportStatusPath; details = @() })
$grownState = Read-JsonFile $reportStatusPath $null
Assert-ChangeTest ([string]$grownState.status -eq 'completed') "The grown-sheet report job failed: $($grownState.message)"
$grownHtml = [IO.File]::ReadAllText((Join-Path (Get-PackChangeReportDir 'ja' $packId 'main' $build7 $build9) 'report.html'), [Text.Encoding]::UTF8)
Assert-ChangeTest (-not $grownHtml.Contains('自動で検出した範囲') -and $grownHtml.Contains('変更箇所を特定できませんでした')) 'Server frames were used for a row whose source page count changed.'
# 描き方の版（profileVersion）が違う解析の画像も使わない。
Write-TestRasterHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') } (Copy-TestProfile $goodProfile 'profileVersion' ([int]$goodProfile.profileVersion - 1))
Assert-ChangeTest ($null -eq (Get-PackChangeReportRasterCache 'ja' $rasterSide)) 'Saved page pictures of another visual-hash profile were used.'
Write-TestVisualHashes 's7' 'v7' @{ 'シートA' = @('K1','K2') }
$Script:VisualHashCache.Clear()

Write-Output 'pack-change selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
