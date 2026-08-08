param(
    [string]$TestRoot = '',
    [switch]$KeepTestData
)

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = if ([string]::IsNullOrWhiteSpace($TestRoot)) { Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-history-logic-' + [Guid]::NewGuid().ToString('N')) } else { [IO.Path]::GetFullPath($TestRoot) }
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_HISTORY_TEST_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_HISTORY_TEST_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_HISTORY_TEST_APPROOT')
    $testBody = @'
function Assert-HistoryTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalProjectsRoot 'history-test-project\data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
Save-AppConfig ([pscustomobject][ordered]@{ schemaVersion=2; lastSubmissionDir=$submissionDir; lastDataDir=$dataDir; lastOutputDir=$outputDir })
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')
Assert-HistoryTest (Test-SourceRetentionEnabled) 'Source retention was not enabled for the history test workspace.'

# Local diff detail caches are immutable comparison identities. Re-rendering either
# snapshot, switching automatic/history scope, or bumping the algorithm must miss the
# old cache even when the two snapshot IDs themselves are unchanged.
$historyCacheV1 = Get-LocalDiffDetailCachePath 'ja' 'cache-key-source' 'base-snapshot' 'current-snapshot' 'base-render' 'current-render-1' 'history'
$historyCacheV2 = Get-LocalDiffDetailCachePath 'ja' 'cache-key-source' 'base-snapshot' 'current-snapshot' 'base-render' 'current-render-2' 'history'
$automaticCache = Get-LocalDiffDetailCachePath 'ja' 'cache-key-source' 'base-snapshot' 'current-snapshot' 'base-render' 'current-render-2' 'automatic'
Assert-HistoryTest ($historyCacheV1 -ne $historyCacheV2) 'A re-rendered snapshot reused the previous local diff detail cache key.'
Assert-HistoryTest ($historyCacheV2 -ne $automaticCache) 'Automatic and arbitrary-history comparisons shared a local diff detail cache key.'
$savedDiffAlgorithmVersion = $Script:DiffDetailAlgorithmVersion
try {
    $Script:DiffDetailAlgorithmVersion = $savedDiffAlgorithmVersion + 1
    $newAlgorithmCache = Get-LocalDiffDetailCachePath 'ja' 'cache-key-source' 'base-snapshot' 'current-snapshot' 'base-render' 'current-render-2' 'history'
    Assert-HistoryTest ($historyCacheV2 -ne $newAlgorithmCache) 'A diff algorithm bump reused the previous local diff detail cache key.'
} finally { $Script:DiffDetailAlgorithmVersion = $savedDiffAlgorithmVersion }

$reviewContext = [pscustomobject][ordered]@{
    available=$true; scope='history'; workbookId='cache-key-source'
    baselineSnapshotId='base-snapshot'; baselineVersionId='base-render'
    currentSnapshotId='current-snapshot'; currentVersionId='current-render-1'
}
$reviewPath = Get-DiffReviewStatePath 'ja' $reviewContext
New-Item -ItemType Directory -Path (Split-Path -Parent $reviewPath) -Force | Out-Null
Write-JsonFile $reviewPath ([ordered]@{
    schemaVersion=1; workbookId='cache-key-source'; scope='history'
    baselineSnapshotId='base-snapshot'; baselineVersionId='base-render'
    currentSnapshotId='current-snapshot'; currentVersionId='current-render-1'
    algorithmVersion=$Script:DiffDetailAlgorithmVersion; confirmedSheetKeys=@('s-review-a'); reviewedAt=(New-NowIso); reviewedBy='selfcheck'
})
$loadedReview = Get-DiffReviewStateForContext 'ja' $reviewContext
Assert-HistoryTest (@($loadedReview.confirmedSheetKeys) -contains 's-review-a') 'Exact-generation diff review state was not restored.'
$rerenderedReviewContext = $reviewContext.psobject.Copy()
$rerenderedReviewContext.currentVersionId = 'current-render-2'
Assert-HistoryTest ((Get-DiffReviewStatePath 'ja' $reviewContext) -ne (Get-DiffReviewStatePath 'ja' $rerenderedReviewContext)) 'Re-rendered comparison reused the previous review-state identity.'
Assert-HistoryTest (@((Get-DiffReviewStateForContext 'ja' $rerenderedReviewContext).confirmedSheetKeys).Count -eq 0) 'Review confirmation leaked into a re-rendered comparison.'

# Register one live source so Capture-RenderInput exercises the normal structure path.
$sourcePath = Join-Path $submissionDir 'history-source.pdf'
$sourceTextPath = Join-Path $submissionDir 'history-source.txt'
[IO.File]::WriteAllText($sourceTextPath, 'immutable history source', (New-Object Text.UTF8Encoding($false)))
$pdfRun = Invoke-NativeCapture (Resolve-JavaExe) @('-jar', (Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'), 'TextToPDF', '-standardFont', 'Helvetica', '-fontSize', '12', $sourcePath, $sourceTextPath)
Assert-HistoryTest ([int]$pdfRun.exitCode -eq 0 -and (Test-Path -LiteralPath $sourcePath)) ("History source PDF creation failed: " + [string]$pdfRun.text)
$batch = Register-SourcesBatch 'ja' @('history-source.pdf') 'ecm' 'pdf'
Assert-HistoryTest ([int]$batch.registeredCount -eq 1 -and [int]$batch.errorCount -eq 0) ("History source registration failed: " + ($batch | ConvertTo-Json -Depth 8 -Compress))
$sourceId = [string]$batch.registered[0].sourceId
$detected = Capture-DetectedSnapshot 'ja' $sourceId 'history-selfcheck'
Assert-HistoryTest ([bool]$detected.ok -and [bool]$detected.committed) ("History snapshot capture failed: " + ($detected | ConvertTo-Json -Depth 5 -Compress))
$registered = Get-Structure 'ja'
$workbook = @($registered.workbooks | Where-Object { [string]$_.workbookId -eq $sourceId } | Select-Object -First 1)[0]
$snapshotId = [string]$workbook.currentSnapshotId
$sourceHash = Normalize-FileHash ([string]$workbook.currentExcelHash)
Assert-HistoryTest (-not [string]::IsNullOrWhiteSpace($snapshotId) -and $null -ne (Get-SnapshotManifest 'ja' $sourceId $snapshotId)) 'Detected snapshot was not completed.'

# An existing retained source is reusable only when its bytes still match the manifest.
$retainedPath = [string](Get-SnapshotSourceState 'ja' $sourceId $snapshotId).sourcePath
[IO.File]::WriteAllText($retainedPath, 'corrupt retained source', (New-Object Text.UTF8Encoding($false)))
$repaired = Save-SnapshotSourceFile 'ja' $sourceId $snapshotId $sourcePath $sourceHash
Assert-HistoryTest ([bool]$repaired.ok -and -not [bool]$repaired.reused) 'A corrupt retained source was incorrectly reused.'
Assert-HistoryTest ((Normalize-FileHash (New-Sha256 $retainedPath)) -eq $sourceHash) 'Save-SnapshotSourceFile did not repair the retained source.'

# Capture must verify retained bytes and recover from an unchanged live source.
[IO.File]::WriteAllText($retainedPath, 'corrupt again', (New-Object Text.UTF8Encoding($false)))
$capture = Capture-RenderInput 'ja' $sourceId $snapshotId 'history-test-job'
Assert-HistoryTest ([bool]$capture.verified -and -not [bool]$capture.ephemeral -and [string]$capture.hash -eq $sourceHash) 'Capture-RenderInput did not recover the immutable source generation.'
Assert-HistoryTest ((Normalize-FileHash (New-Sha256 ([string]$capture.path))) -eq $sourceHash) 'Capture-RenderInput returned bytes that differ from the snapshot hash.'

# Only manifest-complete directories are generations. A crashed pending directory
# must not consume retention allowance or appear in history lists.
$pendingId = 'pending-generation'
$pendingDir = Get-SnapshotDir 'ja' $sourceId $pendingId
New-Item -ItemType Directory -Path $pendingDir -Force | Out-Null
Write-JsonFile (Join-Path $pendingDir 'manifest.pending.json') ([ordered]@{ schemaVersion=1; snapshotId=$pendingId; sourceHash='pending' })
Clear-SnapshotRuntimeCaches 'ja' $sourceId
$ids = @(Get-SnapshotIds 'ja' $sourceId)
Assert-HistoryTest ($ids -contains $snapshotId -and $ids -notcontains $pendingId) 'Pending snapshot directory was exposed as a completed generation.'

# Warm all three caches, remove the backing artifacts, then verify no cached value
# can resurrect the removed generation or content PDF.
$cacheSnapshotId = 'cache-generation'
$cacheVersionId = 'cache-version'
$cacheSnapshotDir = Get-SnapshotDir 'ja' $sourceId $cacheSnapshotId
New-Item -ItemType Directory -Path $cacheSnapshotDir -Force | Out-Null
$cacheManifestPath = Join-Path $cacheSnapshotDir 'manifest.json'
Write-JsonFile $cacheManifestPath ([ordered]@{ schemaVersion=1; snapshotId=$cacheSnapshotId; workbookId=$sourceId; relativePath='history-source.pdf'; sourceHash=$sourceHash; capturedAt=(New-NowIso); status='complete' })
$workspace = Get-WorkspacePath 'ja'
$contentDir = Get-ContentPdfVersionDir $workspace $sourceId $cacheVersionId
New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
$contentPath = Join-Path $contentDir ((Get-WorksheetStorageStem 'Summary') + '.pdf')
[IO.File]::WriteAllText($contentPath, 'test pdf placeholder', (New-Object Text.UTF8Encoding($false)))

# 履歴画面をPDF解析より先に開くと「比較不可」の要約が温まる。
# render record の保存後は履歴ルートの時刻が変わらなくても再集計されること。
Clear-SnapshotSummaryCache 'ja' $sourceId
$beforeRender = @(Get-SnapshotSummaries 'ja' $sourceId | Where-Object { [string]$_.snapshotId -eq $cacheSnapshotId } | Select-Object -First 1)
Assert-HistoryTest ($beforeRender.Count -eq 1 -and -not [bool]$beforeRender[0].visualCompareReady) 'Pre-render history summary did not reproduce the unavailable state.'
$Script:CurrentRenderEnvFingerprint = 'history-selfcheck-env'
Write-RenderRecord 'ja' $sourceId $cacheSnapshotId $cacheVersionId 'normal' $true ([pscustomobject][ordered]@{
    analyzerVersion=1; javaVersion='test'; javaVendor='test'
    sheets=@([ordered]@{ sheetName='Summary'; visualHash='abc'; status='ok' })
})
$afterRender = @(Get-SnapshotSummaries 'ja' $sourceId | Where-Object { [string]$_.snapshotId -eq $cacheSnapshotId } | Select-Object -First 1)
Assert-HistoryTest ($afterRender.Count -eq 1 -and [bool]$afterRender[0].visualCompareReady) 'Render record did not invalidate the stale unavailable history summary.'

# A cache written by a previous runtime may contain the pre-fix false result with
# the same history-directory timestamp. Restart must reject that schema and rebuild.
$summaryCacheKey = ('ja|' + $sourceId).ToLowerInvariant()
$legacySummaryPath = Get-LocalSnapshotSummaryCachePath 'ja' $sourceId
$legacyStamp = [IO.Directory]::GetLastWriteTimeUtc((Get-WorkbookHistoryDir 'ja' $sourceId)).Ticks
Write-JsonFile $legacySummaryPath ([ordered]@{
    schemaVersion=1; stamp=$legacyStamp
    summaries=@([ordered]@{ snapshotId=$cacheSnapshotId; visualCompareReady=$false; visualHashAvailable=$false; contentPdfAvailable=$false })
})
[void]$Script:SnapshotSummaryCache.Remove($summaryCacheKey)
$rebuilt = @(Get-SnapshotSummaries 'ja' $sourceId | Where-Object { [string]$_.snapshotId -eq $cacheSnapshotId } | Select-Object -First 1)
Assert-HistoryTest ($rebuilt.Count -eq 1 -and [bool]$rebuilt[0].visualCompareReady) 'Legacy unavailable summary cache was reused after restart.'

$renderDir = Get-RenderRecordDir 'ja' $sourceId $cacheSnapshotId $cacheVersionId
$visualPath = Join-Path $renderDir 'visual-hashes.json'

# A normal user-selected folder can push immutable history assets past MAX_PATH.
# Direct Test-Path fails there on Windows PowerShell 5.1, but history lookup must not.
$longWorkbookId = 'long-' + ('w' * 120)
$longSnapshotId = 'snapshot-' + ('s' * 40)
$longVersionId = 'version-' + ('v' * 40)
$longRenderPath = Join-Path (Get-RenderRecordDir 'ja' $longWorkbookId $longSnapshotId $longVersionId) 'visual-hashes.json'
Assert-HistoryTest ($longRenderPath.Length -gt 260) 'Long-path history fixture did not exceed MAX_PATH.'
Write-RenderRecord 'ja' $longWorkbookId $longSnapshotId $longVersionId 'normal' $true ([pscustomobject][ordered]@{
    analyzerVersion=1; javaVersion='test'; javaVendor='test'
    sheets=@([ordered]@{ sheetName='Summary'; visualHash='long-path-hash'; status='ok' })
})
Assert-HistoryTest (Test-FileExistsCompat $longRenderPath) 'Long-path render record was not written.'
Assert-HistoryTest ($null -ne (Get-VisualHashes 'ja' $longWorkbookId $longSnapshotId $longVersionId)) 'Long-path visual hashes were reported as missing.'

Assert-HistoryTest ($null -ne (Get-SnapshotManifest 'ja' $sourceId $cacheSnapshotId)) 'Snapshot manifest cache could not be warmed.'
Assert-HistoryTest ($null -ne (Get-VisualHashes 'ja' $sourceId $cacheSnapshotId $cacheVersionId)) 'Visual hash cache could not be warmed.'
$available = Get-HistoryRenderVersionAvailability 'ja' $sourceId $cacheSnapshotId $cacheVersionId
Assert-HistoryTest ([bool]$available.ready) 'Content PDF cache could not be warmed.'

Remove-Item -LiteralPath $cacheManifestPath -Force
Remove-Item -LiteralPath $visualPath -Force
Remove-Item -LiteralPath $contentDir -Recurse -Force
Assert-HistoryTest ($null -eq (Get-SnapshotManifest 'ja' $sourceId $cacheSnapshotId)) 'Deleted manifest survived through the runtime cache.'
Assert-HistoryTest ($null -eq (Get-VisualHashes 'ja' $sourceId $cacheSnapshotId $cacheVersionId)) 'Deleted visual hashes survived through the runtime cache.'
Assert-HistoryTest (@(Get-ContentPdfSheetIndex 'ja' $sourceId $cacheVersionId).Keys.Count -eq 0) 'Deleted content PDFs survived through the runtime cache.'
$unavailable = Get-HistoryRenderVersionAvailability 'ja' $sourceId $cacheSnapshotId $cacheVersionId
Assert-HistoryTest (-not [bool]$unavailable.ready) 'Removed comparison assets were still reported as ready.'

# Pins and leases may protect an existing artifact, but must never materialize a
# ghost artifact directory when the requested target no longer exists.
$missingSnapshotId = 'missing-snapshot'
$missingVersionId = 'missing-version'
Assert-HistoryTest (-not (New-SnapshotPin 'ja' $sourceId $missingSnapshotId 'manual' ([ordered]@{ at=New-NowIso }))) 'A pin was created for a missing snapshot.'
Assert-HistoryTest ([string]::IsNullOrWhiteSpace((New-SnapshotLease 'ja' $sourceId $missingSnapshotId 'compare' 'missing-job'))) 'A lease was created for a missing snapshot.'
Assert-HistoryTest (-not (Test-Path -LiteralPath (Get-SnapshotDir 'ja' $sourceId $missingSnapshotId))) 'Snapshot protection created a ghost generation directory.'
Assert-HistoryTest (-not (New-ContentPdfPin $workspace $sourceId $missingVersionId 'manual' ([ordered]@{ at=New-NowIso }))) 'A pin was created for a missing content PDF generation.'
Assert-HistoryTest ([string]::IsNullOrWhiteSpace((New-ContentPdfLease $workspace $sourceId $missingVersionId 'compare' 'missing-job'))) 'A lease was created for a missing content PDF generation.'
Assert-HistoryTest (-not (Test-Path -LiteralPath (Get-ContentPdfVersionDir $workspace $sourceId $missingVersionId))) 'Content PDF protection created a ghost generation directory.'

Write-Output 'history logic selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_HISTORY_TEST_APPROOT', $oldAppRoot, 'Process')
    if (-not $KeepTestData -and (Test-Path -LiteralPath $testRoot)) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
