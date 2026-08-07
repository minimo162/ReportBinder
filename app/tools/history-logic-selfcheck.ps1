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
Save-AppConfig ([pscustomobject][ordered]@{ schemaVersion=2; lastSubmissionDir=$submissionDir; lastDataDir=$dataDir; lastOutputDir=$outputDir; lastMode='ja' })
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')
Assert-HistoryTest (Test-SourceRetentionEnabled) 'Source retention was not enabled for the history test workspace.'

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
$renderDir = Get-RenderRecordDir 'ja' $sourceId $cacheSnapshotId $cacheVersionId
New-Item -ItemType Directory -Path $renderDir -Force | Out-Null
$visualPath = Join-Path $renderDir 'visual-hashes.json'
Write-JsonFile $visualPath ([ordered]@{ schemaVersion=1; snapshotId=$cacheSnapshotId; versionId=$cacheVersionId; sheets=@([ordered]@{ sheetName='Summary'; visualHash='abc' }) })
$workspace = Get-WorkspacePath 'ja'
$contentDir = Get-ContentPdfVersionDir $workspace $sourceId $cacheVersionId
New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
$contentPath = Join-Path $contentDir ((Get-WorksheetStorageStem 'Summary') + '.pdf')
[IO.File]::WriteAllText($contentPath, 'test pdf placeholder', (New-Object Text.UTF8Encoding($false)))

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
