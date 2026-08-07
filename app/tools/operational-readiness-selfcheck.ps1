param()

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-operations-' + [Guid]::NewGuid().ToString('N'))
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_OPERATIONS_APPROOT', 'Process')
$oldTestRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_OPERATIONS_TESTROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', (Join-Path $testRoot 'local-config'), 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OPERATIONS_APPROOT', $appRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OPERATIONS_TESTROOT', $testRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_OPERATIONS_APPROOT')
    $testBody = @'
function Assert-OperationsTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Get-WorkspacePath([string]$Language, [string]$DataDir = '') {
    return (Join-Path $env:REPORTBINDER_OPERATIONS_TESTROOT $Language)
}

$workspace = Get-WorkspacePath 'ja'
New-Item -ItemType Directory -Path (Join-Path $workspace 'locks') -Force | Out-Null

# A shared-folder lock must fail closed while another process/instance owns it,
# then become available as soon as the owner releases the file handle.
$lockPath = Join-Path $workspace 'locks\structure.lock'
$first = Try-AcquireLockHandle $lockPath
Assert-OperationsTest ($null -ne $first) 'The first structure lock was not acquired.'
$second = Try-AcquireLockHandle $lockPath
Assert-OperationsTest ($null -eq $second) 'A concurrent structure lock was incorrectly acquired.'
Release-LockHandle $first
$third = Try-AcquireLockHandle $lockPath
Assert-OperationsTest ($null -ne $third) 'The structure lock was not released.'
Release-LockHandle $third

# Simulate a power loss after replacing an existing final PDF. Recovery must use
# hashes rather than an in-memory flag, restore the old bytes, and remove staging.
$outputDir = Join-Path $workspace 'output'
$backupDir = Get-FinalTransactionBackupDir 'ja' 'tx-rollback'
New-Item -ItemType Directory -Path $outputDir,$backupDir -Force | Out-Null
$finalPath = Join-Path $outputDir 'final.pdf'
$tempPath = Join-Path $outputDir '~building.pdf'
$backupPath = Join-Path $backupDir 'final.pdf'
[IO.File]::WriteAllBytes($backupPath, [Text.Encoding]::UTF8.GetBytes('old-final-pdf'))
Copy-Item -LiteralPath $backupPath -Destination $finalPath -Force
$oldHash = New-Sha256 $finalPath
[IO.File]::WriteAllBytes($finalPath, [Text.Encoding]::UTF8.GetBytes('new-final-pdf'))
$newHash = New-Sha256 $finalPath
[IO.File]::WriteAllBytes($tempPath, [Text.Encoding]::UTF8.GetBytes('unfinished'))
$journal = [pscustomobject]@{
    transactionId='tx-rollback'; phase='files-replaced'; category='ecm'
    targets=@([pscustomobject]@{ volume='ja-main'; finalPath=$finalPath; tempPath=$tempPath; backupPath=$backupPath; existed=$true; oldPdfHash=$oldHash; newPdfHash=$newHash })
}
Restore-FinalTransaction 'ja' $journal
Assert-OperationsTest ([string]$journal.phase -eq 'rolled-back') 'Interrupted final output was not rolled back.'
Assert-OperationsTest ((New-Sha256 $finalPath) -eq $oldHash) 'The prior final PDF was not restored.'
Assert-OperationsTest (-not (Test-Path -LiteralPath $tempPath)) 'The interrupted staging PDF remains.'
Assert-OperationsTest (-not (Test-Path -LiteralPath $backupDir)) 'The used transaction backup remains.'

# If the destination was edited externally after interruption, recovery must not
# overwrite it. It records a manual-recovery state for an administrator instead.
$manualBackupDir = Get-FinalTransactionBackupDir 'ja' 'tx-manual'
New-Item -ItemType Directory -Path $manualBackupDir -Force | Out-Null
$manualBackup = Join-Path $manualBackupDir 'final.pdf'
[IO.File]::WriteAllBytes($manualBackup, [Text.Encoding]::UTF8.GetBytes('known-old'))
$manualOldHash = New-Sha256 $manualBackup
[IO.File]::WriteAllBytes($finalPath, [Text.Encoding]::UTF8.GetBytes('external-edit'))
$externalHash = New-Sha256 $finalPath
$manual = [pscustomobject]@{
    transactionId='tx-manual'; phase='files-replaced'; category='ecm'
    targets=@([pscustomobject]@{ volume='ja-main'; finalPath=$finalPath; tempPath=''; backupPath=$manualBackup; existed=$true; oldPdfHash=$manualOldHash; newPdfHash=(Get-Sha256Text 'expected-new') })
}
Restore-FinalTransaction 'ja' $manual
Assert-OperationsTest ([string]$manual.phase -eq 'manual-recovery-required') 'Unknown destination state was not stopped for manual recovery.'
Assert-OperationsTest ((New-Sha256 $finalPath) -eq $externalHash) 'Manual-recovery destination was overwritten.'
$manualJournal = Read-JsonFile (Join-Path (Get-FinalTransactionDir 'ja') 'tx-manual.json') $null
Assert-OperationsTest ([string]$manualJournal.phase -eq 'manual-recovery-required') 'Manual recovery journal was not persisted.'

# Diagnostics must always return a useful result, even when Office COM cannot be
# started in a service/non-interactive Windows session.
$diagnostics = Get-SystemDiagnostics 'ja'
Assert-OperationsTest (@('ready','limited','blocked') -contains [string]$diagnostics.status) 'Diagnostics status is invalid.'
Assert-OperationsTest ($null -ne $diagnostics.office.excel -and $null -ne $diagnostics.office.word -and $null -ne $diagnostics.office.powerPoint) 'Office diagnostics are missing.'
Assert-OperationsTest ([bool]$diagnostics.runtime.pdfbox.ready) 'Bundled PDF engine was not detected.'
Assert-OperationsTest ([bool]$diagnostics.runtime.pdfjs.ready) 'Bundled PDF viewer was not detected.'

# Automation settings are validated and persisted, and retry metadata survives
# process restarts so failures can be retried without a tight loop.
$autoSettings = Update-AutoRenderSettings ([pscustomobject]@{ enabled=$false; quietPeriodSeconds=5; maxRetryCount=99; retryBaseSeconds=1; minFreeMegabytes=1; notifyOnCompletion=$true; notifyOnFailure=$true })
$normalizedAuto = Get-AutoRenderSettings
Assert-OperationsTest ([int]$normalizedAuto.quietPeriodSeconds -eq 10 -and [int]$normalizedAuto.maxRetryCount -eq 10 -and [int]$normalizedAuto.retryBaseSeconds -eq 10 -and [int]$normalizedAuto.minFreeMegabytes -eq 100) 'Automation settings were not clamped safely.'
$retryState = New-AutoState 'test-workbook'
Assert-OperationsTest ([int]$retryState.schemaVersion -eq 2 -and [int]$retryState.retryCount -eq 0 -and -not $retryState.notificationId) 'Automation retry state schema is incomplete.'
$retryState.state = 'failed'
$retryState.pendingHash = ('A' * 64)
Assert-OperationsTest (-not (Test-AutoFailureSuperseded $retryState ('A' * 64))) 'The failed source version would be retried in a tight loop.'
Assert-OperationsTest (Test-AutoFailureSuperseded $retryState ('B' * 64)) 'A newly saved source version does not recover from retry exhaustion.'

Write-Output ('operational readiness selfcheck ok (environment={0}, excel={1}, word={2}, powerpoint={3})' -f $diagnostics.status, $diagnostics.office.excel.status, $diagnostics.office.word.status, $diagnostics.office.powerPoint.status)
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OPERATIONS_APPROOT', $oldAppRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OPERATIONS_TESTROOT', $oldTestRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
