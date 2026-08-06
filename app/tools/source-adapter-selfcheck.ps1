param()

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-source-adapter-' + [Guid]::NewGuid().ToString('N'))
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_ADAPTER_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_ADAPTER_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_ADAPTER_APPROOT')
    $testBody = @'
function Assert-AdapterTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalConfigRoot 'data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$samplePath = Join-Path $submissionDir '01_ECM_adapter.xlsx'
[IO.File]::WriteAllBytes($samplePath, [byte[]](1,2,3,4,5))
Save-AppConfig ([pscustomobject][ordered]@{
    schemaVersion = 2; lastSubmissionDir = $submissionDir; lastDataDir = $dataDir; lastOutputDir = $outputDir; lastMode = 'ja'
})
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')

$adapter = Get-SourceAdapterDescriptor 'excel'
Assert-AdapterTest ([string]$adapter.adapterId -eq 'excel-com-v1') 'Excel adapterId is invalid.'
Assert-AdapterTest ([int]$adapter.contractVersion -eq 1) 'Adapter contractVersion is invalid.'
$wordAdapter = Get-SourceAdapterDescriptor 'word'
Assert-AdapterTest ([string]$wordAdapter.adapterId -eq 'word-com-v1' -and [string]$wordAdapter.unitKind -eq 'document') 'Word adapter descriptor is invalid.'
$unsupportedRejected = $false
try { [void](Get-SourceAdapterDescriptor 'powerpoint') } catch { $unsupportedRejected = $true }
Assert-AdapterTest $unsupportedRejected 'Unsupported source type was not rejected.'

$candidates = @(Get-SourceCandidates @('excel'))
Assert-AdapterTest ($candidates.Count -eq 1) 'Excel candidate count is invalid.'
Assert-AdapterTest ([string]$candidates[0].sourceType -eq 'excel' -and [string]$candidates[0].adapterId -eq 'excel-com-v1') 'Candidate adapter metadata is missing.'
$validated = Test-SourceCandidate ([pscustomobject]@{ relativePath='01_ECM_adapter.xlsx'; sourceType='excel' })
Assert-AdapterTest ([bool]$validated.ok) 'Excel candidate validation failed.'

$batch = Register-SourcesBatch 'ja' @('01_ECM_adapter.xlsx','01_ECM_adapter.xlsx') 'pack_ecm' 'excel'
Assert-AdapterTest ([int]$batch.registeredCount -eq 1 -and [int]$batch.errorCount -eq 0) 'Source batch registration failed or duplicate was not removed.'
$sourceId = [string]$batch.registered[0].sourceId
Assert-AdapterTest (-not [string]::IsNullOrWhiteSpace($sourceId)) 'Registered sourceId is missing.'
$structure = Get-Structure 'ja'
Assert-AdapterTest ([int]$structure.schemaVersion -eq 3) 'Registered structure is not schema v3.'
$source = @($structure.sources | Where-Object { [string]$_.sourceId -eq $sourceId } | Select-Object -First 1)
Assert-AdapterTest ($source.Count -eq 1) 'Registered schema v3 source is missing.'
Assert-AdapterTest ([string]$source[0].packId -eq 'pack_ecm' -and [string]$source[0].sourceType -eq 'excel' -and [string]$source[0].adapterId -eq 'excel-com-v1') 'Registered source adapter fields are invalid.'
$context = Get-RegisteredSourceAdapterContext 'ja' $sourceId
Assert-AdapterTest ([string]$context.adapter.adapterId -eq 'excel-com-v1') 'Registered source did not resolve through adapter dispatcher.'
$v4 = ConvertTo-V4StructureCompatibilityView $structure
Assert-AdapterTest (@($v4.workbooks).Count -eq 1 -and [string]$v4.workbooks[0].workbookId -eq $sourceId) 'V4 workbook compatibility was not preserved.'
$templates = @(Get-PackTemplateCatalog 'ja')
Assert-AdapterTest ($templates.Count -eq 3 -and [string]$templates[0].packId -eq 'pack_ecm') 'Built-in pack template catalog is invalid.'
$packs = @(Get-PublicPackList $structure 'ja')
Assert-AdapterTest ($packs.Count -ge 3 -and @($packs | Where-Object { [string]$_.category -eq 'ecm' -and [bool]$_.workflowAvailable }).Count -eq 1) 'Public pack catalog is invalid.'
$updated = Update-SourceMetadata 'ja' $sourceId ([pscustomobject]@{ ownerDepartment='経理部'; required=$false; defaultTargetId='appendix' })
Assert-AdapterTest ([string]$updated.ownerDepartment -eq '経理部' -and -not [bool]$updated.required -and [string]$updated.defaultTargetId -eq 'appendix') 'Source metadata update result is invalid.'
$updatedStructure = Get-Structure 'ja'
$updatedSource = @($updatedStructure.sources | Where-Object { [string]$_.sourceId -eq $sourceId } | Select-Object -First 1)
Assert-AdapterTest ($updatedSource.Count -eq 1 -and [string]$updatedSource[0].ownerDepartment -eq '経理部' -and -not [bool]$updatedSource[0].required) 'Source metadata was not persisted.'
$updatedV4 = ConvertTo-V4StructureCompatibilityView $updatedStructure
Assert-AdapterTest ([string]$updatedV4.workbooks[0].ownerDepartment -eq '経理部' -and -not [bool]$updatedV4.workbooks[0].required) 'Source metadata was not mirrored to V4 compatibility.'
$v2State = Get-V2StatePayload 'ja'
Assert-AdapterTest ([int]$v2State.apiVersion -eq 2 -and [int]$v2State.domainSchemaVersion -eq 3 -and @($v2State.packs).Count -ge 3) 'V2 state payload is invalid.'
Write-Output 'source-adapter selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_ADAPTER_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
