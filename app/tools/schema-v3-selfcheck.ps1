param()

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$fixturePath = Join-Path $toolsRoot 'fixtures\structure-v2-generalized-ja.json'
$expectedPath = Join-Path $toolsRoot 'fixtures\structure-v3-migration-expected.json'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-schema-v3-' + [Guid]::NewGuid().ToString('N'))
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldFixture = [Environment]::GetEnvironmentVariable('REPORTBINDER_SCHEMA_FIXTURE', 'Process')
$oldExpected = [Environment]::GetEnvironmentVariable('REPORTBINDER_SCHEMA_EXPECTED', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_SCHEMA_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_SCHEMA_FIXTURE', $fixturePath, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_SCHEMA_EXPECTED', $expectedPath, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_SCHEMA_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_SCHEMA_APPROOT')
    $testBody = @'
function Assert-SchemaTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
$fixture = Read-JsonFile $env:REPORTBINDER_SCHEMA_FIXTURE $null
$expected = Read-JsonFile $env:REPORTBINDER_SCHEMA_EXPECTED $null
$converted = ConvertTo-StructureV3 $fixture 'ja' (Split-Path -Parent $env:REPORTBINDER_SCHEMA_FIXTURE)
Assert-SchemaTest ([int]$converted.schemaVersion -eq [int]$expected.schemaVersion) 'schemaVersion was not migrated to 3.'
$packIds = @($converted.packs | ForEach-Object { [string]$_.packId })
$actualPackKey = $packIds -join '|'
$expectedPackKey = @($expected.packIds) -join '|'
Assert-SchemaTest ($actualPackKey -eq $expectedPackKey) 'Built-in pack mapping is invalid.'
$sourceIds = @($converted.sources | ForEach-Object { [string]$_.sourceId })
$actualSourceKey = $sourceIds -join '|'
$expectedSourceKey = @($expected.sourceIds) -join '|'
Assert-SchemaTest ($actualSourceKey -eq $expectedSourceKey) 'workbookId was not preserved as sourceId.'
foreach ($itemId in @($expected.items.PSObject.Properties.Name)) {
    $actual = @($converted.items | Where-Object { [string]$_.itemId -eq $itemId } | Select-Object -First 1)
    $wanted = $expected.items.$itemId
    Assert-SchemaTest ($actual.Count -eq 1) "Item was not found: $itemId"
    foreach ($name in @('packId','targetId','order','enabled')) {
        Assert-SchemaTest ([string]$actual[0].$name -eq [string]$wanted.$name) "Item $name does not match: $itemId"
    }
}
foreach ($outputKey in @($expected.outputs.PSObject.Properties.Name)) {
    $actual = Get-DataProperty $converted.outputs $outputKey $null
    $wanted = $expected.outputs.$outputKey
    Assert-SchemaTest ($null -ne $actual) "Output was not found: $outputKey"
    Assert-SchemaTest ([string]$actual.status -eq [string]$wanted.status) "Output status does not match: $outputKey"
    Assert-SchemaTest ([string]$actual.builtFingerprint -eq [string]$wanted.builtFingerprint) "Output fingerprint does not match: $outputKey"
}
Assert-SchemaTest (@($converted.artifacts).Count -eq [int]$expected.artifactCount) 'Artifact count does not match.'
$issueCodes = @($converted.migrationIssues | ForEach-Object { [string]$_.code } | Sort-Object -Unique)
$actualIssueKey = $issueCodes -join '|'
$expectedIssueKey = @($expected.migrationIssueCodes) -join '|'
Assert-SchemaTest ($actualIssueKey -eq $expectedIssueKey) 'Migration issues do not match.'
$v4 = ConvertTo-V4StructureCompatibilityView $converted
Assert-SchemaTest ([int]$v4.schemaVersion -eq [int]$expected.v4SchemaVersion) 'V4 compatibility schemaVersion is invalid.'
Assert-SchemaTest (@($v4.workbooks).Count -eq [int]$expected.v4WorkbookCount) 'V4 compatibility workbook count is invalid.'
Assert-SchemaTest (@($v4.pages).Count -eq [int]$expected.v4PageCount) 'V4 compatibility page count is invalid.'
$firstIds = @($converted.items | ForEach-Object { [string]$_.itemId }) -join '|'
$convertedAgain = ConvertTo-StructureV3 $converted 'ja' (Split-Path -Parent $env:REPORTBINDER_SCHEMA_FIXTURE)
$secondIds = @($convertedAgain.items | ForEach-Object { [string]$_.itemId }) -join '|'
Assert-SchemaTest ($firstIds -eq $secondIds) 'itemId changed during idempotence check.'
$empty = New-EmptyStructure 'ja'
Assert-SchemaTest ([int]$empty.schemaVersion -eq 3 -and (Test-StructureDocument $empty 'ja')) 'New schema v3 structure is invalid.'
$convertedAgain.pages[0].volume = 'ja-appendix'
$convertedAgain.pages[0].order = 90
Sync-StructureV3FromLegacy $convertedAgain 'ja' | Out-Null
$syncedItem = @($convertedAgain.items | Where-Object { [string]$_.itemId -eq 'ecm-item-1' } | Select-Object -First 1)
Assert-SchemaTest ($syncedItem.Count -eq 1 -and [string]$syncedItem[0].targetId -eq 'appendix' -and [int]$syncedItem[0].order -eq 90) 'Legacy mutation was not synchronized to schema v3.'

function Test-FileMigration([string]$FixturePath, [string[]]$ExpectedBackups) {
    $caseId = [IO.Path]::GetFileNameWithoutExtension($FixturePath)
    $dataDir = Join-Path $Script:LocalConfigRoot ('migration-' + $caseId)
    $workspace = Join-Path $dataDir 'ja'
    New-Item -ItemType Directory -Path (Join-Path $workspace 'locks') -Force | Out-Null
    Copy-Item -LiteralPath $FixturePath -Destination (Join-Path $workspace 'structure.json') -Force
    $migration = Initialize-Or-MigrateStructure 'ja' $dataDir
    Assert-SchemaTest ([bool]$migration.migrated) "Migration did not run: $caseId"
    $saved = Read-JsonFile (Join-Path $workspace 'structure.json') $null
    Assert-SchemaTest ([int]$saved.schemaVersion -eq 3) "Saved schemaVersion is not 3: $caseId"
    Assert-SchemaTest (Test-StructureDocument $saved 'ja') "Saved structure is invalid: $caseId"
    foreach ($backupName in $ExpectedBackups) {
        Assert-SchemaTest (Test-Path -LiteralPath (Join-Path $workspace $backupName)) "Migration backup is missing: $backupName"
    }
    $marker = Read-JsonFile (Join-Path $workspace 'schema-version.json') $null
    Assert-SchemaTest ([int]$marker.schemaVersion -eq 3) "Migration marker is invalid: $caseId"
    $second = Initialize-Or-MigrateStructure 'ja' $dataDir
    Assert-SchemaTest (-not [bool]$second.migrated) "Migration is not idempotent: $caseId"
}

$fixtureDir = Split-Path -Parent $env:REPORTBINDER_SCHEMA_FIXTURE
Test-FileMigration $env:REPORTBINDER_SCHEMA_FIXTURE @('structure.json.v2.bak')
Test-FileMigration (Join-Path $fixtureDir 'structure-v1-ja.json') @('structure.json.v1.bak','structure.json.v2.bak')
Write-Output 'schema-v3 selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_SCHEMA_FIXTURE', $oldFixture, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_SCHEMA_EXPECTED', $oldExpected, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_SCHEMA_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
