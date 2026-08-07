param()

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-pack-lifecycle-' + [Guid]::NewGuid().ToString('N'))
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_PACK_TEST_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_PACK_TEST_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_PACK_TEST_APPROOT')
    $testBody = @'
function Assert-PackTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalConfigRoot 'data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
Save-AppConfig ([pscustomobject][ordered]@{
    schemaVersion = 2; lastSubmissionDir = $submissionDir; lastDataDir = $dataDir; lastOutputDir = $outputDir; lastMode = 'en'
})
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('en')

$cancelJobId = 'job_20260807_120000_abcdef12'
$cancelJobDir = Get-RenderJobDir 'en'
$cancelStatusPath = Join-Path $cancelJobDir "$cancelJobId.status.json"
$cancelInputPath = Join-Path $cancelJobDir "$cancelJobId.input.json"
Write-JsonFile $cancelStatusPath ([ordered]@{ ok=$true; jobId=$cancelJobId; status='queued'; total=0; completed=0; failed=0; percent=1; message='queued'; processId=0; results=@(); errors=@(); updatedAt=New-NowIso })
Write-JsonFile $cancelInputPath ([ordered]@{ jobId=$cancelJobId; mode='en'; workbookIds=@(); onlyUpdated=$false; category=''; statusPath=$cancelStatusPath; stdoutPath=(Join-Path $cancelJobDir "$cancelJobId.out.log"); stderrPath=(Join-Path $cancelJobDir "$cancelJobId.err.log") })
$cancelRequest = Request-RenderJobCancellation 'en' $cancelJobId
Assert-PackTest ([bool]$cancelRequest.accepted -and (Test-RenderJobCancellationRequested 'en' $cancelJobId)) 'Render job cancellation request was not persisted.'
$cancelPending = Read-RenderJobStatus 'en' $cancelJobId
Assert-PackTest ([bool]$cancelPending.cancelRequested -and [string]$cancelPending.status -eq 'queued') 'Pending render cancellation was not exposed in job status.'
Invoke-RenderJobFromFile $cancelInputPath
$cancelledJob = Read-RenderJobStatus 'en' $cancelJobId
Assert-PackTest ([string]$cancelledJob.status -eq 'cancelled' -and -not [string]::IsNullOrWhiteSpace([string]$cancelledJob.cancelledAt)) 'Render job did not stop at a safe cancellation boundary.'
$invalidCancelRejected = $false
try { [void](Request-RenderJobCancellation 'en' '..\bad') } catch { $invalidCancelRejected = $true }
Assert-PackTest $invalidCancelRejected 'Unsafe render job cancellation ID was accepted.'

$templates = @(Get-PackTemplateCatalog 'en')
$generic = @($templates | Where-Object { [string]$_.templateId -eq 'builtin-generic-department-pack' })
Assert-PackTest ($templates.Count -eq 4 -and $generic.Count -eq 1) 'Generic pack template is missing.'
Assert-PackTest ((@($generic[0].acceptedSourceTypes) -join '|') -eq 'excel|word|pdf|powerpoint') 'Generic template source types are invalid.'

$userTemplate = Save-PackTemplate 'en' ([pscustomobject]@{
    displayName = 'Finance Review'
    description = 'Finance source review template'
    acceptedSourceTypes = @('excel','pdf')
    sourceRequirements = @([pscustomobject]@{ requirementId='requirement_finance'; displayName='Finance report'; ownerDepartment='Finance'; required=$true; acceptedSourceTypes=@('excel','pdf'); defaultTargetId='main'; dueDate='2026-08-31' })
    targets = @(
        [pscustomobject]@{ targetId='main'; displayName='Review'; required=$true },
        [pscustomobject]@{ targetId='appendix'; displayName='Evidence'; required=$false }
    )
    rules = [pscustomobject]@{ newItemDestination='main'; retainManualOrder=$true; blockBuildWhenRequiredSourceIsStale=$true; blockBuildWhenRequiredSourceFailed=$true }
    output = [pscustomobject]@{ fileNamePattern='{packName}_{targetName}_{yyyyMMdd}.pdf'; tableOfContents=$true }
})
Assert-PackTest ([string]$userTemplate.templateId -match '^template_[0-9a-f]{16}$') 'User template ID is invalid.'
Assert-PackTest ((@(Get-PackTemplateCatalog 'en')).Count -eq 5) 'User template was not added to the catalog.'
$updatedTemplate = Save-PackTemplate 'en' ([pscustomobject]@{
    targets = @(
        [pscustomobject]@{ targetId='main'; displayName='Primary review'; required=$true },
        [pscustomobject]@{ targetId='appendix'; displayName='Evidence'; required=$false }
    )
}) ([string]$userTemplate.templateId)
Assert-PackTest ((Get-IntDataProperty $updatedTemplate 'templateVersion' 0) -eq 2) 'Template version was not incremented.'

$templatePack = New-DocumentPack 'en' ([pscustomobject]@{ displayName='Finance Review Pack'; templateId=[string]$userTemplate.templateId })
Assert-PackTest ([string]$templatePack.templateConfig.targets[0].displayName -eq 'Primary review') 'Pack did not store a template snapshot.'
Assert-PackTest ([string]$templatePack.templateConfig.sourceRequirements[0].requirementId -eq 'requirement_finance') 'Pack did not store source requirements in the template snapshot.'
Assert-PackTest ([bool]$templatePack.settings.includeToc -and [string]$templatePack.settings.outputFileNamePattern -eq '{packName}_{targetName}_{yyyyMMdd}.pdf') 'Template output defaults were not applied.'
[void](Save-PackTemplate 'en' ([pscustomobject]@{
    targets = @(
        [pscustomobject]@{ targetId='main'; displayName='Changed later'; required=$true },
        [pscustomobject]@{ targetId='appendix'; displayName='Evidence'; required=$false }
    )
}) ([string]$userTemplate.templateId))
$templatePackPublic = @(Get-PublicPackList (Get-Structure 'en') 'en' | Where-Object { [string]$_.packId -eq [string]$templatePack.packId })
Assert-PackTest ([string]$templatePackPublic[0].targets[0].displayName -eq 'Primary review') 'Existing pack changed when its template was edited.'
Assert-PackTest ([bool]$templatePackPublic[0].templateUpdateAvailable -and [int]$templatePackPublic[0].availableTemplateVersion -eq 3) 'Template update availability was not advertised.'
$upgradePreview = Get-PackTemplateUpgradePreview 'en' ([string]$templatePack.packId)
Assert-PackTest ([bool]$upgradePreview.updateAvailable -and [int]$upgradePreview.fromVersion -eq 2 -and [int]$upgradePreview.toVersion -eq 3) 'Template upgrade preview is invalid.'
$upgradedPack = Update-PackTemplateSnapshot 'en' ([string]$templatePack.packId)
Assert-PackTest ([int]$upgradedPack.toVersion -eq 3 -and [string]$upgradedPack.templateConfig.targets[0].displayName -eq 'Changed later') 'Template upgrade was not applied.'
$afterUpgradePublic = @(Get-PublicPackList (Get-Structure 'en') 'en' | Where-Object { [string]$_.packId -eq [string]$templatePack.packId })
Assert-PackTest (-not [bool]$afterUpgradePublic[0].templateUpdateAvailable -and [string]$afterUpgradePublic[0].targets[0].displayName -eq 'Changed later') 'Upgraded pack metadata is invalid.'
$usedTemplateDeleteRejected = $false
try { [void](Remove-PackTemplate 'en' ([string]$userTemplate.templateId)) } catch { $usedTemplateDeleteRejected = $true }
Assert-PackTest $usedTemplateDeleteRejected 'Template used by a pack was deleted.'
$disposableTemplate = Save-PackTemplate 'en' ([pscustomobject]@{ displayName='Disposable'; acceptedSourceTypes=@('pdf') })
[void](Remove-PackTemplate 'en' ([string]$disposableTemplate.templateId))
Assert-PackTest (@(Get-PackTemplateCatalog 'en' | Where-Object { [string]$_.templateId -eq [string]$disposableTemplate.templateId }).Count -eq 0) 'Unused template was not deleted.'

$created = New-DocumentPack 'en' ([pscustomobject]@{
    displayName = 'Monthly Department Report'
    settings = [pscustomobject]@{ includeCover=$true; documentTitle='Monthly Department Report'; outputFileNamePattern='{packName}_{yyyyMMdd}.pdf' }
})
Assert-PackTest ([string]$created.packId -match '^pack_[0-9a-f]{16}$') 'Custom packId is invalid.'
Assert-PackTest ([string]$created.templateId -eq 'builtin-generic-department-pack') 'Default template was not assigned.'
Assert-PackTest ([bool]$created.settings.includeCover -and [string]$created.settings.documentTitle -eq 'Monthly Department Report') 'Create settings were not applied.'

$structure = Get-Structure 'en'
$createdPublic = @(Get-PublicPackList $structure 'en' | Where-Object { [string]$_.packId -eq [string]$created.packId })
Assert-PackTest ($createdPublic.Count -eq 1 -and -not [bool]$createdPublic[0].builtIn -and [bool]$createdPublic[0].archivable) 'Created pack public metadata is invalid.'
foreach ($targetId in @('main','appendix')) {
    $output = Get-DataProperty $structure.outputs ("$($created.packId)|$targetId") $null
    Assert-PackTest ($null -ne $output -and [string]$output.status -eq 'not-built') "Output state was not initialized: $targetId"
}

$renamed = Update-PackSettings 'en' ([string]$created.packId) ([pscustomobject]@{
    displayName='Quarterly Department Report'
    settings=[pscustomobject]@{ documentSubtitle='FY2026 Q1' }
})
Assert-PackTest ([string]$renamed.displayName -eq 'Quarterly Department Report') 'Pack rename failed.'
Assert-PackTest ([string]$renamed.settings.documentSubtitle -eq 'FY2026 Q1') 'Pack settings update failed.'

$copy = Copy-DocumentPack 'en' ([string]$created.packId) ([pscustomobject]@{})
Assert-PackTest ([string]$copy.packId -ne [string]$created.packId) 'Duplicated pack reused the source ID.'
Assert-PackTest ([string]$copy.displayName -eq 'Quarterly Department Report (Copy)') 'Duplicated pack name is invalid.'
Assert-PackTest ([string]$copy.duplicatedFromPackId -eq [string]$created.packId) 'Duplicate origin was not recorded.'
Assert-PackTest ([string]$copy.settings.documentSubtitle -eq 'FY2026 Q1') 'Duplicated settings were not copied.'

$copy2 = Copy-DocumentPack 'en' ([string]$created.packId) ([pscustomobject]@{})
Assert-PackTest ([string]$copy2.displayName -eq 'Quarterly Department Report (Copy 2)') 'Duplicate name collision was not resolved.'
$longPack = New-DocumentPack 'en' ([pscustomobject]@{ displayName=('X' * 120) })
$longCopy = Copy-DocumentPack 'en' ([string]$longPack.packId) ([pscustomobject]@{})
Assert-PackTest (([string]$longCopy.displayName).Length -le 120 -and [string]$longCopy.displayName -match '\(Copy\)$') 'Generated duplicate name exceeded the length limit.'

[void](Set-DocumentPackArchived 'en' ([string]$copy.packId) $true)
$active = @(Get-PublicPackList (Get-Structure 'en') 'en')
$all = @(Get-PublicPackList (Get-Structure 'en') 'en' $true)
Assert-PackTest (@($active | Where-Object { [string]$_.packId -eq [string]$copy.packId }).Count -eq 0) 'Archived pack remained in the active list.'
$archived = @($all | Where-Object { [string]$_.packId -eq [string]$copy.packId })
Assert-PackTest ($archived.Count -eq 1 -and [bool]$archived[0].archived -and -not [string]::IsNullOrWhiteSpace([string]$archived[0].archivedAt)) 'Archived pack metadata is invalid.'

[void](Set-DocumentPackArchived 'en' ([string]$copy.packId) $false)
$restored = @(Get-PublicPackList (Get-Structure 'en') 'en' | Where-Object { [string]$_.packId -eq [string]$copy.packId })
Assert-PackTest ($restored.Count -eq 1 -and -not [bool]$restored[0].archived) 'Pack restore failed.'

$duplicateNameRejected = $false
try { [void](New-DocumentPack 'en' ([pscustomobject]@{ displayName='Quarterly Department Report' })) } catch { $duplicateNameRejected = $true }
Assert-PackTest $duplicateNameRejected 'Duplicate active display name was accepted.'
$unknownTemplateRejected = $false
try { [void](New-DocumentPack 'en' ([pscustomobject]@{ displayName='Invalid Template'; templateId='missing-template' })) } catch { $unknownTemplateRejected = $true }
Assert-PackTest $unknownTemplateRejected 'Unknown template was accepted.'
$builtinArchiveRejected = $false
try { [void](Set-DocumentPackArchived 'en' 'pack_ecm' $true) } catch { $builtinArchiveRejected = $true }
Assert-PackTest $builtinArchiveRejected 'Built-in pack was archived.'

$reloaded = Read-StructureUnlocked 'en'
Assert-PackTest (Test-StructureDocument $reloaded 'en') 'Persisted structure is invalid.'
Assert-PackTest (@($reloaded.packs | Where-Object { [string]$_.packId -eq [string]$created.packId }).Count -eq 1) 'Custom pack did not survive persistence.'
Write-Output 'pack-lifecycle selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_PACK_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
