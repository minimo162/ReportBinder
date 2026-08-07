param(
    [string]$TestRoot = '',
    [switch]$KeepTestData
)

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = if ([string]::IsNullOrWhiteSpace($TestRoot)) { Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-custom-pack-workflow-' + [Guid]::NewGuid().ToString('N')) } else { [IO.Path]::GetFullPath($TestRoot) }
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_CUSTOM_PACK_TEST_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_CUSTOM_PACK_TEST_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_CUSTOM_PACK_TEST_APPROOT')
    $testBody = @'
function Assert-CustomPackTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
Assert-CustomPackTest ((Get-LegacyVolumeFromTargetId 'ja' 'unassigned') -eq 'none') 'Unassigned target was converted into a language-prefixed output.'
function New-CustomPackTestPdf([string]$Path, [int]$LineCount) {
    $textPath = [IO.Path]::ChangeExtension($Path, '.txt')
    $lines = @(1..$LineCount | ForEach-Object { "Cross-department report line $_ | Amount $_ | Status OK" })
    [IO.File]::WriteAllLines($textPath, $lines, (New-Object Text.UTF8Encoding($false)))
    $run = Invoke-NativeCapture (Resolve-JavaExe) @('-jar', (Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'), 'TextToPDF', '-standardFont', 'Helvetica', '-fontSize', '12', $Path, $textPath)
    if ([int]$run.exitCode -ne 0) { throw "Test PDF creation failed: $($run.text)" }
}

$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalConfigRoot 'data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
Save-AppConfig ([pscustomobject][ordered]@{ schemaVersion=2; lastSubmissionDir=$submissionDir; lastDataDir=$dataDir; lastOutputDir=$outputDir; lastMode='ja' })
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')

$overdueDueDate = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
$workflowTemplate = Save-PackTemplate 'ja' ([pscustomobject]@{
    displayName='Department review'; acceptedSourceTypes=@('excel','word','pdf')
    sourceRequirements=@([pscustomobject]@{requirementId='requirement_department_report';displayName='Monthly department report';ownerDepartment='Finance';required=$true;acceptedSourceTypes=@('pdf');defaultTargetId='main';dueDate=$overdueDueDate})
    targets=@([pscustomobject]@{targetId='main';displayName='Review packet';required=$true},[pscustomobject]@{targetId='appendix';displayName='Reference';required=$false},[pscustomobject]@{targetId='executive-summary';displayName='Executive summary';required=$false})
    rules=[pscustomobject]@{newItemDestination='main';retainManualOrder=$true;blockBuildWhenRequiredSourceIsStale=$true;blockBuildWhenRequiredSourceFailed=$true}
    output=[pscustomobject]@{fileNamePattern='{packName}_{targetName}.pdf';tableOfContents=$false}
})
$pack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='Monthly department packet'; templateId=[string]$workflowTemplate.templateId; settings=[pscustomobject]@{ documentTitle='Monthly department packet'; includeCover=$true } })
$packId = [string]$pack.packId
$unregisteredReady = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $packId
Assert-CustomPackTest (@($unregisteredReady.blockers | Where-Object { [string]$_.code -eq 'required-source-unregistered' -and [string]$_.requirementId -eq 'requirement_department_report' }).Count -eq 1) 'Unregistered required source was not reported.'
$unregisteredProgress = Get-PackProgressDashboard (Get-Structure 'ja') 'ja'
$unregisteredPackProgress = @($unregisteredProgress.packs | Where-Object { [string]$_.packId -eq $packId } | Select-Object -First 1)[0]
Assert-CustomPackTest ([string]$unregisteredPackProgress.state -eq 'overdue-source' -and [int]$unregisteredPackProgress.overdueRequiredSourceCount -eq 1 -and [int]$unregisteredProgress.overdueRequiredCount -ge 1 -and [string]$unregisteredPackProgress.nearestRequiredDueDate -eq $overdueDueDate) 'Overdue required source was not prioritized in cross-pack progress.'
$openOverduePack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='Overdue department pack'; templateId=[string]$workflowTemplate.templateId })
$openOverdueProgress = Get-PackProgressDashboard (Get-Structure 'ja') 'ja'
$openOverduePackProgress = @($openOverdueProgress.packs | Where-Object { [string]$_.packId -eq [string]$openOverduePack.packId } | Select-Object -First 1)[0]
Assert-CustomPackTest ([string]$openOverduePackProgress.state -eq 'overdue-source' -and [int]$openOverduePackProgress.overdueRequiredSourceCount -eq 1) 'Persistent overdue pack fixture was not created.'
$dueSoonDate = (Get-Date).AddDays(3).ToString('yyyy-MM-dd')
$dueSoonTemplate = Save-PackTemplate 'ja' ([pscustomobject]@{
    displayName='Due soon template'; acceptedSourceTypes=@('pdf')
    sourceRequirements=@([pscustomobject]@{requirementId='requirement_due_soon';displayName='Due soon report';ownerDepartment='Operations';required=$true;acceptedSourceTypes=@('pdf');defaultTargetId='main';dueDate=$dueSoonDate})
    targets=@([pscustomobject]@{targetId='main';displayName='Main';required=$true})
    rules=[pscustomobject]@{newItemDestination='main';retainManualOrder=$true;blockBuildWhenRequiredSourceIsStale=$true;blockBuildWhenRequiredSourceFailed=$true}
    output=[pscustomobject]@{fileNamePattern='{packName}_{targetName}.pdf';tableOfContents=$false}
})
$dueSoonPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='Due soon pack'; templateId=[string]$dueSoonTemplate.templateId })
$deadlineProgress = Get-PackProgressDashboard (Get-Structure 'ja') 'ja'
$dueSoonPackProgress = @($deadlineProgress.packs | Where-Object { [string]$_.packId -eq [string]$dueSoonPack.packId } | Select-Object -First 1)[0]
Assert-CustomPackTest ([string]$dueSoonPackProgress.state -eq 'missing-source' -and [int]$dueSoonPackProgress.dueSoonRequiredSourceCount -eq 1 -and [int]$deadlineProgress.dueSoonRequiredCount -ge 1 -and [string]$dueSoonPackProgress.nearestRequiredDueDate -eq $dueSoonDate) 'Due-soon required source was not summarized in cross-pack progress.'
$pdfPath = Join-Path $submissionDir 'department-report.pdf'
New-CustomPackTestPdf $pdfPath 180

$batch = Register-SourcesBatch 'ja' @('department-report.pdf') $packId 'pdf'
Assert-CustomPackTest ([int]$batch.registeredCount -eq 1 -and [int]$batch.errorCount -eq 0) 'Custom pack source registration failed.'
$sourceId = [string]$batch.registered[0].sourceId
[void](Update-SourceMetadata 'ja' $sourceId ([pscustomobject]@{ requirementId='requirement_department_report' }))
$registered = Get-Structure 'ja'
$workbook = @($registered.workbooks | Where-Object { [string]$_.workbookId -eq $sourceId } | Select-Object -First 1)[0]
Assert-CustomPackTest ([string]$workbook.packId -eq $packId) 'Custom pack assignment was lost after persistence.'
Assert-CustomPackTest ([string]$workbook.requirementId -eq 'requirement_department_report' -and [string]$workbook.ownerDepartment -eq 'Finance' -and [string]$workbook.defaultTargetId -eq 'main') 'Required source assignment defaults were not applied.'
$submittedProgress = Get-PackProgressDashboard (Get-Structure 'ja') 'ja'
$submittedPackProgress = @($submittedProgress.packs | Where-Object { [string]$_.packId -eq $packId } | Select-Object -First 1)[0]
Assert-CustomPackTest ([int]$submittedPackProgress.overdueRequiredSourceCount -eq 0 -and [int]$submittedPackProgress.submittedRequiredSourceCount -eq 1) 'Submitted required source remained overdue.'

$rendered = Render-Source 'ja' $sourceId
Assert-CustomPackTest (@($rendered.rendered).Count -ge 2) 'Custom pack PDF rendering did not create multiple pages.'
$afterRender = Get-Structure 'ja'
$pages = @($afterRender.pages | Where-Object { [string]$_.workbookId -eq $sourceId } | Sort-Object order)
Assert-CustomPackTest ($pages.Count -eq @($rendered.rendered).Count) 'Rendered custom pack pages are missing.'
Assert-CustomPackTest (@($pages | Where-Object { [string]$_.volume -ne 'ja-main' -or -not [bool]$_.enabled }).Count -eq 0) 'Template new-item destination was not applied.'
Assert-CustomPackTest ([string](Get-PackTargetDisplayName 'ja' $pack 'main') -eq 'Review packet') 'Template target display name was not applied.'
$items = @($afterRender.items | Where-Object { [string]$_.sourceId -eq $sourceId })
Assert-CustomPackTest ($items.Count -eq $pages.Count -and @($items | Where-Object { [string]$_.packId -ne $packId }).Count -eq 0) 'V3 items lost the custom pack assignment.'

$mainVolume = 'ja-main'
$reorderBody = [pscustomobject][ordered]@{ packId=$packId; volumes=[pscustomobject][ordered]@{ 'ja-main'=@($pages | ForEach-Object { Resolve-PageId $_ }); 'ja-appendix'=@(); none=@() } }
[void](Reorder-Pages 'ja' $reorderBody)
$mainIds = @($pages | ForEach-Object { Resolve-PageId $_ })
$appendixId = [string]$mainIds[0]
$changedLayout = [pscustomobject][ordered]@{ packId=$packId; volumes=[pscustomobject][ordered]@{ 'ja-main'=@($mainIds | Where-Object { $_ -ne $appendixId }); 'ja-appendix'=@($appendixId); none=@() } }
[void](Reorder-Pages 'ja' $changedLayout)
$layoutSnapshots = @(Get-LayoutSnapshots 'ja' $packId)
Assert-CustomPackTest ($layoutSnapshots.Count -ge 1 -and @($layoutSnapshots | Where-Object { [string]$_.packId -ne $packId }).Count -eq 0) 'Custom pack layout snapshots were not isolated by packId.'
Assert-CustomPackTest (@(Get-LayoutSnapshots 'ja' 'ecm').Count -eq 0) 'Custom pack layout history leaked into ECM.'
$restoreCandidate = $null
foreach ($layoutSnapshot in $layoutSnapshots) {
    $candidatePreview = Get-LayoutRestorePreview 'ja' $packId ([string]$layoutSnapshot.snapshotId)
    if (@($candidatePreview.volumeChanges).Count -eq 1) { $restoreCandidate = $layoutSnapshot; break }
}
Assert-CustomPackTest ($null -ne $restoreCandidate) 'Custom pack layout restore candidate was not found.'
$preview = Get-LayoutRestorePreview 'ja' $packId ([string]$restoreCandidate.snapshotId)
Assert-CustomPackTest ([string]$preview.packId -eq $packId -and [int]$preview.appliedPageCount -eq $pages.Count -and @($preview.volumeChanges).Count -eq 1) 'Custom pack layout restore preview is invalid.'
$restoredLayout = Restore-LayoutSnapshot 'ja' $packId ([string]$restoreCandidate.snapshotId)
Assert-CustomPackTest ([int]$restoredLayout.appliedPageCount -eq $pages.Count -and -not [string]::IsNullOrWhiteSpace([string]$restoredLayout.undoSnapshotId)) 'Custom pack layout restore failed.'
$afterRestore = Get-Structure 'ja'
$restoredPages = @($afterRestore.pages | Where-Object { [string]$_.workbookId -eq $sourceId })
Assert-CustomPackTest (@($restoredPages | Where-Object { [string]$_.volume -ne $mainVolume }).Count -eq 0) 'Custom pack layout did not return to the saved volume assignment.'
$afterRestoreSnapshots = @(Get-LayoutSnapshots 'ja' $packId)
Assert-CustomPackTest ($afterRestoreSnapshots.Count -eq ($layoutSnapshots.Count + 1) -and @($afterRestoreSnapshots | Where-Object { [string]$_.reason -eq 'pre-restore' }).Count -eq 1) 'Pre-restore layout snapshot was not retained.'
$ready = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' $mainVolume $packId
Assert-CustomPackTest ([bool]$ready.canBuild -and [int]$ready.pageCount -gt $pages.Count) 'Custom pack final readiness is not buildable.'
$originalHash = [string]$workbook.currentExcelHash
[void](Update-StructureLocked 'ja' { param($st) $w=@($st.workbooks|Where-Object{[string]$_.workbookId -eq $sourceId})[0]; Set-NoteProperty $w 'currentExcelHash' 'sha256:changed'; Set-NoteProperty $w 'status' 'source-updated' })
$staleReady = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' $mainVolume $packId
Assert-CustomPackTest (-not [bool]$staleReady.canBuild -and @($staleReady.blockers|Where-Object{[string]$_.code -eq 'required-source-stale'}).Count -eq 1) 'Required stale source did not block the build.'
[void](Update-StructureLocked 'ja' { param($st) $w=@($st.workbooks|Where-Object{[string]$_.workbookId -eq $sourceId})[0]; Set-NoteProperty $w 'currentExcelHash' $originalHash; Set-NoteProperty $w 'status' 'rendered'; Set-NoteProperty $w 'lastError' '' })
[void](Update-StructureLocked 'ja' { param($st) $w=@($st.workbooks|Where-Object{[string]$_.workbookId -eq $sourceId})[0]; Set-NoteProperty $w 'lastError' 'conversion failed'; Set-NoteProperty $w 'status' 'render-failed' })
$failedReady = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' $mainVolume $packId
Assert-CustomPackTest (-not [bool]$failedReady.canBuild -and @($failedReady.blockers|Where-Object{[string]$_.code -eq 'required-source-failed'}).Count -eq 1) 'Required failed source did not block the build.'
[void](Update-StructureLocked 'ja' { param($st) $w=@($st.workbooks|Where-Object{[string]$_.workbookId -eq $sourceId})[0]; Set-NoteProperty $w 'lastError' ''; Set-NoteProperty $w 'status' 'rendered' })
$movedPdfPath = $pdfPath + '.missing-test'
Move-Item -LiteralPath $pdfPath -Destination $movedPdfPath
try {
    $missingReady = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' $mainVolume $packId
    Assert-CustomPackTest (-not [bool]$missingReady.canBuild -and @($missingReady.blockers|Where-Object{[string]$_.code -eq 'required-source-missing'}).Count -eq 1) 'Missing required source did not block the build.'
} finally {
    Move-Item -LiteralPath $movedPdfPath -Destination $pdfPath
}
$ready = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' $mainVolume $packId
Assert-CustomPackTest ([bool]$ready.canBuild) 'Readiness did not recover after required source issues were resolved.'
$ecmReady = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' $mainVolume 'ecm'
Assert-CustomPackTest ([int]$ecmReady.pageCount -eq 0) 'Custom pack pages leaked into the ECM output.'

$built = Build-DocumentPackPdf 'ja' $packId 'main'
Assert-CustomPackTest (Test-Path -LiteralPath ([string]$built.outputPdf)) 'Custom pack final PDF was not created.'
Assert-CustomPackTest ((Get-PdfPageCount ([string]$built.outputPdf)) -eq [int]$ready.pageCount) 'Custom pack final PDF page count is incorrect.'
Assert-CustomPackTest (-not [string]::IsNullOrWhiteSpace([string]$built.archivePath) -and (Test-Path -LiteralPath ([string]$built.archivePath))) 'Custom pack final archive was not created.'
$archives = @(Get-FinalArchives 'ja' $packId)
Assert-CustomPackTest ($archives.Count -eq 1 -and [string]$archives[0].packId -eq $packId -and [string]$archives[0].targetId -eq 'main') 'Custom pack final archive metadata is invalid.'
$published = Publish-DocumentPackPdfToShared 'ja' $mainVolume $packId
Assert-CustomPackTest (Test-Path -LiteralPath ([string]$published.sharedPath)) 'Custom pack shared publication was not created.'
Assert-CustomPackTest ([string]$published.packId -eq $packId -and [string]$published.targetId -eq 'main') 'Custom pack shared publication metadata is invalid.'
Assert-CustomPackTest ((Normalize-FileHash (New-Sha256 ([string]$published.sharedPath))) -eq (Normalize-FileHash (New-Sha256 ([string]$built.outputPdf)))) 'Published custom pack PDF differs from the local output.'
$dynamicPageId = [string]$mainIds[0]
$dynamicLayout = [pscustomobject][ordered]@{ packId=$packId; volumes=[pscustomobject][ordered]@{ 'ja-main'=@($mainIds | Where-Object { $_ -ne $dynamicPageId }); 'ja-appendix'=@(); 'ja-executive-summary'=@($dynamicPageId); none=@() } }
[void](Reorder-Pages 'ja' $dynamicLayout)
$dynamicReady = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-executive-summary' $packId
Assert-CustomPackTest ([bool]$dynamicReady.canBuild -and [int]$dynamicReady.pageCount -gt 0 -and [string]$dynamicReady.targetId -eq 'executive-summary') 'Dynamic output target readiness is invalid.'
$dynamicBuilt = Build-DocumentPackPdf 'ja' $packId 'executive-summary'
Assert-CustomPackTest ((Test-Path -LiteralPath ([string]$dynamicBuilt.outputPdf)) -and [string]$dynamicBuilt.targetId -eq 'executive-summary' -and (Get-PdfPageCount ([string]$dynamicBuilt.outputPdf)) -eq [int]$dynamicReady.pageCount) 'Dynamic output target PDF was not created.'
$progress = Get-PackProgressDashboard (Get-Structure 'ja') 'ja'
$packProgress = @($progress.packs | Where-Object { [string]$_.packId -eq $packId } | Select-Object -First 1)[0]
Assert-CustomPackTest ($null -ne $packProgress -and @($packProgress.targets).Count -eq 3 -and [int]$packProgress.sourceCount -eq 1 -and [int]$progress.totalCount -ge 4) 'Cross-pack progress dashboard did not include the dynamic custom pack.'
$saved = Read-StructureUnlocked 'ja'
$outputState = Get-DataProperty $saved.outputs "$packId|main" $null
Assert-CustomPackTest ([string]$outputState.status -in @('built','needs-rebuild') -and -not [string]::IsNullOrWhiteSpace([string]$outputState.builtFingerprint)) 'Custom pack output state was not saved.'
Assert-CustomPackTest (@(Get-PublicPackList $saved 'ja' | Where-Object { [string]$_.packId -eq $packId -and [bool]$_.workflowAvailable }).Count -eq 1) 'Custom pack is not advertised as operational.'

Write-Output "custom-pack workflow selfcheck ok ($($pages.Count) source pages -> $([int]$ready.pageCount) output pages)"
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_CUSTOM_PACK_TEST_APPROOT', $oldAppRoot, 'Process')
    if (-not $KeepTestData -and (Test-Path -LiteralPath $testRoot)) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
