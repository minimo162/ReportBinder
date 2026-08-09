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
function Write-CustomPackStage([string]$Message) { Write-Output ("[custom-pack {0}] {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $Message) }
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
Save-AppConfig ([pscustomobject][ordered]@{ schemaVersion=2; lastSubmissionDir=$submissionDir; lastDataDir=$dataDir; lastOutputDir=$outputDir })
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')

$overdueDueDate = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
$workflowTemplate = Save-PackTemplate 'ja' ([pscustomobject]@{
    displayName='Department review'; acceptedSourceTypes=@('excel','word','pdf')
    sourceRequirements=@([pscustomobject]@{requirementId='requirement_department_report';displayName='Monthly department report';ownerDepartment='Finance';required=$true;acceptedSourceTypes=@('pdf');defaultTargetId='main';dueDate=$overdueDueDate})
    targets=@([pscustomobject]@{targetId='main';displayName='Review packet';required=$true},[pscustomobject]@{targetId='appendix';displayName='Reference';required=$false},[pscustomobject]@{targetId='executive-summary';displayName='Executive summary';required=$false})
    # A source explicitly marked required must block stale/failed output even
    # when an imported legacy template carried permissive rule flags.
    rules=[pscustomobject]@{newItemDestination='main';retainManualOrder=$true;blockBuildWhenRequiredSourceIsStale=$false;blockBuildWhenRequiredSourceFailed=$false}
    output=[pscustomobject]@{fileNamePattern='{packName}_{targetName}.pdf'}
})
$pack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='Monthly department packet'; templateId=[string]$workflowTemplate.templateId; settings=[pscustomobject]@{ documentTitle='Monthly department packet' } })
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
    output=[pscustomobject]@{fileNamePattern='{packName}_{targetName}.pdf'}
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
Assert-CustomPackTest ([bool]$ready.canBuild -and [int]$ready.pageCount -eq $pages.Count) 'Custom pack final readiness must contain only registered source pages.'
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

Write-CustomPackStage 'building final PDF'
$built = Build-DocumentPackPdf 'ja' $packId 'main'
Assert-CustomPackTest (Test-Path -LiteralPath ([string]$built.outputPdf)) 'Custom pack final PDF was not created.'
Assert-CustomPackTest ((Get-PdfPageCount ([string]$built.outputPdf)) -eq [int]$ready.pageCount) 'Custom pack final PDF page count is incorrect.'
if (-not [string]::IsNullOrWhiteSpace([string]$built.archivePath)) {
    Assert-CustomPackTest (Test-Path -LiteralPath ([string]$built.archivePath)) 'Returned custom pack archive path does not exist.'
}
Write-CustomPackStage 'checking draft review state'
$reviewDraft = Get-PackReviewSnapshot (Get-Structure 'ja') 'ja' $packId $true
Assert-CustomPackTest ([string]$reviewDraft.status -eq 'draft' -and [bool]$reviewDraft.canSubmit) 'Built custom pack was not ready for review submission.'
Write-CustomPackStage 'submitting review'
$reviewSubmitted = Invoke-PackReviewAction 'ja' $packId ([pscustomobject]@{ action='submit'; note='ready for review'; actor='workflow-owner' })
Assert-CustomPackTest ([string]$reviewSubmitted.status -eq 'in-review' -and @($reviewSubmitted.events).Count -eq 1) 'Custom pack review submission was not recorded.'
Write-CustomPackStage 'approving review'
$reviewApproved = Invoke-PackReviewAction 'ja' $packId ([pscustomobject]@{ action='approve'; note='approved'; actor='reviewer' })
Assert-CustomPackTest ([string]$reviewApproved.status -eq 'approved' -and [string]$reviewApproved.approvedBy -eq 'reviewer' -and @($reviewApproved.events).Count -eq 2) 'Custom pack approval was not recorded.'
$approvedProgress = @((Get-PackProgressDashboard (Get-Structure 'ja') 'ja').packs | Where-Object { [string]$_.packId -eq $packId })[0]
Assert-CustomPackTest ([string]$approvedProgress.state -eq 'complete' -and [string]$approvedProgress.reviewStatus -eq 'approved') 'Approved custom pack was not complete in the cross-pack dashboard.'
$reviewPage = @((Get-Structure 'ja').pages | Where-Object { [string]$_.workbookId -eq $sourceId } | Select-Object -First 1)[0]
$reviewPageId = Resolve-PageId $reviewPage
$reviewOriginalTitle = [string]$reviewPage.title
[void](Update-StructureLocked 'ja' { param($st) $p=@($st.pages|Where-Object{(Resolve-PageId $_)-eq $reviewPageId})[0]; Set-NoteProperty $p 'title' ($reviewOriginalTitle + ' updated') })
Write-CustomPackStage 'checking stale review state'
$reviewStale = Get-PackReviewSnapshot (Get-Structure 'ja') 'ja' $packId $false
Assert-CustomPackTest ([string]$reviewStale.status -eq 'stale') 'Approved review did not become stale after the reviewed fingerprint changed.'
$staleApprovalError = ''
try { [void](Invoke-PackReviewAction 'ja' $packId ([pscustomobject]@{ action='approve'; note='must fail'; actor='reviewer' })) } catch { $staleApprovalError = $_.Exception.Message }
Assert-CustomPackTest (-not [string]::IsNullOrWhiteSpace($staleApprovalError)) 'A stale review was incorrectly approved.'
[void](Update-StructureLocked 'ja' { param($st) $p=@($st.pages|Where-Object{(Resolve-PageId $_)-eq $reviewPageId})[0]; Set-NoteProperty $p 'title' $reviewOriginalTitle })
$missingReasonError = ''
try { [void](Invoke-PackReviewAction 'ja' $packId ([pscustomobject]@{ action='request-changes'; note=''; actor='reviewer' })) } catch { $missingReasonError = $_.Exception.Message }
Assert-CustomPackTest (-not [string]::IsNullOrWhiteSpace($missingReasonError)) 'A review change request without a reason was accepted.'
Write-CustomPackStage 'requesting changes'
$reviewChanges = Invoke-PackReviewAction 'ja' $packId ([pscustomobject]@{ action='request-changes'; note='clarify the appendix'; actor='reviewer' })
Assert-CustomPackTest ([string]$reviewChanges.status -eq 'changes-requested' -and @($reviewChanges.events).Count -eq 3) 'Review change request was not recorded.'
$changesProgress = @((Get-PackProgressDashboard (Get-Structure 'ja') 'ja').packs | Where-Object { [string]$_.packId -eq $packId })[0]
Assert-CustomPackTest ([string]$changesProgress.state -eq 'complete' -and [string]$changesProgress.reviewStatus -eq 'changes-requested') 'Cross-pack dashboard must report progress and the confirmation record independently.'
Write-CustomPackStage 'resubmitting review'
$reviewResubmitted = Invoke-PackReviewAction 'ja' $packId ([pscustomobject]@{ action='submit'; note='updated response'; actor='workflow-owner' })
Assert-CustomPackTest ([string]$reviewResubmitted.status -eq 'in-review') 'Corrected custom pack could not be resubmitted.'
Write-CustomPackStage 'reopening review'
$reviewReopened = Invoke-PackReviewAction 'ja' $packId ([pscustomobject]@{ action='reopen'; note='continue editing'; actor='workflow-owner' })
Assert-CustomPackTest ([string]$reviewReopened.status -eq 'draft' -and @($reviewReopened.events).Count -eq 5) 'Review audit trail did not preserve the full workflow.'
$archives = @(Get-FinalArchives 'ja' $packId)
Assert-CustomPackTest ($archives.Count -le 1) 'Custom pack returned duplicate final archives.'
if ($archives.Count -eq 1) {
    Assert-CustomPackTest ([string]$archives[0].packId -eq $packId -and [string]$archives[0].targetId -eq 'main') 'Custom pack final archive metadata is invalid.'
}
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
Assert-CustomPackTest ($null -ne $packProgress -and @($packProgress.targets).Count -eq 3 -and [int]$packProgress.sourceCount -eq 1 -and [int]$progress.totalCount -ge 1) 'Cross-pack progress dashboard did not include the dynamic custom pack.'
$saved = Read-StructureUnlocked 'ja'
$outputState = Get-DataProperty $saved.outputs "$packId|main" $null
Assert-CustomPackTest ([string]$outputState.status -in @('built','needs-rebuild') -and -not [string]::IsNullOrWhiteSpace([string]$outputState.builtFingerprint)) 'Custom pack output state was not saved.'
Assert-CustomPackTest (@(Get-PublicPackList $saved 'ja' | Where-Object { [string]$_.packId -eq $packId -and [bool]$_.workflowAvailable }).Count -eq 1) 'Custom pack is not advertised as operational.'

# A saved layout may outlive a custom output target. Restoring it must not revive
# the removed target; the affected page is safely returned to the unassigned tray.
$obsoleteTargetSnapshotId = Save-LayoutSnapshot 'ja' $packId 'obsolete-target-test'
Assert-CustomPackTest (-not [string]::IsNullOrWhiteSpace($obsoleteTargetSnapshotId)) 'Obsolete-target layout snapshot was not saved.'
[void](Update-StructureLocked 'ja' {
    param($st)
    $targetPack = @($st.packs | Where-Object { [string]$_.packId -eq $packId } | Select-Object -First 1)[0]
    $targets = @(Get-Array (Get-DataProperty $targetPack.templateConfig 'targets' @()) | Where-Object { [string]$_.targetId -ne 'executive-summary' })
    Set-NoteProperty $targetPack.templateConfig 'targets' @($targets)
    $targetPage = @($st.pages | Where-Object { (Resolve-PageId $_) -eq $dynamicPageId } | Select-Object -First 1)[0]
    Set-NoteProperty $targetPage 'volume' 'ja-main'
    Set-NoteProperty $targetPage 'enabled' $true
    return $true
})
$obsoletePreview = Get-LayoutRestorePreview 'ja' $packId $obsoleteTargetSnapshotId
Assert-CustomPackTest (@($obsoletePreview.invalidVolumePageIds).Count -eq 1 -and [string]$obsoletePreview.invalidVolumePageIds[0] -eq $dynamicPageId) 'Removed output target was not identified in restore preview.'
$obsoleteChange = @($obsoletePreview.volumeChanges | Where-Object { [string]$_.pageId -eq $dynamicPageId } | Select-Object -First 1)[0]
Assert-CustomPackTest ([string]$obsoleteChange.to -eq 'none' -and [string]$obsoleteChange.storedVolume -eq 'ja-executive-summary' -and [bool]$obsoleteChange.normalized) 'Removed output target was not normalized in restore preview.'
$obsoleteRestore = Restore-LayoutSnapshot 'ja' $packId $obsoleteTargetSnapshotId
$obsoletePage = @((Get-Structure 'ja').pages | Where-Object { (Resolve-PageId $_) -eq $dynamicPageId } | Select-Object -First 1)[0]
Assert-CustomPackTest ([int]$obsoleteRestore.invalidVolumePageCount -eq 1 -and -not [string]::IsNullOrWhiteSpace([string]$obsoleteRestore.undoSnapshotId)) 'Removed output target restore did not retain a safe undo point.'
Assert-CustomPackTest ([string]$obsoletePage.volume -eq 'none' -and -not [bool]$obsoletePage.enabled) 'Removed output target was revived by layout restore.'

$duplicateSnapshotId = Save-LayoutSnapshot 'ja' $packId 'duplicate-page-test'
$duplicateSnapshotPath = Join-Path (Get-LayoutHistoryDir 'ja' $packId) ("{0}.json" -f $duplicateSnapshotId)
$duplicateSnapshot = Read-JsonFile $duplicateSnapshotPath $null
$duplicateSnapshot.pages = @($duplicateSnapshot.pages) + @($duplicateSnapshot.pages[0])
Write-JsonFile $duplicateSnapshotPath $duplicateSnapshot
$duplicateError = ''
try { [void](Get-LayoutRestorePreview 'ja' $packId $duplicateSnapshotId) } catch { $duplicateError = $_.Exception.Message }
Assert-CustomPackTest (-not [string]::IsNullOrWhiteSpace($duplicateError)) 'A corrupted layout snapshot with duplicate page IDs was accepted.'

# 控えの作成に失敗しても、出力の元になった版を守る pin は残ること。
# pin が無い版は保持期間の猶予なしに掃除の対象になるため、ここが崩れると
# 提出したPDFの元原稿が黙って消える。
$snapshotForPins = Add-FinalSnapshotSourceWorkbooks (Get-Structure 'ja') (Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' $mainVolume $packId)
$pinSourceWorkbooks = @(Get-Array (Get-DataProperty $snapshotForPins 'sourceWorkbooks' @()))
Assert-CustomPackTest ($pinSourceWorkbooks.Count -gt 0) 'The build snapshot carried no source workbooks, so the pin test proves nothing.'
$pinBuildId = 'archivefailure' + [Guid]::NewGuid().ToString('N').Substring(0, 8)
Set-FinalPdfSnapshotPins 'ja' $packId '' $mainVolume $pinBuildId $pinSourceWorkbooks ''
$pinMissing = @()
foreach ($sw in $pinSourceWorkbooks) {
    $pinWbId = [string](Get-DataProperty $sw 'workbookId' '')
    $pinSnapshotId = [string](Get-DataProperty $sw 'snapshotId' '')
    if ([string]::IsNullOrWhiteSpace($pinWbId) -or [string]::IsNullOrWhiteSpace($pinSnapshotId)) { continue }
    if (@(Get-SnapshotPins 'ja' $pinWbId $pinSnapshotId) -notcontains ("final-pdf_{0}" -f $pinBuildId)) { $pinMissing += $pinWbId }
}
Assert-CustomPackTest ($pinMissing.Count -eq 0) 'Snapshot pins were not created without an archive.'

# 出力したPDFの保存先を返せること。エクスプローラーは開かない(解決だけを見る)。
$revealPath = Resolve-OutputPdfForReveal 'ja' $packId 'main' ''
Assert-CustomPackTest ((Test-Path -LiteralPath $revealPath) -and $revealPath -eq ([IO.Path]::GetFullPath([string]$built.outputPdf))) 'Reveal did not resolve to the produced PDF.'
$revealOutsideError = ''
# 実在するファイルを出力フォルダーの外に置く。存在しないパスで試すと、境界の検査を
# 外しても「ファイルが見つかりません」で例外になり、検査を外したことに気付けない。
$outsidePdf = Join-Path ([IO.Path]::GetTempPath()) ('rb-outside-' + [Guid]::NewGuid().ToString('N') + '.pdf')
Copy-Item -LiteralPath ([string]$built.outputPdf) -Destination $outsidePdf -Force
$revealOriginalPdf = [string]$built.outputPdf
try {
    [void](Update-StructureLocked 'ja' { param($st) Set-NoteProperty (Get-PackOutputState $st 'ja' $packId 'main' $false) 'outputPdf' $outsidePdf })
    [void](Resolve-OutputPdfForReveal 'ja' $packId 'main' '')
} catch { $revealOutsideError = $_.Exception.Message }
[void](Update-StructureLocked 'ja' { param($st) Set-NoteProperty (Get-PackOutputState $st 'ja' $packId 'main' $false) 'outputPdf' $revealOriginalPdf })
Remove-Item -LiteralPath $outsidePdf -Force -ErrorAction SilentlyContinue
Assert-CustomPackTest ($revealOutsideError -like '*出力フォルダーの外*') 'Reveal accepted a path outside the output folder.'

Write-Output "custom-pack workflow selfcheck ok ($($pages.Count) source pages -> $([int]$ready.pageCount) output pages)"
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_CUSTOM_PACK_TEST_APPROOT', $oldAppRoot, 'Process')
    if (-not $KeepTestData -and (Test-Path -LiteralPath $testRoot)) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
