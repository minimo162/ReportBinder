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
function New-TestPresentationPdf([string]$Path, [int]$ExpectedPages) {
    $textPath = [IO.Path]::ChangeExtension($Path, '.txt')
    $java = Resolve-JavaExe
    $jar = Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'
    foreach ($lineCount in 20..240 | Where-Object { $_ % 10 -eq 0 }) {
        $lines = @(1..$lineCount | ForEach-Object { "Presentation fixture line $_" })
        [IO.File]::WriteAllLines($textPath, $lines, (New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
        $run = Invoke-NativeCapture $java @('-jar', $jar, 'TextToPDF', '-standardFont', 'Helvetica', '-fontSize', '12', $Path, $textPath)
        if ([int]$run.exitCode -ne 0) { throw "Test presentation PDF creation failed: $($run.text)" }
        if ([int](Get-PdfPageCount $Path) -eq $ExpectedPages) { return }
    }
    throw "Could not create a $ExpectedPages-page presentation PDF fixture."
}
$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalConfigRoot 'data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
$samplePath = Join-Path $submissionDir '01_ECM_adapter.xlsx'
[IO.File]::WriteAllBytes($samplePath, [byte[]](1,2,3,4,5))
$pptxSource = Join-Path $Script:LocalConfigRoot 'pptx-source'
New-Item -ItemType Directory -Path (Join-Path $pptxSource 'ppt\slides') -Force | Out-Null
[IO.File]::WriteAllText((Join-Path $pptxSource '[Content_Types].xml'), '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>', [Text.Encoding]::UTF8)
[IO.File]::WriteAllText((Join-Path $pptxSource 'ppt\presentation.xml'), '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>', [Text.Encoding]::UTF8)
foreach ($slide in 1..2) { [IO.File]::WriteAllText((Join-Path $pptxSource "ppt\slides\slide$slide.xml"), "<?xml version=`"1.0`" encoding=`"UTF-8`"?><p:sld xmlns:p=`"http://schemas.openxmlformats.org/presentationml/2006/main`"/>", [Text.Encoding]::UTF8) }
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
$pptxPath = Join-Path $submissionDir '02_ECM_presentation.pptx'
[IO.Compression.ZipFile]::CreateFromDirectory($pptxSource, $pptxPath)
Save-AppConfig ([pscustomobject][ordered]@{
    schemaVersion = 2; lastSubmissionDir = $submissionDir; lastDataDir = $dataDir; lastOutputDir = $outputDir
})
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')

$adapter = Get-SourceAdapterDescriptor 'excel'
Assert-AdapterTest ([string]$adapter.adapterId -eq 'excel-com-v1') 'Excel adapterId is invalid.'
Assert-AdapterTest ([int]$adapter.contractVersion -eq 1) 'Adapter contractVersion is invalid.'
Assert-AdapterTest (@($adapter.extensions) -contains '.xlsm') 'Macro-enabled Excel is not advertised by the adapter.'
$wordAdapter = Get-SourceAdapterDescriptor 'word'
Assert-AdapterTest ([string]$wordAdapter.adapterId -eq 'word-com-v1' -and [string]$wordAdapter.unitKind -eq 'document') 'Word adapter descriptor is invalid.'
$powerPointAdapter = Get-SourceAdapterDescriptor 'powerpoint'
Assert-AdapterTest ([string]$powerPointAdapter.adapterId -eq 'powerpoint-com-v1' -and [string]$powerPointAdapter.unitKind -eq 'slide' -and @($powerPointAdapter.extensions) -contains '.pptx') 'PowerPoint adapter descriptor is invalid.'
$unsupportedRejected = $false
try { [void](Get-SourceAdapterDescriptor 'image') } catch { $unsupportedRejected = $true }
Assert-AdapterTest $unsupportedRejected 'Unsupported source type was not rejected.'

$candidates = @(Get-SourceCandidates @('excel'))
Assert-AdapterTest ($candidates.Count -eq 1) 'Excel candidate count is invalid.'
Assert-AdapterTest ([string]$candidates[0].sourceType -eq 'excel' -and [string]$candidates[0].adapterId -eq 'excel-com-v1') 'Candidate adapter metadata is missing.'
$validated = Test-SourceCandidate ([pscustomobject]@{ relativePath='01_ECM_adapter.xlsx'; sourceType='excel' })
Assert-AdapterTest ([bool]$validated.ok) 'Excel candidate validation failed.'
$validatedXlsm = Test-SourceCandidate ([pscustomobject]@{ relativePath='01_ECM_adapter.xlsm'; sourceType='excel' })
Assert-AdapterTest ([bool]$validatedXlsm.ok -and (Resolve-SourceTypeFromPath '01_ECM_adapter.xlsm') -eq 'excel') 'Macro-enabled Excel candidate validation failed.'
$validatedPptx = Test-SourceCandidate ([pscustomobject]@{ relativePath='02_ECM_presentation.pptx'; sourceType='powerpoint' })
Assert-AdapterTest ([bool]$validatedPptx.ok -and (Resolve-SourceTypeFromPath '02_ECM_presentation.pptx') -eq 'powerpoint') 'PowerPoint candidate validation failed.'
$pptxInspection = Inspect-PowerPointSourceFile $pptxPath
Assert-AdapterTest ([int]$pptxInspection.pageCount -eq 2 -and @($pptxInspection.units).Count -eq 2) 'PowerPoint inspection did not enumerate slides.'
$powerPointCandidates = @(Get-SourceCandidates @('powerpoint'))
Assert-AdapterTest ($powerPointCandidates.Count -eq 1 -and [string]$powerPointCandidates[0].adapterId -eq 'powerpoint-com-v1') 'PowerPoint candidate scanning failed.'

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
$pptBatch = Register-SourcesBatch 'ja' @('02_ECM_presentation.pptx') 'pack_ecm' 'powerpoint'
Assert-AdapterTest ([int]$pptBatch.registeredCount -eq 1 -and [int]$pptBatch.errorCount -eq 0) 'PowerPoint source registration failed.'
$pptContext = Get-RegisteredSourceAdapterContext 'ja' ([string]$pptBatch.registered[0].sourceId)
Assert-AdapterTest ([string]$pptContext.sourceType -eq 'powerpoint' -and [string]$pptContext.adapter.adapterId -eq 'powerpoint-com-v1') 'Registered PowerPoint source did not resolve through the adapter dispatcher.'
$pptSourceId = [string]$pptBatch.registered[0].sourceId
$Script:PowerPointFixturePdf = Join-Path $Script:LocalConfigRoot 'powerpoint-render-fixture.pdf'
New-TestPresentationPdf $Script:PowerPointFixturePdf 2
function Invoke-PowerPointToPdf([string]$InputPath, [string]$OutputPath, [int]$TimeoutSeconds = 0) {
    Copy-Item -LiteralPath $Script:PowerPointFixturePdf -Destination $OutputPath -Force
    return [pscustomobject][ordered]@{ ok=$true; powerPointVersion='selfcheck'; outputPath=$OutputPath }
}
$pptRendered = Render-Source 'ja' $pptSourceId
Assert-AdapterTest ([string]$pptRendered.sourceType -eq 'powerpoint' -and @($pptRendered.rendered).Count -eq 2) 'PowerPoint rendering did not preserve every slide.'
$pptStructure = Get-Structure 'ja'
$pptSource = @($pptStructure.workbooks | Where-Object { [string]$_.workbookId -eq $pptSourceId } | Select-Object -First 1)[0]
$pptPages = @($pptStructure.pages | Where-Object { [string]$_.workbookId -eq $pptSourceId })
Assert-AdapterTest ([int]$pptSource.renderProfileVersion -eq [int]$Script:PowerPointRenderProfileVersion -and [int]$pptSource.sourcePageCount -eq 2) 'PowerPoint render profile metadata is invalid.'
Assert-AdapterTest ($pptPages.Count -eq 2 -and @($pptPages | Where-Object { [string]$_.sheetName -notmatch '^Slide [12]$' -or [string]::IsNullOrWhiteSpace([string]$_.contentPdf) }).Count -eq 0) 'PowerPoint slide page artifacts are missing.'
$structure = Get-Structure 'ja'
$v4 = ConvertTo-V4StructureCompatibilityView $structure
Assert-AdapterTest (@($v4.workbooks).Count -eq 2 -and @($v4.workbooks | Where-Object { [string]$_.workbookId -eq $sourceId }).Count -eq 1) 'V4 workbook compatibility was not preserved.'
$templates = @(Get-PackTemplateCatalog 'ja')
Assert-AdapterTest ($templates.Count -eq 1 -and [string]$templates[0].templateId -eq 'builtin-generic-department-pack' -and [string]$templates[0].packId -eq '') 'Pack template catalog must expose only the general-purpose starting template.'
$packs = @(Get-PublicPackList $structure 'ja')
Assert-AdapterTest ($packs.Count -eq 0) 'Legacy category packs must not appear in the public pack catalog.'
$updated = Update-SourceMetadata 'ja' $sourceId ([pscustomobject]@{ ownerDepartment='経理部'; required=$false; defaultTargetId='appendix' })
Assert-AdapterTest ([string]$updated.ownerDepartment -eq '経理部' -and -not [bool]$updated.required -and [string]$updated.defaultTargetId -eq 'appendix') 'Source metadata update result is invalid.'
$updatedStructure = Get-Structure 'ja'
$updatedSource = @($updatedStructure.sources | Where-Object { [string]$_.sourceId -eq $sourceId } | Select-Object -First 1)
Assert-AdapterTest ($updatedSource.Count -eq 1 -and [string]$updatedSource[0].ownerDepartment -eq '経理部' -and -not [bool]$updatedSource[0].required) 'Source metadata was not persisted.'
$updatedV4 = ConvertTo-V4StructureCompatibilityView $updatedStructure
Assert-AdapterTest ([string]$updatedV4.workbooks[0].ownerDepartment -eq '経理部' -and -not [bool]$updatedV4.workbooks[0].required) 'Source metadata was not mirrored to V4 compatibility.'
$v2State = Get-V2StatePayload 'ja'
Assert-AdapterTest ([int]$v2State.apiVersion -eq 2 -and [int]$v2State.domainSchemaVersion -eq 3 -and @($v2State.packs).Count -eq 0) 'V2 state payload is invalid.'
Write-Output 'source-adapter selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_ADAPTER_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
