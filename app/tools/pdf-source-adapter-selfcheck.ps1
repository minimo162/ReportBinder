param()

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-pdf-adapter-' + [Guid]::NewGuid().ToString('N'))
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_PDF_ADAPTER_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_PDF_ADAPTER_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_PDF_ADAPTER_APPROOT')
    $testBody = @'
function Assert-PdfAdapterTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function New-TestPdf([string]$Path, [int]$LineCount) {
    $textPath = [IO.Path]::ChangeExtension($Path, '.txt')
    $lines = @(1..$LineCount | ForEach-Object { "Department report line $_ with table-like values A$_ B$_ C$_" })
    [IO.File]::WriteAllLines($textPath, $lines, (New-Object Text.UTF8Encoding($false)))
    $java = Resolve-JavaExe
    $jar = Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'
    $run = Invoke-NativeCapture $java @('-jar', $jar, 'TextToPDF', '-standardFont', 'Helvetica', '-fontSize', '12', $Path, $textPath)
    if ([int]$run.exitCode -ne 0) { throw "Test PDF creation failed: $($run.text)" }
}

$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalConfigRoot 'data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
Save-AppConfig ([pscustomobject][ordered]@{
    schemaVersion = 2; lastSubmissionDir = $submissionDir; lastDataDir = $dataDir; lastOutputDir = $outputDir
})
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')

$pdfPath = Join-Path $submissionDir '02_ECM_department-pack.pdf'
New-TestPdf $pdfPath 180
$adapter = Get-SourceAdapterDescriptor 'pdf'
Assert-PdfAdapterTest ([string]$adapter.adapterId -eq 'pdfbox-import-v1' -and [string]$adapter.unitKind -eq 'pdf-page') 'PDF adapter descriptor is invalid.'
$candidates = @(Get-SourceCandidates @('excel','pdf'))
Assert-PdfAdapterTest ($candidates.Count -eq 1 -and [string]$candidates[0].sourceType -eq 'pdf') 'PDF candidate discovery failed.'
$inspection = Inspect-Source ([pscustomobject]@{ relativePath='02_ECM_department-pack.pdf'; sourceType='pdf' })
Assert-PdfAdapterTest ([int]$inspection.pageCount -ge 2 -and @($inspection.units).Count -eq [int]$inspection.pageCount) 'PDF physical page inspection failed.'

$badPath = Join-Path $submissionDir 'broken.pdf'
[IO.File]::WriteAllText($badPath, 'not a pdf')
$brokenRejected = $false
try { [void](Inspect-PdfSourceFile $badPath) } catch { $brokenRejected = $true }
Assert-PdfAdapterTest $brokenRejected 'Broken PDF was not rejected.'

$encryptedPath = Join-Path $submissionDir 'protected.pdf'
$encryptRun = Invoke-NativeCapture (Resolve-JavaExe) @('-jar', (Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'), 'Encrypt', '-O', 'ownerpass', '-U', 'userpass', $pdfPath, $encryptedPath)
Assert-PdfAdapterTest ([int]$encryptRun.exitCode -eq 0) 'Encrypted PDF test fixture creation failed.'
$encryptedRejected = $false
try { [void](Inspect-PdfSourceFile $encryptedPath) } catch { $encryptedRejected = ($_.Exception.Message -match 'PDF') }
Assert-PdfAdapterTest $encryptedRejected 'Password-protected PDF was not rejected.'

$batch = Register-SourcesBatch 'ja' @('02_ECM_department-pack.pdf') 'pack_ecm' ''
Assert-PdfAdapterTest ([int]$batch.registeredCount -eq 1 -and [int]$batch.errorCount -eq 0) 'PDF registration failed.'
$sourceId = [string]$batch.registered[0].sourceId
$rendered = Render-Source 'ja' $sourceId
Assert-PdfAdapterTest ([string]$rendered.sourceType -eq 'pdf' -and @($rendered.rendered).Count -eq [int]$inspection.pageCount) 'PDF rendering did not preserve every physical page.'
$structure = Get-Structure 'ja'
$source = @($structure.workbooks | Where-Object { [string]$_.workbookId -eq $sourceId } | Select-Object -First 1)[0]
$pages = @($structure.pages | Where-Object { [string]$_.workbookId -eq $sourceId })
Assert-PdfAdapterTest ([string]$source.status -notin @('new','source-updated','render-error') -and [int]$source.renderProfileVersion -eq [int]$Script:PdfImportProfileVersion) 'PDF source did not become current.'
Assert-PdfAdapterTest ($pages.Count -eq [int]$inspection.pageCount -and @($pages | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.contentPdf) }).Count -eq 0) 'PDF page artifacts are missing.'

Start-Sleep -Milliseconds 50
New-TestPdf $pdfPath 250
$scan = Scan-Updates 'ja'
$changed = @($scan.changedWorkbookIds | Where-Object { [string]$_ -eq $sourceId })
Assert-PdfAdapterTest ($changed.Count -eq 1) 'PDF replacement was not detected.'
$updated = Get-Structure 'ja'
$updatedSource = @($updated.workbooks | Where-Object { [string]$_.workbookId -eq $sourceId } | Select-Object -First 1)[0]
Assert-PdfAdapterTest ([string]$updatedSource.status -eq 'source-updated') 'PDF replacement did not set source-updated.'

$rerendered = Render-Source 'ja' $sourceId
$after = Get-Structure 'ja'
$afterPages = @($after.pages | Where-Object { [string]$_.workbookId -eq $sourceId })
Assert-PdfAdapterTest (@($rerendered.rendered).Count -gt @($rendered.rendered).Count -and $afterPages.Count -eq @($rerendered.rendered).Count) 'PDF page additions were not synchronized after replacement.'
$v2 = Get-V2StatePayload 'ja'
$v2Source = @($v2.structure.sources | Where-Object { [string]$_.sourceId -eq $sourceId })
Assert-PdfAdapterTest ($v2Source.Count -eq 1 -and [string]$v2Source[0].sourceType -eq 'pdf') 'V2 PDF source projection is invalid.'
$templates = @(Get-PackTemplateCatalog 'ja')
Assert-PdfAdapterTest (@($templates[0].acceptedSourceTypes) -contains 'pdf') 'Built-in pack does not advertise PDF support.'
Write-Output "pdf-source-adapter selfcheck ok ($($afterPages.Count) pages)"
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_PDF_ADAPTER_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
