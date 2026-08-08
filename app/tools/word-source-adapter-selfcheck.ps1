param(
    [string]$FixtureV1 = $env:REPORTBINDER_WORD_FIXTURE_V1,
    [string]$FixtureV2 = $env:REPORTBINDER_WORD_FIXTURE_V2
)

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-word-adapter-' + [Guid]::NewGuid().ToString('N'))
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_WORD_ADAPTER_APPROOT', 'Process')

if ([string]::IsNullOrWhiteSpace($FixtureV1) -or -not (Test-Path -LiteralPath $FixtureV1 -PathType Leaf)) { throw 'FixtureV1 DOCX is required.' }
if ([string]::IsNullOrWhiteSpace($FixtureV2) -or -not (Test-Path -LiteralPath $FixtureV2 -PathType Leaf)) { throw 'FixtureV2 DOCX is required.' }

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_WORD_ADAPTER_APPROOT', $appRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_WORD_FIXTURE_V1_INNER', ([IO.Path]::GetFullPath($FixtureV1)), 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_WORD_FIXTURE_V2_INNER', ([IO.Path]::GetFullPath($FixtureV2)), 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_WORD_ADAPTER_APPROOT')
    $testBody = @'
function Assert-WordAdapterTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
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

$wordPath = Join-Path $submissionDir '03_ECM_department-report.docx'
Copy-Item -LiteralPath $env:REPORTBINDER_WORD_FIXTURE_V1_INNER -Destination $wordPath -Force
$adapter = Get-SourceAdapterDescriptor 'word'
Assert-WordAdapterTest ([string]$adapter.adapterId -eq 'word-com-v1' -and [string]$adapter.unitKind -eq 'document') 'Word adapter descriptor is invalid.'
$candidates = @(Get-SourceCandidates @('excel','word','pdf'))
Assert-WordAdapterTest ($candidates.Count -eq 1 -and [string]$candidates[0].sourceType -eq 'word') 'Word candidate discovery failed.'
$inspection = Inspect-Source ([pscustomobject]@{ relativePath='03_ECM_department-report.docx'; sourceType='word' })
Assert-WordAdapterTest ([int]$inspection.pageCount -eq 0 -and @($inspection.units).Count -eq 1 -and [string]$inspection.units[0].unitKind -eq 'document') 'Word package inspection failed.'

$badPath = Join-Path $submissionDir 'broken.docx'
[IO.File]::WriteAllText($badPath, 'not a docx')
$brokenRejected = $false
try { [void](Inspect-WordSourceFile $badPath) } catch { $brokenRejected = ($_.Exception.Message -match 'Word|DOCX') }
Assert-WordAdapterTest $brokenRejected 'Broken DOCX was not rejected.'
$encryptedPath = Join-Path $submissionDir 'protected.docx'
[IO.File]::WriteAllBytes($encryptedPath, [byte[]](0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1))
$encryptedRejected = $false
try { [void](Inspect-WordSourceFile $encryptedPath) } catch { $encryptedRejected = $true }
Assert-WordAdapterTest $encryptedRejected 'Encrypted-like DOCX was not rejected.'
Remove-Item -LiteralPath $badPath,$encryptedPath -Force

$batch = Register-SourcesBatch 'ja' @('03_ECM_department-report.docx') 'pack_ecm' ''
Assert-WordAdapterTest ([int]$batch.registeredCount -eq 1 -and [int]$batch.errorCount -eq 0) 'Word registration failed.'
$sourceId = [string]$batch.registered[0].sourceId
$rendered = Render-Source 'ja' $sourceId
Assert-WordAdapterTest ([string]$rendered.sourceType -eq 'word' -and @($rendered.rendered).Count -ge 5) 'Word rendering did not produce the expected physical pages.'
$structure = Get-Structure 'ja'
$source = @($structure.workbooks | Where-Object { [string]$_.workbookId -eq $sourceId } | Select-Object -First 1)[0]
$pages = @($structure.pages | Where-Object { [string]$_.workbookId -eq $sourceId })
Assert-WordAdapterTest ([string]$source.status -notin @('new','source-updated','render-error') -and [int]$source.renderProfileVersion -eq [int]$Script:WordRenderProfileVersion) 'Word source did not become current.'
Assert-WordAdapterTest ($pages.Count -eq @($rendered.rendered).Count -and @($pages | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.contentPdf) }).Count -eq 0) 'Word page artifacts are missing.'
$firstPageCount = $pages.Count

Copy-Item -LiteralPath $env:REPORTBINDER_WORD_FIXTURE_V2_INNER -Destination $wordPath -Force
[IO.File]::SetLastWriteTimeUtc($wordPath, [DateTime]::UtcNow.AddSeconds(2))
$scan = Scan-Updates 'ja' $null $true
$changed = @($scan.changedWorkbookIds | Where-Object { [string]$_ -eq $sourceId })
Assert-WordAdapterTest ($changed.Count -eq 1) 'Word replacement was not detected.'
$updated = Get-Structure 'ja'
$updatedSource = @($updated.workbooks | Where-Object { [string]$_.workbookId -eq $sourceId } | Select-Object -First 1)[0]
Assert-WordAdapterTest ([string]$updatedSource.status -eq 'source-updated') 'Word replacement did not set source-updated.'

$rerendered = Render-Source 'ja' $sourceId
$after = Get-Structure 'ja'
$afterPages = @($after.pages | Where-Object { [string]$_.workbookId -eq $sourceId })
Assert-WordAdapterTest (@($rerendered.rendered).Count -gt $firstPageCount -and $afterPages.Count -eq @($rerendered.rendered).Count) 'Word page additions were not synchronized after replacement.'
$summary = Get-SourceChangeSummary 'ja' $sourceId
Assert-WordAdapterTest ($null -ne $summary -and [string]$summary.sourceType -eq 'word') 'Word change summary was not produced.'
Assert-WordAdapterTest ((@($summary.changedSheets).Count + @($summary.addedSheets).Count) -gt 0) 'Word visual changes were not detected.'
$v2 = Get-V2StatePayload 'ja'
$v2Source = @($v2.structure.sources | Where-Object { [string]$_.sourceId -eq $sourceId })
Assert-WordAdapterTest ($v2Source.Count -eq 1 -and [string]$v2Source[0].sourceType -eq 'word') 'V2 Word source projection is invalid.'
$templates = @(Get-PackTemplateCatalog 'ja')
Assert-WordAdapterTest (@($templates[0].acceptedSourceTypes) -contains 'word') 'Built-in pack does not advertise Word support.'
Write-Output "word-source-adapter selfcheck ok ($firstPageCount -> $($afterPages.Count) pages)"
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_WORD_ADAPTER_APPROOT', $oldAppRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_WORD_FIXTURE_V1_INNER', $null, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_WORD_FIXTURE_V2_INNER', $null, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
