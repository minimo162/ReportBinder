param(
  [string]$PdfBoxAppJar = (Join-Path $PSScriptRoot 'pdfbox-app.jar')
)

$ErrorActionPreference = 'Stop'
$src = Join-Path $PSScriptRoot 'src\ReportPdfComposer.java'
$splitSrc = Join-Path $PSScriptRoot 'src\BatchPdfSplitter.java'
$analyzerSrc = Join-Path $PSScriptRoot 'src\PdfPageAnalyzer.java'
$batchRasterizerSrc = Join-Path $PSScriptRoot 'src\PdfBatchRasterizer.java'
$classes = Join-Path $PSScriptRoot 'classes'
$outJar = Join-Path $PSScriptRoot 'ReportPdfComposer.jar'

if (-not (Test-Path -LiteralPath $PdfBoxAppJar)) {
  throw "pdfbox-app.jar が見つかりません: $PdfBoxAppJar`nApache PDFBox 2.x の pdfbox-app.jar を app\lib\pdfbox に配置してください。"
}
if (Test-Path -LiteralPath $classes) { Remove-Item -LiteralPath $classes -Recurse -Force }
New-Item -ItemType Directory -Path $classes -Force | Out-Null

$releaseArgs = @()
$helpText = (& javac --help 2>&1 | Out-String)
if ($helpText -match '--release') {
  # Keep the composer runnable on the bundled Java 17 runtime even when it is rebuilt with a newer JDK.
  $releaseArgs = @('--release', '8')
}
& javac @releaseArgs -encoding UTF-8 -cp $PdfBoxAppJar -d $classes $src $splitSrc $analyzerSrc $batchRasterizerSrc
if ($LASTEXITCODE -ne 0) { throw 'javac failed' }
if (Test-Path -LiteralPath $outJar) { Remove-Item -LiteralPath $outJar -Force }
Push-Location $classes
try {
  & jar cfe $outJar ReportPdfComposer *
  if ($LASTEXITCODE -ne 0) { throw 'jar failed' }
} finally {
  Pop-Location
}
Remove-Item -LiteralPath $classes -Recurse -Force
Write-Host "created: $outJar"
