# Verifies the external runtime payloads used by ReportBinder.
param(
  [switch]$RequirePdfJs,
  [switch]$RequireJava,
  [switch]$RequirePortableJava
)

$ErrorActionPreference = "Stop"
$ToolRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppRoot = [IO.Path]::GetFullPath((Join-Path $ToolRoot ".."))
$PdfBoxDir = Join-Path $AppRoot "lib\pdfbox"
$PdfJsDir = Join-Path $AppRoot "web\pdfjs"
$JavaDir = Join-Path $AppRoot "lib\java"

function Assert-File {
  param(
    [string]$Path,
    [int64]$MinimumBytes = 1
  )
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
    throw "Required file is missing: $Path"
  }
  $item = Get-Item -LiteralPath $Path
  if ($item.Length -lt $MinimumBytes) {
    throw "Required file is unexpectedly small: $Path ($($item.Length) bytes; minimum $MinimumBytes)"
  }
  return $item
}

function Assert-ZipEntry {
  param(
    [string]$ArchivePath,
    [string]$EntryName
  )
  Add-Type -AssemblyName System.IO.Compression | Out-Null
  Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
  $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
  try {
    $entry = $archive.GetEntry($EntryName)
    if ($null -eq $entry) {
      throw "Archive entry is missing: $ArchivePath -> $EntryName"
    }
  } finally {
    $archive.Dispose()
  }
}

function Invoke-Native {
  param(
    [string]$FilePath,
    [string[]]$Arguments
  )
  $previous = $ErrorActionPreference
  $ErrorActionPreference = "Continue"
  try {
    $output = @(& $FilePath @Arguments 2>&1 | ForEach-Object { [string]$_ })
    $exitCode = $LASTEXITCODE
  } finally {
    $ErrorActionPreference = $previous
  }
  if ($exitCode -ne 0) {
    throw "Native command failed ($exitCode): $FilePath $($Arguments -join ' ')`n$($output -join "`n")"
  }
  return ($output -join "`n")
}

$pdfBoxJar = Join-Path $PdfBoxDir "pdfbox-app.jar"
$composerJar = Join-Path $PdfBoxDir "ReportPdfComposer.jar"
$pdfBoxVersionFile = Join-Path $PdfBoxDir "PDFBOX_VERSION.txt"

[void](Assert-File -Path $pdfBoxJar -MinimumBytes 1000000)
[void](Assert-File -Path $composerJar -MinimumBytes 10000)
[void](Assert-File -Path $pdfBoxVersionFile -MinimumBytes 20)
Assert-ZipEntry -ArchivePath $pdfBoxJar -EntryName "org/apache/pdfbox/pdmodel/PDDocument.class"
foreach ($entry in @("ReportPdfComposer.class", "BatchPdfSplitter.class", "PdfPageAnalyzer.class")) {
  Assert-ZipEntry -ArchivePath $composerJar -EntryName $entry
}

$pdfBoxVersionText = Get-Content -LiteralPath $pdfBoxVersionFile -Raw
if ($pdfBoxVersionText -notmatch "Apache PDFBox 2\.0\.37") {
  throw "PDFBox version metadata must identify 2.0.37: $pdfBoxVersionFile"
}
if ($pdfBoxVersionText -notmatch "SHA-512 verified") {
  throw "PDFBox version metadata does not record SHA-512 verification: $pdfBoxVersionFile"
}

if ($RequirePdfJs) {
  $pdfJsMain = Join-Path $PdfJsDir "pdf.min.mjs"
  $pdfJsWorker = Join-Path $PdfJsDir "pdf.worker.min.mjs"
  $pdfJsLicense = Join-Path $PdfJsDir "LICENSE"
  $pdfJsVersion = Join-Path $PdfJsDir "PDFJS_VERSION.txt"

  [void](Assert-File -Path $pdfJsMain -MinimumBytes 100000)
  [void](Assert-File -Path $pdfJsWorker -MinimumBytes 500000)
  [void](Assert-File -Path $pdfJsLicense -MinimumBytes 1000)
  [void](Assert-File -Path $pdfJsVersion -MinimumBytes 20)

  $pdfJsVersionText = Get-Content -LiteralPath $pdfJsVersion -Raw
  if ($pdfJsVersionText -notmatch "pdfjs-dist 5\.7\.284") {
    throw "PDF.js version metadata must identify 5.7.284: $pdfJsVersion"
  }
  if ($pdfJsVersionText -notmatch "verified npm package tarball") {
    throw "PDF.js version metadata does not record npm integrity verification: $pdfJsVersion"
  }
}

$javaExe = Join-Path $JavaDir "bin\java.exe"
if ($RequirePortableJava) {
  [void](Assert-File -Path $javaExe -MinimumBytes 10000)
  [void](Assert-File -Path (Join-Path $JavaDir "release") -MinimumBytes 10)
  [void](Assert-File -Path (Join-Path $JavaDir "NOTICE") -MinimumBytes 10)
  [void](Assert-File -Path (Join-Path $JavaDir "JAVA_VERSION.txt") -MinimumBytes 100)
  $javaInfo = Get-Content -LiteralPath (Join-Path $JavaDir "JAVA_VERSION.txt") -Raw
  if ($javaInfo -notmatch "SHA-256:\s*[A-Fa-f0-9]{64}") {
    throw "Portable Java metadata does not record a verified SHA-256 checksum."
  }
  [void](Invoke-Native -FilePath $javaExe -Arguments @("-version"))
} elseif ($RequireJava) {
  $resolvedJava = $null
  if (Test-Path -LiteralPath $javaExe) {
    $resolvedJava = $javaExe
  } else {
    $command = Get-Command java.exe -ErrorAction SilentlyContinue
    if (-not $command) { $command = Get-Command java -ErrorAction SilentlyContinue }
    if ($command) { $resolvedJava = [string]$command.Source }
  }
  if ([string]::IsNullOrWhiteSpace($resolvedJava)) {
    throw "Java Runtime was not found locally or on PATH."
  }
  [void](Invoke-Native -FilePath $resolvedJava -Arguments @("-version"))
}

Write-Host "Third-party dependency verification OK."
Write-Host "PDFBox 2.0.37: $pdfBoxJar"
if ($RequirePdfJs) { Write-Host "PDF.js 5.7.284: $PdfJsDir" }
if ($RequirePortableJava) { Write-Host "Portable Java: $javaExe" }
