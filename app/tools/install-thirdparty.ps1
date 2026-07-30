# ASCII-only installer to avoid Windows PowerShell 5.1 codepage issues.
# Downloads Apache PDFBox and Mozilla PDF.js into the local ReportBinder app folder.
param(
  [string]$PdfBoxVersion = "2.0.36",
  [string]$PdfJsVersion = "5.7.284",
  [switch]$SkipBuild,
  [switch]$SkipJava
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ToolRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppRoot = Resolve-Path (Join-Path $ToolRoot "..")
$PdfBoxDir = Join-Path $AppRoot "lib\pdfbox"
$PdfJsDir = Join-Path $AppRoot "web\pdfjs"
$JavaDir = Join-Path $AppRoot "lib\java"
$DownloadDir = Join-Path $AppRoot "thirdparty-cache"

New-Item -ItemType Directory -Force -Path $PdfBoxDir, $PdfJsDir, $DownloadDir | Out-Null

function Save-Url {
  param(
    [string[]]$Urls,
    [string]$Destination
  )
  $lastError = $null
  foreach ($url in $Urls) {
    try {
      Write-Host "Downloading: $url"
      Invoke-WebRequest -Uri $url -OutFile $Destination -UseBasicParsing -TimeoutSec 180
      $item = Get-Item -LiteralPath $Destination -ErrorAction Stop
      if ($item.Length -le 0) {
        throw "Downloaded file is empty: $Destination"
      }
      return
    } catch {
      $lastError = $_
      Write-Warning "Failed: $url"
      Write-Warning $_.Exception.Message
    }
  }
  throw "Download failed: $Destination`n$lastError"
}

function Verify-Sha512 {
  param(
    [string]$File,
    [string]$Sha512File
  )
  $shaText = Get-Content -LiteralPath $Sha512File -Raw
  $match = [regex]::Match($shaText, "[A-Fa-f0-9]{128}")
  if (-not $match.Success) {
    throw "SHA512 value not found: $Sha512File"
  }
  $expected = $match.Value.ToLowerInvariant()
  $actual = (Get-FileHash -LiteralPath $File -Algorithm SHA512).Hash.ToLowerInvariant()
  if ($expected -ne $actual) {
    throw "SHA512 mismatch for $File`nexpected: $expected`nactual:   $actual"
  }
  Write-Host "SHA512 OK: $File"
}


function Ensure-JavaRuntime {
  if ($SkipJava) {
    Write-Host "Java runtime install skipped by -SkipJava."
    return
  }

  $localJava = Join-Path $JavaDir "bin\java.exe"
  if (Test-Path -LiteralPath $localJava) {
    Write-Host "Java runtime: already present at $localJava"
    return
  }

  $systemJava = Get-Command java.exe -ErrorAction SilentlyContinue
  if ($systemJava) {
    Write-Host "Java runtime: system Java found at $($systemJava.Source)"
    return
  }

  Write-Host "Java runtime was not found. Downloading portable Eclipse Temurin JRE 17..."
  $jreZip = Join-Path $DownloadDir "temurin-jre-17-windows-x64.zip"
  Save-Url -Urls @(
    "https://api.adoptium.net/v3/binary/latest/17/ga/windows/x64/jre/hotspot/normal/eclipse?project=jdk"
  ) -Destination $jreZip

  $extractTmp = Join-Path $DownloadDir ("temurin-jre-extract-" + [Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Force -Path $extractTmp | Out-Null
  try {
    Expand-Archive -LiteralPath $jreZip -DestinationPath $extractTmp -Force
    $root = @(Get-ChildItem -LiteralPath $extractTmp -Directory | Sort-Object Name | Select-Object -First 1)
    if ($root.Count -eq 0) {
      throw "Portable Java archive did not contain a root directory."
    }
    if (Test-Path -LiteralPath $JavaDir) {
      Remove-Item -LiteralPath $JavaDir -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $JavaDir | Out-Null
    Copy-Item -Path (Join-Path $root[0].FullName "*") -Destination $JavaDir -Recurse -Force
    if (-not (Test-Path -LiteralPath $localJava)) {
      throw "java.exe was not found after extraction: $localJava"
    }
    Set-Content -LiteralPath (Join-Path $JavaDir "JAVA_VERSION.txt") -Value "Portable Eclipse Temurin JRE 17 installed by ReportBinder installer." -Encoding UTF8
    Write-Host "Java runtime installed: $localJava"
  } finally {
    if (Test-Path -LiteralPath $extractTmp) {
      Remove-Item -LiteralPath $extractTmp -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}

# Apache PDFBox 2.0.x app jar. The composer source currently targets PDFBox 2.x APIs.
$pdfBoxFileName = "pdfbox-app-$PdfBoxVersion.jar"
$pdfBoxTemp = Join-Path $DownloadDir $pdfBoxFileName
$pdfBoxSha = "$pdfBoxTemp.sha512"
$pdfBoxFinal = Join-Path $PdfBoxDir "pdfbox-app.jar"

Save-Url -Urls @(
  "https://dlcdn.apache.org/pdfbox/$PdfBoxVersion/$pdfBoxFileName",
  "https://downloads.apache.org/pdfbox/$PdfBoxVersion/$pdfBoxFileName",
  "https://archive.apache.org/dist/pdfbox/$PdfBoxVersion/$pdfBoxFileName"
) -Destination $pdfBoxTemp

Save-Url -Urls @(
  "https://downloads.apache.org/pdfbox/$PdfBoxVersion/$pdfBoxFileName.sha512",
  "https://archive.apache.org/dist/pdfbox/$PdfBoxVersion/$pdfBoxFileName.sha512"
) -Destination $pdfBoxSha

Verify-Sha512 -File $pdfBoxTemp -Sha512File $pdfBoxSha
Copy-Item -LiteralPath $pdfBoxTemp -Destination $pdfBoxFinal -Force
Set-Content -LiteralPath (Join-Path $PdfBoxDir "PDFBOX_VERSION.txt") -Value "Apache PDFBox $PdfBoxVersion`npdfbox-app.jar copied from $pdfBoxFileName" -Encoding UTF8

# Mozilla PDF.js generic build. Keep files local to avoid CDN dependency at runtime.
Save-Url -Urls @(
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@$PdfJsVersion/build/pdf.min.mjs",
  "https://unpkg.com/pdfjs-dist@$PdfJsVersion/build/pdf.min.mjs"
) -Destination (Join-Path $PdfJsDir "pdf.min.mjs")

Save-Url -Urls @(
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@$PdfJsVersion/build/pdf.worker.min.mjs",
  "https://unpkg.com/pdfjs-dist@$PdfJsVersion/build/pdf.worker.min.mjs"
) -Destination (Join-Path $PdfJsDir "pdf.worker.min.mjs")

Save-Url -Urls @(
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@$PdfJsVersion/LICENSE",
  "https://unpkg.com/pdfjs-dist@$PdfJsVersion/LICENSE"
) -Destination (Join-Path $PdfJsDir "LICENSE")

Set-Content -LiteralPath (Join-Path $PdfJsDir "PDFJS_VERSION.txt") -Value "Mozilla PDF.js / pdfjs-dist $PdfJsVersion`nFiles: pdf.min.mjs, pdf.worker.min.mjs" -Encoding UTF8

Write-Host ""
Write-Host "Third-party files installed."
Write-Host "PDFBox: $pdfBoxFinal"
Write-Host "PDF.js: $PdfJsDir"

Ensure-JavaRuntime

if (-not $SkipBuild) {
  $build = Join-Path $PdfBoxDir "build.ps1"
  $composerJar = Join-Path $PdfBoxDir "ReportPdfComposer.jar"

  if (Test-Path -LiteralPath $composerJar) {
    Write-Host "ReportPdfComposer.jar: already present. Build is skipped."
    Write-Host "To rebuild it from source, install OpenJDK/JDK and run app\lib\pdfbox\build.ps1."
  } elseif (Get-Command javac -ErrorAction SilentlyContinue) {
    Write-Host ""
    Write-Host "Building ReportPdfComposer.jar..."
    & $build
  } else {
    Write-Warning "javac was not found and ReportPdfComposer.jar is not present."
    Write-Warning "Install OpenJDK/JDK and run app\lib\pdfbox\build.ps1, or copy ReportPdfComposer.jar to app\lib\pdfbox."
  }
}
