# ASCII-only installer to avoid Windows PowerShell 5.1 codepage issues.
# Downloads and verifies Apache PDFBox, Mozilla PDF.js, and a portable Eclipse Temurin JRE.
param(
  [string]$PdfBoxVersion = "2.0.37",
  [string]$PdfJsVersion = "5.7.284",
  [int]$TemurinFeatureVersion = 17,
  [switch]$SkipBuild,
  [switch]$SkipJava,
  [switch]$PreferSystemJava,
  [switch]$Force
)

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ToolRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppRoot = [IO.Path]::GetFullPath((Join-Path $ToolRoot ".."))
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
  $parent = Split-Path -Parent $Destination
  if ($parent) { New-Item -ItemType Directory -Force -Path $parent | Out-Null }

  foreach ($url in $Urls) {
    $temp = "$Destination.download-$([Guid]::NewGuid().ToString('N'))"
    try {
      Write-Host "Downloading: $url"
      Invoke-WebRequest -Uri $url -OutFile $temp -UseBasicParsing -TimeoutSec 180
      $item = Get-Item -LiteralPath $temp -ErrorAction Stop
      if ($item.Length -le 0) {
        throw "Downloaded file is empty: $temp"
      }
      Move-Item -LiteralPath $temp -Destination $Destination -Force
      return
    } catch {
      $lastError = $_
      Write-Warning "Failed: $url"
      Write-Warning $_.Exception.Message
      Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
  }

  throw "Download failed: $Destination`n$lastError"
}

function Get-Json {
  param([string]$Url)
  Write-Host "Reading metadata: $Url"
  return Invoke-RestMethod -Uri $Url -Method Get -TimeoutSec 180
}

function Verify-ExpectedHash {
  param(
    [string]$File,
    [ValidateSet("SHA256", "SHA512")]
    [string]$Algorithm,
    [string]$Expected
  )

  $expectedValue = ([string]$Expected).Trim().ToLowerInvariant()
  $requiredLength = if ($Algorithm -eq "SHA256") { 64 } else { 128 }
  if ($expectedValue -notmatch ("^[a-f0-9]{" + $requiredLength + "}$")) {
    throw "Invalid $Algorithm value: $Expected"
  }

  $actual = (Get-FileHash -LiteralPath $File -Algorithm $Algorithm).Hash.ToLowerInvariant()
  if ($actual -ne $expectedValue) {
    throw "$Algorithm mismatch for $File`nexpected: $expectedValue`nactual:   $actual"
  }
  Write-Host "$Algorithm OK: $File"
}

function Verify-HashFile {
  param(
    [string]$File,
    [string]$HashFile,
    [ValidateSet("SHA256", "SHA512")]
    [string]$Algorithm
  )

  $length = if ($Algorithm -eq "SHA256") { 64 } else { 128 }
  $hashText = Get-Content -LiteralPath $HashFile -Raw
  $match = [regex]::Match($hashText, "[A-Fa-f0-9]{" + $length + "}")
  if (-not $match.Success) {
    throw "$Algorithm value not found: $HashFile"
  }
  Verify-ExpectedHash -File $File -Algorithm $Algorithm -Expected $match.Value
}

function Verify-SriSha512 {
  param(
    [string]$File,
    [string]$Integrity
  )

  $prefix = "sha512-"
  $value = ([string]$Integrity).Trim()
  if (-not $value.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
    throw "npm package integrity is not SHA-512: $Integrity"
  }

  $expected = $value.Substring($prefix.Length)
  $sha = [Security.Cryptography.SHA512]::Create()
  try {
    $bytes = [IO.File]::ReadAllBytes($File)
    $actual = [Convert]::ToBase64String($sha.ComputeHash($bytes))
  } finally {
    $sha.Dispose()
  }

  if ($actual -ne $expected) {
    throw "npm SHA-512 integrity mismatch for $File"
  }
  Write-Host "npm SHA-512 integrity OK: $File"
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

function Install-PdfBox {
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

  Verify-HashFile -File $pdfBoxTemp -HashFile $pdfBoxSha -Algorithm SHA512
  Copy-Item -LiteralPath $pdfBoxTemp -Destination $pdfBoxFinal -Force
  Set-Content -LiteralPath (Join-Path $PdfBoxDir "PDFBOX_VERSION.txt") `
    -Value "Apache PDFBox $PdfBoxVersion`npdfbox-app.jar copied from $pdfBoxFileName`nSHA-512 verified" -Encoding UTF8
}

function Install-PdfJs {
  $metadataUrl = "https://registry.npmjs.org/pdfjs-dist/$PdfJsVersion"
  $metadata = Get-Json -Url $metadataUrl
  $tarballUrl = [string]$metadata.dist.tarball
  $integrity = [string]$metadata.dist.integrity

  if ([string]::IsNullOrWhiteSpace($tarballUrl) -or [string]::IsNullOrWhiteSpace($integrity)) {
    throw "npm metadata did not contain dist.tarball and dist.integrity for pdfjs-dist $PdfJsVersion"
  }

  $packageFile = Join-Path $DownloadDir "pdfjs-dist-$PdfJsVersion.tgz"
  Save-Url -Urls @($tarballUrl) -Destination $packageFile
  Verify-SriSha512 -File $packageFile -Integrity $integrity

  $tar = Get-Command tar.exe -ErrorAction SilentlyContinue
  if (-not $tar) { $tar = Get-Command tar -ErrorAction SilentlyContinue }
  if (-not $tar) {
    throw "tar.exe was not found. Windows 10/11 includes tar.exe; install it or place PDF.js manually."
  }

  $extractTmp = Join-Path $DownloadDir ("pdfjs-extract-" + [Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Force -Path $extractTmp | Out-Null
  try {
    [void](Invoke-Native -FilePath $tar.Source -Arguments @("-xzf", $packageFile, "-C", $extractTmp))
    $packageRoot = Join-Path $extractTmp "package"
    $files = @(
      @{ Source = (Join-Path $packageRoot "build\pdf.min.mjs"); Destination = (Join-Path $PdfJsDir "pdf.min.mjs"); MinimumBytes = 100000 },
      @{ Source = (Join-Path $packageRoot "build\pdf.worker.min.mjs"); Destination = (Join-Path $PdfJsDir "pdf.worker.min.mjs"); MinimumBytes = 500000 },
      @{ Source = (Join-Path $packageRoot "LICENSE"); Destination = (Join-Path $PdfJsDir "LICENSE"); MinimumBytes = 1000 }
    )

    foreach ($entry in $files) {
      if (-not (Test-Path -LiteralPath $entry.Source)) {
        throw "Required PDF.js file was not found in the verified npm package: $($entry.Source)"
      }
      $item = Get-Item -LiteralPath $entry.Source
      if ($item.Length -lt [int64]$entry.MinimumBytes) {
        throw "PDF.js file is unexpectedly small: $($entry.Source) ($($item.Length) bytes)"
      }
      Copy-Item -LiteralPath $entry.Source -Destination $entry.Destination -Force
    }
  } finally {
    Remove-Item -LiteralPath $extractTmp -Recurse -Force -ErrorAction SilentlyContinue
  }

  Set-Content -LiteralPath (Join-Path $PdfJsDir "PDFJS_VERSION.txt") `
    -Value "Mozilla PDF.js / pdfjs-dist $PdfJsVersion`nSource: verified npm package tarball`nFiles: pdf.min.mjs, pdf.worker.min.mjs" -Encoding UTF8
}

function Install-PortableJava {
  if ($SkipJava) {
    Write-Host "Java runtime install skipped by -SkipJava."
    return $false
  }

  $localJava = Join-Path $JavaDir "bin\java.exe"
  if ((Test-Path -LiteralPath $localJava) -and -not $Force) {
    Write-Host "Portable Java runtime: already present at $localJava"
    return $true
  }

  if ($PreferSystemJava -and -not $Force) {
    $systemJava = Get-Command java.exe -ErrorAction SilentlyContinue
    if (-not $systemJava) { $systemJava = Get-Command java -ErrorAction SilentlyContinue }
    if ($systemJava) {
      Write-Host "System Java selected by -PreferSystemJava: $($systemJava.Source)"
      return $false
    }
  }

  Write-Host "Downloading the latest portable Eclipse Temurin JRE $TemurinFeatureVersion GA for Windows x64..."
  $metadataUrl = "https://api.adoptium.net/v3/assets/latest/$TemurinFeatureVersion/hotspot?architecture=x64&heap_size=normal&image_type=jre&jvm_impl=hotspot&os=windows&project=jdk&vendor=eclipse"
  $assets = @(Get-Json -Url $metadataUrl)
  $asset = @($assets | Where-Object {
    $_.binary -and $_.binary.package -and
    -not [string]::IsNullOrWhiteSpace([string]$_.binary.package.link) -and
    -not [string]::IsNullOrWhiteSpace([string]$_.binary.package.checksum)
  } | Select-Object -First 1)

  if ($asset.Count -eq 0) {
    throw "Adoptium API did not return a Windows x64 JRE asset for Java $TemurinFeatureVersion."
  }

  $package = $asset[0].binary.package
  $releaseName = [string]$asset[0].release_name
  $semver = [string]$asset[0].version.semver
  $downloadUrl = [string]$package.link
  $checksum = [string]$package.checksum
  $jreZip = Join-Path $DownloadDir ("temurin-jre-" + $TemurinFeatureVersion + "-windows-x64.zip")

  Save-Url -Urls @($downloadUrl) -Destination $jreZip
  Verify-ExpectedHash -File $jreZip -Algorithm SHA256 -Expected $checksum

  $extractTmp = Join-Path $DownloadDir ("temurin-jre-extract-" + [Guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Force -Path $extractTmp | Out-Null
  try {
    Expand-Archive -LiteralPath $jreZip -DestinationPath $extractTmp -Force
    $javaExe = @(Get-ChildItem -LiteralPath $extractTmp -Filter java.exe -Recurse -File |
      Where-Object { $_.FullName -match "[\\/]bin[\\/]java\.exe$" } |
      Sort-Object FullName |
      Select-Object -First 1)

    if ($javaExe.Count -eq 0) {
      throw "Portable Java archive did not contain bin\java.exe."
    }

    $runtimeRoot = Split-Path -Parent (Split-Path -Parent $javaExe[0].FullName)
    if (Test-Path -LiteralPath $JavaDir) {
      Remove-Item -LiteralPath $JavaDir -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $JavaDir | Out-Null
    Copy-Item -Path (Join-Path $runtimeRoot "*") -Destination $JavaDir -Recurse -Force

    if (-not (Test-Path -LiteralPath $localJava)) {
      throw "java.exe was not found after extraction: $localJava"
    }

    $versionOutput = Invoke-Native -FilePath $localJava -Arguments @("-version")
    $info = @(
      "Eclipse Temurin portable JRE installed by ReportBinder",
      "Release: $releaseName",
      "SemVer: $semver",
      "Feature version: $TemurinFeatureVersion",
      "Platform: Windows x64 / HotSpot / JRE",
      "SHA-256: $checksum",
      "Source: $downloadUrl",
      "java -version:",
      $versionOutput
    ) -join "`r`n"
    Set-Content -LiteralPath (Join-Path $JavaDir "JAVA_VERSION.txt") -Value $info -Encoding UTF8
    Write-Host "Portable Java runtime installed: $localJava"
  } finally {
    Remove-Item -LiteralPath $extractTmp -Recurse -Force -ErrorAction SilentlyContinue
  }

  return $true
}

Install-PdfBox
Install-PdfJs
$portableJavaInstalled = Install-PortableJava

Write-Host ""
Write-Host "Third-party files installed and verified."
Write-Host "PDFBox: $(Join-Path $PdfBoxDir 'pdfbox-app.jar')"
Write-Host "PDF.js: $PdfJsDir"
if ($portableJavaInstalled) {
  Write-Host "Java: $(Join-Path $JavaDir 'bin\java.exe')"
} elseif ($SkipJava) {
  Write-Host "Java: skipped"
} else {
  Write-Host "Java: system runtime selected"
}

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
    if ($LASTEXITCODE -ne 0) { throw "ReportPdfComposer.jar build failed." }
  } else {
    throw "javac was not found and ReportPdfComposer.jar is not present. Restore the tracked JAR or install a JDK and run app\lib\pdfbox\build.ps1."
  }
}

$verify = Join-Path $ToolRoot "verify-thirdparty.ps1"
if (-not (Test-Path -LiteralPath $verify)) {
  throw "Dependency verification script is missing: $verify"
}

if ($SkipJava) {
  & $verify -RequirePdfJs
} elseif ($PreferSystemJava -and -not $portableJavaInstalled) {
  & $verify -RequirePdfJs -RequireJava
} else {
  & $verify -RequirePdfJs -RequirePortableJava
}
if ($LASTEXITCODE -ne 0) { throw "Third-party dependency verification failed." }
