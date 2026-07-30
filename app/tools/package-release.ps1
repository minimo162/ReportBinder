param(
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'ReportBinderRelease')
)

$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

function Invoke-SelfCheck {
    $script = Join-Path $root 'app\tools\selfcheck.py'
    $python = Get-Command py -ErrorAction SilentlyContinue
    if ($python) { & $python.Source -3 $script; if ($LASTEXITCODE -ne 0) { throw 'selfcheckに失敗しました。' }; return }
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) { & $python.Source $script; if ($LASTEXITCODE -ne 0) { throw 'selfcheckに失敗しました。' }; return }
    throw 'Pythonが見つかりません。先に app\tools\selfcheck.py を実行してください。'
}

function Invoke-ThirdPartyCheck([bool]$RequirePortableJava) {
    $script = Join-Path $root 'app\tools\verify-thirdparty.ps1'
    if (-not (Test-Path -LiteralPath $script)) {
        throw "依存物検証スクリプトが見つかりません: $script"
    }
    if ($RequirePortableJava) {
        & $script -RequirePdfJs -RequirePortableJava
    } else {
        & $script -RequirePdfJs
    }
    if ($LASTEXITCODE -ne 0) {
        throw '第三者依存物の検証に失敗しました。app\tools\install-thirdparty.cmd を実行してください。'
    }
}

function Assert-StagedDependencies([string]$StageRoot, [bool]$IncludeJava) {
    $required = @(
        (Join-Path $StageRoot 'app\lib\pdfbox\pdfbox-app.jar'),
        (Join-Path $StageRoot 'app\lib\pdfbox\ReportPdfComposer.jar'),
        (Join-Path $StageRoot 'app\web\pdfjs\pdf.min.mjs'),
        (Join-Path $StageRoot 'app\web\pdfjs\pdf.worker.min.mjs'),
        (Join-Path $StageRoot 'app\web\pdfjs\LICENSE')
    )
    if ($IncludeJava) {
        $required += (Join-Path $StageRoot 'app\lib\java\bin\java.exe')
        $required += (Join-Path $StageRoot 'app\lib\java\release')
        $required += (Join-Path $StageRoot 'app\lib\java\NOTICE')
        $required += (Join-Path $StageRoot 'app\lib\java\JAVA_VERSION.txt')
    }
    $missing = @($required | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($missing.Count -gt 0) {
        throw ("配布物に必要な依存ファイルがありません:`n" + ($missing -join "`n"))
    }
}

function Remove-RuntimeFiles([string]$StageRoot) {
    foreach ($path in @(
        (Join-Path $StageRoot 'app\config.json'),
        (Join-Path $StageRoot 'app\thirdparty-cache'),
        (Join-Path $StageRoot 'app\logs'),
        (Join-Path $StageRoot 'app\lib\pdfbox\classes')
    )) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force }
    }
    New-Item -ItemType Directory -Path (Join-Path $StageRoot 'app\logs') -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $StageRoot 'app\logs\.gitkeep') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $StageRoot 'app\thirdparty-cache') -Force | Out-Null
    Get-ChildItem -LiteralPath $StageRoot -Recurse -Force -File | Where-Object {
        $_.Name -like 'job_*.json' -or $_.Name -like '~building_*.pdf' -or
        $_.Name -like '*.tmp' -or $_.Name -like '*.pyc'
    } | Remove-Item -Force -ErrorAction SilentlyContinue
    Get-ChildItem -LiteralPath $StageRoot -Recurse -Force -Directory |
        Where-Object { $_.Name -eq '__pycache__' } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Get-ChildItem -LiteralPath $StageRoot -Recurse -Force -Directory |
        Where-Object { $_.Name -eq '_reportbinder' } |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

function New-ReleaseZip([string]$Suffix, [bool]$IncludeJava) {
    $tempBase = Join-Path ([IO.Path]::GetTempPath()) "ReportBinderPackage_$([guid]::NewGuid().ToString('N'))"
    $stage = Join-Path $tempBase 'ReportBinder'
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    try {
        Copy-Item -Path (Join-Path $root '*') -Destination $stage -Recurse -Force
        Remove-RuntimeFiles $stage
        if (-not $IncludeJava) {
            $java = Join-Path $stage 'app\lib\java'
            if (Test-Path -LiteralPath $java) { Remove-Item -LiteralPath $java -Recurse -Force }
        }
        Assert-StagedDependencies -StageRoot $stage -IncludeJava $IncludeJava
        $zip = Join-Path $OutputDir "ReportBinder_V5_${Suffix}_${stamp}.zip"
        if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
        New-Utf8Zip $stage $zip
        return $zip
    } finally {
        if (Test-Path -LiteralPath $tempBase) {
            Remove-Item -LiteralPath $tempBase -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-Utf8Zip([string]$SourceDir, [string]$ZipPath) {
    # Compress-Archive does not reliably mark non-ASCII entry names as UTF-8 for Windows Explorer.
    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $baseParent = Split-Path -Parent $SourceDir
    $fs = [System.IO.File]::Open($ZipPath, [System.IO.FileMode]::Create)
    $archive = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false, [System.Text.Encoding]::UTF8)
    try {
        foreach ($file in (Get-ChildItem -LiteralPath $SourceDir -Recurse -File -Force)) {
            $rel = $file.FullName.Substring($baseParent.Length + 1).Replace([char]92, [char]47)
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive,
                $file.FullName,
                $rel,
                [System.IO.Compression.CompressionLevel]::Optimal
            )
        }
    } finally {
        $archive.Dispose()
        $fs.Dispose()
    }
}

Invoke-SelfCheck
# PDFBox and PDF.js are included in both release variants.
Invoke-ThirdPartyCheck $false
# The offline-complete release must contain its own verified portable JRE even if system Java exists.
Invoke-ThirdPartyCheck $true

$offline = New-ReleaseZip 'オフライン完結版' $true
$online  = New-ReleaseZip 'オンライン導入版' $false
Write-Host "作成しました:`n$offline`n$online"
