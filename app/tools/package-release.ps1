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
    Get-ChildItem -LiteralPath $StageRoot -Recurse -Force -Directory | Where-Object { $_.Name -eq '_reportbinder' } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

function New-ReleaseZip([string]$Suffix, [bool]$IncludeJava) {
    $tempBase = Join-Path ([IO.Path]::GetTempPath()) "ReportBinderPackage_$([guid]::NewGuid().ToString('N'))"
    $stage = Join-Path $tempBase 'ReportBinder'
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    Copy-Item -Path (Join-Path $root '*') -Destination $stage -Recurse -Force
    Remove-RuntimeFiles $stage
    if (-not $IncludeJava) {
        $java = Join-Path $stage 'app\lib\java'
        if (Test-Path -LiteralPath $java) { Remove-Item -LiteralPath $java -Recurse -Force }
    }
    $zip = Join-Path $OutputDir "ReportBinder_V5_${Suffix}_${stamp}.zip"
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
    New-Utf8Zip $stage $zip
    Remove-Item -LiteralPath $tempBase -Recurse -Force
    return $zip
}

function New-Utf8Zip([string]$SourceDir, [string]$ZipPath) {
    # V5-P0(#1): Compress-Archive はエントリ名を UTF-8 バイトで書くが「言語エンコーディングフラグ(bit11)」を立てない。
    # そのため日本語 Windows の標準展開(エクスプローラー)が CP932 と誤解し、日本語ファイル名が文字化けする。
    # ZipArchive を UTF8 エンコーディングで開くと非ASCII名のエントリに bit11 が立ち、標準展開でも文字化けしない。
    # 併せて、ZIP 仕様どおりパス区切りを '/' に統一する(.NET Framework の CreateFromDirectory は '\' を書いてしまう)。
    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $baseParent = Split-Path -Parent $SourceDir
    $fs = [System.IO.File]::Open($ZipPath, [System.IO.FileMode]::Create)
    $archive = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false, [System.Text.Encoding]::UTF8)
    try {
        foreach ($file in (Get-ChildItem -LiteralPath $SourceDir -Recurse -File -Force)) {
            $rel = $file.FullName.Substring($baseParent.Length + 1).Replace([char]92, [char]47)
            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $file.FullName, $rel, [System.IO.Compression.CompressionLevel]::Optimal)
        }
    } finally {
        $archive.Dispose(); $fs.Dispose()
    }
}

Invoke-SelfCheck
$offline = New-ReleaseZip 'オフライン完結版' $true
$online  = New-ReleaseZip 'オンライン導入版' $false
Write-Host "作成しました:`n$offline`n$online"
