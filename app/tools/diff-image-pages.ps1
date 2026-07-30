param(
    [string]$BeforePdf = '',
    [string]$AfterPdf = '',
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,
    [ValidateSet('modified','changed','added','removed','unknown','unchanged')]
    [string]$Kind = 'modified',
    [int]$Dpi = 150,
    [int]$Threshold = 18,
    [int]$MinimumRegionPixels = 16,
    [int]$Padding = 6
)

$ErrorActionPreference = 'Stop'
$toolsRoot = $PSScriptRoot
$appRoot = Split-Path -Parent $toolsRoot
$javaExe = Join-Path $appRoot 'lib\java\bin\java.exe'
$pdfBoxJar = Join-Path $appRoot 'lib\pdfbox\pdfbox-app.jar'
$engineSource = Join-Path $toolsRoot 'DiffImageEngine.cs'

if (-not (Test-Path -LiteralPath $javaExe)) { throw "Java runtime が見つかりません: $javaExe" }
if (-not (Test-Path -LiteralPath $pdfBoxJar)) { throw "PDFBox が見つかりません: $pdfBoxJar" }
if (-not (Test-Path -LiteralPath $engineSource)) { throw "差分画像エンジンが見つかりません: $engineSource" }

$outputFull = [IO.Path]::GetFullPath($OutputDirectory)
if (-not (Test-Path -LiteralPath $outputFull)) { New-Item -ItemType Directory -Path $outputFull -Force | Out-Null }
# PDFBoxの出力先にGUID付き一時名を足すと、深い履歴フォルダでは
# Windows PowerShell 5.1 / JavaのMAX_PATHを超える。作業画像は短いTEMPへ置き、
# 完成した比較PNGだけを履歴キャッシュへ保存する。
$rasterRoot = Join-Path ([IO.Path]::GetTempPath()) ('rb-diff-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $rasterRoot -Force | Out-Null

function Invoke-PdfRaster([string]$PdfPath, [string]$Prefix) {
    if ([string]::IsNullOrWhiteSpace($PdfPath)) { return @() }
    $full = [IO.Path]::GetFullPath($PdfPath)
    if (-not (Test-Path -LiteralPath $full)) { throw "比較対象PDFが見つかりません: $full" }
    $prefixPath = Join-Path $rasterRoot ($Prefix + '-')
    $previousPreference = $ErrorActionPreference
    $nativeOutput = @()
    $nativeExitCode = -1
    try {
        $ErrorActionPreference = 'Continue'
        $nativeOutput = @(& $javaExe -jar $pdfBoxJar PDFToImage -dpi $Dpi -format png -color rgb -outputPrefix $prefixPath $full 2>&1 | ForEach-Object { [string]$_ })
        $nativeExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousPreference
    }
    if ($nativeExitCode -ne 0) {
        throw ("PDFを画像化できませんでした。`n" + ($nativeOutput -join "`n"))
    }
    return @(Get-ChildItem -LiteralPath $rasterRoot -File -Filter ($Prefix + '-*.png') -ErrorAction SilentlyContinue |
        Sort-Object {
            $m = [regex]::Match($_.BaseName, '(\d+)$')
            if ($m.Success) { [int]$m.Groups[1].Value } else { [int]::MaxValue }
        } |
        ForEach-Object { $_.FullName })
}

try {
    if (-not ('ReportBinderDiffEngine' -as [type])) {
        Add-Type -Path $engineSource -ReferencedAssemblies @('System.Drawing')
    }

    $beforePages = @(Invoke-PdfRaster $BeforePdf 'before')
    $afterPages = @(Invoke-PdfRaster $AfterPdf 'after')
    $pageCount = [Math]::Max($beforePages.Count, $afterPages.Count)
    if ($pageCount -eq 0) { throw '比較できるPDFページがありません。' }

    $results = @()
    for ($i = 0; $i -lt $pageCount; $i++) {
        $beforeImage = if ($i -lt $beforePages.Count) { [string]$beforePages[$i] } else { '' }
        $afterImage = if ($i -lt $afterPages.Count) { [string]$afterPages[$i] } else { '' }
        $pageKind = $Kind
        if ($Kind -notin @('added','removed','unknown')) {
            if ([string]::IsNullOrWhiteSpace($beforeImage)) { $pageKind = 'added' }
            elseif ([string]::IsNullOrWhiteSpace($afterImage)) { $pageKind = 'removed' }
        }
        $results += [ReportBinderDiffEngine]::ComparePage(
            $beforeImage,
            $afterImage,
            $outputFull,
            ($i + 1),
            $pageKind,
            $Threshold,
            $MinimumRegionPixels,
            $Padding)
    }
    [ordered]@{
        ok = $true
        dpi = $Dpi
        threshold = $Threshold
        minimumRegionPixels = $MinimumRegionPixels
        padding = $Padding
        beforePageCount = $beforePages.Count
        afterPageCount = $afterPages.Count
        pages = @($results)
    } | ConvertTo-Json -Depth 12 -Compress
} finally {
    if (Test-Path -LiteralPath $rasterRoot) {
        Remove-Item -LiteralPath $rasterRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
