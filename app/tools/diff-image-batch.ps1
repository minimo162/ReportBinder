param(
    [Parameter(Mandatory = $true)]
    [string]$RequestPath,
    [int]$Dpi = 120,
    [int]$Threshold = 24,
    [int]$MinimumRegionPixels = 24,
    [int]$Padding = 5
)

$ErrorActionPreference = 'Stop'
$toolsRoot = $PSScriptRoot
$appRoot = Split-Path -Parent $toolsRoot
$javaExe = Join-Path $appRoot 'lib\java\bin\java.exe'
$composerJar = Join-Path $appRoot 'lib\pdfbox\ReportPdfComposer.jar'
$pdfBoxJar = Join-Path $appRoot 'lib\pdfbox\pdfbox-app.jar'
$engineSource = Join-Path $toolsRoot 'DiffImageEngine.cs'

foreach ($required in @($javaExe, $composerJar, $pdfBoxJar, $engineSource, $RequestPath)) {
    if (-not (Test-Path -LiteralPath $required)) { throw "差分一括処理に必要なファイルが見つかりません: $required" }
}

$request = Get-Content -LiteralPath $RequestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$items = @($request.items)
if ($items.Count -eq 0) {
    [ordered]@{ ok = $true; items = @() } | ConvertTo-Json -Depth 12 -Compress
    exit 0
}

$rasterRoot = Join-Path ([IO.Path]::GetTempPath()) ('rb-diff-batch-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $rasterRoot -Force | Out-Null
$tsvPath = Join-Path $rasterRoot 'requests.tsv'

function ConvertTo-PathBase64([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($Value)))
}

function Get-RasterPages([string]$ItemDirectory, [string]$Prefix) {
    return @(Get-ChildItem -LiteralPath $ItemDirectory -File -Filter ($Prefix + '-*.png') -ErrorAction SilentlyContinue |
        Sort-Object {
            $match = [regex]::Match($_.BaseName, '(\d+)$')
            if ($match.Success) { [int]$match.Groups[1].Value } else { [int]::MaxValue }
        } |
        ForEach-Object { $_.FullName })
}

try {
    $lines = @()
    foreach ($item in $items) {
        $id = [string]$item.id
        if ($id -notmatch '^[A-Za-z0-9_-]+$') { throw "差分一括処理IDが不正です: $id" }
        $lines += ($id + "`t" + (ConvertTo-PathBase64 ([string]$item.beforePdf)) + "`t" + (ConvertTo-PathBase64 ([string]$item.afterPdf)))
    }
    [IO.File]::WriteAllLines($tsvPath, $lines, [Text.UTF8Encoding]::new($false))

    $threads = [Math]::Max(1, [Math]::Min(4, [Environment]::ProcessorCount))
    $nativeOutput = @()
    $nativeExitCode = -1
    $previousPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        $nativeOutput = @(& $javaExe '-Djava.awt.headless=true' '-cp' "$composerJar;$pdfBoxJar" 'PdfBatchRasterizer' '--request' $tsvPath '--output' $rasterRoot '--dpi' $Dpi '--threads' $threads 2>&1 |
            ForEach-Object { [string]$_ })
        $nativeExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousPreference
    }
    if ($nativeExitCode -ne 0) {
        throw ("PDF一括画像化に失敗しました。`n" + ($nativeOutput -join "`n"))
    }

    if (-not ('ReportBinderDiffEngine' -as [type])) {
        Add-Type -Path $engineSource -ReferencedAssemblies @('System.Drawing')
    }

    $results = @()
    foreach ($item in $items) {
        $id = [string]$item.id
        $kind = [string]$item.kind
        $itemRasterDir = Join-Path $rasterRoot $id
        $result = [ordered]@{ id = $id; ok = $false; message = ''; beforePageCount = 0; afterPageCount = 0; pages = @() }
        try {
            foreach ($side in @('before','after')) {
                $errorPath = Join-Path $itemRasterDir ($side + '.error.txt')
                if (Test-Path -LiteralPath $errorPath) {
                    throw (Get-Content -LiteralPath $errorPath -Raw -Encoding UTF8)
                }
            }
            $beforePages = @(Get-RasterPages $itemRasterDir 'before')
            $afterPages = @(Get-RasterPages $itemRasterDir 'after')
            $pageCount = [Math]::Max($beforePages.Count, $afterPages.Count)
            if ($pageCount -eq 0) { throw '比較できるPDFページがありません。' }
            $outputDirectory = [IO.Path]::GetFullPath([string]$item.outputDirectory)
            if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
            $pages = @()
            for ($i = 0; $i -lt $pageCount; $i++) {
                $beforeImage = if ($i -lt $beforePages.Count) { [string]$beforePages[$i] } else { '' }
                $afterImage = if ($i -lt $afterPages.Count) { [string]$afterPages[$i] } else { '' }
                $pageKind = $kind
                if ($kind -notin @('added','removed','unknown')) {
                    if ([string]::IsNullOrWhiteSpace($beforeImage)) { $pageKind = 'added' }
                    elseif ([string]::IsNullOrWhiteSpace($afterImage)) { $pageKind = 'removed' }
                }
                $pages += [ReportBinderDiffEngine]::ComparePage(
                    $beforeImage, $afterImage, $outputDirectory, ($i + 1), $pageKind,
                    $Threshold, $MinimumRegionPixels, $Padding)
            }
            $result.ok = $true
            $result.beforePageCount = $beforePages.Count
            $result.afterPageCount = $afterPages.Count
            $result.pages = @($pages)
        } catch {
            $result.message = $_.Exception.Message
        }
        $results += [pscustomobject]$result
    }
    [ordered]@{
        ok = $true
        dpi = $Dpi
        threshold = $Threshold
        minimumRegionPixels = $MinimumRegionPixels
        padding = $Padding
        items = @($results)
    } | ConvertTo-Json -Depth 12 -Compress
} finally {
    if (Test-Path -LiteralPath $rasterRoot) {
        Remove-Item -LiteralPath $rasterRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
