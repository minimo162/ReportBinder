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
$totalWatch = [Diagnostics.Stopwatch]::StartNew()
$rasterMs = 0
$analysisMs = 0
$cacheHitSides = 0
$rasterizedSides = 0
$unchangedPagesSkipped = 0

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

function Get-CachedRasterPages([string]$Directory, [int]$ExpectedCount) {
    if ([string]::IsNullOrWhiteSpace($Directory) -or -not (Test-Path -LiteralPath $Directory)) { return @() }
    $pages = @(Get-ChildItem -LiteralPath $Directory -File -Filter 'page-*.png' -ErrorAction SilentlyContinue |
        Sort-Object {
            $match = [regex]::Match($_.BaseName, '(\d+)$')
            if ($match.Success) { [int]$match.Groups[1].Value } else { [int]::MaxValue }
        } |
        ForEach-Object { $_.FullName })
    if ($ExpectedCount -gt 0 -and $pages.Count -ne $ExpectedCount) { return @() }
    return $pages
}

try {
    $lines = @()
    $cachedPagesById = @{}
    $needsRaster = $false
    foreach ($item in $items) {
        $id = [string]$item.id
        if ($id -notmatch '^[A-Za-z0-9_-]+$') { throw "差分一括処理IDが不正です: $id" }
        $beforeCached = @(Get-CachedRasterPages ([string]$item.beforeRasterDirectory) ([int]$item.beforePageCount))
        $afterCached = @(Get-CachedRasterPages ([string]$item.afterRasterDirectory) ([int]$item.afterPageCount))
        $beforePdf = $(if ($beforeCached.Count -gt 0) { '' } else { [string]$item.beforePdf })
        $afterPdf = $(if ($afterCached.Count -gt 0) { '' } else { [string]$item.afterPdf })
        if ($beforeCached.Count -gt 0) { $cacheHitSides++ }
        elseif (-not [string]::IsNullOrWhiteSpace($beforePdf)) { $rasterizedSides++; $needsRaster = $true }
        if ($afterCached.Count -gt 0) { $cacheHitSides++ }
        elseif (-not [string]::IsNullOrWhiteSpace($afterPdf)) { $rasterizedSides++; $needsRaster = $true }
        $cachedPagesById[$id] = [ordered]@{ before = @($beforeCached); after = @($afterCached) }
        $lines += ($id + "`t" + (ConvertTo-PathBase64 $beforePdf) + "`t" + (ConvertTo-PathBase64 $afterPdf))
    }
    [IO.File]::WriteAllLines($tsvPath, $lines, [Text.UTF8Encoding]::new($false))

    $threads = [Math]::Max(1, [Math]::Min(4, [Environment]::ProcessorCount))
    $rasterWatch = [Diagnostics.Stopwatch]::StartNew()
    if ($needsRaster) {
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
    }
    $rasterWatch.Stop()
    $rasterMs = [int64]$rasterWatch.ElapsedMilliseconds

    if (-not ('ReportBinderDiffEngine' -as [type])) {
        Add-Type -Path $engineSource -ReferencedAssemblies @('System.Drawing')
    }

    $analysisWatch = [Diagnostics.Stopwatch]::StartNew()
    $results = @()
    $resultById = @{}
    $pagesById = @{}
    $pageRequests = New-Object 'System.Collections.Generic.List[ReportBinderDiffBatchPageRequest]'
    foreach ($item in $items) {
        $id = [string]$item.id
        $kind = [string]$item.kind
        $unchangedPageSet = @{}
        foreach ($rawPageNumber in @($item.unchangedPageNumbers)) {
            $samePageNumber = [int]$rawPageNumber
            if ($samePageNumber -gt 0) { $unchangedPageSet[$samePageNumber] = $true }
        }
        $itemRasterDir = Join-Path $rasterRoot $id
        $result = [ordered]@{ id = $id; ok = $false; message = ''; beforePageCount = 0; afterPageCount = 0; pages = @() }
        try {
            foreach ($side in @('before','after')) {
                $errorPath = Join-Path $itemRasterDir ($side + '.error.txt')
                if (Test-Path -LiteralPath $errorPath) {
                    throw (Get-Content -LiteralPath $errorPath -Raw -Encoding UTF8)
                }
            }
            $cached = $cachedPagesById[$id]
            $beforePages = @($cached.before)
            $afterPages = @($cached.after)
            if ($beforePages.Count -eq 0) { $beforePages = @(Get-RasterPages $itemRasterDir 'before') }
            if ($afterPages.Count -eq 0) { $afterPages = @(Get-RasterPages $itemRasterDir 'after') }
            $pageCount = [Math]::Max($beforePages.Count, $afterPages.Count)
            if ($pageCount -eq 0) { throw '比較できるPDFページがありません。' }
            $outputDirectory = [IO.Path]::GetFullPath([string]$item.outputDirectory)
            if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
            $pagesById[$id] = New-Object System.Collections.ArrayList
            for ($i = 0; $i -lt $pageCount; $i++) {
                $beforeImage = if ($i -lt $beforePages.Count) { [string]$beforePages[$i] } else { '' }
                $afterImage = if ($i -lt $afterPages.Count) { [string]$afterPages[$i] } else { '' }
                $pageKind = $kind
                if ($kind -notin @('added','removed','unknown')) {
                    if ([string]::IsNullOrWhiteSpace($beforeImage)) { $pageKind = 'added' }
                    elseif ([string]::IsNullOrWhiteSpace($afterImage)) { $pageKind = 'removed' }
                    elseif ($unchangedPageSet.ContainsKey($i + 1)) {
                        # SHA-256が完全一致するページは、C#側で画像の複製だけを行う。
                        $pageKind = 'unchanged'
                        $unchangedPagesSkipped++
                    }
                }
                $pageRequests.Add([ReportBinderDiffBatchPageRequest]@{
                    itemId = $id
                    beforePath = $beforeImage
                    afterPath = $afterImage
                    outputDirectory = $outputDirectory
                    pageNumber = ($i + 1)
                    kind = $pageKind
                })
            }
            $result.ok = $true
            $result.beforePageCount = $beforePages.Count
            $result.afterPageCount = $afterPages.Count
        } catch {
            $result.message = $_.Exception.Message
        }
        $resultObject = [pscustomobject]$result
        $resultById[$id] = $resultObject
        $results += $resultObject
    }
    if ($pageRequests.Count -gt 0) {
        $pageResults = [ReportBinderDiffEngine]::ComparePages(
            $pageRequests.ToArray(), $threads, $Threshold, $MinimumRegionPixels, $Padding)
        foreach ($pageResult in @($pageResults)) {
            $id = [string]$pageResult.itemId
            $resultObject = $resultById[$id]
            if ($null -eq $resultObject) { continue }
            if (-not [string]::IsNullOrWhiteSpace([string]$pageResult.error)) {
                $resultObject.ok = $false
                if ([string]::IsNullOrWhiteSpace([string]$resultObject.message)) {
                    $resultObject.message = [string]$pageResult.error
                }
                continue
            }
            if ($pagesById.ContainsKey($id)) {
                [void]$pagesById[$id].Add($pageResult.page)
            }
        }
    }
    foreach ($resultObject in @($results)) {
        $id = [string]$resultObject.id
        if ($pagesById.ContainsKey($id)) {
            $resultObject.pages = @($pagesById[$id] | Sort-Object pageNumber)
        }
    }
    $analysisWatch.Stop()
    $analysisMs = [int64]$analysisWatch.ElapsedMilliseconds
    $totalWatch.Stop()
    [ordered]@{
        ok = $true
        dpi = $Dpi
        threshold = $Threshold
        minimumRegionPixels = $MinimumRegionPixels
        padding = $Padding
        timings = [ordered]@{
            totalMs = [int64]$totalWatch.ElapsedMilliseconds
            rasterMs = $rasterMs
            analysisMs = $analysisMs
            analysisThreads = $threads
            analyzedPages = $pageRequests.Count
            fullyAnalyzedPages = [Math]::Max(0, $pageRequests.Count - $unchangedPagesSkipped)
            unchangedPagesSkipped = $unchangedPagesSkipped
            cacheHitSides = $cacheHitSides
            rasterizedSides = $rasterizedSides
        }
        items = @($results)
    } | ConvertTo-Json -Depth 12 -Compress
} finally {
    if (Test-Path -LiteralPath $rasterRoot) {
        Remove-Item -LiteralPath $rasterRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
