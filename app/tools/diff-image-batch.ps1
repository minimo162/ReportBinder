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
$reusedRasterPageAssets = 0
$rasterHashCache = @{}
$rasterSignatureCache = @{}

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
Add-Type -AssemblyName System.Drawing

function Get-DiffRasterHash([string]$Path, [bool]$ContentOnly = $false) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return '' }
    $cacheKey = ('{0}|{1}' -f $(if ($ContentOnly) { 'content' } else { 'full' }),[IO.Path]::GetFullPath($Path))
    if ($rasterHashCache.ContainsKey($cacheKey)) { return [string]$rasterHashCache[$cacheKey] }
    $stream = $null; $bitmap = $null; $cropped = $null; $memory = $null
    try {
        if ($ContentOnly) {
            $bitmap = [Drawing.Bitmap]::new($Path)
            # Running headers/footers often contain only a date or physical page
            # number. Ignore the outer bands for correspondence, while retaining
            # the full-image hash below so those edits are still reported.
            $top = [Math]::Max(0, [int][Math]::Floor($bitmap.Height * 0.06))
            $bottom = [Math]::Min($bitmap.Height, [int][Math]::Ceiling($bitmap.Height * 0.92))
            $rect = [Drawing.Rectangle]::new(0, $top, $bitmap.Width, [Math]::Max(1, $bottom - $top))
            $cropped = $bitmap.Clone($rect, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $memory = [IO.MemoryStream]::new()
            $cropped.Save($memory, [Drawing.Imaging.ImageFormat]::Png)
            $memory.Position = 0
            $stream = $memory
        } else {
            $stream = [IO.File]::OpenRead($Path)
        }
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            $result = ([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-', '').ToLowerInvariant()
            $rasterHashCache[$cacheKey] = $result
            return $result
        }
        finally { $sha.Dispose() }
    } finally {
        if ($null -ne $stream -and $stream -ne $memory) { $stream.Dispose() }
        if ($null -ne $memory) { $memory.Dispose() }
        if ($null -ne $cropped) { $cropped.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
    }
}

function New-DiffPageMapping([int]$BeforeIndex, [int]$AfterIndex, [string]$Kind, [string]$Method, [bool]$Ambiguous = $false, [string]$Message = '') {
    return [pscustomobject][ordered]@{
        beforeIndex = $BeforeIndex
        afterIndex = $AfterIndex
        kind = $Kind
        method = $Method
        ambiguous = $Ambiguous
        message = $Message
    }
}

function Get-DiffRasterSignature([string]$Path) {
    $cacheKey = [IO.Path]::GetFullPath($Path)
    if ($rasterSignatureCache.ContainsKey($cacheKey)) { return ,([byte[]]$rasterSignatureCache[$cacheKey]) }
    $bitmap = $null; $sample = $null; $graphics = $null
    try {
        $bitmap = [Drawing.Bitmap]::new($Path)
        $size = 64
        $sample = [Drawing.Bitmap]::new($size, $size, [Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $graphics = [Drawing.Graphics]::FromImage($sample)
        $graphics.Clear([Drawing.Color]::White)
        $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBilinear
        $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $sourceTop = [Math]::Max(0, [int][Math]::Floor($bitmap.Height * 0.06))
        $sourceBottom = [Math]::Min($bitmap.Height, [int][Math]::Ceiling($bitmap.Height * 0.92))
        $source = [Drawing.Rectangle]::new(0, $sourceTop, $bitmap.Width, [Math]::Max(1, $sourceBottom - $sourceTop))
        $graphics.DrawImage($bitmap, [Drawing.Rectangle]::new(0, 0, $size, $size), $source, [Drawing.GraphicsUnit]::Pixel)
        $values = [byte[]]::new($size * $size)
        for ($y = 0; $y -lt $size; $y++) {
            for ($x = 0; $x -lt $size; $x++) {
                $color = $sample.GetPixel($x, $y)
                $values[$y * $size + $x] = [byte][Math]::Max(0, [Math]::Min(255, [int][Math]::Round(0.299 * $color.R + 0.587 * $color.G + 0.114 * $color.B)))
            }
        }
        $rasterSignatureCache[$cacheKey] = $values
        return ,$values
    } finally {
        if ($null -ne $graphics) { $graphics.Dispose() }
        if ($null -ne $sample) { $sample.Dispose() }
        if ($null -ne $bitmap) { $bitmap.Dispose() }
    }
}

function Get-DiffRasterDistance($BeforeSignature, $AfterSignature) {
    $before = [byte[]]$BeforeSignature; $after = [byte[]]$AfterSignature
    if ($before.Length -eq 0 -or $before.Length -ne $after.Length) { return 1.0 }
    if ('ReportBinderDiffEngine' -as [type]) { return [ReportBinderDiffEngine]::MeanAbsoluteByteDistance($before, $after) }
    [double]$difference = 0
    for ($i = 0; $i -lt $before.Length; $i++) { $difference += [Math]::Abs([int]$before[$i] - [int]$after[$i]) }
    return [Math]::Min(1.0, $difference / ($before.Length * 255.0))
}

function Get-DiffPageMappings($BeforePages, $AfterPages, [string]$ItemKind, $UnchangedPageSet) {
    $before = @($BeforePages); $after = @($AfterPages)
    $m = $before.Count; $n = $after.Count
    if ($ItemKind -eq 'added' -or $m -eq 0) {
        return @($(for ($j = 0; $j -lt $n; $j++) { New-DiffPageMapping -1 $j 'added' 'one-sided' }))
    }
    if ($ItemKind -eq 'removed' -or $n -eq 0) {
        return @($(for ($i = 0; $i -lt $m; $i++) { New-DiffPageMapping $i -1 'removed' 'one-sided' }))
    }

    # A physical page number is positional, so comparing page N only with page N
    # turns every following page into a false change after an insertion/deletion.
    # Use visual distance over the page sequence: unlike exact hashes, this still
    # aligns pages when a running label/date changes on every page at the same time.
    $beforeFullHashes = @($before | ForEach-Object { Get-DiffRasterHash ([string]$_) })
    $afterFullHashes = @($after | ForEach-Object { Get-DiffRasterHash ([string]$_) })
    $beforeHashes = @($before | ForEach-Object { Get-DiffRasterHash ([string]$_) $true })
    $afterHashes = @($after | ForEach-Object { Get-DiffRasterHash ([string]$_) $true })
    $beforeSignatures = @($before | ForEach-Object { ,(Get-DiffRasterSignature ([string]$_)) })
    $afterSignatures = @($after | ForEach-Object { ,(Get-DiffRasterSignature ([string]$_)) })
    $distance = New-Object 'double[,]' $m,$n
    $rowMinimums = @(); $columnMinimums = @()
    for ($i = 0; $i -lt $m; $i++) {
        [double]$rowMinimum = 1.0
        for ($j = 0; $j -lt $n; $j++) {
            $distance[$i,$j] = if ($beforeHashes[$i] -and $beforeHashes[$i] -eq $afterHashes[$j]) { 0.0 } else { Get-DiffRasterDistance $beforeSignatures[$i] $afterSignatures[$j] }
            $currentDistance = [double]$distance[$i,$j]
            $rowMinimum = [Math]::Min($rowMinimum, $currentDistance)
        }
        $rowMinimums += $rowMinimum
    }
    for ($j = 0; $j -lt $n; $j++) {
        [double]$columnMinimum = 1.0
        for ($i = 0; $i -lt $m; $i++) {
            $currentDistance = [double]$distance[$i,$j]
            $columnMinimum = [Math]::Min($columnMinimum, $currentDistance)
        }
        $columnMinimums += $columnMinimum
    }
    $minimums = @($rowMinimums + $columnMinimums | Sort-Object)
    $alignmentBaseline = if ($minimums.Count -eq 0) { 0.0 } elseif ($minimums.Count % 2) {
        [double]$minimums[[int][Math]::Floor($minimums.Count / 2)]
    } else {
        ([double]$minimums[$minimums.Count / 2 - 1] + [double]$minimums[$minimums.Count / 2]) / 2.0
    }
    # A rewritten page should normally stay paired (substitution cap 0.005),
    # while a run of near-identical pages shifted by one position should prefer
    # one remove+add operation (2 * 0.0028) over several small mismatches.
    [double]$gapCost = 0.0028
    [double]$substitutionCap = 0.005
    $cost = New-Object 'double[,]' ($m + 1),($n + 1)
    $choice = New-Object 'byte[,]' ($m + 1),($n + 1) # 1=pair, 2=remove, 3=add
    for ($i = 1; $i -le $m; $i++) { $cost[$i,0] = $i * $gapCost; $choice[$i,0] = 2 }
    for ($j = 1; $j -le $n; $j++) { $cost[0,$j] = $j * $gapCost; $choice[0,$j] = 3 }
    for ($i = 1; $i -le $m; $i++) {
        for ($j = 1; $j -le $n; $j++) {
            $previousI = $i - 1; $previousJ = $j - 1
            $substitutionDistance = [double]$distance[$previousI,$previousJ]
            $relativeDistance = [Math]::Max(0.0, $substitutionDistance - $alignmentBaseline)
            $pairCost = $cost[$previousI,$previousJ] + [Math]::Min($substitutionCap, $relativeDistance)
            $removeCost = $cost[$previousI,$j] + $gapCost
            $addCost = $cost[$i,$previousJ] + $gapCost
            # Prefer a pair on ties so a wholly rewritten page remains a modified
            # page rather than becoming a noisy remove+add pair.
            if ($pairCost -le $removeCost -and $pairCost -le $addCost) { $cost[$i,$j] = $pairCost; $choice[$i,$j] = 1 }
            elseif ($removeCost -le $addCost) { $cost[$i,$j] = $removeCost; $choice[$i,$j] = 2 }
            else { $cost[$i,$j] = $addCost; $choice[$i,$j] = 3 }
        }
    }
    $reverse = @(); $i = $m; $j = $n
    while ($i -gt 0 -or $j -gt 0) {
        $step = [int]$choice[$i,$j]
        if ($step -eq 1) {
            $beforeIndex = $i - 1; $afterIndex = $j - 1
            $fullExact = $beforeFullHashes[$beforeIndex] -and $beforeFullHashes[$beforeIndex] -eq $afterFullHashes[$afterIndex]
            $contentExact = $beforeHashes[$beforeIndex] -and $beforeHashes[$beforeIndex] -eq $afterHashes[$afterIndex]
            $declaredUnchanged = $beforeIndex -eq $afterIndex -and $UnchangedPageSet.ContainsKey($beforeIndex + 1)
            $pairDistance = [double]$distance[$beforeIndex,$afterIndex]
            $pairRelativeDistance = [Math]::Max(0.0, $pairDistance - $alignmentBaseline)
            $notMutualNearest = $pairDistance -gt ([double]$rowMinimums[$beforeIndex] + 0.00005) -or $pairDistance -gt ([double]$columnMinimums[$afterIndex] + 0.00005)
            $ambiguousSubstitution = -not $fullExact -and -not $contentExact -and -not $declaredUnchanged -and $notMutualNearest -and $pairRelativeDistance -ge ($substitutionCap * 0.98)
            $pairKind = if ($fullExact -or $declaredUnchanged) { 'unchanged' } elseif ($ItemKind -eq 'unknown' -or $ambiguousSubstitution) { 'unknown' } else { 'modified' }
            $method = if ($fullExact) { 'exact-raster-sequence' } elseif ($declaredUnchanged) { 'declared-unchanged' } elseif ($contentExact) { 'content-raster-sequence' } elseif ($ambiguousSubstitution) { 'perceptual-raster-sequence-ambiguous' } else { 'perceptual-raster-sequence' }
            $mappingMessage = if ($ambiguousSubstitution) { 'ページの全面書き換えと追加・削除を画像だけでは区別できないため、対応の確認が必要です。' } else { '' }
            $reverse += New-DiffPageMapping $beforeIndex $afterIndex $pairKind $method $ambiguousSubstitution $mappingMessage
            $i--; $j--
        } elseif ($step -eq 2) {
            $beforeIndex = $i - 1; $hash = [string]$beforeFullHashes[$beforeIndex]
            $beforeDuplicates = @($beforeFullHashes | Where-Object { $_ -eq $hash }).Count
            $afterDuplicates = @($afterFullHashes | Where-Object { $_ -eq $hash }).Count
            $duplicateAmbiguous = $hash -and $beforeDuplicates -gt 1 -and $afterDuplicates -gt 0
            $method = if ($duplicateAmbiguous) { 'sequence-removed-ambiguous-duplicate' } else { 'sequence-removed' }
            $mappingMessage = if ($duplicateAmbiguous) { '同一内容のページが複数あるため、削除された複製の位置は特定できません。' } else { '' }
            $reverse += New-DiffPageMapping $beforeIndex -1 'removed' $method $duplicateAmbiguous $mappingMessage; $i--
        } else {
            $afterIndex = $j - 1; $hash = [string]$afterFullHashes[$afterIndex]
            $beforeDuplicates = @($beforeFullHashes | Where-Object { $_ -eq $hash }).Count
            $afterDuplicates = @($afterFullHashes | Where-Object { $_ -eq $hash }).Count
            $duplicateAmbiguous = $hash -and $afterDuplicates -gt 1 -and $beforeDuplicates -gt 0
            $method = if ($duplicateAmbiguous) { 'sequence-inserted-ambiguous-duplicate' } else { 'sequence-inserted' }
            $mappingMessage = if ($duplicateAmbiguous) { '同一内容のページが複数あるため、追加された複製の位置は特定できません。' } else { '' }
            $reverse += New-DiffPageMapping -1 $afterIndex 'added' $method $duplicateAmbiguous $mappingMessage; $j--
        }
    }
    [array]::Reverse($reverse)
    return @($reverse)
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
        $cachedPagesById[$id] = [ordered]@{
            before = @($beforeCached)
            after = @($afterCached)
            beforePersistent = ($beforeCached.Count -gt 0)
            afterPersistent = ($afterCached.Count -gt 0)
        }
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
            $beforePersistent = [bool]$cached.beforePersistent
            $afterPersistent = [bool]$cached.afterPersistent
            $beforePages = @($cached.before)
            $afterPages = @($cached.after)
            if ($beforePages.Count -eq 0) { $beforePages = @(Get-RasterPages $itemRasterDir 'before') }
            if ($afterPages.Count -eq 0) { $afterPages = @(Get-RasterPages $itemRasterDir 'after') }
            $pageMappings = @(Get-DiffPageMappings $beforePages $afterPages $kind $unchangedPageSet)
            if ($pageMappings.Count -eq 0) { throw '比較できるPDFページがありません。' }
            $outputDirectory = [IO.Path]::GetFullPath([string]$item.outputDirectory)
            if (-not (Test-Path -LiteralPath $outputDirectory)) { New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null }
            $pagesById[$id] = New-Object System.Collections.ArrayList
            $pageMappingByNumber = @{}
            for ($i = 0; $i -lt $pageMappings.Count; $i++) {
                $mapping = $pageMappings[$i]
                $beforeIndex = [int]$mapping.beforeIndex; $afterIndex = [int]$mapping.afterIndex
                $beforeImage = if ($beforeIndex -ge 0) { [string]$beforePages[$beforeIndex] } else { '' }
                $afterImage = if ($afterIndex -ge 0) { [string]$afterPages[$afterIndex] } else { '' }
                $pageKind = [string]$mapping.kind
                if ($pageKind -eq 'unchanged') { $unchangedPagesSkipped++ }
                $pageNumber = $i + 1
                $pageMappingByNumber[$pageNumber] = $mapping
                $copyBefore = (-not $beforePersistent) -or [string]::IsNullOrWhiteSpace($beforeImage)
                $copyAfter = (-not $afterPersistent) -or [string]::IsNullOrWhiteSpace($afterImage)
                if ($beforePersistent -and -not [string]::IsNullOrWhiteSpace($beforeImage)) { $reusedRasterPageAssets++ }
                if ($afterPersistent -and -not [string]::IsNullOrWhiteSpace($afterImage)) { $reusedRasterPageAssets++ }
                $pageRequests.Add([ReportBinderDiffBatchPageRequest]@{
                    itemId = $id
                    beforePath = $beforeImage
                    afterPath = $afterImage
                    outputDirectory = $outputDirectory
                    pageNumber = $pageNumber
                    kind = $pageKind
                    copyBefore = $copyBefore
                    copyAfter = $copyAfter
                })
            }
            $cachedPagesById[$id].pageMappings = $pageMappingByNumber
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
                $mappingByNumber = $cachedPagesById[$id].pageMappings
                if ($null -ne $mappingByNumber) {
                    $mapping = $mappingByNumber[[int]$pageResult.page.pageNumber]
                    if ($null -ne $mapping) {
                        Add-Member -InputObject $pageResult.page -NotePropertyName beforePageNumber -NotePropertyValue $(if ([int]$mapping.beforeIndex -ge 0) { [int]$mapping.beforeIndex + 1 } else { 0 }) -Force
                        Add-Member -InputObject $pageResult.page -NotePropertyName afterPageNumber -NotePropertyValue $(if ([int]$mapping.afterIndex -ge 0) { [int]$mapping.afterIndex + 1 } else { 0 }) -Force
                        Add-Member -InputObject $pageResult.page -NotePropertyName comparisonKind -NotePropertyValue ([string]$mapping.kind) -Force
                        Add-Member -InputObject $pageResult.page -NotePropertyName matchMethod -NotePropertyValue ([string]$mapping.method) -Force
                        Add-Member -InputObject $pageResult.page -NotePropertyName mappingAmbiguous -NotePropertyValue ([bool]$mapping.ambiguous) -Force
                        Add-Member -InputObject $pageResult.page -NotePropertyName mappingMessage -NotePropertyValue ([string]$mapping.message) -Force
                    }
                }
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
            reusedRasterPageAssets = $reusedRasterPageAssets
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
