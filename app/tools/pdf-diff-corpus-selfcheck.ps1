param(
    [string]$TestRoot = '',
    [switch]$KeepTestData
)

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$repoRoot = Split-Path -Parent $appRoot
$testRoot = if ([string]::IsNullOrWhiteSpace($TestRoot)) { Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-pdf-diff-corpus-' + [Guid]::NewGuid().ToString('N')) } else { [IO.Path]::GetFullPath($TestRoot) }
$corpusRoot = Join-Path $testRoot 'corpus'
$renderRoot = Join-Path $testRoot 'rendered'
$fontCacheRoot = Join-Path $testRoot 'font-cache'

try {
    New-Item -ItemType Directory -Path $corpusRoot,$renderRoot,$fontCacheRoot -Force | Out-Null
    $python = (Get-Command python.exe -ErrorAction Stop).Source
    $node = (Get-Command node.exe -ErrorAction Stop).Source
    $java = Join-Path $appRoot 'lib\java\bin\java.exe'
    $pdfbox = Join-Path $appRoot 'lib\pdfbox\pdfbox-app.jar'
    & $python (Join-Path $toolsRoot 'create-pdf-diff-corpus.py') --output $corpusRoot
    if ($LASTEXITCODE -ne 0) { throw 'PDF diff corpus generation failed.' }
    $manifestPath = Join-Path $corpusRoot 'manifest.json'
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($case in @($manifest.cases)) {
        foreach ($property in @('before','after')) {
            $pdfName = [string]$case.$property
            $pdfPath = Join-Path $corpusRoot $pdfName
            $prefix = Join-Path $renderRoot (([IO.Path]::GetFileNameWithoutExtension($pdfName)) + '-')
            & $java ("-Dpdfbox.fontcache={0}" -f $fontCacheRoot) -jar $pdfbox PDFToImage -format bmp -color rgb -dpi ([int]$manifest.dpi) -page ([int]$case.page) -prefix $prefix $pdfPath
            if ($LASTEXITCODE -ne 0) { throw "PDF corpus rendering failed: $pdfName" }
        }
    }
    $alignmentRenderRoot = Join-Path $renderRoot 'page-alignment'
    $alignmentBeforeRoot = Join-Path $alignmentRenderRoot 'before'
    $alignmentAfterRoot = Join-Path $alignmentRenderRoot 'after'
    New-Item -ItemType Directory -Path $alignmentBeforeRoot,$alignmentAfterRoot -Force | Out-Null
    foreach ($side in @('before','after')) {
        $pdfName = [string]$manifest.pageAlignment.$side
        $pdfPath = Join-Path $corpusRoot $pdfName
        $sideRoot = if ($side -eq 'before') { $alignmentBeforeRoot } else { $alignmentAfterRoot }
        $prefix = Join-Path $sideRoot 'page-'
        & $java ("-Dpdfbox.fontcache={0}" -f $fontCacheRoot) -jar $pdfbox PDFToImage -format png -color rgb -dpi ([int]$manifest.dpi) -prefix $prefix $pdfPath
        if ($LASTEXITCODE -ne 0) { throw "PDF page-alignment rendering failed: $pdfName" }
    }
    $riskRenderRoots = @{}
    foreach ($property in @('redesignBefore','redesignAfter','redesignAmbiguous','duplicateBefore','duplicateAfter','scanBefore','scanNoise','scanChanged')) {
        $pdfName = [string]$manifest.riskAlignment.$property
        $pdfPath = Join-Path $corpusRoot $pdfName
        $sideRoot = Join-Path $alignmentRenderRoot $property
        New-Item -ItemType Directory -Path $sideRoot -Force | Out-Null
        $prefix = Join-Path $sideRoot 'page-'
        & $java ("-Dpdfbox.fontcache={0}" -f $fontCacheRoot) -jar $pdfbox PDFToImage -format png -color rgb -dpi ([int]$manifest.dpi) -prefix $prefix $pdfPath
        if ($LASTEXITCODE -ne 0) { throw "PDF risk-alignment rendering failed: $pdfName" }
        $riskRenderRoots[$property] = $sideRoot
    }
    $alignmentRequestPath = Join-Path $testRoot 'page-alignment-request.json'
    $alignmentRequest = [ordered]@{ items = @(
        [ordered]@{ id='forward'; kind='modified'; beforeRasterDirectory=$alignmentBeforeRoot; afterRasterDirectory=$alignmentAfterRoot; beforePageCount=3; afterPageCount=4; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'alignment-forward'); unchangedPageNumbers=@() },
        [ordered]@{ id='reverse'; kind='modified'; beforeRasterDirectory=$alignmentAfterRoot; afterRasterDirectory=$alignmentBeforeRoot; beforePageCount=4; afterPageCount=3; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'alignment-reverse'); unchangedPageNumbers=@() },
        [ordered]@{ id='redesign-forward'; kind='modified'; beforeRasterDirectory=$riskRenderRoots.redesignBefore; afterRasterDirectory=$riskRenderRoots.redesignAfter; beforePageCount=4; afterPageCount=5; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'redesign-forward'); unchangedPageNumbers=@() },
        [ordered]@{ id='redesign-reverse'; kind='modified'; beforeRasterDirectory=$riskRenderRoots.redesignAfter; afterRasterDirectory=$riskRenderRoots.redesignBefore; beforePageCount=5; afterPageCount=4; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'redesign-reverse'); unchangedPageNumbers=@() },
        [ordered]@{ id='redesign-ambiguous'; kind='modified'; beforeRasterDirectory=$riskRenderRoots.redesignBefore; afterRasterDirectory=$riskRenderRoots.redesignAmbiguous; beforePageCount=4; afterPageCount=4; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'redesign-ambiguous'); unchangedPageNumbers=@() },
        [ordered]@{ id='duplicate-forward'; kind='modified'; beforeRasterDirectory=$riskRenderRoots.duplicateBefore; afterRasterDirectory=$riskRenderRoots.duplicateAfter; beforePageCount=4; afterPageCount=5; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'duplicate-forward'); unchangedPageNumbers=@() },
        [ordered]@{ id='duplicate-reverse'; kind='modified'; beforeRasterDirectory=$riskRenderRoots.duplicateAfter; afterRasterDirectory=$riskRenderRoots.duplicateBefore; beforePageCount=5; afterPageCount=4; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'duplicate-reverse'); unchangedPageNumbers=@() },
        [ordered]@{ id='scan-noise'; kind='modified'; beforeRasterDirectory=$riskRenderRoots.scanBefore; afterRasterDirectory=$riskRenderRoots.scanNoise; beforePageCount=1; afterPageCount=1; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'scan-noise'); unchangedPageNumbers=@() },
        [ordered]@{ id='scan-change'; kind='modified'; beforeRasterDirectory=$riskRenderRoots.scanBefore; afterRasterDirectory=$riskRenderRoots.scanChanged; beforePageCount=1; afterPageCount=1; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'scan-change'); unchangedPageNumbers=@() },
        # 変更点PDFの呼び方: 比べる組（pagePairs）を指定する。範囲外・重複の組は捨てる。
        [ordered]@{ id='requested-pairs'; kind='modified'; beforeRasterDirectory=$alignmentBeforeRoot; afterRasterDirectory=$alignmentAfterRoot; beforePageCount=3; afterPageCount=4; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'requested-pairs'); unchangedPageNumbers=@()
            pagePairs=@([ordered]@{ beforePageNumber=3; afterPageNumber=4 }, [ordered]@{ beforePageNumber=0; afterPageNumber=1 }, [ordered]@{ beforePageNumber=3; afterPageNumber=4 }, [ordered]@{ beforePageNumber=9; afterPageNumber=9 }, [ordered]@{ beforePageNumber=1; afterPageNumber=1 }) }
    ) }
    [IO.File]::WriteAllText($alignmentRequestPath, ($alignmentRequest | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
    $alignmentRaw = @(& (Join-Path $toolsRoot 'diff-image-batch.ps1') -RequestPath $alignmentRequestPath -Dpi ([int]$manifest.dpi) -Threshold 24 -MinimumRegionPixels 24 -Padding 5 | ForEach-Object { [string]$_ })
    if ($LASTEXITCODE -ne 0 -or $alignmentRaw.Count -eq 0) { throw 'PDF page-alignment diff generation failed.' }
    $alignmentResult = $alignmentRaw[-1] | ConvertFrom-Json
    foreach ($direction in @('forward','reverse')) {
        $actualItem = @($alignmentResult.items | Where-Object { [string]$_.id -eq $direction } | Select-Object -First 1)
        if ($actualItem.Count -ne 1 -or -not [bool]$actualItem[0].ok) { throw "PDF page-alignment result failed: $direction" }
        $actual = @($actualItem[0].pages | ForEach-Object { "{0}:{1}" -f ([int]$_.beforePageNumber),([int]$_.afterPageNumber) })
        $expected = @($manifest.pageAlignment.$direction | ForEach-Object { "{0}:{1}" -f ([int]$_[0]),([int]$_[1]) })
        if (($actual -join ',') -ne ($expected -join ',')) {
            throw "PDF page-alignment mismatch ($direction): expected=$($expected -join ',') actual=$($actual -join ',')"
        }
    }
    foreach ($direction in @('redesign-forward','redesign-reverse')) {
        $actualItem = @($alignmentResult.items | Where-Object { [string]$_.id -eq $direction } | Select-Object -First 1)
        $actual = @($actualItem[0].pages | ForEach-Object { "{0}:{1}" -f ([int]$_.beforePageNumber),([int]$_.afterPageNumber) })
        $property = if ($direction -eq 'redesign-forward') { 'redesignForward' } else { 'redesignReverse' }
        $expected = @($manifest.riskAlignment.$property | ForEach-Object { "{0}:{1}" -f ([int]$_[0]),([int]$_[1]) })
        if (($actual -join ',') -ne ($expected -join ',')) { throw "PDF redesign alignment mismatch ($direction): expected=$($expected -join ',') actual=$($actual -join ',')" }
    }
    $ambiguousItem = @($alignmentResult.items | Where-Object { [string]$_.id -eq 'redesign-ambiguous' } | Select-Object -First 1)[0]
    $ambiguousPage = @($ambiguousItem.pages | Where-Object { [int]$_.beforePageNumber -eq 2 -and [int]$_.afterPageNumber -eq 2 } | Select-Object -First 1)[0]
    if ([string]$ambiguousPage.comparisonKind -ne 'unknown' -or -not [bool]$ambiguousPage.mappingAmbiguous) { throw 'PDF redesign replacement must be reported as an ambiguous page mapping.' }
    foreach ($direction in @('duplicate-forward','duplicate-reverse')) {
        $duplicateItem = @($alignmentResult.items | Where-Object { [string]$_.id -eq $direction } | Select-Object -First 1)[0]
        $expectedKind = if ($direction -eq 'duplicate-forward') { 'added' } else { 'removed' }
        $ambiguousDuplicates = @($duplicateItem.pages | Where-Object { [string]$_.comparisonKind -eq $expectedKind -and [bool]$_.mappingAmbiguous })
        if ($ambiguousDuplicates.Count -ne 1) { throw "Duplicate PDF page $expectedKind mapping must retain the count and report positional ambiguity." }
    }
    $scanNoiseItem = @($alignmentResult.items | Where-Object { [string]$_.id -eq 'scan-noise' } | Select-Object -First 1)[0]
    $scanChangeItem = @($alignmentResult.items | Where-Object { [string]$_.id -eq 'scan-change' } | Select-Object -First 1)[0]
    if (@($scanNoiseItem.pages[0].regions).Count -ne 0) { throw 'Sparse scan noise must not produce a diff region.' }
    if (@($scanChangeItem.pages[0].regions).Count -lt 1) { throw 'A real checkbox edit on a noisy scan must remain detectable.' }
    # 組を指定すると、その組だけを指定の順に返し（範囲外・重複は捨てる）、全ページの対応付けをしない。
    # 使う画像は永続の画像フォルダーに書かずに outputDirectory へ複製する。前後が両方ある組の枠は、
    # 全ページを対応付けたときの同じ組の枠と同じ。
    $pairsItem = @($alignmentResult.items | Where-Object { [string]$_.id -eq 'requested-pairs' } | Select-Object -First 1)[0]
    $pairsActual = @($pairsItem.pages | ForEach-Object { "{0}:{1}:{2}:{3}" -f ([int]$_.pageNumber),([int]$_.beforePageNumber),([int]$_.afterPageNumber),([string]$_.comparisonKind) })
    if (-not [bool]$pairsItem.ok -or ($pairsActual -join ',') -ne '1:3:4:modified,2:0:1:added,3:1:1:modified') { throw "Requested page pairs were not returned as asked: $($pairsActual -join ',') $($pairsItem.message)" }
    $pairsDir = Join-Path $testRoot 'requested-pairs'
    foreach ($name in @('0001-b.png','0001-a.png','0002-a.png','0003-b.png','0003-a.png')) {
        if (-not (Test-Path -LiteralPath (Join-Path $pairsDir $name))) { throw "A requested page picture was not copied to the output directory: $name" }
    }
    if (@($pairsItem.pages | Where-Object { [string]$_.beforeFile -eq 'render.png' -or [string]$_.afterFile -eq 'render.png' }).Count -ne 0) { throw 'Requested page pairs must not point at the persistent raster folder.' }
    if (@(Get-ChildItem -LiteralPath $alignmentBeforeRoot,$alignmentAfterRoot -File | Where-Object { $_.Name -notlike 'page-*' }).Count -ne 0) { throw 'Requested page pairs wrote into the persistent raster folder.' }
    $forwardItem = @($alignmentResult.items | Where-Object { [string]$_.id -eq 'forward' } | Select-Object -First 1)[0]
    $comparedPairs = 0
    foreach ($pairPage in @($pairsItem.pages | Where-Object { [int]$_.beforePageNumber -gt 0 -and [int]$_.afterPageNumber -gt 0 })) {
        $aligned = @($forwardItem.pages | Where-Object { [int]$_.beforePageNumber -eq [int]$pairPage.beforePageNumber -and [int]$_.afterPageNumber -eq [int]$pairPage.afterPageNumber } | Select-Object -First 1)
        if ($aligned.Count -eq 0) { continue }
        $comparedPairs++
        $box = { param($Page) @($Page.regions | ForEach-Object { '{0:N5},{1:N5},{2:N5},{3:N5}' -f $_.x,$_.y,$_.width,$_.height }) -join ' ' }
        if ((& $box $pairPage) -ne (& $box $aligned[0])) { throw "A requested page pair found other regions than the aligned comparison: $($pairPage.beforePageNumber):$($pairPage.afterPageNumber)" }
    }
    if ($comparedPairs -lt 2) { throw "Requested page pairs were not compared with the aligned comparison (compared=$comparedPairs)." }
    & $node (Join-Path $repoRoot 'tests\pdf-diff-corpus.mjs') $manifestPath $renderRoot
    if ($LASTEXITCODE -ne 0) { throw 'PDF diff corpus quality gate failed.' }
    Write-Output 'PDF diff corpus selfcheck OK.'
} finally {
    if (-not $KeepTestData -and (Test-Path -LiteralPath $testRoot)) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
