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
        [ordered]@{ id='scan-change'; kind='modified'; beforeRasterDirectory=$riskRenderRoots.scanBefore; afterRasterDirectory=$riskRenderRoots.scanChanged; beforePageCount=1; afterPageCount=1; beforePdf=''; afterPdf=''; outputDirectory=(Join-Path $testRoot 'scan-change'); unchangedPageNumbers=@() }
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
    & $node (Join-Path $repoRoot 'tests\pdf-diff-corpus.mjs') $manifestPath $renderRoot
    if ($LASTEXITCODE -ne 0) { throw 'PDF diff corpus quality gate failed.' }
    Write-Output 'PDF diff corpus selfcheck OK.'
} finally {
    if (-not $KeepTestData -and (Test-Path -LiteralPath $testRoot)) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
