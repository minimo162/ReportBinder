param()

# 提出用PDFの出力先パスの重なり（H-1）と、共有用フォルダーへのコピー
# （/api/v2/outputs/publish-all の本体 Publish-DocumentPackAllPdfsToShared）を、
# server.ps1 の実物の関数で確かめる。Java・Office は使わない。組版の代わりに、
# 出力済みのPDFと出力状態を直接置き、入力スナップショットだけを差し替える。
# あわせて、ページ構成の保存（Reorder-Pages）が不完全な並びを拒むこと（H-2）も見る。
$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-output-publish-' + [Guid]::NewGuid().ToString('N'))
$oldConfigRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', 'Process')
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $testRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_OUTPUT_TEST_APPROOT')
    $testBody = @'
function Assert-OutputTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function Assert-Throws([scriptblock]$Action, [string]$ExpectedFragment, [string]$Message) {
    $caught = $null
    try { & $Action } catch { $caught = $_.Exception.Message }
    if ($null -eq $caught) { throw "$Message (no error was raised)" }
    if (-not [string]::IsNullOrWhiteSpace($ExpectedFragment) -and -not $caught.Contains($ExpectedFragment)) { throw "$Message (unexpected error: $caught)" }
}

$submissionDir = Join-Path $Script:LocalConfigRoot 'submission'
$dataDir = Join-Path $Script:LocalConfigRoot 'data'
$outputDir = Join-Path $Script:LocalConfigRoot 'output'
foreach ($dir in @($submissionDir,$dataDir,$outputDir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
Save-AppConfig ([pscustomobject][ordered]@{
    schemaVersion = 2; lastSubmissionDir = $submissionDir; lastDataDir = $dataDir; lastOutputDir = $outputDir
})
$paths = [pscustomobject][ordered]@{ submissionDir=$submissionDir; dataDir=$dataDir; outputDir=$outputDir }
Ensure-Package $paths -Languages @('ja')

# 一式に原稿1件と、指定した出力先へページを置く。PDFの中身は使わない。
function Add-TestPages([string]$PackId, [string[]]$Volumes) {
    [void](Update-StructureLocked 'ja' {
        param($st)
        $workbookId = 'wb_' + [Guid]::NewGuid().ToString('N').Substring(0,12)
        $wb = [pscustomobject][ordered]@{ workbookId=$workbookId; packId=$PackId; fileName="$workbookId.xlsx"; relativePath="$workbookId.xlsx"; displayName=$workbookId; status='rendered' }
        Set-NoteProperty $st 'workbooks' (@(Get-Array $st.workbooks) + @($wb))
        $pages = @(Get-Array $st.pages)
        $order = 10
        foreach ($volume in $Volumes) {
            $pages += [pscustomobject][ordered]@{ pageId="${workbookId}_$order"; workbookId=$workbookId; sheetName="Sheet$order"; title="Sheet$order"; volume=$volume; enabled=$true; order=$order; numberingMode='visible' }
            $order += 10
        }
        Set-NoteProperty $st 'pages' @($pages)
        return $null
    })
}
# 保存時の検査を通らない古い型を、保存済みのデータとして直接置く。
function Set-LegacyPattern([string]$PackId, [string]$Pattern) {
    [void](Update-StructureLocked 'ja' {
        param($st)
        $pack = Get-PackRecord $st $PackId
        $settings = Get-ResolvedPackSettings $pack 'ja'
        $settings.outputFileNamePattern = $Pattern
        Set-NoteProperty $pack 'settings' $settings
        return $null
    })
}
function Get-PlannedOutputPath([string]$PackId, [string]$Volume) {
    $st = Get-Structure 'ja'
    $scope = Resolve-DocumentPackScope $st $PackId $true
    return [IO.Path]::GetFullPath((Join-Path $outputDir ([string](Get-PackTargetOutputNaming $st 'ja' $scope $Volume).outputFileName)))
}

# ---- H-1: 出力先パスの重なり -------------------------------------------------
$sharedName = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='共有名の一式' })
$sharedId = [string]$sharedName.packId
Add-TestPages $sharedId @('ja-main','ja-appendix')
Assert-OutputTest (Test-PackVolumeHasEnabledPages (Get-Structure 'ja') $sharedId 'ja-appendix') 'Test pages were not stored on the appendix target.'
# 既定の型（{targetName} 入り）なら本体と補足は別の名前になる。
Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' $sharedId @('ja-main','ja-appendix')
Set-LegacyPattern $sharedId '{packName}_{yyyyMMdd}.pdf'
Assert-Throws { Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' $sharedId @('ja-main') } '{targetName}' 'Building 本体 to the same path as 補足 was not rejected.'
Assert-Throws { Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' $sharedId @('ja-appendix') } '{targetName}' 'Building 補足 to the same path as 本体 was not rejected.'

# 補足にページが無ければ補足は書き出されないので、古い型でも本体だけは出力できる。
$mainOnly = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='本体だけの一式' })
$mainOnlyId = [string]$mainOnly.packId
Add-TestPages $mainOnlyId @('ja-main')
Set-LegacyPattern $mainOnlyId '{packName}_{yyyyMMdd}.pdf'
Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' $mainOnlyId @('ja-main')

# 他の一式が記録している出力PDFと同じパスにも書かない（{packName} を含まない型どうし）。
$packA = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='一式A'; settings=[pscustomobject]@{ outputFileNamePattern='{targetName}.pdf' } })
$packB = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='一式B'; settings=[pscustomobject]@{ outputFileNamePattern='{targetName}.pdf' } })
$packAPath = Get-PlannedOutputPath ([string]$packA.packId) 'ja-main'
[void](Update-StructureLocked 'ja' {
    param($st)
    $state = Get-PackOutputState $st 'ja' ([string]$packA.packId) 'ja-main' $true
    Set-NoteProperty $state 'outputPdf' $packAPath
    Set-NoteProperty $state 'builtFingerprint' 'fp'
    return $null
})
Assert-Throws { Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' ([string]$packB.packId) @('ja-main') } '一式「一式A」' 'Overwriting the output PDF recorded by another pack was not rejected.'
# 自分が前回書いたPDFへの再出力は重なりではない。
Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' ([string]$packA.packId) @('ja-main')
# 大文字小文字だけが違うパスも同じファイルとして扱う。
[void](Update-StructureLocked 'ja' {
    param($st)
    $state = Get-PackOutputState $st 'ja' ([string]$packA.packId) 'ja-main' $true
    Set-NoteProperty $state 'outputPdf' $packAPath.ToUpperInvariant()
    return $null
})
Assert-Throws { Assert-PackOutputPathsDistinct (Get-Structure 'ja') 'ja' ([string]$packB.packId) @('ja-main') } '一式「一式A」' 'A recorded output path differing only by case was not treated as the same file.'

# 組み込み一式の状態は volumes と outputs の両方に写しがある。自分の写しとは重ならない。
$ecmMainPath = Get-PlannedOutputPath 'pack_ecm' 'ja-main'
[void](Update-StructureLocked 'ja' {
    param($st)
    $state = Get-PackOutputState $st 'ja' 'pack_ecm' 'ja-main' $true
    Set-NoteProperty $state 'outputPdf' $ecmMainPath
    return $null
})
$ecmStructure = Get-Structure 'ja'
$ecmMirror = Get-DataProperty (Get-DataProperty $ecmStructure 'outputs' $null) 'pack_ecm|main' $null
Assert-OutputTest ($null -ne $ecmMirror -and [string](Get-DataProperty $ecmMirror 'outputPdf' '') -eq $ecmMainPath) 'The built-in output mirror was not populated; the self-exclusion case is not exercised.'
Assert-PackOutputPathsDistinct $ecmStructure 'ja' 'pack_ecm' @('ja-main')

# ---- H-2: ページ構成の保存は、送られたレーンのページを欠けなく1回ずつ含む --------
$layoutPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='並べ替えの一式' })
$layoutId = [string]$layoutPack.packId
Add-TestPages $layoutId @('ja-main','ja-main','ja-appendix','none')
function Get-LayoutIds([string]$Volume) {
    $st = Get-Structure 'ja'
    $ids = @{}
    foreach ($wb in @(Get-Array $st.workbooks | Where-Object { Test-WorkbookPack $_ $layoutId })) { $ids[[string]$wb.workbookId] = $true }
    return @(Get-Array $st.pages | Where-Object { $ids.ContainsKey([string]$_.workbookId) -and [string]$_.volume -eq $Volume } | Sort-Object {[double]$_.order} | ForEach-Object { Resolve-PageId $_ })
}
function Get-LayoutSignature { return (@('ja-main','ja-appendix','none') | ForEach-Object { $_ + '=' + ((Get-LayoutIds $_) -join ',') }) -join ';' }
$main = @(Get-LayoutIds 'ja-main'); $appendix = @(Get-LayoutIds 'ja-appendix'); $none = @(Get-LayoutIds 'none')
Assert-OutputTest ($main.Count -eq 2 -and $appendix.Count -eq 1 -and $none.Count -eq 1) 'Layout test pages were not placed as expected.'
$before = Get-LayoutSignature
function New-Layout($Volumes) { return [pscustomobject][ordered]@{ packId=$layoutId; volumes=[pscustomobject]$Volumes } }
Assert-Throws { Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@(); 'ja-appendix'=$appendix; none=$none })) } '不完全' 'An empty lane that still has pages was accepted.'
Assert-Throws { Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@($main[1]); 'ja-appendix'=$appendix; none=$none })) } '不完全' 'A lane missing one of its pages was accepted.'
Assert-Throws { Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@($main + @('page_unknown')); 'ja-appendix'=$appendix; none=$none })) } 'この一式に無いページ' 'An unknown page id was silently ignored.'
Assert-Throws { Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=$main; 'ja-appendix'=@($appendix + @($main[0])); none=$none })) } '2回' 'A page sent in two lanes was accepted.'
Assert-OutputTest ((Get-LayoutSignature) -eq $before) 'A rejected layout save still changed the page composition.'
# 正しい全体の並べ替えと、送ったレーンの間だけの移動（補足から本体へ）は通る。
[void](Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@($main[1],$main[0]); 'ja-appendix'=$appendix; none=$none })))
Assert-OutputTest (((Get-LayoutIds 'ja-main') -join ',') -eq (@($main[1],$main[0]) -join ',')) 'A complete reorder was not applied.'
[void](Reorder-Pages 'ja' (New-Layout ([ordered]@{ 'ja-main'=@($main[1],$main[0],$appendix[0]) })))
Assert-OutputTest (((Get-LayoutIds 'ja-main') -join ',') -eq (@($main[1],$main[0],$appendix[0]) -join ',') -and (Get-LayoutIds 'ja-appendix').Count -eq 0) 'Moving a page from an unsent lane into a sent lane was rejected or misapplied.'

Write-Output 'output-publish selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
