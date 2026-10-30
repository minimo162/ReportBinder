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
Assert-OutputTest (((Get-LayoutIds 'ja-main') -join ',') -eq (@($main[1],$main[0],$appendix[0]) -join ',') -and @(Get-LayoutIds 'ja-appendix').Count -eq 0) 'Moving a page from an unsent lane into a sent lane was rejected or misapplied.'

# ---- M-1: 日付が変わっただけでは「再出力が必要」にしない ----------------------------
# ここではまだ実物の Get-FinalBuildInputSnapshot を使う（下で差し替える）。
$dated = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='日付の一式' })
Add-TestPages ([string]$dated.packId) @('ja-main')
$today = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' ([string]$dated.packId)
$realGetDate = Get-Command Get-Date -CommandType Cmdlet
function Get-Date { param([string]$Format) $tomorrow = (& $realGetDate).AddDays(1); if ($Format) { return $tomorrow.ToString($Format) } return $tomorrow }
try { $tomorrow = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' ([string]$dated.packId) }
finally { Remove-Item Function:\Get-Date }
Assert-OutputTest ([string]$today.outputFileName -ne [string]$tomorrow.outputFileName) 'The date override did not reach the output file name; the M-1 case is not exercised.'
Assert-OutputTest ([string]$today.fingerprint -eq [string]$tomorrow.fingerprint) 'A date change alone changed the output input fingerprint.'

# ---- 共有用フォルダーへのまとめてコピー（publish-all）と出力PDFの照合（M-2） ----------
# 原稿のPDFを置かずに済むよう、入力スナップショットだけを差し替える。指紋は出力先ごとに固定。
function Get-FinalBuildInputSnapshot($Structure, [string]$Language, [string]$Volume, [string]$PackId) {
    $ids = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks | Where-Object { Test-WorkbookPack $_ $PackId })) { $ids[[string]$wb.workbookId] = $true }
    $count = @(Get-Array $Structure.pages | Where-Object { $ids.ContainsKey([string]$_.workbookId) -and [string]$_.volume -eq $Volume -and $_.enabled -ne $false }).Count
    return [pscustomobject]@{ fingerprint="fp|$PackId|$Volume"; blockers=@(); pageCount=$count }
}
# 出力ジョブが確定したときと同じ記録を置く。
function Set-BuiltOutput([string]$PackId, [string]$Volume, [string]$Content, [bool]$RecordIntegrity = $true) {
    $path = Get-PlannedOutputPath $PackId $Volume
    [IO.File]::WriteAllText($path, $Content)
    [void](Update-StructureLocked 'ja' {
        param($st)
        $state = Get-PackOutputState $st 'ja' $PackId $Volume $true
        Set-NoteProperty $state 'outputPdf' $path
        Set-NoteProperty $state 'builtFingerprint' ([string](Get-FinalBuildInputSnapshot $st 'ja' $Volume $PackId).fingerprint)
        Set-NoteProperty $state 'outputIntegrity' $(if ($RecordIntegrity) { Get-OutputPdfIntegrity $path } else { $null })
        return $null
    })
    return $path
}
function Get-SharedEntries { return @(Get-ChildItem -LiteralPath $submissionDir -Force | ForEach-Object { $_.Name }) }
$publishPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='発行の一式' })
$publishId = [string]$publishPack.packId
Add-TestPages $publishId @('ja-main','ja-appendix')
$mainPdf = Set-BuiltOutput $publishId 'ja-main' '%PDF-main'
$appendixPdf = Set-BuiltOutput $publishId 'ja-appendix' '%PDF-appendix'
$published = Publish-DocumentPackAllPdfsToShared 'ja' $publishId
Assert-OutputTest (@($published.files).Count -eq 2) 'publish-all did not copy both targets with pages.'
foreach ($file in @($published.files)) {
    $local = if ([string]$file.targetId -eq 'main') { $mainPdf } else { $appendixPdf }
    Assert-OutputTest ((New-Sha256 ([string]$file.sharedPath)) -eq (New-Sha256 $local)) "publish-all copied a different file for $($file.targetId)."
    Assert-OutputTest ([IO.Path]::GetDirectoryName([string]$file.sharedPath) -eq [string]$published.sharedFolder) 'publish-all split the targets across folders.'
}
Assert-OutputTest (@(Get-SharedEntries | Where-Object { $_ -like '.publishing-*' }).Count -eq 0) 'A staging folder was left behind after publishing.'
# 発行フォルダー名は「日時_一式名」だけ。誰が発行したかは publishedBy に残す。
$publisher = Get-SafePublishUserName
Assert-OutputTest ([string]$published.folderName -match ('^\d{4}_\d{6}_' + [regex]::Escape((Get-SafePublishPackName '発行の一式')) + '$')) "The published folder name is not 'MMdd_HHmmss_<pack>': $($published.folderName)"
Assert-OutputTest (-not ([string]$published.folderName).EndsWith('_' + $publisher)) 'The published folder name still contains the Windows user name.'
Assert-OutputTest ([IO.Path]::GetFileName([string]$published.sharedFolder) -eq [string]$published.folderName) 'folderName does not match the published folder.'
Assert-OutputTest ([string]$published.publishedBy -eq $publisher) 'publishedBy no longer records the Windows user name.'
$sharedBefore = @(Get-SharedEntries)

# 1冊でも最新でなければ何もコピーしない。
[void](Update-StructureLocked 'ja' { param($st) Set-NoteProperty (Get-PackOutputState $st 'ja' $publishId 'ja-appendix' $true) 'builtFingerprint' 'old'; return $null })
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '最新ではありません' 'publish-all copied while one target was stale.'
Assert-OutputTest ((@(Get-SharedEntries) -join '|') -eq ($sharedBefore -join '|')) 'A rejected publish-all still created something in the shared folder.'
$appendixPdf = Set-BuiltOutput $publishId 'ja-appendix' '%PDF-appendix'

# 出力した後に書き換えられたPDFは「最新」と表示せず、発行もしない。
[IO.File]::AppendAllText($mainPdf, 'edited')
$changed = Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $publishId
Assert-OutputTest ([string]$changed.displayState -eq 'needs-rebuild' -and @($changed.staleReasons | Where-Object { [string]$_.type -eq 'output-changed' }).Count -eq 1) 'A PDF changed after output was still shown as latest.'
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '変更または上書き' 'A PDF changed after output was published.'
# サイズも更新日時も同じまま中身だけ違うものは、発行前のハッシュで止める。
$mainPdf = Set-BuiltOutput $publishId 'ja-main' '%PDF-main'
$stamp = (Get-Item -LiteralPath $mainPdf).LastWriteTimeUtc
[IO.File]::WriteAllText($mainPdf, '%PDF-MAIN')
(Get-Item -LiteralPath $mainPdf).LastWriteTimeUtc = $stamp
Assert-OutputTest ([string](Get-FinalBuildReadiness (Get-Structure 'ja') 'ja' 'ja-main' $publishId).displayState -eq 'built') 'The same-size case did not reach the hash check.'
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '中身' 'A PDF whose content differs from the recorded hash was published.'
# 照合の記録が無い（以前の版で出力した）PDFは発行しない。
$mainPdf = Set-BuiltOutput $publishId 'ja-main' '%PDF-main' $false
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '記録がありません' 'A PDF without an output record was published.'
Assert-OutputTest ((@(Get-SharedEntries) -join '|') -eq ($sharedBefore -join '|')) 'A rejected publish still created something in the shared folder.'
$mainPdf = Set-BuiltOutput $publishId 'ja-main' '%PDF-main'

# ページの無い出力先は対象外（補足が空なら本体だけを発行する）。
$mainOnlyPublish = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='本体だけの発行' })
Add-TestPages ([string]$mainOnlyPublish.packId) @('ja-main')
[void](Set-BuiltOutput ([string]$mainOnlyPublish.packId) 'ja-main' '%PDF-only')
$onlyMain = Publish-DocumentPackAllPdfsToShared 'ja' ([string]$mainOnlyPublish.packId)
Assert-OutputTest (@($onlyMain.files).Count -eq 1 -and [string]@($onlyMain.files)[0].targetId -eq 'main') 'publish-all did not skip the target without pages.'

# ---- M-3: 出力中は共有用フォルダーへコピーしない ------------------------------------
$sharedBefore = @(Get-SharedEntries)
# 出力が同じ出力先のPDFを置き換えている間（出力のロックを持っている間）はコピーしない。
$outputLock = Join-Path (Get-WorkspacePath 'ja') ("locks\pack-output_{0}_main.lock" -f $publishId)
New-Item -ItemType Directory -Path (Split-Path -Parent $outputLock) -Force | Out-Null
$held = [IO.File]::Open($outputLock, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
try { Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } 'ロック中' 'Publishing ran while the output lock was held.' }
finally { $held.Dispose() }
# 出力ジョブが動いている間はコピーしない。
$jobId = 'final_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '_' + [Guid]::NewGuid().ToString('N').Substring(0,8)
$jobStatus = Join-Path (Get-RenderJobDir 'ja') "$jobId.status.json"
Write-JsonFile $jobStatus ([ordered]@{ ok=$true; jobId=$jobId; kind='final-build'; status='running'; processId=$PID; startedAt=(New-NowIso); updatedAt=(New-NowIso) })
try { Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '出力中' 'Publishing ran while an output job was active.' }
finally { Remove-Item -LiteralPath $jobStatus -Force }
Assert-OutputTest ((@(Get-SharedEntries) -join '|') -eq ($sharedBefore -join '|')) 'A publish refused during output still created something in the shared folder.'
$afterJob = Publish-DocumentPackAllPdfsToShared 'ja' $publishId
Assert-OutputTest (@($afterJob.files).Count -eq 2) 'Publishing did not resume after the output finished.'

# ---- M-7: 発行の作業用フォルダーは隠し、残骸は次の発行の前に消す -----------------
$me = Get-SafePublishUserName
function New-StagingLeftover([string]$Name, [double]$HoursOld) {
    $dir = New-Item -ItemType Directory -Path (Join-Path $submissionDir $Name) -Force
    [IO.File]::WriteAllText((Join-Path $dir.FullName 'half.pdf'), '%PDF-half')
    $dir.LastWriteTimeUtc = (Get-Date).ToUniversalTime().AddHours(-$HoursOld)
    return $Name
}
$guid = { [Guid]::NewGuid().ToString('N') }
$ownOld = New-StagingLeftover (".publishing-$me-" + (& $guid)) 2
$ownRecent = New-StagingLeftover (".publishing-$me-" + (& $guid)) 0
$otherOld = New-StagingLeftover (".publishing-$me-x-" + (& $guid)) 48
$legacyOld = New-StagingLeftover ('.publishing-' + (& $guid)) 48
$legacyRecent = New-StagingLeftover ('.publishing-' + (& $guid)) 2
# コピーの瞬間に作業用フォルダーの属性を記録する。
$Script:StagingAttributes = @()
function Copy-Item {
    param([string]$LiteralPath, [string]$Destination, [switch]$Force)
    $Script:StagingAttributes += (Get-Item -LiteralPath (Split-Path -Parent $Destination) -Force).Attributes
    Microsoft.PowerShell.Management\Copy-Item -LiteralPath $LiteralPath -Destination $Destination -Force:$Force
}
[void](Publish-DocumentPackAllPdfsToShared 'ja' $publishId)
Remove-Item Function:\Copy-Item
$left = @(Get-SharedEntries)
Assert-OutputTest (-not ($left -contains $ownOld)) 'An old staging folder left by this user was not removed.'
Assert-OutputTest (-not ($left -contains $legacyOld)) 'An old staging folder in the previous naming was not removed.'
Assert-OutputTest ($left -contains $ownRecent -and $left -contains $legacyRecent) 'A staging folder that may still be in use was removed.'
Assert-OutputTest ($left -contains $otherOld) "Another user's staging folder was removed."
Assert-OutputTest ($Script:StagingAttributes.Count -eq 2) 'The staging copy was not observed.'
if ($env:OS -eq 'Windows_NT') {
    Assert-OutputTest (@($Script:StagingAttributes | Where-Object { ($_ -band [IO.FileAttributes]::Hidden) -eq 0 }).Count -eq 0) 'The staging folder was visible in the shared folder while copying.'
}

# ---- 隠し属性は作業用フォルダーだけ。発行フォルダーとPDFには残さない --------------------
# 改名しても属性はフォルダーに残るため、以前は発行フォルダーごと見えなくなっていた。
# 属性そのものは Windows でしか立たないので、どの環境でも「改名後の発行フォルダーと
# 中のPDFを1つずつ見えるようにしたか」を呼び出しで確かめ、Windows では実際の属性も見る。
$realClearHidden = ${function:Clear-PublishHiddenAttribute}
$Script:ClearedPaths = @()
$Script:FailClearOnPublishedFolder = $false
$submissionFull = [IO.Path]::GetFullPath($submissionDir)
function Clear-PublishHiddenAttribute([string]$Path) {
    $full = [IO.Path]::GetFullPath($Path)
    $Script:ClearedPaths += $full
    # 改名した後の発行フォルダー（共有ルート直下で「.」で始まらないもの）だけ失敗させる。
    if ($Script:FailClearOnPublishedFolder -and [IO.Path]::GetDirectoryName($full) -eq $submissionFull -and [IO.Path]::GetFileName($full) -notlike '.*') { return $false }
    return (& $realClearHidden $Path)
}
$onWindows = $env:OS -eq 'Windows_NT'
# Get-Item は -Force が無いと隠しファイルを返さない。照合が隠されたPDFを「見つかりません」に
# しないこと。Windows では隠し属性で、それ以外では「.」で始まる名前で隠す。
$hiddenProbe = Join-Path (Split-Path -Parent $mainPdf) '.hidden-integrity-probe.pdf'
[IO.File]::WriteAllText($hiddenProbe, '%PDF-hidden')
if ($onWindows) { (Get-Item -LiteralPath $hiddenProbe -Force).Attributes = [IO.FileAttributes]::Hidden }
try {
    $hiddenState = [pscustomobject]@{ outputIntegrity = (Get-OutputPdfIntegrity $hiddenProbe) }
    $hiddenProblem = Get-OutputPdfIntegrityProblem $hiddenState $hiddenProbe $true
    Assert-OutputTest ($hiddenProblem -eq '') "A hidden output PDF failed the integrity check: $hiddenProblem"
} finally {
    if ($onWindows) { (Get-Item -LiteralPath $hiddenProbe -Force).Attributes = [IO.FileAttributes]::Normal }
    Remove-Item -LiteralPath $hiddenProbe -Force
}
# 出力したPDFが隠されていても、Copy-Item で写った属性を発行先に残さない。
if ($onWindows) { (Get-Item -LiteralPath $mainPdf -Force).Attributes = (Get-Item -LiteralPath $mainPdf -Force).Attributes -bor [IO.FileAttributes]::Hidden }
try { $visible = Publish-DocumentPackAllPdfsToShared 'ja' $publishId }
finally { if ($onWindows) { (Get-Item -LiteralPath $mainPdf -Force).Attributes = [IO.FileAttributes]::Normal } }
$visibleDir = [IO.Path]::GetFullPath([string]$visible.sharedFolder)
Assert-OutputTest ([IO.Path]::GetFileName($visibleDir) -notlike '.*') 'The published folder kept a staging-style dot name.'
Assert-OutputTest ($Script:ClearedPaths -contains $visibleDir) 'The hidden attribute was not cleared on the published folder after the rename.'
foreach ($file in @($visible.files)) {
    Assert-OutputTest ($Script:ClearedPaths -contains [IO.Path]::GetFullPath([string]$file.sharedPath)) "The hidden attribute was not checked on the published PDF $($file.fileName)."
    Assert-OutputTest (Test-Path -LiteralPath ([string]$file.sharedPath) -PathType Leaf) "The published PDF $($file.fileName) is missing."
}
if ($onWindows) {
    $hiddenItems = @(@(Get-Item -LiteralPath $visibleDir -Force) + @(Get-ChildItem -LiteralPath $visibleDir -Force -Recurse) | Where-Object { ($_.Attributes -band [IO.FileAttributes]::Hidden) -ne 0 })
    Assert-OutputTest ($hiddenItems.Count -eq 0) ('The published folder or its PDFs are hidden: ' + (($hiddenItems | ForEach-Object { $_.FullName }) -join ', '))
    Assert-OutputTest (@(Get-ChildItem -LiteralPath $submissionDir -Directory | Where-Object { $_.FullName -eq $visibleDir }).Count -eq 1) 'The published folder is not listed without -Force.'
}
# 隠し属性を外せなければ、見えない発行フォルダーを残さずに失敗にする。
$sharedBefore = @(Get-SharedEntries)
$Script:FailClearOnPublishedFolder = $true
Assert-Throws { Publish-DocumentPackAllPdfsToShared 'ja' $publishId } '隠し属性を外せません' 'A published folder whose hidden attribute could not be cleared was reported as success.'
Assert-OutputTest ((@(Get-SharedEntries) -join '|') -eq ($sharedBefore -join '|')) 'A published folder that stayed hidden was left in the shared folder.'
$Script:FailClearOnPublishedFolder = $false
Set-Item Function:\Clear-PublishHiddenAttribute $realClearHidden

Write-Output 'output-publish selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
