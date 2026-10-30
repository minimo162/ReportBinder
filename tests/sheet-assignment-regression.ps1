param()

# シートの振り分け（出力先・未振り分け・シートの同一性）を server.ps1 の実物の関数で確かめる。
# Excel は使わない。PDF作成でシートを調べた結果（シート名・位置・非表示・対象）を
# Render-Workbook と同じ順で Update-WorkbookPagesFromInspection と
# Remove-SheetSelectionExcludedPages へ渡す。
#
# 守ること:
# - 半角数字だけのPDF化で、PDF化済みの数字シートを改名・非表示にして外したとき、
#   そのページが未振り分けに残らない（旧名のページも残らない）
# - 表示中の全シートでは、改名・非表示でも配置を引き継ぐ（従来どおり）
# - 改名後に旧名のシートが新しく現れても、ページIDが重ならず保存できる
# - 以前の版が未振り分けに残したページを、起動時の読み込みで整理する
$ErrorActionPreference = 'Stop'
$testsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $testsRoot
$appRoot = Join-Path $repoRoot 'app'
$serverPath = Join-Path $appRoot 'server.ps1'
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-sheet-assignment-' + [Guid]::NewGuid().ToString('N'))
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
function Assert-SheetTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
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

function New-TestWorkbook([string]$PackId, [string]$Mode) {
    $workbookId = 'wb_' + [Guid]::NewGuid().ToString('N').Substring(0,12)
    [void](Update-StructureLocked 'ja' {
        param($st)
        $wb = [pscustomobject][ordered]@{
            workbookId=$workbookId; packId=$PackId; sourceType='excel'; fileName="$workbookId.xlsx"; relativePath="$workbookId.xlsx"; displayName=$workbookId
            status='rendered'; excelSheetSelectionMode=$Mode; lastRenderedSheetSelectionMode=$Mode; lastRenderedSheets=@(); lastRenderedVersionId=''
        }
        Set-NoteProperty $st 'workbooks' (@(Get-Array $st.workbooks) + @($wb))
        return $null
    })
    return $workbookId
}

# Render-Workbook がExcelで調べた結果と同じ形を作り、ロック内と同じ順で反映する。
# $Sheets は @{ name; hidden } の並び（位置は並び順）。PDF化したシートには変換PDFを付ける。
function Invoke-TestRender([string]$WorkbookId, [string]$Mode, [object[]]$Sheets) {
    return Update-StructureLocked 'ja' {
        param($st)
        $wb = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId })[0]
        $all = @(); $allNames = @(); $targets = @(); $excluded = @(); $hidden = @()
        $i = 0
        foreach ($sheet in $Sheets) {
            $i++
            $name = [string]$sheet.name
            $allNames += $name
            $all += [ordered]@{ sheetName=$name; sheetIndex=$i; titleSource='A1'; detectedTitle="$WorkbookId / $name"; printArea='' }
            if ([bool]$sheet.hidden) { $excluded += $name; $hidden += $name }
            elseif ($Mode -eq 'numeric-only' -and -not (Test-StrictNumericSheetName $name)) { $excluded += $name }
            else { $targets += $name }
        }
        $sync = Update-WorkbookPagesFromInspection 'ja' $st $wb $all $allNames $targets $hidden $true
        $removedBySelection = @(Remove-SheetSelectionExcludedPages $st $WorkbookId $excluded $Mode)
        foreach ($name in $targets) {
            $page = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId -and [string]$_.sheetName -eq $name } | Select-Object -First 1)
            if ($page.Count) {
                Set-NoteProperty $page[0] 'contentPdf' ("content-pdf\$WorkbookId\v1\" + (Get-WorksheetStorageStem $name) + '.pdf')
                Set-NoteProperty $page[0] 'status' 'rendered'
                Set-NoteProperty $page[0] 'sheetSelectionExcluded' $false
            }
        }
        Set-WorkbookRenderedSheetSnapshot $wb (Get-DataProperty $sync 'sheetNames' @())
        Set-NoteProperty $wb 'lastRenderedSheetSelectionMode' $Mode
        return [ordered]@{ sync=$sync; removedBySelection=@($removedBySelection) }
    }
}

function Get-TestPages([string]$WorkbookId) {
    return @(Get-Array (Get-Structure 'ja').pages | Where-Object { [string]$_.workbookId -eq $WorkbookId })
}
function Get-TestPage([string]$WorkbookId, [string]$SheetName) {
    return @(Get-TestPages $WorkbookId | Where-Object { [string]$_.sheetName -eq $SheetName } | Select-Object -First 1)
}
function Test-Unassigned($Page) {
    return ([string](Get-DataProperty $Page 'volume' 'none') -eq 'none' -or (Get-DataProperty $Page 'enabled' $true) -eq $false)
}
function Move-TestPage([string]$WorkbookId, [string]$SheetName, [string]$Volume) {
    [void](Update-StructureLocked 'ja' {
        param($st)
        $page = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId -and [string]$_.sheetName -eq $SheetName })[0]
        Set-NoteProperty $page 'volume' $Volume
        Set-NoteProperty $page 'enabled' ($Volume -ne 'none')
        return $null
    })
}
function Get-UnassignedCount([string]$PackId) {
    $row = @((Get-PackProgressDashboard (Get-Structure 'ja') 'ja').packs | Where-Object { [string]$_.packId -eq $PackId })[0]
    return [int]$row.unassignedPageCount
}

$pack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='振り分けの一式' })
$packId = [string]$pack.packId

# ---- 症状: 半角数字シートを外して改名すると、未振り分けに残り続ける -----------------
$wbRename = New-TestWorkbook $packId 'numeric-only'
[void](Invoke-TestRender $wbRename 'numeric-only' @(@{name='1'}, @{name='2'}, @{name='3'}))
Assert-SheetTest (@(Get-TestPages $wbRename).Count -eq 3) 'Numeric sheets did not become pages.'
foreach ($name in @('1','2','3')) { Move-TestPage $wbRename $name 'ja-main' }
$oldId = Resolve-PageId @(Get-TestPage $wbRename '3')[0]
# 一式から外す: 本体から未振り分けへ移し、シート名を半角数字以外へ変える。
Move-TestPage $wbRename '3' 'none'
$result = Invoke-TestRender $wbRename 'numeric-only' @(@{name='1'}, @{name='2'}, @{name='3_参考'})
$left = @(Get-TestPages $wbRename)
Assert-SheetTest (@($left | Where-Object { [string]$_.sheetName -in @('3','3_参考') }).Count -eq 0) 'A numeric sheet removed and renamed to a non-numeric name stayed in the page composition.'
Assert-SheetTest (@($left | Where-Object { (Resolve-PageId $_) -eq $oldId }).Count -eq 0) 'The page of the old sheet name stayed in the page composition.'
Assert-SheetTest (@(Get-Array $result.sync.renamedPages).Count -eq 0) 'A rename into an excluded sheet name was reported as carried over.'
Assert-SheetTest (@($left | Where-Object { Test-Unassigned $_ }).Count -eq 0) 'Pages were left in the unassigned lane.'
Assert-SheetTest ((Get-UnassignedCount $packId) -eq 0) 'The home dashboard still counts unassigned pages.'
Assert-SheetTest (@($left | Where-Object { [string]$_.volume -eq 'ja-main' }).Count -eq 2) 'Pages assigned to 本体 were lost.'

# 本体に入れたまま改名しても同じ（外すだけで出力先も再出力が必要になる）。
$wbDirect = New-TestWorkbook $packId 'numeric-only'
[void](Invoke-TestRender $wbDirect 'numeric-only' @(@{name='10'}, @{name='20'}))
Move-TestPage $wbDirect '10' 'ja-main'; Move-TestPage $wbDirect '20' 'ja-main'
[void](Invoke-TestRender $wbDirect 'numeric-only' @(@{name='10'}, @{name='メモ'}))
Assert-SheetTest (@(Get-TestPages $wbDirect).Count -eq 1 -and [string]@(Get-TestPages $wbDirect)[0].sheetName -eq '10') 'A renamed sheet that left 本体 remained as a page.'

# 非表示にして外した半角数字シートも残らない。再表示すると新しいページとして戻る。
$wbHidden = New-TestWorkbook $packId 'numeric-only'
[void](Invoke-TestRender $wbHidden 'numeric-only' @(@{name='1'}, @{name='2'}))
$hiddenId = Resolve-PageId @(Get-TestPage $wbHidden '2')[0]
$hiddenResult = Invoke-TestRender $wbHidden 'numeric-only' @(@{name='1'}, @{name='2'; hidden=$true})
Assert-SheetTest (@(Get-TestPage $wbHidden '2').Count -eq 0) 'A hidden numeric sheet stayed in the unassigned lane.'
Assert-SheetTest (@($hiddenResult.removedBySelection).Count -eq 1) 'The removed hidden-sheet page was not reported.'
[void](Invoke-TestRender $wbHidden 'numeric-only' @(@{name='1'}, @{name='2'}))
$back = @(Get-TestPage $wbHidden '2')
Assert-SheetTest ($back.Count -eq 1 -and (Resolve-PageId $back[0]) -eq $hiddenId) 'A re-shown sheet did not come back with the same page ID (layout history could not find it).'

# 一度もPDF化していない半角数字以外のシートは、もともとページにならない。
$wbNever = New-TestWorkbook $packId 'numeric-only'
[void](Invoke-TestRender $wbNever 'numeric-only' @(@{name='表紙'}, @{name='1'}))
Assert-SheetTest (@(Get-TestPages $wbNever).Count -eq 1) 'A never-rendered non-numeric sheet became a page.'

# ---- 表示中の全シート: 改名・非表示は配置を引き継ぐ（従来どおり） ----------------------
$wbAll = New-TestWorkbook $packId 'all-visible'
[void](Invoke-TestRender $wbAll 'all-visible' @(@{name='3'}, @{name='4'}))
Move-TestPage $wbAll '3' 'ja-main'; Move-TestPage $wbAll '4' 'ja-appendix'
$keptId = Resolve-PageId @(Get-TestPage $wbAll '3')[0]
$allResult = Invoke-TestRender $wbAll 'all-visible' @(@{name='参考'}, @{name='4'; hidden=$true})
$renamedPage = @(Get-TestPage $wbAll '参考')
Assert-SheetTest ($renamedPage.Count -eq 1 -and (Resolve-PageId $renamedPage[0]) -eq $keptId -and [string]$renamedPage[0].volume -eq 'ja-main') 'All-visible rename no longer carries over the assignment.'
Assert-SheetTest (@(Get-Array $allResult.sync.renamedPages).Count -eq 1) 'All-visible rename was not reported.'
$hiddenAll = @(Get-TestPage $wbAll '4')
Assert-SheetTest ($hiddenAll.Count -eq 1 -and [string]$hiddenAll[0].volume -eq 'ja-appendix') 'All-visible hidden sheet lost its assignment.'

# ---- 改名後に旧名のシートが現れても、ページIDが重ならない ------------------------------
# 改名を引き継いだページは旧名のIDを持つ。新しい「3」に同じIDを付けると items が重複し、
# structure.json の保存が拒否されて、PDF作成もページ構成の保存もできなくなっていた。
[void](Invoke-TestRender $wbAll 'all-visible' @(@{name='参考'}, @{name='4'; hidden=$true}, @{name='3'}))
$pagesAll = @(Get-TestPages $wbAll)
$ids = @($pagesAll | ForEach-Object { Resolve-PageId $_ })
Assert-SheetTest ($pagesAll.Count -eq 3) 'The new sheet with the old name did not become a page.'
Assert-SheetTest (@($ids | Select-Object -Unique).Count -eq $ids.Count) 'A new sheet reused the page ID of a renamed sheet.'
Assert-SheetTest ([string]@(Get-TestPage $wbAll '3')[0].contentPdf -ne [string]@(Get-TestPage $wbAll '参考')[0].contentPdf) 'The new sheet and the renamed sheet share a converted PDF.'

# ---- 以前の版の残骸を起動時に整理する --------------------------------------------------
$wbLegacy = New-TestWorkbook $packId 'numeric-only'
[void](Update-StructureLocked 'ja' {
    param($st)
    $pages = @(Get-Array $st.pages)
    # 以前の版が「未振り分けに戻した」ページ（改名後の名前・旧名のID）。
    $pages += [pscustomobject][ordered]@{ pageId="$wbLegacy-3"; workbookId=$wbLegacy; sheetName='3_旧'; sheetIndex=3; title='旧3'; volume='none'; enabled=$false; order=10; contentPdf=$null; status='not-rendered'; sheetSelectionExcluded=$true; renamedFromSheetName='3' }
    # 利用者が出力先へ移したページは残す。
    $pages += [pscustomobject][ordered]@{ pageId="$wbLegacy-4"; workbookId=$wbLegacy; sheetName='4_旧'; sheetIndex=4; title='旧4'; volume='ja-main'; enabled=$true; order=10; contentPdf=$null; status='not-rendered'; sheetSelectionExcluded=$true }
    # 通常のページは残す。
    $pages += [pscustomobject][ordered]@{ pageId="$wbLegacy-1"; workbookId=$wbLegacy; sheetName='1'; sheetIndex=1; title='1'; volume='none'; enabled=$false; order=20; contentPdf="content-pdf\$wbLegacy\v1\a.pdf"; status='rendered'; sheetSelectionExcluded=$false }
    Set-NoteProperty $st 'pages' @($pages)
    return $null
})
$historyDir = Get-LayoutHistoryDir 'ja' ([string](Get-LayoutScopeInfo (Get-Structure 'ja') $packId $false).storageId)
$snapshotsBefore = @(Get-ChildItem -LiteralPath $historyDir -File -Filter '*.json' -ErrorAction SilentlyContinue).Count
$init = Initialize-Or-MigrateStructure 'ja' $dataDir
Assert-SheetTest ([int](Get-DataProperty $init 'removedSheetSelectionResiduePageCount' 0) -eq 1) 'Leftover excluded pages were not cleaned when the workspace was opened.'
$legacyPages = @(Get-TestPages $wbLegacy)
Assert-SheetTest (@(Get-TestPage $wbLegacy '3_旧').Count -eq 0) 'The leftover page of a renamed sheet was kept.'
Assert-SheetTest (@(Get-TestPage $wbLegacy '4_旧').Count -eq 1 -and @(Get-TestPage $wbLegacy '1').Count -eq 1) 'Cleanup removed a page the user had assigned or a normal page.'
$snapshotsAfter = @(Get-ChildItem -LiteralPath $historyDir -File -Filter '*.json' -ErrorAction SilentlyContinue).Count
Assert-SheetTest ($snapshotsAfter -eq $snapshotsBefore + 1) 'Cleanup did not leave a restore point.'
# 2回目は何もしない。
$again = Initialize-Or-MigrateStructure 'ja' $dataDir
Assert-SheetTest ([int](Get-DataProperty $again 'removedSheetSelectionResiduePageCount' 0) -eq 0) 'Cleanup ran again on clean data.'
Assert-SheetTest (Test-StructureDocument (Read-StructureUnlocked 'ja') 'ja') 'The cleaned structure is not a valid document.'

Write-Output 'sheet-assignment regression ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
