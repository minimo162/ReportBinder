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
# - 物理ページの頁・左右・見開きの番号が、出力先ごとに1頁＝右から数えた値になる（指紋は変わらない）
# - 見開きの試算（Get-SpreadPreview）は保存せず、確定（Reorder-Pages）と同じ検証をし、after が確定後と一致する
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

# ---- 物理ページの頁・左右（見開き） ---------------------------------------------------
# 出力先ごとに1頁＝右から数え、奇数＝右（ReportPdfComposer.java のパンチ穴シフトと同じ）。
# 頁数は変換PDFから数える。ページ数だけを持つ偽のPDFを置く。
function Write-TestContentPdf([string]$WorkbookId, [string]$SheetName, [int]$PageCount) {
    $rel = [string]@(Get-TestPage $WorkbookId $SheetName)[0].contentPdf
    $full = [IO.Path]::GetFullPath((Join-Path (Get-WorkspacePath 'ja') $rel))
    New-Item -ItemType Directory -Path (Split-Path -Parent $full) -Force | Out-Null
    $body = '%PDF-1.4' + "`n" + ((1..$PageCount | ForEach-Object { "$_ 0 obj << /Type /Page >> endobj" }) -join "`n") + "`n%%EOF`n"
    [IO.File]::WriteAllText($full, $body, [Text.Encoding]::ASCII)
}
function Set-TestPagePlacement([string]$WorkbookId, [string]$SheetName, [string]$Volume, [double]$Order, $PageRange) {
    [void](Update-StructureLocked 'ja' {
        param($st)
        $page = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId -and [string]$_.sheetName -eq $SheetName })[0]
        Set-NoteProperty $page 'volume' $Volume
        Set-NoteProperty $page 'enabled' ($Volume -ne 'none')
        Set-NoteProperty $page 'order' $Order
        Set-NoteProperty $page 'pageRange' $PageRange
        return $null
    })
}
function Get-TestPhysicalPages([string]$PackId, [string]$Volume) {
    return @((Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' $Volume $PackId).physicalPages)
}
function Assert-PhysicalPage($Page, [int]$Number, [string]$Side, [int]$Spread, [string]$PageId, [int]$Part, [int]$PartCount, [int]$SourcePage, [string]$Label) {
    Assert-SheetTest ([int]$Page.outputPageNumber -eq $Number) "$Label : outputPageNumber was $($Page.outputPageNumber), expected $Number."
    Assert-SheetTest ([string]$Page.side -eq $Side) "$Label : side was '$($Page.side)', expected '$Side'."
    Assert-SheetTest ([int]$Page.spreadIndex -eq $Spread) "$Label : spreadIndex was $($Page.spreadIndex), expected $Spread."
    Assert-SheetTest ([string]$Page.pageId -eq $PageId) "$Label : pageId was '$($Page.pageId)', expected '$PageId'."
    Assert-SheetTest ([int]$Page.pagePartIndex -eq $Part -and [int]$Page.pagePartCount -eq $PartCount) "$Label : part was $($Page.pagePartIndex)/$($Page.pagePartCount), expected $Part/$PartCount."
    Assert-SheetTest ([int]$Page.sourcePdfPageNumber -eq $SourcePage) "$Label : sourcePdfPageNumber was $($Page.sourcePdfPageNumber), expected $SourcePage."
}

$spreadPack = New-DocumentPack 'ja' ([pscustomobject]@{ displayName='見開きの一式' })
$spreadPackId = [string]$spreadPack.packId
$wbSpread = New-TestWorkbook $spreadPackId 'all-visible'
[void](Invoke-TestRender $wbSpread 'all-visible' @(@{name='表紙'}, @{name='損益'}, @{name='営利'}, @{name='付録'}))
Write-TestContentPdf $wbSpread '表紙' 1
Write-TestContentPdf $wbSpread '損益' 2
Write-TestContentPdf $wbSpread '営利' 3
Write-TestContentPdf $wbSpread '付録' 1
Set-TestPagePlacement $wbSpread '表紙' 'ja-main' 10 $null
Set-TestPagePlacement $wbSpread '損益' 'ja-main' 20 $null
Set-TestPagePlacement $wbSpread '営利' 'ja-main' 30 $null
Set-TestPagePlacement $wbSpread '付録' 'ja-appendix' 10 $null
$idCover = Resolve-PageId @(Get-TestPage $wbSpread '表紙')[0]
$idPl = Resolve-PageId @(Get-TestPage $wbSpread '損益')[0]
$idOp = Resolve-PageId @(Get-TestPage $wbSpread '営利')[0]
$idAppendix = Resolve-PageId @(Get-TestPage $wbSpread '付録')[0]

# 1・2・3頁のシートが混ざると、2頁目以降の左右はシートの境目と無関係に交互になる。
$mainPhysical = @(Get-TestPhysicalPages $spreadPackId 'ja-main')
Assert-SheetTest ($mainPhysical.Count -eq 6) "本体 should have 6 physical pages, got $($mainPhysical.Count)."
Assert-PhysicalPage $mainPhysical[0] 1 'right' 0 $idCover 1 1 1 '本体1頁（表紙）'
Assert-PhysicalPage $mainPhysical[1] 2 'left' 1 $idPl 1 2 1 '本体2頁（損益1/2）'
Assert-PhysicalPage $mainPhysical[2] 3 'right' 1 $idPl 2 2 2 '本体3頁（損益2/2）'
Assert-PhysicalPage $mainPhysical[3] 4 'left' 2 $idOp 1 3 1 '本体4頁（営利1/3）'
Assert-PhysicalPage $mainPhysical[4] 5 'right' 2 $idOp 2 3 2 '本体5頁（営利2/3）'
Assert-PhysicalPage $mainPhysical[5] 6 'left' 3 $idOp 3 3 3 '本体6頁（営利3/3）'

# 出力先ごとに1頁＝右から数え直す。
$appendixPhysical = @(Get-TestPhysicalPages $spreadPackId 'ja-appendix')
Assert-SheetTest ($appendixPhysical.Count -eq 1) "補足 should have 1 physical page, got $($appendixPhysical.Count)."
Assert-PhysicalPage $appendixPhysical[0] 1 'right' 0 $idAppendix 1 1 1 '補足1頁（付録）'

# 指紋は頁・左右を含まない（項目を足しても「再出力が必要」にならない）。
$snapBefore = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' $spreadPackId
$fpJson = ConvertTo-Json $snapBefore.fingerprintInput -Depth 20 -Compress
foreach ($field in @('side','spreadIndex','pagePartIndex','pagePartCount','outputPageNumber')) {
    Assert-SheetTest (-not $fpJson.Contains('"' + $field + '"')) "The fingerprint input contains the physical-page field '$field'."
}
Assert-SheetTest ([string]$snapBefore.fingerprint -eq (Get-Sha256Text $fpJson)) 'The fingerprint is no longer computed from the fingerprint input alone.'
$snapAgain = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' $spreadPackId
Assert-SheetTest ([string]$snapAgain.fingerprint -eq [string]$snapBefore.fingerprint) 'The fingerprint changed without any input change.'

# 使用ページ範囲は、枚数（pagePartCount）と後ろのページの左右に反映される。
Set-TestPagePlacement $wbSpread '損益' 'ja-main' 20 '2'
$rangePhysical = @(Get-TestPhysicalPages $spreadPackId 'ja-main')
Assert-SheetTest ($rangePhysical.Count -eq 5) "本体 with 損益 limited to page 2 should have 5 physical pages, got $($rangePhysical.Count)."
Assert-PhysicalPage $rangePhysical[1] 2 'left' 1 $idPl 1 1 2 '範囲指定後の本体2頁（損益の2ページ目だけ）'
Assert-PhysicalPage $rangePhysical[2] 3 'right' 1 $idOp 1 3 1 '範囲指定後の本体3頁（営利1/3は右へ反転）'
Assert-PhysicalPage $rangePhysical[4] 5 'right' 2 $idOp 3 3 3 '範囲指定後の本体5頁（営利3/3）'
Set-TestPagePlacement $wbSpread '営利' 'ja-main' 30 '2-3'
$rangePhysical = @(Get-TestPhysicalPages $spreadPackId 'ja-main')
Assert-SheetTest ($rangePhysical.Count -eq 4) "本体 with two page ranges should have 4 physical pages, got $($rangePhysical.Count)."
Assert-PhysicalPage $rangePhysical[2] 3 'right' 1 $idOp 1 2 2 '範囲指定後の本体3頁（営利の2ページ目）'
Assert-PhysicalPage $rangePhysical[3] 4 'left' 2 $idOp 2 2 3 '範囲指定後の本体4頁（営利の3ページ目）'
$snapRange = Get-FinalBuildInputSnapshot (Get-Structure 'ja') 'ja' 'ja-main' $spreadPackId
Assert-SheetTest ([string]$snapRange.fingerprint -ne [string]$snapBefore.fingerprint) 'Changing the page range no longer changes the fingerprint.'

# ---- 見開きの試算（保存しない before/after） ------------------------------------------
Set-TestPagePlacement $wbSpread '損益' 'ja-main' 20 $null
Set-TestPagePlacement $wbSpread '営利' 'ja-main' 30 $null
# 未振り分けから入れるページ（2頁）と、変換PDFが無く頁数がわからないページ。
$wbExtra = New-TestWorkbook $spreadPackId 'all-visible'
[void](Invoke-TestRender $wbExtra 'all-visible' @(@{name='追加'}, @{name='未作成'}))
Write-TestContentPdf $wbExtra '追加' 2
$idExtra = Resolve-PageId @(Get-TestPage $wbExtra '追加')[0]
$idMissing = Resolve-PageId @(Get-TestPage $wbExtra '未作成')[0]
Assert-SheetTest ((Test-Unassigned @(Get-TestPage $wbExtra '追加')[0]) -and (Test-Unassigned @(Get-TestPage $wbExtra '未作成')[0])) 'New pages for the preview test were not unassigned.'

function New-TestReorderBody($Volumes, [string]$BaseLayout = $null) {
    $request = [ordered]@{ packId=$spreadPackId; volumes=$Volumes }
    if ($null -ne $BaseLayout) { $request['baseLayout'] = $BaseLayout }
    return (ConvertTo-Json $request -Depth 6 | ConvertFrom-Json)
}
function Get-TestPhysicalKey($Pages) {
    return (@($Pages | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.outputPageNumber, $_.side, $_.spreadIndex, $_.pageId, $_.sourcePdfPageNumber, $_.numberingMode }) -join ';')
}
function Get-TestPreviewError($Body) {
    try { [void](Get-SpreadPreview 'ja' $Body); return $null } catch { return $_.Exception }
}
# 補足1頁の付録を本体の2頁目へ、未振り分けの2枚を補足へ。
$proposal = [ordered]@{ 'ja-main'=@($idCover,$idAppendix,$idPl,$idOp); 'ja-appendix'=@($idExtra,$idMissing); none=@() }
$structurePath = Join-Path (Get-WorkspacePath 'ja') 'structure.json'
$spreadHistoryDir = Get-LayoutHistoryDir 'ja' ([string](Get-LayoutScopeInfo (Get-Structure 'ja') $spreadPackId $false).storageId)
$hashBefore = (Get-FileHash -LiteralPath $structurePath -Algorithm SHA256).Hash
$historyBefore = @(Get-ChildItem -LiteralPath $spreadHistoryDir -File -Filter '*.json' -ErrorAction SilentlyContinue).Count
$readinessMainBefore = Get-TestPhysicalKey (Get-TestPhysicalPages $spreadPackId 'ja-main')
$layoutBefore = Get-PageLayoutFingerprint (Get-Structure 'ja') $spreadPackId

$preview = Get-SpreadPreview 'ja' (New-TestReorderBody $proposal $layoutBefore)
Assert-SheetTest ([string]$preview.layoutFingerprint -eq $layoutBefore) 'The preview did not return the current layout fingerprint.'
Assert-SheetTest ((@($preview.before.Keys) -join ',') -eq 'main,appendix' -and (@($preview.after.Keys) -join ',') -eq 'main,appendix') "The preview targets were '$(@($preview.after.Keys) -join ',')'."
$beforeMain = @($preview.before['main'].physicalPages)
Assert-SheetTest ((Get-TestPhysicalKey $beforeMain) -eq $readinessMainBefore) 'The preview before does not match the readiness physical pages.'
Assert-SheetTest ([int]$preview.before['main'].pageCount -eq 6 -and [int]$preview.before['appendix'].pageCount -eq 1) 'The preview before page counts are wrong.'
Assert-SheetTest (@($preview.before['appendix'].unknownPageIds).Count -eq 0) 'The preview before reported unknown pages in 補足.'
$afterMain = @($preview.after['main'].physicalPages)
Assert-SheetTest ($afterMain.Count -eq 7 -and [int]$preview.after['main'].pageCount -eq 7) "本体 after the move should have 7 physical pages, got $($afterMain.Count)."
Assert-PhysicalPage $afterMain[0] 1 'right' 0 $idCover 1 1 1 '試算後の本体1頁（表紙）'
Assert-PhysicalPage $afterMain[1] 2 'left' 1 $idAppendix 1 1 1 '試算後の本体2頁（付録）'
Assert-PhysicalPage $afterMain[2] 3 'right' 1 $idPl 1 2 1 '試算後の本体3頁（損益1/2は左から右へ反転）'
Assert-PhysicalPage $afterMain[6] 7 'right' 3 $idOp 3 3 3 '試算後の本体7頁（営利3/3）'
$afterAppendix = $preview.after['appendix']
Assert-SheetTest ([int]$afterAppendix.pageCount -eq 2) "補足 after the move should have 2 physical pages, got $($afterAppendix.pageCount)."
Assert-PhysicalPage @($afterAppendix.physicalPages)[0] 1 'right' 0 $idExtra 1 2 1 '試算後の補足1頁（追加1/2）'
Assert-SheetTest ((@($afterAppendix.unknownPageIds) -join ',') -eq $idMissing) "unknownPageIds was '$(@($afterAppendix.unknownPageIds) -join ',')', expected the page without a PDF."
Assert-SheetTest (((@($preview.changedVolumes) | Sort-Object) -join ',') -eq 'ja-appendix,ja-main,none') "changedVolumes was '$(@($preview.changedVolumes) -join ',')'."

# 試算は何も保存しない。共有キャッシュの構造も書き換えない。
Assert-SheetTest ((Get-FileHash -LiteralPath $structurePath -Algorithm SHA256).Hash -eq $hashBefore) 'The preview wrote structure.json.'
Assert-SheetTest (@(Get-ChildItem -LiteralPath $spreadHistoryDir -File -Filter '*.json' -ErrorAction SilentlyContinue).Count -eq $historyBefore) 'The preview saved a layout snapshot.'
Assert-SheetTest ((Get-TestPhysicalKey (Get-TestPhysicalPages $spreadPackId 'ja-main')) -eq $readinessMainBefore) 'The preview changed the cached structure.'
Assert-SheetTest ((Get-PageLayoutFingerprint (Get-Structure 'ja') $spreadPackId) -eq $layoutBefore) 'The preview changed the page layout.'

# 出力先IDでの指定（targets）も同じ結果になる。
$targetsBody = (ConvertTo-Json ([ordered]@{ packId=$spreadPackId; targets=[ordered]@{ main=$proposal['ja-main']; appendix=$proposal['ja-appendix']; unassigned=@() } }) -Depth 6 | ConvertFrom-Json)
$targetsPreview = Get-SpreadPreview 'ja' ([pscustomobject][ordered]@{ packId=$spreadPackId; volumes=(ConvertTo-PageReorderVolumes 'ja' $targetsBody) })
Assert-SheetTest ((Get-TestPhysicalKey ($targetsPreview.after['main'].physicalPages)) -eq (Get-TestPhysicalKey $afterMain)) 'The targets form gave a different preview.'

# 確定と同じ検証をする（未知のID・重複・欠落・不正な出力先）。どれも保存しない。
$badCases = @(
    @{ label='unknown id'; volumes=[ordered]@{ 'ja-main'=@($idCover,$idPl,$idOp,'page_unknown') } },
    @{ label='duplicate id'; volumes=[ordered]@{ 'ja-main'=@($idCover,$idPl,$idOp,$idCover) } },
    @{ label='dropped page'; volumes=[ordered]@{ 'ja-main'=@($idCover,$idPl) } },
    @{ label='invalid volume'; volumes=[ordered]@{ 'ja-other'=@($idCover) } }
)
foreach ($case in $badCases) {
    $err = Get-TestPreviewError (New-TestReorderBody $case.volumes)
    Assert-SheetTest ($err -is [System.ArgumentException]) "The preview accepted a bad request ($($case.label)): $err"
}
$err = Get-TestPreviewError ([pscustomobject][ordered]@{ packId=$spreadPackId })
Assert-SheetTest ($err -is [System.ArgumentException]) 'The preview accepted a request without volumes.'
# 古い並びに対する試算は 409（structure-conflict）で止める。
$err = Get-TestPreviewError (New-TestReorderBody $proposal ('0' * 64))
Assert-SheetTest (Test-StructureConflictError $err) "A stale baseLayout was not rejected as a conflict: $err"
Assert-SheetTest ((Get-FileHash -LiteralPath $structurePath -Algorithm SHA256).Hash -eq $hashBefore) 'A rejected preview wrote structure.json.'

# 試算の after は、同じ案を確定したあとの物理ページと一致する。
$committed = Reorder-Pages 'ja' (New-TestReorderBody $proposal $layoutBefore)
Assert-SheetTest (-not [string]::IsNullOrWhiteSpace([string]$committed.layoutSnapshotId)) 'Committing the move did not save a layout snapshot.'
Assert-SheetTest ((Get-TestPhysicalKey (Get-TestPhysicalPages $spreadPackId 'ja-main')) -eq (Get-TestPhysicalKey $afterMain)) 'The committed 本体 differs from the preview after.'
Assert-SheetTest ((Get-TestPhysicalKey (Get-TestPhysicalPages $spreadPackId 'ja-appendix')) -eq (Get-TestPhysicalKey ($afterAppendix.physicalPages))) 'The committed 補足 differs from the preview after.'
Assert-SheetTest ([string]$committed.layoutFingerprint -ne $layoutBefore) 'The committed layout fingerprint did not change.'
# 確定後に前の指紋で試算すると衝突になる。
$err = Get-TestPreviewError (New-TestReorderBody $proposal $layoutBefore)
Assert-SheetTest (Test-StructureConflictError $err) 'A preview against the pre-commit layout was not rejected.'
# 並びが同じ案は、before と after が同じになる。
$samePreview = Get-SpreadPreview 'ja' (New-TestReorderBody $proposal ([string]$committed.layoutFingerprint))
Assert-SheetTest (@($samePreview.changedVolumes).Count -eq 0) 'An unchanged proposal reported changed volumes.'
Assert-SheetTest ((Get-TestPhysicalKey ($samePreview.before['main'].physicalPages)) -eq (Get-TestPhysicalKey ($samePreview.after['main'].physicalPages))) 'An unchanged proposal changed the physical pages.'

Write-Output 'sheet-assignment regression ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_LOCAL_CONFIG_ROOT', $oldConfigRoot, 'Process')
    [Environment]::SetEnvironmentVariable('REPORTBINDER_OUTPUT_TEST_APPROOT', $oldAppRoot, 'Process')
    if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
