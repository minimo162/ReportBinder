param()

$ErrorActionPreference = 'Stop'
$toolsRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $toolsRoot
$serverPath = Join-Path $appRoot 'server.ps1'
$oldAppRoot = [Environment]::GetEnvironmentVariable('REPORTBINDER_HISTORY_TEST_APPROOT', 'Process')

try {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_HISTORY_TEST_APPROOT', $appRoot, 'Process')
    $server = Get-Content -LiteralPath $serverPath -Raw -Encoding UTF8
    $startupMarker = "if (-not [string]::IsNullOrWhiteSpace(`$AutoSchedulerPath)) {"
    $startupAt = $server.LastIndexOf($startupMarker, [StringComparison]::Ordinal)
    if ($startupAt -lt 0) { throw 'server.ps1 startup boundary was not found.' }
    $definitions = $server.Substring(0, $startupAt)
    $definitions = $definitions.Replace('$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path', '$Script:AppRoot = $env:REPORTBINDER_HISTORY_TEST_APPROOT')
    $testBody = @'
function Assert-HistoryTest([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}
function New-TestSheet([string]$Name, [string]$Hash, [string]$Status = 'ok') {
    return [pscustomobject][ordered]@{ sheetName=$Name; status=$Status; sheetVisualHash=('sha256:' + $Hash.PadLeft(64,'0')); pageCount=1; pageHashes=@(('sha256:' + $Hash.PadLeft(64,'0'))) }
}

$before = @(
    (New-TestSheet 'Page 1' '1'),
    (New-TestSheet 'Page 2' '2'),
    (New-TestSheet 'Page 3' '3')
)
$afterInserted = @(
    (New-TestSheet 'Page 1' '1'),
    (New-TestSheet 'Page 2' '9'),
    (New-TestSheet 'Page 3' '2'),
    (New-TestSheet 'Page 4' '3')
)
$inserted = @(Get-ComparisonUnitMappings $before $afterInserted 'word')
Assert-HistoryTest ($inserted.Count -eq 4) 'Inserted Word page mapping count is invalid.'
Assert-HistoryTest (@($inserted | Where-Object { $_.kind -eq 'added' -and $_.afterSheetName -eq 'Page 2' }).Count -eq 1) 'Inserted Word page was not classified as added.'
Assert-HistoryTest (@($inserted | Where-Object { $_.kind -eq 'unchanged' -and $_.beforeSheetName -eq 'Page 2' -and $_.afterSheetName -eq 'Page 3' -and $_.matchConfidence -eq 1 }).Count -eq 1) 'Shifted Word page was not matched by content.'
Assert-HistoryTest (@($inserted | Where-Object { $_.kind -eq 'modified' }).Count -eq 0) 'A page insertion incorrectly changed all following pages.'

$afterChanged = @(
    (New-TestSheet 'Page 1' '1'),
    (New-TestSheet 'Page 2' '8'),
    (New-TestSheet 'Page 3' '3')
)
$changed = @(Get-ComparisonUnitMappings $before $afterChanged 'pdf')
$changedPage = @($changed | Where-Object { $_.kind -eq 'modified' })
Assert-HistoryTest ($changedPage.Count -eq 1 -and $changedPage[0].beforeSheetName -eq 'Page 2' -and $changedPage[0].afterSheetName -eq 'Page 2') 'A changed PDF page was not paired between exact anchors.'
Assert-HistoryTest ([double]$changedPage[0].matchConfidence -ge 0.9) 'Changed PDF page confidence is unexpectedly low.'

$afterAmbiguous = @(
    (New-TestSheet 'Page 1' '1'),
    (New-TestSheet 'Page 2' '7'),
    (New-TestSheet 'Page 3' '8'),
    (New-TestSheet 'Page 4' '3')
)
$ambiguous = @(Get-ComparisonUnitMappings $before $afterAmbiguous 'word')
$unknown = @($ambiguous | Where-Object { $_.kind -eq 'unknown' })
Assert-HistoryTest ($unknown.Count -eq 1 -and [double]$unknown[0].matchConfidence -lt 0.5 -and $unknown[0].matchMethod -eq 'ambiguous-sequence') 'Ambiguous page correspondence must remain visibly uncertain.'

$excelBefore = @((New-TestSheet 'Summary' '1'))
$excelAfter = @((New-TestSheet 'Summary' '8'))
$excel = @(Get-ComparisonUnitMappings $excelBefore $excelAfter 'excel')
Assert-HistoryTest ($excel.Count -eq 1 -and $excel[0].kind -eq 'modified' -and $excel[0].matchMethod -eq 'stable-name') 'Excel sheet-name correspondence regressed.'

$result = [ordered]@{ changedSheets=@(); unchangedSheets=@(); unknownSheets=@(); addedSheets=@(); removedSheets=@(); unitMappings=@(); mappingConfidence=1.0 }
Set-ComparisonUnitMappingResult $result $inserted
Assert-HistoryTest (@($result.addedSheets) -contains 'Page 2') 'Legacy addedSheets compatibility was not derived from unit mappings.'
Assert-HistoryTest (@($result.unchangedSheets) -contains 'Page 3') 'Legacy unchangedSheets compatibility was not derived from unit mappings.'
$skeleton = New-DiffDetailSkeleton 'ja' ([pscustomobject][ordered]@{
    available=$false; message=''; workbookId='word-test'; workbookName='Word test'; sourceType='word'
    comparison=$result; currentVisualHashes=$null; baselineVisualHashes=$null
    currentSnapshotId='after'; currentVersionId='v2'; baselineSnapshotId='before'; baselineVersionId='v1'
    currentAt='2026-08-06T11:00:00+09:00'; baselineAt='2026-08-06T10:00:00+09:00'; method='test'; scope='history'
})
$shiftedDetail = @($skeleton.sheets | Where-Object { $_.beforeSheetName -eq 'Page 2' -and $_.afterSheetName -eq 'Page 3' })
Assert-HistoryTest ($shiftedDetail.Count -eq 1) 'Generalized diff detail lost the before/after page keys.'
Assert-HistoryTest (-not [string]::IsNullOrWhiteSpace([string]$skeleton.unitLabel)) 'Generalized diff detail unit label is missing.'
$unknownResult = [ordered]@{ changedSheets=@(); unchangedSheets=@(); unknownSheets=@(); addedSheets=@(); removedSheets=@(); unitMappings=@(); mappingConfidence=1.0 }
Set-ComparisonUnitMappingResult $unknownResult $ambiguous
$unknownSkeleton = New-DiffDetailSkeleton 'ja' ([pscustomobject][ordered]@{
    available=$false; message=''; workbookId='word-test'; workbookName='Word test'; sourceType='word'
    comparison=$unknownResult; currentVisualHashes=$null; baselineVisualHashes=$null
    currentSnapshotId='after'; currentVersionId='v2'; baselineSnapshotId='before'; baselineVersionId='v1'
    currentAt=''; baselineAt=''; method='test'; scope='history'
})
Assert-HistoryTest (@($unknownSkeleton.sheets | Where-Object { $_.kind -eq 'unknown' -and $_.status -eq 'unknown' -and [double]$_.matchConfidence -lt 0.5 }).Count -eq 1) 'Uncertain page mapping is not exposed as unknown in diff detail.'
Write-Output 'history-generalization selfcheck ok'
'@
    $script = [scriptblock]::Create($definitions + [Environment]::NewLine + $testBody)
    & $script -Mode ja -NoOpen
} finally {
    [Environment]::SetEnvironmentVariable('REPORTBINDER_HISTORY_TEST_APPROOT', $oldAppRoot, 'Process')
}
