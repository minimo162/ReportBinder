param(
    [Parameter(Mandatory=$true)][string]$InputPathB64,
    [Parameter(Mandatory=$true)][string]$OutputPathB64,
    [Parameter(Mandatory=$true)][string]$ResultPathB64
)

$ErrorActionPreference = 'Stop'

function Decode-Path([string]$Value) {
    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Value))
}

function Release-Com($Object) {
    if ($null -eq $Object) { return }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch { }
}

$inputPath = Decode-Path $InputPathB64
$outputPath = Decode-Path $OutputPathB64
$resultPath = Decode-Path $ResultPathB64
$existingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$word = $null
$documents = $null
$options = $null
$document = $null
$wordPid = 0
$stage = 'start'
$result = [ordered]@{ ok = $false; stage = $stage; errorCode = ''; message = ''; wordVersion = ''; pageCount = 0; outputPath = $outputPath; ownedProcessId = 0 }

try {
    if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw 'WORD_INPUT_MISSING' }
    $parent = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    try { Unblock-File -LiteralPath $inputPath -ErrorAction SilentlyContinue } catch { }

    $stage = 'word-start'
    $word = New-Object -ComObject Word.Application
    try {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ReportBinderWordNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@ -ErrorAction SilentlyContinue
        [uint32]$pidValue = 0
        [void][ReportBinderWordNative]::GetWindowThreadProcessId([IntPtr][int]$word.Hwnd, [ref]$pidValue)
        $wordPid = [int]$pidValue
    } catch { $wordPid = 0 }
    if ($wordPid -le 0) {
        Start-Sleep -Milliseconds 200
        $newWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $existingWordPids -notcontains [int]$_.Id } | ForEach-Object { [int]$_.Id })
        if ($newWordPids.Count -eq 1) { $wordPid = $newWordPids[0] }
    }
    $result.ownedProcessId = $wordPid
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch { }
    try { $word.ScreenUpdating = $false } catch { }
    try {
        $options = $word.Options
        $options.UpdateLinksAtOpen = $false
        $options.UpdateFieldsAtPrint = $false
        $options.PrintBackground = $false
        $options.PrintRevisions = $false
    } catch { }
    try { $result.wordVersion = [string]$word.Version } catch { }

    $stage = 'document-open'
    # Word's IDispatch signature differs slightly across Office builds. Supplying only
    # the stable leading arguments avoids PowerShell marshaling null optional values.
    $documents = $word.Documents
    $document = $documents.Open($inputPath, $false, $true, $false)
    if ($null -eq $document) { throw 'WORD_OPEN_FAILED' }
    try { $document.ShowRevisions = $false } catch { }
    try { $result.pageCount = [int]$document.ComputeStatistics(2, $false) } catch { }

    $stage = 'pdf-export'
    try {
        $document.ExportAsFixedFormat($outputPath, 17, $false, 0, 0, 1, 9999999, 0, $true, $true, 1, $true, $true, $false)
    } catch {
        try { $document.SaveAs2($outputPath, 17, $false, '', $false, '', $false, $false, $false, $false, $false, 0, $false, $false, 0, $false, $false) }
        catch { throw $_ }
    }
    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) { throw 'WORD_PDF_MISSING' }
    if ((Get-Item -LiteralPath $outputPath).Length -le 0) { throw 'WORD_PDF_EMPTY' }

    $stage = 'complete'
    $result.ok = $true
    $result.stage = $stage
} catch {
    $message = [string]$_.Exception.Message
    $code = if ($message -match '^WORD_[A-Z_]+$') { $message } elseif ($stage -eq 'document-open') { 'WORD_OPEN_FAILED' } elseif ($stage -eq 'pdf-export') { 'WORD_EXPORT_FAILED' } elseif ($stage -eq 'word-start') { 'WORD_NOT_AVAILABLE' } else { 'WORD_WORKER_FAILED' }
    $result.ok = $false
    $result.stage = $stage
    $result.errorCode = $code
    $result.message = $message
} finally {
    if ($document) { try { $document.Close($false) } catch { }; Release-Com $document }
    if ($documents) { Release-Com $documents }
    if ($options) { Release-Com $options }
    if ($word) { try { $word.Quit(0) } catch { }; Release-Com $word }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($wordPid -gt 0 -and $existingWordPids -notcontains $wordPid) {
        try {
            $ownedProcess = Get-Process -Id $wordPid -ErrorAction Stop
            if ([string]$ownedProcess.ProcessName -eq 'WINWORD') {
                if (-not $ownedProcess.WaitForExit(3000)) {
                    Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue
                    try { $ownedProcess.WaitForExit(5000) | Out-Null } catch { }
                }
            }
        } catch { }
    }
    try { $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resultPath -Encoding UTF8 }
    catch { }
}

if (-not [bool]$result.ok) { exit 1 }
