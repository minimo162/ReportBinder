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
$existingPowerPointPids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { [int]$_.Id })
$powerPoint = $null
$presentations = $null
$presentation = $null
$powerPointPid = 0
$stage = 'start'
$result = [ordered]@{ ok=$false; stage=$stage; errorCode=''; message=''; powerPointVersion=''; slideCount=0; outputPath=$outputPath; ownedProcessId=0 }

try {
    if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw 'POWERPOINT_INPUT_MISSING' }
    $parent = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    try { Unblock-File -LiteralPath $inputPath -ErrorAction SilentlyContinue } catch { }

    $stage = 'powerpoint-start'
    $powerPoint = New-Object -ComObject PowerPoint.Application
    try {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ReportBinderPowerPointNative {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@ -ErrorAction SilentlyContinue
        [uint32]$pidValue = 0
        [void][ReportBinderPowerPointNative]::GetWindowThreadProcessId([IntPtr][int]$powerPoint.HWND, [ref]$pidValue)
        $powerPointPid = [int]$pidValue
    } catch { $powerPointPid = 0 }
    if ($powerPointPid -le 0) {
        Start-Sleep -Milliseconds 200
        $newPids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $existingPowerPointPids -notcontains [int]$_.Id } | ForEach-Object { [int]$_.Id })
        if ($newPids.Count -eq 1) { $powerPointPid = $newPids[0] }
    }
    $result.ownedProcessId = $powerPointPid
    try { $powerPoint.Visible = 0 } catch { }
    try { $powerPoint.DisplayAlerts = 1 } catch { }
    try { $powerPoint.AutomationSecurity = 3 } catch { }
    try { $result.powerPointVersion = [string]$powerPoint.Version } catch { }

    $stage = 'presentation-open'
    $presentations = $powerPoint.Presentations
    $presentation = $presentations.Open($inputPath, -1, 0, 0)
    if ($null -eq $presentation) { throw 'POWERPOINT_OPEN_FAILED' }
    try { $result.slideCount = [int]$presentation.Slides.Count } catch { }

    $stage = 'pdf-export'
    try { $presentation.ExportAsFixedFormat($outputPath, 2) }
    catch {
        try { $presentation.SaveAs($outputPath, 32) }
        catch { throw $_ }
    }
    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) { throw 'POWERPOINT_PDF_MISSING' }
    if ((Get-Item -LiteralPath $outputPath).Length -le 0) { throw 'POWERPOINT_PDF_EMPTY' }

    $stage = 'complete'
    $result.ok = $true
    $result.stage = $stage
} catch {
    $message = [string]$_.Exception.Message
    $code = if ($message -match '^POWERPOINT_[A-Z_]+$') { $message } elseif ($stage -eq 'presentation-open') { 'POWERPOINT_OPEN_FAILED' } elseif ($stage -eq 'pdf-export') { 'POWERPOINT_EXPORT_FAILED' } elseif ($stage -eq 'powerpoint-start') { 'POWERPOINT_NOT_AVAILABLE' } else { 'POWERPOINT_WORKER_FAILED' }
    $result.ok = $false
    $result.stage = $stage
    $result.errorCode = $code
    $result.message = $message
} finally {
    if ($presentation) { try { $presentation.Close() } catch { }; Release-Com $presentation }
    if ($presentations) { Release-Com $presentations }
    if ($powerPoint) { try { $powerPoint.Quit() } catch { }; Release-Com $powerPoint }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($powerPointPid -gt 0 -and $existingPowerPointPids -notcontains $powerPointPid) {
        try {
            $ownedProcess = Get-Process -Id $powerPointPid -ErrorAction Stop
            if ([string]$ownedProcess.ProcessName -eq 'POWERPNT' -and -not $ownedProcess.WaitForExit(3000)) {
                Stop-Process -Id $powerPointPid -Force -ErrorAction SilentlyContinue
                try { $ownedProcess.WaitForExit(5000) | Out-Null } catch { }
            }
        } catch { }
    }
    try { $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $resultPath -Encoding UTF8 } catch { }
}

if (-not [bool]$result.ok) { exit 1 }
