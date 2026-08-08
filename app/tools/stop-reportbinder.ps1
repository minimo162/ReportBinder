$ErrorActionPreference = 'Continue'
$ToolsDir = $PSScriptRoot
$AppRoot = Split-Path -Parent $ToolsDir
$RootDir = Split-Path -Parent $AppRoot
$server = [IO.Path]::GetFullPath((Join-Path $AppRoot 'server.ps1'))
$launch = [IO.Path]::GetFullPath((Join-Path $AppRoot 'launch.ps1'))

function Get-ReportBinderLocalLaunchDir([string]$PathToAppRoot) {
    $appRootKey = 'default'
    try {
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            $hashBytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($PathToAppRoot.ToLowerInvariant()))
            $appRootKey = -join ($hashBytes[0..7] | ForEach-Object { $_.ToString('x2') })
        } finally { try { $sha.Dispose() } catch {} }
    } catch { $appRootKey = 'default' }
    return (Join-Path ([IO.Path]::GetTempPath()) ("ReportBinder-$appRootKey"))
}

function CommandLine-ContainsPath([string]$CommandLine, [string]$Path) {
    if ([string]::IsNullOrWhiteSpace($CommandLine) -or [string]::IsNullOrWhiteSpace($Path)) { return $false }
    $cmd = $CommandLine.ToLowerInvariant()
    $p1 = ([IO.Path]::GetFullPath($Path)).ToLowerInvariant()
    $p2 = ($p1 -replace '\\','/')
    return ($cmd.Contains($p1) -or $cmd.Contains($p2))
}

$targets = @($server, $launch)
$runtimeVersionsRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ReportBinder\runtime\versions'
if (Test-Path -LiteralPath $runtimeVersionsRoot -PathType Container) {
    foreach ($versionDir in @(Get-ChildItem -LiteralPath $runtimeVersionsRoot -Directory -ErrorAction SilentlyContinue)) {
        foreach ($relativeScript in @('app\server.ps1','app\launch.ps1')) {
            $runtimeScript = Join-Path $versionDir.FullName $relativeScript
            if (Test-Path -LiteralPath $runtimeScript -PathType Leaf) { $targets += [IO.Path]::GetFullPath($runtimeScript) }
        }
    }
}
$targets = @($targets | Sort-Object -Unique)
$processes = @()
try {
    $processes = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
        $cmd = [string]$_.CommandLine
        $matched = $false
        foreach ($target in $targets) {
            if (CommandLine-ContainsPath $cmd $target) { $matched = $true; break }
        }
        $matched
    })
} catch {}

if ($processes.Count -eq 0) {
    Write-Host 'No ReportBinder process was found.'
} else {
    foreach ($p in @($processes | Sort-Object ProcessId -Unique)) {
        try {
            Write-Host ("Stopping ReportBinder process PID={0}" -f $p.ProcessId)
            Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
        } catch {}
    }
}

$logDir = Join-Path $AppRoot 'logs'
if (Test-Path -LiteralPath $logDir) {
    foreach ($pattern in @('open-*-latest.url', 'ReportBinder-*.url', 'ReportBinder-*-Edge.cmd', 'startup-wait-*-latest.html', 'server-*-latest.pid', 'launch-*-latest.pid')) {
        try { Get-ChildItem -LiteralPath $logDir -Filter $pattern -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue } catch {}
    }
}
$localLaunchDir = Get-ReportBinderLocalLaunchDir $AppRoot
if (Test-Path -LiteralPath $localLaunchDir) {
    foreach ($pattern in @('startup-wait-*-latest.html')) {
        try { Get-ChildItem -LiteralPath $localLaunchDir -Filter $pattern -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue } catch {}
    }
}
Write-Host 'ReportBinder stop completed.'
