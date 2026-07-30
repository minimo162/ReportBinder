param(
    [ValidateSet('ja','en')]
    [string]$Mode = 'ja',
    [switch]$Diagnostics
)

$ErrorActionPreference = 'Stop'
$script:AppRoot = $PSScriptRoot
$script:RootDir = Split-Path -Parent $script:AppRoot
$logDir = Join-Path $script:AppRoot 'logs'
if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
# 30日より古いログ・ジョブファイルを削除する(*-latest.* は対象外)。
try {
    $logCutoff = (Get-Date).AddDays(-30)
    Get-ChildItem -LiteralPath $logDir -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch 'latest' -and $_.LastWriteTime -lt $logCutoff -and $_.Name -match '^(startup|server)-' } |
        Remove-Item -Force -ErrorAction SilentlyContinue
} catch { }

# Keep the startup wait page off shared folders. Some enterprise Edge/IE-mode
# policies classify file:// pages opened from a UNC share as IE mode even when
# msedge.exe is used explicitly.  The app itself still runs from AppRoot, but
# the temporary wait page is written under the user's local temp folder and then
# redirects to http://127.0.0.1 when the server is ready.
$appRootKey = 'default'
try {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($script:AppRoot.ToLowerInvariant()))
        $appRootKey = -join ($hashBytes[0..7] | ForEach-Object { $_.ToString('x2') })
    } finally { try { $sha.Dispose() } catch {} }
} catch {
    try { $appRootKey = ([Guid]::NewGuid().ToString('N')).Substring(0, 16) } catch { $appRootKey = 'default' }
}
$localLaunchDir = Join-Path ([IO.Path]::GetTempPath()) ("ReportBinder-$appRootKey")
if (-not (Test-Path -LiteralPath $localLaunchDir)) { New-Item -ItemType Directory -Path $localLaunchDir -Force | Out-Null }

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$log = Join-Path $logDir ("startup-$Mode-$stamp.log")
$latest = Join-Path $logDir ("startup-$Mode-latest.log")
$serverOut = Join-Path $logDir ("server-$Mode-$stamp.out.log")
$serverErr = Join-Path $logDir ("server-$Mode-$stamp.err.log")
$urlFile = Join-Path $logDir ("open-$Mode-latest.url")
$rootUrlFile = Join-Path $logDir ("ReportBinder-$Mode.url")
$serverPidFile = Join-Path $logDir ("server-$Mode-latest.pid")
$launchPidFile = Join-Path $logDir ("launch-$Mode-latest.pid")
$waitPageFile = Join-Path $localLaunchDir ("startup-wait-$Mode-latest.html")
$edgeCmdFile = Join-Path $logDir ("ReportBinder-$Mode-Edge.cmd")
$launchMutex = $null
$launchMutexOwned = $false

function Sync-LatestLog {
    try { Copy-Item -LiteralPath $log -Destination $latest -Force } catch {}
}

function Add-LaunchLog([string]$Message) {
    $line = "{0} {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -LiteralPath $log -Value $line -Encoding UTF8
    Sync-LatestLog
}

function Quote-ProcessArgument([string]$Value) {
    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"','\"') + '"'
}



function Resolve-EdgeExecutableFromCommandText([string]$CommandText) {
    if ([string]::IsNullOrWhiteSpace($CommandText)) { return '' }
    $expanded = [Environment]::ExpandEnvironmentVariables($CommandText.Trim())
    $candidate = ''
    if ($expanded -match '^\s*"([^"]*msedge\.exe)"') { $candidate = $matches[1] }
    elseif ($expanded -match '^\s*([^\s"]*msedge\.exe)') { $candidate = $matches[1] }
    elseif ($expanded -match '"([^"]*msedge\.exe)"') { $candidate = $matches[1] }
    elseif ($expanded -match '([^\s"]*msedge\.exe)') { $candidate = $matches[1] }
    if ([string]::IsNullOrWhiteSpace($candidate)) { return '' }
    $candidate = $candidate.Trim('"')
    try {
        if (Test-Path -LiteralPath $candidate) { return ([IO.Path]::GetFullPath($candidate)) }
    } catch {}
    return ''
}

function Add-EdgeExecutableCandidate($Candidates, [string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    $expanded = [Environment]::ExpandEnvironmentVariables($Value.Trim())
    $parsed = Resolve-EdgeExecutableFromCommandText $expanded
    if (-not [string]::IsNullOrWhiteSpace($parsed)) {
        [void]$Candidates.Add($parsed)
        return
    }
    $raw = $expanded.Trim('"')
    if (-not [string]::IsNullOrWhiteSpace($raw)) { [void]$Candidates.Add($raw) }
}

function Get-RegistryDefaultValueText([string]$Path) {
    try {
        $key = Get-Item -LiteralPath $Path -ErrorAction Stop
        return [string]$key.GetValue('')
    } catch {
        return ''
    }
}

function Get-EdgeExecutablePath {
    $candidates = New-Object System.Collections.Generic.List[string]
    try {
        $cmd = Get-Command msedge.exe -ErrorAction SilentlyContinue
        if ($cmd -and -not [string]::IsNullOrWhiteSpace([string]$cmd.Source)) {
            Add-EdgeExecutableCandidate $candidates ([string]$cmd.Source)
        }
    } catch {}

    foreach ($regPath in @(
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe',
        'HKCU:\SOFTWARE\Classes\MSEdgeHTM\shell\open\command',
        'HKLM:\SOFTWARE\Classes\MSEdgeHTM\shell\open\command',
        'HKCU:\SOFTWARE\Classes\microsoft-edge\shell\open\command',
        'HKLM:\SOFTWARE\Classes\microsoft-edge\shell\open\command'
    )) {
        Add-EdgeExecutableCandidate $candidates (Get-RegistryDefaultValueText $regPath)
    }

    foreach ($base in @(
        [Environment]::GetEnvironmentVariable('ProgramFiles(x86)'),
        [Environment]::GetEnvironmentVariable('ProgramFiles'),
        [Environment]::GetEnvironmentVariable('LocalAppData')
    )) {
        if (-not [string]::IsNullOrWhiteSpace($base)) {
            Add-EdgeExecutableCandidate $candidates (Join-Path $base 'Microsoft\Edge\Application\msedge.exe')
        }
    }

    foreach ($candidate in @($candidates | Select-Object -Unique)) {
        try {
            if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
                return ([IO.Path]::GetFullPath($candidate))
            }
        } catch {}
    }
    return ''
}

function ConvertTo-EdgeOpenTarget([string]$Target) {
    if ([string]::IsNullOrWhiteSpace($Target)) { return '' }
    $trimmed = $Target.Trim()
    try {
        # A Windows drive path such as C:\... must not be mistaken for a URI scheme.
        if ($trimmed -match '^[A-Za-z]:[\\/]' -or $trimmed -match '^\\\\') {
            return ([Uri]([IO.Path]::GetFullPath($trimmed))).AbsoluteUri
        }
        if ($trimmed -match '^file:') { return $trimmed }
        if ($trimmed -match '^[A-Za-z][A-Za-z0-9+.-]*:') { return $trimmed }
        if (Test-Path -LiteralPath $trimmed) {
            return ([Uri]([IO.Path]::GetFullPath($trimmed))).AbsoluteUri
        }
        return $trimmed
    } catch {
        return $trimmed
    }
}

function Write-CmdQuoted([string]$Value) {
    if ($null -eq $Value) { $Value = '' }
    return '"' + ($Value -replace '"','""') + '"'
}

function Write-EdgeOpenCommand([string]$Path, [string]$Target) {
    try {
        $edge = Get-EdgeExecutablePath
        $openTarget = ConvertTo-EdgeOpenTarget $Target
        $lines = New-Object System.Collections.Generic.List[string]
        [void]$lines.Add('@echo off')
        if (-not [string]::IsNullOrWhiteSpace($edge)) {
            [void]$lines.Add('start "" ' + (Write-CmdQuoted $edge) + ' ' + (Write-CmdQuoted $openTarget))
        } elseif ($openTarget -match '^https?://') {
            [void]$lines.Add('start "" ' + (Write-CmdQuoted ('microsoft-edge:' + $openTarget)))
        } else {
            [void]$lines.Add('start "" msedge.exe ' + (Write-CmdQuoted $openTarget))
        }
        Set-Content -LiteralPath $Path -Value ($lines -join "`r`n") -Encoding Default
    } catch {}
}

function New-SessionToken {
    $bytes = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $rng.GetBytes($bytes) }
    finally { try { $rng.Dispose() } catch {} }
    return ([Convert]::ToBase64String($bytes) -replace '[^A-Za-z0-9]', '')
}

function Get-FreePort {
    $listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Parse('127.0.0.1'), 0)
    $listener.Start()
    try { return $listener.LocalEndpoint.Port }
    finally { $listener.Stop() }
}

function Write-UrlShortcut([string]$Path, [string]$Url) {
    $content = "[InternetShortcut]`r`nURL=$Url`r`n"
    Set-Content -LiteralPath $Path -Value $content -Encoding ASCII
}

function Read-UrlShortcut([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    try {
        foreach ($line in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
            if ($line -match '^URL=(.+)$') { return $matches[1].Trim() }
        }
    } catch { }
    return ''
}

function Get-ReportBinderApiUrl([string]$Url, [string]$ApiPath) {
    try {
        $uri = [Uri]$Url
        if ($uri.Host -ne '127.0.0.1') { return '' }
        if ($uri.Port -le 0) { return '' }
        return "http://127.0.0.1:$($uri.Port)$ApiPath$($uri.Query)"
    } catch {
        return ''
    }
}

function Test-ReportBinderTcpHttpUrl([string]$Url, [string]$ApiPath, [int]$TimeoutMs) {
    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    if ($TimeoutMs -lt 200) { $TimeoutMs = 200 }
    $client = $null
    try {
        $uri = [Uri]$Url
        if ($uri.Host -ne '127.0.0.1' -or $uri.Port -le 0) { return $false }
        $target = "$ApiPath$($uri.Query)"
        $client = New-Object Net.Sockets.TcpClient
        $async = $client.BeginConnect('127.0.0.1', $uri.Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { return $false }
        $client.EndConnect($async)
        try { $async.AsyncWaitHandle.Close() } catch {}
        $client.SendTimeout = $TimeoutMs
        $client.ReceiveTimeout = $TimeoutMs
        $stream = $client.GetStream()
        $request = "GET $target HTTP/1.1`r`nHost: 127.0.0.1:$($uri.Port)`r`nConnection: close`r`nUser-Agent: ReportBinderLauncher`r`n`r`n"
        $requestBytes = [Text.Encoding]::ASCII.GetBytes($request)
        $stream.Write($requestBytes, 0, $requestBytes.Length)
        $stream.Flush()
        $buffer = New-Object byte[] 4096
        $ms = New-Object IO.MemoryStream
        while ($true) {
            try { $read = $stream.Read($buffer, 0, $buffer.Length) } catch { break }
            if ($read -le 0) { break }
            $ms.Write($buffer, 0, $read)
            if ($ms.Length -gt 8192) { break }
            $partial = [Text.Encoding]::UTF8.GetString($ms.ToArray())
            if ($partial -match '^HTTP/1\.[01]\s+200\s' -and $partial -match '"ok"\s*:\s*true') { return $true }
        }
        $text = [Text.Encoding]::UTF8.GetString($ms.ToArray())
        return ($text -match '^HTTP/1\.[01]\s+200\s' -and $text -match '"ok"\s*:\s*true')
    } catch {
        return $false
    } finally {
        try { if ($client) { $client.Close() } } catch {}
    }
}

function Test-ReportBinderUrl([string]$Url, [int]$TimeoutSec) {
    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    $timeoutMs = [Math]::Max(250, [Math]::Min(800, $TimeoutSec * 1000))
    # Use only the lightweight ping endpoint. Do not call heavier state APIs during startup.
    return (Test-ReportBinderTcpHttpUrl $Url '/api/ping' $timeoutMs)
}

function Test-ServerStartedMessage([string]$StdOutPath, [int]$Port) {
    if ([string]::IsNullOrWhiteSpace($StdOutPath) -or -not (Test-Path -LiteralPath $StdOutPath)) { return $false }
    try {
        $text = Get-Content -LiteralPath $StdOutPath -Raw -ErrorAction Stop
        return ([string]$text -match "ReportBinder local server started on 127\.0\.0\.1:$Port")
    } catch {
        return $false
    }
}

function Wait-ReportBinderReady([string]$Url, $Process, [string]$StdOutPath, [int]$Port, [int]$TimeoutMs) {
    $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
    $startupMessageLogged = $false
    while ([DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 150
        if ($null -ne $Process) {
            try {
                $Process.Refresh()
                if ($Process.HasExited) {
                    Add-LaunchLog "Server process exited during startup. ExitCode=$($Process.ExitCode)"
                    return $false
                }
            } catch {}
        }
        if (-not $startupMessageLogged -and (Test-ServerStartedMessage $StdOutPath $Port)) {
            Add-LaunchLog "Server startup message detected. Waiting for lightweight ping response."
            $startupMessageLogged = $true
        }
        if (Test-ReportBinderUrl $Url 1) { return $true }
    }
    return $false
}

function ConvertTo-HtmlText([string]$Value) {
    if ($null -eq $Value) { return '' }
    return $Value.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;').Replace('"','&quot;')
}

function ConvertTo-JsStringLiteral([string]$Value) {
    if ($null -eq $Value) { $Value = '' }
    $escaped = $Value.Replace('\','\\').Replace('"','\"').Replace("`r",'\r').Replace("`n",'\n').Replace('<','\u003c').Replace('>','\u003e').Replace('&','\u0026')
    return '"' + $escaped + '"'
}

function Write-StartupWaitPage([string]$Path, [string]$AppUrl, [string]$ReadyImageUrl, [string]$ShortcutPath, [string]$LatestLogPath) {
    $appJs = ConvertTo-JsStringLiteral $AppUrl
    $readyJs = ConvertTo-JsStringLiteral $ReadyImageUrl
    $shortcutHtml = ConvertTo-HtmlText $ShortcutPath
    $logHtml = ConvertTo-HtmlText $LatestLogPath
    $directHtml = ConvertTo-HtmlText $AppUrl
    $appHref = ConvertTo-HtmlText $AppUrl
    $html = @"
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ReportBinder 起動中</title>
<style>
body{font-family:"BIZ UDPGothic","Yu Gothic",Meiryo,sans-serif;margin:0;background:#f6f7f9;color:#1f2937;}
main{max-width:720px;margin:12vh auto;padding:32px;border-radius:18px;background:white;box-shadow:0 12px 32px rgba(15,23,42,.12);}
h1{font-size:24px;margin:0 0 12px;}p{font-size:16px;line-height:1.8;margin:8px 0;}.muted{color:#6b7280}.spinner{width:28px;height:28px;border:4px solid #e5e7eb;border-top-color:#6b7280;border-radius:50%;animation:spin 1s linear infinite;margin-bottom:16px;}@keyframes spin{to{transform:rotate(360deg)}}code{word-break:break-all;background:#f3f4f6;padding:2px 5px;border-radius:5px;}.late{display:none;margin-top:18px;padding:14px;border:1px solid #f59e0b;border-radius:12px;background:#fffbeb;}.late.show{display:block;}.actions{margin:16px 0 4px}.button{display:inline-block;padding:10px 16px;border-radius:10px;background:#111827;color:#fff;text-decoration:none;font-weight:700}.button:hover{background:#374151;}
</style>
</head>
<body>
<main>
<div class="spinner"></div>
<h1>ReportBinder を起動しています</h1>
<p>ローカルサーバーの準備ができたら、自動で画面に切り替わります。</p>
<p class="muted">経過: <span id="elapsed">0</span> 秒</p>
<div class="actions"><a class="button" href="$appHref">準備ができたら画面を開く</a></div>
<div id="late" class="late">
<p>起動に時間がかかっています。もう少し待つと自動で切り替わります。</p>
<p>この画面のまま止まる場合は、上の「準備ができたら画面を開く」ボタン、または以下のEdge起動用ファイルを開いてください。</p>
<p><code>$shortcutHtml</code></p>
<p class="muted">起動ログ: <code>$logHtml</code></p>
<p class="muted">接続先: <code>$directHtml</code></p>
</div>
</main>
<script>
(function(){
  var appUrl = $appJs;
  var readyUrl = $readyJs;
  var started = Date.now();
  var elapsedNode = document.getElementById('elapsed');
  var lateNode = document.getElementById('late');
  function updateElapsed(){
    var elapsed = Math.floor((Date.now() - started) / 1000);
    elapsedNode.textContent = String(elapsed);
    if (elapsed >= 15) lateNode.className = 'late show';
  }
  function go(){ window.location.replace(appUrl); }
  function retry(){ setTimeout(probe, 700); }
  function probe(){
    updateElapsed();
    var url = readyUrl + (readyUrl.indexOf('?') >= 0 ? '&' : '?') + '_=' + Date.now();
    var settled = false;
    var done = function(ok){
      if (settled) return;
      settled = true;
      if (ok) go(); else retry();
    };
    if (window.fetch) {
      try {
        fetch(url, {cache:'no-store', mode:'cors'}).then(function(res){ done(!!res && res.ok); }).catch(function(){
          // Some browsers restrict file:// -> http:// fetch. Fall back to image loading.
          imageProbe(url, done);
        });
        setTimeout(function(){ if (!settled) imageProbe(url, done); }, 1200);
        return;
      } catch(e) {}
    }
    imageProbe(url, done);
  }
  function imageProbe(url, done){
    var img = new Image();
    var timer = setTimeout(function(){ try { img.src = ''; } catch(e) {} done(false); }, 1800);
    img.onload = function(){ clearTimeout(timer); done(true); };
    img.onerror = function(){ clearTimeout(timer); done(false); };
    img.src = url;
  }
  probe();
})();
</script>
</body>
</html>
"@
    Set-Content -LiteralPath $Path -Value $html -Encoding UTF8
}


function Open-EdgeBrowser([string]$Url) {
    # ReportBinder must open Microsoft Edge explicitly. Do not use raw URL / explorer / rundll32
    # fallback because older Windows associations can open Internet Explorer.
    if ([string]::IsNullOrWhiteSpace($Url)) { return $false }
    $target = ConvertTo-EdgeOpenTarget $Url
    $edge = Get-EdgeExecutablePath

    if (-not [string]::IsNullOrWhiteSpace($edge)) {
        try {
            Start-Process -FilePath $edge -ArgumentList @($target) -ErrorAction Stop | Out-Null
            Add-LaunchLog "Browser target opened by Microsoft Edge. Edge=$edge Target=$target"
            return $true
        } catch { Add-LaunchLog ("Browser open via Microsoft Edge failed: " + $_.Exception.Message) }
    } else {
        Add-LaunchLog "Microsoft Edge executable path was not resolved from PATH, registry, or standard locations. Trying protocol/command fallbacks."
    }

    if ($target -match '^https?://') {
        try {
            Start-Process -FilePath ("microsoft-edge:" + $target) -ErrorAction Stop | Out-Null
            Add-LaunchLog "Browser opened by microsoft-edge protocol fallback."
            return $true
        } catch { Add-LaunchLog ("Browser open via microsoft-edge protocol failed: " + $_.Exception.Message) }
    }

    try {
        $cmdExe = Join-Path $env:SystemRoot 'System32\cmd.exe'
        if (-not (Test-Path -LiteralPath $cmdExe)) { $cmdExe = 'cmd.exe' }
        $escapedTarget = $target -replace '"','""'
        $cmdArgs = '/c start "" msedge.exe "' + $escapedTarget + '"'
        Start-Process -FilePath $cmdExe -ArgumentList $cmdArgs -WindowStyle Hidden -ErrorAction Stop | Out-Null
        Add-LaunchLog "Browser opened by msedge command fallback."
        return $true
    } catch { Add-LaunchLog ("Browser open via msedge command fallback failed: " + $_.Exception.Message) }

    return $false
}

function Write-ServerLogsIntoStartupLog([string]$StdOutPath, [string]$StdErrPath) {
    try {
        if (Test-Path -LiteralPath $StdOutPath) {
            Add-LaunchLog "--- server stdout ---"
            Add-Content -LiteralPath $log -Value (Get-Content -LiteralPath $StdOutPath -Raw -ErrorAction SilentlyContinue) -Encoding UTF8
        }
        if (Test-Path -LiteralPath $StdErrPath) {
            Add-LaunchLog "--- server stderr ---"
            Add-Content -LiteralPath $log -Value (Get-Content -LiteralPath $StdErrPath -Raw -ErrorAction SilentlyContinue) -Encoding UTF8
        }
        Sync-LatestLog
    } catch {}
}

function Get-ProcessCommandLineText($ProcessInfo) {
    try { return [string]$ProcessInfo.CommandLine } catch { return '' }
}

function CommandLine-ContainsPath([string]$CommandLine, [string]$Path) {
    if ([string]::IsNullOrWhiteSpace($CommandLine) -or [string]::IsNullOrWhiteSpace($Path)) { return $false }
    $cmd = $CommandLine.ToLowerInvariant()
    $p1 = ([IO.Path]::GetFullPath($Path)).ToLowerInvariant()
    $p2 = ($p1 -replace '\\','/')
    return ($cmd.Contains($p1) -or $cmd.Contains($p2))
}

function Stop-StaleUiServerProcesses([string]$ServerPath, [string]$TargetMode) {
    # V5-P3: 以前は server.ps1 を含むプロセスを無条件に停止していたため、
    # 日本語サーバーを起動すると稼働中の英語サーバーまで落ちていた(逆も同様)。
    # 停止するのは「同じ -Mode の UI サーバー」だけに限定する。
    # -RenderJobPath(PDF作成の子)、-DiffJobPath(差分画像の子)、
    # -AutoSchedulerPath(自動処理の子) は対象外。
    $modePattern = ('\s-mode\s+{0}(\s|$)' -f [regex]::Escape([string]$TargetMode))
    try {
        $stale = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
            $cmd = Get-ProcessCommandLineText $_
            $cmdLower = $cmd.ToLowerInvariant()
            (CommandLine-ContainsPath $cmd $ServerPath) -and
            ($cmdLower -notmatch '\s-renderjobpath\b') -and
            ($cmdLower -notmatch '\s-diffjobpath\b') -and
            ($cmdLower -notmatch '\s-autoschedulerpath\b') -and
            ($cmdLower -match $modePattern)
        })
        foreach ($p in $stale) {
            try {
                Add-LaunchLog "Stopping stale ReportBinder UI server PID=$($p.ProcessId) before new startup."
                Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    } catch {
        Add-LaunchLog ("Stale server cleanup skipped: " + $_.Exception.Message)
    }
}

function Open-ExistingIfRunning([string]$Mode, [string]$UrlFile, [string]$RootUrlFile) {
    foreach ($candidatePath in @($UrlFile, $RootUrlFile)) {
        $existingUrl = Read-UrlShortcut $candidatePath
        if (Test-ReportBinderUrl $existingUrl 1) {
            Add-LaunchLog "Existing ReportBinder server detected. Opening existing browser URL."
            Write-UrlShortcut $UrlFile $existingUrl
            Write-UrlShortcut $RootUrlFile $existingUrl
            $existingEdgeCmdFile = Join-Path (Split-Path -Parent $RootUrlFile) ("ReportBinder-$Mode-Edge.cmd")
            Write-EdgeOpenCommand $existingEdgeCmdFile $existingUrl
            if (-not (Open-EdgeBrowser $existingUrl)) {
                Add-LaunchLog "Could not open Microsoft Edge automatically. Use Edge command: $existingEdgeCmdFile"
                if (-not $Diagnostics) {
                    try {
                        $popup = New-Object -ComObject WScript.Shell
                        [void]$popup.Popup("ReportBinder は起動済みですが、Microsoft Edgeで自動表示できませんでした。`n`n以下を開いてください。`n$existingEdgeCmdFile", 0, "ReportBinder", 48)
                    } catch {}
                }
            }
            Add-LaunchLog "Existing session opened: $existingUrl"
            Sync-LatestLog
            if ($Diagnostics) {
                Write-Host "ReportBinder is already running."
                Write-Host "Mode: $Mode"
                Write-Host "URL: $existingUrl"
                Write-Host "Shortcut: $RootUrlFile"
                Write-Host "Log: $latest"
            }
            exit 0
        }
    }
}

try {
    # Windows can deliver two launcher invocations for one double-click, especially
    # when the shortcut lives on a shared folder. Only one launcher may decide
    # whether to start/open the UI for this app root and language.
    $mutexName = "Local\ReportBinder.Launch.$appRootKey.$Mode"
    $createdNew = $false
    $launchMutex = New-Object System.Threading.Mutex($false, $mutexName, [ref]$createdNew)
    try {
        $launchMutexOwned = $launchMutex.WaitOne(0, $false)
    } catch [System.Threading.AbandonedMutexException] {
        $launchMutexOwned = $true
    }
    if (-not $launchMutexOwned) {
        Add-LaunchLog "Another launcher is already handling startup. This duplicate invocation will exit without opening a tab."
        exit 0
    }

    Set-Content -LiteralPath $launchPidFile -Value ([string]$PID) -Encoding ASCII
    Add-LaunchLog "ReportBinder launcher start. Mode=$Mode"
    Add-LaunchLog "AppRoot=$script:AppRoot"
    Add-LaunchLog "PowerShell=$($PSVersionTable.PSVersion)"

    $server = Join-Path $script:AppRoot 'server.ps1'
    if (-not (Test-Path -LiteralPath $server)) { throw "server.ps1 was not found: $server" }

    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }

    Open-ExistingIfRunning $Mode $urlFile $rootUrlFile

    # If the saved URL is dead, remove unusable UI server processes from this app before starting one clean instance.
    # Render-job processes are excluded so PDF creation can continue after the screen is closed.
    Stop-StaleUiServerProcesses $server $Mode

    $port = Get-FreePort
    $token = New-SessionToken
    $url = "http://127.0.0.1:$port/?token=$token&mode=$Mode"
    Write-UrlShortcut $urlFile $url
    Write-UrlShortcut $rootUrlFile $url
    Write-EdgeOpenCommand $edgeCmdFile $url

    Add-LaunchLog "SelectedPort=$port"
    Add-LaunchLog "URL=$url"
    Add-LaunchLog "URLShortcut=$rootUrlFile"
    Add-LaunchLog "StartupWaitPage=$waitPageFile"
    Add-LaunchLog "ServerStdOut=$serverOut"
    Add-LaunchLog "ServerStdErr=$serverErr"

    $serverQuoted = Quote-ProcessArgument $server
    $argLine = "-NoProfile -ExecutionPolicy Bypass -File $serverQuoted -Mode $Mode -Port $port -Token $token -NoOpen"
    Add-LaunchLog "Starting server.ps1 as a background process. Browser startup wait page will be opened immediately."
    # サーバーをコンソールウィンドウ表示で起動する。ウィンドウを閉じればプロセスごと終了できる。
    # (目立たせたくない場合は Normal を Minimized に、完全に隠す場合は Hidden に変更)
    $proc = Start-Process -FilePath $psExe -ArgumentList $argLine -WindowStyle Normal -PassThru -RedirectStandardOutput $serverOut -RedirectStandardError $serverErr
    Set-Content -LiteralPath $serverPidFile -Value ([string]$proc.Id) -Encoding ASCII
    Add-LaunchLog "ServerProcessId=$($proc.Id)"

    # コンソールウィンドウの描画が先に完了してからブラウザを前面に出す。
    Start-Sleep -Milliseconds 100

    $readyImageUrl = Get-ReportBinderApiUrl $url '/api/ready.gif'
    Write-StartupWaitPage $waitPageFile $url $readyImageUrl $edgeCmdFile $latest
    Add-LaunchLog "Opening startup wait page immediately. It will switch to ReportBinder when the local server is ready."
    $waitPageOpened = Open-EdgeBrowser $waitPageFile
    if (-not $waitPageOpened) {
        Add-LaunchLog "Could not open startup wait page in Microsoft Edge automatically. Will open the real ReportBinder URL after readiness. Edge command: $edgeCmdFile"
        if (-not $Diagnostics) {
            try {
                $popup = New-Object -ComObject WScript.Shell
                [void]$popup.Popup("ReportBinder の起動画面をMicrosoft Edgeで自動表示できませんでした。`nサーバー準備後に本体画面を開きます。`n`n自動表示されない場合は以下を開いてください。`n$edgeCmdFile", 0, "ReportBinder", 48)
            } catch {}
        }
    }

    $readyWaitStarted = [DateTime]::UtcNow
    $ready = Wait-ReportBinderReady $url $proc $serverOut $port 45000
    $readyWaitMs = [int](([DateTime]::UtcNow - $readyWaitStarted).TotalMilliseconds)
    if (-not $ready) {
        try { $proc.Refresh() } catch {}
        if ($null -ne $proc -and $proc.HasExited) {
            Write-ServerLogsIntoStartupLog $serverOut $serverErr
            throw "server.ps1 exited during startup. Check $latest and $serverErr"
        }
        if (-not $waitPageOpened) {
            Add-LaunchLog "Startup wait page was not opened and readiness timed out; opening ReportBinder URL directly as a fallback. URL=$url"
            [void](Open-EdgeBrowser $url)
        } else {
            Add-LaunchLog "Startup wait page is open; browser-side readiness polling will continue. URL=$url"
        }
        Sync-LatestLog
        if ($Diagnostics) {
            Write-Host "ReportBinder startup wait page opened."
            Write-Host "Mode: $Mode"
            Write-Host "URL: $url"
            Write-Host "Wait page: $waitPageFile"
            Write-Host "Shortcut: $rootUrlFile"
            Write-Host "Log: $latest"
            Write-Host "Server stdout: $serverOut"
            Write-Host "Server stderr: $serverErr"
        }
        exit 0
    }

    Add-LaunchLog "Server is ready. Startup wait page should redirect to ReportBinder shortly. ReadyWaitMs=$readyWaitMs URL=$url"
    # The wait page owns the redirect once Edge accepted it. Opening the real URL
    # as a delayed rescue as well races with that redirect and creates two tabs.
    if (-not $waitPageOpened) {
        Add-LaunchLog "Startup wait page was not opened; opening ReportBinder URL directly now that the server is ready."
        [void](Open-EdgeBrowser $url)
    }
    Sync-LatestLog
    if ($Diagnostics) {
        Write-Host "ReportBinder started."
        Write-Host "Mode: $Mode"
        Write-Host "URL: $url"
        Write-Host "Wait page: $waitPageFile"
        Write-Host "Shortcut: $rootUrlFile"
        Write-Host "Log: $latest"
        Write-Host "Server stdout: $serverOut"
        Write-Host "Server stderr: $serverErr"
    }
    exit 0
} catch {
    $msg = $_.Exception.Message
    Add-Content -LiteralPath $log -Value ("ERROR: " + $msg) -Encoding UTF8
    Add-Content -LiteralPath $log -Value ([string]$_) -Encoding UTF8
    Sync-LatestLog
    if (-not $Diagnostics) {
        try {
            $popup = New-Object -ComObject WScript.Shell
            [void]$popup.Popup("ReportBinder failed to start.`n`nLog:`n$latest", 0, "ReportBinder", 16)
        } catch {}
    }
    if ($Diagnostics) {
        Write-Host "ReportBinder failed to start."
        Write-Host "Log: $latest"
        Write-Host $msg
    }
    exit 1
} finally {
    try { Remove-Item -LiteralPath $launchPidFile -Force -ErrorAction SilentlyContinue } catch {}
    if ($launchMutexOwned -and $null -ne $launchMutex) {
        try { $launchMutex.ReleaseMutex() } catch {}
    }
    if ($null -ne $launchMutex) {
        try { $launchMutex.Dispose() } catch {}
    }
}
