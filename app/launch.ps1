param(
    [switch]$Diagnostics,
    [switch]$LocalRuntime,
    [string]$SharedAppRoot = ''
)

$ErrorActionPreference = 'Stop'
$script:AppRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($SharedAppRoot)) { $SharedAppRoot = $script:AppRoot }
$script:SharedAppRoot = [IO.Path]::GetFullPath($SharedAppRoot)
$script:BootstrapWarning = ''
$script:IntegrityFailure = ''
$runtimeInfoPath = Join-Path $script:AppRoot 'runtime-version.json'
$script:RuntimeVersion = 'legacy'
try {
    if (Test-Path -LiteralPath $runtimeInfoPath) {
        $runtimeInfo = Get-Content -LiteralPath $runtimeInfoPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $candidateVersion = ([string]$runtimeInfo.version).Trim()
        if ($candidateVersion -match '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$') { $script:RuntimeVersion = $candidateVersion }
    }
} catch { }

# 完全性マニフェストの除外規則。package-release.ps1 に同じ関数があり、両者が一致して
# いることを selfcheck.py が検査する。片方だけを変えると検証が素通りする。
function Test-IntegrityExcludedPath([string]$RelativePath) {
    if ($RelativePath -eq 'integrity-manifest.json') { return $true }
    if ($RelativePath -eq 'config.json') { return $true }
    $top = ($RelativePath -split '/')[0]
    return ($top -in @('logs', 'thirdparty-cache'))
}

# 共有フォルダーからコピーしたツリーを、%LOCALAPPDATA% へ確定する前に照合する。
# SMB越しのコピー欠落・切り詰めと、ツリーの一部だけを書き換えられた場合を検出する。
# 共有全体を書き換えられる相手はこの検証自体を無効化できるため、配布用共有は
# 発行者以外を読み取り専用にすること（docs/OPERATIONS_GUIDE.md に記載）。
function Test-StagedTreeIntegrity([string]$StageAppRoot) {
    $manifestPath = Join-Path $StageAppRoot 'integrity-manifest.json'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        return [pscustomobject]@{ Verified = $false; Reason = 'no-manifest'; Message = '完全性マニフェストがない配布物です。' }
    }
    $manifest = $null
    try { $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }
    if ($null -eq $manifest -or $null -eq $manifest.files) {
        return [pscustomobject]@{ Verified = $false; Reason = 'broken-manifest'; Message = '完全性マニフェストを読み取れません。' }
    }
    $expected = @{}
    foreach ($entry in $manifest.files.PSObject.Properties) { $expected[[string]$entry.Name] = $entry.Value }
    $prefix = [IO.Path]::GetFullPath($StageAppRoot).TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($file in @(Get-ChildItem -LiteralPath $StageAppRoot -Recurse -File -Force)) {
        $relative = $file.FullName.Substring($prefix.Length) -replace '\\', '/'
        if (Test-IntegrityExcludedPath $relative) { continue }
        if (-not $expected.ContainsKey($relative)) {
            return [pscustomobject]@{ Verified = $false; Reason = 'unexpected-file'; Message = "配布物にない余分なファイルがあります: $relative" }
        }
        [void]$seen.Add($relative)
        # サイズ照合で先に落とすと、改変されたファイルのハッシュ計算を省ける。
        if ([int64]$expected[$relative].size -ne [int64]$file.Length) {
            return [pscustomobject]@{ Verified = $false; Reason = 'size-mismatch'; Message = "ファイルのサイズが配布物と一致しません: $relative" }
        }
        $actual = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($actual -ne ([string]$expected[$relative].sha256).ToLowerInvariant()) {
            return [pscustomobject]@{ Verified = $false; Reason = 'hash-mismatch'; Message = "ファイルの内容が配布物と一致しません: $relative" }
        }
    }
    foreach ($relative in $expected.Keys) {
        if (-not $seen.Contains($relative)) {
            return [pscustomobject]@{ Verified = $false; Reason = 'missing-file'; Message = "配布物のファイルがコピーされていません: $relative" }
        }
    }
    return [pscustomobject]@{ Verified = $true; Reason = 'ok'; Message = ("{0} ファイルを照合しました。" -f $seen.Count); FileCount = $seen.Count }
}

# The shared launcher is intentionally thin. It reads one small version file, installs
# that immutable app version under LocalAppData when necessary, then restarts this
# launcher from the local copy. Heavy PowerShell, web, Java and PDFBox files are never
# loaded from SMB during normal startup.
if (-not $LocalRuntime -and $script:RuntimeVersion -ne 'legacy') {
    $installMutex = $null
    $installOwned = $false
    $stageRoot = ''
    try {
        $runtimeBase = Join-Path (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ReportBinder\runtime\versions') $script:RuntimeVersion
        $localAppRoot = Join-Path $runtimeBase 'app'
        $installedMarker = Join-Path $runtimeBase 'installed.json'
        $ready = (Test-Path -LiteralPath $installedMarker -PathType Leaf) -and
                 (Test-Path -LiteralPath (Join-Path $localAppRoot 'server.ps1') -PathType Leaf) -and
                 (Test-Path -LiteralPath (Join-Path $localAppRoot 'web\app.js') -PathType Leaf)
        if (-not $ready) {
            $mutexName = 'Local\ReportBinder.RuntimeInstall.' + $script:RuntimeVersion
            $created = $false
            $installMutex = New-Object System.Threading.Mutex($false, $mutexName, [ref]$created)
            try { $installOwned = $installMutex.WaitOne(120000, $false) }
            catch [System.Threading.AbandonedMutexException] { $installOwned = $true }
            if (-not $installOwned) { throw 'ローカル版アプリの更新待ちがタイムアウトしました。' }

            $ready = (Test-Path -LiteralPath $installedMarker -PathType Leaf) -and
                     (Test-Path -LiteralPath (Join-Path $localAppRoot 'server.ps1') -PathType Leaf)
            if (-not $ready) {
                $runtimeParent = Split-Path -Parent $runtimeBase
                if (-not (Test-Path -LiteralPath $runtimeParent)) { New-Item -ItemType Directory -Path $runtimeParent -Force | Out-Null }
                if (Test-Path -LiteralPath $runtimeBase) { Remove-Item -LiteralPath $runtimeBase -Recurse -Force }
                $stageRoot = $runtimeBase + '.staging-' + ([Guid]::NewGuid().ToString('N'))
                $stageApp = Join-Path $stageRoot 'app'
                New-Item -ItemType Directory -Path $stageApp -Force | Out-Null
                foreach ($entry in @(Get-ChildItem -LiteralPath $script:AppRoot -Force)) {
                    if ($entry.Name -in @('logs','thirdparty-cache','config.json')) { continue }
                    Copy-Item -LiteralPath $entry.FullName -Destination $stageApp -Recurse -Force
                }
                foreach ($required in @('launch.ps1','server.ps1','runtime-version.json','web\app.js')) {
                    if (-not (Test-Path -LiteralPath (Join-Path $stageApp $required) -PathType Leaf)) {
                        throw "ローカル版アプリのコピーが不完全です: $required"
                    }
                }
                # 改変・破損したツリーを %LOCALAPPDATA% へ確定してしまうと、以後の起動は
                # そちらを使い続ける。確定の前に照合し、一致しない場合は staging を捨てる。
                $integrity = Test-StagedTreeIntegrity $stageApp
                if (-not $integrity.Verified) {
                    if ($integrity.Reason -eq 'no-manifest') {
                        # 旧配布物と開発ツリーからの起動を壊さないため、警告に留める。
                        $script:BootstrapWarning = '配布物の完全性を確認できませんでした（' + $integrity.Message + '）'
                    } else {
                        # 完全性の不一致は可用性より優先する。共有コピーへのフォールバックは
                        # 改変されたツリーをそのまま実行することになるため、起動を止める。
                        $script:IntegrityFailure = $integrity.Message
                        throw ('配布物の内容が壊れているか、書き換えられています。' + $integrity.Message)
                    }
                }
                [ordered]@{
                    version = $script:RuntimeVersion
                    source = $script:SharedAppRoot
                    installedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz')
                    integrityVerified = [bool]$integrity.Verified
                    integrityFileCount = [int]$integrity.FileCount
                } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $stageRoot 'installed.json') -Encoding UTF8
                Move-Item -LiteralPath $stageRoot -Destination $runtimeBase
                $stageRoot = ''
            }
        }

        $psExeBootstrap = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        if (-not (Test-Path -LiteralPath $psExeBootstrap)) { $psExeBootstrap = 'powershell.exe' }
        $localLauncher = Join-Path $localAppRoot 'launch.ps1'
        $qLauncher = '"' + ($localLauncher -replace '"','\"') + '"'
        $qShared = '"' + ($script:SharedAppRoot -replace '"','\"') + '"'
        $bootstrapArgs = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File $qLauncher -LocalRuntime -SharedAppRoot $qShared"
        if ($Diagnostics) { $bootstrapArgs += ' -Diagnostics' }
        Start-Process -FilePath $psExeBootstrap -ArgumentList $bootstrapArgs -WindowStyle Hidden | Out-Null
        exit 0
    } catch {
        $script:BootstrapWarning = $_.Exception.Message
        # Availability wins over speed if the local install cannot be prepared.
        # Continue from the shared copy for this launch and leave a diagnostic locally.
        # 完全性の不一致だけは例外。共有コピーで続行すると改変されたツリーを実行するため、
        # 利用者に理由を提示して起動を中止する。
        if (-not [string]::IsNullOrWhiteSpace($script:IntegrityFailure)) {
            $stopMessage = "ReportBinderを起動できません。`n`n" +
                "配布物の内容が、配布時の記録と一致しません。破損したコピー、または第三者による書き換えの可能性があります。`n`n" +
                $script:IntegrityFailure + "`n`n" +
                '共有フォルダーの管理者に連絡し、配布物を作り直してもらってください。'
            # Add-LaunchLog はこの時点ではまだ定義・初期化されていないため、直接書く。
            try {
                $bootLogDir = Join-Path (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ReportBinder') 'logs'
                if (-not (Test-Path -LiteralPath $bootLogDir)) { New-Item -ItemType Directory -Path $bootLogDir -Force | Out-Null }
                Add-Content -LiteralPath (Join-Path $bootLogDir 'integrity-latest.log') -Encoding UTF8 -Value (
                    '{0} integrity check failed ({1}): {2} source={3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),
                    $integrity.Reason, $script:IntegrityFailure, $script:SharedAppRoot)
            } catch { }
            try {
                Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
                [void][System.Windows.Forms.MessageBox]::Show($stopMessage, 'ReportBinder',
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            } catch { Write-Error $stopMessage }
            if ($stageRoot -and (Test-Path -LiteralPath $stageRoot)) { Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue }
            exit 1
        }
    } finally {
        if ($stageRoot -and (Test-Path -LiteralPath $stageRoot)) { Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue }
        if ($installOwned -and $null -ne $installMutex) { try { $installMutex.ReleaseMutex() } catch { } }
        if ($null -ne $installMutex) { try { $installMutex.Dispose() } catch { } }
    }
}

$script:RootDir = Split-Path -Parent $script:AppRoot
$logDir = Join-Path (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ReportBinder') 'logs'
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
        $hashBytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($script:SharedAppRoot.ToLowerInvariant()))
        $appRootKey = -join ($hashBytes[0..7] | ForEach-Object { $_.ToString('x2') })
    } finally { try { $sha.Dispose() } catch {} }
} catch {
    try { $appRootKey = ([Guid]::NewGuid().ToString('N')).Substring(0, 16) } catch { $appRootKey = 'default' }
}
$localLaunchDir = Join-Path ([IO.Path]::GetTempPath()) ("ReportBinder-$appRootKey")
if (-not (Test-Path -LiteralPath $localLaunchDir)) { New-Item -ItemType Directory -Path $localLaunchDir -Force | Out-Null }

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$log = Join-Path $logDir ("startup-workspace-$stamp.log")
$latest = Join-Path $logDir 'startup-workspace-latest.log'
$serverOut = Join-Path $logDir ("server-workspace-$stamp.out.log")
$serverErr = Join-Path $logDir ("server-workspace-$stamp.err.log")
$urlFile = Join-Path $logDir 'open-workspace-latest.url'
$rootUrlFile = Join-Path $logDir 'ReportBinder-workspace.url'
$serverPidFile = Join-Path $logDir 'server-workspace-latest.pid'
$launchPidFile = Join-Path $logDir 'launch-workspace-latest.pid'
$waitPageFile = Join-Path $localLaunchDir 'startup-wait-workspace-latest.html'
$edgeCmdFile = Join-Path $logDir 'ReportBinder-workspace-Edge.cmd'
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
            $versionOk = ($script:RuntimeVersion -eq 'legacy') -or ($partial -match ('"runtimeVersion"\s*:\s*"' + [regex]::Escape($script:RuntimeVersion) + '"'))
            if ($partial -match '^HTTP/1\.[01]\s+200\s' -and $partial -match '"ok"\s*:\s*true' -and $versionOk) { return $true }
        }
        $text = [Text.Encoding]::UTF8.GetString($ms.ToArray())
        $versionOk = ($script:RuntimeVersion -eq 'legacy') -or ($text -match ('"runtimeVersion"\s*:\s*"' + [regex]::Escape($script:RuntimeVersion) + '"'))
        return ($text -match '^HTTP/1\.[01]\s+200\s' -and $text -match '"ok"\s*:\s*true' -and $versionOk)
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

function ConvertTo-JsStringLiteral([string]$Value) {
    if ($null -eq $Value) { $Value = '' }
    $escaped = $Value.Replace('\','\\').Replace('"','\"').Replace("`r",'\r').Replace("`n",'\n').Replace('<','\u003c').Replace('>','\u003e').Replace('&','\u0026')
    return '"' + $escaped + '"'
}

function Write-StartupWaitPage([string]$Path, [string]$AppUrl, [string]$ReadyImageUrl) {
    $appJs = ConvertTo-JsStringLiteral $AppUrl
    $readyJs = ConvertTo-JsStringLiteral $ReadyImageUrl
    $html = @"
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ReportBinder 起動中</title>
<style>
*{box-sizing:border-box}body{font-family:"Segoe UI","BIZ UDPGothic","BIZ UDPゴシック","Yu Gothic UI",Meiryo,sans-serif;font-size:16px;margin:0;background:#f6f7f8;color:#1b1c1e}
.loading-screen{position:fixed;inset:0;display:grid;place-items:center;padding:24px}
main{width:min(360px,calc(100vw - 48px));display:grid;justify-items:center;gap:8px;padding:30px 28px;border:1px solid #e6e7ea;border-radius:12px;background:#fff;box-shadow:0 4px 16px rgba(27,28,30,.10),0 0 0 1px rgba(27,28,30,.05);text-align:center}
.mark{width:30px;height:30px;margin-bottom:4px;border-radius:9px;background:#5e6ad2;box-shadow:inset 0 0 0 8px #f1f2fb}
h1{margin:0;font-size:18px;line-height:1.45}p{margin:0;color:#6b6f76;font-size:15px;line-height:1.55}
.spinner{width:22px;height:22px;margin-top:10px;border:2px solid #d5d7dc;border-top-color:#5e6ad2;border-radius:50%;animation:spin .8s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}
.late{display:none;margin-top:10px;padding-top:12px;border-top:1px solid #e6e7ea}.late.show{display:block}.late p{font-size:14px}
@media(prefers-reduced-motion:reduce){.spinner{animation:none;border-top-color:#d5d7dc}}
</style>
</head>
<body>
<section class="loading-screen">
<main role="status" aria-live="polite" aria-labelledby="loading-title">
<span class="mark" aria-hidden="true"></span>
<h1 id="loading-title">ReportBinderを準備しています</h1>
<p id="loading-message">アプリを起動しています。</p>
<span class="spinner" aria-hidden="true"></span>
<div id="late" class="late"><p>初回起動や更新後は時間がかかることがあります。このタブを閉じずにお待ちください。</p></div>
</main>
</section>
<script>
(function(){
  var appUrl = $appJs;
  var readyUrl = $readyJs;
  var started = Date.now();
  var statusNode = document.getElementById('loading-message');
  var lateNode = document.getElementById('late');
  function updateStatus(){
    var elapsed = Math.floor((Date.now() - started) / 1000);
    if (elapsed >= 20) {
      statusNode.textContent = '通常より時間がかかっています。';
      lateNode.className = 'late show';
    }
  }
  function go(){ window.location.replace(appUrl); }
  function retry(){ setTimeout(probe, 700); }
  function probe(){
    updateStatus();
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

function Stop-StaleUiServerProcesses([string]$ServerPath) {
    # 単一ワークスペースのUIサーバーだけを起動し直す。
    # -RenderJobPath(PDF作成の子)、-DiffJobPath(差分画像の子)、
    # -AutoSchedulerPath(自動処理の子) は対象外。
    try {
        $stale = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
            $cmd = Get-ProcessCommandLineText $_
            $cmdLower = $cmd.ToLowerInvariant()
            ((CommandLine-ContainsPath $cmd $ServerPath) -or
             ($cmdLower -match '[\\/]reportbinder[\\/]runtime[\\/]versions[\\/][^\\/]+[\\/]app[\\/]server\.ps1')) -and
            ($cmdLower -notmatch '\s-renderjobpath\b') -and
            ($cmdLower -notmatch '\s-diffjobpath\b') -and
            ($cmdLower -notmatch '\s-autoschedulerpath\b')
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

function Open-ExistingIfRunning([string]$UrlFile, [string]$RootUrlFile) {
    foreach ($candidatePath in @($UrlFile, $RootUrlFile)) {
        $existingUrl = Read-UrlShortcut $candidatePath
        if (Test-ReportBinderUrl $existingUrl 1) {
            Add-LaunchLog "Existing ReportBinder server detected. Opening existing browser URL."
            Write-UrlShortcut $UrlFile $existingUrl
            Write-UrlShortcut $RootUrlFile $existingUrl
            $existingEdgeCmdFile = Join-Path (Split-Path -Parent $RootUrlFile) 'ReportBinder-workspace-Edge.cmd'
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
    # whether to start/open the UI for this app root and workspace.
    $mutexName = "Local\ReportBinder.Launch.$appRootKey.workspace"
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
    Add-LaunchLog "ReportBinder launcher start. Workspace=default"
    Add-LaunchLog "AppRoot=$script:AppRoot"
    Add-LaunchLog "SharedAppRoot=$script:SharedAppRoot"
    Add-LaunchLog "RuntimeVersion=$script:RuntimeVersion LocalRuntime=$LocalRuntime"
    if (-not [string]::IsNullOrWhiteSpace($script:BootstrapWarning)) {
        Add-LaunchLog ("Local runtime install failed; using shared copy for this launch: " + $script:BootstrapWarning)
    }
    Add-LaunchLog "PowerShell=$($PSVersionTable.PSVersion)"

    $server = Join-Path $script:AppRoot 'server.ps1'
    if (-not (Test-Path -LiteralPath $server)) { throw "server.ps1 was not found: $server" }

    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }

    Open-ExistingIfRunning $urlFile $rootUrlFile

    # If the saved URL is dead, remove unusable UI server processes from this app before starting one clean instance.
    # Render-job processes are excluded so PDF creation can continue after the screen is closed.
    Stop-StaleUiServerProcesses $server

    $port = Get-FreePort
    $token = New-SessionToken
    $url = "http://127.0.0.1:$port/?token=$token"
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
    $argLine = "-NoProfile -ExecutionPolicy Bypass -File $serverQuoted -Port $port -Token $token -NoOpen -SharedAppRoot $(Quote-ProcessArgument $script:SharedAppRoot)"
    Add-LaunchLog "Starting server.ps1 as a background process. Browser startup wait page will be opened immediately."
    # UIサーバーはバックグラウンドで動作する。診断情報はローカルログへ保存し、通常起動ではコンソールを表示しない。
    $proc = Start-Process -FilePath $psExe -ArgumentList $argLine -WindowStyle Hidden -PassThru -RedirectStandardOutput $serverOut -RedirectStandardError $serverErr
    Set-Content -LiteralPath $serverPidFile -Value ([string]$proc.Id) -Encoding ASCII
    Add-LaunchLog "ServerProcessId=$($proc.Id)"

    $readyImageUrl = Get-ReportBinderApiUrl $url '/api/ready.gif'
    Write-StartupWaitPage $waitPageFile $url $readyImageUrl
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
