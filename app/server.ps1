param(
    [ValidateSet('ja','en')]
    [string]$Mode = 'ja',
    [int]$Port = 0,
    [string]$Token = '',
    [switch]$NoOpen,
    [string]$RenderJobPath = '',
    [string]$FinalJobPath = '',
    [string]$DiffJobPath = '',
    [string]$AutoSchedulerPath = '',
    [int]$ParentProcessId = 0
)

$ErrorActionPreference = 'Stop'

$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:WebRoot = Join-Path $Script:AppRoot 'web'
$Script:DefaultConfigPath = Join-Path $Script:AppRoot 'default-config.json'
$Script:RuntimeVersion = 'legacy'
try {
    $runtimeInfoPath = Join-Path $Script:AppRoot 'runtime-version.json'
    if (Test-Path -LiteralPath $runtimeInfoPath) {
        $runtimeInfo = Get-Content -LiteralPath $runtimeInfoPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $runtimeCandidate = ([string]$runtimeInfo.version).Trim()
        if ($runtimeCandidate -match '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$') { $Script:RuntimeVersion = $runtimeCandidate }
    }
} catch { }
$localConfigOverride = ([string]$env:REPORTBINDER_LOCAL_CONFIG_ROOT).Trim()
if ([string]::IsNullOrWhiteSpace($localConfigOverride)) {
    $Script:LocalConfigRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ReportBinder'
} else {
    $Script:LocalConfigRoot = [IO.Path]::GetFullPath($localConfigOverride)
}
$Script:ConfigPath = Join-Path $Script:LocalConfigRoot 'config.json'
$Script:LocalProjectsRoot = Join-Path $Script:LocalConfigRoot 'projects'
if (-not (Test-Path -LiteralPath $Script:LocalConfigRoot)) { New-Item -ItemType Directory -Path $Script:LocalConfigRoot -Force | Out-Null }
if (-not (Test-Path -LiteralPath $Script:LocalProjectsRoot)) { New-Item -ItemType Directory -Path $Script:LocalProjectsRoot -Force | Out-Null }
if ([string]::IsNullOrWhiteSpace($Token)) {
    $tokenBytes = New-Object byte[] 32
    $tokenRng = [Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $tokenRng.GetBytes($tokenBytes)
    } finally {
        $tokenRng.Dispose()
    }
    $Script:Token = [Convert]::ToBase64String($tokenBytes).TrimEnd('=').Replace('+','-').Replace('/','_')
} else {
    $Script:Token = $Token
}

$Script:ClientAttached = $false
$Script:LastHeartbeatUtc = [DateTime]::UtcNow
$Script:ClientCloseNotifiedUtc = [DateTime]::MinValue
$Script:ShutdownRequested = $false
$Script:CachedJavaExe = ''
# /api/state が読んだ共有フォルダー上のメタデータを、直後の比較画面でも再利用する。
# ファイル更新時刻・サイズまたは版IDが変われば別キー/再読込になる。
$Script:StructureReadCache = @{}
$Script:StructureReadCacheLimit = 4
$Script:LatestComparisonCache = @{}
$Script:LatestComparisonCacheLimit = 256
$Script:SnapshotManifestCache = @{}
$Script:VisualHashCache = @{}
$Script:SnapshotIdCache = @{}
$Script:SnapshotSummaryCache = @{}
$Script:DiffDetailResponseCache = @{}
$Script:LocalRuntimeCacheRoot = Join-Path $Script:LocalConfigRoot 'runtime-cache'
if (-not (Test-Path -LiteralPath $Script:LocalRuntimeCacheRoot)) { New-Item -ItemType Directory -Path $Script:LocalRuntimeCacheRoot -Force | Out-Null }
# 最小化コンソールで起動されたとき、何のウィンドウか分かるようにタイトルを付ける。
try { $host.UI.RawUI.WindowTitle = "ReportBinder サーバー ($Mode) - このウィンドウを閉じると終了します" } catch { }
$Script:AutoRenderInProgress = $false
$Script:ServerStartedUtc = [DateTime]::UtcNow
$Script:IdleTimeoutSeconds = 1800
$Script:NoClientStartupTimeoutSeconds = 600
$Script:ReadyGifBytes = [Convert]::FromBase64String('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==')
$Script:ExcelPrintProfileVersion = 2026072201
$Script:PdfImportProfileVersion = 2026080601
$Script:WordRenderProfileVersion = 2026080601
$Script:WordRenderTimeoutSeconds = 120
$Script:PowerPointRenderProfileVersion = 2026080701
$Script:PowerPointRenderTimeoutSeconds = 120
# Excel は同一プロセス内の同期COM呼び出しなので、Word/PowerPoint のようにワーカーごと
# 落とす方法が使えない。進捗が止まってからの猶予として長めに取る(大きなブックは
# 1シートに数十秒かかる)。詳細は Start-ExcelRenderWatchdog を参照。
$Script:ExcelRenderTimeoutSeconds = 300
# 1シートあたりの上乗せ。複数シートを1回で書き出す経路は、その間まったく心拍を
# 打てないため、シート数ぶん猶予を伸ばさないと正常な変換を殺してしまう。
$Script:ExcelPerSheetAllowanceSeconds = 60
$Script:ExcelWatchdogProcess = $null
$Script:ExcelWatchdogHeartbeatPath = ''
$Script:ExcelWatchdogKilledPath = ''
$Script:ExcelWatchdogFiredSticky = $false
$Script:ExcelWatchdogLogPaths = @()
$Script:OwnedExcelProcessId = 0
# 実行中のジョブが「中止が要求されたか」を答えるスクリプトブロック。外部コマンドの
# 待ちループがこれを1秒ごとに見る。ジョブ外では $null。
$Script:NativeCancelProbe = $null
# 外部コマンド(呼び先はすべて java)の上限。用途ごとに分ける。
# 2026-08-10: 上限を導入した時点、9箇所すべてが既定の120秒で走っていた。120秒は
# Word/PowerPoint の COM 変換に合わせた値で、java の実計算に流用できるものではない。
# 分割はコード自身が2000ページのPDFを許容し、組版の出力先は共有フォルダーにもなる。
# どちらも120秒に触れ得るのに、呼び出し側が時間切れと異常終了を区別していなかったため、
# 健全な入力に「破損していないか確認してください」と表示していた。
$Script:JavaProbeTimeoutSeconds = 30
$Script:PdfSplitTimeoutSeconds = 900
$Script:FinalComposeTimeoutSeconds = 900
# 解析はPDF作成のクリティカルパス外(Invoke-PostRenderAnalysis)。上限を設けた元の理由が
# ここの AWT 初期化の無限待ちなので、他より短く抑える。
$Script:PdfAnalyzeTimeoutSeconds = 300
$Script:FinalPdfComposerProfileVersion = 20260806
$Script:RenderEnvironmentCache = $null
$Script:RenderEnvironmentCompared = $false
$Script:CurrentRenderEnvFingerprint = ''
$Script:CurrentRenderEnvInfo = $null
$Script:LastRenderAttempt = $null
$Script:PdfPageAnalyzerVersion = 1
$Script:SourceAdapterContractVersion = 1
$Script:JavaRuntimeSignature = ''
$Script:PendingAnalysis = $null
$Script:PdfPageAnalyzerAvailable = $null
$Script:AutoSchedulerProcessId = 0
$Script:AutoSchedulerProcess = $null
$Script:AutoSchedulerControlPath = ''
# V5-P2: 設定・パス・履歴容量は Get-WorkspacePath 経由で頻繁に参照される。
# 毎回ディスクを読む(さらに config は書く)と、共有ドライブ上で致命的に遅くなる。短時間だけキャッシュする。
$Script:AppConfigCache = $null
$Script:AppConfigCacheAtUtc = [DateTime]::MinValue
$Script:PathsCache = $null
$Script:PathsCacheAtUtc = [DateTime]::MinValue
$Script:HistorySizeCache = $null
$Script:HistorySizeCacheKey = ''
$Script:HistorySizeCacheAtUtc = [DateTime]::MinValue
$Script:ConfigCacheSeconds = 2
$Script:HistorySizeCacheSeconds = 60
$Script:ConfigMergeChanged = $false

function New-NowIso {
    return (Get-Date).ToString('yyyy-MM-ddTHH:mm:sszzz')
}

function Get-ErrorDetail($ErrorRecord) {
    if ($null -eq $ErrorRecord) { return '' }
    $parts = @()
    # V5-P3: 以前は Exception.ToString() が取れないと、メッセージ1行だけになり原因追跡ができなかった。
    # 型・HResult・.NETスタック・内部例外・発生行を、取れたものから必ず積む。
    try {
        $ex = $ErrorRecord.Exception
        $depth = 0
        while ($null -ne $ex -and $depth -lt 5) {
            $line = ('[{0}] {1}' -f $ex.GetType().FullName, [string]$ex.Message)
            try { $line += (' (HResult=0x{0:X8})' -f [int]$ex.HResult) } catch { }
            $parts += $line
            try { if (-not [string]::IsNullOrWhiteSpace([string]$ex.StackTrace)) { $parts += [string]$ex.StackTrace } } catch { }
            try { $ex = $ex.InnerException } catch { $ex = $null }
            $depth++
        }
    } catch { }
    try {
        $ii = $ErrorRecord.InvocationInfo
        if ($null -ne $ii) {
            $parts += ('at line {0}, char {1}: {2}' -f [int]$ii.ScriptLineNumber, [int]$ii.OffsetInLine, [string]$ii.Line).Trim()
        }
    } catch { }
    try { if ($ErrorRecord.FullyQualifiedErrorId) { $parts += ('errorId: ' + [string]$ErrorRecord.FullyQualifiedErrorId) } } catch { }
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ScriptStackTrace)) {
            $parts += "PowerShell stack:`n$([string]$ErrorRecord.ScriptStackTrace)"
        }
    } catch { }
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.InvocationInfo.PositionMessage)) {
            $parts += [string]$ErrorRecord.InvocationInfo.PositionMessage
        }
    } catch { }
    if ($parts.Count -eq 0) { return [string]$ErrorRecord }
    return ($parts -join "`n")
}



function ConvertTo-NativeArgumentString([string[]]$ArgumentList) {
    # CommandLineToArgvW の規則で1本のコマンドライン文字列にする。
    $parts = @()
    foreach ($argument in @($ArgumentList)) {
        $value = [string]$argument
        if ($value -eq '') { $parts += '""'; continue }
        if ($value -notmatch '[\s"]') { $parts += $value; continue }
        $escaped = [regex]::Replace($value, '(\\*)"', '$1$1\"')
        $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
        $parts += ('"' + $escaped + '"')
    }
    return ($parts -join ' ')
}
function Invoke-NativeCapture([string]$FilePath, [string[]]$ArgumentList, [int]$TimeoutSeconds = 120) {
    # V5-P3: Windows PowerShell 5.1 では、ネイティブコマンドの stderr を 2>&1 で取り込むと
    # ErrorRecord としてパイプラインに流れ、$ErrorActionPreference='Stop' の下では
    # NativeCommandError の例外になる。
    # PDFBox は日本語フォントを含むPDFで警告(Format 14 cmap table ...)を stderr に出すため、
    # 解析や組版が成功していても呼び出し側が「失敗」と誤認していた。
    #
    # 2026-08-09: 従来は `& $FilePath @args 2>&1 | ...` で待っており、待ち時間の上限が
    # 無かった。この関数の呼び出し先はすべて java で、その1つ PdfPageAnalyzer は PDF を
    # ラスタライズするため Windows では -Djava.awt.headless=true を付けても AWT の
    # ツールキット(sun.awt.windows.WToolkit)を生成する。その初期化はウィンドウ
    # ステーションが使えないと上限なしで待ち続けるため、CIのように対話デスクトップが
    # 不安定な環境では java が永久に返らず、呼び出し元ごと固まっていた。
    # ProcessStartInfo で直接起動し、上限を過ぎたらプロセスツリーごと終了させる。
    # 解析は「PDF作成のクリティカルパスの外」(Invoke-PostRenderAnalysis を参照)なので、
    # 例外ではなく exitCode 非0 として返し、呼び出し元の既存の失敗処理に載せる。
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = ConvertTo-NativeArgumentString $ArgumentList
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    # 標準入力を渡さない。継承した端を子や孫が握ると、呼び出し元のパイプが閉じない。
    $psi.RedirectStandardInput = $true
    $proc = $null
    $timedOut = $false
    $cancelled = $false
    $exit = -1
    $stdout = ''
    $stderr = ''
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $proc = [Diagnostics.Process]::Start($psi)
        try { $proc.StandardInput.Close() } catch { }
        $outTask = $proc.StandardOutput.ReadToEndAsync()
        $errTask = $proc.StandardError.ReadToEndAsync()
        # 1秒ずつ待って、そのたびに中止要求を見る。上限だけで待つと、上限を長くした分
        # そのまま「中止を押しても何も起きない時間」になる(組版と分割は15分)。
        $deadline = [DateTime]::UtcNow.AddSeconds([Math]::Max(1, $TimeoutSeconds))
        $exited = $false
        while ($true) {
            if ($proc.WaitForExit(1000)) { $exited = $true; break }
            if ([DateTime]::UtcNow -ge $deadline) { break }
            if ($null -ne $Script:NativeCancelProbe) {
                $shouldCancel = $false
                try { $shouldCancel = [bool](& $Script:NativeCancelProbe) } catch { $shouldCancel = $false }
                if ($shouldCancel) { $cancelled = $true; break }
            }
        }
        if ($exited) {
            try { $proc.WaitForExit() } catch { }
            $exit = [int]$proc.ExitCode
        } elseif ($cancelled) {
            try { Stop-ReportBinderProcessTree ([int]$proc.Id) } catch { }
            try { [void]$proc.WaitForExit(5000) } catch { }
        } else {
            $timedOut = $true
            try { Stop-ReportBinderProcessTree ([int]$proc.Id) } catch { }
            try { [void]$proc.WaitForExit(5000) } catch { }
        }
        try { if ($outTask.Wait(5000)) { $stdout = [string]$outTask.Result } } catch { }
        try { if ($errTask.Wait(5000)) { $stderr = [string]$errTask.Result } } catch { }
    } catch {
        $stderr = [string]$_.Exception.Message
    } finally {
        $ErrorActionPreference = $previous
        if ($null -ne $proc) { try { $proc.Dispose() } catch { } }
    }
    $text = (($stdout + "`n" + $stderr) -replace "`r`n", "`n").Trim()
    if ($timedOut) {
        $text = (("NATIVE_TIMEOUT: " + $FilePath + " が " + [string]$TimeoutSeconds + " 秒で終わらないため中止しました。") + "`n" + $text).Trim()
    }
    if ($cancelled) {
        $text = ("NATIVE_CANCELLED: 中止の要求を受けて外部コマンドを終了しました。" + "`n" + $text).Trim()
    }
    $lines = if ($text -eq '') { @() } else { @($text -split "`n") }
    return [ordered]@{ exitCode = $exit; output = $lines; text = $text; timedOut = $timedOut; cancelled = $cancelled }
}

# Read a file timestamp from an open Windows file handle. On SMB shares this is
# more reliable than directory-enumeration metadata and matches Explorer's
# "更新日時" value. Fall back to System.IO on non-Windows or if the handle call fails.
if ($env:OS -eq 'Windows_NT' -and -not ('ReportBinderNative.FileTimes' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace ReportBinderNative {
    public static class FileTimes {
        private const uint FILE_READ_ATTRIBUTES = 0x00000080;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint FILE_SHARE_DELETE = 0x00000004;
        private const uint OPEN_EXISTING = 3;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(
            string fileName,
            uint desiredAccess,
            uint shareMode,
            IntPtr securityAttributes,
            uint creationDisposition,
            uint flagsAndAttributes,
            IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetFileTime(
            SafeFileHandle fileHandle,
            out long creationTime,
            out long lastAccessTime,
            out long lastWriteTime);

        public static long GetLastWriteFileTimeUtc(string path) {
            using (SafeFileHandle handle = CreateFile(
                path,
                FILE_READ_ATTRIBUTES,
                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                IntPtr.Zero)) {
                if (handle.IsInvalid) {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                long creation, access, write;
                if (!GetFileTime(handle, out creation, out access, out write)) {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                return write;
            }
        }
    }

    // 起動した Office アプリのプロセスIDを、そのウィンドウハンドルから引く。
    // 利用者が自分で開いている Excel を巻き添えにしないため、こちらが起動した
    // 1つだけを特定する必要がある。
    public static class OfficeWindows {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}
'@
}

function Get-FileLastWriteSnapshot([string]$Path) {
    $utc = $null
    if ($env:OS -eq 'Windows_NT' -and ('ReportBinderNative.FileTimes' -as [type])) {
        try {
            $fileTimeUtc = [ReportBinderNative.FileTimes]::GetLastWriteFileTimeUtc($Path)
            $utc = [DateTime]::FromFileTimeUtc($fileTimeUtc)
        } catch { }
    }
    if ($null -eq $utc) {
        $utc = [IO.File]::GetLastWriteTimeUtc($Path)
    }
    $local = $utc.ToLocalTime()
    return [ordered]@{
        utc = $utc
        local = $local
        display = $local.ToString('yyyy/MM/dd HH:mm')
        unixMs = [int64]([DateTimeOffset]::new($utc).ToUnixTimeMilliseconds())
    }
}

function ConvertTo-Win32ExtendedPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path) -or [Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT -or $Path.StartsWith('\\?\')) { return $Path }
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.Length -lt 240) { return $full }
    if ($full.StartsWith('\\')) { return '\\?\UNC\' + $full.Substring(2) }
    return '\\?\' + $full
}

function Test-FileExistsCompat([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return [IO.File]::Exists((ConvertTo-Win32ExtendedPath $Path))
}

function Test-DirectoryExistsCompat([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return [IO.Directory]::Exists((ConvertTo-Win32ExtendedPath $Path))
}

function Read-TextFileShared([string]$Path) {
    $share = [IO.FileShare]([int][IO.FileShare]::ReadWrite -bor [int][IO.FileShare]::Delete)
    $fs = [IO.File]::Open((ConvertTo-Win32ExtendedPath $Path), [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
    try {
        $reader = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8, $true)
        try { return $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    } finally {
        $fs.Dispose()
    }
}

function Read-JsonFile([string]$Path, $DefaultValue) {
    for ($attempt = 0; $attempt -lt 10; $attempt++) {
        try {
            if (-not [IO.File]::Exists((ConvertTo-Win32ExtendedPath $Path))) { return $DefaultValue }
            $text = Read-TextFileShared $Path
            if ([string]::IsNullOrWhiteSpace($text)) { return $DefaultValue }
            return $text | ConvertFrom-Json
        } catch [System.IO.FileNotFoundException] {
            return $DefaultValue
        } catch [System.IO.IOException] {
            if ($attempt -ge 9) { throw }
            Start-Sleep -Milliseconds (60 + (45 * $attempt))
        } catch [System.UnauthorizedAccessException] {
            if ($attempt -ge 9) { throw }
            Start-Sleep -Milliseconds (60 + (45 * $attempt))
        } catch {
            # A progress JSON file can be read exactly while it is being rewritten.
            # Retry parse failures instead of treating a temporary partial file as a fatal job error.
            if ($attempt -ge 9) {
                if ($null -ne $DefaultValue) { return $DefaultValue }
                throw
            }
            Start-Sleep -Milliseconds (70 + (50 * $attempt))
        }
    }
    return $DefaultValue
}

function Write-Utf8NoBomFile([string]$Path, [string]$Text) {
    # Windows PowerShell 5.1's Set-Content -Encoding UTF8 writes a BOM.
    # ReportPdfComposer's compact JSON reader expects plain UTF-8, so write JSON files without BOM.
    $encoding = New-Object System.Text.UTF8Encoding -ArgumentList $false
    [IO.File]::WriteAllText((ConvertTo-Win32ExtendedPath $Path), $Text, $encoding)
}

function Write-Utf8NoBomFileShared([string]$Path, [string]$Text) {
    # Fallback for protected / synced folders where File.Replace can intermittently
    # raise Access Denied. Readers use ReadWrite/Delete sharing and Read-JsonFile
    # retries partial reads, so progress keeps moving instead of staying at 0%.
    $encoding = New-Object System.Text.UTF8Encoding -ArgumentList $false
    $bytes = $encoding.GetBytes($Text)
    $parent = Split-Path -Parent $Path
    $extendedParent = ConvertTo-Win32ExtendedPath $parent
    if (-not [IO.Directory]::Exists($extendedParent)) { [void][IO.Directory]::CreateDirectory($extendedParent) }
    for ($attempt = 0; $attempt -lt 12; $attempt++) {
        $fs = $null
        try {
            $share = [IO.FileShare]([int][IO.FileShare]::ReadWrite -bor [int][IO.FileShare]::Delete)
            $fs = [IO.File]::Open((ConvertTo-Win32ExtendedPath $Path), [IO.FileMode]::Create, [IO.FileAccess]::Write, $share)
            $fs.Write($bytes, 0, $bytes.Length)
            try { $fs.Flush($true) } catch { $fs.Flush() }
            return
        } catch [System.IO.IOException] {
            if ($attempt -ge 11) { throw }
        } catch [System.UnauthorizedAccessException] {
            if ($attempt -ge 11) { throw }
        } finally {
            if ($fs) { try { $fs.Dispose() } catch { } }
        }
        Start-Sleep -Milliseconds (50 + (35 * $attempt))
    }
}

function Move-FileAtomicCompat([string]$SourcePath, [string]$DestinationPath) {
    $sourceIoPath = ConvertTo-Win32ExtendedPath $SourcePath
    $destinationIoPath = ConvertTo-Win32ExtendedPath $DestinationPath
    for ($attempt = 0; $attempt -lt 12; $attempt++) {
        try {
            if ([IO.File]::Exists($destinationIoPath)) {
                [IO.File]::Replace($sourceIoPath, $destinationIoPath, $null, $true)
            } else {
                [IO.File]::Move($sourceIoPath, $destinationIoPath)
            }
            return
        } catch [System.IO.FileNotFoundException] {
            try { [IO.File]::Move($sourceIoPath, $destinationIoPath); return } catch { if ($attempt -ge 11) { throw } }
        } catch [System.IO.IOException] {
            if ($attempt -ge 11) { throw }
        } catch [System.UnauthorizedAccessException] {
            if ($attempt -ge 11) { throw }
        }
        Start-Sleep -Milliseconds (50 + (30 * $attempt))
    }
}

function Write-JsonFile([string]$Path, $Value) {
    $parent = Split-Path -Parent $Path
    $extendedParent = ConvertTo-Win32ExtendedPath $parent
    if (-not [IO.Directory]::Exists($extendedParent)) { [void][IO.Directory]::CreateDirectory($extendedParent) }
    $tmp = "$Path.tmp.$([Guid]::NewGuid().ToString('N'))"
    $json = ConvertTo-Json -InputObject $Value -Depth 50
    $atomicError = ''
    try {
        # 深い共有フォルダでは、最終パスは扱えてもGUID付き一時名だけが
        # MAX_PATHを超えることがある。作成もtry内に置き、直接書込へ縮退する。
        Write-Utf8NoBomFile $tmp $json
        Move-FileAtomicCompat $tmp $Path
        return
    } catch {
        $atomicError = $_.Exception.Message
        # Some Windows/OneDrive/antivirus combinations deny File.Replace on JSON files
        # that are being watched or previewed. Use a shared direct write as a fallback.
        try {
            Write-Utf8NoBomFileShared $Path $json
            return
        } catch {
            throw "JSON保存に失敗しました: $Path / atomic=$atomicError / shared=$($_.Exception.Message)"
        }
    } finally {
        $tmpIoPath = ConvertTo-Win32ExtendedPath $tmp
        if ([IO.File]::Exists($tmpIoPath)) { try { [IO.File]::Delete($tmpIoPath) } catch { } }
    }
}

function Read-TextFileTailSafe([string]$Path, [int]$MaxChars = 4000) {
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return '' }
        $text = Read-TextFileShared $Path
        if ([string]::IsNullOrWhiteSpace($text)) { return '' }
        if ($text.Length -le $MaxChars) { return $text }
        return $text.Substring($text.Length - $MaxChars)
    } catch { return '' }
}

function Set-RenderJobFailedFromStartupProblem([string]$StatusPath, $Job, [string]$Message) {
    if ($null -eq $Job) { return $Job }
    Set-NoteProperty $Job 'status' 'failed'
    Set-NoteProperty $Job 'percent' 100
    Set-NoteProperty $Job 'message' $Message
    $stdoutPath = [string](Get-DataProperty $Job 'stdoutPath' '')
    $stderrPath = [string](Get-DataProperty $Job 'stderrPath' '')
    $stdoutTail = Read-TextFileTailSafe $stdoutPath 3000
    $stderrTail = Read-TextFileTailSafe $stderrPath 3000
    $detail = @()
    if (-not [string]::IsNullOrWhiteSpace($stderrTail)) { $detail += "stderr:`n$stderrTail" }
    if (-not [string]::IsNullOrWhiteSpace($stdoutTail)) { $detail += "stdout:`n$stdoutTail" }
    if ($detail.Count -gt 0) { Set-NoteProperty $Job 'startupLog' ($detail -join "`n`n") }
    $err = [ordered]@{ error = $Message; userError = 'PDF作成プロセスを起動できませんでした。ReportBinderを一度終了してから再実行してください。'; detail = ([string](Get-DataProperty $Job 'startupLog' '')) }
    Set-NoteProperty $Job 'errors' @($err)
    Write-RenderJobStatus $StatusPath $Job
    return $Job
}

function Get-Array($Value) {
    if ($null -eq $Value) { return @() }
    if ($Value -is [System.Array]) { return @($Value) }
    return @($Value)
}


function Set-ArrayProperty($Object, [string]$Name) {
    if ($null -eq $Object) { return }
    $arr = @(Get-Array $Object.$Name)
    if ($Object.PSObject.Properties[$Name]) {
        $Object.$Name = $arr
    } else {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $arr
    }
}

function Normalize-StructureCollections($Structure) {
    if ($null -eq $Structure) { return $Structure }
    $version = 0
    try { $version = [int](Get-DataProperty $Structure 'schemaVersion' 0) } catch { $version = 0 }
    if ($version -eq 3) { Update-LegacyCompatibilityViewFromV3 $Structure }
    Set-ArrayProperty $Structure 'workbooks'
    Set-ArrayProperty $Structure 'pages'
    if (-not (Test-ConfigHasKey $Structure 'volumes') -or $null -eq (Get-DataProperty $Structure 'volumes' $null)) { Set-NoteProperty $Structure 'volumes' ([ordered]@{}) }
    if ($version -eq 3) {
        foreach ($name in @('packs','sources','units','items','artifacts','migrationIssues')) { Set-ArrayProperty $Structure $name }
        if (-not (Test-ConfigHasKey $Structure 'outputs') -or $null -eq (Get-DataProperty $Structure 'outputs' $null)) { Set-NoteProperty $Structure 'outputs' ([ordered]@{}) }
    }
    return $Structure
}

function Test-DirectExcelRelativePath([string]$RelativePath) {
    Test-RelativePath $RelativePath | Out-Null
    if ($RelativePath -match '[\\/]') { throw '提出フォルダ直下のExcelだけ登録できます。子フォルダ内のファイルは対象外です。' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -notin @('.xlsx','.xlsm')) { throw '拡張子が .xlsx または .xlsm のExcelだけ登録できます。' }
    if ([IO.Path]::GetFileName($RelativePath) -like '~$*') { throw 'Excelの一時ファイルは登録できません。' }
    return $true
}

function Get-ConfigKeyNames($Object) {
    if ($null -eq $Object) { return @() }
    if ($Object -is [System.Collections.IDictionary]) { return @($Object.Keys | ForEach-Object { [string]$_ }) }
    return @($Object.PSObject.Properties | ForEach-Object { [string]$_.Name })
}

function Test-ConfigHasKey($Object, [string]$Name) {
    if ($null -eq $Object) { return $false }
    if ($Object -is [System.Collections.IDictionary]) { return $Object.Contains($Name) }
    return ($null -ne $Object.PSObject.Properties[$Name])
}

function Merge-ConfigDefaults($Target, $Defaults) {
    # 既定値をキー単位で再帰的に補完する。既にある値は必ず優先する(利用者の設定を壊さない)。
    # Target/Defaults は PSCustomObject でも ordered hashtable でもよい。
    if ($null -eq $Defaults) { return $Target }
    if ($null -eq $Target) { return $Defaults }
    foreach ($name in (Get-ConfigKeyNames $Defaults)) {
        $defValue = Get-DataProperty $Defaults $name $null
        if (-not (Test-ConfigHasKey $Target $name)) {
            Set-NoteProperty $Target $name $defValue
            # V5-P2: 既定値を実際に補完したときだけ config.json を書き戻す。
            $Script:ConfigMergeChanged = $true
            continue
        }
        $cur = Get-DataProperty $Target $name $null
        $defIsObj = ($defValue -is [pscustomobject]) -or ($defValue -is [System.Collections.IDictionary])
        $curIsObj = ($cur -is [pscustomobject]) -or ($cur -is [System.Collections.IDictionary])
        if ($defIsObj -and $curIsObj) { [void](Merge-ConfigDefaults $cur $defValue) }
    }
    return $Target
}

function Reset-ConfigCaches {
    # 設定・パスを変更したら必ず呼ぶ。
    $Script:AppConfigCache = $null
    $Script:AppConfigCacheAtUtc = [DateTime]::MinValue
    $Script:PathsCache = $null
    $Script:PathsCacheAtUtc = [DateTime]::MinValue
    # dataDir の切替時に旧ワークスペースの容量を返さない。
    $Script:HistorySizeCache = $null
    $Script:HistorySizeCacheKey = ''
    $Script:HistorySizeCacheAtUtc = [DateTime]::MinValue
}

function Get-AppConfig {
    # V5: ローカルconfigをそのまま返すと、既存利用者に新しいキー(autoRender など)が反映されない。
    # 既定値を読み、ローカル優先でキー単位にマージしてから返す。
    # V5-P2: 以前はこの関数が呼ばれるたびに config.json を書き戻していた。
    # Get-Paths -> Get-WorkspacePath 経由でほぼ全関数から呼ばれるため、
    # 検知版の一覧取得などで1件ごとにファイル書き込みが発生していた。
    # 実際に既定値を補完したときだけ書き、結果は短時間キャッシュする。
    if ($null -ne $Script:AppConfigCache -and (([DateTime]::UtcNow - $Script:AppConfigCacheAtUtc).TotalSeconds -lt $Script:ConfigCacheSeconds)) {
        return $Script:AppConfigCache
    }
    $default = [ordered]@{ schemaVersion = 1; lastSubmissionDir = ''; lastDataDir = ''; lastOutputDir = '' }
    $seed = Read-JsonFile $Script:DefaultConfigPath $default
    if (-not (Test-Path -LiteralPath $Script:ConfigPath)) {
        try { Write-JsonFile $Script:ConfigPath $seed } catch { }
        $Script:AppConfigCache = $seed
        $Script:AppConfigCacheAtUtc = [DateTime]::UtcNow
        return $seed
    }
    $local = Read-JsonFile $Script:ConfigPath $default
    $Script:ConfigMergeChanged = $false
    $merged = Merge-ConfigDefaults $local $seed
    if ([string](Get-DataProperty $merged 'schemaVersion' '') -ne '2') {
        Set-NoteProperty $merged 'schemaVersion' 2
        $Script:ConfigMergeChanged = $true
    }
    if ($Script:ConfigMergeChanged) { try { Write-JsonFile $Script:ConfigPath $merged } catch { } }
    $Script:AppConfigCache = $merged
    $Script:AppConfigCacheAtUtc = [DateTime]::UtcNow
    return $merged
}


function Test-InputHistoryEnabled {
    # 履歴・差分は既定で有効。旧 policy.json の承認フラグは参照しない。
    return $true
}

function Test-SourceRetentionEnabled {
    # 利用者が選択した提出Excelの履歴は、ReportBinder専用のローカルプロジェクト内に保持する。
    # 共有提出フォルダーへ履歴・中間PDFを書かない。
    try {
        $paths = Get-Paths
        $data = [IO.Path]::GetFullPath([string]$paths.dataDir)
        $localRoot = [IO.Path]::GetFullPath($Script:LocalProjectsRoot)
        if (-not $localRoot.EndsWith([IO.Path]::DirectorySeparatorChar)) { $localRoot += [IO.Path]::DirectorySeparatorChar }
        return $data.StartsWith($localRoot, [StringComparison]::OrdinalIgnoreCase)
    } catch { return $false }
}

function Get-AutoRenderSettings {
    $config = Get-AppConfig
    $auto = Get-DataProperty $config 'autoRender' $null
    return [ordered]@{
        enabled = [bool](Get-DataProperty $auto 'enabled' $false)
        quietPeriodSeconds = [int](Get-DataProperty $auto 'quietPeriodSeconds' 180)
        requireStableHashCount = [int](Get-DataProperty $auto 'requireStableHashCount' 2)
        deferWhileExcelInUse = [bool](Get-DataProperty $auto 'deferWhileExcelInUse' $true)
        maxRetryCount = [int](Get-DataProperty $auto 'maxRetryCount' 3)
        retryBaseSeconds = [int](Get-DataProperty $auto 'retryBaseSeconds' 60)
        minFreeMegabytes = [int](Get-DataProperty $auto 'minFreeMegabytes' 1024)
        notifyOnCompletion = [bool](Get-DataProperty $auto 'notifyOnCompletion' $true)
        notifyOnFailure = [bool](Get-DataProperty $auto 'notifyOnFailure' $true)
    }
}

function Update-AutoRenderSettings($Patch) {
    $config = Get-AppConfig
    $current = Get-AutoRenderSettings
    $next = [ordered]@{
        enabled = $(if (Test-ConfigHasKey $Patch 'enabled') { [bool](Get-DataProperty $Patch 'enabled' $false) } else { [bool]$current.enabled })
        quietPeriodSeconds = $(if (Test-ConfigHasKey $Patch 'quietPeriodSeconds') { [Math]::Min(3600,[Math]::Max(10,[int](Get-DataProperty $Patch 'quietPeriodSeconds' 180))) } else { [int]$current.quietPeriodSeconds })
        requireStableHashCount = $(if (Test-ConfigHasKey $Patch 'requireStableHashCount') { [Math]::Min(10,[Math]::Max(1,[int](Get-DataProperty $Patch 'requireStableHashCount' 2))) } else { [int]$current.requireStableHashCount })
        deferWhileExcelInUse = $(if (Test-ConfigHasKey $Patch 'deferWhileExcelInUse') { [bool](Get-DataProperty $Patch 'deferWhileExcelInUse' $true) } else { [bool]$current.deferWhileExcelInUse })
        maxRetryCount = $(if (Test-ConfigHasKey $Patch 'maxRetryCount') { [Math]::Min(10,[Math]::Max(0,[int](Get-DataProperty $Patch 'maxRetryCount' 3))) } else { [int]$current.maxRetryCount })
        retryBaseSeconds = $(if (Test-ConfigHasKey $Patch 'retryBaseSeconds') { [Math]::Min(3600,[Math]::Max(10,[int](Get-DataProperty $Patch 'retryBaseSeconds' 60))) } else { [int]$current.retryBaseSeconds })
        minFreeMegabytes = $(if (Test-ConfigHasKey $Patch 'minFreeMegabytes') { [Math]::Min(102400,[Math]::Max(100,[int](Get-DataProperty $Patch 'minFreeMegabytes' 1024))) } else { [int]$current.minFreeMegabytes })
        notifyOnCompletion = $(if (Test-ConfigHasKey $Patch 'notifyOnCompletion') { [bool](Get-DataProperty $Patch 'notifyOnCompletion' $true) } else { [bool]$current.notifyOnCompletion })
        notifyOnFailure = $(if (Test-ConfigHasKey $Patch 'notifyOnFailure') { [bool](Get-DataProperty $Patch 'notifyOnFailure' $true) } else { [bool]$current.notifyOnFailure })
    }
    Set-NoteProperty $config 'autoRender' ([pscustomobject]$next)
    Save-AppConfig $config
    if ([bool]$next.enabled) { if (-not (Test-AutoSchedulerProcessRunning)) { [void](Start-AutoSchedulerProcess (Get-EffectiveLanguage)) } }
    else { Stop-AutoSchedulerProcess }
    return Get-AutoStateSummary (Get-EffectiveLanguage)
}

function Get-InputHistorySettings {
    $config = Get-AppConfig
    $ih = Get-DataProperty $config 'inputHistory' $null
    return [ordered]@{
        retainSourceVersions = [int](Get-DataProperty $ih 'retainSourceVersions' 5)
        retainContentPdfVersions = [int](Get-DataProperty $ih 'retainContentPdfVersions' 3)
        sourceRetentionDaysAfterBuild = (Get-DataProperty $ih 'sourceRetentionDaysAfterBuild' $null)
        softCapMegabytes = [int](Get-DataProperty $ih 'softCapMegabytes' 5120)
        warnAtPercent = [int](Get-DataProperty $ih 'warnAtPercent' 80)
        ephemeralCopyMaxAgeMinutes = [int](Get-DataProperty $ih 'ephemeralCopyMaxAgeMinutes' 30)
    }
}

function Save-AppConfig($Config) {
    # Current values are always per-user. The shared default-config.json is read-only at runtime.
    Write-JsonFile $Script:ConfigPath $Config
    Reset-ConfigCaches
}

function Convert-CmToPt([double]$Cm) { return $Cm * 28.3464567 }


function Get-RelativePathCompat([string]$BasePath, [string]$FullPath) {
    $base = [IO.Path]::GetFullPath($BasePath)
    if (-not $base.EndsWith([IO.Path]::DirectorySeparatorChar)) { $base += [IO.Path]::DirectorySeparatorChar }
    $full = [IO.Path]::GetFullPath($FullPath)
    # フォルダ名に % や # を含むと Uri.MakeRelativeUri / UnescapeDataString が
    # パスを壊す（%20→空白化、#以降欠落）。配下の場合は単純な切り出しで求める。
    if ($full.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($base.Length)
    }
    $baseUri = New-Object System.Uri($base)
    $fullUri = New-Object System.Uri($full)
    $rel = [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($fullUri).ToString())
    return ($rel -replace '/', [IO.Path]::DirectorySeparatorChar)
}

function Normalize-FileHash([string]$Value) {
    # V5: ファイルハッシュは New-Sha256 の形式(64文字・大文字・prefixなし)に統一する。
    # 過去データや手書き設定に 'sha256:' 付き小文字が混ざっていても比較が壊れないよう吸収する。
    # 文字列 fingerprint (Get-Sha256Text) の 'sha256:'+小文字 とは別物なので混ぜないこと。
    $v = [string]$Value
    if ([string]::IsNullOrWhiteSpace($v)) { return '' }
    if ($v -match '^(?i)sha256:') { $v = $v.Substring(7) }
    return $v.Trim().ToUpperInvariant()
}

function New-RbId {
    # V5: ID = タイムスタンプ(ミリ秒) + '_' + GUID8。
    # ハッシュや fingerprint は ID に埋め込まない(manifest の正式フィールドとして持つ)。
    return ((Get-Date).ToString('yyyyMMddTHHmmss.fff') + '_' + ([Guid]::NewGuid().ToString('N').Substring(0,8)))
}

function New-RbVersionId {
    return ('v' + (New-RbId))
}

function New-UniqueDirectory([string]$Parent, [scriptblock]$IdFactory) {
    # 生成後に既存ディレクトリがあれば再生成する(最大3回)。immutable な世代フォルダが混ざるのを防ぐ。
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $id = [string](& $IdFactory)
        $full = Join-Path $Parent $id
        if (-not (Test-Path -LiteralPath $full)) {
            New-Item -ItemType Directory -Path $full -Force | Out-Null
            return [ordered]@{ id = $id; path = $full }
        }
        Start-Sleep -Milliseconds 5
    }
    throw "一意なフォルダ名を生成できませんでした: $Parent"
}

function New-Sha256([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    # Get-FileHash は FileShare.Read で開くため、誰かがExcelで(書き込みアクセス付きで)開いていると
    # 「別のプロセスで使用されています」で失敗する。FileShare.ReadWrite を明示して開けば読める。
    $fs = $null
    $sha = $null
    try {
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $sha = [Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha.ComputeHash($fs)
        $hex = (-join ($hashBytes | ForEach-Object { $_.ToString('x2') })).ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($hex)) { throw "ハッシュを計算できませんでした: $Path" }
        return $hex
    } finally {
        if ($sha) { try { $sha.Dispose() } catch { } }
        if ($fs) { try { $fs.Dispose() } catch { } }
    }
}

function Copy-FileSharedRead([string]$Source, [string]$Destination) {
    # Excelで開かれている(書き込みアクセス保持中の)ファイルもコピーできるよう、
    # 読み取り側を FileShare.ReadWrite で開く。Copy-Item では同じ理由で失敗する。
    $inStream = $null
    $outStream = $null
    try {
        $inStream = [IO.File]::Open($Source, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $outStream = [IO.File]::Open($Destination, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $inStream.CopyTo($outStream)
    } finally {
        if ($outStream) { try { $outStream.Dispose() } catch { } }
        if ($inStream) { try { $inStream.Dispose() } catch { } }
    }
}

function New-StableHash([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    # Old implementation always waited at least 500ms per workbook. That becomes very visible
    # when PDF作成 processes many Excel files. If the file has not been touched for a few seconds,
    # hash it immediately; only recently modified files get a short stability check.
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $ageMs = ([DateTime]::UtcNow - $item.LastWriteTimeUtc).TotalMilliseconds
    if ($ageMs -ge 3000) { return New-Sha256 $Path }

    $lastSize = $item.Length
    $lastWrite = $item.LastWriteTimeUtc
    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 100
        $next = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($next.Length -eq $lastSize -and $next.LastWriteTimeUtc -eq $lastWrite) { break }
        $lastSize = $next.Length
        $lastWrite = $next.LastWriteTimeUtc
    }
    return New-Sha256 $Path
}

function New-Slug([string]$Text) {
    $base = [IO.Path]::GetFileNameWithoutExtension($Text).ToLowerInvariant()
    $base = [regex]::Replace($base, '[^a-z0-9]+', '-')
    $base = $base.Trim('-')
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'item' }
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
    $hash = ([BitConverter]::ToString($sha1.ComputeHash($bytes))).Replace('-', '').Substring(0, 8).ToLowerInvariant()
    return "$base-$hash"
}

function Test-RelativePath([string]$RelativePath) {
    if ([string]::IsNullOrWhiteSpace($RelativePath)) { throw 'relativePath が空です。' }
    if ([IO.Path]::IsPathRooted($RelativePath)) { throw '絶対パスは受け付けません。' }
    if ($RelativePath -match '(^|[\\/])\.\.($|[\\/])') { throw '.. を含むパスは受け付けません。' }
    if ($RelativePath -match '[\x00-\x1F]') { throw '制御文字を含むパスは受け付けません。' }
    return $true
}

function Assert-SafeStorageSegment([string]$Value, [string]$Name = '識別子') {
    # workbookId / snapshotId / versionId are used as single directory or file-name
    # segments. Never let API input introduce separators, drive prefixes, or dot
    # traversal into the history/archive trees.
    $segment = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($segment) -or
        $segment.Length -gt 200 -or
        $segment -in @('.', '..') -or
        $segment -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        throw [System.ArgumentException]::new("$Name が不正です。")
    }
    return $segment
}

function Join-Safe([string]$Root, [string]$RelativePath) {
    Test-RelativePath $RelativePath | Out-Null
    $full = [IO.Path]::GetFullPath((Join-Path $Root $RelativePath))
    $rootFull = [IO.Path]::GetFullPath($Root)
    if (-not $rootFull.EndsWith([IO.Path]::DirectorySeparatorChar)) { $rootFull += [IO.Path]::DirectorySeparatorChar }
    if (-not $full.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { throw '登録済みフォルダ外のパスです。' }
    return $full
}

function Get-Paths {
    # V5-P2: Get-WorkspacePath 経由でほぼ全関数から呼ばれる。共有ドライブ上の
    # common\paths.json を1回の操作で何百回も読み直さないよう短時間キャッシュする。
    if ($null -ne $Script:PathsCache -and (([DateTime]::UtcNow - $Script:PathsCacheAtUtc).TotalSeconds -lt $Script:ConfigCacheSeconds)) {
        return $Script:PathsCache
    }
    $config = Get-AppConfig
    $paths = [ordered]@{
        submissionDir = [string]$config.lastSubmissionDir
        dataDir       = [string]$config.lastDataDir
        outputDir     = [string]$config.lastOutputDir
    }
    $result = $paths
    if ($paths.dataDir -and (Test-Path -LiteralPath (Join-Path $paths.dataDir 'common\paths.json'))) {
        $stored = Read-JsonFile (Join-Path $paths.dataDir 'common\paths.json') $null
        if ($stored) { $result = $stored }
    }
    $Script:PathsCache = $result
    $Script:PathsCacheAtUtc = [DateTime]::UtcNow
    return $result
}


function Get-CanonicalSubmissionPath([string]$SubmissionDir) {
    if ([string]::IsNullOrWhiteSpace($SubmissionDir)) { return '' }
    $full = [IO.Path]::GetFullPath($SubmissionDir).TrimEnd([char[]]@([char]92, [char]47))
    # 同じ共有フォルダーを H:\ と \\server\share の両方で選んでも、同じローカル
    # プロジェクトになるよう、取得できる場合はマップドライブをUNCへ正規化する。
    if ($full -match '^([A-Za-z]):[\\/]') {
        try {
            $drive = Get-PSDrive -Name $matches[1] -ErrorAction Stop
            $displayRoot = ''
            $displayProp = $drive.PSObject.Properties['DisplayRoot']
            if ($null -ne $displayProp) { $displayRoot = [string]$displayProp.Value }
            if ([string]::IsNullOrWhiteSpace($displayRoot) -and ([string]$drive.Root -match '^\\\\')) { $displayRoot = [string]$drive.Root }
            if (-not [string]::IsNullOrWhiteSpace($displayRoot)) {
                $tail = if ($full.Length -gt 3) { $full.Substring(3) } else { '' }
                $full = [IO.Path]::GetFullPath((Join-Path $displayRoot $tail)).TrimEnd([char[]]@([char]92, [char]47))
            }
        } catch { }
    }
    return $full
}

function Get-LocalProjectKey([string]$SubmissionDir) {
    $canonical = Get-CanonicalSubmissionPath $SubmissionDir
    if ([string]::IsNullOrWhiteSpace($canonical)) { return '' }
    $label = [IO.Path]::GetFileName($canonical)
    if ([string]::IsNullOrWhiteSpace($label)) { $label = 'project' }
    $label = [regex]::Replace($label, '[^\p{L}\p{Nd}._-]+', '-').Trim([char[]]@('-','.'))
    if ([string]::IsNullOrWhiteSpace($label)) { $label = 'project' }
    if ($label.Length -gt 32) { $label = $label.Substring(0,32) }
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($canonical.ToLowerInvariant()))
        $hash = -join ($bytes[0..7] | ForEach-Object { $_.ToString('x2') })
    } finally { $sha.Dispose() }
    return "$label-$hash"
}

function Get-DefaultChildPaths([string]$SubmissionDir) {
    if ([string]::IsNullOrWhiteSpace($SubmissionDir)) {
        return [ordered]@{ submissionDir = ''; dataDir = ''; outputDir = ''; projectRoot = ''; projectKey = '' }
    }
    $canonical = Get-CanonicalSubmissionPath $SubmissionDir
    $projectKey = Get-LocalProjectKey $canonical
    $projectRoot = Join-Path $Script:LocalProjectsRoot $projectKey
    return [ordered]@{
        submissionDir = $canonical
        dataDir = (Join-Path $projectRoot 'data')
        outputDir = (Join-Path $projectRoot 'output')
        projectRoot = $projectRoot
        projectKey = $projectKey
    }
}

function Move-IncompleteLocalProjectPath([string]$Path, [string]$ProjectRoot, [string]$Kind) {
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return '' }
    $hasEntries = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop | Select-Object -First 1).Count -gt 0
    if (-not $hasEntries) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
        return ''
    }

    # 移行完了マーカーがないローカルデータは、旧共有コピーの失敗途中か
    # 旧実装の残骸である。削除せず短い名前で退避し、新しい管理領域は空で始める。
    $suffix = (Get-Date -Format 'yyyyMMddHHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0,8)
    $backup = Join-Path $ProjectRoot ('.preclean-' + $Kind + '-' + $suffix)
    Move-Item -LiteralPath $Path -Destination $backup -ErrorAction Stop
    return $backup
}

function Initialize-CleanLocalProject($LocalPaths) {
    $submissionDir = [string]$LocalPaths.submissionDir
    $targetData = [string]$LocalPaths.dataDir
    $targetOutput = [string]$LocalPaths.outputDir
    $projectRoot = [string]$LocalPaths.projectRoot
    if ([string]::IsNullOrWhiteSpace($submissionDir) -or [string]::IsNullOrWhiteSpace($projectRoot)) {
        return [ordered]@{ initialized = $false; reason = 'not-configured' }
    }

    $marker = Join-Path $projectRoot 'local-project.json'
    if (Test-Path -LiteralPath $marker -PathType Leaf) {
        return [ordered]@{ initialized = $false; reason = 'already-local'; marker = $marker }
    }

    if (-not (Test-Path -LiteralPath $projectRoot)) {
        New-Item -ItemType Directory -Path $projectRoot -Force | Out-Null
    }

    # 古い版との同時起動でも旧共有コピーと新しいクリーン初期化が競合しないよう、
    # mutex名は既存版と同じまま維持する。
    $projectMutex = $null
    $projectMutexOwned = $false
    try {
        $created = $false
        $projectMutex = New-Object System.Threading.Mutex($false, ('Local\ReportBinder.ProjectMigration.' + [string]$LocalPaths.projectKey), [ref]$created)
        try { $projectMutexOwned = $projectMutex.WaitOne(1800000, $false) }
        catch [System.Threading.AbandonedMutexException] { $projectMutexOwned = $true }
        if (-not $projectMutexOwned) { throw '利用者ローカル管理領域の初期化待ちがタイムアウトしました。' }
        if (Test-Path -LiteralPath $marker -PathType Leaf) {
            return [ordered]@{ initialized = $false; reason = 'already-local'; marker = $marker }
        }

        # 共有側の _reportbinder / 出力 は存在確認もコピーも行わない。
        # 提出フォルダーはExcelの読込元としてだけ使用する。
        $quarantined = @()
        $oldData = Move-IncompleteLocalProjectPath $targetData $projectRoot 'data'
        if (-not [string]::IsNullOrWhiteSpace($oldData)) {
            $quarantined += [ordered]@{ kind = 'data'; path = $oldData }
        }
        $oldOutput = Move-IncompleteLocalProjectPath $targetOutput $projectRoot 'output'
        if (-not [string]::IsNullOrWhiteSpace($oldOutput)) {
            $quarantined += [ordered]@{ kind = 'output'; path = $oldOutput }
        }

        New-Item -ItemType Directory -Path $targetData -Force | Out-Null
        New-Item -ItemType Directory -Path $targetOutput -Force | Out-Null
        Write-JsonFile $marker ([ordered]@{
            schemaVersion = 2
            projectKey = [string]$LocalPaths.projectKey
            submissionDir = $submissionDir
            dataDir = $targetData
            outputDir = $targetOutput
            initializationMode = 'clean'
            legacySharedImport = $false
            initializedAt = New-NowIso
            quarantinedLocalPaths = @($quarantined)
        })
        return [ordered]@{
            initialized = $true
            reason = 'clean-start'
            marker = $marker
            dataDir = $targetData
            outputDir = $targetOutput
            quarantinedLocalPaths = @($quarantined)
        }
    } finally {
        if ($projectMutexOwned -and $null -ne $projectMutex) { try { $projectMutex.ReleaseMutex() } catch { } }
        if ($null -ne $projectMutex) { try { $projectMutex.Dispose() } catch { } }
    }
}

function Initialize-LocalProjectConfig($Config) {
    $submissionDir = [string](Get-DataProperty $Config 'lastSubmissionDir' '')
    if ([string]::IsNullOrWhiteSpace($submissionDir)) { return $Config }
    $localPaths = Get-DefaultChildPaths $submissionDir
    $currentData = [string](Get-DataProperty $Config 'lastDataDir' '')
    $currentOutput = [string](Get-DataProperty $Config 'lastOutputDir' '')
    $sameData = $false
    $sameOutput = $false
    try { $sameData = ([IO.Path]::GetFullPath($currentData) -eq [IO.Path]::GetFullPath([string]$localPaths.dataDir)) } catch { }
    try { $sameOutput = ([IO.Path]::GetFullPath($currentOutput) -eq [IO.Path]::GetFullPath([string]$localPaths.outputDir)) } catch { }
    if ($sameData -and $sameOutput) {
        [void](Initialize-CleanLocalProject $localPaths)
        return $Config
    }

    [void](Initialize-CleanLocalProject $localPaths)
    Set-NoteProperty $Config 'lastSubmissionDir' ([string]$localPaths.submissionDir)
    Set-NoteProperty $Config 'lastDataDir' ([string]$localPaths.dataDir)
    Set-NoteProperty $Config 'lastOutputDir' ([string]$localPaths.outputDir)
    Write-JsonFile $Script:ConfigPath $Config
    Reset-ConfigCaches
    return $Config
}


function ConvertTo-EnvBase64([string]$Text) {
    if ($null -eq $Text) { $Text = '' }
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text))
}

function Select-FolderDialog([string]$Title, [string]$InitialDir) {
    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = 'フォルダを選択してください' }
    if ([string]::IsNullOrWhiteSpace($InitialDir) -or -not (Test-Path -LiteralPath $InitialDir)) { $InitialDir = [Environment]::GetFolderPath('MyDocuments') }

    $helper = Join-Path $Script:AppRoot 'tools\select-folder.ps1'
    if (-not (Test-Path -LiteralPath $helper)) { throw 'フォルダ選択用の補助スクリプトが見つかりません。' }

    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("reportbinder-folder-{0}.txt" -f ([Guid]::NewGuid().ToString('N')))
    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }

    $oldTitle = [Environment]::GetEnvironmentVariable('REPORTBINDER_PICKER_TITLE_B64', 'Process')
    $oldInitial = [Environment]::GetEnvironmentVariable('REPORTBINDER_PICKER_INITIAL_B64', 'Process')
    $oldOutput = [Environment]::GetEnvironmentVariable('REPORTBINDER_PICKER_OUTPUT_B64', 'Process')
    try {
        [Environment]::SetEnvironmentVariable('REPORTBINDER_PICKER_TITLE_B64', (ConvertTo-EnvBase64 $Title), 'Process')
        [Environment]::SetEnvironmentVariable('REPORTBINDER_PICKER_INITIAL_B64', (ConvertTo-EnvBase64 $InitialDir), 'Process')
        [Environment]::SetEnvironmentVariable('REPORTBINDER_PICKER_OUTPUT_B64', (ConvertTo-EnvBase64 $tmp), 'Process')

        $escapedHelper = $helper -replace "'", "''"
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes("& '$escapedHelper'"))
        $args = "-NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $encoded"
        $proc = Start-Process -FilePath $psExe -ArgumentList $args -WindowStyle Hidden -PassThru
        if (-not $proc.WaitForExit(60000)) {
            try { $proc.Kill() } catch { }
            throw 'フォルダ選択画面が応答しません。画面でパスを直接入力してください。'
        }
        if ($proc.ExitCode -ne 0) { throw "フォルダ選択ダイアログを開けませんでした。ExitCode=$($proc.ExitCode)" }
        if (Test-Path -LiteralPath $tmp) {
            $selected = (Get-Content -LiteralPath $tmp -Raw -Encoding UTF8).Trim()
            if ($selected -and (Test-Path -LiteralPath $selected)) { return $selected }
            return ''
        }
        return ''
    } finally {
        [Environment]::SetEnvironmentVariable('REPORTBINDER_PICKER_TITLE_B64', $oldTitle, 'Process')
        [Environment]::SetEnvironmentVariable('REPORTBINDER_PICKER_INITIAL_B64', $oldInitial, 'Process')
        [Environment]::SetEnvironmentVariable('REPORTBINDER_PICKER_OUTPUT_B64', $oldOutput, 'Process')
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Get-WorkspacePath([string]$Language, [string]$DataDir = '') {
    # First-run folder setup must be able to initialize the selected workspace before
    # the per-user config is committed. Prefer the explicit dataDir when supplied;
    # normal API operations continue to resolve it from Get-Paths.
    $resolvedDataDir = ([string]$DataDir).Trim()
    if ([string]::IsNullOrWhiteSpace($resolvedDataDir)) {
        $paths = Get-Paths
        $resolvedDataDir = ([string]$paths.dataDir).Trim()
    }
    if ([string]::IsNullOrWhiteSpace($resolvedDataDir)) { throw '管理データフォルダが未設定です。提出フォルダを選んでください。' }
    # 一式を唯一の管理単位とし、すべてのパックを同じワークスペースに保存する。
    # 言語は保存領域を分割する設定ではなく、一式やひな形の属性として扱う。
    return Join-Path $resolvedDataDir 'workspace'
}

function New-EmptyVolumeState {
    return [ordered]@{
        status = 'not-built'
        lastBuiltAt = $null
        outputPdf = $null
        builtFingerprint = ''
        staleReasons = @()
        message = ''
    }
}

function Get-VolumeStateKey([string]$Volume, [string]$Category) {
    $cat = Normalize-WorkbookCategory $Category ''
    if ([string]::IsNullOrWhiteSpace($cat)) { return "$Volume|_" }
    return "$Volume|$cat"
}

function Test-DirectWordRelativePath([string]$RelativePath) {
    Test-RelativePath $RelativePath | Out-Null
    if ($RelativePath -match '[\/]') { throw '提出フォルダ直下のWord原稿だけ登録できます。子フォルダ内のファイルは対象外です。' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -ne '.docx') { throw '拡張子が .docx のWord原稿だけ登録できます。' }
    if ([IO.Path]::GetFileName($RelativePath) -like '~$*') { throw 'Wordの一時ファイルは登録できません。' }
    return $true
}

function Get-BuiltinPackId([string]$Category) {
    $cat = Normalize-WorkbookCategory $Category ''
    if ([string]::IsNullOrWhiteSpace($cat)) { return '' }
    return "pack_$cat"
}

function Get-CategoryFromBuiltinPackId([string]$PackId) {
    $value = ([string]$PackId).Trim().ToLowerInvariant()
    if ($value -match '^pack_(ecm|bod|dmm)$') { return [string]$matches[1] }
    return ''
}

function Get-TargetIdFromLegacyVolume([string]$Volume) {
    $value = ([string]$Volume).Trim().ToLowerInvariant()
    if ($value -eq 'none' -or $value -eq 'unassigned') { return 'unassigned' }
    if ($value -match '^(ja|en)-([a-z0-9][a-z0-9_-]{0,47})$') { return [string]$matches[2] }
    if ($value -match '^[a-z0-9][a-z0-9_-]{0,47}$') { return $value }
    return 'unassigned'
}

function Test-DirectPowerPointRelativePath([string]$RelativePath) {
    Test-RelativePath $RelativePath | Out-Null
    if ($RelativePath -match '[\/]') { throw '提出フォルダ直下のPowerPoint原稿だけ登録できます。子フォルダ内のファイルは対象外です。' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -ne '.pptx') { throw '拡張子が .pptx のPowerPoint原稿だけ登録できます。' }
    if ([IO.Path]::GetFileName($RelativePath) -like '~$*') { throw 'PowerPointの一時ファイルは登録できません。' }
    return $true
}

function Get-LegacyVolumeFromTargetId([string]$Language, [string]$TargetId) {
    $target = ([string]$TargetId).Trim().ToLowerInvariant()
    if ($target -in @('unassigned','none')) { return 'none' }
    if ($target -match '^[a-z0-9][a-z0-9_-]{0,47}$') { return "${Language}-$target" }
    return 'none'
}

function New-BuiltinPack([string]$Category, [string]$Language, $Existing = $null) {
    $cat = Require-WorkbookCategory $Category
    $display = $cat.ToUpperInvariant()
    return [pscustomobject][ordered]@{
        packId = Get-BuiltinPackId $cat
        templateId = "builtin-$cat"
        templateVersion = Get-IntDataProperty $Existing 'templateVersion' 1
        displayName = [string](Get-DataProperty $Existing 'displayName' $display)
        language = $Language
        createdAt = Get-DataProperty $Existing 'createdAt' (New-NowIso)
        updatedAt = Get-DataProperty $Existing 'updatedAt' $null
        settings = Get-DataProperty $Existing 'settings' ([ordered]@{})
        legacyCategory = $cat
        archivedAt = $null
        review = Get-DataProperty $Existing 'review' ([ordered]@{ status='draft'; submittedFingerprints=[ordered]@{}; submittedAt=''; submittedBy=''; approvedAt=''; approvedBy=''; note=''; events=@() })
    }
}

function ConvertTo-PackDisplayName([string]$DisplayName) {
    $name = ([string]$DisplayName).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { throw [ArgumentException]::new('一式名を入力してください。') }
    if ($name.Length -gt 120) { throw [ArgumentException]::new('一式名は120文字以内で入力してください。') }
    if ($name -match '[\x00-\x1f\x7f]') { throw [ArgumentException]::new('一式名に制御文字は使用できません。') }
    return $name
}

function Test-PackArchived($Pack) {
    return (-not [string]::IsNullOrWhiteSpace([string](Get-DataProperty $Pack 'archivedAt' '')))
}

function Get-PackRecord($Structure, [string]$PackId) {
    $id = ([string]$PackId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw [ArgumentException]::new('packId が必要です。') }
    $matches = @(Get-Array (Get-DataProperty $Structure 'packs' @()) | Where-Object { [string](Get-DataProperty $_ 'packId' '') -eq $id } | Select-Object -First 1)
    if ($matches.Count -eq 0) { throw [ArgumentException]::new('指定された一式が見つかりません。') }
    return $matches[0]
}

function Get-WorkbookPackId($Workbook) {
    if ($null -eq $Workbook) { return '' }
    $packId = ([string](Get-DataProperty $Workbook 'packId' '')).Trim()
    if (-not [string]::IsNullOrWhiteSpace($packId)) { return $packId }
    $category = Normalize-WorkbookCategory ([string](Get-DataProperty $Workbook 'category' '')) ([string](Get-DataProperty $Workbook 'fileName' ''))
    if (-not [string]::IsNullOrWhiteSpace($category)) { return Get-BuiltinPackId $category }
    return ''
}

function Resolve-DocumentPackScope($Structure, [string]$PackIdOrCategory, [bool]$AllowArchived = $false) {
    $value = ([string]$PackIdOrCategory).Trim()
    $category = Normalize-WorkbookCategory $value ''
    $packId = if (-not [string]::IsNullOrWhiteSpace($category)) { Get-BuiltinPackId $category } else { $value }
    $matches = @(Get-Array (Get-DataProperty $Structure 'packs' @()) | Where-Object { [string](Get-DataProperty $_ 'packId' '') -eq $packId } | Select-Object -First 1)
    if ($matches.Count -gt 0) { $pack = $matches[0] }
    elseif (-not [string]::IsNullOrWhiteSpace($category)) { $pack = New-BuiltinPack $category ([string](Get-DataProperty $Structure 'language' 'ja')) }
    else { throw [ArgumentException]::new('指定された一式が見つかりません。') }
    if ((Test-PackArchived $pack) -and -not $AllowArchived) { throw [InvalidOperationException]::new('アーカイブ済みの一式は操作できません。復元してからやり直してください。') }
    return [pscustomobject][ordered]@{
        packId = $packId
        category = Get-CategoryFromBuiltinPackId $packId
        pack = $pack
        builtIn = (-not [string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId $packId)))
    }
}

function Test-WorkbookPack($Workbook, [string]$PackId) {
    return ([string](Get-WorkbookPackId $Workbook) -eq ([string]$PackId).Trim())
}

function Get-PackOutputState($Structure, [string]$Language, [string]$PackId, [string]$Volume, [bool]$Create = $false) {
    $targetId = Get-TargetIdFromLegacyVolume $Volume
    $scope = Resolve-DocumentPackScope $Structure $PackId $true
    if (@(Get-PackTargetIds $Language $scope.pack) -notcontains $targetId) { throw [ArgumentException]::new('この一式に存在しない出力先です。') }
    $category = Get-CategoryFromBuiltinPackId $PackId
    if (-not [string]::IsNullOrWhiteSpace($category)) {
        $state = Get-DataProperty $Structure.volumes (Get-VolumeStateKey $Volume $category) $null
        if ($null -eq $state -and $Create) {
            $state = New-EmptyVolumeState
            Set-NoteProperty $Structure.volumes (Get-VolumeStateKey $Volume $category) $state
        }
        return $state
    }
    $outputs = Get-DataProperty $Structure 'outputs' $null
    if ($null -eq $outputs) {
        if (-not $Create) { return $null }
        $outputs = [ordered]@{}
        Set-NoteProperty $Structure 'outputs' $outputs
    }
    $key = "$PackId|$targetId"
    $state = Get-DataProperty $outputs $key $null
    if ($null -eq $state -and $Create) {
        $empty = New-EmptyVolumeState
        $state = [pscustomobject][ordered]@{ packId=$PackId; targetId=$targetId; status=[string]$empty.status; lastBuiltAt=$null; outputPdf=$null; builtFingerprint=''; staleReasons=@(); message=''; buildId='' }
        Set-NoteProperty $outputs $key $state
    }
    return $state
}

function Assert-PackDisplayNameAvailable($Structure, [string]$DisplayName, [string]$ExceptPackId = '') {
    $name = ConvertTo-PackDisplayName $DisplayName
    foreach ($pack in @(Get-Array (Get-DataProperty $Structure 'packs' @()))) {
        if ([string](Get-DataProperty $pack 'packId' '') -eq $ExceptPackId) { continue }
        if ([string]::Equals(([string](Get-DataProperty $pack 'displayName' '')).Trim(), $name, [StringComparison]::OrdinalIgnoreCase)) {
            throw [ArgumentException]::new('同じ名前の一式が既にあります。別の名前を入力してください。')
        }
    }
    return $name
}

function Get-UniquePackDisplayName($Structure, [string]$BaseName, [string]$Language) {
    $root = ConvertTo-PackDisplayName $BaseName
    $suffix = $(if ($Language -eq 'en') { 'Copy' } else { 'コピー' })
    for ($number = 1; $number -le 999; $number++) {
        $label = $(if ($number -eq 1) { "($suffix)" } else { "($suffix $number)" })
        $rootLimit = 120 - $label.Length - 1
        $candidateRoot = $(if ($root.Length -gt $rootLimit) { $root.Substring(0, $rootLimit).TrimEnd() } else { $root })
        if ([string]::IsNullOrWhiteSpace($candidateRoot)) { $candidateRoot = 'Pack' }
        $candidate = "$candidateRoot $label"
        $exists = @(
            Get-Array (Get-DataProperty $Structure 'packs' @()) |
                Where-Object { [string]::Equals(([string](Get-DataProperty $_ 'displayName' '')).Trim(), $candidate, [StringComparison]::OrdinalIgnoreCase) }
        ).Count -gt 0
        if (-not $exists) { return $candidate }
    }
    throw '複製した一式に一意な名前を付けられませんでした。'
}

function New-CustomPackId($Structure) {
    $existing = @{}; foreach ($pack in @(Get-Array (Get-DataProperty $Structure 'packs' @()))) { $existing[[string](Get-DataProperty $pack 'packId' '')] = $true }
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $id = 'pack_' + ([Guid]::NewGuid().ToString('N').Substring(0,16))
        if (-not $existing.ContainsKey($id)) { return $id }
    }
    throw '一意な一式IDを生成できませんでした。'
}

function Get-DefaultPackSettings([string]$Language, [string]$DisplayName) {
    $name = ([string]$DisplayName).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'ReportBinder' }
    return [ordered]@{
        documentTitle = $name
        documentSubtitle = ''
        outputFileNamePattern = '{packName}_{targetName}_{yyyyMMdd}.pdf'
    }
}

function Get-ResolvedPackSettings($Pack, [string]$Language) {
    $defaults = Get-DefaultPackSettings $Language ([string](Get-DataProperty $Pack 'displayName' 'ReportBinder'))
    $saved = Get-DataProperty $Pack 'settings' $null
    foreach ($name in @('documentTitle','documentSubtitle','outputFileNamePattern')) {
        $value = [string](Get-DataProperty $saved $name (Get-DataProperty $defaults $name ''))
        $defaults[$name] = $value
    }
    if ([string]::IsNullOrWhiteSpace([string]$defaults.documentTitle)) { $defaults.documentTitle = [string](Get-DataProperty $Pack 'displayName' 'ReportBinder') }
    if ([string]::IsNullOrWhiteSpace([string]$defaults.outputFileNamePattern)) { $defaults.outputFileNamePattern = '{packName}_{targetName}_{yyyyMMdd}.pdf' }
    return $defaults
}

function Get-UpdatedPackSettings($Pack, [string]$Language, $Patch) {
    $settings = Get-ResolvedPackSettings $Pack $Language
    foreach ($name in @('documentTitle','documentSubtitle','outputFileNamePattern')) {
        if ($null -ne (Get-DataProperty $Patch $name $null)) { $settings[$name] = ([string](Get-DataProperty $Patch $name '')).Trim() }
    }
    if ([string]::IsNullOrWhiteSpace([string]$settings.documentTitle)) { throw [ArgumentException]::new('資料タイトルを入力してください。') }
    [void](Resolve-OutputFileNamePattern ([string]$settings.outputFileNamePattern) 'PROJECT' ([string](Get-DataProperty $Pack 'displayName' 'ReportBinder')) (Get-TargetDisplayName $Language ("$Language-main")))
    return $settings
}

function ConvertTo-NormalizedPageRange($Value) {
    if ($null -eq $Value) { return $null }
    $start = 0; $end = 0
    if ($Value -is [string]) {
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        if ($text -notmatch '^(?<start>\d+)(?:\s*-\s*(?<end>\d+))?$') { throw [ArgumentException]::new('ページ範囲は「2」または「2-5」の形式で入力してください。') }
        $start = [int]$matches.start; $end = if ($matches.end) { [int]$matches.end } else { $start }
    } else {
        $start = Get-IntDataProperty $Value 'start' 0
        $end = Get-IntDataProperty $Value 'end' $start
    }
    if ($start -lt 1 -or $end -lt $start) { throw [ArgumentException]::new('ページ範囲は1以上で、開始ページが終了ページを超えないようにしてください。') }
    return [ordered]@{ start = $start; end = $end }
}

function Get-TargetDisplayName([string]$Language, [string]$Volume) {
    if ($Volume -match 'appendix$') { return $(if ($Language -eq 'en') { 'Appendix' } else { '補足' }) }
    return $(if ($Language -eq 'en') { 'Main' } else { '本体' })
}

function Resolve-OutputFileNamePattern([string]$Pattern, [string]$ProjectId, [string]$PackName, [string]$TargetName) {
    $value = ([string]$Pattern).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { $value = '{projectId}_{targetName}.pdf' }
    $unknown = [regex]::Matches($value, '\{[^{}]+\}') | ForEach-Object { $_.Value } | Where-Object { $_ -notin @('{projectId}','{packName}','{targetName}','{yyyyMMdd}') } | Select-Object -Unique
    if (@($unknown).Count -gt 0) { throw [ArgumentException]::new("未対応のファイル名プレースホルダーです: $(@($unknown) -join ', ')") }
    $resolved = $value.Replace('{projectId}', $ProjectId).Replace('{packName}', $PackName).Replace('{targetName}', $TargetName).Replace('{yyyyMMdd}', (Get-Date -Format 'yyyyMMdd'))
    if ($resolved.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $resolved -match '[\\/]') { throw [ArgumentException]::new('出力ファイル名に使用できない文字が含まれています。') }
    if (-not $resolved.EndsWith('.pdf', [StringComparison]::OrdinalIgnoreCase)) { $resolved += '.pdf' }
    return $resolved
}

function Update-PackSettings([string]$Language, [string]$PackId, $Patch) {
    $id = ([string]$PackId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw [ArgumentException]::new('packId が必要です。') }
    return Update-StructureLocked $Language {
        param($structure)
        $pack = Get-PackRecord $structure $id
        if ($null -ne (Get-DataProperty $Patch 'displayName' $null)) {
            $displayName = Assert-PackDisplayNameAvailable $structure ([string](Get-DataProperty $Patch 'displayName' '')) $id
            Set-NoteProperty $pack 'displayName' $displayName
        }
        $settingsPatch = Get-DataProperty $Patch 'settings' $Patch
        $settings = Get-UpdatedPackSettings $pack $Language $settingsPatch
        Set-NoteProperty $pack 'settings' $settings
        Set-NoteProperty $pack 'updatedAt' (New-NowIso)
        Mark-VolumeNeedsRebuild $structure $Language $id @(Get-PackVolumeList $Language $pack $false) 'document-settings' '資料の仕上げ設定を変更しました'
        return [ordered]@{ packId=$id; displayName=[string]$pack.displayName; settings=$settings; archived=(Test-PackArchived $pack); updatedAt=(Get-DataProperty $pack 'updatedAt' $null) }
    }
}

function Get-PackTemplateDir([string]$Language) {
    $dataDir = ([string](Get-DataProperty (Get-Paths) 'dataDir' '')).Trim()
    if ([string]::IsNullOrWhiteSpace($dataDir)) { return '' }
    return (Join-Path (Get-WorkspacePath $Language $dataDir) 'templates')
}

function Get-LocalizedTemplateText($Value, [string]$Language, [string]$Fallback = '') {
    if ($null -eq $Value) { return $Fallback }
    if ($Value -is [string]) {
        $text = ([string]$Value).Trim()
        return $(if ([string]::IsNullOrWhiteSpace($text)) { $Fallback } else { $text })
    }
    $localized = [string](Get-DataProperty $Value $Language '')
    if ([string]::IsNullOrWhiteSpace($localized)) { $localized = [string](Get-DataProperty $Value 'ja' (Get-DataProperty $Value 'en' $Fallback)) }
    return $(if ([string]::IsNullOrWhiteSpace($localized)) { $Fallback } else { $localized.Trim() })
}

function ConvertTo-NormalizedPackTemplate([string]$Language, $Template, [string]$TemplateId = '') {
    $displayName = Get-LocalizedTemplateText (Get-DataProperty $Template 'displayName' '') $Language ''
    if ([string]::IsNullOrWhiteSpace($displayName)) { throw [ArgumentException]::new('テンプレート名を入力してください。') }
    if ($displayName.Length -gt 120) { throw [ArgumentException]::new('テンプレート名は120文字以内で入力してください。') }
    $description = Get-LocalizedTemplateText (Get-DataProperty $Template 'description' '') $Language ''
    if ($description.Length -gt 500) { throw [ArgumentException]::new('テンプレートの説明は500文字以内で入力してください。') }
    $sourceTypes = @(
        Get-Array (Get-DataProperty $Template 'acceptedSourceTypes' @('excel','word','pdf','powerpoint')) |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { $_ -in @('excel','word','pdf','powerpoint') } |
            Select-Object -Unique
    )
    if ($sourceTypes.Count -eq 0) { throw [ArgumentException]::new('テンプレートで扱う原稿形式を1つ以上選んでください。') }
    $targetInput = @(Get-Array (Get-DataProperty $Template 'targets' @()))
    if ($targetInput.Count -eq 0) {
        $targetInput = @(
            [pscustomobject][ordered]@{ targetId='main'; displayName=$(if ($Language -eq 'en') { 'Main' } else { '本体' }); required=$true },
            [pscustomobject][ordered]@{ targetId='appendix'; displayName=$(if ($Language -eq 'en') { 'Appendix' } else { '補足' }); required=$false }
        )
    }
    if ($targetInput.Count -gt 12) { throw [ArgumentException]::new('出力先は12件以内で設定してください。') }
    $targets = @()
    $targetIds = @{}
    foreach ($targetInputItem in $targetInput) {
        $targetId = Assert-SafeStorageSegment (([string](Get-DataProperty $targetInputItem 'targetId' '')).Trim().ToLowerInvariant()) 'targetId'
        if ($targetId -in @('none','unassigned')) { throw [ArgumentException]::new('none と unassigned は出力先IDに使用できません。') }
        if ($targetId.Length -gt 48) { throw [ArgumentException]::new('出力先IDは48文字以内で入力してください。') }
        if ($targetIds.ContainsKey($targetId)) { throw [ArgumentException]::new('出力先IDが重複しています。') }
        $targetIds[$targetId] = $true
        $fallback = $(if ($targetId -eq 'main') { $(if ($Language -eq 'en') { 'Main' } else { '本体' }) } elseif ($targetId -eq 'appendix') { $(if ($Language -eq 'en') { 'Appendix' } else { '補足' }) } else { $targetId })
        $label = Get-LocalizedTemplateText (Get-DataProperty $targetInputItem 'displayName' '') $Language $fallback
        if ([string]::IsNullOrWhiteSpace($label) -or $label.Length -gt 60) { throw [ArgumentException]::new('出力先名は1～60文字で入力してください。') }
        $targets += [pscustomobject][ordered]@{
            targetId = $targetId
            displayName = $label
            required = [bool](Get-DataProperty $targetInputItem 'required' ($targetId -eq 'main'))
        }
    }
    $requirements = @()
    $requirementIds = @{}
    $requirementIndex = 0
    foreach ($requirementInput in @(Get-Array (Get-DataProperty $Template 'sourceRequirements' @()))) {
        $requirementIndex++
        $requirementName = Get-LocalizedTemplateText (Get-DataProperty $requirementInput 'displayName' '') $Language ''
        if ([string]::IsNullOrWhiteSpace($requirementName)) { throw [ArgumentException]::new('必要原稿名を入力してください。') }
        if ($requirementName.Length -gt 120) { throw [ArgumentException]::new('必要原稿名は120文字以内で入力してください。') }
        $requirementId = ([string](Get-DataProperty $requirementInput 'requirementId' '')).Trim()
        if ([string]::IsNullOrWhiteSpace($requirementId)) { $requirementId = 'requirement_' + (Get-Sha256Text "$requirementIndex|$requirementName").Substring(7,16) }
        $requirementId = Assert-SafeStorageSegment $requirementId 'requirementId'
        if ($requirementIds.ContainsKey($requirementId)) { throw [ArgumentException]::new('必要原稿IDが重複しています。') }
        $requirementIds[$requirementId] = $true
        $requirementTypes = @(
            Get-Array (Get-DataProperty $requirementInput 'acceptedSourceTypes' $sourceTypes) |
                ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
                Where-Object { $_ -in $sourceTypes } |
                Select-Object -Unique
        )
        if ($requirementTypes.Count -eq 0) { throw [ArgumentException]::new("必要原稿「$requirementName」で扱う原稿形式を1つ以上選んでください。") }
        $requirementOwner = ([string](Get-DataProperty $requirementInput 'ownerDepartment' '')).Trim()
        if ($requirementOwner.Length -gt 120) { throw [ArgumentException]::new('必要原稿の担当部署は120文字以内で入力してください。') }
        $requirementTarget = ([string](Get-DataProperty $requirementInput 'defaultTargetId' 'unassigned')).Trim().ToLowerInvariant()
        if ($requirementTarget -ne 'unassigned' -and -not $targetIds.ContainsKey($requirementTarget)) { throw [ArgumentException]::new('必要原稿の既定出力先が不正です。') }
        $dueDate = ([string](Get-DataProperty $requirementInput 'dueDate' '')).Trim()
        if (-not [string]::IsNullOrWhiteSpace($dueDate) -and $dueDate -notmatch '^\d{4}-\d{2}-\d{2}$') { throw [ArgumentException]::new('必要原稿の期限はYYYY-MM-DD形式で入力してください。') }
        $requirements += [pscustomobject][ordered]@{
            requirementId = $requirementId
            displayName = $requirementName
            ownerDepartment = $requirementOwner
            required = [bool](Get-DataProperty $requirementInput 'required' $true)
            acceptedSourceTypes = @($requirementTypes)
            defaultTargetId = $requirementTarget
            dueDate = $dueDate
        }
    }
    $rulesInput = Get-DataProperty $Template 'rules' ([ordered]@{})
    $newDestination = ([string](Get-DataProperty $rulesInput 'newItemDestination' 'unassigned')).Trim().ToLowerInvariant()
    if ($newDestination -ne 'unassigned' -and -not $targetIds.ContainsKey($newDestination)) { throw [ArgumentException]::new('新規原稿の既定先が不正です。') }
    $outputInput = Get-DataProperty $Template 'output' ([ordered]@{})
    $filePattern = ([string](Get-DataProperty $outputInput 'fileNamePattern' '{packName}_{targetName}_{yyyyMMdd}.pdf')).Trim()
    if ([string]::IsNullOrWhiteSpace($filePattern) -or $filePattern.Length -gt 180) { throw [ArgumentException]::new('出力ファイル名パターンは1～180文字で入力してください。') }
    [void](Resolve-OutputFileNamePattern $filePattern 'PROJECT' 'PACK' 'TARGET')
    $id = ([string]$TemplateId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { $id = 'template_' + ([Guid]::NewGuid().ToString('N').Substring(0,16)) }
    $id = Assert-SafeStorageSegment $id 'templateId'
    if ($id -like 'builtin-*') { throw [ArgumentException]::new('builtin-で始まるテンプレートIDは利用者定義に使用できません。') }
    return [pscustomobject][ordered]@{
        templateSchemaVersion = 1
        templateId = $id
        templateVersion = [Math]::Max(1, (Get-IntDataProperty $Template 'templateVersion' 1))
        displayName = $displayName
        description = $description
        acceptedSourceTypes = @($sourceTypes)
        sourceRequirements = @($requirements)
        targets = @($targets)
        rules = [pscustomobject][ordered]@{
            newItemDestination = $newDestination
            retainManualOrder = [bool](Get-DataProperty $rulesInput 'retainManualOrder' $true)
            blockBuildWhenRequiredSourceIsStale = [bool](Get-DataProperty $rulesInput 'blockBuildWhenRequiredSourceIsStale' $true)
            blockBuildWhenRequiredSourceFailed = [bool](Get-DataProperty $rulesInput 'blockBuildWhenRequiredSourceFailed' $true)
        }
        output = [pscustomobject][ordered]@{
            fileNamePattern = $filePattern
            pageNumbering = [string](Get-DataProperty $outputInput 'pageNumbering' 'continuous')
            bookmarks = [string](Get-DataProperty $outputInput 'bookmarks' 'from-items')
        }
        workflowAvailable = $true
        builtIn = $false
        editable = $true
        deletable = $true
        updatedAt = New-NowIso
    }
}

function Get-UserPackTemplates([string]$Language) {
    $dir = Get-PackTemplateDir $Language
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir)) { return @() }
    $result = @()
    foreach ($file in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name)) {
        try {
            $raw = Read-JsonFile $file.FullName $null
            if ($null -eq $raw) { continue }
            $normalized = ConvertTo-NormalizedPackTemplate $Language $raw ([string](Get-DataProperty $raw 'templateId' $file.BaseName))
            Set-NoteProperty $normalized 'createdAt' (Get-DataProperty $raw 'createdAt' $null)
            Set-NoteProperty $normalized 'updatedAt' (Get-DataProperty $raw 'updatedAt' $file.LastWriteTimeUtc.ToString('o'))
            $result += $normalized
        } catch { Write-Warning ('利用者定義テンプレートを読み込めません: ' + $file.Name + ' / ' + $_.Exception.Message) }
    }
    return @($result)
}

function Save-PackTemplate([string]$Language, $Request, [string]$TemplateId = '') {
    $existing = $null
    if (-not [string]::IsNullOrWhiteSpace($TemplateId)) {
        $matches = @(Get-UserPackTemplates $Language | Where-Object { [string]$_.templateId -eq $TemplateId } | Select-Object -First 1)
        if ($matches.Count -eq 0) { throw [ArgumentException]::new('編集するテンプレートが見つかりません。') }
        $existing = $matches[0]
    }
    $merged = [ordered]@{}
    foreach ($name in @('displayName','description','acceptedSourceTypes','sourceRequirements','targets','rules','output')) {
        $value = Get-DataProperty $Request $name $null
        if ($null -eq $value -and $null -ne $existing) { $value = Get-DataProperty $existing $name $null }
        if ($null -ne $value) { $merged[$name] = $value }
    }
    $nextVersion = if ($null -eq $existing) { 1 } else { (Get-IntDataProperty $existing 'templateVersion' 1) + 1 }
    $merged.templateVersion = $nextVersion
    $template = ConvertTo-NormalizedPackTemplate $Language ([pscustomobject]$merged) $TemplateId
    $dir = Get-PackTemplateDir $Language
    if ([string]::IsNullOrWhiteSpace($dir)) { throw [ArgumentException]::new('先に提出フォルダと管理データフォルダを設定してください。') }
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $createdAt = if ($null -ne $existing) { Get-DataProperty $existing 'createdAt' (New-NowIso) } else { New-NowIso }
    Set-NoteProperty $template 'createdAt' $createdAt
    Set-NoteProperty $template 'updatedAt' (New-NowIso)
    Write-JsonFile (Join-Path $dir (([string]$template.templateId) + '.json')) $template
    return $template
}

function Remove-PackTemplate([string]$Language, [string]$TemplateId) {
    $id = Assert-SafeStorageSegment $TemplateId 'templateId'
    if ($id -like 'builtin-*') { throw [ArgumentException]::new('組み込みテンプレートは削除できません。') }
    $structure = Get-Structure $Language
    $packsUsingTemplate = @(
        Get-Array (Get-DataProperty $structure 'packs' @()) |
            Where-Object { [string](Get-DataProperty $_ 'templateId' '') -eq $id }
    )
    if ($packsUsingTemplate.Count -gt 0) {
        throw [ArgumentException]::new('このテンプレートを使用している一式があります。先に対象パックを変更またはアーカイブしてください。')
    }
    $path = Join-Path (Get-PackTemplateDir $Language) ($id + '.json')
    if (-not (Test-Path -LiteralPath $path)) { throw [ArgumentException]::new('削除するテンプレートが見つかりません。') }
    Remove-Item -LiteralPath $path -Force
    return [ordered]@{ templateId=$id; removed=$true }
}

function Get-PackTemplateUpgradePreview([string]$Language, [string]$PackId) {
    $structure = Get-Structure $Language
    $pack = Get-PackRecord $structure $PackId
    $templateId = [string](Get-DataProperty $pack 'templateId' '')
    $latest = Get-PackTemplateRecord $Language $templateId
    $current = Get-PackEffectiveTemplate $Language $pack
    $fromVersion = Get-IntDataProperty $pack 'templateVersion' (Get-IntDataProperty $current 'templateVersion' 1)
    $toVersion = Get-IntDataProperty $latest 'templateVersion' 1
    $changes = @()
    $currentTargets = ConvertTo-Json (Get-DataProperty $current 'targets' @()) -Depth 10 -Compress
    $latestTargets = ConvertTo-Json (Get-DataProperty $latest 'targets' @()) -Depth 10 -Compress
    if ($currentTargets -ne $latestTargets) { $changes += '出力先の名前・必須設定' }
    $currentRequirements = ConvertTo-Json (Get-DataProperty $current 'sourceRequirements' @()) -Depth 10 -Compress
    $latestRequirements = ConvertTo-Json (Get-DataProperty $latest 'sourceRequirements' @()) -Depth 10 -Compress
    if ($currentRequirements -ne $latestRequirements) { $changes += '必要原稿リスト' }
    $currentRules = ConvertTo-Json (Get-DataProperty $current 'rules' ([ordered]@{})) -Depth 10 -Compress
    $latestRules = ConvertTo-Json (Get-DataProperty $latest 'rules' ([ordered]@{})) -Depth 10 -Compress
    if ($currentRules -ne $latestRules) { $changes += '新規ページ配置・出力条件' }
    $currentOutput = ConvertTo-Json (Get-DataProperty $current 'output' ([ordered]@{})) -Depth 10 -Compress
    $latestOutput = ConvertTo-Json (Get-DataProperty $latest 'output' ([ordered]@{})) -Depth 10 -Compress
    if ($currentOutput -ne $latestOutput) { $changes += '出力既定値' }
    return [pscustomobject][ordered]@{
        packId = [string](Get-DataProperty $pack 'packId' '')
        packName = [string](Get-DataProperty $pack 'displayName' '')
        templateId = $templateId
        templateName = [string](Get-DataProperty $latest 'displayName' '')
        fromVersion = $fromVersion
        toVersion = $toVersion
        updateAvailable = ($toVersion -gt $fromVersion -or $changes.Count -gt 0)
        changes = @($changes)
        current = New-PackTemplateSnapshot $Language $current
        latest = New-PackTemplateSnapshot $Language $latest
    }
}

function Update-PackTemplateSnapshot([string]$Language, [string]$PackId) {
    return Update-StructureLocked $Language {
        param($structure)
        $pack = Get-PackRecord $structure $PackId
        $latest = Get-PackTemplateRecord $Language ([string](Get-DataProperty $pack 'templateId' ''))
        $currentVersion = Get-IntDataProperty $pack 'templateVersion' 1
        $latestVersion = Get-IntDataProperty $latest 'templateVersion' 1
        $oldTargetIds = @(Get-PackTargetIds $Language $pack)
        $newTargetIds = @(Get-Array (Get-DataProperty $latest 'targets' @()) | ForEach-Object { [string](Get-DataProperty $_ 'targetId' '') })
        $removedTargetIds = @($oldTargetIds | Where-Object { $newTargetIds -notcontains $_ })
        $unassignedPageCount = 0
        if ($removedTargetIds.Count -gt 0) {
            foreach ($page in @(Get-Array (Get-DataProperty $structure 'pages' @()))) {
                $workbook = @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq [string](Get-DataProperty $page 'workbookId' '') } | Select-Object -First 1)
                if ($workbook.Count -eq 0 -or -not (Test-WorkbookPack $workbook[0] ([string]$pack.packId))) { continue }
                $pageTargetId = Get-TargetIdFromLegacyVolume ([string](Get-DataProperty $page 'volume' 'none'))
                if ($removedTargetIds -notcontains $pageTargetId) { continue }
                Set-NoteProperty $page 'volume' 'none'; Set-NoteProperty $page 'enabled' $false; Set-NoteProperty $page 'updatedAt' (New-NowIso)
                $unassignedPageCount++
            }
            $outputs = Get-DataProperty $structure 'outputs' ([ordered]@{})
            foreach ($removedTargetId in $removedTargetIds) {
                $oldState = Get-DataProperty $outputs "$([string]$pack.packId)|$removedTargetId" $null
                if ($null -ne $oldState) { Set-NoteProperty $oldState 'status' 'needs-rebuild'; Add-StaleReason $oldState 'template-target-removed' 'ひな形から出力先が削除されました' }
            }
        }
        Set-NoteProperty $pack 'templateConfig' (New-PackTemplateSnapshot $Language $latest)
        Set-NoteProperty $pack 'templateVersion' $latestVersion
        Set-NoteProperty $pack 'updatedAt' (New-NowIso)
        Add-PackOutputStates $structure $pack $latest
        Mark-VolumeNeedsRebuild $structure $Language ([string]$pack.packId) @(Get-PackVolumeList $Language $pack $false) 'template-updated' '一式のひな形を更新しました'
        return [pscustomobject][ordered]@{ packId=[string]$pack.packId; templateId=[string]$pack.templateId; fromVersion=$currentVersion; toVersion=$latestVersion; removedTargetIds=@($removedTargetIds); unassignedPageCount=$unassignedPageCount; templateConfig=(New-PackTemplateSnapshot $Language $latest) }
    }
}

function Get-PackTemplateCatalog([string]$Language) {
    if ($Language -notin @('ja','en')) { $Language = 'ja' }
    $genericTargets = @(
        [pscustomobject][ordered]@{ targetId = 'main'; displayName = $(if ($Language -eq 'en') { 'Main' } else { '本体' }); required = $true },
        [pscustomobject][ordered]@{ targetId = 'appendix'; displayName = $(if ($Language -eq 'en') { 'Appendix' } else { '補足' }); required = $false }
    )
    $generic = [pscustomobject][ordered]@{
        templateId = 'builtin-generic-department-pack'
        templateVersion = 1
        packId = ''
        displayName = $(if ($Language -eq 'en') { 'Department document pack' } else { '部門一式' })
        description = $(if ($Language -eq 'en') { 'A general-purpose pack for Excel, Word, PowerPoint, and PDF source documents' } else { 'Excel・Word・PowerPoint・PDFをまとめる汎用一式' })
        acceptedSourceTypes = @('excel','word','pdf','powerpoint')
        targets = @($genericTargets)
        rules = [pscustomobject][ordered]@{ newItemDestination='unassigned'; retainManualOrder=$true; blockBuildWhenRequiredSourceIsStale=$true; blockBuildWhenRequiredSourceFailed=$true }
        output = [pscustomobject][ordered]@{ fileNamePattern='{packName}_{targetName}_{yyyyMMdd}.pdf'; pageNumbering='continuous'; bookmarks='from-items' }
        workflowAvailable = $true
        builtIn = $true
    }
    # ECM/BOD/DMM were migration-era categories, not choices a new user should
    # have to understand. Keep their internal compatibility records, but expose
    # only the general-purpose starting point and user-created templates.
    return @($generic) + @(Get-UserPackTemplates $Language)
}

function New-PackTemplateSnapshot([string]$Language, $Template) {
    $rules = Get-DataProperty $Template 'rules' ([ordered]@{})
    $output = Get-DataProperty $Template 'output' ([ordered]@{})
    return [pscustomobject][ordered]@{
        templateSchemaVersion = Get-IntDataProperty $Template 'templateSchemaVersion' 1
        templateId = [string](Get-DataProperty $Template 'templateId' '')
        templateVersion = Get-IntDataProperty $Template 'templateVersion' 1
        displayName = [string](Get-DataProperty $Template 'displayName' '')
        description = [string](Get-DataProperty $Template 'description' '')
        acceptedSourceTypes = [object[]](Get-Array (Get-DataProperty $Template 'acceptedSourceTypes' @('excel','word','pdf','powerpoint')))
        sourceRequirements = [object[]]@(
            Get-Array (Get-DataProperty $Template 'sourceRequirements' @()) | ForEach-Object {
                [pscustomobject][ordered]@{
                    requirementId = [string](Get-DataProperty $_ 'requirementId' '')
                    displayName = [string](Get-DataProperty $_ 'displayName' '')
                    ownerDepartment = [string](Get-DataProperty $_ 'ownerDepartment' '')
                    required = [bool](Get-DataProperty $_ 'required' $true)
                    acceptedSourceTypes = [object[]](Get-Array (Get-DataProperty $_ 'acceptedSourceTypes' @('excel','word','pdf','powerpoint')))
                    defaultTargetId = [string](Get-DataProperty $_ 'defaultTargetId' 'unassigned')
                    dueDate = [string](Get-DataProperty $_ 'dueDate' '')
                }
            }
        )
        targets = [object[]]@(
            Get-Array (Get-DataProperty $Template 'targets' @()) | ForEach-Object {
                [pscustomobject][ordered]@{
                    targetId = [string](Get-DataProperty $_ 'targetId' '')
                    displayName = [string](Get-DataProperty $_ 'displayName' '')
                    required = [bool](Get-DataProperty $_ 'required' $false)
                }
            }
        )
        rules = [pscustomobject][ordered]@{
            newItemDestination = [string](Get-DataProperty $rules 'newItemDestination' 'unassigned')
            retainManualOrder = [bool](Get-DataProperty $rules 'retainManualOrder' $true)
            blockBuildWhenRequiredSourceIsStale = [bool](Get-DataProperty $rules 'blockBuildWhenRequiredSourceIsStale' $true)
            blockBuildWhenRequiredSourceFailed = [bool](Get-DataProperty $rules 'blockBuildWhenRequiredSourceFailed' $true)
        }
        output = [pscustomobject][ordered]@{
            fileNamePattern = [string](Get-DataProperty $output 'fileNamePattern' '{packName}_{targetName}_{yyyyMMdd}.pdf')
            pageNumbering = [string](Get-DataProperty $output 'pageNumbering' 'continuous')
            bookmarks = [string](Get-DataProperty $output 'bookmarks' 'from-items')
        }
    }
}

function Get-PackEffectiveTemplate([string]$Language, $Pack) {
    $saved = Get-DataProperty $Pack 'templateConfig' $null
    if ($null -ne $saved) { return (New-PackTemplateSnapshot $Language $saved) }
    return Get-PackTemplateRecord $Language ([string](Get-DataProperty $Pack 'templateId' 'builtin-generic-department-pack'))
}

function Get-PackTargetIds([string]$Language, $Pack) {
    $ids = @(
        Get-Array (Get-DataProperty (Get-PackEffectiveTemplate $Language $Pack) 'targets' @()) |
            ForEach-Object { ([string](Get-DataProperty $_ 'targetId' '')).Trim().ToLowerInvariant() } |
            Where-Object { $_ -match '^[a-z0-9][a-z0-9_-]{0,47}$' -and $_ -notin @('none','unassigned') } |
            Select-Object -Unique
    )
    if ($ids.Count -eq 0) { return @('main','appendix') }
    return @($ids)
}

function Get-PackVolumeList([string]$Language, $Pack, [bool]$IncludeUnassigned = $true) {
    $volumes = @(Get-PackTargetIds $Language $Pack | ForEach-Object { Get-LegacyVolumeFromTargetId $Language $_ })
    if ($IncludeUnassigned) { $volumes += 'none' }
    return @($volumes)
}

function Get-AllowedPackVolumes($Structure, [string]$Language, [string]$PackIdOrCategory, [bool]$IncludeUnassigned = $true) {
    $scope = Resolve-DocumentPackScope $Structure $PackIdOrCategory $true
    return @(Get-PackVolumeList $Language $scope.pack $IncludeUnassigned)
}

function Assert-PackTargetId([string]$Language, $Pack, [string]$TargetId) {
    $id = ([string]$TargetId).Trim().ToLowerInvariant()
    if (@(Get-PackTargetIds $Language $Pack) -notcontains $id) { throw [ArgumentException]::new('この一式に存在しない出力先です。') }
    return $id
}

function Get-PackTargetDisplayName([string]$Language, $Pack, [string]$TargetId) {
    $target = @(
        Get-Array (Get-DataProperty (Get-PackEffectiveTemplate $Language $Pack) 'targets' @()) |
            Where-Object { [string](Get-DataProperty $_ 'targetId' '') -eq $TargetId } |
            Select-Object -First 1
    )
    if ($target.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string](Get-DataProperty $target[0] 'displayName' ''))) {
        return [string](Get-DataProperty $target[0] 'displayName' '')
    }
    return Get-TargetDisplayName $Language (Get-LegacyVolumeFromTargetId $Language $TargetId)
}

function Get-NewItemTargetId($Structure, [string]$Language, $Workbook) {
    $sourceDefault = ([string](Get-DataProperty $Workbook 'defaultTargetId' 'unassigned')).Trim().ToLowerInvariant()
    $packId = Get-WorkbookPackId $Workbook
    $pack = @(Get-Array (Get-DataProperty $Structure 'packs' @()) | Where-Object { [string](Get-DataProperty $_ 'packId' '') -eq $packId } | Select-Object -First 1)
    if ($pack.Count -eq 0) { return 'unassigned' }
    $allowed = @(Get-PackTargetIds $Language $pack[0])
    if ($allowed -contains $sourceDefault) { return $sourceDefault }
    $template = Get-PackEffectiveTemplate $Language $pack[0]
    $targetId = ([string](Get-DataProperty (Get-DataProperty $template 'rules' ([ordered]@{})) 'newItemDestination' 'unassigned')).Trim().ToLowerInvariant()
    return $(if ($allowed -contains $targetId) { $targetId } else { 'unassigned' })
}

function Get-PublicPackList($Structure, [string]$Language, [bool]$IncludeArchived = $false) {
    $latestTemplates = @{}; foreach ($catalogTemplate in @(Get-PackTemplateCatalog $Language)) { $latestTemplates[[string](Get-DataProperty $catalogTemplate 'templateId' '')] = $catalogTemplate }
    $result = @()
    foreach ($pack in @(Get-Array (Get-DataProperty $Structure 'packs' @()))) {
        # Built-in category packs exist only as an internal legacy mirror. The
        # public application model has one explicit unit: a user-created pack.
        if (-not [string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId ([string](Get-DataProperty $pack 'packId' ''))))) { continue }
        $archived = Test-PackArchived $pack
        if ($archived -and -not $IncludeArchived) { continue }
        $templateId = [string](Get-DataProperty $pack 'templateId' '')
        $template = Get-PackEffectiveTemplate $Language $pack
        $availableTemplateVersion = if ($latestTemplates.ContainsKey($templateId)) { Get-IntDataProperty $latestTemplates[$templateId] 'templateVersion' 1 } else { Get-IntDataProperty $pack 'templateVersion' 1 }
        $category = Get-CategoryFromBuiltinPackId ([string](Get-DataProperty $pack 'packId' ''))
        $result += [pscustomobject][ordered]@{
            packId = [string](Get-DataProperty $pack 'packId' '')
            templateId = $templateId
            templateVersion = Get-IntDataProperty $pack 'templateVersion' 1
            availableTemplateVersion = $availableTemplateVersion
            templateUpdateAvailable = ($availableTemplateVersion -gt (Get-IntDataProperty $pack 'templateVersion' 1))
            displayName = [string](Get-DataProperty $pack 'displayName' '')
            language = [string](Get-DataProperty $pack 'language' $Language)
            category = $category
            acceptedSourceTypes = [object[]](Get-Array (Get-DataProperty $template 'acceptedSourceTypes' @()))
            sourceRequirements = [object[]](Get-Array (Get-DataProperty $template 'sourceRequirements' @()))
            targets = [object[]](Get-Array (Get-DataProperty $template 'targets' @()))
            rules = Get-DataProperty $template 'rules' ([ordered]@{})
            workflowAvailable = (-not $archived)
            builtIn = (-not [string]::IsNullOrWhiteSpace($category))
            duplicable = $true
            renamable = $true
            archivable = [string]::IsNullOrWhiteSpace($category)
            archived = $archived
            archivedAt = Get-DataProperty $pack 'archivedAt' $null
            createdAt = Get-DataProperty $pack 'createdAt' $null
            updatedAt = Get-DataProperty $pack 'updatedAt' $null
            settings = Get-ResolvedPackSettings $pack $Language
        }
    }
    return @($result)
}

function Get-PackTemplateRecord([string]$Language, [string]$TemplateId) {
    $id = ([string]$TemplateId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { $id = 'builtin-generic-department-pack' }
    if ($id -match '^builtin-(ecm|bod|dmm)$') {
        $category = $Matches[1]
        return [pscustomobject][ordered]@{
            templateId=$id; templateVersion=1; packId=(Get-BuiltinPackId $category); displayName=$category.ToUpperInvariant()
            description='Legacy compatibility template'; acceptedSourceTypes=@('excel','word','powerpoint','pdf')
            targets=@(
                [pscustomobject][ordered]@{targetId='main';displayName=$(if($Language -eq 'en'){'Main'}else{'本体'});required=$false},
                [pscustomobject][ordered]@{targetId='appendix';displayName=$(if($Language -eq 'en'){'Appendix'}else{'補足'});required=$false}
            )
            rules=[pscustomobject][ordered]@{newItemDestination='unassigned';retainManualOrder=$true;blockBuildWhenRequiredSourceIsStale=$false;blockBuildWhenRequiredSourceFailed=$false}
            output=[pscustomobject][ordered]@{fileNamePattern='{packName}_{targetName}_{yyyyMMdd}.pdf';pageNumbering='continuous';bookmarks='from-items'}
            workflowAvailable=$false; builtIn=$true
        }
    }
    $matches = @(Get-PackTemplateCatalog $Language | Where-Object { [string](Get-DataProperty $_ 'templateId' '') -eq $id } | Select-Object -First 1)
    if ($matches.Count -eq 0) { throw [ArgumentException]::new('指定された一式テンプレートが見つかりません。') }
    return $matches[0]
}

function Add-PackOutputStates($Structure, $Pack, $Template) {
    $outputs = Get-DataProperty $Structure 'outputs' ([ordered]@{})
    foreach ($target in @(Get-Array (Get-DataProperty $Template 'targets' @()))) {
        $targetId = [string](Get-DataProperty $target 'targetId' '')
        if ([string]::IsNullOrWhiteSpace($targetId) -or $targetId -eq 'unassigned') { continue }
        $key = "$([string]$Pack.packId)|$targetId"
        if (Test-ConfigHasKey $outputs $key) { continue }
        $empty = New-EmptyVolumeState
        Set-NoteProperty $outputs $key ([pscustomobject][ordered]@{
            packId = [string]$Pack.packId; targetId = $targetId; status = [string]$empty.status
            lastBuiltAt = $null; outputPdf = $null; builtFingerprint = ''; staleReasons = @(); message = ''; buildId = ''
        })
    }
    Set-NoteProperty $Structure 'outputs' $outputs
}

function New-DocumentPack([string]$Language, $Request) {
    return Update-StructureLocked $Language {
        param($structure)
        $template = Get-PackTemplateRecord $Language ([string](Get-DataProperty $Request 'templateId' 'builtin-generic-department-pack'))
        $displayName = Assert-PackDisplayNameAvailable $structure ([string](Get-DataProperty $Request 'displayName' ''))
        $now = New-NowIso
        $pack = [pscustomobject][ordered]@{
            packId = New-CustomPackId $structure
            templateId = [string]$template.templateId
            templateVersion = Get-IntDataProperty $template 'templateVersion' 1
            templateConfig = New-PackTemplateSnapshot $Language $template
            displayName = $displayName
            language = $Language
            createdAt = $now
            updatedAt = $now
            settings = [ordered]@{
                outputFileNamePattern = [string](Get-DataProperty (Get-DataProperty $template 'output' ([ordered]@{})) 'fileNamePattern' '{packName}_{targetName}_{yyyyMMdd}.pdf')
            }
            legacyCategory = ''
            archivedAt = $null
            review = [ordered]@{ status='draft'; submittedFingerprints=[ordered]@{}; submittedAt=''; submittedBy=''; approvedAt=''; approvedBy=''; note=''; events=@() }
        }
        $settingsPatch = Get-DataProperty $Request 'settings' $Request
        Set-NoteProperty $pack 'settings' (Get-UpdatedPackSettings $pack $Language $settingsPatch)
        Set-NoteProperty $structure 'packs' @((Get-Array (Get-DataProperty $structure 'packs' @())) + @($pack))
        Add-PackOutputStates $structure $pack $template
        return $pack
    }
}

function Copy-DocumentPack([string]$Language, [string]$PackId, $Request) {
    return Update-StructureLocked $Language {
        param($structure)
        $source = Get-PackRecord $structure $PackId
        $template = Get-PackEffectiveTemplate $Language $source
        $requestedName = [string](Get-DataProperty $Request 'displayName' '')
        $displayName = if ([string]::IsNullOrWhiteSpace($requestedName)) {
            Get-UniquePackDisplayName $structure ([string](Get-DataProperty $source 'displayName' 'ReportBinder')) $Language
        } else {
            Assert-PackDisplayNameAvailable $structure $requestedName
        }
        $now = New-NowIso
        $pack = [pscustomobject][ordered]@{
            packId = New-CustomPackId $structure
            templateId = [string]$template.templateId
            templateVersion = Get-IntDataProperty $source 'templateVersion' (Get-IntDataProperty $template 'templateVersion' 1)
            templateConfig = New-PackTemplateSnapshot $Language $template
            displayName = $displayName
            language = $Language
            createdAt = $now
            updatedAt = $now
            settings = Get-ResolvedPackSettings $source $Language
            legacyCategory = ''
            archivedAt = $null
            duplicatedFromPackId = [string](Get-DataProperty $source 'packId' '')
            review = [ordered]@{ status='draft'; submittedFingerprints=[ordered]@{}; submittedAt=''; submittedBy=''; approvedAt=''; approvedBy=''; note=''; events=@() }
        }
        Set-NoteProperty $structure 'packs' @((Get-Array (Get-DataProperty $structure 'packs' @())) + @($pack))
        Add-PackOutputStates $structure $pack $template
        return $pack
    }
}

function Set-DocumentPackArchived([string]$Language, [string]$PackId, [bool]$Archived) {
    return Update-StructureLocked $Language {
        param($structure)
        $pack = Get-PackRecord $structure $PackId
        if (-not [string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId ([string](Get-DataProperty $pack 'packId' ''))))) {
            throw [ArgumentException]::new('組み込み一式はアーカイブできません。')
        }
        if (-not $Archived) {
            [void](Assert-PackDisplayNameAvailable $structure ([string](Get-DataProperty $pack 'displayName' '')) ([string](Get-DataProperty $pack 'packId' '')))
        }
        Set-NoteProperty $pack 'archivedAt' $(if ($Archived) { New-NowIso } else { $null })
        Set-NoteProperty $pack 'updatedAt' (New-NowIso)
        return $pack
    }
}

function Remove-DocumentPack([string]$Language, [string]$PackId) {
    $id = ([string]$PackId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw [ArgumentException]::new('packId が必要です。') }
    return Update-StructureLocked $Language {
        param($structure)
        $pack = Get-PackRecord $structure $id
        if (-not [string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId ([string](Get-DataProperty $pack 'packId' ''))))) {
            throw [ArgumentException]::new('組み込み一式は削除できません。')
        }
        if (-not (Test-PackArchived $pack)) {
            throw [InvalidOperationException]::new('一式を削除する前にアーカイブしてください。')
        }

        $workbooks = @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { Test-WorkbookPack $_ $id })
        $workbookIds = @{}; foreach ($workbook in $workbooks) { $workbookIds[[string](Get-DataProperty $workbook 'workbookId' '')] = $true }
        $sources = @(Get-Array (Get-DataProperty $structure 'sources' @()) | Where-Object { [string](Get-DataProperty $_ 'packId' '') -eq $id })
        $sourceIds = @{}; foreach ($source in $sources) { $sourceIds[[string](Get-DataProperty $source 'sourceId' '')] = $true }
        foreach ($workbookId in @($workbookIds.Keys)) { $sourceIds[[string]$workbookId] = $true }
        $units = @(Get-Array (Get-DataProperty $structure 'units' @()) | Where-Object { $sourceIds.ContainsKey([string](Get-DataProperty $_ 'sourceId' '')) })
        $unitIds = @{}; foreach ($unit in $units) { $unitIds[[string](Get-DataProperty $unit 'unitId' '')] = $true }
        $pages = @(Get-Array (Get-DataProperty $structure 'pages' @()) | Where-Object { $workbookIds.ContainsKey([string](Get-DataProperty $_ 'workbookId' '')) })

        Set-NoteProperty $structure 'packs' @(Get-Array (Get-DataProperty $structure 'packs' @()) | Where-Object { [string](Get-DataProperty $_ 'packId' '') -ne $id })
        Set-NoteProperty $structure 'workbooks' @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { -not $workbookIds.ContainsKey([string](Get-DataProperty $_ 'workbookId' '')) })
        Set-NoteProperty $structure 'pages' @(Get-Array (Get-DataProperty $structure 'pages' @()) | Where-Object { -not $workbookIds.ContainsKey([string](Get-DataProperty $_ 'workbookId' '')) })
        Set-NoteProperty $structure 'sources' @(Get-Array (Get-DataProperty $structure 'sources' @()) | Where-Object { -not $sourceIds.ContainsKey([string](Get-DataProperty $_ 'sourceId' '')) })
        Set-NoteProperty $structure 'units' @(Get-Array (Get-DataProperty $structure 'units' @()) | Where-Object { -not $sourceIds.ContainsKey([string](Get-DataProperty $_ 'sourceId' '')) })
        Set-NoteProperty $structure 'items' @(Get-Array (Get-DataProperty $structure 'items' @()) | Where-Object {
            [string](Get-DataProperty $_ 'packId' '') -ne $id -and -not $sourceIds.ContainsKey([string](Get-DataProperty $_ 'sourceId' ''))
        })
        Set-NoteProperty $structure 'artifacts' @(Get-Array (Get-DataProperty $structure 'artifacts' @()) | Where-Object {
            [string](Get-DataProperty $_ 'packId' '') -ne $id -and
            -not $sourceIds.ContainsKey([string](Get-DataProperty $_ 'sourceId' '')) -and
            -not $unitIds.ContainsKey([string](Get-DataProperty $_ 'unitId' ''))
        })
        $outputs = Get-DataProperty $structure 'outputs' ([ordered]@{})
        $keptOutputs = [ordered]@{}
        foreach ($key in @(Get-ConfigKeyNames $outputs)) {
            $output = Get-DataProperty $outputs $key $null
            if ([string](Get-DataProperty $output 'packId' '') -eq $id -or $key.StartsWith("$id|", [StringComparison]::Ordinal)) { continue }
            $keptOutputs[$key] = $output
        }
        Set-NoteProperty $structure 'outputs' $keptOutputs
        return [ordered]@{
            packId=$id; displayName=[string](Get-DataProperty $pack 'displayName' '')
            removedWorkbookCount=$workbooks.Count; removedPageCount=$pages.Count
            sourceFilesDeleted=$false
        }
    }
}

function New-WorksheetUnitId([string]$SourceId, [string]$SheetName) {
    if ([string]::IsNullOrWhiteSpace($SourceId) -or [string]::IsNullOrWhiteSpace($SheetName)) { return '' }
    return "unit-$SourceId-$(Get-WorksheetStorageStem $SheetName)"
}

function New-ArtifactId([string]$SourceId, [string]$UnitId, [string]$SnapshotId, [string]$VersionId, [string]$RelativePdfPath) {
    $identity = @($SourceId,$UnitId,$SnapshotId,$VersionId,(([string]$RelativePdfPath -replace '\\','/').ToLowerInvariant())) -join '|'
    if ([string]::IsNullOrWhiteSpace($identity.Replace('|',''))) { return '' }
    return 'artifact-' + (Get-Sha256Text $identity).Substring(7, 20)
}

function New-EmptyStructure([string]$Language) {
    $volumes = [ordered]@{}
    foreach ($volume in @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })) {
        foreach ($category in @('ecm','bod','dmm')) {
            $volumes[(Get-VolumeStateKey $volume $category)] = New-EmptyVolumeState
        }
    }
    $structure = [pscustomobject][ordered]@{
        schemaVersion = 3
        language = $Language
        packs = @()
        sources = @()
        units = @()
        items = @()
        artifacts = @()
        outputs = [ordered]@{}
        migrationIssues = @()
        compatibility = [ordered]@{ schemaVersion = 2; mode = 'legacy-mirror'; updatedAt = New-NowIso }
        workbooks = @()
        pages = @()
        volumes = $volumes
        updatedAt = New-NowIso
    }
    Sync-StructureV3FromLegacy $structure $Language | Out-Null
    return $structure
}

function Ensure-Package($Paths, [string[]]$Languages = @()) {
    foreach ($key in @('submissionDir','dataDir','outputDir')) {
        if ([string]::IsNullOrWhiteSpace([string]$Paths.$key)) { throw "$key が未設定です。" }
    }
    foreach ($dir in @([string]$Paths.submissionDir, [string]$Paths.dataDir, [string]$Paths.outputDir)) {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    }
    $dataDir = [string]$Paths.dataDir
    foreach ($dir in @('common','common\audit','common\tmp','common\locks')) {
        $fullDir = Join-Path $dataDir $dir
        if (-not (Test-Path -LiteralPath $fullDir)) { New-Item -ItemType Directory -Path $fullDir -Force | Out-Null }
    }
    $pkgPath = Join-Path $dataDir 'package.json'
    if (-not (Test-Path -LiteralPath $pkgPath)) {
        Write-JsonFile $pkgPath ([ordered]@{ schemaVersion = 2; app = 'ReportBinder'; createdAt = New-NowIso })
    }
    $commonPathsPath = Join-Path $dataDir 'common\paths.json'
    $savedPaths = Read-JsonFile $commonPathsPath $null
    $pathsChanged = (
        $null -eq $savedPaths -or
        [string](Get-DataProperty $savedPaths 'submissionDir' '') -ne [string]$Paths.submissionDir -or
        [string](Get-DataProperty $savedPaths 'dataDir' '') -ne [string]$Paths.dataDir -or
        [string](Get-DataProperty $savedPaths 'outputDir' '') -ne [string]$Paths.outputDir
    )
    if ($pathsChanged) {
        Write-JsonFile $commonPathsPath ([ordered]@{
            submissionDir = [string]$Paths.submissionDir
            dataDir = [string]$Paths.dataDir
            outputDir = [string]$Paths.outputDir
            updatedAt = New-NowIso
        })
        # V5-P2: セットアップ変更直後に古いキャッシュを返さない。
        Reset-ConfigCaches
    }
    # Since document packs became the only management unit, Japanese and English
    # no longer own separate workspaces. Initializing the same workspace twice
    # with different language labels made the second pass reject the structure
    # created by the first pass. Select one language for the shared workspace and
    # validate the newly written document before committing the local config.
    $workspaceLanguages = @($Languages | Where-Object { $_ -in @('ja','en') } | Select-Object -Unique)
    if ($workspaceLanguages.Count -eq 0) { $workspaceLanguages = @((Get-EffectiveLanguage)) }
    $workspaceLanguage = [string]$workspaceLanguages[0]
    foreach ($lang in @($workspaceLanguage)) {
        foreach ($dir in @('', 'workbooks', 'pages', 'content-pdf', 'exports', 'state', 'locks', 'logs')) {
            $fullDir = Join-Path (Join-Path $dataDir $lang) $dir
            if (-not (Test-Path -LiteralPath $fullDir)) { New-Item -ItemType Directory -Path $fullDir -Force | Out-Null }
        }
        # Use the selected dataDir directly. On first run the local config is intentionally
        # saved only after package initialization succeeds, so Get-Paths is still empty here.
        Initialize-Or-MigrateStructure $lang ([string]$Paths.dataDir) | Out-Null
    }
    $verifiedStructure = Read-StructureUnlocked $workspaceLanguage ([string]$Paths.dataDir)
    if (-not (Test-StructureDocument $verifiedStructure $workspaceLanguage)) {
        throw '作業データを準備できませんでした。提出フォルダを選び直してください。原稿ファイルは変更していません。'
    }
}



function Get-WorksheetStorageStem([string]$SheetName) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes([string]$SheetName))
        $shortHash = -join @($bytes[0..7] | ForEach-Object { $_.ToString('x2') })
        return "sheet-$shortHash"
    } finally {
        if ($sha) { $sha.Dispose() }
    }
}

function New-WorksheetPageId([string]$WorkbookId, [string]$SheetName) {
    if ([string]::IsNullOrWhiteSpace($WorkbookId) -or [string]::IsNullOrWhiteSpace($SheetName)) { return '' }
    return "$WorkbookId-$(Get-WorksheetStorageStem $SheetName)"
}

function Get-PageSheetIndex($Page) {
    $value = Get-DataProperty $Page 'sheetIndex' $null
    $parsed = 0
    if ($null -ne $value -and [int]::TryParse([string]$value, [ref]$parsed) -and $parsed -gt 0) { return $parsed }
    return (Get-SheetOrderNumber ([string](Get-DataProperty $Page 'sheetName' '')))
}

function Resolve-PageId($Page) {
    if ($null -eq $Page) { return '' }
    $existing = [string]$Page.pageId
    if (-not [string]::IsNullOrWhiteSpace($existing)) { return $existing }
    $legacy = [string]$Page.id
    if (-not [string]::IsNullOrWhiteSpace($legacy)) { return $legacy }
    return (New-WorksheetPageId ([string]$Page.workbookId) ([string]$Page.sheetName))
}
function Set-NoteProperty($Object, [string]$Name, $Value) {
    if ($null -eq $Object) { return }
    if ($Object -is [System.Collections.IDictionary]) {
        $Object[$Name] = $Value
        return
    }
    if ($null -eq $Object.PSObject.Properties[$Name]) {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    } else {
        $Object.$Name = $Value
    }
}

function Get-DataProperty($Object, [string]$Name, $DefaultValue = $null) {
    if ($null -eq $Object) { return $DefaultValue }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $DefaultValue
    }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop) { return $DefaultValue }
    return $prop.Value
}


function Get-IntDataProperty($Object, [string]$Name, [int]$DefaultValue = 0) {
    try { return [int](Get-DataProperty $Object $Name $DefaultValue) } catch { return $DefaultValue }
}

function Update-LegacyCompatibilityViewFromV3($Structure) {
    if ($null -eq $Structure) { return }
    $sources = @(Get-Array (Get-DataProperty $Structure 'sources' @()))
    $items = @(Get-Array (Get-DataProperty $Structure 'items' @()))
    $workbooks = @(Get-Array (Get-DataProperty $Structure 'workbooks' @()))
    $pages = @(Get-Array (Get-DataProperty $Structure 'pages' @()))
    if ($workbooks.Count -eq 0 -and $sources.Count -gt 0) {
        $rebuilt = @()
        foreach ($source in $sources) {
            $legacy = Get-DataProperty $source 'legacyWorkbook' $null
            if ($null -eq $legacy) {
                $legacy = [pscustomobject][ordered]@{
                    workbookId = [string](Get-DataProperty $source 'sourceId' '')
                    relativePath = [string](Get-DataProperty $source 'relativePath' '')
                    fileName = [IO.Path]::GetFileName([string](Get-DataProperty $source 'relativePath' ''))
                    displayName = [string](Get-DataProperty $source 'displayName' '')
                    category = Get-CategoryFromBuiltinPackId ([string](Get-DataProperty $source 'packId' ''))
                    status = [string](Get-DataProperty $source 'status' 'new')
                    currentExcelHash = [string](Get-DataProperty $source 'currentSourceHash' '')
                    currentSnapshotId = [string](Get-DataProperty $source 'currentSnapshotId' '')
                    lastRenderedSnapshotId = [string](Get-DataProperty $source 'lastRenderedSnapshotId' '')
                    lastRenderedVersionId = [string](Get-DataProperty $source 'lastRenderedVersionId' '')
                    lastRenderedAt = Get-DataProperty $source 'lastRenderedAt' $null
                }
            }
            $rebuilt += $legacy
        }
        Set-NoteProperty $Structure 'workbooks' @($rebuilt)
    }
    if ($pages.Count -eq 0 -and $items.Count -gt 0) {
        $rebuiltPages = @()
        foreach ($item in $items) {
            $legacy = Get-DataProperty $item 'legacyPage' $null
            if ($null -eq $legacy) {
                $legacy = [pscustomobject][ordered]@{
                    pageId = [string](Get-DataProperty $item 'itemId' '')
                    workbookId = [string](Get-DataProperty $item 'sourceId' '')
                    sheetName = [string](Get-DataProperty $item 'sourceTitle' (Get-DataProperty $item 'title' ''))
                    title = [string](Get-DataProperty $item 'title' '')
                    volume = Get-LegacyVolumeFromTargetId ([string](Get-DataProperty $Structure 'language' 'ja')) ([string](Get-DataProperty $item 'targetId' 'unassigned'))
                    order = Get-DataProperty $item 'order' 0
                    enabled = [bool](Get-DataProperty $item 'enabled' $true)
                    numberingMode = [string](Get-DataProperty $item 'numberingMode' 'visible')
                    numberingManual = [bool](Get-DataProperty $item 'numberingManual' $false)
                    pageRange = Get-DataProperty $item 'pageRange' $null
                    contentPdf = [string](Get-DataProperty $item 'legacyContentPdf' '')
                }
            }
            $rebuiltPages += $legacy
        }
        Set-NoteProperty $Structure 'pages' @($rebuiltPages)
    }
    $volumes = Get-DataProperty $Structure 'volumes' $null
    if ($null -eq $volumes -or @(Get-ConfigKeyNames $volumes).Count -eq 0) {
        $rebuiltVolumes = [ordered]@{}
        $outputs = Get-DataProperty $Structure 'outputs' ([ordered]@{})
        foreach ($key in @(Get-ConfigKeyNames $outputs)) {
            $output = Get-DataProperty $outputs $key $null
            if ($null -eq $output) { continue }
            $packId = [string](Get-DataProperty $output 'packId' '')
            $targetId = [string](Get-DataProperty $output 'targetId' '')
            $category = Get-CategoryFromBuiltinPackId $packId
            if ([string]::IsNullOrWhiteSpace($category)) { continue }
            $legacyVolume = Get-LegacyVolumeFromTargetId ([string](Get-DataProperty $Structure 'language' 'ja')) $targetId
            if ($legacyVolume -eq 'none') { continue }
            $legacyState = Get-DataProperty $output 'legacyVolume' $null
            if ($null -eq $legacyState) {
                $legacyState = [pscustomobject][ordered]@{
                    status = [string](Get-DataProperty $output 'status' 'not-built')
                    lastBuiltAt = Get-DataProperty $output 'lastBuiltAt' $null
                    outputPdf = Get-DataProperty $output 'outputPdf' $null
                    builtFingerprint = [string](Get-DataProperty $output 'builtFingerprint' '')
                    staleReasons = @(Get-Array (Get-DataProperty $output 'staleReasons' @()))
                    message = [string](Get-DataProperty $output 'message' '')
                }
            }
            $rebuiltVolumes[(Get-VolumeStateKey $legacyVolume $category)] = $legacyState
        }
        Set-NoteProperty $Structure 'volumes' $rebuiltVolumes
    }
}

function Sync-StructureV3FromLegacy($Structure, [string]$Language = '') {
    if ($null -eq $Structure) { return $Structure }
    if ([string]::IsNullOrWhiteSpace($Language)) { $Language = [string](Get-DataProperty $Structure 'language' 'ja') }
    if ($Language -notin @('ja','en')) { $Language = 'ja' }
    $oldPackMap = @{}; foreach ($x in @(Get-Array (Get-DataProperty $Structure 'packs' @()))) { $oldPackMap[[string](Get-DataProperty $x 'packId' '')] = $x }
    $oldSourceMap = @{}; foreach ($x in @(Get-Array (Get-DataProperty $Structure 'sources' @()))) { $oldSourceMap[[string](Get-DataProperty $x 'sourceId' '')] = $x }
    $oldUnitMap = @{}; foreach ($x in @(Get-Array (Get-DataProperty $Structure 'units' @()))) { $oldUnitMap[[string](Get-DataProperty $x 'unitId' '')] = $x }
    $oldItemMap = @{}; foreach ($x in @(Get-Array (Get-DataProperty $Structure 'items' @()))) { $oldItemMap[[string](Get-DataProperty $x 'itemId' '')] = $x }
    $oldArtifactMap = @{}; foreach ($x in @(Get-Array (Get-DataProperty $Structure 'artifacts' @()))) { $oldArtifactMap[[string](Get-DataProperty $x 'artifactId' '')] = $x }
    $oldOutputs = Get-DataProperty $Structure 'outputs' ([ordered]@{})
    $packs = @()
    foreach ($category in @('ecm','bod','dmm')) {
        $packId = Get-BuiltinPackId $category
        $existing = if ($oldPackMap.ContainsKey($packId)) { $oldPackMap[$packId] } else { $null }
        $packs += New-BuiltinPack $category $Language $existing
    }
    foreach ($oldPack in @(Get-Array (Get-DataProperty $Structure 'packs' @()))) {
        if ([string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId ([string](Get-DataProperty $oldPack 'packId' ''))))) { $packs += $oldPack }
    }
    $sources = @(); $sourceById = @{}
    foreach ($wb in @(Get-Array (Get-DataProperty $Structure 'workbooks' @()))) {
        $sourceId = [string](Get-DataProperty $wb 'workbookId' '')
        if ([string]::IsNullOrWhiteSpace($sourceId)) { continue }
        $category = Normalize-WorkbookCategory ([string](Get-DataProperty $wb 'category' '')) ([string](Get-DataProperty $wb 'fileName' ''))
        if ([string]::IsNullOrWhiteSpace($category)) { $category = 'ecm' }
        $old = if ($oldSourceMap.ContainsKey($sourceId)) { $oldSourceMap[$sourceId] } else { $null }
        $sourceType = ([string](Get-DataProperty $wb 'sourceType' (Get-DataProperty $old 'sourceType' ''))).Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($sourceType)) {
            $extension = [IO.Path]::GetExtension([string](Get-DataProperty $wb 'relativePath' '')).ToLowerInvariant()
            $sourceType = $(if ($extension -eq '.pdf') { 'pdf' } elseif ($extension -eq '.docx') { 'word' } elseif ($extension -eq '.pptx') { 'powerpoint' } else { 'excel' })
        }
        $defaultAdapterId = $(if ($sourceType -eq 'pdf') { 'pdfbox-import-v1' } elseif ($sourceType -eq 'word') { 'word-com-v1' } elseif ($sourceType -eq 'powerpoint') { 'powerpoint-com-v1' } else { 'excel-com-v1' })
        $savedPackId = [string](Get-DataProperty $wb 'packId' (Get-DataProperty $old 'packId' ''))
        if ([string]::IsNullOrWhiteSpace($savedPackId) -or -not $oldPackMap.ContainsKey($savedPackId)) { $savedPackId = Get-BuiltinPackId $category }
        Set-NoteProperty $wb 'packId' $savedPackId
        Set-NoteProperty $wb 'sourceType' $sourceType
        Set-NoteProperty $wb 'adapterId' ([string](Get-DataProperty $old 'adapterId' (Get-DataProperty $wb 'adapterId' $defaultAdapterId)))
        Set-NoteProperty $wb 'ownerDepartment' ([string](Get-DataProperty $old 'ownerDepartment' (Get-DataProperty $wb 'ownerDepartment' '')))
        Set-NoteProperty $wb 'required' ([bool](Get-DataProperty $old 'required' (Get-DataProperty $wb 'required' $true)))
        Set-NoteProperty $wb 'defaultTargetId' ([string](Get-DataProperty $old 'defaultTargetId' (Get-DataProperty $wb 'defaultTargetId' 'unassigned')))
        Set-NoteProperty $wb 'requirementId' ([string](Get-DataProperty $old 'requirementId' (Get-DataProperty $wb 'requirementId' '')))
        if ($sourceType -eq 'excel') {
            Set-NoteProperty $wb 'excelSheetSelectionMode' (Normalize-ExcelSheetSelection ([string](Get-DataProperty $wb 'excelSheetSelectionMode' (Get-DataProperty $old 'excelSheetSelectionMode' 'all-visible'))))
            Set-NoteProperty $wb 'lastRenderedSheetSelectionMode' (Normalize-ExcelSheetSelection ([string](Get-DataProperty $wb 'lastRenderedSheetSelectionMode' (Get-DataProperty $old 'lastRenderedSheetSelectionMode' 'all-visible'))))
        }
        $source = [pscustomobject][ordered]@{
            sourceId = $sourceId; packId = $savedPackId
            relativePath = [string](Get-DataProperty $wb 'relativePath' (Get-DataProperty $wb 'fileName' ''))
            sourceType = $sourceType; adapterId = [string](Get-DataProperty $wb 'adapterId' $defaultAdapterId)
            displayName = [string](Get-DataProperty $wb 'displayName' (Get-DataProperty $wb 'fileName' $sourceId))
            ownerDepartment = [string](Get-DataProperty $wb 'ownerDepartment' '')
            required = [bool](Get-DataProperty $wb 'required' $true)
            defaultTargetId = [string](Get-DataProperty $wb 'defaultTargetId' 'unassigned')
            requirementId = [string](Get-DataProperty $wb 'requirementId' '')
            excelSheetSelectionMode = $(if ($sourceType -eq 'excel') { Normalize-ExcelSheetSelection ([string](Get-DataProperty $wb 'excelSheetSelectionMode' 'all-visible')) } else { 'all-visible' })
            status = [string](Get-DataProperty $wb 'status' 'new')
            currentSourceHash = [string](Get-DataProperty $wb 'currentExcelHash' '')
            currentSnapshotId = [string](Get-DataProperty $wb 'currentSnapshotId' '')
            lastRenderedSnapshotId = [string](Get-DataProperty $wb 'lastRenderedSnapshotId' '')
            lastRenderedVersionId = [string](Get-DataProperty $wb 'lastRenderedVersionId' '')
            lastRenderedAt = Get-DataProperty $wb 'lastRenderedAt' $null
            sourceMetadata = [ordered]@{ size = Get-DataProperty $wb 'currentExcelSize' 0; modifiedAtUtcTicks = [string](Get-DataProperty $wb 'currentExcelLastWriteUtcTicks' '') }
            legacyWorkbook = $wb
        }
        $sources += $source; $sourceById[$sourceId] = $source
    }
    $units = @(); $unitAdded = @{}; $items = @(); $artifactMap = [ordered]@{}
    foreach ($page in @(Get-Array (Get-DataProperty $Structure 'pages' @()))) {
        $sourceId = [string](Get-DataProperty $page 'workbookId' ''); $sheetName = [string](Get-DataProperty $page 'sheetName' ''); $itemId = Resolve-PageId $page
        if ([string]::IsNullOrWhiteSpace($sourceId) -or [string]::IsNullOrWhiteSpace($itemId)) { continue }
        $unitId = New-WorksheetUnitId $sourceId $sheetName
        $source = if ($sourceById.ContainsKey($sourceId)) { $sourceById[$sourceId] } else { $null }
        $packId = if ($null -ne $source) { [string]$source.packId } else { Get-BuiltinPackId 'ecm' }
        $sourceType = if ($null -ne $source) { [string](Get-DataProperty $source 'sourceType' 'excel') } else { 'excel' }
        $unitKind = if ($sourceType -eq 'pdf') { 'pdf-page' } elseif ($sourceType -eq 'word') { 'document-page' } elseif ($sourceType -eq 'powerpoint') { 'slide' } else { 'worksheet' }
        $adapterId = if ($sourceType -eq 'pdf') { 'pdfbox-import-v1' } elseif ($sourceType -eq 'word') { 'word-com-v1' } elseif ($sourceType -eq 'powerpoint') { 'powerpoint-com-v1' } else { 'excel-com-v1' }
        $versionId = if ($null -ne $source) { [string]$source.lastRenderedVersionId } else { '' }
        $snapshotId = if ($null -ne $source) { [string]$source.lastRenderedSnapshotId } else { '' }
        $relativePdf = [string](Get-DataProperty $page 'contentPdf' ''); $artifactId = ''
        if (-not [string]::IsNullOrWhiteSpace($relativePdf)) {
            $artifactId = New-ArtifactId $sourceId $unitId $snapshotId $versionId $relativePdf
            if (-not $artifactMap.Contains($artifactId)) {
                $oldArtifact = if ($oldArtifactMap.ContainsKey($artifactId)) { $oldArtifactMap[$artifactId] } else { $null }
                $artifactMap[$artifactId] = [pscustomobject][ordered]@{
                    artifactId = $artifactId; sourceId = $sourceId; unitId = $unitId; snapshotId = $snapshotId
                    adapterId = $adapterId; adapterVersion = 1; relativePdfPath = $relativePdf
                    sha256 = [string](Get-DataProperty $oldArtifact 'sha256' ''); pageCount = Get-IntDataProperty $oldArtifact 'pageCount' 0
                    createdAt = Get-DataProperty $oldArtifact 'createdAt' (Get-DataProperty $source 'lastRenderedAt' $null)
                }
            }
        }
        $oldUnit = if ($oldUnitMap.ContainsKey($unitId)) { $oldUnitMap[$unitId] } else { $null }
        $artifactPageCount = if ($artifactMap.Contains($artifactId)) { Get-IntDataProperty $artifactMap[$artifactId] 'pageCount' 0 } else { 0 }
        if (-not $unitAdded.ContainsKey($unitId)) {
            $units += [pscustomobject][ordered]@{
                unitId = $unitId; sourceId = $sourceId; unitKind = $unitKind; sourceKey = $(if ($sourceType -eq 'powerpoint') { "slide:$([int](Get-DataProperty $page 'sheetIndex' 0))" } elseif ($sourceType -in @('pdf','word')) { "page:$([int](Get-DataProperty $page 'sheetIndex' 0))" } else { "worksheet:$sheetName" })
                title = [string](Get-DataProperty $page 'title' $sheetName); sourceIndex = Get-IntDataProperty $page 'sheetIndex' 0
                status = [string](Get-DataProperty $page 'status' 'not-rendered'); renderVersionId = $versionId; artifactId = $artifactId
                physicalPageCount = Get-IntDataProperty $oldUnit 'physicalPageCount' $artifactPageCount
            }
            $unitAdded[$unitId] = $true
        }
        $oldItem = if ($oldItemMap.ContainsKey($itemId)) { $oldItemMap[$itemId] } else { $null }
        $item = [pscustomobject][ordered]@{
            itemId = $itemId; packId = $packId; sourceId = $sourceId; unitId = $unitId
            targetId = Get-TargetIdFromLegacyVolume ([string](Get-DataProperty $page 'volume' 'none'))
            sectionId = [string](Get-DataProperty $oldItem 'sectionId' 'body'); order = Get-DataProperty $page 'order' 0
            enabled = [bool](Get-DataProperty $page 'enabled' $true); title = [string](Get-DataProperty $page 'title' $sheetName); sourceTitle = $sheetName
            numberingMode = [string](Get-DataProperty $page 'numberingMode' 'visible'); numberingManual = [bool](Get-DataProperty $page 'numberingManual' $false)
            artifactId = $artifactId; pageRange = Get-DataProperty $page 'pageRange' (Get-DataProperty $oldItem 'pageRange' $null); legacyContentPdf = $relativePdf; legacyPage = $page
        }
        $items += $item
    }
    $outputs = [ordered]@{}; $legacyVolumes = Get-DataProperty $Structure 'volumes' ([ordered]@{})
    foreach ($category in @('ecm','bod','dmm')) {
        $packId = Get-BuiltinPackId $category
        foreach ($legacyVolume in @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })) {
            $targetId = Get-TargetIdFromLegacyVolume $legacyVolume; $outputKey = "$packId|$targetId"
            $legacyState = Get-DataProperty $legacyVolumes (Get-VolumeStateKey $legacyVolume $category) $null
            if ($null -eq $legacyState) { $legacyState = New-EmptyVolumeState }
            $oldOutput = Get-DataProperty $oldOutputs $outputKey $null
            $outputs[$outputKey] = [pscustomobject][ordered]@{
                packId = $packId; targetId = $targetId; status = [string](Get-DataProperty $legacyState 'status' 'not-built')
                lastBuiltAt = Get-DataProperty $legacyState 'lastBuiltAt' $null; outputPdf = Get-DataProperty $legacyState 'outputPdf' $null
                builtFingerprint = [string](Get-DataProperty $legacyState 'builtFingerprint' ''); staleReasons = @(Get-Array (Get-DataProperty $legacyState 'staleReasons' @()))
                message = [string](Get-DataProperty $legacyState 'message' ''); buildId = [string](Get-DataProperty $oldOutput 'buildId' ''); legacyVolume = $legacyState
            }
        }
    }
    foreach ($oldKey in @(Get-ConfigKeyNames $oldOutputs)) { if (-not $outputs.Contains($oldKey)) { $outputs[$oldKey] = Get-DataProperty $oldOutputs $oldKey $null } }
    Set-NoteProperty $Structure 'schemaVersion' 3; Set-NoteProperty $Structure 'language' $Language
    Set-NoteProperty $Structure 'packs' @($packs); Set-NoteProperty $Structure 'sources' @($sources); Set-NoteProperty $Structure 'units' @($units)
    Set-NoteProperty $Structure 'items' @($items); Set-NoteProperty $Structure 'artifacts' @($artifactMap.Values); Set-NoteProperty $Structure 'outputs' $outputs
    if (-not (Test-ConfigHasKey $Structure 'migrationIssues')) { Set-NoteProperty $Structure 'migrationIssues' @() }
    Set-NoteProperty $Structure 'compatibility' ([ordered]@{ schemaVersion = 2; mode = 'legacy-mirror'; updatedAt = New-NowIso })
    return $Structure
}

function ConvertTo-StructureV3($Structure, [string]$Language, [string]$Workspace = '') {
    $Structure = Normalize-StructureCollections $Structure; [void](Repair-StructurePages $Structure); $issues = @()
    foreach ($wb in @(Get-Array $Structure.workbooks)) {
        $category = Normalize-WorkbookCategory ([string](Get-DataProperty $wb 'category' '')) ([string](Get-DataProperty $wb 'fileName' ''))
        if ([string]::IsNullOrWhiteSpace($category)) {
            $issues += [pscustomobject][ordered]@{ code = 'category-unresolved'; sourceId = [string](Get-DataProperty $wb 'workbookId' ''); relativePath = [string](Get-DataProperty $wb 'relativePath' ''); message = '既存カテゴリを判定できませんでした。ECMとして互換移行した後、一式を確認してください。' }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($Workspace)) {
        foreach ($page in @(Get-Array $Structure.pages)) {
            $relativePdf = [string](Get-DataProperty $page 'contentPdf' ''); if ([string]::IsNullOrWhiteSpace($relativePdf)) { continue }
            try {
                $root = [IO.Path]::GetFullPath($Workspace); if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
                $full = [IO.Path]::GetFullPath((Join-Path $Workspace $relativePdf))
                if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $full)) { throw 'missing' }
            } catch { $issues += [pscustomobject][ordered]@{ code = 'content-pdf-missing'; itemId = Resolve-PageId $page; relativePath = $relativePdf; message = '移行時に変換PDFを確認できませんでした。再度PDFを作成してください。' } }
        }
    }
    Set-NoteProperty $Structure 'migrationIssues' @($issues); Sync-StructureV3FromLegacy $Structure $Language | Out-Null
    return $Structure
}

function ConvertTo-V4StructureCompatibilityView($Structure) {
    if ($null -eq $Structure) { return $null }
    Update-LegacyCompatibilityViewFromV3 $Structure
    return [pscustomobject][ordered]@{
        schemaVersion = 2; language = [string](Get-DataProperty $Structure 'language' '')
        workbooks = @(Get-Array (Get-DataProperty $Structure 'workbooks' @())); pages = @(Get-Array (Get-DataProperty $Structure 'pages' @()))
        volumes = Get-DataProperty $Structure 'volumes' ([ordered]@{}); updatedAt = Get-DataProperty $Structure 'updatedAt' $null
    }
}

function Get-RequiredSourceRenderProfileVersion($Workbook) {
    $sourceType = [string](Get-DataProperty $Workbook 'sourceType' 'excel')
    if ($sourceType -eq 'pdf') { return $Script:PdfImportProfileVersion }
    if ($sourceType -eq 'word') { return $Script:WordRenderProfileVersion }
    if ($sourceType -eq 'powerpoint') { return $Script:PowerPointRenderProfileVersion }
    return $Script:ExcelPrintProfileVersion
}

function Get-SourceUpdatedStatus($Workbook) {
    if ([string](Get-DataProperty $Workbook 'sourceType' 'excel') -ne 'excel') { return 'source-updated' }
    return 'excel-updated'
}

function Test-WorkbookRenderIsCurrent($Workbook) {
    if ($null -eq $Workbook) { return $false }
    $status = [string](Get-DataProperty $Workbook 'status' '')
    if ($status -in @('new','excel-updated','source-updated','missing','render-error','rendering','stale','not-rendered')) { return $false }
    $lastRenderedHash = [string](Get-DataProperty $Workbook 'lastRenderedExcelHash' '')
    if ([string]::IsNullOrWhiteSpace($lastRenderedHash)) { return $false }
    $currentHash = [string](Get-DataProperty $Workbook 'currentExcelHash' '')
    if ((-not [string]::IsNullOrWhiteSpace($currentHash)) -and $currentHash -ne $lastRenderedHash) { return $false }
    $profileVersion = Get-IntDataProperty $Workbook 'renderProfileVersion' 0
    if ($profileVersion -lt (Get-RequiredSourceRenderProfileVersion $Workbook)) { return $false }
    if ([string](Get-DataProperty $Workbook 'sourceType' 'excel') -eq 'excel') {
        $wantedSheetSelection = Normalize-ExcelSheetSelection ([string](Get-DataProperty $Workbook 'excelSheetSelectionMode' 'all-visible'))
        $renderedSheetSelection = Normalize-ExcelSheetSelection ([string](Get-DataProperty $Workbook 'lastRenderedSheetSelectionMode' 'all-visible'))
        if ($wantedSheetSelection -ne $renderedSheetSelection) { return $false }
    }
    return $true
}

function Mark-WorkbookContentStaleForProfile($Structure, $Workbook) {
    $changed = $false
    if ($null -eq $Workbook) { return $changed }
    $lastRenderedHash = [string](Get-DataProperty $Workbook 'lastRenderedExcelHash' '')
    if ([string]::IsNullOrWhiteSpace($lastRenderedHash)) { return $changed }
    $profileVersion = Get-IntDataProperty $Workbook 'renderProfileVersion' 0
    if ($profileVersion -ge (Get-RequiredSourceRenderProfileVersion $Workbook)) { return $changed }
    $updatedStatus = Get-SourceUpdatedStatus $Workbook
    if ([string](Get-DataProperty $Workbook 'status' '') -ne $updatedStatus) {
        Set-NoteProperty $Workbook 'status' $updatedStatus
        $changed = $true
    }
    foreach ($p in @(Get-Array $Structure.pages | Where-Object { [string]$_.workbookId -eq [string]$Workbook.workbookId })) {
        if (-not [string]::IsNullOrWhiteSpace([string]$p.contentPdf) -and [string]$p.status -ne 'stale') {
            Set-NoteProperty $p 'status' 'stale'
            Set-NoteProperty $p 'updatedAt' (New-NowIso)
            $changed = $true
        }
    }
    return $changed
}


function Repair-StructurePages($Structure) {
    $changed = $false
    if ($null -eq $Structure.workbooks) { Set-NoteProperty $Structure 'workbooks' @(); $changed = $true }
    if ($null -eq $Structure.pages) { Set-NoteProperty $Structure 'pages' @(); $changed = $true }
    foreach ($wb in @(Get-Array $Structure.workbooks)) {
        # structure.json created by older builds may not have these properties.
        # Add them explicitly instead of assigning to a missing PSCustomObject property.
        $cat = ''
        try { $cat = Normalize-WorkbookCategory ([string]$wb.category) ([string]$wb.fileName) } catch { $cat = '' }
        if ([string]::IsNullOrWhiteSpace($cat)) {
            $name = [string]$wb.fileName
            if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$wb.displayName }
            if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$wb.relativePath }
            $cat = Normalize-WorkbookCategory '' $name
        }
        if ($null -eq $wb.PSObject.Properties['category'] -or [string]$wb.category -ne $cat) { Set-NoteProperty $wb 'category' $cat; $changed = $true }
        foreach ($pair in @(
            @{Name='lastError'; Value=''},
            @{Name='lastErrorUser'; Value=''},
            @{Name='lastErrorAt'; Value=$null},
            @{Name='lastRenderAttemptHash'; Value=''},
            @{Name='lastRenderLog'; Value=''},
            @{Name='renderProfileVersion'; Value=0},
            @{Name='lastRenderedSheets'; Value=@()},
            @{Name='lastRenderedSheetFingerprint'; Value=''},
            @{Name='currentExcelLastWriteUtcTicks'; Value=''},
            @{Name='warnings'; Value=@()},
            @{Name='excelSheetSelectionMode'; Value='all-visible'},
            @{Name='lastRenderedSheetSelectionMode'; Value='all-visible'}
        )) {
            if ($null -eq $wb.PSObject.Properties[$pair.Name]) { Set-NoteProperty $wb $pair.Name $pair.Value; $changed = $true }
        }
        if ([string](Get-DataProperty $wb 'sourceType' 'excel') -eq 'excel') {
            foreach ($selectionProperty in @('excelSheetSelectionMode','lastRenderedSheetSelectionMode')) {
                try { $normalizedSelection = Normalize-ExcelSheetSelection ([string](Get-DataProperty $wb $selectionProperty 'all-visible')) }
                catch { $normalizedSelection = 'all-visible' }
                if ([string](Get-DataProperty $wb $selectionProperty 'all-visible') -ne $normalizedSelection) { Set-NoteProperty $wb $selectionProperty $normalizedSelection; $changed = $true }
            }
        }
        if (Mark-WorkbookContentStaleForProfile $Structure $wb) { $changed = $true }
    }
    $used = @{}
    foreach ($p in @(Get-Array $Structure.pages)) {
        $pageKey = Resolve-PageId $p
        if (-not [string]::IsNullOrWhiteSpace($pageKey)) {
            $base = $pageKey
            $n = 2
            while ($used.ContainsKey($pageKey)) {
                $pageKey = "$base-$n"
                $n++
            }
            $used[$pageKey] = $true
            if ([string]::IsNullOrWhiteSpace([string]$p.pageId) -or [string]$p.pageId -ne $pageKey) {
                Set-NoteProperty $p 'pageId' $pageKey
                $changed = $true
            }
        }
        if ($null -eq $p.PSObject.Properties['numberingManual']) { Set-NoteProperty $p 'numberingManual' $false; $changed = $true }
        if ($null -eq $p.PSObject.Properties['numberingDefault']) { Set-NoteProperty $p 'numberingDefault' 'first-page-none'; $changed = $true }
        if ($null -eq $p.PSObject.Properties['enabled']) { Set-NoteProperty $p 'enabled' $true; $changed = $true }
        if ($null -eq $p.PSObject.Properties['sheetSelectionExcluded']) { Set-NoteProperty $p 'sheetSelectionExcluded' $false; $changed = $true }
    }
    return $changed
}

function Test-UniqueStructureRecordIds($Records, [string]$IdProperty) {
    $seen = @{}
    foreach ($record in @(Get-Array $Records)) {
        $id = [string](Get-DataProperty $record $IdProperty '')
        if ([string]::IsNullOrWhiteSpace($id) -or $seen.ContainsKey($id)) { return $false }
        $seen[$id] = $true
    }
    return $true
}

function Test-DirectPdfRelativePath([string]$RelativePath) {
    Test-RelativePath $RelativePath | Out-Null
    if ($RelativePath -match '[\\/]') { throw '提出フォルダ直下のPDFだけ登録できます。子フォルダ内のファイルは対象外です。' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -ne '.pdf') { throw '拡張子が .pdf のPDFだけ登録できます。' }
    return $true
}

function Test-StructureV3References($Structure) {
    $packs = @(Get-Array (Get-DataProperty $Structure 'packs' @()))
    $sources = @(Get-Array (Get-DataProperty $Structure 'sources' @()))
    $units = @(Get-Array (Get-DataProperty $Structure 'units' @()))
    $items = @(Get-Array (Get-DataProperty $Structure 'items' @()))
    $artifacts = @(Get-Array (Get-DataProperty $Structure 'artifacts' @()))
    if (-not (Test-UniqueStructureRecordIds $packs 'packId') -or
        -not (Test-UniqueStructureRecordIds $sources 'sourceId') -or
        -not (Test-UniqueStructureRecordIds $units 'unitId') -or
        -not (Test-UniqueStructureRecordIds $items 'itemId') -or
        -not (Test-UniqueStructureRecordIds $artifacts 'artifactId')) { return $false }
    $packIds = @{}; foreach ($x in $packs) { $packIds[[string]$x.packId] = $true }
    $sourceIds = @{}; foreach ($x in $sources) { $sourceIds[[string]$x.sourceId] = $true; if (-not $packIds.ContainsKey([string]$x.packId)) { return $false } }
    $unitIds = @{}; foreach ($x in $units) { $unitIds[[string]$x.unitId] = $true; if (-not $sourceIds.ContainsKey([string]$x.sourceId)) { return $false } }
    $artifactIds = @{}; foreach ($x in $artifacts) { $artifactIds[[string]$x.artifactId] = $true; if (-not $sourceIds.ContainsKey([string]$x.sourceId) -or -not $unitIds.ContainsKey([string]$x.unitId)) { return $false } }
    foreach ($x in $items) {
        if (-not $packIds.ContainsKey([string]$x.packId) -or -not $sourceIds.ContainsKey([string]$x.sourceId) -or -not $unitIds.ContainsKey([string]$x.unitId)) { return $false }
        $artifactId = [string](Get-DataProperty $x 'artifactId' '')
        if (-not [string]::IsNullOrWhiteSpace($artifactId) -and -not $artifactIds.ContainsKey($artifactId)) { return $false }
    }
    $outputs = Get-DataProperty $Structure 'outputs' ([ordered]@{})
    foreach ($key in @(Get-ConfigKeyNames $outputs)) {
        $output = Get-DataProperty $outputs $key $null
        if ($null -eq $output -or -not $packIds.ContainsKey([string](Get-DataProperty $output 'packId' '')) -or [string]::IsNullOrWhiteSpace([string](Get-DataProperty $output 'targetId' ''))) { return $false }
    }
    return $true
}

function Test-StructureDocument($Structure, [string]$Language) {
    if ($null -eq $Structure) { return $false }
    try {
        $version = [int](Get-DataProperty $Structure 'schemaVersion' 0)
        if ($version -lt 1 -or $version -gt 3) { return $false }
        $storedLanguage = [string](Get-DataProperty $Structure 'language' '')
        # schema v3 is a single workspace shared by all pack languages. The
        # structure-level language is retained only for legacy compatibility.
        if ($version -le 2 -and -not [string]::IsNullOrWhiteSpace($storedLanguage) -and $storedLanguage -ne $Language) { return $false }
        if ($version -le 2) {
            if (-not (Test-ConfigHasKey $Structure 'workbooks') -or -not (Test-ConfigHasKey $Structure 'pages')) { return $false }
        } else {
            foreach ($name in @('packs','sources','units','items','artifacts','outputs')) {
                if (-not (Test-ConfigHasKey $Structure $name)) { return $false }
            }
            if (-not (Test-StructureV3References $Structure)) { return $false }
        }
        return $true
    } catch {
        return $false
    }
}

function Read-StructureUnlocked([string]$Language, [string]$DataDir = '') {
    $workspace = Get-WorkspacePath $Language $DataDir
    $path = Join-Path $workspace 'structure.json'
    if (-not (Test-Path -LiteralPath $path)) { return New-EmptyStructure $Language }
    $backupPath = Join-Path $workspace 'structure.json.last-good'
    $structure = $null
    try {
        $structure = Read-JsonFile $path $null
        if (-not (Test-StructureDocument $structure $Language)) { throw 'structure.json の内容が不完全です。' }
    } catch {
        # Never reinterpret an existing but temporarily unreadable shared-file as
        # a brand-new empty workspace. That old behavior allowed the next mutation
        # to unregister every workbook. A verified last-good copy is safe to use;
        # without one, fail closed and leave the original file untouched.
        $primaryError = $_.Exception.Message
        $structure = $null
        try {
            if (Test-Path -LiteralPath $backupPath) {
                $candidate = Read-JsonFile $backupPath $null
                if (Test-StructureDocument $candidate $Language) { $structure = $candidate }
            }
        } catch { $structure = $null }
        if ($null -eq $structure) {
            throw "登録情報を安全に読み込めませんでした。structure.json は上書きしていません。管理フォルダの接続を確認して再起動してください。詳細: $primaryError"
        }
    }
    return Normalize-StructureCollections $structure
}

function Write-StructureUnlocked([string]$Language, $Structure, [string]$DataDir = '') {
    $Structure = Normalize-StructureCollections $Structure
    if ([int](Get-DataProperty $Structure 'schemaVersion' 0) -eq 3) { Sync-StructureV3FromLegacy $Structure $Language | Out-Null }
    if (-not (Test-StructureDocument $Structure $Language)) { throw '不完全な登録情報の保存を拒否しました。' }
    Set-NoteProperty $Structure 'updatedAt' (New-NowIso)
    $workspace = Get-WorkspacePath $Language $DataDir
    $path = Join-Path $workspace 'structure.json'
    $backupPath = Join-Path $workspace 'structure.json.last-good'
    if (Test-Path -LiteralPath $path) {
        try {
            $existing = Read-JsonFile $path $null
            if (Test-StructureDocument $existing $Language) {
                Write-JsonFile $backupPath $existing
            }
        } catch {
            # Preserve the previous last-good copy when the primary cannot be read.
            # The caller's Structure may itself have been recovered from that copy.
        }
    }
    Write-JsonFile $path $Structure
    $saved = Read-JsonFile $path $null
    if (-not (Test-StructureDocument $saved $Language)) {
        throw '登録情報を保存後に検証できませんでした。直前のバックアップを保持しています。'
    }
}

function Get-Structure([string]$Language) {
    # Read-only API path. Repair and schema migration are performed only by Initialize-Or-MigrateStructure.
    # state取得直後の比較画面で同じ共有JSONを再読込しない。外部PCの更新は
    # LastWriteTimeUtc/Lengthの変化で検出し、書込経路では明示的に破棄する。
    $workspace = Get-WorkspacePath $Language
    $path = Join-Path $workspace 'structure.json'
    try { $file = Get-Item -LiteralPath $path -ErrorAction Stop } catch { return (Read-StructureUnlocked $Language) }
    $cacheKey = (([IO.Path]::GetFullPath($path)) + '|' + $Language).ToLowerInvariant()
    $stamp = ([string]$file.Length) + '|' + ([string]$file.LastWriteTimeUtc.Ticks)
    $cached = $Script:StructureReadCache[$cacheKey]
    if ($null -ne $cached -and [string]$cached.stamp -eq $stamp) { return $cached.structure }
    $structure = Read-StructureUnlocked $Language
    if (-not $Script:StructureReadCache.ContainsKey($cacheKey) -and
        $Script:StructureReadCache.Count -ge $Script:StructureReadCacheLimit) {
        $oldestKey = @($Script:StructureReadCache.Keys)[0]
        if ($null -ne $oldestKey) { [void]$Script:StructureReadCache.Remove($oldestKey) }
    }
    $Script:StructureReadCache[$cacheKey] = [pscustomobject][ordered]@{ stamp = $stamp; structure = $structure }
    return $structure
}

# ページ構成の楽観ロック。並べ替えAPIは画面全体の並びを絶対値で送るため、
# これが無いと後から保存した側が相手の変更を無言で巻き戻す。
#
# 判定材料は「配置そのものの指紋」にする。structure全体の版番号だと、5分ごとの
# 更新スキャンやPDF作成のように配置を変えない書き込みでも増えてしまい、
# 実際には衝突していない操作を拒否してしまう。
function Get-PageLayoutFingerprint($Structure, [string]$PackId) {
    # /api/state がパックごとに毎回呼ぶ。Sort-Object とパイプラインは
    # PowerShell 5.1 では固定コストが大きく（20ページでも十数ms）、
    # 状態取得のたびに積み上がる。素のループと配列ソートで組み立てる。
    $workbookIds = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks)) {
        if (Test-WorkbookPack $wb $PackId) { $workbookIds[[string]$wb.workbookId] = $true }
    }
    $parts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($page in @(Get-Array $Structure.pages)) {
        if (-not $workbookIds.ContainsKey([string]$page.workbookId)) { continue }
        $parts.Add(('{0}~{1}~{2}~{3}' -f (Resolve-PageId $page),
            [string](Get-DataProperty $page 'volume' 'none'),
            ([double](Get-DataProperty $page 'order' 0)),
            ([bool](Get-DataProperty $page 'enabled' $false))))
    }
    if ($parts.Count -eq 0) { return 'empty' }
    # 先頭が pageId なので、合成文字列の序数ソートは pageId 順と同じ並びを与える。
    $sorted = $parts.ToArray()
    [Array]::Sort($sorted, [StringComparer]::Ordinal)
    $bytes = [Text.Encoding]::UTF8.GetBytes([string]::Join('|', $sorted))
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function New-StructureConflictError([string]$Expected, [string]$Actual) {
    $conflict = [System.InvalidOperationException]::new('他の画面でページ構成が変更されました。最新の状態を読み込み直してから操作してください。')
    $conflict.Data['reportBinderConflict'] = $true
    $conflict.Data['expectedLayout'] = $Expected
    $conflict.Data['actualLayout'] = $Actual
    return $conflict
}

function Test-StructureConflictError($Exception) {
    if ($null -eq $Exception -or $null -eq $Exception.Data) { return $false }
    try { return [bool]$Exception.Data['reportBinderConflict'] } catch { return $false }
}

function Get-RequestedBaseLayout($Body) {
    $raw = [string](Get-DataProperty $Body 'baseLayout' '')
    if ([string]::IsNullOrWhiteSpace($raw)) { return '' }
    if ($raw -notmatch '^[a-z0-9]{1,128}$') { return '' }
    return $raw
}

# $LayoutScope を渡した呼び出しだけが、ページ配置の衝突判定と指紋の返却を行う。
# パックIDの解決には structure が要るため、ロックを取った中で解決する。
function Update-StructureLocked([string]$Language, [scriptblock]$Mutation, [string]$BaseLayout = '', [string]$LayoutScope = $null) {
    $workspace = Get-WorkspacePath $Language
    $structureLock = Join-Path $workspace 'locks\structure.lock'
    $layoutAware = ($null -ne $LayoutScope)
    return Invoke-WithLock $structureLock {
        $structure = Read-StructureUnlocked $Language
        $layoutPackId = ''
        if ($layoutAware) {
            try { $layoutPackId = [string](Resolve-DocumentPackScope $structure ([string]$LayoutScope) $false).packId } catch { $layoutPackId = '' }
        }
        if ($layoutAware -and -not [string]::IsNullOrWhiteSpace($BaseLayout)) {
            $currentLayout = Get-PageLayoutFingerprint $structure $layoutPackId
            if ($BaseLayout -ne $currentLayout) { throw (New-StructureConflictError $BaseLayout $currentLayout) }
        }
        $result = & $Mutation $structure
        Write-StructureUnlocked $Language $structure
        # 同一ファイルシステム時刻内の連続更新でも古い読取結果を返さない。
        $Script:StructureReadCache.Clear()
        # 保存後の指紋を返す。画面はこれを次回の baseLayout として使う。
        if ($layoutAware -and $result -is [System.Collections.IDictionary]) {
            $result['layoutFingerprint'] = Get-PageLayoutFingerprint $structure $layoutPackId
        }
        return $result
    }
}

function Save-Structure([string]$Language, $Structure) {
    throw 'Save-Structureの直接呼出しは禁止されています。Update-StructureLockedを使用してください。'
}

function Get-EffectiveLanguage {
    # 既存の内部volume IDにだけ使用する互換キー。利用者向けの管理単位ではない。
    return 'ja'
}

function Get-VolumeList([string]$Language) {
    if ($Language -eq 'ja') { return @('ja-main','ja-appendix','none') }
    return @('en-main','en-appendix','none')
}


function Get-DefaultVolume([string]$Language) {
    # New pages are intentionally staged outside the final PDFs until the user assigns them.
    return 'none'
}
function Get-LanguageFromFileName([string]$FileName) {
    $name = [string]$FileName
    if ($name -match '(^|[_-])(J|JA|JPN)([_-]|$)') { return 'ja' }
    if ($name -match '(^|[_-])(E|EN|ENG)([_-]|$)') { return 'en' }
    return $null
}

function Normalize-WorkbookCategory([string]$Category, [string]$FileName = '') {
    $cat = ([string]$Category).Trim().ToLowerInvariant()
    if (@('ecm','bod','dmm') -contains $cat) { return $cat }
    $upper = ([string]$FileName).ToUpperInvariant()
    if ($upper -match '(^|[_-])ECM([_-]|$)' -or $upper.Contains('ECM')) { return 'ecm' }
    if ($upper -match '(^|[_-])BOD([_-]|$)' -or $upper.Contains('BOD')) { return 'bod' }
    if ($upper -match '(^|[_-])DMM([_-]|$)' -or $upper.Contains('DMM')) { return 'dmm' }
    return ''
}

function Require-WorkbookCategory([string]$Category) {
    $cat = Normalize-WorkbookCategory $Category ''
    if (@('ecm','bod','dmm') -notcontains $cat) {
        throw [System.ArgumentException]::new('categoryには ecm / bod / dmm のいずれかを指定してください。')
    }
    return $cat
}

function Get-PageCategory($Structure, $Page) {
    if ($null -eq $Page) { return '' }
    $wb = @(Get-Array $Structure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$Page.workbookId } | Select-Object -First 1)
    if ($wb.Count -eq 0) { return '' }
    return Normalize-WorkbookCategory ([string]$wb[0].category) ([string]$wb[0].fileName)
}

function Test-WorkbookCategory($Workbook, [string]$Category) {
    $cat = Normalize-WorkbookCategory $Category ''
    if ([string]::IsNullOrWhiteSpace($cat)) { return $true }
    $stored = Normalize-WorkbookCategory ([string]$Workbook.category) ''
    if (-not [string]::IsNullOrWhiteSpace($stored)) { return $stored -eq $cat }
    $name = ([string]$Workbook.fileName)
    if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$Workbook.displayName }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$Workbook.relativePath }
    return ((Normalize-WorkbookCategory '' $name) -eq $cat)
}

function Get-ProjectIdFromWorkbooks($Workbooks) {
    foreach ($wb in (Get-Array $Workbooks)) {
        $name = [string]$wb.fileName
        if ($name -match '^(?<p>.+)_(J|E)_') { return $matches['p'] }
    }
    return 'ReportBinder'
}

function Get-CategoryProjectId([string]$ProjectId, [string]$Category) {
    $cat = Require-WorkbookCategory $Category
    $categoryLabel = $cat.ToUpperInvariant()
    $base = ([string]$ProjectId).Trim()
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'ReportBinder' }
    if ($base -match '^(?<prefix>.*?)(?:[_-](?:ECM|BOD|DMM))$') {
        $prefix = (([string]$matches['prefix']) -replace '[_-]+$','')
        if ([string]::IsNullOrWhiteSpace($prefix)) { return $categoryLabel }
        return "${prefix}_${categoryLabel}"
    }
    return "${base}_${categoryLabel}"
}

function Get-OutputFileName([string]$Volume, [string]$ProjectId, [string]$Category) {
    $namedProjectId = Get-CategoryProjectId $ProjectId $Category
    switch ($Volume) {
        'ja-main' { return "${namedProjectId}_J_本体.pdf" }
        'ja-appendix' { return "${namedProjectId}_J_補足.pdf" }
        'en-main' { return "${namedProjectId}_E_Main.pdf" }
        'en-appendix' { return "${namedProjectId}_E_Appendix.pdf" }
        default { throw "未知の成果物: $Volume" }
    }
}

function Get-SheetOrderNumber([string]$SheetName) {
    $sheetNum = 0
    if ([int]::TryParse($SheetName, [ref]$sheetNum)) { return $sheetNum }
    return 999999
}

function Get-FileOrderNumber([string]$FileName) {
    if ($FileName -match '_(?<num>\d{1,4})_') { return [int]$matches['num'] }
    return 999999
}

function Get-OrderHint([string]$FileName, [string]$SheetName) {
    # Default page insertion order is primarily the numeric worksheet name.
    # File order is only a tie breaker when several workbooks have the same sheet number.
    return ((Get-SheetOrderNumber $SheetName) * 100000) + (Get-FileOrderNumber $FileName)
}

function Get-ExcelFilesInSubmission {
    $paths = Get-Paths
    if ([string]::IsNullOrWhiteSpace([string]$paths.submissionDir) -or -not (Test-Path -LiteralPath ([string]$paths.submissionDir))) { return @() }

    # Only show Excel files directly under the submission folder.
    # The default data/output child folders are intentionally ignored.
    $excelExts = @('.xlsx','.xlsm')
    $files = Get-ChildItem -LiteralPath ([string]$paths.submissionDir) -File -ErrorAction SilentlyContinue |
        Where-Object { $excelExts -contains $_.Extension.ToLowerInvariant() -and $_.Name -notlike '~$*' } |
        Sort-Object Name

    $result = @()
    foreach ($entry in $files) {
        # Network shares may briefly retain directory-enumeration metadata. Re-read and refresh
        # each FileInfo so the UI shows the newest save time available from the file server.
        try {
            $f = Get-Item -LiteralPath $entry.FullName -Force -ErrorAction Stop
            $f.Refresh()
        } catch {
            $f = $entry
        }
        $rel = Get-RelativePathCompat ([string]$paths.submissionDir) $f.FullName
        # Safety: direct children only. Do not accept any relative path containing a separator.
        if ($rel -match '[\\/]') { continue }
        # Explorer and the app must show the same minute. Read the timestamp from an
        # open file handle (which bypasses stale SMB directory metadata), convert it on the
        # server PC, and return the final display string without browser timezone conversion.
        $modifiedSnapshot = Get-FileLastWriteSnapshot $f.FullName
        $result += [ordered]@{
            fileName = $f.Name
            relativePath = $rel
            size = $f.Length
            modifiedAtDisplay = [string]$modifiedSnapshot.display
            modifiedAt = ([DateTime]$modifiedSnapshot.local).ToString('yyyy-MM-ddTHH:mm:sszzz')
            modifiedAtUtc = ([DateTime]$modifiedSnapshot.utc).ToString('o')
            modifiedAtUnixMs = [int64]$modifiedSnapshot.unixMs
            modifiedAtUtcTicks = [string]([DateTime]$modifiedSnapshot.utc).Ticks
            detectedLanguage = (Get-LanguageFromFileName $f.Name)
        }
    }
    return $result
}

function Get-PdfFilesInSubmission {
    $paths = Get-Paths
    if ([string]::IsNullOrWhiteSpace([string]$paths.submissionDir) -or -not (Test-Path -LiteralPath ([string]$paths.submissionDir))) { return @() }
    $files = Get-ChildItem -LiteralPath ([string]$paths.submissionDir) -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension.ToLowerInvariant() -eq '.pdf' } |
        Sort-Object Name
    $result = @()
    foreach ($entry in $files) {
        try { $f = Get-Item -LiteralPath $entry.FullName -Force -ErrorAction Stop; $f.Refresh() } catch { $f = $entry }
        $rel = Get-RelativePathCompat ([string]$paths.submissionDir) $f.FullName
        if ($rel -match '[\\/]') { continue }
        $modifiedSnapshot = Get-FileLastWriteSnapshot $f.FullName
        $result += [ordered]@{
            fileName = $f.Name; relativePath = $rel; size = $f.Length
            modifiedAtDisplay = [string]$modifiedSnapshot.display
            modifiedAt = ([DateTime]$modifiedSnapshot.local).ToString('yyyy-MM-ddTHH:mm:sszzz')
            modifiedAtUtc = ([DateTime]$modifiedSnapshot.utc).ToString('o')
            modifiedAtUnixMs = [int64]$modifiedSnapshot.unixMs
            modifiedAtUtcTicks = [string]([DateTime]$modifiedSnapshot.utc).Ticks
            detectedLanguage = (Get-LanguageFromFileName $f.Name)
        }
    }
    return $result
}

function Get-WordFilesInSubmission {
    $paths = Get-Paths
    if ([string]::IsNullOrWhiteSpace([string]$paths.submissionDir) -or -not (Test-Path -LiteralPath ([string]$paths.submissionDir))) { return @() }
    $files = Get-ChildItem -LiteralPath ([string]$paths.submissionDir) -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension.ToLowerInvariant() -eq '.docx' -and $_.Name -notlike '~$*' } |
        Sort-Object Name
    $result = @()
    foreach ($entry in $files) {
        try { $f = Get-Item -LiteralPath $entry.FullName -Force -ErrorAction Stop; $f.Refresh() } catch { $f = $entry }
        $rel = Get-RelativePathCompat ([string]$paths.submissionDir) $f.FullName
        if ($rel -match '[\\/]') { continue }
        $modifiedSnapshot = Get-FileLastWriteSnapshot $f.FullName
        $result += [ordered]@{
            fileName = $f.Name; relativePath = $rel; size = $f.Length
            modifiedAtDisplay = [string]$modifiedSnapshot.display
            modifiedAt = ([DateTime]$modifiedSnapshot.local).ToString('yyyy-MM-ddTHH:mm:sszzz')
            modifiedAtUtc = ([DateTime]$modifiedSnapshot.utc).ToString('o')
            modifiedAtUnixMs = [int64]$modifiedSnapshot.unixMs
            modifiedAtUtcTicks = [string]([DateTime]$modifiedSnapshot.utc).Ticks
            detectedLanguage = (Get-LanguageFromFileName $f.Name)
        }
    }
    return $result
}

function Get-PowerPointFilesInSubmission {
    $paths = Get-Paths
    if ([string]::IsNullOrWhiteSpace([string]$paths.submissionDir) -or -not (Test-Path -LiteralPath ([string]$paths.submissionDir))) { return @() }
    $files = Get-ChildItem -LiteralPath ([string]$paths.submissionDir) -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension.ToLowerInvariant() -eq '.pptx' -and $_.Name -notlike '~$*' } |
        Sort-Object Name
    $result = @()
    foreach ($entry in $files) {
        try { $f = Get-Item -LiteralPath $entry.FullName -Force -ErrorAction Stop; $f.Refresh() } catch { $f = $entry }
        $rel = Get-RelativePathCompat ([string]$paths.submissionDir) $f.FullName
        if ($rel -match '[\/]') { continue }
        $modifiedSnapshot = Get-FileLastWriteSnapshot $f.FullName
        $result += [ordered]@{
            fileName = $f.Name; relativePath = $rel; size = $f.Length
            modifiedAtDisplay = [string]$modifiedSnapshot.display
            modifiedAt = ([DateTime]$modifiedSnapshot.local).ToString('yyyy-MM-ddTHH:mm:sszzz')
            modifiedAtUtc = ([DateTime]$modifiedSnapshot.utc).ToString('o')
            modifiedAtUnixMs = [int64]$modifiedSnapshot.unixMs
            modifiedAtUtcTicks = [string]([DateTime]$modifiedSnapshot.utc).Ticks
            detectedLanguage = (Get-LanguageFromFileName $f.Name)
        }
    }
    return $result
}

function Test-PdfMagic([string]$PdfPath) {
    $stream = $null
    try {
        $stream = [IO.File]::Open($PdfPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $bytes = New-Object byte[] 5
        if ($stream.Read($bytes, 0, 5) -ne 5) { return $false }
        return ([Text.Encoding]::ASCII.GetString($bytes) -eq '%PDF-')
    } catch { return $false } finally { if ($stream) { $stream.Dispose() } }
}

function Invoke-PdfSourceSplit([string]$PdfPath, [string]$OutputDir) {
    if (-not (Test-Path -LiteralPath $PdfPath)) { throw "PDF原稿が見つかりません: $PdfPath" }
    if (-not (Test-PdfMagic $PdfPath)) { throw 'PDFのファイル形式を確認できません。拡張子だけがPDFになっていないか確認してください。' }
    $tool = Get-PdfBatchToolInfo
    if ($null -eq $tool) { throw 'PDF原稿の取り込みに必要なPDFBoxまたはJava Runtimeが見つかりません。' }
    if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
    $prefix = Join-Path $OutputDir 'page'
    $pdfboxJar = Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'
    $run = Invoke-NativeCapture ([string]$tool.javaExe) @('-jar', $pdfboxJar, 'PDFSplit', '-split', '1', '-outputPrefix', $prefix, $PdfPath) $Script:PdfSplitTimeoutSeconds
    $message = [string]$run.text
    # 時間切れを異常終了と混ぜない。健全な大きいPDFに「破損」と言ってしまうと、
    # 利用者は打つ手が無くなる。
    if ($run.timedOut) {
        throw ('このPDFの読み込みが{0}分以内に終わりませんでした。ページ数の多いPDFは分割してから登録してください。' -f [int]($Script:PdfSplitTimeoutSeconds / 60))
    }
    if ([int]$run.exitCode -ne 0) {
        if ($message -match 'password|encrypted|decrypt|InvalidPassword') { throw 'パスワード保護されたPDFは登録できません。保護を解除したPDFを使用してください。' }
        throw ('PDFを安全に読み込めませんでした。破損していないか確認してください。' + $(if ($message) { " 詳細: $message" } else { '' }))
    }
    $files = @(Get-ChildItem -LiteralPath $OutputDir -File -Filter 'page*.pdf' -ErrorAction SilentlyContinue | Sort-Object {
        $m = [regex]::Match($_.BaseName, '(\d+)$'); if ($m.Success) { [int]$m.Groups[1].Value } else { [int]::MaxValue }
    }, Name)
    if ($files.Count -eq 0) { throw 'PDFに取り込み可能なページがありません。' }
    if ($files.Count -gt 2000) { throw '2000ページを超えるPDFは登録できません。分割してから登録してください。' }
    $pages = @(); $number = 0
    foreach ($file in $files) {
        $number++
        if ($file.Length -le 0 -or -not (Test-PdfMagic $file.FullName)) { throw "PDFの $number ページ目を取り込めませんでした。" }
        $pages += [pscustomobject][ordered]@{ pageNumber = $number; sheetName = "Page $number"; sheetIndex = $number; detectedTitle = "Page $number"; pdf = $file.FullName; pageCount = 1; warnings = @() }
    }
    return [pscustomobject][ordered]@{ ok = $true; pageCount = $pages.Count; pages = @($pages); message = $message }
}

function Inspect-PdfSourceFile([string]$PdfPath) {
    $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-pdf-inspect-' + (New-RbId))
    try {
        New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
        $split = Invoke-PdfSourceSplit $PdfPath $tmpDir
        return [pscustomobject][ordered]@{
            pageCount = [int]$split.pageCount
            units = @($split.pages | ForEach-Object { [pscustomobject][ordered]@{ unitKind = 'pdf-page'; sourceKey = "page:$([int]$_.pageNumber)"; title = "Page $([int]$_.pageNumber)"; sourceIndex = [int]$_.pageNumber } })
            warnings = @()
        }
    } finally { if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue } }
}

function Inspect-WordSourceFile([string]$DocxPath) {
    if (-not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) { throw "Word原稿が見つかりません: $DocxPath" }
    $item = Get-Item -LiteralPath $DocxPath
    if ($item.Length -le 0) { throw 'Word原稿が空です。' }
    if ($item.Length -gt 536870912) { throw '512MBを超えるWord原稿は登録できません。分割してから登録してください。' }
    $stream = $null; $zip = $null
    try {
        $stream = [IO.File]::Open($DocxPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $header = New-Object byte[] 8
        $read = $stream.Read($header, 0, $header.Length)
        $stream.Dispose(); $stream = $null
        if ($read -ge 8 -and $header[0] -eq 0xD0 -and $header[1] -eq 0xCF -and $header[2] -eq 0x11 -and $header[3] -eq 0xE0) {
            throw 'パスワード保護または暗号化されたWord原稿は登録できません。保護を解除したDOCXを使用してください。'
        }
        if ($read -lt 4 -or $header[0] -ne 0x50 -or $header[1] -ne 0x4B) { throw 'Word原稿のファイル形式を確認できません。拡張子だけがDOCXになっていないか確認してください。' }
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
        $zip = [IO.Compression.ZipFile]::OpenRead($DocxPath)
        $names = @($zip.Entries | ForEach-Object { ([string]$_.FullName -replace '\\','/').ToLowerInvariant() })
        if ($names -notcontains '[content_types].xml' -or $names -notcontains 'word/document.xml') { throw 'Word原稿の内部構造が壊れています。Wordで開いて別名保存してから再登録してください。' }
        if (@($names | Where-Object { $_ -eq 'word/vbaproject.bin' }).Count -gt 0) { throw 'マクロを含むWord原稿は登録できません。.docx形式で保存してください。' }
        $title = [IO.Path]::GetFileNameWithoutExtension($item.Name)
        return [pscustomobject][ordered]@{
            pageCount = 0
            units = @([pscustomobject][ordered]@{ unitKind = 'document'; sourceKey = 'document:root'; title = $title; sourceIndex = 1 })
            warnings = @()
        }
    } catch {
        if ($_.Exception.Message -match 'Word原稿|DOCX|パスワード|マクロ') { throw }
        throw ('Word原稿を安全に読み込めませんでした。破損していないか確認してください。 詳細: ' + $_.Exception.Message)
    } finally {
        if ($zip) { $zip.Dispose() }
        if ($stream) { $stream.Dispose() }
    }
}

function New-WorkbookObject([string]$RelativePath, [string]$Language, [string]$Category = '') {
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $item = Get-Item -LiteralPath $full
    $hash = New-StableHash $full
    $categoryNormalized = Normalize-WorkbookCategory $Category $item.Name
    $baseId = New-Slug $item.Name
    # ECM / BOD / DMM are independent work sets. A single Excel file may be registered
    # in more than one category, so new workbook IDs include the category prefix.
    $id = $(if (-not [string]::IsNullOrWhiteSpace($categoryNormalized)) { "$categoryNormalized-$baseId" } else { $baseId })
    return [ordered]@{
        workbookId = $id
        language = $Language
        fileName = $item.Name
        relativePath = $RelativePath
        displayName = $item.Name
        category = $categoryNormalized
        sourceType = 'excel'
        # New Excel registrations follow the operational convention that only
        # half-width digit sheet names are submission pages.  Keep the mode
        # explicit in the persisted record so older records (which may not
        # have this property at all) can retain their legacy all-visible
        # behavior during Repair-StructurePages.
        excelSheetSelectionMode = 'numeric-only'
        lastRenderedSheetSelectionMode = 'numeric-only'
        currentExcelModifiedAt = $item.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:sszzz')
        currentExcelLastWriteUtcTicks = [string]$item.LastWriteTimeUtc.Ticks
        currentExcelSize = $item.Length
        currentExcelHash = $hash
        lastRenderedVersionId = $null
        lastRenderedExcelHash = $null
        lastRenderedAt = $null
        lastRenderedSheets = @()
        lastRenderedSheetFingerprint = ''
        status = 'new'
        warnings = @()
        lastError = ''
        lastErrorUser = ''
        lastErrorAt = $null
        lastRenderAttemptHash = ''
        lastRenderLog = ''
        renderProfileVersion = 0
        registeredAt = New-NowIso
    }
}

function Invoke-ComRelease($Obj) {
    if ($null -ne $Obj) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) | Out-Null } catch { }
    }
}


function Open-ExcelWorkbookSafe($Excel, [string]$FullPath, [bool]$ReadOnly) {
    $missing = [Type]::Missing

    # Mark of the Web (Zone.Identifier) が付いたファイルは保護ビュー対象となり、
    # COMの Workbooks.Open が「Workbooks クラスの Open プロパティを取得できません」で失敗する。
    # 事前にZone.Identifierを除去する（内容・更新日時は変わらない）。
    try { Unblock-File -LiteralPath $FullPath -ErrorAction SilentlyContinue } catch { }

    $openError = $null
    try {
        # Use explicit optional arguments. This avoids Excel interpreting a short Open() call differently
        # on some Office builds, and prevents link / read-only prompts.
        return $Excel.Workbooks.Open($FullPath, 0, $ReadOnly, $missing, $missing, $missing, $true, $missing, $missing, $false, $false, $missing, $false, $true, $missing)
    } catch {
        $openError = $_
    }
    try {
        # Compatibility fallback for older COM dispatchers.
        return $Excel.Workbooks.Open($FullPath, 0, $ReadOnly)
    } catch {
        $openError = $_
    }
    # 最終フォールバック: 保護ビュー経由で開いて編集モードへ昇格させる。
    # Unblock-Fileが効かない環境（グループポリシーで保護ビュー強制等）向け。
    try {
        $pvw = $Excel.ProtectedViewWindows.Open($FullPath)
        if ($null -ne $pvw) {
            $book = $pvw.Edit()
            if ($null -ne $book) { return $book }
        }
    } catch { }
    throw $openError
}


function Inspect-ExcelWorkbook([string]$FullPath) {
    $excel = $null
    $book = $null
    $sheets = @()
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.ScreenUpdating = $false
        try { $excel.AskToUpdateLinks = $false } catch { }
        try { $excel.AutomationSecurity = 3 } catch { }
        $book = Open-ExcelWorkbookSafe $excel $FullPath $true
        $sheetCount = 0
        try { $sheetCount = [int]$book.Worksheets.Count } catch { $sheetCount = 0 }
        for ($i = 1; $i -le $sheetCount; $i++) {
            $ws = $null
            try {
                $ws = $book.Worksheets.Item($i)
                $sheetName = [string]$ws.Name
                $visible = ([int]$ws.Visible -eq -1)
                if ($visible) {
                    $a1 = ''
                    try { $a1 = [string]$ws.Range('A1').Text } catch { $a1 = '' }
                    if ([string]::IsNullOrWhiteSpace($a1)) { $a1 = '' }
                    $zoom = $null
                    try { $zoom = $ws.PageSetup.Zoom } catch { }
                    $printArea = ''
                    try { $printArea = [string]$ws.PageSetup.PrintArea } catch { }
                    $sheets += [ordered]@{
                        sheetName = $sheetName
                        sheetIndex = $i
                        titleSource = 'A1'
                        detectedTitle = $a1
                        zoom = $zoom
                        printArea = $printArea
                    }
                }
            } finally {
                Invoke-ComRelease $ws
            }
        }
    } finally {
        if ($book) { try { $book.Close($false) } catch { } ; Invoke-ComRelease $book }
        if ($excel) { try { $excel.Quit() } catch { } ; Invoke-ComRelease $excel }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    return $sheets
}

function Get-SourceAdapterDescriptor([string]$SourceType) {
    $type = ([string]$SourceType).Trim().ToLowerInvariant()
    if ($type -eq 'excel') {
        return [pscustomobject][ordered]@{
            contractVersion = $Script:SourceAdapterContractVersion
            sourceType = 'excel'
            adapterId = 'excel-com-v1'
            adapterVersion = 1
            extensions = @('.xlsx','.xlsm')
            unitKind = 'worksheet'
            candidateScope = 'submission-root'
            supportsInspection = $true
            supportsRendering = $true
            supportsStructuralDiff = $true
            supportsVisualDiff = $true
        }
    }
    if ($type -eq 'pdf') {
        return [pscustomobject][ordered]@{
            contractVersion = $Script:SourceAdapterContractVersion
            sourceType = 'pdf'
            adapterId = 'pdfbox-import-v1'
            adapterVersion = 1
            extensions = @('.pdf')
            unitKind = 'pdf-page'
            candidateScope = 'submission-root'
            supportsInspection = $true
            supportsRendering = $true
            supportsStructuralDiff = $true
            supportsVisualDiff = $true
        }
    }
    if ($type -eq 'word') {
        return [pscustomobject][ordered]@{
            contractVersion = $Script:SourceAdapterContractVersion
            sourceType = 'word'
            adapterId = 'word-com-v1'
            adapterVersion = 1
            extensions = @('.docx')
            unitKind = 'document'
            candidateScope = 'submission-root'
            supportsInspection = $true
            supportsRendering = $true
            supportsStructuralDiff = $true
            supportsVisualDiff = $true
        }
    }
    if ($type -eq 'powerpoint') {
        return [pscustomobject][ordered]@{
            contractVersion = $Script:SourceAdapterContractVersion
            sourceType = 'powerpoint'
            adapterId = 'powerpoint-com-v1'
            adapterVersion = 1
            extensions = @('.pptx')
            unitKind = 'slide'
            candidateScope = 'submission-root'
            supportsInspection = $true
            supportsRendering = $true
            supportsStructuralDiff = $true
            supportsVisualDiff = $true
        }
    }
    throw [System.ArgumentException]::new("未対応の原稿形式です: $SourceType")
}

function Resolve-SourceTypeFromPath([string]$Path, [string]$RequestedType = '') {
    $requested = ([string]$RequestedType).Trim().ToLowerInvariant()
    if (-not [string]::IsNullOrWhiteSpace($requested)) {
        [void](Get-SourceAdapterDescriptor $requested)
        return $requested
    }
    $extension = [IO.Path]::GetExtension([string]$Path).ToLowerInvariant()
    if ($extension -in @('.xlsx','.xlsm')) { return 'excel' }
    if ($extension -eq '.pdf') { return 'pdf' }
    if ($extension -eq '.docx') { return 'word' }
    if ($extension -eq '.pptx') { return 'powerpoint' }
    throw [System.ArgumentException]::new("拡張子に対応する原稿アダプターがありません: $extension")
}

function Test-SourceCandidate($Context) {
    if ($null -eq $Context) { throw [System.ArgumentNullException]::new('Context') }
    $relativePath = [string](Get-DataProperty $Context 'relativePath' '')
    $sourceType = Resolve-SourceTypeFromPath $relativePath ([string](Get-DataProperty $Context 'sourceType' ''))
    $adapter = Get-SourceAdapterDescriptor $sourceType
    if ($sourceType -eq 'excel') { Test-DirectExcelRelativePath $relativePath | Out-Null }
    elseif ($sourceType -eq 'pdf') { Test-DirectPdfRelativePath $relativePath | Out-Null }
    elseif ($sourceType -eq 'word') { Test-DirectWordRelativePath $relativePath | Out-Null }
    elseif ($sourceType -eq 'powerpoint') { Test-DirectPowerPointRelativePath $relativePath | Out-Null }
    return [pscustomobject][ordered]@{ ok = $true; sourceType = $sourceType; adapterId = [string]$adapter.adapterId; relativePath = $relativePath }
}

function Get-SourceCandidates([string[]]$SourceTypes = @('excel')) {
    $requested = @($SourceTypes | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
    if ($requested.Count -eq 0) { $requested = @('excel') }
    $result = @()
    foreach ($sourceType in $requested) {
        $adapter = Get-SourceAdapterDescriptor $sourceType
        if ($sourceType -eq 'excel') {
            foreach ($candidate in @(Get-ExcelFilesInSubmission)) {
                Set-NoteProperty $candidate 'sourceType' $sourceType
                Set-NoteProperty $candidate 'adapterId' ([string]$adapter.adapterId)
                $result += $candidate
            }
        } elseif ($sourceType -eq 'pdf') {
            foreach ($candidate in @(Get-PdfFilesInSubmission)) {
                Set-NoteProperty $candidate 'sourceType' $sourceType
                Set-NoteProperty $candidate 'adapterId' ([string]$adapter.adapterId)
                $result += $candidate
            }
        } elseif ($sourceType -eq 'word') {
            foreach ($candidate in @(Get-WordFilesInSubmission)) {
                Set-NoteProperty $candidate 'sourceType' $sourceType
                Set-NoteProperty $candidate 'adapterId' ([string]$adapter.adapterId)
                $result += $candidate
            }
        } elseif ($sourceType -eq 'powerpoint') {
            foreach ($candidate in @(Get-PowerPointFilesInSubmission)) {
                Set-NoteProperty $candidate 'sourceType' $sourceType
                Set-NoteProperty $candidate 'adapterId' ([string]$adapter.adapterId)
                $result += $candidate
            }
        }
    }
    return @($result | Sort-Object relativePath, sourceType)
}

function Inspect-Source($Context) {
    $validated = Test-SourceCandidate $Context
    $paths = Get-Paths
    $fullPath = [string](Get-DataProperty $Context 'fullPath' '')
    if ([string]::IsNullOrWhiteSpace($fullPath)) { $fullPath = Join-Safe ([string]$paths.submissionDir) ([string]$validated.relativePath) }
    if (-not (Test-Path -LiteralPath $fullPath)) { throw "原稿ファイルが見つかりません: $($validated.relativePath)" }
    if ([string]$validated.sourceType -eq 'excel') {
        $units = @()
        foreach ($sheet in @(Inspect-ExcelWorkbook $fullPath)) {
            $units += [pscustomobject][ordered]@{
                unitKind = 'worksheet'
                sourceKey = "worksheet:$([string]$sheet.sheetName)"
                title = [string]$sheet.detectedTitle
                sourceIndex = [int]$sheet.sheetIndex
                legacySheet = $sheet
            }
        }
        return [pscustomobject][ordered]@{
            sourceType = 'excel'; adapterId = [string]$validated.adapterId; relativePath = [string]$validated.relativePath
            units = @($units); warnings = @()
        }
    }
    if ([string]$validated.sourceType -eq 'pdf') {
        $inspection = Inspect-PdfSourceFile $fullPath
        return [pscustomobject][ordered]@{
            sourceType = 'pdf'; adapterId = [string]$validated.adapterId; relativePath = [string]$validated.relativePath
            units = @($inspection.units); pageCount = [int]$inspection.pageCount; warnings = @($inspection.warnings)
        }
    }
    if ([string]$validated.sourceType -eq 'word') {
        $inspection = Inspect-WordSourceFile $fullPath
        return [pscustomobject][ordered]@{
            sourceType = 'word'; adapterId = [string]$validated.adapterId; relativePath = [string]$validated.relativePath
            units = @($inspection.units); pageCount = 0; warnings = @($inspection.warnings)
        }
    }
    if ([string]$validated.sourceType -eq 'powerpoint') {
        $inspection = Inspect-PowerPointSourceFile $fullPath
        return [pscustomobject][ordered]@{
            sourceType = 'powerpoint'; adapterId = [string]$validated.adapterId; relativePath = [string]$validated.relativePath
            units = @($inspection.units); pageCount = [int]$inspection.pageCount; warnings = @($inspection.warnings)
        }
    }
    throw [System.ArgumentException]::new("検査に対応していない原稿形式です: $($validated.sourceType)")
}

function Get-RegisteredSourceAdapterContext([string]$Language, [string]$SourceId) {
    if ([string]::IsNullOrWhiteSpace($SourceId)) { throw [System.ArgumentException]::new('sourceId が必要です。') }
    $structure = Get-Structure $Language
    $source = @(Get-Array $structure.sources | Where-Object { [string]$_.sourceId -eq $SourceId } | Select-Object -First 1)
    if ($source.Count -gt 0) {
        $sourceType = Resolve-SourceTypeFromPath ([string]$source[0].relativePath) ([string]$source[0].sourceType)
        return [pscustomobject][ordered]@{ source = $source[0]; sourceId = $SourceId; sourceType = $sourceType; adapter = Get-SourceAdapterDescriptor $sourceType }
    }
    $workbook = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $SourceId } | Select-Object -First 1)
    if ($workbook.Count -eq 0) { throw "登録済み原稿が見つかりません: $SourceId" }
    $sourceType = Resolve-SourceTypeFromPath ([string]$workbook[0].relativePath) 'excel'
    return [pscustomobject][ordered]@{ source = $workbook[0]; sourceId = $SourceId; sourceType = $sourceType; adapter = Get-SourceAdapterDescriptor $sourceType }
}

function Register-Source([string]$Language, [string]$RelativePath, [string]$PackId = '', [string]$SourceType = '') {
    $validated = Test-SourceCandidate ([pscustomobject]@{ relativePath = $RelativePath; sourceType = $SourceType })
    $structure = Get-Structure $Language
    $scope = Resolve-DocumentPackScope $structure $PackId $false
    $resolvedPackId = [string]$scope.packId
    $category = if ([bool]$scope.builtIn) { [string]$scope.category } else { 'ecm' }
    if ([string]$validated.sourceType -eq 'excel') {
        $result = Register-Workbook $Language $RelativePath $category $resolvedPackId
        $workbook = Get-DataProperty $result 'workbook' $null
        $source = [pscustomobject][ordered]@{
            sourceId = [string](Get-DataProperty $workbook 'workbookId' '')
            packId = $resolvedPackId
            relativePath = [string](Get-DataProperty $workbook 'relativePath' $RelativePath)
            sourceType = 'excel'
            adapterId = [string]$validated.adapterId
            displayName = [string](Get-DataProperty $workbook 'displayName' '')
            status = [string](Get-DataProperty $workbook 'status' 'new')
        }
        Set-NoteProperty $result 'source' $source
        Set-NoteProperty $result 'units' @()
        return $result
    }
    if ([string]$validated.sourceType -eq 'word') {
        $result = Register-WordSource $Language $RelativePath $category $resolvedPackId
        $workbook = Get-DataProperty $result 'workbook' $null
        $source = [pscustomobject][ordered]@{
            sourceId = [string](Get-DataProperty $workbook 'workbookId' '')
            packId = $resolvedPackId
            relativePath = [string](Get-DataProperty $workbook 'relativePath' $RelativePath)
            sourceType = 'word'; adapterId = [string]$validated.adapterId
            displayName = [string](Get-DataProperty $workbook 'displayName' '')
            status = [string](Get-DataProperty $workbook 'status' 'new')
        }
        Set-NoteProperty $result 'source' $source
        if (-not (Test-ConfigHasKey $result 'units')) { Set-NoteProperty $result 'units' @() }
        return $result
    }
    if ([string]$validated.sourceType -eq 'powerpoint') {
        $result = Register-PowerPointSource $Language $RelativePath $category $resolvedPackId
        $workbook = Get-DataProperty $result 'workbook' $null
        $source = [pscustomobject][ordered]@{
            sourceId = [string](Get-DataProperty $workbook 'workbookId' '')
            packId = $resolvedPackId
            relativePath = [string](Get-DataProperty $workbook 'relativePath' $RelativePath)
            sourceType = 'powerpoint'; adapterId = [string]$validated.adapterId
            displayName = [string](Get-DataProperty $workbook 'displayName' '')
            status = [string](Get-DataProperty $workbook 'status' 'new')
        }
        Set-NoteProperty $result 'source' $source
        if (-not (Test-ConfigHasKey $result 'units')) { Set-NoteProperty $result 'units' @() }
        return $result
    }
    if ([string]$validated.sourceType -eq 'pdf') {
        $result = Register-PdfSource $Language $RelativePath $category $resolvedPackId
        $workbook = Get-DataProperty $result 'workbook' $null
        $source = [pscustomobject][ordered]@{
            sourceId = [string](Get-DataProperty $workbook 'workbookId' '')
            packId = $resolvedPackId
            relativePath = [string](Get-DataProperty $workbook 'relativePath' $RelativePath)
            sourceType = 'pdf'; adapterId = [string]$validated.adapterId
            displayName = [string](Get-DataProperty $workbook 'displayName' '')
            status = [string](Get-DataProperty $workbook 'status' 'new')
        }
        Set-NoteProperty $result 'source' $source
        if (-not (Test-ConfigHasKey $result 'units')) { Set-NoteProperty $result 'units' @() }
        return $result
    }
    throw [System.ArgumentException]::new("登録に対応していない原稿形式です: $($validated.sourceType)")
}

function Register-SourcesBatch([string]$Language, $RelativePaths, [string]$PackId = '', [string]$SourceType = '') {
    $registered = @(); $errors = @(); $seen = @{}
    foreach ($raw in @(Get-Array $RelativePaths)) {
        $relativePath = [string]$raw
        if ([string]::IsNullOrWhiteSpace($relativePath)) { continue }
        $key = ($relativePath -replace '\\','/').ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        try {
            $resolvedType = Resolve-SourceTypeFromPath $relativePath $SourceType
            $result = Register-Source $Language $relativePath $PackId $resolvedType
            $source = $result.source; $workbook = $result.workbook
            $registered += [pscustomobject][ordered]@{
                relativePath = $relativePath; sourceId = [string]$source.sourceId; sourceType = [string]$source.sourceType; adapterId = [string]$source.adapterId
                workbookId = [string]$workbook.workbookId; fileName = [string]$workbook.fileName; displayName = [string]$workbook.displayName
                status = [string]$workbook.status; category = [string]$workbook.category; packId = [string]$source.packId
            }
        } catch { $errors += [pscustomobject][ordered]@{ relativePath = $relativePath; error = $_.Exception.Message; detail = [string]$_ } }
    }
    return [pscustomobject][ordered]@{
        requestedCount = @(Get-Array $RelativePaths).Count; registered = @($registered); errors = @($errors)
        registeredCount = $registered.Count; errorCount = $errors.Count
    }
}

function New-PdfSourceObject([string]$RelativePath, [string]$Language, [string]$Category = '', [int]$PageCount = 0) {
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $item = Get-Item -LiteralPath $full
    $hash = New-StableHash $full
    $categoryNormalized = Normalize-WorkbookCategory $Category $item.Name
    $baseId = New-Slug $item.Name
    $id = $(if (-not [string]::IsNullOrWhiteSpace($categoryNormalized)) { "$categoryNormalized-pdf-$baseId" } else { "pdf-$baseId" })
    return [ordered]@{
        workbookId = $id; language = $Language; fileName = $item.Name; relativePath = $RelativePath; displayName = $item.Name
        category = $categoryNormalized; sourceType = 'pdf'; adapterId = 'pdfbox-import-v1'; sourcePageCount = $PageCount
        currentExcelModifiedAt = $item.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:sszzz')
        currentExcelLastWriteUtcTicks = [string]$item.LastWriteTimeUtc.Ticks; currentExcelSize = $item.Length; currentExcelHash = $hash
        lastRenderedVersionId = $null; lastRenderedExcelHash = $null; lastRenderedAt = $null; lastRenderedSheets = @(); lastRenderedSheetFingerprint = ''
        status = 'new'; warnings = @(); lastError = ''; lastErrorUser = ''; lastErrorAt = $null; lastRenderAttemptHash = ''; lastRenderLog = ''
        renderProfileVersion = 0; registeredAt = New-NowIso
    }
}

function New-WordSourceObject([string]$RelativePath, [string]$Language, [string]$Category = '') {
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $item = Get-Item -LiteralPath $full
    $hash = New-StableHash $full
    $categoryNormalized = Normalize-WorkbookCategory $Category $item.Name
    $baseId = New-Slug $item.Name
    $id = $(if (-not [string]::IsNullOrWhiteSpace($categoryNormalized)) { "$categoryNormalized-word-$baseId" } else { "word-$baseId" })
    return [ordered]@{
        workbookId = $id; language = $Language; fileName = $item.Name; relativePath = $RelativePath; displayName = $item.Name
        category = $categoryNormalized; sourceType = 'word'; adapterId = 'word-com-v1'; sourcePageCount = 0
        currentExcelModifiedAt = $item.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:sszzz')
        currentExcelLastWriteUtcTicks = [string]$item.LastWriteTimeUtc.Ticks; currentExcelSize = $item.Length; currentExcelHash = $hash
        lastRenderedVersionId = $null; lastRenderedExcelHash = $null; lastRenderedAt = $null; lastRenderedSheets = @(); lastRenderedSheetFingerprint = ''
        status = 'new'; warnings = @(); lastError = ''; lastErrorUser = ''; lastErrorAt = $null; lastRenderAttemptHash = ''; lastRenderLog = ''
        renderProfileVersion = 0; registeredAt = New-NowIso
    }
}

function Normalize-ExcelSheetSelection([string]$Selection) {
    $value = ([string]$Selection).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($value)) { return 'all-visible' }
    if ($value -in @('all-visible','numeric-only')) { return $value }
    throw [ArgumentException]::new('Excelシートの対象指定が不正です。all-visible または numeric-only を指定してください。')
}

function Test-StrictNumericSheetName([string]$SheetName) {
    # Trimは行わない。半角数字だけという運用ルールに空白や全角数字を混ぜない。
    $raw = [string]$SheetName
    return ($raw.Length -gt 0 -and $raw -match '^[0-9]+$' -and $raw -notmatch '[\r\n]')
}

function Get-NumericSheetSortRank($Page, $WorkbookMap) {
    $workbookId = [string](Get-DataProperty $Page 'workbookId' '')
    $workbook = if ($null -ne $WorkbookMap -and $WorkbookMap.ContainsKey($workbookId)) { $WorkbookMap[$workbookId] } else { $null }
    $sourceType = [string](Get-DataProperty $workbook 'sourceType' 'excel')
    $sheetName = [string](Get-DataProperty $Page 'sheetName' '')
    $sheetHidden = [bool](Get-DataProperty $Page 'sheetHidden' $false)
    if ($sourceType -ne 'excel' -or $sheetHidden -or -not (Test-StrictNumericSheetName $sheetName)) {
        return [pscustomobject]@{ rank = 1; length = 999999; ordinal = ''; fileOrder = 999999; fileName = ''; sheetIndex = (Get-PageSheetIndex $Page); pageId = (Resolve-PageId $Page) }
    }
    # Do not cast to Int64/Double: workbook conventions can use arbitrary
    # digit lengths. Canonical length then ordinal comparison gives 2 < 10
    # and remains correct for 31+ digit sheet names.
    $ordinal = ($sheetName -replace '^0+', '')
    if ([string]::IsNullOrWhiteSpace($ordinal)) { $ordinal = '0' }
    return [pscustomobject]@{
        rank = 0
        length = $ordinal.Length
        ordinal = $ordinal
        fileOrder = Get-FileOrderNumber ([string](Get-DataProperty $workbook 'fileName' ''))
        fileName = [string](Get-DataProperty $workbook 'fileName' '')
        sheetIndex = (Get-PageSheetIndex $Page)
        pageId = (Resolve-PageId $Page)
    }
}

function Sort-PagesByNumericDefault($Pages, $WorkbookMap) {
    $items = @($Pages)
    if ($items.Count -le 1) { return $items }
    return @($items | Sort-Object `
        @{Expression={ (Get-NumericSheetSortRank $_ $WorkbookMap).rank }; Ascending=$true}, `
        @{Expression={ (Get-NumericSheetSortRank $_ $WorkbookMap).length }; Ascending=$true}, `
        @{Expression={ (Get-NumericSheetSortRank $_ $WorkbookMap).ordinal }; Ascending=$true}, `
        @{Expression={ (Get-NumericSheetSortRank $_ $WorkbookMap).fileOrder }; Ascending=$true}, `
        @{Expression={ (Get-NumericSheetSortRank $_ $WorkbookMap).fileName }; Ascending=$true}, `
        @{Expression={ (Get-NumericSheetSortRank $_ $WorkbookMap).sheetIndex }; Ascending=$true}, `
        @{Expression={ (Get-NumericSheetSortRank $_ $WorkbookMap).pageId }; Ascending=$true})
}

function Sort-NumericPagesWithinAnchors($Pages, $WorkbookMap) {
    # Explicit re-sorting is intentionally narrower than the old tab-order
    # action: only numeric Excel pages move. Word/PDF pages, non-numeric Excel
    # pages, and other in-lane anchors stay in their existing slots.
    $items = @($Pages)
    $numeric = @($items | Where-Object { (Get-NumericSheetSortRank $_ $WorkbookMap).rank -eq 0 })
    if ($numeric.Count -le 1) { return $items }
    $sortedNumeric = @(Sort-PagesByNumericDefault $numeric $WorkbookMap)
    $result = New-Object System.Collections.Generic.List[object]
    $numericIndex = 0
    foreach ($item in $items) {
        if ((Get-NumericSheetSortRank $item $WorkbookMap).rank -eq 0) {
            [void]$result.Add($sortedNumeric[$numericIndex])
            $numericIndex++
        } else {
            [void]$result.Add($item)
        }
    }
    return $result.ToArray()
}

function Set-ExcludedSheetPagesNotRendered($Structure, [string]$WorkbookId, [string[]]$ExcludedSheetNames, [string]$SelectionMode = 'numeric-only') {
    if ((Normalize-ExcelSheetSelection $SelectionMode) -ne 'numeric-only') { return }
    $excluded = @($ExcludedSheetNames | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($excluded.Count -eq 0) { return }
    $excludedSet = @{}
    foreach ($name in $excluded) { $excludedSet[$name] = $true }
    foreach ($page in @(Get-Array $Structure.pages | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $WorkbookId })) {
        $name = [string](Get-DataProperty $page 'sheetName' '')
        if (-not $excludedSet.ContainsKey($name)) { continue }
        # 配置・ページ名・ページ番号・使用範囲は保持する。今回の限定変換で古い
        # contentPdfだけを再利用しないため、出力対象のアーティファクトを外す。
        Set-NoteProperty $page 'volume' 'none'
        Set-NoteProperty $page 'enabled' $false
        Set-NoteProperty $page 'contentPdf' $null
        Set-NoteProperty $page 'status' 'not-rendered'
        Set-NoteProperty $page 'warnings' @('今回のPDF変換では半角数字シートだけを対象にしたため、PDF未作成です。')
        Set-NoteProperty $page 'sheetSelectionExcluded' $true
        Set-NoteProperty $page 'updatedAt' (New-NowIso)
    }
}

function Render-Source([string]$Language, [string]$SourceId, $SharedHost = $null, [bool]$KeepHostOpen = $false, [scriptblock]$ProgressCallback = $null,
                       [string]$SourceOverridePath = '', [string]$SourceSnapshotId = '', [string]$ExpectedSourceHash = '', [string]$ExcelSheetSelection = '') {
    $context = Get-RegisteredSourceAdapterContext $Language $SourceId
    if ([string]$context.sourceType -eq 'excel') {
        if ([string]::IsNullOrWhiteSpace($ExcelSheetSelection)) {
            $ExcelSheetSelection = [string](Get-DataProperty $context.source 'excelSheetSelectionMode' 'all-visible')
        }
        $ExcelSheetSelection = Normalize-ExcelSheetSelection $ExcelSheetSelection
        if ($KeepHostOpen -and $null -eq $SharedHost) {
            throw 'Excelを起動できないため、この原稿をPDF化できませんでした。'
        }
        $result = Render-Workbook $Language $SourceId $SharedHost $KeepHostOpen $ProgressCallback $SourceOverridePath $SourceSnapshotId $ExpectedSourceHash $ExcelSheetSelection
        Set-NoteProperty $result 'sourceId' $SourceId
        Set-NoteProperty $result 'sourceType' 'excel'
        Set-NoteProperty $result 'adapterId' ([string]$context.adapter.adapterId)
        return $result
    }
    if (-not [string]::IsNullOrWhiteSpace($ExcelSheetSelection) -and (Normalize-ExcelSheetSelection $ExcelSheetSelection) -eq 'numeric-only') {
        throw '半角数字シート指定はExcel原稿でのみ利用できます。'
    }
    if ([string]$context.sourceType -eq 'pdf') {
        $result = Render-PdfSource $Language $SourceId $ProgressCallback $SourceOverridePath $SourceSnapshotId $ExpectedSourceHash
        # A batch job collects this value after every source and performs analysis only
        # after the shared rendering lock has been released. Direct calls do the same here.
        if (-not $KeepHostOpen) {
            $pending = $Script:PendingAnalysis
            $Script:PendingAnalysis = $null
            if ($null -ne $pending) {
                try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered ([string](Get-DataProperty $pending 'sheetSelectionMode' 'all-visible'))) }
                catch { Write-Warning ('PDF原稿の画像解析に失敗しました: ' + $_.Exception.Message) }
            }
        }
        return $result
    }
    if ([string]$context.sourceType -eq 'word') {
        $result = Render-WordSource $Language $SourceId $ProgressCallback $SourceOverridePath $SourceSnapshotId $ExpectedSourceHash
        if (-not $KeepHostOpen) {
            $pending = $Script:PendingAnalysis
            $Script:PendingAnalysis = $null
            if ($null -ne $pending) {
                try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered ([string](Get-DataProperty $pending 'sheetSelectionMode' 'all-visible'))) }
                catch { Write-Warning ('Word原稿の画像解析に失敗しました: ' + $_.Exception.Message) }
            }
        }
        return $result
    }
    if ([string]$context.sourceType -eq 'powerpoint') {
        $result = Render-PowerPointSource $Language $SourceId $ProgressCallback $SourceOverridePath $SourceSnapshotId $ExpectedSourceHash
        if (-not $KeepHostOpen) {
            $pending = $Script:PendingAnalysis
            $Script:PendingAnalysis = $null
            if ($null -ne $pending) {
                try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered ([string](Get-DataProperty $pending 'sheetSelectionMode' 'all-visible'))) }
                catch { Write-Warning ('PowerPoint原稿の画像解析に失敗しました: ' + $_.Exception.Message) }
            }
        }
        return $result
    }
    throw [System.ArgumentException]::new("変換に対応していない原稿形式です: $($context.sourceType)")
}

function Get-SourceChangeSummary([string]$Language, [string]$SourceId, $Source = $null) {
    $context = Get-RegisteredSourceAdapterContext $Language $SourceId
    if ([string]$context.sourceType -eq 'excel') {
        $summary = Get-WorkbookChangeSummary $Language $SourceId $Source
        if ($null -eq $summary) { return $null }
        Set-NoteProperty $summary 'sourceId' $SourceId
        Set-NoteProperty $summary 'sourceType' 'excel'
        Set-NoteProperty $summary 'adapterId' ([string]$context.adapter.adapterId)
        return $summary
    }
    if ([string]$context.sourceType -eq 'pdf') {
        $summary = Get-WorkbookChangeSummary $Language $SourceId $Source
        if ($null -eq $summary) { return $null }
        Set-NoteProperty $summary 'sourceId' $SourceId; Set-NoteProperty $summary 'sourceType' 'pdf'; Set-NoteProperty $summary 'adapterId' ([string]$context.adapter.adapterId)
        return $summary
    }
    if ([string]$context.sourceType -eq 'word') {
        $summary = Get-WorkbookChangeSummary $Language $SourceId $Source
        if ($null -eq $summary) { return $null }
        Set-NoteProperty $summary 'sourceId' $SourceId; Set-NoteProperty $summary 'sourceType' 'word'; Set-NoteProperty $summary 'adapterId' ([string]$context.adapter.adapterId)
        return $summary
    }
    if ([string]$context.sourceType -eq 'powerpoint') {
        $summary = Get-WorkbookChangeSummary $Language $SourceId $Source
        if ($null -eq $summary) { return $null }
        Set-NoteProperty $summary 'sourceId' $SourceId; Set-NoteProperty $summary 'sourceType' 'powerpoint'; Set-NoteProperty $summary 'adapterId' ([string]$context.adapter.adapterId)
        return $summary
    }
    throw [System.ArgumentException]::new("差分要約に対応していない原稿形式です: $($context.sourceType)")
}

# 原稿ファイルの名前変更・移動に対応する。日本の事務では版をファイル名で管理する
# (〜_v2.xlsx など)ため、複数回使う利用者はいずれ必ず当たる。
#
# workbookId は据え置く。これは content-pdf\<workbookId>\ や input-history\<workbookId>\、
# locks\ のディレクトリ名そのもので、変えると変換PDFと履歴が孤児になる。ID から
# ファイル名を逆算している箇所は無い(New-Slug の呼び出しは登録時の4箇所だけ)ので、
# 参照先だけ差し替えれば、ページの並び順も出力先の振り分けも保たれる。
# 登録解除→再登録はページを丸ごと消すため、これが唯一の無害な直し方になる。
function Get-RelinkCandidates([string]$Language, [string]$SourceId) {
    $structure = Get-Structure $Language
    $workbook = @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $SourceId } | Select-Object -First 1)
    if ($workbook.Count -eq 0) { throw '指定された原稿が見つかりません。' }
    $w = $workbook[0]
    $sourceType = ([string](Get-DataProperty $w 'sourceType' 'excel')).ToLowerInvariant()
    $packId = [string](Get-WorkbookPackId $w)
    # 同じ一式内で既に使われているファイルは候補にしない(重複参照になる)。
    $taken = @{}
    foreach ($other in @(Get-Array (Get-DataProperty $structure 'workbooks' @()))) {
        if ([string](Get-DataProperty $other 'workbookId' '') -eq $SourceId) { continue }
        if ([string](Get-WorkbookPackId $other) -ne $packId) { continue }
        $rel = ([string](Get-DataProperty $other 'relativePath' '')).Replace('\', '/').ToLowerInvariant()
        if ($rel) { $taken[$rel] = $true }
    }
    $previousSize = [int64](Get-DataProperty $w 'currentExcelSize' -1)
    $previousHash = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
    $renderedHash = Normalize-FileHash ([string](Get-DataProperty $w 'lastRenderedExcelHash' ''))
    $paths = Get-Paths
    $candidates = @()
    foreach ($candidate in @(Get-SourceCandidates @($sourceType))) {
        $rel = [string](Get-DataProperty $candidate 'relativePath' '')
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }
        if ($taken.ContainsKey($rel.Replace('\', '/').ToLowerInvariant())) { continue }
        $full = Join-Safe ([string]$paths.submissionDir) $rel
        $size = -1
        try { $size = [int64](Get-Item -LiteralPath $full -ErrorAction Stop).Length } catch { $size = -1 }
        # 中身が同じかどうかは、まず大きさで絞ってからハッシュする。共有フォルダー上で
        # 全候補をハッシュすると待たされるため(Scan-Updates と同じ考え方)。
        $sameContent = $false
        if ($size -ge 0 -and ($size -eq $previousSize)) {
            $hash = Normalize-FileHash (New-StableHash $full)
            if ($hash -and (($hash -eq $previousHash) -or ($hash -eq $renderedHash))) { $sameContent = $true }
        }
        $candidates += [ordered]@{
            relativePath = $rel
            fileName = [string](Get-DataProperty $candidate 'fileName' '')
            modifiedAt = [string](Get-DataProperty $candidate 'modifiedAt' '')
            size = $size
            sameContent = $sameContent
        }
    }
    return [ordered]@{
        sourceId = $SourceId
        sourceType = $sourceType
        fileName = [string](Get-DataProperty $w 'fileName' '')
        relativePath = [string](Get-DataProperty $w 'relativePath' '')
        candidates = @($candidates | Sort-Object -Property @{Expression={ -[int]$_.sameContent }}, @{Expression={ [string]$_.modifiedAt }; Descending=$true})
    }
}

function Relink-Source([string]$Language, [string]$SourceId, [string]$RelativePath) {
    $id = ([string]$SourceId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw '付け替える原稿が指定されていません。' }
    $rel = ([string]$RelativePath).Trim()
    if ([string]::IsNullOrWhiteSpace($rel)) { throw '付け替え先のファイルが指定されていません。' }
    $paths = Get-Paths
    return Update-StructureLocked $Language {
        param($structure)
        $workbook = @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $id } | Select-Object -First 1)
        if ($workbook.Count -eq 0) { throw '指定された原稿が見つかりません。' }
        $w = $workbook[0]
        if ([string](Get-DataProperty $w 'status' '') -ne 'missing') {
            throw 'この原稿のファイルは見つかっています。付け替えは、ファイルが見つからない原稿にだけ行えます。'
        }
        $sourceType = ([string](Get-DataProperty $w 'sourceType' 'excel')).ToLowerInvariant()
        # 形式をまたぐ付け替えは許さない。ID接頭辞・unitKind・adapterId・
        # renderProfileVersion の意味が同時に壊れる。
        $checked = Test-SourceCandidate ([pscustomobject]@{ relativePath = $rel; sourceType = '' })
        if ([string]$checked.sourceType -ne $sourceType) {
            throw ('同じ種類の原稿にだけ付け替えられます（今は {0}、選んだファイルは {1}）。' -f $sourceType, [string]$checked.sourceType)
        }
        $packId = [string](Get-WorkbookPackId $w)
        foreach ($other in @(Get-Array (Get-DataProperty $structure 'workbooks' @()))) {
            if ([string](Get-DataProperty $other 'workbookId' '') -eq $id) { continue }
            if ([string](Get-WorkbookPackId $other) -ne $packId) { continue }
            if (([string](Get-DataProperty $other 'relativePath' '')).Replace('\','/').ToLowerInvariant() -eq $rel.Replace('\','/').ToLowerInvariant()) {
                throw 'そのファイルは、この一式の別の原稿として登録済みです。'
            }
        }
        $full = Join-Safe ([string]$paths.submissionDir) $rel
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw '付け替え先のファイルが見つかりません。' }
        $item = Get-Item -LiteralPath $full
        # 取り消せる場所を作ってから触る（シート構成の変化と同じ扱い）。
        [void](Save-LayoutSnapshot $Language $packId 'source-relinked' $structure)
        $hash = Normalize-FileHash (New-StableHash $full)
        $renderedHash = Normalize-FileHash ([string](Get-DataProperty $w 'lastRenderedExcelHash' ''))
        Set-NoteProperty $w 'relativePath' $rel
        Set-NoteProperty $w 'fileName' $item.Name
        Set-NoteProperty $w 'displayName' $item.Name
        Set-NoteProperty $w 'currentExcelModifiedAt' ($item.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:sszzz'))
        Set-NoteProperty $w 'currentExcelLastWriteUtcTicks' ([string]$item.LastWriteTimeUtc.Ticks)
        Set-NoteProperty $w 'currentExcelSize' $item.Length
        Set-NoteProperty $w 'currentExcelHash' $hash
        # 旧ファイルの検知版を指したままにしない。次の走査が新しい版を作る。
        Set-NoteProperty $w 'currentSnapshotId' ''
        Set-NoteProperty $w 'lastError' ''
        $sameContent = ($hash -and $renderedHash -and $hash -eq $renderedHash)
        if ($sameContent) {
            # 名前が変わっただけ。変換PDFはそのまま使えるので、作り直しは要らない。
            Set-NoteProperty $w 'status' 'rendered-unchecked'
        } else {
            Set-NoteProperty $w 'status' (Get-SourceUpdatedStatus $w)
            # 中身が違うのに古い変換PDFが提出用PDFに載らないよう、ページを古い印にする。
            foreach ($page in @(Get-Array (Get-DataProperty $structure 'pages' @()))) {
                if ([string](Get-DataProperty $page 'workbookId' '') -ne $id) { continue }
                if ([string](Get-DataProperty $page 'contentPdf' '')) { Set-NoteProperty $page 'status' 'stale' }
            }
            $affected = @(Get-Array (Get-DataProperty $structure 'pages' @()) |
                Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $id } |
                ForEach-Object { [string](Get-DataProperty $_ 'volume' '') } |
                Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
            Mark-VolumeNeedsRebuild $structure $Language $packId $affected 'source-updated' '原稿の参照先を付け替えました'
        }
        # 自動処理の安定待ちは旧ファイル基準なので捨てる。
        try { Remove-Item -LiteralPath (Join-Path (Get-WorkspacePath $Language) ("state\auto-render\{0}.json" -f $id)) -Force -ErrorAction SilentlyContinue } catch { }
        Write-HistoryEvent $Language 'source.relinked' ([ordered]@{ sourceId = $id; packId = $packId; relativePath = $rel; sameContent = $sameContent })
        return [ordered]@{ sourceId = $id; relativePath = $rel; fileName = $item.Name; sameContent = $sameContent; status = [string](Get-DataProperty $w 'status' '') }
    }
}

function Update-SourceMetadata([string]$Language, [string]$SourceId, $Patch) {
    $id = ([string]$SourceId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw '原稿IDが指定されていません。' }
    $hasOwner = Test-ConfigHasKey $Patch 'ownerDepartment'
    $hasRequired = Test-ConfigHasKey $Patch 'required'
    $hasDefaultTarget = Test-ConfigHasKey $Patch 'defaultTargetId'
    $hasRequirement = Test-ConfigHasKey $Patch 'requirementId'
    $hasSheetSelection = Test-ConfigHasKey $Patch 'excelSheetSelectionMode'
    if (-not ($hasOwner -or $hasRequired -or $hasDefaultTarget -or $hasRequirement -or $hasSheetSelection)) { throw '更新する原稿設定が指定されていません。' }
    $owner = if ($hasOwner) { ([string](Get-DataProperty $Patch 'ownerDepartment' '')).Trim() } else { '' }
    if ($owner.Length -gt 120) { throw '担当部署は120文字以内で入力してください。' }
    $defaultTarget = if ($hasDefaultTarget) { ([string](Get-DataProperty $Patch 'defaultTargetId' '')).Trim().ToLowerInvariant() } else { '' }
    $requirementId = if ($hasRequirement) { ([string](Get-DataProperty $Patch 'requirementId' '')).Trim() } else { '' }
    $sheetSelection = if ($hasSheetSelection) { Normalize-ExcelSheetSelection ([string](Get-DataProperty $Patch 'excelSheetSelectionMode' '')) } else { '' }
    return Update-StructureLocked $Language {
        param($structure)
        $source = @(Get-Array (Get-DataProperty $structure 'sources' @()) | Where-Object { [string](Get-DataProperty $_ 'sourceId' '') -eq $id } | Select-Object -First 1)
        if ($source.Count -eq 0) { throw '指定された原稿が見つかりません。' }
        $current = $source[0]
        $sourceType = ([string](Get-DataProperty $current 'sourceType' 'excel')).ToLowerInvariant()
        if ($hasSheetSelection -and $sourceType -ne 'excel') { throw '半角数字シート指定はExcel原稿でのみ利用できます。' }
        $pack = Get-PackRecord $structure ([string](Get-DataProperty $current 'packId' ''))
        if ($hasDefaultTarget -and $defaultTarget -ne 'unassigned' -and @(Get-PackTargetIds $Language $pack) -notcontains $defaultTarget) { throw '既定の振り分け先が不正です。' }
        $workbook = @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $id } | Select-Object -First 1)
        $selectionChanged = $false
        $selectionFallback = 'all-visible'
        if ($workbook.Count -gt 0) { $selectionFallback = [string](Get-DataProperty $workbook[0] 'excelSheetSelectionMode' 'all-visible') }
        if ($hasSheetSelection) {
            $oldSelection = Normalize-ExcelSheetSelection ([string](Get-DataProperty $current 'excelSheetSelectionMode' $selectionFallback))
            $selectionChanged = ($oldSelection -ne $sheetSelection)
            Set-NoteProperty $current 'excelSheetSelectionMode' $sheetSelection
            if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'excelSheetSelectionMode' $sheetSelection }
        }
        $requirement = $null
        if ($hasRequirement -and -not [string]::IsNullOrWhiteSpace($requirementId)) {
            [void](Assert-SafeStorageSegment $requirementId 'requirementId')
            $matches = @(
                Get-Array (Get-DataProperty (Get-PackEffectiveTemplate $Language $pack) 'sourceRequirements' @()) |
                    Where-Object { [string](Get-DataProperty $_ 'requirementId' '') -eq $requirementId } |
                    Select-Object -First 1
            )
            if ($matches.Count -eq 0) { throw '指定された必要原稿枠がこの一式にありません。' }
            $requirement = $matches[0]
            $sourceType = ([string](Get-DataProperty $current 'sourceType' 'excel')).ToLowerInvariant()
            if ($sourceType -notin @(Get-Array (Get-DataProperty $requirement 'acceptedSourceTypes' @('excel','word','pdf','powerpoint')))) { throw 'この必要原稿枠では選択した原稿形式を利用できません。' }
            $alreadyAssigned = @(
                Get-Array (Get-DataProperty $structure 'sources' @()) |
                    Where-Object {
                        [string](Get-DataProperty $_ 'sourceId' '') -ne $id -and
                        [string](Get-DataProperty $_ 'packId' '') -eq [string](Get-DataProperty $current 'packId' '') -and
                        [string](Get-DataProperty $_ 'requirementId' '') -eq $requirementId
                    }
            )
            if ($alreadyAssigned.Count -gt 0) { throw 'この必要原稿枠には別の原稿が割り当て済みです。先に割り当てを解除してください。' }
        }
        if ($hasOwner) {
            Set-NoteProperty $current 'ownerDepartment' $owner
            if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'ownerDepartment' $owner }
        }
        if ($hasRequired) {
            $required = [bool](Get-DataProperty $Patch 'required' $false)
            Set-NoteProperty $current 'required' $required
            if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'required' $required }
        }
        if ($hasDefaultTarget) {
            Set-NoteProperty $current 'defaultTargetId' $defaultTarget
            if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'defaultTargetId' $defaultTarget }
        }
        if ($hasRequirement) {
            Set-NoteProperty $current 'requirementId' $requirementId
            if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'requirementId' $requirementId }
            if ($null -ne $requirement) {
                if (-not $hasOwner -and [string]::IsNullOrWhiteSpace([string](Get-DataProperty $current 'ownerDepartment' ''))) {
                    $requirementOwner = [string](Get-DataProperty $requirement 'ownerDepartment' '')
                    Set-NoteProperty $current 'ownerDepartment' $requirementOwner
                    if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'ownerDepartment' $requirementOwner }
                }
                if (-not $hasRequired) {
                    $requirementRequired = [bool](Get-DataProperty $requirement 'required' $true)
                    Set-NoteProperty $current 'required' $requirementRequired
                    if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'required' $requirementRequired }
                }
                if (-not $hasDefaultTarget -and [string](Get-DataProperty $current 'defaultTargetId' 'unassigned') -eq 'unassigned') {
                    $requirementTarget = [string](Get-DataProperty $requirement 'defaultTargetId' 'unassigned')
                    Set-NoteProperty $current 'defaultTargetId' $requirementTarget
                    if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'defaultTargetId' $requirementTarget }
                }
            }
        }
        if ($selectionChanged) {
            $affectedVolumes = @(Get-Array (Get-DataProperty $structure 'pages' @()) |
                Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $id } |
                ForEach-Object { [string](Get-DataProperty $_ 'volume' '') } |
                Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
            foreach ($page in @(Get-Array (Get-DataProperty $structure 'pages' @()) | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $id })) {
                if (-not [string]::IsNullOrWhiteSpace([string](Get-DataProperty $page 'contentPdf' ''))) { Set-NoteProperty $page 'status' 'stale' }
            }
            if ($workbook.Count -gt 0) { Set-NoteProperty $workbook[0] 'status' (Get-SourceUpdatedStatus $workbook[0]) }
            Mark-VolumeNeedsRebuild $structure $Language ([string](Get-DataProperty $current 'packId' '')) $affectedVolumes 'sheet-selection-mode' 'ExcelシートのPDF化対象を変更しました'
        }
        return [pscustomobject][ordered]@{
            sourceId = $id
            ownerDepartment = [string](Get-DataProperty $current 'ownerDepartment' '')
            required = [bool](Get-DataProperty $current 'required' $true)
            defaultTargetId = [string](Get-DataProperty $current 'defaultTargetId' 'unassigned')
            requirementId = [string](Get-DataProperty $current 'requirementId' '')
            excelSheetSelectionMode = Normalize-ExcelSheetSelection ([string](Get-DataProperty $current 'excelSheetSelectionMode' $selectionFallback))
        }
    }
}

function New-PowerPointSourceObject([string]$RelativePath, [string]$Language, [string]$Category = '', [int]$SlideCount = 0) {
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $item = Get-Item -LiteralPath $full
    $hash = New-StableHash $full
    $categoryNormalized = Normalize-WorkbookCategory $Category $item.Name
    $baseId = New-Slug $item.Name
    $id = $(if (-not [string]::IsNullOrWhiteSpace($categoryNormalized)) { "$categoryNormalized-powerpoint-$baseId" } else { "powerpoint-$baseId" })
    return [ordered]@{
        workbookId=$id; language=$Language; fileName=$item.Name; relativePath=$RelativePath; displayName=$item.Name
        category=$categoryNormalized; sourceType='powerpoint'; adapterId='powerpoint-com-v1'; sourcePageCount=$SlideCount
        currentExcelModifiedAt=$item.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:sszzz')
        currentExcelLastWriteUtcTicks=[string]$item.LastWriteTimeUtc.Ticks; currentExcelSize=$item.Length; currentExcelHash=$hash
        lastRenderedVersionId=$null; lastRenderedExcelHash=$null; lastRenderedAt=$null; lastRenderedSheets=@(); lastRenderedSheetFingerprint=''
        status='new'; warnings=@(); lastError=''; lastErrorUser=''; lastErrorAt=$null; lastRenderAttemptHash=''; lastRenderLog=''
        renderProfileVersion=0; registeredAt=New-NowIso
    }
}

function Inspect-PowerPointSourceFile([string]$PptxPath) {
    if (-not (Test-Path -LiteralPath $PptxPath -PathType Leaf)) { throw "PowerPoint原稿が見つかりません: $PptxPath" }
    $item = Get-Item -LiteralPath $PptxPath
    if ($item.Length -le 0) { throw 'PowerPoint原稿が空です。' }
    if ($item.Length -gt 1073741824) { throw '1GBを超えるPowerPoint原稿は登録できません。分割してから登録してください。' }
    $stream = $null; $zip = $null
    try {
        $stream = [IO.File]::Open($PptxPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $header = New-Object byte[] 8
        $read = $stream.Read($header, 0, $header.Length)
        $stream.Dispose(); $stream = $null
        if ($read -ge 8 -and $header[0] -eq 0xD0 -and $header[1] -eq 0xCF -and $header[2] -eq 0x11 -and $header[3] -eq 0xE0) {
            throw 'パスワード保護または暗号化されたPowerPoint原稿は登録できません。保護を解除したPPTXを使用してください。'
        }
        if ($read -lt 4 -or $header[0] -ne 0x50 -or $header[1] -ne 0x4B) { throw 'PowerPoint原稿のファイル形式を確認できません。拡張子だけがPPTXになっていないか確認してください。' }
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
        $zip = [IO.Compression.ZipFile]::OpenRead($PptxPath)
        $names = @($zip.Entries | ForEach-Object { ([string]$_.FullName -replace '\\','/').ToLowerInvariant() })
        if ($names -notcontains '[content_types].xml' -or $names -notcontains 'ppt/presentation.xml') { throw 'PowerPoint原稿の内部構造が壊れています。PowerPointで開いて別名保存してから再登録してください。' }
        if (@($names | Where-Object { $_ -eq 'ppt/vbaproject.bin' }).Count -gt 0) { throw 'マクロを含むPowerPoint原稿は登録できません。.pptx形式で保存してください。' }
        $slides = @($names | Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' } | Sort-Object { [int]([regex]::Match($_, 'slide(\d+)\.xml$').Groups[1].Value) })
        if ($slides.Count -eq 0) { throw 'PowerPoint原稿にスライドがありません。' }
        if ($slides.Count -gt 5000) { throw '5000スライドを超えるPowerPoint原稿は登録できません。分割してから登録してください。' }
        $units = @()
        for ($index = 1; $index -le $slides.Count; $index++) {
            $units += [pscustomobject][ordered]@{ unitKind='slide'; sourceKey="slide:$index"; title="Slide $index"; sourceIndex=$index }
        }
        return [pscustomobject][ordered]@{ pageCount=$slides.Count; units=@($units); warnings=@() }
    } catch {
        if ($_.Exception.Message -match 'PowerPoint原稿|PPTX|パスワード|マクロ|スライド') { throw }
        throw ('PowerPoint原稿を安全に読み込めませんでした。破損していないか確認してください。 詳細: ' + $_.Exception.Message)
    } finally {
        if ($zip) { $zip.Dispose() }
        if ($stream) { $stream.Dispose() }
    }
}
function Add-NotePropertyIfMissing($Object, [string]$Name, $Value) {
    if ($null -eq $Object) { return }
    if ($Object -is [System.Collections.IDictionary]) {
        if (-not $Object.Contains($Name)) { $Object[$Name] = $Value }
        return
    }
    if ($null -eq $Object.PSObject.Properties[$Name]) {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    }
}


function Renumber-VolumeOrder($Structure, [string]$Volume, [string]$Category) {
    $scope = Resolve-DocumentPackScope $Structure $Category $true
    $packId = [string]$scope.packId
    $wbIds = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks | Where-Object { Test-WorkbookPack $_ $packId })) { $wbIds[[string]$wb.workbookId] = $true }
    $pages = @(Get-Array $Structure.pages | Where-Object {
        $wbIds.ContainsKey([string]$_.workbookId) -and (
            ($Volume -eq 'none' -and ([string]$_.volume -eq 'none' -or $_.enabled -eq $false)) -or
            ($Volume -ne 'none' -and [string]$_.volume -eq $Volume -and $_.enabled -ne $false)
        )
    } | Sort-Object {[double](Get-DataProperty $_ 'order' 0)}, {[string](Resolve-PageId $_)})
    for ($i=0; $i -lt $pages.Count; $i++) { Set-NoteProperty $pages[$i] 'order' (($i+1)*10) }
    return $pages
}


function Insert-PageInSheetOrder($Structure, $NewPage, [string]$Volume, [string]$Category) {
    $scope = Resolve-DocumentPackScope $Structure $Category $true
    $existing = @(Renumber-VolumeOrder $Structure $Volume ([string]$scope.packId))
    $workbookMap = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks)) { $workbookMap[[string](Get-DataProperty $wb 'workbookId' '')] = $wb }
    $newRank = Get-NumericSheetSortRank $NewPage $workbookMap
    $newIsNumeric = $newRank.rank -eq 0
    $numericExisting = @($existing | Where-Object { (Get-NumericSheetSortRank $_ $workbookMap).rank -eq 0 })
    $numericSorted = @(Sort-PagesByNumericDefault $numericExisting $workbookMap)
    $ordered = New-Object System.Collections.Generic.List[object]
    $insertedAtEnd = $false
    if ($newIsNumeric) {
        # A cover/Word/PDF page is an anchor. Replace only existing numeric
        # slots with the sorted numeric sequence; if the new page creates one
        # extra numeric item, put it directly after the last numeric slot (or
        # at the end when this lane had no numeric slot at all).
        $manualNumeric = (($numericExisting | ForEach-Object { Resolve-PageId $_ }) -join "`n") -ne (($numericSorted | ForEach-Object { Resolve-PageId $_ }) -join "`n")
        if ($manualNumeric) {
            foreach ($p in $existing) { [void]$ordered.Add($p) }
            [void]$ordered.Add($NewPage)
            $insertedAtEnd = $true
        } else {
            $withNew = @($numericExisting) + @($NewPage)
            $sortedWithNew = @(Sort-PagesByNumericDefault $withNew $workbookMap)
            $numericIndex = 0
            $lastNumericSlot = -1
            foreach ($p in $existing) {
                if ((Get-NumericSheetSortRank $p $workbookMap).rank -eq 0) {
                    [void]$ordered.Add($sortedWithNew[$numericIndex]); $numericIndex++; $lastNumericSlot = $ordered.Count - 1
                } else { [void]$ordered.Add($p) }
            }
            if ($numericIndex -lt $sortedWithNew.Count) {
                $ordered.Insert($lastNumericSlot + 1, $sortedWithNew[$numericIndex])
            }
            $insertedAtEnd = ($numericExisting.Count -eq 0)
        }
    } else {
        # Preserve the legacy deterministic insertion behavior for nonnumeric
        # Excel and non-Excel pages; only numeric Excel pages use the numeric
        # convention above.
        $manual = @($existing | Where-Object { [bool](Get-DataProperty $_ 'orderManual' $false) }).Count -gt 0
        foreach ($p in $existing) { [void]$ordered.Add($p) }
        if ($manual -or $existing.Count -eq 0) {
            [void]$ordered.Add($NewPage); $insertedAtEnd = $true
        } else {
            $newWorkbook = if ($workbookMap.ContainsKey([string](Get-DataProperty $NewPage 'workbookId' ''))) { $workbookMap[[string](Get-DataProperty $NewPage 'workbookId' '')] } else { $null }
            $newFileOrder = Get-FileOrderNumber ([string](Get-DataProperty $newWorkbook 'fileName' ''))
            $newFileName = [string](Get-DataProperty $newWorkbook 'fileName' '')
            $newSheetIndex = Get-PageSheetIndex $NewPage
            $at = $existing.Count
            for ($i=0; $i -lt $existing.Count; $i++) {
                $p = $existing[$i]; $wb = if ($workbookMap.ContainsKey([string](Get-DataProperty $p 'workbookId' ''))) { $workbookMap[[string](Get-DataProperty $p 'workbookId' '')] } else { $null }
                $compare = Get-FileOrderNumber ([string](Get-DataProperty $wb 'fileName' '')) - $newFileOrder
                if ($compare -eq 0) { $compare = [StringComparer]::OrdinalIgnoreCase.Compare([string](Get-DataProperty $wb 'fileName' ''), $newFileName) }
                if ($compare -eq 0) { $compare = (Get-PageSheetIndex $p) - $newSheetIndex }
                if ($compare -eq 0) { $compare = [StringComparer]::Ordinal.Compare((Resolve-PageId $p), (Resolve-PageId $NewPage)) }
                if ($compare -gt 0) { $at = $i; break }
            }
            $ordered.Insert($at, $NewPage)
        }
    }
    for ($i=0; $i -lt $ordered.Count; $i++) { Set-NoteProperty $ordered[$i] 'order' (($i+1)*10) }
    return [ordered]@{ insertedAtEnd=$insertedAtEnd; pages=$ordered.ToArray() }
}
function Initialize-Or-MigrateStructure([string]$Language, [string]$DataDir = '') {
    $workspace = Get-WorkspacePath $Language $DataDir
    $lockPath = Join-Path $workspace 'locks\structure.lock'
    return Invoke-WithLock $lockPath {
        $path = Join-Path $workspace 'structure.json'
        if (-not (Test-Path -LiteralPath $path)) { Write-StructureUnlocked $Language (New-EmptyStructure $Language) $DataDir; return [ordered]@{ created=$true; migrated=$false } }
        $structure = Read-StructureUnlocked $Language $DataDir
        $version = 1
        try { $version=[int](Get-DataProperty $structure 'schemaVersion' 1) } catch { $version=1 }
        if ($version -gt 3) { throw "この管理データは新しいschemaVersion=$versionです。対応するReportBinderを使用してください。" }
        if ($version -eq 3) {
            $lastGood = Join-Path $workspace 'structure.json.last-good'
            if (-not (Test-Path -LiteralPath $lastGood)) { Write-JsonFile $lastGood $structure }
            return [ordered]@{ created=$false; migrated=$false }
        }
        $backups = @()
        if ($version -eq 1) {
            $v1Backup = Join-Path $workspace 'structure.json.v1.bak'
            if (-not (Test-Path -LiteralPath $v1Backup)) {
                Copy-Item -LiteralPath $path -Destination $v1Backup -ErrorAction Stop
                if ((Get-Item $v1Backup).Length -ne (Get-Item $path).Length -or (New-StableHash $v1Backup) -ne (New-StableHash $path)) { throw 'structure.json.v1.bakの検証に失敗しました。' }
            }
            $backups += $v1Backup
            $structure = Normalize-StructureCollections $structure; [void](Repair-StructurePages $structure)
            $oldVolumes = Get-DataProperty $structure 'volumes' $null; $newVolumes = [ordered]@{}
            foreach ($volume in @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })) {
                foreach ($cat in @('ecm','bod','dmm')) {
                    $fresh = New-EmptyVolumeState
                    if ($cat -eq 'ecm' -and $null -ne $oldVolumes) {
                        $old = Get-DataProperty $oldVolumes $volume $null
                        if ($null -eq $old) { $old = Get-DataProperty $oldVolumes (Get-VolumeStateKey $volume $cat) $null }
                        if ($null -ne $old) { foreach ($name in @('status','lastBuiltAt','outputPdf','message','builtFingerprint','staleReasons')) { $value=Get-DataProperty $old $name $null; if ($null -ne $value) { Set-NoteProperty $fresh $name $value } } }
                    }
                    $newVolumes[(Get-VolumeStateKey $volume $cat)] = $fresh
                }
            }
            Set-NoteProperty $structure 'volumes' $newVolumes; Set-NoteProperty $structure 'schemaVersion' 2
            foreach ($cat in @('ecm','bod','dmm')) { foreach ($volume in @(Get-VolumeList $Language)) { [void](Renumber-VolumeOrder $structure $volume $cat) } }
        }
        $v2Backup = Join-Path $workspace 'structure.json.v2.bak'
        if (-not (Test-Path -LiteralPath $v2Backup)) {
            if ($version -eq 2) {
                Copy-Item -LiteralPath $path -Destination $v2Backup -ErrorAction Stop
                if ((Get-Item $v2Backup).Length -ne (Get-Item $path).Length -or (New-StableHash $v2Backup) -ne (New-StableHash $path)) { throw 'structure.json.v2.bakの検証に失敗しました。' }
            } else {
                Write-JsonFile $v2Backup $structure
                $verifiedV2 = Read-JsonFile $v2Backup $null
                if (-not (Test-StructureDocument $verifiedV2 $Language) -or [int](Get-DataProperty $verifiedV2 'schemaVersion' 0) -ne 2) { throw 'structure.json.v2.bakの検証に失敗しました。' }
            }
        }
        $backups += $v2Backup
        $structure = ConvertTo-StructureV3 $structure $Language $workspace
        Write-StructureUnlocked $Language $structure $DataDir
        $savedV3 = Read-JsonFile $path $null
        if (-not (Test-StructureDocument $savedV3 $Language) -or [int](Get-DataProperty $savedV3 'schemaVersion' 0) -ne 3) { throw 'schemaVersion 3の保存後検証に失敗しました。' }
        Write-JsonFile (Join-Path $workspace 'structure.json.last-good') $savedV3
        Write-JsonFile (Join-Path $workspace 'schema-version.json') ([ordered]@{ schemaVersion=3; migratedAt=(New-NowIso); backups=@($backups) })
        return [ordered]@{ created=$false; migrated=$true; backup=$v2Backup; backups=@($backups); migrationIssues=@(Get-Array $structure.migrationIssues) }
    }
}

function Apply-DefaultNumberingPerVolume([string]$Language, $Structure, [string]$Category) {
    $scope = Resolve-DocumentPackScope $Structure $Category $true
    $packId = [string]$scope.packId
    $wbIds = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks | Where-Object { Test-WorkbookPack $_ $packId })) { $wbIds[[string]$wb.workbookId] = $true }
    foreach ($vol in @(Get-PackVolumeList $Language $scope.pack $false)) {
        $setPages = @(Get-Array $Structure.pages | Where-Object { $wbIds.ContainsKey([string]$_.workbookId) -and $_.enabled -eq $true -and [string]$_.volume -eq $vol } | Sort-Object {[double](Get-DataProperty $_ 'order' 0)}, {[string](Resolve-PageId $_)})
        for ($i=0; $i -lt $setPages.Count; $i++) {
            $p=$setPages[$i]
            Add-NotePropertyIfMissing $p 'numberingManual' $false
            Add-NotePropertyIfMissing $p 'numberingDefault' 'first-page-none'
            if (-not [bool]$p.numberingManual) {
                Set-NoteProperty $p 'numberingMode' $(if ($i -eq 0) { 'none' } else { 'visible' })
                Set-NoteProperty $p 'numberingDefault' 'first-page-none'
            }
        }
    }
}

function ConvertTo-UserRenderError([string]$Message) {
    $text = [string]$Message
    $rawTail = if ($text.Length -gt 250) { ' (元のエラー: ' + $text.Substring(0, 250) + '…)' } else { ' (元のエラー: ' + $text + ')' }
    # 監視プロセスが Excel を終了させていたら、以降のCOM呼び出しはすべてRPCの失敗として
    # 返る。開く・閉じる・一括書き出しなど、どこで踏んでも原因は同じなので最初に判定する。
    if (Test-ExcelRenderWatchdogFired) {
        return 'Excelでの変換が進まなくなったため中止しました。Excelの画面に確認のダイアログが出ていないか確かめてから、もう一度お試しください。原稿が大きい場合は、シート数を減らすか分割すると通ることがあります。'
    }
    if ($text -match 'WORD_TIMEOUT') { return 'Word原稿のPDF変換が時間内に完了しませんでした。Wordの確認画面が開いていないか、文書が破損していないか確認してください。前回成功した変換PDFは保持されています。' }
    if ($text -match 'WORD_NOT_AVAILABLE|ActiveX component can.t create object|Class not registered') { return 'Microsoft Wordを起動できませんでした。このPCにデスクトップ版Wordがインストールされ、通常起動できることを確認してください。' }
    if ($text -match 'WORD_OPEN_FAILED') { return 'Wordが原稿を開けませんでした。パスワード保護、秘密度ラベル、破損、変換確認が必要な文書ではないか確認してください。' }
    if ($text -match 'WORD_EXPORT_FAILED|WORD_PDF_MISSING|WORD_PDF_EMPTY|WORD_PDF_INVALID') { return 'Word原稿からPDFを作成できませんでした。印刷レイアウト、プリンター設定、文書の保護状態を確認してください。' }
    if ($text -match 'Word原稿|DOCX|word-com|WORD_WORKER') { return ('Word原稿をPDF変換できませんでした。' + $rawTail) }
    if ($text -match 'POWERPOINT_TIMEOUT') { return 'PowerPoint原稿のPDF変換が時間内に完了しませんでした。PowerPointの確認画面が開いていないか、原稿が破損していないか確認してください。前回成功した変換PDFは保持されています。' }
    if ($text -match 'POWERPOINT_NOT_AVAILABLE') { return 'Microsoft PowerPointを起動できませんでした。このPCにデスクトップ版PowerPointがインストールされ、通常起動できることを確認してください。' }
    if ($text -match 'POWERPOINT_OPEN_FAILED') { return 'PowerPointが原稿を開けませんでした。パスワード保護、秘密度ラベル、破損、変換確認が必要な原稿ではないか確認してください。' }
    if ($text -match 'POWERPOINT_EXPORT_FAILED|POWERPOINT_PDF_MISSING|POWERPOINT_PDF_EMPTY|POWERPOINT_PDF_INVALID|POWERPOINT_SLIDE_COUNT_MISMATCH') { return 'PowerPoint原稿からPDFを作成できませんでした。スライド設定、原稿の保護状態、埋め込みメディアを確認してください。' }
    if ($text -match 'PowerPoint原稿|PPTX|powerpoint-com|POWERPOINT_WORKER') { return ('PowerPoint原稿をPDF変換できませんでした。' + $rawTail) }
    if ($text -match 'Open プロパティを取得できません|Unable to get the Open property') { return 'Excelがこのファイルを開けませんでした。ファイルが保護ビュー対象（インターネット由来）・暗号化（秘密度ラベル/IRM）・破損のいずれかの可能性があります。ファイルを右クリック→プロパティ→「許可する」にチェック後、もう一度PDF作成してください。' }
    if ($text -match 'このオブジェクトにプロパティ|stateSavedAt|プロパティ.*見つかりません|property.*not found|does not contain a property') { return 'PDF作成の進捗状態を更新できませんでした。ReportBinderを更新してから、もう一度PDF作成してください。' }
    if ($text -match 'PDF化対象のシート|PDF化対象の表示シート') { return 'PDF化対象の表示シートがありません。Excelで少なくとも1つのワークシートを表示してください。' }
    if ($text -match '提出ファイル|ファイルが見つかりません|ブックが見つかりません|not found|missing') { return '提出ファイルが見つかりません。削除・移動・名前変更されていないか確認してください。' }
    if ($text -match 'コピー前後|提出中|使用中|locked|lock|ロック') { return ('Excelが保存中または他の処理中です。保存が終わってから再度確認します。' + $rawTail) }
    if ($text -match 'Excel|COM|HRESULT|ExportAsFixedFormat|RPC') { return ('ExcelでPDF化できませんでした。' + $rawTail) }
    if ($text -match 'PDFを作成できません|PDFが作成されません|空のPDF|0 bytes') { return ('ExcelからPDFが出力されませんでした。印刷範囲とシート設定を確認してください。' + $rawTail) }
    if ($text.Length -gt 120) { return $text.Substring(0, 120) + '…' }
    return $text
}

function Get-SheetNameListFromInspection($Sheets) {
    $names = @()
    foreach ($s in @(Get-Array $Sheets)) {
        $name = [string](Get-DataProperty $s 'sheetName' '')
        if (-not [string]::IsNullOrWhiteSpace($name)) { $names += $name }
    }
    return @($names)
}


function Get-SheetFingerprintFromNames($SheetNames) {
    $names = @($SheetNames | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($names.Count -eq 0) { return '' }
    # Worksheet order is significant now that names are not required to be numeric.
    return ([string]::Join('|', [string[]]$names))
}
function Set-WorkbookRenderedSheetSnapshot($Workbook, $SheetNames) {
    $names = @($SheetNames | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    Set-NoteProperty $Workbook 'lastRenderedSheets' @($names)
    Set-NoteProperty $Workbook 'lastRenderedSheetFingerprint' (Get-SheetFingerprintFromNames $names)
}

function Test-WorkbookRenderedSheetContains($Workbook, [string]$SheetName) {
    if ($null -eq $Workbook) { return $false }
    $lastSheets = @(Get-Array (Get-DataProperty $Workbook 'lastRenderedSheets' @()) | ForEach-Object { [string]$_ })
    if ($lastSheets.Count -eq 0) { return $true } # Compatibility with workbooks rendered by older builds.
    return (@($lastSheets | Where-Object { $_ -eq [string]$SheetName }).Count -gt 0)
}

function Test-PageContentMatchesWorkbookVersion($Page, $Workbook) {
    if ($null -eq $Page -or $null -eq $Workbook) { return $false }
    $versionId = [string](Get-DataProperty $Workbook 'lastRenderedVersionId' '')
    if ([string]::IsNullOrWhiteSpace($versionId)) { return $true } # Older data: fall back to file existence and hash checks.
    $workbookId = [string](Get-DataProperty $Workbook 'workbookId' '')
    if ([string]::IsNullOrWhiteSpace($workbookId)) { return $false }
    $rel = Normalize-RelativeForCompare ([string](Get-DataProperty $Page 'contentPdf' ''))
    $expectedPrefix = Normalize-RelativeForCompare ("content-pdf\$workbookId\$versionId\")
    return ($rel.StartsWith($expectedPrefix))
}

function Add-StaleReason($VolumeState, [string]$Type, [string]$Detail) {
    if ($null -eq $VolumeState) { return }
    $reasons = @(Get-Array (Get-DataProperty $VolumeState 'staleReasons' @()))
    $existing = @($reasons | Where-Object { [string]$_.type -eq $Type } | Select-Object -First 1)
    $others = @($reasons | Where-Object { [string]$_.type -ne $Type })
    $count = 1
    $detailText = [string]$Detail
    if ($detailText -match '(?<n>\d+)件') { $count = [int]$matches['n'] }
    if ($existing.Count -gt 0) {
        $oldCount = Get-IntDataProperty $existing[0] 'count' 0
        if ($oldCount -le 0) {
            $oldDetail = [string](Get-DataProperty $existing[0] 'detail' '')
            if ($oldDetail -match '(?<n>\d+)件') { $oldCount = [int]$matches['n'] } else { $oldCount = 1 }
        }
        $count += $oldCount
        if ($detailText -match '\d+件') { $detailText = [regex]::Replace($detailText, '\d+件', "$count`件", 1) }
    }
    $reason = [ordered]@{ type=$Type; at=(New-NowIso); detail=$detailText; count=$count }
    $nextReasons = @($reason) + @($others)
    Set-NoteProperty $VolumeState 'staleReasons' @($nextReasons | Select-Object -First 10)
}

function Mark-VolumeNeedsRebuild($Structure, [string]$Language, [string]$Category, [string[]]$Volumes, [string]$Type, [string]$Detail) {
    $scope=Resolve-DocumentPackScope $Structure $Category $true
    foreach ($volume in @($Volumes | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($volume) -or $volume -eq 'none') { continue }
        $v=Get-PackOutputState $Structure $Language ([string]$scope.packId) $volume $false
        if ($null -eq $v) { continue }
        $built=[string](Get-DataProperty $v 'builtFingerprint' '')
        if (-not [string]::IsNullOrWhiteSpace($built)) {
            Set-NoteProperty $v 'status' 'needs-rebuild'
            Add-StaleReason $v $Type $Detail
        }
    }
}

function Mark-StructureVolumesNeedRebuild($Structure) {
    # Legacy compatibility only. New code must call Mark-VolumeNeedsRebuild with explicit category and volumes.
}

function Remove-ContentPdfFileSafe([string]$Workspace, [string]$RelativePdf) {
    if ([string]::IsNullOrWhiteSpace($RelativePdf)) { return }
    try {
        $workspaceFull = [IO.Path]::GetFullPath($Workspace)
        if (-not $workspaceFull.EndsWith([IO.Path]::DirectorySeparatorChar)) { $workspaceFull += [IO.Path]::DirectorySeparatorChar }
        $contentRoot = [IO.Path]::GetFullPath((Join-Path $Workspace 'content-pdf'))
        if (-not $contentRoot.EndsWith([IO.Path]::DirectorySeparatorChar)) { $contentRoot += [IO.Path]::DirectorySeparatorChar }
        $full = [IO.Path]::GetFullPath((Join-Path $Workspace $RelativePdf))
        if (-not $full.StartsWith($workspaceFull, [StringComparison]::OrdinalIgnoreCase)) { return }
        if (-not $full.StartsWith($contentRoot, [StringComparison]::OrdinalIgnoreCase)) { return }
        if (Test-Path -LiteralPath $full) { Remove-Item -LiteralPath $full -Force -ErrorAction SilentlyContinue }
        $dir = Split-Path -Parent $full
        while (-not [string]::IsNullOrWhiteSpace($dir)) {
            $dirFull = [IO.Path]::GetFullPath($dir)
            if (-not $dirFull.StartsWith($contentRoot, [StringComparison]::OrdinalIgnoreCase)) { break }
            $children = @(Get-ChildItem -LiteralPath $dirFull -Force -ErrorAction SilentlyContinue)
            if ($children.Count -gt 0) { break }
            Remove-Item -LiteralPath $dirFull -Force -ErrorAction SilentlyContinue
            $dir = Split-Path -Parent $dirFull
        }
    } catch { }
}


# $Sheets には今回のワークシート情報を渡す。通常は描画対象だけだが、Excelの
# numeric-only経路では除外シートのメタデータも含める。
# $PresentSheetNames には非表示を含む「原稿に存在するシート名」、$SelectedSheetNamesには
# 実際に描画するシート名、$HiddenSheetNamesには非表示シート名を渡す。後ろ3引数を
# 省略する経路（Word/PowerPoint/PDFなど）は従来どおり$Sheetsを描画対象として扱う。
function Update-WorkbookPagesFromInspection([string]$Language, $Structure, $Workbook, $Sheets, $PresentSheetNames = $null, $SelectedSheetNames = $null, $HiddenSheetNames = $null, [bool]$SaveSnapshot = $true) {
    $workbookId = [string](Get-DataProperty $Workbook 'workbookId' '')
    $packId = Get-WorkbookPackId $Workbook
    [void](Resolve-DocumentPackScope $Structure $packId $true)
    $pages = @(Get-Array $Structure.pages)
    $sortedSheets = @(Get-Array $Sheets | Sort-Object @{Expression={[int](Get-DataProperty $_ 'sheetIndex' 999999)};Ascending=$true}, @{Expression={Get-SheetOrderNumber ([string](Get-DataProperty $_ 'sheetName' ''))};Ascending=$true}, @{Expression={[string](Get-DataProperty $_ 'sheetName' '')};Ascending=$true})

    $current = @{}
    foreach ($sheetInfo in $sortedSheets) {
        $name = [string](Get-DataProperty $sheetInfo 'sheetName' '')
        if (-not [string]::IsNullOrWhiteSpace($name)) { $current[$name.ToLowerInvariant()] = $true }
    }
    $present = @{}
    if ($null -eq $PresentSheetNames) { foreach ($key in $current.Keys) { $present[$key] = $true } }
    else {
        foreach ($name in @($PresentSheetNames)) {
            $text = [string]$name
            if (-not [string]::IsNullOrWhiteSpace($text)) { $present[$text.ToLowerInvariant()] = $true }
        }
    }
    $selected = @{}
    if ($null -eq $SelectedSheetNames) {
        foreach ($sheetInfo in $sortedSheets) {
            $name = [string](Get-DataProperty $sheetInfo 'sheetName' '')
            if (-not [string]::IsNullOrWhiteSpace($name)) { $selected[$name.ToLowerInvariant()] = $true }
        }
    } else {
        foreach ($name in @($SelectedSheetNames)) {
            $text = [string]$name
            if (-not [string]::IsNullOrWhiteSpace($text)) { $selected[$text.ToLowerInvariant()] = $true }
        }
    }
    $hiddenNames = @{}
    foreach ($name in @($HiddenSheetNames)) {
        $text = [string]$name
        if (-not [string]::IsNullOrWhiteSpace($text)) { $hiddenNames[$text.ToLowerInvariant()] = $true }
    }

    $removed = @()
    $kept = @()
    $hidden = @()
    foreach ($page in $pages) {
        $belongs = ([string]$page.workbookId -eq $workbookId)
        $sheetKey = ([string]$page.sheetName).ToLowerInvariant()
        if (-not $belongs -or $current.ContainsKey($sheetKey)) {
            if ($belongs -and $hiddenNames.ContainsKey($sheetKey)) { Set-NoteProperty $page 'sheetHidden' $true; $hidden += (Resolve-PageId $page) }
            $kept += $page; continue
        }
        # 非表示にしただけのシートは原稿から消えていない。ページを削除すると配置・
        # ページ名・使用範囲・番号設定が失われ、再表示しても未振り分けに戻るだけになる。
        if ($present.ContainsKey($sheetKey)) {
            Set-NoteProperty $page 'sheetHidden' $true
            $hidden += (Resolve-PageId $page)
            $kept += $page
            continue
        }
        $removed += $page
    }
    if ($removed.Count -gt 0 -and $SaveSnapshot) {
        # 削除は取り消せないため、確定前に必ず復元ポイントを残す。
        [void](Save-LayoutSnapshot $Language $packId 'source-sheets-changed' $Structure)
    }
    $pages = @($kept)
    $Structure.pages = $pages

    # Keep existing IDs and assignments. Match existing pages by workbook + sheet name so
    # upgrading from legacy numeric IDs does not reset a user's page composition.
    $knownBySheet = @{}
    foreach ($page in $pages) {
        if ([string]$page.workbookId -ne $workbookId) { continue }
        $sheetKey = ([string]$page.sheetName).ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($sheetKey)) { $knownBySheet[$sheetKey] = $page }
    }

    # シート名の変更は「旧シートの消滅＋新シートの出現」として現れる。位置(sheetIndex)で
    # 対応付け、旧ページを作り直さずに引き継ぐ。pageId を保つことで、レイアウト履歴と
    # Ctrl+Z からの復元経路も生き残る。
    $renameBySheetKey = @{}
    if ($removed.Count -gt 0) {
        $newSheets = @($sortedSheets | Where-Object {
            $name = [string](Get-DataProperty $_ 'sheetName' '')
            $name -and (-not $knownBySheet.ContainsKey($name.ToLowerInvariant()))
        })
        $pool = New-Object System.Collections.ArrayList
        foreach ($page in $removed) { [void]$pool.Add($page) }
        foreach ($sheetInfo in $newSheets) {
            if ($pool.Count -eq 0) { break }
            $sheetIndex = [int](Get-DataProperty $sheetInfo 'sheetIndex' 999999)
            $match = @($pool | Where-Object { [int](Get-DataProperty $_ 'sheetIndex' -1) -eq $sheetIndex } | Select-Object -First 1)
            if ($match.Count -eq 0) {
                # 位置も変わった場合は、A1から拾った見出しが完全一致するときだけ同一視する。
                # 手掛かりなしで対応付けると、無関係な新規シートに前のページの配置と
                # 使用ページ範囲を引き継いでしまい、削除より分かりにくい誤りになる。
                $detected = [string](Get-DataProperty $sheetInfo 'detectedTitle' '')
                if (-not [string]::IsNullOrWhiteSpace($detected)) {
                    $match = @($pool | Where-Object { [string](Get-DataProperty $_ 'detectedTitle' '') -eq $detected } | Select-Object -First 1)
                }
            }
            if ($match.Count -eq 0) { continue }
            $name = [string](Get-DataProperty $sheetInfo 'sheetName' '')
            $renameBySheetKey[$name.ToLowerInvariant()] = $match[0]
            [void]$pool.Remove($match[0])
        }
        # 引き継いだページは「削除された」扱いにしない。
        $reattached = @($renameBySheetKey.Values)
        if ($reattached.Count -gt 0) { $removed = @($pool) }
    }

    $added = @()
    $updated = @()
    $renamed = @()
    $atEnd = @()
    $inOrder = 0
    $newTargetId = Get-NewItemTargetId $Structure $Language $Workbook
    $newVolume = Get-LegacyVolumeFromTargetId $Language $newTargetId
    foreach ($sheetInfo in $sortedSheets) {
        $sheet = [string](Get-DataProperty $sheetInfo 'sheetName' '')
        if ([string]::IsNullOrWhiteSpace($sheet)) { continue }
        $sheetKey = $sheet.ToLowerInvariant()
        $sheetIndex = [int](Get-DataProperty $sheetInfo 'sheetIndex' 999999)
        $title = [string](Get-DataProperty $sheetInfo 'detectedTitle' '')
        if ([string]::IsNullOrWhiteSpace($title)) { $title = "$([string]$Workbook.fileName) / $sheet" }

        if ($knownBySheet.ContainsKey($sheetKey)) {
            $page = $knownBySheet[$sheetKey]
            $pageId = Resolve-PageId $page
            if ([string]::IsNullOrWhiteSpace([string](Get-DataProperty $page 'pageId' ''))) { Set-NoteProperty $page 'pageId' $pageId }
            Set-NoteProperty $page 'sheetName' $sheet
            Set-NoteProperty $page 'sheetIndex' $sheetIndex
            Set-NoteProperty $page 'detectedTitle' $title
            Set-NoteProperty $page 'sheetHidden' ($hiddenNames.ContainsKey($sheetKey))
            if ([string]::IsNullOrWhiteSpace([string]$page.title)) { Set-NoteProperty $page 'title' $title }
            $updated += $pageId
            continue
        }

        if ($renameBySheetKey.ContainsKey($sheetKey)) {
            $page = $renameBySheetKey[$sheetKey]
            $previousSheet = [string]$page.sheetName
            # pageId は据え置く。差し替えるとレイアウト履歴の復元が対象を見失う。
            $pageId = Resolve-PageId $page
            Set-NoteProperty $page 'pageId' $pageId
            Set-NoteProperty $page 'sheetName' $sheet
            Set-NoteProperty $page 'sheetIndex' $sheetIndex
            Set-NoteProperty $page 'detectedTitle' $title
            Set-NoteProperty $page 'sheetHidden' ($hiddenNames.ContainsKey($sheetKey))
            Set-NoteProperty $page 'renamedFromSheetName' $previousSheet
            # 参照先のPDFは旧シート名で作られている。中身は作り直しになる。
            Set-NoteProperty $page 'contentPdf' $null
            Set-NoteProperty $page 'status' 'not-rendered'
            Set-NoteProperty $page 'updatedAt' (New-NowIso)
            if ([string]::IsNullOrWhiteSpace([string]$page.title)) { Set-NoteProperty $page 'title' $title }
            $Structure.pages = @(Get-Array $Structure.pages) + @($page)
            $knownBySheet[$sheetKey] = $page
            $updated += $pageId
            $renamed += [ordered]@{ pageId=$pageId; fromSheetName=$previousSheet; toSheetName=$sheet }
            continue
        }

        # Numeric-only rendering still supplies metadata for excluded visible
        # sheets, but must not create a new page record for a sheet that has
        # never been rendered. Existing records above retain their settings.
        if (-not $selected.ContainsKey($sheetKey)) { continue }

        $pageId = New-WorksheetPageId $workbookId $sheet
        $page = [pscustomobject][ordered]@{
            pageId=$pageId; workbookId=$workbookId; sheetName=$sheet; sheetIndex=$sheetIndex
            titleSource='A1'; detectedTitle=$title; title=$title
            volume=$newVolume; order=0; orderManual=$false
            numberingMode='visible'; numberingManual=$false; numberingDefault='first-page-none'
            enabled=($newVolume -ne 'none'); contentPdf=$null; status='not-rendered'; warnings=@(); sheetSelectionExcluded=$false; updatedAt=New-NowIso
        }
        $insert = Insert-PageInSheetOrder $Structure $page $newVolume $packId
        $Structure.pages = @(Get-Array $Structure.pages) + @($page)
        if ($insert.insertedAtEnd) { $atEnd += $pageId } else { $inOrder++ }
        $added += $pageId
        $knownBySheet[$sheetKey] = $page
    }

    $pagePack = (Resolve-DocumentPackScope $Structure $packId $true).pack
    # New workbooks start in the numeric-sheet convention. On an untouched
    # lane, normalize existing pages as well as newly inserted pages so the
    # first page-composition view is numeric without changing any assignment.
    # Once a user has manually arranged a lane, preserve that lane verbatim on
    # later PDF refreshes.
    $workbookMap = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks | Where-Object { Test-WorkbookPack $_ $packId })) {
        $workbookMap[[string](Get-DataProperty $wb 'workbookId' '')] = $wb
    }
    foreach ($volume in @(Get-PackVolumeList $Language $pagePack $true)) {
        $lane = @(Get-Array $Structure.pages | Where-Object {
            $workbookMap.ContainsKey([string](Get-DataProperty $_ 'workbookId' '')) -and (
                ($volume -eq 'none' -and ([string](Get-DataProperty $_ 'volume' 'none') -eq 'none' -or (Get-DataProperty $_ 'enabled' $true) -eq $false)) -or
                ($volume -ne 'none' -and [string](Get-DataProperty $_ 'volume' '') -eq $volume -and (Get-DataProperty $_ 'enabled' $true) -ne $false)
            )
        } | Sort-Object @{Expression={ [double](Get-DataProperty $_ 'order' 0) };Ascending=$true}, @{Expression={ Resolve-PageId $_ };Ascending=$true})
        if ($lane.Count -eq 0 -or @($lane | Where-Object { [bool](Get-DataProperty $_ 'orderManual' $false) }).Count -gt 0) { continue }
        $sortedLane = @(Sort-NumericPagesWithinAnchors $lane $workbookMap)
        for ($i = 0; $i -lt $sortedLane.Count; $i++) { Set-NoteProperty $sortedLane[$i] 'order' (($i + 1) * 10) }
    }
    foreach ($volume in @(Get-PackVolumeList $Language $pagePack $true)) { [void](Renumber-VolumeOrder $Structure $volume $packId) }
    if ($removed.Count -gt 0) {
        $affected = @($removed | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
        if ($affected.Count -gt 0) {
            Mark-VolumeNeedsRebuild $Structure $Language $packId $affected 'render' '原稿のページ構成が変更されました'
        }
    }
    Apply-DefaultNumberingPerVolume $Language $Structure $packId
    $names = @($sortedSheets | Where-Object { $selected.ContainsKey(([string]$_.sheetName).ToLowerInvariant()) } | ForEach-Object { [string]$_.sheetName })
    return [ordered]@{
        addedPageIds=@($added); updatedPageIds=@($updated); removedPages=@($removed)
        renamedPages=@($renamed); hiddenSheetPageIds=@($hidden)
        sheetNames=$names; sheetFingerprint=(Get-SheetFingerprintFromNames $names)
        addedCount=$added.Count; insertedInOrderCount=$inOrder; insertedAtEndCount=$atEnd.Count
        insertedAtEndPageIds=@($atEnd); newPagesAreUnassigned=($newVolume -eq 'none'); newItemTargetId=$newTargetId
    }
}
function Clear-ExcelHeaderFooterParts($Target) {
    if ($null -eq $Target) { return }
    foreach ($part in @('LeftHeader','CenterHeader','RightHeader','LeftFooter','CenterFooter','RightFooter')) {
        try { $Target.$part = '' } catch { }
        try { $Target.$part.Text = '' } catch { }
    }
}

function Clear-ExcelHeaderFooterPictures($PageSetup) {
    if ($null -eq $PageSetup) { return }
    foreach ($part in @(
        'LeftHeaderPicture','CenterHeaderPicture','RightHeaderPicture',
        'LeftFooterPicture','CenterFooterPicture','RightFooterPicture'
    )) {
        try { $PageSetup.$part.FileName = '' } catch { }
        try { $PageSetup.$part.Filename = '' } catch { }
    }
}

function Set-OpenXmlAttributeValue($Node, [string]$Name, [string]$Value) {
    if ($null -eq $Node) { return }
    $attribute = $Node.Attributes.GetNamedItem($Name)
    if ($null -eq $attribute) {
        $attribute = $Node.OwnerDocument.CreateAttribute($Name)
        [void]$Node.Attributes.Append($attribute)
    }
    $attribute.Value = $Value
}

function Add-WorksheetElementOrdered($Document, $Root, $Node) {
    $rank = @{
        sheetPr=10; dimension=20; sheetViews=30; sheetFormatPr=40; cols=45; sheetData=50
        sheetCalcPr=60; sheetProtection=70; protectedRanges=80; scenarios=90; autoFilter=100
        sortState=110; dataConsolidate=120; customSheetViews=130; mergeCells=140
        phoneticPr=145; conditionalFormatting=150; dataValidations=155; hyperlinks=158
        printOptions=160; pageMargins=170; pageSetup=180; headerFooter=190
        rowBreaks=200; colBreaks=210; customProperties=220; cellWatches=230
        ignoredErrors=240; smartTags=250; drawing=270; legacyDrawing=280
        legacyDrawingHF=290; picture=300; oleObjects=310; controls=320
        webPublishItems=330; tableParts=340; extLst=1000
    }
    $targetRank = if ($rank.ContainsKey([string]$Node.LocalName)) { [int]$rank[[string]$Node.LocalName] } else { 999 }
    foreach ($child in @($Root.ChildNodes)) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $childRank = if ($rank.ContainsKey([string]$child.LocalName)) { [int]$rank[[string]$child.LocalName] } else { 999 }
        if ($childRank -gt $targetRank) {
            [void]$Root.InsertBefore($Node, $child)
            return $Node
        }
    }
    [void]$Root.AppendChild($Node)
    return $Node
}

function Get-OrCreateWorksheetElement($Document, $Root, [string]$LocalName) {
    $node = $Root.SelectSingleNode("./*[local-name()='$LocalName']")
    if ($null -ne $node) { return $node }
    $node = $Document.CreateElement($LocalName, [string]$Root.NamespaceURI)
    return (Add-WorksheetElementOrdered $Document $Root $node)
}

function Prepare-XlsxPrintPackage([string]$XlsxPath) {
    # PageSetup is one of the slowest Excel COM APIs. For OpenXML workbooks, persist the
    # standard one-page print profile directly in the local temporary package before Excel opens it.
    # Excel then only has to open and export the workbook; legacy .xls files retain the COM fallback.
    if ([string]::IsNullOrWhiteSpace($XlsxPath)) { return [ordered]@{ ok=$true; changed=0; sheetsPrepared=0; printSettingsPrepared=$false; skipped=$true } }
    $ext = [IO.Path]::GetExtension($XlsxPath).ToLowerInvariant()
    if (@('.xlsx','.xlsm','.xltx','.xltm') -notcontains $ext) { return [ordered]@{ ok=$true; changed=0; sheetsPrepared=0; printSettingsPrepared=$false; skipped=$true } }
    if (-not (Test-Path -LiteralPath $XlsxPath)) { return [ordered]@{ ok=$true; changed=0; sheetsPrepared=0; printSettingsPrepared=$false; skipped=$true } }

    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue | Out-Null
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
    } catch { }

    $zip = $null
    $changed = 0
    $sheetsPrepared = 0
    try {
        $zip = [System.IO.Compression.ZipFile]::Open($XlsxPath, [System.IO.Compression.ZipArchiveMode]::Update)
        $targets = @($zip.Entries | Where-Object { [string]$_.FullName -match '^xl/(worksheets|chartsheets)/[^/]+[.]xml$' })
        $utf8 = New-Object System.Text.UTF8Encoding -ArgumentList $false
        foreach ($entry in $targets) {
            $name = [string]$entry.FullName
            $xml = ''
            $reader = $null
            try {
                $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8, $true)
                $xml = $reader.ReadToEnd()
            } finally {
                if ($reader) { $reader.Dispose() }
            }
            if ([string]::IsNullOrWhiteSpace($xml)) { continue }

            $document = New-Object System.Xml.XmlDocument
            $document.PreserveWhitespace = $true
            $document.LoadXml($xml)
            $root = $document.DocumentElement
            $documentChanged = $false
            foreach ($headerFooter in @($root.SelectNodes("./*[local-name()='headerFooter']"))) {
                [void]$root.RemoveChild($headerFooter)
                $documentChanged = $true
            }

            if ($name -match '^xl/worksheets/') {
                $sheetPr = $root.SelectSingleNode("./*[local-name()='sheetPr']")
                if ($null -eq $sheetPr) {
                    $sheetPr = $document.CreateElement('sheetPr', [string]$root.NamespaceURI)
                    [void](Add-WorksheetElementOrdered $document $root $sheetPr)
                }
                $pageSetUpPr = $sheetPr.SelectSingleNode("./*[local-name()='pageSetUpPr']")
                if ($null -eq $pageSetUpPr) {
                    $pageSetUpPr = $document.CreateElement('pageSetUpPr', [string]$root.NamespaceURI)
                    [void]$sheetPr.AppendChild($pageSetUpPr)
                }
                Set-OpenXmlAttributeValue $pageSetUpPr 'fitToPage' '1'
                Set-OpenXmlAttributeValue $pageSetUpPr 'autoPageBreaks' '0'

                $printOptions = Get-OrCreateWorksheetElement $document $root 'printOptions'
                Set-OpenXmlAttributeValue $printOptions 'horizontalCentered' '1'
                Set-OpenXmlAttributeValue $printOptions 'verticalCentered' '0'

                $pageMargins = Get-OrCreateWorksheetElement $document $root 'pageMargins'
                Set-OpenXmlAttributeValue $pageMargins 'left' '0.47244094'
                Set-OpenXmlAttributeValue $pageMargins 'right' '0.47244094'
                Set-OpenXmlAttributeValue $pageMargins 'top' '0.31496063'
                Set-OpenXmlAttributeValue $pageMargins 'bottom' '0.31496063'
                Set-OpenXmlAttributeValue $pageMargins 'header' '0'
                Set-OpenXmlAttributeValue $pageMargins 'footer' '0'

                $pageSetup = Get-OrCreateWorksheetElement $document $root 'pageSetup'
                Set-OpenXmlAttributeValue $pageSetup 'fitToWidth' '1'
                Set-OpenXmlAttributeValue $pageSetup 'fitToHeight' '1'
                $scaleAttribute = $pageSetup.Attributes.GetNamedItem('scale')
                if ($null -ne $scaleAttribute) { [void]$pageSetup.Attributes.Remove($scaleAttribute) }
                $sheetsPrepared++
                $documentChanged = $true
            }

            if ($documentChanged) {
                $newXml = $document.OuterXml
                $entry.Delete()
                $newEntry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $writer = $null
                try {
                    $writer = New-Object IO.StreamWriter($newEntry.Open(), $utf8)
                    $writer.Write($newXml)
                } finally {
                    if ($writer) { $writer.Dispose() }
                }
                $changed++
            }
        }
        return [ordered]@{ ok=$true; changed=$changed; sheetsPrepared=$sheetsPrepared; printSettingsPrepared=($sheetsPrepared -gt 0); skipped=$false }
    } finally {
        if ($zip) { $zip.Dispose() }
    }
}

function Clear-ExcelPageSetupHeadersAndFooters($PageSetup) {
    if ($null -eq $PageSetup) { return }

    # Keep COM calls minimal. First-page/even-page header/footer XML is stripped from the
    # temporary XLSX package before opening, so normal PageSetup cleanup is enough here.
    # Toggling DifferentFirstPageHeaderFooter/OddAndEvenPagesHeaderFooter on some Excel builds
    # can be very slow or hang on the first sheet.
    Clear-ExcelHeaderFooterParts $PageSetup
    Clear-ExcelHeaderFooterPictures $PageSetup
    try { $PageSetup.DifferentFirstPageHeaderFooter = $false } catch { }
    try { $PageSetup.OddAndEvenPagesHeaderFooter = $false } catch { }
    try { $PageSetup.ScaleWithDocHeaderFooter = $false } catch { }
    try { $PageSetup.AlignMarginsHeaderFooter = $false } catch { }
    try { $PageSetup.HeaderMargin = 0 } catch { }
    try { $PageSetup.FooterMargin = 0 } catch { }
}

function Clear-ExcelWorksheetHeadersAndFooters($Worksheet, $Excel = $null) {
    if ($null -eq $Worksheet) { return }
    # Flush queued PageSetup changes first. Some Excel versions keep header/footer changes queued
    # while PrintCommunication is false, which can make ExportAsFixedFormat see old footer text.
    try { if ($null -ne $Excel) { $Excel.PrintCommunication = $true } } catch { }
    try { Clear-ExcelPageSetupHeadersAndFooters $Worksheet.PageSetup } catch { }
    try { if ($null -ne $Excel) { $Excel.PrintCommunication = $true } } catch { }
}

function Set-ExcelPrintCommunicationSafe($Excel, [bool]$Enabled) {
    if ($null -eq $Excel) { return $false }
    try {
        $Excel.PrintCommunication = $Enabled
        return $true
    } catch { return $false }
}

function Apply-StandardPrintSettings($Worksheet, $Excel = $null, [bool]$DeferPrintCommunication = $false) {
    $printCommunicationChanged = $false
    $ps = $null
    try {
        # PageSetup はExcel COMの中でも特に重い。複数シートのPDF作成時は、呼び出し元で
        # PrintCommunication を一時停止してからまとめて反映することで、シートごとの待ち時間を減らす。
        if (-not $DeferPrintCommunication) {
            $printCommunicationChanged = Set-ExcelPrintCommunicationSafe $Excel $false
        }

        try { $Worksheet.DisplayPageBreaks = $false } catch { }
        $ps = $Worksheet.PageSetup
        # 暫定PDFは左右1.2cmを基準にする。提出用PDFでは奇数/偶数ページを0.2cmだけ内側へ寄せる(パンチ側1.4cm/外側1.0cm)。
        $ps.TopMargin = Convert-CmToPt 0.8
        $ps.BottomMargin = Convert-CmToPt 0.8
        $ps.LeftMargin = Convert-CmToPt 1.2
        $ps.RightMargin = Convert-CmToPt 1.2
        # 印刷範囲が1ページ幅に満たないシートが左寄りに見えないよう、水平方向のみ中央揃えにする。
        # 幅いっぱいのシートは余白に接するため、センタリングしても位置は変わらない。
        $ps.CenterHorizontally = $true
        $ps.CenterVertically = $false
        # 印刷範囲は尊重しつつ、倍率はExcel設定を無視して1ページに収める。
        $ps.Zoom = $false
        $ps.FitToPagesWide = 1
        $ps.FitToPagesTall = 1
        Clear-ExcelPageSetupHeadersAndFooters $ps
    } finally {
        if ($printCommunicationChanged -and $null -ne $Excel) {
            [void](Set-ExcelPrintCommunicationSafe $Excel $true)
        }
    }

    if ($null -eq $ps) {
        try { Clear-ExcelWorksheetHeadersAndFooters $Worksheet $Excel } catch { }
    }
}

function Get-PdfBatchToolInfo {
    $composerJar = Join-Path $Script:AppRoot 'lib\pdfbox\ReportPdfComposer.jar'
    $pdfboxJar = Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'
    if (-not (Test-Path -LiteralPath $composerJar)) { return $null }
    if (-not (Test-Path -LiteralPath $pdfboxJar)) { return $null }
    try { $javaExe = Resolve-JavaExe } catch { return $null }
    return [ordered]@{ javaExe = $javaExe; classPath = "$composerJar;$pdfboxJar" }
}

function Split-BatchPdfToSheets([string]$BatchPdf, $SheetInfos, [string]$TmpDir) {
    $tool = Get-PdfBatchToolInfo
    if ($null -eq $tool) { return [ordered]@{ ok = $false; reason = 'pdfbox-unavailable'; message = 'PDFBox/Javaがないため一括PDF分割を使いません。' } }
    $mapPath = Join-Path $TmpDir 'batch-split-map.tsv'
    $lines = New-Object System.Collections.Generic.List[string]
    $pageNo = 1
    foreach ($info in @($SheetInfos)) {
        $out = [string]$info.outPdf
        $outParent = Split-Path -Parent $out
        if (-not (Test-Path -LiteralPath $outParent)) { New-Item -ItemType Directory -Path $outParent -Force | Out-Null }
        $lines.Add(("{0}`t{1}`t1" -f $out, $pageNo))
        $pageNo++
    }
    Write-Utf8NoBomFile $mapPath ([string]::Join("`n", $lines))
    $run = Invoke-NativeCapture ([string]$tool.javaExe) @('-cp', [string]$tool.classPath, 'BatchPdfSplitter', '--source', $BatchPdf, '--map', $mapPath) $Script:PdfSplitTimeoutSeconds
    $output = @($run.output)
    $global:LASTEXITCODE = [int]$run.exitCode
    $exit = $LASTEXITCODE
    $outputText = ($output -join "`n")
    if ($run.timedOut) { return [ordered]@{ ok = $false; reason = 'split-timeout'; message = ('PDFの分割が{0}分以内に終わりませんでした。' -f [int]($Script:PdfSplitTimeoutSeconds / 60)) } }
    if ($exit -ne 0) { return [ordered]@{ ok = $false; reason = 'split-failed'; message = $outputText } }
    foreach ($info in @($SheetInfos)) {
        if (-not (Test-Path -LiteralPath ([string]$info.outPdf))) { return [ordered]@{ ok = $false; reason = 'split-missing-output'; message = "分割後PDFが見つかりません: $($info.outPdf)" } }
        try {
            if ((Get-Item -LiteralPath ([string]$info.outPdf)).Length -le 0) { return [ordered]@{ ok = $false; reason = 'split-empty-output'; message = "分割後PDFが空です: $($info.outPdf)" } }
        } catch { return [ordered]@{ ok = $false; reason = 'split-check-failed'; message = $_.Exception.Message } }
    }
    return [ordered]@{ ok = $true; message = $outputText }
}

function Export-WorkbookSheetsToPdfBatch($Excel, $Workbook, $SheetInfos, [string]$TmpDir) {
    $infos = @($SheetInfos)
    if ($infos.Count -le 1) { return [ordered]@{ ok = $false; reason = 'single-sheet'; message = '1シートのため一括PDF化しません。' } }
    if ($null -eq (Get-PdfBatchToolInfo)) { return [ordered]@{ ok = $false; reason = 'pdfbox-unavailable'; message = 'PDFBox/Javaがないため従来方式でPDF化します。' } }

    $batchPdf = Join-Path $TmpDir 'batch-workbook.pdf'
    if (Test-Path -LiteralPath $batchPdf) { Remove-Item -LiteralPath $batchPdf -Force -ErrorAction SilentlyContinue }
    $missing = [Type]::Missing
    $selected = $false
    try {
        $first = $true
        foreach ($info in $infos) {
            $ws = $null
            try {
                $ws = $Workbook.Worksheets.Item([string]$info.sheetName)
                if ($first) { $ws.Select($true) | Out-Null; $first = $false }
                else { $ws.Select($false) | Out-Null }
            } finally {
                Invoke-ComRelease $ws
            }
        }
        $selected = $true
        $active = $Workbook.ActiveSheet
        try {
            # 選択した複数シートを1回でPDF化する。シートごとのExportAsFixedFormat回数を減らすのが狙い。
            # この1回の同期呼び出しの内側では心拍を打てないので、シート数ぶん猶予を伸ばしてから入る。
            Update-ExcelRenderHeartbeat (Get-ExcelBatchAllowanceSeconds $infos.Count)
            $active.ExportAsFixedFormat(0, $batchPdf, 0, $true, $false, $missing, $missing, $false, $missing)
            Update-ExcelRenderHeartbeat
        } finally {
            Invoke-ComRelease $active
        }
        [void](Wait-ForPdfOutput $batchPdf '一括PDF')
        $pageCount = Get-PdfPageCount $batchPdf
        if ($pageCount -ne $infos.Count) {
            return [ordered]@{ ok = $false; reason = 'page-count-mismatch'; message = "一括PDFのページ数が想定と異なります。想定=$($infos.Count) 実際=$pageCount" }
        }
        return (Split-BatchPdfToSheets $batchPdf $infos $TmpDir)
    } catch {
        return [ordered]@{ ok = $false; reason = 'batch-export-failed'; message = $_.Exception.Message }
    } finally {
        if ($selected -and $infos.Count -gt 0) {
            try { $Workbook.Worksheets.Item([string]$infos[0].sheetName).Select($true) | Out-Null } catch { }
        }
    }
}


function Wait-ForPdfOutput([string]$PdfPath, [string]$SheetName) {
    for ($i = 0; $i -lt 80; $i++) {
        if (Test-Path -LiteralPath $PdfPath) {
            try {
                $item = Get-Item -LiteralPath $PdfPath
                if ($item.Length -gt 0) { return $item }
            } catch { }
        }
        Start-Sleep -Milliseconds 250
    }
    if (-not (Test-Path -LiteralPath $PdfPath)) { throw "ExcelからPDFが出力されませんでした: シート $SheetName" }
    $item2 = Get-Item -LiteralPath $PdfPath
    if ($item2.Length -le 0) { throw "Excelから空のPDFが出力されました: シート $SheetName" }
    return $item2
}

function Export-WorksheetToPdfSafe($Excel, $Workbook, $Worksheet, [string]$OutPdf, [string]$SheetName, [bool]$AlreadyPrepared = $false) {
    $missing = [Type]::Missing
    $parent = Split-Path -Parent $OutPdf
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    if (Test-Path -LiteralPath $OutPdf) { Remove-Item -LiteralPath $OutPdf -Force -ErrorAction SilentlyContinue }

    # Export the cleaned worksheet directly. The previous one-sheet-copy method can hang on
    # cover workbooks with shapes or embedded objects. Header/footer definitions are already
    # removed from the temporary XLSX package, so a direct ExportAsFixedFormat is safer and faster.
    try {
        if (-not $AlreadyPrepared) { Apply-StandardPrintSettings $Worksheet $Excel }
        try { $Workbook.Activate() | Out-Null } catch { }
        try { $Worksheet.Activate() | Out-Null } catch { }
        try { $Worksheet.Select($true) | Out-Null } catch { }
        # Type=0 xlTypePDF, Quality=0 xlQualityStandard, IgnorePrintAreas=false.
        # 1シートごとに心拍を打つ。止まったまま一定時間が過ぎたら監視プロセスが
        # こちらの Excel を終了させ、この呼び出しは RPC の失敗として戻る。
        Update-ExcelRenderHeartbeat
        $Worksheet.ExportAsFixedFormat(0, $OutPdf, 0, $true, $false, $missing, $missing, $false, $missing)
        Update-ExcelRenderHeartbeat
        return (Wait-ForPdfOutput $OutPdf $SheetName)
    } catch {
        $directError = $_.Exception.Message
        if (Test-ExcelRenderWatchdogFired) {
            throw ("Excelでの変換が{0}分以上進まなかったため中止しました: シート {1}。Excelの画面に確認のダイアログが出ていないか確かめてから、もう一度お試しください。" -f [int]($Script:ExcelRenderTimeoutSeconds / 60), $SheetName)
        }
        throw "ExcelでPDF化できませんでした: シート $SheetName / 直接出力=[$directError]"
    } finally {
        try { $Workbook.Activate() | Out-Null } catch { }
    }
}

$Script:PdfPageCountCache = @{}
$Script:PdfPageCountCacheLimit = 4096

# 変換PDFのページ数は /api/state のたびに全ページ分が必要になる。毎回ファイル全体を
# 読み直すとdataDir全量の読み込みになるため、パスとサイズ・更新時刻でメモ化する。
function Get-PdfPageCount([string]$PdfPath) {
    $key = ''
    try {
        $item = Get-Item -LiteralPath $PdfPath -ErrorAction Stop
        $key = "$($item.FullName)|$($item.Length)|$($item.LastWriteTimeUtc.Ticks)"
        if ($Script:PdfPageCountCache.ContainsKey($key)) { return [int]$Script:PdfPageCountCache[$key] }
    } catch { $key = '' }
    try {
        $bytes = [IO.File]::ReadAllBytes($PdfPath)
        $text = [Text.Encoding]::ASCII.GetString($bytes)
        $count = ([regex]::Matches($text, '/Type\s*/Page(?!s)\b')).Count
        if ($count -lt 1) { $count = 1 }
        if ($key) {
            if ($Script:PdfPageCountCache.Count -ge $Script:PdfPageCountCacheLimit) { $Script:PdfPageCountCache.Clear() }
            $Script:PdfPageCountCache[$key] = $count
        }
        return $count
    } catch { return 1 }
}


# Excel の PDF 書き出しは同一プロセス内の同期COM呼び出しで、Word/PowerPoint のような
# 別プロセスのワーカーを通らない。固まると呼び出し元ごと止まり、「PDF作成を中止」も
# 現在の原稿が終わるまで効かないため、利用者にはタスクマネージャー以外の脱出手段が無い。
# PowerShell のタイマーやイベントでは救えない(COM呼び出しでパイプラインが塞がっている
# 間、ハンドラは実行されない)。そこで外部の監視プロセスに見張らせる。
# 進捗があるたびに心拍ファイルを更新し、一定時間更新が止まったら、こちらが起動した
# Excel だけを終了させる。COM呼び出しは RPC の失敗として戻り、既存の失敗処理に載る。
function Get-OwnedExcelProcessId($Excel) {
    if (-not ('ReportBinderNative.OfficeWindows' -as [type])) { return 0 }
    try {
        [uint32]$pidValue = 0
        [void][ReportBinderNative.OfficeWindows]::GetWindowThreadProcessId([IntPtr][int]$Excel.Hwnd, [ref]$pidValue)
        return [int]$pidValue
    } catch { return 0 }
}

# 心拍には「次の心拍までに許す時間」も一緒に書く。Excel の PDF 書き出しは、複数シートを
# 1回の ExportAsFixedFormat で出す経路(Export-WorkbookSheetsToPdfBatch)が既定で、
# その1回の同期呼び出しの内側では心拍を打てない。固定の猶予だと、正常に動いている
# 大きなブックを殺してしまうため、これから始める作業の重さに応じて猶予を伸ばす。
function Update-ExcelRenderHeartbeat([int]$AllowanceSeconds = 0) {
    if ([string]::IsNullOrWhiteSpace($Script:ExcelWatchdogHeartbeatPath)) { return }
    $allowance = if ($AllowanceSeconds -gt 0) { $AllowanceSeconds } else { [int]$Script:ExcelRenderTimeoutSeconds }
    try { [IO.File]::WriteAllText($Script:ExcelWatchdogHeartbeatPath, ('{0} {1}' -f [DateTimeOffset]::UtcNow.ToUnixTimeSeconds(), $allowance)) } catch { }
}

# 監視プロセスの作業ファイルは、サーバーが強制終了されると残る。放っておくと
# %TEMP% に無期限に溜まるので、監視を始めるたびに古いものを掃除する。
function Remove-StaleExcelWatchdogFiles {
    $runDir = Join-Path ([IO.Path]::GetTempPath()) 'ReportBinderExcelWatchdog'
    if (-not (Test-Path -LiteralPath $runDir)) { return }
    $cutoff = [DateTime]::UtcNow.AddHours(-6)
    try {
        foreach ($file in @(Get-ChildItem -LiteralPath $runDir -File -ErrorAction SilentlyContinue)) {
            if ($file.LastWriteTimeUtc -lt $cutoff) { Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue }
        }
    } catch { }
}

# シート数に応じた猶予。1シートあたりの上限を積み、下限は既定値。
function Get-ExcelBatchAllowanceSeconds([int]$SheetCount) {
    $perSheet = [int]$Script:ExcelPerSheetAllowanceSeconds
    $scaled = [int]$Script:ExcelRenderTimeoutSeconds + ($perSheet * [Math]::Max(0, $SheetCount))
    if ($scaled -lt [int]$Script:ExcelRenderTimeoutSeconds) { return [int]$Script:ExcelRenderTimeoutSeconds }
    return $scaled
}

function Start-ExcelRenderWatchdog($Excel) {
    # 監視の枠は1組しかない。前の Excel が生きているうちに2つ目を作ると、1つ目の心拍の
    # 場所を見失い、更新されないまま現役の Excel が終了させられる。今は経路の約束だけで
    # 二重起動を避けているので、約束が破れたことをここで検知する。
    if (-not [string]::IsNullOrWhiteSpace($Script:ExcelWatchdogHeartbeatPath) -and (Test-Path -LiteralPath $Script:ExcelWatchdogHeartbeatPath)) {
        throw 'Excelの監視がすでに動いています。前の変換を終えてから次を開始してください。'
    }
    $Script:ExcelWatchdogFiredSticky = $false
    $Script:ExcelWatchdogProcess = $null
    $Script:ExcelWatchdogHeartbeatPath = ''
    $Script:ExcelWatchdogKilledPath = ''
    $Script:OwnedExcelProcessId = 0
    Remove-StaleExcelWatchdogFiles
    if ($Script:ExcelRenderTimeoutSeconds -le 0) { return }
    $excelPid = Get-OwnedExcelProcessId $Excel
    if ($excelPid -le 0) { return }
    $Script:OwnedExcelProcessId = $excelPid
    $runDir = Join-Path ([IO.Path]::GetTempPath()) 'ReportBinderExcelWatchdog'
    try { New-Item -ItemType Directory -Path $runDir -Force | Out-Null } catch { return }
    $runId = [Guid]::NewGuid().ToString('N')
    $Script:ExcelWatchdogHeartbeatPath = Join-Path $runDir ("beat-$runId.txt")
    $Script:ExcelWatchdogKilledPath = Join-Path $runDir ("killed-$runId.txt")
    Update-ExcelRenderHeartbeat
    $beatEscaped = $Script:ExcelWatchdogHeartbeatPath.Replace("'", "''")
    $killedEscaped = $Script:ExcelWatchdogKilledPath.Replace("'", "''")
    $limit = [int]$Script:ExcelRenderTimeoutSeconds
    # 心拍ファイルが消えたら仕事が終わった合図。監視も終える。
    $command = @"
`$ErrorActionPreference = 'SilentlyContinue'
while (`$true) {
    Start-Sleep -Seconds 2
    if (-not (Test-Path -LiteralPath '$beatEscaped')) { break }
    `$raw = ''
    try { `$raw = [IO.File]::ReadAllText('$beatEscaped') } catch { continue }
    # "<unix秒> <その作業に許す秒数>"。猶予は作業ごとに変わる(複数シートの一括書き出しなど)。
    `$parts = @(`$raw.Trim() -split '\s+')
    [long]`$last = 0
    if (`$parts.Count -lt 1 -or -not [long]::TryParse(`$parts[0], [ref]`$last)) { continue }
    [long]`$allowance = $limit
    if (`$parts.Count -ge 2) { [void][long]::TryParse(`$parts[1], [ref]`$allowance) }
    if (`$allowance -lt 1) { `$allowance = $limit }
    if (([DateTimeOffset]::UtcNow.ToUnixTimeSeconds() - `$last) -lt `$allowance) { continue }
    `$target = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
    if (`$null -eq `$target -or `$target.ProcessName -ne 'EXCEL') { break }
    try { [IO.File]::WriteAllText('$killedEscaped', 'timeout') } catch { }
    Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
    break
}
"@
    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
    $Script:ExcelWatchdogLogPaths = @((Join-Path $runDir "out-$runId.log"), (Join-Path $runDir "err-$runId.log"))
    try {
        $Script:ExcelWatchdogProcess = Start-HiddenPowerShellChild $psExe $command $Script:ExcelWatchdogLogPaths[0] $Script:ExcelWatchdogLogPaths[1]
    } catch { $Script:ExcelWatchdogProcess = $null }
}

function Test-ExcelRenderWatchdogFired {
    # 監視を止めた後でも判定できるようにする。失敗の文言を組み立てるのは、
    # 後片付けが終わってからのことがあるため。
    if ($Script:ExcelWatchdogFiredSticky) { return $true }
    if ([string]::IsNullOrWhiteSpace($Script:ExcelWatchdogKilledPath)) { return $false }
    return (Test-Path -LiteralPath $Script:ExcelWatchdogKilledPath)
}

function Stop-ExcelRenderWatchdog {
    if (Test-ExcelRenderWatchdogFired) { $Script:ExcelWatchdogFiredSticky = $true }
    if (-not [string]::IsNullOrWhiteSpace($Script:ExcelWatchdogHeartbeatPath)) {
        Remove-Item -LiteralPath $Script:ExcelWatchdogHeartbeatPath -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $Script:ExcelWatchdogProcess) {
        try { if (-not $Script:ExcelWatchdogProcess.WaitForExit(3000)) { Stop-ReportBinderProcessTree ([int]$Script:ExcelWatchdogProcess.Id) } } catch { }
    }
    if (-not [string]::IsNullOrWhiteSpace($Script:ExcelWatchdogKilledPath)) {
        Remove-Item -LiteralPath $Script:ExcelWatchdogKilledPath -Force -ErrorAction SilentlyContinue
    }
    # 監視プロセスの標準出力・標準エラーのファイルも消す。Excelを起動するたび2つ増える。
    foreach ($logPath in @($Script:ExcelWatchdogLogPaths)) {
        if (-not [string]::IsNullOrWhiteSpace($logPath)) { Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue }
    }
    $Script:ExcelWatchdogLogPaths = @()
    $Script:ExcelWatchdogProcess = $null
    $Script:ExcelWatchdogHeartbeatPath = ''
    $Script:ExcelWatchdogKilledPath = ''
    $Script:OwnedExcelProcessId = 0
}

function New-ExcelApplicationForRender {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    try { $excel.AskToUpdateLinks = $false } catch { }
    try { $excel.AutomationSecurity = 3 } catch { }
    try { $excel.CalculateBeforeSave = $false } catch { }
    # Speed-oriented settings. Submitted workbooks are expected to be saved with calculated values.
    # PDF rendering does not need UI animation, status bar updates, or automatic recalculation.
    try { $excel.DisplayStatusBar = $false } catch { }
    try { $excel.EnableAnimations = $false } catch { }
    try { $excel.UserControl = $false } catch { }
    try { $excel.Calculation = -4135 } catch { } # xlCalculationManual
    Start-ExcelRenderWatchdog $excel
    return $excel
}

function Close-ExcelApplicationForRender($Excel) {
    # Quit() も同期COM呼び出しで、確認ダイアログや保留イベントで固まりうる。
    # 監視を先に止めると、いちばん固まりやすい所だけ無防備になる。閉じ終えてから止める。
    if ($Excel) {
        Update-ExcelRenderHeartbeat
        try { $Excel.Quit() } catch { }
        Invoke-ComRelease $Excel
    }
    Stop-ExcelRenderWatchdog
}

function Get-RenderEnvironment($Excel) {
    $fontArial = Test-Path -LiteralPath 'C:\Windows\Fonts\arial.ttf'
    $fontMsgothic = (Test-Path -LiteralPath 'C:\Windows\Fonts\msgothic.ttc') -or (Test-Path -LiteralPath 'C:\Windows\Fonts\msgothic.ttf')
    $activePrinter = ''
    try { $activePrinter = [string]$Excel.ActivePrinter } catch { }
    return [ordered]@{
        pcName = $env:COMPUTERNAME
        userName = "$env:USERDOMAIN\$env:USERNAME"
        excelVersion = [string]$Excel.Version
        osVersion = [Environment]::OSVersion.VersionString
        activePrinter = $activePrinter
        hasArial = $fontArial
        hasMsgothic = $fontMsgothic
        capturedAt = New-NowIso
    }
}

function Get-CachedRenderEnvironment($Excel) {
    if ($null -eq $Script:RenderEnvironmentCache) {
        $Script:RenderEnvironmentCache = Get-RenderEnvironment $Excel
    }
    return $Script:RenderEnvironmentCache
}

function Get-RenderEnvironmentFingerprint($EnvInfo) {
    # V5-§6.4: 画像ハッシュの比較可否を左右する要素だけを指紋にする。
    # pcName / userName は診断情報であり、比較の主判定には含めない
    # (別PCでも環境が同等なら同じPDFになりうるため)。
    if ($null -eq $EnvInfo) { return '' }
    $parts = @(
        'excelVersion=' + [string](Get-DataProperty $EnvInfo 'excelVersion' '')
        'osVersion='    + [string](Get-DataProperty $EnvInfo 'osVersion' '')
        'printer='      + [string](Get-DataProperty $EnvInfo 'activePrinter' '')
        'arial='        + [string](Get-DataProperty $EnvInfo 'hasArial' $false)
        'msgothic='     + [string](Get-DataProperty $EnvInfo 'hasMsgothic' $false)
        'printProfile=' + [string]$Script:ExcelPrintProfileVersion
    )
    # V5-P1: 画像ハッシュは PDFBox / Java / 解析方式 / DPI / 色にも依存する。
    # これらを含めないと、方式を変えても指紋が一致して互換性のないハッシュを直接比較してしまう。
    try {
        $vp = Get-VisualHashProfile
        $parts += ('pdfBox=' + [string](Get-DataProperty $vp 'pdfBoxVersion' ''))
        $parts += ('dpi=' + [string](Get-DataProperty $vp 'dpi' ''))
        $parts += ('color=' + [string](Get-DataProperty $vp 'colorMode' ''))
        $parts += ('profile=' + [string](Get-DataProperty $vp 'profileVersion' ''))
        $parts += ('analyzer=' + [string]$Script:PdfPageAnalyzerVersion)
        # V5-P1(#11): Java の版を実際に採取して指紋へ含める(空のままだと Java 更新後も指紋が変わらず、
        # 互換性のない画像ハッシュを直接比較してしまう)。
        $parts += ('java=' + [string](Get-JavaRuntimeSignature))
    } catch { }
    return (Get-Sha256Text ($parts -join '|'))
}

function Reset-RenderEnvironmentForJob {
    # V5-§6.4: 環境はサーバー稼働中に変わりうる(通常使うプリンタの変更など)。
    # レンダリングジョブの開始ごとに取り直す。ログ出力の1回制限とは分ける。
    $Script:RenderEnvironmentCache = $null
    $Script:RenderEnvironmentCompared = $false
    $Script:CurrentRenderEnvFingerprint = ''
}

function Compare-And-SaveEnvironment([string]$Language, $EnvInfo) {
    $workspace = Get-WorkspacePath $Language
    $path = Join-Path $workspace 'state\render-env.json'
    $warnings = @()
    if (Test-Path -LiteralPath $path) {
        $old = Read-JsonFile $path $null
        foreach ($key in @('pcName','excelVersion','osVersion','activePrinter','hasArial','hasMsgothic')) {
            if ([string]$old.$key -ne [string]$EnvInfo.$key) {
                $warnings += "PDF化環境が前回と異なります: $key 前回=[$($old.$key)] 今回=[$($EnvInfo.$key)]"
            }
        }
    }
    Write-JsonFile $path $EnvInfo
    return $warnings
}


function Touch-ClientActivity([string]$ClientId) {
    $Script:ClientAttached = $true
    $Script:LastHeartbeatUtc = [DateTime]::UtcNow
    $Script:ClientCloseNotifiedUtc = [DateTime]::MinValue
    return [ordered]@{ ok = $true; clientId = $ClientId; at = New-NowIso }
}

function Notify-ClientClosing([string]$ClientId) {
    $Script:ClientAttached = $true
    $Script:ClientCloseNotifiedUtc = [DateTime]::UtcNow
    return [ordered]@{ ok = $true; clientId = $ClientId; closing = $true; at = New-NowIso }
}

function Request-ServerShutdown([string]$Reason, [string]$Language) {
    if ($Script:AutoRenderInProgress -or (Test-ActiveRenderJobs $Language)) {
        throw 'PDF作成中はReportBinderを終了できません。処理が完了するか、PDF作成を中止してから終了してください。'
    }
    $Script:ShutdownRequested = $true
    return [ordered]@{ ok = $true; shutdown = $true; reason = $Reason; at = New-NowIso }
}

function Test-ActiveRenderJobs([string]$Language) {
    try {
        $dir = Get-RenderJobDir $Language
        if (-not (Test-Path -LiteralPath $dir)) { return $false }
        $cutoff = [DateTime]::UtcNow.AddHours(-12)
        $terminal = @('completed','completed-with-errors','failed','missing','cancelled')
        $now = [DateTime]::UtcNow
        foreach ($file in @(Get-ChildItem -LiteralPath $dir -Filter '*.status.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 30)) {
            if ($file.LastWriteTimeUtc -lt $cutoff) { continue }
            $job = Read-JsonFile $file.FullName $null
            if ($null -eq $job) { continue }
            $status = ([string]$job.status).ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($status)) { continue }
            if ($terminal -contains $status) { continue }
            $processId = Get-IntDataProperty $job 'processId' 0
            if ($processId -gt 0) {
                try { if ($null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)) { return $true } } catch { }
                continue
            }
            if (($now - $file.LastWriteTimeUtc).TotalSeconds -lt 20) { return $true }
        }
    } catch {
        return $false
    }
    return $false
}

function Try-AcquireLockHandle([string]$LockPath) {
    # V5: Invoke-WithLock は body 終了でハンドルを閉じるため、tick をまたいでロックを保持できない。
    # 自動スケジューラーが複数ブックの所有権を持ち続けるための取得専用ヘルパ。
    # 取得できない場合は $null を返す(エラーにしない。他サーバーが担当しているだけ)。
    try {
        $parent = Split-Path -Parent $LockPath
        if ($parent -and -not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $fs = [IO.File]::Open($LockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        try {
            $fs.SetLength(0)
            $bytes = [Text.Encoding]::UTF8.GetBytes("$env:COMPUTERNAME\$env:USERNAME $(New-NowIso)")
            $fs.Write($bytes, 0, $bytes.Length)
            $fs.Flush()
        } catch { }
        return [pscustomobject]@{ Path = $LockPath; Stream = $fs }
    } catch {
        return $null
    }
}

function Release-LockHandle($Handle) {
    if ($null -eq $Handle) { return }
    try { if ($Handle.Stream) { $Handle.Stream.Close(); $Handle.Stream.Dispose() } } catch { }
    try { if ($Handle.Path -and (Test-Path -LiteralPath $Handle.Path)) { Remove-Item -LiteralPath $Handle.Path -Force -ErrorAction SilentlyContinue } } catch { }
}

function Get-RenderEngineLockPath([string]$Language) {
    return (Join-Path (Get-WorkspacePath $Language) 'locks\render-engine.lock')
}

function Get-WorkbookRenderLockPath([string]$Language, [string]$WorkbookId) {
    return (Join-Path (Get-WorkspacePath $Language) ("locks\render_{0}.lock" -f $WorkbookId))
}

function Invoke-WithRenderLock([string]$Language, [string]$WorkbookId, [scriptblock]$Body) {
    # V5-D1: Excel COM は言語ごとに1ジョブへ制限する。
    # デッドロックを避けるため、取得順序を全経路で render-engine -> render_<workbookId> に統一する。
    $enginePath = Get-RenderEngineLockPath $Language
    $bookPath = Get-WorkbookRenderLockPath $Language $WorkbookId
    # Invoke-WithLock deliberately names its scriptblock parameter $Action.
    # If both functions use $Body, PowerShell's dynamic scope makes this wrapper
    # see itself and recursively reacquire the workbook lock.
    return Invoke-WithLock $enginePath {
        Invoke-WithLock $bookPath $Body
    }
}

function Invoke-WithLock([string]$LockPath, [scriptblock]$Action) {
    $parent = Split-Path -Parent $LockPath
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $fs = $null
    # OpenOrCreate + FileShare.None: 排他は「開いているハンドル」で担保する。
    # 以前の CreateNew 方式は、プロセス強制終了や電源断でロックファイルが残ると
    # 手動削除するまで永久に「他の処理がロック中」になっていた。
    # 強制終了直後はSMB側でハンドル解放が遅れることがあるため、短いリトライを入れる。
    for ($lockAttempt = 1; $lockAttempt -le 3; $lockAttempt++) {
        try {
            $fs = [IO.File]::Open($LockPath, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
            break
        } catch {
            $fs = $null
            if ($lockAttempt -lt 3) { Start-Sleep -Seconds 2 }
        }
    }
    if ($null -eq $fs) {
        throw "他の処理がロック中です。誰かが同じ処理を実行中か、強制終了したプロセスのハンドルが残っています。時間をおいて再実行してください。解放後に保持者を確認するには次のファイルを開いてください: $LockPath"
    }
    try {
        $fs.SetLength(0)
        $bytes = [Text.Encoding]::UTF8.GetBytes("$env:COMPUTERNAME\$env:USERNAME $(New-NowIso)")
        $fs.Write($bytes, 0, $bytes.Length)
        return & $Action
    } finally {
        if ($fs) { $fs.Close(); $fs.Dispose() }
        if (Test-Path -LiteralPath $LockPath) { Remove-Item -LiteralPath $LockPath -Force -ErrorAction SilentlyContinue }
    }
}

function Register-Workbook([string]$Language, [string]$RelativePath, [string]$Category = '', [string]$PackId = '') {
    Test-DirectExcelRelativePath $RelativePath | Out-Null
    $cat=Require-WorkbookCategory $Category
    $paths=Get-Paths
    $full=Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $candidate=New-WorkbookObject $RelativePath $Language $cat
    $resolvedPackId = if ([string]::IsNullOrWhiteSpace($PackId)) { Get-BuiltinPackId $cat } else { ([string]$PackId).Trim() }
    if ([string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId $resolvedPackId))) { Set-NoteProperty $candidate 'workbookId' ("$resolvedPackId-$([string]$candidate.workbookId)") }
    Set-NoteProperty $candidate 'packId' $resolvedPackId
    return Update-StructureLocked $Language {
        param($structure)
        $normalized=([string]$RelativePath -replace '\\','/').ToLowerInvariant()
        $existing=@(Get-Array $structure.workbooks | Where-Object {
            ([string]$_.workbookId -eq [string]$candidate.workbookId) -or ((([string]$_.relativePath -replace '\\','/').ToLowerInvariant() -eq $normalized) -and (Test-WorkbookPack $_ $resolvedPackId))
        } | Select-Object -First 1)
        if ($existing.Count -gt 0) {
            $old=$existing[0]
            Set-NoteProperty $candidate 'workbookId' ([string]$old.workbookId)
            foreach ($name in @('lastRenderedVersionId','lastRenderedExcelHash','lastRenderedAt','lastRenderedSheets','lastRenderedSheetFingerprint','lastRenderedSheetSelectionMode','excelSheetSelectionMode','lastRenderLog','renderProfileVersion','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash')) {
                Set-NoteProperty $candidate $name (Get-DataProperty $old $name $null)
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$candidate.lastRenderedExcelHash)) { Set-NoteProperty $candidate 'status' $(if ($candidate.currentExcelHash -ne $candidate.lastRenderedExcelHash) {'excel-updated'} else {[string]$old.status}) }
        }
        $structure.workbooks=@(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -ne [string]$candidate.workbookId }) + @($candidate)
        return [ordered]@{ workbook=$candidate; sheets=@(); registered=$true; inspected=$false }
    }
}

function Register-WorkbooksBatch([string]$Language, $RelativePaths, [string]$Category = '') {
    $categoryNormalized = Require-WorkbookCategory $Category
    return Register-SourcesBatch $Language $RelativePaths (Get-BuiltinPackId $categoryNormalized) 'excel'
}

function Register-PdfSource([string]$Language, [string]$RelativePath, [string]$Category = '', [string]$PackId = '') {
    Test-DirectPdfRelativePath $RelativePath | Out-Null
    $cat = Require-WorkbookCategory $Category
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $inspection = Inspect-PdfSourceFile $full
    $candidate = New-PdfSourceObject $RelativePath $Language $cat ([int]$inspection.pageCount)
    $resolvedPackId = if ([string]::IsNullOrWhiteSpace($PackId)) { Get-BuiltinPackId $cat } else { ([string]$PackId).Trim() }
    if ([string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId $resolvedPackId))) { Set-NoteProperty $candidate 'workbookId' ("$resolvedPackId-$([string]$candidate.workbookId)") }
    Set-NoteProperty $candidate 'packId' $resolvedPackId
    return Update-StructureLocked $Language {
        param($structure)
        $normalized = ([string]$RelativePath -replace '\\','/').ToLowerInvariant()
        $existing = @(Get-Array $structure.workbooks | Where-Object {
            ([string]$_.workbookId -eq [string]$candidate.workbookId) -or ((([string]$_.relativePath -replace '\\','/').ToLowerInvariant() -eq $normalized) -and (Test-WorkbookPack $_ $resolvedPackId))
        } | Select-Object -First 1)
        if ($existing.Count -gt 0) {
            $old = $existing[0]
            Set-NoteProperty $candidate 'workbookId' ([string]$old.workbookId)
            foreach ($name in @('lastRenderedVersionId','lastRenderedSnapshotId','lastRenderedExcelHash','lastRenderedAt','lastRenderedSheets','lastRenderedSheetFingerprint','lastRenderLog','renderProfileVersion','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash','ownerDepartment','required','defaultTargetId','requirementId')) {
                Set-NoteProperty $candidate $name (Get-DataProperty $old $name $null)
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$candidate.lastRenderedExcelHash)) {
                Set-NoteProperty $candidate 'status' $(if ($candidate.currentExcelHash -ne $candidate.lastRenderedExcelHash) { 'source-updated' } else { [string]$old.status })
            }
        }
        $structure.workbooks = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -ne [string]$candidate.workbookId }) + @($candidate)
        return [ordered]@{ workbook = $candidate; units = @($inspection.units); registered = $true; inspected = $true; sourceType = 'pdf' }
    }
}

function Register-WordSource([string]$Language, [string]$RelativePath, [string]$Category = '', [string]$PackId = '') {
    Test-DirectWordRelativePath $RelativePath | Out-Null
    $cat = Require-WorkbookCategory $Category
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $inspection = Inspect-WordSourceFile $full
    $candidate = New-WordSourceObject $RelativePath $Language $cat
    $resolvedPackId = if ([string]::IsNullOrWhiteSpace($PackId)) { Get-BuiltinPackId $cat } else { ([string]$PackId).Trim() }
    if ([string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId $resolvedPackId))) { Set-NoteProperty $candidate 'workbookId' ("$resolvedPackId-$([string]$candidate.workbookId)") }
    Set-NoteProperty $candidate 'packId' $resolvedPackId
    return Update-StructureLocked $Language {
        param($structure)
        $normalized = ([string]$RelativePath -replace '\\','/').ToLowerInvariant()
        $existing = @(Get-Array $structure.workbooks | Where-Object {
            ([string]$_.workbookId -eq [string]$candidate.workbookId) -or ((([string]$_.relativePath -replace '\\','/').ToLowerInvariant() -eq $normalized) -and (Test-WorkbookPack $_ $resolvedPackId))
        } | Select-Object -First 1)
        if ($existing.Count -gt 0) {
            $old = $existing[0]
            Set-NoteProperty $candidate 'workbookId' ([string]$old.workbookId)
            foreach ($name in @('lastRenderedVersionId','lastRenderedSnapshotId','lastRenderedExcelHash','lastRenderedAt','lastRenderedSheets','lastRenderedSheetFingerprint','lastRenderLog','renderProfileVersion','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash','ownerDepartment','required','defaultTargetId','requirementId','sourcePageCount')) {
                Set-NoteProperty $candidate $name (Get-DataProperty $old $name $null)
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$candidate.lastRenderedExcelHash)) {
                Set-NoteProperty $candidate 'status' $(if ($candidate.currentExcelHash -ne $candidate.lastRenderedExcelHash) { 'source-updated' } else { [string]$old.status })
            }
        }
        $structure.workbooks = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -ne [string]$candidate.workbookId }) + @($candidate)
        return [ordered]@{ workbook = $candidate; units = @($inspection.units); registered = $true; inspected = $true; sourceType = 'word' }
    }
}

function Stop-ReportBinderProcessTree([int]$ProcessId) {
    if ($ProcessId -le 0) { return }
    try {
        $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId=$ProcessId" -ErrorAction SilentlyContinue)
        foreach ($child in $children) { Stop-ReportBinderProcessTree ([int]$child.ProcessId) }
    } catch { }
    try { Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue } catch { }
}

function ConvertTo-PathBase64([string]$Path) {
    return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($Path)))
}

function Invoke-WordDocumentToPdf([string]$InputPath, [string]$OutputPath, [int]$TimeoutSeconds = 0) {
    if ($TimeoutSeconds -le 0) { $TimeoutSeconds = $Script:WordRenderTimeoutSeconds }
    $worker = Join-Path $Script:AppRoot 'tools\word-render-worker.ps1'
    if (-not (Test-Path -LiteralPath $worker -PathType Leaf)) { throw 'Word変換ワーカーが見つかりません。' }
    $runId = New-RbId
    $runDir = Split-Path -Parent $OutputPath
    if (-not (Test-Path -LiteralPath $runDir)) { New-Item -ItemType Directory -Path $runDir -Force | Out-Null }
    $resultPath = Join-Path $runDir "word-result-$runId.json"
    $stdoutPath = Join-Path $runDir "word-out-$runId.log"
    $stderrPath = Join-Path $runDir "word-err-$runId.log"
    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
    $workerEscaped = $worker.Replace("'", "''")
    $command = "& '$workerEscaped' -InputPathB64 '$(ConvertTo-PathBase64 $InputPath)' -OutputPathB64 '$(ConvertTo-PathBase64 $OutputPath)' -ResultPathB64 '$(ConvertTo-PathBase64 $resultPath)'"
    $proc = $null
    try {
        $proc = Start-HiddenPowerShellChild $psExe $command $stdoutPath $stderrPath
        if ($null -eq $proc) { throw 'WORD_WORKER_START_FAILED' }
        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            Stop-ReportBinderProcessTree ([int]$proc.Id)
            throw "WORD_TIMEOUT:$TimeoutSeconds"
        }
        try { $proc.WaitForExit() } catch { }
        $result = Read-JsonFile $resultPath $null
        if ($null -eq $result) {
            $details = @()
            if (Test-Path -LiteralPath $stderrPath) { $details += (Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue) }
            if (Test-Path -LiteralPath $stdoutPath) { $details += (Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue) }
            throw ('WORD_WORKER_NO_RESULT:' + ([string]::Join(' ', @($details))).Trim())
        }
        if (-not [bool](Get-DataProperty $result 'ok' $false)) {
            throw ("$([string](Get-DataProperty $result 'errorCode' 'WORD_WORKER_FAILED')):$([string](Get-DataProperty $result 'stage' '')):$([string](Get-DataProperty $result 'message' ''))")
        }
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf) -or (Get-Item -LiteralPath $OutputPath).Length -le 0 -or -not (Test-PdfMagic $OutputPath)) { throw 'WORD_PDF_INVALID' }
        return $result
    } finally {
        foreach ($path in @($resultPath,$stdoutPath,$stderrPath)) { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } }
    }
}

function Render-WordSource([string]$Language, [string]$SourceId, [scriptblock]$ProgressCallback = $null,
                           [string]$SourceOverridePath = '', [string]$SourceSnapshotId = '', [string]$ExpectedSourceHash = '') {
    return Invoke-WithRenderLock $Language $SourceId {
        $workspace = Get-WorkspacePath $Language
        $paths = Get-Paths
        $structure = Get-Structure $Language
        $found = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $SourceId } | Select-Object -First 1)
        if ($found.Count -eq 0) { throw "登録済みWord原稿が見つかりません: $SourceId" }
        $source = $found[0]
        if ([string](Get-DataProperty $source 'sourceType' '') -ne 'word') { throw 'Word原稿ではないためWordアダプターで処理できません。' }
        $livePath = Join-Safe ([string]$paths.submissionDir) ([string]$source.relativePath)
        if (-not (Test-Path -LiteralPath $livePath)) { throw "Word原稿が見つかりません: $($source.relativePath)" }
        $inputPath = $SourceOverridePath
        $ephemeralCaptureId = ''
        if ([string]::IsNullOrWhiteSpace($inputPath) -and (Test-InputHistoryEnabled)) {
            if ([string]::IsNullOrWhiteSpace($SourceSnapshotId)) {
                $captured = Capture-DetectedSnapshot $Language $SourceId 'render'
                if ([bool]$captured.ok) { $SourceSnapshotId = [string]$captured.snapshotId }
            }
            $inputInfo = Capture-RenderInput $Language $SourceId $SourceSnapshotId ''
            if (-not [string]::IsNullOrWhiteSpace([string]$inputInfo.path)) {
                $inputPath = [string]$inputInfo.path; $SourceSnapshotId = [string]$inputInfo.snapshotId
                if ([bool]$inputInfo.ephemeral) { $ephemeralCaptureId = [string]$inputInfo.captureId }
            }
        }
        if ([string]::IsNullOrWhiteSpace($inputPath)) { $inputPath = $livePath }
        [void](Inspect-WordSourceFile $inputPath)
        $sourceHash = Normalize-FileHash (New-Sha256 $inputPath)
        $expectedHash = Normalize-FileHash $ExpectedSourceHash
        if (-not [string]::IsNullOrWhiteSpace($expectedHash) -and $sourceHash -ne $expectedHash) { throw '指定したWord原稿の版と実ファイルが一致しません。更新確認後に再度変換してください。' }
        $versionId = New-RbVersionId
        $contentDir = Get-ContentPdfVersionDir $workspace $SourceId $versionId
        $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-word-' + (New-RbId))
        $success = $false
        try {
            New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
            New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
            $workDocx = Join-Path $tmpDir 'source.docx'
            Copy-FileSharedRead $inputPath $workDocx
            $wholePdf = Join-Path $tmpDir 'document.pdf'
            if ($ProgressCallback) { & $ProgressCallback 'word' $SourceId '' }
            $wordResult = Invoke-WordDocumentToPdf $workDocx $wholePdf $Script:WordRenderTimeoutSeconds
            if ($ProgressCallback) { & $ProgressCallback 'split' $SourceId '' }
            $split = Invoke-PdfSourceSplit $wholePdf $contentDir
            $rendered = @(); $inspection = @()
            foreach ($pageInfo in @($split.pages)) {
                $number = [int]$pageInfo.pageNumber; $name = "Page $number"
                $title = "$([IO.Path]::GetFileNameWithoutExtension([string]$source.fileName)) / $name"
                $inspection += [pscustomobject][ordered]@{ sheetName = $name; sheetIndex = $number; detectedTitle = $title }
                $rendered += [pscustomobject][ordered]@{ sheetName = $name; pdf = [string]$pageInfo.pdf; pageCount = 1; warnings = @() }
            }
            $pageSync = Update-StructureLocked $Language {
                param($st)
                $latest = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq $SourceId } | Select-Object -First 1)
                if ($latest.Count -eq 0) { throw "登録済みWord原稿が見つかりません: $SourceId" }
                $w = $latest[0]; $sync = Update-WorkbookPagesFromInspection $Language $st $w $inspection; $pages = @(Get-Array $st.pages)
                foreach ($r in $rendered) {
                    $page = @($pages | Where-Object { [string]$_.workbookId -eq $SourceId -and [string]$_.sheetName -eq [string]$r.sheetName } | Select-Object -First 1)
                    if ($page.Count -gt 0) {
                        Set-NoteProperty $page[0] 'contentPdf' (Get-RelativePathCompat $workspace ([string]$r.pdf))
                        Set-NoteProperty $page[0] 'status' 'rendered'; Set-NoteProperty $page[0] 'warnings' @(); Set-NoteProperty $page[0] 'updatedAt' (New-NowIso)
                    }
                }
                Set-NoteProperty $w 'sourceType' 'word'; Set-NoteProperty $w 'adapterId' 'word-com-v1'; Set-NoteProperty $w 'sourcePageCount' ([int]$split.pageCount)
                Set-NoteProperty $w 'lastRenderedVersionId' $versionId; Set-NoteProperty $w 'lastRenderedSnapshotId' $SourceSnapshotId
                Set-NoteProperty $w 'lastRenderedExcelHash' $sourceHash; Set-NoteProperty $w 'lastRenderedAt' (New-NowIso)
                Set-NoteProperty $w 'renderProfileVersion' $Script:WordRenderProfileVersion; Set-WorkbookRenderedSheetSnapshot $w (Get-DataProperty $sync 'sheetNames' @())
                Set-NoteProperty $w 'lastError' ''; Set-NoteProperty $w 'lastErrorUser' ''; Set-NoteProperty $w 'lastErrorAt' $null; Set-NoteProperty $w 'lastRenderAttemptHash' $sourceHash
                $currentHash = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
                if ([string]::IsNullOrWhiteSpace($currentHash) -or $currentHash -eq $sourceHash) { Set-NoteProperty $w 'status' 'rendered-unchecked' }
                else { Set-NoteProperty $w 'status' 'source-updated'; foreach ($p in @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $SourceId })) { Set-NoteProperty $p 'status' 'stale' } }
                $cat = Get-WorkbookPackId $w
                $vols = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $SourceId } | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
                Mark-VolumeNeedsRebuild $st $Language $cat $vols 'render' 'Word原稿を1件PDF変換しました'
                return $sync
            }
            $wordVersion = [string](Get-DataProperty $wordResult 'wordVersion' '')
            $Script:CurrentRenderEnvFingerprint = Get-Sha256Text "word-com|$wordVersion|$($Script:WordRenderProfileVersion)"
            $Script:CurrentRenderEnvInfo = [ordered]@{ adapterId = 'word-com-v1'; wordVersion = $wordVersion; profileVersion = $Script:WordRenderProfileVersion }
            $Script:PendingAnalysis = [ordered]@{ language = $Language; workbookId = $SourceId; snapshotId = [string]$SourceSnapshotId; versionId = $versionId; rendered = $rendered }
            Remove-WorkbookContentPdfs $workspace $SourceId $versionId
            $success = $true
            return [ordered]@{ workbookId = $SourceId; sourceId = $SourceId; sourceType = 'word'; adapterId = 'word-com-v1'; versionId = $versionId; rendered = @($rendered); warnings = @(); steps = @('Wordを一時コピーからPDF変換','PDFをページ単位で取り込み'); sheetSync = $pageSync }
        } finally {
            if (-not $success -and (Test-Path -LiteralPath $contentDir)) { Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (-not [string]::IsNullOrWhiteSpace($ephemeralCaptureId)) { Remove-EphemeralCopy $Language $ephemeralCaptureId }
        }
    }
}

function Render-PdfSource([string]$Language, [string]$SourceId, [scriptblock]$ProgressCallback = $null,
                          [string]$SourceOverridePath = '', [string]$SourceSnapshotId = '', [string]$ExpectedSourceHash = '') {
    return Invoke-WithRenderLock $Language $SourceId {
        $workspace = Get-WorkspacePath $Language
        $paths = Get-Paths
        $structure = Get-Structure $Language
        $found = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $SourceId } | Select-Object -First 1)
        if ($found.Count -eq 0) { throw "登録済みPDF原稿が見つかりません: $SourceId" }
        $source = $found[0]
        if ([string](Get-DataProperty $source 'sourceType' '') -ne 'pdf') { throw 'PDF原稿ではないためPDFアダプターで処理できません。' }
        $livePath = Join-Safe ([string]$paths.submissionDir) ([string]$source.relativePath)
        if (-not (Test-Path -LiteralPath $livePath)) { throw "PDF原稿が見つかりません: $($source.relativePath)" }
        $inputPath = $SourceOverridePath
        $ephemeralCaptureId = ''
        if ([string]::IsNullOrWhiteSpace($inputPath) -and (Test-InputHistoryEnabled)) {
            if ([string]::IsNullOrWhiteSpace($SourceSnapshotId)) {
                $captured = Capture-DetectedSnapshot $Language $SourceId 'render'
                if ([bool]$captured.ok) { $SourceSnapshotId = [string]$captured.snapshotId }
            }
            $inputInfo = Capture-RenderInput $Language $SourceId $SourceSnapshotId ''
            if (-not [string]::IsNullOrWhiteSpace([string]$inputInfo.path)) {
                $inputPath = [string]$inputInfo.path
                $SourceSnapshotId = [string]$inputInfo.snapshotId
                if ([bool]$inputInfo.ephemeral) { $ephemeralCaptureId = [string]$inputInfo.captureId }
            }
        }
        if ([string]::IsNullOrWhiteSpace($inputPath)) { $inputPath = $livePath }
        $sourceHash = Normalize-FileHash (New-Sha256 $inputPath)
        $expectedHash = Normalize-FileHash $ExpectedSourceHash
        if (-not [string]::IsNullOrWhiteSpace($expectedHash) -and $sourceHash -ne $expectedHash) { throw '指定したPDF原稿の版と実ファイルが一致しません。更新確認後に再度取り込んでください。' }
        $versionId = New-RbVersionId
        $contentDir = Get-ContentPdfVersionDir $workspace $SourceId $versionId
        $success = $false
        try {
            New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
            if ($ProgressCallback) { & $ProgressCallback 'split' $SourceId '' }
            $split = Invoke-PdfSourceSplit $inputPath $contentDir
            $rendered = @(); $inspection = @()
            foreach ($pageInfo in @($split.pages)) {
                $number = [int]$pageInfo.pageNumber
                $name = "Page $number"
                $title = "$([IO.Path]::GetFileNameWithoutExtension([string]$source.fileName)) / $name"
                $inspection += [pscustomobject][ordered]@{ sheetName = $name; sheetIndex = $number; detectedTitle = $title }
                $rendered += [pscustomobject][ordered]@{ sheetName = $name; pdf = [string]$pageInfo.pdf; pageCount = 1; warnings = @() }
            }
            $pageSync = Update-StructureLocked $Language {
                param($st)
                $latest = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq $SourceId } | Select-Object -First 1)
                if ($latest.Count -eq 0) { throw "登録済みPDF原稿が見つかりません: $SourceId" }
                $w = $latest[0]
                $sync = Update-WorkbookPagesFromInspection $Language $st $w $inspection
                $pages = @(Get-Array $st.pages)
                foreach ($r in $rendered) {
                    $page = @($pages | Where-Object { [string]$_.workbookId -eq $SourceId -and [string]$_.sheetName -eq [string]$r.sheetName } | Select-Object -First 1)
                    if ($page.Count -gt 0) {
                        Set-NoteProperty $page[0] 'contentPdf' (Get-RelativePathCompat $workspace ([string]$r.pdf))
                        Set-NoteProperty $page[0] 'status' 'rendered'; Set-NoteProperty $page[0] 'warnings' @(); Set-NoteProperty $page[0] 'updatedAt' (New-NowIso)
                    }
                }
                Set-NoteProperty $w 'sourceType' 'pdf'; Set-NoteProperty $w 'adapterId' 'pdfbox-import-v1'; Set-NoteProperty $w 'sourcePageCount' ([int]$split.pageCount)
                Set-NoteProperty $w 'lastRenderedVersionId' $versionId; Set-NoteProperty $w 'lastRenderedSnapshotId' $SourceSnapshotId
                Set-NoteProperty $w 'lastRenderedExcelHash' $sourceHash; Set-NoteProperty $w 'lastRenderedAt' (New-NowIso)
                Set-NoteProperty $w 'renderProfileVersion' $Script:PdfImportProfileVersion; Set-WorkbookRenderedSheetSnapshot $w (Get-DataProperty $sync 'sheetNames' @())
                Set-NoteProperty $w 'lastError' ''; Set-NoteProperty $w 'lastErrorUser' ''; Set-NoteProperty $w 'lastErrorAt' $null; Set-NoteProperty $w 'lastRenderAttemptHash' $sourceHash
                $currentHash = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
                if ([string]::IsNullOrWhiteSpace($currentHash) -or $currentHash -eq $sourceHash) { Set-NoteProperty $w 'status' 'rendered-unchecked' }
                else { Set-NoteProperty $w 'status' 'source-updated'; foreach ($p in @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $SourceId })) { Set-NoteProperty $p 'status' 'stale' } }
                $cat = Get-WorkbookPackId $w
                $vols = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $SourceId } | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
                Mark-VolumeNeedsRebuild $st $Language $cat $vols 'render' 'PDF原稿を1件取り込みました'
                return $sync
            }
            $Script:CurrentRenderEnvFingerprint = Get-Sha256Text "pdfbox-import|$($Script:PdfImportProfileVersion)"
            $Script:CurrentRenderEnvInfo = [ordered]@{ adapterId = 'pdfbox-import-v1'; profileVersion = $Script:PdfImportProfileVersion }
            $Script:PendingAnalysis = [ordered]@{ language = $Language; workbookId = $SourceId; snapshotId = [string]$SourceSnapshotId; versionId = $versionId; rendered = $rendered }
            Remove-WorkbookContentPdfs $workspace $SourceId $versionId
            $success = $true
            return [ordered]@{ workbookId = $SourceId; sourceId = $SourceId; sourceType = 'pdf'; adapterId = 'pdfbox-import-v1'; versionId = $versionId; rendered = @($rendered); warnings = @(); steps = @('PDFをページ単位で取り込み'); sheetSync = $pageSync }
        } finally {
            if (-not $success -and (Test-Path -LiteralPath $contentDir)) { Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (-not [string]::IsNullOrWhiteSpace($ephemeralCaptureId)) { Remove-EphemeralCopy $Language $ephemeralCaptureId }
        }
    }
}

function Render-PdfSnapshotForComparison([string]$Language, [string]$SourceId, [string]$SnapshotId) {
    $state = Get-SnapshotSourceState $Language $SourceId $SnapshotId
    if (-not [bool]$state.sourceRetained) { return [ordered]@{ ok = $false; reason = 'source-missing' } }
    return Invoke-WithRenderLock $Language $SourceId {
        $versionId = New-RbVersionId
        $workspace = Get-WorkspacePath $Language
        $contentDir = Get-ContentPdfVersionDir $workspace $SourceId $versionId
        $jobId = New-RbId; $snapshotLease = ''; $contentLease = ''; $success = $false
        try {
            New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
            $snapshotLease = New-SnapshotLease $Language $SourceId $SnapshotId 'compare' $jobId 120 $versionId
            $contentLease = New-ContentPdfLease $workspace $SourceId $versionId 'compare' $jobId 120
            if ([string]::IsNullOrWhiteSpace($snapshotLease) -or [string]::IsNullOrWhiteSpace($contentLease)) { throw '比較資産の保護leaseを作成できませんでした。' }
            $split = Invoke-PdfSourceSplit ([string]$state.sourcePath) $contentDir
            $sheets = @()
            foreach ($pageInfo in @($split.pages)) {
                $name = "Page $([int]$pageInfo.pageNumber)"
                $sheets += @{ sheetName = $name; pdf = [string]$pageInfo.pdf; rasterDirectory = (Get-RenderRasterSheetDir $Language $SourceId $SnapshotId $versionId $name) }
            }
            $Script:CurrentRenderEnvFingerprint = Get-Sha256Text "pdfbox-import|$($Script:PdfImportProfileVersion)"
            $analysis = Invoke-PdfPageAnalyzer $sheets
            if ($null -eq $analysis -or -not [bool]$analysis.ok) { return [ordered]@{ ok = $false; reason = 'analyze-failed' } }
            Write-RenderRecord $Language $SourceId $SnapshotId $versionId 'comparison' $true ([pscustomobject]$analysis.result)
            $availability = Get-HistoryRenderVersionAvailability $Language $SourceId $SnapshotId $versionId
            if (-not [bool]$availability.ready) { return [ordered]@{ ok = $false; reason = 'retention-verify-failed'; message = [string]$availability.reason } }
            $success = $true
            return [ordered]@{ ok = $true; versionId = $versionId; envFingerprint = [string]$Script:CurrentRenderEnvFingerprint }
        } catch { return [ordered]@{ ok = $false; reason = 'error'; message = $_.Exception.Message } }
        finally {
            Remove-SnapshotLease $Language $SourceId $SnapshotId $snapshotLease
            Remove-ContentPdfLease $workspace $SourceId $versionId $contentLease
            if (-not $success -and (Test-Path -LiteralPath $contentDir)) { Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Render-WordSnapshotForComparison([string]$Language, [string]$SourceId, [string]$SnapshotId) {
    $state = Get-SnapshotSourceState $Language $SourceId $SnapshotId
    if (-not [bool]$state.sourceRetained) { return [ordered]@{ ok = $false; reason = 'source-missing' } }
    return Invoke-WithRenderLock $Language $SourceId {
        $versionId = New-RbVersionId
        $workspace = Get-WorkspacePath $Language
        $contentDir = Get-ContentPdfVersionDir $workspace $SourceId $versionId
        $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-word-compare-' + (New-RbId))
        $jobId = New-RbId; $snapshotLease = ''; $contentLease = ''; $success = $false
        try {
            New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
            New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
            $snapshotLease = New-SnapshotLease $Language $SourceId $SnapshotId 'compare' $jobId 180 $versionId
            $contentLease = New-ContentPdfLease $workspace $SourceId $versionId 'compare' $jobId 180
            if ([string]::IsNullOrWhiteSpace($snapshotLease) -or [string]::IsNullOrWhiteSpace($contentLease)) { throw '比較資産の保護leaseを作成できませんでした。' }
            $workDocx = Join-Path $tmpDir 'source.docx'; Copy-FileSharedRead ([string]$state.sourcePath) $workDocx
            $wholePdf = Join-Path $tmpDir 'document.pdf'
            $wordResult = Invoke-WordDocumentToPdf $workDocx $wholePdf $Script:WordRenderTimeoutSeconds
            $split = Invoke-PdfSourceSplit $wholePdf $contentDir
            $sheets = @()
            foreach ($pageInfo in @($split.pages)) {
                $name = "Page $([int]$pageInfo.pageNumber)"
                $sheets += @{ sheetName = $name; pdf = [string]$pageInfo.pdf; rasterDirectory = (Get-RenderRasterSheetDir $Language $SourceId $SnapshotId $versionId $name) }
            }
            $wordVersion = [string](Get-DataProperty $wordResult 'wordVersion' '')
            $Script:CurrentRenderEnvFingerprint = Get-Sha256Text "word-com|$wordVersion|$($Script:WordRenderProfileVersion)"
            $analysis = Invoke-PdfPageAnalyzer $sheets
            if ($null -eq $analysis -or -not [bool]$analysis.ok) { return [ordered]@{ ok = $false; reason = 'analyze-failed' } }
            Write-RenderRecord $Language $SourceId $SnapshotId $versionId 'comparison' $true ([pscustomobject]$analysis.result)
            $availability = Get-HistoryRenderVersionAvailability $Language $SourceId $SnapshotId $versionId
            if (-not [bool]$availability.ready) { return [ordered]@{ ok = $false; reason = 'retention-verify-failed'; message = [string]$availability.reason } }
            $success = $true
            return [ordered]@{ ok = $true; versionId = $versionId; envFingerprint = [string]$Script:CurrentRenderEnvFingerprint }
        } catch { return [ordered]@{ ok = $false; reason = 'error'; message = $_.Exception.Message } }
        finally {
            Remove-SnapshotLease $Language $SourceId $SnapshotId $snapshotLease
            Remove-ContentPdfLease $workspace $SourceId $versionId $contentLease
            if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (-not $success -and (Test-Path -LiteralPath $contentDir)) { Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}


function Get-LastRenderAttemptFor([string]$WorkbookId) {
    $a = $Script:LastRenderAttempt
    if ($null -eq $a) { return [ordered]@{ snapshotId = ''; hash = '' } }
    if ([string](Get-DataProperty $a 'workbookId' '') -ne [string]$WorkbookId) { return [ordered]@{ snapshotId = ''; hash = '' } }
    return [ordered]@{ snapshotId = [string](Get-DataProperty $a 'snapshotId' ''); hash = [string](Get-DataProperty $a 'hash' '') }
}

function Set-WorkbookRenderError([string]$Language, [string]$WorkbookId, [string]$Message, [string]$Detail = '',
                                 [string]$AttemptedSnapshotId = '', [string]$AttemptedHash = '') {
    try {
        $userMessage=ConvertTo-UserRenderError $Message
        $noNumericTarget = ([string]$Message -match '半角数字だけの表示シートがありません')
        Update-StructureLocked $Language {
            param($structure)
            $wb=@(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
            if ($wb.Count -gt 0) {
                # V5-§6.3a: 失敗したのは「試行した版」であって、現在の版とは限らない。
                # 既に新しい版が検知されているなら render-error にせず excel-updated のままにする。
                $attempted = Normalize-FileHash $AttemptedHash
                $latest = Normalize-FileHash ([string](Get-DataProperty $wb[0] 'currentExcelHash' ''))
                $sameVersion = ([string]::IsNullOrWhiteSpace($attempted)) -or ([string]::IsNullOrWhiteSpace($latest)) -or ($attempted -eq $latest)
                if ($sameVersion) {
                    Set-NoteProperty $wb[0] 'status' 'render-error'
                    if (-not $noNumericTarget) {
                        foreach ($p in @(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId -and [string]$_.status -ne 'confirmed' })) { Set-NoteProperty $p 'status' 'render-error'; Set-NoteProperty $p 'warnings' @($userMessage); Set-NoteProperty $p 'updatedAt' (New-NowIso) }
                    }
                } else {
                    Set-NoteProperty $wb[0] 'status' (Get-SourceUpdatedStatus $wb[0])
                    # ページの status は上書きしない(新しい版の状態を壊さないため)
                }
                # 失敗情報は状態と切り離して必ず残す
                Add-NotePropertyIfMissing $wb[0] 'lastRenderErrorSnapshotId' ''
                Add-NotePropertyIfMissing $wb[0] 'lastRenderErrorHash' ''
                Set-NoteProperty $wb[0] 'lastRenderErrorSnapshotId' ([string]$AttemptedSnapshotId)
                Set-NoteProperty $wb[0] 'lastRenderErrorHash' ([string]$AttemptedHash)
                Set-NoteProperty $wb[0] 'lastError' $Message; Set-NoteProperty $wb[0] 'lastErrorUser' $userMessage; Set-NoteProperty $wb[0] 'lastErrorAt' (New-NowIso)
            }
        } | Out-Null
        $workspace=Get-WorkspacePath $Language
        $safeId=[regex]::Replace($WorkbookId,'[^A-Za-z0-9_.-]+','_')
        $logRel=Join-Path 'logs' ("render-error_{0}_{1}.json" -f $safeId,(New-RbId))
        Write-JsonFile (Join-Path $workspace $logRel) ([ordered]@{workbookId=$WorkbookId;message=$Message;userMessage=$userMessage;detail=$Detail;at=New-NowIso})
        Update-StructureLocked $Language { param($structure) $wb=@(Get-Array $structure.workbooks|Where-Object{[string]$_.workbookId -eq $WorkbookId}|Select-Object -First 1); if($wb.Count){Set-NoteProperty $wb[0] 'lastRenderLog' $logRel} } | Out-Null
    } catch { }
}

function Get-ContentPdfMaintenanceLockPath([string]$Workspace, [string]$WorkbookId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    return (Join-Path $Workspace ("locks\content-pdf_{0}.lock" -f $safeWorkbookId))
}

function Remove-WorkbookContentPdfs([string]$Workspace, [string]$WorkbookId, [string]$KeepVersionId) {
    $lockPath = Get-ContentPdfMaintenanceLockPath $Workspace $WorkbookId
    try {
        Invoke-WithLock $lockPath { Remove-WorkbookContentPdfsCore $Workspace $WorkbookId $KeepVersionId } | Out-Null
    } catch {
        # 保持整理の競合・失敗で、完成済みPDF作成そのものを失敗扱いにしない。
        Write-Warning ('content PDFの世代整理を見送りました: ' + $_.Exception.Message)
    }
}

function Remove-WorkbookContentPdfsCore([string]$Workspace, [string]$WorkbookId, [string]$KeepVersionId) {
    # V5-§3.7: 保持世代数と pins/leases による保護を尊重する。
    # 単純に KeepVersionId 以外を全削除すると、差分比較の基準や正式PDFが参照する世代まで消える。
    try {
        # V5-P0: 承認前はディスク使用量を増やさない。V4.1 と同じ「最新以外を削除」に戻す。
        if (-not (Test-InputHistoryEnabled)) { Remove-WorkbookContentPdfsAll $Workspace $WorkbookId $KeepVersionId; return }
        $keepCount = 3
        try { $keepCount = [int](Get-InputHistorySettings).retainContentPdfVersions } catch { }
        if ($keepCount -lt 1) { $keepCount = 1 }
        $bookDir0 = Join-Path $Workspace (Join-Path 'content-pdf' $WorkbookId)
        if (Test-Path -LiteralPath $bookDir0) {
            $dirs = @(Get-ChildItem -LiteralPath $bookDir0 -Directory -ErrorAction SilentlyContinue | Sort-Object Name)
            $keep = @()
            if (-not [string]::IsNullOrWhiteSpace($KeepVersionId)) { $keep += $KeepVersionId }
            $keep += @($dirs | Select-Object -Last $keepCount | ForEach-Object { [string]$_.Name })
            foreach ($d in $dirs) {
                $name = [string]$d.Name
                if ($keep -contains $name) { continue }
                if (Test-ContentPdfProtected $Workspace $WorkbookId $name) { continue }
                Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
                [void]$Script:ContentPdfSheetIndexCache.Remove(([IO.Path]::GetFullPath($d.FullName)).ToLowerInvariant())
            }
        }
        return
    } catch { }
}

function Remove-WorkbookContentPdfsAll([string]$Workspace, [string]$WorkbookId, [string]$KeepVersionId) {
    try {
        $bookDir = Join-Path $Workspace (Join-Path 'content-pdf' $WorkbookId)
        if (-not (Test-Path -LiteralPath $bookDir)) { return }
        foreach ($dir in @(Get-ChildItem -LiteralPath $bookDir -Directory -ErrorAction SilentlyContinue)) {
            if ([string]::IsNullOrWhiteSpace($KeepVersionId) -or [string]$dir.Name -ne $KeepVersionId) {
                Remove-Item -LiteralPath $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
                [void]$Script:ContentPdfSheetIndexCache.Remove(([IO.Path]::GetFullPath($dir.FullName)).ToLowerInvariant())
            }
        }
    } catch { }
}

function Unregister-Workbook([string]$Language, [string]$WorkbookId) {
    if ([string]::IsNullOrWhiteSpace($WorkbookId)) { throw 'workbookId が必要です。' }
    $result=Update-StructureLocked $Language {
        param($structure)
        $found=@(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
        if ($found.Count -eq 0) { throw "登録済みExcelが見つかりません: $WorkbookId" }
        $cat=Get-WorkbookPackId $found[0]
        $removedPages=@(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId })
        $affected=@($removedPages | ForEach-Object {[string]$_.volume} | Where-Object {$_ -and $_ -ne 'none'} | Select-Object -Unique)
        # 取り消せる場所を作ってから消す。ページの振り分け・並び順・ページ名がまとめて
        # 失われる操作なのに、ここだけスナップショットを取っていなかった(他の破壊的な
        # 操作はすべて取っている)。復元プレビューは「過去にだけ存在するページ」を
        # 適用しないので、これが無いと戻す手段が一つも無い。
        if ($removedPages.Count -gt 0) { [void](Save-LayoutSnapshot $Language $cat 'pre-unregister' $structure) }
        $structure.workbooks=@(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -ne $WorkbookId })
        $structure.pages=@(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -ne $WorkbookId })
        Mark-VolumeNeedsRebuild $structure $Language $cat $affected 'unregister' 'Excelを1件登録解除しました'
        return [ordered]@{workbookId=$WorkbookId;fileName=[string]$found[0].fileName;removedPages=$removedPages}
    }
    # V5-§3.8: 登録解除と履歴削除は別操作。履歴が有効なら content-pdf も残す
    # (正式PDFアーカイブが過去の世代を参照している可能性があるため)。
    try {
        if (-not (Test-InputHistoryEnabled)) { Remove-WorkbookContentPdfsAll (Get-WorkspacePath $Language) $WorkbookId '' }
    } catch { }
    return $result
}

function Render-Workbook([string]$Language, [string]$WorkbookId, $SharedExcel = $null, [bool]$KeepExcelOpen = $false, [scriptblock]$ProgressCallback = $null,
                         [string]$SourceOverridePath = '', [string]$SourceSnapshotId = '', [string]$ExpectedSourceHash = '', [string]$ExcelSheetSelection = 'all-visible') {
    $ExcelSheetSelection = Normalize-ExcelSheetSelection $ExcelSheetSelection
    $paths = Get-Paths
    $workspace = Get-WorkspacePath $Language
    $structure = Get-Structure $Language
    $workbook = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
    if ($workbook.Count -eq 0) { throw "Workbookが見つかりません: $WorkbookId" }
    $wb = $workbook[0]
    $sourcePath = Join-Safe ([string]$paths.submissionDir) ([string]$wb.relativePath)
    # V5: 保存済み検知版からレンダリングする場合は、提出フォルダの現物が消えていても続行できる。
    if ([string]::IsNullOrWhiteSpace($SourceOverridePath) -and -not (Test-Path -LiteralPath $sourcePath)) {
        $hasSnapshotInput = $false
        try { $probe = Capture-RenderInput $Language $WorkbookId $SourceSnapshotId ''; $hasSnapshotInput = (-not [string]::IsNullOrWhiteSpace([string]$probe.path)) } catch { }
        if (-not $hasSnapshotInput) {
            Update-StructureLocked $Language { param($st) $x=@(Get-Array $st.workbooks|Where-Object{[string]$_.workbookId -eq $WorkbookId}|Select-Object -First 1);if($x.Count){Set-NoteProperty $x[0] 'status' 'missing'} } | Out-Null
            throw "提出ファイルが見つかりません: $($wb.relativePath)"
        }
    }
    # V5-D1: render-engine -> render_<workbookId> の順で取得する。
    $Script:PendingAnalysis = $null
    $renderResult = Invoke-WithRenderLock $Language $WorkbookId {
        $renderTimer = [Diagnostics.Stopwatch]::StartNew()
        $timingsMs = [ordered]@{}
        Add-NotePropertyIfMissing $wb 'lastError' ''
        Add-NotePropertyIfMissing $wb 'lastErrorUser' ''
        Add-NotePropertyIfMissing $wb 'lastErrorAt' $null
        Add-NotePropertyIfMissing $wb 'lastRenderAttemptHash' ''
        Add-NotePropertyIfMissing $wb 'lastRenderLog' ''
        Set-NoteProperty $wb 'status' 'rendering'
        Set-NoteProperty $wb 'lastError' ''
        Set-NoteProperty $wb 'lastErrorUser' ''
        Set-NoteProperty $wb 'lastErrorAt' $null
        Update-StructureLocked $Language { param($st) $x=@(Get-Array $st.workbooks|Where-Object{[string]$_.workbookId -eq $WorkbookId}|Select-Object -First 1);if($x.Count){Set-NoteProperty $x[0] 'status' 'rendering';Set-NoteProperty $x[0] 'lastError' '';Set-NoteProperty $x[0] 'lastErrorUser' '';Set-NoteProperty $x[0] 'lastErrorAt' $null} } | Out-Null

        # 誰かが保存した直後やウイルススキャン中は読み取りが一時的に失敗するため、少し待ってリトライする。
        # V5-§6.1/§C: 先にレンダリング入力を確定させる。
        # ハッシュは「実際に開くファイル」に対して計算しなければならない。
        # 現行Excelのハッシュを使うと、検知版からレンダリングしたのに最新扱いになりCASが壊れる。
        $inputInfo = $null
        $inputHashVerified = $false
        $ephemeralCaptureId = ''
        if ([string]::IsNullOrWhiteSpace($SourceOverridePath)) {
            try { $inputInfo = Capture-RenderInput $Language $WorkbookId $SourceSnapshotId '' } catch { $inputInfo = $null }
            if ($null -ne $inputInfo -and -not [string]::IsNullOrWhiteSpace([string]$inputInfo.path)) {
                $SourceOverridePath = [string]$inputInfo.path
                $SourceSnapshotId = [string]$inputInfo.snapshotId
                $ExpectedSourceHash = [string]$inputInfo.hash
                $inputHashVerified = [bool](Get-DataProperty $inputInfo 'verified' $false)
                if ([bool]$inputInfo.ephemeral) { $ephemeralCaptureId = [string]$inputInfo.captureId }
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($SourceOverridePath)) {
            if (-not (Test-Path -LiteralPath $SourceOverridePath)) { throw "レンダリング入力が見つかりません: $SourceOverridePath" }
            $sourcePath = $SourceOverridePath
            $item = Get-Item -LiteralPath $sourcePath -ErrorAction Stop
        }
        $sourceHash = ''
        $lastReadError = ''
        if ($inputHashVerified -and -not [string]::IsNullOrWhiteSpace($ExpectedSourceHash)) {
            # Capture-RenderInput already verified this immutable/local captured input.
            # Re-reading the whole XLSX from a shared snapshot would duplicate the slowest I/O.
            $sourceHash = Normalize-FileHash $ExpectedSourceHash
        } else {
            for ($readAttempt = 1; $readAttempt -le 3; $readAttempt++) {
                try {
                    $sourceHash = New-StableHash $sourcePath
                    if (-not [string]::IsNullOrWhiteSpace($sourceHash)) { break }
                } catch { $lastReadError = $_.Exception.Message }
                if ($readAttempt -lt 3) { Start-Sleep -Seconds 2 }
            }
        }
        # V5-§6.3a: catch 経路が「どの版を試行したか」を知るために記録する。
        # 同一プロセス内のレンダリングは render-engine ロックで直列化されているため安全。
        # V5-P0(#3): 固定版(pin)の指定がある場合、試行した版は「意図した ExpectedSourceHash」であって
        # 現物ファイル($sourceHash=現在版)ではない。現物のハッシュを記録すると、版ずれ失敗時に
        # Set-WorkbookRenderError の sameVersion 判定が誤って現在版を render-error 化してしまう。
        $attemptHashForRecord = if (-not [string]::IsNullOrWhiteSpace($ExpectedSourceHash)) { [string]$ExpectedSourceHash } else { [string]$sourceHash }
        $Script:LastRenderAttempt = [ordered]@{ workbookId = $WorkbookId; snapshotId = [string]$SourceSnapshotId; hash = [string]$attemptHashForRecord }
        if (-not [string]::IsNullOrWhiteSpace($ExpectedSourceHash) -and -not [string]::IsNullOrWhiteSpace($sourceHash)) {
            if ((Normalize-FileHash $sourceHash) -ne (Normalize-FileHash $ExpectedSourceHash)) {
                throw 'レンダリング入力が想定した版と一致しません。もう一度PDF作成してください。'
            }
        }
        if ([string]::IsNullOrWhiteSpace($sourceHash)) {
            throw "提出Excelを読み取れませんでした(3回試行)。誰かが保存中か、排他モードで開かれている可能性があります。少し待ってからもう一度PDF作成してください。 $lastReadError"
        }
        $item = Get-Item -LiteralPath $sourcePath
        # V5: 秒単位だと同一秒の2回処理で同じ世代フォルダを共有してしまう。ミリ秒+GUID8にする。
        $versionId = New-RbVersionId
        $contentDir = Join-Path $workspace (Join-Path 'content-pdf' (Join-Path $WorkbookId $versionId))
        New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
        $tmpDir = Join-Path ([string]$paths.dataDir) (Join-Path 'common\tmp' ([Guid]::NewGuid().ToString('N')))
        New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
        $tmpPath = Join-Path $tmpDir ([IO.Path]::GetFileName($sourcePath))
        $copyTimer = [Diagnostics.Stopwatch]::StartNew()
        $lastCopyError = ''
        $copied = $false
        for ($copyAttempt = 1; $copyAttempt -le 3; $copyAttempt++) {
            try {
                Copy-FileSharedRead $sourcePath $tmpPath
                $copied = $true
                break
            } catch { $lastCopyError = $_.Exception.Message }
            if ($copyAttempt -lt 3) { Start-Sleep -Seconds 2 }
        }
        if (-not $copied) {
            throw "提出Excelをコピーできませんでした(3回試行)。誰かが保存中の可能性があります。少し待ってからもう一度PDF作成してください。 $lastCopyError"
        }
        # Copy-ItemはZone.Identifier（Mark of the Web）を一時コピーへ引き継ぐため、
        # 保護ビューによるWorkbooks.Open失敗を防ぐ目的で明示的に除去する。
        try { Unblock-File -LiteralPath $tmpPath -ErrorAction SilentlyContinue } catch { }
        # Full SHA-256 twice per workbook was expensive. The source has already been hashed;
        # after copying, confirm size and source timestamp/size stability instead of hashing the copy again.
        $copyItem = Get-Item -LiteralPath $tmpPath -ErrorAction Stop
        $sourceAfterCopy = Get-Item -LiteralPath $sourcePath -ErrorAction Stop
        if ($copyItem.Length -ne $item.Length -or $sourceAfterCopy.Length -ne $item.Length -or $sourceAfterCopy.LastWriteTimeUtc -ne $item.LastWriteTimeUtc) {
            throw 'コピー中に提出Excelが更新されました。保存が終わってからもう一度PDF作成してください。'
        }
        $timingsMs.localCopy = [int64]$copyTimer.ElapsedMilliseconds
        # V5-§3.1/§6.2: 提出フォルダの現物からレンダリングする場合は、
        # ヘッダー/フッターXMLを加工する前(L直後)に未加工の検知版を保存する。
        if ([string]::IsNullOrWhiteSpace($SourceSnapshotId) -and (Test-InputHistoryEnabled)) {
            try {
                $capMeta = Ensure-SnapshotMetadata $Language $WorkbookId ([string]$wb.relativePath) ([string]$wb.category) $sourcePath $sourceHash 'render'
                if ([bool]$capMeta.ok) {
                    $capId = [string]$capMeta.snapshotId
                    $capOk = $true
                    if ([bool](Get-DataProperty $capMeta 'pending' $false)) {
                        if (Test-SourceRetentionEnabled) {
                            $saved = Save-SnapshotSourceFile $Language $WorkbookId $capId $sourcePath $sourceHash
                            if (-not [bool]$saved.ok) { $capOk = $false }
                        }
                        if ($capOk) { $capOk = Complete-Snapshot $Language $WorkbookId $capId }
                    }
                    if ($capOk) {
                        $SourceSnapshotId = $capId
                        $Script:LastRenderAttempt = [ordered]@{ workbookId = $WorkbookId; snapshotId = $SourceSnapshotId; hash = [string]$sourceHash }
                    }
                }
            } catch { }
        }

        $packageTimer = [Diagnostics.Stopwatch]::StartNew()
        $excelPackagePreparation = $null
        try { $excelPackagePreparation = Prepare-XlsxPrintPackage $tmpPath } catch { $excelPackagePreparation = [ordered]@{ ok=$false; printSettingsPrepared=$false; error=$_.Exception.Message } }
        $packagePrintSettingsPrepared = [bool](Get-DataProperty $excelPackagePreparation 'printSettingsPrepared' $false)
        $timingsMs.packagePreparation = [int64]$packageTimer.ElapsedMilliseconds

        $excel = $SharedExcel
        $ownsExcel = $false
        $book = $null
        $rendered = @()
        $warnings = @()
        $steps = @()
        try {
            if ($ProgressCallback) { & $ProgressCallback 'prepare' $WorkbookId '' }
            if ($null -eq $excel) {
                $steps += 'Excel COMを起動'
                $excelStartTimer = [Diagnostics.Stopwatch]::StartNew()
                $excel = New-ExcelApplicationForRender
                $timingsMs.excelStartup = [int64]$excelStartTimer.ElapsedMilliseconds
                $ownsExcel = $true
            } else {
                $steps += '既存のExcel COMを使用'
                $timingsMs.excelStartup = 0
                try { $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.ScreenUpdating = $false } catch { }
            }
            $envInfo = Get-CachedRenderEnvironment $excel
            $Script:CurrentRenderEnvFingerprint = Get-RenderEnvironmentFingerprint $envInfo
            $Script:CurrentRenderEnvInfo = $envInfo
            if (-not $Script:RenderEnvironmentCompared) {
                $warnings += Compare-And-SaveEnvironment $Language $envInfo
                Write-JsonFile (Join-Path $workspace "logs\render-env_$versionId.json") $envInfo
                $Script:RenderEnvironmentCompared = $true
            }

            if ($packagePrintSettingsPrepared) {
                $steps += "Excelを開く前に印刷設定を準備: $([int](Get-DataProperty $excelPackagePreparation 'sheetsPrepared' 0)) シート"
            } elseif ($excelPackagePreparation -and $excelPackagePreparation.ok -eq $false) {
                $warnings += "Excel内部の印刷設定を事前準備できませんでした。COM設定で続行します: $([string](Get-DataProperty $excelPackagePreparation 'error' ''))"
            }
            $steps += '一時コピーを開く'
            if ($ProgressCallback) { & $ProgressCallback 'open' $WorkbookId '' }
            $excelOpenTimer = [Diagnostics.Stopwatch]::StartNew()
            $book = Open-ExcelWorkbookSafe $excel $tmpPath $true
            $timingsMs.excelOpen = [int64]$excelOpenTimer.ElapsedMilliseconds
            try { $book.CheckCompatibility = $false } catch { }

            $sheetSetupTimer = [Diagnostics.Stopwatch]::StartNew()
            $inspected = @()
            $targetSheetNames = @()
            # 非表示シートも「原稿に存在する」として記録する。表示中のものだけを渡すと、
            # 一時的に非表示にしただけでページ設定が削除されてしまう。
            $allSheetNames = @()
            $sheetRenderInfos = @()
            $excludedSheetNames = @()
            $hiddenSheetNames = @()
            $allSheetInfos = @()
            $sheetCount = 0
            try { $sheetCount = [int]$book.Worksheets.Count } catch { $sheetCount = 0 }
            $deferredPrintCommunication = $false
            if (-not $packagePrintSettingsPrepared) {
                $deferredPrintCommunication = Set-ExcelPrintCommunicationSafe $excel $false
            }
            try {
                for ($i = 1; $i -le $sheetCount; $i++) {
                    $ws = $null
                    try {
                        $ws = $book.Worksheets.Item($i)
                        $sheetName = [string]$ws.Name
                        $visible = ([int]$ws.Visible -eq -1)
                        $allSheetNames += $sheetName
                        $a1 = ''
                        try { $a1 = [string]$ws.Range('A1').Text } catch { }
                        if ([string]::IsNullOrWhiteSpace($a1)) { $a1 = "$($wb.fileName) / $sheetName" }
                        $sheetInfo = [ordered]@{ sheetName = $sheetName; sheetIndex = $i; titleSource = 'A1'; detectedTitle = $a1; printArea = '' }
                        $allSheetInfos += $sheetInfo
                        if (-not $visible) {
                            $excludedSheetNames += $sheetName
                            $hiddenSheetNames += $sheetName
                        } elseif ($ExcelSheetSelection -eq 'numeric-only' -and -not (Test-StrictNumericSheetName $sheetName)) {
                            $excludedSheetNames += $sheetName
                        } else {
                            $targetSheetNames += $sheetName
                            # Reading PageSetup.PrintArea is another slow COM call and is only diagnostic.
                            # Keep it blank in render logs to avoid delaying PDF作成.
                            $printArea = ''
                            $inspected += $sheetInfo

                            if (-not $packagePrintSettingsPrepared) {
                                $steps += "シート $sheetName の印刷設定を調整"
                                if ($ProgressCallback) { & $ProgressCallback 'sheet-setup' $WorkbookId $sheetName }
                                Apply-StandardPrintSettings $ws $excel $deferredPrintCommunication
                            }
                            $outPdf = Join-Path $contentDir ("{0}.pdf" -f (Get-WorksheetStorageStem $sheetName))
                            $sheetRenderInfos += [ordered]@{ sheetName = $sheetName; sheetIndex = $i; outPdf = $outPdf; titleSource = 'A1'; detectedTitle = $a1; printArea = $printArea }
                        }
                    } finally {
                        Invoke-ComRelease $ws
                    }
                }
            } finally {
                if ($deferredPrintCommunication) { [void](Set-ExcelPrintCommunicationSafe $excel $true) }
            }
            $timingsMs.sheetInspectionAndSetup = [int64]$sheetSetupTimer.ElapsedMilliseconds
            if ($targetSheetNames.Count -eq 0) {
                try { if (Test-Path -LiteralPath $contentDir) { Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue } } catch { }
                if ($ExcelSheetSelection -eq 'numeric-only') {
                    throw '半角数字だけの表示シートがありません。シート名を半角数字（例: 1、2、10）にしてから、もう一度PDFを作成してください。PDFやページ構成は変更していません。'
                }
                throw 'PDF化対象の表示シートがありません。Excelで少なくとも1つのワークシートを表示してください。'
            }

            $excelExportTimer = [Diagnostics.Stopwatch]::StartNew()
            $batchResult = $null
            if (@($sheetRenderInfos).Count -gt 1) {
                if ($ProgressCallback) { & $ProgressCallback 'batch' $WorkbookId '' }
                $steps += "複数シートを一括PDF化: $(@($sheetRenderInfos).Count) シート"
                $batchResult = Export-WorkbookSheetsToPdfBatch $excel $book $sheetRenderInfos $tmpDir
                if ($batchResult -and $batchResult.ok -eq $true) {
                    $steps += '一括PDFをシート別PDFへ分割'
                    if ($ProgressCallback) { & $ProgressCallback 'split' $WorkbookId '' }
                } else {
                    $reason = [string](Get-DataProperty $batchResult 'reason' '')
                    $msg = [string](Get-DataProperty $batchResult 'message' '')
                    if ($reason -ne 'single-sheet' -and $reason -ne 'pdfbox-unavailable') {
                        $warnings += "一括PDF化を使えなかったため従来方式で続行します: $msg"
                    }
                }
            }

            if (-not ($batchResult -and $batchResult.ok -eq $true)) {
                foreach ($info in @($sheetRenderInfos)) {
                    $ws = $null
                    try {
                        $sheetName = [string]$info.sheetName
                        $steps += "シート $sheetName をPDF化"
                        if ($ProgressCallback) { & $ProgressCallback 'sheet' $WorkbookId $sheetName }
                        $ws = $book.Worksheets.Item($sheetName)
                        [void](Export-WorksheetToPdfSafe $excel $book $ws ([string]$info.outPdf) $sheetName $true)
                    } finally {
                        Invoke-ComRelease $ws
                    }
                }
            }

            $usedBatchOutput = ($batchResult -and $batchResult.ok -eq $true)
            $timingsMs.excelExportAndSplit = [int64]$excelExportTimer.ElapsedMilliseconds
            foreach ($info in @($sheetRenderInfos)) {
                $sheetName = [string]$info.sheetName
                $outPdf = [string]$info.outPdf
                # Batch export is accepted only when total pages equals target sheet count, and the splitter
                # writes one page per target. Avoid rereading every split PDF just to count pages.
                $pageCount = if ($usedBatchOutput) { 1 } else { Get-PdfPageCount $outPdf }
                $pageWarnings = @()
                if ($pageCount -gt 1) { $pageWarnings += "このシートのPDFは $pageCount ページです。Excelの印刷範囲・改ページ・倍率を確認してください。" }
                $rendered += [ordered]@{ sheetName = $sheetName; pdf = $outPdf; pageCount = $pageCount; warnings = $pageWarnings }
            }
            if ($rendered.Count -eq 0) {
                throw 'PDFを作成できませんでした。Excelの印刷設定または対象シートを確認してください。'
            }
            if ($ExcelSheetSelection -eq 'numeric-only' -and $excludedSheetNames.Count -gt 0) {
                $warnings += "半角数字以外または非表示のExcelシート $($excludedSheetNames.Count) 件は今回のPDFに含めず、既存ページを未振り分けに戻しました。"
            }

            # This copy is used only for the response/warning summary. The real
            # snapshot and layout mutation happen once, against the latest
            # structure, inside Update-StructureLocked below.
            $pageSync = Update-WorkbookPagesFromInspection $Language $structure $wb $allSheetInfos $allSheetNames $targetSheetNames $hiddenSheetNames $false
            Set-ExcludedSheetPagesNotRendered $structure $WorkbookId $excludedSheetNames $ExcelSheetSelection
            $selectionPack = (Resolve-DocumentPackScope $structure (Get-WorkbookPackId $wb) $true).pack
            foreach ($selectionVolume in @(Get-PackVolumeList $Language $selectionPack $true)) { [void](Renumber-VolumeOrder $structure $selectionVolume (Get-WorkbookPackId $wb)) }
            $removedPages = @(Get-Array (Get-DataProperty $pageSync 'removedPages' @()))
            if ($removedPages.Count -gt 0) {
                $removedNames = @($removedPages | ForEach-Object { [string](Get-DataProperty $_ 'sheetName' '') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $removedLabel = [string]::Join('、', $removedNames)
                if ([string]::IsNullOrWhiteSpace($removedLabel)) { $removedLabel = "$($removedPages.Count) ページ" }
                $warnings += "現在のExcelに存在しないシートをページ構成から外しました: $removedLabel（ページ構成の履歴から元に戻せます）"
                $steps += "存在しないシートの古いページを削除: $removedLabel"
                # V5-§3.7: 旧世代の個別削除は廃止(上記と同じ理由)。
            }
            $renamedPages = @(Get-Array (Get-DataProperty $pageSync 'renamedPages' @()))
            if ($renamedPages.Count -gt 0) {
                $renameLabel = [string]::Join('、', @($renamedPages | ForEach-Object {
                    "{0}→{1}" -f [string](Get-DataProperty $_ 'fromSheetName' ''), [string](Get-DataProperty $_ 'toSheetName' '')
                }))
                $warnings += "シート名の変更を引き継ぎました: $renameLabel（配置とページ設定はそのままです）"
                $steps += "シート名変更の引き継ぎ: $renameLabel"
            }
            $hiddenPageIds = @(Get-Array (Get-DataProperty $pageSync 'hiddenSheetPageIds' @()))
            if ($hiddenPageIds.Count -gt 0) {
                $warnings += "非表示のシートが $($hiddenPageIds.Count) 件あります。ページ構成は保持していますが、PDFは作り直されません。提出物に含める場合はシートを再表示してからPDFを作成してください。"
                $steps += "非表示シートのページを保持: $($hiddenPageIds.Count) 件"
            }
            Set-WorkbookRenderedSheetSnapshot $wb (Get-DataProperty $pageSync 'sheetNames' @())
            $pages = @(Get-Array $structure.pages)
            foreach ($r in $rendered) {
                $pageId = "$WorkbookId-$([regex]::Replace([string]$r.sheetName, '[^0-9A-Za-z]+', '-'))"
                $p = @($pages | Where-Object { (Resolve-PageId $_) -eq $pageId -or ([string]$_.workbookId -eq $WorkbookId -and [string]$_.sheetName -eq [string]$r.sheetName) } | Select-Object -First 1)
                if ($p.Count -gt 0) {
                    $rel = Get-RelativePathCompat $workspace ([string]$r.pdf)
                    Set-NoteProperty ($p[0]) 'contentPdf' $rel
                    Set-NoteProperty ($p[0]) 'status' 'rendered'
                    Set-NoteProperty ($p[0]) 'warnings' @($r.warnings)
                    Set-NoteProperty ($p[0]) 'sheetSelectionExcluded' $false
                    Set-NoteProperty ($p[0]) 'updatedAt' (New-NowIso)
                }
            }
            foreach ($p in @($pages | Where-Object { [string]$_.workbookId -eq $WorkbookId -and [string]$_.contentPdf -and [string]$_.status -eq 'confirmed' })) {
                if ($sourceHash -ne [string]$wb.lastRenderedExcelHash) { Set-NoteProperty $p 'status' 'stale' }
            }
            # V5-§2.4: currentExcel* は Scan-Updates の専有。レンダリングは書かない。
            # レンダリング中に別プロセスが検知した新しい版を、古い版のハッシュで上書きしないため。
            Set-NoteProperty $wb 'lastRenderedVersionId' $versionId
            Set-NoteProperty $wb 'lastRenderedExcelHash' $sourceHash
            Set-NoteProperty $wb 'lastRenderedAt' (New-NowIso)
            Set-WorkbookRenderedSheetSnapshot $wb (Get-DataProperty $pageSync 'sheetNames' @())
            Set-NoteProperty $wb 'renderProfileVersion' $Script:ExcelPrintProfileVersion
            Set-NoteProperty $wb 'lastRenderedSheetSelectionMode' $ExcelSheetSelection
            Add-NotePropertyIfMissing $wb 'lastError' ''
            Add-NotePropertyIfMissing $wb 'lastErrorUser' ''
            Add-NotePropertyIfMissing $wb 'lastErrorAt' $null
            Add-NotePropertyIfMissing $wb 'lastRenderAttemptHash' ''
            Add-NotePropertyIfMissing $wb 'lastRenderLog' ''
            Set-NoteProperty $wb 'status' 'rendered-unchecked'
            Set-NoteProperty $wb 'lastError' ''
            Set-NoteProperty $wb 'lastErrorUser' ''
            Set-NoteProperty $wb 'lastErrorAt' $null
            Set-NoteProperty $wb 'lastRenderAttemptHash' $sourceHash
            Set-NoteProperty $wb 'warnings' @($warnings)
            $logRel = "logs\render_$WorkbookId`_$versionId.json"
            Set-NoteProperty $wb 'lastRenderLog' $logRel
            $pageSync = Update-StructureLocked $Language {
                param($st)
                $latest=@(Get-Array $st.workbooks|Where-Object{[string]$_.workbookId -eq $WorkbookId}|Select-Object -First 1)
                if(-not $latest.Count){throw "Workbookが見つかりません: $WorkbookId"}
                $lw=$latest[0]
                $selectionAffectedVolumesLocked = @()
                $selectionNeedsSnapshotLocked = $false
                if ($ExcelSheetSelection -eq 'numeric-only' -and $excludedSheetNames.Count -gt 0) {
                    $excludedSetLocked = @{}; foreach ($excludedName in $excludedSheetNames) { $excludedSetLocked[[string]$excludedName] = $true }
                    $excludedPagesLocked = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId -and $excludedSetLocked.ContainsKey([string]$_.sheetName) })
                    $selectionAffectedVolumesLocked = @($excludedPagesLocked | Where-Object { [string]$_.volume -and [string]$_.volume -ne 'none' } | ForEach-Object { [string]$_.volume } | Select-Object -Unique)
                    $selectionNeedsSnapshotLocked = @($excludedPagesLocked | Where-Object { ([string]$_.volume -and [string]$_.volume -ne 'none') -or $_.enabled -ne $false }).Count -gt 0
                    if ($selectionNeedsSnapshotLocked) { [void](Save-LayoutSnapshot $Language (Get-WorkbookPackId $lw) 'sheet-selection' $st) }
                }
                # If the numeric exclusion already made a restore point, let it
                # cover simultaneous rename/removal changes too. Otherwise the
                # normal source-sheet change path creates its one snapshot.
                $sync=Update-WorkbookPagesFromInspection $Language $st $lw $allSheetInfos $allSheetNames $targetSheetNames $hiddenSheetNames (-not $selectionNeedsSnapshotLocked)
                Set-ExcludedSheetPagesNotRendered $st $WorkbookId $excludedSheetNames $ExcelSheetSelection
                $selectionPackLocked = (Resolve-DocumentPackScope $st (Get-WorkbookPackId $lw) $true).pack
                foreach ($selectionVolume in @(Get-PackVolumeList $Language $selectionPackLocked $true)) { [void](Renumber-VolumeOrder $st $selectionVolume (Get-WorkbookPackId $lw)) }
                foreach ($r in $rendered) {
                    $pageId = "$WorkbookId-$([regex]::Replace([string]$r.sheetName, '[^0-9A-Za-z]+', '-'))"
                    # Unrestricted worksheet names use stable hashed page IDs. The legacy
                    # sanitized ID above cannot identify Japanese-only names, so retain the
                    # same workbook/sheet fallback used by the in-memory render result.
                    $pg = @(Get-Array $st.pages | Where-Object {
                        (Resolve-PageId $_) -eq $pageId -or
                        ([string]$_.workbookId -eq $WorkbookId -and [string]$_.sheetName -eq [string]$r.sheetName)
                    } | Select-Object -First 1)
                    if ($pg.Count) {
                        Set-NoteProperty $pg[0] 'contentPdf' (Get-RelativePathCompat $workspace ([string]$r.pdf))
                        Set-NoteProperty $pg[0] 'status' 'rendered'
                        Set-NoteProperty $pg[0] 'warnings' @($r.warnings)
                        Set-NoteProperty $pg[0] 'sheetSelectionExcluded' $false
                        Set-NoteProperty $pg[0] 'updatedAt' (New-NowIso)
                    }
                }
                # V5-§2.4: currentExcel* はコピーしない(Scan-Updates の専有)。
                foreach($name in @('lastRenderedVersionId','lastRenderedExcelHash','lastRenderedAt','renderProfileVersion','lastRenderedSheetSelectionMode','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash','warnings','lastRenderLog')){Set-NoteProperty $lw $name (Get-DataProperty $wb $name $null)}
                Set-NoteProperty $lw 'lastRenderedSnapshotId' ([string]$SourceSnapshotId)
                Set-NoteProperty $lw 'renderEnvironmentFingerprint' ([string]$Script:CurrentRenderEnvFingerprint)
                # V5-§6.3: ロック内で最新の currentExcelHash と突き合わせて status を決める。
                # レンダリング中に元Excelが更新されていれば、PDFは保存しつつ excel-updated に戻す。
                $latestCurrentHash = Normalize-FileHash ([string](Get-DataProperty $lw 'currentExcelHash' ''))
                $renderedHash = Normalize-FileHash $sourceHash
                if ([string]::IsNullOrWhiteSpace($latestCurrentHash) -or $latestCurrentHash -eq $renderedHash) {
                    Set-NoteProperty $lw 'status' 'rendered-unchecked'
                } else {
                    Set-NoteProperty $lw 'status' 'excel-updated'
                    foreach($pg2 in @(Get-Array $st.pages|Where-Object{[string]$_.workbookId -eq $WorkbookId -and [string]$_.status -eq 'rendered'})){ Set-NoteProperty $pg2 'status' 'stale' }
                }
                Set-WorkbookRenderedSheetSnapshot $lw (Get-DataProperty $sync 'sheetNames' @())
                $cat=Get-WorkbookPackId $lw;$vols=@(Get-Array $st.pages|Where-Object{[string]$_.workbookId -eq $WorkbookId}|ForEach-Object{[string]$_.volume}|Where-Object{$_ -and $_ -ne 'none'}|Select-Object -Unique);$vols=@($vols+$selectionAffectedVolumesLocked|Select-Object -Unique);Mark-VolumeNeedsRebuild $st $Language $cat $vols 'render' 'Excelを1件PDF作成しました'
                return $sync
            }
            $timingsMs.total = [int64]$renderTimer.ElapsedMilliseconds
            Write-JsonFile (Join-Path $workspace $logRel) ([ordered]@{ workbookId = $WorkbookId; sheetSelectionMode = $ExcelSheetSelection; excludedSheetNames = @($excludedSheetNames); rendered = $rendered; warnings = $warnings; steps = $steps; timingsMs = $timingsMs; sheetSync = $pageSync; at = New-NowIso })
            # V5-P1: 解析はここでは実行しない。
            # レンダリング用Excelを開いたまま、レンダリングロックを保持したまま解析すると、
            # 比較用の再レンダリングが2つ目のExcel COMを起動してしまう(1ジョブ制限に反する)。
            # ロック解放後に実行するため、対象だけを記録しておく。
            $Script:PendingAnalysis = [ordered]@{ language = $Language; workbookId = $WorkbookId; snapshotId = [string]$SourceSnapshotId; versionId = $versionId; rendered = $rendered; sheetSelectionMode = $ExcelSheetSelection }
            Remove-WorkbookContentPdfs $workspace $WorkbookId $versionId
        } finally {
            if ($book) { try { $book.Close($false) } catch { } ; Invoke-ComRelease $book }
            if ($ownsExcel -and $excel) { Close-ExcelApplicationForRender $excel }
            if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (-not [string]::IsNullOrWhiteSpace($ephemeralCaptureId)) { Remove-EphemeralCopy $Language $ephemeralCaptureId }
            if ($ownsExcel -or -not $KeepExcelOpen) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
        }
        return [ordered]@{ workbookId = $WorkbookId; versionId = $versionId; sheetSelectionMode = $ExcelSheetSelection; excludedSheetNames = @($excludedSheetNames); rendered = $rendered; warnings = $warnings; steps = $steps; timingsMs = $timingsMs; sheetSync = $pageSync }
    }

    # V5-P1: ここではレンダリングロックもExcelも解放済み。
    # 比較用の再レンダリングが必要になっても、改めて共通ロックを取り直せる。
    # 解析の失敗はPDF作成の失敗にしない(判定は unknown になる)。
    # V5-P3: KeepExcelOpen のとき($Script:PendingAnalysis を呼出元が回収する一括ジョブ)は
    # ここで消してはならない。以前は無条件に $null を代入していたため、
    # Invoke-RenderJobFromFile の $deferredAnalyses が常に空になり、
    # 画像ハッシュの解析が一度も実行されていなかった(renders フォルダが作られない)。
    if ($KeepExcelOpen) { return $renderResult }
    $pending = $Script:PendingAnalysis
    $Script:PendingAnalysis = $null
    if ($null -ne $pending) {
        try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered ([string](Get-DataProperty $pending 'sheetSelectionMode' 'all-visible'))) }
        catch { Write-Warning ('画像ハッシュの解析に失敗しました: ' + $_.Exception.Message) }
    }
    return $renderResult
}

function Scan-Updates([string]$Language, [scriptblock]$ProgressCallback = $null, [bool]$ForceHash = $false) {
    $paths = Get-Paths
    $snapshot = Get-Structure $Language
    $changed = @()
    $scanned = 0; $hashed = 0; $metadataOnly = 0
    $list = @(Get-Array $snapshot.workbooks)
    $index = 0
    foreach ($snap in $list) {
        $index++
        if ($ProgressCallback) { & $ProgressCallback $index $list.Count ([string]$snap.workbookId) ([string]$snap.displayName) }
        $id = [string]$snap.workbookId
        try { $full = Join-Safe ([string]$paths.submissionDir) ([string]$snap.relativePath) } catch { continue }
        if (-not (Test-Path -LiteralPath $full)) {
            $didMissing = Update-StructureLocked $Language {
                param($st)
                $found = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq $id } | Select-Object -First 1)
                if ($found.Count -eq 0) { return $false }
                $w = $found[0]
                $wasMissing = ([string]$w.status -eq 'missing')
                Set-NoteProperty $w 'status' 'missing'
                if (-not $wasMissing) {
                    $cat = Get-WorkbookPackId $w
                    $vols = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $id } | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
                    Mark-VolumeNeedsRebuild $st $Language $cat $vols 'source-updated' '元原稿が見つかりません'
                }
                return (-not $wasMissing)
            }
            if ($didMissing) { $changed += $id }
            continue
        }
        $item = Get-Item -LiteralPath $full
        $scanned++
        $ticks = [string]$item.LastWriteTimeUtc.Ticks
        $size = [int64]$item.Length
        $modified = $item.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:sszzz')
        $knownTicks = [string](Get-DataProperty $snap 'currentExcelLastWriteUtcTicks' '')
        $knownSize = [int64](Get-DataProperty $snap 'currentExcelSize' -1)
        $hash = [string](Get-DataProperty $snap 'currentExcelHash' '')
        $mustHash = $ForceHash -or $knownTicks -ne $ticks -or $knownSize -ne $size -or [string]::IsNullOrWhiteSpace($hash)
        if ($mustHash) { try { $hash = New-StableHash $full; $hashed++ } catch { $hash = '' } } else { $metadataOnly++ }
        $did = Update-StructureLocked $Language {
            param($st)
            $found = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq $id } | Select-Object -First 1)
            if ($found.Count -eq 0) { return $false }
            $w = $found[0]
            $previousHash = [string](Get-DataProperty $w 'currentExcelHash' '')
            $previousStatus = [string](Get-DataProperty $w 'status' '')
            $lastRendered = [string](Get-DataProperty $w 'lastRenderedExcelHash' '')
            $updatedStatus = Get-SourceUpdatedStatus $w
            Set-NoteProperty $w 'currentExcelModifiedAt' $modified
            # ハッシュを取得できなかった回は、更新時刻とサイズも据え置く。
            # 新しいメタデータと古いハッシュを組にして保存すると、次回のスキャンが
            # 「メタデータ一致」で再ハッシュを省略し、更新済みの原稿を最新と誤判定する。
            if (-not [string]::IsNullOrWhiteSpace($hash)) {
                Set-NoteProperty $w 'currentExcelLastWriteUtcTicks' $ticks
                Set-NoteProperty $w 'currentExcelSize' $size
                Set-NoteProperty $w 'currentExcelHash' $hash
            }
            $profileOld = (-not [string]::IsNullOrWhiteSpace($lastRendered)) -and ((Get-IntDataProperty $w 'renderProfileVersion' 0) -lt (Get-RequiredSourceRenderProfileVersion $w))
            $stale = (-not [string]::IsNullOrWhiteSpace($lastRendered)) -and (([string]::IsNullOrWhiteSpace($hash)) -or $hash -ne $lastRendered -or $profileOld)
            if ($stale) {
                Set-NoteProperty $w 'status' $updatedStatus
                foreach ($page in @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $id -and -not [string]::IsNullOrWhiteSpace([string]$_.contentPdf) })) { Set-NoteProperty $page 'status' 'stale' }
                $newlyDetected = ($previousStatus -ne $updatedStatus) -or ($previousHash -ne $hash)
                if ($newlyDetected) {
                    $cat = Get-WorkbookPackId $w
                    $vols = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $id } | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
                    Mark-VolumeNeedsRebuild $st $Language $cat $vols 'source-updated' '元原稿が1件更新されました'
                }
                return $newlyDetected
            }
            if ([string]::IsNullOrWhiteSpace($lastRendered)) {
                if ([string]$w.status -ne 'render-error') { Set-NoteProperty $w 'status' 'new' }
            } else { Set-NoteProperty $w 'status' 'rendered-unchecked' }
            return $false
        }
        if ($did) { $changed += $id }
    }
    return [ordered]@{ changedWorkbookIds=@($changed | Select-Object -Unique); scanned=$scanned; hashed=$hashed; metadataOnly=$metadataOnly; forceHash=[bool]$ForceHash }
}

function Get-AutoRenderWorkbookIds([string]$Language, [string[]]$PreferredIds, [string]$Category = '') {
    $structure = Get-Structure $Language
    $preferred = @($PreferredIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $usePreferred = ($preferred.Count -gt 0)
    $categoryNormalized = Normalize-WorkbookCategory $Category ''
    $ids = @()
    foreach ($wb in @(Get-Array $structure.workbooks)) {
        $id = [string]$wb.workbookId
        if (-not [string]::IsNullOrWhiteSpace($categoryNormalized) -and -not (Test-WorkbookCategory $wb $categoryNormalized)) { continue }
        if ($usePreferred -and (@($preferred | Where-Object { $_ -eq $id }).Count -eq 0)) { continue }
        if ([string]$wb.status -eq 'missing') { continue }
        $needs = $false
        $status = [string]$wb.status
        $currentHash = [string]$wb.currentExcelHash
        $lastRenderedHash = [string]$wb.lastRenderedExcelHash
        $lastAttemptHash = [string]$wb.lastRenderAttemptHash
        $renderProfileOutdated = ((-not [string]::IsNullOrWhiteSpace($lastRenderedHash)) -and ((Get-IntDataProperty $wb 'renderProfileVersion' 0) -lt (Get-RequiredSourceRenderProfileVersion $wb)))
        $sheetSelectionOutdated = $false
        if ([string](Get-DataProperty $wb 'sourceType' 'excel') -eq 'excel') {
            $sheetSelectionOutdated = (Normalize-ExcelSheetSelection ([string](Get-DataProperty $wb 'excelSheetSelectionMode' 'all-visible')) -ne Normalize-ExcelSheetSelection ([string](Get-DataProperty $wb 'lastRenderedSheetSelectionMode' 'all-visible')))
        }
        if ($status -eq 'render-error') {
            # The PDF button is an explicit retry. A previous render-error must not make
            # the button look idle just because the same file hash already failed once.
            $needs = $true
        } else {
            if ([string]::IsNullOrWhiteSpace($lastRenderedHash)) { $needs = $true }
            if ($status -in @('new','excel-updated','source-updated')) { $needs = $true }
            if ($renderProfileOutdated) { $needs = $true }
            if ($sheetSelectionOutdated) { $needs = $true }
            if ((-not [string]::IsNullOrWhiteSpace($currentHash)) -and ($currentHash -ne $lastRenderedHash)) { $needs = $true }
            if ($usePreferred) { $needs = $true }
        }
        if ($needs) { $ids += $id }
    }
    return @($ids | Select-Object -Unique)
}

function Invoke-AutoRender([string]$Language, [string[]]$WorkbookIds) {
    if ($Script:AutoRenderInProgress) {
        return [ordered]@{ skipped = $true; reason = 'auto-render-busy'; results = @(); at = New-NowIso }
    }
    $Script:AutoRenderInProgress = $true
    try {
        Reset-RenderEnvironmentForJob   # V5-§6.4: ジョブ開始ごとに環境を取り直す
        $scan = Scan-Updates $Language $null $false
        $ids = Get-AutoRenderWorkbookIds $Language $WorkbookIds ''
        $results = @()
        foreach ($id in $ids) {
            try {
                $r = Render-Source $Language $id
                $r['ok'] = $true
                $results += $r
            } catch {
                $msg = $_.Exception.Message
                $userMsg = ConvertTo-UserRenderError $msg
                $att = Get-LastRenderAttemptFor $id
                $errorDetail = Get-ErrorDetail $_
                Set-WorkbookRenderError $Language $id $msg $errorDetail ([string]$att.snapshotId) ([string]$att.hash)
                $results += [ordered]@{ ok = $false; workbookId = $id; error = $msg; userError = $userMsg; detail = $errorDetail }
            }
        }
        $hasErrors = $false
        foreach ($rr in $results) { try { if ($rr.Contains('ok') -and $rr['ok'] -eq $false) { $hasErrors = $true } } catch { } }
        return [ordered]@{ skipped = $false; scanned = $scan; workbookIds = $ids; results = $results; hasErrors = $hasErrors; at = New-NowIso }
    } finally {
        $Script:LastHeartbeatUtc = [DateTime]::UtcNow
        $Script:AutoRenderInProgress = $false
    }
}


function Get-RenderJobDir([string]$Language) {
    $workspace = Get-WorkspacePath $Language
    $dir = Join-Path $workspace 'state\jobs'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    return $dir
}

function Write-RenderJobStatus([string]$StatusPath, $Status) {
    if ([string]::IsNullOrWhiteSpace($StatusPath)) { return }
    try {
        Set-NoteProperty $Status 'updatedAt' (New-NowIso)
        $json = ConvertTo-Json -InputObject $Status -Depth 50
        Write-Utf8NoBomFileShared $StatusPath $json
    } catch {
        # Progress JSON is read frequently by the browser while a background PowerShell job writes it.
        # If shared write fails, fall back to the generic writer and log the failure, but do not stop Excel.
        try { Write-JsonFile $StatusPath $Status } catch { }
        try {
            $errPath = "$StatusPath.write-error.log"
            Add-Content -LiteralPath $errPath -Encoding UTF8 -Value ("{0} {1}" -f (New-NowIso), $_.Exception.Message)
        } catch { }
    }
}

function Normalize-RenderJobId([string]$JobId) {
    $value = ([string]$JobId).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    try { $value = [Uri]::UnescapeDataString($value).Trim() } catch { }

    # Browser / fetch / proxy differences should not be able to break progress polling.
    # Accept the exact job id, and also recover it if it was accidentally stringified
    # together with another query string or a small wrapper object.
    if ($value -match '(?i)(job_[0-9]{8}_[0-9]{6}_[0-9a-f]{8})') {
        return $matches[1].ToLowerInvariant()
    }
    return ''
}

function Get-RenderJobCancellationPath([string]$Language, [string]$JobId) {
    $normalizedJobId = Normalize-RenderJobId $JobId
    if ([string]::IsNullOrWhiteSpace($normalizedJobId)) { throw 'PDF作成ジョブの情報を受け取れませんでした。' }
    return (Join-Path (Get-RenderJobDir $Language) "$normalizedJobId.cancel.json")
}

function Test-RenderJobCancellationRequested([string]$Language, [string]$JobId) {
    try { return (Test-Path -LiteralPath (Get-RenderJobCancellationPath $Language $JobId) -PathType Leaf) } catch { return $false }
}

function Request-RenderJobCancellation([string]$Language, [string]$JobId) {
    $normalizedJobId = Normalize-RenderJobId $JobId
    if ([string]::IsNullOrWhiteSpace($normalizedJobId)) { throw '中止するPDF作成ジョブを確認できませんでした。' }
    $jobDir = Get-RenderJobDir $Language
    $statusPath = Join-Path $jobDir "$normalizedJobId.status.json"
    $status = Read-JsonFile $statusPath $null
    if ($null -eq $status) { throw '中止するPDF作成ジョブが見つかりません。' }
    $terminal = @('completed','completed-with-errors','failed','missing','cancelled')
    $statusText = ([string](Get-DataProperty $status 'status' '')).ToLowerInvariant()
    if ($terminal -contains $statusText) {
        return [pscustomobject][ordered]@{ ok=$true; jobId=$normalizedJobId; accepted=$false; alreadyFinished=$true; status=$statusText; message='PDF作成はすでに終了しています。' }
    }
    $requestedAt = New-NowIso
    Write-JsonFile (Get-RenderJobCancellationPath $Language $normalizedJobId) ([ordered]@{ schemaVersion=1; jobId=$normalizedJobId; requestedAt=$requestedAt })
    return [pscustomobject][ordered]@{ ok=$true; jobId=$normalizedJobId; accepted=$true; alreadyFinished=$false; status=$statusText; cancelRequested=$true; requestedAt=$requestedAt; message='中止を受け付けました。現在処理中の原稿が終わると停止します。' }
}

function Set-RenderJobCancelledStatus([string]$StatusPath, $Status) {
    $total = [Math]::Max(0, (Get-IntDataProperty $Status 'total' 0))
    $processed = [Math]::Max(0, (Get-IntDataProperty $Status 'completed' 0) + (Get-IntDataProperty $Status 'failed' 0))
    $percent = if ($total -le 0) { 100 } else { [int][Math]::Min(100, [Math]::Floor(($processed / [double]$total) * 100)) }
    Set-NoteProperty $Status 'status' 'cancelled'
    Set-NoteProperty $Status 'cancelRequested' $true
    Set-NoteProperty $Status 'cancelledAt' (New-NowIso)
    Set-NoteProperty $Status 'percent' $percent
    Set-NoteProperty $Status 'currentWorkbookId' ''
    Set-NoteProperty $Status 'currentWorkbookName' ''
    Set-NoteProperty $Status 'currentSheet' ''
    $summary = if ($processed -gt 0) { "$processed / $total 件を処理した時点でPDF作成を中止しました。" } else { 'PDF作成を開始前に中止しました。' }
    Set-NoteProperty $Status 'message' $summary
    Write-RenderJobStatus $StatusPath $Status
    return $Status
}

function Read-RenderJobStatus([string]$Language, [string]$JobId) {
    $normalizedJobId = Normalize-RenderJobId $JobId
    if ([string]::IsNullOrWhiteSpace($normalizedJobId)) {
        throw 'PDF作成ジョブの情報を受け取れませんでした。もう一度「PDF作成」を押してください。'
    }
    $jobDir = Get-RenderJobDir $Language
    $path = Join-Path $jobDir "$normalizedJobId.status.json"
    for ($attempt = 0; $attempt -lt 8; $attempt++) {
        if (Test-Path -LiteralPath $path) {
            try {
                $job = Read-JsonFile $path $null
                if ($null -ne $job) {
                    try {
                        $statusText = ([string]$job.status).ToLowerInvariant()
                        $terminal = @('completed','completed-with-errors','failed','missing','cancelled')
                        if ($terminal -notcontains $statusText) {
                            $item = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
                            $ageSeconds = if ($item) { ([DateTime]::UtcNow - $item.LastWriteTimeUtc).TotalSeconds } else { 0 }
                            $processId = Get-IntDataProperty $job 'processId' 0
                            if ($processId -gt 0) {
                                $alive = $false
                                try { $alive = $null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue) } catch { $alive = $false }
                                if ((-not $alive) -and $ageSeconds -gt 10) {
                                    $job = Set-RenderJobFailedFromStartupProblem $path $job 'PDF作成プロセスが終了していたため停止しました。もう一度PDF作成を押してください。'
                                } elseif ($alive -and ($statusText -eq 'queued' -or $statusText -eq 'launching') -and $ageSeconds -gt 90) {
                                    # A healthy child process rewrites queued/launching -> running. Give PowerShell/Excel startup
                                    # enough room, but do not leave the UI at the initial percent forever.
                                    $job = Set-RenderJobFailedFromStartupProblem $path $job 'PDF作成プロセスは起動しましたが、ジョブ処理に入れませんでした。Excelを閉じてから、もう一度PDF作成を押してください。'
                                }
                            } elseif ($ageSeconds -gt 10) {
                                $job = Set-RenderJobFailedFromStartupProblem $path $job 'PDF作成ジョブが中断されていました。もう一度PDF作成を押してください。'
                            }
                        }
                    } catch { }
                    $jobStatus = ([string](Get-DataProperty $job 'status' '')).ToLowerInvariant()
                    if (@('completed','completed-with-errors','failed','missing','cancelled') -notcontains $jobStatus -and (Test-RenderJobCancellationRequested $Language $normalizedJobId)) {
                        Set-NoteProperty $job 'cancelRequested' $true
                        Set-NoteProperty $job 'message' '中止を受け付けました。現在処理中の原稿が終わると停止します。'
                    }
                    return $job
                }
            } catch [System.IO.IOException] {
                Start-Sleep -Milliseconds (70 + (35 * $attempt))
            } catch [System.UnauthorizedAccessException] {
                Start-Sleep -Milliseconds (70 + (35 * $attempt))
            }
        } else {
            Start-Sleep -Milliseconds (70 + (35 * $attempt))
        }
    }

    $inputPath = Join-Path $jobDir "$normalizedJobId.input.json"
    $total = 0
    try {
        $input = Read-JsonFile $inputPath $null
        if ($input -and $input.workbookIds) { $total = @(Get-Array $input.workbookIds).Count }
    } catch { }
    return [ordered]@{
        ok = $true
        jobId = $normalizedJobId
        status = 'queued'
        total = $total
        completed = 0
        failed = 0
        percent = 0
        message = 'PDF作成ジョブを準備しています。'
        currentWorkbookId = ''
        currentWorkbookName = ''
        currentSheet = ''
        results = @()
        errors = @()
        updatedAt = 'waiting-for-status-file'
        stateSavedAt = ''
        transientMissing = $true
    }
}

function Get-WorkbookDisplayForStatus([string]$Language, [string]$WorkbookId) {
    try {
        $s = Get-Structure $Language
        $w = @(Get-Array $s.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
        if ($w.Count -gt 0) {
            $name = [string]$w[0].displayName
            if ([string]::IsNullOrWhiteSpace($name)) { $name = [string]$w[0].fileName }
            if (-not [string]::IsNullOrWhiteSpace($name)) { return $name }
        }
    } catch { }
    return $WorkbookId
}


function Get-ActiveRenderJobStatus([string]$Language) {
    try {
        $dir = Get-RenderJobDir $Language
        if (-not (Test-Path -LiteralPath $dir)) { return $null }
        $terminal = @('completed','completed-with-errors','failed','missing','cancelled')
        $cutoff = [DateTime]::UtcNow.AddHours(-4)
        $now = [DateTime]::UtcNow
        foreach ($file in @(Get-ChildItem -LiteralPath $dir -Filter '*.status.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 20)) {
            if ($file.LastWriteTimeUtc -lt $cutoff) { continue }
            $job = Read-JsonFile $file.FullName $null
            if ($null -eq $job) { continue }
            $status = ([string]$job.status).ToLowerInvariant()
            if ([string]::IsNullOrWhiteSpace($status)) { continue }
            if ($terminal -contains $status) { continue }

            $ageSeconds = ($now - $file.LastWriteTimeUtc).TotalSeconds
            $processId = Get-IntDataProperty $job 'processId' 0
            if ($processId -gt 0) {
                $alive = $false
                try { $alive = $null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue) } catch { $alive = $false }
                if ($alive) {
                    if (($status -eq 'queued' -or $status -eq 'launching') -and $ageSeconds -gt 90) {
                        [void](Set-RenderJobFailedFromStartupProblem $file.FullName $job '前回のPDF作成プロセスがジョブ処理に入らず停止扱いになりました。新しくPDF作成できます。')
                        continue
                    }
                    return $job
                }
                # Immediately after Start-Process there can be a short race before the child process is visible.
                if ($ageSeconds -lt 10) { return $job }
                [void](Set-RenderJobFailedFromStartupProblem $file.FullName $job '前回のPDF作成プロセスが終了していたため、新しいPDF作成を開始できます。')
                continue
            }

            # Older builds did not write processId. Do not let those stale queued/running files block the PDF button.
            if ($ageSeconds -lt 5) { return $job }
            [void](Set-RenderJobFailedFromStartupProblem $file.FullName $job '古いPDF作成ジョブを終了扱いにしました。もう一度PDF作成できます。')
        }
    } catch { }
    return $null
}

function Start-HiddenPowerShellChild([string]$PowerShellExe, [string]$Command, [string]$StdoutPath, [string]$StderrPath) {
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Command))
    $args = @('-NoProfile','-STA','-NonInteractive','-ExecutionPolicy','Bypass','-EncodedCommand',$encodedCommand) -join ' '
    try {
        return (Start-Process -FilePath $PowerShellExe -ArgumentList $args -WindowStyle Hidden -RedirectStandardOutput $StdoutPath -RedirectStandardError $StderrPath -PassThru)
    } catch {
        # Some managed launch environments expose both Path and PATH. Windows
        # PowerShell's Start-Process tries to copy them into a case-insensitive
        # dictionary and fails before the child starts. ShellExecute avoids that
        # enumeration; redirect inside the encoded child command instead.
        $outEscaped = ([IO.Path]::GetFullPath($StdoutPath)).Replace("'", "''")
        $errEscaped = ([IO.Path]::GetFullPath($StderrPath)).Replace("'", "''")
        $redirected = "& { $Command } 1>> '$outEscaped' 2>> '$errEscaped'"
        $encodedFallback = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($redirected))
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $PowerShellExe
        $psi.Arguments = @('-NoProfile','-STA','-NonInteractive','-ExecutionPolicy','Bypass','-EncodedCommand',$encodedFallback) -join ' '
        $psi.WorkingDirectory = $Script:AppRoot
        $psi.UseShellExecute = $true
        $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        return [System.Diagnostics.Process]::Start($psi)
    }
}

# ---- 提出用PDFの出力ジョブ -------------------------------------------------
# HTTPサーバーはリクエストを直列に処理するため、出力を要求の中で完結させると
# その間の進捗ポーリングも中止要求も受け付けられない。変換PDFと同じく子プロセスへ
# 出し、状態ファイルをブラウザが読む方式に揃える。
$Script:FinalBuildProgress = $null

function Report-FinalBuildPhase([string]$Message) {
    if ($null -eq $Script:FinalBuildProgress) { return }
    try { & $Script:FinalBuildProgress $Message } catch { }
}

function Normalize-FinalJobId([string]$JobId) {
    $value = ([string]$JobId).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { return '' }
    try { $value = [Uri]::UnescapeDataString($value).Trim() } catch { }
    if ($value -match '(?i)(final_[0-9]{8}_[0-9]{6}_[0-9a-f]{8})') { return $matches[1].ToLowerInvariant() }
    return ''
}

function Get-FinalJobStatusPath([string]$Language, [string]$JobId) {
    $normalized = Normalize-FinalJobId $JobId
    if ([string]::IsNullOrWhiteSpace($normalized)) { throw '提出用PDFの出力ジョブを確認できませんでした。' }
    return (Join-Path (Get-RenderJobDir $Language) "$normalized.status.json")
}

function Get-FinalJobCancellationPath([string]$Language, [string]$JobId) {
    $normalized = Normalize-FinalJobId $JobId
    if ([string]::IsNullOrWhiteSpace($normalized)) { throw '中止する出力ジョブを確認できませんでした。' }
    return (Join-Path (Get-RenderJobDir $Language) "$normalized.cancel.json")
}

function Test-FinalJobCancellationRequested([string]$Language, [string]$JobId) {
    try { return (Test-Path -LiteralPath (Get-FinalJobCancellationPath $Language $JobId) -PathType Leaf) } catch { return $false }
}

function Request-FinalJobCancellation([string]$Language, [string]$JobId) {
    $normalized = Normalize-FinalJobId $JobId
    $statusPath = Get-FinalJobStatusPath $Language $normalized
    $status = Read-JsonFile $statusPath $null
    if ($null -eq $status) { throw '中止する出力ジョブが見つかりません。' }
    $statusText = ([string](Get-DataProperty $status 'status' '')).ToLowerInvariant()
    if (@('completed','completed-with-errors','failed','cancelled') -contains $statusText) {
        return [pscustomobject][ordered]@{ ok=$true; jobId=$normalized; accepted=$false; alreadyFinished=$true; status=$statusText; message='提出用PDFの出力はすでに終了しています。' }
    }
    Write-JsonFile (Get-FinalJobCancellationPath $Language $normalized) ([ordered]@{ schemaVersion=1; jobId=$normalized; requestedAt=(New-NowIso) })
    return [pscustomobject][ordered]@{ ok=$true; jobId=$normalized; accepted=$true; alreadyFinished=$false; status=$statusText; cancelRequested=$true
        message='中止を受け付けました。作成中の1冊は最後まで書き上げてから停止します。' }
}

function Read-FinalJobStatus([string]$Language, [string]$JobId) {
    $normalized = Normalize-FinalJobId $JobId
    $statusPath = Get-FinalJobStatusPath $Language $normalized
    $status = Read-JsonFile $statusPath $null
    if ($null -eq $status) {
        return [pscustomobject][ordered]@{ ok=$true; jobId=$normalized; status='missing'; percent=0; total=0; completed=0; failed=0; message='出力の状態を読み取れませんでした。' }
    }
    $statusText = ([string](Get-DataProperty $status 'status' '')).ToLowerInvariant()
    if (@('completed','completed-with-errors','failed','cancelled') -notcontains $statusText) {
        # 子プロセスが落ちたまま「実行中」で残ると、以後の出力が永久に始められない。
        $processId = Get-IntDataProperty $status 'processId' 0
        $alive = $false
        if ($processId -gt 0) { try { $alive = ($null -ne (Get-Process -Id $processId -ErrorAction Stop)) } catch { $alive = $false } }
        if (-not $alive -and $processId -gt 0) {
            Set-NoteProperty $status 'status' 'failed'
            Set-NoteProperty $status 'percent' 100
            Set-NoteProperty $status 'message' '出力プロセスが予期せず終了しました。もう一度出力してください。'
            Write-RenderJobStatus $statusPath $status
        }
    }
    if ($null -eq $status.PSObject.Properties['ok']) { Set-NoteProperty $status 'ok' $true }
    return ([pscustomobject]$status)
}

function Get-ActiveFinalJobStatus([string]$Language) {
    $dir = Get-RenderJobDir $Language
    $cutoff = (Get-Date).ToUniversalTime().AddHours(-4)
    foreach ($file in @(Get-ChildItem -LiteralPath $dir -Filter 'final_*.status.json' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending)) {
        if ($file.LastWriteTimeUtc -lt $cutoff) { continue }
        $status = Read-JsonFile $file.FullName $null
        if ($null -eq $status) { continue }
        $statusText = ([string](Get-DataProperty $status 'status' '')).ToLowerInvariant()
        if (@('completed','completed-with-errors','failed','cancelled') -contains $statusText) { continue }
        $refreshed = Read-FinalJobStatus $Language ([string](Get-DataProperty $status 'jobId' ''))
        $refreshedText = ([string](Get-DataProperty $refreshed 'status' '')).ToLowerInvariant()
        if (@('completed','completed-with-errors','failed','cancelled','missing') -contains $refreshedText) { continue }
        return $refreshed
    }
    return $null
}

function Start-FinalBuildJob([string]$Language, [string]$PackId, [string[]]$TargetIds) {
    $active = Get-ActiveFinalJobStatus $Language
    if ($null -ne $active) {
        Set-NoteProperty $active 'message' ('提出用PDFを作成中です。 ' + [string]$active.message)
        return $active
    }
    $structure = Get-Structure $Language
    $scope = Resolve-DocumentPackScope $structure $PackId $false
    $allowed = @(Get-PackTargetIds $Language $scope.pack)
    $requested = @($TargetIds | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
    if ($requested.Count -eq 0) { $requested = @($allowed[0]) }
    foreach ($targetId in $requested) {
        if ($allowed -notcontains $targetId) { throw [ArgumentException]::new('この一式に存在しない出力先です。') }
    }
    $jobId = 'final_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '_' + ([Guid]::NewGuid().ToString('N').Substring(0,8))
    $jobDir = Get-RenderJobDir $Language
    $inputPath = Join-Path $jobDir "$jobId.input.json"
    $statusPath = Join-Path $jobDir "$jobId.status.json"
    $stdoutPath = Join-Path $jobDir "$jobId.out.log"
    $stderrPath = Join-Path $jobDir "$jobId.err.log"
    $initial = [pscustomobject][ordered]@{
        ok=$true; jobId=$jobId; kind='final-build'; status='launching'; total=$requested.Count; completed=0; failed=0; percent=1
        message='提出用PDFの作成を開始します。'; currentTargetId=''; currentTargetName=''; phase=''
        packId=[string]$scope.packId; targetIds=@($requested); built=@(); skipped=@(); errors=@()
        processId=0; stdoutPath=$stdoutPath; stderrPath=$stderrPath; startedAt=(New-NowIso); updatedAt=(New-NowIso)
    }
    Write-JsonFile $statusPath $initial
    Write-JsonFile $inputPath ([ordered]@{ jobId=$jobId; mode=$Language; packId=[string]$scope.packId; targetIds=@($requested); statusPath=$statusPath; stdoutPath=$stdoutPath; stderrPath=$stderrPath })
    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
    $script = Join-Path $Script:AppRoot 'server.ps1'
    $jobCommand = "& '$($script.Replace("'", "''"))' -Mode '$($Language.Replace("'", "''"))' -FinalJobPath '$($inputPath.Replace("'", "''"))'"
    try {
        $proc = Start-HiddenPowerShellChild $psExe $jobCommand $stdoutPath $stderrPath
        if ($proc -and $proc.Id) {
            Set-NoteProperty $initial 'processId' ([int]$proc.Id)
            Set-NoteProperty $initial 'message' '出力プロセスを起動しました。準備しています。'
            Set-NoteProperty $initial 'percent' 2
            Write-RenderJobStatus $statusPath $initial
        }
    } catch {
        Set-NoteProperty $initial 'status' 'failed'
        Set-NoteProperty $initial 'percent' 100
        Set-NoteProperty $initial 'message' ("出力プロセスを起動できませんでした: " + $_.Exception.Message)
        Write-RenderJobStatus $statusPath $initial
        throw
    }
    return $initial
}

function Invoke-FinalBuildJobFromFile([string]$JobPath) {
    $job = Read-JsonFile $JobPath $null
    if ($null -eq $job) { throw "Final build job file is not readable: $JobPath" }
    $language = [string]$job.mode
    if ([string]::IsNullOrWhiteSpace($language)) { $language = $Mode }
    $jobId = [string]$job.jobId
    $packId = [string]$job.packId
    $statusPath = [string]$job.statusPath
    $targetIds = @(Get-Array $job.targetIds | ForEach-Object { [string]$_ } | Where-Object { $_ })
    $processId = 0
    try { $processId = [System.Diagnostics.Process]::GetCurrentProcess().Id } catch { $processId = 0 }
    $status = Read-JsonFile $statusPath $null
    if ($null -eq $status) { $status = [pscustomobject][ordered]@{ ok=$true; jobId=$jobId; kind='final-build' } }
    Set-NoteProperty $status 'status' 'running'
    Set-NoteProperty $status 'processId' $processId
    Set-NoteProperty $status 'total' $targetIds.Count
    Set-NoteProperty $status 'completed' 0
    Set-NoteProperty $status 'failed' 0
    Set-NoteProperty $status 'percent' 3
    Set-NoteProperty $status 'message' '出力条件を確認しています。'
    Write-RenderJobStatus $statusPath $status

    $structure = Get-Structure $language
    $scope = Resolve-DocumentPackScope $structure $packId $false
    $built = @(); $skipped = @(); $errors = @()
    $index = 0
    $cancelled = $false

    # 組み込みパックの一括出力は準トランザクションAPIが「本体だけ成功する」状態を防ぐ。
    # 進捗のために1冊ずつのループへ置き換えると、その保証が失われる。
    $allTargets = @(Get-PackTargetIds $language $scope.pack)
    $isTransactionalAll = ([bool]$scope.builtIn) -and ($targetIds.Count -gt 1) -and
        (@($allTargets | Where-Object { $targetIds -notcontains $_ }).Count -eq 0)
    if ($isTransactionalAll) {
        Set-NoteProperty $status 'phase' '一括出力'
        Set-NoteProperty $status 'percent' 10
        Set-NoteProperty $status 'message' 'すべての提出用PDFをまとめて作成しています。途中で分かれた状態にならないよう、一度に書き出します。'
        Write-RenderJobStatus $statusPath $status
        $Script:FinalBuildProgress = {
            param($phaseMessage)
            Set-NoteProperty $status 'phase' ([string]$phaseMessage)
            Set-NoteProperty $status 'message' ('まとめて出力中 : ' + [string]$phaseMessage)
            Write-RenderJobStatus $statusPath $status
        }.GetNewClosure()
        try {
            $volumes = @($targetIds | ForEach-Object { Get-LegacyVolumeFromTargetId $language $_ })
            $result = Invoke-FinalBuildTransaction $language ([string]$scope.category) $volumes
            $built = @(Get-Array (Get-DataProperty $result 'built' @()))
            $skipped = @(Get-Array (Get-DataProperty $result 'skipped' @()))
            Set-NoteProperty $status 'completed' $targetIds.Count
        } catch {
            $message = [string]$_.Exception.Message
            $errors += [ordered]@{ targetId=''; targetName='まとめて出力'; error=$message; userError=(ConvertTo-UserRenderError $message) }
            Set-NoteProperty $status 'failed' $targetIds.Count
        } finally {
            $Script:FinalBuildProgress = $null
        }
        $index = $targetIds.Count
        $targetIds = @()
    }

    # 組版の待ちからも中止要求が見えるようにする（上限は15分）。
    $Script:NativeCancelProbe = { Test-FinalJobCancellationRequested $language $jobId }.GetNewClosure()
    try {
    foreach ($targetId in $targetIds) {
        if (Test-FinalJobCancellationRequested $language $jobId) { $cancelled = $true; break }
        $displayName = [string](Get-PackTargetDisplayName $language $scope.pack $targetId)
        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $targetId }
        $basePercent = if ($targetIds.Count -le 0) { 100 } else { [int](3 + [Math]::Floor(($index / [double]$targetIds.Count) * 94)) }
        Set-NoteProperty $status 'currentTargetId' $targetId
        Set-NoteProperty $status 'currentTargetName' $displayName
        Set-NoteProperty $status 'percent' $basePercent
        Set-NoteProperty $status 'phase' '準備'
        Set-NoteProperty $status 'message' "$displayName を準備しています。"
        Write-RenderJobStatus $statusPath $status
        # 出力の各段階を画面へ返す。1冊の作成でも無反応な時間が生まれないようにする。
        $Script:FinalBuildProgress = {
            param($phaseMessage)
            Set-NoteProperty $status 'phase' ([string]$phaseMessage)
            Set-NoteProperty $status 'message' ("$displayName : " + [string]$phaseMessage)
            Write-RenderJobStatus $statusPath $status
        }.GetNewClosure()
        try {
            $readiness = Get-FinalBuildReadiness (Get-Structure $language) $language (Get-LegacyVolumeFromTargetId $language $targetId) $packId
            if ([int]$readiness.pageCount -le 0) { $skipped += $targetId }
            else { $built += @(Build-DocumentPackPdf $language $packId $targetId) }
            Set-NoteProperty $status 'completed' ([int](Get-IntDataProperty $status 'completed' 0) + 1)
        } catch {
            $message = [string]$_.Exception.Message
            $errors += [ordered]@{ targetId=$targetId; targetName=$displayName; error=$message; userError=(ConvertTo-UserRenderError $message) }
            Set-NoteProperty $status 'failed' ([int](Get-IntDataProperty $status 'failed' 0) + 1)
        } finally {
            $Script:FinalBuildProgress = $null
        }
        Set-NoteProperty $status 'built' @($built)
        Set-NoteProperty $status 'skipped' @($skipped)
        Set-NoteProperty $status 'errors' @($errors)
        Write-RenderJobStatus $statusPath $status
        $index++
    }
    $remaining = @($targetIds | Select-Object -Skip $index)
    Set-NoteProperty $status 'currentTargetId' ''
    Set-NoteProperty $status 'currentTargetName' ''
    Set-NoteProperty $status 'phase' ''
    Set-NoteProperty $status 'percent' 100
    Set-NoteProperty $status 'built' @($built)
    } finally {
        $Script:NativeCancelProbe = $null
    }
    Set-NoteProperty $status 'skipped' @($skipped)
    Set-NoteProperty $status 'errors' @($errors)
    if ($cancelled) {
        Set-NoteProperty $status 'status' 'cancelled'
        Set-NoteProperty $status 'cancelRequested' $true
        Set-NoteProperty $status 'message' ("$($built.Count) 冊を出力した時点で中止しました。残り $($remaining.Count) 冊は作成していません。")
    } elseif ($errors.Count -gt 0) {
        Set-NoteProperty $status 'status' 'completed-with-errors'
        Set-NoteProperty $status 'message' ("$($built.Count) 冊を出力しました。$($errors.Count) 冊は出力できませんでした。")
    } else {
        Set-NoteProperty $status 'status' 'completed'
        $summary = if ($built.Count -gt 0) { "$($built.Count) 冊の提出用PDFを出力しました。" } else { '出力対象のページがありませんでした。' }
        if ($skipped.Count -gt 0) { $summary += " $($skipped.Count) 冊はページが無いため作成していません。" }
        Set-NoteProperty $status 'message' $summary
    }
    Set-NoteProperty $status 'finishedAt' (New-NowIso)
    Write-RenderJobStatus $statusPath $status
    try { Remove-Item -LiteralPath (Get-FinalJobCancellationPath $language $jobId) -Force -ErrorAction SilentlyContinue } catch { }
}

function Start-RenderJob([string]$Language, [string[]]$WorkbookIds, [bool]$OnlyUpdated, [string]$Category = '', $SnapshotPins = $null, [string]$ExcelSheetSelection = '') {
    # Return a job immediately. Expensive update scanning runs inside the background job,
    # so the PDF button does not appear to do nothing on large folders.
    $activeJob = Get-ActiveRenderJobStatus $Language
    if ($null -ne $activeJob) {
        $activeMsg = [string]$activeJob.message
        if ([string]::IsNullOrWhiteSpace($activeMsg)) { $activeMsg = 'PDF作成中です。' }
        else { $activeMsg = "PDF作成中です。 $activeMsg" }
        Set-NoteProperty $activeJob 'message' $activeMsg
        if ($null -eq $activeJob.PSObject.Properties['ok']) { Set-NoteProperty $activeJob 'ok' $true }
        return ([pscustomobject]$activeJob)
    }
    $structure = Get-Structure $Language
    $categoryNormalized = Normalize-WorkbookCategory $Category ''
    $explicitIds = @($WorkbookIds | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    $ids = @()
    if ($explicitIds.Count -gt 0) {
        $known = @(Get-Array $structure.workbooks | Where-Object {
            $id = [string]$_.workbookId
            ($explicitIds -contains $id) -and ([string]$_.status -ne 'missing') -and ([string]::IsNullOrWhiteSpace($categoryNormalized) -or (Test-WorkbookCategory $_ $categoryNormalized))
        } | ForEach-Object { [string]$_.workbookId })
        $ids = @($known | Select-Object -Unique)
    } elseif (-not $OnlyUpdated) {
        $ids = @(Get-Array $structure.workbooks | Where-Object {
            ([string]$_.status -ne 'missing') -and ([string]::IsNullOrWhiteSpace($categoryNormalized) -or (Test-WorkbookCategory $_ $categoryNormalized))
        } | ForEach-Object { [string]$_.workbookId })
    }

    $registeredCount = @(Get-Array $structure.workbooks | Where-Object { ([string]$_.status -ne 'missing') -and ([string]::IsNullOrWhiteSpace($categoryNormalized) -or (Test-WorkbookCategory $_ $categoryNormalized)) }).Count
    if ($registeredCount -eq 0) { $OnlyUpdated = $false }

    # Pin the per-source selection at queue time. A later settings change must not
    # silently alter an already queued job; an empty argument means "use the
    # persisted source setting" for backwards-compatible callers.
    $sheetSelectionByWorkbook = [ordered]@{}
    $selectionPinIds = @($ids)
    if ($selectionPinIds.Count -eq 0 -and $OnlyUpdated) {
        $selectionPinIds = @(Get-Array $structure.workbooks | Where-Object { [string]$_.status -ne 'missing' -and ([string]::IsNullOrWhiteSpace($categoryNormalized) -or (Test-WorkbookCategory $_ $categoryNormalized)) } | ForEach-Object { [string]$_.workbookId })
    }
    foreach ($id in @($selectionPinIds)) {
        $source = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$id } | Select-Object -First 1)
        if ($source.Count -eq 0) { continue }
        $sourceType = ([string](Get-DataProperty $source[0] 'sourceType' 'excel')).ToLowerInvariant()
        if ($sourceType -eq 'excel') {
            $rawSelection = if (-not [string]::IsNullOrWhiteSpace($ExcelSheetSelection)) { $ExcelSheetSelection } else { [string](Get-DataProperty $source[0] 'excelSheetSelectionMode' 'all-visible') }
            $sheetSelectionByWorkbook[[string]$id] = Normalize-ExcelSheetSelection $rawSelection
        } else {
            if (-not [string]::IsNullOrWhiteSpace($ExcelSheetSelection) -and (Normalize-ExcelSheetSelection $ExcelSheetSelection) -eq 'numeric-only') { throw '半角数字シート指定はExcel原稿でのみ利用できます。' }
            $sheetSelectionByWorkbook[[string]$id] = 'all-visible'
        }
    }

    $jobId = 'job_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '_' + ([Guid]::NewGuid().ToString('N').Substring(0,8))
    $jobDir = Get-RenderJobDir $Language
    $inputPath = Join-Path $jobDir "$jobId.input.json"
    $statusPath = Join-Path $jobDir "$jobId.status.json"
    $stdoutPath = Join-Path $jobDir "$jobId.out.log"
    $stderrPath = Join-Path $jobDir "$jobId.err.log"
    $initialTotal = if ($OnlyUpdated -and $ids.Count -eq 0) { $registeredCount } else { $ids.Count }
    $initialMessage = '変換PDFの作成が必要な原稿はありません。'
    if ($registeredCount -eq 0) { $initialMessage = '登録済み原稿がありません。' }
    elseif ($OnlyUpdated -and $ids.Count -eq 0) { $initialMessage = 'PDF作成対象を確認しています。' }
    elseif ($ids.Count -gt 0) { $initialMessage = 'PDF作成を開始します。' }
    $initial = [pscustomobject][ordered]@{
        ok = $true; jobId = $jobId; status = 'queued'; total = $initialTotal; completed = 0; failed = 0; percent = 1;
        message = $initialMessage;
        currentWorkbookId = ''; currentWorkbookName = ''; currentSheet = ''; processId = 0; stdoutPath = $stdoutPath; stderrPath = $stderrPath; results = @(); errors = @(); startedAt = New-NowIso; updatedAt = New-NowIso; stateSavedAt = ''
    }
    Write-JsonFile $statusPath $initial
    $pinsOut = [ordered]@{}
    if ($null -ne $SnapshotPins) { foreach ($k in @($SnapshotPins.Keys)) { $pinsOut[[string]$k] = $SnapshotPins[$k] } }
    $requestedSheetSelection = if ([string]::IsNullOrWhiteSpace($ExcelSheetSelection)) { '' } else { Normalize-ExcelSheetSelection $ExcelSheetSelection }
    Write-JsonFile $inputPath ([ordered]@{ jobId = $jobId; mode = $Language; workbookIds = @($ids); onlyUpdated = $OnlyUpdated; category = $categoryNormalized; snapshotPins = $pinsOut; sheetSelectionByWorkbook = $sheetSelectionByWorkbook; excelSheetSelectionMode = $requestedSheetSelection; statusPath = $statusPath; stdoutPath = $stdoutPath; stderrPath = $stderrPath })
    if ($registeredCount -eq 0 -or ($ids.Count -eq 0 -and -not $OnlyUpdated)) {
        $initial.status = 'completed'; $initial.percent = 100
        if ($registeredCount -eq 0) { $initial.message = '登録済み原稿がありません。' } else { $initial.message = '変換PDFの作成が必要な原稿はありません。' }
        Write-RenderJobStatus $statusPath $initial
        return ([pscustomobject]$initial)
    }

    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
    $script = Join-Path $Script:AppRoot 'server.ps1'
    # Use -EncodedCommand instead of a quoted -File command line. This avoids Windows quoting edge cases
    # where the child PowerShell can start without binding -RenderJobPath, leaving the UI stuck at 0%.
    $jobCommand = "& '$($script.Replace("'", "''"))' -Mode '$($Language.Replace("'", "''"))' -RenderJobPath '$($inputPath.Replace("'", "''"))'"
    Set-NoteProperty $initial 'launchCommandKind' 'EncodedCommand-STA'
    Set-NoteProperty $initial 'status' 'launching'
    Set-NoteProperty $initial 'message' 'PDF作成プロセスを起動しています。'
    Write-RenderJobStatus $statusPath $initial
    try {
        $proc = Start-HiddenPowerShellChild $psExe $jobCommand $stdoutPath $stderrPath
        if ($proc -and $proc.Id) {
            $initial.processId = [int]$proc.Id
            $statusToUpdate = $initial
            try {
                $existingStatus = Read-JsonFile $statusPath $null
                if ($null -ne $existingStatus) { $statusToUpdate = $existingStatus }
            } catch { }
            Set-NoteProperty $statusToUpdate 'processId' ([int]$proc.Id)
            Set-NoteProperty $statusToUpdate 'status' 'launching'
            Set-NoteProperty $statusToUpdate 'percent' ([Math]::Max(2, (Get-IntDataProperty $statusToUpdate 'percent' 0)))
            Set-NoteProperty $statusToUpdate 'message' 'PDF作成プロセスを起動しました。原稿を準備しています。'
            Write-RenderJobStatus $statusPath $statusToUpdate
        }
    } catch {
        $initial.status = 'failed'
        $initial.percent = 100
        $initial.message = "PDF作成プロセスを起動できませんでした: $($_.Exception.Message)"
        $initial.errors = @([ordered]@{ error = $_.Exception.Message; userError = (ConvertTo-UserRenderError $_.Exception.Message); detail = [string]$_ })
        Write-RenderJobStatus $statusPath $initial
        throw
    }
    return ([pscustomobject]$initial)
}

function Invoke-RenderJobFromFile([string]$JobPath) {
    $job = Read-JsonFile $JobPath $null
    if ($null -eq $job) { throw "Render job file is not readable: $JobPath" }
    $language = [string]$job.mode
    if ([string]::IsNullOrWhiteSpace($language)) { $language = $Mode }
    $ids = @(Get-Array $job.workbookIds | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $explicitIds = ($ids.Count -gt 0)
    $onlyUpdated = [bool](Get-DataProperty $job 'onlyUpdated' $false)
    $category = Normalize-WorkbookCategory ([string](Get-DataProperty $job 'category' '')) ''
    $sheetSelectionByWorkbook = Get-DataProperty $job 'sheetSelectionByWorkbook' $null
    $requestedSheetSelection = [string](Get-DataProperty $job 'excelSheetSelectionMode' '')
    $statusPath = [string]$job.statusPath
    $stdoutPath = [string](Get-DataProperty $job 'stdoutPath' '')
    $stderrPath = [string](Get-DataProperty $job 'stderrPath' '')
    $jobId = [string]$job.jobId
    $processId = 0
    try {
        $existingStatus = Read-JsonFile $statusPath $null
        if ($null -ne $existingStatus) { $processId = Get-IntDataProperty $existingStatus 'processId' 0 }
    } catch { }
    if ($processId -le 0) { try { $processId = [System.Diagnostics.Process]::GetCurrentProcess().Id } catch { $processId = 0 } }
    $total = $ids.Count
    $statusMessage = '原稿を準備しています。'
    if ($onlyUpdated -and -not $explicitIds) { $statusMessage = 'PDF作成対象を確認しています。' }
    $status = [pscustomobject][ordered]@{ ok = $true; jobId = $jobId; status = 'running'; total = $total; completed = 0; failed = 0; percent = 3; message = $statusMessage; currentWorkbookId = ''; currentWorkbookName = ''; currentSheet = ''; processId = $processId; stdoutPath = $stdoutPath; stderrPath = $stderrPath; results = @(); errors = @(); startedAt = New-NowIso; updatedAt = New-NowIso; stateSavedAt = '' }
    Write-RenderJobStatus $statusPath $status
    $excel = $null
    $excelStartupError = ''
    $renderLoopSucceeded = $false
    $jobCancelled = $false
    # 外部コマンドの待ちからも中止要求が見えるようにする。これが無いと、java の待ちに
    # 入っている間は上限(最大15分)まで中止が効かない。
    $Script:NativeCancelProbe = { Test-RenderJobCancellationRequested $language $jobId }.GetNewClosure()
    try {
        if (Test-RenderJobCancellationRequested $language $jobId) {
            [void](Set-RenderJobCancelledStatus $statusPath $status)
            return
        }
        if ($onlyUpdated -and -not $explicitIds) {
            $status.status = 'scanning'
            $status.message = '変換PDFの作成が必要な原稿を確認しています。'
            $status.percent = 5
            Write-RenderJobStatus $statusPath $status
            [void](Scan-Updates $language {
                param($scanIndex, $scanTotal, $scanWorkbookId, $scanName)
                $status.status = 'scanning'
                $status.currentWorkbookId = [string]$scanWorkbookId
                $status.currentWorkbookName = [string]$scanName
                $status.currentSheet = ''
                $status.total = [Math]::Max([int]$scanTotal, 1)
                $status.completed = [Math]::Max([int]$scanIndex - 1, 0)
                $status.failed = 0
                $status.percent = [int][Math]::Max(5, [Math]::Min(15, [Math]::Floor(([double]$scanIndex / [Math]::Max(1, [int]$scanTotal)) * 15)))
                $status.message = "PDF作成対象を確認しています: $scanIndex / $scanTotal 件目 $scanName"
                Write-RenderJobStatus $statusPath $status
            })
            $ids = @(Get-AutoRenderWorkbookIds $language @() $category)
            $total = $ids.Count
            $status.total = $total
            $status.completed = 0
            $status.failed = 0
            $status.percent = if ($total -gt 0) { [Math]::Max([int]$status.percent, 15) } else { 100 }
            if (Test-RenderJobCancellationRequested $language $jobId) {
                [void](Set-RenderJobCancelledStatus $statusPath $status)
                return
            }
            if ($total -eq 0) {
                $status.status = 'completed'
                $status.message = '変換PDFの作成が必要な原稿はありません。'
                Write-RenderJobStatus $statusPath $status
                return
            }
            $status.message = "$total 件のPDF作成を開始します。"
            Write-RenderJobStatus $statusPath $status
        } elseif ($total -eq 0) {
            if (Test-RenderJobCancellationRequested $language $jobId) {
                [void](Set-RenderJobCancelledStatus $statusPath $status)
                return
            }
            $status.status = 'completed'
            $status.percent = 100
            $status.message = '変換PDFの作成が必要な原稿はありません。'
            Write-RenderJobStatus $statusPath $status
            return
        }
        if (Test-RenderJobCancellationRequested $language $jobId) {
            [void](Set-RenderJobCancelledStatus $statusPath $status)
            return
        }
        $renderStructure = if ($total -gt 0) { Get-Structure $language } else { $null }
        $excelIds = @()
        if ($null -ne $renderStructure) {
            $excelIds = @($ids | Where-Object {
                $id = $_
                $source = @(Get-Array $renderStructure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$id } | Select-Object -First 1)
                $source.Count -gt 0 -and [string](Get-DataProperty $source[0] 'sourceType' 'excel') -eq 'excel'
            })
        }
        if ($excelIds.Count -gt 0) {
            Reset-RenderEnvironmentForJob   # V5-§6.4: ジョブ開始ごとに環境を取り直す
            $status.message = 'Excelを起動しています。'
            $status.percent = [Math]::Max([int]$status.percent, 8)
            Write-RenderJobStatus $statusPath $status
            try {
                $excel = New-ExcelApplicationForRender
                $status.message = 'Excelの起動が完了しました。PDF化を開始します。'
            } catch {
                $excelStartupError = $_.Exception.Message
                $status.message = 'Excelを起動できません。Excel原稿をエラーとして記録し、他形式の処理を続けます。'
            }
            $status.percent = [Math]::Max([int]$status.percent, 10)
            Write-RenderJobStatus $statusPath $status
        } elseif ($total -gt 0) {
            $hasWord = @($ids | Where-Object {
                $id = $_
                $source = @(Get-Array $renderStructure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$id } | Select-Object -First 1)
                $source.Count -gt 0 -and [string](Get-DataProperty $source[0] 'sourceType' 'excel') -eq 'word'
            }).Count -gt 0
            $hasPowerPoint = @($ids | Where-Object {
                $id = $_
                $source = @(Get-Array $renderStructure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$id } | Select-Object -First 1)
                $source.Count -gt 0 -and [string](Get-DataProperty $source[0] 'sourceType' 'excel') -eq 'powerpoint'
            }).Count -gt 0
            $status.message = $(if ($hasWord) { 'Word原稿をPDF変換しています。' } elseif ($hasPowerPoint) { 'PowerPoint原稿をPDF変換しています。' } else { 'PDF原稿をページ単位で取り込んでいます。' })
            $status.percent = [Math]::Max([int]$status.percent, 10)
            Write-RenderJobStatus $statusPath $status
        }
        $index = 0
        $deferredAnalyses = @()
        foreach ($id in $ids) {
            if (Test-RenderJobCancellationRequested $language $jobId) { $jobCancelled = $true; break }
            $index++
            $name = Get-WorkbookDisplayForStatus $language $id
            $jobSource = if ($null -ne $renderStructure) { @(Get-Array $renderStructure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$id } | Select-Object -First 1) } else { @() }
            $jobSourceType = if ($jobSource.Count -gt 0) { [string](Get-DataProperty $jobSource[0] 'sourceType' 'excel') } else { 'excel' }
            $unitLabel = if ($jobSourceType -eq 'powerpoint') { 'スライド' } elseif ($jobSourceType -in @('pdf','word')) { 'ページ' } else { 'シート' }
            $status.currentWorkbookId = $id
            $status.currentWorkbookName = $name
            $status.currentSheet = ''
            $status.message = "$index / $total 件目: $name をPDF化しています。"
            $status.percent = [int][Math]::Max(10, [Math]::Floor((($index - 1) / [Math]::Max(1, $total)) * 100))
            Write-RenderJobStatus $statusPath $status
            $callback = {
                param($stage, $workbookId, $sheetName)
                # 進捗が動いた印。この callback は open / sheet-setup / batch / split / sheet /
                # word / powerpoint のすべてで呼ばれるので、変換が進んでいる限り心拍が続く。
                # 監視プロセスが見るのはこの心拍であって、Excel を起動してからの経過ではない。
                Update-ExcelRenderHeartbeat
                $status.currentWorkbookId = [string]$workbookId
                $status.currentWorkbookName = $name
                $status.currentSheet = [string]$sheetName
                if ($stage -eq 'word') { $status.message = "$index / $total 件目: $name をWordでPDF変換しています。" }
                elseif ($stage -eq 'powerpoint') { $status.message = "$index / $total 件目: $name をPowerPointでPDF変換しています。" }
                elseif ($stage -eq 'open') { $status.message = "$index / $total 件目: $name を開いています。" }
                elseif ($stage -eq 'sheet-setup') { $status.message = "$index / $total 件目: $name / $unitLabel $sheetName の印刷設定を調整しています。" }
                elseif ($stage -eq 'batch') { $status.message = "$index / $total 件目: $name の複数シートをまとめてPDF化しています。" }
                elseif ($stage -eq 'split') { $status.message = "$index / $total 件目: $name のPDFを$unitLabel 単位に分けています。" }
                elseif ($stage -eq 'sheet') { $status.message = "$index / $total 件目: $name / $unitLabel $sheetName をPDF化しています。" }
                else { $status.message = "$index / $total 件目: $name を準備しています。" }
                $status.percent = [int][Math]::Floor((($index - 1 + 0.35) / [Math]::Max(1, $total)) * 100)
                Write-RenderJobStatus $statusPath $status
            }
            try {
                # V5-P0: 自動処理が固定した検知版があれば、その版でレンダリングする。
                $pin = Get-DataProperty (Get-DataProperty $job 'snapshotPins' $null) $id $null
                $pinSnapshot = [string](Get-DataProperty $pin 'snapshotId' '')
                $pinHash = [string](Get-DataProperty $pin 'expectedHash' '')
                $sheetSelection = [string](Get-DataProperty $sheetSelectionByWorkbook $id '')
                if ([string]::IsNullOrWhiteSpace($sheetSelection) -and -not [string]::IsNullOrWhiteSpace($requestedSheetSelection)) {
                    $sheetSelection = $requestedSheetSelection
                }
                if ([string]::IsNullOrWhiteSpace($sheetSelection) -and $jobSource.Count -gt 0 -and $jobSourceType -eq 'excel') {
                    $sheetSelection = [string](Get-DataProperty $jobSource[0] 'excelSheetSelectionMode' 'all-visible')
                }
                $r = Render-Source $language $id $excel $true $callback '' $pinSnapshot $pinHash $sheetSelection
                $deferredAnalyses += $Script:PendingAnalysis
                $Script:PendingAnalysis = $null
                $r['ok'] = $true
                $status.results = @($status.results) + @($r)
                $status.completed = [int]$status.completed + 1
            } catch {
                $msg = $_.Exception.Message
                $userMsg = ConvertTo-UserRenderError $msg
                $att = Get-LastRenderAttemptFor $id
                $errorDetail = Get-ErrorDetail $_
                Set-WorkbookRenderError $language $id $msg $errorDetail ([string]$att.snapshotId) ([string]$att.hash)
                $err = [ordered]@{ ok = $false; workbookId = $id; workbookName = $name; error = $msg; userError = $userMsg; detail = $errorDetail }
                $status.errors = @($status.errors) + @($err)
                $status.results = @($status.results) + @($err)
                $status.failed = [int]$status.failed + 1
            }
            $status.currentSheet = ''
            $status.percent = [int][Math]::Min(95, [Math]::Floor((([int]$status.completed + [int]$status.failed) / [Math]::Max(1, $total)) * 95))
            Write-RenderJobStatus $statusPath $status
            if (Test-RenderJobCancellationRequested $language $jobId) { $jobCancelled = $true; break }
        }
        if (Test-RenderJobCancellationRequested $language $jobId) { $jobCancelled = $true }
        # 比較画面が開ける「completed」は、比較情報まで保存し終えてから通知する。
        # 先に100%を返すと、ブラウザーが作成途中の共有ファイルを読み始めて停止して見える。
        $status.status = 'analyzing'
        $status.currentWorkbookId = ''
        $status.currentWorkbookName = ''
        $status.currentSheet = ''
        $status.percent = 96
        $status.message = '比較情報を準備しています。'
        Write-RenderJobStatus $statusPath $status
        $renderLoopSucceeded = $true
    } catch {
        $status.status = 'failed'
        $status.message = $_.Exception.Message
        $status.errors = @($status.errors) + @([ordered]@{ error = $_.Exception.Message; userError = (ConvertTo-UserRenderError $_.Exception.Message); detail = [string]$_ })
        Write-RenderJobStatus $statusPath $status
        throw
    } finally {
        $Script:NativeCancelProbe = $null
        Close-ExcelApplicationForRender $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # V5-P1: Excel を閉じ、レンダリングロックも解放してから解析する。
        # 比較用の再レンダリングが必要になっても、改めて共通ロックを取り直せる。
        $analysisItems = @($deferredAnalyses | Where-Object { $null -ne $_ })
        $analysisIndex = 0
        foreach ($pending in $analysisItems) {
            $analysisIndex++
            if ($renderLoopSucceeded) {
                $status.message = "比較情報を準備しています: $analysisIndex / $($analysisItems.Count) 件"
                $status.percent = [int][Math]::Min(99, 96 + [Math]::Floor(($analysisIndex / [Math]::Max(1, $analysisItems.Count)) * 3))
                Write-RenderJobStatus $statusPath $status
            }
            try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered ([string](Get-DataProperty $pending 'sheetSelectionMode' 'all-visible'))) }
            catch { Write-Warning ('画像ハッシュの解析に失敗しました: ' + $_.Exception.Message) }
        }
    }
    if (Test-RenderJobCancellationRequested $language $jobId) { $jobCancelled = $true }
    if ($renderLoopSucceeded) {
        if ($jobCancelled) {
            [void](Set-RenderJobCancelledStatus $statusPath $status)
            return
        } elseif ([int]$status.failed -gt 0) {
            $status.status = 'completed-with-errors'
            $status.message = "$($status.failed) 件でエラーが発生しました。赤いエラー表示を確認してください。"
        } else {
            $status.status = 'completed'
            $status.message = "$($status.completed) 件のPDF作成が完了しました。"
        }
        $status.percent = 100
        Set-NoteProperty $status 'stateSavedAt' (New-NowIso)
        Write-RenderJobStatus $statusPath $status
    }
}


function Reorder-Pages([string]$Language, $Body) {
    $requestedScope = [string](Get-DataProperty $Body 'packId' (Get-DataProperty $Body 'category' ''))
    $baseLayout = Get-RequestedBaseLayout $Body
    return Update-StructureLocked $Language {
        param($structure)
        $cat = [string](Resolve-DocumentPackScope $structure $requestedScope $false).packId
        $allowed = @(Get-PackVolumeList $Language (Resolve-DocumentPackScope $structure $requestedScope $false).pack $true)
        $volumeObject = $Body.volumes
        if ($null -eq $volumeObject) { throw [System.ArgumentException]::new('volumes が必要です。') }
        $workbookIds = @{}
        foreach ($wb in @(Get-Array $structure.workbooks | Where-Object { Test-WorkbookPack $_ $cat })) { $workbookIds[[string]$wb.workbookId] = $true }
        $pageMap = @{}
        foreach ($page in @(Get-Array $structure.pages | Where-Object { $workbookIds.ContainsKey([string]$_.workbookId) })) { $pageMap[(Resolve-PageId $page)] = $page }
        $affected = New-Object System.Collections.Generic.HashSet[string]
        $layoutSnapshotId = ''
        foreach ($property in $volumeObject.PSObject.Properties) {
            $volume = [string]$property.Name
            if ($allowed -notcontains $volume) { throw [System.ArgumentException]::new("不正なvolumeです: $volume") }
            $desiredIds = @(Get-Array $property.Value | ForEach-Object { [string]$_ })
            $currentIds = @(Get-Array $structure.pages | Where-Object {
                $workbookIds.ContainsKey([string]$_.workbookId) -and (
                    ($volume -eq 'none' -and ([string]$_.volume -eq 'none' -or $_.enabled -eq $false)) -or
                    ($volume -ne 'none' -and [string]$_.volume -eq $volume -and $_.enabled -ne $false)
                )
            } | Sort-Object {[double](Get-DataProperty $_ 'order' 0)}, {Resolve-PageId $_} | ForEach-Object { Resolve-PageId $_ })
            $sameSequence = (($currentIds -join "`n") -eq ($desiredIds -join "`n"))
            if ($sameSequence) { continue }
            if ([string]::IsNullOrWhiteSpace($layoutSnapshotId)) { $layoutSnapshotId = Save-LayoutSnapshot $Language $cat 'reorder' $structure }
            if ($volume -ne 'none') { [void]$affected.Add($volume) }
            for ($i=0; $i -lt $desiredIds.Count; $i++) {
                $id = $desiredIds[$i]
                if (-not $pageMap.ContainsKey($id)) { continue }
                $page = $pageMap[$id]
                $oldVolume = [string](Get-DataProperty $page 'volume' 'none')
                if ($oldVolume -ne 'none') { [void]$affected.Add($oldVolume) }
                Set-NoteProperty $page 'volume' $volume
                Set-NoteProperty $page 'enabled' ($volume -ne 'none')
                Set-NoteProperty $page 'order' (($i+1)*10)
                Set-NoteProperty $page 'orderManual' $true
                Set-NoteProperty $page 'updatedAt' (New-NowIso)
            }
        }
        foreach ($volume in $allowed) { [void](Renumber-VolumeOrder $structure $volume $cat) }
        Apply-DefaultNumberingPerVolume $Language $structure $cat
        if ($affected.Count -gt 0) { Mark-VolumeNeedsRebuild $structure $Language $cat @($affected) 'reorder' 'ページ構成を変更しました' }
        return [ordered]@{ pages=$structure.pages; volumes=$structure.volumes; affectedVolumes=@($affected); layoutSnapshotId=$layoutSnapshotId; updatedAt=(New-NowIso) }
    } $baseLayout $requestedScope
}


function Update-Page([string]$Language, $Body) {
    $requestedScope = [string](Get-DataProperty $Body 'packId' (Get-DataProperty $Body 'category' ''))
    $baseLayout = Get-RequestedBaseLayout $Body
    return Update-StructureLocked $Language {
        param($structure)
        $cat = [string](Resolve-DocumentPackScope $structure $requestedScope $false).packId
        $page = @(Get-Array $structure.pages | Where-Object { (Resolve-PageId $_) -eq [string]$Body.pageId } | Select-Object -First 1)
        if ($page.Count -eq 0) { throw "Pageが見つかりません: $($Body.pageId)" }
        $p = $page[0]
        $workbook = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$p.workbookId } | Select-Object -First 1)
        if ($workbook.Count -eq 0 -or -not (Test-WorkbookPack $workbook[0] $cat)) { throw [ArgumentException]::new('指定された一式のページではありません。') }
        $beforeVol = [string]$p.volume
        $structural = $false
        $layoutRequested = @('title','volume','numberingMode','numberingManual','resetNumbering','pageRange','clearPageRange','enabled') | Where-Object { Test-ConfigHasKey $Body $_ }
        $layoutSnapshotId = if ($layoutRequested.Count -gt 0) { Save-LayoutSnapshot $Language $cat 'page-update' $structure } else { '' }

        if ($null -ne $Body.title) { Set-NoteProperty $p 'title' ([string]$Body.title); $structural = $true }
        if ($null -ne $Body.volume) {
            $targetVolume = [string]$Body.volume
            if (@(Get-PackVolumeList $Language (Resolve-DocumentPackScope $structure $requestedScope $false).pack $true) -notcontains $targetVolume) { throw [ArgumentException]::new('不正なvolumeです。') }
            Set-NoteProperty $p 'volume' $targetVolume
            Set-NoteProperty $p 'enabled' ($targetVolume -ne 'none')
            $structural = $true
        }
        if ($null -ne $Body.numberingMode) {
            if (@('none','visible') -notcontains [string]$Body.numberingMode) { throw [ArgumentException]::new('不正なnumberingModeです。') }
            Set-NoteProperty $p 'numberingMode' ([string]$Body.numberingMode)
            Set-NoteProperty $p 'numberingManual' $true
            $structural = $true
        }
        if ($null -ne $Body.numberingManual) { Set-NoteProperty $p 'numberingManual' ([bool]$Body.numberingManual); $structural = $true }
        if ($null -ne $Body.resetNumbering -and [bool]$Body.resetNumbering) { Set-NoteProperty $p 'numberingManual' $false; $structural = $true }
        if ($null -ne (Get-DataProperty $Body 'pageRange' $null)) {
            Set-NoteProperty $p 'pageRange' (ConvertTo-NormalizedPageRange (Get-DataProperty $Body 'pageRange' $null))
            $structural = $true
        }
        if ($null -ne (Get-DataProperty $Body 'clearPageRange' $null) -and [bool](Get-DataProperty $Body 'clearPageRange' $false)) {
            Set-NoteProperty $p 'pageRange' $null
            $structural = $true
        }
        if ($null -ne $Body.enabled) {
            if ([bool]$Body.enabled) {
                if ([string](Get-DataProperty $p 'volume' 'none') -eq 'none') {
                    throw [ArgumentException]::new('出力するページは、本体または補足へ割り当ててください。')
                }
                Set-NoteProperty $p 'enabled' $true
            } else {
                Set-NoteProperty $p 'enabled' $false
                Set-NoteProperty $p 'volume' 'none'
            }
            $structural = $true
        }

        Set-NoteProperty $p 'updatedAt' (New-NowIso)
        foreach ($volume in @(Get-PackVolumeList $Language (Resolve-DocumentPackScope $structure $cat $true).pack $true)) { [void](Renumber-VolumeOrder $structure $volume $cat) }
        Apply-DefaultNumberingPerVolume $Language $structure $cat
        if ($structural) {
            $affected = @($beforeVol, [string]$p.volume) | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique
            Mark-VolumeNeedsRebuild $structure $Language $cat @($affected) 'reorder' 'ページ構成を変更しました'
        }
        return [ordered]@{ page=$p; volumes=$structure.volumes; layoutSnapshotId=$layoutSnapshotId }
    } $baseLayout $requestedScope
}
function Confirm-Page([string]$Language, $Body) {
    $cat=Require-WorkbookCategory ([string]$Body.category)
    return Update-StructureLocked $Language { param($structure) $page=@(Get-Array $structure.pages|Where-Object{(Resolve-PageId $_)-eq [string]$Body.pageId}|Select-Object -First 1);if($page.Count -eq 0){throw "Pageが見つかりません: $($Body.pageId)"};if((Get-PageCategory $structure $page[0]) -ne $cat){throw [ArgumentException]::new('指定カテゴリのページではありません。')};switch([string]$Body.action){'confirm'{Set-NoteProperty $page[0] 'status' 'confirmed'}'reject'{Set-NoteProperty $page[0] 'status' 'rejected'}default{throw [ArgumentException]::new('action は confirm または reject を指定してください。')}};Set-NoteProperty $page[0] 'updatedAt' (New-NowIso);return $page[0] }
}


function Sort-PagesByNumericSheet([string]$Language, $Body) {
    $requestedScope = [string](Get-DataProperty $Body 'packId' (Get-DataProperty $Body 'category' ''))
    $baseLayout = Get-RequestedBaseLayout $Body
    return Update-StructureLocked $Language {
        param($structure)
        $cat = [string](Resolve-DocumentPackScope $structure $requestedScope $false).packId
        $allowedVolumes = @(Get-PackVolumeList $Language (Resolve-DocumentPackScope $structure $cat $true).pack $true)
        $volumes = if ($Body.volumes) {
            @(Get-Array $Body.volumes | ForEach-Object { [string]$_ })
        } else {
            @($allowedVolumes)
        }
        $wbMap = @{}
        foreach ($wb in @(Get-Array $structure.workbooks | Where-Object { Test-WorkbookPack $_ $cat })) {
            $wbMap[[string]$wb.workbookId] = $wb
        }
        $affected = New-Object System.Collections.Generic.HashSet[string]
        $layoutSnapshotId = ''
        foreach ($volume in @($volumes | Select-Object -Unique)) {
            if ($allowedVolumes -notcontains $volume) { throw [ArgumentException]::new("不正なvolumeです: $volume") }
            $current = @(Get-Array $structure.pages | Where-Object {
                $wbMap.ContainsKey([string]$_.workbookId) -and (
                    ($volume -eq 'none' -and ([string]$_.volume -eq 'none' -or $_.enabled -eq $false)) -or
                    ($volume -ne 'none' -and [string]$_.volume -eq $volume -and $_.enabled -ne $false)
                )
            } | Sort-Object {[double](Get-DataProperty $_ 'order' 0)}, {Resolve-PageId $_})
            $sorted = @(Sort-NumericPagesWithinAnchors $current $wbMap)
            $beforeIds = @($current | ForEach-Object { Resolve-PageId $_ })
            $afterIds = @($sorted | ForEach-Object { Resolve-PageId $_ })
            # A sequence-equal lane is a true no-op, even if an older reorder
            # request left orderManual markers behind. Resetting flags alone
            # would create a misleading history entry and an undo with no
            # visible layout to restore.
            $inputChanged = (($beforeIds -join [Environment]::NewLine) -ne ($afterIds -join [Environment]::NewLine))
            if (-not $inputChanged) { continue }
            if ([string]::IsNullOrWhiteSpace($layoutSnapshotId)) { $layoutSnapshotId = Save-LayoutSnapshot $Language $cat 'sort-by-numeric-sheet' $structure }
            for ($i=0; $i -lt $sorted.Count; $i++) {
                $expected = ($i + 1) * 10
                Set-NoteProperty $sorted[$i] 'order' $expected
                Set-NoteProperty $sorted[$i] 'orderManual' $false
                Set-NoteProperty $sorted[$i] 'updatedAt' (New-NowIso)
            }
            if ($inputChanged -and $volume -ne 'none') { [void]$affected.Add($volume) }
        }
        Apply-DefaultNumberingPerVolume $Language $structure $cat
        if ($affected.Count -gt 0) {
            Mark-VolumeNeedsRebuild $structure $Language $cat @($affected) 'reorder' 'ページを半角数字シート順に並べ替えました'
        }
        return [ordered]@{ pages=$structure.pages; volumes=$structure.volumes; packId=$cat; category=(Get-CategoryFromBuiltinPackId $cat); affectedVolumes=@($affected); layoutSnapshotId=$layoutSnapshotId }
    } $baseLayout $requestedScope
}

function Resolve-JavaExe {
    # 一度見つかったパスはキャッシュする(共有フォルダ上のTest-Pathの瞬断対策も兼ねる)。
    if (-not [string]::IsNullOrWhiteSpace($Script:CachedJavaExe)) { return $Script:CachedJavaExe }
    # Prefer a local portable runtime installed by app\tools\install-thirdparty.cmd.
    $direct = Join-Path $Script:AppRoot 'lib\java\bin\java.exe'
    # ネットワーク共有では Test-Path がウイルススキャン等で一時的に false になることがあるためリトライする。
    for ($javaAttempt = 1; $javaAttempt -le 3; $javaAttempt++) {
        if (Test-Path -LiteralPath $direct) { $Script:CachedJavaExe = $direct; return $direct }
        if ($javaAttempt -lt 3) { Start-Sleep -Seconds 2 }
    }
    $javaRoot = Join-Path $Script:AppRoot 'lib\java'
    if (Test-Path -LiteralPath $javaRoot) {
        $found = @(Get-ChildItem -LiteralPath $javaRoot -Filter java.exe -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match '[\\/]bin[\\/]java\.exe$' } |
            Sort-Object FullName |
            Select-Object -First 1)
        if ($found.Count -gt 0) { return [string]$found[0].FullName }
    }
    $cmd = Get-Command java.exe -ErrorAction SilentlyContinue
    if ($cmd) { return [string]$cmd.Source }
    $cmd2 = Get-Command java -ErrorAction SilentlyContinue
    if ($cmd2) { return [string]$cmd2.Source }
    throw ("提出用PDFの作成に必要なJava Runtimeが見つかりません(探した場所: {0})。この場所にjava.exeがあるのにこのエラーが出る場合は、server.ps1の置き場所(AppRoot)がずれています。app\logs\startup-*-latest.log の AppRoot 行を確認してください。java.exe自体がない場合は app\tools\install-thirdparty.cmd を実行してから、もう一度PDFを出力してください。" -f $direct)
}

function Get-JavaRuntimeSignature {
    # V5-P1(#11): 画像ハッシュの環境指紋に含める Java 版の署名。
    # `java -version` の出力(バージョン+ビルド)を採取してハッシュ化し、Java 更新で必ず値が変わるようにする。
    if (-not [string]::IsNullOrWhiteSpace($Script:JavaRuntimeSignature)) { return $Script:JavaRuntimeSignature }
    $sig = 'unknown'
    try {
        $java = Resolve-JavaExe
        # `java -version` はバージョン情報を stderr に出すため 2>&1 で取り込む。
        $probe = Invoke-NativeCapture $java @('-version') $Script:JavaProbeTimeoutSeconds
        $out = [string]$probe.text
        $norm = ($out -replace '\s+', ' ').Trim()
        # 起動に失敗すると例外メッセージが text として返る。それを版の署名にすると、
        # 実行できないJavaの「版が変わった」を延々と記録し続けることになる。
        if ([int]$probe.exitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($norm)) { $sig = (Get-Sha256Text $norm).Substring(0, 16) }
    } catch { $sig = 'unknown' }
    $Script:JavaRuntimeSignature = $sig
    return $sig
}

function Get-Sha256Text([string]$Text) {
    $sha=[Security.Cryptography.SHA256]::Create();try{$bytes=[Text.Encoding]::UTF8.GetBytes($Text);return 'sha256:'+([BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-','').ToLowerInvariant())}finally{$sha.Dispose()}
}

function Get-FinalBuildInputSnapshot($Structure,[string]$Language,[string]$Volume,[string]$Category) {
    $scope=Resolve-DocumentPackScope $Structure $Category $false;$packId=[string]$scope.packId;$cat=[string]$scope.category
    if(@(Get-PackVolumeList $Language $scope.pack $false) -notcontains $Volume){throw [ArgumentException]::new("不正な成果物です: $Volume")}
    $workspace=Get-WorkspacePath $Language;$wbMap=@{};$targets=@(Get-Array $Structure.workbooks|Where-Object{Test-WorkbookPack $_ $packId});foreach($wb in $targets){$wbMap[[string]$wb.workbookId]=$wb}
    $sourceMap=@{};foreach($source in @(Get-Array (Get-DataProperty $Structure 'sources' @()))){$sourceMap[[string](Get-DataProperty $source 'sourceId' '')]=$source}
    $itemMap=@{};foreach($item in @(Get-Array (Get-DataProperty $Structure 'items' @()))){$itemMap[[string](Get-DataProperty $item 'itemId' '')]=$item}
    $unitMap=@{};foreach($unit in @(Get-Array (Get-DataProperty $Structure 'units' @()))){$unitMap[[string](Get-DataProperty $unit 'unitId' '')]=$unit}
    $artifactMap=@{};foreach($artifact in @(Get-Array (Get-DataProperty $Structure 'artifacts' @()))){$artifactMap[[string](Get-DataProperty $artifact 'artifactId' '')]=$artifact}
    $packObject=$scope.pack
    $effectiveTemplate=Get-PackEffectiveTemplate $Language $packObject
    $settings=Get-ResolvedPackSettings $packObject $Language;$targetName=Get-PackTargetDisplayName $Language $packObject (Get-TargetIdFromLegacyVolume $Volume);$projectId=Get-ProjectIdFromWorkbooks $targets;$namedProjectId=if([bool]$scope.builtIn){Get-CategoryProjectId $projectId $cat}else{$projectId}
    $fallbackPackName=if([bool]$scope.builtIn){$cat.ToUpperInvariant()}else{'ReportBinder'}
    $outputFileName=Resolve-OutputFileNamePattern ([string]$settings.outputFileNamePattern) $namedProjectId ([string](Get-DataProperty $packObject 'displayName' $fallbackPackName)) $targetName
    $document=[ordered]@{title=[string]$settings.documentTitle;subtitle=[string]$settings.documentSubtitle;targetName=$targetName;outputFileNamePattern=[string]$settings.outputFileNamePattern;outputFileName=$outputFileName}
    $pages=@(Get-Array $Structure.pages|Where-Object{$_.enabled -eq $true -and [string]$_.volume -eq $Volume -and $wbMap.ContainsKey([string]$_.workbookId)}|Sort-Object {[double]$_.order},{Resolve-PageId $_})
    $blockers=@();$manifest=@();$fpPages=@()
    $requiredRequirementIds=@{}
    foreach($requirement in @(Get-Array (Get-DataProperty $effectiveTemplate 'sourceRequirements' @()))){
        if(-not [bool](Get-DataProperty $requirement 'required' $true)){continue}
        $requirementId=[string](Get-DataProperty $requirement 'requirementId' '')
        if([string]::IsNullOrWhiteSpace($requirementId)){continue}
        $requiredRequirementIds[$requirementId]=$true
        $assigned=@($targets|Where-Object{[string](Get-DataProperty $_ 'requirementId' '') -eq $requirementId})
        if($assigned.Count -eq 0){
            $blockers+=[ordered]@{code='required-source-unregistered';requirementId=$requirementId;sourceId='';pageTitle='';workbookName=[string](Get-DataProperty $requirement 'displayName' '');ownerDepartment=[string](Get-DataProperty $requirement 'ownerDepartment' '');dueDate=[string](Get-DataProperty $requirement 'dueDate' '');message='必須原稿がまだ登録されていません。必要原稿リストから原稿を割り当ててください。'}
        }
    }
    # A source marked "required" is a safety contract, independent of template
    # defaults. Never publish a pack while that source is missing, failed, or
    # stale; otherwise a partial PDF could be labelled as current.
    $submissionRoot=[string](Get-DataProperty (Get-Paths) 'submissionDir' '')
    foreach($requiredSource in @($targets|Where-Object{[bool](Get-DataProperty $_ 'required' $false) -or $requiredRequirementIds.ContainsKey([string](Get-DataProperty $_ 'requirementId' ''))})){
            $requiredName=[string](Get-DataProperty $requiredSource 'displayName' (Get-DataProperty $requiredSource 'fileName' ''))
            $relativePath=[string](Get-DataProperty $requiredSource 'relativePath' '')
            $sourceExists=$false
            if(-not [string]::IsNullOrWhiteSpace($relativePath)){
                try{$sourceExists=Test-Path -LiteralPath (Join-Safe $submissionRoot $relativePath) -PathType Leaf}catch{$sourceExists=$false}
            }
            if(-not $sourceExists){
                $blockers+=[ordered]@{code='required-source-missing';sourceId=[string](Get-DataProperty $requiredSource 'workbookId' '');pageTitle='';workbookName=$requiredName;message='必須原稿のファイルが見つかりません。提出フォルダを確認してください。'}
                continue
            }
            $lastError=[string](Get-DataProperty $requiredSource 'lastError' '')
            $sourceStatus=([string](Get-DataProperty $requiredSource 'status' '')).ToLowerInvariant()
            if(-not [string]::IsNullOrWhiteSpace($lastError) -or $sourceStatus -match 'fail|error'){
                $blockers+=[ordered]@{code='required-source-failed';sourceId=[string](Get-DataProperty $requiredSource 'workbookId' '');pageTitle='';workbookName=$requiredName;message='必須原稿のPDF変換に失敗しています。エラーを解消して再作成してください。'}
                continue
            }
            if(-not (Test-WorkbookRenderIsCurrent $requiredSource)){
                $hasRendered=-not [string]::IsNullOrWhiteSpace([string](Get-DataProperty $requiredSource 'lastRenderedVersionId' ''))
                $blockers+=[ordered]@{code=$(if($hasRendered){'required-source-stale'}else{'required-source-not-rendered'});sourceId=[string](Get-DataProperty $requiredSource 'workbookId' '');pageTitle='';workbookName=$requiredName;message=$(if($hasRendered){'必須原稿が更新されています。先に変換PDFを再作成してください。'}else{'必須原稿の変換PDFが未作成です。先にPDFを作成してください。'})}
            }
    }
    if($pages.Count -eq 0){$blockers+= [ordered]@{code='no-pages';pageTitle='';workbookName='';message='対象ページがありません。ページ構成を確認してください。'}}
    foreach($p in $pages){$pageId=Resolve-PageId $p;$wb=$wbMap[[string]$p.workbookId];$title=[string]$p.title;if([string]::IsNullOrWhiteSpace($title)){$title=[string]$p.sheetName};$wbName=[string]$wb.fileName;$rel=[string]$p.contentPdf;$full='';$size=0L;$ticks=0L;$sourcePageCount=1
        if(-not $rel){$blockers+=[ordered]@{code='content-missing';pageTitle=$title;workbookName=$wbName;message='PDF未作成のページがあります。先にPDF作成してください。'}}
        else{try{$full=[IO.Path]::GetFullPath((Join-Path $workspace $rel));$root=[IO.Path]::GetFullPath($workspace);if(-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)){$root+=[IO.Path]::DirectorySeparatorChar};if(-not $full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'outside'};if(-not(Test-Path $full)){throw 'missing'};$it=Get-Item $full;$size=$it.Length;$ticks=$it.LastWriteTimeUtc.Ticks;$sourcePageCount=Get-PdfPageCount $full}catch{$blockers+=[ordered]@{code='content-file-missing';pageTitle=$title;workbookName=$wbName;message='レンダリング済みPDFが見つかりません。先にPDF作成してください。'}}}
        if(-not(Test-WorkbookRenderIsCurrent $wb)){$blockers+=[ordered]@{code='stale-content';pageTitle=$title;workbookName=$wbName;message='元原稿が更新されています。先に変換PDFを作成してください。'}}
        elseif(-not(Test-WorkbookRenderedSheetContains $wb ([string]$p.sheetName))){$blockers+=[ordered]@{code='sheet-changed';pageTitle=$title;workbookName=$wbName;message='原稿のページ構成が変わっています。先に変換PDFを作成してください。'}}
        elseif(-not(Test-PageContentMatchesWorkbookVersion $p $wb)){$blockers+=[ordered]@{code='old-content';pageTitle=$title;workbookName=$wbName;message='古いPDF参照が残っています。先にPDF作成してください。'}}
        $item=if($itemMap.ContainsKey($pageId)){$itemMap[$pageId]}else{$null};$unitId=[string](Get-DataProperty $item 'unitId' (New-WorksheetUnitId ([string]$p.workbookId) ([string]$p.sheetName)));$unit=if($unitMap.ContainsKey($unitId)){$unitMap[$unitId]}else{$null};$artifactId=[string](Get-DataProperty $item 'artifactId' '');$artifact=if($artifactMap.ContainsKey($artifactId)){$artifactMap[$artifactId]}else{$null};$source=if($sourceMap.ContainsKey([string]$p.workbookId)){$sourceMap[[string]$p.workbookId]}else{$null};$recordedPageCount=Get-IntDataProperty $artifact 'pageCount' (Get-IntDataProperty $unit 'physicalPageCount' 0);if($recordedPageCount -gt 0){$sourcePageCount=$recordedPageCount}
        $rangeValue=Get-DataProperty $p 'pageRange' (Get-DataProperty $item 'pageRange' $null);$range=$null;try{$range=ConvertTo-NormalizedPageRange $rangeValue}catch{$blockers+=[ordered]@{code='invalid-page-range';pageTitle=$title;workbookName=$wbName;message=$_.Exception.Message}}
        $rangeStart=if($null -ne $range){[int]$range.start}else{1};$rangeEnd=if($null -ne $range){[int]$range.end}else{$sourcePageCount}
        if($rangeStart -gt $sourcePageCount -or $rangeEnd -gt $sourcePageCount){$blockers+=[ordered]@{code='page-range-out-of-bounds';pageTitle=$title;workbookName=$wbName;message="ページ範囲 $rangeStart-$rangeEnd は変換PDF（$sourcePageCount ページ）の範囲外です。"}}
        $sourceType=[string](Get-DataProperty $source 'sourceType' (Get-DataProperty $wb 'sourceType' 'excel'));$sourceKey=[string](Get-DataProperty $unit 'sourceKey' ([string]$p.sheetName));$sectionId="source:$([string]$p.workbookId)";$sectionTitle=[string](Get-DataProperty $source 'displayName' (Get-DataProperty $wb 'displayName' $wbName));$bookmarkTitle=if([string]::IsNullOrWhiteSpace($sectionTitle)){$title}else{"$sectionTitle / $title"}
        $normalized=($rel -replace '\\','/').ToLowerInvariant();$fpPages+=[ordered]@{pageId=$pageId;itemId=[string](Get-DataProperty $item 'itemId' $pageId);order=[double]$p.order;enabled=[bool]$p.enabled;volume=[string]$p.volume;title=$title;pageRange=$range;numberingMode=[string]$p.numberingMode;contentPdf=$normalized;contentPdfSize=$size;contentPdfLastWriteUtcTicks=$ticks;lastRenderedVersionId=[string]$wb.lastRenderedVersionId}
        if($full){$manifest+=[ordered]@{pageId=$pageId;itemId=[string](Get-DataProperty $item 'itemId' $pageId);sourceId=[string]$p.workbookId;unitId=$unitId;artifactId=[string](Get-DataProperty $artifact 'artifactId' $artifactId);sourceType=$sourceType;sourceKey=$sourceKey;title=$title;bookmarkTitle=$bookmarkTitle;sectionId=$sectionId;sectionTitle=$sectionTitle;sourcePdf=$full;sourcePageStart=$rangeStart;sourcePageEnd=$rangeEnd;sourcePageCount=$sourcePageCount;numberingMode=[string]$p.numberingMode;punchShiftPt=(Convert-CmToPt 0.2)}}
    }
    $physicalPages=@();$outputPageNumber=0
    foreach($entry in $manifest){
        for($sourcePage=[int]$entry.sourcePageStart;$sourcePage -le [int]$entry.sourcePageEnd;$sourcePage++){$outputPageNumber++;$physicalPages+=[ordered]@{outputPageNumber=$outputPageNumber;kind='content';itemId=[string]$entry.itemId;pageId=[string]$entry.pageId;sourceId=[string]$entry.sourceId;unitId=[string]$entry.unitId;artifactId=[string]$entry.artifactId;sourceType=[string]$entry.sourceType;sourceKey=[string]$entry.sourceKey;sourcePdfPageNumber=$sourcePage;title=[string]$entry.title;sectionId=[string]$entry.sectionId;numberingMode=[string]$entry.numberingMode}}
    }
    $input=[ordered]@{composerProfileVersion=$Script:FinalPdfComposerProfileVersion;language=$Language;packId=$packId;category=$cat;volume=$Volume;document=$document;pages=$fpPages};$json=ConvertTo-Json $input -Depth 20 -Compress;$fingerprint=Get-Sha256Text $json
    return [ordered]@{language=$Language;packId=$packId;category=$cat;volume=$Volume;pages=$pages;manifestPages=$manifest;physicalPages=$physicalPages;document=$document;blockers=@($blockers);itemCount=$pages.Count;pageCount=$outputPageNumber;projectId=$projectId;outputFileName=$outputFileName;fingerprint=$fingerprint;fingerprintInput=$input}
}

function Get-OfficeApplicationDiagnostic([string]$ProgId, [string]$DisplayName) {
    $curVer = ''; $registered = $false
    try {
        $key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("$ProgId\CurVer")
        if ($null -ne $key) { try { $curVer = [string]$key.GetValue(''); $registered = -not [string]::IsNullOrWhiteSpace($curVer) } finally { $key.Dispose() } }
    } catch { }
    $app = $null; $version = ''; $build = ''; $errorCode = ''; $errorMessage = ''
    if ($registered) {
        try {
            $app = New-Object -ComObject $ProgId
            try { $app.Visible = $false } catch { }
            try { $app.DisplayAlerts = $false } catch { }
            try { $version = [string]$app.Version } catch { }
            try { $build = [string]$app.Build } catch { }
        } catch {
            $errorMessage = [string]$_.Exception.Message
            if ($errorMessage -match '80070520|0x80070520') { $errorCode = 'windows-session-unavailable' }
            elseif ($errorMessage -match '80040154|0x80040154') { $errorCode = 'class-not-registered' }
            else { $errorCode = 'com-start-failed' }
        } finally {
            if ($null -ne $app) {
                try { $app.Quit() } catch { }
                try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch { }
            }
        }
    } else { $errorCode = 'not-installed' }
    $major = 0; if ($version -match '^(?<major>\d+)') { $major = [int]$matches.major }
    $compatible = ($registered -and -not [string]::IsNullOrWhiteSpace($version) -and $major -ge 16)
    $available = ($compatible -and [string]::IsNullOrWhiteSpace($errorCode))
    $status = if ($available) { 'ready' } elseif (-not $registered) { 'not-installed' } elseif ($version -and -not $compatible) { 'unsupported' } else { 'unavailable' }
    $action = switch ($status) {
        'ready' { '' }
        'not-installed' { "$DisplayName デスクトップ版をインストールしてください。PDF原稿だけの処理は継続できます。" }
        'unsupported' { "$DisplayName 16.x以降へ更新してください。" }
        default { "$DisplayName は登録済みですが起動できません。Windowsへサインインし直し、デスクトップからReportBinderを起動してください。" }
    }
    return [ordered]@{ appId=$ProgId; displayName=$DisplayName; registered=[bool]$registered; registeredVersion=$curVer; available=[bool]$available; compatible=[bool]$compatible; version=$version; build=$build; status=$status; errorCode=$errorCode; errorMessage=$errorMessage; userAction=$action }
}

function Get-PathFreeBytes([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    try {
        $full = [IO.Path]::GetFullPath($Path); $root = [IO.Path]::GetPathRoot($full)
        if ([string]::IsNullOrWhiteSpace($root)) { return $null }
        return [int64]([IO.DriveInfo]::new($root).AvailableFreeSpace)
    } catch { return $null }
}

function Get-SystemDiagnostics([string]$Language) {
    $excel = Get-OfficeApplicationDiagnostic 'Excel.Application' 'Microsoft Excel'
    $word = Get-OfficeApplicationDiagnostic 'Word.Application' 'Microsoft Word'
    $powerPoint = Get-OfficeApplicationDiagnostic 'PowerPoint.Application' 'Microsoft PowerPoint'
    $composer = Join-Path $Script:AppRoot 'lib\pdfbox\ReportPdfComposer.jar'
    $pdfbox = Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'
    $pdfjsMain = Join-Path $Script:WebRoot 'pdfjs\pdf.min.mjs'
    $pdfjsWorker = Join-Path $Script:WebRoot 'pdfjs\pdf.worker.min.mjs'
    $javaPath = ''; $javaVersion = ''; $javaReady = $false
    # exitCode を見ずにテキストの有無だけで判定すると、実行を禁止されている java や
    # 展開が不完全な同梱JREでも「使える」と診断してしまう。診断画面が最も頼られるのは
    # まさにその場面なので、終了コードまで確認する。
    try { $javaPath = Resolve-JavaExe; $javaProbe = Invoke-NativeCapture $javaPath @('-version') $Script:JavaProbeTimeoutSeconds; $javaVersion = ([string]$javaProbe.text -replace '\s+', ' ').Trim(); $javaReady = ([int]$javaProbe.exitCode -eq 0 -and -not [string]::IsNullOrWhiteSpace($javaVersion)) } catch { $javaVersion = $_.Exception.Message }
    $coreReady = ($javaReady -and (Test-Path -LiteralPath $composer) -and (Test-Path -LiteralPath $pdfbox) -and (Test-Path -LiteralPath $pdfjsMain) -and (Test-Path -LiteralPath $pdfjsWorker))
    $paths = Get-Paths
    $officeReady = ([bool]$excel.available -and [bool]$word.available -and [bool]$powerPoint.available)
    $overall = if (-not $coreReady) { 'blocked' } elseif ($officeReady) { 'ready' } else { 'limited' }
    $summary = if ($overall -eq 'ready') { 'Excel・Word・PowerPoint・PDF原稿を処理できます。' } elseif ($overall -eq 'limited') { 'PDF原稿は処理できます。Office原稿には対応が必要です。' } else { '実行依存ファイルが不足しているため、管理者による修復が必要です。' }
    return [ordered]@{
        capturedAt = New-NowIso; language=$Language; status=$overall; summary=$summary
        office = [ordered]@{ minimumSupportedMajor=16; excel=$excel; word=$word; powerPoint=$powerPoint }
        runtime = [ordered]@{
            java=[ordered]@{ready=[bool]$javaReady;path=$javaPath;version=$javaVersion}
            pdfbox=[ordered]@{ready=[bool]((Test-Path -LiteralPath $composer) -and (Test-Path -LiteralPath $pdfbox));composerBytes=$(if(Test-Path -LiteralPath $composer){(Get-Item -LiteralPath $composer).Length}else{0});libraryBytes=$(if(Test-Path -LiteralPath $pdfbox){(Get-Item -LiteralPath $pdfbox).Length}else{0})}
            pdfjs=[ordered]@{ready=[bool]((Test-Path -LiteralPath $pdfjsMain) -and (Test-Path -LiteralPath $pdfjsWorker))}
        }
        storage = [ordered]@{
            configured=[bool](-not [string]::IsNullOrWhiteSpace([string]$paths.submissionDir)); submissionDir=[string]$paths.submissionDir; dataDir=[string]$paths.dataDir; outputDir=[string]$paths.outputDir
            dataFreeBytes=Get-PathFreeBytes ([string]$paths.dataDir); outputFreeBytes=Get-PathFreeBytes ([string]$paths.outputDir)
        }
    }
}

function Get-FinalBuildFingerprint($Snapshot) { return [string](Get-DataProperty $Snapshot 'fingerprint' '') }

function Get-FinalBuildReadiness($Structure,[string]$Language,[string]$Volume,[string]$Category,[bool]$CheckOutputExists = $true) {
    $scope=Resolve-DocumentPackScope $Structure $Category $false;$packId=[string]$scope.packId;$snap=Get-FinalBuildInputSnapshot $Structure $Language $Volume $packId;$v=Get-PackOutputState $Structure $Language $packId $Volume $false;if($null -eq $v){$v=New-EmptyVolumeState};$built=[string](Get-DataProperty $v 'builtFingerprint' '');$current=[string]$snap.fingerprint;$out=[string](Get-DataProperty $v 'outputPdf' '');$exists=(-not [string]::IsNullOrWhiteSpace($out));if($exists -and $CheckOutputExists){$exists=Test-Path -LiteralPath $out};$status='not-built';if($built){if($built -ne $current -or $snap.blockers.Count -gt 0){$status='needs-rebuild'}else{$status='built'}};$display=if($snap.blockers.Count -gt 0){'blocked'}elseif(-not $built){'not-built'}elseif($built -ne $current){'needs-rebuild'}elseif(-not $exists){'output-missing'}else{'built'}
    Set-NoteProperty $v 'status' $status
    $reasons=@(Get-Array (Get-DataProperty $v 'staleReasons' @()));if($display -eq 'needs-rebuild' -and $reasons.Count -eq 0){$reasons=@([ordered]@{type='fingerprint';at=(Get-DataProperty $Structure 'updatedAt' $null);detail='提出用PDFの入力が変更されました'})}
    return [ordered]@{packId=$packId;targetId=(Get-TargetIdFromLegacyVolume $Volume);canBuild=($snap.blockers.Count -eq 0 -and $snap.pageCount -gt 0);pageCount=$snap.pageCount;status=$status;displayState=$display;builtFingerprint=$built;currentFingerprint=$current;outputPdf=$out;outputPdfExists=[bool]$exists;lastBuiltAt=(Get-DataProperty $v 'lastBuiltAt' $null);blockers=@($snap.blockers);staleReasons=@($reasons);snapshot=$snap}
}

function Render-PowerPointSnapshotForComparison([string]$Language, [string]$SourceId, [string]$SnapshotId) {
    $state = Get-SnapshotSourceState $Language $SourceId $SnapshotId
    if (-not [bool]$state.sourceRetained) { return [ordered]@{ ok=$false; reason='source-missing' } }
    return Invoke-WithRenderLock $Language $SourceId {
        $versionId=New-RbVersionId; $workspace=Get-WorkspacePath $Language
        $contentDir=Get-ContentPdfVersionDir $workspace $SourceId $versionId
        $tmpDir=Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-powerpoint-compare-' + (New-RbId))
        $jobId=New-RbId; $snapshotLease=''; $contentLease=''; $success=$false
        try {
            New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
            New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
            $snapshotLease=New-SnapshotLease $Language $SourceId $SnapshotId 'compare' $jobId 180 $versionId
            $contentLease=New-ContentPdfLease $workspace $SourceId $versionId 'compare' $jobId 180
            if ([string]::IsNullOrWhiteSpace($snapshotLease) -or [string]::IsNullOrWhiteSpace($contentLease)) { throw '比較資産の保護leaseを作成できませんでした。' }
            $workPptx=Join-Path $tmpDir 'source.pptx'; Copy-FileSharedRead ([string]$state.sourcePath) $workPptx
            [void](Inspect-PowerPointSourceFile $workPptx)
            $wholePdf=Join-Path $tmpDir 'presentation.pdf'
            $powerPointResult=Invoke-PowerPointToPdf $workPptx $wholePdf $Script:PowerPointRenderTimeoutSeconds
            $split=Convert-PowerPointSplitPages (Invoke-PdfSourceSplit $wholePdf $contentDir) $contentDir
            $sheets=@()
            foreach ($pageInfo in @($split.pages)) {
                $name="Slide $([int]$pageInfo.pageNumber)"
                $sheets += @{ sheetName=$name; pdf=[string]$pageInfo.pdf; rasterDirectory=(Get-RenderRasterSheetDir $Language $SourceId $SnapshotId $versionId $name) }
            }
            $powerPointVersion=[string](Get-DataProperty $powerPointResult 'powerPointVersion' '')
            $Script:CurrentRenderEnvFingerprint=Get-Sha256Text "powerpoint-com|$powerPointVersion|$($Script:PowerPointRenderProfileVersion)"
            $analysis=Invoke-PdfPageAnalyzer $sheets
            if ($null -eq $analysis -or -not [bool]$analysis.ok) { return [ordered]@{ ok=$false; reason='analyze-failed' } }
            Write-RenderRecord $Language $SourceId $SnapshotId $versionId 'comparison' $true ([pscustomobject]$analysis.result)
            $availability=Get-HistoryRenderVersionAvailability $Language $SourceId $SnapshotId $versionId
            if (-not [bool]$availability.ready) { return [ordered]@{ ok=$false; reason='retention-verify-failed'; message=[string]$availability.reason } }
            $success=$true
            return [ordered]@{ ok=$true; versionId=$versionId; envFingerprint=[string]$Script:CurrentRenderEnvFingerprint }
        } catch { return [ordered]@{ ok=$false; reason='error'; message=$_.Exception.Message } }
        finally {
            Remove-SnapshotLease $Language $SourceId $SnapshotId $snapshotLease
            Remove-ContentPdfLease $workspace $SourceId $versionId $contentLease
            if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (-not $success -and (Test-Path -LiteralPath $contentDir)) { Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
}

function Convert-PowerPointSplitPages($Split, [string]$ContentDir) {
    $pages = @()
    foreach ($pageInfo in @(Get-Array (Get-DataProperty $Split 'pages' @()))) {
        $number = [int](Get-DataProperty $pageInfo 'pageNumber' 0)
        if ($number -le 0) { throw 'PowerPointのスライド番号を確認できませんでした。' }
        $name = "Slide $number"
        $sourcePath = [string](Get-DataProperty $pageInfo 'pdf' '')
        if ([string]::IsNullOrWhiteSpace($sourcePath) -or -not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) { throw "PowerPointの $number スライド目を取り込めませんでした。" }
        $destination = Join-Path $ContentDir ((Get-WorksheetStorageStem $name) + '.pdf')
        if (-not [IO.Path]::GetFullPath($sourcePath).Equals([IO.Path]::GetFullPath($destination), [StringComparison]::OrdinalIgnoreCase)) {
            Move-Item -LiteralPath $sourcePath -Destination $destination -Force
        }
        $pages += [pscustomobject][ordered]@{
            pageNumber = $number; sheetName = $name; sheetIndex = $number; detectedTitle = $name
            pdf = $destination; pageCount = 1; warnings = @()
        }
    }
    return [pscustomobject][ordered]@{ ok=$true; pageCount=$pages.Count; pages=@($pages) }
}

function Render-PowerPointSource([string]$Language, [string]$SourceId, [scriptblock]$ProgressCallback = $null,
                                 [string]$SourceOverridePath = '', [string]$SourceSnapshotId = '', [string]$ExpectedSourceHash = '') {
    return Invoke-WithRenderLock $Language $SourceId {
        $workspace = Get-WorkspacePath $Language
        $paths = Get-Paths
        $structure = Get-Structure $Language
        $found = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $SourceId } | Select-Object -First 1)
        if ($found.Count -eq 0) { throw "登録済みPowerPoint原稿が見つかりません: $SourceId" }
        $source = $found[0]
        if ([string](Get-DataProperty $source 'sourceType' '') -ne 'powerpoint') { throw 'PowerPoint原稿ではないためPowerPointアダプターで処理できません。' }
        $livePath = Join-Safe ([string]$paths.submissionDir) ([string]$source.relativePath)
        if (-not (Test-Path -LiteralPath $livePath)) { throw "PowerPoint原稿が見つかりません: $($source.relativePath)" }
        $inputPath = $SourceOverridePath
        $ephemeralCaptureId = ''
        if ([string]::IsNullOrWhiteSpace($inputPath) -and (Test-InputHistoryEnabled)) {
            if ([string]::IsNullOrWhiteSpace($SourceSnapshotId)) {
                $captured = Capture-DetectedSnapshot $Language $SourceId 'render'
                if ([bool]$captured.ok) { $SourceSnapshotId = [string]$captured.snapshotId }
            }
            $inputInfo = Capture-RenderInput $Language $SourceId $SourceSnapshotId ''
            if (-not [string]::IsNullOrWhiteSpace([string]$inputInfo.path)) {
                $inputPath = [string]$inputInfo.path; $SourceSnapshotId = [string]$inputInfo.snapshotId
                if ([bool]$inputInfo.ephemeral) { $ephemeralCaptureId = [string]$inputInfo.captureId }
            }
        }
        if ([string]::IsNullOrWhiteSpace($inputPath)) { $inputPath = $livePath }
        $pptInspection = Inspect-PowerPointSourceFile $inputPath
        $sourceHash = Normalize-FileHash (New-Sha256 $inputPath)
        $expectedHash = Normalize-FileHash $ExpectedSourceHash
        if (-not [string]::IsNullOrWhiteSpace($expectedHash) -and $sourceHash -ne $expectedHash) { throw '指定したPowerPoint原稿の版と実ファイルが一致しません。更新確認後に再度変換してください。' }
        $versionId = New-RbVersionId
        $contentDir = Get-ContentPdfVersionDir $workspace $SourceId $versionId
        $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('reportbinder-powerpoint-' + (New-RbId))
        $success = $false
        try {
            New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
            New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
            $workPptx = Join-Path $tmpDir 'source.pptx'
            Copy-FileSharedRead $inputPath $workPptx
            $wholePdf = Join-Path $tmpDir 'presentation.pdf'
            if ($ProgressCallback) { & $ProgressCallback 'powerpoint' $SourceId '' }
            $powerPointResult = Invoke-PowerPointToPdf $workPptx $wholePdf $Script:PowerPointRenderTimeoutSeconds
            if ($ProgressCallback) { & $ProgressCallback 'split' $SourceId '' }
            $split = Convert-PowerPointSplitPages (Invoke-PdfSourceSplit $wholePdf $contentDir) $contentDir
            if ([int]$pptInspection.pageCount -ne [int]$split.pageCount) { throw "POWERPOINT_SLIDE_COUNT_MISMATCH:$([int]$pptInspection.pageCount):$([int]$split.pageCount)" }
            $rendered = @(); $inspection = @()
            foreach ($pageInfo in @($split.pages)) {
                $number = [int]$pageInfo.pageNumber; $name = "Slide $number"
                $title = "$([IO.Path]::GetFileNameWithoutExtension([string]$source.fileName)) / $name"
                $inspection += [pscustomobject][ordered]@{ sheetName=$name; sheetIndex=$number; detectedTitle=$title }
                $rendered += [pscustomobject][ordered]@{ sheetName=$name; pdf=[string]$pageInfo.pdf; pageCount=1; warnings=@() }
            }
            $pageSync = Update-StructureLocked $Language {
                param($st)
                $latest = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq $SourceId } | Select-Object -First 1)
                if ($latest.Count -eq 0) { throw "登録済みPowerPoint原稿が見つかりません: $SourceId" }
                $w=$latest[0]; $sync=Update-WorkbookPagesFromInspection $Language $st $w $inspection; $pages=@(Get-Array $st.pages)
                foreach ($r in $rendered) {
                    $page = @($pages | Where-Object { [string]$_.workbookId -eq $SourceId -and [string]$_.sheetName -eq [string]$r.sheetName } | Select-Object -First 1)
                    if ($page.Count -gt 0) {
                        Set-NoteProperty $page[0] 'contentPdf' (Get-RelativePathCompat $workspace ([string]$r.pdf))
                        Set-NoteProperty $page[0] 'status' 'rendered'; Set-NoteProperty $page[0] 'warnings' @(); Set-NoteProperty $page[0] 'updatedAt' (New-NowIso)
                    }
                }
                Set-NoteProperty $w 'sourceType' 'powerpoint'; Set-NoteProperty $w 'adapterId' 'powerpoint-com-v1'; Set-NoteProperty $w 'sourcePageCount' ([int]$split.pageCount)
                Set-NoteProperty $w 'lastRenderedVersionId' $versionId; Set-NoteProperty $w 'lastRenderedSnapshotId' $SourceSnapshotId
                Set-NoteProperty $w 'lastRenderedExcelHash' $sourceHash; Set-NoteProperty $w 'lastRenderedAt' (New-NowIso)
                Set-NoteProperty $w 'renderProfileVersion' $Script:PowerPointRenderProfileVersion; Set-WorkbookRenderedSheetSnapshot $w (Get-DataProperty $sync 'sheetNames' @())
                Set-NoteProperty $w 'lastError' ''; Set-NoteProperty $w 'lastErrorUser' ''; Set-NoteProperty $w 'lastErrorAt' $null; Set-NoteProperty $w 'lastRenderAttemptHash' $sourceHash
                $currentHash = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
                if ([string]::IsNullOrWhiteSpace($currentHash) -or $currentHash -eq $sourceHash) { Set-NoteProperty $w 'status' 'rendered-unchecked' }
                else { Set-NoteProperty $w 'status' 'source-updated'; foreach ($p in @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $SourceId })) { Set-NoteProperty $p 'status' 'stale' } }
                $cat=Get-WorkbookPackId $w
                $vols=@(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $SourceId } | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
                Mark-VolumeNeedsRebuild $st $Language $cat $vols 'render' 'PowerPoint原稿を1件PDF変換しました'
                return $sync
            }
            $powerPointVersion = [string](Get-DataProperty $powerPointResult 'powerPointVersion' '')
            $Script:CurrentRenderEnvFingerprint = Get-Sha256Text "powerpoint-com|$powerPointVersion|$($Script:PowerPointRenderProfileVersion)"
            $Script:CurrentRenderEnvInfo = [ordered]@{ adapterId='powerpoint-com-v1'; powerPointVersion=$powerPointVersion; profileVersion=$Script:PowerPointRenderProfileVersion }
            $Script:PendingAnalysis = [ordered]@{ language=$Language; workbookId=$SourceId; snapshotId=[string]$SourceSnapshotId; versionId=$versionId; rendered=$rendered }
            Remove-WorkbookContentPdfs $workspace $SourceId $versionId
            $success = $true
            return [ordered]@{ workbookId=$SourceId; sourceId=$SourceId; sourceType='powerpoint'; adapterId='powerpoint-com-v1'; versionId=$versionId; rendered=@($rendered); warnings=@(); steps=@('PowerPointを一時コピーからPDF変換','PDFをスライド単位で取り込み'); sheetSync=$pageSync }
        } finally {
            if (-not $success -and (Test-Path -LiteralPath $contentDir)) { Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (-not [string]::IsNullOrWhiteSpace($ephemeralCaptureId)) { Remove-EphemeralCopy $Language $ephemeralCaptureId }
        }
    }
}

function Invoke-PowerPointToPdf([string]$InputPath, [string]$OutputPath, [int]$TimeoutSeconds = 0) {
    if ($TimeoutSeconds -le 0) { $TimeoutSeconds = $Script:PowerPointRenderTimeoutSeconds }
    $worker = Join-Path $Script:AppRoot 'tools\powerpoint-render-worker.ps1'
    if (-not (Test-Path -LiteralPath $worker -PathType Leaf)) { throw 'PowerPoint変換ワーカーが見つかりません。' }
    $runId = New-RbId
    $runDir = Split-Path -Parent $OutputPath
    if (-not (Test-Path -LiteralPath $runDir)) { New-Item -ItemType Directory -Path $runDir -Force | Out-Null }
    $resultPath = Join-Path $runDir "powerpoint-result-$runId.json"
    $stdoutPath = Join-Path $runDir "powerpoint-out-$runId.log"
    $stderrPath = Join-Path $runDir "powerpoint-err-$runId.log"
    $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
    $workerEscaped = $worker.Replace("'", "''")
    $command = "& '$workerEscaped' -InputPathB64 '$(ConvertTo-PathBase64 $InputPath)' -OutputPathB64 '$(ConvertTo-PathBase64 $OutputPath)' -ResultPathB64 '$(ConvertTo-PathBase64 $resultPath)'"
    $proc = $null
    try {
        $proc = Start-HiddenPowerShellChild $psExe $command $stdoutPath $stderrPath
        if ($null -eq $proc) { throw 'POWERPOINT_WORKER_START_FAILED' }
        if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
            Stop-ReportBinderProcessTree ([int]$proc.Id)
            throw "POWERPOINT_TIMEOUT:$TimeoutSeconds"
        }
        try { $proc.WaitForExit() } catch { }
        $result = Read-JsonFile $resultPath $null
        if ($null -eq $result) {
            $details = @()
            if (Test-Path -LiteralPath $stderrPath) { $details += (Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue) }
            if (Test-Path -LiteralPath $stdoutPath) { $details += (Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue) }
            throw ('POWERPOINT_WORKER_NO_RESULT:' + ([string]::Join(' ', @($details))).Trim())
        }
        if (-not [bool](Get-DataProperty $result 'ok' $false)) {
            throw ("$([string](Get-DataProperty $result 'errorCode' 'POWERPOINT_WORKER_FAILED')):$([string](Get-DataProperty $result 'stage' '')):$([string](Get-DataProperty $result 'message' ''))")
        }
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf) -or (Get-Item -LiteralPath $OutputPath).Length -le 0 -or -not (Test-PdfMagic $OutputPath)) { throw 'POWERPOINT_PDF_INVALID' }
        return $result
    } finally {
        foreach ($path in @($resultPath,$stdoutPath,$stderrPath)) { if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } }
    }
}

function Register-PowerPointSource([string]$Language, [string]$RelativePath, [string]$Category = '', [string]$PackId = '') {
    Test-DirectPowerPointRelativePath $RelativePath | Out-Null
    $cat = Require-WorkbookCategory $Category
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $inspection = Inspect-PowerPointSourceFile $full
    $candidate = New-PowerPointSourceObject $RelativePath $Language $cat ([int]$inspection.pageCount)
    $resolvedPackId = if ([string]::IsNullOrWhiteSpace($PackId)) { Get-BuiltinPackId $cat } else { ([string]$PackId).Trim() }
    if ([string]::IsNullOrWhiteSpace((Get-CategoryFromBuiltinPackId $resolvedPackId))) { Set-NoteProperty $candidate 'workbookId' ("$resolvedPackId-$([string]$candidate.workbookId)") }
    Set-NoteProperty $candidate 'packId' $resolvedPackId
    return Update-StructureLocked $Language {
        param($structure)
        $normalized = ([string]$RelativePath -replace '\\','/').ToLowerInvariant()
        $existing = @(Get-Array $structure.workbooks | Where-Object {
            ([string]$_.workbookId -eq [string]$candidate.workbookId) -or ((([string]$_.relativePath -replace '\\','/').ToLowerInvariant() -eq $normalized) -and (Test-WorkbookPack $_ $resolvedPackId))
        } | Select-Object -First 1)
        if ($existing.Count -gt 0) {
            $old = $existing[0]
            Set-NoteProperty $candidate 'workbookId' ([string]$old.workbookId)
            foreach ($name in @('lastRenderedVersionId','lastRenderedSnapshotId','lastRenderedExcelHash','lastRenderedAt','lastRenderedSheets','lastRenderedSheetFingerprint','lastRenderLog','renderProfileVersion','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash','ownerDepartment','required','defaultTargetId','requirementId','sourcePageCount')) {
                Set-NoteProperty $candidate $name (Get-DataProperty $old $name $null)
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$candidate.lastRenderedExcelHash)) {
                Set-NoteProperty $candidate 'status' $(if ($candidate.currentExcelHash -ne $candidate.lastRenderedExcelHash) { 'source-updated' } else { [string]$old.status })
            }
        }
        $structure.workbooks = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -ne [string]$candidate.workbookId }) + @($candidate)
        return [ordered]@{ workbook=$candidate; units=@($inspection.units); registered=$true; inspected=$true; sourceType='powerpoint' }
    }
}

function Get-AllFinalReadiness($Structure,[string]$Language,[bool]$CheckOutputExists = $true) {
    $result=[ordered]@{};foreach($cat in @('ecm','bod','dmm')){$vols=[ordered]@{};foreach($volume in @(Get-VolumeList $Language|Where-Object{$_ -ne 'none'})){$vols[$volume]=Get-FinalBuildReadiness $Structure $Language $volume $cat $CheckOutputExists};$result[$cat]=[ordered]@{volumes=$vols}};return $result
}

function Get-PackProgressDashboard($Structure, [string]$Language) {
    $rows = @()
    $today = (Get-Date).Date
    $dueSoonLimit = $today.AddDays(7)
    foreach ($pack in @(Get-PublicPackList $Structure $Language $false)) {
        $packId = [string](Get-DataProperty $pack 'packId' '')
        $workbooks = @(Get-Array (Get-DataProperty $Structure 'workbooks' @()) | Where-Object { Test-WorkbookPack $_ $packId })
        $workbookIds = @{}; foreach ($workbook in $workbooks) { $workbookIds[[string](Get-DataProperty $workbook 'workbookId' '')] = $true }
        $pages = @(Get-Array (Get-DataProperty $Structure 'pages' @()) | Where-Object { $workbookIds.ContainsKey([string](Get-DataProperty $_ 'workbookId' '')) })
        $requirements = @(Get-Array (Get-DataProperty $pack 'sourceRequirements' @()))
        $required = @($requirements | Where-Object { [bool](Get-DataProperty $_ 'required' $true) })
        $submittedRequired = 0
        $overdueRequired = 0
        $dueSoonRequired = 0
        $nearestDueDate = ''
        foreach ($requirement in $required) {
            $requirementId = [string](Get-DataProperty $requirement 'requirementId' '')
            $submitted = @($workbooks | Where-Object { [string](Get-DataProperty $_ 'requirementId' '') -eq $requirementId -and [string](Get-DataProperty $_ 'status' '') -ne 'missing' }).Count -gt 0
            if ($submitted) { $submittedRequired++; continue }
            $dueDateText = [string](Get-DataProperty $requirement 'dueDate' '')
            if ([string]::IsNullOrWhiteSpace($dueDateText)) { continue }
            try {
                $dueDate = [DateTime]::ParseExact($dueDateText, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).Date
                if ([string]::IsNullOrWhiteSpace($nearestDueDate) -or [string]::CompareOrdinal($dueDateText, $nearestDueDate) -lt 0) { $nearestDueDate = $dueDateText }
                if ($dueDate -lt $today) { $overdueRequired++ }
                elseif ($dueDate -le $dueSoonLimit) { $dueSoonRequired++ }
            } catch { }
        }
        # The dashboard and final-build checks must share one freshness rule.
        # A duplicated approximation used to label a workbook "変換待ち" while
        # the workbook card correctly said its conversion PDF was current.
        $needsRender = @($workbooks | Where-Object { -not (Test-WorkbookRenderIsCurrent $_) }).Count

        $unassigned = @($pages | Where-Object { [bool](Get-DataProperty $_ 'enabled' $true) -eq $false -or [string](Get-DataProperty $_ 'volume' 'none') -eq 'none' }).Count
        $targetRows = @(); $allBlockers = @(); $targetFingerprints = [ordered]@{}
        foreach ($target in @(Get-Array (Get-DataProperty $pack 'targets' @()))) {
            $targetId = [string](Get-DataProperty $target 'targetId' '')
            try {
                $readiness = Get-FinalBuildReadiness $Structure $Language (Get-LegacyVolumeFromTargetId $Language $targetId) $packId $false
                $targetFingerprints[$targetId] = [string]$readiness.currentFingerprint
                $targetRows += [pscustomobject][ordered]@{ targetId=$targetId; displayName=[string](Get-DataProperty $target 'displayName' $targetId); required=[bool](Get-DataProperty $target 'required' $false); pageCount=[int]$readiness.pageCount; displayState=[string]$readiness.displayState; canBuild=[bool]$readiness.canBuild }
                $allBlockers += @(Get-Array $readiness.blockers | Where-Object { [string](Get-DataProperty $_ 'code' '') -ne 'no-pages' })
            } catch {
                $targetRows += [pscustomobject][ordered]@{ targetId=$targetId; displayName=[string](Get-DataProperty $target 'displayName' $targetId); required=[bool](Get-DataProperty $target 'required' $false); pageCount=0; displayState='blocked'; canBuild=$false }
                $allBlockers += [pscustomobject][ordered]@{ code='readiness-error'; message=$_.Exception.Message }
            }
        }
        $blockerKeys = @{}; $blockers = @()
        foreach ($blocker in $allBlockers) {
            $key = @([string](Get-DataProperty $blocker 'code' ''),[string](Get-DataProperty $blocker 'sourceId' ''),[string](Get-DataProperty $blocker 'requirementId' '')) -join '|'
            if ($blockerKeys.ContainsKey($key)) { continue }; $blockerKeys[$key] = $true; $blockers += $blocker
        }
        $activeTargets = @($targetRows | Where-Object { [int]$_.pageCount -gt 0 })
        $state = 'complete'; $nextAction = 'final'
        if ($workbooks.Count -eq 0) { $state='not-started'; $nextAction='excel' }
        if ($overdueRequired -gt 0) { $state='overdue-source'; $nextAction='excel' }
        elseif ($submittedRequired -lt $required.Count) { $state='missing-source'; $nextAction='excel' }
        elseif ($needsRender -gt 0) { $state='needs-render'; $nextAction='excel' }
        elseif ($unassigned -gt 0) { $state='unassigned'; $nextAction='pages' }
        elseif ($blockers.Count -gt 0) { $state='blocked'; $nextAction='excel' }
        elseif ($activeTargets.Count -eq 0) { $state='no-pages'; $nextAction='pages' }
        elseif (@($activeTargets | Where-Object { [string]$_.displayState -in @('needs-rebuild','output-missing','not-built') }).Count -gt 0) { $state='needs-output'; $nextAction='final' }
        $storedPack = Get-PackRecord $Structure $packId
        $storedReview = Get-DataProperty $storedPack 'review' ([ordered]@{})
        $reviewStoredStatus = [string](Get-DataProperty $storedReview 'status' 'draft')
        if ($reviewStoredStatus -notin @('draft','in-review','approved','changes-requested')) { $reviewStoredStatus = 'draft' }
        $reviewStatus = $reviewStoredStatus
        if ($reviewStoredStatus -in @('in-review','approved')) {
            $submittedFingerprints = Get-DataProperty $storedReview 'submittedFingerprints' ([ordered]@{})
            $reviewTargets = @($targetRows | Where-Object { [bool]$_.required })
            if ($reviewTargets.Count -eq 0) { $reviewTargets = @($targetRows | Select-Object -First 1) }
            foreach ($reviewTarget in $reviewTargets) {
                $reviewTargetId = [string]$reviewTarget.targetId
                if ([string](Get-DataProperty $submittedFingerprints $reviewTargetId '') -ne [string]$targetFingerprints[$reviewTargetId]) { $reviewStatus = 'stale'; break }
            }
        }
        # 確認の記録は %LOCALAPPDATA% 配下の利用者ごとのデータにしか残らず、他の
        # 利用者からは参照できない。作業の進み具合(原稿・変換PDF・ページ構成・出力)
        # とは別物なので、未確認を理由に complete から降格させない。確認の状態は
        # reviewStatus として別に返し、画面側で個別に表示する。
        if ($state -eq 'complete' -and $reviewStatus -eq 'stale') { $nextAction = 'final' }
        $rows += [pscustomobject][ordered]@{
            packId=$packId; displayName=[string](Get-DataProperty $pack 'displayName' ''); category=[string](Get-DataProperty $pack 'category' '')
            sourceCount=$workbooks.Count; requiredSourceCount=$required.Count; submittedRequiredSourceCount=$submittedRequired; overdueRequiredSourceCount=$overdueRequired; dueSoonRequiredSourceCount=$dueSoonRequired; nearestRequiredDueDate=$nearestDueDate; needsRenderCount=$needsRender
            pageCount=$pages.Count; unassignedPageCount=$unassigned; blockerCount=$blockers.Count; blockers=@($blockers | Select-Object -First 3)
            targets=@($targetRows); state=$state; nextAction=$nextAction; reviewStatus=$reviewStatus; reviewStoredStatus=$reviewStoredStatus; reviewEventCount=@(Get-Array (Get-DataProperty $storedReview 'events' @())).Count; updatedAt=Get-DataProperty $pack 'updatedAt' $null
        }
    }
    $attention = @($rows | Where-Object { [string]$_.state -ne 'complete' })
    return [pscustomobject][ordered]@{
        packs=@($rows); totalCount=$rows.Count; completeCount=@($rows | Where-Object { [string]$_.state -eq 'complete' }).Count
        attentionCount=$attention.Count; missingRequiredCount=[int](($rows | ForEach-Object { [Math]::Max(0,[int]$_.requiredSourceCount-[int]$_.submittedRequiredSourceCount) } | Measure-Object -Sum).Sum)
        overdueRequiredCount=[int](($rows | ForEach-Object { [int]$_.overdueRequiredSourceCount } | Measure-Object -Sum).Sum)
        dueSoonRequiredCount=[int](($rows | ForEach-Object { [int]$_.dueSoonRequiredSourceCount } | Measure-Object -Sum).Sum)
        needsRenderCount=[int](($rows | ForEach-Object { [int]$_.needsRenderCount } | Measure-Object -Sum).Sum)
        unassignedPageCount=[int](($rows | ForEach-Object { [int]$_.unassignedPageCount } | Measure-Object -Sum).Sum)
    }
}

function Build-FinalPdf([string]$Language,[string]$Volume,[string]$Category='') {
    # category は fail closed。ここでも明示的に検証する。
    $cat = Require-WorkbookCategory $Category
    # 単体出力もまとめて出力も、履歴とアーカイブを持つ同じトランザクションエンジンを通す。
    $r = Invoke-FinalBuildTransaction $Language $cat @($Volume)
    $built = @(Get-Array $r.built)
    if ($built.Count -eq 0) { throw [InvalidOperationException]::new('対象ページがありません。ページ構成を確認してください。') }
    return $built[0]
}

function Add-FinalSnapshotSourceWorkbooks($Structure, $Snapshot) {
    $seen = @{}
    $sourceWorkbooks = @()
    foreach ($page in @(Get-Array (Get-DataProperty $Snapshot 'pages' @()))) {
        $workbookId = [string](Get-DataProperty $page 'workbookId' '')
        if ([string]::IsNullOrWhiteSpace($workbookId) -or $seen.ContainsKey($workbookId)) { continue }
        $seen[$workbookId] = $true
        $workbook = @(Get-Array $Structure.workbooks | Where-Object { [string]$_.workbookId -eq $workbookId } | Select-Object -First 1)
        if ($workbook.Count -eq 0) { continue }
        $w = $workbook[0]
        $sourceWorkbooks += [ordered]@{
            workbookId = $workbookId
            fileName = [string](Get-DataProperty $w 'fileName' '')
            relativePath = [string](Get-DataProperty $w 'relativePath' '')
            currentExcelLastWriteUtcTicks = [string](Get-DataProperty $w 'currentExcelLastWriteUtcTicks' '')
            currentExcelSize = [int64](Get-DataProperty $w 'currentExcelSize' -1)
            currentExcelHash = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
            snapshotId = [string](Get-DataProperty $w 'lastRenderedSnapshotId' '')
            versionId = [string](Get-DataProperty $w 'lastRenderedVersionId' '')
            sourceHash = Normalize-FileHash ([string](Get-DataProperty $w 'lastRenderedExcelHash' ''))
            renderEnvironmentFingerprint = [string](Get-DataProperty $w 'renderEnvironmentFingerprint' '')
            renderEnvironment = Get-DataProperty $w 'renderEnvironment' $null
        }
    }
    Set-NoteProperty $Snapshot 'sourceWorkbooks' @($sourceWorkbooks)
    return $Snapshot
}

function Build-DocumentPackPdf([string]$Language, [string]$PackId, [string]$TargetId) {
    $structure = Get-Structure $Language
    $scope = Resolve-DocumentPackScope $structure $PackId $false
    $TargetId = Assert-PackTargetId $Language $scope.pack $TargetId
    $volume = Get-LegacyVolumeFromTargetId $Language $TargetId
    if ([bool]$scope.builtIn) { return Build-FinalPdf $Language $volume ([string]$scope.category) }

    $paths = Get-Paths
    $workspace = Get-WorkspacePath $Language
    Report-FinalBuildPhase '元原稿の更新を確認しています'
    try { [void](Scan-Updates $Language $null $false) } catch { }
    $lockPath = Join-Path $workspace ("locks\pack-output_{0}_{1}.lock" -f ([string]$scope.packId), $TargetId)
    return Invoke-WithLock $lockPath {
        $composerJar = Join-Path $Script:AppRoot 'lib\pdfbox\ReportPdfComposer.jar'
        $pdfboxJar = Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'
        if (-not (Test-Path $composerJar)) { throw 'ReportPdfComposer.jar がありません。' }
        if (-not (Test-Path $pdfboxJar)) { throw 'pdfbox-app.jar がありません。' }
        $snapshotBefore = Update-StructureLocked $Language {
            param($st)
            Apply-DefaultNumberingPerVolume $Language $st ([string]$scope.packId)
            return Get-FinalBuildInputSnapshot $st $Language $volume ([string]$scope.packId)
        }
        [void](Add-FinalSnapshotSourceWorkbooks (Get-Structure $Language) $snapshotBefore)
        if ($snapshotBefore.blockers.Count -gt 0) { throw [InvalidOperationException]::new([string]$snapshotBefore.blockers[0].message) }
        $fingerprint = [string]$snapshotBefore.fingerprint
        $outputName = [string]$snapshotBefore.outputFileName
        $outputPath = Join-Path ([string]$paths.outputDir) $outputName
        $tempPath = Join-Path ([string]$paths.outputDir) ("~building_{0}_{1}.pdf" -f ([string]$scope.packId), $TargetId)
        if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $outputPath) {
            $handle = $null
            try { $handle = [IO.File]::Open($outputPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None) }
            catch { throw "出力先の提出用PDFが開かれているため上書きできません: $outputName" }
            finally { if ($handle) { $handle.Dispose() } }
        }
        $manifest = [ordered]@{
            schemaVersion=3; language=$Language; packId=[string]$scope.packId; targetId=$TargetId; volume=$volume
            projectId=[string]$snapshotBefore.projectId; inputFingerprint=$fingerprint; outputPdf=$tempPath; createdAt=New-NowIso
            document=$snapshotBefore.document; pageNumber=[ordered]@{font='Arial';fontSize=8;bottomPt=18;format='hyphenated';countHidden=$true}
            physicalPages=$snapshotBefore.physicalPages; pages=$snapshotBefore.manifestPages
        }
        $manifestPath = Join-Path $workspace ("exports\manifest_{0}_{1}.json" -f ([string]$scope.packId), $TargetId)
        Write-JsonFile $manifestPath $manifest
        Report-FinalBuildPhase ('ページを結合しています（' + [string]$snapshotBefore.pageCount + 'ページ）')
        $java = Resolve-JavaExe
        $run = Invoke-NativeCapture $java @('-cp',"$composerJar;$pdfboxJar",'ReportPdfComposer','--manifest',$manifestPath) $Script:FinalComposeTimeoutSeconds
        if ($run.timedOut) { throw ('PDFの結合が{0}分以内に終わりませんでした。出力先がネットワーク上のフォルダーの場合は、いったんPC内のフォルダーに出力してみてください。' -f [int]($Script:FinalComposeTimeoutSeconds / 60)) }
        if ([int]$run.exitCode -ne 0) { throw "PDFBox組版に失敗しました。exit=$([int]$run.exitCode)`n$([string]$run.text)" }
        if (-not (Test-Path $tempPath) -or (Get-Item $tempPath).Length -le 0) { throw '提出用PDFを作成できませんでした。' }
        Report-FinalBuildPhase '出力先へ保存しています'
        $buildId = New-RbId
        $commit = Update-StructureLocked $Language {
            param($st)
            $after = Get-FinalBuildInputSnapshot $st $Language $volume ([string]$scope.packId)
            if ([string]$after.fingerprint -ne $fingerprint) { return [ordered]@{ changed=$true } }
            Move-Item -LiteralPath $tempPath -Destination $outputPath -Force
            $state = Get-PackOutputState $st $Language ([string]$scope.packId) $volume $true
            Set-NoteProperty $state 'builtFingerprint' $fingerprint
            Set-NoteProperty $state 'lastBuiltAt' (New-NowIso)
            Set-NoteProperty $state 'outputPdf' $outputPath
            Set-NoteProperty $state 'staleReasons' @()
            Set-NoteProperty $state 'message' ([string]$run.text)
            Set-NoteProperty $state 'status' 'built'
            Set-NoteProperty $state 'buildId' $buildId
            return [ordered]@{ changed=$false; readiness=(Get-FinalBuildReadiness $st $Language $volume ([string]$scope.packId)) }
        }
        if ([bool]$commit.changed) {
            if (Test-Path -LiteralPath $tempPath) { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue }
            throw 'PDF作成中にページ構成またはPDF入力が変更されました。最新の状態で再度出力してください。'
        }
        $manifest.outputPdf = $outputPath
        Write-JsonFile $manifestPath $manifest
        $archivePath = ''
        $archiveError = ''
        try { $archivePath = New-FinalArchive $Language ([string]$scope.packId) $volume $buildId $outputPath $manifest $snapshotBefore }
        catch {
            $archiveError = $_.Exception.Message
            # PDFは既に出力先へ書けているので、ここで作成そのものを失敗にはしない。
            # ただしアーカイブが作れなかったときに黙って進むと、出力の元になった版が
            # 保持期間の猶予なしに掃除で消える。保護(pin)だけは作り直す。
            try {
                Set-FinalPdfSnapshotPins $Language ([string]$scope.packId) '' $volume $buildId (Get-DataProperty $snapshotBefore 'sourceWorkbooks' @()) ''
            } catch {
                $archiveError = $archiveError + ' / ' + $_.Exception.Message
            }
            # 握りつぶさない。利用者にも記録にも残す。
            Write-HistoryEvent $Language 'final.archive.failed' ([ordered]@{ packId=[string]$scope.packId; targetId=$TargetId; volume=$volume; buildId=$buildId; message=$archiveError })
        }
        Write-HistoryEvent $Language 'final.built' ([ordered]@{ packId=[string]$scope.packId; category=''; targetId=$TargetId; volume=$volume; buildId=$buildId; outputPdf=$outputPath; archivePath=$archivePath; archiveError=$archiveError })
        return [ordered]@{ packId=[string]$scope.packId; targetId=$TargetId; volume=$volume; outputPdf=$outputPath; inputFingerprint=$fingerprint; buildId=$buildId; archivePath=$archivePath; archiveError=$archiveError; message=[string]$run.text; readiness=$commit.readiness }
    }
}

function Get-StatePayload([string]$Language) {
    $paths = Get-Paths
    $configured = $false
    if ($paths -and [string]$paths.submissionDir -and [string]$paths.dataDir -and [string]$paths.outputDir) { $configured = $true }
    $structure = $null
    $structureLoadError = ''
    if ($configured -and (Test-Path -LiteralPath ([string]$paths.dataDir))) {
        try {
            $structure = Get-Structure $Language
        } catch {
            $structureLoadError = $_.Exception.Message
            $structure = New-EmptyStructure $Language
        }
    } else {
        $structure = New-EmptyStructure $Language
    }
    $workbooks = @(Get-Array $structure.workbooks)
    $pages = @(Get-Array $structure.pages)
    $summary = [ordered]@{
        excelUpdated = @($workbooks | Where-Object { [string]$_.status -in @('excel-updated','source-updated') }).Count
        uncheckedPages = @($pages | Where-Object { [string]$_.status -in @('rendered','stale','not-rendered') }).Count
        confirmedPages = @($pages | Where-Object { [string]$_.status -eq 'confirmed' }).Count
        renderErrors = @($workbooks | Where-Object { [string]$_.status -eq 'render-error' }).Count
        pdfReadyWorkbooks = @($workbooks | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.lastRenderedExcelHash) -and [string]$_.status -ne 'render-error' }).Count
        pdfPendingWorkbooks = @($workbooks | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.lastRenderedExcelHash) -or [string]$_.status -in @('new','excel-updated','source-updated','render-error') }).Count
        pdfReadyPages = @($pages | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.contentPdf) }).Count
        totalWorkbooks = $workbooks.Count
        totalPages = $pages.Count
    }
    if ($configured -and -not $structureLoadError) {
        foreach ($cat in @('ecm','bod','dmm')) {
            foreach ($volume in @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })) {
                $key=Get-VolumeStateKey $volume $cat; $v=Get-DataProperty $structure.volumes $key $null
                if ($null -ne $v) { $out=[string](Get-DataProperty $v 'outputPdf' ''); Set-NoteProperty $v 'outputPdfExists' (-not [string]::IsNullOrWhiteSpace($out)) }
            }
        }
    }
    # 初期表示では共有フォルダー上の比較JSONをExcel件数分読まない。
    # PDF作成時にstructureへ保存した小さな要約だけを返す。
    $changeSummaries = [ordered]@{}
    $inputHistoryOn = $false
    try { $inputHistoryOn = (Test-InputHistoryEnabled) } catch { }
    if ($configured -and -not $structureLoadError -and $inputHistoryOn) {
        foreach ($w in $workbooks) {
            $cs = Get-DataProperty $w 'latestComparisonSummary' $null
            if ($null -ne $cs) { $changeSummaries[[string]$w.workbookId] = $cs }
        }
    }
    $autoSummary = $null
    try { if ($configured) { $autoSummary = Get-AutoStateSummary $Language -Fast } } catch { }

    $pdfjsDir = Join-Path $Script:WebRoot 'pdfjs'
    $pdfjsClassic = ((Test-Path -LiteralPath (Join-Path $pdfjsDir 'pdf.min.js')) -and (Test-Path -LiteralPath (Join-Path $pdfjsDir 'pdf.worker.min.js')))
    $pdfjsModule = ((Test-Path -LiteralPath (Join-Path $pdfjsDir 'pdf.min.mjs')) -and (Test-Path -LiteralPath (Join-Path $pdfjsDir 'pdf.worker.min.mjs')))
    $pdfjsMode = 'none'
    if ($pdfjsModule) { $pdfjsMode = 'module' } elseif ($pdfjsClassic) { $pdfjsMode = 'classic' }
    $packTemplates = @(Get-PackTemplateCatalog $Language)
    $publicPacks = @(Get-PublicPackList $structure $Language)
    $packProgress = [pscustomobject][ordered]@{ packs=@(); totalCount=0; completeCount=0; attentionCount=0; missingRequiredCount=0; overdueRequiredCount=0; dueSoonRequiredCount=0; needsRenderCount=0; unassignedPageCount=0 }
    if ($configured -and -not $structureLoadError) { try { $packProgress = Get-PackProgressDashboard $structure $Language } catch { Write-Warning ('一式進捗を集計できません: ' + $_.Exception.Message) } }
    # ページ構成の楽観ロック用。画面はこれを baseLayout として送り返す。
    $layoutFingerprints = [ordered]@{}
    if ($configured -and -not $structureLoadError) {
        try { foreach ($pack in $publicPacks) { $layoutFingerprints[[string]$pack.packId] = Get-PageLayoutFingerprint $structure ([string]$pack.packId) } }
        catch { Write-Warning ('ページ構成の指紋を計算できません: ' + $_.Exception.Message) }
    }
    return [ordered]@{
        ok = $true
        token = $Script:Token
        configured = $configured
        paths = $paths
        packTemplates = $packTemplates
        packs = $publicPacks
        structure = ConvertTo-V4StructureCompatibilityView $structure
        layoutFingerprints = $layoutFingerprints
        structureLoadError = $structureLoadError
        finalReadiness = $(if ($configured -and -not $structureLoadError) { Get-AllFinalReadiness $structure $Language $false } else { [ordered]@{} })
        packProgress = $packProgress
        summary = $summary
        recentErrors = @(Get-Array $workbooks | Where-Object { [string]$_.status -eq 'render-error' -or -not [string]::IsNullOrWhiteSpace([string]$_.lastError) } | ForEach-Object { [ordered]@{ workbookId = [string]$_.workbookId; fileName = [string]$_.fileName; displayName = [string]$_.displayName; message = [string]$_.lastErrorUser; detail = [string]$_.lastError; at = [string]$_.lastErrorAt } })
        pdfjsPresent = ($pdfjsClassic -or $pdfjsModule)
        pdfjsMode = $pdfjsMode
        excelPrintProfileVersion = $Script:ExcelPrintProfileVersion
        pdfImportProfileVersion = $Script:PdfImportProfileVersion
        wordRenderProfileVersion = $Script:WordRenderProfileVersion
        powerPointRenderProfileVersion = $Script:PowerPointRenderProfileVersion
        inputHistoryEnabled = $inputHistoryOn
        changeSummaries = $changeSummaries
        auto = $autoSummary
        autoRenderInProgress = $Script:AutoRenderInProgress
        shutdownOnTabClose = $false
    }
}

function Get-V2StatePayload([string]$Language) {
    $legacy = Get-StatePayload $Language
    $configured = [bool](Get-DataProperty $legacy 'configured' $false)
    $structureLoadError = [string](Get-DataProperty $legacy 'structureLoadError' '')
    $structure = if ($configured -and [string]::IsNullOrWhiteSpace($structureLoadError)) { Get-Structure $Language } else { New-EmptyStructure $Language }
    $publicPacks = @(Get-PublicPackList $structure $Language)
    # The legacy state contains readiness only for the three migration-era
    # categories.  V2 mutations must also carry the active custom pack's
    # fingerprint; otherwise the browser has to guess that a reorder is stale.
    $finalReadiness = Get-DataProperty $legacy 'finalReadiness' ([ordered]@{})
    if ($configured -and [string]::IsNullOrWhiteSpace($structureLoadError)) {
        foreach ($pack in $publicPacks) {
            $packId = [string](Get-DataProperty $pack 'packId' '')
            if ([string]::IsNullOrWhiteSpace($packId)) { continue }
            $volumes = [ordered]@{}
            foreach ($targetId in @(Get-PackTargetIds $Language $pack)) {
                $volume = Get-LegacyVolumeFromTargetId $Language $targetId
                try { $volumes[$volume] = Get-FinalBuildReadiness $structure $Language $volume $packId $true }
                catch { }
            }
            Set-NoteProperty $finalReadiness $packId ([ordered]@{ volumes = $volumes })
        }
    }
    $sources = @()
    foreach ($source in @(Get-Array (Get-DataProperty $structure 'sources' @()))) {
        $sources += [pscustomobject][ordered]@{
            sourceId = [string](Get-DataProperty $source 'sourceId' '')
            packId = [string](Get-DataProperty $source 'packId' '')
            relativePath = [string](Get-DataProperty $source 'relativePath' '')
            sourceType = [string](Get-DataProperty $source 'sourceType' '')
            adapterId = [string](Get-DataProperty $source 'adapterId' '')
            displayName = [string](Get-DataProperty $source 'displayName' '')
            ownerDepartment = [string](Get-DataProperty $source 'ownerDepartment' '')
            required = [bool](Get-DataProperty $source 'required' $true)
            defaultTargetId = [string](Get-DataProperty $source 'defaultTargetId' 'unassigned')
            requirementId = [string](Get-DataProperty $source 'requirementId' '')
            excelSheetSelectionMode = Normalize-ExcelSheetSelection ([string](Get-DataProperty $source 'excelSheetSelectionMode' 'all-visible'))
            status = [string](Get-DataProperty $source 'status' '')
            currentSnapshotId = [string](Get-DataProperty $source 'currentSnapshotId' '')
            lastRenderedSnapshotId = [string](Get-DataProperty $source 'lastRenderedSnapshotId' '')
            lastRenderedVersionId = [string](Get-DataProperty $source 'lastRenderedVersionId' '')
            lastRenderedAt = Get-DataProperty $source 'lastRenderedAt' $null
        }
    }
    return [ordered]@{
        ok = $true
        apiVersion = 2
        domainSchemaVersion = 3
        token = Get-DataProperty $legacy 'token' $Script:Token
        configured = $configured
        paths = Get-DataProperty $legacy 'paths' $null
        structureLoadError = $structureLoadError
        packTemplates = @(Get-PackTemplateCatalog $Language)
        packs = $publicPacks
        structure = [ordered]@{
            schemaVersion = 3
            packs = $publicPacks
            sources = @($sources)
            units = @(Get-Array (Get-DataProperty $structure 'units' @()))
            items = @(Get-Array (Get-DataProperty $structure 'items' @()))
            artifacts = @(Get-Array (Get-DataProperty $structure 'artifacts' @()))
            outputs = Get-DataProperty $structure 'outputs' ([ordered]@{})
            migrationIssues = @(Get-Array (Get-DataProperty $structure 'migrationIssues' @()))
            updatedAt = Get-DataProperty $structure 'updatedAt' $null
        }
        finalReadiness = $finalReadiness
        packProgress = Get-DataProperty $legacy 'packProgress' ([ordered]@{})
        summary = Get-DataProperty $legacy 'summary' ([ordered]@{})
        recentErrors = @(Get-Array (Get-DataProperty $legacy 'recentErrors' @()))
        inputHistoryEnabled = [bool](Get-DataProperty $legacy 'inputHistoryEnabled' $false)
        changeSummaries = Get-DataProperty $legacy 'changeSummaries' ([ordered]@{})
    }
}

function Read-BodyJson($Request) {
    $reader = New-Object IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    $text = $reader.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($text)) { return [pscustomobject]@{} }
    return $text | ConvertFrom-Json
}

function Touch-ResponseActivity {
    if ($Script:ClientAttached -and $Script:ClientCloseNotifiedUtc -eq [DateTime]::MinValue) { $Script:LastHeartbeatUtc = [DateTime]::UtcNow }
}

function Write-TextResponse($Context, [int]$Status, [string]$Body, [string]$ContentType, [bool]$AllowCors = $false) {
    Touch-ResponseActivity
    $bytes = [Text.Encoding]::UTF8.GetBytes($Body)
    if (Test-TcpContext $Context) { Write-TcpResponse $Context $Status $bytes $ContentType $AllowCors; return }
    $Context.Response.StatusCode = $Status
    $Context.Response.ContentType = $ContentType
    $Context.Response.Headers['Cache-Control'] = 'no-store'
    if ($AllowCors) { $Context.Response.Headers['Access-Control-Allow-Origin'] = '*' }
    $Context.Response.ContentLength64 = $bytes.Length
    $Context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $Context.Response.OutputStream.Close()
}

function Write-JsonResponse($Context, [int]$Status, $Object, [bool]$AllowCors = $false) {
    $json = ConvertTo-Json -InputObject $Object -Depth 50
    Write-TextResponse $Context $Status $json 'application/json; charset=utf-8' $AllowCors
}

function Write-BytesResponse($Context, [int]$Status, [byte[]]$Bytes, [string]$ContentType, [bool]$AllowCors = $false, [string]$CacheControl = 'no-store') {
    Touch-ResponseActivity
    if ([string]::IsNullOrWhiteSpace($CacheControl)) { $CacheControl = 'no-store' }
    $CacheControl = $CacheControl -replace "[\r\n]", ''
    if (Test-TcpContext $Context) { Write-TcpResponse $Context $Status $Bytes $ContentType $AllowCors $CacheControl; return }
    $Context.Response.StatusCode = $Status
    $Context.Response.ContentType = $ContentType
    $Context.Response.Headers['Cache-Control'] = $CacheControl
    if ($AllowCors) { $Context.Response.Headers['Access-Control-Allow-Origin'] = '*' }
    $Context.Response.ContentLength64 = $Bytes.Length
    $Context.Response.OutputStream.Write($Bytes, 0, $Bytes.Length)
    $Context.Response.OutputStream.Close()
}


function Write-FileResponse($Context, [int]$Status, [string]$FullPath, [string]$ContentType, [bool]$AllowCors = $false, [string]$CacheControl = 'no-store') {
    Touch-ResponseActivity
    if ([string]::IsNullOrWhiteSpace($CacheControl)) { $CacheControl = 'no-store' }
    $CacheControl = $CacheControl -replace "[\r\n]", ''
    $ContentType = ([string]$ContentType) -replace "[\r\n]", ''
    $fileInfo = [IO.FileInfo]::new((ConvertTo-Win32ExtendedPath $FullPath))
    if (-not $fileInfo.Exists) { throw '配信するファイルが見つかりません。' }
    $source = [IO.File]::Open($fileInfo.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    if (Test-TcpContext $Context) {
        try {
            $statusText = Get-HttpStatusText $Status
            $corsHeader = ''
            if ($AllowCors) { $corsHeader = "Access-Control-Allow-Origin: *`r`n" }
            $header = "HTTP/1.1 $Status $statusText`r`nContent-Type: $ContentType`r`nContent-Length: $($fileInfo.Length)`r`nCache-Control: $CacheControl`r`n${corsHeader}Connection: close`r`n`r`n"
            $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
            $target = $Context.TcpStream
            $target.Write($headerBytes, 0, $headerBytes.Length)
            $source.CopyTo($target, 65536)
            $target.Flush()
        } finally {
            $source.Dispose()
            try { if ($Context.TcpStream) { $Context.TcpStream.Close() } } catch {}
            try { if ($Context.TcpClient) { $Context.TcpClient.Close() } } catch {}
        }
        return
    }
    try {
        $Context.Response.StatusCode = $Status
        $Context.Response.ContentType = $ContentType
        $Context.Response.Headers['Cache-Control'] = $CacheControl
        if ($AllowCors) { $Context.Response.Headers['Access-Control-Allow-Origin'] = '*' }
        $Context.Response.ContentLength64 = $fileInfo.Length
        $source.CopyTo($Context.Response.OutputStream, 65536)
    } finally {
        $source.Dispose()
        $Context.Response.OutputStream.Close()
    }
}

function Get-Mime([string]$Path) {
    switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.html' { return 'text/html; charset=utf-8' }
        '.css' { return 'text/css; charset=utf-8' }
        '.js' { return 'application/javascript; charset=utf-8' }
        '.mjs' { return 'application/javascript; charset=utf-8' }
        '.json' { return 'application/json; charset=utf-8' }
        '.pdf' { return 'application/pdf' }
        '.png' { return 'image/png' }
        '.svg' { return 'image/svg+xml' }
        default { return 'application/octet-stream' }
    }
}

function Get-RequestCookieValue($Request, [string]$Name) {
    try {
        $cookieHeader = [string]$Request.Headers['Cookie']
        if ([string]::IsNullOrWhiteSpace($cookieHeader)) { return '' }
        foreach ($part in ($cookieHeader -split ';')) {
            $item = $part.Trim()
            $eq = $item.IndexOf('=')
            if ($eq -le 0) { continue }
            $n = $item.Substring(0, $eq).Trim()
            if ($n -ne $Name) { continue }
            return [Uri]::UnescapeDataString($item.Substring($eq + 1))
        }
    } catch {}
    return ''
}

function Test-FixedTimeTokenEquals([string]$Candidate, [string]$Expected) {
    if ($null -eq $Candidate) { $Candidate = '' }
    if ($null -eq $Expected) { $Expected = '' }
    $utf8 = [Text.Encoding]::UTF8
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $candidateHash = $sha.ComputeHash($utf8.GetBytes($Candidate))
        $expectedHash = $sha.ComputeHash($utf8.GetBytes($Expected))
    } finally {
        $sha.Dispose()
    }
    $difference = 0
    for ($i = 0; $i -lt $candidateHash.Length; $i++) {
        $difference = $difference -bor ($candidateHash[$i] -bxor $expectedHash[$i])
    }
    return ($difference -eq 0 -and $Candidate.Length -eq $Expected.Length)
}

# ループバックのcookieはポートで分離されない。127.0.0.1の別ポートで動く任意のページが
# ReportBinderTokenを読み、同一サイト扱いのままAPIを実行できてしまう。
# Host/Originを検証して、自分のオリジン以外からのAPI呼び出しを拒否する。
function Test-RequestOrigin($Request) {
    $allowed = @("127.0.0.1:$($Script:Port)", "localhost:$($Script:Port)")
    $hostHeader = [string]$Request.Headers['Host']
    if (-not [string]::IsNullOrWhiteSpace($hostHeader) -and ($allowed -notcontains $hostHeader)) { return $false }
    $origin = [string]$Request.Headers['Origin']
    if ([string]::IsNullOrWhiteSpace($origin)) { return $true }
    foreach ($a in $allowed) { if ($origin -eq "http://$a") { return $true } }
    return $false
}

function Test-Token($Request) {
    $q = [string]$Request.QueryString['token']
    $t = [string]$Request.QueryString['t']
    $h = [string]$Request.Headers['X-ReportBinder-Token']
    $c = Get-RequestCookieValue $Request 'ReportBinderToken'
    return ((Test-FixedTimeTokenEquals $q $Script:Token) -or
        (Test-FixedTimeTokenEquals $t $Script:Token) -or
        (Test-FixedTimeTokenEquals $h $Script:Token) -or
        (Test-FixedTimeTokenEquals $c $Script:Token))
}

function Serve-Static($Context, [string]$Path) {
    if ($Path -eq '/') { $Path = '/index.html' }
    $rel = $Path.TrimStart('/') -replace '/', [IO.Path]::DirectorySeparatorChar
    if ($rel -match '(^|[\\/])\.\.($|[\\/])') { Write-JsonResponse $Context 400 ([ordered]@{ ok = $false; error = '.. is not allowed' }); return }
    $file = [IO.Path]::GetFullPath((Join-Path $Script:WebRoot $rel))
    $root = [IO.Path]::GetFullPath($Script:WebRoot)
    if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
    if (-not $file.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $file)) {
        Write-JsonResponse $Context 404 ([ordered]@{ ok = $false; error = 'not found' }); return
    }
    Write-BytesResponse $Context 200 ([IO.File]::ReadAllBytes($file)) (Get-Mime $file)
}


function Normalize-WorkspaceRelativePath([string]$RelativePath) {
    if ([string]::IsNullOrWhiteSpace($RelativePath)) { return '' }
    $rel = $RelativePath.Trim()
    if ($rel -eq 'undefined' -or $rel -eq 'null') { return '' }
    if ([IO.Path]::IsPathRooted($rel)) { throw 'PDFパスが不正です。' }
    if ($rel -match '(^|[\\/])\.\.($|[\\/])') { throw 'PDFパスが不正です。' }
    if ($rel -match '[\x00-\x1F]') { throw 'PDFパスが不正です。' }
    if ([IO.Path]::GetExtension($rel).ToLowerInvariant() -ne '.pdf') { throw 'PDFファイルだけ表示できます。' }
    return $rel
}


function Normalize-RelativeForCompare([string]$RelativePath) {
    return (([string]$RelativePath) -replace '\\','/').Trim().ToLowerInvariant()
}

function Serve-ContentPdfByValues($Context, [string]$Language, [string]$PageId, [string]$WorkbookId, [string]$SheetName, [string]$ContentPdf) {
    $workspace = Get-WorkspacePath $Language
    $structure = Get-Structure $Language
    $pageId = ([string]$PageId).Trim()
    if ($pageId -eq 'undefined' -or $pageId -eq 'null') { $pageId = '' }
    $workbookId = ([string]$WorkbookId).Trim()
    if ($workbookId -eq 'undefined' -or $workbookId -eq 'null') { $workbookId = '' }
    $sheetName = ([string]$SheetName).Trim()
    if ($sheetName -eq 'undefined' -or $sheetName -eq 'null') { $sheetName = '' }
    $contentRel = Normalize-WorkspaceRelativePath $ContentPdf
    $page = @()

    if (-not [string]::IsNullOrWhiteSpace($pageId)) {
        $page = @(Get-Array $structure.pages | Where-Object {
            (Resolve-PageId $_) -eq $pageId -or [string]$_.pageId -eq $pageId -or [string]$_.id -eq $pageId
        } | Select-Object -First 1)
    }
    if ($page.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($workbookId) -and -not [string]::IsNullOrWhiteSpace($sheetName)) {
        $page = @(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -eq $workbookId -and [string]$_.sheetName -eq $sheetName } | Select-Object -First 1)
    }
    if ($page.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($contentRel)) {
        $cmp = Normalize-RelativeForCompare $contentRel
        $page = @(Get-Array $structure.pages | Where-Object { (Normalize-RelativeForCompare ([string]$_.contentPdf)) -eq $cmp } | Select-Object -First 1)
    }

    $rel = ''
    if ($page.Count -gt 0) { $rel = [string]$page[0].contentPdf }
    if ([string]::IsNullOrWhiteSpace($rel) -and -not [string]::IsNullOrWhiteSpace($contentRel)) { $rel = $contentRel }
    $rel = Normalize-WorkspaceRelativePath $rel
    if ([string]::IsNullOrWhiteSpace($rel)) { throw 'PDF情報が不足しています。ページ構成を更新してからPDFを開いてください。' }

    $full = [IO.Path]::GetFullPath((Join-Path $workspace $rel))
    $workspaceFull = [IO.Path]::GetFullPath($workspace)
    if (-not $workspaceFull.EndsWith([IO.Path]::DirectorySeparatorChar)) { $workspaceFull += [IO.Path]::DirectorySeparatorChar }
    if (-not $full.StartsWith($workspaceFull, [StringComparison]::OrdinalIgnoreCase)) { throw 'ワークスペース外のファイルは表示できません。' }
    if (-not (Test-Path -LiteralPath $full)) { throw 'PDFファイルが見つかりません。PDF作成をやり直してください。' }
    if ((Get-Item -LiteralPath $full).Length -le 0) { throw 'PDFファイルが空です。PDF作成をやり直してください。' }
    Write-FileResponse $Context 200 $full 'application/pdf' $false 'no-store'
}

function Resolve-ContentPdfFullPath([string]$Workspace, [string]$RelativePdf) {
    if ([string]::IsNullOrWhiteSpace($RelativePdf)) { throw 'content-pdf が未作成です。' }
    Test-RelativePath $RelativePdf | Out-Null
    if ([IO.Path]::GetExtension($RelativePdf).ToLowerInvariant() -ne '.pdf') { throw 'PDFファイルだけ表示できます。' }
    $full = [IO.Path]::GetFullPath((Join-Path $Workspace $RelativePdf))
    $workspaceFull = [IO.Path]::GetFullPath($Workspace)
    if (-not $workspaceFull.EndsWith([IO.Path]::DirectorySeparatorChar)) { $workspaceFull += [IO.Path]::DirectorySeparatorChar }
    if (-not $full.StartsWith($workspaceFull, [StringComparison]::OrdinalIgnoreCase)) { throw 'ワークスペース外のファイルは表示できません。' }
    if (-not (Test-Path -LiteralPath $full)) { throw 'PDFファイルが見つかりません。PDF作成をやり直してください。' }
    return $full
}

function Serve-ContentPdf($Context, [string]$Language) {
    $pageId = [string]$Context.Request.QueryString['pageId']
    if ([string]::IsNullOrWhiteSpace($pageId)) { $pageId = [string]$Context.Request.QueryString['id'] }
    $contentPdf = [string]$Context.Request.QueryString['contentPdf']
    if ([string]::IsNullOrWhiteSpace($contentPdf)) { $contentPdf = [string]$Context.Request.QueryString['pdf'] }
    if ([string]::IsNullOrWhiteSpace($contentPdf)) { $contentPdf = [string]$Context.Request.QueryString['path'] }
    Serve-ContentPdfByValues $Context $Language $pageId ([string]$Context.Request.QueryString['workbookId']) ([string]$Context.Request.QueryString['sheetName']) $contentPdf
}

function Serve-ContentPdfFromBody($Context, [string]$Language, $Body) {
    $pageId = [string]$Body.pageId
    if ([string]::IsNullOrWhiteSpace($pageId)) { $pageId = [string]$Body.id }
    $contentPdf = [string]$Body.contentPdf
    if ([string]::IsNullOrWhiteSpace($contentPdf)) { $contentPdf = [string]$Body.pdf }
    if ([string]::IsNullOrWhiteSpace($contentPdf)) { $contentPdf = [string]$Body.path }
    Serve-ContentPdfByValues $Context $Language $pageId ([string]$Body.workbookId) ([string]$Body.sheetName) $contentPdf
}

function Serve-FinalPdfByVolume($Context,[string]$Language,[string]$Volume,[string]$Category) {
    if ((Get-VolumeList $Language | Where-Object { $_ -ne 'none' }) -notcontains $Volume) { throw [System.ArgumentException]::new('volumeには本体または補足を指定してください。') }
    $cat=Require-WorkbookCategory $Category;$volume=([string]$Volume).Trim();if(@(Get-VolumeList $Language|Where-Object{$_ -ne 'none'}) -notcontains $volume){throw [ArgumentException]::new('不正な成果物です。')};$structure=Get-Structure $Language;$v=Get-DataProperty $structure.volumes (Get-VolumeStateKey $volume $cat) $null;if($null -eq $v -or -not [string]$v.outputPdf){throw '提出用PDFはまだ作成されていません。'};$paths=Get-Paths;$full=[IO.Path]::GetFullPath([string]$v.outputPdf);$root=[IO.Path]::GetFullPath([string]$paths.outputDir);if(-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)){$root+=[IO.Path]::DirectorySeparatorChar};if(-not $full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw '出力フォルダ外のPDFは表示できません。'};if(-not(Test-Path $full)){throw '提出用PDFファイルが見つかりません。'};Write-BytesResponse $Context 200 ([IO.File]::ReadAllBytes($full)) 'application/pdf'
}

function Serve-FinalPdf($Context,[string]$Language) {
    Serve-FinalPdfByVolume $Context $Language ([string]$Context.Request.QueryString['volume']) ([string]$Context.Request.QueryString['category'])
}

function Serve-DocumentPackPdf($Context, [string]$Language, [string]$PackId, [string]$TargetId) {
    $structure = Get-Structure $Language
    $scope = Resolve-DocumentPackScope $structure $PackId $true
    $TargetId = Assert-PackTargetId $Language $scope.pack $TargetId
    $volume = Get-LegacyVolumeFromTargetId $Language $TargetId
    $state = Get-PackOutputState $structure $Language ([string]$scope.packId) $volume $false
    $output = [string](Get-DataProperty $state 'outputPdf' '')
    if ([string]::IsNullOrWhiteSpace($output)) { throw '提出用PDFはまだ作成されていません。' }
    $paths = Get-Paths
    $full = [IO.Path]::GetFullPath($output)
    $root = [IO.Path]::GetFullPath([string]$paths.outputDir)
    if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
    if (-not $full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)) { throw '出力フォルダ外のPDFは表示できません。' }
    if (-not (Test-Path -LiteralPath $full)) { throw '提出用PDFファイルが見つかりません。' }
    Write-FileResponse $Context 200 $full 'application/pdf'
}

# 出力したPDFの保存先をエクスプローラーで開き、そのファイルを選択状態にする。
# ブラウザの別タブでは中身を見られるだけで、添付やコピーのために実体を掴めないため。
# 出力フォルダーの外は開かない(パスは structure 由来だが、確認は Serve と同じ形で行う)。
function Resolve-OutputPdfForReveal([string]$Language, [string]$PackOrCategory, [string]$TargetId, [string]$Volume) {
    $structure = Get-Structure $Language
    # Resolve-DocumentPackScope は packId と category のどちらでも受ける。組み込みパックは
    # 画面が category と volume を送るため、targetId が無いときは volume をそのまま使う。
    $scope = Resolve-DocumentPackScope $structure $PackOrCategory $true
    if (-not [string]::IsNullOrWhiteSpace($TargetId)) {
        $resolvedTarget = Assert-PackTargetId $Language $scope.pack $TargetId
        $volume = Get-LegacyVolumeFromTargetId $Language $resolvedTarget
    } else {
        $volume = ([string]$Volume).Trim()
        if (@(Get-VolumeList $Language | Where-Object { $_ -ne 'none' }) -notcontains $volume) { throw [ArgumentException]::new('不正な出力先です。') }
    }
    $state = Get-PackOutputState $structure $Language ([string]$scope.packId) $volume $false
    $output = [string](Get-DataProperty $state 'outputPdf' '')
    if ([string]::IsNullOrWhiteSpace($output)) { throw '提出用PDFはまだ作成されていません。' }
    $paths = Get-Paths
    $full = [IO.Path]::GetFullPath($output)
    $root = [IO.Path]::GetFullPath([string]$paths.outputDir)
    if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
    if (-not $full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)) { throw '出力フォルダーの外にあるPDFは開けません。' }
    if (-not (Test-Path -LiteralPath $full)) { throw '提出用PDFファイルが見つかりません。出力フォルダーから移動または削除された可能性があります。' }
    return $full
}

function Open-OutputPdfLocation([string]$FullPath) {
    # 実測(2026-08-10)で確かめた形。どちらを崩してもエクスプローラーは何も開かない。
    #   - `/select` の直後のカンマは必須。空白にすると窓が出ない。
    #   - パスの引用符も必須。出力ファイル名は利用者が編集できるパターン
    #     ({packName}_{targetName}_{yyyyMMdd}.pdf)から作られるためカンマを含みうる。
    #     引用しないと、カンマを含む名前で窓が出ない。
    $argument = '/select,"' + $FullPath + '"'
    [void][Diagnostics.Process]::Start((New-Object Diagnostics.ProcessStartInfo -Property @{
        FileName = 'explorer.exe'
        Arguments = $argument
        UseShellExecute = $true
    }))
}

function Get-SafePublishUserName {
    $name = ([string]$env:USERNAME).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'user' }
    $name = [regex]::Replace($name, '[<>:"/\\|?*\x00-\x1F]+', '_')
    $name = [regex]::Replace($name, '\s+', '_').Trim([char[]]@('_','.'))
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'user' }
    if ($name.Length -gt 40) { $name = $name.Substring(0,40) }
    return $name
}

function Get-SafePublishPackName([string]$DisplayName) {
    $name = ([string]$DisplayName).Trim()
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'document-pack' }
    $name = [regex]::Replace($name, '[<>:"/\\|?*\x00-\x1F]+', '_')
    $name = [regex]::Replace($name, '\s+', '_').Trim([char[]]@('_','.'))
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'document-pack' }
    if ($name.Length -gt 48) { $name = $name.Substring(0,48) }
    return $name
}

function Publish-DocumentPackPdfToShared([string]$Language, [string]$Volume, [string]$PackIdOrCategory) {
    $paths = Get-Paths
    $structure = Get-Structure $Language
    $scope = Resolve-DocumentPackScope $structure $PackIdOrCategory $false
    if (@(Get-PackVolumeList $Language $scope.pack $false) -notcontains $Volume) { throw [ArgumentException]::new('共有発行するPDFの種類が不正です。') }
    $packId = [string]$scope.packId
    $cat = [string]$scope.category
    $ready = Get-FinalBuildReadiness $structure $Language $Volume $packId
    if ([string]$ready.displayState -ne 'built') {
        throw '提出用PDFが最新ではありません。最新状態で再出力してから共有発行してください。'
    }
    $source = [string]$ready.outputPdf
    if ([string]::IsNullOrWhiteSpace($source) -or -not (Test-Path -LiteralPath $source -PathType Leaf)) {
        throw '共有発行できる提出用PDFがありません。先に提出用PDFを出力してください。'
    }

    $localOutputRoot = [IO.Path]::GetFullPath([string]$paths.outputDir)
    if (-not $localOutputRoot.EndsWith([IO.Path]::DirectorySeparatorChar)) { $localOutputRoot += [IO.Path]::DirectorySeparatorChar }
    $sourceFull = [IO.Path]::GetFullPath($source)
    if (-not $sourceFull.StartsWith($localOutputRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw '共有発行元が利用者ローカルの出力フォルダー外です。提出用PDFを再出力してください。'
    }

    $submissionDir = [IO.Path]::GetFullPath([string]$paths.submissionDir)
    $userName = Get-SafePublishUserName
    $packName = Get-SafePublishPackName ([string](Get-DataProperty $scope.pack 'displayName' $packId))
    $stamp = Get-Date -Format 'MMdd_HHmmss'
    $baseFolderName = "{0}_{1}_{2}" -f $stamp, $packName, $userName
    $folderName = $baseFolderName
    $publishDir = Join-Path $submissionDir $folderName
    $suffix = 2
    while (Test-Path -LiteralPath $publishDir) {
        $folderName = "{0}_{1}" -f $baseFolderName, $suffix
        $publishDir = Join-Path $submissionDir $folderName
        $suffix++
    }

    # PDFが半端な状態で共有側に見えないよう、同じ共有ルートの隠し一時フォルダーで
    # コピーと検証を完了してから、発行フォルダー名へ一度だけ切り替える。
    $stagingDir = Join-Path $submissionDir ('.publishing-' + [Guid]::NewGuid().ToString('N'))
    $fileName = [IO.Path]::GetFileName($sourceFull)
    $stagedPdf = Join-Path $stagingDir $fileName
    try {
        New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null
        Copy-Item -LiteralPath $sourceFull -Destination $stagedPdf -Force
        $sourceLength = (Get-Item -LiteralPath $sourceFull -ErrorAction Stop).Length
        $uploadedLength = (Get-Item -LiteralPath $stagedPdf -ErrorAction Stop).Length
        if ($sourceLength -le 0 -or $uploadedLength -ne $sourceLength) {
            throw '共有フォルダーへのコピーサイズが一致しません。'
        }
        while (Test-Path -LiteralPath $publishDir) {
            $folderName = "{0}_{1}" -f $baseFolderName, $suffix
            $publishDir = Join-Path $submissionDir $folderName
            $suffix++
        }
        Move-Item -LiteralPath $stagingDir -Destination $publishDir
        $stagingDir = ''
    } finally {
        if ($stagingDir -and (Test-Path -LiteralPath $stagingDir)) {
            Remove-Item -LiteralPath $stagingDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    $destination = Join-Path $publishDir $fileName

    $publishedAt = New-NowIso
    Write-HistoryEvent $Language 'final.published' ([ordered]@{
        packId = $packId
        category = $cat
        targetId = Get-TargetIdFromLegacyVolume $Volume
        volume = $Volume
        source = $sourceFull
        destination = $destination
        sharedFolder = $publishDir
        folderName = $folderName
        fileName = $fileName
        userName = $userName
        packName = $packName
        publishedAt = $publishedAt
    })
    return [ordered]@{
        volume = $Volume
        targetId = Get-TargetIdFromLegacyVolume $Volume
        packId = $packId
        category = $cat
        sourcePdf = $sourceFull
        sharedPath = $destination
        sharedFolder = $publishDir
        folderName = $folderName
        fileName = $fileName
        publishedBy = $userName
        packName = $packName
        publishedAt = $publishedAt
    }
}

function Publish-FinalPdfToShared([string]$Language, [string]$Volume, [string]$Category) {
    $cat = Require-WorkbookCategory $Category
    return Publish-DocumentPackPdfToShared $Language $Volume (Get-BuiltinPackId $cat)
}

function Handle-Api($Context) {
    $language = Get-EffectiveLanguage
    $path = $Context.Request.Url.AbsolutePath
    $method = $Context.Request.HttpMethod.ToUpperInvariant()
    try {
        # These lightweight endpoints are only for startup readiness checks.
        # They intentionally do not require a session token so the local wait page
        # can detect readiness even if the browser strips or delays query handling.
        if ($method -eq 'OPTIONS') {
            Write-TextResponse $Context 204 '' 'text/plain; charset=utf-8'; return
        }
        if ($method -eq 'GET' -and $path -eq '/api/ready.gif') {
            Write-BytesResponse $Context 200 $Script:ReadyGifBytes 'image/gif' $true; return
        }
        if ($method -eq 'GET' -and $path -eq '/api/ping') {
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; runtimeVersion = $Script:RuntimeVersion; at = New-NowIso }) $true; return
        }
        if (-not (Test-RequestOrigin $Context.Request)) { Write-JsonResponse $Context 403 ([ordered]@{ ok = $false; error = 'invalid origin' }); return }
        if (-not (Test-Token $Context.Request)) { Write-JsonResponse $Context 403 ([ordered]@{ ok = $false; error = 'invalid token' }); return }
        Touch-ClientActivity '' | Out-Null
        if ($method -eq 'POST' -and $path -eq '/api/heartbeat') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 (Touch-ClientActivity ([string]$body.clientId)); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/client/close') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 (Notify-ClientClosing ([string]$body.clientId)); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/shutdown') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 (Request-ServerShutdown ([string]$body.reason) $language); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/auto/render') {
            $body = Read-BodyJson $Context.Request
            $ids = @()
            if ($body.workbookIds) { $ids = @(Get-Array $body.workbookIds | ForEach-Object { [string]$_ }) }
            $result = Invoke-AutoRender $language $ids
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/state') {
            Write-JsonResponse $Context 200 (Get-V2StatePayload $language); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/diagnostics/run') {
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; diagnostics=(Get-SystemDiagnostics $language) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/pack-templates') {
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; packTemplates = @(Get-PackTemplateCatalog $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/pack-templates') {
            $body = Read-BodyJson $Context.Request
            $result = Save-PackTemplate $language $body
            Write-JsonResponse $Context 201 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; packTemplates = @(Get-PackTemplateCatalog $language) }); return
        }
        if ($method -eq 'PATCH' -and $path -match '^/api/v2/pack-templates/([^/]+)$') {
            $body = Read-BodyJson $Context.Request
            $templateId = [Uri]::UnescapeDataString([string]$matches[1])
            $result = Save-PackTemplate $language $body $templateId
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; packTemplates = @(Get-PackTemplateCatalog $language) }); return
        }
        if ($method -eq 'DELETE' -and $path -match '^/api/v2/pack-templates/([^/]+)$') {
            $templateId = [Uri]::UnescapeDataString([string]$matches[1])
            $result = Remove-PackTemplate $language $templateId
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; packTemplates = @(Get-PackTemplateCatalog $language) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/packs') {
            $packPaths = Get-Paths
            $packDataDir = [string](Get-DataProperty $packPaths 'dataDir' '')
            $structure = if (-not [string]::IsNullOrWhiteSpace($packDataDir) -and (Test-Path -LiteralPath $packDataDir)) { Get-Structure $language } else { New-EmptyStructure $language }
            $includeArchived = ([string]$Context.Request.QueryString['includeArchived']).Trim().ToLowerInvariant() -in @('1','true','yes')
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; packs = @(Get-PublicPackList $structure $language $includeArchived) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/packs') {
            $body = Read-BodyJson $Context.Request
            $result = New-DocumentPack $language $body
            Write-JsonResponse $Context 201 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; packs = @(Get-PublicPackList (Get-Structure $language) $language) }); return
        }
        if ($method -eq 'POST' -and $path -match '^/api/v2/packs/([^/]+)/duplicate$') {
            $body = Read-BodyJson $Context.Request
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            $result = Copy-DocumentPack $language $packId $body
            Write-JsonResponse $Context 201 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; packs = @(Get-PublicPackList (Get-Structure $language) $language) }); return
        }
        if ($method -eq 'GET' -and $path -match '^/api/v2/packs/([^/]+)/template-upgrade$') {
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; preview = (Get-PackTemplateUpgradePreview $language $packId) }); return
        }
        if ($method -eq 'POST' -and $path -match '^/api/v2/packs/([^/]+)/template-upgrade$') {
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            $result = Update-PackTemplateSnapshot $language $packId
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; packs = @(Get-PublicPackList (Get-Structure $language) $language) }); return
        }
        if ($method -eq 'POST' -and $path -match '^/api/v2/packs/([^/]+)/archive$') {
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            $result = Set-DocumentPackArchived $language $packId $true
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; packs = @(Get-PublicPackList (Get-Structure $language) $language) }); return
        }
        if ($method -eq 'POST' -and $path -match '^/api/v2/packs/([^/]+)/restore$') {
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            $result = Set-DocumentPackArchived $language $packId $false
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; packs = @(Get-PublicPackList (Get-Structure $language) $language) }); return
        }
        if ($method -eq 'DELETE' -and $path -match '^/api/v2/packs/([^/]+)$') {
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            $result = Remove-DocumentPack $language $packId
            Write-HistoryEvent $language 'pack.deleted' ([ordered]@{ packId=$packId; displayName=[string]$result.displayName; removedWorkbookCount=[int]$result.removedWorkbookCount; removedPageCount=[int]$result.removedPageCount })
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; result=$result; state=(Get-StatePayload $language); packs=@(Get-PublicPackList (Get-Structure $language) $language $true) }); return
        }
        if ($method -eq 'GET' -and $path -match '^/api/v2/packs/([^/]+)/review$') {
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; review=(Get-PackReviewSnapshot (Get-Structure $language) $language $packId $true) }); return
        }
        if ($method -eq 'POST' -and $path -match '^/api/v2/packs/([^/]+)/review$') {
            $body = Read-BodyJson $Context.Request
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; review=(Invoke-PackReviewAction $language $packId $body); state=(Get-StatePayload $language) }); return
        }
        if ($method -eq 'PATCH' -and $path -match '^/api/v2/packs/([^/]+)$') {
            $body = Read-BodyJson $Context.Request
            $packId = [Uri]::UnescapeDataString([string]$matches[1])
            $result = Update-PackSettings $language $packId $body
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; state = (Get-StatePayload $language) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/source-candidates') {
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; scannedAt = (New-NowIso); candidates = @(Get-SourceCandidates @('excel','word','pdf','powerpoint')) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/sources/register-batch') {
            $body = Read-BodyJson $Context.Request
            $relativePaths = @(Get-Array (Get-DataProperty $body 'relativePaths' @()) | ForEach-Object { [string]$_ })
            if ($relativePaths.Count -eq 0) { throw '登録する原稿が選択されていません。' }
            $packId = [string](Get-DataProperty $body 'packId' '')
            if ([string]::IsNullOrWhiteSpace($packId)) { $packId = Get-BuiltinPackId (Require-WorkbookCategory ([string](Get-DataProperty $body 'category' ''))) }
            $result = Register-SourcesBatch $language $relativePaths $packId ([string](Get-DataProperty $body 'sourceType' ''))
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; state = (Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/sources/unregister') {
            $body = Read-BodyJson $Context.Request
            $sourceId = [string](Get-DataProperty $body 'sourceId' '')
            if ([string]::IsNullOrWhiteSpace($sourceId)) { throw 'sourceId が必要です。' }
            $result = Unregister-Workbook $language $sourceId
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; state = (Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/sources/scan-updates') {
            $result = Scan-Updates $language
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; result = $result; state = (Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/sources/render/start') {
            $body = Read-BodyJson $Context.Request
            $sourceIds = @(Get-Array (Get-DataProperty $body 'sourceIds' @()) | ForEach-Object { [string]$_ })
            $renderPackId = [string](Get-DataProperty $body 'packId' '')
            if ($sourceIds.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($renderPackId)) {
                $renderStructure = Get-Structure $language
                $renderScope = Resolve-DocumentPackScope $renderStructure $renderPackId $false
                $sourceIds = @(Get-Array $renderStructure.workbooks | Where-Object { Test-WorkbookPack $_ ([string]$renderScope.packId) } | ForEach-Object { [string]$_.workbookId })
            }
            $result = Start-RenderJob $language $sourceIds ([bool](Get-DataProperty $body 'onlyUpdated' $false)) ([string](Get-DataProperty $body 'category' '')) $null ([string](Get-DataProperty $body 'excelSheetSelectionMode' ''))
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; apiVersion = 2; jobId = [string]$result.jobId; job = $result }); return
        }
        if ($method -eq 'PATCH' -and $path -match '^/api/v2/sources/([^/]+)$') {
            $body = Read-BodyJson $Context.Request
            $sourceId = [Uri]::UnescapeDataString([string]$matches[1])
            $updated = Update-SourceMetadata $language $sourceId $body
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; source = $updated; state = (Get-StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/items/reorder') {
            $body = Read-BodyJson $Context.Request
            $volumes = Get-DataProperty $body 'volumes' $null
            if ($null -eq $volumes) {
                $targets = Get-DataProperty $body 'targets' $null
                if ($null -eq $targets) { throw 'targets が必要です。' }
                $volumeMap = [ordered]@{}
                foreach ($property in $targets.PSObject.Properties) {
                    $targetId = ([string]$property.Name).Trim().ToLowerInvariant()
                    $volume = if ($targetId -in @('none','unassigned')) { 'none' } else { Get-LegacyVolumeFromTargetId $language $targetId }
                    $volumeMap[$volume] = @(Get-Array $property.Value)
                }
                $volumes = [pscustomobject]$volumeMap
            }
            $request = [pscustomobject][ordered]@{ packId=[string](Get-DataProperty $body 'packId' (Get-DataProperty $body 'category' '')); volumes=$volumes }
            # Keep the optimistic-lock fingerprint supplied by the browser.
            # Rebuilding the request without it silently disabled conflict
            # detection on the V2 route and allowed a stale tab to overwrite a
            # newer page composition.
            if (Test-ConfigHasKey $body 'baseLayout') { Set-NoteProperty $request 'baseLayout' ([string](Get-DataProperty $body 'baseLayout' '')) }
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; result=(Reorder-Pages $language $request); state=(Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'PATCH' -and $path -match '^/api/v2/items/([^/]+)$') {
            $body = Read-BodyJson $Context.Request
            $itemId = [Uri]::UnescapeDataString([string]$matches[1])
            $request = [pscustomobject][ordered]@{ packId=[string](Get-DataProperty $body 'packId' (Get-DataProperty $body 'category' '')); pageId=$itemId }
            if (Test-ConfigHasKey $body 'baseLayout') { Set-NoteProperty $request 'baseLayout' ([string](Get-DataProperty $body 'baseLayout' '')) }
            foreach ($name in @('title','numberingMode','numberingManual','resetNumbering','pageRange','clearPageRange','enabled')) { if (Test-ConfigHasKey $body $name) { Set-NoteProperty $request $name (Get-DataProperty $body $name $null) } }
            if (Test-ConfigHasKey $body 'targetId') { Set-NoteProperty $request 'volume' (Get-LegacyVolumeFromTargetId $language ([string](Get-DataProperty $body 'targetId' ''))) }
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; result=(Update-Page $language $request); state=(Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/layout/snapshots') {
            $packId = [string]$Context.Request.QueryString['packId']
            $structure = Get-Structure $language
            $scope = Get-LayoutScopeInfo $structure $packId $true
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; packId=[string]$scope.packId; snapshots=(Get-LayoutSnapshots $language ([string]$scope.packId)) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/layout/restore/preview') {
            $body = Read-BodyJson $Context.Request
            $packId = [string](Get-DataProperty $body 'packId' '')
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; preview=(Get-LayoutRestorePreview $language $packId ([string](Get-DataProperty $body 'snapshotId' ''))) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/layout/restore') {
            $body = Read-BodyJson $Context.Request
            $packId = [string](Get-DataProperty $body 'packId' '')
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; result=(Restore-LayoutSnapshot $language $packId ([string](Get-DataProperty $body 'snapshotId' ''))); state=(Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/outputs/readiness') {
            $packId = [string]$Context.Request.QueryString['packId']
            $structure = Get-Structure $language
            $scope = Resolve-DocumentPackScope $structure $packId $false
            $targets = [ordered]@{}
            foreach ($targetId in @(Get-PackTargetIds $language $scope.pack)) {
                $volume = Get-LegacyVolumeFromTargetId $language $targetId
                $targets[$targetId] = Get-FinalBuildReadiness $structure $language $volume ([string]$scope.packId)
            }
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; packId=[string]$scope.packId; targets=$targets }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/outputs/build') {
            $body = Read-BodyJson $Context.Request
            $packId = [string](Get-DataProperty $body 'packId' '')
            $targetIds = @(Get-Array (Get-DataProperty $body 'targetIds' @()))
            $structure = Get-Structure $language
            $scope = Resolve-DocumentPackScope $structure $packId $false
            $allowedTargetIds = @(Get-PackTargetIds $language $scope.pack)
            if ($targetIds.Count -eq 0) { $targetIds = @([string](Get-DataProperty $body 'targetId' $allowedTargetIds[0])) }
            # 出力は数十秒かかることがある。サーバーはリクエストを直列に処理するため、
            # ここで完結させると進捗ポーリングも中止も受け付けられない。ジョブとして返す。
            $job = Start-FinalBuildJob $language ([string]$scope.packId) @($targetIds | ForEach-Object { [string]$_ })
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; job=$job }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/outputs/build/status') {
            $jobId = [string]$Context.Request.QueryString['jobId']
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; job=(Read-FinalJobStatus $language $jobId); state=(Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/outputs/build/cancel') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; result=(Request-FinalJobCancellation $language ([string](Get-DataProperty $body 'jobId' ''))) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/outputs/file') {
            $body = Read-BodyJson $Context.Request
            Serve-DocumentPackPdf $Context $language ([string](Get-DataProperty $body 'packId' '')) ([string](Get-DataProperty $body 'targetId' 'main')); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/sources/relink-candidates') {
            $relinkFor = [string]$Context.Request.QueryString['sourceId']
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; result=(Get-RelinkCandidates $language $relinkFor) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/sources/relink') {
            $body = Read-BodyJson $Context.Request
            $relinkResult = Relink-Source $language ([string](Get-DataProperty $body 'sourceId' '')) ([string](Get-DataProperty $body 'relativePath' ''))
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; result=$relinkResult }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/outputs/reveal') {
            $body = Read-BodyJson $Context.Request
            $revealPackId = [string](Get-DataProperty $body 'packId' '')
            $revealTargetId = [string](Get-DataProperty $body 'targetId' '')
            if ([string]::IsNullOrWhiteSpace($revealPackId)) { $revealPackId = [string](Get-DataProperty $body 'category' '') }
            $revealPath = Resolve-OutputPdfForReveal $language $revealPackId $revealTargetId ([string](Get-DataProperty $body 'volume' ''))
            Open-OutputPdfLocation $revealPath
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; path=$revealPath }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/v2/outputs/publish') {
            $body = Read-BodyJson $Context.Request
            $packId = [string](Get-DataProperty $body 'packId' '')
            $targetId = ([string](Get-DataProperty $body 'targetId' '')).Trim().ToLowerInvariant()
            $publishStructure = Get-Structure $language
            $publishScope = Resolve-DocumentPackScope $publishStructure $packId $false
            $targetId = Assert-PackTargetId $language $publishScope.pack $targetId
            $volume = Get-LegacyVolumeFromTargetId $language $targetId
            $result = Publish-DocumentPackPdfToShared $language $volume $packId
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; result=$result }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/v2/outputs/archives') {
            $packId = [string]$Context.Request.QueryString['packId']
            $structure = Get-Structure $language
            $scope = Resolve-DocumentPackScope $structure $packId $true
            Write-JsonResponse $Context 200 ([ordered]@{ ok=$true; apiVersion=2; packId=[string]$scope.packId; archives=(Get-FinalArchives $language ([string]$scope.packId)) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/state') {
            Write-JsonResponse $Context 200 (Get-StatePayload $language); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/final/readiness') {
            $cat=Require-WorkbookCategory ([string]$Context.Request.QueryString['category'])
            $structure=Get-Structure $language;$vols=[ordered]@{}
            foreach($volume in @(Get-VolumeList $language|Where-Object{$_ -ne 'none'})){$vols[$volume]=Get-FinalBuildReadiness $structure $language $volume $cat}
            Write-JsonResponse $Context 200 ([ordered]@{ok=$true;category=$cat;volumes=$vols});return
        }
        if ($method -eq 'GET' -and $path -eq '/api/submission-files') {
            # 空の配列を式のまま渡すと PowerShell が展開して $null になり、
            # ConvertTo-Json が "files":{} を返す。受け取った画面は asArray() で
            # 要素1件と数え、原稿が0件のときだけ出るはずの案内
            # (「サブフォルダーの中は探しません」)へ決して到達しなかった。
            $sourceFiles = Get-SourceCandidates @('excel','word','pdf','powerpoint')
            if ($null -eq $sourceFiles) { $sourceFiles = @() }
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; scannedAt = (New-NowIso); files = $sourceFiles }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/submission/select-and-start') {
            $body = Read-BodyJson $Context.Request
            $initial = [string]$body.initialDir
            if ([string]::IsNullOrWhiteSpace($initial)) { $initial = [string](Get-Paths).submissionDir }
            # 画面はこのフォルダーを一貫して「原稿フォルダー」と呼ぶ。押したボタンと
            # 開いたダイアログで呼び名が変わると、選ぶ対象を取り違える。
            $selected = Select-FolderDialog '原稿フォルダーを選ぶ' $initial
            if ([string]::IsNullOrWhiteSpace($selected)) {
                Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; cancelled = $true; state = (Get-StatePayload $language) }); return
            }
            $newPaths = Get-DefaultChildPaths $selected
            $migration = Initialize-CleanLocalProject $newPaths
            Ensure-Package $newPaths -Languages @($language)
            $config = Get-AppConfig
            $config.lastSubmissionDir = [string]$newPaths.submissionDir
            $config.lastDataDir = [string]$newPaths.dataDir
            $config.lastOutputDir = [string]$newPaths.outputDir
            Save-AppConfig $config
            # /api/submission-files と同じ空配列の展開が起きる。こちらが「フォルダーを
            # 選ぶ」を押した人が通る主経路で、原稿0件のフォルダーを選ぶと files が
            # {} になり、画面側は1件と数えて再取得(loadFilesSilently)も飛ばしていた。
            $selectedFiles = Get-SourceCandidates @('excel','word','pdf','powerpoint')
            if ($null -eq $selectedFiles) { $selectedFiles = @() }
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; paths = $newPaths; path = [string]$newPaths.submissionDir; migration = $migration; state = (Get-StatePayload $language); scannedAt = (New-NowIso); files = $selectedFiles }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/dialog/folder') {
            $body = Read-BodyJson $Context.Request
            $selected = Select-FolderDialog ([string]$body.title) ([string]$body.initialDir)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; path = $selected }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/paths/defaults') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; paths = (Get-DefaultChildPaths ([string]$body.submissionDir)) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/file') {
            Serve-ContentPdf $Context $language; return
        }
        if ($method -eq 'POST' -and $path -eq '/api/file') {
            $body = Read-BodyJson $Context.Request
            Serve-ContentPdfFromBody $Context $language $body; return
        }
        if ($method -eq 'GET' -and $path -eq '/api/final/file') {
            Serve-FinalPdf $Context $language; return
        }
        if ($method -eq 'POST' -and $path -eq '/api/final/file') {
            $body = Read-BodyJson $Context.Request
            Serve-FinalPdfByVolume $Context $language ([string]$body.volume) ([string]$body.category); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/paths') {
            $body = Read-BodyJson $Context.Request
            # 管理データと通常出力は常に利用者ローカル。APIから共有側の任意パスを
            # 指定して、複数利用者の状態が再び混ざる経路を残さない。
            $requestedSubmissionDir = ([string]$body.submissionDir).Trim()
            if ([string]::IsNullOrWhiteSpace($requestedSubmissionDir) -or -not (Test-Path -LiteralPath $requestedSubmissionDir -PathType Container)) {
                throw '入力した原稿フォルダが見つかりません。場所を確認するか、「フォルダを選ぶ」を使用してください。'
            }
            $newPaths = Get-DefaultChildPaths $requestedSubmissionDir
            $migration = Initialize-CleanLocalProject $newPaths
            Ensure-Package $newPaths -Languages @($language)
            $config = Get-AppConfig
            $config.lastSubmissionDir = [string]$newPaths.submissionDir
            $config.lastDataDir = [string]$newPaths.dataDir
            $config.lastOutputDir = [string]$newPaths.outputDir
            Save-AppConfig $config
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; paths = $newPaths }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/workbooks/register-batch') {
            $body = Read-BodyJson $Context.Request
            $rels = @()
            if ($body.relativePaths) { $rels = @(Get-Array $body.relativePaths | ForEach-Object { [string]$_ }) }
            elseif ($body.relativePath) { $rels = @([string]$body.relativePath) }
            if ($rels.Count -eq 0) { throw '登録する原稿が選択されていません。' }
            $cat = Require-WorkbookCategory ([string]$body.category)
            $result = Register-SourcesBatch $language $rels (Get-BuiltinPackId $cat) ''
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result; state = (Get-StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/workbooks/register') {
            $body = Read-BodyJson $Context.Request
            $cat = Require-WorkbookCategory ([string]$body.category)
            $result = Register-Source $language ([string]$body.relativePath) (Get-BuiltinPackId $cat) ([string](Get-DataProperty $body 'sourceType' ''))
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result; state = (Get-StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/workbooks/unregister') {
            $body = Read-BodyJson $Context.Request
            $result = Unregister-Workbook $language ([string]$body.workbookId)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result }); return
        }
        if (($method -eq 'GET' -or $method -eq 'POST') -and $path -eq '/api/jobs/status') {
            $jobId = ''
            if ($method -eq 'POST') {
                $body = Read-BodyJson $Context.Request
                $jobId = [string](Get-DataProperty $body 'jobId' '')
                if ([string]::IsNullOrWhiteSpace($jobId)) { $jobId = [string](Get-DataProperty $body 'JobId' '') }
                if ([string]::IsNullOrWhiteSpace($jobId)) { $jobId = [string](Get-DataProperty $body 'id' '') }
            }
            if ([string]::IsNullOrWhiteSpace($jobId)) { $jobId = [string]$Context.Request.QueryString['jobId'] }
            if ([string]::IsNullOrWhiteSpace($jobId)) { $jobId = [string]$Context.Request.QueryString['JobId'] }
            if ([string]::IsNullOrWhiteSpace($jobId)) { $jobId = [string]$Context.Request.QueryString['id'] }
            Write-JsonResponse $Context 200 (Read-RenderJobStatus $language $jobId); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/jobs/cancel') {
            $body = Read-BodyJson $Context.Request
            $jobId = [string](Get-DataProperty $body 'jobId' '')
            Write-JsonResponse $Context 200 (Request-RenderJobCancellation $language $jobId); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/workbooks/render/start') {
            $body = Read-BodyJson $Context.Request
            $ids = @()
            if ($body.workbookIds) { $ids = @(Get-Array $body.workbookIds | ForEach-Object { [string]$_ }) }
            $result = Start-RenderJob $language $ids ([bool]$body.onlyUpdated) ([string]$body.category) $null ([string](Get-DataProperty $body 'excelSheetSelectionMode' ''))
            $jobIdForResponse = Normalize-RenderJobId ([string](Get-DataProperty $result 'jobId' ''))
            if ([string]::IsNullOrWhiteSpace($jobIdForResponse)) { try { $jobIdForResponse = Normalize-RenderJobId ([string]$result.jobId) } catch { } }
            if ([string]::IsNullOrWhiteSpace($jobIdForResponse)) { $jobIdForResponse = Normalize-RenderJobId ([string]$result) }
            if ([string]::IsNullOrWhiteSpace($jobIdForResponse)) { throw 'PDF作成ジョブの情報を作成できませんでした。もう一度「PDF作成」を押してください。' }
            Set-NoteProperty $result 'jobId' $jobIdForResponse
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; jobId = $jobIdForResponse; job = $result; state = (Get-StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/workbooks/render') {
            $body = Read-BodyJson $Context.Request
            $structure = Get-Structure $language
            $ids = @()
            if ($body.workbookIds) { $ids = @(Get-Array $body.workbookIds | ForEach-Object { [string]$_ }) }
            elseif ($body.onlyUpdated -eq $true) { $ids = @(Get-Array $structure.workbooks | Where-Object { [string]$_.status -in @('new','excel-updated','source-updated','render-error') -or [string]$_.currentExcelHash -ne [string]$_.lastRenderedExcelHash } | ForEach-Object { [string]$_.workbookId }) }
            else { $ids = @(Get-Array $structure.workbooks | ForEach-Object { [string]$_.workbookId }) }
            $results = @()
            foreach ($id in $ids) {
                try {
                    $r = Render-Source $language $id $null $false $null '' '' '' ([string](Get-DataProperty $body 'excelSheetSelectionMode' ''))
                    $r['ok'] = $true
                    $results += $r
                } catch {
                    $msg = $_.Exception.Message
                    $userMsg = ConvertTo-UserRenderError $msg
                    $att = Get-LastRenderAttemptFor $id
                    $errorDetail = Get-ErrorDetail $_
                    Set-WorkbookRenderError $language $id $msg $errorDetail ([string]$att.snapshotId) ([string]$att.hash)
                    $results += [ordered]@{ ok = $false; workbookId = $id; error = $msg; userError = $userMsg; detail = $errorDetail }
                }
            }
            $hasErrors = $false
            foreach ($rr in $results) { try { if ($rr.Contains('ok') -and $rr['ok'] -eq $false) { $hasErrors = $true } } catch { } }
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; results = $results; hasErrors = $hasErrors; state = (Get-StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/pages/reorder') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = (Reorder-Pages $language $body); state = (Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -in @('/api/pages/sort-by-numeric-sheet','/api/pages/sort-by-sheet')) {
            $body=Read-BodyJson $Context.Request
            # Keep the old route as a compatibility alias. Its behavior is now
            # explicitly half-width numeric-sheet ordering rather than Excel
            # tab order.
            Write-JsonResponse $Context 200 ([ordered]@{ok=$true;result=(Sort-PagesByNumericSheet $language $body);state=(Get-V2StatePayload $language)});return
        }
        if ($method -eq 'POST' -and $path -eq '/api/pages/update') {
            $body = Read-BodyJson $Context.Request
            $pageResult = Update-Page $language $body
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; page = $pageResult.page; result = $pageResult; state = (Get-V2StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/pages/confirm') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; page = (Confirm-Page $language $body) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/scan-updates') {
            $body = Read-BodyJson $Context.Request
            $forceHash = $false
            try { $forceHash = [bool]$body.forceHash -or [bool]$body.force } catch { $forceHash = $false }
            $scanResult = Scan-Updates $language $null $forceHash
            # V5-§3.1/§3.3: 検知と同時に保存する。静止待ちより前。
            $captured = Update-InputHistoryAfterScan $language
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $scanResult; capturedSnapshots = @($captured) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/final/build') {
            $body = Read-BodyJson $Context.Request
            $result = Build-FinalPdf $language ([string]$body.volume) ([string]$body.category)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/final/build-all') {
            $body = Read-BodyJson $Context.Request
            $cat = Require-WorkbookCategory ([string]$body.category)
            $vols = @(Get-VolumeList $language | Where-Object { $_ -ne 'none' })
            $result = Invoke-FinalBuildTransaction $language $cat $vols
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/final/publish') {
            $body = Read-BodyJson $Context.Request
            $result = Publish-FinalPdfToShared $language ([string]$body.volume) ([string]$body.category)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result }); return
        }

        # ---- V5: 履歴・差分・自動処理 ----
        if ($method -eq 'GET' -and $path -eq '/api/history/timeline') {
            $limit = 100
            try { $limit = [int]$Context.Request.QueryString['limit'] } catch { }
            if ($limit -le 0 -or $limit -gt 500) { $limit = 100 }
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; events = (Get-HistoryTimeline $language $limit) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/history/snapshots') {
            $wbId = [string]$Context.Request.QueryString['workbookId']
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; snapshots = (Get-SnapshotSummaries $language $wbId) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/history/diff') {
            $wbId = [string]$Context.Request.QueryString['workbookId']
            $fromId = [string]$Context.Request.QueryString['fromSnapshotId']
            $toId = [string]$Context.Request.QueryString['toSnapshotId']
            $diff = $null
            if ([string]::IsNullOrWhiteSpace($fromId) -and [string]::IsNullOrWhiteSpace($toId)) {
                $diff = Get-LatestComparison $language $wbId
            } else {
                $fromVer = Get-PreferredHistoryRenderVersion $language $wbId $fromId
                $toVer = Get-PreferredHistoryRenderVersion $language $wbId $toId
                $diff = if ($fromVer -and $toVer) { Get-StoredComparison $language $wbId $fromId $toId $fromVer $toVer 'history' } else { $null }
            }
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; diff = $diff }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/history/diff-detail') {
            $wbId = [string]$Context.Request.QueryString['workbookId']
            $fromId = [string]$Context.Request.QueryString['fromSnapshotId']
            $toId = [string]$Context.Request.QueryString['toSnapshotId']
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; detail = (Get-DiffDetail $language $wbId $fromId $toId) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/history/diff-review') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; review = (Set-DiffReviewState $language $body) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/history/diff/prepare') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; job = (Start-DiffDetailJob $language ([string]$body.workbookId) ([string]$body.fromSnapshotId) ([string]$body.toSnapshotId) ([string]$body.sheetKey)) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/history/diff-page') {
            Serve-DiffPage $Context $language `
                ([string]$Context.Request.QueryString['workbookId']) `
                ([string]$Context.Request.QueryString['currentSnapshotId']) `
                ([string]$Context.Request.QueryString['baselineSnapshotId']) `
                ([string]$Context.Request.QueryString['sheetKey']) `
                ([string]$Context.Request.QueryString['pageNumber']) `
                ([string]$Context.Request.QueryString['asset']) `
                ([string]$Context.Request.QueryString['scope'])
            return
        }
        if ($method -eq 'GET' -and $path -eq '/api/history/content-pdf') {
            Serve-HistoryContentPdf $Context $language ([string]$Context.Request.QueryString['workbookId']) ([string]$Context.Request.QueryString['versionId']) ([string]$Context.Request.QueryString['sheetName']) ([string]$Context.Request.QueryString['snapshotId']); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/history/pin') {
            $body = Read-BodyJson $Context.Request
            $wbId = Assert-SafeStorageSegment ([string]$body.workbookId) 'workbookId'
            $snapshotId = Assert-SafeStorageSegment ([string]$body.snapshotId) 'snapshotId'
            [void](Set-ManualSnapshotPin $language $wbId $snapshotId $true)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/history/unpin') {
            $body = Read-BodyJson $Context.Request
            $wbId = Assert-SafeStorageSegment ([string]$body.workbookId) 'workbookId'
            $snapshotId = Assert-SafeStorageSegment ([string]$body.snapshotId) 'snapshotId'
            [void](Set-ManualSnapshotPin $language $wbId $snapshotId $false)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/layout/snapshots') {
            $cat = Require-WorkbookCategory ([string]$Context.Request.QueryString['category'])
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; snapshots = (Get-LayoutSnapshots $language $cat) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/layout/restore/preview') {
            $body = Read-BodyJson $Context.Request
            $cat = Require-WorkbookCategory ([string]$body.category)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; preview = (Get-LayoutRestorePreview $language $cat ([string]$body.snapshotId)) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/layout/restore') {
            $body = Read-BodyJson $Context.Request
            $cat = Require-WorkbookCategory ([string]$body.category)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = (Restore-LayoutSnapshot $language $cat ([string]$body.snapshotId)) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/final/archives') {
            $cat = Require-WorkbookCategory ([string]$Context.Request.QueryString['category'])
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; archives = (Get-FinalArchives $language $cat) }); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/auto/state') {
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; auto = (Get-AutoStateSummary $language) }); return
        }
        if ($method -eq 'PATCH' -and $path -eq '/api/auto/settings') {
            $body = Read-BodyJson $Context.Request
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; auto = (Update-AutoRenderSettings $body) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/auto/run-now') {
            $body = Read-BodyJson $Context.Request
            $wbId = Assert-SafeStorageSegment ([string]$body.workbookId) 'workbookId'
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = (Request-AutoRunNow $language $wbId) }); return
        }
        Write-JsonResponse $Context 404 ([ordered]@{ ok = $false; error = 'unknown api route' })
    } catch [System.ArgumentException] {
        Write-JsonResponse $Context 400 ([ordered]@{ ok = $false; error = $_.Exception.Message })
    } catch {
        # 楽観ロックの衝突は、利用者側で解決できる状態ずれ。サーバー障害と区別する。
        if (Test-StructureConflictError $_.Exception) {
            Write-JsonResponse $Context 409 ([ordered]@{
                ok = $false; code = 'structure-conflict'; error = $_.Exception.Message
                currentLayout = [string]$_.Exception.Data['actualLayout']
            })
            return
        }
        Write-JsonResponse $Context 500 ([ordered]@{ ok = $false; error = $_.Exception.Message; detail = (Get-ErrorDetail $_) })
    }
}


function Test-TcpContext($Context) {
    return ($null -ne $Context -and $null -ne $Context.PSObject.Properties['IsTcp'] -and $Context.IsTcp -eq $true)
}

function Get-HttpStatusText([int]$Status) {
    switch ($Status) {
        200 { return 'OK' }
        204 { return 'No Content' }
        400 { return 'Bad Request' }
        403 { return 'Forbidden' }
        404 { return 'Not Found' }
        500 { return 'Internal Server Error' }
        default { return 'OK' }
    }
}

function Write-TcpResponse($Context, [int]$Status, [byte[]]$Bytes, [string]$ContentType, [bool]$AllowCors = $false, [string]$CacheControl = 'no-store') {
    try {
        $statusText = Get-HttpStatusText $Status
        if ([string]::IsNullOrWhiteSpace($CacheControl)) { $CacheControl = 'no-store' }
        $CacheControl = $CacheControl -replace "[\r\n]", ''
        # CORSヘッダーは起動待ちページ(file://)が読む必要のある軽量エンドポイントだけに付ける。
        # トークン保護APIに Access-Control-Allow-Origin: * を付けると、トークン漏えい時に
        # 任意のWebページから応答を読めてしまうため、既定では付けない（同一オリジンには不要）。
        $corsHeader = ''
        if ($AllowCors) { $corsHeader = "Access-Control-Allow-Origin: *`r`n" }
        $header = "HTTP/1.1 $Status $statusText`r`nContent-Type: $ContentType`r`nContent-Length: $($Bytes.Length)`r`nCache-Control: $CacheControl`r`n${corsHeader}Connection: close`r`n`r`n"
        $headerBytes = [Text.Encoding]::ASCII.GetBytes($header)
        $stream = $Context.TcpStream
        $stream.Write($headerBytes, 0, $headerBytes.Length)
        if ($Bytes.Length -gt 0) { $stream.Write($Bytes, 0, $Bytes.Length) }
        $stream.Flush()
    } finally {
        try { if ($Context.TcpStream) { $Context.TcpStream.Close() } } catch {}
        try { if ($Context.TcpClient) { $Context.TcpClient.Close() } } catch {}
    }
}

function Find-HttpHeaderEnd([byte[]]$Bytes) {
    if ($Bytes.Length -lt 4) { return -1 }
    for ($i = 3; $i -lt $Bytes.Length; $i++) {
        if ($Bytes[$i - 3] -eq 13 -and $Bytes[$i - 2] -eq 10 -and $Bytes[$i - 1] -eq 13 -and $Bytes[$i] -eq 10) {
            return ($i - 3)
        }
    }
    return -1
}

function New-NameValueCollectionCompat {
    try { return New-Object System.Collections.Specialized.NameValueCollection ([StringComparer]::OrdinalIgnoreCase) }
    catch { return New-Object System.Collections.Specialized.NameValueCollection }
}

function Decode-UrlPart([string]$Value) {
    if ($null -eq $Value) { return '' }
    return [Uri]::UnescapeDataString(($Value -replace '\+', ' '))
}

function New-QueryStringCollection([string]$Query) {
    $nvc = New-NameValueCollectionCompat
    # NameValueCollection is enumerable. Returning it normally makes PowerShell
    # unwrap its values into Object[] (or a scalar for a single query item), so
    # Request.QueryString['token'] can never retrieve the token. Keep the
    # collection as one pipeline object on every return path.
    if ([string]::IsNullOrEmpty($Query)) { return ,$nvc }
    $q = $Query
    if ($q.StartsWith('?')) { $q = $q.Substring(1) }
    if ([string]::IsNullOrEmpty($q)) { return ,$nvc }
    foreach ($pair in ($q -split '&')) {
        if ([string]::IsNullOrEmpty($pair)) { continue }
        $eq = $pair.IndexOf('=')
        if ($eq -ge 0) {
            $name = Decode-UrlPart $pair.Substring(0, $eq)
            $value = Decode-UrlPart $pair.Substring($eq + 1)
        } else {
            $name = Decode-UrlPart $pair
            $value = ''
        }
        $nvc.Add($name, $value)
    }
    return ,$nvc
}

function Read-TcpHttpContext($TcpClient, [int]$Port) {
    $stream = $TcpClient.GetStream()
    # Browsers can open speculative/idle local connections before sending a request.
    # The server handles requests sequentially, so a long header timeout can freeze startup.
    $stream.ReadTimeout = 2000
    $buffer = New-Object byte[] 8192
    $ms = New-Object IO.MemoryStream
    $headerEnd = -1
    while ($headerEnd -lt 0) {
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -le 0) { throw 'Empty HTTP request.' }
        $ms.Write($buffer, 0, $read)
        $data = $ms.ToArray()
        $headerEnd = Find-HttpHeaderEnd $data
        if ($ms.Length -gt 65536 -and $headerEnd -lt 0) { throw 'HTTP header is too large.' }
    }

    $data = $ms.ToArray()
    $headerText = [Text.Encoding]::ASCII.GetString($data, 0, $headerEnd)
    $lines = $headerText -split "`r?`n"
    if ($lines.Count -lt 1) { throw 'Invalid HTTP request.' }
    $requestLine = $lines[0] -split ' '
    if ($requestLine.Count -lt 2) { throw 'Invalid HTTP request line.' }
    $method = $requestLine[0]
    $target = $requestLine[1]

    $headers = New-NameValueCollectionCompat
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $colon = $line.IndexOf(':')
        if ($colon -gt 0) {
            $headers.Add($line.Substring(0, $colon).Trim(), $line.Substring($colon + 1).Trim())
        }
    }

    $contentLength = 0
    $contentLengthText = [string]$headers['Content-Length']
    if (-not [string]::IsNullOrWhiteSpace($contentLengthText)) {
        if (-not [int]::TryParse($contentLengthText, [ref]$contentLength) -or $contentLength -lt 0) {
            throw [System.ArgumentException]::new('Invalid Content-Length header.')
        }
        # ReportBinder only accepts small JSON commands. A limit keeps a malformed
        # or hostile local request from blocking the single-threaded listener.
        if ($contentLength -gt 1048576) {
            throw [System.ArgumentException]::new('Request body is too large.')
        }
    }
    if ($contentLength -gt 0) { $stream.ReadTimeout = 10000 }
    $bodyMs = New-Object IO.MemoryStream
    $bodyStart = $headerEnd + 4
    if ($data.Length -gt $bodyStart -and $contentLength -gt 0) {
        $available = [Math]::Min($data.Length - $bodyStart, $contentLength)
        if ($available -gt 0) { $bodyMs.Write($data, $bodyStart, $available) }
    }
    while ($bodyMs.Length -lt $contentLength) {
        $needed = [Math]::Min($buffer.Length, $contentLength - [int]$bodyMs.Length)
        $read = $stream.Read($buffer, 0, $needed)
        if ($read -le 0) { break }
        $bodyMs.Write($buffer, 0, $read)
    }
    if ($bodyMs.Length -ne $contentLength) {
        throw [System.ArgumentException]::new('Incomplete HTTP request body.')
    }
    $bodyMs.Position = 0

    $pathOnly = $target
    $query = ''
    $qmark = $target.IndexOf('?')
    if ($qmark -ge 0) {
        $pathOnly = $target.Substring(0, $qmark)
        $query = $target.Substring($qmark + 1)
    }
    if ([string]::IsNullOrWhiteSpace($pathOnly)) { $pathOnly = '/' }
    $uri = New-Object System.Uri("http://127.0.0.1:$Port$target")

    $request = [pscustomobject]@{
        Url = $uri
        HttpMethod = $method
        Headers = $headers
        QueryString = (New-QueryStringCollection $query)
        InputStream = $bodyMs
        ContentEncoding = [Text.Encoding]::UTF8
    }
    return [pscustomobject]@{
        IsTcp = $true
        TcpClient = $TcpClient
        TcpStream = $stream
        Request = $request
        Response = [pscustomobject]@{}
    }
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


function Open-ReportBinderBrowser([string]$Url) {
    Write-Host "ReportBinder URL: $Url"
    $target = ConvertTo-EdgeOpenTarget $Url
    $edge = Get-EdgeExecutablePath
    if (-not [string]::IsNullOrWhiteSpace($edge)) {
        try { Start-Process -FilePath $edge -ArgumentList @($target) -ErrorAction Stop | Out-Null; return $true }
        catch { Write-Warning ("Could not open Microsoft Edge: " + $_.Exception.Message) }
    } else {
        Write-Warning "Microsoft Edge executable path was not resolved. Trying protocol/command fallbacks."
    }
    if ($target -match '^https?://') {
        try { Start-Process -FilePath ("microsoft-edge:" + $target) -ErrorAction Stop | Out-Null; return $true }
        catch { Write-Warning ("Could not open Microsoft Edge via protocol fallback: " + $_.Exception.Message) }
    }
    try {
        $cmdExe = Join-Path $env:SystemRoot 'System32\cmd.exe'
        if (-not (Test-Path -LiteralPath $cmdExe)) { $cmdExe = 'cmd.exe' }
        $escapedTarget = $target -replace '"','""'
        $cmdArgs = '/c start "" msedge.exe "' + $escapedTarget + '"'
        Start-Process -FilePath $cmdExe -ArgumentList $cmdArgs -WindowStyle Hidden -ErrorAction Stop | Out-Null
        return $true
    } catch { Write-Warning ("Could not open Microsoft Edge via command fallback: " + $_.Exception.Message) }
    return $false
}

function Start-LocalTcpServer([int]$ListenPort, [string]$OpenUrl, [bool]$SkipOpen) {
    # listener を先に確立し、起動待ちをスケジューラー子プロセスの初期化で遅らせない。
    # 以降の失敗時も finally で必ず子プロセスを停止する。
    $tcp = $null
    try {
        $tcp = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Parse('127.0.0.1'), $ListenPort)
        $tcp.Start()
        try { [void](Start-AutoSchedulerProcess (Get-EffectiveLanguage)) } catch { Write-Warning $_.Exception.Message }
        Write-Host "ReportBinder local server started on 127.0.0.1:$ListenPort"
        Write-Host "ReportBinder URL: $OpenUrl"
        Write-Host "ReportBinder will stop after 30 minutes without browser activity. PDF creation jobs continue even if the tab is closed."
        if (-not $SkipOpen) { Open-ReportBinderBrowser $OpenUrl }
        while ($true) {
            $now = [DateTime]::UtcNow
            if ($Script:ShutdownRequested) {
                Write-Host "ReportBinder shutdown requested."
                break
            }
            if (-not $Script:ClientAttached -and (($now - $Script:ServerStartedUtc).TotalSeconds -ge $Script:NoClientStartupTimeoutSeconds)) {
                if (-not (Test-ActiveRenderJobs $Mode)) {
                    Write-Host "ReportBinder browser was not opened or attached. Stopping local server."
                    break
                }
            }
            if ($Script:ClientAttached -and (($now - $Script:LastHeartbeatUtc).TotalSeconds -ge $Script:IdleTimeoutSeconds)) {
                if (Test-ActiveRenderJobs $Mode) {
                    Write-Host "ReportBinder is idle, but a PDF creation job is still running. Keeping local server alive."
                    $Script:LastHeartbeatUtc = $now
                } else {
                    Write-Host "ReportBinder idle timeout reached. Stopping local server."
                    break
                }
            }

            if (-not $tcp.Pending()) {
                Start-Sleep -Milliseconds 250
                continue
            }

            $client = $tcp.AcceptTcpClient()
            try {
                $context = Read-TcpHttpContext $client $ListenPort
                $path = $context.Request.Url.AbsolutePath
                if ($path.StartsWith('/api/')) { Handle-Api $context } else { Serve-Static $context $path }
            } catch {
                try {
                    $ctx = [pscustomobject]@{ IsTcp = $true; TcpClient = $client; TcpStream = $client.GetStream(); Request = $null; Response = [pscustomobject]@{} }
                    $err = [ordered]@{ ok = $false; error = $_.Exception.Message }
                    $status = 500
                    if ($_.Exception -is [System.ArgumentException]) { $status = 400 }
                    elseif (Test-StructureConflictError $_.Exception) { $status = 409; $err['code'] = 'structure-conflict' }
                    Write-JsonResponse $ctx $status $err
                } catch {
                    try { $client.Close() } catch {}
                }
            }
        }
    } finally {
        # 親PID監視だけに頼らず、通常終了・listener起動失敗のどちらでも明示停止する。
        try { Stop-AutoSchedulerProcess } catch { }
        if ($tcp) { try { $tcp.Stop() } catch { } }
    }
}

function Get-FreePort {
    $listener = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Parse('127.0.0.1'), 0)
    $listener.Start()
    $p = $listener.LocalEndpoint.Port
    $listener.Stop()
    return $p
}


# =====================================================================
# V5 Stage 2 — Phase 1A: 検知版スナップショット (input history)
# =====================================================================

function Get-InputHistoryRoot([string]$Language) {
    return (Join-Path (Get-WorkspacePath $Language) 'input-history')
}
function Get-WorkbookHistoryDir([string]$Language, [string]$WorkbookId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    return (Join-Path (Get-InputHistoryRoot $Language) $safeWorkbookId)
}
function Get-SnapshotDir([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    return (Join-Path (Get-WorkbookHistoryDir $Language $WorkbookId) $safeSnapshotId)
}
function Get-EphemeralJobRoot([string]$Language) {
    return (Join-Path (Get-WorkspacePath $Language) 'state\jobs')
}

function Write-HistoryEvent([string]$Language, [string]$Type, $Data) {
    # 履歴タイムラインUIを廃止したため、共有フォルダーへの1イベント1ファイル書込も行わない。
    # 版比較に必要なsnapshot/manifest/comparisonは別経路で保持される。
    return
}

function Get-SnapshotManifest([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $cacheKey = ($Language + '|' + $safeWorkbookId + '|' + $safeSnapshotId).ToLowerInvariant()
    $path = Join-Path (Get-SnapshotDir $Language $safeWorkbookId $safeSnapshotId) 'manifest.json'
    if ($Script:SnapshotManifestCache.ContainsKey($cacheKey)) {
        if (Test-Path -LiteralPath $path -PathType Leaf) { return $Script:SnapshotManifestCache[$cacheKey] }
        [void]$Script:SnapshotManifestCache.Remove($cacheKey)
    }
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try {
        $manifest = Read-JsonFile $path $null
        if ($null -ne $manifest) { $Script:SnapshotManifestCache[$cacheKey] = $manifest }
        return $manifest
    } catch { return $null }
}

function Get-SnapshotIds([string]$Language, [string]$WorkbookId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $dir = Get-WorkbookHistoryDir $Language $safeWorkbookId
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    $stamp = [IO.Directory]::GetLastWriteTimeUtc($dir).Ticks
    $cacheKey = ($Language + '|' + $safeWorkbookId).ToLowerInvariant()
    $cached = $Script:SnapshotIdCache[$cacheKey]
    if ($null -ne $cached -and [Int64](Get-DataProperty $cached 'stamp' -1) -eq $stamp) {
        return @(Get-Array (Get-DataProperty $cached 'ids' @()))
    }
    # manifest.json is the completion marker. Pending/crashed directories are not
    # history generations and must not consume the recent-generation allowance.
    # Enumerate completion markers in one directory walk instead of issuing one
    # remote Test-Path call per generation on a shared workspace.
    $ids = @(Get-ChildItem -LiteralPath $dir -File -Filter 'manifest.json' -Recurse -Depth 1 -ErrorAction SilentlyContinue |
             Where-Object { [string]$_.Directory.Parent.FullName -eq [string]$dir } |
             Sort-Object { $_.Directory.Name } | ForEach-Object { [string]$_.Directory.Name })
    $Script:SnapshotIdCache[$cacheKey] = [pscustomobject][ordered]@{ stamp = $stamp; ids = $ids }
    return $ids
}

function Get-PackReviewSnapshot($Structure, [string]$Language, [string]$PackId, [bool]$CheckOutputExists = $true) {
    $pack = Get-PackRecord $Structure $PackId
    $template = Get-PackEffectiveTemplate $Language $pack
    $targets = @(Get-Array (Get-DataProperty $template 'targets' @()) | Where-Object { [bool](Get-DataProperty $_ 'required' $false) })
    if ($targets.Count -eq 0) { $targets = @(Get-Array (Get-DataProperty $template 'targets' @()) | Select-Object -First 1) }
    $fingerprints = [ordered]@{}; $targetStates = @(); $canSubmit = ($targets.Count -gt 0)
    foreach ($target in $targets) {
        $targetId = [string](Get-DataProperty $target 'targetId' '')
        $readiness = Get-FinalBuildReadiness $Structure $Language (Get-LegacyVolumeFromTargetId $Language $targetId) $PackId $CheckOutputExists
        $fingerprints[$targetId] = [string](Get-DataProperty $readiness 'currentFingerprint' '')
        $ready = [string](Get-DataProperty $readiness 'displayState' '') -eq 'built'
        if (-not $ready) { $canSubmit = $false }
        $targetStates += [pscustomobject][ordered]@{
            targetId=$targetId; displayName=[string](Get-DataProperty $target 'displayName' $targetId)
            ready=$ready; displayState=[string](Get-DataProperty $readiness 'displayState' '')
            fingerprint=[string](Get-DataProperty $readiness 'currentFingerprint' '')
            outputPdf=[string](Get-DataProperty $readiness 'outputPdf' '')
        }
    }
    $stored = Get-DataProperty $pack 'review' ([ordered]@{})
    $storedStatus = [string](Get-DataProperty $stored 'status' 'draft')
    if ($storedStatus -notin @('draft','in-review','approved','changes-requested')) { $storedStatus = 'draft' }
    $submitted = Get-DataProperty $stored 'submittedFingerprints' ([ordered]@{})
    $stale = $false
    if ($storedStatus -in @('in-review','approved')) {
        foreach ($target in $targets) {
            $targetId = [string](Get-DataProperty $target 'targetId' '')
            if ([string](Get-DataProperty $submitted $targetId '') -ne [string]$fingerprints[$targetId]) { $stale = $true; break }
        }
    }
    return [pscustomobject][ordered]@{
        packId=$PackId; status=$(if ($stale) { 'stale' } else { $storedStatus }); storedStatus=$storedStatus
        canSubmit=[bool]$canSubmit; stale=[bool]$stale; targetStates=@($targetStates); currentFingerprints=$fingerprints
        submittedFingerprints=$submitted; submittedAt=[string](Get-DataProperty $stored 'submittedAt' '')
        submittedBy=[string](Get-DataProperty $stored 'submittedBy' ''); approvedAt=[string](Get-DataProperty $stored 'approvedAt' '')
        approvedBy=[string](Get-DataProperty $stored 'approvedBy' ''); note=[string](Get-DataProperty $stored 'note' '')
        events=@(Get-Array (Get-DataProperty $stored 'events' @()) | Select-Object -Last 100)
    }
}

function Copy-PackReviewFingerprints($Fingerprints) {
    $copy = [ordered]@{}
    if ($null -eq $Fingerprints) { return $copy }
    if ($Fingerprints -is [Collections.IDictionary]) {
        foreach ($key in @($Fingerprints.Keys)) { $copy[[string]$key] = [string]$Fingerprints[$key] }
        return $copy
    }
    foreach ($property in @($Fingerprints.PSObject.Properties)) { $copy[[string]$property.Name] = [string]$property.Value }
    return $copy
}

function Invoke-PackReviewAction([string]$Language, [string]$PackId, $Body) {
    $id = Assert-SafeStorageSegment $PackId 'packId'
    # Avoid the name `$action`: Invoke-WithLock has a case-insensitive `$Action`
    # parameter, which would shadow this value inside the mutation scriptblock.
    $reviewAction = ([string](Get-DataProperty $Body 'action' '')).Trim().ToLowerInvariant()
    if ($reviewAction -notin @('submit','approve','request-changes','reopen')) { throw [ArgumentException]::new('review action が不正です。') }
    $note = ([string](Get-DataProperty $Body 'note' '')).Trim()
    if ($note.Length -gt 1000) { throw [ArgumentException]::new('レビューコメントは1000文字以内で入力してください。') }
    if ($reviewAction -eq 'request-changes' -and [string]::IsNullOrWhiteSpace($note)) { throw [ArgumentException]::new('差し戻し理由を入力してください。') }
    $actor = ([string](Get-DataProperty $Body 'actor' $env:USERNAME)).Trim()
    if ([string]::IsNullOrWhiteSpace($actor)) { $actor = 'local-user' }
    if ($actor.Length -gt 120) { throw [ArgumentException]::new('確認者名は120文字以内で入力してください。') }
    # PDF readiness evaluation may invoke the PDF runtime. Keep that work out of
    # the structure file lock, then reject the action if another writer changed
    # the structure before the short persistence transaction begins.
    $reviewStructure = Get-Structure $Language
    $expectedStructureUpdatedAt = [string](Get-DataProperty $reviewStructure 'updatedAt' '')
    $snapshot = Get-PackReviewSnapshot $reviewStructure $Language $id $true
    return Update-StructureLocked $Language {
        param($structure)
        if ([string](Get-DataProperty $structure 'updatedAt' '') -ne $expectedStructureUpdatedAt) {
            throw [InvalidOperationException]::new('資料構成が同時に更新されました。最新状態を確認して、もう一度操作してください。')
        }
        $pack = Get-PackRecord $structure $id
        $stored = Get-DataProperty $pack 'review' ([ordered]@{})
        $now = New-NowIso
        switch ($reviewAction) {
            'submit' {
                if (-not [bool]$snapshot.canSubmit) { throw [InvalidOperationException]::new('必須の提出用PDFをすべて最新にしてからレビューへ提出してください。') }
                Set-NoteProperty $stored 'status' 'in-review'; Set-NoteProperty $stored 'submittedFingerprints' (Copy-PackReviewFingerprints $snapshot.currentFingerprints)
                Set-NoteProperty $stored 'submittedAt' $now; Set-NoteProperty $stored 'submittedBy' $actor
                Set-NoteProperty $stored 'approvedAt' ''; Set-NoteProperty $stored 'approvedBy' ''
            }
            'approve' {
                if ([string]$snapshot.status -ne 'in-review') { throw [InvalidOperationException]::new('レビュー提出後で、内容が更新されていない資料だけを承認できます。') }
                Set-NoteProperty $stored 'status' 'approved'; Set-NoteProperty $stored 'approvedAt' $now; Set-NoteProperty $stored 'approvedBy' $actor
            }
            'request-changes' {
                if ([string]$snapshot.status -notin @('in-review','approved')) { throw [InvalidOperationException]::new('レビュー中または承認済みの資料だけを差し戻せます。') }
                Set-NoteProperty $stored 'status' 'changes-requested'; Set-NoteProperty $stored 'approvedAt' ''; Set-NoteProperty $stored 'approvedBy' ''
            }
            'reopen' { Set-NoteProperty $stored 'status' 'draft'; Set-NoteProperty $stored 'submittedFingerprints' ([ordered]@{}); Set-NoteProperty $stored 'submittedAt' ''; Set-NoteProperty $stored 'submittedBy' ''; Set-NoteProperty $stored 'approvedAt' ''; Set-NoteProperty $stored 'approvedBy' '' }
        }
        Set-NoteProperty $stored 'note' $note
        $eventFingerprints = Copy-PackReviewFingerprints (Get-DataProperty $stored 'submittedFingerprints' ([ordered]@{}))
        $reviewEvent = [pscustomobject][ordered]@{
            eventId='review-' + [Guid]::NewGuid().ToString('N').Substring(0,16); action=$reviewAction; status=[string](Get-DataProperty $stored 'status' 'draft')
            actor=$actor; note=$note; at=$now; fingerprints=$eventFingerprints
        }
        $events = @(Get-Array (Get-DataProperty $stored 'events' @())) + @($reviewEvent)
        Set-NoteProperty $stored 'events' @($events | Select-Object -Last 100)
        Set-NoteProperty $pack 'review' $stored; Set-NoteProperty $pack 'updatedAt' $now
        # The readiness snapshot above already contains the exact fingerprints
        # validated by this action. Recomputing it here renders every source PDF
        # a second time and makes a single review click unnecessarily expensive.
        Set-NoteProperty $snapshot 'status' ([string](Get-DataProperty $stored 'status' 'draft'))
        Set-NoteProperty $snapshot 'storedStatus' ([string](Get-DataProperty $stored 'status' 'draft'))
        Set-NoteProperty $snapshot 'stale' $false
        Set-NoteProperty $snapshot 'submittedFingerprints' (Copy-PackReviewFingerprints (Get-DataProperty $stored 'submittedFingerprints' ([ordered]@{})))
        Set-NoteProperty $snapshot 'submittedAt' ([string](Get-DataProperty $stored 'submittedAt' ''))
        Set-NoteProperty $snapshot 'submittedBy' ([string](Get-DataProperty $stored 'submittedBy' ''))
        Set-NoteProperty $snapshot 'approvedAt' ([string](Get-DataProperty $stored 'approvedAt' ''))
        Set-NoteProperty $snapshot 'approvedBy' ([string](Get-DataProperty $stored 'approvedBy' ''))
        Set-NoteProperty $snapshot 'note' ([string](Get-DataProperty $stored 'note' ''))
        Set-NoteProperty $snapshot 'events' @(Get-Array (Get-DataProperty $stored 'events' @()) | Select-Object -Last 100)
        return $snapshot
    }
}

function Clear-SnapshotRuntimeCaches([string]$Language, [string]$WorkbookId, [string]$SnapshotId = '') {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $workbookPrefix = ($Language + '|' + $safeWorkbookId + '|').ToLowerInvariant()
    $snapshotKey = if ([string]::IsNullOrWhiteSpace($SnapshotId)) { '' } else { ($workbookPrefix + (Assert-SafeStorageSegment $SnapshotId 'snapshotId').ToLowerInvariant()) }
    foreach ($key in @($Script:SnapshotManifestCache.Keys)) {
        if (($snapshotKey -and [string]$key -eq $snapshotKey) -or (-not $snapshotKey -and [string]$key -like ($workbookPrefix + '*'))) {
            [void]$Script:SnapshotManifestCache.Remove($key)
        }
    }
    foreach ($key in @($Script:VisualHashCache.Keys)) {
        if (($snapshotKey -and [string]$key -like ($snapshotKey + '|*')) -or (-not $snapshotKey -and [string]$key -like ($workbookPrefix + '*'))) {
            [void]$Script:VisualHashCache.Remove($key)
        }
    }
    Clear-SnapshotSummaryCache $Language $safeWorkbookId
}

function Find-SnapshotBySourceHash([string]$Language, [string]$WorkbookId, [string]$SourceHash) {
    $target = Normalize-FileHash $SourceHash
    if ([string]::IsNullOrWhiteSpace($target)) { return '' }
    foreach ($id in (Get-SnapshotIds $Language $WorkbookId)) {
        $m = Get-SnapshotManifest $Language $WorkbookId $id
        if ($null -eq $m) { continue }
        if ((Normalize-FileHash ([string](Get-DataProperty $m 'sourceHash' ''))) -eq $target) { return $id }
    }
    return ''
}

function Get-SnapshotSourceState([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    # V5-IV-5: 現物の有無は immutable な manifest には書かない。
    $dir = Get-SnapshotDir $Language $WorkbookId $SnapshotId
    $statePath = Join-Path $dir 'source-state.json'
    $manifest = Get-SnapshotManifest $Language $WorkbookId $SnapshotId
    $extension = [IO.Path]::GetExtension([string](Get-DataProperty $manifest 'relativePath' '')).ToLowerInvariant()
    if ($extension -notin @('.xlsx','.xlsm','.docx','.pptx','.pdf')) { $extension = $(if (Test-Path -LiteralPath (Join-Path $dir 'source.pdf')) { '.pdf' } elseif (Test-Path -LiteralPath (Join-Path $dir 'source.pptx')) { '.pptx' } elseif (Test-Path -LiteralPath (Join-Path $dir 'source.docx')) { '.docx' } elseif (Test-Path -LiteralPath (Join-Path $dir 'source.xlsm')) { '.xlsm' } else { '.xlsx' }) }
    $sourcePath = Join-Path $dir ("source$extension")
    $exists = Test-Path -LiteralPath $sourcePath
    $state = $null
    if (Test-Path -LiteralPath $statePath) { try { $state = Read-JsonFile $statePath $null } catch { } }
    return [ordered]@{
        sourceRetained = [bool]$exists
        sourcePath = $sourcePath
        removedAt = [string](Get-DataProperty $state 'removedAt' '')
        removedReason = [string](Get-DataProperty $state 'removedReason' '')
    }
}

function Set-SnapshotSourceState([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [bool]$Retained, [string]$Reason) {
    $dir = Get-SnapshotDir $Language $WorkbookId $SnapshotId
    if (-not (Test-Path -LiteralPath $dir)) { return }
    Write-JsonFile (Join-Path $dir 'source-state.json') ([ordered]@{
        schemaVersion = 1; sourceRetained = $Retained
        removedAt = $(if ($Retained) { '' } else { New-NowIso })
        removedReason = $(if ($Retained) { '' } else { [string]$Reason })
    })
}

# ---- pins / leases -------------------------------------------------

function Get-SnapshotPinDir([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    return (Join-Path (Get-SnapshotDir $Language $WorkbookId $SnapshotId) 'pins')
}
function Get-SnapshotLeaseDir([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    return (Join-Path (Get-SnapshotDir $Language $WorkbookId $SnapshotId) 'leases')
}

function New-SnapshotPin([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$PinName, $Data) {
    # V5-IV-2: 保護はファイルの作成・削除で表す。配列の書き換えはしない。
    try {
        $safePinName = Assert-SafeStorageSegment $PinName 'pinName'
        if ($null -eq (Get-SnapshotManifest $Language $WorkbookId $SnapshotId)) { return $false }
        $dir = Get-SnapshotPinDir $Language $WorkbookId $SnapshotId
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Write-JsonFile (Join-Path $dir ("{0}.json" -f $safePinName)) $Data
        return $true
    } catch { return $false }
}

function Remove-SnapshotPin([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$PinName) {
    try {
        $path = Join-Path (Get-SnapshotPinDir $Language $WorkbookId $SnapshotId) ("{0}.json" -f $PinName)
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
    } catch { }
}

function Get-SnapshotPins([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    $dir = Get-SnapshotPinDir $Language $WorkbookId $SnapshotId
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    return @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.BaseName })
}

function Set-ManualSnapshotPin([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [bool]$Pinned) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $historyLock = Join-Path (Get-WorkspacePath $Language) 'locks\history-cleanup.lock'
    return Invoke-WithLock $historyLock {
        if ($null -eq (Get-SnapshotManifest $Language $safeWorkbookId $safeSnapshotId)) {
            throw [ArgumentException]::new('保護する履歴版が見つかりません。')
        }
        if ($Pinned) {
            if (-not (New-SnapshotPin $Language $safeWorkbookId $safeSnapshotId 'manual' ([ordered]@{ pinnedAt = New-NowIso }))) {
                throw '履歴の保護情報を保存できませんでした。'
            }
        } else {
            Remove-SnapshotPin $Language $safeWorkbookId $safeSnapshotId 'manual'
        }
        [void](Update-SnapshotSummaryCacheEntry $Language $safeWorkbookId $safeSnapshotId)
        return $true
    }
}

function New-SnapshotLease([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$Purpose, [string]$JobId, [int]$MinutesValid = 30, [string]$VersionId = '') {
    try {
        $safePurpose = Assert-SafeStorageSegment $Purpose 'purpose'
        $safeJobId = Assert-SafeStorageSegment $JobId 'jobId'
        if ($null -eq (Get-SnapshotManifest $Language $WorkbookId $SnapshotId)) { return '' }
        $dir = Get-SnapshotLeaseDir $Language $WorkbookId $SnapshotId
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $name = ('{0}_{1}.json' -f $safePurpose, $safeJobId)
        Write-JsonFile (Join-Path $dir $name) ([ordered]@{
            jobId = $JobId; purpose = $Purpose; versionId = $VersionId
            createdAt = New-NowIso; heartbeatAt = New-NowIso
            expiresAt = ([DateTime]::UtcNow.AddMinutes($MinutesValid).ToString('o'))
            pcName = $env:COMPUTERNAME
        })
        return $name
    } catch { return '' }
}

function Remove-SnapshotLease([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$LeaseName) {
    try {
        if ([string]::IsNullOrWhiteSpace($LeaseName)) { return }
        $path = Join-Path (Get-SnapshotLeaseDir $Language $WorkbookId $SnapshotId) $LeaseName
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
    } catch { }
}

function Test-LeaseActive($LeaseFile) {
    try {
        $j = Read-JsonFile $LeaseFile.FullName $null
        $exp = [string](Get-DataProperty $j 'expiresAt' '')
        if ([string]::IsNullOrWhiteSpace($exp)) { return $false }
        return ([DateTime]::Parse($exp).ToUniversalTime() -gt [DateTime]::UtcNow)
    } catch { return $false }
}

function Get-ActiveLeaseCount([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    $dir = Get-SnapshotLeaseDir $Language $WorkbookId $SnapshotId
    if (-not (Test-Path -LiteralPath $dir)) { return 0 }
    $n = 0
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue)) {
        if (Test-LeaseActive $f) { $n++ } else { try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue } catch { } }
    }
    return $n
}


function Clear-ExpiredLeases([string]$Language) {
    # Snapshot and content-PDF leases both expire after abnormal termination.
    try {
        $root = Get-InputHistoryRoot $Language
        if (Test-Path -LiteralPath $root) {
            foreach ($wb in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
                foreach ($sn in @(Get-ChildItem -LiteralPath $wb.FullName -Directory -ErrorAction SilentlyContinue)) {
                    [void](Get-ActiveLeaseCount $Language $wb.Name $sn.Name)
                }
            }
        }
    } catch { }
    try {
        $contentRoot = Join-Path (Get-WorkspacePath $Language) 'content-pdf'
        if (Test-Path -LiteralPath $contentRoot) {
            foreach ($wb in @(Get-ChildItem -LiteralPath $contentRoot -Directory -ErrorAction SilentlyContinue)) {
                foreach ($ver in @(Get-ChildItem -LiteralPath $wb.FullName -Directory -ErrorAction SilentlyContinue)) {
                    $leaseDir = Join-Path $ver.FullName 'leases'
                    if (-not (Test-Path -LiteralPath $leaseDir)) { continue }
                    foreach ($f in @(Get-ChildItem -LiteralPath $leaseDir -File -Filter '*.json' -ErrorAction SilentlyContinue)) {
                        if (-not (Test-LeaseActive $f)) { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue }
                    }
                }
            }
        }
    } catch { }
}

function Ensure-SnapshotMetadata([string]$Language, [string]$WorkbookId, [string]$RelativePath, [string]$Category, [string]$SourcePath, [string]$SourceHash, [string]$CaptureReason) {
    # V5-§C: メタデータの重複排除だけを担当する。入力ファイルの確保は Capture-RenderInput の役目。
    if (-not (Test-InputHistoryEnabled)) { return [ordered]@{ ok = $false; reason = 'not-approved'; snapshotId = '' } }
    $hash = Normalize-FileHash $SourceHash
    if ([string]::IsNullOrWhiteSpace($hash)) { return [ordered]@{ ok = $false; reason = 'no-hash'; snapshotId = '' } }

    $lockPath = Join-Path (Get-WorkspacePath $Language) ("locks\snapshot_{0}.lock" -f $WorkbookId)
    return Invoke-WithLock $lockPath {
        $existing = Find-SnapshotBySourceHash $Language $WorkbookId $hash
        if (-not [string]::IsNullOrWhiteSpace($existing)) {
            Write-HistoryEvent $Language 'input.snapshot.deduplicated' ([ordered]@{ workbookId = $WorkbookId; snapshotId = $existing; sourceHash = $hash; captureReason = $CaptureReason })
            return [ordered]@{ ok = $true; snapshotId = $existing; isNew = $false }
        }
        $ids = @(Get-SnapshotIds $Language $WorkbookId)
        $previous = $(if ($ids.Count -gt 0) { [string]$ids[-1] } else { '' })
        $parent = Get-WorkbookHistoryDir $Language $WorkbookId
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
        $created = New-UniqueDirectory $parent { New-RbId }
        $snapshotId = [string]$created.id
        $item = $null
        try { $item = Get-Item -LiteralPath $SourcePath -ErrorAction Stop } catch { }
        # V5-§3.2: manifest.json の存在が完成マーカー。先に pending として書き、
        # source のコピー・検証が終わってから manifest.json へ改名する(Complete-Snapshot)。
        Write-JsonFile (Join-Path ([string]$created.path) 'manifest.pending.json') ([ordered]@{
            schemaVersion = 1
            snapshotId = $snapshotId
            workbookId = $WorkbookId
            language = $Language
            category = $Category
            relativePath = $RelativePath
            sourceHashAlgorithm = 'SHA-256'
            sourceHash = $hash
            sourceSize = $(if ($item) { $item.Length } else { 0 })
            sourceLastWriteUtcTicks = $(if ($item) { [string]$item.LastWriteTimeUtc.Ticks } else { '' })
            sourceModifiedAt = $(if ($item) { $item.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:sszzz') } else { '' })
            detectedAt = New-NowIso
            captureReason = $CaptureReason
            previousSnapshotId = $previous
            status = 'complete'
            parserVersion = 1
            capturedBy = [ordered]@{ pcName = $env:COMPUTERNAME; userName = "$env:USERDOMAIN\$env:USERNAME" }
        })
        Set-SnapshotSourceState $Language $WorkbookId $snapshotId $false 'not-captured'
        return [ordered]@{ ok = $true; snapshotId = $snapshotId; isNew = $true; pending = $true }
    }
}

function Complete-Snapshot([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    # pending から manifest.json へ改名して完成させる。以後この検知版は参照可能になる。
    $dir = Get-SnapshotDir $Language $WorkbookId $SnapshotId
    $pending = Join-Path $dir 'manifest.pending.json'
    $final = Join-Path $dir 'manifest.json'
    if (Test-Path -LiteralPath $final) { return $true }
    if (-not (Test-Path -LiteralPath $pending)) { return $false }
    try {
        Move-Item -LiteralPath $pending -Destination $final -Force
        Clear-SnapshotRuntimeCaches $Language $WorkbookId $SnapshotId
        Write-HistoryEvent $Language 'input.snapshot.created' ([ordered]@{ workbookId = $WorkbookId; snapshotId = $SnapshotId })
        return $true
    } catch { return $false }
}

function Save-SnapshotSourceFile([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$SourcePath, [string]$SourceHash) {
    # V5-§3.2 手順4-8: コピー -> ハッシュ検証 -> 履歴フォルダへ移動 -> 再ハッシュ。
    if (-not (Test-SourceRetentionEnabled)) { return [ordered]@{ ok = $false; reason = 'retention-not-approved' } }
    $dir = Get-SnapshotDir $Language $WorkbookId $SnapshotId
    if (-not (Test-Path -LiteralPath $dir)) { return [ordered]@{ ok = $false; reason = 'snapshot-missing' } }
    $extension = [IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    if ($extension -notin @('.xlsx','.xlsm','.docx','.pptx','.pdf')) { $extension = '.xlsx' }
    $dest = Join-Path $dir ("source$extension")
    $expectedHash = Normalize-FileHash $SourceHash
    if (Test-Path -LiteralPath $dest) {
        $existingHash = Normalize-FileHash (New-Sha256 $dest)
        if (-not [string]::IsNullOrWhiteSpace($existingHash) -and $existingHash -eq $expectedHash) {
            Set-SnapshotSourceState $Language $WorkbookId $SnapshotId $true ''
            return [ordered]@{ ok = $true; path = $dest; reused = $true }
        }
    }

    $tmpDir = Join-Path (Get-WorkspacePath $Language) ('state\tmp\' + (New-RbId))
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
    $tmp = Join-Path $tmpDir ("source$extension")
    try {
        $before = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
        $copied = $false; $lastError = ''
        for ($i = 1; $i -le 3; $i++) {
            try { Copy-FileSharedRead $SourcePath $tmp; $copied = $true; break } catch { $lastError = $_.Exception.Message }
            if ($i -lt 3) { Start-Sleep -Seconds 2 }
        }
        if (-not $copied) { return [ordered]@{ ok = $false; reason = 'copy-failed'; message = $lastError } }
        try { Unblock-File -LiteralPath $tmp -ErrorAction SilentlyContinue } catch { }
        $copyHash = Normalize-FileHash (New-Sha256 $tmp)
        if ($copyHash -ne $expectedHash) { return [ordered]@{ ok = $false; reason = 'hash-mismatch' } }
        $after = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
        if ($after.Length -ne $before.Length -or $after.LastWriteTimeUtc -ne $before.LastWriteTimeUtc) {
            return [ordered]@{ ok = $false; reason = 'source-changed-during-copy' }
        }
        Move-Item -LiteralPath $tmp -Destination $dest -Force
        $finalHash = Normalize-FileHash (New-Sha256 $dest)
        if ($finalHash -ne $copyHash) {
            Remove-Item -LiteralPath $dest -Force -ErrorAction SilentlyContinue
            return [ordered]@{ ok = $false; reason = 'verify-failed' }
        }
        Set-SnapshotSourceState $Language $WorkbookId $SnapshotId $true ''
        return [ordered]@{ ok = $true; path = $dest; reused = $false }
    } catch {
        return [ordered]@{ ok = $false; reason = 'error'; message = $_.Exception.Message }
    } finally {
        if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Capture-DetectedSnapshot([string]$Language, [string]$WorkbookId, [string]$CaptureReason) {
    # V5-§3.1/§3.3: 検知と同時に保存する。静止待ちより前。
    if (-not (Test-InputHistoryEnabled)) { return [ordered]@{ ok = $false; reason = 'not-approved' } }
    $paths = Get-Paths
    $structure = Get-Structure $Language
    $wb = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
    if ($wb.Count -eq 0) { return [ordered]@{ ok = $false; reason = 'workbook-missing' } }
    $w = $wb[0]
    $sourcePath = Join-Safe ([string]$paths.submissionDir) ([string]$w.relativePath)
    if (-not (Test-Path -LiteralPath $sourcePath)) { return [ordered]@{ ok = $false; reason = 'file-missing' } }

    $expectedPrevious = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
    $item = Get-Item -LiteralPath $sourcePath -ErrorAction Stop
    $capturedHash = Normalize-FileHash (New-Sha256 $sourcePath)
    if ([string]::IsNullOrWhiteSpace($capturedHash)) { return [ordered]@{ ok = $false; reason = 'hash-failed' } }
    $capturedTicks = [string]$item.LastWriteTimeUtc.Ticks
    $capturedSize = $item.Length

    $meta = Ensure-SnapshotMetadata $Language $WorkbookId ([string]$w.relativePath) ([string]$w.category) $sourcePath $capturedHash $CaptureReason
    if (-not [bool]$meta.ok) { return $meta }
    $snapshotId = [string]$meta.snapshotId
    if ([bool](Get-DataProperty $meta 'pending' $false)) {
        if (Test-SourceRetentionEnabled) {
            $saved = Save-SnapshotSourceFile $Language $WorkbookId $snapshotId $sourcePath $capturedHash
            if (-not [bool]$saved.ok) { return [ordered]@{ ok = $false; reason = [string]$saved.reason } }
        }
        if (-not (Complete-Snapshot $Language $WorkbookId $snapshotId)) { return [ordered]@{ ok = $false; reason = 'complete-failed' } }
    }

    # V5-§3.3: コミットは compare-and-set。PC の時刻ではなく提出ファイルの実状態で判定する。
    $committed = Update-StructureLocked $Language {
        param($st)
        $x = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
        if (-not $x.Count) { return [ordered]@{ committed = $false; reason = 'workbook-missing' } }
        $t = $x[0]
        $live = $null
        try { $live = Get-Item -LiteralPath $sourcePath -ErrorAction Stop } catch { }
        $stillCurrent = ($null -ne $live) -and
                        ([string]$live.LastWriteTimeUtc.Ticks -eq $capturedTicks) -and
                        ($live.Length -eq $capturedSize)
        if (-not $stillCurrent) { return [ordered]@{ committed = $false; reason = 'superseded' } }
        # V5-P1(#9): ticks/size が同一でも内容が書き換わっている可能性があるため、
        # コミット直前に実ファイルを再ハッシュし、捕捉時のハッシュと一致することを確認する。
        $liveHash = Normalize-FileHash (New-Sha256 $sourcePath)
        if ([string]::IsNullOrWhiteSpace($liveHash) -or $liveHash -ne $capturedHash) {
            return [ordered]@{ committed = $false; reason = 'superseded' }
        }
        $latest = Normalize-FileHash ([string](Get-DataProperty $t 'currentExcelHash' ''))
        if (-not ([string]::IsNullOrWhiteSpace($latest)) -and $latest -ne $expectedPrevious -and $latest -ne $capturedHash) {
            return [ordered]@{ committed = $false; reason = 'concurrent-update' }
        }
        Add-NotePropertyIfMissing $t 'currentSnapshotId' ''
        Add-NotePropertyIfMissing $t 'lastDetectedAt' ''
        Set-NoteProperty $t 'currentExcelHash' $capturedHash
        Set-NoteProperty $t 'currentExcelLastWriteUtcTicks' $capturedTicks
        Set-NoteProperty $t 'currentExcelSize' $capturedSize
        Set-NoteProperty $t 'currentExcelModifiedAt' ($live.LastWriteTime.ToString('yyyy-MM-ddTHH:mm:sszzz'))
        Set-NoteProperty $t 'currentSnapshotId' $snapshotId
        Set-NoteProperty $t 'lastDetectedAt' (New-NowIso)
        $rendered = Normalize-FileHash ([string](Get-DataProperty $t 'lastRenderedExcelHash' ''))
        if ($rendered -and $rendered -ne $capturedHash -and [string]$t.status -ne 'render-error') { Set-NoteProperty $t 'status' (Get-SourceUpdatedStatus $t) }
        return [ordered]@{ committed = $true }
    }
    return [ordered]@{ ok = $true; snapshotId = $snapshotId; isNew = [bool]$meta.isNew; committed = [bool]$committed.committed; reason = [string]$committed.reason }
}

function Update-InputHistoryAfterScan([string]$Language) {
    # Scan-Updates の直後に呼ぶ。現在ハッシュに対応する検知版が無いブックだけを保存する。
    if (-not (Test-InputHistoryEnabled)) { return @() }
    $result = @()
    try {
        $structure = Get-Structure $Language
        foreach ($w in @(Get-Array $structure.workbooks)) {
            $id = [string]$w.workbookId
            $cur = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
            if ([string]::IsNullOrWhiteSpace($cur)) { continue }
            $curSnap = [string](Get-DataProperty $w 'currentSnapshotId' '')
            if (-not [string]::IsNullOrWhiteSpace($curSnap)) {
                $m = Get-SnapshotManifest $Language $id $curSnap
                if ($null -ne $m -and (Normalize-FileHash ([string](Get-DataProperty $m 'sourceHash' ''))) -eq $cur) { continue }
            }
            $r = Capture-DetectedSnapshot $Language $id 'scan'
            if ([bool]$r.ok) { $result += [ordered]@{ workbookId = $id; snapshotId = [string]$r.snapshotId } }
        }
    } catch { }
    return $result
}

function Capture-RenderInput([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$JobId) {
    # V5-§C-2: 検知版が既にあっても、レンダリング予定なら入力ファイルは必ず確保する。
    $paths = Get-Paths
    $structure = Get-Structure $Language
    $wb = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
    if ($wb.Count -eq 0) { throw "Workbookが見つかりません: $WorkbookId" }
    $w = $wb[0]
    $livePath = Join-Safe ([string]$paths.submissionDir) ([string]$w.relativePath)
    $snap = $SnapshotId
    if ([string]::IsNullOrWhiteSpace($snap)) { $snap = [string](Get-DataProperty $w 'currentSnapshotId' '') }

    if (-not (Test-InputHistoryEnabled) -or [string]::IsNullOrWhiteSpace($snap)) {
        # 履歴機能が無効: 従来どおり提出フォルダの現物を直接使う。
        return [ordered]@{ path = ''; snapshotId = ''; ephemeral = $false; hash = '' }
    }
    $m = Get-SnapshotManifest $Language $WorkbookId $snap
    if ($null -eq $m) { return [ordered]@{ path = ''; snapshotId = ''; ephemeral = $false; hash = '' } }
    $hash = Normalize-FileHash ([string](Get-DataProperty $m 'sourceHash' ''))

    if (Test-SourceRetentionEnabled) {
        $state = Get-SnapshotSourceState $Language $WorkbookId $snap
        if ([bool]$state.sourceRetained) {
            $retainedHash = Normalize-FileHash (New-Sha256 ([string]$state.sourcePath))
            if (-not [string]::IsNullOrWhiteSpace($retainedHash) -and $retainedHash -eq $hash) {
                return [ordered]@{ path = [string]$state.sourcePath; snapshotId = $snap; ephemeral = $false; hash = $hash; verified = $true }
            }
            # Never render a corrupted retained source as the requested immutable
            # generation. Mark it unavailable, then attempt recovery from the live file.
            try { Remove-Item -LiteralPath ([string]$state.sourcePath) -Force -ErrorAction SilentlyContinue } catch { }
            Set-SnapshotSourceState $Language $WorkbookId $snap $false 'hash-mismatch'
        }
        # 現物が消えている: 提出フォルダの現物が同じハッシュなら復元する。
        if (Test-Path -LiteralPath $livePath) {
            $liveHash = Normalize-FileHash (New-Sha256 $livePath)
            if ($liveHash -eq $hash) {
                $r = Save-SnapshotSourceFile $Language $WorkbookId $snap $livePath $hash
                if ([bool]$r.ok) { return [ordered]@{ path = [string]$r.path; snapshotId = $snap; ephemeral = $false; hash = $hash; verified = $true } }
            }
        }
    }

    # 縮退モード、または現物を復元できない場合は一時コピーを作る。
    if (-not (Test-Path -LiteralPath $livePath)) { return [ordered]@{ path = ''; snapshotId = $snap; ephemeral = $false; hash = $hash } }
    $captureId = $(if ([string]::IsNullOrWhiteSpace($JobId)) { New-RbId } else { $JobId })
    $dir = Join-Path (Get-EphemeralJobRoot $Language) $captureId
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $extension = [IO.Path]::GetExtension($livePath).ToLowerInvariant()
    if ($extension -notin @('.xlsx','.xlsm','.docx','.pptx','.pdf')) { $extension = '.xlsx' }
    $tmp = Join-Path $dir ("source$extension")
    if (-not (Test-Path -LiteralPath $tmp)) {
        Copy-FileSharedRead $livePath $tmp
        try { Unblock-File -LiteralPath $tmp -ErrorAction SilentlyContinue } catch { }
    }
    $tmpHash = Normalize-FileHash (New-Sha256 $tmp)
    if ($tmpHash -ne $hash) {
        # 提出フォルダの現物は既に別の版になっている。
        # この検知版の入力としては使えない(H1のPDFとしてH2を組んでしまうため)。
        Remove-EphemeralCopy $Language $captureId
        return [ordered]@{ path = ''; snapshotId = ''; ephemeral = $false; hash = '' }
    }
    return [ordered]@{ path = $tmp; snapshotId = $snap; ephemeral = $true; hash = $tmpHash; captureId = $captureId; verified = $true }
}

function Remove-EphemeralCopy([string]$Language, [string]$CaptureId) {
    try {
        if ([string]::IsNullOrWhiteSpace($CaptureId)) { return }
        $dir = Join-Path (Get-EphemeralJobRoot $Language) $CaptureId
        if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    } catch { }
}

function Clear-StaleEphemeralCopies([string]$Language) {
    # V5-§1.3: 非保持承認なのに現物が長期間残らないよう、上限時間で必ず消す。
    try {
        $root = Get-EphemeralJobRoot $Language
        if (-not (Test-Path -LiteralPath $root)) { return }
        $maxAge = [int](Get-InputHistorySettings).ephemeralCopyMaxAgeMinutes
        if ($maxAge -le 0) { $maxAge = 30 }
        $limit = [DateTime]::UtcNow.AddMinutes(-1 * $maxAge)
        foreach ($d in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
            if ($d.LastWriteTimeUtc -lt $limit) {
                Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
                Write-HistoryEvent $Language 'history.cleanup' ([ordered]@{ kind = 'ephemeral-copy'; captureId = [string]$d.Name })
            }
        }
    } catch { }
}


# ---- 掃除 (V5-§3.6 / §3.7) -----------------------------------------

function Get-ProtectedSnapshotIds([string]$Language, [string]$WorkbookId) {
    # V5-§C-3: 保持数の枠とは無関係に常時保護する検知版。
    # H1 -> H2 -> 再H1 では manifest の detectedAt が古いままなので、日時順の「直近N版」では現在版を消しうる。
    $ids = @()
    try {
        $structure = Get-Structure $Language
        $wb = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
        if ($wb.Count -gt 0) {
            $ids += [string](Get-DataProperty $wb[0] 'currentSnapshotId' '')
            $ids += [string](Get-DataProperty $wb[0] 'lastRenderedSnapshotId' '')
        }
    } catch { }
    try {
        $auto = Read-AutoState $Language $WorkbookId
        if ($null -ne $auto) { $ids += [string](Get-DataProperty $auto 'pendingSnapshotId' '') }
    } catch { }
    try {
        $ptr = Get-ComparisonBaselinePointer $Language $WorkbookId
        if ($null -ne $ptr) { $ids += [string](Get-DataProperty $ptr 'snapshotId' '') }
    } catch { }
    return @($ids | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Get-FinalPdfPinAgeDays([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    # 正式PDFで使われた版のうち、最も新しい出力からの経過日数を返す。pin が無ければ -1。
    $dir = Get-SnapshotPinDir $Language $WorkbookId $SnapshotId
    if (-not (Test-Path -LiteralPath $dir)) { return -1 }
    $newest = $null
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter 'final-pdf_*.json' -ErrorAction SilentlyContinue)) {
        if ($null -eq $newest -or $f.LastWriteTimeUtc -gt $newest) { $newest = $f.LastWriteTimeUtc }
    }
    if ($null -eq $newest) { return -1 }
    return ([DateTime]::UtcNow - $newest).TotalDays
}

function Invoke-InputHistoryCleanup([string]$Language) {
    if (-not (Test-InputHistoryEnabled)) { return }
    $lockPath = Join-Path (Get-WorkspacePath $Language) 'locks\history-cleanup.lock'
    $handle = Try-AcquireLockHandle $lockPath
    if ($null -eq $handle) { return }   # 他サーバーが掃除中
    try {
        $cfg = Get-InputHistorySettings
        $root = Get-InputHistoryRoot $Language
        if (-not (Test-Path -LiteralPath $root)) { return }
        foreach ($wbDir in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
            $workbookId = [string]$wbDir.Name
            $protected = @(Get-ProtectedSnapshotIds $Language $workbookId)
            # 未完成世代(manifest なし)を削除
            $removedIncomplete = $false
            foreach ($d in @(Get-ChildItem -LiteralPath $wbDir.FullName -Directory -ErrorAction SilentlyContinue)) {
                if (-not (Test-Path -LiteralPath (Join-Path $d.FullName 'manifest.json'))) {
                    if ($d.LastWriteTimeUtc -lt [DateTime]::UtcNow.AddHours(-6)) {
                        Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        $removedIncomplete = $true
                    }
                }
            }
            if ($removedIncomplete) { Clear-SnapshotRuntimeCaches $Language $workbookId }
            $all = @(Get-SnapshotIds $Language $workbookId)
            if ($all.Count -eq 0) { continue }
            $keepRecent = @($all | Select-Object -Last ([Math]::Max(1, [int]$cfg.retainSourceVersions)))
            foreach ($sn in $all) {
                if ($protected -contains $sn) { continue }
                if ($keepRecent -contains $sn) { continue }
                if ((Get-ActiveLeaseCount $Language $workbookId $sn) -gt 0) { continue }
                $pins = @(Get-SnapshotPins $Language $workbookId $sn)
                if ($pins -contains 'comparison-baseline') { continue }
                if ($pins -contains 'manual') { continue }
                $finalPins = @($pins | Where-Object { $_ -like 'final-pdf_*' })

                # --- 段階1: source.xlsx だけを削除する ---
                $state = Get-SnapshotSourceState $Language $workbookId $sn
                if ([bool]$state.sourceRetained) {
                    $canRemoveSource = $true
                    if ($finalPins.Count -gt 0) {
                        $days = [double](Get-InputHistorySettings).sourceRetentionDaysAfterBuild
                        $configured = (Get-InputHistorySettings).sourceRetentionDaysAfterBuild
                        if ($null -eq $configured) { $canRemoveSource = $false }
                        else {
                            $age = Get-FinalPdfPinAgeDays $Language $workbookId $sn
                            if ($age -lt [double]$configured) { $canRemoveSource = $false }
                        }
                    }
                    if ($canRemoveSource) {
                        Remove-Item -LiteralPath ([string]$state.sourcePath) -Force -ErrorAction SilentlyContinue
                        Set-SnapshotSourceState $Language $workbookId $sn $false 'retention'
                        Clear-SnapshotRuntimeCaches $Language $workbookId $sn
                        Write-HistoryEvent $Language 'input.source.removed' ([ordered]@{ workbookId = $workbookId; snapshotId = $sn; reason = 'retention' })
                    }
                }

                # --- 段階2: pin が1件も無ければフォルダごと削除する ---
                if ($finalPins.Count -eq 0 -and $pins.Count -eq 0) {
                    $m = Get-SnapshotManifest $Language $workbookId $sn
                    if ($null -ne $m -and [string](Get-DataProperty $m 'status' '') -eq 'complete') {
                        Remove-Item -LiteralPath (Get-SnapshotDir $Language $workbookId $sn) -Recurse -Force -ErrorAction SilentlyContinue
                        Clear-SnapshotRuntimeCaches $Language $workbookId $sn
                        Write-HistoryEvent $Language 'history.cleanup' ([ordered]@{ kind = 'snapshot'; workbookId = $workbookId; snapshotId = $sn })
                    }
                }
            }
        }
    } catch {
        # 掃除の失敗はレンダリングの失敗にしない。
        Write-Warning ("履歴の掃除に失敗しました: " + $_.Exception.Message)
    } finally {
        Release-LockHandle $handle
        # V5-P2: 掃除で容量が変わるため、次回の /api/state で実測させる。
        Reset-InputHistorySizeCache
    }
}

function Get-InputHistorySizeMb([string]$Language) {
    # input-history 配下の全再帰列挙は共有ドライブ上で非常に重い。
    # /api/state のポーリングごとに実行せず60秒キャッシュする。
    # dataDir を切り替えた直後に旧ワークスペースの値を返さないよう、ルートパスもキーに含める。
    $root = ''
    try { $root = Get-InputHistoryRoot $Language } catch { }
    $cacheKey = ('{0}|{1}' -f $Language, [string]$root).ToLowerInvariant()
    if ($null -ne $Script:HistorySizeCache -and $Script:HistorySizeCacheKey -eq $cacheKey -and (([DateTime]::UtcNow - $Script:HistorySizeCacheAtUtc).TotalSeconds -lt $Script:HistorySizeCacheSeconds)) {
        return $Script:HistorySizeCache
    }
    $value = 0
    try {
        if (-not [string]::IsNullOrWhiteSpace($root) -and (Test-Path -LiteralPath $root)) {
            $bytes = (Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            if ($null -ne $bytes) { $value = [Math]::Round(($bytes / 1MB), 1) }
        }
    } catch { $value = 0 }
    $Script:HistorySizeCache = $value
    $Script:HistorySizeCacheKey = $cacheKey
    $Script:HistorySizeCacheAtUtc = [DateTime]::UtcNow
    return $value
}

function Reset-InputHistorySizeCache {
    $Script:HistorySizeCache = $null
    $Script:HistorySizeCacheKey = ''
    $Script:HistorySizeCacheAtUtc = [DateTime]::MinValue
}

# ---- content-pdf の保護 (V5-§3.7) ----------------------------------

function Get-ContentPdfVersionDir([string]$Workspace, [string]$WorkbookId, [string]$VersionId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $safeVersionId = Assert-SafeStorageSegment $VersionId 'versionId'
    return (Join-Path $Workspace (Join-Path 'content-pdf' (Join-Path $safeWorkbookId $safeVersionId)))
}

function New-ContentPdfPin([string]$Workspace, [string]$WorkbookId, [string]$VersionId, [string]$PinName, $Data) {
    # V5-P0(#4): 失敗を握りつぶさず $true/$false で返す。呼出元が結果を検査してロールバックできるようにする。
    try {
        $safePinName = Assert-SafeStorageSegment $PinName 'pinName'
        $versionDir = Get-ContentPdfVersionDir $Workspace $WorkbookId $VersionId
        if (-not (Test-Path -LiteralPath $versionDir -PathType Container)) { return $false }
        $dir = Join-Path $versionDir 'pins'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $pinPath = Join-Path $dir ("{0}.json" -f $safePinName)
        Write-JsonFile $pinPath $Data
        return (Test-Path -LiteralPath $pinPath)
    } catch { return $false }
}

function Remove-ContentPdfPin([string]$Workspace, [string]$WorkbookId, [string]$VersionId, [string]$PinName) {
    try {
        if ([string]::IsNullOrWhiteSpace($VersionId) -or [string]::IsNullOrWhiteSpace($PinName)) { return }
        $path = Join-Path (Join-Path (Get-ContentPdfVersionDir $Workspace $WorkbookId $VersionId) 'pins') ("{0}.json" -f $PinName)
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
    } catch { }
}


function New-ContentPdfLease([string]$Workspace, [string]$WorkbookId, [string]$VersionId, [string]$Purpose, [string]$JobId, [int]$MinutesValid = 120) {
    try {
        $safePurpose = Assert-SafeStorageSegment $Purpose 'purpose'
        $safeJobId = Assert-SafeStorageSegment $JobId 'jobId'
        $versionDir = Get-ContentPdfVersionDir $Workspace $WorkbookId $VersionId
        if (-not (Test-Path -LiteralPath $versionDir -PathType Container)) { return '' }
        $dir = Join-Path $versionDir 'leases'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $name = ('{0}_{1}.json' -f $safePurpose, $safeJobId)
        $path = Join-Path $dir $name
        Write-JsonFile $path ([ordered]@{
            jobId = $JobId; purpose = $Purpose; versionId = $VersionId
            createdAt = New-NowIso; heartbeatAt = New-NowIso
            expiresAt = ([DateTime]::UtcNow.AddMinutes($MinutesValid).ToString('o'))
            pcName = $env:COMPUTERNAME
        })
        if (-not (Test-Path -LiteralPath $path)) { return '' }
        return $name
    } catch { return '' }
}

function Remove-ContentPdfLease([string]$Workspace, [string]$WorkbookId, [string]$VersionId, [string]$LeaseName) {
    try {
        if ([string]::IsNullOrWhiteSpace($LeaseName)) { return }
        $path = Join-Path (Join-Path (Get-ContentPdfVersionDir $Workspace $WorkbookId $VersionId) 'leases') $LeaseName
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
    } catch { }
}

function Refresh-LeaseFile([string]$Path, [int]$MinutesValid = 120) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $lease = Read-JsonFile $Path $null
        if ($null -eq $lease) { return $false }
        Set-NoteProperty $lease 'heartbeatAt' (New-NowIso)
        Set-NoteProperty $lease 'expiresAt' ([DateTime]::UtcNow.AddMinutes($MinutesValid).ToString('o'))
        Write-JsonFile $Path $lease
        return $true
    } catch { return $false }
}

function New-DiffJobLeases([string]$Language, $Context, [string]$JobId) {
    $workspace = Get-WorkspacePath $Language
    $historyLock = Join-Path $workspace 'locks\history-cleanup.lock'
    $contentLock = Get-ContentPdfMaintenanceLockPath $workspace ([string]$Context.workbookId)
    return Invoke-WithLock $historyLock {
        Invoke-WithLock $contentLock {
            $leases = @()
            try {
                foreach ($side in @(
                    [ordered]@{ role = 'baseline'; snapshotId = [string]$Context.baselineSnapshotId; versionId = [string]$Context.baselineVersionId },
                    [ordered]@{ role = 'current'; snapshotId = [string]$Context.currentSnapshotId; versionId = [string]$Context.currentVersionId }
                )) {
                    if ($null -eq (Get-SnapshotManifest $Language ([string]$Context.workbookId) ([string]$side.snapshotId))) {
                        throw '比較対象の履歴版が整理されたため、処理を開始できません。'
                    }
                    $availability = Get-HistoryRenderVersionAvailability $Language ([string]$Context.workbookId) ([string]$side.snapshotId) ([string]$side.versionId)
                    if (-not [bool]$availability.ready) { throw [string]$availability.reason }
                    $snapshotLease = New-SnapshotLease $Language ([string]$Context.workbookId) ([string]$side.snapshotId) ('diff-' + [string]$side.role) $JobId 120 ([string]$side.versionId)
                    if ([string]::IsNullOrWhiteSpace($snapshotLease)) { throw '履歴版の保護leaseを作成できませんでした。' }
                    $lease = [ordered]@{
                        role = [string]$side.role
                        snapshotId = [string]$side.snapshotId
                        versionId = [string]$side.versionId
                        snapshotLeaseName = $snapshotLease
                        contentLeaseName = ''
                    }
                    $leases += $lease
                    $contentLease = New-ContentPdfLease $workspace ([string]$Context.workbookId) ([string]$side.versionId) ('diff-' + [string]$side.role) $JobId 120
                    if ([string]::IsNullOrWhiteSpace($contentLease)) { throw 'content PDFの保護leaseを作成できませんでした。' }
                    $lease.contentLeaseName = $contentLease
                }
                return @($leases)
            } catch {
                foreach ($lease in @($leases)) {
                    Remove-SnapshotLease $Language ([string]$Context.workbookId) ([string]$lease.snapshotId) ([string]$lease.snapshotLeaseName)
                    Remove-ContentPdfLease $workspace ([string]$Context.workbookId) ([string]$lease.versionId) ([string]$lease.contentLeaseName)
                }
                throw
            }
        }
    }
}

function Refresh-DiffJobLeases($Job) {
    $language = [string](Get-DataProperty $Job 'mode' $Mode)
    $workbookId = [string](Get-DataProperty $Job 'workbookId' '')
    $workspace = Get-WorkspacePath $language
    foreach ($lease in @(Get-Array (Get-DataProperty $Job 'leases' @()))) {
        $snapshotPath = Join-Path (Get-SnapshotLeaseDir $language $workbookId ([string]$lease.snapshotId)) ([string]$lease.snapshotLeaseName)
        $contentPath = Join-Path (Join-Path (Get-ContentPdfVersionDir $workspace $workbookId ([string]$lease.versionId)) 'leases') ([string]$lease.contentLeaseName)
        if (-not (Refresh-LeaseFile $snapshotPath 120)) { throw '履歴版の保護leaseを更新できませんでした。' }
        if (-not (Refresh-LeaseFile $contentPath 120)) { throw 'content PDFの保護leaseを更新できませんでした。' }
    }
}

function Remove-DiffJobLeases($Job) {
    try {
        $language = [string](Get-DataProperty $Job 'mode' $Mode)
        $workbookId = [string](Get-DataProperty $Job 'workbookId' '')
        $workspace = Get-WorkspacePath $language
        foreach ($lease in @(Get-Array (Get-DataProperty $Job 'leases' @()))) {
            Remove-SnapshotLease $language $workbookId ([string]$lease.snapshotId) ([string]$lease.snapshotLeaseName)
            Remove-ContentPdfLease $workspace $workbookId ([string]$lease.versionId) ([string]$lease.contentLeaseName)
        }
    } catch { }
}

function Test-ContentPdfProtected([string]$Workspace, [string]$WorkbookId, [string]$VersionId) {
    $base = Get-ContentPdfVersionDir $Workspace $WorkbookId $VersionId
    foreach ($sub in @('pins','leases')) {
        $d = Join-Path $base $sub
        if (Test-Path -LiteralPath $d) {
            $files = @(Get-ChildItem -LiteralPath $d -File -Filter '*.json' -ErrorAction SilentlyContinue)
            if ($sub -eq 'pins' -and $files.Count -gt 0) { return $true }
            if ($sub -eq 'leases') { foreach ($f in $files) { if (Test-LeaseActive $f) { return $true } } }
        }
    }
    return $false
}


# =====================================================================
# V5 Stage 3 — Phase 2A: 検知パイプラインと自動スケジューラー
# =====================================================================

function Get-AutoStateDir([string]$Language) {
    return (Join-Path (Get-WorkspacePath $Language) 'state\auto-render')
}
function Read-AutoState([string]$Language, [string]$WorkbookId) {
    try {
        $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
        $p = Join-Path (Get-AutoStateDir $Language) ("{0}.json" -f $safeWorkbookId)
        if (-not (Test-Path -LiteralPath $p)) { return $null }
        return (Read-JsonFile $p $null)
    } catch { return $null }
}
function Write-AutoState([string]$Language, [string]$WorkbookId, $State) {
    try {
        $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
        $dir = Get-AutoStateDir $Language
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Set-NoteProperty $State 'workbookId' $safeWorkbookId
        Set-NoteProperty $State 'updatedAt' (New-NowIso)
        Write-JsonFile (Join-Path $dir ("{0}.json" -f $safeWorkbookId)) $State
    } catch { }
}
function New-AutoState([string]$WorkbookId) {
    return [ordered]@{
        schemaVersion = 2; workbookId = $WorkbookId; updatedAt = New-NowIso
        pendingSnapshotId = ''; pendingHash = ''; stableCount = 0
        firstDetectedAt = ''; lastSeenAt = ''; quietDeadline = ''
        state = 'idle'; deferReason = ''; ownerPcName = ''; ownerJobId = ''
        retryCount = 0; nextRetryAt = ''; lastError = ''; lastResult = ''; lastResultAt = ''; notificationId = ''
    }
}

function Test-InteractiveExcelInUse([string]$SourcePath) {
    # V5-§5.3: 単純な Get-Process EXCEL では止まりすぎる。
    # ReportBinder のレンダリング用 Excel は Visible=$false なのでメインウィンドウを持たない。
    # 対話操作されている Excel と、対象ファイルのロックファイルだけを見る。
    try {
        $procs = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
        if ($procs.Count -gt 0) { return $true }
    } catch { }
    try {
        if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
            $dir = Split-Path -Parent $SourcePath
            $name = [IO.Path]::GetFileName($SourcePath)
            $lock = Join-Path $dir ('~$' + $name)
            if (Test-Path -LiteralPath $lock) { return $true }
        }
    } catch { }
    return $false
}

function Test-RenderEngineBusy([string]$Language) {
    $h = Try-AcquireLockHandle (Get-RenderEngineLockPath $Language)
    if ($null -eq $h) { return $true }
    Release-LockHandle $h
    return $false
}

function Start-AutoRenderJobForWorkbook([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    # V5-P0: 静止確認したのは「その検知版」なので、ジョブ実行時に別の版へすり替わってはいけない。
    # 対象snapshotを明示的に固定してジョブへ渡す。
    try {
        $pins = @{}
        if (-not [string]::IsNullOrWhiteSpace($SnapshotId)) {
            $m = Get-SnapshotManifest $Language $WorkbookId $SnapshotId
            if ($null -ne $m) {
                $pins[$WorkbookId] = [ordered]@{ snapshotId = $SnapshotId; expectedHash = [string](Get-DataProperty $m 'sourceHash' '') }
            }
        }
        $job = Start-RenderJob $Language @($WorkbookId) $false '' $pins
        return [ordered]@{ ok = $true; jobId = [string]$job.jobId }
    } catch {
        return [ordered]@{ ok = $false; message = $_.Exception.Message }
    }
}

function Test-AutoFailureSuperseded($State, [string]$CurrentHash) {
    if ([string](Get-DataProperty $State 'state' '') -ne 'failed') { return $false }
    $current = Normalize-FileHash $CurrentHash
    $failedHash = Normalize-FileHash ([string](Get-DataProperty $State 'pendingHash' ''))
    return (-not [string]::IsNullOrWhiteSpace($current) -and $current -ne $failedHash)
}

function Invoke-AutoSchedulerTick([string]$Language, $OwnedLocks) {
    # V5-§5.5/§A-3: tick 内でスリープしない状態機械。ブックごとに直列で180秒待たない。
    $settings = Get-AutoRenderSettings
    if (-not [bool]$settings.enabled) { return }
    # V5-P0: ブラウザの30秒タイマーに依存しない。スケジューラー自身が検知し、検知と同時に保存する。
    try { [void](Scan-Updates $Language $null $false) } catch { Write-Warning $_.Exception.Message }
    try { [void](Update-InputHistoryAfterScan $Language) } catch { Write-Warning $_.Exception.Message }
    $paths = Get-Paths
    $structure = $null
    try { $structure = Get-Structure $Language } catch { return }
    $now = [DateTime]::UtcNow

    foreach ($w in @(Get-Array $structure.workbooks)) {
        $id = [string]$w.workbookId
        if ([string]::IsNullOrWhiteSpace($id)) { continue }

        # V5-§5.5: ブック単位の所有権。取得できないブックは他サーバーの担当なので黙って飛ばす。
        if (-not $OwnedLocks.ContainsKey($id)) {
            $lockPath = Join-Path (Get-WorkspacePath $Language) ("locks\auto-owner_{0}.lock" -f $id)
            $h = Try-AcquireLockHandle $lockPath
            if ($null -eq $h) { continue }
            $OwnedLocks[$id] = $h
        }

        $state = Read-AutoState $Language $id
        if ($null -eq $state) { $state = New-AutoState $id }

        $sourcePath = Join-Safe ([string]$paths.submissionDir) ([string]$w.relativePath)
        if (-not (Test-Path -LiteralPath $sourcePath)) {
            Set-NoteProperty $state 'state' 'idle'; Set-NoteProperty $state 'deferReason' 'file-missing'
            Write-AutoState $Language $id $state; continue
        }

        if ([string](Get-DataProperty $state 'state' '') -eq 'rendering') {
            $jobId = [string](Get-DataProperty $state 'ownerJobId' '')
            $job = if ([string]::IsNullOrWhiteSpace($jobId)) { $null } else { Read-RenderJobStatus $Language $jobId }
            $jobStatus = [string](Get-DataProperty $job 'status' '')
            if ($jobStatus -in @('queued','running','analyzing')) { continue }
            if ($jobStatus -eq 'completed') {
                Set-NoteProperty $state 'state' 'idle'; Set-NoteProperty $state 'retryCount' 0; Set-NoteProperty $state 'nextRetryAt' ''; Set-NoteProperty $state 'lastError' ''
                Set-NoteProperty $state 'lastResult' 'completed'; Set-NoteProperty $state 'lastResultAt' (New-NowIso); Set-NoteProperty $state 'notificationId' (New-RbId)
                Write-AutoState $Language $id $state; Write-HistoryEvent $Language 'auto.completed' ([ordered]@{ workbookId=$id; jobId=$jobId }); continue
            }
            if ($jobStatus -in @('completed-with-errors','failed')) {
                $retryCount = [int](Get-DataProperty $state 'retryCount' 0) + 1
                $lastError = [string](Get-DataProperty $job 'message' '自動PDF作成に失敗しました。')
                if ([string]::IsNullOrWhiteSpace($lastError)) { $lastError = [string](Get-DataProperty $job 'error' '自動PDF作成に失敗しました。') }
                Set-NoteProperty $state 'retryCount' $retryCount; Set-NoteProperty $state 'lastError' $lastError; Set-NoteProperty $state 'lastResult' 'failed'; Set-NoteProperty $state 'lastResultAt' (New-NowIso); Set-NoteProperty $state 'notificationId' (New-RbId)
                if ($retryCount -le [int]$settings.maxRetryCount) {
                    $delay = [Math]::Min(3600,[int]$settings.retryBaseSeconds * [Math]::Pow(2,[Math]::Max(0,$retryCount-1)))
                    Set-NoteProperty $state 'state' 'retry-wait'; Set-NoteProperty $state 'deferReason' 'retry-backoff'; Set-NoteProperty $state 'nextRetryAt' ($now.AddSeconds($delay).ToString('o'))
                    Write-HistoryEvent $Language 'auto.retry-scheduled' ([ordered]@{ workbookId=$id; jobId=$jobId; retryCount=$retryCount; nextRetryAt=[string]$state.nextRetryAt; error=$lastError })
                } else {
                    Set-NoteProperty $state 'state' 'failed'; Set-NoteProperty $state 'deferReason' 'retry-exhausted'; Set-NoteProperty $state 'nextRetryAt' ''
                    Write-HistoryEvent $Language 'auto.failed' ([ordered]@{ workbookId=$id; jobId=$jobId; retryCount=$retryCount; error=$lastError })
                }
                Write-AutoState $Language $id $state; continue
            }
            Set-NoteProperty $state 'state' 'waiting'; Set-NoteProperty $state 'deferReason' 'job-status-missing'; Write-AutoState $Language $id $state
        }

        if ([string](Get-DataProperty $state 'state' '') -eq 'retry-wait') {
            $nextRetryText = [string](Get-DataProperty $state 'nextRetryAt' '')
            $retryReady = $true; if (-not [string]::IsNullOrWhiteSpace($nextRetryText)) { try { $retryReady = ([DateTime]::Parse($nextRetryText).ToUniversalTime() -le $now) } catch { } }
            if (-not $retryReady) { continue }
            Set-NoteProperty $state 'state' 'waiting'; Set-NoteProperty $state 'deferReason' ''; Set-NoteProperty $state 'quietDeadline' ($now.AddSeconds(-1).ToString('o')); Set-NoteProperty $state 'stableCount' 99
        }

        $currentHash = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
        $renderedHash = Normalize-FileHash ([string](Get-DataProperty $w 'lastRenderedExcelHash' ''))
        # 同じ版で失敗を繰り返さない。ただし利用者が原稿を保存し直した場合は、
        # 新しいハッシュを新規検知として扱い、待機・再試行状態へ自動復帰する。
        if ([string](Get-DataProperty $state 'state' '') -eq 'failed' -and -not (Test-AutoFailureSuperseded $state $currentHash)) { continue }
        if ([string]::IsNullOrWhiteSpace($currentHash) -or $currentHash -eq $renderedHash) {
            if ([string](Get-DataProperty $state 'state' '') -ne 'idle') {
                Set-NoteProperty $state 'state' 'idle'; Set-NoteProperty $state 'pendingSnapshotId' ''
                Set-NoteProperty $state 'pendingHash' ''; Set-NoteProperty $state 'stableCount' 0
                Write-AutoState $Language $id $state
            }
            continue
        }

        $pendingHash = Normalize-FileHash ([string](Get-DataProperty $state 'pendingHash' ''))
        if ($pendingHash -ne $currentHash) {
            # 新しい版を検知。待機をやり直す(古い一時コピーは破棄)。
            $oldCapture = [string](Get-DataProperty $state 'ephemeralCaptureId' '')
            if (-not [string]::IsNullOrWhiteSpace($oldCapture)) { Remove-EphemeralCopy $Language $oldCapture }
            Set-NoteProperty $state 'pendingHash' $currentHash
            Set-NoteProperty $state 'pendingSnapshotId' ([string](Get-DataProperty $w 'currentSnapshotId' ''))
            Set-NoteProperty $state 'stableCount' 1
            Set-NoteProperty $state 'firstDetectedAt' (New-NowIso)
            Set-NoteProperty $state 'quietDeadline' ($now.AddSeconds([int]$settings.quietPeriodSeconds).ToString('o'))
            Set-NoteProperty $state 'state' 'waiting'
            Set-NoteProperty $state 'deferReason' ''
            Set-NoteProperty $state 'ownerPcName' $env:COMPUTERNAME
            Set-NoteProperty $state 'retryCount' 0; Set-NoteProperty $state 'nextRetryAt' ''; Set-NoteProperty $state 'lastError' ''
            Write-AutoState $Language $id $state
            Write-HistoryEvent $Language 'auto.detected' ([ordered]@{ workbookId = $id; hash = $currentHash })
            continue
        }

        Set-NoteProperty $state 'stableCount' ([int](Get-DataProperty $state 'stableCount' 0) + 1)
        Set-NoteProperty $state 'lastSeenAt' (New-NowIso)

        $deadlineText = [string](Get-DataProperty $state 'quietDeadline' '')
        $deadlineReached = $false
        if (-not [string]::IsNullOrWhiteSpace($deadlineText)) {
            try { $deadlineReached = ([DateTime]::Parse($deadlineText).ToUniversalTime() -le $now) } catch { $deadlineReached = $true }
        }
        if (-not $deadlineReached -or [int](Get-DataProperty $state 'stableCount' 0) -lt [int]$settings.requireStableHashCount) {
            Set-NoteProperty $state 'state' 'waiting'; Write-AutoState $Language $id $state; continue
        }
        if ([bool]$settings.deferWhileExcelInUse -and (Test-InteractiveExcelInUse $sourcePath)) {
            Set-NoteProperty $state 'state' 'deferred'; Set-NoteProperty $state 'deferReason' 'excel-in-use'
            Write-AutoState $Language $id $state
            Write-HistoryEvent $Language 'auto.deferred' ([ordered]@{ workbookId = $id; reason = 'excel-in-use' })
            continue
        }
        if (Test-RenderEngineBusy $Language) {
            Set-NoteProperty $state 'state' 'deferred'; Set-NoteProperty $state 'deferReason' 'job-busy'
            Write-AutoState $Language $id $state; continue
        }
        $minimumFreeBytes = [int64]$settings.minFreeMegabytes * 1MB
        $dataFreeBytes = Get-PathFreeBytes ([string]$paths.dataDir); $outputFreeBytes = Get-PathFreeBytes ([string]$paths.outputDir)
        if (($dataFreeBytes -gt 0 -and $dataFreeBytes -lt $minimumFreeBytes) -or ($outputFreeBytes -gt 0 -and $outputFreeBytes -lt $minimumFreeBytes)) {
            Set-NoteProperty $state 'state' 'deferred'; Set-NoteProperty $state 'deferReason' 'disk-low'; Set-NoteProperty $state 'lastError' '空き容量が不足しているため自動PDF作成を保留しています。'
            Write-AutoState $Language $id $state; continue
        }

        Set-NoteProperty $state 'state' 'rendering'; Set-NoteProperty $state 'deferReason' ''
        Write-AutoState $Language $id $state
        $started = Start-AutoRenderJobForWorkbook $Language $id ([string](Get-DataProperty $state 'pendingSnapshotId' ''))
        if ([bool]$started.ok) {
            Set-NoteProperty $state 'ownerJobId' ([string]$started.jobId)
            Write-HistoryEvent $Language 'render.started' ([ordered]@{ workbookId = $id; jobId = [string]$started.jobId; trigger = 'auto' })
        } else {
            $retryCount = [int](Get-DataProperty $state 'retryCount' 0) + 1; Set-NoteProperty $state 'retryCount' $retryCount; Set-NoteProperty $state 'lastError' ([string]$started.message)
            if ($retryCount -le [int]$settings.maxRetryCount) { $delay=[Math]::Min(3600,[int]$settings.retryBaseSeconds*[Math]::Pow(2,[Math]::Max(0,$retryCount-1))); Set-NoteProperty $state 'state' 'retry-wait'; Set-NoteProperty $state 'deferReason' 'start-failed'; Set-NoteProperty $state 'nextRetryAt' ($now.AddSeconds($delay).ToString('o')) }
            else { Set-NoteProperty $state 'state' 'failed'; Set-NoteProperty $state 'deferReason' 'retry-exhausted'; Set-NoteProperty $state 'lastResult' 'failed'; Set-NoteProperty $state 'lastResultAt' (New-NowIso); Set-NoteProperty $state 'notificationId' (New-RbId) }
        }
        Write-AutoState $Language $id $state
    }
}

function Invoke-AutoSchedulerFromFile([string]$ControlPath, [int]$ParentProcessId) {
    # V5-§5.7: 静止待ちは HTTP リスナーの中で行わない。
    # V4 のサーバーは単一スレッドの AcceptTcpClient ループなので、Handle-Api 内で待つと画面が止まる。
    $language = 'ja'
    try {
        $control = Read-JsonFile $ControlPath $null
        if ($null -ne $control) { $language = [string](Get-DataProperty $control 'language' 'ja') }
    } catch { }
    $owned = @{}
    $historyCleanupPending = $true
    try {
        Clear-ExpiredLeases $language
        Clear-StaleEphemeralCopies $language
        Recover-AutoStates $language
        while ($true) {
            # 親PID監視に加え、stopファイルを100ms単位で確認する。
            # 従来の10秒Sleepでは正常終了でも親が子をKillする必要があり、finallyの状態復旧が走らなかった。
            if ($ParentProcessId -le 0 -or $null -eq (Get-Process -Id $ParentProcessId -ErrorAction SilentlyContinue)) { break }
            if (Test-Path -LiteralPath ($ControlPath + '.stop')) { break }
            try { Invoke-AutoSchedulerTick $language $owned } catch { Write-Warning $_.Exception.Message }
            try { Clear-StaleEphemeralCopies $language } catch { }
            # 停止確認は 500ms 間隔。ControlPath は共有ドライブ上にあるため、
            # 100ms 間隔だと利用者ごとに毎秒10回のSMB Test-Path が常時発生する。
            # 親側は WaitForExit(2500) 待つので、500ms でも正常終了と後片付けは間に合う。
            $stopRequested = $false
            for ($i = 0; $i -lt 20; $i++) {
                if ($ParentProcessId -le 0 -or $null -eq (Get-Process -Id $ParentProcessId -ErrorAction SilentlyContinue) -or (Test-Path -LiteralPath ($ControlPath + '.stop'))) {
                    $stopRequested = $true
                    break
                }
                Start-Sleep -Milliseconds 500
            }
            if ($stopRequested) { break }
            # History retention walks every saved workbook/version (including cached
            # raster pages). Run it after the UI has had time to become ready instead
            # of making every application launch wait for the full directory scan.
            if ($historyCleanupPending) {
                try { Invoke-InputHistoryCleanup $language } catch { Write-Warning $_.Exception.Message }
                $historyCleanupPending = $false
            }
        }
    } finally {
        foreach ($k in @($owned.Keys)) { Release-LockHandle $owned[$k] }
        try {
            foreach ($f in @(Get-ChildItem -LiteralPath (Get-AutoStateDir $language) -File -Filter '*.json' -ErrorAction SilentlyContinue)) {
                $st = Read-JsonFile $f.FullName $null
                if ($null -ne $st -and @('rendering','ready') -contains [string](Get-DataProperty $st 'state' '')) {
                    Set-NoteProperty $st 'state' 'waiting'
                    Write-JsonFile $f.FullName $st
                }
            }
        } catch { }
        # 正常終了のたびに制御JSON/.stopを残さない。
        foreach ($path in @($ControlPath, ($ControlPath + '.stop'))) {
            try { if ($path -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } } catch { }
        }
    }
}

function Recover-AutoStates([string]$Language) {
    # V5-§5.5: 起動時リカバリー。
    try {
        $dir = Get-AutoStateDir $Language
        if (-not (Test-Path -LiteralPath $dir)) { return }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue)) {
            $st = Read-JsonFile $f.FullName $null
            if ($null -eq $st) { continue }
            $state = [string](Get-DataProperty $st 'state' '')
            $changed = $false
            if ($state -eq 'rendering') { Set-NoteProperty $st 'state' 'waiting'; $changed = $true }
            elseif ($state -eq 'ready') { Set-NoteProperty $st 'state' 'waiting'; $changed = $true }
            $pending = [string](Get-DataProperty $st 'pendingSnapshotId' '')
            $wbId = [string](Get-DataProperty $st 'workbookId' '')
            if (-not [string]::IsNullOrWhiteSpace($pending) -and -not [string]::IsNullOrWhiteSpace($wbId)) {
                if ($null -eq (Get-SnapshotManifest $Language $wbId $pending)) {
                    Set-NoteProperty $st 'pendingSnapshotId' ''; Set-NoteProperty $st 'pendingHash' ''
                    Set-NoteProperty $st 'stableCount' 0; Set-NoteProperty $st 'state' 'idle'; $changed = $true
                }
            }
            if ($changed) { Write-JsonFile $f.FullName $st }
        }
    } catch { }
}

function Start-AutoSchedulerProcess([string]$Language) {
    $controlPath = ''
    try {
        $settings = Get-AutoRenderSettings
        if (-not [bool]$settings.enabled) { return $null }
        $dir = Join-Path (Get-WorkspacePath $Language) 'state'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        # 制御ファイルはPC名+親PIDでプロセス単位に分離する。
        $controlName = 'auto-scheduler_{0}_{1}.json' -f ([regex]::Replace([string]$env:COMPUTERNAME, '[^A-Za-z0-9_.-]+', '_')), $PID
        $controlPath = Join-Path $dir $controlName
        Write-JsonFile $controlPath ([ordered]@{ schemaVersion = 1; language = $Language; startedAt = New-NowIso; parentPid = $PID; pcName = [string]$env:COMPUTERNAME })
        if (Test-Path -LiteralPath ($controlPath + '.stop')) { Remove-Item -LiteralPath ($controlPath + '.stop') -Force -ErrorAction SilentlyContinue }
        # AppRoot / 制御ファイルのパスに空白が含まれても子プロセスが起動できるよう明示的に引用する。
        $serverScript = Join-Path $Script:AppRoot 'server.ps1'
        $psi = @('-NoProfile','-ExecutionPolicy','Bypass','-File', ('"{0}"' -f $serverScript),
                 '-Mode', $Language, '-AutoSchedulerPath', ('"{0}"' -f $controlPath), '-ParentProcessId', [string]$PID)
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $psi -WindowStyle Hidden -PassThru
        $Script:AutoSchedulerProcess = $proc
        $Script:AutoSchedulerProcessId = $proc.Id
        $Script:AutoSchedulerControlPath = $controlPath
        return $proc
    } catch {
        foreach ($path in @($controlPath, ($controlPath + '.stop'))) {
            try { if ($path -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } } catch { }
        }
        Write-Warning ('自動スケジューラーを起動できませんでした: ' + $_.Exception.Message)
        return $null
    }
}

function Test-AutoSchedulerProcessRunning {
    try {
        if ($null -ne $Script:AutoSchedulerProcess) {
            $Script:AutoSchedulerProcess.Refresh()
            return (-not $Script:AutoSchedulerProcess.HasExited)
        }
        if ($Script:AutoSchedulerProcessId -gt 0) {
            return ($null -ne (Get-Process -Id $Script:AutoSchedulerProcessId -ErrorAction SilentlyContinue))
        }
    } catch { }
    return $false
}

function Stop-AutoSchedulerProcess {
    $controlPath = [string]$Script:AutoSchedulerControlPath
    try {
        if ($controlPath) {
            Set-Content -LiteralPath ($controlPath + '.stop') -Value 'stop' -Encoding ASCII -ErrorAction SilentlyContinue
        }
        $p = $Script:AutoSchedulerProcess
        if ($null -eq $p -and $Script:AutoSchedulerProcessId -gt 0) {
            $p = Get-Process -Id $Script:AutoSchedulerProcessId -ErrorAction SilentlyContinue
        }
        if ($null -ne $p) {
            # 子側は100ms単位でstopを確認するため、まず正常終了とfinallyの後片付けを待つ。
            $exited = $false
            try { $exited = $p.WaitForExit(2500) } catch { }
            if (-not $exited) {
                try { $p.Kill() } catch { }
                try { [void]$p.WaitForExit(1000) } catch { }
            }
        }
    } catch { }
    finally {
        foreach ($path in @($controlPath, ($controlPath + '.stop'))) {
            try { if ($path -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } } catch { }
        }
        $Script:AutoSchedulerProcess = $null
        $Script:AutoSchedulerProcessId = 0
        $Script:AutoSchedulerControlPath = ''
    }
}


# =====================================================================
# V5 Stage 4 — Phase 2B: 画像ハッシュと比較
# =====================================================================

$Script:VisualHashProfileVersion = 3
$Script:VisualHashPageDistanceLimit = 0.075
$Script:VisualHashAverageDistanceLimit = 0.040
$Script:VisualHashDpi = 120

function Get-RenderRecordDir([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId) {
    return (Join-Path (Get-SnapshotDir $Language $WorkbookId $SnapshotId) (Join-Path 'renders' $VersionId))
}

function Get-RenderRasterSheetDir([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId, [string]$SheetName) {
    $recordDir = Get-RenderRecordDir $Language $WorkbookId $SnapshotId $VersionId
    return (Join-Path $recordDir (Join-Path 'raster-v1' (Get-DiffSheetKey $SheetName)))
}

function Get-VisualHashProfile {
    $pdfBoxVersion = ''
    try {
        $vf = Join-Path $Script:AppRoot 'lib\pdfbox\PDFBOX_VERSION.txt'
        if (Test-Path -LiteralPath $vf) { $pdfBoxVersion = ((Get-Content -LiteralPath $vf -Raw) -replace '\s+', ' ').Trim() }
    } catch { }
    return [ordered]@{
        profileVersion = $Script:VisualHashProfileVersion
        pdfBoxVersion = $pdfBoxVersion
        dpi = $Script:VisualHashDpi
        colorMode = 'RGB'
    }
}

function Get-HexHammingRatio([string]$Left, [string]$Right) {
    $a = ([string]$Left).Trim().ToUpperInvariant()
    $b = ([string]$Right).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($a) -or $a.Length -ne $b.Length) { return 1.0 }
    $differentBits = 0
    for ($i = 0; $i -lt $a.Length; $i++) {
        try {
            $xor = ([Convert]::ToInt32($a[$i].ToString(), 16) -bxor [Convert]::ToInt32($b[$i].ToString(), 16))
        } catch { return 1.0 }
        while ($xor -gt 0) {
            $differentBits += ($xor -band 1)
            $xor = $xor -shr 1
        }
    }
    return ([double]$differentBits / [Math]::Max(1, $a.Length * 4))
}

function Test-SheetVisualEquivalent($Before, $After) {
    if ($null -eq $Before -or $null -eq $After) { return $false }
    if ((Get-IntDataProperty $Before 'pageCount' -1) -ne (Get-IntDataProperty $After 'pageCount' -2)) { return $false }
    $beforeText = Normalize-FileHash ([string](Get-DataProperty $Before 'textHash' ''))
    $afterText = Normalize-FileHash ([string](Get-DataProperty $After 'textHash' ''))
    $hasComparableText = -not [string]::IsNullOrWhiteSpace($beforeText) -and -not [string]::IsNullOrWhiteSpace($afterText)
    if ($hasComparableText -and $beforeText -ne $afterText) { return $false }

    $beforePages = @(Get-Array (Get-DataProperty $Before 'pagePerceptualHashes' @()))
    $afterPages = @(Get-Array (Get-DataProperty $After 'pagePerceptualHashes' @()))
    $beforeExactPages = @(Get-Array (Get-DataProperty $Before 'pageHashes' @()))
    $afterExactPages = @(Get-Array (Get-DataProperty $After 'pageHashes' @()))
    # Current analyzer records normalized RGB pixels. When both generations have
    # those hashes, they are authoritative: the old 32x32 grayscale perceptual hash
    # can hide a cell fill, pale color, border color, or small drawing resize.
    if ($beforeExactPages.Count -gt 0 -and $beforeExactPages.Count -eq $afterExactPages.Count) {
        for ($i = 0; $i -lt $beforeExactPages.Count; $i++) {
            if ((Normalize-FileHash ([string]$beforeExactPages[$i])) -ne
                (Normalize-FileHash ([string]$afterExactPages[$i]))) { return $false }
        }
        return $true
    }
    # Legacy records without normalized page hashes retain the perceptual fallback.
    if ($beforePages.Count -eq 0 -or $beforePages.Count -ne $afterPages.Count) { return $false }
    $total = 0.0
    $pageLimit = $(if ($hasComparableText) { $Script:VisualHashPageDistanceLimit } else { 0.035 })
    $averageLimit = $(if ($hasComparableText) { $Script:VisualHashAverageDistanceLimit } else { 0.020 })
    for ($i = 0; $i -lt $beforePages.Count; $i++) {
        $distance = Get-HexHammingRatio ([string]$beforePages[$i]) ([string]$afterPages[$i])
        if ($distance -gt $pageLimit) { return $false }
        $total += $distance
    }
    return (($total / [Math]::Max(1, $beforePages.Count)) -le $averageLimit)
}

function Test-PdfPageAnalyzerAvailable {
    if ($null -ne $Script:PdfPageAnalyzerAvailable) { return [bool]$Script:PdfPageAnalyzerAvailable }
    $Script:PdfPageAnalyzerAvailable = $false
    try {
        $jar = Join-Path $Script:AppRoot 'lib\pdfbox\ReportPdfComposer.jar'
        if (-not (Test-Path -LiteralPath $jar)) { return $false }
        Add-Type -AssemblyName 'System.IO.Compression.FileSystem' -ErrorAction SilentlyContinue
        $zip = [IO.Compression.ZipFile]::OpenRead($jar)
        try { $Script:PdfPageAnalyzerAvailable = @($zip.Entries | Where-Object { $_.FullName -eq 'PdfPageAnalyzer.class' }).Count -gt 0 }
        finally { $zip.Dispose() }
    } catch { $Script:PdfPageAnalyzerAvailable = $false }
    return [bool]$Script:PdfPageAnalyzerAvailable
}

function Invoke-PdfPageAnalyzer([hashtable[]]$Sheets) {
    # ブックごとに Java を1回だけ起動して全シートを解析する。
    if ($null -eq $Sheets -or $Sheets.Count -eq 0) { return $null }
    $tool = $null
    try { $tool = Get-PdfBatchToolInfo } catch { }
    $javaExe = ''
    $cp = ''
    try {
        $javaExe = Resolve-JavaExe
        $cp = (Join-Path $Script:AppRoot 'lib\pdfbox\ReportPdfComposer.jar') + ';' + (Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar')
    } catch { return $null }
    # V5-P0: JAR に PdfPageAnalyzer.class が入っていない配布物では、解析だけを黙って諦める。
    # (build.ps1 を実行して JAR を再生成すると有効になる)
    if (-not (Test-PdfPageAnalyzerAvailable)) { return [ordered]@{ ok = $false; message = 'PdfPageAnalyzer が JAR に含まれていません。app\lib\pdfbox\build.ps1 を実行してください。' } }
    $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('rb-analyze-' + (New-RbId))
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
    try {
        $req = Join-Path $tmpDir 'request.json'
        $res = Join-Path $tmpDir 'result.json'
        Write-JsonFile $req ([ordered]@{ sheets = @($Sheets | ForEach-Object {
            [ordered]@{
                sheetName = [string]$_.sheetName
                pdf = [string]$_.pdf
                rasterDirectory = [string](Get-DataProperty $_ 'rasterDirectory' '')
            }
        }) })
        $run = Invoke-NativeCapture $javaExe @('-Djava.awt.headless=true', '-cp', $cp, 'PdfPageAnalyzer', '--input', $req, '--output', $res, '--dpi', [string]$Script:VisualHashDpi) $Script:PdfAnalyzeTimeoutSeconds
        if ($run.timedOut) {
            # 解析はPDF作成のクリティカルパス外。失敗として返すが、記録には理由を残す。
            return [ordered]@{ ok = $false; message = ('見た目の解析が{0}分以内に終わりませんでした。' -f [int]($Script:PdfAnalyzeTimeoutSeconds / 60)) }
        }
        if ([int]$run.exitCode -ne 0 -or -not (Test-Path -LiteralPath $res)) {
            return [ordered]@{ ok = $false; message = ("exit=" + [string]$run.exitCode + "`n" + [string]$run.text) }
        }
        $parsed = Read-JsonFile $res $null
        return [ordered]@{ ok = $true; result = $parsed }
    } catch {
        return [ordered]@{ ok = $false; message = $_.Exception.Message }
    } finally {
        if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Write-RenderRecord([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId, [string]$Purpose, [bool]$ContentPdfRetained, $Analysis, [string]$SheetSelectionMode = 'all-visible') {
    # V5-IV-3: レンダリング結果は renders\<versionId>\ に世代ごとに置き、書いたら変更しない。
    if ([string]::IsNullOrWhiteSpace($SnapshotId) -or [string]::IsNullOrWhiteSpace($VersionId)) { return }
    try {
        $dir = Get-RenderRecordDir $Language $WorkbookId $SnapshotId $VersionId
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $manifestPath = Join-Path $dir 'render-manifest.json'
        if (-not (Test-FileExistsCompat $manifestPath)) {
            Write-JsonFile $manifestPath ([ordered]@{
                schemaVersion = 1
                purpose = $Purpose
                # Immutable render record: this says whether content PDF was produced for this render.
                # Current retention is determined from the actual content-pdf directory, never from this manifest.
                contentPdfProduced = $ContentPdfRetained
                contentPdfRetained = $ContentPdfRetained # legacy compatibility; do not use as current-state truth
                sourceSnapshotId = $SnapshotId
                workbookId = $WorkbookId
                versionId = $VersionId
                createdAt = New-NowIso
                status = 'complete'
                renderEnvironmentFingerprint = [string]$Script:CurrentRenderEnvFingerprint
                renderEnvironment = $Script:CurrentRenderEnvInfo
                excelPrintProfileVersion = $Script:ExcelPrintProfileVersion
                sheetSelectionMode = Normalize-ExcelSheetSelection $SheetSelectionMode
                visualHashProfile = (Get-VisualHashProfile)
            })
        }
        $hashPath = Join-Path $dir 'visual-hashes.json'
        if ($null -ne $Analysis -and -not (Test-FileExistsCompat $hashPath)) {
            Write-JsonFile $hashPath ([ordered]@{
                schemaVersion = 1
                snapshotId = $SnapshotId
                versionId = $VersionId
                renderEnvironmentFingerprint = [string]$Script:CurrentRenderEnvFingerprint
                visualHashProfile = (Get-VisualHashProfile)
                analyzerVersion = [int](Get-DataProperty $Analysis 'analyzerVersion' 1)
                javaVersion = [string](Get-DataProperty $Analysis 'javaVersion' '')
                javaVendor = [string](Get-DataProperty $Analysis 'javaVendor' '')
                sheets = @(Get-Array (Get-DataProperty $Analysis 'sheets' @()))
            })
        }
        # 履歴画面は重い世代走査を避けるため要約をキャッシュしている。
        # render-manifest / visual-hashes は履歴ルート直下を更新しないため、
        # ここで破棄しないと「画像ハッシュなし」の作成前判定が残り続ける。
        Clear-SnapshotSummaryCache $Language $WorkbookId
    } catch { }
}

function Get-SnapshotRenderSheetSelection([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId = '') {
    # A historical comparison must reproduce the sheet set used by that
    # historical render, not today's source setting. Older records have no
    # field and therefore retain the legacy all-visible behavior.
    try {
        if (-not [string]::IsNullOrWhiteSpace($VersionId)) {
            $exactManifest = Join-Path (Get-RenderRecordDir $Language $WorkbookId $SnapshotId $VersionId) 'render-manifest.json'
            if (-not (Test-Path -LiteralPath $exactManifest -PathType Leaf)) { return 'all-visible' }
            $exact = Read-JsonFile $exactManifest $null
            if ($null -eq $exact) { return 'all-visible' }
            $exactSelection = [string](Get-DataProperty $exact 'sheetSelectionMode' '')
            if ([string]::IsNullOrWhiteSpace($exactSelection)) { return 'all-visible' }
            return Normalize-ExcelSheetSelection $exactSelection
        }
        $rendersRoot = Join-Path (Get-SnapshotDir $Language $WorkbookId $SnapshotId) 'renders'
        if (-not (Test-Path -LiteralPath $rendersRoot -PathType Container)) { return 'all-visible' }
        $records = @(Get-ChildItem -LiteralPath $rendersRoot -File -Filter 'render-manifest.json' -Recurse -Depth 1 -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc -Descending)
        foreach ($record in $records) {
            $manifest = Read-JsonFile ([string]$record.FullName) $null
            if ($null -eq $manifest) { continue }
            $selection = [string](Get-DataProperty $manifest 'sheetSelectionMode' '')
            if (-not [string]::IsNullOrWhiteSpace($selection)) { return Normalize-ExcelSheetSelection $selection }
        }
    } catch { }
    return 'all-visible'
}

function Get-VisualHashes([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $safeVersionId = Assert-SafeStorageSegment $VersionId 'versionId'
    $cacheKey = ($Language + '|' + $safeWorkbookId + '|' + $safeSnapshotId + '|' + $safeVersionId).ToLowerInvariant()
    $p = Join-Path (Get-RenderRecordDir $Language $safeWorkbookId $safeSnapshotId $safeVersionId) 'visual-hashes.json'
    if ($Script:VisualHashCache.ContainsKey($cacheKey)) {
        if (Test-FileExistsCompat $p) { return $Script:VisualHashCache[$cacheKey] }
        [void]$Script:VisualHashCache.Remove($cacheKey)
    }
    if (-not (Test-FileExistsCompat $p)) { return $null }
    try {
        $hashes = Read-JsonFile $p $null
        if ($null -ne $hashes) { $Script:VisualHashCache[$cacheKey] = $hashes }
        return $hashes
    } catch { return $null }
}

function Get-RenderVersionIds([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    $dir = Join-Path (Get-SnapshotDir $Language $WorkbookId $SnapshotId) 'renders'
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    return @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object { [string]$_.Name })
}

# ---- baseline ポインタ (V5-§6.6) -----------------------------------

function Get-ComparisonBaselinePointerPath([string]$Language, [string]$WorkbookId) {
    return (Join-Path (Get-WorkbookHistoryDir $Language $WorkbookId) 'comparison-baseline.json')
}
function Get-ComparisonBaselinePointer([string]$Language, [string]$WorkbookId) {
    $p = Get-ComparisonBaselinePointerPath $Language $WorkbookId
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    try { return (Read-JsonFile $p $null) } catch { return $null }
}
function Set-ComparisonBaseline([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId, [string]$EnvFingerprint) {
    # 切替順序: 新snapshot/content pin -> ポインタ置換 -> 旧pin削除。途中停止時は保護過多側に倒す。
    $workspace = Get-WorkspacePath $Language
    $historyLock = Join-Path $workspace 'locks\history-cleanup.lock'
    $contentLock = Get-ContentPdfMaintenanceLockPath $workspace $WorkbookId
    Invoke-WithLock $historyLock {
        Invoke-WithLock $contentLock {
            if ($null -eq (Get-SnapshotManifest $Language $WorkbookId $SnapshotId)) { throw '比較基準の履歴版が見つかりません。' }
            $availability = Get-HistoryRenderVersionAvailability $Language $WorkbookId $SnapshotId $VersionId
            if (-not [bool]$availability.ready) { throw ('比較基準を保護できません: ' + [string]$availability.reason) }
            $old = Get-ComparisonBaselinePointer $Language $WorkbookId
            $oldSnapshot = [string](Get-DataProperty $old 'snapshotId' '')
            $oldVersion = [string](Get-DataProperty $old 'versionId' '')
            $pinData = [ordered]@{
                snapshotId = $SnapshotId; versionId = $VersionId
                renderEnvironmentFingerprint = $EnvFingerprint
                visualHashProfileVersion = $Script:VisualHashProfileVersion
                pinnedAt = New-NowIso
            }
            if (-not (New-SnapshotPin $Language $WorkbookId $SnapshotId 'comparison-baseline' $pinData)) {
                throw '比較基準の履歴版を保護できませんでした。'
            }
            if (-not (New-ContentPdfPin $workspace $WorkbookId $VersionId 'comparison-baseline' $pinData)) {
                Remove-SnapshotPin $Language $WorkbookId $SnapshotId 'comparison-baseline'
                throw '比較基準のcontent PDFを保護できませんでした。'
            }
            Write-JsonFile (Get-ComparisonBaselinePointerPath $Language $WorkbookId) ([ordered]@{
                schemaVersion = 2; snapshotId = $SnapshotId; versionId = $VersionId
                renderEnvironmentFingerprint = $EnvFingerprint; updatedAt = New-NowIso
            })
            if (-not [string]::IsNullOrWhiteSpace($oldSnapshot) -and $oldSnapshot -ne $SnapshotId) {
                Remove-SnapshotPin $Language $WorkbookId $oldSnapshot 'comparison-baseline'
            }
            if (-not [string]::IsNullOrWhiteSpace($oldVersion) -and $oldVersion -ne $VersionId) {
                Remove-ContentPdfPin $workspace $WorkbookId $oldVersion 'comparison-baseline'
            }
        }
    } | Out-Null
}

function Get-LatestComparisonAssetPointerPath([string]$Language, [string]$WorkbookId) {
    return (Join-Path (Get-WorkbookHistoryDir $Language $WorkbookId) 'latest-comparison-assets.json')
}

function Get-LatestComparisonAssetPointer([string]$Language, [string]$WorkbookId) {
    $path = Get-LatestComparisonAssetPointerPath $Language $WorkbookId
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try { return (Read-JsonFile $path $null) } catch { return $null }
}

function Set-LatestComparisonAssets([string]$Language, [string]$WorkbookId, $Comparison) {
    # 変更バッジが示す直近の自動比較は、次回比較基準の切替とは別に両版を保護する。
    # 固定role名を別々に使うため、旧currentが新baselineになる場合も安全に切り替えられる。
    if ($null -eq $Comparison -or [string](Get-DataProperty $Comparison 'scope' '') -ne 'automatic' -or
        [string](Get-DataProperty $Comparison 'status' '') -ne 'complete') { throw '直近比較として保護できる自動比較結果がありません。' }
    foreach ($field in @('baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
        if ([string]::IsNullOrWhiteSpace([string](Get-DataProperty $Comparison $field ''))) { throw "直近比較の識別子が不足しています: $field" }
    }
    $workspace = Get-WorkspacePath $Language
    $historyLock = Join-Path $workspace 'locks\history-cleanup.lock'
    $contentLock = Get-ContentPdfMaintenanceLockPath $workspace $WorkbookId
    Invoke-WithLock $historyLock {
        Invoke-WithLock $contentLock {
            $baselineAvailability = Get-HistoryRenderVersionAvailability $Language $WorkbookId ([string]$Comparison.baselineSnapshotId) ([string]$Comparison.baselineVersionId)
            $currentAvailability = Get-HistoryRenderVersionAvailability $Language $WorkbookId ([string]$Comparison.currentSnapshotId) ([string]$Comparison.currentVersionId)
            if (-not [bool]$baselineAvailability.ready -or -not [bool]$currentAvailability.ready) {
                throw '直近比較の画像ハッシュと同一世代のcontent PDFを保護できません。'
            }
            $old = Get-LatestComparisonAssetPointer $Language $WorkbookId
            $pinData = [ordered]@{
                scope = 'automatic'
                baselineSnapshotId = [string]$Comparison.baselineSnapshotId
                baselineVersionId = [string]$Comparison.baselineVersionId
                currentSnapshotId = [string]$Comparison.currentSnapshotId
                currentVersionId = [string]$Comparison.currentVersionId
                comparedAt = [string](Get-DataProperty $Comparison 'comparedAt' '')
                pinnedAt = New-NowIso
            }
            $specs = @(
                [ordered]@{ role = 'baseline'; pinName = 'latest-comparison-baseline'; snapshotId = [string]$Comparison.baselineSnapshotId; versionId = [string]$Comparison.baselineVersionId },
                [ordered]@{ role = 'current'; pinName = 'latest-comparison-current'; snapshotId = [string]$Comparison.currentSnapshotId; versionId = [string]$Comparison.currentVersionId }
            )
            # 途中失敗時は削除せず保護過多側へ倒す。旧pointer/pinも残るため比較資産は失われない。
            foreach ($spec in $specs) {
                $data = [ordered]@{}
                foreach ($key in @($pinData.Keys)) { $data[$key] = $pinData[$key] }
                $data.role = [string]$spec.role
                if (-not (New-SnapshotPin $Language $WorkbookId ([string]$spec.snapshotId) ([string]$spec.pinName) $data)) {
                    throw '直近比較の履歴版を保護できませんでした。'
                }
                if (-not (New-ContentPdfPin $workspace $WorkbookId ([string]$spec.versionId) ([string]$spec.pinName) $data)) {
                    throw '直近比較のcontent PDFを保護できませんでした。'
                }
            }
            $pointerPath = Get-LatestComparisonAssetPointerPath $Language $WorkbookId
            Write-JsonFile $pointerPath ([ordered]@{
                schemaVersion = 1
                baselineSnapshotId = [string]$Comparison.baselineSnapshotId
                baselineVersionId = [string]$Comparison.baselineVersionId
                currentSnapshotId = [string]$Comparison.currentSnapshotId
                currentVersionId = [string]$Comparison.currentVersionId
                comparedAt = [string](Get-DataProperty $Comparison 'comparedAt' '')
                updatedAt = New-NowIso
            })
            $savedPointer = Read-JsonFile $pointerPath $null
            if ($null -eq $savedPointer) { throw '直近比較の保護ポインタを保存後に再読込できませんでした。' }
            foreach ($field in @('baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
                if ([string](Get-DataProperty $savedPointer $field '') -ne [string](Get-DataProperty $Comparison $field '')) {
                    throw "直近比較の保護ポインタ検証に失敗しました: $field"
                }
            }
            foreach ($oldSpec in @(
                [ordered]@{ pinName = 'latest-comparison-baseline'; snapshotId = [string](Get-DataProperty $old 'baselineSnapshotId' ''); versionId = [string](Get-DataProperty $old 'baselineVersionId' ''); newSnapshotId = [string]$Comparison.baselineSnapshotId; newVersionId = [string]$Comparison.baselineVersionId },
                [ordered]@{ pinName = 'latest-comparison-current'; snapshotId = [string](Get-DataProperty $old 'currentSnapshotId' ''); versionId = [string](Get-DataProperty $old 'currentVersionId' ''); newSnapshotId = [string]$Comparison.currentSnapshotId; newVersionId = [string]$Comparison.currentVersionId }
            )) {
                if (-not [string]::IsNullOrWhiteSpace([string]$oldSpec.snapshotId) -and [string]$oldSpec.snapshotId -ne [string]$oldSpec.newSnapshotId) {
                    Remove-SnapshotPin $Language $WorkbookId ([string]$oldSpec.snapshotId) ([string]$oldSpec.pinName)
                }
                if (-not [string]::IsNullOrWhiteSpace([string]$oldSpec.versionId) -and [string]$oldSpec.versionId -ne [string]$oldSpec.newVersionId) {
                    Remove-ContentPdfPin $workspace $WorkbookId ([string]$oldSpec.versionId) ([string]$oldSpec.pinName)
                }
            }
        }
    } | Out-Null
}

# ---- 比較専用レンダリング (V5-§6.7) --------------------------------


function Render-SnapshotForComparison([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$BaselineVersionId = '') {
    # structure.json等は変更しないが、画像ハッシュと同一世代のcontent PDFは保持する。
    # これにより再レンダリング比較でも、判定対象と画面表示対象が必ず一致する。
    try {
        $structure = Get-Structure $Language
        $source = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
        if ($source.Count -gt 0) {
            $sourceType = [string](Get-DataProperty $source[0] 'sourceType' 'excel')
            if ($sourceType -eq 'pdf') { return Render-PdfSnapshotForComparison $Language $WorkbookId $SnapshotId }
            if ($sourceType -eq 'word') { return Render-WordSnapshotForComparison $Language $WorkbookId $SnapshotId }
            if ($sourceType -eq 'powerpoint') { return Render-PowerPointSnapshotForComparison $Language $WorkbookId $SnapshotId }
        }
    } catch { }
    $state = Get-SnapshotSourceState $Language $WorkbookId $SnapshotId
    if (-not [bool]$state.sourceRetained) { return [ordered]@{ ok = $false; reason = 'source-missing' } }
    $comparisonSheetSelection = Get-SnapshotRenderSheetSelection $Language $WorkbookId $SnapshotId $BaselineVersionId
    return Invoke-WithRenderLock $Language $WorkbookId {
        $versionId = New-RbVersionId
        $tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('rb-cmp-' + (New-RbId))
        $workspace = Get-WorkspacePath $Language
        $contentDir = Get-ContentPdfVersionDir $workspace $WorkbookId $versionId
        New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
        New-Item -ItemType Directory -Path $contentDir -Force | Out-Null
        $leaseJobId = New-RbId
        $snapshotLease = ''; $contentLease = ''
        try {
            $historyLock = Join-Path $workspace 'locks\history-cleanup.lock'
            $contentLock = Get-ContentPdfMaintenanceLockPath $workspace $WorkbookId
            $leaseResult = Invoke-WithLock $historyLock {
                Invoke-WithLock $contentLock {
                    if ($null -eq (Get-SnapshotManifest $Language $WorkbookId $SnapshotId)) { throw '比較元の履歴版が整理されました。' }
                    $newSnapshotLease = New-SnapshotLease $Language $WorkbookId $SnapshotId 'compare' $leaseJobId 120 $versionId
                    if ([string]::IsNullOrWhiteSpace($newSnapshotLease)) { throw '履歴版の保護leaseを作成できませんでした。' }
                    $newContentLease = New-ContentPdfLease $workspace $WorkbookId $versionId 'compare' $leaseJobId 120
                    if ([string]::IsNullOrWhiteSpace($newContentLease)) {
                        Remove-SnapshotLease $Language $WorkbookId $SnapshotId $newSnapshotLease
                        throw 'content PDFの保護leaseを作成できませんでした。'
                    }
                    return [pscustomobject][ordered]@{ snapshotLease = $newSnapshotLease; contentLease = $newContentLease }
                }
            }
            $snapshotLease = [string]$leaseResult.snapshotLease
            $contentLease = [string]$leaseResult.contentLease
        } catch {
            Remove-SnapshotLease $Language $WorkbookId $SnapshotId $snapshotLease
            Remove-ContentPdfLease $workspace $WorkbookId $versionId $contentLease
            Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue
            return [ordered]@{ ok = $false; reason = 'lease-failed'; message = $_.Exception.Message }
        }
        $excel = $null; $book = $null; $success = $false
        try {
            $workExtension = [IO.Path]::GetExtension([string]$state.sourcePath).ToLowerInvariant(); if ($workExtension -notin @('.xlsx','.xlsm')) { $workExtension = '.xlsx' }
            $work = Join-Path $tmpDir ("source$workExtension")
            Copy-FileSharedRead ([string]$state.sourcePath) $work
            try { Unblock-File -LiteralPath $work -ErrorAction SilentlyContinue } catch { }
            $comparisonPackagePreparation = $null
            try { $comparisonPackagePreparation = Prepare-XlsxPrintPackage $work } catch { $comparisonPackagePreparation = [ordered]@{ ok=$false; printSettingsPrepared=$false } }
            $comparisonPackagePrepared = [bool](Get-DataProperty $comparisonPackagePreparation 'printSettingsPrepared' $false)
            $excel = New-ExcelApplicationForRender
            $envInfo = Get-RenderEnvironment $excel
            $Script:CurrentRenderEnvFingerprint = Get-RenderEnvironmentFingerprint $envInfo
            $Script:CurrentRenderEnvInfo = $envInfo
            $book = Open-ExcelWorkbookSafe $excel $work $true
            $sheets = @()
            $sheetCount = 0
            try { $sheetCount = [int]$book.Worksheets.Count } catch { $sheetCount = 0 }
            for ($i = 1; $i -le $sheetCount; $i++) {
                $ws = $null
                try {
                    $ws = $book.Worksheets.Item($i)
                    $sheetName = [string]$ws.Name
                    if ([int]$ws.Visible -ne -1) { continue }
                    if ($comparisonSheetSelection -eq 'numeric-only' -and -not (Test-StrictNumericSheetName $sheetName)) { continue }
                    $outPdf = Join-Path $contentDir ("{0}.pdf" -f (Get-WorksheetStorageStem $sheetName))
                    [void](Export-WorksheetToPdfSafe $excel $book $ws $outPdf $sheetName $comparisonPackagePrepared)
                    if (Test-Path -LiteralPath $outPdf) {
                        $sheets += @{
                            sheetName = $sheetName
                            pdf = $outPdf
                            rasterDirectory = (Get-RenderRasterSheetDir $Language $WorkbookId $SnapshotId $versionId $sheetName)
                        }
                    }
                } catch {
                } finally { Invoke-ComRelease $ws }
            }
            if ($sheets.Count -eq 0) { return [ordered]@{ ok = $false; reason = 'no-sheets' } }
            $analysis = Invoke-PdfPageAnalyzer $sheets
            if ($null -eq $analysis -or -not [bool]$analysis.ok) { return [ordered]@{ ok = $false; reason = 'analyze-failed' } }
            Write-RenderRecord $Language $WorkbookId $SnapshotId $versionId 'comparison' $true ([pscustomobject]$analysis.result) $comparisonSheetSelection
            $availability = Get-HistoryRenderVersionAvailability $Language $WorkbookId $SnapshotId $versionId
            if (-not [bool]$availability.ready) { return [ordered]@{ ok = $false; reason = 'retention-verify-failed'; message = [string]$availability.reason } }
            $success = $true
            return [ordered]@{ ok = $true; versionId = $versionId; envFingerprint = [string]$Script:CurrentRenderEnvFingerprint }
        } catch {
            return [ordered]@{ ok = $false; reason = 'error'; message = $_.Exception.Message }
        } finally {
            if ($book) { try { $book.Close($false) } catch { } ; Invoke-ComRelease $book }
            if ($excel) { Close-ExcelApplicationForRender $excel }
            Remove-SnapshotLease $Language $WorkbookId $SnapshotId $snapshotLease
            Remove-ContentPdfLease $workspace $WorkbookId $versionId $contentLease
            if (-not $success -and (Test-Path -LiteralPath $contentDir)) { Remove-Item -LiteralPath $contentDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }
}

function Get-ComparisonSourceType([string]$Language, [string]$WorkbookId) {
    try {
        $context = Get-RegisteredSourceAdapterContext $Language $WorkbookId
        $sourceType = ([string](Get-DataProperty $context 'sourceType' 'excel')).Trim().ToLowerInvariant()
        if ($sourceType -in @('excel','word','pdf','powerpoint')) { return $sourceType }
    } catch { }
    return 'excel'
}

function Test-ComparisonSheetEquivalent($Before, $After) {
    if ($null -eq $Before -or $null -eq $After) { return $false }
    if ([string](Get-DataProperty $Before 'status' '') -ne 'ok' -or [string](Get-DataProperty $After 'status' '') -ne 'ok') { return $false }
    $beforeHash = Normalize-FileHash ([string](Get-DataProperty $Before 'sheetVisualHash' ''))
    $afterHash = Normalize-FileHash ([string](Get-DataProperty $After 'sheetVisualHash' ''))
    return ((-not [string]::IsNullOrWhiteSpace($beforeHash) -and $beforeHash -eq $afterHash) -or (Test-SheetVisualEquivalent $Before $After))
}

function New-ComparisonUnitMapping($Before, $After, [string]$Kind, [double]$Confidence, [string]$Method, [string]$Message = '') {
    $beforeName = $(if ($null -ne $Before) { [string](Get-DataProperty $Before 'sheetName' '') } else { '' })
    $afterName = $(if ($null -ne $After) { [string](Get-DataProperty $After 'sheetName' '') } else { '' })
    return [pscustomobject][ordered]@{
        beforeSheetName = $beforeName
        afterSheetName = $afterName
        displayName = $(if (-not [string]::IsNullOrWhiteSpace($afterName)) { $afterName } else { $beforeName })
        kind = $Kind
        matchConfidence = [Math]::Max(0, [Math]::Min(1, $Confidence))
        matchMethod = $Method
        message = $Message
    }
}

function Get-ComparisonSheetSequenceNumber($Sheet) {
    $name = [string](Get-DataProperty $Sheet 'sheetName' '')
    if ($name -match '(?<number>\d+)\s*$') { return [int]$matches['number'] }
    return Get-SheetOrderNumber $name
}

function Get-ComparisonUnitMappings($BaselineSheets, $CurrentSheets, [string]$SourceType = 'excel') {
    $before = @(Get-Array $BaselineSheets)
    $after = @(Get-Array $CurrentSheets)
    if ($SourceType -eq 'excel') {
        $beforeMap = @{}
        foreach ($sheet in $before) {
            $name = [string](Get-DataProperty $sheet 'sheetName' '')
            if (-not [string]::IsNullOrWhiteSpace($name)) { $beforeMap[$name] = $sheet }
        }
        $seen = @{}; $result = @()
        foreach ($sheet in $after) {
            $name = [string](Get-DataProperty $sheet 'sheetName' '')
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $seen[$name] = $true
            if (-not $beforeMap.ContainsKey($name)) { $result += New-ComparisonUnitMapping $null $sheet 'added' 1 'name-unmatched'; continue }
            $old = $beforeMap[$name]
            if ([string](Get-DataProperty $old 'status' '') -ne 'ok' -or [string](Get-DataProperty $sheet 'status' '') -ne 'ok') {
                $result += New-ComparisonUnitMapping $old $sheet 'unknown' 0 'analysis-unavailable' '比較用画像を解析できないため判定不能です。'
            } elseif (Test-ComparisonSheetEquivalent $old $sheet) {
                $result += New-ComparisonUnitMapping $old $sheet 'unchanged' 1 'stable-name-and-hash'
            } else {
                $result += New-ComparisonUnitMapping $old $sheet 'modified' 1 'stable-name'
            }
        }
        foreach ($name in @($beforeMap.Keys)) {
            if (-not $seen.ContainsKey([string]$name)) { $result += New-ComparisonUnitMapping $beforeMap[$name] $null 'removed' 1 'name-unmatched' }
        }
        return @($result)
    }

    # Word/PDFは物理ページ名が位置に依存する。完全一致ページの最長共通部分列を
    # アンカーにし、途中への挿入・削除で後続ページを誤って「変更」にしない。
    $before = @($before | Sort-Object @{Expression={Get-ComparisonSheetSequenceNumber $_};Ascending=$true})
    $after = @($after | Sort-Object @{Expression={Get-ComparisonSheetSequenceNumber $_};Ascending=$true})
    $m = $before.Count; $n = $after.Count
    if ($m -eq 0) { return @($after | ForEach-Object { New-ComparisonUnitMapping $null $_ 'added' 1 'sequence-inserted' }) }
    if ($n -eq 0) { return @($before | ForEach-Object { New-ComparisonUnitMapping $_ $null 'removed' 1 'sequence-removed' }) }
    $equivalent = New-Object 'bool[,]' $m,$n
    $hashCountsBefore = @{}; $hashCountsAfter = @{}
    foreach ($sheet in $before) { $h = Normalize-FileHash ([string](Get-DataProperty $sheet 'sheetVisualHash' '')); if ($h) { $hashCountsBefore[$h] = 1 + [int](Get-DataProperty $hashCountsBefore $h 0) } }
    foreach ($sheet in $after) { $h = Normalize-FileHash ([string](Get-DataProperty $sheet 'sheetVisualHash' '')); if ($h) { $hashCountsAfter[$h] = 1 + [int](Get-DataProperty $hashCountsAfter $h 0) } }
    for ($i = 0; $i -lt $m; $i++) { for ($j = 0; $j -lt $n; $j++) { $equivalent[$i,$j] = Test-ComparisonSheetEquivalent ($before[$i]) ($after[$j]) } }
    $lcs = New-Object 'int[,]' ($m + 1),($n + 1)
    for ($i = $m - 1; $i -ge 0; $i--) {
        for ($j = $n - 1; $j -ge 0; $j--) {
            $nextI = $i + 1; $nextJ = $j + 1
            if ($equivalent[$i,$j]) { $lcs[$i,$j] = 1 + $lcs[$nextI,$nextJ] }
            else {
                $skipBefore = [int]$lcs[$nextI,$j]; $skipAfter = [int]$lcs[$i,$nextJ]
                $lcs[$i,$j] = [Math]::Max($skipBefore, $skipAfter)
            }
        }
    }
    $anchors = @(); $i = 0; $j = 0
    while ($i -lt $m -and $j -lt $n) {
        $nextI = $i + 1; $nextJ = $j + 1
        if ($equivalent[$i,$j] -and $lcs[$i,$j] -eq (1 + $lcs[$nextI,$nextJ])) { $anchors += [pscustomobject]@{ before=$i; after=$j }; $i++; $j++ }
        else {
            $skipBefore = [int]$lcs[$nextI,$j]; $skipAfter = [int]$lcs[$i,$nextJ]
            if ($skipBefore -ge $skipAfter) { $i++ } else { $j++ }
        }
    }
    $anchors += [pscustomobject]@{ before=$m; after=$n }
    $result = @(); $beforeStart = 0; $afterStart = 0
    foreach ($anchor in $anchors) {
        $beforeEnd = [int]$anchor.before; $afterEnd = [int]$anchor.after
        $beforeGap = $beforeEnd - $beforeStart; $afterGap = $afterEnd - $afterStart
        if ($beforeGap -eq 0) {
            for ($k = 0; $k -lt $afterGap; $k++) { $result += New-ComparisonUnitMapping $null ($after[$afterStart + $k]) 'added' 1 'sequence-inserted' }
        } elseif ($afterGap -eq 0) {
            for ($k = 0; $k -lt $beforeGap; $k++) { $result += New-ComparisonUnitMapping ($before[$beforeStart + $k]) $null 'removed' 1 'sequence-removed' }
        } elseif ($beforeGap -eq $afterGap) {
            for ($k = 0; $k -lt $beforeGap; $k++) {
                $old = $before[$beforeStart + $k]; $cur = $after[$afterStart + $k]
                if ([string](Get-DataProperty $old 'status' '') -ne 'ok' -or [string](Get-DataProperty $cur 'status' '') -ne 'ok') {
                    $result += New-ComparisonUnitMapping $old $cur 'unknown' 0 'analysis-unavailable' '比較用画像を解析できないため判定不能です。'
                } else {
                    $confidence = $(if ($beforeGap -eq 1) { 0.95 } else { 0.8 })
                    $result += New-ComparisonUnitMapping $old $cur 'modified' $confidence 'sequence-between-anchors' '前後の一致ページを基準に対応付けました。'
                }
            }
        } else {
            $pairs = [Math]::Min($beforeGap, $afterGap)
            for ($k = 0; $k -lt $pairs; $k++) {
                $result += New-ComparisonUnitMapping ($before[$beforeStart + $k]) ($after[$afterStart + $k]) 'unknown' 0.35 'ambiguous-sequence' 'ページの追加・削除と変更が同じ区間にあり、対応を確定できません。'
            }
            for ($k = $pairs; $k -lt $beforeGap; $k++) { $result += New-ComparisonUnitMapping ($before[$beforeStart + $k]) $null 'removed' 0.65 'ambiguous-sequence-unmatched' }
            for ($k = $pairs; $k -lt $afterGap; $k++) { $result += New-ComparisonUnitMapping $null ($after[$afterStart + $k]) 'added' 0.65 'ambiguous-sequence-unmatched' }
        }
        if ($beforeEnd -lt $m -and $afterEnd -lt $n) {
            $old = $before[$beforeEnd]; $cur = $after[$afterEnd]
            $hash = Normalize-FileHash ([string](Get-DataProperty $old 'sheetVisualHash' ''))
            $ambiguousHash = (-not [string]::IsNullOrWhiteSpace($hash) -and ([int](Get-DataProperty $hashCountsBefore $hash 0) -gt 1 -or [int](Get-DataProperty $hashCountsAfter $hash 0) -gt 1))
            $result += New-ComparisonUnitMapping $old $cur 'unchanged' $(if ($ambiguousHash) { 0.85 } else { 1.0 }) $(if ($ambiguousHash) { 'exact-hash-sequence-duplicate' } else { 'exact-hash-sequence' })
        }
        $beforeStart = $beforeEnd + 1; $afterStart = $afterEnd + 1
    }
    return @($result)
}

function Set-ComparisonUnitMappingResult($Result, $Mappings) {
    $list = @(Get-Array $Mappings)
    $Result.unitMappings = @($list)
    $Result.changedSheets = @($list | Where-Object { [string]$_.kind -eq 'modified' } | ForEach-Object { [string]$_.displayName })
    $Result.unchangedSheets = @($list | Where-Object { [string]$_.kind -eq 'unchanged' } | ForEach-Object { [string]$_.displayName })
    $Result.unknownSheets = @($list | Where-Object { [string]$_.kind -eq 'unknown' } | ForEach-Object { [string]$_.displayName })
    $Result.addedSheets = @($list | Where-Object { [string]$_.kind -eq 'added' } | ForEach-Object { [string]$_.displayName })
    $Result.removedSheets = @($list | Where-Object { [string]$_.kind -eq 'removed' } | ForEach-Object { [string]$_.displayName })
    $confidences = @($list | ForEach-Object { [double](Get-DataProperty $_ 'matchConfidence' 1) })
    $Result.mappingConfidence = $(if ($confidences.Count -gt 0) { [double](@($confidences | Measure-Object -Minimum)[0].Minimum) } else { 1.0 })
}

function Compare-SnapshotVisual([string]$Language, [string]$WorkbookId, [string]$CurrentSnapshotId, [string]$CurrentVersionId) {
    # 表現は「見落としなし」ではなく「提出用PDFの見た目を基準とした高精度な判定」。
    $result = [ordered]@{
        schemaVersion = 3; status = 'unavailable'
        scope = 'automatic'
        baselineSnapshotId = ''; baselineVersionId = ''
        currentSnapshotId = $CurrentSnapshotId; currentVersionId = $CurrentVersionId
        comparedAt = New-NowIso; method = ''
        sourceType = Get-ComparisonSourceType $Language $WorkbookId
        unitMappings = @(); mappingConfidence = 1.0
        changedSheets = @(); unchangedSheets = @(); unknownSheets = @()
        addedSheets = @(); removedSheets = @(); message = ''
    }
    $cur = Get-VisualHashes $Language $WorkbookId $CurrentSnapshotId $CurrentVersionId
    if ($null -eq $cur) { $result.message = '今回版の画像ハッシュがありません。'; return $result }
    $curEnv = [string](Get-DataProperty $cur 'renderEnvironmentFingerprint' '')

    $ptr = Get-ComparisonBaselinePointer $Language $WorkbookId
    $baseSnap = [string](Get-DataProperty $ptr 'snapshotId' '')
    $baseVer = [string](Get-DataProperty $ptr 'versionId' '')
    # The baseline pointer is authoritative. Falling back to manifest.previousSnapshotId
    # is unsafe after registration recovery or hash de-duplication: the "previous"
    # snapshot can be months old even though the user just recreated a PDF.
    if ([string]::IsNullOrWhiteSpace($baseSnap) -or $baseSnap -eq $CurrentSnapshotId) {
        $result.message = '前回の比較基準がありません。今回版を新しい基準にします。'
        return $result
    }

    # 自動比較も履歴比較と同じく、判定ハッシュと表示PDFの世代一致を必須にする。
    $currentAvailability = Get-HistoryRenderVersionAvailability $Language $WorkbookId $CurrentSnapshotId $CurrentVersionId
    if (-not [bool]$currentAvailability.ready) {
        $result.message = '今回版の画像ハッシュと同一世代のcontent PDFがそろっていないため、自動比較できません。'
        return $result
    }
    $base = $null
    if (-not [string]::IsNullOrWhiteSpace($baseVer)) { $base = Get-VisualHashes $Language $WorkbookId $baseSnap $baseVer }
    $baseAvailability = $null
    if ($null -ne $base -and -not [string]::IsNullOrWhiteSpace($baseVer)) {
        $baseAvailability = Get-HistoryRenderVersionAvailability $Language $WorkbookId $baseSnap $baseVer
    }
    $curProfile = Get-DataProperty $cur 'visualHashProfile' $null
    $baseProfile = $(if ($null -ne $base) { Get-DataProperty $base 'visualHashProfile' $null } else { $null })
    $environmentChanged = ($null -ne $base -and (
        [string](Get-DataProperty $base 'renderEnvironmentFingerprint' '') -ne $curEnv -or
        (Get-IntDataProperty $base 'analyzerVersion' 0) -ne (Get-IntDataProperty $cur 'analyzerVersion' 0) -or
        (Get-IntDataProperty $baseProfile 'profileVersion' 0) -ne (Get-IntDataProperty $curProfile 'profileVersion' 0)))
    $assetsMissing = ($null -eq $baseAvailability -or -not [bool]$baseAvailability.ready)
    $method = 'stored-hash'
    if ($null -eq $base -or $environmentChanged -or $assetsMissing) {
        # 環境差・ハッシュ欠落・同一世代PDF欠落のいずれでも、保存済みExcelから一組を再生成する。
        $re = Render-SnapshotForComparison $Language $WorkbookId $baseSnap $baseVer
        if ([bool]$re.ok) {
            $baseVer = [string]$re.versionId
            $base = Get-VisualHashes $Language $WorkbookId $baseSnap $baseVer
            $baseAvailability = Get-HistoryRenderVersionAvailability $Language $WorkbookId $baseSnap $baseVer
            $method = 're-rendered'
        } else {
            $result.message = '前回版の比較資産を同一世代で確保できないため、自動比較できません。今回版を新しい比較基準とします。'
            return $result
        }
    }
    if ($null -eq $base -or $null -eq $baseAvailability -or -not [bool]$baseAvailability.ready) {
        $result.message = '前回版の画像ハッシュと同一世代のcontent PDFを確認できません。'
        return $result
    }

    $mappings = Get-ComparisonUnitMappings (Get-DataProperty $base 'sheets' @()) (Get-DataProperty $cur 'sheets' @()) ([string]$result.sourceType)
    $result.status = 'complete'
    $result.baselineSnapshotId = $baseSnap
    $result.baselineVersionId = $baseVer
    $result.method = $method
    Set-ComparisonUnitMappingResult $result $mappings
    $dir = Join-Path (Get-RenderRecordDir $Language $WorkbookId $CurrentSnapshotId $CurrentVersionId) 'comparisons'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $comparisonKey = (Get-Sha256Text ("{0}|{1}|{2}|{3}|automatic" -f $baseSnap, $baseVer, $CurrentSnapshotId, $CurrentVersionId)).Substring(7, 16)
    $comparisonPath = Join-Path $dir ("cmp-{0}.json" -f $comparisonKey)
    Write-JsonFile $comparisonPath $result
    $saved = Read-JsonFile $comparisonPath $null
    if ($null -eq $saved) { throw '自動比較結果を保存後に再読込できませんでした。' }
    foreach ($field in @('scope','baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
        if ([string](Get-DataProperty $saved $field '') -ne [string](Get-DataProperty $result $field '')) {
            throw "自動比較結果の保存検証に失敗しました: $field"
        }
    }
    try {
        Set-LatestComparisonAssets $Language $WorkbookId $saved
    } catch {
        # 保護できない比較を最新変更バッジへ公開しない。
        Remove-Item -LiteralPath $comparisonPath -Force -ErrorAction SilentlyContinue
        throw
    }
    try { Publish-LatestComparisonCaches $Language $WorkbookId $saved $cur $base } catch {
        Write-Warning ('比較情報のローカルキャッシュ作成に失敗しました: ' + $_.Exception.Message)
    }
    Write-HistoryEvent $Language 'compare.completed' ([ordered]@{ workbookId = $WorkbookId; snapshotId = $CurrentSnapshotId; changed = @($result.changedSheets); unknown = @($result.unknownSheets) })
    return $saved
}

function Invoke-PostRenderAnalysis([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId, $Rendered, [string]$SheetSelectionMode = 'all-visible') {
    # V5-§6.5: PDF作成のクリティカルパスの外。失敗しても PDF 作成は成功扱いのまま。
    if (-not (Test-InputHistoryEnabled)) { return $null }
    if ([string]::IsNullOrWhiteSpace($SnapshotId)) { return $null }
    try {
        $sheets = @()
        foreach ($r in @(Get-Array $Rendered)) {
            $pdf = [string](Get-DataProperty $r 'pdf' '')
            $name = [string](Get-DataProperty $r 'sheetName' '')
            # Render-Workbook already validated every output PDF before queuing this analysis.
            # Avoid one extra SMB stat per sheet on the comparison preparation path.
            if ($pdf -and $name) {
                $sheets += @{
                    sheetName = $name
                    pdf = $pdf
                    rasterDirectory = (Get-RenderRasterSheetDir $Language $WorkbookId $SnapshotId $VersionId $name)
                }
            }
        }
        if ($sheets.Count -eq 0) { return $null }
        $analysis = Invoke-PdfPageAnalyzer $sheets
        $parsed = $null
        if ($null -ne $analysis -and [bool]$analysis.ok) { $parsed = [pscustomobject]$analysis.result }
        Write-RenderRecord $Language $WorkbookId $SnapshotId $VersionId 'normal' $true $parsed $SheetSelectionMode
        if ($null -eq $parsed) { return $null }
        $cmp = Compare-SnapshotVisual $Language $WorkbookId $SnapshotId $VersionId
        # complete/unavailableのどちらでも解析結果を版の記録へ残す。
        # 非同期解析の競合や環境差で比較できない場合に、理由を後から確認できる。
        try {
            $analysisRecordDir = Get-RenderRecordDir $Language $WorkbookId $SnapshotId $VersionId
            Write-JsonFile (Join-Path $analysisRecordDir 'comparison-analysis.json') $cmp
        } catch { }
        # V5-§6.6: baseline を更新するのは今回の解析に成功したときだけ。
        # unknown 版を基準にすると次回の比較元が失われる。
        $hasOk = @(Get-Array (Get-DataProperty $parsed 'sheets' @()) | Where-Object { [string]$_.status -eq 'ok' }).Count -gt 0
        if ($hasOk) { Set-ComparisonBaseline $Language $WorkbookId $SnapshotId $VersionId ([string]$Script:CurrentRenderEnvFingerprint) }
        # 比較・baseline保護まで終わった最終状態を、complete/unavailable に
        # かかわらず履歴一覧へ反映する。作成途中の要約を再利用しない。
        [void](Update-SnapshotSummaryCacheEntry $Language $WorkbookId $SnapshotId)
        return $cmp
    } catch {
        Clear-SnapshotSummaryCache $Language $WorkbookId
        Write-Warning ('画像ハッシュの解析に失敗しました: ' + $_.Exception.Message)
        return $null
    }
}

function Get-LatestComparison([string]$Language, [string]$WorkbookId, $Workbook = $null) {
    try {
        # V5-P2: 呼出元が既に structure を読んでいる場合は再読込しない。
        # /api/state はブック1件ごとにここへ来るため、90件なら structure.json を90回読んでいた。
        $target = $Workbook
        if ($null -eq $target) {
            $structure = Get-Structure $Language
            $wb = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
            if ($wb.Count -eq 0) { return $null }
            $target = $wb[0]
        }
        $snap = [string](Get-DataProperty $target 'lastRenderedSnapshotId' '')
        $ver = [string](Get-DataProperty $target 'lastRenderedVersionId' '')
        if ([string]::IsNullOrWhiteSpace($snap) -or [string]::IsNullOrWhiteSpace($ver)) { return $null }
        $cacheKey = ($Language + '|' + $WorkbookId + '|' + $snap + '|' + $ver).ToLowerInvariant()
        if ($Script:LatestComparisonCache.ContainsKey($cacheKey)) {
            return $Script:LatestComparisonCache[$cacheKey]
        }
        # 既存データも比較フォルダーを列挙せず、レンダリング時の確定ファイルを直接読む。
        $recordDir = Get-RenderRecordDir $Language $WorkbookId $snap $ver
        $analysisPath = Join-Path $recordDir 'comparison-analysis.json'
        if (Test-Path -LiteralPath $analysisPath) {
            try {
                $candidate = Read-JsonFile $analysisPath $null
                $scope = [string](Get-DataProperty $candidate 'scope' '')
                $candidateSnapshot = [string](Get-DataProperty $candidate 'currentSnapshotId' '')
                $candidateVersion = [string](Get-DataProperty $candidate 'currentVersionId' '')
                if ($null -ne $candidate -and ([string]::IsNullOrWhiteSpace($scope) -or $scope -eq 'automatic') -and
                    ([string]::IsNullOrWhiteSpace($candidateSnapshot) -or $candidateSnapshot -eq $snap) -and
                    ([string]::IsNullOrWhiteSpace($candidateVersion) -or $candidateVersion -eq $ver)) {
                    $Script:LatestComparisonCache[$cacheKey] = $candidate
                    return $candidate
                }
            } catch { }
        }
        $dir = Join-Path $recordDir 'comparisons'
        if (-not (Test-Path -LiteralPath $dir)) { return $null }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending)) {
            try {
                $candidate = Read-JsonFile $f.FullName $null
                if ($null -eq $candidate) { continue }
                $scope = [string](Get-DataProperty $candidate 'scope' '')
                # scope無しは旧版互換。履歴画面から作った任意比較は、
                # 最新版の変更バッジや自動比較基準として扱わない。
                if (-not ([string]::IsNullOrWhiteSpace($scope) -or $scope -eq 'automatic')) { continue }
                # 新形式は現在版のsnapshot/versionも完全一致させる。旧形式で項目が無い場合だけ、
                # 物理的に現在版renderフォルダ内にあることを根拠に互換読込する。
                $candidateSnapshot = [string](Get-DataProperty $candidate 'currentSnapshotId' '')
                $candidateVersion = [string](Get-DataProperty $candidate 'currentVersionId' '')
                if (-not [string]::IsNullOrWhiteSpace($candidateSnapshot) -and $candidateSnapshot -ne $snap) { continue }
                if (-not [string]::IsNullOrWhiteSpace($candidateVersion) -and $candidateVersion -ne $ver) { continue }
                if (-not $Script:LatestComparisonCache.ContainsKey($cacheKey) -and
                    $Script:LatestComparisonCache.Count -ge $Script:LatestComparisonCacheLimit) {
                    $oldestKey = @($Script:LatestComparisonCache.Keys)[0]
                    if ($null -ne $oldestKey) { [void]$Script:LatestComparisonCache.Remove($oldestKey) }
                }
                # snapshot/version付きの自動比較は不変なので、stateで読んだ結果を詳細表示へ引き継ぐ。
                $Script:LatestComparisonCache[$cacheKey] = $candidate
                return $candidate
            } catch { }
        }
        return $null
    } catch { return $null }
}

# =====================================================================
# V5 Stage 5 — Phase 1B / 2C: アーカイブ・レイアウト履歴・出力トランザクション
# =====================================================================

function Get-ArchiveRoot([string]$Language) { return (Join-Path (Get-WorkspacePath $Language) 'exports\archive') }
function Get-LayoutHistoryDir([string]$Language, [string]$StorageId) {
    $safeStorageId = Assert-SafeStorageSegment $StorageId 'packId'
    return (Join-Path (Get-WorkspacePath $Language) (Join-Path 'layout-history' $safeStorageId))
}
function Get-LayoutScopeInfo($Structure, [string]$PackIdOrCategory, [bool]$AllowArchived = $false) {
    $scope = Resolve-DocumentPackScope $Structure $PackIdOrCategory $AllowArchived
    $storageId = if ([string]::IsNullOrWhiteSpace([string]$scope.category)) { [string]$scope.packId } else { [string]$scope.category }
    return [pscustomobject][ordered]@{
        packId = [string]$scope.packId
        category = [string]$scope.category
        storageId = (Assert-SafeStorageSegment $storageId 'packId')
        pack = $scope.pack
        builtIn = [bool]$scope.builtIn
    }
}
function Invoke-LayoutSnapshotRetention([string]$Directory, [int]$Keep = 100) {
    if (-not (Test-Path -LiteralPath $Directory)) { return }
    $limit = [Math]::Max(1, $Keep)
    foreach ($old in @(Get-ChildItem -LiteralPath $Directory -File -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -Skip $limit)) {
        Remove-Item -LiteralPath $old.FullName -Force -ErrorAction SilentlyContinue
    }
}
function Get-FinalTransactionDir([string]$Language) { return (Join-Path (Get-WorkspacePath $Language) 'state\final-transactions') }
function Get-FinalTransactionBackupDir([string]$Language, [string]$TransactionId) { return (Join-Path (Get-WorkspacePath $Language) (Join-Path 'state\final-backups' $TransactionId)) }
function Remove-FinalTransactionBackupDir([string]$Language, [string]$TransactionId) {
    if ([string]::IsNullOrWhiteSpace($TransactionId)) { return }
    try {
        $dir = Get-FinalTransactionBackupDir $Language $TransactionId
        if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue }
    } catch { }
}

# ---- レイアウト投影スナップショットと限定復元 (V5-§4.2) -------------

function ConvertTo-NormalizedLayoutSnapshotPage($Page, [string[]]$AllowedVolumes) {
    $pageId = [string](Get-DataProperty $Page 'pageId' '')
    if ([string]::IsNullOrWhiteSpace($pageId)) { throw [ArgumentException]::new('ページ構成履歴にpageIdがありません。') }
    $storedVolume = [string](Get-DataProperty $Page 'volume' 'none')
    $volumeInvalid = ($AllowedVolumes -notcontains $storedVolume)
    $volume = $(if ($volumeInvalid) { 'none' } else { $storedVolume })
    $numberingMode = [string](Get-DataProperty $Page 'numberingMode' 'visible')
    if ($numberingMode -notin @('none','visible')) { throw [ArgumentException]::new("ページ構成履歴のnumberingModeが不正です: $pageId") }
    try { $order = [double](Get-DataProperty $Page 'order' 0) } catch { throw [ArgumentException]::new("ページ構成履歴のorderが不正です: $pageId") }
    if ([double]::IsNaN($order) -or [double]::IsInfinity($order)) { throw [ArgumentException]::new("ページ構成履歴のorderが不正です: $pageId") }
    $range = ConvertTo-NormalizedPageRange (Get-DataProperty $Page 'pageRange' $null)
    return [pscustomobject][ordered]@{
        pageId = $pageId
        title = [string](Get-DataProperty $Page 'title' '')
        volume = $volume
        storedVolume = $storedVolume
        volumeInvalid = $volumeInvalid
        enabled = ($volume -ne 'none')
        order = $order
        orderManual = [bool](Get-DataProperty $Page 'orderManual' $false)
        numberingMode = $numberingMode
        numberingManual = [bool](Get-DataProperty $Page 'numberingManual' $false)
        pageRange = $range
    }
}

function Get-NormalizedLayoutSnapshotPages($Snapshot, [string[]]$AllowedVolumes) {
    $pages = @()
    $seen = @{}
    foreach ($page in @(Get-Array (Get-DataProperty $Snapshot 'pages' @()))) {
        $normalized = ConvertTo-NormalizedLayoutSnapshotPage $page $AllowedVolumes
        $pageId = [string]$normalized.pageId
        if ($seen.ContainsKey($pageId)) { throw [ArgumentException]::new("ページ構成履歴に重複したpageIdがあります: $pageId") }
        $seen[$pageId] = $true
        $pages += $normalized
    }
    return @($pages)
}

function Save-LayoutSnapshot([string]$Language, [string]$PackIdOrCategory, [string]$Reason, $Structure = $null) {
    if (-not (Test-InputHistoryEnabled)) { return '' }
    try {
        $st = $Structure
        if ($null -eq $st) { $st = Get-Structure $Language }
        $scope = Get-LayoutScopeInfo $st $PackIdOrCategory $false
        $pages = @()
        foreach ($p in @(Get-Array $st.pages)) {
            $wb = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq [string]$p.workbookId } | Select-Object -First 1)
            if ($wb.Count -eq 0) { continue }
            if (-not (Test-WorkbookPack $wb[0] ([string]$scope.packId))) { continue }
            # V5-§4.2: レイアウト項目だけを保存する。structure 全体は保存しない。
            $pages += [ordered]@{
                pageId = (Resolve-PageId $p)
                title = [string]$p.title
                volume = [string]$p.volume
                enabled = [bool]$p.enabled
                order = [double]$p.order
                orderManual = [bool](Get-DataProperty $p 'orderManual' $false)
                numberingMode = [string](Get-DataProperty $p 'numberingMode' 'visible')
                numberingManual = [bool](Get-DataProperty $p 'numberingManual' $false)
                pageRange = Get-DataProperty $p 'pageRange' $null
            }
        }
        if ($pages.Count -eq 0) { return '' }
        $dir = Get-LayoutHistoryDir $Language ([string]$scope.storageId)
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $id = New-RbId
        Write-JsonFile (Join-Path $dir ("{0}.json" -f $id)) ([ordered]@{
            schemaVersion = 2; snapshotId = $id; language = $Language; packId = [string]$scope.packId; category = [string]$scope.category
            createdAt = New-NowIso; reason = $Reason; pages = $pages
        })
        Invoke-LayoutSnapshotRetention $dir 100
        Write-HistoryEvent $Language 'layout.changed' ([ordered]@{ packId = [string]$scope.packId; category = [string]$scope.category; reason = $Reason; snapshotId = $id; pageCount = $pages.Count })
        return $id
    } catch { return '' }
}

function Get-LayoutSnapshots([string]$Language, [string]$PackIdOrCategory) {
    $structure = Get-Structure $Language
    $scope = Get-LayoutScopeInfo $structure $PackIdOrCategory $true
    $dir = Get-LayoutHistoryDir $Language ([string]$scope.storageId)
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    $out = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 100)) {
        try {
            $j = Read-JsonFile $f.FullName $null
            $out += [ordered]@{
                snapshotId = [string](Get-DataProperty $j 'snapshotId' $f.BaseName)
                createdAt = [string](Get-DataProperty $j 'createdAt' '')
                reason = [string](Get-DataProperty $j 'reason' '')
                pageCount = @(Get-Array (Get-DataProperty $j 'pages' @())).Count
                packId = [string]$scope.packId
                category = [string]$scope.category
            }
        } catch { }
    }
    return $out
}

function Read-LayoutSnapshot([string]$Language, [string]$PackIdOrCategory, [string]$SnapshotId) {
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $structure = Get-Structure $Language
    $scope = Get-LayoutScopeInfo $structure $PackIdOrCategory $true
    $p = Join-Path (Get-LayoutHistoryDir $Language ([string]$scope.storageId)) ("{0}.json" -f $safeSnapshotId)
    if (-not (Test-Path -LiteralPath $p)) { throw "レイアウト履歴が見つかりません: $SnapshotId" }
    $snapshot = Read-JsonFile $p $null
    if ($null -eq $snapshot) { throw "レイアウト履歴を読み込めません: $SnapshotId" }
    $recordedSnapshotId = [string](Get-DataProperty $snapshot 'snapshotId' '')
    $recordedLanguage = [string](Get-DataProperty $snapshot 'language' '')
    if (-not [string]::IsNullOrWhiteSpace($recordedSnapshotId) -and $recordedSnapshotId -ne $safeSnapshotId) { throw 'ページ構成履歴の識別子がファイル名と一致しません。' }
    if (-not [string]::IsNullOrWhiteSpace($recordedLanguage) -and $recordedLanguage -ne $Language) { throw '別の言語のページ構成履歴は復元できません。' }
    $recordedPackId = [string](Get-DataProperty $snapshot 'packId' '')
    $recordedCategory = [string](Get-DataProperty $snapshot 'category' '')
    if (-not [string]::IsNullOrWhiteSpace($recordedPackId) -and $recordedPackId -ne [string]$scope.packId) { throw '別の一式のページ構成履歴は復元できません。' }
    if ([string]::IsNullOrWhiteSpace($recordedPackId) -and -not [bool]$scope.builtIn) { throw '旧形式のページ構成履歴は任意一式へ復元できません。' }
    if ([string]::IsNullOrWhiteSpace($recordedPackId) -and $recordedCategory -ne [string]$scope.category) { throw '別のカテゴリのページ構成履歴は復元できません。' }
    return $snapshot
}

function Get-LayoutRestorePreview([string]$Language, [string]$PackIdOrCategory, [string]$SnapshotId) {
    $snap = Read-LayoutSnapshot $Language $PackIdOrCategory $SnapshotId
    $st = Get-Structure $Language
    $scope = Get-LayoutScopeInfo $st $PackIdOrCategory $true
    $currentIds = @{}
    foreach ($p in @(Get-Array $st.pages)) {
        $wb = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq [string]$p.workbookId } | Select-Object -First 1)
        if ($wb.Count -eq 0) { continue }
        if (-not (Test-WorkbookPack $wb[0] ([string]$scope.packId))) { continue }
        $currentIds[(Resolve-PageId $p)] = $p
    }
    $allowedVolumes = @(Get-PackVolumeList $Language $scope.pack $true)
    $applied = 0; $pastOnly = @(); $volumeChanges = @(); $invalidVolumePageIds = @()
    $snapIds = @{}
    foreach ($normalized in @(Get-NormalizedLayoutSnapshotPages $snap $allowedVolumes)) {
        $pageKey = [string]$normalized.pageId
        $snapIds[$pageKey] = $true
        if (-not $currentIds.ContainsKey($pageKey)) { $pastOnly += $pageKey; continue }
        $applied++
        $cur = $currentIds[$pageKey]
        if ([bool]$normalized.volumeInvalid) { $invalidVolumePageIds += $pageKey }
        if ([string]$cur.volume -ne [string]$normalized.volume) {
            $volumeChanges += [ordered]@{ pageId = $pageKey; title = [string]$cur.title; from = [string]$cur.volume; to = [string]$normalized.volume; storedVolume = [string]$normalized.storedVolume; normalized = [bool]$normalized.volumeInvalid }
        }
    }
    $currentOnly = @($currentIds.Keys | Where-Object { -not $snapIds.ContainsKey($_) })
    return [ordered]@{
        snapshotId = $SnapshotId
        packId = [string]$scope.packId
        category = [string]$scope.category
        createdAt = [string](Get-DataProperty $snap 'createdAt' '')
        reason = [string](Get-DataProperty $snap 'reason' '')
        appliedPageCount = $applied
        pastOnlyPageIds = @($pastOnly)
        currentOnlyPageIds = @($currentOnly)
        volumeChanges = @($volumeChanges)
        invalidVolumePageIds = @($invalidVolumePageIds)
        requiresRebuild = ($applied -gt 0)
    }
}

function Restore-LayoutSnapshot([string]$Language, [string]$PackIdOrCategory, [string]$SnapshotId) {
    $snap = Read-LayoutSnapshot $Language $PackIdOrCategory $SnapshotId
    $scope = Get-LayoutScopeInfo (Get-Structure $Language) $PackIdOrCategory $false
    $restoreResult = Update-StructureLocked $Language {
        param($st)
        $lockedScope = Get-LayoutScopeInfo $st ([string]$scope.packId) $false
        $allowedVolumes = @(Get-PackVolumeList $Language $lockedScope.pack $true)
        $workbookIds = @{}
        foreach ($wb in @(Get-Array $st.workbooks | Where-Object { Test-WorkbookPack $_ ([string]$lockedScope.packId) })) { $workbookIds[[string]$wb.workbookId] = $true }
        $map = @{}
        foreach ($p in @(Get-Array $st.pages | Where-Object { $workbookIds.ContainsKey([string]$_.workbookId) })) { $map[(Resolve-PageId $p)] = $p }
        $records = @()
        foreach ($normalized in @(Get-NormalizedLayoutSnapshotPages $snap $allowedVolumes)) {
            $pageKey = [string]$normalized.pageId
            if (-not $map.ContainsKey($pageKey)) { continue }
            $records += [pscustomobject][ordered]@{ page = $map[$pageKey]; snapshot = $normalized }
        }
        if ($records.Count -eq 0) { return [ordered]@{ applied = 0; undoSnapshotId = ''; invalidVolumePageCount = 0 } }
        # Save the exact locked pre-restore state. A failed undo snapshot must abort
        # before structure.json is changed.
        $undoId = Save-LayoutSnapshot $Language ([string]$lockedScope.packId) 'pre-restore' $st
        if ([string]::IsNullOrWhiteSpace($undoId)) { throw '復元直前のページ構成を保存できないため、復元を中止しました。' }
        $n = 0; $invalidVolumePageCount = 0
        foreach ($record in $records) {
            $p = $record.page
            $sp = $record.snapshot
            if ([bool]$sp.volumeInvalid) { $invalidVolumePageCount++ }
            # V5-§4.2: 適用してよいのはレイアウト項目のみ。
            # contentPdf / status / warnings / currentExcelHash / lastRendered* / volumes は触らない。
            Set-NoteProperty $p 'title' ([string]$sp.title)
            Set-NoteProperty $p 'volume' ([string]$sp.volume)
            Set-NoteProperty $p 'enabled' ([bool]$sp.enabled)
            Set-NoteProperty $p 'order' ([double]$sp.order)
            Set-NoteProperty $p 'orderManual' ([bool]$sp.orderManual)
            Set-NoteProperty $p 'numberingMode' ([string]$sp.numberingMode)
            Set-NoteProperty $p 'numberingManual' ([bool]$sp.numberingManual)
            Set-NoteProperty $p 'pageRange' (Get-DataProperty $sp 'pageRange' $null)
            Set-NoteProperty $p 'updatedAt' (New-NowIso)
            $n++
        }
        foreach ($volume in $allowedVolumes) { [void](Renumber-VolumeOrder $st $volume ([string]$lockedScope.packId)) }
        Apply-DefaultNumberingPerVolume $Language $st ([string]$lockedScope.packId)
        $vols = @(Get-PackVolumeList $Language $lockedScope.pack $false)
        Mark-VolumeNeedsRebuild $st $Language ([string]$lockedScope.packId) $vols 'layout-restored' 'ページ構成を過去の状態へ戻しました'
        return [ordered]@{ applied = $n; undoSnapshotId = $undoId; invalidVolumePageCount = $invalidVolumePageCount }
    }
    Write-HistoryEvent $Language 'layout.restored' ([ordered]@{ packId = [string]$scope.packId; category = [string]$scope.category; snapshotId = $SnapshotId; undoSnapshotId = [string]$restoreResult.undoSnapshotId; appliedPageCount = [int]$restoreResult.applied })
    return [ordered]@{ packId = [string]$scope.packId; category = [string]$scope.category; appliedPageCount = [int]$restoreResult.applied; undoSnapshotId = [string]$restoreResult.undoSnapshotId; invalidVolumePageCount = [int]$restoreResult.invalidVolumePageCount }
}

# ---- 提出用PDFアーカイブ (V5-§4.1) -----------------------------------

# pin は、提出用PDFが参照している source / content-pdf を後日の掃除から守る唯一の
# 仕組みである(Invoke-InputHistoryCleanup を参照。pin が無い版は保持期間の猶予なしに
# 削除の対象になる)。アーカイブの作成に失敗しても、この保護だけは残す必要があるため、
# アーカイブ本体から切り離してある。
function Set-FinalPdfSnapshotPins([string]$Language, [string]$PackId, [string]$Category, [string]$Volume, [string]$BuildId, $SourceWorkbooks, [string]$ArchivePath) {
    foreach ($sw in @(Get-Array $SourceWorkbooks)) {
        $wbId = [string](Get-DataProperty $sw 'workbookId' '')
        $snapshotId = [string](Get-DataProperty $sw 'snapshotId' '')
        $versionId = [string](Get-DataProperty $sw 'versionId' '')
        if ([string]::IsNullOrWhiteSpace($wbId) -or [string]::IsNullOrWhiteSpace($snapshotId)) { continue }
        $pinOk = New-SnapshotPin $Language $wbId $snapshotId ("final-pdf_{0}" -f $BuildId) ([ordered]@{
            buildId = $BuildId; snapshotId = $snapshotId; versionId = $versionId
            volume = $Volume; packId = $PackId; category = $Category; archivePath = $ArchivePath
        })
        if (-not $pinOk) { throw ("スナップショット保護(pin)を作成できませんでした: {0} / {1}" -f $wbId, $snapshotId) }
        if ($versionId) {
            $cpPinOk = New-ContentPdfPin (Get-WorkspacePath $Language) $wbId $versionId ("final-pdf_{0}" -f $BuildId) ([ordered]@{
                buildId = $BuildId; volume = $Volume; packId = $PackId; category = $Category
            })
            if (-not $cpPinOk) { throw ("content-pdf 保護(pin)を作成できませんでした: {0} / {1}" -f $wbId, $versionId) }
        }
    }
}

function New-FinalArchive([string]$Language, [string]$Category, [string]$Volume, [string]$BuildId, [string]$OutputPdf, $Manifest, $Snapshot) {
    # 冪等: 一時フォルダで完成させてから buildId フォルダへ移動する。
    try {
        $archiveCategory = Normalize-WorkbookCategory $Category ''
        $archivePackId = if ([string]::IsNullOrWhiteSpace($archiveCategory)) { Assert-SafeStorageSegment $Category 'packId' } else { Get-BuiltinPackId $archiveCategory }
        $archiveStorageId = if ([string]::IsNullOrWhiteSpace($archiveCategory)) { $archivePackId } else { $archiveCategory }
        $target = Join-Path (Join-Path (Join-Path (Get-ArchiveRoot $Language) $archiveStorageId) $Volume) $BuildId
        if (Test-Path -LiteralPath $target) { return $target }
        $stage = $target + '.staging'
        if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
        New-Item -ItemType Directory -Path $stage -Force | Out-Null
        Copy-Item -LiteralPath $OutputPdf -Destination (Join-Path $stage 'final.pdf') -Force
        $sha = Normalize-FileHash (New-Sha256 (Join-Path $stage 'final.pdf'))
        Set-Content -LiteralPath (Join-Path $stage 'sha256.txt') -Value $sha -Encoding ASCII
        Write-JsonFile (Join-Path $stage 'manifest.json') $Manifest

        # V5-P0: 出力に「実際に使った」不変の情報だけを記録する。
        # ここで構造データを読み直すと、出力後に別レンダリングが完了した場合に
        # 正式PDFで使っていない snapshot / versionId を記録してしまう。
        $envs = [ordered]@{}
        $sourceWorkbooks = @()
        foreach ($sw in @(Get-Array (Get-DataProperty $Snapshot 'sourceWorkbooks' @()))) {
            $fp = [string](Get-DataProperty $sw 'renderEnvironmentFingerprint' '')
            if ($fp -and -not $envs.Contains($fp)) { $envs[$fp] = (Get-DataProperty $sw 'renderEnvironment' $null) }
            $sourceWorkbooks += $sw
        }
        Write-JsonFile (Join-Path $stage 'metadata.json') ([ordered]@{
            schemaVersion = 2; buildId = $BuildId; language = $Language; packId = $archivePackId; category = $archiveCategory; targetId = (Get-TargetIdFromLegacyVolume $Volume); volume = $Volume
            projectId = [string](Get-DataProperty $Snapshot 'projectId' ''); builtAt = New-NowIso
            inputFingerprint = [string](Get-DataProperty $Snapshot 'fingerprint' '')
            outputPdfSha256 = $sha; outputFileName = [IO.Path]::GetFileName($OutputPdf)
            pageCount = @(Get-Array (Get-DataProperty $Snapshot 'pages' @())).Count
            sourceWorkbooks = @($sourceWorkbooks)
            excelPrintProfileVersion = $Script:ExcelPrintProfileVersion
            renderEnvironments = $envs
            composerEnvironment = [ordered]@{ pcName = $env:COMPUTERNAME; osVersion = [Environment]::OSVersion.VersionString }
            builtBy = [ordered]@{ pcName = $env:COMPUTERNAME; userName = "$env:USERDOMAIN\$env:USERNAME" }
        })
        Move-Item -LiteralPath $stage -Destination $target -Force

        # V5-P0: pin はアーカイブが正式フォルダへ移動できてから作る。
        # 先に作ると、移動に失敗したときにアーカイブが無いのに pin だけ残る。
        Set-FinalPdfSnapshotPins $Language $archivePackId $archiveCategory $Volume $BuildId $sourceWorkbooks $target
        Write-HistoryEvent $Language 'final.archive.created' ([ordered]@{ packId = $archivePackId; category = $archiveCategory; targetId = (Get-TargetIdFromLegacyVolume $Volume); volume = $Volume; buildId = $BuildId; path = $target })
        return $target
    } catch {
        # V5-P0: 握りつぶさない。呼出元がロールバックする。
        try { if (Test-Path -LiteralPath ($target + '.staging')) { Remove-Item -LiteralPath ($target + '.staging') -Recurse -Force -ErrorAction SilentlyContinue } } catch { }
        throw ("提出用PDFのアーカイブに失敗しました: " + $_.Exception.Message)
    }
}

function Remove-FinalArchiveArtifacts([string]$Language, [string]$Category, [string]$Volume, [string]$BuildId) {
    # ロールバック時に、部分的にできたアーカイブと pin を掃除する。
    try {
        $target = Join-Path (Join-Path (Join-Path (Get-ArchiveRoot $Language) $Category) $Volume) $BuildId
        foreach ($path in @($target, ($target + '.staging'))) {
            if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue }
        }
        $root = Get-InputHistoryRoot $Language
        if (Test-Path -LiteralPath $root) {
            foreach ($wbDir in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
                foreach ($snDir in @(Get-ChildItem -LiteralPath $wbDir.FullName -Directory -ErrorAction SilentlyContinue)) {
                    Remove-SnapshotPin $Language $wbDir.Name $snDir.Name ("final-pdf_{0}" -f $BuildId)
                }
            }
        }
        $cpRoot = Join-Path (Get-WorkspacePath $Language) 'content-pdf'
        if (Test-Path -LiteralPath $cpRoot) {
            foreach ($f in @(Get-ChildItem -LiteralPath $cpRoot -Recurse -File -Filter ("final-pdf_{0}.json" -f $BuildId) -ErrorAction SilentlyContinue)) {
                Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
            }
        }
    } catch { }
}

function Get-FinalArchives([string]$Language, [string]$Category) {
    $structure = Get-Structure $Language
    $scope = Resolve-DocumentPackScope $structure $Category $true
    $packId = Assert-SafeStorageSegment ([string]$scope.packId) 'packId'
    $archiveStorageId = if ([bool]$scope.builtIn) { [string]$scope.category } else { $packId }
    $root = Join-Path (Get-ArchiveRoot $Language) $archiveStorageId
    if (-not (Test-Path -LiteralPath $root)) { return @() }
    $out = @()
    foreach ($volDir in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
        foreach ($b in @(Get-ChildItem -LiteralPath $volDir.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 50)) {
            try {
                $meta = Read-JsonFile (Join-Path $b.FullName 'metadata.json') $null
                if ($null -eq $meta) { continue }
                $out += [ordered]@{
                    buildId = [string](Get-DataProperty $meta 'buildId' $b.Name)
                    packId = [string](Get-DataProperty $meta 'packId' $packId)
                    category = [string](Get-DataProperty $meta 'category' ([string]$scope.category))
                    targetId = [string](Get-DataProperty $meta 'targetId' (Get-TargetIdFromLegacyVolume ([string]$volDir.Name)))
                    volume = [string]$volDir.Name
                    builtAt = [string](Get-DataProperty $meta 'builtAt' '')
                    pageCount = [int](Get-DataProperty $meta 'pageCount' 0)
                    outputFileName = [string](Get-DataProperty $meta 'outputFileName' '')
                    outputPdfSha256 = [string](Get-DataProperty $meta 'outputPdfSha256' '')
                    builtBy = (Get-DataProperty $meta 'builtBy' $null)
                    path = [string]$b.FullName
                }
            } catch { }
        }
    }
    return @($out | Sort-Object { [string]$_.builtAt } -Descending)
}


# ---- 正式出力トランザクション (V5-§7.2) ----------------------------

function Write-FinalJournal([string]$Language, $Journal) {
    $dir = Get-FinalTransactionDir $Language
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Set-NoteProperty $Journal 'updatedAt' (New-NowIso)
    Write-JsonFile (Join-Path $dir ("{0}.json" -f [string]$Journal.transactionId)) $Journal
}
function Set-JournalPhase([string]$Language, $Journal, [string]$Phase) {
    # V5-§7.2: phase はその副作用を「始める前」に書く。後で書くと復旧できない窓が残る。
    Set-NoteProperty $Journal 'phase' $Phase
    Write-FinalJournal $Language $Journal
}
function Get-JournalTarget($Journal, [string]$Volume) {
    foreach ($t in @(Get-Array $Journal.targets)) { if ([string]$t.volume -eq $Volume) { return $t } }
    return $null
}

function Assert-FinalBuildSourcesUnchanged($Snapshots, [string[]]$Volumes) {
    # 提出用PDFの対象になったExcelだけを、共有フォルダー上のサイズ・更新時刻で確認する。
    # Scan-Updates は全登録Excelを走査して structure.json もブックごとに更新するため、
    # 最終出力のたびに呼ぶ必要はない。
    $paths = Get-Paths
    $seen = @{}
    foreach ($v in $Volumes) {
        $snap = Get-DataProperty $Snapshots $v $null
        foreach ($w in @(Get-Array (Get-DataProperty $snap 'sourceWorkbooks' @()))) {
            $id = [string](Get-DataProperty $w 'workbookId' '')
            if ([string]::IsNullOrWhiteSpace($id) -or $seen.ContainsKey($id)) { continue }
            $seen[$id] = $true
            $relativePath = [string](Get-DataProperty $w 'relativePath' '')
            $name = [string](Get-DataProperty $w 'fileName' $relativePath)
            $full = Join-Safe ([string]$paths.submissionDir) $relativePath
            try { $item = Get-Item -LiteralPath $full -ErrorAction Stop }
            catch { throw [InvalidOperationException]::new("$name が提出フォルダーに見つかりません。先にPDF作成状態を確認してください。") }
            $knownTicks = [string](Get-DataProperty $w 'currentExcelLastWriteUtcTicks' '')
            $knownSize = [int64](Get-DataProperty $w 'currentExcelSize' -1)
            if (-not [string]::IsNullOrWhiteSpace($knownTicks) -and $knownSize -ge 0) {
                if ([string]$item.LastWriteTimeUtc.Ticks -ne $knownTicks -or [int64]$item.Length -ne $knownSize) {
                    throw [InvalidOperationException]::new("$name の元原稿が更新されています。先に変換PDFを作成してください。")
                }
                continue
            }
            # 旧データでメタデータがない場合だけ、対象ファイル1件のハッシュ確認へ縮退する。
            $knownHash = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
            $liveHash = Normalize-FileHash (New-StableHash $full)
            if ([string]::IsNullOrWhiteSpace($knownHash) -or $liveHash -ne $knownHash) {
                throw [InvalidOperationException]::new("$name の元原稿が更新されています。先に変換PDFを作成してください。")
            }
        }
    }
}

function Invoke-FinalBuildTransaction([string]$Language, [string]$Category, [string[]]$Volumes) {
    # 単体出力もまとめて出力も、必ずこの1本を通る(単体だけ障害復旧が無い状態を作らない)。
    $cat = Require-WorkbookCategory $Category
    $paths = Get-Paths
    $workspace = Get-WorkspacePath $Language

    $allowed = @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })
    $requested = @($Volumes | Where-Object { $allowed -contains $_ })
    if ($requested.Count -eq 0) { throw [System.ArgumentException]::new('volumeには本体または補足を指定してください。') }

    # V5-§7.2: 対象は pageCount > 0 の volume のみ。
    # Get-FinalBuildInputSnapshot は 0ページを no-pages blocker にするため、
    # 補足が空の案件で「まとめて出力」が常に失敗してしまう。
    $lockPath = Join-Path $workspace ("locks\final-build_{0}.lock" -f $cat)
    return Invoke-WithLock $lockPath {
        $snapshots = @{}
        $targets = @()
        $skipped = @()
        # 既定番号の適用と全volumeの入力固定を、structure.json 1回の更新で済ませる。
        $initialSnapshots = Update-StructureLocked $Language {
            param($st)
            Apply-DefaultNumberingPerVolume $Language $st $cat
            $all = [ordered]@{}
            foreach ($v in $requested) {
                $sn = Get-FinalBuildInputSnapshot $st $Language $v $cat
                # V5-P0: アーカイブ用の情報はこの時点で固定する(後で structure を読み直さない)。
                $seen = @{}
                $sw = @()
                foreach ($pg in @(Get-Array (Get-DataProperty $sn 'pages' @()))) {
                    $wid = [string]$pg.workbookId
                    if ([string]::IsNullOrWhiteSpace($wid) -or $seen.ContainsKey($wid)) { continue }
                    $seen[$wid] = $true
                    $w = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq $wid } | Select-Object -First 1)
                    if ($w.Count -eq 0) { continue }
                    $sw += [ordered]@{
                        workbookId = $wid
                        fileName = [string]$w[0].fileName
                        relativePath = [string]$w[0].relativePath
                        currentExcelLastWriteUtcTicks = [string](Get-DataProperty $w[0] 'currentExcelLastWriteUtcTicks' '')
                        currentExcelSize = [int64](Get-DataProperty $w[0] 'currentExcelSize' -1)
                        currentExcelHash = Normalize-FileHash ([string](Get-DataProperty $w[0] 'currentExcelHash' ''))
                        snapshotId = [string](Get-DataProperty $w[0] 'lastRenderedSnapshotId' '')
                        versionId = [string](Get-DataProperty $w[0] 'lastRenderedVersionId' '')
                        sourceHash = Normalize-FileHash ([string](Get-DataProperty $w[0] 'lastRenderedExcelHash' ''))
                        renderEnvironmentFingerprint = [string](Get-DataProperty $w[0] 'renderEnvironmentFingerprint' '')
                        renderEnvironment = (Get-DataProperty $w[0] 'renderEnvironment' $null)
                    }
                }
                Set-NoteProperty $sn 'sourceWorkbooks' @($sw)
                $all[$v] = $sn
            }
            return $all
        }
        foreach ($v in $requested) {
            $snap = Get-DataProperty $initialSnapshots $v $null
            if ($null -eq $snap -or [int]$snap.pageCount -le 0) { $skipped += $v; continue }
            $blockers = @(Get-Array $snap.blockers | Where-Object { [string]$_.code -ne 'no-pages' })
            if ($blockers.Count -gt 0) { throw [InvalidOperationException]::new(("{0}：{1}" -f (Get-VolumeLabelForMessage $v), [string]$blockers[0].message)) }
            $snapshots[$v] = $snap
            $targets += $v
        }
        if ($targets.Count -eq 0) { return [ordered]@{ built = @(); skipped = @($skipped); message = '出力対象がありません。' } }
        Assert-FinalBuildSourcesUnchanged $snapshots $targets

        $composerJar = Join-Path $Script:AppRoot 'lib\pdfbox\ReportPdfComposer.jar'
        $pdfboxJar = Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'
        if (-not (Test-Path $composerJar)) { throw 'ReportPdfComposer.jar がありません。' }
        if (-not (Test-Path $pdfboxJar)) { throw 'pdfbox-app.jar がありません。' }

        $txId = New-RbId
        $journal = [ordered]@{
            schemaVersion = 1; transactionId = $txId; language = $Language; category = $cat
            startedAt = New-NowIso; phase = 'prepared'
            targets = @($targets | ForEach-Object { [ordered]@{ volume = $_; buildId = (New-RbId); backupCreated = $false; fileReplaced = $false; oldPdfHash = ''; newPdfHash = ''; existed = $false; finalPath = ''; backupPath = ''; tempPath = '' } })
            beforeFingerprints = [ordered]@{}
            oldVolumeStates = [ordered]@{}
            newVolumeStates = [ordered]@{}
        }
        foreach ($v in $targets) { $journal.beforeFingerprints[$v] = [string]$snapshots[$v].fingerprint }
        Write-FinalJournal $Language $journal

        $backupDir = Join-Path $workspace ("state\final-backups\" + $txId)
        try {
            # 5. 出力先PDFが開かれていないかを、対象すべてまとめて確認
            foreach ($v in $targets) {
                $t = Get-JournalTarget $journal $v
                $outName = [string]$snapshots[$v].outputFileName
                $outPath = Join-Path ([string]$paths.outputDir) $outName
                Set-NoteProperty $t 'finalPath' $outPath
                Set-NoteProperty $t 'existed' ([bool](Test-Path -LiteralPath $outPath))
                if (Test-Path -LiteralPath $outPath) {
                    $f = $null
                    try { $f = [IO.File]::Open($outPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
                    catch { throw "出力先の提出用PDFが開かれているため上書きできません: $outName" }
                    finally { if ($f) { $f.Dispose() } }
                }
            }
            Write-FinalJournal $Language $journal

            # 7-8. 一時ファイルへ組版
            foreach ($v in $targets) {
                $t = Get-JournalTarget $journal $v
                $tmp = Join-Path ([string]$paths.outputDir) ("~building_{0}_{1}_{2}.pdf" -f $v, $cat, $txId)
                if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
                Set-NoteProperty $t 'tempPath' $tmp
                $snap = $snapshots[$v]
                $manifest = [ordered]@{ schemaVersion=3; language=$Language; category=$cat; volume=$v; projectId=[string]$snap.projectId; inputFingerprint=[string]$snap.fingerprint; outputPdf=$tmp; createdAt=New-NowIso; document=$snap.document; pageNumber=[ordered]@{font='Arial';fontSize=8;bottomPt=18;format='hyphenated';countHidden=$true}; physicalPages=$snap.physicalPages; pages=$snap.manifestPages }
                $manifestPath = Join-Path ([IO.Path]::GetTempPath()) ("ReportBinder_final_{0}_{1}_{2}.json" -f $v, $cat, $txId)
                try {
                    Write-JsonFile $manifestPath $manifest
                    $java = Resolve-JavaExe
                    $run = Invoke-NativeCapture $java @('-cp', "$composerJar;$pdfboxJar", 'ReportPdfComposer', '--manifest', $manifestPath) $Script:FinalComposeTimeoutSeconds
                    $exit = [int]$run.exitCode
                    $text = [string]$run.text
                    if ($run.timedOut) { throw ('PDFの結合が{0}分以内に終わりませんでした。出力先がネットワーク上のフォルダーの場合は、いったんPC内のフォルダーに出力してみてください。' -f [int]($Script:FinalComposeTimeoutSeconds / 60)) }
                    if ($exit -ne 0) { throw "PDFBox組版に失敗しました。exit=$exit`n$text" }
                } finally {
                    if (Test-Path -LiteralPath $manifestPath) { Remove-Item -LiteralPath $manifestPath -Force -ErrorAction SilentlyContinue }
                }
                if (-not (Test-Path $tmp) -or (Get-Item $tmp).Length -le 0) { throw '提出用PDFを作成できませんでした。' }
                Set-NoteProperty $t 'newPdfHash' (Normalize-FileHash (New-Sha256 $tmp))
                Set-NoteProperty $t 'manifest' $manifest
            }
            Write-FinalJournal $Language $journal

            # 9. fingerprint 再確認。読取だけなので structure.json をvolumeごとに再保存しない。
            $afterStructure = Get-Structure $Language
            foreach ($v in $targets) {
                $after = Get-FinalBuildInputSnapshot $afterStructure $Language $v $cat
                if ([string]$after.fingerprint -ne [string]$journal.beforeFingerprints[$v]) {
                    throw 'PDF作成中にページ構成またはPDF入力が変更されました。最新の状態で再度出力してください。'
                }
            }

            # 10. バックアップ
            Set-JournalPhase $Language $journal 'backups-created'
            if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
            foreach ($v in $targets) {
                $t = Get-JournalTarget $journal $v
                if ([bool]$t.existed) {
                    $bak = Join-Path $backupDir ("{0}.pdf" -f $v)
                    Copy-Item -LiteralPath ([string]$t.finalPath) -Destination $bak -Force
                    Set-NoteProperty $t 'backupPath' $bak
                    Set-NoteProperty $t 'oldPdfHash' (Normalize-FileHash (New-Sha256 $bak))
                }
                Set-NoteProperty $t 'backupCreated' $true
                Write-FinalJournal $Language $journal
            }

            # 11. 差し替え
            Set-JournalPhase $Language $journal 'replacing-files'
            foreach ($v in $targets) {
                $t = Get-JournalTarget $journal $v
                Move-Item -LiteralPath ([string]$t.tempPath) -Destination ([string]$t.finalPath) -Force
                Set-NoteProperty $t 'fileReplaced' $true
                Write-FinalJournal $Language $journal
            }
            Set-JournalPhase $Language $journal 'files-replaced'

            # 12. structure 更新(ロック内で fingerprint 再確認、old/new を先に書き終える)
            Set-JournalPhase $Language $journal 'structure-committing'
            $commit = Update-StructureLocked $Language {
                param($st)
                foreach ($v in $targets) {
                    $after = Get-FinalBuildInputSnapshot $st $Language $v $cat
                    if ([string]$after.fingerprint -ne [string]$journal.beforeFingerprints[$v]) { return [ordered]@{ changed = $true; volume = $v } }
                }
                foreach ($v in $targets) {
                    $key = Get-VolumeStateKey $v $cat
                    $old = Get-DataProperty $st.volumes $key $null
                    $journal.oldVolumeStates[$v] = $(if ($null -eq $old) { $null } else { [ordered]@{
                        builtFingerprint = [string](Get-DataProperty $old 'builtFingerprint' '')
                        status = [string](Get-DataProperty $old 'status' '')
                        outputPdf = [string](Get-DataProperty $old 'outputPdf' '')
                        lastBuiltAt = [string](Get-DataProperty $old 'lastBuiltAt' '')
                        staleReasons = @(Get-Array (Get-DataProperty $old 'staleReasons' @()))
                    } })
                    $t = Get-JournalTarget $journal $v
                    $journal.newVolumeStates[$v] = [ordered]@{
                        builtFingerprint = [string]$journal.beforeFingerprints[$v]
                        status = 'built'; outputPdf = [string]$t.finalPath; lastBuiltAt = New-NowIso; staleReasons = @()
                    }
                }
                Write-FinalJournal $Language $journal
                foreach ($v in $targets) {
                    $key = Get-VolumeStateKey $v $cat
                    $vs = Get-DataProperty $st.volumes $key $null
                    if ($null -eq $vs) { $vs = New-EmptyVolumeState; Set-NoteProperty $st.volumes $key $vs }
                    $n = $journal.newVolumeStates[$v]
                    Set-NoteProperty $vs 'builtFingerprint' ([string]$n.builtFingerprint)
                    Set-NoteProperty $vs 'lastBuiltAt' ([string]$n.lastBuiltAt)
                    Set-NoteProperty $vs 'outputPdf' ([string]$n.outputPdf)
                    Set-NoteProperty $vs 'staleReasons' @()
                    # 同じstructureロック内でfingerprintを確認済み。再度全ページを走査しない。
                    Set-NoteProperty $vs 'status' 'built'
                }
                return [ordered]@{ changed = $false }
            }
            if ([bool]$commit.changed) { throw 'PDF作成中にページ構成またはPDF入力が変更されました。最新の状態で再度出力してください。' }
            Set-JournalPhase $Language $journal 'structure-committed'

            # 13. アーカイブと pin(冪等)
            Set-JournalPhase $Language $journal 'archiving'
            $built = @()
            foreach ($v in $targets) {
                $t = Get-JournalTarget $journal $v
                $savedManifest = Get-DataProperty $t 'manifest' $null
                $savedManifestPath = Join-Path $workspace ("exports\manifest_{0}_{1}.json" -f $v, $cat)
                if ($null -ne $savedManifest) {
                    Set-NoteProperty $savedManifest 'outputPdf' ([string]$t.finalPath)
                    Write-JsonFile $savedManifestPath $savedManifest
                }
                # 保存済み出力・ページ構成履歴は廃止。共有フォルダーへの複製を作らない。
                $archive = ''
                Write-HistoryEvent $Language 'final.built' ([ordered]@{ category = $cat; volume = $v; buildId = [string]$t.buildId; outputPdf = [string]$t.finalPath })
                $built += [ordered]@{ volume = $v; category = $cat; outputPdf = [string]$t.finalPath; manifestPath = $savedManifestPath; buildId = [string]$t.buildId; archivePath = $archive; inputFingerprint = [string]$journal.beforeFingerprints[$v] }
            }
            Set-JournalPhase $Language $journal 'completed'
            if (Test-Path -LiteralPath $backupDir) { Remove-Item -LiteralPath $backupDir -Recurse -Force -ErrorAction SilentlyContinue }
            return [ordered]@{ built = @($built); skipped = @($skipped); transactionId = $txId }
        } catch {
            foreach ($t in @(Get-Array $journal.targets)) { Remove-FinalArchiveArtifacts $Language $cat ([string]$t.volume) ([string]$t.buildId) }
            Restore-FinalTransaction $Language $journal
            Write-HistoryEvent $Language 'final.build.failed' ([ordered]@{ category = $cat; transactionId = $txId; message = $_.Exception.Message })
            throw
        }
    }
}

function Restore-FinalTransaction([string]$Language, $Journal) {
    # V5-§B-1: 復旧判定はフラグではなく実ファイルのハッシュを正とする。
    # fileReplaced=true を書く前に落ちる窓があるため。
    try {
        $phase = [string](Get-DataProperty $Journal 'phase' '')
        foreach ($t in @(Get-Array $Journal.targets)) {
            $final = [string](Get-DataProperty $t 'finalPath' '')
            if ([string]::IsNullOrWhiteSpace($final)) { continue }
            $tmp = [string](Get-DataProperty $t 'tempPath' '')
            if ($tmp -and (Test-Path -LiteralPath $tmp)) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
            $newHash = Normalize-FileHash ([string](Get-DataProperty $t 'newPdfHash' ''))
            $oldHash = Normalize-FileHash ([string](Get-DataProperty $t 'oldPdfHash' ''))
            $existed = [bool](Get-DataProperty $t 'existed' $false)
            $backup = [string](Get-DataProperty $t 'backupPath' '')
            $current = ''
            if (Test-Path -LiteralPath $final) { $current = Normalize-FileHash (New-Sha256 $final) }

            if ([string]::IsNullOrWhiteSpace($current)) {
                if (-not $existed) { continue }            # 元から無く、今も無い
                if ($backup -and (Test-Path -LiteralPath $backup)) { Copy-Item -LiteralPath $backup -Destination $final -Force }
                continue
            }
            if ($newHash -and $current -eq $newHash) {
                if (-not $existed) { Remove-Item -LiteralPath $final -Force -ErrorAction SilentlyContinue }
                elseif ($backup -and (Test-Path -LiteralPath $backup)) { Copy-Item -LiteralPath $backup -Destination $final -Force }
                continue
            }
            if ($oldHash -and $current -eq $oldHash) { continue }   # 未差し替え
            # どれとも一致しない: 自動で上書きしない。
            Set-NoteProperty $Journal 'phase' 'manual-recovery-required'
            Set-NoteProperty $Journal 'manualRecoveryReason' ("出力先PDFが記録されたどの状態とも一致しません: " + $final)
            Write-FinalJournal $Language $Journal
            return
        }
        # structure の巻き戻し(対象volumeの項目だけ)
        if (@('structure-committing','structure-committed','archiving') -contains $phase) {
            $olds = Get-DataProperty $Journal 'oldVolumeStates' $null
            $news = Get-DataProperty $Journal 'newVolumeStates' $null
            if ($null -ne $olds) {
                $cat = [string](Get-DataProperty $Journal 'category' '')
                Update-StructureLocked $Language {
                    param($st)
                    foreach ($t in @(Get-Array $Journal.targets)) {
                        $v = [string]$t.volume
                        $key = Get-VolumeStateKey $v $cat
                        $vs = Get-DataProperty $st.volumes $key $null
                        if ($null -eq $vs) { continue }
                        $newState = Get-DataProperty $news $v $null
                        if ($null -ne $newState -and [string](Get-DataProperty $vs 'builtFingerprint' '') -ne [string](Get-DataProperty $newState 'builtFingerprint' '')) { continue }
                        $oldState = Get-DataProperty $olds $v $null
                        if ($null -eq $oldState) { continue }
                        Set-NoteProperty $vs 'builtFingerprint' ([string](Get-DataProperty $oldState 'builtFingerprint' ''))
                        Set-NoteProperty $vs 'status' ([string](Get-DataProperty $oldState 'status' 'not-built'))
                        Set-NoteProperty $vs 'outputPdf' ([string](Get-DataProperty $oldState 'outputPdf' ''))
                        Set-NoteProperty $vs 'lastBuiltAt' ([string](Get-DataProperty $oldState 'lastBuiltAt' ''))
                        Set-NoteProperty $vs 'staleReasons' @(Get-Array (Get-DataProperty $oldState 'staleReasons' @()))
                    }
                } | Out-Null
            }
        }
        Set-NoteProperty $Journal 'phase' 'rolled-back'
        Write-FinalJournal $Language $Journal
        # 巻き戻しに使い終えた旧PDFバックアップを残さない。
        Remove-FinalTransactionBackupDir $Language ([string](Get-DataProperty $Journal 'transactionId' ''))
    } catch { }
}

function Recover-FinalTransactions([string]$Language) {
    # V5: 起動時復旧も final-build ロックを取る(複数サーバーが同じ復旧を走らせない)。
    try {
        $dir = Get-FinalTransactionDir $Language
        if (-not (Test-Path -LiteralPath $dir)) { return }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue)) {
            $j = Read-JsonFile $f.FullName $null
            if ($null -eq $j) { continue }
            $phase = [string](Get-DataProperty $j 'phase' '')
            if (@('completed','rolled-back') -contains $phase) {
                # 完了済み/巻き戻し済みではバックアップは不要。前回削除に失敗していても起動時に再試行する。
                # ジャーナル本体は30日残し、manual-recovery-required は担当者が確認するまで残す。
                try {
                    Remove-FinalTransactionBackupDir $Language ([string](Get-DataProperty $j 'transactionId' $f.BaseName))
                    if ($f.LastWriteTimeUtc -lt [DateTime]::UtcNow.AddDays(-30)) {
                        Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
                    }
                } catch { }
                continue
            }
            if ($phase -eq 'manual-recovery-required') { continue }
            $cat = [string](Get-DataProperty $j 'category' '')
            $lockPath = Join-Path (Get-WorkspacePath $Language) ("locks\final-build_{0}.lock" -f $cat)
            $h = Try-AcquireLockHandle $lockPath
            if ($null -eq $h) { continue }
            try {
                # V5-P1(#8): 通常の失敗経路(Invoke-FinalBuildTransaction の catch)と同じく、
                # 起動時復旧でも部分的にできたアーカイブ/pin を掃除してから巻き戻す。
                # archiving 中に強制終了すると、一部の巻だけアーカイブ/pin が残り、structure だけ巻き戻る。
                foreach ($t in @(Get-Array $j.targets)) { Remove-FinalArchiveArtifacts $Language $cat ([string]$t.volume) ([string]$t.buildId) }
                Restore-FinalTransaction $Language $j
            } finally { Release-LockHandle $h }
        }
    } catch { }
}

function Get-VolumeLabelForMessage([string]$Volume) {
    switch ($Volume) {
        'ja-main' { return '本体' } 'ja-appendix' { return '補足' }
        'en-main' { return 'Main' } 'en-appendix' { return 'Appendix' }
        default { return $Volume }
    }
}


# ---- API 補助 (V5) --------------------------------------------------

function Get-HistoryTimeline([string]$Language, [int]$Limit) {
    try {
        $dir = Join-Path (Get-WorkspacePath $Language) 'history\events'
        if (-not (Test-Path -LiteralPath $dir)) { return @() }
        $out = @()
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First $Limit)) {
            try {
                $j = Read-JsonFile $f.FullName $null
                if ($null -ne $j) { $out += $j }
            } catch { }
        }
        return $out
    } catch { return @() }
}


function Get-ContentPdfSheetIndex([string]$Language, [string]$WorkbookId, [string]$VersionId) {
    $workspace = Get-WorkspacePath $Language
    $safeVersionId = Assert-SafeStorageSegment $VersionId 'versionId'
    $dir = Get-ContentPdfVersionDir $workspace $WorkbookId $safeVersionId
    $cacheKey = ([IO.Path]::GetFullPath($dir)).ToLowerInvariant()
    $cached = $Script:ContentPdfSheetIndexCache[$cacheKey]
    if ($null -ne $cached) {
        $cachedIndex = Get-DataProperty $cached 'index' @{}
        $allFilesExist = $true
        foreach ($cachedPath in @($cachedIndex.Values)) {
            if (-not (Test-FileExistsCompat ([string]$cachedPath))) { $allFilesExist = $false; break }
        }
        # An empty cached generation has no file whose existence can prove that
        # its parent still exists, so validate the directory in that case.
        if ($cachedIndex.Count -eq 0 -and -not (Test-DirectoryExistsCompat $dir)) { $allFilesExist = $false }
        if ($allFilesExist) { return $cachedIndex }
    }
    if (-not (Test-DirectoryExistsCompat $dir)) {
        [void]$Script:ContentPdfSheetIndexCache.Remove($cacheKey)
        return @{}
    }
    $stamp = [IO.Directory]::GetLastWriteTimeUtc($dir).Ticks
    $root = [IO.Path]::GetFullPath($workspace)
    if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
    $index = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.pdf' -ErrorAction SilentlyContinue)) {
        $full = [IO.Path]::GetFullPath($file.FullName)
        if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { continue }
        $key = ([string]$file.BaseName).ToLowerInvariant()
        if (-not $index.ContainsKey($key) -and (Test-FileExistsCompat $full)) { $index[$key] = $full }
    }
    if (-not $Script:ContentPdfSheetIndexCache.ContainsKey($cacheKey) -and
        $Script:ContentPdfSheetIndexCache.Count -ge $Script:ContentPdfSheetIndexCacheLimit) {
        $oldestKey = @($Script:ContentPdfSheetIndexCache.Keys)[0]
        if ($null -ne $oldestKey) { [void]$Script:ContentPdfSheetIndexCache.Remove($oldestKey) }
    }
    $Script:ContentPdfSheetIndexCache[$cacheKey] = [pscustomobject][ordered]@{ stamp = $stamp; index = $index }
    return $index
}


function Resolve-ContentPdfSheetPathFromIndex($Index, [string]$SheetName) {
    if ($null -eq $Index -or [string]::IsNullOrWhiteSpace($SheetName)) { return '' }
    # Hashed storage names are safe for every Excel sheet name. Legacy candidates keep
    # previously rendered numeric/ASCII workbooks readable during an upgrade.
    foreach ($candidate in @(
        (Get-WorksheetStorageStem $SheetName),
        $SheetName,
        ([regex]::Replace($SheetName, '[^0-9A-Za-z]+', '-'))
    )) {
        $key = ([string]$candidate).ToLowerInvariant()
        if ($Index.ContainsKey($key)) { return [string]$Index[$key] }
    }
    return ''
}
function Resolve-ContentPdfSheetPathExact([string]$Language, [string]$WorkbookId, [string]$VersionId, [string]$SheetName) {
    if ([string]::IsNullOrWhiteSpace($VersionId) -or [string]::IsNullOrWhiteSpace($SheetName)) { return '' }
    $index = Get-ContentPdfSheetIndex $Language $WorkbookId $VersionId
    return (Resolve-ContentPdfSheetPathFromIndex $index $SheetName)
}

function Get-HistoryRenderVersionAvailability([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId) {
    $result = [ordered]@{
        versionId = [string]$VersionId
        visualHashAvailable = $false
        contentPdfAvailable = $false
        ready = $false
        missingSheets = @()
        reason = ''
    }
    $hashes = Get-VisualHashes $Language $WorkbookId $SnapshotId $VersionId
    if ($null -eq $hashes) {
        $result.reason = '画像ハッシュが保存されていません。'
        return [pscustomobject]$result
    }
    $result.visualHashAvailable = $true
    $sheets = @(Get-Array (Get-DataProperty $hashes 'sheets' @()))
    if ($sheets.Count -eq 0) {
        $result.reason = '比較対象シートの画像ハッシュがありません。'
        return [pscustomobject]$result
    }
    # 版ディレクトリの列挙と更新時刻確認は1回だけ行い、全シートをローカル索引で照合する。
    $pdfIndex = Get-ContentPdfSheetIndex $Language $WorkbookId $VersionId
    $missing = @()
    foreach ($sheet in $sheets) {
        $name = [string](Get-DataProperty $sheet 'sheetName' '')
        if ([string]::IsNullOrWhiteSpace($name)) {
            $missing += '[sheetName missing]'
            continue
        }
        if ([string]::IsNullOrWhiteSpace((Resolve-ContentPdfSheetPathFromIndex $pdfIndex $name))) { $missing += $name }
    }
    $result.missingSheets = @($missing)
    $result.contentPdfAvailable = ($missing.Count -eq 0)
    $result.ready = ([bool]$result.visualHashAvailable -and [bool]$result.contentPdfAvailable)
    if (-not [bool]$result.ready) {
        $result.reason = $(if ($missing.Count -gt 0) { '同じレンダリング世代のcontent PDFが不足しています。' } else { 'content PDFが保存されていません。' })
    }
    return [pscustomobject]$result
}


function Get-LocalSnapshotSummaryCachePath([string]$Language, [string]$WorkbookId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    return (Join-Path $Script:LocalRuntimeCacheRoot ("snapshots-{0}-{1}.json" -f $Language, $safeWorkbookId))
}

function Clear-SnapshotSummaryCache([string]$Language, [string]$WorkbookId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $cacheKey = ($Language + '|' + $safeWorkbookId).ToLowerInvariant()
    [void]$Script:SnapshotSummaryCache.Remove($cacheKey)
    [void]$Script:SnapshotIdCache.Remove($cacheKey)
    try {
        $path = Get-LocalSnapshotSummaryCachePath $Language $safeWorkbookId
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
    } catch { }
}

function Get-SnapshotSummaryEntry([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $m = Get-SnapshotManifest $Language $safeWorkbookId $safeSnapshotId
    if ($null -eq $m) { return $null }
    $state = Get-SnapshotSourceState $Language $safeWorkbookId $safeSnapshotId
    $versions = @(Get-RenderVersionIds $Language $safeWorkbookId $safeSnapshotId)
    $preferred = ''; $visualAvailable = $false; $contentAvailable = $false
    $reason = 'PDF作成済みの比較可能な版がありません。'
    foreach ($versionId in @($versions | Sort-Object -Descending)) {
        $availability = Get-HistoryRenderVersionAvailability $Language $safeWorkbookId $safeSnapshotId ([string]$versionId)
        if ([bool]$availability.visualHashAvailable) { $visualAvailable = $true }
        if ([bool]$availability.contentPdfAvailable) { $contentAvailable = $true }
        if ([string]::IsNullOrWhiteSpace($preferred) -and [bool]$availability.ready) {
            $preferred = [string]$versionId; $reason = ''
        } elseif ([string]::IsNullOrWhiteSpace($preferred) -and -not [string]::IsNullOrWhiteSpace([string]$availability.reason)) {
            $reason = [string]$availability.reason
        }
    }
    return [pscustomobject][ordered]@{
        snapshotId = $safeSnapshotId; detectedAt = [string](Get-DataProperty $m 'detectedAt' '')
        captureReason = [string](Get-DataProperty $m 'captureReason' '')
        sourceHash = [string](Get-DataProperty $m 'sourceHash' '')
        sourceRetained = [bool]$state.sourceRetained
        pins = @(Get-SnapshotPins $Language $safeWorkbookId $safeSnapshotId)
        renderVersionIds = @($versions); visualCompareReady = (-not [string]::IsNullOrWhiteSpace($preferred))
        preferredVersionId = $preferred; visualHashAvailable = $visualAvailable
        contentPdfAvailable = $contentAvailable; unavailableReason = $reason
    }
}

function Update-SnapshotSummaryCacheEntry([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    # PDF作成やpin変更で影響する通常1版だけを差し替え、履歴画面で40版を再走査しない。
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $historyDir = Get-WorkbookHistoryDir $Language $safeWorkbookId
    if (-not (Test-Path -LiteralPath $historyDir)) { Clear-SnapshotSummaryCache $Language $safeWorkbookId; return $false }
    $cacheKey = ($Language + '|' + $safeWorkbookId).ToLowerInvariant()
    $record = $Script:SnapshotSummaryCache[$cacheKey]
    $localPath = Get-LocalSnapshotSummaryCachePath $Language $safeWorkbookId
    if ($null -eq $record) {
        try { if (Test-Path -LiteralPath $localPath) { $record = Read-JsonFile $localPath $null } } catch { $record = $null }
    }
    $entry = Get-SnapshotSummaryEntry $Language $safeWorkbookId $safeSnapshotId
    if ($null -eq $entry) { Clear-SnapshotSummaryCache $Language $safeWorkbookId; return $false }

    # 新規版追加や期限整理も反映するため、軽量な版ID一覧だけは1回確認する。
    [void]$Script:SnapshotIdCache.Remove($cacheKey)
    $recentIds = @((Get-SnapshotIds $Language $safeWorkbookId) | Select-Object -Last 40)
    $map = @{}
    foreach ($old in @(Get-Array (Get-DataProperty $record 'summaries' @()))) {
        $oldId = [string](Get-DataProperty $old 'snapshotId' '')
        if (-not [string]::IsNullOrWhiteSpace($oldId)) { $map[$oldId] = $old }
    }
    $map[$safeSnapshotId] = $entry
    $summaries = @()
    foreach ($id in @($recentIds | Sort-Object -Descending)) {
        $key = [string]$id
        if (-not $map.ContainsKey($key)) {
            Clear-SnapshotSummaryCache $Language $safeWorkbookId
            return $false
        }
        $summaries += $map[$key]
    }
    $stamp = [IO.Directory]::GetLastWriteTimeUtc($historyDir).Ticks
    $updated = [pscustomobject][ordered]@{ schemaVersion = 2; stamp = $stamp; summaries = @($summaries); savedAt = New-NowIso }
    $Script:SnapshotSummaryCache[$cacheKey] = $updated
    try { Write-JsonFile $localPath $updated } catch { }
    return $true
}

function Get-SnapshotSummaries([string]$Language, [string]$WorkbookId) {
    if ([string]::IsNullOrWhiteSpace($WorkbookId)) { return @() }
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $historyDir = Get-WorkbookHistoryDir $Language $safeWorkbookId
    if (-not (Test-Path -LiteralPath $historyDir)) { return @() }
    $stamp = [IO.Directory]::GetLastWriteTimeUtc($historyDir).Ticks
    $cacheKey = ($Language + '|' + $safeWorkbookId).ToLowerInvariant()
    $memory = $Script:SnapshotSummaryCache[$cacheKey]
    if ($null -ne $memory -and [Int64](Get-DataProperty $memory 'stamp' -1) -eq $stamp) {
        return @(Get-Array (Get-DataProperty $memory 'summaries' @()))
    }
    $localPath = Get-LocalSnapshotSummaryCachePath $Language $safeWorkbookId
    try {
        if (Test-Path -LiteralPath $localPath) {
            $local = Read-JsonFile $localPath $null
            if ($null -ne $local -and
                [int](Get-DataProperty $local 'schemaVersion' 0) -eq 2 -and
                [Int64](Get-DataProperty $local 'stamp' -1) -eq $stamp) {
                $summaries = @(Get-Array (Get-DataProperty $local 'summaries' @()))
                $Script:SnapshotSummaryCache[$cacheKey] = [pscustomobject][ordered]@{ stamp = $stamp; summaries = $summaries }
                return $summaries
            }
        }
    } catch { }
    $out = @()
    foreach ($id in @((Get-SnapshotIds $Language $safeWorkbookId) | Select-Object -Last 40)) {
        $entry = Get-SnapshotSummaryEntry $Language $safeWorkbookId ([string]$id)
        if ($null -ne $entry) { $out += $entry }
    }
    $summaries = @($out | Sort-Object { [string]$_.snapshotId } -Descending)
    $record = [pscustomobject][ordered]@{ schemaVersion = 2; stamp = $stamp; summaries = $summaries; savedAt = New-NowIso }
    $Script:SnapshotSummaryCache[$cacheKey] = $record
    try { Write-JsonFile $localPath $record } catch { }
    return $summaries
}

function Get-StoredComparison(
    [string]$Language,
    [string]$WorkbookId,
    [string]$FromSnapshotId,
    [string]$ToSnapshotId,
    [string]$FromVersionId = '',
    [string]$ToVersionId = '',
    [string]$Scope = 'history'
) {
    if ([string]::IsNullOrWhiteSpace($WorkbookId) -or [string]::IsNullOrWhiteSpace($FromSnapshotId) -or [string]::IsNullOrWhiteSpace($ToSnapshotId)) { return $null }
    $versions = if (-not [string]::IsNullOrWhiteSpace($ToVersionId)) { @($ToVersionId) } else { @(Get-RenderVersionIds $Language $WorkbookId $ToSnapshotId) }
    foreach ($ver in $versions) {
        $dir = Join-Path (Get-RenderRecordDir $Language $WorkbookId $ToSnapshotId ([string]$ver)) 'comparisons'
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue | Sort-Object LastWriteTimeUtc -Descending)) {
            try {
                $candidate = Read-JsonFile $f.FullName $null
                if ($null -eq $candidate) { continue }
                if ([string](Get-DataProperty $candidate 'scope' '') -ne $Scope) { continue }
                if ([string](Get-DataProperty $candidate 'baselineSnapshotId' '') -ne $FromSnapshotId) { continue }
                if ([string](Get-DataProperty $candidate 'currentSnapshotId' '') -ne $ToSnapshotId) { continue }
                if (-not [string]::IsNullOrWhiteSpace($FromVersionId) -and
                    [string](Get-DataProperty $candidate 'baselineVersionId' '') -ne $FromVersionId) { continue }
                if (-not [string]::IsNullOrWhiteSpace($ToVersionId) -and
                    [string](Get-DataProperty $candidate 'currentVersionId' '') -ne $ToVersionId) { continue }
                return $candidate
            } catch { }
        }
    }
    return $null
}

function Serve-HistoryContentPdf($Context, [string]$Language, [string]$WorkbookId, [string]$VersionId, [string]$SheetName, [string]$SnapshotId = '') {
    if ([string]::IsNullOrWhiteSpace($WorkbookId) -or [string]::IsNullOrWhiteSpace($SheetName)) {
        throw 'workbookId / sheetName が必要です。'
    }
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $safeSnapshotId = ''
    if (-not [string]::IsNullOrWhiteSpace($SnapshotId)) { $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId' }
    if ([string]::IsNullOrWhiteSpace($VersionId)) {
        if ([string]::IsNullOrWhiteSpace($safeSnapshotId)) { throw 'versionId または snapshotId が必要です。' }
        $VersionId = Get-PreferredHistoryRenderVersion $Language $safeWorkbookId $safeSnapshotId
        if ([string]::IsNullOrWhiteSpace($VersionId)) { throw '指定した検知版のPDFは保持されていません。' }
    }
    $safeVersionId = Assert-SafeStorageSegment $VersionId 'versionId'
    if (-not [string]::IsNullOrWhiteSpace($safeSnapshotId)) {
        # renders配下を毎回列挙せず、指定された版ディレクトリを直接確認する。
        $renderRecordDir = Get-RenderRecordDir $Language $safeWorkbookId $safeSnapshotId $safeVersionId
        if (-not (Test-Path -LiteralPath $renderRecordDir -PathType Container)) {
            throw '指定したPDF世代は、この履歴版に属していません。'
        }
    }
    $full = Resolve-ContentPdfSheetPathExact $Language $safeWorkbookId $safeVersionId $SheetName
    if ([string]::IsNullOrWhiteSpace($full)) { throw '指定した世代のPDFが見つかりません。' }
    # URLはsnapshot/version/sheetで不変。全バイトを先にメモリへ読むのをやめ、
    # PDF.jsが受信済みデータから解析を始められるようストリーミングする。
    Write-FileResponse $Context 200 $full 'application/pdf' $false 'private, max-age=31536000, immutable'
}


function Get-AutoStateSummary([string]$Language, [switch]$Fast) {
    $settings = Get-AutoRenderSettings
    $historySettings = Get-InputHistorySettings
    if ($Fast) {
        return [ordered]@{
            enabled = [bool]$settings.enabled; inputHistoryApproved = $true
            sourceRetentionApproved = [bool](Test-SourceRetentionEnabled)
            quietPeriodSeconds = [int]$settings.quietPeriodSeconds
            maxRetryCount = [int]$settings.maxRetryCount; retryBaseSeconds = [int]$settings.retryBaseSeconds
            minFreeMegabytes = [int]$settings.minFreeMegabytes; notifyOnCompletion = [bool]$settings.notifyOnCompletion; notifyOnFailure = [bool]$settings.notifyOnFailure
            schedulerRunning = $false; historySizeMb = 0
            softCapMegabytes = [int]$historySettings.softCapMegabytes; items = @()
        }
    }
    $items = @()
    try {
        $dir = Get-AutoStateDir $Language
        if (Test-Path -LiteralPath $dir) {
            foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue)) {
                $j = Read-JsonFile $f.FullName $null
                if ($null -eq $j) { continue }
                $itemState = [string](Get-DataProperty $j 'state' '')
                $lastResultAt = [string](Get-DataProperty $j 'lastResultAt' '')
                if ($itemState -eq 'idle' -and [string]::IsNullOrWhiteSpace($lastResultAt)) { continue }
                $items += [ordered]@{
                    workbookId = [string](Get-DataProperty $j 'workbookId' $f.BaseName)
                    state = $itemState
                    deferReason = [string](Get-DataProperty $j 'deferReason' '')
                    quietDeadline = [string](Get-DataProperty $j 'quietDeadline' '')
                    firstDetectedAt = [string](Get-DataProperty $j 'firstDetectedAt' '')
                    ownerPcName = [string](Get-DataProperty $j 'ownerPcName' '')
                    retryCount = [int](Get-DataProperty $j 'retryCount' 0)
                    nextRetryAt = [string](Get-DataProperty $j 'nextRetryAt' '')
                    lastError = [string](Get-DataProperty $j 'lastError' '')
                    lastResult = [string](Get-DataProperty $j 'lastResult' '')
                    lastResultAt = $lastResultAt
                    notificationId = [string](Get-DataProperty $j 'notificationId' '')
                }
            }
        }
    } catch { }
    return [ordered]@{
        enabled = [bool]$settings.enabled; inputHistoryApproved = $true
        sourceRetentionApproved = [bool](Test-SourceRetentionEnabled)
        quietPeriodSeconds = [int]$settings.quietPeriodSeconds
        maxRetryCount = [int]$settings.maxRetryCount; retryBaseSeconds = [int]$settings.retryBaseSeconds
        minFreeMegabytes = [int]$settings.minFreeMegabytes; notifyOnCompletion = [bool]$settings.notifyOnCompletion; notifyOnFailure = [bool]$settings.notifyOnFailure
        schedulerRunning = (Test-AutoSchedulerProcessRunning)
        historySizeMb = (Get-InputHistorySizeMb $Language)
        softCapMegabytes = [int]$historySettings.softCapMegabytes; items = @($items)
    }
}

function Request-AutoRunNow([string]$Language, [string]$WorkbookId) {
    # 延期中のブックを即時実行する。スケジューラー子プロセスが次の tick で拾う。
    if ([string]::IsNullOrWhiteSpace($WorkbookId)) { throw 'workbookId が必要です。' }
    $state = Read-AutoState $Language $WorkbookId
    if ($null -eq $state) { $state = New-AutoState $WorkbookId }
    Set-NoteProperty $state 'quietDeadline' ([DateTime]::UtcNow.AddSeconds(-1).ToString('o'))
    Set-NoteProperty $state 'stableCount' 99
    Set-NoteProperty $state 'deferReason' ''
    Set-NoteProperty $state 'state' 'waiting'
    Set-NoteProperty $state 'retryCount' 0
    Set-NoteProperty $state 'nextRetryAt' ''
    Set-NoteProperty $state 'lastError' ''
    Write-AutoState $Language $WorkbookId $state
    return [ordered]@{ workbookId = $WorkbookId; requested = $true }
}

# =====================================================================
# V5 差分詳細・視覚比較
# =====================================================================

$Script:DiffDetailAlgorithmVersion = 23
$Script:DiffDetailDpi = 120
$Script:DiffDetailThreshold = 24
$Script:DiffDetailMinimumRegionPixels = 24
$Script:DiffDetailPadding = 5
# content PDFの版ディレクトリは生成完了後は不変。シート名索引を共有し、
# 同じネットワークフォルダーをシート数分だけ再列挙しない。
$Script:ContentPdfSheetIndexCache = @{}
$Script:ContentPdfSheetIndexCacheLimit = 64

function Get-DiffSnapshotDate([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$Fallback = '') {
    try {
        $m = Get-SnapshotManifest $Language $WorkbookId $SnapshotId
        $detected = [string](Get-DataProperty $m 'detectedAt' '')
        if (-not [string]::IsNullOrWhiteSpace($detected)) { return $detected }
    } catch { }
    return $Fallback
}


function Get-PreferredHistoryRenderVersion([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    foreach ($versionId in @((Get-RenderVersionIds $Language $WorkbookId $SnapshotId) | Sort-Object -Descending)) {
        $availability = Get-HistoryRenderVersionAvailability $Language $WorkbookId $SnapshotId ([string]$versionId)
        if ([bool]$availability.ready) { return [string]$versionId }
    }
    return ''
}


function New-HistoricalSnapshotComparison(
    [string]$Language,
    [string]$WorkbookId,
    [string]$BaselineSnapshotId,
    [string]$BaselineVersionId,
    [string]$CurrentSnapshotId,
    [string]$CurrentVersionId
) {
    $base = Get-VisualHashes $Language $WorkbookId $BaselineSnapshotId $BaselineVersionId
    $current = Get-VisualHashes $Language $WorkbookId $CurrentSnapshotId $CurrentVersionId
    $result = [ordered]@{
        schemaVersion = 3
        status = 'unavailable'
        scope = 'history'
        baselineSnapshotId = $BaselineSnapshotId
        baselineVersionId = $BaselineVersionId
        currentSnapshotId = $CurrentSnapshotId
        currentVersionId = $CurrentVersionId
        comparedAt = New-NowIso
        method = 'stored-hash-history'
        confidence = 1.0
        sourceType = Get-ComparisonSourceType $Language $WorkbookId
        unitMappings = @()
        mappingConfidence = 1.0
        changedSheets = @()
        unchangedSheets = @()
        unknownSheets = @()
        addedSheets = @()
        removedSheets = @()
        message = ''
    }
    if ($null -eq $base -or $null -eq $current) {
        $result.message = '選択した版の画像ハッシュが保存されていないため比較できません。'
        return [pscustomobject]$result
    }
    $baseEnvironment = [string](Get-DataProperty $base 'renderEnvironmentFingerprint' '')
    $currentEnvironment = [string](Get-DataProperty $current 'renderEnvironmentFingerprint' '')
    if (-not [string]::IsNullOrWhiteSpace($baseEnvironment) -and
        -not [string]::IsNullOrWhiteSpace($currentEnvironment) -and
        $baseEnvironment -ne $currentEnvironment) {
        $result.method = 'stored-hash-history-environment-mismatch'
        $result.confidence = 0.65
        $result.message = '2版のPDF作成環境が異なります。表示結果を目視で確認してください。'
    }
    $mappings = Get-ComparisonUnitMappings (Get-DataProperty $base 'sheets' @()) (Get-DataProperty $current 'sheets' @()) ([string]$result.sourceType)
    $result.status = 'complete'
    Set-ComparisonUnitMappingResult $result $mappings
    return [pscustomobject]$result
}


function Save-HistoricalSnapshotComparison([string]$Language, [string]$WorkbookId, $Comparison) {
    if ($null -eq $Comparison -or [string](Get-DataProperty $Comparison 'status' '') -ne 'complete') { throw '保存できる履歴比較結果がありません。' }
    foreach ($field in @('baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
        if ([string]::IsNullOrWhiteSpace([string](Get-DataProperty $Comparison $field ''))) { throw "履歴比較結果の識別子が不足しています: $field" }
    }
    if ([string](Get-DataProperty $Comparison 'scope' '') -ne 'history') { throw '履歴比較以外はこの保存経路を使用できません。' }
    $currentSnapshotId = [string]$Comparison.currentSnapshotId
    $currentVersionId = [string]$Comparison.currentVersionId
    $dir = Join-Path (Get-RenderRecordDir $Language $WorkbookId $currentSnapshotId $currentVersionId) 'comparisons'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $key = (Get-Sha256Text ("history|{0}|{1}|{2}|{3}" -f $Comparison.baselineSnapshotId, $Comparison.baselineVersionId, $currentSnapshotId, $currentVersionId)).Substring(7, 16)
    $path = Join-Path $dir ("hcmp-{0}.json" -f $key)
    Write-JsonFile $path $Comparison
    $saved = Read-JsonFile $path $null
    if ($null -eq $saved) { throw '履歴比較結果を保存後に再読込できませんでした。' }
    foreach ($field in @('scope','baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
        if ([string](Get-DataProperty $saved $field '') -ne [string](Get-DataProperty $Comparison $field '')) {
            throw "履歴比較結果の保存検証に失敗しました: $field"
        }
    }
    Write-HistoryEvent $Language 'compare.history.created' ([ordered]@{
        workbookId = $WorkbookId
        baselineSnapshotId = [string]$Comparison.baselineSnapshotId
        baselineVersionId = [string]$Comparison.baselineVersionId
        currentSnapshotId = $currentSnapshotId
        currentVersionId = $currentVersionId
        changed = @(Get-Array $Comparison.changedSheets)
        unknown = @(Get-Array $Comparison.unknownSheets)
    })
    return $saved
}


function Get-DiffDetailContext(
    [string]$Language,
    [string]$WorkbookId,
    [string]$BaselineSnapshotId = '',
    [string]$CurrentSnapshotId = ''
) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $structure = Get-Structure $Language
    $matches = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $safeWorkbookId } | Select-Object -First 1)
    if ($matches.Count -eq 0) { throw '登録済み原稿が見つかりません。' }
    $workbook = $matches[0]
    $displayName = [string](Get-DataProperty $workbook 'displayName' '')
    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = [string](Get-DataProperty $workbook 'fileName' $safeWorkbookId) }
    $baseResult = [ordered]@{
        available = $false; status = 'unavailable'; message = ''
        workbookId = $safeWorkbookId; workbookName = $displayName; workbook = $workbook
        comparison = $null; comparisonPersisted = $false
        currentSnapshotId = ''; currentVersionId = ''; baselineSnapshotId = ''; baselineVersionId = ''
        currentVisualHashes = $null; baselineVisualHashes = $null
        currentAt = ''; baselineAt = ''; method = ''; scope = 'automatic'
        sourceType = [string](Get-DataProperty $workbook 'sourceType' 'excel')
    }
    if (-not (Test-InputHistoryEnabled)) {
        $baseResult.message = '入力履歴が無効なため、差分詳細は利用できません。'
        return [pscustomobject]$baseResult
    }
    $hasBaseline = -not [string]::IsNullOrWhiteSpace($BaselineSnapshotId)
    $hasCurrent = -not [string]::IsNullOrWhiteSpace($CurrentSnapshotId)
    if ($hasBaseline -xor $hasCurrent) { throw '履歴比較では比較元と比較先の両方を指定してください。' }
    if ($hasBaseline -and $hasCurrent) {
        $baselineSnapshotId = Assert-SafeStorageSegment $BaselineSnapshotId 'baselineSnapshotId'
        $currentSnapshotId = Assert-SafeStorageSegment $CurrentSnapshotId 'currentSnapshotId'
        $baseResult.scope = 'history'
        if ($baselineSnapshotId -eq $currentSnapshotId) { $baseResult.message = '異なる2版を選択してください。'; return [pscustomobject]$baseResult }
        if ($null -eq (Get-SnapshotManifest $Language $safeWorkbookId $baselineSnapshotId) -or
            $null -eq (Get-SnapshotManifest $Language $safeWorkbookId $currentSnapshotId)) {
            $baseResult.message = '選択した履歴版が見つかりません。保存期限または履歴整理を確認してください。'
            return [pscustomobject]$baseResult
        }
        $baselineVersionId = Get-PreferredHistoryRenderVersion $Language $safeWorkbookId $baselineSnapshotId
        $currentVersionId = Get-PreferredHistoryRenderVersion $Language $safeWorkbookId $currentSnapshotId
        if ([string]::IsNullOrWhiteSpace($baselineVersionId) -or [string]::IsNullOrWhiteSpace($currentVersionId)) {
            $baseResult.message = '選択した2版は、画像ハッシュと同一世代のcontent PDFがそろっていないため視覚比較できません。'
            return [pscustomobject]$baseResult
        }
        $comparison = Get-StoredComparison $Language $safeWorkbookId $baselineSnapshotId $currentSnapshotId $baselineVersionId $currentVersionId 'history'
        if ($null -eq $comparison) {
            $comparison = New-HistoricalSnapshotComparison $Language $safeWorkbookId $baselineSnapshotId $baselineVersionId $currentSnapshotId $currentVersionId
        } else { $baseResult.comparisonPersisted = $true }
        if ($null -eq $comparison -or [string](Get-DataProperty $comparison 'status' '') -ne 'complete') {
            $baseResult.message = [string](Get-DataProperty $comparison 'message' '選択した2版を比較できません。')
            return [pscustomobject]$baseResult
        }
        $baseResult.available = $true; $baseResult.status = 'available'; $baseResult.comparison = $comparison
        $baseResult.currentSnapshotId = $currentSnapshotId; $baseResult.currentVersionId = $currentVersionId
        $baseResult.baselineSnapshotId = $baselineSnapshotId; $baseResult.baselineVersionId = $baselineVersionId
        $baseResult.currentAt = Get-DiffSnapshotDate $Language $safeWorkbookId $currentSnapshotId ''
        $baseResult.baselineAt = Get-DiffSnapshotDate $Language $safeWorkbookId $baselineSnapshotId ''
        $baseResult.method = [string](Get-DataProperty $comparison 'method' 'stored-hash-history')
        return [pscustomobject]$baseResult
    }
    $status = [string](Get-DataProperty $workbook 'status' '')
    $currentHash = Normalize-FileHash ([string](Get-DataProperty $workbook 'currentExcelHash' ''))
    $renderedHash = Normalize-FileHash ([string](Get-DataProperty $workbook 'lastRenderedExcelHash' ''))
    if ([string]::IsNullOrWhiteSpace($renderedHash) -or $status -in @('new','excel-updated','source-updated','render-error','rendering') -or
        ((-not [string]::IsNullOrWhiteSpace($currentHash)) -and $currentHash -ne $renderedHash)) {
        $baseResult.message = '最新のPDFを作成してから差分を確認してください。'; return [pscustomobject]$baseResult
    }
    $currentSnapshotId = [string](Get-DataProperty $workbook 'lastRenderedSnapshotId' '')
    $currentVersionId = [string](Get-DataProperty $workbook 'lastRenderedVersionId' '')
    if ([string]::IsNullOrWhiteSpace($currentSnapshotId) -or [string]::IsNullOrWhiteSpace($currentVersionId)) {
        $baseResult.message = '現在版を一意に識別できないため、差分詳細を開けません。'; return [pscustomobject]$baseResult
    }
    $currentSnapshotId = Assert-SafeStorageSegment $currentSnapshotId 'currentSnapshotId'
    $currentVersionId = Assert-SafeStorageSegment $currentVersionId 'currentVersionId'
    $comparison = Get-LatestComparison $Language $safeWorkbookId $workbook
    if ($null -eq $comparison) { $baseResult.message = '保存済みの比較結果がありません。PDFを再作成してください。'; return [pscustomobject]$baseResult }
    $baselineSnapshotId = [string](Get-DataProperty $comparison 'baselineSnapshotId' '')
    $baselineVersionId = [string](Get-DataProperty $comparison 'baselineVersionId' '')
    if ([string](Get-DataProperty $comparison 'status' '') -ne 'complete' -or
        [string]::IsNullOrWhiteSpace($baselineSnapshotId) -or [string]::IsNullOrWhiteSpace($baselineVersionId)) {
        $baseResult.message = [string](Get-DataProperty $comparison 'message' '前回の比較基準がありません。'); return [pscustomobject]$baseResult
    }
    $baselineSnapshotId = Assert-SafeStorageSegment $baselineSnapshotId 'baselineSnapshotId'
    $baselineVersionId = Assert-SafeStorageSegment $baselineVersionId 'baselineVersionId'
    # 自動比較の版は作成時に全PDFを検証・固定済み。詳細を開くたびに全シートを
    # 再列挙せず、軽量な版ディレクトリとハッシュだけを確認する。
    # 個別PDFは表示要求時に Serve-HistoryContentPdf が厳密に検証する。
    $currentHashes = Get-VisualHashes $Language $safeWorkbookId $currentSnapshotId $currentVersionId
    $baselineHashes = Get-VisualHashes $Language $safeWorkbookId $baselineSnapshotId $baselineVersionId
    if ($null -eq $currentHashes -or $null -eq $baselineHashes) {
        $baseResult.message = '自動比較に使った画像ハッシュが保持されていません。PDFを再作成してください。'
        return [pscustomobject]$baseResult
    }
    $baseResult.currentVisualHashes = $currentHashes
    $baseResult.baselineVisualHashes = $baselineHashes
    $baseResult.available = $true; $baseResult.status = 'available'; $baseResult.comparison = $comparison; $baseResult.comparisonPersisted = $true
    $baseResult.currentSnapshotId = $currentSnapshotId; $baseResult.currentVersionId = $currentVersionId
    $baseResult.baselineSnapshotId = $baselineSnapshotId; $baseResult.baselineVersionId = $baselineVersionId
    $baseResult.currentAt = [string](Get-DataProperty $workbook 'lastRenderedAt' ([string](Get-DataProperty $comparison 'comparedAt' '')))
    $baseResult.baselineAt = [string](Get-DataProperty $comparison 'baselineAt' '')
    $baseResult.method = [string](Get-DataProperty $comparison 'method' '')
    return [pscustomobject]$baseResult
}


function Get-DiffPairKey($Context) {
    return (Get-Sha256Text ("{0}|{1}|{2}|{3}|{4}|{5}|{6}" -f
        [string]$Context.scope, [string]$Context.workbookId,
        [string]$Context.baselineSnapshotId, [string]$Context.baselineVersionId,
        [string]$Context.currentSnapshotId, [string]$Context.currentVersionId,
        $Script:DiffDetailAlgorithmVersion)).Substring(7, 24)
}

function Get-DiffReviewStatePath([string]$Language, $Context) {
    $safeWorkbookId = Assert-SafeStorageSegment ([string]$Context.workbookId) 'workbookId'
    $reviewRoot = Join-Path $Script:LocalConfigRoot 'diff-reviews'
    return (Join-Path $reviewRoot ("review-{0}-{1}-{2}.json" -f $Language, $safeWorkbookId, (Get-DiffPairKey $Context)))
}

function Get-DiffReviewStateForContext([string]$Language, $Context) {
    $empty = [pscustomobject][ordered]@{
        schemaVersion = 1
        workbookId = [string]$Context.workbookId
        scope = [string]$Context.scope
        baselineSnapshotId = [string]$Context.baselineSnapshotId
        baselineVersionId = [string]$Context.baselineVersionId
        currentSnapshotId = [string]$Context.currentSnapshotId
        currentVersionId = [string]$Context.currentVersionId
        algorithmVersion = $Script:DiffDetailAlgorithmVersion
        confirmedSheetKeys = @()
        reviewedAt = ''
        reviewedBy = ''
    }
    if (-not [bool]$Context.available) { return $empty }
    $path = Get-DiffReviewStatePath $Language $Context
    if (-not (Test-Path -LiteralPath $path)) { return $empty }
    try { $saved = Read-JsonFile $path $null } catch { return $empty }
    if ($null -eq $saved -or (Get-IntDataProperty $saved 'algorithmVersion' 0) -ne $Script:DiffDetailAlgorithmVersion) { return $empty }
    foreach ($field in @('workbookId','scope','baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
        if ([string](Get-DataProperty $saved $field '') -ne [string](Get-DataProperty $empty $field '')) { return $empty }
    }
    $empty.confirmedSheetKeys = @(Get-Array (Get-DataProperty $saved 'confirmedSheetKeys' @()) | ForEach-Object { [string]$_ } | Where-Object { $_ } | Select-Object -Unique)
    $empty.reviewedAt = [string](Get-DataProperty $saved 'reviewedAt' '')
    $empty.reviewedBy = [string](Get-DataProperty $saved 'reviewedBy' '')
    return $empty
}

function Set-DiffReviewState([string]$Language, $Body) {
    $workbookId = Assert-SafeStorageSegment ([string](Get-DataProperty $Body 'workbookId' '')) 'workbookId'
    $fromSnapshotId = [string](Get-DataProperty $Body 'fromSnapshotId' '')
    $toSnapshotId = [string](Get-DataProperty $Body 'toSnapshotId' '')
    $context = Get-DiffDetailContext $Language $workbookId $fromSnapshotId $toSnapshotId
    if (-not [bool]$context.available) { throw [ArgumentException]::new([string](Get-DataProperty $context 'message' '比較対象を確認できません。')) }
    foreach ($field in @('baselineVersionId','currentVersionId')) {
        $requested = [string](Get-DataProperty $Body $field '')
        if (-not [string]::IsNullOrWhiteSpace($requested) -and $requested -ne [string](Get-DataProperty $context $field '')) {
            throw [ArgumentException]::new('比較対象が更新されました。差分画面を開き直してください。')
        }
    }
    $sheetKey = Assert-SafeStorageSegment ([string](Get-DataProperty $Body 'sheetKey' '')) 'sheetKey'
    $detail = New-DiffDetailSkeleton $Language $context
    if (@(Get-Array $detail.sheets | Where-Object { [string]$_.sheetKey -eq $sheetKey }).Count -eq 0) {
        throw [ArgumentException]::new('確認対象のページ項目が見つかりません。')
    }
    $review = Get-DiffReviewStateForContext $Language $context
    $keys = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($existing in @(Get-Array $review.confirmedSheetKeys)) { if ($existing) { [void]$keys.Add([string]$existing) } }
    $confirmed = [bool](Get-DataProperty $Body 'confirmed' $true)
    if ($confirmed) { [void]$keys.Add($sheetKey) } else { [void]$keys.Remove($sheetKey) }
    $review.confirmedSheetKeys = @($keys | Sort-Object)
    $review.reviewedAt = New-NowIso
    $review.reviewedBy = [string]$env:USERNAME
    $path = Get-DiffReviewStatePath $Language $context
    $parent = Split-Path -Parent $path
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    Write-JsonFile $path $review
    $cachePath = Get-LocalDiffDetailCachePath $Language $workbookId ([string]$context.baselineSnapshotId) ([string]$context.currentSnapshotId) ([string]$context.baselineVersionId) ([string]$context.currentVersionId) ([string]$context.scope)
    $cacheKey = ([IO.Path]::GetFullPath($cachePath)).ToLowerInvariant()
    [void]$Script:DiffDetailResponseCache.Remove($cacheKey)
    if (Test-Path -LiteralPath $cachePath) { Remove-Item -LiteralPath $cachePath -Force -ErrorAction SilentlyContinue }
    return $review
}

function Get-DiffLaunchLockPath([string]$Language, $Context) {
    return (Join-Path (Get-WorkspacePath $Language) ("locks\diff-launch_{0}.lock" -f (Get-DiffPairKey $Context)))
}
function Get-DiffGenerationLockPath([string]$Language, $Context) {
    return (Join-Path (Get-WorkspacePath $Language) ("locks\diff-generate_{0}.lock" -f (Get-DiffPairKey $Context)))
}

function Get-DiffDetailCacheDir($Context) {
    $record = Get-RenderRecordDir (Get-EffectiveLanguage) ([string]$Context.workbookId) ([string]$Context.currentSnapshotId) ([string]$Context.currentVersionId)
    $baseline = Assert-SafeStorageSegment ([string]$Context.baselineSnapshotId) 'baselineSnapshotId'
    $cacheKey = (Get-Sha256Text ("{0}|{1}|{2}|{3}" -f [string]$Context.scope, $baseline, [string]$Context.baselineVersionId, $Script:DiffDetailAlgorithmVersion)).Substring(7, 16)
    return (Join-Path $record (Join-Path 'comparisons' ("d{0}" -f $cacheKey)))
}

function Get-DiffDetailCacheDirForLanguage([string]$Language, $Context) {
    $record = Get-RenderRecordDir $Language ([string]$Context.workbookId) ([string]$Context.currentSnapshotId) ([string]$Context.currentVersionId)
    $baseline = Assert-SafeStorageSegment ([string]$Context.baselineSnapshotId) 'baselineSnapshotId'
    $cacheKey = (Get-Sha256Text ("{0}|{1}|{2}|{3}" -f [string]$Context.scope, $baseline, [string]$Context.baselineVersionId, $Script:DiffDetailAlgorithmVersion)).Substring(7, 16)
    return (Join-Path $record (Join-Path 'comparisons' ("d{0}" -f $cacheKey)))
}

function Get-DiffSheetKey([string]$SheetName) {
    $hash = Get-Sha256Text $SheetName
    return ('s-' + $hash.Substring(7, 10))
}

function Test-DiffDetailMatchesContext($Detail, $Context) {
    if ($null -eq $Detail -or $null -eq $Context) { return $false }
    if ((Get-IntDataProperty $Detail 'algorithmVersion' 0) -ne $Script:DiffDetailAlgorithmVersion) { return $false }
    if ([string](Get-DataProperty $Detail 'workbookId' '') -ne [string]$Context.workbookId) { return $false }
    $comparison = Get-DataProperty $Detail 'comparison' $null
    if ($null -eq $comparison) { return $false }
    foreach ($field in @('scope','baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
        if ([string](Get-DataProperty $comparison $field '') -ne [string](Get-DataProperty $Context $field '')) { return $false }
    }
    return $true
}

function Get-DiffHashSheetMap($Hashes) {
    $map = @{}
    if ($null -eq $Hashes) { return $map }
    foreach ($sheet in @(Get-Array (Get-DataProperty $Hashes 'sheets' @()))) {
        $map[[string](Get-DataProperty $sheet 'sheetName' '')] = $sheet
    }
    return $map
}

function New-DiffDetailSkeleton([string]$Language, $Context) {
    $comparison = $Context.comparison
    $addedSet = @{}
    $removedSet = @{}
    $unknownSet = @{}
    foreach ($name in @(Get-Array (Get-DataProperty $comparison 'addedSheets' @()))) { $addedSet[[string]$name] = $true }
    foreach ($name in @(Get-Array (Get-DataProperty $comparison 'removedSheets' @()))) { $removedSet[[string]$name] = $true }
    foreach ($name in @(Get-Array (Get-DataProperty $comparison 'unknownSheets' @()))) { $unknownSet[[string]$name] = $true }

    $items = @(); $seen = @()
    $mappingRecords = @(Get-Array (Get-DataProperty $comparison 'unitMappings' @()))
    if ($mappingRecords.Count -eq 0) {
        foreach ($group in @(
            [ordered]@{ kind = 'modified'; names = @(Get-Array (Get-DataProperty $comparison 'changedSheets' @())) },
            [ordered]@{ kind = 'added'; names = @(Get-Array (Get-DataProperty $comparison 'addedSheets' @())) },
            [ordered]@{ kind = 'removed'; names = @(Get-Array (Get-DataProperty $comparison 'removedSheets' @())) },
            [ordered]@{ kind = 'unknown'; names = @(Get-Array (Get-DataProperty $comparison 'unknownSheets' @())) },
            [ordered]@{ kind = 'unchanged'; names = @(Get-Array (Get-DataProperty $comparison 'unchangedSheets' @())) }
        )) {
            foreach ($rawName in @($group.names)) {
                $name = [string]$rawName
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $mappingRecords += [pscustomobject][ordered]@{ beforeSheetName = $(if ([string]$group.kind -ne 'added') { $name } else { '' }); afterSheetName = $(if ([string]$group.kind -ne 'removed') { $name } else { '' }); displayName = $name; kind = [string]$group.kind; matchConfidence = 1.0; matchMethod = 'legacy-name'; message = '' }
            }
        }
    }
    foreach ($mapping in $mappingRecords) {
        $beforeName = [string](Get-DataProperty $mapping 'beforeSheetName' '')
        $afterName = [string](Get-DataProperty $mapping 'afterSheetName' '')
        $name = [string](Get-DataProperty $mapping 'displayName' $(if ($afterName) { $afterName } else { $beforeName }))
        $kind = [string](Get-DataProperty $mapping 'kind' 'unknown')
        $identity = "$beforeName`n$afterName`n$kind"
        if ([string]::IsNullOrWhiteSpace($name) -or $seen -contains $identity) { continue }
        $seen += $identity
        $mappingMessage = [string](Get-DataProperty $mapping 'message' '')
        $item = [pscustomobject][ordered]@{
            sheetName = $name
            beforeSheetName = $beforeName
            afterSheetName = $afterName
            sheetKey = Get-DiffSheetKey $identity
            kind = $kind
            matchConfidence = [double](Get-DataProperty $mapping 'matchConfidence' 1)
            matchMethod = [string](Get-DataProperty $mapping 'matchMethod' '')
            beforePages = 0
            afterPages = 0
            pageCount = 0
            unchangedPageNumbers = @()
            regionCount = 0
            # PDFはそのまま保持し、表示中のページだけをブラウザで描画・比較する。
            status = 'ready'
            message = $(if ($mappingMessage) { $mappingMessage } elseif ($kind -eq 'unknown') { '対応を確定できないため、左右のPDFを目視確認してください。' } else { '表示したページをブラウザで比較します。' })
            confirmed = $false
            pages = @()
        }
        if ($kind -eq 'unknown') { $item.status = 'unknown' }
        $items += $item
    }
    $currentHashes = Get-DataProperty $Context 'currentVisualHashes' $null
    $baselineHashes = Get-DataProperty $Context 'baselineVisualHashes' $null
    # 自動比較ではContextで読んだハッシュを再利用する。履歴比較など未設定の経路だけ
    # ここで1回読み、同じ共有JSONへの重複アクセスを避ける。
    if ([bool]$Context.available) {
        if ($null -eq $currentHashes) {
            $currentHashes = Get-VisualHashes $Language ([string]$Context.workbookId) ([string]$Context.currentSnapshotId) ([string]$Context.currentVersionId)
        }
        if ($null -eq $baselineHashes) {
            $baselineHashes = Get-VisualHashes $Language ([string]$Context.workbookId) ([string]$Context.baselineSnapshotId) ([string]$Context.baselineVersionId)
        }
    }
    $currentMap = Get-DiffHashSheetMap $currentHashes
    $baselineMap = Get-DiffHashSheetMap $baselineHashes
    foreach ($item in $items) {
        $beforeSheet = $null
        $afterSheet = $null
        if ($baselineMap.ContainsKey([string]$item.beforeSheetName)) {
            $beforeSheet = $baselineMap[[string]$item.beforeSheetName]
            $item.beforePages = Get-IntDataProperty $beforeSheet 'pageCount' 0
        }
        if ($currentMap.ContainsKey([string]$item.afterSheetName)) {
            $afterSheet = $currentMap[[string]$item.afterSheetName]
            $item.afterPages = Get-IntDataProperty $afterSheet 'pageCount' 0
        }
        $item.pageCount = [Math]::Max([int]$item.beforePages, [int]$item.afterPages)

        # 120 DPI・RGBの正規化画素ハッシュが完全一致するページは、差分領域解析を省略できる。
        # 知覚ハッシュは使わずSHA-256の完全一致だけを採用するため、変更の見落としは発生しない。
        $beforeHashes = @(Get-Array (Get-DataProperty $beforeSheet 'pageHashes' @()))
        $afterHashes = @(Get-Array (Get-DataProperty $afterSheet 'pageHashes' @()))
        $samePages = @()
        $comparablePages = [Math]::Min($beforeHashes.Count, $afterHashes.Count)
        for ($pageIndex = 0; $pageIndex -lt $comparablePages; $pageIndex++) {
            $beforeHash = Normalize-FileHash ([string]$beforeHashes[$pageIndex])
            $afterHash = Normalize-FileHash ([string]$afterHashes[$pageIndex])
            if (-not [string]::IsNullOrWhiteSpace($beforeHash) -and $beforeHash -eq $afterHash) {
                $samePages += ($pageIndex + 1)
            }
        }
        $item.unchangedPageNumbers = @($samePages)
    }
    $review = Get-DiffReviewStateForContext $Language $Context
    $confirmedSheetKeys = @{}
    foreach ($sheetKey in @(Get-Array (Get-DataProperty $review 'confirmedSheetKeys' @()))) { $confirmedSheetKeys[[string]$sheetKey] = $true }
    foreach ($item in $items) { $item.confirmed = $confirmedSheetKeys.ContainsKey([string]$item.sheetKey) }
    $modifiedCount = @($items | Where-Object { [string]$_.kind -eq 'modified' }).Count
    $addedCount = @($items | Where-Object { [string]$_.kind -eq 'added' }).Count
    $removedCount = @($items | Where-Object { [string]$_.kind -eq 'removed' }).Count
    $unknownCount = @($items | Where-Object { [string]$_.kind -eq 'unknown' }).Count
    $unchangedCount = @($items | Where-Object { [string]$_.kind -eq 'unchanged' }).Count
    return [pscustomobject][ordered]@{
        schemaVersion = 1
        algorithmVersion = $Script:DiffDetailAlgorithmVersion
        status = $(if ([bool]$Context.available) { 'ready' } else { 'unavailable' })
        message = [string]$Context.message
        workbookId = [string]$Context.workbookId
        workbookName = [string]$Context.workbookName
        sourceType = [string](Get-DataProperty $Context 'sourceType' 'excel')
        unitLabel = $(if ([string](Get-DataProperty $Context 'sourceType' 'excel') -eq 'powerpoint') { 'スライド' } elseif ([string](Get-DataProperty $Context 'sourceType' 'excel') -in @('word','pdf')) { 'ページ' } else { 'シート' })
        comparison = [ordered]@{
            baselineSnapshotId = [string]$Context.baselineSnapshotId
            baselineVersionId = [string]$Context.baselineVersionId
            currentSnapshotId = [string]$Context.currentSnapshotId
            currentVersionId = [string]$Context.currentVersionId
            baselineAt = [string]$Context.baselineAt
            currentAt = [string]$Context.currentAt
            comparedAt = [string](Get-DataProperty $comparison 'comparedAt' '')
            method = [string]$Context.method
            scope = [string]$Context.scope
            confidence = [double](Get-DataProperty $comparison 'confidence' 1.0)
        }
        summary = [ordered]@{
            changed = $modifiedCount
            added = $addedCount
            removed = $removedCount
            unknown = $unknownCount
            unchanged = $unchangedCount
            confirmed = @($items | Where-Object { [bool]$_.confirmed }).Count
        }
        review = $review
        generation = [ordered]@{ status = 'completed'; jobId = ''; percent = 100; message = '表示ページをブラウザで比較します。'; currentSheet = '' }
        sheets = @($items)
        generatedAt = ''
    }
}

function Get-LocalDiffDetailCachePath(
    [string]$Language,
    [string]$WorkbookId,
    [string]$BaselineSnapshotId = '',
    [string]$CurrentSnapshotId = '',
    [string]$BaselineVersionId = '',
    [string]$CurrentVersionId = '',
    [string]$Scope = ''
) {
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $scope = $(if (-not [string]::IsNullOrWhiteSpace($Scope)) { $Scope } elseif (-not [string]::IsNullOrWhiteSpace($BaselineSnapshotId) -and -not [string]::IsNullOrWhiteSpace($CurrentSnapshotId)) { 'history' } else { 'automatic' })
    $pair = (Get-Sha256Text ("{0}|{1}|{2}|{3}|{4}|{5}" -f
        $scope, $BaselineSnapshotId, $BaselineVersionId,
        $CurrentSnapshotId, $CurrentVersionId, $Script:DiffDetailAlgorithmVersion)).Substring(7, 24)
    return (Join-Path $Script:LocalRuntimeCacheRoot ("diff-{0}-{1}-{2}.json" -f $Language, $safeWorkbookId, $pair))
}

function Get-LocalDiffDetailCacheIdentity([string]$Language, [string]$WorkbookId, [string]$BaselineSnapshotId = '', [string]$CurrentSnapshotId = '') {
    $hasBaseline = -not [string]::IsNullOrWhiteSpace($BaselineSnapshotId)
    $hasCurrent = -not [string]::IsNullOrWhiteSpace($CurrentSnapshotId)
    if ($hasBaseline -xor $hasCurrent) { return $null }
    if ($hasBaseline -and $hasCurrent) {
        $baselineVersionId = Get-PreferredHistoryRenderVersion $Language $WorkbookId $BaselineSnapshotId
        $currentVersionId = Get-PreferredHistoryRenderVersion $Language $WorkbookId $CurrentSnapshotId
        if ([string]::IsNullOrWhiteSpace($baselineVersionId) -or [string]::IsNullOrWhiteSpace($currentVersionId)) { return $null }
        return [pscustomobject][ordered]@{
            scope = 'history'
            baselineSnapshotId = $BaselineSnapshotId; baselineVersionId = $baselineVersionId
            currentSnapshotId = $CurrentSnapshotId; currentVersionId = $currentVersionId
        }
    }
    $structure = Get-Structure $Language
    $workbook = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
    if ($workbook.Count -eq 0) { return $null }
    $comparison = Get-LatestComparison $Language $WorkbookId $workbook[0]
    if ($null -eq $comparison -or [string](Get-DataProperty $comparison 'status' '') -ne 'complete') { return $null }
    $identity = [pscustomobject][ordered]@{
        scope = 'automatic'
        baselineSnapshotId = [string](Get-DataProperty $comparison 'baselineSnapshotId' '')
        baselineVersionId = [string](Get-DataProperty $comparison 'baselineVersionId' '')
        currentSnapshotId = [string](Get-DataProperty $workbook[0] 'lastRenderedSnapshotId' '')
        currentVersionId = [string](Get-DataProperty $workbook[0] 'lastRenderedVersionId' '')
    }
    foreach ($field in @('baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
        if ([string]::IsNullOrWhiteSpace([string](Get-DataProperty $identity $field ''))) { return $null }
    }
    if ([string](Get-DataProperty $comparison 'currentSnapshotId' '') -ne [string]$identity.currentSnapshotId -or
        [string](Get-DataProperty $comparison 'currentVersionId' '') -ne [string]$identity.currentVersionId) { return $null }
    return $identity
}

function Save-LocalDiffDetailCache([string]$Language, $Detail, [string]$BaselineSnapshotId = '', [string]$CurrentSnapshotId = '') {
    if ($null -eq $Detail) { return }
    $workbookId = [string](Get-DataProperty $Detail 'workbookId' '')
    if ([string]::IsNullOrWhiteSpace($workbookId)) { return }
    $comparison = Get-DataProperty $Detail 'comparison' $null
    if ($null -eq $comparison) { return }
    $baselineSnapshot = [string](Get-DataProperty $comparison 'baselineSnapshotId' $BaselineSnapshotId)
    $currentSnapshot = [string](Get-DataProperty $comparison 'currentSnapshotId' $CurrentSnapshotId)
    $baselineVersion = [string](Get-DataProperty $comparison 'baselineVersionId' '')
    $currentVersion = [string](Get-DataProperty $comparison 'currentVersionId' '')
    $scope = [string](Get-DataProperty $comparison 'scope' $(if (-not [string]::IsNullOrWhiteSpace($BaselineSnapshotId)) { 'history' } else { 'automatic' }))
    $path = Get-LocalDiffDetailCachePath $Language $workbookId $baselineSnapshot $currentSnapshot $baselineVersion $currentVersion $scope
    $cacheKey = ([IO.Path]::GetFullPath($path)).ToLowerInvariant()
    $Script:DiffDetailResponseCache[$cacheKey] = $Detail
    try { Write-JsonFile $path $Detail } catch { }
}

function Get-LocalDiffDetailCache([string]$Language, [string]$WorkbookId, [string]$BaselineSnapshotId = '', [string]$CurrentSnapshotId = '') {
    $identity = Get-LocalDiffDetailCacheIdentity $Language $WorkbookId $BaselineSnapshotId $CurrentSnapshotId
    if ($null -eq $identity) { return $null }
    $path = Get-LocalDiffDetailCachePath $Language $WorkbookId ([string]$identity.baselineSnapshotId) ([string]$identity.currentSnapshotId) ([string]$identity.baselineVersionId) ([string]$identity.currentVersionId) ([string]$identity.scope)
    $cacheKey = ([IO.Path]::GetFullPath($path)).ToLowerInvariant()
    $detail = $Script:DiffDetailResponseCache[$cacheKey]
    if ($null -eq $detail -and (Test-Path -LiteralPath $path)) {
        try { $detail = Read-JsonFile $path $null } catch { $detail = $null }
        if ($null -ne $detail) { $Script:DiffDetailResponseCache[$cacheKey] = $detail }
    }
    if ($null -eq $detail -or (Get-IntDataProperty $detail 'algorithmVersion' 0) -ne $Script:DiffDetailAlgorithmVersion -or
        [string](Get-DataProperty $detail 'workbookId' '') -ne $WorkbookId) { return $null }
    $comparison = Get-DataProperty $detail 'comparison' $null
    if ($null -eq $comparison) { return $null }
    foreach ($field in @('scope','baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId')) {
        if ([string](Get-DataProperty $comparison $field '') -ne [string](Get-DataProperty $identity $field '')) { return $null }
    }
    return $detail
}

function Publish-LatestComparisonCaches([string]$Language, [string]$WorkbookId, $Comparison, $CurrentHashes, $BaselineHashes) {
    if ($null -eq $Comparison -or [string](Get-DataProperty $Comparison 'status' '') -ne 'complete') { return }
    $summary = [ordered]@{
        status = [string](Get-DataProperty $Comparison 'status' '')
        method = [string](Get-DataProperty $Comparison 'method' '')
        changedSheets = @(Get-Array (Get-DataProperty $Comparison 'changedSheets' @()))
        unchangedSheets = @(Get-Array (Get-DataProperty $Comparison 'unchangedSheets' @()))
        unknownSheets = @(Get-Array (Get-DataProperty $Comparison 'unknownSheets' @()))
        addedSheets = @(Get-Array (Get-DataProperty $Comparison 'addedSheets' @()))
        removedSheets = @(Get-Array (Get-DataProperty $Comparison 'removedSheets' @()))
        message = [string](Get-DataProperty $Comparison 'message' '')
        comparedAt = [string](Get-DataProperty $Comparison 'comparedAt' '')
    }
    $workbook = Update-StructureLocked $Language {
        param($structure)
        $matches = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
        if ($matches.Count -eq 0) { return $null }
        Set-NoteProperty $matches[0] 'latestComparisonSummary' $summary
        return $matches[0]
    }
    if ($null -eq $workbook) { return }
    $displayName = [string](Get-DataProperty $workbook 'displayName' '')
    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = [string](Get-DataProperty $workbook 'fileName' $WorkbookId) }
    $context = [pscustomobject][ordered]@{
        available = $true; status = 'available'; message = ''
        workbookId = $WorkbookId; workbookName = $displayName; workbook = $workbook
        comparison = $Comparison; comparisonPersisted = $true
        currentSnapshotId = [string]$Comparison.currentSnapshotId
        currentVersionId = [string]$Comparison.currentVersionId
        baselineSnapshotId = [string]$Comparison.baselineSnapshotId
        baselineVersionId = [string]$Comparison.baselineVersionId
        currentVisualHashes = $CurrentHashes; baselineVisualHashes = $BaselineHashes
        currentAt = Get-DiffSnapshotDate $Language $WorkbookId ([string]$Comparison.currentSnapshotId) ([string]$Comparison.comparedAt)
        baselineAt = Get-DiffSnapshotDate $Language $WorkbookId ([string]$Comparison.baselineSnapshotId) ''
        method = [string](Get-DataProperty $Comparison 'method' '')
        scope = 'automatic'
    }
    $detail = New-DiffDetailSkeleton $Language $context
    Save-LocalDiffDetailCache $Language $detail
    $latestKey = ($Language + '|' + $WorkbookId + '|' + [string]$Comparison.currentSnapshotId + '|' + [string]$Comparison.currentVersionId).ToLowerInvariant()
    $Script:LatestComparisonCache[$latestKey] = $Comparison
    [void](Update-SnapshotSummaryCacheEntry $Language $WorkbookId ([string]$Comparison.currentSnapshotId))
}

function Get-DiffDetail(
    [string]$Language,
    [string]$WorkbookId,
    [string]$BaselineSnapshotId = '',
    [string]$CurrentSnapshotId = ''
) {
    $cached = Get-LocalDiffDetailCache $Language $WorkbookId $BaselineSnapshotId $CurrentSnapshotId
    if ($null -ne $cached) {
        Set-NoteProperty $cached 'performance' ([ordered]@{ contextMs = 0; skeletonMs = 0; totalMs = 0; source = 'local-cache' })
        return $cached
    }
    # 比較PNGと差分JSONは事前生成しない。対象版・シート・ページ数だけを返し、
    # 表示中の1ページをPDF.jsとWeb Workerでブラウザ内比較する。
    $totalTimer = [Diagnostics.Stopwatch]::StartNew()
    $context = Get-DiffDetailContext $Language $WorkbookId $BaselineSnapshotId $CurrentSnapshotId
    $contextMs = $totalTimer.ElapsedMilliseconds
    $detail = New-DiffDetailSkeleton $Language $context
    $totalTimer.Stop()
    Set-NoteProperty $detail 'performance' ([ordered]@{
        contextMs = $contextMs
        skeletonMs = [Math]::Max(0, $totalTimer.ElapsedMilliseconds - $contextMs)
        totalMs = $totalTimer.ElapsedMilliseconds
        source = 'shared-fallback'
    })
    Save-LocalDiffDetailCache $Language $detail $BaselineSnapshotId $CurrentSnapshotId
    if ([string]::IsNullOrWhiteSpace($BaselineSnapshotId) -and [string]::IsNullOrWhiteSpace($CurrentSnapshotId) -and
        [bool]$context.available -and [string]$context.scope -eq 'automatic') {
        try { Publish-LatestComparisonCaches $Language $WorkbookId $context.comparison $context.currentVisualHashes $context.baselineVisualHashes } catch { }
    }
    return $detail
}

function Resolve-DiffContentPdfPath([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId, [string]$SheetName) {
    # Strict identity: never fall back to another render version. Hashes and displayed PDF must be the same generation.
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $safeVersionId = Assert-SafeStorageSegment $VersionId 'versionId'
    if (@(Get-RenderVersionIds $Language $WorkbookId $safeSnapshotId) -notcontains $safeVersionId) { return '' }
    return (Resolve-ContentPdfSheetPathExact $Language $WorkbookId $safeVersionId $SheetName)
}

function Invoke-DiffImagePageGeneration([string]$BeforePdf, [string]$AfterPdf, [string]$OutputDirectory, [string]$Kind) {
    $scriptPath = Join-Path $Script:AppRoot 'tools\diff-image-pages.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath)) { throw '差分画像生成ツールが見つかりません。' }
    $arguments = @{
        OutputDirectory = $OutputDirectory
        Kind = $Kind
        Dpi = $Script:DiffDetailDpi
        Threshold = $Script:DiffDetailThreshold
        MinimumRegionPixels = $Script:DiffDetailMinimumRegionPixels
        Padding = $Script:DiffDetailPadding
    }
    if (-not [string]::IsNullOrWhiteSpace($BeforePdf)) { $arguments.BeforePdf = $BeforePdf }
    if (-not [string]::IsNullOrWhiteSpace($AfterPdf)) { $arguments.AfterPdf = $AfterPdf }
    $raw = @(& $scriptPath @arguments | ForEach-Object { [string]$_ })
    $text = ($raw -join '')
    if ([string]::IsNullOrWhiteSpace($text)) { throw '差分画像生成結果を取得できませんでした。' }
    return ($text | ConvertFrom-Json)
}

function Invoke-DiffImageBatchGeneration($Items) {
    $scriptPath = Join-Path $Script:AppRoot 'tools\diff-image-batch.ps1'
    if (-not (Test-Path -LiteralPath $scriptPath)) { throw '差分画像一括生成ツールが見つかりません。' }
    $requestPath = Join-Path ([IO.Path]::GetTempPath()) ('rb-diff-request-' + [Guid]::NewGuid().ToString('N') + '.json')
    try {
        Write-JsonFile $requestPath ([ordered]@{ schemaVersion = 1; items = @($Items) })
        $raw = @(& $scriptPath -RequestPath $requestPath -Dpi $Script:DiffDetailDpi `
            -Threshold $Script:DiffDetailThreshold -MinimumRegionPixels $Script:DiffDetailMinimumRegionPixels `
            -Padding $Script:DiffDetailPadding | ForEach-Object { [string]$_ })
        $text = ($raw -join '')
        if ([string]::IsNullOrWhiteSpace($text)) { throw '差分画像一括生成結果を取得できませんでした。' }
        $result = $text | ConvertFrom-Json
        if (-not [bool](Get-DataProperty $result 'ok' $false)) { throw '差分画像の一括生成に失敗しました。' }
        return $result
    } finally {
        Remove-Item -LiteralPath $requestPath -Force -ErrorAction SilentlyContinue
    }
}


function Start-DiffDetailJob(
    [string]$Language,
    [string]$WorkbookId,
    [string]$BaselineSnapshotId = '',
    [string]$CurrentSnapshotId = '',
    [string]$SheetKey = ''
) {
    $context = Get-DiffDetailContext $Language $WorkbookId $BaselineSnapshotId $CurrentSnapshotId
    if (-not [bool]$context.available) {
        return [pscustomobject][ordered]@{ ok = $true; jobId = ''; status = 'unavailable'; percent = 100; message = [string]$context.message }
    }
    $safeSheetKey = ''
    if (-not [string]::IsNullOrWhiteSpace($SheetKey)) { $safeSheetKey = Assert-SafeStorageSegment $SheetKey 'sheetKey' }
    $launchLock = Get-DiffLaunchLockPath $Language $context
    return Invoke-WithLock $launchLock {
        $freshContext = Get-DiffDetailContext $Language $WorkbookId $BaselineSnapshotId $CurrentSnapshotId
        if (-not [bool]$freshContext.available) { throw [string]$freshContext.message }
        foreach ($field in @('scope','currentSnapshotId','currentVersionId','baselineSnapshotId','baselineVersionId')) {
            if ([string](Get-DataProperty $context $field '') -ne [string](Get-DataProperty $freshContext $field '')) {
                throw '比較対象が更新されました。差分詳細を開き直してください。'
            }
        }
        $context = $freshContext
        $skeleton = New-DiffDetailSkeleton $Language $context
        if (-not [string]::IsNullOrWhiteSpace($safeSheetKey) -and
            @($skeleton.sheets | Where-Object { [string]$_.sheetKey -eq $safeSheetKey }).Count -eq 0) {
            throw '指定したシートは比較対象に含まれていません。'
        }
        $cacheDir = Get-DiffDetailCacheDirForLanguage $Language $context
        if (-not (Test-Path -LiteralPath $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null }
        $detailPath = Join-Path $cacheDir 'diff-detail.json'
        $existing = $null
        if (Test-Path -LiteralPath $detailPath) { try { $existing = Read-JsonFile $detailPath $null } catch { } }
        if (Test-DiffDetailMatchesContext $existing $context) {
            if ([string]::IsNullOrWhiteSpace($safeSheetKey) -and [string](Get-DataProperty $existing 'status' '') -eq 'ready') {
                return [pscustomobject][ordered]@{ ok = $true; jobId = ''; status = 'completed'; percent = 100; message = '差分詳細は作成済みです。'; detailReady = $true }
            }
            if (-not [string]::IsNullOrWhiteSpace($safeSheetKey)) {
                $target = @(Get-Array $existing.sheets | Where-Object { [string]$_.sheetKey -eq $safeSheetKey } | Select-Object -First 1)
                if ($target.Count -gt 0 -and [string](Get-DataProperty $target[0] 'status' '') -eq 'ready') {
                    return [pscustomobject][ordered]@{ ok = $true; jobId = ''; status = 'completed'; percent = 100; message = 'このシートの差分画像は作成済みです。'; detailReady = $true }
                }
            }
        }
        $pointerPath = Join-Path $cacheDir 'diff-job.json'
        if (Test-Path -LiteralPath $pointerPath) {
            try {
                $pointer = Read-JsonFile $pointerPath $null
                $activeId = [string](Get-DataProperty $pointer 'jobId' '')
                if (-not [string]::IsNullOrWhiteSpace($activeId)) {
                    $active = Read-RenderJobStatus $Language $activeId
                    if (@('completed','completed-with-errors','failed','missing','cancelled') -notcontains [string](Get-DataProperty $active 'status' '')) {
                        Set-NoteProperty $active 'joinedExistingDiffJob' $true
                        return $active
                    }
                }
            } catch { }
        }
        $jobId = 'job_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '_' + ([Guid]::NewGuid().ToString('N').Substring(0,8))
        $leases = @(New-DiffJobLeases $Language $context $jobId)
        $jobDir = Get-RenderJobDir $Language
        $inputPath = Join-Path $jobDir "$jobId.diff.input.json"
        $statusPath = Join-Path $jobDir "$jobId.status.json"
        $stdoutPath = Join-Path $jobDir "$jobId.diff.out.log"
        $stderrPath = Join-Path $jobDir "$jobId.diff.err.log"
        # 比較画像はシート単位で遅延生成するため、1ジョブの対象は最大1シート。
        $sheetCount = $(if ($skeleton.sheets.Count -gt 0) { 1 } else { 0 })
        $initial = [pscustomobject][ordered]@{
            ok = $true; jobType = 'diff-detail'; jobId = $jobId; status = 'queued'; total = $sheetCount
            completed = 0; failed = 0; percent = 1; message = '差分詳細を準備しています。'
            currentWorkbookId = [string]$context.workbookId; currentWorkbookName = [string]$context.workbookName
            currentSheet = ''; processId = 0; stdoutPath = $stdoutPath; stderrPath = $stderrPath
            results = @(); errors = @(); startedAt = New-NowIso; updatedAt = New-NowIso; stateSavedAt = ''
        }
        try {
            if ([string]$context.scope -eq 'history' -and -not [bool]$context.comparisonPersisted) {
                $saved = Save-HistoricalSnapshotComparison $Language ([string]$context.workbookId) $context.comparison
                $context.comparison = $saved
                $context.comparisonPersisted = $true
            }
            Write-RenderJobStatus $statusPath $initial
            Write-JsonFile $inputPath ([ordered]@{
                jobId = $jobId; mode = $Language; workbookId = [string]$context.workbookId
                currentSnapshotId = [string]$context.currentSnapshotId; currentVersionId = [string]$context.currentVersionId
                baselineSnapshotId = [string]$context.baselineSnapshotId; baselineVersionId = [string]$context.baselineVersionId
                scope = [string]$context.scope; sheetKey = $safeSheetKey
                cacheDir = $cacheDir; detailPath = $detailPath; statusPath = $statusPath
                stdoutPath = $stdoutPath; stderrPath = $stderrPath; leases = @($leases)
                generationLockPath = (Get-DiffGenerationLockPath $Language $context)
            })
            Write-JsonFile $pointerPath ([ordered]@{ jobId = $jobId; sheetKey = $safeSheetKey; createdAt = New-NowIso })
            $psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
            if (-not (Test-Path -LiteralPath $psExe)) { $psExe = 'powershell.exe' }
            $serverScript = Join-Path $Script:AppRoot 'server.ps1'
            $command = "& '$($serverScript.Replace("'", "''"))' -Mode '$($Language.Replace("'", "''"))' -DiffJobPath '$($inputPath.Replace("'", "''"))'"
            $proc = Start-HiddenPowerShellChild $psExe $command $stdoutPath $stderrPath
            if (-not $proc -or -not $proc.Id) { throw '差分画像の作成プロセスIDを取得できませんでした。' }
            $initial.processId = [int]$proc.Id; $initial.status = 'launching'; $initial.percent = 2
            $initial.message = '差分画像の作成プロセスを起動しました。'
            Write-RenderJobStatus $statusPath $initial
            return $initial
        } catch {
            $initial.status = 'failed'; $initial.percent = 100; $initial.message = '差分画像の作成プロセスを起動できませんでした。'
            $initial.errors = @([ordered]@{ error = $_.Exception.Message; detail = Get-ErrorDetail $_ })
            try { Write-RenderJobStatus $statusPath $initial } catch { }
            Remove-DiffJobLeases ([pscustomobject][ordered]@{ mode=$Language; workbookId=[string]$context.workbookId; leases=@($leases) })
            throw
        }
    }
}


function Invoke-DiffDetailJobCore($Job) {
    $language = [string](Get-DataProperty $Job 'mode' $Mode)
    $workbookId = [string](Get-DataProperty $Job 'workbookId' '')
    $statusPath = [string](Get-DataProperty $Job 'statusPath' '')
    $detailPath = [string](Get-DataProperty $Job 'detailPath' '')
    $cacheDir = [IO.Path]::GetFullPath([string](Get-DataProperty $Job 'cacheDir' ''))
    $jobId = [string](Get-DataProperty $Job 'jobId' '')
    $requestedSheetKey = [string](Get-DataProperty $Job 'sheetKey' '')
    $status = [pscustomobject][ordered]@{
        ok = $true; jobType = 'diff-detail'; jobId = $jobId; status = 'running'; total = 0; completed = 0; failed = 0
        percent = 3; message = '比較対象を確認しています。'; currentWorkbookId = $workbookId; currentWorkbookName = ''
        currentSheet = ''; processId = [System.Diagnostics.Process]::GetCurrentProcess().Id
        results = @(); errors = @(); startedAt = New-NowIso; updatedAt = New-NowIso; stateSavedAt = ''
    }
    Write-RenderJobStatus $statusPath $status
    $detail = $null
    try {
        Refresh-DiffJobLeases $Job
        $jobScope = [string](Get-DataProperty $Job 'scope' 'automatic')
        $context = if ($jobScope -eq 'history') {
            Get-DiffDetailContext $language $workbookId ([string]$Job.baselineSnapshotId) ([string]$Job.currentSnapshotId)
        } else { Get-DiffDetailContext $language $workbookId }
        if (-not [bool]$context.available) { throw [string]$context.message }
        foreach ($field in @('currentSnapshotId','currentVersionId','baselineSnapshotId','baselineVersionId','scope')) {
            if ([string](Get-DataProperty $Job $field '') -ne [string](Get-DataProperty $context $field '')) {
                throw '比較対象が更新されました。変更バッジを開き直してください。'
            }
        }
        $expectedCache = [IO.Path]::GetFullPath((Get-DiffDetailCacheDirForLanguage $language $context))
        if ($expectedCache -ne $cacheDir) { throw '差分キャッシュの保存先が不正です。' }
        if (-not (Test-Path -LiteralPath $cacheDir)) { New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null }
        # 全体再試行でも既存詳細を読み込み、failedになった変更なしシートを保持する。
        # 読み込んだ内容は直後のTest-DiffDetailMatchesContextで比較対象との完全一致を検証する。
        if (Test-Path -LiteralPath $detailPath) {
            $detail = Read-JsonFile $detailPath $null
        }
        if (-not (Test-DiffDetailMatchesContext $detail $context)) {
            $detail = New-DiffDetailSkeleton $language $context
        }
        $allSheets = @(Get-Array $detail.sheets)
        $workIndexes = @()
        if ([string]::IsNullOrWhiteSpace($requestedSheetKey)) {
            # 旧Web UIなどsheetKeyを送らない呼び出しでも、全シート一括生成には戻さない。
            # 変更シートを優先し、1ジョブにつき1シートだけ処理する。
            $preferredIndex = -1
            for ($n = 0; $n -lt $allSheets.Count; $n++) {
                $sheetKind = [string](Get-DataProperty $allSheets[$n] 'kind' '')
                $sheetStatus = [string](Get-DataProperty $allSheets[$n] 'status' '')
                if ($sheetKind -ne 'unchanged' -and @('pending','deferred','failed') -contains $sheetStatus) {
                    $preferredIndex = $n
                    break
                }
            }
            if ($preferredIndex -lt 0) {
                for ($n = 0; $n -lt $allSheets.Count; $n++) {
                    $sheetStatus = [string](Get-DataProperty $allSheets[$n] 'status' '')
                    if (@('pending','deferred','failed') -contains $sheetStatus) { $preferredIndex = $n; break }
                }
            }
            if ($preferredIndex -ge 0) { $workIndexes += $preferredIndex }
        } else {
            for ($n = 0; $n -lt $allSheets.Count; $n++) { if ([string]$allSheets[$n].sheetKey -eq $requestedSheetKey) { $workIndexes += $n; break } }
            if ($workIndexes.Count -eq 0) { throw '指定した項目は比較対象に含まれていません。' }
        }
        $detail.sheets = @($allSheets)
        $detail.status = 'generating'
        $detail.generation = [ordered]@{ status = 'running'; jobId = $jobId; percent = 3; message = '差分画像を作成しています。'; currentSheet = '' }
        Write-JsonFile $detailPath $detail
        $status.total = $workIndexes.Count; $status.currentWorkbookName = [string]$context.workbookName
        Write-RenderJobStatus $statusPath $status

        # 変更シートごとに Java/PDFBox を2回起動していた旧経路を避ける。
        # 全シートの新旧PDFを1つのJVMへ渡し、最大4並列でラスタライズしてから一括解析する。
        $batchRequest = @()
        $batchIdByIndex = @{}
        $batchPreparationErrors = @{}
        for ($position = 0; $position -lt $workIndexes.Count; $position++) {
            $i = [int]$workIndexes[$position]
            $sheet = $detail.sheets[$i]
            $name = [string]$sheet.sheetName
            $beforeName = [string](Get-DataProperty $sheet 'beforeSheetName' $name)
            $afterName = [string](Get-DataProperty $sheet 'afterSheetName' $name)
            $kind = [string]$sheet.kind
            try {
                $beforePdf = ''; $afterPdf = ''
                if ($kind -ne 'added') { $beforePdf = Resolve-DiffContentPdfPath $language $workbookId ([string]$context.baselineSnapshotId) ([string]$context.baselineVersionId) $beforeName }
                if ($kind -ne 'removed') { $afterPdf = Resolve-DiffContentPdfPath $language $workbookId ([string]$context.currentSnapshotId) ([string]$context.currentVersionId) $afterName }
                $missing = (($kind -in @('modified','unchanged') -and ([string]::IsNullOrWhiteSpace($beforePdf) -or [string]::IsNullOrWhiteSpace($afterPdf))) -or
                    ($kind -eq 'added' -and [string]::IsNullOrWhiteSpace($afterPdf)) -or ($kind -eq 'removed' -and [string]::IsNullOrWhiteSpace($beforePdf)))
                if ($missing) { throw '同一レンダリング世代のcontent PDFが見つかりません。' }
                if ($kind -eq 'unknown' -and [string]::IsNullOrWhiteSpace($beforePdf) -and [string]::IsNullOrWhiteSpace($afterPdf)) { throw '比較元・比較先のPDFを確認できません。' }
                $sheetDir = Join-Path $cacheDir (Join-Path 'p' ([string]$sheet.sheetKey))
                if (Test-Path -LiteralPath $sheetDir) { Remove-Item -LiteralPath $sheetDir -Recurse -Force -ErrorAction SilentlyContinue }
                New-Item -ItemType Directory -Path $sheetDir -Force | Out-Null
                $batchId = 'i' + $position.ToString('0000')
                $batchIdByIndex[[string]$i] = $batchId
                $batchRequest += [ordered]@{
                    id = $batchId
                    beforePdf = $beforePdf
                    afterPdf = $afterPdf
                    beforeRasterDirectory = $(if ($kind -ne 'added') {
                        Get-RenderRasterSheetDir $language $workbookId ([string]$context.baselineSnapshotId) ([string]$context.baselineVersionId) $beforeName
                    } else { '' })
                    afterRasterDirectory = $(if ($kind -ne 'removed') {
                        Get-RenderRasterSheetDir $language $workbookId ([string]$context.currentSnapshotId) ([string]$context.currentVersionId) $afterName
                    } else { '' })
                    beforePageCount = Get-IntDataProperty $sheet 'beforePages' 0
                    afterPageCount = Get-IntDataProperty $sheet 'afterPages' 0
                    unchangedPageNumbers = @(Get-Array (Get-DataProperty $sheet 'unchangedPageNumbers' @()))
                    outputDirectory = $sheetDir
                    kind = $kind
                }
            } catch {
                $batchPreparationErrors[[string]$i] = $_.Exception.Message
            }
        }
        $batchResultMap = @{}
        $batchTimings = $null
        if ($batchRequest.Count -gt 0) {
            Refresh-DiffJobLeases $Job
            $status.message = "新旧PDFをまとめて画像化・解析しています（$($batchRequest.Count)シート）。"
            $status.currentSheet = ''
            $status.percent = 5
            $detail.generation = [ordered]@{ status = 'running'; jobId = $jobId; percent = 5; message = $status.message; currentSheet = '' }
            Write-JsonFile $detailPath $detail; Write-RenderJobStatus $statusPath $status
            $batchGenerated = Invoke-DiffImageBatchGeneration $batchRequest
            $batchTimings = Get-DataProperty $batchGenerated 'timings' $null
            foreach ($batchItem in @(Get-Array (Get-DataProperty $batchGenerated 'items' @()))) {
                $batchResultMap[[string](Get-DataProperty $batchItem 'id' '')] = $batchItem
            }
        }

        for ($position = 0; $position -lt $workIndexes.Count; $position++) {
            Refresh-DiffJobLeases $Job
            $i = [int]$workIndexes[$position]
            $sheet = $detail.sheets[$i]
            $name = [string]$sheet.sheetName; $kind = [string]$sheet.kind
            $sheet.status = 'generating'; $sheet.message = ''
            $status.currentSheet = $name
            $status.message = "差分画像を作成しています: $($position + 1) / $($workIndexes.Count) 項目 $name"
            $status.percent = [int][Math]::Max(5, [Math]::Min(95, [Math]::Floor((([double]$position) / [Math]::Max(1, $workIndexes.Count)) * 90) + 5))
            Write-RenderJobStatus $statusPath $status
            try {
                if ($batchPreparationErrors.ContainsKey([string]$i)) { throw [string]$batchPreparationErrors[[string]$i] }
                $batchId = [string]$batchIdByIndex[[string]$i]
                if ([string]::IsNullOrWhiteSpace($batchId) -or -not $batchResultMap.ContainsKey($batchId)) {
                    throw '差分画像一括生成結果に対象シートがありません。'
                }
                $generated = $batchResultMap[$batchId]
                if (-not [bool](Get-DataProperty $generated 'ok' $false)) {
                    throw [string](Get-DataProperty $generated 'message' '差分画像を生成できませんでした。')
                }
                $pages = @(); $regionTotal = 0; $hasUnknownPage = $false
                foreach ($page in @(Get-Array $generated.pages)) {
                    $regions = @(Get-Array (Get-DataProperty $page 'regions' @())); $regionTotal += $regions.Count
                    if ([string](Get-DataProperty $page 'status' '') -eq 'unknown') { $hasUnknownPage = $true }
                    $pages += [pscustomobject][ordered]@{
                        pageNumber = Get-IntDataProperty $page 'pageNumber' 0; width = Get-IntDataProperty $page 'width' 0; height = Get-IntDataProperty $page 'height' 0
                        beforePageNumber = Get-IntDataProperty $page 'beforePageNumber' 0; afterPageNumber = Get-IntDataProperty $page 'afterPageNumber' 0
                        comparisonKind = [string](Get-DataProperty $page 'comparisonKind' ''); matchMethod = [string](Get-DataProperty $page 'matchMethod' '')
                        mappingAmbiguous = [bool](Get-DataProperty $page 'mappingAmbiguous' $false); mappingMessage = [string](Get-DataProperty $page 'mappingMessage' '')
                        pageSizeChanged = [bool](Get-DataProperty $page 'pageSizeChanged' $false); status = [string](Get-DataProperty $page 'status' 'ready')
                        message = [string](Get-DataProperty $page 'message' ''); confidence = [double](Get-DataProperty $page 'confidence' 1)
                        changedRatio = [double](Get-DataProperty $page 'changedRatio' 0); beforeAsset = [string](Get-DataProperty $page 'beforeFile' '')
                        afterAsset = [string](Get-DataProperty $page 'afterFile' ''); beforeMaskAsset = [string](Get-DataProperty $page 'beforeMaskFile' '')
                        beforeOverlayAsset = [string](Get-DataProperty $page 'beforeOverlayFile' ''); maskAsset = [string](Get-DataProperty $page 'afterMaskFile' '')
                        overlayAsset = [string](Get-DataProperty $page 'afterOverlayFile' ''); regions = @($regions)
                    }
                }
                $sheet.pages = @($pages); $sheet.beforePages = Get-IntDataProperty $generated 'beforePageCount' 0
                $sheet.afterPages = Get-IntDataProperty $generated 'afterPageCount' 0
                $sheet.pageCount = [Math]::Max([int]$sheet.beforePages, [int]$sheet.afterPages); $sheet.regionCount = $regionTotal
                $sheet.status = $(if ($kind -eq 'unknown' -or $hasUnknownPage) { 'unknown' } else { 'ready' })
                if ($kind -eq 'unknown') { $sheet.message = '信頼できる差分領域を判定できないため、強調表示は行いません。' }
            } catch {
                # 画像生成そのものの失敗は、比較上の「判定不能」と区別する。
                # failed のまま残すことで、全体再試行またはシート再選択から再生成できる。
                $sheet.status = 'failed'; $sheet.message = $_.Exception.Message; $sheet.pages = @(); $sheet.regionCount = 0
                $status.failed++; $status.errors = @($status.errors) + @([ordered]@{ sheetName = $name; error = $_.Exception.Message; detail = Get-ErrorDetail $_ })
            }
            $status.completed = $position + 1
            $status.percent = [int][Math]::Max(8, [Math]::Min(98, [Math]::Floor((([double]($position + 1)) / [Math]::Max(1, $workIndexes.Count)) * 93) + 5))
            $detail.sheets[$i] = $sheet
            $detail.generation = [ordered]@{ status = 'running'; jobId = $jobId; percent = $status.percent; message = $status.message; currentSheet = $name }
            Write-JsonFile $detailPath $detail; Write-RenderJobStatus $statusPath $status
        }
        # 1シートだけの遅延生成でも 'ready' を書いていたため、全シート生成が途中で失敗した後に
        # 「変更なし」シートを1件開くと status が ready に上書きされ、pending のまま残った
        # シートが二度と生成できなくなっていた(Start-DiffDetailJob が「作成済み」を返す)。
        # 実際の残件から status を決める。
        $pendingSheets = @(Get-Array $detail.sheets | Where-Object { @('pending','generating') -contains [string](Get-DataProperty $_ 'status' '') }).Count
        $failedSheets = @(Get-Array $detail.sheets | Where-Object { [string](Get-DataProperty $_ 'status' '') -eq 'failed' }).Count
        if ($pendingSheets -eq 0 -and $failedSheets -eq 0) {
            $detail.status = 'ready'; $detail.message = ''
        } else {
            $detail.status = 'failed'
            $parts = @()
            if ($pendingSheets -gt 0) { $parts += "未作成 $pendingSheets 件" }
            if ($failedSheets -gt 0) { $parts += "作成失敗 $failedSheets 件" }
            $detail.message = ('差分画像に未完了の項目があります（' + ($parts -join '、') + '）。再試行してください。')
        }
        $detail.generatedAt = New-NowIso
        Set-NoteProperty $detail 'performance' $batchTimings
        $detail.generation = [ordered]@{ status = 'completed'; jobId = $jobId; percent = 100; message = '差分詳細を作成しました。'; currentSheet = '' }
        Write-JsonFile $detailPath $detail
        $status.status = $(if ($status.failed -gt 0) { 'completed-with-errors' } else { 'completed' }); $status.percent = 100; $status.currentSheet = ''
        $status.message = $(if ($status.failed -gt 0) { '一部のシートを除き、差分詳細を作成しました。' } else { '差分詳細を作成しました。' })
        $status.results = @([ordered]@{ workbookId = $workbookId; detailPath = $detailPath; sheetCount = $workIndexes.Count; performance = $batchTimings })
        Write-RenderJobStatus $statusPath $status
    } catch {
        $status.status = 'failed'; $status.percent = 100; $status.message = '差分詳細を作成できませんでした。'
        $status.errors = @([ordered]@{ error = $_.Exception.Message; detail = Get-ErrorDetail $_ }); Write-RenderJobStatus $statusPath $status
        if ($null -ne $detail) {
            try {
                $detail.status = 'failed'
                if (-not [string]::IsNullOrWhiteSpace($requestedSheetKey)) {
                    $target = @(Get-Array $detail.sheets | Where-Object { [string](Get-DataProperty $_ 'sheetKey' '') -eq $requestedSheetKey } | Select-Object -First 1)
                    if ($target.Count -gt 0) {
                        $target[0].status = 'failed'
                        $target[0].message = $_.Exception.Message
                        $target[0].pages = @()
                        $target[0].regionCount = 0
                    }
                }
                $detail.message = $_.Exception.Message
                $detail.generation = [ordered]@{ status = 'failed'; jobId = $jobId; percent = 100; message = $_.Exception.Message; currentSheet = '' }
                Write-JsonFile $detailPath $detail
            } catch { }
        }
    }
}


function Invoke-DiffDetailJobFromFile([string]$JobPath) {
    $job = Read-JsonFile $JobPath $null
    if ($null -eq $job) { throw "Diff job file is not readable: $JobPath" }
    try {
        $language = [string](Get-DataProperty $job 'mode' $Mode)
        $workbookId = [string](Get-DataProperty $job 'workbookId' '')
        $jobScope = [string](Get-DataProperty $job 'scope' 'automatic')
        $context = if ($jobScope -eq 'history') {
            Get-DiffDetailContext $language $workbookId ([string]$job.baselineSnapshotId) ([string]$job.currentSnapshotId)
        } else { Get-DiffDetailContext $language $workbookId }
        if (-not [bool]$context.available) { throw [string]$context.message }
        $expectedLock = [IO.Path]::GetFullPath((Get-DiffGenerationLockPath $language $context))
        $providedLockText = [string](Get-DataProperty $job 'generationLockPath' '')
        if ([string]::IsNullOrWhiteSpace($providedLockText)) { throw '差分生成ロックの保存先がありません。' }
        $providedLock = [IO.Path]::GetFullPath($providedLockText)
        if ($expectedLock -ne $providedLock) { throw '差分生成ロックの保存先が不正です。' }
        Invoke-WithLock $expectedLock { Invoke-DiffDetailJobCore $job }
    } finally {
        Remove-DiffJobLeases $job
    }
}

function Serve-DiffPage($Context, [string]$Language, [string]$WorkbookId, [string]$CurrentSnapshotId, [string]$BaselineSnapshotId, [string]$SheetKey, [string]$PageNumberText, [string]$Asset, [string]$Scope = 'automatic') {
    $allowedAssets = @('before','after','before-mask','before-overlay','mask','overlay')
    if ($allowedAssets -notcontains $Asset) { throw [ArgumentException]::new('asset が不正です。') }
    $pageNumber = 0
    if (-not [int]::TryParse($PageNumberText, [ref]$pageNumber) -or $pageNumber -le 0) {
        throw [ArgumentException]::new('pageNumber が不正です。')
    }
    if ([string]::IsNullOrWhiteSpace($Scope)) { $Scope = 'automatic' }
    if ($Scope -notin @('automatic','history')) { throw [ArgumentException]::new('scope が不正です。') }
    $diffContext = if ($Scope -eq 'history') {
        Get-DiffDetailContext $Language $WorkbookId $BaselineSnapshotId $CurrentSnapshotId
    } else {
        Get-DiffDetailContext $Language $WorkbookId
    }
    if (-not [bool]$diffContext.available) { throw [string]$diffContext.message }
    if ((Assert-SafeStorageSegment $CurrentSnapshotId 'currentSnapshotId') -ne [string]$diffContext.currentSnapshotId -or
        (Assert-SafeStorageSegment $BaselineSnapshotId 'baselineSnapshotId') -ne [string]$diffContext.baselineSnapshotId) {
        throw '比較対象が更新されました。差分詳細を開き直してください。'
    }
    $safeSheetKey = Assert-SafeStorageSegment $SheetKey 'sheetKey'
    $cacheDir = Get-DiffDetailCacheDirForLanguage $Language $diffContext
    $detailPath = Join-Path $cacheDir 'diff-detail.json'
    if (-not (Test-Path -LiteralPath $detailPath)) { throw '差分詳細はまだ作成されていません。' }
    $detail = Read-JsonFile $detailPath $null
    if (-not (Test-DiffDetailMatchesContext $detail $diffContext)) {
        throw '差分詳細の比較対象が一致しません。差分詳細を再作成してください。'
    }
    $sheet = @(Get-Array $detail.sheets | Where-Object { [string]$_.sheetKey -eq $safeSheetKey } | Select-Object -First 1)
    if ($sheet.Count -eq 0) { throw '指定した項目は比較対象に含まれていません。' }
    $page = @(Get-Array $sheet[0].pages | Where-Object { (Get-IntDataProperty $_ 'pageNumber' 0) -eq $pageNumber } | Select-Object -First 1)
    if ($page.Count -eq 0) { throw '指定したページは比較対象に含まれていません。' }
    $property = switch ($Asset) {
        'before' { 'beforeAsset' }
        'after' { 'afterAsset' }
        'before-mask' { 'beforeMaskAsset' }
        'before-overlay' { 'beforeOverlayAsset' }
        'mask' { 'maskAsset' }
        'overlay' { 'overlayAsset' }
    }
    $fileName = [string](Get-DataProperty $page[0] $property '')
    if ([string]::IsNullOrWhiteSpace($fileName)) { throw '差分画像がありません。' }

    $full = ''
    $root = ''
    if ($Asset -in @('before','after') -and $fileName -eq 'render.png') {
        # レンダリング時に保存した120 DPI PNGを直接返し、比較キャッシュへの重複コピーを省く。
        $sheetName = $(if ($Asset -eq 'before') { [string](Get-DataProperty $sheet[0] 'beforeSheetName' (Get-DataProperty $sheet[0] 'sheetName' '')) } else { [string](Get-DataProperty $sheet[0] 'afterSheetName' (Get-DataProperty $sheet[0] 'sheetName' '')) })
        if ($Asset -eq 'before') {
            $rasterDir = Get-RenderRasterSheetDir $Language ([string]$diffContext.workbookId) ([string]$diffContext.baselineSnapshotId) ([string]$diffContext.baselineVersionId) $sheetName
        } else {
            $rasterDir = Get-RenderRasterSheetDir $Language ([string]$diffContext.workbookId) ([string]$diffContext.currentSnapshotId) ([string]$diffContext.currentVersionId) $sheetName
        }
        $root = [IO.Path]::GetFullPath($rasterDir)
        $full = [IO.Path]::GetFullPath((Join-Path $root ('page-{0:0000}.png' -f $pageNumber)))
    } else {
        if ($fileName -notmatch '^[A-Za-z0-9._-]+\.png$' -or $fileName -eq 'render.png') { throw '差分画像がありません。' }
        $root = [IO.Path]::GetFullPath($cacheDir)
        $full = [IO.Path]::GetFullPath((Join-Path $cacheDir (Join-Path 'p' (Join-Path $safeSheetKey $fileName))))
    }
    if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
    if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $full)) {
        throw '差分画像が見つかりません。'
    }
    # URLには比較元・比較先・アルゴリズム版が含まれるため、元ラスタを直接返す場合もimmutableでよい。
    Write-BytesResponse $Context 200 ([IO.File]::ReadAllBytes($full)) 'image/png' $false 'private, max-age=31536000, immutable'
}

function Get-WorkbookChangeSummary([string]$Language, [string]$WorkbookId, $Workbook = $null) {
    $cmp = Get-LatestComparison $Language $WorkbookId $Workbook
    if ($null -eq $cmp) { return $null }
    return [ordered]@{
        status = [string](Get-DataProperty $cmp 'status' '')
        method = [string](Get-DataProperty $cmp 'method' '')
        changedSheets = @(Get-Array (Get-DataProperty $cmp 'changedSheets' @()))
        unchangedSheets = @(Get-Array (Get-DataProperty $cmp 'unchangedSheets' @()))
        unknownSheets = @(Get-Array (Get-DataProperty $cmp 'unknownSheets' @()))
        # V5-P1(#10): 追加・削除シートも state に載せる。比較結果は持っているのにここで捨てていた。
        addedSheets = @(Get-Array (Get-DataProperty $cmp 'addedSheets' @()))
        removedSheets = @(Get-Array (Get-DataProperty $cmp 'removedSheets' @()))
        message = [string](Get-DataProperty $cmp 'message' '')
        comparedAt = [string](Get-DataProperty $cmp 'comparedAt' '')
    }
}

function Invoke-StartupRecovery([string]$Language) {
    # lease・一時コピー・自動状態は直後に起動するスケジューラー子プロセスが
    # 復旧するため、UI サーバー側では二重に共有ドライブを走査しない。
    try { Recover-FinalTransactions $Language } catch { }
}

# V5-P0: スケジューラーモード。これが無いと子プロセスが通常サーバーとして起動し、
# さらに孫スケジューラーを起動して無限に増殖する。通常サーバー初期化より前に置くこと。
if (-not [string]::IsNullOrWhiteSpace($AutoSchedulerPath)) {
    Invoke-AutoSchedulerFromFile -ControlPath $AutoSchedulerPath -ParentProcessId $ParentProcessId
    return
}

if (-not [string]::IsNullOrWhiteSpace($DiffJobPath)) {
    Invoke-DiffDetailJobFromFile $DiffJobPath
    return
}

if (-not [string]::IsNullOrWhiteSpace($RenderJobPath)) {
    Invoke-RenderJobFromFile $RenderJobPath
    return
}

if (-not [string]::IsNullOrWhiteSpace($FinalJobPath)) {
    Invoke-FinalBuildJobFromFile $FinalJobPath
    return
}

if ($Port -le 0) { $Port = Get-FreePort }
$config0 = Get-AppConfig
$config0 = Initialize-LocalProjectConfig $config0
try { $startupPaths=Get-Paths; if($startupPaths.dataDir -and (Test-Path -LiteralPath ([string]$startupPaths.dataDir))){Ensure-Package $startupPaths -Languages @((Get-EffectiveLanguage))} } catch { Write-Warning $_.Exception.Message }

# 未完了の提出用PDFトランザクションだけは、UI操作を受け付ける前に復旧する。
try { Invoke-StartupRecovery (Get-EffectiveLanguage) } catch { Write-Warning $_.Exception.Message }

$prefix = "http://127.0.0.1:$Port/"
$url = ("http://127.0.0.1:{0}/?token={1}" -f $Port, $Script:Token)

# Use a small TcpListener-based HTTP server instead of HttpListener.
# This avoids URL ACL / administrator-rights issues on locked-down Windows PCs.
Start-LocalTcpServer $Port $url ([bool]$NoOpen)
