param(
    [ValidateSet('ja','en')]
    [string]$Mode = 'ja',
    [int]$Port = 0,
    [string]$Token = '',
    [switch]$NoOpen,
    [string]$RenderJobPath = '',
    [string]$DiffJobPath = '',
    [string]$AutoSchedulerPath = '',
    [int]$ParentProcessId = 0
)

$ErrorActionPreference = 'Stop'

$Script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:WebRoot = Join-Path $Script:AppRoot 'web'
$Script:DefaultConfigPath = Join-Path $Script:AppRoot 'default-config.json'
$localConfigOverride = ([string]$env:REPORTBINDER_LOCAL_CONFIG_ROOT).Trim()
if ([string]::IsNullOrWhiteSpace($localConfigOverride)) {
    $Script:LocalConfigRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ReportBinder'
} else {
    $Script:LocalConfigRoot = [IO.Path]::GetFullPath($localConfigOverride)
}
$Script:ConfigPath = Join-Path $Script:LocalConfigRoot 'config.json'
if (-not (Test-Path -LiteralPath $Script:LocalConfigRoot)) { New-Item -ItemType Directory -Path $Script:LocalConfigRoot -Force | Out-Null }
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
# 最小化コンソールで起動されたとき、何のウィンドウか分かるようにタイトルを付ける。
try { $host.UI.RawUI.WindowTitle = "ReportBinder サーバー ($Mode) - このウィンドウを閉じると終了します" } catch { }
$Script:AutoRenderInProgress = $false
$Script:ServerStartedUtc = [DateTime]::UtcNow
$Script:IdleTimeoutSeconds = 1800
$Script:NoClientStartupTimeoutSeconds = 600
$Script:ReadyGifBytes = [Convert]::FromBase64String('R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==')
$Script:ExcelPrintProfileVersion = 2026072201
$Script:FinalPdfComposerProfileVersion = 20260604
$Script:RenderEnvironmentCache = $null
$Script:RenderEnvironmentCompared = $false
$Script:CurrentRenderEnvFingerprint = ''
$Script:CurrentRenderEnvInfo = $null
$Script:LastRenderAttempt = $null
$Script:PdfPageAnalyzerVersion = 1
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



function Invoke-NativeCapture([string]$FilePath, [string[]]$ArgumentList) {
    # V5-P3: Windows PowerShell 5.1 では、ネイティブコマンドの stderr を 2>&1 で取り込むと
    # ErrorRecord としてパイプラインに流れ、$ErrorActionPreference='Stop' の下では
    # NativeCommandError の例外になる。
    # PDFBox は日本語フォントを含むPDFで警告(Format 14 cmap table ...)を stderr に出すため、
    # 解析や組版が成功していても呼び出し側が「失敗」と誤認していた。
    # ここだけ Continue に落として出力を文字列として回収する。
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $lines = @()
    $exit = -1
    try {
        $lines = @(& $FilePath @ArgumentList 2>&1 | ForEach-Object { [string]$_ })
        $exit = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previous
    }
    return [ordered]@{ exitCode = $exit; output = $lines; text = ($lines -join "`n") }
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

function Read-TextFileShared([string]$Path) {
    $share = [IO.FileShare]([int][IO.FileShare]::ReadWrite -bor [int][IO.FileShare]::Delete)
    $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, $share)
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
            if (-not (Test-Path -LiteralPath $Path)) { return $DefaultValue }
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
    [IO.File]::WriteAllText($Path, $Text, $encoding)
}

function Write-Utf8NoBomFileShared([string]$Path, [string]$Text) {
    # Fallback for protected / synced folders where File.Replace can intermittently
    # raise Access Denied. Readers use ReadWrite/Delete sharing and Read-JsonFile
    # retries partial reads, so progress keeps moving instead of staying at 0%.
    $encoding = New-Object System.Text.UTF8Encoding -ArgumentList $false
    $bytes = $encoding.GetBytes($Text)
    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    for ($attempt = 0; $attempt -lt 12; $attempt++) {
        $fs = $null
        try {
            $share = [IO.FileShare]([int][IO.FileShare]::ReadWrite -bor [int][IO.FileShare]::Delete)
            $fs = [IO.File]::Open($Path, [IO.FileMode]::Create, [IO.FileAccess]::Write, $share)
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
    for ($attempt = 0; $attempt -lt 12; $attempt++) {
        try {
            if (Test-Path -LiteralPath $DestinationPath) {
                [IO.File]::Replace($SourcePath, $DestinationPath, $null, $true)
            } else {
                [IO.File]::Move($SourcePath, $DestinationPath)
            }
            return
        } catch [System.IO.FileNotFoundException] {
            try { [IO.File]::Move($SourcePath, $DestinationPath); return } catch { if ($attempt -ge 11) { throw } }
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
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
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
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
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
    Set-ArrayProperty $Structure 'workbooks'
    Set-ArrayProperty $Structure 'pages'
    if ($null -eq $Structure.volumes) {
        $Structure | Add-Member -NotePropertyName 'volumes' -NotePropertyValue ([ordered]@{}) -Force
    }
    return $Structure
}

function Test-DirectExcelRelativePath([string]$RelativePath) {
    Test-RelativePath $RelativePath | Out-Null
    if ($RelativePath -match '[\\/]') { throw '提出フォルダ直下のExcelだけ登録できます。子フォルダ内のファイルは対象外です。' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -ne '.xlsx') { throw '拡張子が .xlsx のExcelだけ登録できます。' }
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
    $default = [ordered]@{ schemaVersion = 1; lastSubmissionDir = ''; lastDataDir = ''; lastOutputDir = ''; lastMode = $Mode }
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
    # 提出Excelの現物は、管理データが提出フォルダ配下にある場合だけ保持する。
    # policy.json ではなく実パスを検証し、外部フォルダへの意図しない複製を防ぐ。
    try {
        $paths = Get-Paths
        $sub = [IO.Path]::GetFullPath([string]$paths.submissionDir)
        if (-not $sub.EndsWith([IO.Path]::DirectorySeparatorChar)) { $sub += [IO.Path]::DirectorySeparatorChar }
        $data = [IO.Path]::GetFullPath([string]$paths.dataDir)
        if (-not $data.StartsWith($sub, [StringComparison]::OrdinalIgnoreCase)) { return $false }
    } catch { return $false }
    return $true
}

function Get-AutoRenderSettings {
    $config = Get-AppConfig
    $auto = Get-DataProperty $config 'autoRender' $null
    return [ordered]@{
        enabled = [bool](Get-DataProperty $auto 'enabled' $false)
        quietPeriodSeconds = [int](Get-DataProperty $auto 'quietPeriodSeconds' 180)
        requireStableHashCount = [int](Get-DataProperty $auto 'requireStableHashCount' 2)
        deferWhileExcelInUse = [bool](Get-DataProperty $auto 'deferWhileExcelInUse' $true)
    }
}

function Get-InputHistorySettings {
    $config = Get-AppConfig
    $ih = Get-DataProperty $config 'inputHistory' $null
    return [ordered]@{
        retainSourceVersions = [int](Get-DataProperty $ih 'retainSourceVersions' 2)
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


function Get-DefaultChildPaths([string]$SubmissionDir) {
    if ([string]::IsNullOrWhiteSpace($SubmissionDir)) {
        return [ordered]@{ submissionDir = ''; dataDir = ''; outputDir = '' }
    }
    $trimmed = $SubmissionDir.TrimEnd([char[]]@([char]92, [char]47))
    return [ordered]@{
        submissionDir = $trimmed
        dataDir = (Join-Path $trimmed '_reportbinder')
        outputDir = (Join-Path $trimmed '出力')
    }
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
        $proc = Start-Process -FilePath $psExe -ArgumentList $args -WindowStyle Hidden -PassThru -Wait
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
    return Join-Path $resolvedDataDir $Language
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

function New-EmptyStructure([string]$Language) {
    $volumes = [ordered]@{}
    foreach ($volume in @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })) {
        foreach ($category in @('ecm','bod','dmm')) {
            $volumes[(Get-VolumeStateKey $volume $category)] = New-EmptyVolumeState
        }
    }
    return [ordered]@{
        schemaVersion = 2
        language = $Language
        workbooks = @()
        pages = @()
        volumes = $volumes
        updatedAt = New-NowIso
    }
}

function Ensure-Package($Paths, [string[]]$Languages = @('ja','en')) {
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
    foreach ($lang in @($Languages | Where-Object { $_ -in @('ja','en') } | Select-Object -Unique)) {
        foreach ($dir in @('', 'workbooks', 'pages', 'content-pdf', 'exports', 'state', 'locks', 'logs')) {
            $fullDir = Join-Path (Join-Path $dataDir $lang) $dir
            if (-not (Test-Path -LiteralPath $fullDir)) { New-Item -ItemType Directory -Path $fullDir -Force | Out-Null }
        }
        # Use the selected dataDir directly. On first run the local config is intentionally
        # saved only after package initialization succeeds, so Get-Paths is still empty here.
        Initialize-Or-MigrateStructure $lang ([string]$Paths.dataDir) | Out-Null
    }
}


function Resolve-PageId($Page) {
    if ($null -eq $Page) { return '' }
    $existing = [string]$Page.pageId
    if (-not [string]::IsNullOrWhiteSpace($existing)) { return $existing }
    $legacy = [string]$Page.id
    if (-not [string]::IsNullOrWhiteSpace($legacy)) { return $legacy }
    $wb = [string]$Page.workbookId
    $sheetName = [string]$Page.sheetName
    if ([string]::IsNullOrWhiteSpace($wb) -or [string]::IsNullOrWhiteSpace($sheetName)) { return '' }
    $sheetKey = [regex]::Replace($sheetName, '[^0-9A-Za-z]+', '-')
    if ([string]::IsNullOrWhiteSpace($sheetKey)) { $sheetKey = 'sheet' }
    return "$wb-$sheetKey"
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

function Test-WorkbookRenderIsCurrent($Workbook) {
    if ($null -eq $Workbook) { return $false }
    $status = [string](Get-DataProperty $Workbook 'status' '')
    if ($status -in @('new','excel-updated','missing','render-error','rendering','stale','not-rendered')) { return $false }
    $lastRenderedHash = [string](Get-DataProperty $Workbook 'lastRenderedExcelHash' '')
    if ([string]::IsNullOrWhiteSpace($lastRenderedHash)) { return $false }
    $currentHash = [string](Get-DataProperty $Workbook 'currentExcelHash' '')
    if ((-not [string]::IsNullOrWhiteSpace($currentHash)) -and $currentHash -ne $lastRenderedHash) { return $false }
    $profileVersion = Get-IntDataProperty $Workbook 'renderProfileVersion' 0
    if ($profileVersion -lt $Script:ExcelPrintProfileVersion) { return $false }
    return $true
}

function Mark-WorkbookContentStaleForProfile($Structure, $Workbook) {
    $changed = $false
    if ($null -eq $Workbook) { return $changed }
    $lastRenderedHash = [string](Get-DataProperty $Workbook 'lastRenderedExcelHash' '')
    if ([string]::IsNullOrWhiteSpace($lastRenderedHash)) { return $changed }
    $profileVersion = Get-IntDataProperty $Workbook 'renderProfileVersion' 0
    if ($profileVersion -ge $Script:ExcelPrintProfileVersion) { return $changed }
    if ([string](Get-DataProperty $Workbook 'status' '') -ne 'excel-updated') {
        Set-NoteProperty $Workbook 'status' 'excel-updated'
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
            @{Name='warnings'; Value=@()}
        )) {
            if ($null -eq $wb.PSObject.Properties[$pair.Name]) { Set-NoteProperty $wb $pair.Name $pair.Value; $changed = $true }
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
    }
    return $changed
}

function Test-StructureDocument($Structure, [string]$Language) {
    if ($null -eq $Structure) { return $false }
    try {
        $version = [int](Get-DataProperty $Structure 'schemaVersion' 0)
        if ($version -lt 1 -or $version -gt 2) { return $false }
        $storedLanguage = [string](Get-DataProperty $Structure 'language' '')
        if (-not [string]::IsNullOrWhiteSpace($storedLanguage) -and $storedLanguage -ne $Language) { return $false }
        if (-not (Test-ConfigHasKey $Structure 'workbooks') -or
            -not (Test-ConfigHasKey $Structure 'pages')) { return $false }
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

function Update-StructureLocked([string]$Language, [scriptblock]$Mutation) {
    $workspace = Get-WorkspacePath $Language
    $structureLock = Join-Path $workspace 'locks\structure.lock'
    return Invoke-WithLock $structureLock {
        $structure = Read-StructureUnlocked $Language
        $result = & $Mutation $structure
        Write-StructureUnlocked $Language $structure
        # 同一ファイルシステム時刻内の連続更新でも古い読取結果を返さない。
        $Script:StructureReadCache.Clear()
        return $result
    }
}

function Save-Structure([string]$Language, $Structure) {
    throw 'Save-Structureの直接呼出しは禁止されています。Update-StructureLockedを使用してください。'
}

function Get-EffectiveLanguage {
    if ($Mode -eq 'en') { return 'en' }
    return 'ja'
}

function Get-VolumeList([string]$Language) {
    if ($Language -eq 'ja') { return @('ja-main','ja-appendix','none') }
    return @('en-main','en-appendix','none')
}

function Get-DefaultVolume([string]$Language) {
    if ($Language -eq 'ja') { return 'ja-main' }
    return 'en-main'
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
    $excelExts = @('.xlsx')
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
                if ($visible -and $sheetName -match '^[0-9]+$') {
                    $a1 = ''
                    try { $a1 = [string]$ws.Range('A1').Text } catch { $a1 = '' }
                    if ([string]::IsNullOrWhiteSpace($a1)) { $a1 = '' }
                    $zoom = $null
                    try { $zoom = $ws.PageSetup.Zoom } catch { }
                    $printArea = ''
                    try { $printArea = [string]$ws.PageSetup.PrintArea } catch { }
                    $sheets += [ordered]@{
                        sheetName = $sheetName
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
    $cat = Require-WorkbookCategory $Category
    $wbIds = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks | Where-Object { Test-WorkbookCategory $_ $cat })) { $wbIds[[string]$wb.workbookId] = $true }
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
    $cat = Require-WorkbookCategory $Category
    $existing = @(Renumber-VolumeOrder $Structure $Volume $cat)
    $manual = @($existing | Where-Object { [bool](Get-DataProperty $_ 'orderManual' $false) }).Count -gt 0
    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($p in $existing) { [void]$ordered.Add($p) }
    $insertedAtEnd = $manual
    if ($manual -or $existing.Count -eq 0) {
        [void]$ordered.Add($NewPage)
    } else {
        $newWb = @(Get-Array $Structure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$NewPage.workbookId } | Select-Object -First 1)
        $newFile = if ($newWb.Count) { [string]$newWb[0].fileName } else { '' }
        $newSheet = Get-SheetOrderNumber ([string]$NewPage.sheetName)
        $newFileOrder = Get-FileOrderNumber $newFile
        $at = $existing.Count
        for ($i=0; $i -lt $existing.Count; $i++) {
            $p = $existing[$i]
            $wb = @(Get-Array $Structure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$p.workbookId } | Select-Object -First 1)
            $file = if ($wb.Count) { [string]$wb[0].fileName } else { '' }
            $sheet = Get-SheetOrderNumber ([string]$p.sheetName)
            $fileOrder = Get-FileOrderNumber $file
            if ($sheet -gt $newSheet -or ($sheet -eq $newSheet -and $fileOrder -gt $newFileOrder) -or ($sheet -eq $newSheet -and $fileOrder -eq $newFileOrder -and [StringComparer]::OrdinalIgnoreCase.Compare($file,$newFile) -gt 0)) { $at=$i; break }
        }
        $ordered.Insert($at,$NewPage)
    }
    for ($i=0; $i -lt $ordered.Count; $i++) { Set-NoteProperty $ordered[$i] 'order' (($i+1)*10) }
    # Windows PowerShell 5.1 can throw "Argument types do not match" when
    # @() directly converts List[object] through PSToObjectArrayBinder.
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
        if ($version -gt 2) { throw "この管理データは新しいschemaVersion=$versionです。対応するReportBinderを使用してください。" }
        if ($version -eq 2) {
            $lastGood = Join-Path $workspace 'structure.json.last-good'
            if (-not (Test-Path -LiteralPath $lastGood)) { Write-JsonFile $lastGood $structure }
            return [ordered]@{ created=$false; migrated=$false }
        }
        $backup = Join-Path $workspace 'structure.json.v1.bak'
        if (-not (Test-Path -LiteralPath $backup)) {
            Copy-Item -LiteralPath $path -Destination $backup -ErrorAction Stop
            if ((Get-Item $backup).Length -ne (Get-Item $path).Length -or (New-StableHash $backup) -ne (New-StableHash $path)) { throw 'structure.json.v1.bakの検証に失敗しました。' }
        }
        $structure = Normalize-StructureCollections $structure
        [void](Repair-StructurePages $structure)
        $oldVolumes = Get-DataProperty $structure 'volumes' $null
        $newVolumes = [ordered]@{}
        foreach ($volume in @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })) {
            foreach ($cat in @('ecm','bod','dmm')) {
                $fresh = New-EmptyVolumeState
                if ($cat -eq 'ecm' -and $null -ne $oldVolumes) {
                    $old = Get-DataProperty $oldVolumes $volume $null
                    if ($null -eq $old) { $old = Get-DataProperty $oldVolumes (Get-VolumeStateKey $volume $cat) $null }
                    if ($null -ne $old) {
                        foreach ($name in @('status','lastBuiltAt','outputPdf','message','builtFingerprint','staleReasons')) {
                            $value=Get-DataProperty $old $name $null
                            if ($null -ne $value) { Set-NoteProperty $fresh $name $value }
                        }
                    }
                }
                $newVolumes[(Get-VolumeStateKey $volume $cat)] = $fresh
            }
        }
        Set-NoteProperty $structure 'volumes' $newVolumes
        Set-NoteProperty $structure 'schemaVersion' 2
        foreach ($cat in @('ecm','bod','dmm')) { foreach ($volume in @(Get-VolumeList $Language)) { [void](Renumber-VolumeOrder $structure $volume $cat) } }
        Write-StructureUnlocked $Language $structure $DataDir
        Write-JsonFile (Join-Path $workspace 'schema-version.json') ([ordered]@{ schemaVersion=2; migratedAt=(New-NowIso); backup=$backup })
        return [ordered]@{ created=$false; migrated=$true; backup=$backup }
    }
}

function Apply-DefaultNumberingPerVolume([string]$Language, $Structure, [string]$Category) {
    $cat = Require-WorkbookCategory $Category
    $wbIds = @{}
    foreach ($wb in @(Get-Array $Structure.workbooks | Where-Object { Test-WorkbookCategory $_ $cat })) { $wbIds[[string]$wb.workbookId] = $true }
    foreach ($vol in @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })) {
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
    if ($text -match 'Open プロパティを取得できません|Unable to get the Open property') { return 'Excelがこのファイルを開けませんでした。ファイルが保護ビュー対象（インターネット由来）・暗号化（秘密度ラベル/IRM）・破損のいずれかの可能性があります。ファイルを右クリック→プロパティ→「許可する」にチェック後、もう一度PDF作成してください。' }
    if ($text -match 'このオブジェクトにプロパティ|stateSavedAt|プロパティ.*見つかりません|property.*not found|does not contain a property') { return 'PDF作成の進捗状態を更新できませんでした。ReportBinderを更新してから、もう一度PDF作成してください。' }
    if ($text -match 'PDF化対象のシート|半角数字') { return 'PDF化対象のシートがありません。シート名を半角数字だけにしてください。例: 1, 2, 3' }
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
    return ([string]::Join('|', @($names | Sort-Object { Get-SheetOrderNumber $_ }, { [string]$_ })))
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
    $cat=Require-WorkbookCategory $Category
    foreach ($volume in @($Volumes | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($volume) -or $volume -eq 'none') { continue }
        $key=Get-VolumeStateKey $volume $cat
        $v=Get-DataProperty $Structure.volumes $key $null
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

function Update-WorkbookPagesFromInspection([string]$Language, $Structure, $Workbook, $Sheets) {
    $workbookId=[string](Get-DataProperty $Workbook 'workbookId' '')
    $cat=Require-WorkbookCategory ([string](Get-DataProperty $Workbook 'category' ''))
    $pages=@(Get-Array $Structure.pages)
    $sortedSheets=@(Get-Array $Sheets | Sort-Object @{Expression={Get-SheetOrderNumber ([string](Get-DataProperty $_ 'sheetName' ''))};Ascending=$true}, @{Expression={[string](Get-DataProperty $_ 'sheetName' '')};Ascending=$true})
    $current=@{}; foreach($s in $sortedSheets){$n=[string](Get-DataProperty $s 'sheetName' '');if($n){$current[$n]=$true}}
    $removed=@(); $kept=@()
    foreach($p in $pages){ if([string]$p.workbookId -eq $workbookId -and (-not $current.ContainsKey([string]$p.sheetName))){$removed+= $p}else{$kept+=$p} }
    $pages=@($kept); $Structure.pages=$pages
    $known=@{}; foreach($p in $pages){$known[(Resolve-PageId $p)]=$p}
    $added=@();$updated=@();$atEnd=@();$inOrder=0
    foreach($s in $sortedSheets){
        $sheet=[string](Get-DataProperty $s 'sheetName' ''); if(-not $sheet){continue}
        $pageId="$workbookId-$([regex]::Replace($sheet,'[^0-9A-Za-z]+','-'))"
        $title=[string](Get-DataProperty $s 'detectedTitle' '');if(-not $title){$title="$([string]$Workbook.fileName) / $sheet"}
        if($known.ContainsKey($pageId)){
            $p=$known[$pageId];Set-NoteProperty $p 'sheetName' $sheet;Set-NoteProperty $p 'detectedTitle' $title;if(-not [string]$p.title){Set-NoteProperty $p 'title' $title};$updated+=$pageId
        }else{
            # V5-P3: pages は JSON 由来の PSCustomObject で構成される。新規ページだけ OrderedDictionary にすると
            #        Windows PowerShell 5.1 が Sort-Object 等の型比較で「引数の型が一致しません」を投げることがある。
            #        コレクションの要素型を必ず揃える。
            $p=[pscustomobject][ordered]@{pageId=$pageId;workbookId=$workbookId;sheetName=$sheet;titleSource='A1';detectedTitle=$title;title=$title;volume=(Get-DefaultVolume $Language);order=0;orderManual=$false;numberingMode='visible';numberingManual=$false;numberingDefault='first-page-none';enabled=$true;contentPdf=$null;status='not-rendered';warnings=@();updatedAt=New-NowIso}
            $ins=Insert-PageInSheetOrder $Structure $p ([string]$p.volume) $cat
            $Structure.pages=@(Get-Array $Structure.pages)+@($p)
            if($ins.insertedAtEnd){$atEnd+=$pageId}else{$inOrder++}
            $added+=$pageId;$known[$pageId]=$p
        }
    }
    foreach($vol in @(Get-VolumeList $Language)){[void](Renumber-VolumeOrder $Structure $vol $cat)}
    if($removed.Count -gt 0 -or $added.Count -gt 0){
        $affected=@($removed|ForEach-Object{[string]$_.volume})
        if($added.Count -gt 0){$affected+=@((Get-DefaultVolume $Language))}
        Mark-VolumeNeedsRebuild $Structure $Language $cat @($affected|Where-Object{$_ -and $_ -ne 'none'}|Select-Object -Unique) 'render' 'Excelのシート構成が変更されました'
    }
    Apply-DefaultNumberingPerVolume $Language $Structure $cat
    $names=@($sortedSheets|ForEach-Object{[string]$_.sheetName})
    return [ordered]@{addedPageIds=@($added);updatedPageIds=@($updated);removedPages=@($removed);sheetNames=$names;sheetFingerprint=(Get-SheetFingerprintFromNames $names);addedCount=$added.Count;insertedInOrderCount=$inOrder;insertedAtEndCount=$atEnd.Count;insertedAtEndPageIds=@($atEnd)}
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

function Remove-XlsxHeaderFooterXml([string]$XlsxPath) {
    # Remove Excel header/footer definitions from the temporary XLSX package before Excel opens it.
    # This avoids relying on slow/fragile COM FirstPage/EvenPage PageSetup calls and prevents
    # footer page numbers such as &P from being exported into the content PDF.
    if ([string]::IsNullOrWhiteSpace($XlsxPath)) { return [ordered]@{ ok = $true; changed = 0; skipped = $true } }
    $ext = [IO.Path]::GetExtension($XlsxPath).ToLowerInvariant()
    if (@('.xlsx','.xlsm','.xltx','.xltm') -notcontains $ext) { return [ordered]@{ ok = $true; changed = 0; skipped = $true } }
    if (-not (Test-Path -LiteralPath $XlsxPath)) { return [ordered]@{ ok = $true; changed = 0; skipped = $true } }

    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue | Out-Null
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
    } catch { }

    $zip = $null
    $changed = 0
    try {
        $zip = [System.IO.Compression.ZipFile]::Open($XlsxPath, [System.IO.Compression.ZipArchiveMode]::Update)
        $targets = @($zip.Entries | Where-Object { [string]$_.FullName -match '^xl/(worksheets|chartsheets)/[^/]+\.xml$' })
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

            $hasHeaderFooter = ([regex]::IsMatch($xml, '(?is)<(?:\w+:)?headerFooter\b'))
            if (-not $hasHeaderFooter) { continue }

            $newXml = [regex]::Replace($xml, '(?is)<(?:\w+:)?headerFooter\b[^>]*(?:/>|>.*?</(?:\w+:)?headerFooter>)', '')
            # Only touch pageMargins when the sheet actually had header/footer content. Rewriting every
            # sheet XML just to change blank header/footer margins is slow and does not affect page numbers.
            $newXml = [regex]::Replace($newXml, '(?is)<(?:\w+:)?pageMargins\b[^>]*>', [System.Text.RegularExpressions.MatchEvaluator]{
                param($m)
                $tag = [string]$m.Value
                if ($tag -notmatch '\sheader=') { $tag = $tag -replace '/?>$', ' header="0"$0' }
                else { $tag = [regex]::Replace($tag, 'header="[^"]*"', 'header="0"') }
                if ($tag -notmatch '\sfooter=') { $tag = $tag -replace '/?>$', ' footer="0"$0' }
                else { $tag = [regex]::Replace($tag, 'footer="[^"]*"', 'footer="0"') }
                return $tag
            })

            if ($newXml -ne $xml) {
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
        return [ordered]@{ ok = $true; changed = $changed; skipped = $false }
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
        # 暫定PDFは左右1.2cmを基準にする。最終PDFでは奇数/偶数ページを0.2cmだけ内側へ寄せる(パンチ側1.4cm/外側1.0cm)。
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
    $run = Invoke-NativeCapture ([string]$tool.javaExe) @('-cp', [string]$tool.classPath, 'BatchPdfSplitter', '--source', $BatchPdf, '--map', $mapPath)
    $output = @($run.output)
    $global:LASTEXITCODE = [int]$run.exitCode
    $exit = $LASTEXITCODE
    $outputText = ($output -join "`n")
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
            $active.ExportAsFixedFormat(0, $batchPdf, 0, $true, $false, $missing, $missing, $false, $missing)
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
        $Worksheet.ExportAsFixedFormat(0, $OutPdf, 0, $true, $false, $missing, $missing, $false, $missing)
        return (Wait-ForPdfOutput $OutPdf $SheetName)
    } catch {
        $directError = $_.Exception.Message
        throw "ExcelでPDF化できませんでした: シート $SheetName / 直接出力=[$directError]"
    } finally {
        try { $Workbook.Activate() | Out-Null } catch { }
    }
}

function Get-PdfPageCount([string]$PdfPath) {
    try {
        $bytes = [IO.File]::ReadAllBytes($PdfPath)
        $text = [Text.Encoding]::ASCII.GetString($bytes)
        $count = ([regex]::Matches($text, '/Type\s*/Page(?!s)\b')).Count
        if ($count -lt 1) { return 1 }
        return $count
    } catch { return 1 }
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
    return $excel
}

function Close-ExcelApplicationForRender($Excel) {
    if ($Excel) {
        try { $Excel.Quit() } catch { }
        Invoke-ComRelease $Excel
    }
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

function Request-ServerShutdown([string]$Reason) {
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

function Register-Workbook([string]$Language, [string]$RelativePath, [string]$Category = '') {
    Test-DirectExcelRelativePath $RelativePath | Out-Null
    $cat=Require-WorkbookCategory $Category
    $paths=Get-Paths
    $full=Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "提出ファイルが見つかりません: $RelativePath" }
    $candidate=New-WorkbookObject $RelativePath $Language $cat
    return Update-StructureLocked $Language {
        param($structure)
        $normalized=([string]$RelativePath -replace '\\','/').ToLowerInvariant()
        $existing=@(Get-Array $structure.workbooks | Where-Object {
            ([string]$_.workbookId -eq [string]$candidate.workbookId) -or ((([string]$_.relativePath -replace '\\','/').ToLowerInvariant() -eq $normalized) -and (Test-WorkbookCategory $_ $cat))
        } | Select-Object -First 1)
        if ($existing.Count -gt 0) {
            $old=$existing[0]
            Set-NoteProperty $candidate 'workbookId' ([string]$old.workbookId)
            foreach ($name in @('lastRenderedVersionId','lastRenderedExcelHash','lastRenderedAt','lastRenderedSheets','lastRenderedSheetFingerprint','lastRenderLog','renderProfileVersion','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash')) {
                Set-NoteProperty $candidate $name (Get-DataProperty $old $name $null)
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$candidate.lastRenderedExcelHash)) { Set-NoteProperty $candidate 'status' $(if ($candidate.currentExcelHash -ne $candidate.lastRenderedExcelHash) {'excel-updated'} else {[string]$old.status}) }
        }
        $structure.workbooks=@(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -ne [string]$candidate.workbookId }) + @($candidate)
        Mark-VolumeNeedsRebuild $structure $Language $cat @((Get-DefaultVolume $Language)) 'register' 'Excelを1件登録しました'
        return [ordered]@{ workbook=$candidate; sheets=@(); registered=$true; inspected=$false }
    }
}

function Register-WorkbooksBatch([string]$Language, $RelativePaths, [string]$Category = '') {
    $registered = @()
    $errors = @()
    $seen = @{}
    foreach ($raw in (Get-Array $RelativePaths)) {
        $rel = [string]$raw
        if ([string]::IsNullOrWhiteSpace($rel)) { continue }
        $key = ($rel -replace '\\', '/').ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        try {
            $r = Register-Workbook $Language $rel $Category
            $wb = $r.workbook
            $registered += [ordered]@{
                relativePath = $rel
                workbookId = [string]$wb.workbookId
                fileName = [string]$wb.fileName
                displayName = [string]$wb.displayName
                status = [string]$wb.status
                category = [string]$wb.category
            }
        } catch {
            $errors += [ordered]@{ relativePath = $rel; error = $_.Exception.Message; detail = [string]$_ }
        }
    }
    return [ordered]@{
        requestedCount = @(Get-Array $RelativePaths).Count
        registered = @($registered)
        errors = @($errors)
        registeredCount = @($registered).Count
        errorCount = @($errors).Count
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
                    foreach ($p in @(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId -and [string]$_.status -ne 'confirmed' })) { Set-NoteProperty $p 'status' 'render-error'; Set-NoteProperty $p 'warnings' @($userMessage); Set-NoteProperty $p 'updatedAt' (New-NowIso) }
                } else {
                    Set-NoteProperty $wb[0] 'status' 'excel-updated'
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
        $cat=Require-WorkbookCategory ([string]$found[0].category)
        $removedPages=@(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId })
        $affected=@($removedPages | ForEach-Object {[string]$_.volume} | Where-Object {$_ -and $_ -ne 'none'} | Select-Object -Unique)
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
                         [string]$SourceOverridePath = '', [string]$SourceSnapshotId = '', [string]$ExpectedSourceHash = '') {
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
        $ephemeralCaptureId = ''
        if ([string]::IsNullOrWhiteSpace($SourceOverridePath)) {
            try { $inputInfo = Capture-RenderInput $Language $WorkbookId $SourceSnapshotId '' } catch { $inputInfo = $null }
            if ($null -ne $inputInfo -and -not [string]::IsNullOrWhiteSpace([string]$inputInfo.path)) {
                $SourceOverridePath = [string]$inputInfo.path
                $SourceSnapshotId = [string]$inputInfo.snapshotId
                $ExpectedSourceHash = [string]$inputInfo.hash
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
        for ($readAttempt = 1; $readAttempt -le 3; $readAttempt++) {
            try {
                $sourceHash = New-StableHash $sourcePath
                if (-not [string]::IsNullOrWhiteSpace($sourceHash)) { break }
            } catch { $lastReadError = $_.Exception.Message }
            if ($readAttempt -lt 3) { Start-Sleep -Seconds 2 }
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

        $excelPackageScrub = $null
        try { $excelPackageScrub = Remove-XlsxHeaderFooterXml $tmpPath } catch { $excelPackageScrub = [ordered]@{ ok = $false; error = $_.Exception.Message } }

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
                $excel = New-ExcelApplicationForRender
                $ownsExcel = $true
            } else {
                $steps += '既存のExcel COMを使用'
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

            if ($excelPackageScrub -and $excelPackageScrub.ok -eq $true -and [int](Get-DataProperty $excelPackageScrub 'changed' 0) -gt 0) {
                $steps += "Excel内部のヘッダー/フッターXMLを削除: $([int](Get-DataProperty $excelPackageScrub 'changed' 0)) 件"
            } elseif ($excelPackageScrub -and $excelPackageScrub.ok -eq $false) {
                $warnings += "Excel内部ヘッダー/フッターXMLの事前削除に失敗しました。COM設定で削除を続行します: $([string](Get-DataProperty $excelPackageScrub 'error' ''))"
            }
            $steps += '一時コピーを開く'
            if ($ProgressCallback) { & $ProgressCallback 'open' $WorkbookId '' }
            $book = Open-ExcelWorkbookSafe $excel $tmpPath $true
            try { $book.CheckCompatibility = $false } catch { }

            $inspected = @()
            $targetSheetNames = @()
            $sheetRenderInfos = @()
            $sheetCount = 0
            try { $sheetCount = [int]$book.Worksheets.Count } catch { $sheetCount = 0 }
            $deferredPrintCommunication = Set-ExcelPrintCommunicationSafe $excel $false
            try {
                for ($i = 1; $i -le $sheetCount; $i++) {
                    $ws = $null
                    try {
                        $ws = $book.Worksheets.Item($i)
                        $sheetName = [string]$ws.Name
                        $visible = ([int]$ws.Visible -eq -1)
                        if ($visible -and $sheetName -match '^[0-9]+$') {
                            $targetSheetNames += $sheetName
                            $a1 = ''
                            try { $a1 = [string]$ws.Range('A1').Text } catch { }
                            if ([string]::IsNullOrWhiteSpace($a1)) { $a1 = "$($wb.fileName) / $sheetName" }
                            # Reading PageSetup.PrintArea is another slow COM call and is only diagnostic.
                            # Keep it blank in render logs to avoid delaying PDF作成.
                            $printArea = ''
                            $inspected += [ordered]@{ sheetName = $sheetName; titleSource = 'A1'; detectedTitle = $a1; printArea = $printArea }

                            $steps += "シート $sheetName の印刷設定を調整"
                            if ($ProgressCallback) { & $ProgressCallback 'sheet-setup' $WorkbookId $sheetName }
                            Apply-StandardPrintSettings $ws $excel $deferredPrintCommunication
                            $outPdf = Join-Path $contentDir "$sheetName.pdf"
                            $sheetRenderInfos += [ordered]@{ sheetName = $sheetName; outPdf = $outPdf; titleSource = 'A1'; detectedTitle = $a1; printArea = $printArea }
                        }
                    } finally {
                        Invoke-ComRelease $ws
                    }
                }
            } finally {
                if ($deferredPrintCommunication) { [void](Set-ExcelPrintCommunicationSafe $excel $true) }
            }
            if ($targetSheetNames.Count -eq 0) {
                $emptyPageSync = Update-WorkbookPagesFromInspection $Language $structure $wb @()
                # V5-§3.7: 旧世代の個別削除は廃止。世代単位の掃除(Remove-WorkbookContentPdfs)に一本化する。
                # 個別に消すと、保持しているはずの世代フォルダの中身が欠損する。
                Set-WorkbookRenderedSheetSnapshot $wb @()
                Update-StructureLocked $Language { param($st) $x=@(Get-Array $st.workbooks|Where-Object{[string]$_.workbookId -eq $WorkbookId}|Select-Object -First 1);if($x.Count){[void](Update-WorkbookPagesFromInspection $Language $st $x[0] @());Set-WorkbookRenderedSheetSnapshot $x[0] @()} } | Out-Null
                throw 'PDF化対象のシートがありません。シート名が半角数字のみ（例: 1, 2, 003）のシートを用意してください。'
            }

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

            $pageSync = Update-WorkbookPagesFromInspection $Language $structure $wb $inspected
            $removedPages = @(Get-Array (Get-DataProperty $pageSync 'removedPages' @()))
            if ($removedPages.Count -gt 0) {
                $removedNames = @($removedPages | ForEach-Object { [string](Get-DataProperty $_ 'sheetName' '') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $removedLabel = [string]::Join('、', $removedNames)
                if ([string]::IsNullOrWhiteSpace($removedLabel)) { $removedLabel = "$($removedPages.Count) ページ" }
                $warnings += "現在のExcelに存在しないシートをページ構成から外しました: $removedLabel"
                $steps += "存在しないシートの古いページを削除: $removedLabel"
                # V5-§3.7: 旧世代の個別削除は廃止(上記と同じ理由)。
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
                $lw=$latest[0];$sync=Update-WorkbookPagesFromInspection $Language $st $lw $inspected
                foreach($r in $rendered){$pageId="$WorkbookId-$([regex]::Replace([string]$r.sheetName,'[^0-9A-Za-z]+','-'))";$pg=@(Get-Array $st.pages|Where-Object{(Resolve-PageId $_)-eq $pageId}|Select-Object -First 1);if($pg.Count){Set-NoteProperty $pg[0] 'contentPdf' (Get-RelativePathCompat $workspace ([string]$r.pdf));Set-NoteProperty $pg[0] 'status' 'rendered';Set-NoteProperty $pg[0] 'warnings' @($r.warnings);Set-NoteProperty $pg[0] 'updatedAt' (New-NowIso)}}
                # V5-§2.4: currentExcel* はコピーしない(Scan-Updates の専有)。
                foreach($name in @('lastRenderedVersionId','lastRenderedExcelHash','lastRenderedAt','renderProfileVersion','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash','warnings','lastRenderLog')){Set-NoteProperty $lw $name (Get-DataProperty $wb $name $null)}
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
                $cat=Require-WorkbookCategory ([string]$lw.category);$vols=@(Get-Array $st.pages|Where-Object{[string]$_.workbookId -eq $WorkbookId}|ForEach-Object{[string]$_.volume}|Where-Object{$_ -and $_ -ne 'none'}|Select-Object -Unique);Mark-VolumeNeedsRebuild $st $Language $cat $vols 'render' 'Excelを1件PDF作成しました'
                return $sync
            }
            Write-JsonFile (Join-Path $workspace $logRel) ([ordered]@{ workbookId = $WorkbookId; rendered = $rendered; warnings = $warnings; steps = $steps; sheetSync = $pageSync; at = New-NowIso })
            # V5-P1: 解析はここでは実行しない。
            # レンダリング用Excelを開いたまま、レンダリングロックを保持したまま解析すると、
            # 比較用の再レンダリングが2つ目のExcel COMを起動してしまう(1ジョブ制限に反する)。
            # ロック解放後に実行するため、対象だけを記録しておく。
            $Script:PendingAnalysis = [ordered]@{ language = $Language; workbookId = $WorkbookId; snapshotId = [string]$SourceSnapshotId; versionId = $versionId; rendered = $rendered }
            Remove-WorkbookContentPdfs $workspace $WorkbookId $versionId
        } finally {
            if ($book) { try { $book.Close($false) } catch { } ; Invoke-ComRelease $book }
            if ($ownsExcel -and $excel) { Close-ExcelApplicationForRender $excel }
            if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (-not [string]::IsNullOrWhiteSpace($ephemeralCaptureId)) { Remove-EphemeralCopy $Language $ephemeralCaptureId }
            if ($ownsExcel -or -not $KeepExcelOpen) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
        }
        return [ordered]@{ workbookId = $WorkbookId; versionId = $versionId; rendered = $rendered; warnings = $warnings; steps = $steps; sheetSync = $pageSync }
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
        try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered) }
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
                    $cat = Require-WorkbookCategory ([string]$w.category)
                    $vols = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $id } | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
                    Mark-VolumeNeedsRebuild $st $Language $cat $vols 'excel-updated' '元Excelが見つかりません'
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
            Set-NoteProperty $w 'currentExcelModifiedAt' $modified
            Set-NoteProperty $w 'currentExcelLastWriteUtcTicks' $ticks
            Set-NoteProperty $w 'currentExcelSize' $size
            if (-not [string]::IsNullOrWhiteSpace($hash)) { Set-NoteProperty $w 'currentExcelHash' $hash }
            $profileOld = (-not [string]::IsNullOrWhiteSpace($lastRendered)) -and ((Get-IntDataProperty $w 'renderProfileVersion' 0) -lt $Script:ExcelPrintProfileVersion)
            $stale = (-not [string]::IsNullOrWhiteSpace($lastRendered)) -and (([string]::IsNullOrWhiteSpace($hash)) -or $hash -ne $lastRendered -or $profileOld)
            if ($stale) {
                Set-NoteProperty $w 'status' 'excel-updated'
                foreach ($page in @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $id -and -not [string]::IsNullOrWhiteSpace([string]$_.contentPdf) })) { Set-NoteProperty $page 'status' 'stale' }
                $newlyDetected = ($previousStatus -ne 'excel-updated') -or ($previousHash -ne $hash)
                if ($newlyDetected) {
                    $cat = Require-WorkbookCategory ([string]$w.category)
                    $vols = @(Get-Array $st.pages | Where-Object { [string]$_.workbookId -eq $id } | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
                    Mark-VolumeNeedsRebuild $st $Language $cat $vols 'excel-updated' '元Excelが1件更新されました'
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
        $renderProfileOutdated = ((-not [string]::IsNullOrWhiteSpace($lastRenderedHash)) -and ((Get-IntDataProperty $wb 'renderProfileVersion' 0) -lt $Script:ExcelPrintProfileVersion))
        if ($status -eq 'render-error') {
            # The PDF button is an explicit retry. A previous render-error must not make
            # the button look idle just because the same file hash already failed once.
            $needs = $true
        } else {
            if ([string]::IsNullOrWhiteSpace($lastRenderedHash)) { $needs = $true }
            if ($status -in @('new','excel-updated')) { $needs = $true }
            if ($renderProfileOutdated) { $needs = $true }
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
                $r = Render-Workbook $Language $id
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

function Start-RenderJob([string]$Language, [string[]]$WorkbookIds, [bool]$OnlyUpdated, [string]$Category = '', $SnapshotPins = $null) {
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

    $jobId = 'job_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '_' + ([Guid]::NewGuid().ToString('N').Substring(0,8))
    $jobDir = Get-RenderJobDir $Language
    $inputPath = Join-Path $jobDir "$jobId.input.json"
    $statusPath = Join-Path $jobDir "$jobId.status.json"
    $stdoutPath = Join-Path $jobDir "$jobId.out.log"
    $stderrPath = Join-Path $jobDir "$jobId.err.log"
    $initialTotal = if ($OnlyUpdated -and $ids.Count -eq 0) { $registeredCount } else { $ids.Count }
    $initialMessage = 'PDF作成が必要なExcelはありません。'
    if ($registeredCount -eq 0) { $initialMessage = '登録済みExcelがありません。' }
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
    Write-JsonFile $inputPath ([ordered]@{ jobId = $jobId; mode = $Language; workbookIds = @($ids); onlyUpdated = $OnlyUpdated; category = $categoryNormalized; snapshotPins = $pinsOut; statusPath = $statusPath; stdoutPath = $stdoutPath; stderrPath = $stderrPath })
    if ($registeredCount -eq 0 -or ($ids.Count -eq 0 -and -not $OnlyUpdated)) {
        $initial.status = 'completed'; $initial.percent = 100
        if ($registeredCount -eq 0) { $initial.message = '登録済みExcelがありません。' } else { $initial.message = 'PDF作成が必要なExcelはありません。' }
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
            Set-NoteProperty $statusToUpdate 'message' 'PDF作成プロセスを起動しました。Excelを準備しています。'
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
    $statusMessage = 'Excelを準備しています。'
    if ($onlyUpdated -and -not $explicitIds) { $statusMessage = 'PDF作成対象を確認しています。' }
    $status = [pscustomobject][ordered]@{ ok = $true; jobId = $jobId; status = 'running'; total = $total; completed = 0; failed = 0; percent = 3; message = $statusMessage; currentWorkbookId = ''; currentWorkbookName = ''; currentSheet = ''; processId = $processId; stdoutPath = $stdoutPath; stderrPath = $stderrPath; results = @(); errors = @(); startedAt = New-NowIso; updatedAt = New-NowIso; stateSavedAt = '' }
    Write-RenderJobStatus $statusPath $status
    $excel = $null
    try {
        if ($onlyUpdated -and -not $explicitIds) {
            $status.status = 'scanning'
            $status.message = 'PDF作成が必要なExcelを確認しています。'
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
            if ($total -eq 0) {
                $status.status = 'completed'
                $status.message = 'PDF作成が必要なExcelはありません。'
                Write-RenderJobStatus $statusPath $status
                return
            }
            $status.message = "$total 件のPDF作成を開始します。"
            Write-RenderJobStatus $statusPath $status
        } elseif ($total -eq 0) {
            $status.status = 'completed'
            $status.percent = 100
            $status.message = 'PDF作成が必要なExcelはありません。'
            Write-RenderJobStatus $statusPath $status
            return
        }
        if ($total -gt 0) {
            Reset-RenderEnvironmentForJob   # V5-§6.4: ジョブ開始ごとに環境を取り直す
            $status.message = 'Excelを起動しています。'
            $status.percent = [Math]::Max([int]$status.percent, 8)
            Write-RenderJobStatus $statusPath $status
            $excel = New-ExcelApplicationForRender
            $status.message = 'Excelの起動が完了しました。PDF化を開始します。'
            $status.percent = [Math]::Max([int]$status.percent, 10)
            Write-RenderJobStatus $statusPath $status
        }
        $index = 0
        $deferredAnalyses = @()
        foreach ($id in $ids) {
            $index++
            $name = Get-WorkbookDisplayForStatus $language $id
            $status.currentWorkbookId = $id
            $status.currentWorkbookName = $name
            $status.currentSheet = ''
            $status.message = "$index / $total 件目: $name をPDF化しています。"
            $status.percent = [int][Math]::Max(10, [Math]::Floor((($index - 1) / [Math]::Max(1, $total)) * 100))
            Write-RenderJobStatus $statusPath $status
            $callback = {
                param($stage, $workbookId, $sheetName)
                $status.currentWorkbookId = [string]$workbookId
                $status.currentWorkbookName = $name
                $status.currentSheet = [string]$sheetName
                if ($stage -eq 'open') { $status.message = "$index / $total 件目: $name を開いています。" }
                elseif ($stage -eq 'sheet-setup') { $status.message = "$index / $total 件目: $name / シート $sheetName の印刷設定を調整しています。" }
                elseif ($stage -eq 'batch') { $status.message = "$index / $total 件目: $name の複数シートをまとめてPDF化しています。" }
                elseif ($stage -eq 'split') { $status.message = "$index / $total 件目: $name のPDFをシート別に分けています。" }
                elseif ($stage -eq 'sheet') { $status.message = "$index / $total 件目: $name / シート $sheetName をPDF化しています。" }
                else { $status.message = "$index / $total 件目: $name を準備しています。" }
                $status.percent = [int][Math]::Floor((($index - 1 + 0.35) / [Math]::Max(1, $total)) * 100)
                Write-RenderJobStatus $statusPath $status
            }
            try {
                # V5-P0: 自動処理が固定した検知版があれば、その版でレンダリングする。
                $pin = Get-DataProperty (Get-DataProperty $job 'snapshotPins' $null) $id $null
                $pinSnapshot = [string](Get-DataProperty $pin 'snapshotId' '')
                $pinHash = [string](Get-DataProperty $pin 'expectedHash' '')
                $r = Render-Workbook $language $id $excel $true $callback '' $pinSnapshot $pinHash
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
            $status.percent = [int][Math]::Floor((([int]$status.completed + [int]$status.failed) / [Math]::Max(1, $total)) * 100)
            Write-RenderJobStatus $statusPath $status
        }
        if ([int]$status.failed -gt 0) {
            $status.status = 'completed-with-errors'
            $status.message = "$($status.failed) 件でエラーが発生しました。赤いエラー表示を確認してください。"
        } else {
            $status.status = 'completed'
            $status.message = "$($status.completed) 件のPDF作成が完了しました。"
        }
        $status.percent = 100
        Set-NoteProperty $status 'stateSavedAt' (New-NowIso)
        Write-RenderJobStatus $statusPath $status
    } catch {
        $status.status = 'failed'
        $status.message = $_.Exception.Message
        $status.errors = @($status.errors) + @([ordered]@{ error = $_.Exception.Message; userError = (ConvertTo-UserRenderError $_.Exception.Message); detail = [string]$_ })
        Write-RenderJobStatus $statusPath $status
        throw
    } finally {
        Close-ExcelApplicationForRender $excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # V5-P1: Excel を閉じ、レンダリングロックも解放してから解析する。
        # 比較用の再レンダリングが必要になっても、改めて共通ロックを取り直せる。
        foreach ($pending in @($deferredAnalyses | Where-Object { $null -ne $_ })) {
            try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered) }
            catch { Write-Warning ('画像ハッシュの解析に失敗しました: ' + $_.Exception.Message) }
        }
    }
}


function Reorder-Pages([string]$Language, $Body) {
    $cat = Require-WorkbookCategory ([string]$Body.category)
    return Update-StructureLocked $Language {
        param($structure)
        $allowed = @(Get-VolumeList $Language)
        $volumeObject = $Body.volumes
        if ($null -eq $volumeObject) { throw [System.ArgumentException]::new('volumes が必要です。') }
        $workbookIds = @{}
        foreach ($wb in @(Get-Array $structure.workbooks | Where-Object { Test-WorkbookCategory $_ $cat })) { $workbookIds[[string]$wb.workbookId] = $true }
        $pageMap = @{}
        foreach ($page in @(Get-Array $structure.pages | Where-Object { $workbookIds.ContainsKey([string]$_.workbookId) })) { $pageMap[(Resolve-PageId $page)] = $page }
        $affected = New-Object System.Collections.Generic.HashSet[string]
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
        return [ordered]@{ pages=$structure.pages; affectedVolumes=@($affected); updatedAt=(New-NowIso) }
    }
}

function Update-Page([string]$Language, $Body) {
    $cat=Require-WorkbookCategory ([string]$Body.category)
    return Update-StructureLocked $Language {
        param($structure)
        $page=@(Get-Array $structure.pages|Where-Object{(Resolve-PageId $_)-eq [string]$Body.pageId}|Select-Object -First 1);if($page.Count -eq 0){throw "Pageが見つかりません: $($Body.pageId)"};$p=$page[0]
        if((Get-PageCategory $structure $p) -ne $cat){throw [ArgumentException]::new('指定カテゴリのページではありません。')}
        $beforeVol=[string]$p.volume;$beforeEnabled=[bool](Get-DataProperty $p 'enabled' $true);$beforeNum=[string]$p.numberingMode;$structural=$false
        if($null -ne $Body.title){Set-NoteProperty $p 'title' ([string]$Body.title)}
        if($null -ne $Body.volume){if((Get-VolumeList $Language)-notcontains [string]$Body.volume){throw [ArgumentException]::new('不正なvolumeです。')};Set-NoteProperty $p 'volume' ([string]$Body.volume);$structural=$true}
        if($null -ne $Body.numberingMode){if(@('none','visible')-notcontains [string]$Body.numberingMode){throw [ArgumentException]::new('不正なnumberingModeです。')};Set-NoteProperty $p 'numberingMode' ([string]$Body.numberingMode);Set-NoteProperty $p 'numberingManual' $true;$structural=$true}
        if($null -ne $Body.numberingManual){Set-NoteProperty $p 'numberingManual' ([bool]$Body.numberingManual);$structural=$true}
        if($null -ne $Body.resetNumbering -and [bool]$Body.resetNumbering){Set-NoteProperty $p 'numberingManual' $false;$structural=$true}
        if($null -ne $Body.enabled){Set-NoteProperty $p 'enabled' ([bool]$Body.enabled);$structural=$true}
        Set-NoteProperty $p 'updatedAt' (New-NowIso);foreach($vol in @(Get-VolumeList $Language)){[void](Renumber-VolumeOrder $structure $vol $cat)};Apply-DefaultNumberingPerVolume $Language $structure $cat
        if($structural){$affected=@($beforeVol,[string]$p.volume)|Where-Object{$_ -and $_ -ne 'none'}|Select-Object -Unique;Mark-VolumeNeedsRebuild $structure $Language $cat @($affected) 'reorder' 'ページ構成を変更しました'}
        return $p
    }
}

function Confirm-Page([string]$Language, $Body) {
    $cat=Require-WorkbookCategory ([string]$Body.category)
    return Update-StructureLocked $Language { param($structure) $page=@(Get-Array $structure.pages|Where-Object{(Resolve-PageId $_)-eq [string]$Body.pageId}|Select-Object -First 1);if($page.Count -eq 0){throw "Pageが見つかりません: $($Body.pageId)"};if((Get-PageCategory $structure $page[0]) -ne $cat){throw [ArgumentException]::new('指定カテゴリのページではありません。')};switch([string]$Body.action){'confirm'{Set-NoteProperty $page[0] 'status' 'confirmed'}'reject'{Set-NoteProperty $page[0] 'status' 'rejected'}default{throw [ArgumentException]::new('action は confirm または reject を指定してください。')}};Set-NoteProperty $page[0] 'updatedAt' (New-NowIso);return $page[0] }
}

function Sort-PagesBySheet([string]$Language, $Body) {
    $cat = Require-WorkbookCategory ([string]$Body.category)
    return Update-StructureLocked $Language {
        param($structure)
        $volumes = @()
        if ($Body.volumes) {
            $volumes = @(Get-Array $Body.volumes | ForEach-Object { [string]$_ })
        } else {
            $volumes = @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })
        }
        $wbMap = @{}
        foreach ($wb in @(Get-Array $structure.workbooks | Where-Object { Test-WorkbookCategory $_ $cat })) {
            $wbMap[[string]$wb.workbookId] = $wb
        }
        $affected = New-Object System.Collections.Generic.HashSet[string]
        foreach ($volume in @($volumes | Select-Object -Unique)) {
            if ((Get-VolumeList $Language) -notcontains $volume) { throw [ArgumentException]::new("不正なvolumeです: $volume") }
            $current = @(Get-Array $structure.pages | Where-Object {
                $wbMap.ContainsKey([string]$_.workbookId) -and (
                    ($volume -eq 'none' -and ([string]$_.volume -eq 'none' -or $_.enabled -eq $false)) -or
                    ($volume -ne 'none' -and [string]$_.volume -eq $volume -and $_.enabled -ne $false)
                )
            } | Sort-Object {[double](Get-DataProperty $_ 'order' 0)}, {Resolve-PageId $_})
            $sorted = @($current | Sort-Object `
                @{Expression={Get-SheetOrderNumber ([string]$_.sheetName)};Ascending=$true}, `
                @{Expression={Get-FileOrderNumber ([string]$wbMap[[string]$_.workbookId].fileName)};Ascending=$true}, `
                @{Expression={[string]$wbMap[[string]$_.workbookId].fileName};Ascending=$true}, `
                @{Expression={Resolve-PageId $_};Ascending=$true})
            $beforeIds = @($current | ForEach-Object { Resolve-PageId $_ })
            $afterIds = @($sorted | ForEach-Object { Resolve-PageId $_ })
            $inputChanged = (($beforeIds -join "`n") -ne ($afterIds -join "`n"))
            for ($i=0; $i -lt $sorted.Count; $i++) {
                $expected = ($i + 1) * 10
                if ([double](Get-DataProperty $sorted[$i] 'order' 0) -ne $expected) { $inputChanged = $true }
                Set-NoteProperty $sorted[$i] 'order' $expected
                Set-NoteProperty $sorted[$i] 'orderManual' $false
                Set-NoteProperty $sorted[$i] 'updatedAt' (New-NowIso)
            }
            if ($inputChanged -and $volume -ne 'none') { [void]$affected.Add($volume) }
        }
        Apply-DefaultNumberingPerVolume $Language $structure $cat
        if ($affected.Count -gt 0) {
            Mark-VolumeNeedsRebuild $structure $Language $cat @($affected) 'reorder' 'ページをシート名順に並べ替えました'
        }
        return [ordered]@{ pages=$structure.pages; category=$cat; affectedVolumes=@($affected) }
    }
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
    throw ("最終PDFの作成に必要なJava Runtimeが見つかりません(探した場所: {0})。この場所にjava.exeがあるのにこのエラーが出る場合は、server.ps1の置き場所(AppRoot)がずれています。app\logs\startup-*-latest.log の AppRoot 行を確認してください。java.exe自体がない場合は app\tools\install-thirdparty.cmd を実行してから、もう一度PDFを出力してください。" -f $direct)
}

function Get-JavaRuntimeSignature {
    # V5-P1(#11): 画像ハッシュの環境指紋に含める Java 版の署名。
    # `java -version` の出力(バージョン+ビルド)を採取してハッシュ化し、Java 更新で必ず値が変わるようにする。
    if (-not [string]::IsNullOrWhiteSpace($Script:JavaRuntimeSignature)) { return $Script:JavaRuntimeSignature }
    $sig = 'unknown'
    try {
        $java = Resolve-JavaExe
        # `java -version` はバージョン情報を stderr に出すため 2>&1 で取り込む。
        $out = [string](Invoke-NativeCapture $java @('-version')).text
        $norm = ($out -replace '\s+', ' ').Trim()
        if (-not [string]::IsNullOrWhiteSpace($norm)) { $sig = (Get-Sha256Text $norm).Substring(0, 16) }
    } catch { $sig = 'unknown' }
    $Script:JavaRuntimeSignature = $sig
    return $sig
}

function Get-Sha256Text([string]$Text) {
    $sha=[Security.Cryptography.SHA256]::Create();try{$bytes=[Text.Encoding]::UTF8.GetBytes($Text);return 'sha256:'+([BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-','').ToLowerInvariant())}finally{$sha.Dispose()}
}

function Get-FinalBuildInputSnapshot($Structure,[string]$Language,[string]$Volume,[string]$Category) {
    $cat=Require-WorkbookCategory $Category
    if(@(Get-VolumeList $Language|Where-Object{$_ -ne 'none'}) -notcontains $Volume){throw [ArgumentException]::new("不正な成果物です: $Volume")}
    $workspace=Get-WorkspacePath $Language;$wbMap=@{};$targets=@(Get-Array $Structure.workbooks|Where-Object{Test-WorkbookCategory $_ $cat});foreach($wb in $targets){$wbMap[[string]$wb.workbookId]=$wb}
    $pages=@(Get-Array $Structure.pages|Where-Object{$_.enabled -eq $true -and [string]$_.volume -eq $Volume -and $wbMap.ContainsKey([string]$_.workbookId)}|Sort-Object {[double]$_.order},{Resolve-PageId $_})
    $blockers=@();$manifest=@();$fpPages=@()
    if($pages.Count -eq 0){$blockers+= [ordered]@{code='no-pages';pageTitle='';workbookName='';message='対象ページがありません。ページ構成を確認してください。'}}
    foreach($p in $pages){$wb=$wbMap[[string]$p.workbookId];$title=[string]$p.title;$wbName=[string]$wb.fileName;$rel=[string]$p.contentPdf;$full='';$size=0L;$ticks=0L
        if(-not $rel){$blockers+=[ordered]@{code='content-missing';pageTitle=$title;workbookName=$wbName;message='PDF未作成のページがあります。先にPDF作成してください。'}}
        else{try{$full=[IO.Path]::GetFullPath((Join-Path $workspace $rel));$root=[IO.Path]::GetFullPath($workspace);if(-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)){$root+=[IO.Path]::DirectorySeparatorChar};if(-not $full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'outside'};if(-not(Test-Path $full)){throw 'missing'};$it=Get-Item $full;$size=$it.Length;$ticks=$it.LastWriteTimeUtc.Ticks}catch{$blockers+=[ordered]@{code='content-file-missing';pageTitle=$title;workbookName=$wbName;message='レンダリング済みPDFが見つかりません。先にPDF作成してください。'}}}
        if(-not(Test-WorkbookRenderIsCurrent $wb)){$blockers+=[ordered]@{code='stale-content';pageTitle=$title;workbookName=$wbName;message='元Excelが更新されています。先にPDF作成してください。'}}
        elseif(-not(Test-WorkbookRenderedSheetContains $wb ([string]$p.sheetName))){$blockers+=[ordered]@{code='sheet-changed';pageTitle=$title;workbookName=$wbName;message='Excelのシート構成が変わっています。先にPDF作成してください。'}}
        elseif(-not(Test-PageContentMatchesWorkbookVersion $p $wb)){$blockers+=[ordered]@{code='old-content';pageTitle=$title;workbookName=$wbName;message='古いPDF参照が残っています。先にPDF作成してください。'}}
        $normalized=($rel -replace '\\','/').ToLowerInvariant();$fpPages+=[ordered]@{pageId=(Resolve-PageId $p);order=[double]$p.order;enabled=[bool]$p.enabled;volume=[string]$p.volume;numberingMode=[string]$p.numberingMode;contentPdf=$normalized;contentPdfSize=$size;contentPdfLastWriteUtcTicks=$ticks;lastRenderedVersionId=[string]$wb.lastRenderedVersionId}
        if($full){$manifest+=[ordered]@{pageId=(Resolve-PageId $p);title=$title;sourcePdf=$full;numberingMode=[string]$p.numberingMode;punchShiftPt=(Convert-CmToPt 0.2)}}
    }
    $input=[ordered]@{composerProfileVersion=$Script:FinalPdfComposerProfileVersion;language=$Language;category=$cat;volume=$Volume;pages=$fpPages};$json=ConvertTo-Json $input -Depth 20 -Compress;$fingerprint=Get-Sha256Text $json
    return [ordered]@{language=$Language;category=$cat;volume=$Volume;pages=$pages;manifestPages=$manifest;blockers=@($blockers);pageCount=$pages.Count;projectId=(Get-ProjectIdFromWorkbooks $targets);fingerprint=$fingerprint;fingerprintInput=$input}
}

function Get-FinalBuildFingerprint($Snapshot) { return [string](Get-DataProperty $Snapshot 'fingerprint' '') }

function Get-FinalBuildReadiness($Structure,[string]$Language,[string]$Volume,[string]$Category) {
    $cat=Require-WorkbookCategory $Category;$snap=Get-FinalBuildInputSnapshot $Structure $Language $Volume $cat;$key=Get-VolumeStateKey $Volume $cat;$v=Get-DataProperty $Structure.volumes $key (New-EmptyVolumeState);$built=[string](Get-DataProperty $v 'builtFingerprint' '');$current=[string]$snap.fingerprint;$out=[string](Get-DataProperty $v 'outputPdf' '');$exists=(-not [string]::IsNullOrWhiteSpace($out))-and(Test-Path -LiteralPath $out);$status='not-built';if($built){if($built -ne $current -or $snap.blockers.Count -gt 0){$status='needs-rebuild'}else{$status='built'}};$display=if($snap.blockers.Count -gt 0){'blocked'}elseif(-not $built){'not-built'}elseif($built -ne $current){'needs-rebuild'}elseif(-not $exists){'output-missing'}else{'built'}
    Set-NoteProperty $v 'status' $status
    $reasons=@(Get-Array (Get-DataProperty $v 'staleReasons' @()));if($display -eq 'needs-rebuild' -and $reasons.Count -eq 0){$reasons=@([ordered]@{type='fingerprint';at=(Get-DataProperty $Structure 'updatedAt' $null);detail='最終PDFの入力が変更されました'})}
    return [ordered]@{canBuild=($snap.blockers.Count -eq 0 -and $snap.pageCount -gt 0);pageCount=$snap.pageCount;status=$status;displayState=$display;builtFingerprint=$built;currentFingerprint=$current;outputPdf=$out;outputPdfExists=[bool]$exists;lastBuiltAt=(Get-DataProperty $v 'lastBuiltAt' $null);blockers=@($snap.blockers);staleReasons=@($reasons);snapshot=$snap}
}

function Get-AllFinalReadiness($Structure,[string]$Language) {
    $result=[ordered]@{};foreach($cat in @('ecm','bod','dmm')){$vols=[ordered]@{};foreach($volume in @(Get-VolumeList $Language|Where-Object{$_ -ne 'none'})){$vols[$volume]=Get-FinalBuildReadiness $Structure $Language $volume $cat};$result[$cat]=[ordered]@{volumes=$vols}};return $result
}

function Build-FinalPdfLegacy([string]$Language,[string]$Volume,[string]$Category='') {
    if ((Get-VolumeList $Language | Where-Object { $_ -ne 'none' }) -notcontains $Volume) { throw [System.ArgumentException]::new('volumeには本体または補足を指定してください。') }
    $cat=Require-WorkbookCategory $Category;$paths=Get-Paths;$workspace=Get-WorkspacePath $Language;try{[void](Scan-Updates $Language $null $false)}catch{}
    $lockPath=Join-Path $workspace "locks\volume_${Volume}_${cat}.lock"
    return Invoke-WithLock $lockPath {
        $composerJar=Join-Path $Script:AppRoot 'lib\pdfbox\ReportPdfComposer.jar';$pdfboxJar=Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar';if(-not(Test-Path $composerJar)){throw 'ReportPdfComposer.jar がありません。'};if(-not(Test-Path $pdfboxJar)){throw 'pdfbox-app.jar がありません。'}
        $snapshotBefore=Update-StructureLocked $Language {param($st) Apply-DefaultNumberingPerVolume $Language $st $cat;return Get-FinalBuildInputSnapshot $st $Language $Volume $cat}
        if($snapshotBefore.blockers.Count -gt 0){throw [InvalidOperationException]::new([string]$snapshotBefore.blockers[0].message)}
        $fpBefore=[string]$snapshotBefore.fingerprint;$projectId=[string]$snapshotBefore.projectId;$outName=Get-OutputFileName $Volume $projectId $cat;$outPath=Join-Path ([string]$paths.outputDir) $outName;$tmp=Join-Path ([string]$paths.outputDir) "~building_${Volume}_${cat}.pdf";if(Test-Path $tmp){Remove-Item $tmp -Force -ErrorAction SilentlyContinue}
        if(Test-Path $outPath){$f=$null;try{$f=[IO.File]::Open($outPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)}catch{throw "出力先の最終PDFが開かれているため上書きできません: $outName"}finally{if($f){$f.Dispose()}}}
        $manifest=[ordered]@{schemaVersion=2;language=$Language;category=$cat;volume=$Volume;projectId=$projectId;inputFingerprint=$fpBefore;outputPdf=$tmp;createdAt=New-NowIso;pageNumber=[ordered]@{font='Arial';fontSize=8;bottomPt=18;format='hyphenated';countHidden=$true};pages=$snapshotBefore.manifestPages};$manifestPath=Join-Path $workspace "exports\manifest_${Volume}_${cat}.json";Write-JsonFile $manifestPath $manifest
        $java=Resolve-JavaExe;$run=Invoke-NativeCapture $java @('-cp',"$composerJar;$pdfboxJar",'ReportPdfComposer','--manifest',$manifestPath);$exit=[int]$run.exitCode;$text=[string]$run.text;if($exit -ne 0){throw "PDFBox組版に失敗しました。exit=$exit`n$text"};if(-not(Test-Path $tmp)-or(Get-Item $tmp).Length -le 0){Remove-Item $tmp -Force -ErrorAction SilentlyContinue;throw '最終PDFを作成できませんでした。'}
        $commit=Update-StructureLocked $Language {param($st)$after=Get-FinalBuildInputSnapshot $st $Language $Volume $cat;if([string]$after.fingerprint -ne $fpBefore){return [ordered]@{changed=$true;after=$after}};Move-Item -LiteralPath $tmp -Destination $outPath -Force;$key=Get-VolumeStateKey $Volume $cat;$v=Get-DataProperty $st.volumes $key $null;if($null -eq $v){$v=New-EmptyVolumeState;Set-NoteProperty $st.volumes $key $v};Set-NoteProperty $v 'builtFingerprint' $fpBefore;Set-NoteProperty $v 'lastBuiltAt' (New-NowIso);Set-NoteProperty $v 'outputPdf' $outPath;Set-NoteProperty $v 'staleReasons' @();Set-NoteProperty $v 'message' $text;$ready=Get-FinalBuildReadiness $st $Language $Volume $cat;if($ready.blockers.Count -gt 0){Set-NoteProperty $v 'status' 'needs-rebuild';Add-StaleReason $v 'excel-updated' '元Excelが更新されたため、PDFを再作成後に最終PDFを再出力してください'}else{Set-NoteProperty $v 'status' 'built'};return [ordered]@{changed=$false;readiness=$ready}}
        if($commit.changed){Remove-Item $tmp -Force -ErrorAction SilentlyContinue;throw 'PDF作成中にページ構成またはPDF入力が変更されました。最新の状態で再度出力してください。'}
        return [ordered]@{volume=$Volume;category=$cat;outputPdf=$outPath;inputFingerprint=$fpBefore;message=$text;readiness=$commit.readiness}
    }
}


function Invoke-FinalBuildAllLegacy([string]$Language, [string]$Category, [string[]]$Volumes) {
    # V4.1互換実装。現在のルートからは呼ばないが、旧形式の復旧・仕様照合用に保持する。
    # V4.1 経路(Build-FinalPdfLegacy)を巻ごとに回す。
    # トランザクション版と同様、ページが無い巻はスキップし、本当のブロッカーは Build-FinalPdfLegacy 側で送出する。
    $cat = Require-WorkbookCategory $Category
    $allowed = @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })
    $requested = @($Volumes | Where-Object { $allowed -contains $_ })
    if ($requested.Count -eq 0) { throw [System.ArgumentException]::new('volumeには本体または補足を指定してください。') }
    $built = @()
    $skipped = @()
    foreach ($v in $requested) {
        $structure = Get-Structure $Language
        $rd = Get-FinalBuildReadiness $structure $Language $v $cat
        if ([int]$rd.pageCount -le 0) { $skipped += $v; continue }
        $built += @(Build-FinalPdfLegacy $Language $v $cat)
    }
    return [ordered]@{ built = @($built); skipped = @($skipped); message = (if ($built.Count -eq 0) { '出力対象がありません。' } else { '' }) }
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
        excelUpdated = @($workbooks | Where-Object { [string]$_.status -eq 'excel-updated' }).Count
        uncheckedPages = @($pages | Where-Object { [string]$_.status -in @('rendered','stale','not-rendered') }).Count
        confirmedPages = @($pages | Where-Object { [string]$_.status -eq 'confirmed' }).Count
        renderErrors = @($workbooks | Where-Object { [string]$_.status -eq 'render-error' }).Count
        pdfReadyWorkbooks = @($workbooks | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.lastRenderedExcelHash) -and [string]$_.status -ne 'render-error' }).Count
        pdfPendingWorkbooks = @($workbooks | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.lastRenderedExcelHash) -or [string]$_.status -in @('new','excel-updated','render-error') }).Count
        pdfReadyPages = @($pages | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.contentPdf) }).Count
        totalWorkbooks = $workbooks.Count
        totalPages = $pages.Count
    }
    if ($configured -and -not $structureLoadError) {
        foreach ($cat in @('ecm','bod','dmm')) {
            foreach ($volume in @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })) {
                $key=Get-VolumeStateKey $volume $cat; $v=Get-DataProperty $structure.volumes $key $null
                if ($null -ne $v) { $out=[string](Get-DataProperty $v 'outputPdf' ''); Set-NoteProperty $v 'outputPdfExists' ((-not [string]::IsNullOrWhiteSpace($out)) -and (Test-Path -LiteralPath $out)) }
            }
        }
    }
    # V5: シート単位の変更判定を state に載せる(履歴が承認されている場合のみ)。
    $changeSummaries = [ordered]@{}
    $inputHistoryOn = $false
    try { $inputHistoryOn = (Test-InputHistoryEnabled) } catch { }
    if ($configured -and -not $structureLoadError -and $inputHistoryOn) {
        foreach ($w in $workbooks) {
            try {
                $cs = Get-WorkbookChangeSummary $Language ([string]$w.workbookId) $w
                if ($null -ne $cs) { $changeSummaries[[string]$w.workbookId] = $cs }
            } catch { }
        }
    }
    $autoSummary = $null
    try { if ($configured) { $autoSummary = Get-AutoStateSummary $Language } } catch { }

    $pdfjsDir = Join-Path $Script:WebRoot 'pdfjs'
    $pdfjsClassic = ((Test-Path -LiteralPath (Join-Path $pdfjsDir 'pdf.min.js')) -and (Test-Path -LiteralPath (Join-Path $pdfjsDir 'pdf.worker.min.js')))
    $pdfjsModule = ((Test-Path -LiteralPath (Join-Path $pdfjsDir 'pdf.min.mjs')) -and (Test-Path -LiteralPath (Join-Path $pdfjsDir 'pdf.worker.min.mjs')))
    $pdfjsMode = 'none'
    if ($pdfjsModule) { $pdfjsMode = 'module' } elseif ($pdfjsClassic) { $pdfjsMode = 'classic' }
    return [ordered]@{
        ok = $true
        mode = $Mode
        language = $Language
        token = $Script:Token
        configured = $configured
        paths = $paths
        structure = $structure
        structureLoadError = $structureLoadError
        finalReadiness = $(if ($configured -and -not $structureLoadError) { Get-AllFinalReadiness $structure $Language } else { [ordered]@{} })
        summary = $summary
        recentErrors = @(Get-Array $workbooks | Where-Object { [string]$_.status -eq 'render-error' -or -not [string]::IsNullOrWhiteSpace([string]$_.lastError) } | ForEach-Object { [ordered]@{ workbookId = [string]$_.workbookId; fileName = [string]$_.fileName; displayName = [string]$_.displayName; message = [string]$_.lastErrorUser; detail = [string]$_.lastError; at = [string]$_.lastErrorAt } })
        pdfjsPresent = ($pdfjsClassic -or $pdfjsModule)
        pdfjsMode = $pdfjsMode
        excelPrintProfileVersion = $Script:ExcelPrintProfileVersion
        inputHistoryEnabled = $inputHistoryOn
        changeSummaries = $changeSummaries
        auto = $autoSummary
        autoRenderInProgress = $Script:AutoRenderInProgress
        shutdownOnTabClose = $false
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
    $fileInfo = [IO.FileInfo]::new($FullPath)
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
    Write-BytesResponse $Context 200 ([IO.File]::ReadAllBytes($full)) 'application/pdf'
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
    $cat=Require-WorkbookCategory $Category;$volume=([string]$Volume).Trim();if(@(Get-VolumeList $Language|Where-Object{$_ -ne 'none'}) -notcontains $volume){throw [ArgumentException]::new('不正な成果物です。')};$structure=Get-Structure $Language;$v=Get-DataProperty $structure.volumes (Get-VolumeStateKey $volume $cat) $null;if($null -eq $v -or -not [string]$v.outputPdf){throw '最終PDFはまだ作成されていません。'};$paths=Get-Paths;$full=[IO.Path]::GetFullPath([string]$v.outputPdf);$root=[IO.Path]::GetFullPath([string]$paths.outputDir);if(-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)){$root+=[IO.Path]::DirectorySeparatorChar};if(-not $full.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw '出力フォルダ外のPDFは表示できません。'};if(-not(Test-Path $full)){throw '最終PDFファイルが見つかりません。'};Write-BytesResponse $Context 200 ([IO.File]::ReadAllBytes($full)) 'application/pdf'
}

function Serve-FinalPdf($Context,[string]$Language) {
    Serve-FinalPdfByVolume $Context $Language ([string]$Context.Request.QueryString['volume']) ([string]$Context.Request.QueryString['category'])
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
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; mode = $Mode; at = New-NowIso }) $true; return
        }
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
            Write-JsonResponse $Context 200 (Request-ServerShutdown ([string]$body.reason)); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/auto/render') {
            $body = Read-BodyJson $Context.Request
            $ids = @()
            if ($body.workbookIds) { $ids = @(Get-Array $body.workbookIds | ForEach-Object { [string]$_ }) }
            $result = Invoke-AutoRender $language $ids
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result }); return
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
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; scannedAt = (New-NowIso); files = (Get-ExcelFilesInSubmission) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/submission/select-and-start') {
            $body = Read-BodyJson $Context.Request
            $initial = [string]$body.initialDir
            if ([string]::IsNullOrWhiteSpace($initial)) { $initial = [string](Get-Paths).submissionDir }
            $selected = Select-FolderDialog '提出フォルダを選択' $initial
            if ([string]::IsNullOrWhiteSpace($selected)) {
                Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; cancelled = $true; state = (Get-StatePayload $language) }); return
            }
            $newPaths = Get-DefaultChildPaths $selected
            Ensure-Package $newPaths
            $config = Get-AppConfig
            $config.lastSubmissionDir = [string]$newPaths.submissionDir
            $config.lastDataDir = [string]$newPaths.dataDir
            $config.lastOutputDir = [string]$newPaths.outputDir
            $config.lastMode = $Mode
            Save-AppConfig $config
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; paths = $newPaths; path = [string]$newPaths.submissionDir; state = (Get-StatePayload $language); scannedAt = (New-NowIso); files = (Get-ExcelFilesInSubmission) }); return
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
            $newPaths = [ordered]@{ submissionDir = [string]$body.submissionDir; dataDir = [string]$body.dataDir; outputDir = [string]$body.outputDir }
            if ([string]::IsNullOrWhiteSpace([string]$newPaths.dataDir) -or [string]::IsNullOrWhiteSpace([string]$newPaths.outputDir)) {
                $defaults = Get-DefaultChildPaths ([string]$newPaths.submissionDir)
                if ([string]::IsNullOrWhiteSpace([string]$newPaths.dataDir)) { $newPaths.dataDir = [string]$defaults.dataDir }
                if ([string]::IsNullOrWhiteSpace([string]$newPaths.outputDir)) { $newPaths.outputDir = [string]$defaults.outputDir }
            }
            Ensure-Package $newPaths
            $config = Get-AppConfig
            $config.lastSubmissionDir = [string]$newPaths.submissionDir
            $config.lastDataDir = [string]$newPaths.dataDir
            $config.lastOutputDir = [string]$newPaths.outputDir
            $config.lastMode = $Mode
            Save-AppConfig $config
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; paths = $newPaths }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/workbooks/register-batch') {
            $body = Read-BodyJson $Context.Request
            $rels = @()
            if ($body.relativePaths) { $rels = @(Get-Array $body.relativePaths | ForEach-Object { [string]$_ }) }
            elseif ($body.relativePath) { $rels = @([string]$body.relativePath) }
            if ($rels.Count -eq 0) { throw '登録するExcelが選択されていません。' }
            $result = Register-WorkbooksBatch $language $rels ([string]$body.category)
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = $result; state = (Get-StatePayload $language) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/workbooks/register') {
            $body = Read-BodyJson $Context.Request
            $result = Register-Workbook $language ([string]$body.relativePath) ([string]$body.category)
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
        if ($method -eq 'POST' -and $path -eq '/api/workbooks/render/start') {
            $body = Read-BodyJson $Context.Request
            $ids = @()
            if ($body.workbookIds) { $ids = @(Get-Array $body.workbookIds | ForEach-Object { [string]$_ }) }
            $result = Start-RenderJob $language $ids ([bool]$body.onlyUpdated) ([string]$body.category)
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
            elseif ($body.onlyUpdated -eq $true) { $ids = @(Get-Array $structure.workbooks | Where-Object { [string]$_.status -in @('new','excel-updated','render-error') -or [string]$_.currentExcelHash -ne [string]$_.lastRenderedExcelHash } | ForEach-Object { [string]$_.workbookId }) }
            else { $ids = @(Get-Array $structure.workbooks | ForEach-Object { [string]$_.workbookId }) }
            $results = @()
            foreach ($id in $ids) {
                try {
                    $r = Render-Workbook $language $id
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
            [void](Save-LayoutSnapshot $language (Require-WorkbookCategory ([string]$body.category)) 'reorder')
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = (Reorder-Pages $language $body) }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/pages/sort-by-sheet') {
            $body=Read-BodyJson $Context.Request
            [void](Save-LayoutSnapshot $language (Require-WorkbookCategory ([string]$body.category)) 'sort-by-sheet')
            Write-JsonResponse $Context 200 ([ordered]@{ok=$true;result=(Sort-PagesBySheet $language $body)});return
        }
        if ($method -eq 'POST' -and $path -eq '/api/pages/update') {
            $body = Read-BodyJson $Context.Request
            # V5: 出力先・ページ番号・出力対象・タイトルの変更もレイアウト履歴に残す。
            [void](Save-LayoutSnapshot $language (Require-WorkbookCategory ([string]$body.category)) 'page-update')
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; page = (Update-Page $language $body) }); return
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
        if ($method -eq 'GET' -and $path -eq '/api/history/render-page') {
            Serve-HistoryRasterPage $Context $language ([string]$Context.Request.QueryString['workbookId']) ([string]$Context.Request.QueryString['snapshotId']) ([string]$Context.Request.QueryString['versionId']) ([string]$Context.Request.QueryString['sheetName']) ([int]$Context.Request.QueryString['pageNumber']); return
        }
        if ($method -eq 'GET' -and $path -eq '/api/history/content-pdf') {
            Serve-HistoryContentPdf $Context $language ([string]$Context.Request.QueryString['workbookId']) ([string]$Context.Request.QueryString['versionId']) ([string]$Context.Request.QueryString['sheetName']) ([string]$Context.Request.QueryString['snapshotId']); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/history/pin') {
            $body = Read-BodyJson $Context.Request
            $wbId = Assert-SafeStorageSegment ([string]$body.workbookId) 'workbookId'
            $snapshotId = Assert-SafeStorageSegment ([string]$body.snapshotId) 'snapshotId'
            if (-not (New-SnapshotPin $language $wbId $snapshotId 'manual' ([ordered]@{ pinnedAt = New-NowIso }))) {
                throw '履歴の保護情報を保存できませんでした。'
            }
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true }); return
        }
        if ($method -eq 'POST' -and $path -eq '/api/history/unpin') {
            $body = Read-BodyJson $Context.Request
            $wbId = Assert-SafeStorageSegment ([string]$body.workbookId) 'workbookId'
            $snapshotId = Assert-SafeStorageSegment ([string]$body.snapshotId) 'snapshotId'
            Remove-SnapshotPin $language $wbId $snapshotId 'manual'
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
        if ($method -eq 'POST' -and $path -eq '/api/auto/run-now') {
            $body = Read-BodyJson $Context.Request
            $wbId = Assert-SafeStorageSegment ([string]$body.workbookId) 'workbookId'
            Write-JsonResponse $Context 200 ([ordered]@{ ok = $true; result = (Request-AutoRunNow $language $wbId) }); return
        }
        Write-JsonResponse $Context 404 ([ordered]@{ ok = $false; error = 'unknown api route' })
    } catch [System.ArgumentException] {
        Write-JsonResponse $Context 400 ([ordered]@{ ok = $false; error = $_.Exception.Message })
    } catch {
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
                    $status = $(if ($_.Exception -is [System.ArgumentException]) { 400 } else { 500 })
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
    # V5-§4.3: 1イベント1ファイル。共有ドライブ上でのJSONL追記は行が混ざるため使わない。
    if (-not (Test-InputHistoryEnabled)) { return }
    try {
        $dir = Join-Path (Get-WorkspacePath $Language) 'history\events'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $safeType = [regex]::Replace([string]$Type, '[^A-Za-z0-9._-]+', '-')
        $name = ('{0}_{1}.json' -f (New-RbId), $safeType)
        $payload = [ordered]@{
            schemaVersion = 1; eventType = [string]$Type; at = New-NowIso
            pcName = $env:COMPUTERNAME; userName = "$env:USERDOMAIN\$env:USERNAME"
            data = $Data
        }
        Write-JsonFile (Join-Path $dir $name) $payload
    } catch { }
}

function Get-SnapshotManifest([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    $path = Join-Path (Get-SnapshotDir $Language $WorkbookId $SnapshotId) 'manifest.json'
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    try { return (Read-JsonFile $path $null) } catch { return $null }
}

function Get-SnapshotIds([string]$Language, [string]$WorkbookId) {
    $dir = Get-WorkbookHistoryDir $Language $WorkbookId
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    return @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue |
             Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'manifest.json') } |
             Sort-Object Name | ForEach-Object { [string]$_.Name })
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
    $sourcePath = Join-Path $dir 'source.xlsx'
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
        $dir = Get-SnapshotPinDir $Language $WorkbookId $SnapshotId
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Write-JsonFile (Join-Path $dir ("{0}.json" -f $PinName)) $Data
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

function New-SnapshotLease([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$Purpose, [string]$JobId, [int]$MinutesValid = 30, [string]$VersionId = '') {
    try {
        $dir = Get-SnapshotLeaseDir $Language $WorkbookId $SnapshotId
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $name = ('{0}_{1}.json' -f $Purpose, $JobId)
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
        Write-HistoryEvent $Language 'input.snapshot.created' ([ordered]@{ workbookId = $WorkbookId; snapshotId = $SnapshotId })
        return $true
    } catch { return $false }
}

function Save-SnapshotSourceFile([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$SourcePath, [string]$SourceHash) {
    # V5-§3.2 手順4-8: コピー -> ハッシュ検証 -> 履歴フォルダへ移動 -> 再ハッシュ。
    if (-not (Test-SourceRetentionEnabled)) { return [ordered]@{ ok = $false; reason = 'retention-not-approved' } }
    $dir = Get-SnapshotDir $Language $WorkbookId $SnapshotId
    if (-not (Test-Path -LiteralPath $dir)) { return [ordered]@{ ok = $false; reason = 'snapshot-missing' } }
    $dest = Join-Path $dir 'source.xlsx'
    if (Test-Path -LiteralPath $dest) { return [ordered]@{ ok = $true; path = $dest; reused = $true } }

    $tmpDir = Join-Path (Get-WorkspacePath $Language) ('state\tmp\' + (New-RbId))
    New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
    $tmp = Join-Path $tmpDir 'source.xlsx'
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
        if ($copyHash -ne (Normalize-FileHash $SourceHash)) { return [ordered]@{ ok = $false; reason = 'hash-mismatch' } }
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
        if ($rendered -and $rendered -ne $capturedHash -and [string]$t.status -ne 'render-error') { Set-NoteProperty $t 'status' 'excel-updated' }
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
            return [ordered]@{ path = [string]$state.sourcePath; snapshotId = $snap; ephemeral = $false; hash = $hash }
        }
        # 現物が消えている: 提出フォルダの現物が同じハッシュなら復元する。
        if (Test-Path -LiteralPath $livePath) {
            $liveHash = Normalize-FileHash (New-Sha256 $livePath)
            if ($liveHash -eq $hash) {
                $r = Save-SnapshotSourceFile $Language $WorkbookId $snap $livePath $hash
                if ([bool]$r.ok) { return [ordered]@{ path = [string]$r.path; snapshotId = $snap; ephemeral = $false; hash = $hash } }
            }
        }
    }

    # 縮退モード、または現物を復元できない場合は一時コピーを作る。
    if (-not (Test-Path -LiteralPath $livePath)) { return [ordered]@{ path = ''; snapshotId = $snap; ephemeral = $false; hash = $hash } }
    $captureId = $(if ([string]::IsNullOrWhiteSpace($JobId)) { New-RbId } else { $JobId })
    $dir = Join-Path (Get-EphemeralJobRoot $Language) $captureId
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = Join-Path $dir 'source.xlsx'
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
    return [ordered]@{ path = $tmp; snapshotId = $snap; ephemeral = $true; hash = $tmpHash; captureId = $captureId }
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
            $all = @(Get-SnapshotIds $Language $workbookId)
            # 未完成世代(manifest なし)を削除
            foreach ($d in @(Get-ChildItem -LiteralPath $wbDir.FullName -Directory -ErrorAction SilentlyContinue)) {
                if (-not (Test-Path -LiteralPath (Join-Path $d.FullName 'manifest.json'))) {
                    if ($d.LastWriteTimeUtc -lt [DateTime]::UtcNow.AddHours(-6)) {
                        Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            }
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
                        Write-HistoryEvent $Language 'input.source.removed' ([ordered]@{ workbookId = $workbookId; snapshotId = $sn; reason = 'retention' })
                    }
                }

                # --- 段階2: pin が1件も無ければフォルダごと削除する ---
                if ($finalPins.Count -eq 0 -and $pins.Count -eq 0) {
                    $m = Get-SnapshotManifest $Language $workbookId $sn
                    if ($null -ne $m -and [string](Get-DataProperty $m 'status' '') -eq 'complete') {
                        Remove-Item -LiteralPath (Get-SnapshotDir $Language $workbookId $sn) -Recurse -Force -ErrorAction SilentlyContinue
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
        $dir = Join-Path (Get-ContentPdfVersionDir $Workspace $WorkbookId $VersionId) 'pins'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $pinPath = Join-Path $dir ("{0}.json" -f $PinName)
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
        $dir = Join-Path (Get-ContentPdfVersionDir $Workspace $WorkbookId $VersionId) 'leases'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $name = ('{0}_{1}.json' -f $Purpose, $JobId)
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
        schemaVersion = 1; workbookId = $WorkbookId; updatedAt = New-NowIso
        pendingSnapshotId = ''; pendingHash = ''; stableCount = 0
        firstDetectedAt = ''; lastSeenAt = ''; quietDeadline = ''
        state = 'idle'; deferReason = ''; ownerPcName = ''; ownerJobId = ''
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

        $currentHash = Normalize-FileHash ([string](Get-DataProperty $w 'currentExcelHash' ''))
        $renderedHash = Normalize-FileHash ([string](Get-DataProperty $w 'lastRenderedExcelHash' ''))
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

        Set-NoteProperty $state 'state' 'rendering'; Set-NoteProperty $state 'deferReason' ''
        Write-AutoState $Language $id $state
        $started = Start-AutoRenderJobForWorkbook $Language $id ([string](Get-DataProperty $state 'pendingSnapshotId' ''))
        if ([bool]$started.ok) {
            Set-NoteProperty $state 'ownerJobId' ([string]$started.jobId)
            Write-HistoryEvent $Language 'render.started' ([ordered]@{ workbookId = $id; jobId = [string]$started.jobId; trigger = 'auto' })
        } else {
            Set-NoteProperty $state 'state' 'waiting'
            Set-NoteProperty $state 'deferReason' 'start-failed'
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

$Script:VisualHashProfileVersion = 2
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
        $run = Invoke-NativeCapture $javaExe @('-Djava.awt.headless=true', '-cp', $cp, 'PdfPageAnalyzer', '--input', $req, '--output', $res, '--dpi', [string]$Script:VisualHashDpi)
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

function Write-RenderRecord([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId, [string]$Purpose, [bool]$ContentPdfRetained, $Analysis) {
    # V5-IV-3: レンダリング結果は renders\<versionId>\ に世代ごとに置き、書いたら変更しない。
    if ([string]::IsNullOrWhiteSpace($SnapshotId) -or [string]::IsNullOrWhiteSpace($VersionId)) { return }
    try {
        $dir = Get-RenderRecordDir $Language $WorkbookId $SnapshotId $VersionId
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $manifestPath = Join-Path $dir 'render-manifest.json'
        if (-not (Test-Path -LiteralPath $manifestPath)) {
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
                visualHashProfile = (Get-VisualHashProfile)
            })
        }
        $hashPath = Join-Path $dir 'visual-hashes.json'
        if ($null -ne $Analysis -and -not (Test-Path -LiteralPath $hashPath)) {
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
    } catch { }
}

function Get-VisualHashes([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId) {
    $p = Join-Path (Get-RenderRecordDir $Language $WorkbookId $SnapshotId $VersionId) 'visual-hashes.json'
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    try { return (Read-JsonFile $p $null) } catch { return $null }
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


function Render-SnapshotForComparison([string]$Language, [string]$WorkbookId, [string]$SnapshotId) {
    # structure.json等は変更しないが、画像ハッシュと同一世代のcontent PDFは保持する。
    # これにより再レンダリング比較でも、判定対象と画面表示対象が必ず一致する。
    $state = Get-SnapshotSourceState $Language $WorkbookId $SnapshotId
    if (-not [bool]$state.sourceRetained) { return [ordered]@{ ok = $false; reason = 'source-missing' } }
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
            $work = Join-Path $tmpDir 'source.xlsx'
            Copy-FileSharedRead ([string]$state.sourcePath) $work
            try { Unblock-File -LiteralPath $work -ErrorAction SilentlyContinue } catch { }
            try { [void](Remove-XlsxHeaderFooterXml $work) } catch { }
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
                    if (-not (([int]$ws.Visible -eq -1) -and $sheetName -match '^[0-9]+$')) { continue }
                    $outPdf = Join-Path $contentDir ("{0}.pdf" -f $sheetName)
                    [void](Export-WorksheetToPdfSafe $excel $book $ws $outPdf $sheetName $false)
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
            Write-RenderRecord $Language $WorkbookId $SnapshotId $versionId 'comparison' $true ([pscustomobject]$analysis.result)
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

function Compare-SnapshotVisual([string]$Language, [string]$WorkbookId, [string]$CurrentSnapshotId, [string]$CurrentVersionId) {
    # 表現は「見落としなし」ではなく「最終PDFの見た目を基準とした高精度な判定」。
    $result = [ordered]@{
        schemaVersion = 2; status = 'unavailable'
        scope = 'automatic'
        baselineSnapshotId = ''; baselineVersionId = ''
        currentSnapshotId = $CurrentSnapshotId; currentVersionId = $CurrentVersionId
        comparedAt = New-NowIso; method = ''
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
        $re = Render-SnapshotForComparison $Language $WorkbookId $baseSnap
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

    $baseMap = @{}
    foreach ($s in @(Get-Array (Get-DataProperty $base 'sheets' @()))) { $baseMap[[string]$s.sheetName] = $s }
    $curNames = @{}
    $changed = @(); $unchanged = @(); $unknown = @(); $added = @(); $removed = @()
    foreach ($s in @(Get-Array (Get-DataProperty $cur 'sheets' @()))) {
        $name = [string]$s.sheetName
        $curNames[$name] = $true
        if ([string](Get-DataProperty $s 'status' '') -ne 'ok') { $unknown += $name; continue }
        if (-not $baseMap.ContainsKey($name)) { $added += $name; continue }
        $b = $baseMap[$name]
        if ([string](Get-DataProperty $b 'status' '') -ne 'ok') { $unknown += $name; continue }
        if ((Normalize-FileHash ([string](Get-DataProperty $b 'sheetVisualHash' ''))) -eq (Normalize-FileHash ([string](Get-DataProperty $s 'sheetVisualHash' ''))) -or
            (Test-SheetVisualEquivalent $b $s)) { $unchanged += $name }
        else { $changed += $name }
    }
    $result.status = 'complete'
    $result.baselineSnapshotId = $baseSnap
    $result.baselineVersionId = $baseVer
    $result.method = $method
    # V5-P1: 今回版のシートだけを回すと、削除されたシートが差分に出ない。
    foreach ($k in @($baseMap.Keys)) { if (-not $curNames.ContainsKey([string]$k)) { $removed += [string]$k } }
    $result.changedSheets = @($changed)
    $result.unchangedSheets = @($unchanged)
    $result.unknownSheets = @($unknown)
    $result.addedSheets = @($added)
    $result.removedSheets = @($removed)
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
    Write-HistoryEvent $Language 'compare.completed' ([ordered]@{ workbookId = $WorkbookId; snapshotId = $CurrentSnapshotId; changed = @($changed); unknown = @($unknown) })
    return $saved
}

function Invoke-PostRenderAnalysis([string]$Language, [string]$WorkbookId, [string]$SnapshotId, [string]$VersionId, $Rendered) {
    # V5-§6.5: PDF作成のクリティカルパスの外。失敗しても PDF 作成は成功扱いのまま。
    if (-not (Test-InputHistoryEnabled)) { return $null }
    if ([string]::IsNullOrWhiteSpace($SnapshotId)) { return $null }
    try {
        $sheets = @()
        foreach ($r in @(Get-Array $Rendered)) {
            $pdf = [string](Get-DataProperty $r 'pdf' '')
            $name = [string](Get-DataProperty $r 'sheetName' '')
            if ($pdf -and $name -and (Test-Path -LiteralPath $pdf)) {
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
        Write-RenderRecord $Language $WorkbookId $SnapshotId $VersionId 'normal' $true $parsed
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
        return $cmp
    } catch {
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
        $dir = Join-Path (Get-RenderRecordDir $Language $WorkbookId $snap $ver) 'comparisons'
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

function Get-ArchiveRoot([string]$Language) { return (Join-Path (Get-WorkspacePath $Language) 'exports\archive') }
function Get-LayoutHistoryDir([string]$Language, [string]$Category) { return (Join-Path (Get-WorkspacePath $Language) (Join-Path 'layout-history' $Category)) }
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

function Save-LayoutSnapshot([string]$Language, [string]$Category, [string]$Reason, $Structure = $null) {
    if (-not (Test-InputHistoryEnabled)) { return '' }
    try {
        $st = $Structure
        if ($null -eq $st) { $st = Get-Structure $Language }
        $pages = @()
        foreach ($p in @(Get-Array $st.pages)) {
            $wb = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq [string]$p.workbookId } | Select-Object -First 1)
            if ($wb.Count -eq 0) { continue }
            if (-not (Test-WorkbookCategory $wb[0] $Category)) { continue }
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
            }
        }
        if ($pages.Count -eq 0) { return '' }
        $dir = Get-LayoutHistoryDir $Language $Category
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $id = New-RbId
        Write-JsonFile (Join-Path $dir ("{0}.json" -f $id)) ([ordered]@{
            schemaVersion = 1; snapshotId = $id; language = $Language; category = $Category
            createdAt = New-NowIso; reason = $Reason; pages = $pages
        })
        Write-HistoryEvent $Language 'layout.changed' ([ordered]@{ category = $Category; reason = $Reason; snapshotId = $id; pageCount = $pages.Count })
        return $id
    } catch { return '' }
}

function Get-LayoutSnapshots([string]$Language, [string]$Category) {
    $dir = Get-LayoutHistoryDir $Language $Category
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
            }
        } catch { }
    }
    return $out
}

function Read-LayoutSnapshot([string]$Language, [string]$Category, [string]$SnapshotId) {
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $p = Join-Path (Get-LayoutHistoryDir $Language $Category) ("{0}.json" -f $safeSnapshotId)
    if (-not (Test-Path -LiteralPath $p)) { throw "レイアウト履歴が見つかりません: $SnapshotId" }
    return (Read-JsonFile $p $null)
}

function Get-LayoutRestorePreview([string]$Language, [string]$Category, [string]$SnapshotId) {
    $snap = Read-LayoutSnapshot $Language $Category $SnapshotId
    $st = Get-Structure $Language
    $currentIds = @{}
    foreach ($p in @(Get-Array $st.pages)) {
        $wb = @(Get-Array $st.workbooks | Where-Object { [string]$_.workbookId -eq [string]$p.workbookId } | Select-Object -First 1)
        if ($wb.Count -eq 0) { continue }
        if (-not (Test-WorkbookCategory $wb[0] $Category)) { continue }
        $currentIds[(Resolve-PageId $p)] = $p
    }
    $applied = 0; $pastOnly = @(); $volumeChanges = @()
    $snapIds = @{}
    foreach ($sp in @(Get-Array (Get-DataProperty $snap 'pages' @()))) {
        $pageKey = [string]$sp.pageId
        $snapIds[$pageKey] = $true
        if (-not $currentIds.ContainsKey($pageKey)) { $pastOnly += $pageKey; continue }
        $applied++
        $cur = $currentIds[$pageKey]
        if ([string]$cur.volume -ne [string]$sp.volume) {
            $volumeChanges += [ordered]@{ pageId = $pageKey; title = [string]$cur.title; from = [string]$cur.volume; to = [string]$sp.volume }
        }
    }
    $currentOnly = @($currentIds.Keys | Where-Object { -not $snapIds.ContainsKey($_) })
    return [ordered]@{
        snapshotId = $SnapshotId
        createdAt = [string](Get-DataProperty $snap 'createdAt' '')
        reason = [string](Get-DataProperty $snap 'reason' '')
        appliedPageCount = $applied
        pastOnlyPageIds = @($pastOnly)
        currentOnlyPageIds = @($currentOnly)
        volumeChanges = @($volumeChanges)
        requiresRebuild = $true
    }
}

function Restore-LayoutSnapshot([string]$Language, [string]$Category, [string]$SnapshotId) {
    $snap = Read-LayoutSnapshot $Language $Category $SnapshotId
    # 復元の直前にも保存しておき、「復元を取り消す」を可能にする。
    $undoId = Save-LayoutSnapshot $Language $Category 'pre-restore'
    $applied = Update-StructureLocked $Language {
        param($st)
        $map = @{}
        foreach ($p in @(Get-Array $st.pages)) { $map[(Resolve-PageId $p)] = $p }
        $n = 0
        foreach ($sp in @(Get-Array (Get-DataProperty $snap 'pages' @()))) {
            $pageKey = [string]$sp.pageId
            if (-not $map.ContainsKey($pageKey)) { continue }
            $p = $map[$pageKey]
            # V5-§4.2: 適用してよいのはレイアウト項目のみ。
            # contentPdf / status / warnings / currentExcelHash / lastRendered* / volumes は触らない。
            Set-NoteProperty $p 'title' ([string]$sp.title)
            Set-NoteProperty $p 'volume' ([string]$sp.volume)
            Set-NoteProperty $p 'enabled' ([bool]$sp.enabled)
            Set-NoteProperty $p 'order' ([double]$sp.order)
            Set-NoteProperty $p 'orderManual' ([bool]$sp.orderManual)
            Set-NoteProperty $p 'numberingMode' ([string]$sp.numberingMode)
            Set-NoteProperty $p 'numberingManual' ([bool]$sp.numberingManual)
            Set-NoteProperty $p 'updatedAt' (New-NowIso)
            $n++
        }
        Apply-DefaultNumberingPerVolume $Language $st $Category
        $vols = @(Get-VolumeList $Language | Where-Object { $_ -ne 'none' })
        Mark-VolumeNeedsRebuild $st $Language $Category $vols 'layout-restored' 'ページ構成を過去の状態へ戻しました'
        return $n
    }
    Write-HistoryEvent $Language 'layout.restored' ([ordered]@{ category = $Category; snapshotId = $SnapshotId; undoSnapshotId = $undoId; appliedPageCount = $applied })
    return [ordered]@{ appliedPageCount = $applied; undoSnapshotId = $undoId }
}

# ---- 最終PDFアーカイブ (V5-§4.1) -----------------------------------

function New-FinalArchive([string]$Language, [string]$Category, [string]$Volume, [string]$BuildId, [string]$OutputPdf, $Manifest, $Snapshot) {
    # 冪等: 一時フォルダで完成させてから buildId フォルダへ移動する。
    try {
        $target = Join-Path (Join-Path (Join-Path (Get-ArchiveRoot $Language) $Category) $Volume) $BuildId
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
            schemaVersion = 1; buildId = $BuildId; language = $Language; category = $Category; volume = $Volume
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
        foreach ($sw in $sourceWorkbooks) {
            $wbId = [string](Get-DataProperty $sw 'workbookId' '')
            $snapshotId = [string](Get-DataProperty $sw 'snapshotId' '')
            $versionId = [string](Get-DataProperty $sw 'versionId' '')
            if ([string]::IsNullOrWhiteSpace($wbId) -or [string]::IsNullOrWhiteSpace($snapshotId)) { continue }
            # V5-P0(#4): pin を作れなければトランザクションを失敗させる。
            # pin は正式PDFが参照する source/content-pdf を後日の掃除から守る唯一の仕組みであり、
            # 作成に失敗したまま completed にすると、参照先が削除され得る。
            $pinOk = New-SnapshotPin $Language $wbId $snapshotId ("final-pdf_{0}" -f $BuildId) ([ordered]@{
                buildId = $BuildId; snapshotId = $snapshotId; versionId = $versionId
                volume = $Volume; category = $Category; archivePath = $target
            })
            if (-not $pinOk) { throw ("スナップショット保護(pin)を作成できませんでした: {0} / {1}" -f $wbId, $snapshotId) }
            if ($versionId) {
                $cpPinOk = New-ContentPdfPin (Get-WorkspacePath $Language) $wbId $versionId ("final-pdf_{0}" -f $BuildId) ([ordered]@{
                    buildId = $BuildId; volume = $Volume; category = $Category
                })
                if (-not $cpPinOk) { throw ("content-pdf 保護(pin)を作成できませんでした: {0} / {1}" -f $wbId, $versionId) }
            }
        }
        Write-HistoryEvent $Language 'final.archive.created' ([ordered]@{ category = $Category; volume = $Volume; buildId = $BuildId; path = $target })
        return $target
    } catch {
        # V5-P0: 握りつぶさない。呼出元がロールバックする。
        try { if (Test-Path -LiteralPath ($target + '.staging')) { Remove-Item -LiteralPath ($target + '.staging') -Recurse -Force -ErrorAction SilentlyContinue } } catch { }
        throw ("最終PDFのアーカイブに失敗しました: " + $_.Exception.Message)
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
    $root = Join-Path (Get-ArchiveRoot $Language) $Category
    if (-not (Test-Path -LiteralPath $root)) { return @() }
    $out = @()
    foreach ($volDir in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
        foreach ($b in @(Get-ChildItem -LiteralPath $volDir.FullName -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending | Select-Object -First 50)) {
            try {
                $meta = Read-JsonFile (Join-Path $b.FullName 'metadata.json') $null
                if ($null -eq $meta) { continue }
                $out += [ordered]@{
                    buildId = [string](Get-DataProperty $meta 'buildId' $b.Name)
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

function Invoke-FinalBuildTransaction([string]$Language, [string]$Category, [string[]]$Volumes) {
    # 単体出力もまとめて出力も、必ずこの1本を通る(単体だけ障害復旧が無い状態を作らない)。
    $cat = Require-WorkbookCategory $Category
    $paths = Get-Paths
    $workspace = Get-WorkspacePath $Language
    try { [void](Scan-Updates $Language $null $false) } catch { }

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
        foreach ($v in $requested) {
            $snap = Update-StructureLocked $Language {
                param($st)
                Apply-DefaultNumberingPerVolume $Language $st $cat
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
                        snapshotId = [string](Get-DataProperty $w[0] 'lastRenderedSnapshotId' '')
                        versionId = [string](Get-DataProperty $w[0] 'lastRenderedVersionId' '')
                        sourceHash = Normalize-FileHash ([string](Get-DataProperty $w[0] 'lastRenderedExcelHash' ''))
                        renderEnvironmentFingerprint = [string](Get-DataProperty $w[0] 'renderEnvironmentFingerprint' '')
                        renderEnvironment = (Get-DataProperty $w[0] 'renderEnvironment' $null)
                    }
                }
                Set-NoteProperty $sn 'sourceWorkbooks' @($sw)
                return $sn
            }
            if ([int]$snap.pageCount -le 0) { $skipped += $v; continue }
            $blockers = @(Get-Array $snap.blockers | Where-Object { [string]$_.code -ne 'no-pages' })
            if ($blockers.Count -gt 0) { throw [InvalidOperationException]::new(("{0}：{1}" -f (Get-VolumeLabelForMessage $v), [string]$blockers[0].message)) }
            $snapshots[$v] = $snap
            $targets += $v
        }
        if ($targets.Count -eq 0) { return [ordered]@{ built = @(); skipped = @($skipped); message = '出力対象がありません。' } }

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
                $outName = Get-OutputFileName $v ([string]$snapshots[$v].projectId) $Category
                $outPath = Join-Path ([string]$paths.outputDir) $outName
                Set-NoteProperty $t 'finalPath' $outPath
                Set-NoteProperty $t 'existed' ([bool](Test-Path -LiteralPath $outPath))
                if (Test-Path -LiteralPath $outPath) {
                    $f = $null
                    try { $f = [IO.File]::Open($outPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
                    catch { throw "出力先の最終PDFが開かれているため上書きできません: $outName" }
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
                $manifest = [ordered]@{ schemaVersion=2; language=$Language; category=$cat; volume=$v; projectId=[string]$snap.projectId; inputFingerprint=[string]$snap.fingerprint; outputPdf=$tmp; createdAt=New-NowIso; pageNumber=[ordered]@{font='Arial';fontSize=8;bottomPt=18;format='hyphenated';countHidden=$true}; pages=$snap.manifestPages }
                $manifestPath = Join-Path $workspace ("exports\manifest_{0}_{1}.json" -f $v, $cat)
                Write-JsonFile $manifestPath $manifest
                $java = Resolve-JavaExe
                $run = Invoke-NativeCapture $java @('-cp', "$composerJar;$pdfboxJar", 'ReportPdfComposer', '--manifest', $manifestPath)
                $exit = [int]$run.exitCode
                $text = [string]$run.text
                if ($exit -ne 0) { throw "PDFBox組版に失敗しました。exit=$exit`n$text" }
                if (-not (Test-Path $tmp) -or (Get-Item $tmp).Length -le 0) { throw '最終PDFを作成できませんでした。' }
                Set-NoteProperty $t 'newPdfHash' (Normalize-FileHash (New-Sha256 $tmp))
                Set-NoteProperty $t 'manifest' $manifest
            }
            Write-FinalJournal $Language $journal

            # 9. fingerprint 再確認
            foreach ($v in $targets) {
                $after = Update-StructureLocked $Language { param($st) return Get-FinalBuildInputSnapshot $st $Language $v $cat }
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
                    $ready = Get-FinalBuildReadiness $st $Language $v $cat
                    if ($ready.blockers.Count -gt 0) {
                        Set-NoteProperty $vs 'status' 'needs-rebuild'
                        Add-StaleReason $vs 'excel-updated' '元Excelが更新されたため、PDFを再作成後に最終PDFを再出力してください'
                    } else { Set-NoteProperty $vs 'status' 'built' }
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
                $archive = New-FinalArchive $Language $cat $v ([string]$t.buildId) ([string]$t.finalPath) (Get-DataProperty $t 'manifest' $null) $snapshots[$v]
                Save-LayoutSnapshot $Language $cat 'final-build' | Out-Null
                Write-HistoryEvent $Language 'final.built' ([ordered]@{ category = $cat; volume = $v; buildId = [string]$t.buildId; outputPdf = [string]$t.finalPath })
                $built += [ordered]@{ volume = $v; category = $cat; outputPdf = [string]$t.finalPath; buildId = [string]$t.buildId; archivePath = $archive; inputFingerprint = [string]$journal.beforeFingerprints[$v] }
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
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) { return @{} }
    $stamp = [IO.Directory]::GetLastWriteTimeUtc($dir).Ticks
    $cacheKey = ([IO.Path]::GetFullPath($dir)).ToLowerInvariant()
    $cached = $Script:ContentPdfSheetIndexCache[$cacheKey]
    if ($null -ne $cached -and [Int64](Get-DataProperty $cached 'stamp' -1) -eq $stamp) {
        return (Get-DataProperty $cached 'index' @{})
    }
    $root = [IO.Path]::GetFullPath($workspace)
    if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
    $index = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.pdf' -ErrorAction SilentlyContinue)) {
        $full = [IO.Path]::GetFullPath($file.FullName)
        if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { continue }
        $key = ([string]$file.BaseName).ToLowerInvariant()
        if (-not $index.ContainsKey($key)) { $index[$key] = $full }
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
    foreach ($candidate in @($SheetName, ([regex]::Replace($SheetName, '[^0-9A-Za-z]+', '-')))) {
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


function Get-SnapshotSummaries([string]$Language, [string]$WorkbookId) {
    if ([string]::IsNullOrWhiteSpace($WorkbookId)) { return @() }
    $out = @()
    foreach ($id in (Get-SnapshotIds $Language $WorkbookId)) {
        $m = Get-SnapshotManifest $Language $WorkbookId $id
        if ($null -eq $m) { continue }
        $state = Get-SnapshotSourceState $Language $WorkbookId $id
        $versions = @(Get-RenderVersionIds $Language $WorkbookId $id)
        $preferred = ''
        $visualAvailable = $false
        $contentAvailable = $false
        $reason = 'PDF作成済みの比較可能な版がありません。'
        foreach ($versionId in @($versions | Sort-Object -Descending)) {
            $availability = Get-HistoryRenderVersionAvailability $Language $WorkbookId $id ([string]$versionId)
            if ([bool]$availability.visualHashAvailable) { $visualAvailable = $true }
            if ([bool]$availability.contentPdfAvailable) { $contentAvailable = $true }
            if ([string]::IsNullOrWhiteSpace($preferred) -and [bool]$availability.ready) {
                $preferred = [string]$versionId
                $reason = ''
            } elseif ([string]::IsNullOrWhiteSpace($preferred) -and -not [string]::IsNullOrWhiteSpace([string]$availability.reason)) {
                $reason = [string]$availability.reason
            }
        }
        $out += [ordered]@{
            snapshotId = $id
            detectedAt = [string](Get-DataProperty $m 'detectedAt' '')
            captureReason = [string](Get-DataProperty $m 'captureReason' '')
            sourceHash = [string](Get-DataProperty $m 'sourceHash' '')
            sourceRetained = [bool]$state.sourceRetained
            pins = @(Get-SnapshotPins $Language $WorkbookId $id)
            renderVersionIds = @($versions)
            visualCompareReady = (-not [string]::IsNullOrWhiteSpace($preferred))
            preferredVersionId = $preferred
            visualHashAvailable = $visualAvailable
            contentPdfAvailable = $contentAvailable
            unavailableReason = $reason
        }
    }
    return @($out | Sort-Object { [string]$_.snapshotId } -Descending)
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


function Serve-HistoryRasterPage(
    $Context,
    [string]$Language,
    [string]$WorkbookId,
    [string]$SnapshotId,
    [string]$VersionId,
    [string]$SheetName,
    [int]$PageNumber
) {
    if ([string]::IsNullOrWhiteSpace($WorkbookId) -or [string]::IsNullOrWhiteSpace($SnapshotId) -or
        [string]::IsNullOrWhiteSpace($VersionId) -or [string]::IsNullOrWhiteSpace($SheetName) -or $PageNumber -lt 1) {
        throw [System.ArgumentException]::new('workbookId / snapshotId / versionId / sheetName / pageNumber が必要です。')
    }
    $safeWorkbookId = Assert-SafeStorageSegment $WorkbookId 'workbookId'
    $safeSnapshotId = Assert-SafeStorageSegment $SnapshotId 'snapshotId'
    $safeVersionId = Assert-SafeStorageSegment $VersionId 'versionId'
    # snapshotとversionの組み合わせを直接検証し、別世代のラスタを混在させない。
    $renderRecordDir = Get-RenderRecordDir $Language $safeWorkbookId $safeSnapshotId $safeVersionId
    if (-not (Test-Path -LiteralPath $renderRecordDir -PathType Container)) {
        throw [System.ArgumentException]::new('指定したレンダリング世代が見つかりません。')
    }
    $rasterDir = Get-RenderRasterSheetDir $Language $safeWorkbookId $safeSnapshotId $safeVersionId $SheetName
    $root = [IO.Path]::GetFullPath($rasterDir)
    $full = [IO.Path]::GetFullPath((Join-Path $root ('page-{0:0000}.png' -f $PageNumber)))
    if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
    if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $full -PathType Leaf)) {
        Write-JsonResponse $Context 404 ([ordered]@{ ok = $false; error = 'render-raster-not-found' })
        return
    }
    # 比較時に再画像化は行わず、PDF作成後のハッシュ解析で保存済みの120 DPI PNGをそのまま返す。
    Write-FileResponse $Context 200 $full 'image/png' $false 'private, max-age=31536000, immutable'
}

function Get-AutoStateSummary([string]$Language) {
    $settings = Get-AutoRenderSettings
    $items = @()
    try {
        $dir = Get-AutoStateDir $Language
        if (Test-Path -LiteralPath $dir) {
            foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Filter '*.json' -ErrorAction SilentlyContinue)) {
                $j = Read-JsonFile $f.FullName $null
                if ($null -eq $j) { continue }
                if ([string](Get-DataProperty $j 'state' '') -eq 'idle') { continue }
                $items += [ordered]@{
                    workbookId = [string](Get-DataProperty $j 'workbookId' $f.BaseName)
                    state = [string](Get-DataProperty $j 'state' '')
                    deferReason = [string](Get-DataProperty $j 'deferReason' '')
                    quietDeadline = [string](Get-DataProperty $j 'quietDeadline' '')
                    firstDetectedAt = [string](Get-DataProperty $j 'firstDetectedAt' '')
                    ownerPcName = [string](Get-DataProperty $j 'ownerPcName' '')
                }
            }
        }
    } catch { }
    return [ordered]@{
        enabled = [bool]$settings.enabled
        # 旧Web UIとの互換性のためプロパティ名を維持する。policy.json の値ではない。
        inputHistoryApproved = $true
        sourceRetentionApproved = [bool](Test-SourceRetentionEnabled)
        quietPeriodSeconds = [int]$settings.quietPeriodSeconds
        schedulerRunning = (Test-AutoSchedulerProcessRunning)
        historySizeMb = (Get-InputHistorySizeMb $Language)
        softCapMegabytes = [int](Get-InputHistorySettings).softCapMegabytes
        items = @($items)
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
    Write-AutoState $Language $WorkbookId $state
    return [ordered]@{ workbookId = $WorkbookId; requested = $true }
}

# =====================================================================
# V5 差分詳細・視覚比較
# =====================================================================

$Script:DiffDetailAlgorithmVersion = 15
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
        schemaVersion = 2
        status = 'unavailable'
        scope = 'history'
        baselineSnapshotId = $BaselineSnapshotId
        baselineVersionId = $BaselineVersionId
        currentSnapshotId = $CurrentSnapshotId
        currentVersionId = $CurrentVersionId
        comparedAt = New-NowIso
        method = 'stored-hash-history'
        confidence = 1.0
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
    $baseMap = @{}
    foreach ($sheet in @(Get-Array (Get-DataProperty $base 'sheets' @()))) {
        $name = [string](Get-DataProperty $sheet 'sheetName' '')
        if (-not [string]::IsNullOrWhiteSpace($name)) { $baseMap[$name] = $sheet }
    }
    $currentNames = @{}
    $changed = @(); $unchanged = @(); $unknown = @(); $added = @(); $removed = @()
    foreach ($sheet in @(Get-Array (Get-DataProperty $current 'sheets' @()))) {
        $name = [string](Get-DataProperty $sheet 'sheetName' '')
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $currentNames[$name] = $true
        if ([string](Get-DataProperty $sheet 'status' '') -ne 'ok') { $unknown += $name; continue }
        if (-not $baseMap.ContainsKey($name)) { $added += $name; continue }
        $before = $baseMap[$name]
        if ([string](Get-DataProperty $before 'status' '') -ne 'ok') { $unknown += $name; continue }
        if ((Normalize-FileHash ([string](Get-DataProperty $before 'sheetVisualHash' ''))) -eq
            (Normalize-FileHash ([string](Get-DataProperty $sheet 'sheetVisualHash' ''))) -or
            (Test-SheetVisualEquivalent $before $sheet)) { $unchanged += $name } else { $changed += $name }
    }
    foreach ($name in @($baseMap.Keys)) { if (-not $currentNames.ContainsKey([string]$name)) { $removed += [string]$name } }
    $result.status = 'complete'
    $result.changedSheets = @($changed)
    $result.unchangedSheets = @($unchanged)
    $result.unknownSheets = @($unknown)
    $result.addedSheets = @($added)
    $result.removedSheets = @($removed)
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
    if ($matches.Count -eq 0) { throw '登録済みExcelが見つかりません。' }
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
    if ([string]::IsNullOrWhiteSpace($renderedHash) -or $status -in @('new','excel-updated','render-error','rendering') -or
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
    $workspace = Get-WorkspacePath $Language
    $currentPdfDir = Get-ContentPdfVersionDir $workspace $safeWorkbookId $currentVersionId
    $baselinePdfDir = Get-ContentPdfVersionDir $workspace $safeWorkbookId $baselineVersionId
    if ($null -eq $currentHashes -or $null -eq $baselineHashes -or
        -not (Test-Path -LiteralPath $currentPdfDir -PathType Container) -or
        -not (Test-Path -LiteralPath $baselinePdfDir -PathType Container)) {
        $baseResult.message = '自動比較に使った画像ハッシュまたはcontent PDF世代が保持されていません。PDFを再作成してください。'
        return [pscustomobject]$baseResult
    }
    $baseResult.currentVisualHashes = $currentHashes
    $baseResult.baselineVisualHashes = $baselineHashes
    $baseResult.available = $true; $baseResult.status = 'available'; $baseResult.comparison = $comparison; $baseResult.comparisonPersisted = $true
    $baseResult.currentSnapshotId = $currentSnapshotId; $baseResult.currentVersionId = $currentVersionId
    $baseResult.baselineSnapshotId = $baselineSnapshotId; $baseResult.baselineVersionId = $baselineVersionId
    $baseResult.currentAt = Get-DiffSnapshotDate $Language $safeWorkbookId $currentSnapshotId ([string](Get-DataProperty $workbook 'lastRenderedAt' ''))
    $baseResult.baselineAt = Get-DiffSnapshotDate $Language $safeWorkbookId $baselineSnapshotId ''
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

    $items = @()
    $seen = @{}
    foreach ($group in @(
        [ordered]@{ kind = 'modified'; names = @(Get-Array (Get-DataProperty $comparison 'changedSheets' @())) },
        [ordered]@{ kind = 'added'; names = @(Get-Array (Get-DataProperty $comparison 'addedSheets' @())) },
        [ordered]@{ kind = 'removed'; names = @(Get-Array (Get-DataProperty $comparison 'removedSheets' @())) },
        [ordered]@{ kind = 'unknown'; names = @(Get-Array (Get-DataProperty $comparison 'unknownSheets' @())) },
        [ordered]@{ kind = 'unchanged'; names = @(Get-Array (Get-DataProperty $comparison 'unchangedSheets' @())) }
    )) {
        foreach ($rawName in @($group.names)) {
            $name = [string]$rawName
            if ([string]::IsNullOrWhiteSpace($name) -or $seen.ContainsKey($name)) { continue }
            if ([string]$group.kind -eq 'modified' -and ($addedSet.ContainsKey($name) -or $removedSet.ContainsKey($name) -or $unknownSet.ContainsKey($name))) { continue }
            $seen[$name] = $true
            $items += [pscustomobject][ordered]@{
                sheetName = $name
                sheetKey = Get-DiffSheetKey $name
                kind = [string]$group.kind
                beforePages = 0
                afterPages = 0
                pageCount = 0
                unchangedPageNumbers = @()
                regionCount = 0
                # PDFはそのまま保持し、表示中のページだけをブラウザで描画・比較する。
                status = 'ready'
                message = '表示したページをブラウザで比較します。'
                confirmed = $false
                pages = @()
            }
        }
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
        if ($baselineMap.ContainsKey([string]$item.sheetName)) {
            $beforeSheet = $baselineMap[[string]$item.sheetName]
            $item.beforePages = Get-IntDataProperty $beforeSheet 'pageCount' 0
        }
        if ($currentMap.ContainsKey([string]$item.sheetName)) {
            $afterSheet = $currentMap[[string]$item.sheetName]
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
        }
        generation = [ordered]@{ status = 'completed'; jobId = ''; percent = 100; message = '表示ページをブラウザで比較します。'; currentSheet = '' }
        sheets = @($items)
        generatedAt = ''
    }
}

function Get-DiffDetail(
    [string]$Language,
    [string]$WorkbookId,
    [string]$BaselineSnapshotId = '',
    [string]$CurrentSnapshotId = ''
) {
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
    })
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
            if ($workIndexes.Count -eq 0) { throw '指定したシートは比較対象に含まれていません。' }
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
            $kind = [string]$sheet.kind
            try {
                $beforePdf = ''; $afterPdf = ''
                if ($kind -ne 'added') { $beforePdf = Resolve-DiffContentPdfPath $language $workbookId ([string]$context.baselineSnapshotId) ([string]$context.baselineVersionId) $name }
                if ($kind -ne 'removed') { $afterPdf = Resolve-DiffContentPdfPath $language $workbookId ([string]$context.currentSnapshotId) ([string]$context.currentVersionId) $name }
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
                        Get-RenderRasterSheetDir $language $workbookId ([string]$context.baselineSnapshotId) ([string]$context.baselineVersionId) $name
                    } else { '' })
                    afterRasterDirectory = $(if ($kind -ne 'removed') {
                        Get-RenderRasterSheetDir $language $workbookId ([string]$context.currentSnapshotId) ([string]$context.currentVersionId) $name
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
            $status.message = "差分画像を作成しています: $($position + 1) / $($workIndexes.Count) シート $name"
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
            $detail.message = ('差分画像に未完了のシートがあります（' + ($parts -join '、') + '）。再試行してください。')
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
    if ($sheet.Count -eq 0) { throw '指定したシートは比較対象に含まれていません。' }
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
        $sheetName = [string](Get-DataProperty $sheet[0] 'sheetName' '')
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

if ($Port -le 0) { $Port = Get-FreePort }
$config0 = Get-AppConfig
if ([string](Get-DataProperty $config0 'lastMode' '') -ne $Mode) {
    $config0.lastMode = $Mode
    Save-AppConfig $config0
}
try { $startupPaths=Get-Paths; if($startupPaths.dataDir -and (Test-Path -LiteralPath ([string]$startupPaths.dataDir))){Ensure-Package $startupPaths -Languages @((Get-EffectiveLanguage))} } catch { Write-Warning $_.Exception.Message }

# 未完了の最終PDFトランザクションだけは、UI操作を受け付ける前に復旧する。
try { Invoke-StartupRecovery (Get-EffectiveLanguage) } catch { Write-Warning $_.Exception.Message }

$prefix = "http://127.0.0.1:$Port/"
$url = ("http://127.0.0.1:{0}/?token={1}&mode={2}" -f $Port, $Script:Token, $Mode)

# Use a small TcpListener-based HTTP server instead of HttpListener.
# This avoids URL ACL / administrator-rights issues on locked-down Windows PCs.
Start-LocalTcpServer $Port $url ([bool]$NoOpen)
