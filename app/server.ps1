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
# /api/state ãŒèª­ã‚“ã å…±æœ‰ãƒ•ã‚©ãƒ«ãƒ€ãƒ¼ä¸Šã®ãƒ¡ã‚¿ãƒ‡ãƒ¼ã‚¿ã‚’ã€ç›´å¾Œã®æ¯”è¼ƒç”»é¢ã§ã‚‚å†åˆ©ç”¨ã™ã‚‹ã€‚
# ãƒ•ã‚¡ã‚¤ãƒ«æ›´æ–°æ™‚åˆ»ãƒ»ã‚µã‚¤ã‚ºã¾ãŸã¯ç‰ˆIDãŒå¤‰ã‚ã‚Œã°åˆ¥ã‚­ãƒ¼/å†èª­è¾¼ã«ãªã‚‹ã€‚
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
# æœ€å°åŒ–ã‚³ãƒ³ã‚½ãƒ¼ãƒ«ã§èµ·å‹•ã•ã‚ŒãŸã¨ãã€ä½•ã®ã‚¦ã‚£ãƒ³ãƒ‰ã‚¦ã‹åˆ†ã‹ã‚‹ã‚ˆã†ã«ã‚¿ã‚¤ãƒˆãƒ«ã‚’ä»˜ã‘ã‚‹ã€‚
try { $host.UI.RawUI.WindowTitle = "ReportBinder ã‚µãƒ¼ãƒãƒ¼ ($Mode) - ã“ã®ã‚¦ã‚£ãƒ³ãƒ‰ã‚¦ã‚’é–‰ã˜ã‚‹ã¨çµ‚äº†ã—ã¾ã™" } catch { }
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
# Excel ã¯åŒä¸€ãƒ—ãƒ­ã‚»ã‚¹å†…ã®åŒæœŸCOMå‘¼ã³å‡ºã—ãªã®ã§ã€Word/PowerPoint ã®ã‚ˆã†ã«ãƒ¯ãƒ¼ã‚«ãƒ¼ã”ã¨
# è½ã¨ã™æ–¹æ³•ãŒä½¿ãˆãªã„ã€‚é€²æ—ãŒæ­¢ã¾ã£ã¦ã‹ã‚‰ã®çŒ¶äºˆã¨ã—ã¦é•·ã‚ã«å–ã‚‹(å¤§ããªãƒ–ãƒƒã‚¯ã¯
# 1ã‚·ãƒ¼ãƒˆã«æ•°åç§’ã‹ã‹ã‚‹)ã€‚è©³ç´°ã¯ Start-ExcelRenderWatchdog ã‚’å‚ç…§ã€‚
$Script:ExcelRenderTimeoutSeconds = 300
# 1ã‚·ãƒ¼ãƒˆã‚ãŸã‚Šã®ä¸Šä¹—ã›ã€‚è¤‡æ•°ã‚·ãƒ¼ãƒˆã‚’1å›žã§æ›¸ãå‡ºã™çµŒè·¯ã¯ã€ãã®é–“ã¾ã£ãŸãå¿ƒæ‹ã‚’
# æ‰“ã¦ãªã„ãŸã‚ã€ã‚·ãƒ¼ãƒˆæ•°ã¶ã‚“çŒ¶äºˆã‚’ä¼¸ã°ã•ãªã„ã¨æ­£å¸¸ãªå¤‰æ›ã‚’æ®ºã—ã¦ã—ã¾ã†ã€‚
$Script:ExcelPerSheetAllowanceSeconds = 60
$Script:ExcelWatchdogProcess = $null
$Script:ExcelWatchdogHeartbeatPath = ''
$Script:ExcelWatchdogKilledPath = ''
$Script:ExcelWatchdogFiredSticky = $false
$Script:ExcelWatchdogLogPaths = @()
$Script:OwnedExcelProcessId = 0
# å®Ÿè¡Œä¸­ã®ã‚¸ãƒ§ãƒ–ãŒã€Œä¸­æ­¢ãŒè¦æ±‚ã•ã‚ŒãŸã‹ã€ã‚’ç­”ãˆã‚‹ã‚¹ã‚¯ãƒªãƒ—ãƒˆãƒ–ãƒ­ãƒƒã‚¯ã€‚å¤–éƒ¨ã‚³ãƒžãƒ³ãƒ‰ã®
# å¾…ã¡ãƒ«ãƒ¼ãƒ—ãŒã“ã‚Œã‚’1ç§’ã”ã¨ã«è¦‹ã‚‹ã€‚ã‚¸ãƒ§ãƒ–å¤–ã§ã¯ $nullã€‚
$Script:NativeCancelProbe = $null
# å¤–éƒ¨ã‚³ãƒžãƒ³ãƒ‰(å‘¼ã³å…ˆã¯ã™ã¹ã¦ java)ã®ä¸Šé™ã€‚ç”¨é€”ã”ã¨ã«åˆ†ã‘ã‚‹ã€‚
# 2026-08-10: ä¸Šé™ã‚’å°Žå…¥ã—ãŸæ™‚ç‚¹ã€9ç®‡æ‰€ã™ã¹ã¦ãŒæ—¢å®šã®120ç§’ã§èµ°ã£ã¦ã„ãŸã€‚120ç§’ã¯
# Word/PowerPoint ã® COM å¤‰æ›ã«åˆã‚ã›ãŸå€¤ã§ã€java ã®å®Ÿè¨ˆç®—ã«æµç”¨ã§ãã‚‹ã‚‚ã®ã§ã¯ãªã„ã€‚
# åˆ†å‰²ã¯ã‚³ãƒ¼ãƒ‰è‡ªèº«ãŒ2000ãƒšãƒ¼ã‚¸ã®PDFã‚’è¨±å®¹ã—ã€çµ„ç‰ˆã®å‡ºåŠ›å…ˆã¯å…±æœ‰ãƒ•ã‚©ãƒ«ãƒ€ãƒ¼ã«ã‚‚ãªã‚‹ã€‚
# ã©ã¡ã‚‰ã‚‚120ç§’ã«è§¦ã‚Œå¾—ã‚‹ã®ã«ã€å‘¼ã³å‡ºã—å´ãŒæ™‚é–“åˆ‡ã‚Œã¨ç•°å¸¸çµ‚äº†ã‚’åŒºåˆ¥ã—ã¦ã„ãªã‹ã£ãŸãŸã‚ã€
# å¥å…¨ãªå…¥åŠ›ã«ã€Œç ´æã—ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€ã¨è¡¨ç¤ºã—ã¦ã„ãŸã€‚
$Script:JavaProbeTimeoutSeconds = 30
$Script:PdfSplitTimeoutSeconds = 900
$Script:FinalComposeTimeoutSeconds = 900
# è§£æžã¯PDFä½œæˆã®ã‚¯ãƒªãƒ†ã‚£ã‚«ãƒ«ãƒ‘ã‚¹å¤–(Invoke-PostRenderAnalysis)ã€‚ä¸Šé™ã‚’è¨­ã‘ãŸå…ƒã®ç†ç”±ãŒ
# ã“ã“ã® AWT åˆæœŸåŒ–ã®ç„¡é™å¾…ã¡ãªã®ã§ã€ä»–ã‚ˆã‚ŠçŸ­ãæŠ‘ãˆã‚‹ã€‚
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
# V5-P2: è¨­å®šãƒ»ãƒ‘ã‚¹ãƒ»å±¥æ­´å®¹é‡ã¯ Get-WorkspacePath çµŒç”±ã§é »ç¹ã«å‚ç…§ã•ã‚Œã‚‹ã€‚
# æ¯Žå›žãƒ‡ã‚£ã‚¹ã‚¯ã‚’èª­ã‚€(ã•ã‚‰ã« config ã¯æ›¸ã)ã¨ã€å…±æœ‰ãƒ‰ãƒ©ã‚¤ãƒ–ä¸Šã§è‡´å‘½çš„ã«é…ããªã‚‹ã€‚çŸ­æ™‚é–“ã ã‘ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã™ã‚‹ã€‚
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
    # V5-P3: ä»¥å‰ã¯ Exception.ToString() ãŒå–ã‚Œãªã„ã¨ã€ãƒ¡ãƒƒã‚»ãƒ¼ã‚¸1è¡Œã ã‘ã«ãªã‚ŠåŽŸå› è¿½è·¡ãŒã§ããªã‹ã£ãŸã€‚
    # åž‹ãƒ»HResultãƒ».NETã‚¹ã‚¿ãƒƒã‚¯ãƒ»å†…éƒ¨ä¾‹å¤–ãƒ»ç™ºç”Ÿè¡Œã‚’ã€å–ã‚ŒãŸã‚‚ã®ã‹ã‚‰å¿…ãšç©ã‚€ã€‚
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
    # CommandLineToArgvW ã®è¦å‰‡ã§1æœ¬ã®ã‚³ãƒžãƒ³ãƒ‰ãƒ©ã‚¤ãƒ³æ–‡å­—åˆ—ã«ã™ã‚‹ã€‚
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
    # V5-P3: Windows PowerShell 5.1 ã§ã¯ã€ãƒã‚¤ãƒ†ã‚£ãƒ–ã‚³ãƒžãƒ³ãƒ‰ã® stderr ã‚’ 2>&1 ã§å–ã‚Šè¾¼ã‚€ã¨
    # ErrorRecord ã¨ã—ã¦ãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³ã«æµã‚Œã€$ErrorActionPreference='Stop' ã®ä¸‹ã§ã¯
    # NativeCommandError ã®ä¾‹å¤–ã«ãªã‚‹ã€‚
    # PDFBox ã¯æ—¥æœ¬èªžãƒ•ã‚©ãƒ³ãƒˆã‚’å«ã‚€PDFã§è­¦å‘Š(Format 14 cmap table ...)ã‚’ stderr ã«å‡ºã™ãŸã‚ã€
    # è§£æžã‚„çµ„ç‰ˆãŒæˆåŠŸã—ã¦ã„ã¦ã‚‚å‘¼ã³å‡ºã—å´ãŒã€Œå¤±æ•—ã€ã¨èª¤èªã—ã¦ã„ãŸã€‚
    #
    # 2026-08-09: å¾“æ¥ã¯ `& $FilePath @args 2>&1 | ...` ã§å¾…ã£ã¦ãŠã‚Šã€å¾…ã¡æ™‚é–“ã®ä¸Šé™ãŒ
    # ç„¡ã‹ã£ãŸã€‚ã“ã®é–¢æ•°ã®å‘¼ã³å‡ºã—å…ˆã¯ã™ã¹ã¦ java ã§ã€ãã®1ã¤ PdfPageAnalyzer ã¯ PDF ã‚’
    # ãƒ©ã‚¹ã‚¿ãƒ©ã‚¤ã‚ºã™ã‚‹ãŸã‚ Windows ã§ã¯ -Djava.awt.headless=true ã‚’ä»˜ã‘ã¦ã‚‚ AWT ã®
    # ãƒ„ãƒ¼ãƒ«ã‚­ãƒƒãƒˆ(sun.awt.windows.WToolkit)ã‚’ç”Ÿæˆã™ã‚‹ã€‚ãã®åˆæœŸåŒ–ã¯ã‚¦ã‚£ãƒ³ãƒ‰ã‚¦
    # ã‚¹ãƒ†ãƒ¼ã‚·ãƒ§ãƒ³ãŒä½¿ãˆãªã„ã¨ä¸Šé™ãªã—ã§å¾…ã¡ç¶šã‘ã‚‹ãŸã‚ã€CIã®ã‚ˆã†ã«å¯¾è©±ãƒ‡ã‚¹ã‚¯ãƒˆãƒƒãƒ—ãŒ
    # ä¸å®‰å®šãªç’°å¢ƒã§ã¯ java ãŒæ°¸ä¹…ã«è¿”ã‚‰ãšã€å‘¼ã³å‡ºã—å…ƒã”ã¨å›ºã¾ã£ã¦ã„ãŸã€‚
    # ProcessStartInfo ã§ç›´æŽ¥èµ·å‹•ã—ã€ä¸Šé™ã‚’éŽãŽãŸã‚‰ãƒ—ãƒ­ã‚»ã‚¹ãƒ„ãƒªãƒ¼ã”ã¨çµ‚äº†ã•ã›ã‚‹ã€‚
    # è§£æžã¯ã€ŒPDFä½œæˆã®ã‚¯ãƒªãƒ†ã‚£ã‚«ãƒ«ãƒ‘ã‚¹ã®å¤–ã€(Invoke-PostRenderAnalysis ã‚’å‚ç…§)ãªã®ã§ã€
    # ä¾‹å¤–ã§ã¯ãªã exitCode éž0 ã¨ã—ã¦è¿”ã—ã€å‘¼ã³å‡ºã—å…ƒã®æ—¢å­˜ã®å¤±æ•—å‡¦ç†ã«è¼‰ã›ã‚‹ã€‚
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = ConvertTo-NativeArgumentString $ArgumentList
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    # æ¨™æº–å…¥åŠ›ã‚’æ¸¡ã•ãªã„ã€‚ç¶™æ‰¿ã—ãŸç«¯ã‚’å­ã‚„å­«ãŒæ¡ã‚‹ã¨ã€å‘¼ã³å‡ºã—å…ƒã®ãƒ‘ã‚¤ãƒ—ãŒé–‰ã˜ãªã„ã€‚
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
        # 1ç§’ãšã¤å¾…ã£ã¦ã€ãã®ãŸã³ã«ä¸­æ­¢è¦æ±‚ã‚’è¦‹ã‚‹ã€‚ä¸Šé™ã ã‘ã§å¾…ã¤ã¨ã€ä¸Šé™ã‚’é•·ãã—ãŸåˆ†
        # ãã®ã¾ã¾ã€Œä¸­æ­¢ã‚’æŠ¼ã—ã¦ã‚‚ä½•ã‚‚èµ·ããªã„æ™‚é–“ã€ã«ãªã‚‹(çµ„ç‰ˆã¨åˆ†å‰²ã¯15åˆ†)ã€‚
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
        $text = (("NATIVE_TIMEOUT: " + $FilePath + " ãŒ " + [string]$TimeoutSeconds + " ç§’ã§çµ‚ã‚ã‚‰ãªã„ãŸã‚ä¸­æ­¢ã—ã¾ã—ãŸã€‚") + "`n" + $text).Trim()
    }
    if ($cancelled) {
        $text = ("NATIVE_CANCELLED: ä¸­æ­¢ã®è¦æ±‚ã‚’å—ã‘ã¦å¤–éƒ¨ã‚³ãƒžãƒ³ãƒ‰ã‚’çµ‚äº†ã—ã¾ã—ãŸã€‚" + "`n" + $text).Trim()
    }
    $lines = if ($text -eq '') { @() } else { @($text -split "`n") }
    return [ordered]@{ exitCode = $exit; output = $lines; text = $text; timedOut = $timedOut; cancelled = $cancelled }
}

# Read a file timestamp from an open Windows file handle. On SMB shares this is
# more reliable than directory-enumeration metadata and matches Explorer's
# "æ›´æ–°æ—¥æ™‚" value. Fall back to System.IO on non-Windows or if the handle call fails.
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

    // èµ·å‹•ã—ãŸ Office ã‚¢ãƒ—ãƒªã®ãƒ—ãƒ­ã‚»ã‚¹IDã‚’ã€ãã®ã‚¦ã‚£ãƒ³ãƒ‰ã‚¦ãƒãƒ³ãƒ‰ãƒ«ã‹ã‚‰å¼•ãã€‚
    // åˆ©ç”¨è€…ãŒè‡ªåˆ†ã§é–‹ã„ã¦ã„ã‚‹ Excel ã‚’å·»ãæ·»ãˆã«ã—ãªã„ãŸã‚ã€ã“ã¡ã‚‰ãŒèµ·å‹•ã—ãŸ
    // 1ã¤ã ã‘ã‚’ç‰¹å®šã™ã‚‹å¿…è¦ãŒã‚ã‚‹ã€‚
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
        # æ·±ã„å…±æœ‰ãƒ•ã‚©ãƒ«ãƒ€ã§ã¯ã€æœ€çµ‚ãƒ‘ã‚¹ã¯æ‰±ãˆã¦ã‚‚GUIDä»˜ãä¸€æ™‚åã ã‘ãŒ
        # MAX_PATHã‚’è¶…ãˆã‚‹ã“ã¨ãŒã‚ã‚‹ã€‚ä½œæˆã‚‚tryå†…ã«ç½®ãã€ç›´æŽ¥æ›¸è¾¼ã¸ç¸®é€€ã™ã‚‹ã€‚
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
            throw "JSONä¿å­˜ã«å¤±æ•—ã—ã¾ã—ãŸ: $Path / atomic=$atomicError / shared=$($_.Exception.Message)"
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
    $err = [ordered]@{ error = $Message; userError = 'PDFä½œæˆãƒ—ãƒ­ã‚»ã‚¹ã‚’èµ·å‹•ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚ReportBinderã‚’ä¸€åº¦çµ‚äº†ã—ã¦ã‹ã‚‰å†å®Ÿè¡Œã—ã¦ãã ã•ã„ã€‚'; detail = ([string](Get-DataProperty $Job 'startupLog' '')) }
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
    if ($RelativePath -match '[\\/]') { throw 'æå‡ºãƒ•ã‚©ãƒ«ãƒ€ç›´ä¸‹ã®Excelã ã‘ç™»éŒ²ã§ãã¾ã™ã€‚å­ãƒ•ã‚©ãƒ«ãƒ€å†…ã®ãƒ•ã‚¡ã‚¤ãƒ«ã¯å¯¾è±¡å¤–ã§ã™ã€‚' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -notin @('.xlsx','.xlsm')) { throw 'æ‹¡å¼µå­ãŒ .xlsx ã¾ãŸã¯ .xlsm ã®Excelã ã‘ç™»éŒ²ã§ãã¾ã™ã€‚' }
    if ([IO.Path]::GetFileName($RelativePath) -like '~$*') { throw 'Excelã®ä¸€æ™‚ãƒ•ã‚¡ã‚¤ãƒ«ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚' }
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
    # æ—¢å®šå€¤ã‚’ã‚­ãƒ¼å˜ä½ã§å†å¸°çš„ã«è£œå®Œã™ã‚‹ã€‚æ—¢ã«ã‚ã‚‹å€¤ã¯å¿…ãšå„ªå…ˆã™ã‚‹(åˆ©ç”¨è€…ã®è¨­å®šã‚’å£Šã•ãªã„)ã€‚
    # Target/Defaults ã¯ PSCustomObject ã§ã‚‚ ordered hashtable ã§ã‚‚ã‚ˆã„ã€‚
    if ($null -eq $Defaults) { return $Target }
    if ($null -eq $Target) { return $Defaults }
    foreach ($name in (Get-ConfigKeyNames $Defaults)) {
        $defValue = Get-DataProperty $Defaults $name $null
        if (-not (Test-ConfigHasKey $Target $name)) {
            Set-NoteProperty $Target $name $defValue
            # V5-P2: æ—¢å®šå€¤ã‚’å®Ÿéš›ã«è£œå®Œã—ãŸã¨ãã ã‘ config.json ã‚’æ›¸ãæˆ»ã™ã€‚
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
    # è¨­å®šãƒ»ãƒ‘ã‚¹ã‚’å¤‰æ›´ã—ãŸã‚‰å¿…ãšå‘¼ã¶ã€‚
    $Script:AppConfigCache = $null
    $Script:AppConfigCacheAtUtc = [DateTime]::MinValue
    $Script:PathsCache = $null
    $Script:PathsCacheAtUtc = [DateTime]::MinValue
    # dataDir ã®åˆ‡æ›¿æ™‚ã«æ—§ãƒ¯ãƒ¼ã‚¯ã‚¹ãƒšãƒ¼ã‚¹ã®å®¹é‡ã‚’è¿”ã•ãªã„ã€‚
    $Script:HistorySizeCache = $null
    $Script:HistorySizeCacheKey = ''
    $Script:HistorySizeCacheAtUtc = [DateTime]::MinValue
}

function Get-AppConfig {
    # V5: ãƒ­ãƒ¼ã‚«ãƒ«configã‚’ãã®ã¾ã¾è¿”ã™ã¨ã€æ—¢å­˜åˆ©ç”¨è€…ã«æ–°ã—ã„ã‚­ãƒ¼(autoRender ãªã©)ãŒåæ˜ ã•ã‚Œãªã„ã€‚
    # æ—¢å®šå€¤ã‚’èª­ã¿ã€ãƒ­ãƒ¼ã‚«ãƒ«å„ªå…ˆã§ã‚­ãƒ¼å˜ä½ã«ãƒžãƒ¼ã‚¸ã—ã¦ã‹ã‚‰è¿”ã™ã€‚
    # V5-P2: ä»¥å‰ã¯ã“ã®é–¢æ•°ãŒå‘¼ã°ã‚Œã‚‹ãŸã³ã« config.json ã‚’æ›¸ãæˆ»ã—ã¦ã„ãŸã€‚
    # Get-Paths -> Get-WorkspacePath çµŒç”±ã§ã»ã¼å…¨é–¢æ•°ã‹ã‚‰å‘¼ã°ã‚Œã‚‹ãŸã‚ã€
    # æ¤œçŸ¥ç‰ˆã®ä¸€è¦§å–å¾—ãªã©ã§1ä»¶ã”ã¨ã«ãƒ•ã‚¡ã‚¤ãƒ«æ›¸ãè¾¼ã¿ãŒç™ºç”Ÿã—ã¦ã„ãŸã€‚
    # å®Ÿéš›ã«æ—¢å®šå€¤ã‚’è£œå®Œã—ãŸã¨ãã ã‘æ›¸ãã€çµæžœã¯çŸ­æ™‚é–“ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã™ã‚‹ã€‚
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
    # å±¥æ­´ãƒ»å·®åˆ†ã¯æ—¢å®šã§æœ‰åŠ¹ã€‚æ—§ policy.json ã®æ‰¿èªãƒ•ãƒ©ã‚°ã¯å‚ç…§ã—ãªã„ã€‚
    return $true
}

function Test-SourceRetentionEnabled {
    # åˆ©ç”¨è€…ãŒé¸æŠžã—ãŸæå‡ºExcelã®å±¥æ­´ã¯ã€ReportBinderå°‚ç”¨ã®ãƒ­ãƒ¼ã‚«ãƒ«ãƒ—ãƒ­ã‚¸ã‚§ã‚¯ãƒˆå†…ã«ä¿æŒã™ã‚‹ã€‚
    # å…±æœ‰æå‡ºãƒ•ã‚©ãƒ«ãƒ€ãƒ¼ã¸å±¥æ­´ãƒ»ä¸­é–“PDFã‚’æ›¸ã‹ãªã„ã€‚
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
    # ãƒ•ã‚©ãƒ«ãƒ€åã« % ã‚„ # ã‚’å«ã‚€ã¨ Uri.MakeRelativeUri / UnescapeDataString ãŒ
    # ãƒ‘ã‚¹ã‚’å£Šã™ï¼ˆ%20â†’ç©ºç™½åŒ–ã€#ä»¥é™æ¬ è½ï¼‰ã€‚é…ä¸‹ã®å ´åˆã¯å˜ç´”ãªåˆ‡ã‚Šå‡ºã—ã§æ±‚ã‚ã‚‹ã€‚
    if ($full.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) {
        return $full.Substring($base.Length)
    }
    $baseUri = New-Object System.Uri($base)
    $fullUri = New-Object System.Uri($full)
    $rel = [System.Uri]::UnescapeDataString($baseUri.MakeRelativeUri($fullUri).ToString())
    return ($rel -replace '/', [IO.Path]::DirectorySeparatorChar)
}

function Normalize-FileHash([string]$Value) {
    # V5: ãƒ•ã‚¡ã‚¤ãƒ«ãƒãƒƒã‚·ãƒ¥ã¯ New-Sha256 ã®å½¢å¼(64æ–‡å­—ãƒ»å¤§æ–‡å­—ãƒ»prefixãªã—)ã«çµ±ä¸€ã™ã‚‹ã€‚
    # éŽåŽ»ãƒ‡ãƒ¼ã‚¿ã‚„æ‰‹æ›¸ãè¨­å®šã« 'sha256:' ä»˜ãå°æ–‡å­—ãŒæ··ã–ã£ã¦ã„ã¦ã‚‚æ¯”è¼ƒãŒå£Šã‚Œãªã„ã‚ˆã†å¸åŽã™ã‚‹ã€‚
    # æ–‡å­—åˆ— fingerprint (Get-Sha256Text) ã® 'sha256:'+å°æ–‡å­— ã¨ã¯åˆ¥ç‰©ãªã®ã§æ··ãœãªã„ã“ã¨ã€‚
    $v = [string]$Value
    if ([string]::IsNullOrWhiteSpace($v)) { return '' }
    if ($v -match '^(?i)sha256:') { $v = $v.Substring(7) }
    return $v.Trim().ToUpperInvariant()
}

function New-RbId {
    # V5: ID = ã‚¿ã‚¤ãƒ ã‚¹ã‚¿ãƒ³ãƒ—(ãƒŸãƒªç§’) + '_' + GUID8ã€‚
    # ãƒãƒƒã‚·ãƒ¥ã‚„ fingerprint ã¯ ID ã«åŸ‹ã‚è¾¼ã¾ãªã„(manifest ã®æ­£å¼ãƒ•ã‚£ãƒ¼ãƒ«ãƒ‰ã¨ã—ã¦æŒã¤)ã€‚
    return ((Get-Date).ToString('yyyyMMddTHHmmss.fff') + '_' + ([Guid]::NewGuid().ToString('N').Substring(0,8)))
}

function New-RbVersionId {
    return ('v' + (New-RbId))
}

function New-UniqueDirectory([string]$Parent, [scriptblock]$IdFactory) {
    # ç”Ÿæˆå¾Œã«æ—¢å­˜ãƒ‡ã‚£ãƒ¬ã‚¯ãƒˆãƒªãŒã‚ã‚Œã°å†ç”Ÿæˆã™ã‚‹(æœ€å¤§3å›ž)ã€‚immutable ãªä¸–ä»£ãƒ•ã‚©ãƒ«ãƒ€ãŒæ··ã–ã‚‹ã®ã‚’é˜²ãã€‚
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $id = [string](& $IdFactory)
        $full = Join-Path $Parent $id
        if (-not (Test-Path -LiteralPath $full)) {
            New-Item -ItemType Directory -Path $full -Force | Out-Null
            return [ordered]@{ id = $id; path = $full }
        }
        Start-Sleep -Milliseconds 5
    }
    throw "ä¸€æ„ãªãƒ•ã‚©ãƒ«ãƒ€åã‚’ç”Ÿæˆã§ãã¾ã›ã‚“ã§ã—ãŸ: $Parent"
}

function New-Sha256([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return '' }
    # Get-FileHash ã¯ FileShare.Read ã§é–‹ããŸã‚ã€èª°ã‹ãŒExcelã§(æ›¸ãè¾¼ã¿ã‚¢ã‚¯ã‚»ã‚¹ä»˜ãã§)é–‹ã„ã¦ã„ã‚‹ã¨
    # ã€Œåˆ¥ã®ãƒ—ãƒ­ã‚»ã‚¹ã§ä½¿ç”¨ã•ã‚Œã¦ã„ã¾ã™ã€ã§å¤±æ•—ã™ã‚‹ã€‚FileShare.ReadWrite ã‚’æ˜Žç¤ºã—ã¦é–‹ã‘ã°èª­ã‚ã‚‹ã€‚
    $fs = $null
    $sha = $null
    try {
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        $sha = [Security.Cryptography.SHA256]::Create()
        $hashBytes = $sha.ComputeHash($fs)
        $hex = (-join ($hashBytes | ForEach-Object { $_.ToString('x2') })).ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($hex)) { throw "ãƒãƒƒã‚·ãƒ¥ã‚’è¨ˆç®—ã§ãã¾ã›ã‚“ã§ã—ãŸ: $Path" }
        return $hex
    } finally {
        if ($sha) { try { $sha.Dispose() } catch { } }
        if ($fs) { try { $fs.Dispose() } catch { } }
    }
}

function Copy-FileSharedRead([string]$Source, [string]$Destination) {
    # Excelã§é–‹ã‹ã‚Œã¦ã„ã‚‹(æ›¸ãè¾¼ã¿ã‚¢ã‚¯ã‚»ã‚¹ä¿æŒä¸­ã®)ãƒ•ã‚¡ã‚¤ãƒ«ã‚‚ã‚³ãƒ”ãƒ¼ã§ãã‚‹ã‚ˆã†ã€
    # èª­ã¿å–ã‚Šå´ã‚’ FileShare.ReadWrite ã§é–‹ãã€‚Copy-Item ã§ã¯åŒã˜ç†ç”±ã§å¤±æ•—ã™ã‚‹ã€‚
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
    # when PDFä½œæˆ processes many Excel files. If the file has not been touched for a few seconds,
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
    if ([string]::IsNullOrWhiteSpace($RelativePath)) { throw 'relativePath ãŒç©ºã§ã™ã€‚' }
    if ([IO.Path]::IsPathRooted($RelativePath)) { throw 'çµ¶å¯¾ãƒ‘ã‚¹ã¯å—ã‘ä»˜ã‘ã¾ã›ã‚“ã€‚' }
    if ($RelativePath -match '(^|[\\/])\.\.($|[\\/])') { throw '.. ã‚’å«ã‚€ãƒ‘ã‚¹ã¯å—ã‘ä»˜ã‘ã¾ã›ã‚“ã€‚' }
    if ($RelativePath -match '[\x00-\x1F]') { throw 'åˆ¶å¾¡æ–‡å­—ã‚’å«ã‚€ãƒ‘ã‚¹ã¯å—ã‘ä»˜ã‘ã¾ã›ã‚“ã€‚' }
    return $true
}

function Assert-SafeStorageSegment([string]$Value, [string]$Name = 'è­˜åˆ¥å­') {
    # workbookId / snapshotId / versionId are used as single directory or file-name
    # segments. Never let API input introduce separators, drive prefixes, or dot
    # traversal into the history/archive trees.
    $segment = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($segment) -or
        $segment.Length -gt 200 -or
        $segment -in @('.', '..') -or
        $segment -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        throw [System.ArgumentException]::new("$Name ãŒä¸æ­£ã§ã™ã€‚")
    }
    return $segment
}

function Join-Safe([string]$Root, [string]$RelativePath) {
    Test-RelativePath $RelativePath | Out-Null
    $full = [IO.Path]::GetFullPath((Join-Path $Root $RelativePath))
    $rootFull = [IO.Path]::GetFullPath($Root)
    if (-not $rootFull.EndsWith([IO.Path]::DirectorySeparatorChar)) { $rootFull += [IO.Path]::DirectorySeparatorChar }
    if (-not $full.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) { throw 'ç™»éŒ²æ¸ˆã¿ãƒ•ã‚©ãƒ«ãƒ€å¤–ã®ãƒ‘ã‚¹ã§ã™ã€‚' }
    return $full
}

function Get-Paths {
    # V5-P2: Get-WorkspacePath çµŒç”±ã§ã»ã¼å…¨é–¢æ•°ã‹ã‚‰å‘¼ã°ã‚Œã‚‹ã€‚å…±æœ‰ãƒ‰ãƒ©ã‚¤ãƒ–ä¸Šã®
    # common\paths.json ã‚’1å›žã®æ“ä½œã§ä½•ç™¾å›žã‚‚èª­ã¿ç›´ã•ãªã„ã‚ˆã†çŸ­æ™‚é–“ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã™ã‚‹ã€‚
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
    # åŒã˜å…±æœ‰ãƒ•ã‚©ãƒ«ãƒ€ãƒ¼ã‚’ H:\ ã¨ \\server\share ã®ä¸¡æ–¹ã§é¸ã‚“ã§ã‚‚ã€åŒã˜ãƒ­ãƒ¼ã‚«ãƒ«
    # ãƒ—ãƒ­ã‚¸ã‚§ã‚¯ãƒˆã«ãªã‚‹ã‚ˆã†ã€å–å¾—ã§ãã‚‹å ´åˆã¯ãƒžãƒƒãƒ—ãƒ‰ãƒ©ã‚¤ãƒ–ã‚’UNCã¸æ­£è¦åŒ–ã™ã‚‹ã€‚
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

    # ç§»è¡Œå®Œäº†ãƒžãƒ¼ã‚«ãƒ¼ãŒãªã„ãƒ­ãƒ¼ã‚«ãƒ«ãƒ‡ãƒ¼ã‚¿ã¯ã€æ—§å…±æœ‰ã‚³ãƒ”ãƒ¼ã®å¤±æ•—é€”ä¸­ã‹
    # æ—§å®Ÿè£…ã®æ®‹éª¸ã§ã‚ã‚‹ã€‚å‰Šé™¤ã›ãšçŸ­ã„åå‰ã§é€€é¿ã—ã€æ–°ã—ã„ç®¡ç†é ˜åŸŸã¯ç©ºã§å§‹ã‚ã‚‹ã€‚
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

    # å¤ã„ç‰ˆã¨ã®åŒæ™‚èµ·å‹•ã§ã‚‚æ—§å…±æœ‰ã‚³ãƒ”ãƒ¼ã¨æ–°ã—ã„ã‚¯ãƒªãƒ¼ãƒ³åˆæœŸåŒ–ãŒç«¶åˆã—ãªã„ã‚ˆã†ã€
    # mutexåã¯æ—¢å­˜ç‰ˆã¨åŒã˜ã¾ã¾ç¶­æŒã™ã‚‹ã€‚
    $projectMutex = $null
    $projectMutexOwned = $false
    try {
        $created = $false
        $projectMutex = New-Object System.Threading.Mutex($false, ('Local\ReportBinder.ProjectMigration.' + [string]$LocalPaths.projectKey), [ref]$created)
        try { $projectMutexOwned = $projectMutex.WaitOne(1800000, $false) }
        catch [System.Threading.AbandonedMutexException] { $projectMutexOwned = $true }
        if (-not $projectMutexOwned) { throw 'åˆ©ç”¨è€…ãƒ­ãƒ¼ã‚«ãƒ«ç®¡ç†é ˜åŸŸã®åˆæœŸåŒ–å¾…ã¡ãŒã‚¿ã‚¤ãƒ ã‚¢ã‚¦ãƒˆã—ã¾ã—ãŸã€‚' }
        if (Test-Path -LiteralPath $marker -PathType Leaf) {
            return [ordered]@{ initialized = $false; reason = 'already-local'; marker = $marker }
        }

        # å…±æœ‰å´ã® _reportbinder / å‡ºåŠ› ã¯å­˜åœ¨ç¢ºèªã‚‚ã‚³ãƒ”ãƒ¼ã‚‚è¡Œã‚ãªã„ã€‚
        # æå‡ºãƒ•ã‚©ãƒ«ãƒ€ãƒ¼ã¯Excelã®èª­è¾¼å…ƒã¨ã—ã¦ã ã‘ä½¿ç”¨ã™ã‚‹ã€‚
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
    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = 'ãƒ•ã‚©ãƒ«ãƒ€ã‚’é¸æŠžã—ã¦ãã ã•ã„' }
    if ([string]::IsNullOrWhiteSpace($InitialDir) -or -not (Test-Path -LiteralPath $InitialDir)) { $InitialDir = [Environment]::GetFolderPath('MyDocuments') }

    $helper = Join-Path $Script:AppRoot 'tools\select-folder.ps1'
    if (-not (Test-Path -LiteralPath $helper)) { throw 'ãƒ•ã‚©ãƒ«ãƒ€é¸æŠžç”¨ã®è£œåŠ©ã‚¹ã‚¯ãƒªãƒ—ãƒˆãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }

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
            throw 'ãƒ•ã‚©ãƒ«ãƒ€é¸æŠžç”»é¢ãŒå¿œç­”ã—ã¾ã›ã‚“ã€‚ç”»é¢ã§ãƒ‘ã‚¹ã‚’ç›´æŽ¥å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚'
        }
        if ($proc.ExitCode -ne 0) { throw "ãƒ•ã‚©ãƒ«ãƒ€é¸æŠžãƒ€ã‚¤ã‚¢ãƒ­ã‚°ã‚’é–‹ã‘ã¾ã›ã‚“ã§ã—ãŸã€‚ExitCode=$($proc.ExitCode)" }
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
    if ([string]::IsNullOrWhiteSpace($resolvedDataDir)) { throw 'ç®¡ç†ãƒ‡ãƒ¼ã‚¿ãƒ•ã‚©ãƒ«ãƒ€ãŒæœªè¨­å®šã§ã™ã€‚æå‡ºãƒ•ã‚©ãƒ«ãƒ€ã‚’é¸ã‚“ã§ãã ã•ã„ã€‚' }
    # ä¸€å¼ã‚’å”¯ä¸€ã®ç®¡ç†å˜ä½ã¨ã—ã€ã™ã¹ã¦ã®ãƒ‘ãƒƒã‚¯ã‚’åŒã˜ãƒ¯ãƒ¼ã‚¯ã‚¹ãƒšãƒ¼ã‚¹ã«ä¿å­˜ã™ã‚‹ã€‚
    # è¨€èªžã¯ä¿å­˜é ˜åŸŸã‚’åˆ†å‰²ã™ã‚‹è¨­å®šã§ã¯ãªãã€ä¸€å¼ã‚„ã²ãªå½¢ã®å±žæ€§ã¨ã—ã¦æ‰±ã†ã€‚
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
    if ($RelativePath -match '[\/]') { throw 'æå‡ºãƒ•ã‚©ãƒ«ãƒ€ç›´ä¸‹ã®WordåŽŸç¨¿ã ã‘ç™»éŒ²ã§ãã¾ã™ã€‚å­ãƒ•ã‚©ãƒ«ãƒ€å†…ã®ãƒ•ã‚¡ã‚¤ãƒ«ã¯å¯¾è±¡å¤–ã§ã™ã€‚' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -ne '.docx') { throw 'æ‹¡å¼µå­ãŒ .docx ã®WordåŽŸç¨¿ã ã‘ç™»éŒ²ã§ãã¾ã™ã€‚' }
    if ([IO.Path]::GetFileName($RelativePath) -like '~$*') { throw 'Wordã®ä¸€æ™‚ãƒ•ã‚¡ã‚¤ãƒ«ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚' }
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
    if ($RelativePath -match '[\/]') { throw 'æå‡ºãƒ•ã‚©ãƒ«ãƒ€ç›´ä¸‹ã®PowerPointåŽŸç¨¿ã ã‘ç™»éŒ²ã§ãã¾ã™ã€‚å­ãƒ•ã‚©ãƒ«ãƒ€å†…ã®ãƒ•ã‚¡ã‚¤ãƒ«ã¯å¯¾è±¡å¤–ã§ã™ã€‚' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -ne '.pptx') { throw 'æ‹¡å¼µå­ãŒ .pptx ã®PowerPointåŽŸç¨¿ã ã‘ç™»éŒ²ã§ãã¾ã™ã€‚' }
    if ([IO.Path]::GetFileName($RelativePath) -like '~$*') { throw 'PowerPointã®ä¸€æ™‚ãƒ•ã‚¡ã‚¤ãƒ«ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚' }
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
    if ([string]::IsNullOrWhiteSpace($name)) { throw [ArgumentException]::new('ä¸€å¼åã‚’å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
    if ($name.Length -gt 120) { throw [ArgumentException]::new('ä¸€å¼åã¯120æ–‡å­—ä»¥å†…ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
    if ($name -match '[\x00-\x1f\x7f]') { throw [ArgumentException]::new('ä¸€å¼åã«åˆ¶å¾¡æ–‡å­—ã¯ä½¿ç”¨ã§ãã¾ã›ã‚“ã€‚') }
    return $name
}

function Test-PackArchived($Pack) {
    return (-not [string]::IsNullOrWhiteSpace([string](Get-DataProperty $Pack 'archivedAt' '')))
}

function Get-PackRecord($Structure, [string]$PackId) {
    $id = ([string]$PackId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw [ArgumentException]::new('packId ãŒå¿…è¦ã§ã™ã€‚') }
    $matches = @(Get-Array (Get-DataProperty $Structure 'packs' @()) | Where-Object { [string](Get-DataProperty $_ 'packId' '') -eq $id } | Select-Object -First 1)
    if ($matches.Count -eq 0) { throw [ArgumentException]::new('æŒ‡å®šã•ã‚ŒãŸä¸€å¼ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚') }
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
    else { throw [ArgumentException]::new('æŒ‡å®šã•ã‚ŒãŸä¸€å¼ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚') }
    if ((Test-PackArchived $pack) -and -not $AllowArchived) { throw [InvalidOperationException]::new('ã‚¢ãƒ¼ã‚«ã‚¤ãƒ–æ¸ˆã¿ã®ä¸€å¼ã¯æ“ä½œã§ãã¾ã›ã‚“ã€‚å¾©å…ƒã—ã¦ã‹ã‚‰ã‚„ã‚Šç›´ã—ã¦ãã ã•ã„ã€‚') }
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
    if (@(Get-PackTargetIds $Language $scope.pack) -notcontains $targetId) { throw [ArgumentException]::new('ã“ã®ä¸€å¼ã«å­˜åœ¨ã—ãªã„å‡ºåŠ›å…ˆã§ã™ã€‚') }
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
            throw [ArgumentException]::new('åŒã˜åå‰ã®ä¸€å¼ãŒæ—¢ã«ã‚ã‚Šã¾ã™ã€‚åˆ¥ã®åå‰ã‚’å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚')
        }
    }
    return $name
}

function Get-UniquePackDisplayName($Structure, [string]$BaseName, [string]$Language) {
    $root = ConvertTo-PackDisplayName $BaseName
    $suffix = $(if ($Language -eq 'en') { 'Copy' } else { 'ã‚³ãƒ”ãƒ¼' })
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
    throw 'è¤‡è£½ã—ãŸä¸€å¼ã«ä¸€æ„ãªåå‰ã‚’ä»˜ã‘ã‚‰ã‚Œã¾ã›ã‚“ã§ã—ãŸã€‚'
}

function New-CustomPackId($Structure) {
    $existing = @{}; foreach ($pack in @(Get-Array (Get-DataProperty $Structure 'packs' @()))) { $existing[[string](Get-DataProperty $pack 'packId' '')] = $true }
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $id = 'pack_' + ([Guid]::NewGuid().ToString('N').Substring(0,16))
        if (-not $existing.ContainsKey($id)) { return $id }
    }
    throw 'ä¸€æ„ãªä¸€å¼IDã‚’ç”Ÿæˆã§ãã¾ã›ã‚“ã§ã—ãŸã€‚'
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
    if ([string]::IsNullOrWhiteSpace([string]$settings.documentTitle)) { throw [ArgumentException]::new('è³‡æ–™ã‚¿ã‚¤ãƒˆãƒ«ã‚’å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
    [void](Resolve-OutputFileNamePattern ([string]$settings.outputFileNamePattern) 'PROJECT' ([string](Get-DataProperty $Pack 'displayName' 'ReportBinder')) (Get-TargetDisplayName $Language ("$Language-main")))
    return $settings
}

function ConvertTo-NormalizedPageRange($Value) {
    if ($null -eq $Value) { return $null }
    $start = 0; $end = 0
    if ($Value -is [string]) {
        $text = ([string]$Value).Trim()
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        if ($text -notmatch '^(?<start>\d+)(?:\s*-\s*(?<end>\d+))?$') { throw [ArgumentException]::new('ãƒšãƒ¼ã‚¸ç¯„å›²ã¯ã€Œ2ã€ã¾ãŸã¯ã€Œ2-5ã€ã®å½¢å¼ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
        $start = [int]$matches.start; $end = if ($matches.end) { [int]$matches.end } else { $start }
    } else {
        $start = Get-IntDataProperty $Value 'start' 0
        $end = Get-IntDataProperty $Value 'end' $start
    }
    if ($start -lt 1 -or $end -lt $start) { throw [ArgumentException]::new('ãƒšãƒ¼ã‚¸ç¯„å›²ã¯1ä»¥ä¸Šã§ã€é–‹å§‹ãƒšãƒ¼ã‚¸ãŒçµ‚äº†ãƒšãƒ¼ã‚¸ã‚’è¶…ãˆãªã„ã‚ˆã†ã«ã—ã¦ãã ã•ã„ã€‚') }
    return [ordered]@{ start = $start; end = $end }
}

function Get-TargetDisplayName([string]$Language, [string]$Volume) {
    if ($Volume -match 'appendix$') { return $(if ($Language -eq 'en') { 'Appendix' } else { 'è£œè¶³' }) }
    return $(if ($Language -eq 'en') { 'Main' } else { 'æœ¬ä½“' })
}

function Resolve-OutputFileNamePattern([string]$Pattern, [string]$ProjectId, [string]$PackName, [string]$TargetName) {
    $value = ([string]$Pattern).Trim()
    if ([string]::IsNullOrWhiteSpace($value)) { $value = '{projectId}_{targetName}.pdf' }
    $unknown = [regex]::Matches($value, '\{[^{}]+\}') | ForEach-Object { $_.Value } | Where-Object { $_ -notin @('{projectId}','{packName}','{targetName}','{yyyyMMdd}') } | Select-Object -Unique
    if (@($unknown).Count -gt 0) { throw [ArgumentException]::new("æœªå¯¾å¿œã®ãƒ•ã‚¡ã‚¤ãƒ«åãƒ—ãƒ¬ãƒ¼ã‚¹ãƒ›ãƒ«ãƒ€ãƒ¼ã§ã™: $(@($unknown) -join ', ')") }
    $resolved = $value.Replace('{projectId}', $ProjectId).Replace('{packName}', $PackName).Replace('{targetName}', $TargetName).Replace('{yyyyMMdd}', (Get-Date -Format 'yyyyMMdd'))
    if ($resolved.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $resolved -match '[\\/]') { throw [ArgumentException]::new('å‡ºåŠ›ãƒ•ã‚¡ã‚¤ãƒ«åã«ä½¿ç”¨ã§ããªã„æ–‡å­—ãŒå«ã¾ã‚Œã¦ã„ã¾ã™ã€‚') }
    if (-not $resolved.EndsWith('.pdf', [StringComparison]::OrdinalIgnoreCase)) { $resolved += '.pdf' }
    return $resolved
}

function Update-PackSettings([string]$Language, [string]$PackId, $Patch) {
    $id = ([string]$PackId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw [ArgumentException]::new('packId ãŒå¿…è¦ã§ã™ã€‚') }
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
        Mark-VolumeNeedsRebuild $structure $Language $id @(Get-PackVolumeList $Language $pack $false) 'document-settings' 'è³‡æ–™ã®ä»•ä¸Šã’è¨­å®šã‚’å¤‰æ›´ã—ã¾ã—ãŸ'
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
    if ([string]::IsNullOrWhiteSpace($displayName)) { throw [ArgumentException]::new('ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆåã‚’å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
    if ($displayName.Length -gt 120) { throw [ArgumentException]::new('ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆåã¯120æ–‡å­—ä»¥å†…ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
    $description = Get-LocalizedTemplateText (Get-DataProperty $Template 'description' '') $Language ''
    if ($description.Length -gt 500) { throw [ArgumentException]::new('ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆã®èª¬æ˜Žã¯500æ–‡å­—ä»¥å†…ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
    $sourceTypes = @(
        Get-Array (Get-DataProperty $Template 'acceptedSourceTypes' @('excel','word','pdf','powerpoint')) |
            ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
            Where-Object { $_ -in @('excel','word','pdf','powerpoint') } |
            Select-Object -Unique
    )
    if ($sourceTypes.Count -eq 0) { throw [ArgumentException]::new('ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆã§æ‰±ã†åŽŸç¨¿å½¢å¼ã‚’1ã¤ä»¥ä¸Šé¸ã‚“ã§ãã ã•ã„ã€‚') }
    $targetInput = @(Get-Array (Get-DataProperty $Template 'targets' @()))
    if ($targetInput.Count -eq 0) {
        $targetInput = @(
            [pscustomobject][ordered]@{ targetId='main'; displayName=$(if ($Language -eq 'en') { 'Main' } else { 'æœ¬ä½“' }); required=$true },
            [pscustomobject][ordered]@{ targetId='appendix'; displayName=$(if ($Language -eq 'en') { 'Appendix' } else { 'è£œè¶³' }); required=$false }
        )
    }
    if ($targetInput.Count -gt 12) { throw [ArgumentException]::new('å‡ºåŠ›å…ˆã¯12ä»¶ä»¥å†…ã§è¨­å®šã—ã¦ãã ã•ã„ã€‚') }
    $targets = @()
    $targetIds = @{}
    foreach ($targetInputItem in $targetInput) {
        $targetId = Assert-SafeStorageSegment (([string](Get-DataProperty $targetInputItem 'targetId' '')).Trim().ToLowerInvariant()) 'targetId'
        if ($targetId -in @('none','unassigned')) { throw [ArgumentException]::new('none ã¨ unassigned ã¯å‡ºåŠ›å…ˆIDã«ä½¿ç”¨ã§ãã¾ã›ã‚“ã€‚') }
        if ($targetId.Length -gt 48) { throw [ArgumentException]::new('å‡ºåŠ›å…ˆIDã¯48æ–‡å­—ä»¥å†…ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
        if ($targetIds.ContainsKey($targetId)) { throw [ArgumentException]::new('å‡ºåŠ›å…ˆIDãŒé‡è¤‡ã—ã¦ã„ã¾ã™ã€‚') }
        $targetIds[$targetId] = $true
        $fallback = $(if ($targetId -eq 'main') { $(if ($Language -eq 'en') { 'Main' } else { 'æœ¬ä½“' }) } elseif ($targetId -eq 'appendix') { $(if ($Language -eq 'en') { 'Appendix' } else { 'è£œè¶³' }) } else { $targetId })
        $label = Get-LocalizedTemplateText (Get-DataProperty $targetInputItem 'displayName' '') $Language $fallback
        if ([string]::IsNullOrWhiteSpace($label) -or $label.Length -gt 60) { throw [ArgumentException]::new('å‡ºåŠ›å…ˆåã¯1ï½ž60æ–‡å­—ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
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
        if ([string]::IsNullOrWhiteSpace($requirementName)) { throw [ArgumentException]::new('å¿…è¦åŽŸç¨¿åã‚’å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
        if ($requirementName.Length -gt 120) { throw [ArgumentException]::new('å¿…è¦åŽŸç¨¿åã¯120æ–‡å­—ä»¥å†…ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
        $requirementId = ([string](Get-DataProperty $requirementInput 'requirementId' '')).Trim()
        if ([string]::IsNullOrWhiteSpace($requirementId)) { $requirementId = 'requirement_' + (Get-Sha256Text "$requirementIndex|$requirementName").Substring(7,16) }
        $requirementId = Assert-SafeStorageSegment $requirementId 'requirementId'
        if ($requirementIds.ContainsKey($requirementId)) { throw [ArgumentException]::new('å¿…è¦åŽŸç¨¿IDãŒé‡è¤‡ã—ã¦ã„ã¾ã™ã€‚') }
        $requirementIds[$requirementId] = $true
        $requirementTypes = @(
            Get-Array (Get-DataProperty $requirementInput 'acceptedSourceTypes' $sourceTypes) |
                ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
                Where-Object { $_ -in $sourceTypes } |
                Select-Object -Unique
        )
        if ($requirementTypes.Count -eq 0) { throw [ArgumentException]::new("å¿…è¦åŽŸç¨¿ã€Œ$requirementNameã€ã§æ‰±ã†åŽŸç¨¿å½¢å¼ã‚’1ã¤ä»¥ä¸Šé¸ã‚“ã§ãã ã•ã„ã€‚") }
        $requirementOwner = ([string](Get-DataProperty $requirementInput 'ownerDepartment' '')).Trim()
        if ($requirementOwner.Length -gt 120) { throw [ArgumentException]::new('å¿…è¦åŽŸç¨¿ã®æ‹…å½“éƒ¨ç½²ã¯120æ–‡å­—ä»¥å†…ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
        $requirementTarget = ([string](Get-DataProperty $requirementInput 'defaultTargetId' 'unassigned')).Trim().ToLowerInvariant()
        if ($requirementTarget -ne 'unassigned' -and -not $targetIds.ContainsKey($requirementTarget)) { throw [ArgumentException]::new('å¿…è¦åŽŸç¨¿ã®æ—¢å®šå‡ºåŠ›å…ˆãŒä¸æ­£ã§ã™ã€‚') }
        $dueDate = ([string](Get-DataProperty $requirementInput 'dueDate' '')).Trim()
        if (-not [string]::IsNullOrWhiteSpace($dueDate) -and $dueDate -notmatch '^\d{4}-\d{2}-\d{2}$') { throw [ArgumentException]::new('å¿…è¦åŽŸç¨¿ã®æœŸé™ã¯YYYY-MM-DDå½¢å¼ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
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
    if ($newDestination -ne 'unassigned' -and -not $targetIds.ContainsKey($newDestination)) { throw [ArgumentException]::new('æ–°è¦åŽŸç¨¿ã®æ—¢å®šå…ˆãŒä¸æ­£ã§ã™ã€‚') }
    $outputInput = Get-DataProperty $Template 'output' ([ordered]@{})
    $filePattern = ([string](Get-DataProperty $outputInput 'fileNamePattern' '{packName}_{targetName}_{yyyyMMdd}.pdf')).Trim()
    if ([string]::IsNullOrWhiteSpace($filePattern) -or $filePattern.Length -gt 180) { throw [ArgumentException]::new('å‡ºåŠ›ãƒ•ã‚¡ã‚¤ãƒ«åãƒ‘ã‚¿ãƒ¼ãƒ³ã¯1ï½ž180æ–‡å­—ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚') }
    [void](Resolve-OutputFileNamePattern $filePattern 'PROJECT' 'PACK' 'TARGET')
    $id = ([string]$TemplateId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { $id = 'template_' + ([Guid]::NewGuid().ToString('N').Substring(0,16)) }
    $id = Assert-SafeStorageSegment $id 'templateId'
    if ($id -like 'builtin-*') { throw [ArgumentException]::new('builtin-ã§å§‹ã¾ã‚‹ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆIDã¯åˆ©ç”¨è€…å®šç¾©ã«ä½¿ç”¨ã§ãã¾ã›ã‚“ã€‚') }
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
        } catch { Write-Warning ('åˆ©ç”¨è€…å®šç¾©ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆã‚’èª­ã¿è¾¼ã‚ã¾ã›ã‚“: ' + $file.Name + ' / ' + $_.Exception.Message) }
    }
    return @($result)
}

function Save-PackTemplate([string]$Language, $Request, [string]$TemplateId = '') {
    $existing = $null
    if (-not [string]::IsNullOrWhiteSpace($TemplateId)) {
        $matches = @(Get-UserPackTemplates $Language | Where-Object { [string]$_.templateId -eq $TemplateId } | Select-Object -First 1)
        if ($matches.Count -eq 0) { throw [ArgumentException]::new('ç·¨é›†ã™ã‚‹ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚') }
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
    if ([string]::IsNullOrWhiteSpace($dir)) { throw [ArgumentException]::new('å…ˆã«æå‡ºãƒ•ã‚©ãƒ«ãƒ€ã¨ç®¡ç†ãƒ‡ãƒ¼ã‚¿ãƒ•ã‚©ãƒ«ãƒ€ã‚’è¨­å®šã—ã¦ãã ã•ã„ã€‚') }
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $createdAt = if ($null -ne $existing) { Get-DataProperty $existing 'createdAt' (New-NowIso) } else { New-NowIso }
    Set-NoteProperty $template 'createdAt' $createdAt
    Set-NoteProperty $template 'updatedAt' (New-NowIso)
    Write-JsonFile (Join-Path $dir (([string]$template.templateId) + '.json')) $template
    return $template
}

function Remove-PackTemplate([string]$Language, [string]$TemplateId) {
    $id = Assert-SafeStorageSegment $TemplateId 'templateId'
    if ($id -like 'builtin-*') { throw [ArgumentException]::new('çµ„ã¿è¾¼ã¿ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆã¯å‰Šé™¤ã§ãã¾ã›ã‚“ã€‚') }
    $structure = Get-Structure $Language
    $packsUsingTemplate = @(
        Get-Array (Get-DataProperty $structure 'packs' @()) |
            Where-Object { [string](Get-DataProperty $_ 'templateId' '') -eq $id }
    )
    if ($packsUsingTemplate.Count -gt 0) {
        throw [ArgumentException]::new('ã“ã®ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆã‚’ä½¿ç”¨ã—ã¦ã„ã‚‹ä¸€å¼ãŒã‚ã‚Šã¾ã™ã€‚å…ˆã«å¯¾è±¡ãƒ‘ãƒƒã‚¯ã‚’å¤‰æ›´ã¾ãŸã¯ã‚¢ãƒ¼ã‚«ã‚¤ãƒ–ã—ã¦ãã ã•ã„ã€‚')
    }
    $path = Join-Path (Get-PackTemplateDir $Language) ($id + '.json')
    if (-not (Test-Path -LiteralPath $path)) { throw [ArgumentException]::new('å‰Šé™¤ã™ã‚‹ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚') }
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
    if ($currentTargets -ne $latestTargets) { $changes += 'å‡ºåŠ›å…ˆã®åå‰ãƒ»å¿…é ˆè¨­å®š' }
    $currentRequirements = ConvertTo-Json (Get-DataProperty $current 'sourceRequirements' @()) -Depth 10 -Compress
    $latestRequirements = ConvertTo-Json (Get-DataProperty $latest 'sourceRequirements' @()) -Depth 10 -Compress
    if ($currentRequirements -ne $latestRequirements) { $changes += 'å¿…è¦åŽŸç¨¿ãƒªã‚¹ãƒˆ' }
    $currentRules = ConvertTo-Json (Get-DataProperty $current 'rules' ([ordered]@{})) -Depth 10 -Compress
    $latestRules = ConvertTo-Json (Get-DataProperty $latest 'rules' ([ordered]@{})) -Depth 10 -Compress
    if ($currentRules -ne $latestRules) { $changes += 'æ–°è¦ãƒšãƒ¼ã‚¸é…ç½®ãƒ»å‡ºåŠ›æ¡ä»¶' }
    $currentOutput = ConvertTo-Json (Get-DataProperty $current 'output' ([ordered]@{})) -Depth 10 -Compress
    $latestOutput = ConvertTo-Json (Get-DataProperty $latest 'output' ([ordered]@{})) -Depth 10 -Compress
    if ($currentOutput -ne $latestOutput) { $changes += 'å‡ºåŠ›æ—¢å®šå€¤' }
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
                if ($null -ne $oldState) { Set-NoteProperty $oldState 'status' 'needs-rebuild'; Add-StaleReason $oldState 'template-target-removed' 'ã²ãªå½¢ã‹ã‚‰å‡ºåŠ›å…ˆãŒå‰Šé™¤ã•ã‚Œã¾ã—ãŸ' }
            }
        }
        Set-NoteProperty $pack 'templateConfig' (New-PackTemplateSnapshot $Language $latest)
        Set-NoteProperty $pack 'templateVersion' $latestVersion
        Set-NoteProperty $pack 'updatedAt' (New-NowIso)
        Add-PackOutputStates $structure $pack $latest
        Mark-VolumeNeedsRebuild $structure $Language ([string]$pack.packId) @(Get-PackVolumeList $Language $pack $false) 'template-updated' 'ä¸€å¼ã®ã²ãªå½¢ã‚’æ›´æ–°ã—ã¾ã—ãŸ'
        return [pscustomobject][ordered]@{ packId=[string]$pack.packId; templateId=[string]$pack.templateId; fromVersion=$currentVersion; toVersion=$latestVersion; removedTargetIds=@($removedTargetIds); unassignedPageCount=$unassignedPageCount; templateConfig=(New-PackTemplateSnapshot $Language $latest) }
    }
}

function Get-PackTemplateCatalog([string]$Language) {
    if ($Language -notin @('ja','en')) { $Language = 'ja' }
    $genericTargets = @(
        [pscustomobject][ordered]@{ targetId = 'main'; displayName = $(if ($Language -eq 'en') { 'Main' } else { 'æœ¬ä½“' }); required = $true },
        [pscustomobject][ordered]@{ targetId = 'appendix'; displayName = $(if ($Language -eq 'en') { 'Appendix' } else { 'è£œè¶³' }); required = $false }
    )
    $generic = [pscustomobject][ordered]@{
        templateId = 'builtin-generic-department-pack'
        templateVersion = 1
        packId = ''
        displayName = $(if ($Language -eq 'en') { 'Department document pack' } else { 'éƒ¨é–€ä¸€å¼' })
        description = $(if ($Language -eq 'en') { 'A general-purpose pack for Excel, Word, PowerPoint, and PDF source documents' } else { 'Excelãƒ»Wordãƒ»PowerPointãƒ»PDFã‚’ã¾ã¨ã‚ã‚‹æ±Žç”¨ä¸€å¼' })
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
    if (@(Get-PackTargetIds $Language $Pack) -notcontains $id) { throw [ArgumentException]::new('ã“ã®ä¸€å¼ã«å­˜åœ¨ã—ãªã„å‡ºåŠ›å…ˆã§ã™ã€‚') }
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
                [pscustomobject][ordered]@{targetId='main';displayName=$(if($Language -eq 'en'){'Main'}else{'æœ¬ä½“'});required=$false},
                [pscustomobject][ordered]@{targetId='appendix';displayName=$(if($Language -eq 'en'){'Appendix'}else{'è£œè¶³'});required=$false}
            )
            rules=[pscustomobject][ordered]@{newItemDestination='unassigned';retainManualOrder=$true;blockBuildWhenRequiredSourceIsStale=$false;blockBuildWhenRequiredSourceFailed=$false}
            output=[pscustomobject][ordered]@{fileNamePattern='{packName}_{targetName}_{yyyyMMdd}.pdf';pageNumbering='continuous';bookmarks='from-items'}
            workflowAvailable=$false; builtIn=$true
        }
    }
    $matches = @(Get-PackTemplateCatalog $Language | Where-Object { [string](Get-DataProperty $_ 'templateId' '') -eq $id } | Select-Object -First 1)
    if ($matches.Count -eq 0) { throw [ArgumentException]::new('æŒ‡å®šã•ã‚ŒãŸä¸€å¼ãƒ†ãƒ³ãƒ—ãƒ¬ãƒ¼ãƒˆãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚') }
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
            throw [ArgumentException]::new('çµ„ã¿è¾¼ã¿ä¸€å¼ã¯ã‚¢ãƒ¼ã‚«ã‚¤ãƒ–ã§ãã¾ã›ã‚“ã€‚')
        }
        if (-not $Archived) {
            [void](Assert-PackDisplayNameAvailable $structure ([string](Get-DataProperty $pack 'displayName' '')) ([string](Get-DataProperty $pack 'packId' '')))
        }
        Set-NoteProperty $pack 'archivedAt' $(if ($Archived) { New-NowIso } else { $null })
        Set-NoteProperty $pack 'updatedAt' (New-NowIso)
        return $pack
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
        if ([string]::IsNullOrWhiteSpace([string]$Paths.$key)) { throw "$key ãŒæœªè¨­å®šã§ã™ã€‚" }
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
        # V5-P2: ã‚»ãƒƒãƒˆã‚¢ãƒƒãƒ—å¤‰æ›´ç›´å¾Œã«å¤ã„ã‚­ãƒ£ãƒƒã‚·ãƒ¥ã‚’è¿”ã•ãªã„ã€‚
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
        throw 'ä½œæ¥­ãƒ‡ãƒ¼ã‚¿ã‚’æº–å‚™ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚æå‡ºãƒ•ã‚©ãƒ«ãƒ€ã‚’é¸ã³ç›´ã—ã¦ãã ã•ã„ã€‚åŽŸç¨¿ãƒ•ã‚¡ã‚¤ãƒ«ã¯å¤‰æ›´ã—ã¦ã„ã¾ã›ã‚“ã€‚'
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
        $source = [pscustomobject][ordered]@{
            sourceId = $sourceId; packId = $savedPackId
            relativePath = [string](Get-DataProperty $wb 'relativePath' (Get-DataProperty $wb 'fileName' ''))
            sourceType = $sourceType; adapterId = [string](Get-DataProperty $wb 'adapterId' $defaultAdapterId)
            displayName = [string](Get-DataProperty $wb 'displayName' (Get-DataProperty $wb 'fileName' $sourceId))
            ownerDepartment = [string](Get-DataProperty $wb 'ownerDepartment' '')
            required = [bool](Get-DataProperty $wb 'required' $true)
            defaultTargetId = [string](Get-DataProperty $wb 'defaultTargetId' 'unassigned')
            requirementId = [string](Get-DataProperty $wb 'requirementId' '')
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
            $issues += [pscustomobject][ordered]@{ code = 'category-unresolved'; sourceId = [string](Get-DataProperty $wb 'workbookId' ''); relativePath = [string](Get-DataProperty $wb 'relativePath' ''); message = 'æ—¢å­˜ã‚«ãƒ†ã‚´ãƒªã‚’åˆ¤å®šã§ãã¾ã›ã‚“ã§ã—ãŸã€‚ECMã¨ã—ã¦äº’æ›ç§»è¡Œã—ãŸå¾Œã€ä¸€å¼ã‚’ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($Workspace)) {
        foreach ($page in @(Get-Array $Structure.pages)) {
            $relativePdf = [string](Get-DataProperty $page 'contentPdf' ''); if ([string]::IsNullOrWhiteSpace($relativePdf)) { continue }
            try {
                $root = [IO.Path]::GetFullPath($Workspace); if (-not $root.EndsWith([IO.Path]::DirectorySeparatorChar)) { $root += [IO.Path]::DirectorySeparatorChar }
                $full = [IO.Path]::GetFullPath((Join-Path $Workspace $relativePdf))
                if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $full)) { throw 'missing' }
            } catch { $issues += [pscustomobject][ordered]@{ code = 'content-pdf-missing'; itemId = Resolve-PageId $page; relativePath = $relativePdf; message = 'ç§»è¡Œæ™‚ã«å¤‰æ›PDFã‚’ç¢ºèªã§ãã¾ã›ã‚“ã§ã—ãŸã€‚å†åº¦PDFã‚’ä½œæˆã—ã¦ãã ã•ã„ã€‚' } }
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
    if ($RelativePath -match '[\\/]') { throw 'æå‡ºãƒ•ã‚©ãƒ«ãƒ€ç›´ä¸‹ã®PDFã ã‘ç™»éŒ²ã§ãã¾ã™ã€‚å­ãƒ•ã‚©ãƒ«ãƒ€å†…ã®ãƒ•ã‚¡ã‚¤ãƒ«ã¯å¯¾è±¡å¤–ã§ã™ã€‚' }
    $ext = [IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    if ($ext -ne '.pdf') { throw 'æ‹¡å¼µå­ãŒ .pdf ã®PDFã ã‘ç™»éŒ²ã§ãã¾ã™ã€‚' }
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
        if (-not (Test-StructureDocument $structure $Language)) { throw 'structure.json ã®å†…å®¹ãŒä¸å®Œå…¨ã§ã™ã€‚' }
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
            throw "ç™»éŒ²æƒ…å ±ã‚’å®‰å…¨ã«èª­ã¿è¾¼ã‚ã¾ã›ã‚“ã§ã—ãŸã€‚structure.json ã¯ä¸Šæ›¸ãã—ã¦ã„ã¾ã›ã‚“ã€‚ç®¡ç†ãƒ•ã‚©ãƒ«ãƒ€ã®æŽ¥ç¶šã‚’ç¢ºèªã—ã¦å†èµ·å‹•ã—ã¦ãã ã•ã„ã€‚è©³ç´°: $primaryError"
        }
    }
    return Normalize-StructureCollections $structure
}

function Write-StructureUnlocked([string]$Language, $Structure, [string]$DataDir = '') {
    $Structure = Normalize-StructureCollections $Structure
    if ([int](Get-DataProperty $Structure 'schemaVersion' 0) -eq 3) { Sync-StructureV3FromLegacy $Structure $Language | Out-Null }
    if (-not (Test-StructureDocument $Structure $Language)) { throw 'ä¸å®Œå…¨ãªç™»éŒ²æƒ…å ±ã®ä¿å­˜ã‚’æ‹’å¦ã—ã¾ã—ãŸã€‚' }
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
        throw 'ç™»éŒ²æƒ…å ±ã‚’ä¿å­˜å¾Œã«æ¤œè¨¼ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚ç›´å‰ã®ãƒãƒƒã‚¯ã‚¢ãƒƒãƒ—ã‚’ä¿æŒã—ã¦ã„ã¾ã™ã€‚'
    }
}

function Get-Structure([string]$Language) {
    # Read-only API path. Repair and schema migration are performed only by Initialize-Or-MigrateStructure.
    # stateå–å¾—ç›´å¾Œã®æ¯”è¼ƒç”»é¢ã§åŒã˜å…±æœ‰JSONã‚’å†èª­è¾¼ã—ãªã„ã€‚å¤–éƒ¨PCã®æ›´æ–°ã¯
    # LastWriteTimeUtc/Lengthã®å¤‰åŒ–ã§æ¤œå‡ºã—ã€æ›¸è¾¼çµŒè·¯ã§ã¯æ˜Žç¤ºçš„ã«ç ´æ£„ã™ã‚‹ã€‚
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

# ãƒšãƒ¼ã‚¸æ§‹æˆã®æ¥½è¦³ãƒ­ãƒƒã‚¯ã€‚ä¸¦ã¹æ›¿ãˆAPIã¯ç”»é¢å…¨ä½“ã®ä¸¦ã³ã‚’çµ¶å¯¾å€¤ã§é€ã‚‹ãŸã‚ã€
# ã“ã‚ŒãŒç„¡ã„ã¨å¾Œã‹ã‚‰ä¿å­˜ã—ãŸå´ãŒç›¸æ‰‹ã®å¤‰æ›´ã‚’ç„¡è¨€ã§å·»ãæˆ»ã™ã€‚
#
# åˆ¤å®šææ–™ã¯ã€Œé…ç½®ãã®ã‚‚ã®ã®æŒ‡ç´‹ã€ã«ã™ã‚‹ã€‚structureå…¨ä½“ã®ç‰ˆç•ªå·ã ã¨ã€5åˆ†ã”ã¨ã®
# æ›´æ–°ã‚¹ã‚­ãƒ£ãƒ³ã‚„PDFä½œæˆã®ã‚ˆã†ã«é…ç½®ã‚’å¤‰ãˆãªã„æ›¸ãè¾¼ã¿ã§ã‚‚å¢—ãˆã¦ã—ã¾ã„ã€
# å®Ÿéš›ã«ã¯è¡çªã—ã¦ã„ãªã„æ“ä½œã‚’æ‹’å¦ã—ã¦ã—ã¾ã†ã€‚
function Get-PageLayoutFingerprint($Structure, [string]$PackId) {
    # /api/state ãŒãƒ‘ãƒƒã‚¯ã”ã¨ã«æ¯Žå›žå‘¼ã¶ã€‚Sort-Object ã¨ãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³ã¯
    # PowerShell 5.1 ã§ã¯å›ºå®šã‚³ã‚¹ãƒˆãŒå¤§ããï¼ˆ20ãƒšãƒ¼ã‚¸ã§ã‚‚åæ•°msï¼‰ã€
    # çŠ¶æ…‹å–å¾—ã®ãŸã³ã«ç©ã¿ä¸ŠãŒã‚‹ã€‚ç´ ã®ãƒ«ãƒ¼ãƒ—ã¨é…åˆ—ã‚½ãƒ¼ãƒˆã§çµ„ã¿ç«‹ã¦ã‚‹ã€‚
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
    # å…ˆé ­ãŒ pageId ãªã®ã§ã€åˆæˆæ–‡å­—åˆ—ã®åºæ•°ã‚½ãƒ¼ãƒˆã¯ pageId é †ã¨åŒã˜ä¸¦ã³ã‚’ä¸Žãˆã‚‹ã€‚
    $sorted = $parts.ToArray()
    [Array]::Sort($sorted, [StringComparer]::Ordinal)
    $bytes = [Text.Encoding]::UTF8.GetBytes([string]::Join('|', $sorted))
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function New-StructureConflictError([string]$Expected, [string]$Actual) {
    $conflict = [System.InvalidOperationException]::new('ä»–ã®ç”»é¢ã§ãƒšãƒ¼ã‚¸æ§‹æˆãŒå¤‰æ›´ã•ã‚Œã¾ã—ãŸã€‚æœ€æ–°ã®çŠ¶æ…‹ã‚’èª­ã¿è¾¼ã¿ç›´ã—ã¦ã‹ã‚‰æ“ä½œã—ã¦ãã ã•ã„ã€‚')
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

# $LayoutScope ã‚’æ¸¡ã—ãŸå‘¼ã³å‡ºã—ã ã‘ãŒã€ãƒšãƒ¼ã‚¸é…ç½®ã®è¡çªåˆ¤å®šã¨æŒ‡ç´‹ã®è¿”å´ã‚’è¡Œã†ã€‚
# ãƒ‘ãƒƒã‚¯IDã®è§£æ±ºã«ã¯ structure ãŒè¦ã‚‹ãŸã‚ã€ãƒ­ãƒƒã‚¯ã‚’å–ã£ãŸä¸­ã§è§£æ±ºã™ã‚‹ã€‚
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
        # åŒä¸€ãƒ•ã‚¡ã‚¤ãƒ«ã‚·ã‚¹ãƒ†ãƒ æ™‚åˆ»å†…ã®é€£ç¶šæ›´æ–°ã§ã‚‚å¤ã„èª­å–çµæžœã‚’è¿”ã•ãªã„ã€‚
        $Script:StructureReadCache.Clear()
        # ä¿å­˜å¾Œã®æŒ‡ç´‹ã‚’è¿”ã™ã€‚ç”»é¢ã¯ã“ã‚Œã‚’æ¬¡å›žã® baseLayout ã¨ã—ã¦ä½¿ã†ã€‚
        if ($layoutAware -and $result -is [System.Collections.IDictionary]) {
            $result['layoutFingerprint'] = Get-PageLayoutFingerprint $structure $layoutPackId
        }
        return $result
    }
}

function Save-Structure([string]$Language, $Structure) {
    throw 'Save-Structureã®ç›´æŽ¥å‘¼å‡ºã—ã¯ç¦æ­¢ã•ã‚Œã¦ã„ã¾ã™ã€‚Update-StructureLockedã‚’ä½¿ç”¨ã—ã¦ãã ã•ã„ã€‚'
}

function Get-EffectiveLanguage {
    # æ—¢å­˜ã®å†…éƒ¨volume IDã«ã ã‘ä½¿ç”¨ã™ã‚‹äº’æ›ã‚­ãƒ¼ã€‚åˆ©ç”¨è€…å‘ã‘ã®ç®¡ç†å˜ä½ã§ã¯ãªã„ã€‚
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
        throw [System.ArgumentException]::new('categoryã«ã¯ ecm / bod / dmm ã®ã„ãšã‚Œã‹ã‚’æŒ‡å®šã—ã¦ãã ã•ã„ã€‚')
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
        'ja-main' { return "${namedProjectId}_J_æœ¬ä½“.pdf" }
        'ja-appendix' { return "${namedProjectId}_J_è£œè¶³.pdf" }
        'en-main' { return "${namedProjectId}_E_Main.pdf" }
        'en-appendix' { return "${namedProjectId}_E_Appendix.pdf" }
        default { throw "æœªçŸ¥ã®æˆæžœç‰©: $Volume" }
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
    if (-not (Test-Path -LiteralPath $PdfPath)) { throw "PDFåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $PdfPath" }
    if (-not (Test-PdfMagic $PdfPath)) { throw 'PDFã®ãƒ•ã‚¡ã‚¤ãƒ«å½¢å¼ã‚’ç¢ºèªã§ãã¾ã›ã‚“ã€‚æ‹¡å¼µå­ã ã‘ãŒPDFã«ãªã£ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
    $tool = Get-PdfBatchToolInfo
    if ($null -eq $tool) { throw 'PDFåŽŸç¨¿ã®å–ã‚Šè¾¼ã¿ã«å¿…è¦ãªPDFBoxã¾ãŸã¯Java RuntimeãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }
    if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
    $prefix = Join-Path $OutputDir 'page'
    $pdfboxJar = Join-Path $Script:AppRoot 'lib\pdfbox\pdfbox-app.jar'
    $run = Invoke-NativeCapture ([string]$tool.javaExe) @('-jar', $pdfboxJar, 'PDFSplit', '-split', '1', '-outputPrefix', $prefix, $PdfPath) $Script:PdfSplitTimeoutSeconds
    $message = [string]$run.text
    # æ™‚é–“åˆ‡ã‚Œã‚’ç•°å¸¸çµ‚äº†ã¨æ··ãœãªã„ã€‚å¥å…¨ãªå¤§ãã„PDFã«ã€Œç ´æã€ã¨è¨€ã£ã¦ã—ã¾ã†ã¨ã€
    # åˆ©ç”¨è€…ã¯æ‰“ã¤æ‰‹ãŒç„¡ããªã‚‹ã€‚
    if ($run.timedOut) {
        throw ('ã“ã®PDFã®èª­ã¿è¾¼ã¿ãŒ{0}åˆ†ä»¥å†…ã«çµ‚ã‚ã‚Šã¾ã›ã‚“ã§ã—ãŸã€‚ãƒšãƒ¼ã‚¸æ•°ã®å¤šã„PDFã¯åˆ†å‰²ã—ã¦ã‹ã‚‰ç™»éŒ²ã—ã¦ãã ã•ã„ã€‚' -f [int]($Script:PdfSplitTimeoutSeconds / 60))
    }
    if ([int]$run.exitCode -ne 0) {
        if ($message -match 'password|encrypted|decrypt|InvalidPassword') { throw 'ãƒ‘ã‚¹ãƒ¯ãƒ¼ãƒ‰ä¿è­·ã•ã‚ŒãŸPDFã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚ä¿è­·ã‚’è§£é™¤ã—ãŸPDFã‚’ä½¿ç”¨ã—ã¦ãã ã•ã„ã€‚' }
        throw ('PDFã‚’å®‰å…¨ã«èª­ã¿è¾¼ã‚ã¾ã›ã‚“ã§ã—ãŸã€‚ç ´æã—ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚' + $(if ($message) { " è©³ç´°: $message" } else { '' }))
    }
    $files = @(Get-ChildItem -LiteralPath $OutputDir -File -Filter 'page*.pdf' -ErrorAction SilentlyContinue | Sort-Object {
        $m = [regex]::Match($_.BaseName, '(\d+)$'); if ($m.Success) { [int]$m.Groups[1].Value } else { [int]::MaxValue }
    }, Name)
    if ($files.Count -eq 0) { throw 'PDFã«å–ã‚Šè¾¼ã¿å¯èƒ½ãªãƒšãƒ¼ã‚¸ãŒã‚ã‚Šã¾ã›ã‚“ã€‚' }
    if ($files.Count -gt 2000) { throw '2000ãƒšãƒ¼ã‚¸ã‚’è¶…ãˆã‚‹PDFã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚åˆ†å‰²ã—ã¦ã‹ã‚‰ç™»éŒ²ã—ã¦ãã ã•ã„ã€‚' }
    $pages = @(); $number = 0
    foreach ($file in $files) {
        $number++
        if ($file.Length -le 0 -or -not (Test-PdfMagic $file.FullName)) { throw "PDFã® $number ãƒšãƒ¼ã‚¸ç›®ã‚’å–ã‚Šè¾¼ã‚ã¾ã›ã‚“ã§ã—ãŸã€‚" }
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
    if (-not (Test-Path -LiteralPath $DocxPath -PathType Leaf)) { throw "WordåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $DocxPath" }
    $item = Get-Item -LiteralPath $DocxPath
    if ($item.Length -le 0) { throw 'WordåŽŸç¨¿ãŒç©ºã§ã™ã€‚' }
    if ($item.Length -gt 536870912) { throw '512MBã‚’è¶…ãˆã‚‹WordåŽŸç¨¿ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚åˆ†å‰²ã—ã¦ã‹ã‚‰ç™»éŒ²ã—ã¦ãã ã•ã„ã€‚' }
    $stream = $null; $zip = $null
    try {
        $stream = [IO.File]::Open($DocxPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $header = New-Object byte[] 8
        $read = $stream.Read($header, 0, $header.Length)
        $stream.Dispose(); $stream = $null
        if ($read -ge 8 -and $header[0] -eq 0xD0 -and $header[1] -eq 0xCF -and $header[2] -eq 0x11 -and $header[3] -eq 0xE0) {
            throw 'ãƒ‘ã‚¹ãƒ¯ãƒ¼ãƒ‰ä¿è­·ã¾ãŸã¯æš—å·åŒ–ã•ã‚ŒãŸWordåŽŸç¨¿ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚ä¿è­·ã‚’è§£é™¤ã—ãŸDOCXã‚’ä½¿ç”¨ã—ã¦ãã ã•ã„ã€‚'
        }
        if ($read -lt 4 -or $header[0] -ne 0x50 -or $header[1] -ne 0x4B) { throw 'WordåŽŸç¨¿ã®ãƒ•ã‚¡ã‚¤ãƒ«å½¢å¼ã‚’ç¢ºèªã§ãã¾ã›ã‚“ã€‚æ‹¡å¼µå­ã ã‘ãŒDOCXã«ãªã£ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
        $zip = [IO.Compression.ZipFile]::OpenRead($DocxPath)
        $names = @($zip.Entries | ForEach-Object { ([string]$_.FullName -replace '\\','/').ToLowerInvariant() })
        if ($names -notcontains '[content_types].xml' -or $names -notcontains 'word/document.xml') { throw 'WordåŽŸç¨¿ã®å†…éƒ¨æ§‹é€ ãŒå£Šã‚Œã¦ã„ã¾ã™ã€‚Wordã§é–‹ã„ã¦åˆ¥åä¿å­˜ã—ã¦ã‹ã‚‰å†ç™»éŒ²ã—ã¦ãã ã•ã„ã€‚' }
        if (@($names | Where-Object { $_ -eq 'word/vbaproject.bin' }).Count -gt 0) { throw 'ãƒžã‚¯ãƒ­ã‚’å«ã‚€WordåŽŸç¨¿ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚.docxå½¢å¼ã§ä¿å­˜ã—ã¦ãã ã•ã„ã€‚' }
        $title = [IO.Path]::GetFileNameWithoutExtension($item.Name)
        return [pscustomobject][ordered]@{
            pageCount = 0
            units = @([pscustomobject][ordered]@{ unitKind = 'document'; sourceKey = 'document:root'; title = $title; sourceIndex = 1 })
            warnings = @()
        }
    } catch {
        if ($_.Exception.Message -match 'WordåŽŸç¨¿|DOCX|ãƒ‘ã‚¹ãƒ¯ãƒ¼ãƒ‰|ãƒžã‚¯ãƒ­') { throw }
        throw ('WordåŽŸç¨¿ã‚’å®‰å…¨ã«èª­ã¿è¾¼ã‚ã¾ã›ã‚“ã§ã—ãŸã€‚ç ´æã—ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚ è©³ç´°: ' + $_.Exception.Message)
    } finally {
        if ($zip) { $zip.Dispose() }
        if ($stream) { $stream.Dispose() }
    }
}

function New-WorkbookObject([string]$RelativePath, [string]$Language, [string]$Category = '') {
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $RelativePath" }
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

    # Mark of the Web (Zone.Identifier) ãŒä»˜ã„ãŸãƒ•ã‚¡ã‚¤ãƒ«ã¯ä¿è­·ãƒ“ãƒ¥ãƒ¼å¯¾è±¡ã¨ãªã‚Šã€
    # COMã® Workbooks.Open ãŒã€ŒWorkbooks ã‚¯ãƒ©ã‚¹ã® Open ãƒ—ãƒ­ãƒ‘ãƒ†ã‚£ã‚’å–å¾—ã§ãã¾ã›ã‚“ã€ã§å¤±æ•—ã™ã‚‹ã€‚
    # äº‹å‰ã«Zone.Identifierã‚’é™¤åŽ»ã™ã‚‹ï¼ˆå†…å®¹ãƒ»æ›´æ–°æ—¥æ™‚ã¯å¤‰ã‚ã‚‰ãªã„ï¼‰ã€‚
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
    # æœ€çµ‚ãƒ•ã‚©ãƒ¼ãƒ«ãƒãƒƒã‚¯: ä¿è­·ãƒ“ãƒ¥ãƒ¼çµŒç”±ã§é–‹ã„ã¦ç·¨é›†ãƒ¢ãƒ¼ãƒ‰ã¸æ˜‡æ ¼ã•ã›ã‚‹ã€‚
    # Unblock-FileãŒåŠ¹ã‹ãªã„ç’°å¢ƒï¼ˆã‚°ãƒ«ãƒ¼ãƒ—ãƒãƒªã‚·ãƒ¼ã§ä¿è­·ãƒ“ãƒ¥ãƒ¼å¼·åˆ¶ç­‰ï¼‰å‘ã‘ã€‚
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
    throw [System.ArgumentException]::new("æœªå¯¾å¿œã®åŽŸç¨¿å½¢å¼ã§ã™: $SourceType")
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
    throw [System.ArgumentException]::new("æ‹¡å¼µå­ã«å¯¾å¿œã™ã‚‹åŽŸç¨¿ã‚¢ãƒ€ãƒ—ã‚¿ãƒ¼ãŒã‚ã‚Šã¾ã›ã‚“: $extension")
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
    if (-not (Test-Path -LiteralPath $fullPath)) { throw "åŽŸç¨¿ãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $($validated.relativePath)" }
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
    throw [System.ArgumentException]::new("æ¤œæŸ»ã«å¯¾å¿œã—ã¦ã„ãªã„åŽŸç¨¿å½¢å¼ã§ã™: $($validated.sourceType)")
}

function Get-RegisteredSourceAdapterContext([string]$Language, [string]$SourceId) {
    if ([string]::IsNullOrWhiteSpace($SourceId)) { throw [System.ArgumentException]::new('sourceId ãŒå¿…è¦ã§ã™ã€‚') }
    $structure = Get-Structure $Language
    $source = @(Get-Array $structure.sources | Where-Object { [string]$_.sourceId -eq $SourceId } | Select-Object -First 1)
    if ($source.Count -gt 0) {
        $sourceType = Resolve-SourceTypeFromPath ([string]$source[0].relativePath) ([string]$source[0].sourceType)
        return [pscustomobject][ordered]@{ source = $source[0]; sourceId = $SourceId; sourceType = $sourceType; adapter = Get-SourceAdapterDescriptor $sourceType }
    }
    $workbook = @(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $SourceId } | Select-Object -First 1)
    if ($workbook.Count -eq 0) { throw "ç™»éŒ²æ¸ˆã¿åŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $SourceId" }
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
    throw [System.ArgumentException]::new("ç™»éŒ²ã«å¯¾å¿œã—ã¦ã„ãªã„åŽŸç¨¿å½¢å¼ã§ã™: $($validated.sourceType)")
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
    if (-not (Test-Path -LiteralPath $full)) { throw "æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $RelativePath" }
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
    if (-not (Test-Path -LiteralPath $full)) { throw "æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $RelativePath" }
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

function Render-Source([string]$Language, [string]$SourceId, $SharedHost = $null, [bool]$KeepHostOpen = $false, [scriptblock]$ProgressCallback = $null,
                       [string]$SourceOverridePath = '', [string]$SourceSnapshotId = '', [string]$ExpectedSourceHash = '') {
    $context = Get-RegisteredSourceAdapterContext $Language $SourceId
    if ([string]$context.sourceType -eq 'excel') {
        if ($KeepHostOpen -and $null -eq $SharedHost) {
            throw 'Excelã‚’èµ·å‹•ã§ããªã„ãŸã‚ã€ã“ã®åŽŸç¨¿ã‚’PDFåŒ–ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚'
        }
        $result = Render-Workbook $Language $SourceId $SharedHost $KeepHostOpen $ProgressCallback $SourceOverridePath $SourceSnapshotId $ExpectedSourceHash
        Set-NoteProperty $result 'sourceId' $SourceId
        Set-NoteProperty $result 'sourceType' 'excel'
        Set-NoteProperty $result 'adapterId' ([string]$context.adapter.adapterId)
        return $result
    }
    if ([string]$context.sourceType -eq 'pdf') {
        $result = Render-PdfSource $Language $SourceId $ProgressCallback $SourceOverridePath $SourceSnapshotId $ExpectedSourceHash
        # A batch job collects this value after every source and performs analysis only
        # after the shared rendering lock has been released. Direct calls do the same here.
        if (-not $KeepHostOpen) {
            $pending = $Script:PendingAnalysis
            $Script:PendingAnalysis = $null
            if ($null -ne $pending) {
                try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered) }
                catch { Write-Warning ('PDFåŽŸç¨¿ã®ç”»åƒè§£æžã«å¤±æ•—ã—ã¾ã—ãŸ: ' + $_.Exception.Message) }
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
                try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered) }
                catch { Write-Warning ('WordåŽŸç¨¿ã®ç”»åƒè§£æžã«å¤±æ•—ã—ã¾ã—ãŸ: ' + $_.Exception.Message) }
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
                try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered) }
                catch { Write-Warning ('PowerPointåŽŸç¨¿ã®ç”»åƒè§£æžã«å¤±æ•—ã—ã¾ã—ãŸ: ' + $_.Exception.Message) }
            }
        }
        return $result
    }
    throw [System.ArgumentException]::new("å¤‰æ›ã«å¯¾å¿œã—ã¦ã„ãªã„åŽŸç¨¿å½¢å¼ã§ã™: $($context.sourceType)")
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
    throw [System.ArgumentException]::new("å·®åˆ†è¦ç´„ã«å¯¾å¿œã—ã¦ã„ãªã„åŽŸç¨¿å½¢å¼ã§ã™: $($context.sourceType)")
}

# åŽŸç¨¿ãƒ•ã‚¡ã‚¤ãƒ«ã®åå‰å¤‰æ›´ãƒ»ç§»å‹•ã«å¯¾å¿œã™ã‚‹ã€‚æ—¥æœ¬ã®äº‹å‹™ã§ã¯ç‰ˆã‚’ãƒ•ã‚¡ã‚¤ãƒ«åã§ç®¡ç†ã™ã‚‹
# (ã€œ_v2.xlsx ãªã©)ãŸã‚ã€è¤‡æ•°å›žä½¿ã†åˆ©ç”¨è€…ã¯ã„ãšã‚Œå¿…ãšå½“ãŸã‚‹ã€‚
#
# workbookId ã¯æ®ãˆç½®ãã€‚ã“ã‚Œã¯ content-pdf\<workbookId>\ ã‚„ input-history\<workbookId>\ã€
# locks\ ã®ãƒ‡ã‚£ãƒ¬ã‚¯ãƒˆãƒªåãã®ã‚‚ã®ã§ã€å¤‰ãˆã‚‹ã¨å¤‰æ›PDFã¨å±¥æ­´ãŒå­¤å…ã«ãªã‚‹ã€‚ID ã‹ã‚‰
# ãƒ•ã‚¡ã‚¤ãƒ«åã‚’é€†ç®—ã—ã¦ã„ã‚‹ç®‡æ‰€ã¯ç„¡ã„(New-Slug ã®å‘¼ã³å‡ºã—ã¯ç™»éŒ²æ™‚ã®4ç®‡æ‰€ã ã‘)ã®ã§ã€
# å‚ç…§å…ˆã ã‘å·®ã—æ›¿ãˆã‚Œã°ã€ãƒšãƒ¼ã‚¸ã®ä¸¦ã³é †ã‚‚å‡ºåŠ›å…ˆã®æŒ¯ã‚Šåˆ†ã‘ã‚‚ä¿ãŸã‚Œã‚‹ã€‚
# ç™»éŒ²è§£é™¤â†’å†ç™»éŒ²ã¯ãƒšãƒ¼ã‚¸ã‚’ä¸¸ã”ã¨æ¶ˆã™ãŸã‚ã€ã“ã‚ŒãŒå”¯ä¸€ã®ç„¡å®³ãªç›´ã—æ–¹ã«ãªã‚‹ã€‚
function Get-RelinkCandidates([string]$Language, [string]$SourceId) {
    $structure = Get-Structure $Language
    $workbook = @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $SourceId } | Select-Object -First 1)
    if ($workbook.Count -eq 0) { throw 'æŒ‡å®šã•ã‚ŒãŸåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }
    $w = $workbook[0]
    $sourceType = ([string](Get-DataProperty $w 'sourceType' 'excel')).ToLowerInvariant()
    $packId = [string](Get-WorkbookPackId $w)
    # åŒã˜ä¸€å¼å†…ã§æ—¢ã«ä½¿ã‚ã‚Œã¦ã„ã‚‹ãƒ•ã‚¡ã‚¤ãƒ«ã¯å€™è£œã«ã—ãªã„(é‡è¤‡å‚ç…§ã«ãªã‚‹)ã€‚
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
        # ä¸­èº«ãŒåŒã˜ã‹ã©ã†ã‹ã¯ã€ã¾ãšå¤§ãã•ã§çµžã£ã¦ã‹ã‚‰ãƒãƒƒã‚·ãƒ¥ã™ã‚‹ã€‚å…±æœ‰ãƒ•ã‚©ãƒ«ãƒ€ãƒ¼ä¸Šã§
        # å…¨å€™è£œã‚’ãƒãƒƒã‚·ãƒ¥ã™ã‚‹ã¨å¾…ãŸã•ã‚Œã‚‹ãŸã‚(Scan-Updates ã¨åŒã˜è€ƒãˆæ–¹)ã€‚
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
    if ([string]::IsNullOrWhiteSpace($id)) { throw 'ä»˜ã‘æ›¿ãˆã‚‹åŽŸç¨¿ãŒæŒ‡å®šã•ã‚Œã¦ã„ã¾ã›ã‚“ã€‚' }
    $rel = ([string]$RelativePath).Trim()
    if ([string]::IsNullOrWhiteSpace($rel)) { throw 'ä»˜ã‘æ›¿ãˆå…ˆã®ãƒ•ã‚¡ã‚¤ãƒ«ãŒæŒ‡å®šã•ã‚Œã¦ã„ã¾ã›ã‚“ã€‚' }
    $paths = Get-Paths
    return Update-StructureLocked $Language {
        param($structure)
        $workbook = @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $id } | Select-Object -First 1)
        if ($workbook.Count -eq 0) { throw 'æŒ‡å®šã•ã‚ŒãŸåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }
        $w = $workbook[0]
        if ([string](Get-DataProperty $w 'status' '') -ne 'missing') {
            throw 'ã“ã®åŽŸç¨¿ã®ãƒ•ã‚¡ã‚¤ãƒ«ã¯è¦‹ã¤ã‹ã£ã¦ã„ã¾ã™ã€‚ä»˜ã‘æ›¿ãˆã¯ã€ãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚‰ãªã„åŽŸç¨¿ã«ã ã‘è¡Œãˆã¾ã™ã€‚'
        }
        $sourceType = ([string](Get-DataProperty $w 'sourceType' 'excel')).ToLowerInvariant()
        # å½¢å¼ã‚’ã¾ãŸãä»˜ã‘æ›¿ãˆã¯è¨±ã•ãªã„ã€‚IDæŽ¥é ­è¾žãƒ»unitKindãƒ»adapterIdãƒ»
        # renderProfileVersion ã®æ„å‘³ãŒåŒæ™‚ã«å£Šã‚Œã‚‹ã€‚
        $checked = Test-SourceCandidate ([pscustomobject]@{ relativePath = $rel; sourceType = '' })
        if ([string]$checked.sourceType -ne $sourceType) {
            throw ('åŒã˜ç¨®é¡žã®åŽŸç¨¿ã«ã ã‘ä»˜ã‘æ›¿ãˆã‚‰ã‚Œã¾ã™ï¼ˆä»Šã¯ {0}ã€é¸ã‚“ã ãƒ•ã‚¡ã‚¤ãƒ«ã¯ {1}ï¼‰ã€‚' -f $sourceType, [string]$checked.sourceType)
        }
        $packId = [string](Get-WorkbookPackId $w)
        foreach ($other in @(Get-Array (Get-DataProperty $structure 'workbooks' @()))) {
            if ([string](Get-DataProperty $other 'workbookId' '') -eq $id) { continue }
            if ([string](Get-WorkbookPackId $other) -ne $packId) { continue }
            if (([string](Get-DataProperty $other 'relativePath' '')).Replace('\','/').ToLowerInvariant() -eq $rel.Replace('\','/').ToLowerInvariant()) {
                throw 'ãã®ãƒ•ã‚¡ã‚¤ãƒ«ã¯ã€ã“ã®ä¸€å¼ã®åˆ¥ã®åŽŸç¨¿ã¨ã—ã¦ç™»éŒ²æ¸ˆã¿ã§ã™ã€‚'
            }
        }
        $full = Join-Safe ([string]$paths.submissionDir) $rel
        if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw 'ä»˜ã‘æ›¿ãˆå…ˆã®ãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }
        $item = Get-Item -LiteralPath $full
        # å–ã‚Šæ¶ˆã›ã‚‹å ´æ‰€ã‚’ä½œã£ã¦ã‹ã‚‰è§¦ã‚‹ï¼ˆã‚·ãƒ¼ãƒˆæ§‹æˆã®å¤‰åŒ–ã¨åŒã˜æ‰±ã„ï¼‰ã€‚
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
        # æ—§ãƒ•ã‚¡ã‚¤ãƒ«ã®æ¤œçŸ¥ç‰ˆã‚’æŒ‡ã—ãŸã¾ã¾ã«ã—ãªã„ã€‚æ¬¡ã®èµ°æŸ»ãŒæ–°ã—ã„ç‰ˆã‚’ä½œã‚‹ã€‚
        Set-NoteProperty $w 'currentSnapshotId' ''
        Set-NoteProperty $w 'lastError' ''
        $sameContent = ($hash -and $renderedHash -and $hash -eq $renderedHash)
        if ($sameContent) {
            # åå‰ãŒå¤‰ã‚ã£ãŸã ã‘ã€‚å¤‰æ›PDFã¯ãã®ã¾ã¾ä½¿ãˆã‚‹ã®ã§ã€ä½œã‚Šç›´ã—ã¯è¦ã‚‰ãªã„ã€‚
            Set-NoteProperty $w 'status' 'rendered-unchecked'
        } else {
            Set-NoteProperty $w 'status' (Get-SourceUpdatedStatus $w)
            # ä¸­èº«ãŒé•ã†ã®ã«å¤ã„å¤‰æ›PDFãŒæå‡ºç”¨PDFã«è¼‰ã‚‰ãªã„ã‚ˆã†ã€ãƒšãƒ¼ã‚¸ã‚’å¤ã„å°ã«ã™ã‚‹ã€‚
            foreach ($page in @(Get-Array (Get-DataProperty $structure 'pages' @()))) {
                if ([string](Get-DataProperty $page 'workbookId' '') -ne $id) { continue }
                if ([string](Get-DataProperty $page 'contentPdf' '')) { Set-NoteProperty $page 'status' 'stale' }
            }
            $affected = @(Get-Array (Get-DataProperty $structure 'pages' @()) |
                Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $id } |
                ForEach-Object { [string](Get-DataProperty $_ 'volume' '') } |
                Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
            Mark-VolumeNeedsRebuild $structure $Language $packId $affected 'source-updated' 'åŽŸç¨¿ã®å‚ç…§å…ˆã‚’ä»˜ã‘æ›¿ãˆã¾ã—ãŸ'
        }
        # è‡ªå‹•å‡¦ç†ã®å®‰å®šå¾…ã¡ã¯æ—§ãƒ•ã‚¡ã‚¤ãƒ«åŸºæº–ãªã®ã§æ¨ã¦ã‚‹ã€‚
        try { Remove-Item -LiteralPath (Join-Path (Get-WorkspacePath $Language) ("state\auto-render\{0}.json" -f $id)) -Force -ErrorAction SilentlyContinue } catch { }
        Write-HistoryEvent $Language 'source.relinked' ([ordered]@{ sourceId = $id; packId = $packId; relativePath = $rel; sameContent = $sameContent })
        return [ordered]@{ sourceId = $id; relativePath = $rel; fileName = $item.Name; sameContent = $sameContent; status = [string](Get-DataProperty $w 'status' '') }
    }
}

function Update-SourceMetadata([string]$Language, [string]$SourceId, $Patch) {
    $id = ([string]$SourceId).Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { throw 'åŽŸç¨¿IDãŒæŒ‡å®šã•ã‚Œã¦ã„ã¾ã›ã‚“ã€‚' }
    $hasOwner = Test-ConfigHasKey $Patch 'ownerDepartment'
    $hasRequired = Test-ConfigHasKey $Patch 'required'
    $hasDefaultTarget = Test-ConfigHasKey $Patch 'defaultTargetId'
    $hasRequirement = Test-ConfigHasKey $Patch 'requirementId'
    if (-not ($hasOwner -or $hasRequired -or $hasDefaultTarget -or $hasRequirement)) { throw 'æ›´æ–°ã™ã‚‹åŽŸç¨¿è¨­å®šãŒæŒ‡å®šã•ã‚Œã¦ã„ã¾ã›ã‚“ã€‚' }
    $owner = if ($hasOwner) { ([string](Get-DataProperty $Patch 'ownerDepartment' '')).Trim() } else { '' }
    if ($owner.Length -gt 120) { throw 'æ‹…å½“éƒ¨ç½²ã¯120æ–‡å­—ä»¥å†…ã§å…¥åŠ›ã—ã¦ãã ã•ã„ã€‚' }
    $defaultTarget = if ($hasDefaultTarget) { ([string](Get-DataProperty $Patch 'defaultTargetId' '')).Trim().ToLowerInvariant() } else { '' }
    $requirementId = if ($hasRequirement) { ([string](Get-DataProperty $Patch 'requirementId' '')).Trim() } else { '' }
    return Update-StructureLocked $Language {
        param($structure)
        $source = @(Get-Array (Get-DataProperty $structure 'sources' @()) | Where-Object { [string](Get-DataProperty $_ 'sourceId' '') -eq $id } | Select-Object -First 1)
        if ($source.Count -eq 0) { throw 'æŒ‡å®šã•ã‚ŒãŸåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }
        $current = $source[0]
        $pack = Get-PackRecord $structure ([string](Get-DataProperty $current 'packId' ''))
        if ($hasDefaultTarget -and $defaultTarget -ne 'unassigned' -and @(Get-PackTargetIds $Language $pack) -notcontains $defaultTarget) { throw 'æ—¢å®šã®æŒ¯ã‚Šåˆ†ã‘å…ˆãŒä¸æ­£ã§ã™ã€‚' }
        $workbook = @(Get-Array (Get-DataProperty $structure 'workbooks' @()) | Where-Object { [string](Get-DataProperty $_ 'workbookId' '') -eq $id } | Select-Object -First 1)
        $requirement = $null
        if ($hasRequirement -and -not [string]::IsNullOrWhiteSpace($requirementId)) {
            [void](Assert-SafeStorageSegment $requirementId 'requirementId')
            $matches = @(
                Get-Array (Get-DataProperty (Get-PackEffectiveTemplate $Language $pack) 'sourceRequirements' @()) |
                    Where-Object { [string](Get-DataProperty $_ 'requirementId' '') -eq $requirementId } |
                    Select-Object -First 1
            )
            if ($matches.Count -eq 0) { throw 'æŒ‡å®šã•ã‚ŒãŸå¿…è¦åŽŸç¨¿æž ãŒã“ã®ä¸€å¼ã«ã‚ã‚Šã¾ã›ã‚“ã€‚' }
            $requirement = $matches[0]
            $sourceType = ([string](Get-DataProperty $current 'sourceType' 'excel')).ToLowerInvariant()
            if ($sourceType -notin @(Get-Array (Get-DataProperty $requirement 'acceptedSourceTypes' @('excel','word','pdf','powerpoint')))) { throw 'ã“ã®å¿…è¦åŽŸç¨¿æž ã§ã¯é¸æŠžã—ãŸåŽŸç¨¿å½¢å¼ã‚’åˆ©ç”¨ã§ãã¾ã›ã‚“ã€‚' }
            $alreadyAssigned = @(
                Get-Array (Get-DataProperty $structure 'sources' @()) |
                    Where-Object {
                        [string](Get-DataProperty $_ 'sourceId' '') -ne $id -and
                        [string](Get-DataProperty $_ 'packId' '') -eq [string](Get-DataProperty $current 'packId' '') -and
                        [string](Get-DataProperty $_ 'requirementId' '') -eq $requirementId
                    }
            )
            if ($alreadyAssigned.Count -gt 0) { throw 'ã“ã®å¿…è¦åŽŸç¨¿æž ã«ã¯åˆ¥ã®åŽŸç¨¿ãŒå‰²ã‚Šå½“ã¦æ¸ˆã¿ã§ã™ã€‚å…ˆã«å‰²ã‚Šå½“ã¦ã‚’è§£é™¤ã—ã¦ãã ã•ã„ã€‚' }
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
        return [pscustomobject][ordered]@{
            sourceId = $id
            ownerDepartment = [string](Get-DataProperty $current 'ownerDepartment' '')
            required = [bool](Get-DataProperty $current 'required' $true)
            defaultTargetId = [string](Get-DataProperty $current 'defaultTargetId' 'unassigned')
            requirementId = [string](Get-DataProperty $current 'requirementId' '')
        }
    }
}

function New-PowerPointSourceObject([string]$RelativePath, [string]$Language, [string]$Category = '', [int]$SlideCount = 0) {
    $paths = Get-Paths
    $full = Join-Safe ([string]$paths.submissionDir) $RelativePath
    if (-not (Test-Path -LiteralPath $full)) { throw "æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $RelativePath" }
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
    if (-not (Test-Path -LiteralPath $PptxPath -PathType Leaf)) { throw "PowerPointåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $PptxPath" }
    $item = Get-Item -LiteralPath $PptxPath
    if ($item.Length -le 0) { throw 'PowerPointåŽŸç¨¿ãŒç©ºã§ã™ã€‚' }
    if ($item.Length -gt 1073741824) { throw '1GBã‚’è¶…ãˆã‚‹PowerPointåŽŸç¨¿ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚åˆ†å‰²ã—ã¦ã‹ã‚‰ç™»éŒ²ã—ã¦ãã ã•ã„ã€‚' }
    $stream = $null; $zip = $null
    try {
        $stream = [IO.File]::Open($PptxPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $header = New-Object byte[] 8
        $read = $stream.Read($header, 0, $header.Length)
        $stream.Dispose(); $stream = $null
        if ($read -ge 8 -and $header[0] -eq 0xD0 -and $header[1] -eq 0xCF -and $header[2] -eq 0x11 -and $header[3] -eq 0xE0) {
            throw 'ãƒ‘ã‚¹ãƒ¯ãƒ¼ãƒ‰ä¿è­·ã¾ãŸã¯æš—å·åŒ–ã•ã‚ŒãŸPowerPointåŽŸç¨¿ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚ä¿è­·ã‚’è§£é™¤ã—ãŸPPTXã‚’ä½¿ç”¨ã—ã¦ãã ã•ã„ã€‚'
        }
        if ($read -lt 4 -or $header[0] -ne 0x50 -or $header[1] -ne 0x4B) { throw 'PowerPointåŽŸç¨¿ã®ãƒ•ã‚¡ã‚¤ãƒ«å½¢å¼ã‚’ç¢ºèªã§ãã¾ã›ã‚“ã€‚æ‹¡å¼µå­ã ã‘ãŒPPTXã«ãªã£ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
        $zip = [IO.Compression.ZipFile]::OpenRead($PptxPath)
        $names = @($zip.Entries | ForEach-Object { ([string]$_.FullName -replace '\\','/').ToLowerInvariant() })
        if ($names -notcontains '[content_types].xml' -or $names -notcontains 'ppt/presentation.xml') { throw 'PowerPointåŽŸç¨¿ã®å†…éƒ¨æ§‹é€ ãŒå£Šã‚Œã¦ã„ã¾ã™ã€‚PowerPointã§é–‹ã„ã¦åˆ¥åä¿å­˜ã—ã¦ã‹ã‚‰å†ç™»éŒ²ã—ã¦ãã ã•ã„ã€‚' }
        if (@($names | Where-Object { $_ -eq 'ppt/vbaproject.bin' }).Count -gt 0) { throw 'ãƒžã‚¯ãƒ­ã‚’å«ã‚€PowerPointåŽŸç¨¿ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚.pptxå½¢å¼ã§ä¿å­˜ã—ã¦ãã ã•ã„ã€‚' }
        $slides = @($names | Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' } | Sort-Object { [int]([regex]::Match($_, 'slide(\d+)\.xml$').Groups[1].Value) })
        if ($slides.Count -eq 0) { throw 'PowerPointåŽŸç¨¿ã«ã‚¹ãƒ©ã‚¤ãƒ‰ãŒã‚ã‚Šã¾ã›ã‚“ã€‚' }
        if ($slides.Count -gt 5000) { throw '5000ã‚¹ãƒ©ã‚¤ãƒ‰ã‚’è¶…ãˆã‚‹PowerPointåŽŸç¨¿ã¯ç™»éŒ²ã§ãã¾ã›ã‚“ã€‚åˆ†å‰²ã—ã¦ã‹ã‚‰ç™»éŒ²ã—ã¦ãã ã•ã„ã€‚' }
        $units = @()
        for ($index = 1; $index -le $slides.Count; $index++) {
            $units += [pscustomobject][ordered]@{ unitKind='slide'; sourceKey="slide:$index"; title="Slide $index"; sourceIndex=$index }
        }
        return [pscustomobject][ordered]@{ pageCount=$slides.Count; units=@($units); warnings=@() }
    } catch {
        if ($_.Exception.Message -match 'PowerPointåŽŸç¨¿|PPTX|ãƒ‘ã‚¹ãƒ¯ãƒ¼ãƒ‰|ãƒžã‚¯ãƒ­|ã‚¹ãƒ©ã‚¤ãƒ‰') { throw }
        throw ('PowerPointåŽŸç¨¿ã‚’å®‰å…¨ã«èª­ã¿è¾¼ã‚ã¾ã›ã‚“ã§ã—ãŸã€‚ç ´æã—ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚ è©³ç´°: ' + $_.Exception.Message)
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
    $manual = @($existing | Where-Object { [bool](Get-DataProperty $_ 'orderManual' $false) }).Count -gt 0
    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($p in $existing) { [void]$ordered.Add($p) }
    $insertedAtEnd = $manual
    if ($manual -or $existing.Count -eq 0) {
        [void]$ordered.Add($NewPage)
    } else {
        $newWb = @(Get-Array $Structure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$NewPage.workbookId } | Select-Object -First 1)
        $newFile = if ($newWb.Count) { [string]$newWb[0].fileName } else { '' }
        $newFileOrder = Get-FileOrderNumber $newFile
        $newSheetIndex = Get-PageSheetIndex $NewPage
        $at = $existing.Count
        for ($i=0; $i -lt $existing.Count; $i++) {
            $p = $existing[$i]
            $wb = @(Get-Array $Structure.workbooks | Where-Object { [string]$_.workbookId -eq [string]$p.workbookId } | Select-Object -First 1)
            $file = if ($wb.Count) { [string]$wb[0].fileName } else { '' }
            $fileOrder = Get-FileOrderNumber $file
            $sheetIndex = Get-PageSheetIndex $p
            $fileCompare = [StringComparer]::OrdinalIgnoreCase.Compare($file, $newFile)
            if ($fileOrder -gt $newFileOrder -or
                ($fileOrder -eq $newFileOrder -and $fileCompare -gt 0) -or
                ($fileOrder -eq $newFileOrder -and $fileCompare -eq 0 -and $sheetIndex -gt $newSheetIndex)) {
                $at = $i
                break
            }
        }
        $ordered.Insert($at, $NewPage)
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
        if ($version -gt 3) { throw "ã“ã®ç®¡ç†ãƒ‡ãƒ¼ã‚¿ã¯æ–°ã—ã„schemaVersion=$versionã§ã™ã€‚å¯¾å¿œã™ã‚‹ReportBinderã‚’ä½¿ç”¨ã—ã¦ãã ã•ã„ã€‚" }
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
                if ((Get-Item $v1Backup).Length -ne (Get-Item $path).Length -or (New-StableHash $v1Backup) -ne (New-StableHash $path)) { throw 'structure.json.v1.bakã®æ¤œè¨¼ã«å¤±æ•—ã—ã¾ã—ãŸã€‚' }
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
                if ((Get-Item $v2Backup).Length -ne (Get-Item $path).Length -or (New-StableHash $v2Backup) -ne (New-StableHash $path)) { throw 'structure.json.v2.bakã®æ¤œè¨¼ã«å¤±æ•—ã—ã¾ã—ãŸã€‚' }
            } else {
                Write-JsonFile $v2Backup $structure
                $verifiedV2 = Read-JsonFile $v2Backup $null
                if (-not (Test-StructureDocument $verifiedV2 $Language) -or [int](Get-DataProperty $verifiedV2 'schemaVersion' 0) -ne 2) { throw 'structure.json.v2.bakã®æ¤œè¨¼ã«å¤±æ•—ã—ã¾ã—ãŸã€‚' }
            }
        }
        $backups += $v2Backup
        $structure = ConvertTo-StructureV3 $structure $Language $workspace
        Write-StructureUnlocked $Language $structure $DataDir
        $savedV3 = Read-JsonFile $path $null
        if (-not (Test-StructureDocument $savedV3 $Language) -or [int](Get-DataProperty $savedV3 'schemaVersion' 0) -ne 3) { throw 'schemaVersion 3ã®ä¿å­˜å¾Œæ¤œè¨¼ã«å¤±æ•—ã—ã¾ã—ãŸã€‚' }
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
    $rawTail = if ($text.Length -gt 250) { ' (å…ƒã®ã‚¨ãƒ©ãƒ¼: ' + $text.Substring(0, 250) + 'â€¦)' } else { ' (å…ƒã®ã‚¨ãƒ©ãƒ¼: ' + $text + ')' }
    # ç›£è¦–ãƒ—ãƒ­ã‚»ã‚¹ãŒ Excel ã‚’çµ‚äº†ã•ã›ã¦ã„ãŸã‚‰ã€ä»¥é™ã®COMå‘¼ã³å‡ºã—ã¯ã™ã¹ã¦RPCã®å¤±æ•—ã¨ã—ã¦
    # è¿”ã‚‹ã€‚é–‹ããƒ»é–‰ã˜ã‚‹ãƒ»ä¸€æ‹¬æ›¸ãå‡ºã—ãªã©ã€ã©ã“ã§è¸ã‚“ã§ã‚‚åŽŸå› ã¯åŒã˜ãªã®ã§æœ€åˆã«åˆ¤å®šã™ã‚‹ã€‚
    if (Test-ExcelRenderWatchdogFired) {
        return 'Excelã§ã®å¤‰æ›ãŒé€²ã¾ãªããªã£ãŸãŸã‚ä¸­æ­¢ã—ã¾ã—ãŸã€‚Excelã®ç”»é¢ã«ç¢ºèªã®ãƒ€ã‚¤ã‚¢ãƒ­ã‚°ãŒå‡ºã¦ã„ãªã„ã‹ç¢ºã‹ã‚ã¦ã‹ã‚‰ã€ã‚‚ã†ä¸€åº¦ãŠè©¦ã—ãã ã•ã„ã€‚åŽŸç¨¿ãŒå¤§ãã„å ´åˆã¯ã€ã‚·ãƒ¼ãƒˆæ•°ã‚’æ¸›ã‚‰ã™ã‹åˆ†å‰²ã™ã‚‹ã¨é€šã‚‹ã“ã¨ãŒã‚ã‚Šã¾ã™ã€‚'
    }
    if ($text -match 'WORD_TIMEOUT') { return 'WordåŽŸç¨¿ã®PDFå¤‰æ›ãŒæ™‚é–“å†…ã«å®Œäº†ã—ã¾ã›ã‚“ã§ã—ãŸã€‚Wordã®ç¢ºèªç”»é¢ãŒé–‹ã„ã¦ã„ãªã„ã‹ã€æ–‡æ›¸ãŒç ´æã—ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚å‰å›žæˆåŠŸã—ãŸå¤‰æ›PDFã¯ä¿æŒã•ã‚Œã¦ã„ã¾ã™ã€‚' }
    if ($text -match 'WORD_NOT_AVAILABLE|ActiveX component can.t create object|Class not registered') { return 'Microsoft Wordã‚’èµ·å‹•ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚ã“ã®PCã«ãƒ‡ã‚¹ã‚¯ãƒˆãƒƒãƒ—ç‰ˆWordãŒã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ã•ã‚Œã€é€šå¸¸èµ·å‹•ã§ãã‚‹ã“ã¨ã‚’ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'WORD_OPEN_FAILED') { return 'WordãŒåŽŸç¨¿ã‚’é–‹ã‘ã¾ã›ã‚“ã§ã—ãŸã€‚ãƒ‘ã‚¹ãƒ¯ãƒ¼ãƒ‰ä¿è­·ã€ç§˜å¯†åº¦ãƒ©ãƒ™ãƒ«ã€ç ´æã€å¤‰æ›ç¢ºèªãŒå¿…è¦ãªæ–‡æ›¸ã§ã¯ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'WORD_EXPORT_FAILED|WORD_PDF_MISSING|WORD_PDF_EMPTY|WORD_PDF_INVALID') { return 'WordåŽŸç¨¿ã‹ã‚‰PDFã‚’ä½œæˆã§ãã¾ã›ã‚“ã§ã—ãŸã€‚å°åˆ·ãƒ¬ã‚¤ã‚¢ã‚¦ãƒˆã€ãƒ—ãƒªãƒ³ã‚¿ãƒ¼è¨­å®šã€æ–‡æ›¸ã®ä¿è­·çŠ¶æ…‹ã‚’ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'WordåŽŸç¨¿|DOCX|word-com|WORD_WORKER') { return ('WordåŽŸç¨¿ã‚’PDFå¤‰æ›ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚' + $rawTail) }
    if ($text -match 'POWERPOINT_TIMEOUT') { return 'PowerPointåŽŸç¨¿ã®PDFå¤‰æ›ãŒæ™‚é–“å†…ã«å®Œäº†ã—ã¾ã›ã‚“ã§ã—ãŸã€‚PowerPointã®ç¢ºèªç”»é¢ãŒé–‹ã„ã¦ã„ãªã„ã‹ã€åŽŸç¨¿ãŒç ´æã—ã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚å‰å›žæˆåŠŸã—ãŸå¤‰æ›PDFã¯ä¿æŒã•ã‚Œã¦ã„ã¾ã™ã€‚' }
    if ($text -match 'POWERPOINT_NOT_AVAILABLE') { return 'Microsoft PowerPointã‚’èµ·å‹•ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚ã“ã®PCã«ãƒ‡ã‚¹ã‚¯ãƒˆãƒƒãƒ—ç‰ˆPowerPointãŒã‚¤ãƒ³ã‚¹ãƒˆãƒ¼ãƒ«ã•ã‚Œã€é€šå¸¸èµ·å‹•ã§ãã‚‹ã“ã¨ã‚’ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'POWERPOINT_OPEN_FAILED') { return 'PowerPointãŒåŽŸç¨¿ã‚’é–‹ã‘ã¾ã›ã‚“ã§ã—ãŸã€‚ãƒ‘ã‚¹ãƒ¯ãƒ¼ãƒ‰ä¿è­·ã€ç§˜å¯†åº¦ãƒ©ãƒ™ãƒ«ã€ç ´æã€å¤‰æ›ç¢ºèªãŒå¿…è¦ãªåŽŸç¨¿ã§ã¯ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'POWERPOINT_EXPORT_FAILED|POWERPOINT_PDF_MISSING|POWERPOINT_PDF_EMPTY|POWERPOINT_PDF_INVALID|POWERPOINT_SLIDE_COUNT_MISMATCH') { return 'PowerPointåŽŸç¨¿ã‹ã‚‰PDFã‚’ä½œæˆã§ãã¾ã›ã‚“ã§ã—ãŸã€‚ã‚¹ãƒ©ã‚¤ãƒ‰è¨­å®šã€åŽŸç¨¿ã®ä¿è­·çŠ¶æ…‹ã€åŸ‹ã‚è¾¼ã¿ãƒ¡ãƒ‡ã‚£ã‚¢ã‚’ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'PowerPointåŽŸç¨¿|PPTX|powerpoint-com|POWERPOINT_WORKER') { return ('PowerPointåŽŸç¨¿ã‚’PDFå¤‰æ›ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚' + $rawTail) }
    if ($text -match 'Open ãƒ—ãƒ­ãƒ‘ãƒ†ã‚£ã‚’å–å¾—ã§ãã¾ã›ã‚“|Unable to get the Open property') { return 'ExcelãŒã“ã®ãƒ•ã‚¡ã‚¤ãƒ«ã‚’é–‹ã‘ã¾ã›ã‚“ã§ã—ãŸã€‚ãƒ•ã‚¡ã‚¤ãƒ«ãŒä¿è­·ãƒ“ãƒ¥ãƒ¼å¯¾è±¡ï¼ˆã‚¤ãƒ³ã‚¿ãƒ¼ãƒãƒƒãƒˆç”±æ¥ï¼‰ãƒ»æš—å·åŒ–ï¼ˆç§˜å¯†åº¦ãƒ©ãƒ™ãƒ«/IRMï¼‰ãƒ»ç ´æã®ã„ãšã‚Œã‹ã®å¯èƒ½æ€§ãŒã‚ã‚Šã¾ã™ã€‚ãƒ•ã‚¡ã‚¤ãƒ«ã‚’å³ã‚¯ãƒªãƒƒã‚¯â†’ãƒ—ãƒ­ãƒ‘ãƒ†ã‚£â†’ã€Œè¨±å¯ã™ã‚‹ã€ã«ãƒã‚§ãƒƒã‚¯å¾Œã€ã‚‚ã†ä¸€åº¦PDFä½œæˆã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'ã“ã®ã‚ªãƒ–ã‚¸ã‚§ã‚¯ãƒˆã«ãƒ—ãƒ­ãƒ‘ãƒ†ã‚£|stateSavedAt|ãƒ—ãƒ­ãƒ‘ãƒ†ã‚£.*è¦‹ã¤ã‹ã‚Šã¾ã›ã‚“|property.*not found|does not contain a property') { return 'PDFä½œæˆã®é€²æ—çŠ¶æ…‹ã‚’æ›´æ–°ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚ReportBinderã‚’æ›´æ–°ã—ã¦ã‹ã‚‰ã€ã‚‚ã†ä¸€åº¦PDFä½œæˆã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'PDFåŒ–å¯¾è±¡ã®ã‚·ãƒ¼ãƒˆ|PDFåŒ–å¯¾è±¡ã®è¡¨ç¤ºã‚·ãƒ¼ãƒˆ') { return 'PDFåŒ–å¯¾è±¡ã®è¡¨ç¤ºã‚·ãƒ¼ãƒˆãŒã‚ã‚Šã¾ã›ã‚“ã€‚Excelã§å°‘ãªãã¨ã‚‚1ã¤ã®ãƒ¯ãƒ¼ã‚¯ã‚·ãƒ¼ãƒˆã‚’è¡¨ç¤ºã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'æå‡ºãƒ•ã‚¡ã‚¤ãƒ«|ãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“|ãƒ–ãƒƒã‚¯ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“|not found|missing') { return 'æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚å‰Šé™¤ãƒ»ç§»å‹•ãƒ»åå‰å¤‰æ›´ã•ã‚Œã¦ã„ãªã„ã‹ç¢ºèªã—ã¦ãã ã•ã„ã€‚' }
    if ($text -match 'ã‚³ãƒ”ãƒ¼å‰å¾Œ|æå‡ºä¸­|ä½¿ç”¨ä¸­|locked|lock|ãƒ­ãƒƒã‚¯') { return ('ExcelãŒä¿å­˜ä¸­ã¾ãŸã¯ä»–ã®å‡¦ç†ä¸­ã§ã™ã€‚ä¿å­˜ãŒçµ‚ã‚ã£ã¦ã‹ã‚‰å†åº¦ç¢ºèªã—ã¾ã™ã€‚' + $rawTail) }
    if ($text -match 'Excel|COM|HRESULT|ExportAsFixedFormat|RPC') { return ('Excelã§PDFåŒ–ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚' + $rawTail) }
    if ($text -match 'PDFã‚’ä½œæˆã§ãã¾ã›ã‚“|PDFãŒä½œæˆã•ã‚Œã¾ã›ã‚“|ç©ºã®PDF|0 bytes') { return ('Excelã‹ã‚‰PDFãŒå‡ºåŠ›ã•ã‚Œã¾ã›ã‚“ã§ã—ãŸã€‚å°åˆ·ç¯„å›²ã¨ã‚·ãƒ¼ãƒˆè¨­å®šã‚’ç¢ºèªã—ã¦ãã ã•ã„ã€‚' + $rawTail) }
    if ($text.Length -gt 120) { return $text.Substring(0, 120) + 'â€¦' }
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
    if ($detailText -match '(?<n>\d+)ä»¶') { $count = [int]$matches['n'] }
    if ($existing.Count -gt 0) {
        $oldCount = Get-IntDataProperty $existing[0] 'count' 0
        if ($oldCount -le 0) {
            $oldDetail = [string](Get-DataProperty $existing[0] 'detail' '')
            if ($oldDetail -match '(?<n>\d+)ä»¶') { $oldCount = [int]$matches['n'] } else { $oldCount = 1 }
        }
        $count += $oldCount
        if ($detailText -match '\d+ä»¶') { $detailText = [regex]::Replace($detailText, '\d+ä»¶', "$count`ä»¶", 1) }
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


# $Sheets ã«ã¯ã€Œä»Šå›žæç”»ã§ãã‚‹ã‚·ãƒ¼ãƒˆã€ï¼ˆExcelã§ã¯è¡¨ç¤ºä¸­ã®ã‚‚ã®ï¼‰ãŒå…¥ã‚‹ã€‚
# $PresentSheetNames ã«ã¯éžè¡¨ç¤ºã‚’å«ã‚€ã€ŒåŽŸç¨¿ã«å­˜åœ¨ã™ã‚‹ã‚·ãƒ¼ãƒˆåã€ã‚’æ¸¡ã™ã€‚çœç•¥æ™‚ã¯
# $Sheets ã¨åŒã˜ã¨ã¿ãªã™ï¼ˆWord/PowerPoint/PDFã®ã‚ˆã†ã«éžè¡¨ç¤ºã®æ¦‚å¿µãŒãªã„çµŒè·¯ï¼‰ã€‚
function Update-WorkbookPagesFromInspection([string]$Language, $Structure, $Workbook, $Sheets, $PresentSheetNames = $null) {
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

    $removed = @()
    $kept = @()
    $hidden = @()
    foreach ($page in $pages) {
        $belongs = ([string]$page.workbookId -eq $workbookId)
        $sheetKey = ([string]$page.sheetName).ToLowerInvariant()
        if (-not $belongs -or $current.ContainsKey($sheetKey)) { $kept += $page; continue }
        # éžè¡¨ç¤ºã«ã—ãŸã ã‘ã®ã‚·ãƒ¼ãƒˆã¯åŽŸç¨¿ã‹ã‚‰æ¶ˆãˆã¦ã„ãªã„ã€‚ãƒšãƒ¼ã‚¸ã‚’å‰Šé™¤ã™ã‚‹ã¨é…ç½®ãƒ»
        # ãƒšãƒ¼ã‚¸åãƒ»ä½¿ç”¨ç¯„å›²ãƒ»ç•ªå·è¨­å®šãŒå¤±ã‚ã‚Œã€å†è¡¨ç¤ºã—ã¦ã‚‚æœªæŒ¯ã‚Šåˆ†ã‘ã«æˆ»ã‚‹ã ã‘ã«ãªã‚‹ã€‚
        if ($present.ContainsKey($sheetKey)) {
            Set-NoteProperty $page 'sheetHidden' $true
            $hidden += (Resolve-PageId $page)
            $kept += $page
            continue
        }
        $removed += $page
    }
    if ($removed.Count -gt 0) {
        # å‰Šé™¤ã¯å–ã‚Šæ¶ˆã›ãªã„ãŸã‚ã€ç¢ºå®šå‰ã«å¿…ãšå¾©å…ƒãƒã‚¤ãƒ³ãƒˆã‚’æ®‹ã™ã€‚
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

    # ã‚·ãƒ¼ãƒˆåã®å¤‰æ›´ã¯ã€Œæ—§ã‚·ãƒ¼ãƒˆã®æ¶ˆæ»…ï¼‹æ–°ã‚·ãƒ¼ãƒˆã®å‡ºç¾ã€ã¨ã—ã¦ç¾ã‚Œã‚‹ã€‚ä½ç½®(sheetIndex)ã§
    # å¯¾å¿œä»˜ã‘ã€æ—§ãƒšãƒ¼ã‚¸ã‚’ä½œã‚Šç›´ã•ãšã«å¼•ãç¶™ãã€‚pageId ã‚’ä¿ã¤ã“ã¨ã§ã€ãƒ¬ã‚¤ã‚¢ã‚¦ãƒˆå±¥æ­´ã¨
    # Ctrl+Z ã‹ã‚‰ã®å¾©å…ƒçµŒè·¯ã‚‚ç”Ÿãæ®‹ã‚‹ã€‚
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
                # ä½ç½®ã‚‚å¤‰ã‚ã£ãŸå ´åˆã¯ã€A1ã‹ã‚‰æ‹¾ã£ãŸè¦‹å‡ºã—ãŒå®Œå…¨ä¸€è‡´ã™ã‚‹ã¨ãã ã‘åŒä¸€è¦–ã™ã‚‹ã€‚
                # æ‰‹æŽ›ã‹ã‚Šãªã—ã§å¯¾å¿œä»˜ã‘ã‚‹ã¨ã€ç„¡é–¢ä¿‚ãªæ–°è¦ã‚·ãƒ¼ãƒˆã«å‰ã®ãƒšãƒ¼ã‚¸ã®é…ç½®ã¨
                # ä½¿ç”¨ãƒšãƒ¼ã‚¸ç¯„å›²ã‚’å¼•ãç¶™ã„ã§ã—ã¾ã„ã€å‰Šé™¤ã‚ˆã‚Šåˆ†ã‹ã‚Šã«ãã„èª¤ã‚Šã«ãªã‚‹ã€‚
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
        # å¼•ãç¶™ã„ã ãƒšãƒ¼ã‚¸ã¯ã€Œå‰Šé™¤ã•ã‚ŒãŸã€æ‰±ã„ã«ã—ãªã„ã€‚
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
            Set-NoteProperty $page 'sheetHidden' $false
            if ([string]::IsNullOrWhiteSpace([string]$page.title)) { Set-NoteProperty $page 'title' $title }
            $updated += $pageId
            continue
        }

        if ($renameBySheetKey.ContainsKey($sheetKey)) {
            $page = $renameBySheetKey[$sheetKey]
            $previousSheet = [string]$page.sheetName
            # pageId ã¯æ®ãˆç½®ãã€‚å·®ã—æ›¿ãˆã‚‹ã¨ãƒ¬ã‚¤ã‚¢ã‚¦ãƒˆå±¥æ­´ã®å¾©å…ƒãŒå¯¾è±¡ã‚’è¦‹å¤±ã†ã€‚
            $pageId = Resolve-PageId $page
            Set-NoteProperty $page 'pageId' $pageId
            Set-NoteProperty $page 'sheetName' $sheet
            Set-NoteProperty $page 'sheetIndex' $sheetIndex
            Set-NoteProperty $page 'detectedTitle' $title
            Set-NoteProperty $page 'sheetHidden' $false
            Set-NoteProperty $page 'renamedFromSheetName' $previousSheet
            # å‚ç…§å…ˆã®PDFã¯æ—§ã‚·ãƒ¼ãƒˆåã§ä½œã‚‰ã‚Œã¦ã„ã‚‹ã€‚ä¸­èº«ã¯ä½œã‚Šç›´ã—ã«ãªã‚‹ã€‚
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

        $pageId = New-WorksheetPageId $workbookId $sheet
        $page = [pscustomobject][ordered]@{
            pageId=$pageId; workbookId=$workbookId; sheetName=$sheet; sheetIndex=$sheetIndex
            titleSource='A1'; detectedTitle=$title; title=$title
            volume=$newVolume; order=0; orderManual=$false
            numberingMode='visible'; numberingManual=$false; numberingDefault='first-page-none'
            enabled=($newVolume -ne 'none'); contentPdf=$null; status='not-rendered'; warnings=@(); updatedAt=New-NowIso
        }
        $insert = Insert-PageInSheetOrder $Structure $page $newVolume $packId
        $Structure.pages = @(Get-Array $Structure.pages) + @($page)
        if ($insert.insertedAtEnd) { $atEnd += $pageId } else { $inOrder++ }
        $added += $pageId
        $knownBySheet[$sheetKey] = $page
    }

    $pagePack = (Resolve-DocumentPackScope $Structure $packId $true).pack
    foreach ($volume in @(Get-PackVolumeList $Language $pagePack $true)) { [void](Renumber-VolumeOrder $Structure $volume $packId) }
    if ($removed.Count -gt 0) {
        $affected = @($removed | ForEach-Object { [string]$_.volume } | Where-Object { $_ -and $_ -ne 'none' } | Select-Object -Unique)
        if ($affected.Count -gt 0) {
            Mark-VolumeNeedsRebuild $Structure $Language $packId $affected 'render' 'åŽŸç¨¿ã®ãƒšãƒ¼ã‚¸æ§‹æˆãŒå¤‰æ›´ã•ã‚Œã¾ã—ãŸ'
        }
    }
    Apply-DefaultNumberingPerVolume $Language $Structure $packId
    $names = @($sortedSheets | ForEach-Object { [string]$_.sheetName })
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
        # PageSetup ã¯Excel COMã®ä¸­ã§ã‚‚ç‰¹ã«é‡ã„ã€‚è¤‡æ•°ã‚·ãƒ¼ãƒˆã®PDFä½œæˆæ™‚ã¯ã€å‘¼ã³å‡ºã—å…ƒã§
        # PrintCommunication ã‚’ä¸€æ™‚åœæ­¢ã—ã¦ã‹ã‚‰ã¾ã¨ã‚ã¦åæ˜ ã™ã‚‹ã“ã¨ã§ã€ã‚·ãƒ¼ãƒˆã”ã¨ã®å¾…ã¡æ™‚é–“ã‚’æ¸›ã‚‰ã™ã€‚
        if (-not $DeferPrintCommunication) {
            $printCommunicationChanged = Set-ExcelPrintCommunicationSafe $Excel $false
        }

        try { $Worksheet.DisplayPageBreaks = $false } catch { }
        $ps = $Worksheet.PageSetup
        # æš«å®šPDFã¯å·¦å³1.2cmã‚’åŸºæº–ã«ã™ã‚‹ã€‚æå‡ºç”¨PDFã§ã¯å¥‡æ•°/å¶æ•°ãƒšãƒ¼ã‚¸ã‚’0.2cmã ã‘å†…å´ã¸å¯„ã›ã‚‹(ãƒ‘ãƒ³ãƒå´1.4cm/å¤–å´1.0cm)ã€‚
        $ps.TopMargin = Convert-CmToPt 0.8
        $ps.BottomMargin = Convert-CmToPt 0.8
        $ps.LeftMargin = Convert-CmToPt 1.2
        $ps.RightMargin = Convert-CmToPt 1.2
        # å°åˆ·ç¯„å›²ãŒ1ãƒšãƒ¼ã‚¸å¹…ã«æº€ãŸãªã„ã‚·ãƒ¼ãƒˆãŒå·¦å¯„ã‚Šã«è¦‹ãˆãªã„ã‚ˆã†ã€æ°´å¹³æ–¹å‘ã®ã¿ä¸­å¤®æƒãˆã«ã™ã‚‹ã€‚
        # å¹…ã„ã£ã±ã„ã®ã‚·ãƒ¼ãƒˆã¯ä½™ç™½ã«æŽ¥ã™ã‚‹ãŸã‚ã€ã‚»ãƒ³ã‚¿ãƒªãƒ³ã‚°ã—ã¦ã‚‚ä½ç½®ã¯å¤‰ã‚ã‚‰ãªã„ã€‚
        $ps.CenterHorizontally = $true
        $ps.CenterVertically = $false
        # å°åˆ·ç¯„å›²ã¯å°Šé‡ã—ã¤ã¤ã€å€çŽ‡ã¯Excelè¨­å®šã‚’ç„¡è¦–ã—ã¦1ãƒšãƒ¼ã‚¸ã«åŽã‚ã‚‹ã€‚
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
    if ($null -eq $tool) { return [ordered]@{ ok = $false; reason = 'pdfbox-unavailable'; message = 'PDFBox/JavaãŒãªã„ãŸã‚ä¸€æ‹¬PDFåˆ†å‰²ã‚’ä½¿ã„ã¾ã›ã‚“ã€‚' } }
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
    if ($run.timedOut) { return [ordered]@{ ok = $false; reason = 'split-timeout'; message = ('PDFã®åˆ†å‰²ãŒ{0}åˆ†ä»¥å†…ã«çµ‚ã‚ã‚Šã¾ã›ã‚“ã§ã—ãŸã€‚' -f [int]($Script:PdfSplitTimeoutSeconds / 60)) } }
    if ($exit -ne 0) { return [ordered]@{ ok = $false; reason = 'split-failed'; message = $outputText } }
    foreach ($info in @($SheetInfos)) {
        if (-not (Test-Path -LiteralPath ([string]$info.outPdf))) { return [ordered]@{ ok = $false; reason = 'split-missing-output'; message = "åˆ†å‰²å¾ŒPDFãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $($info.outPdf)" } }
        try {
            if ((Get-Item -LiteralPath ([string]$info.outPdf)).Length -le 0) { return [ordered]@{ ok = $false; reason = 'split-empty-output'; message = "åˆ†å‰²å¾ŒPDFãŒç©ºã§ã™: $($info.outPdf)" } }
        } catch { return [ordered]@{ ok = $false; reason = 'split-check-failed'; message = $_.Exception.Message } }
    }
    return [ordered]@{ ok = $true; message = $outputText }
}

function Export-WorkbookSheetsToPdfBatch($Excel, $Workbook, $SheetInfos, [string]$TmpDir) {
    $infos = @($SheetInfos)
    if ($infos.Count -le 1) { return [ordered]@{ ok = $false; reason = 'single-sheet'; message = '1ã‚·ãƒ¼ãƒˆã®ãŸã‚ä¸€æ‹¬PDFåŒ–ã—ã¾ã›ã‚“ã€‚' } }
    if ($null -eq (Get-PdfBatchToolInfo)) { return [ordered]@{ ok = $false; reason = 'pdfbox-unavailable'; message = 'PDFBox/JavaãŒãªã„ãŸã‚å¾“æ¥æ–¹å¼ã§PDFåŒ–ã—ã¾ã™ã€‚' } }

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
            # é¸æŠžã—ãŸè¤‡æ•°ã‚·ãƒ¼ãƒˆã‚’1å›žã§PDFåŒ–ã™ã‚‹ã€‚ã‚·ãƒ¼ãƒˆã”ã¨ã®ExportAsFixedFormatå›žæ•°ã‚’æ¸›ã‚‰ã™ã®ãŒç‹™ã„ã€‚
            # ã“ã®1å›žã®åŒæœŸå‘¼ã³å‡ºã—ã®å†…å´ã§ã¯å¿ƒæ‹ã‚’æ‰“ã¦ãªã„ã®ã§ã€ã‚·ãƒ¼ãƒˆæ•°ã¶ã‚“çŒ¶äºˆã‚’ä¼¸ã°ã—ã¦ã‹ã‚‰å…¥ã‚‹ã€‚
            Update-ExcelRenderHeartbeat (Get-ExcelBatchAllowanceSeconds $infos.Count)
            $active.ExportAsFixedFormat(0, $batchPdf, 0, $true, $false, $missing, $missing, $false, $missing)
            Update-ExcelRenderHeartbeat
        } finally {
            Invoke-ComRelease $active
        }
        [void](Wait-ForPdfOutput $batchPdf 'ä¸€æ‹¬PDF')
        $pageCount = Get-PdfPageCount $batchPdf
        if ($pageCount -ne $infos.Count) {
            return [ordered]@{ ok = $false; reason = 'page-count-mismatch'; message = "ä¸€æ‹¬PDFã®ãƒšãƒ¼ã‚¸æ•°ãŒæƒ³å®šã¨ç•°ãªã‚Šã¾ã™ã€‚æƒ³å®š=$($infos.Count) å®Ÿéš›=$pageCount" }
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
    if (-not (Test-Path -LiteralPath $PdfPath)) { throw "Excelã‹ã‚‰PDFãŒå‡ºåŠ›ã•ã‚Œã¾ã›ã‚“ã§ã—ãŸ: ã‚·ãƒ¼ãƒˆ $SheetName" }
    $item2 = Get-Item -LiteralPath $PdfPath
    if ($item2.Length -le 0) { throw "Excelã‹ã‚‰ç©ºã®PDFãŒå‡ºåŠ›ã•ã‚Œã¾ã—ãŸ: ã‚·ãƒ¼ãƒˆ $SheetName" }
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
        # 1ã‚·ãƒ¼ãƒˆã”ã¨ã«å¿ƒæ‹ã‚’æ‰“ã¤ã€‚æ­¢ã¾ã£ãŸã¾ã¾ä¸€å®šæ™‚é–“ãŒéŽãŽãŸã‚‰ç›£è¦–ãƒ—ãƒ­ã‚»ã‚¹ãŒ
        # ã“ã¡ã‚‰ã® Excel ã‚’çµ‚äº†ã•ã›ã€ã“ã®å‘¼ã³å‡ºã—ã¯ RPC ã®å¤±æ•—ã¨ã—ã¦æˆ»ã‚‹ã€‚
        Update-ExcelRenderHeartbeat
        $Worksheet.ExportAsFixedFormat(0, $OutPdf, 0, $true, $false, $missing, $missing, $false, $missing)
        Update-ExcelRenderHeartbeat
        return (Wait-ForPdfOutput $OutPdf $SheetName)
    } catch {
        $directError = $_.Exception.Message
        if (Test-ExcelRenderWatchdogFired) {
            throw ("Excelã§ã®å¤‰æ›ãŒ{0}åˆ†ä»¥ä¸Šé€²ã¾ãªã‹ã£ãŸãŸã‚ä¸­æ­¢ã—ã¾ã—ãŸ: ã‚·ãƒ¼ãƒˆ {1}ã€‚Excelã®ç”»é¢ã«ç¢ºèªã®ãƒ€ã‚¤ã‚¢ãƒ­ã‚°ãŒå‡ºã¦ã„ãªã„ã‹ç¢ºã‹ã‚ã¦ã‹ã‚‰ã€ã‚‚ã†ä¸€åº¦ãŠè©¦ã—ãã ã•ã„ã€‚" -f [int]($Script:ExcelRenderTimeoutSeconds / 60), $SheetName)
        }
        throw "Excelã§PDFåŒ–ã§ãã¾ã›ã‚“ã§ã—ãŸ: ã‚·ãƒ¼ãƒˆ $SheetName / ç›´æŽ¥å‡ºåŠ›=[$directError]"
    } finally {
        try { $Workbook.Activate() | Out-Null } catch { }
    }
}

$Script:PdfPageCountCache = @{}
$Script:PdfPageCountCacheLimit = 4096

# å¤‰æ›PDFã®ãƒšãƒ¼ã‚¸æ•°ã¯ /api/state ã®ãŸã³ã«å…¨ãƒšãƒ¼ã‚¸åˆ†ãŒå¿…è¦ã«ãªã‚‹ã€‚æ¯Žå›žãƒ•ã‚¡ã‚¤ãƒ«å…¨ä½“ã‚’
# èª­ã¿ç›´ã™ã¨dataDirå…¨é‡ã®èª­ã¿è¾¼ã¿ã«ãªã‚‹ãŸã‚ã€ãƒ‘ã‚¹ã¨ã‚µã‚¤ã‚ºãƒ»æ›´æ–°æ™‚åˆ»ã§ãƒ¡ãƒ¢åŒ–ã™ã‚‹ã€‚
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


# Excel ã® PDF æ›¸ãå‡ºã—ã¯åŒä¸€ãƒ—ãƒ­ã‚»ã‚¹å†…ã®åŒæœŸCOMå‘¼ã³å‡ºã—ã§ã€Word/PowerPoint ã®ã‚ˆã†ãª
# åˆ¥ãƒ—ãƒ­ã‚»ã‚¹ã®ãƒ¯ãƒ¼ã‚«ãƒ¼ã‚’é€šã‚‰ãªã„ã€‚å›ºã¾ã‚‹ã¨å‘¼ã³å‡ºã—å…ƒã”ã¨æ­¢ã¾ã‚Šã€ã€ŒPDFä½œæˆã‚’ä¸­æ­¢ã€ã‚‚
# ç¾åœ¨ã®åŽŸç¨¿ãŒçµ‚ã‚ã‚‹ã¾ã§åŠ¹ã‹ãªã„ãŸã‚ã€åˆ©ç”¨è€…ã«ã¯ã‚¿ã‚¹ã‚¯ãƒžãƒãƒ¼ã‚¸ãƒ£ãƒ¼ä»¥å¤–ã®è„±å‡ºæ‰‹æ®µãŒç„¡ã„ã€‚
# PowerShell ã®ã‚¿ã‚¤ãƒžãƒ¼ã‚„ã‚¤ãƒ™ãƒ³ãƒˆã§ã¯æ•‘ãˆãªã„(COMå‘¼ã³å‡ºã—ã§ãƒ‘ã‚¤ãƒ—ãƒ©ã‚¤ãƒ³ãŒå¡žãŒã£ã¦ã„ã‚‹
# é–“ã€ãƒãƒ³ãƒ‰ãƒ©ã¯å®Ÿè¡Œã•ã‚Œãªã„)ã€‚ãã“ã§å¤–éƒ¨ã®ç›£è¦–ãƒ—ãƒ­ã‚»ã‚¹ã«è¦‹å¼µã‚‰ã›ã‚‹ã€‚
# é€²æ—ãŒã‚ã‚‹ãŸã³ã«å¿ƒæ‹ãƒ•ã‚¡ã‚¤ãƒ«ã‚’æ›´æ–°ã—ã€ä¸€å®šæ™‚é–“æ›´æ–°ãŒæ­¢ã¾ã£ãŸã‚‰ã€ã“ã¡ã‚‰ãŒèµ·å‹•ã—ãŸ
# Excel ã ã‘ã‚’çµ‚äº†ã•ã›ã‚‹ã€‚COMå‘¼ã³å‡ºã—ã¯ RPC ã®å¤±æ•—ã¨ã—ã¦æˆ»ã‚Šã€æ—¢å­˜ã®å¤±æ•—å‡¦ç†ã«è¼‰ã‚‹ã€‚
function Get-OwnedExcelProcessId($Excel) {
    if (-not ('ReportBinderNative.OfficeWindows' -as [type])) { return 0 }
    try {
        [uint32]$pidValue = 0
        [void][ReportBinderNative.OfficeWindows]::GetWindowThreadProcessId([IntPtr][int]$Excel.Hwnd, [ref]$pidValue)
        return [int]$pidValue
    } catch { return 0 }
}

# å¿ƒæ‹ã«ã¯ã€Œæ¬¡ã®å¿ƒæ‹ã¾ã§ã«è¨±ã™æ™‚é–“ã€ã‚‚ä¸€ç·’ã«æ›¸ãã€‚Excel ã® PDF æ›¸ãå‡ºã—ã¯ã€è¤‡æ•°ã‚·ãƒ¼ãƒˆã‚’
# 1å›žã® ExportAsFixedFormat ã§å‡ºã™çµŒè·¯(Export-WorkbookSheetsToPdfBatch)ãŒæ—¢å®šã§ã€
# ãã®1å›žã®åŒæœŸå‘¼ã³å‡ºã—ã®å†…å´ã§ã¯å¿ƒæ‹ã‚’æ‰“ã¦ãªã„ã€‚å›ºå®šã®çŒ¶äºˆã ã¨ã€æ­£å¸¸ã«å‹•ã„ã¦ã„ã‚‹
# å¤§ããªãƒ–ãƒƒã‚¯ã‚’æ®ºã—ã¦ã—ã¾ã†ãŸã‚ã€ã“ã‚Œã‹ã‚‰å§‹ã‚ã‚‹ä½œæ¥­ã®é‡ã•ã«å¿œã˜ã¦çŒ¶äºˆã‚’ä¼¸ã°ã™ã€‚
function Update-ExcelRenderHeartbeat([int]$AllowanceSeconds = 0) {
    if ([string]::IsNullOrWhiteSpace($Script:ExcelWatchdogHeartbeatPath)) { return }
    $allowance = if ($AllowanceSeconds -gt 0) { $AllowanceSeconds } else { [int]$Script:ExcelRenderTimeoutSeconds }
    try { [IO.File]::WriteAllText($Script:ExcelWatchdogHeartbeatPath, ('{0} {1}' -f [DateTimeOffset]::UtcNow.ToUnixTimeSeconds(), $allowance)) } catch { }
}

# ç›£è¦–ãƒ—ãƒ­ã‚»ã‚¹ã®ä½œæ¥­ãƒ•ã‚¡ã‚¤ãƒ«ã¯ã€ã‚µãƒ¼ãƒãƒ¼ãŒå¼·åˆ¶çµ‚äº†ã•ã‚Œã‚‹ã¨æ®‹ã‚‹ã€‚æ”¾ã£ã¦ãŠãã¨
# %TEMP% ã«ç„¡æœŸé™ã«æºœã¾ã‚‹ã®ã§ã€ç›£è¦–ã‚’å§‹ã‚ã‚‹ãŸã³ã«å¤ã„ã‚‚ã®ã‚’æŽƒé™¤ã™ã‚‹ã€‚
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

# ã‚·ãƒ¼ãƒˆæ•°ã«å¿œã˜ãŸçŒ¶äºˆã€‚1ã‚·ãƒ¼ãƒˆã‚ãŸã‚Šã®ä¸Šé™ã‚’ç©ã¿ã€ä¸‹é™ã¯æ—¢å®šå€¤ã€‚
function Get-ExcelBatchAllowanceSeconds([int]$SheetCount) {
    $perSheet = [int]$Script:ExcelPerSheetAllowanceSeconds
    $scaled = [int]$Script:ExcelRenderTimeoutSeconds + ($perSheet * [Math]::Max(0, $SheetCount))
    if ($scaled -lt [int]$Script:ExcelRenderTimeoutSeconds) { return [int]$Script:ExcelRenderTimeoutSeconds }
    return $scaled
}

function Start-ExcelRenderWatchdog($Excel) {
    # ç›£è¦–ã®æž ã¯1çµ„ã—ã‹ãªã„ã€‚å‰ã® Excel ãŒç”Ÿãã¦ã„ã‚‹ã†ã¡ã«2ã¤ç›®ã‚’ä½œã‚‹ã¨ã€1ã¤ç›®ã®å¿ƒæ‹ã®
    # å ´æ‰€ã‚’è¦‹å¤±ã„ã€æ›´æ–°ã•ã‚Œãªã„ã¾ã¾ç¾å½¹ã® Excel ãŒçµ‚äº†ã•ã›ã‚‰ã‚Œã‚‹ã€‚ä»Šã¯çµŒè·¯ã®ç´„æŸã ã‘ã§
    # äºŒé‡èµ·å‹•ã‚’é¿ã‘ã¦ã„ã‚‹ã®ã§ã€ç´„æŸãŒç ´ã‚ŒãŸã“ã¨ã‚’ã“ã“ã§æ¤œçŸ¥ã™ã‚‹ã€‚
    if (-not [string]::IsNullOrWhiteSpace($Script:ExcelWatchdogHeartbeatPath) -and (Test-Path -LiteralPath $Script:ExcelWatchdogHeartbeatPath)) {
        throw 'Excelã®ç›£è¦–ãŒã™ã§ã«å‹•ã„ã¦ã„ã¾ã™ã€‚å‰ã®å¤‰æ›ã‚’çµ‚ãˆã¦ã‹ã‚‰æ¬¡ã‚’é–‹å§‹ã—ã¦ãã ã•ã„ã€‚'
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
    # å¿ƒæ‹ãƒ•ã‚¡ã‚¤ãƒ«ãŒæ¶ˆãˆãŸã‚‰ä»•äº‹ãŒçµ‚ã‚ã£ãŸåˆå›³ã€‚ç›£è¦–ã‚‚çµ‚ãˆã‚‹ã€‚
    $command = @"
`$ErrorActionPreference = 'SilentlyContinue'
while (`$true) {
    Start-Sleep -Seconds 2
    if (-not (Test-Path -LiteralPath '$beatEscaped')) { break }
    `$raw = ''
    try { `$raw = [IO.File]::ReadAllText('$beatEscaped') } catch { continue }
    # "<unixç§’> <ãã®ä½œæ¥­ã«è¨±ã™ç§’æ•°>"ã€‚çŒ¶äºˆã¯ä½œæ¥­ã”ã¨ã«å¤‰ã‚ã‚‹(è¤‡æ•°ã‚·ãƒ¼ãƒˆã®ä¸€æ‹¬æ›¸ãå‡ºã—ãªã©)ã€‚
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
    # ç›£è¦–ã‚’æ­¢ã‚ãŸå¾Œã§ã‚‚åˆ¤å®šã§ãã‚‹ã‚ˆã†ã«ã™ã‚‹ã€‚å¤±æ•—ã®æ–‡è¨€ã‚’çµ„ã¿ç«‹ã¦ã‚‹ã®ã¯ã€
    # å¾Œç‰‡ä»˜ã‘ãŒçµ‚ã‚ã£ã¦ã‹ã‚‰ã®ã“ã¨ãŒã‚ã‚‹ãŸã‚ã€‚
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
    # ç›£è¦–ãƒ—ãƒ­ã‚»ã‚¹ã®æ¨™æº–å‡ºåŠ›ãƒ»æ¨™æº–ã‚¨ãƒ©ãƒ¼ã®ãƒ•ã‚¡ã‚¤ãƒ«ã‚‚æ¶ˆã™ã€‚Excelã‚’èµ·å‹•ã™ã‚‹ãŸã³2ã¤å¢—ãˆã‚‹ã€‚
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
    # Quit() ã‚‚åŒæœŸCOMå‘¼ã³å‡ºã—ã§ã€ç¢ºèªãƒ€ã‚¤ã‚¢ãƒ­ã‚°ã‚„ä¿ç•™ã‚¤ãƒ™ãƒ³ãƒˆã§å›ºã¾ã‚Šã†ã‚‹ã€‚
    # ç›£è¦–ã‚’å…ˆã«æ­¢ã‚ã‚‹ã¨ã€ã„ã¡ã°ã‚“å›ºã¾ã‚Šã‚„ã™ã„æ‰€ã ã‘ç„¡é˜²å‚™ã«ãªã‚‹ã€‚é–‰ã˜çµ‚ãˆã¦ã‹ã‚‰æ­¢ã‚ã‚‹ã€‚
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
    # V5-Â§6.4: ç”»åƒãƒãƒƒã‚·ãƒ¥ã®æ¯”è¼ƒå¯å¦ã‚’å·¦å³ã™ã‚‹è¦ç´ ã ã‘ã‚’æŒ‡ç´‹ã«ã™ã‚‹ã€‚
    # pcName / userName ã¯è¨ºæ–­æƒ…å ±ã§ã‚ã‚Šã€æ¯”è¼ƒã®ä¸»åˆ¤å®šã«ã¯å«ã‚ãªã„
    # (åˆ¥PCã§ã‚‚ç’°å¢ƒãŒåŒç­‰ãªã‚‰åŒã˜PDFã«ãªã‚Šã†ã‚‹ãŸã‚)ã€‚
    if ($null -eq $EnvInfo) { return '' }
    $parts = @(
        'excelVersion=' + [string](Get-DataProperty $EnvInfo 'excelVersion' '')
        'osVersion='    + [string](Get-DataProperty $EnvInfo 'osVersion' '')
        'printer='      + [string](Get-DataProperty $EnvInfo 'activePrinter' '')
        'arial='        + [string](Get-DataProperty $EnvInfo 'hasArial' $false)
        'msgothic='     + [string](Get-DataProperty $EnvInfo 'hasMsgothic' $false)
        'printProfile=' + [string]$Script:ExcelPrintProfileVersion
    )
    # V5-P1: ç”»åƒãƒãƒƒã‚·ãƒ¥ã¯ PDFBox / Java / è§£æžæ–¹å¼ / DPI / è‰²ã«ã‚‚ä¾å­˜ã™ã‚‹ã€‚
    # ã“ã‚Œã‚‰ã‚’å«ã‚ãªã„ã¨ã€æ–¹å¼ã‚’å¤‰ãˆã¦ã‚‚æŒ‡ç´‹ãŒä¸€è‡´ã—ã¦äº’æ›æ€§ã®ãªã„ãƒãƒƒã‚·ãƒ¥ã‚’ç›´æŽ¥æ¯”è¼ƒã—ã¦ã—ã¾ã†ã€‚
    try {
        $vp = Get-VisualHashProfile
        $parts += ('pdfBox=' + [string](Get-DataProperty $vp 'pdfBoxVersion' ''))
        $parts += ('dpi=' + [string](Get-DataProperty $vp 'dpi' ''))
        $parts += ('color=' + [string](Get-DataProperty $vp 'colorMode' ''))
        $parts += ('profile=' + [string](Get-DataProperty $vp 'profileVersion' ''))
        $parts += ('analyzer=' + [string]$Script:PdfPageAnalyzerVersion)
        # V5-P1(#11): Java ã®ç‰ˆã‚’å®Ÿéš›ã«æŽ¡å–ã—ã¦æŒ‡ç´‹ã¸å«ã‚ã‚‹(ç©ºã®ã¾ã¾ã ã¨ Java æ›´æ–°å¾Œã‚‚æŒ‡ç´‹ãŒå¤‰ã‚ã‚‰ãšã€
        # äº’æ›æ€§ã®ãªã„ç”»åƒãƒãƒƒã‚·ãƒ¥ã‚’ç›´æŽ¥æ¯”è¼ƒã—ã¦ã—ã¾ã†)ã€‚
        $parts += ('java=' + [string](Get-JavaRuntimeSignature))
    } catch { }
    return (Get-Sha256Text ($parts -join '|'))
}

function Reset-RenderEnvironmentForJob {
    # V5-Â§6.4: ç’°å¢ƒã¯ã‚µãƒ¼ãƒãƒ¼ç¨¼åƒä¸­ã«å¤‰ã‚ã‚Šã†ã‚‹(é€šå¸¸ä½¿ã†ãƒ—ãƒªãƒ³ã‚¿ã®å¤‰æ›´ãªã©)ã€‚
    # ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ã‚¸ãƒ§ãƒ–ã®é–‹å§‹ã”ã¨ã«å–ã‚Šç›´ã™ã€‚ãƒ­ã‚°å‡ºåŠ›ã®1å›žåˆ¶é™ã¨ã¯åˆ†ã‘ã‚‹ã€‚
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
                $warnings += "PDFåŒ–ç’°å¢ƒãŒå‰å›žã¨ç•°ãªã‚Šã¾ã™: $key å‰å›ž=[$($old.$key)] ä»Šå›ž=[$($EnvInfo.$key)]"
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
        throw 'PDFä½œæˆä¸­ã¯ReportBinderã‚’çµ‚äº†ã§ãã¾ã›ã‚“ã€‚å‡¦ç†ãŒå®Œäº†ã™ã‚‹ã‹ã€PDFä½œæˆã‚’ä¸­æ­¢ã—ã¦ã‹ã‚‰çµ‚äº†ã—ã¦ãã ã•ã„ã€‚'
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
    # V5: Invoke-WithLock ã¯ body çµ‚äº†ã§ãƒãƒ³ãƒ‰ãƒ«ã‚’é–‰ã˜ã‚‹ãŸã‚ã€tick ã‚’ã¾ãŸã„ã§ãƒ­ãƒƒã‚¯ã‚’ä¿æŒã§ããªã„ã€‚
    # è‡ªå‹•ã‚¹ã‚±ã‚¸ãƒ¥ãƒ¼ãƒ©ãƒ¼ãŒè¤‡æ•°ãƒ–ãƒƒã‚¯ã®æ‰€æœ‰æ¨©ã‚’æŒã¡ç¶šã‘ã‚‹ãŸã‚ã®å–å¾—å°‚ç”¨ãƒ˜ãƒ«ãƒ‘ã€‚
    # å–å¾—ã§ããªã„å ´åˆã¯ $null ã‚’è¿”ã™(ã‚¨ãƒ©ãƒ¼ã«ã—ãªã„ã€‚ä»–ã‚µãƒ¼ãƒãƒ¼ãŒæ‹…å½“ã—ã¦ã„ã‚‹ã ã‘)ã€‚
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
    # V5-D1: Excel COM ã¯è¨€èªžã”ã¨ã«1ã‚¸ãƒ§ãƒ–ã¸åˆ¶é™ã™ã‚‹ã€‚
    # ãƒ‡ãƒƒãƒ‰ãƒ­ãƒƒã‚¯ã‚’é¿ã‘ã‚‹ãŸã‚ã€å–å¾—é †åºã‚’å…¨çµŒè·¯ã§ render-engine -> render_<workbookId> ã«çµ±ä¸€ã™ã‚‹ã€‚
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
    # OpenOrCreate + FileShare.None: æŽ’ä»–ã¯ã€Œé–‹ã„ã¦ã„ã‚‹ãƒãƒ³ãƒ‰ãƒ«ã€ã§æ‹…ä¿ã™ã‚‹ã€‚
    # ä»¥å‰ã® CreateNew æ–¹å¼ã¯ã€ãƒ—ãƒ­ã‚»ã‚¹å¼·åˆ¶çµ‚äº†ã‚„é›»æºæ–­ã§ãƒ­ãƒƒã‚¯ãƒ•ã‚¡ã‚¤ãƒ«ãŒæ®‹ã‚‹ã¨
    # æ‰‹å‹•å‰Šé™¤ã™ã‚‹ã¾ã§æ°¸ä¹…ã«ã€Œä»–ã®å‡¦ç†ãŒãƒ­ãƒƒã‚¯ä¸­ã€ã«ãªã£ã¦ã„ãŸã€‚
    # å¼·åˆ¶çµ‚äº†ç›´å¾Œã¯SMBå´ã§ãƒãƒ³ãƒ‰ãƒ«è§£æ”¾ãŒé…ã‚Œã‚‹ã“ã¨ãŒã‚ã‚‹ãŸã‚ã€çŸ­ã„ãƒªãƒˆãƒ©ã‚¤ã‚’å…¥ã‚Œã‚‹ã€‚
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
        throw "ä»–ã®å‡¦ç†ãŒãƒ­ãƒƒã‚¯ä¸­ã§ã™ã€‚èª°ã‹ãŒåŒã˜å‡¦ç†ã‚’å®Ÿè¡Œä¸­ã‹ã€å¼·åˆ¶çµ‚äº†ã—ãŸãƒ—ãƒ­ã‚»ã‚¹ã®ãƒãƒ³ãƒ‰ãƒ«ãŒæ®‹ã£ã¦ã„ã¾ã™ã€‚æ™‚é–“ã‚’ãŠã„ã¦å†å®Ÿè¡Œã—ã¦ãã ã•ã„ã€‚è§£æ”¾å¾Œã«ä¿æŒè€…ã‚’ç¢ºèªã™ã‚‹ã«ã¯æ¬¡ã®ãƒ•ã‚¡ã‚¤ãƒ«ã‚’é–‹ã„ã¦ãã ã•ã„: $LockPath"
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
    if (-not (Test-Path -LiteralPath $full)) { throw "æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $RelativePath" }
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
            foreach ($name in @('lastRenderedVersionId','lastRenderedExcelHash','lastRenderedAt','lastRenderedSheets','lastRenderedSheetFingerprint','lastRenderLog','renderProfileVersion','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash')) {
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
    if (-not (Test-Path -LiteralPath $full)) { throw "æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $RelativePath" }
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
    if (-not (Test-Path -LiteralPath $full)) { throw "æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $RelativePath" }
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
    if (-not (Test-Path -LiteralPath $worker -PathType Leaf)) { throw 'Wordå¤‰æ›ãƒ¯ãƒ¼ã‚«ãƒ¼ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }
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
        if ($found.Count -eq 0) { throw "ç™»éŒ²æ¸ˆã¿WordåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $SourceId" }
        $source = $found[0]
        if ([string](Get-DataProperty $source 'sourceType' '') -ne 'word') { throw 'WordåŽŸç¨¿ã§ã¯ãªã„ãŸã‚Wordã‚¢ãƒ€ãƒ—ã‚¿ãƒ¼ã§å‡¦ç†ã§ãã¾ã›ã‚“ã€‚' }
        $livePath = Join-Safe ([string]$paths.submissionDir) ([string]$source.relativePath)
        if (-not (Test-Path -LiteralPath $livePath)) { throw "WordåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $($source.relativePath)" }
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
        if (-not [string]::IsNullOrWhiteSpace($expectedHash) -and $sourceHash -ne $expectedHash) { throw 'æŒ‡å®šã—ãŸWordåŽŸç¨¿ã®ç‰ˆã¨å®Ÿãƒ•ã‚¡ã‚¤ãƒ«ãŒä¸€è‡´ã—ã¾ã›ã‚“ã€‚æ›´æ–°ç¢ºèªå¾Œã«å†åº¦å¤‰æ›ã—ã¦ãã ã•ã„ã€‚' }
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
                if ($latest.Count -eq 0) { throw "ç™»éŒ²æ¸ˆã¿WordåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $SourceId" }
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
                Mark-VolumeNeedsRebuild $st $Language $cat $vols 'render' 'WordåŽŸç¨¿ã‚’1ä»¶PDFå¤‰æ›ã—ã¾ã—ãŸ'
                return $sync
            }
            $wordVersion = [string](Get-DataProperty $wordResult 'wordVersion' '')
            $Script:CurrentRenderEnvFingerprint = Get-Sha256Text "word-com|$wordVersion|$($Script:WordRenderProfileVersion)"
            $Script:CurrentRenderEnvInfo = [ordered]@{ adapterId = 'word-com-v1'; wordVersion = $wordVersion; profileVersion = $Script:WordRenderProfileVersion }
            $Script:PendingAnalysis = [ordered]@{ language = $Language; workbookId = $SourceId; snapshotId = [string]$SourceSnapshotId; versionId = $versionId; rendered = $rendered }
            Remove-WorkbookContentPdfs $workspace $SourceId $versionId
            $success = $true
            return [ordered]@{ workbookId = $SourceId; sourceId = $SourceId; sourceType = 'word'; adapterId = 'word-com-v1'; versionId = $versionId; rendered = @($rendered); warnings = @(); steps = @('Wordã‚’ä¸€æ™‚ã‚³ãƒ”ãƒ¼ã‹ã‚‰PDFå¤‰æ›','PDFã‚’ãƒšãƒ¼ã‚¸å˜ä½ã§å–ã‚Šè¾¼ã¿'); sheetSync = $pageSync }
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
        if ($found.Count -eq 0) { throw "ç™»éŒ²æ¸ˆã¿PDFåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $SourceId" }
        $source = $found[0]
        if ([string](Get-DataProperty $source 'sourceType' '') -ne 'pdf') { throw 'PDFåŽŸç¨¿ã§ã¯ãªã„ãŸã‚PDFã‚¢ãƒ€ãƒ—ã‚¿ãƒ¼ã§å‡¦ç†ã§ãã¾ã›ã‚“ã€‚' }
        $livePath = Join-Safe ([string]$paths.submissionDir) ([string]$source.relativePath)
        if (-not (Test-Path -LiteralPath $livePath)) { throw "PDFåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $($source.relativePath)" }
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
        if (-not [string]::IsNullOrWhiteSpace($expectedHash) -and $sourceHash -ne $expectedHash) { throw 'æŒ‡å®šã—ãŸPDFåŽŸç¨¿ã®ç‰ˆã¨å®Ÿãƒ•ã‚¡ã‚¤ãƒ«ãŒä¸€è‡´ã—ã¾ã›ã‚“ã€‚æ›´æ–°ç¢ºèªå¾Œã«å†åº¦å–ã‚Šè¾¼ã‚“ã§ãã ã•ã„ã€‚' }
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
                if ($latest.Count -eq 0) { throw "ç™»éŒ²æ¸ˆã¿PDFåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $SourceId" }
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
                Mark-VolumeNeedsRebuild $st $Language $cat $vols 'render' 'PDFåŽŸç¨¿ã‚’1ä»¶å–ã‚Šè¾¼ã¿ã¾ã—ãŸ'
                return $sync
            }
            $Script:CurrentRenderEnvFingerprint = Get-Sha256Text "pdfbox-import|$($Script:PdfImportProfileVersion)"
            $Script:CurrentRenderEnvInfo = [ordered]@{ adapterId = 'pdfbox-import-v1'; profileVersion = $Script:PdfImportProfileVersion }
            $Script:PendingAnalysis = [ordered]@{ language = $Language; workbookId = $SourceId; snapshotId = [string]$SourceSnapshotId; versionId = $versionId; rendered = $rendered }
            Remove-WorkbookContentPdfs $workspace $SourceId $versionId
            $success = $true
            return [ordered]@{ workbookId = $SourceId; sourceId = $SourceId; sourceType = 'pdf'; adapterId = 'pdfbox-import-v1'; versionId = $versionId; rendered = @($rendered); warnings = @(); steps = @('PDFã‚’ãƒšãƒ¼ã‚¸å˜ä½ã§å–ã‚Šè¾¼ã¿'); sheetSync = $pageSync }
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
            if ([string]::IsNullOrWhiteSpace($snapshotLease) -or [string]::IsNullOrWhiteSpace($contentLease)) { throw 'æ¯”è¼ƒè³‡ç”£ã®ä¿è­·leaseã‚’ä½œæˆã§ãã¾ã›ã‚“ã§ã—ãŸã€‚' }
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
            if ([string]::IsNullOrWhiteSpace($snapshotLease) -or [string]::IsNullOrWhiteSpace($contentLease)) { throw 'æ¯”è¼ƒè³‡ç”£ã®ä¿è­·leaseã‚’ä½œæˆã§ãã¾ã›ã‚“ã§ã—ãŸã€‚' }
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
        Update-StructureLocked $Language {
            param($structure)
            $wb=@(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
            if ($wb.Count -gt 0) {
                # V5-Â§6.3a: å¤±æ•—ã—ãŸã®ã¯ã€Œè©¦è¡Œã—ãŸç‰ˆã€ã§ã‚ã£ã¦ã€ç¾åœ¨ã®ç‰ˆã¨ã¯é™ã‚‰ãªã„ã€‚
                # æ—¢ã«æ–°ã—ã„ç‰ˆãŒæ¤œçŸ¥ã•ã‚Œã¦ã„ã‚‹ãªã‚‰ render-error ã«ã›ãš excel-updated ã®ã¾ã¾ã«ã™ã‚‹ã€‚
                $attempted = Normalize-FileHash $AttemptedHash
                $latest = Normalize-FileHash ([string](Get-DataProperty $wb[0] 'currentExcelHash' ''))
                $sameVersion = ([string]::IsNullOrWhiteSpace($attempted)) -or ([string]::IsNullOrWhiteSpace($latest)) -or ($attempted -eq $latest)
                if ($sameVersion) {
                    Set-NoteProperty $wb[0] 'status' 'render-error'
                    foreach ($p in @(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId -and [string]$_.status -ne 'confirmed' })) { Set-NoteProperty $p 'status' 'render-error'; Set-NoteProperty $p 'warnings' @($userMessage); Set-NoteProperty $p 'updatedAt' (New-NowIso) }
                } else {
                    Set-NoteProperty $wb[0] 'status' (Get-SourceUpdatedStatus $wb[0])
                    # ãƒšãƒ¼ã‚¸ã® status ã¯ä¸Šæ›¸ãã—ãªã„(æ–°ã—ã„ç‰ˆã®çŠ¶æ…‹ã‚’å£Šã•ãªã„ãŸã‚)
                }
                # å¤±æ•—æƒ…å ±ã¯çŠ¶æ…‹ã¨åˆ‡ã‚Šé›¢ã—ã¦å¿…ãšæ®‹ã™
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
        # ä¿æŒæ•´ç†ã®ç«¶åˆãƒ»å¤±æ•—ã§ã€å®Œæˆæ¸ˆã¿PDFä½œæˆãã®ã‚‚ã®ã‚’å¤±æ•—æ‰±ã„ã«ã—ãªã„ã€‚
        Write-Warning ('content PDFã®ä¸–ä»£æ•´ç†ã‚’è¦‹é€ã‚Šã¾ã—ãŸ: ' + $_.Exception.Message)
    }
}

function Remove-WorkbookContentPdfsCore([string]$Workspace, [string]$WorkbookId, [string]$KeepVersionId) {
    # V5-Â§3.7: ä¿æŒä¸–ä»£æ•°ã¨ pins/leases ã«ã‚ˆã‚‹ä¿è­·ã‚’å°Šé‡ã™ã‚‹ã€‚
    # å˜ç´”ã« KeepVersionId ä»¥å¤–ã‚’å…¨å‰Šé™¤ã™ã‚‹ã¨ã€å·®åˆ†æ¯”è¼ƒã®åŸºæº–ã‚„æ­£å¼PDFãŒå‚ç…§ã™ã‚‹ä¸–ä»£ã¾ã§æ¶ˆãˆã‚‹ã€‚
    try {
        # V5-P0: æ‰¿èªå‰ã¯ãƒ‡ã‚£ã‚¹ã‚¯ä½¿ç”¨é‡ã‚’å¢—ã‚„ã•ãªã„ã€‚V4.1 ã¨åŒã˜ã€Œæœ€æ–°ä»¥å¤–ã‚’å‰Šé™¤ã€ã«æˆ»ã™ã€‚
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
    if ([string]::IsNullOrWhiteSpace($WorkbookId)) { throw 'workbookId ãŒå¿…è¦ã§ã™ã€‚' }
    $result=Update-StructureLocked $Language {
        param($structure)
        $found=@(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -eq $WorkbookId } | Select-Object -First 1)
        if ($found.Count -eq 0) { throw "ç™»éŒ²æ¸ˆã¿ExcelãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $WorkbookId" }
        $cat=Get-WorkbookPackId $found[0]
        $removedPages=@(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -eq $WorkbookId })
        $affected=@($removedPages | ForEach-Object {[string]$_.volume} | Where-Object {$_ -and $_ -ne 'none'} | Select-Object -Unique)
        # å–ã‚Šæ¶ˆã›ã‚‹å ´æ‰€ã‚’ä½œã£ã¦ã‹ã‚‰æ¶ˆã™ã€‚ãƒšãƒ¼ã‚¸ã®æŒ¯ã‚Šåˆ†ã‘ãƒ»ä¸¦ã³é †ãƒ»ãƒšãƒ¼ã‚¸åãŒã¾ã¨ã‚ã¦
        # å¤±ã‚ã‚Œã‚‹æ“ä½œãªã®ã«ã€ã“ã“ã ã‘ã‚¹ãƒŠãƒƒãƒ—ã‚·ãƒ§ãƒƒãƒˆã‚’å–ã£ã¦ã„ãªã‹ã£ãŸ(ä»–ã®ç ´å£Šçš„ãª
        # æ“ä½œã¯ã™ã¹ã¦å–ã£ã¦ã„ã‚‹)ã€‚å¾©å…ƒãƒ—ãƒ¬ãƒ“ãƒ¥ãƒ¼ã¯ã€ŒéŽåŽ»ã«ã ã‘å­˜åœ¨ã™ã‚‹ãƒšãƒ¼ã‚¸ã€ã‚’
        # é©ç”¨ã—ãªã„ã®ã§ã€ã“ã‚ŒãŒç„¡ã„ã¨æˆ»ã™æ‰‹æ®µãŒä¸€ã¤ã‚‚ç„¡ã„ã€‚
        if ($removedPages.Count -gt 0) { [void](Save-LayoutSnapshot $Language $cat 'pre-unregister' $structure) }
        $structure.workbooks=@(Get-Array $structure.workbooks | Where-Object { [string]$_.workbookId -ne $WorkbookId })
        $structure.pages=@(Get-Array $structure.pages | Where-Object { [string]$_.workbookId -ne $WorkbookId })
        Mark-VolumeNeedsRebuild $structure $Language $cat $affected 'unregister' 'Excelã‚’1ä»¶ç™»éŒ²è§£é™¤ã—ã¾ã—ãŸ'
        return [ordered]@{workbookId=$WorkbookId;fileName=[string]$found[0].fileName;removedPages=$removedPages}
    }
    # V5-Â§3.8: ç™»éŒ²è§£é™¤ã¨å±¥æ­´å‰Šé™¤ã¯åˆ¥æ“ä½œã€‚å±¥æ­´ãŒæœ‰åŠ¹ãªã‚‰ content-pdf ã‚‚æ®‹ã™
    # (æ­£å¼PDFã‚¢ãƒ¼ã‚«ã‚¤ãƒ–ãŒéŽåŽ»ã®ä¸–ä»£ã‚’å‚ç…§ã—ã¦ã„ã‚‹å¯èƒ½æ€§ãŒã‚ã‚‹ãŸã‚)ã€‚
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
    if ($workbook.Count -eq 0) { throw "WorkbookãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $WorkbookId" }
    $wb = $workbook[0]
    $sourcePath = Join-Safe ([string]$paths.submissionDir) ([string]$wb.relativePath)
    # V5: ä¿å­˜æ¸ˆã¿æ¤œçŸ¥ç‰ˆã‹ã‚‰ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ã™ã‚‹å ´åˆã¯ã€æå‡ºãƒ•ã‚©ãƒ«ãƒ€ã®ç¾ç‰©ãŒæ¶ˆãˆã¦ã„ã¦ã‚‚ç¶šè¡Œã§ãã‚‹ã€‚
    if ([string]::IsNullOrWhiteSpace($SourceOverridePath) -and -not (Test-Path -LiteralPath $sourcePath)) {
        $hasSnapshotInput = $false
        try { $probe = Capture-RenderInput $Language $WorkbookId $SourceSnapshotId ''; $hasSnapshotInput = (-not [string]::IsNullOrWhiteSpace([string]$probe.path)) } catch { }
        if (-not $hasSnapshotInput) {
            Update-StructureLocked $Language { param($st) $x=@(Get-Array $st.workbooks|Where-Object{[string]$_.workbookId -eq $WorkbookId}|Select-Object -First 1);if($x.Count){Set-NoteProperty $x[0] 'status' 'missing'} } | Out-Null
            throw "æå‡ºãƒ•ã‚¡ã‚¤ãƒ«ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $($wb.relativePath)"
        }
    }
    # V5-D1: render-engine -> render_<workbookId> ã®é †ã§å–å¾—ã™ã‚‹ã€‚
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

        # èª°ã‹ãŒä¿å­˜ã—ãŸç›´å¾Œã‚„ã‚¦ã‚¤ãƒ«ã‚¹ã‚¹ã‚­ãƒ£ãƒ³ä¸­ã¯èª­ã¿å–ã‚ŠãŒä¸€æ™‚çš„ã«å¤±æ•—ã™ã‚‹ãŸã‚ã€å°‘ã—å¾…ã£ã¦ãƒªãƒˆãƒ©ã‚¤ã™ã‚‹ã€‚
        # V5-Â§6.1/Â§C: å…ˆã«ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°å…¥åŠ›ã‚’ç¢ºå®šã•ã›ã‚‹ã€‚
        # ãƒãƒƒã‚·ãƒ¥ã¯ã€Œå®Ÿéš›ã«é–‹ããƒ•ã‚¡ã‚¤ãƒ«ã€ã«å¯¾ã—ã¦è¨ˆç®—ã—ãªã‘ã‚Œã°ãªã‚‰ãªã„ã€‚
        # ç¾è¡ŒExcelã®ãƒãƒƒã‚·ãƒ¥ã‚’ä½¿ã†ã¨ã€æ¤œçŸ¥ç‰ˆã‹ã‚‰ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ã—ãŸã®ã«æœ€æ–°æ‰±ã„ã«ãªã‚ŠCASãŒå£Šã‚Œã‚‹ã€‚
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
            if (-not (Test-Path -LiteralPath $SourceOverridePath)) { throw "ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°å…¥åŠ›ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $SourceOverridePath" }
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
        # V5-Â§6.3a: catch çµŒè·¯ãŒã€Œã©ã®ç‰ˆã‚’è©¦è¡Œã—ãŸã‹ã€ã‚’çŸ¥ã‚‹ãŸã‚ã«è¨˜éŒ²ã™ã‚‹ã€‚
        # åŒä¸€ãƒ—ãƒ­ã‚»ã‚¹å†…ã®ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ã¯ render-engine ãƒ­ãƒƒã‚¯ã§ç›´åˆ—åŒ–ã•ã‚Œã¦ã„ã‚‹ãŸã‚å®‰å…¨ã€‚
        # V5-P0(#3): å›ºå®šç‰ˆ(pin)ã®æŒ‡å®šãŒã‚ã‚‹å ´åˆã€è©¦è¡Œã—ãŸç‰ˆã¯ã€Œæ„å›³ã—ãŸ ExpectedSourceHashã€ã§ã‚ã£ã¦
        # ç¾ç‰©ãƒ•ã‚¡ã‚¤ãƒ«($sourceHash=ç¾åœ¨ç‰ˆ)ã§ã¯ãªã„ã€‚ç¾ç‰©ã®ãƒãƒƒã‚·ãƒ¥ã‚’è¨˜éŒ²ã™ã‚‹ã¨ã€ç‰ˆãšã‚Œå¤±æ•—æ™‚ã«
        # Set-WorkbookRenderError ã® sameVersion åˆ¤å®šãŒèª¤ã£ã¦ç¾åœ¨ç‰ˆã‚’ render-error åŒ–ã—ã¦ã—ã¾ã†ã€‚
        $attemptHashForRecord = if (-not [string]::IsNullOrWhiteSpace($ExpectedSourceHash)) { [string]$ExpectedSourceHash } else { [string]$sourceHash }
        $Script:LastRenderAttempt = [ordered]@{ workbookId = $WorkbookId; snapshotId = [string]$SourceSnapshotId; hash = [string]$attemptHashForRecord }
        if (-not [string]::IsNullOrWhiteSpace($ExpectedSourceHash) -and -not [string]::IsNullOrWhiteSpace($sourceHash)) {
            if ((Normalize-FileHash $sourceHash) -ne (Normalize-FileHash $ExpectedSourceHash)) {
                throw 'ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°å…¥åŠ›ãŒæƒ³å®šã—ãŸç‰ˆã¨ä¸€è‡´ã—ã¾ã›ã‚“ã€‚ã‚‚ã†ä¸€åº¦PDFä½œæˆã—ã¦ãã ã•ã„ã€‚'
            }
        }
        if ([string]::IsNullOrWhiteSpace($sourceHash)) {
            throw "æå‡ºExcelã‚’èª­ã¿å–ã‚Œã¾ã›ã‚“ã§ã—ãŸ(3å›žè©¦è¡Œ)ã€‚èª°ã‹ãŒä¿å­˜ä¸­ã‹ã€æŽ’ä»–ãƒ¢ãƒ¼ãƒ‰ã§é–‹ã‹ã‚Œã¦ã„ã‚‹å¯èƒ½æ€§ãŒã‚ã‚Šã¾ã™ã€‚å°‘ã—å¾…ã£ã¦ã‹ã‚‰ã‚‚ã†ä¸€åº¦PDFä½œæˆã—ã¦ãã ã•ã„ã€‚ $lastReadError"
        }
        $item = Get-Item -LiteralPath $sourcePath
        # V5: ç§’å˜ä½ã ã¨åŒä¸€ç§’ã®2å›žå‡¦ç†ã§åŒã˜ä¸–ä»£ãƒ•ã‚©ãƒ«ãƒ€ã‚’å…±æœ‰ã—ã¦ã—ã¾ã†ã€‚ãƒŸãƒªç§’+GUID8ã«ã™ã‚‹ã€‚
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
            throw "æå‡ºExcelã‚’ã‚³ãƒ”ãƒ¼ã§ãã¾ã›ã‚“ã§ã—ãŸ(3å›žè©¦è¡Œ)ã€‚èª°ã‹ãŒä¿å­˜ä¸­ã®å¯èƒ½æ€§ãŒã‚ã‚Šã¾ã™ã€‚å°‘ã—å¾…ã£ã¦ã‹ã‚‰ã‚‚ã†ä¸€åº¦PDFä½œæˆã—ã¦ãã ã•ã„ã€‚ $lastCopyError"
        }
        # Copy-Itemã¯Zone.Identifierï¼ˆMark of the Webï¼‰ã‚’ä¸€æ™‚ã‚³ãƒ”ãƒ¼ã¸å¼•ãç¶™ããŸã‚ã€
        # ä¿è­·ãƒ“ãƒ¥ãƒ¼ã«ã‚ˆã‚‹Workbooks.Openå¤±æ•—ã‚’é˜²ãç›®çš„ã§æ˜Žç¤ºçš„ã«é™¤åŽ»ã™ã‚‹ã€‚
        try { Unblock-File -LiteralPath $tmpPath -ErrorAction SilentlyContinue } catch { }
        # Full SHA-256 twice per workbook was expensive. The source has already been hashed;
        # after copying, confirm size and source timestamp/size stability instead of hashing the copy again.
        $copyItem = Get-Item -LiteralPath $tmpPath -ErrorAction Stop
        $sourceAfterCopy = Get-Item -LiteralPath $sourcePath -ErrorAction Stop
        if ($copyItem.Length -ne $item.Length -or $sourceAfterCopy.Length -ne $item.Length -or $sourceAfterCopy.LastWriteTimeUtc -ne $item.LastWriteTimeUtc) {
            throw 'ã‚³ãƒ”ãƒ¼ä¸­ã«æå‡ºExcelãŒæ›´æ–°ã•ã‚Œã¾ã—ãŸã€‚ä¿å­˜ãŒçµ‚ã‚ã£ã¦ã‹ã‚‰ã‚‚ã†ä¸€åº¦PDFä½œæˆã—ã¦ãã ã•ã„ã€‚'
        }
        $timingsMs.localCopy = [int64]$copyTimer.ElapsedMilliseconds
        # V5-Â§3.1/Â§6.2: æå‡ºãƒ•ã‚©ãƒ«ãƒ€ã®ç¾ç‰©ã‹ã‚‰ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ã™ã‚‹å ´åˆã¯ã€
        # ãƒ˜ãƒƒãƒ€ãƒ¼/ãƒ•ãƒƒã‚¿ãƒ¼XMLã‚’åŠ å·¥ã™ã‚‹å‰(Lç›´å¾Œ)ã«æœªåŠ å·¥ã®æ¤œçŸ¥ç‰ˆã‚’ä¿å­˜ã™ã‚‹ã€‚
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
                $steps += 'Excel COMã‚’èµ·å‹•'
                $excelStartTimer = [Diagnostics.Stopwatch]::StartNew()
                $excel = New-ExcelApplicationForRender
                $timingsMs.excelStartup = [int64]$excelStartTimer.ElapsedMilliseconds
                $ownsExcel = $true
            } else {
                $steps += 'æ—¢å­˜ã®Excel COMã‚’ä½¿ç”¨'
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
                $steps += "Excelã‚’é–‹ãå‰ã«å°åˆ·è¨­å®šã‚’æº–å‚™: $([int](Get-DataProperty $excelPackagePreparation 'sheetsPrepared' 0)) ã‚·ãƒ¼ãƒˆ"
            } elseif ($excelPackagePreparation -and $excelPackagePreparation.ok -eq $false) {
                $warnings += "Excelå†…éƒ¨ã®å°åˆ·è¨­å®šã‚’äº‹å‰æº–å‚™ã§ãã¾ã›ã‚“ã§ã—ãŸã€‚COMè¨­å®šã§ç¶šè¡Œã—ã¾ã™: $([string](Get-DataProperty $excelPackagePreparation 'error' ''))"
            }
            $steps += 'ä¸€æ™‚ã‚³ãƒ”ãƒ¼ã‚’é–‹ã'
            if ($ProgressCallback) { & $ProgressCallback 'open' $WorkbookId '' }
            $excelOpenTimer = [Diagnostics.Stopwatch]::StartNew()
            $book = Open-ExcelWorkbookSafe $excel $tmpPath $true
            $timingsMs.excelOpen = [int64]$excelOpenTimer.ElapsedMilliseconds
            try { $book.CheckCompatibility = $false } catch { }

            $sheetSetupTimer = [Diagnostics.Stopwatch]::StartNew()
            $inspected = @()
            $targetSheetNames = @()
            # éžè¡¨ç¤ºã‚·ãƒ¼ãƒˆã‚‚ã€ŒåŽŸç¨¿ã«å­˜åœ¨ã™ã‚‹ã€ã¨ã—ã¦è¨˜éŒ²ã™ã‚‹ã€‚è¡¨ç¤ºä¸­ã®ã‚‚ã®ã ã‘ã‚’æ¸¡ã™ã¨ã€
            # ä¸€æ™‚çš„ã«éžè¡¨ç¤ºã«ã—ãŸã ã‘ã§ãƒšãƒ¼ã‚¸è¨­å®šãŒå‰Šé™¤ã•ã‚Œã¦ã—ã¾ã†ã€‚
            $allSheetNames = @()
            $sheetRenderInfos = @()
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
                        if ($visible) {
                            $targetSheetNames += $sheetName
                            $a1 = ''
                            try { $a1 = [string]$ws.Range('A1').Text } catch { }
                            if ([string]::IsNullOrWhiteSpace($a1)) { $a1 = "$($wb.fileName) / $sheetName" }
                            # Reading PageSetup.PrintArea is another slow COM call and is only diagnostic.
                            # Keep it blank in render logs to avoid delaying PDFä½œæˆ.
                            $printArea = ''
                            $inspected += [ordered]@{ sheetName = $sheetName; sheetIndex = $i; titleSource = 'A1'; detectedTitle = $a1; printArea = $printArea }

                            if (-not $packagePrintSettingsPrepared) {
                                $steps += "ã‚·ãƒ¼ãƒˆ $sheetName ã®å°åˆ·è¨­å®šã‚’èª¿æ•´"
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
                $emptyPageSync = Update-WorkbookPagesFromInspection $Language $structure $wb @() $allSheetNames
                # V5-Â§3.7: æ—§ä¸–ä»£ã®å€‹åˆ¥å‰Šé™¤ã¯å»ƒæ­¢ã€‚ä¸–ä»£å˜ä½ã®æŽƒé™¤(Remove-WorkbookContentPdfs)ã«ä¸€æœ¬åŒ–ã™ã‚‹ã€‚
                # å€‹åˆ¥ã«æ¶ˆã™ã¨ã€ä¿æŒã—ã¦ã„ã‚‹ã¯ãšã®ä¸–ä»£ãƒ•ã‚©ãƒ«ãƒ€ã®ä¸­èº«ãŒæ¬ æã™ã‚‹ã€‚
                Set-WorkbookRenderedSheetSnapshot $wb @()
                Update-StructureLocked $Language { param($st) $x=@(Get-Array $st.workbooks|Where-Object{[string]$_.workbookId -eq $WorkbookId}|Select-Object -First 1);if($x.Count){[void](Update-WorkbookPagesFromInspection $Language $st $x[0] @() $allSheetNames);Set-WorkbookRenderedSheetSnapshot $x[0] @()} } | Out-Null
                throw 'PDFåŒ–å¯¾è±¡ã®è¡¨ç¤ºã‚·ãƒ¼ãƒˆãŒã‚ã‚Šã¾ã›ã‚“ã€‚Excelã§å°‘ãªãã¨ã‚‚1ã¤ã®ãƒ¯ãƒ¼ã‚¯ã‚·ãƒ¼ãƒˆã‚’è¡¨ç¤ºã—ã¦ãã ã•ã„ã€‚'
            }

            $excelExportTimer = [Diagnostics.Stopwatch]::StartNew()
            $batchResult = $null
            if (@($sheetRenderInfos).Count -gt 1) {
                if ($ProgressCallback) { & $ProgressCallback 'batch' $WorkbookId '' }
                $steps += "è¤‡æ•°ã‚·ãƒ¼ãƒˆã‚’ä¸€æ‹¬PDFåŒ–: $(@($sheetRenderInfos).Count) ã‚·ãƒ¼ãƒˆ"
                $batchResult = Export-WorkbookSheetsToPdfBatch $excel $book $sheetRenderInfos $tmpDir
                if ($batchResult -and $batchResult.ok -eq $true) {
                    $steps += 'ä¸€æ‹¬PDFã‚’ã‚·ãƒ¼ãƒˆåˆ¥PDFã¸åˆ†å‰²'
                    if ($ProgressCallback) { & $ProgressCallback 'split' $WorkbookId '' }
                } else {
                    $reason = [string](Get-DataProperty $batchResult 'reason' '')
                    $msg = [string](Get-DataProperty $batchResult 'message' '')
                    if ($reason -ne 'single-sheet' -and $reason -ne 'pdfbox-unavailable') {
                        $warnings += "ä¸€æ‹¬PDFåŒ–ã‚’ä½¿ãˆãªã‹ã£ãŸãŸã‚å¾“æ¥æ–¹å¼ã§ç¶šè¡Œã—ã¾ã™: $msg"
                    }
                }
            }

            if (-not ($batchResult -and $batchResult.ok -eq $true)) {
                foreach ($info in @($sheetRenderInfos)) {
                    $ws = $null
                    try {
                        $sheetName = [string]$info.sheetName
                        $steps += "ã‚·ãƒ¼ãƒˆ $sheetName ã‚’PDFåŒ–"
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
                if ($pageCount -gt 1) { $pageWarnings += "ã“ã®ã‚·ãƒ¼ãƒˆã®PDFã¯ $pageCount ãƒšãƒ¼ã‚¸ã§ã™ã€‚Excelã®å°åˆ·ç¯„å›²ãƒ»æ”¹ãƒšãƒ¼ã‚¸ãƒ»å€çŽ‡ã‚’ç¢ºèªã—ã¦ãã ã•ã„ã€‚" }
                $rendered += [ordered]@{ sheetName = $sheetName; pdf = $outPdf; pageCount = $pageCount; warnings = $pageWarnings }
            }
            if ($rendered.Count -eq 0) {
                throw 'PDFã‚’ä½œæˆã§ãã¾ã›ã‚“ã§ã—ãŸã€‚Excelã®å°åˆ·è¨­å®šã¾ãŸã¯å¯¾è±¡ã‚·ãƒ¼ãƒˆã‚’ç¢ºèªã—ã¦ãã ã•ã„ã€‚'
            }

            $pageSync = Update-WorkbookPagesFromInspection $Language $structure $wb $inspected $allSheetNames
            $removedPages = @(Get-Array (Get-DataProperty $pageSync 'removedPages' @()))
            if ($removedPages.Count -gt 0) {
                $removedNames = @($removedPages | ForEach-Object { [string](Get-DataProperty $_ 'sheetName' '') } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $removedLabel = [string]::Join('ã€', $removedNames)
                if ([string]::IsNullOrWhiteSpace($removedLabel)) { $removedLabel = "$($removedPages.Count) ãƒšãƒ¼ã‚¸" }
                $warnings += "ç¾åœ¨ã®Excelã«å­˜åœ¨ã—ãªã„ã‚·ãƒ¼ãƒˆã‚’ãƒšãƒ¼ã‚¸æ§‹æˆã‹ã‚‰å¤–ã—ã¾ã—ãŸ: $removedLabelï¼ˆãƒšãƒ¼ã‚¸æ§‹æˆã®å±¥æ­´ã‹ã‚‰å…ƒã«æˆ»ã›ã¾ã™ï¼‰"
                $steps += "å­˜åœ¨ã—ãªã„ã‚·ãƒ¼ãƒˆã®å¤ã„ãƒšãƒ¼ã‚¸ã‚’å‰Šé™¤: $removedLabel"
                # V5-Â§3.7: æ—§ä¸–ä»£ã®å€‹åˆ¥å‰Šé™¤ã¯å»ƒæ­¢(ä¸Šè¨˜ã¨åŒã˜ç†ç”±)ã€‚
            }
            $renamedPages = @(Get-Array (Get-DataProperty $pageSync 'renamedPages' @()))
            if ($renamedPages.Count -gt 0) {
                $renameLabel = [string]::Join('ã€', @($renamedPages | ForEach-Object {
                    "{0}â†’{1}" -f [string](Get-DataProperty $_ 'fromSheetName' ''), [string](Get-DataProperty $_ 'toSheetName' '')
                }))
                $warnings += "ã‚·ãƒ¼ãƒˆåã®å¤‰æ›´ã‚’å¼•ãç¶™ãŽã¾ã—ãŸ: $renameLabelï¼ˆé…ç½®ã¨ãƒšãƒ¼ã‚¸è¨­å®šã¯ãã®ã¾ã¾ã§ã™ï¼‰"
                $steps += "ã‚·ãƒ¼ãƒˆåå¤‰æ›´ã®å¼•ãç¶™ãŽ: $renameLabel"
            }
            $hiddenPageIds = @(Get-Array (Get-DataProperty $pageSync 'hiddenSheetPageIds' @()))
            if ($hiddenPageIds.Count -gt 0) {
                $warnings += "éžè¡¨ç¤ºã®ã‚·ãƒ¼ãƒˆãŒ $($hiddenPageIds.Count) ä»¶ã‚ã‚Šã¾ã™ã€‚ãƒšãƒ¼ã‚¸æ§‹æˆã¯ä¿æŒã—ã¦ã„ã¾ã™ãŒã€PDFã¯ä½œã‚Šç›´ã•ã‚Œã¾ã›ã‚“ã€‚æå‡ºç‰©ã«å«ã‚ã‚‹å ´åˆã¯ã‚·ãƒ¼ãƒˆã‚’å†è¡¨ç¤ºã—ã¦ã‹ã‚‰PDFã‚’ä½œæˆã—ã¦ãã ã•ã„ã€‚"
                $steps += "éžè¡¨ç¤ºã‚·ãƒ¼ãƒˆã®ãƒšãƒ¼ã‚¸ã‚’ä¿æŒ: $($hiddenPageIds.Count) ä»¶"
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
            # V5-Â§2.4: currentExcel* ã¯ Scan-Updates ã®å°‚æœ‰ã€‚ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ã¯æ›¸ã‹ãªã„ã€‚
            # ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ä¸­ã«åˆ¥ãƒ—ãƒ­ã‚»ã‚¹ãŒæ¤œçŸ¥ã—ãŸæ–°ã—ã„ç‰ˆã‚’ã€å¤ã„ç‰ˆã®ãƒãƒƒã‚·ãƒ¥ã§ä¸Šæ›¸ãã—ãªã„ãŸã‚ã€‚
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
                if(-not $latest.Count){throw "WorkbookãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“: $WorkbookId"}
                $lw=$latest[0];$sync=Update-WorkbookPagesFromInspection $Language $st $lw $inspected $allSheetNames
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
                        Set-NoteProperty $pg[0] 'updatedAt' (New-NowIso)
                    }
                }
                # V5-Â§2.4: currentExcel* ã¯ã‚³ãƒ”ãƒ¼ã—ãªã„(Scan-Updates ã®å°‚æœ‰)ã€‚
                foreach($name in @('lastRenderedVersionId','lastRenderedExcelHash','lastRenderedAt','renderProfileVersion','lastError','lastErrorUser','lastErrorAt','lastRenderAttemptHash','warnings','lastRenderLog')){Set-NoteProperty $lw $name (Get-DataProperty $wb $name $null)}
                Set-NoteProperty $lw 'lastRenderedSnapshotId' ([string]$SourceSnapshotId)
                Set-NoteProperty $lw 'renderEnvironmentFingerprint' ([string]$Script:CurrentRenderEnvFingerprint)
                # V5-Â§6.3: ãƒ­ãƒƒã‚¯å†…ã§æœ€æ–°ã® currentExcelHash ã¨çªãåˆã‚ã›ã¦ status ã‚’æ±ºã‚ã‚‹ã€‚
                # ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ä¸­ã«å…ƒExcelãŒæ›´æ–°ã•ã‚Œã¦ã„ã‚Œã°ã€PDFã¯ä¿å­˜ã—ã¤ã¤ excel-updated ã«æˆ»ã™ã€‚
                $latestCurrentHash = Normalize-FileHash ([string](Get-DataProperty $lw 'currentExcelHash' ''))
                $renderedHash = Normalize-FileHash $sourceHash
                if ([string]::IsNullOrWhiteSpace($latestCurrentHash) -or $latestCurrentHash -eq $renderedHash) {
                    Set-NoteProperty $lw 'status' 'rendered-unchecked'
                } else {
                    Set-NoteProperty $lw 'status' 'excel-updated'
                    foreach($pg2 in @(Get-Array $st.pages|Where-Object{[string]$_.workbookId -eq $WorkbookId -and [string]$_.status -eq 'rendered'})){ Set-NoteProperty $pg2 'status' 'stale' }
                }
                Set-WorkbookRenderedSheetSnapshot $lw (Get-DataProperty $sync 'sheetNames' @())
                $cat=Get-WorkbookPackId $lw;$vols=@(Get-Array $st.pages|Where-Object{[string]$_.workbookId -eq $WorkbookId}|ForEach-Object{[string]$_.volume}|Where-Object{$_ -and $_ -ne 'none'}|Select-Object -Unique);Mark-VolumeNeedsRebuild $st $Language $cat $vols 'render' 'Excelã‚’1ä»¶PDFä½œæˆã—ã¾ã—ãŸ'
                return $sync
            }
            $timingsMs.total = [int64]$renderTimer.ElapsedMilliseconds
            Write-JsonFile (Join-Path $workspace $logRel) ([ordered]@{ workbookId = $WorkbookId; rendered = $rendered; warnings = $warnings; steps = $steps; timingsMs = $timingsMs; sheetSync = $pageSync; at = New-NowIso })
            # V5-P1: è§£æžã¯ã“ã“ã§ã¯å®Ÿè¡Œã—ãªã„ã€‚
            # ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ç”¨Excelã‚’é–‹ã„ãŸã¾ã¾ã€ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ãƒ­ãƒƒã‚¯ã‚’ä¿æŒã—ãŸã¾ã¾è§£æžã™ã‚‹ã¨ã€
            # æ¯”è¼ƒç”¨ã®å†ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ãŒ2ã¤ç›®ã®Excel COMã‚’èµ·å‹•ã—ã¦ã—ã¾ã†(1ã‚¸ãƒ§ãƒ–åˆ¶é™ã«åã™ã‚‹)ã€‚
            # ãƒ­ãƒƒã‚¯è§£æ”¾å¾Œã«å®Ÿè¡Œã™ã‚‹ãŸã‚ã€å¯¾è±¡ã ã‘ã‚’è¨˜éŒ²ã—ã¦ãŠãã€‚
            $Script:PendingAnalysis = [ordered]@{ language = $Language; workbookId = $WorkbookId; snapshotId = [string]$SourceSnapshotId; versionId = $versionId; rendered = $rendered }
            Remove-WorkbookContentPdfs $workspace $WorkbookId $versionId
        } finally {
            if ($book) { try { $book.Close($false) } catch { } ; Invoke-ComRelease $book }
            if ($ownsExcel -and $excel) { Close-ExcelApplicationForRender $excel }
            if (Test-Path -LiteralPath $tmpDir) { Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
            if (-not [string]::IsNullOrWhiteSpace($ephemeralCaptureId)) { Remove-EphemeralCopy $Language $ephemeralCaptureId }
            if ($ownsExcel -or -not $KeepExcelOpen) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
        }
        return [ordered]@{ workbookId = $WorkbookId; versionId = $versionId; rendered = $rendered; warnings = $warnings; steps = $steps; timingsMs = $timingsMs; sheetSync = $pageSync }
    }

    # V5-P1: ã“ã“ã§ã¯ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ãƒ­ãƒƒã‚¯ã‚‚Excelã‚‚è§£æ”¾æ¸ˆã¿ã€‚
    # æ¯”è¼ƒç”¨ã®å†ãƒ¬ãƒ³ãƒ€ãƒªãƒ³ã‚°ãŒå¿…è¦ã«ãªã£ã¦ã‚‚ã€æ”¹ã‚ã¦å…±é€šãƒ­ãƒƒã‚¯ã‚’å–ã‚Šç›´ã›ã‚‹ã€‚
    # è§£æžã®å¤±æ•—ã¯PDFä½œæˆã®å¤±æ•—ã«ã—ãªã„(åˆ¤å®šã¯ unknown ã«ãªã‚‹)ã€‚
    # V5-P3: KeepExcelOpen ã®ã¨ã($Script:PendingAnalysis ã‚’å‘¼å‡ºå…ƒãŒå›žåŽã™ã‚‹ä¸€æ‹¬ã‚¸ãƒ§ãƒ–)ã¯
    # ã“ã“ã§æ¶ˆã—ã¦ã¯ãªã‚‰ãªã„ã€‚ä»¥å‰ã¯ç„¡æ¡ä»¶ã« $null ã‚’ä»£å…¥ã—ã¦ã„ãŸãŸã‚ã€
    # Invoke-RenderJobFromFile ã® $deferredAnalyses ãŒå¸¸ã«ç©ºã«ãªã‚Šã€
    # ç”»åƒãƒãƒƒã‚·ãƒ¥ã®è§£æžãŒä¸€åº¦ã‚‚å®Ÿè¡Œã•ã‚Œã¦ã„ãªã‹ã£ãŸ(renders ãƒ•ã‚©ãƒ«ãƒ€ãŒä½œã‚‰ã‚Œãªã„)ã€‚
    if ($KeepExcelOpen) { return $renderResult }
    $pending = $Script:PendingAnalysis
    $Script:PendingAnalysis = $null
    if ($null -ne $pending) {
        try { [void](Invoke-PostRenderAnalysis ([string]$pending.language) ([string]$pending.workbookId) ([string]$pending.snapshotId) ([string]$pending.versionId) $pending.rendered) }
        catch { Write-Warning ('ç”»åƒãƒãƒƒã‚·ãƒ¥ã®è§£æžã«å¤±æ•—ã—ã¾ã—ãŸ: ' + $_.Exception.Message) }
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
                    Mark-VolumeNeedsRebuild $st $Language $cat $vols 'source-updated' 'å…ƒåŽŸç¨¿ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“'
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
            # ãƒãƒƒã‚·ãƒ¥ã‚’å–å¾—ã§ããªã‹ã£ãŸå›žã¯ã€æ›´æ–°æ™‚åˆ»ã¨ã‚µã‚¤ã‚ºã‚‚æ®ãˆç½®ãã€‚
            # æ–°ã—ã„ãƒ¡ã‚¿ãƒ‡ãƒ¼ã‚¿ã¨å¤ã„ãƒãƒƒã‚·ãƒ¥ã‚’çµ„ã«ã—ã¦ä¿å­˜ã™ã‚‹ã¨ã€æ¬¡å›žã®ã‚¹ã‚­ãƒ£ãƒ³ãŒ
            # ã€Œãƒ¡ã‚¿ãƒ‡ãƒ¼ã‚¿ä¸€è‡´ã€ã§å†ãƒãƒƒã‚·ãƒ¥ã‚’çœç•¥ã—ã€æ›´æ–°æ¸ˆã¿ã®åŽŸç¨¿ã‚’æœ€æ–°ã¨èª¤åˆ¤å®šã™ã‚‹ã€‚
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
                    Mark-VolumeNeedsRebuild $st $Language $cat $vols 'source-updated' 'å…ƒåŽŸç¨¿ãŒ1ä»¶æ›´æ–°ã•ã‚Œã¾ã—ãŸ'
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
        if ($status -eq 'render-error') {
            # The PDF button is an explicit retry. A previous render-error must not make
            # the button look idle just because the same file hash already failed once.
            $needs = $true
        } else {
            if ([string]::IsNullOrWhiteSpace($lastRenderedHash)) { $needs = $true }
            if ($status -in @('new','excel-updated','source-updated')) { $needs = $true }
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
        Reset-RenderEnvironmentForJob   # V5-Â§6.4: ã‚¸ãƒ§ãƒ–é–‹å§‹ã”ã¨ã«ç’°å¢ƒã‚’å–ã‚Šç›´ã™
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
    if ([string]::IsNullOrWhiteSpace($normalizedJobId)) { throw 'PDFä½œæˆã‚¸ãƒ§ãƒ–ã®æƒ…å ±ã‚’å—ã‘å–ã‚Œã¾ã›ã‚“ã§ã—ãŸã€‚' }
    return (Join-Path (Get-RenderJobDir $Language) "$normalizedJobId.cancel.json")
}

function Test-RenderJobCancellationRequested([string]$Language, [string]$JobId) {
    try { return (Test-Path -LiteralPath (Get-RenderJobCancellationPath $Language $JobId) -PathType Leaf) } catch { return $false }
}

function Request-RenderJobCancellation([string]$Language, [string]$JobId) {
    $normalizedJobId = Normalize-RenderJobId $JobId
    if ([string]::IsNullOrWhiteSpace($normalizedJobId)) { throw 'ä¸­æ­¢ã™ã‚‹PDFä½œæˆã‚¸ãƒ§ãƒ–ã‚’ç¢ºèªã§ãã¾ã›ã‚“ã§ã—ãŸã€‚' }
    $jobDir = Get-RenderJobDir $Language
    $statusPath = Join-Path $jobDir "$normalizedJobId.status.json"
    $status = Read-JsonFile $statusPath $null
    if ($null -eq $status) { throw 'ä¸­æ­¢ã™ã‚‹PDFä½œæˆã‚¸ãƒ§ãƒ–ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }
    $terminal = @('completed','completed-with-errors','failed','missing','cancelled')
    $statusText = ([string](Get-DataProperty $status 'status' '')).ToLowerInvariant()
    if ($terminal -contains $statusText) {
        return [pscustomobject][ordered]@{ ok=$true; jobId=$normalizedJobId; accepted=$false; alreadyFinished=$true; status=$statusText; message='PDFä½œæˆã¯ã™ã§ã«çµ‚äº†ã—ã¦ã„ã¾ã™ã€‚' }
    }
    $requestedAt = New-NowIso
    Write-JsonFile (Get-RenderJobCancellationPath $Language $normalizedJobId) ([ordered]@{ schemaVersion=1; jobId=$normalizedJobId; requestedAt=$requestedAt })
    return [pscustomobject][ordered]@{ ok=$true; jobId=$normalizedJobId; accepted=$true; alreadyFinished=$false; status=$statusText; cancelRequested=$true; requestedAt=$requestedAt; message='ä¸­æ­¢ã‚’å—ã‘ä»˜ã‘ã¾ã—ãŸã€‚ç¾åœ¨å‡¦ç†ä¸­ã®åŽŸç¨¿ãŒçµ‚ã‚ã‚‹ã¨åœæ­¢ã—ã¾ã™ã€‚' }
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
    $summary = if ($processed -gt 0) { "$processed / $total ä»¶ã‚’å‡¦ç†ã—ãŸæ™‚ç‚¹ã§PDFä½œæˆã‚’ä¸­æ­¢ã—ã¾ã—ãŸã€‚" } else { 'PDFä½œæˆã‚’é–‹å§‹å‰ã«ä¸­æ­¢ã—ã¾ã—ãŸã€‚' }
    Set-NoteProperty $Status 'message' $summary
    Write-RenderJobStatus $StatusPath $Status
    return $Status
}

function Read-RenderJobStatus([string]$Language, [string]$JobId) {
    $normalizedJobId = Normalize-RenderJobId $JobId
    if ([string]::IsNullOrWhiteSpace($normalizedJobId)) {
        throw 'PDFä½œæˆã‚¸ãƒ§ãƒ–ã®æƒ…å ±ã‚’å—ã‘å–ã‚Œã¾ã›ã‚“ã§ã—ãŸã€‚ã‚‚ã†ä¸€åº¦ã€ŒPDFä½œæˆã€ã‚’æŠ¼ã—ã¦ãã ã•ã„ã€‚'
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
                                    $job = Set-RenderJobFailedFromStartupProblem $path $job 'PDFä½œæˆãƒ—ãƒ­ã‚»ã‚¹ãŒçµ‚äº†ã—ã¦ã„ãŸãŸã‚åœæ­¢ã—ã¾ã—ãŸã€‚ã‚‚ã†ä¸€åº¦PDFä½œæˆã‚’æŠ¼ã—ã¦ãã ã•ã„ã€‚'
                                } elseif ($alive -and ($statusText -eq 'queued' -or $statusText -eq 'launching') -and $ageSeconds -gt 90) {
                                    # A healthy child process rewrites queued/launching -> running. Give PowerShell/Excel startup
                                    # enough room, but do not leave the UI at the initial percent forever.
                                    $job = Set-RenderJobFailedFromStartupProblem $path $job 'PDFä½œæˆãƒ—ãƒ­ã‚»ã‚¹ã¯èµ·å‹•ã—ã¾ã—ãŸãŒã€ã‚¸ãƒ§ãƒ–å‡¦ç†ã«å…¥ã‚Œã¾ã›ã‚“ã§ã—ãŸã€‚Excelã‚’é–‰ã˜ã¦ã‹ã‚‰ã€ã‚‚ã†ä¸€åº¦PDFä½œæˆã‚’æŠ¼ã—ã¦ãã ã•ã„ã€‚'
                                }
                            } elseif ($ageSeconds -gt 10) {
                                $job = Set-RenderJobFailedFromStartupProblem $path $job 'PDFä½œæˆã‚¸ãƒ§ãƒ–ãŒä¸­æ–­ã•ã‚Œã¦ã„ã¾ã—ãŸã€‚ã‚‚ã†ä¸€åº¦PDFä½œæˆã‚’æŠ¼ã—ã¦ãã ã•ã„ã€‚'
                            }
                        }
                    } catch { }
                    $jobStatus = ([string](Get-DataProperty $job 'status' '')).ToLowerInvariant()
                    if (@('completed','completed-with-errors','failed','missing','cancelled') -notcontains $jobStatus -and (Test-RenderJobCancellationRequested $Language $normalizedJobId)) {
                        Set-NoteProperty $job 'cancelRequested' $true
                        Set-NoteProperty $job 'message' 'ä¸­æ­¢ã‚’å—ã‘ä»˜ã‘ã¾ã—ãŸã€‚ç¾åœ¨å‡¦ç†ä¸­ã®åŽŸç¨¿ãŒçµ‚ã‚ã‚‹ã¨åœæ­¢ã—ã¾ã™ã€‚'
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
        message = 'PDFä½œæˆã‚¸ãƒ§ãƒ–ã‚’æº–å‚™ã—ã¦ã„ã¾ã™ã€‚'
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
                        [void](Set-RenderJobFailedFromStartupProblem $file.FullName $job 'å‰å›žã®PDFä½œæˆãƒ—ãƒ­ã‚»ã‚¹ãŒã‚¸ãƒ§ãƒ–å‡¦ç†ã«å…¥ã‚‰ãšåœæ­¢æ‰±ã„ã«ãªã‚Šã¾ã—ãŸã€‚æ–°ã—ãPDFä½œæˆã§ãã¾ã™ã€‚')
                        continue
                    }
                    return $job
                }
                # Immediately after Start-Process there can be a short race before the child process is visible.
                if ($ageSeconds -lt 10) { return $job }
                [void](Set-RenderJobFailedFromStartupProblem $file.FullName $job 'å‰å›žã®PDFä½œæˆãƒ—ãƒ­ã‚»ã‚¹ãŒçµ‚äº†ã—ã¦ã„ãŸãŸã‚ã€æ–°ã—ã„PDFä½œæˆã‚’é–‹å§‹ã§ãã¾ã™ã€‚')
                continue
            }

            # Older builds did not write processId. Do not let those stale queued/running files block the PDF button.
            if ($ageSeconds -lt 5) { return $job }
            [void](Set-RenderJobFailedFromStartupProblem $file.FullName $job 'å¤ã„PDFä½œæˆã‚¸ãƒ§ãƒ–ã‚’çµ‚äº†æ‰±ã„ã«ã—ã¾ã—ãŸã€‚ã‚‚ã†ä¸€åº¦PDFä½œæˆã§ãã¾ã™ã€‚')
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

# ---- æå‡ºç”¨PDFã®å‡ºåŠ›ã‚¸ãƒ§ãƒ– -------------------------------------------------
# HTTPã‚µãƒ¼ãƒãƒ¼ã¯ãƒªã‚¯ã‚¨ã‚¹ãƒˆã‚’ç›´åˆ—ã«å‡¦ç†ã™ã‚‹ãŸã‚ã€å‡ºåŠ›ã‚’è¦æ±‚ã®ä¸­ã§å®Œçµã•ã›ã‚‹ã¨
# ãã®é–“ã®é€²æ—ãƒãƒ¼ãƒªãƒ³ã‚°ã‚‚ä¸­æ­¢è¦æ±‚ã‚‚å—ã‘ä»˜ã‘ã‚‰ã‚Œãªã„ã€‚å¤‰æ›PDFã¨åŒã˜ãå­ãƒ—ãƒ­ã‚»ã‚¹ã¸
# å‡ºã—ã€çŠ¶æ…‹ãƒ•ã‚¡ã‚¤ãƒ«ã‚’ãƒ–ãƒ©ã‚¦ã‚¶ãŒèª­ã‚€æ–¹å¼ã«æƒãˆã‚‹ã€‚
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
    if ([string]::IsNullOrWhiteSpace($normalized)) { throw 'æå‡ºç”¨PDFã®å‡ºåŠ›ã‚¸ãƒ§ãƒ–ã‚’ç¢ºèªã§ãã¾ã›ã‚“ã§ã—ãŸã€‚' }
    return (Join-Path (Get-RenderJobDir $Language) "$normalized.status.json")
}

function Get-FinalJobCancellationPath([string]$Language, [string]$JobId) {
    $normalized = Normalize-FinalJobId $JobId
    if ([string]::IsNullOrWhiteSpace($normalized)) { throw 'ä¸­æ­¢ã™ã‚‹å‡ºåŠ›ã‚¸ãƒ§ãƒ–ã‚’ç¢ºèªã§ãã¾ã›ã‚“ã§ã—ãŸã€‚' }
    return (Join-Path (Get-RenderJobDir $Language) "$normalized.cancel.json")
}

function Test-FinalJobCancellationRequested([string]$Language, [string]$JobId) {
    try { return (Test-Path -LiteralPath (Get-FinalJobCancellationPath $Language $JobId) -PathType Leaf) } catch { return $false }
}

function Request-FinalJobCancellation([string]$Language, [string]$JobId) {
    $normalized = Normalize-FinalJobId $JobId
    $statusPath = Get-FinalJobStatusPath $Language $normalized
    $status = Read-JsonFile $statusPath $null
    if ($null -eq $status) { throw 'ä¸­æ­¢ã™ã‚‹å‡ºåŠ›ã‚¸ãƒ§ãƒ–ãŒè¦‹ã¤ã‹ã‚Šã¾ã›ã‚“ã€‚' }
    $statusText = ([string](Get-DataProperty $status 'status' '')).ToLowerInvariant()
    if (@('completed','completed-with-errors','failed','cancelled') -contains $statusText) {
        return [pscustomobject][ordered]@{ ok=$true; jobId=$normalized; accepted=$false; alreadyFinished=$true; status=$statusText; message='æå‡ºç”¨PDFã®å‡ºåŠ›ã¯ã™ã§ã«çµ‚äº†ã—ã¦ã„ã¾ã™ã€‚' }
    }
    Write-JsonFile (Get-FinalJobCancellationPath $Language $normalized) ([ordered]@{ schemaVersion=1; jobId=$normalized; requestedAt=(New-NowIso) })
    return [pscustomobject][ordered]@{ ok=$true; jobId=$normalized; accepted=$true; alreadyFinished=$false; status=$statusText; cancelRequested=$true
        message='ä¸­æ­¢ã‚’å—ã‘ä»˜ã‘ã¾ã—ãŸã€‚ä½œæˆä¸­ã®1å†Šã¯æœ€å¾Œã¾ã§æ›¸ãä¸Šã’ã¦ã‹ã‚‰åœæ­¢ã—ã¾ã™ã€‚' }
}

function Read-FinalJobStatus([string]$Language, [string]$JobId) {
    $normalized = Normalize-FinalJobId $JobId
    $statusPath = Get-FinalJobStatusPath $Language $normalized
    $status = Read-JsonFile $statusPath $null
    if ($null -eq $status) {
        return [pscustomobject][ordered]@{ ok=$true; jobId=$normalized; status='missing'; percent=0; total=0; completed=0; failed=0; message='å‡ºåŠ›ã®çŠ¶æ…‹ã‚’èª­ã¿å–ã‚Œã¾ã›ã‚“ã§ã—ãŸã€‚' }
    }
    $statusText = ([string](Get-DataProperty $status 'status' '')).ToLowerInvariant()
    if (@('completed','completed-with-errors','failed','cancelled') -notcontains $statusText) {
        # å­ãƒ—ãƒ­ã‚»ã‚¹ãŒè½ã¡ãŸã¾ã¾ã€Œå®Ÿè¡Œä¸­ã€ã§æ®‹ã‚‹ã¨ã€ä»¥å¾Œã®å‡ºåŠ›ãŒæ°¸ä¹…ã«å§‹ã‚ã‚‰ã‚Œãªã„ã€‚
        $processId = Get-IntDataProperty $status 'processId' 0
        $alive = $false
        if ($processId -gt 0) { try { $alive = ($null -ne (Get-Process -Id $processId -ErrorAction Stop)) } catch { $alive = $false } }
        if (-not $alive -and $processId -gt 0) {
            Set-NoteProperty $status 'status' 'failed'
            Set-NoteProperty $status 'percent' 100
            Set-NoteProperty $status 'message' 'å‡ºåŠ›ãƒ—ãƒ­ã‚»ã‚¹ãŒäºˆæœŸã›ãšçµ‚äº†ã—ã¾ã—ãŸã€‚ã‚‚ã†ä¸€åº¦å‡ºåŠ›ã—ã¦ãã ã•ã„ã€‚'
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
        Set-NoteProperty $active 'message' ('æå‡ºç”¨PDFã‚’ä½œæˆä¸­ã§ã™ã€‚ ' + [string]$active.message)
        return $active
    }
    $structure = Get-Structure $Language
    $scope = Resolve-DocumentPackScope $structure $PackId $false
    $allowed = @(Get-PackTargetIds $Language $scope.pack)
    $requested = @($TargetIds | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } | Where-Object { $_ } | Select-Object -Unique)
    if ($requested.Count -eq 0) { $requested = @($allowed[0]) }
    foreach ($targetId in $requested) {
        if ($allowed -notcontains $targetId) { throw [ArgumentException]::new('ã“ã®ä¸€å¼ã«å­˜åœ¨ã—ãªã„å‡ºåŠ›å…ˆã§ã™ã€‚') }
    }
    $jobId = 'final_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '_' + ([Guid]::NewGuid().ToString('N').Substring(0,8))
    $jobDir = Get-RenderJobDir $Language
    $inputPath = Join-Path $jobDir "$jobId.input.json"
    $statusPath = Join-Path $jobDir "$jobId.status.json"
    $stdoutPath = Join-Path $jobDir "$jobId.out.log"
    $stderrPath = Join-Path $jobDir "$jobId.err.log"
    $initial = [pscustomobject][ordered]@{
        ok=$true; jobId=$jobId; kind='final-build'; status='launching'; total=$requested.Count; completed=0; failed=0; percent=1
        message='æå‡ºç”¨PDFã®ä½œæˆã‚’é–‹å§‹ã—ã¾ã™ã€‚'; currentTargetId=''; currentTargetName=''; phase=''
        packId=[string]$scope.packId; targetIds=@($requested); built=@(); skipped=@(); errors=@()
        processId=0; stdoutPath=$stdoutPath; stderrPath=$stderrPath; startedAt=(New-NowIso); updatedAt=(New-NowIso)
    }
    Write-JsonFile $statusPath $initial
    Write-JsonFile $inputPath ([ordered]@{ jobId=$jobId; mode=$Language; packId=[string]$scope.p×_4ßfòµë(š+myÖ—7Æ”æÖSÕ·7G&–æuÒ„vWBÔFF&÷W'G’GF&vWBvF—7Æ”æÖRrGF&vWD–B“²&WV—&VCÕ¶&ööÅÒ„vWBÔFF&÷W'G’GF&vWBw&WV—&VBrFfÇ6R“²vT6÷VçCÕ¶–çEÒG&VF–æW72çvT6÷VçC²F—7Æ•7FFSÕ·7G&–æuÒG&VF–æW72æF—7Æ•7FFS²6ä'V–ÆCÕ¶&ööÅÒG&VF–æW72æ6ä'V–ÆBÐ¢FÆÄ&Æö6¶W'2³Ò„vWBÔ'&’G&VF–æW72æ&Æö6¶W'2Âv†W&RÔö&¦V7B²·7G&–æuÒ„vWBÔFF&÷W'G’Eòv6öFRrrr’ÖæRvæò×vW2rÒ¢Ò6F6‚°¢GF&vWE&÷w2³Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²F&vWD–CÒGF&vWD–C²F—7Æ”æÖSÕ·7G&–æuÒ„vWBÔFF&÷W'G’GF&vWBvF—7Æ”æÖRrGF&vWD–B“²&WV—&VCÕ¶&ööÅÒ„vWBÔFF&÷W'G’GF&vWBw&WV—&VBrFfÇ6R“²vT6÷VçCÓ²F—7Æ•7FFSÒv&Æö6¶VBs²6ä'V–ÆCÒFfÇ6RÐ¢FÆÄ&Æö6¶W'2³Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²6öFSÒw&VF–æW72ÖW'&÷"s²ÖW76vSÒEòäW†6WF–öâäÖW76vRÐ¢Ð¢Ð¢F&Æö6¶W$¶W—2Ò·Ó²F&Æö6¶W'2Ò‚¢f÷&V6‚‚F&Æö6¶W"–âFÆÄ&Æö6¶W'2’°¢F¶W’Ò…·7G&–æuÒ„vWBÔFF&÷W'G’F&Æö6¶W"v6öFRrrr’Å·7G&–æuÒ„vWBÔFF&÷W'G’F&Æö6¶W"w6÷W&6T–Brrr’Å·7G&–æuÒ„vWBÔFF&÷W'G’F&Æö6¶W"w&WV—&VÖVçD–Brrr’’Ö¦ö–âwÂp¢–b‚F&Æö6¶W$¶W—2ä6öçF–ç4¶W’‚F¶W’’’²6öçF–çVRÓ²F&Æö6¶W$¶W—5²F¶W•ÒÒGG'VS²F&Æö6¶W'2³ÒF&Æö6¶W ¢Ð¢F7F—fUF&vWG2Ò‚GF&vWE&÷w2Âv†W&RÔö&¦V7B²¶–çEÒEòçvT6÷VçBÖwBÒ¢G7FFRÒv6ö×ÆWFRs²FæW‡D7F–öâÒvf–æÂp¢–b‚Gv÷&¶&öö·2ä6÷VçBÖW’²G7FFSÒvæ÷B×7F'FVBs²FæW‡D7F–öãÒvW†6VÂrÐ¢–b‚F÷fW&GVU&WV—&VBÖwB’²G7FFSÒv÷fW&GVR×6÷W&6Rs²FæW‡D7F–öãÒvW†6VÂrÐ¢VÇ6V–b‚G7V&Ö—GFVE&WV—&VBÖÇBG&WV—&VBä6÷VçB’²G7FFSÒvÖ—76–ær×6÷W&6Rs²FæW‡D7F–öãÒvW†6VÂrÐ¢VÇ6V–b‚FæVVG5&VæFW"ÖwB’²G7FFSÒvæVVG2×&VæFW"s²FæW‡D7F–öãÒvW†6VÂrÐ¢VÇ6V–b‚GVæ76–væVBÖwB’²G7FFSÒwVæ76–væVBs²FæW‡D7F–öãÒwvW2rÐ¢VÇ6V–b‚F&Æö6¶W'2ä6÷VçBÖwB’²G7FFSÒv&Æö6¶VBs²FæW‡D7F–öãÒvW†6VÂrÐ¢VÇ6V–b‚F7F—fUF&vWG2ä6÷VçBÖW’²G7FFSÒvæò×vW2s²FæW‡D7F–öãÒwvW2rÐ¢VÇ6V–b„‚F7F—fUF&vWG2Âv†W&RÔö&¦V7B²·7G&–æuÒEòæF—7Æ•7FFRÖ–â‚væVVG2×&V'V–ÆBrÂv÷WGWBÖÖ—76–ærrÂvæ÷BÖ'V–ÇBr’Ò’ä6÷VçBÖwB’²G7FFSÒvæVVG2Ö÷WGWBs²FæW‡D7F–öãÒvf–æÂrÐ¢G7F÷&VE6²ÒvWBÕ6µ&V6÷&BE7G'V7GW&RG6´–@¢G7F÷&VE&Wf–WrÒvWBÔFF&÷W'G’G7F÷&VE6²w&Wf–Wrr…¶÷&FW&VEÔ·Ò¢G&Wf–Wu7F÷&VE7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VE&Wf–Wrw7FGW2rvG&gBr¢–b‚G&Wf–Wu7F÷&VE7FGW2Öæ÷F–â‚vG&gBrÂv–â×&Wf–WrrÂv&÷fVBrÂv6†ævW2×&WVW7FVBr’’²G&Wf–Wu7F÷&VE7FGW2ÒvG&gBrÐ¢G&Wf–Wu7FGW2ÒG&Wf–Wu7F÷&VE7FGW0¢–b‚G&Wf–Wu7F÷&VE7FGW2Ö–â‚v–â×&Wf–WrrÂv&÷fVBr’’°¢G7V&Ö—GFVDf–ævW'&–çG2ÒvWBÔFF&÷W'G’G7F÷&VE&Wf–Wrw7V&Ö—GFVDf–ævW'&–çG2r…¶÷&FW&VEÔ·Ò¢G&Wf–WuF&vWG2Ò‚GF&vWE&÷w2Âv†W&RÔö&¦V7B²¶&ööÅÒEòç&WV—&VBÒ¢–b‚G&Wf–WuF&vWG2ä6÷VçBÖW’²G&Wf–WuF&vWG2Ò‚GF&vWE&÷w2Â6VÆV7BÔö&¦V7BÔf—'7B’Ð¢f÷&V6‚‚G&Wf–WuF&vWB–âG&Wf–WuF&vWG2’°¢G&Wf–WuF&vWD–BÒ·7G&–æuÒG&Wf–WuF&vWBçF&vWD–@¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G7V&Ö—GFVDf–ævW'&–çG2G&Wf–WuF&vWD–Brr’ÖæR·7G&–æuÒGF&vWDf–ævW'&–çG5²G&Wf–WuF&vWD–EÒ’²G&Wf–Wu7FGW2Òw7FÆRs²'&V²Ð¢Ð¢Ð¢2z+®Š¨Þ8îŠ‰Ž˜Ë.8òTÄô4ÄDDR˜XÞKˆ¾8îXŠžyJŽˆ^8N8Ž8î88~8;Î8+þ8¾8~8¾jè¾8(ž8®8K¹n8à¢2XŠžyJŽˆ^8¾8(ž8þXø.xZ~8~8Þ8®8N8.KÙÎjZÞ8î˜.8þX[~Y‚ŽXéþz‹þ8;¾ZHžhùµDn8;¾89®8;Î8+Žjx¾h‰8;¾X{®X©²¢28Ž8þXŠ^xšž8®8î8~8iÊ®z+®Š¨Þ8).ynyK8²6ö×ÆWFR8¾8(ž™˜ÞjÎ8^8¾8®8N8.z+®Š¨Þ8îx«nhX¾8ð¢2&Wf–Wu7FGW28Ž8~8nXŠ^8¾‹ùN8~8yK¾™Ú.XN8~X¾XŠ^8¾ŠŽzK®8ž8(¾8 ¢–b‚G7FFRÖWv6ö×ÆWFRrÖæBG&Wf–Wu7FGW2ÖWw7FÆRr’²FæW‡D7F–öâÒvf–æÂrÐ¢G&÷w2³Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢6´–CÒG6´–C²F—7Æ”æÖSÕ·7G&–æuÒ„vWBÔFF&÷W'G’G6²vF—7Æ”æÖRrrr“²6FVv÷'“Õ·7G&–æuÒ„vWBÔFF&÷W'G’G6²v6FVv÷'’rrr¢6÷W&6T6÷VçCÒGv÷&¶&öö·2ä6÷VçC²&WV—&VE6÷W&6T6÷VçCÒG&WV—&VBä6÷VçC²7V&Ö—GFVE&WV—&VE6÷W&6T6÷VçCÒG7V&Ö—GFVE&WV—&VC²÷fW&GVU&WV—&VE6÷W&6T6÷VçCÒF÷fW&GVU&WV—&VC²GVU6ööå&WV—&VE6÷W&6T6÷VçCÒFGVU6ööå&WV—&VC²æV&W7E&WV—&VDGVTFFSÒFæV&W7DGVTFFS²æVVG5&VæFW$6÷VçCÒFæVVG5&VæFW ¢vT6÷VçCÒGvW2ä6÷VçC²Væ76–væVEvT6÷VçCÒGVæ76–væVC²&Æö6¶W$6÷VçCÒF&Æö6¶W'2ä6÷VçC²&Æö6¶W'3Ô‚F&Æö6¶W'2Â6VÆV7BÔö&¦V7BÔf—'7B2¢F&vWG3Ô‚GF&vWE&÷w2“²7FFSÒG7FFS²æW‡D7F–öãÒFæW‡D7F–öã²&Wf–Wu7FGW3ÒG&Wf–Wu7FGW3²&Wf–Wu7F÷&VE7FGW3ÒG&Wf–Wu7F÷&VE7FGW3²&Wf–WtWfVçD6÷VçCÔ„vWBÔ'&’„vWBÔFF&÷W'G’G7F÷&VE&Wf–WrvWfVçG2r‚’’’ä6÷VçC²WFFVDCÔvWBÔFF&÷W'G’G6²wWFFVDBrFçVÆÀ¢Ð¢Ð¢FGFVçF–öâÒ‚G&÷w2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FFRÖæRv6ö×ÆWFRrÒ¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢6·3Ô‚G&÷w2“²F÷FÄ6÷VçCÒG&÷w2ä6÷VçC²6ö×ÆWFT6÷VçCÔ‚G&÷w2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FFRÖWv6ö×ÆWFRrÒ’ä6÷Vç@¢GFVçF–öä6÷VçCÒFGFVçF–öâä6÷VçC²Ö—76–æu&WV—&VD6÷VçCÕ¶–çEÒ‚‚G&÷w2Âf÷$V6‚Ôö&¦V7B²´ÖF…Ó£¤Ö‚ƒÅ¶–çEÒEòç&WV—&VE6÷W&6T6÷VçBÕ¶–çEÒEòç7V&Ö—GFVE&WV—&VE6÷W&6T6÷VçB’ÒÂÖV7W&RÔö&¦V7BÕ7VÒ’å7VÒ¢÷fW&GVU&WV—&VD6÷VçCÕ¶–çEÒ‚‚G&÷w2Âf÷$V6‚Ôö&¦V7B²¶–çEÒEòæ÷fW&GVU&WV—&VE6÷W&6T6÷VçBÒÂÖV7W&RÔö&¦V7BÕ7VÒ’å7VÒ¢GVU6ööå&WV—&VD6÷VçCÕ¶–çEÒ‚‚G&÷w2Âf÷$V6‚Ôö&¦V7B²¶–çEÒEòæGVU6ööå&WV—&VE6÷W&6T6÷VçBÒÂÖV7W&RÔö&¦V7BÕ7VÒ’å7VÒ¢æVVG5&VæFW$6÷VçCÕ¶–çEÒ‚‚G&÷w2Âf÷$V6‚Ôö&¦V7B²¶–çEÒEòææVVG5&VæFW$6÷VçBÒÂÖV7W&RÔö&¦V7BÕ7VÒ’å7VÒ¢Væ76–væVEvT6÷VçCÕ¶–çEÒ‚‚G&÷w2Âf÷$V6‚Ôö&¦V7B²¶–çEÒEòçVæ76–væVEvT6÷VçBÒÂÖV7W&RÔö&¦V7BÕ7VÒ’å7VÒ¢Ð§Ð ¦gVæ7F–öâ'V–ÆBÔf–æÅFb…·7G&–æuÒDÆæwVvRÅ·7G&–æuÒEföÇVÖRÅ·7G&–æuÒD6FVv÷'“Òrr’°¢26FVv÷'’8òf–Â6Æ÷6VN8.8>8>8~8(.iˆîzK®y¨N8¾jIÎŠ‹Î8ž8(¾8 ¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’D6FVv÷'¢2XÙŽKÙ>X{®X©¾8(.8î8Ž8(8nX{®X©¾8(.8[^jÛN8Ž8*.8;Î8*¾8*N89n8).hÈ8NYÎ8Ž88Ž8:ž8;>8+n8*þ8+~8:~8;>8*Ž8;>8+Ž8;>8).˜	®8ž8 ¢G"Ò–çfö¶RÔf–æÄ'V–ÆEG&ç67F–öâDÆæwVvRF6B‚EföÇVÖR¢F'V–ÇBÒ„vWBÔ'&’G"æ'V–ÇB¢–b‚F'V–ÇBä6÷VçBÖW’²F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚~Zûî‹89®8;Î8+Ž8Î8.8(®8î8¾8)>8.89®8;Î8+Žjx¾h‰8).z+®Š¨Þ8~8n8þ88^8N8"r’Ð¢&WGW&âF'V–ÇE³Ð§Ð ¦gVæ7F–öâFBÔf–æÅ6æ6†÷E6÷W&6Uv÷&¶&öö·2‚E7G'V7GW&RÂE6æ6†÷B’°¢G6VVâÒ·Ð¢G6÷W&6Uv÷&¶&öö·2Ò‚¢f÷&V6‚‚GvR–â„vWBÔ'&’„vWBÔFF&÷W'G’E6æ6†÷BwvW2r‚’’’’°¢Gv÷&¶&öö´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRwv÷&¶&öö´–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Gv÷&¶&öö´–B’Ö÷"G6VVâä6öçF–ç4¶W’‚Gv÷&¶&öö´–B’’²6öçF–çVRÐ¢G6VVå²Gv÷&¶&öö´–EÒÒGG'VP¢Gv÷&¶&öö²Ò„vWBÔ'&’E7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWGv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Gv÷&¶&öö²ä6÷VçBÖW’²6öçF–çVRÐ¢GrÒGv÷&¶&ööµ³Ð¢G6÷W&6Uv÷&¶&öö·2³Ò¶÷&FW&VEÔ°¢v÷&¶&öö´–BÒGv÷&¶&öö´–@¢f–ÆTæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’Grvf–ÆTæÖRrrr¢&VÆF—fUF‚Ò·7G&–æuÒ„vWBÔFF&÷W'G’Grw&VÆF—fUF‚rrr¢7W'&VçDW†6VÄÆ7Ew&—FUWF5F–6·2Ò·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÄÆ7Ew&—FUWF5F–6·2rrr¢7W'&VçDW†6VÅ6—¦RÒ¶–çCcEÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÅ6—¦RrÓ¢7W'&VçDW†6VÄ†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÄ†6‚rrr’¢6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’GrvÆ7E&VæFW&VE6æ6†÷D–Brrr¢fW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’GrvÆ7E&VæFW&VEfW'6–öä–Brrr¢6÷W&6T†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’GrvÆ7E&VæFW&VDW†6VÄ†6‚rrr’¢&VæFW$Vçf—&öæÖVçDf–ævW'&–çBÒ·7G&–æuÒ„vWBÔFF&÷W'G’Grw&VæFW$Vçf—&öæÖVçDf–ævW'&–çBrrr¢&VæFW$Vçf—&öæÖVçBÒvWBÔFF&÷W'G’Grw&VæFW$Vçf—&öæÖVçBrFçVÆÀ¢Ð¢Ð¢6WBÔæ÷FU&÷W'G’E6æ6†÷Bw6÷W&6Uv÷&¶&öö·2r‚G6÷W&6Uv÷&¶&öö·2¢&WGW&âE6æ6†÷@§Ð ¦gVæ7F–öâ'V–ÆBÔFö7VÖVçE6µFb…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–BÂ·7G&–æuÒEF&vWD–B’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG7G'V7GW&RE6´–BFfÇ6P¢EF&vWD–BÒ76W'BÕ6µF&vWD–BDÆæwVvRG66÷Rç6²EF&vWD–@¢GföÇVÖRÒvWBÔÆVv7•föÇVÖTg&öÕF&vWD–BDÆæwVvREF&vWD–@¢–b…¶&ööÅÒG66÷Ræ'V–ÇD–â’²&WGW&â'V–ÆBÔf–æÅFbDÆæwVvRGföÇVÖR…·7G&–æuÒG66÷Ræ6FVv÷'’’Ð ¢GF‡2ÒvWBÕF‡0¢Gv÷&·76RÒvWBÕv÷&·76UF‚DÆæwVvP¢&W÷'BÔf–æÄ'V–ÆE†6R~XX>Xéþz‹þ8îi»Nik8).z+®Š¨Þ8~8n8N8î8’p¢G'’²·fö–EÒ…66âÕWFFW2DÆæwVvRFçVÆÂFfÇ6R’Ò6F6‚²Ð¢FÆö6µF‚Ò¦ö–âÕF‚Gv÷&·76R‚&Æö6·5Ç6²Ö÷WGWE÷³Õ÷³ÒæÆö6²"Öb…·7G&–æuÒG66÷Rç6´–B’ÂEF&vWD–B¢&WGW&â–çfö¶RÕv—F„Æö6²FÆö6µF‚°¢F6ö×÷6W$¦"Ò¦ö–âÕF‚E67&—C¤&ö÷BvÆ–%ÇFf&÷…Å&W÷'EFd6ö×÷6W"æ¦"p¢GFf&÷„¦"Ò¦ö–âÕF‚E67&—C¤&ö÷BvÆ–%ÇFf&÷…ÇFf&÷‚Öæ¦"p¢–b‚Öæ÷B…FW7BÕF‚F6ö×÷6W$¦"’’²F‡&÷ru&W÷'EFd6ö×÷6W"æ¦"8Î8.8(®8î8¾8)>8"rÐ¢–b‚Öæ÷B…FW7BÕF‚GFf&÷„¦"’’²F‡&÷rwFf&÷‚Öæ¦"8Î8.8(®8î8¾8)>8"rÐ¢G6æ6†÷D&Vf÷&RÒWFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7B¢Ç’ÔFVfVÇDçVÖ&W&–æuW%föÇVÖRDÆæwVvRG7B…·7G&–æuÒG66÷Rç6´–B¢&WGW&âvWBÔf–æÄ'V–ÆD–çWE6æ6†÷BG7BDÆæwVvRGföÇVÖR…·7G&–æuÒG66÷Rç6´–B¢Ð¢·fö–EÒ„FBÔf–æÅ6æ6†÷E6÷W&6Uv÷&¶&öö·2„vWBÕ7G'V7GW&RDÆæwVvR’G6æ6†÷D&Vf÷&R¢–b‚G6æ6†÷D&Vf÷&Ræ&Æö6¶W'2ä6÷VçBÖwB’²F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr…·7G&–æuÒG6æ6†÷D&Vf÷&Ræ&Æö6¶W'5³ÒæÖW76vR’Ð¢Ff–ævW'&–çBÒ·7G&–æuÒG6æ6†÷D&Vf÷&Ræf–ævW'&–ç@¢F÷WGWDæÖRÒ·7G&–æuÒG6æ6†÷D&Vf÷&Ræ÷WGWDf–ÆTæÖP¢F÷WGWEF‚Ò¦ö–âÕF‚…·7G&–æuÒGF‡2æ÷WGWDF—"’F÷WGWDæÖP¢GFV×F‚Ò¦ö–âÕF‚…·7G&–æuÒGF‡2æ÷WGWDF—"’‚'æ'V–ÆF–æu÷³Õ÷³ÒçFb"Öb…·7G&–æuÒG66÷Rç6´–B’ÂEF&vWD–B¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GFV×F‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GFV×F‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚F÷WGWEF‚’°¢F†æFÆRÒFçVÆÀ¢G'’²F†æFÆRÒ´”òäf–ÆUÓ£¤÷Vâ‚F÷WGWEF‚Å´”òäf–ÆTÖöFUÓ£¤÷VâÅ´”òäf–ÆT66W75Ó£¥&VEw&—FRÅ´”òäf–ÆU6†&UÓ£¤æöæR’Ð¢6F6‚²F‡&÷r.X{®X©¾XXŽ8îhùX{®yJ…Dn8Î™h¾8¾8(Î8n8N8(¾8þ8(Kˆ®i»Ž8Þ8~8Þ8î8¾8)3¢F÷WGWDæÖR"Ð¢f–æÆÇ’²–b‚F†æFÆR’²F†æFÆRäF—7÷6R‚’ÒÐ¢Ð¢FÖæ–fW7BÒ¶÷&FW&VEÔ°¢66†VÖfW'6–öãÓ3²ÆæwVvSÒDÆæwVvS²6´–CÕ·7G&–æuÒG66÷Rç6´–C²F&vWD–CÒEF&vWD–C²föÇVÖSÒGföÇVÖP¢&ö¦V7D–CÕ·7G&–æuÒG6æ6†÷D&Vf÷&Rç&ö¦V7D–C²–çWDf–ævW'&–çCÒFf–ævW'&–çC²÷WGWEFcÒGFV×Fƒ²7&VFVDCÔæWrÔæ÷t—6ð¢Fö7VÖVçCÒG6æ6†÷D&Vf÷&RæFö7VÖVçC²vTçVÖ&W#Õ¶÷&FW&VEÔ¶föçCÒt&–Âs¶föçE6—¦SÓƒ¶&÷GFöÕCÓƒ¶f÷&ÖCÒv‡—†VæFVBs¶6÷VçD†–FFVãÒGG'VWÐ¢‡—6–6ÅvW3ÒG6æ6†÷D&Vf÷&Rç‡—6–6ÅvW3²vW3ÒG6æ6†÷D&Vf÷&RæÖæ–fW7EvW0¢Ð¢FÖæ–fW7EF‚Ò¦ö–âÕF‚Gv÷&·76R‚&W‡÷'G5ÆÖæ–fW7E÷³Õ÷³Òæ§6öâ"Öb…·7G&–æuÒG66÷Rç6´–B’ÂEF&vWD–B¢w&—FRÔ§6öäf–ÆRFÖæ–fW7EF‚FÖæ–fW7@¢&W÷'BÔf–æÄ'V–ÆE†6R‚~89®8;Î8+Ž8).{YYŽ8~8n8N8î8žûÈ‚r²·7G&–æuÒG6æ6†÷D&Vf÷&RçvT6÷VçB²~89®8;Î8+ŽûÈ’r¢F¦fÒ&W6öÇfRÔ¦fW†P¢G'VâÒ–çfö¶RÔæF—fT6GW&RF¦f‚rÖ7rÂ"F6ö×÷6W$¦#²GFf&÷„¦""Âu&W÷'EFd6ö×÷6W"rÂrÒÖÖæ–fW7BrÂFÖæ–fW7EF‚’E67&—C¤f–æÄ6ö×÷6UF–ÖV÷WE6V6öæG0¢–b‚G'VâçF–ÖVD÷WB’²F‡&÷r‚uDn8î{YYŽ8Ç³ÞXˆnKº^Xh^8¾{X.8(þ8(®8î8¾8)>8~8~8þ8.X{®X©¾XXŽ8Î88Þ88>88Ž8:þ8;Î8*þKˆ®8î89^8*ž8:¾888;Î8îZNYŽ8þ88N8>8þ8)5>Xh^8î89^8*ž8:¾888;Î8¾X{®X©¾8~8n8þ8n8þ88^8N8"rÖb¶–çEÒ‚E67&—C¤f–æÄ6ö×÷6UF–ÖV÷WE6V6öæG2òc’’Ð¢–b…¶–çEÒG'VâæW†—D6öFRÖæR’²F‡&÷r%Dd&÷Ž{XNx˜Ž8¾ZKiY~8~8î8~8þ8&W†—CÒB…¶–çEÒG'VâæW†—D6öFR–âB…·7G&–æuÒG'VâçFW‡B’"Ð¢–b‚Öæ÷B…FW7BÕF‚GFV×F‚’Ö÷"„vWBÔ—FVÒGFV×F‚’äÆVæwF‚ÖÆR’²F‡&÷r~hùX{®yJ…Dn8).KÙÎh‰8~8Þ8î8¾8)>8~8~8þ8"rÐ¢&W÷'BÔf–æÄ'V–ÆE†6R~X{®X©¾XXŽ8ŽKùÞZÙŽ8~8n8N8î8’p¢F'V–ÆD–BÒæWrÕ&$–@¢F6öÖÖ—BÒWFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7B¢FgFW"ÒvWBÔf–æÄ'V–ÆD–çWE6æ6†÷BG7BDÆæwVvRGföÇVÖR…·7G&–æuÒG66÷Rç6´–B¢–b…·7G&–æuÒFgFW"æf–ævW'&–çBÖæRFf–ævW'&–çB’²&WGW&â¶÷&FW&VEÔ²6†ævVCÒGG'VRÒÐ¢Ö÷fRÔ—FVÒÔÆ—FW&ÅF‚GFV×F‚ÔFW7F–æF–öâF÷WGWEF‚Ôf÷&6P¢G7FFRÒvWBÕ6´÷WGWE7FFRG7BDÆæwVvR…·7G&–æuÒG66÷Rç6´–B’GföÇVÖRGG'VP¢6WBÔæ÷FU&÷W'G’G7FFRv'V–ÇDf–ævW'&–çBrFf–ævW'&–ç@¢6WBÔæ÷FU&÷W'G’G7FFRvÆ7D'V–ÇDBr„æWrÔæ÷t—6ò¢6WBÔæ÷FU&÷W'G’G7FFRv÷WGWEFbrF÷WGWEF€¢6WBÔæ÷FU&÷W'G’G7FFRw7FÆU&V6öç2r‚¢6WBÔæ÷FU&÷W'G’G7FFRvÖW76vRr…·7G&–æuÒG'VâçFW‡B¢6WBÔæ÷FU&÷W'G’G7FFRw7FGW2rv'V–ÇBp¢6WBÔæ÷FU&÷W'G’G7FFRv'V–ÆD–BrF'V–ÆD–@¢&WGW&â¶÷&FW&VEÔ²6†ævVCÒFfÇ6S²&VF–æW73Ò„vWBÔf–æÄ'V–ÆE&VF–æW72G7BDÆæwVvRGföÇVÖR…·7G&–æuÒG66÷Rç6´–B’’Ð¢Ð¢–b…¶&ööÅÒF6öÖÖ—Bæ6†ævVB’°¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GFV×F‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GFV×F‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢F‡&÷ruDnKÙÎh‰KŠÞ8¾89®8;Î8+Žjx¾h‰8î8þ8õDnXZ^X©¾8ÎZHži»N8^8(Î8î8~8þ8.iÈik8îx«nhX¾8~XhÞ[ªnX{®X©¾8~8n8þ88^8N8"p¢Ð¢FÖæ–fW7Bæ÷WGWEFbÒF÷WGWEF€¢w&—FRÔ§6öäf–ÆRFÖæ–fW7EF‚FÖæ–fW7@¢F&6†—fUF‚Òrp¢F&6†—fTW'&÷"Òrp¢G'’²F&6†—fUF‚ÒæWrÔf–æÄ&6†—fRDÆæwVvR…·7G&–æuÒG66÷Rç6´–B’GföÇVÖRF'V–ÆD–BF÷WGWEF‚FÖæ–fW7BG6æ6†÷D&Vf÷&RÐ¢6F6‚°¢F&6†—fTW'&÷"ÒEòäW†6WF–öâäÖW76vP¢2Dn8þiz.8¾X{®X©¾XXŽ8Ži»Ž88n8N8(¾8î8~88>8>8~KÙÎh‰8Þ8î8(.8î8).ZKiY~8¾8þ8~8®8N8 ¢28þ88~8*.8;Î8*¾8*N89n8ÎKÙÎ8(Î8®8¾8>8þ8Ž8Þ8¾›¹ž8>8n˜.8(8Ž8X{®X©¾8îXX>8¾8®8>8þx˜Ž8À¢2KùÞhÈiÉþ™i>8îxËnK¨Ž8®8~8¾hè>™šN8~khŽ8Ž8(¾8.KùÞŠÛr‡–âž888þKÙÎ8(®y»N8ž8 ¢G'’°¢6WBÔf–æÅFe6æ6†÷E–ç2DÆæwVvR…·7G&–æuÒG66÷Rç6´–B’rrGföÇVÖRF'V–ÆD–B„vWBÔFF&÷W'G’G6æ6†÷D&Vf÷&Rw6÷W&6Uv÷&¶&öö·2r‚’’rp¢Ò6F6‚°¢F&6†—fTW'&÷"ÒF&6†—fTW'&÷"²ròr²EòäW†6WF–öâäÖW76vP¢Ð¢2hú8(®8N8n8^8®8N8.XŠžyJŽˆ^8¾8(.Š‰Ž˜Ë.8¾8(.jè¾8ž8 ¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvf–æÂæ&6†—fRæf–ÆVBr…¶÷&FW&VEÔ²6´–CÕ·7G&–æuÒG66÷Rç6´–C²F&vWD–CÒEF&vWD–C²föÇVÖSÒGföÇVÖS²'V–ÆD–CÒF'V–ÆD–C²ÖW76vSÒF&6†—fTW'&÷"Ò¢Ð¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvf–æÂæ'V–ÇBr…¶÷&FW&VEÔ²6´–CÕ·7G&–æuÒG66÷Rç6´–C²6FVv÷'“Òrs²F&vWD–CÒEF&vWD–C²föÇVÖSÒGföÇVÖS²'V–ÆD–CÒF'V–ÆD–C²÷WGWEFcÒF÷WGWEFƒ²&6†—fUFƒÒF&6†—fUFƒ²&6†—fTW'&÷#ÒF&6†—fTW'&÷"Ò¢&WGW&â¶÷&FW&VEÔ²6´–CÕ·7G&–æuÒG66÷Rç6´–C²F&vWD–CÒEF&vWD–C²föÇVÖSÒGföÇVÖS²÷WGWEFcÒF÷WGWEFƒ²–çWDf–ævW'&–çCÒFf–ævW'&–çC²'V–ÆD–CÒF'V–ÆD–C²&6†—fUFƒÒF&6†—fUFƒ²&6†—fTW'&÷#ÒF&6†—fTW'&÷#²ÖW76vSÕ·7G&–æuÒG'VâçFW‡C²&VF–æW73ÒF6öÖÖ—Bç&VF–æW72Ð¢Ð§Ð ¦gVæ7F–öâvWBÕ7FFU–ÆöB…·7G&–æuÒDÆæwVvR’°¢GF‡2ÒvWBÕF‡0¢F6öæf–wW&VBÒFfÇ6P¢–b‚GF‡2ÖæB·7G&–æuÒGF‡2ç7V&Ö—76–öäF—"ÖæB·7G&–æuÒGF‡2æFFF—"ÖæB·7G&–æuÒGF‡2æ÷WGWDF—"’²F6öæf–wW&VBÒGG'VRÐ¢G7G'V7GW&RÒFçVÆÀ¢G7G'V7GW&TÆöDW'&÷"Òrp¢–b‚F6öæf–wW&VBÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚…·7G&–æuÒGF‡2æFFF—"’’’°¢G'’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢Ò6F6‚°¢G7G'V7GW&TÆöDW'&÷"ÒEòäW†6WF–öâäÖW76vP¢G7G'V7GW&RÒæWrÔV×G•7G'V7GW&RDÆæwVvP¢Ð¢ÒVÇ6R°¢G7G'V7GW&RÒæWrÔV×G•7G'V7GW&RDÆæwVvP¢Ð¢Gv÷&¶&öö·2Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2¢GvW2Ò„vWBÔ'&’G7G'V7GW&RçvW2¢G7VÖÖ'’Ò¶÷&FW&VEÔ°¢W†6VÅWFFVBÒ‚Gv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FGW2Ö–â‚vW†6VÂ×WFFVBrÂw6÷W&6R×WFFVBr’Ò’ä6÷Vç@¢Væ6†V6¶VEvW2Ò‚GvW2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FGW2Ö–â‚w&VæFW&VBrÂw7FÆRrÂvæ÷B×&VæFW&VBr’Ò’ä6÷Vç@¢6öæf—&ÖVEvW2Ò‚GvW2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FGW2ÖWv6öæf—&ÖVBrÒ’ä6÷Vç@¢&VæFW$W'&÷'2Ò‚Gv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FGW2ÖWw&VæFW"ÖW'&÷"rÒ’ä6÷Vç@¢Fe&VG•v÷&¶&öö·2Ò‚Gv÷&¶&öö·2Âv†W&RÔö&¦V7B²Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒEòæÆ7E&VæFW&VDW†6VÄ†6‚’ÖæB·7G&–æuÒEòç7FGW2ÖæRw&VæFW"ÖW'&÷"rÒ’ä6÷Vç@¢FeVæF–æuv÷&¶&öö·2Ò‚Gv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒEòæÆ7E&VæFW&VDW†6VÄ†6‚’Ö÷"·7G&–æuÒEòç7FGW2Ö–â‚væWrrÂvW†6VÂ×WFFVBrÂw6÷W&6R×WFFVBrÂw&VæFW"ÖW'&÷"r’Ò’ä6÷Vç@¢Fe&VG•vW2Ò‚GvW2Âv†W&RÔö&¦V7B²Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒEòæ6öçFVçEFb’Ò’ä6÷Vç@¢F÷FÅv÷&¶&öö·2ÒGv÷&¶&öö·2ä6÷Vç@¢F÷FÅvW2ÒGvW2ä6÷Vç@¢Ð¢–b‚F6öæf–wW&VBÖæBÖæ÷BG7G'V7GW&TÆöDW'&÷"’°¢f÷&V6‚‚F6B–â‚vV6ÒrÂv&öBrÂvFÖÒr’’°¢f÷&V6‚‚GföÇVÖR–â„vWBÕföÇVÖTÆ—7BDÆæwVvRÂv†W&RÔö&¦V7B²EòÖæRvæöæRrÒ’’°¢F¶W“ÔvWBÕföÇVÖU7FFT¶W’GföÇVÖRF6C²GcÔvWBÔFF&÷W'G’G7G'V7GW&RçföÇVÖW2F¶W’FçVÆÀ¢–b‚FçVÆÂÖæRGb’²F÷WCÕ·7G&–æuÒ„vWBÔFF&÷W'G’Gbv÷WGWEFbrrr“²6WBÔæ÷FU&÷W'G’Gbv÷WGWEFdW†—7G2r‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F÷WB’’Ð¢Ð¢Ð¢Ð¢2X‰ÞiÉþŠŽzK®8~8þX[iÈž89^8*ž8:¾888;ÎKˆ®8îjùN‹È4¥4ôî8)$W†6VÎK»ni[XˆnŠªÞ8î8®8N8 ¢2DnKÙÎh‰i˜.8·7G'V7GW&^8ŽKùÞZÙŽ8~8þ[þ8^8®Šh{HN888).‹ùN8ž8 ¢F6†ævU7VÖÖ&–W2Ò¶÷&FW&VEÔ·Ð¢F–çWD†—7F÷'”öâÒFfÇ6P¢G'’²F–çWD†—7F÷'”öâÒ…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’Ò6F6‚²Ð¢–b‚F6öæf–wW&VBÖæBÖæ÷BG7G'V7GW&TÆöDW'&÷"ÖæBF–çWD†—7F÷'”öâ’°¢f÷&V6‚‚Gr–âGv÷&¶&öö·2’°¢F72ÒvWBÔFF&÷W'G’GrvÆFW7D6ö×&—6öå7VÖÖ'’rFçVÆÀ¢–b‚FçVÆÂÖæRF72’²F6†ævU7VÖÖ&–W5µ·7G&–æuÒGrçv÷&¶&öö´–EÒÒF72Ð¢Ð¢Ð¢FWFõ7VÖÖ'’ÒFçVÆÀ¢G'’²–b‚F6öæf–wW&VB’²FWFõ7VÖÖ'’ÒvWBÔWFõ7FFU7VÖÖ'’DÆæwVvRÔf7BÒÒ6F6‚²Ð ¢GFf§4F—"Ò¦ö–âÕF‚E67&—C¥vV%&ö÷BwFf§2p¢GFf§46Æ76–2Ò‚…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚GFf§4F—"wFbæÖ–âæ§2r’’ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚GFf§4F—"wFbçv÷&¶W"æÖ–âæ§2r’’¢GFf§4ÖöGVÆRÒ‚…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚GFf§4F—"wFbæÖ–âæÖ§2r’’ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚GFf§4F—"wFbçv÷&¶W"æÖ–âæÖ§2r’’¢GFf§4ÖöFRÒvæöæRp¢–b‚GFf§4ÖöGVÆR’²GFf§4ÖöFRÒvÖöGVÆRrÒVÇ6V–b‚GFf§46Æ76–2’²GFf§4ÖöFRÒv6Æ76–2rÐ¢G6µFV×ÆFW2Ò„vWBÕ6µFV×ÆFT6FÆörDÆæwVvR¢GV&Æ–56·2Ò„vWBÕV&Æ–56´Æ—7BG7G'V7GW&RDÆæwVvR¢G6µ&öw&W72Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²6·3Ô‚“²F÷FÄ6÷VçCÓ²6ö×ÆWFT6÷VçCÓ²GFVçF–öä6÷VçCÓ²Ö—76–æu&WV—&VD6÷VçCÓ²÷fW&GVU&WV—&VD6÷VçCÓ²GVU6ööå&WV—&VD6÷VçCÓ²æVVG5&VæFW$6÷VçCÓ²Væ76–væVEvT6÷VçCÓÐ¢–b‚F6öæf–wW&VBÖæBÖæ÷BG7G'V7GW&TÆöDW'&÷"’²G'’²G6µ&öw&W72ÒvWBÕ6µ&öw&W74F6†&ö&BG7G'V7GW&RDÆæwVvRÒ6F6‚²w&—FRÕv&æ–ær‚~Kˆ[Èþ˜.hÙ~8).™¸nŠˆŽ8~8Þ8î8¾8)3¢r²EòäW†6WF–öâäÖW76vR’ÒÐ¢289®8;Î8+Žjx¾h‰8îj[ÞŠk>8:Þ88>8*þyJŽ8.yK¾™Ú.8þ8>8(Î8)"&6TÆ–÷WB8Ž8~8n˜8(®‹ùN8ž8 ¢FÆ–÷WDf–ævW'&–çG2Ò¶÷&FW&VEÔ·Ð¢–b‚F6öæf–wW&VBÖæBÖæ÷BG7G'V7GW&TÆöDW'&÷"’°¢G'’²f÷&V6‚‚G6²–âGV&Æ–56·2’²FÆ–÷WDf–ævW'&–çG5µ·7G&–æuÒG6²ç6´–EÒÒvWBÕvTÆ–÷WDf–ævW'&–çBG7G'V7GW&R…·7G&–æuÒG6²ç6´–B’ÒÐ¢6F6‚²w&—FRÕv&æ–ær‚~89®8;Î8+Žjx¾h‰8îhÈ~{H¾8).ŠˆŽzé~8~8Þ8î8¾8)3¢r²EòäW†6WF–öâäÖW76vR’Ð¢Ð¢&WGW&â¶÷&FW&VEÔ°¢ö²ÒGG'VP¢Fö¶VâÒE67&—C¥Fö¶Và¢6öæf–wW&VBÒF6öæf–wW&V@¢F‡2ÒGF‡0¢6µFV×ÆFW2ÒG6µFV×ÆFW0¢6·2ÒGV&Æ–56·0¢7G'V7GW&RÒ6öçfW'EFòÕcE7G'V7GW&T6ö×F–&–Æ—G•f–WrG7G'V7GW&P¢Æ–÷WDf–ævW'&–çG2ÒFÆ–÷WDf–ævW'&–çG0¢7G'V7GW&TÆöDW'&÷"ÒG7G'V7GW&TÆöDW'&÷ ¢f–æÅ&VF–æW72ÒB†–b‚F6öæf–wW&VBÖæBÖæ÷BG7G'V7GW&TÆöDW'&÷"’²vWBÔÆÄf–æÅ&VF–æW72G7G'V7GW&RDÆæwVvRFfÇ6RÒVÇ6R²¶÷&FW&VEÔ·ÒÒ¢6µ&öw&W72ÒG6µ&öw&W70¢7VÖÖ'’ÒG7VÖÖ'¢&V6VçDW'&÷'2Ò„vWBÔ'&’Gv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FGW2ÖWw&VæFW"ÖW'&÷"rÖ÷"Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒEòæÆ7DW'&÷"’ÒÂf÷$V6‚Ôö&¦V7B²¶÷&FW&VEÔ²v÷&¶&öö´–BÒ·7G&–æuÒEòçv÷&¶&öö´–C²f–ÆTæÖRÒ·7G&–æuÒEòæf–ÆTæÖS²F—7Æ”æÖRÒ·7G&–æuÒEòæF—7Æ”æÖS²ÖW76vRÒ·7G&–æuÒEòæÆ7DW'&÷%W6W#²FWF–ÂÒ·7G&–æuÒEòæÆ7DW'&÷#²BÒ·7G&–æuÒEòæÆ7DW'&÷$BÒÒ¢Ff§5&W6VçBÒ‚GFf§46Æ76–2Ö÷"GFf§4ÖöGVÆR¢Ff§4ÖöFRÒGFf§4ÖöFP¢W†6VÅ&–çE&öf–ÆUfW'6–öâÒE67&—C¤W†6VÅ&–çE&öf–ÆUfW'6–öà¢Fd–×÷'E&öf–ÆUfW'6–öâÒE67&—C¥Fd–×÷'E&öf–ÆUfW'6–öà¢v÷&E&VæFW%&öf–ÆUfW'6–öâÒE67&—C¥v÷&E&VæFW%&öf–ÆUfW'6–öà¢÷vW%ö–çE&VæFW%&öf–ÆUfW'6–öâÒE67&—C¥÷vW%ö–çE&VæFW%&öf–ÆUfW'6–öà¢–çWD†—7F÷'”Væ&ÆVBÒF–çWD†—7F÷'”öà¢6†ævU7VÖÖ&–W2ÒF6†ævU7VÖÖ&–W0¢WFòÒFWFõ7VÖÖ'¢WFõ&VæFW$–å&öw&W72ÒE67&—C¤WFõ&VæFW$–å&öw&W70¢6‡WFF÷väöåF$6Æ÷6RÒFfÇ6P¢Ð§Ð ¦gVæ7F–öâvWBÕc%7FFU–ÆöB…·7G&–æuÒDÆæwVvR’°¢FÆVv7’ÒvWBÕ7FFU–ÆöBDÆæwVvP¢F6öæf–wW&VBÒ¶&ööÅÒ„vWBÔFF&÷W'G’FÆVv7’v6öæf–wW&VBrFfÇ6R¢G7G'V7GW&TÆöDW'&÷"Ò·7G&–æuÒ„vWBÔFF&÷W'G’FÆVv7’w7G'V7GW&TÆöDW'&÷"rrr¢G7G'V7GW&RÒ–b‚F6öæf–wW&VBÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G7G'V7GW&TÆöDW'&÷"’’²vWBÕ7G'V7GW&RDÆæwVvRÒVÇ6R²æWrÔV×G•7G'V7GW&RDÆæwVvRÐ¢G6÷W&6W2Ò‚¢f÷&V6‚‚G6÷W&6R–â„vWBÔ'&’„vWBÔFF&÷W'G’G7G'V7GW&Rw6÷W&6W2r‚’’’’°¢G6÷W&6W2³Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢6÷W&6T–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6Rw6÷W&6T–Brrr¢6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6Rw6´–Brrr¢&VÆF—fUF‚Ò·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6Rw&VÆF—fUF‚rrr¢6÷W&6UG—RÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6Rw6÷W&6UG—Rrrr¢FFW$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6RvFFW$–Brrr¢F—7Æ”æÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6RvF—7Æ”æÖRrrr¢÷væW$FW'FÖVçBÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6Rv÷væW$FW'FÖVçBrrr¢&WV—&VBÒ¶&ööÅÒ„vWBÔFF&÷W'G’G6÷W&6Rw&WV—&VBrGG'VR¢FVfVÇEF&vWD–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6RvFVfVÇEF&vWD–BrwVæ76–væVBr¢&WV—&VÖVçD–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6Rw&WV—&VÖVçD–Brrr¢7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6Rw7FGW2rrr¢7W'&VçE6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6Rv7W'&VçE6æ6†÷D–Brrr¢Æ7E&VæFW&VE6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6RvÆ7E&VæFW&VE6æ6†÷D–Brrr¢Æ7E&VæFW&VEfW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6RvÆ7E&VæFW&VEfW'6–öä–Brrr¢Æ7E&VæFW&VDBÒvWBÔFF&÷W'G’G6÷W&6RvÆ7E&VæFW&VDBrFçVÆÀ¢Ð¢Ð¢&WGW&â¶÷&FW&VEÔ°¢ö²ÒGG'VP¢•fW'6–öâÒ ¢FöÖ–å66†VÖfW'6–öâÒ0¢Fö¶VâÒvWBÔFF&÷W'G’FÆVv7’wFö¶VârE67&—C¥Fö¶Và¢6öæf–wW&VBÒF6öæf–wW&V@¢F‡2ÒvWBÔFF&÷W'G’FÆVv7’wF‡2rFçVÆÀ¢7G'V7GW&TÆöDW'&÷"ÒG7G'V7GW&TÆöDW'&÷ ¢6µFV×ÆFW2Ò„vWBÕ6µFV×ÆFT6FÆörDÆæwVvR¢6·2Ò„vWBÕV&Æ–56´Æ—7BG7G'V7GW&RDÆæwVvR¢7G'V7GW&RÒ¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ0¢6·2Ò„vWBÕV&Æ–56´Æ—7BG7G'V7GW&RDÆæwVvR¢6÷W&6W2Ò‚G6÷W&6W2¢Væ—G2Ò„vWBÔ'&’„vWBÔFF&÷W'G’G7G'V7GW&RwVæ—G2r‚’’¢—FV×2Ò„vWBÔ'&’„vWBÔFF&÷W'G’G7G'V7GW&Rv—FV×2r‚’’¢'F–f7G2Ò„vWBÔ'&’„vWBÔFF&÷W'G’G7G'V7GW&Rv'F–f7G2r‚’’¢÷WGWG2ÒvWBÔFF&÷W'G’G7G'V7GW&Rv÷WGWG2r…¶÷&FW&VEÔ·Ò¢Ö–w&F–öä—77VW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’G7G'V7GW&RvÖ–w&F–öä—77VW2r‚’’¢WFFVDBÒvWBÔFF&÷W'G’G7G'V7GW&RwWFFVDBrFçVÆÀ¢Ð¢f–æÅ&VF–æW72ÒvWBÔFF&÷W'G’FÆVv7’vf–æÅ&VF–æW72r…¶÷&FW&VEÔ·Ò¢6µ&öw&W72ÒvWBÔFF&÷W'G’FÆVv7’w6µ&öw&W72r…¶÷&FW&VEÔ·Ò¢7VÖÖ'’ÒvWBÔFF&÷W'G’FÆVv7’w7VÖÖ'’r…¶÷&FW&VEÔ·Ò¢&V6VçDW'&÷'2Ò„vWBÔ'&’„vWBÔFF&÷W'G’FÆVv7’w&V6VçDW'&÷'2r‚’’¢–çWD†—7F÷'”Væ&ÆVBÒ¶&ööÅÒ„vWBÔFF&÷W'G’FÆVv7’v–çWD†—7F÷'”Væ&ÆVBrFfÇ6R¢6†ævU7VÖÖ&–W2ÒvWBÔFF&÷W'G’FÆVv7’v6†ævU7VÖÖ&–W2r…¶÷&FW&VEÔ·Ò¢Ð§Ð ¦gVæ7F–öâ&VBÔ&öG”§6öâ‚E&WVW7B’°¢G&VFW"ÒæWrÔö&¦V7B”òå7G&VÕ&VFW"‚E&WVW7Bä–çWE7G&VÒÂE&WVW7Bä6öçFVçDVæ6öF–ær¢GFW‡BÒG&VFW"å&VEFôVæB‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GFW‡B’’²&WGW&â·67W7FöÖö&¦V7EÔ·ÒÐ¢&WGW&âGFW‡BÂ6öçfW'Dg&öÒÔ§6öà§Ð ¦gVæ7F–öâF÷V6‚Õ&W7öç6T7F—f—G’°¢–b‚E67&—C¤6Æ–VçDGF6†VBÖæBE67&—C¤6Æ–VçD6Æ÷6Tæ÷F–f–VEWF2ÖW´FFUF–ÖUÓ£¤Ö–åfÇVR’²E67&—C¤Æ7D†V'F&VEWF2Ò´FFUF–ÖUÓ£¥WF4æ÷rÐ§Ð ¦gVæ7F–öâw&—FRÕFW‡E&W7öç6R‚D6öçFW‡BÂ¶–çEÒE7FGW2Â·7G&–æuÒD&öG’Â·7G&–æuÒD6öçFVçEG—RÂ¶&ööÅÒDÆÆ÷t6÷'2ÒFfÇ6R’°¢F÷V6‚Õ&W7öç6T7F—f—G¢F'—FW2ÒµFW‡BäVæ6öF–æuÓ£¥UDc‚ävWD'—FW2‚D&öG’¢–b…FW7BÕF76öçFW‡BD6öçFW‡B’²w&—FRÕF7&W7öç6RD6öçFW‡BE7FGW2F'—FW2D6öçFVçEG—RDÆÆ÷t6÷'3²&WGW&âÐ¢D6öçFW‡Bå&W7öç6Rå7FGW46öFRÒE7FGW0¢D6öçFW‡Bå&W7öç6Rä6öçFVçEG—RÒD6öçFVçEG—P¢D6öçFW‡Bå&W7öç6Rä†VFW'5²t66†RÔ6öçG&öÂuÒÒvæò×7F÷&Rp¢–b‚DÆÆ÷t6÷'2’²D6öçFW‡Bå&W7öç6Rä†VFW'5²t66W72Ô6öçG&öÂÔÆÆ÷rÔ÷&–v–âuÒÒr¢rÐ¢D6öçFW‡Bå&W7öç6Rä6öçFVçDÆVæwFƒcBÒF'—FW2äÆVæwF€¢D6öçFW‡Bå&W7öç6Rä÷WGWE7G&VÒåw&—FR‚F'—FW2ÂÂF'—FW2äÆVæwF‚¢D6öçFW‡Bå&W7öç6Rä÷WGWE7G&VÒä6Æ÷6R‚§Ð ¦gVæ7F–öâw&—FRÔ§6öå&W7öç6R‚D6öçFW‡BÂ¶–çEÒE7FGW2ÂDö&¦V7BÂ¶&ööÅÒDÆÆ÷t6÷'2ÒFfÇ6R’°¢F§6öâÒ6öçfW'EFòÔ§6öâÔ–çWDö&¦V7BDö&¦V7BÔFWF‚S ¢w&—FRÕFW‡E&W7öç6RD6öçFW‡BE7FGW2F§6öâvÆ–6F–öâö§6öã²6†'6WC×WFbÓ‚rDÆÆ÷t6÷'0§Ð ¦gVæ7F–öâw&—FRÔ'—FW5&W7öç6R‚D6öçFW‡BÂ¶–çEÒE7FGW2Â¶'—FUµÕÒD'—FW2Â·7G&–æuÒD6öçFVçEG—RÂ¶&ööÅÒDÆÆ÷t6÷'2ÒFfÇ6RÂ·7G&–æuÒD66†T6öçG&öÂÒvæò×7F÷&Rr’°¢F÷V6‚Õ&W7öç6T7F—f—G¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D66†T6öçG&öÂ’’²D66†T6öçG&öÂÒvæò×7F÷&RrÐ¢D66†T6öçG&öÂÒD66†T6öçG&öÂ×&WÆ6R%µÇ%ÆåÒ"Ârp¢–b…FW7BÕF76öçFW‡BD6öçFW‡B’²w&—FRÕF7&W7öç6RD6öçFW‡BE7FGW2D'—FW2D6öçFVçEG—RDÆÆ÷t6÷'2D66†T6öçG&öÃ²&WGW&âÐ¢D6öçFW‡Bå&W7öç6Rå7FGW46öFRÒE7FGW0¢D6öçFW‡Bå&W7öç6Rä6öçFVçEG—RÒD6öçFVçEG—P¢D6öçFW‡Bå&W7öç6Rä†VFW'5²t66†RÔ6öçG&öÂuÒÒD66†T6öçG&öÀ¢–b‚DÆÆ÷t6÷'2’²D6öçFW‡Bå&W7öç6Rä†VFW'5²t66W72Ô6öçG&öÂÔÆÆ÷rÔ÷&–v–âuÒÒr¢rÐ¢D6öçFW‡Bå&W7öç6Rä6öçFVçDÆVæwFƒcBÒD'—FW2äÆVæwF€¢D6öçFW‡Bå&W7öç6Rä÷WGWE7G&VÒåw&—FR‚D'—FW2ÂÂD'—FW2äÆVæwF‚¢D6öçFW‡Bå&W7öç6Rä÷WGWE7G&VÒä6Æ÷6R‚§Ð  ¦gVæ7F–öâw&—FRÔf–ÆU&W7öç6R‚D6öçFW‡BÂ¶–çEÒE7FGW2Â·7G&–æuÒDgVÆÅF‚Â·7G&–æuÒD6öçFVçEG—RÂ¶&ööÅÒDÆÆ÷t6÷'2ÒFfÇ6RÂ·7G&–æuÒD66†T6öçG&öÂÒvæò×7F÷&Rr’°¢F÷V6‚Õ&W7öç6T7F—f—G¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D66†T6öçG&öÂ’’²D66†T6öçG&öÂÒvæò×7F÷&RrÐ¢D66†T6öçG&öÂÒD66†T6öçG&öÂ×&WÆ6R%µÇ%ÆåÒ"Ârp¢D6öçFVçEG—RÒ…·7G&–æuÒD6öçFVçEG—R’×&WÆ6R%µÇ%ÆåÒ"Ârp¢Ff–ÆT–æfòÒ´”òäf–ÆT–æfõÓ£¦æWr‚„6öçfW'EFòÕv–ã3$W‡FVæFVEF‚DgVÆÅF‚’¢–b‚Öæ÷BFf–ÆT–æfòäW†—7G2’²F‡&÷r~˜XÞKú8ž8(¾89^8*8*N8:¾8ÎŠh¾8N8¾8(®8î8¾8)>8"rÐ¢G6÷W&6RÒ´”òäf–ÆUÓ£¤÷Vâ‚Ff–ÆT–æfòägVÆÄæÖRÂ´”òäf–ÆTÖöFUÓ£¤÷VâÂ´”òäf–ÆT66W75Ó£¥&VBÂ´”òäf–ÆU6†&UÓ£¥&VB¢–b…FW7BÕF76öçFW‡BD6öçFW‡B’°¢G'’°¢G7FGW5FW‡BÒvWBÔ‡GG7FGW5FW‡BE7FGW0¢F6÷'4†VFW"Òrp¢–b‚DÆÆ÷t6÷'2’²F6÷'4†VFW"Ò$66W72Ô6öçG&öÂÔÆÆ÷rÔ÷&–v–ã¢¦&â"Ð¢F†VFW"Ò$…EEóãE7FGW2G7FGW5FW‡F&ä6öçFVçBÕG—S¢D6öçFVçEG—V&ä6öçFVçBÔÆVæwFƒ¢B‚Ff–ÆT–æfòäÆVæwF‚–&ä66†RÔ6öçG&öÃ¢D66†T6öçG&öÆ&âG¶6÷'4†VFW'Ô6öææV7F–öã¢6Æ÷6V&æ&â ¢F†VFW$'—FW2ÒµFW‡BäVæ6öF–æuÓ£¤44”’ävWD'—FW2‚F†VFW"¢GF&vWBÒD6öçFW‡BåF77G&VÐ¢GF&vWBåw&—FR‚F†VFW$'—FW2ÂÂF†VFW$'—FW2äÆVæwF‚¢G6÷W&6Rä6÷•Fò‚GF&vWBÂcSS3b¢GF&vWBäfÇW6‚‚¢Òf–æÆÇ’°¢G6÷W&6RäF—7÷6R‚¢G'’²–b‚D6öçFW‡BåF77G&VÒ’²D6öçFW‡BåF77G&VÒä6Æ÷6R‚’ÒÒ6F6‚·Ð¢G'’²–b‚D6öçFW‡BåF76Æ–VçB’²D6öçFW‡BåF76Æ–VçBä6Æ÷6R‚’ÒÒ6F6‚·Ð¢Ð¢&WGW&à¢Ð¢G'’°¢D6öçFW‡Bå&W7öç6Rå7FGW46öFRÒE7FGW0¢D6öçFW‡Bå&W7öç6Rä6öçFVçEG—RÒD6öçFVçEG—P¢D6öçFW‡Bå&W7öç6Rä†VFW'5²t66†RÔ6öçG&öÂuÒÒD66†T6öçG&öÀ¢–b‚DÆÆ÷t6÷'2’²D6öçFW‡Bå&W7öç6Rä†VFW'5²t66W72Ô6öçG&öÂÔÆÆ÷rÔ÷&–v–âuÒÒr¢rÐ¢D6öçFW‡Bå&W7öç6Rä6öçFVçDÆVæwFƒcBÒFf–ÆT–æfòäÆVæwF€¢G6÷W&6Rä6÷•Fò‚D6öçFW‡Bå&W7öç6Rä÷WGWE7G&VÒÂcSS3b¢Òf–æÆÇ’°¢G6÷W&6RäF—7÷6R‚¢D6öçFW‡Bå&W7öç6Rä÷WGWE7G&VÒä6Æ÷6R‚¢Ð§Ð ¦gVæ7F–öâvWBÔÖ–ÖR…·7G&–æuÒEF‚’°¢7v—F6‚…´”òåF…Ó£¤vWDW‡FVç6–öâ‚EF‚’åFôÆ÷vW$–çf&–çB‚’’°¢ræ‡FÖÂr²&WGW&âwFW‡Bö‡FÖÃ²6†'6WC×WFbÓ‚rÐ¢ræ772r²&WGW&âwFW‡Bö773²6†'6WC×WFbÓ‚rÐ¢ræ§2r²&WGW&âvÆ–6F–öâö¦f67&—C²6†'6WC×WFbÓ‚rÐ¢ræÖ§2r²&WGW&âvÆ–6F–öâö¦f67&—C²6†'6WC×WFbÓ‚rÐ¢ræ§6öâr²&WGW&âvÆ–6F–öâö§6öã²6†'6WC×WFbÓ‚rÐ¢rçFbr²&WGW&âvÆ–6F–öâ÷FbrÐ¢rçærr²&WGW&âv–ÖvR÷ærrÐ¢rç7frr²&WGW&âv–ÖvR÷7fr·†ÖÂrÐ¢FVfVÇB²&WGW&âvÆ–6F–öâöö7FWB×7G&VÒrÐ¢Ð§Ð ¦gVæ7F–öâvWBÕ&WVW7D6öö¶–UfÇVR‚E&WVW7BÂ·7G&–æuÒDæÖR’°¢G'’°¢F6öö¶–T†VFW"Ò·7G&–æuÒE&WVW7Bä†VFW'5²t6öö¶–RuÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öö¶–T†VFW"’’²&WGW&ârrÐ¢f÷&V6‚‚G'B–â‚F6öö¶–T†VFW"×7Æ—Bs²r’’°¢F—FVÒÒG'BåG&–Ò‚¢FWÒF—FVÒä–æFW„öb‚sÒr¢–b‚FWÖÆR’²6öçF–çVRÐ¢FâÒF—FVÒå7V'7G&–ærƒÂFW’åG&–Ò‚¢–b‚FâÖæRDæÖR’²6öçF–çVRÐ¢&WGW&âµW&•Ó£¥VæW66TFF7G&–ær‚F—FVÒå7V'7G&–ær‚FW²’¢Ð¢Ò6F6‚·Ð¢&WGW&ârp§Ð ¦gVæ7F–öâFW7BÔf—†VEF–ÖUFö¶VäWVÇ2…·7G&–æuÒD6æF–FFRÂ·7G&–æuÒDW‡V7FVB’°¢–b‚FçVÆÂÖWD6æF–FFR’²D6æF–FFRÒrrÐ¢–b‚FçVÆÂÖWDW‡V7FVB’²DW‡V7FVBÒrrÐ¢GWFc‚ÒµFW‡BäVæ6öF–æuÓ£¥UDc€¢G6†Òµ6V7W&—G’ä7'—Föw&‡’å4„#SeÓ£¤7&VFR‚¢G'’°¢F6æF–FFT†6‚ÒG6†ä6ö×WFT†6‚‚GWFc‚ävWD'—FW2‚D6æF–FFR’¢FW‡V7FVD†6‚ÒG6†ä6ö×WFT†6‚‚GWFc‚ävWD'—FW2‚DW‡V7FVB’¢Òf–æÆÇ’°¢G6†äF—7÷6R‚¢Ð¢FF–ffW&Væ6RÒ ¢f÷"‚F’Ò²F’ÖÇBF6æF–FFT†6‚äÆVæwFƒ²F’²²’°¢FF–ffW&Væ6RÒFF–ffW&Væ6RÖ&÷"‚F6æF–FFT†6…²F•ÒÖ'†÷"FW‡V7FVD†6…²F•Ò¢Ð¢&WGW&â‚FF–ffW&Væ6RÖWÖæBD6æF–FFRäÆVæwF‚ÖWDW‡V7FVBäÆVæwF‚§Ð ¢28:¾8;Î89~8988>8*þ8æ6öö¶–^8þ89Þ8;Î88Ž8~Xˆn™º.8^8(Î8®8N8##rããã8îXŠ^89Þ8;Î88Ž8~X¹^8þK»¾hHþ8î89®8;Î8+Ž8À¢2&W÷'D&–æFW%Fö¶Vî8).ŠªÞ8þ8YÎKˆ8+^8*N88Žh›8N8î8î8äž8).ZéþŠÎ8~8Þ8n8~8î8n8 ¢2†÷7Bô÷&–v–î8).jIÎŠ‹Î8~8n8ˆz®Xˆn8î8*®8:®8+Ž8;>Kº^ZIn8¾8(ž8äžYÎ8>X{®8~8).h¹.Y
n8ž8(¾8 ¦gVæ7F–öâFW7BÕ&WVW7D÷&–v–â‚E&WVW7B’°¢FÆÆ÷vVBÒ‚##rããã¢B‚E67&—C¥÷'B’"Â&Æö6Æ†÷7C¢B‚E67&—C¥÷'B’"¢F†÷7D†VFW"Ò·7G&–æuÒE&WVW7Bä†VFW'5²t†÷7BuÐ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F†÷7D†VFW"’ÖæB‚FÆÆ÷vVBÖæ÷F6öçF–ç2F†÷7D†VFW"’’²&WGW&âFfÇ6RÐ¢F÷&–v–âÒ·7G&–æuÒE&WVW7Bä†VFW'5²t÷&–v–âuÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F÷&–v–â’’²&WGW&âGG'VRÐ¢f÷&V6‚‚F–âFÆÆ÷vVB’²–b‚F÷&–v–âÖW&‡GG¢òòF"’²&WGW&âGG'VRÒÐ¢&WGW&âFfÇ6P§Ð ¦gVæ7F–öâFW7BÕFö¶Vâ‚E&WVW7B’°¢GÒ·7G&–æuÒE&WVW7BåVW'•7G&–æu²wFö¶VâuÐ¢GBÒ·7G&–æuÒE&WVW7BåVW'•7G&–æu²wBuÐ¢F‚Ò·7G&–æuÒE&WVW7Bä†VFW'5²u‚Õ&W÷'D&–æFW"ÕFö¶VâuÐ¢F2ÒvWBÕ&WVW7D6öö¶–UfÇVRE&WVW7Bu&W÷'D&–æFW%Fö¶Vâp¢&WGW&â‚…FW7BÔf—†VEF–ÖUFö¶VäWVÇ2GE67&—C¥Fö¶Vâ’Ö÷ ¢…FW7BÔf—†VEF–ÖUFö¶VäWVÇ2GBE67&—C¥Fö¶Vâ’Ö÷ ¢…FW7BÔf—†VEF–ÖUFö¶VäWVÇ2F‚E67&—C¥Fö¶Vâ’Ö÷ ¢…FW7BÔf—†VEF–ÖUFö¶VäWVÇ2F2E67&—C¥Fö¶Vâ’§Ð ¦gVæ7F–öâ6W'fRÕ7FF–2‚D6öçFW‡BÂ·7G&–æuÒEF‚’°¢–b‚EF‚ÖWròr’²EF‚Òrö–æFW‚æ‡FÖÂrÐ¢G&VÂÒEF‚åG&–Õ7F'B‚ròr’×&WÆ6RròrÂ´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6† ¢–b‚G&VÂÖÖF6‚r…çÅµÅÂõÒ•ÂåÂâ‚GÅµÅÂõÒ’r’²w&—FRÔ§6öå&W7öç6RD6öçFW‡BC…¶÷&FW&VEÔ²ö²ÒFfÇ6S²W'&÷"Òrââ—2æ÷BÆÆ÷vVBrÒ“²&WGW&âÐ¢Ff–ÆRÒ´”òåF…Ó£¤vWDgVÆÅF‚‚„¦ö–âÕF‚E67&—C¥vV%&ö÷BG&VÂ’¢G&ö÷BÒ´”òåF…Ó£¤vWDgVÆÅF‚‚E67&—C¥vV%&ö÷B¢–b‚Öæ÷BG&ö÷BäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’’²G&ö÷B³Ò´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"Ð¢–b‚Öæ÷BFf–ÆRå7F'G5v—F‚‚G&ö÷BÂµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’Ö÷"Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚Ff–ÆR’’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡BCB…¶÷&FW&VEÔ²ö²ÒFfÇ6S²W'&÷"Òvæ÷Bf÷VæBrÒ“²&WGW&à¢Ð¢w&—FRÔ'—FW5&W7öç6RD6öçFW‡B#…´”òäf–ÆUÓ£¥&VDÆÄ'—FW2‚Ff–ÆR’’„vWBÔÖ–ÖRFf–ÆR§Ð  ¦gVæ7F–öâæ÷&ÖÆ—¦RÕv÷&·76U&VÆF—fUF‚…·7G&–æuÒE&VÆF—fUF‚’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E&VÆF—fUF‚’’²&WGW&ârrÐ¢G&VÂÒE&VÆF—fUF‚åG&–Ò‚¢–b‚G&VÂÖWwVæFVf–æVBrÖ÷"G&VÂÖWvçVÆÂr’²&WGW&ârrÐ¢–b…´”òåF…Ó£¤—5F…&ö÷FVB‚G&VÂ’’²F‡&÷ruDn898+ž8ÎKˆÞjÚ>8~8ž8"rÐ¢–b‚G&VÂÖÖF6‚r…çÅµÅÂõÒ•ÂåÂâ‚GÅµÅÂõÒ’r’²F‡&÷ruDn898+ž8ÎKˆÞjÚ>8~8ž8"rÐ¢–b‚G&VÂÖÖF6‚uµÇƒÕÇƒeÒr’²F‡&÷ruDn898+ž8ÎKˆÞjÚ>8~8ž8"rÐ¢–b…´”òåF…Ó£¤vWDW‡FVç6–öâ‚G&VÂ’åFôÆ÷vW$–çf&–çB‚’ÖæRrçFbr’²F‡&÷ruDn89^8*8*N8:¾88ŠŽzK®8~8Þ8î8ž8"rÐ¢&WGW&âG&VÀ§Ð  ¦gVæ7F–öâæ÷&ÖÆ—¦RÕ&VÆF—fTf÷$6ö×&R…·7G&–æuÒE&VÆF—fUF‚’°¢&WGW&â‚…·7G&–æuÒE&VÆF—fUF‚’×&WÆ6RuÅÂrÂròr’åG&–Ò‚’åFôÆ÷vW$–çf&–çB‚§Ð ¦gVæ7F–öâ6W'fRÔ6öçFVçEFd'•fÇVW2‚D6öçFW‡BÂ·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEvT–BÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6†VWDæÖRÂ·7G&–æuÒD6öçFVçEFb’°¢Gv÷&·76RÒvWBÕv÷&·76UF‚DÆæwVvP¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢GvT–BÒ…·7G&–æuÒEvT–B’åG&–Ò‚¢–b‚GvT–BÖWwVæFVf–æVBrÖ÷"GvT–BÖWvçVÆÂr’²GvT–BÒrrÐ¢Gv÷&¶&öö´–BÒ…·7G&–æuÒEv÷&¶&öö´–B’åG&–Ò‚¢–b‚Gv÷&¶&öö´–BÖWwVæFVf–æVBrÖ÷"Gv÷&¶&öö´–BÖWvçVÆÂr’²Gv÷&¶&öö´–BÒrrÐ¢G6†VWDæÖRÒ…·7G&–æuÒE6†VWDæÖR’åG&–Ò‚¢–b‚G6†VWDæÖRÖWwVæFVf–æVBrÖ÷"G6†VWDæÖRÖWvçVÆÂr’²G6†VWDæÖRÒrrÐ¢F6öçFVçE&VÂÒæ÷&ÖÆ—¦RÕv÷&·76U&VÆF—fUF‚D6öçFVçEF`¢GvRÒ‚ ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GvT–B’’°¢GvRÒ„vWBÔ'&’G7G'V7GW&RçvW2Âv†W&RÔö&¦V7B°¢…&W6öÇfRÕvT–BEò’ÖWGvT–BÖ÷"·7G&–æuÒEòçvT–BÖWGvT–BÖ÷"·7G&–æuÒEòæ–BÖWGvT–@¢ÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢Ð¢–b‚GvRä6÷VçBÖWÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Gv÷&¶&öö´–B’ÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6†VWDæÖR’’°¢GvRÒ„vWBÔ'&’G7G'V7GW&RçvW2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWGv÷&¶&öö´–BÖæB·7G&–æuÒEòç6†VWDæÖRÖWG6†VWDæÖRÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢Ð¢–b‚GvRä6÷VçBÖWÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öçFVçE&VÂ’’°¢F6×Òæ÷&ÖÆ—¦RÕ&VÆF—fTf÷$6ö×&RF6öçFVçE&VÀ¢GvRÒ„vWBÔ'&’G7G'V7GW&RçvW2Âv†W&RÔö&¦V7B²„æ÷&ÖÆ—¦RÕ&VÆF—fTf÷$6ö×&R…·7G&–æuÒEòæ6öçFVçEFb’’ÖWF6×ÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢Ð ¢G&VÂÒrp¢–b‚GvRä6÷VçBÖwB’²G&VÂÒ·7G&–æuÒGvU³Òæ6öçFVçEFbÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&VÂ’ÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öçFVçE&VÂ’’²G&VÂÒF6öçFVçE&VÂÐ¢G&VÂÒæ÷&ÖÆ—¦RÕv÷&·76U&VÆF—fUF‚G&VÀ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&VÂ’’²F‡&÷ruDnh8^Z8ÎKˆÞ‹k>8~8n8N8î8ž8.89®8;Î8+Žjx¾h‰8).i»Nik8~8n8¾8(•Dn8).™h¾8N8n8þ88^8N8"rÐ ¢FgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚„¦ö–âÕF‚Gv÷&·76RG&VÂ’¢Gv÷&·76TgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚Gv÷&·76R¢–b‚Öæ÷BGv÷&·76TgVÆÂäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’’²Gv÷&·76TgVÆÂ³Ò´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"Ð¢–b‚Öæ÷BFgVÆÂå7F'G5v—F‚‚Gv÷&·76TgVÆÂÂµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’’²F‡&÷r~8:þ8;Î8*þ8+ž89®8;Î8+žZIn8î89^8*8*N8:¾8þŠŽzK®8~8Þ8î8¾8)>8"rÐ¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FgVÆÂ’’²F‡&÷ruDn89^8*8*N8:¾8ÎŠh¾8N8¾8(®8î8¾8)>8%DnKÙÎh‰8).8(N8(®y»N8~8n8þ88^8N8"rÐ¢–b‚„vWBÔ—FVÒÔÆ—FW&ÅF‚FgVÆÂ’äÆVæwF‚ÖÆR’²F‡&÷ruDn89^8*8*N8:¾8Îz›®8~8ž8%DnKÙÎh‰8).8(N8(®y»N8~8n8þ88^8N8"rÐ¢w&—FRÔf–ÆU&W7öç6RD6öçFW‡B#FgVÆÂvÆ–6F–öâ÷FbrFfÇ6Rvæò×7F÷&Rp§Ð ¦gVæ7F–öâ&W6öÇfRÔ6öçFVçEFdgVÆÅF‚…·7G&–æuÒEv÷&·76RÂ·7G&–æuÒE&VÆF—fUFb’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E&VÆF—fUFb’’²F‡&÷rv6öçFVçB×Fb8ÎiÊ®KÙÎh‰8~8ž8"rÐ¢FW7BÕ&VÆF—fUF‚E&VÆF—fUFbÂ÷WBÔçVÆÀ¢–b…´”òåF…Ó£¤vWDW‡FVç6–öâ‚E&VÆF—fUFb’åFôÆ÷vW$–çf&–çB‚’ÖæRrçFbr’²F‡&÷ruDn89^8*8*N8:¾88ŠŽzK®8~8Þ8î8ž8"rÐ¢FgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚„¦ö–âÕF‚Ev÷&·76RE&VÆF—fUFb’¢Gv÷&·76TgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚Ev÷&·76R¢–b‚Öæ÷BGv÷&·76TgVÆÂäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’’²Gv÷&·76TgVÆÂ³Ò´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"Ð¢–b‚Öæ÷BFgVÆÂå7F'G5v—F‚‚Gv÷&·76TgVÆÂÂµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’’²F‡&÷r~8:þ8;Î8*þ8+ž89®8;Î8+žZIn8î89^8*8*N8:¾8þŠŽzK®8~8Þ8î8¾8)>8"rÐ¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FgVÆÂ’’²F‡&÷ruDn89^8*8*N8:¾8ÎŠh¾8N8¾8(®8î8¾8)>8%DnKÙÎh‰8).8(N8(®y»N8~8n8þ88^8N8"rÐ¢&WGW&âFgVÆÀ§Ð ¦gVæ7F–öâ6W'fRÔ6öçFVçEFb‚D6öçFW‡BÂ·7G&–æuÒDÆæwVvR’°¢GvT–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wvT–BuÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GvT–B’’²GvT–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v–BuÒÐ¢F6öçFVçEFbÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v6öçFVçEFbuÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öçFVçEFb’’²F6öçFVçEFbÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wFbuÒÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öçFVçEFb’’²F6öçFVçEFbÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wF‚uÒÐ¢6W'fRÔ6öçFVçEFd'•fÇVW2D6öçFW‡BDÆæwVvRGvT–B…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wv÷&¶&öö´–BuÒ’…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w6†VWDæÖRuÒ’F6öçFVçEF`§Ð ¦gVæ7F–öâ6W'fRÔ6öçFVçEFdg&öÔ&öG’‚D6öçFW‡BÂ·7G&–æuÒDÆæwVvRÂD&öG’’°¢GvT–BÒ·7G&–æuÒD&öG’çvT–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GvT–B’’²GvT–BÒ·7G&–æuÒD&öG’æ–BÐ¢F6öçFVçEFbÒ·7G&–æuÒD&öG’æ6öçFVçEF`¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öçFVçEFb’’²F6öçFVçEFbÒ·7G&–æuÒD&öG’çFbÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öçFVçEFb’’²F6öçFVçEFbÒ·7G&–æuÒD&öG’çF‚Ð¢6W'fRÔ6öçFVçEFd'•fÇVW2D6öçFW‡BDÆæwVvRGvT–B…·7G&–æuÒD&öG’çv÷&¶&öö´–B’…·7G&–æuÒD&öG’ç6†VWDæÖR’F6öçFVçEF`§Ð ¦gVæ7F–öâ6W'fRÔf–æÅFd'•föÇVÖR‚D6öçFW‡BÅ·7G&–æuÒDÆæwVvRÅ·7G&–æuÒEföÇVÖRÅ·7G&–æuÒD6FVv÷'’’°¢–b‚„vWBÕföÇVÖTÆ—7BDÆæwVvRÂv†W&RÔö&¦V7B²EòÖæRvæöæRrÒ’Öæ÷F6öçF–ç2EföÇVÖR’²F‡&÷rµ7—7FVÒä&wVÖVçDW†6WF–öåÓ£¦æWr‚wföÇVÖ^8¾8þiÊÎKÙ>8î8þ8þŠ9Î‹k>8).hÈ~Zé®8~8n8þ88^8N8"r’Ð¢F6CÕ&WV—&RÕv÷&¶&öö´6FVv÷'’D6FVv÷'“²GföÇVÖSÒ…·7G&–æuÒEföÇVÖR’åG&–Ò‚“¶–b„„vWBÕföÇVÖTÆ—7BDÆæwVvWÅv†W&RÔö&¦V7G²EòÖæRvæöæRwÒ’Öæ÷F6öçF–ç2GföÇVÖR—·F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~KˆÞjÚ>8®h‰iéÎxšž8~8ž8"r—Ó²G7G'V7GW&SÔvWBÕ7G'V7GW&RDÆæwVvS²GcÔvWBÔFF&÷W'G’G7G'V7GW&RçföÇVÖW2„vWBÕföÇVÖU7FFT¶W’GföÇVÖRF6B’FçVÆÃ¶–b‚FçVÆÂÖWGbÖ÷"Öæ÷B·7G&–æuÒGbæ÷WGWEFb—·F‡&÷r~hùX{®yJ…Dn8þ8î8KÙÎh‰8^8(Î8n8N8î8¾8)>8"wÓ²GF‡3ÔvWBÕF‡3²FgVÆÃÕ´”òåF…Ó£¤vWDgVÆÅF‚…·7G&–æuÒGbæ÷WGWEFb“²G&ö÷CÕ´”òåF…Ó£¤vWDgVÆÅF‚…·7G&–æuÒGF‡2æ÷WGWDF—"“¶–b‚Öæ÷BG&ö÷BäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’—²G&ö÷B³Õ´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†'Ó¶–b‚Öæ÷BFgVÆÂå7F'G5v—F‚‚G&ö÷BÅµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’—·F‡&÷r~X{®X©¾89^8*ž8:¾88ZIn8åDn8þŠŽzK®8~8Þ8î8¾8)>8"wÓ¶–b‚Öæ÷B…FW7BÕF‚FgVÆÂ’—·F‡&÷r~hùX{®yJ…Dn89^8*8*N8:¾8ÎŠh¾8N8¾8(®8î8¾8)>8"wÓµw&—FRÔ'—FW5&W7öç6RD6öçFW‡B#…´”òäf–ÆUÓ£¥&VDÆÄ'—FW2‚FgVÆÂ’’vÆ–6F–öâ÷Fbp§Ð ¦gVæ7F–öâ6W'fRÔf–æÅFb‚D6öçFW‡BÅ·7G&–æuÒDÆæwVvR’°¢6W'fRÔf–æÅFd'•föÇVÖRD6öçFW‡BDÆæwVvR…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wföÇVÖRuÒ’…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v6FVv÷'’uÒ§Ð ¦gVæ7F–öâ6W'fRÔFö7VÖVçE6µFb‚D6öçFW‡BÂ·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–BÂ·7G&–æuÒEF&vWD–B’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG7G'V7GW&RE6´–BGG'VP¢EF&vWD–BÒ76W'BÕ6µF&vWD–BDÆæwVvRG66÷Rç6²EF&vWD–@¢GföÇVÖRÒvWBÔÆVv7•föÇVÖTg&öÕF&vWD–BDÆæwVvREF&vWD–@¢G7FFRÒvWBÕ6´÷WGWE7FFRG7G'V7GW&RDÆæwVvR…·7G&–æuÒG66÷Rç6´–B’GföÇVÖRFfÇ6P¢F÷WGWBÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRv÷WGWEFbrrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F÷WGWB’’²F‡&÷r~hùX{®yJ…Dn8þ8î8KÙÎh‰8^8(Î8n8N8î8¾8)>8"rÐ¢GF‡2ÒvWBÕF‡0¢FgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚F÷WGWB¢G&ö÷BÒ´”òåF…Ó£¤vWDgVÆÅF‚…·7G&–æuÒGF‡2æ÷WGWDF—"¢–b‚Öæ÷BG&ö÷BäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’’²G&ö÷B³Ò´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"Ð¢–b‚Öæ÷BFgVÆÂå7F'G5v—F‚‚G&ö÷BÅµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’’²F‡&÷r~X{®X©¾89^8*ž8:¾88ZIn8åDn8þŠŽzK®8~8Þ8î8¾8)>8"rÐ¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FgVÆÂ’’²F‡&÷r~hùX{®yJ…Dn89^8*8*N8:¾8ÎŠh¾8N8¾8(®8î8¾8)>8"rÐ¢w&—FRÔf–ÆU&W7öç6RD6öçFW‡B#FgVÆÂvÆ–6F–öâ÷Fbp§Ð ¢2X{®X©¾8~8õDn8îKùÞZÙŽXXŽ8).8*Ž8*þ8+ž89~8:Þ8;Î8:ž8;Î8~™h¾8Þ88Þ8î89^8*8*N8:¾8).˜Žh©îx«nhX¾8¾8ž8(¾8 ¢289n8:ž8*n8+n8îXŠ^8+þ89n8~8þKŠÞ‹ª¾8).Šh¾8(ž8(Î8(¾888~8k{¾K¹Ž8(N8+>89N8;Î8î8þ8(8¾ZéþKÙ>8).hëN8(8®8N8þ8(8 ¢2X{®X©¾89^8*ž8:¾888;Î8îZIn8þ™h¾8¾8®8BŽ898+ž8ò7G'V7GW&RyKiÚ^88Î8z+®Š¨Þ8ò6W'fR8ŽYÎ8Ž[Ú.8~ŠÎ8bž8 ¦gVæ7F–öâ&W6öÇfRÔ÷WGWEFdf÷%&WfVÂ…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´÷$6FVv÷'’Â·7G&–æuÒEF&vWD–BÂ·7G&–æuÒEföÇVÖR’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢2&W6öÇfRÔFö7VÖVçE6µ66÷R8ò6´–B8‚6FVv÷'’8î8ž88(ž8~8(.Xù~88(¾8.{XN8þ‹ëÎ8þ8988>8*þ8ð¢2yK¾™Ú.8Â6FVv÷'’8‚föÇVÖR8).˜8(¾8þ8(8F&vWD–B8ÎxJ8N8Ž8Þ8òföÇVÖR8).8Þ8î8î8îKÛþ8n8 ¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG7G'V7GW&RE6´÷$6FVv÷'’GG'VP¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EF&vWD–B’’°¢G&W6öÇfVEF&vWBÒ76W'BÕ6µF&vWD–BDÆæwVvRG66÷Rç6²EF&vWD–@¢GföÇVÖRÒvWBÔÆVv7•föÇVÖTg&öÕF&vWD–BDÆæwVvRG&W6öÇfVEF&vW@¢ÒVÇ6R°¢GföÇVÖRÒ…·7G&–æuÒEföÇVÖR’åG&–Ò‚¢–b„„vWBÕföÇVÖTÆ—7BDÆæwVvRÂv†W&RÔö&¦V7B²EòÖæRvæöæRrÒ’Öæ÷F6öçF–ç2GföÇVÖR’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~KˆÞjÚ>8®X{®X©¾XXŽ8~8ž8"r’Ð¢Ð¢G7FFRÒvWBÕ6´÷WGWE7FFRG7G'V7GW&RDÆæwVvR…·7G&–æuÒG66÷Rç6´–B’GföÇVÖRFfÇ6P¢F÷WGWBÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRv÷WGWEFbrrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F÷WGWB’’²F‡&÷r~hùX{®yJ…Dn8þ8î8KÙÎh‰8^8(Î8n8N8î8¾8)>8"rÐ¢GF‡2ÒvWBÕF‡0¢FgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚F÷WGWB¢G&ö÷BÒ´”òåF…Ó£¤vWDgVÆÅF‚…·7G&–æuÒGF‡2æ÷WGWDF—"¢–b‚Öæ÷BG&ö÷BäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’’²G&ö÷B³Ò´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"Ð¢–b‚Öæ÷BFgVÆÂå7F'G5v—F‚‚G&ö÷BÅµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’’²F‡&÷r~X{®X©¾89^8*ž8:¾888;Î8îZIn8¾8.8(µDn8þ™h¾88î8¾8)>8"rÐ¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FgVÆÂ’’²F‡&÷r~hùX{®yJ…Dn89^8*8*N8:¾8ÎŠh¾8N8¾8(®8î8¾8)>8.X{®X©¾89^8*ž8:¾888;Î8¾8(žz{¾X¹^8î8þ8þX˜®™šN8^8(Î8þXúþˆ;Þh
~8Î8.8(®8î8ž8"rÐ¢&WGW&âFgVÆÀ§Ð ¦gVæ7F–öâ÷VâÔ÷WGWEFdÆö6F–öâ…·7G&–æuÒDgVÆÅF‚’°¢2ZéþkŠÂƒ##bÓ‚Óž8~z+®8¾8(8þ[Ú.8.8ž88(ž8).[Jž8~8n8(.8*Ž8*þ8+ž89~8:Þ8;Î8:ž8;Î8þKÙ^8(.™h¾8¾8®8N8 ¢2Ò÷6VÆV7F8îy»N[èÎ8î8*¾8;>89î8þ[ø^šŽ8.z›®y›Þ8¾8ž8(¾8Žz©>8ÎX{®8®8N8 ¢2Ò898+ž8î[É^yJŽzÊn8(.[ø^šŽ8.X{®X©¾89^8*8*N8:¾YÞ8þXŠžyJŽˆ^8Î{zŽ™¸n8~8Þ8(¾898+þ8;Î8;0¢2‡·6´æÖWÕ÷·F&vWDæÖWÕ÷·———”ÔÖFGÒçFbž8¾8(žKÙÎ8(ž8(Î8(¾8þ8(8*¾8;>89î8).Y
¾8þ8n8(¾8 ¢2[É^yJŽ8~8®8N8Ž88*¾8;>89î8).Y
¾8(YÞX˜Þ8~z©>8ÎX{®8®8N8 ¢F&wVÖVçBÒr÷6VÆV7BÂ"r²DgVÆÅF‚²r"p¢·fö–EÕ´F–væ÷7F–72å&ö6W75Ó£¥7F'B‚„æWrÔö&¦V7BF–væ÷7F–72å&ö6W757F'D–æfòÕ&÷W'G’°¢f–ÆTæÖRÒvW‡Æ÷&W"æW†Rp¢&wVÖVçG2ÒF&wVÖVç@¢W6U6†VÆÄW†V7WFRÒGG'VP¢Ò’§Ð ¦gVæ7F–öâvWBÕ6fUV&Æ—6…W6W$æÖR°¢FæÖRÒ…·7G&–æuÒFVçc¥U4U$äÔR’åG&–Ò‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’’²FæÖRÒwW6W"rÐ¢FæÖRÒ·&VvW…Ó£¥&WÆ6R‚FæÖRÂu³Ãã¢"õÅÇÃò¥ÇƒÕÇƒeÒ²rÂuòr¢FæÖRÒ·&VvW…Ó£¥&WÆ6R‚FæÖRÂuÇ2²rÂuòr’åG&–Ò…¶6†%µÕÔ‚uòrÂrâr’¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’’²FæÖRÒwW6W"rÐ¢–b‚FæÖRäÆVæwF‚ÖwBC’²FæÖRÒFæÖRå7V'7G&–ærƒÃC’Ð¢&WGW&âFæÖP§Ð ¦gVæ7F–öâvWBÕ6fUV&Æ—6…6´æÖR…·7G&–æuÒDF—7Æ”æÖR’°¢FæÖRÒ…·7G&–æuÒDF—7Æ”æÖR’åG&–Ò‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’’²FæÖRÒvFö7VÖVçB×6²rÐ¢FæÖRÒ·&VvW…Ó£¥&WÆ6R‚FæÖRÂu³Ãã¢"õÅÇÃò¥ÇƒÕÇƒeÒ²rÂuòr¢FæÖRÒ·&VvW…Ó£¥&WÆ6R‚FæÖRÂuÇ2²rÂuòr’åG&–Ò…¶6†%µÕÔ‚uòrÂrâr’¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’’²FæÖRÒvFö7VÖVçB×6²rÐ¢–b‚FæÖRäÆVæwF‚ÖwBC‚’²FæÖRÒFæÖRå7V'7G&–ærƒÃC‚’Ð¢&WGW&âFæÖP§Ð ¦gVæ7F–öâV&Æ—6‚ÔFö7VÖVçE6µFeFõ6†&VB…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEföÇVÖRÂ·7G&–æuÒE6´–D÷$6FVv÷'’’°¢GF‡2ÒvWBÕF‡0¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG7G'V7GW&RE6´–D÷$6FVv÷'’FfÇ6P¢–b„„vWBÕ6µföÇVÖTÆ—7BDÆæwVvRG66÷Rç6²FfÇ6R’Öæ÷F6öçF–ç2EföÇVÖR’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~X[iÈžy›®ŠÎ8ž8(µDn8îzŠîšî8ÎKˆÞjÚ>8~8ž8"r’Ð¢G6´–BÒ·7G&–æuÒG66÷Rç6´–@¢F6BÒ·7G&–æuÒG66÷Ræ6FVv÷'¢G&VG’ÒvWBÔf–æÄ'V–ÆE&VF–æW72G7G'V7GW&RDÆæwVvREföÇVÖRG6´–@¢–b…·7G&–æuÒG&VG’æF—7Æ•7FFRÖæRv'V–ÇBr’°¢F‡&÷r~hùX{®yJ…Dn8ÎiÈik8~8þ8.8(®8î8¾8)>8.iÈikx«nhX¾8~XhÞX{®X©¾8~8n8¾8(žX[iÈžy›®ŠÎ8~8n8þ88^8N8"p¢Ð¢G6÷W&6RÒ·7G&–æuÒG&VG’æ÷WGWEF`¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6÷W&6R’Ö÷"Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G6÷W&6RÕF…G—RÆVb’’°¢F‡&÷r~X[iÈžy›®ŠÎ8~8Þ8(¾hùX{®yJ…Dn8Î8.8(®8î8¾8)>8.XXŽ8¾hùX{®yJ…Dn8).X{®X©¾8~8n8þ88^8N8"p¢Ð ¢FÆö6Ä÷WGWE&ö÷BÒ´”òåF…Ó£¤vWDgVÆÅF‚…·7G&–æuÒGF‡2æ÷WGWDF—"¢–b‚Öæ÷BFÆö6Ä÷WGWE&ö÷BäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’’²FÆö6Ä÷WGWE&ö÷B³Ò´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"Ð¢G6÷W&6TgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚G6÷W&6R¢–b‚Öæ÷BG6÷W&6TgVÆÂå7F'G5v—F‚‚FÆö6Ä÷WGWE&ö÷BÂµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’’°¢F‡&÷r~X[iÈžy›®ŠÎXX>8ÎXŠžyJŽˆ^8:Þ8;Î8*¾8:¾8îX{®X©¾89^8*ž8:¾888;ÎZIn8~8ž8.hùX{®yJ…Dn8).XhÞX{®X©¾8~8n8þ88^8N8"p¢Ð ¢G7V&Ö—76–öäF—"Ò´”òåF…Ó£¤vWDgVÆÅF‚…·7G&–æuÒGF‡2ç7V&Ö—76–öäF—"¢GW6W$æÖRÒvWBÕ6fUV&Æ—6…W6W$æÖP¢G6´æÖRÒvWBÕ6fUV&Æ—6…6´æÖR…·7G&–æuÒ„vWBÔFF&÷W'G’G66÷Rç6²vF—7Æ”æÖRrG6´–B’¢G7F×ÒvWBÔFFRÔf÷&ÖBtÔÖFEô„†Ö×72p¢F&6TföÆFW$æÖRÒ'³Õ÷³Õ÷³'Ò"ÖbG7F×ÂG6´æÖRÂGW6W$æÖP¢FföÆFW$æÖRÒF&6TföÆFW$æÖP¢GV&Æ—6„F—"Ò¦ö–âÕF‚G7V&Ö—76–öäF—"FföÆFW$æÖP¢G7Vff—‚Ò ¢v†–ÆR…FW7BÕF‚ÔÆ—FW&ÅF‚GV&Æ—6„F—"’°¢FföÆFW$æÖRÒ'³Õ÷³Ò"ÖbF&6TföÆFW$æÖRÂG7Vff—€¢GV&Æ—6„F—"Ò¦ö–âÕF‚G7V&Ö—76–öäF—"FföÆFW$æÖP¢G7Vff—‚²°¢Ð ¢2Dn8ÎXØ®zºþ8®x«nhX¾8~X[iÈžXN8¾Šh¾8Ž8®8N8(Ž8n8YÎ8ŽX[iÈž8:¾8;Î88Ž8î™ª8~Kˆi˜.89^8*ž8:¾888;Î8p¢28+>89N8;Î8ŽjIÎŠ‹Î8).ZèÎK¨n8~8n8¾8(ž8y›®ŠÎ89^8*ž8:¾888;ÎYÞ8ŽKˆ[ªn88Xˆ~8(®i»þ8Ž8(¾8 ¢G7Fv–ætF—"Ò¦ö–âÕF‚G7V&Ö—76–öäF—"‚rçV&Æ—6†–ærÒr²´wV–EÓ£¤æWtwV–B‚’åFõ7G&–ær‚târ’¢Ff–ÆTæÖRÒ´”òåF…Ó£¤vWDf–ÆTæÖR‚G6÷W&6TgVÆÂ¢G7FvVEFbÒ¦ö–âÕF‚G7Fv–ætF—"Ff–ÆTæÖP¢G'’°¢æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚G7Fv–ætF—"Ôf÷&6RÂ÷WBÔçVÆÀ¢6÷’Ô—FVÒÔÆ—FW&ÅF‚G6÷W&6TgVÆÂÔFW7F–æF–öâG7FvVEFbÔf÷&6P¢G6÷W&6TÆVæwF‚Ò„vWBÔ—FVÒÔÆ—FW&ÅF‚G6÷W&6TgVÆÂÔW'&÷$7F–öâ7F÷’äÆVæwF€¢GWÆöFVDÆVæwF‚Ò„vWBÔ—FVÒÔÆ—FW&ÅF‚G7FvVEFbÔW'&÷$7F–öâ7F÷’äÆVæwF€¢–b‚G6÷W&6TÆVæwF‚ÖÆRÖ÷"GWÆöFVDÆVæwF‚ÖæRG6÷W&6TÆVæwF‚’°¢F‡&÷r~X[iÈž89^8*ž8:¾888;Î8Ž8î8+>89N8;Î8+^8*N8+®8ÎKˆˆ{N8~8î8¾8)>8"p¢Ð¢v†–ÆR…FW7BÕF‚ÔÆ—FW&ÅF‚GV&Æ—6„F—"’°¢FföÆFW$æÖRÒ'³Õ÷³Ò"ÖbF&6TföÆFW$æÖRÂG7Vff—€¢GV&Æ—6„F—"Ò¦ö–âÕF‚G7V&Ö—76–öäF—"FföÆFW$æÖP¢G7Vff—‚²°¢Ð¢Ö÷fRÔ—FVÒÔÆ—FW&ÅF‚G7Fv–ætF—"ÔFW7F–æF–öâGV&Æ—6„F— ¢G7Fv–ætF—"Òrp¢Òf–æÆÇ’°¢–b‚G7Fv–ætF—"ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚G7Fv–ætF—"’’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚G7Fv–ætF—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢Ð¢Ð¢FFW7F–æF–öâÒ¦ö–âÕF‚GV&Æ—6„F—"Ff–ÆTæÖP ¢GV&Æ—6†VDBÒæWrÔæ÷t—6ð¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvf–æÂçV&Æ—6†VBr…¶÷&FW&VEÔ°¢6´–BÒG6´–@¢6FVv÷'’ÒF6@¢F&vWD–BÒvWBÕF&vWD–Dg&öÔÆVv7•föÇVÖREföÇVÖP¢föÇVÖRÒEföÇVÖP¢6÷W&6RÒG6÷W&6TgVÆÀ¢FW7F–æF–öâÒFFW7F–æF–öà¢6†&VDföÆFW"ÒGV&Æ—6„F— ¢föÆFW$æÖRÒFföÆFW$æÖP¢f–ÆTæÖRÒFf–ÆTæÖP¢W6W$æÖRÒGW6W$æÖP¢6´æÖRÒG6´æÖP¢V&Æ—6†VDBÒGV&Æ—6†VD@¢Ò¢&WGW&â¶÷&FW&VEÔ°¢föÇVÖRÒEföÇVÖP¢F&vWD–BÒvWBÕF&vWD–Dg&öÔÆVv7•föÇVÖREföÇVÖP¢6´–BÒG6´–@¢6FVv÷'’ÒF6@¢6÷W&6UFbÒG6÷W&6TgVÆÀ¢6†&VEF‚ÒFFW7F–æF–öà¢6†&VDföÆFW"ÒGV&Æ—6„F— ¢föÆFW$æÖRÒFföÆFW$æÖP¢f–ÆTæÖRÒFf–ÆTæÖP¢V&Æ—6†VD'’ÒGW6W$æÖP¢6´æÖRÒG6´æÖP¢V&Æ—6†VDBÒGV&Æ—6†VD@¢Ð§Ð ¦gVæ7F–öâV&Æ—6‚Ôf–æÅFeFõ6†&VB…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEföÇVÖRÂ·7G&–æuÒD6FVv÷'’’°¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’D6FVv÷'¢&WGW&âV&Æ—6‚ÔFö7VÖVçE6µFeFõ6†&VBDÆæwVvREföÇVÖR„vWBÔ'V–ÇF–å6´–BF6B§Ð ¦gVæ7F–öâ†æFÆRÔ’‚D6öçFW‡B’°¢FÆæwVvRÒvWBÔVffV7F—fTÆæwVvP¢GF‚ÒD6öçFW‡Bå&WVW7BåW&Âä'6öÇWFUF€¢FÖWF†öBÒD6öçFW‡Bå&WVW7Bä‡GGÖWF†öBåFõWW$–çf&–çB‚¢G'’°¢2F†W6RÆ–v‡GvV–v‡BVæGö–çG2&RöæÇ’f÷"7F'GW&VF–æW726†V6·2à¢2F†W’–çFVçF–öæÆÇ’Fòæ÷B&WV—&R6W76–öâFö¶Vâ6òF†RÆö6Âv—BvP¢26âFWFV7B&VF–æW72WfVâ–bF†R'&÷w6W"7G&—2÷"FVÆ—2VW'’†æFÆ–ærà¢–b‚FÖWF†öBÖWtõD”ôå2r’°¢w&—FRÕFW‡E&W7öç6RD6öçFW‡B#BrrwFW‡B÷Æ–ã²6†'6WC×WFbÓ‚s²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷&VG’æv–br’°¢w&—FRÔ'—FW5&W7öç6RD6öçFW‡B#E67&—C¥&VG”v–d'—FW2v–ÖvRöv–brGG'VS²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷–ærr’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²'VçF–ÖUfW'6–öâÒE67&—C¥'VçF–ÖUfW'6–öã²BÒæWrÔæ÷t—6òÒ’GG'VS²&WGW&à¢Ð¢–b‚Öæ÷B…FW7BÕ&WVW7D÷&–v–âD6öçFW‡Bå&WVW7B’’²w&—FRÔ§6öå&W7öç6RD6öçFW‡BC2…¶÷&FW&VEÔ²ö²ÒFfÇ6S²W'&÷"Òv–çfÆ–B÷&–v–ârÒ“²&WGW&âÐ¢–b‚Öæ÷B…FW7BÕFö¶VâD6öçFW‡Bå&WVW7B’’²w&—FRÔ§6öå&W7öç6RD6öçFW‡BC2…¶÷&FW&VEÔ²ö²ÒFfÇ6S²W'&÷"Òv–çfÆ–BFö¶VârÒ“²&WGW&âÐ¢F÷V6‚Ô6Æ–VçD7F—f—G’rrÂ÷WBÔçVÆÀ¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’ö†V'F&VBr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…F÷V6‚Ô6Æ–VçD7F—f—G’…·7G&–æuÒF&öG’æ6Æ–VçD–B’“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’ö6Æ–VçBö6Æ÷6Rr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#„æ÷F–g’Ô6Æ–VçD6Æ÷6–ær…·7G&–æuÒF&öG’æ6Æ–VçD–B’“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷6‡WFF÷vâr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…&WVW7BÕ6W'fW%6‡WFF÷vâ…·7G&–æuÒF&öG’ç&V6öâ’FÆæwVvR“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öWFò÷&VæFW"r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F–G2Ò‚¢–b‚F&öG’çv÷&¶&öö´–G2’²F–G2Ò„vWBÔ'&’F&öG’çv÷&¶&öö´–G2Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ’Ð¢G&W7VÇBÒ–çfö¶RÔWFõ&VæFW"FÆæwVvRF–G0¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG&W7VÇBÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"÷7FFRr’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#„vWBÕc%7FFU–ÆöBFÆæwVvR“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öF–væ÷7F–72÷'Vâr’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²F–væ÷7F–73Ò„vWBÕ7—7FVÔF–væ÷7F–72FÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"÷6²×FV×ÆFW2r’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²6µFV×ÆFW2Ò„vWBÕ6µFV×ÆFT6FÆörFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"÷6²×FV×ÆFW2r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&W7VÇBÒ6fRÕ6µFV×ÆFRFÆæwVvRF&öG¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²6µFV×ÆFW2Ò„vWBÕ6µFV×ÆFT6FÆörFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuD4‚rÖæBGF‚ÖÖF6‚uâö’÷c"÷6²×FV×ÆFW2ò…µâõÒ²’Br’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢GFV×ÆFT–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢G&W7VÇBÒ6fRÕ6µFV×ÆFRFÆæwVvRF&öG’GFV×ÆFT–@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²6µFV×ÆFW2Ò„vWBÕ6µFV×ÆFT6FÆörFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWtDTÄUDRrÖæBGF‚ÖÖF6‚uâö’÷c"÷6²×FV×ÆFW2ò…µâõÒ²’Br’°¢GFV×ÆFT–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢G&W7VÇBÒ&VÖ÷fRÕ6µFV×ÆFRFÆæwVvRGFV×ÆFT–@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²6µFV×ÆFW2Ò„vWBÕ6µFV×ÆFT6FÆörFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"÷6·2r’°¢G6µF‡2ÒvWBÕF‡0¢G6´FFF—"Ò·7G&–æuÒ„vWBÔFF&÷W'G’G6µF‡2vFFF—"rrr¢G7G'V7GW&RÒ–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6´FFF—"’ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚G6´FFF—"’’²vWBÕ7G'V7GW&RFÆæwVvRÒVÇ6R²æWrÔV×G•7G'V7GW&RFÆæwVvRÐ¢F–æ6ÇVFT&6†—fVBÒ…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v–æ6ÇVFT&6†—fVBuÒ’åG&–Ò‚’åFôÆ÷vW$–çf&–çB‚’Ö–â‚srÂwG'VRrÂw–W2r¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²6·2Ò„vWBÕV&Æ–56´Æ—7BG7G'V7GW&RFÆæwVvRF–æ6ÇVFT&6†—fVB’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"÷6·2r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&W7VÇBÒæWrÔFö7VÖVçE6²FÆæwVvRF&öG¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²6·2Ò„vWBÕV&Æ–56´Æ—7B„vWBÕ7G'V7GW&RFÆæwVvR’FÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖÖF6‚uâö’÷c"÷6·2ò…µâõÒ²’öGWÆ–6FRBr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6´–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢G&W7VÇBÒ6÷’ÔFö7VÖVçE6²FÆæwVvRG6´–BF&öG¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²6·2Ò„vWBÕV&Æ–56´Æ—7B„vWBÕ7G'V7GW&RFÆæwVvR’FÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖÖF6‚uâö’÷c"÷6·2ò…µâõÒ²’÷FV×ÆFR×Ww&FRBr’°¢G6´–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&Wf–WrÒ„vWBÕ6µFV×ÆFUWw&FU&Wf–WrFÆæwVvRG6´–B’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖÖF6‚uâö’÷c"÷6·2ò…µâõÒ²’÷FV×ÆFR×Ww&FRBr’°¢G6´–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢G&W7VÇBÒWFFRÕ6µFV×ÆFU6æ6†÷BFÆæwVvRG6´–@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²6·2Ò„vWBÕV&Æ–56´Æ—7B„vWBÕ7G'V7GW&RFÆæwVvR’FÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖÖF6‚uâö’÷c"÷6·2ò…µâõÒ²’ö&6†—fRBr’°¢G6´–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢G&W7VÇBÒ6WBÔFö7VÖVçE6´&6†—fVBFÆæwVvRG6´–BGG'VP¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²6·2Ò„vWBÕV&Æ–56´Æ—7B„vWBÕ7G'V7GW&RFÆæwVvR’FÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖÖF6‚uâö’÷c"÷6·2ò…µâõÒ²’÷&W7F÷&RBr’°¢G6´–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢G&W7VÇBÒ6WBÔFö7VÖVçE6´&6†—fVBFÆæwVvRG6´–BFfÇ6P¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²6·2Ò„vWBÕV&Æ–56´Æ—7B„vWBÕ7G'V7GW&RFÆæwVvR’FÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖÖF6‚uâö’÷c"÷6·2ò…µâõÒ²’÷&Wf–WrBr’°¢G6´–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&Wf–WsÒ„vWBÕ6µ&Wf–Wu6æ6†÷B„vWBÕ7G'V7GW&RFÆæwVvR’FÆæwVvRG6´–BGG'VR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖÖF6‚uâö’÷c"÷6·2ò…µâõÒ²’÷&Wf–WrBr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6´–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&Wf–WsÒ„–çfö¶RÕ6µ&Wf–Wt7F–öâFÆæwVvRG6´–BF&öG’“²7FFSÒ„vWBÕ7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuD4‚rÖæBGF‚ÖÖF6‚uâö’÷c"÷6·2ò…µâõÒ²’Br’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6´–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢G&W7VÇBÒWFFRÕ6µ6WGF–æw2FÆæwVvRG6´–BF&öG¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²7FFRÒ„vWBÕ7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"÷6÷W&6RÖ6æF–FFW2r’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²66ææVDBÒ„æWrÔæ÷t—6ò“²6æF–FFW2Ò„vWBÕ6÷W&6T6æF–FFW2‚vW†6VÂrÂwv÷&BrÂwFbrÂw÷vW'ö–çBr’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"÷6÷W&6W2÷&Vv—7FW"Ö&F6‚r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&VÆF—fUF‡2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F&öG’w&VÆF—fUF‡2r‚’’Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ¢–b‚G&VÆF—fUF‡2ä6÷VçBÖW’²F‡&÷r~y›¾˜Ë.8ž8(¾Xéþz‹þ8Î˜Žh©î8^8(Î8n8N8î8¾8)>8"rÐ¢G6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6´–B’’²G6´–BÒvWBÔ'V–ÇF–å6´–B…&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v6FVv÷'’rrr’’’Ð¢G&W7VÇBÒ&Vv—7FW"Õ6÷W&6W4&F6‚FÆæwVvRG&VÆF—fUF‡2G6´–B…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6÷W&6UG—Rrrr’¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²7FFRÒ„vWBÕc%7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"÷6÷W&6W2÷Vç&Vv—7FW"r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6÷W&6T–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6÷W&6T–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6÷W&6T–B’’²F‡&÷rw6÷W&6T–B8Î[ø^Šh8~8ž8"rÐ¢G&W7VÇBÒVç&Vv—7FW"Õv÷&¶&öö²FÆæwVvRG6÷W&6T–@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²7FFRÒ„vWBÕc%7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"÷6÷W&6W2÷66â×WFFW2r’°¢G&W7VÇBÒ66âÕWFFW2FÆæwVvP¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²&W7VÇBÒG&W7VÇC²7FFRÒ„vWBÕc%7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"÷6÷W&6W2÷&VæFW"÷7F'Br’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6÷W&6T–G2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F&öG’w6÷W&6T–G2r‚’’Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ¢G&VæFW%6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Brrr¢–b‚G6÷W&6T–G2ä6÷VçBÖWÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&VæFW%6´–B’’°¢G&VæFW%7G'V7GW&RÒvWBÕ7G'V7GW&RFÆæwVvP¢G&VæFW%66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG&VæFW%7G'V7GW&RG&VæFW%6´–BFfÇ6P¢G6÷W&6T–G2Ò„vWBÔ'&’G&VæFW%7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²FW7BÕv÷&¶&ööµ6²Eò…·7G&–æuÒG&VæFW%66÷Rç6´–B’ÒÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÒ¢Ð¢G&W7VÇBÒ7F'BÕ&VæFW$¦ö"FÆæwVvRG6÷W&6T–G2…¶&ööÅÒ„vWBÔFF&÷W'G’F&öG’vöæÇ•WFFVBrFfÇ6R’’…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v6FVv÷'’rrr’¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²•fW'6–öâÒ#²¦ö$–BÒ·7G&–æuÒG&W7VÇBæ¦ö$–C²¦ö"ÒG&W7VÇBÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuD4‚rÖæBGF‚ÖÖF6‚uâö’÷c"÷6÷W&6W2ò…µâõÒ²’Br’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6÷W&6T–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢GWFFVBÒWFFRÕ6÷W&6TÖWFFFFÆæwVvRG6÷W&6T–BF&öG¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²6÷W&6RÒGWFFVC²7FFRÒ„vWBÕ7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"ö—FV×2÷&V÷&FW"r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢GföÇVÖW2ÒvWBÔFF&÷W'G’F&öG’wföÇVÖW2rFçVÆÀ¢–b‚FçVÆÂÖWGföÇVÖW2’°¢GF&vWG2ÒvWBÔFF&÷W'G’F&öG’wF&vWG2rFçVÆÀ¢–b‚FçVÆÂÖWGF&vWG2’²F‡&÷rwF&vWG28Î[ø^Šh8~8ž8"rÐ¢GföÇVÖTÖÒ¶÷&FW&VEÔ·Ð¢f÷&V6‚‚G&÷W'G’–âGF&vWG2å4ö&¦V7Bå&÷W'F–W2’°¢GF&vWD–BÒ…·7G&–æuÒG&÷W'G’äæÖR’åG&–Ò‚’åFôÆ÷vW$–çf&–çB‚¢GföÇVÖRÒ–b‚GF&vWD–BÖ–â‚væöæRrÂwVæ76–væVBr’’²væöæRrÒVÇ6R²vWBÔÆVv7•föÇVÖTg&öÕF&vWD–BFÆæwVvRGF&vWD–BÐ¢GföÇVÖTÖ²GföÇVÖUÒÒ„vWBÔ'&’G&÷W'G’åfÇVR¢Ð¢GföÇVÖW2Ò·67W7FöÖö&¦V7EÒGföÇVÖTÖ ¢Ð¢G&WVW7BÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²6´–CÕ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Br„vWBÔFF&÷W'G’F&öG’v6FVv÷'’rrr’“²föÇVÖW3ÒGföÇVÖW2Ð¢2¶VWF†R÷F–Ö—7F–2ÖÆö6²f–ævW'&–çB7WÆ–VB'’F†R'&÷w6W"à¢2&V'V–ÆF–ærF†R&WVW7Bv—F†÷WB—B6–ÆVçFÇ’F—6&ÆVB6öæfÆ–7@¢2FWFV7F–öâöâF†Rc"&÷WFRæBÆÆ÷vVB7FÆRF"Fò÷fW'w&—FR¢2æWvW"vR6ö×÷6—F–öâà¢–b…FW7BÔ6öæf–t†4¶W’F&öG’v&6TÆ–÷WBr’²6WBÔæ÷FU&÷W'G’G&WVW7Bv&6TÆ–÷WBr…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v&6TÆ–÷WBrrr’’Ð¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&W7VÇCÒ…&V÷&FW"ÕvW2FÆæwVvRG&WVW7B“²7FFSÒ„vWBÕc%7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuD4‚rÖæBGF‚ÖÖF6‚uâö’÷c"ö—FV×2ò…µâõÒ²’Br’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F—FVÔ–BÒµW&•Ó£¥VæW66TFF7G&–ær…·7G&–æuÒFÖF6†W5³Ò¢G&WVW7BÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²6´–CÕ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Br„vWBÔFF&÷W'G’F&öG’v6FVv÷'’rrr’“²vT–CÒF—FVÔ–BÐ¢–b…FW7BÔ6öæf–t†4¶W’F&öG’v&6TÆ–÷WBr’²6WBÔæ÷FU&÷W'G’G&WVW7Bv&6TÆ–÷WBr…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v&6TÆ–÷WBrrr’’Ð¢f÷&V6‚‚FæÖR–â‚wF—FÆRrÂvçVÖ&W&–ætÖöFRrÂvçVÖ&W&–ætÖçVÂrÂw&W6WDçVÖ&W&–ærrÂwvU&ævRrÂv6ÆV%vU&ævRrÂvVæ&ÆVBr’’²–b…FW7BÔ6öæf–t†4¶W’F&öG’FæÖR’²6WBÔæ÷FU&÷W'G’G&WVW7BFæÖR„vWBÔFF&÷W'G’F&öG’FæÖRFçVÆÂ’ÒÐ¢–b…FW7BÔ6öæf–t†4¶W’F&öG’wF&vWD–Br’²6WBÔæ÷FU&÷W'G’G&WVW7BwföÇVÖRr„vWBÔÆVv7•föÇVÖTg&öÕF&vWD–BFÆæwVvR…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’wF&vWD–Brrr’’’Ð¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&W7VÇCÒ…WFFRÕvRFÆæwVvRG&WVW7B“²7FFSÒ„vWBÕc%7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"öÆ–÷WB÷6æ6†÷G2r’°¢G6´–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w6´–BuÐ¢G7G'V7GW&RÒvWBÕ7G'V7GW&RFÆæwVvP¢G66÷RÒvWBÔÆ–÷WE66÷T–æfòG7G'V7GW&RG6´–BGG'VP¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²6´–CÕ·7G&–æuÒG66÷Rç6´–C²6æ6†÷G3Ò„vWBÔÆ–÷WE6æ6†÷G2FÆæwVvR…·7G&–æuÒG66÷Rç6´–B’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"öÆ–÷WB÷&W7F÷&R÷&Wf–Wrr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Brrr¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&Wf–WsÒ„vWBÔÆ–÷WE&W7F÷&U&Wf–WrFÆæwVvRG6´–B…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6æ6†÷D–Brrr’’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"öÆ–÷WB÷&W7F÷&Rr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Brrr¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&W7VÇCÒ…&W7F÷&RÔÆ–÷WE6æ6†÷BFÆæwVvRG6´–B…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6æ6†÷D–Brrr’’“²7FFSÒ„vWBÕc%7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"ö÷WGWG2÷&VF–æW72r’°¢G6´–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w6´–BuÐ¢G7G'V7GW&RÒvWBÕ7G'V7GW&RFÆæwVvP¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG7G'V7GW&RG6´–BFfÇ6P¢GF&vWG2Ò¶÷&FW&VEÔ·Ð¢f÷&V6‚‚GF&vWD–B–â„vWBÕ6µF&vWD–G2FÆæwVvRG66÷Rç6²’’°¢GföÇVÖRÒvWBÔÆVv7•föÇVÖTg&öÕF&vWD–BFÆæwVvRGF&vWD–@¢GF&vWG5²GF&vWD–EÒÒvWBÔf–æÄ'V–ÆE&VF–æW72G7G'V7GW&RFÆæwVvRGföÇVÖR…·7G&–æuÒG66÷Rç6´–B¢Ð¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²6´–CÕ·7G&–æuÒG66÷Rç6´–C²F&vWG3ÒGF&vWG2Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"ö÷WGWG2ö'V–ÆBr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Brrr¢GF&vWD–G2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F&öG’wF&vWD–G2r‚’’¢G7G'V7GW&RÒvWBÕ7G'V7GW&RFÆæwVvP¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG7G'V7GW&RG6´–BFfÇ6P¢FÆÆ÷vVEF&vWD–G2Ò„vWBÕ6µF&vWD–G2FÆæwVvRG66÷Rç6²¢–b‚GF&vWD–G2ä6÷VçBÖW’²GF&vWD–G2Ò…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’wF&vWD–BrFÆÆ÷vVEF&vWD–G5³Ò’’Ð¢2X{®X©¾8þi[XØzy.8¾8¾8(¾8>8Ž8Î8.8(¾8.8+^8;Î898;Î8þ8:®8*þ8*Ž8+ž88Ž8).y»NX‰~8¾Xznyn8ž8(¾8þ8(8¢28>8>8~ZèÎ{Y8^8¾8(¾8Ž˜.hÙ~89Þ8;Î8:®8;>8+8(.KŠÞjÚ.8(.Xù~8K¹Ž88(ž8(Î8®8N8.8+Ž8:~89n8Ž8~8n‹ùN8ž8 ¢F¦ö"Ò7F'BÔf–æÄ'V–ÆD¦ö"FÆæwVvR…·7G&–æuÒG66÷Rç6´–B’‚GF&vWD–G2Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²¦ö#ÒF¦ö"Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"ö÷WGWG2ö'V–ÆB÷7FGW2r’°¢F¦ö$–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v¦ö$–BuÐ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²¦ö#Ò…&VBÔf–æÄ¦ö%7FGW2FÆæwVvRF¦ö$–B“²7FFSÒ„vWBÕc%7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"ö÷WGWG2ö'V–ÆBö6æ6VÂr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&W7VÇCÒ…&WVW7BÔf–æÄ¦ö$6æ6VÆÆF–öâFÆæwVvR…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v¦ö$–Brrr’’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"ö÷WGWG2öf–ÆRr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢6W'fRÔFö7VÖVçE6µFbD6öçFW‡BFÆæwVvR…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Brrr’’…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’wF&vWD–BrvÖ–âr’“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"÷6÷W&6W2÷&VÆ–æ²Ö6æF–FFW2r’°¢G&VÆ–æ´f÷"Ò·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w6÷W&6T–BuÐ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&W7VÇCÒ„vWBÕ&VÆ–æ´6æF–FFW2FÆæwVvRG&VÆ–æ´f÷"’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"÷6÷W&6W2÷&VÆ–æ²r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&VÆ–æµ&W7VÇBÒ&VÆ–æ²Õ6÷W&6RFÆæwVvR…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6÷W&6T–Brrr’’…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w&VÆF—fUF‚rrr’¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&W7VÇCÒG&VÆ–æµ&W7VÇBÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"ö÷WGWG2÷&WfVÂr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&WfVÅ6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Brrr¢G&WfVÅF&vWD–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’wF&vWD–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&WfVÅ6´–B’’²G&WfVÅ6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v6FVv÷'’rrr’Ð¢G&WfVÅF‚Ò&W6öÇfRÔ÷WGWEFdf÷%&WfVÂFÆæwVvRG&WfVÅ6´–BG&WfVÅF&vWD–B…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’wföÇVÖRrrr’¢÷VâÔ÷WGWEFdÆö6F–öâG&WfVÅF€¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²FƒÒG&WfVÅF‚Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷c"ö÷WGWG2÷V&Æ—6‚r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6´–Brrr¢GF&vWD–BÒ…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’wF&vWD–Brrr’’åG&–Ò‚’åFôÆ÷vW$–çf&–çB‚¢GV&Æ—6…7G'V7GW&RÒvWBÕ7G'V7GW&RFÆæwVvP¢GV&Æ—6…66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RGV&Æ—6…7G'V7GW&RG6´–BFfÇ6P¢GF&vWD–BÒ76W'BÕ6µF&vWD–BFÆæwVvRGV&Æ—6…66÷Rç6²GF&vWD–@¢GföÇVÖRÒvWBÔÆVv7•föÇVÖTg&öÕF&vWD–BFÆæwVvRGF&vWD–@¢G&W7VÇBÒV&Æ—6‚ÔFö7VÖVçE6µFeFõ6†&VBFÆæwVvRGföÇVÖRG6´–@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²&W7VÇCÒG&W7VÇBÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷c"ö÷WGWG2ö&6†—fW2r’°¢G6´–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w6´–BuÐ¢G7G'V7GW&RÒvWBÕ7G'V7GW&RFÆæwVvP¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG7G'V7GW&RG6´–BGG'VP¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö³ÒGG'VS²•fW'6–öãÓ#²6´–CÕ·7G&–æuÒG66÷Rç6´–C²&6†—fW3Ò„vWBÔf–æÄ&6†—fW2FÆæwVvR…·7G&–æuÒG66÷Rç6´–B’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷7FFRr’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#„vWBÕ7FFU–ÆöBFÆæwVvR“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’öf–æÂ÷&VF–æW72r’°¢F6CÕ&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v6FVv÷'’uÒ¢G7G'V7GW&SÔvWBÕ7G'V7GW&RFÆæwVvS²GföÇ3Õ¶÷&FW&VEÔ·Ð¢f÷&V6‚‚GföÇVÖR–â„vWBÕföÇVÖTÆ—7BFÆæwVvWÅv†W&RÔö&¦V7G²EòÖæRvæöæRwÒ’—²GföÇ5²GföÇVÖUÓÔvWBÔf–æÄ'V–ÆE&VF–æW72G7G'V7GW&RFÆæwVvRGföÇVÖRF6GÐ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ¶ö³ÒGG'VS¶6FVv÷'“ÒF6C·föÇVÖW3ÒGföÇ7Ò“·&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’÷7V&Ö—76–öâÖf–ÆW2r’°¢2z›®8î˜XÞX‰~8).[Èþ8î8î8îkŠ8ž8‚÷vW%6†VÆÂ8Î[^™h¾8~8bFçVÆÂ8¾8®8(®8¢26öçfW'EFòÔ§6öâ8Â&f–ÆW2#§·Ò8).‹ùN8ž8.Xù~8Xùn8>8þyK¾™Ú.8ò4'&’‚’8p¢2Šh{JK»n8Ži[8Ž8Xéþz‹þ8ÃK»n8î8Ž8Þ88X{®8(¾8þ8®8îjŽXhP¢2Ž8Î8+^89n89^8*ž8:¾888;Î8îKŠÞ8þhê.8~8î8¾8)>8Òž8Žk®8~8nX‹˜N8~8®8¾8>8þ8 ¢G6÷W&6Tf–ÆW2ÒvWBÕ6÷W&6T6æF–FFW2‚vW†6VÂrÂwv÷&BrÂwFbrÂw÷vW'ö–çBr¢–b‚FçVÆÂÖWG6÷W&6Tf–ÆW2’²G6÷W&6Tf–ÆW2Ò‚’Ð¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²66ææVDBÒ„æWrÔæ÷t—6ò“²f–ÆW2ÒG6÷W&6Tf–ÆW2Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷7V&Ö—76–öâ÷6VÆV7BÖæB×7F'Br’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F–æ—F–ÂÒ·7G&–æuÒF&öG’æ–æ—F–ÄF— ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F–æ—F–Â’’²F–æ—F–ÂÒ·7G&–æuÒ„vWBÕF‡2’ç7V&Ö—76–öäF—"Ð¢2yK¾™Ú.8þ8>8î89^8*ž8:¾888;Î8).Kˆ‹*¾8~8n8ÎXéþz‹þ89^8*ž8:¾888;Î8Þ8ŽYÎ8n8.h«Î8~8þ89Î8+þ8;>8€¢2™h¾8N8þ888*N8*.8:Þ8+8~YÎ8>YÞ8ÎZHž8(þ8(¾8Ž8˜Ž8nZûî‹8).Xùn8(®˜^8Ž8(¾8 ¢G6VÆV7FVBÒ6VÆV7BÔföÆFW$F–Æör~Xéþz‹þ89^8*ž8:¾888;Î8).˜Ž8brF–æ—F–À¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6VÆV7FVB’’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²6æ6VÆÆVBÒGG'VS²7FFRÒ„vWBÕ7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢FæWuF‡2ÒvWBÔFVfVÇD6†–ÆEF‡2G6VÆV7FV@¢FÖ–w&F–öâÒ–æ—F–Æ—¦RÔ6ÆVäÆö6Å&ö¦V7BFæWuF‡0¢Vç7W&RÕ6¶vRFæWuF‡2ÔÆæwVvW2‚FÆæwVvR¢F6öæf–rÒvWBÔ6öæf–p¢F6öæf–ræÆ7E7V&Ö—76–öäF—"Ò·7G&–æuÒFæWuF‡2ç7V&Ö—76–öäF— ¢F6öæf–ræÆ7DFFF—"Ò·7G&–æuÒFæWuF‡2æFFF— ¢F6öæf–ræÆ7D÷WGWDF—"Ò·7G&–æuÒFæWuF‡2æ÷WGWDF— ¢6fRÔ6öæf–rF6öæf–p¢2ö’÷7V&Ö—76–öâÖf–ÆW28ŽYÎ8Žz›®˜XÞX‰~8î[^™h¾8Î‹[~8Þ8(¾8.8>88(ž8Î8Î89^8*ž8:¾888;Î8) ¢2˜Ž8n8Þ8).h«Î8~8þK«®8Î˜	®8(¾K‹¾{XÎ‹zþ8~8Xéþz‹óK»n8î89^8*ž8:¾888;Î8).˜Ž8n8‚f–ÆW28À¢2·Ò8¾8®8(®8yK¾™Ú.XN8óK»n8Ži[8Ž8nXhÞXùn[ér†ÆöDf–ÆW56–ÆVçFÇ’ž8(.š9¾88~8n8N8þ8 ¢G6VÆV7FVDf–ÆW2ÒvWBÕ6÷W&6T6æF–FFW2‚vW†6VÂrÂwv÷&BrÂwFbrÂw÷vW'ö–çBr¢–b‚FçVÆÂÖWG6VÆV7FVDf–ÆW2’²G6VÆV7FVDf–ÆW2Ò‚’Ð¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²F‡2ÒFæWuF‡3²F‚Ò·7G&–æuÒFæWuF‡2ç7V&Ö—76–öäF—#²Ö–w&F–öâÒFÖ–w&F–öã²7FFRÒ„vWBÕ7FFU–ÆöBFÆæwVvR“²66ææVDBÒ„æWrÔæ÷t—6ò“²f–ÆW2ÒG6VÆV7FVDf–ÆW2Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öF–ÆöröföÆFW"r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G6VÆV7FVBÒ6VÆV7BÔföÆFW$F–Æör…·7G&–æuÒF&öG’çF—FÆR’…·7G&–æuÒF&öG’æ–æ—F–ÄF—"¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²F‚ÒG6VÆV7FVBÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷F‡2öFVfVÇG2r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²F‡2Ò„vWBÔFVfVÇD6†–ÆEF‡2…·7G&–æuÒF&öG’ç7V&Ö—76–öäF—"’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’öf–ÆRr’°¢6W'fRÔ6öçFVçEFbD6öçFW‡BFÆæwVvS²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öf–ÆRr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢6W'fRÔ6öçFVçEFdg&öÔ&öG’D6öçFW‡BFÆæwVvRF&öG“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’öf–æÂöf–ÆRr’°¢6W'fRÔf–æÅFbD6öçFW‡BFÆæwVvS²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öf–æÂöf–ÆRr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢6W'fRÔf–æÅFd'•föÇVÖRD6öçFW‡BFÆæwVvR…·7G&–æuÒF&öG’çföÇVÖR’…·7G&–æuÒF&öG’æ6FVv÷'’“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷F‡2r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢2zêyn88~8;Î8+þ8Ž˜	®[‹ŽX{®X©¾8þ[‹Ž8¾XŠžyJŽˆ^8:Þ8;Î8*¾8:¾8$ž8¾8(žX[iÈžXN8îK»¾hHþ898+ž8) ¢2hÈ~Zé®8~8n8ŠH~i[XŠžyJŽˆ^8îx«nhX¾8ÎXhÞ8>k{~8n8(¾{XÎ‹zþ8).jè¾8^8®8N8 ¢G&WVW7FVE7V&Ö—76–öäF—"Ò…·7G&–æuÒF&öG’ç7V&Ö—76–öäF—"’åG&–Ò‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&WVW7FVE7V&Ö—76–öäF—"’Ö÷"Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G&WVW7FVE7V&Ö—76–öäF—"ÕF…G—R6öçF–æW"’’°¢F‡&÷r~XZ^X©¾8~8þXéþz‹þ89^8*ž8:¾888ÎŠh¾8N8¾8(®8î8¾8)>8.ZNh˜8).z+®Š¨Þ8ž8(¾8¾88Î89^8*ž8:¾888).˜Ž8n8Þ8).KÛþyJŽ8~8n8þ88^8N8"p¢Ð¢FæWuF‡2ÒvWBÔFVfVÇD6†–ÆEF‡2G&WVW7FVE7V&Ö—76–öäF— ¢FÖ–w&F–öâÒ–æ—F–Æ—¦RÔ6ÆVäÆö6Å&ö¦V7BFæWuF‡0¢Vç7W&RÕ6¶vRFæWuF‡2ÔÆæwVvW2‚FÆæwVvR¢F6öæf–rÒvWBÔ6öæf–p¢F6öæf–ræÆ7E7V&Ö—76–öäF—"Ò·7G&–æuÒFæWuF‡2ç7V&Ö—76–öäF— ¢F6öæf–ræÆ7DFFF—"Ò·7G&–æuÒFæWuF‡2æFFF— ¢F6öæf–ræÆ7D÷WGWDF—"Ò·7G&–æuÒFæWuF‡2æ÷WGWDF— ¢6fRÔ6öæf–rF6öæf–p¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²F‡2ÒFæWuF‡2Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷v÷&¶&öö·2÷&Vv—7FW"Ö&F6‚r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&VÇ2Ò‚¢–b‚F&öG’ç&VÆF—fUF‡2’²G&VÇ2Ò„vWBÔ'&’F&öG’ç&VÆF—fUF‡2Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ’Ð¢VÇ6V–b‚F&öG’ç&VÆF—fUF‚’²G&VÇ2Ò…·7G&–æuÒF&öG’ç&VÆF—fUF‚’Ð¢–b‚G&VÇ2ä6÷VçBÖW’²F‡&÷r~y›¾˜Ë.8ž8(¾Xéþz‹þ8Î˜Žh©î8^8(Î8n8N8î8¾8)>8"rÐ¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒF&öG’æ6FVv÷'’¢G&W7VÇBÒ&Vv—7FW"Õ6÷W&6W4&F6‚FÆæwVvRG&VÇ2„vWBÔ'V–ÇF–å6´–BF6B’rp¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG&W7VÇC²7FFRÒ„vWBÕ7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷v÷&¶&öö·2÷&Vv—7FW"r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒF&öG’æ6FVv÷'’¢G&W7VÇBÒ&Vv—7FW"Õ6÷W&6RFÆæwVvR…·7G&–æuÒF&öG’ç&VÆF—fUF‚’„vWBÔ'V–ÇF–å6´–BF6B’…·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’w6÷W&6UG—Rrrr’¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG&W7VÇC²7FFRÒ„vWBÕ7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷v÷&¶&öö·2÷Vç&Vv—7FW"r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&W7VÇBÒVç&Vv—7FW"Õv÷&¶&öö²FÆæwVvR…·7G&–æuÒF&öG’çv÷&¶&öö´–B¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG&W7VÇBÒ“²&WGW&à¢Ð¢–b‚‚FÖWF†öBÖWttUBrÖ÷"FÖWF†öBÖWuõ5Br’ÖæBGF‚ÖWrö’ö¦ö'2÷7FGW2r’°¢F¦ö$–BÒrp¢–b‚FÖWF†öBÖWuõ5Br’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F¦ö$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v¦ö$–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–B’’²F¦ö$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’t¦ö$–Brrr’Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–B’’²F¦ö$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v–Brrr’Ð¢Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–B’’²F¦ö$–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v¦ö$–BuÒÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–B’’²F¦ö$–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²t¦ö$–BuÒÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–B’’²F¦ö$–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v–BuÒÐ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…&VBÕ&VæFW$¦ö%7FGW2FÆæwVvRF¦ö$–B“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’ö¦ö'2ö6æ6VÂr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F¦ö$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&öG’v¦ö$–Brrr¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…&WVW7BÕ&VæFW$¦ö$6æ6VÆÆF–öâFÆæwVvRF¦ö$–B“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷v÷&¶&öö·2÷&VæFW"÷7F'Br’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F–G2Ò‚¢–b‚F&öG’çv÷&¶&öö´–G2’²F–G2Ò„vWBÔ'&’F&öG’çv÷&¶&öö´–G2Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ’Ð¢G&W7VÇBÒ7F'BÕ&VæFW$¦ö"FÆæwVvRF–G2…¶&ööÅÒF&öG’æöæÇ•WFFVB’…·7G&–æuÒF&öG’æ6FVv÷'’¢F¦ö$–Df÷%&W7öç6RÒæ÷&ÖÆ—¦RÕ&VæFW$¦ö$–B…·7G&–æuÒ„vWBÔFF&÷W'G’G&W7VÇBv¦ö$–Brrr’¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–Df÷%&W7öç6R’’²G'’²F¦ö$–Df÷%&W7öç6RÒæ÷&ÖÆ—¦RÕ&VæFW$¦ö$–B…·7G&–æuÒG&W7VÇBæ¦ö$–B’Ò6F6‚²ÒÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–Df÷%&W7öç6R’’²F¦ö$–Df÷%&W7öç6RÒæ÷&ÖÆ—¦RÕ&VæFW$¦ö$–B…·7G&–æuÒG&W7VÇB’Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–Df÷%&W7öç6R’’²F‡&÷ruDnKÙÎh‰8+Ž8:~89n8îh8^Z8).KÙÎh‰8~8Þ8î8¾8)>8~8~8þ8.8(.8nKˆ[ªn8ÅDnKÙÎh‰8Þ8).h«Î8~8n8þ88^8N8"rÐ¢6WBÔæ÷FU&÷W'G’G&W7VÇBv¦ö$–BrF¦ö$–Df÷%&W7öç6P¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²¦ö$–BÒF¦ö$–Df÷%&W7öç6S²¦ö"ÒG&W7VÇC²7FFRÒ„vWBÕ7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷v÷&¶&öö·2÷&VæFW"r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G7G'V7GW&RÒvWBÕ7G'V7GW&RFÆæwVvP¢F–G2Ò‚¢–b‚F&öG’çv÷&¶&öö´–G2’²F–G2Ò„vWBÔ'&’F&öG’çv÷&¶&öö´–G2Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ’Ð¢VÇ6V–b‚F&öG’æöæÇ•WFFVBÖWGG'VR’²F–G2Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FGW2Ö–â‚væWrrÂvW†6VÂ×WFFVBrÂw6÷W&6R×WFFVBrÂw&VæFW"ÖW'&÷"r’Ö÷"·7G&–æuÒEòæ7W'&VçDW†6VÄ†6‚ÖæR·7G&–æuÒEòæÆ7E&VæFW&VDW†6VÄ†6‚ÒÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÒ’Ð¢VÇ6R²F–G2Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÒ’Ð¢G&W7VÇG2Ò‚¢f÷&V6‚‚F–B–âF–G2’°¢G'’°¢G"Ò&VæFW"Õ6÷W&6RFÆæwVvRF–@¢G%²vö²uÒÒGG'VP¢G&W7VÇG2³ÒG ¢Ò6F6‚°¢F×6rÒEòäW†6WF–öâäÖW76vP¢GW6W$×6rÒ6öçfW'EFòÕW6W%&VæFW$W'&÷"F×6p¢FGBÒvWBÔÆ7E&VæFW$GFV×Df÷"F–@¢FW'&÷$FWF–ÂÒvWBÔW'&÷$FWF–ÂEð¢6WBÕv÷&¶&ööµ&VæFW$W'&÷"FÆæwVvRF–BF×6rFW'&÷$FWF–Â…·7G&–æuÒFGBç6æ6†÷D–B’…·7G&–æuÒFGBæ†6‚¢G&W7VÇG2³Ò¶÷&FW&VEÔ²ö²ÒFfÇ6S²v÷&¶&öö´–BÒF–C²W'&÷"ÒF×6s²W6W$W'&÷"ÒGW6W$×6s²FWF–ÂÒFW'&÷$FWF–ÂÐ¢Ð¢Ð¢F†4W'&÷'2ÒFfÇ6P¢f÷&V6‚‚G'"–âG&W7VÇG2’²G'’²–b‚G'"ä6öçF–ç2‚vö²r’ÖæBG'%²vö²uÒÖWFfÇ6R’²F†4W'&÷'2ÒGG'VRÒÒ6F6‚²ÒÐ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇG2ÒG&W7VÇG3²†4W'&÷'2ÒF†4W'&÷'3²7FFRÒ„vWBÕ7FFU–ÆöBFÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷vW2÷&V÷&FW"r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒ…&V÷&FW"ÕvW2FÆæwVvRF&öG’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷vW2÷6÷'BÖ'’×6†VWBr’°¢F&öG“Õ&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ¶ö³ÒGG'VS·&W7VÇCÒ…6÷'BÕvW4'•6†VWBFÆæwVvRF&öG’—Ò“·&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷vW2÷WFFRr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢GvU&W7VÇBÒWFFRÕvRFÆæwVvRF&öG¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²vRÒGvU&W7VÇBçvS²&W7VÇBÒGvU&W7VÇBÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷vW2ö6öæf—&Òr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²vRÒ„6öæf—&ÒÕvRFÆæwVvRF&öG’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’÷66â×WFFW2r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢Ff÷&6T†6‚ÒFfÇ6P¢G'’²Ff÷&6T†6‚Ò¶&ööÅÒF&öG’æf÷&6T†6‚Ö÷"¶&ööÅÒF&öG’æf÷&6RÒ6F6‚²Ff÷&6T†6‚ÒFfÇ6RÐ¢G66å&W7VÇBÒ66âÕWFFW2FÆæwVvRFçVÆÂFf÷&6T†6€¢2cRÜ*s2ãü*s2ã3¢jIÎyú^8ŽYÎi˜.8¾KùÞZÙŽ8ž8(¾8.™ÙžjÚ.[è^88(Ž8(®X˜Þ8 ¢F6GW&VBÒWFFRÔ–çWD†—7F÷'”gFW%66âFÆæwVvP¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG66å&W7VÇC²6GW&VE6æ6†÷G2Ò‚F6GW&VB’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öf–æÂö'V–ÆBr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&W7VÇBÒ'V–ÆBÔf–æÅFbFÆæwVvR…·7G&–æuÒF&öG’çföÇVÖR’…·7G&–æuÒF&öG’æ6FVv÷'’¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG&W7VÇBÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öf–æÂö'V–ÆBÖÆÂr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒF&öG’æ6FVv÷'’¢GföÇ2Ò„vWBÕföÇVÖTÆ—7BFÆæwVvRÂv†W&RÔö&¦V7B²EòÖæRvæöæRrÒ¢G&W7VÇBÒ–çfö¶RÔf–æÄ'V–ÆEG&ç67F–öâFÆæwVvRF6BGföÇ0¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG&W7VÇBÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öf–æÂ÷V&Æ—6‚r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢G&W7VÇBÒV&Æ—6‚Ôf–æÅFeFõ6†&VBFÆæwVvR…·7G&–æuÒF&öG’çföÇVÖR’…·7G&–æuÒF&öG’æ6FVv÷'’¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG&W7VÇBÒ“²&WGW&à¢Ð ¢2ÒÒÒÒcS¢[^jÛN8;¾[zîXˆn8;¾ˆz®X¹^XznybÒÒÒÐ¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’ö†—7F÷'’÷F–ÖVÆ–æRr’°¢FÆ–Ö—BÒ ¢G'’²FÆ–Ö—BÒ¶–çEÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²vÆ–Ö—BuÒÒ6F6‚²Ð¢–b‚FÆ–Ö—BÖÆRÖ÷"FÆ–Ö—BÖwBS’²FÆ–Ö—BÒÐ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²WfVçG2Ò„vWBÔ†—7F÷'•F–ÖVÆ–æRFÆæwVvRFÆ–Ö—B’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’ö†—7F÷'’÷6æ6†÷G2r’°¢Gv$–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wv÷&¶&öö´–BuÐ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²6æ6†÷G2Ò„vWBÕ6æ6†÷E7VÖÖ&–W2FÆæwVvRGv$–B’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’ö†—7F÷'’öF–fbr’°¢Gv$–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wv÷&¶&öö´–BuÐ¢Fg&öÔ–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²vg&öÕ6æ6†÷D–BuÐ¢GFô–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wFõ6æ6†÷D–BuÐ¢FF–fbÒFçVÆÀ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Fg&öÔ–B’ÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GFô–B’’°¢FF–fbÒvWBÔÆFW7D6ö×&—6öâFÆæwVvRGv$–@¢ÒVÇ6R°¢Fg&öÕfW"ÒvWBÕ&VfW'&VD†—7F÷'•&VæFW%fW'6–öâFÆæwVvRGv$–BFg&öÔ–@¢GFõfW"ÒvWBÕ&VfW'&VD†—7F÷'•&VæFW%fW'6–öâFÆæwVvRGv$–BGFô–@¢FF–fbÒ–b‚Fg&öÕfW"ÖæBGFõfW"’²vWBÕ7F÷&VD6ö×&—6öâFÆæwVvRGv$–BFg&öÔ–BGFô–BFg&öÕfW"GFõfW"v†—7F÷'’rÒVÇ6R²FçVÆÂÐ¢Ð¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²F–fbÒFF–fbÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’ö†—7F÷'’öF–fbÖFWF–Âr’°¢Gv$–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wv÷&¶&öö´–BuÐ¢Fg&öÔ–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²vg&öÕ6æ6†÷D–BuÐ¢GFô–BÒ·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wFõ6æ6†÷D–BuÐ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²FWF–ÂÒ„vWBÔF–fdFWF–ÂFÆæwVvRGv$–BFg&öÔ–BGFô–B’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’ö†—7F÷'’öF–fb×&Wf–Wrr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&Wf–WrÒ…6WBÔF–fe&Wf–Wu7FFRFÆæwVvRF&öG’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’ö†—7F÷'’öF–fb÷&W&Rr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²¦ö"Ò…7F'BÔF–fdFWF–Ä¦ö"FÆæwVvR…·7G&–æuÒF&öG’çv÷&¶&öö´–B’…·7G&–æuÒF&öG’æg&öÕ6æ6†÷D–B’…·7G&–æuÒF&öG’çFõ6æ6†÷D–B’…·7G&–æuÒF&öG’ç6†VWD¶W’’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’ö†—7F÷'’öF–fb×vRr’°¢6W'fRÔF–fevRD6öçFW‡BFÆæwVvR ¢…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wv÷&¶&öö´–BuÒ’ ¢…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v7W'&VçE6æ6†÷D–BuÒ’ ¢…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v&6VÆ–æU6æ6†÷D–BuÒ’ ¢…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w6†VWD¶W’uÒ’ ¢…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wvTçVÖ&W"uÒ’ ¢…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v76WBuÒ’ ¢…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w66÷RuÒ¢&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’ö†—7F÷'’ö6öçFVçB×Fbr’°¢6W'fRÔ†—7F÷'”6öçFVçEFbD6öçFW‡BFÆæwVvR…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wv÷&¶&öö´–BuÒ’…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²wfW'6–öä–BuÒ’…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w6†VWDæÖRuÒ’…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²w6æ6†÷D–BuÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’ö†—7F÷'’÷–âr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢Gv$–BÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒF&öG’çv÷&¶&öö´–B’wv÷&¶&öö´–Bp¢G6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒF&öG’ç6æ6†÷D–B’w6æ6†÷D–Bp¢·fö–EÒ…6WBÔÖçVÅ6æ6†÷E–âFÆæwVvRGv$–BG6æ6†÷D–BGG'VR¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VRÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’ö†—7F÷'’÷Vç–âr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢Gv$–BÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒF&öG’çv÷&¶&öö´–B’wv÷&¶&öö´–Bp¢G6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒF&öG’ç6æ6†÷D–B’w6æ6†÷D–Bp¢·fö–EÒ…6WBÔÖçVÅ6æ6†÷E–âFÆæwVvRGv$–BG6æ6†÷D–BFfÇ6R¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VRÒ“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’öÆ–÷WB÷6æ6†÷G2r’°¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v6FVv÷'’uÒ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²6æ6†÷G2Ò„vWBÔÆ–÷WE6æ6†÷G2FÆæwVvRF6B’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öÆ–÷WB÷&W7F÷&R÷&Wf–Wrr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒF&öG’æ6FVv÷'’¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&Wf–WrÒ„vWBÔÆ–÷WE&W7F÷&U&Wf–WrFÆæwVvRF6B…·7G&–æuÒF&öG’ç6æ6†÷D–B’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öÆ–÷WB÷&W7F÷&Rr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒF&öG’æ6FVv÷'’¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒ…&W7F÷&RÔÆ–÷WE6æ6†÷BFÆæwVvRF6B…·7G&–æuÒF&öG’ç6æ6†÷D–B’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’öf–æÂö&6†—fW2r’°¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’…·7G&–æuÒD6öçFW‡Bå&WVW7BåVW'•7G&–æu²v6FVv÷'’uÒ¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&6†—fW2Ò„vWBÔf–æÄ&6†—fW2FÆæwVvRF6B’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWttUBrÖæBGF‚ÖWrö’öWFò÷7FFRr’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²WFòÒ„vWBÔWFõ7FFU7VÖÖ'’FÆæwVvR’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuD4‚rÖæBGF‚ÖWrö’öWFò÷6WGF–æw2r’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²WFòÒ…WFFRÔWFõ&VæFW%6WGF–æw2F&öG’’Ò“²&WGW&à¢Ð¢–b‚FÖWF†öBÖWuõ5BrÖæBGF‚ÖWrö’öWFò÷'VâÖæ÷rr’°¢F&öG’Ò&VBÔ&öG”§6öâD6öçFW‡Bå&WVW7@¢Gv$–BÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒF&öG’çv÷&¶&öö´–B’wv÷&¶&öö´–Bp¢w&—FRÔ§6öå&W7öç6RD6öçFW‡B#…¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒ…&WVW7BÔWFõ'Väæ÷rFÆæwVvRGv$–B’Ò“²&WGW&à¢Ð¢w&—FRÔ§6öå&W7öç6RD6öçFW‡BCB…¶÷&FW&VEÔ²ö²ÒFfÇ6S²W'&÷"ÒwVæ¶æ÷vâ’&÷WFRrÒ¢Ò6F6‚µ7—7FVÒä&wVÖVçDW†6WF–öåÒ°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡BC…¶÷&FW&VEÔ²ö²ÒFfÇ6S²W'&÷"ÒEòäW†6WF–öâäÖW76vRÒ¢Ò6F6‚°¢2j[ÞŠk>8:Þ88>8*þ8îŠÞz¨8þ8XŠžyJŽˆ^XN8~Šz>k®8~8Þ8(¾x«nhX¾8®8(Î8.8+^8;Î898;Î™©ÎZë>8ŽXË®XŠ^8ž8(¾8 ¢–b…FW7BÕ7G'V7GW&T6öæfÆ–7DW'&÷"EòäW†6WF–öâ’°¢w&—FRÔ§6öå&W7öç6RD6öçFW‡BC’…¶÷&FW&VEÔ°¢ö²ÒFfÇ6S²6öFRÒw7G'V7GW&RÖ6öæfÆ–7Bs²W'&÷"ÒEòäW†6WF–öâäÖW76vP¢7W'&VçDÆ–÷WBÒ·7G&–æuÒEòäW†6WF–öâäFF²v7GVÄÆ–÷WBuÐ¢Ò¢&WGW&à¢Ð¢w&—FRÔ§6öå&W7öç6RD6öçFW‡BS…¶÷&FW&VEÔ²ö²ÒFfÇ6S²W'&÷"ÒEòäW†6WF–öâäÖW76vS²FWF–ÂÒ„vWBÔW'&÷$FWF–ÂEò’Ò¢Ð§Ð  ¦gVæ7F–öâFW7BÕF76öçFW‡B‚D6öçFW‡B’°¢&WGW&â‚FçVÆÂÖæRD6öçFW‡BÖæBFçVÆÂÖæRD6öçFW‡Bå4ö&¦V7Bå&÷W'F–W5²t—5F7uÒÖæBD6öçFW‡Bä—5F7ÖWGG'VR§Ð ¦gVæ7F–öâvWBÔ‡GG7FGW5FW‡B…¶–çEÒE7FGW2’°¢7v—F6‚‚E7FGW2’°¢#²&WGW&âtô²rÐ¢#B²&WGW&âtæò6öçFVçBrÐ¢C²&WGW&ât&B&WVW7BrÐ¢C2²&WGW&âtf÷&&–FFVârÐ¢CB²&WGW&âtæ÷Bf÷VæBrÐ¢S²&WGW&ât–çFW&æÂ6W'fW"W'&÷"rÐ¢FVfVÇB²&WGW&âtô²rÐ¢Ð§Ð ¦gVæ7F–öâw&—FRÕF7&W7öç6R‚D6öçFW‡BÂ¶–çEÒE7FGW2Â¶'—FUµÕÒD'—FW2Â·7G&–æuÒD6öçFVçEG—RÂ¶&ööÅÒDÆÆ÷t6÷'2ÒFfÇ6RÂ·7G&–æuÒD66†T6öçG&öÂÒvæò×7F÷&Rr’°¢G'’°¢G7FGW5FW‡BÒvWBÔ‡GG7FGW5FW‡BE7FGW0¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D66†T6öçG&öÂ’’²D66†T6öçG&öÂÒvæò×7F÷&RrÐ¢D66†T6öçG&öÂÒD66†T6öçG&öÂ×&WÆ6R%µÇ%ÆåÒ"Ârp¢24õ%>89Ž88>888;Î8þ‹[~X¹^[è^889®8;Î8+‚†f–ÆS¢òòž8ÎŠªÞ8([ø^Šh8î8.8(¾‹»Þ˜xþ8*Ž8;>88ž89Þ8*N8;>88Ž888¾K¹Ž88(¾8 ¢288Ž8;Î8*þ8;>KùÞŠÛtž8²66W72Ô6öçG&öÂÔÆÆ÷rÔ÷&–v–ã¢¢8).K¹Ž88(¾8Ž888Ž8;Î8*þ8;>kÈþ8Ž8Ni˜.8°¢2K»¾hHþ8åvV.89®8;Î8+Ž8¾8(ž[ùÎzÙN8).ŠªÞ8(8n8~8î8n8þ8(8iz.Zé®8~8þK¹Ž88®8NûÈŽYÎKˆ8*®8:®8+Ž8;>8¾8þKˆÞŠhûÈž8 ¢F6÷'4†VFW"Òrp¢–b‚DÆÆ÷t6÷'2’²F6÷'4†VFW"Ò$66W72Ô6öçG&öÂÔÆÆ÷rÔ÷&–v–ã¢¦&â"Ð¢F†VFW"Ò$…EEóãE7FGW2G7FGW5FW‡F&ä6öçFVçBÕG—S¢D6öçFVçEG—V&ä6öçFVçBÔÆVæwFƒ¢B‚D'—FW2äÆVæwF‚–&ä66†RÔ6öçG&öÃ¢D66†T6öçG&öÆ&âG¶6÷'4†VFW'Ô6öææV7F–öã¢6Æ÷6V&æ&â ¢F†VFW$'—FW2ÒµFW‡BäVæ6öF–æuÓ£¤44”’ävWD'—FW2‚F†VFW"¢G7G&VÒÒD6öçFW‡BåF77G&VÐ¢G7G&VÒåw&—FR‚F†VFW$'—FW2ÂÂF†VFW$'—FW2äÆVæwF‚¢–b‚D'—FW2äÆVæwF‚ÖwB’²G7G&VÒåw&—FR‚D'—FW2ÂÂD'—FW2äÆVæwF‚’Ð¢G7G&VÒäfÇW6‚‚¢Òf–æÆÇ’°¢G'’²–b‚D6öçFW‡BåF77G&VÒ’²D6öçFW‡BåF77G&VÒä6Æ÷6R‚’ÒÒ6F6‚·Ð¢G'’²–b‚D6öçFW‡BåF76Æ–VçB’²D6öçFW‡BåF76Æ–VçBä6Æ÷6R‚’ÒÒ6F6‚·Ð¢Ð§Ð ¦gVæ7F–öâf–æBÔ‡GG†VFW$VæB…¶'—FUµÕÒD'—FW2’°¢–b‚D'—FW2äÆVæwF‚ÖÇBB’²&WGW&âÓÐ¢f÷"‚F’Ò3²F’ÖÇBD'—FW2äÆVæwFƒ²F’²²’°¢–b‚D'—FW5²F’Ò5ÒÖW2ÖæBD'—FW5²F’Ò%ÒÖWÖæBD'—FW5²F’ÒÒÖW2ÖæBD'—FW5²F•ÒÖW’°¢&WGW&â‚F’Ò2¢Ð¢Ð¢&WGW&âÓ§Ð ¦gVæ7F–öâæWrÔæÖUfÇVT6öÆÆV7F–öä6ö×B°¢G'’²&WGW&âæWrÔö&¦V7B7—7FVÒä6öÆÆV7F–öç2å7V6–Æ—¦VBäæÖUfÇVT6öÆÆV7F–öâ…µ7G&–æt6ö×&W%Ó£¤÷&F–æÄ–væ÷&T66R’Ð¢6F6‚²&WGW&âæWrÔö&¦V7B7—7FVÒä6öÆÆV7F–öç2å7V6–Æ—¦VBäæÖUfÇVT6öÆÆV7F–öâÐ§Ð ¦gVæ7F–öâFV6öFRÕW&Å'B…·7G&–æuÒEfÇVR’°¢–b‚FçVÆÂÖWEfÇVR’²&WGW&ârrÐ¢&WGW&âµW&•Ó£¥VæW66TFF7G&–ær‚‚EfÇVR×&WÆ6RuÂ²rÂrr’§Ð ¦gVæ7F–öâæWrÕVW'•7G&–æt6öÆÆV7F–öâ…·7G&–æuÒEVW'’’°¢Fçf2ÒæWrÔæÖUfÇVT6öÆÆV7F–öä6ö×@¢2æÖUfÇVT6öÆÆV7F–öâ—2VçVÖW&&ÆRâ&WGW&æ–ær—Bæ÷&ÖÆÇ’Ö¶W2÷vW%6†VÆÀ¢2Vçw&—G2fÇVW2–çFòö&¦V7EµÒ†÷"66Æ"f÷"6–ævÆRVW'’—FVÒ’Â6ð¢2&WVW7BåVW'•7G&–æu²wFö¶VâuÒ6âæWfW"&WG&–WfRF†RFö¶Vââ¶VWF†P¢26öÆÆV7F–öâ2öæR—VÆ–æRö&¦V7BöâWfW'’&WGW&âF‚à¢–b…·7G&–æuÓ£¤—4çVÆÄ÷$V×G’‚EVW'’’’²&WGW&âÂFçf2Ð¢GÒEVW'¢–b‚Gå7F'G5v—F‚‚sòr’’²GÒGå7V'7G&–ærƒ’Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷$V×G’‚G’’²&WGW&âÂFçf2Ð¢f÷&V6‚‚G—"–â‚G×7Æ—Brbr’’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷$V×G’‚G—"’’²6öçF–çVRÐ¢FWÒG—"ä–æFW„öb‚sÒr¢–b‚FWÖvR’°¢FæÖRÒFV6öFRÕW&Å'BG—"å7V'7G&–ærƒÂFW¢GfÇVRÒFV6öFRÕW&Å'BG—"å7V'7G&–ær‚FW²¢ÒVÇ6R°¢FæÖRÒFV6öFRÕW&Å'BG— ¢GfÇVRÒrp¢Ð¢Fçf2äFB‚FæÖRÂGfÇVR¢Ð¢&WGW&âÂFçf0§Ð ¦gVæ7F–öâ&VBÕF7‡GG6öçFW‡B‚EF76Æ–VçBÂ¶–çEÒE÷'B’°¢G7G&VÒÒEF76Æ–VçBävWE7G&VÒ‚¢2'&÷w6W'26â÷Vâ7V7VÆF—fRö–FÆRÆö6Â6öææV7F–öç2&Vf÷&R6VæF–ær&WVW7Bà¢2F†R6W'fW"†æFÆW2&WVW7G26WVVçF–ÆÇ’Â6òÆöær†VFW"F–ÖV÷WB6âg&VW¦R7F'GWà¢G7G&VÒå&VEF–ÖV÷WBÒ# ¢F'VffW"ÒæWrÔö&¦V7B'—FUµÒƒ“ ¢F×2ÒæWrÔö&¦V7B”òäÖVÖ÷'•7G&VÐ¢F†VFW$VæBÒÓ¢v†–ÆR‚F†VFW$VæBÖÇB’°¢G&VBÒG7G&VÒå&VB‚F'VffW"ÂÂF'VffW"äÆVæwF‚¢–b‚G&VBÖÆR’²F‡&÷rtV×G’…EE&WVW7BârÐ¢F×2åw&—FR‚F'VffW"ÂÂG&VB¢FFFÒF×2åFô'&’‚¢F†VFW$VæBÒf–æBÔ‡GG†VFW$VæBFFF¢–b‚F×2äÆVæwF‚ÖwBcSS3bÖæBF†VFW$VæBÖÇB’²F‡&÷rt…EE†VFW"—2FöòÆ&vRârÐ¢Ð ¢FFFÒF×2åFô'&’‚¢F†VFW%FW‡BÒµFW‡BäVæ6öF–æuÓ£¤44”’ävWE7G&–ær‚FFFÂÂF†VFW$VæB¢FÆ–æW2ÒF†VFW%FW‡B×7Æ—B&#öâ ¢–b‚FÆ–æW2ä6÷VçBÖÇB’²F‡&÷rt–çfÆ–B…EE&WVW7BârÐ¢G&WVW7DÆ–æRÒFÆ–æW5³Ò×7Æ—Brp¢–b‚G&WVW7DÆ–æRä6÷VçBÖÇB"’²F‡&÷rt–çfÆ–B…EE&WVW7BÆ–æRârÐ¢FÖWF†öBÒG&WVW7DÆ–æU³Ð¢GF&vWBÒG&WVW7DÆ–æU³Ð ¢F†VFW'2ÒæWrÔæÖUfÇVT6öÆÆV7F–öä6ö×@¢f÷"‚F’Ò²F’ÖÇBFÆ–æW2ä6÷VçC²F’²²’°¢FÆ–æRÒFÆ–æW5²F•Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FÆ–æR’’²6öçF–çVRÐ¢F6öÆöâÒFÆ–æRä–æFW„öb‚s¢r¢–b‚F6öÆöâÖwB’°¢F†VFW'2äFB‚FÆ–æRå7V'7G&–ærƒÂF6öÆöâ’åG&–Ò‚’ÂFÆ–æRå7V'7G&–ær‚F6öÆöâ²’åG&–Ò‚’¢Ð¢Ð ¢F6öçFVçDÆVæwF‚Ò ¢F6öçFVçDÆVæwF…FW‡BÒ·7G&–æuÒF†VFW'5²t6öçFVçBÔÆVæwF‚uÐ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öçFVçDÆVæwF…FW‡B’’°¢–b‚Öæ÷B¶–çEÓ£¥G'•'6R‚F6öçFVçDÆVæwF…FW‡BÂ·&VeÒF6öçFVçDÆVæwF‚’Ö÷"F6öçFVçDÆVæwF‚ÖÇB’°¢F‡&÷rµ7—7FVÒä&wVÖVçDW†6WF–öåÓ£¦æWr‚t–çfÆ–B6öçFVçBÔÆVæwF‚†VFW"âr¢Ð¢2&W÷'D&–æFW"öæÇ’66WG26ÖÆÂ¥4ôâ6öÖÖæG2âÆ–Ö—B¶VW2ÖÆf÷&ÖV@¢2÷"†÷7F–ÆRÆö6Â&WVW7Bg&öÒ&Æö6¶–ærF†R6–ævÆR×F‡&VFVBÆ—7FVæW"à¢–b‚F6öçFVçDÆVæwF‚ÖwBCƒSsb’°¢F‡&÷rµ7—7FVÒä&wVÖVçDW†6WF–öåÓ£¦æWr‚u&WVW7B&öG’—2FöòÆ&vRâr¢Ð¢Ð¢–b‚F6öçFVçDÆVæwF‚ÖwB’²G7G&VÒå&VEF–ÖV÷WBÒÐ¢F&öG”×2ÒæWrÔö&¦V7B”òäÖVÖ÷'•7G&VÐ¢F&öG•7F'BÒF†VFW$VæB²@¢–b‚FFFäÆVæwF‚ÖwBF&öG•7F'BÖæBF6öçFVçDÆVæwF‚ÖwB’°¢Ff–Æ&ÆRÒ´ÖF…Ó£¤Ö–â‚FFFäÆVæwF‚ÒF&öG•7F'BÂF6öçFVçDÆVæwF‚¢–b‚Ff–Æ&ÆRÖwB’²F&öG”×2åw&—FR‚FFFÂF&öG•7F'BÂFf–Æ&ÆR’Ð¢Ð¢v†–ÆR‚F&öG”×2äÆVæwF‚ÖÇBF6öçFVçDÆVæwF‚’°¢FæVVFVBÒ´ÖF…Ó£¤Ö–â‚F'VffW"äÆVæwF‚ÂF6öçFVçDÆVæwF‚Ò¶–çEÒF&öG”×2äÆVæwF‚¢G&VBÒG7G&VÒå&VB‚F'VffW"ÂÂFæVVFVB¢–b‚G&VBÖÆR’²'&V²Ð¢F&öG”×2åw&—FR‚F'VffW"ÂÂG&VB¢Ð¢–b‚F&öG”×2äÆVæwF‚ÖæRF6öçFVçDÆVæwF‚’°¢F‡&÷rµ7—7FVÒä&wVÖVçDW†6WF–öåÓ£¦æWr‚t–æ6ö×ÆWFR…EE&WVW7B&öG’âr¢Ð¢F&öG”×2å÷6—F–öâÒ  ¢GF„öæÇ’ÒGF&vW@¢GVW'’Òrp¢GÖ&²ÒGF&vWBä–æFW„öb‚sòr¢–b‚GÖ&²ÖvR’°¢GF„öæÇ’ÒGF&vWBå7V'7G&–ærƒÂGÖ&²¢GVW'’ÒGF&vWBå7V'7G&–ær‚GÖ&²²¢Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GF„öæÇ’’’²GF„öæÇ’ÒròrÐ¢GW&’ÒæWrÔö&¦V7B7—7FVÒåW&’‚&‡GG¢òó#rããã¢E÷'BGF&vWB" ¢G&WVW7BÒ·67W7FöÖö&¦V7EÔ°¢W&ÂÒGW&¢‡GGÖWF†öBÒFÖWF†ö@¢†VFW'2ÒF†VFW'0¢VW'•7G&–ærÒ„æWrÕVW'•7G&–æt6öÆÆV7F–öâGVW'’¢–çWE7G&VÒÒF&öG”×0¢6öçFVçDVæ6öF–ærÒµFW‡BäVæ6öF–æuÓ£¥UDc€¢Ð¢&WGW&â·67W7FöÖö&¦V7EÔ°¢—5F7ÒGG'VP¢F76Æ–VçBÒEF76Æ–Vç@¢F77G&VÒÒG7G&VÐ¢&WVW7BÒG&WVW7@¢&W7öç6RÒ·67W7FöÖö&¦V7EÔ·Ð¢Ð§Ð   ¦gVæ7F–öâ&W6öÇfRÔVFvTW†V7WF&ÆTg&öÔ6öÖÖæEFW‡B…·7G&–æuÒD6öÖÖæEFW‡B’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D6öÖÖæEFW‡B’’²&WGW&ârrÐ¢FW‡æFVBÒ´Vçf—&öæÖVçEÓ£¤W‡æDVçf—&öæÖVçEf&–&ÆW2‚D6öÖÖæEFW‡BåG&–Ò‚’¢F6æF–FFRÒrp¢–b‚FW‡æFVBÖÖF6‚uåÇ2¢"…µâ%Ò¦×6VFvUÂæW†R’"r’²F6æF–FFRÒFÖF6†W5³ÒÐ¢VÇ6V–b‚FW‡æFVBÖÖF6‚uåÇ2¢…µåÇ2%Ò¦×6VFvUÂæW†R’r’²F6æF–FFRÒFÖF6†W5³ÒÐ¢VÇ6V–b‚FW‡æFVBÖÖF6‚r"…µâ%Ò¦×6VFvUÂæW†R’"r’²F6æF–FFRÒFÖF6†W5³ÒÐ¢VÇ6V–b‚FW‡æFVBÖÖF6‚r…µåÇ2%Ò¦×6VFvUÂæW†R’r’²F6æF–FFRÒFÖF6†W5³ÒÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6æF–FFR’’²&WGW&ârrÐ¢F6æF–FFRÒF6æF–FFRåG&–Ò‚r"r¢G'’°¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚F6æF–FFR’²&WGW&â…´”òåF…Ó£¤vWDgVÆÅF‚‚F6æF–FFR’’Ð¢Ò6F6‚·Ð¢&WGW&ârp§Ð ¦gVæ7F–öâFBÔVFvTW†V7WF&ÆT6æF–FFR‚D6æF–FFW2Â·7G&–æuÒEfÇVR’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EfÇVR’’²&WGW&âÐ¢FW‡æFVBÒ´Vçf—&öæÖVçEÓ£¤W‡æDVçf—&öæÖVçEf&–&ÆW2‚EfÇVRåG&–Ò‚’¢G'6VBÒ&W6öÇfRÔVFvTW†V7WF&ÆTg&öÔ6öÖÖæEFW‡BFW‡æFV@¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G'6VB’’°¢·fö–EÒD6æF–FFW2äFB‚G'6VB¢&WGW&à¢Ð¢G&rÒFW‡æFVBåG&–Ò‚r"r¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&r’’²·fö–EÒD6æF–FFW2äFB‚G&r’Ð§Ð ¦gVæ7F–öâvWBÕ&Vv—7G'”FVfVÇEfÇVUFW‡B…·7G&–æuÒEF‚’°¢G'’°¢F¶W’ÒvWBÔ—FVÒÔÆ—FW&ÅF‚EF‚ÔW'&÷$7F–öâ7F÷ ¢&WGW&â·7G&–æuÒF¶W’ävWEfÇVR‚rr¢Ò6F6‚°¢&WGW&ârp¢Ð§Ð ¦gVæ7F–öâvWBÔVFvTW†V7WF&ÆUF‚°¢F6æF–FFW2ÒæWrÔö&¦V7B7—7FVÒä6öÆÆV7F–öç2ävVæW&–2äÆ—7E·7G&–æuÐ¢G'’°¢F6ÖBÒvWBÔ6öÖÖæB×6VFvRæW†RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢–b‚F6ÖBÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒF6ÖBå6÷W&6R’’°¢FBÔVFvTW†V7WF&ÆT6æF–FFRF6æF–FFW2…·7G&–æuÒF6ÖBå6÷W&6R¢Ð¢Ò6F6‚·Ð ¢f÷&V6‚‚G&VuF‚–â€¢t„´5S¥Å4ôeEt$UÄÖ–7&÷6ögEÅv–æF÷w5Ä7W'&VçEfW'6–öåÄF‡5Æ×6VFvRæW†RrÀ¢t„´ÄÓ¥Å4ôeEt$UÄÖ–7&÷6ögEÅv–æF÷w5Ä7W'&VçEfW'6–öåÄF‡5Æ×6VFvRæW†RrÀ¢t„´ÄÓ¥Å4ôeEt$UÅtõscC3$æöFUÄÖ–7&÷6ögEÅv–æF÷w5Ä7W'&VçEfW'6–öåÄF‡5Æ×6VFvRæW†RrÀ¢t„´5S¥Å4ôeEt$UÄ6Æ76W5ÄÕ4VFvT…DÕÇ6†VÆÅÆ÷VåÆ6öÖÖæBrÀ¢t„´ÄÓ¥Å4ôeEt$UÄ6Æ76W5ÄÕ4VFvT…DÕÇ6†VÆÅÆ÷VåÆ6öÖÖæBrÀ¢t„´5S¥Å4ôeEt$UÄ6Æ76W5ÆÖ–7&÷6ögBÖVFvUÇ6†VÆÅÆ÷VåÆ6öÖÖæBrÀ¢t„´ÄÓ¥Å4ôeEt$UÄ6Æ76W5ÆÖ–7&÷6ögBÖVFvUÇ6†VÆÅÆ÷VåÆ6öÖÖæBp¢’’°¢FBÔVFvTW†V7WF&ÆT6æF–FFRF6æF–FFW2„vWBÕ&Vv—7G'”FVfVÇEfÇVUFW‡BG&VuF‚¢Ð ¢f÷&V6‚‚F&6R–â€¢´Vçf—&öæÖVçEÓ£¤vWDVçf—&öæÖVçEf&–&ÆR‚u&öw&Ôf–ÆW2‡ƒƒb’r’À¢´Vçf—&öæÖVçEÓ£¤vWDVçf—&öæÖVçEf&–&ÆR‚u&öw&Ôf–ÆW2r’À¢´Vçf—&öæÖVçEÓ£¤vWDVçf—&öæÖVçEf&–&ÆR‚tÆö6ÄFFr¢’’°¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6R’’°¢FBÔVFvTW†V7WF&ÆT6æF–FFRF6æF–FFW2„¦ö–âÕF‚F&6RtÖ–7&÷6ögEÄVFvUÄÆ–6F–öåÆ×6VFvRæW†Rr¢Ð¢Ð ¢f÷&V6‚‚F6æF–FFR–â‚F6æF–FFW2Â6VÆV7BÔö&¦V7BÕVæ—VR’’°¢G'’°¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6æF–FFR’ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚F6æF–FFR’’°¢&WGW&â…´”òåF…Ó£¤vWDgVÆÅF‚‚F6æF–FFR’¢Ð¢Ò6F6‚·Ð¢Ð¢&WGW&ârp§Ð ¦gVæ7F–öâ6öçfW'EFòÔVFvT÷VåF&vWB…·7G&–æuÒEF&vWB’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EF&vWB’’²&WGW&ârrÐ¢GG&–ÖÖVBÒEF&vWBåG&–Ò‚¢G'’°¢2v–æF÷w2G&—fRF‚7V6‚23¥Ââââ×W7Bæ÷B&RÖ—7F¶Vâf÷"U$’66†VÖRà¢–b‚GG&–ÖÖVBÖÖF6‚uå´Õ¦×¥Ó¥µÅÂõÒrÖ÷"GG&–ÖÖVBÖÖF6‚uåÅÅÅÂr’°¢&WGW&â…µW&•Ò…´”òåF…Ó£¤vWDgVÆÅF‚‚GG&–ÖÖVB’’’ä'6öÇWFUW&¢Ð¢–b‚GG&–ÖÖVBÖÖF6‚uæf–ÆS¢r’²&WGW&âGG&–ÖÖVBÐ¢–b‚GG&–ÖÖVBÖÖF6‚uå´Õ¦×¥Õ´Õ¦×£Ó’²âÕÒ£¢r’²&WGW&âGG&–ÖÖVBÐ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GG&–ÖÖVB’°¢&WGW&â…µW&•Ò…´”òåF…Ó£¤vWDgVÆÅF‚‚GG&–ÖÖVB’’’ä'6öÇWFUW&¢Ð¢&WGW&âGG&–ÖÖV@¢Ò6F6‚°¢&WGW&âGG&–ÖÖV@¢Ð§Ð  ¦gVæ7F–öâ÷VâÕ&W÷'D&–æFW$'&÷w6W"…·7G&–æuÒEW&Â’°¢w&—FRÔ†÷7B%&W÷'D&–æFW"U$Ã¢EW&Â ¢GF&vWBÒ6öçfW'EFòÔVFvT÷VåF&vWBEW&À¢FVFvRÒvWBÔVFvTW†V7WF&ÆUF€¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FVFvR’’°¢G'’²7F'BÕ&ö6W72Ôf–ÆUF‚FVFvRÔ&wVÖVçDÆ—7B‚GF&vWB’ÔW'&÷$7F–öâ7F÷Â÷WBÔçVÆÃ²&WGW&âGG'VRÐ¢6F6‚²w&—FRÕv&æ–ær‚$6÷VÆBæ÷B÷VâÖ–7&÷6ögBVFvS¢"²EòäW†6WF–öâäÖW76vR’Ð¢ÒVÇ6R°¢w&—FRÕv&æ–ær$Ö–7&÷6ögBVFvRW†V7WF&ÆRF‚v2æ÷B&W6öÇfVBâG'––ær&÷Fö6öÂö6öÖÖæBfÆÆ&6·2â ¢Ð¢–b‚GF&vWBÖÖF6‚uæ‡GG3ó¢òòr’°¢G'’²7F'BÕ&ö6W72Ôf–ÆUF‚‚&Ö–7&÷6ögBÖVFvS¢"²GF&vWB’ÔW'&÷$7F–öâ7F÷Â÷WBÔçVÆÃ²&WGW&âGG'VRÐ¢6F6‚²w&—FRÕv&æ–ær‚$6÷VÆBæ÷B÷VâÖ–7&÷6ögBVFvRf–&÷Fö6öÂfÆÆ&6³¢"²EòäW†6WF–öâäÖW76vR’Ð¢Ð¢G'’°¢F6ÖDW†RÒ¦ö–âÕF‚FVçc¥7—7FVÕ&ö÷Bu7—7FVÓ3%Æ6ÖBæW†Rp¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚F6ÖDW†R’’²F6ÖDW†RÒv6ÖBæW†RrÐ¢FW66VEF&vWBÒGF&vWB×&WÆ6Rr"rÂr""p¢F6ÖD&w2Òrö27F'B""×6VFvRæW†R"r²FW66VEF&vWB²r"p¢7F'BÕ&ö6W72Ôf–ÆUF‚F6ÖDW†RÔ&wVÖVçDÆ—7BF6ÖD&w2Õv–æF÷u7G–ÆR†–FFVâÔW'&÷$7F–öâ7F÷Â÷WBÔçVÆÀ¢&WGW&âGG'VP¢Ò6F6‚²w&—FRÕv&æ–ær‚$6÷VÆBæ÷B÷VâÖ–7&÷6ögBVFvRf–6öÖÖæBfÆÆ&6³¢"²EòäW†6WF–öâäÖW76vR’Ð¢&WGW&âFfÇ6P§Ð ¦gVæ7F–öâ7F'BÔÆö6ÅF76W'fW"…¶–çEÒDÆ—7FVå÷'BÂ·7G&–æuÒD÷VåW&ÂÂ¶&ööÅÒE6¶—÷Vâ’°¢2Æ—7FVæW"8).XXŽ8¾z+®z¸¾8~8‹[~X¹^[è^88).8+ž8+8+Ž8:^8;Î8:ž8;ÎZÙ89~8:Þ8+¾8+ž8îX‰ÞiÉþXÉn8~˜^8(ž8¾8®8N8 ¢2Kº^™˜Þ8îZKiY~i˜.8("f–æÆÇ’8~[ø^8®ZÙ89~8:Þ8+¾8+ž8).XÎjÚ.8ž8(¾8 ¢GF7ÒFçVÆÀ¢G'’°¢GF7ÒæWrÔö&¦V7BæWBå6ö6¶WG2åF7Æ—7FVæW"…´æWBä•FG&W75Ó£¥'6R‚s#rãããr’ÂDÆ—7FVå÷'B¢GF7å7F'B‚¢G'’²·fö–EÒ…7F'BÔWFõ66†VGVÆW%&ö6W72„vWBÔVffV7F—fTÆæwVvR’’Ò6F6‚²w&—FRÕv&æ–ærEòäW†6WF–öâäÖW76vRÐ¢w&—FRÔ†÷7B%&W÷'D&–æFW"Æö6Â6W'fW"7F'FVBöâ#rããã¢DÆ—7FVå÷'B ¢w&—FRÔ†÷7B%&W÷'D&–æFW"U$Ã¢D÷VåW&Â ¢w&—FRÔ†÷7B%&W÷'D&–æFW"v–ÆÂ7F÷gFW"3Ö–çWFW2v—F†÷WB'&÷w6W"7F—f—G’âDb7&VF–öâ¦ö'26öçF–çVRWfVâ–bF†RF"—26Æ÷6VBâ ¢–b‚Öæ÷BE6¶—÷Vâ’²÷VâÕ&W÷'D&–æFW$'&÷w6W"D÷VåW&ÂÐ¢v†–ÆR‚GG'VR’°¢Fæ÷rÒ´FFUF–ÖUÓ£¥WF4æ÷p¢–b‚E67&—C¥6‡WFF÷vå&WVW7FVB’°¢w&—FRÔ†÷7B%&W÷'D&–æFW"6‡WFF÷vâ&WVW7FVBâ ¢'&V°¢Ð¢–b‚Öæ÷BE67&—C¤6Æ–VçDGF6†VBÖæB‚‚Fæ÷rÒE67&—C¥6W'fW%7F'FVEWF2’åF÷FÅ6V6öæG2ÖvRE67&—C¤æô6Æ–VçE7F'GWF–ÖV÷WE6V6öæG2’’°¢–b‚Öæ÷B…FW7BÔ7F—fU&VæFW$¦ö'2DÖöFR’’°¢w&—FRÔ†÷7B%&W÷'D&–æFW"'&÷w6W"v2æ÷B÷VæVB÷"GF6†VBâ7F÷–ærÆö6Â6W'fW"â ¢'&V°¢Ð¢Ð¢–b‚E67&—C¤6Æ–VçDGF6†VBÖæB‚‚Fæ÷rÒE67&—C¤Æ7D†V'F&VEWF2’åF÷FÅ6V6öæG2ÖvRE67&—C¤–FÆUF–ÖV÷WE6V6öæG2’’°¢–b…FW7BÔ7F—fU&VæFW$¦ö'2DÖöFR’°¢w&—FRÔ†÷7B%&W÷'D&–æFW"—2–FÆRÂ'WBDb7&VF–öâ¦ö"—27F–ÆÂ'Vææ–ærâ¶VW–ærÆö6Â6W'fW"Æ—fRâ ¢E67&—C¤Æ7D†V'F&VEWF2ÒFæ÷p¢ÒVÇ6R°¢w&—FRÔ†÷7B%&W÷'D&–æFW"–FÆRF–ÖV÷WB&V6†VBâ7F÷–ærÆö6Â6W'fW"â ¢'&V°¢Ð¢Ð ¢–b‚Öæ÷BGF7åVæF–ær‚’’°¢7F'BÕ6ÆVWÔÖ–ÆÆ—6V6öæG2#S ¢6öçF–çVP¢Ð ¢F6Æ–VçBÒGF7ä66WEF76Æ–VçB‚¢G'’°¢F6öçFW‡BÒ&VBÕF7‡GG6öçFW‡BF6Æ–VçBDÆ—7FVå÷'@¢GF‚ÒF6öçFW‡Bå&WVW7BåW&Âä'6öÇWFUF€¢–b‚GF‚å7F'G5v—F‚‚rö’òr’’²†æFÆRÔ’F6öçFW‡BÒVÇ6R²6W'fRÕ7FF–2F6öçFW‡BGF‚Ð¢Ò6F6‚°¢G'’°¢F7G‚Ò·67W7FöÖö&¦V7EÔ²—5F7ÒGG'VS²F76Æ–VçBÒF6Æ–VçC²F77G&VÒÒF6Æ–VçBävWE7G&VÒ‚“²&WVW7BÒFçVÆÃ²&W7öç6RÒ·67W7FöÖö&¦V7EÔ·ÒÐ¢FW'"Ò¶÷&FW&VEÔ²ö²ÒFfÇ6S²W'&÷"ÒEòäW†6WF–öâäÖW76vRÐ¢G7FGW2ÒS ¢–b‚EòäW†6WF–öâÖ—2µ7—7FVÒä&wVÖVçDW†6WF–öåÒ’²G7FGW2ÒCÐ¢VÇ6V–b…FW7BÕ7G'V7GW&T6öæfÆ–7DW'&÷"EòäW†6WF–öâ’²G7FGW2ÒC“²FW'%²v6öFRuÒÒw7G'V7GW&RÖ6öæfÆ–7BrÐ¢w&—FRÔ§6öå&W7öç6RF7G‚G7FGW2FW' ¢Ò6F6‚°¢G'’²F6Æ–VçBä6Æ÷6R‚’Ò6F6‚·Ð¢Ð¢Ð¢Ð¢Òf–æÆÇ’°¢2Šj¥”Nyº>Šin888¾šÎ8(ž8®8˜	®[‹Ž{X.K¨n8;¶Æ—7FVæW.‹[~X¹^ZKiY~8î8ž88(ž8~8(.iˆîzK®XÎjÚ.8ž8(¾8 ¢G'’²7F÷ÔWFõ66†VGVÆW%&ö6W72Ò6F6‚²Ð¢–b‚GF7’²G'’²GF7å7F÷‚’Ò6F6‚²ÒÐ¢Ð§Ð ¦gVæ7F–öâvWBÔg&VU÷'B°¢FÆ—7FVæW"ÒæWrÔö&¦V7BæWBå6ö6¶WG2åF7Æ—7FVæW"…´æWBä•FG&W75Ó£¥'6R‚s#rãããr’Â¢FÆ—7FVæW"å7F'B‚¢GÒFÆ—7FVæW"äÆö6ÄVæGö–çBå÷'@¢FÆ—7FVæW"å7F÷‚¢&WGW&âG §Ð  ¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ¢2cR7FvR"(	B†6R¢jIÎyú^x˜Ž8+ž88®88>89~8+~8:~88>88‚†–çWB†—7F÷'’¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ ¦gVæ7F–öâvWBÔ–çWD†—7F÷'•&ö÷B…·7G&–æuÒDÆæwVvR’°¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’v–çWBÖ†—7F÷'’r§Ð¦gVæ7F–öâvWBÕv÷&¶&öö´†—7F÷'”F—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢&WGW&â„¦ö–âÕF‚„vWBÔ–çWD†—7F÷'•&ö÷BDÆæwVvR’G6fUv÷&¶&öö´–B§Ð¦gVæ7F–öâvWBÕ6æ6†÷DF—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Bp¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&¶&öö´†—7F÷'”F—"DÆæwVvREv÷&¶&öö´–B’G6fU6æ6†÷D–B§Ð¦gVæ7F–öâvWBÔW†VÖW&Ä¦ö%&ö÷B…·7G&–æuÒDÆæwVvR’°¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’w7FFUÆ¦ö'2r§Ð ¦gVæ7F–öâw&—FRÔ†—7F÷'”WfVçB…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEG—RÂDFF’°¢2[^jÛN8+þ8*N8:8:ž8*N8;5Tž8).[¸>jÚ.8~8þ8þ8(8X[iÈž89^8*ž8:¾888;Î8Ž8ã8*N89ž8;>88ƒ89^8*8*N8:¾i»Ž‹ëÎ8(.ŠÎ8(þ8®8N8 ¢2x˜ŽjùN‹È>8¾[ø^Šh8§6æ6†÷BöÖæ–fW7Bö6ö×&—6öî8þXŠ^{XÎ‹zþ8~KùÞhÈ8^8(Î8(¾8 ¢&WGW&à§Ð ¦gVæ7F–öâvWBÕ6æ6†÷DÖæ–fW7B…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Bp¢F66†T¶W’Ò‚DÆæwVvR²wÂr²G6fUv÷&¶&öö´–B²wÂr²G6fU6æ6†÷D–B’åFôÆ÷vW$–çf&–çB‚¢GF‚Ò¦ö–âÕF‚„vWBÕ6æ6†÷DF—"DÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–B’vÖæ–fW7Bæ§6öâp¢–b‚E67&—C¥6æ6†÷DÖæ–fW7D66†Rä6öçF–ç4¶W’‚F66†T¶W’’’°¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚ÕF…G—RÆVb’²&WGW&âE67&—C¥6æ6†÷DÖæ–fW7D66†U²F66†T¶W•ÒÐ¢·fö–EÒE67&—C¥6æ6†÷DÖæ–fW7D66†Rå&VÖ÷fR‚F66†T¶W’¢Ð¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’’²&WGW&âFçVÆÂÐ¢G'’°¢FÖæ–fW7BÒ&VBÔ§6öäf–ÆRGF‚FçVÆÀ¢–b‚FçVÆÂÖæRFÖæ–fW7B’²E67&—C¥6æ6†÷DÖæ–fW7D66†U²F66†T¶W•ÒÒFÖæ–fW7BÐ¢&WGW&âFÖæ–fW7@¢Ò6F6‚²&WGW&âFçVÆÂÐ§Ð ¦gVæ7F–öâvWBÕ6æ6†÷D–G2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢FF—"ÒvWBÕv÷&¶&öö´†—7F÷'”F—"DÆæwVvRG6fUv÷&¶&öö´–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&â‚’Ð¢G7F×Ò´”òäF—&V7F÷'•Ó£¤vWDÆ7Ew&—FUF–ÖUWF2‚FF—"’åF–6·0¢F66†T¶W’Ò‚DÆæwVvR²wÂr²G6fUv÷&¶&öö´–B’åFôÆ÷vW$–çf&–çB‚¢F66†VBÒE67&—C¥6æ6†÷D–D66†U²F66†T¶W•Ð¢–b‚FçVÆÂÖæRF66†VBÖæB´–çCcEÒ„vWBÔFF&÷W'G’F66†VBw7F×rÓ’ÖWG7F×’°¢&WGW&â„vWBÔ'&’„vWBÔFF&÷W'G’F66†VBv–G2r‚’’¢Ð¢2Öæ–fW7Bæ§6öâ—2F†R6ö×ÆWF–öâÖ&¶W"âVæF–ærö7&6†VBF—&V7F÷&–W2&Ræ÷@¢2†—7F÷'’vVæW&F–öç2æB×W7Bæ÷B6öç7VÖRF†R&V6VçBÖvVæW&F–öâÆÆ÷væ6Rà¢2VçVÖW&FR6ö×ÆWF–öâÖ&¶W'2–âöæRF—&V7F÷'’vÆ²–ç7FVBöb—77V–æröæP¢2&VÖ÷FRFW7BÕF‚6ÆÂW"vVæW&F–öâöâ6†&VBv÷&·76Rà¢F–G2Ò„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"vÖæ–fW7Bæ§6öârÕ&V7W'6RÔFWF‚ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÀ¢v†W&RÔö&¦V7B²·7G&–æuÒEòäF—&V7F÷'’å&VçBägVÆÄæÖRÖW·7G&–æuÒFF—"ÒÀ¢6÷'BÔö&¦V7B²EòäF—&V7F÷'’äæÖRÒÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòäF—&V7F÷'’äæÖRÒ¢E67&—C¥6æ6†÷D–D66†U²F66†T¶W•ÒÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²7F×ÒG7F×²–G2ÒF–G2Ð¢&WGW&âF–G0§Ð ¦gVæ7F–öâvWBÕ6µ&Wf–Wu6æ6†÷B‚E7G'V7GW&RÂ·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–BÂ¶&ööÅÒD6†V6´÷WGWDW†—7G2ÒGG'VR’°¢G6²ÒvWBÕ6µ&V6÷&BE7G'V7GW&RE6´–@¢GFV×ÆFRÒvWBÕ6´VffV7F—fUFV×ÆFRDÆæwVvRG6°¢GF&vWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’GFV×ÆFRwF&vWG2r‚’’Âv†W&RÔö&¦V7B²¶&ööÅÒ„vWBÔFF&÷W'G’Eòw&WV—&VBrFfÇ6R’Ò¢–b‚GF&vWG2ä6÷VçBÖW’²GF&vWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’GFV×ÆFRwF&vWG2r‚’’Â6VÆV7BÔö&¦V7BÔf—'7B’Ð¢Ff–ævW'&–çG2Ò¶÷&FW&VEÔ·Ó²GF&vWE7FFW2Ò‚“²F6å7V&Ö—BÒ‚GF&vWG2ä6÷VçBÖwB¢f÷&V6‚‚GF&vWB–âGF&vWG2’°¢GF&vWD–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’GF&vWBwF&vWD–Brrr¢G&VF–æW72ÒvWBÔf–æÄ'V–ÆE&VF–æW72E7G'V7GW&RDÆæwVvR„vWBÔÆVv7•föÇVÖTg&öÕF&vWD–BDÆæwVvRGF&vWD–B’E6´–BD6†V6´÷WGWDW†—7G0¢Ff–ævW'&–çG5²GF&vWD–EÒÒ·7G&–æuÒ„vWBÔFF&÷W'G’G&VF–æW72v7W'&VçDf–ævW'&–çBrrr¢G&VG’Ò·7G&–æuÒ„vWBÔFF&÷W'G’G&VF–æW72vF—7Æ•7FFRrrr’ÖWv'V–ÇBp¢–b‚Öæ÷BG&VG’’²F6å7V&Ö—BÒFfÇ6RÐ¢GF&vWE7FFW2³Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢F&vWD–CÒGF&vWD–C²F—7Æ”æÖSÕ·7G&–æuÒ„vWBÔFF&÷W'G’GF&vWBvF—7Æ”æÖRrGF&vWD–B¢&VG“ÒG&VG“²F—7Æ•7FFSÕ·7G&–æuÒ„vWBÔFF&÷W'G’G&VF–æW72vF—7Æ•7FFRrrr¢f–ævW'&–çCÕ·7G&–æuÒ„vWBÔFF&÷W'G’G&VF–æW72v7W'&VçDf–ævW'&–çBrrr¢÷WGWEFcÕ·7G&–æuÒ„vWBÔFF&÷W'G’G&VF–æW72v÷WGWEFbrrr¢Ð¢Ð¢G7F÷&VBÒvWBÔFF&÷W'G’G6²w&Wf–Wrr…¶÷&FW&VEÔ·Ò¢G7F÷&VE7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBw7FGW2rvG&gBr¢–b‚G7F÷&VE7FGW2Öæ÷F–â‚vG&gBrÂv–â×&Wf–WrrÂv&÷fVBrÂv6†ævW2×&WVW7FVBr’’²G7F÷&VE7FGW2ÒvG&gBrÐ¢G7V&Ö—GFVBÒvWBÔFF&÷W'G’G7F÷&VBw7V&Ö—GFVDf–ævW'&–çG2r…¶÷&FW&VEÔ·Ò¢G7FÆRÒFfÇ6P¢–b‚G7F÷&VE7FGW2Ö–â‚v–â×&Wf–WrrÂv&÷fVBr’’°¢f÷&V6‚‚GF&vWB–âGF&vWG2’°¢GF&vWD–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’GF&vWBwF&vWD–Brrr¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G7V&Ö—GFVBGF&vWD–Brr’ÖæR·7G&–æuÒFf–ævW'&–çG5²GF&vWD–EÒ’²G7FÆRÒGG'VS²'&V²Ð¢Ð¢Ð¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢6´–CÒE6´–C²7FGW3ÒB†–b‚G7FÆR’²w7FÆRrÒVÇ6R²G7F÷&VE7FGW2Ò“²7F÷&VE7FGW3ÒG7F÷&VE7FGW0¢6å7V&Ö—CÕ¶&ööÅÒF6å7V&Ö—C²7FÆSÕ¶&ööÅÒG7FÆS²F&vWE7FFW3Ô‚GF&vWE7FFW2“²7W'&VçDf–ævW'&–çG3ÒFf–ævW'&–çG0¢7V&Ö—GFVDf–ævW'&–çG3ÒG7V&Ö—GFVC²7V&Ö—GFVDCÕ·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBw7V&Ö—GFVDBrrr¢7V&Ö—GFVD'“Õ·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBw7V&Ö—GFVD'’rrr“²&÷fVDCÕ·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBv&÷fVDBrrr¢&÷fVD'“Õ·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBv&÷fVD'’rrr“²æ÷FSÕ·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBvæ÷FRrrr¢WfVçG3Ô„vWBÔ'&’„vWBÔFF&÷W'G’G7F÷&VBvWfVçG2r‚’’Â6VÆV7BÔö&¦V7BÔÆ7B¢Ð§Ð ¦gVæ7F–öâ6÷’Õ6µ&Wf–Wtf–ævW'&–çG2‚Df–ævW'&–çG2’°¢F6÷’Ò¶÷&FW&VEÔ·Ð¢–b‚FçVÆÂÖWDf–ævW'&–çG2’²&WGW&âF6÷’Ð¢–b‚Df–ævW'&–çG2Ö—2´6öÆÆV7F–öç2ä”F–7F–öæ'•Ò’°¢f÷&V6‚‚F¶W’–â‚Df–ævW'&–çG2ä¶W—2’’²F6÷•µ·7G&–æuÒF¶W•ÒÒ·7G&–æuÒDf–ævW'&–çG5²F¶W•ÒÐ¢&WGW&âF6÷¢Ð¢f÷&V6‚‚G&÷W'G’–â‚Df–ævW'&–çG2å4ö&¦V7Bå&÷W'F–W2’’²F6÷•µ·7G&–æuÒG&÷W'G’äæÖUÒÒ·7G&–æuÒG&÷W'G’åfÇVRÐ¢&WGW&âF6÷§Ð ¦gVæ7F–öâ–çfö¶RÕ6µ&Wf–Wt7F–öâ…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–BÂD&öG’’°¢F–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6´–Bw6´–Bp¢2fö–BF†RæÖRF7F–öæ¢–çfö¶RÕv—F„Æö6²†266RÖ–ç6Vç6—F—fRD7F–öæ ¢2&ÖWFW"Âv†–6‚v÷VÆB6†F÷rF†—2fÇVR–ç6–FRF†R×WFF–öâ67&—F&Æö6²à¢G&Wf–Wt7F–öâÒ…·7G&–æuÒ„vWBÔFF&÷W'G’D&öG’v7F–öârrr’’åG&–Ò‚’åFôÆ÷vW$–çf&–çB‚¢–b‚G&Wf–Wt7F–öâÖæ÷F–â‚w7V&Ö—BrÂv&÷fRrÂw&WVW7BÖ6†ævW2rÂw&V÷Vâr’’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚w&Wf–Wr7F–öâ8ÎKˆÞjÚ>8~8ž8"r’Ð¢Fæ÷FRÒ…·7G&–æuÒ„vWBÔFF&÷W'G’D&öG’væ÷FRrrr’’åG&–Ò‚¢–b‚Fæ÷FRäÆVæwF‚ÖwB’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~8:Î89>8:^8;Î8+>8:8;>88Ž8óih~ZÙ~Kº^Xh^8~XZ^X©¾8~8n8þ88^8N8"r’Ð¢–b‚G&Wf–Wt7F–öâÖWw&WVW7BÖ6†ævW2rÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Fæ÷FR’’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~[zî8~h‹¾8~ynyK8).XZ^X©¾8~8n8þ88^8N8"r’Ð¢F7F÷"Ò…·7G&–æuÒ„vWBÔFF&÷W'G’D&öG’v7F÷"rFVçc¥U4U$äÔR’’åG&–Ò‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7F÷"’’²F7F÷"ÒvÆö6Â×W6W"rÐ¢–b‚F7F÷"äÆVæwF‚ÖwB#’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~z+®Š¨Þˆ^YÞ8ó#ih~ZÙ~Kº^Xh^8~XZ^X©¾8~8n8þ88^8N8"r’Ð¢2Db&VF–æW72WfÇVF–öâÖ’–çfö¶RF†RDb'VçF–ÖRâ¶VWF†Bv÷&²÷WBö`¢2F†R7G'V7GW&Rf–ÆRÆö6²ÂF†Vâ&V¦V7BF†R7F–öâ–bæ÷F†W"w&—FW"6†ævV@¢2F†R7G'V7GW&R&Vf÷&RF†R6†÷'BW'6—7FVæ6RG&ç67F–öâ&Vv–ç2à¢G&Wf–Wu7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢FW‡V7FVE7G'V7GW&UWFFVDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’G&Wf–Wu7G'V7GW&RwWFFVDBrrr¢G6æ6†÷BÒvWBÕ6µ&Wf–Wu6æ6†÷BG&Wf–Wu7G'V7GW&RDÆæwVvRF–BGG'VP¢&WGW&âWFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7G'V7GW&R¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G7G'V7GW&RwWFFVDBrrr’ÖæRFW‡V7FVE7G'V7GW&UWFFVDB’°¢F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚~‹8~iižjx¾h‰8ÎYÎi˜.8¾i»Nik8^8(Î8î8~8þ8.iÈikx«nhX¾8).z+®Š¨Þ8~8n88(.8nKˆ[ªni8ÞKÙÎ8~8n8þ88^8N8"r¢Ð¢G6²ÒvWBÕ6µ&V6÷&BG7G'V7GW&RF–@¢G7F÷&VBÒvWBÔFF&÷W'G’G6²w&Wf–Wrr…¶÷&FW&VEÔ·Ò¢Fæ÷rÒæWrÔæ÷t—6ð¢7v—F6‚‚G&Wf–Wt7F–öâ’°¢w7V&Ö—Br°¢–b‚Öæ÷B¶&ööÅÒG6æ6†÷Bæ6å7V&Ö—B’²F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚~[ø^šŽ8îhùX{®yJ…Dn8).8ž8ž8niÈik8¾8~8n8¾8(ž8:Î89>8:^8;Î8ŽhùX{®8~8n8þ88^8N8"r’Ð¢6WBÔæ÷FU&÷W'G’G7F÷&VBw7FGW2rv–â×&Wf–Wrs²6WBÔæ÷FU&÷W'G’G7F÷&VBw7V&Ö—GFVDf–ævW'&–çG2r„6÷’Õ6µ&Wf–Wtf–ævW'&–çG2G6æ6†÷Bæ7W'&VçDf–ævW'&–çG2¢6WBÔæ÷FU&÷W'G’G7F÷&VBw7V&Ö—GFVDBrFæ÷s²6WBÔæ÷FU&÷W'G’G7F÷&VBw7V&Ö—GFVD'’rF7F÷ ¢6WBÔæ÷FU&÷W'G’G7F÷&VBv&÷fVDBrrs²6WBÔæ÷FU&÷W'G’G7F÷&VBv&÷fVD'’rrp¢Ð¢v&÷fRr°¢–b…·7G&–æuÒG6æ6†÷Bç7FGW2ÖæRv–â×&Wf–Wrr’²F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚~8:Î89>8:^8;ÎhùX{®[èÎ8~8Xh^Zëž8Îi»Nik8^8(Î8n8N8®8N‹8~iiž888).h›þŠ¨Þ8~8Þ8î8ž8"r’Ð¢6WBÔæ÷FU&÷W'G’G7F÷&VBw7FGW2rv&÷fVBs²6WBÔæ÷FU&÷W'G’G7F÷&VBv&÷fVDBrFæ÷s²6WBÔæ÷FU&÷W'G’G7F÷&VBv&÷fVD'’rF7F÷ ¢Ð¢w&WVW7BÖ6†ævW2r°¢–b…·7G&–æuÒG6æ6†÷Bç7FGW2Öæ÷F–â‚v–â×&Wf–WrrÂv&÷fVBr’’²F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚~8:Î89>8:^8;ÎKŠÞ8î8þ8þh›þŠ¨ÞkˆŽ8þ8î‹8~iiž888).[zî8~h‹¾8¾8î8ž8"r’Ð¢6WBÔæ÷FU&÷W'G’G7F÷&VBw7FGW2rv6†ævW2×&WVW7FVBs²6WBÔæ÷FU&÷W'G’G7F÷&VBv&÷fVDBrrs²6WBÔæ÷FU&÷W'G’G7F÷&VBv&÷fVD'’rrp¢Ð¢w&V÷Vâr²6WBÔæ÷FU&÷W'G’G7F÷&VBw7FGW2rvG&gBs²6WBÔæ÷FU&÷W'G’G7F÷&VBw7V&Ö—GFVDf–ævW'&–çG2r…¶÷&FW&VEÔ·Ò“²6WBÔæ÷FU&÷W'G’G7F÷&VBw7V&Ö—GFVDBrrs²6WBÔæ÷FU&÷W'G’G7F÷&VBw7V&Ö—GFVD'’rrs²6WBÔæ÷FU&÷W'G’G7F÷&VBv&÷fVDBrrs²6WBÔæ÷FU&÷W'G’G7F÷&VBv&÷fVD'’rrrÐ¢Ð¢6WBÔæ÷FU&÷W'G’G7F÷&VBvæ÷FRrFæ÷FP¢FWfVçDf–ævW'&–çG2Ò6÷’Õ6µ&Wf–Wtf–ævW'&–çG2„vWBÔFF&÷W'G’G7F÷&VBw7V&Ö—GFVDf–ævW'&–çG2r…¶÷&FW&VEÔ·Ò’¢G&Wf–WtWfVçBÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢WfVçD–CÒw&Wf–WrÒr²´wV–EÓ£¤æWtwV–B‚’åFõ7G&–ær‚târ’å7V'7G&–ærƒÃb“²7F–öãÒG&Wf–Wt7F–öã²7FGW3Õ·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBw7FGW2rvG&gBr¢7F÷#ÒF7F÷#²æ÷FSÒFæ÷FS²CÒFæ÷s²f–ævW'&–çG3ÒFWfVçDf–ævW'&–çG0¢Ð¢FWfVçG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’G7F÷&VBvWfVçG2r‚’’’²‚G&Wf–WtWfVçB¢6WBÔæ÷FU&÷W'G’G7F÷&VBvWfVçG2r‚FWfVçG2Â6VÆV7BÔö&¦V7BÔÆ7B¢6WBÔæ÷FU&÷W'G’G6²w&Wf–WrrG7F÷&VC²6WBÔæ÷FU&÷W'G’G6²wWFFVDBrFæ÷p¢2F†R&VF–æW726æ6†÷B&÷fRÇ&VG’6öçF–ç2F†RW†7Bf–ævW'&–çG0¢2fÆ–FFVB'’F†—27F–öââ&V6ö×WF–ær—B†W&R&VæFW'2WfW'’6÷W&6RD`¢26V6öæBF–ÖRæBÖ¶W26–ævÆR&Wf–Wr6Æ–6²VææV6W76&–Ç’W‡Vç6—fRà¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bw7FGW2r…·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBw7FGW2rvG&gBr’¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bw7F÷&VE7FGW2r…·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBw7FGW2rvG&gBr’¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bw7FÆRrFfÇ6P¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bw7V&Ö—GFVDf–ævW'&–çG2r„6÷’Õ6µ&Wf–Wtf–ævW'&–çG2„vWBÔFF&÷W'G’G7F÷&VBw7V&Ö—GFVDf–ævW'&–çG2r…¶÷&FW&VEÔ·Ò’’¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bw7V&Ö—GFVDBr…·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBw7V&Ö—GFVDBrrr’¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bw7V&Ö—GFVD'’r…·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBw7V&Ö—GFVD'’rrr’¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bv&÷fVDBr…·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBv&÷fVDBrrr’¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bv&÷fVD'’r…·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBv&÷fVD'’rrr’¢6WBÔæ÷FU&÷W'G’G6æ6†÷Bvæ÷FRr…·7G&–æuÒ„vWBÔFF&÷W'G’G7F÷&VBvæ÷FRrrr’¢6WBÔæ÷FU&÷W'G’G6æ6†÷BvWfVçG2r„vWBÔ'&’„vWBÔFF&÷W'G’G7F÷&VBvWfVçG2r‚’’Â6VÆV7BÔö&¦V7BÔÆ7B¢&WGW&âG6æ6†÷@¢Ð§Ð ¦gVæ7F–öâ6ÆV"Õ6æ6†÷E'VçF–ÖT66†W2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÒrr’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢Gv÷&¶&ööµ&Vf—‚Ò‚DÆæwVvR²wÂr²G6fUv÷&¶&öö´–B²wÂr’åFôÆ÷vW$–çf&–çB‚¢G6æ6†÷D¶W’Ò–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6æ6†÷D–B’’²rrÒVÇ6R²‚Gv÷&¶&ööµ&Vf—‚²„76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Br’åFôÆ÷vW$–çf&–çB‚’’Ð¢f÷&V6‚‚F¶W’–â‚E67&—C¥6æ6†÷DÖæ–fW7D66†Rä¶W—2’’°¢–b‚‚G6æ6†÷D¶W’ÖæB·7G&–æuÒF¶W’ÖWG6æ6†÷D¶W’’Ö÷"‚Öæ÷BG6æ6†÷D¶W’ÖæB·7G&–æuÒF¶W’ÖÆ–¶R‚Gv÷&¶&ööµ&Vf—‚²r¢r’’’°¢·fö–EÒE67&—C¥6æ6†÷DÖæ–fW7D66†Rå&VÖ÷fR‚F¶W’¢Ð¢Ð¢f÷&V6‚‚F¶W’–â‚E67&—C¥f—7VÄ†6„66†Rä¶W—2’’°¢–b‚‚G6æ6†÷D¶W’ÖæB·7G&–æuÒF¶W’ÖÆ–¶R‚G6æ6†÷D¶W’²wÂ¢r’’Ö÷"‚Öæ÷BG6æ6†÷D¶W’ÖæB·7G&–æuÒF¶W’ÖÆ–¶R‚Gv÷&¶&ööµ&Vf—‚²r¢r’’’°¢·fö–EÒE67&—C¥f—7VÄ†6„66†Rå&VÖ÷fR‚F¶W’¢Ð¢Ð¢6ÆV"Õ6æ6†÷E7VÖÖ'”66†RDÆæwVvRG6fUv÷&¶&öö´–@§Ð ¦gVæ7F–öâf–æBÕ6æ6†÷D'•6÷W&6T†6‚…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6÷W&6T†6‚’°¢GF&vWBÒæ÷&ÖÆ—¦RÔf–ÆT†6‚E6÷W&6T†6€¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GF&vWB’’²&WGW&ârrÐ¢f÷&V6‚‚F–B–â„vWBÕ6æ6†÷D–G2DÆæwVvREv÷&¶&öö´–B’’°¢FÒÒvWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BF–@¢–b‚FçVÆÂÖWFÒ’²6öçF–çVRÐ¢–b‚„æ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’FÒw6÷W&6T†6‚rrr’’’ÖWGF&vWB’²&WGW&âF–BÐ¢Ð¢&WGW&ârp§Ð ¦gVæ7F–öâvWBÕ6æ6†÷E6÷W&6U7FFR…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢2cRÔ•bÓS¢xûîxšž8îiÈžxJ8ò–Ö×WF&ÆR8¢Öæ–fW7B8¾8þi»Ž8¾8®8N8 ¢FF—"ÒvWBÕ6æ6†÷DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢G7FFUF‚Ò¦ö–âÕF‚FF—"w6÷W&6R×7FFRæ§6öâp¢FÖæ–fW7BÒvWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢FW‡FVç6–öâÒ´”òåF…Ó£¤vWDW‡FVç6–öâ…·7G&–æuÒ„vWBÔFF&÷W'G’FÖæ–fW7Bw&VÆF—fUF‚rrr’’åFôÆ÷vW$–çf&–çB‚¢–b‚FW‡FVç6–öâÖæ÷F–â‚rç†Ç7‚rÂrç†Ç6ÒrÂræFö7‚rÂrçG‚rÂrçFbr’’²FW‡FVç6–öâÒB†–b…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚FF—"w6÷W&6RçFbr’’²rçFbrÒVÇ6V–b…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚FF—"w6÷W&6RçG‚r’’²rçG‚rÒVÇ6V–b…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚FF—"w6÷W&6RæFö7‚r’’²ræFö7‚rÒVÇ6V–b…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚FF—"w6÷W&6Rç†Ç6Òr’’²rç†Ç6ÒrÒVÇ6R²rç†Ç7‚rÒ’Ð¢G6÷W&6UF‚Ò¦ö–âÕF‚FF—"‚'6÷W&6RFW‡FVç6–öâ"¢FW†—7G2ÒFW7BÕF‚ÔÆ—FW&ÅF‚G6÷W&6UF€¢G7FFRÒFçVÆÀ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚G7FFUF‚’²G'’²G7FFRÒ&VBÔ§6öäf–ÆRG7FFUF‚FçVÆÂÒ6F6‚²ÒÐ¢&WGW&â¶÷&FW&VEÔ°¢6÷W&6U&WF–æVBÒ¶&ööÅÒFW†—7G0¢6÷W&6UF‚ÒG6÷W&6UF€¢&VÖ÷fVDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRw&VÖ÷fVDBrrr¢&VÖ÷fVE&V6öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRw&VÖ÷fVE&V6öârrr¢Ð§Ð ¦gVæ7F–öâ6WBÕ6æ6†÷E6÷W&6U7FFR…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ¶&ööÅÒE&WF–æVBÂ·7G&–æuÒE&V6öâ’°¢FF—"ÒvWBÕ6æ6†÷DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&âÐ¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚FF—"w6÷W&6R×7FFRæ§6öâr’…¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ²6÷W&6U&WF–æVBÒE&WF–æV@¢&VÖ÷fVDBÒB†–b‚E&WF–æVB’²rrÒVÇ6R²æWrÔæ÷t—6òÒ¢&VÖ÷fVE&V6öâÒB†–b‚E&WF–æVB’²rrÒVÇ6R²·7G&–æuÒE&V6öâÒ¢Ò§Ð ¢2ÒÒÒÒ–ç2òÆV6W2ÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÐ ¦gVæ7F–öâvWBÕ6æ6†÷E–äF—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢&WGW&â„¦ö–âÕF‚„vWBÕ6æ6†÷DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’w–ç2r§Ð¦gVæ7F–öâvWBÕ6æ6†÷DÆV6TF—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢&WGW&â„¦ö–âÕF‚„vWBÕ6æ6†÷DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’vÆV6W2r§Ð ¦gVæ7F–öâæWrÕ6æ6†÷E–â…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒE–äæÖRÂDFF’°¢2cRÔ•bÓ#¢KùÞŠÛ~8þ89^8*8*N8:¾8îKÙÎh‰8;¾X˜®™šN8~ŠŽ8ž8.˜XÞX‰~8îi»Ž8Þhù¾8Ž8þ8~8®8N8 ¢G'’°¢G6fU–äæÖRÒ76W'BÕ6fU7F÷&vU6VvÖVçBE–äæÖRw–äæÖRp¢–b‚FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’’²&WGW&âFfÇ6RÐ¢FF—"ÒvWBÕ6æ6†÷E–äF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚FF—"‚'³Òæ§6öâ"ÖbG6fU–äæÖR’’DFF¢&WGW&âGG'VP¢Ò6F6‚²&WGW&âFfÇ6RÐ§Ð ¦gVæ7F–öâ&VÖ÷fRÕ6æ6†÷E–â…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒE–äæÖR’°¢G'’°¢GF‚Ò¦ö–âÕF‚„vWBÕ6æ6†÷E–äF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’‚'³Òæ§6öâ"ÖbE–äæÖR¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâvWBÕ6æ6†÷E–ç2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢FF—"ÒvWBÕ6æ6†÷E–äF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&â‚’Ð¢&WGW&â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòä&6TæÖRÒ§Ð ¦gVæ7F–öâ6WBÔÖçVÅ6æ6†÷E–â…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ¶&ööÅÒE–ææVB’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Bp¢F†—7F÷'”Æö6²Ò¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’vÆö6·5Æ†—7F÷'’Ö6ÆVçWæÆö6²p¢&WGW&â–çfö¶RÕv—F„Æö6²F†—7F÷'”Æö6²°¢–b‚FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–B’’°¢F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~KùÞŠÛ~8ž8(¾[^jÛNx˜Ž8ÎŠh¾8N8¾8(®8î8¾8)>8"r¢Ð¢–b‚E–ææVB’°¢–b‚Öæ÷B„æWrÕ6æ6†÷E–âDÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–BvÖçVÂr…¶÷&FW&VEÔ²–ææVDBÒæWrÔæ÷t—6òÒ’’’°¢F‡&÷r~[^jÛN8îKùÞŠÛ~h8^Z8).KùÞZÙŽ8~8Þ8î8¾8)>8~8~8þ8"p¢Ð¢ÒVÇ6R°¢&VÖ÷fRÕ6æ6†÷E–âDÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–BvÖçVÂp¢Ð¢·fö–EÒ…WFFRÕ6æ6†÷E7VÖÖ'”66†TVçG'’DÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–B¢&WGW&âGG'VP¢Ð§Ð ¦gVæ7F–öâæWrÕ6æ6†÷DÆV6R…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEW'÷6RÂ·7G&–æuÒD¦ö$–BÂ¶–çEÒDÖ–çWFW5fÆ–BÒ3Â·7G&–æuÒEfW'6–öä–BÒrr’°¢G'’°¢G6fUW'÷6RÒ76W'BÕ6fU7F÷&vU6VvÖVçBEW'÷6RwW'÷6Rp¢G6fT¦ö$–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBD¦ö$–Bv¦ö$–Bp¢–b‚FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’’²&WGW&ârrÐ¢FF—"ÒvWBÕ6æ6†÷DÆV6TF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢FæÖRÒ‚w³Õ÷³Òæ§6öârÖbG6fUW'÷6RÂG6fT¦ö$–B¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚FF—"FæÖR’…¶÷&FW&VEÔ°¢¦ö$–BÒD¦ö$–C²W'÷6RÒEW'÷6S²fW'6–öä–BÒEfW'6–öä–@¢7&VFVDBÒæWrÔæ÷t—6ó²†V'F&VDBÒæWrÔæ÷t—6ð¢W‡—&W4BÒ…´FFUF–ÖUÓ£¥WF4æ÷räFDÖ–çWFW2‚DÖ–çWFW5fÆ–B’åFõ7G&–ær‚vòr’¢4æÖRÒFVçc¤4ôÕUDU$äÔP¢Ò¢&WGW&âFæÖP¢Ò6F6‚²&WGW&ârrÐ§Ð ¦gVæ7F–öâ&VÖ÷fRÕ6æ6†÷DÆV6R…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒDÆV6TæÖR’°¢G'’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚DÆV6TæÖR’’²&WGW&âÐ¢GF‚Ò¦ö–âÕF‚„vWBÕ6æ6†÷DÆV6TF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’DÆV6TæÖP¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâFW7BÔÆV6T7F—fR‚DÆV6Tf–ÆR’°¢G'’°¢F¢Ò&VBÔ§6öäf–ÆRDÆV6Tf–ÆRägVÆÄæÖRFçVÆÀ¢FW‡Ò·7G&–æuÒ„vWBÔFF&÷W'G’F¢vW‡—&W4Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FW‡’’²&WGW&âFfÇ6RÐ¢&WGW&â…´FFUF–ÖUÓ£¥'6R‚FW‡’åFõVæ—fW'6ÅF–ÖR‚’ÖwB´FFUF–ÖUÓ£¥WF4æ÷r¢Ò6F6‚²&WGW&âFfÇ6RÐ§Ð ¦gVæ7F–öâvWBÔ7F—fTÆV6T6÷VçB…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢FF—"ÒvWBÕ6æ6†÷DÆV6TF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&âÐ¢FâÒ ¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢–b…FW7BÔÆV6T7F—fRFb’²Fâ²²ÒVÇ6R²G'’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FbägVÆÄæÖRÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒ6F6‚²ÒÐ¢Ð¢&WGW&âFà§Ð  ¦gVæ7F–öâ6ÆV"ÔW‡—&VDÆV6W2…·7G&–æuÒDÆæwVvR’°¢26æ6†÷BæB6öçFVçBÕDbÆV6W2&÷F‚W‡—&RgFW"&æ÷&ÖÂFW&Ö–æF–öâà¢G'’°¢G&ö÷BÒvWBÔ–çWD†—7F÷'•&ö÷BDÆæwVvP¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚G&ö÷B’°¢f÷&V6‚‚Gv"–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚G&ö÷BÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢f÷&V6‚‚G6â–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚Gv"ägVÆÄæÖRÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢·fö–EÒ„vWBÔ7F—fTÆV6T6÷VçBDÆæwVvRGv"äæÖRG6âäæÖR¢Ð¢Ð¢Ð¢Ò6F6‚²Ð¢G'’°¢F6öçFVçE&ö÷BÒ¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’v6öçFVçB×Fbp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚F6öçFVçE&ö÷B’°¢f÷&V6‚‚Gv"–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚F6öçFVçE&ö÷BÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢f÷&V6‚‚GfW"–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚Gv"ägVÆÄæÖRÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢FÆV6TF—"Ò¦ö–âÕF‚GfW"ägVÆÄæÖRvÆV6W2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FÆV6TF—"’’²6öçF–çVRÐ¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FÆV6TF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢–b‚Öæ÷B…FW7BÔÆV6T7F—fRFb’’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FbägVÆÄæÖRÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ð¢Ð¢Ð¢Ð¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâVç7W&RÕ6æ6†÷DÖWFFF…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE&VÆF—fUF‚Â·7G&–æuÒD6FVv÷'’Â·7G&–æuÒE6÷W&6UF‚Â·7G&–æuÒE6÷W&6T†6‚Â·7G&–æuÒD6GW&U&V6öâ’°¢2cRÜ*t3¢8:8+þ88~8;Î8+þ8î˜xÞŠH~hé.™šN888).h¸^[Ù>8ž8(¾8.XZ^X©¾89^8*8*N8:¾8îz+®KùÞ8ò6GW&RÕ&VæFW$–çWB8î[Ûžyºî8 ¢–b‚Öæ÷B…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvæ÷BÖ&÷fVBs²6æ6†÷D–BÒrrÒÐ¢F†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚E6÷W&6T†6€¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F†6‚’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvæòÖ†6‚s²6æ6†÷D–BÒrrÒÐ ¢FÆö6µF‚Ò¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’‚&Æö6·5Ç6æ6†÷E÷³ÒæÆö6²"ÖbEv÷&¶&öö´–B¢&WGW&â–çfö¶RÕv—F„Æö6²FÆö6µF‚°¢FW†—7F–ærÒf–æBÕ6æ6†÷D'•6÷W&6T†6‚DÆæwVvREv÷&¶&öö´–BF†6€¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FW†—7F–ær’’°¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRv–çWBç6æ6†÷BæFVGWÆ–6FVBr…¶÷&FW&VEÔ²v÷&¶&öö´–BÒEv÷&¶&öö´–C²6æ6†÷D–BÒFW†—7F–æs²6÷W&6T†6‚ÒF†6ƒ²6GW&U&V6öâÒD6GW&U&V6öâÒ¢&WGW&â¶÷&FW&VEÔ²ö²ÒGG'VS²6æ6†÷D–BÒFW†—7F–æs²—4æWrÒFfÇ6RÐ¢Ð¢F–G2Ò„vWBÕ6æ6†÷D–G2DÆæwVvREv÷&¶&öö´–B¢G&Wf–÷W2ÒB†–b‚F–G2ä6÷VçBÖwB’²·7G&–æuÒF–G5²ÓÒÒVÇ6R²rrÒ¢G&VçBÒvWBÕv÷&¶&öö´†—7F÷'”F—"DÆæwVvREv÷&¶&öö´–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G&VçB’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚G&VçBÔf÷&6RÂ÷WBÔçVÆÂÐ¢F7&VFVBÒæWrÕVæ—VTF—&V7F÷'’G&VçB²æWrÕ&$–BÐ¢G6æ6†÷D–BÒ·7G&–æuÒF7&VFVBæ–@¢F—FVÒÒFçVÆÀ¢G'’²F—FVÒÒvWBÔ—FVÒÔÆ—FW&ÅF‚E6÷W&6UF‚ÔW'&÷$7F–öâ7F÷Ò6F6‚²Ð¢2cRÜ*s2ã#¢Öæ–fW7Bæ§6öâ8îZÙŽYÊŽ8ÎZèÎh‰89î8;Î8*¾8;Î8.XXŽ8²VæF–ær8Ž8~8ni»Ž8Þ8¢26÷W&6R8î8+>89N8;Î8;¾jIÎŠ‹Î8Î{X.8(þ8>8n8¾8(’Öæ–fW7Bæ§6öâ8ŽiKžYÞ8ž8(²„6ö×ÆWFRÕ6æ6†÷Bž8 ¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚…·7G&–æuÒF7&VFVBçF‚’vÖæ–fW7BçVæF–æræ§6öâr’…¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ¢6æ6†÷D–BÒG6æ6†÷D–@¢v÷&¶&öö´–BÒEv÷&¶&öö´–@¢ÆæwVvRÒDÆæwVvP¢6FVv÷'’ÒD6FVv÷'¢&VÆF—fUF‚ÒE&VÆF—fUF€¢6÷W&6T†6„Æv÷&—F†ÒÒu4„Ó#Sbp¢6÷W&6T†6‚ÒF†6€¢6÷W&6U6—¦RÒB†–b‚F—FVÒ’²F—FVÒäÆVæwF‚ÒVÇ6R²Ò¢6÷W&6TÆ7Ew&—FUWF5F–6·2ÒB†–b‚F—FVÒ’²·7G&–æuÒF—FVÒäÆ7Ew&—FUF–ÖUWF2åF–6·2ÒVÇ6R²rrÒ¢6÷W&6TÖöF–f–VDBÒB†–b‚F—FVÒ’²F—FVÒäÆ7Ew&—FUF–ÖRåFõ7G&–ær‚w———’ÔÔÒÖFED„ƒ¦ÖÓ§77§§¢r’ÒVÇ6R²rrÒ¢FWFV7FVDBÒæWrÔæ÷t—6ð¢6GW&U&V6öâÒD6GW&U&V6öà¢&Wf–÷W56æ6†÷D–BÒG&Wf–÷W0¢7FGW2Òv6ö×ÆWFRp¢'6W%fW'6–öâÒ¢6GW&VD'’Ò¶÷&FW&VEÔ²4æÖRÒFVçc¤4ôÕUDU$äÔS²W6W$æÖRÒ"FVçc¥U4U$DôÔ”åÂFVçc¥U4U$äÔR"Ð¢Ò¢6WBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvREv÷&¶&öö´–BG6æ6†÷D–BFfÇ6Rvæ÷BÖ6GW&VBp¢&WGW&â¶÷&FW&VEÔ²ö²ÒGG'VS²6æ6†÷D–BÒG6æ6†÷D–C²—4æWrÒGG'VS²VæF–ærÒGG'VRÐ¢Ð§Ð ¦gVæ7F–öâ6ö×ÆWFRÕ6æ6†÷B…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢2VæF–ær8¾8(’Öæ–fW7Bæ§6öâ8ŽiKžYÞ8~8nZèÎh‰8^8¾8(¾8.Kº^[èÎ8>8îjIÎyú^x˜Ž8þXø.xZ~Xúþˆ;Þ8¾8®8(¾8 ¢FF—"ÒvWBÕ6æ6†÷DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢GVæF–ærÒ¦ö–âÕF‚FF—"vÖæ–fW7BçVæF–æræ§6öâp¢Ff–æÂÒ¦ö–âÕF‚FF—"vÖæ–fW7Bæ§6öâp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚Ff–æÂ’²&WGW&âGG'VRÐ¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚GVæF–ær’’²&WGW&âFfÇ6RÐ¢G'’°¢Ö÷fRÔ—FVÒÔÆ—FW&ÅF‚GVæF–ærÔFW7F–æF–öâFf–æÂÔf÷&6P¢6ÆV"Õ6æ6†÷E'VçF–ÖT66†W2DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRv–çWBç6æ6†÷Bæ7&VFVBr…¶÷&FW&VEÔ²v÷&¶&öö´–BÒEv÷&¶&öö´–C²6æ6†÷D–BÒE6æ6†÷D–BÒ¢&WGW&âGG'VP¢Ò6F6‚²&WGW&âFfÇ6RÐ§Ð ¦gVæ7F–öâ6fRÕ6æ6†÷E6÷W&6Tf–ÆR…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒE6÷W&6UF‚Â·7G&–æuÒE6÷W&6T†6‚’°¢2cRÜ*s2ã"h˜¾šcBÓƒ¢8+>89N8;ÂÓâ88þ88>8+~8:^jIÎŠ‹ÂÓâ[^jÛN89^8*ž8:¾888Žz{¾X¹RÓâXhÞ88þ88>8+~8:^8 ¢–b‚Öæ÷B…FW7BÕ6÷W&6U&WFVçF–öäVæ&ÆVB’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒw&WFVçF–öâÖæ÷BÖ&÷fVBrÒÐ¢FF—"ÒvWBÕ6æ6†÷DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒw6æ6†÷BÖÖ—76–ærrÒÐ¢FW‡FVç6–öâÒ´”òåF…Ó£¤vWDW‡FVç6–öâ‚E6÷W&6UF‚’åFôÆ÷vW$–çf&–çB‚¢–b‚FW‡FVç6–öâÖæ÷F–â‚rç†Ç7‚rÂrç†Ç6ÒrÂræFö7‚rÂrçG‚rÂrçFbr’’²FW‡FVç6–öâÒrç†Ç7‚rÐ¢FFW7BÒ¦ö–âÕF‚FF—"‚'6÷W&6RFW‡FVç6–öâ"¢FW‡V7FVD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚E6÷W&6T†6€¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FFW7B’°¢FW†—7F–æt†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbFFW7B¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FW†—7F–æt†6‚’ÖæBFW†—7F–æt†6‚ÖWFW‡V7FVD†6‚’°¢6WBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BGG'VRrp¢&WGW&â¶÷&FW&VEÔ²ö²ÒGG'VS²F‚ÒFFW7C²&WW6VBÒGG'VRÐ¢Ð¢Ð ¢GF×F—"Ò¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’‚w7FFUÇF×Âr²„æWrÕ&$–B’¢æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚GF×F—"Ôf÷&6RÂ÷WBÔçVÆÀ¢GF×Ò¦ö–âÕF‚GF×F—"‚'6÷W&6RFW‡FVç6–öâ"¢G'’°¢F&Vf÷&RÒvWBÔ—FVÒÔÆ—FW&ÅF‚E6÷W&6UF‚ÔW'&÷$7F–öâ7F÷ ¢F6÷–VBÒFfÇ6S²FÆ7DW'&÷"Òrp¢f÷"‚F’Ò²F’ÖÆR3²F’²²’°¢G'’²6÷’Ôf–ÆU6†&VE&VBE6÷W&6UF‚GF×²F6÷–VBÒGG'VS²'&V²Ò6F6‚²FÆ7DW'&÷"ÒEòäW†6WF–öâäÖW76vRÐ¢–b‚F’ÖÇB2’²7F'BÕ6ÆVWÕ6V6öæG2"Ð¢Ð¢–b‚Öæ÷BF6÷–VB’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒv6÷’Öf–ÆVBs²ÖW76vRÒFÆ7DW'&÷"ÒÐ¢G'’²Væ&Æö6²Ôf–ÆRÔÆ—FW&ÅF‚GF×ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒ6F6‚²Ð¢F6÷”†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbGF×¢–b‚F6÷”†6‚ÖæRFW‡V7FVD†6‚’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒv†6‚ÖÖ—6ÖF6‚rÒÐ¢FgFW"ÒvWBÔ—FVÒÔÆ—FW&ÅF‚E6÷W&6UF‚ÔW'&÷$7F–öâ7F÷ ¢–b‚FgFW"äÆVæwF‚ÖæRF&Vf÷&RäÆVæwF‚Ö÷"FgFW"äÆ7Ew&—FUF–ÖUWF2ÖæRF&Vf÷&RäÆ7Ew&—FUF–ÖUWF2’°¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒw6÷W&6RÖ6†ævVBÖGW&–ærÖ6÷’rÐ¢Ð¢Ö÷fRÔ—FVÒÔÆ—FW&ÅF‚GF×ÔFW7F–æF–öâFFW7BÔf÷&6P¢Ff–æÄ†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbFFW7B¢–b‚Ff–æÄ†6‚ÖæRF6÷”†6‚’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FFW7BÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒwfW&–g’Öf–ÆVBrÐ¢Ð¢6WBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BGG'VRrp¢&WGW&â¶÷&FW&VEÔ²ö²ÒGG'VS²F‚ÒFFW7C²&WW6VBÒFfÇ6RÐ¢Ò6F6‚°¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvW'&÷"s²ÖW76vRÒEòäW†6WF–öâäÖW76vRÐ¢Òf–æÆÇ’°¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF×F—"’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF×F—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ð§Ð ¦gVæ7F–öâ6GW&RÔFWFV7FVE6æ6†÷B…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒD6GW&U&V6öâ’°¢2cRÜ*s2ãü*s2ã3¢jIÎyú^8ŽYÎi˜.8¾KùÞZÙŽ8ž8(¾8.™ÙžjÚ.[è^88(Ž8(®X˜Þ8 ¢–b‚Öæ÷B…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvæ÷BÖ&÷fVBrÒÐ¢GF‡2ÒvWBÕF‡0¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢Gv"Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWEv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Gv"ä6÷VçBÖW’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒwv÷&¶&öö²ÖÖ—76–ærrÒÐ¢GrÒGv%³Ð¢G6÷W&6UF‚Ò¦ö–âÕ6fR…·7G&–æuÒGF‡2ç7V&Ö—76–öäF—"’…·7G&–æuÒGrç&VÆF—fUF‚¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G6÷W&6UF‚’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvf–ÆRÖÖ—76–ærrÒÐ ¢FW‡V7FVE&Wf–÷W2Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÄ†6‚rrr’¢F—FVÒÒvWBÔ—FVÒÔÆ—FW&ÅF‚G6÷W&6UF‚ÔW'&÷$7F–öâ7F÷ ¢F6GW&VD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbG6÷W&6UF‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6GW&VD†6‚’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒv†6‚Öf–ÆVBrÒÐ¢F6GW&VEF–6·2Ò·7G&–æuÒF—FVÒäÆ7Ew&—FUF–ÖUWF2åF–6·0¢F6GW&VE6—¦RÒF—FVÒäÆVæwF€ ¢FÖWFÒVç7W&RÕ6æ6†÷DÖWFFFDÆæwVvREv÷&¶&öö´–B…·7G&–æuÒGrç&VÆF—fUF‚’…·7G&–æuÒGræ6FVv÷'’’G6÷W&6UF‚F6GW&VD†6‚D6GW&U&V6öà¢–b‚Öæ÷B¶&ööÅÒFÖWFæö²’²&WGW&âFÖWFÐ¢G6æ6†÷D–BÒ·7G&–æuÒFÖWFç6æ6†÷D–@¢–b…¶&ööÅÒ„vWBÔFF&÷W'G’FÖWFwVæF–ærrFfÇ6R’’°¢–b…FW7BÕ6÷W&6U&WFVçF–öäVæ&ÆVB’°¢G6fVBÒ6fRÕ6æ6†÷E6÷W&6Tf–ÆRDÆæwVvREv÷&¶&öö´–BG6æ6†÷D–BG6÷W&6UF‚F6GW&VD†6€¢–b‚Öæ÷B¶&ööÅÒG6fVBæö²’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒ·7G&–æuÒG6fVBç&V6öâÒÐ¢Ð¢–b‚Öæ÷B„6ö×ÆWFRÕ6æ6†÷BDÆæwVvREv÷&¶&öö´–BG6æ6†÷D–B’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒv6ö×ÆWFRÖf–ÆVBrÒÐ¢Ð ¢2cRÜ*s2ã3¢8+>89þ88>88Ž8ò6ö×&RÖæB×6WN8%28îi˜.X‹¾8~8þ8®8þhùX{®89^8*8*N8:¾8îZéþx«nhX¾8~XŠNZé®8ž8(¾8 ¢F6öÖÖ—GFVBÒWFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7B¢G‚Ò„vWBÔ'&’G7Bçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWEv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Öæ÷BG‚ä6÷VçB’²&WGW&â¶÷&FW&VEÔ²6öÖÖ—GFVBÒFfÇ6S²&V6öâÒwv÷&¶&öö²ÖÖ—76–ærrÒÐ¢GBÒG…³Ð¢FÆ—fRÒFçVÆÀ¢G'’²FÆ—fRÒvWBÔ—FVÒÔÆ—FW&ÅF‚G6÷W&6UF‚ÔW'&÷$7F–öâ7F÷Ò6F6‚²Ð¢G7F–ÆÄ7W'&VçBÒ‚FçVÆÂÖæRFÆ—fR’Öæ@¢…·7G&–æuÒFÆ—fRäÆ7Ew&—FUF–ÖUWF2åF–6·2ÖWF6GW&VEF–6·2’Öæ@¢‚FÆ—fRäÆVæwF‚ÖWF6GW&VE6—¦R¢–b‚Öæ÷BG7F–ÆÄ7W'&VçB’²&WGW&â¶÷&FW&VEÔ²6öÖÖ—GFVBÒFfÇ6S²&V6öâÒw7WW'6VFVBrÒÐ¢2cRÕ‚3’“¢F–6·2÷6—¦R8ÎYÎKˆ8~8(.Xh^Zëž8Îi»Ž8Þhù¾8(þ8>8n8N8(¾Xúþˆ;Þh
~8Î8.8(¾8þ8(8¢28+>89þ88>88Žy»NX˜Þ8¾Zéþ89^8*8*N8:¾8).XhÞ88þ88>8+~8:^8~8hÙ^hØži˜.8î88þ88>8+~8:^8ŽKˆˆ{N8ž8(¾8>8Ž8).z+®Š¨Þ8ž8(¾8 ¢FÆ—fT†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbG6÷W&6UF‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FÆ—fT†6‚’Ö÷"FÆ—fT†6‚ÖæRF6GW&VD†6‚’°¢&WGW&â¶÷&FW&VEÔ²6öÖÖ—GFVBÒFfÇ6S²&V6öâÒw7WW'6VFVBrÐ¢Ð¢FÆFW7BÒæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’GBv7W'&VçDW†6VÄ†6‚rrr’¢–b‚Öæ÷B…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FÆFW7B’’ÖæBFÆFW7BÖæRFW‡V7FVE&Wf–÷W2ÖæBFÆFW7BÖæRF6GW&VD†6‚’°¢&WGW&â¶÷&FW&VEÔ²6öÖÖ—GFVBÒFfÇ6S²&V6öâÒv6öæ7W'&VçB×WFFRrÐ¢Ð¢FBÔæ÷FU&÷W'G”–dÖ—76–ærGBv7W'&VçE6æ6†÷D–Brrp¢FBÔæ÷FU&÷W'G”–dÖ—76–ærGBvÆ7DFWFV7FVDBrrp¢6WBÔæ÷FU&÷W'G’GBv7W'&VçDW†6VÄ†6‚rF6GW&VD†6€¢6WBÔæ÷FU&÷W'G’GBv7W'&VçDW†6VÄÆ7Ew&—FUWF5F–6·2rF6GW&VEF–6·0¢6WBÔæ÷FU&÷W'G’GBv7W'&VçDW†6VÅ6—¦RrF6GW&VE6—¦P¢6WBÔæ÷FU&÷W'G’GBv7W'&VçDW†6VÄÖöF–f–VDBr‚FÆ—fRäÆ7Ew&—FUF–ÖRåFõ7G&–ær‚w———’ÔÔÒÖFED„ƒ¦ÖÓ§77§§¢r’¢6WBÔæ÷FU&÷W'G’GBv7W'&VçE6æ6†÷D–BrG6æ6†÷D–@¢6WBÔæ÷FU&÷W'G’GBvÆ7DFWFV7FVDBr„æWrÔæ÷t—6ò¢G&VæFW&VBÒæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’GBvÆ7E&VæFW&VDW†6VÄ†6‚rrr’¢–b‚G&VæFW&VBÖæBG&VæFW&VBÖæRF6GW&VD†6‚ÖæB·7G&–æuÒGBç7FGW2ÖæRw&VæFW"ÖW'&÷"r’²6WBÔæ÷FU&÷W'G’GBw7FGW2r„vWBÕ6÷W&6UWFFVE7FGW2GB’Ð¢&WGW&â¶÷&FW&VEÔ²6öÖÖ—GFVBÒGG'VRÐ¢Ð¢&WGW&â¶÷&FW&VEÔ²ö²ÒGG'VS²6æ6†÷D–BÒG6æ6†÷D–C²—4æWrÒ¶&ööÅÒFÖWFæ—4æWs²6öÖÖ—GFVBÒ¶&ööÅÒF6öÖÖ—GFVBæ6öÖÖ—GFVC²&V6öâÒ·7G&–æuÒF6öÖÖ—GFVBç&V6öâÐ§Ð ¦gVæ7F–öâWFFRÔ–çWD†—7F÷'”gFW%66â…·7G&–æuÒDÆæwVvR’°¢266âÕWFFW28îy»N[èÎ8¾YÎ8n8.xûîYÊŽ88þ88>8+~8:^8¾Zûî[ùÎ8ž8(¾jIÎyú^x˜Ž8ÎxJ8N89n88>8*þ888).KùÞZÙŽ8ž8(¾8 ¢–b‚Öæ÷B…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’’²&WGW&â‚’Ð¢G&W7VÇBÒ‚¢G'’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢f÷&V6‚‚Gr–â„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2’’°¢F–BÒ·7G&–æuÒGrçv÷&¶&öö´–@¢F7W"Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÄ†6‚rrr’¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W"’’²6öçF–çVRÐ¢F7W%6æÒ·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçE6æ6†÷D–Brrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W%6æ’’°¢FÒÒvWBÕ6æ6†÷DÖæ–fW7BDÆæwVvRF–BF7W%6æ ¢–b‚FçVÆÂÖæRFÒÖæB„æ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’FÒw6÷W&6T†6‚rrr’’’ÖWF7W"’²6öçF–çVRÐ¢Ð¢G"Ò6GW&RÔFWFV7FVE6æ6†÷BDÆæwVvRF–Bw66âp¢–b…¶&ööÅÒG"æö²’²G&W7VÇB³Ò¶÷&FW&VEÔ²v÷&¶&öö´–BÒF–C²6æ6†÷D–BÒ·7G&–æuÒG"ç6æ6†÷D–BÒÐ¢Ð¢Ò6F6‚²Ð¢&WGW&âG&W7VÇ@§Ð ¦gVæ7F–öâ6GW&RÕ&VæFW$–çWB…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒD¦ö$–B’°¢2cRÜ*t2Ó#¢jIÎyú^x˜Ž8Îiz.8¾8.8>8n8(.88:Î8;>888:®8;>8+K¨ŽZé®8®8(žXZ^X©¾89^8*8*N8:¾8þ[ø^8®z+®KùÞ8ž8(¾8 ¢GF‡2ÒvWBÕF‡0¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢Gv"Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWEv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Gv"ä6÷VçBÖW’²F‡&÷r%v÷&¶&öö¾8ÎŠh¾8N8¾8(®8î8¾8)3¢Ev÷&¶&öö´–B"Ð¢GrÒGv%³Ð¢FÆ—fUF‚Ò¦ö–âÕ6fR…·7G&–æuÒGF‡2ç7V&Ö—76–öäF—"’…·7G&–æuÒGrç&VÆF—fUF‚¢G6æÒE6æ6†÷D–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6æ’’²G6æÒ·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçE6æ6†÷D–Brrr’Ð ¢–b‚Öæ÷B…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6æ’’°¢2[^jÛNj™þˆ;Þ8ÎxJX«“¢[é>iÚ^8ž8®8(®hùX{®89^8*ž8:¾888îxûîxšž8).y»Nhê^KÛþ8n8 ¢&WGW&â¶÷&FW&VEÔ²F‚Òrs²6æ6†÷D–BÒrs²W†VÖW&ÂÒFfÇ6S²†6‚ÒrrÐ¢Ð¢FÒÒvWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BG6æ ¢–b‚FçVÆÂÖWFÒ’²&WGW&â¶÷&FW&VEÔ²F‚Òrs²6æ6†÷D–BÒrs²W†VÖW&ÂÒFfÇ6S²†6‚ÒrrÒÐ¢F†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’FÒw6÷W&6T†6‚rrr’ ¢–b…FW7BÕ6÷W&6U&WFVçF–öäVæ&ÆVB’°¢G7FFRÒvWBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvREv÷&¶&öö´–BG6æ ¢–b…¶&ööÅÒG7FFRç6÷W&6U&WF–æVB’°¢G&WF–æVD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#Sb…·7G&–æuÒG7FFRç6÷W&6UF‚’¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&WF–æVD†6‚’ÖæBG&WF–æVD†6‚ÖWF†6‚’°¢&WGW&â¶÷&FW&VEÔ²F‚Ò·7G&–æuÒG7FFRç6÷W&6UFƒ²6æ6†÷D–BÒG6æ²W†VÖW&ÂÒFfÇ6S²†6‚ÒF†6ƒ²fW&–f–VBÒGG'VRÐ¢Ð¢2æWfW"&VæFW"6÷''WFVB&WF–æVB6÷W&6R2F†R&WVW7FVB–Ö×WF&ÆP¢2vVæW&F–öââÖ&²—BVæf–Æ&ÆRÂF†VâGFV×B&V6÷fW'’g&öÒF†RÆ—fRf–ÆRà¢G'’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚…·7G&–æuÒG7FFRç6÷W&6UF‚’Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒ6F6‚²Ð¢6WBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvREv÷&¶&öö´–BG6æFfÇ6Rv†6‚ÖÖ—6ÖF6‚p¢Ð¢2xûîxšž8ÎkhŽ8Ž8n8N8(³¢hùX{®89^8*ž8:¾888îxûîxšž8ÎYÎ8Ž88þ88>8+~8:^8®8(ž[êžXX>8ž8(¾8 ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FÆ—fUF‚’°¢FÆ—fT†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbFÆ—fUF‚¢–b‚FÆ—fT†6‚ÖWF†6‚’°¢G"Ò6fRÕ6æ6†÷E6÷W&6Tf–ÆRDÆæwVvREv÷&¶&öö´–BG6æFÆ—fUF‚F†6€¢–b…¶&ööÅÒG"æö²’²&WGW&â¶÷&FW&VEÔ²F‚Ò·7G&–æuÒG"çFƒ²6æ6†÷D–BÒG6æ²W†VÖW&ÂÒFfÇ6S²†6‚ÒF†6ƒ²fW&–f–VBÒGG'VRÒÐ¢Ð¢Ð¢Ð ¢2{Šî˜8:.8;Î88ž88î8þ8þxûîxšž8).[êžXX>8~8Þ8®8NZNYŽ8þKˆi˜.8+>89N8;Î8).KÙÎ8(¾8 ¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FÆ—fUF‚’’²&WGW&â¶÷&FW&VEÔ²F‚Òrs²6æ6†÷D–BÒG6æ²W†VÖW&ÂÒFfÇ6S²†6‚ÒF†6‚ÒÐ¢F6GW&T–BÒB†–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D¦ö$–B’’²æWrÕ&$–BÒVÇ6R²D¦ö$–BÒ¢FF—"Ò¦ö–âÕF‚„vWBÔW†VÖW&Ä¦ö%&ö÷BDÆæwVvR’F6GW&T–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢FW‡FVç6–öâÒ´”òåF…Ó£¤vWDW‡FVç6–öâ‚FÆ—fUF‚’åFôÆ÷vW$–çf&–çB‚¢–b‚FW‡FVç6–öâÖæ÷F–â‚rç†Ç7‚rÂrç†Ç6ÒrÂræFö7‚rÂrçG‚rÂrçFbr’’²FW‡FVç6–öâÒrç†Ç7‚rÐ¢GF×Ò¦ö–âÕF‚FF—"‚'6÷W&6RFW‡FVç6–öâ"¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚GF×’’°¢6÷’Ôf–ÆU6†&VE&VBFÆ—fUF‚GF× ¢G'’²Væ&Æö6²Ôf–ÆRÔÆ—FW&ÅF‚GF×ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒ6F6‚²Ð¢Ð¢GF×†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbGF×¢–b‚GF×†6‚ÖæRF†6‚’°¢2hùX{®89^8*ž8:¾888îxûîxšž8þiz.8¾XŠ^8îx˜Ž8¾8®8>8n8N8(¾8 ¢28>8îjIÎyú^x˜Ž8îXZ^X©¾8Ž8~8n8þKÛþ8Ž8®8B„ƒ8åDn8Ž8~8dƒ.8).{XN8)>8~8~8î8n8þ8(ž8 ¢&VÖ÷fRÔW†VÖW&Ä6÷’DÆæwVvRF6GW&T–@¢&WGW&â¶÷&FW&VEÔ²F‚Òrs²6æ6†÷D–BÒrs²W†VÖW&ÂÒFfÇ6S²†6‚ÒrrÐ¢Ð¢&WGW&â¶÷&FW&VEÔ²F‚ÒGF×²6æ6†÷D–BÒG6æ²W†VÖW&ÂÒGG'VS²†6‚ÒGF×†6ƒ²6GW&T–BÒF6GW&T–C²fW&–f–VBÒGG'VRÐ§Ð ¦gVæ7F–öâ&VÖ÷fRÔW†VÖW&Ä6÷’…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒD6GW&T–B’°¢G'’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D6GW&T–B’’²&WGW&âÐ¢FF—"Ò¦ö–âÕF‚„vWBÔW†VÖW&Ä¦ö%&ö÷BDÆæwVvR’D6GW&T–@¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FF—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâ6ÆV"Õ7FÆTW†VÖW&Ä6÷–W2…·7G&–æuÒDÆæwVvR’°¢2cRÜ*sã3¢™ÙîKùÞhÈh›þŠ¨Þ8®8î8¾xûîxšž8Î™[~iÉþ™i>jè¾8(ž8®8N8(Ž8n8Kˆ®™™i˜.™i>8~[ø^8®khŽ8ž8 ¢G'’°¢G&ö÷BÒvWBÔW†VÖW&Ä¦ö%&ö÷BDÆæwVvP¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G&ö÷B’’²&WGW&âÐ¢FÖ„vRÒ¶–çEÒ„vWBÔ–çWD†—7F÷'•6WGF–æw2’æW†VÖW&Ä6÷”Ö„vTÖ–çWFW0¢–b‚FÖ„vRÖÆR’²FÖ„vRÒ3Ð¢FÆ–Ö—BÒ´FFUF–ÖUÓ£¥WF4æ÷räFDÖ–çWFW2‚Ó¢FÖ„vR¢f÷&V6‚‚FB–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚G&ö÷BÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢–b‚FBäÆ7Ew&—FUF–ÖUWF2ÖÇBFÆ–Ö—B’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FBägVÆÄæÖRÕ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRv†—7F÷'’æ6ÆVçWr…¶÷&FW&VEÔ²¶–æBÒvW†VÖW&ÂÖ6÷’s²6GW&T–BÒ·7G&–æuÒFBäæÖRÒ¢Ð¢Ð¢Ò6F6‚²Ð§Ð  ¢2ÒÒÒÒhè>™šB…cRÜ*s2ãbò*s2ãr’ÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÐ ¦gVæ7F–öâvWBÕ&÷FV7FVE6æ6†÷D–G2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢2cRÜ*t2Ó3¢KùÞhÈi[8îiê8Ž8þxJ™j.Kø.8¾[‹Ži˜.KùÞŠÛ~8ž8(¾jIÎyú^x˜Ž8 ¢2ƒÓâƒ"ÓâXhÔƒ8~8òÖæ–fW7B8âFWFV7FVDB8ÎXúN8N8î8î8®8î8~8iz^i˜.šn8î8Îy»N‹ùîx˜Ž8Þ8~8þxûîYÊŽx˜Ž8).khŽ8~8n8(¾8 ¢F–G2Ò‚¢G'’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢Gv"Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWEv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Gv"ä6÷VçBÖwB’°¢F–G2³Ò·7G&–æuÒ„vWBÔFF&÷W'G’Gv%³Òv7W'&VçE6æ6†÷D–Brrr¢F–G2³Ò·7G&–æuÒ„vWBÔFF&÷W'G’Gv%³ÒvÆ7E&VæFW&VE6æ6†÷D–Brrr¢Ð¢Ò6F6‚²Ð¢G'’°¢FWFòÒ&VBÔWFõ7FFRDÆæwVvREv÷&¶&öö´–@¢–b‚FçVÆÂÖæRFWFò’²F–G2³Ò·7G&–æuÒ„vWBÔFF&÷W'G’FWFòwVæF–æu6æ6†÷D–Brrr’Ð¢Ò6F6‚²Ð¢G'’°¢GG"ÒvWBÔ6ö×&—6öä&6VÆ–æUö–çFW"DÆæwVvREv÷&¶&öö´–@¢–b‚FçVÆÂÖæRGG"’²F–G2³Ò·7G&–æuÒ„vWBÔFF&÷W'G’GG"w6æ6†÷D–Brrr’Ð¢Ò6F6‚²Ð¢&WGW&â‚F–G2Âv†W&RÔö&¦V7B²Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Eò’ÒÂ6VÆV7BÔö&¦V7BÕVæ—VR§Ð ¦gVæ7F–öâvWBÔf–æÅFe–ävTF—2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢2jÚ>[ÈõDn8~KÛþ8(þ8(Î8þx˜Ž8î8n88iÈ8(.ik8~8NX{®X©¾8¾8(ž8î{XÎ˜îiz^i[8).‹ùN8ž8'–â8ÎxJ88(Î8Ó8 ¢FF—"ÒvWBÕ6æ6†÷E–äF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&âÓÐ¢FæWvW7BÒFçVÆÀ¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"vf–æÂ×Feò¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢–b‚FçVÆÂÖWFæWvW7BÖ÷"FbäÆ7Ew&—FUF–ÖUWF2ÖwBFæWvW7B’²FæWvW7BÒFbäÆ7Ew&—FUF–ÖUWF2Ð¢Ð¢–b‚FçVÆÂÖWFæWvW7B’²&WGW&âÓÐ¢&WGW&â…´FFUF–ÖUÓ£¥WF4æ÷rÒFæWvW7B’åF÷FÄF—0§Ð ¦gVæ7F–öâ–çfö¶RÔ–çWD†—7F÷'”6ÆVçW…·7G&–æuÒDÆæwVvR’°¢–b‚Öæ÷B…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’’²&WGW&âÐ¢FÆö6µF‚Ò¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’vÆö6·5Æ†—7F÷'’Ö6ÆVçWæÆö6²p¢F†æFÆRÒG'’Ô7V—&TÆö6´†æFÆRFÆö6µF€¢–b‚FçVÆÂÖWF†æFÆR’²&WGW&âÒ2K¹n8+^8;Î898;Î8Îhè>™šNKŠÐ¢G'’°¢F6frÒvWBÔ–çWD†—7F÷'•6WGF–æw0¢G&ö÷BÒvWBÔ–çWD†—7F÷'•&ö÷BDÆæwVvP¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G&ö÷B’’²&WGW&âÐ¢f÷&V6‚‚Gv$F—"–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚G&ö÷BÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢Gv÷&¶&öö´–BÒ·7G&–æuÒGv$F—"äæÖP¢G&÷FV7FVBÒ„vWBÕ&÷FV7FVE6æ6†÷D–G2DÆæwVvRGv÷&¶&öö´–B¢2iÊ®ZèÎh‰K‰nKº2†Öæ–fW7B8®8rž8).X˜®™š@¢G&VÖ÷fVD–æ6ö×ÆWFRÒFfÇ6P¢f÷&V6‚‚FB–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚Gv$F—"ägVÆÄæÖRÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚„¦ö–âÕF‚FBägVÆÄæÖRvÖæ–fW7Bæ§6öâr’’’°¢–b‚FBäÆ7Ew&—FUF–ÖUWF2ÖÇB´FFUF–ÖUÓ£¥WF4æ÷räFD†÷W'2‚Ób’’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FBägVÆÄæÖRÕ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢G&VÖ÷fVD–æ6ö×ÆWFRÒGG'VP¢Ð¢Ð¢Ð¢–b‚G&VÖ÷fVD–æ6ö×ÆWFR’²6ÆV"Õ6æ6†÷E'VçF–ÖT66†W2DÆæwVvRGv÷&¶&öö´–BÐ¢FÆÂÒ„vWBÕ6æ6†÷D–G2DÆæwVvRGv÷&¶&öö´–B¢–b‚FÆÂä6÷VçBÖW’²6öçF–çVRÐ¢F¶VW&V6VçBÒ‚FÆÂÂ6VÆV7BÔö&¦V7BÔÆ7B…´ÖF…Ó£¤Ö‚ƒÂ¶–çEÒF6frç&WF–å6÷W&6UfW'6–öç2’’¢f÷&V6‚‚G6â–âFÆÂ’°¢–b‚G&÷FV7FVBÖ6öçF–ç2G6â’²6öçF–çVRÐ¢–b‚F¶VW&V6VçBÖ6öçF–ç2G6â’²6öçF–çVRÐ¢–b‚„vWBÔ7F—fTÆV6T6÷VçBDÆæwVvRGv÷&¶&öö´–BG6â’ÖwB’²6öçF–çVRÐ¢G–ç2Ò„vWBÕ6æ6†÷E–ç2DÆæwVvRGv÷&¶&öö´–BG6â¢–b‚G–ç2Ö6öçF–ç2v6ö×&—6öâÖ&6VÆ–æRr’²6öçF–çVRÐ¢–b‚G–ç2Ö6öçF–ç2vÖçVÂr’²6öçF–çVRÐ¢Ff–æÅ–ç2Ò‚G–ç2Âv†W&RÔö&¦V7B²EòÖÆ–¶Rvf–æÂ×Feò¢rÒ ¢2ÒÒÒjë^™¨ã¢6÷W&6Rç†Ç7‚888).X˜®™šN8ž8(²ÒÒÐ¢G7FFRÒvWBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvRGv÷&¶&öö´–BG6à¢–b…¶&ööÅÒG7FFRç6÷W&6U&WF–æVB’°¢F6å&VÖ÷fU6÷W&6RÒGG'VP¢–b‚Ff–æÅ–ç2ä6÷VçBÖwB’°¢FF—2Ò¶F÷V&ÆUÒ„vWBÔ–çWD†—7F÷'•6WGF–æw2’ç6÷W&6U&WFVçF–öäF—4gFW$'V–Æ@¢F6öæf–wW&VBÒ„vWBÔ–çWD†—7F÷'•6WGF–æw2’ç6÷W&6U&WFVçF–öäF—4gFW$'V–Æ@¢–b‚FçVÆÂÖWF6öæf–wW&VB’²F6å&VÖ÷fU6÷W&6RÒFfÇ6RÐ¢VÇ6R°¢FvRÒvWBÔf–æÅFe–ävTF—2DÆæwVvRGv÷&¶&öö´–BG6à¢–b‚FvRÖÇB¶F÷V&ÆUÒF6öæf–wW&VB’²F6å&VÖ÷fU6÷W&6RÒFfÇ6RÐ¢Ð¢Ð¢–b‚F6å&VÖ÷fU6÷W&6R’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚…·7G&–æuÒG7FFRç6÷W&6UF‚’Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢6WBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvRGv÷&¶&öö´–BG6âFfÇ6Rw&WFVçF–öâp¢6ÆV"Õ6æ6†÷E'VçF–ÖT66†W2DÆæwVvRGv÷&¶&öö´–BG6à¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRv–çWBç6÷W&6Rç&VÖ÷fVBr…¶÷&FW&VEÔ²v÷&¶&öö´–BÒGv÷&¶&öö´–C²6æ6†÷D–BÒG6ã²&V6öâÒw&WFVçF–öârÒ¢Ð¢Ð ¢2ÒÒÒjë^™¨ã#¢–â8ÃK»n8(.xJ88(Î889^8*ž8:¾888N8ŽX˜®™šN8ž8(²ÒÒÐ¢–b‚Ff–æÅ–ç2ä6÷VçBÖWÖæBG–ç2ä6÷VçBÖW’°¢FÒÒvWBÕ6æ6†÷DÖæ–fW7BDÆæwVvRGv÷&¶&öö´–BG6à¢–b‚FçVÆÂÖæRFÒÖæB·7G&–æuÒ„vWBÔFF&÷W'G’FÒw7FGW2rrr’ÖWv6ö×ÆWFRr’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚„vWBÕ6æ6†÷DF—"DÆæwVvRGv÷&¶&öö´–BG6â’Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢6ÆV"Õ6æ6†÷E'VçF–ÖT66†W2DÆæwVvRGv÷&¶&öö´–BG6à¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRv†—7F÷'’æ6ÆVçWr…¶÷&FW&VEÔ²¶–æBÒw6æ6†÷Bs²v÷&¶&öö´–BÒGv÷&¶&öö´–C²6æ6†÷D–BÒG6âÒ¢Ð¢Ð¢Ð¢Ð¢Ò6F6‚°¢2hè>™šN8îZKiY~8þ8:Î8;>888:®8;>8+8îZKiY~8¾8~8®8N8 ¢w&—FRÕv&æ–ær‚.[^jÛN8îhè>™šN8¾ZKiY~8~8î8~8ó¢"²EòäW†6WF–öâäÖW76vR¢Òf–æÆÇ’°¢&VÆV6RÔÆö6´†æFÆRF†æFÆP¢2cRÕ#¢hè>™šN8~Zëž˜xþ8ÎZHž8(þ8(¾8þ8(8jÊY¹î8âö’÷7FFR8~ZéþkŠÎ8^8¾8(¾8 ¢&W6WBÔ–çWD†—7F÷'•6—¦T66†P¢Ð§Ð ¦gVæ7F–öâvWBÔ–çWD†—7F÷'•6—¦TÖ"…·7G&–æuÒDÆæwVvR’°¢2–çWBÖ†—7F÷'’˜XÞKˆ¾8îXZŽXhÞ[‹X‰~hÉž8þX[iÈž88ž8:ž8*N89nKˆ®8~™Ùî[‹Ž8¾˜xÞ8N8 ¢2ö’÷7FFR8î89Þ8;Î8:®8;>8+8N8Ž8¾ZéþŠÎ8¾8£czy.8*Þ8:>88>8+~8:^8ž8(¾8 ¢2FFF—"8).Xˆ~8(®i»þ8Ž8þy»N[èÎ8¾iz~8:þ8;Î8*þ8+ž89®8;Î8+ž8îX
N8).‹ùN8^8®8N8(Ž8n88:¾8;Î88Ž898+ž8(.8*Þ8;Î8¾Y
¾8(8(¾8 ¢G&ö÷BÒrp¢G'’²G&ö÷BÒvWBÔ–çWD†—7F÷'•&ö÷BDÆæwVvRÒ6F6‚²Ð¢F66†T¶W’Ò‚w³×Ç³ÒrÖbDÆæwVvRÂ·7G&–æuÒG&ö÷B’åFôÆ÷vW$–çf&–çB‚¢–b‚FçVÆÂÖæRE67&—C¤†—7F÷'•6—¦T66†RÖæBE67&—C¤†—7F÷'•6—¦T66†T¶W’ÖWF66†T¶W’ÖæB‚…´FFUF–ÖUÓ£¥WF4æ÷rÒE67&—C¤†—7F÷'•6—¦T66†TEWF2’åF÷FÅ6V6öæG2ÖÇBE67&—C¤†—7F÷'•6—¦T66†U6V6öæG2’’°¢&WGW&âE67&—C¤†—7F÷'•6—¦T66†P¢Ð¢GfÇVRÒ ¢G'’°¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&ö÷B’ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚G&ö÷B’’°¢F'—FW2Ò„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚G&ö÷BÕ&V7W'6RÔf–ÆRÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂÖV7W&RÔö&¦V7BÕ&÷W'G’ÆVæwF‚Õ7VÒ’å7VÐ¢–b‚FçVÆÂÖæRF'—FW2’²GfÇVRÒ´ÖF…Ó£¥&÷VæB‚‚F'—FW2òÔ"’Â’Ð¢Ð¢Ò6F6‚²GfÇVRÒÐ¢E67&—C¤†—7F÷'•6—¦T66†RÒGfÇVP¢E67&—C¤†—7F÷'•6—¦T66†T¶W’ÒF66†T¶W¢E67&—C¤†—7F÷'•6—¦T66†TEWF2Ò´FFUF–ÖUÓ£¥WF4æ÷p¢&WGW&âGfÇVP§Ð ¦gVæ7F–öâ&W6WBÔ–çWD†—7F÷'•6—¦T66†R°¢E67&—C¤†—7F÷'•6—¦T66†RÒFçVÆÀ¢E67&—C¤†—7F÷'•6—¦T66†T¶W’Òrp¢E67&—C¤†—7F÷'•6—¦T66†TEWF2Ò´FFUF–ÖUÓ£¤Ö–åfÇVP§Ð ¢2ÒÒÒÒ6öçFVçB×Fb8îKùÞŠÛr…cRÜ*s2ãr’ÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÐ ¦gVæ7F–öâvWBÔ6öçFVçEFefW'6–öäF—"…·7G&–æuÒEv÷&·76RÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G6fUfW'6–öä–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEfW'6–öä–BwfW'6–öä–Bp¢&WGW&â„¦ö–âÕF‚Ev÷&·76R„¦ö–âÕF‚v6öçFVçB×Fbr„¦ö–âÕF‚G6fUv÷&¶&öö´–BG6fUfW'6–öä–B’’§Ð ¦gVæ7F–öâæWrÔ6öçFVçEFe–â…·7G&–æuÒEv÷&·76RÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒE–äæÖRÂDFF’°¢2cRÕ‚3B“¢ZKiY~8).hú8(®8N8n8^8¢GG'VRòFfÇ6R8~‹ùN8ž8.YÎX{®XX>8Î{YiéÎ8).jIÎiû¾8~8n8:Þ8;Î8:¾8988>8*þ8~8Þ8(¾8(Ž8n8¾8ž8(¾8 ¢G'’°¢G6fU–äæÖRÒ76W'BÕ6fU7F÷&vU6VvÖVçBE–äæÖRw–äæÖRp¢GfW'6–öäF—"ÒvWBÔ6öçFVçEFefW'6–öäF—"Ev÷&·76REv÷&¶&öö´–BEfW'6–öä–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚GfW'6–öäF—"ÕF…G—R6öçF–æW"’’²&WGW&âFfÇ6RÐ¢FF—"Ò¦ö–âÕF‚GfW'6–öäF—"w–ç2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢G–åF‚Ò¦ö–âÕF‚FF—"‚'³Òæ§6öâ"ÖbG6fU–äæÖR¢w&—FRÔ§6öäf–ÆRG–åF‚DFF¢&WGW&â…FW7BÕF‚ÔÆ—FW&ÅF‚G–åF‚¢Ò6F6‚²&WGW&âFfÇ6RÐ§Ð ¦gVæ7F–öâ&VÖ÷fRÔ6öçFVçEFe–â…·7G&–æuÒEv÷&·76RÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒE–äæÖR’°¢G'’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EfW'6–öä–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E–äæÖR’’²&WGW&âÐ¢GF‚Ò¦ö–âÕF‚„¦ö–âÕF‚„vWBÔ6öçFVçEFefW'6–öäF—"Ev÷&·76REv÷&¶&öö´–BEfW'6–öä–B’w–ç2r’‚'³Òæ§6öâ"ÖbE–äæÖR¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ò6F6‚²Ð§Ð  ¦gVæ7F–öâæWrÔ6öçFVçEFdÆV6R…·7G&–æuÒEv÷&·76RÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒEW'÷6RÂ·7G&–æuÒD¦ö$–BÂ¶–çEÒDÖ–çWFW5fÆ–BÒ#’°¢G'’°¢G6fUW'÷6RÒ76W'BÕ6fU7F÷&vU6VvÖVçBEW'÷6RwW'÷6Rp¢G6fT¦ö$–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBD¦ö$–Bv¦ö$–Bp¢GfW'6–öäF—"ÒvWBÔ6öçFVçEFefW'6–öäF—"Ev÷&·76REv÷&¶&öö´–BEfW'6–öä–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚GfW'6–öäF—"ÕF…G—R6öçF–æW"’’²&WGW&ârrÐ¢FF—"Ò¦ö–âÕF‚GfW'6–öäF—"vÆV6W2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢FæÖRÒ‚w³Õ÷³Òæ§6öârÖbG6fUW'÷6RÂG6fT¦ö$–B¢GF‚Ò¦ö–âÕF‚FF—"FæÖP¢w&—FRÔ§6öäf–ÆRGF‚…¶÷&FW&VEÔ°¢¦ö$–BÒD¦ö$–C²W'÷6RÒEW'÷6S²fW'6–öä–BÒEfW'6–öä–@¢7&VFVDBÒæWrÔæ÷t—6ó²†V'F&VDBÒæWrÔæ÷t—6ð¢W‡—&W4BÒ…´FFUF–ÖUÓ£¥WF4æ÷räFDÖ–çWFW2‚DÖ–çWFW5fÆ–B’åFõ7G&–ær‚vòr’¢4æÖRÒFVçc¤4ôÕUDU$äÔP¢Ò¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’’²&WGW&ârrÐ¢&WGW&âFæÖP¢Ò6F6‚²&WGW&ârrÐ§Ð ¦gVæ7F–öâ&VÖ÷fRÔ6öçFVçEFdÆV6R…·7G&–æuÒEv÷&·76RÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒDÆV6TæÖR’°¢G'’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚DÆV6TæÖR’’²&WGW&âÐ¢GF‚Ò¦ö–âÕF‚„¦ö–âÕF‚„vWBÔ6öçFVçEFefW'6–öäF—"Ev÷&·76REv÷&¶&öö´–BEfW'6–öä–B’vÆV6W2r’DÆV6TæÖP¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâ&Vg&W6‚ÔÆV6Tf–ÆR…·7G&–æuÒEF‚Â¶–çEÒDÖ–çWFW5fÆ–BÒ#’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EF‚’Ö÷"Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚EF‚’’²&WGW&âFfÇ6RÐ¢G'’°¢FÆV6RÒ&VBÔ§6öäf–ÆREF‚FçVÆÀ¢–b‚FçVÆÂÖWFÆV6R’²&WGW&âFfÇ6RÐ¢6WBÔæ÷FU&÷W'G’FÆV6Rv†V'F&VDBr„æWrÔæ÷t—6ò¢6WBÔæ÷FU&÷W'G’FÆV6RvW‡—&W4Br…´FFUF–ÖUÓ£¥WF4æ÷räFDÖ–çWFW2‚DÖ–çWFW5fÆ–B’åFõ7G&–ær‚vòr’¢w&—FRÔ§6öäf–ÆREF‚FÆV6P¢&WGW&âGG'VP¢Ò6F6‚²&WGW&âFfÇ6RÐ§Ð ¦gVæ7F–öâæWrÔF–fd¦ö$ÆV6W2…·7G&–æuÒDÆæwVvRÂD6öçFW‡BÂ·7G&–æuÒD¦ö$–B’°¢Gv÷&·76RÒvWBÕv÷&·76UF‚DÆæwVvP¢F†—7F÷'”Æö6²Ò¦ö–âÕF‚Gv÷&·76RvÆö6·5Æ†—7F÷'’Ö6ÆVçWæÆö6²p¢F6öçFVçDÆö6²ÒvWBÔ6öçFVçEFdÖ–çFVææ6TÆö6µF‚Gv÷&·76R…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B¢&WGW&â–çfö¶RÕv—F„Æö6²F†—7F÷'”Æö6²°¢–çfö¶RÕv—F„Æö6²F6öçFVçDÆö6²°¢FÆV6W2Ò‚¢G'’°¢f÷&V6‚‚G6–FR–â€¢¶÷&FW&VEÔ²&öÆRÒv&6VÆ–æRs²6æ6†÷D–BÒ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æU6æ6†÷D–C²fW'6–öä–BÒ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æUfW'6–öä–BÒÀ¢¶÷&FW&VEÔ²&öÆRÒv7W'&VçBs²6æ6†÷D–BÒ·7G&–æuÒD6öçFW‡Bæ7W'&VçE6æ6†÷D–C²fW'6–öä–BÒ·7G&–æuÒD6öçFW‡Bæ7W'&VçEfW'6–öä–BÐ¢’’°¢–b‚FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvR…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒG6–FRç6æ6†÷D–B’’’°¢F‡&÷r~jùN‹È>Zûî‹8î[^jÛNx˜Ž8Îi[Nyn8^8(Î8þ8þ8(8Xznyn8).™h¾Zx¾8~8Þ8î8¾8)>8"p¢Ð¢Ff–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvR…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒG6–FRç6æ6†÷D–B’…·7G&–æuÒG6–FRçfW'6–öä–B¢–b‚Öæ÷B¶&ööÅÒFf–Æ&–Æ—G’ç&VG’’²F‡&÷r·7G&–æuÒFf–Æ&–Æ—G’ç&V6öâÐ¢G6æ6†÷DÆV6RÒæWrÕ6æ6†÷DÆV6RDÆæwVvR…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒG6–FRç6æ6†÷D–B’‚vF–fbÒr²·7G&–æuÒG6–FRç&öÆR’D¦ö$–B#…·7G&–æuÒG6–FRçfW'6–öä–B¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6æ6†÷DÆV6R’’²F‡&÷r~[^jÛNx˜Ž8îKùÞŠÛvÆV6^8).KÙÎh‰8~8Þ8î8¾8)>8~8~8þ8"rÐ¢FÆV6RÒ¶÷&FW&VEÔ°¢&öÆRÒ·7G&–æuÒG6–FRç&öÆP¢6æ6†÷D–BÒ·7G&–æuÒG6–FRç6æ6†÷D–@¢fW'6–öä–BÒ·7G&–æuÒG6–FRçfW'6–öä–@¢6æ6†÷DÆV6TæÖRÒG6æ6†÷DÆV6P¢6öçFVçDÆV6TæÖRÒrp¢Ð¢FÆV6W2³ÒFÆV6P¢F6öçFVçDÆV6RÒæWrÔ6öçFVçEFdÆV6RGv÷&·76R…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒG6–FRçfW'6–öä–B’‚vF–fbÒr²·7G&–æuÒG6–FRç&öÆR’D¦ö$–B# ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6öçFVçDÆV6R’’²F‡&÷rv6öçFVçBDn8îKùÞŠÛvÆV6^8).KÙÎh‰8~8Þ8î8¾8)>8~8~8þ8"rÐ¢FÆV6Ræ6öçFVçDÆV6TæÖRÒF6öçFVçDÆV6P¢Ð¢&WGW&â‚FÆV6W2¢Ò6F6‚°¢f÷&V6‚‚FÆV6R–â‚FÆV6W2’’°¢&VÖ÷fRÕ6æ6†÷DÆV6RDÆæwVvR…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒFÆV6Rç6æ6†÷D–B’…·7G&–æuÒFÆV6Rç6æ6†÷DÆV6TæÖR¢&VÖ÷fRÔ6öçFVçEFdÆV6RGv÷&·76R…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒFÆV6RçfW'6–öä–B’…·7G&–æuÒFÆV6Ræ6öçFVçDÆV6TæÖR¢Ð¢F‡&÷p¢Ð¢Ð¢Ð§Ð ¦gVæ7F–öâ&Vg&W6‚ÔF–fd¦ö$ÆV6W2‚D¦ö"’°¢FÆæwVvRÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"vÖöFRrDÖöFR¢Gv÷&¶&öö´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"wv÷&¶&öö´–Brrr¢Gv÷&·76RÒvWBÕv÷&·76UF‚FÆæwVvP¢f÷&V6‚‚FÆV6R–â„vWBÔ'&’„vWBÔFF&÷W'G’D¦ö"vÆV6W2r‚’’’’°¢G6æ6†÷EF‚Ò¦ö–âÕF‚„vWBÕ6æ6†÷DÆV6TF—"FÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒFÆV6Rç6æ6†÷D–B’’…·7G&–æuÒFÆV6Rç6æ6†÷DÆV6TæÖR¢F6öçFVçEF‚Ò¦ö–âÕF‚„¦ö–âÕF‚„vWBÔ6öçFVçEFefW'6–öäF—"Gv÷&·76RGv÷&¶&öö´–B…·7G&–æuÒFÆV6RçfW'6–öä–B’’vÆV6W2r’…·7G&–æuÒFÆV6Ræ6öçFVçDÆV6TæÖR¢–b‚Öæ÷B…&Vg&W6‚ÔÆV6Tf–ÆRG6æ6†÷EF‚#’’²F‡&÷r~[^jÛNx˜Ž8îKùÞŠÛvÆV6^8).i»Nik8~8Þ8î8¾8)>8~8~8þ8"rÐ¢–b‚Öæ÷B…&Vg&W6‚ÔÆV6Tf–ÆRF6öçFVçEF‚#’’²F‡&÷rv6öçFVçBDn8îKùÞŠÛvÆV6^8).i»Nik8~8Þ8î8¾8)>8~8~8þ8"rÐ¢Ð§Ð ¦gVæ7F–öâ&VÖ÷fRÔF–fd¦ö$ÆV6W2‚D¦ö"’°¢G'’°¢FÆæwVvRÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"vÖöFRrDÖöFR¢Gv÷&¶&öö´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"wv÷&¶&öö´–Brrr¢Gv÷&·76RÒvWBÕv÷&·76UF‚FÆæwVvP¢f÷&V6‚‚FÆV6R–â„vWBÔ'&’„vWBÔFF&÷W'G’D¦ö"vÆV6W2r‚’’’’°¢&VÖ÷fRÕ6æ6†÷DÆV6RFÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒFÆV6Rç6æ6†÷D–B’…·7G&–æuÒFÆV6Rç6æ6†÷DÆV6TæÖR¢&VÖ÷fRÔ6öçFVçEFdÆV6RGv÷&·76RGv÷&¶&öö´–B…·7G&–æuÒFÆV6RçfW'6–öä–B’…·7G&–æuÒFÆV6Ræ6öçFVçDÆV6TæÖR¢Ð¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâFW7BÔ6öçFVçEFe&÷FV7FVB…·7G&–æuÒEv÷&·76RÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–B’°¢F&6RÒvWBÔ6öçFVçEFefW'6–öäF—"Ev÷&·76REv÷&¶&öö´–BEfW'6–öä–@¢f÷&V6‚‚G7V"–â‚w–ç2rÂvÆV6W2r’’°¢FBÒ¦ö–âÕF‚F&6RG7V ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FB’°¢Ff–ÆW2Ò„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FBÔf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR¢–b‚G7V"ÖWw–ç2rÖæBFf–ÆW2ä6÷VçBÖwB’²&WGW&âGG'VRÐ¢–b‚G7V"ÖWvÆV6W2r’²f÷&V6‚‚Fb–âFf–ÆW2’²–b…FW7BÔÆV6T7F—fRFb’²&WGW&âGG'VRÒÒÐ¢Ð¢Ð¢&WGW&âFfÇ6P§Ð  ¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ¢2cR7FvR2(	B†6R$¢jIÎyú^898*N89~8:ž8*N8;>8Žˆz®X¹^8+ž8+8+Ž8:^8;Î8:ž8;À¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ ¦gVæ7F–öâvWBÔWFõ7FFTF—"…·7G&–æuÒDÆæwVvR’°¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’w7FFUÆWFò×&VæFW"r§Ð¦gVæ7F–öâ&VBÔWFõ7FFR…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢G'’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢GÒ¦ö–âÕF‚„vWBÔWFõ7FFTF—"DÆæwVvR’‚'³Òæ§6öâ"ÖbG6fUv÷&¶&öö´–B¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G’’²&WGW&âFçVÆÂÐ¢&WGW&â…&VBÔ§6öäf–ÆRGFçVÆÂ¢Ò6F6‚²&WGW&âFçVÆÂÐ§Ð¦gVæ7F–öâw&—FRÔWFõ7FFR…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂE7FFR’°¢G'’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢FF—"ÒvWBÔWFõ7FFTF—"DÆæwVvP¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢6WBÔæ÷FU&÷W'G’E7FFRwv÷&¶&öö´–BrG6fUv÷&¶&öö´–@¢6WBÔæ÷FU&÷W'G’E7FFRwWFFVDBr„æWrÔæ÷t—6ò¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚FF—"‚'³Òæ§6öâ"ÖbG6fUv÷&¶&öö´–B’’E7FFP¢Ò6F6‚²Ð§Ð¦gVæ7F–öâæWrÔWFõ7FFR…·7G&–æuÒEv÷&¶&öö´–B’°¢&WGW&â¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ#²v÷&¶&öö´–BÒEv÷&¶&öö´–C²WFFVDBÒæWrÔæ÷t—6ð¢VæF–æu6æ6†÷D–BÒrs²VæF–æt†6‚Òrs²7F&ÆT6÷VçBÒ ¢f—'7DFWFV7FVDBÒrs²Æ7E6VVäBÒrs²V–WDFVFÆ–æRÒrp¢7FFRÒv–FÆRs²FVfW%&V6öâÒrs²÷væW%4æÖRÒrs²÷væW$¦ö$–BÒrp¢&WG'”6÷VçBÒ²æW‡E&WG'”BÒrs²Æ7DW'&÷"Òrs²Æ7E&W7VÇBÒrs²Æ7E&W7VÇDBÒrs²æ÷F–f–6F–öä–BÒrp¢Ð§Ð ¦gVæ7F–öâFW7BÔ–çFW&7F—fTW†6VÄ–åW6R…·7G&–æuÒE6÷W&6UF‚’°¢2cRÜ*sRã3¢XÙŽ{IN8¢vWBÕ&ö6W72U„4TÂ8~8þjÚ.8î8(®8ž8î8(¾8 ¢2&W÷'D&–æFW"8î8:Î8;>888:®8;>8+yJ‚W†6VÂ8òf—6–&ÆSÒFfÇ6R8®8î8~8:8*N8;>8*n8*>8;>88ž8*n8).hÈ8þ8®8N8 ¢2ZûîŠ›i8ÞKÙÎ8^8(Î8n8N8(²W†6VÂ8Ž8Zûî‹89^8*8*N8:¾8î8:Þ88>8*þ89^8*8*N8:¾888).Šh¾8(¾8 ¢G'’°¢G&ö72Ò„vWBÕ&ö6W72ÔæÖRU„4TÂÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂv†W&RÔö&¦V7B²EòäÖ–åv–æF÷t†æFÆRÖæRÒ¢–b‚G&ö72ä6÷VçBÖwB’²&WGW&âGG'VRÐ¢Ò6F6‚²Ð¢G'’°¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6÷W&6UF‚’’°¢FF—"Ò7Æ—BÕF‚Õ&VçBE6÷W&6UF€¢FæÖRÒ´”òåF…Ó£¤vWDf–ÆTæÖR‚E6÷W&6UF‚¢FÆö6²Ò¦ö–âÕF‚FF—"‚wâBr²FæÖR¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FÆö6²’²&WGW&âGG'VRÐ¢Ð¢Ò6F6‚²Ð¢&WGW&âFfÇ6P§Ð ¦gVæ7F–öâFW7BÕ&VæFW$Væv–æT'W7’…·7G&–æuÒDÆæwVvR’°¢F‚ÒG'’Ô7V—&TÆö6´†æFÆR„vWBÕ&VæFW$Væv–æTÆö6µF‚DÆæwVvR¢–b‚FçVÆÂÖWF‚’²&WGW&âGG'VRÐ¢&VÆV6RÔÆö6´†æFÆRF€¢&WGW&âFfÇ6P§Ð ¦gVæ7F–öâ7F'BÔWFõ&VæFW$¦ö$f÷%v÷&¶&öö²…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢2cRÕ¢™ÙžjÚ.z+®Š¨Þ8~8þ8î8þ8Î8Þ8îjIÎyú^x˜Ž8Þ8®8î8~88+Ž8:~89nZéþŠÎi˜.8¾XŠ^8îx˜Ž8Ž8ž8(®i»þ8(þ8>8n8þ8N88®8N8 ¢2Zûî‹6æ6†÷N8).iˆîzK®y¨N8¾Y»®Zé®8~8n8+Ž8:~89n8ŽkŠ8ž8 ¢G'’°¢G–ç2Ò·Ð¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6æ6†÷D–B’’°¢FÒÒvWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚FçVÆÂÖæRFÒ’°¢G–ç5²Ev÷&¶&öö´–EÒÒ¶÷&FW&VEÔ²6æ6†÷D–BÒE6æ6†÷D–C²W‡V7FVD†6‚Ò·7G&–æuÒ„vWBÔFF&÷W'G’FÒw6÷W&6T†6‚rrr’Ð¢Ð¢Ð¢F¦ö"Ò7F'BÕ&VæFW$¦ö"DÆæwVvR‚Ev÷&¶&öö´–B’FfÇ6RrrG–ç0¢&WGW&â¶÷&FW&VEÔ²ö²ÒGG'VS²¦ö$–BÒ·7G&–æuÒF¦ö"æ¦ö$–BÐ¢Ò6F6‚°¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²ÖW76vRÒEòäW†6WF–öâäÖW76vRÐ¢Ð§Ð ¦gVæ7F–öâFW7BÔWFôf–ÇW&U7WW'6VFVB‚E7FFRÂ·7G&–æuÒD7W'&VçD†6‚’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’E7FFRw7FFRrrr’ÖæRvf–ÆVBr’²&WGW&âFfÇ6RÐ¢F7W'&VçBÒæ÷&ÖÆ—¦RÔf–ÆT†6‚D7W'&VçD†6€¢Ff–ÆVD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’E7FFRwVæF–æt†6‚rrr’¢&WGW&â‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçB’ÖæBF7W'&VçBÖæRFf–ÆVD†6‚§Ð ¦gVæ7F–öâ–çfö¶RÔWFõ66†VGVÆW%F–6²…·7G&–æuÒDÆæwVvRÂD÷væVDÆö6·2’°¢2cRÜ*sRãRü*tÓ3¢F–6²Xh^8~8+ž8:®8;Î89~8~8®8Nx«nhX¾j™þj+8.89n88>8*þ8N8Ž8¾y»NX‰~8sƒzy.[è^8þ8®8N8 ¢G6WGF–æw2ÒvWBÔWFõ&VæFW%6WGF–æw0¢–b‚Öæ÷B¶&ööÅÒG6WGF–æw2æVæ&ÆVB’²&WGW&âÐ¢2cRÕ¢89n8:ž8*n8+n8ã3zy.8+þ8*N89î8;Î8¾KéÞZÙŽ8~8®8N8.8+ž8+8+Ž8:^8;Î8:ž8;Îˆz®‹ª¾8ÎjIÎyú^8~8jIÎyú^8ŽYÎi˜.8¾KùÞZÙŽ8ž8(¾8 ¢G'’²·fö–EÒ…66âÕWFFW2DÆæwVvRFçVÆÂFfÇ6R’Ò6F6‚²w&—FRÕv&æ–ærEòäW†6WF–öâäÖW76vRÐ¢G'’²·fö–EÒ…WFFRÔ–çWD†—7F÷'”gFW%66âDÆæwVvR’Ò6F6‚²w&—FRÕv&æ–ærEòäW†6WF–öâäÖW76vRÐ¢GF‡2ÒvWBÕF‡0¢G7G'V7GW&RÒFçVÆÀ¢G'’²G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvRÒ6F6‚²&WGW&âÐ¢Fæ÷rÒ´FFUF–ÖUÓ£¥WF4æ÷p ¢f÷&V6‚‚Gr–â„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2’’°¢F–BÒ·7G&–æuÒGrçv÷&¶&öö´–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F–B’’²6öçF–çVRÐ ¢2cRÜ*sRãS¢89n88>8*þXÙŽKØÞ8îh˜iÈžjŠž8.Xùn[é~8~8Þ8®8N89n88>8*þ8þK¹n8+^8;Î898;Î8îh¸^[Ù>8®8î8~›¹ž8>8nš9¾88ž8 ¢–b‚Öæ÷BD÷væVDÆö6·2ä6öçF–ç4¶W’‚F–B’’°¢FÆö6µF‚Ò¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’‚&Æö6·5ÆWFòÖ÷væW%÷³ÒæÆö6²"ÖbF–B¢F‚ÒG'’Ô7V—&TÆö6´†æFÆRFÆö6µF€¢–b‚FçVÆÂÖWF‚’²6öçF–çVRÐ¢D÷væVDÆö6·5²F–EÒÒF€¢Ð ¢G7FFRÒ&VBÔWFõ7FFRDÆæwVvRF–@¢–b‚FçVÆÂÖWG7FFR’²G7FFRÒæWrÔWFõ7FFRF–BÐ ¢G6÷W&6UF‚Ò¦ö–âÕ6fR…·7G&–æuÒGF‡2ç7V&Ö—76–öäF—"’…·7G&–æuÒGrç&VÆF—fUF‚¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G6÷W&6UF‚’’°¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrv–FÆRs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârvf–ÆRÖÖ—76–ærp¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFS²6öçF–çVP¢Ð ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRw7FFRrrr’ÖWw&VæFW&–ærr’°¢F¦ö$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRv÷væW$¦ö$–Brrr¢F¦ö"Ò–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¦ö$–B’’²FçVÆÂÒVÇ6R²&VBÕ&VæFW$¦ö%7FGW2DÆæwVvRF¦ö$–BÐ¢F¦ö%7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’F¦ö"w7FGW2rrr¢–b‚F¦ö%7FGW2Ö–â‚wVWVVBrÂw'Vææ–ærrÂvæÇ—¦–ærr’’²6öçF–çVRÐ¢–b‚F¦ö%7FGW2ÖWv6ö×ÆWFVBr’°¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrv–FÆRs²6WBÔæ÷FU&÷W'G’G7FFRw&WG'”6÷VçBr²6WBÔæ÷FU&÷W'G’G7FFRvæW‡E&WG'”Brrs²6WBÔæ÷FU&÷W'G’G7FFRvÆ7DW'&÷"rrp¢6WBÔæ÷FU&÷W'G’G7FFRvÆ7E&W7VÇBrv6ö×ÆWFVBs²6WBÔæ÷FU&÷W'G’G7FFRvÆ7E&W7VÇDBr„æWrÔæ÷t—6ò“²6WBÔæ÷FU&÷W'G’G7FFRvæ÷F–f–6F–öä–Br„æWrÕ&$–B¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFS²w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvWFòæ6ö×ÆWFVBr…¶÷&FW&VEÔ²v÷&¶&öö´–CÒF–C²¦ö$–CÒF¦ö$–BÒ“²6öçF–çVP¢Ð¢–b‚F¦ö%7FGW2Ö–â‚v6ö×ÆWFVB×v—F‚ÖW'&÷'2rÂvf–ÆVBr’’°¢G&WG'”6÷VçBÒ¶–çEÒ„vWBÔFF&÷W'G’G7FFRw&WG'”6÷VçBr’²¢FÆ7DW'&÷"Ò·7G&–æuÒ„vWBÔFF&÷W'G’F¦ö"vÖW76vRr~ˆz®X¹UDnKÙÎh‰8¾ZKiY~8~8î8~8þ8"r¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FÆ7DW'&÷"’’²FÆ7DW'&÷"Ò·7G&–æuÒ„vWBÔFF&÷W'G’F¦ö"vW'&÷"r~ˆz®X¹UDnKÙÎh‰8¾ZKiY~8~8î8~8þ8"r’Ð¢6WBÔæ÷FU&÷W'G’G7FFRw&WG'”6÷VçBrG&WG'”6÷VçC²6WBÔæ÷FU&÷W'G’G7FFRvÆ7DW'&÷"rFÆ7DW'&÷#²6WBÔæ÷FU&÷W'G’G7FFRvÆ7E&W7VÇBrvf–ÆVBs²6WBÔæ÷FU&÷W'G’G7FFRvÆ7E&W7VÇDBr„æWrÔæ÷t—6ò“²6WBÔæ÷FU&÷W'G’G7FFRvæ÷F–f–6F–öä–Br„æWrÕ&$–B¢–b‚G&WG'”6÷VçBÖÆR¶–çEÒG6WGF–æw2æÖ…&WG'”6÷VçB’°¢FFVÆ’Ò´ÖF…Ó£¤Ö–âƒ3cÅ¶–çEÒG6WGF–æw2ç&WG'”&6U6V6öæG2¢´ÖF…Ó£¥÷rƒ"Å´ÖF…Ó£¤Ö‚ƒÂG&WG'”6÷VçBÓ’’¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrw&WG'’×v—Bs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârw&WG'’Ö&6¶öfbs²6WBÔæ÷FU&÷W'G’G7FFRvæW‡E&WG'”Br‚Fæ÷räFE6V6öæG2‚FFVÆ’’åFõ7G&–ær‚vòr’¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvWFòç&WG'’×66†VGVÆVBr…¶÷&FW&VEÔ²v÷&¶&öö´–CÒF–C²¦ö$–CÒF¦ö$–C²&WG'”6÷VçCÒG&WG'”6÷VçC²æW‡E&WG'”CÕ·7G&–æuÒG7FFRææW‡E&WG'”C²W'&÷#ÒFÆ7DW'&÷"Ò¢ÒVÇ6R°¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrvf–ÆVBs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârw&WG'’ÖW††W7FVBs²6WBÔæ÷FU&÷W'G’G7FFRvæW‡E&WG'”Brrp¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvWFòæf–ÆVBr…¶÷&FW&VEÔ²v÷&¶&öö´–CÒF–C²¦ö$–CÒF¦ö$–C²&WG'”6÷VçCÒG&WG'”6÷VçC²W'&÷#ÒFÆ7DW'&÷"Ò¢Ð¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFS²6öçF–çVP¢Ð¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrwv—F–ærs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârv¦ö"×7FGW2ÖÖ—76–ærs²w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFP¢Ð ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRw7FFRrrr’ÖWw&WG'’×v—Br’°¢FæW‡E&WG'•FW‡BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRvæW‡E&WG'”Brrr¢G&WG'•&VG’ÒGG'VS²–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæW‡E&WG'•FW‡B’’²G'’²G&WG'•&VG’Ò…´FFUF–ÖUÓ£¥'6R‚FæW‡E&WG'•FW‡B’åFõVæ—fW'6ÅF–ÖR‚’ÖÆRFæ÷r’Ò6F6‚²ÒÐ¢–b‚Öæ÷BG&WG'•&VG’’²6öçF–çVRÐ¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrwv—F–ærs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârrs²6WBÔæ÷FU&÷W'G’G7FFRwV–WDFVFÆ–æRr‚Fæ÷räFE6V6öæG2‚Ó’åFõ7G&–ær‚vòr’“²6WBÔæ÷FU&÷W'G’G7FFRw7F&ÆT6÷VçBr“¢Ð ¢F7W'&VçD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÄ†6‚rrr’¢G&VæFW&VD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’GrvÆ7E&VæFW&VDW†6VÄ†6‚rrr’¢2YÎ8Žx˜Ž8~ZKiY~8).{›8(®‹ùN8^8®8N8.8þ88~XŠžyJŽˆ^8ÎXéþz‹þ8).KùÞZÙŽ8~y»N8~8þZNYŽ8þ8¢2ik8~8N88þ88>8+~8:^8).ikŠhþjIÎyú^8Ž8~8nh›8N8[è^j™þ8;¾XhÞŠšnŠÎx«nhX¾8Žˆz®X¹^[êž[‹8ž8(¾8 ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRw7FFRrrr’ÖWvf–ÆVBrÖæBÖæ÷B…FW7BÔWFôf–ÇW&U7WW'6VFVBG7FFRF7W'&VçD†6‚’’²6öçF–çVRÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçD†6‚’Ö÷"F7W'&VçD†6‚ÖWG&VæFW&VD†6‚’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRw7FFRrrr’ÖæRv–FÆRr’°¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrv–FÆRs²6WBÔæ÷FU&÷W'G’G7FFRwVæF–æu6æ6†÷D–Brrp¢6WBÔæ÷FU&÷W'G’G7FFRwVæF–æt†6‚rrs²6WBÔæ÷FU&÷W'G’G7FFRw7F&ÆT6÷VçBr ¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFP¢Ð¢6öçF–çVP¢Ð ¢GVæF–æt†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRwVæF–æt†6‚rrr’¢–b‚GVæF–æt†6‚ÖæRF7W'&VçD†6‚’°¢2ik8~8Nx˜Ž8).jIÎyú^8.[è^j™þ8).8(N8(®y»N8’ŽXúN8NKˆi˜.8+>89N8;Î8þzNj8Bž8 ¢FöÆD6GW&RÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRvW†VÖW&Ä6GW&T–Brrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FöÆD6GW&R’’²&VÖ÷fRÔW†VÖW&Ä6÷’DÆæwVvRFöÆD6GW&RÐ¢6WBÔæ÷FU&÷W'G’G7FFRwVæF–æt†6‚rF7W'&VçD†6€¢6WBÔæ÷FU&÷W'G’G7FFRwVæF–æu6æ6†÷D–Br…·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçE6æ6†÷D–Brrr’¢6WBÔæ÷FU&÷W'G’G7FFRw7F&ÆT6÷VçBr¢6WBÔæ÷FU&÷W'G’G7FFRvf—'7DFWFV7FVDBr„æWrÔæ÷t—6ò¢6WBÔæ÷FU&÷W'G’G7FFRwV–WDFVFÆ–æRr‚Fæ÷räFE6V6öæG2…¶–çEÒG6WGF–æw2çV–WEW&–öE6V6öæG2’åFõ7G&–ær‚vòr’¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrwv—F–ærp¢6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârrp¢6WBÔæ÷FU&÷W'G’G7FFRv÷væW%4æÖRrFVçc¤4ôÕUDU$äÔP¢6WBÔæ÷FU&÷W'G’G7FFRw&WG'”6÷VçBr²6WBÔæ÷FU&÷W'G’G7FFRvæW‡E&WG'”Brrs²6WBÔæ÷FU&÷W'G’G7FFRvÆ7DW'&÷"rrp¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFP¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvWFòæFWFV7FVBr…¶÷&FW&VEÔ²v÷&¶&öö´–BÒF–C²†6‚ÒF7W'&VçD†6‚Ò¢6öçF–çVP¢Ð ¢6WBÔæ÷FU&÷W'G’G7FFRw7F&ÆT6÷VçBr…¶–çEÒ„vWBÔFF&÷W'G’G7FFRw7F&ÆT6÷VçBr’²¢6WBÔæ÷FU&÷W'G’G7FFRvÆ7E6VVäBr„æWrÔæ÷t—6ò ¢FFVFÆ–æUFW‡BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRwV–WDFVFÆ–æRrrr¢FFVFÆ–æU&V6†VBÒFfÇ6P¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FFVFÆ–æUFW‡B’’°¢G'’²FFVFÆ–æU&V6†VBÒ…´FFUF–ÖUÓ£¥'6R‚FFVFÆ–æUFW‡B’åFõVæ—fW'6ÅF–ÖR‚’ÖÆRFæ÷r’Ò6F6‚²FFVFÆ–æU&V6†VBÒGG'VRÐ¢Ð¢–b‚Öæ÷BFFVFÆ–æU&V6†VBÖ÷"¶–çEÒ„vWBÔFF&÷W'G’G7FFRw7F&ÆT6÷VçBr’ÖÇB¶–çEÒG6WGF–æw2ç&WV—&U7F&ÆT†6„6÷VçB’°¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrwv—F–ærs²w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFS²6öçF–çVP¢Ð¢–b…¶&ööÅÒG6WGF–æw2æFVfW%v†–ÆTW†6VÄ–åW6RÖæB…FW7BÔ–çFW&7F—fTW†6VÄ–åW6RG6÷W&6UF‚’’°¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrvFVfW'&VBs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârvW†6VÂÖ–â×W6Rp¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFP¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvWFòæFVfW'&VBr…¶÷&FW&VEÔ²v÷&¶&öö´–BÒF–C²&V6öâÒvW†6VÂÖ–â×W6RrÒ¢6öçF–çVP¢Ð¢–b…FW7BÕ&VæFW$Væv–æT'W7’DÆæwVvR’°¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrvFVfW'&VBs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârv¦ö"Ö'W7’p¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFS²6öçF–çVP¢Ð¢FÖ–æ–×VÔg&VT'—FW2Ò¶–çCcEÒG6WGF–æw2æÖ–äg&VTÖVv'—FW2¢Ô ¢FFFg&VT'—FW2ÒvWBÕF„g&VT'—FW2…·7G&–æuÒGF‡2æFFF—"“²F÷WGWDg&VT'—FW2ÒvWBÕF„g&VT'—FW2…·7G&–æuÒGF‡2æ÷WGWDF—"¢–b‚‚FFFg&VT'—FW2ÖwBÖæBFFFg&VT'—FW2ÖÇBFÖ–æ–×VÔg&VT'—FW2’Ö÷"‚F÷WGWDg&VT'—FW2ÖwBÖæBF÷WGWDg&VT'—FW2ÖÇBFÖ–æ–×VÔg&VT'—FW2’’°¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrvFVfW'&VBs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârvF—6²ÖÆ÷rs²6WBÔæ÷FU&÷W'G’G7FFRvÆ7DW'&÷"r~z›®8ÞZëž˜xþ8ÎKˆÞ‹k>8~8n8N8(¾8þ8(ˆz®X¹UDnKÙÎh‰8).KùÞyYž8~8n8N8î8ž8"p¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFS²6öçF–çVP¢Ð ¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrw&VæFW&–ærs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârrp¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFP¢G7F'FVBÒ7F'BÔWFõ&VæFW$¦ö$f÷%v÷&¶&öö²DÆæwVvRF–B…·7G&–æuÒ„vWBÔFF&÷W'G’G7FFRwVæF–æu6æ6†÷D–Brrr’¢–b…¶&ööÅÒG7F'FVBæö²’°¢6WBÔæ÷FU&÷W'G’G7FFRv÷væW$¦ö$–Br…·7G&–æuÒG7F'FVBæ¦ö$–B¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRw&VæFW"ç7F'FVBr…¶÷&FW&VEÔ²v÷&¶&öö´–BÒF–C²¦ö$–BÒ·7G&–æuÒG7F'FVBæ¦ö$–C²G&–vvW"ÒvWFòrÒ¢ÒVÇ6R°¢G&WG'”6÷VçBÒ¶–çEÒ„vWBÔFF&÷W'G’G7FFRw&WG'”6÷VçBr’²²6WBÔæ÷FU&÷W'G’G7FFRw&WG'”6÷VçBrG&WG'”6÷VçC²6WBÔæ÷FU&÷W'G’G7FFRvÆ7DW'&÷"r…·7G&–æuÒG7F'FVBæÖW76vR¢–b‚G&WG'”6÷VçBÖÆR¶–çEÒG6WGF–æw2æÖ…&WG'”6÷VçB’²FFVÆ“Õ´ÖF…Ó£¤Ö–âƒ3cÅ¶–çEÒG6WGF–æw2ç&WG'”&6U6V6öæG2¥´ÖF…Ó£¥÷rƒ"Å´ÖF…Ó£¤Ö‚ƒÂG&WG'”6÷VçBÓ’’“²6WBÔæ÷FU&÷W'G’G7FFRw7FFRrw&WG'’×v—Bs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârw7F'BÖf–ÆVBs²6WBÔæ÷FU&÷W'G’G7FFRvæW‡E&WG'”Br‚Fæ÷räFE6V6öæG2‚FFVÆ’’åFõ7G&–ær‚vòr’’Ð¢VÇ6R²6WBÔæ÷FU&÷W'G’G7FFRw7FFRrvf–ÆVBs²6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârw&WG'’ÖW††W7FVBs²6WBÔæ÷FU&÷W'G’G7FFRvÆ7E&W7VÇBrvf–ÆVBs²6WBÔæ÷FU&÷W'G’G7FFRvÆ7E&W7VÇDBr„æWrÔæ÷t—6ò“²6WBÔæ÷FU&÷W'G’G7FFRvæ÷F–f–6F–öä–Br„æWrÕ&$–B’Ð¢Ð¢w&—FRÔWFõ7FFRDÆæwVvRF–BG7FFP¢Ð§Ð ¦gVæ7F–öâ–çfö¶RÔWFõ66†VGVÆW$g&öÔf–ÆR…·7G&–æuÒD6öçG&öÅF‚Â¶–çEÒE&VçE&ö6W74–B’°¢2cRÜ*sRãs¢™ÙžjÚ.[è^88ò…EE8:®8+ž88®8;Î8îKŠÞ8~ŠÎ8(þ8®8N8 ¢2cB8î8+^8;Î898;Î8þXÙŽKˆ8+ž8:Î88>88ž8â66WEF76Æ–VçB8:¾8;Î89~8®8î8~8†æFÆRÔ’Xh^8~[è^8N8ŽyK¾™Ú.8ÎjÚ.8î8(¾8 ¢FÆæwVvRÒv¦p¢G'’°¢F6öçG&öÂÒ&VBÔ§6öäf–ÆRD6öçG&öÅF‚FçVÆÀ¢–b‚FçVÆÂÖæRF6öçG&öÂ’²FÆæwVvRÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6öçG&öÂvÆæwVvRrv¦r’Ð¢Ò6F6‚²Ð¢F÷væVBÒ·Ð¢F†—7F÷'”6ÆVçWVæF–ærÒGG'VP¢G'’°¢6ÆV"ÔW‡—&VDÆV6W2FÆæwVvP¢6ÆV"Õ7FÆTW†VÖW&Ä6÷–W2FÆæwVvP¢&V6÷fW"ÔWFõ7FFW2FÆæwVvP¢v†–ÆR‚GG'VR’°¢2Šj¥”Nyº>Šin8¾Xª8Ž87F÷89^8*8*N8:¾8)#×>XÙŽKØÞ8~z+®Š¨Þ8ž8(¾8 ¢2[é>iÚ^8ãzy%6ÆVW8~8þjÚ>[‹Ž{X.K¨n8~8(.Šj®8ÎZÙ8)$¶–ÆÎ8ž8(¾[ø^Šh8Î8.8(®8f–æÆÇž8îx«nhX¾[êžiz~8Î‹[8(ž8®8¾8>8þ8 ¢–b‚E&VçE&ö6W74–BÖÆRÖ÷"FçVÆÂÖW„vWBÕ&ö6W72Ô–BE&VçE&ö6W74–BÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’²'&V²Ð¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚‚D6öçG&öÅF‚²rç7F÷r’’²'&V²Ð¢G'’²–çfö¶RÔWFõ66†VGVÆW%F–6²FÆæwVvRF÷væVBÒ6F6‚²w&—FRÕv&æ–ærEòäW†6WF–öâäÖW76vRÐ¢G'’²6ÆV"Õ7FÆTW†VÖW&Ä6÷–W2FÆæwVvRÒ6F6‚²Ð¢2XÎjÚ.z+®Š¨Þ8òS×2™i>™©N8$6öçG&öÅF‚8þX[iÈž88ž8:ž8*N89nKˆ®8¾8.8(¾8þ8(8¢2×2™i>™©N88ŽXŠžyJŽˆ^8N8Ž8¾jøîzy#Y¹î8å4Ô"FW7BÕF‚8Î[‹Ži˜.y›®yIþ8ž8(¾8 ¢2Šj®XN8òv—Df÷$W†—Bƒ#S’[è^8N8î8~8S×28~8(.jÚ>[‹Ž{X.K¨n8Ž[èÎx˜~K¹Ž88þ™i>8¾YŽ8n8 ¢G7F÷&WVW7FVBÒFfÇ6P¢f÷"‚F’Ò²F’ÖÇB#²F’²²’°¢–b‚E&VçE&ö6W74–BÖÆRÖ÷"FçVÆÂÖW„vWBÕ&ö6W72Ô–BE&VçE&ö6W74–BÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’Ö÷"…FW7BÕF‚ÔÆ—FW&ÅF‚‚D6öçG&öÅF‚²rç7F÷r’’’°¢G7F÷&WVW7FVBÒGG'VP¢'&V°¢Ð¢7F'BÕ6ÆVWÔÖ–ÆÆ—6V6öæG2S ¢Ð¢–b‚G7F÷&WVW7FVB’²'&V²Ð¢2†—7F÷'’&WFVçF–öâvÆ·2WfW'’6fVBv÷&¶&öö²÷fW'6–öâ†–æ6ÇVF–ær66†V@¢2&7FW"vW2’â'Vâ—BgFW"F†RT’†2†BF–ÖRFò&V6öÖR&VG’–ç7FV@¢2öbÖ¶–ærWfW'’Æ–6F–öâÆVæ6‚v—Bf÷"F†RgVÆÂF—&V7F÷'’66âà¢–b‚F†—7F÷'”6ÆVçWVæF–ær’°¢G'’²–çfö¶RÔ–çWD†—7F÷'”6ÆVçWFÆæwVvRÒ6F6‚²w&—FRÕv&æ–ærEòäW†6WF–öâäÖW76vRÐ¢F†—7F÷'”6ÆVçWVæF–ærÒFfÇ6P¢Ð¢Ð¢Òf–æÆÇ’°¢f÷&V6‚‚F²–â‚F÷væVBä¶W—2’’²&VÆV6RÔÆö6´†æFÆRF÷væVE²FµÒÐ¢G'’°¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚„vWBÔWFõ7FFTF—"FÆæwVvR’Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢G7BÒ&VBÔ§6öäf–ÆRFbägVÆÄæÖRFçVÆÀ¢–b‚FçVÆÂÖæRG7BÖæB‚w&VæFW&–ærrÂw&VG’r’Ö6öçF–ç2·7G&–æuÒ„vWBÔFF&÷W'G’G7Bw7FFRrrr’’°¢6WBÔæ÷FU&÷W'G’G7Bw7FFRrwv—F–ærp¢w&—FRÔ§6öäf–ÆRFbägVÆÄæÖRG7@¢Ð¢Ð¢Ò6F6‚²Ð¢2jÚ>[‹Ž{X.K¨n8î8þ8>8¾X‹n[ê¥4ôâòç7F÷8).jè¾8^8®8N8 ¢f÷&V6‚‚GF‚–â‚D6öçG&öÅF‚Â‚D6öçG&öÅF‚²rç7F÷r’’’°¢G'’²–b‚GF‚ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒÒ6F6‚²Ð¢Ð¢Ð§Ð ¦gVæ7F–öâ&V6÷fW"ÔWFõ7FFW2…·7G&–æuÒDÆæwVvR’°¢2cRÜ*sRãS¢‹[~X¹^i˜.8:®8*¾898:®8;Î8 ¢G'’°¢FF—"ÒvWBÔWFõ7FFTF—"DÆæwVvP¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&âÐ¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢G7BÒ&VBÔ§6öäf–ÆRFbägVÆÄæÖRFçVÆÀ¢–b‚FçVÆÂÖWG7B’²6öçF–çVRÐ¢G7FFRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7Bw7FFRrrr¢F6†ævVBÒFfÇ6P¢–b‚G7FFRÖWw&VæFW&–ærr’²6WBÔæ÷FU&÷W'G’G7Bw7FFRrwv—F–ærs²F6†ævVBÒGG'VRÐ¢VÇ6V–b‚G7FFRÖWw&VG’r’²6WBÔæ÷FU&÷W'G’G7Bw7FFRrwv—F–ærs²F6†ævVBÒGG'VRÐ¢GVæF–ærÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7BwVæF–æu6æ6†÷D–Brrr¢Gv$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7Bwv÷&¶&öö´–Brrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GVæF–ær’ÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Gv$–B’’°¢–b‚FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvRGv$–BGVæF–ær’’°¢6WBÔæ÷FU&÷W'G’G7BwVæF–æu6æ6†÷D–Brrs²6WBÔæ÷FU&÷W'G’G7BwVæF–æt†6‚rrp¢6WBÔæ÷FU&÷W'G’G7Bw7F&ÆT6÷VçBr²6WBÔæ÷FU&÷W'G’G7Bw7FFRrv–FÆRs²F6†ævVBÒGG'VP¢Ð¢Ð¢–b‚F6†ævVB’²w&—FRÔ§6öäf–ÆRFbägVÆÄæÖRG7BÐ¢Ð¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâ7F'BÔWFõ66†VGVÆW%&ö6W72…·7G&–æuÒDÆæwVvR’°¢F6öçG&öÅF‚Òrp¢G'’°¢G6WGF–æw2ÒvWBÔWFõ&VæFW%6WGF–æw0¢–b‚Öæ÷B¶&ööÅÒG6WGF–æw2æVæ&ÆVB’²&WGW&âFçVÆÂÐ¢FF—"Ò¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’w7FFRp¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢2X‹n[ê89^8*8*N8:¾8õ>YÒ¾Šj¥”N8~89~8:Þ8+¾8+žXÙŽKØÞ8¾Xˆn™º.8ž8(¾8 ¢F6öçG&öÄæÖRÒvWFò×66†VGVÆW%÷³Õ÷³Òæ§6öârÖb…·&VvW…Ó£¥&WÆ6R…·7G&–æuÒFVçc¤4ôÕUDU$äÔRÂuµäÕ¦×£Ó•òâÕÒ²rÂuòr’’ÂE”@¢F6öçG&öÅF‚Ò¦ö–âÕF‚FF—"F6öçG&öÄæÖP¢w&—FRÔ§6öäf–ÆRF6öçG&öÅF‚…¶÷&FW&VEÔ²66†VÖfW'6–öâÒ²ÆæwVvRÒDÆæwVvS²7F'FVDBÒæWrÔæ÷t—6ó²&VçE–BÒE”C²4æÖRÒ·7G&–æuÒFVçc¤4ôÕUDU$äÔRÒ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚‚F6öçG&öÅF‚²rç7F÷r’’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚‚F6öçG&öÅF‚²rç7F÷r’Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢2&ö÷BòX‹n[ê89^8*8*N8:¾8î898+ž8¾z›®y›Þ8ÎY
¾8î8(Î8n8(.ZÙ89~8:Þ8+¾8+ž8Î‹[~X¹^8~8Þ8(¾8(Ž8niˆîzK®y¨N8¾[É^yJŽ8ž8(¾8 ¢G6W'fW%67&—BÒ¦ö–âÕF‚E67&—C¤&ö÷Bw6W'fW"ç3p¢G6’Ò‚rÔæõ&öf–ÆRrÂrÔW†V7WF–öåöÆ–7’rÂt'—72rÂrÔf–ÆRrÂ‚r'³Ò"rÖbG6W'fW%67&—B’À¢rÔÖöFRrÂDÆæwVvRÂrÔWFõ66†VGVÆW%F‚rÂ‚r'³Ò"rÖbF6öçG&öÅF‚’ÂrÕ&VçE&ö6W74–BrÂ·7G&–æuÒE”B¢G&ö2Ò7F'BÕ&ö6W72Ôf–ÆUF‚w÷vW'6†VÆÂæW†RrÔ&wVÖVçDÆ—7BG6’Õv–æF÷u7G–ÆR†–FFVâÕ75F‡'P¢E67&—C¤WFõ66†VGVÆW%&ö6W72ÒG&ö0¢E67&—C¤WFõ66†VGVÆW%&ö6W74–BÒG&ö2ä–@¢E67&—C¤WFõ66†VGVÆW$6öçG&öÅF‚ÒF6öçG&öÅF€¢&WGW&âG&ö0¢Ò6F6‚°¢f÷&V6‚‚GF‚–â‚F6öçG&öÅF‚Â‚F6öçG&öÅF‚²rç7F÷r’’’°¢G'’²–b‚GF‚ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒÒ6F6‚²Ð¢Ð¢w&—FRÕv&æ–ær‚~ˆz®X¹^8+ž8+8+Ž8:^8;Î8:ž8;Î8).‹[~X¹^8~8Þ8î8¾8)>8~8~8ó¢r²EòäW†6WF–öâäÖW76vR¢&WGW&âFçVÆÀ¢Ð§Ð ¦gVæ7F–öâFW7BÔWFõ66†VGVÆW%&ö6W75'Vææ–ær°¢G'’°¢–b‚FçVÆÂÖæRE67&—C¤WFõ66†VGVÆW%&ö6W72’°¢E67&—C¤WFõ66†VGVÆW%&ö6W72å&Vg&W6‚‚¢&WGW&â‚Öæ÷BE67&—C¤WFõ66†VGVÆW%&ö6W72ä†4W†—FVB¢Ð¢–b‚E67&—C¤WFõ66†VGVÆW%&ö6W74–BÖwB’°¢&WGW&â‚FçVÆÂÖæR„vWBÕ&ö6W72Ô–BE67&—C¤WFõ66†VGVÆW%&ö6W74–BÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’¢Ð¢Ò6F6‚²Ð¢&WGW&âFfÇ6P§Ð ¦gVæ7F–öâ7F÷ÔWFõ66†VGVÆW%&ö6W72°¢F6öçG&öÅF‚Ò·7G&–æuÒE67&—C¤WFõ66†VGVÆW$6öçG&öÅF€¢G'’°¢–b‚F6öçG&öÅF‚’°¢6WBÔ6öçFVçBÔÆ—FW&ÅF‚‚F6öçG&öÅF‚²rç7F÷r’ÕfÇVRw7F÷rÔVæ6öF–ær44”’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢Ð¢GÒE67&—C¤WFõ66†VGVÆW%&ö6W70¢–b‚FçVÆÂÖWGÖæBE67&—C¤WFõ66†VGVÆW%&ö6W74–BÖwB’°¢GÒvWBÕ&ö6W72Ô–BE67&—C¤WFõ66†VGVÆW%&ö6W74–BÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢Ð¢–b‚FçVÆÂÖæRG’°¢2ZÙXN8ó×>XÙŽKØÞ8w7F÷8).z+®Š¨Þ8ž8(¾8þ8(88î8®jÚ>[‹Ž{X.K¨n8†f–æÆÇž8î[èÎx˜~K¹Ž88).[è^8N8 ¢FW†—FVBÒFfÇ6P¢G'’²FW†—FVBÒGåv—Df÷$W†—Bƒ#S’Ò6F6‚²Ð¢–b‚Öæ÷BFW†—FVB’°¢G'’²Gä¶–ÆÂ‚’Ò6F6‚²Ð¢G'’²·fö–EÒGåv—Df÷$W†—Bƒ’Ò6F6‚²Ð¢Ð¢Ð¢Ò6F6‚²Ð¢f–æÆÇ’°¢f÷&V6‚‚GF‚–â‚F6öçG&öÅF‚Â‚F6öçG&öÅF‚²rç7F÷r’’’°¢G'’²–b‚GF‚ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒÒ6F6‚²Ð¢Ð¢E67&—C¤WFõ66†VGVÆW%&ö6W72ÒFçVÆÀ¢E67&—C¤WFõ66†VGVÆW%&ö6W74–BÒ ¢E67&—C¤WFõ66†VGVÆW$6öçG&öÅF‚Òrp¢Ð§Ð  ¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ¢2cR7FvRB(	B†6R$#¢yK¾X8þ88þ88>8+~8:^8ŽjùN‹È0¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ ¢E67&—C¥f—7VÄ†6…&öf–ÆUfW'6–öâÒ0¢E67&—C¥f—7VÄ†6…vTF—7Fæ6TÆ–Ö—BÒãsP¢E67&—C¥f—7VÄ†6„fW&vTF—7Fæ6TÆ–Ö—BÒãC ¢E67&—C¥f—7VÄ†6„G’Ò#  ¦gVæ7F–öâvWBÕ&VæFW%&V6÷&DF—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEfW'6–öä–B’°¢&WGW&â„¦ö–âÕF‚„vWBÕ6æ6†÷DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’„¦ö–âÕF‚w&VæFW'2rEfW'6–öä–B’§Ð ¦gVæ7F–öâvWBÕ&VæFW%&7FW%6†VWDF—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒE6†VWDæÖR’°¢G&V6÷&DF—"ÒvWBÕ&VæFW%&V6÷&DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–@¢&WGW&â„¦ö–âÕF‚G&V6÷&DF—"„¦ö–âÕF‚w&7FW"×cr„vWBÔF–fe6†VWD¶W’E6†VWDæÖR’’§Ð ¦gVæ7F–öâvWBÕf—7VÄ†6…&öf–ÆR°¢GFd&÷…fW'6–öâÒrp¢G'’°¢GfbÒ¦ö–âÕF‚E67&—C¤&ö÷BvÆ–%ÇFf&÷…ÅDd$õ…õdU%4”ôâçG‡Bp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚Gfb’²GFd&÷…fW'6–öâÒ‚„vWBÔ6öçFVçBÔÆ—FW&ÅF‚GfbÕ&r’×&WÆ6RuÇ2²rÂrr’åG&–Ò‚’Ð¢Ò6F6‚²Ð¢&WGW&â¶÷&FW&VEÔ°¢&öf–ÆUfW'6–öâÒE67&—C¥f—7VÄ†6…&öf–ÆUfW'6–öà¢Fd&÷…fW'6–öâÒGFd&÷…fW'6–öà¢G’ÒE67&—C¥f—7VÄ†6„G¢6öÆ÷$ÖöFRÒu$t"p¢Ð§Ð ¦gVæ7F–öâvWBÔ†W„†ÖÖ–æu&F–ò…·7G&–æuÒDÆVgBÂ·7G&–æuÒE&–v‡B’°¢FÒ…·7G&–æuÒDÆVgB’åG&–Ò‚’åFõWW$–çf&–çB‚¢F"Ò…·7G&–æuÒE&–v‡B’åG&–Ò‚’åFõWW$–çf&–çB‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F’Ö÷"FäÆVæwF‚ÖæRF"äÆVæwF‚’²&WGW&âãÐ¢FF–ffW&VçD&—G2Ò ¢f÷"‚F’Ò²F’ÖÇBFäÆVæwFƒ²F’²²’°¢G'’°¢G†÷"Ò…´6öçfW'EÓ£¥Fô–çC3"‚F²F•ÒåFõ7G&–ær‚’Âb’Ö'†÷"´6öçfW'EÓ£¥Fô–çC3"‚F%²F•ÒåFõ7G&–ær‚’Âb’¢Ò6F6‚²&WGW&âãÐ¢v†–ÆR‚G†÷"ÖwB’°¢FF–ffW&VçD&—G2³Ò‚G†÷"Ö&æB¢G†÷"ÒG†÷"×6‡"¢Ð¢Ð¢&WGW&â…¶F÷V&ÆUÒFF–ffW&VçD&—G2ò´ÖF…Ó£¤Ö‚ƒÂFäÆVæwF‚¢B’§Ð ¦gVæ7F–öâFW7BÕ6†VWEf—7VÄWV—fÆVçB‚D&Vf÷&RÂDgFW"’°¢–b‚FçVÆÂÖWD&Vf÷&RÖ÷"FçVÆÂÖWDgFW"’²&WGW&âFfÇ6RÐ¢–b‚„vWBÔ–çDFF&÷W'G’D&Vf÷&RwvT6÷VçBrÓ’ÖæR„vWBÔ–çDFF&÷W'G’DgFW"wvT6÷VçBrÓ"’’²&WGW&âFfÇ6RÐ¢F&Vf÷&UFW‡BÒæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’D&Vf÷&RwFW‡D†6‚rrr’¢FgFW%FW‡BÒæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’DgFW"wFW‡D†6‚rrr’¢F†46ö×&&ÆUFW‡BÒÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&Vf÷&UFW‡B’ÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FgFW%FW‡B¢–b‚F†46ö×&&ÆUFW‡BÖæBF&Vf÷&UFW‡BÖæRFgFW%FW‡B’²&WGW&âFfÇ6RÐ ¢F&Vf÷&UvW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’D&Vf÷&RwvUW&6WGVÄ†6†W2r‚’’¢FgFW%vW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’DgFW"wvUW&6WGVÄ†6†W2r‚’’¢F&Vf÷&TW†7EvW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’D&Vf÷&RwvT†6†W2r‚’’¢FgFW$W†7EvW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’DgFW"wvT†6†W2r‚’’¢27W'&VçBæÇ—¦W"&V6÷&G2æ÷&ÖÆ—¦VB$t"—†VÇ2âv†Vâ&÷F‚vVæW&F–öç2†fP¢2F†÷6R†6†W2ÂF†W’&RWF†÷&—FF—fS¢F†RöÆB3'ƒ3"w&—66ÆRW&6WGVÂ†6€¢26â†–FR6VÆÂf–ÆÂÂÆR6öÆ÷"Â&÷&FW"6öÆ÷"Â÷"6ÖÆÂG&v–ær&W6—¦Rà¢–b‚F&Vf÷&TW†7EvW2ä6÷VçBÖwBÖæBF&Vf÷&TW†7EvW2ä6÷VçBÖWFgFW$W†7EvW2ä6÷VçB’°¢f÷"‚F’Ò²F’ÖÇBF&Vf÷&TW†7EvW2ä6÷VçC²F’²²’°¢–b‚„æ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒF&Vf÷&TW†7EvW5²F•Ò’’ÖæP¢„æ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒFgFW$W†7EvW5²F•Ò’’’²&WGW&âFfÇ6RÐ¢Ð¢&WGW&âGG'VP¢Ð¢2ÆVv7’&V6÷&G2v—F†÷WBæ÷&ÖÆ—¦VBvR†6†W2&WF–âF†RW&6WGVÂfÆÆ&6²à¢–b‚F&Vf÷&UvW2ä6÷VçBÖWÖ÷"F&Vf÷&UvW2ä6÷VçBÖæRFgFW%vW2ä6÷VçB’²&WGW&âFfÇ6RÐ¢GF÷FÂÒã ¢GvTÆ–Ö—BÒB†–b‚F†46ö×&&ÆUFW‡B’²E67&—C¥f—7VÄ†6…vTF—7Fæ6TÆ–Ö—BÒVÇ6R²ã3RÒ¢FfW&vTÆ–Ö—BÒB†–b‚F†46ö×&&ÆUFW‡B’²E67&—C¥f—7VÄ†6„fW&vTF—7Fæ6TÆ–Ö—BÒVÇ6R²ã#Ò¢f÷"‚F’Ò²F’ÖÇBF&Vf÷&UvW2ä6÷VçC²F’²²’°¢FF—7Fæ6RÒvWBÔ†W„†ÖÖ–æu&F–ò…·7G&–æuÒF&Vf÷&UvW5²F•Ò’…·7G&–æuÒFgFW%vW5²F•Ò¢–b‚FF—7Fæ6RÖwBGvTÆ–Ö—B’²&WGW&âFfÇ6RÐ¢GF÷FÂ³ÒFF—7Fæ6P¢Ð¢&WGW&â‚‚GF÷FÂò´ÖF…Ó£¤Ö‚ƒÂF&Vf÷&UvW2ä6÷VçB’’ÖÆRFfW&vTÆ–Ö—B§Ð ¦gVæ7F–öâFW7BÕFevTæÇ—¦W$f–Æ&ÆR°¢–b‚FçVÆÂÖæRE67&—C¥FevTæÇ—¦W$f–Æ&ÆR’²&WGW&â¶&ööÅÒE67&—C¥FevTæÇ—¦W$f–Æ&ÆRÐ¢E67&—C¥FevTæÇ—¦W$f–Æ&ÆRÒFfÇ6P¢G'’°¢F¦"Ò¦ö–âÕF‚E67&—C¤&ö÷BvÆ–%ÇFf&÷…Å&W÷'EFd6ö×÷6W"æ¦"p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚F¦"’’²&WGW&âFfÇ6RÐ¢FBÕG—RÔ76VÖ&Ç”æÖRu7—7FVÒä”òä6ö×&W76–öâäf–ÆU7—7FVÒrÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢G¦—Ò´”òä6ö×&W76–öâå¦—f–ÆUÓ£¤÷Vå&VB‚F¦"¢G'’²E67&—C¥FevTæÇ—¦W$f–Æ&ÆRÒ‚G¦—äVçG&–W2Âv†W&RÔö&¦V7B²EòägVÆÄæÖRÖWuFevTæÇ—¦W"æ6Æ72rÒ’ä6÷VçBÖwBÐ¢f–æÆÇ’²G¦—äF—7÷6R‚’Ð¢Ò6F6‚²E67&—C¥FevTæÇ—¦W$f–Æ&ÆRÒFfÇ6RÐ¢&WGW&â¶&ööÅÒE67&—C¥FevTæÇ—¦W$f–Æ&ÆP§Ð ¦gVæ7F–öâ–çfö¶RÕFevTæÇ—¦W"…¶†6‡F&ÆUµÕÒE6†VWG2’°¢289n88>8*þ8N8Ž8²¦f8)#Y¹î88‹[~X¹^8~8nXZŽ8+~8;Î88Ž8).Šz>ié8ž8(¾8 ¢–b‚FçVÆÂÖWE6†VWG2Ö÷"E6†VWG2ä6÷VçBÖW’²&WGW&âFçVÆÂÐ¢GFööÂÒFçVÆÀ¢G'’²GFööÂÒvWBÕFd&F6…FööÄ–æfòÒ6F6‚²Ð¢F¦fW†RÒrp¢F7Òrp¢G'’°¢F¦fW†RÒ&W6öÇfRÔ¦fW†P¢F7Ò„¦ö–âÕF‚E67&—C¤&ö÷BvÆ–%ÇFf&÷…Å&W÷'EFd6ö×÷6W"æ¦"r’²s²r²„¦ö–âÕF‚E67&—C¤&ö÷BvÆ–%ÇFf&÷…ÇFf&÷‚Öæ¦"r¢Ò6F6‚²&WGW&âFçVÆÂÐ¢2cRÕ¢¤"8²FevTæÇ—¦W"æ6Æ728ÎXZ^8>8n8N8®8N˜XÞ[ˆ>xšž8~8þ8Šz>ié888).›¹ž8>8nŠºn8(8(¾8 ¢2†'V–ÆBç38).ZéþŠÎ8~8b¤"8).XhÞyIþh‰8ž8(¾8ŽiÈžX«ž8¾8®8(²¢–b‚Öæ÷B…FW7BÕFevTæÇ—¦W$f–Æ&ÆR’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²ÖW76vRÒuFevTæÇ—¦W"8Â¤"8¾Y
¾8î8(Î8n8N8î8¾8)>8&ÆÆ–%ÇFf&÷…Æ'V–ÆBç38).ZéþŠÎ8~8n8þ88^8N8"rÒÐ¢GF×F—"Ò¦ö–âÕF‚…´”òåF…Ó£¤vWEFV×F‚‚’’‚w&"ÖæÇ—¦RÒr²„æWrÕ&$–B’¢æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚GF×F—"Ôf÷&6RÂ÷WBÔçVÆÀ¢G'’°¢G&WÒ¦ö–âÕF‚GF×F—"w&WVW7Bæ§6öâp¢G&W2Ò¦ö–âÕF‚GF×F—"w&W7VÇBæ§6öâp¢w&—FRÔ§6öäf–ÆRG&W…¶÷&FW&VEÔ²6†VWG2Ò‚E6†VWG2Âf÷$V6‚Ôö&¦V7B°¢¶÷&FW&VEÔ°¢6†VWDæÖRÒ·7G&–æuÒEòç6†VWDæÖP¢FbÒ·7G&–æuÒEòçF`¢&7FW$F—&V7F÷'’Ò·7G&–æuÒ„vWBÔFF&÷W'G’Eòw&7FW$F—&V7F÷'’rrr¢Ð¢Ò’Ò¢G'VâÒ–çfö¶RÔæF—fT6GW&RF¦fW†R‚rÔF¦fæwBæ†VFÆW73×G'VRrÂrÖ7rÂF7ÂuFevTæÇ—¦W"rÂrÒÖ–çWBrÂG&WÂrÒÖ÷WGWBrÂG&W2ÂrÒÖG’rÂ·7G&–æuÒE67&—C¥f—7VÄ†6„G’’E67&—C¥FdæÇ—¦UF–ÖV÷WE6V6öæG0¢–b‚G'VâçF–ÖVD÷WB’°¢2Šz>ié8õDnKÙÎh‰8î8*þ8:®88n8*>8*¾8:¾898+žZIn8.ZKiY~8Ž8~8n‹ùN8ž8Î8Š‰Ž˜Ë.8¾8þynyK8).jè¾8ž8 ¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²ÖW76vRÒ‚~Šh¾8þyºî8îŠz>ié8Ç³ÞXˆnKº^Xh^8¾{X.8(þ8(®8î8¾8)>8~8~8þ8"rÖb¶–çEÒ‚E67&—C¥FdæÇ—¦UF–ÖV÷WE6V6öæG2òc’’Ð¢Ð¢–b…¶–çEÒG'VâæW†—D6öFRÖæRÖ÷"Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G&W2’’°¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²ÖW76vRÒ‚&W†—CÒ"²·7G&–æuÒG'VâæW†—D6öFR²&â"²·7G&–æuÒG'VâçFW‡B’Ð¢Ð¢G'6VBÒ&VBÔ§6öäf–ÆRG&W2FçVÆÀ¢&WGW&â¶÷&FW&VEÔ²ö²ÒGG'VS²&W7VÇBÒG'6VBÐ¢Ò6F6‚°¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²ÖW76vRÒEòäW†6WF–öâäÖW76vRÐ¢Òf–æÆÇ’°¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF×F—"’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF×F—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ð§Ð ¦gVæ7F–öâw&—FRÕ&VæFW%&V6÷&B…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒEW'÷6RÂ¶&ööÅÒD6öçFVçEFe&WF–æVBÂDæÇ—6—2’°¢2cRÔ•bÓ3¢8:Î8;>888:®8;>8+{YiéÎ8ò&VæFW'5ÃÇfW'6–öä–CåÂ8¾K‰nKº>8N8Ž8¾{Úî8Þ8i»Ž8N8þ8(žZHži»N8~8®8N8 ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6æ6†÷D–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EfW'6–öä–B’’²&WGW&âÐ¢G'’°¢FF—"ÒvWBÕ&VæFW%&V6÷&DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢FÖæ–fW7EF‚Ò¦ö–âÕF‚FF—"w&VæFW"ÖÖæ–fW7Bæ§6öâp¢–b‚Öæ÷B…FW7BÔf–ÆTW†—7G46ö×BFÖæ–fW7EF‚’’°¢w&—FRÔ§6öäf–ÆRFÖæ–fW7EF‚…¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ¢W'÷6RÒEW'÷6P¢2–Ö×WF&ÆR&VæFW"&V6÷&C¢F†—26—2v†WF†W"6öçFVçBDbv2&öGV6VBf÷"F†—2&VæFW"à¢27W'&VçB&WFVçF–öâ—2FWFW&Ö–æVBg&öÒF†R7GVÂ6öçFVçB×FbF—&V7F÷'’ÂæWfW"g&öÒF†—2Öæ–fW7Bà¢6öçFVçEFe&öGV6VBÒD6öçFVçEFe&WF–æV@¢6öçFVçEFe&WF–æVBÒD6öçFVçEFe&WF–æVB2ÆVv7’6ö×F–&–Æ—G“²Fòæ÷BW6R27W'&VçB×7FFRG'WF€¢6÷W&6U6æ6†÷D–BÒE6æ6†÷D–@¢v÷&¶&öö´–BÒEv÷&¶&öö´–@¢fW'6–öä–BÒEfW'6–öä–@¢7&VFVDBÒæWrÔæ÷t—6ð¢7FGW2Òv6ö×ÆWFRp¢&VæFW$Vçf—&öæÖVçDf–ævW'&–çBÒ·7G&–æuÒE67&—C¤7W'&VçE&VæFW$Vçdf–ævW'&–ç@¢&VæFW$Vçf—&öæÖVçBÒE67&—C¤7W'&VçE&VæFW$Vçd–æfð¢W†6VÅ&–çE&öf–ÆUfW'6–öâÒE67&—C¤W†6VÅ&–çE&öf–ÆUfW'6–öà¢f—7VÄ†6…&öf–ÆRÒ„vWBÕf—7VÄ†6…&öf–ÆR¢Ò¢Ð¢F†6…F‚Ò¦ö–âÕF‚FF—"wf—7VÂÖ†6†W2æ§6öâp¢–b‚FçVÆÂÖæRDæÇ—6—2ÖæBÖæ÷B…FW7BÔf–ÆTW†—7G46ö×BF†6…F‚’’°¢w&—FRÔ§6öäf–ÆRF†6…F‚…¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ¢6æ6†÷D–BÒE6æ6†÷D–@¢fW'6–öä–BÒEfW'6–öä–@¢&VæFW$Vçf—&öæÖVçDf–ævW'&–çBÒ·7G&–æuÒE67&—C¤7W'&VçE&VæFW$Vçdf–ævW'&–ç@¢f—7VÄ†6…&öf–ÆRÒ„vWBÕf—7VÄ†6…&öf–ÆR¢æÇ—¦W%fW'6–öâÒ¶–çEÒ„vWBÔFF&÷W'G’DæÇ—6—2væÇ—¦W%fW'6–öâr¢¦ffW'6–öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’DæÇ—6—2v¦ffW'6–öârrr¢¦ffVæF÷"Ò·7G&–æuÒ„vWBÔFF&÷W'G’DæÇ—6—2v¦ffVæF÷"rrr¢6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’DæÇ—6—2w6†VWG2r‚’’¢Ò¢Ð¢2[^jÛNyK¾™Ú.8þ˜xÞ8NK‰nKº>‹[iû¾8).˜þ88(¾8þ8(Šh{HN8).8*Þ8:>88>8+~8:^8~8n8N8(¾8 ¢2&VæFW"ÖÖæ–fW7Bòf—7VÂÖ†6†W28þ[^jÛN8:¾8;Î88Žy»NKˆ¾8).i»Nik8~8®8N8þ8(8¢28>8>8~zNj8N8~8®8N8Ž8ÎyK¾X8þ88þ88>8+~8:^8®8~8Þ8îKÙÎh‰X˜ÞXŠNZé®8Îjè¾8(®{i®88(¾8 ¢6ÆV"Õ6æ6†÷E7VÖÖ'”66†RDÆæwVvREv÷&¶&öö´–@¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâvWBÕf—7VÄ†6†W2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEfW'6–öä–B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Bp¢G6fUfW'6–öä–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEfW'6–öä–BwfW'6–öä–Bp¢F66†T¶W’Ò‚DÆæwVvR²wÂr²G6fUv÷&¶&öö´–B²wÂr²G6fU6æ6†÷D–B²wÂr²G6fUfW'6–öä–B’åFôÆ÷vW$–çf&–çB‚¢GÒ¦ö–âÕF‚„vWBÕ&VæFW%&V6÷&DF—"DÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–BG6fUfW'6–öä–B’wf—7VÂÖ†6†W2æ§6öâp¢–b‚E67&—C¥f—7VÄ†6„66†Rä6öçF–ç4¶W’‚F66†T¶W’’’°¢–b…FW7BÔf–ÆTW†—7G46ö×BG’²&WGW&âE67&—C¥f—7VÄ†6„66†U²F66†T¶W•ÒÐ¢·fö–EÒE67&—C¥f—7VÄ†6„66†Rå&VÖ÷fR‚F66†T¶W’¢Ð¢–b‚Öæ÷B…FW7BÔf–ÆTW†—7G46ö×BG’’²&WGW&âFçVÆÂÐ¢G'’°¢F†6†W2Ò&VBÔ§6öäf–ÆRGFçVÆÀ¢–b‚FçVÆÂÖæRF†6†W2’²E67&—C¥f—7VÄ†6„66†U²F66†T¶W•ÒÒF†6†W2Ð¢&WGW&âF†6†W0¢Ò6F6‚²&WGW&âFçVÆÂÐ§Ð ¦gVæ7F–öâvWBÕ&VæFW%fW'6–öä–G2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢FF—"Ò¦ö–âÕF‚„vWBÕ6æ6†÷DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’w&VæFW'2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&â‚’Ð¢&WGW&â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"ÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂ6÷'BÔö&¦V7BæÖRÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòäæÖRÒ§Ð ¢2ÒÒÒÒ&6VÆ–æR89Þ8*N8;>8+ò…cRÜ*sbãb’ÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÐ ¦gVæ7F–öâvWBÔ6ö×&—6öä&6VÆ–æUö–çFW%F‚…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&¶&öö´†—7F÷'”F—"DÆæwVvREv÷&¶&öö´–B’v6ö×&—6öâÖ&6VÆ–æRæ§6öâr§Ð¦gVæ7F–öâvWBÔ6ö×&—6öä&6VÆ–æUö–çFW"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢GÒvWBÔ6ö×&—6öä&6VÆ–æUö–çFW%F‚DÆæwVvREv÷&¶&öö´–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G’’²&WGW&âFçVÆÂÐ¢G'’²&WGW&â…&VBÔ§6öäf–ÆRGFçVÆÂ’Ò6F6‚²&WGW&âFçVÆÂÐ§Ð¦gVæ7F–öâ6WBÔ6ö×&—6öä&6VÆ–æR…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒDVçdf–ævW'&–çB’°¢2Xˆ~i»þšn[¨ó¢ik6æ6†÷Bö6öçFVçB–âÓâ89Þ8*N8;>8+þ{Úîhù²Óâizw–îX˜®™šN8.˜	NKŠÞXÎjÚ.i˜.8þKùÞŠÛ~˜îZI®XN8¾X	.8ž8 ¢Gv÷&·76RÒvWBÕv÷&·76UF‚DÆæwVvP¢F†—7F÷'”Æö6²Ò¦ö–âÕF‚Gv÷&·76RvÆö6·5Æ†—7F÷'’Ö6ÆVçWæÆö6²p¢F6öçFVçDÆö6²ÒvWBÔ6öçFVçEFdÖ–çFVææ6TÆö6µF‚Gv÷&·76REv÷&¶&öö´–@¢–çfö¶RÕv—F„Æö6²F†—7F÷'”Æö6²°¢–çfö¶RÕv—F„Æö6²F6öçFVçDÆö6²°¢–b‚FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’’²F‡&÷r~jùN‹È>Yû®k©n8î[^jÛNx˜Ž8ÎŠh¾8N8¾8(®8î8¾8)>8"rÐ¢Ff–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–@¢–b‚Öæ÷B¶&ööÅÒFf–Æ&–Æ—G’ç&VG’’²F‡&÷r‚~jùN‹È>Yû®k©n8).KùÞŠÛ~8~8Þ8î8¾8)3¢r²·7G&–æuÒFf–Æ&–Æ—G’ç&V6öâ’Ð¢FöÆBÒvWBÔ6ö×&—6öä&6VÆ–æUö–çFW"DÆæwVvREv÷&¶&öö´–@¢FöÆE6æ6†÷BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBw6æ6†÷D–Brrr¢FöÆEfW'6–öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBwfW'6–öä–Brrr¢G–äFFÒ¶÷&FW&VEÔ°¢6æ6†÷D–BÒE6æ6†÷D–C²fW'6–öä–BÒEfW'6–öä–@¢&VæFW$Vçf—&öæÖVçDf–ævW'&–çBÒDVçdf–ævW'&–ç@¢f—7VÄ†6…&öf–ÆUfW'6–öâÒE67&—C¥f—7VÄ†6…&öf–ÆUfW'6–öà¢–ææVDBÒæWrÔæ÷t—6ð¢Ð¢–b‚Öæ÷B„æWrÕ6æ6†÷E–âDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–Bv6ö×&—6öâÖ&6VÆ–æRrG–äFF’’°¢F‡&÷r~jùN‹È>Yû®k©n8î[^jÛNx˜Ž8).KùÞŠÛ~8~8Þ8î8¾8)>8~8~8þ8"p¢Ð¢–b‚Öæ÷B„æWrÔ6öçFVçEFe–âGv÷&·76REv÷&¶&öö´–BEfW'6–öä–Bv6ö×&—6öâÖ&6VÆ–æRrG–äFF’’°¢&VÖ÷fRÕ6æ6†÷E–âDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–Bv6ö×&—6öâÖ&6VÆ–æRp¢F‡&÷r~jùN‹È>Yû®k©n8æ6öçFVçBDn8).KùÞŠÛ~8~8Þ8î8¾8)>8~8~8þ8"p¢Ð¢w&—FRÔ§6öäf–ÆR„vWBÔ6ö×&—6öä&6VÆ–æUö–çFW%F‚DÆæwVvREv÷&¶&öö´–B’…¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ#²6æ6†÷D–BÒE6æ6†÷D–C²fW'6–öä–BÒEfW'6–öä–@¢&VæFW$Vçf—&öæÖVçDf–ævW'&–çBÒDVçdf–ævW'&–çC²WFFVDBÒæWrÔæ÷t—6ð¢Ò¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FöÆE6æ6†÷B’ÖæBFöÆE6æ6†÷BÖæRE6æ6†÷D–B’°¢&VÖ÷fRÕ6æ6†÷E–âDÆæwVvREv÷&¶&öö´–BFöÆE6æ6†÷Bv6ö×&—6öâÖ&6VÆ–æRp¢Ð¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FöÆEfW'6–öâ’ÖæBFöÆEfW'6–öâÖæREfW'6–öä–B’°¢&VÖ÷fRÔ6öçFVçEFe–âGv÷&·76REv÷&¶&öö´–BFöÆEfW'6–öâv6ö×&—6öâÖ&6VÆ–æRp¢Ð¢Ð¢ÒÂ÷WBÔçVÆÀ§Ð ¦gVæ7F–öâvWBÔÆFW7D6ö×&—6öä76WEö–çFW%F‚…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&¶&öö´†—7F÷'”F—"DÆæwVvREv÷&¶&öö´–B’vÆFW7BÖ6ö×&—6öâÖ76WG2æ§6öâr§Ð ¦gVæ7F–öâvWBÔÆFW7D6ö×&—6öä76WEö–çFW"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢GF‚ÒvWBÔÆFW7D6ö×&—6öä76WEö–çFW%F‚DÆæwVvREv÷&¶&öö´–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’’²&WGW&âFçVÆÂÐ¢G'’²&WGW&â…&VBÔ§6öäf–ÆRGF‚FçVÆÂ’Ò6F6‚²&WGW&âFçVÆÂÐ§Ð ¦gVæ7F–öâ6WBÔÆFW7D6ö×&—6öä76WG2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂD6ö×&—6öâ’°¢2ZHži»N8988>8+Ž8ÎzK®8žy»N‹ù8îˆz®X¹^jùN‹È>8þ8jÊY¹îjùN‹È>Yû®k©n8îXˆ~i»þ8Ž8þXŠ^8¾KŠx˜Ž8).KùÞŠÛ~8ž8(¾8 ¢2Y»®Zé§&öÆ^YÞ8).XŠ^8^8¾KÛþ8n8þ8(8izv7W'&VçN8Îik&6VÆ–æ^8¾8®8(¾ZNYŽ8(.ZèžXZŽ8¾Xˆ~8(®i»þ8Ž8(ž8(Î8(¾8 ¢–b‚FçVÆÂÖWD6ö×&—6öâÖ÷"·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâw66÷Rrrr’ÖæRvWFöÖF–2rÖ÷ ¢·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâw7FGW2rrr’ÖæRv6ö×ÆWFRr’²F‡&÷r~y»N‹ùjùN‹È>8Ž8~8nKùÞŠÛ~8~8Þ8(¾ˆz®X¹^jùN‹È>{YiéÎ8Î8.8(®8î8¾8)>8"rÐ¢f÷&V6‚‚Ff–VÆB–â‚v&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâFf–VÆBrr’’’²F‡&÷r.y»N‹ùjùN‹È>8îŠÙŽXŠ^ZÙ8ÎKˆÞ‹k>8~8n8N8î8“¢Ff–VÆB"Ð¢Ð¢Gv÷&·76RÒvWBÕv÷&·76UF‚DÆæwVvP¢F†—7F÷'”Æö6²Ò¦ö–âÕF‚Gv÷&·76RvÆö6·5Æ†—7F÷'’Ö6ÆVçWæÆö6²p¢F6öçFVçDÆö6²ÒvWBÔ6öçFVçEFdÖ–çFVææ6TÆö6µF‚Gv÷&·76REv÷&¶&öö´–@¢–çfö¶RÕv—F„Æö6²F†—7F÷'”Æö6²°¢–çfö¶RÕv—F„Æö6²F6öçFVçDÆö6²°¢F&6VÆ–æTf–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvREv÷&¶&öö´–B…·7G&–æuÒD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒD6ö×&—6öâæ&6VÆ–æUfW'6–öä–B¢F7W'&VçDf–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvREv÷&¶&öö´–B…·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒD6ö×&—6öâæ7W'&VçEfW'6–öä–B¢–b‚Öæ÷B¶&ööÅÒF&6VÆ–æTf–Æ&–Æ—G’ç&VG’Ö÷"Öæ÷B¶&ööÅÒF7W'&VçDf–Æ&–Æ—G’ç&VG’’°¢F‡&÷r~y»N‹ùjùN‹È>8îyK¾X8þ88þ88>8+~8:^8ŽYÎKˆK‰nKº>8æ6öçFVçBDn8).KùÞŠÛ~8~8Þ8î8¾8)>8"p¢Ð¢FöÆBÒvWBÔÆFW7D6ö×&—6öä76WEö–çFW"DÆæwVvREv÷&¶&öö´–@¢G–äFFÒ¶÷&FW&VEÔ°¢66÷RÒvWFöÖF–2p¢&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–@¢&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æUfW'6–öä–@¢7W'&VçE6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–@¢7W'&VçEfW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçEfW'6–öä–@¢6ö×&VDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâv6ö×&VDBrrr¢–ææVDBÒæWrÔæ÷t—6ð¢Ð¢G7V72Ò€¢¶÷&FW&VEÔ²&öÆRÒv&6VÆ–æRs²–äæÖRÒvÆFW7BÖ6ö×&—6öâÖ&6VÆ–æRs²6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–C²fW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æUfW'6–öä–BÒÀ¢¶÷&FW&VEÔ²&öÆRÒv7W'&VçBs²–äæÖRÒvÆFW7BÖ6ö×&—6öâÖ7W'&VçBs²6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–C²fW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçEfW'6–öä–BÐ¢¢2˜	NKŠÞZKiY~i˜.8þX˜®™šN8¾8®KùÞŠÛ~˜îZI®XN8ŽX	.8ž8.izwö–çFW"÷–î8(.jè¾8(¾8þ8(jùN‹È>‹8~yJ>8þZK8(þ8(Î8®8N8 ¢f÷&V6‚‚G7V2–âG7V72’°¢FFFÒ¶÷&FW&VEÔ·Ð¢f÷&V6‚‚F¶W’–â‚G–äFFä¶W—2’’²FFF²F¶W•ÒÒG–äFF²F¶W•ÒÐ¢FFFç&öÆRÒ·7G&–æuÒG7V2ç&öÆP¢–b‚Öæ÷B„æWrÕ6æ6†÷E–âDÆæwVvREv÷&¶&öö´–B…·7G&–æuÒG7V2ç6æ6†÷D–B’…·7G&–æuÒG7V2ç–äæÖR’FFF’’°¢F‡&÷r~y»N‹ùjùN‹È>8î[^jÛNx˜Ž8).KùÞŠÛ~8~8Þ8î8¾8)>8~8~8þ8"p¢Ð¢–b‚Öæ÷B„æWrÔ6öçFVçEFe–âGv÷&·76REv÷&¶&öö´–B…·7G&–æuÒG7V2çfW'6–öä–B’…·7G&–æuÒG7V2ç–äæÖR’FFF’’°¢F‡&÷r~y»N‹ùjùN‹È>8æ6öçFVçBDn8).KùÞŠÛ~8~8Þ8î8¾8)>8~8~8þ8"p¢Ð¢Ð¢Gö–çFW%F‚ÒvWBÔÆFW7D6ö×&—6öä76WEö–çFW%F‚DÆæwVvREv÷&¶&öö´–@¢w&—FRÔ§6öäf–ÆRGö–çFW%F‚…¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ¢&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–@¢&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æUfW'6–öä–@¢7W'&VçE6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–@¢7W'&VçEfW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçEfW'6–öä–@¢6ö×&VDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâv6ö×&VDBrrr¢WFFVDBÒæWrÔæ÷t—6ð¢Ò¢G6fVEö–çFW"Ò&VBÔ§6öäf–ÆRGö–çFW%F‚FçVÆÀ¢–b‚FçVÆÂÖWG6fVEö–çFW"’²F‡&÷r~y»N‹ùjùN‹È>8îKùÞŠÛ~89Þ8*N8;>8+þ8).KùÞZÙŽ[èÎ8¾XhÞŠªÞ‹ëÎ8~8Þ8î8¾8)>8~8~8þ8"rÐ¢f÷&V6‚‚Ff–VÆB–â‚v&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G6fVEö–çFW"Ff–VÆBrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâFf–VÆBrr’’°¢F‡&÷r.y»N‹ùjùN‹È>8îKùÞŠÛ~89Þ8*N8;>8+þjIÎŠ‹Î8¾ZKiY~8~8î8~8ó¢Ff–VÆB ¢Ð¢Ð¢f÷&V6‚‚FöÆE7V2–â€¢¶÷&FW&VEÔ²–äæÖRÒvÆFW7BÖ6ö×&—6öâÖ&6VÆ–æRs²6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBv&6VÆ–æU6æ6†÷D–Brrr“²fW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBv&6VÆ–æUfW'6–öä–Brrr“²æWu6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–C²æWufW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æUfW'6–öä–BÒÀ¢¶÷&FW&VEÔ²–äæÖRÒvÆFW7BÖ6ö×&—6öâÖ7W'&VçBs²6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBv7W'&VçE6æ6†÷D–Brrr“²fW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBv7W'&VçEfW'6–öä–Brrr“²æWu6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–C²æWufW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçEfW'6–öä–BÐ¢’’°¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒFöÆE7V2ç6æ6†÷D–B’ÖæB·7G&–æuÒFöÆE7V2ç6æ6†÷D–BÖæR·7G&–æuÒFöÆE7V2ææWu6æ6†÷D–B’°¢&VÖ÷fRÕ6æ6†÷E–âDÆæwVvREv÷&¶&öö´–B…·7G&–æuÒFöÆE7V2ç6æ6†÷D–B’…·7G&–æuÒFöÆE7V2ç–äæÖR¢Ð¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒFöÆE7V2çfW'6–öä–B’ÖæB·7G&–æuÒFöÆE7V2çfW'6–öä–BÖæR·7G&–æuÒFöÆE7V2ææWufW'6–öä–B’°¢&VÖ÷fRÔ6öçFVçEFe–âGv÷&·76REv÷&¶&öö´–B…·7G&–æuÒFöÆE7V2çfW'6–öä–B’…·7G&–æuÒFöÆE7V2ç–äæÖR¢Ð¢Ð¢Ð¢ÒÂ÷WBÔçVÆÀ§Ð ¢2ÒÒÒÒjùN‹È>[.yJŽ8:Î8;>888:®8;>8+…cRÜ*sbãr’ÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÐ  ¦gVæ7F–öâ&VæFW"Õ6æ6†÷Df÷$6ö×&—6öâ…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢27G'V7GW&Ræ§6öîzØž8þZHži»N8~8®8N8Î8yK¾X8þ88þ88>8+~8:^8ŽYÎKˆK‰nKº>8æ6öçFVçBDn8þKùÞhÈ8ž8(¾8 ¢28>8(Î8¾8(Ž8(®XhÞ8:Î8;>888:®8;>8+jùN‹È>8~8(.8XŠNZé®Zûî‹8ŽyK¾™Ú.ŠŽzK®Zûî‹8Î[ø^8®Kˆˆ{N8ž8(¾8 ¢G'’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢G6÷W&6RÒ„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWEv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚G6÷W&6Rä6÷VçBÖwB’°¢G6÷W&6UG—RÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6÷W&6U³Òw6÷W&6UG—RrvW†6VÂr¢–b‚G6÷W&6UG—RÖWwFbr’²&WGW&â&VæFW"ÕFe6æ6†÷Df÷$6ö×&—6öâDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BÐ¢–b‚G6÷W&6UG—RÖWwv÷&Br’²&WGW&â&VæFW"Õv÷&E6æ6†÷Df÷$6ö×&—6öâDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BÐ¢–b‚G6÷W&6UG—RÖWw÷vW'ö–çBr’²&WGW&â&VæFW"Õ÷vW%ö–çE6æ6†÷Df÷$6ö×&—6öâDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BÐ¢Ð¢Ò6F6‚²Ð¢G7FFRÒvWBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢–b‚Öæ÷B¶&ööÅÒG7FFRç6÷W&6U&WF–æVB’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒw6÷W&6RÖÖ—76–ærrÒÐ¢&WGW&â–çfö¶RÕv—F…&VæFW$Æö6²DÆæwVvREv÷&¶&öö´–B°¢GfW'6–öä–BÒæWrÕ&%fW'6–öä–@¢GF×F—"Ò¦ö–âÕF‚…´”òåF…Ó£¤vWEFV×F‚‚’’‚w&"Ö6×Òr²„æWrÕ&$–B’¢Gv÷&·76RÒvWBÕv÷&·76UF‚DÆæwVvP¢F6öçFVçDF—"ÒvWBÔ6öçFVçEFefW'6–öäF—"Gv÷&·76REv÷&¶&öö´–BGfW'6–öä–@¢æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚GF×F—"Ôf÷&6RÂ÷WBÔçVÆÀ¢æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚F6öçFVçDF—"Ôf÷&6RÂ÷WBÔçVÆÀ¢FÆV6T¦ö$–BÒæWrÕ&$–@¢G6æ6†÷DÆV6RÒrs²F6öçFVçDÆV6RÒrp¢G'’°¢F†—7F÷'”Æö6²Ò¦ö–âÕF‚Gv÷&·76RvÆö6·5Æ†—7F÷'’Ö6ÆVçWæÆö6²p¢F6öçFVçDÆö6²ÒvWBÔ6öçFVçEFdÖ–çFVææ6TÆö6µF‚Gv÷&·76REv÷&¶&öö´–@¢FÆV6U&W7VÇBÒ–çfö¶RÕv—F„Æö6²F†—7F÷'”Æö6²°¢–çfö¶RÕv—F„Æö6²F6öçFVçDÆö6²°¢–b‚FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’’²F‡&÷r~jùN‹È>XX>8î[^jÛNx˜Ž8Îi[Nyn8^8(Î8î8~8þ8"rÐ¢FæWu6æ6†÷DÆV6RÒæWrÕ6æ6†÷DÆV6RDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–Bv6ö×&RrFÆV6T¦ö$–B#GfW'6–öä–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæWu6æ6†÷DÆV6R’’²F‡&÷r~[^jÛNx˜Ž8îKùÞŠÛvÆV6^8).KÙÎh‰8~8Þ8î8¾8)>8~8~8þ8"rÐ¢FæWt6öçFVçDÆV6RÒæWrÔ6öçFVçEFdÆV6RGv÷&·76REv÷&¶&öö´–BGfW'6–öä–Bv6ö×&RrFÆV6T¦ö$–B# ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæWt6öçFVçDÆV6R’’°¢&VÖ÷fRÕ6æ6†÷DÆV6RDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BFæWu6æ6†÷DÆV6P¢F‡&÷rv6öçFVçBDn8îKùÞŠÛvÆV6^8).KÙÎh‰8~8Þ8î8¾8)>8~8~8þ8"p¢Ð¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²6æ6†÷DÆV6RÒFæWu6æ6†÷DÆV6S²6öçFVçDÆV6RÒFæWt6öçFVçDÆV6RÐ¢Ð¢Ð¢G6æ6†÷DÆV6RÒ·7G&–æuÒFÆV6U&W7VÇBç6æ6†÷DÆV6P¢F6öçFVçDÆV6RÒ·7G&–æuÒFÆV6U&W7VÇBæ6öçFVçDÆV6P¢Ò6F6‚°¢&VÖ÷fRÕ6æ6†÷DÆV6RDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BG6æ6†÷DÆV6P¢&VÖ÷fRÔ6öçFVçEFdÆV6RGv÷&·76REv÷&¶&öö´–BGfW'6–öä–BF6öçFVçDÆV6P¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚F6öçFVçDF—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvÆV6RÖf–ÆVBs²ÖW76vRÒEòäW†6WF–öâäÖW76vRÐ¢Ð¢FW†6VÂÒFçVÆÃ²F&öö²ÒFçVÆÃ²G7V66W72ÒFfÇ6P¢G'’°¢Gv÷&´W‡FVç6–öâÒ´”òåF…Ó£¤vWDW‡FVç6–öâ…·7G&–æuÒG7FFRç6÷W&6UF‚’åFôÆ÷vW$–çf&–çB‚“²–b‚Gv÷&´W‡FVç6–öâÖæ÷F–â‚rç†Ç7‚rÂrç†Ç6Òr’’²Gv÷&´W‡FVç6–öâÒrç†Ç7‚rÐ¢Gv÷&²Ò¦ö–âÕF‚GF×F—"‚'6÷W&6RGv÷&´W‡FVç6–öâ"¢6÷’Ôf–ÆU6†&VE&VB…·7G&–æuÒG7FFRç6÷W&6UF‚’Gv÷&°¢G'’²Væ&Æö6²Ôf–ÆRÔÆ—FW&ÅF‚Gv÷&²ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒ6F6‚²Ð¢F6ö×&—6öå6¶vU&W&F–öâÒFçVÆÀ¢G'’²F6ö×&—6öå6¶vU&W&F–öâÒ&W&RÕ†Ç7…&–çE6¶vRGv÷&²Ò6F6‚²F6ö×&—6öå6¶vU&W&F–öâÒ¶÷&FW&VEÔ²ö³ÒFfÇ6S²&–çE6WGF–æw5&W&VCÒFfÇ6RÒÐ¢F6ö×&—6öå6¶vU&W&VBÒ¶&ööÅÒ„vWBÔFF&÷W'G’F6ö×&—6öå6¶vU&W&F–öâw&–çE6WGF–æw5&W&VBrFfÇ6R¢FW†6VÂÒæWrÔW†6VÄÆ–6F–öäf÷%&VæFW ¢FVçd–æfòÒvWBÕ&VæFW$Vçf—&öæÖVçBFW†6VÀ¢E67&—C¤7W'&VçE&VæFW$Vçdf–ævW'&–çBÒvWBÕ&VæFW$Vçf—&öæÖVçDf–ævW'&–çBFVçd–æfð¢E67&—C¤7W'&VçE&VæFW$Vçd–æfòÒFVçd–æfð¢F&öö²Ò÷VâÔW†6VÅv÷&¶&ööµ6fRFW†6VÂGv÷&²GG'VP¢G6†VWG2Ò‚¢G6†VWD6÷VçBÒ ¢G'’²G6†VWD6÷VçBÒ¶–çEÒF&öö²åv÷&·6†VWG2ä6÷VçBÒ6F6‚²G6†VWD6÷VçBÒÐ¢f÷"‚F’Ò²F’ÖÆRG6†VWD6÷VçC²F’²²’°¢Gw2ÒFçVÆÀ¢G'’°¢Gw2ÒF&öö²åv÷&·6†VWG2ä—FVÒ‚F’¢G6†VWDæÖRÒ·7G&–æuÒGw2äæÖP¢–b…¶–çEÒGw2åf—6–&ÆRÖæRÓ’²6öçF–çVRÐ¢F÷WEFbÒ¦ö–âÕF‚F6öçFVçDF—"‚'³ÒçFb"Öb„vWBÕv÷&·6†VWE7F÷&vU7FVÒG6†VWDæÖR’¢·fö–EÒ„W‡÷'BÕv÷&·6†VWEFõFe6fRFW†6VÂF&öö²Gw2F÷WEFbG6†VWDæÖRF6ö×&—6öå6¶vU&W&VB¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚F÷WEFb’°¢G6†VWG2³Ò°¢6†VWDæÖRÒG6†VWDæÖP¢FbÒF÷WEF`¢&7FW$F—&V7F÷'’Ò„vWBÕ&VæFW%&7FW%6†VWDF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BGfW'6–öä–BG6†VWDæÖR¢Ð¢Ð¢Ò6F6‚°¢Òf–æÆÇ’²–çfö¶RÔ6öÕ&VÆV6RGw2Ð¢Ð¢–b‚G6†VWG2ä6÷VçBÖW’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvæò×6†VWG2rÒÐ¢FæÇ—6—2Ò–çfö¶RÕFevTæÇ—¦W"G6†VWG0¢–b‚FçVÆÂÖWFæÇ—6—2Ö÷"Öæ÷B¶&ööÅÒFæÇ—6—2æö²’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvæÇ—¦RÖf–ÆVBrÒÐ¢w&—FRÕ&VæFW%&V6÷&BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BGfW'6–öä–Bv6ö×&—6öârGG'VR…·67W7FöÖö&¦V7EÒFæÇ—6—2ç&W7VÇB¢Ff–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BGfW'6–öä–@¢–b‚Öæ÷B¶&ööÅÒFf–Æ&–Æ—G’ç&VG’’²&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒw&WFVçF–öâ×fW&–g’Öf–ÆVBs²ÖW76vRÒ·7G&–æuÒFf–Æ&–Æ—G’ç&V6öâÒÐ¢G7V66W72ÒGG'VP¢&WGW&â¶÷&FW&VEÔ²ö²ÒGG'VS²fW'6–öä–BÒGfW'6–öä–C²Vçdf–ævW'&–çBÒ·7G&–æuÒE67&—C¤7W'&VçE&VæFW$Vçdf–ævW'&–çBÐ¢Ò6F6‚°¢&WGW&â¶÷&FW&VEÔ²ö²ÒFfÇ6S²&V6öâÒvW'&÷"s²ÖW76vRÒEòäW†6WF–öâäÖW76vRÐ¢Òf–æÆÇ’°¢–b‚F&öö²’²G'’²F&öö²ä6Æ÷6R‚FfÇ6R’Ò6F6‚²Ò²–çfö¶RÔ6öÕ&VÆV6RF&öö²Ð¢–b‚FW†6VÂ’²6Æ÷6RÔW†6VÄÆ–6F–öäf÷%&VæFW"FW†6VÂÐ¢&VÖ÷fRÕ6æ6†÷DÆV6RDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BG6æ6†÷DÆV6P¢&VÖ÷fRÔ6öçFVçEFdÆV6RGv÷&·76REv÷&¶&öö´–BGfW'6–öä–BF6öçFVçDÆV6P¢–b‚Öæ÷BG7V66W72ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚F6öçFVçDF—"’’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚F6öçFVçDF—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF×F—"’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF×F—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢´t5Ó£¤6öÆÆV7B‚“²´t5Ó£¥v—Df÷%VæF–ætf–æÆ—¦W'2‚¢Ð¢Ð§Ð ¦gVæ7F–öâvWBÔ6ö×&—6öå6÷W&6UG—R…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢G'’°¢F6öçFW‡BÒvWBÕ&Vv—7FW&VE6÷W&6TFFW$6öçFW‡BDÆæwVvREv÷&¶&öö´–@¢G6÷W&6UG—RÒ…·7G&–æuÒ„vWBÔFF&÷W'G’F6öçFW‡Bw6÷W&6UG—RrvW†6VÂr’’åG&–Ò‚’åFôÆ÷vW$–çf&–çB‚¢–b‚G6÷W&6UG—RÖ–â‚vW†6VÂrÂwv÷&BrÂwFbrÂw÷vW'ö–çBr’’²&WGW&âG6÷W&6UG—RÐ¢Ò6F6‚²Ð¢&WGW&âvW†6VÂp§Ð ¦gVæ7F–öâFW7BÔ6ö×&—6öå6†VWDWV—fÆVçB‚D&Vf÷&RÂDgFW"’°¢–b‚FçVÆÂÖWD&Vf÷&RÖ÷"FçVÆÂÖWDgFW"’²&WGW&âFfÇ6RÐ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’D&Vf÷&Rw7FGW2rrr’ÖæRvö²rÖ÷"·7G&–æuÒ„vWBÔFF&÷W'G’DgFW"w7FGW2rrr’ÖæRvö²r’²&WGW&âFfÇ6RÐ¢F&Vf÷&T†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’D&Vf÷&Rw6†VWEf—7VÄ†6‚rrr’¢FgFW$†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’DgFW"w6†VWEf—7VÄ†6‚rrr’¢&WGW&â‚‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&Vf÷&T†6‚’ÖæBF&Vf÷&T†6‚ÖWFgFW$†6‚’Ö÷"…FW7BÕ6†VWEf—7VÄWV—fÆVçBD&Vf÷&RDgFW"’§Ð ¦gVæ7F–öâæWrÔ6ö×&—6öåVæ—DÖ–ær‚D&Vf÷&RÂDgFW"Â·7G&–æuÒD¶–æBÂ¶F÷V&ÆUÒD6öæf–FVæ6RÂ·7G&–æuÒDÖWF†öBÂ·7G&–æuÒDÖW76vRÒrr’°¢F&Vf÷&TæÖRÒB†–b‚FçVÆÂÖæRD&Vf÷&R’²·7G&–æuÒ„vWBÔFF&÷W'G’D&Vf÷&Rw6†VWDæÖRrrr’ÒVÇ6R²rrÒ¢FgFW$æÖRÒB†–b‚FçVÆÂÖæRDgFW"’²·7G&–æuÒ„vWBÔFF&÷W'G’DgFW"w6†VWDæÖRrrr’ÒVÇ6R²rrÒ¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢&Vf÷&U6†VWDæÖRÒF&Vf÷&TæÖP¢gFW%6†VWDæÖRÒFgFW$æÖP¢F—7Æ”æÖRÒB†–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FgFW$æÖR’’²FgFW$æÖRÒVÇ6R²F&Vf÷&TæÖRÒ¢¶–æBÒD¶–æ@¢ÖF6„6öæf–FVæ6RÒ´ÖF…Ó£¤Ö‚ƒÂ´ÖF…Ó£¤Ö–âƒÂD6öæf–FVæ6R’¢ÖF6„ÖWF†öBÒDÖWF†ö@¢ÖW76vRÒDÖW76vP¢Ð§Ð ¦gVæ7F–öâvWBÔ6ö×&—6öå6†VWE6WVVæ6TçVÖ&W"‚E6†VWB’°¢FæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’E6†VWBw6†VWDæÖRrrr¢–b‚FæÖRÖÖF6‚rƒóÆçVÖ&W#åÆB²•Ç2¢Br’²&WGW&â¶–çEÒFÖF6†W5²vçVÖ&W"uÒÐ¢&WGW&âvWBÕ6†VWD÷&FW$çVÖ&W"FæÖP§Ð ¦gVæ7F–öâvWBÔ6ö×&—6öåVæ—DÖ–æw2‚D&6VÆ–æU6†VWG2ÂD7W'&VçE6†VWG2Â·7G&–æuÒE6÷W&6UG—RÒvW†6VÂr’°¢F&Vf÷&RÒ„vWBÔ'&’D&6VÆ–æU6†VWG2¢FgFW"Ò„vWBÔ'&’D7W'&VçE6†VWG2¢–b‚E6÷W&6UG—RÖWvW†6VÂr’°¢F&Vf÷&TÖÒ·Ð¢f÷&V6‚‚G6†VWB–âF&Vf÷&R’°¢FæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBw6†VWDæÖRrrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’’²F&Vf÷&TÖ²FæÖUÒÒG6†VWBÐ¢Ð¢G6VVâÒ·Ó²G&W7VÇBÒ‚¢f÷&V6‚‚G6†VWB–âFgFW"’°¢FæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBw6†VWDæÖRrrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’’²6öçF–çVRÐ¢G6VVå²FæÖUÒÒGG'VP¢–b‚Öæ÷BF&Vf÷&TÖä6öçF–ç4¶W’‚FæÖR’’²G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFçVÆÂG6†VWBvFFVBrvæÖR×VæÖF6†VBs²6öçF–çVRÐ¢FöÆBÒF&Vf÷&TÖ²FæÖUÐ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBw7FGW2rrr’ÖæRvö²rÖ÷"·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBw7FGW2rrr’ÖæRvö²r’°¢G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFöÆBG6†VWBwVæ¶æ÷vârvæÇ—6—2×Væf–Æ&ÆRr~jùN‹È>yJŽyK¾X8þ8).Šz>ié8~8Þ8®8N8þ8(XŠNZé®KˆÞˆ;Þ8~8ž8"p¢ÒVÇ6V–b…FW7BÔ6ö×&—6öå6†VWDWV—fÆVçBFöÆBG6†VWB’°¢G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFöÆBG6†VWBwVæ6†ævVBrw7F&ÆRÖæÖRÖæBÖ†6‚p¢ÒVÇ6R°¢G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFöÆBG6†VWBvÖöF–f–VBrw7F&ÆRÖæÖRp¢Ð¢Ð¢f÷&V6‚‚FæÖR–â‚F&Vf÷&TÖä¶W—2’’°¢–b‚Öæ÷BG6VVâä6öçF–ç4¶W’…·7G&–æuÒFæÖR’’²G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærF&Vf÷&TÖ²FæÖUÒFçVÆÂw&VÖ÷fVBrvæÖR×VæÖF6†VBrÐ¢Ð¢&WGW&â‚G&W7VÇB¢Ð ¢2v÷&BõDn8þxšžyn89®8;Î8+ŽYÞ8ÎKØÞ{Úî8¾KéÞZÙŽ8ž8(¾8.ZèÎXZŽKˆˆ{N89®8;Î8+Ž8îiÈ™[~X[˜	®˜:ŽXˆnX‰~8) ¢28*.8;>8*¾8;Î8¾8~8˜	NKŠÞ8Ž8îhËþXZ^8;¾X˜®™šN8~[èÎ{i®89®8;Î8+Ž8).ŠªN8>8n8ÎZHži»N8Þ8¾8~8®8N8 ¢F&Vf÷&RÒ‚F&Vf÷&RÂ6÷'BÔö&¦V7B´W‡&W76–öã×´vWBÔ6ö×&—6öå6†VWE6WVVæ6TçVÖ&W"E÷Ó´66VæF–æsÒGG'VWÒ¢FgFW"Ò‚FgFW"Â6÷'BÔö&¦V7B´W‡&W76–öã×´vWBÔ6ö×&—6öå6†VWE6WVVæ6TçVÖ&W"E÷Ó´66VæF–æsÒGG'VWÒ¢FÒÒF&Vf÷&Rä6÷VçC²FâÒFgFW"ä6÷Vç@¢–b‚FÒÖW’²&WGW&â‚FgFW"Âf÷$V6‚Ôö&¦V7B²æWrÔ6ö×&—6öåVæ—DÖ–ærFçVÆÂEòvFFVBrw6WVVæ6RÖ–ç6W'FVBrÒ’Ð¢–b‚FâÖW’²&WGW&â‚F&Vf÷&RÂf÷$V6‚Ôö&¦V7B²æWrÔ6ö×&—6öåVæ—DÖ–ærEòFçVÆÂw&VÖ÷fVBrw6WVVæ6R×&VÖ÷fVBrÒ’Ð¢FWV—fÆVçBÒæWrÔö&¦V7Bv&ööÅ²ÅÒrFÒÂFà¢F†6„6÷VçG4&Vf÷&RÒ·Ó²F†6„6÷VçG4gFW"Ò·Ð¢f÷&V6‚‚G6†VWB–âF&Vf÷&R’²F‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBw6†VWEf—7VÄ†6‚rrr’“²–b‚F‚’²F†6„6÷VçG4&Vf÷&U²F…ÒÒ²¶–çEÒ„vWBÔFF&÷W'G’F†6„6÷VçG4&Vf÷&RF‚’ÒÐ¢f÷&V6‚‚G6†VWB–âFgFW"’²F‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBw6†VWEf—7VÄ†6‚rrr’“²–b‚F‚’²F†6„6÷VçG4gFW%²F…ÒÒ²¶–çEÒ„vWBÔFF&÷W'G’F†6„6÷VçG4gFW"F‚’ÒÐ¢f÷"‚F’Ò²F’ÖÇBFÓ²F’²²’²f÷"‚F¢Ò²F¢ÖÇBFã²F¢²²’²FWV—fÆVçE²F’ÂF¥ÒÒFW7BÔ6ö×&—6öå6†VWDWV—fÆVçB‚F&Vf÷&U²F•Ò’‚FgFW%²F¥Ò’ÒÐ¢FÆ72ÒæWrÔö&¦V7Bv–çE²ÅÒr‚FÒ²’Â‚Fâ²¢f÷"‚F’ÒFÒÒ²F’ÖvR²F’ÒÒ’°¢f÷"‚F¢ÒFâÒ²F¢ÖvR²F¢ÒÒ’°¢FæW‡D’ÒF’²²FæW‡D¢ÒF¢²¢–b‚FWV—fÆVçE²F’ÂF¥Ò’²FÆ75²F’ÂF¥ÒÒ²FÆ75²FæW‡D’ÂFæW‡D¥ÒÐ¢VÇ6R°¢G6¶—&Vf÷&RÒ¶–çEÒFÆ75²FæW‡D’ÂF¥Ó²G6¶—gFW"Ò¶–çEÒFÆ75²F’ÂFæW‡D¥Ð¢FÆ75²F’ÂF¥ÒÒ´ÖF…Ó£¤Ö‚‚G6¶—&Vf÷&RÂG6¶—gFW"¢Ð¢Ð¢Ð¢Fæ6†÷'2Ò‚“²F’Ò²F¢Ò ¢v†–ÆR‚F’ÖÇBFÒÖæBF¢ÖÇBFâ’°¢FæW‡D’ÒF’²²FæW‡D¢ÒF¢²¢–b‚FWV—fÆVçE²F’ÂF¥ÒÖæBFÆ75²F’ÂF¥ÒÖWƒ²FÆ75²FæW‡D’ÂFæW‡D¥Ò’’²Fæ6†÷'2³Ò·67W7FöÖö&¦V7EÔ²&Vf÷&SÒF“²gFW#ÒF¢Ó²F’²³²F¢²²Ð¢VÇ6R°¢G6¶—&Vf÷&RÒ¶–çEÒFÆ75²FæW‡D’ÂF¥Ó²G6¶—gFW"Ò¶–çEÒFÆ75²F’ÂFæW‡D¥Ð¢–b‚G6¶—&Vf÷&RÖvRG6¶—gFW"’²F’²²ÒVÇ6R²F¢²²Ð¢Ð¢Ð¢Fæ6†÷'2³Ò·67W7FöÖö&¦V7EÔ²&Vf÷&SÒFÓ²gFW#ÒFâÐ¢G&W7VÇBÒ‚“²F&Vf÷&U7F'BÒ²FgFW%7F'BÒ ¢f÷&V6‚‚Fæ6†÷"–âFæ6†÷'2’°¢F&Vf÷&TVæBÒ¶–çEÒFæ6†÷"æ&Vf÷&S²FgFW$VæBÒ¶–çEÒFæ6†÷"ægFW ¢F&Vf÷&TvÒF&Vf÷&TVæBÒF&Vf÷&U7F'C²FgFW$vÒFgFW$VæBÒFgFW%7F'@¢–b‚F&Vf÷&TvÖW’°¢f÷"‚F²Ò²F²ÖÇBFgFW$v²F²²²’²G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFçVÆÂ‚FgFW%²FgFW%7F'B²FµÒ’vFFVBrw6WVVæ6RÖ–ç6W'FVBrÐ¢ÒVÇ6V–b‚FgFW$vÖW’°¢f÷"‚F²Ò²F²ÖÇBF&Vf÷&Tv²F²²²’²G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ær‚F&Vf÷&U²F&Vf÷&U7F'B²FµÒ’FçVÆÂw&VÖ÷fVBrw6WVVæ6R×&VÖ÷fVBrÐ¢ÒVÇ6V–b‚F&Vf÷&TvÖWFgFW$v’°¢f÷"‚F²Ò²F²ÖÇBF&Vf÷&Tv²F²²²’°¢FöÆBÒF&Vf÷&U²F&Vf÷&U7F'B²FµÓ²F7W"ÒFgFW%²FgFW%7F'B²FµÐ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBw7FGW2rrr’ÖæRvö²rÖ÷"·7G&–æuÒ„vWBÔFF&÷W'G’F7W"w7FGW2rrr’ÖæRvö²r’°¢G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFöÆBF7W"wVæ¶æ÷vârvæÇ—6—2×Væf–Æ&ÆRr~jùN‹È>yJŽyK¾X8þ8).Šz>ié8~8Þ8®8N8þ8(XŠNZé®KˆÞˆ;Þ8~8ž8"p¢ÒVÇ6R°¢F6öæf–FVæ6RÒB†–b‚F&Vf÷&TvÖW’²ã“RÒVÇ6R²ã‚Ò¢G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFöÆBF7W"vÖöF–f–VBrF6öæf–FVæ6Rw6WVVæ6RÖ&WGvVVâÖæ6†÷'2r~X˜Þ[èÎ8îKˆˆ{N89®8;Î8+Ž8).Yû®k©n8¾Zûî[ùÎK¹Ž88î8~8þ8"p¢Ð¢Ð¢ÒVÇ6R°¢G—'2Ò´ÖF…Ó£¤Ö–â‚F&Vf÷&TvÂFgFW$v¢f÷"‚F²Ò²F²ÖÇBG—'3²F²²²’°¢G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ær‚F&Vf÷&U²F&Vf÷&U7F'B²FµÒ’‚FgFW%²FgFW%7F'B²FµÒ’wVæ¶æ÷vârã3RvÖ&–wV÷W2×6WVVæ6Rr~89®8;Î8+Ž8î‹ûÞXª8;¾X˜®™šN8ŽZHži»N8ÎYÎ8ŽXË®™i>8¾8.8(®8Zûî[ùÎ8).z+®Zé®8~8Þ8î8¾8)>8"p¢Ð¢f÷"‚F²ÒG—'3²F²ÖÇBF&Vf÷&Tv²F²²²’²G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ær‚F&Vf÷&U²F&Vf÷&U7F'B²FµÒ’FçVÆÂw&VÖ÷fVBrãcRvÖ&–wV÷W2×6WVVæ6R×VæÖF6†VBrÐ¢f÷"‚F²ÒG—'3²F²ÖÇBFgFW$v²F²²²’²G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFçVÆÂ‚FgFW%²FgFW%7F'B²FµÒ’vFFVBrãcRvÖ&–wV÷W2×6WVVæ6R×VæÖF6†VBrÐ¢Ð¢–b‚F&Vf÷&TVæBÖÇBFÒÖæBFgFW$VæBÖÇBFâ’°¢FöÆBÒF&Vf÷&U²F&Vf÷&TVæEÓ²F7W"ÒFgFW%²FgFW$VæEÐ¢F†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBw6†VWEf—7VÄ†6‚rrr’¢FÖ&–wV÷W4†6‚Ò‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F†6‚’ÖæB…¶–çEÒ„vWBÔFF&÷W'G’F†6„6÷VçG4&Vf÷&RF†6‚’ÖwBÖ÷"¶–çEÒ„vWBÔFF&÷W'G’F†6„6÷VçG4gFW"F†6‚’ÖwB’¢G&W7VÇB³ÒæWrÔ6ö×&—6öåVæ—DÖ–ærFöÆBF7W"wVæ6†ævVBrB†–b‚FÖ&–wV÷W4†6‚’²ãƒRÒVÇ6R²ãÒ’B†–b‚FÖ&–wV÷W4†6‚’²vW†7BÖ†6‚×6WVVæ6RÖGWÆ–6FRrÒVÇ6R²vW†7BÖ†6‚×6WVVæ6RrÒ¢Ð¢F&Vf÷&U7F'BÒF&Vf÷&TVæB²²FgFW%7F'BÒFgFW$VæB²¢Ð¢&WGW&â‚G&W7VÇB§Ð ¦gVæ7F–öâ6WBÔ6ö×&—6öåVæ—DÖ–æu&W7VÇB‚E&W7VÇBÂDÖ–æw2’°¢FÆ—7BÒ„vWBÔ'&’DÖ–æw2¢E&W7VÇBçVæ—DÖ–æw2Ò‚FÆ—7B¢E&W7VÇBæ6†ævVE6†VWG2Ò‚FÆ—7BÂv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWvÖöF–f–VBrÒÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòæF—7Æ”æÖRÒ¢E&W7VÇBçVæ6†ævVE6†VWG2Ò‚FÆ—7BÂv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWwVæ6†ævVBrÒÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòæF—7Æ”æÖRÒ¢E&W7VÇBçVæ¶æ÷vå6†VWG2Ò‚FÆ—7BÂv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWwVæ¶æ÷vârÒÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòæF—7Æ”æÖRÒ¢E&W7VÇBæFFVE6†VWG2Ò‚FÆ—7BÂv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWvFFVBrÒÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòæF—7Æ”æÖRÒ¢E&W7VÇBç&VÖ÷fVE6†VWG2Ò‚FÆ—7BÂv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWw&VÖ÷fVBrÒÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòæF—7Æ”æÖRÒ¢F6öæf–FVæ6W2Ò‚FÆ—7BÂf÷$V6‚Ôö&¦V7B²¶F÷V&ÆUÒ„vWBÔFF&÷W'G’EòvÖF6„6öæf–FVæ6Rr’Ò¢E&W7VÇBæÖ–æt6öæf–FVæ6RÒB†–b‚F6öæf–FVæ6W2ä6÷VçBÖwB’²¶F÷V&ÆUÒ„‚F6öæf–FVæ6W2ÂÖV7W&RÔö&¦V7BÔÖ–æ–×VÒ•³ÒäÖ–æ–×VÒ’ÒVÇ6R²ãÒ§Ð ¦gVæ7F–öâ6ö×&RÕ6æ6†÷Ef—7VÂ…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒD7W'&VçE6æ6†÷D–BÂ·7G&–æuÒD7W'&VçEfW'6–öä–B’°¢2ŠŽxûî8þ8ÎŠh¾‰Þ8Ž8~8®8~8Þ8~8þ8®8þ8ÎhùX{®yJ…Dn8îŠh¾8þyºî8).Yû®k©n8Ž8~8þš¹Ž{+î[ªn8®XŠNZé®8Þ8 ¢G&W7VÇBÒ¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ3²7FGW2ÒwVæf–Æ&ÆRp¢66÷RÒvWFöÖF–2p¢&6VÆ–æU6æ6†÷D–BÒrs²&6VÆ–æUfW'6–öä–BÒrp¢7W'&VçE6æ6†÷D–BÒD7W'&VçE6æ6†÷D–C²7W'&VçEfW'6–öä–BÒD7W'&VçEfW'6–öä–@¢6ö×&VDBÒæWrÔæ÷t—6ó²ÖWF†öBÒrp¢6÷W&6UG—RÒvWBÔ6ö×&—6öå6÷W&6UG—RDÆæwVvREv÷&¶&öö´–@¢Væ—DÖ–æw2Ò‚“²Ö–æt6öæf–FVæ6RÒã ¢6†ævVE6†VWG2Ò‚“²Væ6†ævVE6†VWG2Ò‚“²Væ¶æ÷vå6†VWG2Ò‚¢FFVE6†VWG2Ò‚“²&VÖ÷fVE6†VWG2Ò‚“²ÖW76vRÒrp¢Ð¢F7W"ÒvWBÕf—7VÄ†6†W2DÆæwVvREv÷&¶&öö´–BD7W'&VçE6æ6†÷D–BD7W'&VçEfW'6–öä–@¢–b‚FçVÆÂÖWF7W"’²G&W7VÇBæÖW76vRÒ~K¸®Y¹îx˜Ž8îyK¾X8þ88þ88>8+~8:^8Î8.8(®8î8¾8)>8"s²&WGW&âG&W7VÇBÐ¢F7W$VçbÒ·7G&–æuÒ„vWBÔFF&÷W'G’F7W"w&VæFW$Vçf—&öæÖVçDf–ævW'&–çBrrr ¢GG"ÒvWBÔ6ö×&—6öä&6VÆ–æUö–çFW"DÆæwVvREv÷&¶&öö´–@¢F&6U6æÒ·7G&–æuÒ„vWBÔFF&÷W'G’GG"w6æ6†÷D–Brrr¢F&6UfW"Ò·7G&–æuÒ„vWBÔFF&÷W'G’GG"wfW'6–öä–Brrr¢2F†R&6VÆ–æRö–çFW"—2WF†÷&—FF—fRâfÆÆ–ær&6²FòÖæ–fW7Bç&Wf–÷W56æ6†÷D–@¢2—2Vç6fRgFW"&Vv—7G&F–öâ&V6÷fW'’÷"†6‚FRÖGWÆ–6F–öã¢F†R'&Wf–÷W2 ¢26æ6†÷B6â&RÖöçF‡2öÆBWfVâF†÷Vv‚F†RW6W"§W7B&V7&VFVBDbà¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6U6æ’Ö÷"F&6U6æÖWD7W'&VçE6æ6†÷D–B’°¢G&W7VÇBæÖW76vRÒ~X˜ÞY¹î8îjùN‹È>Yû®k©n8Î8.8(®8î8¾8)>8.K¸®Y¹îx˜Ž8).ik8~8NYû®k©n8¾8~8î8ž8"p¢&WGW&âG&W7VÇ@¢Ð ¢2ˆz®X¹^jùN‹È>8(.[^jÛNjùN‹È>8ŽYÎ8Ž8þ8XŠNZé®88þ88>8+~8:^8ŽŠŽzK¥Dn8îK‰nKº>Kˆˆ{N8).[ø^šŽ8¾8ž8(¾8 ¢F7W'&VçDf–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvREv÷&¶&öö´–BD7W'&VçE6æ6†÷D–BD7W'&VçEfW'6–öä–@¢–b‚Öæ÷B¶&ööÅÒF7W'&VçDf–Æ&–Æ—G’ç&VG’’°¢G&W7VÇBæÖW76vRÒ~K¸®Y¹îx˜Ž8îyK¾X8þ88þ88>8+~8:^8ŽYÎKˆK‰nKº>8æ6öçFVçBDn8Î8Þ8(Þ8>8n8N8®8N8þ8(8ˆz®X¹^jùN‹È>8~8Þ8î8¾8)>8"p¢&WGW&âG&W7VÇ@¢Ð¢F&6RÒFçVÆÀ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6UfW"’’²F&6RÒvWBÕf—7VÄ†6†W2DÆæwVvREv÷&¶&öö´–BF&6U6æF&6UfW"Ð¢F&6Tf–Æ&–Æ—G’ÒFçVÆÀ¢–b‚FçVÆÂÖæRF&6RÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6UfW"’’°¢F&6Tf–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvREv÷&¶&öö´–BF&6U6æF&6UfW ¢Ð¢F7W%&öf–ÆRÒvWBÔFF&÷W'G’F7W"wf—7VÄ†6…&öf–ÆRrFçVÆÀ¢F&6U&öf–ÆRÒB†–b‚FçVÆÂÖæRF&6R’²vWBÔFF&÷W'G’F&6Rwf—7VÄ†6…&öf–ÆRrFçVÆÂÒVÇ6R²FçVÆÂÒ¢FVçf—&öæÖVçD6†ævVBÒ‚FçVÆÂÖæRF&6RÖæB€¢·7G&–æuÒ„vWBÔFF&÷W'G’F&6Rw&VæFW$Vçf—&öæÖVçDf–ævW'&–çBrrr’ÖæRF7W$VçbÖ÷ ¢„vWBÔ–çDFF&÷W'G’F&6RvæÇ—¦W%fW'6–öâr’ÖæR„vWBÔ–çDFF&÷W'G’F7W"væÇ—¦W%fW'6–öâr’Ö÷ ¢„vWBÔ–çDFF&÷W'G’F&6U&öf–ÆRw&öf–ÆUfW'6–öâr’ÖæR„vWBÔ–çDFF&÷W'G’F7W%&öf–ÆRw&öf–ÆUfW'6–öâr’’¢F76WG4Ö—76–ærÒ‚FçVÆÂÖWF&6Tf–Æ&–Æ—G’Ö÷"Öæ÷B¶&ööÅÒF&6Tf–Æ&–Æ—G’ç&VG’¢FÖWF†öBÒw7F÷&VBÖ†6‚p¢–b‚FçVÆÂÖWF&6RÖ÷"FVçf—&öæÖVçD6†ævVBÖ÷"F76WG4Ö—76–ær’°¢2y+Z(>[zî8;¾88þ88>8+~8:^jÊ‰Þ8;¾YÎKˆK‰nKº5DnjÊ‰Þ8î8N8®8(Î8~8(.8KùÞZÙŽkˆŽ8ôW†6VÎ8¾8(žKˆ{XN8).XhÞyIþh‰8ž8(¾8 ¢G&RÒ&VæFW"Õ6æ6†÷Df÷$6ö×&—6öâDÆæwVvREv÷&¶&öö´–BF&6U6æ ¢–b…¶&ööÅÒG&Ræö²’°¢F&6UfW"Ò·7G&–æuÒG&RçfW'6–öä–@¢F&6RÒvWBÕf—7VÄ†6†W2DÆæwVvREv÷&¶&öö´–BF&6U6æF&6UfW ¢F&6Tf–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvREv÷&¶&öö´–BF&6U6æF&6UfW ¢FÖWF†öBÒw&R×&VæFW&VBp¢ÒVÇ6R°¢G&W7VÇBæÖW76vRÒ~X˜ÞY¹îx˜Ž8îjùN‹È>‹8~yJ>8).YÎKˆK‰nKº>8~z+®KùÞ8~8Þ8®8N8þ8(8ˆz®X¹^jùN‹È>8~8Þ8î8¾8)>8.K¸®Y¹îx˜Ž8).ik8~8NjùN‹È>Yû®k©n8Ž8~8î8ž8"p¢&WGW&âG&W7VÇ@¢Ð¢Ð¢–b‚FçVÆÂÖWF&6RÖ÷"FçVÆÂÖWF&6Tf–Æ&–Æ—G’Ö÷"Öæ÷B¶&ööÅÒF&6Tf–Æ&–Æ—G’ç&VG’’°¢G&W7VÇBæÖW76vRÒ~X˜ÞY¹îx˜Ž8îyK¾X8þ88þ88>8+~8:^8ŽYÎKˆK‰nKº>8æ6öçFVçBDn8).z+®Š¨Þ8~8Þ8î8¾8)>8"p¢&WGW&âG&W7VÇ@¢Ð ¢FÖ–æw2ÒvWBÔ6ö×&—6öåVæ—DÖ–æw2„vWBÔFF&÷W'G’F&6Rw6†VWG2r‚’’„vWBÔFF&÷W'G’F7W"w6†VWG2r‚’’…·7G&–æuÒG&W7VÇBç6÷W&6UG—R¢G&W7VÇBç7FGW2Òv6ö×ÆWFRp¢G&W7VÇBæ&6VÆ–æU6æ6†÷D–BÒF&6U6æ ¢G&W7VÇBæ&6VÆ–æUfW'6–öä–BÒF&6UfW ¢G&W7VÇBæÖWF†öBÒFÖWF†ö@¢6WBÔ6ö×&—6öåVæ—DÖ–æu&W7VÇBG&W7VÇBFÖ–æw0¢FF—"Ò¦ö–âÕF‚„vWBÕ&VæFW%&V6÷&DF—"DÆæwVvREv÷&¶&öö´–BD7W'&VçE6æ6†÷D–BD7W'&VçEfW'6–öä–B’v6ö×&—6öç2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢F6ö×&—6öä¶W’Ò„vWBÕ6†#SeFW‡B‚'³×Ç³×Ç³'×Ç³7×ÆWFöÖF–2"ÖbF&6U6æÂF&6UfW"ÂD7W'&VçE6æ6†÷D–BÂD7W'&VçEfW'6–öä–B’’å7V'7G&–ærƒrÂb¢F6ö×&—6öåF‚Ò¦ö–âÕF‚FF—"‚&6××³Òæ§6öâ"ÖbF6ö×&—6öä¶W’¢w&—FRÔ§6öäf–ÆRF6ö×&—6öåF‚G&W7VÇ@¢G6fVBÒ&VBÔ§6öäf–ÆRF6ö×&—6öåF‚FçVÆÀ¢–b‚FçVÆÂÖWG6fVB’²F‡&÷r~ˆz®X¹^jùN‹È>{YiéÎ8).KùÞZÙŽ[èÎ8¾XhÞŠªÞ‹ëÎ8~8Þ8î8¾8)>8~8~8þ8"rÐ¢f÷&V6‚‚Ff–VÆB–â‚w66÷RrÂv&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G6fVBFf–VÆBrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’G&W7VÇBFf–VÆBrr’’°¢F‡&÷r.ˆz®X¹^jùN‹È>{YiéÎ8îKùÞZÙŽjIÎŠ‹Î8¾ZKiY~8~8î8~8ó¢Ff–VÆB ¢Ð¢Ð¢G'’°¢6WBÔÆFW7D6ö×&—6öä76WG2DÆæwVvREv÷&¶&öö´–BG6fV@¢Ò6F6‚°¢2KùÞŠÛ~8~8Þ8®8NjùN‹È>8).iÈikZHži»N8988>8+Ž8ŽXZÎ™h¾8~8®8N8 ¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚F6ö×&—6öåF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢F‡&÷p¢Ð¢G'’²V&Æ—6‚ÔÆFW7D6ö×&—6öä66†W2DÆæwVvREv÷&¶&öö´–BG6fVBF7W"F&6RÒ6F6‚°¢w&—FRÕv&æ–ær‚~jùN‹È>h8^Z8î8:Þ8;Î8*¾8:¾8*Þ8:>88>8+~8:^KÙÎh‰8¾ZKiY~8~8î8~8ó¢r²EòäW†6WF–öâäÖW76vR¢Ð¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRv6ö×&Ræ6ö×ÆWFVBr…¶÷&FW&VEÔ²v÷&¶&öö´–BÒEv÷&¶&öö´–C²6æ6†÷D–BÒD7W'&VçE6æ6†÷D–C²6†ævVBÒ‚G&W7VÇBæ6†ævVE6†VWG2“²Væ¶æ÷vâÒ‚G&W7VÇBçVæ¶æ÷vå6†VWG2’Ò¢&WGW&âG6fV@§Ð ¦gVæ7F–öâ–çfö¶RÕ÷7E&VæFW$æÇ—6—2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEfW'6–öä–BÂE&VæFW&VB’°¢2cRÜ*sbãS¢DnKÙÎh‰8î8*þ8:®88n8*>8*¾8:¾898+ž8îZIn8.ZKiY~8~8n8("DbKÙÎh‰8þh‰X©þh›8N8î8î8î8 ¢–b‚Öæ÷B…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’’²&WGW&âFçVÆÂÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6æ6†÷D–B’’²&WGW&âFçVÆÂÐ¢G'’°¢G6†VWG2Ò‚¢f÷&V6‚‚G"–â„vWBÔ'&’E&VæFW&VB’’°¢GFbÒ·7G&–æuÒ„vWBÔFF&÷W'G’G"wFbrrr¢FæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G"w6†VWDæÖRrrr¢2&VæFW"Õv÷&¶&öö²Ç&VG’fÆ–FFVBWfW'’÷WGWBDb&Vf÷&RVWV–ærF†—2æÇ—6—2à¢2fö–BöæRW‡G&4Ô"7FBW"6†VWBöâF†R6ö×&—6öâ&W&F–öâF‚à¢–b‚GFbÖæBFæÖR’°¢G6†VWG2³Ò°¢6†VWDæÖRÒFæÖP¢FbÒGF`¢&7FW$F—&V7F÷'’Ò„vWBÕ&VæFW%&7FW%6†VWDF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–BFæÖR¢Ð¢Ð¢Ð¢–b‚G6†VWG2ä6÷VçBÖW’²&WGW&âFçVÆÂÐ¢FæÇ—6—2Ò–çfö¶RÕFevTæÇ—¦W"G6†VWG0¢G'6VBÒFçVÆÀ¢–b‚FçVÆÂÖæRFæÇ—6—2ÖæB¶&ööÅÒFæÇ—6—2æö²’²G'6VBÒ·67W7FöÖö&¦V7EÒFæÇ—6—2ç&W7VÇBÐ¢w&—FRÕ&VæFW%&V6÷&BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–Bvæ÷&ÖÂrGG'VRG'6V@¢–b‚FçVÆÂÖWG'6VB’²&WGW&âFçVÆÂÐ¢F6×Ò6ö×&RÕ6æ6†÷Ef—7VÂDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–@¢26ö×ÆWFR÷Væf–Æ&Æ^8î8ž88(ž8~8(.Šz>ié{YiéÎ8).x˜Ž8îŠ‰Ž˜Ë.8Žjè¾8ž8 ¢2™ÙîYÎiÉþŠz>ié8îz»nYŽ8(Ny+Z(>[zî8~jùN‹È>8~8Þ8®8NZNYŽ8¾8ynyK8).[èÎ8¾8(žz+®Š¨Þ8~8Þ8(¾8 ¢G'’°¢FæÇ—6—5&V6÷&DF—"ÒvWBÕ&VæFW%&V6÷&DF—"DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–@¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚FæÇ—6—5&V6÷&DF—"v6ö×&—6öâÖæÇ—6—2æ§6öâr’F6× ¢Ò6F6‚²Ð¢2cRÜ*sbãc¢&6VÆ–æR8).i»Nik8ž8(¾8î8þK¸®Y¹î8îŠz>ié8¾h‰X©þ8~8þ8Ž8Þ888 ¢2Væ¶æ÷vâx˜Ž8).Yû®k©n8¾8ž8(¾8ŽjÊY¹î8îjùN‹È>XX>8ÎZK8(þ8(Î8(¾8 ¢F†4ö²Ò„vWBÔ'&’„vWBÔFF&÷W'G’G'6VBw6†VWG2r‚’’Âv†W&RÔö&¦V7B²·7G&–æuÒEòç7FGW2ÖWvö²rÒ’ä6÷VçBÖwB ¢–b‚F†4ö²’²6WBÔ6ö×&—6öä&6VÆ–æRDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–B…·7G&–æuÒE67&—C¤7W'&VçE&VæFW$Vçdf–ævW'&–çB’Ð¢2jùN‹È>8;¶&6VÆ–æ^KùÞŠÛ~8î8~{X.8(þ8>8þiÈ{X.x«nhX¾8).86ö×ÆWFR÷Væf–Æ&ÆR8°¢28¾8¾8(þ8(ž8®[^jÛNKˆŠj~8ŽXøÞiŠ8ž8(¾8.KÙÎh‰˜	NKŠÞ8îŠh{HN8).XhÞXŠžyJŽ8~8®8N8 ¢·fö–EÒ…WFFRÕ6æ6†÷E7VÖÖ'”66†TVçG'’DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B¢&WGW&âF6× ¢Ò6F6‚°¢6ÆV"Õ6æ6†÷E7VÖÖ'”66†RDÆæwVvREv÷&¶&öö´–@¢w&—FRÕv&æ–ær‚~yK¾X8þ88þ88>8+~8:^8îŠz>ié8¾ZKiY~8~8î8~8ó¢r²EòäW†6WF–öâäÖW76vR¢&WGW&âFçVÆÀ¢Ð§Ð ¦gVæ7F–öâvWBÔÆFW7D6ö×&—6öâ…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂEv÷&¶&öö²ÒFçVÆÂ’°¢G'’°¢2cRÕ#¢YÎX{®XX>8Îiz.8²7G'V7GW&R8).ŠªÞ8)>8~8N8(¾ZNYŽ8þXhÞŠªÞ‹ëÎ8~8®8N8 ¢2ö’÷7FFR8þ89n88>8*óK»n8N8Ž8¾8>8>8ŽiÚ^8(¾8þ8(8“K»n8®8(’7G'V7GW&Ræ§6öâ8)#“Y¹îŠªÞ8)>8~8N8þ8 ¢GF&vWBÒEv÷&¶&öö°¢–b‚FçVÆÂÖWGF&vWB’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢Gv"Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWEv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Gv"ä6÷VçBÖW’²&WGW&âFçVÆÂÐ¢GF&vWBÒGv%³Ð¢Ð¢G6æÒ·7G&–æuÒ„vWBÔFF&÷W'G’GF&vWBvÆ7E&VæFW&VE6æ6†÷D–Brrr¢GfW"Ò·7G&–æuÒ„vWBÔFF&÷W'G’GF&vWBvÆ7E&VæFW&VEfW'6–öä–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6æ’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GfW"’’²&WGW&âFçVÆÂÐ¢F66†T¶W’Ò‚DÆæwVvR²wÂr²Ev÷&¶&öö´–B²wÂr²G6æ²wÂr²GfW"’åFôÆ÷vW$–çf&–çB‚¢–b‚E67&—C¤ÆFW7D6ö×&—6öä66†Rä6öçF–ç4¶W’‚F66†T¶W’’’°¢&WGW&âE67&—C¤ÆFW7D6ö×&—6öä66†U²F66†T¶W•Ð¢Ð¢2iz.ZÙŽ88~8;Î8+þ8(.jùN‹È>89^8*ž8:¾888;Î8).X‰~hÉž8¾8®88:Î8;>888:®8;>8+i˜.8îz+®Zé®89^8*8*N8:¾8).y»Nhê^ŠªÞ8(8 ¢G&V6÷&DF—"ÒvWBÕ&VæFW%&V6÷&DF—"DÆæwVvREv÷&¶&öö´–BG6æGfW ¢FæÇ—6—5F‚Ò¦ö–âÕF‚G&V6÷&DF—"v6ö×&—6öâÖæÇ—6—2æ§6öâp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FæÇ—6—5F‚’°¢G'’°¢F6æF–FFRÒ&VBÔ§6öäf–ÆRFæÇ—6—5F‚FçVÆÀ¢G66÷RÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRw66÷Rrrr¢F6æF–FFU6æ6†÷BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRv7W'&VçE6æ6†÷D–Brrr¢F6æF–FFUfW'6–öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRv7W'&VçEfW'6–öä–Brrr¢–b‚FçVÆÂÖæRF6æF–FFRÖæB…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G66÷R’Ö÷"G66÷RÖWvWFöÖF–2r’Öæ@¢…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6æF–FFU6æ6†÷B’Ö÷"F6æF–FFU6æ6†÷BÖWG6æ’Öæ@¢…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6æF–FFUfW'6–öâ’Ö÷"F6æF–FFUfW'6–öâÖWGfW"’’°¢E67&—C¤ÆFW7D6ö×&—6öä66†U²F66†T¶W•ÒÒF6æF–FFP¢&WGW&âF6æF–FFP¢Ð¢Ò6F6‚²Ð¢Ð¢FF—"Ò¦ö–âÕF‚G&V6÷&DF—"v6ö×&—6öç2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&âFçVÆÂÐ¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂ6÷'BÔö&¦V7BÆ7Ew&—FUF–ÖUWF2ÔFW66VæF–ær’’°¢G'’°¢F6æF–FFRÒ&VBÔ§6öäf–ÆRFbägVÆÄæÖRFçVÆÀ¢–b‚FçVÆÂÖWF6æF–FFR’²6öçF–çVRÐ¢G66÷RÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRw66÷Rrrr¢266÷^xJ8~8þiz~x˜ŽK©.hù¾8.[^jÛNyK¾™Ú.8¾8(žKÙÎ8>8þK»¾hHþjùN‹È>8þ8¢2iÈikx˜Ž8îZHži»N8988>8+Ž8(Nˆz®X¹^jùN‹È>Yû®k©n8Ž8~8nh›8(þ8®8N8 ¢–b‚Öæ÷B…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G66÷R’Ö÷"G66÷RÖWvWFöÖF–2r’’²6öçF–çVRÐ¢2ik[Ú.[Èþ8þxûîYÊŽx˜Ž8ç6æ6†÷B÷fW'6–öî8(.ZèÎXZŽKˆˆ{N8^8¾8(¾8.iz~[Ú.[Èþ8~š^yºî8ÎxJ8NZNYŽ888¢2xšžyny¨N8¾xûîYÊŽx˜‡&VæFW.89^8*ž8:¾88Xh^8¾8.8(¾8>8Ž8).jžhº8¾K©.hù¾ŠªÞ‹ëÎ8ž8(¾8 ¢F6æF–FFU6æ6†÷BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRv7W'&VçE6æ6†÷D–Brrr¢F6æF–FFUfW'6–öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRv7W'&VçEfW'6–öä–Brrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6æF–FFU6æ6†÷B’ÖæBF6æF–FFU6æ6†÷BÖæRG6æ’²6öçF–çVRÐ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F6æF–FFUfW'6–öâ’ÖæBF6æF–FFUfW'6–öâÖæRGfW"’²6öçF–çVRÐ¢–b‚Öæ÷BE67&—C¤ÆFW7D6ö×&—6öä66†Rä6öçF–ç4¶W’‚F66†T¶W’’Öæ@¢E67&—C¤ÆFW7D6ö×&—6öä66†Rä6÷VçBÖvRE67&—C¤ÆFW7D6ö×&—6öä66†TÆ–Ö—B’°¢FöÆFW7D¶W’Ò‚E67&—C¤ÆFW7D6ö×&—6öä66†Rä¶W—2•³Ð¢–b‚FçVÆÂÖæRFöÆFW7D¶W’’²·fö–EÒE67&—C¤ÆFW7D6ö×&—6öä66†Rå&VÖ÷fR‚FöÆFW7D¶W’’Ð¢Ð¢26æ6†÷B÷fW'6–öîK¹Ž8Þ8îˆz®X¹^jùN‹È>8þKˆÞZHž8®8î8~87FF^8~ŠªÞ8)>8{YiéÎ8).Š›>{KŠŽzK®8Ž[É^8Þ{iž88 ¢E67&—C¤ÆFW7D6ö×&—6öä66†U²F66†T¶W•ÒÒF6æF–FFP¢&WGW&âF6æF–FFP¢Ò6F6‚²Ð¢Ð¢&WGW&âFçVÆÀ¢Ò6F6‚²&WGW&âFçVÆÂÐ§Ð ¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ¢2cR7FvRR(	B†6R"ò$3¢8*.8;Î8*¾8*N89n8;¾8:Î8*N8*.8*n88Ž[^jÛN8;¾X{®X©¾88Ž8:ž8;>8+n8*þ8+~8:~8;0¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ ¦gVæ7F–öâvWBÔ&6†—fU&ö÷B…·7G&–æuÒDÆæwVvR’²&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’vW‡÷'G5Æ&6†—fRr’Ð¦gVæ7F–öâvWBÔÆ–÷WD†—7F÷'”F—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE7F÷&vT–B’°¢G6fU7F÷&vT–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE7F÷&vT–Bw6´–Bp¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’„¦ö–âÕF‚vÆ–÷WBÖ†—7F÷'’rG6fU7F÷&vT–B’§Ð¦gVæ7F–öâvWBÔÆ–÷WE66÷T–æfò‚E7G'V7GW&RÂ·7G&–æuÒE6´–D÷$6FVv÷'’Â¶&ööÅÒDÆÆ÷t&6†—fVBÒFfÇ6R’°¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RE7G'V7GW&RE6´–D÷$6FVv÷'’DÆÆ÷t&6†—fV@¢G7F÷&vT–BÒ–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒG66÷Ræ6FVv÷'’’’²·7G&–æuÒG66÷Rç6´–BÒVÇ6R²·7G&–æuÒG66÷Ræ6FVv÷'’Ð¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢6´–BÒ·7G&–æuÒG66÷Rç6´–@¢6FVv÷'’Ò·7G&–æuÒG66÷Ræ6FVv÷'¢7F÷&vT–BÒ„76W'BÕ6fU7F÷&vU6VvÖVçBG7F÷&vT–Bw6´–Br¢6²ÒG66÷Rç6°¢'V–ÇD–âÒ¶&ööÅÒG66÷Ræ'V–ÇD–à¢Ð§Ð¦gVæ7F–öâ–çfö¶RÔÆ–÷WE6æ6†÷E&WFVçF–öâ…·7G&–æuÒDF—&V7F÷'’Â¶–çEÒD¶VWÒ’°¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚DF—&V7F÷'’’’²&WGW&âÐ¢FÆ–Ö—BÒ´ÖF…Ó£¤Ö‚ƒÂD¶VW¢f÷&V6‚‚FöÆB–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚DF—&V7F÷'’Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂ6÷'BÔö&¦V7BæÖRÔFW66VæF–ærÂ6VÆV7BÔö&¦V7BÕ6¶—FÆ–Ö—B’’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FöÆBägVÆÄæÖRÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢Ð§Ð¦gVæ7F–öâvWBÔf–æÅG&ç67F–öäF—"…·7G&–æuÒDÆæwVvR’²&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’w7FFUÆf–æÂ×G&ç67F–öç2r’Ð¦gVæ7F–öâvWBÔf–æÅG&ç67F–öä&6·WF—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEG&ç67F–öä–B’²&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’„¦ö–âÕF‚w7FFUÆf–æÂÖ&6·W2rEG&ç67F–öä–B’’Ð¦gVæ7F–öâ&VÖ÷fRÔf–æÅG&ç67F–öä&6·WF—"…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEG&ç67F–öä–B’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EG&ç67F–öä–B’’²&WGW&âÐ¢G'’°¢FF—"ÒvWBÔf–æÅG&ç67F–öä&6·WF—"DÆæwVvREG&ç67F–öä–@¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FF—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ò6F6‚²Ð§Ð ¢2ÒÒÒÒ8:Î8*N8*.8*n88Žh©^[Û8+ž88®88>89~8+~8:~88>88Ž8Ž™™Zé®[êžXX2…cRÜ*sBã"’ÒÒÒÒÒÒÒÒÒÒÒÒÐ ¦gVæ7F–öâ6öçfW'EFòÔæ÷&ÖÆ—¦VDÆ–÷WE6æ6†÷EvR‚EvRÂ·7G&–æuµÕÒDÆÆ÷vVEföÇVÖW2’°¢GvT–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’EvRwvT–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GvT–B’’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~89®8;Î8+Žjx¾h‰[^jÛN8·vT–N8Î8.8(®8î8¾8)>8"r’Ð¢G7F÷&VEföÇVÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’EvRwföÇVÖRrvæöæRr¢GföÇVÖT–çfÆ–BÒ‚DÆÆ÷vVEföÇVÖW2Öæ÷F6öçF–ç2G7F÷&VEföÇVÖR¢GföÇVÖRÒB†–b‚GföÇVÖT–çfÆ–B’²væöæRrÒVÇ6R²G7F÷&VEföÇVÖRÒ¢FçVÖ&W&–ætÖöFRÒ·7G&–æuÒ„vWBÔFF&÷W'G’EvRvçVÖ&W&–ætÖöFRrwf—6–&ÆRr¢–b‚FçVÖ&W&–ætÖöFRÖæ÷F–â‚væöæRrÂwf—6–&ÆRr’’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚.89®8;Î8+Žjx¾h‰[^jÛN8æçVÖ&W&–ætÖöF^8ÎKˆÞjÚ>8~8“¢GvT–B"’Ð¢G'’²F÷&FW"Ò¶F÷V&ÆUÒ„vWBÔFF&÷W'G’EvRv÷&FW"r’Ò6F6‚²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚.89®8;Î8+Žjx¾h‰[^jÛN8æ÷&FW.8ÎKˆÞjÚ>8~8“¢GvT–B"’Ð¢–b…¶F÷V&ÆUÓ£¤—4æâ‚F÷&FW"’Ö÷"¶F÷V&ÆUÓ£¤—4–æf–æ—G’‚F÷&FW"’’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚.89®8;Î8+Žjx¾h‰[^jÛN8æ÷&FW.8ÎKˆÞjÚ>8~8“¢GvT–B"’Ð¢G&ævRÒ6öçfW'EFòÔæ÷&ÖÆ—¦VEvU&ævR„vWBÔFF&÷W'G’EvRwvU&ævRrFçVÆÂ¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢vT–BÒGvT–@¢F—FÆRÒ·7G&–æuÒ„vWBÔFF&÷W'G’EvRwF—FÆRrrr¢föÇVÖRÒGföÇVÖP¢7F÷&VEföÇVÖRÒG7F÷&VEföÇVÖP¢föÇVÖT–çfÆ–BÒGföÇVÖT–çfÆ–@¢Væ&ÆVBÒ‚GföÇVÖRÖæRvæöæRr¢÷&FW"ÒF÷&FW ¢÷&FW$ÖçVÂÒ¶&ööÅÒ„vWBÔFF&÷W'G’EvRv÷&FW$ÖçVÂrFfÇ6R¢çVÖ&W&–ætÖöFRÒFçVÖ&W&–ætÖöFP¢çVÖ&W&–ætÖçVÂÒ¶&ööÅÒ„vWBÔFF&÷W'G’EvRvçVÖ&W&–ætÖçVÂrFfÇ6R¢vU&ævRÒG&ævP¢Ð§Ð ¦gVæ7F–öâvWBÔæ÷&ÖÆ—¦VDÆ–÷WE6æ6†÷EvW2‚E6æ6†÷BÂ·7G&–æuµÕÒDÆÆ÷vVEföÇVÖW2’°¢GvW2Ò‚¢G6VVâÒ·Ð¢f÷&V6‚‚GvR–â„vWBÔ'&’„vWBÔFF&÷W'G’E6æ6†÷BwvW2r‚’’’’°¢Fæ÷&ÖÆ—¦VBÒ6öçfW'EFòÔæ÷&ÖÆ—¦VDÆ–÷WE6æ6†÷EvRGvRDÆÆ÷vVEföÇVÖW0¢GvT–BÒ·7G&–æuÒFæ÷&ÖÆ—¦VBçvT–@¢–b‚G6VVâä6öçF–ç4¶W’‚GvT–B’’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚.89®8;Î8+Žjx¾h‰[^jÛN8¾˜xÞŠH~8~8÷vT–N8Î8.8(®8î8“¢GvT–B"’Ð¢G6VVå²GvT–EÒÒGG'VP¢GvW2³ÒFæ÷&ÖÆ—¦V@¢Ð¢&WGW&â‚GvW2§Ð ¦gVæ7F–öâ6fRÔÆ–÷WE6æ6†÷B…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–D÷$6FVv÷'’Â·7G&–æuÒE&V6öâÂE7G'V7GW&RÒFçVÆÂ’°¢–b‚Öæ÷B…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’’²&WGW&ârrÐ¢G'’°¢G7BÒE7G'V7GW&P¢–b‚FçVÆÂÖWG7B’²G7BÒvWBÕ7G'V7GW&RDÆæwVvRÐ¢G66÷RÒvWBÔÆ–÷WE66÷T–æfòG7BE6´–D÷$6FVv÷'’FfÇ6P¢GvW2Ò‚¢f÷&V6‚‚G–â„vWBÔ'&’G7BçvW2’’°¢Gv"Ò„vWBÔ'&’G7Bçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖW·7G&–æuÒGçv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Gv"ä6÷VçBÖW’²6öçF–çVRÐ¢–b‚Öæ÷B…FW7BÕv÷&¶&ööµ6²Gv%³Ò…·7G&–æuÒG66÷Rç6´–B’’’²6öçF–çVRÐ¢2cRÜ*sBã#¢8:Î8*N8*.8*n88Žš^yºî888).KùÞZÙŽ8ž8(¾8'7G'V7GW&RXZŽKÙ>8þKùÞZÙŽ8~8®8N8 ¢GvW2³Ò¶÷&FW&VEÔ°¢vT–BÒ…&W6öÇfRÕvT–BG¢F—FÆRÒ·7G&–æuÒGçF—FÆP¢föÇVÖRÒ·7G&–æuÒGçföÇVÖP¢Væ&ÆVBÒ¶&ööÅÒGæVæ&ÆV@¢÷&FW"Ò¶F÷V&ÆUÒGæ÷&FW ¢÷&FW$ÖçVÂÒ¶&ööÅÒ„vWBÔFF&÷W'G’Gv÷&FW$ÖçVÂrFfÇ6R¢çVÖ&W&–ætÖöFRÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvçVÖ&W&–ætÖöFRrwf—6–&ÆRr¢çVÖ&W&–ætÖçVÂÒ¶&ööÅÒ„vWBÔFF&÷W'G’GvçVÖ&W&–ætÖçVÂrFfÇ6R¢vU&ævRÒvWBÔFF&÷W'G’GwvU&ævRrFçVÆÀ¢Ð¢Ð¢–b‚GvW2ä6÷VçBÖW’²&WGW&ârrÐ¢FF—"ÒvWBÔÆ–÷WD†—7F÷'”F—"DÆæwVvR…·7G&–æuÒG66÷Rç7F÷&vT–B¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢F–BÒæWrÕ&$–@¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚FF—"‚'³Òæ§6öâ"ÖbF–B’’…¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ#²6æ6†÷D–BÒF–C²ÆæwVvRÒDÆæwVvS²6´–BÒ·7G&–æuÒG66÷Rç6´–C²6FVv÷'’Ò·7G&–æuÒG66÷Ræ6FVv÷'¢7&VFVDBÒæWrÔæ÷t—6ó²&V6öâÒE&V6öã²vW2ÒGvW0¢Ò¢–çfö¶RÔÆ–÷WE6æ6†÷E&WFVçF–öâFF—" ¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvÆ–÷WBæ6†ævVBr…¶÷&FW&VEÔ²6´–BÒ·7G&–æuÒG66÷Rç6´–C²6FVv÷'’Ò·7G&–æuÒG66÷Ræ6FVv÷'“²&V6öâÒE&V6öã²6æ6†÷D–BÒF–C²vT6÷VçBÒGvW2ä6÷VçBÒ¢&WGW&âF–@¢Ò6F6‚²&WGW&ârrÐ§Ð ¦gVæ7F–öâvWBÔÆ–÷WE6æ6†÷G2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–D÷$6FVv÷'’’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢G66÷RÒvWBÔÆ–÷WE66÷T–æfòG7G'V7GW&RE6´–D÷$6FVv÷'’GG'VP¢FF—"ÒvWBÔÆ–÷WD†—7F÷'”F—"DÆæwVvR…·7G&–æuÒG66÷Rç7F÷&vT–B¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&â‚’Ð¢F÷WBÒ‚¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂ6÷'BÔö&¦V7BæÖRÔFW66VæF–ærÂ6VÆV7BÔö&¦V7BÔf—'7B’’°¢G'’°¢F¢Ò&VBÔ§6öäf–ÆRFbägVÆÄæÖRFçVÆÀ¢F÷WB³Ò¶÷&FW&VEÔ°¢6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢w6æ6†÷D–BrFbä&6TæÖR¢7&VFVDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢v7&VFVDBrrr¢&V6öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢w&V6öârrr¢vT6÷VçBÒ„vWBÔ'&’„vWBÔFF&÷W'G’F¢wvW2r‚’’’ä6÷Vç@¢6´–BÒ·7G&–æuÒG66÷Rç6´–@¢6FVv÷'’Ò·7G&–æuÒG66÷Ræ6FVv÷'¢Ð¢Ò6F6‚²Ð¢Ð¢&WGW&âF÷W@§Ð ¦gVæ7F–öâ&VBÔÆ–÷WE6æ6†÷B…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–D÷$6FVv÷'’Â·7G&–æuÒE6æ6†÷D–B’°¢G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Bp¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢G66÷RÒvWBÔÆ–÷WE66÷T–æfòG7G'V7GW&RE6´–D÷$6FVv÷'’GG'VP¢GÒ¦ö–âÕF‚„vWBÔÆ–÷WD†—7F÷'”F—"DÆæwVvR…·7G&–æuÒG66÷Rç7F÷&vT–B’’‚'³Òæ§6öâ"ÖbG6fU6æ6†÷D–B¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G’’²F‡&÷r.8:Î8*N8*.8*n88Ž[^jÛN8ÎŠh¾8N8¾8(®8î8¾8)3¢E6æ6†÷D–B"Ð¢G6æ6†÷BÒ&VBÔ§6öäf–ÆRGFçVÆÀ¢–b‚FçVÆÂÖWG6æ6†÷B’²F‡&÷r.8:Î8*N8*.8*n88Ž[^jÛN8).ŠªÞ8þ‹ëÎ8(8î8¾8)3¢E6æ6†÷D–B"Ð¢G&V6÷&FVE6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6æ6†÷Bw6æ6†÷D–Brrr¢G&V6÷&FVDÆæwVvRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6æ6†÷BvÆæwVvRrrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&V6÷&FVE6æ6†÷D–B’ÖæBG&V6÷&FVE6æ6†÷D–BÖæRG6fU6æ6†÷D–B’²F‡&÷r~89®8;Î8+Žjx¾h‰[^jÛN8îŠÙŽXŠ^ZÙ8Î89^8*8*N8:¾YÞ8ŽKˆˆ{N8~8î8¾8)>8"rÐ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&V6÷&FVDÆæwVvR’ÖæBG&V6÷&FVDÆæwVvRÖæRDÆæwVvR’²F‡&÷r~XŠ^8îŠˆŠ©î8î89®8;Î8+Žjx¾h‰[^jÛN8þ[êžXX>8~8Þ8î8¾8)>8"rÐ¢G&V6÷&FVE6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6æ6†÷Bw6´–Brrr¢G&V6÷&FVD6FVv÷'’Ò·7G&–æuÒ„vWBÔFF&÷W'G’G6æ6†÷Bv6FVv÷'’rrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&V6÷&FVE6´–B’ÖæBG&V6÷&FVE6´–BÖæR·7G&–æuÒG66÷Rç6´–B’²F‡&÷r~XŠ^8îKˆ[Èþ8î89®8;Î8+Žjx¾h‰[^jÛN8þ[êžXX>8~8Þ8î8¾8)>8"rÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&V6÷&FVE6´–B’ÖæBÖæ÷B¶&ööÅÒG66÷Ræ'V–ÇD–â’²F‡&÷r~iz~[Ú.[Èþ8î89®8;Î8+Žjx¾h‰[^jÛN8þK»¾hHþKˆ[Èþ8Ž[êžXX>8~8Þ8î8¾8)>8"rÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&V6÷&FVE6´–B’ÖæBG&V6÷&FVD6FVv÷'’ÖæR·7G&–æuÒG66÷Ræ6FVv÷'’’²F‡&÷r~XŠ^8î8*¾88n8+N8:®8î89®8;Î8+Žjx¾h‰[^jÛN8þ[êžXX>8~8Þ8î8¾8)>8"rÐ¢&WGW&âG6æ6†÷@§Ð ¦gVæ7F–öâvWBÔÆ–÷WE&W7F÷&U&Wf–Wr…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–D÷$6FVv÷'’Â·7G&–æuÒE6æ6†÷D–B’°¢G6æÒ&VBÔÆ–÷WE6æ6†÷BDÆæwVvRE6´–D÷$6FVv÷'’E6æ6†÷D–@¢G7BÒvWBÕ7G'V7GW&RDÆæwVvP¢G66÷RÒvWBÔÆ–÷WE66÷T–æfòG7BE6´–D÷$6FVv÷'’GG'VP¢F7W'&VçD–G2Ò·Ð¢f÷&V6‚‚G–â„vWBÔ'&’G7BçvW2’’°¢Gv"Ò„vWBÔ'&’G7Bçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖW·7G&–æuÒGçv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Gv"ä6÷VçBÖW’²6öçF–çVRÐ¢–b‚Öæ÷B…FW7BÕv÷&¶&ööµ6²Gv%³Ò…·7G&–æuÒG66÷Rç6´–B’’’²6öçF–çVRÐ¢F7W'&VçD–G5²…&W6öÇfRÕvT–BG•ÒÒG ¢Ð¢FÆÆ÷vVEföÇVÖW2Ò„vWBÕ6µföÇVÖTÆ—7BDÆæwVvRG66÷Rç6²GG'VR¢FÆ–VBÒ²G7DöæÇ’Ò‚“²GföÇVÖT6†ævW2Ò‚“²F–çfÆ–EföÇVÖUvT–G2Ò‚¢G6æ–G2Ò·Ð¢f÷&V6‚‚Fæ÷&ÖÆ—¦VB–â„vWBÔæ÷&ÖÆ—¦VDÆ–÷WE6æ6†÷EvW2G6æFÆÆ÷vVEföÇVÖW2’’°¢GvT¶W’Ò·7G&–æuÒFæ÷&ÖÆ—¦VBçvT–@¢G6æ–G5²GvT¶W•ÒÒGG'VP¢–b‚Öæ÷BF7W'&VçD–G2ä6öçF–ç4¶W’‚GvT¶W’’’²G7DöæÇ’³ÒGvT¶W“²6öçF–çVRÐ¢FÆ–VB²°¢F7W"ÒF7W'&VçD–G5²GvT¶W•Ð¢–b…¶&ööÅÒFæ÷&ÖÆ—¦VBçföÇVÖT–çfÆ–B’²F–çfÆ–EföÇVÖUvT–G2³ÒGvT¶W’Ð¢–b…·7G&–æuÒF7W"çföÇVÖRÖæR·7G&–æuÒFæ÷&ÖÆ—¦VBçföÇVÖR’°¢GföÇVÖT6†ævW2³Ò¶÷&FW&VEÔ²vT–BÒGvT¶W“²F—FÆRÒ·7G&–æuÒF7W"çF—FÆS²g&öÒÒ·7G&–æuÒF7W"çföÇVÖS²FòÒ·7G&–æuÒFæ÷&ÖÆ—¦VBçföÇVÖS²7F÷&VEföÇVÖRÒ·7G&–æuÒFæ÷&ÖÆ—¦VBç7F÷&VEföÇVÖS²æ÷&ÖÆ—¦VBÒ¶&ööÅÒFæ÷&ÖÆ—¦VBçföÇVÖT–çfÆ–BÐ¢Ð¢Ð¢F7W'&VçDöæÇ’Ò‚F7W'&VçD–G2ä¶W—2Âv†W&RÔö&¦V7B²Öæ÷BG6æ–G2ä6öçF–ç4¶W’‚Eò’Ò¢&WGW&â¶÷&FW&VEÔ°¢6æ6†÷D–BÒE6æ6†÷D–@¢6´–BÒ·7G&–æuÒG66÷Rç6´–@¢6FVv÷'’Ò·7G&–æuÒG66÷Ræ6FVv÷'¢7&VFVDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6æv7&VFVDBrrr¢&V6öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6æw&V6öârrr¢Æ–VEvT6÷VçBÒFÆ–V@¢7DöæÇ•vT–G2Ò‚G7DöæÇ’¢7W'&VçDöæÇ•vT–G2Ò‚F7W'&VçDöæÇ’¢föÇVÖT6†ævW2Ò‚GföÇVÖT6†ævW2¢–çfÆ–EföÇVÖUvT–G2Ò‚F–çfÆ–EföÇVÖUvT–G2¢&WV—&W5&V'V–ÆBÒ‚FÆ–VBÖwB¢Ð§Ð ¦gVæ7F–öâ&W7F÷&RÔÆ–÷WE6æ6†÷B…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–D÷$6FVv÷'’Â·7G&–æuÒE6æ6†÷D–B’°¢G6æÒ&VBÔÆ–÷WE6æ6†÷BDÆæwVvRE6´–D÷$6FVv÷'’E6æ6†÷D–@¢G66÷RÒvWBÔÆ–÷WE66÷T–æfò„vWBÕ7G'V7GW&RDÆæwVvR’E6´–D÷$6FVv÷'’FfÇ6P¢G&W7F÷&U&W7VÇBÒWFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7B¢FÆö6¶VE66÷RÒvWBÔÆ–÷WE66÷T–æfòG7B…·7G&–æuÒG66÷Rç6´–B’FfÇ6P¢FÆÆ÷vVEföÇVÖW2Ò„vWBÕ6µföÇVÖTÆ—7BDÆæwVvRFÆö6¶VE66÷Rç6²GG'VR¢Gv÷&¶&öö´–G2Ò·Ð¢f÷&V6‚‚Gv"–â„vWBÔ'&’G7Bçv÷&¶&öö·2Âv†W&RÔö&¦V7B²FW7BÕv÷&¶&ööµ6²Eò…·7G&–æuÒFÆö6¶VE66÷Rç6´–B’Ò’’²Gv÷&¶&öö´–G5µ·7G&–æuÒGv"çv÷&¶&öö´–EÒÒGG'VRÐ¢FÖÒ·Ð¢f÷&V6‚‚G–â„vWBÔ'&’G7BçvW2Âv†W&RÔö&¦V7B²Gv÷&¶&öö´–G2ä6öçF–ç4¶W’…·7G&–æuÒEòçv÷&¶&öö´–B’Ò’’²FÖ²…&W6öÇfRÕvT–BG•ÒÒGÐ¢G&V6÷&G2Ò‚¢f÷&V6‚‚Fæ÷&ÖÆ—¦VB–â„vWBÔæ÷&ÖÆ—¦VDÆ–÷WE6æ6†÷EvW2G6æFÆÆ÷vVEföÇVÖW2’’°¢GvT¶W’Ò·7G&–æuÒFæ÷&ÖÆ—¦VBçvT–@¢–b‚Öæ÷BFÖä6öçF–ç4¶W’‚GvT¶W’’’²6öçF–çVRÐ¢G&V6÷&G2³Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²vRÒFÖ²GvT¶W•Ó²6æ6†÷BÒFæ÷&ÖÆ—¦VBÐ¢Ð¢–b‚G&V6÷&G2ä6÷VçBÖW’²&WGW&â¶÷&FW&VEÔ²Æ–VBÒ²VæFõ6æ6†÷D–BÒrs²–çfÆ–EföÇVÖUvT6÷VçBÒÒÐ¢26fRF†RW†7BÆö6¶VB&R×&W7F÷&R7FFRâf–ÆVBVæFò6æ6†÷B×W7B&÷'@¢2&Vf÷&R7G'V7GW&Ræ§6öâ—26†ævVBà¢GVæFô–BÒ6fRÔÆ–÷WE6æ6†÷BDÆæwVvR…·7G&–æuÒFÆö6¶VE66÷Rç6´–B’w&R×&W7F÷&RrG7@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GVæFô–B’’²F‡&÷r~[êžXX>y»NX˜Þ8î89®8;Î8+Žjx¾h‰8).KùÞZÙŽ8~8Þ8®8N8þ8(8[êžXX>8).KŠÞjÚ.8~8î8~8þ8"rÐ¢FâÒ²F–çfÆ–EföÇVÖUvT6÷VçBÒ ¢f÷&V6‚‚G&V6÷&B–âG&V6÷&G2’°¢GÒG&V6÷&BçvP¢G7ÒG&V6÷&Bç6æ6†÷@¢–b…¶&ööÅÒG7çföÇVÖT–çfÆ–B’²F–çfÆ–EföÇVÖUvT6÷VçB²²Ð¢2cRÜ*sBã#¢˜žyJŽ8~8n8(Ž8N8î8þ8:Î8*N8*.8*n88Žš^yºî8î8þ8 ¢26öçFVçEFbò7FGW2òv&æ–æw2ò7W'&VçDW†6VÄ†6‚òÆ7E&VæFW&VB¢òföÇVÖW28þŠzn8(ž8®8N8 ¢6WBÔæ÷FU&÷W'G’GwF—FÆRr…·7G&–æuÒG7çF—FÆR¢6WBÔæ÷FU&÷W'G’GwföÇVÖRr…·7G&–æuÒG7çföÇVÖR¢6WBÔæ÷FU&÷W'G’GvVæ&ÆVBr…¶&ööÅÒG7æVæ&ÆVB¢6WBÔæ÷FU&÷W'G’Gv÷&FW"r…¶F÷V&ÆUÒG7æ÷&FW"¢6WBÔæ÷FU&÷W'G’Gv÷&FW$ÖçVÂr…¶&ööÅÒG7æ÷&FW$ÖçVÂ¢6WBÔæ÷FU&÷W'G’GvçVÖ&W&–ætÖöFRr…·7G&–æuÒG7æçVÖ&W&–ætÖöFR¢6WBÔæ÷FU&÷W'G’GvçVÖ&W&–ætÖçVÂr…¶&ööÅÒG7æçVÖ&W&–ætÖçVÂ¢6WBÔæ÷FU&÷W'G’GwvU&ævRr„vWBÔFF&÷W'G’G7wvU&ævRrFçVÆÂ¢6WBÔæ÷FU&÷W'G’GwWFFVDBr„æWrÔæ÷t—6ò¢Fâ²°¢Ð¢f÷&V6‚‚GföÇVÖR–âFÆÆ÷vVEföÇVÖW2’²·fö–EÒ…&VçVÖ&W"ÕföÇVÖT÷&FW"G7BGföÇVÖR…·7G&–æuÒFÆö6¶VE66÷Rç6´–B’’Ð¢Ç’ÔFVfVÇDçVÖ&W&–æuW%föÇVÖRDÆæwVvRG7B…·7G&–æuÒFÆö6¶VE66÷Rç6´–B¢GföÇ2Ò„vWBÕ6µföÇVÖTÆ—7BDÆæwVvRFÆö6¶VE66÷Rç6²FfÇ6R¢Ö&²ÕföÇVÖTæVVG5&V'V–ÆBG7BDÆæwVvR…·7G&–æuÒFÆö6¶VE66÷Rç6´–B’GföÇ2vÆ–÷WB×&W7F÷&VBr~89®8;Î8+Žjx¾h‰8).˜îXë¾8îx«nhX¾8Žh‹¾8~8î8~8òp¢&WGW&â¶÷&FW&VEÔ²Æ–VBÒFã²VæFõ6æ6†÷D–BÒGVæFô–C²–çfÆ–EföÇVÖUvT6÷VçBÒF–çfÆ–EföÇVÖUvT6÷VçBÐ¢Ð¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvÆ–÷WBç&W7F÷&VBr…¶÷&FW&VEÔ²6´–BÒ·7G&–æuÒG66÷Rç6´–C²6FVv÷'’Ò·7G&–æuÒG66÷Ræ6FVv÷'“²6æ6†÷D–BÒE6æ6†÷D–C²VæFõ6æ6†÷D–BÒ·7G&–æuÒG&W7F÷&U&W7VÇBçVæFõ6æ6†÷D–C²Æ–VEvT6÷VçBÒ¶–çEÒG&W7F÷&U&W7VÇBæÆ–VBÒ¢&WGW&â¶÷&FW&VEÔ²6´–BÒ·7G&–æuÒG66÷Rç6´–C²6FVv÷'’Ò·7G&–æuÒG66÷Ræ6FVv÷'“²Æ–VEvT6÷VçBÒ¶–çEÒG&W7F÷&U&W7VÇBæÆ–VC²VæFõ6æ6†÷D–BÒ·7G&–æuÒG&W7F÷&U&W7VÇBçVæFõ6æ6†÷D–C²–çfÆ–EföÇVÖUvT6÷VçBÒ¶–çEÒG&W7F÷&U&W7VÇBæ–çfÆ–EföÇVÖUvT6÷VçBÐ§Ð ¢2ÒÒÒÒhùX{®yJ…Dn8*.8;Î8*¾8*N89b…cRÜ*sBã’ÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÐ ¢2–â8þ8hùX{®yJ…Dn8ÎXø.xZ~8~8n8N8(²6÷W&6Rò6öçFVçB×Fb8).[èÎiz^8îhè>™šN8¾8(žZèŽ8(¾YJþKˆ8à¢2K¹^{XN8þ8~8.8(²„–çfö¶RÔ–çWD†—7F÷'”6ÆVçW8).Xø.xZ~8'–â8ÎxJ8Nx˜Ž8þKùÞhÈiÉþ™i>8îxËnK¨Ž8®8~8°¢2X˜®™šN8îZûî‹8¾8®8(²ž8.8*.8;Î8*¾8*N89n8îKÙÎh‰8¾ZKiY~8~8n8(.88>8îKùÞŠÛ~888þjè¾8ž[ø^Šh8Î8.8(¾8þ8(8¢28*.8;Î8*¾8*N89niÊÎKÙ>8¾8(žXˆ~8(®™º.8~8n8.8(¾8 ¦gVæ7F–öâ6WBÔf–æÅFe6æ6†÷E–ç2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒE6´–BÂ·7G&–æuÒD6FVv÷'’Â·7G&–æuÒEföÇVÖRÂ·7G&–æuÒD'V–ÆD–BÂE6÷W&6Uv÷&¶&öö·2Â·7G&–æuÒD&6†—fUF‚’°¢f÷&V6‚‚G7r–â„vWBÔ'&’E6÷W&6Uv÷&¶&öö·2’’°¢Gv$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7rwv÷&¶&öö´–Brrr¢G6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7rw6æ6†÷D–Brrr¢GfW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7rwfW'6–öä–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Gv$–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6æ6†÷D–B’’²6öçF–çVRÐ¢G–äö²ÒæWrÕ6æ6†÷E–âDÆæwVvRGv$–BG6æ6†÷D–B‚&f–æÂ×Fe÷³Ò"ÖbD'V–ÆD–B’…¶÷&FW&VEÔ°¢'V–ÆD–BÒD'V–ÆD–C²6æ6†÷D–BÒG6æ6†÷D–C²fW'6–öä–BÒGfW'6–öä–@¢föÇVÖRÒEföÇVÖS²6´–BÒE6´–C²6FVv÷'’ÒD6FVv÷'“²&6†—fUF‚ÒD&6†—fUF€¢Ò¢–b‚Öæ÷BG–äö²’²F‡&÷r‚.8+ž88®88>89~8+~8:~88>88ŽKùÞŠÛr‡–âž8).KÙÎh‰8~8Þ8î8¾8)>8~8~8ó¢³Òò³Ò"ÖbGv$–BÂG6æ6†÷D–B’Ð¢–b‚GfW'6–öä–B’°¢F7–äö²ÒæWrÔ6öçFVçEFe–â„vWBÕv÷&·76UF‚DÆæwVvR’Gv$–BGfW'6–öä–B‚&f–æÂ×Fe÷³Ò"ÖbD'V–ÆD–B’…¶÷&FW&VEÔ°¢'V–ÆD–BÒD'V–ÆD–C²föÇVÖRÒEföÇVÖS²6´–BÒE6´–C²6FVv÷'’ÒD6FVv÷'¢Ò¢–b‚Öæ÷BF7–äö²’²F‡&÷r‚&6öçFVçB×FbKùÞŠÛr‡–âž8).KÙÎh‰8~8Þ8î8¾8)>8~8~8ó¢³Òò³Ò"ÖbGv$–BÂGfW'6–öä–B’Ð¢Ð¢Ð§Ð ¦gVæ7F–öâæWrÔf–æÄ&6†—fR…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒD6FVv÷'’Â·7G&–æuÒEföÇVÖRÂ·7G&–æuÒD'V–ÆD–BÂ·7G&–æuÒD÷WGWEFbÂDÖæ–fW7BÂE6æ6†÷B’°¢2Xj®zØ“¢Kˆi˜.89^8*ž8:¾888~ZèÎh‰8^8¾8n8¾8(’'V–ÆD–B89^8*ž8:¾888Žz{¾X¹^8ž8(¾8 ¢G'’°¢F&6†—fT6FVv÷'’Òæ÷&ÖÆ—¦RÕv÷&¶&öö´6FVv÷'’D6FVv÷'’rp¢F&6†—fU6´–BÒ–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6†—fT6FVv÷'’’’²76W'BÕ6fU7F÷&vU6VvÖVçBD6FVv÷'’w6´–BrÒVÇ6R²vWBÔ'V–ÇF–å6´–BF&6†—fT6FVv÷'’Ð¢F&6†—fU7F÷&vT–BÒ–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6†—fT6FVv÷'’’’²F&6†—fU6´–BÒVÇ6R²F&6†—fT6FVv÷'’Ð¢GF&vWBÒ¦ö–âÕF‚„¦ö–âÕF‚„¦ö–âÕF‚„vWBÔ&6†—fU&ö÷BDÆæwVvR’F&6†—fU7F÷&vT–B’EföÇVÖR’D'V–ÆD–@¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF&vWB’²&WGW&âGF&vWBÐ¢G7FvRÒGF&vWB²rç7Fv–ærp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚G7FvR’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚G7FvRÕ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚G7FvRÔf÷&6RÂ÷WBÔçVÆÀ¢6÷’Ô—FVÒÔÆ—FW&ÅF‚D÷WGWEFbÔFW7F–æF–öâ„¦ö–âÕF‚G7FvRvf–æÂçFbr’Ôf÷&6P¢G6†Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#Sb„¦ö–âÕF‚G7FvRvf–æÂçFbr’¢6WBÔ6öçFVçBÔÆ—FW&ÅF‚„¦ö–âÕF‚G7FvRw6†#SbçG‡Br’ÕfÇVRG6†ÔVæ6öF–ær44”¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚G7FvRvÖæ–fW7Bæ§6öâr’DÖæ–fW7@ ¢2cRÕ¢X{®X©¾8¾8ÎZéþ™©¾8¾KÛþ8>8þ8ÞKˆÞZHž8îh8^Z888).Š‰Ž˜Ë.8ž8(¾8 ¢28>8>8~jx¾˜
88~8;Î8+þ8).ŠªÞ8þy»N8ž8Ž8X{®X©¾[èÎ8¾XŠ^8:Î8;>888:®8;>8+8ÎZèÎK¨n8~8þZNYŽ8°¢2jÚ>[ÈõDn8~KÛþ8>8n8N8®8B6æ6†÷BòfW'6–öä–B8).Š‰Ž˜Ë.8~8n8~8î8n8 ¢FVçg2Ò¶÷&FW&VEÔ·Ð¢G6÷W&6Uv÷&¶&öö·2Ò‚¢f÷&V6‚‚G7r–â„vWBÔ'&’„vWBÔFF&÷W'G’E6æ6†÷Bw6÷W&6Uv÷&¶&öö·2r‚’’’’°¢FgÒ·7G&–æuÒ„vWBÔFF&÷W'G’G7rw&VæFW$Vçf—&öæÖVçDf–ævW'&–çBrrr¢–b‚FgÖæBÖæ÷BFVçg2ä6öçF–ç2‚Fg’’²FVçg5²FgÒÒ„vWBÔFF&÷W'G’G7rw&VæFW$Vçf—&öæÖVçBrFçVÆÂ’Ð¢G6÷W&6Uv÷&¶&öö·2³ÒG7p¢Ð¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚G7FvRvÖWFFFæ§6öâr’…¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ#²'V–ÆD–BÒD'V–ÆD–C²ÆæwVvRÒDÆæwVvS²6´–BÒF&6†—fU6´–C²6FVv÷'’ÒF&6†—fT6FVv÷'“²F&vWD–BÒ„vWBÕF&vWD–Dg&öÔÆVv7•föÇVÖREföÇVÖR“²föÇVÖRÒEföÇVÖP¢&ö¦V7D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’E6æ6†÷Bw&ö¦V7D–Brrr“²'V–ÇDBÒæWrÔæ÷t—6ð¢–çWDf–ævW'&–çBÒ·7G&–æuÒ„vWBÔFF&÷W'G’E6æ6†÷Bvf–ævW'&–çBrrr¢÷WGWEFe6†#SbÒG6†²÷WGWDf–ÆTæÖRÒ´”òåF…Ó£¤vWDf–ÆTæÖR‚D÷WGWEFb¢vT6÷VçBÒ„vWBÔ'&’„vWBÔFF&÷W'G’E6æ6†÷BwvW2r‚’’’ä6÷Vç@¢6÷W&6Uv÷&¶&öö·2Ò‚G6÷W&6Uv÷&¶&öö·2¢W†6VÅ&–çE&öf–ÆUfW'6–öâÒE67&—C¤W†6VÅ&–çE&öf–ÆUfW'6–öà¢&VæFW$Vçf—&öæÖVçG2ÒFVçg0¢6ö×÷6W$Vçf—&öæÖVçBÒ¶÷&FW&VEÔ²4æÖRÒFVçc¤4ôÕUDU$äÔS²÷5fW'6–öâÒ´Vçf—&öæÖVçEÓ£¤õ5fW'6–öâåfW'6–öå7G&–ærÐ¢'V–ÇD'’Ò¶÷&FW&VEÔ²4æÖRÒFVçc¤4ôÕUDU$äÔS²W6W$æÖRÒ"FVçc¥U4U$DôÔ”åÂFVçc¥U4U$äÔR"Ð¢Ò¢Ö÷fRÔ—FVÒÔÆ—FW&ÅF‚G7FvRÔFW7F–æF–öâGF&vWBÔf÷&6P ¢2cRÕ¢–â8þ8*.8;Î8*¾8*N89n8ÎjÚ>[Èþ89^8*ž8:¾888Žz{¾X¹^8~8Þ8n8¾8(žKÙÎ8(¾8 ¢2XXŽ8¾KÙÎ8(¾8Ž8z{¾X¹^8¾ZKiY~8~8þ8Ž8Þ8¾8*.8;Î8*¾8*N89n8ÎxJ8N8î8²–â88jè¾8(¾8 ¢6WBÔf–æÅFe6æ6†÷E–ç2DÆæwVvRF&6†—fU6´–BF&6†—fT6FVv÷'’EföÇVÖRD'V–ÆD–BG6÷W&6Uv÷&¶&öö·2GF&vW@¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvf–æÂæ&6†—fRæ7&VFVBr…¶÷&FW&VEÔ²6´–BÒF&6†—fU6´–C²6FVv÷'’ÒF&6†—fT6FVv÷'“²F&vWD–BÒ„vWBÕF&vWD–Dg&öÔÆVv7•föÇVÖREföÇVÖR“²föÇVÖRÒEföÇVÖS²'V–ÆD–BÒD'V–ÆD–C²F‚ÒGF&vWBÒ¢&WGW&âGF&vW@¢Ò6F6‚°¢2cRÕ¢hú8(®8N8n8^8®8N8.YÎX{®XX>8Î8:Þ8;Î8:¾8988>8*þ8ž8(¾8 ¢G'’²–b…FW7BÕF‚ÔÆ—FW&ÅF‚‚GF&vWB²rç7Fv–ærr’’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚‚GF&vWB²rç7Fv–ærr’Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÒÒ6F6‚²Ð¢F‡&÷r‚.hùX{®yJ…Dn8î8*.8;Î8*¾8*N89n8¾ZKiY~8~8î8~8ó¢"²EòäW†6WF–öâäÖW76vR¢Ð§Ð ¦gVæ7F–öâ&VÖ÷fRÔf–æÄ&6†—fT'F–f7G2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒD6FVv÷'’Â·7G&–æuÒEföÇVÖRÂ·7G&–æuÒD'V–ÆD–B’°¢28:Þ8;Î8:¾8988>8*þi˜.8¾8˜:ŽXˆny¨N8¾8~8Þ8þ8*.8;Î8*¾8*N89n8‚–â8).hè>™šN8ž8(¾8 ¢G'’°¢GF&vWBÒ¦ö–âÕF‚„¦ö–âÕF‚„¦ö–âÕF‚„vWBÔ&6†—fU&ö÷BDÆæwVvR’D6FVv÷'’’EföÇVÖR’D'V–ÆD–@¢f÷&V6‚‚GF‚–â‚GF&vWBÂ‚GF&vWB²rç7Fv–ærr’’’°¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ð¢G&ö÷BÒvWBÔ–çWD†—7F÷'•&ö÷BDÆæwVvP¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚G&ö÷B’°¢f÷&V6‚‚Gv$F—"–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚G&ö÷BÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢f÷&V6‚‚G6äF—"–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚Gv$F—"ägVÆÄæÖRÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢&VÖ÷fRÕ6æ6†÷E–âDÆæwVvRGv$F—"äæÖRG6äF—"äæÖR‚&f–æÂ×Fe÷³Ò"ÖbD'V–ÆD–B¢Ð¢Ð¢Ð¢F7&ö÷BÒ¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’v6öçFVçB×Fbp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚F7&ö÷B’°¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚F7&ö÷BÕ&V7W'6RÔf–ÆRÔf–ÇFW"‚&f–æÂ×Fe÷³Òæ§6öâ"ÖbD'V–ÆD–B’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FbägVÆÄæÖRÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢Ð¢Ð¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâvWBÔf–æÄ&6†—fW2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒD6FVv÷'’’°¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢G66÷RÒ&W6öÇfRÔFö7VÖVçE6µ66÷RG7G'V7GW&RD6FVv÷'’GG'VP¢G6´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒG66÷Rç6´–B’w6´–Bp¢F&6†—fU7F÷&vT–BÒ–b…¶&ööÅÒG66÷Ræ'V–ÇD–â’²·7G&–æuÒG66÷Ræ6FVv÷'’ÒVÇ6R²G6´–BÐ¢G&ö÷BÒ¦ö–âÕF‚„vWBÔ&6†—fU&ö÷BDÆæwVvR’F&6†—fU7F÷&vT–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G&ö÷B’’²&WGW&â‚’Ð¢F÷WBÒ‚¢f÷&V6‚‚GföÄF—"–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚G&ö÷BÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢f÷&V6‚‚F"–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚GföÄF—"ägVÆÄæÖRÔF—&V7F÷'’ÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂ6÷'BÔö&¦V7BæÖRÔFW66VæF–ærÂ6VÆV7BÔö&¦V7BÔf—'7BS’’°¢G'’°¢FÖWFÒ&VBÔ§6öäf–ÆR„¦ö–âÕF‚F"ägVÆÄæÖRvÖWFFFæ§6öâr’FçVÆÀ¢–b‚FçVÆÂÖWFÖWF’²6öçF–çVRÐ¢F÷WB³Ò¶÷&FW&VEÔ°¢'V–ÆD–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖWFv'V–ÆD–BrF"äæÖR¢6´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖWFw6´–BrG6´–B¢6FVv÷'’Ò·7G&–æuÒ„vWBÔFF&÷W'G’FÖWFv6FVv÷'’r…·7G&–æuÒG66÷Ræ6FVv÷'’’¢F&vWD–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖWFwF&vWD–Br„vWBÕF&vWD–Dg&öÔÆVv7•föÇVÖR…·7G&–æuÒGföÄF—"äæÖR’’¢föÇVÖRÒ·7G&–æuÒGföÄF—"äæÖP¢'V–ÇDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖWFv'V–ÇDBrrr¢vT6÷VçBÒ¶–çEÒ„vWBÔFF&÷W'G’FÖWFwvT6÷VçBr¢÷WGWDf–ÆTæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖWFv÷WGWDf–ÆTæÖRrrr¢÷WGWEFe6†#SbÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖWFv÷WGWEFe6†#Sbrrr¢'V–ÇD'’Ò„vWBÔFF&÷W'G’FÖWFv'V–ÇD'’rFçVÆÂ¢F‚Ò·7G&–æuÒF"ägVÆÄæÖP¢Ð¢Ò6F6‚²Ð¢Ð¢Ð¢&WGW&â‚F÷WBÂ6÷'BÔö&¦V7B²·7G&–æuÒEòæ'V–ÇDBÒÔFW66VæF–ær§Ð  ¢2ÒÒÒÒjÚ>[ÈþX{®X©¾88Ž8:ž8;>8+n8*þ8+~8:~8;2…cRÜ*srã"’ÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÐ ¦gVæ7F–öâw&—FRÔf–æÄ¦÷W&æÂ…·7G&–æuÒDÆæwVvRÂD¦÷W&æÂ’°¢FF—"ÒvWBÔf–æÅG&ç67F–öäF—"DÆæwVvP¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢6WBÔæ÷FU&÷W'G’D¦÷W&æÂwWFFVDBr„æWrÔæ÷t—6ò¢w&—FRÔ§6öäf–ÆR„¦ö–âÕF‚FF—"‚'³Òæ§6öâ"Öb·7G&–æuÒD¦÷W&æÂçG&ç67F–öä–B’’D¦÷W&æÀ§Ð¦gVæ7F–öâ6WBÔ¦÷W&æÅ†6R…·7G&–æuÒDÆæwVvRÂD¦÷W&æÂÂ·7G&–æuÒE†6R’°¢2cRÜ*srã#¢†6R8þ8Þ8îXšþKÙÎyJŽ8).8ÎZx¾8(8(¾X˜Þ8Þ8¾i»Ž8þ8.[èÎ8~i»Ž8þ8Ž[êžiz~8~8Þ8®8Nz©>8Îjè¾8(¾8 ¢6WBÔæ÷FU&÷W'G’D¦÷W&æÂw†6RrE†6P¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRD¦÷W&æÀ§Ð¦gVæ7F–öâvWBÔ¦÷W&æÅF&vWB‚D¦÷W&æÂÂ·7G&–æuÒEföÇVÖR’°¢f÷&V6‚‚GB–â„vWBÔ'&’D¦÷W&æÂçF&vWG2’’²–b…·7G&–æuÒGBçföÇVÖRÖWEföÇVÖR’²&WGW&âGBÒÐ¢&WGW&âFçVÆÀ§Ð ¦gVæ7F–öâ76W'BÔf–æÄ'V–ÆE6÷W&6W5Væ6†ævVB‚E6æ6†÷G2Â·7G&–æuµÕÒEföÇVÖW2’°¢2hùX{®yJ…Dn8îZûî‹8¾8®8>8ôW†6VÎ888).8X[iÈž89^8*ž8:¾888;ÎKˆ®8î8+^8*N8+®8;¾i»Niki˜.X‹¾8~z+®Š¨Þ8ž8(¾8 ¢266âÕWFFW28þXZŽy›¾˜Ë$W†6VÎ8).‹[iû¾8~8b7G'V7GW&Ræ§6öâ8(.89n88>8*þ8N8Ž8¾i»Nik8ž8(¾8þ8(8¢2iÈ{X.X{®X©¾8î8þ8>8¾YÎ8n[ø^Šh8þ8®8N8 ¢GF‡2ÒvWBÕF‡0¢G6VVâÒ·Ð¢f÷&V6‚‚Gb–âEföÇVÖW2’°¢G6æÒvWBÔFF&÷W'G’E6æ6†÷G2GbFçVÆÀ¢f÷&V6‚‚Gr–â„vWBÔ'&’„vWBÔFF&÷W'G’G6æw6÷W&6Uv÷&¶&öö·2r‚’’’’°¢F–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’Grwv÷&¶&öö´–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F–B’Ö÷"G6VVâä6öçF–ç4¶W’‚F–B’’²6öçF–çVRÐ¢G6VVå²F–EÒÒGG'VP¢G&VÆF—fUF‚Ò·7G&–æuÒ„vWBÔFF&÷W'G’Grw&VÆF—fUF‚rrr¢FæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’Grvf–ÆTæÖRrG&VÆF—fUF‚¢FgVÆÂÒ¦ö–âÕ6fR…·7G&–æuÒGF‡2ç7V&Ö—76–öäF—"’G&VÆF—fUF€¢G'’²F—FVÒÒvWBÔ—FVÒÔÆ—FW&ÅF‚FgVÆÂÔW'&÷$7F–öâ7F÷Ð¢6F6‚²F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚"FæÖR8ÎhùX{®89^8*ž8:¾888;Î8¾Šh¾8N8¾8(®8î8¾8)>8.XXŽ8µDnKÙÎh‰x«nhX¾8).z+®Š¨Þ8~8n8þ88^8N8""’Ð¢F¶æ÷våF–6·2Ò·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÄÆ7Ew&—FUWF5F–6·2rrr¢F¶æ÷vå6—¦RÒ¶–çCcEÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÅ6—¦RrÓ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¶æ÷våF–6·2’ÖæBF¶æ÷vå6—¦RÖvR’°¢–b…·7G&–æuÒF—FVÒäÆ7Ew&—FUF–ÖUWF2åF–6·2ÖæRF¶æ÷våF–6·2Ö÷"¶–çCcEÒF—FVÒäÆVæwF‚ÖæRF¶æ÷vå6—¦R’°¢F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚"FæÖR8îXX>Xéþz‹þ8Îi»Nik8^8(Î8n8N8î8ž8.XXŽ8¾ZHžhùµDn8).KÙÎh‰8~8n8þ88^8N8""¢Ð¢6öçF–çVP¢Ð¢2iz~88~8;Î8+þ8~8:8+þ88~8;Î8+þ8Î8®8NZNYŽ888Zûî‹89^8*8*N8:³K»n8î88þ88>8+~8:^z+®Š¨Þ8Ž{Šî˜8ž8(¾8 ¢F¶æ÷vä†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Grv7W'&VçDW†6VÄ†6‚rrr’¢FÆ—fT†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ7F&ÆT†6‚FgVÆÂ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F¶æ÷vä†6‚’Ö÷"FÆ—fT†6‚ÖæRF¶æ÷vä†6‚’°¢F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚"FæÖR8îXX>Xéþz‹þ8Îi»Nik8^8(Î8n8N8î8ž8.XXŽ8¾ZHžhùµDn8).KÙÎh‰8~8n8þ88^8N8""¢Ð¢Ð¢Ð§Ð ¦gVæ7F–öâ–çfö¶RÔf–æÄ'V–ÆEG&ç67F–öâ…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒD6FVv÷'’Â·7G&–æuµÕÒEföÇVÖW2’°¢2XÙŽKÙ>X{®X©¾8(.8î8Ž8(8nX{®X©¾8(.8[ø^8®8>8ãiÊÎ8).˜	®8(²ŽXÙŽKÙ>88™©ÎZë>[êžiz~8ÎxJ8Nx«nhX¾8).KÙÎ8(ž8®8Bž8 ¢F6BÒ&WV—&RÕv÷&¶&öö´6FVv÷'’D6FVv÷'¢GF‡2ÒvWBÕF‡0¢Gv÷&·76RÒvWBÕv÷&·76UF‚DÆæwVvP ¢FÆÆ÷vVBÒ„vWBÕföÇVÖTÆ—7BDÆæwVvRÂv†W&RÔö&¦V7B²EòÖæRvæöæRrÒ¢G&WVW7FVBÒ‚EföÇVÖW2Âv†W&RÔö&¦V7B²FÆÆ÷vVBÖ6öçF–ç2EòÒ¢–b‚G&WVW7FVBä6÷VçBÖW’²F‡&÷rµ7—7FVÒä&wVÖVçDW†6WF–öåÓ£¦æWr‚wföÇVÖ^8¾8þiÊÎKÙ>8î8þ8þŠ9Î‹k>8).hÈ~Zé®8~8n8þ88^8N8"r’Ð ¢2cRÜ*srã#¢Zûî‹8òvT6÷VçBâ8âföÇVÖR8î8þ8 ¢2vWBÔf–æÄ'V–ÆD–çWE6æ6†÷B8ò89®8;Î8+Ž8)"æò×vW2&Æö6¶W"8¾8ž8(¾8þ8(8¢2Š9Î‹k>8Îz›®8îjŽK»n8~8Î8î8Ž8(8nX{®X©¾8Þ8Î[‹Ž8¾ZKiY~8~8n8~8î8n8 ¢FÆö6µF‚Ò¦ö–âÕF‚Gv÷&·76R‚&Æö6·5Æf–æÂÖ'V–ÆE÷³ÒæÆö6²"ÖbF6B¢&WGW&â–çfö¶RÕv—F„Æö6²FÆö6µF‚°¢G6æ6†÷G2Ò·Ð¢GF&vWG2Ò‚¢G6¶—VBÒ‚¢2iz.Zé®yZ®Xû~8î˜žyJŽ8ŽXZ‡föÇVÖ^8îXZ^X©¾Y»®Zé®8).87G'V7GW&Ræ§6öâY¹î8îi»Nik8~kˆŽ8î8¾8(¾8 ¢F–æ—F–Å6æ6†÷G2ÒWFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7B¢Ç’ÔFVfVÇDçVÖ&W&–æuW%föÇVÖRDÆæwVvRG7BF6@¢FÆÂÒ¶÷&FW&VEÔ·Ð¢f÷&V6‚‚Gb–âG&WVW7FVB’°¢G6âÒvWBÔf–æÄ'V–ÆD–çWE6æ6†÷BG7BDÆæwVvRGbF6@¢2cRÕ¢8*.8;Î8*¾8*N89nyJŽ8îh8^Z8þ8>8îi˜.x+ž8~Y»®Zé®8ž8(²Ž[èÎ8r7G'V7GW&R8).ŠªÞ8þy»N8^8®8Bž8 ¢G6VVâÒ·Ð¢G7rÒ‚¢f÷&V6‚‚Gr–â„vWBÔ'&’„vWBÔFF&÷W'G’G6âwvW2r‚’’’’°¢Gv–BÒ·7G&–æuÒGrçv÷&¶&öö´–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Gv–B’Ö÷"G6VVâä6öçF–ç4¶W’‚Gv–B’’²6öçF–çVRÐ¢G6VVå²Gv–EÒÒGG'VP¢GrÒ„vWBÔ'&’G7Bçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWGv–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Grä6÷VçBÖW’²6öçF–çVRÐ¢G7r³Ò¶÷&FW&VEÔ°¢v÷&¶&öö´–BÒGv–@¢f–ÆTæÖRÒ·7G&–æuÒGu³Òæf–ÆTæÖP¢&VÆF—fUF‚Ò·7G&–æuÒGu³Òç&VÆF—fUF€¢7W'&VçDW†6VÄÆ7Ew&—FUWF5F–6·2Ò·7G&–æuÒ„vWBÔFF&÷W'G’Gu³Òv7W'&VçDW†6VÄÆ7Ew&—FUWF5F–6·2rrr¢7W'&VçDW†6VÅ6—¦RÒ¶–çCcEÒ„vWBÔFF&÷W'G’Gu³Òv7W'&VçDW†6VÅ6—¦RrÓ¢7W'&VçDW†6VÄ†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Gu³Òv7W'&VçDW†6VÄ†6‚rrr’¢6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gu³ÒvÆ7E&VæFW&VE6æ6†÷D–Brrr¢fW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gu³ÒvÆ7E&VæFW&VEfW'6–öä–Brrr¢6÷W&6T†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Gu³ÒvÆ7E&VæFW&VDW†6VÄ†6‚rrr’¢&VæFW$Vçf—&öæÖVçDf–ævW'&–çBÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gu³Òw&VæFW$Vçf—&öæÖVçDf–ævW'&–çBrrr¢&VæFW$Vçf—&öæÖVçBÒ„vWBÔFF&÷W'G’Gu³Òw&VæFW$Vçf—&öæÖVçBrFçVÆÂ¢Ð¢Ð¢6WBÔæ÷FU&÷W'G’G6âw6÷W&6Uv÷&¶&öö·2r‚G7r¢FÆÅ²GeÒÒG6à¢Ð¢&WGW&âFÆÀ¢Ð¢f÷&V6‚‚Gb–âG&WVW7FVB’°¢G6æÒvWBÔFF&÷W'G’F–æ—F–Å6æ6†÷G2GbFçVÆÀ¢–b‚FçVÆÂÖWG6æÖ÷"¶–çEÒG6æçvT6÷VçBÖÆR’²G6¶—VB³ÒGc²6öçF–çVRÐ¢F&Æö6¶W'2Ò„vWBÔ'&’G6ææ&Æö6¶W'2Âv†W&RÔö&¦V7B²·7G&–æuÒEòæ6öFRÖæRvæò×vW2rÒ¢–b‚F&Æö6¶W'2ä6÷VçBÖwB’²F‡&÷r´–çfÆ–D÷W&F–öäW†6WF–öåÓ£¦æWr‚‚'³ÞûÉ§³Ò"Öb„vWBÕföÇVÖTÆ&VÄf÷$ÖW76vRGb’Â·7G&–æuÒF&Æö6¶W'5³ÒæÖW76vR’’Ð¢G6æ6†÷G5²GeÒÒG6æ ¢GF&vWG2³ÒG`¢Ð¢–b‚GF&vWG2ä6÷VçBÖW’²&WGW&â¶÷&FW&VEÔ²'V–ÇBÒ‚“²6¶—VBÒ‚G6¶—VB“²ÖW76vRÒ~X{®X©¾Zûî‹8Î8.8(®8î8¾8)>8"rÒÐ¢76W'BÔf–æÄ'V–ÆE6÷W&6W5Væ6†ævVBG6æ6†÷G2GF&vWG0 ¢F6ö×÷6W$¦"Ò¦ö–âÕF‚E67&—C¤&ö÷BvÆ–%ÇFf&÷…Å&W÷'EFd6ö×÷6W"æ¦"p¢GFf&÷„¦"Ò¦ö–âÕF‚E67&—C¤&ö÷BvÆ–%ÇFf&÷…ÇFf&÷‚Öæ¦"p¢–b‚Öæ÷B…FW7BÕF‚F6ö×÷6W$¦"’’²F‡&÷ru&W÷'EFd6ö×÷6W"æ¦"8Î8.8(®8î8¾8)>8"rÐ¢–b‚Öæ÷B…FW7BÕF‚GFf&÷„¦"’’²F‡&÷rwFf&÷‚Öæ¦"8Î8.8(®8î8¾8)>8"rÐ ¢GG„–BÒæWrÕ&$–@¢F¦÷W&æÂÒ¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ²G&ç67F–öä–BÒGG„–C²ÆæwVvRÒDÆæwVvS²6FVv÷'’ÒF6@¢7F'FVDBÒæWrÔæ÷t—6ó²†6RÒw&W&VBp¢F&vWG2Ò‚GF&vWG2Âf÷$V6‚Ôö&¦V7B²¶÷&FW&VEÔ²föÇVÖRÒEó²'V–ÆD–BÒ„æWrÕ&$–B“²&6·W7&VFVBÒFfÇ6S²f–ÆU&WÆ6VBÒFfÇ6S²öÆEFd†6‚Òrs²æWuFd†6‚Òrs²W†—7FVBÒFfÇ6S²f–æÅF‚Òrs²&6·WF‚Òrs²FV×F‚ÒrrÒÒ¢&Vf÷&Tf–ævW'&–çG2Ò¶÷&FW&VEÔ·Ð¢öÆEföÇVÖU7FFW2Ò¶÷&FW&VEÔ·Ð¢æWuföÇVÖU7FFW2Ò¶÷&FW&VEÔ·Ð¢Ð¢f÷&V6‚‚Gb–âGF&vWG2’²F¦÷W&æÂæ&Vf÷&Tf–ævW'&–çG5²GeÒÒ·7G&–æuÒG6æ6†÷G5²GeÒæf–ævW'&–çBÐ¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRF¦÷W&æÀ ¢F&6·WF—"Ò¦ö–âÕF‚Gv÷&·76R‚'7FFUÆf–æÂÖ&6·W5Â"²GG„–B¢G'’°¢2RâX{®X©¾XX…Dn8Î™h¾8¾8(Î8n8N8®8N8¾8).8Zûî‹8ž8ž8n8î8Ž8(8nz+®Š¨Ð¢f÷&V6‚‚Gb–âGF&vWG2’°¢GBÒvWBÔ¦÷W&æÅF&vWBF¦÷W&æÂG`¢F÷WDæÖRÒ·7G&–æuÒG6æ6†÷G5²GeÒæ÷WGWDf–ÆTæÖP¢F÷WEF‚Ò¦ö–âÕF‚…·7G&–æuÒGF‡2æ÷WGWDF—"’F÷WDæÖP¢6WBÔæ÷FU&÷W'G’GBvf–æÅF‚rF÷WEF€¢6WBÔæ÷FU&÷W'G’GBvW†—7FVBr…¶&ööÅÒ…FW7BÕF‚ÔÆ—FW&ÅF‚F÷WEF‚’¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚F÷WEF‚’°¢FbÒFçVÆÀ¢G'’²FbÒ´”òäf–ÆUÓ£¤÷Vâ‚F÷WEF‚Â´”òäf–ÆTÖöFUÓ£¤÷VâÂ´”òäf–ÆT66W75Ó£¥&VEw&—FRÂ´”òäf–ÆU6†&UÓ£¤æöæR’Ð¢6F6‚²F‡&÷r.X{®X©¾XXŽ8îhùX{®yJ…Dn8Î™h¾8¾8(Î8n8N8(¾8þ8(Kˆ®i»Ž8Þ8~8Þ8î8¾8)3¢F÷WDæÖR"Ð¢f–æÆÇ’²–b‚Fb’²FbäF—7÷6R‚’ÒÐ¢Ð¢Ð¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRF¦÷W&æÀ ¢2rÓ‚âKˆi˜.89^8*8*N8:¾8Ž{XNx˜€¢f÷&V6‚‚Gb–âGF&vWG2’°¢GBÒvWBÔ¦÷W&æÅF&vWBF¦÷W&æÂG`¢GF×Ò¦ö–âÕF‚…·7G&–æuÒGF‡2æ÷WGWDF—"’‚'æ'V–ÆF–æu÷³Õ÷³Õ÷³'ÒçFb"ÖbGbÂF6BÂGG„–B¢–b…FW7BÕF‚GF×’²&VÖ÷fRÔ—FVÒGF×Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢6WBÔæ÷FU&÷W'G’GBwFV×F‚rGF× ¢G6æÒG6æ6†÷G5²GeÐ¢FÖæ–fW7BÒ¶÷&FW&VEÔ²66†VÖfW'6–öãÓ3²ÆæwVvSÒDÆæwVvS²6FVv÷'“ÒF6C²föÇVÖSÒGc²&ö¦V7D–CÕ·7G&–æuÒG6æç&ö¦V7D–C²–çWDf–ævW'&–çCÕ·7G&–æuÒG6ææf–ævW'&–çC²÷WGWEFcÒGF×²7&VFVDCÔæWrÔæ÷t—6ó²Fö7VÖVçCÒG6ææFö7VÖVçC²vTçVÖ&W#Õ¶÷&FW&VEÔ¶föçCÒt&–Âs¶föçE6—¦SÓƒ¶&÷GFöÕCÓƒ¶f÷&ÖCÒv‡—†VæFVBs¶6÷VçD†–FFVãÒGG'VWÓ²‡—6–6ÅvW3ÒG6æç‡—6–6ÅvW3²vW3ÒG6ææÖæ–fW7EvW2Ð¢FÖæ–fW7EF‚Ò¦ö–âÕF‚…´”òåF…Ó£¤vWEFV×F‚‚’’‚%&W÷'D&–æFW%öf–æÅ÷³Õ÷³Õ÷³'Òæ§6öâ"ÖbGbÂF6BÂGG„–B¢G'’°¢w&—FRÔ§6öäf–ÆRFÖæ–fW7EF‚FÖæ–fW7@¢F¦fÒ&W6öÇfRÔ¦fW†P¢G'VâÒ–çfö¶RÔæF—fT6GW&RF¦f‚rÖ7rÂ"F6ö×÷6W$¦#²GFf&÷„¦""Âu&W÷'EFd6ö×÷6W"rÂrÒÖÖæ–fW7BrÂFÖæ–fW7EF‚’E67&—C¤f–æÄ6ö×÷6UF–ÖV÷WE6V6öæG0¢FW†—BÒ¶–çEÒG'VâæW†—D6öFP¢GFW‡BÒ·7G&–æuÒG'VâçFW‡@¢–b‚G'VâçF–ÖVD÷WB’²F‡&÷r‚uDn8î{YYŽ8Ç³ÞXˆnKº^Xh^8¾{X.8(þ8(®8î8¾8)>8~8~8þ8.X{®X©¾XXŽ8Î88Þ88>88Ž8:þ8;Î8*þKˆ®8î89^8*ž8:¾888;Î8îZNYŽ8þ88N8>8þ8)5>Xh^8î89^8*ž8:¾888;Î8¾X{®X©¾8~8n8þ8n8þ88^8N8"rÖb¶–çEÒ‚E67&—C¤f–æÄ6ö×÷6UF–ÖV÷WE6V6öæG2òc’’Ð¢–b‚FW†—BÖæR’²F‡&÷r%Dd&÷Ž{XNx˜Ž8¾ZKiY~8~8î8~8þ8&W†—CÒFW†—FâGFW‡B"Ð¢Òf–æÆÇ’°¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FÖæ–fW7EF‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FÖæ–fW7EF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ð¢–b‚Öæ÷B…FW7BÕF‚GF×’Ö÷"„vWBÔ—FVÒGF×’äÆVæwF‚ÖÆR’²F‡&÷r~hùX{®yJ…Dn8).KÙÎh‰8~8Þ8î8¾8)>8~8~8þ8"rÐ¢6WBÔæ÷FU&÷W'G’GBvæWuFd†6‚r„æ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbGF×’¢6WBÔæ÷FU&÷W'G’GBvÖæ–fW7BrFÖæ–fW7@¢Ð¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRF¦÷W&æÀ ¢2’âf–ævW'&–çBXhÞz+®Š¨Þ8.ŠªÞXùn888®8î8r7G'V7GW&Ræ§6öâ8)'föÇVÖ^8N8Ž8¾XhÞKùÞZÙŽ8~8®8N8 ¢FgFW%7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢f÷&V6‚‚Gb–âGF&vWG2’°¢FgFW"ÒvWBÔf–æÄ'V–ÆD–çWE6æ6†÷BFgFW%7G'V7GW&RDÆæwVvRGbF6@¢–b…·7G&–æuÒFgFW"æf–ævW'&–çBÖæR·7G&–æuÒF¦÷W&æÂæ&Vf÷&Tf–ævW'&–çG5²GeÒ’°¢F‡&÷ruDnKÙÎh‰KŠÞ8¾89®8;Î8+Žjx¾h‰8î8þ8õDnXZ^X©¾8ÎZHži»N8^8(Î8î8~8þ8.iÈik8îx«nhX¾8~XhÞ[ªnX{®X©¾8~8n8þ88^8N8"p¢Ð¢Ð ¢2â8988>8*þ8*.88>89p¢6WBÔ¦÷W&æÅ†6RDÆæwVvRF¦÷W&æÂv&6·W2Ö7&VFVBp¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚F&6·WF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚F&6·WF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢f÷&V6‚‚Gb–âGF&vWG2’°¢GBÒvWBÔ¦÷W&æÅF&vWBF¦÷W&æÂG`¢–b…¶&ööÅÒGBæW†—7FVB’°¢F&²Ò¦ö–âÕF‚F&6·WF—"‚'³ÒçFb"ÖbGb¢6÷’Ô—FVÒÔÆ—FW&ÅF‚…·7G&–æuÒGBæf–æÅF‚’ÔFW7F–æF–öâF&²Ôf÷&6P¢6WBÔæ÷FU&÷W'G’GBv&6·WF‚rF&°¢6WBÔæ÷FU&÷W'G’GBvöÆEFd†6‚r„æ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbF&²’¢Ð¢6WBÔæ÷FU&÷W'G’GBv&6·W7&VFVBrGG'VP¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRF¦÷W&æÀ¢Ð ¢2â[zî8~i»þ8€¢6WBÔ¦÷W&æÅ†6RDÆæwVvRF¦÷W&æÂw&WÆ6–ærÖf–ÆW2p¢f÷&V6‚‚Gb–âGF&vWG2’°¢GBÒvWBÔ¦÷W&æÅF&vWBF¦÷W&æÂG`¢Ö÷fRÔ—FVÒÔÆ—FW&ÅF‚…·7G&–æuÒGBçFV×F‚’ÔFW7F–æF–öâ…·7G&–æuÒGBæf–æÅF‚’Ôf÷&6P¢6WBÔæ÷FU&÷W'G’GBvf–ÆU&WÆ6VBrGG'VP¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRF¦÷W&æÀ¢Ð¢6WBÔ¦÷W&æÅ†6RDÆæwVvRF¦÷W&æÂvf–ÆW2×&WÆ6VBp ¢2"â7G'V7GW&Ri»NikŽ8:Þ88>8*þXh^8rf–ævW'&–çBXhÞz+®Š¨Þ8öÆBöæWr8).XXŽ8¾i»Ž8Þ{X.8Ž8(²¢6WBÔ¦÷W&æÅ†6RDÆæwVvRF¦÷W&æÂw7G'V7GW&RÖ6öÖÖ—GF–ærp¢F6öÖÖ—BÒWFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7B¢f÷&V6‚‚Gb–âGF&vWG2’°¢FgFW"ÒvWBÔf–æÄ'V–ÆD–çWE6æ6†÷BG7BDÆæwVvRGbF6@¢–b…·7G&–æuÒFgFW"æf–ævW'&–çBÖæR·7G&–æuÒF¦÷W&æÂæ&Vf÷&Tf–ævW'&–çG5²GeÒ’²&WGW&â¶÷&FW&VEÔ²6†ævVBÒGG'VS²föÇVÖRÒGbÒÐ¢Ð¢f÷&V6‚‚Gb–âGF&vWG2’°¢F¶W’ÒvWBÕföÇVÖU7FFT¶W’GbF6@¢FöÆBÒvWBÔFF&÷W'G’G7BçföÇVÖW2F¶W’FçVÆÀ¢F¦÷W&æÂæöÆEföÇVÖU7FFW5²GeÒÒB†–b‚FçVÆÂÖWFöÆB’²FçVÆÂÒVÇ6R²¶÷&FW&VEÔ°¢'V–ÇDf–ævW'&–çBÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBv'V–ÇDf–ævW'&–çBrrr¢7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBw7FGW2rrr¢÷WGWEFbÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBv÷WGWEFbrrr¢Æ7D'V–ÇDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBvÆ7D'V–ÇDBrrr¢7FÆU&V6öç2Ò„vWBÔ'&’„vWBÔFF&÷W'G’FöÆBw7FÆU&V6öç2r‚’’¢ÒÒ¢GBÒvWBÔ¦÷W&æÅF&vWBF¦÷W&æÂG`¢F¦÷W&æÂææWuföÇVÖU7FFW5²GeÒÒ¶÷&FW&VEÔ°¢'V–ÇDf–ævW'&–çBÒ·7G&–æuÒF¦÷W&æÂæ&Vf÷&Tf–ævW'&–çG5²GeÐ¢7FGW2Òv'V–ÇBs²÷WGWEFbÒ·7G&–æuÒGBæf–æÅFƒ²Æ7D'V–ÇDBÒæWrÔæ÷t—6ó²7FÆU&V6öç2Ò‚¢Ð¢Ð¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRF¦÷W&æÀ¢f÷&V6‚‚Gb–âGF&vWG2’°¢F¶W’ÒvWBÕföÇVÖU7FFT¶W’GbF6@¢Gg2ÒvWBÔFF&÷W'G’G7BçföÇVÖW2F¶W’FçVÆÀ¢–b‚FçVÆÂÖWGg2’²Gg2ÒæWrÔV×G•föÇVÖU7FFS²6WBÔæ÷FU&÷W'G’G7BçföÇVÖW2F¶W’Gg2Ð¢FâÒF¦÷W&æÂææWuföÇVÖU7FFW5²GeÐ¢6WBÔæ÷FU&÷W'G’Gg2v'V–ÇDf–ævW'&–çBr…·7G&–æuÒFâæ'V–ÇDf–ævW'&–çB¢6WBÔæ÷FU&÷W'G’Gg2vÆ7D'V–ÇDBr…·7G&–æuÒFâæÆ7D'V–ÇDB¢6WBÔæ÷FU&÷W'G’Gg2v÷WGWEFbr…·7G&–æuÒFâæ÷WGWEFb¢6WBÔæ÷FU&÷W'G’Gg2w7FÆU&V6öç2r‚¢2YÎ8‡7G'V7GW&^8:Þ88>8*þXh^8vf–ævW'&–çN8).z+®Š¨ÞkˆŽ8þ8.XhÞ[ªnXZŽ89®8;Î8+Ž8).‹[iû¾8~8®8N8 ¢6WBÔæ÷FU&÷W'G’Gg2w7FGW2rv'V–ÇBp¢Ð¢&WGW&â¶÷&FW&VEÔ²6†ævVBÒFfÇ6RÐ¢Ð¢–b…¶&ööÅÒF6öÖÖ—Bæ6†ævVB’²F‡&÷ruDnKÙÎh‰KŠÞ8¾89®8;Î8+Žjx¾h‰8î8þ8õDnXZ^X©¾8ÎZHži»N8^8(Î8î8~8þ8.iÈik8îx«nhX¾8~XhÞ[ªnX{®X©¾8~8n8þ88^8N8"rÐ¢6WBÔ¦÷W&æÅ†6RDÆæwVvRF¦÷W&æÂw7G'V7GW&RÖ6öÖÖ—GFVBp ¢22â8*.8;Î8*¾8*N89n8‚–âŽXj®zØ’¢6WBÔ¦÷W&æÅ†6RDÆæwVvRF¦÷W&æÂv&6†—f–ærp¢F'V–ÇBÒ‚¢f÷&V6‚‚Gb–âGF&vWG2’°¢GBÒvWBÔ¦÷W&æÅF&vWBF¦÷W&æÂG`¢G6fVDÖæ–fW7BÒvWBÔFF&÷W'G’GBvÖæ–fW7BrFçVÆÀ¢G6fVDÖæ–fW7EF‚Ò¦ö–âÕF‚Gv÷&·76R‚&W‡÷'G5ÆÖæ–fW7E÷³Õ÷³Òæ§6öâ"ÖbGbÂF6B¢–b‚FçVÆÂÖæRG6fVDÖæ–fW7B’°¢6WBÔæ÷FU&÷W'G’G6fVDÖæ–fW7Bv÷WGWEFbr…·7G&–æuÒGBæf–æÅF‚¢w&—FRÔ§6öäf–ÆRG6fVDÖæ–fW7EF‚G6fVDÖæ–fW7@¢Ð¢2KùÞZÙŽkˆŽ8þX{®X©¾8;¾89®8;Î8+Žjx¾h‰[^jÛN8þ[¸>jÚ.8.X[iÈž89^8*ž8:¾888;Î8Ž8îŠH~Š;Þ8).KÙÎ8(ž8®8N8 ¢F&6†—fRÒrp¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvf–æÂæ'V–ÇBr…¶÷&FW&VEÔ²6FVv÷'’ÒF6C²föÇVÖRÒGc²'V–ÆD–BÒ·7G&–æuÒGBæ'V–ÆD–C²÷WGWEFbÒ·7G&–æuÒGBæf–æÅF‚Ò¢F'V–ÇB³Ò¶÷&FW&VEÔ²föÇVÖRÒGc²6FVv÷'’ÒF6C²÷WGWEFbÒ·7G&–æuÒGBæf–æÅFƒ²Öæ–fW7EF‚ÒG6fVDÖæ–fW7EFƒ²'V–ÆD–BÒ·7G&–æuÒGBæ'V–ÆD–C²&6†—fUF‚ÒF&6†—fS²–çWDf–ævW'&–çBÒ·7G&–æuÒF¦÷W&æÂæ&Vf÷&Tf–ævW'&–çG5²GeÒÐ¢Ð¢6WBÔ¦÷W&æÅ†6RDÆæwVvRF¦÷W&æÂv6ö×ÆWFVBp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚F&6·WF—"’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚F&6·WF—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢&WGW&â¶÷&FW&VEÔ²'V–ÇBÒ‚F'V–ÇB“²6¶—VBÒ‚G6¶—VB“²G&ç67F–öä–BÒGG„–BÐ¢Ò6F6‚°¢f÷&V6‚‚GB–â„vWBÔ'&’F¦÷W&æÂçF&vWG2’’²&VÖ÷fRÔf–æÄ&6†—fT'F–f7G2DÆæwVvRF6B…·7G&–æuÒGBçföÇVÖR’…·7G&–æuÒGBæ'V–ÆD–B’Ð¢&W7F÷&RÔf–æÅG&ç67F–öâDÆæwVvRF¦÷W&æÀ¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRvf–æÂæ'V–ÆBæf–ÆVBr…¶÷&FW&VEÔ²6FVv÷'’ÒF6C²G&ç67F–öä–BÒGG„–C²ÖW76vRÒEòäW†6WF–öâäÖW76vRÒ¢F‡&÷p¢Ð¢Ð§Ð ¦gVæ7F–öâ&W7F÷&RÔf–æÅG&ç67F–öâ…·7G&–æuÒDÆæwVvRÂD¦÷W&æÂ’°¢2cRÜ*t"Ó¢[êžiz~XŠNZé®8þ89^8:ž8+8~8þ8®8þZéþ89^8*8*N8:¾8î88þ88>8+~8:^8).jÚ>8Ž8ž8(¾8 ¢2f–ÆU&WÆ6VC×G'VR8).i»Ž8þX˜Þ8¾‰Þ88(¾z©>8Î8.8(¾8þ8(8 ¢G'’°¢G†6RÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦÷W&æÂw†6Rrrr¢f÷&V6‚‚GB–â„vWBÔ'&’D¦÷W&æÂçF&vWG2’’°¢Ff–æÂÒ·7G&–æuÒ„vWBÔFF&÷W'G’GBvf–æÅF‚rrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Ff–æÂ’’²6öçF–çVRÐ¢GF×Ò·7G&–æuÒ„vWBÔFF&÷W'G’GBwFV×F‚rrr¢–b‚GF×ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚GF×’’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF×Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢FæWt†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’GBvæWuFd†6‚rrr’¢FöÆD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’GBvöÆEFd†6‚rrr’¢FW†—7FVBÒ¶&ööÅÒ„vWBÔFF&÷W'G’GBvW†—7FVBrFfÇ6R¢F&6·WÒ·7G&–æuÒ„vWBÔFF&÷W'G’GBv&6·WF‚rrr¢F7W'&VçBÒrp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚Ff–æÂ’²F7W'&VçBÒæ÷&ÖÆ—¦RÔf–ÆT†6‚„æWrÕ6†#SbFf–æÂ’Ð ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçB’’°¢–b‚Öæ÷BFW†—7FVB’²6öçF–çVRÒ2XX>8¾8(žxJ8þ8K¸®8(.xJ8@¢–b‚F&6·WÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚F&6·W’’²6÷’Ô—FVÒÔÆ—FW&ÅF‚F&6·WÔFW7F–æF–öâFf–æÂÔf÷&6RÐ¢6öçF–çVP¢Ð¢–b‚FæWt†6‚ÖæBF7W'&VçBÖWFæWt†6‚’°¢–b‚Öæ÷BFW†—7FVB’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚Ff–æÂÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢VÇ6V–b‚F&6·WÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚F&6·W’’²6÷’Ô—FVÒÔÆ—FW&ÅF‚F&6·WÔFW7F–æF–öâFf–æÂÔf÷&6RÐ¢6öçF–çVP¢Ð¢–b‚FöÆD†6‚ÖæBF7W'&VçBÖWFöÆD†6‚’²6öçF–çVRÒ2iÊ®[zî8~i»þ8€¢28ž8(Î8Ž8(.Kˆˆ{N8~8®8C¢ˆz®X¹^8~Kˆ®i»Ž8Þ8~8®8N8 ¢6WBÔæ÷FU&÷W'G’D¦÷W&æÂw†6RrvÖçVÂ×&V6÷fW'’×&WV—&VBp¢6WBÔæ÷FU&÷W'G’D¦÷W&æÂvÖçVÅ&V6÷fW'•&V6öâr‚.X{®X©¾XX…Dn8ÎŠ‰Ž˜Ë.8^8(Î8þ8ž8îx«nhX¾8Ž8(.Kˆˆ{N8~8î8¾8)3¢"²Ff–æÂ¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRD¦÷W&æÀ¢&WGW&à¢Ð¢27G'V7GW&R8î[{¾8Þh‹¾8rŽZûî‹föÇVÖ^8îš^yºî88¢–b„‚w7G'V7GW&RÖ6öÖÖ—GF–ærrÂw7G'V7GW&RÖ6öÖÖ—GFVBrÂv&6†—f–ærr’Ö6öçF–ç2G†6R’°¢FöÆG2ÒvWBÔFF&÷W'G’D¦÷W&æÂvöÆEföÇVÖU7FFW2rFçVÆÀ¢FæWw2ÒvWBÔFF&÷W'G’D¦÷W&æÂvæWuföÇVÖU7FFW2rFçVÆÀ¢–b‚FçVÆÂÖæRFöÆG2’°¢F6BÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦÷W&æÂv6FVv÷'’rrr¢WFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7B¢f÷&V6‚‚GB–â„vWBÔ'&’D¦÷W&æÂçF&vWG2’’°¢GbÒ·7G&–æuÒGBçföÇVÖP¢F¶W’ÒvWBÕföÇVÖU7FFT¶W’GbF6@¢Gg2ÒvWBÔFF&÷W'G’G7BçföÇVÖW2F¶W’FçVÆÀ¢–b‚FçVÆÂÖWGg2’²6öçF–çVRÐ¢FæWu7FFRÒvWBÔFF&÷W'G’FæWw2GbFçVÆÀ¢–b‚FçVÆÂÖæRFæWu7FFRÖæB·7G&–æuÒ„vWBÔFF&÷W'G’Gg2v'V–ÇDf–ævW'&–çBrrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’FæWu7FFRv'V–ÇDf–ævW'&–çBrrr’’²6öçF–çVRÐ¢FöÆE7FFRÒvWBÔFF&÷W'G’FöÆG2GbFçVÆÀ¢–b‚FçVÆÂÖWFöÆE7FFR’²6öçF–çVRÐ¢6WBÔæ÷FU&÷W'G’Gg2v'V–ÇDf–ævW'&–çBr…·7G&–æuÒ„vWBÔFF&÷W'G’FöÆE7FFRv'V–ÇDf–ævW'&–çBrrr’¢6WBÔæ÷FU&÷W'G’Gg2w7FGW2r…·7G&–æuÒ„vWBÔFF&÷W'G’FöÆE7FFRw7FGW2rvæ÷BÖ'V–ÇBr’¢6WBÔæ÷FU&÷W'G’Gg2v÷WGWEFbr…·7G&–æuÒ„vWBÔFF&÷W'G’FöÆE7FFRv÷WGWEFbrrr’¢6WBÔæ÷FU&÷W'G’Gg2vÆ7D'V–ÇDBr…·7G&–æuÒ„vWBÔFF&÷W'G’FöÆE7FFRvÆ7D'V–ÇDBrrr’¢6WBÔæ÷FU&÷W'G’Gg2w7FÆU&V6öç2r„vWBÔ'&’„vWBÔFF&÷W'G’FöÆE7FFRw7FÆU&V6öç2r‚’’¢Ð¢ÒÂ÷WBÔçVÆÀ¢Ð¢Ð¢6WBÔæ÷FU&÷W'G’D¦÷W&æÂw†6Rrw&öÆÆVBÖ&6²p¢w&—FRÔf–æÄ¦÷W&æÂDÆæwVvRD¦÷W&æÀ¢2[{¾8Þh‹¾8~8¾KÛþ8N{X.8Ž8þizuDn8988>8*þ8*.88>89~8).jè¾8^8®8N8 ¢&VÖ÷fRÔf–æÅG&ç67F–öä&6·WF—"DÆæwVvR…·7G&–æuÒ„vWBÔFF&÷W'G’D¦÷W&æÂwG&ç67F–öä–Brrr’¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâ&V6÷fW"Ôf–æÅG&ç67F–öç2…·7G&–æuÒDÆæwVvR’°¢2cS¢‹[~X¹^i˜.[êžiz~8("f–æÂÖ'V–ÆB8:Þ88>8*þ8).Xùn8(²ŽŠH~i[8+^8;Î898;Î8ÎYÎ8Ž[êžiz~8).‹[8(ž8¾8®8Bž8 ¢G'’°¢FF—"ÒvWBÔf–æÅG&ç67F–öäF—"DÆæwVvP¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&âÐ¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢F¢Ò&VBÔ§6öäf–ÆRFbägVÆÄæÖRFçVÆÀ¢–b‚FçVÆÂÖWF¢’²6öçF–çVRÐ¢G†6RÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢w†6Rrrr¢–b„‚v6ö×ÆWFVBrÂw&öÆÆVBÖ&6²r’Ö6öçF–ç2G†6R’°¢2ZèÎK¨nkˆŽ8òþ[{¾8Þh‹¾8~kˆŽ8þ8~8þ8988>8*þ8*.88>89~8þKˆÞŠh8.X˜ÞY¹îX˜®™šN8¾ZKiY~8~8n8N8n8(.‹[~X¹^i˜.8¾XhÞŠšnŠÎ8ž8(¾8 ¢28+Ž8:>8;Î88®8:¾iÊÎKÙ>8ó3iz^jè¾8~8ÖçVÂ×&V6÷fW'’×&WV—&VB8þh¸^[Ù>ˆ^8Îz+®Š¨Þ8ž8(¾8î8~jè¾8ž8 ¢G'’°¢&VÖ÷fRÔf–æÅG&ç67F–öä&6·WF—"DÆæwVvR…·7G&–æuÒ„vWBÔFF&÷W'G’F¢wG&ç67F–öä–BrFbä&6TæÖR’¢–b‚FbäÆ7Ew&—FUF–ÖUWF2ÖÇB´FFUF–ÖUÓ£¥WF4æ÷räFDF—2‚Ó3’’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚FbägVÆÄæÖRÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢Ð¢Ò6F6‚²Ð¢6öçF–çVP¢Ð¢–b‚G†6RÖWvÖçVÂ×&V6÷fW'’×&WV—&VBr’²6öçF–çVRÐ¢F6BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢v6FVv÷'’rrr¢FÆö6µF‚Ò¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’‚&Æö6·5Æf–æÂÖ'V–ÆE÷³ÒæÆö6²"ÖbF6B¢F‚ÒG'’Ô7V—&TÆö6´†æFÆRFÆö6µF€¢–b‚FçVÆÂÖWF‚’²6öçF–çVRÐ¢G'’°¢2cRÕ‚3‚“¢˜	®[‹Ž8îZKiY~{XÎ‹zò„–çfö¶RÔf–æÄ'V–ÆEG&ç67F–öâ8â6F6‚ž8ŽYÎ8Ž8þ8¢2‹[~X¹^i˜.[êžiz~8~8(.˜:ŽXˆny¨N8¾8~8Þ8þ8*.8;Î8*¾8*N89b÷–â8).hè>™šN8~8n8¾8(ž[{¾8Þh‹¾8ž8 ¢2&6†—f–ærKŠÞ8¾[Ë~X‹n{X.K¨n8ž8(¾8Ž8Kˆ˜:Ž8î[{¾888*.8;Î8*¾8*N89b÷–â8Îjè¾8(®87G'V7GW&R88[{¾8Þh‹¾8(¾8 ¢f÷&V6‚‚GB–â„vWBÔ'&’F¢çF&vWG2’’²&VÖ÷fRÔf–æÄ&6†—fT'F–f7G2DÆæwVvRF6B…·7G&–æuÒGBçföÇVÖR’…·7G&–æuÒGBæ'V–ÆD–B’Ð¢&W7F÷&RÔf–æÅG&ç67F–öâDÆæwVvRF ¢Òf–æÆÇ’²&VÆV6RÔÆö6´†æFÆRF‚Ð¢Ð¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâvWBÕföÇVÖTÆ&VÄf÷$ÖW76vR…·7G&–æuÒEföÇVÖR’°¢7v—F6‚‚EföÇVÖR’°¢v¦ÖÖ–âr²&WGW&â~iÊÎKÙ2rÒv¦ÖVæF—‚r²&WGW&â~Š9Î‹k2rÐ¢vVâÖÖ–âr²&WGW&âtÖ–ârÒvVâÖVæF—‚r²&WGW&âtVæF—‚rÐ¢FVfVÇB²&WGW&âEföÇVÖRÐ¢Ð§Ð  ¢2ÒÒÒÒ’Š9ÎXª’…cR’ÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÒÐ ¦gVæ7F–öâvWBÔ†—7F÷'•F–ÖVÆ–æR…·7G&–æuÒDÆæwVvRÂ¶–çEÒDÆ–Ö—B’°¢G'’°¢FF—"Ò¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’v†—7F÷'•ÆWfVçG2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²&WGW&â‚’Ð¢F÷WBÒ‚¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂ6÷'BÔö&¦V7BæÖRÔFW66VæF–ærÂ6VÆV7BÔö&¦V7BÔf—'7BDÆ–Ö—B’’°¢G'’°¢F¢Ò&VBÔ§6öäf–ÆRFbägVÆÄæÖRFçVÆÀ¢–b‚FçVÆÂÖæRF¢’²F÷WB³ÒF¢Ð¢Ò6F6‚²Ð¢Ð¢&WGW&âF÷W@¢Ò6F6‚²&WGW&â‚’Ð§Ð  ¦gVæ7F–öâvWBÔ6öçFVçEFe6†VWD–æFW‚…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–B’°¢Gv÷&·76RÒvWBÕv÷&·76UF‚DÆæwVvP¢G6fUfW'6–öä–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEfW'6–öä–BwfW'6–öä–Bp¢FF—"ÒvWBÔ6öçFVçEFefW'6–öäF—"Gv÷&·76REv÷&¶&öö´–BG6fUfW'6–öä–@¢F66†T¶W’Ò…´”òåF…Ó£¤vWDgVÆÅF‚‚FF—"’’åFôÆ÷vW$–çf&–çB‚¢F66†VBÒE67&—C¤6öçFVçEFe6†VWD–æFW„66†U²F66†T¶W•Ð¢–b‚FçVÆÂÖæRF66†VB’°¢F66†VD–æFW‚ÒvWBÔFF&÷W'G’F66†VBv–æFW‚r·Ð¢FÆÄf–ÆW4W†—7BÒGG'VP¢f÷&V6‚‚F66†VEF‚–â‚F66†VD–æFW‚åfÇVW2’’°¢–b‚Öæ÷B…FW7BÔf–ÆTW†—7G46ö×B…·7G&–æuÒF66†VEF‚’’’²FÆÄf–ÆW4W†—7BÒFfÇ6S²'&V²Ð¢Ð¢2âV×G’66†VBvVæW&F–öâ†2æòf–ÆRv†÷6RW†—7FVæ6R6â&÷fRF†@¢2—G2&VçB7F–ÆÂW†—7G2Â6òfÆ–FFRF†RF—&V7F÷'’–âF†B66Rà¢–b‚F66†VD–æFW‚ä6÷VçBÖWÖæBÖæ÷B…FW7BÔF—&V7F÷'”W†—7G46ö×BFF—"’’²FÆÄf–ÆW4W†—7BÒFfÇ6RÐ¢–b‚FÆÄf–ÆW4W†—7B’²&WGW&âF66†VD–æFW‚Ð¢Ð¢–b‚Öæ÷B…FW7BÔF—&V7F÷'”W†—7G46ö×BFF—"’’°¢·fö–EÒE67&—C¤6öçFVçEFe6†VWD–æFW„66†Rå&VÖ÷fR‚F66†T¶W’¢&WGW&â·Ð¢Ð¢G7F×Ò´”òäF—&V7F÷'•Ó£¤vWDÆ7Ew&—FUF–ÖUWF2‚FF—"’åF–6·0¢G&ö÷BÒ´”òåF…Ó£¤vWDgVÆÅF‚‚Gv÷&·76R¢–b‚Öæ÷BG&ö÷BäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’’²G&ö÷B³Ò´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"Ð¢F–æFW‚Ò·Ð¢f÷&V6‚‚Ff–ÆR–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢çFbrÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢FgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚Ff–ÆRägVÆÄæÖR¢–b‚Öæ÷BFgVÆÂå7F'G5v—F‚‚G&ö÷BÂµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’’²6öçF–çVRÐ¢F¶W’Ò…·7G&–æuÒFf–ÆRä&6TæÖR’åFôÆ÷vW$–çf&–çB‚¢–b‚Öæ÷BF–æFW‚ä6öçF–ç4¶W’‚F¶W’’ÖæB…FW7BÔf–ÆTW†—7G46ö×BFgVÆÂ’’²F–æFW…²F¶W•ÒÒFgVÆÂÐ¢Ð¢–b‚Öæ÷BE67&—C¤6öçFVçEFe6†VWD–æFW„66†Rä6öçF–ç4¶W’‚F66†T¶W’’Öæ@¢E67&—C¤6öçFVçEFe6†VWD–æFW„66†Rä6÷VçBÖvRE67&—C¤6öçFVçEFe6†VWD–æFW„66†TÆ–Ö—B’°¢FöÆFW7D¶W’Ò‚E67&—C¤6öçFVçEFe6†VWD–æFW„66†Rä¶W—2•³Ð¢–b‚FçVÆÂÖæRFöÆFW7D¶W’’²·fö–EÒE67&—C¤6öçFVçEFe6†VWD–æFW„66†Rå&VÖ÷fR‚FöÆFW7D¶W’’Ð¢Ð¢E67&—C¤6öçFVçEFe6†VWD–æFW„66†U²F66†T¶W•ÒÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²7F×ÒG7F×²–æFW‚ÒF–æFW‚Ð¢&WGW&âF–æFW€§Ð  ¦gVæ7F–öâ&W6öÇfRÔ6öçFVçEFe6†VWEF„g&öÔ–æFW‚‚D–æFW‚Â·7G&–æuÒE6†VWDæÖR’°¢–b‚FçVÆÂÖWD–æFW‚Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6†VWDæÖR’’²&WGW&ârrÐ¢2†6†VB7F÷&vRæÖW2&R6fRf÷"WfW'’W†6VÂ6†VWBæÖRâÆVv7’6æF–FFW2¶VW ¢2&Wf–÷W6Ç’&VæFW&VBçVÖW&–2ô44”’v÷&¶&öö·2&VF&ÆRGW&–ærâWw&FRà¢f÷&V6‚‚F6æF–FFR–â€¢„vWBÕv÷&·6†VWE7F÷&vU7FVÒE6†VWDæÖR’À¢E6†VWDæÖRÀ¢…·&VvW…Ó£¥&WÆ6R‚E6†VWDæÖRÂuµãÓ”Õ¦×¥Ò²rÂrÒr’¢’’°¢F¶W’Ò…·7G&–æuÒF6æF–FFR’åFôÆ÷vW$–çf&–çB‚¢–b‚D–æFW‚ä6öçF–ç4¶W’‚F¶W’’’²&WGW&â·7G&–æuÒD–æFW…²F¶W•ÒÐ¢Ð¢&WGW&ârp§Ð¦gVæ7F–öâ&W6öÇfRÔ6öçFVçEFe6†VWEF„W†7B…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒE6†VWDæÖR’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EfW'6–öä–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6†VWDæÖR’’²&WGW&ârrÐ¢F–æFW‚ÒvWBÔ6öçFVçEFe6†VWD–æFW‚DÆæwVvREv÷&¶&öö´–BEfW'6–öä–@¢&WGW&â…&W6öÇfRÔ6öçFVçEFe6†VWEF„g&öÔ–æFW‚F–æFW‚E6†VWDæÖR§Ð ¦gVæ7F–öâvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEfW'6–öä–B’°¢G&W7VÇBÒ¶÷&FW&VEÔ°¢fW'6–öä–BÒ·7G&–æuÒEfW'6–öä–@¢f—7VÄ†6„f–Æ&ÆRÒFfÇ6P¢6öçFVçEFdf–Æ&ÆRÒFfÇ6P¢&VG’ÒFfÇ6P¢Ö—76–æu6†VWG2Ò‚¢&V6öâÒrp¢Ð¢F†6†W2ÒvWBÕf—7VÄ†6†W2DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–BEfW'6–öä–@¢–b‚FçVÆÂÖWF†6†W2’°¢G&W7VÇBç&V6öâÒ~yK¾X8þ88þ88>8+~8:^8ÎKùÞZÙŽ8^8(Î8n8N8î8¾8)>8"p¢&WGW&â·67W7FöÖö&¦V7EÒG&W7VÇ@¢Ð¢G&W7VÇBçf—7VÄ†6„f–Æ&ÆRÒGG'VP¢G6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F†6†W2w6†VWG2r‚’’¢–b‚G6†VWG2ä6÷VçBÖW’°¢G&W7VÇBç&V6öâÒ~jùN‹È>Zûî‹8+~8;Î88Ž8îyK¾X8þ88þ88>8+~8:^8Î8.8(®8î8¾8)>8"p¢&WGW&â·67W7FöÖö&¦V7EÒG&W7VÇ@¢Ð¢2x˜Ž88~8*>8:Î8*þ88Ž8:®8îX‰~hÉž8Ži»Niki˜.X‹¾z+®Š¨Þ8óY¹î88ŠÎ8N8XZŽ8+~8;Î88Ž8).8:Þ8;Î8*¾8:¾{J.[É^8~xZ~YŽ8ž8(¾8 ¢GFd–æFW‚ÒvWBÔ6öçFVçEFe6†VWD–æFW‚DÆæwVvREv÷&¶&öö´–BEfW'6–öä–@¢FÖ—76–ærÒ‚¢f÷&V6‚‚G6†VWB–âG6†VWG2’°¢FæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBw6†VWDæÖRrrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’’°¢FÖ—76–ær³Òu·6†VWDæÖRÖ—76–æuÒp¢6öçF–çVP¢Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚…&W6öÇfRÔ6öçFVçEFe6†VWEF„g&öÔ–æFW‚GFd–æFW‚FæÖR’’’²FÖ—76–ær³ÒFæÖRÐ¢Ð¢G&W7VÇBæÖ—76–æu6†VWG2Ò‚FÖ—76–ær¢G&W7VÇBæ6öçFVçEFdf–Æ&ÆRÒ‚FÖ—76–ærä6÷VçBÖW¢G&W7VÇBç&VG’Ò…¶&ööÅÒG&W7VÇBçf—7VÄ†6„f–Æ&ÆRÖæB¶&ööÅÒG&W7VÇBæ6öçFVçEFdf–Æ&ÆR¢–b‚Öæ÷B¶&ööÅÒG&W7VÇBç&VG’’°¢G&W7VÇBç&V6öâÒB†–b‚FÖ—76–ærä6÷VçBÖwB’²~YÎ8Ž8:Î8;>888:®8;>8+K‰nKº>8æ6öçFVçBDn8ÎKˆÞ‹k>8~8n8N8î8ž8"rÒVÇ6R²v6öçFVçBDn8ÎKùÞZÙŽ8^8(Î8n8N8î8¾8)>8"rÒ¢Ð¢&WGW&â·67W7FöÖö&¦V7EÒG&W7VÇ@§Ð  ¦gVæ7F–öâvWBÔÆö6Å6æ6†÷E7VÖÖ'”66†UF‚…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢&WGW&â„¦ö–âÕF‚E67&—C¤Æö6Å'VçF–ÖT66†U&ö÷B‚'6æ6†÷G2×³Ò×³Òæ§6öâ"ÖbDÆæwVvRÂG6fUv÷&¶&öö´–B’§Ð ¦gVæ7F–öâ6ÆV"Õ6æ6†÷E7VÖÖ'”66†R…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢F66†T¶W’Ò‚DÆæwVvR²wÂr²G6fUv÷&¶&öö´–B’åFôÆ÷vW$–çf&–çB‚¢·fö–EÒE67&—C¥6æ6†÷E7VÖÖ'”66†Rå&VÖ÷fR‚F66†T¶W’¢·fö–EÒE67&—C¥6æ6†÷D–D66†Rå&VÖ÷fR‚F66†T¶W’¢G'’°¢GF‚ÒvWBÔÆö6Å6æ6†÷E7VÖÖ'”66†UF‚DÆæwVvRG6fUv÷&¶&öö´–@¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚GF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢Ò6F6‚²Ð§Ð ¦gVæ7F–öâvWBÕ6æ6†÷E7VÖÖ'”VçG'’…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Bp¢FÒÒvWBÕ6æ6†÷DÖæ–fW7BDÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–@¢–b‚FçVÆÂÖWFÒ’²&WGW&âFçVÆÂÐ¢G7FFRÒvWBÕ6æ6†÷E6÷W&6U7FFRDÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–@¢GfW'6–öç2Ò„vWBÕ&VæFW%fW'6–öä–G2DÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–B¢G&VfW'&VBÒrs²Gf—7VÄf–Æ&ÆRÒFfÇ6S²F6öçFVçDf–Æ&ÆRÒFfÇ6P¢G&V6öâÒuDnKÙÎh‰kˆŽ8þ8îjùN‹È>Xúþˆ;Þ8®x˜Ž8Î8.8(®8î8¾8)>8"p¢f÷&V6‚‚GfW'6–öä–B–â‚GfW'6–öç2Â6÷'BÔö&¦V7BÔFW66VæF–ær’’°¢Ff–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–B…·7G&–æuÒGfW'6–öä–B¢–b…¶&ööÅÒFf–Æ&–Æ—G’çf—7VÄ†6„f–Æ&ÆR’²Gf—7VÄf–Æ&ÆRÒGG'VRÐ¢–b…¶&ööÅÒFf–Æ&–Æ—G’æ6öçFVçEFdf–Æ&ÆR’²F6öçFVçDf–Æ&ÆRÒGG'VRÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&VfW'&VB’ÖæB¶&ööÅÒFf–Æ&–Æ—G’ç&VG’’°¢G&VfW'&VBÒ·7G&–æuÒGfW'6–öä–C²G&V6öâÒrp¢ÒVÇ6V–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&VfW'&VB’ÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒFf–Æ&–Æ—G’ç&V6öâ’’°¢G&V6öâÒ·7G&–æuÒFf–Æ&–Æ—G’ç&V6öà¢Ð¢Ð¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢6æ6†÷D–BÒG6fU6æ6†÷D–C²FWFV7FVDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÒvFWFV7FVDBrrr¢6GW&U&V6öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÒv6GW&U&V6öârrr¢6÷W&6T†6‚Ò·7G&–æuÒ„vWBÔFF&÷W'G’FÒw6÷W&6T†6‚rrr¢6÷W&6U&WF–æVBÒ¶&ööÅÒG7FFRç6÷W&6U&WF–æV@¢–ç2Ò„vWBÕ6æ6†÷E–ç2DÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–B¢&VæFW%fW'6–öä–G2Ò‚GfW'6–öç2“²f—7VÄ6ö×&U&VG’Ò‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&VfW'&VB’¢&VfW'&VEfW'6–öä–BÒG&VfW'&VC²f—7VÄ†6„f–Æ&ÆRÒGf—7VÄf–Æ&ÆP¢6öçFVçEFdf–Æ&ÆRÒF6öçFVçDf–Æ&ÆS²Væf–Æ&ÆU&V6öâÒG&V6öà¢Ð§Ð ¦gVæ7F–öâWFFRÕ6æ6†÷E7VÖÖ'”66†TVçG'’…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢2DnKÙÎh‰8(G–îZHži»N8~[Û™ûþ8ž8(¾˜	®[‹ƒx˜Ž888).[zî8~i»þ8Ž8[^jÛNyK¾™Ú.8sCx˜Ž8).XhÞ‹[iû¾8~8®8N8 ¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Bp¢F†—7F÷'”F—"ÒvWBÕv÷&¶&öö´†—7F÷'”F—"DÆæwVvRG6fUv÷&¶&öö´–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚F†—7F÷'”F—"’’²6ÆV"Õ6æ6†÷E7VÖÖ'”66†RDÆæwVvRG6fUv÷&¶&öö´–C²&WGW&âFfÇ6RÐ¢F66†T¶W’Ò‚DÆæwVvR²wÂr²G6fUv÷&¶&öö´–B’åFôÆ÷vW$–çf&–çB‚¢G&V6÷&BÒE67&—C¥6æ6†÷E7VÖÖ'”66†U²F66†T¶W•Ð¢FÆö6ÅF‚ÒvWBÔÆö6Å6æ6†÷E7VÖÖ'”66†UF‚DÆæwVvRG6fUv÷&¶&öö´–@¢–b‚FçVÆÂÖWG&V6÷&B’°¢G'’²–b…FW7BÕF‚ÔÆ—FW&ÅF‚FÆö6ÅF‚’²G&V6÷&BÒ&VBÔ§6öäf–ÆRFÆö6ÅF‚FçVÆÂÒÒ6F6‚²G&V6÷&BÒFçVÆÂÐ¢Ð¢FVçG'’ÒvWBÕ6æ6†÷E7VÖÖ'”VçG'’DÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–@¢–b‚FçVÆÂÖWFVçG'’’²6ÆV"Õ6æ6†÷E7VÖÖ'”66†RDÆæwVvRG6fUv÷&¶&öö´–C²&WGW&âFfÇ6RÐ ¢2ikŠhþx˜Ž‹ûÞXª8(NiÉþ™™i[Nyn8(.XøÞiŠ8ž8(¾8þ8(8‹»Þ˜xþ8®x˜„”NKˆŠj~888óY¹îz+®Š¨Þ8ž8(¾8 ¢·fö–EÒE67&—C¥6æ6†÷D–D66†Rå&VÖ÷fR‚F66†T¶W’¢G&V6VçD–G2Ò‚„vWBÕ6æ6†÷D–G2DÆæwVvRG6fUv÷&¶&öö´–B’Â6VÆV7BÔö&¦V7BÔÆ7BC¢FÖÒ·Ð¢f÷&V6‚‚FöÆB–â„vWBÔ'&’„vWBÔFF&÷W'G’G&V6÷&Bw7VÖÖ&–W2r‚’’’’°¢FöÆD–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’FöÆBw6æ6†÷D–Brrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FöÆD–B’’²FÖ²FöÆD–EÒÒFöÆBÐ¢Ð¢FÖ²G6fU6æ6†÷D–EÒÒFVçG'¢G7VÖÖ&–W2Ò‚¢f÷&V6‚‚F–B–â‚G&V6VçD–G2Â6÷'BÔö&¦V7BÔFW66VæF–ær’’°¢F¶W’Ò·7G&–æuÒF–@¢–b‚Öæ÷BFÖä6öçF–ç4¶W’‚F¶W’’’°¢6ÆV"Õ6æ6†÷E7VÖÖ'”66†RDÆæwVvRG6fUv÷&¶&öö´–@¢&WGW&âFfÇ6P¢Ð¢G7VÖÖ&–W2³ÒFÖ²F¶W•Ð¢Ð¢G7F×Ò´”òäF—&V7F÷'•Ó£¤vWDÆ7Ew&—FUF–ÖUWF2‚F†—7F÷'”F—"’åF–6·0¢GWFFVBÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²66†VÖfW'6–öâÒ#²7F×ÒG7F×²7VÖÖ&–W2Ò‚G7VÖÖ&–W2“²6fVDBÒæWrÔæ÷t—6òÐ¢E67&—C¥6æ6†÷E7VÖÖ'”66†U²F66†T¶W•ÒÒGWFFV@¢G'’²w&—FRÔ§6öäf–ÆRFÆö6ÅF‚GWFFVBÒ6F6‚²Ð¢&WGW&âGG'VP§Ð ¦gVæ7F–öâvWBÕ6æ6†÷E7VÖÖ&–W2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Ev÷&¶&öö´–B’’²&WGW&â‚’Ð¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢F†—7F÷'”F—"ÒvWBÕv÷&¶&öö´†—7F÷'”F—"DÆæwVvRG6fUv÷&¶&öö´–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚F†—7F÷'”F—"’’²&WGW&â‚’Ð¢G7F×Ò´”òäF—&V7F÷'•Ó£¤vWDÆ7Ew&—FUF–ÖUWF2‚F†—7F÷'”F—"’åF–6·0¢F66†T¶W’Ò‚DÆæwVvR²wÂr²G6fUv÷&¶&öö´–B’åFôÆ÷vW$–çf&–çB‚¢FÖVÖ÷'’ÒE67&—C¥6æ6†÷E7VÖÖ'”66†U²F66†T¶W•Ð¢–b‚FçVÆÂÖæRFÖVÖ÷'’ÖæB´–çCcEÒ„vWBÔFF&÷W'G’FÖVÖ÷'’w7F×rÓ’ÖWG7F×’°¢&WGW&â„vWBÔ'&’„vWBÔFF&÷W'G’FÖVÖ÷'’w7VÖÖ&–W2r‚’’¢Ð¢FÆö6ÅF‚ÒvWBÔÆö6Å6æ6†÷E7VÖÖ'”66†UF‚DÆæwVvRG6fUv÷&¶&öö´–@¢G'’°¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FÆö6ÅF‚’°¢FÆö6ÂÒ&VBÔ§6öäf–ÆRFÆö6ÅF‚FçVÆÀ¢–b‚FçVÆÂÖæRFÆö6ÂÖæ@¢¶–çEÒ„vWBÔFF&÷W'G’FÆö6Âw66†VÖfW'6–öâr’ÖW"Öæ@¢´–çCcEÒ„vWBÔFF&÷W'G’FÆö6Âw7F×rÓ’ÖWG7F×’°¢G7VÖÖ&–W2Ò„vWBÔ'&’„vWBÔFF&÷W'G’FÆö6Âw7VÖÖ&–W2r‚’’¢E67&—C¥6æ6†÷E7VÖÖ'”66†U²F66†T¶W•ÒÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²7F×ÒG7F×²7VÖÖ&–W2ÒG7VÖÖ&–W2Ð¢&WGW&âG7VÖÖ&–W0¢Ð¢Ð¢Ò6F6‚²Ð¢F÷WBÒ‚¢f÷&V6‚‚F–B–â‚„vWBÕ6æ6†÷D–G2DÆæwVvRG6fUv÷&¶&öö´–B’Â6VÆV7BÔö&¦V7BÔÆ7BC’’°¢FVçG'’ÒvWBÕ6æ6†÷E7VÖÖ'”VçG'’DÆæwVvRG6fUv÷&¶&öö´–B…·7G&–æuÒF–B¢–b‚FçVÆÂÖæRFVçG'’’²F÷WB³ÒFVçG'’Ð¢Ð¢G7VÖÖ&–W2Ò‚F÷WBÂ6÷'BÔö&¦V7B²·7G&–æuÒEòç6æ6†÷D–BÒÔFW66VæF–ær¢G&V6÷&BÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²66†VÖfW'6–öâÒ#²7F×ÒG7F×²7VÖÖ&–W2ÒG7VÖÖ&–W3²6fVDBÒæWrÔæ÷t—6òÐ¢E67&—C¥6æ6†÷E7VÖÖ'”66†U²F66†T¶W•ÒÒG&V6÷&@¢G'’²w&—FRÔ§6öäf–ÆRFÆö6ÅF‚G&V6÷&BÒ6F6‚²Ð¢&WGW&âG7VÖÖ&–W0§Ð ¦gVæ7F–öâvWBÕ7F÷&VD6ö×&—6öâ€¢·7G&–æuÒDÆæwVvRÀ¢·7G&–æuÒEv÷&¶&öö´–BÀ¢·7G&–æuÒDg&öÕ6æ6†÷D–BÀ¢·7G&–æuÒEFõ6æ6†÷D–BÀ¢·7G&–æuÒDg&öÕfW'6–öä–BÒrrÀ¢·7G&–æuÒEFõfW'6–öä–BÒrrÀ¢·7G&–æuÒE66÷RÒv†—7F÷'’p¢’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Ev÷&¶&öö´–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Dg&öÕ6æ6†÷D–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EFõ6æ6†÷D–B’’²&WGW&âFçVÆÂÐ¢GfW'6–öç2Ò–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EFõfW'6–öä–B’’²‚EFõfW'6–öä–B’ÒVÇ6R²„vWBÕ&VæFW%fW'6–öä–G2DÆæwVvREv÷&¶&öö´–BEFõ6æ6†÷D–B’Ð¢f÷&V6‚‚GfW"–âGfW'6–öç2’°¢FF—"Ò¦ö–âÕF‚„vWBÕ&VæFW%&V6÷&DF—"DÆæwVvREv÷&¶&öö´–BEFõ6æ6†÷D–B…·7G&–æuÒGfW"’’v6ö×&—6öç2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²6öçF–çVRÐ¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÂ6÷'BÔö&¦V7BÆ7Ew&—FUF–ÖUWF2ÔFW66VæF–ær’’°¢G'’°¢F6æF–FFRÒ&VBÔ§6öäf–ÆRFbägVÆÄæÖRFçVÆÀ¢–b‚FçVÆÂÖWF6æF–FFR’²6öçF–çVRÐ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRw66÷Rrrr’ÖæRE66÷R’²6öçF–çVRÐ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRv&6VÆ–æU6æ6†÷D–Brrr’ÖæRDg&öÕ6æ6†÷D–B’²6öçF–çVRÐ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRv7W'&VçE6æ6†÷D–Brrr’ÖæREFõ6æ6†÷D–B’²6öçF–çVRÐ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Dg&öÕfW'6–öä–B’Öæ@¢·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRv&6VÆ–æUfW'6–öä–Brrr’ÖæRDg&öÕfW'6–öä–B’²6öçF–çVRÐ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EFõfW'6–öä–B’Öæ@¢·7G&–æuÒ„vWBÔFF&÷W'G’F6æF–FFRv7W'&VçEfW'6–öä–Brrr’ÖæREFõfW'6–öä–B’²6öçF–çVRÐ¢&WGW&âF6æF–FFP¢Ò6F6‚²Ð¢Ð¢Ð¢&WGW&âFçVÆÀ§Ð ¦gVæ7F–öâ6W'fRÔ†—7F÷'”6öçFVçEFb‚D6öçFW‡BÂ·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒE6†VWDæÖRÂ·7G&–æuÒE6æ6†÷D–BÒrr’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Ev÷&¶&öö´–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6†VWDæÖR’’°¢F‡&÷rwv÷&¶&öö´–Bò6†VWDæÖR8Î[ø^Šh8~8ž8"p¢Ð¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G6fU6æ6†÷D–BÒrp¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6æ6†÷D–B’’²G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–BrÐ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EfW'6–öä–B’’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6fU6æ6†÷D–B’’²F‡&÷rwfW'6–öä–B8î8þ8ò6æ6†÷D–B8Î[ø^Šh8~8ž8"rÐ¢EfW'6–öä–BÒvWBÕ&VfW'&VD†—7F÷'•&VæFW%fW'6–öâDÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚EfW'6–öä–B’’²F‡&÷r~hÈ~Zé®8~8þjIÎyú^x˜Ž8åDn8þKùÞhÈ8^8(Î8n8N8î8¾8)>8"rÐ¢Ð¢G6fUfW'6–öä–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEfW'6–öä–BwfW'6–öä–Bp¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6fU6æ6†÷D–B’’°¢2&VæFW'>˜XÞKˆ¾8).jøîY¹îX‰~hÉž8¾8®8hÈ~Zé®8^8(Î8þx˜Ž88~8*>8:Î8*þ88Ž8:®8).y»Nhê^z+®Š¨Þ8ž8(¾8 ¢G&VæFW%&V6÷&DF—"ÒvWBÕ&VæFW%&V6÷&DF—"DÆæwVvRG6fUv÷&¶&öö´–BG6fU6æ6†÷D–BG6fUfW'6–öä–@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G&VæFW%&V6÷&DF—"ÕF…G—R6öçF–æW"’’°¢F‡&÷r~hÈ~Zé®8~8õDnK‰nKº>8þ88>8î[^jÛNx˜Ž8¾[î8~8n8N8î8¾8)>8"p¢Ð¢Ð¢FgVÆÂÒ&W6öÇfRÔ6öçFVçEFe6†VWEF„W†7BDÆæwVvRG6fUv÷&¶&öö´–BG6fUfW'6–öä–BE6†VWDæÖP¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FgVÆÂ’’²F‡&÷r~hÈ~Zé®8~8þK‰nKº>8åDn8ÎŠh¾8N8¾8(®8î8¾8)>8"rÐ¢2U$Î8÷6æ6†÷B÷fW'6–öâ÷6†VWN8~KˆÞZHž8.XZŽ898*N88Ž8).XXŽ8¾8:8:.8:®8ŽŠªÞ8(8î8).8(N8(8¢2Dbæ§>8ÎXù~KúkˆŽ8þ88~8;Î8+þ8¾8(žŠz>ié8).Zx¾8(8(ž8(Î8(¾8(Ž8n8+ž88Ž8:®8;Î89þ8;>8+8ž8(¾8 ¢w&—FRÔf–ÆU&W7öç6RD6öçFW‡B#FgVÆÂvÆ–6F–öâ÷FbrFfÇ6Rw&—fFRÂÖ‚ÖvSÓ3S3cÂ–Ö×WF&ÆRp§Ð  ¦gVæ7F–öâvWBÔWFõ7FFU7VÖÖ'’…·7G&–æuÒDÆæwVvRÂ·7v—F6…ÒDf7B’°¢G6WGF–æw2ÒvWBÔWFõ&VæFW%6WGF–æw0¢F†—7F÷'•6WGF–æw2ÒvWBÔ–çWD†—7F÷'•6WGF–æw0¢–b‚Df7B’°¢&WGW&â¶÷&FW&VEÔ°¢Væ&ÆVBÒ¶&ööÅÒG6WGF–æw2æVæ&ÆVC²–çWD†—7F÷'”&÷fVBÒGG'VP¢6÷W&6U&WFVçF–öä&÷fVBÒ¶&ööÅÒ…FW7BÕ6÷W&6U&WFVçF–öäVæ&ÆVB¢V–WEW&–öE6V6öæG2Ò¶–çEÒG6WGF–æw2çV–WEW&–öE6V6öæG0¢Ö…&WG'”6÷VçBÒ¶–çEÒG6WGF–æw2æÖ…&WG'”6÷VçC²&WG'”&6U6V6öæG2Ò¶–çEÒG6WGF–æw2ç&WG'”&6U6V6öæG0¢Ö–äg&VTÖVv'—FW2Ò¶–çEÒG6WGF–æw2æÖ–äg&VTÖVv'—FW3²æ÷F–g”öä6ö×ÆWF–öâÒ¶&ööÅÒG6WGF–æw2ææ÷F–g”öä6ö×ÆWF–öã²æ÷F–g”öäf–ÇW&RÒ¶&ööÅÒG6WGF–æw2ææ÷F–g”öäf–ÇW&P¢66†VGVÆW%'Vææ–ærÒFfÇ6S²†—7F÷'•6—¦TÖ"Ò ¢6ögD6ÖVv'—FW2Ò¶–çEÒF†—7F÷'•6WGF–æw2ç6ögD6ÖVv'—FW3²—FV×2Ò‚¢Ð¢Ð¢F—FV×2Ò‚¢G'’°¢FF—"ÒvWBÔWFõ7FFTF—"DÆæwVvP¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’°¢f÷&V6‚‚Fb–â„vWBÔ6†–ÆD—FVÒÔÆ—FW&ÅF‚FF—"Ôf–ÆRÔf–ÇFW"r¢æ§6öârÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVR’’°¢F¢Ò&VBÔ§6öäf–ÆRFbägVÆÄæÖRFçVÆÀ¢–b‚FçVÆÂÖWF¢’²6öçF–çVRÐ¢F—FVÕ7FFRÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢w7FFRrrr¢FÆ7E&W7VÇDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢vÆ7E&W7VÇDBrrr¢–b‚F—FVÕ7FFRÖWv–FÆRrÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FÆ7E&W7VÇDB’’²6öçF–çVRÐ¢F—FV×2³Ò¶÷&FW&VEÔ°¢v÷&¶&öö´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢wv÷&¶&öö´–BrFbä&6TæÖR¢7FFRÒF—FVÕ7FFP¢FVfW%&V6öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢vFVfW%&V6öârrr¢V–WDFVFÆ–æRÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢wV–WDFVFÆ–æRrrr¢f—'7DFWFV7FVDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢vf—'7DFWFV7FVDBrrr¢÷væW%4æÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢v÷væW%4æÖRrrr¢&WG'”6÷VçBÒ¶–çEÒ„vWBÔFF&÷W'G’F¢w&WG'”6÷VçBr¢æW‡E&WG'”BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢væW‡E&WG'”Brrr¢Æ7DW'&÷"Ò·7G&–æuÒ„vWBÔFF&÷W'G’F¢vÆ7DW'&÷"rrr¢Æ7E&W7VÇBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢vÆ7E&W7VÇBrrr¢Æ7E&W7VÇDBÒFÆ7E&W7VÇD@¢æ÷F–f–6F–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¢væ÷F–f–6F–öä–Brrr¢Ð¢Ð¢Ð¢Ò6F6‚²Ð¢&WGW&â¶÷&FW&VEÔ°¢Væ&ÆVBÒ¶&ööÅÒG6WGF–æw2æVæ&ÆVC²–çWD†—7F÷'”&÷fVBÒGG'VP¢6÷W&6U&WFVçF–öä&÷fVBÒ¶&ööÅÒ…FW7BÕ6÷W&6U&WFVçF–öäVæ&ÆVB¢V–WEW&–öE6V6öæG2Ò¶–çEÒG6WGF–æw2çV–WEW&–öE6V6öæG0¢Ö…&WG'”6÷VçBÒ¶–çEÒG6WGF–æw2æÖ…&WG'”6÷VçC²&WG'”&6U6V6öæG2Ò¶–çEÒG6WGF–æw2ç&WG'”&6U6V6öæG0¢Ö–äg&VTÖVv'—FW2Ò¶–çEÒG6WGF–æw2æÖ–äg&VTÖVv'—FW3²æ÷F–g”öä6ö×ÆWF–öâÒ¶&ööÅÒG6WGF–æw2ææ÷F–g”öä6ö×ÆWF–öã²æ÷F–g”öäf–ÇW&RÒ¶&ööÅÒG6WGF–æw2ææ÷F–g”öäf–ÇW&P¢66†VGVÆW%'Vææ–ærÒ…FW7BÔWFõ66†VGVÆW%&ö6W75'Vææ–ær¢†—7F÷'•6—¦TÖ"Ò„vWBÔ–çWD†—7F÷'•6—¦TÖ"DÆæwVvR¢6ögD6ÖVv'—FW2Ò¶–çEÒF†—7F÷'•6WGF–æw2ç6ögD6ÖVv'—FW3²—FV×2Ò‚F—FV×2¢Ð§Ð ¦gVæ7F–öâ&WVW7BÔWFõ'Väæ÷r…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–B’°¢2[»niÉþKŠÞ8î89n88>8*þ8).XÛ>i˜.ZéþŠÎ8ž8(¾8.8+ž8+8+Ž8:^8;Î8:ž8;ÎZÙ89~8:Þ8+¾8+ž8ÎjÊ8âF–6²8~h»î8n8 ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Ev÷&¶&öö´–B’’²F‡&÷rwv÷&¶&öö´–B8Î[ø^Šh8~8ž8"rÐ¢G7FFRÒ&VBÔWFõ7FFRDÆæwVvREv÷&¶&öö´–@¢–b‚FçVÆÂÖWG7FFR’²G7FFRÒæWrÔWFõ7FFREv÷&¶&öö´–BÐ¢6WBÔæ÷FU&÷W'G’G7FFRwV–WDFVFÆ–æRr…´FFUF–ÖUÓ£¥WF4æ÷räFE6V6öæG2‚Ó’åFõ7G&–ær‚vòr’¢6WBÔæ÷FU&÷W'G’G7FFRw7F&ÆT6÷VçBr“¢6WBÔæ÷FU&÷W'G’G7FFRvFVfW%&V6öârrp¢6WBÔæ÷FU&÷W'G’G7FFRw7FFRrwv—F–ærp¢6WBÔæ÷FU&÷W'G’G7FFRw&WG'”6÷VçBr ¢6WBÔæ÷FU&÷W'G’G7FFRvæW‡E&WG'”Brrp¢6WBÔæ÷FU&÷W'G’G7FFRvÆ7DW'&÷"rrp¢w&—FRÔWFõ7FFRDÆæwVvREv÷&¶&öö´–BG7FFP¢&WGW&â¶÷&FW&VEÔ²v÷&¶&öö´–BÒEv÷&¶&öö´–C²&WVW7FVBÒGG'VRÐ§Ð ¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ¢2cR[zîXˆnŠ›>{K8;¾ŠinŠi®jùN‹È0¢2ÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÓÐ ¢E67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öâÒ#0¢E67&—C¤F–fdFWF–ÄG’Ò# ¢E67&—C¤F–fdFWF–ÅF‡&W6†öÆBÒ#@¢E67&—C¤F–fdFWF–ÄÖ–æ–×VÕ&Vv–öå—†VÇ2Ò#@¢E67&—C¤F–fdFWF–ÅFF–ærÒP¢26öçFVçBDn8îx˜Ž88~8*>8:Î8*þ88Ž8:®8þyIþh‰ZèÎK¨n[èÎ8þKˆÞZHž8.8+~8;Î88ŽYÞ{J.[É^8).X[iÈž8~8¢2YÎ8Ž88Þ88>88Ž8:þ8;Î8*þ89^8*ž8:¾888;Î8).8+~8;Î88Ži[Xˆn88XhÞX‰~hÉž8~8®8N8 ¢E67&—C¤6öçFVçEFe6†VWD–æFW„66†RÒ·Ð¢E67&—C¤6öçFVçEFe6†VWD–æFW„66†TÆ–Ö—BÒc@ ¦gVæ7F–öâvWBÔF–fe6æ6†÷DFFR…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒDfÆÆ&6²Òrr’°¢G'’°¢FÒÒvWBÕ6æ6†÷DÖæ–fW7BDÆæwVvREv÷&¶&öö´–BE6æ6†÷D–@¢FFWFV7FVBÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÒvFWFV7FVDBrrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FFWFV7FVB’’²&WGW&âFFWFV7FVBÐ¢Ò6F6‚²Ð¢&WGW&âDfÆÆ&6°§Ð  ¦gVæ7F–öâvWBÕ&VfW'&VD†—7F÷'•&VæFW%fW'6–öâ…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–B’°¢f÷&V6‚‚GfW'6–öä–B–â‚„vWBÕ&VæFW%fW'6–öä–G2DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B’Â6÷'BÔö&¦V7BÔFW66VæF–ær’’°¢Ff–Æ&–Æ—G’ÒvWBÔ†—7F÷'•&VæFW%fW'6–öäf–Æ&–Æ—G’DÆæwVvREv÷&¶&öö´–BE6æ6†÷D–B…·7G&–æuÒGfW'6–öä–B¢–b…¶&ööÅÒFf–Æ&–Æ—G’ç&VG’’²&WGW&â·7G&–æuÒGfW'6–öä–BÐ¢Ð¢&WGW&ârp§Ð  ¦gVæ7F–öâæWrÔ†—7F÷&–6Å6æ6†÷D6ö×&—6öâ€¢·7G&–æuÒDÆæwVvRÀ¢·7G&–æuÒEv÷&¶&öö´–BÀ¢·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÀ¢·7G&–æuÒD&6VÆ–æUfW'6–öä–BÀ¢·7G&–æuÒD7W'&VçE6æ6†÷D–BÀ¢·7G&–æuÒD7W'&VçEfW'6–öä–@¢’°¢F&6RÒvWBÕf—7VÄ†6†W2DÆæwVvREv÷&¶&öö´–BD&6VÆ–æU6æ6†÷D–BD&6VÆ–æUfW'6–öä–@¢F7W'&VçBÒvWBÕf—7VÄ†6†W2DÆæwVvREv÷&¶&öö´–BD7W'&VçE6æ6†÷D–BD7W'&VçEfW'6–öä–@¢G&W7VÇBÒ¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ0¢7FGW2ÒwVæf–Æ&ÆRp¢66÷RÒv†—7F÷'’p¢&6VÆ–æU6æ6†÷D–BÒD&6VÆ–æU6æ6†÷D–@¢&6VÆ–æUfW'6–öä–BÒD&6VÆ–æUfW'6–öä–@¢7W'&VçE6æ6†÷D–BÒD7W'&VçE6æ6†÷D–@¢7W'&VçEfW'6–öä–BÒD7W'&VçEfW'6–öä–@¢6ö×&VDBÒæWrÔæ÷t—6ð¢ÖWF†öBÒw7F÷&VBÖ†6‚Ö†—7F÷'’p¢6öæf–FVæ6RÒã ¢6÷W&6UG—RÒvWBÔ6ö×&—6öå6÷W&6UG—RDÆæwVvREv÷&¶&öö´–@¢Væ—DÖ–æw2Ò‚¢Ö–æt6öæf–FVæ6RÒã ¢6†ævVE6†VWG2Ò‚¢Væ6†ævVE6†VWG2Ò‚¢Væ¶æ÷vå6†VWG2Ò‚¢FFVE6†VWG2Ò‚¢&VÖ÷fVE6†VWG2Ò‚¢ÖW76vRÒrp¢Ð¢–b‚FçVÆÂÖWF&6RÖ÷"FçVÆÂÖWF7W'&VçB’°¢G&W7VÇBæÖW76vRÒ~˜Žh©î8~8þx˜Ž8îyK¾X8þ88þ88>8+~8:^8ÎKùÞZÙŽ8^8(Î8n8N8®8N8þ8(jùN‹È>8~8Þ8î8¾8)>8"p¢&WGW&â·67W7FöÖö&¦V7EÒG&W7VÇ@¢Ð¢F&6TVçf—&öæÖVçBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F&6Rw&VæFW$Vçf—&öæÖVçDf–ævW'&–çBrrr¢F7W'&VçDVçf—&öæÖVçBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F7W'&VçBw&VæFW$Vçf—&öæÖVçDf–ævW'&–çBrrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6TVçf—&öæÖVçB’Öæ@¢Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçDVçf—&öæÖVçB’Öæ@¢F&6TVçf—&öæÖVçBÖæRF7W'&VçDVçf—&öæÖVçB’°¢G&W7VÇBæÖWF†öBÒw7F÷&VBÖ†6‚Ö†—7F÷'’ÖVçf—&öæÖVçBÖÖ—6ÖF6‚p¢G&W7VÇBæ6öæf–FVæ6RÒãcP¢G&W7VÇBæÖW76vRÒs.x˜Ž8åDnKÙÎh‰y+Z(>8Îy[8®8(®8î8ž8.ŠŽzK®{YiéÎ8).yºîŠin8~z+®Š¨Þ8~8n8þ88^8N8"p¢Ð¢FÖ–æw2ÒvWBÔ6ö×&—6öåVæ—DÖ–æw2„vWBÔFF&÷W'G’F&6Rw6†VWG2r‚’’„vWBÔFF&÷W'G’F7W'&VçBw6†VWG2r‚’’…·7G&–æuÒG&W7VÇBç6÷W&6UG—R¢G&W7VÇBç7FGW2Òv6ö×ÆWFRp¢6WBÔ6ö×&—6öåVæ—DÖ–æu&W7VÇBG&W7VÇBFÖ–æw0¢&WGW&â·67W7FöÖö&¦V7EÒG&W7VÇ@§Ð  ¦gVæ7F–öâ6fRÔ†—7F÷&–6Å6æ6†÷D6ö×&—6öâ…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂD6ö×&—6öâ’°¢–b‚FçVÆÂÖWD6ö×&—6öâÖ÷"·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâw7FGW2rrr’ÖæRv6ö×ÆWFRr’²F‡&÷r~KùÞZÙŽ8~8Þ8(¾[^jÛNjùN‹È>{YiéÎ8Î8.8(®8î8¾8)>8"rÐ¢f÷&V6‚‚Ff–VÆB–â‚v&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâFf–VÆBrr’’’²F‡&÷r.[^jÛNjùN‹È>{YiéÎ8îŠÙŽXŠ^ZÙ8ÎKˆÞ‹k>8~8n8N8î8“¢Ff–VÆB"Ð¢Ð¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâw66÷Rrrr’ÖæRv†—7F÷'’r’²F‡&÷r~[^jÛNjùN‹È>Kº^ZIn8þ8>8îKùÞZÙŽ{XÎ‹zþ8).KÛþyJŽ8~8Þ8î8¾8)>8"rÐ¢F7W'&VçE6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–@¢F7W'&VçEfW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçEfW'6–öä–@¢FF—"Ò¦ö–âÕF‚„vWBÕ&VæFW%&V6÷&DF—"DÆæwVvREv÷&¶&öö´–BF7W'&VçE6æ6†÷D–BF7W'&VçEfW'6–öä–B’v6ö×&—6öç2p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚FF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢F¶W’Ò„vWBÕ6†#SeFW‡B‚&†—7F÷'—Ç³×Ç³×Ç³'×Ç³7Ò"ÖbD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–BÂD6ö×&—6öâæ&6VÆ–æUfW'6–öä–BÂF7W'&VçE6æ6†÷D–BÂF7W'&VçEfW'6–öä–B’’å7V'7G&–ærƒrÂb¢GF‚Ò¦ö–âÕF‚FF—"‚&†6××³Òæ§6öâ"ÖbF¶W’¢w&—FRÔ§6öäf–ÆRGF‚D6ö×&—6öà¢G6fVBÒ&VBÔ§6öäf–ÆRGF‚FçVÆÀ¢–b‚FçVÆÂÖWG6fVB’²F‡&÷r~[^jÛNjùN‹È>{YiéÎ8).KùÞZÙŽ[èÎ8¾XhÞŠªÞ‹ëÎ8~8Þ8î8¾8)>8~8~8þ8"rÐ¢f÷&V6‚‚Ff–VÆB–â‚w66÷RrÂv&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G6fVBFf–VÆBrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâFf–VÆBrr’’°¢F‡&÷r.[^jÛNjùN‹È>{YiéÎ8îKùÞZÙŽjIÎŠ‹Î8¾ZKiY~8~8î8~8ó¢Ff–VÆB ¢Ð¢Ð¢w&—FRÔ†—7F÷'”WfVçBDÆæwVvRv6ö×&Ræ†—7F÷'’æ7&VFVBr…¶÷&FW&VEÔ°¢v÷&¶&öö´–BÒEv÷&¶&öö´–@¢&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–@¢&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æUfW'6–öä–@¢7W'&VçE6æ6†÷D–BÒF7W'&VçE6æ6†÷D–@¢7W'&VçEfW'6–öä–BÒF7W'&VçEfW'6–öä–@¢6†ævVBÒ„vWBÔ'&’D6ö×&—6öâæ6†ævVE6†VWG2¢Væ¶æ÷vâÒ„vWBÔ'&’D6ö×&—6öâçVæ¶æ÷vå6†VWG2¢Ò¢&WGW&âG6fV@§Ð  ¦gVæ7F–öâvWBÔF–fdFWF–Ä6öçFW‡B€¢·7G&–æuÒDÆæwVvRÀ¢·7G&–æuÒEv÷&¶&öö´–BÀ¢·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÒrrÀ¢·7G&–æuÒD7W'&VçE6æ6†÷D–BÒrp¢’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢FÖF6†W2Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWG6fUv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚FÖF6†W2ä6÷VçBÖW’²F‡&÷r~y›¾˜Ë.kˆŽ8þXéþz‹þ8ÎŠh¾8N8¾8(®8î8¾8)>8"rÐ¢Gv÷&¶&öö²ÒFÖF6†W5³Ð¢FF—7Æ”æÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²vF—7Æ”æÖRrrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FF—7Æ”æÖR’’²FF—7Æ”æÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²vf–ÆTæÖRrG6fUv÷&¶&öö´–B’Ð¢F&6U&W7VÇBÒ¶÷&FW&VEÔ°¢f–Æ&ÆRÒFfÇ6S²7FGW2ÒwVæf–Æ&ÆRs²ÖW76vRÒrp¢v÷&¶&öö´–BÒG6fUv÷&¶&öö´–C²v÷&¶&öö´æÖRÒFF—7Æ”æÖS²v÷&¶&öö²ÒGv÷&¶&öö°¢6ö×&—6öâÒFçVÆÃ²6ö×&—6öåW'6—7FVBÒFfÇ6P¢7W'&VçE6æ6†÷D–BÒrs²7W'&VçEfW'6–öä–BÒrs²&6VÆ–æU6æ6†÷D–BÒrs²&6VÆ–æUfW'6–öä–BÒrp¢7W'&VçEf—7VÄ†6†W2ÒFçVÆÃ²&6VÆ–æUf—7VÄ†6†W2ÒFçVÆÀ¢7W'&VçDBÒrs²&6VÆ–æTBÒrs²ÖWF†öBÒrs²66÷RÒvWFöÖF–2p¢6÷W&6UG—RÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²w6÷W&6UG—RrvW†6VÂr¢Ð¢–b‚Öæ÷B…FW7BÔ–çWD†—7F÷'”Væ&ÆVB’’°¢F&6U&W7VÇBæÖW76vRÒ~XZ^X©¾[^jÛN8ÎxJX«ž8®8þ8(8[zîXˆnŠ›>{K8þXŠžyJŽ8~8Þ8î8¾8)>8"p¢&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢F†4&6VÆ–æRÒÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D&6VÆ–æU6æ6†÷D–B¢F†47W'&VçBÒÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D7W'&VçE6æ6†÷D–B¢–b‚F†4&6VÆ–æR×†÷"F†47W'&VçB’²F‡&÷r~[^jÛNjùN‹È>8~8þjùN‹È>XX>8ŽjùN‹È>XXŽ8îKŠikž8).hÈ~Zé®8~8n8þ88^8N8"rÐ¢–b‚F†4&6VÆ–æRÖæBF†47W'&VçB’°¢F&6VÆ–æU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBD&6VÆ–æU6æ6†÷D–Bv&6VÆ–æU6æ6†÷D–Bp¢F7W'&VçE6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBD7W'&VçE6æ6†÷D–Bv7W'&VçE6æ6†÷D–Bp¢F&6U&W7VÇBç66÷RÒv†—7F÷'’p¢–b‚F&6VÆ–æU6æ6†÷D–BÖWF7W'&VçE6æ6†÷D–B’²F&6U&W7VÇBæÖW76vRÒ~y[8®8(³.x˜Ž8).˜Žh©î8~8n8þ88^8N8"s²&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇBÐ¢–b‚FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvRG6fUv÷&¶&öö´–BF&6VÆ–æU6æ6†÷D–B’Ö÷ ¢FçVÆÂÖW„vWBÕ6æ6†÷DÖæ–fW7BDÆæwVvRG6fUv÷&¶&öö´–BF7W'&VçE6æ6†÷D–B’’°¢F&6U&W7VÇBæÖW76vRÒ~˜Žh©î8~8þ[^jÛNx˜Ž8ÎŠh¾8N8¾8(®8î8¾8)>8.KùÞZÙŽiÉþ™™8î8þ8þ[^jÛNi[Nyn8).z+®Š¨Þ8~8n8þ88^8N8"p¢&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢F&6VÆ–æUfW'6–öä–BÒvWBÕ&VfW'&VD†—7F÷'•&VæFW%fW'6–öâDÆæwVvRG6fUv÷&¶&öö´–BF&6VÆ–æU6æ6†÷D–@¢F7W'&VçEfW'6–öä–BÒvWBÕ&VfW'&VD†—7F÷'•&VæFW%fW'6–öâDÆæwVvRG6fUv÷&¶&öö´–BF7W'&VçE6æ6†÷D–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6VÆ–æUfW'6–öä–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçEfW'6–öä–B’’°¢F&6U&W7VÇBæÖW76vRÒ~˜Žh©î8~8ó.x˜Ž8þ8yK¾X8þ88þ88>8+~8:^8ŽYÎKˆK‰nKº>8æ6öçFVçBDn8Î8Þ8(Þ8>8n8N8®8N8þ8(ŠinŠi®jùN‹È>8~8Þ8î8¾8)>8"p¢&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢F6ö×&—6öâÒvWBÕ7F÷&VD6ö×&—6öâDÆæwVvRG6fUv÷&¶&öö´–BF&6VÆ–æU6æ6†÷D–BF7W'&VçE6æ6†÷D–BF&6VÆ–æUfW'6–öä–BF7W'&VçEfW'6–öä–Bv†—7F÷'’p¢–b‚FçVÆÂÖWF6ö×&—6öâ’°¢F6ö×&—6öâÒæWrÔ†—7F÷&–6Å6æ6†÷D6ö×&—6öâDÆæwVvRG6fUv÷&¶&öö´–BF&6VÆ–æU6æ6†÷D–BF&6VÆ–æUfW'6–öä–BF7W'&VçE6æ6†÷D–BF7W'&VçEfW'6–öä–@¢ÒVÇ6R²F&6U&W7VÇBæ6ö×&—6öåW'6—7FVBÒGG'VRÐ¢–b‚FçVÆÂÖWF6ö×&—6öâÖ÷"·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâw7FGW2rrr’ÖæRv6ö×ÆWFRr’°¢F&6U&W7VÇBæÖW76vRÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâvÖW76vRr~˜Žh©î8~8ó.x˜Ž8).jùN‹È>8~8Þ8î8¾8)>8"r¢&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢F&6U&W7VÇBæf–Æ&ÆRÒGG'VS²F&6U&W7VÇBç7FGW2Òvf–Æ&ÆRs²F&6U&W7VÇBæ6ö×&—6öâÒF6ö×&—6öà¢F&6U&W7VÇBæ7W'&VçE6æ6†÷D–BÒF7W'&VçE6æ6†÷D–C²F&6U&W7VÇBæ7W'&VçEfW'6–öä–BÒF7W'&VçEfW'6–öä–@¢F&6U&W7VÇBæ&6VÆ–æU6æ6†÷D–BÒF&6VÆ–æU6æ6†÷D–C²F&6U&W7VÇBæ&6VÆ–æUfW'6–öä–BÒF&6VÆ–æUfW'6–öä–@¢F&6U&W7VÇBæ7W'&VçDBÒvWBÔF–fe6æ6†÷DFFRDÆæwVvRG6fUv÷&¶&öö´–BF7W'&VçE6æ6†÷D–Brp¢F&6U&W7VÇBæ&6VÆ–æTBÒvWBÔF–fe6æ6†÷DFFRDÆæwVvRG6fUv÷&¶&öö´–BF&6VÆ–æU6æ6†÷D–Brp¢F&6U&W7VÇBæÖWF†öBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâvÖWF†öBrw7F÷&VBÖ†6‚Ö†—7F÷'’r¢&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢G7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²w7FGW2rrr¢F7W'&VçD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²v7W'&VçDW†6VÄ†6‚rrr’¢G&VæFW&VD†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²vÆ7E&VæFW&VDW†6VÄ†6‚rrr’¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&VæFW&VD†6‚’Ö÷"G7FGW2Ö–â‚væWrrÂvW†6VÂ×WFFVBrÂw6÷W&6R×WFFVBrÂw&VæFW"ÖW'&÷"rÂw&VæFW&–ærr’Ö÷ ¢‚‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçD†6‚’’ÖæBF7W'&VçD†6‚ÖæRG&VæFW&VD†6‚’’°¢F&6U&W7VÇBæÖW76vRÒ~iÈik8åDn8).KÙÎh‰8~8n8¾8(ž[zîXˆn8).z+®Š¨Þ8~8n8þ88^8N8"s²&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢F7W'&VçE6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²vÆ7E&VæFW&VE6æ6†÷D–Brrr¢F7W'&VçEfW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²vÆ7E&VæFW&VEfW'6–öä–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçE6æ6†÷D–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçEfW'6–öä–B’’°¢F&6U&W7VÇBæÖW76vRÒ~xûîYÊŽx˜Ž8).KˆhHþ8¾ŠÙŽXŠ^8~8Þ8®8N8þ8(8[zîXˆnŠ›>{K8).™h¾88î8¾8)>8"s²&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢F7W'&VçE6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBF7W'&VçE6æ6†÷D–Bv7W'&VçE6æ6†÷D–Bp¢F7W'&VçEfW'6–öä–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBF7W'&VçEfW'6–öä–Bv7W'&VçEfW'6–öä–Bp¢F6ö×&—6öâÒvWBÔÆFW7D6ö×&—6öâDÆæwVvRG6fUv÷&¶&öö´–BGv÷&¶&öö°¢–b‚FçVÆÂÖWF6ö×&—6öâ’²F&6U&W7VÇBæÖW76vRÒ~KùÞZÙŽkˆŽ8þ8îjùN‹È>{YiéÎ8Î8.8(®8î8¾8)>8%Dn8).XhÞKÙÎh‰8~8n8þ88^8N8"s²&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇBÐ¢F&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv&6VÆ–æU6æ6†÷D–Brrr¢F&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv&6VÆ–æUfW'6–öä–Brrr¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâw7FGW2rrr’ÖæRv6ö×ÆWFRrÖ÷ ¢·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6VÆ–æU6æ6†÷D–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6VÆ–æUfW'6–öä–B’’°¢F&6U&W7VÇBæÖW76vRÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâvÖW76vRr~X˜ÞY¹î8îjùN‹È>Yû®k©n8Î8.8(®8î8¾8)>8"r“²&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢F&6VÆ–æU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBF&6VÆ–æU6æ6†÷D–Bv&6VÆ–æU6æ6†÷D–Bp¢F&6VÆ–æUfW'6–öä–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBF&6VÆ–æUfW'6–öä–Bv&6VÆ–æUfW'6–öä–Bp¢2ˆz®X¹^jùN‹È>8îx˜Ž8þKÙÎh‰i˜.8¾XZ…Dn8).jIÎŠ‹Î8;¾Y»®Zé®kˆŽ8þ8.Š›>{K8).™h¾8þ8þ8>8¾XZŽ8+~8;Î88Ž8) ¢2XhÞX‰~hÉž8¾8®8‹»Þ˜xþ8®x˜Ž88~8*>8:Î8*þ88Ž8:®8Ž88þ88>8+~8:^888).z+®Š¨Þ8ž8(¾8 ¢2X¾XŠUDn8þŠŽzK®Šhk.i˜.8²6W'fRÔ†—7F÷'”6öçFVçEFb8ÎXë>Zøn8¾jIÎŠ‹Î8ž8(¾8 ¢F7W'&VçD†6†W2ÒvWBÕf—7VÄ†6†W2DÆæwVvRG6fUv÷&¶&öö´–BF7W'&VçE6æ6†÷D–BF7W'&VçEfW'6–öä–@¢F&6VÆ–æT†6†W2ÒvWBÕf—7VÄ†6†W2DÆæwVvRG6fUv÷&¶&öö´–BF&6VÆ–æU6æ6†÷D–BF&6VÆ–æUfW'6–öä–@¢–b‚FçVÆÂÖWF7W'&VçD†6†W2Ö÷"FçVÆÂÖWF&6VÆ–æT†6†W2’°¢F&6U&W7VÇBæÖW76vRÒ~ˆz®X¹^jùN‹È>8¾KÛþ8>8þyK¾X8þ88þ88>8+~8:^8ÎKùÞhÈ8^8(Î8n8N8î8¾8)>8%Dn8).XhÞKÙÎh‰8~8n8þ88^8N8"p¢&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@¢Ð¢F&6U&W7VÇBæ7W'&VçEf—7VÄ†6†W2ÒF7W'&VçD†6†W0¢F&6U&W7VÇBæ&6VÆ–æUf—7VÄ†6†W2ÒF&6VÆ–æT†6†W0¢F&6U&W7VÇBæf–Æ&ÆRÒGG'VS²F&6U&W7VÇBç7FGW2Òvf–Æ&ÆRs²F&6U&W7VÇBæ6ö×&—6öâÒF6ö×&—6öã²F&6U&W7VÇBæ6ö×&—6öåW'6—7FVBÒGG'VP¢F&6U&W7VÇBæ7W'&VçE6æ6†÷D–BÒF7W'&VçE6æ6†÷D–C²F&6U&W7VÇBæ7W'&VçEfW'6–öä–BÒF7W'&VçEfW'6–öä–@¢F&6U&W7VÇBæ&6VÆ–æU6æ6†÷D–BÒF&6VÆ–æU6æ6†÷D–C²F&6U&W7VÇBæ&6VÆ–æUfW'6–öä–BÒF&6VÆ–æUfW'6–öä–@¢F&6U&W7VÇBæ7W'&VçDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²vÆ7E&VæFW&VDBr…·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv6ö×&VDBrrr’’¢F&6U&W7VÇBæ&6VÆ–æTBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv&6VÆ–æTBrrr¢F&6U&W7VÇBæÖWF†öBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâvÖWF†öBrrr¢&WGW&â·67W7FöÖö&¦V7EÒF&6U&W7VÇ@§Ð  ¦gVæ7F–öâvWBÔF–fe—$¶W’‚D6öçFW‡B’°¢&WGW&â„vWBÕ6†#SeFW‡B‚'³×Ç³×Ç³'×Ç³7×Ç³G×Ç³W×Ç³gÒ"Ö`¢·7G&–æuÒD6öçFW‡Bç66÷RÂ·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–BÀ¢·7G&–æuÒD6öçFW‡Bæ&6VÆ–æU6æ6†÷D–BÂ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æUfW'6–öä–BÀ¢·7G&–æuÒD6öçFW‡Bæ7W'&VçE6æ6†÷D–BÂ·7G&–æuÒD6öçFW‡Bæ7W'&VçEfW'6–öä–BÀ¢E67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öâ’’å7V'7G&–ærƒrÂ#B§Ð ¦gVæ7F–öâvWBÔF–fe&Wf–Wu7FFUF‚…·7G&–æuÒDÆæwVvRÂD6öçFW‡B’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’wv÷&¶&öö´–Bp¢G&Wf–Wu&ö÷BÒ¦ö–âÕF‚E67&—C¤Æö6Ä6öæf–u&ö÷BvF–fb×&Wf–Ww2p¢&WGW&â„¦ö–âÕF‚G&Wf–Wu&ö÷B‚'&Wf–Wr×³Ò×³Ò×³'Òæ§6öâ"ÖbDÆæwVvRÂG6fUv÷&¶&öö´–BÂ„vWBÔF–fe—$¶W’D6öçFW‡B’’§Ð ¦gVæ7F–öâvWBÔF–fe&Wf–Wu7FFTf÷$6öçFW‡B…·7G&–æuÒDÆæwVvRÂD6öçFW‡B’°¢FV×G’Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ¢v÷&¶&öö´–BÒ·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–@¢66÷RÒ·7G&–æuÒD6öçFW‡Bç66÷P¢&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æU6æ6†÷D–@¢&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æUfW'6–öä–@¢7W'&VçE6æ6†÷D–BÒ·7G&–æuÒD6öçFW‡Bæ7W'&VçE6æ6†÷D–@¢7W'&VçEfW'6–öä–BÒ·7G&–æuÒD6öçFW‡Bæ7W'&VçEfW'6–öä–@¢Æv÷&—F†ÕfW'6–öâÒE67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öà¢6öæf—&ÖVE6†VWD¶W—2Ò‚¢&Wf–WvVDBÒrp¢&Wf–WvVD'’Òrp¢Ð¢–b‚Öæ÷B¶&ööÅÒD6öçFW‡Bæf–Æ&ÆR’²&WGW&âFV×G’Ð¢GF‚ÒvWBÔF–fe&Wf–Wu7FFUF‚DÆæwVvRD6öçFW‡@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’’²&WGW&âFV×G’Ð¢G'’²G6fVBÒ&VBÔ§6öäf–ÆRGF‚FçVÆÂÒ6F6‚²&WGW&âFV×G’Ð¢–b‚FçVÆÂÖWG6fVBÖ÷"„vWBÔ–çDFF&÷W'G’G6fVBvÆv÷&—F†ÕfW'6–öâr’ÖæRE67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öâ’²&WGW&âFV×G’Ð¢f÷&V6‚‚Ff–VÆB–â‚wv÷&¶&öö´–BrÂw66÷RrÂv&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’G6fVBFf–VÆBrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’FV×G’Ff–VÆBrr’’²&WGW&âFV×G’Ð¢Ð¢FV×G’æ6öæf—&ÖVE6†VWD¶W—2Ò„vWBÔ'&’„vWBÔFF&÷W'G’G6fVBv6öæf—&ÖVE6†VWD¶W—2r‚’’Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒÂv†W&RÔö&¦V7B²EòÒÂ6VÆV7BÔö&¦V7BÕVæ—VR¢FV×G’ç&Wf–WvVDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6fVBw&Wf–WvVDBrrr¢FV×G’ç&Wf–WvVD'’Ò·7G&–æuÒ„vWBÔFF&÷W'G’G6fVBw&Wf–WvVD'’rrr¢&WGW&âFV×G§Ð ¦gVæ7F–öâ6WBÔF–fe&Wf–Wu7FFR…·7G&–æuÒDÆæwVvRÂD&öG’’°¢Gv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒ„vWBÔFF&÷W'G’D&öG’wv÷&¶&öö´–Brrr’’wv÷&¶&öö´–Bp¢Fg&öÕ6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’D&öG’vg&öÕ6æ6†÷D–Brrr¢GFõ6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’D&öG’wFõ6æ6†÷D–Brrr¢F6öçFW‡BÒvWBÔF–fdFWF–Ä6öçFW‡BDÆæwVvRGv÷&¶&öö´–BFg&öÕ6æ6†÷D–BGFõ6æ6†÷D–@¢–b‚Öæ÷B¶&ööÅÒF6öçFW‡Bæf–Æ&ÆR’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr…·7G&–æuÒ„vWBÔFF&÷W'G’F6öçFW‡BvÖW76vRr~jùN‹È>Zûî‹8).z+®Š¨Þ8~8Þ8î8¾8)>8"r’’Ð¢f÷&V6‚‚Ff–VÆB–â‚v&6VÆ–æUfW'6–öä–BrÂv7W'&VçEfW'6–öä–Br’’°¢G&WVW7FVBÒ·7G&–æuÒ„vWBÔFF&÷W'G’D&öG’Ff–VÆBrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&WVW7FVB’ÖæBG&WVW7FVBÖæR·7G&–æuÒ„vWBÔFF&÷W'G’F6öçFW‡BFf–VÆBrr’’°¢F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~jùN‹È>Zûî‹8Îi»Nik8^8(Î8î8~8þ8.[zîXˆnyK¾™Ú.8).™h¾8Þy»N8~8n8þ88^8N8"r¢Ð¢Ð¢G6†VWD¶W’Ò76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒ„vWBÔFF&÷W'G’D&öG’w6†VWD¶W’rrr’’w6†VWD¶W’p¢FFWF–ÂÒæWrÔF–fdFWF–Å6¶VÆWFöâDÆæwVvRF6öçFW‡@¢–b„„vWBÔ'&’FFWF–Âç6†VWG2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç6†VWD¶W’ÖWG6†VWD¶W’Ò’ä6÷VçBÖW’°¢F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚~z+®Š¨ÞZûî‹8î89®8;Î8+Žš^yºî8ÎŠh¾8N8¾8(®8î8¾8)>8"r¢Ð¢G&Wf–WrÒvWBÔF–fe&Wf–Wu7FFTf÷$6öçFW‡BDÆæwVvRF6öçFW‡@¢F¶W—2Ò´6öÆÆV7F–öç2ävVæW&–2ä†6…6WE·7G&–æuÕÓ£¦æWr…µ7G&–æt6ö×&W%Ó£¤÷&F–æÂ¢f÷&V6‚‚FW†—7F–ær–â„vWBÔ'&’G&Wf–Wræ6öæf—&ÖVE6†VWD¶W—2’’²–b‚FW†—7F–ær’²·fö–EÒF¶W—2äFB…·7G&–æuÒFW†—7F–ær’ÒÐ¢F6öæf—&ÖVBÒ¶&ööÅÒ„vWBÔFF&÷W'G’D&öG’v6öæf—&ÖVBrGG'VR¢–b‚F6öæf—&ÖVB’²·fö–EÒF¶W—2äFB‚G6†VWD¶W’’ÒVÇ6R²·fö–EÒF¶W—2å&VÖ÷fR‚G6†VWD¶W’’Ð¢G&Wf–Wræ6öæf—&ÖVE6†VWD¶W—2Ò‚F¶W—2Â6÷'BÔö&¦V7B¢G&Wf–Wrç&Wf–WvVDBÒæWrÔæ÷t—6ð¢G&Wf–Wrç&Wf–WvVD'’Ò·7G&–æuÒFVçc¥U4U$äÔP¢GF‚ÒvWBÔF–fe&Wf–Wu7FFUF‚DÆæwVvRF6öçFW‡@¢G&VçBÒ7Æ—BÕF‚Õ&VçBGF€¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G&VçB’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚G&VçBÔf÷&6RÂ÷WBÔçVÆÂÐ¢w&—FRÔ§6öäf–ÆRGF‚G&Wf–Wp¢F66†UF‚ÒvWBÔÆö6ÄF–fdFWF–Ä66†UF‚DÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒF6öçFW‡Bæ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒF6öçFW‡Bæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒF6öçFW‡Bæ&6VÆ–æUfW'6–öä–B’…·7G&–æuÒF6öçFW‡Bæ7W'&VçEfW'6–öä–B’…·7G&–æuÒF6öçFW‡Bç66÷R¢F66†T¶W’Ò…´”òåF…Ó£¤vWDgVÆÅF‚‚F66†UF‚’’åFôÆ÷vW$–çf&–çB‚¢·fö–EÒE67&—C¤F–fdFWF–Å&W7öç6T66†Rå&VÖ÷fR‚F66†T¶W’¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚F66†UF‚’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚F66†UF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢&WGW&âG&Wf–Wp§Ð ¦gVæ7F–öâvWBÔF–fdÆVæ6„Æö6µF‚…·7G&–æuÒDÆæwVvRÂD6öçFW‡B’°¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’‚&Æö6·5ÆF–fbÖÆVæ6…÷³ÒæÆö6²"Öb„vWBÔF–fe—$¶W’D6öçFW‡B’’§Ð¦gVæ7F–öâvWBÔF–fdvVæW&F–öäÆö6µF‚…·7G&–æuÒDÆæwVvRÂD6öçFW‡B’°¢&WGW&â„¦ö–âÕF‚„vWBÕv÷&·76UF‚DÆæwVvR’‚&Æö6·5ÆF–fbÖvVæW&FU÷³ÒæÆö6²"Öb„vWBÔF–fe—$¶W’D6öçFW‡B’’§Ð ¦gVæ7F–öâvWBÔF–fdFWF–Ä66†TF—"‚D6öçFW‡B’°¢G&V6÷&BÒvWBÕ&VæFW%&V6÷&DF—"„vWBÔVffV7F—fTÆæwVvR’…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒD6öçFW‡Bæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒD6öçFW‡Bæ7W'&VçEfW'6–öä–B¢F&6VÆ–æRÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒD6öçFW‡Bæ&6VÆ–æU6æ6†÷D–B’v&6VÆ–æU6æ6†÷D–Bp¢F66†T¶W’Ò„vWBÕ6†#SeFW‡B‚'³×Ç³×Ç³'×Ç³7Ò"Öb·7G&–æuÒD6öçFW‡Bç66÷RÂF&6VÆ–æRÂ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æUfW'6–öä–BÂE67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öâ’’å7V'7G&–ærƒrÂb¢&WGW&â„¦ö–âÕF‚G&V6÷&B„¦ö–âÕF‚v6ö×&—6öç2r‚&G³Ò"ÖbF66†T¶W’’’§Ð ¦gVæ7F–öâvWBÔF–fdFWF–Ä66†TF—$f÷$ÆæwVvR…·7G&–æuÒDÆæwVvRÂD6öçFW‡B’°¢G&V6÷&BÒvWBÕ&VæFW%&V6÷&DF—"DÆæwVvR…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒD6öçFW‡Bæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒD6öçFW‡Bæ7W'&VçEfW'6–öä–B¢F&6VÆ–æRÒ76W'BÕ6fU7F÷&vU6VvÖVçB…·7G&–æuÒD6öçFW‡Bæ&6VÆ–æU6æ6†÷D–B’v&6VÆ–æU6æ6†÷D–Bp¢F66†T¶W’Ò„vWBÕ6†#SeFW‡B‚'³×Ç³×Ç³'×Ç³7Ò"Öb·7G&–æuÒD6öçFW‡Bç66÷RÂF&6VÆ–æRÂ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æUfW'6–öä–BÂE67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öâ’’å7V'7G&–ærƒrÂb¢&WGW&â„¦ö–âÕF‚G&V6÷&B„¦ö–âÕF‚v6ö×&—6öç2r‚&G³Ò"ÖbF66†T¶W’’’§Ð ¦gVæ7F–öâvWBÔF–fe6†VWD¶W’…·7G&–æuÒE6†VWDæÖR’°¢F†6‚ÒvWBÕ6†#SeFW‡BE6†VWDæÖP¢&WGW&â‚w2Òr²F†6‚å7V'7G&–ærƒrÂ’§Ð ¦gVæ7F–öâFW7BÔF–fdFWF–ÄÖF6†W46öçFW‡B‚DFWF–ÂÂD6öçFW‡B’°¢–b‚FçVÆÂÖWDFWF–ÂÖ÷"FçVÆÂÖWD6öçFW‡B’²&WGW&âFfÇ6RÐ¢–b‚„vWBÔ–çDFF&÷W'G’DFWF–ÂvÆv÷&—F†ÕfW'6–öâr’ÖæRE67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öâ’²&WGW&âFfÇ6RÐ¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’DFWF–Âwv÷&¶&öö´–Brrr’ÖæR·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’²&WGW&âFfÇ6RÐ¢F6ö×&—6öâÒvWBÔFF&÷W'G’DFWF–Âv6ö×&—6öârFçVÆÀ¢–b‚FçVÆÂÖWF6ö×&—6öâ’²&WGW&âFfÇ6RÐ¢f÷&V6‚‚Ff–VÆB–â‚w66÷RrÂv&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâFf–VÆBrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’D6öçFW‡BFf–VÆBrr’’²&WGW&âFfÇ6RÐ¢Ð¢&WGW&âGG'VP§Ð ¦gVæ7F–öâvWBÔF–fd†6…6†VWDÖ‚D†6†W2’°¢FÖÒ·Ð¢–b‚FçVÆÂÖWD†6†W2’²&WGW&âFÖÐ¢f÷&V6‚‚G6†VWB–â„vWBÔ'&’„vWBÔFF&÷W'G’D†6†W2w6†VWG2r‚’’’’°¢FÖµ·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBw6†VWDæÖRrrr•ÒÒG6†VW@¢Ð¢&WGW&âFÖ §Ð ¦gVæ7F–öâæWrÔF–fdFWF–Å6¶VÆWFöâ…·7G&–æuÒDÆæwVvRÂD6öçFW‡B’°¢F6ö×&—6öâÒD6öçFW‡Bæ6ö×&—6öà¢FFFVE6WBÒ·Ð¢G&VÖ÷fVE6WBÒ·Ð¢GVæ¶æ÷vå6WBÒ·Ð¢f÷&V6‚‚FæÖR–â„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâvFFVE6†VWG2r‚’’’’²FFFVE6WEµ·7G&–æuÒFæÖUÒÒGG'VRÐ¢f÷&V6‚‚FæÖR–â„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâw&VÖ÷fVE6†VWG2r‚’’’’²G&VÖ÷fVE6WEµ·7G&–æuÒFæÖUÒÒGG'VRÐ¢f÷&V6‚‚FæÖR–â„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâwVæ¶æ÷vå6†VWG2r‚’’’’²GVæ¶æ÷vå6WEµ·7G&–æuÒFæÖUÒÒGG'VRÐ ¢F—FV×2Ò‚“²G6VVâÒ‚¢FÖ–æu&V6÷&G2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâwVæ—DÖ–æw2r‚’’¢–b‚FÖ–æu&V6÷&G2ä6÷VçBÖW’°¢f÷&V6‚‚Fw&÷W–â€¢¶÷&FW&VEÔ²¶–æBÒvÖöF–f–VBs²æÖW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâv6†ævVE6†VWG2r‚’’’ÒÀ¢¶÷&FW&VEÔ²¶–æBÒvFFVBs²æÖW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâvFFVE6†VWG2r‚’’’ÒÀ¢¶÷&FW&VEÔ²¶–æBÒw&VÖ÷fVBs²æÖW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâw&VÖ÷fVE6†VWG2r‚’’’ÒÀ¢¶÷&FW&VEÔ²¶–æBÒwVæ¶æ÷vâs²æÖW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâwVæ¶æ÷vå6†VWG2r‚’’’ÒÀ¢¶÷&FW&VEÔ²¶–æBÒwVæ6†ævVBs²æÖW2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6ö×&—6öâwVæ6†ævVE6†VWG2r‚’’’Ð¢’’°¢f÷&V6‚‚G&tæÖR–â‚Fw&÷WææÖW2’’°¢FæÖRÒ·7G&–æuÒG&tæÖP¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’’²6öçF–çVRÐ¢FÖ–æu&V6÷&G2³Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²&Vf÷&U6†VWDæÖRÒB†–b…·7G&–æuÒFw&÷Wæ¶–æBÖæRvFFVBr’²FæÖRÒVÇ6R²rrÒ“²gFW%6†VWDæÖRÒB†–b…·7G&–æuÒFw&÷Wæ¶–æBÖæRw&VÖ÷fVBr’²FæÖRÒVÇ6R²rrÒ“²F—7Æ”æÖRÒFæÖS²¶–æBÒ·7G&–æuÒFw&÷Wæ¶–æC²ÖF6„6öæf–FVæ6RÒã²ÖF6„ÖWF†öBÒvÆVv7’ÖæÖRs²ÖW76vRÒrrÐ¢Ð¢Ð¢Ð¢f÷&V6‚‚FÖ–ær–âFÖ–æu&V6÷&G2’°¢F&Vf÷&TæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖ–ærv&Vf÷&U6†VWDæÖRrrr¢FgFW$æÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖ–ærvgFW%6†VWDæÖRrrr¢FæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖ–ærvF—7Æ”æÖRrB†–b‚FgFW$æÖR’²FgFW$æÖRÒVÇ6R²F&Vf÷&TæÖRÒ’¢F¶–æBÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖ–ærv¶–æBrwVæ¶æ÷vâr¢F–FVçF—G’Ò"F&Vf÷&TæÖVâFgFW$æÖVâF¶–æB ¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FæÖR’Ö÷"G6VVâÖ6öçF–ç2F–FVçF—G’’²6öçF–çVRÐ¢G6VVâ³ÒF–FVçF—G¢FÖ–ætÖW76vRÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖ–ærvÖW76vRrrr¢F—FVÒÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢6†VWDæÖRÒFæÖP¢&Vf÷&U6†VWDæÖRÒF&Vf÷&TæÖP¢gFW%6†VWDæÖRÒFgFW$æÖP¢6†VWD¶W’ÒvWBÔF–fe6†VWD¶W’F–FVçF—G¢¶–æBÒF¶–æ@¢ÖF6„6öæf–FVæ6RÒ¶F÷V&ÆUÒ„vWBÔFF&÷W'G’FÖ–ærvÖF6„6öæf–FVæ6Rr¢ÖF6„ÖWF†öBÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÖ–ærvÖF6„ÖWF†öBrrr¢&Vf÷&UvW2Ò ¢gFW%vW2Ò ¢vT6÷VçBÒ ¢Væ6†ævVEvTçVÖ&W'2Ò‚¢&Vv–öä6÷VçBÒ ¢2Dn8þ8Þ8î8î8îKùÞhÈ8~8ŠŽzK®KŠÞ8î89®8;Î8+Ž888).89n8:ž8*n8+n8~høþyK¾8;¾jùN‹È>8ž8(¾8 ¢7FGW2Òw&VG’p¢ÖW76vRÒB†–b‚FÖ–ætÖW76vR’²FÖ–ætÖW76vRÒVÇ6V–b‚F¶–æBÖWwVæ¶æ÷vâr’²~Zûî[ùÎ8).z+®Zé®8~8Þ8®8N8þ8(8[znXû>8åDn8).yºîŠinz+®Š¨Þ8~8n8þ88^8N8"rÒVÇ6R²~ŠŽzK®8~8þ89®8;Î8+Ž8).89n8:ž8*n8+n8~jùN‹È>8~8î8ž8"rÒ¢6öæf—&ÖVBÒFfÇ6P¢vW2Ò‚¢Ð¢–b‚F¶–æBÖWwVæ¶æ÷vâr’²F—FVÒç7FGW2ÒwVæ¶æ÷vârÐ¢F—FV×2³ÒF—FVÐ¢Ð¢F7W'&VçD†6†W2ÒvWBÔFF&÷W'G’D6öçFW‡Bv7W'&VçEf—7VÄ†6†W2rFçVÆÀ¢F&6VÆ–æT†6†W2ÒvWBÔFF&÷W'G’D6öçFW‡Bv&6VÆ–æUf—7VÄ†6†W2rFçVÆÀ¢2ˆz®X¹^jùN‹È>8~8ô6öçFW‡N8~ŠªÞ8)>888þ88>8+~8:^8).XhÞXŠžyJŽ8ž8(¾8.[^jÛNjùN‹È>8®8žiÊ®ŠŠÞZé®8î{XÎ‹zþ88¢28>8>8sY¹îŠªÞ8þ8YÎ8ŽX[iÈ”¥4ôî8Ž8î˜xÞŠH~8*.8*þ8+¾8+ž8).˜þ88(¾8 ¢–b…¶&ööÅÒD6öçFW‡Bæf–Æ&ÆR’°¢–b‚FçVÆÂÖWF7W'&VçD†6†W2’°¢F7W'&VçD†6†W2ÒvWBÕf—7VÄ†6†W2DÆæwVvR…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒD6öçFW‡Bæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒD6öçFW‡Bæ7W'&VçEfW'6–öä–B¢Ð¢–b‚FçVÆÂÖWF&6VÆ–æT†6†W2’°¢F&6VÆ–æT†6†W2ÒvWBÕf—7VÄ†6†W2DÆæwVvR…·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒD6öçFW‡Bæ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒD6öçFW‡Bæ&6VÆ–æUfW'6–öä–B¢Ð¢Ð¢F7W'&VçDÖÒvWBÔF–fd†6…6†VWDÖF7W'&VçD†6†W0¢F&6VÆ–æTÖÒvWBÔF–fd†6…6†VWDÖF&6VÆ–æT†6†W0¢f÷&V6‚‚F—FVÒ–âF—FV×2’°¢F&Vf÷&U6†VWBÒFçVÆÀ¢FgFW%6†VWBÒFçVÆÀ¢–b‚F&6VÆ–æTÖä6öçF–ç4¶W’…·7G&–æuÒF—FVÒæ&Vf÷&U6†VWDæÖR’’°¢F&Vf÷&U6†VWBÒF&6VÆ–æTÖµ·7G&–æuÒF—FVÒæ&Vf÷&U6†VWDæÖUÐ¢F—FVÒæ&Vf÷&UvW2ÒvWBÔ–çDFF&÷W'G’F&Vf÷&U6†VWBwvT6÷VçBr ¢Ð¢–b‚F7W'&VçDÖä6öçF–ç4¶W’…·7G&–æuÒF—FVÒægFW%6†VWDæÖR’’°¢FgFW%6†VWBÒF7W'&VçDÖµ·7G&–æuÒF—FVÒægFW%6†VWDæÖUÐ¢F—FVÒægFW%vW2ÒvWBÔ–çDFF&÷W'G’FgFW%6†VWBwvT6÷VçBr ¢Ð¢F—FVÒçvT6÷VçBÒ´ÖF…Ó£¤Ö‚…¶–çEÒF—FVÒæ&Vf÷&UvW2Â¶–çEÒF—FVÒægFW%vW2 ¢2#Ež8;µ$t.8îjÚ>ŠhþXÉnyK¾{J88þ88>8+~8:^8ÎZèÎXZŽKˆˆ{N8ž8(¾89®8;Î8+Ž8þ8[zîXˆnš	ŽYùþŠz>ié8).yÈyZ^8~8Þ8(¾8 ¢2yú^Ši®88þ88>8+~8:^8þKÛþ8(þ8¥4„Ó#Sn8îZèÎXZŽKˆˆ{N888).hêyJŽ8ž8(¾8þ8(8ZHži»N8îŠh¾‰Þ8Ž8~8þy›®yIþ8~8®8N8 ¢F&Vf÷&T†6†W2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F&Vf÷&U6†VWBwvT†6†W2r‚’’¢FgFW$†6†W2Ò„vWBÔ'&’„vWBÔFF&÷W'G’FgFW%6†VWBwvT†6†W2r‚’’¢G6ÖUvW2Ò‚¢F6ö×&&ÆUvW2Ò´ÖF…Ó£¤Ö–â‚F&Vf÷&T†6†W2ä6÷VçBÂFgFW$†6†W2ä6÷VçB¢f÷"‚GvT–æFW‚Ò²GvT–æFW‚ÖÇBF6ö×&&ÆUvW3²GvT–æFW‚²²’°¢F&Vf÷&T†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒF&Vf÷&T†6†W5²GvT–æFW…Ò¢FgFW$†6‚Òæ÷&ÖÆ—¦RÔf–ÆT†6‚…·7G&–æuÒFgFW$†6†W5²GvT–æFW…Ò¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&Vf÷&T†6‚’ÖæBF&Vf÷&T†6‚ÖWFgFW$†6‚’°¢G6ÖUvW2³Ò‚GvT–æFW‚²¢Ð¢Ð¢F—FVÒçVæ6†ævVEvTçVÖ&W'2Ò‚G6ÖUvW2¢Ð¢G&Wf–WrÒvWBÔF–fe&Wf–Wu7FFTf÷$6öçFW‡BDÆæwVvRD6öçFW‡@¢F6öæf—&ÖVE6†VWD¶W—2Ò·Ð¢f÷&V6‚‚G6†VWD¶W’–â„vWBÔ'&’„vWBÔFF&÷W'G’G&Wf–Wrv6öæf—&ÖVE6†VWD¶W—2r‚’’’’²F6öæf—&ÖVE6†VWD¶W—5µ·7G&–æuÒG6†VWD¶W•ÒÒGG'VRÐ¢f÷&V6‚‚F—FVÒ–âF—FV×2’²F—FVÒæ6öæf—&ÖVBÒF6öæf—&ÖVE6†VWD¶W—2ä6öçF–ç4¶W’…·7G&–æuÒF—FVÒç6†VWD¶W’’Ð¢FÖöF–f–VD6÷VçBÒ‚F—FV×2Âv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWvÖöF–f–VBrÒ’ä6÷Vç@¢FFFVD6÷VçBÒ‚F—FV×2Âv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWvFFVBrÒ’ä6÷Vç@¢G&VÖ÷fVD6÷VçBÒ‚F—FV×2Âv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWw&VÖ÷fVBrÒ’ä6÷Vç@¢GVæ¶æ÷vä6÷VçBÒ‚F—FV×2Âv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWwVæ¶æ÷vârÒ’ä6÷Vç@¢GVæ6†ævVD6÷VçBÒ‚F—FV×2Âv†W&RÔö&¦V7B²·7G&–æuÒEòæ¶–æBÖWwVæ6†ævVBrÒ’ä6÷Vç@¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢66†VÖfW'6–öâÒ¢Æv÷&—F†ÕfW'6–öâÒE67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öà¢7FGW2ÒB†–b…¶&ööÅÒD6öçFW‡Bæf–Æ&ÆR’²w&VG’rÒVÇ6R²wVæf–Æ&ÆRrÒ¢ÖW76vRÒ·7G&–æuÒD6öçFW‡BæÖW76vP¢v÷&¶&öö´–BÒ·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´–@¢v÷&¶&öö´æÖRÒ·7G&–æuÒD6öçFW‡Bçv÷&¶&öö´æÖP¢6÷W&6UG—RÒ·7G&–æuÒ„vWBÔFF&÷W'G’D6öçFW‡Bw6÷W&6UG—RrvW†6VÂr¢Væ—DÆ&VÂÒB†–b…·7G&–æuÒ„vWBÔFF&÷W'G’D6öçFW‡Bw6÷W&6UG—RrvW†6VÂr’ÖWw÷vW'ö–çBr’²~8+ž8:ž8*N88’rÒVÇ6V–b…·7G&–æuÒ„vWBÔFF&÷W'G’D6öçFW‡Bw6÷W&6UG—RrvW†6VÂr’Ö–â‚wv÷&BrÂwFbr’’²~89®8;Î8+‚rÒVÇ6R²~8+~8;Î88‚rÒ¢6ö×&—6öâÒ¶÷&FW&VEÔ°¢&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æU6æ6†÷D–@¢&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æUfW'6–öä–@¢7W'&VçE6æ6†÷D–BÒ·7G&–æuÒD6öçFW‡Bæ7W'&VçE6æ6†÷D–@¢7W'&VçEfW'6–öä–BÒ·7G&–æuÒD6öçFW‡Bæ7W'&VçEfW'6–öä–@¢&6VÆ–æTBÒ·7G&–æuÒD6öçFW‡Bæ&6VÆ–æT@¢7W'&VçDBÒ·7G&–æuÒD6öçFW‡Bæ7W'&VçD@¢6ö×&VDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv6ö×&VDBrrr¢ÖWF†öBÒ·7G&–æuÒD6öçFW‡BæÖWF†ö@¢66÷RÒ·7G&–æuÒD6öçFW‡Bç66÷P¢6öæf–FVæ6RÒ¶F÷V&ÆUÒ„vWBÔFF&÷W'G’F6ö×&—6öâv6öæf–FVæ6Rrã¢Ð¢7VÖÖ'’Ò¶÷&FW&VEÔ°¢6†ævVBÒFÖöF–f–VD6÷Vç@¢FFVBÒFFFVD6÷Vç@¢&VÖ÷fVBÒG&VÖ÷fVD6÷Vç@¢Væ¶æ÷vâÒGVæ¶æ÷vä6÷Vç@¢Væ6†ævVBÒGVæ6†ævVD6÷Vç@¢6öæf—&ÖVBÒ‚F—FV×2Âv†W&RÔö&¦V7B²¶&ööÅÒEòæ6öæf—&ÖVBÒ’ä6÷Vç@¢Ð¢&Wf–WrÒG&Wf–Wp¢vVæW&F–öâÒ¶÷&FW&VEÔ²7FGW2Òv6ö×ÆWFVBs²¦ö$–BÒrs²W&6VçBÒ²ÖW76vRÒ~ŠŽzK®89®8;Î8+Ž8).89n8:ž8*n8+n8~jùN‹È>8~8î8ž8"s²7W'&VçE6†VWBÒrrÐ¢6†VWG2Ò‚F—FV×2¢vVæW&FVDBÒrp¢Ð§Ð ¦gVæ7F–öâvWBÔÆö6ÄF–fdFWF–Ä66†UF‚€¢·7G&–æuÒDÆæwVvRÀ¢·7G&–æuÒEv÷&¶&öö´–BÀ¢·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÒrrÀ¢·7G&–æuÒD7W'&VçE6æ6†÷D–BÒrrÀ¢·7G&–æuÒD&6VÆ–æUfW'6–öä–BÒrrÀ¢·7G&–æuÒD7W'&VçEfW'6–öä–BÒrrÀ¢·7G&–æuÒE66÷RÒrp¢’°¢G6fUv÷&¶&öö´–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEv÷&¶&öö´–Bwv÷&¶&öö´–Bp¢G66÷RÒB†–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E66÷R’’²E66÷RÒVÇ6V–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D&6VÆ–æU6æ6†÷D–B’ÖæBÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D7W'&VçE6æ6†÷D–B’’²v†—7F÷'’rÒVÇ6R²vWFöÖF–2rÒ¢G—"Ò„vWBÕ6†#SeFW‡B‚'³×Ç³×Ç³'×Ç³7×Ç³G×Ç³WÒ"Ö`¢G66÷RÂD&6VÆ–æU6æ6†÷D–BÂD&6VÆ–æUfW'6–öä–BÀ¢D7W'&VçE6æ6†÷D–BÂD7W'&VçEfW'6–öä–BÂE67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öâ’’å7V'7G&–ærƒrÂ#B¢&WGW&â„¦ö–âÕF‚E67&—C¤Æö6Å'VçF–ÖT66†U&ö÷B‚&F–fb×³Ò×³Ò×³'Òæ§6öâ"ÖbDÆæwVvRÂG6fUv÷&¶&öö´–BÂG—"’§Ð ¦gVæ7F–öâvWBÔÆö6ÄF–fdFWF–Ä66†T–FVçF—G’…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÒrrÂ·7G&–æuÒD7W'&VçE6æ6†÷D–BÒrr’°¢F†4&6VÆ–æRÒÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D&6VÆ–æU6æ6†÷D–B¢F†47W'&VçBÒÖæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D7W'&VçE6æ6†÷D–B¢–b‚F†4&6VÆ–æR×†÷"F†47W'&VçB’²&WGW&âFçVÆÂÐ¢–b‚F†4&6VÆ–æRÖæBF†47W'&VçB’°¢F&6VÆ–æUfW'6–öä–BÒvWBÕ&VfW'&VD†—7F÷'•&VæFW%fW'6–öâDÆæwVvREv÷&¶&öö´–BD&6VÆ–æU6æ6†÷D–@¢F7W'&VçEfW'6–öä–BÒvWBÕ&VfW'&VD†—7F÷'•&VæFW%fW'6–öâDÆæwVvREv÷&¶&öö´–BD7W'&VçE6æ6†÷D–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&6VÆ–æUfW'6–öä–B’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7W'&VçEfW'6–öä–B’’²&WGW&âFçVÆÂÐ¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢66÷RÒv†—7F÷'’p¢&6VÆ–æU6æ6†÷D–BÒD&6VÆ–æU6æ6†÷D–C²&6VÆ–æUfW'6–öä–BÒF&6VÆ–æUfW'6–öä–@¢7W'&VçE6æ6†÷D–BÒD7W'&VçE6æ6†÷D–C²7W'&VçEfW'6–öä–BÒF7W'&VçEfW'6–öä–@¢Ð¢Ð¢G7G'V7GW&RÒvWBÕ7G'V7GW&RDÆæwVvP¢Gv÷&¶&öö²Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWEv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚Gv÷&¶&öö²ä6÷VçBÖW’²&WGW&âFçVÆÂÐ¢F6ö×&—6öâÒvWBÔÆFW7D6ö×&—6öâDÆæwVvREv÷&¶&öö´–BGv÷&¶&ööµ³Ð¢–b‚FçVÆÂÖWF6ö×&—6öâÖ÷"·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâw7FGW2rrr’ÖæRv6ö×ÆWFRr’²&WGW&âFçVÆÂÐ¢F–FVçF—G’Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢66÷RÒvWFöÖF–2p¢&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv&6VÆ–æU6æ6†÷D–Brrr¢&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv&6VÆ–æUfW'6–öä–Brrr¢7W'&VçE6æ6†÷D–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&ööµ³ÒvÆ7E&VæFW&VE6æ6†÷D–Brrr¢7W'&VçEfW'6–öä–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&ööµ³ÒvÆ7E&VæFW&VEfW'6–öä–Brrr¢Ð¢f÷&V6‚‚Ff–VÆB–â‚v&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R…·7G&–æuÒ„vWBÔFF&÷W'G’F–FVçF—G’Ff–VÆBrr’’’²&WGW&âFçVÆÂÐ¢Ð¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv7W'&VçE6æ6†÷D–Brrr’ÖæR·7G&–æuÒF–FVçF—G’æ7W'&VçE6æ6†÷D–BÖ÷ ¢·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv7W'&VçEfW'6–öä–Brrr’ÖæR·7G&–æuÒF–FVçF—G’æ7W'&VçEfW'6–öä–B’²&WGW&âFçVÆÂÐ¢&WGW&âF–FVçF—G§Ð ¦gVæ7F–öâ6fRÔÆö6ÄF–fdFWF–Ä66†R…·7G&–æuÒDÆæwVvRÂDFWF–ÂÂ·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÒrrÂ·7G&–æuÒD7W'&VçE6æ6†÷D–BÒrr’°¢–b‚FçVÆÂÖWDFWF–Â’²&WGW&âÐ¢Gv÷&¶&öö´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’DFWF–Âwv÷&¶&öö´–Brrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Gv÷&¶&öö´–B’’²&WGW&âÐ¢F6ö×&—6öâÒvWBÔFF&÷W'G’DFWF–Âv6ö×&—6öârFçVÆÀ¢–b‚FçVÆÂÖWF6ö×&—6öâ’²&WGW&âÐ¢F&6VÆ–æU6æ6†÷BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv&6VÆ–æU6æ6†÷D–BrD&6VÆ–æU6æ6†÷D–B¢F7W'&VçE6æ6†÷BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv7W'&VçE6æ6†÷D–BrD7W'&VçE6æ6†÷D–B¢F&6VÆ–æUfW'6–öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv&6VÆ–æUfW'6–öä–Brrr¢F7W'&VçEfW'6–öâÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâv7W'&VçEfW'6–öä–Brrr¢G66÷RÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâw66÷RrB†–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D&6VÆ–æU6æ6†÷D–B’’²v†—7F÷'’rÒVÇ6R²vWFöÖF–2rÒ’¢GF‚ÒvWBÔÆö6ÄF–fdFWF–Ä66†UF‚DÆæwVvRGv÷&¶&öö´–BF&6VÆ–æU6æ6†÷BF7W'&VçE6æ6†÷BF&6VÆ–æUfW'6–öâF7W'&VçEfW'6–öâG66÷P¢F66†T¶W’Ò…´”òåF…Ó£¤vWDgVÆÅF‚‚GF‚’’åFôÆ÷vW$–çf&–çB‚¢E67&—C¤F–fdFWF–Å&W7öç6T66†U²F66†T¶W•ÒÒDFWF–À¢G'’²w&—FRÔ§6öäf–ÆRGF‚DFWF–ÂÒ6F6‚²Ð§Ð ¦gVæ7F–öâvWBÔÆö6ÄF–fdFWF–Ä66†R…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÒrrÂ·7G&–æuÒD7W'&VçE6æ6†÷D–BÒrr’°¢F–FVçF—G’ÒvWBÔÆö6ÄF–fdFWF–Ä66†T–FVçF—G’DÆæwVvREv÷&¶&öö´–BD&6VÆ–æU6æ6†÷D–BD7W'&VçE6æ6†÷D–@¢–b‚FçVÆÂÖWF–FVçF—G’’²&WGW&âFçVÆÂÐ¢GF‚ÒvWBÔÆö6ÄF–fdFWF–Ä66†UF‚DÆæwVvREv÷&¶&öö´–B…·7G&–æuÒF–FVçF—G’æ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒF–FVçF—G’æ7W'&VçE6æ6†÷D–B’…·7G&–æuÒF–FVçF—G’æ&6VÆ–æUfW'6–öä–B’…·7G&–æuÒF–FVçF—G’æ7W'&VçEfW'6–öä–B’…·7G&–æuÒF–FVçF—G’ç66÷R¢F66†T¶W’Ò…´”òåF…Ó£¤vWDgVÆÅF‚‚GF‚’’åFôÆ÷vW$–çf&–çB‚¢FFWF–ÂÒE67&—C¤F–fdFWF–Å&W7öç6T66†U²F66†T¶W•Ð¢–b‚FçVÆÂÖWFFWF–ÂÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚GF‚’’°¢G'’²FFWF–ÂÒ&VBÔ§6öäf–ÆRGF‚FçVÆÂÒ6F6‚²FFWF–ÂÒFçVÆÂÐ¢–b‚FçVÆÂÖæRFFWF–Â’²E67&—C¤F–fdFWF–Å&W7öç6T66†U²F66†T¶W•ÒÒFFWF–ÂÐ¢Ð¢–b‚FçVÆÂÖWFFWF–ÂÖ÷"„vWBÔ–çDFF&÷W'G’FFWF–ÂvÆv÷&—F†ÕfW'6–öâr’ÖæRE67&—C¤F–fdFWF–ÄÆv÷&—F†ÕfW'6–öâÖ÷ ¢·7G&–æuÒ„vWBÔFF&÷W'G’FFWF–Âwv÷&¶&öö´–Brrr’ÖæREv÷&¶&öö´–B’²&WGW&âFçVÆÂÐ¢F6ö×&—6öâÒvWBÔFF&÷W'G’FFWF–Âv6ö×&—6öârFçVÆÀ¢–b‚FçVÆÂÖWF6ö×&—6öâ’²&WGW&âFçVÆÂÐ¢f÷&V6‚‚Ff–VÆB–â‚w66÷RrÂv&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–Br’’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’F6ö×&—6öâFf–VÆBrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’F–FVçF—G’Ff–VÆBrr’’²&WGW&âFçVÆÂÐ¢Ð¢&WGW&âFFWF–À§Ð ¦gVæ7F–öâV&Æ—6‚ÔÆFW7D6ö×&—6öä66†W2…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂD6ö×&—6öâÂD7W'&VçD†6†W2ÂD&6VÆ–æT†6†W2’°¢–b‚FçVÆÂÖWD6ö×&—6öâÖ÷"·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâw7FGW2rrr’ÖæRv6ö×ÆWFRr’²&WGW&âÐ¢G7VÖÖ'’Ò¶÷&FW&VEÔ°¢7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâw7FGW2rrr¢ÖWF†öBÒ·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâvÖWF†öBrrr¢6†ævVE6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’D6ö×&—6öâv6†ævVE6†VWG2r‚’’¢Væ6†ævVE6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’D6ö×&—6öâwVæ6†ævVE6†VWG2r‚’’¢Væ¶æ÷vå6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’D6ö×&—6öâwVæ¶æ÷vå6†VWG2r‚’’¢FFVE6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’D6ö×&—6öâvFFVE6†VWG2r‚’’¢&VÖ÷fVE6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’D6ö×&—6öâw&VÖ÷fVE6†VWG2r‚’’¢ÖW76vRÒ·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâvÖW76vRrrr¢6ö×&VDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâv6ö×&VDBrrr¢Ð¢Gv÷&¶&öö²ÒWFFRÕ7G'V7GW&TÆö6¶VBDÆæwVvR°¢&Ò‚G7G'V7GW&R¢FÖF6†W2Ò„vWBÔ'&’G7G'V7GW&Rçv÷&¶&öö·2Âv†W&RÔö&¦V7B²·7G&–æuÒEòçv÷&¶&öö´–BÖWEv÷&¶&öö´–BÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚FÖF6†W2ä6÷VçBÖW’²&WGW&âFçVÆÂÐ¢6WBÔæ÷FU&÷W'G’FÖF6†W5³ÒvÆFW7D6ö×&—6öå7VÖÖ'’rG7VÖÖ'¢&WGW&âFÖF6†W5³Ð¢Ð¢–b‚FçVÆÂÖWGv÷&¶&öö²’²&WGW&âÐ¢FF—7Æ”æÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²vF—7Æ”æÖRrrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FF—7Æ”æÖR’’²FF—7Æ”æÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gv÷&¶&öö²vf–ÆTæÖRrEv÷&¶&öö´–B’Ð¢F6öçFW‡BÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢f–Æ&ÆRÒGG'VS²7FGW2Òvf–Æ&ÆRs²ÖW76vRÒrp¢v÷&¶&öö´–BÒEv÷&¶&öö´–C²v÷&¶&öö´æÖRÒFF—7Æ”æÖS²v÷&¶&öö²ÒGv÷&¶&öö°¢6ö×&—6öâÒD6ö×&—6öã²6ö×&—6öåW'6—7FVBÒGG'VP¢7W'&VçE6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–@¢7W'&VçEfW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ7W'&VçEfW'6–öä–@¢&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–@¢&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒD6ö×&—6öâæ&6VÆ–æUfW'6–öä–@¢7W'&VçEf—7VÄ†6†W2ÒD7W'&VçD†6†W3²&6VÆ–æUf—7VÄ†6†W2ÒD&6VÆ–æT†6†W0¢7W'&VçDBÒvWBÔF–fe6æ6†÷DFFRDÆæwVvREv÷&¶&öö´–B…·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒD6ö×&—6öâæ6ö×&VDB¢&6VÆ–æTBÒvWBÔF–fe6æ6†÷DFFRDÆæwVvREv÷&¶&öö´–B…·7G&–æuÒD6ö×&—6öâæ&6VÆ–æU6æ6†÷D–B’rp¢ÖWF†öBÒ·7G&–æuÒ„vWBÔFF&÷W'G’D6ö×&—6öâvÖWF†öBrrr¢66÷RÒvWFöÖF–2p¢Ð¢FFWF–ÂÒæWrÔF–fdFWF–Å6¶VÆWFöâDÆæwVvRF6öçFW‡@¢6fRÔÆö6ÄF–fdFWF–Ä66†RDÆæwVvRFFWF–À¢FÆFW7D¶W’Ò‚DÆæwVvR²wÂr²Ev÷&¶&öö´–B²wÂr²·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–B²wÂr²·7G&–æuÒD6ö×&—6öâæ7W'&VçEfW'6–öä–B’åFôÆ÷vW$–çf&–çB‚¢E67&—C¤ÆFW7D6ö×&—6öä66†U²FÆFW7D¶W•ÒÒD6ö×&—6öà¢·fö–EÒ…WFFRÕ6æ6†÷E7VÖÖ'”66†TVçG'’DÆæwVvREv÷&¶&öö´–B…·7G&–æuÒD6ö×&—6öâæ7W'&VçE6æ6†÷D–B’§Ð ¦gVæ7F–öâvWBÔF–fdFWF–Â€¢·7G&–æuÒDÆæwVvRÀ¢·7G&–æuÒEv÷&¶&öö´–BÀ¢·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÒrrÀ¢·7G&–æuÒD7W'&VçE6æ6†÷D–BÒrp¢’°¢F66†VBÒvWBÔÆö6ÄF–fdFWF–Ä66†RDÆæwVvREv÷&¶&öö´–BD&6VÆ–æU6æ6†÷D–BD7W'&VçE6æ6†÷D–@¢–b‚FçVÆÂÖæRF66†VB’°¢6WBÔæ÷FU&÷W'G’F66†VBwW&f÷&Öæ6Rr…¶÷&FW&VEÔ²6öçFW‡D×2Ò²6¶VÆWFöä×2Ò²F÷FÄ×2Ò²6÷W&6RÒvÆö6ÂÖ66†RrÒ¢&WGW&âF66†V@¢Ð¢2jùN‹È5ä~8Ž[zîXˆd¥4ôî8þK¨¾X˜ÞyIþh‰8~8®8N8.Zûî‹x˜Ž8;¾8+~8;Î88Ž8;¾89®8;Î8+Ži[888).‹ùN8~8¢2ŠŽzK®KŠÞ8ã89®8;Î8+Ž8)%Dbæ§>8…vV"v÷&¶W.8~89n8:ž8*n8+nXh^jùN‹È>8ž8(¾8 ¢GF÷FÅF–ÖW"Ò´F–væ÷7F–72å7F÷vF6…Ó£¥7F'DæWr‚¢F6öçFW‡BÒvWBÔF–fdFWF–Ä6öçFW‡BDÆæwVvREv÷&¶&öö´–BD&6VÆ–æU6æ6†÷D–BD7W'&VçE6æ6†÷D–@¢F6öçFW‡D×2ÒGF÷FÅF–ÖW"äVÆ6VDÖ–ÆÆ—6V6öæG0¢FFWF–ÂÒæWrÔF–fdFWF–Å6¶VÆWFöâDÆæwVvRF6öçFW‡@¢GF÷FÅF–ÖW"å7F÷‚¢6WBÔæ÷FU&÷W'G’FFWF–ÂwW&f÷&Öæ6Rr…¶÷&FW&VEÔ°¢6öçFW‡D×2ÒF6öçFW‡D×0¢6¶VÆWFöä×2Ò´ÖF…Ó£¤Ö‚ƒÂGF÷FÅF–ÖW"äVÆ6VDÖ–ÆÆ—6V6öæG2ÒF6öçFW‡D×2¢F÷FÄ×2ÒGF÷FÅF–ÖW"äVÆ6VDÖ–ÆÆ—6V6öæG0¢6÷W&6RÒw6†&VBÖfÆÆ&6²p¢Ò¢6fRÔÆö6ÄF–fdFWF–Ä66†RDÆæwVvRFFWF–ÂD&6VÆ–æU6æ6†÷D–BD7W'&VçE6æ6†÷D–@¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D&6VÆ–æU6æ6†÷D–B’ÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D7W'&VçE6æ6†÷D–B’Öæ@¢¶&ööÅÒF6öçFW‡Bæf–Æ&ÆRÖæB·7G&–æuÒF6öçFW‡Bç66÷RÖWvWFöÖF–2r’°¢G'’²V&Æ—6‚ÔÆFW7D6ö×&—6öä66†W2DÆæwVvREv÷&¶&öö´–BF6öçFW‡Bæ6ö×&—6öâF6öçFW‡Bæ7W'&VçEf—7VÄ†6†W2F6öçFW‡Bæ&6VÆ–æUf—7VÄ†6†W2Ò6F6‚²Ð¢Ð¢&WGW&âFFWF–À§Ð ¦gVæ7F–öâ&W6öÇfRÔF–fd6öçFVçEFeF‚…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒE6æ6†÷D–BÂ·7G&–æuÒEfW'6–öä–BÂ·7G&–æuÒE6†VWDæÖR’°¢27G&–7B–FVçF—G“¢æWfW"fÆÂ&6²Fòæ÷F†W"&VæFW"fW'6–öââ†6†W2æBF—7Æ–VBDb×W7B&RF†R6ÖRvVæW&F–öâà¢G6fU6æ6†÷D–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBE6æ6†÷D–Bw6æ6†÷D–Bp¢G6fUfW'6–öä–BÒ76W'BÕ6fU7F÷&vU6VvÖVçBEfW'6–öä–BwfW'6–öä–Bp¢–b„„vWBÕ&VæFW%fW'6–öä–G2DÆæwVvREv÷&¶&öö´–BG6fU6æ6†÷D–B’Öæ÷F6öçF–ç2G6fUfW'6–öä–B’²&WGW&ârrÐ¢&WGW&â…&W6öÇfRÔ6öçFVçEFe6†VWEF„W†7BDÆæwVvREv÷&¶&öö´–BG6fUfW'6–öä–BE6†VWDæÖR§Ð ¦gVæ7F–öâ–çfö¶RÔF–fd–ÖvUvTvVæW&F–öâ…·7G&–æuÒD&Vf÷&UFbÂ·7G&–æuÒDgFW%FbÂ·7G&–æuÒD÷WGWDF—&V7F÷'’Â·7G&–æuÒD¶–æB’°¢G67&—EF‚Ò¦ö–âÕF‚E67&—C¤&ö÷BwFööÇ5ÆF–fbÖ–ÖvR×vW2ç3p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G67&—EF‚’’²F‡&÷r~[zîXˆnyK¾X8þyIþh‰88N8;Î8:¾8ÎŠh¾8N8¾8(®8î8¾8)>8"rÐ¢F&wVÖVçG2Ò°¢÷WGWDF—&V7F÷'’ÒD÷WGWDF—&V7F÷'¢¶–æBÒD¶–æ@¢G’ÒE67&—C¤F–fdFWF–ÄG¢F‡&W6†öÆBÒE67&—C¤F–fdFWF–ÅF‡&W6†öÆ@¢Ö–æ–×VÕ&Vv–öå—†VÇ2ÒE67&—C¤F–fdFWF–ÄÖ–æ–×VÕ&Vv–öå—†VÇ0¢FF–ærÒE67&—C¤F–fdFWF–ÅFF–æp¢Ð¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚D&Vf÷&UFb’’²F&wVÖVçG2ä&Vf÷&UFbÒD&Vf÷&UFbÐ¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚DgFW%Fb’’²F&wVÖVçG2ägFW%FbÒDgFW%FbÐ¢G&rÒ‚bG67&—EF‚&wVÖVçG2Âf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ¢GFW‡BÒ‚G&rÖ¦ö–ârr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GFW‡B’’²F‡&÷r~[zîXˆnyK¾X8þyIþh‰{YiéÎ8).Xùn[é~8~8Þ8î8¾8)>8~8~8þ8"rÐ¢&WGW&â‚GFW‡BÂ6öçfW'Dg&öÒÔ§6öâ§Ð ¦gVæ7F–öâ–çfö¶RÔF–fd–ÖvT&F6„vVæW&F–öâ‚D—FV×2’°¢G67&—EF‚Ò¦ö–âÕF‚E67&—C¤&ö÷BwFööÇ5ÆF–fbÖ–ÖvRÖ&F6‚ç3p¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G67&—EF‚’’²F‡&÷r~[zîXˆnyK¾X8þKˆhºÎyIþh‰88N8;Î8:¾8ÎŠh¾8N8¾8(®8î8¾8)>8"rÐ¢G&WVW7EF‚Ò¦ö–âÕF‚…´”òåF…Ó£¤vWEFV×F‚‚’’‚w&"ÖF–fb×&WVW7BÒr²´wV–EÓ£¤æWtwV–B‚’åFõ7G&–ær‚târ’²ræ§6öâr¢G'’°¢w&—FRÔ§6öäf–ÆRG&WVW7EF‚…¶÷&FW&VEÔ²66†VÖfW'6–öâÒ²—FV×2Ò‚D—FV×2’Ò¢G&rÒ‚bG67&—EF‚Õ&WVW7EF‚G&WVW7EF‚ÔG’E67&—C¤F–fdFWF–ÄG’ ¢ÕF‡&W6†öÆBE67&—C¤F–fdFWF–ÅF‡&W6†öÆBÔÖ–æ–×VÕ&Vv–öå—†VÇ2E67&—C¤F–fdFWF–ÄÖ–æ–×VÕ&Vv–öå—†VÇ2 ¢ÕFF–ærE67&—C¤F–fdFWF–ÅFF–ærÂf÷$V6‚Ôö&¦V7B²·7G&–æuÒEòÒ¢GFW‡BÒ‚G&rÖ¦ö–ârr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚GFW‡B’’²F‡&÷r~[zîXˆnyK¾X8þKˆhºÎyIþh‰{YiéÎ8).Xùn[é~8~8Þ8î8¾8)>8~8~8þ8"rÐ¢G&W7VÇBÒGFW‡BÂ6öçfW'Dg&öÒÔ§6öà¢–b‚Öæ÷B¶&ööÅÒ„vWBÔFF&÷W'G’G&W7VÇBvö²rFfÇ6R’’²F‡&÷r~[zîXˆnyK¾X8þ8îKˆhºÎyIþh‰8¾ZKiY~8~8î8~8þ8"rÐ¢&WGW&âG&W7VÇ@¢Òf–æÆÇ’°¢&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚G&WVW7EF‚Ôf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVP¢Ð§Ð  ¦gVæ7F–öâ7F'BÔF–fdFWF–Ä¦ö"€¢·7G&–æuÒDÆæwVvRÀ¢·7G&–æuÒEv÷&¶&öö´–BÀ¢·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÒrrÀ¢·7G&–æuÒD7W'&VçE6æ6†÷D–BÒrrÀ¢·7G&–æuÒE6†VWD¶W’Òrp¢’°¢F6öçFW‡BÒvWBÔF–fdFWF–Ä6öçFW‡BDÆæwVvREv÷&¶&öö´–BD&6VÆ–æU6æ6†÷D–BD7W'&VçE6æ6†÷D–@¢–b‚Öæ÷B¶&ööÅÒF6öçFW‡Bæf–Æ&ÆR’°¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²ö²ÒGG'VS²¦ö$–BÒrs²7FGW2ÒwVæf–Æ&ÆRs²W&6VçBÒ²ÖW76vRÒ·7G&–æuÒF6öçFW‡BæÖW76vRÐ¢Ð¢G6fU6†VWD¶W’Òrp¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E6†VWD¶W’’’²G6fU6†VWD¶W’Ò76W'BÕ6fU7F÷&vU6VvÖVçBE6†VWD¶W’w6†VWD¶W’rÐ¢FÆVæ6„Æö6²ÒvWBÔF–fdÆVæ6„Æö6µF‚DÆæwVvRF6öçFW‡@¢&WGW&â–çfö¶RÕv—F„Æö6²FÆVæ6„Æö6²°¢Fg&W6„6öçFW‡BÒvWBÔF–fdFWF–Ä6öçFW‡BDÆæwVvREv÷&¶&öö´–BD&6VÆ–æU6æ6†÷D–BD7W'&VçE6æ6†÷D–@¢–b‚Öæ÷B¶&ööÅÒFg&W6„6öçFW‡Bæf–Æ&ÆR’²F‡&÷r·7G&–æuÒFg&W6„6öçFW‡BæÖW76vRÐ¢f÷&V6‚‚Ff–VÆB–â‚w66÷RrÂv7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–BrÂv&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–Br’’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’F6öçFW‡BFf–VÆBrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’Fg&W6„6öçFW‡BFf–VÆBrr’’°¢F‡&÷r~jùN‹È>Zûî‹8Îi»Nik8^8(Î8î8~8þ8.[zîXˆnŠ›>{K8).™h¾8Þy»N8~8n8þ88^8N8"p¢Ð¢Ð¢F6öçFW‡BÒFg&W6„6öçFW‡@¢G6¶VÆWFöâÒæWrÔF–fdFWF–Å6¶VÆWFöâDÆæwVvRF6öçFW‡@¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6fU6†VWD¶W’’Öæ@¢‚G6¶VÆWFöâç6†VWG2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç6†VWD¶W’ÖWG6fU6†VWD¶W’Ò’ä6÷VçBÖW’°¢F‡&÷r~hÈ~Zé®8~8þ8+~8;Î88Ž8þjùN‹È>Zûî‹8¾Y
¾8î8(Î8n8N8î8¾8)>8"p¢Ð¢F66†TF—"ÒvWBÔF–fdFWF–Ä66†TF—$f÷$ÆæwVvRDÆæwVvRF6öçFW‡@¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚F66†TF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚F66†TF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢FFWF–ÅF‚Ò¦ö–âÕF‚F66†TF—"vF–fbÖFWF–Âæ§6öâp¢FW†—7F–ærÒFçVÆÀ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FFWF–ÅF‚’²G'’²FW†—7F–ærÒ&VBÔ§6öäf–ÆRFFWF–ÅF‚FçVÆÂÒ6F6‚²ÒÐ¢–b…FW7BÔF–fdFWF–ÄÖF6†W46öçFW‡BFW†—7F–ærF6öçFW‡B’°¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6fU6†VWD¶W’’ÖæB·7G&–æuÒ„vWBÔFF&÷W'G’FW†—7F–ærw7FGW2rrr’ÖWw&VG’r’°¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²ö²ÒGG'VS²¦ö$–BÒrs²7FGW2Òv6ö×ÆWFVBs²W&6VçBÒ²ÖW76vRÒ~[zîXˆnŠ›>{K8þKÙÎh‰kˆŽ8þ8~8ž8"s²FWF–Å&VG’ÒGG'VRÐ¢Ð¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G6fU6†VWD¶W’’’°¢GF&vWBÒ„vWBÔ'&’FW†—7F–ærç6†VWG2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç6†VWD¶W’ÖWG6fU6†VWD¶W’ÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚GF&vWBä6÷VçBÖwBÖæB·7G&–æuÒ„vWBÔFF&÷W'G’GF&vWE³Òw7FGW2rrr’ÖWw&VG’r’°¢&WGW&â·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²ö²ÒGG'VS²¦ö$–BÒrs²7FGW2Òv6ö×ÆWFVBs²W&6VçBÒ²ÖW76vRÒ~8>8î8+~8;Î88Ž8î[zîXˆnyK¾X8þ8þKÙÎh‰kˆŽ8þ8~8ž8"s²FWF–Å&VG’ÒGG'VRÐ¢Ð¢Ð¢Ð¢Gö–çFW%F‚Ò¦ö–âÕF‚F66†TF—"vF–fbÖ¦ö"æ§6öâp¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚Gö–çFW%F‚’°¢G'’°¢Gö–çFW"Ò&VBÔ§6öäf–ÆRGö–çFW%F‚FçVÆÀ¢F7F—fT–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’Gö–çFW"v¦ö$–Brrr¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F7F—fT–B’’°¢F7F—fRÒ&VBÕ&VæFW$¦ö%7FGW2DÆæwVvRF7F—fT–@¢–b„‚v6ö×ÆWFVBrÂv6ö×ÆWFVB×v—F‚ÖW'&÷'2rÂvf–ÆVBrÂvÖ—76–ærrÂv6æ6VÆÆVBr’Öæ÷F6öçF–ç2·7G&–æuÒ„vWBÔFF&÷W'G’F7F—fRw7FGW2rrr’’°¢6WBÔæ÷FU&÷W'G’F7F—fRv¦ö–æVDW†—7F–ætF–fd¦ö"rGG'VP¢&WGW&âF7F—fP¢Ð¢Ð¢Ò6F6‚²Ð¢Ð¢F¦ö$–BÒv¦ö%òr²„vWBÔFFR’åFõ7G&–ær‚w———”ÔÖFEô„†Ö×72r’²uòr²…´wV–EÓ£¤æWtwV–B‚’åFõ7G&–ær‚târ’å7V'7G&–ærƒÃ‚’¢FÆV6W2Ò„æWrÔF–fd¦ö$ÆV6W2DÆæwVvRF6öçFW‡BF¦ö$–B¢F¦ö$F—"ÒvWBÕ&VæFW$¦ö$F—"DÆæwVvP¢F–çWEF‚Ò¦ö–âÕF‚F¦ö$F—""F¦ö$–BæF–fbæ–çWBæ§6öâ ¢G7FGW5F‚Ò¦ö–âÕF‚F¦ö$F—""F¦ö$–Bç7FGW2æ§6öâ ¢G7FF÷WEF‚Ò¦ö–âÕF‚F¦ö$F—""F¦ö$–BæF–fbæ÷WBæÆör ¢G7FFW'%F‚Ò¦ö–âÕF‚F¦ö$F—""F¦ö$–BæF–fbæW'"æÆör ¢2jùN‹È>yK¾X8þ8þ8+~8;Î88ŽXÙŽKØÞ8~˜^[»nyIþh‰8ž8(¾8þ8(88+Ž8:~89n8îZûî‹8þiÈZJs8+~8;Î88Ž8 ¢G6†VWD6÷VçBÒB†–b‚G6¶VÆWFöâç6†VWG2ä6÷VçBÖwB’²ÒVÇ6R²Ò¢F–æ—F–ÂÒ·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢ö²ÒGG'VS²¦ö%G—RÒvF–fbÖFWF–Âs²¦ö$–BÒF¦ö$–C²7FGW2ÒwVWVVBs²F÷FÂÒG6†VWD6÷Vç@¢6ö×ÆWFVBÒ²f–ÆVBÒ²W&6VçBÒ²ÖW76vRÒ~[zîXˆnŠ›>{K8).k©nX)ž8~8n8N8î8ž8"p¢7W'&VçEv÷&¶&öö´–BÒ·7G&–æuÒF6öçFW‡Bçv÷&¶&öö´–C²7W'&VçEv÷&¶&öö´æÖRÒ·7G&–æuÒF6öçFW‡Bçv÷&¶&öö´æÖP¢7W'&VçE6†VWBÒrs²&ö6W74–BÒ²7FF÷WEF‚ÒG7FF÷WEFƒ²7FFW'%F‚ÒG7FFW'%F€¢&W7VÇG2Ò‚“²W'&÷'2Ò‚“²7F'FVDBÒæWrÔæ÷t—6ó²WFFVDBÒæWrÔæ÷t—6ó²7FFU6fVDBÒrp¢Ð¢G'’°¢–b…·7G&–æuÒF6öçFW‡Bç66÷RÖWv†—7F÷'’rÖæBÖæ÷B¶&ööÅÒF6öçFW‡Bæ6ö×&—6öåW'6—7FVB’°¢G6fVBÒ6fRÔ†—7F÷&–6Å6æ6†÷D6ö×&—6öâDÆæwVvR…·7G&–æuÒF6öçFW‡Bçv÷&¶&öö´–B’F6öçFW‡Bæ6ö×&—6öà¢F6öçFW‡Bæ6ö×&—6öâÒG6fV@¢F6öçFW‡Bæ6ö×&—6öåW'6—7FVBÒGG'VP¢Ð¢w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚F–æ—F–À¢w&—FRÔ§6öäf–ÆRF–çWEF‚…¶÷&FW&VEÔ°¢¦ö$–BÒF¦ö$–C²ÖöFRÒDÆæwVvS²v÷&¶&öö´–BÒ·7G&–æuÒF6öçFW‡Bçv÷&¶&öö´–@¢7W'&VçE6æ6†÷D–BÒ·7G&–æuÒF6öçFW‡Bæ7W'&VçE6æ6†÷D–C²7W'&VçEfW'6–öä–BÒ·7G&–æuÒF6öçFW‡Bæ7W'&VçEfW'6–öä–@¢&6VÆ–æU6æ6†÷D–BÒ·7G&–æuÒF6öçFW‡Bæ&6VÆ–æU6æ6†÷D–C²&6VÆ–æUfW'6–öä–BÒ·7G&–æuÒF6öçFW‡Bæ&6VÆ–æUfW'6–öä–@¢66÷RÒ·7G&–æuÒF6öçFW‡Bç66÷S²6†VWD¶W’ÒG6fU6†VWD¶W¢66†TF—"ÒF66†TF—#²FWF–ÅF‚ÒFFWF–ÅFƒ²7FGW5F‚ÒG7FGW5F€¢7FF÷WEF‚ÒG7FF÷WEFƒ²7FFW'%F‚ÒG7FFW'%Fƒ²ÆV6W2Ò‚FÆV6W2¢vVæW&F–öäÆö6µF‚Ò„vWBÔF–fdvVæW&F–öäÆö6µF‚DÆæwVvRF6öçFW‡B¢Ò¢w&—FRÔ§6öäf–ÆRGö–çFW%F‚…¶÷&FW&VEÔ²¦ö$–BÒF¦ö$–C²6†VWD¶W’ÒG6fU6†VWD¶W“²7&VFVDBÒæWrÔæ÷t—6òÒ¢G4W†RÒ¦ö–âÕF‚FVçc¥7—7FVÕ&ö÷Bu7—7FVÓ3%Åv–æF÷w5÷vW%6†VÆÅÇcãÇ÷vW'6†VÆÂæW†Rp¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚G4W†R’’²G4W†RÒw÷vW'6†VÆÂæW†RrÐ¢G6W'fW%67&—BÒ¦ö–âÕF‚E67&—C¤&ö÷Bw6W'fW"ç3p¢F6öÖÖæBÒ"brB‚G6W'fW%67&—Bå&WÆ6R‚"r"Â"rr"’’rÔÖöFRrB‚DÆæwVvRå&WÆ6R‚"r"Â"rr"’’rÔF–fd¦ö%F‚rB‚F–çWEF‚å&WÆ6R‚"r"Â"rr"’’r ¢G&ö2Ò7F'BÔ†–FFVå÷vW%6†VÆÄ6†–ÆBG4W†RF6öÖÖæBG7FF÷WEF‚G7FFW'%F€¢–b‚Öæ÷BG&ö2Ö÷"Öæ÷BG&ö2ä–B’²F‡&÷r~[zîXˆnyK¾X8þ8îKÙÎh‰89~8:Þ8+¾8+””N8).Xùn[é~8~8Þ8î8¾8)>8~8~8þ8"rÐ¢F–æ—F–Âç&ö6W74–BÒ¶–çEÒG&ö2ä–C²F–æ—F–Âç7FGW2ÒvÆVæ6†–ærs²F–æ—F–ÂçW&6VçBÒ ¢F–æ—F–ÂæÖW76vRÒ~[zîXˆnyK¾X8þ8îKÙÎh‰89~8:Þ8+¾8+ž8).‹[~X¹^8~8î8~8þ8"p¢w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚F–æ—F–À¢&WGW&âF–æ—F–À¢Ò6F6‚°¢F–æ—F–Âç7FGW2Òvf–ÆVBs²F–æ—F–ÂçW&6VçBÒ²F–æ—F–ÂæÖW76vRÒ~[zîXˆnyK¾X8þ8îKÙÎh‰89~8:Þ8+¾8+ž8).‹[~X¹^8~8Þ8î8¾8)>8~8~8þ8"p¢F–æ—F–ÂæW'&÷'2Ò…¶÷&FW&VEÔ²W'&÷"ÒEòäW†6WF–öâäÖW76vS²FWF–ÂÒvWBÔW'&÷$FWF–ÂEòÒ¢G'’²w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚F–æ—F–ÂÒ6F6‚²Ð¢&VÖ÷fRÔF–fd¦ö$ÆV6W2…·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ²ÖöFSÒDÆæwVvS²v÷&¶&öö´–CÕ·7G&–æuÒF6öçFW‡Bçv÷&¶&öö´–C²ÆV6W3Ô‚FÆV6W2’Ò¢F‡&÷p¢Ð¢Ð§Ð  ¦gVæ7F–öâ–çfö¶RÔF–fdFWF–Ä¦ö$6÷&R‚D¦ö"’°¢FÆæwVvRÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"vÖöFRrDÖöFR¢Gv÷&¶&öö´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"wv÷&¶&öö´–Brrr¢G7FGW5F‚Ò·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"w7FGW5F‚rrr¢FFWF–ÅF‚Ò·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"vFWF–ÅF‚rrr¢F66†TF—"Ò´”òåF…Ó£¤vWDgVÆÅF‚…·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"v66†TF—"rrr’¢F¦ö$–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"v¦ö$–Brrr¢G&WVW7FVE6†VWD¶W’Ò·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"w6†VWD¶W’rrr¢G7FGW2Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢ö²ÒGG'VS²¦ö%G—RÒvF–fbÖFWF–Âs²¦ö$–BÒF¦ö$–C²7FGW2Òw'Vææ–ærs²F÷FÂÒ²6ö×ÆWFVBÒ²f–ÆVBÒ ¢W&6VçBÒ3²ÖW76vRÒ~jùN‹È>Zûî‹8).z+®Š¨Þ8~8n8N8î8ž8"s²7W'&VçEv÷&¶&öö´–BÒGv÷&¶&öö´–C²7W'&VçEv÷&¶&öö´æÖRÒrp¢7W'&VçE6†VWBÒrs²&ö6W74–BÒµ7—7FVÒäF–væ÷7F–72å&ö6W75Ó£¤vWD7W'&VçE&ö6W72‚’ä–@¢&W7VÇG2Ò‚“²W'&÷'2Ò‚“²7F'FVDBÒæWrÔæ÷t—6ó²WFFVDBÒæWrÔæ÷t—6ó²7FFU6fVDBÒrp¢Ð¢w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚G7FGW0¢FFWF–ÂÒFçVÆÀ¢G'’°¢&Vg&W6‚ÔF–fd¦ö$ÆV6W2D¦ö ¢F¦ö%66÷RÒ·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"w66÷RrvWFöÖF–2r¢F6öçFW‡BÒ–b‚F¦ö%66÷RÖWv†—7F÷'’r’°¢vWBÔF–fdFWF–Ä6öçFW‡BFÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒD¦ö"æ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒD¦ö"æ7W'&VçE6æ6†÷D–B¢ÒVÇ6R²vWBÔF–fdFWF–Ä6öçFW‡BFÆæwVvRGv÷&¶&öö´–BÐ¢–b‚Öæ÷B¶&ööÅÒF6öçFW‡Bæf–Æ&ÆR’²F‡&÷r·7G&–æuÒF6öçFW‡BæÖW76vRÐ¢f÷&V6‚‚Ff–VÆB–â‚v7W'&VçE6æ6†÷D–BrÂv7W'&VçEfW'6–öä–BrÂv&6VÆ–æU6æ6†÷D–BrÂv&6VÆ–æUfW'6–öä–BrÂw66÷Rr’’°¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’D¦ö"Ff–VÆBrr’ÖæR·7G&–æuÒ„vWBÔFF&÷W'G’F6öçFW‡BFf–VÆBrr’’°¢F‡&÷r~jùN‹È>Zûî‹8Îi»Nik8^8(Î8î8~8þ8.ZHži»N8988>8+Ž8).™h¾8Þy»N8~8n8þ88^8N8"p¢Ð¢Ð¢FW‡V7FVD66†RÒ´”òåF…Ó£¤vWDgVÆÅF‚‚„vWBÔF–fdFWF–Ä66†TF—$f÷$ÆæwVvRFÆæwVvRF6öçFW‡B’¢–b‚FW‡V7FVD66†RÖæRF66†TF—"’²F‡&÷r~[zîXˆn8*Þ8:>88>8+~8:^8îKùÞZÙŽXXŽ8ÎKˆÞjÚ>8~8ž8"rÐ¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚F66†TF—"’’²æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚F66†TF—"Ôf÷&6RÂ÷WBÔçVÆÂÐ¢2XZŽKÙ>XhÞŠšnŠÎ8~8(.iz.ZÙŽŠ›>{K8).ŠªÞ8þ‹ëÎ8þ8f–ÆVN8¾8®8>8þZHži»N8®8~8+~8;Î88Ž8).KùÞhÈ8ž8(¾8 ¢2ŠªÞ8þ‹ëÎ8)>8Xh^Zëž8þy»N[èÎ8åFW7BÔF–fdFWF–ÄÖF6†W46öçFW‡N8~jùN‹È>Zûî‹8Ž8îZèÎXZŽKˆˆ{N8).jIÎŠ‹Î8ž8(¾8 ¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚FFWF–ÅF‚’°¢FFWF–ÂÒ&VBÔ§6öäf–ÆRFFWF–ÅF‚FçVÆÀ¢Ð¢–b‚Öæ÷B…FW7BÔF–fdFWF–ÄÖF6†W46öçFW‡BFFWF–ÂF6öçFW‡B’’°¢FFWF–ÂÒæWrÔF–fdFWF–Å6¶VÆWFöâFÆæwVvRF6öçFW‡@¢Ð¢FÆÅ6†VWG2Ò„vWBÔ'&’FFWF–Âç6†VWG2¢Gv÷&´–æFW†W2Ò‚¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&WVW7FVE6†VWD¶W’’’°¢2izuvV"Tž8®8—6†VWD¶Wž8).˜8(ž8®8NYÎ8>X{®8~8~8(.8XZŽ8+~8;Î88ŽKˆhºÎyIþh‰8¾8þh‹¾8^8®8N8 ¢2ZHži»N8+~8;Î88Ž8).XJ®XXŽ8~88+Ž8:~89n8¾8N8Ó8+~8;Î88Ž88Xznyn8ž8(¾8 ¢G&VfW'&VD–æFW‚ÒÓ¢f÷"‚FâÒ²FâÖÇBFÆÅ6†VWG2ä6÷VçC²Fâ²²’°¢G6†VWD¶–æBÒ·7G&–æuÒ„vWBÔFF&÷W'G’FÆÅ6†VWG5²FåÒv¶–æBrrr¢G6†VWE7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’FÆÅ6†VWG5²FåÒw7FGW2rrr¢–b‚G6†VWD¶–æBÖæRwVæ6†ævVBrÖæB‚wVæF–ærrÂvFVfW'&VBrÂvf–ÆVBr’Ö6öçF–ç2G6†VWE7FGW2’°¢G&VfW'&VD–æFW‚ÒFà¢'&V°¢Ð¢Ð¢–b‚G&VfW'&VD–æFW‚ÖÇB’°¢f÷"‚FâÒ²FâÖÇBFÆÅ6†VWG2ä6÷VçC²Fâ²²’°¢G6†VWE7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’FÆÅ6†VWG5²FåÒw7FGW2rrr¢–b„‚wVæF–ærrÂvFVfW'&VBrÂvf–ÆVBr’Ö6öçF–ç2G6†VWE7FGW2’²G&VfW'&VD–æFW‚ÒFã²'&V²Ð¢Ð¢Ð¢–b‚G&VfW'&VD–æFW‚ÖvR’²Gv÷&´–æFW†W2³ÒG&VfW'&VD–æFW‚Ð¢ÒVÇ6R°¢f÷"‚FâÒ²FâÖÇBFÆÅ6†VWG2ä6÷VçC²Fâ²²’²–b…·7G&–æuÒFÆÅ6†VWG5²FåÒç6†VWD¶W’ÖWG&WVW7FVE6†VWD¶W’’²Gv÷&´–æFW†W2³ÒFã²'&V²ÒÐ¢–b‚Gv÷&´–æFW†W2ä6÷VçBÖW’²F‡&÷r~hÈ~Zé®8~8þš^yºî8þjùN‹È>Zûî‹8¾Y
¾8î8(Î8n8N8î8¾8)>8"rÐ¢Ð¢FFWF–Âç6†VWG2Ò‚FÆÅ6†VWG2¢FFWF–Âç7FGW2ÒvvVæW&F–ærp¢FFWF–ÂævVæW&F–öâÒ¶÷&FW&VEÔ²7FGW2Òw'Vææ–ærs²¦ö$–BÒF¦ö$–C²W&6VçBÒ3²ÖW76vRÒ~[zîXˆnyK¾X8þ8).KÙÎh‰8~8n8N8î8ž8"s²7W'&VçE6†VWBÒrrÐ¢w&—FRÔ§6öäf–ÆRFFWF–ÅF‚FFWF–À¢G7FGW2çF÷FÂÒGv÷&´–æFW†W2ä6÷VçC²G7FGW2æ7W'&VçEv÷&¶&öö´æÖRÒ·7G&–æuÒF6öçFW‡Bçv÷&¶&öö´æÖP¢w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚G7FGW0 ¢2ZHži»N8+~8;Î88Ž8N8Ž8²¦fõDd&÷‚8)#.Y¹î‹[~X¹^8~8n8N8þiz~{XÎ‹zþ8).˜þ88(¾8 ¢2XZŽ8+~8;Î88Ž8îikizuDn8)#8N8ä¥dÞ8ŽkŠ8~8iÈZJsNKŠnX‰~8~8:ž8+ž8+þ8:ž8*N8+®8~8n8¾8(žKˆhºÎŠz>ié8ž8(¾8 ¢F&F6…&WVW7BÒ‚¢F&F6„–D'”–æFW‚Ò·Ð¢F&F6…&W&F–öäW'&÷'2Ò·Ð¢f÷"‚G÷6—F–öâÒ²G÷6—F–öâÖÇBGv÷&´–æFW†W2ä6÷VçC²G÷6—F–öâ²²’°¢F’Ò¶–çEÒGv÷&´–æFW†W5²G÷6—F–öåÐ¢G6†VWBÒFFWF–Âç6†VWG5²F•Ð¢FæÖRÒ·7G&–æuÒG6†VWBç6†VWDæÖP¢F&Vf÷&TæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBv&Vf÷&U6†VWDæÖRrFæÖR¢FgFW$æÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWBvgFW%6†VWDæÖRrFæÖR¢F¶–æBÒ·7G&–æuÒG6†VWBæ¶–æ@¢G'’°¢F&Vf÷&UFbÒrs²FgFW%FbÒrp¢–b‚F¶–æBÖæRvFFVBr’²F&Vf÷&UFbÒ&W6öÇfRÔF–fd6öçFVçEFeF‚FÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒF6öçFW‡Bæ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒF6öçFW‡Bæ&6VÆ–æUfW'6–öä–B’F&Vf÷&TæÖRÐ¢–b‚F¶–æBÖæRw&VÖ÷fVBr’²FgFW%FbÒ&W6öÇfRÔF–fd6öçFVçEFeF‚FÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒF6öçFW‡Bæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒF6öçFW‡Bæ7W'&VçEfW'6–öä–B’FgFW$æÖRÐ¢FÖ—76–ærÒ‚‚F¶–æBÖ–â‚vÖöF–f–VBrÂwVæ6†ævVBr’ÖæB…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&Vf÷&UFb’Ö÷"·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FgFW%Fb’’’Ö÷ ¢‚F¶–æBÖWvFFVBrÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FgFW%Fb’’Ö÷"‚F¶–æBÖWw&VÖ÷fVBrÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&Vf÷&UFb’’¢–b‚FÖ—76–ær’²F‡&÷r~YÎKˆ8:Î8;>888:®8;>8+K‰nKº>8æ6öçFVçBDn8ÎŠh¾8N8¾8(®8î8¾8)>8"rÐ¢–b‚F¶–æBÖWwVæ¶æ÷vârÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&Vf÷&UFb’ÖæB·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚FgFW%Fb’’²F‡&÷r~jùN‹È>XX>8;¾jùN‹È>XXŽ8åDn8).z+®Š¨Þ8~8Þ8î8¾8)>8"rÐ¢G6†VWDF—"Ò¦ö–âÕF‚F66†TF—"„¦ö–âÕF‚wr…·7G&–æuÒG6†VWBç6†VWD¶W’’¢–b…FW7BÕF‚ÔÆ—FW&ÅF‚G6†VWDF—"’²&VÖ÷fRÔ—FVÒÔÆ—FW&ÅF‚G6†VWDF—"Õ&V7W'6RÔf÷&6RÔW'&÷$7F–öâ6–ÆVçFÇ”6öçF–çVRÐ¢æWrÔ—FVÒÔ—FVÕG—RF—&V7F÷'’ÕF‚G6†VWDF—"Ôf÷&6RÂ÷WBÔçVÆÀ¢F&F6„–BÒv’r²G÷6—F–öâåFõ7G&–ær‚sr¢F&F6„–D'”–æFW…µ·7G&–æuÒF•ÒÒF&F6„–@¢F&F6…&WVW7B³Ò¶÷&FW&VEÔ°¢–BÒF&F6„–@¢&Vf÷&UFbÒF&Vf÷&UF`¢gFW%FbÒFgFW%F`¢&Vf÷&U&7FW$F—&V7F÷'’ÒB†–b‚F¶–æBÖæRvFFVBr’°¢vWBÕ&VæFW%&7FW%6†VWDF—"FÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒF6öçFW‡Bæ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒF6öçFW‡Bæ&6VÆ–æUfW'6–öä–B’F&Vf÷&TæÖP¢ÒVÇ6R²rrÒ¢gFW%&7FW$F—&V7F÷'’ÒB†–b‚F¶–æBÖæRw&VÖ÷fVBr’°¢vWBÕ&VæFW%&7FW%6†VWDF—"FÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒF6öçFW‡Bæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒF6öçFW‡Bæ7W'&VçEfW'6–öä–B’FgFW$æÖP¢ÒVÇ6R²rrÒ¢&Vf÷&UvT6÷VçBÒvWBÔ–çDFF&÷W'G’G6†VWBv&Vf÷&UvW2r ¢gFW%vT6÷VçBÒvWBÔ–çDFF&÷W'G’G6†VWBvgFW%vW2r ¢Væ6†ævVEvTçVÖ&W'2Ò„vWBÔ'&’„vWBÔFF&÷W'G’G6†VWBwVæ6†ævVEvTçVÖ&W'2r‚’’¢÷WGWDF—&V7F÷'’ÒG6†VWDF— ¢¶–æBÒF¶–æ@¢Ð¢Ò6F6‚°¢F&F6…&W&F–öäW'&÷'5µ·7G&–æuÒF•ÒÒEòäW†6WF–öâäÖW76vP¢Ð¢Ð¢F&F6…&W7VÇDÖÒ·Ð¢F&F6…F–Ö–æw2ÒFçVÆÀ¢–b‚F&F6…&WVW7Bä6÷VçBÖwB’°¢&Vg&W6‚ÔF–fd¦ö$ÆV6W2D¦ö ¢G7FGW2æÖW76vRÒ.ikizuDn8).8î8Ž8(8nyK¾X8þXÉn8;¾Šz>ié8~8n8N8î8žûÈ‚B‚F&F6…&WVW7Bä6÷VçBž8+~8;Î88ŽûÈž8" ¢G7FGW2æ7W'&VçE6†VWBÒrp¢G7FGW2çW&6VçBÒP¢FFWF–ÂævVæW&F–öâÒ¶÷&FW&VEÔ²7FGW2Òw'Vææ–ærs²¦ö$–BÒF¦ö$–C²W&6VçBÒS²ÖW76vRÒG7FGW2æÖW76vS²7W'&VçE6†VWBÒrrÐ¢w&—FRÔ§6öäf–ÆRFFWF–ÅF‚FFWF–Ã²w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚G7FGW0¢F&F6„vVæW&FVBÒ–çfö¶RÔF–fd–ÖvT&F6„vVæW&F–öâF&F6…&WVW7@¢F&F6…F–Ö–æw2ÒvWBÔFF&÷W'G’F&F6„vVæW&FVBwF–Ö–æw2rFçVÆÀ¢f÷&V6‚‚F&F6„—FVÒ–â„vWBÔ'&’„vWBÔFF&÷W'G’F&F6„vVæW&FVBv—FV×2r‚’’’’°¢F&F6…&W7VÇDÖµ·7G&–æuÒ„vWBÔFF&÷W'G’F&F6„—FVÒv–Brrr•ÒÒF&F6„—FVÐ¢Ð¢Ð ¢f÷"‚G÷6—F–öâÒ²G÷6—F–öâÖÇBGv÷&´–æFW†W2ä6÷VçC²G÷6—F–öâ²²’°¢&Vg&W6‚ÔF–fd¦ö$ÆV6W2D¦ö ¢F’Ò¶–çEÒGv÷&´–æFW†W5²G÷6—F–öåÐ¢G6†VWBÒFFWF–Âç6†VWG5²F•Ð¢FæÖRÒ·7G&–æuÒG6†VWBç6†VWDæÖS²F¶–æBÒ·7G&–æuÒG6†VWBæ¶–æ@¢G6†VWBç7FGW2ÒvvVæW&F–ærs²G6†VWBæÖW76vRÒrp¢G7FGW2æ7W'&VçE6†VWBÒFæÖP¢G7FGW2æÖW76vRÒ.[zîXˆnyK¾X8þ8).KÙÎh‰8~8n8N8î8“¢B‚G÷6—F–öâ²’òB‚Gv÷&´–æFW†W2ä6÷VçB’š^yºâFæÖR ¢G7FGW2çW&6VçBÒ¶–çEÕ´ÖF…Ó£¤Ö‚ƒRÂ´ÖF…Ó£¤Ö–âƒ“RÂ´ÖF…Ó£¤fÆö÷"‚‚…¶F÷V&ÆUÒG÷6—F–öâ’ò´ÖF…Ó£¤Ö‚ƒÂGv÷&´–æFW†W2ä6÷VçB’’¢“’²R’¢w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚G7FGW0¢G'’°¢–b‚F&F6…&W&F–öäW'&÷'2ä6öçF–ç4¶W’…·7G&–æuÒF’’’²F‡&÷r·7G&–æuÒF&F6…&W&F–öäW'&÷'5µ·7G&–æuÒF•ÒÐ¢F&F6„–BÒ·7G&–æuÒF&F6„–D'”–æFW…µ·7G&–æuÒF•Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚F&F6„–B’Ö÷"Öæ÷BF&F6…&W7VÇDÖä6öçF–ç4¶W’‚F&F6„–B’’°¢F‡&÷r~[zîXˆnyK¾X8þKˆhºÎyIþh‰{YiéÎ8¾Zûî‹8+~8;Î88Ž8Î8.8(®8î8¾8)>8"p¢Ð¢FvVæW&FVBÒF&F6…&W7VÇDÖ²F&F6„–EÐ¢–b‚Öæ÷B¶&ööÅÒ„vWBÔFF&÷W'G’FvVæW&FVBvö²rFfÇ6R’’°¢F‡&÷r·7G&–æuÒ„vWBÔFF&÷W'G’FvVæW&FVBvÖW76vRr~[zîXˆnyK¾X8þ8).yIþh‰8~8Þ8î8¾8)>8~8~8þ8"r¢Ð¢GvW2Ò‚“²G&Vv–öåF÷FÂÒ²F†5Væ¶æ÷våvRÒFfÇ6P¢f÷&V6‚‚GvR–â„vWBÔ'&’FvVæW&FVBçvW2’’°¢G&Vv–öç2Ò„vWBÔ'&’„vWBÔFF&÷W'G’GvRw&Vv–öç2r‚’’“²G&Vv–öåF÷FÂ³ÒG&Vv–öç2ä6÷Vç@¢–b…·7G&–æuÒ„vWBÔFF&÷W'G’GvRw7FGW2rrr’ÖWwVæ¶æ÷vâr’²F†5Væ¶æ÷våvRÒGG'VRÐ¢GvW2³Ò·67W7FöÖö&¦V7EÕ¶÷&FW&VEÔ°¢vTçVÖ&W"ÒvWBÔ–çDFF&÷W'G’GvRwvTçVÖ&W"r²v–GF‚ÒvWBÔ–çDFF&÷W'G’GvRwv–GF‚r²†V–v‡BÒvWBÔ–çDFF&÷W'G’GvRv†V–v‡Br ¢&Vf÷&UvTçVÖ&W"ÒvWBÔ–çDFF&÷W'G’GvRv&Vf÷&UvTçVÖ&W"r²gFW%vTçVÖ&W"ÒvWBÔ–çDFF&÷W'G’GvRvgFW%vTçVÖ&W"r ¢6ö×&—6öä¶–æBÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRv6ö×&—6öä¶–æBrrr“²ÖF6„ÖWF†öBÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRvÖF6„ÖWF†öBrrr¢Ö–ætÖ&–wV÷W2Ò¶&ööÅÒ„vWBÔFF&÷W'G’GvRvÖ–ætÖ&–wV÷W2rFfÇ6R“²Ö–ætÖW76vRÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRvÖ–ætÖW76vRrrr¢vU6—¦T6†ævVBÒ¶&ööÅÒ„vWBÔFF&÷W'G’GvRwvU6—¦T6†ævVBrFfÇ6R“²7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’GvRw7FGW2rw&VG’r¢ÖW76vRÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRvÖW76vRrrr“²6öæf–FVæ6RÒ¶F÷V&ÆUÒ„vWBÔFF&÷W'G’GvRv6öæf–FVæ6Rr¢6†ævVE&F–òÒ¶F÷V&ÆUÒ„vWBÔFF&÷W'G’GvRv6†ævVE&F–òr“²&Vf÷&T76WBÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRv&Vf÷&Tf–ÆRrrr¢gFW$76WBÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRvgFW$f–ÆRrrr“²&Vf÷&TÖ6´76WBÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRv&Vf÷&TÖ6´f–ÆRrrr¢&Vf÷&T÷fW&Æ”76WBÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRv&Vf÷&T÷fW&Æ”f–ÆRrrr“²Ö6´76WBÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRvgFW$Ö6´f–ÆRrrr¢÷fW&Æ”76WBÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvRvgFW$÷fW&Æ”f–ÆRrrr“²&Vv–öç2Ò‚G&Vv–öç2¢Ð¢Ð¢G6†VWBçvW2Ò‚GvW2“²G6†VWBæ&Vf÷&UvW2ÒvWBÔ–çDFF&÷W'G’FvVæW&FVBv&Vf÷&UvT6÷VçBr ¢G6†VWBægFW%vW2ÒvWBÔ–çDFF&÷W'G’FvVæW&FVBvgFW%vT6÷VçBr ¢G6†VWBçvT6÷VçBÒ´ÖF…Ó£¤Ö‚…¶–çEÒG6†VWBæ&Vf÷&UvW2Â¶–çEÒG6†VWBægFW%vW2“²G6†VWBç&Vv–öä6÷VçBÒG&Vv–öåF÷FÀ¢G6†VWBç7FGW2ÒB†–b‚F¶–æBÖWwVæ¶æ÷vârÖ÷"F†5Væ¶æ÷våvR’²wVæ¶æ÷vârÒVÇ6R²w&VG’rÒ¢–b‚F¶–æBÖWwVæ¶æ÷vâr’²G6†VWBæÖW76vRÒ~KúšÎ8~8Þ8(¾[zîXˆnš	ŽYùþ8).XŠNZé®8~8Þ8®8N8þ8(8[Ë~Š«þŠŽzK®8þŠÎ8N8î8¾8)>8"rÐ¢Ò6F6‚°¢2yK¾X8þyIþh‰8Þ8î8(.8î8îZKiY~8þ8jùN‹È>Kˆ®8î8ÎXŠNZé®KˆÞˆ;Þ8Þ8ŽXË®XŠ^8ž8(¾8 ¢2f–ÆVB8î8î8îjè¾8ž8>8Ž8~8XZŽKÙ>XhÞŠšnŠÎ8î8þ8þ8+~8;Î88ŽXhÞ˜Žh©î8¾8(žXhÞyIþh‰8~8Þ8(¾8 ¢G6†VWBç7FGW2Òvf–ÆVBs²G6†VWBæÖW76vRÒEòäW†6WF–öâäÖW76vS²G6†VWBçvW2Ò‚“²G6†VWBç&Vv–öä6÷VçBÒ ¢G7FGW2æf–ÆVB²³²G7FGW2æW'&÷'2Ò‚G7FGW2æW'&÷'2’²…¶÷&FW&VEÔ²6†VWDæÖRÒFæÖS²W'&÷"ÒEòäW†6WF–öâäÖW76vS²FWF–ÂÒvWBÔW'&÷$FWF–ÂEòÒ¢Ð¢G7FGW2æ6ö×ÆWFVBÒG÷6—F–öâ²¢G7FGW2çW&6VçBÒ¶–çEÕ´ÖF…Ó£¤Ö‚ƒ‚Â´ÖF…Ó£¤Ö–âƒ“‚Â´ÖF…Ó£¤fÆö÷"‚‚…¶F÷V&ÆUÒ‚G÷6—F–öâ²’’ò´ÖF…Ó£¤Ö‚ƒÂGv÷&´–æFW†W2ä6÷VçB’’¢“2’²R’¢FFWF–Âç6†VWG5²F•ÒÒG6†VW@¢FFWF–ÂævVæW&F–öâÒ¶÷&FW&VEÔ²7FGW2Òw'Vææ–ærs²¦ö$–BÒF¦ö$–C²W&6VçBÒG7FGW2çW&6VçC²ÖW76vRÒG7FGW2æÖW76vS²7W'&VçE6†VWBÒFæÖRÐ¢w&—FRÔ§6öäf–ÆRFFWF–ÅF‚FFWF–Ã²w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚G7FGW0¢Ð¢28+~8;Î88Ž888î˜^[»nyIþh‰8~8("w&VG’r8).i»Ž8N8n8N8þ8þ8(8XZŽ8+~8;Î88ŽyIþh‰8Î˜	NKŠÞ8~ZKiY~8~8þ[èÎ8°¢28ÎZHži»N8®8~8Þ8+~8;Î88Ž8)#K»n™h¾8þ8‚7FGW28Â&VG’8¾Kˆ®i»Ž8Þ8^8(Î8VæF–ær8î8î8îjè¾8>8ð¢28+~8;Î88Ž8ÎK¨Î[ªn8ŽyIþh‰8~8Þ8®8þ8®8>8n8N8ò…7F'BÔF–fdFWF–Ä¦ö"8Î8ÎKÙÎh‰kˆŽ8þ8Þ8).‹ùN8’ž8 ¢2Zéþ™©¾8îjè¾K»n8¾8(’7FGW28).k®8(8(¾8 ¢GVæF–æu6†VWG2Ò„vWBÔ'&’FFWF–Âç6†VWG2Âv†W&RÔö&¦V7B²‚wVæF–ærrÂvvVæW&F–ærr’Ö6öçF–ç2·7G&–æuÒ„vWBÔFF&÷W'G’Eòw7FGW2rrr’Ò’ä6÷Vç@¢Ff–ÆVE6†VWG2Ò„vWBÔ'&’FFWF–Âç6†VWG2Âv†W&RÔö&¦V7B²·7G&–æuÒ„vWBÔFF&÷W'G’Eòw7FGW2rrr’ÖWvf–ÆVBrÒ’ä6÷Vç@¢–b‚GVæF–æu6†VWG2ÖWÖæBFf–ÆVE6†VWG2ÖW’°¢FFWF–Âç7FGW2Òw&VG’s²FFWF–ÂæÖW76vRÒrp¢ÒVÇ6R°¢FFWF–Âç7FGW2Òvf–ÆVBp¢G'G2Ò‚¢–b‚GVæF–æu6†VWG2ÖwB’²G'G2³Ò.iÊ®KÙÎh‰GVæF–æu6†VWG2K»b"Ð¢–b‚Ff–ÆVE6†VWG2ÖwB’²G'G2³Ò.KÙÎh‰ZKiYrFf–ÆVE6†VWG2K»b"Ð¢FFWF–ÂæÖW76vRÒ‚~[zîXˆnyK¾X8þ8¾iÊ®ZèÎK¨n8îš^yºî8Î8.8(®8î8žûÈ‚r²‚G'G2Ö¦ö–â~8r’²~ûÈž8.XhÞŠšnŠÎ8~8n8þ88^8N8"r¢Ð¢FFWF–ÂævVæW&FVDBÒæWrÔæ÷t—6ð¢6WBÔæ÷FU&÷W'G’FFWF–ÂwW&f÷&Öæ6RrF&F6…F–Ö–æw0¢FFWF–ÂævVæW&F–öâÒ¶÷&FW&VEÔ²7FGW2Òv6ö×ÆWFVBs²¦ö$–BÒF¦ö$–C²W&6VçBÒ²ÖW76vRÒ~[zîXˆnŠ›>{K8).KÙÎh‰8~8î8~8þ8"s²7W'&VçE6†VWBÒrrÐ¢w&—FRÔ§6öäf–ÆRFFWF–ÅF‚FFWF–À¢G7FGW2ç7FGW2ÒB†–b‚G7FGW2æf–ÆVBÖwB’²v6ö×ÆWFVB×v—F‚ÖW'&÷'2rÒVÇ6R²v6ö×ÆWFVBrÒ“²G7FGW2çW&6VçBÒ²G7FGW2æ7W'&VçE6†VWBÒrp¢G7FGW2æÖW76vRÒB†–b‚G7FGW2æf–ÆVBÖwB’²~Kˆ˜:Ž8î8+~8;Î88Ž8).™šN8Þ8[zîXˆnŠ›>{K8).KÙÎh‰8~8î8~8þ8"rÒVÇ6R²~[zîXˆnŠ›>{K8).KÙÎh‰8~8î8~8þ8"rÒ¢G7FGW2ç&W7VÇG2Ò…¶÷&FW&VEÔ²v÷&¶&öö´–BÒGv÷&¶&öö´–C²FWF–ÅF‚ÒFFWF–ÅFƒ²6†VWD6÷VçBÒGv÷&´–æFW†W2ä6÷VçC²W&f÷&Öæ6RÒF&F6…F–Ö–æw2Ò¢w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚G7FGW0¢Ò6F6‚°¢G7FGW2ç7FGW2Òvf–ÆVBs²G7FGW2çW&6VçBÒ²G7FGW2æÖW76vRÒ~[zîXˆnŠ›>{K8).KÙÎh‰8~8Þ8î8¾8)>8~8~8þ8"p¢G7FGW2æW'&÷'2Ò…¶÷&FW&VEÔ²W'&÷"ÒEòäW†6WF–öâäÖW76vS²FWF–ÂÒvWBÔW'&÷$FWF–ÂEòÒ“²w&—FRÕ&VæFW$¦ö%7FGW2G7FGW5F‚G7FGW0¢–b‚FçVÆÂÖæRFFWF–Â’°¢G'’°¢FFWF–Âç7FGW2Òvf–ÆVBp¢–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&WVW7FVE6†VWD¶W’’’°¢GF&vWBÒ„vWBÔ'&’FFWF–Âç6†VWG2Âv†W&RÔö&¦V7B²·7G&–æuÒ„vWBÔFF&÷W'G’Eòw6†VWD¶W’rrr’ÖWG&WVW7FVE6†VWD¶W’ÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚GF&vWBä6÷VçBÖwB’°¢GF&vWE³Òç7FGW2Òvf–ÆVBp¢GF&vWE³ÒæÖW76vRÒEòäW†6WF–öâäÖW76vP¢GF&vWE³ÒçvW2Ò‚¢GF&vWE³Òç&Vv–öä6÷VçBÒ ¢Ð¢Ð¢FFWF–ÂæÖW76vRÒEòäW†6WF–öâäÖW76vP¢FFWF–ÂævVæW&F–öâÒ¶÷&FW&VEÔ²7FGW2Òvf–ÆVBs²¦ö$–BÒF¦ö$–C²W&6VçBÒ²ÖW76vRÒEòäW†6WF–öâäÖW76vS²7W'&VçE6†VWBÒrrÐ¢w&—FRÔ§6öäf–ÆRFFWF–ÅF‚FFWF–À¢Ò6F6‚²Ð¢Ð¢Ð§Ð  ¦gVæ7F–öâ–çfö¶RÔF–fdFWF–Ä¦ö$g&öÔf–ÆR…·7G&–æuÒD¦ö%F‚’°¢F¦ö"Ò&VBÔ§6öäf–ÆRD¦ö%F‚FçVÆÀ¢–b‚FçVÆÂÖWF¦ö"’²F‡&÷r$F–fb¦ö"f–ÆR—2æ÷B&VF&ÆS¢D¦ö%F‚"Ð¢G'’°¢FÆæwVvRÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¦ö"vÖöFRrDÖöFR¢Gv÷&¶&öö´–BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¦ö"wv÷&¶&öö´–Brrr¢F¦ö%66÷RÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¦ö"w66÷RrvWFöÖF–2r¢F6öçFW‡BÒ–b‚F¦ö%66÷RÖWv†—7F÷'’r’°¢vWBÔF–fdFWF–Ä6öçFW‡BFÆæwVvRGv÷&¶&öö´–B…·7G&–æuÒF¦ö"æ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒF¦ö"æ7W'&VçE6æ6†÷D–B¢ÒVÇ6R²vWBÔF–fdFWF–Ä6öçFW‡BFÆæwVvRGv÷&¶&öö´–BÐ¢–b‚Öæ÷B¶&ööÅÒF6öçFW‡Bæf–Æ&ÆR’²F‡&÷r·7G&–æuÒF6öçFW‡BæÖW76vRÐ¢FW‡V7FVDÆö6²Ò´”òåF…Ó£¤vWDgVÆÅF‚‚„vWBÔF–fdvVæW&F–öäÆö6µF‚FÆæwVvRF6öçFW‡B’¢G&÷f–FVDÆö6µFW‡BÒ·7G&–æuÒ„vWBÔFF&÷W'G’F¦ö"vvVæW&F–öäÆö6µF‚rrr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚G&÷f–FVDÆö6µFW‡B’’²F‡&÷r~[zîXˆnyIþh‰8:Þ88>8*þ8îKùÞZÙŽXXŽ8Î8.8(®8î8¾8)>8"rÐ¢G&÷f–FVDÆö6²Ò´”òåF…Ó£¤vWDgVÆÅF‚‚G&÷f–FVDÆö6µFW‡B¢–b‚FW‡V7FVDÆö6²ÖæRG&÷f–FVDÆö6²’²F‡&÷r~[zîXˆnyIþh‰8:Þ88>8*þ8îKùÞZÙŽXXŽ8ÎKˆÞjÚ>8~8ž8"rÐ¢–çfö¶RÕv—F„Æö6²FW‡V7FVDÆö6²²–çfö¶RÔF–fdFWF–Ä¦ö$6÷&RF¦ö"Ð¢Òf–æÆÇ’°¢&VÖ÷fRÔF–fd¦ö$ÆV6W2F¦ö ¢Ð§Ð ¦gVæ7F–öâ6W'fRÔF–fevR‚D6öçFW‡BÂ·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂ·7G&–æuÒD7W'&VçE6æ6†÷D–BÂ·7G&–æuÒD&6VÆ–æU6æ6†÷D–BÂ·7G&–æuÒE6†VWD¶W’Â·7G&–æuÒEvTçVÖ&W%FW‡BÂ·7G&–æuÒD76WBÂ·7G&–æuÒE66÷RÒvWFöÖF–2r’°¢FÆÆ÷vVD76WG2Ò‚v&Vf÷&RrÂvgFW"rÂv&Vf÷&RÖÖ6²rÂv&Vf÷&RÖ÷fW&Æ’rÂvÖ6²rÂv÷fW&Æ’r¢–b‚FÆÆ÷vVD76WG2Öæ÷F6öçF–ç2D76WB’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚v76WB8ÎKˆÞjÚ>8~8ž8"r’Ð¢GvTçVÖ&W"Ò ¢–b‚Öæ÷B¶–çEÓ£¥G'•'6R‚EvTçVÖ&W%FW‡BÂ·&VeÒGvTçVÖ&W"’Ö÷"GvTçVÖ&W"ÖÆR’°¢F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚wvTçVÖ&W"8ÎKˆÞjÚ>8~8ž8"r¢Ð¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E66÷R’’²E66÷RÒvWFöÖF–2rÐ¢–b‚E66÷RÖæ÷F–â‚vWFöÖF–2rÂv†—7F÷'’r’’²F‡&÷r´&wVÖVçDW†6WF–öåÓ£¦æWr‚w66÷R8ÎKˆÞjÚ>8~8ž8"r’Ð¢FF–fd6öçFW‡BÒ–b‚E66÷RÖWv†—7F÷'’r’°¢vWBÔF–fdFWF–Ä6öçFW‡BDÆæwVvREv÷&¶&öö´–BD&6VÆ–æU6æ6†÷D–BD7W'&VçE6æ6†÷D–@¢ÒVÇ6R°¢vWBÔF–fdFWF–Ä6öçFW‡BDÆæwVvREv÷&¶&öö´–@¢Ð¢–b‚Öæ÷B¶&ööÅÒFF–fd6öçFW‡Bæf–Æ&ÆR’²F‡&÷r·7G&–æuÒFF–fd6öçFW‡BæÖW76vRÐ¢–b‚„76W'BÕ6fU7F÷&vU6VvÖVçBD7W'&VçE6æ6†÷D–Bv7W'&VçE6æ6†÷D–Br’ÖæR·7G&–æuÒFF–fd6öçFW‡Bæ7W'&VçE6æ6†÷D–BÖ÷ ¢„76W'BÕ6fU7F÷&vU6VvÖVçBD&6VÆ–æU6æ6†÷D–Bv&6VÆ–æU6æ6†÷D–Br’ÖæR·7G&–æuÒFF–fd6öçFW‡Bæ&6VÆ–æU6æ6†÷D–B’°¢F‡&÷r~jùN‹È>Zûî‹8Îi»Nik8^8(Î8î8~8þ8.[zîXˆnŠ›>{K8).™h¾8Þy»N8~8n8þ88^8N8"p¢Ð¢G6fU6†VWD¶W’Ò76W'BÕ6fU7F÷&vU6VvÖVçBE6†VWD¶W’w6†VWD¶W’p¢F66†TF—"ÒvWBÔF–fdFWF–Ä66†TF—$f÷$ÆæwVvRDÆæwVvRFF–fd6öçFW‡@¢FFWF–ÅF‚Ò¦ö–âÕF‚F66†TF—"vF–fbÖFWF–Âæ§6öâp¢–b‚Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FFWF–ÅF‚’’²F‡&÷r~[zîXˆnŠ›>{K8þ8î8KÙÎh‰8^8(Î8n8N8î8¾8)>8"rÐ¢FFWF–ÂÒ&VBÔ§6öäf–ÆRFFWF–ÅF‚FçVÆÀ¢–b‚Öæ÷B…FW7BÔF–fdFWF–ÄÖF6†W46öçFW‡BFFWF–ÂFF–fd6öçFW‡B’’°¢F‡&÷r~[zîXˆnŠ›>{K8îjùN‹È>Zûî‹8ÎKˆˆ{N8~8î8¾8)>8.[zîXˆnŠ›>{K8).XhÞKÙÎh‰8~8n8þ88^8N8"p¢Ð¢G6†VWBÒ„vWBÔ'&’FFWF–Âç6†VWG2Âv†W&RÔö&¦V7B²·7G&–æuÒEòç6†VWD¶W’ÖWG6fU6†VWD¶W’ÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚G6†VWBä6÷VçBÖW’²F‡&÷r~hÈ~Zé®8~8þš^yºî8þjùN‹È>Zûî‹8¾Y
¾8î8(Î8n8N8î8¾8)>8"rÐ¢GvRÒ„vWBÔ'&’G6†VWE³ÒçvW2Âv†W&RÔö&¦V7B²„vWBÔ–çDFF&÷W'G’EòwvTçVÖ&W"r’ÖWGvTçVÖ&W"ÒÂ6VÆV7BÔö&¦V7BÔf—'7B¢–b‚GvRä6÷VçBÖW’²F‡&÷r~hÈ~Zé®8~8þ89®8;Î8+Ž8þjùN‹È>Zûî‹8¾Y
¾8î8(Î8n8N8î8¾8)>8"rÐ¢G&÷W'G’Ò7v—F6‚‚D76WB’°¢v&Vf÷&Rr²v&Vf÷&T76WBrÐ¢vgFW"r²vgFW$76WBrÐ¢v&Vf÷&RÖÖ6²r²v&Vf÷&TÖ6´76WBrÐ¢v&Vf÷&RÖ÷fW&Æ’r²v&Vf÷&T÷fW&Æ”76WBrÐ¢vÖ6²r²vÖ6´76WBrÐ¢v÷fW&Æ’r²v÷fW&Æ”76WBrÐ¢Ð¢Ff–ÆTæÖRÒ·7G&–æuÒ„vWBÔFF&÷W'G’GvU³ÒG&÷W'G’rr¢–b…·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Ff–ÆTæÖR’’²F‡&÷r~[zîXˆnyK¾X8þ8Î8.8(®8î8¾8)>8"rÐ ¢FgVÆÂÒrp¢G&ö÷BÒrp¢–b‚D76WBÖ–â‚v&Vf÷&RrÂvgFW"r’ÖæBFf–ÆTæÖRÖWw&VæFW"çærr’°¢28:Î8;>888:®8;>8+i˜.8¾KùÞZÙŽ8~8ó#E’ä~8).y»Nhê^‹ùN8~8jùN‹È>8*Þ8:>88>8+~8:^8Ž8î˜xÞŠH~8+>89N8;Î8).yÈ8þ8 ¢G6†VWDæÖRÒB†–b‚D76WBÖWv&Vf÷&Rr’²·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWE³Òv&Vf÷&U6†VWDæÖRr„vWBÔFF&÷W'G’G6†VWE³Òw6†VWDæÖRrrr’’ÒVÇ6R²·7G&–æuÒ„vWBÔFF&÷W'G’G6†VWE³ÒvgFW%6†VWDæÖRr„vWBÔFF&÷W'G’G6†VWE³Òw6†VWDæÖRrrr’’Ò¢–b‚D76WBÖWv&Vf÷&Rr’°¢G&7FW$F—"ÒvWBÕ&VæFW%&7FW%6†VWDF—"DÆæwVvR…·7G&–æuÒFF–fd6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒFF–fd6öçFW‡Bæ&6VÆ–æU6æ6†÷D–B’…·7G&–æuÒFF–fd6öçFW‡Bæ&6VÆ–æUfW'6–öä–B’G6†VWDæÖP¢ÒVÇ6R°¢G&7FW$F—"ÒvWBÕ&VæFW%&7FW%6†VWDF—"DÆæwVvR…·7G&–æuÒFF–fd6öçFW‡Bçv÷&¶&öö´–B’…·7G&–æuÒFF–fd6öçFW‡Bæ7W'&VçE6æ6†÷D–B’…·7G&–æuÒFF–fd6öçFW‡Bæ7W'&VçEfW'6–öä–B’G6†VWDæÖP¢Ð¢G&ö÷BÒ´”òåF…Ó£¤vWDgVÆÅF‚‚G&7FW$F—"¢FgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚„¦ö–âÕF‚G&ö÷B‚wvR×³£ÒçærrÖbGvTçVÖ&W"’’¢ÒVÇ6R°¢–b‚Ff–ÆTæÖRÖæ÷FÖF6‚uå´Õ¦×£Ó’åòÕÒµÂçærBrÖ÷"Ff–ÆTæÖRÖWw&VæFW"çærr’²F‡&÷r~[zîXˆnyK¾X8þ8Î8.8(®8î8¾8)>8"rÐ¢G&ö÷BÒ´”òåF…Ó£¤vWDgVÆÅF‚‚F66†TF—"¢FgVÆÂÒ´”òåF…Ó£¤vWDgVÆÅF‚‚„¦ö–âÕF‚F66†TF—"„¦ö–âÕF‚wr„¦ö–âÕF‚G6fU6†VWD¶W’Ff–ÆTæÖR’’’¢Ð¢–b‚Öæ÷BG&ö÷BäVæG5v—F‚…´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"’’²G&ö÷B³Ò´”òåF…Ó£¤F—&V7F÷'•6W&F÷$6†"Ð¢–b‚Öæ÷BFgVÆÂå7F'G5v—F‚‚G&ö÷BÂµ7G&–æt6ö×&—6öåÓ£¤÷&F–æÄ–væ÷&T66R’Ö÷"Öæ÷B…FW7BÕF‚ÔÆ—FW&ÅF‚FgVÆÂ’’°¢F‡&÷r~[zîXˆnyK¾X8þ8ÎŠh¾8N8¾8(®8î8¾8)>8"p¢Ð¢2U$Î8¾8þjùN‹È>XX>8;¾jùN‹È>XXŽ8;¾8*.8:¾8+N8:®8+®8:x˜Ž8ÎY
¾8î8(Î8(¾8þ8(8XX>8:ž8+ž8+þ8).y»Nhê^‹ùN8žZNYŽ8(&–Ö×WF&Æ^8~8(Ž8N8 ¢w&—FRÔ'—FW5&W7öç6RD6öçFW‡B#…´”òäf–ÆUÓ£¥&VDÆÄ'—FW2‚FgVÆÂ’’v–ÖvR÷ærrFfÇ6Rw&—fFRÂÖ‚ÖvSÓ3S3cÂ–Ö×WF&ÆRp§Ð ¦gVæ7F–öâvWBÕv÷&¶&öö´6†ævU7VÖÖ'’…·7G&–æuÒDÆæwVvRÂ·7G&–æuÒEv÷&¶&öö´–BÂEv÷&¶&öö²ÒFçVÆÂ’°¢F6×ÒvWBÔÆFW7D6ö×&—6öâDÆæwVvREv÷&¶&öö´–BEv÷&¶&öö°¢–b‚FçVÆÂÖWF6×’²&WGW&âFçVÆÂÐ¢&WGW&â¶÷&FW&VEÔ°¢7FGW2Ò·7G&–æuÒ„vWBÔFF&÷W'G’F6×w7FGW2rrr¢ÖWF†öBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6×vÖWF†öBrrr¢6†ævVE6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6×v6†ævVE6†VWG2r‚’’¢Væ6†ævVE6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6×wVæ6†ævVE6†VWG2r‚’’¢Væ¶æ÷vå6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6×wVæ¶æ÷vå6†VWG2r‚’’¢2cRÕ‚3“¢‹ûÞXª8;¾X˜®™šN8+~8;Î88Ž8("7FFR8¾‹Èž8¾8(¾8.jùN‹È>{YiéÎ8þhÈ8>8n8N8(¾8î8¾8>8>8~hÚŽ8n8n8N8þ8 ¢FFVE6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6×vFFVE6†VWG2r‚’’¢&VÖ÷fVE6†VWG2Ò„vWBÔ'&’„vWBÔFF&÷W'G’F6×w&VÖ÷fVE6†VWG2r‚’’¢ÖW76vRÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6×vÖW76vRrrr¢6ö×&VDBÒ·7G&–æuÒ„vWBÔFF&÷W'G’F6×v6ö×&VDBrrr¢Ð§Ð ¦gVæ7F–öâ–çfö¶RÕ7F'GW&V6÷fW'’…·7G&–æuÒDÆæwVvR’°¢2ÆV6^8;¾Kˆi˜.8+>89N8;Î8;¾ˆz®X¹^x«nhX¾8þy»N[èÎ8¾‹[~X¹^8ž8(¾8+ž8+8+Ž8:^8;Î8:ž8;ÎZÙ89~8:Þ8+¾8+ž8À¢2[êžiz~8ž8(¾8þ8(8T’8+^8;Î898;ÎXN8~8þK¨Î˜xÞ8¾X[iÈž88ž8:ž8*N89n8).‹[iû¾8~8®8N8 ¢G'’²&V6÷fW"Ôf–æÅG&ç67F–öç2DÆæwVvRÒ6F6‚²Ð§Ð ¢2cRÕ¢8+ž8+8+Ž8:^8;Î8:ž8;Î8:.8;Î88ž8.8>8(Î8ÎxJ8N8ŽZÙ89~8:Þ8+¾8+ž8Î˜	®[‹Ž8+^8;Î898;Î8Ž8~8n‹[~X¹^8~8¢28^8(ž8¾ZÚ¾8+ž8+8+Ž8:^8;Î8:ž8;Î8).‹[~X¹^8~8nxJ™™8¾Z)~jén8ž8(¾8.˜	®[‹Ž8+^8;Î898;ÎX‰ÞiÉþXÉn8(Ž8(®X˜Þ8¾{Úî8þ8>8Ž8 ¦–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚DWFõ66†VGVÆW%F‚’’°¢–çfö¶RÔWFõ66†VGVÆW$g&öÔf–ÆRÔ6öçG&öÅF‚DWFõ66†VGVÆW%F‚Õ&VçE&ö6W74–BE&VçE&ö6W74–@¢&WGW&à§Ð ¦–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚DF–fd¦ö%F‚’’°¢–çfö¶RÔF–fdFWF–Ä¦ö$g&öÔf–ÆRDF–fd¦ö%F€¢&WGW&à§Ð ¦–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚E&VæFW$¦ö%F‚’’°¢–çfö¶RÕ&VæFW$¦ö$g&öÔf–ÆRE&VæFW$¦ö%F€¢&WGW&à§Ð ¦–b‚Öæ÷B·7G&–æuÓ£¤—4çVÆÄ÷%v†—FU76R‚Df–æÄ¦ö%F‚’’°¢–çfö¶RÔf–æÄ'V–ÆD¦ö$g&öÔf–ÆRDf–æÄ¦ö%F€¢&WGW&à§Ð ¦–b‚E÷'BÖÆR’²E÷'BÒvWBÔg&VU÷'BÐ¢F6öæf–sÒvWBÔ6öæf–p¢F6öæf–sÒ–æ—F–Æ—¦RÔÆö6Å&ö¦V7D6öæf–rF6öæf–s §G'’²G7F'GWF‡3ÔvWBÕF‡3²–b‚G7F'GWF‡2æFFF—"ÖæB…FW7BÕF‚ÔÆ—FW&ÅF‚…·7G&–æuÒG7F'GWF‡2æFFF—"’’—´Vç7W&RÕ6¶vRG7F'GWF‡2ÔÆæwVvW2‚„vWBÔVffV7F—fTÆæwVvR’—ÒÒ6F6‚²w&—FRÕv&æ–ærEòäW†6WF–öâäÖW76vRÐ ¢2iÊ®ZèÎK¨n8îhùX{®yJ…Dn88Ž8:ž8;>8+n8*þ8+~8:~8;>888þ8Tži8ÞKÙÎ8).Xù~8K¹Ž88(¾X˜Þ8¾[êžiz~8ž8(¾8 §G'’²–çfö¶RÕ7F'GW&V6÷fW'’„vWBÔVffV7F—fTÆæwVvR’Ò6F6‚²w&—FRÕv&æ–ærEòäW†6WF–öâäÖW76vRÐ ¢G&Vf—‚Ò&‡GG¢òó#rããã¢E÷'Bò ¢GW&ÂÒ‚&‡GG¢òó#rããã§³Òó÷Fö¶Vã×³Ò"ÖbE÷'BÂE67&—C¥Fö¶Vâ ¢2W6R6ÖÆÂF7Æ—7FVæW"Ö&6VB…EE6W'fW"–ç7FVBöb‡GGÆ—7FVæW"à¢2F†—2fö–G2U$Â4ÂòFÖ–æ—7G&F÷"×&–v‡G2—77VW2öâÆö6¶VBÖF÷vâv–æF÷w272à¥7F'BÔÆö6ÅF76W'fW"E÷'BGW&Â…¶&ööÅÒDæô÷Vâ