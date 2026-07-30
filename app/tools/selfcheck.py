from pathlib import Path
import json, re, zipfile

root = Path(__file__).resolve().parents[2]
required = [
    '日本語管理.vbs','英語管理.vbs','README.md','THIRD_PARTY_NOTICES.md','app/server.ps1','app/default-config.json','app/launch.ps1',
    'app/web/index.html','app/web/style.css','app/web/app.js','app/lib/pdfbox/ReportPdfComposer.jar',
    'app/lib/pdfbox/src/ReportPdfComposer.java','app/lib/pdfbox/src/BatchPdfSplitter.java','app/lib/pdfbox/build.ps1',
    'app/tools/install-thirdparty.ps1','app/tools/install-thirdparty.cmd','app/tools/verify-thirdparty.ps1','app/tools/select-folder.ps1','app/tools/package-release.ps1',
    'app/tools/diff-image-pages.ps1','app/tools/DiffImageEngine.cs',
    'app/tools/fixtures/structure-v1-ja.json','app/tools/fixtures/structure-mixed-order.json','docs/API.md','docs/THIRD_PARTY_SETUP.md','docs/ReportBinder_UIUX改修指示書_V4.md'
]
missing=[x for x in required if not (root/x).exists()]
if missing: raise SystemExit('missing: '+', '.join(missing))
if (root/'app/config.json').exists(): raise SystemExit('runtime app/config.json must not be distributed')
json.loads((root/'app/default-config.json').read_text(encoding='utf-8'))
for ps1 in ['app/server.ps1','app/launch.ps1','app/lib/pdfbox/build.ps1','app/tools/install-thirdparty.ps1','app/tools/verify-thirdparty.ps1','app/tools/select-folder.ps1','app/tools/package-release.ps1','app/tools/diff-image-pages.ps1']:
    if not (root/ps1).read_bytes().startswith(b'\xef\xbb\xbf'): raise SystemExit(f'PowerShell must be UTF-8 BOM: {ps1}')

server=(root/'app/server.ps1').read_text(encoding='utf-8-sig')


for needle in [
    '[Security.Cryptography.RandomNumberGenerator]::Create()',
    'function Test-FixedTimeTokenEquals',
    '[Security.Cryptography.SHA256]::Create()',
    '$difference = $difference -bor ($candidateHash[$i] -bxor $expectedHash[$i])'
]:
    if needle not in server: raise SystemExit(f'token hardening missing: {needle}')
if 'Get-Random -Minimum 0 -Maximum 256' in server:
    raise SystemExit('server token fallback must use a cryptographic RNG')
test_token_block=server.split('function Test-Token($Request)',1)[1].split('\nfunction ',1)[0]
if '-eq $Script:Token' in test_token_block:
    raise SystemExit('token comparison must use fixed-time byte comparison')
diff_skeleton_block=server.split('function New-DiffDetailSkeleton',1)[1].split('function Get-DiffDetail',1)[0]
if 'if ([bool]$Context.available)' not in diff_skeleton_block:
    raise SystemExit('unavailable diff detail must not resolve empty history ids')
compare_block=server.split('function Compare-SnapshotVisual',1)[1].split('function Invoke-PostRenderAnalysis',1)[0]
for needle in ["previousSnapshotId", "Get-RenderVersionIds $Language $WorkbookId $previousSnapshotId", "Sort-Object -Descending"]:
    if needle not in compare_block: raise SystemExit(f'comparison baseline recovery missing: {needle}')
for needle in ['$comparisonKey = (Get-Sha256Text', '"cmp-{0}.json"']:
    if needle not in compare_block: raise SystemExit(f'comparison path shortening missing: {needle}')
post_analysis_block=server.split('function Invoke-PostRenderAnalysis',1)[1].split('function Get-LatestComparison',1)[0]
if "'comparison-analysis.json'" not in post_analysis_block:
    raise SystemExit('post-render comparison analysis must remain auditable')
stored_comparison_block=server.split('function Get-StoredComparison',1)[1].split('\nfunction ',1)[0]
for field in ['scope','baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId']:
    if f"Get-DataProperty $candidate '{field}'" not in stored_comparison_block:
        raise SystemExit(f'stored comparison identity must include {field}')
latest_comparison_block=server.split('function Get-LatestComparison([string]',1)[1].split('\nfunction ',1)[0]
if "'automatic'" not in latest_comparison_block or "'scope'" not in latest_comparison_block:
    raise SystemExit('historical comparisons must not replace the automatic latest comparison')
for needed in ["currentSnapshotId", "currentVersionId", "$candidateSnapshot -ne $snap", "$candidateVersion -ne $ver"]:
    if needed not in latest_comparison_block:
        raise SystemExit(f'latest automatic comparison identity check missing: {needed}')

# Historical comparison GET paths are read-only; persistence is explicit and verified.
_history_calc=server.split('function New-HistoricalSnapshotComparison',1)[1].split('\nfunction ',1)[0]
for forbidden in ['Write-JsonFile','Write-HistoryEvent']:
    if forbidden in _history_calc: raise SystemExit(f'historical comparison calculation must be pure: {forbidden}')
_history_save=server.split('function Save-HistoricalSnapshotComparison',1)[1].split('\nfunction ',1)[0]
for needed in ['Read-JsonFile $path', "'scope'", "'baselineVersionId'", "'currentVersionId'", "'compare.history.created'"]:
    if needed not in _history_save: raise SystemExit(f'historical comparison save verification missing: {needed}')
if _history_save.index('Read-JsonFile $path') > _history_save.index("'compare.history.created'"):
    raise SystemExit('history event must be written only after comparison save verification')
_auto_compare=server.split('function Compare-SnapshotVisual',1)[1].split('\nfunction ',1)[0]
for needed in ['Get-HistoryRenderVersionAvailability', 'Render-SnapshotForComparison', 'Read-JsonFile $comparisonPath', "'currentVersionId'", "'compare.completed'"]:
    if needed not in _auto_compare: raise SystemExit(f'automatic comparison integrity missing: {needed}')
if _auto_compare.index('Read-JsonFile $comparisonPath') > _auto_compare.index("'compare.completed'"):
    raise SystemExit('automatic comparison event must follow save verification')
if 'Set-LatestComparisonAssets' not in _auto_compare or _auto_compare.index('Set-LatestComparisonAssets') > _auto_compare.index("'compare.completed'"):
    raise SystemExit('automatic comparison assets must be pinned before publishing the completion event')
if 'Remove-Item -LiteralPath $comparisonPath' not in _auto_compare:
    raise SystemExit('unprotected automatic comparison metadata must be discarded')
_latest_assets=server.split('function Set-LatestComparisonAssets',1)[1].split('\nfunction ',1)[0]
for needed in ['latest-comparison-baseline','latest-comparison-current','New-SnapshotPin','New-ContentPdfPin','Get-HistoryRenderVersionAvailability','Read-JsonFile $pointerPath']:
    if needed not in _latest_assets: raise SystemExit(f'latest automatic comparison asset protection missing: {needed}')

# Hashes and content PDFs must come from the exact same render version.
_resolve_diff=server.split('function Resolve-DiffContentPdfPath',1)[1].split('\nfunction ',1)[0]
if 'Resolve-ContentPdfSheetPathExact' not in _resolve_diff:
    raise SystemExit('diff content PDF lookup must use exact render-version resolution')
for forbidden in ['contentPdfRetained','candidateVersion','foreach ($ver in $versions)']:
    if forbidden in _resolve_diff: raise SystemExit(f'diff content PDF lookup must not fall back: {forbidden}')
_availability=server.split('function Get-HistoryRenderVersionAvailability',1)[1].split('\nfunction ',1)[0]
for needed in ['visualHashAvailable','contentPdfAvailable','missingSheets','Resolve-ContentPdfSheetPathExact']:
    if needed not in _availability: raise SystemExit(f'history visual-compare readiness missing: {needed}')

# Pair generation is cross-PC serialized and both inputs are leased until finally cleanup.
for fn in ['function Get-DiffLaunchLockPath','function Get-DiffGenerationLockPath','function New-DiffJobLeases','function Refresh-DiffJobLeases','function Remove-DiffJobLeases','function Get-ContentPdfMaintenanceLockPath']:
    if fn not in server: raise SystemExit('missing diff concurrency protection: '+fn)
_start_diff=server.split('function Start-DiffDetailJob',1)[1].split('\nfunction ',1)[0]
for needed in ['Get-DiffLaunchLockPath','New-DiffJobLeases','generationLockPath','Save-HistoricalSnapshotComparison']:
    if needed not in _start_diff: raise SystemExit(f'diff launch protection missing: {needed}')
_worker_wrap=server.split('function Invoke-DiffDetailJobFromFile',1)[1].split('\nfunction ',1)[0]
if 'finally {' not in _worker_wrap or 'Remove-DiffJobLeases $job' not in _worker_wrap:
    raise SystemExit('diff worker must release leases on every exit path')
_baseline=server.split('function Set-ComparisonBaseline',1)[1].split('\n# ---- 比較専用レンダリング',1)[0]
for needed in ['New-ContentPdfPin','Remove-ContentPdfPin','Get-HistoryRenderVersionAvailability','Get-ContentPdfMaintenanceLockPath']:
    if needed not in _baseline: raise SystemExit(f'automatic comparison baseline protection missing: {needed}')

# Unchanged sheets are lazy, and concurrent lazy requests are retried after the active pair job.
_skeleton=server.split('function New-DiffDetailSkeleton',1)[1].split('\nfunction Get-DiffDetail',1)[0]
if "'deferred'" not in _skeleton: raise SystemExit('unchanged sheets must be deferred')
_worker=server.split('function Invoke-DiffDetailJobCore',1)[1].split('\nfunction ',1)[0]
if "$sheetKind -ne 'unchanged'" not in _worker or "$sheetStatus -eq 'failed'" not in _worker or 'requestedSheetKey' not in _worker:
    raise SystemExit('initial diff job must defer unchanged sheets but retry failed lazy sheets')
_appjs_early=(root/'app/web/app.js').read_text(encoding='utf-8-sig')
for needed in ['for(let attempt=0;attempt<3;attempt++)', 'joinedExistingDiffJob', "!['deferred','failed'].includes(targetStatus)", "targetStatus==='failed'&&!joinedExisting"]:
    if needed not in _appjs_early: raise SystemExit(f'lazy diff retry missing: {needed}')
for needle in ['function New-HistoricalSnapshotComparison',
               "scope = 'history'",
               'function Get-PreferredHistoryRenderVersion',
               'fromSnapshotId',
               'toSnapshotId']:
    if needle not in server: raise SystemExit(f'historical visual comparison missing: {needle}')
for needle in ['[string]$Context.scope', "Join-Path 'comparisons' (\"d{0}\"", "return ('s-' + $hash.Substring(7, 10))", "Join-Path 'p'"]:
    if needle not in server: raise SystemExit(f'diff cache path shortening missing: {needle}')
if "Join-Path 'diff-pages'" in server:
    raise SystemExit('long diff-pages cache path remains')
write_json_block=server.split('function Write-JsonFile',1)[1].split('function Read-TextFileTailSafe',1)[0]
try_pos=write_json_block.find('try {')
tmp_write_pos=write_json_block.find('Write-Utf8NoBomFile $tmp $json')
if try_pos < 0 or tmp_write_pos < try_pos:
    raise SystemExit('atomic JSON temp write must be inside fallback try block')
for route in ['/api/state','/api/paths','/api/workbooks/register-batch','/api/workbooks/render/start','/api/jobs/status','/api/pages/reorder','/api/pages/sort-by-sheet','/api/final/readiness','/api/final/build','/api/final/file','/api/scan-updates','/api/history/diff-detail','/api/history/diff/prepare','/api/history/diff-page']:
    if route not in server: raise SystemExit(f'route not found: {route}')
for needle in [
    'Update-StructureLocked','Read-StructureUnlocked','Write-StructureUnlocked','Initialize-Or-MigrateStructure','structure.json.v1.bak',
    'Get-VolumeStateKey','builtFingerprint','Get-FinalBuildInputSnapshot','contentPdfLastWriteUtcTicks','contentPdfSize',
    'manifest_${Volume}_${cat}.json','~building_${Volume}_${cat}.pdf','Require-WorkbookCategory','System.ArgumentException',
    'Mark-VolumeNeedsRebuild','staleReasons','Sort-PagesBySheet','Insert-PageInSheetOrder','Renumber-VolumeOrder',
    'CenterHorizontally = $true','LeftMargin = Convert-CmToPt 1.2','RightMargin = Convert-CmToPt 1.2','punchShiftPt=(Convert-CmToPt 0.2)',
    'lib\\java\\bin\\java.exe','Apply-DefaultNumberingPerVolume','first-page-none','ExcelPrintProfileVersion'
]:
    if needle not in server: raise SystemExit(f'server feature not found: {needle}')
if len(re.findall(r'(?<!function )Save-Structure\s',server)):
    raise SystemExit('direct Save-Structure use is forbidden')
# First-run setup initializes the selected dataDir before local config is committed.
for signature in [
    "function Get-WorkspacePath([string]$Language, [string]$DataDir = '')",
    "function Read-StructureUnlocked([string]$Language, [string]$DataDir = '')",
    "function Write-StructureUnlocked([string]$Language, $Structure, [string]$DataDir = '')",
    "function Initialize-Or-MigrateStructure([string]$Language, [string]$DataDir = '')",
    "Initialize-Or-MigrateStructure $lang ([string]$Paths.dataDir)"
]:
    if signature not in server: raise SystemExit(f'first-run workspace initialization regression: {signature}')
# Migration must be isolated from read-only Get-Structure.
get_block=server.split('function Get-Structure',1)[1].split('function Update-StructureLocked',1)[0]
if 'Write-StructureUnlocked' in get_block or re.search(r'(?m)^\s*Initialize-Or-MigrateStructure\b', get_block): raise SystemExit('Get-Structure must stay read-only')
# Fingerprint must represent rendered inputs, not the source Excel hash.
fp_block=server.split('function Get-FinalBuildInputSnapshot',1)[1].split('function Get-FinalBuildFingerprint',1)[0]
if 'currentExcelHash' in fp_block: raise SystemExit('currentExcelHash must not be in final build fingerprint')
for needle in ['contentPdfSize','contentPdfLastWriteUtcTicks','lastRenderedVersionId']:
    if needle not in fp_block: raise SystemExit(f'fingerprint input missing: {needle}')
# Category must fail closed on all final paths.
for fn in ['Build-FinalPdf','Serve-FinalPdfByVolume','Get-FinalBuildInputSnapshot']:
    block=server.split(f'function {fn}',1)[1].split('\nfunction ',1)[0]
    if 'Require-WorkbookCategory' not in block: raise SystemExit(f'category not required: {fn}')
# Long-running Excel/Java work must be outside the structure transaction body.
build=server.split('function Build-FinalPdf',1)[1].split('function Get-StatePayload',1)[0]
java_pos=build.find('ReportPdfComposer'); commit_pos=build.find('$commit=Update-StructureLocked')
if java_pos < 0 or commit_pos < java_pos: raise SystemExit('final composer/commit order is invalid')

appjs=(root/'app/web/app.js').read_text(encoding='utf-8-sig')
for needle in ['renderGlobalHeader','renderStepBar','renderNavBadges','aggregateFinalState','volumeReadiness','isEditing','lastPageBoardRenderSignature','sortPagesBySheet','/api/pages/sort-by-sheet','category:activePreset','openFinalVolume(volume, category=activePreset)','notice-actions','insertedAtEndCount','modalReturnFocus','render-all-btn','openDiffDetail','moveDiffRegion','syncDiffScroll','/api/history/diff/prepare']:
    if needle not in appjs: raise SystemExit(f'ui feature not found: {needle}')
for dead in ["bind('refresh-btn'","bind('load-files-btn'","bind('save-paths-btn'","bind('render-updated-btn'","bind('render-selected-pages-btn'","$('category-heading')","$('category-caption')"]:
    if dead in appjs: raise SystemExit(f'dead ui code remains: {dead}')
if "body:{volumes:collectBoardVolumes()}" in appjs: raise SystemExit('page reorder must send category')

html=(root/'app/web/index.html').read_text(encoding='utf-8')
for needle in ['workspace-top','step-bar','data-preset="ecm"','global-final-status','nav-excel-count','sort-by-sheet-btn','build-all-btn','render-all-btn','unregister-selected-btn','final-main-fix','final-appendix-fix','notice-actions','language-popover','<symbol id="i-home"','id="diff-modal"','id="diff-before-viewport"','id="diff-after-viewport"','id="diff-mode-overlay"']:
    if needle not in html: raise SystemExit(f'html feature not found: {needle}')
for dead in ['collapse-hint','pagination-lite','info-dot','menu-col','recent-folder-list','page-output-summary','category-card','workflow-card']:
    if dead in html: raise SystemExit(f'dead ui remains: {dead}')
for symbol in '⌂□▦▤▣△⋮◉⌄↻⌕⊘›':
    if symbol in html: raise SystemExit(f'font symbol remains in html: {symbol}')

css=(root/'app/web/style.css').read_text(encoding='utf-8')
for needle in ['--accent:#5e6ad2','--sidebar-w:240px','backdrop-filter','box-shadow:none','.badge.attention','background:var(--attention-subtle)','font-weight:600','.modal-card','.drop-placeholder','.notice{position:fixed','.diff-dialog{width:min(1800px,94vw)','.diff-viewers.overlay-mode','@media(max-width:1100px)']:
    if needle not in css: raise SystemExit(f'css feature not found: {needle}')
if 'font-weight:900' in css or 'radial-gradient' in css: raise SystemExit('old visual style remains')

# V4.5 Excel screen width and Explorer-matching timestamp checks.
for needle in ['availableFilesScannedAt','formatDateTimeWithSeconds','formatSubmissionFileModifiedAt','modifiedAtDisplay','setFileSelectionSummary','title="Excelファイルの最終保存日時"']:
    if needle not in appjs: raise SystemExit(f'v4.5 Excel UI feature not found: {needle}')
if '<th class="numeric-col">シート数</th>' in appjs or '<th class="numeric-col">ページ数</th>' in appjs:
    raise SystemExit('Excel tables must not show sheet/page count columns')
if 'formatDateTimeWithSeconds(f.modifiedAt' in appjs:
    raise SystemExit('submission file modified time must display to the minute')

if 'const byEpoch = formatLocalDateTimeMinuteFromUnixMs(file?.modifiedAtUnixMs)' in appjs:
    raise SystemExit('submission file time must not be converted again in the browser')
if html.find('id="select-render-needed-btn"') > html.find('id="render-selected-btn"'):
    raise SystemExit('PDF render button must be next to the needed-selection action')
for needle in ['unregistered-card','registered-card','workbook-toolbar','render-target-hint']:
    if needle not in html: raise SystemExit(f'v4.3 Excel HTML feature not found: {needle}')
for needle in ['font-family:"BIZ UDPGothic"','grid-template-columns:minmax(460px,.82fr) minmax(650px,1.18fr)','font-size:15px','.workbook-table{width:100%;min-width:0']:
    if needle not in css: raise SystemExit(f'v4.3 readability CSS feature not found: {needle}')
for needle in ['.workbook-table .pdf-status-col{width:320px}', '.excel-grid{grid-template-columns:minmax(340px,.65fr) minmax(720px,1.35fr)', '.excel-grid>.card+.card{margin-top:0}', '.excel-grid .file-name-cell strong{overflow:visible;text-overflow:clip;white-space:normal;overflow-wrap:anywhere', '.pdf-status-stack{display:grid;gap:5px}', '@media(max-width:1320px){.excel-grid{grid-template-columns:1fr}']:
    if needle not in css: raise SystemExit(f'change column readability CSS missing: {needle}')
for needle in ['<div class="file-meta"', '更新日時：', 'PDF作成日時：', '<th class="pdf-status-col">PDF状況</th>', '差分は作成後に確認します', '前回PDFとの差分']:
    if needle not in appjs: raise SystemExit(f'file identity UX feature missing: {needle}')
if '<th class="date-col" title="Excelファイルの最終保存日時">更新日時</th>' in appjs or '<th class="date-col" title="最後にページPDFを作成した日時">PDF作成日時</th>' in appjs:
    raise SystemExit('Excel timestamps must be secondary metadata under the file name')
if '<th class="change-col">変更</th>' in appjs or '<th class="status-col">状態</th>' in appjs:
    raise SystemExit('change and state must be merged into one PDF status column')
for needle in ['$f.Refresh()','ReportBinderNative.FileTimes','Get-FileLastWriteSnapshot','modifiedAtDisplay','scannedAt = (New-NowIso)']:
    if needle not in server: raise SystemExit(f'v4.5 file freshness feature not found: {needle}')

# Fixtures document the real v1 shape and order normalization contract.
v1=json.loads((root/'app/tools/fixtures/structure-v1-ja.json').read_text(encoding='utf-8'))
if set(v1['volumes']) != {'ja-main','ja-appendix'}: raise SystemExit('v1 fixture must use language-specific old two-key volumes')
expected={f'ja-{kind}|{cat}' for kind in ('main','appendix') for cat in ('ecm','bod','dmm')}
if len(expected)!=6: raise SystemExit('six-key migration model invalid')
order=json.loads((root/'app/tools/fixtures/structure-mixed-order.json').read_text(encoding='utf-8'))
renumbered=[(i+1)*10 for i,_ in enumerate(sorted(order['orders']))]
if renumbered != order['expected']: raise SystemExit('10-step order fixture invalid')

logs=[x for x in (root/'app/logs').glob('*') if x.name!='.gitkeep'] if (root/'app/logs').exists() else []
if logs: raise SystemExit('runtime logs must not be distributed')
cache=[x for x in (root/'app/thirdparty-cache').glob('*') if x.is_file()] if (root/'app/thirdparty-cache').exists() else []
if cache: raise SystemExit('download cache must not be distributed')
if (root/'app/lib/pdfbox/classes').exists():
    raise SystemExit('compiled class staging directory must not be distributed')
python_cache=[x for x in root.rglob('*') if x.name == '__pycache__' or x.suffix == '.pyc']
if python_cache:
    raise SystemExit('Python cache must not be distributed: ' + ', '.join(str(x.relative_to(root)) for x in python_cache))

with zipfile.ZipFile(root/'app/lib/pdfbox/ReportPdfComposer.jar') as zf:
    for cls in ['ReportPdfComposer.class','BatchPdfSplitter.class','BatchPdfSplitter$PdfBox.class']:
        if cls not in set(zf.namelist()): raise SystemExit(f'jar class missing: {cls}')

root_files={x.name for x in root.iterdir() if x.is_file()}
extra=sorted(root_files-{'日本語管理.vbs','英語管理.vbs','README.md','THIRD_PARTY_NOTICES.md','CHANGELOG_V4.md','CHANGELOG_V5.md','.gitignore'})
if extra: raise SystemExit('unexpected top files: '+', '.join(extra))

# PowerShell automatic variable $PID is read-only and variable names are case-insensitive.
# Assigning to a local variable named $pid therefore breaks PDF rendering at runtime.
if re.search(r'(?im)\$pid\s*=', server):
    raise SystemExit('reserved PowerShell variable $PID is assigned')

# ---- V5 Stage 1: field ownership, hash format, IDs, locks ----

# Render-Workbook must not write currentExcel* (Scan-Updates owns those fields).
_rw = server.index('function Render-Workbook(')
_rw_end = server.find('function Set-WorkbookRenderError', _rw)
_rw_body = server[_rw:_rw_end if _rw_end > 0 else len(server)]
if re.search(r"Set-NoteProperty \$wb 'currentExcel", _rw_body):
    raise SystemExit('Render-Workbook must not write currentExcel* (Scan-Updates owns them)')
if re.search(r"foreach\(\$name in @\('currentExcelModifiedAt'", server):
    raise SystemExit('render commit must not copy currentExcel* into the locked structure')

# Failure status must be decided against the attempted version, on every catch path.
if 'AttemptedHash' not in server:
    raise SystemExit('Set-WorkbookRenderError must accept the attempted version')
_calls = re.findall(r'Set-WorkbookRenderError \$\w+ \$id \$msg \$errorDetail(.*)', server)
if len(_calls) != 3:
    raise SystemExit(f'expected 3 Set-WorkbookRenderError call sites, found {len(_calls)}')
for tail in _calls:
    if 'att.hash' not in tail:
        raise SystemExit('every Set-WorkbookRenderError call must pass the attempted version')

# File hashes and string fingerprints use different formats and must not be compared directly.
if 'function Normalize-FileHash' not in server:
    raise SystemExit('Normalize-FileHash is required (New-Sha256 is uppercase without prefix)')

# IDs must not be second-granular.
if re.search(r"\$versionId\s*=\s*'v'\s*\+\s*\(Get-Date\)", server):
    raise SystemExit('versionId must be millisecond+GUID8, not second-granular')
for fn in ['function New-RbId', 'function New-RbVersionId']:
    if fn not in server: raise SystemExit('missing: ' + fn)

# Excel COM is limited to one job per language; lock order must be engine -> workbook.
if 'function Invoke-WithRenderLock' not in server:
    raise SystemExit('Invoke-WithRenderLock is required (render-engine -> render_<workbookId>)')
if 'locks\\workbook_$WorkbookId.lock' in server:
    raise SystemExit('legacy workbook lock name remains')
_render_lock = server.split('function Invoke-WithRenderLock', 1)[1].split('\nfunction ', 1)[0]
for needed in ['Invoke-WithLock $enginePath', 'Invoke-WithLock $bookPath $Body']:
    if needed not in _render_lock:
        raise SystemExit('render lock order regression: ' + needed)
_generic_lock = server.split('function Invoke-WithLock', 1)[1].split('\nfunction ', 1)[0]
if '[scriptblock]$Action' not in _generic_lock or 'return & $Action' not in _generic_lock:
    raise SystemExit('generic lock action must not shadow Invoke-WithRenderLock $Body')
if '[scriptblock]$Body' in _generic_lock:
    raise SystemExit('generic lock must not use the dynamically shadowing $Body parameter name')
for fn in ['function Try-AcquireLockHandle', 'function Release-LockHandle']:
    if fn not in server: raise SystemExit('missing: ' + fn)
_insert = server.split('function Insert-PageInSheetOrder', 1)[1].split('\nfunction ', 1)[0]
if 'pages=@($ordered)' in _insert or 'pages=$ordered.ToArray()' not in _insert:
    raise SystemExit('generic page list must use ToArray() for Windows PowerShell 5.1')

# Approval is a workspace-level policy; personal config must not gate it.
for fn in ['function Get-WorkspacePolicy', 'function Test-InputHistoryEnabled', 'function Test-SourceRetentionEnabled']:
    if fn not in server: raise SystemExit('missing: ' + fn)
if 'function Merge-ConfigDefaults' not in server:
    raise SystemExit('Get-AppConfig must merge default-config so existing users receive new keys')

_cfg = json.loads((root/'app/default-config.json').read_text(encoding='utf-8'))
for key in ('inputHistory', 'autoRender'):
    if key not in _cfg: raise SystemExit(f'default-config.json is missing {key}')
if _cfg['autoRender'].get('enabled') is not False:
    raise SystemExit('autoRender must ship disabled')

# Environment is captured per job, not once per server process.
if 'function Reset-RenderEnvironmentForJob' not in server:
    raise SystemExit('render environment must be re-captured per job')

# ---- V5 Stage 2-5: history, pipeline, visual hashes, transaction ----
for fn in ['function Ensure-SnapshotMetadata', 'function Capture-RenderInput', 'function Save-SnapshotSourceFile',
           'function Capture-DetectedSnapshot', 'function Invoke-InputHistoryCleanup',
           'function New-SnapshotPin', 'function New-SnapshotLease', 'function Clear-ExpiredLeases',
           'function Invoke-AutoSchedulerFromFile', 'function Invoke-AutoSchedulerTick',
           'function Test-InteractiveExcelInUse', 'function Recover-AutoStates',
           'function Invoke-PdfPageAnalyzer', 'function Compare-SnapshotVisual',
           'function Render-SnapshotForComparison', 'function Set-ComparisonBaseline',
           'function Save-LayoutSnapshot', 'function Restore-LayoutSnapshot', 'function Get-LayoutRestorePreview',
           'function New-FinalArchive', 'function Invoke-FinalBuildTransaction',
           'function Restore-FinalTransaction', 'function Recover-FinalTransactions']:
    if fn not in server: raise SystemExit('missing: ' + fn)

# Snapshot capture happens at detection, never after the quiet period.
if 'Update-InputHistoryAfterScan' not in server:
    raise SystemExit('scan must capture snapshots at detection time')

# The scheduler runs as a child process, not inside the single-threaded HTTP listener.
if "AutoSchedulerPath" not in server:
    raise SystemExit('server.ps1 must support -AutoSchedulerPath')
_handle = server.split('function Handle-Api', 1)[1].split('\nfunction ', 1)[0]
if re.search(r'Start-Sleep\s+-Seconds\s+([1-9]\d+)', _handle):
    raise SystemExit('Handle-Api must not block the listener for long periods')

# Restore must apply layout fields only.
_restore = server.split('function Restore-LayoutSnapshot', 1)[1].split('\nfunction ', 1)[0]
for forbidden in ["'contentPdf'", "'currentExcelHash'", "'lastRenderedVersionId'", "'lastRenderedExcelHash'"]:
    if f'Set-NoteProperty $p {forbidden}' in _restore:
        raise SystemExit('layout restore must not touch ' + forbidden)

# The journal phase must be written before its side effect.
_txn = server.split('function Invoke-FinalBuildTransaction', 1)[1].split('\nfunction ', 1)[0]
for phase in ['backups-created', 'replacing-files', 'structure-committing', 'archiving', 'completed']:
    if f"'{phase}'" not in _txn: raise SystemExit('journal phase missing: ' + phase)
if '$journal.oldVolumeStates' not in _txn or '$journal.newVolumeStates' not in _txn:
    raise SystemExit('journal must record old and new volume states for rollback')
# no-pages must not block a combined build.
if "-ne 'no-pages'" not in _txn:
    raise SystemExit('build-all must skip empty volumes instead of failing on no-pages')

# Recovery trusts file hashes, not progress flags.
_rec = server.split('function Restore-FinalTransaction', 1)[1].split('\nfunction ', 1)[0]
if 'newPdfHash' not in _rec or 'manual-recovery-required' not in _rec:
    raise SystemExit('recovery must compare real PDF hashes and stop on mismatch')

# The analyzer must exist and be built.
_analyzer = root / 'app/lib/pdfbox/src/PdfPageAnalyzer.java'
if not _analyzer.exists(): raise SystemExit('PdfPageAnalyzer.java missing')
_ajava = _analyzer.read_text(encoding='utf-8')
if 'ImageType.RGB' not in _ajava: raise SystemExit('visual hash must rasterize in RGB')
if 'toUpperCase' not in _ajava: raise SystemExit('pixel hashes must match New-Sha256 formatting')
if 'PdfPageAnalyzer.java' not in (root / 'app/lib/pdfbox/build.ps1').read_text(encoding='utf-8'):
    raise SystemExit('build.ps1 must compile PdfPageAnalyzer')
if "'-Djava.awt.headless=true'" not in server:
    raise SystemExit('PdfPageAnalyzer must run headless for scheduler/remote environments')

# ---- V5 P0: failures the previous selfcheck did not catch ----

# The scheduler child must be dispatched before the normal server starts,
# otherwise it boots another server and spawns its own child, recursively.
_tail = server[server.index('$prefix = "http://127.0.0.1:$Port/"') - 4000:]
if 'Invoke-AutoSchedulerFromFile -ControlPath $AutoSchedulerPath' not in server:
    raise SystemExit('missing -AutoSchedulerPath dispatch (child would start another server)')
_dispatch = server.index('Invoke-AutoSchedulerFromFile -ControlPath $AutoSchedulerPath')
_serverstart = server.index('Start-LocalTcpServer $Port')
if _dispatch > _serverstart:
    raise SystemExit('-AutoSchedulerPath dispatch must precede server start')

# NameValueCollection is enumerable. The custom TCP server must return the
# query-string collection as one object or token/category/jobId lookups fail.
_query = server.split('function New-QueryStringCollection', 1)[1].split('\nfunction ', 1)[0]
if _query.count('return ,$nvc') != 3:
    raise SystemExit('query-string parser must preserve NameValueCollection with unary-comma returns')

# API-supplied history identifiers become path segments and must fail closed.
if 'function Assert-SafeStorageSegment' not in server:
    raise SystemExit('history/storage identifiers must be validated before path construction')
for fn in ['function Get-WorkbookHistoryDir', 'function Get-SnapshotDir',
           'function Get-ContentPdfVersionDir', 'function Read-LayoutSnapshot']:
    blk = server.split(fn, 1)[1].split('\nfunction ', 1)[0]
    if 'Assert-SafeStorageSegment' not in blk:
        raise SystemExit(f'{fn} must validate its storage path segments')
for route in ['/api/history/pin', '/api/history/unpin', '/api/auto/run-now']:
    route_tail = server.split(f"$path -eq '{route}'", 1)[1].split('\n        if ', 1)[0]
    if 'Assert-SafeStorageSegment' not in route_tail:
        raise SystemExit(f'{route} must reject unsafe identifiers')

# The custom single-threaded HTTP listener must not accept unbounded bodies.
_tcp_read = server.split('function Read-TcpHttpContext', 1)[1].split('\nfunction ', 1)[0]
for needed in ['1048576', 'Invalid Content-Length header.', 'Incomplete HTTP request body.']:
    if needed not in _tcp_read:
        raise SystemExit(f'custom HTTP request limit regression: {needed}')

# The scheduler must detect changes itself; the browser timer is not the source of truth.
_tick = server.split('function Invoke-AutoSchedulerTick', 1)[1].split('\nfunction ', 1)[0]
for needed in ['Scan-Updates', 'Update-InputHistoryAfterScan']:
    if needed not in _tick: raise SystemExit(f'scheduler must call {needed}')

# The waited-on snapshot must be pinned into the job, not re-resolved at render time.
if 'snapshotPins' not in server: raise SystemExit('auto job must pin the waited-on snapshot')

# The hash must describe the file actually opened.
_rw = server.split('function Render-Workbook(', 1)[1].split('\nfunction ', 1)[0]
if _rw.index('Capture-RenderInput') > _rw.index('$sourceHash = New-StableHash'):
    raise SystemExit('render input must be resolved before hashing')

# Per-file deletion of old content-pdf breaks retained generations.
if 'Remove-ContentPdfFileSafe $workspace' in server:
    raise SystemExit('per-file content-pdf deletion must be removed')

# manifest.json is the completion marker, so it is written after the source copy.
_ensure = server.split('function Ensure-SnapshotMetadata', 1)[1].split('\nfunction ', 1)[0]
if "'manifest.json'" in _ensure: raise SystemExit('Ensure-SnapshotMetadata must write manifest.pending.json')
if 'function Complete-Snapshot' not in server: raise SystemExit('missing Complete-Snapshot')

# An unapproved workspace must behave exactly like V4.1.
_bfp = server.split('function Build-FinalPdf(', 1)[1].split('\nfunction ', 1)[0]
if 'Build-FinalPdfLegacy' not in _bfp:
    raise SystemExit('unapproved workspaces must fall back to the V4.1 build path')
for fn in ['function Write-HistoryEvent', 'function Save-LayoutSnapshot']:
    blk = server.split(fn, 1)[1].split('\nfunction ', 1)[0]
    if 'Test-InputHistoryEnabled' not in blk: raise SystemExit(f'{fn} must be gated on approval')

# Archive failures must roll the transaction back, not be swallowed.
_arch = server.split('function New-FinalArchive', 1)[1].split('\nfunction ', 1)[0]
if "return ''" in _arch: raise SystemExit('New-FinalArchive must not swallow failures')
if 'throw (' not in _arch: raise SystemExit('New-FinalArchive must rethrow on failure')
if _arch.index('New-SnapshotPin') < _arch.index('Move-Item -LiteralPath $stage'):
    raise SystemExit('pins must be created after the archive is moved into place')
if 'Get-Structure' in _arch: raise SystemExit('archive metadata must come from the immutable build snapshot')
if 'function Remove-FinalArchiveArtifacts' not in server:
    raise SystemExit('rollback must clean partial archives and pins')

# Visual analysis must not run while the render lock and Excel are held.
if 'Invoke-PostRenderAnalysis' in _rw.split('Invoke-WithRenderLock', 1)[1].split('\n    # V5-P1: ここではレンダリングロック', 1)[0]:
    raise SystemExit('analysis must run after the render lock is released')

# The fingerprint must cover the analysis toolchain.
_fp = server.split('function Get-RenderEnvironmentFingerprint', 1)[1].split('\nfunction ', 1)[0]
for needed in ['pdfBox=', 'dpi=', 'analyzer=']:
    if needed not in _fp: raise SystemExit(f'render environment fingerprint missing {needed}')

# Added and removed sheets are part of the diff.
_cmp = server.split('function Compare-SnapshotVisual', 1)[1].split('\nfunction ', 1)[0]
for needed in ['addedSheets', 'removedSheets']:
    if needed not in _cmp: raise SystemExit(f'comparison must report {needed}')
for needed in ['function addedSheetSet', '追加 ${added}', "badge('追加','attention')", 'addedSheetSet(p.workbookId).has(name)']:
    if needed not in appjs:
        raise SystemExit(f'added sheets must appear in the change column/filter: {needed}')

# Detailed visual diff must be background-generated, cache-addressed by IDs,
# and serve only manifest-listed PNG assets.
for needed in ['function Get-DiffDetailContext', 'function Start-DiffDetailJob',
               'function Invoke-DiffDetailJobFromFile', 'function Serve-DiffPage',
               '$Script:DiffDetailThreshold = 18', '$Script:DiffDetailDpi = 150',
               'Assert-SafeStorageSegment $CurrentSnapshotId',
               "allowedAssets = @('before','after','before-mask','before-overlay','mask','overlay')"]:
    if needed not in server:
        raise SystemExit(f'detailed visual diff feature missing: {needed}')
for needed in ['diffDetailRequestPath', 'diffPrepareRequestBody',
               'fromSnapshotId:from', 'toSnapshotId:to', 'body.sheetKey=sheetKey',
               '選んだ2版を視覚比較', '比較元と比較先を入れ替え',
               'renderSnapshotHistorySelectionHint', 'visualCompareReady', 'unavailableReason',
               "scope:String(cmp.scope||'automatic')", "['deferred','failed'].includes"]:
    if needed not in appjs:
        raise SystemExit(f'historical visual diff UI missing: {needed}')
if "if (-not [string]::IsNullOrWhiteSpace($DiffJobPath))" not in server:
    raise SystemExit('diff-detail job child mode is missing')
diff_script=(root/'app/tools/diff-image-pages.ps1').read_text(encoding='utf-8-sig')
for needed in ['PDFToImage','ReportBinderDiffEngine','MinimumRegionPixels','Threshold','ConvertTo-Json']:
    if needed not in diff_script:
        raise SystemExit(f'diff image wrapper feature missing: {needed}')
if "Join-Path ([IO.Path]::GetTempPath()) ('rb-diff-'" not in diff_script:
    raise SystemExit('diff raster work files must stay under a short TEMP path')
if "Join-Path $outputFull ('.raster-'" in diff_script:
    raise SystemExit('long diff raster work path remains under the history cache')
diff_engine=(root/'app/tools/DiffImageEngine.cs').read_text(encoding='utf-8')
for needed in ['FindRegions','minimumRegionPixels','pageSizeChanged','beforeMaskFile','afterOverlayFile','DashStyle.Dash']:
    if needed not in diff_engine:
        raise SystemExit(f'diff image engine feature missing: {needed}')

# The shipped jar must actually contain the analyzer, or the feature silently does nothing.
with zipfile.ZipFile(root/'app/lib/pdfbox/ReportPdfComposer.jar') as zf:
    if 'PdfPageAnalyzer.class' not in set(zf.namelist()):
        raise SystemExit('ReportPdfComposer.jar must contain PdfPageAnalyzer.class')

# --- V5-P2: hot-path I/O regressions -------------------------------------
# Get-AppConfig must not rewrite config.json on every read: it is reached from
# Get-Paths -> Get-WorkspacePath, i.e. from almost every function.
_cfg = server.split('function Get-AppConfig', 1)[1].split('\nfunction ', 1)[0]
if '$Script:ConfigMergeChanged' not in _cfg:
    raise SystemExit('Get-AppConfig must only write config.json when defaults were merged in')
if '$Script:AppConfigCache' not in _cfg:
    raise SystemExit('Get-AppConfig must cache its result')
for fn, cache in [('function Get-Paths', '$Script:PathsCache'),
                  ('function Get-WorkspacePolicy', '$Script:WorkspacePolicyCache'),
                  ('function Get-InputHistorySizeMb', '$Script:HistorySizeCache')]:
    blk = server.split(fn, 1)[1].split('\nfunction ', 1)[0]
    if cache not in blk: raise SystemExit(f'{fn} must be cached ({cache})')
if 'function Reset-ConfigCaches' not in server:
    raise SystemExit('missing Reset-ConfigCaches')
for fn in ['function Save-AppConfig', 'function Ensure-Package']:
    blk = server.split(fn, 1)[1].split('\nfunction ', 1)[0]
    if 'Reset-ConfigCaches' not in blk: raise SystemExit(f'{fn} must invalidate the config caches')

# The state poll must not reload structure.json once per workbook.
_glc = server.split('function Get-LatestComparison([string]', 1)[1].split('\nfunction ', 1)[0]
if '$Workbook = $null' not in _glc:
    raise SystemExit('Get-LatestComparison must accept an already-loaded workbook')
if 'Get-WorkbookChangeSummary $Language ([string]$w.workbookId) $w' not in server:
    raise SystemExit('Get-StatePayload must pass the loaded workbook into the change summary')

# The scheduler child must be stopped explicitly, not only by parent-PID polling.
_srv = server.split('function Start-LocalTcpServer', 1)[1].split('\nfunction ', 1)[0]
if 'Stop-AutoSchedulerProcess' not in _srv:
    raise SystemExit('Start-LocalTcpServer must stop the auto-scheduler child on exit')
if _srv.index('try {') > _srv.index('$tcp.Start()'):
    raise SystemExit('listener startup must be covered by the scheduler-stop finally block')

# Workspace switches must not reuse the previous input-history size.
_hsize = server.split('function Get-InputHistorySizeMb', 1)[1].split('\nfunction ', 1)[0]
if '$Script:HistorySizeCacheKey' not in _hsize or 'Get-InputHistoryRoot' not in _hsize:
    raise SystemExit('history-size cache must be keyed by the current input-history root')
_reset = server.split('function Reset-ConfigCaches', 1)[1].split('\nfunction ', 1)[0]
if '$Script:HistorySizeCache' not in _reset:
    raise SystemExit('Reset-ConfigCaches must invalidate the history-size cache')

# Scheduler shutdown should be graceful and should not leave control files behind.
_child = server.split('function Invoke-AutoSchedulerFromFile', 1)[1].split('\nfunction ', 1)[0]
if 'Start-Sleep -Seconds 10' in _child or 'Start-Sleep -Milliseconds 500' not in _child:
    raise SystemExit('auto scheduler must check for stop requests during the 10-second interval')
if "Remove-Item -LiteralPath $path" not in _child:
    raise SystemExit('auto scheduler must remove its control files on exit')
_stop = server.split('function Stop-AutoSchedulerProcess', 1)[1].split('\nfunction ', 1)[0]
for needed in ['WaitForExit(2500)', '$Script:AutoSchedulerProcessId = 0', "Remove-Item -LiteralPath $path"]:
    if needed not in _stop: raise SystemExit(f'scheduler stop cleanup missing: {needed}')

# Rolled-back transactions must not leave old output PDFs in final-backups forever.
_restore = server.split('function Restore-FinalTransaction', 1)[1].split('\nfunction ', 1)[0]
if 'Remove-FinalTransactionBackupDir' not in _restore:
    raise SystemExit('rolled-back transactions must remove their backup directory')
_recover = server.split('function Recover-FinalTransactions', 1)[1].split('\nfunction ', 1)[0]
if 'AddDays(-30)' not in _recover or 'Remove-FinalTransactionBackupDir' not in _recover:
    raise SystemExit('journal pruning must also remove matching final-backup directories')

# All documentation must state that V5 preview uses the browser viewer, not PDF.js.
for rel in ['README.md', 'THIRD_PARTY_NOTICES.md', 'docs/THIRD_PARTY_SETUP.md', 'app/web/pdfjs/README.md']:
    doc = (root/rel).read_text(encoding='utf-8-sig')
    if 'ブラウザ内蔵' not in doc:
        raise SystemExit(f'PDF.js usage documentation is incomplete: {rel}')

# Release ZIPs must set the UTF-8 name flag, or Japanese Windows' built-in
# extractor mangles 日本語管理.vbs and the launchers become unusable.
_pkg = (root/'app/tools/package-release.ps1').read_text(encoding='utf-8-sig')
if 'New-Utf8Zip' not in _pkg or 'System.Text.Encoding]::UTF8' not in _pkg:
    raise SystemExit('package-release.ps1 must build the ZIP with UTF-8 entry names')
if any(l.strip().startswith('Compress-Archive') for l in _pkg.splitlines()):
    raise SystemExit('Compress-Archive does not set the UTF-8 name flag; use New-Utf8Zip')

# A shared-Excel batch job collects $Script:PendingAnalysis itself; Render-Workbook
# must not clear it first, or visual-hash analysis never runs for any batch render.
_tail = _rw.split('$renderResult = Invoke-WithRenderLock', 1)[1]
if '$KeepExcelOpen) { return $renderResult }' not in _tail:
    raise SystemExit('Render-Workbook must hand PendingAnalysis to the caller when KeepExcelOpen')
_pos_clear = _tail.index('$Script:PendingAnalysis = $null')
_pos_guard = _tail.index('if ($KeepExcelOpen) { return $renderResult }')
if _pos_clear < _pos_guard:
    raise SystemExit('PendingAnalysis must not be cleared before the KeepExcelOpen guard')
_job = server.split('function Invoke-RenderJobFromFile', 1)[1].split('\nfunction ', 1)[0]
if '$deferredAnalyses' not in _job or 'Invoke-PostRenderAnalysis' not in _job:
    raise SystemExit('batch render job must run the deferred visual-hash analysis')

# PS 5.1 turns native stderr into a terminating error under $ErrorActionPreference='Stop'.
# PDFBox writes font warnings to stderr, so every Java call must go through the helper.
if 'function Invoke-NativeCapture' not in server:
    raise SystemExit('missing Invoke-NativeCapture')
_nc = server.split('function Invoke-NativeCapture', 1)[1].split('\nfunction ', 1)[0]
if "$ErrorActionPreference = 'Continue'" not in _nc:
    raise SystemExit('Invoke-NativeCapture must relax ErrorActionPreference around the call')
for line_no, line in enumerate(server.splitlines(), 1):
    if '2>&1' in line and 'Invoke-NativeCapture' not in line and not line.strip().startswith('#'):
        if '$lines = @(& $FilePath' in line:
            continue
        raise SystemExit(f'raw native 2>&1 capture outside Invoke-NativeCapture at line {line_no}')

# Starting one language must not kill the other language's running server.
launch = (root/'app/launch.ps1').read_text(encoding='utf-8-sig')
_stale = launch.split('function Stop-StaleUiServerProcesses', 1)[1].split('\nfunction ', 1)[0]
for needed in ['$TargetMode', '-autoschedulerpath', '-renderjobpath', '-diffjobpath']:
    if needed not in _stale:
        raise SystemExit(f'stale-server cleanup must be scoped ({needed})')
if 'Stop-StaleUiServerProcesses $server $Mode' not in launch:
    raise SystemExit('stale-server cleanup must be called with the current mode')

# 2026-07-30 review fixes ------------------------------------------------------
# A lazily generated single sheet must not mark the whole diff detail as ready,
# or sheets left pending by a failed full run can never be generated again.
_diff_core = server.split('function Invoke-DiffDetailJobCore', 1)[1].split('\nfunction ', 1)[0]
if '$pendingSheets' not in _diff_core or '$failedSheets' not in _diff_core:
    raise SystemExit('diff detail status must be derived from pending and failed sheets')
if "if (Test-Path -LiteralPath $detailPath)" not in _diff_core or "if (-not [string]::IsNullOrWhiteSpace($requestedSheetKey) -and (Test-Path -LiteralPath $detailPath))" in _diff_core:
    raise SystemExit('full diff retry must reload existing detail so failed unchanged sheets remain retryable')
if "Set-NoteProperty $active 'joinedExistingDiffJob' $true" not in server:
    raise SystemExit('lazy diff retry must distinguish joining an existing job from its own failure')
for needed in ["@('pending','generating') -contains", "$sheet.status = 'failed'", "$detail.status = 'ready'", "$detail.status = 'failed'"]:
    if needed not in _diff_core:
        raise SystemExit(f'diff detail completion status check missing: {needed}')
if "$detail.status = 'ready'; $detail.message = ''; $detail.generatedAt" in _diff_core:
    raise SystemExit('diff detail must not be marked ready unconditionally')

engine = (root/'app/tools/DiffImageEngine.cs').read_text(encoding='utf-8-sig')
# Add-Type on Windows PowerShell 5.1 compiles with the C# 5 CodeDom provider.
for forbidden in ['=>', '$"', 'nameof(', '?.', 'out var ', 'using static ']:
    if forbidden in engine:
        raise SystemExit(f'DiffImageEngine.cs must stay C# 5 compatible for Add-Type: {forbidden}')
# DrawImageUnscaled rescales by the source DPI metadata despite its name.
if 'DrawImageUnscaled(' in engine:
    raise SystemExit('page normalization must use an explicit destination rectangle, not DrawImageUnscaled')
# Change pixels within 2x padding belong to one region; without this a single edited
# number becomes one box per glyph and a shifted row produces thousands of boxes.
for needed in ['private static bool[] Dilate(', 'bool[] grouped = Dilate(candidate, width, height, padding);', 'if (!grouped[seed] || visited[seed]) continue;']:
    if needed not in engine:
        raise SystemExit(f'diff region merging missing: {needed}')
if 'MaximumRegionsPerPage' not in engine or 'pixelRegions.Count > MaximumRegionsPerPage' not in engine:
    raise SystemExit('diff engine must cap the highlighted region count per page')
# Asset names stay short: the cache path is already ~236 chars on the real share.
_engine_code = '\n'.join(l for l in engine.splitlines() if not l.strip().startswith('//'))
if '"page-' in _engine_code or '-before.png' in _engine_code or '-overlay.png' in _engine_code:
    raise SystemExit('diff asset file names must stay short for MAX_PATH')
for needed in ['string prefix = pageNumber.ToString("0000")', 'stem + "-b.png"', 'stem + "-a.png"']:
    if needed not in engine:
        raise SystemExit(f'short diff asset naming missing: {needed}')
if '$Script:DiffDetailAlgorithmVersion = 6' not in server:
    raise SystemExit('cache layout changes must bump DiffDetailAlgorithmVersion to 6')
if server.count(')).Substring(7, 16)') < 2:
    raise SystemExit('diff detail cache keys must use at least 64 bits')
if 'function Test-DiffDetailMatchesContext' not in server:
    raise SystemExit('stored diff detail identity validation is missing')
_match = server.split('function Test-DiffDetailMatchesContext', 1)[1].split('\nfunction ', 1)[0]
for needed in ['workbookId', 'scope', 'baselineSnapshotId', 'baselineVersionId', 'currentSnapshotId', 'currentVersionId']:
    if needed not in _match:
        raise SystemExit(f'diff detail identity validation missing: {needed}')
_serve_diff = server.split('function Serve-DiffPage', 1)[1].split('\nfunction ', 1)[0]
if 'Test-DiffDetailMatchesContext $detail $diffContext' not in _serve_diff:
    raise SystemExit('diff page assets must validate the stored detail identity before serving')
appjs = (root/'app/web/app.js').read_text(encoding='utf-8-sig')
if "addEventListener('click',prepareDiffDetail)" in appjs:
    raise SystemExit('diff retry click must not pass MouseEvent as sheetKey')
for needed in ["addEventListener('click',()=>prepareDiffDetail())", "if(selected&&['deferred','failed'].includes", "!['deferred','failed'].includes(targetStatus)", "targetStatus==='failed'&&!joinedExisting", "sheetStatus==='failed'"]:
    if needed not in appjs:
        raise SystemExit(f'diff retry UI regression: {needed}')

# Third-party installation must stay reproducible and fail closed.
installer = (root/'app/tools/install-thirdparty.ps1').read_text(encoding='utf-8-sig')
for needed in ['2.0.37', 'Verify-HashFile', 'Verify-SriSha512', 'dist.integrity', 'binary.package.checksum', 'PreferSystemJava', 'verify-thirdparty.ps1']:
    if needed not in installer: raise SystemExit(f'third-party installer hardening missing: {needed}')
verifier = (root/'app/tools/verify-thirdparty.ps1').read_text(encoding='utf-8-sig')
for needed in ['RequirePortableJava', 'SHA-512 verified', 'verified npm package tarball', 'PdfPageAnalyzer.class']:
    if needed not in verifier: raise SystemExit(f'third-party verifier check missing: {needed}')
package_release = (root/'app/tools/package-release.ps1').read_text(encoding='utf-8-sig')
for needed in ['Invoke-ThirdPartyCheck $true', 'Assert-StagedDependencies', 'JAVA_VERSION.txt']:
    if needed not in package_release: raise SystemExit(f'release dependency gate missing: {needed}')
if 'app/thirdparty-cache/' not in (root/'.gitignore').read_text(encoding='utf-8'):
    raise SystemExit('third-party cache must be ignored by git')

print('selfcheck ok')
