from pathlib import Path
import json, os, re, shutil, subprocess, zipfile

root = Path(__file__).resolve().parents[2]
required = [
    '資料をPDFにまとめる.cmd','共有フォルダー用フォルダー作成.cmd','README.md','THIRD_PARTY_NOTICES.md','app/server.ps1','app/default-config.json','app/runtime-version.json','app/launch.ps1',
    'app/web/index.html','app/web/style.css','app/web/app.js','app/web/diff-worker.js','app/lib/pdfbox/ReportPdfComposer.jar',
    'app/lib/pdfbox/src/ReportPdfComposer.java','app/lib/pdfbox/src/BatchPdfSplitter.java','app/lib/pdfbox/src/PdfBatchRasterizer.java','app/lib/pdfbox/build.ps1',
    'app/tools/install-thirdparty.ps1','app/tools/install-thirdparty.cmd','app/tools/verify-thirdparty.ps1','app/tools/select-folder.ps1','app/tools/package-release.ps1',
    'app/tools/diff-image-pages.ps1','app/tools/diff-image-batch.ps1','app/tools/DiffImageEngine.cs',
    'app/tools/fixtures/structure-v1-ja.json','app/tools/fixtures/structure-mixed-order.json',
    'app/tools/fixtures/structure-v2-generalized-ja.json','app/tools/fixtures/structure-v3-migration-expected.json','app/tools/schema-v3-selfcheck.ps1','app/tools/source-adapter-selfcheck.ps1',
    'app/tools/pdf-source-adapter-selfcheck.ps1','app/tools/word-source-adapter-selfcheck.ps1','app/tools/word-render-worker.ps1','app/tools/powerpoint-render-worker.ps1','app/tools/create-word-adapter-fixtures.py','app/tools/history-generalization-selfcheck.ps1','app/tools/final-composition-selfcheck.py','app/tools/operational-readiness-selfcheck.ps1',
    'app/tools/create-scale-benchmark-fixtures.py','app/tools/create-scale-benchmark-workbooks.mjs','app/tools/scale-benchmark.ps1','app/tools/ci-selfcheck.ps1','app/tools/history-logic-selfcheck.ps1','app/tools/create-pdf-diff-corpus.py','app/tools/pdf-diff-corpus-selfcheck.ps1','tests/pdf-diff-corpus.mjs','docs/PDF_DIFF_CORPUS.md','app/tools/ci-requirements.txt',
    '.github/workflows/thirdparty-check.yml','.github/workflows/release.yml',
    'tests/diff-regression.mjs','docs/API.md','docs/THIRD_PARTY_SETUP.md','docs/ReportBinder_UIUX改修指示書_V4.md','docs/GENERALIZED_DOCUMENT_PACK_DESIGN.md','docs/OPERATIONS_GUIDE.md','docs/SCALE_BENCHMARK.md','docs/benchmarks/scale-benchmark-windows-20260807.json'
]
missing=[x for x in required if not (root/x).exists()]
if missing: raise SystemExit('missing: '+', '.join(missing))
if (root/'app/config.json').exists(): raise SystemExit('runtime app/config.json must not be distributed')
json.loads((root/'app/default-config.json').read_text(encoding='utf-8'))

scale_benchmark=(root/'app/tools/scale-benchmark.ps1').read_text(encoding='utf-8')
for needed in ['Interactive Office session is required', "'/api/scan-updates'", "onlyUpdated=$true",
               "'/api/final/build-all'", 'ConvertTo-ReadinessSummary', 'acceptanceThresholds',
               'exactlyThreeSourcesDetected', 'Remove-BenchmarkWorkspace']:
    if needed not in scale_benchmark: raise SystemExit(f'scale benchmark acceptance coverage missing: {needed}')
scale_fixture=(root/'app/tools/create-scale-benchmark-fixtures.py').read_text(encoding='utf-8')
for needed in ['--update-only', 'create_pdf', 'create_word', 'WD_BREAK.PAGE', 'document.add_table']:
    if needed not in scale_fixture: raise SystemExit(f'scale PDF/Word fixture coverage missing: {needed}')
scale_workbooks=(root/'app/tools/create-scale-benchmark-workbooks.mjs').read_text(encoding='utf-8')
for needed in ['@oai/artifact-tool', '--update-only', 'workbook.inspect', 'Dept04', 'large-excel-01-v2-preview.png']:
    if needed not in scale_workbooks: raise SystemExit(f'scale workbook fixture coverage missing: {needed}')
scale_report=json.loads((root/'docs/benchmarks/scale-benchmark-windows-20260807.json').read_text(encoding='utf-8'))
if scale_report.get('schemaVersion') != 2 or not scale_report.get('acceptance', {}).get('passed'):
    raise SystemExit('large mixed-source benchmark acceptance report is missing or failed')
if scale_report.get('workload', {}).get('updatedSourceCount') != 3:
    raise SystemExit('large benchmark must detect exactly one updated PDF, Word, and Excel source')
ci_selfcheck=(root/'app/tools/ci-selfcheck.ps1').read_text(encoding='utf-8')
for needed in ['Assert-PowerShellSyntax', 'Assert-VersionDocumented', 'Diff regression',
               'Final composition regression', 'Operational readiness regression',
               'History logic regression', 'ZIP release smoke test', 'Assert-PackageContents']:
    if needed not in ci_selfcheck: raise SystemExit(f'CI acceptance coverage missing: {needed}')
ci_workflow=(root/'.github/workflows/thirdparty-check.yml').read_text(encoding='utf-8')
for needed in ['pull_request:', 'branches:', '- main', 'install-thirdparty.ps1 -Force',
               'ci-requirements.txt', 'build.ps1', 'ci-selfcheck.ps1']:
    if needed not in ci_workflow: raise SystemExit(f'GitHub CI workflow coverage missing: {needed}')
release_workflow=(root/'.github/workflows/release.yml').read_text(encoding='utf-8')
for needed in ['workflow_dispatch:', 'package-release.ps1', '-SharedFolderOnly', 'actions/upload-artifact@v6']:
    if needed not in release_workflow: raise SystemExit(f'GitHub release workflow coverage missing: {needed}')
for ps1 in ['app/server.ps1','app/launch.ps1','app/lib/pdfbox/build.ps1','app/tools/install-thirdparty.ps1','app/tools/verify-thirdparty.ps1','app/tools/select-folder.ps1','app/tools/package-release.ps1','app/tools/diff-image-pages.ps1','app/tools/diff-image-batch.ps1']:
    if not (root/ps1).read_bytes().startswith(b'\xef\xbb\xbf'): raise SystemExit(f'PowerShell must be UTF-8 BOM: {ps1}')

launcher_cmd=(root/'資料をPDFにまとめる.cmd').read_text(encoding='utf-8-sig')
for needed in ['app\\launch.ps1', 'start "" /b', 'powershell.exe']:
    if needed not in launcher_cmd:
        raise SystemExit(f'CMD launcher regression: {needed}')
if '-Mode ' in launcher_cmd:
    raise SystemExit('the user-facing launcher must not expose a language mode')
shared_folder_cmd=(root/'共有フォルダー用フォルダー作成.cmd').read_text(encoding='utf-8-sig')
for needed in ['app\\tools\\package-release.ps1', '-SharedFolderOnly', 'pause']:
    if needed not in shared_folder_cmd:
        raise SystemExit(f'shared-folder CMD regression: {needed}')

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
if 'previousSnapshotId' in '\n'.join(
        line for line in compare_block.splitlines() if not line.strip().startswith('#')):
    raise SystemExit('automatic comparison must not fall back to an arbitrarily old previous snapshot')
if '$baseSnap -eq $CurrentSnapshotId' not in compare_block:
    raise SystemExit('current version must become the baseline without self-comparison')
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
_pdf_index=server.split('function Get-ContentPdfSheetIndex',1)[1].split('\nfunction ',1)[0]
if '[void]$Script:ContentPdfSheetIndexCache.Remove($oldestKey)' not in _pdf_index:
    raise SystemExit('PDF index cache eviction must not leak a Boolean into the function pipeline')
_availability=server.split('function Get-HistoryRenderVersionAvailability',1)[1].split('\nfunction ',1)[0]
for needed in ['visualHashAvailable','contentPdfAvailable','missingSheets','Get-ContentPdfSheetIndex','Resolve-ContentPdfSheetPathFromIndex']:
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

# Comparison metadata is returned immediately; the browser renders and analyzes only the visible page.
_skeleton=server.split('function New-DiffDetailSkeleton',1)[1].split('\nfunction Get-DiffDetail',1)[0]
for needed in ["status = 'ready'", 'unchangedPageNumbers', '表示したページをブラウザで比較します。']:
    if needed not in _skeleton: raise SystemExit(f'browser comparison skeleton missing: {needed}')
_get_detail=server.split('function Get-DiffDetail(',1)[1].split('\nfunction ',1)[0]
if '$detail = New-DiffDetailSkeleton $Language $context' not in _get_detail or "'performance'" not in _get_detail:
    raise SystemExit('diff detail must return metadata without starting or reading a comparison-image job')
for forbidden in ['diff-detail.json', 'diff-job.json', 'Read-RenderJobStatus']:
    if forbidden in _get_detail: raise SystemExit(f'diff detail still depends on server-generated comparison assets: {forbidden}')
_appjs_early=(root/'app/web/app.js').read_text(encoding='utf-8-sig')
_diff_worker=(root/'app/web/diff-worker.js').read_text(encoding='utf-8-sig')
for needed in ['ensureDiffPdfJs', 'fetchDiffPdfDocument', 'renderDiffPdfPage', 'buildDiffBrowserPage', 'diffBrowserPageCache', "new Worker(new URL('diff-worker.js?v=20260808_v23'"]:
    if needed not in _appjs_early: raise SystemExit(f'browser PDF comparison missing: {needed}')
for needed in ['buildRowDescriptors', 'alignRows', 'mappedRowY', 'ROW_ALIGNMENT_BAND',
               'choosePixelRowMapping', 'pixel-scale-and-row-shift', 'clearlyImproves',
               'choosePixelColumnMapping', 'mappedColumnX', 'refinePixelColumnMapping',
               'scaleX', 'scaleY', 'maxLocalShiftJump', 'tolerantPixelDifference',
               'mergeNearbyComponents', 'gapX=2', 'self.onmessage']:
    if needed not in _diff_worker: raise SystemExit(f'human-aligned browser diff worker missing: {needed}')
if 'for(let dy=-4;dy<=4;dy++)' in _diff_worker:
    raise SystemExit('browser diff worker still uses fixed +/-4px-only alignment')
for needed in ['analysis.alignmentAdjusted', 'changedRatio:Number(analysis?.changedRatio||0)',
               '行・列の追加や幅・倍率による位置ずれを補正']:
    if needed not in _appjs_early:
        raise SystemExit(f'human-aligned browser diff UI missing: {needed}')
for forbidden in ['prepareDiffDetail(', "api('/api/history/diff/prepare'", '/api/history/diff-page']:
    if forbidden in _appjs_early: raise SystemExit(f'browser still starts server comparison-image generation: {forbidden}')
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
for route in ['/api/state','/api/v2/state','/api/diagnostics/run','/api/v2/pack-templates','/api/v2/packs','/api/v2/sources/','/api/paths','/api/workbooks/register-batch','/api/workbooks/render/start','/api/jobs/status','/api/jobs/cancel','/api/pages/reorder','/api/pages/sort-by-sheet','/api/final/readiness','/api/final/build','/api/final/publish','/api/final/file','/api/scan-updates','/api/history/diff-detail','/api/history/diff/prepare','/api/history/diff-page']:
    if route not in server: raise SystemExit(f'route not found: {route}')
for needed in ['/api/history/diff-review','function Get-DiffReviewStatePath','function Get-DiffReviewStateForContext','function Set-DiffReviewState',
               'confirmedSheetKeys','id="diff-unreviewed-only"','id="diff-confirm-sheet"','function toggleDiffSheetReviewed']:
    review_sources = server + (root/'app/web/index.html').read_text(encoding='utf-8') + (root/'app/web/app.js').read_text(encoding='utf-8-sig')
    if needed not in review_sources: raise SystemExit(f'diff review workflow missing: {needed}')
for needed in ["^/api/v2/packs/([^/]+)/review$", 'function Get-PackReviewSnapshot', 'function Invoke-PackReviewAction',
               'submittedFingerprints', "'review-stale'", 'id="pack-review-submit"', 'id="pack-review-approve"',
               'id="pack-review-request-changes"', 'id="pack-review-events"', 'function performPackReviewAction']:
    review_sources = server + (root/'app/web/index.html').read_text(encoding='utf-8') + (root/'app/web/app.js').read_text(encoding='utf-8-sig')
    if needed not in review_sources: raise SystemExit(f'pack approval workflow missing: {needed}')
for needle in ['function New-DocumentPack', 'function Copy-DocumentPack', 'function Set-DocumentPackArchived',
               "^/api/v2/packs/([^/]+)/duplicate$", "^/api/v2/packs/([^/]+)/archive$", "^/api/v2/packs/([^/]+)/restore$"]:
    if needle not in server: raise SystemExit(f'pack lifecycle API missing: {needle}')
for needle in [
    'Update-StructureLocked','Read-StructureUnlocked','Write-StructureUnlocked','Initialize-Or-MigrateStructure','structure.json.v1.bak',
    'Get-VolumeStateKey','builtFingerprint','Get-FinalBuildInputSnapshot','contentPdfLastWriteUtcTicks','contentPdfSize',
    'manifest_${Volume}_${cat}.json','~building_${Volume}_${cat}.pdf','Require-WorkbookCategory','System.ArgumentException',
    'Mark-VolumeNeedsRebuild','staleReasons','Sort-PagesBySheet','Insert-PageInSheetOrder','Renumber-VolumeOrder',
    'CenterHorizontally = $true','LeftMargin = Convert-CmToPt 1.2','RightMargin = Convert-CmToPt 1.2','punchShiftPt=(Convert-CmToPt 0.2)',
    'lib\\java\\bin\\java.exe','Apply-DefaultNumberingPerVolume','first-page-none','ExcelPrintProfileVersion',
    'ConvertTo-NormalizedPageRange','Resolve-OutputFileNamePattern','physicalPages=$snap.physicalPages','schemaVersion=3'
]:
    if needle not in server: raise SystemExit(f'server feature not found: {needle}')
if len(re.findall(r'(?<!function )Save-Structure\s',server)):
    raise SystemExit('direct Save-Structure use is forbidden')

# OpenXML workbooks must receive the standard print profile before Excel opens them,
# avoiding the slow PageSetup COM round-trip on every numeric worksheet.
if 'function Remove-XlsxHeaderFooterXml' in server:
    raise SystemExit('header/footer-only package preparation must be replaced by full print preparation')
_print_package=server.split('function Prepare-XlsxPrintPackage',1)[1].split('\nfunction Clear-ExcelPageSetupHeadersAndFooters',1)[0]
for needed in ["pageSetUpPr 'fitToPage' '1'", "printOptions 'horizontalCentered' '1'",
               "pageMargins 'left' '0.47244094'", "pageMargins 'top' '0.31496063'",
               "pageSetup 'fitToWidth' '1'", "pageSetup 'fitToHeight' '1'",
               'printSettingsPrepared=($sheetsPrepared -gt 0)']:
    if needed not in _print_package:
        raise SystemExit(f'OpenXML print preparation missing: {needed}')
_render_workbook=server.split('function Render-Workbook(',1)[1].split('\nfunction ',1)[0]
for needed in ['Prepare-XlsxPrintPackage $tmpPath', '$packagePrintSettingsPrepared',
               'if (-not $packagePrintSettingsPrepared)', 'timingsMs = $timingsMs',
               '$timingsMs.excelOpen', '$timingsMs.excelExportAndSplit']:
    if needed not in _render_workbook:
        raise SystemExit(f'PDF render optimization/measurement missing: {needed}')
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
for needed in ['$Script:StructureReadCache', 'LastWriteTimeUtc.Ticks', '$file.Length']:
    if needed not in get_block: raise SystemExit(f'Get-Structure cold-read cache missing: {needed}')
_update_structure = server.split('function Update-StructureLocked',1)[1].split('\nfunction ',1)[0]
if '$Script:StructureReadCache.Clear()' not in _update_structure:
    raise SystemExit('structure writes must invalidate the read cache')
# Fingerprint must represent rendered inputs, not the source Excel hash.
fp_block=server.split('function Get-FinalBuildInputSnapshot',1)[1].split('function Get-FinalBuildFingerprint',1)[0]
if 'currentExcelHash' in fp_block: raise SystemExit('currentExcelHash must not be in final build fingerprint')
for needle in ['contentPdfSize','contentPdfLastWriteUtcTicks','lastRenderedVersionId']:
    if needle not in fp_block: raise SystemExit(f'fingerprint input missing: {needle}')
# Legacy category endpoints must fail closed; the shared snapshot resolves either a
# validated built-in category or an existing document pack.
for fn in ['Build-FinalPdf','Serve-FinalPdfByVolume']:
    block=server.split(f'function {fn}',1)[1].split('\nfunction ',1)[0]
    if 'Require-WorkbookCategory' not in block: raise SystemExit(f'category not required: {fn}')
snapshot_block=server.split('function Get-FinalBuildInputSnapshot',1)[1].split('\nfunction ',1)[0]
if 'Resolve-DocumentPackScope' not in snapshot_block:
    raise SystemExit('final build snapshot must require an existing document pack')
# Long-running Excel/Java work must be outside the structure transaction body.
build=server.split('function Build-FinalPdf',1)[1].split('function Get-StatePayload',1)[0]
java_pos=build.find('ReportPdfComposer'); commit_pos=build.find('$commit=Update-StructureLocked')
if java_pos < 0 or commit_pos < java_pos: raise SystemExit('final composer/commit order is invalid')

appjs=(root/'app/web/app.js').read_text(encoding='utf-8-sig')
for needle in ['renderGlobalHeader','renderStepBar','renderNavBadges','aggregateFinalState','volumeReadiness','isEditing','lastPageBoardRenderSignature','sortPagesBySheet','/api/pages/sort-by-sheet','category:activePreset','openFinalVolume(volume, category=activePreset)','notice-actions','insertedAtEndCount','modalReturnFocus','render-all-btn','openDiffDetail','moveDiffRegion','syncDiffScroll','fetchDiffPdfDocument','buildDiffBrowserPage','rememberPageLayoutUndo','preparePageThumbnails','page-filter-empty','pageMatchesSearch','applyPageThumbnailSize','savePageFromThumbnail','restorePageSettings','data-thumb-editor','data-thumb-page-range','savePackSettings','pack-output-pattern']:
    if needle not in appjs: raise SystemExit(f'ui feature not found: {needle}')
for needle in ['function renderPackSwitcher', 'function openPackEditor', 'function submitPackEditor', 'function duplicatePack', 'function archivePack', 'function restorePack', 'includeArchived=true', 'ReportBinderPackId']:
    if needle not in appjs: raise SystemExit(f'pack management UI behavior missing: {needle}')
for dead in ["bind('refresh-btn'","bind('load-files-btn'","bind('save-paths-btn'","bind('render-updated-btn'","bind('render-selected-pages-btn'","$('category-heading')","$('category-caption')"]:
    if dead in appjs: raise SystemExit(f'dead ui code remains: {dead}')
if "body:{volumes:collectBoardVolumes()}" in appjs: raise SystemExit('page reorder must send category')
if 'class="drag-handle page-thumb-drag"' in appjs:
    raise SystemExit('thumbnail drag handle must not cover the page preview')
if 'class="sr-only" type="checkbox" data-page-check' not in appjs:
    raise SystemExit('thumbnail selection must remain keyboard-accessible without a visible overlay')

html=(root/'app/web/index.html').read_text(encoding='utf-8')
for needle in ['workspace-top','step-bar','id="pack-menu-button"','id="pack-menu-list"','id="create-pack-btn"','id="pack-editor-modal"','global-final-status','nav-excel-count','sort-by-sheet-btn','build-all-btn','render-all-btn','unregister-selected-btn','id="final-target-grid"','id="bulk-target-buttons"','notice-actions','app-menu-popover','<symbol id="i-home"','<symbol id="i-edit"','id="diff-modal"','id="diff-before-viewport"','id="diff-after-viewport"','id="diff-mode-overlay"','id="page-view-thumbnail-btn"','id="page-layout-undo-btn"','id="page-command-bar"','id="page-search-input"','id="page-thumbnail-size"']:
    if needle not in html: raise SystemExit(f'html feature not found: {needle}')
for dead in ['collapse-hint','pagination-lite','info-dot','menu-col','recent-folder-list','page-output-summary','category-card','workflow-card']:
    if dead in html: raise SystemExit(f'dead ui remains: {dead}')
for symbol in '⌂□▦▤▣△⋮◉⌄↻⌕⊘›':
    if symbol in html: raise SystemExit(f'font symbol remains in html: {symbol}')

css=(root/'app/web/style.css').read_text(encoding='utf-8')
for needle in ['--accent:#5e6ad2','--sidebar-w:248px','backdrop-filter','box-shadow:none','.badge.attention','background:var(--attention-subtle)','font-weight:600','.modal-card','.drop-placeholder','.notice{position:fixed','.diff-dialog{width:min(1800px,94vw)','.diff-viewers.overlay-mode','@media(max-width:1100px)','.thumbnail-grid','.page-command-bar.has-selection','.page-thumb-editor','.page-thumb-card.editing']:
    if needle not in css: raise SystemExit(f'css feature not found: {needle}')
if 'font-weight:900' in css or 'radial-gradient' in css: raise SystemExit('old visual style remains')

# V4.5 Excel screen width and Explorer-matching timestamp checks.
for needle in ['availableFilesScannedAt','formatDateTimeWithSeconds','formatSubmissionFileModifiedAt','modifiedAtDisplay','setFileSelectionSummary','title="原稿ファイルの最終保存日時"']:
    if needle not in appjs: raise SystemExit(f'v4.5 Excel UI feature not found: {needle}')
if '<th class="numeric-col">シート数</th>' in appjs or '<th class="numeric-col">ページ数</th>' in appjs:
    raise SystemExit('Excel tables must not show sheet/page count columns')
if 'formatDateTimeWithSeconds(f.modifiedAt' in appjs:
    raise SystemExit('submission file modified time must display to the minute')

if 'const byEpoch = formatLocalDateTimeMinuteFromUnixMs(file?.modifiedAtUnixMs)' in appjs:
    raise SystemExit('submission file time must not be converted again in the browser')
if 'id="select-render-needed-btn"' in html or 'selectRenderNeededWorkbooks' in appjs:
    raise SystemExit('needed PDF targets must not require a separate selection action')
for needle in ['`必要な${info.count}件の変換PDFを作成`', "btn.disabled = info.mode === 'none'", 'await renderUpdated($(\'render-selected-btn\'))']:
    if needle not in appjs: raise SystemExit(f'one-click needed PDF rendering missing: {needle}')
if appjs.count("selectedWorkbooks.clear();\n    lastWorkbookRangeAnchor = '';\n    renderAll();") < 1 or "await refresh();selectedWorkbooks.clear();lastWorkbookRangeAnchor='';renderAll();" not in appjs:
    raise SystemExit('workbook selection must reset after update scan and PDF rendering')
for needle in ['unregistered-card','registered-card','workbook-context-bar','render-target-hint']:
    if needle not in html: raise SystemExit(f'v4.3 Excel HTML feature not found: {needle}')
for needle in ['font-family:"Segoe UI","BIZ UDPGothic","BIZ UDPゴシック"','grid-template-columns:minmax(460px,.82fr) minmax(650px,1.18fr)','font-size:16px','.workbook-table{width:100%;min-width:0']:
    if needle not in css: raise SystemExit(f'v4.3 readability CSS feature not found: {needle}')
for needle in ['.workbook-table .pdf-status-col{width:320px}', '.excel-grid{grid-template-columns:minmax(340px,.65fr) minmax(720px,1.35fr)', '.excel-grid>.card+.card{margin-top:0}', '.excel-grid .file-name-cell strong{overflow:visible;text-overflow:clip;white-space:normal;overflow-wrap:anywhere', '.pdf-status-stack{display:grid;gap:5px}', '@media(max-width:1320px){.excel-grid{grid-template-columns:1fr}']:
    if needle not in css: raise SystemExit(f'change column readability CSS missing: {needle}')
for needle in ['<div class="file-meta"', '更新日時：', '変換PDF作成日時：', '<th class="pdf-status-col">変換PDFの状態</th>', '原稿更新あり', 'PDFを再作成して更新を反映してください', '見た目変更 ${affected}', 'class="pdf-comparison-link"', 'data-open-history', 'sourceTypeBadge', 'data-source-owner', 'data-source-required']:
    if needle not in appjs: raise SystemExit(f'file identity UX feature missing: {needle}')
for needle in ['class="source-advanced-settings"', '<summary>原稿の扱いを変更</summary>', '最終PDFに必須', '更新で増えたページの追加先', '既に並べたページは移動しません', 'data-source-settings-details', 'activePackDefaultTargetLabel()', '一式の既定：']:
    if needle not in appjs: raise SystemExit(f'advanced source settings disclosure missing: {needle}')
if 'ひな形に従う' in appjs:
    raise SystemExit('source settings must explain the default instead of saying template-driven')
if "closest('input,button,a,select,textarea,label,summary,details')" not in appjs:
    raise SystemExit('opening advanced source settings must not toggle workbook selection')
for needle in ['class="pack-switcher"', 'aria-label="今回まとめる一式"', '原稿を登録・PDF化', '未登録原稿', '登録済み原稿']:
    if needle not in html: raise SystemExit(f'generalized document pack UI feature missing: {needle}')
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
v2_general=json.loads((root/'app/tools/fixtures/structure-v2-generalized-ja.json').read_text(encoding='utf-8'))
v3_expected=json.loads((root/'app/tools/fixtures/structure-v3-migration-expected.json').read_text(encoding='utf-8'))
if v2_general['schemaVersion'] != 2 or v3_expected['schemaVersion'] != 3:
    raise SystemExit('schema v2-to-v3 fixture versions are invalid')
if set(v3_expected['packIds']) != {'pack_ecm','pack_bod','pack_dmm'}:
    raise SystemExit('schema v3 built-in pack fixture is invalid')
for needed in ['function Sync-StructureV3FromLegacy','function ConvertTo-StructureV3','function ConvertTo-V4StructureCompatibilityView',
               'structure.json.v2.bak',"Set-NoteProperty $Structure 'schemaVersion' 3"]:
    if needed not in server: raise SystemExit(f'schema v3 compatibility feature missing: {needed}')
for needed in ['function Get-SourceAdapterDescriptor','function Test-SourceCandidate','function Get-SourceCandidates','function Inspect-Source',
               'function Register-Source','function Register-SourcesBatch','function Render-Source','function Get-SourceChangeSummary',
               "adapterId = 'excel-com-v1'",'Render-Source $language $id $excel $true']:
    if needed not in server: raise SystemExit(f'source adapter feature missing: {needed}')
powershell = shutil.which('powershell.exe') or shutil.which('powershell')
if powershell:
    schema_check=subprocess.run(
        [powershell,'-NoProfile','-ExecutionPolicy','Bypass','-File',str(root/'app/tools/schema-v3-selfcheck.ps1')],
        cwd=root, text=True, encoding='utf-8', errors='replace', stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=90
    )
    if schema_check.returncode != 0 or 'schema-v3 selfcheck ok' not in schema_check.stdout:
        raise SystemExit('schema v3 PowerShell selfcheck failed:\n'+schema_check.stdout)
    adapter_check=subprocess.run(
        [powershell,'-NoProfile','-ExecutionPolicy','Bypass','-File',str(root/'app/tools/source-adapter-selfcheck.ps1')],
        cwd=root, text=True, encoding='utf-8', errors='replace', stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=90
    )
    if adapter_check.returncode != 0 or 'source-adapter selfcheck ok' not in adapter_check.stdout:
        raise SystemExit('source adapter PowerShell selfcheck failed:\n'+adapter_check.stdout)
    operations_check=subprocess.run(
        [powershell,'-NoProfile','-ExecutionPolicy','Bypass','-File',str(root/'app/tools/operational-readiness-selfcheck.ps1')],
        cwd=root, text=True, encoding='utf-8', errors='replace', stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=90
    )
    if operations_check.returncode != 0 or 'operational readiness selfcheck ok' not in operations_check.stdout:
        raise SystemExit('operational readiness PowerShell selfcheck failed:\n'+operations_check.stdout)

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
    for cls in ['ReportPdfComposer.class','ReportPdfComposer$ContentSpec.class','BatchPdfSplitter.class','BatchPdfSplitter$PdfBox.class']:
        if cls not in set(zf.namelist()): raise SystemExit(f'jar class missing: {cls}')

root_files={x.name for x in root.iterdir() if x.is_file()}
extra=sorted(root_files-{'資料をPDFにまとめる.cmd','共有フォルダー用フォルダー作成.cmd','README.md','THIRD_PARTY_NOTICES.md','CHANGELOG_V4.md','CHANGELOG_V5.md','.gitignore'})
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

# History/diff is enabled by default; the old workspace policy gate must not return.
for fn in ['function Test-InputHistoryEnabled', 'function Test-SourceRetentionEnabled']:
    if fn not in server: raise SystemExit('missing: ' + fn)
if 'function Get-WorkspacePolicy' in server or '$Script:WorkspacePolicyCache' in server:
    raise SystemExit('legacy policy.json gate/cache must be removed')
_history_enabled = server.split('function Test-InputHistoryEnabled', 1)[1].split('\nfunction ', 1)[0]
if 'return $true' not in _history_enabled:
    raise SystemExit('input history and diff must be enabled without policy.json')
_source_retention = server.split('function Test-SourceRetentionEnabled', 1)[1].split('\nfunction ', 1)[0]
if '$Script:LocalProjectsRoot' not in _source_retention or 'StartsWith($localRoot, [StringComparison]::OrdinalIgnoreCase)' not in _source_retention:
    raise SystemExit('source retention must remain limited to the per-user local project root')
if (root/'docs/POLICY_SAMPLE.json').exists():
    raise SystemExit('obsolete POLICY_SAMPLE.json must not be distributed')
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

# Final PDF output always uses the transactional history/archive path.
_bfp = server.split('function Build-FinalPdf(', 1)[1].split('\nfunction ', 1)[0]
if 'Invoke-FinalBuildTransaction' not in _bfp or 'Build-FinalPdfLegacy' in _bfp:
    raise SystemExit('Build-FinalPdf must always use the transactional path')
_build_all_route = server.split("'/api/final/build-all'", 1)[1].split("'/api/", 1)[0]
if 'Invoke-FinalBuildAllLegacy' in _build_all_route:
    raise SystemExit('build-all route must not fall back to the legacy path')

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
for needed in ['function addedSheetSet', 'const added = addedNames.size', 'affected=changed+added+removed', "badge('追加','attention')", 'addedSheetSet(p.workbookId).has(name)']:
    if needed not in appjs:
        raise SystemExit(f'added sheets must appear in the change column/filter: {needed}')

# Detailed visual diff is rendered from the source PDFs in the browser.
for needed in ['function Get-DiffDetailContext', 'function Serve-HistoryContentPdf',
               '$Script:VisualHashDpi = 120',
               'Resolve-ContentPdfSheetPathExact', 'function Get-ContentPdfSheetIndex',
               'function Write-FileResponse',
               "Write-FileResponse $Context 200 $full 'application/pdf' $false 'private, max-age=31536000, immutable'"]:
    if needed not in server:
        raise SystemExit(f'browser PDF comparison server support missing: {needed}')
for needed in ['diffDetailRequestPath', 'diffHistoryPdfParams', 'fetchDiffPdfDocument',
               'fromSnapshotId:from', 'toSnapshotId:to',
               '選んだ2版を比較', '比較元・比較先を入れ替え',
               'renderSnapshotHistorySelectionHint', 'visualCompareReady', 'unavailableReason']:
    if needed not in appjs:
        raise SystemExit(f'historical browser PDF comparison UI missing: {needed}')
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
for needed in ['$Script:LatestComparisonCache', '$cacheKey', '$snap', '$ver']:
    if needed not in _glc:
        raise SystemExit(f'latest automatic comparison cache missing: {needed}')
_state_payload = server.split('function Get-StatePayload', 1)[1].split('\nfunction ', 1)[0]
if 'Get-WorkbookChangeSummary' in _state_payload:
    raise SystemExit('initial state must not read one comparison file per workbook')
for needed in ["Get-DataProperty $w 'latestComparisonSummary'", 'Get-AutoStateSummary $Language -Fast',
               'Get-AllFinalReadiness $structure $Language $false']:
    if needed not in _state_payload:
        raise SystemExit(f'fast initial state path missing: {needed}')


# Background scans are throttled and history data is not requested twice per render.
for needed in ['300000', 'lastUpdateScanAt', "setTimeout(() => scanUpdatesSilently({withFiles:true}), 30000)"]:
    if needed not in appjs:
        raise SystemExit(f'shared-folder scan throttle missing: {needed}')
_render_all = appjs.split('function renderAll', 1)[1].split('\nfunction ', 1)[0]
if "if(activeView==='history')loadHistoryPanels()" in _render_all:
    raise SystemExit('renderAll must not duplicate the history request already started by setActiveView')
if "const previewUrl=apiUrl('/api/history/content-pdf'" not in appjs:
    raise SystemExit('history PDF preview must use direct browser streaming')

# Obsolete timeline/layout/archive UI no longer creates shared-folder files.
_history_writer = server.split('function Write-HistoryEvent', 1)[1].split('\nfunction ', 1)[0]
if 'return' not in _history_writer or 'Write-JsonFile' in _history_writer:
    raise SystemExit('obsolete history timeline must not write event files')
_transaction = server.split('function Invoke-FinalBuildTransaction', 1)[1].split('\nfunction ', 1)[0]
if 'New-FinalArchive $Language $cat $v' in _transaction or "Save-LayoutSnapshot $Language $cat 'final-build'" in _transaction:
    raise SystemExit('final build must not create removed archive/layout-history artifacts')

# Existing automatic comparisons use their deterministic analysis file and skip redundant SMB validation.
_glc = server.split('function Get-LatestComparison([string]', 1)[1].split('\nfunction ', 1)[0]
if "Join-Path $recordDir 'comparison-analysis.json'" not in _glc or _glc.index('comparison-analysis.json') > _glc.index("Join-Path $recordDir 'comparisons'"):
    raise SystemExit('existing automatic comparison must use the direct analysis file before directory enumeration')
_diff_context = server.split('function Get-DiffDetailContext', 1)[1].split('\nfunction ', 1)[0]
_auto_context = _diff_context.split("$status = [string]", 1)[1]
for forbidden in ['currentPdfDir', 'baselinePdfDir', 'Get-DiffSnapshotDate $Language $safeWorkbookId $currentSnapshotId',
                  'Get-DiffSnapshotDate $Language $safeWorkbookId $baselineSnapshotId']:
    if forbidden in _auto_context:
        raise SystemExit(f'automatic comparison open still performs redundant SMB validation: {forbidden}')

# Shared-folder hot paths use local immutable caches and mutation responses.
for needed in ['$Script:SnapshotManifestCache', '$Script:VisualHashCache', '$Script:SnapshotSummaryCache',
               '$Script:LocalRuntimeCacheRoot', 'function Get-LocalSnapshotSummaryCachePath',
               'function Get-LocalDiffDetailCache', 'function Get-LocalDiffDetailCacheIdentity',
               '$BaselineVersionId', '$CurrentVersionId', '$Script:DiffDetailAlgorithmVersion',
               "foreach ($field in @('scope','baselineSnapshotId','baselineVersionId','currentSnapshotId','currentVersionId'))",
               'function Publish-LatestComparisonCaches',
               "source = 'local-cache'"]:
    if needed not in server:
        raise SystemExit(f'local shared-folder cache missing: {needed}')
_snapshot_ids = server.split('function Get-SnapshotIds', 1)[1].split('\nfunction ', 1)[0]
if "Where-Object { Test-Path" in _snapshot_ids:
    raise SystemExit('snapshot listing must not stat every manifest')
for needed in ["-Filter 'manifest.json'", "-Depth 1", 'function Clear-SnapshotRuntimeCaches']:
    if needed not in (_snapshot_ids + server):
        raise SystemExit(f'completed snapshot/cache integrity safeguard missing: {needed}')
_pdf_index = server.split('function Get-ContentPdfSheetIndex', 1)[1].split('\nfunction ', 1)[0]
if _pdf_index.index('$cached =') > _pdf_index.index('Test-DirectoryExistsCompat $dir'):
    raise SystemExit('immutable content PDF index must be checked before SMB directory access')
_preview = server.split('function Serve-ContentPdfByValues', 1)[1].split('\nfunction ', 1)[0]
if "Write-FileResponse $Context 200 $full 'application/pdf'" not in _preview or 'ReadAllBytes($full)' in _preview:
    raise SystemExit('page preview must stream the PDF instead of buffering it')
for route_marker in [
    "Save-LayoutSnapshot $language (Require-WorkbookCategory ([string]$body.category)) 'reorder'",
    "Save-LayoutSnapshot $language (Require-WorkbookCategory ([string]$body.category)) 'sort-by-sheet'",
    "Save-LayoutSnapshot $language (Require-WorkbookCategory ([string]$body.category)) 'page-update'",
]:
    if route_marker in server:
        raise SystemExit(f'page mutation still writes obsolete layout history: {route_marker}')
for needed in ['function applyPageMutationResult', 'applyPageMutationResult(response)',
               "const previewUrl=apiUrl('/api/file'", 'function loadFinalReadiness']:
    if needed not in appjs:
        raise SystemExit(f'client-side fast mutation/preview path missing: {needed}')
_history_loader = appjs.split('function loadHistoryPanels', 1)[1].split('\n}', 1)[0]
if 'loadSnapshotHistory(' not in _history_loader or any(x in _history_loader for x in ['loadHistoryTimeline', 'loadLayoutSnapshots', 'loadFinalArchives']):
    raise SystemExit('history view must load only the selected workbook versions')
for obsolete in ['id="history-timeline"', 'id="layout-history"', 'id="final-archives"']:
    if obsolete in html:
        raise SystemExit(f'obsolete history panel remains: {obsolete}')

# The scheduler child must be stopped explicitly, not only by parent-PID polling.
_srv = server.split('function Start-LocalTcpServer', 1)[1].split('\nfunction ', 1)[0]
if 'Stop-AutoSchedulerProcess' not in _srv:
    raise SystemExit('Start-LocalTcpServer must stop the auto-scheduler child on exit')
if _srv.index('try {') > _srv.index('$tcp.Start()'):
    raise SystemExit('listener startup must be covered by the scheduler-stop finally block')
if _srv.index('$tcp.Start()') > _srv.index('Start-AutoSchedulerProcess'):
    raise SystemExit('listener must start before the scheduler child to keep startup responsive')

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
if 'Invoke-InputHistoryCleanup $language' not in _child:
    raise SystemExit('input-history cleanup must run in the deferred scheduler process')
_startup_recovery = server.split('function Invoke-StartupRecovery', 1)[1].split('\nfunction ', 1)[0]
for forbidden in ['Clear-ExpiredLeases', 'Clear-StaleEphemeralCopies', 'Recover-AutoStates', 'Invoke-InputHistoryCleanup']:
    if forbidden in _startup_recovery:
        raise SystemExit(f'UI startup must not repeat deferred recovery work: {forbidden}')
if 'Ensure-Package $startupPaths -Languages @((Get-EffectiveLanguage))' not in server:
    raise SystemExit('normal startup must initialize only the active language package')
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

# Release ZIPs must set the UTF-8 name flag so Japanese filenames remain usable.
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

# The one workspace may restart its own stale UI process, never worker processes.
launch = (root/'app/launch.ps1').read_text(encoding='utf-8-sig')
_stale = launch.split('function Stop-StaleUiServerProcesses', 1)[1].split('\nfunction ', 1)[0]
for needed in ['-autoschedulerpath', '-renderjobpath', '-diffjobpath']:
    if needed not in _stale:
        raise SystemExit(f'stale-server cleanup must exclude workers ({needed})')
if 'Stop-StaleUiServerProcesses $server' not in launch or '$TargetMode' in _stale:
    raise SystemExit('stale-server cleanup must target the single workspace')
for needed in ['System.Threading.Mutex', '$launchMutexOwned', 'duplicate invocation will exit without opening a tab']:
    if needed not in launch:
        raise SystemExit(f'launcher single-instance guard missing: {needed}')
for needed in ['ReportBinderを準備しています', 'アプリを起動しています。', '通常より時間がかかっています。']:
    if needed not in launch:
        raise SystemExit(f'startup wait page guidance missing: {needed}')
if '準備ができたら画面を開く' in launch or 'class="actions"' in launch:
    raise SystemExit('startup wait page must not ask the user to open the app manually')
if '$readyWaitMs -ge 12000' in launch:
    raise SystemExit('delayed direct browser fallback can race the wait-page redirect and open a duplicate tab')

# Existing structure.json must never be converted to an empty workspace after a
# transient shared-folder read or parse failure.
for needed in ['function Test-StructureDocument', 'structure.json.last-good',
               'structure.json は上書きしていません', '不完全な登録情報の保存を拒否しました']:
    if needed not in server:
        raise SystemExit(f'structure fail-closed protection missing: {needed}')
_read_structure = server.split('function Read-StructureUnlocked', 1)[1].split('\nfunction ', 1)[0]
if 'Read-JsonFile $path (New-EmptyStructure' in _read_structure:
    raise SystemExit('existing structure read failures must not fall back to New-EmptyStructure')

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
if os.name == 'nt':
    # The release runtime is Windows PowerShell 5.1. Parse every changed PowerShell
    # entry point and compile the C# diff engine instead of relying on text checks.
    ps_paths = [
        root/'app/server.ps1',
        root/'app/launch.ps1',
        root/'app/tools/package-release.ps1',
        root/'app/tools/diff-image-batch.ps1',
    ]
    quoted_paths = ','.join("'" + str(path).replace("'", "''") + "'" for path in ps_paths)
    engine_path = str(root/'app/tools/DiffImageEngine.cs').replace("'", "''")
    runtime_check = (
        "$failed=$false;"
        f"foreach($path in @({quoted_paths})){{"
        "$tokens=$null;$errors=$null;"
        "[void][System.Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors);"
        "if($errors.Count -gt 0){$errors|ForEach-Object{Write-Error ($path+': '+$_.Message)};$failed=$true}"
        "};"
        "if($failed){exit 1};"
        f"Add-Type -Path '{engine_path}' -ReferencedAssemblies @('System.Drawing')"
    )
    checked = subprocess.run(
        ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', runtime_check],
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace',
    )
    if checked.returncode:
        raise SystemExit('Windows PowerShell/C# compile check failed:\n' + checked.stdout + checked.stderr)
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
if '$Script:DiffDetailAlgorithmVersion = 23' not in server:
    raise SystemExit('ambiguous physical page alignment changes must bump DiffDetailAlgorithmVersion to 23')
diff_batch = (root/'app/tools/diff-image-batch.ps1').read_text(encoding='utf-8-sig')
for needed in ['function Get-DiffRasterSignature', '[double]$gapCost = 0.0028', '[double]$substitutionCap = 0.005',
               '$alignmentBaseline', "'perceptual-raster-sequence-ambiguous'", 'beforePageNumber', 'afterPageNumber',
               'comparisonKind', 'mappingAmbiguous', 'mappingMessage']:
    if needed not in diff_batch:
        raise SystemExit(f'physical PDF page alignment missing: {needed}')
for needed in ['beforePageNumber = Get-IntDataProperty', 'afterPageNumber = Get-IntDataProperty', "comparisonKind = [string]", 'mappingAmbiguous = [bool]', 'mappingMessage = [string]']:
    if needed not in server:
        raise SystemExit(f'physical PDF page mapping persistence missing: {needed}')
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
if "'private, max-age=31536000, immutable'" not in _serve_diff:
    raise SystemExit('immutable diff page assets must opt into browser caching')
_bytes_response = server.split('function Write-BytesResponse', 1)[1].split('\nfunction ', 1)[0]
if 'Write-TcpResponse $Context $Status $Bytes $ContentType $AllowCors $CacheControl' not in _bytes_response:
    raise SystemExit('diff asset cache policy must reach the TcpListener response path')
appjs = (root/'app/web/app.js').read_text(encoding='utf-8-sig')
diff_worker = (root/'app/web/diff-worker.js').read_text(encoding='utf-8-sig')
for needed in ["canvas.style.visibility='hidden'", 'beginDiffBrowserRender', 'task.cancel()', 'DIFF_PAGE_CACHE_LIMIT = 6', 'setDiffBrowserProgress', 'PDFを読み込んでいます…']:
    if needed not in appjs:
        raise SystemExit(f'browser page-switch responsiveness missing: {needed}')
for needed in ['getImageData(0,0,width,height)', 'postMessage({id,width,height', 'diffAnalysisPending']:
    if needed not in appjs:
        raise SystemExit(f'non-blocking browser diff analysis missing: {needed}')
for needed in ["const geometry=region?.[side]||region", 'Number(geometry.x||0)', 'Number(geometry.width||0)']:
    if needed not in appjs:
        raise SystemExit(f'side-specific diff region rendering missing: {needed}')
for needed in ['mappedPage?.comparisonKind', 'mappedPage.beforePageNumber', 'mappedPage.afterPageNumber', 'mappedPage?.mappingMessage',
               'buildDocumentTextRowStructureDiffResult(documentTextPair[0],documentTextPair[1],width,height,beforePageNumber,afterPageNumber)']:
    if needed not in appjs:
        raise SystemExit(f'browser physical page mapping missing: {needed}')

# Third-party installation must stay reproducible and fail closed.
installer = (root/'app/tools/install-thirdparty.ps1').read_text(encoding='utf-8-sig')
for needed in ['2.0.37', 'Verify-HashFile', 'Verify-SriSha512', 'dist.integrity', 'binary.package.checksum', 'PreferSystemJava', 'verify-thirdparty.ps1']:
    if needed not in installer: raise SystemExit(f'third-party installer hardening missing: {needed}')
verifier = (root/'app/tools/verify-thirdparty.ps1').read_text(encoding='utf-8-sig')
for needed in ['RequirePortableJava', 'SHA-512 verified', 'verified npm package tarball', 'PdfPageAnalyzer.class']:
    if needed not in verifier: raise SystemExit(f'third-party verifier check missing: {needed}')
package_release = (root/'app/tools/package-release.ps1').read_text(encoding='utf-8-sig')
for needed in ['Invoke-ThirdPartyCheck $true', 'Assert-StagedDependencies', 'JAVA_VERSION.txt',
               'function Remove-ReleaseDevelopmentFiles', 'function Write-ReleaseManifest',
               "'release-manifest.json'", "'ReportPdfComposer$ContentSpec.class'",
               "$excludedTopLevelNames", "'tmp'"]:
    if needed not in package_release: raise SystemExit(f'release dependency gate missing: {needed}')
for needed in ['[switch]$SharedFolderOnly', 'function Remove-SharedFolderDevelopmentFiles', 'function Assert-SharedFolderLayout',
               'ReportBinder_共有フォルダー用_', "'資料をPDFにまとめる.cmd'", 'Start-Process']:
    if needed not in package_release: raise SystemExit(f'shared-folder release feature missing: {needed}')
_shared_cleanup = package_release.split('function Remove-SharedFolderDevelopmentFiles', 1)[1].split('\nfunction ', 1)[0]
for needed in ["'app\\lib\\pdfbox\\src'", "'app\\tools\\fixtures'", "'app\\tools\\selfcheck.py'",
               "'app\\tools\\package-release.ps1'", "'共有フォルダー用フォルダー作成.cmd'",
               "'app\\tools\\scale-benchmark.ps1'", "'app\\tools\\ci-selfcheck.ps1'", "'app\\tools\\history-logic-selfcheck.ps1'", "'app\\tools\\ci-requirements.txt'", "'docs\\benchmarks'"]:
    if needed not in _shared_cleanup:
        raise SystemExit(f'shared-folder cleanup omission: {needed}')
if "-IncludeJava $true" not in package_release:
    raise SystemExit('shared-folder release must always include portable Java')
_shared_release_entry = package_release.split('if ($SharedFolderOnly) {', 1)[1].split('\n}\n\nInvoke-SelfCheck', 1)[0]
if 'Invoke-SelfCheck' in _shared_release_entry:
    raise SystemExit('shared-folder creation must not run the development-tree selfcheck')
if '\nInvoke-SelfCheck\n# PDFBox and PDF.js' not in package_release:
    raise SystemExit('ZIP releases must keep the fail-closed development-tree selfcheck')
for needed in ["GetFolderPath('LocalApplicationData')", "'ReportBinder\\release'",
               'function Copy-DirectoryContents', 'function Publish-SharedFolderStage',
               "'_作成中.txt'", 'ReportBinderShared_', 'Publish-SharedFolderStage -StageRoot $stage']:
    if needed not in package_release:
        raise SystemExit(f'OneDrive-safe shared release feature missing: {needed}')
_publish_shared = package_release.split('function Publish-SharedFolderStage', 1)[1].split('\nfunction ', 1)[0]
for needed in ['Move-Item -LiteralPath $StageRoot', 'catch {', 'Copy-DirectoryContents',
               'Assert-StagedDependencies -StageRoot $TargetRoot', 'Assert-SharedFolderLayout $TargetRoot',
               'Remove-Item -LiteralPath $TargetRoot -Recurse -Force']:
    if needed not in _publish_shared:
        raise SystemExit(f'shared release publish fallback missing: {needed}')
_new_shared = package_release.split('function New-SharedFolderRelease', 1)[1].split('\nfunction ', 1)[0]
if 'Move-Item -LiteralPath $stage' in _new_shared:
    raise SystemExit('shared release must publish through the OneDrive-safe fallback helper')
if 'app/thirdparty-cache/' not in (root/'.gitignore').read_text(encoding='utf-8'):
    raise SystemExit('third-party cache must be ignored by git')

# V5.1: history comparison must be a first-class destination and list actions must be visible before the list.
for needed in ['data-view-nav="history"', 'data-view-panel="history"', '<h2>変更履歴・比較</h2>', '比較する原稿', 'class="list-action-bar contextual-action-bar"']:
    if needed not in html: raise SystemExit(f'history/list action UX missing: {needed}')
if html.index('id="register-selected-btn"') > html.index('id="file-list"'):
    raise SystemExit('unregistered Excel actions must appear before the file list')
if "history: '変更履歴・比較'" not in appjs or "activeView === 'history'" not in appjs:
    raise SystemExit('history navigation is not wired')

# Final PDF names must use the selected category, not a category inherited from an Excel file name.
for needed in ['function Get-CategoryProjectId', 'Resolve-OutputFileNamePattern', '$outName=[string]$snapshotBefore.outputFileName', '$outName = [string]$snapshots[$v].outputFileName']:
    if needed not in server: raise SystemExit(f'category-aware final filename missing: {needed}')
if '$outName=Get-OutputFileName $Volume $projectId;' in server:
    raise SystemExit('legacy final output still omits category')

# V5.2/V5.4: stable visual hashes plus browser-side PDF rendering and diff analysis.
for needed in [
    '$Script:VisualHashProfileVersion = 3',
    '$Script:VisualHashDpi = 120',
    'function Test-SheetVisualEquivalent',
    '$Script:DiffDetailAlgorithmVersion = 23',
]:
    if needed not in server:
        raise SystemExit(f'comparison tolerance/browser setting missing: {needed}')
diff_engine = (root/'app/tools/DiffImageEngine.cs').read_text(encoding='utf-8-sig')
for needed in [
    'MaximumModifiedLabelsPerPage = 12',
    'FindBestOffset',
    'MergeNearbyRegions',
    'SaveBaseImage',
    'sparseFullPage',
    'changedIntegral',
    'Parallel.For',
    'MeanAbsoluteByteDistance',
    'FindRoot',
    'MaximumRegionsPerPage * 3',
]:
    if needed not in diff_engine:
        raise SystemExit(f'diff-region noise control missing: {needed}')
analyzer = (root/'app/lib/pdfbox/src/PdfPageAnalyzer.java').read_text(encoding='utf-8-sig')
for needed in ['ANALYZER_VERSION = 2', 'pageHashes', 'normalizedPixelHash', 'pagePerceptualHashes', 'perceptualHash']:
    if needed not in analyzer:
        raise SystemExit(f'perceptual visual hash missing: {needed}')
_visual_equivalence = server.split('function Test-SheetVisualEquivalent', 1)[1].split('\nfunction ', 1)[0]
for needed in ["Get-DataProperty $Before 'pageHashes'", "Get-DataProperty $After 'pageHashes'",
               'Legacy records without normalized page hashes retain the perceptual fallback']:
    if needed not in _visual_equivalence:
        raise SystemExit(f'exact page-hash comparison priority missing: {needed}')
if _visual_equivalence.index("Get-DataProperty $Before 'pageHashes'") > _visual_equivalence.index('Get-HexHammingRatio'):
    raise SystemExit('exact page hashes must be authoritative before the legacy perceptual fallback')

# V5.3: all changed sheets share one bounded-concurrency Java raster job,
# and region decoration is drawn from JSON instead of four PNG layers per page.
batch_script = (root/'app/tools/diff-image-batch.ps1').read_text(encoding='utf-8-sig')
for needed in ['PdfBatchRasterizer', 'ProcessorCount', 'ReportBinderDiffEngine', 'items = @($results)', 'rasterMs', 'analysisMs']:
    if needed not in batch_script:
        raise SystemExit(f'batched diff rasterization missing: {needed}')
for needed in ['Get-CachedRasterPages', 'cacheHitSides', '$needsRaster', 'beforePersistent', 'afterPersistent']:
    if needed not in batch_script:
        raise SystemExit(f'persistent render-raster reuse missing: {needed}')
for needed in ['ReportBinderDiffBatchPageRequest', 'ComparePages(', 'analysisThreads', 'analyzedPages']:
    if needed not in batch_script:
        raise SystemExit(f'parallel diff analysis missing: {needed}')
for needed in ['unchangedPageNumbers', "if ($pageKind -eq 'unchanged')", "'exact-raster-sequence'", 'unchangedPagesSkipped', 'fullyAnalyzedPages']:
    if needed not in batch_script:
        raise SystemExit(f'exact-page diff fast path missing from batch script: {needed}')
for needed in ['BuildSimplePage', 'forcedKind == "unchanged"', 'TryGetImageSize']:
    if needed not in diff_engine:
        raise SystemExit(f'non-analysis diff fast path missing from engine: {needed}')
_diff_skeleton = server.split('function New-DiffDetailSkeleton', 1)[1].split('\nfunction ', 1)[0]
for needed in ['pageHashes', 'Normalize-FileHash', 'unchangedPageNumbers', "status = 'ready'"]:
    if needed not in _diff_skeleton:
        raise SystemExit(f'exact page hash/browser propagation missing: {needed}')
_diff_core = server.split('function Invoke-DiffDetailJobCore', 1)[1].split('\nfunction ', 1)[0]
for needed in ['$preferredIndex', '$workIndexes += $preferredIndex', '1ジョブにつき1シートだけ処理する']:
    if needed not in _diff_core:
        raise SystemExit(f'one-sheet lazy generation missing: {needed}')
if "$sheetKind -ne 'unchanged' -or $sheetStatus -eq 'failed'" in _diff_core:
    raise SystemExit('diff job must not eagerly queue every changed sheet')
for needed in ['ensureDiffPdfJs', 'renderDiffPdfPage',
               'analyzeDiffCanvases', 'buildDiffBrowserPage', 'clearDiffBrowserResources',
               'fetchDiffDetailResponse', 'prefetchAutomaticDiffDetail',
               'disableRange:false', 'disableStream:false', 'clearDiffBrowserResources(true)']:
    if needed not in appjs:
        raise SystemExit(f'browser on-demand PDF comparison missing: {needed}')
_open_diff = appjs.split('async function openDiffDetail', 1)[1].split('\nfunction ', 1)[0]
if 'ensureDiffPdfJs' not in _open_diff:
    raise SystemExit('initial comparison open must warm PDF.js in parallel with metadata loading')
for forbidden in ['renderDiffRasterPage', 'renderDiffSourcePage', "'/api/history/render-page'"]:
    if forbidden in appjs:
        raise SystemExit(f'saved raster browser path must be removed: {forbidden}')
if "path -eq '/api/history/render-page'" in server or 'function Serve-HistoryRasterPage' in server:
    raise SystemExit('saved raster server endpoint must be removed')
_fetch_pdf = appjs.split('async function fetchDiffPdfDocument', 1)[1].split('\nfunction ', 1)[0]
for forbidden in ["cache:'no-store'", 'arrayBuffer()']:
    if forbidden in _fetch_pdf:
        raise SystemExit(f'PDF fetch must stream and remain browser-cacheable: {forbidden}')
for forbidden in ['prepareDiffDetail(', 'diffPrepareRequestBody(', 'diffAssetUrl(', 'setDiffImage(']:
    if forbidden in appjs:
        raise SystemExit(f'legacy server-image comparison remains in browser: {forbidden}')
batch_java = (root/'app/lib/pdfbox/src/PdfBatchRasterizer.java').read_text(encoding='utf-8-sig')
for needed in ['newFixedThreadPool', 'Math.min(4', 'renderSafely', 'ImageIO.write']:
    if needed not in batch_java:
        raise SystemExit(f'bounded parallel PDF rasterizer missing: {needed}')
for needed in ['function Invoke-DiffImageBatchGeneration', '$batchRequest', '$batchResultMap',
               '新旧PDFをまとめて画像化・解析しています']:
    if needed not in server:
        raise SystemExit(f'diff job does not use the batch path: {needed}')
for needed in ['function Get-RenderRasterSheetDir', 'rasterDirectory', 'beforeRasterDirectory', 'afterRasterDirectory']:
    if needed not in server:
        raise SystemExit(f'render-time raster cache wiring missing: {needed}')
for needed in ['copyBefore', 'copyAfter', 'reusedRasterPageAssets']:
    if needed not in batch_script:
        raise SystemExit(f'comparison PNG copy avoidance missing from batch script: {needed}')
for needed in ['public bool copyBefore', 'PrepareBaseAsset', 'return "render.png"', 'if (copyBefore)']:
    if needed not in diff_engine:
        raise SystemExit(f'comparison PNG copy avoidance missing from engine: {needed}')
_serve_diff = server.split('function Serve-DiffPage', 1)[1].split('\nfunction ', 1)[0]
for needed in ["fileName -eq 'render.png'", 'Get-RenderRasterSheetDir', "'page-{0:0000}.png'"]:
    if needed not in _serve_diff:
        raise SystemExit(f'direct render-raster serving missing: {needed}')
for needed in ['id="diff-before-regions"', 'id="diff-after-regions"', 'canvas id="diff-before-base"', 'canvas id="diff-after-base"', 'id="diff-export-summary"', 'id="final-preflight-summary"', 'id="final-preflight-list"', 'id="final-preflight-refresh"', 'id="app-exit-button"', 'id="change-source-folder-btn"', 'id="shutdown-screen"', 'id="main-content"', 'id="page-volume-tabs"', 'id="source-next-action"', 'id="app-loading-screen"', 'app.js?v=20260808_v144', 'style.css?v=20260808_v90']:
    if needed not in html:
        raise SystemExit(f'browser canvas diff markup/cache version missing: {needed}')
for needed in ['function renderDiffRegionLayer', "document.createElement('span')", 'diff-region-layer',
               'function buildDiffSummaryCsv', 'function exportDiffSummary', "text/csv;charset=utf-8"]:
    if needed not in appjs:
        raise SystemExit(f'browser region rendering missing: {needed}')
compare_page = diff_engine.split('public static ReportBinderDiffPage ComparePage', 1)[1]
if 'SaveLayerImages(' in compare_page:
    raise SystemExit('ComparePage must not encode full-page mask/overlay PNGs')
with zipfile.ZipFile(root/'app/lib/pdfbox/ReportPdfComposer.jar') as zf:
    if 'PdfBatchRasterizer.class' not in set(zf.namelist()):
        raise SystemExit('ReportPdfComposer.jar must contain PdfBatchRasterizer.class')

composer_source = (root/'app/lib/pdfbox/src/ReportPdfComposer.java').read_text(encoding='utf-8')
for needed in ['Map<String, PDDocument> sourceDocuments', 'out.save(outFile)', 'for (PDDocument src : sourceDocuments.values())']:
    if needed not in composer_source:
        raise SystemExit(f'composer source-document lifetime regression: {needed}')
if composer_source.index('out.save(outFile)') > composer_source.index('for (PDDocument src : sourceDocuments.values())'):
    raise SystemExit('composer must save the destination before closing imported source documents')

# 2026-07-31 comparison/final-output pipeline fixes ---------------------------
_render_job = server.split('function Invoke-RenderJobFromFile', 1)[1].split('\nfunction ', 1)[0]
if _render_job.index("$status.status = 'analyzing'") > _render_job.index('Invoke-PostRenderAnalysis'):
    raise SystemExit('render job must announce comparison analysis before running deferred analysis')
if _render_job.rindex("$status.status = 'completed'") < _render_job.rindex('Invoke-PostRenderAnalysis'):
    raise SystemExit('render job must not announce completion before comparison analysis is saved')
for needed in ['$renderLoopSucceeded', '比較情報を準備しています']:
    if needed not in _render_job:
        raise SystemExit(f'render comparison handoff is missing: {needed}')
_capture_input = server.split('function Capture-RenderInput', 1)[1].split('\nfunction ', 1)[0]
_render_workbook = server.split('function Render-Workbook', 1)[1].split('\nfunction ', 1)[0]
if 'verified = $true' not in _capture_input or '$inputHashVerified' not in _render_workbook:
    raise SystemExit('verified render input hash reuse is missing')
_post_analysis = server.split('function Invoke-PostRenderAnalysis', 1)[1].split('\nfunction ', 1)[0]
if '$pdf -and $name -and (Test-Path' in _post_analysis:
    raise SystemExit('post-render analysis must not repeat one SMB stat per rendered sheet')
_final_build = server.split('function Invoke-FinalBuildTransaction', 1)[1].split('\nfunction ', 1)[0]
for forbidden in ['Scan-Updates $Language', '$ready = Get-FinalBuildReadiness']:
    if forbidden in _final_build:
        raise SystemExit(f'final build still contains redundant shared-folder work: {forbidden}')
for needed in ['Assert-FinalBuildSourcesUnchanged $snapshots $targets',
               '$initialSnapshots = Update-StructureLocked',
               '$afterStructure = Get-Structure $Language',
               '[IO.Path]::GetTempPath()']:
    if needed not in _final_build:
        raise SystemExit(f'optimized final build path is missing: {needed}')
if 'function Assert-FinalBuildSourcesUnchanged' not in server:
    raise SystemExit('targeted final-build Excel metadata validation is missing')
if 'signal: options.signal' not in appjs or 'DIFF_DETAIL_TIMEOUT_MS = 15000' not in appjs:
    raise SystemExit('comparison metadata request timeout is missing')
_fetch_detail = appjs.split('async function fetchDiffDetailResponse', 1)[1].split('\nfunction ', 1)[0]
for needed in ['new AbortController()', 'attempt<2', 'diffDetailResponseCache.delete(key)']:
    if needed not in _fetch_detail:
        raise SystemExit(f'comparison metadata retry/recovery is missing: {needed}')
if 'app.js?v=20260808_v144' not in html:
    raise SystemExit('comparison request fix must bump the app cache version')

# 2026-07-31 history selection rendering fixes -------------------------------
_history_loader = appjs.split('async function loadSnapshotHistory', 1)[1].split('\nasync function ', 1)[0]
for needed in ['snapshotHistoryLoadSerial', 'fetchSnapshotHistoryList', 'requestSerial!==snapshotHistoryLoadSerial']:
    if needed not in _history_loader:
        raise SystemExit(f'history response race guard is missing: {needed}')
_fetch_snapshots = appjs.split('async function fetchSnapshotHistoryList', 1)[1].split('\nfunction ', 1)[0]
for needed in ['snapshotHistoryResponseCache', 'snapshotHistoryRequestCache', 'if(!promise)']:
    if needed not in _fetch_snapshots:
        raise SystemExit(f'history request cache/deduplication is missing: {needed}')
if 'if(!promise||force)' in _fetch_snapshots:
    raise SystemExit('forced history refresh must join an identical in-flight request')
_set_view = appjs.split('function setActiveView', 1)[1].split('\nfunction ', 1)[0]
for needed in ['viewChanged', '!historyPanelsInitialized', 'options.reloadPanels']:
    if needed not in _set_view:
        raise SystemExit(f'history tab side-effect guard is missing: {needed}')
_swap_part = _history_loader.split("const swap=$('snap-swap-btn')", 1)[1].split("const db=$('snap-diff-btn')", 1)[0]
if 'loadSnapshotHistory' in _swap_part or 'syncSnapshotHistorySelectionInputs' not in _swap_part:
    raise SystemExit('snapshot swap must update controls without refetching or rebuilding the table')
if 'function Update-SnapshotSummaryCacheEntry' not in server or 'function Get-SnapshotSummaryEntry' not in server:
    raise SystemExit('incremental snapshot summary cache update is missing')
_publish_cache = server.split('function Publish-LatestComparisonCaches', 1)[1].split('\nfunction ', 1)[0]
if 'Update-SnapshotSummaryCacheEntry' not in _publish_cache or 'Clear-SnapshotSummaryCache' in _publish_cache:
    raise SystemExit('render completion must keep the snapshot summary cache warm')
if 'app.js?v=20260808_v144' not in html:
    raise SystemExit('history rendering fix must bump the app cache version')

# 2026-08-01 per-user local runtime/project architecture ----------------------
runtime_info=json.loads((root/'app/runtime-version.json').read_text(encoding='utf-8'))
if not re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9._-]{0,63}', str(runtime_info.get('version',''))):
    raise SystemExit('runtime version must be a safe immutable directory name')
launch=(root/'app/launch.ps1').read_text(encoding='utf-8-sig')
for needed in ['[switch]$LocalRuntime', "'ReportBinder\\runtime\\versions'", 'installed.json',
               '$script:SharedAppRoot', 'runtimeVersion']:
    if needed not in launch:
        raise SystemExit(f'local runtime bootstrap is missing: {needed}')
for needed in ['startup-workspace-', 'ReportBinder-workspace.url', 'ReportBinder.Launch.$appRootKey.workspace']:
    if needed not in launch:
        raise SystemExit(f'single-workspace launcher is missing: {needed}')
if '-Mode $Mode' in launch or '[string]$Mode' in launch:
    raise SystemExit('launcher must not expose or forward a language mode')
_workspace_path = server.split('function Get-WorkspacePath',1)[1].split('\nfunction ',1)[0]
if "Join-Path $resolvedDataDir 'workspace'" not in _workspace_path:
    raise SystemExit('all document packs must use the single workspace directory')
if 'Join-Path $resolvedDataDir $Language' in _workspace_path:
    raise SystemExit('language-specific workspace storage must stay removed')
for obsolete in ['includeCover','includeToc','includeSectionDividers','template-include-toc']:
    if obsolete in (server + html + appjs):
        raise SystemExit(f'generated front-matter setting must stay removed: {obsolete}')
composer_source=(root/'app/lib/pdfbox/src/ReportPdfComposer.java').read_text(encoding='utf-8')
for obsolete in ['includeCover','includeToc','includeSectionDividers','drawGeneratedPage']:
    if obsolete in composer_source:
        raise SystemExit(f'composer must use registered source pages only: {obsolete}')
for needed in ['function Get-LocalProjectKey', '$Script:LocalProjectsRoot',
               "dataDir = (Join-Path $projectRoot 'data')",
               "outputDir = (Join-Path $projectRoot 'output')",
               'function Initialize-LocalProjectConfig',
               'function Initialize-CleanLocalProject',
               'function Move-IncompleteLocalProjectPath']:
    if needed not in server:
        raise SystemExit(f'per-user local project storage is missing: {needed}')
_clean_project = server.split('function Initialize-CleanLocalProject',1)[1].split('\nfunction ',1)[0]
for needed in ["initializationMode = 'clean'", 'legacySharedImport = $false',
               "reason = 'clean-start'", 'Move-IncompleteLocalProjectPath $targetData',
               'Move-IncompleteLocalProjectPath $targetOutput',
               "'Local\\ReportBinder.ProjectMigration.'", 'Write-JsonFile $marker']:
    if needed not in _clean_project:
        raise SystemExit(f'clean local project initialization missing: {needed}')
if 'Test-Path -LiteralPath $submissionDir' in _clean_project:
    raise SystemExit('clean local initialization must not stat the shared submission folder')
_quarantine_local = server.split('function Move-IncompleteLocalProjectPath',1)[1].split('\nfunction ',1)[0]
for needed in ["'.preclean-' + $Kind", 'Move-Item -LiteralPath $Path -Destination $backup',
               'Get-ChildItem -LiteralPath $Path -Force']:
    if needed not in _quarantine_local:
        raise SystemExit(f'incomplete local project quarantine missing: {needed}')
for forbidden in ['function Invoke-LegacyProjectMigration', 'function Copy-DirectoryContents',
                  'function New-LegacyMigrationStagePath', 'function Update-MigratedOutputPaths',
                  "Join-Path $submissionDir '_reportbinder'", "Join-Path $submissionDir '出力'"]:
    if forbidden in server:
        raise SystemExit(f'legacy shared-folder import must stay removed: {forbidden}')
if server.count('Initialize-CleanLocalProject') != 5:
    raise SystemExit('all startup and folder-selection paths must use clean local initialization')
_default_paths = server.split('function Get-DefaultChildPaths',1)[1].split('\nfunction ',1)[0]
for forbidden in ["Join-Path $trimmed '_reportbinder'", "Join-Path $trimmed '出力'"]:
    if forbidden in _default_paths:
        raise SystemExit(f'shared folder is still the default working storage: {forbidden}')
_publish = server.split('function Publish-DocumentPackPdfToShared',1)[1].split('\nfunction ',1)[0]
for needed in ['Get-SafePublishUserName', "Get-Date -Format 'MMdd_HHmmss'",
               'Get-SafePublishPackName',
               '"{0}_{1}_{2}" -f $stamp, $packName, $userName',
               '$fileName = [IO.Path]::GetFileName($sourceFull)', "'.publishing-'",
               'Move-Item -LiteralPath $stagingDir -Destination $publishDir',
               "$ready.displayState -ne 'built'"]:
    if needed not in _publish:
        raise SystemExit(f'safe shared publishing is missing: {needed}')
for forbidden in ["Join-Path ([string]$paths.submissionDir) '共有発行'",
                  "Get-Date -Format 'yyyyMMdd_HHmmss'",
                  '$fileName = "{0}_{1}_{2}.pdf"']:
    if forbidden in _publish:
        raise SystemExit(f'obsolete shared publish layout remains: {forbidden}')
for needed in ["'/api/final/publish'", "'/api/v2/outputs/publish'", 'data-final-publish',
               "publishable=available&&String(r.displayState||'')==='built'"]:
    if needed not in (server + html + appjs):
        raise SystemExit(f'shared publish UI/API is missing: {needed}')
for needed in ["'/api/v2/layout/snapshots'", "'/api/v2/layout/restore/preview'", "'/api/v2/layout/restore'",
               'function Get-LayoutScopeInfo', 'Test-WorkbookPack $wb[0] ([string]$scope.packId)',
               'layoutSnapshotId=$layoutSnapshotId', 'id="page-layout-revisions-btn"',
               'id="page-layout-revisions-modal"', "custom?'/api/v2/layout/restore/preview'"]:
    if needed not in (server + html + appjs):
        raise SystemExit(f'pack-scoped layout history is missing: {needed}')
_layout_restore = server.split('function Restore-LayoutSnapshot', 1)[1].split('\nfunction ', 1)[0]
for needed in ["Set-NoteProperty $p 'pageRange'", '$workbookIds.ContainsKey',
               'Get-NormalizedLayoutSnapshotPages', "Save-LayoutSnapshot $Language ([string]$lockedScope.packId) 'pre-restore' $st",
               'invalidVolumePageCount']:
    if needed not in _layout_restore:
        raise SystemExit(f'pack-scoped layout restore invariant is missing: {needed}')
for needed in ['function ConvertTo-NormalizedPackTemplate', 'function Save-PackTemplate',
               'function Remove-PackTemplate', 'function New-PackTemplateSnapshot',
               'function Get-PackEffectiveTemplate', "'/api/v2/pack-templates'",
               "code='required-source-missing'", "'required-source-stale'",
               "'required-source-failed'", 'Get-NewItemTargetId']:
    if needed not in server:
        raise SystemExit(f'user template workflow is missing: {needed}')
for needed in ['id="manage-pack-templates-btn"', 'id="template-manager-modal"',
               'id="template-new-destination"', 'data-source-default-target',
               'function openTemplateManager', 'function submitTemplateManager']:
    if needed not in (html + appjs):
        raise SystemExit(f'user template UI is missing: {needed}')

# 2026-08-03 two-axis comparison and quiet startup --------------------------
if str(runtime_info.get('version','')) != '2026.08.08.30':
    raise SystemExit('release quality gate must bump the immutable runtime version')
for needed in ['id="confirm-modal"', 'id="file-context-bar"', 'id="workbook-context-bar"', 'id="source-first-run"', 'data-progress-view="excel"', '提出用PDF']:
    if needed not in html:
        raise SystemExit(f'workflow UX markup is missing: {needed}')
for needed in ['function confirmAction', 'function isConfirmModalOpen', 'snapshot-timeline', 'data-snapshot-from', 'data-snapshot-to']:
    if needed not in appjs:
        raise SystemExit(f'workflow UX behavior is missing: {needed}')
stop_script=(root/'app/tools/stop-reportbinder.ps1').read_text(encoding='utf-8-sig')
for needed in ['ReportBinder\\runtime\\versions', "'app\\server.ps1'", "'app\\launch.ps1'"]:
    if needed not in stop_script:
        raise SystemExit(f'local runtime stop coverage is missing: {needed}')
if 'confirm(' in appjs or 'data-step-view=' in html:
    raise SystemExit('blocking browser confirms and duplicate step navigation must not return')
for needed in ['.contextual-action-bar.has-selection', '.confirm-dialog', '.snapshot-timeline-item']:
    if needed not in css:
        raise SystemExit(f'workflow UX styling is missing: {needed}')
for needed in ['id="submissionDir" type="text"', 'id="use-submission-path-btn"', '入力した場所を使う']:
    if needed not in html:
        raise SystemExit(f'submission folder path fallback is missing: {needed}')
for needed in ["bind('use-submission-path-btn'", "event.key==='Enter'", "setTimeout(()=>$('change-source-folder-btn')?.focus()"]:
    if needed not in appjs:
        raise SystemExit(f'submission folder keyboard/fallback wiring is missing: {needed}')
_folder_picker = server.split('function Select-FolderDialog', 1)[1].split('\nfunction ', 1)[0]
for needed in ['$proc.WaitForExit(60000)', '$proc.Kill()', '画面でパスを直接入力してください']:
    if needed not in _folder_picker:
        raise SystemExit(f'folder picker timeout recovery is missing: {needed}')

# 2026-08-08 first-time novice onboarding -----------------------------------
for needed in ['ReportBinderでできること', '複数の原稿を、必要な順番で1つのPDFにまとめる',
               'id="dashboard-onboarding"', '原稿を集める', '順番を整える', '1つのPDFにする',
               'フォルダを選べない場合', '保存先の詳しい説明']:
    if needed not in html:
        raise SystemExit(f'first-time purpose/onboarding copy is missing: {needed}')
for needed in ['はじめる：原稿が入ったフォルダを選ぶ', "action.textContent='はじめる'",
               'function continueFirstRunAfterFolderSelection', "setTimeout(()=>openPackEditor('create'),120)",
               "if(!wasRename){setActiveView('excel')", "button.disabled=!configured()&&!selected"]:
    if needed not in appjs:
        raise SystemExit(f'first-time single-path behavior is missing: {needed}')
if 'data-create-first-pack' in appjs:
    raise SystemExit('the initial dashboard must not show a competing pack-creation CTA')
_ensure_package = server.split('function Ensure-Package', 1)[1].split('\nfunction ', 1)[0]
for needed in ['$workspaceLanguages[0]', 'Read-StructureUnlocked $workspaceLanguage', 'Test-StructureDocument $verifiedStructure']:
    if needed not in _ensure_package:
        raise SystemExit(f'shared-workspace first-run verification is missing: {needed}')
if "foreach ($lang in @($Languages" in _ensure_package:
    raise SystemExit('the shared workspace must not be initialized once per display language')
for needed in ['choosePixelColumnMapping', 'mappedColumnX', 'refinePixelColumnMapping',
               "clearlyImproves(pixel.score,pixel.identityScore,.7)",
               "columnAlignment.split>=0", 'column-identity']:
    if needed not in diff_worker:
        raise SystemExit(f'two-axis diff alignment missing: {needed}')
for needed in ['MIN_LAYOUT_SCALE=.82', 'MAX_LAYOUT_SCALE=1.18',
               'for(let scaleStep=-8;scaleStep<=8;scaleStep++)',
               'function structuralColumnInkBounds', 'structuralColumnBox',
               'Residual antialiasing after an Excel fit-to-page scale',
               'MAX_LAYOUT_REGIONS=12', 'function findTableGridBand', 'lineCount:group.length',
               'cluster.lastX-cluster.firstX+1<=8', 'horizontalCandidates',
               "horizontalBest.detection='horizontal+vertical'", "detection.startsWith('horizontal')",
               'function strongestStructuralRowBox', 'rowAlignment.activeMinX',
               'function extractTableVerticalRules', 'function detectColumnWidthBoundaryChange',
               "kind:'column-width'", "columnMode=columnStructureChange?`column-${columnStructureChange.kind}`:columnBoundaryChange?'column-boundary-width'",
               'const beforeMinX=Math.max(tableBand.minX,left-edgePadding)', 'mappedAfterBoundary', 'changes.length>=3',
               'afterMinX=Math.max(tableBand.minX,afterLeft-edgePadding)', '.slice(0,1)',
               'structuralColumnBox.beforeBox', 'structuralColumnBox.afterBox',
               'if(columnStructureChange||columnBoundaryChange||!reliableColumnRules)', 'localizedResiduals', 'collectMaskComponents', 'const normalizeBox=box=>',
               'function strongestLocalBox', 'structuralRowBoxes', 'structuralSplit',
               'function findTableGridBands', 'function extractHeaderCellRules',
               'function detectColumnWidthBoundaryChanges', 'structuralColumnBoxes=[]',
               'difference>8',
               'verticalBand=null', 'activeMinY', 'dominantHorizontal']:
    if needed not in diff_worker:
        raise SystemExit(f'wide layout-scale correction missing: {needed}')
if 'candidate=Math.max(0,expected-3)' in diff_worker:
    raise SystemExit('diff refinement must not overfit content with a +/-3px per-line search')
if '-WindowStyle Hidden -PassThru -RedirectStandardOutput' not in launch:
    raise SystemExit('normal startup must keep the local server console hidden')
for forbidden in ['WindowStyle Normal', '起動ログ: <code>', '接続先: <code>', '<details class="dev-log">']:
    if forbidden in launch + html:
        raise SystemExit(f'development diagnostics are still visible during normal startup: {forbidden}')
if '<pre id="log" hidden></pre>' not in html:
    raise SystemExit('browser diagnostics sink must remain available without a visible development panel')

# A modified sheet must always lead the user to a visible highlight. Pages known to
# be unchanged are skipped initially, and sub-threshold pixel changes use a local
# fallback before the UI resorts to a whole-page indication.
for needed in ['function preferredDiffPageIndex', 'unchangedPageNumbers',
               'preferredDiffPageIndex(sheet)', 'preferredDiffPageIndex(preferred)',
               'このページに変更はありません。同じ原稿内の別ページに変更があります。',
               "if(!regions.length&&!noiseSuppressed){regions=[fullDiffRegion('modified'", 'analysis.fallbackUsed']:
    if needed not in appjs:
        raise SystemExit(f'visible modified-page highlight fallback missing: {needed}')
for needed in ['function buildFallbackRegion', 'looseCounts', 'difference>8',
               'if(!regions.length)', 'fallbackUsed=true']:
    if needed not in diff_worker:
        raise SystemExit(f'small-difference region fallback missing: {needed}')
for forbidden in ['buildMicroTextFallbackRegion', 'exactCounts',
                  'MICRO_EXACT_THRESHOLD', 'MIN_MICRO_PIXELS']:
    if forbidden in diff_worker:
        raise SystemExit(f'noise-prone exact-pixel fallback remains: {forbidden}')
for needed in ['extractDiffPdfTextItems', 'buildNumericTextDiffResult',
               'buildNumericTextDiffRegions', 'mergeDiffRegionsWithText',
               'diffPdfNumericSignature', 'diffPdfNumericFragments',

               'parentIndex', 'fragmentIndex', 'diffPdfTextTemplatesMatch',

               'diffPdfNonNumericFingerprint', 'diffRegionsShareTextRow',
               'unmatchedDiffPdfNumericItems', 'pairChangedDiffPdfNumbers',
               'diffPdfNumericCenterDistance', 'matched.pairs.length>8',
               'diffPdfNonNumericFingerprint(beforeItems)===diffPdfNonNumericFingerprint(afterItems)',
               'isDiffNumericText', "source:'pdf-text'",
               'const textPromise=beforeRaw&&afterRaw', 'numericOnly',
               'selectDiffSemanticResult', "mode:'numeric'", 'rowStructureAdjusted',
               '数値が変わった箇所だけを強調しました。']:
    if needed not in appjs:
        raise SystemExit(f'PDF numeric-text diff supplement missing: {needed}')
for forbidden in ['groupChangedDiffPdfText', 'diffPdfTextGroupHasNumericChange',
                  'items.map((item,index)=>({...item,index,signature:diffPdfNumericSignature(item.text)}))']:
    if forbidden in appjs:
        raise SystemExit(f'order-dependent numeric text matching remains: {forbidden}')
if "!regions.length||analysis.fallbackUsed||Number(analysis.changedRatio||0)<.01" in appjs:
    raise SystemExit('numeric text inspection must not be skipped because raster noise exceeded one percent')


# 2026-08-05 semantic row insertion/deletion localization --------------------
for needed in ['dedupeDiffPdfRowItems', 'groupDiffPdfTextRows', 'matchDiffPdfTextRows',
               'buildTextRowStructureDiffResult', 'new Uint16Array',
               "kind=delta>0?'added':'removed'", "source:'pdf-row'",
               'const pageRowDiff=buildTextRowStructureDiffResult',
               'diffPdfTextItemsInBand', 'buildLocalizedTextRowStructureDiffResult', 'diffHasLocalizedRowSignal',
               'pageRowDiff.confident?pageRowDiff:buildLocalizedTextRowStructureDiffResult',
               'rowDiff.confident&&diffHasLocalizedRowSignal', '追加・削除された行だけを強調しました。']:
    if needed not in appjs:
        raise SystemExit(f'semantic PDF row localization missing: {needed}')
for needed in ['rowStructureAdjusted', 'rowStructureRegions',
               'components=components.filter(component=>rowBoxes.some',
               'structuralSplit>=0&&Math.abs(columnAlignment.structuralJump)>=3']:
    if needed not in diff_worker:
        raise SystemExit(f'structural diff regression guard missing: {needed}')
for needed in ['function extractTableHorizontalRules', 'function tableRowRasterDistance', 'function detectTableRowInsertionDeletion',
               'tableRowStructureDetected:!!tableRowStructureChange', 'tableRowStructureBand', 'function detectRowHeightChanges',
               'function detectAlignedRowHeightChange', "kind:'row-height'",
               'rowHeightAdjusted', "const rowMode=rowHeightAdjusted?'row-boundary-height'",
               'headerNearTable', 'function detectColumnInsertionDeletion', "kind:added?'added':'removed'",
               'columnStructureKind', '`column-${columnStructureChange.kind}`',
               "kind:component.kind==='added'||component.kind==='removed'?component.kind:'modified'"]:
    if needed not in diff_worker:
        raise SystemExit(f'row-height or column add/remove localization missing: {needed}')
diff_regression=(root/'tests/diff-regression.mjs').read_text(encoding='utf-8')
for needed in ['row-height region needs side-specific geometry', "columnAdded.regions[0].kind,'added'",
               "columnRemoved.regions[0].kind,'removed'", 'added column is localized', 'removed column is localized',
               'duplicate PDF glyph rows must not multiply an inserted row',
               'side table must not split the inserted row',
               'column-width change highlights the whole column on both sides',
               'a width change is not an inserted column', 'a taller row is not an inserted column',
               'more than twenty vertical rules are not one credible table column model',
               'sparse noise overlapping unchanged text must be suppressed',
               'dense mixed-content report still recognizes one inserted row',
               'rich report column resize highlights the complete table column',
               'many scattered antialiasing specks on a dense report must be suppressed together',
               'two changed values stay detectable among paragraphs and multiple tables',
               'actual rich-report font-color pixels are not classified as raster noise',
               'page text comparison uses separate before/after physical page numbers after an insertion',
               'reverse page text comparison preserves the shifted physical page mapping']:
    if needed not in diff_regression:
        raise SystemExit(f'structural regression test missing: {needed}')
pdf_corpus_generator=(root/'app/tools/create-pdf-diff-corpus.py').read_text(encoding='utf-8')
pdf_corpus_selfcheck=(root/'app/tools/pdf-diff-corpus-selfcheck.ps1').read_text(encoding='utf-8-sig')
for needed in ['riskAlignment', 'page-duplicate-before.pdf', 'page-redesign-ambiguous.pdf', 'scan-noise-before.pdf']:
    if needed not in pdf_corpus_generator:
        raise SystemExit(f'PDF residual-risk corpus missing: {needed}')
for needed in ['redesign-ambiguous', 'duplicate-forward', 'scan-noise', 'mappingAmbiguous']:
    if needed not in pdf_corpus_selfcheck:
        raise SystemExit(f'PDF residual-risk quality gate missing: {needed}')
if 'centerX=(tableBand.minX+tableBand.maxX)/2;half=Math.max(8,(tableBand.maxX-tableBand.minX+1)/2)' in diff_worker:
    raise SystemExit('unreliable column rules must not highlight the whole table')


# 2026-08-05 Word prose/resume/column layout localization -------------------
for needed in ['diffPdfTextDocumentCache', 'extractDiffPdfDocumentTextPages',
               'diffPdfTextBoxUnion', 'mergeAdjacentDiffTextRegions',
               'findDiffPdfTextColumnSplit', 'buildColumnTextRowStructureDiffResult',
               'buildDocumentTextRowStructureDiffResult', 'buildTextLayoutShiftDiffResult',
               "source:'pdf-document-row'", "source:'pdf-layout'", "mode:'document-reflow'",
               'if(semantic.suppressRaster)noiseSuppressed=true']:
    if needed not in appjs:
        raise SystemExit(f'Word prose/resume PDF localization missing: {needed}')
for needed in ['wrapped paragraph lines merge into one edit',
               'repaginated unchanged text is not an edit on page 2',
               'Page-wide rows can also be masked by a two-column resume',
               "spacing.regions[0].source,'pdf-layout'"]:
    if needed not in diff_regression:
        raise SystemExit(f'Word prose/resume regression missing: {needed}')


# 2026-08-05 identical-layout raster noise suppression ----------------------
for needed in ['diffPdfTextLayoutFingerprint', 'diffRegionOverlapsPdfText',
               'shouldSuppressDiffRasterNoise', 'diffRegionOverlapsPdfText(region,beforeItems',
               'changedRatio>=.001', 'density<.035', 'changedRatio<.00055&&density<.12',
               'if(!regions.length&&!noiseSuppressed)',
               '微小な画像描画ノイズを除外しました。']:
    if needed not in appjs:
        raise SystemExit(f'identical-layout raster noise guard missing: {needed}')


# 2026-08-04 unrestricted worksheet names and assignment inbox -------------
for forbidden in ["sheetName -match '^[0-9]+", '半角数字だけにしてください', 'シート名が半角数字のみ']:
    if forbidden in server:
        raise SystemExit(f'numeric-only worksheet restriction remains: {forbidden}')
for needed in ['function Get-WorksheetStorageStem', 'function New-WorksheetPageId',
               'sheetIndex = $i', '$newVolume = Get-LegacyVolumeFromTargetId', 'enabled=($newVolume -ne \'none\')',
               "newPagesAreUnassigned=($newVolume -eq 'none')", 'Get-WorksheetStorageStem $sheetName',
               "Set-NoteProperty $page 'enabled' ($volume -ne 'none')"]:
    if needed not in server:
        raise SystemExit(f'unrestricted worksheet/inbox behavior missing: {needed}')
_update_pages = server.split('function Update-WorkbookPagesFromInspection',1)[1].split('\nfunction ',1)[0]
for needed in ['$knownBySheet', 'Insert-PageInSheetOrder $Structure $page $newVolume', '$removed.Count -gt 0']:
    if needed not in _update_pages:
        raise SystemExit(f'existing page assignment preservation missing: {needed}')
_locked_render_commit = server.split('$pageSync = Update-StructureLocked $Language',1)[1].split('$timingsMs.total',1)[0]
for needed in ['[string]$_.workbookId -eq $WorkbookId', '[string]$_.sheetName -eq [string]$r.sheetName']:
    if needed not in _locked_render_commit:
        raise SystemExit(f'unrestricted worksheet render commit fallback missing: {needed}')
for needed in ['未振り分け（出力しない）', '出力先へ移したページだけ',
               '原稿内の順序に整える', '未振り分けへ戻す', 'assignment-guide']:
    if needed not in html + appjs:
        raise SystemExit(f'page assignment inbox UI missing: {needed}')


for needed in ['ページをクリックして選ぶ', '移動先を押す', '慣れたら直接ドラッグも使えます',
               'createPageDragGhost', 'directCardDrag', 'page-drag-ghost']:
    if needed not in html + appjs + css:
        raise SystemExit(f'direct page movement UX missing: {needed}')
for needed in ['data-view-nav="pages"', '<span>ページ構成</span>', 'class="page-tools-details"', '<summary>表示・検索・履歴</summary>', 'page-command-bar:not(.has-selection) .page-destination-actions']:
    if needed not in html + css:
        raise SystemExit(f'page composition progressive disclosure missing: {needed}')
if "pages: 'ページ構成'" not in appjs:
    raise SystemExit('page composition view name must match the navigation label')
if 'grid-template-columns:repeat(3,minmax(0,1fr))' not in css:
    raise SystemExit('thumbnail assignment lanes must remain simultaneously visible on desktop')
for needed in ['boardSavePromise=boardSavePromise.then(persist,persist)',
               'newerBoardExists=requestRevision!==boardSaveRevision',
               'pendingBoardSaveVolumes=afterVolumes',
               'undo?.afterVolumes||collectBoardVolumes()',
               'rememberPageLayoutUndo(undo.label,undo.volumes,afterVolumes)',
               'void saveBoardOrder()',
               'savedBeforeHistory=await boardSavePromise',
               'beforeVolumes=JSON.parse(beforeSignature)',
               'afterVolumes=JSON.parse(drag.previewSignature||afterSignature)']:
    if needed not in appjs:
        raise SystemExit(f'page movement save serialization missing: {needed}')


# 2026-08-06 page-preview organizer controls -------------------------------
for needed in ['id="preview-page-tools"', 'id="preview-prev-page"',
               'id="preview-next-page"', 'id="preview-page-position"',
               'id="preview-target-buttons"', 'id="preview-move-none"']:
    if needed not in html:
        raise SystemExit(f'page-preview organizer markup missing: {needed}')
for needed in ['function previewOrganizerPageIds', 'function syncPreviewOrganizerControls',
               'function navigatePreviewPage', 'function moveCurrentPreviewPageToVolume',
               "['ArrowLeft','ArrowRight'].includes(e.key)", 'originalFocus?.isConnected',
               "currentCard?.querySelector('[data-preview-page]')"]:
    if needed not in appjs:
        raise SystemExit(f'page-preview organizer behavior missing: {needed}')
for needed in ['.preview-page-tools', '.preview-page-assign', '.notice{z-index:120}']:
    if needed not in css:
        raise SystemExit(f'page-preview organizer styling missing: {needed}')


# 2026-08-06 multi-step page layout undo/redo ------------------------------
for needed in ['id="page-layout-undo-btn"', 'id="page-layout-redo-btn"',
               'aria-label="ページ構成の操作履歴"', 'id="i-undo"', 'id="i-redo"']:
    if needed not in html:
        raise SystemExit(f'page layout history markup missing: {needed}')
for needed in ['pageLayoutUndoStack', 'pageLayoutRedoStack', 'PAGE_LAYOUT_HISTORY_LIMIT = 20',
               'function clonePageVolumes', 'function pageVolumeSnapshotsEqual',
               'function clearPageLayoutHistory', 'async function applyPageLayoutHistory',
               'async function redoLastPageLayout', "shortcutKey==='z'", "shortcutKey==='y'",
               'pageLayoutRedoStack=[]', 'syncPageSelectionUi();']:
    if needed not in appjs:
        raise SystemExit(f'multi-step page layout history behavior missing: {needed}')
for needed in ['.page-history-actions', '.page-history-actions .btn']:
    if needed not in css:
        raise SystemExit(f'page layout history styling missing: {needed}')


# 2026-08-06 keyboard-first page board ------------------------------------
for needed in ['class="page-keyboard-hint"', '<kbd>Space</kbd>', '<kbd>Ctrl+A</kbd>',
               '<kbd>Alt+矢印</kbd>']:
    if needed not in html:
        raise SystemExit(f'page keyboard guide missing: {needed}')
for needed in ['focusPageId', 'function visiblePageRowsForKeyboard',
               'function pageRowKeyboardTarget', 'function handlePageRowNavigation',
               "e.key==='Spacebar'", "shortcutKey==='a'", "row.setAttribute('aria-keyshortcuts'",
               "row.setAttribute('aria-label',`${n}ページ目 ${title}`)"]:
    if needed not in appjs:
        raise SystemExit(f'page keyboard behavior missing: {needed}')
for needed in ['.page-keyboard-hint', '.page-row:focus-visible']:
    if needed not in css:
        raise SystemExit(f'page keyboard styling missing: {needed}')


# Browser startup must not be interrupted before navigation handlers and the
# initial state load are registered. A duplicate async prefix is valid syntax as
# two statements, but raises ReferenceError at runtime in the browser.
if re.search(r'\basync\s+async\s+function\b', appjs):
    raise SystemExit('app.js contains a duplicated async function prefix')
for needed in ['function saveBoardOrder()', 'async function sortPagesBySheet(btn=null)']:
    if needed not in appjs:
        raise SystemExit(f'page-board async function is missing: {needed}')
if "if(data.state)state=normalizeStatePayload(data.state)" in appjs:
    raise SystemExit('source registration incorrectly applies the V2 domain state to the legacy UI view')
if "selectedFiles.clear();lastFileRangeAnchor='';await refresh();" not in appjs:
    raise SystemExit('source registration does not refresh the compatibility UI state')
if "packId:String(pack?.packId||activePackId||'')" not in appjs:
    raise SystemExit('page V2 requests do not consistently identify the active document pack')
for needed in [
    "packId=[string](Get-DataProperty $body 'packId' (Get-DataProperty $body 'category' '')); volumes=$volumes",
    "packId=[string](Get-DataProperty $body 'packId' (Get-DataProperty $body 'category' '')); pageId=$itemId",
]:
    if needed not in server:
        raise SystemExit(f'page V2 route compatibility is missing: {needed}')
for needed in ["$excelStartupError = ''", 'Excel原稿をエラーとして記録し、他形式の処理を続けます。', 'if ($KeepHostOpen -and $null -eq $SharedHost)']:
    if needed not in server:
        raise SystemExit(f'mixed-source render isolation is missing: {needed}')


# 2026-08-06 Word source adapter -------------------------------------------
for needed in ["sourceType = 'word'", "adapterId = 'word-com-v1'",
               'function Inspect-WordSourceFile', 'function Register-WordSource',
               'function Render-WordSource', 'function Render-WordSnapshotForComparison',
               'function Invoke-WordDocumentToPdf', '$Script:WordRenderTimeoutSeconds = 120',
               'wordRenderProfileVersion = $Script:WordRenderProfileVersion',
               "Get-SourceCandidates @('excel','word','pdf','powerpoint')"]:
    if needed not in server:
        raise SystemExit(f'Word source adapter behavior missing: {needed}')
for needed in ["id=\"i-file-word\"", 'Excel・Word・PowerPoint・PDFを集め']:
    if needed not in html:
        raise SystemExit(f'Word source UI markup missing: {needed}')
for needed in ['CURRENT_WORD_RENDER_PROFILE_VERSION', "word:'i-file-word'",
               "sourceTypeValue(source) === 'word'", "['excel','word','powerpoint','pdf']"]:
    if needed not in appjs:
        raise SystemExit(f'Word source UI behavior missing: {needed}')
for needed in ['ConvertTo-Win32ExtendedPath', "return '\\\\?\\UNC\\'", 'Write-Utf8NoBomFileShared']:
    if needed not in server:
        raise SystemExit(f'long comparison path support missing: {needed}')


# 2026-08-07 PowerPoint source adapter ------------------------------------
for needed in ["sourceType = 'powerpoint'", "adapterId = 'powerpoint-com-v1'",
               'function Inspect-PowerPointSourceFile', 'function Register-PowerPointSource',
               'function Render-PowerPointSource', 'function Render-PowerPointSnapshotForComparison',
               'function Convert-PowerPointSplitPages', '(Get-WorksheetStorageStem $name)',
               'function Invoke-PowerPointToPdf', '$Script:PowerPointRenderTimeoutSeconds = 120',
               'powerPointRenderProfileVersion = $Script:PowerPointRenderProfileVersion',
               "Get-OfficeApplicationDiagnostic 'PowerPoint.Application' 'Microsoft PowerPoint'"]:
    if needed not in server:
        raise SystemExit(f'PowerPoint source adapter behavior missing: {needed}')
for needed in ['id="i-file-powerpoint"', 'id="template-source-powerpoint"']:
    if needed not in html:
        raise SystemExit(f'PowerPoint source UI markup missing: {needed}')
for needed in ['CURRENT_POWERPOINT_RENDER_PROFILE_VERSION', "powerpoint:'i-file-powerpoint'",
               "sourceTypeValue(source) === 'powerpoint'", 'office.powerPoint']:
    if needed not in appjs:
        raise SystemExit(f'PowerPoint source UI behavior missing: {needed}')
for needed in ['New-Object -ComObject PowerPoint.Application', '$powerPoint.AutomationSecurity = 3',
               '$presentations.Open($inputPath, -1, 0, 0)', 'Get-Process POWERPNT']:
    if needed not in (root/'app/tools/powerpoint-render-worker.ps1').read_text(encoding='utf-8'):
        raise SystemExit(f'PowerPoint render worker safety missing: {needed}')


# 2026-08-07 safe render job cancellation ---------------------------------
for needed in ['function Request-RenderJobCancellation', 'function Test-RenderJobCancellationRequested',
               'function Set-RenderJobCancelledStatus', "status' 'cancelled'", "'/api/jobs/cancel'"]:
    if needed not in server:
        raise SystemExit(f'render job cancellation behavior missing: {needed}')
for needed in ['id="progress-cancel"', 'PDF作成を中止']:
    if needed not in html:
        raise SystemExit(f'render job cancellation UI markup missing: {needed}')
for needed in ['activeRenderJobId', 'cancelActiveRenderJob', "'/api/jobs/cancel'",
               "['completed','completed-with-errors','failed','cancelled']"]:
    if needed not in appjs:
        raise SystemExit(f'render job cancellation UI behavior missing: {needed}')

# 2026-08-07 cross-department deadline progress -----------------------------
for needed in ["$state='overdue-source'", 'overdueRequiredSourceCount=$overdueRequired',
               'dueSoonRequiredSourceCount=$dueSoonRequired', 'overdueRequiredCount=',
               'dueSoonRequiredCount=']:
    if needed not in server:
        raise SystemExit(f'cross-pack deadline progress missing: {needed}')
for needed in ["'overdue-source':['期限超過','danger']", 'overdueRequiredCount',
               'dueSoonRequiredCount', 'nearestRequiredDueDate']:
    if needed not in appjs:
        raise SystemExit(f'cross-pack deadline UI missing: {needed}')


# 2026-08-06 generalized history comparison --------------------------------
for needed in ['function Get-ComparisonUnitMappings', 'function Set-ComparisonUnitMappingResult',
               "'exact-hash-sequence'", "'ambiguous-sequence'", 'unitMappings = @()',
               'beforeSheetName = $beforeName', 'afterSheetName = $afterName',
               '$Script:DiffDetailAlgorithmVersion = 23']:
    if needed not in server:
        raise SystemExit(f'generalized history comparison missing: {needed}')
for needed in ['function diffUnitDisplayName', 'function diffMatchConfidenceLabel',
               'sheet?.beforeSheetName', 'sheet?.afterSheetName', '対応不確実', '対応推定']:
    if needed not in appjs:
        raise SystemExit(f'generalized history UI missing: {needed}')


# 2026-08-07 deployment and operational readiness ---------------------------
for needed in ['function Get-OfficeApplicationDiagnostic', 'function Get-SystemDiagnostics',
               "'windows-session-unavailable'", 'minimumSupportedMajor=16',
               "status=$overall", "PDF原稿は処理できます"]:
    if needed not in server:
        raise SystemExit(f'environment diagnostics missing: {needed}')
for needed in ['id="run-diagnostics-btn"', 'id="diagnostics-result"', '動作環境を診断']:
    if needed not in html:
        raise SystemExit(f'environment diagnostics markup missing: {needed}')
for needed in ['function renderDiagnosticsResult', 'async function runSystemDiagnostics',
               "api('/api/diagnostics/run'", 'diagnosticsResult']:
    if needed not in appjs:
        raise SystemExit(f'environment diagnostics behavior missing: {needed}')
for needed in ['.diagnostic-list', '.diagnostic-row', '.diagnostics-summary']:
    if needed not in css:
        raise SystemExit(f'environment diagnostics styling missing: {needed}')

print('selfcheck ok')
