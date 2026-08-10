const qs = new URLSearchParams(location.search);
function readCookie(name) {
  const prefix = `${name}=`;
  for (const part of document.cookie.split(';')) {
    const item = part.trim();
    if (item.startsWith(prefix)) return decodeURIComponent(item.slice(prefix.length));
  }
  return '';
}
const urlToken = qs.get('token') || '';
if (urlToken) {
  try { sessionStorage.setItem('ReportBinderToken', urlToken); } catch {}
  try { document.cookie = `ReportBinderToken=${encodeURIComponent(urlToken)}; Path=/; SameSite=Lax`; } catch {}
  // トークンはAPI全体に対する資格情報。アドレスバーとブラウザ履歴に残すと、
  // 履歴同期や拡張機能経由で漏れる。保存した直後にURLから取り除く。
  try { history.replaceState(null, '', location.pathname); } catch {}
}
const token = urlToken || (() => { try { return sessionStorage.getItem('ReportBinderToken') || ''; } catch { return ''; } })() || readCookie('ReportBinderToken') || '';
function withToken(path) {
  const sep = path.includes('?') ? '&' : '?';
  return `${path}${sep}token=${encodeURIComponent(token)}`;
}
const clientId = (globalThis.crypto && globalThis.crypto.randomUUID)
  ? globalThis.crypto.randomUUID()
  : `client-${Date.now()}-${Math.random().toString(16).slice(2)}`;

let state = null;
let availableFiles = [];
let availableFilesScannedAt = '';
let activePreset = '';
let activePackId = (() => { try { return sessionStorage.getItem('ReportBinderPackId') || ''; } catch { return ''; } })();
let packAdminPacks = [];
let archivedPacksExpanded = false;
let packEditorMode = 'create';
let packEditorPackId = '';
let packEditorReturnFocus = null;
let templateManagerReturnFocus = null;
let lastSubmissionDefaultBase = '';
let draggingRow = null;
let boardSaveTimer = null;
let boardSavePromise = Promise.resolve(true);
let boardSaveRevision = 0;
let pendingBoardSaveVolumes = null;
let selectedFiles = new Set();
let selectedWorkbooks = new Set();
let expandedSourceSettings = new Set();
let selectedPages = new Set();
let pageBoardView = (() => { try { return sessionStorage.getItem('ReportBinderPageBoardView') === 'detail' ? 'detail' : 'thumbnail'; } catch { return 'thumbnail'; } })();
let activePageVolume = (() => { try { return sessionStorage.getItem('ReportBinderActivePageVolume') || 'none'; } catch { return 'none'; } })();
// Opening on the 未振り分け tab showed an empty board whenever everything was
// already sorted, which reads as "my pages are gone". Pick a tab that has pages
// once per session, then respect whatever the user selects.
let pageVolumeAutoPicked = (() => { try { return !!sessionStorage.getItem('ReportBinderActivePageVolume'); } catch { return false; } })();
let pageFilterText = '';
let pageThumbnailSize = (() => { try { const value=Number(sessionStorage.getItem('ReportBinderPageThumbnailSize'));return Number.isFinite(value)?Math.min(260,Math.max(150,value)):190; } catch { return 190; } })();
let pageLayoutUndoStack = [];
let pageLayoutRedoStack = [];
let pageLayoutHistoryBusy = false;
let pageMutationBusy = false;
let pendingPageLayoutUndo = null;
let pageThumbnailObserver = null;
let pageThumbnailRenderActive = 0;
const pageThumbnailRenderQueue = [];
const pageThumbnailCache = new Map();
const pageVolumeCollapseOverrides = new Map();
const PAGE_THUMBNAIL_CACHE_LIMIT = 48;
const PAGE_THUMBNAIL_RENDER_LIMIT = 3;
const PAGE_LAYOUT_HISTORY_LIMIT = 20;
let lastFileRangeAnchor = '';
let lastWorkbookRangeAnchor = '';
let lastPageRangeAnchor = '';
let currentPreviewObjectUrl = '';
let currentPreviewPageId = '';
let previewPageIds = [];
let previewOrganizerMode = false;
let updateMonitorTimer = null;
let updateMonitorInFlight = false;
let updateVisibilityBound = false;
let lastUpdateScanAt = 0;
let renderJobActive = false;
let activeRenderJobId = '';
let activeRenderCancelRequested = false;
// 提出用PDFの出力ジョブ。進捗パネルは変換PDFと共用するため、中止の宛先を区別する。
let activeFinalJobId = '';
let activeFinalCancelRequested = false;
let activeRenderJobStatus = null;
let folderPickerBusy = false;
let diagnosticsResult = null;
let packProgressAttentionOnly = false;
let finalPreflightCheckedAt = '';
let packReviewState = null;
let packReviewInFlight = null;
let autoStateInFlight = null;
const seenAutoNotificationIds = new Set();
let activeView = (() => { try { const stored=sessionStorage.getItem('ReportBinderView')||'excel'; return stored==='folders'?'excel':stored; } catch { return 'excel'; } })();
let fileFilterText = '';
const pdfObjectUrls = new Set();
let noticeTimer = null;
let lastPageBoardRenderSignature = '';
let modalReturnFocus = null;
let confirmReturnFocus = null;
let confirmResolver = null;
let diffReturnFocus = null;
let diffPollToken = 0;
let diffSyncingScroll = false;
let diffBrowserRenderSerial = 0;
let diffPdfJsPromise = null;
let diffAnalysisWorker = null;
let diffAnalysisRequestSerial = 0;
const diffAnalysisPending = new Map();
const diffActiveRenderTasks = new Set();
const diffPdfDocumentCache = new Map();
const diffPdfTextDocumentCache = new Map();
const diffBrowserPageCache = new Map();
const diffDetailResponseCache = new Map();
const DIFF_RENDER_SCALE = 120 / 72;
const DIFF_PAGE_CACHE_LIMIT = 6;
const DIFF_PDF_CACHE_LIMIT = 16;
const DIFF_DETAIL_CACHE_LIMIT = 12;
// 行構造の比較に渡すテキスト項目数の上限。巻全体を平坦化した配列も通るため、
// 上限が無いとページ数×1ページの項目数だけ膨らんで実用時間を超える。
const DIFF_TEXT_ROW_ITEM_LIMIT = 4000;
const DIFF_DETAIL_TIMEOUT_MS = 15000;
const SNAPSHOT_HISTORY_CACHE_MS = 60000;
let historyPanelsInitialized = false;
let snapshotHistoryLoadSerial = 0;
let snapshotHistoryWorkbookOptionsSignature = '';
const snapshotHistoryResponseCache = new Map();
const snapshotHistoryRequestCache = new Map();
const diffViewState = {
  workbookId: '',
  fromSnapshotId: '',
  toSnapshotId: '',
  detail: null,
  selectedSheetKey: '',
  pageIndex: 0,
  regionIndex: -1,
  filter: 'all',
  unreviewedOnly: false,
  mode: 'side',
  mobileTab: 'before'
};
let highlightedPageIds = new Set();
let pageHighlightTimer = null;
// サーバー側 $Script:ExcelPrintProfileVersion のフォールバック値。
// 実行時は /api/state の excelPrintProfileVersion を優先する（二重管理によるズレを防ぐ）。
const CURRENT_RENDER_PROFILE_VERSION = 2026060403;
const CURRENT_PDF_IMPORT_PROFILE_VERSION = 2026080601;
const CURRENT_WORD_RENDER_PROFILE_VERSION = 2026080601;
const CURRENT_POWERPOINT_RENDER_PROFILE_VERSION = 2026080701;
function currentRenderProfileVersion() {
  const v = Number(state?.excelPrintProfileVersion || 0);
  return Number.isFinite(v) && v > 0 ? v : CURRENT_RENDER_PROFILE_VERSION;
}

const $ = (id) => document.getElementById(id);

const VIEW_NAMES = {
  dashboard: 'ホーム',
  excel: '原稿を登録・PDF化',
  history: '変更履歴・比較',
  pages: 'ページ構成',
  final: '提出用PDF'
};

function setActiveView(view, options = {}) {
  const next = VIEW_NAMES[view] ? view : 'excel';
  const viewChanged = next !== activeView;
  activeView = next;
  try { sessionStorage.setItem('ReportBinderView', activeView); } catch {}
  document.querySelectorAll('[data-view-panel]').forEach(panel => {
    panel.classList.toggle('active', panel.getAttribute('data-view-panel') === activeView);
  });
  document.querySelectorAll('[data-view-nav]').forEach(btn => {
    const on = btn.getAttribute('data-view-nav') === activeView;
    btn.classList.toggle('active', on);
    // Step-bar entries are steps in a progress sequence, the sidebar entries are
    // pages; both carry data-view-nav so the value has to be picked per element.
    if (on) btn.setAttribute('aria-current', btn.hasAttribute('data-progress-view') ? 'step' : 'page');
    else btn.removeAttribute('aria-current');
  });
  // renderStepBar() only ran from renderAll(), so the "you are here" mark on the
  // step bar lagged behind plain navigation. It is visible during work now, so
  // it has to track every view change.
  if (state) renderStepBar();
  if (!options.noScroll) {
    hideMessage();
    const reduceMotion = window.matchMedia?.('(prefers-reduced-motion:reduce)').matches;
    window.scrollTo({top: 0, behavior: (options.instant || reduceMotion) ? 'auto' : 'smooth'});
    requestAnimationFrame(() => {
      const heading=document.querySelector(`[data-view-panel="${CSS.escape(activeView)}"] h2`);
      if(!heading)return;
      heading.setAttribute('tabindex','-1');
      heading.focus({preventScroll:true});
      heading.addEventListener('blur',()=>heading.removeAttribute('tabindex'),{once:true});
    });
  }
  // renderAll also reapplies the active tab. Do not turn that DOM refresh into another
  // shared-folder request; load only on first initialization, actual tab entry, or explicit refresh.
  if (activeView === 'history' && state && (viewChanged || !historyPanelsInitialized || options.reloadPanels)) {
    void loadHistoryPanels({force:!!options.reloadPanels});
  }
  // aggregateFinalState() (header badge, nav dot, dashboard next-action) reads
  // finalReadiness, which used to be fetched only when the final screen was
  // opened - so a fresh launch showed "提出用PDF 未出力" for a pack that was
  // already output. Fetch it once for the active pack whatever the view.
  if (state && !finalReadinessLoaded()) {
    void loadFinalReadiness({render:false}).then(()=>{
      renderGlobalHeader();renderNavBadges();renderDashboardOverview();
      // Repaint the output cards unconditionally, not just while the final view
      // is showing. At boot activePackId is not resolved yet, so this late load
      // is often the first one that succeeds; skipping the cards left them at
      // 0ページ until the user opened the screen, and the correction was then
      // visible as a flash right after the panel appeared.
      renderFinalOverview();renderVolumeLinks();
    });
  }
  if (activeView === 'final' && state) {
    // These three used to be fired here unguarded. renderAll() calls
    // setActiveView(), so every repaint restarted all three, and each one
    // repainted on its own when its request landed - the screen visibly
    // reshuffled a few hundred ms after it had already been drawn.
    if(viewChanged||options.reloadPanels||!finalReadinessLoaded())void loadFinalPanels({force: !!options.reloadPanels});
  }
  if (activeView === 'excel' && state?.configured) void loadAutoStateDetail();
}

function formatDateTime(value) {
  const text = String(value || '').trim();
  if (!text) return '—';
  const normalized = text.replace('T', ' ').replace(/\.\d+Z?$/, '').replace(/Z$/, '');
  const m = normalized.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s+(\d{1,2}):(\d{2})/);
  if (m) return `${m[1]}/${m[2].padStart(2,'0')}/${m[3].padStart(2,'0')} ${m[4].padStart(2,'0')}:${m[5]}`;
  return text.length > 22 ? text.slice(0, 22) : text;
}
function formatDateOnly(value) {
  const text = String(value || '').trim();
  if (!text) return '—';
  const m = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  return m ? `${m[1]}/${m[2].padStart(2,'0')}/${m[3].padStart(2,'0')}` : text;
}
function compactDateTime(value) {
  const f = formatDateTime(value);
  return f === '—' ? '' : f;
}
function formatDateTimeWithSeconds(value) {
  const text = String(value || '').trim();
  if (!text) return '—';
  const normalized = text.replace('T', ' ').replace(/\.\d+Z?$/, '').replace(/Z$/, '');
  const m = normalized.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m) return `${m[1]}/${m[2].padStart(2,'0')}/${m[3].padStart(2,'0')} ${m[4].padStart(2,'0')}:${m[5]}:${(m[6] || '00').padStart(2,'0')}`;
  return formatDateTime(value);
}
function formatLocalDateTimeMinuteFromUnixMs(value) {
  const ms = Number(value);
  if (!Number.isFinite(ms) || ms <= 0) return '';
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return '';
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function formatSubmissionFileModifiedAt(file) {
  // The server returns the Windows/Explorer local display value. Do not convert it
  // again in the browser, because timezone conversion and SMB metadata caching can
  // make the minute differ from Explorer.
  const direct = String(file?.modifiedAtDisplay || '').trim();
  if (direct) return direct;
  return formatDateTime(file?.modifiedAt || file?.modifiedAtUtc || file?.updatedAt || '');
}
function fileListScanSuffix() {
  const formatted = formatDateTimeWithSeconds(availableFilesScannedAt);
  return formatted === '—' ? '' : ` ・ 一覧更新 ${formatted.slice(11)}`;
}
function setFileSelectionSummary(total, selected) {
  const summary = $('file-selection-summary');
  if (summary) summary.textContent = `${total}件中 ${selected}件を選択${fileListScanSuffix()}`;
}
function workbookDisplayName(w) { return String(w?.displayName || w?.fileName || w?.relativePath || w?.workbookId || '').trim() || '原稿'; }
function fileDisplayName(f) { return String(f?.fileName || f?.relativePath || '').trim() || '原稿'; }
function pageCountsByVolume() {
  const pages = pagesForActivePreset();
  const first=activeTargetVolumes()[0]||mainVolume();
  const counts={none:pages.filter(p=>p.enabled===false||String(p.volume||first)==='none').length,total:pages.length,volumes:{}};
  for(const volume of activeTargetVolumes())counts.volumes[volume]=pages.filter(p=>p.enabled!==false&&String(p.volume||first)===volume).length;
  counts.main=counts.volumes[mainVolume()]||0;counts.appendix=counts.volumes[appendixVolume()]||0;
  return counts;
}
function latestWorkbookStamp(w) { return w?.lastRenderedAt || w?.updatedAt || w?.modifiedAt || w?.registeredAt || ''; }
function outputFileName(path) { return String(path || '').split(/[\\/]/).filter(Boolean).pop() || 'PDF'; }
function escapeAttr(s) { return escapeHtml(s).replace(/`/g, '&#96;'); }

function asArray(value) {
  if (value == null) return [];
  return Array.isArray(value) ? value : [value];
}
function normalizeStatePayload(payload) {
  const p = payload || {};
  p.packTemplates = asArray(p.packTemplates);
  p.packs = asArray(p.packs);
  p.structure = p.structure || {};
  p.structure.workbooks = asArray(p.structure.workbooks);
  p.structure.pages = asArray(p.structure.pages);
  p.structure.volumes = p.structure.volumes || {};
  p.finalReadiness = p.finalReadiness || {};
  p.recentErrors = asArray(p.recentErrors);
  return p;
}


function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function normPath(s) { return String(s || '').replace(/\\/g, '/').replace(/\/+/g, '/').toLowerCase(); }
function pathJoin(base, child) { return String(base || '').replace(/[\\/]+$/,'') + '\\' + child; }
function badge(text, cls='') { return `<span class="badge ${cls}">${escapeHtml(text)}</span>`; }
// Volume IDs retain a private compatibility prefix. Language belongs to a
// document pack/template and never selects another application workspace.
const INTERNAL_VOLUME_PREFIX = 'ja';
function mainVolume() { return `${INTERNAL_VOLUME_PREFIX}-main`; }
function appendixVolume() { return `${INTERNAL_VOLUME_PREFIX}-appendix`; }
function targetVolume(targetId) { return `${INTERNAL_VOLUME_PREFIX}-${String(targetId || '').trim().toLowerCase()}`; }
function targetIdFromVolume(volume) {
  const match=String(volume||'').toLowerCase().match(/^(?:ja|en)-([a-z0-9][a-z0-9_-]{0,47})$/);
  return match?match[1]:String(volume||'').toLowerCase()==='none'?'unassigned':'';
}
function activeTargetDefinitions() {
  const targets=asArray(activePackRecord()?.targets).filter(target=>String(target?.targetId||'').trim());
  return targets;
}
function activePackDefaultTargetLabel() {
  const targetId=String(activePackRecord()?.rules?.newItemDestination||'unassigned');
  if(!targetId||targetId==='unassigned')return '未振り分け（ページ構成で確認）';
  const target=activeTargetDefinitions().find(item=>String(item?.targetId||'')===targetId);
  return String(target?.displayName||targetId);
}
function activeTargetVolumes() { return activeTargetDefinitions().map(target=>targetVolume(target.targetId)); }

function volumeLabel(v) { return ({'ja-main':'本体','ja-appendix':'補足','en-main':'Main','en-appendix':'Appendix','none':'出力しない'})[v] || v; }
function assignmentVolumeLabel(v) { return String(v)==='none'?'未振り分け':volumeLabel(v); }
function configured() { return !!(state?.configured && state?.paths?.submissionDir && state?.paths?.dataDir && state?.paths?.outputDir); }
function getWorkbook(workbookId) { return (asArray(state?.structure?.workbooks)).find(w => String(w.workbookId) === String(workbookId)); }
function sheetKey(sheetName) { return String(sheetName ?? '').replace(/[^0-9A-Za-z]+/g, '-'); }
function resolvedPageId(page) {
  if (!page) return '';
  const explicit = String(page.pageId || page.id || '').trim();
  if (explicit) return explicit;
  const wb = String(page.workbookId || '').trim();
  const sheet = String(page.sheetName ?? '').trim();
  if (!wb || !sheet) return '';
  return `${wb}-${sheetKey(sheet) || 'sheet'}`;
}
function getPage(pageId) {
  const target = String(pageId || '').trim();
  if (!target) return null;
  return (asArray(state?.structure?.pages)).find(p => resolvedPageId(p) === target || String(p.pageId || '') === target || String(p.id || '') === target) || null;
}
function getPageByWorkbookSheet(workbookId, sheetName) {
  const wb = String(workbookId || '');
  const sheet = String(sheetName || '');
  if (!wb || !sheet) return null;
  return (asArray(state?.structure?.pages)).find(p => String(p.workbookId || '') === wb && String(p.sheetName || '') === sheet) || null;
}


function volumeReadiness(volume, preset=activePreset) {
  const key=activePackRecord()?.category ? preset : (activePackId||preset);
  return state?.finalReadiness?.[key]?.volumes?.[volume] || {
    canBuild:false, pageCount:0, status:'not-built', displayState:'not-built', blockers:[], staleReasons:[], outputPdf:'', outputPdfExists:false
  };
}
function activeReadinessItems() {
  return activeTargetVolumes().map(volume => ({volume, ready:volumeReadiness(volume)})).filter(x => Number(x.ready.pageCount || 0) > 0);
}
function aggregateFinalState() {
  const items = activeReadinessItems();
  if (!items.length) return {state:'not-built', text:'提出用PDFは未出力'};
  const displays = items.map(x => String(x.ready.displayState || 'not-built'));
  if (displays.includes('blocked')) return {state:'blocked', text:'提出用PDFを出力できません'};
  if (displays.includes('needs-rebuild')) return {state:'needs-rebuild', text:'提出用PDFの再出力が必要'};
  if (displays.includes('output-missing')) return {state:'output-missing', text:'前回出力が見つかりません'};
  if (displays.includes('not-built')) return {state:'not-built', text:'提出用PDFは未出力'};
  return {state:'built', text:'提出用PDFは最新です'};
}
function finalIsComplete() {
  const items = activeReadinessItems();
  return items.length > 0 && items.every(x => x.ready.status === 'built' && x.ready.outputPdfExists === true && asArray(x.ready.blockers).length === 0);
}
function unresolvedPageCount() {
  return pagesForActivePreset().filter(p => p.enabled === false || String(p.volume || 'none') === 'none').length;
}
function isEditing() {
  const el = document.activeElement;
  return !!el && el.matches('input:not([readonly]), select, textarea');
}
function isModalOpen() { return !!$('preview-modal') && !$('preview-modal').classList.contains('hidden'); }
function isConfirmModalOpen() { return !!$('confirm-modal') && !$('confirm-modal').classList.contains('hidden'); }
function isPackEditorOpen() { return !!$('pack-editor-modal') && !$('pack-editor-modal').classList.contains('hidden'); }
function isTemplateManagerOpen() { return !!$('template-manager-modal') && !$('template-manager-modal').classList.contains('hidden'); }
function isPageLayoutRevisionsOpen() { return !!$('page-layout-revisions-modal') && !$('page-layout-revisions-modal').classList.contains('hidden'); }
function isAnyModalOpen() { return isModalOpen() || isDiffModalOpen() || isConfirmModalOpen() || isPackEditorOpen() || isTemplateManagerOpen() || isPageLayoutRevisionsOpen(); }
function shouldPreserveEditorDom() { return isEditing() || draggingRow !== null || isAnyModalOpen(); }
function iconUse(id, cls='') { return `<svg class="icon ${cls}" aria-hidden="true"><use href="#${id}"></use></svg>`; }

function statusLabel(s) {
  return ({
    missing:'ファイルなし', new:'PDF未作成', rendering:'PDF作成中', 'rendered-unchecked':'PDF済み', confirmed:'PDF済み',
    'excel-updated':'更新あり', 'source-updated':'更新あり', rejected:'差し戻し', finalized:'完了', rendered:'PDF済み', stale:'更新あり',
    'not-rendered':'PDF未作成', 'render-error':'作成エラー', 'needs-rebuild':'再出力待ち', built:'出力済み', 'not-built':'未出力'
  })[s] || (s || '未設定');
}
function statusClass(s) {
  if (s === 'missing' || s === 'rejected' || s === 'render-error') return 'danger';
  if (s === 'excel-updated' || s === 'source-updated' || s === 'stale' || s === 'needs-rebuild' || s === 'new' || s === 'not-rendered' || s === 'rendering') return 'attention';
  return 'neutral';
}

function isLatestPdfWorkbook(w) {
  if (!w || String(w.status || '') === 'missing' || String(w.status || '') === 'render-error') return false;
  const current = String(w.currentExcelHash || '').trim();
  const rendered = String(w.lastRenderedExcelHash || '').trim();
  const profileVersion = Number(w.renderProfileVersion || 0);
  if (!w.lastRenderedAt || !rendered) return false;
  if (profileVersion < requiredRenderProfileVersion(w)) return false;
  if (current && current !== rendered) return false;
  if (['new','excel-updated','source-updated','not-rendered'].includes(String(w.status || ''))) return false;
  return true;
}
function latestPdfStatusForWorkbook(w) {
  return isLatestPdfWorkbook(w) ? badge('変換PDFは最新', 'neutral') : badge('変換PDF作成が必要', 'attention');
}

function latestPdfDetailForWorkbook(w) {
  if (!w) return '';
  if (String(w.status || '') === 'render-error') return w.lastErrorUser || userFriendlyError(w.lastError || 'PDF作成エラー');
  return '';
}
function pdfStatusForPage(p) {
  const wb = getWorkbook(p?.workbookId);
  if (wb && !isLatestPdfWorkbook(wb)) return badge('変換PDF作成が必要', 'attention');
  if (p.status === 'render-error') return badge('作成エラー', 'danger');
  if (p.contentPdf) return badge('変換PDFは最新', 'neutral');
  return badge('変換PDF作成が必要', 'attention');
}

function pagePreviewAvailable(p) {
  const wb = getWorkbook(p?.workbookId);
  return !!(p?.contentPdf && (!wb || isLatestPdfWorkbook(wb)));
}
function pagePdfSubtext(p) {
  const wb = getWorkbook(p?.workbookId);
  if (wb && !isLatestPdfWorkbook(wb)) return '先にPDF作成してください';
  return p?.contentPdf ? 'クリックで確認' : 'PDF作成してください';
}
// サーバーが居なくなったときの案内。fetch の失敗は素の "Failed to fetch" で、
// 何が起きたのかも、どう戻ればよいのかも分からない。30分の無操作やノートPCの
// スリープで自動終了する仕様なので、これは事故ではなく日常の状態。
const SERVER_GONE_MESSAGE = 'ReportBinderが終了しています。デスクトップの「資料をPDFにまとめる.cmd」をもう一度実行してから、この画面を読み込み直してください。（30分間操作がないと自動で終了します）';
function userFriendlyError(message) {
  const text = String(message || '');
  const raw = text.length > 300 ? text.slice(0, 300) + '…' : text;
  if (/structure\.json|登録情報を安全に読み込めません|不完全な登録情報/i.test(text)) {
    return '作業データを読み込めませんでした。原稿ファイルは変更されていません。別の原稿フォルダーを選ぶか、ReportBinderを終了してもう一度開いてください。';
  }
  // サーバーが日本語で説明しているメッセージは、そのまま表示するのが最も正確。
  // 技術的なエラー(例外・スタックトレース等)のときだけ、対処ヒントを付けて原文も併記する。
  const isTechnical = /Exception|StackTrace|HRESULT|at\s+java|COMException|System\./i.test(text);
  const hasJapanese = /[ぁ-んァ-ヶ一-龠]/.test(text);
  let hint = '';
  if (/用語 'java'|'java'\s*は.*認識され|java\.exe が見つかりません|Java Runtimeが(見つかり|あり)ません/i.test(text)) {
    hint = '提出用PDFの作成に必要な追加ソフトが入っていません。情報システム担当者に「ReportBinderのセットアップ」の実行を依頼してください。';
  } else if (/別のプロセスが使用中|being used by another process|使用中です/i.test(text)) {
    hint = 'ファイルが他のアプリで開かれています。該当ファイルを閉じてから再実行してください。';
  } else if (/アクセスが拒否|Access is denied|UnauthorizedAccess/i.test(text)) {
    hint = 'ファイルへのアクセスが拒否されました。共有フォルダーの権限や一時的な競合の可能性があります。少し待って再試行してください。';
  } else if (/COMException|HRESULT|RPC/i.test(text)) {
    hint = 'Excelの操作に失敗しました。Excelをすべて閉じてから再試行してください。';
  } else if (/FileNotFoundException|not found|見つかりません/i.test(text)) {
    hint = 'ファイルが見つかりません。削除・移動・名前変更されていないか確認してください。';
  } else if (/OutOfMemory/i.test(text)) {
    hint = 'メモリ不足で処理できませんでした。他のアプリを閉じてから再試行してください。';
  }
  // 対処ヒントが取れたら必ず付ける。ヒント表の分岐は「別のプロセスが使用中」
  // 「アクセスが拒否」など日本語のOS/PowerShellエラーを狙って書かれているので、
  // 日本語を理由にここより手前で打ち切ってはいけない。
  if (!hint) return raw;
  if (hasJapanese && !isTechnical) return hint + '\n(' + raw + ')';
  return hint + '\n(元のエラー: ' + raw + ')';
}

function hasUsefulDetail(detail) {
  if (detail == null) return false;
  if (Array.isArray(detail)) return detail.length > 0;
  if (typeof detail === 'object') {
    return Object.values(detail).some(v => Array.isArray(v) ? v.length > 0 : (v !== undefined && v !== null && String(v) !== ''));
  }
  return String(detail) !== '';
}
function showMessage(type, title, message, detail, actions=[], autoHideMs=null) {
  const box = $('notice');
  if (noticeTimer) { clearTimeout(noticeTimer); noticeTimer = null; }
  if (type === 'danger') {
    if (box) box.classList.add('hidden');
    showErrorPanel(title, message, detail || message, actions);
    return;
  }
  if (!box) return;
  box.className = `notice ${type || ''}`;
  box.classList.remove('hidden','fade-out');
  $('notice-title').textContent = title || '';
  $('notice-message').style.whiteSpace = 'pre-line';
  $('notice-message').textContent = message || '';
  const actionBox = $('notice-actions');
  if (actionBox) {
    actionBox.innerHTML = '';
    for (const action of asArray(actions)) {
      if (!action?.label) continue;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `btn ${action.primary ? 'primary' : 'secondary'}`;
      btn.textContent = action.label;
      btn.addEventListener('click', async () => {
        const previousTitle = $('notice-title')?.textContent || '';
        const previousMessage = $('notice-message')?.textContent || '';
        const siblings = Array.from(actionBox.querySelectorAll('button'));
        siblings.forEach(b => { b.disabled = true; });
        btn.classList.add('busy');
        try {
          if (action.view) setActiveView(action.view);
          if (typeof action.handler === 'function') await action.handler();
        } finally {
          siblings.forEach(b => { b.disabled = false; });
          btn.classList.remove('busy');
        }
        if (($('notice-title')?.textContent || '') === previousTitle && ($('notice-message')?.textContent || '') === previousMessage) hideMessage();
      });
      actionBox.appendChild(btn);
    }
  }
  if (hasUsefulDetail(detail)) log(title || 'message', detail);
  if (type === 'ok') hideErrorPanel();
  const delay = autoHideMs ?? (type === 'warn' ? 8000 : 4000);
  if (delay > 0) noticeTimer = setTimeout(hideMessage, delay);
}

function hideMessage() {
  const box = $('notice');
  if (!box || box.classList.contains('hidden')) return;
  box.classList.add('fade-out');
  setTimeout(() => box.classList.add('hidden'), 180);
}

function extractErrorItems(detail, fallbackMessage='') {
  const items = [];
  const push = (name, msg, raw) => {
    const message = userFriendlyError(raw?.userError || msg || raw?.message || raw?.detail || fallbackMessage || '処理できませんでした。');
    const label = String(name || raw?.workbookName || raw?.displayName || raw?.fileName || raw?.relativePath || raw?.workbookId || '').trim();
    items.push({label, message});
  };
  if (Array.isArray(detail)) {
    for (const d of detail) {
      if (!d) continue;
      if (d.ok === false || d.error || d.userError || d.message || d.detail) push('', d.userError || d.error || d.message || d.detail, d);
      else if (Array.isArray(d.errors)) d.errors.forEach(e => push('', e.userError || e.error || e.message || e.detail, e));
    }
  } else if (detail && typeof detail === 'object') {
    if (Array.isArray(detail.errors)) detail.errors.forEach(e => push('', e.userError || e.error || e.message || e.detail, e));
    else if (Array.isArray(detail.results)) detail.results.forEach(e => { if (e.ok === false || e.error || e.userError) push('', e.userError || e.error || e.message || e.detail, e); });
    else if (detail.userError || detail.error || detail.message || detail.detail) push('', detail.userError || detail.error || detail.message || detail.detail, detail);
  } else if (fallbackMessage) {
    push('', fallbackMessage, null);
  }
  return items.filter(Boolean).slice(0, 8);
}
function showErrorPanel(title, summary, detail, actions=[]) {
  const panel = $('error-panel');
  if (!panel) return;
  const friendlySummary=userFriendlyError(summary);
  const items = extractErrorItems(detail, summary).filter(item=>String(item.message||'').trim()!==String(friendlySummary||'').trim());
  // Reveal before writing: a display:none region is not rendered, so screen
  // readers never picked up the text change and every failure was announced
  // as silence.
  panel.classList.remove('hidden');
  $('error-title').textContent = title || 'エラーがあります';
  $('error-summary').textContent = friendlySummary || '内容を確認してください。';
  const list=$('error-list');
  list.innerHTML = items.map(i => `<li>${i.label ? `<strong>${escapeHtml(i.label)}</strong>：` : ''}${escapeHtml(i.message)}</li>`).join('');
  list.classList.toggle('hidden',items.length===0);
  const actionBox = $('error-actions');
  if (actionBox) {
    actionBox.innerHTML = '';
    for (const action of asArray(actions)) {
      if (!action?.label) continue;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `btn ${action.primary ? 'primary' : 'secondary'}`;
      btn.textContent = action.label;
      btn.addEventListener('click', async () => {
        const previousTitle = $('error-title')?.textContent || '';
        const previousSummary = $('error-summary')?.textContent || '';
        const siblings = Array.from(actionBox.querySelectorAll('button'));
        siblings.forEach(b => { b.disabled = true; });
        btn.classList.add('busy');
        try {
          if (action.view) setActiveView(action.view);
          if (typeof action.handler === 'function') await action.handler();
        } finally {
          siblings.forEach(b => { b.disabled = false; });
          btn.classList.remove('busy');
        }
        if (($('error-title')?.textContent || '') === previousTitle && ($('error-summary')?.textContent || '') === previousSummary) hideErrorPanel();
      });
      actionBox.appendChild(btn);
    }
  }
  // Pull the viewport and focus to the panel; it sits at the top of <main> and
  // was routinely off-screen when a failure happened deep in the page board.
  panel.scrollIntoView({block:'nearest'});
  panel.focus({preventScroll:true});
}
function hideErrorPanel() {
  const panel = $('error-panel');
  if (panel) panel.classList.add('hidden');
}

function log(message, obj) {
  const time = new Date().toLocaleTimeString();
  const detail = obj ? `\n${typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2)}` : '';
  const target = $('log');
  if (target) target.textContent = `[${time}] ${message}${detail}\n\n` + target.textContent;
}

async function api(path, options = {}) {
  const headers = Object.assign({'X-ReportBinder-Token': token}, options.headers || {});
  const hasBody = Object.prototype.hasOwnProperty.call(options, 'body');
  const rawBody = hasBody ? options.body : undefined;
  const isBinaryBody = hasBody && (rawBody instanceof FormData || rawBody instanceof Blob);
  if (hasBody && !isBinaryBody) headers['Content-Type'] = 'application/json';
  let res;
  try{
    res = await fetch(withToken(path), {
      method: options.method || 'GET',
      headers,
      body: hasBody ? (isBinaryBody ? rawBody : JSON.stringify(rawBody ?? {})) : undefined,
      keepalive: !!options.keepalive,
      signal: options.signal,
      cache: 'no-store'
    });
  }catch(e){
    // 中止(AbortController)は呼び出し側が扱うので、そのまま通す。
    if(e && e.name === 'AbortError') throw e;
    // それ以外の fetch の失敗は、ほぼ常に「ローカルサーバーが終了している」。
    // 30分の無操作やスリープで自動終了するため日常的に起きるが、既定では
    // 英語の "Failed to fetch" しか出ず、戻り方が画面のどこにも無かった。
    throw new Error(SERVER_GONE_MESSAGE);
  }
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); } catch { data = {ok: false, error: text}; }
  if (!res.ok || data.ok === false) {
    const err = new Error(data.error || `HTTP ${res.status}`);
    err.detail = data.detail || text;
    err.payload = data;
    err.status = res.status;
    err.code = String(data.code || '');
    throw err;
  }
  return data;
}



function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
// 進捗パネル全体をライブ領域にすると、0.85秒ごとのポーリング更新が読み上げキューを
// 埋め尽くし、他の要素を読めなくする。節目だけを専用のライブ領域へ流す。
let lastProgressAnnounceKey = '';
function announceProgressMilestone(pct, status, terminal, total, done) {
  const box = $('progress-announce');
  if (!box) return;
  let key = '', text = '';
  if (terminal) {
    key = `terminal:${status}`;
    text = status === 'cancelled' ? 'PDF作成を中止しました。'
      : status === 'completed' ? 'PDF作成が完了しました。'
      : (status === 'completed-with-errors' || status === 'failed') ? 'PDF作成が終わりました。確認が必要な項目があります。'
      : 'PDF作成が完了しました。';
  } else {
    const step = Math.floor(Math.max(0, pct) / 25) * 25;
    key = `step:${step}`;
    text = step <= 0
      ? 'PDF作成を開始しました。'
      : (total ? `PDF作成中、${step}パーセント。${Math.min(done, total)} / ${total}件。` : `PDF作成中、${step}パーセント。`);
  }
  if (key === lastProgressAnnounceKey) return;
  lastProgressAnnounceKey = key;
  box.textContent = text;
}
function updateProgressPanel(job) {
  const panel = $('progress-panel');
  if (!panel || !job) return;
  job = Object.assign({}, activeRenderJobStatus || {}, job);
  activeRenderJobStatus = job;
  let pct = Math.max(0, Math.min(100, Number(job.percent || 0)));
  const status = String(job.status || '');
  if (pct <= 0 && ['queued','running','waiting'].includes(status)) pct = status === 'queued' ? 1 : 3;
  const total = Math.max(0, Number(job.total || 0));
  const done = Math.max(0, Number(job.completed || 0) + Number(job.failed || 0));
  const checking = status === 'waiting' || status === 'missing' || job.transient || job.transientMissing;
  $('progress-title').textContent = status === 'cancelled'
    ? 'PDF作成を中止しました'
    : status === 'completed'
    ? 'PDF作成が完了しました'
    : (status === 'completed-with-errors' || status === 'failed'
      ? 'PDF作成を確認してください'
      : (checking ? 'PDF作成を確認中' : 'PDF作成中'));
  panel.style.setProperty('--progress', `${pct}%`);
  $('progress-percent').textContent = `${pct}%`;
  const countLabel = $('progress-count-label');
  if (countLabel) countLabel.textContent = total ? `${Math.min(done, total)} / ${total}件` : '';
  $('progress-fill').style.width = `${pct}%`;
  const currentSource = getWorkbook(job.currentWorkbookId);
  const target = job.currentSheet ? `${job.currentWorkbookName || ''} / ${sourceUnitReference(currentSource, job.currentSheet)}` : (job.currentWorkbookName || '');
  $('progress-message').textContent = job.message || (target ? `${target} を処理しています。` : '処理しています。');
  const track = panel.querySelector('.progress-track');
  const terminal = pct >= 100 || ['completed','completed-with-errors','failed','cancelled'].includes(status);
  if (track) {
    // 不定状態では値を公開しない（支援技術は「進行中」として扱う）。
    if ((checking || pct <= 0) && !terminal) track.removeAttribute('aria-valuenow');
    else track.setAttribute('aria-valuenow', String(pct));
    track.setAttribute('aria-valuetext', total ? `${Math.min(done, total)} / ${total}件 (${pct}%)` : `${pct}%`);
  }
  announceProgressMilestone(pct, status, terminal, total, done);
  const cancelButton = $('progress-cancel');
  if (cancelButton) {
    const requested = activeRenderCancelRequested || !!job.cancelRequested;
    cancelButton.hidden = !activeRenderJobId || terminal;
    cancelButton.disabled = requested;
    cancelButton.textContent = requested ? '中止を受け付けました' : 'PDF作成を中止';
  }
  panel.classList.toggle('indeterminate', (checking || pct <= 0) && !terminal);
  panel.classList.remove('hidden');
}
function hideProgressPanel() {
  const panel = $('progress-panel');
  if (panel) panel.classList.add('hidden');
  lastProgressAnnounceKey = '';
  const box = $('progress-announce');
  if (box) box.textContent = '';
}
function normalizeRenderJobFromStart(started) {
  if (!started) return null;
  const candidates = [started.job, started.renderJob, started.result, started.data, started]
    .flatMap(x => Array.isArray(x) ? x : [x])
    .filter(Boolean);
  for (const job of candidates) {
    if (typeof job !== 'object') continue;
    const jobId = String(job.jobId || job.JobId || job.id || started.jobId || started.JobId || '').trim();
    if (jobId) return Object.assign({}, job, {jobId});
  }
  const direct = String(started.jobId || started.JobId || '').trim();
  if (direct) return {jobId: direct, status: started.status || 'queued', percent: started.percent || 0, message: started.message || 'PDF作成を開始しました。'};
  return null;
}
function normalizeRenderJobId(value) {
  const text = String(value || '').trim();
  const match = text.match(/job_\d{8}_\d{6}_[0-9a-fA-F]{8}/);
  return match ? match[0].toLowerCase() : '';
}
function assertJobId(job, raw) {
  const jobId = normalizeRenderJobId(job?.jobId || job?.JobId || job?.id || raw?.jobId || raw?.JobId || raw?.id || '');
  if (jobId) return jobId;
  const err = new Error('PDF作成を開始できませんでした。もう一度「PDF作成」を押してください。');
  err.detail = raw || job || 'ジョブIDが画面へ返りませんでした。';
  console.warn('ReportBinder render start response without jobId', raw || job);
  throw err;
}
function isTransientRenderStatusError(e) {
  const parts = [e?.message, e?.detail, e?.payload?.message, e?.payload?.error, e?.payload?.status, e?.payload?.jobId]
    .map(v => typeof v === 'string' ? v : (v ? JSON.stringify(v) : ''));
  const text = parts.join('\n');
  return /アクセスが拒否|access is denied|access denied|UnauthorizedAccess|PDF作成ジョブが見つかりません|status"?\s*:?\s*"?missing|missing|別のプロセス|being used/i.test(text);
}
async function waitForRenderJob(jobId) {
  jobId = assertJobId({jobId}, {jobId});
  let last = null;
  let unchangedCount = 0;
  let lastUpdated = '';
  let transientStatusErrors = 0;
  let transientStatuses = 0;
  while (true) {
    // Send the job id in the JSON body instead of relying only on query-string parsing.
    // This prevents empty / malformed jobId polling when the URL is rewritten or a browser
    // keeps an old query string around. The server still accepts GET for compatibility.
    try {
      last = await api('/api/jobs/status', {method:'POST', body:{jobId}});
      transientStatusErrors = 0;
    } catch (e) {
      if (isTransientRenderStatusError(e) && transientStatusErrors < 60) {
        transientStatusErrors += 1;
        updateProgressPanel(Object.assign({status:'waiting', transient:true, percent: last?.percent || 0, total: last?.total || 0, completed: last?.completed || 0, failed: last?.failed || 0}, last || {}, {message:'PDF作成の進捗を確認しています。'}));
        await sleep(900);
        continue;
      }
      throw e;
    }
    updateProgressPanel(last);
    const status = String(last.status || '');
    if (['completed','completed-with-errors','failed','cancelled'].includes(status)) return last;
    if (status === 'waiting' || status === 'missing' || last.transient || last.transientMissing) {
      transientStatuses += 1;
      if (transientStatuses > 70) {
        const err = new Error('変換の進み具合を確認できません。処理が続いている場合があります。1分ほど待ってから「更新確認」を押してください。');
        err.detail = last;
        throw err;
      }
      await sleep(900);
      continue;
    }
    transientStatuses = 0;
    const updated = String(last.updatedAt || '');
    if (updated && updated === lastUpdated) unchangedCount += 1;
    else { unchangedCount = 0; lastUpdated = updated; }
    if (unchangedCount > 600) {
      const err = new Error('PDF作成の進捗が更新されていません。Excelの確認ダイアログが出ていないか、またはExcelが応答していない可能性があります。');
      err.detail = last;
      throw err;
    }
    await sleep(850);
  }
}

async function cancelActiveRenderJob() {
  const jobId = normalizeRenderJobId(activeRenderJobId);
  if (!jobId || activeRenderCancelRequested) return;
  const accepted = await confirmAction({
    title:'PDF作成を中止しますか？',
    message:'未処理の原稿を残して、PDF作成ジョブを停止します。',
    detail:'現在処理中の原稿は、安全に完了した後で停止します。完了済みの変換PDFとページ構成は保持されます。',
    confirmLabel:'PDF作成を中止',
    danger:true
  });
  if (!accepted) return;
  activeRenderCancelRequested = true;
  updateProgressPanel({status:'running',cancelRequested:true,percent:0,message:'中止を受け付けています。'});
  try {
    const response = await api('/api/jobs/cancel',{method:'POST',body:{jobId}});
    updateProgressPanel({status:String(response.status||'running'),cancelRequested:true,percent:0,message:response.message||'中止を受け付けました。現在処理中の原稿が終わると停止します。'});
  } catch (error) {
    activeRenderCancelRequested = false;
    showMessage('danger','PDF作成を中止できませんでした',userFriendlyError(error.message||'中止要求を送信できませんでした。'),error.detail||error);
  }
}

function apiUrl(path, params={}) {
  const url = new URL(withToken(path), window.location.origin);
  for (const [k, v] of Object.entries(params || {})) {
    if (v !== undefined && v !== null && String(v) !== '') url.searchParams.set(k, String(v));
  }
  return url.toString();
}

async function fetchPdfObjectUrl(path, params={}) {
  const postPdf = (path === '/api/file' || path === '/api/final/file' || path === '/api/v2/outputs/file');
  const headers = {'X-ReportBinder-Token': token, 'Accept': 'application/pdf,application/json'};
  if (postPdf) headers['Content-Type'] = 'application/json';
  let res;
  try{
    res = await fetch(apiUrl(path, postPdf ? {} : params), {
      method: postPdf ? 'POST' : 'GET',
      cache: 'no-store',
      headers,
      body: postPdf ? JSON.stringify(params || {}) : undefined
    });
  }catch(e){
    if(e && e.name === 'AbortError') throw e;
    throw new Error(SERVER_GONE_MESSAGE);
  }
  const contentType = (res.headers.get('content-type') || '').toLowerCase();
  if (!res.ok || contentType.includes('application/json')) {
    let detail = null;
    // The raw status code told the user nothing actionable; keep it in detail.
    let message = `PDFを取得できませんでした。時間をおいて、もう一度お試しください。改善しない場合はReportBinderを開き直してください。（状態 ${res.status}）`;
    try {
      detail = await res.json();
      message = detail.error || detail.message || message;
    } catch {
      try { message = await res.text() || message; } catch {}
    }
    const err = new Error(message);
    err.detail = detail;
    throw err;
  }
  const blob = await res.blob();
  if (!blob || blob.size === 0) throw new Error('PDFファイルが空です。PDF作成をやり直してください。');
  const pdfBlob = blob.type === 'application/pdf' ? blob : new Blob([blob], {type:'application/pdf'});
  const objectUrl = URL.createObjectURL(pdfBlob);
  pdfObjectUrls.add(objectUrl);
  return objectUrl;
}

function revokePdfObjectUrl(url) {
  if (url && pdfObjectUrls.has(url)) {
    try { URL.revokeObjectURL(url); } catch {}
    pdfObjectUrls.delete(url);
  }
}

function clearPreviewObjectUrl() {
  if (currentPreviewObjectUrl) {
    revokePdfObjectUrl(currentPreviewObjectUrl);
    currentPreviewObjectUrl = '';
  }
}


async function loadFilesSilently() {
  if (!configured()) return [];
  const data = await api('/api/submission-files');
  availableFiles = asArray(data.files);
  availableFilesScannedAt = String(data.scannedAt || '');
  renderFolderOverview();
  selectedFiles = new Set([...selectedFiles].filter(r => visibleUnregisteredFiles(availableFiles).some(f => normPath(f.relativePath) === normPath(r))));
  return availableFiles;
}
async function scanUpdatesSilently({withFiles=false}={}) {
  // A background refresh during a build re-renders the final cards and hands
  // back a fresh, enabled build button, which let a second job start on top of
  // the first and orphaned the running job's cancel button.
  if (!configured() || updateMonitorInFlight || renderJobActive || activeFinalJobId) return;
  updateMonitorInFlight = true;
  try {
    await api('/api/scan-updates', {method:'POST', body:{force:false}});
    state = normalizeStatePayload(await api('/api/state'));
    if (withFiles && !shouldPreserveEditorDom()) await loadFilesSilently();
    renderAll({preserveEditors:true});
  } catch (e) {
    log('更新確認でエラー', e.detail || e.stack || e.message);
  } finally { lastUpdateScanAt=Date.now(); updateMonitorInFlight = false; }
}

function startUpdateMonitor() {
  if (updateMonitorTimer) clearInterval(updateMonitorTimer);
  // 共有フォルダー全走査は5分間隔。必要なときは画面の更新ボタンで即時確認できる。
  updateMonitorTimer = setInterval(() => scanUpdatesSilently({withFiles:true}), 300000);
  if (!updateVisibilityBound) {
    document.addEventListener('visibilitychange', () => {
      if (!document.hidden && Date.now()-lastUpdateScanAt>=300000) scanUpdatesSilently({withFiles:true});
    });
    updateVisibilityBound = true;
  }
}

let finalReadinessInFlight = null;
// Repaint guards. Entering the final screen re-rendered every block even when
// the data was identical, so the page visibly redrew ~250ms after it appeared.
// Keep signatures off the DOM (a dataset attribute would carry the whole HTML).
const renderSignatures=new WeakMap();
function setHtmlIfChanged(el, html){
  if(!el)return false;
  if(renderSignatures.get(el)===html)return false;
  el.innerHTML=html; renderSignatures.set(el,html); return true;
}
function setTextIfChanged(el, text){
  if(!el||el.textContent===text)return false;
  el.textContent=text; return true;
}
function setClassIfChanged(el, cls){
  if(!el||el.className===cls)return false;
  el.className=cls; return true;
}
let finalPanelsInFlight=null;
function finalReadinessKey(){
  const pack=activePackRecord();
  if(!pack?.packId)return '';
  return pack.category?activePreset:String(pack.packId);
}
// True once readiness for the active pack is actually in state. Used instead of
// a "loaded" flag: an early call during boot can bail before fetching anything,
// and a flag set up-front would then suppress every later attempt, leaving the
// screen showing 0 ページ forever.
function finalReadinessLoaded(){
  const key=finalReadinessKey();
  if(!key)return true;
  return !!state?.finalReadiness?.[key];
}
// Load everything the final screen needs, then paint once.
async function loadFinalPanels(options = {}){
  if(!finalPanelsInFlight){
    finalPanelsInFlight=Promise.allSettled([
      loadFinalReadiness({render:false}),
      loadFinalArchives(),
      loadPackReview({render:false})
    ]).finally(()=>{finalPanelsInFlight=null;});
  }
  await finalPanelsInFlight;
  if(options.render===false) return;
  if(activeView==='final'){renderGlobalHeader();renderNavBadges();renderFinalOverview();renderVolumeLinks();}
}
async function loadFinalReadiness(options = {}){
  if(!state||!configured())return;
  const pack=activePackRecord(),preset=activePreset,key=pack?.category?preset:String(pack?.packId||activePackId);
  if(!pack?.packId)return;
  if(finalReadinessInFlight)return finalReadinessInFlight;
  const request=pack?.category?api(`/api/final/readiness?category=${encodeURIComponent(preset)}`):api(`/api/v2/outputs/readiness?packId=${encodeURIComponent(key)}`);
  finalReadinessInFlight=request.then(r=>{
    state.finalReadiness=state.finalReadiness||{};
    const volumes=r.volumes||Object.fromEntries(Object.entries(r.targets||{}).map(([targetId,ready])=>[targetVolume(targetId),ready]));
    state.finalReadiness[key]={volumes};
    finalPreflightCheckedAt=new Date().toLocaleString('sv-SE',{timeZone:'Asia/Tokyo'});
    if(options.render!==false&&activeView==='final'&&String(activePackId)===String(pack?.packId||activePackId)){renderGlobalHeader();renderNavBadges();renderFinalOverview();renderVolumeLinks();}
  }).catch(e=>log('提出用PDF状態の確認でエラー',e.detail||e.message)).finally(()=>{finalReadinessInFlight=null;});
  return finalReadinessInFlight;
}

async function refresh(options = {}) {
  try {
    // finalReadiness is filled in by loadFinalReadiness(), not by /api/state.
    // Replacing state wholesale dropped it, so the final screen painted
    // "0ページ / 出力できません" and corrected itself once the refetch landed -
    // a guaranteed wrong-then-right flash on every refresh.
    const carriedReadiness = state?.finalReadiness;
    state = normalizeStatePayload(await api('/api/state'));
    if (carriedReadiness) state.finalReadiness = Object.assign({}, carriedReadiness, state.finalReadiness || {});
    const reloadHistory = activeView === 'history' && historyPanelsInitialized;
    snapshotHistoryResponseCache.clear();
    // render:false lets a caller fold several sequential loads into one paint.
    if (options.render !== false) renderAll();
    if (reloadHistory) void loadHistoryPanels({force:true});
    return state;
  } catch (e) {
    showMessage('danger', '状態を読み込めません', userFriendlyError(e.message), e.detail || e.stack || e.message);
    throw e;
  }
}
async function refreshAndLoad(btn) {
  await runBusy(btn, async () => {
    // 更新検知の失敗を握りつぶして「更新しました」と言うと、実際には古い判定のままなのに
    // 利用者は確認済みだと誤解する。読み込み自体は続けたうえで、確認できなかったことを伝える。
    let scanError = null;
    if (configured()) {
      try { await api('/api/scan-updates', {method:'POST', body:{force:false}}); }
      catch (e) { scanError = e; }
    }
    await refresh();
    if (configured()) await loadFiles(null);
    if (scanError) {
      showMessage('warn', '元原稿の更新を確認できませんでした',
        `${userFriendlyError(scanError.message)}\n登録状態は読み込み直しました。少し待ってからもう一度「更新」を押してください。`,
        scanError.detail || scanError.stack || scanError.message);
      return;
    }
    showMessage('ok', '画面を更新しました', '提出フォルダーと登録状態を読み込み直しました。');
  }, false);
}

async function refreshAll(btn=null) {
  await runBusy(btn, async () => {
    await refresh();
    if (configured()) await loadFiles(null);
    showMessage('ok', '更新しました', '最新の状態を読み込みました。');
  }, false);
}

async function scanAndRefresh(btn) {
  await runBusy(btn, async () => {
    if (!configured()) { await refresh(); return; }
    showMessage('', '原稿の更新を確認しています', '登録済み原稿の保存日時・サイズ・内容を確認します。');
    const r = await api('/api/scan-updates', {method:'POST', body:{force:true}});
    state = normalizeStatePayload(await api('/api/state'));
    if (configured()) await loadFiles(null);
    selectedWorkbooks.clear();
    lastWorkbookRangeAnchor = '';
    renderAll();
    const n = r.result?.changedWorkbookIds?.length || 0;
    const scanned = Number(r.result?.scanned || 0);
    const detail = n
      ? `${n}件の原稿をPDF変換してください。`
      : (scanned ? `${scanned}件を確認しました。` : '登録済み原稿はありません。');
    showMessage(n ? 'warn' : 'ok', n ? '原稿に更新があります' : '更新はありません', detail, r.result || null);
  }, false);
}

function renderAll(options={}) {
  const preserve = !!options.preserveEditors && shouldPreserveEditorDom();
  const appTitle = $('app-title');
  if (appTitle) appTitle.textContent = 'ReportBinder';
  setActiveView(activeView, {noScroll:true});
  updatePresetTabs();
  renderPackRequiredState();
  renderPathInputs();
  renderGlobalHeader();
  renderStepBar();
  renderNavBadges();
  renderSummary();
  renderDashboardOverview();
  renderFolderOverview();
  renderPageOverview();
  renderFinalOverview();
  renderAutoStatus();
  renderVolumeLinks();
  updateBulkSelectionLabel();
  if (!preserve) {
    renderFileList(availableFiles);
    renderWorkbooks();
    renderPages();
  }
  if (state.structureLoadError) {
    showMessage('danger', '管理データを読み込めません', userFriendlyError(state.structureLoadError), state.structureLoadError);
    return;
  }
  const errors = Number(state.summary?.renderErrors || 0);
  if (!configured()) hideErrorPanel();
  else if (errors > 0) {
    const recent = asArray(state.recentErrors);
    const first = recent[0];
    showMessage('danger', 'PDF作成でエラーがあります', first?.message || `${errors}件の原稿をPDF変換できません。`, recent);
  } else hideErrorPanel();
}

function renderPathInputs() {
  const submission = $('submissionDir'); if (submission) submission.value = state?.paths?.submissionDir || '';
  const sourceFolder = $('source-folder-path'); if (sourceFolder) { sourceFolder.textContent=state?.paths?.submissionDir||'—';sourceFolder.title=state?.paths?.submissionDir||''; }
  const choose=$('change-source-folder-btn');if(choose){const ready=configured();choose.className=`btn ${ready?'secondary':'primary'}`;choose.textContent=ready?'変更':'フォルダーを選ぶ';}
  const scan=$('scan-btn');if(scan)scan.disabled=!configured();
  const data = $('dataDir'); if (data) data.textContent = state?.paths?.dataDir || '—';
  const output = $('outputDir'); if (output) output.textContent = state?.paths?.outputDir || '—';
}


function renderGlobalHeader() {
  const agg = aggregateFinalState();
  const el = $('global-final-status');
  const hasPack=!!activePackRecord();
  document.body.classList.toggle('has-active-pack',hasPack);
  document.body.classList.toggle('is-configured',configured());
  $('source-first-run')?.classList.toggle('hidden',configured());
  if (!el) return;
  el.classList.toggle('hidden',!hasPack);
  if(!hasPack){el.innerHTML='';el.removeAttribute('aria-label');return;}
  const unresolved=unresolvedPageCount();
  const statusText=unresolved?`未振り分け ${unresolved}ページ`:agg.text;
  el.className = 'global-status';
  if (agg.state === 'blocked') el.classList.add('danger');
  else if (unresolved || ['needs-rebuild','output-missing'].includes(agg.state)) el.classList.add('attention');
  const dot = unresolved || ['blocked','needs-rebuild','output-missing'].includes(agg.state) ? '<span class="status-dot"></span>' : '';
  // Unassigned pages are a step-2 problem, so point the shortcut at step 2
  // instead of sending people to the final screen to be bounced back.
  const label = unresolved ? 'ページ構成' : '提出用PDF';
  el.dataset.viewShortcut = unresolved ? 'pages' : 'final';
  el.innerHTML = `${dot}<span class="global-status-label">${label}</span>${escapeHtml(statusText.replace(/^提出用PDF(?:は|の)?/,'').trim())}`;
  el.setAttribute('aria-label',statusText.startsWith(label)?statusText:`${label} ${statusText}`);
}
const PACK_FREE_VIEWS=['dashboard','excel'];
function renderPackRequiredState(){
  const hasPack=!!activePackRecord();
  document.querySelectorAll('[data-view-nav]').forEach(button=>{
    const view=String(button.dataset.viewNav||'');
    if(!VIEW_NAMES[view])return;
    const locked=!PACK_FREE_VIEWS.includes(view)&&!hasPack;
    // Left focusable on purpose: a real `disabled` button drops out of the tab
    // order and swallows its own title, so keyboard users saw a dimmed item
    // with no way to learn why it was unavailable.
    button.disabled=false;
    button.classList.toggle('nav-locked',locked);
    button.setAttribute('aria-disabled',String(locked));
    button.title=locked?'先に原稿フォルダーを設定し、資料パックを作成してください':'';
  });
  if(!hasPack&&!PACK_FREE_VIEWS.includes(activeView))setActiveView('dashboard',{noScroll:true});
}
function explainPackRequired(){
  setActiveView('dashboard');
  showMessage('warn','先に資料パックを作成してください','原稿フォルダーを選び、今回まとめる資料パックに名前を付けると、ページ構成と提出用PDFに進めます。',null,[{label:'原稿フォルダーを選ぶ',primary:true,handler:()=>setActiveView('excel')}],0);
}
function setStepState(id, kind, fallback) {
  const el=$(id); if(!el) return;
  el.className=`step-state ${kind || ''}`;
  el.textContent = kind === 'complete' ? '✓' : (kind === 'attention' ? '!' : fallback);
}
function renderStepBar() {
  const workbooks=workbooksForActivePreset(); const pages=pagesForActivePreset();
  const needs=workbooks.filter(w=>!isLatestPdfWorkbook(w)).length;
  const currentMap={excel:'step-excel-state',pages:'step-pages-state',final:'step-final-state'};
  setStepState('step-excel-state', configured()&&workbooks.length&&needs===0?'complete':'attention','1');
  setStepState('step-pages-state', pages.some(p=>p.enabled!==false && String(p.volume||'none')!=='none')?'complete':(pages.length?'attention':''),'2');
  // activeReadinessItems() counts targets that HAVE pages, not problems, so this
  // used to show "!" next to "出力準備が整っています". Warn only on real blockers.
  const finalBlockers=activeReadinessItems().flatMap(item=>asArray(item.ready.blockers)).length;
  setStepState('step-final-state', finalIsComplete()?'complete':(finalBlockers?'attention':''),'3');
  document.querySelectorAll('[data-progress-view]').forEach(btn=>{
    const view=btn.dataset.progressView; btn.classList.toggle('current',view===activeView);
    if(view===activeView)btn.setAttribute('aria-current','step');else btn.removeAttribute('aria-current');
    const st=$(currentMap[view]); if(st) st.classList.toggle('current', view===activeView);
  });
}
function renderNavBadges() {
  const needs=workbooksForActivePreset().filter(w=>!isLatestPdfWorkbook(w)).length;
  const unresolved=unresolvedPageCount();
  const excel=$('nav-excel-count'); if(excel){excel.textContent=needs;excel.classList.toggle('hidden',needs===0);}
  const pages=$('nav-pages-count'); if(pages){pages.textContent=`未振り分け ${unresolved}`;pages.classList.toggle('hidden',unresolved===0);}
  const agg=aggregateFinalState(); const dot=$('nav-final-dot'); if(dot) dot.classList.toggle('hidden',!['blocked','needs-rebuild','output-missing'].includes(agg.state));
}

function renderSummary() {
  const summaryEl = $('summary'); if (!summaryEl) return;
  const workbooks=workbooksForActivePreset(); const pages=pagesForActivePreset();
  const needsPdf=workbooks.filter(w=>!isLatestPdfWorkbook(w)).length;
  const cards=[
    {label:'登録済み原稿',value:workbooks.length,unit:'件'},
    {label:'変換PDFは最新',value:workbooks.length-needsPdf,unit:'件'},
    {label:'変換PDF作成が必要',value:needsPdf,unit:'件'},
    {label:'総ページ数',value:pages.length,unit:'ページ'}
  ];
  const scopeLabel=activePackRecord()?presetLabel():'今回まとめる一式の作成前';
  summaryEl.innerHTML=cards.map(c=>`<article class="stat-card"><div class="stat-label">${escapeHtml(c.label)}</div><div class="stat-value">${c.value}<span class="stat-unit">${escapeHtml(c.unit)}</span></div><div class="stat-sub">${escapeHtml(scopeLabel)}</div></article>`).join('');
}


function renderDashboardOverview() {
  const workbooks=workbooksForActivePreset(); const pages=pagesForActivePreset();
  const needsPdf=workbooks.filter(w=>!isLatestPdfWorkbook(w)).length;
  const readiness=activeReadinessItems();
  const blockers=readiness.flatMap(x=>asArray(x.ready.blockers));
  const unassigned=unresolvedPageCount(); const agg=aggregateFinalState();
  const onboarding=$('dashboard-onboarding');if(onboarding)onboarding.classList.toggle('hidden',configured()&&!!activePackRecord()&&workbooks.length>0);
  const hasPack=!!activePackRecord();
  $('summary')?.classList.toggle('hidden',!hasPack);
  $('dashboard-pack-progress')?.classList.toggle('hidden',!hasPack);
  $('dashboard-recent')?.classList.toggle('hidden',!hasPack);
  const title=$('dashboard-next-title'), caption=$('dashboard-next-caption'), action=$('dashboard-action-btn');
  if(title&&caption&&action){
    if(!configured()){title.textContent='はじめる：原稿が入ったフォルダーを選ぶ';caption.textContent='フォルダー内のExcel・Word・PowerPoint・PDFを見つけます。原稿は移動・削除しません。';action.textContent='はじめる';action.dataset.dashboardAction='folder';}
    else if(!activePackRecord()){title.textContent='今回まとめる一式に名前を付ける';caption.textContent='会議名や提出先など、あとで見つけやすい名前を付けます。';action.textContent='名前を付ける';action.dataset.dashboardAction='create-pack';}
    else if(workbooks.length===0){title.textContent='まとめる原稿を登録する';caption.textContent=`「${presetLabel()}」に入れる原稿を選びます。`;action.textContent='原稿を選ぶ';action.dataset.dashboardAction='excel';}
    else if(needsPdf>0){title.textContent=`変換PDFを作成してください（${needsPdf}件）`;caption.textContent='原稿の最新内容をページ構成用のPDFへ変換します。';action.textContent='PDF変換へ';action.dataset.dashboardAction='excel';}
    else if(blockers.length>0){title.textContent='提出用PDFを出力できません';caption.textContent=blockers[0]?.message||'原稿・変換PDFを確認してください。';action.textContent='原稿・変換PDFへ';action.dataset.dashboardAction='excel';}
    else if(unassigned>0){title.textContent=`ページ構成を確認してください（${unassigned}ページ）`;caption.textContent='出力先が未設定のページがあります。';action.textContent='ページ構成へ進む';action.dataset.dashboardAction='pages';}
    else if(['needs-rebuild','output-missing','not-built'].includes(agg.state)){title.textContent=agg.text;caption.textContent='各出力先の状態を確認して出力してください。';action.textContent='提出用PDFへ進む';action.dataset.dashboardAction='final';}
    else{title.textContent='最新の状態です';caption.textContent='提出用PDFとページ構成は最新です。';action.textContent='出力フォルダーを確認';action.dataset.dashboardAction='final';}
  }
  const cat=$('dashboard-category-summary');if(cat)cat.textContent=activePackRecord()?`${presetLabel()} / 登録済み原稿：${workbooks.length}件 / 変換PDFの作成が必要：${needsPdf}件`:'「今回まとめる一式」に、原稿・ページの順番・更新履歴をまとめて保存します。';
  const activity=$('dashboard-activity');if(!activity)return;
  const rows=[...workbooks].sort((a,b)=>String(latestWorkbookStamp(b)).localeCompare(String(latestWorkbookStamp(a)))).slice(0,5).map(w=>{
    if(String(w.status||'')==='render-error')return{text:`${workbookDisplayName(w)} のPDF変換でエラーがあります`,badge:'作成エラー',cls:'danger',time:latestWorkbookStamp(w)};
    if(isLatestPdfWorkbook(w))return{text:`${workbookDisplayName(w)} の変換PDFを作成しました`,badge:'作成済み',cls:'neutral',time:w.lastRenderedAt||latestWorkbookStamp(w)};
    return{text:`${workbookDisplayName(w)} は変換PDF作成が必要です`,badge:'作成が必要',cls:'attention',time:latestWorkbookStamp(w)};
  });
  if(!rows.length)rows.push({text:configured()?(activePackRecord()?`「${presetLabel()}」へ原稿を登録してください`:'今回まとめる一式に名前を付けてください'):'上の「はじめる」から原稿フォルダーを選んでください',badge:'はじめに',cls:'neutral',time:''});
  activity.innerHTML=rows.map(r=>`<div class="activity-row"><span class="activity-text">${escapeHtml(r.text)}</span>${badge(r.badge,r.cls)}<span class="activity-time">${escapeHtml(formatDateTime(r.time))}</span></div>`).join('');
  renderPackProgressDashboard();
}

// Progress and the personal confirmation record are independent now, so the row
// carries both: 完了 says the work is done, this says whether it was eyeballed.
function reviewChip(reviewStatus){
  const status=String(reviewStatus||'draft');
  if(status==='approved')return badge('確認済み','ok');
  if(status==='stale')return badge('確認後に変更あり','attention');
  return badge('未確認','neutral');
}
function renderPackProgressDashboard(){
  const progress=state?.packProgress||{},summary=$('pack-progress-summary'),box=$('pack-progress-list'),toggle=$('pack-progress-attention-only');if(!box)return;if(toggle)toggle.checked=packProgressAttentionOnly;
  if(summary)summary.textContent=Number(progress.totalCount||0)?`全${Number(progress.totalCount||0)}件・完了 ${Number(progress.completeCount||0)}件・要対応 ${Number(progress.attentionCount||0)}件・未提出必須原稿 ${Number(progress.missingRequiredCount||0)}件・期限超過 ${Number(progress.overdueRequiredCount||0)}件・期限7日以内 ${Number(progress.dueSoonRequiredCount||0)}件`:'作成中の一式がここに表示されます。';
  const priority={'overdue-source':0,blocked:1,'missing-source':2,'needs-render':3,unassigned:4,'needs-output':5,'review-stale':6,'review-changes':7,'review-draft':8,'in-review':9,'no-pages':10,'not-started':11,complete:12};let rows=asArray(progress.packs);if(packProgressAttentionOnly)rows=rows.filter(pack=>String(pack.state)!=='complete');rows=[...rows].sort((a,b)=>(priority[String(a.state)]??13)-(priority[String(b.state)]??13)||String(a.displayName||'').localeCompare(String(b.displayName||''),'ja'));
  const stateInfo=value=>({complete:['完了','ok'],'not-started':['未着手','neutral'],'overdue-source':['期限超過','danger'],'missing-source':['必須原稿待ち','danger'],'needs-render':['変換待ち','attention'],unassigned:['未振り分け','attention'],blocked:['出力不可','danger'],'no-pages':['ページ構成待ち','attention'],'needs-output':['出力が必要','attention'],'review-draft':['未確認','neutral'],'in-review':['未確認','neutral'],'review-changes':['未確認','neutral'],'review-stale':['確認後に変更あり','attention']}[String(value)]||['要確認','attention']);
  if(!rows.length){box.innerHTML=packProgressAttentionOnly?'<div class="empty-state">対応が必要な一式はありません。</div>':`<div class="pack-progress-empty"><div><strong>${configured()?'今回まとめる一式はまだありません':'まだ作業は始まっていません'}</strong><span>${configured()?'上の「次にやること」から名前を付けます。':'上の「はじめる」から原稿フォルダーを選びます。'}</span></div></div>`;return;}
  box.innerHTML=rows.map(pack=>{const [status,cls]=stateInfo(pack.state),required=Number(pack.requiredSourceCount||0),submitted=Number(pack.submittedRequiredSourceCount||0),overdue=Number(pack.overdueRequiredSourceCount||0),dueSoon=Number(pack.dueSoonRequiredSourceCount||0),nearestDue=String(pack.nearestRequiredDueDate||''),targets=asArray(pack.targets),action=String(pack.nextAction||'excel'),actionLabel=action==='pages'?'ページ構成':action==='final'?'提出用PDF':'原稿を確認',dueText=overdue?`・期限超過 ${overdue}件`:dueSoon?`・期限7日以内 ${dueSoon}件`:nearestDue?`・次の期限 ${formatDateOnly(nearestDue)}`:'';return `<article class="pack-progress-row ${String(pack.packId||'')===String(activePackId)?'active':''}"><div class="pack-progress-main"><div><strong>${escapeHtml(pack.displayName||pack.packId||'資料パック')}</strong>${badge(status,cls)}${reviewChip(pack.reviewStatus)}</div><span>原稿 ${Number(pack.sourceCount||0)}件${required?`・必須 ${submitted}/${required}`:''}${dueText}・未振り分け ${Number(pack.unassignedPageCount||0)}ページ</span></div><div class="pack-progress-targets">${targets.map(target=>`<span class="pack-target-chip ${escapeAttr(target.displayState||'not-built')}">${escapeHtml(target.displayName||target.targetId)} ${Number(target.pageCount||0)}ページ</span>`).join('')}</div><button class="btn ${String(pack.state)==='complete'?'ghost':'secondary'} compact" type="button" data-pack-progress-open="${escapeAttr(pack.packId||'')}" data-pack-progress-action="${escapeAttr(action)}">${escapeHtml(actionLabel)}</button></article>`;}).join('');
  box.querySelectorAll('[data-pack-progress-open]').forEach(button=>button.addEventListener('click',async()=>{await applyPresetSelection(String(button.dataset.packProgressOpen||''));setActiveView(String(button.dataset.packProgressAction||'excel'));}));
}


function renderFolderOverview() {
  const count=$('folder-file-count'); if(count) count.textContent=`${availableFiles.length || state?.summary?.totalWorkbooks || 0}件`;
  const scan=$('folder-scan-time'); if(scan) scan.textContent=formatDateTime(state?.structure?.updatedAt||'');
  const result=$('folder-scan-result'); if(result){const errors=Number(state?.summary?.renderErrors||0);result.className=`badge ${errors?'danger':'neutral'}`;result.textContent=errors?'要確認':(configured()?'正常':'未設定');}
  renderDiagnosticsResult();
}

function formatCapacity(bytes){const value=Number(bytes);if(!Number.isFinite(value)||value<0)return'—';if(value>=1024**3)return`${(value/1024**3).toFixed(value<10*1024**3?1:0)} GB`;return`${Math.round(value/1024**2)} MB`;}
function renderDiagnosticsResult(){
  const box=$('diagnostics-result');if(!box)return;if(!diagnosticsResult){box.innerHTML='<div class="empty-state">まだ診断していません。ReportBinderをはじめて使うときや、PCの更新後、変換に失敗したときに実行してください。</div>';return;}
  const d=diagnosticsResult,office=d.office||{},runtime=d.runtime||{},storage=d.storage||{};
  const statusLabel=value=>({ready:'利用可能',limited:'一部利用可能',blocked:'修復が必要','not-installed':'未導入',unsupported:'非対応版',unavailable:'起動不可'})[String(value||'')]||'要確認';
  const statusClass=value=>String(value)==='ready'?'neutral':String(value)==='blocked'?'danger':'attention';
  const appRow=app=>{const item=app||{},detail=item.available?`Version ${item.version||'16'}${item.build?` / Build ${item.build}`:''}`:(item.userAction||'利用できません。');return`<div class="diagnostic-row"><div><strong>${escapeHtml(item.displayName||'Office')}</strong><span>${escapeHtml(detail)}</span></div>${badge(statusLabel(item.status),statusClass(item.status))}</div>`;};
  box.innerHTML=`<div class="diagnostics-summary"><div><strong>${escapeHtml(statusLabel(d.status))}</strong><p>${escapeHtml(d.summary||'')}</p></div>${badge(statusLabel(d.status),statusClass(d.status))}</div><div class="diagnostic-list">${appRow(office.excel)}${appRow(office.word)}${appRow(office.powerPoint)}<div class="diagnostic-row"><div><strong>PDF処理エンジン</strong><span>Java・PDFBox・PDF.js</span></div>${badge(runtime.java?.ready&&runtime.pdfbox?.ready&&runtime.pdfjs?.ready?'利用可能':'修復が必要',runtime.java?.ready&&runtime.pdfbox?.ready&&runtime.pdfjs?.ready?'neutral':'danger')}</div><div class="diagnostic-row"><div><strong>ローカル保存容量</strong><span>管理データ ${formatCapacity(storage.dataFreeBytes)} / 出力 ${formatCapacity(storage.outputFreeBytes)}</span></div>${badge(storage.configured?'確認済み':'未設定',storage.configured?'neutral':'attention')}</div></div><p class="caption diagnostics-time">診断日時 ${escapeHtml(formatDateTime(d.capturedAt))}</p>`;
}

async function runSystemDiagnostics(btn){
  await runBusy(btn,async()=>{try{const response=await api('/api/diagnostics/run',{method:'POST',body:{}});diagnosticsResult=response.diagnostics||null;renderDiagnosticsResult();const limited=String(diagnosticsResult?.status||'')!=='ready';showMessage(limited?'warn':'ok',limited?'一部の機能に対応が必要です':'動作環境を確認しました',diagnosticsResult?.summary||'',limited?[diagnosticsResult?.office?.excel?.userAction,diagnosticsResult?.office?.word?.userAction,diagnosticsResult?.office?.powerPoint?.userAction].filter(Boolean):null);}catch(error){showMessage('danger','動作環境を診断できません',userFriendlyError(error.message),error.detail||error.stack||error.message);}});
}



function renderPageOverview() {
  const counts=pageCountsByVolume(); const strip=$('page-summary-strip');
  const summaryItems=[['未振り分け',counts.none,'attention'],...activeTargetDefinitions().map(target=>[`${target.displayName||target.targetId} PDF`,counts.volumes[targetVolume(target.targetId)]||0,''])];
  if(strip)strip.innerHTML=summaryItems.map(([label,n,cls])=>`<div class="page-summary-item ${cls&&n?'attention':''}"><div class="summary-label">${escapeHtml(label)}</div><div><span class="summary-value">${n}</span> ページ</div></div>`).join('');
  const manual=pagesForActivePreset().some(p=>p.orderManual===true); const label=$('sort-state-label'); if(label){label.textContent=manual?'手動':'原稿内の順序';label.classList.toggle('manual',manual);}
  // Step 1 hands off to step 2 with a primary button; step 2 had no equivalent
  // and left people wondering whether the sorting work was finished.
  const assigned=pagesForActivePreset().some(p=>p.enabled!==false&&String(p.volume||'none')!=='none');
  const sortingDone=assigned&&unresolvedPageCount()===0;
  $('pages-next-action')?.classList.toggle('hidden',!sortingDone);
  // Hide the "first, do this" guide once there is nothing left to sort; showing
  // both it and "ページの振り分けが終わりました" made the state unreadable.
  document.querySelector('.assignment-guide')?.classList.toggle('hidden',sortingDone);
  updateBulkSelectionLabel();
}
function renderFinalOverview() {
  const settings=packForPreset()?.settings||{};
  const setValue=(id,value)=>{const el=$(id);if(el)el.value=String(value??'');};
  const setChecked=(id,value)=>{const el=$(id);if(el)el.checked=!!value;};
  setValue('pack-output-pattern',settings.outputFileNamePattern||'{packName}_{targetName}_{yyyyMMdd}.pdf');
  const entries=activeTargetDefinitions().map(target=>({target,volume:targetVolume(target.targetId),ready:volumeReadiness(targetVolume(target.targetId))}));
  const preflight=renderFinalPreflight(entries);
  const unassigned=Number(preflight?.unassigned||0);
  const visibleEntries=entries.filter(entry=>entry.target?.required!==false||Number(entry.ready?.pageCount||0)>0);
  const grid=$('final-target-grid');
  // Skip the rebuild when nothing changed. loadFinalPanels() repaints once after
  // its fetch, and blowing away innerHTML for an identical result made the cards
  // blink for no reason. Same signature trick renderPages() uses.
  const gridHtml=visibleEntries.map(({target,volume,ready})=>`<section class="card final-card" data-final-volume="${escapeAttr(volume)}"><div class="card-head"><div><h3>${escapeHtml(target.displayName||target.targetId)}（提出用PDF）</h3><p><strong class="large-number">${Number(ready.pageCount||0)}</strong> ページ</p></div><svg class="icon card-icon"><use href="#i-file-pdf"/></svg></div><div class="freshness" data-final-state></div><div class="final-path hidden" data-final-path></div><div class="card-actions"><button class="btn secondary" type="button" data-final-build>${unassigned?`${unassigned}ページを除外して${escapeHtml(target.displayName||target.targetId)}だけ出力`:`${escapeHtml(target.displayName||target.targetId)}だけ出力`}</button><button class="btn secondary hidden" type="button" data-final-publish>共有用フォルダーにコピー</button><button class="btn ghost hidden" type="button" data-final-fix>原稿・変換PDFへ →</button><button class="btn secondary hidden" type="button" data-final-reveal>保存先フォルダーを開く</button><a class="btn ghost hidden" href="#" data-final-open>前回出力を開く</a></div></section>`).join('');
  const gridChanged=setHtmlIfChanged(grid,gridHtml);
  const renderFreshness=(ready,card)=>{
    const box=card?.querySelector('[data-final-state]'),btn=card?.querySelector('[data-final-build]'),fix=card?.querySelector('[data-final-fix]');if(!box)return;
    const blockers=asArray(ready.blockers),reasons=asArray(ready.staleReasons).slice(0,3);const display=String(ready.displayState||'not-built');
    let title='提出用PDFは未出力',cls='neutral',body='出力すると、この出力先の提出用PDFを作成します。';
    if(Number(ready.pageCount||0)===0){title='出力ページがありません';body='ページ構成でこの出力先にページを設定してください。';}
    else if(display==='blocked'){title='提出用PDFを出力できません';cls='danger';body=`<ul class="blocker-list">${blockers.slice(0,3).map(b=>`<li>${escapeHtml(b.message||'出力条件を確認してください。')}</li>`).join('')}</ul>`;}
    else if(display==='needs-rebuild'){title='再出力が必要';cls='attention';body=reasons.length?`<ul class="reason-list">${reasons.map(r=>`<li><span>${escapeHtml(r.detail||'入力が変更されました')}</span><time>${escapeHtml(formatDateTime(r.at))}</time></li>`).join('')}</ul>`:'入力が変更されています。';}
    else if(display==='output-missing'){title='前回出力が見つかりません';cls='attention';body='ファイルが削除または移動されています。もう一度出力してください。';}
    else if(unassigned){title=`要確認：未振り分け ${unassigned}ページ`;cls='attention';body='このまま出力すると、未振り分けページはPDFに入りません。';}
    else if(display==='built'){title='提出用PDFは最新です';body=`最終出力 ${escapeHtml(formatDateTime(ready.lastBuiltAt))}`;}
    const dot=cls==='neutral'?'':'<span class="status-dot"></span>';
    setHtmlIfChanged(box,`<div class="freshness-title ${cls}">${dot}${escapeHtml(title)}</div>${typeof body==='string'&&body.startsWith('<')?body:`<p class="caption">${escapeHtml(body)}</p>`}`);
    if(btn)btn.disabled=Number(ready.pageCount||0)===0||blockers.length>0;
    if(fix)fix.classList.toggle('hidden',blockers.length===0||Number(ready.pageCount||0)===0);
  };
  // Bind only when the markup was actually replaced: re-running addEventListener
  // on surviving nodes would stack duplicate handlers and fire a second build.
  for(const entry of visibleEntries){const card=grid?.querySelector(`[data-final-volume="${CSS.escape(entry.volume)}"]`);renderFreshness(entry.ready,card);if(!gridChanged)continue;card?.querySelector('[data-final-build]')?.addEventListener('click',event=>buildVolume(entry.volume,event.currentTarget));card?.querySelector('[data-final-publish]')?.addEventListener('click',event=>publishFinalVolume(entry.volume,event.currentTarget));card?.querySelector('[data-final-fix]')?.addEventListener('click',()=>setActiveView('excel'));card?.querySelector('[data-final-reveal]')?.addEventListener('click',event=>revealFinalVolume(entry.volume,event.currentTarget));card?.querySelector('[data-final-open]')?.addEventListener('click',event=>{event.preventDefault();openFinalVolume(entry.volume,activePreset);});}
  const allBtn=$('build-all-btn'),buildable=entries.map(entry=>entry.ready).filter(ready=>Number(ready.pageCount||0)>0);if(allBtn){const hardProblems=Number(preflight?.hardProblemCount||0);allBtn.disabled=buildable.length===0||hardProblems>0||!!activeFinalJobId;setClassIfChanged(allBtn,`btn ${activeFinalJobId?'primary busy':'primary'}`);setTextIfChanged(allBtn,hardProblems?`先に${hardProblems}項目を確認`:'提出用PDFをまとめて出力');}
  const history=$('final-history');if(history){const rows=entries.filter(entry=>entry.ready.outputPdf||entry.ready.lastBuiltAt).map(entry=>({label:`${entry.target.displayName||entry.target.targetId} PDF`,r:entry.ready}));setHtmlIfChanged(history,rows.length?`<div class="history-row header"><span>出力日時</span><span>種類</span><span>ページ数</span><span>状態</span><span>ファイル</span></div>${rows.map(({label,r})=>`<div class="history-row"><span>${escapeHtml(formatDateTime(r.lastBuiltAt))}</span><span>${escapeHtml(label)}</span><span>${Number(r.pageCount||0)}ページ</span><span>${badge(r.displayState==='built'?'最新':r.displayState==='output-missing'?'ファイルなし':'再出力必要',r.displayState==='built'?'neutral':'attention')}</span><span class="history-file">${escapeHtml(outputFileName(r.outputPdf))}</span></div>`).join('')}`:'<div class="empty-state">まだ出力していません。「提出用PDFをまとめて出力」を押すと、ここに出力日時が残ります。</div>');}
  renderPackReview();
  renderVolumeLinks();
}

function renderFinalPreflight(entries=[]) {
  const summary=$('final-preflight-summary'),list=$('final-preflight-list');if(!summary||!list)return {ready:false,problemCount:0,hardProblemCount:0,unassigned:0};
  const workbooks=workbooksForActivePreset(),requirements=activeSourceRequirements(),requiredRequirements=requirements.filter(item=>item?.required!==false);
  const requiredWorkbooks=workbooks.filter(workbook=>workbook?.required!==false);
  const missingRequired=requiredRequirements.filter(requirement=>!workbooks.some(workbook=>String(workbook?.requirementId||'')===String(requirement?.requirementId||'')&&String(workbook?.status||'')!=='missing'));
  const conversionPending=workbooks.filter(workbook=>!isLatestPdfWorkbook(workbook));
  const unassigned=unresolvedPageCount();
  const requiredTargets=entries.filter(entry=>entry.target?.required!==false),emptyRequiredTargets=requiredTargets.filter(entry=>Number(entry.ready?.pageCount||0)===0);
  const blockedTargets=entries.filter(entry=>Number(entry.ready?.pageCount||0)>0&&asArray(entry.ready?.blockers).filter(blocker=>String(blocker?.code||'')!=='no-pages').length>0);
  const settings=packForPreset()?.settings||{},pattern=String(settings.outputFileNamePattern||'{packName}_{targetName}_{yyyyMMdd}.pdf').trim();
  const checks=[];
  checks.push({label:'必須原稿',state:missingRequired.length?'danger':'ok',detail:requiredRequirements.length?(missingRequired.length?`${missingRequired.length}件が未提出です`:`必要原稿 ${requiredRequirements.length}件を提出済みです`):(requiredWorkbooks.length?`${requiredWorkbooks.length}件を必須原稿として管理しています`:'必須原稿の指定はありません'),view:'excel'});
  checks.push({label:'変換PDF',state:workbooks.length===0||conversionPending.length?'danger':'ok',detail:workbooks.length===0?'原稿がまだ登録されていません':conversionPending.length?`${conversionPending.length}件の作成・更新が必要です`:`${workbooks.length}件すべて最新です`,view:'excel'});
  checks.push({label:'ページ構成',state:unassigned?'attention':'ok',detail:unassigned?`${unassigned}ページが未振り分けです`:'未振り分けページはありません',view:'pages'});
  const targetProblems=emptyRequiredTargets.length+blockedTargets.length;
  checks.push({label:'出力先',state:targetProblems?'danger':'ok',detail:emptyRequiredTargets.length?`${emptyRequiredTargets.map(entry=>entry.target.displayName||entry.target.targetId).join('・')}にページがありません`:blockedTargets.length?`${blockedTargets.length}件の出力先に解消が必要な問題があります`:'必須の出力先を出力できます',view:emptyRequiredTargets.length?'pages':'excel'});
  checks.push({label:'出力ファイル名',state:pattern?'ok':'danger',detail:pattern?'出力名を設定済み':'出力ファイル名を設定してください',view:'final',focus:'pack-output-pattern'});
  const problems=checks.filter(check=>check.state!=='ok'),completed=checks.filter(check=>check.state==='ok');
  const problemCount=problems.length,hardProblemCount=problems.filter(check=>check.label!=='ページ構成').length,ready=problemCount===0;
  // The whole block is re-rendered by every renderAll(), so an aria-live region
  // here re-read the full preflight text on unrelated edits. Announce only when
  // the outcome actually changes.
  const announce=$('final-preflight-announce');
  if(announce){
    const key=`${ready?'ready':'attention'}:${problemCount}`;
    if(announce.dataset.key!==key){
      announce.dataset.key=key;
      announce.textContent=ready?'出力前チェック：すべての確認項目を満たしています。':`出力前チェック：${problemCount}項目の確認が必要です。`;
    }
  }
  summary.className=`final-preflight-summary ${ready?'ready':'attention'}`;
  setHtmlIfChanged(summary,`<div><strong>${ready?'出力準備が整っています':'出力前に確認が必要です'}</strong><span>${ready?'すべての確認項目を満たしています。':`${problemCount}項目を確認してください。`}</span></div>${badge(ready?'準備完了':`${problemCount}項目`,ready?'ok':'attention')}<time>確認 ${escapeHtml(finalPreflightCheckedAt?formatDateTime(finalPreflightCheckedAt):'状態取得中')}</time>`);
  const row=check=>`<div class="final-preflight-row ${escapeAttr(check.state)}"><span class="final-preflight-mark" aria-hidden="true">${check.state==='ok'?'✓':'!'}</span><div><strong>${escapeHtml(check.label)}</strong><span>${escapeHtml(check.detail)}</span></div>${check.state==='ok'?'':`<button class="btn ${check.label==='ページ構成'?'primary':'secondary'} compact" type="button" data-preflight-view="${escapeAttr(check.view)}"${check.focus?` data-preflight-focus="${escapeAttr(check.focus)}"`:''} aria-label="${escapeAttr(`${check.label}を確認する`)}">${check.label==='ページ構成'?`${unassigned}ページを確認`:`${check.label}を確認`}</button>`}</div>`;
  const listChanged=setHtmlIfChanged(list,problems.map(row).join('')+(completed.length?`<details class="preflight-complete" ${ready?'open':''}><summary>確認済み ${completed.length}項目</summary><div>${completed.map(row).join('')}</div></details>`:''));
  if(listChanged)list.querySelectorAll('[data-preflight-view]').forEach(button=>button.addEventListener('click',()=>{
    setActiveView(String(button.dataset.preflightView||'final'));
    // Without this the "出力ファイル名を確認" button only scrolled to the top of
    // the screen it was already on, leaving the field shut inside its <details>.
    const targetId=String(button.dataset.preflightFocus||'');
    if(!targetId)return;
    const field=$(targetId);
    if(!field)return;
    field.closest('details')?.setAttribute('open','');
    field.scrollIntoView({block:'center'});
    field.focus();
  }));
  return {ready,problemCount,hardProblemCount,unassigned,checks};
}

function renderPackReview(){
  const review=packReviewState,summary=$('pack-review-summary'),eventsBox=$('pack-review-events'),statusBadge=$('pack-review-status'),summaryLabel=$('pack-review-summary-label');
  const buttons={approve:$('pack-review-approve'),reopen:$('pack-review-reopen')};
  if(!summary||!eventsBox||!statusBadge)return;
  if(!review||String(review.packId||'')!==String(activePackRecord()?.packId||'')){
    if(summaryLabel)summaryLabel.textContent='状態を確認中';
    statusBadge.className='badge neutral';statusBadge.textContent='状態取得中';summary.innerHTML='<p class="caption">確認の記録を読み込んでいます。</p>';eventsBox.innerHTML='<div class="empty-state">読み込んでいます。</div>';
    Object.values(buttons).forEach(button=>{if(button)button.disabled=true;});return;
  }
  // This record lives only in the current user's local workspace and the actor is
  // always $env:USERNAME, so a submit / approve / send-back workflow could never
  // involve a second person. It is presented as a personal "checked it" note.
  const labels={draft:'未確認','in-review':'未確認',approved:'確認済み','changes-requested':'未確認',stale:'確認後に変更あり'};
  const classes={draft:'neutral','in-review':'neutral',approved:'ok','changes-requested':'neutral',stale:'attention'};
  const status=String(review.status||'draft');setClassIfChanged(statusBadge,`badge ${classes[status]||'neutral'}`);setTextIfChanged(statusBadge,labels[status]||status);
  if(summaryLabel)setTextIfChanged(summaryLabel,status==='approved'?`確認済み ${formatDateTime(review.approvedAt)}`:(status==='stale'?'確認後に変更あり':'自分用の確認メモ'));
  const targets=asArray(review.targetStates),readyCount=targets.filter(target=>target.ready).length;
  const details=[];
  if(status==='approved')details.push(`${formatDateTime(review.approvedAt)} に確認済みとして記録しました。`);
  if(status==='stale')details.push('確認したあとに原稿・ページ構成・出力が変わりました。出力し直して、もう一度確認してください。');
  if(review.note)details.push(`メモ：${review.note}`);
  const draftGuidance=review.canSubmit?'提出用PDFを開いて内容を確かめたら、「確認済みにする」で記録できます。':'必須の提出用PDFを最新にすると、確認済みとして記録できます。';
  setHtmlIfChanged(summary,`<div><strong>必須出力 ${readyCount} / ${targets.length}件が最新</strong><span>${escapeHtml(details.join(' ')||draftGuidance)}</span></div>`);
  if(buttons.approve)buttons.approve.disabled=!review.canSubmit||status==='approved';
  if(buttons.reopen)buttons.reopen.disabled=status!=='approved';
  const actionLabels={approve:'確認済みにしました',reopen:'確認を取り消しました'};
  // 'submit' is an internal step of 確認済みにする and 'request-changes' no longer
  // has a sender; neither means anything in a single-user record.
  const rows=asArray(review.events).filter(event=>['approve','reopen'].includes(String(event.action||''))).slice().reverse();
  setHtmlIfChanged(eventsBox,rows.length?rows.map(event=>`<div class="pack-review-event"><span class="pack-review-event-mark"></span><div><strong>${escapeHtml(actionLabels[String(event.action||'')]||event.action||'更新')}</strong><span>${event.note?escapeHtml(event.note):'メモなし'}</span></div><time>${escapeHtml(formatDateTimeWithSeconds(event.at))}</time></div>`).join(''):'<div class="empty-state">確認の記録はまだありません。「確認済みにする」を押すと、確認した日時がここに残ります。</div>');
}

async function performPackReviewAction(action,button){
  const pack=activePackRecord();if(!pack?.packId)return;
  const note=String($('pack-review-note')?.value||'').trim();
  await runBusy(button,async()=>{
    // The stored workflow still goes draft -> in-review -> approved. One press
    // means "I checked it", so walk both steps here rather than exposing a
    // submission step that has no recipient.
    if(action==='approve'&&String(packReviewState?.status||'draft')!=='in-review'){
      await api(`/api/v2/packs/${encodeURIComponent(pack.packId)}/review`,{method:'POST',body:{action:'submit',note:''}});
    }
    const response=await api(`/api/v2/packs/${encodeURIComponent(pack.packId)}/review`,{method:'POST',body:{action,note}});
    if(response.state)state=normalizeStatePayload(response.state);
    packReviewState=response.review||null;
    if($('pack-review-note'))$('pack-review-note').value='';
    renderAll();renderPackReview();
    const labels={approve:'確認済みとして記録しました',reopen:'確認の記録を取り消しました'};
    showMessage(action==='request-changes'?'warn':'ok',labels[action]||'レビュー状態を更新しました',pack.displayName||'資料パック');
  },false);
}

async function refreshFinalPreflight(btn) {
  await runBusy(btn,async()=>{
    if(configured())await api('/api/scan-updates',{method:'POST',body:{force:true}});
    state=normalizeStatePayload(await api('/api/state'));
    if(configured())await loadFiles(null);
    await loadFinalPanels({force:true, render:false});
    finalPreflightCheckedAt=new Date().toLocaleString('sv-SE',{timeZone:'Asia/Tokyo'});
    renderAll();
    showMessage('ok','出力前チェックを更新しました','原稿・変換PDF・ページ構成・出力設定を最新状態で確認しました。');
  },false);
}


function updateBulkSelectionLabel() {
  const count = selectedPages.size;
  const el = $('bulk-selected-count');
  if (el) el.textContent = count ? `${count}ページ選択中。移動先を押してください` : 'ページをクリックして選ぶと移動できます';
  const bar = $('page-command-bar');
  if (bar) bar.classList.toggle('has-selection', count > 0);
  document.querySelectorAll('[data-bulk-target],#bulk-none-btn,#clear-selected-pages-btn').forEach(button=>{
    const destination=button.matches('#bulk-none-btn')?'none':String(button.dataset.bulkTarget||'');
    // pageMutationBusy: a background re-render used to re-enable these mid-move,
    // letting a second destination be pressed with the same page ids in flight.
    button.disabled=pageMutationBusy||count===0||(destination&&destination===activePageVolume);
  });
  const undoEntry=pageLayoutUndoStack[pageLayoutUndoStack.length-1],redoEntry=pageLayoutRedoStack[pageLayoutRedoStack.length-1],undo=$('page-layout-undo-btn'),redo=$('page-layout-redo-btn');
  if(undo){undo.disabled=pageLayoutHistoryBusy||!undoEntry;undo.title=undoEntry?`${undoEntry.label}を元に戻す（Ctrl+Z）`:'元に戻せるページ構成変更はありません';undo.setAttribute('aria-label',undo.title);}
  if(redo){redo.disabled=pageLayoutHistoryBusy||!redoEntry;redo.title=redoEntry?`${redoEntry.label}をやり直す（Ctrl+Shift+Z）`:'やり直せるページ構成変更はありません';redo.setAttribute('aria-label',redo.title);}
}

async function savePackSettings(btn){
  const pack=packForPreset();if(!pack?.packId){showMessage('danger','資料パック設定を保存できません','対象の資料パックが見つかりません。');return;}
  const body={outputFileNamePattern:String($('pack-output-pattern')?.value||'').trim()};
  if(!body.outputFileNamePattern){showMessage('warn','出力ファイル名を入力してください','使用できる変数は {packName}・{targetName}・{yyyyMMdd} です。');$('pack-output-pattern')?.focus();return;}
  await runBusy(btn,async()=>{try{const response=await api(`/api/v2/packs/${encodeURIComponent(pack.packId)}`,{method:'PATCH',body});state=normalizeStatePayload(response.state||await api('/api/state'));renderAll();showMessage('ok','資料の仕上げ設定を保存しました','次回の提出用PDF出力から反映されます。');}catch(error){showMessage('danger','資料の仕上げ設定を保存できません',userFriendlyError(error.message),error.detail||error.stack||error.message);}});
}
function requiredRenderProfileVersion(source) {
  if (sourceTypeValue(source) === 'pdf') {
    const v = Number(state?.pdfImportProfileVersion || 0);
    return Number.isFinite(v) && v > 0 ? v : CURRENT_PDF_IMPORT_PROFILE_VERSION;
  }
  if (sourceTypeValue(source) === 'word') {
    const v = Number(state?.wordRenderProfileVersion || 0);
    return Number.isFinite(v) && v > 0 ? v : CURRENT_WORD_RENDER_PROFILE_VERSION;
  }
  if (sourceTypeValue(source) === 'powerpoint') {
    const v = Number(state?.powerPointRenderProfileVersion || 0);
    return Number.isFinite(v) && v > 0 ? v : CURRENT_POWERPOINT_RENDER_PROFILE_VERSION;
  }
  return currentRenderProfileVersion();
}

function closeConfirmModal(accepted=false) {
  const modal=$('confirm-modal');
  if(!modal||modal.classList.contains('hidden'))return;
  modal.classList.add('hidden');
  const resolve=confirmResolver,focusTarget=confirmReturnFocus;
  confirmResolver=null;confirmReturnFocus=null;
  if(focusTarget?.isConnected&&typeof focusTarget.focus==='function')focusTarget.focus();
  if(resolve)resolve(accepted);
}
// ファイル名を変えた・移動した原稿を、ページの並び順や出力先の振り分けを保ったまま
// 別のファイルへ結び直す。登録解除→再登録はページを丸ごと消してしまうため。
let relinkSourceId='';
async function openRelinkDialog(sourceId, trigger=null) {
  if(!sourceId)return;
  await runBusy(trigger,async()=>{
    const res=await api(`/api/v2/sources/relink-candidates?sourceId=${encodeURIComponent(sourceId)}`);
    const info=res?.result||{};
    const candidates=asArray(info.candidates);
    const select=$('relink-candidate');
    if(!candidates.length){
      showMessage('warn','付け替えられるファイルがありません',`原稿フォルダーの直下に、まだ登録されていない同じ種類のファイルが見つかりません。ファイルを別の場所へ移した場合は、元のフォルダーに戻してからもう一度お試しください。`);
      return;
    }
    relinkSourceId=sourceId;
    setTextIfChanged($('relink-message'),`「${info.fileName||sourceId}」が見つかりません。どのファイルに付け替えますか。`);
    select.innerHTML=candidates.map(c=>`<option value="${escapeAttr(c.relativePath)}">${escapeHtml(c.fileName||c.relativePath)}${c.sameContent?'（中身が同じ）':''}</option>`).join('');
    const same=candidates.find(c=>c.sameContent);
    if(same)select.value=String(same.relativePath);
    updateRelinkNote(candidates);
    select.onchange=()=>updateRelinkNote(candidates);
    $('relink-modal').classList.remove('hidden');
    select.focus();
  });
}
function updateRelinkNote(candidates) {
  const select=$('relink-candidate');
  const chosen=asArray(candidates).find(c=>String(c.relativePath)===String(select?.value||''));
  setTextIfChanged($('relink-note'),chosen?.sameContent
    ? '中身が同じファイルです。名前が変わっただけなので、変換PDFを作り直す必要はありません。'
    : '中身が異なるファイルです。付け替えたあと、変換PDFの作成と提出用PDFの再出力が必要になります。ページの並び順と出力先はそのまま残ります。');
}
function closeRelinkDialog() { $('relink-modal')?.classList.add('hidden'); relinkSourceId=''; }
async function submitRelink(btn) {
  const relativePath=String($('relink-candidate')?.value||'');
  if(!relinkSourceId||!relativePath)return;
  const sourceId=relinkSourceId;
  await runBusy(btn,async()=>{
    const res=await api('/api/v2/sources/relink',{method:'POST',body:{sourceId,relativePath}});
    closeRelinkDialog();
    await refresh();
    const r=res?.result||{};
    showMessage('ok','原稿を付け替えました',r.sameContent
      ? `${r.fileName} に付け替えました。中身は同じなので、変換PDFはそのまま使えます。`
      : `${r.fileName} に付け替えました。ページの並び順と出力先はそのままです。変換PDFを作成し直してください。`);
  });
}
function confirmAction({title='操作を実行しますか？',message='',detail='',confirmLabel='実行する',danger=false}={}) {
  if(isConfirmModalOpen())closeConfirmModal(false);
  confirmReturnFocus=document.activeElement;
  $('confirm-title').textContent=title;
  $('confirm-message').textContent=message;
  const detailBox=$('confirm-detail');
  detailBox.textContent=detail||'';
  detailBox.classList.toggle('hidden',!detail);
  const accept=$('confirm-accept');
  accept.textContent=confirmLabel;
  accept.className=`btn ${danger?'danger':'primary'}`;
  $('confirm-modal').classList.remove('hidden');
  requestAnimationFrame(()=>$('confirm-cancel')?.focus());
  return new Promise(resolve=>{confirmResolver=resolve;});
}

function closeAppMenu(focusButton=false){
  const popover=$('app-menu-popover'),button=$('app-menu-button');
  if(!popover||popover.classList.contains('hidden'))return;
  popover.classList.add('hidden');button?.setAttribute('aria-expanded','false');
  if(focusButton)button?.focus();
}
function appHasActiveWork(){
  return renderJobActive||!!activeRenderJobId||!!document.querySelector('.btn.busy');
}
async function requestAppShutdown(button){
  closeAppMenu();
  if(appHasActiveWork()){
    showMessage('warn','処理中は終了できません','実行中の処理が完了するか、PDF作成を中止してから終了してください。');
    return;
  }
  const accepted=await confirmAction({title:'ReportBinderを終了しますか？',message:'アプリのバックグラウンド処理を終了します。',detail:'編集中の入力内容を確認してから終了してください。作成済みのPDFや保存済みのページ構成は保持されます。',confirmLabel:'終了する',danger:true});
  if(!accepted)return;
  if(appHasActiveWork()){
    showMessage('warn','処理が開始されたため終了できません','実行中の処理が完了してから、もう一度終了してください。');
    return;
  }
  if(button){button.disabled=true;button.classList.add('busy');}
  try{
    await api('/api/shutdown',{method:'POST',body:{reason:'user-menu'},keepalive:true});
    document.body.classList.add('app-ended');
    const screen=$('shutdown-screen');screen?.classList.remove('hidden');screen?.focus();
  }catch(error){
    showMessage('danger','ReportBinderを終了できませんでした',userFriendlyError(error.message),error.detail||error.message);
    if(button){button.disabled=false;button.classList.remove('busy');}
  }
}

function findPackForManagement(packId) {
  const id=String(packId||'');
  return packAdminPacks.find(pack=>String(pack?.packId||'')===id)
    || asArray(state?.packs).find(pack=>String(pack?.packId||'')===id)
    || null;
}
function updatePackTemplateDescription() {
  const templateId=String($('pack-editor-template')?.value||'');
  const template=asArray(state?.packTemplates).find(item=>String(item?.templateId||'')===templateId);
  const description=$('pack-editor-template-description');if(description)description.textContent=String(template?.description||'Excel・Word・PowerPoint・PDFをまとめる資料パックです。');
}
function openPackEditor(mode='create',packId='') {
  const modal=$('pack-editor-modal');if(!modal)return;
  if(mode!=='rename'&&!configured()){
    setActiveView('excel');
    showMessage('warn','先に原稿フォルダーを選んでください','フォルダーを選ぶと、続けて今回まとめる一式に名前を付けられます。');
    setTimeout(()=>$('change-source-folder-btn')?.focus(),0);
    return;
  }
  const pack=findPackForManagement(packId);
  if(mode==='rename'&&!pack){showMessage('warn','資料パックが見つかりません','一覧を読み込み直してください。');return;}
  packEditorMode=mode==='rename'?'rename':'create';packEditorPackId=packEditorMode==='rename'?String(pack.packId||''):'';packEditorReturnFocus=$('pack-menu-button')||document.activeElement;
  $('pack-editor-title').textContent=packEditorMode==='rename'?'一式の名前を変更':'今回まとめる一式に名前を付ける';
  $('pack-editor-description').textContent=packEditorMode==='rename'?'登録した原稿やページの順番はそのまま、表示する名前だけを変更します。':'会議名や提出先など、あとで見つけやすい名前を付けます。';
  const name=$('pack-editor-name');if(name)name.value=packEditorMode==='rename'?String(pack.displayName||''):'';
  const select=$('pack-editor-template');
  const templates=asArray(state?.packTemplates);
  const field=$('pack-editor-template-field');field?.classList.toggle('hidden',packEditorMode==='rename'||templates.length<=1);
  if(select){
    select.innerHTML=templates.map(template=>`<option value="${escapeAttr(template.templateId||'')}">${escapeHtml(template.displayName||template.templateId||'ひな形')}</option>`).join('');
    select.value=packEditorMode==='rename'?String(pack?.templateId||''):String(templates.find(template=>String(template.templateId)==='builtin-generic-department-pack')?.templateId||templates[0]?.templateId||'');
  }
  $('pack-editor-save').textContent=packEditorMode==='rename'?'名前を変更':'作成';updatePackTemplateDescription();
  closePackMenu();modal.classList.remove('hidden');requestAnimationFrame(()=>name?.focus());
}
function closePackEditor(returnFocus=true) {
  const modal=$('pack-editor-modal');if(!modal||modal.classList.contains('hidden'))return;
  modal.classList.add('hidden');const focus=packEditorReturnFocus;packEditorReturnFocus=null;packEditorPackId='';
  if(returnFocus&&focus?.isConnected&&typeof focus.focus==='function')focus.focus();
}
async function refreshPacksAfterMutation() {
  packAdminPacks=[];await refresh();await loadPackAdminPacks();renderPackSwitcher();
}
async function submitPackEditor(btn) {
  const name=String($('pack-editor-name')?.value||'').trim();
  if(!name){showMessage('warn','資料パック名を入力してください','用途が分かる名前を入力してください。');$('pack-editor-name')?.focus();return;}
  await runBusy(btn,async()=>{
    let response;
    if(packEditorMode==='rename')response=await api(`/api/v2/packs/${encodeURIComponent(packEditorPackId)}`,{method:'PATCH',body:{displayName:name}});
    else response=await api('/api/v2/packs',{method:'POST',body:{displayName:name,templateId:String($('pack-editor-template')?.value||'builtin-generic-department-pack')}});
    const createdName=String(response?.result?.displayName||name),wasRename=packEditorMode==='rename';
    closePackEditor(false);await refreshPacksAfterMutation();
    if(!wasRename&&response?.result?.packId)await applyPresetSelection(String(response.result.packId));
    if(!wasRename){setActiveView('excel');if(!availableFiles.length)await loadFilesSilently();}
    showMessage('ok',wasRename?'名前を変更しました':'今回まとめる一式を作成しました',wasRename?`${createdName}として表示します。`:`「${createdName}」にまとめる原稿を選んでください。`);
  },false);
}
const templateManagerFieldIds=['template-name','template-description','template-source-excel','template-source-word','template-source-powerpoint','template-source-pdf','template-new-destination','template-output-pattern','template-block-stale','template-block-failed'];
function newTemplateTarget(){return {targetId:`output_${Date.now().toString(16).slice(-8)}`,displayName:'',required:false};}
function collectTemplateTargets(){return [...document.querySelectorAll('[data-template-target]')].map(row=>({targetId:String(row.querySelector('[data-target-id]')?.value||'').trim().toLowerCase(),displayName:String(row.querySelector('[data-target-name]')?.value||'').trim(),required:!!row.querySelector('[data-target-required]')?.checked}));}
function templateTargetOptions(selected='unassigned'){return `<option value="unassigned" ${selected==='unassigned'?'selected':''}>未振り分け</option>${collectTemplateTargets().filter(target=>target.targetId).map(target=>`<option value="${escapeAttr(target.targetId)}" ${selected===target.targetId?'selected':''}>${escapeHtml(target.displayName||target.targetId)}</option>`).join('')}`;}
function renderTemplateTargets(targets=[],readOnly=false){
  const box=$('template-targets-list');if(!box)return;const items=asArray(targets).length?asArray(targets):[{targetId:'main',displayName:'本体',required:true},{targetId:'appendix',displayName:'補足',required:false}];box.innerHTML=items.map((target,index)=>`<div class="template-target-row" data-template-target><label><span>出力先ID</span><input data-target-id type="text" maxlength="48" pattern="[a-z0-9][a-z0-9_\\-]*" value="${escapeAttr(target.targetId||'')}" ${readOnly?'disabled':''}></label><label><span>表示名</span><input data-target-name type="text" maxlength="60" value="${escapeAttr(target.displayName||'')}" ${readOnly?'disabled':''}></label><label class="template-check"><input data-target-required type="checkbox" ${target.required?'checked':''} ${readOnly?'disabled':''}> 必須</label><button class="btn ghost compact" type="button" data-remove-target ${readOnly||items.length<=1?'disabled':''}>削除</button></div>`).join('');
  box.querySelectorAll('[data-remove-target]').forEach((button,index)=>button.addEventListener('click',()=>{const requirements=collectTemplateRequirements(),destination=String($('template-new-destination')?.value||'unassigned'),next=collectTemplateTargets().filter((_,itemIndex)=>itemIndex!==index);renderTemplateTargets(next,false);renderTemplateDestinationOptions(destination);renderTemplateRequirements(requirements,false);}));
  box.querySelectorAll('[data-target-id],[data-target-name]').forEach(input=>input.addEventListener('change',()=>{const requirements=collectTemplateRequirements(),destination=String($('template-new-destination')?.value||'unassigned');renderTemplateDestinationOptions(destination);renderTemplateRequirements(requirements,readOnly);}));
}
function renderTemplateDestinationOptions(selected='unassigned'){const select=$('template-new-destination');if(select)select.innerHTML=templateTargetOptions(selected);}
function newTemplateRequirement(){return {requirementId:`requirement_${Date.now().toString(16)}${Math.random().toString(16).slice(2,8)}`,displayName:'',ownerDepartment:'',required:true,acceptedSourceTypes:['excel','word','powerpoint','pdf'],defaultTargetId:'unassigned',dueDate:''};}
function collectTemplateRequirements(){return [...document.querySelectorAll('[data-template-requirement]')].map(row=>({requirementId:String(row.dataset.templateRequirement||''),displayName:String(row.querySelector('[data-requirement-name]')?.value||'').trim(),ownerDepartment:String(row.querySelector('[data-requirement-owner]')?.value||'').trim(),required:!!row.querySelector('[data-requirement-required]')?.checked,acceptedSourceTypes:[...row.querySelectorAll('[data-requirement-type]:checked')].map(input=>String(input.value||'')),defaultTargetId:String(row.querySelector('[data-requirement-target]')?.value||'unassigned'),dueDate:String(row.querySelector('[data-requirement-due]')?.value||'')}));}
function renderTemplateRequirements(requirements=[],readOnly=false){
  const box=$('template-requirements-list');if(!box)return;const items=asArray(requirements);box.innerHTML=items.length?items.map((r,index)=>{const types=asArray(r.acceptedSourceTypes).map(String),id=String(r.requirementId||newTemplateRequirement().requirementId);return `<div class="template-requirement-row" data-template-requirement="${escapeAttr(id)}"><div class="template-requirement-head"><strong>必要原稿 ${index+1}</strong><button class="btn ghost compact" type="button" data-remove-requirement ${readOnly?'disabled':''}>削除</button></div><div class="template-manager-grid"><label><span>原稿名</span><input type="text" maxlength="120" value="${escapeAttr(r.displayName||'')}" data-requirement-name ${readOnly?'disabled':''}></label><label><span>担当部署</span><input type="text" maxlength="120" value="${escapeAttr(r.ownerDepartment||'')}" data-requirement-owner ${readOnly?'disabled':''}></label><label><span>期限</span><input type="date" value="${escapeAttr(r.dueDate||'')}" data-requirement-due ${readOnly?'disabled':''}></label><label><span>新規ページの配置</span><select data-requirement-target ${readOnly?'disabled':''}>${templateTargetOptions(String(r.defaultTargetId||'unassigned'))}</select></label></div><div class="template-inline-options"><label><input type="checkbox" data-requirement-required ${r.required!==false?'checked':''} ${readOnly?'disabled':''}> 必須</label>${['excel','word','powerpoint','pdf'].map(type=>`<label><input type="checkbox" value="${type}" data-requirement-type ${types.includes(type)?'checked':''} ${readOnly?'disabled':''}> ${{excel:'Excel',word:'Word',powerpoint:'PowerPoint',pdf:'PDF'}[type]}</label>`).join('')}</div></div>`;}).join(''):'<div class="template-requirements-empty">必要原稿は未設定です。必要に応じて追加してください。</div>';
  box.querySelectorAll('[data-remove-requirement]').forEach(button=>button.addEventListener('click',()=>{const current=collectTemplateRequirements();const id=String(button.closest('[data-template-requirement]')?.dataset.templateRequirement||'');renderTemplateRequirements(current.filter(item=>item.requirementId!==id),false);}));
}
function templateManagerSelected(){const id=String($('template-manager-select')?.value||'');return asArray(state?.packTemplates).find(item=>String(item?.templateId||'')===id)||null;}
function renderTemplateManagerSelect(selectedId=''){
  const select=$('template-manager-select');if(!select)return;
  const templates=asArray(state?.packTemplates);select.innerHTML=`<option value="">＋ 新しいひな形</option>${templates.map(t=>`<option value="${escapeAttr(t.templateId||'')}">${escapeHtml(t.displayName||t.templateId||'ひな形')}${t.builtIn?'（組み込み）':''}</option>`).join('')}`;
  select.value=templates.some(t=>String(t.templateId||'')===String(selectedId||''))?String(selectedId):'';populateTemplateManager();
}
function populateTemplateManager(){
  const t=templateManagerSelected(),builtIn=!!t?.builtIn,targets=asArray(t?.targets),rules=t?.rules||{},output=t?.output||{},sourceTypes=asArray(t?.acceptedSourceTypes).map(String);
  const setValue=(id,value)=>{const el=$(id);if(el)el.value=String(value??'');},setChecked=(id,value)=>{const el=$(id);if(el)el.checked=!!value;};
  setValue('template-name',t?.displayName||'');setValue('template-description',t?.description||'');setChecked('template-source-excel',t?sourceTypes.includes('excel'):true);setChecked('template-source-word',t?sourceTypes.includes('word'):true);setChecked('template-source-powerpoint',t?sourceTypes.includes('powerpoint'):true);setChecked('template-source-pdf',t?sourceTypes.includes('pdf'):true);
  renderTemplateTargets(targets,builtIn);renderTemplateDestinationOptions(rules.newItemDestination||'unassigned');setValue('template-output-pattern',output.fileNamePattern||'{packName}_{targetName}_{yyyyMMdd}.pdf');setChecked('template-block-stale',t?rules.blockBuildWhenRequiredSourceIsStale:true);setChecked('template-block-failed',t?rules.blockBuildWhenRequiredSourceFailed:true);
  renderTemplateRequirements(t?.sourceRequirements||[],builtIn);
  templateManagerFieldIds.forEach(id=>{const el=$(id);if(el)el.disabled=builtIn;});const save=$('template-manager-save'),remove=$('template-manager-delete'),duplicate=$('template-manager-duplicate'),exportButton=$('template-manager-export'),stateText=$('template-manager-state'),usageCount=t?asArray(packAdminPacks.length?packAdminPacks:state?.packs).filter(pack=>String(pack?.templateId||'')===String(t.templateId||'')).length:0;if(save){save.disabled=builtIn;save.textContent=t?'変更を保存':'作成';}if(remove)remove.classList.toggle('hidden',!t||builtIn);if(duplicate)duplicate.classList.toggle('hidden',!t);if(exportButton)exportButton.classList.toggle('hidden',!t);if(stateText)stateText.textContent=builtIn?`組み込みひな形・使用中 ${usageCount}件。内容を基にする場合は「複製」を選んでください。`:(t?`バージョン ${Number(t.templateVersion||1)}・使用中 ${usageCount}件・変更は新しく作る資料パックから反映されます。`:'新しいひな形を作成します。');
  const addRequirement=$('template-add-requirement'),addTarget=$('template-add-target');if(addRequirement)addRequirement.disabled=builtIn;if(addTarget)addTarget.disabled=builtIn;
}
function openTemplateManager(){const modal=$('template-manager-modal');if(!modal)return;templateManagerReturnFocus=$('manage-pack-templates-btn')||document.activeElement;closePackMenu();modal.classList.remove('hidden');renderTemplateManagerSelect('');requestAnimationFrame(()=>$('template-name')?.focus());}
function closeTemplateManager(returnFocus=true){const modal=$('template-manager-modal');if(!modal||modal.classList.contains('hidden'))return;modal.classList.add('hidden');const focus=templateManagerReturnFocus;templateManagerReturnFocus=null;if(returnFocus&&focus?.isConnected&&typeof focus.focus==='function')focus.focus();}
function templateManagerRequest(){
  const acceptedSourceTypes=[['template-source-excel','excel'],['template-source-word','word'],['template-source-powerpoint','powerpoint'],['template-source-pdf','pdf']].filter(([id])=>!!$(id)?.checked).map(([,type])=>type);
  return {displayName:String($('template-name')?.value||'').trim(),description:String($('template-description')?.value||'').trim(),acceptedSourceTypes,sourceRequirements:collectTemplateRequirements(),targets:collectTemplateTargets(),rules:{newItemDestination:String($('template-new-destination')?.value||'unassigned'),retainManualOrder:true,blockBuildWhenRequiredSourceIsStale:!!$('template-block-stale')?.checked,blockBuildWhenRequiredSourceFailed:!!$('template-block-failed')?.checked},output:{fileNamePattern:String($('template-output-pattern')?.value||'').trim(),pageNumbering:'continuous',bookmarks:'from-items'}};
}
async function submitTemplateManager(btn){
  const body=templateManagerRequest(),selected=templateManagerSelected();if(!body.displayName){showMessage('warn','ひな形名を入力してください','用途が分かる名前を入力してください。');$('template-name')?.focus();return;}if(!body.acceptedSourceTypes.length){showMessage('warn','原稿形式を選んでください','Excel・Word・PowerPoint・PDFから1つ以上選択してください。');return;}
  const invalidTarget=body.targets.find(target=>!target.targetId||!target.displayName||!/^[a-z0-9][a-z0-9_-]{0,47}$/.test(target.targetId));if(!body.targets.length||invalidTarget||new Set(body.targets.map(target=>target.targetId)).size!==body.targets.length){showMessage('warn','出力先の設定を確認してください','出力先IDは英小文字・数字・ハイフン・アンダースコアで重複なく入力し、表示名も設定してください。');return;}const invalidRequirement=body.sourceRequirements.find(item=>!item.displayName||!item.acceptedSourceTypes.length);if(invalidRequirement){showMessage('warn','必要原稿の設定を確認してください','原稿名を入力し、扱う形式を1つ以上選択してください。');return;}
  await runBusy(btn,async()=>{const response=await api(selected?`/api/v2/pack-templates/${encodeURIComponent(selected.templateId)}`:'/api/v2/pack-templates',{method:selected?'PATCH':'POST',body});state.packTemplates=asArray(response.packTemplates);renderTemplateManagerSelect(String(response?.result?.templateId||''));showMessage('ok',selected?'ひな形を更新しました':'ひな形を作成しました','新しく作る資料パックで選択できます。');},false);
}
async function deleteManagedTemplate(btn){const selected=templateManagerSelected();if(!selected||selected.builtIn)return;const accepted=await confirmAction({title:'このひな形を削除しますか？',message:`「${selected.displayName||'ひな形'}」を削除します。`,detail:'このひな形からすでに作った資料パックは、そのまま使えます（原稿・ページ構成・出力先の設定は残ります）。削除すると、今後の資料パック作成でこのひな形を選べなくなり、必要原稿リストと出力条件の設定は元に戻せません。',confirmLabel:'削除',danger:true});if(!accepted)return;await runBusy(btn,async()=>{const response=await api(`/api/v2/pack-templates/${encodeURIComponent(selected.templateId)}`,{method:'DELETE'});state.packTemplates=asArray(response.packTemplates);renderTemplateManagerSelect('');showMessage('ok','ひな形を削除しました','資料パック作成時の一覧から削除しました。');},false);}
async function duplicateManagedTemplate(btn){const selected=templateManagerSelected();if(!selected)return;const body=templateManagerRequest();body.displayName=`${body.displayName||selected.displayName||'ひな形'}（コピー）`;await runBusy(btn,async()=>{const response=await api('/api/v2/pack-templates',{method:'POST',body});state.packTemplates=asArray(response.packTemplates);renderTemplateManagerSelect(String(response?.result?.templateId||''));showMessage('ok','ひな形を複製しました','用途に合わせて名前や必要原稿を編集できます。');},false);}
function exportManagedTemplate(){const selected=templateManagerSelected();if(!selected)return;const data=templateManagerRequest();data.templateSchemaVersion=1;const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=`ReportBinder-template-${String(selected.displayName||'template').replace(/[\\/:*?"<>|]/g,'_')}.json`;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);showMessage('ok','ひな形ひな形をファイルに書き出しました','別のPCや利用者へ渡して読み込めます。');}
async function importManagedTemplateFile(file,btn){if(!file)return;try{const body=JSON.parse(await file.text());delete body.templateId;delete body.templateVersion;const response=await api('/api/v2/pack-templates',{method:'POST',body});state.packTemplates=asArray(response.packTemplates);renderTemplateManagerSelect(String(response?.result?.templateId||''));showMessage('ok','ひな形ひな形をファイルから読み込みました','内容を確認してから資料パックを作成してください。');}catch(error){showMessage('danger','ひな形ひな形を読み込めません',userFriendlyError(error.message),error.detail||error.message);}finally{if(btn)btn.value='';}}
async function duplicatePack(packId,btn) {
  const pack=findPackForManagement(packId);if(!pack)return;
  closePackMenu();
  await runBusy(btn,async()=>{const response=await api(`/api/v2/packs/${encodeURIComponent(packId)}/duplicate`,{method:'POST',body:{}});await refreshPacksAfterMutation();showMessage('ok','資料パックを複製しました',`${response?.result?.displayName||pack.displayName}を新しい資料パックとして作成しました。`);},false);
}
async function archivePack(packId) {
  const pack=findPackForManagement(packId);if(!pack)return;
  closePackMenu();
  const accepted=await confirmAction({title:'資料パックをアーカイブしますか？',message:`「${pack.displayName||'資料パック'}」を通常の一覧から隠します。`,detail:'登録情報は削除されず、アーカイブ済み一覧から復元できます。',confirmLabel:'アーカイブ',danger:true});
  if(!accepted)return;
  await runBusy(null,async()=>{await api(`/api/v2/packs/${encodeURIComponent(packId)}/archive`,{method:'POST',body:{}});await refreshPacksAfterMutation();showMessage('ok','資料パックをアーカイブしました','必要になったら資料パック一覧から復元できます。');},false);
}
async function restorePack(packId,btn) {
  const pack=findPackForManagement(packId);if(!pack)return;
  closePackMenu();
  await runBusy(btn,async()=>{await api(`/api/v2/packs/${encodeURIComponent(packId)}/restore`,{method:'POST',body:{}});await refreshPacksAfterMutation();showMessage('ok','資料パックを復元しました',`${pack.displayName||'資料パック'}を通常の一覧へ戻しました。`);},false);
}
async function upgradePackTemplate(packId,btn){
  closePackMenu();
  try{
    const response=await api(`/api/v2/packs/${encodeURIComponent(packId)}/template-upgrade`),preview=response.preview||{},changes=asArray(preview.changes);
    if(!preview.updateAvailable){showMessage('ok','ひな形は最新です','この資料パックに適用できる更新はありません。');return;}
    const accepted=await confirmAction({title:'新しいひな形を適用しますか？',message:`${preview.packName||'資料パック'}を v${Number(preview.fromVersion||1)} から v${Number(preview.toVersion||1)} へ更新します。`,detail:[...changes.map(item=>`・${item}`),'','既存原稿とページ構成は保持され、提出用PDFは再出力が必要になります。'].join('\n'),confirmLabel:'ひな形を更新'});
    if(!accepted)return;
    await runBusy(btn,async()=>{await api(`/api/v2/packs/${encodeURIComponent(packId)}/template-upgrade`,{method:'POST',body:{}});await refreshPacksAfterMutation();showMessage('ok','ひな形を更新しました','必要原稿リストと出力条件を更新しました。既存の原稿・ページ構成は保持しています。');},false);
  }catch(error){showMessage('danger','ひな形を更新できません',userFriendlyError(error.message),error.detail||error.message);}
}
function handlePackMenuAction(event) {
  const select=event.target.closest('[data-pack-select]');if(select){void applyPresetSelection(select.dataset.packSelect);closePackMenu();return;}
  const unavailable=event.target.closest('[data-pack-unavailable]');if(unavailable){showMessage('warn','この資料パックは利用できません','アーカイブ済みの場合は復元してから選択してください。');return;}
  const action=event.target.closest('[data-pack-action]');if(!action)return;
  const packId=String(action.dataset.packId||'');
  if(action.dataset.packAction==='rename')openPackEditor('rename',packId);
  else if(action.dataset.packAction==='duplicate')void duplicatePack(packId,action);
  else if(action.dataset.packAction==='archive')void archivePack(packId);
  else if(action.dataset.packAction==='restore')void restorePack(packId,action);
  else if(action.dataset.packAction==='template-upgrade')void upgradePackTemplate(packId,action);
}

function syncPageBoardViewControls() {
  const thumbnail = $('page-view-thumbnail-btn');
  const detail = $('page-view-detail-btn');
  if (thumbnail) { thumbnail.classList.toggle('active', pageBoardView === 'thumbnail'); thumbnail.setAttribute('aria-pressed', String(pageBoardView === 'thumbnail')); }
  if (detail) { detail.classList.toggle('active', pageBoardView === 'detail'); detail.setAttribute('aria-pressed', String(pageBoardView === 'detail')); }
  const sizeControl = $('page-thumbnail-size-control');
  if (sizeControl) sizeControl.classList.toggle('hidden', pageBoardView !== 'thumbnail');
  applyPageThumbnailSize();
}

function applyPageThumbnailSize(value=pageThumbnailSize) {
  pageThumbnailSize = Math.min(260,Math.max(150,Number(value)||190));
  document.documentElement.style.setProperty('--page-thumbnail-min',`${pageThumbnailSize}px`);
  const input=$('page-thumbnail-size'),label=$('page-thumbnail-size-label');
  if(input&&Number(input.value)!==pageThumbnailSize)input.value=String(pageThumbnailSize);
  if(label)label.textContent=pageThumbnailSize<=165?'小':pageThumbnailSize>=225?'大':'標準';
}

function setPageBoardView(view) {
  const next = view === 'detail' ? 'detail' : 'thumbnail';
  if (pageBoardView === next) return;
  pageBoardView = next;
  try { sessionStorage.setItem('ReportBinderPageBoardView', pageBoardView); } catch {}
  lastPageBoardRenderSignature = '';
  syncPageBoardViewControls();
  renderPages();
}

function clonePageVolumes(volumes){return Object.fromEntries(Object.entries(volumes||{}).map(([key,ids])=>[key,[...asArray(ids).map(String)]]));}
function pageVolumeSnapshotsEqual(a,b){const left=clonePageVolumes(a),right=clonePageVolumes(b),keys=[...new Set([...Object.keys(left),...Object.keys(right)])].sort();return keys.every(key=>JSON.stringify(left[key]||[])===JSON.stringify(right[key]||[]));}
function trimPageLayoutHistory(stack){if(stack.length>PAGE_LAYOUT_HISTORY_LIMIT)stack.splice(0,stack.length-PAGE_LAYOUT_HISTORY_LIMIT);}
function clearPageLayoutHistory(){pageLayoutUndoStack=[];pageLayoutRedoStack=[];pageLayoutHistoryBusy=false;pendingPageLayoutUndo=null;pendingBoardSaveVolumes=null;if(boardSaveTimer){clearTimeout(boardSaveTimer);boardSaveTimer=null;}updateBulkSelectionLabel();}
function rememberPageLayoutUndo(label,volumes=null,afterVolumes=null){
  const before=clonePageVolumes(volumes||collectBoardVolumes()),after=clonePageVolumes(afterVolumes||collectBoardVolumes());if(!Object.values(before).some(ids=>ids.length)||pageVolumeSnapshotsEqual(before,after))return false;pageLayoutUndoStack.push({label:String(label||'ページ構成の変更'),before,after});trimPageLayoutHistory(pageLayoutUndoStack);pageLayoutRedoStack=[];updateBulkSelectionLabel();return true;
}
async function applyPageLayoutHistory(direction,btn=null){
  if(pageLayoutHistoryBusy)return;if(pendingPageLayoutUndo){if(boardSaveTimer){clearTimeout(boardSaveTimer);boardSaveTimer=null;}const saved=await saveBoardOrder();if(!saved)return;}const savedBeforeHistory=await boardSavePromise;if(!savedBeforeHistory)return;
  const redo=direction==='redo',source=redo?pageLayoutRedoStack:pageLayoutUndoStack,target=redo?pageLayoutUndoStack:pageLayoutRedoStack,entry=source[source.length-1];if(!entry)return;pageLayoutHistoryBusy=true;updateBulkSelectionLabel();const button=btn||$(redo?'page-layout-redo-btn':'page-layout-undo-btn');if(button){button.disabled=true;button.classList.add('busy');}
  try{const response=await api('/api/v2/items/reorder',{method:'POST',body:pageApiBody({volumes:redo?entry.after:entry.before})});source.pop();target.push(entry);trimPageLayoutHistory(target);applyPageMutationResult(response);selectedPages.clear();lastPageRangeAnchor='';syncPageSelectionUi();showMessage('ok',redo?'ページ構成をやり直しました':'ページ構成を元に戻しました',entry.label,null,[{label:redo?'元に戻す':'やり直す',handler:()=>redo?undoLastPageLayout():redoLastPageLayout()}],8000);}
  catch(error){showMessage('danger',redo?'ページ構成をやり直せません':'ページ構成を元に戻せません',userFriendlyError(error.message),error.detail||error.stack||error.message);}
  finally{pageLayoutHistoryBusy=false;if(button)button.classList.remove('busy');updateBulkSelectionLabel();}
}
async function undoLastPageLayout(btn=null){return applyPageLayoutHistory('undo',btn);}
async function redoLastPageLayout(btn=null){return applyPageLayoutHistory('redo',btn);}

function renderVolumeLinks() {
  document.querySelectorAll('[data-final-volume]').forEach(card=>{
    const volume=String(card.dataset.finalVolume||''),a=card.querySelector('[data-final-open]'),publish=card.querySelector('[data-final-publish]'),r=volumeReadiness(volume);
    const available=!!(r.outputPdf && r.outputPdfExists);
    const publishable=available&&String(r.displayState||'')==='built';
    // Three equal-weight buttons side by side gave no clue what to do next.
    // Once the PDF exists, checking it is the next step, so it becomes the
    // primary action and moves to the front; re-outputting drops to tertiary.
    const fresh=available&&String(r.displayState||'')==='built';
    if(a){
      if(available){a.href='#';a.dataset.volume=volume;a.dataset.category=activePreset;setTextIfChanged(a,fresh?'出力したPDFを開く':'前回出力を開く');setClassIfChanged(a,fresh?'btn primary':'btn secondary');a.classList.remove('hidden');}
      else{a.removeAttribute('href');a.classList.add('hidden');}
    }
    const buildBtn=card.querySelector('[data-final-build]');
    setClassIfChanged(buildBtn,fresh?'btn ghost':'btn primary');
    if(publish){publish.disabled=!publishable;setClassIfChanged(publish,`btn secondary${publishable?'':' hidden'}`);}
    // 出力したPDFは別タブで見るだけでは掴めない。メールに添付する・共有フォルダーへ
    // 置くには実体が要るので、保存先そのものを画面に出し、フォルダーを開けるようにする。
    const pathBox=card.querySelector('[data-final-path]'),reveal=card.querySelector('[data-final-reveal]');
    if(pathBox){
      if(available){setHtmlIfChanged(pathBox,`<span class="final-path-label">保存先</span><span class="final-path-value" title="${escapeAttr(String(r.outputPdf||''))}">${escapeHtml(String(r.outputPdf||''))}</span>`);pathBox.classList.remove('hidden');}
      else{pathBox.classList.add('hidden');}
    }
    if(reveal){reveal.dataset.volume=volume;setClassIfChanged(reveal,`btn secondary${available?'':' hidden'}`);}
    card.classList.toggle('has-output',fresh);
  });
}


function normalizeCategoryValue(value, name='') {
  const cat = String(value || '').trim().toLowerCase();
  if (['ecm','bod','dmm'].includes(cat)) return cat;
  const n = String(name || '').toUpperCase();
  if (n.includes('ECM')) return 'ecm';
  if (n.includes('BOD')) return 'bod';
  if (n.includes('DMM')) return 'dmm';
  return '';
}
function workbookCategoryValue(w) {
  return normalizeCategoryValue(w?.category, w?.fileName || w?.displayName || w?.relativePath || w?.workbookId || '');
}
function registeredCategoryPathKey(relativePath, category) {
  return `${normPath(relativePath)}|${String(category || '').toLowerCase()}`;
}
function getRegisteredByRelativePath() {
  const map = new Map();
  for (const w of asArray(state?.structure?.workbooks)) {
    const cat = workbookCategoryValue(w);
    if (cat) map.set(registeredCategoryPathKey(w.relativePath, cat), w);
  }
  return map;
}
function isXlsxFile(file) {
  const name = String(file?.fileName || file?.relativePath || '').toLowerCase();
  return name.endsWith('.xlsx') || name.endsWith('.xlsm');
}
function categoryMatchesName(name, preset = activePreset) {
  const n = String(name || '').toUpperCase();
  if (preset === 'ecm') return n.includes('ECM');
  if (preset === 'bod') return n.includes('BOD');
  if (preset === 'dmm') return n.includes('DMM');
  return false;
}
function presetMatches(file, preset) {
  const accepted = asArray(packForPreset(preset)?.acceptedSourceTypes).map(x => String(x || '').toLowerCase());
  return (accepted.length ? accepted : ['excel','word','powerpoint','pdf']).includes(sourceTypeValue(file));
}
function workbookMatchesPreset(w, preset = activePreset) {
  if(activePackId){
    const explicit=String(w?.packId||'');
    if(explicit)return explicit===String(activePackId);
    const active=activePackRecord();
    if(active?.category)return workbookCategoryValue(w)===String(active.category).toLowerCase();
  }
  const category = workbookCategoryValue(w);
  if (['ecm','bod','dmm'].includes(category)) return category === preset;
  return categoryMatchesName(w?.fileName || w?.displayName || w?.relativePath || w?.workbookId, preset);
}
function packForPreset(preset = activePreset) {
  const category = String(preset || '').toLowerCase();
  const packs = asArray(state?.packs);
  if (category === String(activePreset || '').toLowerCase() && activePackId) {
    const selected = packs.find(pack => String(pack?.packId || '') === activePackId);
    if (selected) return selected;
  }
  return packs.find(pack => String(pack?.category || '').toLowerCase() === category)
    || asArray(state?.packs).find(pack => String(pack?.packId || '').toLowerCase() === `pack_${category}`)
    || null;
}
function activePackRecord() { return packForPreset(activePreset); }
function activePackIsBuiltIn(){return !!String(activePackRecord()?.category||'');}
function pageApiBody(extra={}){
  const pack=activePackRecord();
  const packId=String(pack?.packId||activePackId||'');
  // 並べ替えAPIは画面全体の並びを絶対値で送るため、読み込んだ時点の配置指紋を
  // 添えてサーバーに照合させる。他のタブが先に保存していれば409で拒否される。
  const body=Object.assign({packId},extra);
  const base=activeLayoutFingerprint(packId);
  if(base)body.baseLayout=base;
  return body;
}
function activeLayoutFingerprint(packId){
  return String(state?.layoutFingerprints?.[String(packId||'')]||'');
}
function rememberLayoutFingerprint(packId,fingerprint){
  const id=String(packId||''),value=String(fingerprint||'');
  if(!id||!value||!state)return;
  if(!state.layoutFingerprints)state.layoutFingerprints={};
  state.layoutFingerprints[id]=value;
}
function reconcileActivePackSelection() {
  const workflowPacks = asArray(state?.packs).filter(pack => pack?.workflowAvailable && !pack?.archived);
  let selected = workflowPacks.find(pack => String(pack?.packId || '') === activePackId)
    || workflowPacks.find(pack => String(pack?.category || '').toLowerCase() === String(activePreset || '').toLowerCase())
    || workflowPacks[0]
    || null;
  if (!selected) {
    activePackId='';activePreset='';
    try { sessionStorage.removeItem('ReportBinderPackId');sessionStorage.removeItem('ReportBinderCategory'); } catch {}
    return null;
  }
  activePackId = String(selected.packId || '');
  activePreset = String(selected.category||'').toLowerCase();
  try {
    sessionStorage.setItem('ReportBinderPackId', activePackId);
    sessionStorage.setItem('ReportBinderCategory', activePreset);
  } catch {}
  return selected;
}
function presetLabel(preset = activePreset) {
  return String(packForPreset(preset)?.displayName || '資料パック未作成');
}
function sourceTypeValue(source) {
  const explicit = String(source?.sourceType || '').trim().toLowerCase();
  if (explicit) return explicit;
  const path = String(source?.relativePath || source?.fileName || '').toLowerCase();
  if (path.endsWith('.xlsx') || path.endsWith('.xlsm') || path.endsWith('.xls')) return 'excel';
  if (path.endsWith('.docx') || path.endsWith('.doc')) return 'word';
  if (path.endsWith('.pptx')) return 'powerpoint';
  if (path.endsWith('.pdf')) return 'pdf';
  return 'file';
}
function sourceTypeLabel(source) {
  return ({excel:'Excel', word:'Word', powerpoint:'PowerPoint', pdf:'PDF', file:'ファイル'})[sourceTypeValue(source)] || 'ファイル';
}
function sourceTypeBadge(source) {
  const type = sourceTypeValue(source);
  return `<span class="source-type-badge ${escapeAttr(type)}">${escapeHtml(sourceTypeLabel(source))}</span>`;
}
function sourceIconId(source) { return ({excel:'i-file-excel',word:'i-file-word',powerpoint:'i-file-powerpoint',pdf:'i-file-pdf'})[sourceTypeValue(source)] || 'i-file-pdf'; }
function sourceUnitLabel(source) { return sourceTypeValue(source) === 'excel' ? 'シート' : sourceTypeValue(source)==='powerpoint'?'スライド':'ページ'; }
function sourceUnitReference(source, key) {
  const value = String(key || '');
  return sourceTypeValue(source) === 'excel' ? `シート ${value}` : sourceTypeValue(source)==='powerpoint'?`スライド ${value.replace(/^Slide\s+/i,'')}`:`ページ ${value.replace(/^Page\s+/i, '')}`;
}
function filesForActivePreset(files) {
  return asArray(files).filter(f => presetMatches(f, activePreset));
}
function workbooksForActivePreset(list = state?.structure?.workbooks) {
  return asArray(list).filter(w => workbookMatchesPreset(w, activePreset));
}
function pagesForActivePreset(list = state?.structure?.pages) {
  return asArray(list).filter(p => {
    const wb = getWorkbook(p.workbookId);
    return wb ? workbookMatchesPreset(wb, activePreset) : categoryMatchesName(p.workbookId || p.title || '', activePreset);
  });
}
function sheetNumberValue(sheetName) {
  const s = String(sheetName ?? '').trim();
  return /^\d+$/.test(s) ? Number(s) : Number.MAX_SAFE_INTEGER;
}
function fileOrderValue(name) {
  const m = String(name || '').match(/_(\d{1,4})_/);
  return m ? Number(m[1]) : Number.MAX_SAFE_INTEGER;
}
function updateCategoryLabels() { /* category is always visible in the global header */ }

function updatePresetTabs() {
  reconcileActivePackSelection();
  renderPackSwitcher();
  updateCategoryLabels();
}
function packMenuRow(pack, archived=false) {
  const id=String(pack?.packId||''),name=String(pack?.displayName||'名称未設定'),workflow=!!pack?.workflowAvailable;
  const active=!archived&&workflow&&id===activePackId;
  const status=archived?'アーカイブ済み':(workflow?(active?'利用中':'選択して使用'):'現在は利用できません');
  const selectAttrs=archived?'disabled':(workflow?`data-pack-select="${escapeAttr(id)}"`:`data-pack-unavailable="${escapeAttr(id)}"`);
  const archiveAction=!archived&&pack?.archivable?`<button class="danger" type="button" data-pack-action="archive" data-pack-id="${escapeAttr(id)}">アーカイブ</button>`:'';
  const restoreAction=archived?`<button type="button" data-pack-action="restore" data-pack-id="${escapeAttr(id)}">復元</button>`:'';
  const templateUpgrade=!archived&&pack?.templateUpdateAvailable?`<button class="attention" type="button" data-pack-action="template-upgrade" data-pack-id="${escapeAttr(id)}">ひな形更新</button>`:'';
  return `<div class="pack-menu-row" data-pack-row="${escapeAttr(id)}"><button class="pack-option${active?' active':''}" type="button" ${selectAttrs} ${active?'aria-current="true"':''}><span class="pack-option-marker" aria-hidden="true"></span><span class="pack-option-copy"><strong>${escapeHtml(name)}</strong><small>${escapeHtml(status)}${pack?.templateUpdateAvailable?`・ひな形 v${Number(pack.availableTemplateVersion||1)}あり`:''}</small></span></button><div class="pack-row-actions">${restoreAction}${templateUpgrade}<button type="button" data-pack-action="duplicate" data-pack-id="${escapeAttr(id)}">複製</button><button type="button" data-pack-action="rename" data-pack-id="${escapeAttr(id)}">名前を変更</button>${archiveAction}</div></div>`;
}
function renderPackSwitcher() {
  const selected=activePackRecord();
  const label=$('active-pack-name'),button=$('pack-menu-button');
  if(label)label.textContent=String(selected?.displayName||(configured()?'名前を付ける':'未作成'));
  if(button){button.disabled=!configured()&&!selected;button.classList.toggle('empty',!selected);button.setAttribute('aria-label',selected?`資料パック「${String(selected.displayName||'名称未設定')}」を切り替え・管理`:(configured()?'今回まとめる資料パックに名前を付ける':'原稿フォルダーを設定すると資料パックを作成できます'));button.setAttribute('aria-haspopup',selected?'dialog':'true');}
  const all=packAdminPacks.length?packAdminPacks:asArray(state?.packs);
  const active=all.filter(pack=>!pack?.archived),archived=all.filter(pack=>!!pack?.archived);
  const list=$('pack-menu-list');if(list)list.innerHTML=active.length?active.map(pack=>packMenuRow(pack,false)).join(''):'<div class="pack-menu-empty">使用できる資料パックがありません。</div>';
  const toggle=$('toggle-archived-packs');
  if(toggle){toggle.classList.toggle('hidden',archived.length===0);toggle.textContent=`アーカイブ済み（${archived.length}件）${archivedPacksExpanded?'を閉じる':'を表示'}`;toggle.setAttribute('aria-expanded',String(archivedPacksExpanded));}
  const archivedList=$('archived-pack-list');
  if(archivedList){archivedList.classList.toggle('hidden',!archivedPacksExpanded||archived.length===0);archivedList.innerHTML=archived.map(pack=>packMenuRow(pack,true)).join('');}
}
// Non-modal popovers must not linger behind the focus ring: if Tab (or a click)
// takes focus out of them, close them the same way an outside click does.
function closePopoverOnFocusOut(container, close){
  if(!container||container.dataset.focusOutBound==='1')return;
  container.dataset.focusOutBound='1';
  container.addEventListener('focusout',event=>{
    if(container.classList.contains('hidden'))return;
    const next=event.relatedTarget;
    if(next&&(container.contains(next)||next.getAttribute?.('aria-controls')===container.id))return;
    if(next===null)return; // focus left the document entirely; keep it open
    close(false);
  });
}
function focusFirstIn(container){
  if(!container)return;
  const target=container.querySelector('a[href],button:not([disabled]),input:not([disabled]),select:not([disabled]),[tabindex]:not([tabindex="-1"])');
  target?.focus();
}
function closePackMenu(returnFocus=false) {
  const menu=$('pack-menu'),button=$('pack-menu-button');if(!menu)return;
  menu.classList.add('hidden');button?.setAttribute('aria-expanded','false');
  if(returnFocus)button?.focus();
}
async function loadPackAdminPacks() {
  const response=await api('/api/v2/packs?includeArchived=true');
  packAdminPacks=asArray(response.packs);
  renderPackSwitcher();
  return packAdminPacks;
}
async function togglePackMenu() {
  const menu=$('pack-menu'),button=$('pack-menu-button');if(!menu||!button)return;
  if(!configured()){setActiveView('excel');setTimeout(()=>$('change-source-folder-btn')?.focus(),0);return;}
  if(!activePackRecord()&&!asArray(state?.packs).length){openPackEditor('create');return;}
  if(!menu.classList.contains('hidden')){closePackMenu(true);return;}
  menu.classList.remove('hidden');button.setAttribute('aria-expanded','true');renderPackSwitcher();
  try{await loadPackAdminPacks();}catch(error){showMessage('danger','資料パック一覧を読み込めません',userFriendlyError(error.message),error.detail||error.message);}
  // role="dialog" + aria-modal announces "dialog" but focus stayed on the
  // trigger, so Tab walked straight past the menu into the page behind it.
  // Must run after loadPackAdminPacks(), which rebuilds the list innerHTML.
  if(!menu.classList.contains('hidden')){closePopoverOnFocusOut(menu,closePackMenu);focusFirstIn(menu);}
}
function visibleUnregisteredFiles(files) {
  const registered=new Set(workbooksForActivePreset().map(w=>String(w.relativePath||'').replace(/\\/g,'/').toLowerCase()));
  return filesForActivePreset(files).filter(f => !registered.has(String(f.relativePath||'').replace(/\\/g,'/').toLowerCase()));
}

function capturePageBoardScroll() {
  const focusedRow=document.activeElement?.closest?.('.page-row');
  const snap = {windowX: window.scrollX || 0, windowY: window.scrollY || 0, focusPageId:String(focusedRow?.dataset?.pageId||''), wraps: {}};
  document.querySelectorAll('[data-volume]').forEach(tbody => {
    const volume = tbody.getAttribute('data-volume') || '';
    const wrap = tbody.closest('.table-wrap,.thumbnail-wrap');
    if (volume && wrap) snap.wraps[volume] = {top: wrap.scrollTop || 0, left: wrap.scrollLeft || 0};
  });
  return snap;
}
function restorePageBoardScroll(snap) {
  if (!snap) return;
  requestAnimationFrame(() => {
    for (const [volume, pos] of Object.entries(snap.wraps || {})) {
      const tbody = [...document.querySelectorAll('[data-volume]')].find(t => t.getAttribute('data-volume') === volume);
      const wrap = tbody?.closest?.('.table-wrap,.thumbnail-wrap');
      if (wrap) { wrap.scrollTop = pos.top || 0; wrap.scrollLeft = pos.left || 0; }
    }
    if (Number.isFinite(snap.windowY)) window.scrollTo(snap.windowX || 0, snap.windowY || 0);
    if(snap.focusPageId){const focused=[...document.querySelectorAll('.page-row')].find(row=>String(row.dataset.pageId||'')===snap.focusPageId&&!row.classList.contains('page-filter-hidden'));focused?.focus({preventScroll:true});}
  });
}
function updateSelectionRange(values, selectedSet, clickedValue, checked, shiftKey, anchorValue) {
  const list = values.map(v => String(v || '')).filter(Boolean);
  const clicked = String(clickedValue || '');
  const currentIndex = list.indexOf(clicked);
  if (currentIndex >= 0 && shiftKey && anchorValue) {
    const anchorIndex = list.indexOf(String(anchorValue || ''));
    if (anchorIndex >= 0) {
      const from = Math.min(anchorIndex, currentIndex);
      const to = Math.max(anchorIndex, currentIndex);
      for (let i = from; i <= to; i++) {
        if (checked) selectedSet.add(list[i]); else selectedSet.delete(list[i]);
      }
      return clicked;
    }
  }
  if (checked) selectedSet.add(clicked); else selectedSet.delete(clicked);
  return clicked;
}
function syncTableSelectionUi(box, checkboxSelector, selectAllSelector, selectedSet, values) {
  if (!box) return;
  const list = values.map(v => String(v || '')).filter(Boolean);
  const allSelected = list.length > 0 && list.every(id => selectedSet.has(id));
  const someSelected = list.some(id => selectedSet.has(id));
  const selectAll = box.querySelector(selectAllSelector);
  if (selectAll) {
    selectAll.checked = allSelected;
    selectAll.indeterminate = someSelected && !allSelected;
  }
  box.querySelectorAll(checkboxSelector).forEach(ch => {
    const checked = selectedSet.has(String(ch.value || ''));
    ch.checked = checked;
    ch.closest('tr')?.classList.toggle('selected-row', checked);
  });
}
function visibleFileSelectionValues(files=availableFiles) {
  return visibleUnregisteredFiles(files).map(f => String(f.relativePath || '')).filter(Boolean);
}
function syncFileSelectionUi(files=availableFiles) {
  syncTableSelectionUi($('file-list'), '[data-file-check]', '[data-select-all-files]', selectedFiles, visibleFileSelectionValues(files));
  const count=selectedFiles.size,bar=$('file-context-bar'),label=$('file-context-count'),register=$('register-selected-btn'),clear=$('clear-selected-files-btn');
  if(bar)bar.classList.toggle('has-selection',count>0);
  if(label)label.textContent=`${count}件を選択中`;
  if(register)register.disabled=count===0;
  if(clear)clear.disabled=count===0;
}
function handleFileCheckboxToggle(ch, shiftKey=false) {
  lastFileRangeAnchor = updateSelectionRange(visibleFileSelectionValues(), selectedFiles, ch.value, ch.checked, shiftKey, lastFileRangeAnchor);
  syncFileSelectionUi();
}
function visibleWorkbookSelectionValues() {
  return workbooksForActivePreset().map(w => String(w.workbookId || '')).filter(Boolean);
}
function syncWorkbookSelectionUi() {
  syncTableSelectionUi($('workbook-list'), '[data-workbook-check]', '[data-select-all-workbooks]', selectedWorkbooks, visibleWorkbookSelectionValues());
  const count=selectedWorkbooks.size,bar=$('workbook-context-bar'),label=$('workbook-context-count'),render=$('render-selected-btn'),clear=$('clear-selected-workbooks-btn'),unregister=$('unregister-selected-btn');
  if(bar)bar.classList.toggle('has-selection',count>0);
  if(label)label.textContent=`${count}件を選択中`;
  if(render)render.disabled=count===0;
  if(clear)clear.disabled=count===0;
  if(unregister)unregister.disabled=count===0;
  if(count===0)bar?.querySelector('details')?.removeAttribute('open');
  updateRenderTargetUi();
}
function handleWorkbookCheckboxToggle(ch, shiftKey=false) {
  lastWorkbookRangeAnchor = updateSelectionRange(visibleWorkbookSelectionValues(), selectedWorkbooks, ch.value, ch.checked, shiftKey, lastWorkbookRangeAnchor);
  syncWorkbookSelectionUi();
}
function visiblePageSelectionValues() {
  return [...document.querySelectorAll('.page-row:not(.page-filter-hidden)')].map(row => String(row.getAttribute('data-page-id') || '')).filter(Boolean);
}
function syncPageSelectionUi() {
  updateBulkSelectionLabel();
  document.querySelectorAll('.page-row').forEach(row => {
    const id = String(row.getAttribute('data-page-id') || '');
    const checked = selectedPages.has(id);
    row.classList.toggle('selected-row', checked);
    // aria-selected は tr(role=row) でのみ有効。article のカードでは無視されるため付けない。
    // カード側の選択状態は、同梱の sr-only チェックボックス(下の ch.checked)が公開する。
    if(row.matches('tr'))row.setAttribute('aria-selected',String(checked));
    else row.removeAttribute('aria-selected');
    const ch = row.querySelector('[data-page-check]');
    if (ch) ch.checked = checked;
  });
  document.querySelectorAll('[data-select-all-pages]').forEach(ch => {
    const volumeName = String(ch.dataset.selectAllPages || '');
    const tbody = [...document.querySelectorAll('tbody[data-volume]')].find(t => t.getAttribute('data-volume') === volumeName);
    const ids = tbody ? [...tbody.querySelectorAll('.page-row')].map(r => String(r.getAttribute('data-page-id') || '')).filter(Boolean) : [];
    const allSelected = ids.length > 0 && ids.every(id => selectedPages.has(id));
    const someSelected = ids.some(id => selectedPages.has(id));
    ch.checked = allSelected;
    ch.indeterminate = someSelected && !allSelected;
  });
}
function handlePageCheckboxToggle(ch, shiftKey=false) {
  lastPageRangeAnchor = updateSelectionRange(visiblePageSelectionValues(), selectedPages, ch.value, ch.checked, shiftKey, lastPageRangeAnchor);
  syncPageSelectionUi();
}
function currentFilteredFileSelectionValues() {
  const query = String(fileFilterText || '').trim().toLowerCase();
  return visibleUnregisteredFiles(availableFiles)
    .filter(f => !query || `${fileDisplayName(f)} ${f.relativePath || ''}`.toLowerCase().includes(query))
    .map(f => String(f.relativePath || ''))
    .filter(Boolean);
}
function selectVisibleFiles() {
  const values = currentFilteredFileSelectionValues();
  values.forEach(id => selectedFiles.add(id));
  lastFileRangeAnchor = values[values.length - 1] || '';
  syncFileSelectionUi();
  renderFileList(availableFiles);
}
function clearVisibleFilesSelection() {
  currentFilteredFileSelectionValues().forEach(id => selectedFiles.delete(id));
  lastFileRangeAnchor = '';
  syncFileSelectionUi();
  renderFileList(availableFiles);
}
function clearWorkbookSelection() {
  selectedWorkbooks.clear();
  lastWorkbookRangeAnchor = '';
  syncWorkbookSelectionUi();
}
function checkedWorkbookIdsFromUi() {
  const visible = new Set(visibleWorkbookSelectionValues());
  return [...document.querySelectorAll('#workbook-list [data-workbook-check]:checked')]
    .map(ch => String(ch.value || '').trim())
    .filter(id => id && visible.has(id));
}
function normalizeWorkbookIdListForActivePreset(ids) {
  const visible = new Set(visibleWorkbookSelectionValues());
  const out = [];
  const seen = new Set();
  for (const raw of asArray(ids)) {
    const id = String(raw || '').trim();
    if (!id || seen.has(id) || !visible.has(id)) continue;
    seen.add(id);
    out.push(id);
  }
  return out;
}
function selectAllActivePages() {
  const ids = visiblePageSelectionValues();
  ids.forEach(id => selectedPages.add(id));
  lastPageRangeAnchor = ids[ids.length - 1] || '';
  syncPageSelectionUi();
  if (!ids.length) showMessage('warn', 'ページがありません', '登録済み原稿をPDF変換するとページが表示されます。');
}
function clearActivePageSelection() {
  pagesForActivePreset().map(p => resolvedPageId(p)).filter(Boolean).forEach(id => selectedPages.delete(id));
  lastPageRangeAnchor = '';
  syncPageSelectionUi();
}

async function moveSelectedPagesToVolume(volume, btn) {
  await movePageIdsToVolume([...selectedPages], volume, btn);
}
// Shared by the 移動先 buttons and by dropping a drag onto a volume tab, so both
// routes produce the same reorder request, undo entry and toast.
async function movePageIdsToVolume(pageIds, volume, btn) {
  const target=String(volume||'');const ids=asArray(pageIds).map(String).filter(id=>!!getPage(id));
  if(!ids.length){showMessage('warn','ページを選択してください','移動するページにチェックを入れてください。');return;}
  if(pageMutationBusy)return;
  pageMutationBusy=true;
  updateBulkSelectionLabel();
  try{
  await runBusy(btn,async()=>{
    const beforeVolumes=collectBoardVolumes();
    const first=activeTargetVolumes()[0]||mainVolume(),volumes=Object.fromEntries([...activeTargetVolumes(),'none'].map(key=>[key,[]])),selected=new Set(ids);
    for(const p of [...pagesForActivePreset()].sort(pageSort)){const id=resolvedPageId(p);if(!id||selected.has(id))continue;const v=(p.enabled===false||String(p.volume||first)==='none')?'none':String(p.volume||first);if(!volumes[v])volumes[v]=[];volumes[v].push(id);}
    if(!volumes[target])volumes[target]=[];volumes[target].push(...ids);
    const response=await api('/api/v2/items/reorder',{method:'POST',body:pageApiBody({volumes})});applyPageMutationResult(response);rememberPageLayoutUndo(`${ids.length}ページの移動`,beforeVolumes);selectedPages.clear();lastPageRangeAnchor='';lastPageBoardRenderSignature='';renderPages();
    showMessage('ok','ページ構成を保存しました',`${ids.length}ページを${assignmentVolumeLabel(target)}へ移動しました。`,null,[{label:'元に戻す',handler:()=>undoLastPageLayout()},{label:'提出用PDFへ',view:'final'}],8000);
  });
  }finally{pageMutationBusy=false;updateBulkSelectionLabel();}
  syncPageSelectionUi();
}

async function loadFiles(btn) {
  if (!configured()) {
    showMessage('warn', '原稿フォルダーを選んでください', 'フォルダーを選ぶと原稿を表示できます。');
    return [];
  }
  await runBusy(btn, async () => {
    const data = await api('/api/submission-files');
    availableFiles = asArray(data.files);
    availableFilesScannedAt = String(data.scannedAt || '');
    selectedFiles = new Set([...selectedFiles].filter(r => visibleUnregisteredFiles(availableFiles).some(f => normPath(f.relativePath) === normPath(r))));
    renderFileList(availableFiles);
    renderFolderOverview();
    renderDashboardOverview();
    showMessage(availableFiles.length ? 'ok' : 'warn', availableFiles.length ? '原稿一覧を更新しました' : '対応する原稿が見つかりません', '原稿フォルダー直下の対応形式だけを表示しています。');
  }, false);
  return availableFiles;
}
function syncSourcePanelDensity() {
  const grid=document.querySelector('.excel-grid'),unregisteredCard=document.querySelector('.unregistered-card'),registeredCard=document.querySelector('.registered-card');
  if(!grid||!unregisteredCard||!registeredCard)return;
  const unregisteredCount=configured()?visibleUnregisteredFiles(availableFiles).length:0;
  const registeredCount=workbooksForActivePreset().length;
  const showUnregistered=unregisteredCount>0||registeredCount===0;
  const showRegistered=registeredCount>0;
  unregisteredCard.classList.toggle('hidden',!showUnregistered);
  registeredCard.classList.toggle('hidden',!showRegistered);
  grid.classList.toggle('single-panel',showUnregistered!==showRegistered);
  const next=$('source-next-action'),needsPdf=workbooksForActivePreset().filter(workbook=>!isLatestPdfWorkbook(workbook)).length;
  if(next)next.classList.toggle('hidden',registeredCount===0||needsPdf>0);
}
function renderFileList(files) {
  queueMicrotask(syncSourcePanelDensity);
  const box = $('file-list');
  if (!box) return;
  updatePresetTabs();
  if (!configured()) {
    box.className = 'table-shell empty-state';
    box.textContent = '上の「フォルダーを選ぶ」から原稿フォルダーを設定してください。';
    setFileSelectionSummary(0, 0);
    return;
  }
  const query = String(fileFilterText || '').trim().toLowerCase();
  const allSelectable = visibleUnregisteredFiles(files);
  const selectable = query
    ? allSelectable.filter(f => `${fileDisplayName(f)} ${f.relativePath || ''}`.toLowerCase().includes(query))
    : allSelectable;
  const selectedCount = () => [...selectedFiles].filter(id => allSelectable.some(f => String(f.relativePath || '') === id)).length;
  if (!selectable.length) {
    box.className = 'table-shell empty-state';
    // 「未登録の原稿はありません。」は3つの別の状況で出ていた: 全部登録済み、
    // フォルダー直下に対応形式が無い、資料パックの受け入れ形式で弾かれた。
    // どれも「作業完了」と読めるため、原因に気付けず黙って先へ進んでしまう。
    box.innerHTML = allSelectable.length
      ? '<div><strong>検索条件に一致する未登録原稿はありません。</strong></div>'
      : (files.length
        ? `<div><strong>この資料パックに登録できる原稿はありません。</strong></div><div class="subtext">フォルダーには ${files.length} 件ありますが、すべて登録済みか、この資料パックが受け付けない形式です。</div>`
        : '<div><strong>このフォルダーの直下に原稿が見つかりません。</strong></div><div class="subtext">対応する形式は .xlsx / .xlsm / .docx / .pptx / .pdf です。サブフォルダーの中は探しません。</div>');
    if (!allSelectable.length) { selectedFiles.clear(); lastFileRangeAnchor = ''; }
    setFileSelectionSummary(allSelectable.length, selectedCount());
    syncFileSelectionUi(files);
    return;
  }
  selectedFiles = new Set([...selectedFiles].filter(id => allSelectable.some(f => String(f.relativePath || '') === id)));
  const allSelected = selectable.length > 0 && selectable.every(f => selectedFiles.has(String(f.relativePath || '')));
  const someSelected = selectable.some(f => selectedFiles.has(String(f.relativePath || '')));
  box.className = 'table-shell';
  box.innerHTML = `<table class="data-table file-table">
    <thead><tr><th class="check-col"><input type="checkbox" data-select-all-files ${allSelected ? 'checked' : ''} aria-label="表示中の未登録原稿をすべて選択"></th><th>原稿ファイル</th></tr></thead>
    <tbody>${selectable.map(f => {
      const rel = String(f.relativePath || '');
      const checked = selectedFiles.has(rel);
      const displayName = fileDisplayName(f);
      const detail = normPath(rel) !== normPath(displayName) ? `<div class="subtext">${escapeHtml(rel)}</div>` : '';
      const modified = formatSubmissionFileModifiedAt(f);
      return `<tr data-file-row class="${checked ? 'selected-row' : ''}">
        <td class="check-col"><input type="checkbox" data-file-check value="${escapeAttr(rel)}" ${checked ? 'checked' : ''} aria-label="${escapeAttr(displayName)}を選択"></td>
        <td><div class="file-name-cell"><span class="file-icon ${escapeAttr(sourceTypeValue(f))}">${iconUse(sourceIconId(f))}</span><div><div class="source-name-line"><strong title="${escapeAttr(displayName)}">${escapeHtml(displayName)}</strong>${sourceTypeBadge(f)}</div>${detail}<div class="file-meta" title="原稿ファイルの最終保存日時">更新日時：${escapeHtml(modified)}</div></div></div></td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
  setFileSelectionSummary(allSelectable.length, selectedCount());
  syncFileSelectionUi(files);
  const selectAll = box.querySelector('[data-select-all-files]');
  if (selectAll) {
    selectAll.indeterminate = someSelected && !allSelected;
    selectAll.addEventListener('change', (e) => {
      if (e.target.checked) {
        for (const f of selectable) selectedFiles.add(String(f.relativePath || ''));
      } else {
        for (const f of selectable) selectedFiles.delete(String(f.relativePath || ''));
      }
      lastFileRangeAnchor = '';
      syncFileSelectionUi(files);
      renderFileList(files);
    });
  }
  box.querySelectorAll('[data-file-check]').forEach(ch => ch.addEventListener('click', (e) => {
    e.stopPropagation();
    handleFileCheckboxToggle(ch, e.shiftKey);
    renderFileList(files);
  }));
  box.querySelectorAll('[data-file-row]').forEach(row => row.addEventListener('click', (e) => {
    if (e.target.closest('input, button, a, select, textarea')) return;
    const ch = row.querySelector('[data-file-check]');
    if (!ch) return;
    ch.checked = !ch.checked;
    handleFileCheckboxToggle(ch, e.shiftKey);
    renderFileList(files);
  }));
}
async function applyPresetSelection(presetOrButton) {
  // Switching packs mid-job made the finishing job apply its results (selection,
  // highlights, progress counts) to whichever pack happened to be active.
  if(renderJobActive||activeFinalJobId){showMessage('warn','処理中は資料パックを切り替えられません','PDFの作成が終わってから切り替えてください。中止する場合は進捗パネルの「中止」を押してください。');return;}
  const requested=typeof presetOrButton==='string'?presetOrButton:(presetOrButton?.dataset?.packSelect||presetOrButton?.dataset?.preset||activePackId||activePreset);
  const pack=asArray(state?.packs).find(item=>String(item?.packId||'')===String(requested))
    || asArray(state?.packs).find(item=>String(item?.category||'').toLowerCase()===String(requested||'').toLowerCase());
  if(!pack?.workflowAvailable){showMessage('warn','この資料パックは利用できません','アーカイブ済みの場合は復元してから選択してください。');return;}
  const preset=String(pack.category||'').toLowerCase();
  const nextPackId=String(pack.packId||'');const presetChanged=preset!==activePreset||nextPackId!==activePackId;
  if(presetChanged&&pendingPageLayoutUndo){if(boardSaveTimer){clearTimeout(boardSaveTimer);boardSaveTimer=null;}await saveBoardOrder();}
  activePreset=preset;activePackId=nextPackId;try{sessionStorage.setItem('ReportBinderCategory',activePreset);sessionStorage.setItem('ReportBinderPackId',activePackId);}catch{}
  selectedFiles.clear();selectedWorkbooks.clear();selectedPages.clear();lastFileRangeAnchor='';lastWorkbookRangeAnchor='';lastPageRangeAnchor='';lastPageBoardRenderSignature='';
  if(presetChanged)clearPageLayoutHistory();else updateBulkSelectionLabel();pageVolumeCollapseOverrides.clear();
  if(configured()&&!availableFiles.length)await loadFilesSilently();
  if(activeView==='history'){
    historyPanelsInitialized=false;
    snapshotHistoryLoadSerial++;
  }
  if(activeView==='final'&&presetChanged){packReviewState=null;void loadPackReview();}
  renderAll();
}

async function registerSelected(btn) {
  const rels=[...selectedFiles];if(!rels.length){showMessage('warn','原稿を選択してください','登録するExcel・Word・PowerPoint・PDFにチェックを入れてください。');return;}
  // 資料パック未作成のまま送ると、サーバーが内部互換用の識別子(ecm/bod/dmm)を含む
  // 例外を返し、対処のわからないエラーだけが残る。手前で必要な操作へ案内する。
  if(!activePackRecord()){
    showMessage('warn','先に「今回まとめる一式」に名前を付けてください',
      '名前を付けると、選んだ原稿をそこへ登録できます。',null,
      [{label:'名前を付ける',primary:true,handler:()=>openPackEditor('create')}],0);
    return;
  }
  await runBusy(btn,async()=>{
    const data=await api('/api/v2/sources/register-batch',{method:'POST',body:{relativePaths:rels,packId:activePackId}});const result=data.result||{},registered=asArray(result.registered),errors=asArray(result.errors);
    // 登録APIのstateは汎用V2ドメインで、旧UIが参照するworkbooks/pagesを含まない。
    // 互換viewを持つ正規stateを取り直し、登録直後に0件へ見える状態ずれを防ぐ。
    selectedFiles.clear();lastFileRangeAnchor='';await refresh();
    selectedWorkbooks=new Set(registered.map(x=>String(x.workbookId||'')).filter(Boolean));lastWorkbookRangeAnchor='';renderAll();await loadFiles(null);
    if(registered.length&&errors.length===0)showMessage('ok',`${registered.length}件を登録しました`,`登録した${registered.length}件が選択されています。`,null,[{label:`登録した${registered.length}件をPDF作成`,primary:true,view:'excel',handler:()=>renderSelectedWorkbooks($('render-selected-btn'))}],0);
    else if(registered.length)showMessage('warn',`${registered.length}件を登録しました`,`${errors.length}件は登録できませんでした。`,errors,[{label:'登録分をPDF作成',view:'excel',handler:()=>renderSelectedWorkbooks($('render-selected-btn'))}],0);
    else showMessage('warn','登録できませんでした','詳細を確認してください。',errors.length?errors:result,[],0);
  });
}


async function unregisterSelected(btn) {
  const ids = [...selectedWorkbooks];
  if (!ids.length) { showMessage('warn', '原稿を選択してください', '登録解除する原稿にチェックを入れてください。'); return; }
  const accepted=await confirmAction({title:'原稿の登録を解除しますか？',message:`選択した${ids.length}件をReportBinderの管理対象から外します。`,detail:'作成済みの変換PDFと元の原稿ファイルは削除されません。',confirmLabel:'登録を解除',danger:true});
  if (!accepted) return;
  await runBusy(btn, async () => {
    for (const workbookId of ids) await api('/api/workbooks/unregister', {method:'POST', body:{workbookId}});
    selectedWorkbooks.clear();
    selectedPages.clear();
    lastWorkbookRangeAnchor = '';
    lastPageRangeAnchor = '';
    await refresh();
    await loadFiles(null);
    showMessage('ok', '登録を解除しました', `${ids.length}件を解除しました。`);
  });
}

// ---- V5: シート単位の変更表示 ----
// 表現は「見落としなし」ではなく「提出用PDFの見た目を基準とした高精度な判定」。
let showChangedOnly = false;
function changeSummaryFor(workbookId){
  const map = state?.changeSummaries;
  if (!map) return null;
  return map[String(workbookId||'')] || null;
}
function changedSheetSet(workbookId){
  return new Set(asArray(changeSummaryFor(workbookId)?.changedSheets).map(String));
}
function addedSheetSet(workbookId){
  return new Set(asArray(changeSummaryFor(workbookId)?.addedSheets).map(String));
}
function unknownSheetSet(workbookId){
  return new Set(asArray(changeSummaryFor(workbookId)?.unknownSheets).map(String));
}
function comparisonSummaryForWorkbook(w){
  if (!state?.inputHistoryEnabled) return '';
  const cs = changeSummaryFor(w?.workbookId);
  if (!cs) return '';
  if (String(cs.status||'') !== 'complete') return '見た目を比較できません';
  const addedNames=new Set(asArray(cs.addedSheets).map(String));
  const removedNames=new Set(asArray(cs.removedSheets).map(String));
  const unknownNames=new Set(asArray(cs.unknownSheets).map(String));
  const changed = asArray(cs.changedSheets).map(String).filter(name=>!addedNames.has(name)&&!removedNames.has(name)&&!unknownNames.has(name)).length;
  const added = addedNames.size;
  const unknown = unknownNames.size;
  // Added and removed sheets are changes too. Keep them visible in the workbook
  // summary even though only a current (added) sheet can have a page row.
  const removed = removedNames.size;
  const affected=changed+added+removed;
  if (affected > 0) return `見た目変更 ${affected}${sourceUnitLabel(w)}`;
  if (unknown > 0) return `見た目を判定できません ${unknown}${sourceUnitLabel(w)}`;
  return '見た目変更なし';
}
function workbookPdfStatusCell(w){
  const status = String(w?.status || '');
  const detail = latestPdfDetailForWorkbook(w);
  if (status === 'render-error') {
    return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('PDF作成エラー','danger')}</div><div class="pdf-status-detail">${escapeHtml(detail || '差分を確認できません')}</div></div>`;
  }
  if (status === 'missing') {
    // ファイルが見つからない原稿を「原稿更新あり／PDFを再作成してください」と
    // 案内していた。その指示どおり押しても、対象の絞り込みが missing を外すため
    // 何も起きず、同時にヒントは「すべての変換PDFは最新です。」と出ていた。
    return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('ファイルなし','danger')}</div><div class="pdf-status-detail">原稿フォルダーにこのファイルがありません。名前を変えた・移動した・削除した場合は、下のボタンで別のファイルに付け替えてください。</div><div class="pdf-status-actions"><button class="btn secondary" type="button" data-relink-source="${escapeAttr(w?.workbookId||'')}">別のファイルに付け替える</button></div></div>`;
  }
  if (status === 'rendering') {
    return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('PDF作成中','attention')}</div><div class="pdf-status-detail">完了後に差分を判定します</div></div>`;
  }
  if (!isLatestPdfWorkbook(w)) {
    return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('原稿更新あり','attention')}</div><div class="pdf-status-detail">PDFを再作成して更新を反映してください</div></div>`;
  }
  const change = comparisonSummaryForWorkbook(w);
  const canCompare=String(changeSummaryFor(w?.workbookId)?.status||'')==='complete';
  const comparisonLine=change?(canCompare?`<button class="pdf-comparison-link" type="button" data-open-history="${escapeAttr(w?.workbookId||'')}" aria-label="${escapeAttr(`${workbookDisplayName(w)}、${change}、詳細を見る`)}"><span>${escapeHtml(change)}</span><strong>詳細を見る</strong></button>`:`<div class="pdf-status-detail">${escapeHtml(change)}</div>`):'';
  return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('変換PDFは最新','neutral')}</div>${comparisonLine}</div>`;
}

function openWorkbookComparisonHistory(workbookId){
  const id=String(workbookId||'');if(!id)return;
  snapshotHistoryState.workbookId=id;
  snapshotHistoryState.snapshots=[];
  snapshotHistoryState.fromId='';
  snapshotHistoryState.toId='';
  try{sessionStorage.setItem('ReportBinderSnapshotWorkbook',id);}catch{}
  setActiveView('history',{reloadPanels:true});
}
function pageChangeBadge(p){
  if (!state?.inputHistoryEnabled) return '';
  const cs = changeSummaryFor(p?.workbookId);
  if (!cs || String(cs.status||'') !== 'complete') return '';
  const name = String(p?.sheetName||'');
  if (changedSheetSet(p.workbookId).has(name)) return badge('変更あり','attention');
  if (addedSheetSet(p.workbookId).has(name)) return badge('追加','attention');
  if (unknownSheetSet(p.workbookId).has(name)) return badge('判定不能','neutral');
  return '';
}
function pageIsChanged(p){
  const cs = changeSummaryFor(p?.workbookId);
  if (!cs || String(cs.status||'') !== 'complete') return true;
  const name = String(p?.sheetName||'');
  return changedSheetSet(p.workbookId).has(name) || addedSheetSet(p.workbookId).has(name) || unknownSheetSet(p.workbookId).has(name);
}

// ---- 形式共通の差分詳細ダイアログ ----
function isDiffModalOpen(){
  const modal=$('diff-modal');
  return !!modal&&!modal.classList.contains('hidden');
}
function diffKindMeta(kind){
  const map={
    modified:{label:'変更',mark:'M'},
    added:{label:'追加',mark:'A'},
    removed:{label:'削除',mark:'D'},
    unknown:{label:'判定不能',mark:'?'},
    unchanged:{label:'変更なし',mark:'—'}
  };
  return map[String(kind||'')]||map.unknown;
}
function diffFilteredSheets(){
  const sheets=asArray(diffViewState.detail?.sheets);
  let filtered=diffViewState.filter==='all'?sheets:diffViewState.filter==='confirmed'?sheets.filter(s=>!!s.confirmed):sheets.filter(s=>String(s.kind||'')===diffViewState.filter);
  if(diffViewState.unreviewedOnly)filtered=filtered.filter(s=>!s.confirmed);
  return filtered;
}

async function loadPackReview(options = {}){
  const pack=activePackRecord();if(!state||!configured()||!pack?.packId)return;
  const packId=String(pack.packId);
  if(packReviewInFlight)return packReviewInFlight;
  packReviewInFlight=api(`/api/v2/packs/${encodeURIComponent(packId)}/review`).then(response=>{
    if(String(activePackRecord()?.packId||'')!==packId)return;
    packReviewState=response.review||null;if(options.render!==false)renderPackReview();
  }).catch(error=>{log('レビュー状態の確認でエラー',error.detail||error.message);}).finally(()=>{packReviewInFlight=null;});
  return packReviewInFlight;
}
function diffCurrentSheet(){
  const sheets=asArray(diffViewState.detail?.sheets);
  return sheets.find(s=>String(s.sheetKey||'')===diffViewState.selectedSheetKey)||null;
}
function diffSource(){
  return getWorkbook(diffViewState.workbookId)||{sourceType:diffViewState.detail?.sourceType||'excel'};
}
function diffUnitDisplayName(sheet){
  const source=diffSource();
  const before=String(sheet?.beforeSheetName||'');
  const after=String(sheet?.afterSheetName||'');
  const beforeLabel=before?sourceUnitReference(source,before):'';
  const afterLabel=after?sourceUnitReference(source,after):'';
  if(beforeLabel&&afterLabel&&before!==after)return `${beforeLabel} → ${afterLabel}`;
  return afterLabel||beforeLabel||String(sheet?.sheetName||'項目');
}
function diffMatchConfidenceLabel(sheet){
  const kind=String(sheet?.kind||'unknown');
  if(kind==='added'||kind==='removed')return '対応対象なし';
  const confidence=Math.max(0,Math.min(1,Number(sheet?.matchConfidence??1)));
  if(kind==='unknown'||confidence<.5)return `対応不確実 ${Math.round(confidence*100)}%`;
  if(confidence<.999)return `対応推定 ${Math.round(confidence*100)}%`;
  return '対応確実';
}
function diffCsvCell(value){
  const text=String(value??'').replace(/\r?\n/g,' ').replace(/"/g,'""');
  return `"${text}"`;
}
function buildDiffSummaryCsv(detail,workbookName){
  const cmp=detail?.comparison||{},scope=String(cmp.scope||'')==='history'?'任意2版':'自動比較';
  const headers=['原稿','比較種別','比較元日時','比較先日時','比較元項目','比較先項目','判定','確認状態','対応確信度','対応方法','比較元ページ数','比較先ページ数','詳細'];
  const rows=asArray(detail?.sheets).map(sheet=>{
    const kind=String(sheet?.kind||'unknown'),confidence=(kind==='added'||kind==='removed')?'':`${Math.round(Math.max(0,Math.min(1,Number(sheet?.matchConfidence??1)))*100)}%`;
    return [workbookName,scope,formatDateTime(cmp.baselineAt),formatDateTime(cmp.currentAt),String(sheet?.beforeSheetName||''),String(sheet?.afterSheetName||''),diffKindMeta(kind).label,sheet?.confirmed?'確認済み':'未確認',confidence,String(sheet?.matchMethod||''),Number(sheet?.beforePages||0),Number(sheet?.afterPages||0),String(sheet?.message||'')];
  });
  return [headers,...rows].map(row=>row.map(diffCsvCell).join(',')).join('\r\n');
}
function exportDiffSummary(){
  const detail=diffViewState.detail,sheets=asArray(detail?.sheets);if(!detail||!sheets.length)return;
  const workbookName=String(detail?.workbookName||getWorkbook(diffViewState.workbookId)?.displayName||'source');
  const csv=buildDiffSummaryCsv(detail,workbookName),blob=new Blob(['\uFEFF',csv],{type:'text/csv;charset=utf-8'}),url=URL.createObjectURL(blob),a=document.createElement('a');
  const stamp=new Date().toISOString().replace(/[-:]/g,'').replace('T','-').slice(0,15),stem=workbookName.replace(/[\\/:*?"<>|]/g,'_').slice(0,80)||'source';
  a.href=url;a.download=`ReportBinder-diff-${stem}-${stamp}.csv`;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
  const button=$('diff-export-summary');if(button){button.textContent='CSVを出力しました';setTimeout(()=>{if(button.isConnected)button.textContent='差分レポートCSV';},1800);}
}
function ensureDiffPdfJs(){
  if(diffPdfJsPromise)return diffPdfJsPromise;
  diffPdfJsPromise=(async()=>{
    let lib=null;
    if(String(state?.pdfjsMode||'')==='classic'){
      if(!window.pdfjsLib){
        await new Promise((resolve,reject)=>{
          const script=document.createElement('script');
          script.src=new URL('pdfjs/pdf.min.js',location.href).href;
          script.onload=resolve;
          script.onerror=()=>reject(new Error('PDF.jsを読み込めませんでした。'));
          document.head.appendChild(script);
        });
      }
      lib=window.pdfjsLib;
      if(lib?.GlobalWorkerOptions)lib.GlobalWorkerOptions.workerSrc=new URL('pdfjs/pdf.worker.min.js',location.href).href;
    }else{
      lib=await import(new URL('pdfjs/pdf.min.mjs',location.href).href);
      if(lib?.GlobalWorkerOptions)lib.GlobalWorkerOptions.workerSrc=new URL('pdfjs/pdf.worker.min.mjs',location.href).href;
    }
    if(!lib?.getDocument)throw new Error('PDF.jsを利用できません。外部ライブラリを再配置してください。');
    return lib;
  })().catch(error=>{diffPdfJsPromise=null;throw error;});
  return diffPdfJsPromise;
}
function diffComparisonIdentity(){
  const cmp=diffViewState.detail?.comparison||{};
  return [diffViewState.workbookId,cmp.scope||'automatic',cmp.baselineSnapshotId||'',cmp.baselineVersionId||'',cmp.currentSnapshotId||'',cmp.currentVersionId||''].join('|');
}
function diffSheetPageCount(sheet){
  return Math.max(0,Number(sheet?.pageCount||0),Number(sheet?.beforePages||0),Number(sheet?.afterPages||0),asArray(sheet?.pages).length);
}
function preferredDiffPageIndex(sheet){
  const pageCount=diffSheetPageCount(sheet);
  if(String(sheet?.kind||'')!=='modified'||pageCount<=1)return 0;
  const mappedPages=asArray(sheet?.pages);
  if(mappedPages.length){
    const changed=mappedPages.findIndex(page=>String(page?.comparisonKind||'modified')!=='unchanged');
    if(changed>=0)return changed;
  }
  const unchanged=new Set(asArray(sheet?.unchangedPageNumbers).map(Number).filter(Number.isFinite));
  for(let pageNumber=1;pageNumber<=pageCount;pageNumber++)if(!unchanged.has(pageNumber))return pageNumber-1;
  return 0;
}
function diffBrowserPageKey(sheet,pageIndex){
  return `${diffComparisonIdentity()}|${String(sheet?.sheetKey||'')}|${Number(pageIndex||0)}`;
}
function diffCurrentPage(){
  const sheet=diffCurrentSheet();
  return sheet?(diffBrowserPageCache.get(diffBrowserPageKey(sheet,diffViewState.pageIndex))?.page||null):null;
}
function diffPdfDocumentKey(side,sheet){
  return `${diffComparisonIdentity()}|${side}|${String(sheet?.sheetKey||'')}`;
}
function diffHistoryPdfParams(side,sheet){
  const cmp=diffViewState.detail?.comparison||{};
  const before=side==='before';
  const sideName=before?String(sheet?.beforeSheetName||sheet?.sheetName||''):String(sheet?.afterSheetName||sheet?.sheetName||'');
  return {workbookId:diffViewState.workbookId,snapshotId:String(before?cmp.baselineSnapshotId||'':cmp.currentSnapshotId||''),versionId:String(before?cmp.baselineVersionId||'':cmp.currentVersionId||''),sheetName:sideName};
}
async function fetchDiffPdfDocument(side,sheet){
  const key=diffPdfDocumentKey(side,sheet);
  if(diffPdfDocumentCache.has(key)){
    const cached=diffPdfDocumentCache.get(key);
    diffPdfDocumentCache.delete(key);diffPdfDocumentCache.set(key,cached);
    return cached;
  }
  const promise=(async()=>{
    const lib=await ensureDiffPdfJs();
    const loadingTask=lib.getDocument({
      url:apiUrl('/api/history/content-pdf',diffHistoryPdfParams(side,sheet)),
      httpHeaders:{'X-ReportBinder-Token':token,'Accept':'application/pdf'},
      disableRange:false,
      disableStream:false
    });
    try{return await loadingTask.promise;}
    catch(error){throw new Error('PDFを取得できませんでした。'+String(error?.message||''));}
  })();
  diffPdfDocumentCache.set(key,promise);
  while(diffPdfDocumentCache.size>DIFF_PDF_CACHE_LIMIT){
    const oldest=diffPdfDocumentCache.keys().next().value;
    if(oldest===key)break;
    const oldPromise=diffPdfDocumentCache.get(oldest);
    diffPdfDocumentCache.delete(oldest);
    Promise.resolve(oldPromise).then(doc=>doc?.destroy?.()).catch(()=>{});
  }
  try{return await promise;}catch(error){diffPdfDocumentCache.delete(key);throw error;}
}

function beginDiffBrowserRender(){
  diffBrowserRenderSerial++;
  for(const task of diffActiveRenderTasks){try{task.cancel();}catch{}}
  diffActiveRenderTasks.clear();
  return diffBrowserRenderSerial;
}
function clearDiffCanvas(id){
  const canvas=$(id);if(!canvas)return;
  canvas.style.visibility='hidden';canvas.style.opacity='';
  canvas.width=1;canvas.height=1;
  canvas.getContext('2d')?.clearRect(0,0,1,1);
}
function copyDiffCanvas(id,source){
  const canvas=$(id);if(!canvas||!source)return;
  canvas.width=source.width;canvas.height=source.height;
  const context=canvas.getContext('2d',{alpha:false});
  context.fillStyle='#fff';context.fillRect(0,0,canvas.width,canvas.height);context.drawImage(source,0,0);
  canvas.style.visibility='';
}
function createWhiteDiffCanvas(width,height){
  const canvas=document.createElement('canvas');
  canvas.width=Math.max(1,width);canvas.height=Math.max(1,height);
  const context=canvas.getContext('2d',{alpha:false});
  context.fillStyle='#fff';context.fillRect(0,0,canvas.width,canvas.height);
  return canvas;
}
async function renderDiffPdfPage(side,sheet,pageNumber,serial){
  const configuredPages=Number(side==='before'?sheet?.beforePages||0:sheet?.afterPages||0);
  if(configuredPages>0&&pageNumber>configuredPages)return null;
  const doc=await fetchDiffPdfDocument(side,sheet);
  if(serial!==diffBrowserRenderSerial||pageNumber>doc.numPages)return null;
  const page=await doc.getPage(pageNumber);
  if(serial!==diffBrowserRenderSerial)return null;
  const viewport=page.getViewport({scale:DIFF_RENDER_SCALE});
  const canvas=createWhiteDiffCanvas(Math.ceil(viewport.width),Math.ceil(viewport.height));
  const task=page.render({canvasContext:canvas.getContext('2d',{alpha:false}),viewport,background:'rgb(255,255,255)'});
  diffActiveRenderTasks.add(task);
  try{await task.promise;}finally{diffActiveRenderTasks.delete(task);}
  return serial===diffBrowserRenderSerial?{canvas,page,viewport}:null;
}
function normalizeDiffPdfText(value){
  return String(value||'').normalize('NFKC').replace(/\s+/g,' ').trim();
}
function mapDiffPdfTextContentItems(content,viewport){
  const scale=Math.abs(Number(viewport?.scale||DIFF_RENDER_SCALE))||DIFF_RENDER_SCALE;
  return asArray(content?.items).map(item=>{
    const text=normalizeDiffPdfText(item?.str);if(!text||!Array.isArray(item?.transform))return null;
    const point=viewport.convertToViewportPoint(Number(item.transform[4]||0),Number(item.transform[5]||0));
    const height=Math.max(1,Math.abs(Number(item.height||Math.hypot(Number(item.transform[2]||0),Number(item.transform[3]||0))))*scale);
    return {text,x:Number(point[0]||0),y:Number(point[1]||0)-height,width:Math.max(1,Math.abs(Number(item.width||0))*scale),height};
  }).filter(Boolean);
}
async function extractDiffPdfTextItems(rendered,serial){
  if(!rendered?.page||!rendered?.viewport)return [];
  const content=await rendered.page.getTextContent();
  if(serial!==diffBrowserRenderSerial)return [];
  return mapDiffPdfTextContentItems(content,rendered.viewport);
}
async function extractDiffPdfDocumentTextPages(side,sheet){
  const key=`${diffPdfDocumentKey(side,sheet)}|text-v1`;
  if(diffPdfTextDocumentCache.has(key))return diffPdfTextDocumentCache.get(key);
  const promise=(async()=>{
    const doc=await fetchDiffPdfDocument(side,sheet),pages=[];
    for(let pageNumber=1;pageNumber<=doc.numPages;pageNumber++){
      const page=await doc.getPage(pageNumber),viewport=page.getViewport({scale:DIFF_RENDER_SCALE});
      pages.push(mapDiffPdfTextContentItems(await page.getTextContent(),viewport));
    }
    return pages;
  })();
  diffPdfTextDocumentCache.set(key,promise);
  while(diffPdfTextDocumentCache.size>DIFF_PDF_CACHE_LIMIT)diffPdfTextDocumentCache.delete(diffPdfTextDocumentCache.keys().next().value);
  try{return await promise;}catch(error){diffPdfTextDocumentCache.delete(key);throw error;}
}
function diffPdfNumericFragments(value){
  const text=normalizeDiffPdfText(value),pattern=/[△▲▼+\-−]?\(?\d[\d,]*(?:\.\d+)?(?:[%％])?\)?/g,fragments=[];
  for(const match of text.matchAll(pattern))fragments.push({text:match[0],signature:match[0].replace(/[\s,]/g,'').replace(/％/g,'%'),offset:Number(match.index||0)});
  return fragments;
}
function diffPdfNumericSignature(value){
  return diffPdfNumericFragments(value).map(item=>item.signature).join('|');
}
function isDiffNumericText(value){
  return !!diffPdfNumericSignature(value);
}
function diffPdfTextTemplate(value){
  return normalizeDiffPdfText(value)
    .replace(/[△▲▼+\-−]?\(?\d[\d,]*(?:\.\d+)?(?:[%％])?\)?/g,'#')
    .replace(/\s+/g,' ').trim();
}
function diffPdfTextTemplatesMatch(beforeItems,afterItems){
  const counts=items=>{
    const map=new Map();
    for(const item of items){
      const key=diffPdfTextTemplate(item.text);if(!key)continue;
      map.set(key,(map.get(key)||0)+1);
    }
    return map;
  };
  const before=counts(beforeItems),after=counts(afterItems);
  if(before.size!==after.size)return false;
  for(const [key,count] of before)if(after.get(key)!==count)return false;
  return true;
}
function diffPdfNonNumericFingerprint(items){
  const sorted=[...items].sort((a,b)=>{
    const ay=Number(a.y||0)+Number(a.height||0)/2,by=Number(b.y||0)+Number(b.height||0)/2;
    return Math.abs(ay-by)>4?ay-by:Number(a.x||0)-Number(b.x||0);
  });
  return sorted.map(item=>normalizeDiffPdfText(item.text)
    .replace(/[△▲▼+\-−]?\(?\d[\d,]*(?:\.\d+)?(?:[%％])?\)?/g,'')
    .replace(/\s+/g,'')).filter(Boolean).join('');
}
function diffPdfTextLayoutFingerprint(items){
  return [...items].sort((a,b)=>Number(a.y||0)-Number(b.y||0)||Number(a.x||0)-Number(b.x||0)).map(item=>{
    const q=value=>Math.round(Number(value||0)/2);
    return `${normalizeDiffPdfText(item.text)}@${q(item.x)},${q(item.y)},${q(item.width)},${q(item.height)}`;
  }).join('\n');
}
function diffRegionOverlapsPdfText(region,items,width,height){
  const box=region?.after||region?.before||region,left=Number(box?.x||0)*width,top=Number(box?.y||0)*height;
  const right=left+Number(box?.width||0)*width,bottom=top+Number(box?.height||0)*height;
  return items.some(item=>Math.min(right,Number(item.x||0)+Number(item.width||0))-Math.max(left,Number(item.x||0))>1&&
    Math.min(bottom,Number(item.y||0)+Number(item.height||0))-Math.max(top,Number(item.y||0))>1);
}
function shouldSuppressDiffRasterNoise(beforeItems,afterItems,analysis,regions,width,height){
  if(!beforeItems.length||!afterItems.length||!regions.length||regions.length>24||analysis?.alignmentAdjusted||analysis?.fallbackUsed)return false;
  if(diffPdfTextLayoutFingerprint(beforeItems)!==diffPdfTextLayoutFingerprint(afterItems))return false;
  const changedRatio=Number(analysis?.changedRatio||0);
  if(changedRatio>=.001)return false;
  return regions.every(region=>{
    const box=region?.after||region?.before||region,rw=Math.max(1,Number(box?.width||0)*width),rh=Math.max(1,Number(box?.height||0)*height);
    const area=rw*rh,density=Number(region?.pixelCount||0)/Math.max(1,area),aspect=Math.max(rw/rh,rh/rw);
    if(rw>=width*.15||rh>=height*.1||aspect>=8)return false;
    const overlapsText=diffRegionOverlapsPdfText(region,beforeItems,width,height)||diffRegionOverlapsPdfText(region,afterItems,width,height);
    if(overlapsText)return changedRatio<.00045&&density<.035&&Number(region?.pixelCount||0)<80;
    return changedRatio<.00055&&density<.12;
  });
}
function diffPdfNumericItems(items){
  const fragments=items.flatMap((item,parentIndex)=>diffPdfNumericFragments(item.text).map((fragment,fragmentIndex)=>{
    const text=normalizeDiffPdfText(item.text),total=Math.max(1,[...text].length);
    const start=[...text.slice(0,fragment.offset)].length/total,length=Math.max(1,[...fragment.text].length)/total;
    return {...item,text:fragment.text,x:Number(item.x||0)+Number(item.width||0)*start,width:Math.max(1,Number(item.width||0)*length),parentIndex,fragmentIndex,signature:fragment.signature};
  }));
  // PDF.js may expose the same Excel glyph run twice on nearly identical baselines.
  // Treat those as one number without collapsing legitimate repeats in adjacent rows.
  const unique=[];
  for(const fragment of fragments){
    const duplicate=unique.some(previous=>{
      if(previous.signature!==fragment.signature)return false;
      const dx=Math.abs((previous.x+previous.width/2)-(fragment.x+fragment.width/2));
      const dy=Math.abs((previous.y+previous.height/2)-(fragment.y+fragment.height/2));
      return dx<=Math.max(3,Math.min(previous.width,fragment.width)*.35)&&
        dy<=Math.max(4,Math.max(previous.height,fragment.height)*1.35);
    });
    if(!duplicate)unique.push(fragment);
  }
  return unique;
}
function diffPdfNumericCenterDistance(before,after,width,height){
  const beforeX=before.x+before.width/2,afterX=after.x+after.width/2;
  const beforeY=before.y+before.height/2,afterY=after.y+after.height/2;
  const dx=Math.abs(beforeX-afterX)/Math.max(1,width),dy=Math.abs(beforeY-afterY)/Math.max(1,height);
  return {dx,dy,score:dy*6+dx};
}
function unmatchedDiffPdfNumericItems(beforeItems,afterItems,width,height){
  const before=diffPdfNumericItems(beforeItems),after=diffPdfNumericItems(afterItems);
  const beforeUsed=new Uint8Array(before.length),afterUsed=new Uint8Array(after.length),bySignature=new Map();
  for(let index=0;index<after.length;index++){
    const bucket=bySignature.get(after[index].signature)||[];bucket.push(index);bySignature.set(after[index].signature,bucket);
  }
  const candidates=[];
  for(let beforeIndex=0;beforeIndex<before.length;beforeIndex++)for(const afterIndex of bySignature.get(before[beforeIndex].signature)||[]){
    candidates.push({beforeIndex,afterIndex,...diffPdfNumericCenterDistance(before[beforeIndex],after[afterIndex],width,height)});
  }
  candidates.sort((a,b)=>a.score-b.score);
  for(const candidate of candidates){
    if(beforeUsed[candidate.beforeIndex]||afterUsed[candidate.afterIndex])continue;
    beforeUsed[candidate.beforeIndex]=1;afterUsed[candidate.afterIndex]=1;
  }
  return {
    before:before.filter((item,index)=>!beforeUsed[index]),
    after:after.filter((item,index)=>!afterUsed[index])
  };
}
function pairChangedDiffPdfNumbers(beforeItems,afterItems,width,height){
  const unmatched=unmatchedDiffPdfNumericItems(beforeItems,afterItems,width,height),candidates=[];
  for(let beforeIndex=0;beforeIndex<unmatched.before.length;beforeIndex++)for(let afterIndex=0;afterIndex<unmatched.after.length;afterIndex++){
    const distance=diffPdfNumericCenterDistance(unmatched.before[beforeIndex],unmatched.after[afterIndex],width,height);
    const rowTolerance=Math.max(.012,(unmatched.before[beforeIndex].height+unmatched.after[afterIndex].height)*1.5/Math.max(1,height));
    const columnTolerance=Math.max(.04,(unmatched.before[beforeIndex].width+unmatched.after[afterIndex].width)*2/Math.max(1,width));
    if(distance.dy<=rowTolerance&&distance.dx<=Math.min(.12,columnTolerance))candidates.push({beforeIndex,afterIndex,...distance});
  }
  candidates.sort((a,b)=>a.score-b.score);
  const beforeUsed=new Uint8Array(unmatched.before.length),afterUsed=new Uint8Array(unmatched.after.length),pairs=[];
  for(const candidate of candidates){
    if(beforeUsed[candidate.beforeIndex]||afterUsed[candidate.afterIndex])continue;
    beforeUsed[candidate.beforeIndex]=1;afterUsed[candidate.afterIndex]=1;
    pairs.push({before:unmatched.before[candidate.beforeIndex],after:unmatched.after[candidate.afterIndex]});
  }
  return {pairs,unmatchedBefore:unmatched.before.length,unmatchedAfter:unmatched.after.length};
}
function diffPdfTextPixelBox(items,padding,width,height){
  if(!items.length)return null;
  const minX=Math.max(0,Math.min(...items.map(item=>item.x))-padding),minY=Math.max(0,Math.min(...items.map(item=>item.y))-padding);
  const maxX=Math.min(width,Math.max(...items.map(item=>item.x+item.width))+padding),maxY=Math.min(height,Math.max(...items.map(item=>item.y+item.height))+padding);
  return {x:minX/width,y:minY/height,width:Math.max(1,maxX-minX)/width,height:Math.max(1,maxY-minY)/height};
}
function diffPdfTextBoxUnion(boxes){
  const valid=boxes.filter(Boolean);if(!valid.length)return null;
  const minX=Math.min(...valid.map(box=>Number(box.x||0))),minY=Math.min(...valid.map(box=>Number(box.y||0)));
  const maxX=Math.max(...valid.map(box=>Number(box.x||0)+Number(box.width||0)));
  const maxY=Math.max(...valid.map(box=>Number(box.y||0)+Number(box.height||0)));
  return {x:minX,y:minY,width:Math.max(.000001,maxX-minX),height:Math.max(.000001,maxY-minY)};
}
function mergeAdjacentDiffTextRegions(regions){
  const merged=[];
  for(const region of [...regions].sort((a,b)=>Number(a.y||0)-Number(b.y||0)||Number(a.x||0)-Number(b.x||0))){
    const previous=merged[merged.length-1],changedBox=value=>value?.kind==='removed'?(value?.before||value?.after||value):(value?.after||value?.before||value);
    const a=changedBox(previous),b=changedBox(region);
    if(previous&&previous.kind===region.kind){
      const gap=Number(b.y||0)-(Number(a.y||0)+Number(a.height||0));
      const overlap=Math.min(Number(a.x||0)+Number(a.width||0),Number(b.x||0)+Number(b.width||0))-Math.max(Number(a.x||0),Number(b.x||0));
      const horizontalRatio=overlap/Math.max(.000001,Math.min(Number(a.width||0),Number(b.width||0)));
      const near=gap>=-.01&&gap<=Math.max(.018,Number(a.height||0)*1.25,Number(b.height||0)*1.25);
      if(near&&(horizontalRatio>=.25||Math.abs(Number(a.x||0)-Number(b.x||0))<=.035)){
        const before=diffPdfTextBoxUnion([previous.before,region.before]),after=diffPdfTextBoxUnion([previous.after,region.after]);
        const combined=diffPdfTextBoxUnion([previous,region,before,after]);
        Object.assign(previous,combined,{before:before||after||combined,after:after||before||combined});
        continue;
      }
    }
    merged.push({...region,before:region.before?{...region.before}:region.before,after:region.after?{...region.after}:region.after});
  }
  return merged;
}
// 同一テキストどうしだけが重複候補になるので、正規化文字列でバケットに分けてから
// 近接判定する。総当たりだと巻全体の比較(数万項目)で数分単位のフリーズになる。
// 正規化(NFKC)も項目ごとに1回だけ行う。
function dedupeDiffPdfRowItems(items){
  const unique=[],buckets=new Map();
  const prepared=items.map(item=>({
    item,
    text:normalizeDiffPdfText(item.text),
    centerX:Number(item.x||0)+Number(item.width||0)/2,
    centerY:Number(item.y||0)+Number(item.height||0)/2,
    width:Number(item.width||0),
    height:Number(item.height||0)
  }));
  prepared.sort((a,b)=>Number(a.item.y||0)-Number(b.item.y||0)||Number(a.item.x||0)-Number(b.item.x||0));
  for(const entry of prepared){
    let bucket=buckets.get(entry.text);
    if(!bucket){bucket=[];buckets.set(entry.text,bucket);}
    const duplicate=bucket.some(previous=>{
      const height=Math.max(previous.height,entry.height);
      return Math.abs(entry.centerX-previous.centerX)<=Math.max(3,Math.min(previous.width,entry.width)*.12)&&
        Math.abs(entry.centerY-previous.centerY)<=Math.max(4,height*.9);
    });
    if(!duplicate){bucket.push(entry);unique.push(entry.item);}
  }
  return unique;
}
function groupDiffPdfTextRows(items){
  const rows=[];
  for(const item of dedupeDiffPdfRowItems(items).sort((a,b)=>(Number(a.y||0)+Number(a.height||0)/2)-(Number(b.y||0)+Number(b.height||0)/2)||Number(a.x||0)-Number(b.x||0))){
    const center=Number(item.y||0)+Number(item.height||0)/2,last=rows[rows.length-1],tolerance=Math.max(3,Math.min(12,Number(item.height||0)*.7));
    if(!last||Math.abs(center-last.center)>Math.max(tolerance,last.tolerance)){rows.push({items:[item],center,tolerance});continue;}
    last.items.push(item);last.center=(last.center*(last.items.length-1)+center)/last.items.length;last.tolerance=Math.max(last.tolerance,tolerance);
  }
  return rows.map(row=>{
    row.items.sort((a,b)=>Number(a.x||0)-Number(b.x||0));
    row.signature=diffPdfTextTemplate(row.items.map(item=>item.text).join(' ')).replace(/\s+/g,'');
    return row;
  }).filter(row=>row.signature);
}
function matchDiffPdfTextRows(beforeRows,afterRows){
  const n=beforeRows.length,m=afterRows.length,stride=m+1,dp=new Uint16Array((n+1)*(m+1));
  for(let i=n-1;i>=0;i--)for(let j=m-1;j>=0;j--)dp[i*stride+j]=beforeRows[i].signature===afterRows[j].signature?dp[(i+1)*stride+j+1]+1:Math.max(dp[(i+1)*stride+j],dp[i*stride+j+1]);
  const matches=[];let i=0,j=0;
  while(i<n&&j<m){
    if(beforeRows[i].signature===afterRows[j].signature){matches.push([i,j]);i++;j++;}
    else if(dp[(i+1)*stride+j]>=dp[i*stride+j+1])i++;else j++;
  }
  return matches;
}
function buildTextRowStructureDiffResult(beforeItems,afterItems,width,height){
  const empty={regions:[],confident:false};
  // 兄弟の判定器と同じく規模で打ち切る。ここは巻全体を平坦化した配列も受け取るため、
  // 上限が無いと行グループ化と行マッチングが実用時間を超える。
  if(beforeItems.length>DIFF_TEXT_ROW_ITEM_LIMIT||afterItems.length>DIFF_TEXT_ROW_ITEM_LIMIT)return empty;
  const beforeRows=groupDiffPdfTextRows(beforeItems),afterRows=groupDiffPdfTextRows(afterItems);
  const delta=afterRows.length-beforeRows.length;
  if(!delta||Math.abs(delta)>6||beforeRows.length<3||afterRows.length<3)return empty;
  const matches=matchDiffPdfTextRows(beforeRows,afterRows),beforeMatched=new Set(matches.map(pair=>pair[0])),afterMatched=new Set(matches.map(pair=>pair[1]));
  let beforeMissing=beforeRows.map((row,index)=>({row,index})).filter(item=>!beforeMatched.has(item.index));
  let afterMissing=afterRows.map((row,index)=>({row,index})).filter(item=>!afterMatched.has(item.index));
  // Repagination can move a heading across otherwise matched rows. Cancel exact
  // unmatched signatures on both sides before deciding what was truly added.
  const movedAfter=new Set(),movedBefore=new Set();
  for(let beforeIndex=0;beforeIndex<beforeMissing.length;beforeIndex++){
    const afterIndex=afterMissing.findIndex((entry,index)=>!movedAfter.has(index)&&entry.row.signature===beforeMissing[beforeIndex].row.signature);
    if(afterIndex>=0){movedBefore.add(beforeIndex);movedAfter.add(afterIndex);}
  }
  beforeMissing=beforeMissing.filter((entry,index)=>!movedBefore.has(index));
  afterMissing=afterMissing.filter((entry,index)=>!movedAfter.has(index));
  if(delta>0&&(beforeMissing.length||afterMissing.length!==delta)||delta<0&&(afterMissing.length||beforeMissing.length!==-delta))return empty;
  const missing=delta>0?afterMissing:beforeMissing,kind=delta>0?'added':'removed',regions=[];
  for(const entry of missing){
    const changedBox=diffPdfTextPixelBox(entry.row.items,5,width,height);if(!changedBox)continue;
    const otherRows=delta>0?beforeRows:afterRows,insertionIndex=matches.filter(pair=>(delta>0?pair[1]:pair[0])<entry.index).length;
    const anchorRow=otherRows[Math.min(insertionIndex,otherRows.length-1)]||otherRows[otherRows.length-1],anchorBox=anchorRow?diffPdfTextPixelBox(anchorRow.items,5,width,height):changedBox;
    const beforeBox=delta>0?{...anchorBox,x:changedBox.x,width:changedBox.width}:changedBox;
    const afterBox=delta>0?changedBox:{...anchorBox,x:changedBox.x,width:changedBox.width};
    regions.push({regionId:'',kind,...changedBox,before:beforeBox,after:afterBox,confidence:1,pixelCount:0,source:'pdf-row'});
  }
  const confident=regions.length===Math.abs(delta);
  return {regions:confident?mergeAdjacentDiffTextRegions(regions):regions,confident};
}
function findDiffPdfTextColumnSplit(beforeItems,afterItems,width){
  const usable=[...beforeItems,...afterItems].filter(item=>Number(item.width||0)<width*.58);
  if(usable.length<12)return null;
  const centers=usable.map(item=>Number(item.x||0)+Number(item.width||0)/2).sort((a,b)=>a-b),candidates=[];
  for(let index=1;index<centers.length;index++){
    const gap=centers[index]-centers[index-1],split=(centers[index]+centers[index-1])/2;
    if(gap<width*.12||split<width*.28||split>width*.72)continue;
    const sideCounts=items=>items.reduce((counts,item)=>{(Number(item.x||0)+Number(item.width||0)/2)<split?counts[0]++:counts[1]++;return counts;},[0,0]);
    const beforeCounts=sideCounts(beforeItems),afterCounts=sideCounts(afterItems);
    if(Math.min(...beforeCounts,...afterCounts)<3)continue;
    const crossing=usable.filter(item=>Number(item.x||0)<split&&Number(item.x||0)+Number(item.width||0)>split).length;
    if(crossing>Math.max(2,usable.length*.08))continue;
    candidates.push({split,score:gap-crossing*width*.04});
  }
  return candidates.sort((a,b)=>b.score-a.score)[0]?.split||null;
}
function buildColumnTextRowStructureDiffResult(beforeItems,afterItems,width,height){
  const empty={regions:[],confident:false},split=findDiffPdfTextColumnSplit(beforeItems,afterItems,width);if(!split)return empty;
  const side=(items,left)=>items.filter(item=>((Number(item.x||0)+Number(item.width||0)/2)<split)===left),results=[];
  for(const left of [true,false]){
    const result=buildTextRowStructureDiffResult(side(beforeItems,left),side(afterItems,left),width,height);
    if(result.confident)results.push(result);
  }
  return results.length===1?results[0]:empty;
}
function buildDocumentTextRowStructureDiffResult(beforePages,afterPages,width,height,beforePageNumber,afterPageNumber=beforePageNumber){
  const empty={regions:[],confident:false,documentChanged:false};
  if(!Array.isArray(beforePages)||!Array.isArray(afterPages)||Math.max(beforePages.length,afterPages.length)<2)return empty;
  const stride=height+20,pageCount=Math.max(beforePages.length,afterPages.length),totalHeight=stride*pageCount;
  // Repeated running headers/footers interrupt the reading order when text that
  // merely repaginates crosses a page boundary. Exclude only the outer 5% from
  // the document-wide LCS; page-local comparison still checks that furniture.
  const flatten=pages=>pages.flatMap((items,index)=>items.filter(item=>{
    const center=Number(item.y||0)+Number(item.height||0)/2;return center>=height*.05&&center<=height*.95;
  }).map(item=>({...item,y:Number(item.y||0)+index*stride})));
  const result=buildTextRowStructureDiffResult(flatten(beforePages),flatten(afterPages),width,totalHeight);
  if(!result.confident)return empty;
  const localBox=(box,pageNumber)=>{
    if(!box)return null;
    const target=Math.max(0,Number(pageNumber||1)-1);
    const absoluteY=Number(box.y||0)*totalHeight,index=Math.max(0,Math.min(pageCount-1,Math.floor(absoluteY/stride)));
    if(index!==target)return null;
    return {...box,y:Math.max(0,(absoluteY-index*stride)/height),height:Math.min(1,Number(box.height||0)*totalHeight/height)};
  };
  const regions=[];
  for(const region of result.regions){
    let before=localBox(region.before,beforePageNumber),after=localBox(region.after,afterPageNumber);
    // Row diff gives an added/removed row a synthetic box on the opposite side
    // for overlay display. That box has no document-page identity, so decide
    // membership from the side on which the row actually exists.
    if(region.kind==='added'&&!after)continue;
    if(region.kind==='removed'&&!before)continue;
    if(!before&&!after)continue;
    if(region.kind==='added'&&!before)before=after;
    if(region.kind==='removed'&&!after)after=before;
    const combined=diffPdfTextBoxUnion([before,after]);
    regions.push({...region,...combined,before:before||after||combined,after:after||before||combined,source:'pdf-document-row'});
  }
  return {regions:mergeAdjacentDiffTextRegions(regions),confident:true,documentChanged:result.regions.length>0};
}
function buildTextLayoutShiftDiffResult(beforeItems,afterItems,width,height){
  const empty={regions:[],confident:false},bodyRows=items=>groupDiffPdfTextRows(items).filter(row=>row.center>=height*.05&&row.center<=height*.95);
  const beforeRows=bodyRows(beforeItems),afterRows=bodyRows(afterItems);
  if(beforeRows.length<4||beforeRows.length!==afterRows.length||beforeRows.length>500)return empty;
  if(beforeRows.some((row,index)=>row.signature!==afterRows[index].signature))return empty;
  const offsets=beforeRows.map((row,index)=>afterRows[index].center-row.center);
  const rowHeight=Math.max(1,...beforeRows.flatMap(row=>row.items.map(item=>Number(item.height||0))));
  const threshold=Math.max(1.5,rowHeight*.14),jumps=[];
  for(let index=1;index<offsets.length;index++)if(Math.abs(offsets[index]-offsets[index-1])>=threshold)jumps.push(index);
  if(!jumps.length||jumps.length>6)return empty;
  const start=Math.max(0,jumps[0]-1),end=Math.min(beforeRows.length-1,jumps[jumps.length-1]);
  const beforeBox=diffPdfTextPixelBox(beforeRows.slice(start,end+1).flatMap(row=>row.items),5,width,height);
  const afterBox=diffPdfTextPixelBox(afterRows.slice(start,end+1).flatMap(row=>row.items),5,width,height);
  const combined=diffPdfTextBoxUnion([beforeBox,afterBox]);if(!combined)return empty;
  return {regions:[{regionId:'',kind:'modified',...combined,before:beforeBox||combined,after:afterBox||combined,confidence:1,pixelCount:0,source:'pdf-layout'}],confident:true};
}
function diffPdfTextItemsInBand(items,band,width,height){
  if(!band)return [];
  const left=Number(band.x||0)*width,top=Number(band.y||0)*height;
  const right=left+Number(band.width||0)*width,bottom=top+Number(band.height||0)*height;
  return items.filter(item=>{
    const centerX=Number(item.x||0)+Number(item.width||0)/2,centerY=Number(item.y||0)+Number(item.height||0)/2;
    return centerX>=left&&centerX<=right&&centerY>=top&&centerY<=bottom;
  });
}
function buildLocalizedTextRowStructureDiffResult(beforeItems,afterItems,width,height,analysis){
  if(!analysis?.tableRowStructureDetected||!analysis?.tableRowStructureBand)return {regions:[],confident:false};
  return buildTextRowStructureDiffResult(
    diffPdfTextItemsInBand(beforeItems,analysis.tableRowStructureBand,width,height),
    diffPdfTextItemsInBand(afterItems,analysis.tableRowStructureBand,width,height),width,height);
}
function buildNumericTextDiffResult(beforeItems,afterItems,width,height){
  const empty={regions:[],changedGroupCount:0,numericGroupCount:0,numericOnly:false};
  if(beforeItems.length>1500||afterItems.length>1500)return empty;
  const matched=pairChangedDiffPdfNumbers(beforeItems,afterItems,width,height);
  // A small semantic edit should yield only a few spatial pairs. Reject a noisy
  // extraction-order/layout cascade instead of flooding the page with highlights.
  if(!matched.pairs.length||matched.pairs.length>8)return {...empty,changedGroupCount:matched.unmatchedBefore+matched.unmatchedAfter};
  const regions=[];
  for(const pair of matched.pairs){
    const beforeBox=diffPdfTextPixelBox([pair.before],4,width,height),afterBox=diffPdfTextPixelBox([pair.after],4,width,height);
    const combined=diffPdfTextPixelBox([pair.before,pair.after],4,width,height);
    if(!combined)continue;
    regions.push({regionId:'',kind:'modified',...combined,before:beforeBox||afterBox||combined,after:afterBox||beforeBox||combined,confidence:1,pixelCount:0,source:'pdf-text'});
  }
  const sameNonNumericContent=diffPdfTextTemplatesMatch(beforeItems,afterItems)||
    diffPdfNonNumericFingerprint(beforeItems)===diffPdfNonNumericFingerprint(afterItems);
  const numericOnly=regions.length>0&&matched.unmatchedBefore===regions.length&&matched.unmatchedAfter===regions.length&&sameNonNumericContent;
  return {regions,changedGroupCount:matched.unmatchedBefore+matched.unmatchedAfter,numericGroupCount:regions.length,numericOnly};
}
function buildNumericTextDiffRegions(beforeItems,afterItems,width,height){
  return buildNumericTextDiffResult(beforeItems,afterItems,width,height).regions;
}
function diffPdfTextItemDistance(before,after,width,height){
  const beforeX=Number(before.x||0)+Number(before.width||0)/2,afterX=Number(after.x||0)+Number(after.width||0)/2;
  const beforeY=Number(before.y||0)+Number(before.height||0)/2,afterY=Number(after.y||0)+Number(after.height||0)/2;
  const dx=Math.abs(beforeX-afterX)/Math.max(1,width),dy=Math.abs(beforeY-afterY)/Math.max(1,height);
  return {dx,dy,score:dy*6+dx};
}
function buildTextFragmentDiffResult(beforeItems,afterItems,width,height){
  const empty={regions:[],confident:false};
  if(beforeItems.length>1500||afterItems.length>1500)return empty;
  const clean=items=>dedupeDiffPdfRowItems(items).map(item=>({...item,signature:normalizeDiffPdfText(item.text)})).filter(item=>item.signature);
  const before=clean(beforeItems),after=clean(afterItems),beforeUsed=new Uint8Array(before.length),afterUsed=new Uint8Array(after.length),exact=[];
  for(let beforeIndex=0;beforeIndex<before.length;beforeIndex++)for(let afterIndex=0;afterIndex<after.length;afterIndex++){
    if(before[beforeIndex].signature!==after[afterIndex].signature)continue;
    exact.push({beforeIndex,afterIndex,...diffPdfTextItemDistance(before[beforeIndex],after[afterIndex],width,height)});
  }
  exact.sort((a,b)=>a.score-b.score);
  for(const candidate of exact){
    if(beforeUsed[candidate.beforeIndex]||afterUsed[candidate.afterIndex])continue;
    beforeUsed[candidate.beforeIndex]=1;afterUsed[candidate.afterIndex]=1;
  }
  const unmatchedBefore=before.map((item,index)=>({item,index})).filter(entry=>!beforeUsed[entry.index]);
  const unmatchedAfter=after.map((item,index)=>({item,index})).filter(entry=>!afterUsed[entry.index]);
  if(!unmatchedBefore.length&&!unmatchedAfter.length)return empty;
  if(unmatchedBefore.length+unmatchedAfter.length>24)return empty;
  const changedCandidates=[];
  for(let beforeIndex=0;beforeIndex<unmatchedBefore.length;beforeIndex++)for(let afterIndex=0;afterIndex<unmatchedAfter.length;afterIndex++){
    const a=unmatchedBefore[beforeIndex].item,b=unmatchedAfter[afterIndex].item,distance=diffPdfTextItemDistance(a,b,width,height);
    const rowTolerance=Math.max(.014,(Number(a.height||0)+Number(b.height||0))*1.6/Math.max(1,height));
    const columnTolerance=Math.max(.035,(Number(a.width||0)+Number(b.width||0))*1.5/Math.max(1,width));
    if(distance.dy<=rowTolerance&&distance.dx<=Math.min(.16,columnTolerance))changedCandidates.push({beforeIndex,afterIndex,...distance});
  }
  changedCandidates.sort((a,b)=>a.score-b.score);
  const pairedBefore=new Uint8Array(unmatchedBefore.length),pairedAfter=new Uint8Array(unmatchedAfter.length),regions=[];
  for(const candidate of changedCandidates){
    if(pairedBefore[candidate.beforeIndex]||pairedAfter[candidate.afterIndex])continue;
    pairedBefore[candidate.beforeIndex]=1;pairedAfter[candidate.afterIndex]=1;
    const a=unmatchedBefore[candidate.beforeIndex].item,b=unmatchedAfter[candidate.afterIndex].item;
    const beforeBox=diffPdfTextPixelBox([a],4,width,height),afterBox=diffPdfTextPixelBox([b],4,width,height),combined=diffPdfTextBoxUnion([beforeBox,afterBox]);
    if(combined)regions.push({regionId:'',kind:'modified',...combined,before:beforeBox||combined,after:afterBox||combined,confidence:1,pixelCount:0,source:'pdf-fragment'});
  }
  for(let index=0;index<unmatchedBefore.length;index++)if(!pairedBefore[index]){
    const box=diffPdfTextPixelBox([unmatchedBefore[index].item],4,width,height);if(box)regions.push({regionId:'',kind:'removed',...box,before:box,after:box,confidence:1,pixelCount:0,source:'pdf-fragment'});
  }
  for(let index=0;index<unmatchedAfter.length;index++)if(!pairedAfter[index]){
    const box=diffPdfTextPixelBox([unmatchedAfter[index].item],4,width,height);if(box)regions.push({regionId:'',kind:'added',...box,before:box,after:box,confidence:1,pixelCount:0,source:'pdf-fragment'});
  }
  return {regions,confident:regions.length>0&&regions.length<=12};
}
function diffRegionsOverlap(first,second){
  const a=first?.after||first?.before||first,b=second?.after||second?.before||second;
  const left=Math.max(Number(a?.x||0),Number(b?.x||0)),top=Math.max(Number(a?.y||0),Number(b?.y||0));
  const right=Math.min(Number(a?.x||0)+Number(a?.width||0),Number(b?.x||0)+Number(b?.width||0));
  const bottom=Math.min(Number(a?.y||0)+Number(a?.height||0),Number(b?.y||0)+Number(b?.height||0));
  if(right<=left||bottom<=top)return false;
  const intersection=(right-left)*(bottom-top),smaller=Math.max(.0000001,Math.min(Number(a?.width||0)*Number(a?.height||0),Number(b?.width||0)*Number(b?.height||0)));
  return intersection/smaller>=.2;
}
function diffRegionsShareTextRow(imageRegion,textRegion){
  for(const side of ['before','after']){
    const image=imageRegion?.[side]||imageRegion,text=textRegion?.[side]||textRegion;
    const imageTop=Number(image?.y||0),imageHeight=Number(image?.height||0),textTop=Number(text?.y||0),textHeight=Number(text?.height||0);
    const overlap=Math.min(imageTop+imageHeight,textTop+textHeight)-Math.max(imageTop,textTop);
    const verticalRatio=overlap/Math.max(.0000001,Math.min(imageHeight,textHeight));
    const rowLike=Number(image?.width||0)>=Math.max(.18,Number(text?.width||0)*2.5)&&imageHeight<=Math.max(.12,textHeight*4);
    if(verticalRatio>=.55&&rowLike)return true;
  }
  return false;
}
function mergeDiffRegionsWithText(imageRegions,textRegions){
  const merged=[...textRegions];
  for(const region of imageRegions)if(!textRegions.some(textRegion=>diffRegionsOverlap(region,textRegion)||diffRegionsShareTextRow(region,textRegion)))merged.push(region);
  merged.sort((a,b)=>Number(a?.y||0)-Number(b?.y||0)||Number(a?.x||0)-Number(b?.x||0));
  return merged.map((region,index)=>({...region,regionId:`browser-r${String(index+1).padStart(4,'0')}`}));
}
function diffHasLocalizedRowSignal(analysis,imageRegions,rowRegions){
  if(analysis?.rowStructureAdjusted||analysis?.tableRowStructureDetected)return true;
  if(analysis?.fallbackUsed||!imageRegions.length||!rowRegions.length)return false;
  return rowRegions.every(textRegion=>imageRegions.some(imageRegion=>{
    for(const side of ['before','after']){
      const image=imageRegion?.[side]||imageRegion,text=textRegion?.[side]||textRegion;
      const imageTop=Number(image?.y||0),imageHeight=Number(image?.height||0),textTop=Number(text?.y||0),textHeight=Number(text?.height||0);
      const overlap=Math.min(imageTop+imageHeight,textTop+textHeight)-Math.max(imageTop,textTop);
      const imageWidth=Number(image?.width||0),textWidth=Number(text?.width||0);
      if(overlap/Math.max(.0000001,Math.min(imageHeight,textHeight))>=.45&&
        imageWidth>=Math.max(.12,textWidth*.65)&&imageHeight<=Math.max(.1,textHeight*4.5))return true;
    }
    return false;
  }));
}
function selectDiffSemanticResult(beforeItems,afterItems,width,height,analysis,imageRegions,documentRowDiff=null){
  const textDiff=buildNumericTextDiffResult(beforeItems,afterItems,width,height),textRegions=textDiff.regions;
  // A small numeric-only edit is stronger evidence than PDF.js row grouping. Excel
  // PDFs sometimes duplicate a glyph run on a nearby baseline and fake an extra row.
  if(textDiff.numericOnly&&textRegions.length>0&&textRegions.length<=8){
    return {mode:'numeric',regions:mergeDiffRegionsWithText([],textRegions),message:'PDF内の文字情報を照合し、数値が変わった箇所だけを強調しました。'};
  }
  if(documentRowDiff?.confident&&documentRowDiff.documentChanged){
    if(documentRowDiff.regions.length)return {mode:'row',regions:mergeDiffRegionsWithText([],documentRowDiff.regions),message:'PDF文書全体の文字行を照合し、追加・削除された段落だけを強調しました。'};
    return {mode:'document-reflow',regions:[],suppressRaster:true,message:'変更箇所は別ページにあり、このページは改ページによる移動だけのため強調を省略しました。'};
  }
  const pageRowDiff=buildTextRowStructureDiffResult(beforeItems,afterItems,width,height);
  const tableRowDiff=pageRowDiff.confident?pageRowDiff:buildLocalizedTextRowStructureDiffResult(beforeItems,afterItems,width,height,analysis);
  const rowDiff=tableRowDiff.confident?tableRowDiff:buildColumnTextRowStructureDiffResult(beforeItems,afterItems,width,height);
  // Text-row LCS alone is not enough. Require an independent raster row-shift signal
  // so harmless PDF text fragmentation cannot replace a valid numeric result.
  if(rowDiff.confident&&diffHasLocalizedRowSignal(analysis,imageRegions,rowDiff.regions)){
    return {mode:'row',regions:mergeDiffRegionsWithText([],rowDiff.regions),message:'PDF内の文字行を照合し、追加・削除された行だけを強調しました。'};
  }
  const layoutDiff=buildTextLayoutShiftDiffResult(beforeItems,afterItems,width,height);
  if(layoutDiff.confident&&diffHasLocalizedRowSignal(analysis,imageRegions,layoutDiff.regions)){
    return {mode:'layout',regions:mergeDiffRegionsWithText([],layoutDiff.regions),message:'PDF内の文字配置を照合し、行間や境界が変わった箇所だけを強調しました。'};
  }
  const fragmentDiff=buildTextFragmentDiffResult(beforeItems,afterItems,width,height);
  if(fragmentDiff.confident){
    // Preserve high-confidence structural bands, but replace low-confidence raster
    // fragments with exact PDF text boxes. This keeps cumulative cell edits visible
    // without exposing every antialiasing fragment produced by a simultaneous resize.
    const structuralRegions=imageRegions.filter(region=>region?.before&&region?.after||Number(region?.confidence||0)>=.9);
    return {mode:'text-fragment',regions:mergeDiffRegionsWithText(structuralRegions,fragmentDiff.regions),message:'PDF内の文字情報を照合し、各世代で追加・削除・変更された箇所を補足しました。'};
  }
  if(textRegions.length){
    return {mode:'numeric-supplement',regions:mergeDiffRegionsWithText(imageRegions,textRegions),message:'PDF内の文字情報を照合し、数値が変わった箇所を補足しました。'};
  }
  return {mode:'image',regions:imageRegions,message:''};
}
function getDiffAnalysisWorker(){
  if(diffAnalysisWorker)return diffAnalysisWorker;
  diffAnalysisWorker=new Worker(new URL('diff-worker.js?v=20260808_v23',location.href));
  diffAnalysisWorker.onmessage=event=>{
    const payload=event.data||{},pending=diffAnalysisPending.get(payload.id);
    if(!pending)return;
    diffAnalysisPending.delete(payload.id);
    if(payload.error)pending.reject(new Error(payload.error));else pending.resolve(payload);
  };
  diffAnalysisWorker.onerror=event=>{
    const error=new Error(event.message||'ブラウザ内の差分解析に失敗しました。');
    for(const pending of diffAnalysisPending.values())pending.reject(error);
    diffAnalysisPending.clear();
    try{diffAnalysisWorker.terminate();}catch{}
    diffAnalysisWorker=null;
  };
  return diffAnalysisWorker;
}
function analyzeDiffCanvases(beforeCanvas,afterCanvas,width,height){
  const before=beforeCanvas.getContext('2d',{willReadFrequently:true}).getImageData(0,0,width,height);
  const after=afterCanvas.getContext('2d',{willReadFrequently:true}).getImageData(0,0,width,height);
  const id=++diffAnalysisRequestSerial;
  return new Promise((resolve,reject)=>{
    diffAnalysisPending.set(id,{resolve,reject});
    try{getDiffAnalysisWorker().postMessage({id,width,height,before:before.data.buffer,after:after.data.buffer},[before.data.buffer,after.data.buffer]);}
    catch(error){diffAnalysisPending.delete(id);reject(error);}
  });
}
function fullDiffRegion(kind,width,height){
  const inset=Math.max(6,Math.floor(Math.min(width,height)/250));
  return {regionId:'browser-r0001',kind,x:inset/width,y:inset/height,width:Math.max(1,width-inset*2)/width,height:Math.max(1,height-inset*2)/height,confidence:1,pixelCount:width*height};
}
function cacheDiffBrowserPage(key,result){
  if(diffBrowserPageCache.has(key))diffBrowserPageCache.delete(key);
  diffBrowserPageCache.set(key,result);
  while(diffBrowserPageCache.size>DIFF_PAGE_CACHE_LIMIT)diffBrowserPageCache.delete(diffBrowserPageCache.keys().next().value);
}
function setDiffBrowserProgress(active,message='',percent=0){
  const box=$('diff-progress');
  box?.classList.toggle('hidden',!active);
  if($('diff-progress-text'))$('diff-progress-text').textContent=message||'表示ページを比較しています。';
  if($('diff-progress-fill'))$('diff-progress-fill').style.width=`${Math.max(0,Math.min(100,Number(percent)||0))}%`;
}
function clearDiffBrowserResources(dropDocuments=false){
  beginDiffBrowserRender();
  for(const pending of diffAnalysisPending.values())pending.reject(new Error('比較画面を切り替えました。'));
  diffAnalysisPending.clear();
  if(dropDocuments){
    if(diffAnalysisWorker){try{diffAnalysisWorker.terminate();}catch{}diffAnalysisWorker=null;}
    for(const promise of diffPdfDocumentCache.values())Promise.resolve(promise).then(doc=>doc?.destroy?.()).catch(()=>{});
    diffPdfDocumentCache.clear();
    diffPdfTextDocumentCache.clear();
  }
  diffBrowserPageCache.clear();setDiffBrowserProgress(false);
  for(const id of ['diff-before-base','diff-after-underlay','diff-after-base'])clearDiffCanvas(id);
}
async function buildDiffBrowserPage(sheet,pageIndex,serial){
  const pageNumber=pageIndex+1,mappedPage=asArray(sheet?.pages)[pageIndex]||null,kind=String(mappedPage?.comparisonKind||sheet?.kind||'modified');
  const beforePageNumber=mappedPage&&Object.hasOwn(mappedPage,'beforePageNumber')?Number(mappedPage.beforePageNumber||0):pageNumber;
  const afterPageNumber=mappedPage&&Object.hasOwn(mappedPage,'afterPageNumber')?Number(mappedPage.afterPageNumber||0):pageNumber;
  const needBefore=kind!=='added'&&beforePageNumber>0&&beforePageNumber<=Math.max(Number(sheet?.beforePages||0),1);
  const needAfter=kind!=='removed'&&afterPageNumber>0&&afterPageNumber<=Math.max(Number(sheet?.afterPages||0),1);
  setDiffBrowserProgress(true,'PDFを表示用に描画しています。',25);
  const [beforeRaw,afterRaw]=await Promise.all([needBefore?renderDiffPdfPage('before',sheet,beforePageNumber,serial):Promise.resolve(null),needAfter?renderDiffPdfPage('after',sheet,afterPageNumber,serial):Promise.resolve(null)]);
  if(serial!==diffBrowserRenderSerial)return null;
  const width=Math.max(beforeRaw?.canvas?.width||0,afterRaw?.canvas?.width||0),height=Math.max(beforeRaw?.canvas?.height||0,afterRaw?.canvas?.height||0);
  if(!width||!height)throw new Error('表示できるPDFページがありません。');
  const normalizeCanvas=rendered=>{
    const raw=rendered?.canvas||null;
    if(raw&&raw.width===width&&raw.height===height)return raw;
    const canvas=createWhiteDiffCanvas(width,height);
    if(raw)canvas.getContext('2d',{alpha:false}).drawImage(raw,0,0);
    return canvas;
  };
  // 通常は前後PDFが同じ寸法なので、描画済みCanvasをそのまま解析・表示し、
  // 同じ全面Canvasの再作成とdrawImageを2回分省略する。
  const beforeCanvas=normalizeCanvas(beforeRaw),afterCanvas=normalizeCanvas(afterRaw);
  let regions=[],status='ready',message=String(mappedPage?.mappingMessage||''),analysis=null,noiseSuppressed=false;
  const exactSame=kind==='unchanged'||(!mappedPage&&asArray(sheet?.unchangedPageNumbers).map(Number).includes(pageNumber));
  if(kind==='added')regions=[fullDiffRegion('added',width,height)];
  else if(kind==='removed')regions=[fullDiffRegion('removed',width,height)];
  else if(kind==='unknown'){regions=[fullDiffRegion('unknown',width,height)];status='unknown';message=message||'信頼できる差分領域を判定できません。';}
  else if(kind!=='unchanged'&&!exactSame){
    setDiffBrowserProgress(true,'表示ページの違いを解析しています。',70);
    try{
      // 画像側が位置ずれを差分として拾っても数値照合を省略しない。PDF文字抽出は
      // ワーカー解析と並行して開始し、変更ページの待ち時間増加を抑える。
      const textPromise=beforeRaw&&afterRaw
        ?Promise.all([extractDiffPdfTextItems(beforeRaw,serial),extractDiffPdfTextItems(afterRaw,serial)]).catch(()=>null)
        :Promise.resolve(null);
      const documentTextPromise=beforeRaw&&afterRaw&&diffSheetPageCount(sheet)>1
        ?Promise.all([extractDiffPdfDocumentTextPages('before',sheet),extractDiffPdfDocumentTextPages('after',sheet)]).catch(()=>null)
        :Promise.resolve(null);
      analysis=await analyzeDiffCanvases(beforeCanvas,afterCanvas,width,height);if(serial!==diffBrowserRenderSerial)return null;regions=asArray(analysis.regions);
      if(analysis.fallbackUsed)message='小さい差分を検出したため、最も可能性の高い箇所を強調しています。';
      else if(analysis.alignmentAdjusted)message='行・列の追加や幅・倍率による位置ずれを補正して差分を絞り込みました。';
      setDiffBrowserProgress(true,'PDF内の文字を照合しています。',85);
      const textPair=await textPromise;
      const documentTextPair=await documentTextPromise;
      if(serial!==diffBrowserRenderSerial)return null;
      if(textPair){
        const documentRowDiff=documentTextPair?buildDocumentTextRowStructureDiffResult(documentTextPair[0],documentTextPair[1],width,height,beforePageNumber,afterPageNumber):null;
        const semantic=selectDiffSemanticResult(textPair[0],textPair[1],width,height,analysis,regions,documentRowDiff);
        regions=semantic.regions;
        if(semantic.message)message=semantic.message;
        if(semantic.suppressRaster)noiseSuppressed=true;
        if(semantic.mode!=='row'&&shouldSuppressDiffRasterNoise(textPair[0],textPair[1],analysis,regions,width,height)){
          regions=[];noiseSuppressed=true;message='PDF内の文字と配置が一致したため、微小な画像描画ノイズを除外しました。';
        }
      }
      if(!regions.length&&!noiseSuppressed){regions=[fullDiffRegion('modified',width,height)];status='unknown';message='変更は検出されましたが位置を絞り込めないため、ページ全体を強調しています。';}
    }
    catch(error){regions=[fullDiffRegion('unknown',width,height)];status='unknown';message=userFriendlyError(error.message);}
  }else if(kind==='modified'&&exactSame)message='このページに変更はありません。同じ原稿内の別ページに変更があります。';
  const page={pageNumber,beforePageNumber:needBefore?beforePageNumber:0,afterPageNumber:needAfter?afterPageNumber:0,comparisonKind:kind,width,height,pageSizeChanged:!!beforeRaw&&!!afterRaw&&(beforeRaw.canvas.width!==afterRaw.canvas.width||beforeRaw.canvas.height!==afterRaw.canvas.height),status,message,confidence:status==='unknown'?0:1,changedRatio:Number(analysis?.changedRatio||0),regionCount:regions.length,alignmentAdjusted:!!analysis?.alignmentAdjusted,regions};
  return {sheetKey:String(sheet?.sheetKey||''),pageIndex,page,beforeCanvas,afterCanvas};
}
function paintDiffBrowserPage(result){
  if(!result)return;
  copyDiffCanvas('diff-before-base',result.beforeCanvas);copyDiffCanvas('diff-after-underlay',result.beforeCanvas);copyDiffCanvas('diff-after-base',result.afterCanvas);
  const page=result.page;
  renderDiffRegionLayer('before',page.regions);renderDiffRegionLayer('after',page.regions);
  if(diffViewState.regionIndex>=page.regions.length)diffViewState.regionIndex=-1;
  if($('diff-region-count'))$('diff-region-count').textContent=page.regions.length?(diffViewState.regionIndex>=0?`${diffViewState.regionIndex+1} / ${page.regions.length}件`:`${page.regions.length}件`):'0件';
  $('diff-prev-region').disabled=!diffRegionTargets().length;$('diff-next-region').disabled=!diffRegionTargets().length;
  setDiffPaneEmpty('before',page.comparisonKind==='added'?'前回版には存在しません':'');
  setDiffPaneEmpty('after',page.comparisonKind==='removed'?'現在版では削除されています':'');
  applyDiffHighlightSettings();applyDiffMode();updateDiffStageScale();
}
function diffDetailRequestPath(){
  const params=new URLSearchParams({workbookId:diffViewState.workbookId});
  if(diffViewState.fromSnapshotId&&diffViewState.toSnapshotId){
    params.set('fromSnapshotId',diffViewState.fromSnapshotId);
    params.set('toSnapshotId',diffViewState.toSnapshotId);
  }
  return `/api/history/diff-detail?${params.toString()}`;
}
function prefetchAutomaticDiffDetail(workbookId){
  const id=String(workbookId||'');if(!id)return;
  const path='/api/history/diff-detail?'+new URLSearchParams({workbookId:id}).toString();
  void fetchDiffDetailResponse(path,id).catch(()=>{});
}

async function fetchDiffDetailResponse(path,workbookId){
  const workbook=getWorkbook(workbookId)||{};
  const key=path+'|'+String(workbook.lastRenderedSnapshotId||'')+'|'+String(workbook.lastRenderedVersionId||'');
  if(diffDetailResponseCache.has(key)){
    const cached=diffDetailResponseCache.get(key);
    diffDetailResponseCache.delete(key);diffDetailResponseCache.set(key,cached);
    return cached;
  }
  const promise=(async()=>{
    let lastError=null;
    for(let attempt=0;attempt<2;attempt++){
      const controller=new AbortController();
      const timer=setTimeout(()=>controller.abort(),DIFF_DETAIL_TIMEOUT_MS);
      try{return await api(path,{signal:controller.signal});}
      catch(error){
        lastError=error;
        if(error?.name!=='AbortError')throw error;
        if(attempt===0){
          if(isDiffModalOpen()&&diffViewState.workbookId===String(workbookId||'')&&$('diff-subtitle')){
            $('diff-subtitle').textContent='共有フォルダーの応答を再確認しています。';
          }
          await sleep(250);
          continue;
        }
      }finally{clearTimeout(timer);}
    }
    const error=new Error('共有フォルダーから比較情報を取得できませんでした。比較画面を閉じて、もう一度開いてください。');
    error.cause=lastError;throw error;
  })();
  diffDetailResponseCache.set(key,promise);
  while(diffDetailResponseCache.size>DIFF_DETAIL_CACHE_LIMIT)diffDetailResponseCache.delete(diffDetailResponseCache.keys().next().value);
  try{return await promise;}catch(error){diffDetailResponseCache.delete(key);throw error;}
}

function updateDiffProgress(generation={}){
  const box=$('diff-progress');
  const status=String(generation.status||'');
  const active=!['','idle','completed','failed'].includes(status);
  box?.classList.toggle('hidden',!active);
  const percent=Math.max(0,Math.min(100,Number(generation.percent||0)));
  if($('diff-progress-text'))$('diff-progress-text').textContent=String(generation.message||'差分画像を準備しています。');
  if($('diff-progress-fill'))$('diff-progress-fill').style.width=`${percent}%`;
}
function renderDiffSummary(){
  const box=$('diff-summary');if(!box)return;
  const summary=diffViewState.detail?.summary||{};
  const definitions=[
    ['all','すべて',asArray(diffViewState.detail?.sheets).length],
    ['modified','変更',Number(summary.changed||0)],
    ['added','追加',Number(summary.added||0)],
    ['removed','削除',Number(summary.removed||0)],
    ['unknown','判定不能',Number(summary.unknown||0)],
    ['unchanged','変更なし',Number(summary.unchanged||0)],
    ['confirmed','確認済み',Number(summary.confirmed||0)]
  ];
  box.innerHTML=definitions.map(([key,label,count])=>`<button class="${escapeAttr(key)} ${diffViewState.filter===key?'active':''}" type="button" data-diff-filter="${escapeAttr(key)}">${escapeHtml(label)} ${count}</button>`).join('');
  box.querySelectorAll('[data-diff-filter]').forEach(btn=>btn.addEventListener('click',()=>{
    diffViewState.filter=btn.dataset.diffFilter||'all';
    if(diffViewState.filter==='confirmed'){
      diffViewState.unreviewedOnly=false;
      if($('diff-unreviewed-only'))$('diff-unreviewed-only').checked=false;
    }
    const visible=diffFilteredSheets();
    if(!visible.some(s=>String(s.sheetKey||'')===diffViewState.selectedSheetKey)){
      diffViewState.selectedSheetKey=String(visible[0]?.sheetKey||'');
      diffViewState.pageIndex=preferredDiffPageIndex(visible[0]);diffViewState.regionIndex=-1;
    }
    renderDiffSummary();renderDiffSheetList();renderDiffPage();
  }));
}
function renderDiffSheetList(){
  const box=$('diff-sheet-list');if(!box)return;
  const sheets=diffFilteredSheets();
  if(!sheets.length){box.innerHTML='<div class="empty-state">比較対象のページ項目はありません。</div>';return;}
  box.innerHTML=sheets.map(sheet=>{
    const meta=diffKindMeta(sheet.kind);
    const pages=Math.max(Number(sheet.beforePages||0),Number(sheet.afterPages||0),asArray(sheet.pages).length);
    const sheetStatus=String(sheet.status||'');
    const cached=[...diffBrowserPageCache.values()].filter(item=>item.sheetKey===String(sheet.sheetKey||'')).length;
    const confidence=diffMatchConfidenceLabel(sheet);
    const browserDetail=sourceTypeValue(diffSource())==='excel'?`${pages}ページ・ブラウザ比較${cached?`済 ${cached}`:'（表示時）'}`:`ブラウザ比較${cached?'済':'（表示時）'}`;
    const detail=sheetStatus==='unknown'?(sheet.message||'対応を確定できないため目視確認が必要です。'):browserDetail;
    const displayName=diffUnitDisplayName(sheet);
    return `<button class="diff-sheet-item ${sheet.confirmed?'confirmed':''} ${String(sheet.sheetKey||'')===diffViewState.selectedSheetKey?'active':''}" type="button" data-diff-sheet="${escapeAttr(sheet.sheetKey||'')}" title="${escapeAttr(`${displayName}・${confidence}${sheet.confirmed?'・確認済み':''}`)}"><span class="diff-sheet-kind ${escapeAttr(sheet.kind||'unknown')}">${escapeHtml(meta.mark)}</span><span class="diff-sheet-copy"><strong>${escapeHtml(displayName)}</strong><small>${escapeHtml(meta.label)}・${escapeHtml(confidence)}・${escapeHtml(detail)}</small></span><span class="diff-sheet-reviewed" aria-label="${sheet.confirmed?'確認済み':'未確認'}">${sheet.confirmed?'✓':''}</span></button>`;
  }).join('');
  box.querySelectorAll('[data-diff-sheet]').forEach(btn=>{
    btn.addEventListener('click',()=>selectDiffSheet(btn.dataset.diffSheet||''));
    btn.addEventListener('keydown',e=>{
      if(!['ArrowDown','ArrowUp'].includes(e.key))return;
      e.preventDefault();
      const buttons=[...box.querySelectorAll('[data-diff-sheet]')];
      const index=buttons.indexOf(btn);
      const next=buttons[Math.max(0,Math.min(buttons.length-1,index+(e.key==='ArrowDown'?1:-1)))];
      if(!next)return;
      // selectDiffSheet re-renders this list, destroying the node we just
      // focused, so focus has to be re-applied to its replacement afterwards.
      const nextKey=next.dataset.diffSheet||'';
      selectDiffSheet(nextKey,false);
      box.querySelector(`[data-diff-sheet="${CSS.escape(nextKey)}"]`)?.focus();
    });
  });
}
function selectDiffSheet(sheetKey,focusList=true){
  const sheet=asArray(diffViewState.detail?.sheets).find(s=>String(s.sheetKey||'')===String(sheetKey||''));
  if(!sheet)return;
  diffViewState.selectedSheetKey=String(sheet.sheetKey||'');diffViewState.pageIndex=preferredDiffPageIndex(sheet);diffViewState.regionIndex=-1;
  renderDiffSheetList();renderDiffPage();
  if(focusList)$('diff-before-viewport')?.focus();
}
function renderDiffReviewControl(){
  const button=$('diff-confirm-sheet'),sheet=diffCurrentSheet();if(!button)return;
  button.disabled=!sheet;
  button.classList.toggle('active',!!sheet?.confirmed);
  button.textContent=sheet?.confirmed?'確認済みを解除':'この項目を確認済みにする';
  button.setAttribute('aria-pressed',String(!!sheet?.confirmed));
}
async function toggleDiffSheetReviewed(button=$('diff-confirm-sheet')){
  const sheet=diffCurrentSheet(),cmp=diffViewState.detail?.comparison||{};if(!sheet||!diffViewState.detail)return;
  const confirmed=!sheet.confirmed;
  await runBusy(button,async()=>{
    const response=await api('/api/history/diff-review',{method:'POST',body:{
      workbookId:diffViewState.workbookId,fromSnapshotId:diffViewState.fromSnapshotId,toSnapshotId:diffViewState.toSnapshotId,
      baselineVersionId:String(cmp.baselineVersionId||''),currentVersionId:String(cmp.currentVersionId||''),sheetKey:String(sheet.sheetKey||''),confirmed
    }});
    const confirmedKeys=new Set(asArray(response.review?.confirmedSheetKeys).map(String));
    for(const item of asArray(diffViewState.detail?.sheets))item.confirmed=confirmedKeys.has(String(item.sheetKey||''));
    diffViewState.detail.review=response.review||null;
    if(diffViewState.detail.summary)diffViewState.detail.summary.confirmed=confirmedKeys.size;
    diffDetailResponseCache.clear();
    const visible=diffFilteredSheets();
    if(!visible.some(item=>String(item.sheetKey||'')===diffViewState.selectedSheetKey)){
      diffViewState.selectedSheetKey=String(visible[0]?.sheetKey||'');diffViewState.pageIndex=preferredDiffPageIndex(visible[0]);diffViewState.regionIndex=-1;
    }
    renderDiffSummary();renderDiffSheetList();renderDiffPage();
  },false);
}
function setDiffPaneEmpty(side,message){
  const stage=$(`diff-${side}-stage`),empty=$(`diff-${side}-empty`);
  stage?.classList.toggle('hidden',!!message);
  empty?.classList.toggle('hidden',!message);
  if(empty)empty.textContent=message||'';
}
function updateDiffStageScale(){
  if(!isDiffModalOpen())return;
  const page=diffCurrentPage();if(!page)return;
  const width=Math.max(1,Number(page.width||1)),height=Math.max(1,Number(page.height||1));
  const zoom=String($('diff-zoom')?.value||'fit-width');
  const visibleViewports=[$('diff-before-viewport'),$('diff-after-viewport')].filter(v=>v&&v.offsetParent!==null);
  let scale=1;
  if(zoom==='fit-width'){
    const available=Math.max(160,Math.min(...visibleViewports.map(v=>v.clientWidth||160))-40);
    scale=available/width;
  }else if(zoom==='fit-page'){
    const availableWidth=Math.max(160,Math.min(...visibleViewports.map(v=>v.clientWidth||160))-40);
    const availableHeight=Math.max(180,Math.min(...visibleViewports.map(v=>v.clientHeight||180))-40);
    scale=Math.min(availableWidth/width,availableHeight/height);
  }else{
    scale=Math.max(.25,Math.min(3,Number(zoom)||1));
  }
  for(const id of ['diff-before-stage','diff-after-stage']){
    const stage=$(id);if(!stage)continue;
    stage.style.width=`${Math.max(1,Math.round(width*scale))}px`;
    stage.style.height=`${Math.max(1,Math.round(height*scale))}px`;
  }
  requestAnimationFrame(focusCurrentDiffRegion);
}
function applyDiffHighlightSettings(){
  const enabled=!!$('diff-highlight')?.checked;
  const opacity=Number($('diff-density')?.value||.25);
  document.querySelectorAll('#diff-modal .diff-region-layer').forEach(el=>{
    el.style.display=enabled?'block':'none';
    el.style.setProperty('--diff-region-opacity',String(opacity));
  });
}
function renderDiffRegionLayer(side,regions){
  const layer=$(`diff-${side}-regions`);if(!layer)return;
  layer.replaceChildren();
  const visible=asArray(regions).filter(region=>{
    const kind=String(region?.kind||'modified');
    return kind==='modified'||kind==='unknown'||(kind==='added'&&side==='after')||(kind==='removed'&&side==='before');
  });
  const showModifiedTags=visible.length<=12;
  for(const region of visible){
    const kind=['modified','added','removed','unknown'].includes(String(region?.kind||''))?String(region.kind):'modified';
    const geometry=region?.[side]||region;
    const box=document.createElement('span');
    box.className=`diff-region-box ${kind}`;
    const x=Math.max(0,Math.min(1,Number(geometry?.x||0)));
    const y=Math.max(0,Math.min(1,Number(geometry?.y||0)));
    const width=Math.max(0,Math.min(1-x,Number(geometry?.width||0)));
    const height=Math.max(0,Math.min(1-y,Number(geometry?.height||0)));
    box.style.left=`${x*100}%`;box.style.top=`${y*100}%`;box.style.width=`${width*100}%`;box.style.height=`${height*100}%`;
    if(y<.018)box.classList.add('tag-inside');
    if(kind!=='modified'||showModifiedTags){
      const tag=document.createElement('span');
      tag.className='diff-region-tag';
      tag.textContent=kind==='added'?'A':kind==='removed'?'D':kind==='unknown'?'?':'M';
      box.appendChild(tag);
    }
    layer.appendChild(box);
  }
}
function applyDiffMode(){
  const overlay=diffViewState.mode==='overlay';
  $('diff-viewers')?.classList.toggle('overlay-mode',overlay);
  $('diff-mobile-tabs')?.classList.toggle('hidden',overlay);
  $('diff-mode-side')?.classList.toggle('active',!overlay);
  $('diff-mode-overlay')?.classList.toggle('active',overlay);
  $('diff-mode-side')?.setAttribute('aria-pressed',String(!overlay));
  $('diff-mode-overlay')?.setAttribute('aria-pressed',String(overlay));
  $('diff-opacity-wrap')?.classList.toggle('hidden',!overlay);
  $('diff-after-underlay')?.classList.toggle('hidden',!overlay);
  const opacity=Math.max(0,Math.min(100,Number($('diff-opacity')?.value||50)));
  if($('diff-after-base'))$('diff-after-base').style.opacity=overlay?String(opacity/100):'1';
  if($('diff-opacity-value'))$('diff-opacity-value').textContent=`${opacity}%`;
  requestAnimationFrame(updateDiffStageScale);
}
function renderDiffPage(){
  const sheet=diffCurrentSheet();
  renderDiffReviewControl();
  const historical=String(diffViewState.detail?.comparison?.scope||'')==='history';
  const beforeLabel=historical?'比較元':'前回版',afterLabel=historical?'比較先':'現在版';
  const messageBox=$('diff-state-message'),serial=beginDiffBrowserRender();
  for(const id of ['diff-before-base','diff-after-underlay','diff-after-base'])clearDiffCanvas(id);
  renderDiffRegionLayer('before',[]);renderDiffRegionLayer('after',[]);
  if(!sheet){
    setDiffBrowserProgress(false);
    if(messageBox){messageBox.textContent=diffViewState.detail?.message||'比較対象のページ項目がありません。';messageBox.classList.remove('hidden');}
    setDiffPaneEmpty('before','比較するページがありません。');setDiffPaneEmpty('after','比較するページがありません。');return;
  }
  const pageTotal=diffSheetPageCount(sheet);
  diffViewState.pageIndex=Math.max(0,Math.min(diffViewState.pageIndex,Math.max(0,pageTotal-1)));
  const pageNumber=diffViewState.pageIndex+1,mappedPage=asArray(sheet?.pages)[diffViewState.pageIndex]||null;
  const beforePageNumber=mappedPage&&Object.hasOwn(mappedPage,'beforePageNumber')?Number(mappedPage.beforePageNumber||0):pageNumber;
  const afterPageNumber=mappedPage&&Object.hasOwn(mappedPage,'afterPageNumber')?Number(mappedPage.afterPageNumber||0):pageNumber;
  if($('diff-before-page-label'))$('diff-before-page-label').textContent=beforePageNumber?`${beforePageNumber} / ${Number(sheet?.beforePages||pageTotal)}ページ`:'—';
  if($('diff-after-page-label'))$('diff-after-page-label').textContent=afterPageNumber?`${afterPageNumber} / ${Number(sheet?.afterPages||pageTotal)}ページ`:'—';
  if($('diff-page-count'))$('diff-page-count').textContent=`${pageNumber} / ${pageTotal}件`;
  $('diff-prev-page').disabled=diffViewState.pageIndex<=0;$('diff-next-page').disabled=diffViewState.pageIndex>=pageTotal-1;
  diffViewState.regionIndex=-1;if($('diff-region-count'))$('diff-region-count').textContent='0件';
  $('diff-prev-region').disabled=true;$('diff-next-region').disabled=true;
  if(messageBox){messageBox.textContent='';messageBox.classList.add('hidden');}
  if(!pageTotal){
    setDiffBrowserProgress(false);
    setDiffPaneEmpty('before',sheet.kind==='added'?`${beforeLabel}には存在しません`:'表示できるPDFページがありません。');
    setDiffPaneEmpty('after',sheet.kind==='removed'?`${afterLabel}では削除されています`:'表示できるPDFページがありません。');return;
  }
  const key=diffBrowserPageKey(sheet,diffViewState.pageIndex),cached=diffBrowserPageCache.get(key);
  if(cached){
    diffBrowserPageCache.delete(key);diffBrowserPageCache.set(key,cached);
    setDiffBrowserProgress(false);paintDiffBrowserPage(cached);renderDiffSheetList();return;
  }
  setDiffPaneEmpty('before',sheet.kind==='added'?`${beforeLabel}には存在しません`:'PDFを読み込んでいます…');
  setDiffPaneEmpty('after',sheet.kind==='removed'?`${afterLabel}では削除されています`:'PDFを読み込んでいます…');
  void buildDiffBrowserPage(sheet,diffViewState.pageIndex,serial).then(result=>{
    if(!result||serial!==diffBrowserRenderSerial||!isDiffModalOpen())return;
    cacheDiffBrowserPage(key,result);setDiffBrowserProgress(false);paintDiffBrowserPage(result);renderDiffSheetList();
    const messages=[result.page.message,result.page.pageSizeChanged?'ページサイズが異なるため、左上を基準に揃えて表示しています。':''].filter(Boolean);
    if(messageBox){messageBox.textContent=messages.join(' ');messageBox.classList.toggle('hidden',!messages.length);}
  }).catch(error=>{
    if(serial!==diffBrowserRenderSerial||!isDiffModalOpen())return;
    setDiffBrowserProgress(false);setDiffPaneEmpty('before','PDFを表示できませんでした。');setDiffPaneEmpty('after','PDFを表示できませんでした。');
    if(messageBox){messageBox.textContent=userFriendlyError(error.message);messageBox.classList.remove('hidden');}
  });
}
function focusCurrentDiffRegion(){
  const page=diffCurrentPage(),region=asArray(page?.regions)[diffViewState.regionIndex]||null;
  for(const side of ['before','after']){
    const focus=$(`diff-${side}-focus`);
    if(!focus)continue;
    focus.classList.toggle('hidden',!region);
    if(!region)continue;
    const geometry=region?.[side]||region;
    focus.style.left=`${Number(geometry.x||0)*100}%`;
    focus.style.top=`${Number(geometry.y||0)*100}%`;
    focus.style.width=`${Number(geometry.width||0)*100}%`;
    focus.style.height=`${Number(geometry.height||0)*100}%`;
    const stage=$(`diff-${side}-stage`),viewport=$(`diff-${side}-viewport`);
    if(stage&&viewport&&viewport.offsetParent!==null){
      const centerX=(Number(geometry.x||0)+Number(geometry.width||0)/2)*stage.offsetWidth;
      const centerY=(Number(geometry.y||0)+Number(geometry.height||0)/2)*stage.offsetHeight;
      viewport.scrollTo({left:Math.max(0,centerX-viewport.clientWidth/2),top:Math.max(0,centerY-viewport.clientHeight/2),behavior:'smooth'});
    }
  }
}
function diffRegionTargets(){
  const order=new Map(asArray(diffViewState.detail?.sheets).map((sheet,index)=>[String(sheet.sheetKey||''),index]));
  const visibleSheetKeys=new Set(diffFilteredSheets().map(sheet=>String(sheet.sheetKey||'')));
  const targets=[];
  for(const item of diffBrowserPageCache.values())if(visibleSheetKeys.has(String(item.sheetKey||'')))asArray(item.page?.regions).forEach((region,regionIndex)=>targets.push({sheetKey:item.sheetKey,pageIndex:item.pageIndex,regionIndex,region}));
  return targets.sort((x,y)=>(order.get(x.sheetKey)||0)-(order.get(y.sheetKey)||0)||x.pageIndex-y.pageIndex||x.regionIndex-y.regionIndex);
}
function moveDiffRegion(direction){
  const targets=diffRegionTargets();if(!targets.length)return;
  let index=targets.findIndex(t=>t.sheetKey===diffViewState.selectedSheetKey&&t.pageIndex===diffViewState.pageIndex&&t.regionIndex===diffViewState.regionIndex);
  if(index<0)index=direction>0?-1:targets.length;
  index=Math.max(0,Math.min(targets.length-1,index+direction));
  const target=targets[index];
  diffViewState.selectedSheetKey=target.sheetKey;
  diffViewState.pageIndex=target.pageIndex;
  diffViewState.regionIndex=target.regionIndex;
  renderDiffSheetList();renderDiffPage();
  setTimeout(focusCurrentDiffRegion,80);
}
function setDiffPage(index){
  const count=diffSheetPageCount(diffCurrentSheet());if(!count)return;
  diffViewState.pageIndex=Math.max(0,Math.min(count-1,index));diffViewState.regionIndex=-1;renderDiffPage();
}
function renderDiffDetail(detail){
  diffViewState.detail=detail||null;
  const stateBox=$('diff-state-message');
  if(stateBox){stateBox.textContent='';stateBox.classList.add('hidden');}
  const workbookName=String(detail?.workbookName||getWorkbook(diffViewState.workbookId)?.displayName||'Excel');
  if($('diff-title'))$('diff-title').textContent=`差分詳細：${workbookName}`;
  const cmp=detail?.comparison||{},historical=String(cmp.scope||'')==='history';
  if($('diff-subtitle'))$('diff-subtitle').textContent=`${historical?'比較元':'前回版'} ${formatDateTime(cmp.baselineAt)} → ${historical?'比較先':'現在版'} ${formatDateTime(cmp.currentAt)} ・ 表示ページをブラウザで比較`;
  if($('diff-before-label'))$('diff-before-label').textContent=historical?'比較元':'前回版';
  if($('diff-after-label'))$('diff-after-label').textContent=historical?'比較先':'現在版';
  const tabs=$('diff-mobile-tabs')?.querySelectorAll('button');
  if(tabs?.[0])tabs[0].textContent=historical?'比較元':'前回版';
  if(tabs?.[1])tabs[1].textContent=historical?'比較先':'現在版';
  setDiffBrowserProgress(false);
  const sheets=asArray(detail?.sheets);
  if($('diff-export-summary'))$('diff-export-summary').disabled=!sheets.length;
  if(!sheets.some(s=>String(s.sheetKey||'')===diffViewState.selectedSheetKey)){
    const preferred=sheets.find(s=>['modified','added','removed','unknown'].includes(String(s.kind||'')))||sheets[0];
    diffViewState.selectedSheetKey=String(preferred?.sheetKey||'');diffViewState.pageIndex=preferredDiffPageIndex(preferred);diffViewState.regionIndex=-1;
  }
  renderDiffSummary();renderDiffSheetList();renderDiffPage();
  if(['unavailable','failed'].includes(String(detail?.status||''))&&stateBox){stateBox.textContent=detail?.message||'差分詳細を表示できません。';stateBox.classList.remove('hidden');}
}
async function openDiffDetail(workbookId,opener,historyRange=null){
  const id=String(workbookId||'');if(!id)return;
  diffReturnFocus=opener||document.activeElement;clearDiffBrowserResources();
  diffViewState.workbookId=id;diffViewState.fromSnapshotId=String(historyRange?.fromSnapshotId||'');diffViewState.toSnapshotId=String(historyRange?.toSnapshotId||'');
  diffViewState.detail=null;diffViewState.selectedSheetKey='';diffViewState.pageIndex=0;diffViewState.regionIndex=-1;diffViewState.filter='all';diffViewState.unreviewedOnly=false;diffViewState.mode='side';diffViewState.mobileTab='before';
  if($('diff-unreviewed-only'))$('diff-unreviewed-only').checked=false;
  if($('diff-export-summary'))$('diff-export-summary').disabled=true;
  $('diff-modal')?.classList.remove('hidden');document.body.style.overflow='hidden';
  if($('diff-title'))$('diff-title').textContent='差分詳細：'+workbookDisplayName(getWorkbook(id));
  if($('diff-subtitle'))$('diff-subtitle').textContent='比較情報を読み込んでいます。';
  if($('diff-sheet-list'))$('diff-sheet-list').innerHTML='<div class="empty-state">読み込んでいます。</div>';
  setDiffPaneEmpty('before','比較情報を読み込んでいます。');setDiffPaneEmpty('after','比較情報を読み込んでいます。');$('diff-close')?.focus();
  const requestPath=diffDetailRequestPath();
  // 比較情報の取得と並行してPDF.js・差分Workerを準備し、初回ページの開始待ちを短くする。
  const detailRequest=fetchDiffDetailResponse(requestPath,id);
  void ensureDiffPdfJs().catch(()=>{});
  try{getDiffAnalysisWorker();}catch{}
  try{const loaded=await detailRequest;if(!isDiffModalOpen()||diffViewState.workbookId!==id)return;renderDiffDetail(loaded.detail);}
  // 成功側と同じ古さガードを掛ける。これが無いと、閉じた後や別の原稿へ切り替えた後に
  // 届いた失敗応答が、いま表示している比較画面を上書きしてしまう。
  catch(error){if(!isDiffModalOpen()||diffViewState.workbookId!==id)return;renderDiffDetail({status:'failed',message:userFriendlyError(error.message),workbookName:workbookDisplayName(getWorkbook(id)),sheets:[],summary:{},generation:{status:'failed'}});}
}

function closeDiffDetail(){
  if(!isDiffModalOpen())return;
  diffPollToken++;clearDiffBrowserResources();$('diff-modal')?.classList.add('hidden');document.body.style.overflow='';
  let target=diffReturnFocus;
  if(!target?.isConnected){const workbookId=String(diffViewState.workbookId||'');target=[...document.querySelectorAll('[data-open-diff]')].find(el=>String(el.dataset.openDiff||'')===workbookId)||null;}
  diffReturnFocus=null;if(target&&typeof target.focus==='function')target.focus();
}

async function saveSourceMetadata(sourceId, patch, control) {
  const id=String(sourceId||'');if(!id)return;
  if(control)control.disabled=true;
  try{
    const response=await api(`/api/v2/sources/${encodeURIComponent(id)}`,{method:'PATCH',body:patch});
    state=normalizeStatePayload(response.state||await api('/api/state'));
    if(control)control.disabled=false;
    renderAll({preserveEditors:true});
    showMessage('ok','原稿設定を保存しました',`${workbookDisplayName(getWorkbook(id))}の担当情報を更新しました。`,null,null,3500);
  }catch(error){
    if(control)control.disabled=false;renderWorkbooks();
    showMessage('danger','原稿設定を保存できません',userFriendlyError(error.message),error.detail||null);
  }
}
function activeSourceRequirements(){return asArray(activePackRecord()?.sourceRequirements);}
function renderSourceRequirementStatus(workbooks=workbooksForActivePreset()){
  const box=$('source-requirement-status');if(!box)return;const requirements=activeSourceRequirements();box.classList.toggle('hidden',requirements.length===0);if(!requirements.length){box.innerHTML='';return;}
  const rows=requirements.map(requirement=>{const assigned=asArray(workbooks).find(w=>String(w?.requirementId||'')===String(requirement?.requirementId||''));let stateLabel='未提出',stateClass=requirement?.required===false?'neutral':'danger';if(assigned){if(String(assigned.status||'')==='missing'){stateLabel='ファイルなし';stateClass='danger';}else if(isLatestPdfWorkbook(assigned)){stateLabel='提出・変換済み';stateClass='ok';}else{stateLabel='提出済み・変換待ち';stateClass='attention';}}const details=[requirement.ownerDepartment||'',requirement.dueDate?`期限 ${formatDateOnly(requirement.dueDate)}`:'',requirement.required===false?'任意':'必須'].filter(Boolean).join('・');return `<div class="source-requirement-item ${assigned?'assigned':'unassigned'}"><div><strong>${escapeHtml(requirement.displayName||'必要原稿')}</strong><span>${escapeHtml(details)}</span></div><div><span class="badge ${stateClass}">${escapeHtml(stateLabel)}</span><small>${escapeHtml(assigned?workbookDisplayName(assigned):'登録後、原稿設定から割り当てます')}</small></div></div>`;});
  const complete=rows.length&&requirements.every(requirement=>requirement.required===false||asArray(workbooks).some(w=>String(w?.requirementId||'')===String(requirement?.requirementId||'')&&String(w?.status||'')!=='missing'));
  box.innerHTML=`<div class="source-requirement-title"><strong>必要原稿</strong><span>${complete?'必須原稿はすべて提出済みです':'未提出の必須原稿があります'}</span></div><div class="source-requirement-list">${rows.join('')}</div>`;
}
function sourceSettingsHtml(w){
  const id=String(w?.workbookId||''),owner=String(w?.ownerDepartment||''),required=w?.required!==false,defaultTarget=String(w?.defaultTargetId||'unassigned');
  const sourceType=sourceTypeValue(w),requirements=activeSourceRequirements().filter(item=>asArray(item?.acceptedSourceTypes).map(String).includes(sourceType)),assignedElsewhere=new Map(workbooksForActivePreset().filter(item=>String(item?.workbookId||'')!==id&&item?.requirementId).map(item=>[String(item.requirementId),workbookDisplayName(item)])),requirementId=String(w?.requirementId||'');
  const requirementSelect=requirements.length?`<label class="source-requirement-label">必要原稿<select class="source-requirement-select" data-source-requirement="${escapeAttr(id)}" aria-label="${escapeAttr(workbookDisplayName(w))}の必要原稿"><option value="">割り当てなし</option>${requirements.map(item=>{const rid=String(item.requirementId||''),occupied=assignedElsewhere.get(rid);return `<option value="${escapeAttr(rid)}" ${rid===requirementId?'selected':''} ${occupied?'disabled':''}>${escapeHtml(item.displayName||'必要原稿')}${item.required===false?'（任意）':'（必須）'}${occupied?`・${escapeHtml(occupied)}へ割当済み`:''}</option>`;}).join('')}</select></label>`:'';
  const targetOptions=activeTargetDefinitions().map(target=>`<option value="${escapeAttr(target.targetId)}" ${defaultTarget===String(target.targetId)?'selected':''}>${escapeHtml(target.displayName||target.targetId)}</option>`).join('');
  return `<details class="source-advanced-settings" data-source-settings-details="${escapeAttr(id)}" ${expandedSourceSettings.has(id)?'open':''}><summary>原稿の扱いを変更</summary><div class="source-settings" aria-label="原稿の詳細設定">${requirementSelect}<label class="source-owner-label"><span>担当部署（任意）<small>進捗一覧に表示する管理用メモです</small></span><input class="source-owner-input" type="text" maxlength="120" value="${escapeAttr(owner)}" placeholder="例：営業部" data-source-owner="${escapeAttr(id)}" aria-label="${escapeAttr(workbookDisplayName(w))}の担当部署"></label><label class="source-required-label"><input type="checkbox" data-source-required="${escapeAttr(id)}" ${required?'checked':''}><span>提出用PDFに必須<small>未作成・更新あり・エラーの間は提出用PDFを出力しません</small></span></label><label class="source-default-target-label"><span>更新で増えたページの追加先<small>既に並べたページは移動しません</small></span><select class="source-default-target" data-source-default-target="${escapeAttr(id)}" aria-label="${escapeAttr(workbookDisplayName(w))}の追加ページ配置"><option value="unassigned" ${defaultTarget==='unassigned'?'selected':''}>一式の既定：${escapeHtml(activePackDefaultTargetLabel())}</option>${targetOptions}</select></label></div></details>`;
}
function syncDiffScroll(source,target){
  if(diffSyncingScroll||diffViewState.mode!=='side'||!source||!target)return;
  diffSyncingScroll=true;
  const maxX=Math.max(1,source.scrollWidth-source.clientWidth),maxY=Math.max(1,source.scrollHeight-source.clientHeight);
  const targetMaxX=Math.max(0,target.scrollWidth-target.clientWidth),targetMaxY=Math.max(0,target.scrollHeight-target.clientHeight);
  target.scrollLeft=(source.scrollLeft/maxX)*targetMaxX;
  target.scrollTop=(source.scrollTop/maxY)*targetMaxY;
  requestAnimationFrame(()=>{diffSyncingScroll=false;});
}

function renderWorkbooks() {
  queueMicrotask(syncSourcePanelDensity);
  const list=workbooksForActivePreset();const box=$('workbook-list');if(!box)return;const summary=$('workbook-selection-summary');
  renderSourceRequirementStatus(list);
  if(!list.length){box.className='table-shell empty-state';box.textContent=`${presetLabel()}の登録済み原稿はまだありません。`;selectedWorkbooks.clear();lastWorkbookRangeAnchor='';if(summary)summary.textContent='0件中 0件を選択';syncWorkbookSelectionUi();return;}
  selectedWorkbooks=new Set([...selectedWorkbooks].filter(id=>list.some(w=>String(w.workbookId)===String(id))));const allSelected=list.every(w=>selectedWorkbooks.has(String(w.workbookId||''))),someSelected=list.some(w=>selectedWorkbooks.has(String(w.workbookId||'')));
  box.className='table-shell';box.innerHTML=`<table class="data-table workbook-table"><thead><tr><th class="check-col"><input type="checkbox" data-select-all-workbooks ${allSelected?'checked':''} aria-label="登録済み原稿をすべて選択"></th><th>原稿ファイル</th><th class="pdf-status-col">変換PDFの状態</th></tr></thead><tbody>${list.map(w=>{const id=String(w.workbookId||''),checked=selectedWorkbooks.has(id),renderedAt=formatDateTime(w.lastRenderedAt||'');return `<tr data-workbook-row class="${checked?'selected-row':''} ${w.status==='render-error'?'error-row':''}"><td class="check-col"><input type="checkbox" data-workbook-check value="${escapeAttr(id)}" ${checked?'checked':''} aria-label="${escapeAttr(workbookDisplayName(w))}を選択"></td><td><div class="file-name-cell"><span class="file-icon ${escapeAttr(sourceTypeValue(w))}">${iconUse(sourceIconId(w))}</span><div class="source-file-detail"><div class="source-name-line"><strong title="${escapeAttr(workbookDisplayName(w))}">${escapeHtml(workbookDisplayName(w))}</strong>${sourceTypeBadge(w)}</div><div class="file-meta" title="最後に変換PDFを作成した日時">変換PDF作成日時：${escapeHtml(renderedAt)}</div>${sourceSettingsHtml(w)}</div></div></td><td class="pdf-status-col">${workbookPdfStatusCell(w)}</td></tr>`;}).join('')}</tbody></table>`;
  if(summary)summary.textContent=`${list.length}件中 ${selectedWorkbooks.size}件を選択`;const selectAll=box.querySelector('[data-select-all-workbooks]');if(selectAll){selectAll.indeterminate=someSelected&&!allSelected;selectAll.addEventListener('change',e=>{selectedWorkbooks=e.target.checked?new Set(list.map(w=>String(w.workbookId||'')).filter(Boolean)):new Set();lastWorkbookRangeAnchor='';syncWorkbookSelectionUi();renderWorkbooks();});}
  box.querySelectorAll('[data-workbook-check]').forEach(ch=>ch.addEventListener('click',e=>{e.stopPropagation();handleWorkbookCheckboxToggle(ch,e.shiftKey);renderWorkbooks();}));box.querySelectorAll('[data-open-history]').forEach(btn=>btn.addEventListener('click',e=>{e.stopPropagation();openWorkbookComparisonHistory(btn.dataset.openHistory||'');}));
  box.querySelectorAll('[data-relink-source]').forEach(btn=>btn.addEventListener('click',e=>{e.stopPropagation();openRelinkDialog(btn.dataset.relinkSource||'',btn);}));
  box.querySelectorAll('[data-source-owner]').forEach(input=>input.addEventListener('change',()=>saveSourceMetadata(input.dataset.sourceOwner,{ownerDepartment:input.value},input)));
  box.querySelectorAll('[data-source-required]').forEach(input=>input.addEventListener('change',()=>saveSourceMetadata(input.dataset.sourceRequired,{required:input.checked},input)));
  box.querySelectorAll('[data-source-default-target]').forEach(select=>select.addEventListener('change',()=>saveSourceMetadata(select.dataset.sourceDefaultTarget,{defaultTargetId:select.value},select)));
  box.querySelectorAll('[data-source-requirement]').forEach(select=>select.addEventListener('change',()=>saveSourceMetadata(select.dataset.sourceRequirement,{requirementId:select.value},select)));
  box.querySelectorAll('[data-source-settings-details]').forEach(details=>details.addEventListener('toggle',()=>{const id=String(details.dataset.sourceSettingsDetails||'');if(!id)return;if(details.open)expandedSourceSettings.add(id);else expandedSourceSettings.delete(id);}));
  box.querySelectorAll('[data-workbook-row]').forEach(row=>row.addEventListener('click',e=>{if(e.target.closest('input,button,a,select,textarea,label,summary,details'))return;const ch=row.querySelector('[data-workbook-check]');if(!ch)return;ch.checked=!ch.checked;handleWorkbookCheckboxToggle(ch,e.shiftKey);renderWorkbooks();}));syncWorkbookSelectionUi();
}

function summarizeRenderResults(results) {
  const list = results || [];
  const errors = list.filter(r => r.ok === false || r.error);
  const warnings = [];
  for (const r of list) {
    const wbName = getWorkbook(r.workbookId)?.displayName || getWorkbook(r.workbookId)?.fileName || '';
    const prefix = wbName ? wbName + '：' : '';
    for (const w of r.warnings || []) if (!/縮尺例外|Zoom|倍率例外/.test(String(w))) warnings.push(prefix + w);
    for (const p of r.rendered || []) for (const w of p.warnings || []) if (!/縮尺例外|Zoom|倍率例外/.test(String(w))) warnings.push(prefix + (p.sheetName ? '[' + p.sheetName + '] ' : '') + w);
  }
  if (errors.length) {
    const names = errors.map(e => getWorkbook(e.workbookId)?.displayName || getWorkbook(e.workbookId)?.fileName || e.workbookId).filter(Boolean).slice(0, 2).join('、');
    const first = userFriendlyError(errors[0].userError || errors[0].error || errors[0].message || 'PDF化できませんでした。');
    return {kind:'danger', title:'PDF作成を確認してください', message:names ? `${names}：${first}` : first, detail:list};
  }
  if (warnings.length) {
    const shown = warnings.slice(0, 6).join('\n');
    const more = warnings.length > 6 ? `\n…ほか ${warnings.length - 6}件。「原稿を登録・PDF化」の登録済み原稿一覧で、対象の原稿の状態欄を確認してください` : '';
    return {kind:'warn', title:'PDFを作成しました(要確認)', message: shown + more, detail:list};
  }
  if (list.length) return {kind:'ok', title:'PDFを作成しました', message:`${list.length}件の原稿を処理しました。`, detail:list};
  return {kind:'ok', title:'PDF作成は不要です', message:'対象はありません。', detail:null};
}
async function renderWorkbookIds(ids, btn, options={}) {
  let workbookIds=normalizeWorkbookIdListForActivePreset(ids);let onlyUpdated=!!options.onlyUpdated&&!workbookIds.length;const activeWorkbooks=workbooksForActivePreset().filter(w=>String(w.status||'')!=='missing');
  if(onlyUpdated){workbookIds=defaultPdfTargetWorkbookIds();onlyUpdated=false;}
  if(!workbookIds.length&&!onlyUpdated){hideProgressPanel();showMessage(activeWorkbooks.length?'ok':'warn',activeWorkbooks.length?'変換PDF作成が必要な原稿はありません':'登録済み原稿がありません',activeWorkbooks.length?'登録済み原稿の変換PDFは最新です。':'先に原稿を登録してください。');return;}
  renderJobActive=true;activeRenderJobId='';activeRenderCancelRequested=false;activeRenderJobStatus=null;const estimateTotal=workbookIds.length||defaultPdfTargetWorkbookIds().length||(onlyUpdated?activeWorkbooks.length:0);const startMessage=options.startMessage||(onlyUpdated?'PDF作成対象を確認しています。':`${workbookIds.length}件の原稿をPDF変換します。`);
  updateProgressPanel({status:'queued',total:estimateTotal,completed:0,failed:0,percent:0,message:startMessage});
  try{await runBusy(btn,async()=>{hideErrorPanel();showMessage('','PDF作成を開始しています',startMessage);const body={packId:activePackId};if(workbookIds.length)body.sourceIds=workbookIds;if(onlyUpdated)body.onlyUpdated=true;
    const started=await api('/api/v2/sources/render/start',{method:'POST',body});const job=normalizeRenderJobFromStart(started);const jobId=assertJobId(job,started);activeRenderJobId=jobId;updateProgressPanel(Object.assign({total:estimateTotal,completed:0,failed:0},job));const finished=await waitForRenderJob(jobId);const appendedIds=asArray(finished.results).flatMap(r=>asArray(r?.sheetSync?.insertedAtEndPageIds)).map(String).filter(Boolean);setHighlightedPages(appendedIds);await refresh();selectedWorkbooks.clear();lastWorkbookRangeAnchor='';renderAll();await loadFiles(null);
    const msg=summarizeRenderResults(finished.results||[]);const hasErrors=finished.status==='failed'||finished.status==='completed-with-errors'||asArray(finished.errors).length>0;
    if(finished.status==='cancelled'){showMessage('warn','PDF作成を中止しました',finished.message||'未処理の原稿を残して停止しました。',finished.results||finished);setTimeout(hideProgressPanel,2400);}
    else if(finished.status==='failed')showMessage('danger','PDF作成が停止しました',userFriendlyError(finished.message||'PDF作成を完了できませんでした。'),finished.errors||finished);
    else if(asArray(finished.errors).length)showMessage('danger','PDF作成でエラーがあります',`${finished.failed||finished.errors.length}件の原稿をPDF変換できませんでした。`,finished.errors);
    else{const endCount=asArray(finished.results).reduce((n,r)=>n+Number(r?.sheetSync?.insertedAtEndCount||0),0);const suffix=endCount?`\n新しい${endCount}ページを末尾に追加しました。`:'';const actions=[{label:'ページ構成を確認',primary:true,view:'pages',handler:endCount?scrollToHighlightedPages:null}];if(endCount)actions.push({label:'原稿内の順序に整える',view:'pages',handler:sortPagesBySheet});else actions.push({label:'提出用PDFへ',view:'final'});showMessage(msg.kind,msg.title,msg.message+suffix,msg.detail,actions,0);setTimeout(hideProgressPanel,1800);}
    if(hasErrors)updateProgressPanel(finished);
  },false);}finally{renderJobActive=false;activeRenderJobId='';activeRenderCancelRequested=false;activeRenderJobStatus=null;const cancel=$('progress-cancel');if(cancel)cancel.hidden=true;updateRenderTargetUi();}
}

function defaultPdfTargetWorkbookIds() {
  return workbooksForActivePreset()
    .filter(w => String(w.status || '') !== 'missing' && !isLatestPdfWorkbook(w))
    .map(w => w.workbookId)
    .filter(Boolean);
}
function getWorkbookRenderTargetInfo() {
  const checkedIds = checkedWorkbookIdsFromUi();
  if (checkedIds.length) return {mode:'selected', ids:checkedIds, count:checkedIds.length};
  const selectedIds = normalizeWorkbookIdListForActivePreset([...selectedWorkbooks]);
  if (selectedIds.length) return {mode:'selected', ids:selectedIds, count:selectedIds.length};
  const neededIds = defaultPdfTargetWorkbookIds();
  if (neededIds.length) return {mode:'needed', ids:neededIds, count:neededIds.length};
  return {mode:'none', ids:[], count:0};
}
function renderTargetStartMessage(info) {
  const count = Number(info?.count || 0);
  if (info?.mode === 'selected') return `${presetLabel()}で選択した原稿 ${count}件をPDF変換します。`;
  if (info?.mode === 'needed') return `未選択のため、${presetLabel()}で変換PDF作成が必要な${count}件をまとめて処理します。`;
  return '変換PDF作成が必要な原稿はありません。';
}
function updateRenderTargetUi() {
  const hint = $('render-target-hint');
  const btn = $('render-selected-btn');
  if (!hint && !btn) return;
  const info = getWorkbookRenderTargetInfo();
  const activeCount = workbooksForActivePreset().filter(w => String(w.status || '') !== 'missing').length;
  const missingCount = workbooksForActivePreset().filter(w => String(w.status || '') === 'missing').length;
  if (btn) {
    btn.disabled = info.mode === 'none';
    if (info.mode === 'selected') btn.textContent = `選択した${info.count}件の変換PDFを作成`;
    else if (info.mode === 'needed') btn.textContent = `必要な${info.count}件の変換PDFを作成`;
    else btn.textContent = '変換PDFを作成';
  }
  if (hint) {
    hint.classList.remove('target-selected', 'target-needed', 'target-none');
    hint.classList.add(info.mode === 'selected' ? 'target-selected' : (info.mode === 'needed' ? 'target-needed' : 'target-none'));
    if (info.mode === 'selected') hint.textContent = `${presetLabel()}で選択中: ${info.count}件の変換PDFを作成します。`;
    else if (info.mode === 'needed') hint.textContent = `${presetLabel()}で変換PDFの作成が必要な原稿は${info.count}件です。ボタンを押すと必要な分だけ作成します。`;
    // 見つからない原稿を数えずに「すべて最新」と言うと、赤い「ファイルなし」が
    // 並んでいる横で矛盾した案内が出る。件数を先に伝える。
    else if (missingCount) hint.textContent = `${missingCount}件の原稿はファイルが見つかりません。付け替えるか、登録を解除してください。残りの変換PDFは最新です。`;
    else hint.textContent = activeCount ? 'すべての変換PDFは最新です。' : `${presetLabel()}の登録済み原稿はありません。`;
  }
}
async function renderSelectedWorkbooks(btn) {
  const info = getWorkbookRenderTargetInfo();
  if (info.mode === 'selected') {
    selectedWorkbooks = new Set(info.ids.map(id => String(id || '')).filter(Boolean));
    syncWorkbookSelectionUi();
    await renderWorkbookIds(info.ids, btn, {targetMode:'selected', startMessage:renderTargetStartMessage(info)});
    return;
  }
  if (info.mode === 'needed') {
    // Do not begin with a full update scan when the user presses PDF作成.
    // The scan can make the progress bar sit at 0% before Excel rendering starts.
    // Use the current screen state and start rendering the needed workbooks immediately.
    await renderWorkbookIds(info.ids, btn, {targetMode:'needed', startMessage:renderTargetStartMessage(info)});
    return;
  }
  hideProgressPanel();
  updateRenderTargetUi();
  showMessage('ok', '変換PDF作成が必要な原稿はありません', '登録済み原稿の変換PDFは最新です。');
}
async function renderUpdated(btn) {
  const neededIds = defaultPdfTargetWorkbookIds();
  if (!neededIds.length) {
    hideProgressPanel();
    showMessage('ok', '変換PDF作成が必要な原稿はありません', '登録済み原稿の変換PDFは最新です。');
    return;
  }
  const info = {mode:'needed', count:neededIds.length};
  await renderWorkbookIds(neededIds, btn, {targetMode:'needed', startMessage:renderTargetStartMessage(info)});
}
async function renderAllWorkbooks(btn) {
  const ids = workbooksForActivePreset().map(w => w.workbookId);
  if (!ids.length) { showMessage('warn','登録済み原稿がありません','先に原稿を登録してください。'); return; }
  // Every other destructive action here confirms first. This one re-runs Office
  // for every file and knocks all finished 提出用PDF back to "要再出力".
  const accepted = await confirmAction({
    title:`登録済み原稿 ${ids.length}件をすべて作り直しますか？`,
    message:'更新のない原稿も含めて、変換PDFをすべて作り直します。',
    detail:'完了までに時間がかかり、作成済みの提出用PDFはすべて「再出力が必要」になります。更新された原稿だけでよい場合は「更新確認」を使ってください。',
    confirmLabel:'すべて作り直す',
    danger:true
  });
  if (!accepted) return;
  await renderWorkbookIds(ids, btn, {targetMode:'all', startMessage:`登録済み原稿 ${ids.length}件の変換PDFを作成します。`});
}
async function renderSelectedPages(btn) {
  const pageIds = [...selectedPages];
  if (!pageIds.length) { showMessage('warn', 'ページを選択してください', 'PDF作成するページにチェックを入れてください。'); return; }
  const ids = [...new Set(pageIds.map(pid => getPage(pid)?.workbookId).filter(Boolean))];
  await renderWorkbookIds(ids, btn);
}


function pageSort(a,b) {
  const ao=Number(a.order),bo=Number(b.order);
  if(Number.isFinite(ao)&&Number.isFinite(bo)&&ao!==bo)return ao-bo;
  const aw=getWorkbook(a.workbookId),bw=getWorkbook(b.workbookId);
  const af=fileOrderValue(aw?.fileName||aw?.displayName||a.workbookId),bf=fileOrderValue(bw?.fileName||bw?.displayName||b.workbookId);
  if(af!==bf)return af-bf;
  const fileCompare=String(aw?.fileName||'').localeCompare(String(bw?.fileName||''),'ja',{numeric:true,sensitivity:'base'});
  if(fileCompare)return fileCompare;
  const ai=Number(a.sheetIndex),bi=Number(b.sheetIndex);
  if(Number.isFinite(ai)&&Number.isFinite(bi)&&ai!==bi)return ai-bi;
  const sheetCompare=String(a.sheetName||'').localeCompare(String(b.sheetName||''),'ja',{numeric:true,sensitivity:'base'});
  if(sheetCompare)return sheetCompare;
  return resolvedPageId(a).localeCompare(resolvedPageId(b));
}
function numberingSelectValue(page) { return page.numberingManual ? (page.numberingMode || 'visible') : 'auto'; }
function pageRangeText(page) {
  const range=page?.pageRange;if(!range)return'';const start=Number(range.start||0),end=Number(range.end||start);if(start<1)return'';return start===end?String(start):`${start}-${end}`;
}
function numberingText(page, indexInVolume) {
  if (!page.numberingManual) return indexInVolume === 0 ? '自動（表示なし）' : '自動（表示）';
  return page.numberingMode === 'none' ? '表示なし' : '表示';
}
function setHighlightedPages(ids) {
  highlightedPageIds = new Set(asArray(ids).map(String).filter(Boolean));
  if (pageHighlightTimer) { clearTimeout(pageHighlightTimer); pageHighlightTimer = null; }
  if (!highlightedPageIds.size) return;
  lastPageBoardRenderSignature = '';
  pageHighlightTimer = setTimeout(() => {
    highlightedPageIds.clear();
    document.querySelectorAll('.new-page-row').forEach(row => row.classList.remove('new-page-row'));
  }, 12000);
}
function scrollToHighlightedPages() {
  const first = document.querySelector('.new-page-row');
  if (first) { first.scrollIntoView({block:'center',behavior:'smooth'}); first.focus({preventScroll:true}); }
}

function pageMatchesSearch(page){
  const query=String(pageFilterText||'').trim().toLocaleLowerCase('ja');if(!query)return true;
  const wb=getWorkbook(page?.workbookId);const haystack=[page?.title,page?.sheetName,wb?.displayName,wb?.fileName,wb?.relativePath].map(value=>String(value||'')).join(' ').toLocaleLowerCase('ja');
  return haystack.includes(query);
}
function applyPageFilters(pages){
  return pages.filter(page=>(!showChangedOnly||!state?.inputHistoryEnabled||pageIsChanged(page))&&pageMatchesSearch(page));
}

function renderPageTargetControls(){
  const targets=activeTargetDefinitions(),bulk=$('bulk-target-buttons'),preview=$('preview-target-buttons');
  if(bulk){bulk.innerHTML=targets.map(target=>`<button class="btn secondary" type="button" data-bulk-target="${escapeAttr(targetVolume(target.targetId))}" disabled>${escapeHtml(target.displayName||target.targetId)}へ入れる</button>`).join('');bulk.querySelectorAll('[data-bulk-target]').forEach(button=>button.addEventListener('click',()=>moveSelectedPagesToVolume(button.dataset.bulkTarget,button)));}
  if(preview){preview.innerHTML=targets.map(target=>`<button class="btn secondary compact" type="button" data-preview-target="${escapeAttr(targetVolume(target.targetId))}">${escapeHtml(target.displayName||target.targetId)}へ</button>`).join('');preview.querySelectorAll('[data-preview-target]').forEach(button=>button.addEventListener('click',()=>moveCurrentPreviewPageToVolume(button.dataset.previewTarget,button)));}
}

function renderPages() {
  const box=$('page-board');if(!box)return;renderPageTargetControls();const sourcePages=[...pagesForActivePreset()].sort(pageSort);const visiblePages=applyPageFilters(sourcePages);const visibleIds=new Set(visiblePages.map(resolvedPageId));const allPages=sourcePages;
  const needsPdf=workbooksForActivePreset().filter(w=>!isLatestPdfWorkbook(w));const agg=aggregateFinalState();
  const signature=JSON.stringify([pageBoardView,activePageVolume,pageFilterText,showChangedOnly,[...pageVolumeCollapseOverrides.entries()],allPages.map(p=>[resolvedPageId(p),p.order,p.volume,p.enabled,p.title,pageRangeText(p),p.numberingMode,p.numberingManual,p.orderManual,p.status,p.contentPdf,p.sheetIndex]),needsPdf.map(w=>w.workbookId),agg.state]);
  if(lastPageBoardRenderSignature===signature&&box.childElementCount)return;lastPageBoardRenderSignature=signature;
  const scrollSnap=capturePageBoardScroll();selectedPages=new Set([...selectedPages].filter(id=>allPages.some(p=>resolvedPageId(p)===id)));
  const banners=[];const unassigned=allPages.filter(p=>String(p.volume)==='none'||p.enabled===false).length;
  if(unassigned)banners.push(`<div class="page-status-banner attention"><strong>未振り分けのページが ${unassigned}ページあります</strong><span>出力先へ移したページだけが提出用PDFに含まれます。</span></div>`);
  if(needsPdf.length)banners.push(`<div class="page-status-banner"><strong>変換PDF作成が必要な原稿があります</strong><span>${needsPdf.length}件。先に変換PDFを作成してください。</span></div>`);
  if(agg.state==='needs-rebuild')banners.push('<div class="page-status-banner"><strong>ページ構成または変換PDFが変更されています</strong><span>提出用PDFの再出力が必要です。</span></div>');
  const panels=[{volume:'none',title:'未振り分け（出力しない）',description:'新規ページはここに入ります。必要なページを選び、出力先へ移してください。'},...activeTargetDefinitions().map(target=>({volume:targetVolume(target.targetId),title:String(target.displayName||target.targetId),description:`ここへ入れたページが「${target.displayName||target.targetId}」PDFに含まれます。`}))];
  if(!panels.some(panel=>panel.volume===activePageVolume))activePageVolume=panels[0]?.volume||'none';
  if(!pageVolumeAutoPicked&&allPages.length){
    const pageCountFor=volume=>allPages.filter(page=>volume==='none'?(String(page.volume)==='none'||page.enabled===false):(String(page.volume||activeTargetVolumes()[0]||mainVolume())===volume&&page.enabled!==false)).length;
    if(!pageCountFor(activePageVolume)){
      const populated=panels.find(panel=>pageCountFor(panel.volume)>0);
      if(populated)activePageVolume=populated.volume;
    }
    pageVolumeAutoPicked=true;
  }
  const tabs=$('page-volume-tabs');if(tabs){tabs.innerHTML=panels.map(panel=>{const count=allPages.filter(page=>panel.volume==='none'?(String(page.volume)==='none'||page.enabled===false):(String(page.volume||activeTargetVolumes()[0]||mainVolume())===panel.volume&&page.enabled!==false)).length;const selected=panel.volume===activePageVolume,target=panel.volume==='none'?null:activeTargetDefinitions().find(item=>targetVolume(item.targetId)===panel.volume),optionalEmpty=target?.required===false&&count===0&&!selected,label=panel.volume==='none'?'未振り分け':optionalEmpty?`＋ ${panel.title}を使う`:panel.title;return `<button class="page-volume-tab ${selected?'active':''} ${optionalEmpty?'optional-empty':''}" type="button" aria-pressed="${selected}" data-page-volume-tab="${escapeAttr(panel.volume)}"><span>${escapeHtml(label)}</span>${optionalEmpty?'':`<b>${count}</b>`}</button>`;}).join('');tabs.querySelectorAll('[data-page-volume-tab]').forEach(button=>button.addEventListener('click',()=>{activePageVolume=String(button.dataset.pageVolumeTab||'none');try{sessionStorage.setItem('ReportBinderActivePageVolume',activePageVolume);}catch{}lastPageBoardRenderSignature='';selectedPages.clear();renderPages();}));}
  if(sourcePages.length>0&&visiblePages.length===0){const searching=!!String(pageFilterText||'').trim();const title=searching?'一致するページがありません':'変更されたページはありません';const hint=searching?'検索語を短くするか、クリアしてください。':'「変更分のみ」を解除すると、すべてのページを表示できます。';banners.push(`<div class="page-filter-empty"><strong>${title}</strong><span>${hint}</span></div>`);}
  box.className=`board ${pageBoardView==='thumbnail'?'thumbnail-board':'detail-board'}`;box.innerHTML=banners.join('')+panels.map(panel=>dynamicVolumePanelHtml(panel,allPages,visibleIds)).join('');attachBoardEvents();preparePageThumbnails();restorePageBoardScroll(scrollSnap);syncPageBoardViewControls();syncPageSelectionUi();
}
function dynamicVolumePanelHtml(panel,allPages,visibleIds){
  const volume=panel.volume,first=activeTargetVolumes()[0]||mainVolume(),pages=allPages.filter(page=>volume==='none'?(String(page.volume)==='none'||page.enabled===false):(String(page.volume||first)===volume&&page.enabled!==false)),visibleCount=pages.filter(page=>visibleIds.has(resolvedPageId(page))).length,collapsed=pageVolumeCollapseOverrides.has(volume)?!!pageVolumeCollapseOverrides.get(volume):(pageBoardView==='detail'&&pages.length===0),filtering=(showChangedOnly&&state?.inputHistoryEnabled)||!!String(pageFilterText||'').trim(),countText=filtering?`${visibleCount} / ${pages.length}ページ`:`${pages.length}ページ`;
  const content=pageBoardView==='thumbnail'?`<div class="thumbnail-wrap volume-content"><div class="thumbnail-grid" data-volume="${escapeAttr(volume)}">${pages.map((page,index)=>pageThumbnailHtml(page,index,!visibleIds.has(resolvedPageId(page)))).join('')||'<div class="empty-row page-empty-drop">ここへドロップ</div>'}</div></div>`:`<div class="table-wrap volume-content"><table class="page-table"><thead><tr><th class="check-col"><input type="checkbox" data-select-all-pages="${escapeAttr(volume)}" aria-label="${escapeAttr(panel.title)}をすべて選択"></th><th class="seq-col">順</th><th class="drag-col">移動</th><th>ページ名</th><th>PDF</th><th>番号</th></tr></thead><tbody data-volume="${escapeAttr(volume)}">${pages.map((page,index)=>pageRowHtml(page,index,!visibleIds.has(resolvedPageId(page)))).join('')||'<tr class="empty-row"><td colspan="6">ここへドロップ</td></tr>'}</tbody></table></div>`;
  return `<section class="volume-panel ${volume==='none'?'inbox':''} ${volume===activePageVolume?'active-volume':'inactive-volume'} ${collapsed?'collapsed':''}" data-volume-panel="${escapeAttr(volume)}"><button class="volume-head" type="button" data-toggle-volume="${escapeAttr(volume)}" aria-expanded="${collapsed?'false':'true'}"><div><h3>${escapeHtml(panel.title)}</h3><p>${escapeHtml(panel.description)}</p></div><span class="volume-head-meta"><b>${escapeHtml(countText)}</b><svg class="icon"><use href="#i-chevron-down"/></svg></span></button>${content}</section>`;
}
// 絞り込みで隠れるページは、中身を作らず data-page-id だけの器にする。
// 器を残すのは collectBoardVolumes() が並び順をDOMから絶対値で読み取るため。
// 完全に消すと、絞り込み中のドラッグが隠れたページを欠いた並びを送ってしまう。
function filteredOutPageHtml(pid,isRow){
  return isRow
    ? `<tr class="page-row page-filter-hidden" data-page-id="${escapeAttr(pid)}" aria-hidden="true"></tr>`
    : `<article class="page-row page-filter-hidden" data-page-id="${escapeAttr(pid)}" aria-hidden="true"></article>`;
}
function pageRowHtml(p, idx, filteredOut=false) {
  if(filteredOut)return filteredOutPageHtml(resolvedPageId(p),true);
  const wb=getWorkbook(p.workbookId),pid=resolvedPageId(p),hasPdf=pagePreviewAvailable(p);const warnings=asArray(p.warnings).filter(w=>!/縮尺例外|Zoom|倍率例外/.test(String(w))).map(userFriendlyError);const checked=selectedPages.has(pid),highlighted=highlightedPageIds.has(pid);
  // Table headers do not name form controls, so without these every row read
  // out as an unlabelled checkbox / textbox / combobox in a list of hundreds.
  const rowLabel=`${idx+1}ページ目 ${p.title||p.sheetName||'ページ'}`;
  return `<tr tabindex="0" class="page-row ${filteredOut?'page-filter-hidden':''} ${checked?'selected-row':''} ${highlighted?'new-page-row':''} ${p.status==='render-error'?'error-row':''}" data-page-id="${escapeAttr(pid)}" data-workbook-id="${escapeAttr(p.workbookId||'')}" data-sheet-name="${escapeAttr(p.sheetName||'')}" data-content-pdf="${escapeAttr(p.contentPdf||'')}"><td class="check-col"><input type="checkbox" data-page-check value="${escapeAttr(pid)}" aria-label="${escapeAttr(`${rowLabel}を選択`)}" ${checked?'checked':''}></td><td class="seq-col"><span class="seq-badge" data-seq-cell>${idx+1}</span></td><td class="drag-col"><span class="drag-handle" title="ドラッグして移動">${iconUse('i-drag')}</span></td><td class="page-main-cell"><input class="title-input" data-page-title aria-label="${escapeAttr(`${rowLabel}のページ名`)}" value="${escapeAttr(p.title||'')}"><button class="file-link btn ghost" type="button" ${hasPdf?`data-preview-page="${escapeAttr(pid)}"`:''}>${escapeHtml(wb?.displayName||wb?.fileName||'')} / ${escapeHtml(sourceUnitReference(wb,p.sheetName))}</button>${warnings.length?`<div class="page-warning">${warnings.map(escapeHtml).join('<br>')}</div>`:''}</td><td class="${hasPdf?'preview-trigger':''}" ${hasPdf?`data-preview-page="${escapeAttr(pid)}"`:''}>${pdfStatusForPage(p)} ${pageChangeBadge(p)}<div class="subtext">${escapeHtml(pagePdfSubtext(p))}</div></td><td><select data-page-numbering aria-label="${escapeAttr(`${rowLabel}のページ番号表示`)}"><option value="auto" ${numberingSelectValue(p)==='auto'?'selected':''}>自動</option><option value="none" ${numberingSelectValue(p)==='none'?'selected':''}>表示なし</option><option value="visible" ${numberingSelectValue(p)==='visible'?'selected':''}>表示</option></select><div class="subtext">${escapeHtml(numberingText(p,idx))}</div><label class="page-range-inline">使用ページ<input data-page-range aria-label="${escapeAttr(`${rowLabel}で使用するページ範囲`)}" value="${escapeAttr(pageRangeText(p))}" placeholder="すべて"></label></td></tr>`;
}
function pageThumbnailHtml(p,idx,filteredOut=false){
  if(filteredOut)return filteredOutPageHtml(resolvedPageId(p),false);
  const wb=getWorkbook(p.workbookId),pid=resolvedPageId(p),hasPdf=pagePreviewAvailable(p);const warnings=asArray(p.warnings).filter(w=>!/縮尺例外|Zoom|倍率例外/.test(String(w))).map(userFriendlyError);const checked=selectedPages.has(pid),highlighted=highlightedPageIds.has(pid);const source=`${wb?.displayName||wb?.fileName||''} / ${sourceUnitReference(wb,p.sheetName)}`,numbering=numberingSelectValue(p),range=pageRangeText(p),displayTitle=p.title||p.sheetName||'ページ',editorId=`page-thumb-editor-${pid}`,change=pageChangeBadge(p);return `<article tabindex="0" class="page-row page-thumb-card ${filteredOut?'page-filter-hidden':''} ${checked?'selected-row':''} ${highlighted?'new-page-row':''} ${p.status==='render-error'?'error-row':''}" data-page-id="${escapeAttr(pid)}" data-workbook-id="${escapeAttr(p.workbookId||'')}" data-sheet-name="${escapeAttr(p.sheetName||'')}" data-content-pdf="${escapeAttr(p.contentPdf||'')}" aria-label="${escapeAttr(`${idx+1}ページ目 ${displayTitle}。クリックで選択`)}" title="クリックで選択"><input class="sr-only" type="checkbox" data-page-check value="${escapeAttr(pid)}" aria-label="${escapeAttr(`${displayTitle}を選択`)}" ${checked?'checked':''}><div class="page-thumb-paper"><canvas data-page-thumbnail="${escapeAttr(pid)}" aria-hidden="true"></canvas><div class="page-thumb-placeholder">${hasPdf?'プレビューを読み込み中':'PDF未作成'}</div><span class="page-thumb-seq" data-seq-cell>${idx+1}</span></div><div class="page-thumb-copy"><strong>${escapeHtml(displayTitle)}</strong><span>${escapeHtml(source)}</span>${warnings.length?`<div class="page-warning">${warnings.map(escapeHtml).join('<br>')}</div>`:''}</div><div class="page-thumb-meta"><span>${change||(!hasPdf?pdfStatusForPage(p):'')}</span><span class="page-thumb-settings-summary">${hasPdf?`<button class="btn ghost compact page-thumb-preview" type="button" data-preview-page="${escapeAttr(pid)}">プレビュー</button>`:''}<button class="btn ghost compact page-thumb-edit" type="button" data-thumb-edit aria-label="${escapeAttr(`${displayTitle}の設定を編集`)}" aria-controls="${escapeAttr(editorId)}" aria-expanded="false">${iconUse('i-edit')}設定</button></span></div><div id="${escapeAttr(editorId)}" class="page-thumb-editor" data-thumb-editor hidden><label>ページ名<input data-thumb-page-title value="${escapeAttr(p.title||'')}" data-original-value="${escapeAttr(p.title||'')}"></label><label>元PDFから使う範囲（複数ページの場合）<input data-thumb-page-range value="${escapeAttr(range)}" data-original-value="${escapeAttr(range)}" placeholder="すべて（例: 2-5）" inputmode="numeric"></label><label>ページ番号<select data-thumb-page-numbering data-original-value="${escapeAttr(numbering)}"><option value="auto" ${numbering==='auto'?'selected':''}>自動</option><option value="none" ${numbering==='none'?'selected':''}>表示なし</option><option value="visible" ${numbering==='visible'?'selected':''}>表示</option></select></label><div class="page-thumb-editor-actions"><button class="btn ghost compact" type="button" data-thumb-cancel>取消</button><button class="btn primary compact" type="button" data-thumb-save>保存</button></div></div></article>`;
}

function pageThumbnailCacheKey(page){return `${resolvedPageId(page)}|${String(page?.contentPdf||'')}|${String(page?.updatedAt||'')}`;}
async function buildPageThumbnail(page){
  const key=pageThumbnailCacheKey(page);if(pageThumbnailCache.has(key)){const cached=pageThumbnailCache.get(key);pageThumbnailCache.delete(key);pageThumbnailCache.set(key,cached);return cached;}
  const promise=(async()=>{const lib=await ensureDiffPdfJs();const loadingTask=lib.getDocument({url:apiUrl('/api/file',{pageId:resolvedPageId(page),workbookId:page?.workbookId||'',sheetName:page?.sheetName||''}),httpHeaders:{'X-ReportBinder-Token':token,'Accept':'application/pdf'},disableRange:false,disableStream:false});let doc=null;try{doc=await loadingTask.promise;const pdfPage=await doc.getPage(1);const base=pdfPage.getViewport({scale:1});const pixelRatio=Math.min(2,Math.max(1,window.devicePixelRatio||1));const targetWidth=220*pixelRatio;const viewport=pdfPage.getViewport({scale:targetWidth/Math.max(1,base.width)});const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.ceil(viewport.width));canvas.height=Math.max(1,Math.ceil(viewport.height));const context=canvas.getContext('2d',{alpha:false});context.fillStyle='#fff';context.fillRect(0,0,canvas.width,canvas.height);await pdfPage.render({canvasContext:context,viewport,background:'rgb(255,255,255)'}).promise;return canvas;}finally{try{await doc?.destroy?.();}catch{}}})();
  pageThumbnailCache.set(key,promise);while(pageThumbnailCache.size>PAGE_THUMBNAIL_CACHE_LIMIT)pageThumbnailCache.delete(pageThumbnailCache.keys().next().value);try{return await promise;}catch(error){pageThumbnailCache.delete(key);throw error;}
}
function pumpPageThumbnailQueue(){
  while(pageThumbnailRenderActive<PAGE_THUMBNAIL_RENDER_LIMIT&&pageThumbnailRenderQueue.length){const canvas=pageThumbnailRenderQueue.shift();if(!canvas?.isConnected||canvas.dataset.thumbnailState==='loading'||canvas.closest('.page-filter-hidden'))continue;const page=getPage(canvas.dataset.pageThumbnail);if(!page||!pagePreviewAvailable(page))continue;canvas.dataset.thumbnailState='loading';pageThumbnailRenderActive++;buildPageThumbnail(page).then(source=>{if(!canvas.isConnected)return;canvas.width=source.width;canvas.height=source.height;canvas.getContext('2d',{alpha:false}).drawImage(source,0,0);canvas.dataset.thumbnailState='ready';const placeholder=canvas.parentElement?.querySelector('.page-thumb-placeholder');if(placeholder)placeholder.classList.add('hidden');}).catch(()=>{if(!canvas.isConnected)return;canvas.dataset.thumbnailState='error';const placeholder=canvas.parentElement?.querySelector('.page-thumb-placeholder');if(placeholder)placeholder.textContent='プレビューできません';}).finally(()=>{pageThumbnailRenderActive--;pumpPageThumbnailQueue();});}
}
function queuePageThumbnail(canvas){if(!canvas||canvas.dataset.thumbnailState||pageThumbnailRenderQueue.includes(canvas))return;pageThumbnailRenderQueue.push(canvas);pumpPageThumbnailQueue();}
// 画面外へ出たサムネイルの canvas を解放する。A4縦は 440x622x4B ≒ 1.09MB あり、
// 174ページを一度スクロールし切ると描画済みcanvasだけで100MBを超える。
// 解放しても data-page-id は DOM に残るので、並べ替えの絶対位置は変わらない。
function releasePageThumbnailCanvas(canvas){
  // 描画中(loading)を捨てると描き直しが二重に走る。完了済みだけ解放する。
  if(!canvas||canvas.dataset.thumbnailState!=='ready')return;
  canvas.width=0;canvas.height=0;
  delete canvas.dataset.thumbnailState;
  const placeholder=canvas.parentElement?.querySelector('.page-thumb-placeholder');
  if(placeholder)placeholder.classList.remove('hidden');
}
function preparePageThumbnails(){
  if(pageThumbnailObserver){pageThumbnailObserver.disconnect();pageThumbnailObserver=null;}pageThumbnailRenderQueue.length=0;if(pageBoardView!=='thumbnail')return;
  const canvases=[...document.querySelectorAll('[data-page-thumbnail]')];
  if(!('IntersectionObserver'in window)){canvases.forEach(queuePageThumbnail);return;}
  pageThumbnailObserver=new IntersectionObserver(entries=>{
    for(const entry of entries){
      if(entry.isIntersecting)queuePageThumbnail(entry.target);
      else releasePageThumbnailCanvas(entry.target);
    }
  },{rootMargin:'240px 0px'});
  // unobserve せず観測を続ける。離脱時に解放し、再入場で描き直す。
  canvases.forEach(canvas=>{if(!canvas.closest('.page-filter-hidden'))pageThumbnailObserver.observe(canvas);});
}

function clearDropHighlights() {
  document.querySelectorAll('.drop-active').forEach(el => el.classList.remove('drop-active'));
  document.querySelectorAll('.drop-placeholder').forEach(el => el.remove());
  document.querySelectorAll('.drop-target-tab').forEach(el => el.classList.remove('drop-target-tab'));
}
// Only one volume panel is rendered at a time (.inactive-volume{display:none}),
// so there is no in-board drop target for the other volumes. The tabs stand in
// for them: drop a drag on a tab to send the pages to that volume.
function findVolumeTabAt(x, y) {
  for (const el of document.elementsFromPoint(x, y)) {
    const tab = el.closest?.('[data-page-volume-tab]');
    if (tab) return tab;
  }
  return null;
}
function volumeOfRow(row) {
  return String(row?.closest('[data-volume]')?.getAttribute('data-volume') || '');
}
function highlightVolumeTabTarget(tab) {
  document.querySelectorAll('.drop-target-tab').forEach(el => { if (el !== tab) el.classList.remove('drop-target-tab'); });
  tab?.classList.add('drop-target-tab');
}
function renumberBoardRows() {
  document.querySelectorAll('[data-volume]').forEach(tbody => {
    let n = 1;
    [...tbody.querySelectorAll('.page-row,.drop-placeholder')].forEach(row => {
      if (row.classList.contains('empty-row')) return;
      const seq = row.querySelector('[data-seq-cell], [data-placeholder-seq]');
      if (seq) seq.textContent = String(n);
      if(row.classList.contains('page-thumb-card')){const title=row.querySelector('.page-thumb-copy strong')?.textContent?.trim()||'ページ';row.setAttribute('aria-label',`${n}ページ目 ${title}`);}
      n++;
    });
  });
}
function getRowsForDrag(sourceRow) {
  const sourceId = sourceRow?.getAttribute('data-page-id') || '';
  if (sourceId && selectedPages.has(sourceId) && selectedPages.size > 1) {
    const selectedRows = [...document.querySelectorAll('.page-row:not(.page-filter-hidden)')].filter(r => selectedPages.has(r.getAttribute('data-page-id')));
    if (selectedRows.some(r => r === sourceRow)) return selectedRows;
  }
  return sourceRow ? [sourceRow] : [];
}
function currentBoardSignature() { return JSON.stringify(collectBoardVolumes()); }
function dragRowIds(drag=draggingRow) {
  return asArray(drag?.rows).map(r => String(r.getAttribute('data-page-id') || '')).filter(Boolean);
}
function previewBoardVolumesForDrop(tbody, afterRow) {
  const drag = draggingRow;
  const movingIds = dragRowIds(drag);
  const moving = new Set(movingIds);
  const volumes = {};
  document.querySelectorAll('[data-volume]').forEach(tb => {
    volumes[tb.getAttribute('data-volume')] = [...tb.querySelectorAll('.page-row')]
      .map(row => String(row.getAttribute('data-page-id') || ''))
      .filter(id => id && !moving.has(id));
  });
  const volumeName = tbody?.getAttribute('data-volume') || '';
  if (!Object.prototype.hasOwnProperty.call(volumes, volumeName)) volumes[volumeName] = [];
  const afterId = afterRow ? String(afterRow.getAttribute('data-page-id') || '') : '';
  let insertAt = afterId ? volumes[volumeName].indexOf(afterId) : volumes[volumeName].length;
  if (insertAt < 0) insertAt = volumes[volumeName].length;
  volumes[volumeName].splice(insertAt, 0, ...movingIds);
  return volumes;
}
function createDropPlaceholder(sourceRows) {
  const rows = asArray(sourceRows);
  const title = rows.length > 1
    ? `${rows.length}ページをまとめて移動`
    : (rows[0]?.querySelector('.title-input')?.value || rows[0]?.querySelector('.page-thumb-copy strong')?.textContent || rows[0]?.querySelector('.page-main-cell .subtext')?.textContent || 'このページ');
  const thumbnail = rows[0]?.classList?.contains('page-thumb-card');
  const ph = document.createElement(thumbnail?'div':'tr');
  ph.className = `drop-placeholder ${thumbnail?'thumbnail-drop-placeholder':''}`;
  ph.innerHTML = thumbnail?`<div class="placeholder-card">ここへ移動<br><strong>${escapeHtml(title)}</strong></div>`:`<td colspan="6"><div class="placeholder-card">ここへ移動：${escapeHtml(title)}</div></td>`;
  return ph;
}
function createPageDragGhost(sourceRows) {
  const rows=asArray(sourceRows),first=rows[0];
  const title=first?.querySelector('.page-thumb-copy strong')?.textContent?.trim()||first?.querySelector('.title-input')?.value||'ページ';
  const ghost=document.createElement('div');ghost.className='page-drag-ghost';ghost.setAttribute('aria-hidden','true');
  ghost.innerHTML=rows.length>1?`<strong>${rows.length}ページ</strong><span>${escapeHtml(title)} ほか</span>`:`<strong>${escapeHtml(title)}</strong><span>移動先へドロップ</span>`;
  document.body.appendChild(ghost);return ghost;
}
function positionPageDragGhost(ghost,x,y){
  if(!ghost)return;const margin=12,width=ghost.offsetWidth||210,height=ghost.offsetHeight||58;const left=Math.max(margin,Math.min(window.innerWidth-width-margin,x+18));const top=Math.max(margin,Math.min(window.innerHeight-height-margin,y+18));ghost.style.transform=`translate3d(${left}px,${top}px,0)`;
}
function removeEmptyRow(tbody) {
  const empty = tbody?.querySelector?.('.empty-row');
  if (empty) empty.remove();
}
function findDropTbodyAt(x, y) {
  const elements = document.elementsFromPoint(x, y);
  for (const el of elements) {
    const tbody = el.closest?.('[data-volume]');
    if (tbody) return tbody;
    const panel = el.closest?.('.volume-panel');
    if (panel) {
      panel.classList.remove('collapsed');
      const toggle=panel.querySelector('[data-toggle-volume]');if(toggle)toggle.setAttribute('aria-expanded','true');
      const volume=panel.getAttribute('data-volume-panel')||'';if(volume)pageVolumeCollapseOverrides.set(volume,false);
      const candidate = panel.querySelector('[data-volume]');
      if (candidate) return candidate;
    }
  }
  // Fallback for cases where the pointer is over a sticky header or a scrollbar.
  let best = null;
  let bestDistance = Infinity;
  document.querySelectorAll('[data-volume]').forEach(tbody => {
    const panel = tbody.closest('.volume-panel');
    const rect = panel?.getBoundingClientRect?.();
    if (!rect) return;
    const inX = x >= rect.left && x <= rect.right;
    if (!inX) return;
    const dy = y < rect.top ? rect.top - y : (y > rect.bottom ? y - rect.bottom : 0);
    if (dy < bestDistance) { bestDistance = dy; best = tbody; }
  });
  return best;
}
function autoScrollDuringDragPointer(e, tbody) {
  const margin = 86;
  const speed = 22;
  if (e.clientY > window.innerHeight - margin) window.scrollBy(0, speed);
  else if (e.clientY < margin) window.scrollBy(0, -speed);
  const wrap = tbody?.closest?.('.table-wrap,.thumbnail-wrap');
  if (wrap) {
    const rect = wrap.getBoundingClientRect();
    if (e.clientY > rect.bottom - 64) wrap.scrollTop += 24;
    else if (e.clientY < rect.top + 64) wrap.scrollTop -= 24;
  }
}
function getDragAfterRow(tbody, x, y) {
  const rows = [...tbody.querySelectorAll('.page-row:not(.dragging):not(.page-filter-hidden)')];
  if(tbody.classList.contains('thumbnail-grid')){
    for(const child of rows){const box=child.getBoundingClientRect();const centerY=box.top+box.height/2;if(y<centerY){if(y<box.top-8||x<box.left+box.width/2)return child;}}
    return null;
  }
  return rows.reduce((closest, child) => {
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > closest.offset) return {offset, element: child};
    return closest;
  }, {offset: Number.NEGATIVE_INFINITY, element: null}).element;
}
function moveDropPlaceholder(tbody, x, y) {
  if (!draggingRow || !draggingRow.placeholder || !tbody) return;
  const ph = draggingRow.placeholder;
  const after = getDragAfterRow(tbody, x, y);
  const previewVolumes = previewBoardVolumesForDrop(tbody, after);
  const previewSignature = JSON.stringify(previewVolumes);
  const samePlace = previewSignature === draggingRow.originalSignature;
  draggingRow.previewSignature = previewSignature;
  draggingRow.isNoop = samePlace;

  document.querySelectorAll('.volume-panel.drop-active').forEach(p => p.classList.remove('drop-active'));
  if (samePlace) {
    // Same-position hover is not a real move. Do not show "ここへ移動" there,
    // because it makes users think dropping will change the order.
    if (ph.parentElement) ph.remove();
    renumberBoardRows();
    return;
  }

  removeEmptyRow(tbody);
  tbody.closest('.volume-panel')?.classList.add('drop-active');
  if (after) tbody.insertBefore(ph, after);
  else tbody.appendChild(ph);
  renumberBoardRows();
}
function finishPointerDrag(commit) {
  const drag = draggingRow;
  if (!drag) return;
  const ph = drag.placeholder;
  const beforeSignature = drag.originalSignature || currentBoardSignature();
  let moved = false;
  if (commit && !drag.isNoop && ph && ph.parentElement) {
    for (const row of drag.rows) ph.parentElement.insertBefore(row, ph);
    moved = true;
  }
  for (const row of drag.rows) row.classList.remove('dragging');
  drag.ghost?.remove();
  draggingRow = null;
  document.body.classList.remove('is-dragging-page');
  clearDropHighlights();
  renumberBoardRows();
  const afterSignature = currentBoardSignature();
  let beforeVolumes=drag.originalVolumes,afterVolumes=null;try{beforeVolumes=JSON.parse(beforeSignature);}catch{}try{afterVolumes=JSON.parse(drag.previewSignature||afterSignature);}catch{}
  setTimeout(()=>{if(drag.sourceRow?.dataset)delete drag.sourceRow.dataset.suppressClick;},0);
  if (commit && moved && afterSignature !== beforeSignature) scheduleBoardSave({label:`${drag.rows.length}ページの並べ替え`,volumes:beforeVolumes,afterVolumes});
}
function beginPointerPageDrag(e, row) {
  if (!row || e.button !== 0) return;
  const fromHandle=!!e.target.closest('.drag-handle');
  const directCardDrag=row.classList.contains('page-thumb-card')&&!e.target.closest('input, select, option, button, a, .preview-trigger, .page-thumb-check, .page-thumb-editor');
  if (!fromHandle&&!directCardDrag) return;
  if (fromHandle) e.preventDefault();
  const captureTarget = e.currentTarget || row;
  const startX = e.clientX;
  const startY = e.clientY;
  let started = false;
  const startDrag = (ev) => {
    if (started) return;
    started = true;
    ev.preventDefault();row.dataset.suppressClick='true';
    const rows = getRowsForDrag(row);
    const ghost=createPageDragGhost(rows);positionPageDragGhost(ghost,ev.clientX,ev.clientY);
    draggingRow = {sourceRow: row, rows, placeholder: createDropPlaceholder(rows), ghost, originalSignature: currentBoardSignature(), originalVolumes:collectBoardVolumes(), previewSignature: '', isNoop: true};
    for (const r of rows) r.classList.add('dragging');
    document.body.classList.add('is-dragging-page');
    const tbody = findDropTbodyAt(ev.clientX, ev.clientY) || row.closest('[data-volume]');
    if (tbody) moveDropPlaceholder(tbody, ev.clientX, ev.clientY);
  };
  try { captureTarget.setPointerCapture?.(e.pointerId); } catch {}
  const cleanup = (ev) => {
    try { captureTarget.releasePointerCapture?.(ev.pointerId); } catch {}
    // 保留中のフレームを残すと、ドロップ後の確定済みDOMをもう一度動かしてしまう。
    if (moveFrame) { cancelAnimationFrame(moveFrame); moveFrame = 0; }
    pendingMove = null;
    document.removeEventListener('pointermove', onMove, true);
    document.removeEventListener('pointerup', onUp, true);
    document.removeEventListener('pointercancel', onCancel, true);
  };
  // pointermove は毎秒60回来る。1回ごとに盤面を全走査して getBoundingClientRect と
  // DOM書き込みを交互に行うと、毎イベントで全文書レイアウトが強制され、155ページの
  // 詳細表示では16.7msのフレーム予算を確実に超える。rAFで1フレーム1回に束ねる。
  let pendingMove = null;
  let moveFrame = 0;
  const applyMove = () => {
    moveFrame = 0;
    const ev = pendingMove;
    pendingMove = null;
    if (!ev || !draggingRow) return;
    positionPageDragGhost(draggingRow.ghost, ev.clientX, ev.clientY);
    const tab = findVolumeTabAt(ev.clientX, ev.clientY);
    draggingRow.tabTarget = tab && volumeOfRow(draggingRow.sourceRow) !== String(tab.dataset.pageVolumeTab || '') ? tab : null;
    highlightVolumeTabTarget(draggingRow.tabTarget);
    if (draggingRow.tabTarget) return;
    const tbody = findDropTbodyAt(ev.clientX, ev.clientY);
    autoScrollDuringDragPointer(ev, tbody);
    if (tbody) moveDropPlaceholder(tbody, ev.clientX, ev.clientY);
  };
  const onMove = (ev) => {
    const moved = Math.abs(ev.clientX - startX) + Math.abs(ev.clientY - startY);
    if (!started && moved < 5) return;
    if (!started) startDrag(ev);
    if (!draggingRow) return;
    ev.preventDefault();
    pendingMove = {clientX: ev.clientX, clientY: ev.clientY, target: ev.target};
    if (!moveFrame) moveFrame = requestAnimationFrame(applyMove);
  };
  const onUp = (ev) => {
    cleanup(ev);
    if (!started) { return; }
    const tab = draggingRow ? findVolumeTabAt(ev.clientX, ev.clientY) : null;
    if (tab) {
      const target = String(tab.dataset.pageVolumeTab || '');
      const ids = draggingRow.rows.map(r => r.getAttribute('data-page-id')).filter(Boolean);
      const sameVolume = volumeOfRow(draggingRow.sourceRow) === target;
      // Cancel first: that restores the board to its pre-drag DOM, which
      // movePageIdsToVolume reads as the undo baseline.
      finishPointerDrag(false);
      if (!sameVolume && ids.length) void movePageIdsToVolume(ids, target, null);
      return;
    }
    const targetTbody = findDropTbodyAt(ev.clientX, ev.clientY);
    const commit = !!(draggingRow && draggingRow.placeholder && targetTbody && draggingRow.placeholder.parentElement && !draggingRow.isNoop);
    finishPointerDrag(commit);
  };
  const onCancel = (ev) => {
    cleanup(ev);
    if (started) finishPointerDrag(false);
  };
  document.addEventListener('pointermove', onMove, true);
  document.addEventListener('pointerup', onUp, true);
  document.addEventListener('pointercancel', onCancel, true);
}

function setPageThumbnailEditor(row,open,reset=false){
  if(!row)return;if(open)document.querySelectorAll('.page-thumb-card.editing').forEach(card=>{if(card!==row)setPageThumbnailEditor(card,false,true);});const editor=row.querySelector('[data-thumb-editor]'),toggle=row.querySelector('[data-thumb-edit]');if(!editor||!toggle)return;row.classList.toggle('editing',!!open);editor.hidden=!open;toggle.setAttribute('aria-expanded',String(!!open));
  if(!open&&reset){const title=editor.querySelector('[data-thumb-page-title]'),range=editor.querySelector('[data-thumb-page-range]'),numbering=editor.querySelector('[data-thumb-page-numbering]');if(title)title.value=String(title.dataset.originalValue||'');if(range)range.value=String(range.dataset.originalValue||'');if(numbering)numbering.value=String(numbering.dataset.originalValue||'auto');}
  if(open){const title=editor.querySelector('[data-thumb-page-title]');title?.focus();title?.select();}
}

function visiblePageRowsForKeyboard(){return [...document.querySelectorAll('.page-row:not(.page-filter-hidden)')].filter(row=>row.offsetParent!==null);}
function pageRowKeyboardTarget(row,key){
  const rows=visiblePageRowsForKeyboard(),index=rows.indexOf(row);if(index<0||!rows.length)return null;if(key==='Home')return rows[0];if(key==='End')return rows[rows.length-1];
  if(pageBoardView==='thumbnail'&&(key==='ArrowUp'||key==='ArrowDown')){const current=row.getBoundingClientRect(),currentX=(current.left+current.right)/2,currentY=(current.top+current.bottom)/2,direction=key==='ArrowUp'?-1:1,sameVolume=rows.filter(candidate=>candidate.parentElement===row.parentElement&&candidate!==row),ranked=sameVolume.map(candidate=>{const rect=candidate.getBoundingClientRect(),x=(rect.left+rect.right)/2,y=(rect.top+rect.bottom)/2,dy=(y-currentY)*direction;return{candidate,dy,score:dy*10000+Math.abs(x-currentX)};}).filter(item=>item.dy>4).sort((a,b)=>a.score-b.score);if(ranked.length)return ranked[0].candidate;}
  const delta=(key==='ArrowLeft'||key==='ArrowUp')?-1:1,next=index+delta;return next>=0&&next<rows.length?rows[next]:null;
}
function handlePageRowNavigation(row,e){
  if(!['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Home','End'].includes(e.key))return false;const target=pageRowKeyboardTarget(row,e.key);if(!target)return false;e.preventDefault();target.focus();target.scrollIntoView({block:'nearest',inline:'nearest'});return true;
}

// 以前は再描画のたびに行ごと約8個のリスナーを張り直していた。174ページで1描画
// あたり約1,400個のクロージャを生成・破棄することになり、検索の1打鍵ごとに
// それが繰り返されていた。#page-board に1組だけ委譲リスナーを置く。
// 委譲なら、後から差し込まれたカード(遅延展開分)にもそのまま効く。
let boardEventsDelegated=false;
function handleBoardRowKeydown(row,e){
  if(e.target.matches('input,select,textarea,button,a,[contenteditable="true"]'))return;
  if(!e.altKey){
    if(e.key===' '||e.key==='Spacebar'){e.preventDefault();const ch=row.querySelector('[data-page-check]');if(ch){ch.checked=!selectedPages.has(ch.value);handlePageCheckboxToggle(ch,e.shiftKey);}return;}
    if(e.key==='Enter'){const preview=row.querySelector('[data-preview-page]');if(preview){e.preventDefault();previewPage(preview.dataset.previewPage||row.dataset.pageId,pageFallbackFromRow(row));}return;}
    handlePageRowNavigation(row,e);return;
  }
  if(!['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'].includes(e.key))return;
  e.preventDefault();const tbody=row.parentElement;const rows=[...tbody.querySelectorAll('.page-row:not(.page-filter-hidden)')];const idx=rows.indexOf(row);if(idx<0)return;const beforeVolumes=collectBoardVolumes();let moved=false;
  if(e.key==='ArrowLeft'||e.key==='ArrowRight'){
    const volumeOrder=['none',...activeTargetVolumes()],current=String(tbody.dataset.volume||''),currentIndex=volumeOrder.indexOf(current);if(currentIndex<0)return;const targetIndex=e.shiftKey?(e.key==='ArrowLeft'?0:volumeOrder.length-1):currentIndex+(e.key==='ArrowLeft'?-1:1);if(targetIndex<0||targetIndex>=volumeOrder.length||targetIndex===currentIndex)return;const targetVolume=volumeOrder[targetIndex],target=[...document.querySelectorAll('[data-volume]')].find(container=>String(container.dataset.volume||'')===targetVolume);if(!target)return;const targetPanel=target.closest('.volume-panel');targetPanel?.classList.remove('collapsed');targetPanel?.querySelector('[data-toggle-volume]')?.setAttribute('aria-expanded','true');pageVolumeCollapseOverrides.set(targetVolume,false);for(const movingRow of getRowsForDrag(row)){target.appendChild(movingRow);moved=true;}
  }else if(e.shiftKey){if(e.key==='ArrowUp'&&idx>0){tbody.insertBefore(row,rows[0]);moved=true;}else if(e.key==='ArrowDown'&&idx<rows.length-1){tbody.appendChild(row);moved=true;}}
  else if(e.key==='ArrowUp'&&idx>0){tbody.insertBefore(row,rows[idx-1]);moved=true;}else if(e.key==='ArrowDown'&&idx<rows.length-1){tbody.insertBefore(rows[idx+1],row);moved=true;}
  if(!moved)return;renumberBoardRows();scheduleBoardSave({label:e.key==='ArrowLeft'||e.key==='ArrowRight'?'キーボードでの出力先移動':'キーボードでの並べ替え',volumes:beforeVolumes});row.focus();
}
function volumeIdsForSelectAll(volume){
  const container=[...document.querySelectorAll('[data-volume]')].find(t=>String(t.dataset.volume||'')===String(volume||''));
  return container?[...container.querySelectorAll('.page-row:not(.page-filter-hidden)')].map(r=>r.dataset.pageId).filter(Boolean):[];
}
function delegateBoardEvents(box){
  if(boardEventsDelegated)return;
  boardEventsDelegated=true;
  box.addEventListener('pointerdown',e=>{const row=e.target.closest('.page-row');if(row)beginPointerPageDrag(e,row);});
  box.addEventListener('click',e=>{
    const check=e.target.closest('[data-page-check]');
    if(check){e.stopPropagation();handlePageCheckboxToggle(check,e.shiftKey);return;}
    const toggle=e.target.closest('[data-toggle-volume]');
    if(toggle){const panel=toggle.closest('.volume-panel'),volume=String(toggle.dataset.toggleVolume||'');const collapsed=!panel.classList.contains('collapsed');panel.classList.toggle('collapsed',collapsed);toggle.setAttribute('aria-expanded',String(!collapsed));pageVolumeCollapseOverrides.set(volume,collapsed);lastPageBoardRenderSignature='';if(!collapsed)preparePageThumbnails();return;}
    const edit=e.target.closest('[data-thumb-edit]');
    if(edit){e.stopPropagation();const row=edit.closest('.page-thumb-card');setPageThumbnailEditor(row,!row?.classList.contains('editing'));return;}
    const cancel=e.target.closest('[data-thumb-cancel]');
    if(cancel){e.stopPropagation();setPageThumbnailEditor(cancel.closest('.page-thumb-card'),false,true);return;}
    const save=e.target.closest('[data-thumb-save]');
    if(save){e.stopPropagation();void savePageFromThumbnail(save.closest('.page-thumb-card'),save);return;}
    const preview=e.target.closest('[data-preview-page]');
    if(preview){if(draggingRow)return;if(e.target&&e.target.matches('input,select,option'))return;const row=preview.closest('.page-row');previewPage(preview.dataset.previewPage||row?.dataset.pageId,pageFallbackFromRow(row));return;}
    // .page-row, not .page-thumb-card: the detail view renders <tr class="page-row">
    // and used to ignore row clicks entirely, so the on-screen instruction
    // "ページをクリックして選ぶ" did nothing there and hitting the widest target
    // (the title input) silently started a rename instead.
    const card=e.target.closest('.page-row');
    if(card){
      if(card.dataset.suppressClick==='true'){delete card.dataset.suppressClick;e.preventDefault();e.stopPropagation();return;}
      if(e.target.closest('input,button,a,select,.drag-handle,.badge,.page-thumb-check,.page-thumb-editor,.page-thumb-preview,[data-preview-page]'))return;
      const ch=card.querySelector('[data-page-check]');if(!ch)return;
      ch.checked=!selectedPages.has(ch.value);handlePageCheckboxToggle(ch,e.shiftKey);
    }
  });
  box.addEventListener('dblclick',e=>{const row=e.target.closest('.page-row:not(.page-thumb-card)');if(!row)return;if(draggingRow||e.target.closest('input,select'))return;if(!row.querySelector('[data-preview-page]'))return;previewPage(row.dataset.pageId,pageFallbackFromRow(row));});
  box.addEventListener('keydown',e=>{
    const editor=e.target.closest('[data-thumb-editor]');
    if(editor){
      if(e.key==='Escape'){e.preventDefault();setPageThumbnailEditor(editor.closest('.page-thumb-card'),false,true);editor.closest('.page-thumb-card')?.querySelector('[data-thumb-edit]')?.focus();return;}
      if(e.key==='Enter'&&e.target.matches('[data-thumb-page-title]')){e.preventDefault();editor.querySelector('[data-thumb-save]')?.click();return;}
    }
    const row=e.target.closest('.page-row');
    if(row)handleBoardRowKeydown(row,e);
  });
  box.addEventListener('change',e=>{
    const selectAll=e.target.closest('[data-select-all-pages]');
    if(selectAll){const ids=volumeIdsForSelectAll(selectAll.dataset.selectAllPages);if(selectAll.checked)ids.forEach(id=>selectedPages.add(id));else ids.forEach(id=>selectedPages.delete(id));lastPageRangeAnchor='';syncPageSelectionUi();return;}
    if(e.target.closest('[data-page-range]')){savePageFromRow(e.target.closest('.page-row'),false);return;}
    if(e.target.closest('[data-page-numbering]')){savePageFromRow(e.target.closest('.page-row'),true);}
  });
  // blur は伝播しないため focusout を使う。
  box.addEventListener('focusout',e=>{if(e.target.closest('[data-page-title]'))savePageFromRow(e.target.closest('.page-row'),false);});
}
function attachBoardEvents() {
  const box=$('page-board');
  if(!box)return;
  delegateBoardEvents(box);
  // 委譲できない「その時点の状態」だけを描画のたびに反映する。
  box.querySelectorAll('.page-row').forEach(row=>row.setAttribute('aria-keyshortcuts','Space Enter ArrowUp ArrowDown ArrowLeft ArrowRight Home End Alt+ArrowUp Alt+ArrowDown Alt+ArrowLeft Alt+ArrowRight Control+A'));
  box.querySelectorAll('[data-select-all-pages]').forEach(ch=>{
    const ids=volumeIdsForSelectAll(ch.dataset.selectAllPages);
    const all=ids.length&&ids.every(id=>selectedPages.has(id)),some=ids.some(id=>selectedPages.has(id));
    ch.checked=!!all;ch.indeterminate=some&&!all;
  });
}

function collectBoardVolumes() {
  const volumes = {};
  document.querySelectorAll('[data-volume]').forEach(tbody => {
    volumes[tbody.getAttribute('data-volume')] = [...tbody.querySelectorAll('.page-row')].map(row => row.getAttribute('data-page-id'));
  });
  return volumes;
}
function applyPageMutationResult(payload){
  const result=payload?.result||payload||{};
  const structure=state?.structure;
  if(!structure)return;
  // 保存が通ったら、次の操作の基準を新しい指紋へ進める。ここを忘れると
  // 自分の直前の変更を「他のタブの変更」と誤認して2回目以降が必ず失敗する。
  if(result.layoutFingerprint)rememberLayoutFingerprint(activePackRecord()?.packId||activePackId,result.layoutFingerprint);
  if(Array.isArray(result.pages))structure.pages=result.pages;
  else if(result.page){
    const id=resolvedPageId(result.page);
    const index=asArray(structure.pages).findIndex(p=>resolvedPageId(p)===id);
    if(index>=0)structure.pages[index]=result.page;else structure.pages.push(result.page);
  }
  if(result.volumes)structure.volumes=result.volumes;
  const pages=asArray(structure.pages);
  state.summary=state.summary||{};
  state.summary.totalPages=pages.length;
  state.summary.confirmedPages=pages.filter(p=>String(p.status||'')==='confirmed').length;
  state.summary.uncheckedPages=pages.filter(p=>['rendered','stale','not-rendered'].includes(String(p.status||''))).length;
  state.summary.pdfReadyPages=pages.filter(p=>String(p.contentPdf||'').trim()).length;
  lastPageBoardRenderSignature='';
  renderNavBadges();renderPageOverview();renderFinalOverview();renderPages();if(isModalOpen())syncPreviewOrganizerControls();
}

function scheduleBoardSave(undo=null) {
  boardSaveRevision+=1;
  const afterVolumes=clonePageVolumes(undo?.afterVolumes||collectBoardVolumes());pendingBoardSaveVolumes=afterVolumes;
  if(undo){
    rememberPageLayoutUndo(undo.label,undo.volumes,afterVolumes);
    if(!pendingPageLayoutUndo)pendingPageLayoutUndo=undo;
  }
  if(boardSaveTimer){clearTimeout(boardSaveTimer);boardSaveTimer=null;}
  void saveBoardOrder();
}
// 他のタブが先にページ構成を保存していた場合。こちらの並びで上書きすると相手の変更が
// 消えるので、画面を最新へ戻し、何が起きたかと次にどうするかを伝える。
async function handlePageLayoutConflict(rejectedVolumes,requestRevision){
  const failedUndo=pageLayoutUndoStack[pageLayoutUndoStack.length-1];
  if(failedUndo&&pageVolumeSnapshotsEqual(failedUndo.after,rejectedVolumes))pageLayoutUndoStack.pop();
  try{await refresh();}catch{}
  if(requestRevision!==boardSaveRevision)return;
  selectedPages.clear();lastPageRangeAnchor='';lastPageBoardRenderSignature='';
  renderPages();
  updateBulkSelectionLabel();
  showMessage('warn','ページ構成が別の画面で変更されました',
    'この画面の並びは保存していません。最新の状態を読み込み直したので、内容を確認してからもう一度操作してください。',
    null,[],0);
}
function saveBoardOrder() {
  if(boardSaveTimer){clearTimeout(boardSaveTimer);boardSaveTimer=null;}
  const undo=pendingPageLayoutUndo;pendingPageLayoutUndo=null;
  const volumes=clonePageVolumes(pendingBoardSaveVolumes||collectBoardVolumes()),requestRevision=boardSaveRevision;pendingBoardSaveVolumes=null;
  const persist=async()=>{
    try{
      const response=await api('/api/v2/items/reorder',{method:'POST',body:pageApiBody({volumes})});
      const newerBoardExists=requestRevision!==boardSaveRevision;
      if(!newerBoardExists)applyPageMutationResult(response);
      if(!newerBoardExists)showMessage('ok','ページ構成を保存しました','未振り分け・本体・補足の割り当てと並びを反映しました。',null,[{label:'元に戻す',handler:()=>undoLastPageLayout()},{label:'提出用PDFへ',view:'final'}],8000);
      return true;
    }catch(e){
      if(e.code==='structure-conflict'){void handlePageLayoutConflict(volumes,requestRevision);return false;}
      showMessage('danger','並び替えを保存できません',userFriendlyError(e.message),e.detail||e.stack||e.message);
      lastPageBoardRenderSignature='';
      // 保存できなかった並びを画面に残すと、collectBoardVolumes() が次の操作でそれを一緒に
      // 送ってしまい、拒否されたはずの移動が無言で確定する。サーバの状態へ戻す。
      const newerBoardExists=requestRevision!==boardSaveRevision;
      if(!newerBoardExists){
        const failedUndo=pageLayoutUndoStack[pageLayoutUndoStack.length-1];
        if(failedUndo&&pageVolumeSnapshotsEqual(failedUndo.after,volumes))pageLayoutUndoStack.pop();
        selectedPages.clear();lastPageRangeAnchor='';
        renderPages();
        updateBulkSelectionLabel();
      }
      return false;
    }
  };
  boardSavePromise=boardSavePromise.then(persist,persist);
  return boardSavePromise;
}
async function savePageFromRow(row, numberingChanged=false) {
  if(!row)return;
  const pageId=row.getAttribute('data-page-id'),page=getPage(pageId);
  const body=pageApiBody({pageId,title:row.querySelector('[data-page-title]')?.value||page?.title||''});
  const pageRangeInput=row.querySelector('[data-page-range]');if(pageRangeInput){const range=String(pageRangeInput.value||'').trim();if(range)body.pageRange=range;else body.clearPageRange=true;}
  const numbering=row.querySelector('[data-page-numbering]')?.value||'auto';
  if(numberingChanged){if(numbering==='auto')body.resetNumbering=true;else{body.numberingMode=numbering;body.numberingManual=true;}}
  try{
    const response=await api('/api/pages/update',{method:'POST',body});
    applyPageMutationResult(response);
  }catch(e){if(e.code==='structure-conflict'){await handlePageSettingsConflict();return;}showMessage('danger','ページを保存できません',userFriendlyError(e.message),e.detail||e.stack||e.message);}
}

async function handlePageSettingsConflict(){
  try{await refresh();}catch{}
  lastPageBoardRenderSignature='';
  renderPages();
  showMessage('warn','ページ構成が別の画面で変更されました',
    'この変更は保存していません。最新の状態を読み込み直したので、内容を確認してからもう一度操作してください。',
    null,[],0);
}
function pageSettingsBody(pageId,title,numbering,pageRange=''){
  const body=pageApiBody({pageId:String(pageId||''),title:String(title||'').trim()});const range=String(pageRange||'').trim();if(range)body.pageRange=range;else body.clearPageRange=true;if(numbering==='auto')body.resetNumbering=true;else{body.numberingMode=numbering==='none'?'none':'visible';body.numberingManual=true;}return body;
}
async function restorePageSettings(previous){
  try{const response=await api('/api/pages/update',{method:'POST',body:pageSettingsBody(previous.pageId,previous.title,previous.numbering,previous.pageRange)});applyPageMutationResult(response);showMessage('ok','ページ設定を元に戻しました',previous.title);}
  catch(error){if(error.code==='structure-conflict'){await handlePageSettingsConflict();return;}showMessage('danger','ページ設定を元に戻せません',userFriendlyError(error.message),error.detail||error.stack||error.message);}
}
async function savePageFromThumbnail(row,btn){
  if(!row)return;const pageId=String(row.dataset.pageId||''),page=getPage(pageId),titleInput=row.querySelector('[data-thumb-page-title]'),rangeInput=row.querySelector('[data-thumb-page-range]'),numberingInput=row.querySelector('[data-thumb-page-numbering]'),title=String(titleInput?.value||'').trim(),pageRange=String(rangeInput?.value||'').trim(),numbering=String(numberingInput?.value||'auto');if(!title){showMessage('warn','ページ名を入力してください','ページ名は空にできません。');titleInput?.focus();return;}if(pageRange&&!/^\d+(\s*-\s*\d+)?$/.test(pageRange)){showMessage('warn','ページ範囲を確認してください','「2」または「2-5」の形式で入力してください。');rangeInput?.focus();return;}const previous={pageId,title:String(page?.title||page?.sheetName||''),pageRange:pageRangeText(page),numbering:numberingSelectValue(page)};
  await runBusy(btn,async()=>{try{const response=await api('/api/pages/update',{method:'POST',body:pageSettingsBody(pageId,title,numbering,pageRange)});applyPageMutationResult(response);showMessage('ok','ページ設定を保存しました',title,null,[{label:'元に戻す',handler:()=>restorePageSettings(previous)}],8000);}catch(error){if(error.code==='structure-conflict'){await handlePageSettingsConflict();return;}showMessage('danger','ページ設定を保存できません',userFriendlyError(error.message),error.detail||error.stack||error.message);}});
}




async function sortPagesBySheet(btn=null) {
  await runBusy(btn||$('sort-by-sheet-btn'),async()=>{
    const beforeVolumes=collectBoardVolumes();
    const response=await api('/api/pages/sort-by-sheet',{method:'POST',body:pageApiBody({volumes:[...activeTargetVolumes(),'none']})});
    applyPageMutationResult(response);
    const changed=rememberPageLayoutUndo('原稿内の順序への整列',beforeVolumes);
    showMessage('ok',changed?'原稿内の順序に整えました':'すでに原稿内の順序です',changed?'未振り分け・本体・補足の各表を整列しました。':'ページの並びに変更はありませんでした。',null,changed?[{label:'元に戻す',handler:()=>undoLastPageLayout()},{label:'提出用PDFへ',view:'final'}]:[{label:'提出用PDFへ',view:'final'}],8000);
  });
}
function pageFallbackFromRow(row) {
  return {
    pageId: row?.getAttribute('data-page-id') || '',
    workbookId: row?.getAttribute('data-workbook-id') || '',
    sheetName: row?.getAttribute('data-sheet-name') || '',
    contentPdf: row?.getAttribute('data-content-pdf') || ''
  };
}
function firstSelectedPage() {
  for (const id of selectedPages) {
    const page = getPage(id);
    if (page) return page;
  }
  return null;
}
async function previewSelectedPage(btn) {
  const page = firstSelectedPage();
  if (!page) { showMessage('warn', 'ページを選択してください', '確認したいページを1つ選んでください。行をダブルクリックしても開けます。'); return; }
  await previewPage(resolvedPageId(page), page);
}
function assignedPageVolume(page){return !page||page.enabled===false||String(page.volume)==='none'?'none':String(page.volume||(activeTargetVolumes()[0]||mainVolume()));}
function previewOrganizerPageIds(pageId){
  const current=getPage(pageId);if(!current)return[];const volume=assignedPageVolume(current),visible=visiblePageSelectionValues(),useVisible=visible.includes(String(pageId));const ordered=useVisible?visible.map(getPage).filter(Boolean):[...pagesForActivePreset()].sort(pageSort);return ordered.filter(page=>assignedPageVolume(page)===volume&&pagePreviewAvailable(page)).map(resolvedPageId);
}
function syncPreviewOrganizerControls(){
  const tools=$('preview-page-tools'),page=getPage(currentPreviewPageId);const active=!!(previewOrganizerMode&&page);if(tools)tools.classList.toggle('hidden',!active);if(!active){previewPageIds=[];return;}
  previewPageIds=previewOrganizerPageIds(currentPreviewPageId);if(!previewPageIds.includes(currentPreviewPageId))previewPageIds=[currentPreviewPageId];const index=Math.max(0,previewPageIds.indexOf(currentPreviewPageId)),volume=assignedPageVolume(page),position=$('preview-page-position'),previous=$('preview-prev-page'),next=$('preview-next-page');if(position)position.textContent=`${assignmentVolumeLabel(volume)} ${index+1} / ${previewPageIds.length}`;if(previous)previous.disabled=index<=0;if(next)next.disabled=index>=previewPageIds.length-1;
  document.querySelectorAll('[data-preview-target],#preview-move-none').forEach(button=>{const target=button.id==='preview-move-none'?'none':String(button.dataset.previewTarget||''),current=volume===target;button.disabled=current;button.classList.toggle('active',current);button.setAttribute('aria-pressed',String(current));});
}
async function navigatePreviewPage(delta){
  if(!previewOrganizerMode)return;syncPreviewOrganizerControls();const index=previewPageIds.indexOf(currentPreviewPageId),nextIndex=index+Number(delta||0);if(index<0||nextIndex<0||nextIndex>=previewPageIds.length)return;const page=getPage(previewPageIds[nextIndex]);if(page)await previewPage(resolvedPageId(page),page,{preserveFocus:true});
}
async function moveCurrentPreviewPageToVolume(volume,btn){
  const page=getPage(currentPreviewPageId),target=String(volume||'');if(!page||assignedPageVolume(page)===target)return;await runBusy(btn,async()=>{const beforeVolumes=collectBoardVolumes(),volumes=Object.fromEntries([...activeTargetVolumes(),'none'].map(key=>[key,[]])),pageId=resolvedPageId(page);for(const candidate of[...pagesForActivePreset()].sort(pageSort)){const id=resolvedPageId(candidate);if(!id||id===pageId)continue;const assigned=assignedPageVolume(candidate);if(!volumes[assigned])volumes[assigned]=[];volumes[assigned].push(id);}if(!volumes[target])volumes[target]=[];volumes[target].push(pageId);try{const response=await api('/api/v2/items/reorder',{method:'POST',body:pageApiBody({volumes})});applyPageMutationResult(response);rememberPageLayoutUndo('プレビューからのページ移動',beforeVolumes);syncPreviewOrganizerControls();showMessage('ok','ページを移動しました',`${page.title||page.sheetName||'ページ'}を${assignmentVolumeLabel(target)}へ移動しました。`,null,[{label:'元に戻す',handler:()=>undoLastPageLayout()}],8000);}catch(error){showMessage('danger','ページを移動できません',userFriendlyError(error.message),error.detail||error.stack||error.message);}});
}
async function previewPage(pageId, fallback={}, options={}) {
  const page=getPage(pageId)||getPageByWorkbookSheet(fallback.workbookId,fallback.sheetName);
  const pid=resolvedPageId(page)||String(pageId||'').trim();
  const contentPdf=String(page?.contentPdf||fallback.contentPdf||'').trim();
  if(!contentPdf){showMessage('warn','PDF未作成です','登録済み原稿を選択して「PDF作成」を押してください。');return;}
  const wasOpen=isModalOpen();if(!wasOpen&&!options.preserveFocus)modalReturnFocus=document.activeElement;
  const wb=getWorkbook(page?.workbookId||fallback.workbookId);
  $('preview-title').textContent=page?.title||fallback.title||'PDF確認';
  $('preview-subtitle').textContent=`${wb?.displayName||wb?.fileName||''}${(page?.sheetName||fallback.sheetName)?` / ${sourceUnitReference(wb,page?.sheetName||fallback.sheetName)}`:''}`;
  clearPreviewObjectUrl();
  const previewUrl=apiUrl('/api/file',{
    pageId:pid,
    workbookId:page?.workbookId||fallback.workbookId||'',
    sheetName:page?.sheetName||fallback.sheetName||''
  });
  $('preview-open-new').href=previewUrl;
  $('pdf-frame').src=`${previewUrl}#toolbar=1&navpanes=0`;
  previewOrganizerMode=!!page&&activeView==='pages';currentPreviewPageId=previewOrganizerMode?pid:'';syncPreviewOrganizerControls();
  $('preview-modal').classList.remove('hidden');
  if(!wasOpen)$('preview-close').focus();
}




function closePreview() {
  const modal=$('preview-modal');if(!modal||modal.classList.contains('hidden'))return;const returnPageId=currentPreviewPageId,originalFocus=modalReturnFocus;modal.classList.add('hidden');$('pdf-frame').src='about:blank';clearPreviewObjectUrl();previewOrganizerMode=false;currentPreviewPageId='';previewPageIds=[];$('preview-page-tools')?.classList.add('hidden');const currentCard=returnPageId?[...document.querySelectorAll('[data-page-id]')].find(row=>String(row.dataset.pageId||'')===returnPageId):null,focusTarget=originalFocus?.isConnected?originalFocus:currentCard?.querySelector('[data-preview-page]');if(focusTarget&&typeof focusTarget.focus==='function')focusTarget.focus();modalReturnFocus=null;
}


// Used by the success toast when more than one PDF was produced: put the cards
// that carry the per-target "PDFを開く" buttons in front of the user.
function focusFinalOutputs(){
  const grid=$('final-target-grid');
  if(!grid)return;
  grid.scrollIntoView({block:'center',behavior:window.matchMedia?.('(prefers-reduced-motion:reduce)').matches?'auto':'smooth'});
  const first=grid.querySelector('[data-final-open]:not(.hidden)');
  if(first){first.setAttribute('tabindex','-1');first.focus({preventScroll:true});}
}
async function openFinalVolume(volume, category=activePreset) {
  // 名前付きターゲットにして同じタブを使い回す。'_blank' だと開くたびにタブが増え、
  // どれが今見ているPDFなのか分からなくなる。
  let blank=null;try{blank=window.open('about:blank','reportbinder-pdf');}catch{}if(blank){try{blank.opener=null;}catch{}}
  try{const custom=!activePackIsBuiltIn(),path=custom?'/api/v2/outputs/file':'/api/final/file',body=custom?{packId:activePackId,targetId:targetIdFromVolume(volume)}:{volume,category};const objectUrl=await fetchPdfObjectUrl(path,body);if(blank&&!blank.closed)blank.location.href=objectUrl;else window.open(objectUrl,'reportbinder-pdf');
    // PDFは別タブで開く。戻り方が分からず迷う人がいるため、こちら側に案内を残す。
    // 自動では消さない(利用者は別タブへ行っており、戻ってきたときに読めないと意味がない)。
    showMessage('ok','別のタブでPDFを開きました','確認が終わったら、ブラウザのタブを切り替えてこの画面（ReportBinder）に戻ってください。',null,[],0);}
  catch(e){if(blank){try{blank.close();}catch{}}showMessage('danger','PDFを開けません',userFriendlyError(e.message),e.detail||e.stack||e.message);}
}
// 別タブのPDFは見るだけで、掴んで持ち出せない。メールに添付する・共有フォルダーへ
// 置くには実体が要るので、保存先のフォルダーをエクスプローラーで開いて選択状態にする。
async function revealFinalVolume(volume, trigger=null) {
  await runBusy(trigger,async()=>{
    const custom=!activePackIsBuiltIn();
    await api('/api/v2/outputs/reveal',{method:'POST',body:custom?{packId:activePackId,targetId:targetIdFromVolume(volume)}:{volume,category:activePreset}});
    showMessage('ok','保存先フォルダーを開きました','エクスプローラーでPDFが選択された状態になっています。そこからメールに添付したり、共有フォルダーへコピーしたりできます。');
  });
}


function pathElementValue(id) {
  const el=$(id);if(!el)return '';return String(('value' in el?el.value:el.textContent)||'').trim().replace(/^—$/,'');
}
function setPathElementValue(id,value) {
  const el=$(id);if(!el)return;if('value' in el)el.value=value||'';else el.textContent=value||'—';
}
async function applyDefaultChildPaths(force=false) {
  const sub = pathElementValue('submissionDir');
  if (!sub) return;
  try {
    const result=await api('/api/paths/defaults',{method:'POST',body:{submissionDir:sub}});
    const next=result?.paths||{};
    if(force||!pathElementValue('dataDir'))setPathElementValue('dataDir',next.dataDir||'');
    if(force||!pathElementValue('outputDir'))setPathElementValue('outputDir',next.outputDir||'');
    lastSubmissionDefaultBase=sub;
  } catch {}
}
function continueFirstRunAfterFolderSelection(){
  if(activePackRecord()||asArray(state?.packs).length)return;
  setTimeout(()=>openPackEditor('create'),120);
}
async function chooseSubmissionFolder(btn) {
  if (folderPickerBusy) return;
  const replacingExistingFolder=!!activePackRecord()||asArray(state?.packs).length>0;
  const previousFolder=String(state?.paths?.submissionDir||'');
  // 原稿フォルダーを変えると、作業データ領域ごと別のワークスペースに切り替わる。
  // 登録済みの原稿も作った一式も画面から消えるため、押した瞬間に実行してはいけない。
  if (replacingExistingFolder) {
    const packCount=asArray(state?.packs).filter(pack=>!pack?.archived).length;
    const sourceCount=asArray(state?.structure?.workbooks).length;
    const accepted=await confirmAction({
      title:'原稿フォルダーを変更しますか？',
      message:`いま表示している原稿 ${sourceCount}件と一式 ${packCount}件は、この画面から見えなくなります。`,
      detail:`現在のフォルダー：${previousFolder||'(未設定)'}\n\n作業内容が消えるわけではありません。元に戻すには、このフォルダーをもう一度選び直してください。`,
      confirmLabel:'変更する',
      danger:true
    });
    if(!accepted)return;
  }
  folderPickerBusy = true;
  const old = btn?.textContent;
  if (btn) { btn.disabled = true; btn.textContent = '選択画面を開く'; }
  showMessage('', '原稿フォルダーを選んでください', 'フォルダー選択画面を開いています。');
  try {
    await sleep(40); // Let the button/message repaint before the local server opens the native dialog.
    const data = await api('/api/submission/select-and-start', {
      method:'POST',
      body:{initialDir: state?.paths?.submissionDir || ''}
    });
    if (data.cancelled) {
      hideMessage();
      return;
    }
    const nextState = data.state || (await api('/api/state'));
    state = normalizeStatePayload(nextState);
    availableFiles = asArray(data.files || []);
    availableFilesScannedAt = String(data.scannedAt || '');
    selectedFiles.clear();
    selectedWorkbooks.clear();
    selectedPages.clear();
    lastFileRangeAnchor = '';
    lastWorkbookRangeAnchor = '';
    lastPageRangeAnchor = '';
    renderAll();
    if (!availableFiles.length) await loadFilesSilently(); else renderFileList(availableFiles);
    if (replacingExistingFolder) {
      // 旧パスを残さないと、戻りたくなったときに選び直す先が分からなくなる。
      // 画面のパス表示は既に新しい値へ上書きされている。
      showMessage('ok', '原稿フォルダーを変更しました', '新しいフォルダーの原稿を確認しました。',
        previousFolder?`前のフォルダー：${previousFolder}`:null,
        previousFolder?[{label:'前のフォルダーに戻す',handler:()=>restorePreviousSubmissionFolder(previousFolder)}]:[], 0);
    } else {
      showMessage('ok', '原稿フォルダーを設定しました', '続けて、今回まとめる一式に名前を付けます。');
    }
    if(replacingExistingFolder)setActiveView('excel');else continueFirstRunAfterFolderSelection();
  } catch(e) {
    showMessage('danger', '原稿フォルダーを選べません', userFriendlyError(e.message), e.detail || e.stack || e.message);
  } finally {
    folderPickerBusy = false;
    if (btn) { btn.disabled = false; btn.textContent = old; }
  }
}

// Kept for compatibility with older local pages; the current UX uses chooseSubmissionFolder().
// 変更直後に「やっぱり戻したい」を1操作で満たす。ネイティブのフォルダー選択画面から
// 元のパスを探し直させると、パス表示が既に上書きされているため到達できない。
async function restorePreviousSubmissionFolder(previousFolder) {
  const target=String(previousFolder||'').trim();
  if(!target)return;
  try{
    await api('/api/paths', {method:'POST', body:{submissionDir:target, dataDir:'', outputDir:''}});
    await refresh();
    await loadFiles(null);
    selectedFiles.clear();selectedWorkbooks.clear();selectedPages.clear();
    lastFileRangeAnchor='';lastWorkbookRangeAnchor='';lastPageRangeAnchor='';
    renderAll();
    showMessage('ok','前の原稿フォルダーに戻しました',target);
  }catch(e){
    showMessage('danger','前のフォルダーに戻せません',userFriendlyError(e.message),e.detail||e.stack||e.message);
  }
}
async function savePaths(btn) {
  const submissionDir = pathElementValue('submissionDir').trim();
  if (!submissionDir) {
    showMessage('warn', '原稿フォルダーを入力してください', 'フォルダーのパスを入力するか、「フォルダーを選ぶ」を使用してください。');
    $('submissionDir')?.focus();
    return;
  }
  const body = { submissionDir, dataDir: pathElementValue('dataDir'), outputDir: pathElementValue('outputDir') };
  await runBusy(btn, async () => {
    await api('/api/paths', {method:'POST', body});
    await refresh();
    await loadFiles(null);
    showMessage('ok', '原稿フォルダーを設定しました', '続けて、今回まとめる一式に名前を付けます。');
    continueFirstRunAfterFolderSelection();
  });
}
// 控え(アーカイブ)の作成に失敗したことをサーバーは archiveError で返していたが、
// どこからも読まれておらず画面に一切出ていなかった。控えは、出力の元になった原稿の
// 版を後日の掃除から守る仕組みと同じ経路にあるため、黙って落とすと後から追えなくなる。
function showBuildArchiveWarning(job) {
  const failed=asArray(job?.built).filter(item=>String(item?.archiveError||'').trim());
  if(!failed.length)return;
  const detail=String(failed[0].archiveError||'');
  setTimeout(()=>showMessage('warn','提出用PDFは出力しましたが、控えを保存できませんでした',`${detail} 提出用PDF自体は保存されています。出力フォルダーの空き容量とアクセス権を確認してから、もう一度出力してください。`,failed,[],0),700);
}
function showPostBuildWarning(volume) {
  const after=volumeReadiness(volume),blockers=asArray(after.blockers);
  if(blockers.length){setTimeout(()=>showMessage('warn','提出用PDFは出力しましたが、追加対応が必要です',blockers[0]?.message||'変換PDFを作成し直してから、提出用PDFを再出力してください。',blockers,[{label:'原稿・変換PDFへ',view:'excel'}],0),900);}
  else if(String(after.displayState)==='needs-rebuild'){setTimeout(()=>showMessage('warn','提出用PDFの再出力が必要です','出力中に入力が変更されました。最新状態で再度出力してください。',after.staleReasons,[{label:'提出用PDFを確認',view:'final'}],0),900);}
}
async function buildVolume(volume, btn) {
  const ready=volumeReadiness(volume);if(asArray(ready.blockers).length){setActiveView('excel');showMessage('warn','先に原稿の変換PDFを作成してください',ready.blockers[0]?.message||'出力条件を確認してください。');return;}
  const unassigned=unresolvedPageCount();
  if(unassigned){const accepted=await confirmAction({title:`未振り分け ${unassigned}ページを除外しますか？`,message:'未振り分けのページは、今回の提出用PDFに入りません。',detail:'意図しない原稿落ちを防ぐため、通常は「キャンセル」してページ構成を確認してください。',confirmLabel:'除外して出力'});if(!accepted)return;}
  await runBusy(btn,async()=>{
    const job=await runFinalBuildJob([targetIdFromVolume(volume)]);
    if(!job)return;
    if(job.status==='cancelled'){showMessage('warn','提出用PDFの出力を中止しました',String(job.message||''),null,[],0);return;}
    if(job.status==='failed'||asArray(job.errors).length){
      showMessage('danger','提出用PDFを出力できませんでした',userFriendlyError(asArray(job.errors)[0]?.userError||asArray(job.errors)[0]?.error||job.message||''),job.errors,
        [{label:'動作環境を診断',primary:true,handler:()=>runSystemDiagnostics($('run-diagnostics-btn'))}]);
      return;
    }
    const built=asArray(job.built)[0];
    showMessage('ok',`${volumeLabel(volume)}PDFを出力しました`,built?.outputPdf?`${String(built.outputPdf)} に保存しました。`:'出力フォルダーを確認してください。',built,[{label:'PDFを開く',primary:true,handler:()=>openFinalVolume(volume,activePreset)},{label:'保存先フォルダーを開く',handler:()=>revealFinalVolume(volume)}],0);
    // 控えの作成に失敗しても提出用PDF自体はできている。ただし黙って進むと、
    // 出力の元になった原稿の版がいつ消えたか誰にも分からなくなる。
    showBuildArchiveWarning(job);
    showPostBuildWarning(volume);
  },false);
}

// 提出用PDFの出力はジョブとして走らせ、変換PDFと同じ進捗パネルで件数・段階・中止を出す。
// 単発awaitのままだと、数十秒のあいだ4秒で消えるトーストしか手掛かりがない。
async function runFinalBuildJob(targetIds){
  const ids=asArray(targetIds).map(id=>String(id||'')).filter(Boolean);
  if(!ids.length){showMessage('warn','出力先がありません','ページ構成で本体または補足にページを設定してください。');return null;}
  const started=await api('/api/v2/outputs/build',{method:'POST',body:{packId:activePackId,targetIds:ids}});
  const jobId=String(started.job?.jobId||'');
  if(!jobId)throw new Error('提出用PDFの出力を開始できませんでした。もう一度「提出用PDFをまとめて出力」を押してください。');
  activeFinalJobId=jobId;activeFinalCancelRequested=false;
  updateFinalProgressPanel(started.job);
  try{
    for(;;){
      await sleep(700);
      const polled=await api(`/api/v2/outputs/build/status?jobId=${encodeURIComponent(jobId)}`);
      const job=polled.job||{};
      updateFinalProgressPanel(job);
      if(['completed','completed-with-errors','failed','cancelled','missing'].includes(String(job.status||'')))
        {
          // Was: refresh() painted, then loadFinalReadiness() painted again a
          // network round trip later, with the progress panel still up and then
          // vanishing - three visual states in a row, seen as flicker.
          await refresh({render:false});
          await loadFinalPanels({force:true, render:false});
          hideProgressPanel();
          renderAll();
          return job;
        }
    }
  } finally {
    activeFinalJobId='';activeFinalCancelRequested=false;
    hideProgressPanel();
  }
}
function updateFinalProgressPanel(job){
  if(!job)return;
  // 変換PDFの進捗パネルをそのまま使う。中止ボタンの宛先だけ出力ジョブへ切り替える。
  updateProgressPanel({
    status:String(job.status||'running'),
    percent:Number(job.percent||0),
    total:Number(job.total||0),
    completed:Number(job.completed||0),
    failed:Number(job.failed||0),
    message:String(job.message||''),
    currentWorkbookName:'',
    currentSheet:''
  });
  const title=$('progress-title');
  if(title)title.textContent=String(job.status||'')==='cancelled'?'提出用PDFの出力を中止しました'
    :String(job.status||'')==='completed'?'提出用PDFを出力しました'
    :(String(job.status||'')==='failed'||String(job.status||'')==='completed-with-errors')?'提出用PDFの出力を確認してください'
    :'提出用PDFを作成中';
  const cancel=$('progress-cancel');
  if(cancel){
    const terminal=['completed','completed-with-errors','failed','cancelled','missing'].includes(String(job.status||''));
    cancel.hidden=!activeFinalJobId||terminal;
    cancel.disabled=activeFinalCancelRequested;
    cancel.textContent=activeFinalCancelRequested?'中止を受け付けました':'出力を中止';
  }
}
async function cancelActiveFinalJob(){
  if(!activeFinalJobId||activeFinalCancelRequested)return;
  activeFinalCancelRequested=true;
  const cancel=$('progress-cancel');
  if(cancel){cancel.disabled=true;cancel.textContent='中止を受け付けました';}
  try{
    const response=await api('/api/v2/outputs/build/cancel',{method:'POST',body:{jobId:activeFinalJobId}});
    showMessage('warn','出力の中止を受け付けました',String(response.result?.message||'作成中の1冊は最後まで書き上げてから停止します。'),null,[],0);
  }catch(e){
    activeFinalCancelRequested=false;
    showMessage('danger','出力を中止できません',userFriendlyError(e.message),e.detail||e.stack||e.message);
  }
}

async function publishFinalVolume(volume,btn) {
  await runBusy(btn,async()=>{
    const custom=!activePackIsBuiltIn();
    const response=await api(custom?'/api/v2/outputs/publish':'/api/final/publish',{method:'POST',body:custom?{packId:activePackId,targetId:targetIdFromVolume(volume)}:{volume,category:activePreset}});
    const result=response.result||{};
    showMessage('ok',`${volumeLabel(volume)}PDFを共有用フォルダーにコピーしました`,
      `${result.folderName||'発行フォルダー'} に ${result.fileName||'PDF'} を保存しました。`,
      result,[],0);
  });
}

async function buildAllVolumes(btn) {
  const unassigned=unresolvedPageCount();
  // Match buildVolume: the per-target button lets you confirm and go ahead, so
  // refusing outright here made two adjacent buttons disagree about the rule.
  if(unassigned){
    const accepted=await confirmAction({title:`未振り分け ${unassigned}ページを除外しますか？`,message:'未振り分けのページは、今回の提出用PDFに入りません。',detail:'ページ構成へ戻って振り分ける場合は「キャンセル」を押してください。',confirmLabel:'除外して出力'});
    if(!accepted){setActiveView('pages');return;}
  }
  // V5: 本体だけ成功する状態を作らないよう、組み込みパックの一括出力はサーバー側の
  // 準トランザクションAPIを1回だけ呼ぶ。ジョブ化してもその分岐はサーバー側で維持している。
  await runBusy(btn, async () => {
    const job=await runFinalBuildJob(activeTargetDefinitions().map(target=>String(target.targetId)));
    if(!job)return;
    if(job.status==='cancelled'){showMessage('warn','提出用PDFの出力を中止しました',String(job.message||''),null,[],0);return;}
    const built=asArray(job.built),skipped=asArray(job.skipped),errors=asArray(job.errors);
    if(errors.length){
      showMessage('danger','提出用PDFを出力できませんでした',userFriendlyError(errors[0]?.userError||errors[0]?.error||job.message||''),errors,
        [{label:'動作環境を診断',primary:true,handler:()=>runSystemDiagnostics($('run-diagnostics-btn'))}]);
      return;
    }
    if (built.length) {
      // The old action was {view:'final'} with no handler. You are already on the
      // final screen when you press 出力, so it navigated to where you stood and
      // looked like a dead button. Open the PDF instead.
      const openable=built.filter(item=>item&&item.volume);
      const names=built.map(item=>outputFileName(item?.outputPdf)).filter(Boolean);
      const detailText=names.length?`${names.join('、')} を保存しました。`:`${built.length}件を出力しました。`;
      const actions=openable.length===1
        ? [{label:'出力したPDFを開く',primary:true,handler:()=>openFinalVolume(String(openable[0].volume))}]
        : [{label:'出力したPDFを確認',primary:true,handler:()=>{setActiveView('final');focusFinalOutputs();}}];
      showMessage('ok','提出用PDFを出力しました', detailText, {built, skipped}, actions, 0);
      showBuildArchiveWarning(job);
    } else {
      showMessage('warn','出力対象がありません','ページ構成で本体または補足にページを設定してください。');
    }
  }, false);
}


async function runBusy(btn, fn, showProcessing=true) {
  // Disabling the focused button hands focus to <body>, and most of these
  // handlers re-render their own button, so remember the id and put focus back
  // rather than dumping keyboard users at the top of the page after every action.
  const btnId=btn?.id||'';
  const hadFocus=!!btn&&document.activeElement===btn;
  if(btn){btn.disabled=true;btn.classList.add('busy');btn.setAttribute('aria-busy','true');}
  if(showProcessing)showMessage('','処理中です','完了すると画面が更新されます。');
  try{await fn();}catch(e){hideProgressPanel();showMessage('danger','処理できませんでした',userFriendlyError(e.message),e.detail||e.stack||e.message);}
  finally{
    const back=(btnId&&$(btnId))||btn;
    if(back){back.disabled=false;back.classList.remove('busy');back.removeAttribute('aria-busy');}
    if(hadFocus&&back&&back.isConnected&&document.activeElement===document.body)back.focus({preventScroll:true});
  }
}


// ハートビートの失敗を黙って捨てると、画面は正常に見えたまま、次にボタンを押した
// 瞬間に初めて壊れたと分かる。続けて落ちたらサーバーは居ないと判断して知らせる。
// 1回だけの失敗では出さない（一時的な取りこぼしで驚かせないため）。
let heartbeatFailures = 0;
async function sendHeartbeat() {
  try {
    await api('/api/heartbeat', {method:'POST', body:{clientId}, keepalive:true});
    heartbeatFailures = 0;
  } catch {
    heartbeatFailures++;
    if (heartbeatFailures === 3) showServerGoneScreen();
  }
}
function showServerGoneScreen() {
  const screen=$('shutdown-screen');
  // 利用者が自分で終了した場合は既に出ている。その文言（このタブを閉じてください）を
  // 上書きしない。押した本人にとっては事故ではないため。
  if(!screen || !screen.classList.contains('hidden')) return;
  setTextIfChanged($('shutdown-title'), 'ReportBinderが終了しました');
  setTextIfChanged($('shutdown-detail'), SERVER_GONE_MESSAGE);
  screen.classList.remove('hidden');
  screen.focus();
}
function notifyClientClosing() {
  try {
    const payload = new Blob([JSON.stringify({clientId})], {type:'application/json'});
    navigator.sendBeacon(`/api/client/close?token=${encodeURIComponent(token)}`, payload);
  } catch { try { fetch(`/api/client/close?token=${encodeURIComponent(token)}`, {method:'POST', body:'{}', keepalive:true}); } catch {} }
}
function startLifecycle() {
  sendHeartbeat();
  setInterval(sendHeartbeat, 15000);
  // Do not stop the local server just because a browser tab is closed or refreshed.
  // Long PDF creation jobs should keep running; the server exits later after idle timeout.
}

function bind(id, event, handler) {
  const el = $(id);
  if (el) el.addEventListener(event, handler);
}

document.querySelectorAll('[data-view-nav]').forEach(btn=>btn.addEventListener('click',()=>{
  if(btn.classList.contains('nav-locked')){explainPackRequired();return;}
  setActiveView(btn.dataset.viewNav);
}));
bind('pages-next-final','click',()=>setActiveView('final'));
document.querySelectorAll('[data-view-shortcut]').forEach(btn=>btn.addEventListener('click',()=>setActiveView(btn.dataset.viewShortcut)));
setActiveView(activeView,{noScroll:true,instant:true});
$('relink-cancel')?.addEventListener('click',closeRelinkDialog);
$('relink-accept')?.addEventListener('click',event=>submitRelink(event.currentTarget));
$('relink-modal')?.addEventListener('keydown',event=>{if(event.key==='Escape'){event.preventDefault();closeRelinkDialog();}});
const fileFilterInput=$('file-filter');if(fileFilterInput)fileFilterInput.addEventListener('input',()=>{fileFilterText=fileFilterInput.value||'';renderFileList(availableFiles);});
bind('dashboard-action-btn','click',async()=>{const action=$('dashboard-action-btn')?.dataset.dashboardAction||'excel';if(action==='folder'){setActiveView('excel');setTimeout(()=>$('change-source-folder-btn')?.focus(),0);}else if(action==='create-pack'){openPackEditor('create');}else if(action==='render'){setActiveView('excel');await renderUpdated($('render-selected-btn'));}else setActiveView(action==='pages'?'pages':action==='final'?'final':'excel');});
bind('notice-close','click',hideMessage);bind('error-close','click',hideErrorPanel);bind('scan-btn','click',()=>scanAndRefresh($('scan-btn')));
bind('app-load-retry','click',()=>loadInitialAppState($('app-load-retry')));
bind('progress-cancel','click',()=>{if(activeFinalJobId)return void cancelActiveFinalJob();void cancelActiveRenderJob();});
bind('run-diagnostics-btn','click',()=>runSystemDiagnostics($('run-diagnostics-btn')));
bind('template-add-target','click',()=>{const requirements=collectTemplateRequirements(),destination=String($('template-new-destination')?.value||'unassigned');renderTemplateTargets([...collectTemplateTargets(),newTemplateTarget()],false);renderTemplateDestinationOptions(destination);renderTemplateRequirements(requirements,false);});
bind('pack-progress-attention-only','change',event=>{packProgressAttentionOnly=!!event.target.checked;renderPackProgressDashboard();});
bind('confirm-cancel','click',()=>closeConfirmModal(false));bind('confirm-accept','click',()=>closeConfirmModal(true));const confirmModal=$('confirm-modal');if(confirmModal)confirmModal.addEventListener('click',e=>{if(e.target===confirmModal)closeConfirmModal(false);});
bind('pack-menu-button','click',()=>togglePackMenu());bind('create-pack-btn','click',()=>openPackEditor('create'));bind('manage-pack-templates-btn','click',openTemplateManager);bind('toggle-archived-packs','click',()=>{archivedPacksExpanded=!archivedPacksExpanded;renderPackSwitcher();});
const packMenu=$('pack-menu');if(packMenu)packMenu.addEventListener('click',handlePackMenuAction);
bind('pack-editor-cancel','click',()=>closePackEditor());bind('pack-editor-template','change',updatePackTemplateDescription);const packEditorForm=$('pack-editor-form');if(packEditorForm)packEditorForm.addEventListener('submit',event=>{event.preventDefault();void submitPackEditor($('pack-editor-save'));});const packEditorModal=$('pack-editor-modal');if(packEditorModal)packEditorModal.addEventListener('click',event=>{if(event.target===packEditorModal)closePackEditor();});
bind('template-manager-cancel','click',()=>closeTemplateManager());bind('template-manager-select','change',populateTemplateManager);bind('template-add-requirement','click',()=>renderTemplateRequirements([...collectTemplateRequirements(),newTemplateRequirement()],false));bind('template-manager-delete','click',()=>deleteManagedTemplate($('template-manager-delete')));bind('template-manager-duplicate','click',()=>duplicateManagedTemplate($('template-manager-duplicate')));bind('template-manager-export','click',exportManagedTemplate);bind('template-manager-import','click',()=>$('template-manager-import-file')?.click());bind('template-manager-import-file','change',event=>importManagedTemplateFile(event.target.files?.[0],event.target));const templateManagerForm=$('template-manager-form');if(templateManagerForm)templateManagerForm.addEventListener('submit',event=>{event.preventDefault();void submitTemplateManager($('template-manager-save'));});const templateManagerModal=$('template-manager-modal');if(templateManagerModal)templateManagerModal.addEventListener('click',event=>{if(event.target===templateManagerModal)closeTemplateManager();});
document.addEventListener('click',event=>{if(!$('pack-menu')?.classList.contains('hidden')&&!event.target.closest('.pack-select-wrap'))closePackMenu();if(!$('app-menu-popover')?.classList.contains('hidden')&&!event.target.closest('.app-menu-wrap'))closeAppMenu();});
bind('use-submission-path-btn','click',()=>savePaths($('use-submission-path-btn')));bind('submissionDir','keydown',event=>{if(event.key==='Enter'){event.preventDefault();savePaths($('use-submission-path-btn'));}});
bind('clear-selected-files-btn','click',clearVisibleFilesSelection);bind('clear-selected-workbooks-btn','click',clearWorkbookSelection);bind('select-all-pages-btn','click',selectAllActivePages);bind('clear-selected-pages-btn','click',clearActivePageSelection);
bind('page-view-thumbnail-btn','click',()=>setPageBoardView('thumbnail'));bind('page-view-detail-btn','click',()=>setPageBoardView('detail'));bind('page-layout-undo-btn','click',()=>undoLastPageLayout($('page-layout-undo-btn')));bind('page-layout-redo-btn','click',()=>redoLastPageLayout($('page-layout-redo-btn')));bind('page-layout-revisions-btn','click',()=>openPageLayoutRevisions());
bind('page-layout-revisions-refresh','click',()=>loadLayoutSnapshots());bind('page-layout-revisions-close','click',()=>closePageLayoutRevisions());const pageLayoutRevisionsModal=$('page-layout-revisions-modal');if(pageLayoutRevisionsModal)pageLayoutRevisionsModal.addEventListener('click',event=>{if(event.target===pageLayoutRevisionsModal)closePageLayoutRevisions();});
const pageSearchInput=$('page-search-input'),pageSearchClear=$('page-search-clear');
// 1打鍵ごとに盤面を作り直すと、6文字入力で完全な破棄・再構築が6回走る。
// preparePageThumbnails がそのたびに描画キューを捨てるため、進行中のPDF描画も
// 打鍵のたびに破棄・再実行されていた。入力が落ち着いてから1回だけ描画する。
let pageSearchDebounceTimer=null;
if(pageSearchInput)pageSearchInput.addEventListener('input',()=>{
  pageFilterText=String(pageSearchInput.value||'');
  pageSearchClear?.classList.toggle('hidden',!pageFilterText);
  if(pageSearchDebounceTimer)clearTimeout(pageSearchDebounceTimer);
  pageSearchDebounceTimer=setTimeout(()=>{
    pageSearchDebounceTimer=null;
    // Keep the selection: renderPages() already drops ids that no longer exist,
    // and clearing here threw away work every time a filter character was typed.
    lastPageRangeAnchor='';lastPageBoardRenderSignature='';renderPages();
  },180);
});
if(pageSearchClear)pageSearchClear.addEventListener('click',()=>{pageFilterText='';if(pageSearchInput){pageSearchInput.value='';pageSearchInput.focus();}pageSearchClear.classList.add('hidden');lastPageRangeAnchor='';lastPageBoardRenderSignature='';renderPages();});
const pageThumbnailSizeInput=$('page-thumbnail-size');if(pageThumbnailSizeInput){pageThumbnailSizeInput.value=String(pageThumbnailSize);pageThumbnailSizeInput.addEventListener('input',()=>applyPageThumbnailSize(pageThumbnailSizeInput.value));pageThumbnailSizeInput.addEventListener('change',()=>{try{sessionStorage.setItem('ReportBinderPageThumbnailSize',String(pageThumbnailSize));}catch{}});}
bind('bulk-none-btn','click',()=>moveSelectedPagesToVolume('none',$('bulk-none-btn')));
bind('register-selected-btn','click',()=>registerSelected($('register-selected-btn')));bind('unregister-selected-btn','click',()=>unregisterSelected($('unregister-selected-btn')));bind('render-selected-btn','click',()=>renderSelectedWorkbooks($('render-selected-btn')));bind('render-all-btn','click',()=>renderAllWorkbooks($('render-all-btn')));bind('sort-by-sheet-btn','click',()=>sortPagesBySheet($('sort-by-sheet-btn')));
bind('build-all-btn','click',()=>buildAllVolumes($('build-all-btn')));
bind('source-next-pages','click',()=>setActiveView('pages'));
bind('change-source-folder-btn','click',()=>chooseSubmissionFolder($('change-source-folder-btn')));
bind('save-pack-settings-btn','click',()=>savePackSettings($('save-pack-settings-btn')));
bind('final-preflight-refresh','click',()=>refreshFinalPreflight($('final-preflight-refresh')));
bind('pack-review-approve','click',()=>performPackReviewAction('approve',$('pack-review-approve')));
bind('pack-review-reopen','click',()=>performPackReviewAction('reopen',$('pack-review-reopen')));
const changedOnlyToggle=$('changed-only-toggle');
if(changedOnlyToggle)changedOnlyToggle.addEventListener('change',()=>{showChangedOnly=!!changedOnlyToggle.checked;lastPageRangeAnchor='';lastPageBoardRenderSignature='';renderPages();});
bind('history-refresh-btn','click',()=>loadHistoryPanels({force:true}));
// V5-P1(#12): 版履歴パネルの対象Excel切り替え。
bind('snapshot-history-workbook','change',e=>{snapshotHistoryState.workbookId=String(e.target.value||'');try{sessionStorage.setItem('ReportBinderSnapshotWorkbook',snapshotHistoryState.workbookId);}catch{}snapshotHistoryState.snapshots=[];snapshotHistoryState.fromId='';snapshotHistoryState.toId='';const db=$('snapshot-history-diff');if(db)db.innerHTML='';void loadSnapshotHistory();});
bind('app-menu-button','click',()=>{const pop=$('app-menu-popover'),btn=$('app-menu-button');const hidden=pop.classList.toggle('hidden');btn.setAttribute('aria-expanded',String(!hidden));if(!hidden){closePopoverOnFocusOut(pop,closeAppMenu);requestAnimationFrame(()=>focusFirstIn(pop));}});bind('app-exit-button','click',()=>requestAppShutdown($('app-exit-button')));
bind('preview-close','click',closePreview);const previewModal=$('preview-modal');if(previewModal)previewModal.addEventListener('click',e=>{if(e.target===previewModal)closePreview();});
bind('preview-prev-page','click',()=>navigatePreviewPage(-1));bind('preview-next-page','click',()=>navigatePreviewPage(1));bind('preview-move-none','click',()=>moveCurrentPreviewPageToVolume('none',$('preview-move-none')));
bind('diff-close','click',closeDiffDetail);
bind('diff-export-summary','click',exportDiffSummary);
const diffModal=$('diff-modal');if(diffModal)diffModal.addEventListener('click',e=>{if(e.target===diffModal)closeDiffDetail();});
bind('diff-mode-side','click',()=>{diffViewState.mode='side';applyDiffMode();});
bind('diff-mode-overlay','click',()=>{diffViewState.mode='overlay';applyDiffMode();});
bind('diff-highlight','change',applyDiffHighlightSettings);
bind('diff-unreviewed-only','change',event=>{
  diffViewState.unreviewedOnly=!!event.target.checked;
  if(diffViewState.unreviewedOnly&&diffViewState.filter==='confirmed')diffViewState.filter='all';
  const visible=diffFilteredSheets();
  if(!visible.some(sheet=>String(sheet.sheetKey||'')===diffViewState.selectedSheetKey)){
    diffViewState.selectedSheetKey=String(visible[0]?.sheetKey||'');diffViewState.pageIndex=preferredDiffPageIndex(visible[0]);diffViewState.regionIndex=-1;
  }
  renderDiffSummary();renderDiffSheetList();renderDiffPage();
});
bind('diff-density','change',applyDiffHighlightSettings);
bind('diff-zoom','change',updateDiffStageScale);
bind('diff-opacity','input',applyDiffMode);
bind('diff-prev-page','click',()=>setDiffPage(diffViewState.pageIndex-1));
bind('diff-next-page','click',()=>setDiffPage(diffViewState.pageIndex+1));
bind('diff-prev-region','click',()=>moveDiffRegion(-1));
bind('diff-next-region','click',()=>moveDiffRegion(1));
bind('diff-confirm-sheet','click',()=>toggleDiffSheetReviewed($('diff-confirm-sheet')));
document.querySelectorAll('[data-diff-tab]').forEach(btn=>btn.addEventListener('click',()=>{
  diffViewState.mobileTab=btn.dataset.diffTab||'before';
  document.querySelectorAll('[data-diff-tab]').forEach(x=>{x.classList.toggle('active',x===btn);x.setAttribute('aria-pressed',String(x===btn));});
  $('diff-viewers')?.classList.toggle('show-after',diffViewState.mobileTab==='after');
  requestAnimationFrame(updateDiffStageScale);
}));
const beforeDiffViewport=$('diff-before-viewport'),afterDiffViewport=$('diff-after-viewport');
beforeDiffViewport?.addEventListener('scroll',()=>syncDiffScroll(beforeDiffViewport,afterDiffViewport),{passive:true});
afterDiffViewport?.addEventListener('scroll',()=>syncDiffScroll(afterDiffViewport,beforeDiffViewport),{passive:true});
window.addEventListener('resize',()=>{if(isDiffModalOpen())updateDiffStageScale();});
window.addEventListener('keydown',e=>{
  const shortcutKey=String(e.key||'').toLowerCase(),editableTarget=e.target.matches('input,select,textarea,[contenteditable="true"]');
  if(isDiffModalOpen()&&!editableTarget&&!e.ctrlKey&&!e.metaKey&&!e.altKey&&shortcutKey==='r'){e.preventDefault();void toggleDiffSheetReviewed();return;}
  if(activeView==='pages'&&!isAnyModalOpen()&&(e.ctrlKey||e.metaKey)&&!e.altKey&&!editableTarget&&shortcutKey==='a'){e.preventDefault();selectAllActivePages();return;}
  if(activeView==='pages'&&!isAnyModalOpen()&&(e.ctrlKey||e.metaKey)&&!e.altKey&&!editableTarget){const wantsUndo=shortcutKey==='z'&&!e.shiftKey,wantsRedo=(shortcutKey==='z'&&e.shiftKey)||(shortcutKey==='y'&&!e.shiftKey),available=wantsUndo?(pageLayoutUndoStack.length>0||!!pendingPageLayoutUndo):(wantsRedo&&pageLayoutRedoStack.length>0);if(available){e.preventDefault();void(wantsUndo?undoLastPageLayout():redoLastPageLayout());return;}}
  if(e.key==='Escape'){
    // Cancel an in-flight drag first. Previously Escape fell through to
    // clearActivePageSelection(), which unhighlighted the rows but left the
    // drag running, so the "cancelled" move still committed on drop.
    if(draggingRow){e.preventDefault();finishPointerDrag(false);return;}
    if(isConfirmModalOpen())closeConfirmModal(false);else if(isPageLayoutRevisionsOpen())closePageLayoutRevisions();else if(isTemplateManagerOpen())closeTemplateManager();else if(isPackEditorOpen())closePackEditor();else if(isDiffModalOpen())closeDiffDetail();else if(isModalOpen())closePreview();else if(!$('app-menu-popover')?.classList.contains('hidden'))closeAppMenu(true);else if(!$('pack-menu')?.classList.contains('hidden'))closePackMenu(true);else if(activeView==='pages'&&selectedPages.size)clearActivePageSelection();
    return;
  }
  if(isModalOpen()&&!isDiffModalOpen()&&previewOrganizerMode&&!e.altKey&&!e.ctrlKey&&!e.metaKey&&!e.target.matches('input,select,textarea')&&['ArrowLeft','ArrowRight'].includes(e.key)){e.preventDefault();void navigatePreviewPage(e.key==='ArrowLeft'?-1:1);return;}
  if(e.key!=='Tab')return;
  const modal=isConfirmModalOpen()?$('confirm-modal'):(isPageLayoutRevisionsOpen()?$('page-layout-revisions-modal'):(isTemplateManagerOpen()?$('template-manager-modal'):(isPackEditorOpen()?$('pack-editor-modal'):(isDiffModalOpen()?$('diff-modal'):(isModalOpen()?$('preview-modal'):null)))));
  // The pack switcher and the app menu are NOT trapped: they have no scrim and a
  // click anywhere outside dismisses them, so they are non-modal popovers. They
  // get focus-in, Escape and focus-return instead; tabbing past closes them.
  if(!modal)return;
  const focusable=[...modal.querySelectorAll('a[href],button:not([disabled]),input:not([disabled]),select:not([disabled]),[tabindex]:not([tabindex="-1"]),iframe')].filter(el=>el.offsetParent!==null);
  if(!focusable.length)return;
  const first=focusable[0],last=focusable[focusable.length-1];
  if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}
  else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}
});

let initialUpdateScanScheduled=false;
async function loadInitialAppState(button=null) {
  const screen=$('app-loading-screen'),title=$('app-loading-title'),message=$('app-loading-message'),layout=document.querySelector('.app-layout');
  document.body.classList.add('app-initializing');
  document.body.classList.remove('app-load-failed');
  if(screen)screen.setAttribute('aria-live','polite');
  if(title)title.textContent='ReportBinderを準備しています';
  if(message)message.textContent='原稿やページの状態を読み込んでいます。';
  if(button){button.classList.add('hidden');button.disabled=true;}
  if(layout){layout.setAttribute('aria-hidden','true');layout.setAttribute('inert','');}
  try {
    state=normalizeStatePayload(await api('/api/state'));
    if(configured()){
      try{await loadFilesSilently();}
      catch(e){log('原稿一覧の初期読み込みでエラー',e.detail||e.stack||e.message);}
    }
    // Paint once to reconcile the active pack (renderAll does that), then fetch
    // readiness, then paint again. Both happen while .app-layout is still
    // visibility:hidden, so the user only ever sees the settled result. Fetching
    // before the first renderAll does not work: activePackId is not resolved yet
    // and loadFinalReadiness() bails out, which left the header and the output
    // cards showing 未出力 / 0ページ until something else repainted them.
    renderAll();
    try{
      if(activeView==='final')await loadFinalPanels({render:false});
      else await loadFinalReadiness({render:false});
    }catch(e){log('提出用PDF状態の初期読み込みでエラー',e.detail||e.message);}
    renderAll();
    if(layout){layout.removeAttribute('aria-hidden');layout.removeAttribute('inert');}
    document.body.classList.remove('app-initializing','app-load-failed');
    if(configured()&&!initialUpdateScanScheduled){
      initialUpdateScanScheduled=true;
      // Heavy update checks can make the first folder/dialog action feel unresponsive on a single local server thread.
      // Run the first scan after the user has had time to interact; normal monitoring continues afterwards.
      setTimeout(() => scanUpdatesSilently({withFiles:true}), 30000);
    }
  } catch(e) {
    log('初期データを読み込めません',e.detail||e.stack||e.message);
    document.body.classList.add('app-load-failed');
    if(screen)screen.setAttribute('aria-live','assertive');
    if(title)title.textContent='資料を読み込めませんでした';
    if(message)message.textContent='接続を確認して、もう一度読み込んでください。';
    if(button){button.classList.remove('hidden');button.disabled=false;}
  }
}
(async function init() {
  startLifecycle();
  startUpdateMonitor();
  await loadInitialAppState($('app-load-retry'));
})();

window.addEventListener('pagehide', () => { clearDiffBrowserResources(true); for (const u of [...pdfObjectUrls]) revokePdfObjectUrl(u); });


// ---- V5: 自動処理の状態表示と履歴タイムライン ----
async function loadAutoStateDetail(){
  if(autoStateInFlight)return autoStateInFlight;autoStateInFlight=api('/api/auto/state').then(response=>{state.auto=response.auto||state.auto;renderAutoStatus();showAutoNotifications(state.auto);}).catch(error=>log('自動PDF作成の状態を取得できません',error.detail||error.message)).finally(()=>{autoStateInFlight=null;});return autoStateInFlight;
}
function showAutoNotifications(auto){
  const enabledResults=new Set([...(auto?.notifyOnCompletion?['completed']:[]),...(auto?.notifyOnFailure?['failed']:[])]);const item=asArray(auto?.items).filter(entry=>entry.notificationId&&enabledResults.has(String(entry.lastResult||''))&&!seenAutoNotificationIds.has(String(entry.notificationId))).sort((a,b)=>String(b.lastResultAt||'').localeCompare(String(a.lastResultAt||'')))[0];if(!item)return;seenAutoNotificationIds.add(String(item.notificationId));const wb=getWorkbook(item.workbookId),name=wb?workbookDisplayName(wb):String(item.workbookId||'原稿');if(String(item.lastResult)==='completed')showMessage('ok','自動PDF作成が完了しました',name);else showMessage('danger','自動PDF作成に失敗しました',`${name}：${item.lastError||'原稿を確認して再実行してください。'}`,null,[{label:'原稿を確認',view:'excel'}],0);
}
async function saveAutoSettings(btn){
  const body={enabled:!!$('auto-enabled')?.checked,quietPeriodSeconds:Number($('auto-quiet-seconds')?.value||180),maxRetryCount:Number($('auto-max-retries')?.value||3),minFreeMegabytes:Number($('auto-min-free')?.value||1024),notifyOnCompletion:!!$('auto-notify-success')?.checked,notifyOnFailure:!!$('auto-notify-failure')?.checked};await runBusy(btn,async()=>{const response=await api('/api/auto/settings',{method:'PATCH',body});state.auto=response.auto||body;renderAutoStatus();showMessage('ok','自動PDF作成の設定を保存しました',body.enabled?'原稿更新後、静止を確認して自動変換します。':'自動PDF作成を停止しました。');},false);
}
function renderAutoStatus(){
  const box=$('auto-status');
  if(!box)return;
  const auto=state?.auto;
  if(!auto||!auto.inputHistoryApproved){box.classList.add('hidden');box.innerHTML='';return;}
  box.classList.remove('hidden');
  const items=asArray(auto.items);
  const cap=Number(auto.softCapMegabytes||0),size=Number(auto.historySizeMb||0);
  const overCap=cap>0&&size>=cap*0.8;
  const lines=items.slice(0,8).map(it=>{
    const wb=getWorkbook(it.workbookId);
    const name=wb?workbookDisplayName(wb):String(it.workbookId||'');
    const reason={'excel-in-use':'Excelが使用されています','job-busy':'他のPDF作成が実行中です','start-failed':'開始できませんでした','file-missing':'ファイルが見つかりません','retry-backoff':`${Number(it.retryCount||0)}回目の再試行待ち（${formatDateTime(it.nextRetryAt)}）`,'retry-exhausted':it.lastError||'再試行回数の上限に達しました','disk-low':'空き容量が不足しています','job-status-missing':'実行状態を再確認しています'}[String(it.deferReason||'')]||'';
    const label={idle:it.lastResult==='completed'?'完了':'待機中',waiting:'変更を検知（静止待ち）',deferred:'保留中',rendering:'PDF作成中',ready:'実行待ち','retry-wait':'再試行待ち',failed:'自動処理失敗'}[String(it.state||'')]||String(it.state||'');
    const retryable=['deferred','retry-wait','failed'].includes(String(it.state||''));return `<div class="auto-row"><span>${escapeHtml(name)}</span><span>${escapeHtml(label)}</span><span class="subtext">${escapeHtml(reason||it.lastResultAt?`${reason}${reason?'・':''}${formatDateTime(it.lastResultAt)}`:'')}</span>${retryable?`<button class="btn ghost compact" type="button" data-auto-run="${escapeAttr(it.workbookId)}">今すぐ実行</button>`:''}</div>`;
  }).join('');
  const attentionCount=items.filter(item=>!['idle',''].includes(String(item.state||''))).length;
  const statusText=attentionCount?`${attentionCount}件に対応が必要`:auto.enabled?`有効・変更後 ${Math.round(Number(auto.quietPeriodSeconds||180)/60)}分待機`:'無効';
  box.innerHTML=`<details class="secondary-disclosure auto-disclosure"><summary><span>自動PDF作成</span><small>${escapeHtml(statusText)}</small></summary><div class="secondary-disclosure-body"><div class="disclosure-actions"><p>${attentionCount?'保留中または再試行待ちの原稿があります。':'原稿更新後のPDF作成を自動化できます。'}</p><button class="btn ghost compact" type="button" data-auto-settings-toggle>設定を変更</button></div><div class="auto-settings" data-auto-settings hidden><label><input id="auto-enabled" type="checkbox" ${auto.enabled?'checked':''}> 自動PDF作成を有効にする</label><label>静止待ち<select id="auto-quiet-seconds">${[[60,'1分'],[180,'3分'],[300,'5分'],[600,'10分']].map(([value,label])=>`<option value="${value}" ${Number(auto.quietPeriodSeconds||180)===value?'selected':''}>${label}</option>`).join('')}</select></label><label>失敗時の再試行<select id="auto-max-retries">${[0,1,3,5].map(value=>`<option value="${value}" ${Number(auto.maxRetryCount||3)===value?'selected':''}>${value}回</option>`).join('')}</select></label><label>必要な空き容量<select id="auto-min-free">${[[512,'0.5 GB'],[1024,'1 GB'],[2048,'2 GB'],[5120,'5 GB']].map(([value,label])=>`<option value="${value}" ${Number(auto.minFreeMegabytes||1024)===value?'selected':''}>${label}</option>`).join('')}</select></label><label><input id="auto-notify-success" type="checkbox" ${auto.notifyOnCompletion!==false?'checked':''}> 完了を通知</label><label><input id="auto-notify-failure" type="checkbox" ${auto.notifyOnFailure!==false?'checked':''}> 失敗を通知</label><button class="btn primary compact" type="button" data-auto-settings-save>保存</button></div>`
    +(items.length?`<div class="auto-list">${lines}</div>`:'<div class="caption">自動処理の待ち行列は空です。原稿を更新すると、ここに検知した変更が並びます。</div>')
    +(overCap?`<p class="warn-strip">変更履歴の使用量が ${size}MB です（上限の目安 ${cap}MB）。</p>`:'')
    +'</div></details>';
  box.querySelectorAll('[data-auto-run]').forEach(b=>b.addEventListener('click',()=>runBusy(b,async()=>{
    await api('/api/auto/run-now',{method:'POST',body:{workbookId:b.getAttribute('data-auto-run')}});
    await refresh();
    showMessage('ok','実行を要求しました','変換の開始を待っています。');
  },false)));
  box.querySelector('[data-auto-settings-toggle]')?.addEventListener('click',()=>{const settings=box.querySelector('[data-auto-settings]');if(settings)settings.hidden=!settings.hidden;});box.querySelector('[data-auto-settings-save]')?.addEventListener('click',event=>saveAutoSettings(event.currentTarget));
}

async function loadHistoryTimeline(){
  const box=$('history-timeline');
  if(!box)return;
  if(!state?.inputHistoryEnabled){box.innerHTML='<div class="caption">変更履歴の記録がオフになっているため、更新前後の比較は利用できません。比較を使うには、情報システム担当者に変更履歴の有効化を依頼してください。</div>';return;}
  try{
    const r=await api('/api/history/timeline?limit=50');
    const events=asArray(r.events);
    if(!events.length){box.innerHTML='<div class="caption">履歴はまだありません。</div>';return;}
    const label={'input.snapshot.created':'提出Excelの変更を検知','input.snapshot.deduplicated':'同じ内容を再検知','input.source.removed':'保存期限により現物を削除','render.started':'PDF作成を開始','render.completed':'PDF作成が完了','render.failed':'PDF作成に失敗','compare.completed':'変更を判定','layout.changed':'ページ構成を変更','layout.restored':'ページ構成を復元','final.built':'提出用PDFを出力','final.build.failed':'提出用PDFの出力に失敗','final.archive.created':'提出用PDFを保存','final.published':'提出用PDFを共有発行','auto.detected':'自動処理が変更を検知','auto.deferred':'自動処理を保留','history.cleanup':'履歴を整理'};
    box.innerHTML=events.map(e=>{
      const t=String(e.eventType||'');
      return `<div class="history-row"><span class="date-col">${escapeHtml(formatDateTime(e.at))}</span><span>${escapeHtml(label[t]||t)}</span><span class="subtext">${escapeHtml(String(e.data?.workbookId||e.data?.category||''))}</span></div>`;
    }).join('');
  }catch(e){box.innerHTML=`<div class="caption">履歴を読み込めません：${escapeHtml(userFriendlyError(e.message))}</div>`;}
}


// ---- ページ構成の永続履歴・限定復元 ----
let pageLayoutRevisionsReturnFocus=null;
function openPageLayoutRevisions(returnFocus=null){
  const modal=$('page-layout-revisions-modal');if(!modal)return;
  pageLayoutRevisionsReturnFocus=returnFocus||document.activeElement;modal.classList.remove('hidden');
  requestAnimationFrame(()=>$('page-layout-revisions-close')?.focus());
  void loadLayoutSnapshots();
}
function closePageLayoutRevisions(returnFocus=true){
  const modal=$('page-layout-revisions-modal');if(!modal||modal.classList.contains('hidden'))return;
  modal.classList.add('hidden');const focus=pageLayoutRevisionsReturnFocus;pageLayoutRevisionsReturnFocus=null;
  if(returnFocus&&focus?.isConnected&&typeof focus.focus==='function')focus.focus();
}
async function loadLayoutSnapshots(){
  const box=$('page-layout-revisions-list');
  if(!box)return;
  if(!state?.inputHistoryEnabled){box.innerHTML='<div class="caption">変更履歴の記録がオフになっているため、更新前後の比較は利用できません。比較を使うには、情報システム担当者に変更履歴の有効化を依頼してください。</div>';return;}
  box.innerHTML='<div class="empty-state">保存履歴を読み込んでいます。</div>';
  try{
    const custom=!activePackIsBuiltIn();
    const r=await api(custom?`/api/v2/layout/snapshots?packId=${encodeURIComponent(activePackId)}`:`/api/layout/snapshots?category=${encodeURIComponent(activePreset)}`);
    const list=asArray(r.snapshots);
    if(!list.length){box.innerHTML='<div class="empty-state">保存履歴はまだありません。ページを移動・並べ替え・設定変更すると、その直前の構成を自動保存します。</div>';return;}
    const reasonLabel={reorder:'ページの移動・並べ替え','sort-by-sheet':'原稿内の順序に整える','page-update':'ページ設定の変更','volume-change':'出力先の変更',numbering:'ページ番号の変更',enabled:'出力対象の変更',restore:'復元',
      'pre-restore':'復元の直前'};
    box.innerHTML=list.map(x=>`<div class="page-layout-revision"><div><strong>${escapeHtml(reasonLabel[String(x.reason||'')]||String(x.reason||'ページ構成の変更'))}</strong><span>${escapeHtml(formatDateTime(x.createdAt))}・${Number(x.pageCount||0)}ページ</span></div><button class="btn secondary compact" type="button" data-restore-layout="${escapeAttr(x.snapshotId)}">この構成に戻す</button></div>`).join('');
    box.querySelectorAll('[data-restore-layout]').forEach(b=>b.addEventListener('click',()=>previewLayoutRestore(b.getAttribute('data-restore-layout'),b)));
  }catch(e){box.innerHTML=`<div class="empty-state">履歴を読み込めません：${escapeHtml(userFriendlyError(e.message))}</div>`;}
}

async function previewLayoutRestore(snapshotId, btn){
  // Both the preview fetch and the restore itself can take seconds; with no
  // busy state people pressed again and the second press replaced the confirm
  // dialog that the first press was still opening.
  if(btn){if(btn.disabled)return;btn.disabled=true;btn.classList.add('busy');}
  try{
    const custom=!activePackIsBuiltIn(),body=custom?{packId:activePackId,snapshotId}:{category:activePreset,snapshotId};
    const r=await api(custom?'/api/v2/layout/restore/preview':'/api/layout/restore/preview',{method:'POST',body});
    const p=r.preview||{};
    const volumeLines=asArray(p.volumeChanges).slice(0,10).map(v=>`・${v.title||v.pageId}：${volumeLabel(v.from)} → ${volumeLabel(v.to)}`).join('\n');
      const detail=[`適用されるページ：${Number(p.appliedPageCount||0)}`,
      `過去にだけ存在するページ：${asArray(p.pastOnlyPageIds).length}（適用しません）`,
      `現在にだけ存在するページ：${asArray(p.currentOnlyPageIds).length}（そのまま残ります）`,
      volumeLines?`\n出力先の変更：\n${volumeLines}`:'',
      '','復元後、提出用PDFの再出力が必要になります。'].filter(Boolean).join('\n');
    closePageLayoutRevisions(false);
    const accepted=await confirmAction({title:'このページ構成に戻しますか？',message:`${formatDateTime(p.createdAt)} の構成を適用します。`,detail,confirmLabel:'この構成に戻す'});
    if(!accepted){openPageLayoutRevisions($('page-layout-revisions-btn'));return;}
    const res=await api(custom?'/api/v2/layout/restore':'/api/layout/restore',{method:'POST',body});
    clearPageLayoutHistory();
    await refresh();
    $('page-layout-revisions-btn')?.focus();
    showMessage('ok','ページ構成を復元しました',`${Number(res.result?.appliedPageCount||0)}ページに適用しました。取り消す場合は保存履歴の「復元の直前」から戻せます。`,null,
                [{label:'提出用PDFを確認',view:'final'}],0);
  }catch(e){showMessage('danger','復元できません',userFriendlyError(e.message));}
  finally{if(btn&&btn.isConnected){btn.disabled=false;btn.classList.remove('busy');}}
}

function volumeLabel(v){
  const volume=String(v||'');if(volume==='none')return '出力しない';const targetId=targetIdFromVolume(volume);if(targetId){const target=activeTargetDefinitions().find(item=>String(item?.targetId||'')===targetId);if(target?.displayName)return String(target.displayName);}return {'ja-main':'本体','ja-appendix':'補足','en-main':'Main','en-appendix':'Appendix'}[volume]||volume;
}

async function loadFinalArchives(){
  const box=$('final-archive-history');
  if(!box)return;
  if(!activePackRecord()?.packId){box.innerHTML='<div class="caption">資料パックを作成すると出力履歴が表示されます。</div>';return;}
  try{
    const r=await api(activePackIsBuiltIn()?`/api/final/archives?category=${encodeURIComponent(activePreset)}`:`/api/v2/outputs/archives?packId=${encodeURIComponent(activePackId)}`);
    const list=asArray(r.archives);
    if(!list.length){setHtmlIfChanged(box,'<div class="caption">保存された出力はまだありません。提出用PDFを出力すると、その時点のPDFを自動で保存します。</div>');return;}
    setHtmlIfChanged(box,list.slice(0,20).map(a=>`<div class="history-row"><span class="date-col">${escapeHtml(formatDateTime(a.builtAt))}</span><span>${escapeHtml(volumeLabel(a.volume))}</span><span class="subtext">${escapeHtml(String(a.outputFileName||''))}（${Number(a.pageCount||0)}ページ）</span></div>`).join(''));
  }catch(e){box.innerHTML=`<div class="caption">読み込めません：${escapeHtml(userFriendlyError(e.message))}</div>`;}
}

// ---- V5-P1(#12): 版の履歴・差分・保護（手動pin / 2版差分 / 過去content-pdf表示）----
let lastSnapshotHistoryWorkbook = '';
try { lastSnapshotHistoryWorkbook = sessionStorage.getItem('ReportBinderSnapshotWorkbook') || ''; } catch {}
let snapshotHistoryState = { workbookId:lastSnapshotHistoryWorkbook, snapshots:[], fromId:'', toId:'' };

function renderSnapshotHistorySelectionHint(){
  const box=$('snapshot-history-diff');if(!box)return;
  const from=snapshotHistoryState.snapshots.find(s=>String(s.snapshotId)===String(snapshotHistoryState.fromId));
  const to=snapshotHistoryState.snapshots.find(s=>String(s.snapshotId)===String(snapshotHistoryState.toId));
  if(!from||!to){box.innerHTML='<div class="caption">比較可能な版から比較元と比較先を選んでください。</div>';return;}
  const same=String(from.snapshotId)===String(to.snapshotId);
  const unavailable=!from.visualCompareReady||!to.visualCompareReady;
  const reason=unavailable?String((!from.visualCompareReady?from:to).unavailableReason||'比較用PDFがありません。'):'';
  box.innerHTML=`<div class="snapshot-comparison-hint"><strong>比較元</strong><span>${escapeHtml(formatDateTime(from.detectedAt))}</span><b aria-hidden="true">→</b><strong>比較先</strong><span>${escapeHtml(formatDateTime(to.detectedAt))}</span></div>`
    +`<div class="caption">${escapeHtml(same?'同じ版が選ばれています。異なる2版を選んでください。':unavailable?reason:'古い版を比較元、新しい版を比較先にすると、更新の流れを追いやすくなります。')}</div>`;
  const button=$('snap-diff-btn');if(button)button.disabled=same||unavailable;
}

function populateSnapshotHistoryWorkbooks(){
  const sel=$('snapshot-history-workbook');
  if(!sel)return;
  const list=workbooksForActivePreset();
  const prev=snapshotHistoryState.workbookId || sel.value || '';
  const signature=JSON.stringify(list.map(w=>[String(w.workbookId||''),workbookDisplayName(w)]));
  if(signature!==snapshotHistoryWorkbookOptionsSignature){
    sel.innerHTML='<option value="">選択してください</option>'+list.map(w=>`<option value="${escapeAttr(w.workbookId)}">${escapeHtml(workbookDisplayName(w))}</option>`).join('');
    snapshotHistoryWorkbookOptionsSignature=signature;
  }
  if(prev && list.some(w=>String(w.workbookId)===String(prev))){ sel.value=prev; snapshotHistoryState.workbookId=String(prev); }
  else {
    snapshotHistoryState.workbookId='';
    snapshotHistoryState.snapshots=[];
    snapshotHistoryState.fromId='';
    snapshotHistoryState.toId='';
    sel.value='';
  }
}

function snapshotHistoryCachedRecord(workbookId){
  const record=snapshotHistoryResponseCache.get(String(workbookId||''));
  if(!record||Date.now()-Number(record.savedAt||0)>SNAPSHOT_HISTORY_CACHE_MS)return null;
  return record;
}

async function fetchSnapshotHistoryList(workbookId,force=false){
  const id=String(workbookId||'');
  const cached=!force?snapshotHistoryCachedRecord(id):null;
  if(cached)return cached.snapshots;
  let promise=snapshotHistoryRequestCache.get(id);
  // A forced refresh bypasses the completed cache, but still joins an identical
  // request already in flight so the single-threaded local server is not queued twice.
  if(!promise){
    promise=api(`/api/history/snapshots?workbookId=${encodeURIComponent(id)}`);
    snapshotHistoryRequestCache.set(id,promise);
  }
  try{
    const response=await promise;
    const list=asArray(response.snapshots);
    snapshotHistoryResponseCache.set(id,{savedAt:Date.now(),snapshots:list});
    return list;
  }finally{
    if(snapshotHistoryRequestCache.get(id)===promise)snapshotHistoryRequestCache.delete(id);
  }
}

function syncSnapshotHistorySelectionInputs(){
  document.querySelectorAll('[data-snapshot-from]').forEach(el=>{const selected=String(el.dataset.snapshotFrom)===String(snapshotHistoryState.fromId);el.classList.toggle('active',selected);el.setAttribute('aria-pressed',String(selected));});
  document.querySelectorAll('[data-snapshot-to]').forEach(el=>{const selected=String(el.dataset.snapshotTo)===String(snapshotHistoryState.toId);el.classList.toggle('active',selected);el.setAttribute('aria-pressed',String(selected));});
  document.querySelectorAll('.snapshot-timeline-item').forEach(el=>{const id=String(el.dataset.snapshotId||'');el.classList.toggle('selected-from',id===String(snapshotHistoryState.fromId));el.classList.toggle('selected-to',id===String(snapshotHistoryState.toId));});
  renderSnapshotHistorySelectionHint();
}

async function loadSnapshotHistory(options={}){
  historyPanelsInitialized=true;
  const force=!!options.force;
  const requestSerial=++snapshotHistoryLoadSerial;
  const box=$('snapshot-history'), diffBox=$('snapshot-history-diff');
  if(!box)return;
  if(!state?.inputHistoryEnabled){snapshotHistoryState.snapshots=[];box.innerHTML='<div class="caption">変更履歴の記録がオフになっているため、更新前後の比較は利用できません。比較を使うには、情報システム担当者に変更履歴の有効化を依頼してください。</div>';if(diffBox)diffBox.innerHTML='';return;}
  populateSnapshotHistoryWorkbooks();
  const wbId=String(snapshotHistoryState.workbookId||'');
  if(!wbId){snapshotHistoryState.snapshots=[];box.innerHTML='<div class="caption">対象原稿を選ぶと、版の一覧・差分・保護を表示します。</div>';if(diffBox)diffBox.innerHTML='';return;}
  const cached=!force?snapshotHistoryCachedRecord(wbId):null;
  if(!cached)box.innerHTML='<div class="caption">版の一覧を読み込んでいます。</div>';
  try{
    const list=await fetchSnapshotHistoryList(wbId,force);
    if(requestSerial!==snapshotHistoryLoadSerial||String(snapshotHistoryState.workbookId||'')!==wbId)return;
    snapshotHistoryState.snapshots=list;
    if(!list.length){box.innerHTML='<div class="caption">この原稿の保存された版はまだありません。</div>';if(diffBox)diffBox.innerHTML='';return;}
    const ready=list.filter(s=>!!s.visualCompareReady);
    if(!snapshotHistoryState.toId||!ready.some(s=>String(s.snapshotId)===String(snapshotHistoryState.toId)))snapshotHistoryState.toId=String(ready[0]?.snapshotId||'');
    if(!snapshotHistoryState.fromId||!ready.some(s=>String(s.snapshotId)===String(snapshotHistoryState.fromId)))snapshotHistoryState.fromId=String(ready[1]?.snapshotId||'');
    box.innerHTML=`<ol class="snapshot-timeline" aria-label="保存された版">${list.map((s,index)=>{
      const id=String(s.snapshotId||'');const pinned=asArray(s.pins).map(String).includes('manual');
      const retained=s.sourceRetained?'あり':'期限切れ';const hash=String(s.sourceHash||'').slice(0,10);
      const ready=!!s.visualCompareReady;const reason=String(s.unavailableReason||'比較用PDFがありません。');
      const disabled=ready?'':`disabled title="${escapeAttr(reason)}"`;
      const selectedFrom=String(snapshotHistoryState.fromId)===id,selectedTo=String(snapshotHistoryState.toId)===id;
      return `<li class="snapshot-timeline-item ${selectedFrom?'selected-from ':''}${selectedTo?'selected-to':''}" data-snapshot-id="${escapeAttr(id)}"><span class="snapshot-timeline-marker" aria-hidden="true"></span><div class="snapshot-card"><div class="snapshot-card-main"><div class="snapshot-title"><time>${escapeHtml(formatDateTime(s.detectedAt))}</time>${index===0?badge('最新版','neutral'):''}${pinned?badge('保護中','neutral'):''}</div><p>${ready?'この版は視覚比較に使えます。':escapeHtml(reason)}</p><details><summary>版の詳細</summary><span>現物: ${escapeHtml(retained)} ／ 識別: ${escapeHtml(hash)}</span></details></div><div class="snapshot-card-actions"><div class="snapshot-role-buttons" aria-label="比較での役割"><button class="btn ghost compact ${selectedFrom?'active':''}" type="button" data-snapshot-from="${escapeAttr(id)}" aria-pressed="${selectedFrom}" ${disabled}>比較元にする</button><button class="btn ghost compact ${selectedTo?'active':''}" type="button" data-snapshot-to="${escapeAttr(id)}" aria-pressed="${selectedTo}" ${disabled}>比較先にする</button></div><button class="btn ghost compact" type="button" data-snap-pin="${escapeAttr(id)}" data-pinned="${pinned?'1':'0'}">${pinned?'保護を解除':'この版を保護'}</button></div></div></li>`;
    }).join('')}</ol><div class="snapshot-footer-actions"><button class="btn ghost compact" id="snap-swap-btn" type="button" ${ready.length<2?'disabled':''}>比較元・比較先を入れ替え</button><button class="btn primary compact" id="snap-diff-btn" type="button" ${ready.length<2?'disabled':''}>選んだ2版を比較</button></div>`;
    box.querySelectorAll('[data-snap-pin]').forEach(b=>b.addEventListener('click',()=>toggleSnapshotPin(wbId,b.getAttribute('data-snap-pin'),b.getAttribute('data-pinned')==='1',b)));
    box.querySelectorAll('[data-snapshot-from]').forEach(el=>el.addEventListener('click',()=>{snapshotHistoryState.fromId=el.dataset.snapshotFrom;syncSnapshotHistorySelectionInputs();}));
    box.querySelectorAll('[data-snapshot-to]').forEach(el=>el.addEventListener('click',()=>{snapshotHistoryState.toId=el.dataset.snapshotTo;syncSnapshotHistorySelectionInputs();}));
    const swap=$('snap-swap-btn');if(swap)swap.addEventListener('click',()=>{const previousFrom=snapshotHistoryState.fromId;snapshotHistoryState.fromId=snapshotHistoryState.toId;snapshotHistoryState.toId=previousFrom;syncSnapshotHistorySelectionInputs();});
    const db=$('snap-diff-btn');if(db)db.addEventListener('click',()=>loadSnapshotDiff(wbId));
    renderSnapshotHistorySelectionHint();
  }catch(e){
    if(requestSerial!==snapshotHistoryLoadSerial||String(snapshotHistoryState.workbookId||'')!==wbId)return;
    box.innerHTML=`<div class="caption">版の履歴を読み込めません：${escapeHtml(userFriendlyError(e.message))}</div>`;
  }
}

async function toggleSnapshotPin(workbookId, snapshotId, pinned, btn){
  // Without a busy state the button looked inert during the round trip, and a
  // double press sent pin and unpin back to back.
  if(btn){if(btn.disabled)return;btn.disabled=true;btn.classList.add('busy');}
  try{
    await api(pinned?'/api/history/unpin':'/api/history/pin',{method:'POST',body:{workbookId,snapshotId}});
    showMessage('ok',pinned?'保護を解除しました':'この版を保護しました',pinned?'':'保存期限による自動削除から守ります。');
    snapshotHistoryResponseCache.delete(String(workbookId||''));
    await loadSnapshotHistory({force:true});
  }catch(e){showMessage('danger','操作できません',userFriendlyError(e.message));}
  finally{if(btn){btn.disabled=false;btn.classList.remove('busy');}}
}

async function loadSnapshotDiff(workbookId){
  const diffBox=$('snapshot-history-diff');if(!diffBox)return;
  const from=String(snapshotHistoryState.fromId||''), to=String(snapshotHistoryState.toId||'');
  const fromItem=snapshotHistoryState.snapshots.find(s=>String(s.snapshotId)===from);
  const toItem=snapshotHistoryState.snapshots.find(s=>String(s.snapshotId)===to);
  if(!from||!to){diffBox.innerHTML='<div class="caption">比較可能な版から比較元と比較先を選んでください。</div>';return;}
  if(from===to){diffBox.innerHTML='<div class="caption">異なる2版を選んでください。</div>';return;}
  if(!fromItem?.visualCompareReady||!toItem?.visualCompareReady){diffBox.innerHTML='<div class="caption">この2つの版は比較できません。どちらかの版が古く、比較に使うPDFが残っていないためです。新しい2つの版を選び直してください。</div>';return;}
  diffBox.innerHTML='<div class="caption">選択した履歴版の差分を準備します。自動比較の基準は変更されません。</div>';
  await openDiffDetail(workbookId,$('snap-diff-btn'),{fromSnapshotId:from,toSnapshotId:to});
}

async function viewHistoryContentPdf(workbookId, snapshotId, versionId, sheetName){
  if(!versionId){showMessage('warn','この版のPDFは保存されていません','この版ではページPDFが作成されていないため表示できません。');return;}
  modalReturnFocus=document.activeElement;
  previewOrganizerMode=false;currentPreviewPageId='';previewPageIds=[];$('preview-page-tools')?.classList.add('hidden');
  const wb=getWorkbook(workbookId);
  $('preview-title').textContent='過去の版のPDF';
  $('preview-subtitle').textContent=`${wb?.displayName||wb?.fileName||''} / ${sourceUnitReference(wb,sheetName)}`;
  clearPreviewObjectUrl();
  const previewUrl=apiUrl('/api/history/content-pdf',{workbookId,versionId,sheetName,snapshotId});
  $('preview-open-new').href=previewUrl;
  $('pdf-frame').src=`${previewUrl}#toolbar=1&navpanes=0`;
  $('preview-modal').classList.remove('hidden');
  $('preview-close').focus();
}



function loadHistoryPanels(options={}){
  historyPanelsInitialized=true;
  return loadSnapshotHistory(options);
}
