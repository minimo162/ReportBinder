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
let activePreset = (() => { try { const v = sessionStorage.getItem('ReportBinderCategory'); return ['ecm','bod','dmm'].includes(v) ? v : 'ecm'; } catch { return 'ecm'; } })();
let lastSubmissionDefaultBase = '';
let draggingRow = null;
let boardSaveTimer = null;
let selectedFiles = new Set();
let selectedWorkbooks = new Set();
let selectedPages = new Set();
let lastFileRangeAnchor = '';
let lastWorkbookRangeAnchor = '';
let lastPageRangeAnchor = '';
let currentPreviewObjectUrl = '';
let updateMonitorTimer = null;
let updateMonitorInFlight = false;
let updateVisibilityBound = false;
let renderJobActive = false;
let folderPickerBusy = false;
let activeView = (() => { try { return sessionStorage.getItem('ReportBinderView') || 'dashboard'; } catch { return 'dashboard'; } })();
let fileFilterText = '';
const pdfObjectUrls = new Set();
let noticeTimer = null;
let lastPageBoardRenderSignature = '';
let modalReturnFocus = null;
let diffReturnFocus = null;
let diffPollToken = 0;
let diffSyncingScroll = false;
const diffViewState = {
  workbookId: '',
  fromSnapshotId: '',
  toSnapshotId: '',
  detail: null,
  selectedSheetKey: '',
  pageIndex: 0,
  regionIndex: -1,
  filter: 'all',
  mode: 'side',
  mobileTab: 'before'
};
let highlightedPageIds = new Set();
let pageHighlightTimer = null;
// サーバー側 $Script:ExcelPrintProfileVersion のフォールバック値。
// 実行時は /api/state の excelPrintProfileVersion を優先する（二重管理によるズレを防ぐ）。
const CURRENT_RENDER_PROFILE_VERSION = 2026060403;
function currentRenderProfileVersion() {
  const v = Number(state?.excelPrintProfileVersion || 0);
  return Number.isFinite(v) && v > 0 ? v : CURRENT_RENDER_PROFILE_VERSION;
}

const $ = (id) => document.getElementById(id);

const VIEW_NAMES = {
  dashboard: 'ダッシュボード',
  folders: '提出フォルダ',
  excel: 'Excel登録・PDF作成',
  history: '履歴・比較',
  pages: 'ページ構成',
  final: '最終PDF'
};

function setActiveView(view, options = {}) {
  const next = VIEW_NAMES[view] ? view : 'dashboard';
  activeView = next;
  try { sessionStorage.setItem('ReportBinderView', activeView); } catch {}
  document.querySelectorAll('[data-view-panel]').forEach(panel => {
    panel.classList.toggle('active', panel.getAttribute('data-view-panel') === activeView);
  });
  document.querySelectorAll('[data-view-nav]').forEach(btn => {
    const on = btn.getAttribute('data-view-nav') === activeView;
    btn.classList.toggle('active', on);
    if (on) btn.setAttribute('aria-current', 'page'); else btn.removeAttribute('aria-current');
  });
  if (!options.noScroll) window.scrollTo({top: 0, behavior: options.instant ? 'auto' : 'smooth'});
  if (activeView === 'history' && state) loadHistoryPanels();
}

function formatDateTime(value) {
  const text = String(value || '').trim();
  if (!text) return '—';
  const normalized = text.replace('T', ' ').replace(/\.\d+Z?$/, '').replace(/Z$/, '');
  const m = normalized.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})\s+(\d{1,2}):(\d{2})/);
  if (m) return `${m[1]}/${m[2].padStart(2,'0')}/${m[3].padStart(2,'0')} ${m[4].padStart(2,'0')}:${m[5]}`;
  return text.length > 22 ? text.slice(0, 22) : text;
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
function workbookDisplayName(w) { return String(w?.displayName || w?.fileName || w?.relativePath || w?.workbookId || '').trim() || 'Excel'; }
function fileDisplayName(f) { return String(f?.fileName || f?.relativePath || '').trim() || 'Excel'; }
function pageCountsByVolume() {
  const pages = pagesForActivePreset();
  const main = mainVolume();
  const appendix = appendixVolume();
  return {
    main: pages.filter(p => p.enabled !== false && String(p.volume || main) === main).length,
    appendix: pages.filter(p => p.enabled !== false && String(p.volume || main) === appendix).length,
    none: pages.filter(p => p.enabled === false || String(p.volume || main) === 'none').length,
    total: pages.length
  };
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
function modeTitle(mode) { return ({ja:'日本語管理', en:'英語管理'})[mode] || mode; }
function mainVolume() { return state?.language === 'en' ? 'en-main' : 'ja-main'; }
function appendixVolume() { return state?.language === 'en' ? 'en-appendix' : 'ja-appendix'; }
function volumeLabel(v) { return ({'ja-main':'本体','ja-appendix':'補足','en-main':'Main','en-appendix':'Appendix','none':'出力しない'})[v] || v; }
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
  return state?.finalReadiness?.[preset]?.volumes?.[volume] || {
    canBuild:false, pageCount:0, status:'not-built', displayState:'not-built', blockers:[], staleReasons:[], outputPdf:'', outputPdfExists:false
  };
}
function activeReadinessItems() {
  return [mainVolume(), appendixVolume()].map(volume => ({volume, ready:volumeReadiness(volume)})).filter(x => Number(x.ready.pageCount || 0) > 0);
}
function aggregateFinalState() {
  const items = activeReadinessItems();
  if (!items.length) return {state:'not-built', text:'最終PDFは未出力'};
  const displays = items.map(x => String(x.ready.displayState || 'not-built'));
  if (displays.includes('blocked')) return {state:'blocked', text:'最終PDFを出力できません'};
  if (displays.includes('needs-rebuild')) return {state:'needs-rebuild', text:'最終PDFの再出力が必要'};
  if (displays.includes('output-missing')) return {state:'output-missing', text:'前回出力が見つかりません'};
  if (displays.includes('not-built')) return {state:'not-built', text:'最終PDFは未出力'};
  return {state:'built', text:'最終PDFは最新です'};
}
function finalIsComplete() {
  const items = activeReadinessItems();
  return items.length > 0 && items.every(x => x.ready.status === 'built' && x.ready.outputPdfExists === true && asArray(x.ready.blockers).length === 0);
}
function unresolvedPageCount() {
  return pagesForActivePreset().filter(p => p.enabled !== false && String(p.volume || 'none') === 'none').length;
}
function isEditing() {
  const el = document.activeElement;
  return !!el && el.matches('input:not([readonly]), select, textarea');
}
function isModalOpen() { return !!$('preview-modal') && !$('preview-modal').classList.contains('hidden'); }
function shouldPreserveEditorDom() { return isEditing() || draggingRow !== null || isModalOpen(); }
function iconUse(id, cls='') { return `<svg class="icon ${cls}" aria-hidden="true"><use href="#${id}"></use></svg>`; }

function statusLabel(s) {
  return ({
    missing:'ファイルなし', new:'PDF未作成', rendering:'PDF作成中', 'rendered-unchecked':'PDF済み', confirmed:'PDF済み',
    'excel-updated':'更新あり', rejected:'差し戻し', finalized:'完了', rendered:'PDF済み', stale:'更新あり',
    'not-rendered':'PDF未作成', 'render-error':'作成エラー', 'needs-rebuild':'再出力待ち', built:'出力済み', 'not-built':'未出力'
  })[s] || (s || '未設定');
}
function statusClass(s) {
  if (s === 'missing' || s === 'rejected' || s === 'render-error') return 'danger';
  if (s === 'excel-updated' || s === 'stale' || s === 'needs-rebuild' || s === 'new' || s === 'not-rendered' || s === 'rendering') return 'attention';
  return 'neutral';
}

function isLatestPdfWorkbook(w) {
  if (!w || String(w.status || '') === 'missing' || String(w.status || '') === 'render-error') return false;
  const current = String(w.currentExcelHash || '').trim();
  const rendered = String(w.lastRenderedExcelHash || '').trim();
  const profileVersion = Number(w.renderProfileVersion || 0);
  if (!w.lastRenderedAt || !rendered) return false;
  if (profileVersion < currentRenderProfileVersion()) return false;
  if (current && current !== rendered) return false;
  if (['new','excel-updated','not-rendered'].includes(String(w.status || ''))) return false;
  return true;
}
function latestPdfStatusForWorkbook(w) {
  return isLatestPdfWorkbook(w) ? badge('最新PDFあり', 'neutral') : badge('PDF作成が必要', 'attention');
}

function latestPdfDetailForWorkbook(w) {
  if (!w) return '';
  if (String(w.status || '') === 'render-error') return w.lastErrorUser || userFriendlyError(w.lastError || 'PDF作成エラー');
  return '';
}
function pdfStatusForPage(p) {
  const wb = getWorkbook(p?.workbookId);
  if (wb && !isLatestPdfWorkbook(wb)) return badge('PDF作成が必要', 'attention');
  if (p.status === 'render-error') return badge('作成エラー', 'danger');
  if (p.contentPdf) return badge('最新PDFあり', 'neutral');
  return badge('PDF作成が必要', 'attention');
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
function userFriendlyError(message) {
  const text = String(message || '');
  const raw = text.length > 300 ? text.slice(0, 300) + '…' : text;
  // サーバーが日本語で説明しているメッセージは、そのまま表示するのが最も正確。
  // 技術的なエラー(例外・スタックトレース等)のときだけ、対処ヒントを付けて原文も併記する。
  const isTechnical = /Exception|StackTrace|HRESULT|at\s+java|COMException|System\./i.test(text);
  const hasJapanese = /[ぁ-んァ-ヶ一-龠]/.test(text);
  if (hasJapanese && !isTechnical) return raw;
  let hint = '';
  if (/用語 'java'|'java'\s*は.*認識され|java\.exe が見つかりません|Java Runtimeが(見つかり|あり)ません/i.test(text)) {
    hint = '最終PDF作成用のJava Runtimeが見つかりません。app\\tools\\install-thirdparty.cmd を実行してください。';
  } else if (/別のプロセスが使用中|being used by another process|使用中です/i.test(text)) {
    hint = 'ファイルが他のアプリで開かれています。該当ファイルを閉じてから再実行してください。';
  } else if (/アクセスが拒否|Access is denied|UnauthorizedAccess/i.test(text)) {
    hint = 'ファイルへのアクセスが拒否されました。共有フォルダの権限や一時的な競合の可能性があります。少し待って再試行してください。';
  } else if (/COMException|HRESULT|RPC/i.test(text)) {
    hint = 'Excelの操作に失敗しました。Excelをすべて閉じてから再試行してください。';
  } else if (/FileNotFoundException|not found|見つかりません/i.test(text)) {
    hint = 'ファイルが見つかりません。削除・移動・名前変更されていないか確認してください。';
  } else if (/OutOfMemory/i.test(text)) {
    hint = 'メモリ不足で処理できませんでした。他のアプリを閉じてから再試行してください。';
  }
  if (!hint) return raw;
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
    showErrorPanel(title, message, detail || message);
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
        if (action.view) setActiveView(action.view);
        if (typeof action.handler === 'function') await action.handler();
        hideMessage();
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
function showErrorPanel(title, summary, detail) {
  const panel = $('error-panel');
  if (!panel) return;
  const items = extractErrorItems(detail, summary);
  $('error-title').textContent = title || 'エラーがあります';
  $('error-summary').textContent = summary || '内容を確認してください。';
  $('error-list').innerHTML = items.length
    ? items.map(i => `<li>${i.label ? `<strong>${escapeHtml(i.label)}</strong>：` : ''}${escapeHtml(i.message)}</li>`).join('')
    : `<li>${escapeHtml(summary || '内容を確認してください。')}</li>`;
  panel.classList.remove('hidden');
}
function hideErrorPanel() {
  const panel = $('error-panel');
  if (panel) panel.classList.add('hidden');
}

function log(message, obj) {
  const time = new Date().toLocaleTimeString();
  const detail = obj ? `\n${typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2)}` : '';
  $('log').textContent = `[${time}] ${message}${detail}\n\n` + $('log').textContent;
}

async function api(path, options = {}) {
  const headers = Object.assign({'X-ReportBinder-Token': token}, options.headers || {});
  const hasBody = Object.prototype.hasOwnProperty.call(options, 'body');
  const rawBody = hasBody ? options.body : undefined;
  const isBinaryBody = hasBody && (rawBody instanceof FormData || rawBody instanceof Blob);
  if (hasBody && !isBinaryBody) headers['Content-Type'] = 'application/json';
  const res = await fetch(withToken(path), {
    method: options.method || 'GET',
    headers,
    body: hasBody ? (isBinaryBody ? rawBody : JSON.stringify(rawBody ?? {})) : undefined,
    keepalive: !!options.keepalive,
    cache: 'no-store'
  });
  const text = await res.text();
  let data;
  try { data = JSON.parse(text); } catch { data = {ok: false, error: text}; }
  if (!res.ok || data.ok === false) {
    const err = new Error(data.error || `HTTP ${res.status}`);
    err.detail = data.detail || text;
    err.payload = data;
    throw err;
  }
  return data;
}



function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
function updateProgressPanel(job) {
  const panel = $('progress-panel');
  if (!panel || !job) return;
  let pct = Math.max(0, Math.min(100, Number(job.percent || 0)));
  const status = String(job.status || '');
  if (pct <= 0 && ['queued','running','waiting'].includes(status)) pct = status === 'queued' ? 1 : 3;
  const total = Math.max(0, Number(job.total || 0));
  const done = Math.max(0, Number(job.completed || 0) + Number(job.failed || 0));
  const checking = status === 'waiting' || status === 'missing' || job.transient || job.transientMissing;
  $('progress-title').textContent = status === 'completed'
    ? 'PDF作成が完了しました'
    : (status === 'completed-with-errors' || status === 'failed'
      ? 'PDF作成を確認してください'
      : (checking ? 'PDF作成を確認中' : 'PDF作成中'));
  panel.style.setProperty('--progress', `${pct}%`);
  $('progress-percent').textContent = `${pct}%`;
  const countLabel = $('progress-count-label');
  if (countLabel) countLabel.textContent = total ? `${Math.min(done, total)} / ${total}件` : '';
  $('progress-fill').style.width = `${pct}%`;
  const target = job.currentSheet ? `${job.currentWorkbookName || ''} / シート ${job.currentSheet}` : (job.currentWorkbookName || '');
  $('progress-message').textContent = job.message || (target ? `${target} を処理しています。` : '処理しています。');
  panel.querySelector('.progress-track')?.setAttribute('aria-valuenow', String(pct));
  panel.classList.toggle('indeterminate', (checking || pct <= 0) && !['completed','completed-with-errors','failed'].includes(status));
  panel.classList.remove('hidden');
}
function hideProgressPanel() {
  const panel = $('progress-panel');
  if (panel) panel.classList.add('hidden');
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
    if (['completed','completed-with-errors','failed'].includes(status)) return last;
    if (status === 'waiting' || status === 'missing' || last.transient || last.transientMissing) {
      transientStatuses += 1;
      if (transientStatuses > 70) {
        const err = new Error('PDF作成ジョブの進捗を確認できません。処理中の場合は少し待って画面を更新してください。');
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

function apiUrl(path, params={}) {
  const url = new URL(withToken(path), window.location.origin);
  for (const [k, v] of Object.entries(params || {})) {
    if (v !== undefined && v !== null && String(v) !== '') url.searchParams.set(k, String(v));
  }
  return url.toString();
}

async function fetchPdfObjectUrl(path, params={}) {
  const postPdf = (path === '/api/file' || path === '/api/final/file');
  const headers = {'X-ReportBinder-Token': token, 'Accept': 'application/pdf,application/json'};
  if (postPdf) headers['Content-Type'] = 'application/json';
  const res = await fetch(apiUrl(path, postPdf ? {} : params), {
    method: postPdf ? 'POST' : 'GET',
    cache: 'no-store',
    headers,
    body: postPdf ? JSON.stringify(params || {}) : undefined
  });
  const contentType = (res.headers.get('content-type') || '').toLowerCase();
  if (!res.ok || contentType.includes('application/json')) {
    let detail = null;
    let message = `PDFを取得できませんでした。HTTP ${res.status}`;
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
  if (!configured() || updateMonitorInFlight || renderJobActive) return;
  updateMonitorInFlight = true;
  try {
    await api('/api/scan-updates', {method:'POST', body:{force:false}});
    state = normalizeStatePayload(await api('/api/state'));
    if (withFiles && !shouldPreserveEditorDom()) await loadFilesSilently();
    renderAll({preserveEditors:true});
  } catch (e) {
    log('更新確認でエラー', e.detail || e.stack || e.message);
  } finally { updateMonitorInFlight = false; }
}

function startUpdateMonitor() {
  if (updateMonitorTimer) clearInterval(updateMonitorTimer);
  updateMonitorTimer = setInterval(() => scanUpdatesSilently({withFiles:true}), 30000);
  if (!updateVisibilityBound) {
    document.addEventListener('visibilitychange', () => {
      if (!document.hidden) scanUpdatesSilently({withFiles:true});
    });
    updateVisibilityBound = true;
  }
}

async function refresh() {
  try {
    state = normalizeStatePayload(await api('/api/state'));
    renderAll();
    return state;
  } catch (e) {
    showMessage('danger', '状態を読み込めません', userFriendlyError(e.message), e.detail || e.stack || e.message);
    throw e;
  }
}
async function refreshAndLoad(btn) {
  await runBusy(btn, async () => {
    if (configured()) {
      try { await api('/api/scan-updates', {method:'POST', body:{force:false}}); } catch {}
    }
    await refresh();
    if (configured()) await loadFiles(null);
    showMessage('ok', '画面を更新しました', '提出フォルダと登録状態を読み込み直しました。');
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
    showMessage('', '元Excelの更新を確認しています', '登録済みExcelの保存日時・サイズ・内容を確認します。');
    const r = await api('/api/scan-updates', {method:'POST', body:{force:true}});
    state = normalizeStatePayload(await api('/api/state'));
    if (configured()) await loadFiles(null);
    renderAll();
    const n = r.result?.changedWorkbookIds?.length || 0;
    const scanned = Number(r.result?.scanned || 0);
    const detail = n
      ? `${n}件のExcelをPDF作成してください。`
      : (scanned ? `${scanned}件を確認しました。` : '登録済みExcelはありません。');
    showMessage(n ? 'warn' : 'ok', n ? '元Excelに更新があります' : '更新はありません', detail, r.result || null);
  }, false);
}

function renderAll(options={}) {
  const preserve = !!options.preserveEditors && shouldPreserveEditorDom();
  const appTitle = $('app-title');
  if (appTitle) appTitle.textContent = modeTitle(state?.mode || 'ja');
  const modeBadge = $('mode-badge');
  if (modeBadge) modeBadge.firstChild.textContent = state?.language === 'en' ? 'English ' : '日本語 ';
  const current = $('language-current'); if (current) current.textContent = modeTitle(state?.mode || 'ja');
  const guide = $('language-guide'); if (guide) guide.textContent = state?.language === 'en' ? '日本語管理.vbs を起動してください。' : '英語管理.vbs を起動してください。';
  setActiveView(activeView, {noScroll:true});
  updatePresetTabs();
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
  if(activeView==='history')loadHistoryPanels();
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
    showMessage('danger', 'PDF作成でエラーがあります', first?.message || `${errors}件のExcelでPDFを作成できません。`, recent);
  } else hideErrorPanel();
}

function renderPathInputs() {
  const submission = $('submissionDir'); if (submission) submission.value = state?.paths?.submissionDir || '';
  const data = $('dataDir'); if (data) data.textContent = state?.paths?.dataDir || '—';
  const output = $('outputDir'); if (output) output.textContent = state?.paths?.outputDir || '—';
}


function renderGlobalHeader() {
  const agg = aggregateFinalState();
  const el = $('global-final-status');
  if (!el) return;
  el.className = 'global-status';
  if (agg.state === 'blocked') el.classList.add('danger');
  else if (['needs-rebuild','output-missing'].includes(agg.state)) el.classList.add('attention');
  const dot = ['blocked','needs-rebuild','output-missing'].includes(agg.state) ? '<span class="status-dot"></span>' : '';
  el.innerHTML = `${dot}${escapeHtml(agg.text)}`;
}
function setStepState(id, kind, fallback) {
  const el=$(id); if(!el) return;
  el.className=`step-state ${kind || ''}`;
  el.textContent = kind === 'complete' ? '✓' : (kind === 'attention' ? '!' : fallback);
}
function renderStepBar() {
  const workbooks=workbooksForActivePreset(); const pages=pagesForActivePreset();
  const needs=workbooks.filter(w=>!isLatestPdfWorkbook(w)).length;
  const currentMap={folders:'step-folders-state',excel:'step-excel-state',pages:'step-pages-state',final:'step-final-state'};
  setStepState('step-folders-state', configured()?'complete':'attention','1');
  setStepState('step-excel-state', workbooks.length && needs===0?'complete':(workbooks.length?'attention':''),'2');
  setStepState('step-pages-state', pages.some(p=>p.enabled!==false && String(p.volume||'none')!=='none')?'complete':(pages.length?'attention':''),'3');
  setStepState('step-final-state', finalIsComplete()?'complete':(activeReadinessItems().length?'attention':''),'4');
  document.querySelectorAll('[data-step-view]').forEach(btn=>{
    const view=btn.dataset.stepView; btn.classList.toggle('current',view===activeView);
    const st=$(currentMap[view]); if(view===activeView && st && !st.classList.contains('attention')) st.classList.add('current');
  });
}
function renderNavBadges() {
  const needs=workbooksForActivePreset().filter(w=>!isLatestPdfWorkbook(w)).length;
  const unresolved=unresolvedPageCount();
  const excel=$('nav-excel-count'); if(excel){excel.textContent=needs;excel.classList.toggle('hidden',needs===0);}
  const pages=$('nav-pages-count'); if(pages){pages.textContent=unresolved;pages.classList.toggle('hidden',unresolved===0);}
  const agg=aggregateFinalState(); const dot=$('nav-final-dot'); if(dot) dot.classList.toggle('hidden',!['blocked','needs-rebuild','output-missing'].includes(agg.state));
}

function renderSummary() {
  const summaryEl = $('summary'); if (!summaryEl) return;
  const workbooks=workbooksForActivePreset(); const pages=pagesForActivePreset();
  const needsPdf=workbooks.filter(w=>!isLatestPdfWorkbook(w)).length;
  const cards=[
    {label:'登録済みExcel',value:workbooks.length,unit:'件'},
    {label:'最新PDFあり',value:workbooks.length-needsPdf,unit:'件'},
    {label:'PDF作成が必要',value:needsPdf,unit:'件'},
    {label:'総ページ数',value:pages.length,unit:'ページ'}
  ];
  summaryEl.innerHTML=cards.map(c=>`<article class="stat-card"><div class="stat-label">${escapeHtml(c.label)}</div><div class="stat-value">${c.value}<span class="stat-unit">${escapeHtml(c.unit)}</span></div><div class="stat-sub">${escapeHtml(presetLabel())}カテゴリ</div></article>`).join('');
}


function renderDashboardOverview() {
  const workbooks=workbooksForActivePreset(); const pages=pagesForActivePreset();
  const needsPdf=workbooks.filter(w=>!isLatestPdfWorkbook(w)).length;
  const readiness=activeReadinessItems();
  const blockers=readiness.flatMap(x=>asArray(x.ready.blockers));
  const unassigned=unresolvedPageCount(); const agg=aggregateFinalState();
  const title=$('dashboard-next-title'), caption=$('dashboard-next-caption'), action=$('dashboard-action-btn');
  if(title&&caption&&action){
    if(!configured()){title.textContent='提出フォルダを設定してください';caption.textContent='提出Excelが入っているフォルダを選びます。';action.textContent='提出フォルダを選ぶ';action.dataset.dashboardAction='folder';}
    else if(workbooks.length===0){title.textContent='Excelを登録してください';caption.textContent=`${presetLabel()}のExcelを選んで登録します。`;action.textContent='Excel登録へ進む';action.dataset.dashboardAction='excel';}
    else if(needsPdf>0){title.textContent=`PDF必要分を作成してください（${needsPdf}件）`;caption.textContent='元Excelの最新内容をページPDFへ変換します。';action.textContent='PDF必要分を作成';action.dataset.dashboardAction='render';}
    else if(blockers.length>0){title.textContent='最終PDFを出力できません';caption.textContent=blockers[0]?.message||'Excel登録・PDF作成を確認してください。';action.textContent='Excel登録・PDF作成へ';action.dataset.dashboardAction='excel';}
    else if(unassigned>0){title.textContent=`ページ構成を確認してください（${unassigned}ページ）`;caption.textContent='出力先が未設定のページがあります。';action.textContent='ページ構成へ進む';action.dataset.dashboardAction='pages';}
    else if(['needs-rebuild','output-missing','not-built'].includes(agg.state)){title.textContent=agg.text;caption.textContent='本体・補足の状態を確認して出力してください。';action.textContent='最終PDFへ進む';action.dataset.dashboardAction='final';}
    else{title.textContent='最新の状態です';caption.textContent='最終PDFとページ構成は最新です。';action.textContent='出力フォルダを確認';action.dataset.dashboardAction='final';}
  }
  const cat=$('dashboard-category-summary');if(cat)cat.textContent=`登録済みExcel：${workbooks.length}件 / PDF作成が必要：${needsPdf}件`;
  const activity=$('dashboard-activity');if(!activity)return;
  const rows=[...workbooks].sort((a,b)=>String(latestWorkbookStamp(b)).localeCompare(String(latestWorkbookStamp(a)))).slice(0,5).map(w=>{
    if(String(w.status||'')==='render-error')return{text:`${workbookDisplayName(w)} のPDF作成でエラーがあります`,badge:'作成エラー',cls:'danger',time:latestWorkbookStamp(w)};
    if(isLatestPdfWorkbook(w))return{text:`${workbookDisplayName(w)} のPDFを作成しました`,badge:'PDF作成済み',cls:'neutral',time:w.lastRenderedAt||latestWorkbookStamp(w)};
    return{text:`${workbookDisplayName(w)} はPDF作成が必要です`,badge:'PDF作成が必要',cls:'attention',time:latestWorkbookStamp(w)};
  });
  if(!rows.length)rows.push({text:configured()?`${presetLabel()}のExcelを登録してください`:'ReportBinderへようこそ',badge:configured()?'Excel登録':'お知らせ',cls:'neutral',time:''});
  activity.innerHTML=rows.map(r=>`<div class="activity-row"><span class="activity-text">${escapeHtml(r.text)}</span>${badge(r.badge,r.cls)}<span class="activity-time">${escapeHtml(formatDateTime(r.time))}</span></div>`).join('');
}


function renderFolderOverview() {
  const count=$('folder-file-count'); if(count) count.textContent=`${availableFiles.length || state?.summary?.totalWorkbooks || 0} 件`;
  const scan=$('folder-scan-time'); if(scan) scan.textContent=formatDateTime(state?.structure?.updatedAt||'');
  const result=$('folder-scan-result'); if(result){const errors=Number(state?.summary?.renderErrors||0);result.className=`badge ${errors?'danger':'neutral'}`;result.textContent=errors?'要確認':(configured()?'正常':'未設定');}
}


function renderPageOverview() {
  const counts=pageCountsByVolume(); const strip=$('page-summary-strip');
  if(strip)strip.innerHTML=[['本体',counts.main],['補足',counts.appendix],['出力しない',counts.none]].map(([label,n])=>`<div class="page-summary-item"><div class="summary-label">${label}</div><div><span class="summary-value">${n}</span> ページ</div></div>`).join('');
  const manual=pagesForActivePreset().some(p=>p.orderManual===true); const label=$('sort-state-label'); if(label){label.textContent=manual?'手動':'シート名順';label.classList.toggle('manual',manual);}
  updateBulkSelectionLabel();
}


function renderFinalOverview() {
  const setText=(id,v)=>{const el=$(id);if(el)el.textContent=v;};
  const main=volumeReadiness(mainVolume()), appendix=volumeReadiness(appendixVolume());
  setText('final-main-pages',Number(main.pageCount||0));setText('final-appendix-pages',Number(appendix.pageCount||0));
  const renderFreshness=(ready,containerId,buttonId,fixId)=>{
    const box=$(containerId),btn=$(buttonId),fix=$(fixId);if(!box)return;
    const blockers=asArray(ready.blockers),reasons=asArray(ready.staleReasons).slice(0,3);const display=String(ready.displayState||'not-built');
    let title='最終PDFは未出力',cls='neutral',body='出力すると、このカテゴリ専用の最終PDFを作成します。';
    if(Number(ready.pageCount||0)===0){title='出力ページがありません';body='ページ構成で本体または補足にページを設定してください。';}
    else if(display==='blocked'){title='最終PDFを出力できません';cls='danger';body=`<ul class="blocker-list">${blockers.slice(0,3).map(b=>`<li>${escapeHtml(b.message||'出力条件を確認してください。')}</li>`).join('')}</ul>`;}
    else if(display==='needs-rebuild'){title='再出力が必要';cls='attention';body=reasons.length?`<ul class="reason-list">${reasons.map(r=>`<li><span>${escapeHtml(r.detail||'入力が変更されました')}</span><time>${escapeHtml(formatDateTime(r.at))}</time></li>`).join('')}</ul>`:'入力が変更されています。';}
    else if(display==='output-missing'){title='前回出力が見つかりません';cls='attention';body='ファイルが削除または移動されています。もう一度出力してください。';}
    else if(display==='built'){title='最終PDFは最新です';body=`最終出力 ${escapeHtml(formatDateTime(ready.lastBuiltAt))}`;}
    const dot=cls==='neutral'?'':'<span class="status-dot"></span>';
    box.innerHTML=`<div class="freshness-title ${cls}">${dot}${escapeHtml(title)}</div>${typeof body==='string'&&body.startsWith('<')?body:`<p class="caption">${escapeHtml(body)}</p>`}`;
    if(btn)btn.disabled=Number(ready.pageCount||0)===0||blockers.length>0;
    if(fix)fix.classList.toggle('hidden',blockers.length===0||Number(ready.pageCount||0)===0);
  };
  renderFreshness(main,'final-main-state','build-main-btn','final-main-fix');renderFreshness(appendix,'final-appendix-state','build-appendix-btn','final-appendix-fix');const allBtn=$('build-all-btn');if(allBtn){const targets=[main,appendix].filter(r=>Number(r.pageCount||0)>0);allBtn.disabled=(targets.length===0||targets.some(r=>asArray(r.blockers).filter(b=>String(b.code||'')!=='no-pages').length>0));}
  const history=$('final-history');if(history){const rows=[];for(const [label,r] of [['本体PDF',main],['補足PDF',appendix]])if(r.outputPdf||r.lastBuiltAt)rows.push({label,r});history.innerHTML=rows.length?`<div class="history-row header"><span>出力日時</span><span>種類</span><span>ページ数</span><span>状態</span><span>ファイル</span></div>${rows.map(({label,r})=>`<div class="history-row"><span>${escapeHtml(formatDateTime(r.lastBuiltAt))}</span><span>${label}</span><span>${Number(r.pageCount||0)} ページ</span><span>${badge(r.displayState==='built'?'最新':r.displayState==='output-missing'?'ファイルなし':'再出力必要',r.displayState==='built'?'neutral':'attention')}</span><span class="history-file">${escapeHtml(outputFileName(r.outputPdf))}</span></div>`).join('')}`:'<div class="empty-state">まだ出力履歴はありません。</div>';}
}


function updateBulkSelectionLabel() {
  const el = $('bulk-selected-count');
  if (el) el.textContent = `（選択中：${selectedPages.size}ページ）`;
}

function renderVolumeLinks() {
  for (const [volume,id] of [[mainVolume(),'open-main-link'],[appendixVolume(),'open-appendix-link']]) {
    const a=$(id);if(!a)continue;const r=volumeReadiness(volume);
    if(r.outputPdf && r.outputPdfExists){a.href='#';a.dataset.volume=volume;a.dataset.category=activePreset;a.textContent=r.displayState==='built'?'PDFを開く':'前回出力を開く';a.classList.remove('hidden');}
    else{a.removeAttribute('href');a.classList.add('hidden');}
  }
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
  return name.endsWith('.xlsx');
}
function categoryMatchesName(name, preset = activePreset) {
  const n = String(name || '').toUpperCase();
  if (preset === 'ecm') return n.includes('ECM');
  if (preset === 'bod') return n.includes('BOD');
  if (preset === 'dmm') return n.includes('DMM');
  return false;
}
function presetMatches(file, preset) {
  return isXlsxFile(file);
}
function workbookMatchesPreset(w, preset = activePreset) {
  const category = workbookCategoryValue(w);
  if (['ecm','bod','dmm'].includes(category)) return category === preset;
  return categoryMatchesName(w?.fileName || w?.displayName || w?.relativePath || w?.workbookId, preset);
}
function presetLabel(preset = activePreset) {
  return ({ecm:'ECM', bod:'BOD', dmm:'DMM'})[preset] || 'ECM';
}
function filesForActivePreset(files) {
  // 登録候補は厳密なファイル名判定をしない。提出フォルダ直下の .xlsx だけ表示する。
  return asArray(files).filter(f => isXlsxFile(f));
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
  document.querySelectorAll('[data-preset]').forEach(btn => {
    const active = btn.dataset.preset === activePreset;
    btn.classList.toggle('active', active);
    btn.setAttribute('aria-pressed', active ? 'true' : 'false');
  });
  updateCategoryLabels();
}
function visibleUnregisteredFiles(files) {
  const registered = getRegisteredByRelativePath();
  return filesForActivePreset(files).filter(f => !registered.has(registeredCategoryPathKey(f.relativePath, activePreset)));
}

function capturePageBoardScroll() {
  const snap = {windowX: window.scrollX || 0, windowY: window.scrollY || 0, wraps: {}};
  document.querySelectorAll('tbody[data-volume]').forEach(tbody => {
    const volume = tbody.getAttribute('data-volume') || '';
    const wrap = tbody.closest('.table-wrap');
    if (volume && wrap) snap.wraps[volume] = {top: wrap.scrollTop || 0, left: wrap.scrollLeft || 0};
  });
  return snap;
}
function restorePageBoardScroll(snap) {
  if (!snap) return;
  requestAnimationFrame(() => {
    for (const [volume, pos] of Object.entries(snap.wraps || {})) {
      const tbody = [...document.querySelectorAll('tbody[data-volume]')].find(t => t.getAttribute('data-volume') === volume);
      const wrap = tbody?.closest?.('.table-wrap');
      if (wrap) { wrap.scrollTop = pos.top || 0; wrap.scrollLeft = pos.left || 0; }
    }
    if (Number.isFinite(snap.windowY)) window.scrollTo(snap.windowX || 0, snap.windowY || 0);
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
  updateRenderTargetUi();
}
function handleWorkbookCheckboxToggle(ch, shiftKey=false) {
  lastWorkbookRangeAnchor = updateSelectionRange(visibleWorkbookSelectionValues(), selectedWorkbooks, ch.value, ch.checked, shiftKey, lastWorkbookRangeAnchor);
  syncWorkbookSelectionUi();
}
function visiblePageSelectionValues() {
  return [...document.querySelectorAll('.page-row')].map(row => String(row.getAttribute('data-page-id') || '')).filter(Boolean);
}
function syncPageSelectionUi() {
  updateBulkSelectionLabel();
  document.querySelectorAll('.page-row').forEach(row => {
    const id = String(row.getAttribute('data-page-id') || '');
    const checked = selectedPages.has(id);
    row.classList.toggle('selected-row', checked);
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
function selectRenderNeededWorkbooks() {
  const ids = defaultPdfTargetWorkbookIds();
  selectedWorkbooks = new Set(ids.map(id => String(id || '')).filter(Boolean));
  lastWorkbookRangeAnchor = ids[ids.length - 1] || '';
  syncWorkbookSelectionUi();
  if (ids.length) showMessage('ok', 'PDF必要分を選択しました', `${presetLabel()}のPDF必要分 ${ids.length}件を選択しました。`);
  else showMessage('ok', 'PDF作成が必要なExcelはありません', '登録済みExcelは最新です。');
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
  const ids = pagesForActivePreset().sort(pageSort).map(p => resolvedPageId(p)).filter(Boolean);
  ids.forEach(id => selectedPages.add(id));
  lastPageRangeAnchor = ids[ids.length - 1] || '';
  syncPageSelectionUi();
  if (!ids.length) showMessage('warn', 'ページがありません', '登録済みExcelをPDF作成するとページが表示されます。');
}
function clearActivePageSelection() {
  pagesForActivePreset().map(p => resolvedPageId(p)).filter(Boolean).forEach(id => selectedPages.delete(id));
  lastPageRangeAnchor = '';
  syncPageSelectionUi();
}

async function moveSelectedPagesToVolume(volume, btn) {
  const target=String(volume||'');const ids=[...selectedPages].map(String).filter(id=>!!getPage(id));
  if(!ids.length){showMessage('warn','ページを選択してください','移動するページにチェックを入れてください。');return;}
  await runBusy(btn,async()=>{
    const main=mainVolume(),appendix=appendixVolume();const volumes={[main]:[],[appendix]:[],none:[]};const selected=new Set(ids);
    for(const p of [...pagesForActivePreset()].sort(pageSort)){const id=resolvedPageId(p);if(!id||selected.has(id))continue;const v=(p.enabled===false||String(p.volume||main)==='none')?'none':String(p.volume||main);if(!volumes[v])volumes[v]=[];volumes[v].push(id);}
    if(!volumes[target])volumes[target]=[];volumes[target].push(...ids);
    await api('/api/pages/reorder',{method:'POST',body:{category:activePreset,volumes}});selectedPages.clear();lastPageRangeAnchor='';lastPageBoardRenderSignature='';await refresh();
    showMessage('ok','ページ構成を保存しました',`${ids.length}ページを${volumeLabel(target)}に設定しました。`,null,[{label:'最終PDFへ',view:'final'}]);
  });
}

async function loadFiles(btn) {
  if (!configured()) {
    showMessage('warn', '提出フォルダを選んでください', '提出フォルダを選ぶとExcelを表示できます。');
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
    showMessage(availableFiles.length ? 'ok' : 'warn', availableFiles.length ? 'Excel一覧を更新しました' : 'Excelが見つかりません', '提出フォルダ直下のExcelだけを表示しています。');
  }, false);
  return availableFiles;
}
function renderFileList(files) {
  const box = $('file-list');
  if (!box) return;
  updatePresetTabs();
  if (!configured()) {
    box.className = 'table-shell empty-state';
    box.textContent = '提出フォルダを選んでください。';
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
    box.innerHTML = allSelectable.length ? '<div><strong>検索条件に一致する未登録Excelはありません。</strong></div>' : '<div><strong>未登録のExcelはありません。</strong></div>';
    if (!allSelectable.length) { selectedFiles.clear(); lastFileRangeAnchor = ''; }
    setFileSelectionSummary(allSelectable.length, selectedCount());
    return;
  }
  selectedFiles = new Set([...selectedFiles].filter(id => allSelectable.some(f => String(f.relativePath || '') === id)));
  const allSelected = selectable.length > 0 && selectable.every(f => selectedFiles.has(String(f.relativePath || '')));
  const someSelected = selectable.some(f => selectedFiles.has(String(f.relativePath || '')));
  box.className = 'table-shell';
  box.innerHTML = `<table class="data-table file-table">
    <thead><tr><th class="check-col"><input type="checkbox" data-select-all-files ${allSelected ? 'checked' : ''} aria-label="表示中の未登録Excelをすべて選択"></th><th>ファイル名</th></tr></thead>
    <tbody>${selectable.map(f => {
      const rel = String(f.relativePath || '');
      const checked = selectedFiles.has(rel);
      const displayName = fileDisplayName(f);
      const detail = normPath(rel) !== normPath(displayName) ? `<div class="subtext">${escapeHtml(rel)}</div>` : '';
      const modified = formatSubmissionFileModifiedAt(f);
      return `<tr data-file-row class="${checked ? 'selected-row' : ''}">
        <td class="check-col"><input type="checkbox" data-file-check value="${escapeAttr(rel)}" ${checked ? 'checked' : ''} aria-label="${escapeAttr(displayName)}を選択"></td>
        <td><div class="file-name-cell"><span class="file-icon excel">${iconUse('i-file-excel')}</span><div><strong title="${escapeAttr(displayName)}">${escapeHtml(displayName)}</strong>${detail}<div class="file-meta" title="Excelファイルの最終保存日時">更新日時：${escapeHtml(modified)}</div></div></div></td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;
  setFileSelectionSummary(allSelectable.length, selectedCount());
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
  const preset=typeof presetOrButton==='string'?presetOrButton:(presetOrButton?.dataset?.preset||activePreset);
  if(!['ecm','bod','dmm'].includes(preset))return;
  activePreset=preset;try{sessionStorage.setItem('ReportBinderCategory',activePreset);}catch{}
  selectedFiles.clear();selectedWorkbooks.clear();selectedPages.clear();lastFileRangeAnchor='';lastWorkbookRangeAnchor='';lastPageRangeAnchor='';lastPageBoardRenderSignature='';
  if(configured()&&!availableFiles.length)await loadFilesSilently();
  renderAll();
}

async function registerSelected(btn) {
  const rels=[...selectedFiles];if(!rels.length){showMessage('warn','Excelを選択してください','登録するExcelにチェックを入れてください。');return;}
  await runBusy(btn,async()=>{
    const data=await api('/api/workbooks/register-batch',{method:'POST',body:{relativePaths:rels,category:activePreset}});const result=data.result||{},registered=asArray(result.registered),errors=asArray(result.errors);
    selectedFiles.clear();lastFileRangeAnchor='';if(data.state)state=normalizeStatePayload(data.state);else await refresh();
    selectedWorkbooks=new Set(registered.map(x=>String(x.workbookId||'')).filter(Boolean));lastWorkbookRangeAnchor='';renderAll();await loadFiles(null);
    if(registered.length&&errors.length===0)showMessage('ok',`${registered.length}件を登録しました`,`登録した${registered.length}件が選択されています。`,null,[{label:`登録した${registered.length}件をPDF作成`,primary:true,view:'excel',handler:()=>renderSelectedWorkbooks($('render-selected-btn'))}],0);
    else if(registered.length)showMessage('warn',`${registered.length}件を登録しました`,`${errors.length}件は登録できませんでした。`,errors,[{label:'登録分をPDF作成',view:'excel',handler:()=>renderSelectedWorkbooks($('render-selected-btn'))}],0);
    else showMessage('warn','登録できませんでした','詳細を確認してください。',errors.length?errors:result,[],0);
  });
}


async function unregisterSelected(btn) {
  const ids = [...selectedWorkbooks];
  if (!ids.length) { showMessage('warn', 'Excelを選択してください', '登録解除するExcelにチェックを入れてください。'); return; }
  if (!confirm(`${ids.length}件の登録を解除します。PDFファイル自体は削除しません。`)) return;
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
// 表現は「見落としなし」ではなく「最終PDFの見た目を基準とした高精度な判定」。
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
function changeBadgeForWorkbook(w){
  if (!state?.inputHistoryEnabled) return '';
  const cs = changeSummaryFor(w?.workbookId);
  if (!cs) return '';
  const actionBadge=(text,cls='neutral')=>`<button class="badge change-action ${cls}" type="button" data-open-diff="${escapeAttr(w?.workbookId||'')}" aria-label="${escapeAttr(`${workbookDisplayName(w)}、${text}、差分詳細を開く`)}">${escapeHtml(text)}</button>`;
  if (String(cs.status||'') !== 'complete') return actionBadge('比較不能','neutral');
  const addedNames=new Set(asArray(cs.addedSheets).map(String));
  const removedNames=new Set(asArray(cs.removedSheets).map(String));
  const unknownNames=new Set(asArray(cs.unknownSheets).map(String));
  const changed = asArray(cs.changedSheets).map(String).filter(name=>!addedNames.has(name)&&!removedNames.has(name)&&!unknownNames.has(name)).length;
  const added = addedNames.size;
  const unknown = unknownNames.size;
  // Added and removed sheets are changes too. Keep them visible in the workbook
  // summary even though only a current (added) sheet can have a page row.
  const removed = removedNames.size;
  if (changed > 0 || added > 0 || removed > 0) {
    const parts = [];
    if (changed > 0) parts.push(`変更 ${changed}`);
    if (added > 0) parts.push(`追加 ${added}`);
    if (removed > 0) parts.push(`削除 ${removed}`);
    return actionBadge(`${parts.join(' / ')}シート`,'attention');
  }
  if (unknown > 0) return actionBadge(`判定不能 ${unknown}`,'neutral');
  return actionBadge('変更なし','neutral');
}
function workbookPdfStatusCell(w){
  const status = String(w?.status || '');
  const detail = latestPdfDetailForWorkbook(w);
  if (status === 'render-error') {
    return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('PDF作成エラー','danger')}</div><div class="pdf-status-detail">${escapeHtml(detail || '差分を確認できません')}</div></div>`;
  }
  if (status === 'rendering') {
    return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('PDF作成中','attention')}</div><div class="pdf-status-detail">完了後に差分を判定します</div></div>`;
  }
  if (!isLatestPdfWorkbook(w)) {
    return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('PDF作成が必要','attention')}</div><div class="pdf-status-detail">差分は作成後に確認します</div></div>`;
  }
  const change = changeBadgeForWorkbook(w);
  return `<div class="pdf-status-stack"><div class="pdf-status-line">${badge('最新PDF','neutral')}${change}</div>${change?'<div class="pdf-status-detail">前回PDFとの差分</div>':''}</div>`;
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

// ---- V5: 差分詳細ダイアログ ----
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
  if(diffViewState.filter==='all')return sheets;
  return sheets.filter(s=>String(s.kind||'')===diffViewState.filter);
}
function diffCurrentSheet(){
  const sheets=asArray(diffViewState.detail?.sheets);
  return sheets.find(s=>String(s.sheetKey||'')===diffViewState.selectedSheetKey)||null;
}
function diffCurrentPage(){
  const sheet=diffCurrentSheet();
  return asArray(sheet?.pages)[diffViewState.pageIndex]||null;
}
function diffAssetUrl(sheet,page,asset){
  const cmp=diffViewState.detail?.comparison||{};
  const params=new URLSearchParams({
    workbookId:diffViewState.workbookId,
    currentSnapshotId:String(cmp.currentSnapshotId||''),
    baselineSnapshotId:String(cmp.baselineSnapshotId||''),
    sheetKey:String(sheet?.sheetKey||''),
    pageNumber:String(page?.pageNumber||1),
    asset,
    scope:String(cmp.scope||'automatic')
  });
  return withToken(`/api/history/diff-page?${params.toString()}`);
}
function diffDetailRequestPath(){
  const params=new URLSearchParams({workbookId:diffViewState.workbookId});
  if(diffViewState.fromSnapshotId&&diffViewState.toSnapshotId){
    params.set('fromSnapshotId',diffViewState.fromSnapshotId);
    params.set('toSnapshotId',diffViewState.toSnapshotId);
  }
  return `/api/history/diff-detail?${params.toString()}`;
}
function diffPrepareRequestBody(sheetKey=''){
  const body={workbookId:diffViewState.workbookId};
  if(diffViewState.fromSnapshotId&&diffViewState.toSnapshotId){
    body.fromSnapshotId=diffViewState.fromSnapshotId;
    body.toSnapshotId=diffViewState.toSnapshotId;
  }
  if(sheetKey)body.sheetKey=sheetKey;
  return body;
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
    ['unchanged','変更なし',Number(summary.unchanged||0)]
  ];
  box.innerHTML=definitions.map(([key,label,count])=>`<button class="${escapeAttr(key)} ${diffViewState.filter===key?'active':''}" type="button" data-diff-filter="${escapeAttr(key)}">${escapeHtml(label)} ${count}</button>`).join('');
  box.querySelectorAll('[data-diff-filter]').forEach(btn=>btn.addEventListener('click',()=>{
    diffViewState.filter=btn.dataset.diffFilter||'all';
    const visible=diffFilteredSheets();
    if(!visible.some(s=>String(s.sheetKey||'')===diffViewState.selectedSheetKey)){
      diffViewState.selectedSheetKey=String(visible[0]?.sheetKey||'');
      diffViewState.pageIndex=0;diffViewState.regionIndex=-1;
    }
    renderDiffSummary();renderDiffSheetList();renderDiffPage();
    const selected=visible.find(s=>String(s.sheetKey||'')===diffViewState.selectedSheetKey);
    if(selected&&['deferred','failed'].includes(String(selected.status||'')))void selectDiffSheet(selected.sheetKey,false);
  }));
}
function renderDiffSheetList(){
  const box=$('diff-sheet-list');if(!box)return;
  const sheets=diffFilteredSheets();
  if(!sheets.length){box.innerHTML='<div class="empty-state">対象のシートはありません。</div>';return;}
  box.innerHTML=sheets.map(sheet=>{
    const meta=diffKindMeta(sheet.kind);
    const pages=Math.max(Number(sheet.beforePages||0),Number(sheet.afterPages||0),asArray(sheet.pages).length);
    const sheetStatus=String(sheet.status||'');
    const detail=sheetStatus==='unknown'?(sheet.message||'判定不能'):sheetStatus==='failed'?'作成失敗・再選択で再試行':sheetStatus==='deferred'?'選択時に画像を作成':sheetStatus==='generating'?'画像を作成中':`${pages}ページ・${Number(sheet.regionCount||0)}領域`;
    return `<button class="diff-sheet-item ${String(sheet.sheetKey||'')===diffViewState.selectedSheetKey?'active':''}" type="button" data-diff-sheet="${escapeAttr(sheet.sheetKey||'')}" title="${escapeAttr(sheet.sheetName||'')}"><span class="diff-sheet-kind ${escapeAttr(sheet.kind||'unknown')}">${escapeHtml(meta.mark)}</span><span class="diff-sheet-copy"><strong>${escapeHtml(sheet.sheetName||'')}</strong><small>${escapeHtml(meta.label)}・${escapeHtml(detail)}</small></span></button>`;
  }).join('');
  box.querySelectorAll('[data-diff-sheet]').forEach(btn=>{
    btn.addEventListener('click',()=>selectDiffSheet(btn.dataset.diffSheet||''));
    btn.addEventListener('keydown',e=>{
      if(!['ArrowDown','ArrowUp'].includes(e.key))return;
      e.preventDefault();
      const buttons=[...box.querySelectorAll('[data-diff-sheet]')];
      const index=buttons.indexOf(btn);
      const next=buttons[Math.max(0,Math.min(buttons.length-1,index+(e.key==='ArrowDown'?1:-1)))];
      next?.focus();if(next)selectDiffSheet(next.dataset.diffSheet||'',false);
    });
  });
}
const diffSheetPrepareInFlight=new Set();
async function selectDiffSheet(sheetKey,focusList=true){
  const sheet=asArray(diffViewState.detail?.sheets).find(s=>String(s.sheetKey||'')===String(sheetKey||''));
  if(!sheet)return;
  diffViewState.selectedSheetKey=String(sheet.sheetKey||'');
  diffViewState.pageIndex=0;
  diffViewState.regionIndex=-1;
  sheet.confirmed=true;
  renderDiffSheetList();
  renderDiffPage();
  if(focusList)$('diff-before-viewport')?.focus();
  if(['deferred','failed'].includes(String(sheet.status||''))&&!diffSheetPrepareInFlight.has(String(sheet.sheetKey||''))){
    const key=String(sheet.sheetKey||'');
    diffSheetPrepareInFlight.add(key);
    sheet.status='generating';sheet.message='このシートの画像を準備しています。';
    renderDiffSheetList();renderDiffPage();
    try{await prepareDiffDetail(key);}finally{diffSheetPrepareInFlight.delete(key);}
  }
}
function setDiffImage(id,src,alt,onload){
  const img=$(id);if(!img)return;
  img.alt=alt||'';
  img.onload=()=>{if(typeof onload==='function')onload();};
  img.onerror=()=>{img.removeAttribute('src');};
  if(src)img.src=src;else img.removeAttribute('src');
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
  document.querySelectorAll('#diff-modal .diff-image-mask').forEach(el=>{el.style.display=enabled?'block':'none';el.style.opacity=String(opacity);});
  document.querySelectorAll('#diff-modal .diff-image-overlay').forEach(el=>{el.style.display=enabled?'block':'none';});
}
function applyDiffMode(){
  const overlay=diffViewState.mode==='overlay';
  $('diff-viewers')?.classList.toggle('overlay-mode',overlay);
  $('diff-mobile-tabs')?.classList.toggle('hidden',overlay);
  $('diff-mode-side')?.classList.toggle('active',!overlay);
  $('diff-mode-overlay')?.classList.toggle('active',overlay);
  $('diff-opacity-wrap')?.classList.toggle('hidden',!overlay);
  $('diff-after-underlay')?.classList.toggle('hidden',!overlay);
  const opacity=Math.max(0,Math.min(100,Number($('diff-opacity')?.value||50)));
  if($('diff-after-base'))$('diff-after-base').style.opacity=overlay?String(opacity/100):'1';
  if($('diff-opacity-value'))$('diff-opacity-value').textContent=`${opacity}%`;
  requestAnimationFrame(updateDiffStageScale);
}
function renderDiffPage(){
  const sheet=diffCurrentSheet();
  const historical=String(diffViewState.detail?.comparison?.scope||'')==='history';
  const beforeLabel=historical?'比較元':'前回版';
  const afterLabel=historical?'比較先':'現在版';
  const messageBox=$('diff-state-message');
  if(!sheet){
    if(messageBox){messageBox.textContent=diffViewState.detail?.message||'比較対象のシートがありません。';messageBox.classList.remove('hidden');}
    setDiffPaneEmpty('before','比較するページがありません。');setDiffPaneEmpty('after','比較するページがありません。');
    return;
  }
  const pages=asArray(sheet.pages);
  diffViewState.pageIndex=Math.max(0,Math.min(diffViewState.pageIndex,Math.max(0,pages.length-1)));
  const page=pages[diffViewState.pageIndex]||null;
  const sheetStatus=String(sheet.status||'');
  const sheetMessage=sheetStatus==='unknown'?String(sheet.message||'信頼できる差分領域を判定できません。'):sheetStatus==='failed'?String(sheet.message||'差分画像を作成できませんでした。シートを再選択して再試行してください。'):'';
  const pageMessage=String(page?.message||'');
  const sizeMessage=page?.pageSizeChanged?'ページサイズが異なるため、左上を基準に揃えて表示しています。':'';
  const visibleMessage=[sheetMessage,pageMessage,sizeMessage].filter(Boolean).join(' ');
  if(messageBox){messageBox.textContent=visibleMessage;messageBox.classList.toggle('hidden',!visibleMessage);}
  if(!page){
    const pending=sheetStatus==='deferred'?'このシートを選択すると画像を作成します。':sheetStatus==='failed'?(sheet.message||'差分画像を作成できませんでした。シートを再選択して再試行してください。'):sheetStatus==='generating'||String(diffViewState.detail?.status||'')==='generating'?'差分画像を準備しています。':'表示できるページ画像がありません。';
    setDiffPaneEmpty('before',sheet.kind==='added'?`${beforeLabel}には存在しません`:pending);
    setDiffPaneEmpty('after',sheet.kind==='removed'?`${afterLabel}では削除されています`:pending);
    if($('diff-page-count'))$('diff-page-count').textContent=`0 / ${Math.max(Number(sheet.pageCount||0),0)}ページ`;
    if($('diff-region-count'))$('diff-region-count').textContent='0件';
    return;
  }
  const beforeEmpty=sheet.kind==='added'?`${beforeLabel}には存在しません`:'';
  const afterEmpty=sheet.kind==='removed'?`${afterLabel}では削除されています`:'';
  setDiffPaneEmpty('before',beforeEmpty);setDiffPaneEmpty('after',afterEmpty);
  const pageNo=Number(page.pageNumber||diffViewState.pageIndex+1);
  const pageTotal=pages.length;
  if($('diff-before-page-label'))$('diff-before-page-label').textContent=`${pageNo} / ${pageTotal}ページ`;
  if($('diff-after-page-label'))$('diff-after-page-label').textContent=`${pageNo} / ${pageTotal}ページ`;
  if($('diff-page-count'))$('diff-page-count').textContent=`${pageNo} / ${pageTotal}ページ`;
  $('diff-prev-page').disabled=diffViewState.pageIndex<=0;
  $('diff-next-page').disabled=diffViewState.pageIndex>=pageTotal-1;
  const altBase=`${sheet.sheetName}、${pageNo}ページ`;
  const beforeUrl=page.beforeAsset?diffAssetUrl(sheet,page,'before'):'';
  const afterUrl=page.afterAsset?diffAssetUrl(sheet,page,'after'):'';
  setDiffImage('diff-before-base',beforeEmpty?'':beforeUrl,`${beforeLabel}、${altBase}`,updateDiffStageScale);
  setDiffImage('diff-after-underlay',beforeUrl,'');
  setDiffImage('diff-after-base',afterEmpty?'':afterUrl,`${afterLabel}、${altBase}`,updateDiffStageScale);
  setDiffImage('diff-before-mask',page.beforeMaskAsset?diffAssetUrl(sheet,page,'before-mask'):'','');
  setDiffImage('diff-before-overlay',page.beforeOverlayAsset?diffAssetUrl(sheet,page,'before-overlay'):'','');
  setDiffImage('diff-after-mask',page.maskAsset?diffAssetUrl(sheet,page,'mask'):'','');
  setDiffImage('diff-after-overlay',page.overlayAsset?diffAssetUrl(sheet,page,'overlay'):'','');
  const regions=asArray(page.regions);
  if(diffViewState.regionIndex>=regions.length)diffViewState.regionIndex=-1;
  if($('diff-region-count'))$('diff-region-count').textContent=regions.length?(diffViewState.regionIndex>=0?`${diffViewState.regionIndex+1} / ${regions.length}件`:`${regions.length}件`):'0件';
  $('diff-prev-region').disabled=!diffRegionTargets().length;
  $('diff-next-region').disabled=!diffRegionTargets().length;
  applyDiffHighlightSettings();applyDiffMode();updateDiffStageScale();
  prefetchAdjacentDiffPage(sheet,diffViewState.pageIndex+1);
}
function prefetchAdjacentDiffPage(sheet,index){
  const page=asArray(sheet?.pages)[index];if(!page)return;
  for(const asset of ['before','after']){const img=new Image();img.src=diffAssetUrl(sheet,page,asset);}
}
function focusCurrentDiffRegion(){
  const page=diffCurrentPage(),region=asArray(page?.regions)[diffViewState.regionIndex]||null;
  for(const side of ['before','after']){
    const focus=$(`diff-${side}-focus`);
    if(!focus)continue;
    focus.classList.toggle('hidden',!region);
    if(!region)continue;
    focus.style.left=`${Number(region.x||0)*100}%`;
    focus.style.top=`${Number(region.y||0)*100}%`;
    focus.style.width=`${Number(region.width||0)*100}%`;
    focus.style.height=`${Number(region.height||0)*100}%`;
    const stage=$(`diff-${side}-stage`),viewport=$(`diff-${side}-viewport`);
    if(stage&&viewport&&viewport.offsetParent!==null){
      const centerX=(Number(region.x||0)+Number(region.width||0)/2)*stage.offsetWidth;
      const centerY=(Number(region.y||0)+Number(region.height||0)/2)*stage.offsetHeight;
      viewport.scrollTo({left:Math.max(0,centerX-viewport.clientWidth/2),top:Math.max(0,centerY-viewport.clientHeight/2),behavior:'smooth'});
    }
  }
}
function diffRegionTargets(){
  const targets=[];
  for(const sheet of asArray(diffViewState.detail?.sheets)){
    asArray(sheet.pages).forEach((page,pageIndex)=>asArray(page.regions).forEach((region,regionIndex)=>targets.push({sheetKey:String(sheet.sheetKey||''),pageIndex,regionIndex,region})));
  }
  return targets;
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
  const pages=asArray(diffCurrentSheet()?.pages);
  if(!pages.length)return;
  diffViewState.pageIndex=Math.max(0,Math.min(pages.length-1,index));
  diffViewState.regionIndex=-1;
  renderDiffPage();
}
function renderDiffDetail(detail){
  diffViewState.detail=detail||null;
  const stateBox=$('diff-state-message');
  if(stateBox){stateBox.textContent='';stateBox.classList.add('hidden');}
  const workbookName=String(detail?.workbookName||getWorkbook(diffViewState.workbookId)?.displayName||'Excel');
  if($('diff-title'))$('diff-title').textContent=`差分詳細：${workbookName}`;
  const cmp=detail?.comparison||{};
  const historical=String(cmp.scope||'')==='history';
  if($('diff-subtitle'))$('diff-subtitle').textContent=`${historical?'比較元':'前回版'} ${formatDateTime(cmp.baselineAt)} → ${historical?'比較先':'現在版'} ${formatDateTime(cmp.currentAt)}`;
  if($('diff-before-label'))$('diff-before-label').textContent=historical?'比較元':'前回版';
  if($('diff-after-label'))$('diff-after-label').textContent=historical?'比較先':'現在版';
  const tabs=$('diff-mobile-tabs')?.querySelectorAll('button');
  if(tabs?.[0])tabs[0].textContent=historical?'比較元':'前回版';
  if(tabs?.[1])tabs[1].textContent=historical?'比較先':'現在版';
  updateDiffProgress(detail?.generation||{});
  const sheets=asArray(detail?.sheets);
  if(!sheets.some(s=>String(s.sheetKey||'')===diffViewState.selectedSheetKey)){
    const preferred=sheets.find(s=>['modified','added','removed','unknown'].includes(String(s.kind||'')))||sheets[0];
    diffViewState.selectedSheetKey=String(preferred?.sheetKey||'');
    diffViewState.pageIndex=0;diffViewState.regionIndex=-1;
  }
  renderDiffSummary();renderDiffSheetList();renderDiffPage();
  if(['unavailable','failed'].includes(String(detail?.status||''))){
    const box=$('diff-state-message');
    if(box){
      const retry=String(detail?.status||'')==='failed';
      box.innerHTML=`<span>${escapeHtml(detail?.message||'差分詳細を表示できません。')}</span>${retry?' <button class="btn secondary compact" type="button" data-diff-retry>再試行</button>':''}`;
      box.classList.remove('hidden');
      box.querySelector('[data-diff-retry]')?.addEventListener('click',()=>prepareDiffDetail());
    }
  }
}
async function pollDiffJob(jobId,tokenId){
  while(isDiffModalOpen()&&tokenId===diffPollToken){
    const job=await api(`/api/jobs/status?jobId=${encodeURIComponent(jobId)}`);
    updateDiffProgress(job);
    const status=String(job.status||'');
    if(['completed','completed-with-errors','failed','cancelled','missing'].includes(status)){
      const loaded=await api(diffDetailRequestPath());
      if(tokenId===diffPollToken&&isDiffModalOpen())renderDiffDetail(loaded.detail);
      return loaded.detail||null;
    }
    await sleep(700);
  }
  return null;
}
async function prepareDiffDetail(sheetKey=''){
  if(!diffViewState.workbookId)return;
  try{
    // Another lazy sheet or the initial job may already own the pair lock. In that case
    // wait for it, then retry this requested sheet instead of leaving it deferred.
    for(let attempt=0;attempt<3;attempt++){
      const started=await api('/api/history/diff/prepare',{method:'POST',body:diffPrepareRequestBody(sheetKey)});
      const job=started.job||{};
      const joinedExisting=!!job.joinedExistingDiffJob;
      let detail=null;
      if(job.jobId){
        const tokenId=++diffPollToken;
        updateDiffProgress(job);
        detail=await pollDiffJob(job.jobId,tokenId);
      }else{
        const loaded=await api(diffDetailRequestPath());
        detail=loaded.detail||null;
        renderDiffDetail(detail);
      }
      if(!sheetKey)return;
      const target=asArray(detail?.sheets).find(s=>String(s.sheetKey||'')===String(sheetKey));
      const targetStatus=String(target?.status||'');
      if(!target||!['deferred','failed'].includes(targetStatus))return;
      // A failure from the job we started is final for this click. Retry only when
      // we merely joined another in-flight job that did not complete this sheet.
      if(targetStatus==='failed'&&!joinedExisting)return;
      if(attempt<2)await sleep(150);
    }
    throw new Error('選択したシートの画像作成を開始できませんでした。もう一度選択してください。');
  }catch(e){
    renderDiffDetail(Object.assign({},diffViewState.detail||{},{status:'failed',message:userFriendlyError(e.message),generation:{status:'failed'}}));
  }
}
async function openDiffDetail(workbookId,opener,historyRange=null){
  const id=String(workbookId||'');if(!id)return;
  diffReturnFocus=opener||document.activeElement;
  diffViewState.workbookId=id;
  diffViewState.fromSnapshotId=String(historyRange?.fromSnapshotId||'');
  diffViewState.toSnapshotId=String(historyRange?.toSnapshotId||'');
  diffViewState.detail=null;diffViewState.selectedSheetKey='';diffViewState.pageIndex=0;diffViewState.regionIndex=-1;diffViewState.filter='all';diffViewState.mode='side';diffViewState.mobileTab='before';
  $('diff-modal')?.classList.remove('hidden');
  document.body.style.overflow='hidden';
  if($('diff-title'))$('diff-title').textContent=`差分詳細：${workbookDisplayName(getWorkbook(id))}`;
  if($('diff-subtitle'))$('diff-subtitle').textContent='比較情報を読み込んでいます。';
  if($('diff-sheet-list'))$('diff-sheet-list').innerHTML='<div class="empty-state">読み込んでいます。</div>';
  setDiffPaneEmpty('before','比較情報を読み込んでいます。');setDiffPaneEmpty('after','比較情報を読み込んでいます。');
  $('diff-close')?.focus();
  try{
    const loaded=await api(diffDetailRequestPath());
    if(!isDiffModalOpen()||diffViewState.workbookId!==id)return;
    renderDiffDetail(loaded.detail);
    const status=String(loaded.detail?.status||'');
    if(status==='not-generated')await prepareDiffDetail();
    else if(status==='generating'&&loaded.detail?.generation?.jobId){
      const tokenId=++diffPollToken;
      await pollDiffJob(loaded.detail.generation.jobId,tokenId);
    }
  }catch(e){
    renderDiffDetail({status:'failed',message:userFriendlyError(e.message),workbookName:workbookDisplayName(getWorkbook(id)),sheets:[],summary:{},generation:{status:'failed'}});
  }
}
function closeDiffDetail(){
  if(!isDiffModalOpen())return;
  diffPollToken++;
  $('diff-modal')?.classList.add('hidden');
  document.body.style.overflow='';
  let target=diffReturnFocus;
  if(!target?.isConnected){
    const workbookId=String(diffViewState.workbookId||'');
    target=[...document.querySelectorAll('[data-open-diff]')].find(el=>String(el.dataset.openDiff||'')===workbookId)||null;
  }
  diffReturnFocus=null;
  if(target&&typeof target.focus==='function')target.focus();
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
  const list=workbooksForActivePreset();const box=$('workbook-list');if(!box)return;const summary=$('workbook-selection-summary');
  if(!list.length){box.className='table-shell empty-state';box.textContent=`${presetLabel()}の登録済みExcelはまだありません。`;selectedWorkbooks.clear();lastWorkbookRangeAnchor='';if(summary)summary.textContent='0件中 0件を選択';updateRenderTargetUi();return;}
  selectedWorkbooks=new Set([...selectedWorkbooks].filter(id=>list.some(w=>String(w.workbookId)===String(id))));const allSelected=list.every(w=>selectedWorkbooks.has(String(w.workbookId||''))),someSelected=list.some(w=>selectedWorkbooks.has(String(w.workbookId||'')));
  box.className='table-shell';box.innerHTML=`<table class="data-table workbook-table"><thead><tr><th class="check-col"><input type="checkbox" data-select-all-workbooks ${allSelected?'checked':''}></th><th>ファイル名</th><th class="pdf-status-col">PDF状況</th></tr></thead><tbody>${list.map(w=>{const id=String(w.workbookId||''),checked=selectedWorkbooks.has(id),renderedAt=formatDateTime(w.lastRenderedAt||'');return `<tr data-workbook-row class="${checked?'selected-row':''} ${w.status==='render-error'?'error-row':''}"><td class="check-col"><input type="checkbox" data-workbook-check value="${escapeAttr(id)}" ${checked?'checked':''}></td><td><div class="file-name-cell"><span class="file-icon excel">${iconUse('i-file-excel')}</span><div><strong title="${escapeAttr(workbookDisplayName(w))}">${escapeHtml(workbookDisplayName(w))}</strong><div class="file-meta" title="最後にページPDFを作成した日時">PDF作成日時：${escapeHtml(renderedAt)}</div></div></div></td><td class="pdf-status-col">${workbookPdfStatusCell(w)}</td></tr>`;}).join('')}</tbody></table>`;
  if(summary)summary.textContent=`${list.length}件中 ${selectedWorkbooks.size}件を選択`;const selectAll=box.querySelector('[data-select-all-workbooks]');if(selectAll){selectAll.indeterminate=someSelected&&!allSelected;selectAll.addEventListener('change',e=>{selectedWorkbooks=e.target.checked?new Set(list.map(w=>String(w.workbookId||'')).filter(Boolean)):new Set();lastWorkbookRangeAnchor='';syncWorkbookSelectionUi();renderWorkbooks();});}
  box.querySelectorAll('[data-workbook-check]').forEach(ch=>ch.addEventListener('click',e=>{e.stopPropagation();handleWorkbookCheckboxToggle(ch,e.shiftKey);renderWorkbooks();}));box.querySelectorAll('[data-open-diff]').forEach(btn=>btn.addEventListener('click',e=>{e.stopPropagation();openDiffDetail(btn.dataset.openDiff||'',btn);}));box.querySelectorAll('[data-workbook-row]').forEach(row=>row.addEventListener('click',e=>{if(e.target.closest('input,button,a,select,textarea'))return;const ch=row.querySelector('[data-workbook-check]');if(!ch)return;ch.checked=!ch.checked;handleWorkbookCheckboxToggle(ch,e.shiftKey);renderWorkbooks();}));updateRenderTargetUi();
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
    const more = warnings.length > 6 ? `\n…ほか ${warnings.length - 6} 件(詳細はログ参照)` : '';
    return {kind:'warn', title:'PDFを作成しました(要確認)', message: shown + more, detail:list};
  }
  if (list.length) return {kind:'ok', title:'PDFを作成しました', message:`${list.length}件のExcelを処理しました。`, detail:list};
  return {kind:'ok', title:'PDF作成は不要です', message:'対象はありません。', detail:null};
}
async function renderWorkbookIds(ids, btn, options={}) {
  const workbookIds=normalizeWorkbookIdListForActivePreset(ids);const onlyUpdated=!!options.onlyUpdated&&!workbookIds.length;const activeWorkbooks=workbooksForActivePreset().filter(w=>String(w.status||'')!=='missing');
  if(!workbookIds.length&&!onlyUpdated){hideProgressPanel();showMessage(activeWorkbooks.length?'ok':'warn',activeWorkbooks.length?'PDF作成が必要なExcelはありません':'登録済みExcelがありません',activeWorkbooks.length?'登録済みExcelは最新です。':'先にExcelを登録してください。');return;}
  renderJobActive=true;const estimateTotal=workbookIds.length||defaultPdfTargetWorkbookIds().length||(onlyUpdated?activeWorkbooks.length:0);const startMessage=options.startMessage||(onlyUpdated?'PDF作成対象を確認しています。':`${workbookIds.length}件のExcelをPDF作成します。`);
  updateProgressPanel({status:'queued',total:estimateTotal,completed:0,failed:0,percent:0,message:startMessage});
  try{await runBusy(btn,async()=>{hideErrorPanel();showMessage('','PDF作成を開始しています',startMessage);const body={category:activePreset};if(workbookIds.length)body.workbookIds=workbookIds;if(onlyUpdated)body.onlyUpdated=true;
    const started=await api('/api/workbooks/render/start',{method:'POST',body});const job=normalizeRenderJobFromStart(started);const jobId=assertJobId(job,started);updateProgressPanel(Object.assign({total:estimateTotal,completed:0,failed:0},job));const finished=await waitForRenderJob(jobId);const appendedIds=asArray(finished.results).flatMap(r=>asArray(r?.sheetSync?.insertedAtEndPageIds)).map(String).filter(Boolean);setHighlightedPages(appendedIds);await refresh();renderAll();await loadFiles(null);
    const msg=summarizeRenderResults(finished.results||[]);const hasErrors=finished.status==='failed'||finished.status==='completed-with-errors'||asArray(finished.errors).length>0;
    if(finished.status==='failed')showMessage('danger','PDF作成が停止しました',userFriendlyError(finished.message||'PDF作成を完了できませんでした。'),finished.errors||finished);
    else if(asArray(finished.errors).length)showMessage('danger','PDF作成でエラーがあります',`${finished.failed||finished.errors.length}件のExcelでPDFを作成できませんでした。`,finished.errors);
    else{const endCount=asArray(finished.results).reduce((n,r)=>n+Number(r?.sheetSync?.insertedAtEndCount||0),0);const suffix=endCount?`\n新しい${endCount}ページを末尾に追加しました。`:'';const actions=[{label:'ページ構成を確認',primary:true,view:'pages',handler:endCount?scrollToHighlightedPages:null}];if(endCount)actions.push({label:'シート名順に並べ替え',view:'pages',handler:sortPagesBySheet});else actions.push({label:'最終PDFへ',view:'final'});showMessage(msg.kind,msg.title,msg.message+suffix,msg.detail,actions,0);setTimeout(hideProgressPanel,1800);}
    if(hasErrors)updateProgressPanel(finished);
  },false);}finally{renderJobActive=false;updateRenderTargetUi();}
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
  if (info?.mode === 'selected') return `${presetLabel()}で選択したExcel ${count}件をPDF作成します。`;
  if (info?.mode === 'needed') return `未選択のため、${presetLabel()}のPDF必要分 ${count}件をまとめて作成します。`;
  return 'PDF作成が必要なExcelはありません。';
}
function updateRenderTargetUi() {
  const hint = $('render-target-hint');
  const btn = $('render-selected-btn');
  if (!hint && !btn) return;
  const info = getWorkbookRenderTargetInfo();
  const activeCount = workbooksForActivePreset().filter(w => String(w.status || '') !== 'missing').length;
  if (btn) {
    if (info.mode === 'selected') btn.textContent = '選択分をPDF作成';
    else if (info.mode === 'needed') btn.textContent = 'PDF必要分を作成';
    else btn.textContent = 'PDF作成';
  }
  if (hint) {
    hint.classList.remove('target-selected', 'target-needed', 'target-none');
    hint.classList.add(info.mode === 'selected' ? 'target-selected' : (info.mode === 'needed' ? 'target-needed' : 'target-none'));
    if (info.mode === 'selected') hint.textContent = `${presetLabel()}で選択中: ${info.count}件をPDF作成します。`;
    else if (info.mode === 'needed') hint.textContent = `未選択: ${presetLabel()}のPDF必要分 ${info.count}件を作成します。`;
    else hint.textContent = activeCount ? 'PDF作成が必要なExcelはありません。' : `${presetLabel()}の登録済みExcelはありません。`;
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
  showMessage('ok', 'PDF作成が必要なExcelはありません', '登録済みExcelは最新です。');
}
async function renderUpdated(btn) {
  const neededIds = defaultPdfTargetWorkbookIds();
  if (!neededIds.length) {
    hideProgressPanel();
    showMessage('ok', 'PDF作成が必要なExcelはありません', '登録済みExcelは最新です。');
    return;
  }
  const info = {mode:'needed', count:neededIds.length};
  await renderWorkbookIds(neededIds, btn, {targetMode:'needed', startMessage:renderTargetStartMessage(info)});
}
async function renderAllWorkbooks(btn) {
  const ids = workbooksForActivePreset().map(w => w.workbookId);
  await renderWorkbookIds(ids, btn, {targetMode:'all', startMessage:`${presetLabel()}の登録済みExcel ${ids.length}件をPDF作成します。`});
}
async function renderSelectedPages(btn) {
  const pageIds = [...selectedPages];
  if (!pageIds.length) { showMessage('warn', 'シートを選択してください', 'PDF作成するシートにチェックを入れてください。'); return; }
  const ids = [...new Set(pageIds.map(pid => getPage(pid)?.workbookId).filter(Boolean))];
  await renderWorkbookIds(ids, btn);
}

function pageSort(a,b) {
  const ao = Number(a.order);
  const bo = Number(b.order);
  if (Number.isFinite(ao) && Number.isFinite(bo) && ao !== bo) return ao - bo;
  const as = sheetNumberValue(a.sheetName);
  const bs = sheetNumberValue(b.sheetName);
  if (as !== bs) return as - bs;
  const aw = getWorkbook(a.workbookId);
  const bw = getWorkbook(b.workbookId);
  const af = fileOrderValue(aw?.fileName || aw?.displayName || a.workbookId);
  const bf = fileOrderValue(bw?.fileName || bw?.displayName || b.workbookId);
  if (af !== bf) return af - bf;
  return resolvedPageId(a).localeCompare(resolvedPageId(b));
}
function numberingSelectValue(page) { return page.numberingManual ? (page.numberingMode || 'visible') : 'auto'; }
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

function applyChangedOnlyFilter(pages){
  if (!showChangedOnly || !state?.inputHistoryEnabled) return pages;
  const filtered = pages.filter(pageIsChanged);
  return filtered.length ? filtered : pages;
}
function renderPages() {
  const box=$('page-board');if(!box)return;const allPages=applyChangedOnlyFilter([...pagesForActivePreset()].sort(pageSort));
  const needsPdf=workbooksForActivePreset().filter(w=>!isLatestPdfWorkbook(w));const agg=aggregateFinalState();
  const signature=JSON.stringify([allPages.map(p=>[resolvedPageId(p),p.order,p.volume,p.enabled,p.title,p.numberingMode,p.numberingManual,p.orderManual,p.status,p.contentPdf]),needsPdf.map(w=>w.workbookId),agg.state]);
  if(lastPageBoardRenderSignature===signature&&box.childElementCount)return;lastPageBoardRenderSignature=signature;
  const scrollSnap=capturePageBoardScroll();selectedPages=new Set([...selectedPages].filter(id=>allPages.some(p=>resolvedPageId(p)===id)));
  const banners=[];if(needsPdf.length)banners.push(`<div class="page-status-banner"><strong>PDF作成が必要なExcelがあります</strong><span>${needsPdf.length}件。先にPDF作成してください。</span></div>`);if(agg.state==='needs-rebuild')banners.push('<div class="page-status-banner"><strong>ページ構成またはPDF入力が変更されています</strong><span>最終PDFの再出力が必要です。</span></div>');
  box.className='board';box.innerHTML=banners.join('')+[mainVolume(),appendixVolume(),'none'].map(volume=>{const pages=allPages.filter(p=>volume==='none'?(String(p.volume)==='none'||p.enabled===false):(String(p.volume||mainVolume())===volume&&p.enabled!==false));return `<section class="volume-panel" data-volume-panel="${volume}"><div class="volume-head"><h3>${escapeHtml(volumeLabel(volume))}</h3><span>${pages.length}ページ</span></div><div class="table-wrap"><table class="page-table"><thead><tr><th class="check-col"><input type="checkbox" data-select-all-pages="${volume}" aria-label="${escapeHtml(volumeLabel(volume))}をすべて選択"></th><th class="seq-col">順</th><th class="drag-col">移動</th><th>ページ名</th><th>PDF</th><th>番号</th></tr></thead><tbody data-volume="${volume}">${pages.map((p,i)=>pageRowHtml(p,i)).join('')||'<tr class="empty-row"><td colspan="6">ここに置く</td></tr>'}</tbody></table></div></section>`;}).join('');attachBoardEvents();restorePageBoardScroll(scrollSnap);
}

function pageRowHtml(p, idx) {
  const wb=getWorkbook(p.workbookId),pid=resolvedPageId(p),hasPdf=pagePreviewAvailable(p);const warnings=asArray(p.warnings).filter(w=>!/縮尺例外|Zoom|倍率例外/.test(String(w))).map(userFriendlyError);const checked=selectedPages.has(pid),highlighted=highlightedPageIds.has(pid);
  return `<tr tabindex="0" class="page-row ${checked?'selected-row':''} ${highlighted?'new-page-row':''} ${p.status==='render-error'?'error-row':''}" data-page-id="${escapeAttr(pid)}" data-workbook-id="${escapeAttr(p.workbookId||'')}" data-sheet-name="${escapeAttr(p.sheetName||'')}" data-content-pdf="${escapeAttr(p.contentPdf||'')}"><td class="check-col"><input type="checkbox" data-page-check value="${escapeAttr(pid)}" ${checked?'checked':''}></td><td class="seq-col"><span class="seq-badge" data-seq-cell>${idx+1}</span></td><td class="drag-col"><span class="drag-handle" title="ドラッグして移動">${iconUse('i-drag')}</span></td><td class="page-main-cell"><input class="title-input" data-page-title value="${escapeAttr(p.title||'')}"><button class="file-link btn ghost" type="button" ${hasPdf?`data-preview-page="${escapeAttr(pid)}"`:''}>${escapeHtml(wb?.displayName||wb?.fileName||'')} / シート ${escapeHtml(p.sheetName||'')}</button>${warnings.length?`<div class="page-warning">${warnings.map(escapeHtml).join('<br>')}</div>`:''}</td><td class="${hasPdf?'preview-trigger':''}" ${hasPdf?`data-preview-page="${escapeAttr(pid)}"`:''}>${pdfStatusForPage(p)} ${pageChangeBadge(p)}<div class="subtext">${escapeHtml(pagePdfSubtext(p))}</div></td><td><select data-page-numbering><option value="auto" ${numberingSelectValue(p)==='auto'?'selected':''}>自動</option><option value="none" ${numberingSelectValue(p)==='none'?'selected':''}>表示なし</option><option value="visible" ${numberingSelectValue(p)==='visible'?'selected':''}>表示</option></select><div class="subtext">${escapeHtml(numberingText(p,idx))}</div></td></tr>`;
}

function clearDropHighlights() {
  document.querySelectorAll('.drop-active').forEach(el => el.classList.remove('drop-active'));
  document.querySelectorAll('.drop-placeholder').forEach(el => el.remove());
}
function renumberBoardRows() {
  document.querySelectorAll('tbody[data-volume]').forEach(tbody => {
    let n = 1;
    [...tbody.querySelectorAll('tr')].forEach(row => {
      if (row.classList.contains('empty-row')) return;
      const seq = row.querySelector('[data-seq-cell], [data-placeholder-seq]');
      if (seq) seq.textContent = String(n++);
    });
  });
}
function getRowsForDrag(sourceRow) {
  const sourceId = sourceRow?.getAttribute('data-page-id') || '';
  if (sourceId && selectedPages.has(sourceId) && selectedPages.size > 1) {
    const selectedRows = [...document.querySelectorAll('.page-row')].filter(r => selectedPages.has(r.getAttribute('data-page-id')));
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
  document.querySelectorAll('tbody[data-volume]').forEach(tb => {
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
    : (rows[0]?.querySelector('.title-input')?.value || rows[0]?.querySelector('.page-main-cell .subtext')?.textContent || 'このページ');
  const ph = document.createElement('tr');
  ph.className = 'drop-placeholder';
  ph.innerHTML = `<td colspan="6"><div class="placeholder-card">ここへ移動：${escapeHtml(title)}</div></td>`;
  return ph;
}
function removeEmptyRow(tbody) {
  const empty = tbody?.querySelector?.('.empty-row');
  if (empty) empty.remove();
}
function findDropTbodyAt(x, y) {
  const elements = document.elementsFromPoint(x, y);
  for (const el of elements) {
    const tbody = el.closest?.('tbody[data-volume]');
    if (tbody) return tbody;
    const panel = el.closest?.('.volume-panel');
    if (panel) {
      const candidate = panel.querySelector('tbody[data-volume]');
      if (candidate) return candidate;
    }
  }
  // Fallback for cases where the pointer is over a sticky header or a scrollbar.
  let best = null;
  let bestDistance = Infinity;
  document.querySelectorAll('tbody[data-volume]').forEach(tbody => {
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
  const wrap = tbody?.closest?.('.table-wrap');
  if (wrap) {
    const rect = wrap.getBoundingClientRect();
    if (e.clientY > rect.bottom - 64) wrap.scrollTop += 24;
    else if (e.clientY < rect.top + 64) wrap.scrollTop -= 24;
  }
}
function getDragAfterRow(tbody, y) {
  const rows = [...tbody.querySelectorAll('.page-row:not(.dragging)')];
  return rows.reduce((closest, child) => {
    const box = child.getBoundingClientRect();
    const offset = y - box.top - box.height / 2;
    if (offset < 0 && offset > closest.offset) return {offset, element: child};
    return closest;
  }, {offset: Number.NEGATIVE_INFINITY, element: null}).element;
}
function moveDropPlaceholder(tbody, y) {
  if (!draggingRow || !draggingRow.placeholder || !tbody) return;
  const ph = draggingRow.placeholder;
  const after = getDragAfterRow(tbody, y);
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
  draggingRow = null;
  document.body.classList.remove('is-dragging-page');
  clearDropHighlights();
  renumberBoardRows();
  const afterSignature = currentBoardSignature();
  if (commit && moved && afterSignature !== beforeSignature) scheduleBoardSave();
}
function beginPointerPageDrag(e, row) {
  if (!row || e.button !== 0) return;
  if (!e.target.closest('.drag-handle')) return;
  if (e.target.closest('input, select, option, button, a, .preview-trigger')) return;
  e.preventDefault();
  const captureTarget = e.currentTarget || row;
  const startX = e.clientX;
  const startY = e.clientY;
  let started = false;
  const startDrag = (ev) => {
    if (started) return;
    started = true;
    const rows = getRowsForDrag(row);
    draggingRow = {sourceRow: row, rows, placeholder: createDropPlaceholder(rows), originalSignature: currentBoardSignature(), previewSignature: '', isNoop: true};
    for (const r of rows) r.classList.add('dragging');
    document.body.classList.add('is-dragging-page');
    const tbody = findDropTbodyAt(ev.clientX, ev.clientY) || row.closest('tbody[data-volume]');
    if (tbody) moveDropPlaceholder(tbody, ev.clientY);
  };
  try { captureTarget.setPointerCapture?.(e.pointerId); } catch {}
  const cleanup = (ev) => {
    try { captureTarget.releasePointerCapture?.(ev.pointerId); } catch {}
    document.removeEventListener('pointermove', onMove, true);
    document.removeEventListener('pointerup', onUp, true);
    document.removeEventListener('pointercancel', onCancel, true);
  };
  const onMove = (ev) => {
    const moved = Math.abs(ev.clientX - startX) + Math.abs(ev.clientY - startY);
    if (!started && moved < 5) return;
    if (!started) startDrag(ev);
    if (!draggingRow) return;
    ev.preventDefault();
    const tbody = findDropTbodyAt(ev.clientX, ev.clientY);
    autoScrollDuringDragPointer(ev, tbody);
    if (tbody) moveDropPlaceholder(tbody, ev.clientY);
  };
  const onUp = (ev) => {
    cleanup(ev);
    if (!started) { return; }
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

function attachBoardEvents() {
  document.querySelectorAll('.page-row').forEach(row=>{
    row.addEventListener('pointerdown',e=>beginPointerPageDrag(e,row));
    row.addEventListener('keydown',e=>{
      if(!e.altKey||!['ArrowUp','ArrowDown'].includes(e.key)||e.target.matches('input,select,textarea'))return;e.preventDefault();const tbody=row.parentElement;const rows=[...tbody.querySelectorAll('.page-row')];const idx=rows.indexOf(row);if(idx<0)return;
      if(e.shiftKey){if(e.key==='ArrowUp')tbody.insertBefore(row,rows[0]);else tbody.appendChild(row);}else if(e.key==='ArrowUp'&&idx>0)tbody.insertBefore(row,rows[idx-1]);else if(e.key==='ArrowDown'&&idx<rows.length-1)tbody.insertBefore(rows[idx+1],row);
      renumberBoardRows();scheduleBoardSave();row.focus();
    });
  });
  document.querySelectorAll('[data-page-check]').forEach(ch=>ch.addEventListener('click',e=>{e.stopPropagation();handlePageCheckboxToggle(ch,e.shiftKey);}));
  document.querySelectorAll('[data-select-all-pages]').forEach(ch=>{const volume=String(ch.dataset.selectAllPages||''),tbody=[...document.querySelectorAll('tbody[data-volume]')].find(t=>t.dataset.volume===volume),ids=tbody?[...tbody.querySelectorAll('.page-row')].map(r=>r.dataset.pageId).filter(Boolean):[];const all=ids.length&&ids.every(id=>selectedPages.has(id)),some=ids.some(id=>selectedPages.has(id));ch.checked=all;ch.indeterminate=some&&!all;ch.addEventListener('change',()=>{if(ch.checked)ids.forEach(id=>selectedPages.add(id));else ids.forEach(id=>selectedPages.delete(id));lastPageRangeAnchor='';syncPageSelectionUi();});});
  document.querySelectorAll('[data-page-title]').forEach(input=>input.addEventListener('blur',()=>savePageFromRow(input.closest('.page-row'),false)));
  document.querySelectorAll('[data-page-numbering]').forEach(sel=>sel.addEventListener('change',()=>savePageFromRow(sel.closest('.page-row'),true)));
  document.querySelectorAll('[data-preview-page]').forEach(el=>el.addEventListener('click',e=>{if(draggingRow)return;if(e.target&&e.target.matches('input,select,option'))return;const row=el.closest('.page-row');previewPage(el.dataset.previewPage||row?.dataset.pageId,pageFallbackFromRow(row));}));
  document.querySelectorAll('.page-row').forEach(row=>row.addEventListener('dblclick',e=>{if(draggingRow||e.target.closest('input,select'))return;if(!row.querySelector('[data-preview-page]'))return;previewPage(row.dataset.pageId,pageFallbackFromRow(row));}));
}

function collectBoardVolumes() {
  const volumes = {};
  document.querySelectorAll('tbody[data-volume]').forEach(tbody => {
    volumes[tbody.getAttribute('data-volume')] = [...tbody.querySelectorAll('.page-row')].map(row => row.getAttribute('data-page-id'));
  });
  return volumes;
}
function scheduleBoardSave() { clearTimeout(boardSaveTimer); boardSaveTimer = setTimeout(saveBoardOrder, 250); }
async function saveBoardOrder() {
  try{await api('/api/pages/reorder',{method:'POST',body:{category:activePreset,volumes:collectBoardVolumes()}});lastPageBoardRenderSignature='';await refresh();showMessage('ok','ページ構成を保存しました','本体・補足の並びを反映しました。',null,[{label:'最終PDFへ',view:'final'}]);}
  catch(e){showMessage('danger','並び替えを保存できません',userFriendlyError(e.message),e.detail||e.stack||e.message);lastPageBoardRenderSignature='';await refresh();}
}

async function savePageFromRow(row, numberingChanged=false) {
  if(!row)return;const pageId=row.getAttribute('data-page-id'),page=getPage(pageId);const body={category:activePreset,pageId,title:row.querySelector('[data-page-title]')?.value||page?.title||''};const numbering=row.querySelector('[data-page-numbering]')?.value||'auto';if(numberingChanged){if(numbering==='auto')body.resetNumbering=true;else{body.numberingMode=numbering;body.numberingManual=true;}}
  try{await api('/api/pages/update',{method:'POST',body});lastPageBoardRenderSignature='';await refresh();}catch(e){showMessage('danger','ページを保存できません',userFriendlyError(e.message),e.detail||e.stack||e.message);}
}


async function sortPagesBySheet(btn=null) {
  if(!confirm('本体・補足・出力しないの割り当てはそのままで、それぞれの表の中だけをシート名の数字順に並べ替えます。よろしいですか？'))return;
  await runBusy(btn||$('sort-by-sheet-btn'),async()=>{await api('/api/pages/sort-by-sheet',{method:'POST',body:{category:activePreset,volumes:[mainVolume(),appendixVolume(),'none']}});lastPageBoardRenderSignature='';await refresh();showMessage('ok','シート名順に並べ替えました','本体・補足・出力しないの各表を整列しました。',null,[{label:'最終PDFへ',view:'final'}]);});
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
async function previewPage(pageId, fallback={}) {
  const page=getPage(pageId)||getPageByWorkbookSheet(fallback.workbookId,fallback.sheetName);const pid=resolvedPageId(page)||String(pageId||'').trim();const contentPdf=String(page?.contentPdf||fallback.contentPdf||'').trim();if(!contentPdf){showMessage('warn','PDF未作成です','登録済みExcelを選択して「PDF作成」を押してください。');return;}
  modalReturnFocus=document.activeElement;const wb=getWorkbook(page?.workbookId||fallback.workbookId);$('preview-title').textContent=page?.title||fallback.title||'PDF確認';$('preview-subtitle').textContent=`${wb?.displayName||wb?.fileName||''}${(page?.sheetName||fallback.sheetName)?' / シート '+(page?.sheetName||fallback.sheetName):''}`;$('preview-open-new').removeAttribute('href');$('pdf-frame').src='about:blank';$('preview-modal').classList.remove('hidden');$('preview-close').focus();
  try{clearPreviewObjectUrl();currentPreviewObjectUrl=await fetchPdfObjectUrl('/api/file',{pageId:pid,id:pid,workbookId:page?.workbookId||fallback.workbookId||'',sheetName:page?.sheetName||fallback.sheetName||'',contentPdf});$('preview-open-new').href=currentPreviewObjectUrl;$('pdf-frame').src=`${currentPreviewObjectUrl}#toolbar=1&navpanes=0`;}catch(e){closePreview();showMessage('danger','PDFプレビューを開けません',userFriendlyError(e.message),e.detail||e.stack||e.message);}
}


function closePreview() {
  const modal=$('preview-modal');if(!modal||modal.classList.contains('hidden'))return;modal.classList.add('hidden');$('pdf-frame').src='about:blank';clearPreviewObjectUrl();if(modalReturnFocus&&typeof modalReturnFocus.focus==='function')modalReturnFocus.focus();modalReturnFocus=null;
}


async function openFinalVolume(volume, category=activePreset) {
  let blank=null;try{blank=window.open('about:blank','_blank');}catch{}if(blank){try{blank.opener=null;}catch{}}
  try{const objectUrl=await fetchPdfObjectUrl('/api/final/file',{volume,category});if(blank&&!blank.closed)blank.location.href=objectUrl;else window.open(objectUrl,'_blank');}
  catch(e){if(blank){try{blank.close();}catch{}}showMessage('danger','PDFを開けません',userFriendlyError(e.message),e.detail||e.stack||e.message);}
}


function pathElementValue(id) {
  const el=$(id);if(!el)return '';return String(('value' in el?el.value:el.textContent)||'').trim().replace(/^—$/,'');
}
function setPathElementValue(id,value) {
  const el=$(id);if(!el)return;if('value' in el)el.value=value||'';else el.textContent=value||'—';
}
function applyDefaultChildPaths(force=false) {
  const sub = pathElementValue('submissionDir');
  if (!sub) return;
  const nextData = pathJoin(sub, '_reportbinder');
  const nextOut = pathJoin(sub, '出力');
  const currentData=pathElementValue('dataDir'),currentOut=pathElementValue('outputDir');
  if (force || !currentData || currentData === pathJoin(lastSubmissionDefaultBase, '_reportbinder')) setPathElementValue('dataDir',nextData);
  if (force || !currentOut || currentOut === pathJoin(lastSubmissionDefaultBase, '出力')) setPathElementValue('outputDir',nextOut);
  lastSubmissionDefaultBase = sub;
}
async function chooseSubmissionFolder(btn) {
  if (folderPickerBusy) return;
  folderPickerBusy = true;
  const old = btn?.textContent;
  if (btn) { btn.disabled = true; btn.textContent = '選択画面を開く'; }
  showMessage('', '提出フォルダを選んでください', 'フォルダ選択画面を開いています。');
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
    showMessage('ok', '提出フォルダを設定しました', 'Excelを選んで登録してください。');
  } catch(e) {
    showMessage('danger', '提出フォルダを選べません', userFriendlyError(e.message), e.detail || e.stack || e.message);
  } finally {
    folderPickerBusy = false;
    if (btn) { btn.disabled = false; btn.textContent = old; }
  }
}

// Kept for compatibility with older local pages; the current UX uses chooseSubmissionFolder().
async function savePaths(btn) {
  const body = { submissionDir: pathElementValue('submissionDir'), dataDir: pathElementValue('dataDir'), outputDir: pathElementValue('outputDir') };
  await runBusy(btn, async () => {
    await api('/api/paths', {method:'POST', body});
    await refresh();
    await loadFiles(null);
    showMessage('ok', '提出フォルダを設定しました', 'Excelを選んで登録してください。');
  });
}
function showPostBuildWarning(volume) {
  const after=volumeReadiness(volume),blockers=asArray(after.blockers);
  if(blockers.length){setTimeout(()=>showMessage('warn','最終PDFは出力しましたが、追加対応が必要です',blockers[0]?.message||'ExcelをPDF作成し直してから、最終PDFを再出力してください。',blockers,[{label:'Excel登録・PDF作成へ',view:'excel'}],0),900);}
  else if(String(after.displayState)==='needs-rebuild'){setTimeout(()=>showMessage('warn','最終PDFの再出力が必要です','出力中に入力が変更されました。最新状態で再度出力してください。',after.staleReasons,[{label:'最終PDFを確認',view:'final'}],0),900);}
}
async function buildVolume(volume, btn) {
  const ready=volumeReadiness(volume);if(asArray(ready.blockers).length){setActiveView('excel');showMessage('warn','先にExcelのPDFを作成してください',ready.blockers[0]?.message||'出力条件を確認してください。');return;}
  await runBusy(btn,async()=>{const result=await api('/api/final/build',{method:'POST',body:{volume,category:activePreset}});await refresh();showMessage('ok',`${volumeLabel(volume)}PDFを出力しました`,result.result?.outputPdf||'出力フォルダを確認してください。',result.result,[{label:'PDFを開く',primary:true,handler:()=>openFinalVolume(volume,activePreset)}],0);showPostBuildWarning(volume);});
}

async function buildAllVolumes(btn) {
  // V5: 本体だけ成功する状態を作らないよう、サーバー側の準トランザクションAPIを1回だけ呼ぶ。
  await runBusy(btn, async () => {
    const r = await api('/api/final/build-all', {method:'POST', body:{category: activePreset}});
    const built = asArray(r.result?.built);
    const skipped = asArray(r.result?.skipped);
    await refresh();
    if (built.length) {
      showMessage('ok','最終PDFを出力しました', `${built.length}件を出力しました。`, {built, skipped},
                  [{label:'出力したPDFを確認', view:'final'}], 0);
    } else {
      showMessage('warn','出力対象がありません','ページ構成で本体または補足にページを設定してください。');
    }
  });
}


async function runBusy(btn, fn, showProcessing=true) {
  if(btn){btn.disabled=true;btn.classList.add('busy');}
  if(showProcessing)showMessage('','処理中です','完了すると画面が更新されます。');
  try{await fn();}catch(e){hideProgressPanel();showMessage('danger','処理できませんでした',userFriendlyError(e.message),e.detail||e.stack||e.message);}finally{if(btn){btn.disabled=false;btn.classList.remove('busy');}}
}


async function sendHeartbeat() { try { await api('/api/heartbeat', {method:'POST', body:{clientId}, keepalive:true}); } catch { } }
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

document.querySelectorAll('[data-view-nav]').forEach(btn=>btn.addEventListener('click',()=>setActiveView(btn.dataset.viewNav)));
document.querySelectorAll('[data-view-shortcut]').forEach(btn=>btn.addEventListener('click',()=>setActiveView(btn.dataset.viewShortcut)));
document.querySelectorAll('[data-step-view]').forEach(btn=>btn.addEventListener('click',()=>setActiveView(btn.dataset.stepView)));
setActiveView(activeView,{noScroll:true,instant:true});
const fileFilterInput=$('file-filter');if(fileFilterInput)fileFilterInput.addEventListener('input',()=>{fileFilterText=fileFilterInput.value||'';renderFileList(availableFiles);});
bind('dashboard-action-btn','click',async()=>{const action=$('dashboard-action-btn')?.dataset.dashboardAction||'excel';if(action==='folder'){setActiveView('folders');await chooseSubmissionFolder($('choose-submission-btn'));}else if(action==='render'){setActiveView('excel');selectRenderNeededWorkbooks();await renderSelectedWorkbooks($('render-selected-btn'));}else setActiveView(action==='pages'?'pages':action==='final'?'final':'excel');});
bind('notice-close','click',hideMessage);bind('scan-btn','click',()=>scanAndRefresh($('scan-btn')));bind('scan-folder-btn','click',()=>scanAndRefresh($('scan-folder-btn')));bind('choose-submission-btn','click',()=>chooseSubmissionFolder($('choose-submission-btn')));
document.querySelectorAll('[data-preset]').forEach(btn=>btn.addEventListener('click',()=>applyPresetSelection(btn.dataset.preset)));
bind('select-visible-files-btn','click',selectVisibleFiles);bind('clear-selected-files-btn','click',clearVisibleFilesSelection);bind('select-render-needed-btn','click',selectRenderNeededWorkbooks);bind('clear-selected-workbooks-btn','click',clearWorkbookSelection);bind('select-all-pages-btn','click',selectAllActivePages);bind('clear-selected-pages-btn','click',clearActivePageSelection);
bind('bulk-main-btn','click',()=>moveSelectedPagesToVolume(mainVolume(),$('bulk-main-btn')));bind('bulk-appendix-btn','click',()=>moveSelectedPagesToVolume(appendixVolume(),$('bulk-appendix-btn')));bind('bulk-none-btn','click',()=>moveSelectedPagesToVolume('none',$('bulk-none-btn')));
bind('register-selected-btn','click',()=>registerSelected($('register-selected-btn')));bind('unregister-selected-btn','click',()=>unregisterSelected($('unregister-selected-btn')));bind('render-selected-btn','click',()=>renderSelectedWorkbooks($('render-selected-btn')));bind('render-all-btn','click',()=>renderAllWorkbooks($('render-all-btn')));bind('sort-by-sheet-btn','click',()=>sortPagesBySheet($('sort-by-sheet-btn')));
bind('build-main-btn','click',()=>buildVolume(mainVolume(),$('build-main-btn')));bind('build-appendix-btn','click',()=>buildVolume(appendixVolume(),$('build-appendix-btn')));bind('build-all-btn','click',()=>buildAllVolumes($('build-all-btn')));
const changedOnlyToggle=$('changed-only-toggle');
if(changedOnlyToggle)changedOnlyToggle.addEventListener('change',()=>{showChangedOnly=!!changedOnlyToggle.checked;renderPages();});
bind('history-refresh-btn','click',()=>loadHistoryPanels());
// V5-P1(#12): 版履歴パネルの対象Excel切り替え。
bind('snapshot-history-workbook','change',e=>{snapshotHistoryState.workbookId=String(e.target.value||'');snapshotHistoryState.fromId='';snapshotHistoryState.toId='';const db=$('snapshot-history-diff');if(db)db.innerHTML='';loadSnapshotHistory();});
bind('final-main-fix','click',()=>setActiveView('excel'));bind('final-appendix-fix','click',()=>setActiveView('excel'));
bind('open-main-link','click',e=>{e.preventDefault();openFinalVolume(mainVolume(),activePreset);});bind('open-appendix-link','click',e=>{e.preventDefault();openFinalVolume(appendixVolume(),activePreset);});
bind('mode-badge','click',()=>{const pop=$('language-popover'),btn=$('mode-badge');const hidden=pop.classList.toggle('hidden');btn.setAttribute('aria-expanded',String(!hidden));});
bind('preview-close','click',closePreview);const previewModal=$('preview-modal');if(previewModal)previewModal.addEventListener('click',e=>{if(e.target===previewModal)closePreview();});
bind('diff-close','click',closeDiffDetail);
const diffModal=$('diff-modal');if(diffModal)diffModal.addEventListener('click',e=>{if(e.target===diffModal)closeDiffDetail();});
bind('diff-mode-side','click',()=>{diffViewState.mode='side';applyDiffMode();});
bind('diff-mode-overlay','click',()=>{diffViewState.mode='overlay';applyDiffMode();});
bind('diff-highlight','change',applyDiffHighlightSettings);
bind('diff-density','change',applyDiffHighlightSettings);
bind('diff-zoom','change',updateDiffStageScale);
bind('diff-opacity','input',applyDiffMode);
bind('diff-prev-page','click',()=>setDiffPage(diffViewState.pageIndex-1));
bind('diff-next-page','click',()=>setDiffPage(diffViewState.pageIndex+1));
bind('diff-prev-region','click',()=>moveDiffRegion(-1));
bind('diff-next-region','click',()=>moveDiffRegion(1));
document.querySelectorAll('[data-diff-tab]').forEach(btn=>btn.addEventListener('click',()=>{
  diffViewState.mobileTab=btn.dataset.diffTab||'before';
  document.querySelectorAll('[data-diff-tab]').forEach(x=>x.classList.toggle('active',x===btn));
  $('diff-viewers')?.classList.toggle('show-after',diffViewState.mobileTab==='after');
  requestAnimationFrame(updateDiffStageScale);
}));
const beforeDiffViewport=$('diff-before-viewport'),afterDiffViewport=$('diff-after-viewport');
beforeDiffViewport?.addEventListener('scroll',()=>syncDiffScroll(beforeDiffViewport,afterDiffViewport),{passive:true});
afterDiffViewport?.addEventListener('scroll',()=>syncDiffScroll(afterDiffViewport,beforeDiffViewport),{passive:true});
window.addEventListener('resize',()=>{if(isDiffModalOpen())updateDiffStageScale();});
window.addEventListener('keydown',e=>{
  if(e.key==='Escape'){
    if(isDiffModalOpen())closeDiffDetail();else closePreview();
    return;
  }
  if(e.key!=='Tab')return;
  const modal=isDiffModalOpen()?$('diff-modal'):(isModalOpen()?$('preview-modal'):null);
  if(!modal)return;
  const focusable=[...modal.querySelectorAll('a[href],button:not([disabled]),input:not([disabled]),select:not([disabled]),[tabindex]:not([tabindex="-1"]),iframe')].filter(el=>el.offsetParent!==null);
  if(!focusable.length)return;
  const first=focusable[0],last=focusable[focusable.length-1];
  if(e.shiftKey&&document.activeElement===first){e.preventDefault();last.focus();}
  else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first.focus();}
});

(async function init() {
  startLifecycle();
  startUpdateMonitor();
  await refresh();
  if (configured()) {
    try {
      await loadFilesSilently();
      renderFileList(availableFiles);
    } catch (e) {
      log('Excel一覧の初期読み込みでエラー', e.detail || e.stack || e.message);
    }
    // Heavy update checks can make the first folder/dialog action feel unresponsive on a single local server thread.
    // Run the first scan after the user has had time to interact; normal monitoring continues afterwards.
    setTimeout(() => scanUpdatesSilently({withFiles:true}), 8000);
  }
})();

window.addEventListener('pagehide', () => { for (const u of [...pdfObjectUrls]) revokePdfObjectUrl(u); });


// ---- V5: 自動処理の状態表示と履歴タイムライン ----
function renderAutoStatus(){
  const box=$('auto-status');
  if(!box)return;
  const auto=state?.auto;
  if(!auto||!auto.inputHistoryApproved){box.classList.add('hidden');box.innerHTML='';return;}
  box.classList.remove('hidden');
  const items=asArray(auto.items);
  const cap=Number(auto.softCapMegabytes||0),size=Number(auto.historySizeMb||0);
  const overCap=cap>0&&size>=cap*0.8;
  const lines=items.slice(0,5).map(it=>{
    const wb=getWorkbook(it.workbookId);
    const name=wb?workbookDisplayName(wb):String(it.workbookId||'');
    const reason={'excel-in-use':'Excelが使用されています','job-busy':'他のPDF作成が実行中です','start-failed':'開始できませんでした','file-missing':'ファイルが見つかりません'}[String(it.deferReason||'')]||'';
    const label={waiting:'変更を検知（静止待ち）',deferred:'保留中',rendering:'PDF作成中',ready:'実行待ち'}[String(it.state||'')]||String(it.state||'');
    return `<div class="auto-row"><span>${escapeHtml(name)}</span><span>${escapeHtml(label)}</span><span class="subtext">${escapeHtml(reason)}</span>${String(it.state||'')==='deferred'?`<button class="btn ghost compact" type="button" data-auto-run="${escapeAttr(it.workbookId)}">今すぐ実行</button>`:''}</div>`;
  }).join('');
  box.innerHTML=`<div class="card-head"><h3>自動PDF作成</h3><span class="caption">${auto.enabled?'有効':'無効'}${auto.enabled?`（変更後 ${Math.round(Number(auto.quietPeriodSeconds||180)/60)}分待機）`:''}</span></div>`
    +(items.length?`<div class="auto-list">${lines}</div>`:'<div class="caption">保留中の変更はありません。</div>')
    +(overCap?`<p class="warn-strip">変更履歴の使用量が ${size}MB です（上限の目安 ${cap}MB）。</p>`:'');
  box.querySelectorAll('[data-auto-run]').forEach(b=>b.addEventListener('click',async()=>{
    try{await api('/api/auto/run-now',{method:'POST',body:{workbookId:b.getAttribute('data-auto-run')}});await refresh();}
    catch(e){showMessage('danger','実行を要求できません',userFriendlyError(e.message));}
  }));
}

async function loadHistoryTimeline(){
  const box=$('history-timeline');
  if(!box)return;
  if(!state?.inputHistoryEnabled){box.innerHTML='<div class="caption">変更履歴は無効です。</div>';return;}
  try{
    const r=await api('/api/history/timeline?limit=50');
    const events=asArray(r.events);
    if(!events.length){box.innerHTML='<div class="caption">履歴はまだありません。</div>';return;}
    const label={'input.snapshot.created':'提出Excelの変更を検知','input.snapshot.deduplicated':'同じ内容を再検知','input.source.removed':'保存期限により現物を削除','render.started':'PDF作成を開始','render.completed':'PDF作成が完了','render.failed':'PDF作成に失敗','compare.completed':'変更を判定','layout.changed':'ページ構成を変更','layout.restored':'ページ構成を復元','final.built':'最終PDFを出力','final.build.failed':'最終PDFの出力に失敗','final.archive.created':'最終PDFを保存','auto.detected':'自動処理が変更を検知','auto.deferred':'自動処理を保留','history.cleanup':'履歴を整理'};
    box.innerHTML=events.map(e=>{
      const t=String(e.eventType||'');
      return `<div class="history-row"><span class="date-col">${escapeHtml(formatDateTime(e.at))}</span><span>${escapeHtml(label[t]||t)}</span><span class="subtext">${escapeHtml(String(e.data?.workbookId||e.data?.category||''))}</span></div>`;
    }).join('');
  }catch(e){box.innerHTML=`<div class="caption">履歴を読み込めません：${escapeHtml(userFriendlyError(e.message))}</div>`;}
}


// ---- V5-P1: レイアウト履歴・復元・アーカイブのUI配線 ----
async function loadLayoutSnapshots(){
  const box=$('layout-history');
  if(!box)return;
  if(!state?.inputHistoryEnabled){box.innerHTML='<div class="caption">変更履歴は無効です。</div>';return;}
  try{
    const r=await api(`/api/layout/snapshots?category=${encodeURIComponent(activePreset)}`);
    const list=asArray(r.snapshots);
    if(!list.length){box.innerHTML='<div class="caption">保存されたページ構成はまだありません。</div>';return;}
    const reasonLabel={reorder:'並べ替え','sort-by-sheet':'シート名順に並べ替え','volume-change':'出力先の変更',numbering:'ページ番号の変更',enabled:'出力対象の変更','final-build':'最終PDF出力時',restore:'復元',
      'pre-restore':'復元の直前'};
    box.innerHTML=list.map(x=>`<div class="history-row"><span class="date-col">${escapeHtml(formatDateTime(x.createdAt))}</span><span>${escapeHtml(reasonLabel[String(x.reason||'')]||String(x.reason||''))}</span><span class="subtext">${Number(x.pageCount||0)}ページ</span><button class="btn ghost compact" type="button" data-restore-layout="${escapeAttr(x.snapshotId)}">この状態に戻す</button></div>`).join('');
    box.querySelectorAll('[data-restore-layout]').forEach(b=>b.addEventListener('click',()=>previewLayoutRestore(b.getAttribute('data-restore-layout'))));
  }catch(e){box.innerHTML=`<div class="caption">履歴を読み込めません：${escapeHtml(userFriendlyError(e.message))}</div>`;}
}

async function previewLayoutRestore(snapshotId){
  try{
    const r=await api('/api/layout/restore/preview',{method:'POST',body:{category:activePreset,snapshotId}});
    const p=r.preview||{};
    const volumeLines=asArray(p.volumeChanges).slice(0,10).map(v=>`・${v.title||v.pageId}：${volumeLabel(v.from)} → ${volumeLabel(v.to)}`).join('\n');
    const text=[`${formatDateTime(p.createdAt)} の構成に戻します`,'',
      `適用されるページ：${Number(p.appliedPageCount||0)}`,
      `過去にだけ存在するページ：${asArray(p.pastOnlyPageIds).length}（適用しません）`,
      `現在にだけ存在するページ：${asArray(p.currentOnlyPageIds).length}（そのまま残ります）`,
      volumeLines?`\n出力先の変更：\n${volumeLines}`:'',
      '','復元後、最終PDFの再出力が必要になります。'].filter(Boolean).join('\n');
    if(!confirm(text))return;
    const res=await api('/api/layout/restore',{method:'POST',body:{category:activePreset,snapshotId}});
    await refresh();
    await loadLayoutSnapshots();
    showMessage('ok','ページ構成を復元しました',`${Number(res.result?.appliedPageCount||0)}ページに適用しました。取り消す場合は履歴の「復元の直前」から戻せます。`,null,
                [{label:'最終PDFを確認',view:'final'}],0);
  }catch(e){showMessage('danger','復元できません',userFriendlyError(e.message));}
}

function volumeLabel(v){
  return {'ja-main':'本体','ja-appendix':'補足','en-main':'Main','en-appendix':'Appendix','none':'出力しない'}[String(v||'')]||String(v||'');
}

async function loadFinalArchives(){
  const box=$('final-archives');
  if(!box)return;
  if(!state?.inputHistoryEnabled){box.innerHTML='<div class="caption">変更履歴は無効です。</div>';return;}
  try{
    const r=await api(`/api/final/archives?category=${encodeURIComponent(activePreset)}`);
    const list=asArray(r.archives);
    if(!list.length){box.innerHTML='<div class="caption">保存された出力はまだありません。</div>';return;}
    box.innerHTML=list.slice(0,20).map(a=>`<div class="history-row"><span class="date-col">${escapeHtml(formatDateTime(a.builtAt))}</span><span>${escapeHtml(volumeLabel(a.volume))}</span><span class="subtext">${escapeHtml(String(a.outputFileName||''))}（${Number(a.pageCount||0)}ページ）</span></div>`).join('');
  }catch(e){box.innerHTML=`<div class="caption">読み込めません：${escapeHtml(userFriendlyError(e.message))}</div>`;}
}

// ---- V5-P1(#12): 版の履歴・差分・保護（手動pin / 2版差分 / 過去content-pdf表示）----
let snapshotHistoryState = { workbookId:'', snapshots:[], fromId:'', toId:'' };

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
  sel.innerHTML='<option value="">選択してください</option>'+list.map(w=>`<option value="${escapeAttr(w.workbookId)}">${escapeHtml(workbookDisplayName(w))}</option>`).join('');
  if(prev && list.some(w=>String(w.workbookId)===String(prev))){ sel.value=prev; snapshotHistoryState.workbookId=String(prev); }
  else { snapshotHistoryState.workbookId=''; }
}

async function loadSnapshotHistory(){
  const box=$('snapshot-history'), diffBox=$('snapshot-history-diff');
  if(!box)return;
  if(!state?.inputHistoryEnabled){box.innerHTML='<div class="caption">変更履歴は無効です。</div>';if(diffBox)diffBox.innerHTML='';return;}
  populateSnapshotHistoryWorkbooks();
  const wbId=String(snapshotHistoryState.workbookId||'');
  if(!wbId){box.innerHTML='<div class="caption">対象Excelを選ぶと、版の一覧・差分・保護を表示します。</div>';if(diffBox)diffBox.innerHTML='';return;}
  try{
    const r=await api(`/api/history/snapshots?workbookId=${encodeURIComponent(wbId)}`);
    const list=asArray(r.snapshots);snapshotHistoryState.snapshots=list;
    if(!list.length){box.innerHTML='<div class="caption">この Excel の保存された版はまだありません。</div>';if(diffBox)diffBox.innerHTML='';return;}
    const ready=list.filter(s=>!!s.visualCompareReady);
    if(!snapshotHistoryState.toId||!ready.some(s=>String(s.snapshotId)===String(snapshotHistoryState.toId)))snapshotHistoryState.toId=String(ready[0]?.snapshotId||'');
    if(!snapshotHistoryState.fromId||!ready.some(s=>String(s.snapshotId)===String(snapshotHistoryState.fromId)))snapshotHistoryState.fromId=String(ready[1]?.snapshotId||'');
    box.innerHTML=`<table class="data-table"><thead><tr><th>検知日時</th><th>ハッシュ</th><th>現物</th><th>視覚比較</th><th>保護</th><th class="check-col">基準</th><th class="check-col">対象</th></tr></thead><tbody>${list.map(s=>{
      const id=String(s.snapshotId||'');const pinned=asArray(s.pins).map(String).includes('manual');
      const retained=s.sourceRetained?'あり':'期限切れ';const hash=String(s.sourceHash||'').slice(0,10);
      const ready=!!s.visualCompareReady;const reason=String(s.unavailableReason||'比較用PDFがありません。');
      const compareState=ready?badge('可能','neutral'):`<span class="subtext" title="${escapeAttr(reason)}">不可</span>`;
      const disabled=ready?'':`disabled title="${escapeAttr(reason)}"`;
      return `<tr><td class="date-col">${escapeHtml(formatDateTime(s.detectedAt))}</td><td class="subtext">${escapeHtml(hash)}</td><td class="subtext">${escapeHtml(retained)}</td><td>${compareState}</td><td>${pinned?badge('保護中','neutral'):''} <button class="btn ghost compact" type="button" data-snap-pin="${escapeAttr(id)}" data-pinned="${pinned?'1':'0'}">${pinned?'保護解除':'保護する'}</button></td><td class="check-col"><input type="radio" name="snap-from" value="${escapeAttr(id)}" ${String(snapshotHistoryState.fromId)===id?'checked':''} ${disabled}></td><td class="check-col"><input type="radio" name="snap-to" value="${escapeAttr(id)}" ${String(snapshotHistoryState.toId)===id?'checked':''} ${disabled}></td></tr>`;
    }).join('')}</tbody></table><div class="card-actions"><button class="btn ghost compact" id="snap-swap-btn" type="button" ${ready.length<2?'disabled':''}>比較元と比較先を入れ替え</button><button class="btn secondary compact" id="snap-diff-btn" type="button" ${ready.length<2?'disabled':''}>選んだ2版を視覚比較</button></div>`;
    box.querySelectorAll('[data-snap-pin]').forEach(b=>b.addEventListener('click',()=>toggleSnapshotPin(wbId,b.getAttribute('data-snap-pin'),b.getAttribute('data-pinned')==='1')));
    box.querySelectorAll('input[name="snap-from"]').forEach(el=>el.addEventListener('change',()=>{snapshotHistoryState.fromId=el.value;renderSnapshotHistorySelectionHint();}));
    box.querySelectorAll('input[name="snap-to"]').forEach(el=>el.addEventListener('change',()=>{snapshotHistoryState.toId=el.value;renderSnapshotHistorySelectionHint();}));
    const swap=$('snap-swap-btn');if(swap)swap.addEventListener('click',()=>{const previousFrom=snapshotHistoryState.fromId;snapshotHistoryState.fromId=snapshotHistoryState.toId;snapshotHistoryState.toId=previousFrom;loadSnapshotHistory();});
    const db=$('snap-diff-btn');if(db)db.addEventListener('click',()=>loadSnapshotDiff(wbId));
    renderSnapshotHistorySelectionHint();
  }catch(e){box.innerHTML=`<div class="caption">版の履歴を読み込めません：${escapeHtml(userFriendlyError(e.message))}</div>`;}
}

async function toggleSnapshotPin(workbookId, snapshotId, pinned){
  try{
    await api(pinned?'/api/history/unpin':'/api/history/pin',{method:'POST',body:{workbookId,snapshotId}});
    showMessage('ok',pinned?'保護を解除しました':'この版を保護しました',pinned?'':'保存期限による自動削除から守ります。');
    await loadSnapshotHistory();
  }catch(e){showMessage('danger','操作できません',userFriendlyError(e.message));}
}

async function loadSnapshotDiff(workbookId){
  const diffBox=$('snapshot-history-diff');if(!diffBox)return;
  const from=String(snapshotHistoryState.fromId||''), to=String(snapshotHistoryState.toId||'');
  const fromItem=snapshotHistoryState.snapshots.find(s=>String(s.snapshotId)===from);
  const toItem=snapshotHistoryState.snapshots.find(s=>String(s.snapshotId)===to);
  if(!from||!to){diffBox.innerHTML='<div class="caption">比較可能な版から比較元と比較先を選んでください。</div>';return;}
  if(from===to){diffBox.innerHTML='<div class="caption">異なる2版を選んでください。</div>';return;}
  if(!fromItem?.visualCompareReady||!toItem?.visualCompareReady){diffBox.innerHTML='<div class="caption">画像ハッシュと同一世代のcontent PDFがそろった版だけ比較できます。</div>';return;}
  diffBox.innerHTML='<div class="caption">選択した履歴版の差分を準備します。自動比較の基準は変更されません。</div>';
  await openDiffDetail(workbookId,$('snap-diff-btn'),{fromSnapshotId:from,toSnapshotId:to});
}

async function viewHistoryContentPdf(workbookId, snapshotId, versionId, sheetName){
  if(!versionId){showMessage('warn','この版のPDFは保存されていません','この版ではページPDFが作成されていないため表示できません。');return;}
  modalReturnFocus=document.activeElement;
  const wb=getWorkbook(workbookId);
  $('preview-title').textContent='過去の版のPDF';
  $('preview-subtitle').textContent=`${wb?.displayName||wb?.fileName||''} / シート ${sheetName}`;
  $('preview-open-new').removeAttribute('href');$('pdf-frame').src='about:blank';$('preview-modal').classList.remove('hidden');$('preview-close').focus();
  try{
    clearPreviewObjectUrl();
    currentPreviewObjectUrl=await fetchPdfObjectUrl('/api/history/content-pdf',{workbookId,versionId,sheetName,snapshotId});
    $('preview-open-new').href=currentPreviewObjectUrl;
    $('pdf-frame').src=`${currentPreviewObjectUrl}#toolbar=1&navpanes=0`;
  }catch(e){closePreview();showMessage('danger','過去のPDFを開けません',userFriendlyError(e.message),e.detail||e.stack||e.message);}
}

function loadHistoryPanels(){
  loadHistoryTimeline();
  loadLayoutSnapshots();
  loadFinalArchives();
  loadSnapshotHistory();
}
