// 変更履歴・比較まわりのレビュー指摘（H1・H3・M3）の回帰テスト。
//
// H1: 登録画面の「見た目変更 N」は structure の要約で描く。要約が今の変換PDF
//     （lastRenderedSnapshotId / lastRenderedVersionId）の比較でなければ出さない。
// H3: 同じ原稿で比較ペアを切り替えたとき、前のペアの応答で差分画面を上書きしない。
//     「確認済みにする」の応答も、開始時の画面にだけ反映する。
// M3: 差分詳細のクライアントキャッシュは ready の応答だけを持つ。refresh() で捨てる。
//
// app.js の実物の関数を取り出し、偽の環境で動かして確かめる。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
import vm from 'node:vm';

const appPath = fileURLToPath(new URL('../app/web/app.js', import.meta.url));
const source = readFileSync(appPath, 'utf8');

function extractFunction(name) {
  let start = source.indexOf(`async function ${name}(`);
  if (start < 0) start = source.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `function ${name} is missing from app.js`);
  const open = source.indexOf('{', source.indexOf(')', start));
  let depth = 0;
  for (let i = open; i < source.length; i++) {
    if (source[i] === '{') depth++;
    else if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(`unterminated function ${name}`);
}

function deferred() {
  let resolve, reject;
  const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

const stubs = `
var state = null;
function asArray(value){return value==null?[]:(Array.isArray(value)?value:[value]);}
function badge(text,cls){return '['+cls+':'+text+']';}
function sourceUnitLabel(){return 'シート';}
`;

// ---- H1: 要約が今の変換PDFの比較でなければ出さない ----------------------------
{
  const ctx = vm.createContext({});
  vm.runInContext(stubs + [
    'getWorkbook', 'changeSummaryFor', 'changedSheetSet', 'addedSheetSet', 'unknownSheetSet',
    'comparisonSummaryForWorkbook', 'pageChangeBadge', 'pageIsChanged',
  ].map(extractFunction).join('\n'), ctx);
  const summary = (ids, extra = {}) => ({ status: 'complete', changedSheets: ['S1', 'S2'], addedSheets: [], removedSheets: [], unknownSheets: [], ...ids, ...extra });
  const setState = (workbook, cs) => {
    ctx.state = { inputHistoryEnabled: true, structure: { workbooks: [workbook] }, changeSummaries: { w1: cs } };
  };
  const wb = { workbookId: 'w1', lastRenderedSnapshotId: 'snap2', lastRenderedVersionId: 'ver2' };

  // 一致していれば今までどおり出す。
  setState(wb, summary({ currentSnapshotId: 'snap2', currentVersionId: 'ver2' }));
  assert.equal(ctx.comparisonSummaryForWorkbook(wb), '見た目変更 2シート');
  assert.match(ctx.pageChangeBadge({ workbookId: 'w1', sheetName: 'S1' }), /変更あり/);
  assert.equal(ctx.pageIsChanged({ workbookId: 'w1', sheetName: 'S3' }), false);

  // 同じ snapshot に新しい version を作った（原稿を変えずに作り直した）後の古い要約。
  setState(wb, summary({ currentSnapshotId: 'snap2', currentVersionId: 'ver1' }));
  assert.equal(ctx.changeSummaryFor('w1'), null);
  assert.equal(ctx.comparisonSummaryForWorkbook(wb), '');
  assert.equal(ctx.pageChangeBadge({ workbookId: 'w1', sheetName: 'S1' }), '');
  // 比較できない扱い（絞り込みでは隠さない）。
  assert.equal(ctx.pageIsChanged({ workbookId: 'w1', sheetName: 'S3' }), true);

  // 別の snapshot の要約、ID の無い旧形式の要約も出さない。
  setState(wb, summary({ currentSnapshotId: 'snap1', currentVersionId: 'ver2' }));
  assert.equal(ctx.comparisonSummaryForWorkbook(wb), '');
  setState(wb, summary({}));
  assert.equal(ctx.comparisonSummaryForWorkbook(wb), '');
  // まだ変換PDFが無い原稿に要約だけ残っていても出さない。
  setState({ workbookId: 'w1' }, summary({ currentSnapshotId: '', currentVersionId: '' }));
  assert.equal(ctx.changeSummaryFor('w1'), null);

  // 今の版の unavailable 要約は「見た目を比較できません」と出す（前回の変更数を出さない）。
  setState(wb, summary({ currentSnapshotId: 'snap2', currentVersionId: 'ver2' }, { status: 'unavailable', changedSheets: [] }));
  assert.equal(ctx.comparisonSummaryForWorkbook(wb), '見た目を比較できません');
  assert.equal(ctx.pageChangeBadge({ workbookId: 'w1', sheetName: 'S1' }), '');
}

// ---- H3: ペアを切り替えたら前の応答で上書きしない ---------------------------
async function testOpenDiffDetailSerial() {
  const ctx = vm.createContext({ URLSearchParams });
  const pending = new Map();
  const rendered = [];
  vm.runInContext(`
var diffViewState = {};
var diffDetailOpenSerial = 0;
var diffReturnFocus = null;
var modalOpen = false;
var document = { activeElement: null, body: { style: {} } };
function $(id){
  if(id==='diff-modal')return { classList:{ remove(){ modalOpen=true; }, add(){ modalOpen=false; }, contains(){ return !modalOpen; } } };
  return null;
}
function isDiffModalOpen(){return modalOpen;}
function clearDiffBrowserResources(){}
function workbookDisplayName(){return 'wb';}
function getWorkbook(){return {};}
function setDiffPaneEmpty(){}
function ensureDiffPdfJs(){return Promise.resolve();}
function getDiffAnalysisWorker(){}
function userFriendlyError(message){return message;}
` + extractFunction('diffDetailRequestPath') + '\n' + extractFunction('openDiffDetail'), ctx);
  ctx.fetchDiffDetailResponse = path => { const d = deferred(); pending.set(path, d); return d.promise; };
  ctx.renderDiffDetail = detail => rendered.push(detail);

  const first = ctx.openDiffDetail('w1', null, { fromSnapshotId: 'A', toSnapshotId: 'B' });
  const second = ctx.openDiffDetail('w1', null, { fromSnapshotId: 'B', toSnapshotId: 'C' });
  const pathAB = [...pending.keys()].find(p => p.includes('fromSnapshotId=A'));
  const pathBC = [...pending.keys()].find(p => p.includes('fromSnapshotId=B'));
  assert.ok(pathAB && pathBC, 'both pairs were requested');
  // 新しいペアが先に届き、古いペアが後から届く。
  pending.get(pathBC).resolve({ detail: { status: 'ready', pair: 'B-C' } });
  await second;
  pending.get(pathAB).resolve({ detail: { status: 'ready', pair: 'A-B' } });
  await first;
  assert.deepEqual(rendered.map(d => d.pair), ['B-C'], 'stale A-B response must not overwrite B-C');
  assert.equal(ctx.diffViewState.fromSnapshotId, 'B');

  // 古いペアの失敗応答も描かない。
  rendered.length = 0; pending.clear();
  const third = ctx.openDiffDetail('w1', null, { fromSnapshotId: 'A', toSnapshotId: 'B' });
  const fourth = ctx.openDiffDetail('w1', null, { fromSnapshotId: 'B', toSnapshotId: 'C' });
  [...pending.entries()].find(([p]) => p.includes('fromSnapshotId=A'))[1].reject(new Error('timeout'));
  await third;
  assert.equal(rendered.length, 0, 'stale failure must not render');
  [...pending.entries()].find(([p]) => p.includes('fromSnapshotId=B'))[1].resolve({ detail: { status: 'ready', pair: 'B-C' } });
  await fourth;
  assert.deepEqual(rendered.map(d => d.pair), ['B-C']);
}

async function testReviewToggleKeepsOwnDetail() {
  const ctx = vm.createContext({});
  vm.runInContext(stubs + `
var diffViewState = {};
var diffDetailResponseCache = new Map();
var renders = 0;
function diffCurrentSheet(){return asArray(diffViewState.detail?.sheets)[0]||null;}
async function runBusy(button,fn){return fn();}
function diffFilteredSheets(){return asArray(diffViewState.detail?.sheets);}
function preferredDiffPageIndex(){return 0;}
function renderDiffSummary(){renders++;}
function renderDiffSheetList(){}
function renderDiffPage(){}
` + extractFunction('toggleDiffSheetReviewed'), ctx);
  const reviewRequest = deferred();
  ctx.api = () => reviewRequest.promise;
  const oldDetail = { comparison: { baselineVersionId: 'v1', currentVersionId: 'v2' }, sheets: [{ sheetKey: 'k1', confirmed: false }], summary: { confirmed: 0 } };
  const newDetail = { comparison: {}, sheets: [{ sheetKey: 'k1', confirmed: false }], summary: { confirmed: 0 } };
  ctx.diffViewState.detail = oldDetail;
  ctx.diffDetailResponseCache.set('x', Promise.resolve());
  const toggling = ctx.toggleDiffSheetReviewed(null);
  // 応答待ちの間に別の比較へ切り替わった。
  ctx.diffViewState.detail = newDetail;
  reviewRequest.resolve({ review: { confirmedSheetKeys: ['k1'] } });
  await toggling;
  assert.equal(newDetail.sheets[0].confirmed, false, 'review must not be written into another comparison');
  assert.equal(newDetail.summary.confirmed, 0);
  assert.equal(ctx.renders, 0);
  assert.equal(ctx.diffDetailResponseCache.size, 0, 'review change still invalidates cached details');

  // 切り替わっていなければ今までどおり反映する。
  const again = deferred();
  ctx.api = () => again.promise;
  ctx.diffViewState.detail = oldDetail;
  const toggling2 = ctx.toggleDiffSheetReviewed(null);
  again.resolve({ review: { confirmedSheetKeys: ['k1'] } });
  await toggling2;
  assert.equal(oldDetail.sheets[0].confirmed, true);
  assert.equal(oldDetail.summary.confirmed, 1);
  assert.equal(ctx.renders, 1);
}

// ---- M3: ready 以外の応答はキャッシュに残さない --------------------------------
async function testDetailCacheKeepsOnlyReady() {
  const ctx = vm.createContext({ setTimeout, clearTimeout, AbortController });
  vm.runInContext(`
var diffDetailResponseCache = new Map();
var DIFF_DETAIL_TIMEOUT_MS = 1000, DIFF_DETAIL_CACHE_LIMIT = 8;
var diffViewState = {};
function getWorkbook(){return {lastRenderedSnapshotId:'s',lastRenderedVersionId:'v'};}
function isDiffModalOpen(){return false;}
function sleep(){return Promise.resolve();}
` + extractFunction('fetchDiffDetailResponse'), ctx);
  const responses = [{ detail: { status: 'unavailable' } }, { detail: { status: 'ready' } }];
  let calls = 0;
  ctx.api = async () => responses[Math.min(calls++, responses.length - 1)];
  const first = await ctx.fetchDiffDetailResponse('/api/history/diff-detail?workbookId=w1', 'w1');
  assert.equal(first.detail.status, 'unavailable');
  assert.equal(ctx.diffDetailResponseCache.size, 0, 'unavailable response must not stay cached');
  const second = await ctx.fetchDiffDetailResponse('/api/history/diff-detail?workbookId=w1', 'w1');
  assert.equal(second.detail.status, 'ready');
  assert.equal(calls, 2, 'reopening after unavailable must refetch');
  const third = await ctx.fetchDiffDetailResponse('/api/history/diff-detail?workbookId=w1', 'w1');
  assert.equal(third.detail.status, 'ready');
  assert.equal(calls, 2, 'ready response is served from cache');

  const refresh = extractFunction('refresh');
  assert.ok(refresh.includes('diffDetailResponseCache.clear()'), 'refresh() must clear the diff detail cache');
}

await testOpenDiffDetailSerial();
await testReviewToggleKeepsOwnDetail();
await testDetailCacheKeepsOnlyReady();
console.log('history-compare-review-regression: OK');
