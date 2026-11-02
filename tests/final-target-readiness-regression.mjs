// 出力前チェックの「出力先」が、振り分け済みの出力先を「ページがありません」と誤って
// 判定した不具合の回帰テスト。
//
// readiness の pageCount は manifest に入った物理出力頁数で、変換PDFが未作成だと
// 振り分け済みでも0になる。出力先にページがあるかは振り分け件数 itemCount で見る。
// また readiness が届く前（volumeReadiness の既定値）は判定せず、確認中として扱う。
//
// app.js の実物の関数を取り出し、偽のDOMで動かして確かめる。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const appPath = fileURLToPath(new URL('../app/web/app.js', import.meta.url));
const source = readFileSync(appPath, 'utf8');

function extractFunction(name) {
  const start = source.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `function ${name} is missing from app.js`);
  // 既定値 `entries=[]` の括弧を本体と取り違えないよう、引数の閉じ括弧の後から探す。
  const open = source.indexOf('{', source.indexOf(')', start));
  let depth = 0;
  for (let i = open; i < source.length; i++) {
    if (source[i] === '{') depth++;
    else if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(`unterminated function ${name}`);
}

// 必須原稿・変換PDFの問題を一時的に起こすための切り替え。
const overrides = { conversionPending: false };
const fakeElement = () => ({ className: '', innerHTML: '', textContent: '', dataset: {}, querySelectorAll: () => [] });
const elements = { 'final-preflight-summary': fakeElement(), 'final-preflight-list': fakeElement(), 'final-preflight-announce': fakeElement() };
const env = {
  $: id => elements[id] || null,
  asArray: value => (Array.isArray(value) ? value : value == null ? [] : [value]),
  escapeHtml: value => String(value ?? ''),
  escapeAttr: value => String(value ?? ''),
  badge: (text, cls) => `<span class="${cls}">${text}</span>`,
  setHtmlIfChanged: (el, html) => { if (el.innerHTML === html) return false; el.innerHTML = html; return true; },
  formatDateTime: value => String(value ?? ''),
  setActiveView: () => {},
  // 必須原稿・変換PDF・ページ構成・出力ファイル名は問題なしにしておく。
  workbooksForActivePreset: () => [{ workbookId: 'w1', status: 'ready' }],
  activeSourceRequirements: () => [],
  isLatestPdfWorkbook: () => !overrides.conversionPending,
  unresolvedPageCount: () => 0,
  packForPreset: () => ({ settings: { outputFileNamePattern: '{packName}_{targetName}.pdf' } }),
  finalPreflightCheckedAt: '',
  state: { finalReadiness: {} },
  activePackRecord: () => null,
  activePackId: 'pack1',
  activePreset: 'pack1',
};
const names = Object.keys(env);
const factory = new Function(...names, [
  extractFunction('volumeReadiness'),
  extractFunction('targetItemCount'),
  extractFunction('targetReadinessPending'),
  extractFunction('renderFinalPreflight'),
  'return { volumeReadiness, targetItemCount, targetReadinessPending, renderFinalPreflight };',
].join('\n'));
const { volumeReadiness, targetItemCount, targetReadinessPending, renderFinalPreflight } = factory(...names.map(name => env[name]));

const target = (targetId, required = true) => ({ targetId, displayName: targetId, required });
const entry = (targetId, ready, required = true) => ({ target: target(targetId, required), volume: targetId, ready });
const targetCheck = result => result.checks.find(check => check.label === '出力先');

// ---- itemCount の読み方 ----
assert.equal(targetItemCount({ pageCount: 0, snapshot: { itemCount: 3 } }), 3, 'readiness の snapshot.itemCount を使う');
assert.equal(targetItemCount({ pageCount: 0, itemCount: 2 }), 2, '直下の itemCount を使う');
assert.equal(targetItemCount({ pageCount: 5 }), 5, 'itemCount の無い古い応答は pageCount で代用する');
assert.equal(targetItemCount({ pageCount: 5, snapshot: { itemCount: 0 } }), 0, 'itemCount が0なら pageCount を見ない');
assert.equal(targetItemCount(undefined), 0);

// ---- 未取得の印 ----
const fallback = volumeReadiness('ja-main');
assert.equal(fallback.loaded, false, 'readiness 未取得の既定値には loaded:false を付ける');
assert.ok(targetReadinessPending(fallback));
assert.ok(!targetReadinessPending({ pageCount: 0, snapshot: { itemCount: 0 } }), 'サーバーの応答は確認済みとして扱う');

// ---- 振り分け済みで物理頁0（変換PDF未作成）: 「ページがありません」と言わない ----
{
  const ready = { pageCount: 0, displayState: 'blocked', snapshot: { itemCount: 4 },
    blockers: [{ code: 'required-source-not-rendered', message: '必須原稿の変換PDFが未作成です。' }] };
  const result = renderFinalPreflight([entry('ja-main', ready)]);
  const check = targetCheck(result);
  assert.equal(check.state, 'ok', `振り分け済みの出力先は空ではない: ${check.detail}`);
  assert.ok(!check.detail.includes('ページがありません'));
  assert.equal(result.hardProblemCount, 0, 'required-source- の blocker は出力先の問題に重ねて数えない');
}

// ---- 本当に振り分けが0件: 今までどおり danger ----
{
  const ready = { pageCount: 0, displayState: 'blocked', snapshot: { itemCount: 0 }, blockers: [{ code: 'no-pages', message: 'ページがありません' }] };
  const check = targetCheck(renderFinalPreflight([entry('ja-main', ready)]));
  assert.equal(check.state, 'danger');
  assert.ok(check.detail.includes('ja-mainにページがありません'));
  assert.equal(check.view, 'pages');
}

// ---- 任意の出力先が空なのは問題にしない ----
{
  const ready = { pageCount: 0, displayState: 'not-built', snapshot: { itemCount: 0 }, blockers: [{ code: 'no-pages' }] };
  assert.equal(targetCheck(renderFinalPreflight([entry('ja-appendix', ready, false)])).state, 'ok');
}

// ---- required-source- 以外の blocker は振り分け済みの出力先の問題として数える ----
{
  const ready = { pageCount: 2, displayState: 'blocked', snapshot: { itemCount: 2 }, blockers: [{ code: 'page-source-missing', message: '原稿が見つかりません' }] };
  const result = renderFinalPreflight([entry('ja-main', ready)]);
  assert.equal(targetCheck(result).state, 'danger');
  assert.ok(targetCheck(result).detail.includes('1件の出力先'));
  assert.equal(result.hardProblemCount, 1);
}

// ---- 振り分け0件の出力先の blocker は「問題あり」と二重に数えない ----
{
  const ready = { pageCount: 0, displayState: 'blocked', snapshot: { itemCount: 0 }, blockers: [{ code: 'page-source-missing' }] };
  const result = renderFinalPreflight([entry('ja-appendix', ready, false)]);
  assert.equal(targetCheck(result).state, 'ok');
}

// ---- サーバーで確認に失敗した出力先は「出力できません」側で数える ----
{
  const ready = { readinessError: true, pageCount: 0, displayState: 'blocked', blockers: [{ code: 'readiness-failed', message: '確認できませんでした' }] };
  const check = targetCheck(renderFinalPreflight([entry('ja-main', ready)]));
  assert.equal(check.state, 'danger');
  assert.ok(!check.detail.includes('ページがありません'), check.detail);
}

// ---- readiness 未取得: 確認中で、出力を止める問題には数えない ----
{
  const result = renderFinalPreflight([entry('ja-main', volumeReadiness('ja-main')), entry('ja-appendix', volumeReadiness('ja-appendix'))]);
  const check = targetCheck(result);
  assert.equal(check.state, 'attention', '未取得の出力先は danger にしない');
  assert.equal(check.pending, true);
  assert.ok(!check.detail.includes('ページがありません'), check.detail);
  assert.equal(result.hardProblemCount, 0, '確認中は hardProblem に数えない');
}

// ---- 一部だけ届いていて、届いた側に本当の問題があれば danger を優先する ----
{
  const empty = { pageCount: 0, displayState: 'blocked', snapshot: { itemCount: 0 }, blockers: [{ code: 'no-pages' }] };
  const result = renderFinalPreflight([entry('ja-main', empty), entry('ja-appendix', volumeReadiness('ja-appendix'))]);
  assert.equal(targetCheck(result).state, 'danger');
  assert.equal(result.hardProblemCount, 1);
}

// ---- 確認中だけのとき: 問題の件数に入れず、「確認が必要」とも言わない ----
{
  const result = renderFinalPreflight([entry('ja-main', volumeReadiness('ja-main'))]);
  assert.equal(result.problemCount, 0, '確認中は問題の件数に入れない');
  assert.equal(result.ready, false, '確認中は準備完了にしない');
  const summary = elements['final-preflight-summary'];
  assert.ok(summary.className.includes('checking'), summary.className);
  assert.ok(summary.innerHTML.includes('確認中'), summary.innerHTML);
  assert.ok(!summary.innerHTML.includes('確認が必要'), summary.innerHTML);
  assert.ok(!summary.innerHTML.includes('項目を確認してください'), summary.innerHTML);
  assert.ok(!elements['final-preflight-announce'].textContent.includes('確認が必要'), elements['final-preflight-announce'].textContent);
}

// ---- 確認中と本当の問題（変換PDFの未作成）が同時にあるとき: 本当の問題だけを1件と数える ----
{
  overrides.conversionPending = true;
  const result = renderFinalPreflight([entry('ja-main', volumeReadiness('ja-main'))]);
  overrides.conversionPending = false;
  assert.equal(result.problemCount, 1, '確認中を除いた本当の問題だけを数える');
  assert.equal(result.ready, false);
  const summary = elements['final-preflight-summary'];
  assert.ok(summary.className.includes('attention'), summary.className);
  assert.ok(summary.innerHTML.includes('出力前に確認が必要です'), summary.innerHTML);
  assert.ok(summary.innerHTML.includes('1項目を確認してください'), summary.innerHTML);
  assert.ok(elements['final-preflight-announce'].textContent.includes('1項目の確認が必要'));
}

// ---- 取り直しても確認できなかった出力先: danger で数え、文言は確定した問題と分ける ----
{
  const ready = { readinessError: true, pageCount: 0, displayState: 'blocked', blockers: [{ code: 'readiness-failed', message: '確認できませんでした' }] };
  const result = renderFinalPreflight([entry('ja-main', ready)]);
  assert.equal(targetCheck(result).state, 'danger');
  assert.equal(targetCheck(result).detail, '出力先の状態を確認できませんでした（再確認）');
  assert.equal(result.hardProblemCount, 1, '確認できなかった出力先は出力を止める');
}

// ---- 出力先カードとまとめて出力も pageCount ではなく itemCount で見る ----
const renderFinalBody = source.slice(source.indexOf('const renderFreshness=(ready,card)=>{'), source.indexOf('function renderFinalPreflight('));
assert.ok(renderFinalBody.includes('targetItemCount(ready)===0'), 'card empty state must use itemCount');
assert.ok(renderFinalBody.includes('btn.disabled=pending||empty||blockers.length>0'), 'card build button must use itemCount and wait for readiness');
assert.ok(renderFinalBody.includes("fix.classList.toggle('hidden',pending||empty||blockers.length===0)"), 'card fix button must use itemCount');
assert.ok(renderFinalBody.includes('!targetReadinessPending(ready)&&targetItemCount(ready)>0'), 'build-all must use itemCount');
assert.ok(!/Number\(ready\.pageCount\|\|0\)===0/.test(renderFinalBody), 'card must not judge emptiness by physical pageCount');
assert.ok(source.includes("entry.target?.required!==false||targetItemCount(entry.ready)>0"), 'optional targets must be shown by itemCount');

// ---- 出力完了の前に始まった readiness 要求を、完了後の表示に使わない ----
// 症状: PDFを何度作成しても「再出力が必要」のまま、再起動で直る。
// loadFinalPanels({force:true}) が進行中の古い要求を使い回し、出力前の応答で表示していた。
async function readinessHarness() {
  const requests = [];
  const deferred = () => { let resolve; const promise = new Promise(r => { resolve = r; }); return { promise, resolve }; };
  const api = url => {
    if (!url.includes('readiness')) return Promise.resolve({});
    const d = deferred();
    requests.push({ url, resolve: status => d.resolve({ volumes: { 'ja-main': { status, displayState: status } } }) });
    return d.promise;
  };
  const env = {
    api,
    configured: () => true,
    activePackRecord: () => ({ packId: 'custom1' }),
    activePackId: 'custom1', activePreset: 'custom1',
    targetVolume: id => id,
    activeTargetVolumes: () => ['ja-main'],
    log: (...args) => { throw new Error(`unexpected log: ${args.join(' ')}`); },
    activeView: 'final',
    renderGlobalHeader() {}, renderNavBadges() {}, renderStepBar() {}, renderFinalOverview() {}, renderVolumeLinks() {},
    renderSpreadBoardIfVisible() {},
    loadFinalArchives: () => Promise.resolve(), loadPackReview: () => Promise.resolve(),
  };
  const names = Object.keys(env);
  const factory = new Function(...names, [
    'let state = { finalReadiness: {} }, finalPreflightCheckedAt = "";',
    'const finalReadinessInFlight = new Map();',
    'let finalReadinessEpoch = 0, finalPanelsInFlight = null;',
    extractFunction('finalReadinessKey'),
    extractFunction('advanceFinalReadinessEpoch'),
    // extractFunction は `function` から切り出すので async を付け直す。
    'async ' + extractFunction('loadFinalReadiness'),
    'async ' + extractFunction('loadFinalPanels'),
    'return { loadFinalPanels, loadFinalReadiness, advanceFinalReadinessEpoch, readiness: () => state.finalReadiness.custom1?.volumes?.["ja-main"]?.status };',
  ].join('\n'));
  return { requests, ...factory(...names.map(name => env[name])) };
}
const flush = () => new Promise(r => setTimeout(r, 0));

{
  // 出力前に読込が始まり、出力完了後に force で読み直す。
  const h = await readinessHarness();
  const before = h.loadFinalPanels({ render: false });
  assert.equal(h.requests.length, 1);
  const after = h.loadFinalPanels({ force: true, render: false });
  assert.equal(h.requests.length, 2, 'force は進行中の readiness 要求を使い回さず取り直す');
  h.requests[0].resolve('needs-rebuild'); // 出力前の応答が後から届く
  h.requests[1].resolve('built');
  await Promise.all([before, after]); await flush();
  assert.equal(h.readiness(), 'built', '出力前に始まった応答で「再出力が必要」に戻さない');
  assert.equal(h.requests.length, 2, '新しい要求が既にあれば捨てた応答の再取得はそれを使い回す');
}

{
  // 古い応答が新しい応答より後に届いても上書きしない。
  const h = await readinessHarness();
  const before = h.loadFinalPanels({ render: false });
  const after = h.loadFinalPanels({ force: true, render: false });
  h.requests[1].resolve('built');
  await after; await flush();
  h.requests[0].resolve('needs-rebuild');
  await before; await flush();
  assert.equal(h.readiness(), 'built');
  assert.equal(h.requests.length, 2);
}

{
  // 世代ずれで捨てた応答は、取り直しが無ければ1回だけ取り直す。
  const h = await readinessHarness();
  const first = h.loadFinalReadiness({ render: false });
  h.advanceFinalReadinessEpoch();
  h.requests[0].resolve('needs-rebuild');
  await first; await flush();
  assert.equal(h.readiness(), undefined, '世代ずれの応答は使わない');
  assert.equal(h.requests.length, 2, '捨てた応答の代わりを1回取り直す');
  h.advanceFinalReadinessEpoch();
  h.requests[1].resolve('needs-rebuild');
  await flush(); await flush();
  assert.equal(h.requests.length, 2, '取り直しは1回だけ（繰り返さない）');
  // 取り直しが世代内に届けば反映する。
  const again = h.loadFinalReadiness({ render: false });
  h.requests[2].resolve('built');
  await again; await flush();
  assert.equal(h.readiness(), 'built');
}

// ---- state をまるごと差し替えても、readiness が届く前に描かない ----
// /api/state の finalReadiness は ecm/bod/dmm だけで、カスタム一式の readiness が消える。
// readinessResponses: 呼ばれた順に返す readiness 応答（'error' は通信失敗）。
function stateHarness(readinessResponses = []) {
  const calls = [];
  const observed = [];
  const queue = readinessResponses.slice();
  const api = url => {
    calls.push(url);
    if (url.includes('readiness')) {
      const next = queue.length ? queue.shift() : { targets: { 'ja-main': { status: 'built', displayState: 'built' } } };
      return next === 'error' ? Promise.reject(new Error('network')) : Promise.resolve(JSON.parse(JSON.stringify(next)));
    }
    if (url === '/api/state') return Promise.resolve({ finalReadiness: { ecm: { volumes: {} }, bod: { volumes: {} }, dmm: { volumes: {} } } });
    if (url.includes('/review')) return Promise.resolve({ state: { finalReadiness: { ecm: { volumes: {} } } }, review: { packId: 'custom1' } });
    return Promise.resolve({});
  };
  const logs = [];
  const env = {
    api, observed, logs,
    configured: () => true,
    activePackRecord: () => ({ packId: 'custom1' }),
    activePackId: 'custom1', activePreset: 'custom1',
    targetVolume: id => id,
    activeTargetVolumes: () => ['ja-main'],
    log: (...args) => logs.push(args.join(' ')),
    activeView: 'final',
    renderGlobalHeader() {}, renderNavBadges() {}, renderStepBar() {}, renderVolumeLinks() {}, renderPackReview() {},
    renderSpreadBoardIfVisible() {},
    shouldPreserveEditorDom: () => false,
    loadFilesSilently: () => Promise.resolve([]),
    runBusy: (button, fn) => fn(),
    showMessage() {},
    $: () => null,
    asArray: value => (Array.isArray(value) ? value : value == null ? [] : [value]),
  };
  const names = Object.keys(env);
  const factory = new Function(...names, [
    'let state = { finalReadiness: { custom1: { volumes: { "ja-main": { status: "built", displayState: "built" } } } } }, finalPreflightCheckedAt = "";',
    'const finalReadinessInFlight = new Map();',
    'let finalReadinessEpoch = 0, updateMonitorInFlight = false, renderJobActive = false, activeFinalJobId = "", lastUpdateScanAt = 0, packReviewState = null;',
    // 描画した時点の出力先の状態を記録する。
    'function renderFinalOverview(){ observed.push(volumeReadiness("ja-main")); }',
    'function renderAll(){ renderFinalOverview(); }',
    extractFunction('normalizeStatePayload'),
    extractFunction('volumeReadiness'),
    extractFunction('targetReadinessPending'),
    extractFunction('finalReadinessKey'),
    extractFunction('finalReadinessLoaded'),
    extractFunction('advanceFinalReadinessEpoch'),
    'async ' + extractFunction('replaceStatePayload'),
    'async ' + extractFunction('loadFinalReadiness'),
    'async ' + extractFunction('scanUpdatesSilently'),
    'async ' + extractFunction('performPackReviewAction'),
    'return { scanUpdatesSilently, performPackReviewAction, loadFinalReadiness, volume: () => state.finalReadiness.custom1?.volumes?.["ja-main"] };',
  ].join('\n'));
  return { calls, observed, logs, ...factory(...names.map(name => env[name])) };
}

for (const [label, run] of [['scanUpdatesSilently', h => h.scanUpdatesSilently({ withFiles: true })], ['performPackReviewAction', h => h.performPackReviewAction('approve', null)]]) {
  const h = stateHarness();
  await run(h); await flush();
  assert.ok(h.observed.length > 0, `${label} は描き直す`);
  assert.ok(h.observed.every(ready => !targetReadinessPending(ready)), `${label}: readiness が届く前に描かない`);
  assert.equal(h.volume()?.status, 'built');
  assert.equal(h.calls.filter(url => url.includes('readiness')).length, 1, `${label}: カスタム一式の readiness を1回取り直す`);
}

// 差し替え箇所が同じ順序（区切り番号→差し替え→readiness 取得→描画）を使っていること。
for (const name of ['scanUpdatesSilently', 'scanAndRefresh', 'performPackReviewAction', 'saveSourceMetadata']) {
  const body = extractFunction(name);
  assert.ok(!/state\s*=\s*normalizeStatePayload/.test(body), `${name} は state を直接差し替えない`);
  const replaced = body.indexOf('await replaceStatePayload(');
  assert.ok(replaced >= 0 && replaced < body.indexOf('renderAll('), `${name} は readiness を待ってから描く`);
}
{
  const body = extractFunction('refreshFinalPreflight');
  assert.ok(body.indexOf('advanceFinalReadinessEpoch()') < body.indexOf('normalizeStatePayload('));
  assert.ok(body.indexOf('await loadFinalPanels(') < body.indexOf('loadFiles('), '差し替えと readiness の取得の間に原稿一覧の読込を挟まない');
}

// ---- readiness-failed が1回だけ返る: 1回だけ取り直し、その間 danger を描かない ----
const failed = { targets: { 'ja-main': { readinessError: true, canBuild: false, pageCount: 0, displayState: 'blocked', blockers: [{ code: 'readiness-failed', message: '確認できませんでした' }] } } };
{
  const h = stateHarness([failed]);
  await h.loadFinalReadiness({ render: true }); await flush();
  assert.equal(h.calls.filter(url => url.includes('readiness')).length, 2, '確認できなかったら1回だけ取り直す');
  assert.ok(h.observed.length > 0);
  assert.ok(h.observed.every(ready => !ready.readinessError || targetReadinessPending(ready)), '取り直しの間は確認中として描く');
  assert.equal(h.volume().status, 'built', '取り直しが成功すれば届いた状態で描く');
  assert.ok(!h.volume().readinessError);
}

// ---- 2回続けて失敗: danger として数える ----
{
  const h = stateHarness([failed, failed]);
  await h.loadFinalReadiness({ render: false }); await flush();
  assert.equal(h.calls.filter(url => url.includes('readiness')).length, 2, '取り直しは1回だけ');
  const ready = h.volume();
  assert.ok(ready.readinessError && !targetReadinessPending(ready), '2回失敗したら確認中で隠さない');
  const result = renderFinalPreflight([entry('ja-main', ready)]);
  assert.equal(targetCheck(result).state, 'danger');
  assert.equal(result.hardProblemCount, 1);
}

// ---- 取り直しの要求自体が失敗: 確認中のまま止めない ----
{
  const h = stateHarness([failed, 'error']);
  await h.loadFinalReadiness({ render: false }); await flush();
  const ready = h.volume();
  assert.ok(ready.readinessError && !targetReadinessPending(ready), '取り直しが通信失敗でも確認中のまま止めない');
  assert.equal(h.logs.length, 1);
}

// ---- entry はあるのに出力先が欠けている: 確認中のまま止めない ----
{
  const h = stateHarness([{ targets: {} }, { targets: {} }]);
  await h.loadFinalReadiness({ render: false }); await flush();
  assert.equal(h.calls.filter(url => url.includes('readiness')).length, 2, '欠けた出力先も1回だけ取り直す');
  const ready = h.volume();
  assert.ok(ready && !targetReadinessPending(ready), '欠けた出力先を確認中のままにしない');
  assert.equal(targetCheck(renderFinalPreflight([entry('ja-main', ready)])).state, 'danger');
}
{
  const h = stateHarness([{ targets: {} }]);
  await h.loadFinalReadiness({ render: false }); await flush();
  assert.equal(h.volume().status, 'built', '取り直しで届けばそのまま使う');
}

console.log('final target readiness regression ok');
