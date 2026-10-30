// 位置を指定した移動の before/after（見開きの差分）の回帰テスト（見開き PR5）。
//
// app.js の実物の classifySpreadDiff()・spreadDiffSummary()・buildSpreadLayout()・planPageMove()
// を取り出し、差分の種類（移動・左右の反転・頁だけずれた・出力先が変わった・頁未確定）を
// 確かめる。とくに「1頁のシートを動かすと間にあったページだけ左右が反転し、2頁のシートでは
// 頁がずれるだけ」になること。あわせて、試算から確定までの流れ（runSpreadMovePreview・
// confirmSpreadMove・closeSpreadMoveDialog）を api() を差し替えて動かし、取消で何も保存
// されないこと、確定が試算の指紋を baseLayout にして reorder へ送り Undo に積むこと、
// 楽観ロックで止まったら確定させずに試算し直しを案内することを確かめる。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const appPath = fileURLToPath(new URL('../app/web/app.js', import.meta.url));
const source = readFileSync(appPath, 'utf8');

function extractFunction(name) {
  const start = source.search(new RegExp(`(?:async\\s+)?function ${name}\\(`));
  assert.ok(start >= 0, `function ${name} is missing from app.js`);
  // 既定値 `options={}` の波括弧を本体と取り違えないよう、引数の閉じ括弧の後から探す。
  const open = source.indexOf('{', source.indexOf(')', start));
  let depth = 0;
  for (let i = open; i < source.length; i++) {
    if (source[i] === '{') depth++;
    else if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(`unterminated function ${name}`);
}

const { buildSpreadLayout, planPageMove, classifySpreadDiff, spreadDiffSummary } = new Function(`
  ${extractFunction('buildSpreadLayout')}
  ${extractFunction('planPageMove')}
  ${extractFunction('classifySpreadDiff')}
  ${extractFunction('spreadDiffSummary')}
  return { buildSpreadLayout, planPageMove, classifySpreadDiff, spreadDiffSummary };
`)();

// サーバー（Get-FinalBuildInputSnapshot）と同じ規則で physicalPages を作る。出力先ごとに
// 1頁＝右から数える。counts に無いページは PDF未作成として physicalPages に出さない。
function physicalPagesFor(laneIds, counts) {
  const pages = [];
  let n = 0;
  for (const id of laneIds) {
    const count = counts[id];
    if (!count) continue;
    for (let part = 1; part <= count; part++) {
      n++;
      pages.push({ pageId: id, outputPageNumber: n, sourcePdfPageNumber: part, side: n % 2 === 1 ? 'right' : 'left', spreadIndex: Math.floor(n / 2), pagePartIndex: part, pagePartCount: count, title: id });
    }
  }
  return pages;
}
function layoutsFor(volumes, counts) {
  const layouts = {};
  for (const [volume, ids] of Object.entries(volumes)) {
    if (volume === 'none') continue;
    layouts[volume] = buildSpreadLayout(ids, physicalPagesFor(ids, counts));
  }
  return layouts;
}
function diffFor(before, after, counts, moved) {
  return classifySpreadDiff(layoutsFor(before, counts), layoutsFor(after, counts), moved);
}
const kindsOf = (diff, volume) => diff.entries.filter(entry => entry.volume === volume).map(entry => `${entry.key}:${entry.kind}`);
const labels = { main: '本体', appendix: '補足', none: '未振り分け' };
const label = volume => labels[volume] || volume;

// 1. 1頁のシートを後ろへ動かすと、間にあったページだけ左右が反転する。前後のページは変わらない。
{
  const counts = { A: 1, B: 1, C: 1, D: 1, E: 1 };
  const before = { main: ['A', 'B', 'C', 'D', 'E'], none: [] };
  const plan = planPageMove(before, ['B'], 'main', 4);
  assert.equal(plan.ok, true);
  assert.deepEqual(plan.volumes.main, ['A', 'C', 'D', 'B', 'E']);
  const diff = diffFor(before, plan.volumes, counts, plan.movedIds);
  assert.deepEqual(kindsOf(diff, 'main'), ['A|1:same', 'C|1:side', 'D|1:side', 'B|1:moved', 'E|1:same'],
    'only the pages between the old and the new position may flip sides');
  assert.equal(diff.counts.side, 2, 'C (3 right → 2 left) and D (4 left → 3 right) flip; B itself stays on the left (2 → 4)');
  const summary = spreadDiffSummary(diff, label);
  assert.equal(summary.alert, true, 'any side flip must raise the alert');
  assert.match(summary.text, /左右が反転 \d+枚（本体2〜3頁）/);
  assert.match(summary.text, /最初に影響する頁: 本体2頁/);
  assert.deepEqual([...diff.volumes.main.afterSpreads].sort(), [1, 2], 'only the changed spreads are listed for the changed-only view');
  assert.equal(diff.firstAffected.spreadIndex, 1);
}

// 2. 2頁のシートを動かしても、間にあったページは頁がずれるだけで左右は反転しない。
{
  const counts = { A: 1, B: 2, C: 1, D: 1, E: 1 };
  const before = { main: ['A', 'B', 'C', 'D', 'E'], none: [] };
  const plan = planPageMove(before, ['B'], 'main', 4);
  const diff = diffFor(before, plan.volumes, counts, plan.movedIds);
  assert.deepEqual(kindsOf(diff, 'main'), ['A|1:same', 'C|1:shift', 'D|1:shift', 'B|1:moved', 'B|2:moved', 'E|1:same']);
  assert.equal(diff.counts.side, 0, 'moving a 2-page sheet must not report a side flip for the pages in between');
  assert.equal(diff.counts.shift, 2);
  const c = diff.entries.find(entry => entry.key === 'C|1');
  assert.deepEqual([c.before.outputPageNumber, c.after.outputPageNumber, c.before.side, c.after.side], [4, 2, 'left', 'left']);
  const summary = spreadDiffSummary(diff, label);
  assert.equal(summary.alert, false, 'a pure shift must not raise the side-flip alert');
  assert.match(summary.text, /頁だけずれた 2枚/);
  assert.match(summary.text, /移動 2枚/);
}

// 3. 補足6→本体3。本体の3番目以降と補足の6番目より後ろがずれ、動かしたページは出力先が変わる。
{
  const counts = { m1: 1, m2: 1, m3: 1, m4: 1, a1: 1, a2: 1, a3: 1, a4: 1, a5: 1, a6: 1, a7: 1 };
  const before = { main: ['m1', 'm2', 'm3', 'm4'], appendix: ['a1', 'a2', 'a3', 'a4', 'a5', 'a6', 'a7'], none: [] };
  const plan = planPageMove(before, ['a6'], 'main', 3);
  assert.deepEqual(plan.volumes.main, ['m1', 'm2', 'a6', 'm3', 'm4']);
  const diff = diffFor(before, plan.volumes, counts, plan.movedIds);
  assert.deepEqual(kindsOf(diff, 'main'), ['m1|1:same', 'm2|1:same', 'a6|1:moved', 'm3|1:side', 'm4|1:side']);
  assert.deepEqual(kindsOf(diff, 'appendix'), ['a1|1:same', 'a2|1:same', 'a3|1:same', 'a4|1:same', 'a5|1:same', 'a7|1:side']);
  const moved = diff.entries.find(entry => entry.key === 'a6|1');
  assert.equal(moved.volumeChanged, true);
  assert.equal(moved.fromVolume, 'appendix');
  assert.equal(moved.volume, 'main');
  assert.deepEqual(diff.changedVolumes.sort(), ['appendix', 'main']);
  const summary = spreadDiffSummary(diff, label);
  assert.match(summary.text, /本体3〜5頁/);
  assert.match(summary.text, /補足6頁/);
  assert.match(summary.text, /最初に影響する頁: 本体3頁/);
}

// 4. 未振り分けへ戻すと「出力先から外れる」。before にだけあるページも数える。
{
  const counts = { A: 1, B: 1, C: 1 };
  const before = { main: ['A', 'B', 'C'], none: [] };
  const plan = planPageMove(before, ['A'], 'none');
  const diff = diffFor(before, plan.volumes, counts, plan.movedIds);
  assert.equal(diff.counts.removed, 1);
  assert.equal(diff.removed[0].key, 'A|1');
  assert.equal(diff.removed[0].kind, 'removed');
  assert.deepEqual(kindsOf(diff, 'main'), ['B|1:side', 'C|1:side']);
  assert.ok(diff.volumes.main.beforeSpreads.has(0), 'the spread the removed page left must be listed on the before side');
  assert.match(spreadDiffSummary(diff, label).text, /出力先から外れる 1枚/);
}

// 5. 頁数がわからないページ（PDF未作成）より後ろは未確定。前にあるページの組が同じなら変化なし、
//    違えば「頁未確定」とし、数え直した位置で左右の反転を言わない。
{
  const counts = { A: 1, C: 1, D: 1, E: 1 }; // X は PDF未作成
  const before = { main: ['A', 'X', 'C', 'D', 'E'], none: [] };
  const plan = planPageMove(before, ['C'], 'main', 4);
  assert.deepEqual(plan.volumes.main, ['A', 'X', 'D', 'C', 'E']);
  const diff = diffFor(before, plan.volumes, counts, plan.movedIds);
  assert.deepEqual(kindsOf(diff, 'main'), ['A|1:same', 'X|?:same', 'D|1:uncertain', 'C|1:moved', 'E|1:same']);
  assert.equal(diff.counts.side, 0, 'an uncertain page must never be reported as a side flip');
  assert.equal(diff.counts.uncertain, 1);
  const summary = spreadDiffSummary(diff, label);
  assert.equal(summary.alert, false);
  assert.match(summary.text, /頁未確定 1枚/);
  assert.match(summary.text, /本体の\d+枚目（頁未確定）/);
}

// 6. 並びが同じ案では差分が無い。
{
  const counts = { A: 1, B: 2 };
  const before = { main: ['A', 'B'], none: [] };
  const plan = planPageMove(before, ['B'], 'main', 2);
  const diff = diffFor(before, plan.volumes, counts, plan.movedIds);
  assert.equal(diff.changed, false);
  assert.deepEqual(spreadDiffSummary(diff, label), { text: '並びは変わりません', alert: false });
}

// ---- 試算から確定までの流れ ----
const flow = new Function('hooks', `
  let spreadMoveSession=null,spreadMoveReturnFocus=null,pageMutationBusy=false,activePageVolume='all',pageVolumeAutoPicked=false,lastPageRangeAnchor='',lastPageBoardRenderSignature='';
  let selectedPages=new Set();
  const asArray=value=>Array.isArray(value)?value:[];
  const $=id=>hooks.elements[id]||null;
  const document={querySelectorAll:()=>[]};
  const sessionStorage={setItem(){}};
  const requestAnimationFrame=fn=>fn();
  const api=(...args)=>hooks.api(...args);
  const showMessage=(...args)=>hooks.messages.push(args);
  const refresh=async()=>{hooks.refreshes++;};
  const pageLaneVolumesForMove=()=>JSON.parse(JSON.stringify(hooks.lanes));
  const activeLayoutFingerprint=()=>hooks.fingerprint;
  const activeTargetVolumes=()=>['main','appendix'];
  const assignmentVolumeLabel=volume=>({main:'本体',appendix:'補足',none:'未振り分け'})[volume]||volume;
  const getPage=id=>({pageId:id,title:id});
  const userFriendlyError=message=>String(message||'');
  const pageVolumeSnapshotsEqual=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
  const waitForQueuedPageLayoutMutation=async()=>!pageMutationBusy;
  const applyPageMutationResult=response=>hooks.applied.push(response);
  const rememberPageLayoutUndo=(...args)=>{hooks.undo.push(args);return true;};
  const undoLastPageLayout=()=>{};
  const setPageSaveStatus=(kind)=>hooks.status.push(kind);
  const updateBulkSelectionLabel=()=>{},renderPageOverview=()=>{},renderPages=()=>{},syncPageSelectionUi=()=>{};
  const renderGlobalHeader=()=>{},renderNavBadges=()=>{},renderStepBar=()=>{},renderSummary=()=>{},renderDashboardOverview=()=>{},renderFinalOverview=()=>{},renderVolumeLinks=()=>{};
  const renderSpreadMoveDialog=()=>{hooks.renders++;};
  ${extractFunction('buildSpreadLayout')}
  ${extractFunction('planPageMove')}
  ${extractFunction('classifySpreadDiff')}
  ${extractFunction('spreadLayoutsFromPreview')}
  ${extractFunction('runSpreadMovePreview')}
  ${extractFunction('reloadSpreadMoveBase')}
  ${extractFunction('confirmSpreadMove')}
  ${extractFunction('closeSpreadMoveDialog')}
  function start(pageIds,target,position){
    spreadMoveSession={pageIds,target,position,packId:'pack_t',beforeVolumes:pageLaneVolumesForMove(),baseLayout:activeLayoutFingerprint(),plan:null,preview:null,diff:null,error:'',conflict:false,loading:false,saving:false,seq:0,scrolledSeq:0,changedOnly:true,pane:'after'};
    return spreadMoveSession;
  }
  return { start, runSpreadMovePreview, reloadSpreadMoveBase, confirmSpreadMove, closeSpreadMoveDialog, session:()=>spreadMoveSession, busy:()=>pageMutationBusy, selected:()=>[...selectedPages] };
`);

const COUNTS = { A: 1, B: 1, C: 1, D: 1 };
function previewResponse(beforeVolumes, afterVolumes, fingerprint) {
  const side = volumes => ({ main: { targetId: 'main', volume: 'main', physicalPages: physicalPagesFor(volumes.main, COUNTS), unknownPageIds: [] }, appendix: { targetId: 'appendix', volume: 'appendix', physicalPages: physicalPagesFor(volumes.appendix, COUNTS), unknownPageIds: [] } });
  return { ok: true, preview: { layoutFingerprint: fingerprint, before: side(beforeVolumes), after: side(afterVolumes) } };
}
function newHooks() {
  const hooks = { elements: {}, messages: [], applied: [], undo: [], status: [], calls: [], renders: 0, refreshes: 0, fingerprint: 'fp1', lanes: { main: ['A', 'B', 'C', 'D'], appendix: [], none: [] } };
  hooks.api = async (path, options) => {
    hooks.calls.push({ path, body: JSON.parse(JSON.stringify(options?.body || {})) });
    if (path === '/api/v2/outputs/spread-preview') return previewResponse(hooks.lanes, options.body.volumes, hooks.fingerprint);
    if (path === '/api/v2/items/reorder') return { ok: true, result: { layoutFingerprint: 'fp2' } };
    throw new Error(`unexpected ${path}`);
  };
  return hooks;
}

// 7. 試算は spread-preview にだけ送り、確定するまで reorder を呼ばない。試算中は確定できない。
{
  const hooks = newHooks(), ui = flow(hooks);
  const session = ui.start(['D'], 'main', 2);
  assert.equal(await ui.confirmSpreadMove(), false, 'confirm must be refused before the preview has finished');
  assert.equal(hooks.calls.length, 0);
  assert.equal(await ui.runSpreadMovePreview(), true);
  assert.deepEqual(hooks.calls.map(call => call.path), ['/api/v2/outputs/spread-preview']);
  assert.deepEqual(hooks.calls[0].body, { packId: 'pack_t', volumes: { main: ['A', 'D', 'B', 'C'], appendix: [], none: [] }, baseLayout: 'fp1' });
  assert.equal(session.diff.changed, true);
  assert.equal(session.diff.counts.side, 2, 'B and C flip; D itself stays on the left (4 → 2)');
}

// 8. 取消では何も保存しない（reorder を呼ばず、Undo も保存履歴も増えない）。
{
  const hooks = newHooks(), ui = flow(hooks);
  ui.start(['D'], 'main', 2);
  await ui.runSpreadMovePreview();
  assert.equal(ui.closeSpreadMoveDialog(false), true);
  assert.equal(ui.session(), null);
  assert.deepEqual(hooks.calls.map(call => call.path), ['/api/v2/outputs/spread-preview'], 'cancel must not call the reorder API');
  assert.equal(hooks.undo.length, 0);
  assert.equal(hooks.applied.length, 0);
  assert.equal(await ui.confirmSpreadMove(), false, 'a closed dialog must not confirm');
}

// 9. 確定は試算の指紋を baseLayout にして reorder へ送り、before/after を Undo に積む。
{
  const hooks = newHooks(), ui = flow(hooks);
  ui.start(['D'], 'main', 2);
  await ui.runSpreadMovePreview();
  assert.equal(await ui.confirmSpreadMove(), true);
  const reorder = hooks.calls.find(call => call.path === '/api/v2/items/reorder');
  assert.deepEqual(reorder.body, { packId: 'pack_t', volumes: { main: ['A', 'D', 'B', 'C'], appendix: [], none: [] }, baseLayout: 'fp1' });
  assert.equal(hooks.applied.length, 1);
  assert.equal(hooks.undo.length, 1);
  assert.deepEqual(hooks.undo[0][1], { main: ['A', 'B', 'C', 'D'], appendix: [], none: [] }, 'Undo must restore the order the preview started from');
  assert.deepEqual(hooks.undo[0][2], { main: ['A', 'D', 'B', 'C'], appendix: [], none: [] });
  assert.equal(ui.session(), null, 'a confirmed move closes the dialog');
  assert.equal(ui.busy(), false, 'the busy flag must be released after saving');
  assert.deepEqual(ui.selected(), ['D']);
}

// 10. 楽観ロック（409）で止まったら、Undo に積まず、ダイアログを残して試算し直しを案内する。
{
  const hooks = newHooks(), ui = flow(hooks);
  const session = ui.start(['D'], 'main', 2);
  await ui.runSpreadMovePreview();
  hooks.api = async (path, options) => {
    hooks.calls.push({ path, body: options?.body });
    const error = new Error('conflict'); error.code = 'structure-conflict'; error.status = 409; throw error;
  };
  assert.equal(await ui.confirmSpreadMove(), false);
  assert.equal(hooks.undo.length, 0, 'a rejected save must not be recorded for Undo');
  assert.equal(hooks.applied.length, 0);
  assert.equal(ui.session(), session, 'the dialog stays open so the user can re-run the preview');
  assert.equal(session.conflict, true);
  assert.match(session.error, /試算し直して/);
  assert.equal(ui.busy(), false);
  const before = hooks.calls.length;
  assert.equal(await ui.confirmSpreadMove(), false, 'confirm stays blocked until the preview is re-run');
  assert.equal(hooks.calls.length, before);
  // 試算し直すと、最新の並びと指紋で作り直し、確定できるようになる。
  const fresh = newHooks();
  hooks.api = fresh.api; fresh.calls = hooks.calls;
  hooks.lanes = { main: ['B', 'A', 'C', 'D'], appendix: [], none: [] }; hooks.fingerprint = 'fp9';
  fresh.lanes = hooks.lanes; fresh.fingerprint = 'fp9';
  assert.equal(await ui.reloadSpreadMoveBase(), true);
  assert.equal(hooks.refreshes, 1);
  assert.equal(session.conflict, false);
  assert.equal(session.baseLayout, 'fp9');
  assert.deepEqual(session.plan.volumes.main, ['B', 'D', 'A', 'C']);
  assert.equal(hooks.calls.at(-1).body.baseLayout, 'fp9');
}

// 11. 試算の後で並びが変わっていたら、サーバーへ送らずに試算し直しを求める。
{
  const hooks = newHooks(), ui = flow(hooks);
  const session = ui.start(['D'], 'main', 2);
  await ui.runSpreadMovePreview();
  hooks.lanes = { main: ['B', 'A', 'C', 'D'], appendix: [], none: [] };
  assert.equal(await ui.confirmSpreadMove(), false);
  assert.equal(hooks.calls.filter(call => call.path === '/api/v2/items/reorder').length, 0);
  assert.equal(session.conflict, true);
}

// 12. 試算の 409 も確定させない。古い試算の応答は捨てる。
{
  const hooks = newHooks(), ui = flow(hooks);
  const session = ui.start(['D'], 'main', 2);
  hooks.api = async () => { const error = new Error('conflict'); error.code = 'structure-conflict'; throw error; };
  assert.equal(await ui.runSpreadMovePreview(), false);
  assert.equal(session.conflict, true);
  assert.equal(session.preview, null);

  const hooks2 = newHooks(), ui2 = flow(hooks2);
  const session2 = ui2.start(['D'], 'main', 2);
  const pending = [];
  hooks2.api = (path, options) => new Promise(resolve => pending.push(() => resolve(previewResponse(hooks2.lanes, options.body.volumes, 'fp1'))));
  const first = ui2.runSpreadMovePreview();
  session2.position = 'end';
  const second = ui2.runSpreadMovePreview();
  pending[1]();
  assert.equal(await second, true);
  pending[0]();
  assert.equal(await first, false, 'a stale preview response must be dropped');
  assert.deepEqual(session2.plan.volumes.main, ['A', 'B', 'C', 'D']);
  assert.equal(session2.diff.changed, false, 'the latest choice (end = no change) must win');
}

// 13. 範囲外の位置は試算せずに知らせる。保存中は閉じない。
{
  const hooks = newHooks(), ui = flow(hooks);
  const session = ui.start(['D'], 'main', 9);
  assert.equal(await ui.runSpreadMovePreview(), false);
  assert.equal(hooks.calls.length, 0);
  assert.match(session.error, /1〜4番目/);
  session.saving = true;
  assert.equal(ui.closeSpreadMoveDialog(false), false, 'the dialog must not close while the reorder is saving');
}

// 14. 画面のつなぎ: 入口、キー操作、モーダルの扱い。見開きの中ではドラッグしない。
assert.match(extractFunction('handleBoardRowKeydown'), /openSpreadMoveForCard\(row\)/, 'M on a card must open the position move');
assert.match(extractFunction('handleSpreadCellKeydown'), /openSpreadMoveForCard\(cell\)/, 'M on a spread cell must open the position move');
assert.match(extractFunction('isAnyModalOpen'), /isSpreadMoveOpen\(\)/, 'the before/after dialog must block board shortcuts like Ctrl+Z');
assert.ok(extractFunction('beginPointerPageDrag').includes("if (pageBoardView === 'spread') return;"), 'the spread view must stay drag-free');
for (const name of ['openSpreadMoveDialog', 'runSpreadMovePreview', 'reloadSpreadMoveBase', 'closeSpreadMoveDialog']) {
  assert.ok(!extractFunction(name).includes('/api/v2/items/reorder'), `${name} must not save`);
}
assert.ok(source.includes("bind('spread-move-btn','click'"), 'the destination bar must offer the position move');
assert.ok(source.includes("bind('preview-move-position','click'"), 'the preview must offer the position move');

console.log('spread diff regression ok');
