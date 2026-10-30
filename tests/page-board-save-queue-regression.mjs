// ページ構成の保存の待ち行列の回帰テスト（H-4）。
//
// saveBoardOrder() は並べ替えの保存を boardSavePromise の直列の待ち行列へつなぐ。保存に
// 失敗すると、その Promise は false で解決したまま次の保存まで残る。元に戻す・やり直し・
// シート順の並べ替え・一式の切り替えがそれを直接 await していたため、1回の失敗の後は
// 保存が終わっていても無言で止まり続けた。
//
// app.js の実物の saveBoardOrder() と awaitBoardSave() を取り出し、api() を差し替えて
// 「失敗した保存の後は待たずに進める」「保存中なら、その保存の結果を待つ」ことを確かめる。
// あわせて、出力・共有コピーの前に並べ替え以外のページ変更も待つこと（M-14）を確かめる。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const appPath = fileURLToPath(new URL('../app/web/app.js', import.meta.url));
const source = readFileSync(appPath, 'utf8');

function extractFunction(name) {
  const start = source.search(new RegExp(`(?:async\\s+)?function ${name}\\(`));
  assert.ok(start >= 0, `function ${name} is missing from app.js`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let i = open; i < source.length; i++) {
    if (source[i] === '{') depth++;
    else if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(`unterminated function ${name}`);
}

const harness = new Function('hooks', `
  let boardSaveTimer=null,boardSavePromise=Promise.resolve(true),boardSaveRevision=0,pendingBoardSaveVolumes=null;
  let boardSaveMutationBusy=false,pageMutationBusy=false,pageLayoutHistoryBusy=false,pendingPageLayoutUndo=null;
  let lastPageBoardRenderSignature='',lastPageRangeAnchor='',activePackId='pack_test';
  const pageLayoutUndoStack=[],selectedPages=new Set();
  const api=(...args)=>hooks.api(...args);
  const showMessage=(...args)=>hooks.messages.push(args);
  const setPageSaveStatus=()=>{},updateBulkSelectionLabel=()=>{},renderPageOverview=()=>{},renderPages=()=>{};
  const applyPageMutationResult=()=>{},rememberPageMutationFingerprint=()=>{},undoLastPageLayout=()=>{};
  const handlePageLayoutConflict=async()=>{},userFriendlyError=message=>String(message||'');
  const activePackRecord=()=>({packId:activePackId}),activeLayoutFingerprint=()=>'',collectBoardVolumes=()=>({'ja-main':['p1']});
  const clonePageVolumes=volumes=>JSON.parse(JSON.stringify(volumes||{})),pageVolumeSnapshotsEqual=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
  ${extractFunction('saveBoardOrder')}
  ${extractFunction('awaitBoardSave')}
  ${extractFunction('waitForQueuedPageLayoutMutation')}
  ${extractFunction('sleep')}
  ${extractFunction('awaitPageMutationsIdle')}
  // scheduleBoardSave() と同じ順で旗を立ててから保存を始める。
  function startBoardSave(){boardSaveMutationBusy=true;pageMutationBusy=true;boardSaveRevision+=1;pendingBoardSaveVolumes={'ja-main':['p1']};return saveBoardOrder();}
  // 並べ替え以外のページ変更（移動・元に戻す・ページ名の保存）は旗だけを立てて送る。
  function setOtherMutation(busy,history=false){pageMutationBusy=busy;pageLayoutHistoryBusy=history&&busy;}
  return { startBoardSave, awaitBoardSave, waitForQueuedPageLayoutMutation, awaitPageMutationsIdle, setOtherMutation, state:()=>({boardSaveMutationBusy,pageMutationBusy}) };
`);

const hooks = { messages: [], api: null };
const queue = harness(hooks);

// 1. 保存に失敗した後は、前回の失敗を持ち越さずに進める。
hooks.api = async () => { throw new Error('network down'); };
assert.equal(await queue.startBoardSave(), false, 'a failed save must resolve to false');
assert.deepEqual(queue.state(), { boardSaveMutationBusy: false, pageMutationBusy: false }, 'a failed save must release the busy flags');
assert.equal(await queue.awaitBoardSave(), true, 'after a failed save has finished, callers must not keep seeing the old failure');
assert.equal(await queue.waitForQueuedPageLayoutMutation(), true, 'page mutations must be allowed again after a failed save has finished');

// 2. 保存中なら、その保存の結果を待つ（成功）。
let release;
hooks.api = () => new Promise(resolve => { release = () => resolve({ ok: true }); });
void queue.startBoardSave();
let settled = false;
const waiting = queue.awaitBoardSave().then(value => { settled = true; return value; });
await new Promise(resolve => setTimeout(resolve, 10));
assert.equal(settled, false, 'callers must wait while a board save is still in flight');
release();
assert.equal(await waiting, true, 'callers must see the in-flight save succeed');

// 3. 保存中に失敗したら、その失敗は待っている側へ届く（出力を止めるため）。
let reject = null;
hooks.api = () => new Promise((_, rejectSave) => { reject = () => rejectSave(new Error('conflict')); });
void queue.startBoardSave();
const waitingFailure = queue.awaitBoardSave();
while (!reject) await new Promise(resolve => setTimeout(resolve, 1));
reject();
assert.equal(await waitingFailure, false, 'a save that fails while being awaited must be reported as a failure');
assert.equal(await queue.awaitBoardSave(), true, 'and must not block later operations once finished');

// 4. 出力・共有コピーの前は、並べ替え以外のページ変更（移動・元に戻すなど）の完了も待つ（M-14）。
queue.setOtherMutation(true, true);
let idleSettled = false;
const idle = queue.awaitPageMutationsIdle().then(value => { idleSettled = true; return value; });
await new Promise(resolve => setTimeout(resolve, 120));
assert.equal(idleSettled, false, 'output must wait while a move or undo is still being saved');
queue.setOtherMutation(false);
assert.equal(await idle, true, 'output must proceed once the move or undo has finished');
// 待っている間に並べ替えの保存が失敗したら出力しない。
let rejectPending = null;
hooks.api = () => new Promise((_, rejectSave) => { rejectPending = () => rejectSave(new Error('conflict')); });
void queue.startBoardSave();
const idleAfterFailure = queue.awaitPageMutationsIdle();
while (!rejectPending) await new Promise(resolve => setTimeout(resolve, 1));
await new Promise(resolve => setTimeout(resolve, 10));
rejectPending();
assert.equal(await idleAfterFailure, false, 'output must not start when the pending board save fails');
// 終わらない変更は時間切れで止める（いつまでも押せないままにしない）。
queue.setOtherMutation(true);
assert.equal(await queue.awaitPageMutationsIdle(60), false, 'a mutation that never finishes must time out');
queue.setOtherMutation(false);

// 出力・共有コピーの入口がすべて待つこと。
for (const name of ['runFinalBuildJob', 'publishFinalVolume', 'publishAllFinalVolumes']) {
  assert.match(extractFunction(name), /await awaitPageMutationsIdle\(\)/, `${name} must wait for every pending page mutation`);
}

// 出力中は共有用フォルダーへコピーしない（M-3）。1冊ずつのボタンも、ジョブの開始・終了で描き直す。
for (const name of ['publishFinalVolume', 'publishAllFinalVolumes']) {
  assert.match(extractFunction(name), /if\(activeFinalJobId\)\{showMessage\(/, `${name} must refuse while an output job is running`);
}
assert.match(extractFunction('renderVolumeLinks'), /publish\.disabled=!publishable\|\|!!activeFinalJobId/, 'per-target copy buttons must be disabled while an output job is running');
const buildJob = extractFunction('runFinalBuildJob');
assert.ok(buildJob.indexOf('renderVolumeLinks()') > buildJob.indexOf('activeFinalJobId=jobId'), 'copy buttons must be redrawn when an output job starts');
assert.ok(buildJob.lastIndexOf('renderVolumeLinks()') > buildJob.indexOf("activeFinalJobId='';"), 'copy buttons must be redrawn when an output job ends');

// 待ち行列を直接 await する呼び出しを残さない。残すと同じ不具合が戻る。
const direct = source.match(/await\s+boardSavePromise\b/g) || [];
assert.equal(direct.length, 0, 'await awaitBoardSave() instead of the raw boardSavePromise queue');

// 位置を指定した移動（planPageMove）。純粋関数として取り出して確かめる。
const planPageMove = new Function(`${extractFunction('planPageMove')}\nreturn planPageMove;`)();
const lanesFixture = () => ({ 'ja-main': ['m1', 'm2', 'm3', 'm4', 'm5'], 'ja-supp': ['s1', 's2', 's3', 's4', 's5', 's6'], none: ['u1'] });
{
  // 補足6→本体3
  const input = lanesFixture();
  const frozen = JSON.stringify(input);
  const plan = planPageMove(input, ['s6'], 'ja-main', 3);
  assert.equal(plan.ok, true);
  assert.deepEqual(plan.volumes['ja-main'], ['m1', 'm2', 's6', 'm3', 'm4', 'm5'], 'supplement 6 must become main 3');
  assert.deepEqual(plan.volumes['ja-supp'], ['s1', 's2', 's3', 's4', 's5']);
  assert.deepEqual(plan.volumes.none, ['u1']);
  assert.equal(JSON.stringify(input), frozen, 'planPageMove must not mutate its input');
}
{
  // 範囲の移動は、選んだ順ではなく今の並び順のまま1かたまりで入れる。
  const plan = planPageMove(lanesFixture(), ['s4', 's2', 's3'], 'ja-main', 2);
  assert.deepEqual(plan.volumes['ja-main'], ['m1', 's2', 's3', 's4', 'm2', 'm3', 'm4', 'm5'], 'a moved range must keep its board order');
  assert.deepEqual(plan.movedIds, ['s2', 's3', 's4']);
  assert.deepEqual(plan.volumes['ja-supp'], ['s1', 's5', 's6']);
  // 出力先をまたいで選んでも、出力先の順→並び順でまとめる。
  const mixed = planPageMove(lanesFixture(), ['s1', 'm4'], 'none', 1);
  assert.deepEqual(mixed.volumes.none, ['m4', 's1', 'u1']);
}
{
  // 同じ出力先の中での移動。位置は動かすページを抜いた後の並びで数える。
  const forward = planPageMove(lanesFixture(), ['m1'], 'ja-main', 3);
  assert.deepEqual(forward.volumes['ja-main'], ['m2', 'm3', 'm1', 'm4', 'm5']);
  const backward = planPageMove(lanesFixture(), ['m5'], 'ja-main', 1);
  assert.deepEqual(backward.volumes['ja-main'], ['m5', 'm1', 'm2', 'm3', 'm4']);
  const inPlace = planPageMove(lanesFixture(), ['m2', 'm3'], 'ja-main', 2);
  assert.deepEqual(inPlace.volumes['ja-main'], ['m1', 'm2', 'm3', 'm4', 'm5'], 'moving a range onto itself must not change the order');
}
{
  // 末尾（省略・null・'end'・最後の次の番号）。今までの「移動先」ボタンと同じ結果になる。
  for (const position of [undefined, null, 'end', 6]) {
    const plan = planPageMove(lanesFixture(), ['s6', 's1'], 'ja-main', position);
    assert.deepEqual(plan.volumes['ja-main'], ['m1', 'm2', 'm3', 'm4', 'm5', 's1', 's6'], `position ${String(position)} must append`);
  }
  // 同じ出力先の中で末尾へ。抜いた後は4枚なので5番目が末尾。
  assert.deepEqual(planPageMove(lanesFixture(), ['m1'], 'ja-main', 5).volumes['ja-main'], ['m2', 'm3', 'm4', 'm5', 'm1']);
  // 空の出力先や、まだ並びに無い出力先へも入れられる。
  assert.deepEqual(planPageMove({ 'ja-main': ['m1'], 'ja-supp': [] }, ['m1'], 'ja-supp', 1).volumes, { 'ja-main': [], 'ja-supp': ['m1'] });
  assert.deepEqual(planPageMove({ 'ja-main': ['m1'] }, ['m1'], 'en-main').volumes, { 'ja-main': [], 'en-main': ['m1'] });
  // 並びに無いIDは指定順で後ろに付ける（今までの移動と同じく落とさない）。
  assert.deepEqual(planPageMove(lanesFixture(), ['x9', 's2'], 'none').movedIds, ['s2', 'x9']);
  // 文字列の番号も受け付ける。
  assert.deepEqual(planPageMove(lanesFixture(), ['s6'], 'ja-main', '3').volumes['ja-main'], ['m1', 'm2', 's6', 'm3', 'm4', 'm5']);
}
{
  // 範囲外の位置や移動先の無い指定は拒否し、並びを変えない。
  for (const position of [0, -1, 7, 2.5, NaN, '', 'abc', '3a', true]) {
    const input = lanesFixture();
    const plan = planPageMove(input, ['s6'], 'ja-main', position);
    assert.equal(plan.ok, false, `position ${String(position)} must be rejected`);
    assert.equal(plan.error, 'position-out-of-range');
    assert.equal(plan.maxPosition, 6);
    assert.deepEqual(plan.volumes, lanesFixture(), 'a rejected plan must keep the original order');
  }
  assert.equal(planPageMove(lanesFixture(), ['m1'], 'ja-main', 6).ok, false, 'within the same lane the end is one past the remaining pages');
  assert.equal(planPageMove(lanesFixture(), ['s6'], '', 1).error, 'target-missing');
  assert.equal(planPageMove(lanesFixture(), [], 'ja-main', 1).error, 'no-pages');
}
// 「移動先」ボタンとプレビューからの移動は、末尾に足す処理ではなく planPageMove を通す。
for (const name of ['movePageIdsToVolume', 'moveCurrentPreviewPageToVolume']) {
  const body = extractFunction(name);
  assert.match(body, /planPageMove\(/, `${name} must build the reorder through planPageMove`);
  assert.doesNotMatch(body, /\]\.push\((?:\.\.\.ids|pageId)\)/, `${name} must not append the moved pages by hand`);
}

console.log('page-board save queue regression ok');
