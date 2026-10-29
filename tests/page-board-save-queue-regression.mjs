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

// 待ち行列を直接 await する呼び出しを残さない。残すと同じ不具合が戻る。
const direct = source.match(/await\s+boardSavePromise\b/g) || [];
assert.equal(direct.length, 0, 'await awaitBoardSave() instead of the raw boardSavePromise queue');

console.log('page-board save queue regression ok');
