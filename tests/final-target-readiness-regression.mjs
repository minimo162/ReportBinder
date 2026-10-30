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
  isLatestPdfWorkbook: () => true,
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

// ---- 出力先カードとまとめて出力も pageCount ではなく itemCount で見る ----
const renderFinalBody = source.slice(source.indexOf('const renderFreshness=(ready,card)=>{'), source.indexOf('function renderFinalPreflight('));
assert.ok(renderFinalBody.includes('targetItemCount(ready)===0'), 'card empty state must use itemCount');
assert.ok(renderFinalBody.includes('btn.disabled=pending||empty||blockers.length>0'), 'card build button must use itemCount and wait for readiness');
assert.ok(renderFinalBody.includes("fix.classList.toggle('hidden',pending||empty||blockers.length===0)"), 'card fix button must use itemCount');
assert.ok(renderFinalBody.includes('!targetReadinessPending(ready)&&targetItemCount(ready)>0'), 'build-all must use itemCount');
assert.ok(!/Number\(ready\.pageCount\|\|0\)===0/.test(renderFinalBody), 'card must not judge emptiness by physical pageCount');
assert.ok(source.includes("entry.target?.required!==false||targetItemCount(entry.ready)>0"), 'optional targets must be shown by itemCount');

console.log('final target readiness regression ok');
