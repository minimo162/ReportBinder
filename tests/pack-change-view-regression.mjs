// 提出用PDF一式の変更点の画面（#pack-change-modal）の回帰テスト。
//
// - 比べる組み合わせを切り替えたとき、前の組み合わせの応答（成功・失敗とも）で画面を上書きしない
//   （openDiffDetail と同じ通し番号＋要求パスの照合）。閉じた後に届いた応答でも描かない。
// - 「くわしく見る」が行の sourceId・fromSnapshotId・toSnapshotId の順で openDiffDetail を呼ぶ。
// - 頁・左右はサーバーの値をそのまま出し、左右が記録されていなければ出さない（見開きの契約）。
// - ずれただけのページ・変わっていないページは最初は折りたたむ。左右が入れ替わったページは隠さない。
// - 出力先カードの入口: 組み込み一式・保存済み出力1つ・別の一式の一覧では出さない。
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
function extractConst(name) {
  const start = source.indexOf(`const ${name}=`);
  assert.ok(start >= 0, `const ${name} is missing from app.js`);
  const end = source.indexOf('};', start);
  return source.slice(start, end + 2);
}

function deferred() {
  let resolve, reject;
  const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}
const flush = () => new Promise(resolve => setImmediate(resolve));
// ブラウザでは window にある URLSearchParams・URL を vm の中でも使えるようにする。
const browserContext = (extra = {}) => vm.createContext({ URLSearchParams, URL, ...extra });

const stubs = `
function asArray(value){return value==null?[]:(Array.isArray(value)?value:[value]);}
function escapeHtml(s){return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function escapeAttr(s){return escapeHtml(s);}
function formatDateTime(v){return String(v||'');}
function userFriendlyError(m){return String(m||'');}
`;
const helpers = [
  'packChangeRequestPath', 'packChangeSideLabel', 'packChangePageLabel', 'packChangeKindLabels',
  'packChangeIsShiftOnly', 'packChangeIsQuiet', 'packChangeVisibleRows', 'packChangeRowHtml',
  'packChangeDefaultToBuildId', 'packChangeEntryState', 'packChangeSummaryDefinitions',
].map(extractFunction).join('\n') + '\n' + extractConst('PACK_CHANGE_KIND_LABELS');

// ---- 頁・左右はサーバーの値だけで出す -----------------------------------------
{
  const ctx = browserContext();
  vm.runInContext(stubs + helpers, ctx);
  assert.equal(ctx.packChangePageLabel(12, 'right'), 'p.12（右）');
  assert.equal(ctx.packChangePageLabel(11, 'left'), 'p.11（左）');
  // 古い形式の版は左右が空。頁番号の奇数・偶数から左右を推定して埋めない。
  assert.equal(ctx.packChangePageLabel(3, ''), 'p.3');
  assert.equal(ctx.packChangePageLabel(4, undefined), 'p.4');
  assert.equal(ctx.packChangePageLabel(0, 'right'), 'なし');
  assert.equal(ctx.packChangePageLabel(null, ''), 'なし');
  const html = ctx.packChangeRowHtml({ toPage: 5, toSide: '', fromPage: 4, fromSide: '', kinds: ['moved'], sourceName: 'a.xlsx', sheetName: 'S1' }, 0);
  assert.ok(!/（右）|（左）/.test(html), 'must not infer sides for pages without a recorded side');
  // 利用者が付けた原稿名・シート名はエスケープする。
  const escaped = ctx.packChangeRowHtml({ toPage: 1, toSide: 'right', kinds: ['content'], sourceName: '<script>x</script>', sheetName: 'S"1' }, 0);
  assert.ok(!escaped.includes('<script>'), 'source names must be escaped');
  // くわしく見るは detailAvailable のときだけ。
  assert.ok(!ctx.packChangeRowHtml({ toPage: 1, kinds: ['content'], sourceId: 'w1', detailAvailable: false }, 0).includes('data-pack-change-detail'));
  assert.ok(ctx.packChangeRowHtml({ toPage: 1, kinds: ['content'], sourceId: 'w1', detailAvailable: true }, 3).includes('data-pack-change-detail="3"'));

  // 要求パス: 省略時は既定の組み合わせ。指定はそのまま渡す。
  assert.equal(ctx.packChangeRequestPath('pack_1', 'main'), '/api/v2/outputs/changes?packId=pack_1&targetId=main');
  assert.equal(ctx.packChangeRequestPath('pack_1', 'main', 'b1', ''), '/api/v2/outputs/changes?packId=pack_1&targetId=main&fromBuildId=b1');
  assert.equal(ctx.packChangeRequestPath('pack_1', 'main', 'b1', 'b2'), '/api/v2/outputs/changes?packId=pack_1&targetId=main&fromBuildId=b1&toBuildId=b2');

  // 左右の丸ボタンは、両方の版に左右の記録があるときだけ出す（無いのに 0 と出さない）。
  const keys = changes => ctx.packChangeSummaryDefinitions(changes).map(item => item.key);
  assert.ok(keys({ hasSides: true, counts: { side: 0 } }).includes('side'));
  assert.ok(!keys({ hasSides: false, counts: { side: 0 } }).includes('side'));
}

// ---- ずれただけ・変わっていないページの折りたたみ ------------------------------
{
  const ctx = browserContext();
  vm.runInContext(stubs + helpers, ctx);
  const pages = [
    { toPage: 1, kinds: ['unchanged'] },
    { toPage: 2, kinds: ['content'] },
    { toPage: 3, kinds: ['added'] },
    { toPage: 4, kinds: ['moved'], shiftOnly: true },
    // ずれた結果、左右が入れ替わったページは見開きの確認で一番大事なので隠さない。
    { toPage: 5, kinds: ['moved', 'side'], shiftOnly: true },
    { toPage: 6, kinds: ['moved'], shiftOnly: false },
    { toPage: 7, kinds: ['same-look'] },
    { toPage: 8, kinds: ['moved', 'content'], shiftOnly: true },
  ];
  const pagesOf = rows => rows.map(row => row.toPage).join(',');
  assert.equal(pagesOf(ctx.packChangeVisibleRows(pages, 'all', false)), '2,3,5,6,8');
  assert.equal(pagesOf(ctx.packChangeVisibleRows(pages, 'all', true)), '1,2,3,4,5,6,7,8');
  // 「場所が変わった」で絞り込んだら、ずれただけのページも出す。
  assert.equal(pagesOf(ctx.packChangeVisibleRows(pages, 'moved', false)), '4,5,6,8');
  assert.equal(pagesOf(ctx.packChangeVisibleRows(pages, 'side', false)), '5');
  assert.deepEqual([...ctx.packChangeKindLabels(pages[3])], ['前のページが増減したためずれた']);
  assert.deepEqual([...ctx.packChangeKindLabels(pages[5])], ['場所が変わった']);
}

// ---- 出力先カードの入口 ---------------------------------------------------------
{
  const ctx = browserContext();
  vm.runInContext(stubs + helpers, ctx);
  const built = { outputPdf: 'C:\\out\\a.pdf', lastBuiltAt: '2026-11-02T09:40:00' };
  const archives = list => ({ packId: 'pack_1', loaded: true, list });
  const two = [{ volume: 'v-main', buildId: 'b2', isCurrent: true }, { volume: 'v-main', buildId: 'b1' }, { volume: 'v-app', buildId: 'c1' }];
  let entry = ctx.packChangeEntryState('v-main', built, false, archives(two), 'pack_1');
  assert.equal(entry.show, true);
  assert.equal(entry.label, '前回の出力からの変更点を見る');
  entry = ctx.packChangeEntryState('v-main', built, false, archives([{ ...two[0] }, { ...two[1], isLastPublished: true }]), 'pack_1');
  assert.equal(entry.label, '前回コピーしたときからの変更点を見る');
  // 保存済み出力が1つだけ: ボタンは出さず「はじめての出力」と出す。
  entry = ctx.packChangeEntryState('v-app', built, false, archives(two), 'pack_1');
  assert.equal(entry.show, false); assert.equal(entry.note, true);
  // 組み込み一式・未出力・別の一式の一覧・未取得では何も出さない。
  for (const args of [
    ['v-main', built, true, archives(two), 'pack_1'],
    ['v-main', {}, false, archives(two), 'pack_1'],
    ['v-main', built, false, archives(two), 'pack_2'],
    ['v-main', built, false, { packId: 'pack_1', loaded: false, list: two }, 'pack_1'],
  ]) {
    entry = ctx.packChangeEntryState(...args);
    assert.equal(entry.show, false); assert.equal(entry.note, false);
  }
  // 「今の版と比べる」の相手は isCurrent の版。無ければ最後に出力した版。
  assert.equal(ctx.packChangeDefaultToBuildId([{ volume: 'v', buildId: 'b3' }, { volume: 'v', buildId: 'b2', isCurrent: true }], 'v'), 'b2');
  assert.equal(ctx.packChangeDefaultToBuildId([{ volume: 'v', buildId: 'b3' }, { volume: 'v', buildId: 'b2' }], 'v'), 'b3');
  assert.equal(ctx.packChangeDefaultToBuildId([{ volume: 'w', buildId: 'b3' }], 'v'), '');
}

// ---- 古い応答で新しい比較を上書きしない ------------------------------------------
function makeLoadContext() {
  const elements = new Map();
  const el = id => {
    if (!elements.has(id)) {
      const classes = new Set();
      elements.set(id, { id, innerHTML: '', textContent: '', value: '', checked: false, open: false,
        classList: { add: c => classes.add(c), remove: c => classes.delete(c), contains: c => classes.has(c), toggle: (c, on) => { if (on === undefined ? !classes.has(c) : on) classes.add(c); else classes.delete(c); } } });
    }
    return elements.get(id);
  };
  const ctx = browserContext({ calls: [], rendered: [], errors: [], pending: [] });
  vm.runInContext(stubs + helpers + `
var activePackId='pack_1';
var finalArchivesState={packId:'pack_1',loaded:true,list:[{volume:'v-main',buildId:'b1'},{volume:'v-main',buildId:'b2'}]};
var packChangeState={volume:'v-main',targetId:'main',fromBuildId:'',toBuildId:'',requestPath:'',changes:null,archives:[],filter:'all',showAll:false};
var packChangeOpenSerial=0;
var packChangeReturnFocus=null;
function $(id){return __el(id);}
function api(path){const d=__deferred();pending.push({path,d});return d.promise;}
function renderPackChanges(changes,archives){rendered.push(changes);}
function renderPackChangeError(message){errors.push(message);}
` + ['isPackChangeOpen', 'loadPackChangeArchives', 'loadPackChanges', 'comparePackChangeVersions', 'closePackChanges', 'openPackChangeDetail'].map(extractFunction).join('\n'), Object.assign(ctx, { __el: el, __deferred: deferred }));
  return { ctx, el };
}
{
  const { ctx } = makeLoadContext();
  // 既定の組み合わせを開き、届く前に「ほかの版と比べる」で b1→b2 を選ぶ。
  const first = ctx.loadPackChanges({});
  ctx.__el('pack-change-from').value = 'b1';
  ctx.__el('pack-change-to').value = 'b2';
  const second = ctx.comparePackChangeVersions();
  assert.equal(ctx.pending.length, 2);
  assert.match(ctx.pending[1].path, /fromBuildId=b1&toBuildId=b2/);
  // 新しい方が先に届き、古い方が後から届く。
  ctx.pending[1].d.resolve({ changes: { status: 'ready', tag: 'new' } });
  await second;
  ctx.pending[0].d.resolve({ changes: { status: 'ready', tag: 'old' } });
  await first;
  assert.deepEqual(ctx.rendered.map(c => c.tag), ['new'], 'a stale response must not overwrite the newer comparison');

  // 古い要求の失敗も描かない。
  const third = ctx.loadPackChanges({ fromBuildId: 'b1' });
  const fourth = ctx.loadPackChanges({});
  ctx.pending[3].d.resolve({ changes: { status: 'ready', tag: 'latest' } });
  await fourth;
  ctx.pending[2].d.reject(new Error('boom'));
  await third;
  assert.deepEqual(ctx.errors, [], 'a stale failure must not overwrite the newer comparison');
  assert.deepEqual(ctx.rendered.map(c => c.tag), ['new', 'latest']);

  // 閉じた後に届いた応答も描かない。
  const fifth = ctx.loadPackChanges({});
  ctx.closePackChanges();
  ctx.pending[4].d.resolve({ changes: { status: 'ready', tag: 'after-close' } });
  await fifth;
  await flush();
  assert.deepEqual(ctx.rendered.map(c => c.tag), ['new', 'latest']);

  // 同じ版どうしは要求しない。
  ctx.__el('pack-change-from').value = 'b2';
  ctx.__el('pack-change-to').value = 'b2';
  const before = ctx.pending.length;
  ctx.comparePackChangeVersions();
  assert.equal(ctx.pending.length, before);
  assert.match(ctx.__el('pack-change-other-hint').textContent, /同じ版/);
}

// ---- 「くわしく見る」は正しい版で原稿単位の比較を開く ----------------------------
{
  const { ctx } = makeLoadContext();
  const seen = [];
  ctx.openDiffDetail = (...args) => { seen.push(args); };
  const opener = { id: 'btn' };
  ctx.openPackChangeDetail({ sourceId: 'wb_1', fromSnapshotId: 'snap_old', toSnapshotId: 'snap_new', detailAvailable: true }, opener);
  assert.equal(seen.length, 1);
  assert.equal(seen[0][0], 'wb_1');
  assert.equal(seen[0][1], opener);
  assert.equal(seen[0][2].fromSnapshotId, 'snap_old', 'from must be the earlier version');
  assert.equal(seen[0][2].toSnapshotId, 'snap_new', 'to must be the later version');
  ctx.openPackChangeDetail({ sourceId: 'wb_1', fromSnapshotId: 'a', toSnapshotId: 'b', detailAvailable: false }, opener);
  ctx.openPackChangeDetail({ sourceId: '', fromSnapshotId: 'a', toSnapshotId: 'b', detailAvailable: true }, opener);
  assert.equal(seen.length, 1);
}

console.log('pack change view regression ok');
