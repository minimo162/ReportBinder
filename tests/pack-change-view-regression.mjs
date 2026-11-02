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

// ---- 変更点PDF（確認用） -------------------------------------------------------------
const reportHelpers = ['packChangeReportBody', 'packChangeReportView', 'packChangeReportCopyNote'].map(extractFunction).join('\n') + '\n'
  + source.slice(source.indexOf('const PACK_CHANGE_REPORT_TERMINAL='), source.indexOf('];', source.indexOf('const PACK_CHANGE_REPORT_TERMINAL=')) + 2);
const readyChanges = (extra = {}) => ({ status: 'ready', from: { buildId: 'b1' }, to: { buildId: 'b3' }, report: { exists: false, reason: 'not-created' }, ...extra });
{
  const ctx = browserContext();
  vm.runInContext(stubs + reportHelpers, ctx);
  // 作るのは画面に出ている2つの版。既定の組み合わせでも buildId を両方送る（作成までに出力・コピーがあっても組み合わせが変わらない）。
  assert.deepEqual({ ...ctx.packChangeReportBody(readyChanges(), 'pack_1', 'main') }, { packId: 'pack_1', targetId: 'main', fromBuildId: 'b1', toBuildId: 'b3' });
  assert.equal(ctx.packChangeReportBody({ status: 'no-baseline', to: { buildId: 'b3' } }, 'pack_1', 'main'), null);
  assert.equal(ctx.packChangeReportBody(readyChanges({ from: null }), 'pack_1', 'main'), null);

  // 比べられないときは欄ごと出さない。
  assert.equal(ctx.packChangeReportView({ status: 'unsupported' }, null).show, false);
  let view = ctx.packChangeReportView(readyChanges(), null);
  assert.equal(view.show, true); assert.equal(view.canCreate, true); assert.equal(view.canOpen, false);
  assert.equal(view.createLabel, '変更点をPDFにする');
  // 作成中は押せない。進み具合を出す。
  view = ctx.packChangeReportView(readyChanges(), { status: 'running', percent: 40, message: '印刷しています。' });
  assert.equal(view.busy, true); assert.equal(view.percent, 40); assert.match(view.status, /40%/);
  // 終わったら「開く」「フォルダーを開く」。
  view = ctx.packChangeReportView(readyChanges(), { status: 'completed' });
  assert.equal(view.canOpen, true); assert.equal(view.canCreate, false); assert.match(view.status, /変更点PDFを作りました/);
  // Edge で作れなかったときは、サーバーの説明をそのまま出し、もう一度押せる。
  view = ctx.packChangeReportView(readyChanges(), { status: 'failed', message: '管理者に『Edge のヘッドレス印刷』が使えるか確認してください。' });
  assert.equal(view.canCreate, true); assert.equal(view.canOpen, false); assert.match(view.status, /ヘッドレス印刷/);
  // 作成済み・古くなったもの。
  view = ctx.packChangeReportView(readyChanges({ report: { exists: true, createdAt: '2026-11-02' } }), null);
  assert.equal(view.canOpen, true); assert.equal(view.canCreate, false);
  view = ctx.packChangeReportView(readyChanges({ report: { exists: false, reason: 'outdated' } }), null);
  assert.equal(view.canOpen, false); assert.equal(view.createLabel, '変更点PDFを作り直す');
}

function makeReportContext() {
  const elements = new Map();
  const el = id => {
    if (!elements.has(id)) {
      const classes = new Set();
      elements.set(id, { id, innerHTML: '', textContent: '', value: '', disabled: false, checked: false, open: false,
        classList: { add: c => classes.add(c), remove: c => classes.delete(c), contains: c => classes.has(c), toggle: (c, on) => { if (on === undefined ? !classes.has(c) : on) classes.add(c); else classes.delete(c); } } });
    }
    return elements.get(id);
  };
  const ctx = browserContext({ calls: [], pending: [], renders: [] });
  vm.runInContext(stubs + helpers + reportHelpers + `
var activePackId='pack_1';
var finalArchivesState={packId:'pack_1',loaded:true,list:[]};
var packChangeState={volume:'v-main',targetId:'main',fromBuildId:'',toBuildId:'',requestPath:'',changes:null,archives:[],filter:'all',showAll:false};
var packChangeOpenSerial=0;
var packChangeReturnFocus=null;
var packChangeReportJob=null;
function $(id){return __el(id);}
function api(path,options){const d=__deferred();pending.push({path,options,d});return d.promise;}
function sleep(){return Promise.resolve();}
function renderPackChanges(changes){packChangeState.changes=changes;}
function renderPackChangeError(){}
function renderPackChangeReport(){renders.push(packChangeReportJob?{...packChangeReportJob}:null);}
` + ['isPackChangeOpen', 'loadPackChangeArchives', 'loadPackChanges', 'closePackChanges', 'createPackChangeReport'].map(extractFunction).join('\n'), Object.assign(ctx, { __el: el, __deferred: deferred }));
  return ctx;
}
{
  const ctx = makeReportContext();
  // 既定の組み合わせを開いた状態（b1 → b3）。
  const load = ctx.loadPackChanges({});
  ctx.pending.shift().d.resolve({ changes: readyChanges() });
  await load;
  assert.equal(ctx.packChangeState.changes.to.buildId, 'b3');

  // 作成 → 状態を見る → 完了。送るのは画面の2つの版。
  const creating = ctx.createPackChangeReport();
  const start = ctx.pending.shift();
  assert.equal(start.path, '/api/v2/outputs/changes/report');
  assert.deepEqual({ ...start.options.body }, { packId: 'pack_1', targetId: 'main', fromBuildId: 'b1', toBuildId: 'b3' });
  // 作成中にもう一度押しても、2本目は始めない。
  await ctx.createPackChangeReport();
  assert.equal(ctx.pending.length, 0);
  start.d.resolve({ job: { jobId: 'job_1', status: 'launching', percent: 2 } });
  await flush(); await flush();
  const poll = ctx.pending.shift();
  assert.equal(poll.path, '/api/v2/outputs/changes/report/status?jobId=job_1');
  poll.d.resolve({ job: { jobId: 'job_1', status: 'completed', percent: 100, report: { exists: true, reportId: 'b1__b3' } } });
  await creating;
  assert.equal(ctx.packChangeReportJob.status, 'completed');
  assert.equal(ctx.packChangeState.changes.report.exists, true, 'a finished report must be offered for opening');
}
{
  // 作成中に別の組み合わせへ切り替えたら、前の組み合わせのジョブの結果で画面を書き換えない。
  const ctx = makeReportContext();
  const load = ctx.loadPackChanges({});
  ctx.pending.shift().d.resolve({ changes: readyChanges() });
  await load;
  const creating = ctx.createPackChangeReport();
  ctx.pending.shift().d.resolve({ job: { jobId: 'job_1', status: 'running', percent: 30 } });
  await flush(); await flush();
  const poll = ctx.pending.shift();
  const switched = ctx.loadPackChanges({ fromBuildId: 'b2', toBuildId: 'b3' });
  assert.equal(ctx.packChangeReportJob, null, 'switching the pair must drop the previous report job');
  ctx.pending.shift().d.resolve({ changes: readyChanges({ from: { buildId: 'b2' } }) });
  await switched;
  const rendersBefore = ctx.renders.length;
  poll.d.resolve({ job: { jobId: 'job_1', status: 'completed', report: { exists: true } } });
  await creating;
  assert.equal(ctx.renders.length, rendersBefore, 'a stale report job must not repaint the new comparison');
  assert.equal(ctx.packChangeReportJob, null);
  assert.equal(ctx.packChangeState.changes.report.exists, false, 'a report for the old pair must not be offered for the new pair');
  assert.equal(ctx.pending.length, 0, 'polling must stop after the pair changed');
}
{
  // 閉じた後の失敗も描かない。開いているときの失敗はサーバーの説明を出す。
  const ctx = makeReportContext();
  const load = ctx.loadPackChanges({});
  ctx.pending.shift().d.resolve({ changes: readyChanges() });
  await load;
  let creating = ctx.createPackChangeReport();
  ctx.pending.shift().d.resolve({ job: { jobId: 'job_2', status: 'running' } });
  await flush(); await flush();
  ctx.pending.shift().d.resolve({ job: { jobId: 'job_2', status: 'failed', message: 'この PC では変更点PDFを作れませんでした。' } });
  await creating;
  assert.equal(ctx.packChangeReportJob.status, 'failed');
  assert.match(ctx.packChangeReportJob.message, /作れませんでした/);
  creating = ctx.createPackChangeReport();
  ctx.closePackChanges();
  const rendersBefore = ctx.renders.length;
  ctx.pending.shift().d.reject(new Error('boom'));
  await creating;
  assert.equal(ctx.renders.length, rendersBefore, 'a failure after closing must not repaint');
}

// ---- 共有用フォルダーへのコピーに変更点PDFを同梱する ------------------------------------
const publishHelpers = ['publishChangeReportPlan', 'publishChangeReportNote'].map(extractFunction).join('\n');
{
  const ctx = browserContext();
  vm.runInContext(stubs + reportHelpers + publishHelpers, ctx);
  const item = (targetId, report, status = 'ready') => ({ volume: `v-${targetId}`, targetId, label: targetId, changes: { status, from: { buildId: 'b1' }, to: { buildId: 'b2' }, report } });
  // 付けられるのは attachable（前回コピーした版 → 今の提出用PDF）だけ。作成済みなら ready、未作成なら create。
  const plan = ctx.publishChangeReportPlan([
    item('main', { attachable: true, publishable: true, exists: true }),
    item('appendix', { attachable: true, publishable: false, exists: false }),
    item('picked', { attachable: false, publishable: false, exists: true }),
    item('first', { attachable: true, publishable: true }, 'no-baseline'),
  ]);
  assert.equal(plan.offer, true);
  assert.deepEqual([...plan.ready.map(i => i.targetId)], ['main']);
  assert.deepEqual([...plan.create.map(i => i.targetId)], ['appendix'], 'a report for another pair must never be offered for the copy');
  assert.equal(ctx.publishChangeReportPlan([item('picked', { attachable: false, exists: true })]).offer, false);
  // 欄の説明: 同梱できる／作り直すとできる／この組み合わせはしない。
  assert.match(ctx.packChangeReportView(readyChanges({ report: { exists: true, attachable: true, publishable: true } }), null).status, /いっしょにコピーできます/);
  assert.match(ctx.packChangeReportView(readyChanges({ report: { exists: true, attachable: true, publishable: false } }), null).status, /作り直すとコピーできます/);
  assert.match(ctx.packChangeReportView(readyChanges({ report: { exists: true, attachable: false } }), null).status, /共有用フォルダーにはコピーしません/);
  // 通知: 付けたものを書き、はじめての出力先など当たり前の理由は書かない。
  const note = ctx.publishChangeReportNote({ changeReports: [{ fileName: 'a_変更点.pdf' }], changeReportsSkipped: [{ targetName: '補足', reason: 'not-published', message: 'まだ' }, { targetName: '本体2', reason: 'outdated', message: '作り直すとコピーできます。' }] }, true);
  assert.match(note, /a_変更点\.pdf/); assert.match(note, /本体2の変更点PDFは付けていません/); assert.ok(!note.includes('補足'));
  assert.equal(ctx.publishChangeReportNote({ changeReports: [{ fileName: 'a_変更点.pdf' }] }, false), '');
}

function makePublishContext({ changes, confirm = true, checked = true, reportJob = { status: 'completed' } }) {
  const elements = new Map();
  const el = id => {
    if (!elements.has(id)) {
      const classes = new Set(['hidden']);
      elements.set(id, { id, innerHTML: '', textContent: '', value: '', disabled: false, checked: false, className: '', onchange: null,
        classList: { add: c => classes.add(c), remove: c => classes.delete(c), contains: c => classes.has(c), toggle: (c, on) => { if (on === undefined ? !classes.has(c) : on) classes.add(c); else classes.delete(c); } } });
    }
    return elements.get(id);
  };
  const ctx = browserContext({ calls: [], messages: [], confirms: [] });
  vm.runInContext(stubs + helpers + reportHelpers + publishHelpers + `
var activePackId='pack_1',activePreset='p',activeFinalJobId='';
var confirmResolver=null,confirmReturnFocus=null;
var document={activeElement:null};
function requestAnimationFrame(){}
function $(id){return __el(id);}
function isConfirmModalOpen(){return false;}
function closeConfirmModal(){}
function sleep(){return Promise.resolve();}
function volumeLabel(v){return v==='v-main'?'本体':v;}
function targetIdFromVolume(v){return String(v).replace(/^v-/,'');}
function activePackIsBuiltIn(){return false;}
function awaitPageMutationsIdle(){return Promise.resolve(true);}
function showMessage(type,title,message){messages.push({type,title,message});}
async function runBusy(btn,fn){await fn();}
async function api(path,options){
  calls.push({path,body:options?.body||null});
  if(path.startsWith('/api/v2/outputs/changes?'))return {changes:__changes};
  if(path==='/api/v2/outputs/changes/report')return {job:{jobId:'job_1',status:'running'}};
  if(path.startsWith('/api/v2/outputs/changes/report/status'))return {job:__reportJob};
  if(path==='/api/v2/outputs/publish')return {result:{sharedPath:'S:\\\\x\\\\a.pdf',changeReports:options.body.includeChangeReports?[{fileName:'a_変更点.pdf'}]:[]}};
  throw new Error('unexpected '+path);
}
` + ['confirmAction', 'confirmCheckState', 'loadPublishChangeReportPlan', 'confirmPublishChangeReports', 'runPackChangeReportJob', 'createPublishChangeReports', 'preparePublishChangeReports', 'publishFinalVolume'].map(extractFunction).join('\n'), Object.assign(ctx, { __el: el, __changes: changes, __reportJob: reportJob }));
  // 確認画面の「コピーする」「キャンセル」とチェックを真似る。
  ctx.closeConfirmModal = () => {};
  const realConfirm = ctx.confirmAction;
  ctx.confirmAction = opts => {
    const p = realConfirm(opts);
    ctx.confirms.push({ opts, accept: el('confirm-accept').textContent, wrapHidden: el('confirm-check-wrap').classList.contains('hidden'), checked: el('confirm-check').checked });
    if (!checked) { el('confirm-check').checked = false; el('confirm-check').onchange?.(); ctx.confirms.at(-1).acceptAfterUncheck = el('confirm-accept').textContent; }
    ctx.confirmResolver(confirm);
    return p;
  };
  return ctx;
}
const attachChanges = report => ({ status: 'ready', from: { buildId: 'b1' }, to: { buildId: 'b2' }, report });
{
  // 作成済み: チェックありの確認を出し、そのままコピーすると includeChangeReports:true。
  const ctx = makePublishContext({ changes: attachChanges({ attachable: true, publishable: true, exists: true }) });
  await ctx.publishFinalVolume('v-main', null);
  assert.equal(ctx.confirms.length, 1);
  assert.equal(ctx.confirms[0].wrapHidden, false); assert.equal(ctx.confirms[0].checked, true, 'the copy-together box must start checked');
  assert.equal(ctx.confirms[0].accept, 'コピーする');
  const publish = ctx.calls.find(c => c.path === '/api/v2/outputs/publish');
  assert.equal(publish.body.includeChangeReports, true);
  assert.equal(ctx.calls.filter(c => c.path === '/api/v2/outputs/changes/report').length, 0, 'a made report must not be made again');
  assert.match(ctx.messages.at(-1).message, /a_変更点\.pdf/);
}
{
  // 未作成: ボタンは「変更点PDFを作ってからコピーする」。先に作ってからコピーする。
  const ctx = makePublishContext({ changes: attachChanges({ attachable: true, publishable: false, exists: false }) });
  await ctx.publishFinalVolume('v-main', null);
  assert.equal(ctx.confirms[0].accept, '変更点PDFを作ってからコピーする');
  const order = ctx.calls.map(c => c.path.split('?')[0]);
  const created = order.indexOf('/api/v2/outputs/changes/report'), published = order.indexOf('/api/v2/outputs/publish');
  assert.ok(created >= 0 && published > created, `the report must be made before the copy: ${order.join(', ')}`);
  assert.deepEqual({ ...ctx.calls[created].body }, { packId: 'pack_1', targetId: 'main', fromBuildId: 'b1', toBuildId: 'b2' });
  assert.equal(ctx.calls[published].body.includeChangeReports, true);
}
{
  // 作れなかったら、何もコピーしない。
  const ctx = makePublishContext({ changes: attachChanges({ attachable: true, publishable: false }), reportJob: { status: 'failed', message: 'この PC では変更点PDFを作れませんでした。' } });
  await ctx.publishFinalVolume('v-main', null);
  assert.equal(ctx.calls.filter(c => c.path === '/api/v2/outputs/publish').length, 0, 'a failed report must stop the copy');
  assert.match(ctx.messages.at(-1).title, /コピーしていません/);
}
{
  // チェックを外すと、作らずに提出用PDFだけ。ボタンの文言も戻る。
  const ctx = makePublishContext({ changes: attachChanges({ attachable: true, publishable: false }), checked: false });
  await ctx.publishFinalVolume('v-main', null);
  assert.equal(ctx.confirms[0].acceptAfterUncheck, 'コピーする');
  assert.equal(ctx.calls.filter(c => c.path === '/api/v2/outputs/changes/report').length, 0);
  assert.equal(ctx.calls.find(c => c.path === '/api/v2/outputs/publish').body.includeChangeReports, false);
}
{
  // キャンセルなら何もしない。付けられるものが無ければ、今までどおり確認なしでコピーする。
  let ctx = makePublishContext({ changes: attachChanges({ attachable: true, publishable: true, exists: true }), confirm: false });
  await ctx.publishFinalVolume('v-main', null);
  assert.equal(ctx.calls.filter(c => c.path === '/api/v2/outputs/publish').length, 0);
  ctx = makePublishContext({ changes: attachChanges({ attachable: false, exists: true }) });
  await ctx.publishFinalVolume('v-main', null);
  assert.equal(ctx.confirms.length, 0);
  assert.equal(ctx.calls.find(c => c.path === '/api/v2/outputs/publish').body.includeChangeReports, false);
  // 確認の部品: check を渡さない確認ではチェック欄を出さず、状態も false。
  ctx.confirms.length = 0;
  void ctx.confirmAction({ title: 'x' });
  assert.equal(ctx.confirms[0].wrapHidden, true);
  assert.equal(ctx.confirmCheckState(), false);
}

console.log('pack change view regression ok');
