// 横パネル「PDFイメージを横に表示」（操作しながら見る）の回帰テスト。
//
// サムネイル・詳細の盤面の右に、表示中の出力先の物理ページを小さく並べる。並びは盤面のDOM
// （collectBoardVolumes）から読むので、ドラッグ・キー操作の直後、保存（と readiness の読み直し）を
// 待たずに新しい順で出る。409 の競合で盤面が状態へ戻れば、パネルも戻る。
// パネルは #page-board の外に置き、レーンの属性もカードのクラスも持たない（並びの保存に混ざらない）。
//
// app.js の実物の関数を取り出し、偽の DOM・rAF で動かして確かめる。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const appPath = fileURLToPath(new URL('../app/web/app.js', import.meta.url));
const source = readFileSync(appPath, 'utf8');

function extractFunction(name) {
  const start = source.indexOf(`function ${name}(`);
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

// 盤面のDOM（collectBoardVolumes の戻り値）と状態（readiness）を別々に持つ。
const env = {
  board: { none: ['U'], 'v-main': ['A', 'B', 'C'], 'v-sub': ['S'] },
  pages: [],
  targets: [{ targetId: 'main', displayName: '本体' }, { targetId: 'sub', displayName: '補足' }],
  readiness: {},
  activePageVolume: 'all',
  selectedPages: new Set(),
  pageBoardView: 'thumbnail',
  wide: true,
  enabled: true,
  frames: [],
  observed: 0,
  dequeued: 0,
};
const workbooks = new Map([['wb', { workbookId: 'wb', latest: true }], ['wb-old', { workbookId: 'wb-old', latest: false }]]);
const pane = {
  innerHTML: '', scrollTop: 0, hidden: false,
  get childElementCount() { return this.innerHTML ? 1 : 0; },
  // 選択の強調・描画待ちの取り消しは、テストでは中身を見ない。
  querySelectorAll() { return []; },
};

const api = new Function('env', 'workbooks', 'pane', `
  const window={matchMedia:()=>({matches:env.wide})};
  const requestAnimationFrame=fn=>{env.frames.push(fn);return env.frames.length;};
  const cancelAnimationFrame=()=>{};
  const sessionStorage={setItem(){}};
  const $=id=>id==='page-pdf-image-pane'?pane:null;
  const asArray=value=>Array.isArray(value)?value:[];
  const escapeHtml=value=>String(value??'');
  const escapeAttr=value=>String(value??'');
  const resolvedPageId=page=>String(page?.pageId||'');
  const getWorkbook=id=>workbooks.get(String(id||''))||null;
  const isLatestPdfWorkbook=w=>!!w?.latest;
  const pagePreviewAvailable=page=>!!page?.contentPdf&&isLatestPdfWorkbook(getWorkbook(page.workbookId));
  const pageThumbnailCacheKey=(page,n)=>resolvedPageId(page)+'|'+String(page?.contentPdf||'')+'|p'+n;
  const pdfImagePaperRatios=new Map();
  const pdfImageObservedPageCounts=new Map();
  const PDF_IMAGE_PANE_MIN_WIDTH=1100;
  const activeTargetDefinitions=()=>env.targets;
  const targetVolume=id=>'v-'+id;
  const activeTargetVolumes=()=>env.targets.map(t=>targetVolume(t.targetId));
  const volumeReadiness=volume=>env.readiness[volume]||null;
  const pagesForActivePreset=()=>env.pages;
  const collectBoardVolumes=()=>JSON.parse(JSON.stringify(env.board));
  const dequeuePageThumbnail=()=>{env.dequeued++;};
  const observePdfImagePaneThumbnails=()=>{env.observed++;pdfImagePaneObserveDirty=false;};
  const state={};
  const activePageVolume=env.activePageVolume;
  const selectedPages=env.selectedPages;
  let pdfImagePaneFrame=0,pdfImagePaneVolumes=null,pdfImagePaneSignature='',pdfImagePaneObserveDirty=false,pdfImagePaneLastVolume='';
  ${extractFunction('clonePageVolumes')}
  ${extractFunction('pageFirstSourcePage')}
  ${extractFunction('pageLastSourcePage')}
  ${extractFunction('collectPdfImagePageCounts')}
  ${extractFunction('buildPdfImageLayout')}
  ${extractFunction('pdfImagePaneWideEnough')}
  ${extractFunction('schedulePdfImagePaneRefresh')}
  ${extractFunction('pdfImagePaneTargetVolume')}
  ${extractFunction('pdfImagePaneSheetHtml')}
  ${extractFunction('renderPdfImagePane')}
  function isReadOnlyBoardView(view=env.pageBoardView){return view==='spread'||view==='pdf';}
  function pdfImagePaneVisible(){return env.enabled&&!isReadOnlyBoardView()&&pdfImagePaneWideEnough();}
  return {
    schedulePdfImagePaneRefresh, pdfImagePaneTargetVolume,
    setSelected(ids){selectedPages.clear();ids.forEach(id=>selectedPages.add(id));},
  };
`)(env, workbooks, pane);

const flush = () => { const frames = env.frames.splice(0); frames.forEach(fn => fn()); return frames.length; };
const paneOrder = () => [...pane.innerHTML.matchAll(/data-pane-page-id="([^"]+)"[^>]*data-source-page="(\d+)"/g)].map(m => `${m[1]}:${m[2]}`);
const paneHead = () => (pane.innerHTML.match(/<div class="pdf-pane-head">(.*?)<small>/) || [])[1] || '';
const parts = (pageId, start, count) => Array.from({ length: count }, (_, i) => ({ pageId, pagePartIndex: i + 1, pagePartCount: count, sourcePdfPageNumber: start + i }));
const page = (pageId, extra = {}) => ({ pageId, workbookId: 'wb', contentPdf: `content-pdf/${pageId}.pdf`, ...extra });

env.pages = [page('U'), page('A'), page('B', { pageRange: '2-4' }), page('C'), page('S')];
env.readiness = {
  'v-main': { snapshot: { physicalPages: [...parts('A', 1, 1), ...parts('B', 2, 3), ...parts('C', 1, 1)] } },
  'v-sub': { snapshot: { physicalPages: parts('S', 1, 1) } },
};

// 表示する出力先: タブの出力先 →「全体」では最後に選んだページの出力先 → 前回の出力先 → 最初の出力先。
{
  const volumes = env.board, targets = ['v-main', 'v-sub'];
  assert.equal(api.pdfImagePaneTargetVolume(volumes, targets, 'v-sub', new Set(), ''), 'v-sub', 'a focused target tab wins');
  assert.equal(api.pdfImagePaneTargetVolume(volumes, targets, 'all', new Set(['A', 'S']), ''), 'v-sub', 'the last selected page decides in the overview');
  assert.equal(api.pdfImagePaneTargetVolume(volumes, targets, 'all', new Set(['S', 'A']), ''), 'v-main');
  assert.equal(api.pdfImagePaneTargetVolume(volumes, targets, 'none', new Set(['U']), 'v-sub'), 'v-sub', 'unassigned pages fall back to the previous target');
  assert.equal(api.pdfImagePaneTargetVolume(volumes, targets, 'all', new Set(), 'gone'), 'v-main', 'otherwise the first target');
  assert.equal(api.pdfImagePaneTargetVolume(volumes, [], 'all', new Set(), ''), '', 'no targets, no pane content');
}

// 初回: 盤面の並びどおり、複数ページのシートは使用ページ範囲の全ページを出す。
api.schedulePdfImagePaneRefresh();
assert.equal(flush(), 1);
assert.deepEqual(paneOrder(), ['A:1', 'B:2', 'B:3', 'B:4', 'C:1']);
assert.match(paneHead(), /本体（提出用PDF）.*<span>5 ページ<\/span>/);
assert.equal(env.observed, 1, 'new pane sheets are observed by the pane observer');
for (const banned of ['data-volume', 'page-row', 'data-page-id=']) assert.ok(!pane.innerHTML.includes(banned), `pane markup must not look like a board lane/card: ${banned}`);

// rAF で1回にまとめる。並びが変わらなければ作り直さない（描いた canvas を捨てない）。
api.schedulePdfImagePaneRefresh();
api.schedulePdfImagePaneRefresh();
const before = pane.innerHTML;
assert.equal(flush(), 1, 'refreshes within a frame are batched');
assert.equal(pane.innerHTML, before);
assert.equal(env.observed, 1, 'an unchanged pane keeps its canvases and observer');

// ドラッグで盤面のDOMだけが変わった直後（保存前・readiness は古いまま）でも、新しい順で出る。
const stateOrder = JSON.parse(JSON.stringify(env.board));
env.board['v-main'] = ['C', 'A', 'B'];
api.schedulePdfImagePaneRefresh(env.board); // scheduleBoardSave と同じ呼び方
flush();
assert.deepEqual(paneOrder(), ['C:1', 'A:1', 'B:2', 'B:3', 'B:4'], 'the pane follows the DOM order before the save');
assert.match(paneHead(), /<span>5 ページ<\/span>/, 'page counts do not change by reordering, so the count stays exact');

// 409 の競合などで盤面が状態へ戻ったら、パネルも戻る（renderPages の末尾からの呼び出し）。
env.board = stateOrder;
api.schedulePdfImagePaneRefresh();
flush();
assert.deepEqual(paneOrder(), ['A:1', 'B:2', 'B:3', 'B:4', 'C:1'], 'the pane reverts with the board');

// 未振り分けから移したばかりで頁数の分からないページは、1枚分の場所で「約」。
env.board = { none: [], 'v-main': ['A', 'U', 'B', 'C'], 'v-sub': ['S'] };
api.schedulePdfImagePaneRefresh(env.board);
flush();
assert.deepEqual(paneOrder(), ['A:1', 'U:1', 'B:2', 'B:3', 'B:4', 'C:1']);
assert.match(paneHead(), /<span>約 6 ページ<\/span>/);
assert.match(pane.innerHTML, /頁数確認中/);

// 古い変換PDFは allow-stale で描き、印を付ける。PDF未作成は canvas を作らない。
env.pages = env.pages.map(p => (p.pageId === 'A' ? { ...p, workbookId: 'wb-old' } : p.pageId === 'C' ? { ...p, contentPdf: '' } : p));
api.schedulePdfImagePaneRefresh();
flush();
assert.match(pane.innerHTML, /data-pane-page-id="A"[^>]*>[^]*?data-thumbnail-allow-stale="1"[^]*?変換PDFが古い/);
assert.ok(!/data-page-thumbnail="C"/.test(pane.innerHTML), 'a page without a converted PDF has no canvas');
assert.match(pane.innerHTML, /PDF未作成/);

// 「全体」で補足のページを選ぶと、補足のPDFを出す。選択を外しても補足のまま（前回の出力先）。
api.setSelected(['S']);
api.schedulePdfImagePaneRefresh();
flush();
assert.deepEqual(paneOrder(), ['S:1']);
assert.match(paneHead(), /補足（提出用PDF）/);
api.setSelected([]);
api.schedulePdfImagePaneRefresh();
flush();
assert.match(paneHead(), /補足（提出用PDF）/, 'clearing the selection keeps the last shown target');

// 見るだけの表示（見開き・PDFイメージ）と狭い画面では組まない。
env.pageBoardView = 'pdf';
api.schedulePdfImagePaneRefresh();
assert.equal(env.frames.length, 0, 'the pane is not shown in read-only board views');
env.pageBoardView = 'thumbnail';
env.wide = false;
api.schedulePdfImagePaneRefresh();
assert.equal(env.frames.length, 0, 'the pane is not shown on narrow screens');
env.wide = true;

// 呼び出し点: 保存前（DOM の並び）・描き直し・選択・readiness 到着・頁数判明・表示切替。
const save = extractFunction('scheduleBoardSave');
assert.ok(save.indexOf('schedulePdfImagePaneRefresh(collectBoardVolumes());') >= 0 && save.indexOf('schedulePdfImagePaneRefresh(') < save.indexOf('saveBoardOrder()'),
  'scheduleBoardSave must refresh the pane from the DOM before the save request');
for (const [name, needle] of [['renderPages', 'schedulePdfImagePaneRefresh();'], ['syncPageSelectionUi', 'schedulePdfImagePaneRefresh();'],
  ['renderSpreadBoardIfVisible', 'schedulePdfImagePaneRefresh();'], ['pumpPageThumbnailQueue', 'schedulePdfImagePaneRefresh();'],
  ['syncPageBoardViewControls', 'syncPdfImagePaneControls();']]) {
  assert.ok(extractFunction(name).includes(needle), `${name} must keep the pane in step: ${needle}`);
}
// 盤面の Observer はパネルの canvas を見ない（パネルは root=パネルの専用 Observer）。
assert.ok(extractFunction('preparePageThumbnails').includes("!canvas.closest('#page-pdf-image-pane')"));
assert.ok(extractFunction('observePdfImagePaneThumbnails').includes('root:pane'));
// 見開きの頁・左右（サーバー確定値）には触れない。
for (const name of ['renderPdfImagePane', 'pdfImagePaneSheetHtml', 'pdfImagePaneTargetVolume']) {
  for (const banned of ['outputPageNumber', 'spreadIndex', 'buildSpreadLayout']) assert.ok(!extractFunction(name).includes(banned), `${name} must not use ${banned}`);
}

console.log('pdf image pane regression ok');
