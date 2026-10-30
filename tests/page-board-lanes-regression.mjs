// ページ構成の並びが保存されない不具合の回帰テスト。
//
// collectBoardVolumes() は画面上のレーン（[data-volume]）から並びを読み取り、
// 並べ替え前後の並びが違うときだけ保存する。提出用PDF画面のボタンにも data-volume が
// 付いていたため、文書全体を探すと本体・補足の並びが空で上書きされ、前後が同じに見えて
// 保存されず、PDFは変更前の構成で作られていた。
//
// app.js の実物の関数を取り出し、提出用PDF画面に同じ data-volume を持つ要素がある
// 偽のDOMで動かして、レーンの並びがそのまま読み取れることを確かめる。
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

class FakeElement {
  constructor(attrs = {}, children = []) { this.attrs = attrs; this.children = children; }
  getAttribute(name) { return Object.prototype.hasOwnProperty.call(this.attrs, name) ? this.attrs[name] : null; }
  matches(selector) {
    if (selector === '[data-volume]') return this.getAttribute('data-volume') !== null;
    if (selector === '.page-row') return String(this.attrs.class || '').split(/\s+/).includes('page-row');
    throw new Error(`unsupported selector ${selector}`);
  }
  descendants() { return this.children.flatMap(child => [child, ...child.descendants()]); }
  querySelectorAll(selector) { return this.descendants().filter(el => el.matches(selector)); }
}

const row = id => new FakeElement({ class: 'page-row', 'data-page-id': id });
const lane = (volume, ids) => new FakeElement({ 'data-volume': volume }, ids.map(row));
const board = new FakeElement({ id: 'page-board' }, [
  lane('none', ['u1']),
  lane('ja-main', ['p2', 'p1', 'p3']),
  lane('ja-appendix', ['a1']),
]);
// 提出用PDF画面は index.html でページ構成画面より後ろにある。ここの要素が
// 同じ data-volume を持っていても、レーンの並びを上書きしてはならない。
const finalScreen = new FakeElement({ id: 'final-target-grid' }, [
  new FakeElement({ 'data-final-volume': 'ja-main' }, [new FakeElement({ 'data-volume': 'ja-main' })]),
  new FakeElement({ 'data-final-volume': 'ja-appendix' }, [new FakeElement({ 'data-volume': 'ja-appendix' })]),
]);
const root = new FakeElement({}, [board, finalScreen]);
const document = { querySelectorAll: selector => root.querySelectorAll(selector) };
const $ = id => (id === 'page-board' ? board : null);

const factory = new Function('document', '$',
  `${extractFunction('pageBoardLanes')}\n${extractFunction('collectBoardVolumes')}\nreturn { collectBoardVolumes };`);
const { collectBoardVolumes } = factory(document, $);

assert.deepEqual(collectBoardVolumes(), {
  none: ['u1'],
  'ja-main': ['p2', 'p1', 'p3'],
  'ja-appendix': ['a1'],
}, 'board lanes must be read only from #page-board');

// 並べ替え後の並びが前と違って見えること（＝保存が呼ばれること）。
const before = JSON.stringify(collectBoardVolumes());
board.children[1].children.reverse();
assert.notEqual(JSON.stringify(collectBoardVolumes()), before, 'a reorder inside 本体 must change the collected layout');

// 文書全体から [data-volume] を探す経路と、レーン以外へ data-volume を付ける経路を残さない。
assert.ok(!source.includes("document.querySelectorAll('[data-volume]')"),
  'page-board lane lookups must be scoped to #page-board');
assert.ok(!/\.dataset\.volume\s*=/.test(source),
  'only page-board lanes may carry data-volume; final-screen controls must not');

// ---- 見開き表示（見るだけ） ----
// 台紙は physicalPages（提出用PDFの作成と同じ計算）から組む。1頁は台紙0の右、2・3頁は台紙1。
const { buildSpreadLayout } = new Function(`${extractFunction('buildSpreadLayout')}\nreturn { buildSpreadLayout };`)();
const phys = (pageId, from, count) => Array.from({ length: count }, (_, i) => {
  const n = from + i;
  return { outputPageNumber: n, side: n % 2 === 1 ? 'right' : 'left', spreadIndex: Math.floor(n / 2), pageId, pagePartIndex: i + 1, pagePartCount: count, sourcePdfPageNumber: i + 1, title: pageId };
});
const slotIds = spread => [spread.left, spread.right].map(cell => (cell ? `${cell.pageId}${cell.partIndex}` : '-'));

// 1・2・3ページのシートが混ざった本体: A(1頁) B(2頁) C(3頁) → 6頁・4見開き。
{
  const layout = buildSpreadLayout(['A', 'B', 'C'], [...phys('A', 1, 1), ...phys('B', 2, 2), ...phys('C', 4, 3)]);
  assert.equal(layout.pageCount, 6);
  assert.equal(layout.uncertainFrom, 0, 'a fresh layout must be certain');
  assert.deepEqual(layout.spreads.map(slotIds), [['-', 'A1'], ['B1', 'B2'], ['C1', 'C2'], ['C3', '-']],
    'page 1 is right of spread 0 with an empty left; later pages pair as left|right');
  const b2 = layout.cells.find(cell => cell.pageId === 'B' && cell.partIndex === 2);
  assert.deepEqual([b2.outputPageNumber, b2.side, b2.partCount, b2.laneIndex], [3, 'right', 2, 1]);
  assert.deepEqual(layout.cells.find(cell => cell.pageId === 'C' && cell.partIndex === 3).side, 'left');
}
// 頁と左右はサーバーの値を使い、クライアントで作り直さない（サーバーの side を信じる）。
{
  const pages = phys('A', 1, 1);
  pages[0].side = 'left-from-server';
  assert.equal(buildSpreadLayout(['A'], pages).cells[0].side, 'left-from-server');
}
// PDF未作成で physicalPages に出てこないページは「頁数未確定」の1マス。後ろは未確定。
{
  const layout = buildSpreadLayout(['A', 'X', 'B'], [...phys('A', 1, 1), ...phys('B', 2, 2)]);
  assert.deepEqual(layout.unknownPageIds, ['X']);
  assert.equal(layout.uncertainFrom, 2);
  assert.equal(layout.uncertainReason, 'unknown');
  const [a, x, b1, b2] = layout.cells;
  assert.ok(!a.uncertain && a.outputPageNumber === 1, 'pages before the unknown one stay certain');
  assert.ok(x.unknown && x.uncertain && x.outputPageNumber === 0);
  assert.ok(b1.uncertain && b2.uncertain && b1.outputPageNumber === 0 && b2.outputPageNumber === 0,
    'pages after an unknown page must not show the server numbers as final');
  assert.deepEqual(layout.spreads.map(slotIds), [['-', 'A1'], ['X1', 'B1'], ['B2', '-']]);
}
// 並べ替え直後（readiness の読み直し前）は、並びと physicalPages が食い違う。古い頁を確定値に見せない。
{
  const stale = buildSpreadLayout(['B', 'A'], [...phys('A', 1, 1), ...phys('B', 2, 2)]);
  assert.equal(stale.uncertainReason, 'stale');
  assert.ok(stale.cells.every(cell => cell.uncertain && cell.outputPageNumber === 0));
  assert.deepEqual(stale.spreads.map(slotIds), [['-', 'B1'], ['B2', 'A1']], 'placement still follows the board order');
  const orphan = buildSpreadLayout(['A'], [...phys('A', 1, 1), ...phys('Z', 2, 1)]);
  assert.deepEqual(orphan.orphanPageIds, ['Z']);
  assert.ok(orphan.cells.every(cell => cell.uncertain), 'a page moved out of the target leaves the old numbers stale');
  const pending = buildSpreadLayout(['A'], phys('A', 1, 1), { pending: true });
  assert.equal(pending.uncertainReason, 'pending');
  assert.ok(pending.cells[0].uncertain, 'a pending readiness reload marks the layout as updating');
  assert.deepEqual(buildSpreadLayout([], []).spreads, []);
}

// 見開きの DOM: 並びの器（data-volume）は出力先ごとに1つで、各ページを1回だけ持つ。
// 台紙・マスに data-volume を付けると、collectBoardVolumes が重複した並びを読んでしまう。
{
  const stubs = `
    const selectedPages=new Set(),highlightedPageIds=new Set();
    const asArray=value=>value==null?[]:(Array.isArray(value)?value:[value]);
    const escapeAttr=value=>escapeHtml(value).replace(/\`/g,'&#96;');
    const resolvedPageId=page=>String(page?.pageId||'');
    const getWorkbook=()=>({displayName:'営業.xlsx'}),sourceUnitReference=(wb,sheet)=>String(sheet||'');
    const pagePreviewAvailable=page=>!!page?.contentPdf;
    const finalReadinessKey=()=>'pack',state={finalReadiness:{pack:{receivedAt:0}}};
    const readiness={'ja-main':{blockers:[],snapshot:{physicalPages:hooks.physicalPages}}};
    const volumeReadiness=volume=>readiness[volume]||{};`;
  const html = new Function('hooks', `${stubs}
    ${extractFunction('escapeHtml')}
    ${extractFunction('filteredOutPageHtml')}
    ${extractFunction('buildSpreadLayout')}
    ${extractFunction('spreadReadinessTimeText')}
    ${extractFunction('spreadCellHtml')}
    ${extractFunction('spreadVolumeContentHtml')}
    return spreadVolumeContentHtml(hooks.panel,hooks.pages,new Set(hooks.pages.map(page=>page.pageId)));
  `)({
    panel: { volume: 'ja-main', title: '本体' },
    pages: ['A', 'B', 'C'].map(pageId => ({ pageId, title: pageId, sheetName: pageId, contentPdf: `${pageId}.pdf` })),
    physicalPages: [...phys('A', 1, 1), ...phys('B', 2, 2), ...phys('C', 4, 3)],
  });
  assert.equal((html.match(/data-volume=/g) || []).length, 1, 'only the hidden lane index may carry data-volume');
  const index = html.match(/<div class="spread-lane-index" data-volume="ja-main" hidden>(.*?)<\/div>/);
  assert.ok(index, 'the spread panel keeps a hidden lane index for collectBoardVolumes');
  assert.deepEqual([...index[1].matchAll(/data-page-id="([^"]+)"/g)].map(match => match[1]), ['A', 'B', 'C']);
  assert.ok(!/class="spread-cell[^"]*page-row/.test(html) && !/class="page-row[^"]*spread-cell/.test(html),
    'spread cells must not be .page-row, or drag/keyboard reorder would treat them as cards');
  assert.equal((html.match(/class="spread-cell /g) || []).length, 6, 'one cell per physical page');
  assert.equal((html.match(/class="spread-sheet"/g) || []).length, 4);
  assert.ok(html.includes('data-thumbnail-page="3"'), 'the third page of a sheet draws page 3 of its PDF');
  assert.ok(html.includes('（空き）'), 'spread 0 draws an empty left slot');
  assert.ok(html.includes('2/2') && html.includes('次の見開きへ続く'), 'multi-page sheets show part numbers and continuation');

  // 同じ DOM 形を偽の要素で組み、collectBoardVolumes の結果が変わらないことを確かめる。
  const cell = (id) => new FakeElement({ class: 'spread-cell', 'data-page-id': id });
  const spreadBoard = new FakeElement({ id: 'page-board' }, [
    lane('none', ['u1']),
    new FakeElement({ class: 'spread-wrap' }, [
      lane('ja-main', ['A', 'B', 'C']),
      new FakeElement({ class: 'spread-list' }, [cell('A'), cell('B'), cell('B'), cell('C'), cell('C'), cell('C')]),
    ]),
  ]);
  const spreadFactory = new Function('$', `${extractFunction('pageBoardLanes')}\n${extractFunction('collectBoardVolumes')}\nreturn collectBoardVolumes;`);
  assert.deepEqual(spreadFactory(id => (id === 'page-board' ? spreadBoard : null))(), { none: ['u1'], 'ja-main': ['A', 'B', 'C'] },
    'the spread view must not change what collectBoardVolumes reads');
}

// サムネイルのキャッシュは、描くページ番号ごとに分ける。サムネイル表示は従来どおり1枚目。
{
  const { pageThumbnailCacheKey } = new Function(`
    const getWorkbook=()=>({lastRenderedSnapshotId:'s',lastRenderedVersionId:'v'}),resolvedPageId=page=>String(page?.pageId||'');
    ${extractFunction('pageThumbnailCacheKey')}
    return { pageThumbnailCacheKey };`)();
  const page = { pageId: 'B', contentPdf: 'B.pdf' };
  assert.equal(pageThumbnailCacheKey(page), pageThumbnailCacheKey(page, 1));
  assert.notEqual(pageThumbnailCacheKey(page, 1), pageThumbnailCacheKey(page, 2));
}

// 見開きではドラッグと Alt+矢印の並べ替えを止め、Undo の before は状態から読む。
assert.ok(extractFunction('beginPointerPageDrag').includes("if (pageBoardView === 'spread') return;"),
  'pointer drag must not start in the spread view');
assert.ok(extractFunction('handleBoardRowKeydown').includes("pageBoardView==='spread'&&e.altKey"),
  'Alt+arrow reorder must be stopped in the spread view');
assert.ok(extractFunction('movePageIdsToVolume').includes("pageBoardView==='spread'?pageVolumesFromState():collectBoardVolumes()"),
  'the undo baseline for a move from the spread view must come from state');

console.log('page-board lanes regression ok');
