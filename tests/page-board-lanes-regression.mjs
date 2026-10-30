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

// サムネイルのキャッシュは、描くページ番号ごとに分ける。番号を省けば1枚目。
{
  const { pageThumbnailCacheKey } = new Function(`
    const getWorkbook=()=>({lastRenderedSnapshotId:'s',lastRenderedVersionId:'v'}),resolvedPageId=page=>String(page?.pageId||'');
    ${extractFunction('pageDocumentCacheKey')}
    ${extractFunction('pageThumbnailCacheKey')}
    return { pageThumbnailCacheKey };`)();
  const page = { pageId: 'B', contentPdf: 'B.pdf' };
  assert.equal(pageThumbnailCacheKey(page), pageThumbnailCacheKey(page, 1));
  assert.notEqual(pageThumbnailCacheKey(page, 1), pageThumbnailCacheKey(page, 2));
}

// サムネイルは使用ページ範囲の先頭ページを描き、複数ページの項目には「全N枚」を出す。
{
  const { pageFirstSourcePage, pagePartCountsFromPhysicalPages } = new Function(`
    const resolvedPageId=page=>String(page?.pageId||'');
    ${extractFunction('pageFirstSourcePage')}
    ${extractFunction('pagePartCountsFromPhysicalPages')}
    return { pageFirstSourcePage, pagePartCountsFromPhysicalPages };`)();
  assert.equal(pageFirstSourcePage({ pageId: 'A' }), 1, 'no range draws page 1');
  assert.equal(pageFirstSourcePage({ pageId: 'A', pageRange: null }), 1);
  assert.equal(pageFirstSourcePage({ pageId: 'A', pageRange: { start: 3, end: 5 } }), 3, 'range 3-5 draws page 3');
  assert.equal(pageFirstSourcePage({ pageId: 'A', pageRange: { start: 4, end: 4 } }), 4, 'single page 4 draws page 4');
  assert.equal(pageFirstSourcePage({ pageId: 'A', pageRange: '3-5' }), 3);
  assert.equal(pageFirstSourcePage({ pageId: 'A', pageRange: '4' }), 4);
  assert.equal(pageFirstSourcePage({ pageId: 'A', pageRange: { start: 0 } }), 1, 'invalid starts fall back to page 1');

  const parts = (pageId, start, count, from = 1) => Array.from({ length: count }, (_, i) => ({
    pageId, outputPageNumber: from + i, pagePartIndex: i + 1, pagePartCount: count, sourcePdfPageNumber: start + i,
  }));
  const pages = [
    { pageId: 'A' },
    { pageId: 'B' },
    { pageId: 'C', pageRange: { start: 3, end: 5 } },
    { pageId: 'D', pageRange: { start: 2, end: 3 } },
    { pageId: 'E', pageRange: { start: 1, end: 2 } },
    { pageId: 'F' },
  ];
  const counts = pagePartCountsFromPhysicalPages(pages, [
    [...parts('A', 1, 1), ...parts('B', 1, 5, 2), ...parts('C', 3, 3, 7)],
    // D は範囲の開始を変えた直後で、readiness は古い範囲（1-2）のまま。
    [...parts('D', 1, 2), ...parts('E', 1, 4, 3)],
    undefined,
  ]);
  assert.deepEqual([...counts], [['B', 5], ['C', 3]],
    'only multi-page items whose server range matches the current range get a count');
  assert.ok(!counts.has('A'), 'single-page items get no badge');
  assert.ok(!counts.has('D'), 'a stale start (range changed, readiness not reloaded) gets no badge');
  assert.ok(!counts.has('E'), 'more parts than the current range allows is stale and gets no badge');
  assert.ok(!counts.has('F'), 'pages missing from physicalPages get no badge');
  // 移した直後で、古い出力先と新しい出力先の readiness が食い違うときは出さない。
  const moved = pagePartCountsFromPhysicalPages([{ pageId: 'B' }], [parts('B', 1, 5), parts('B', 1, 3)]);
  assert.equal(moved.size, 0, 'conflicting counts across targets get no badge');
  assert.equal(pagePartCountsFromPhysicalPages([{ pageId: 'B' }], [parts('B', 1, 5), parts('B', 1, 5)]).get('B'), 5,
    'the same count in two snapshots is not doubled');
  assert.equal(pagePartCountsFromPhysicalPages(pages, null).size, 0);

  const card = extractFunction('pageThumbnailHtml');
  assert.ok(card.includes('data-thumbnail-page="${escapeAttr(pageFirstSourcePage(p))}"'),
    'the thumbnail canvas must draw the first page of the used range');
  assert.ok(card.includes('page-thumb-parts') && card.includes('pageThumbnailPartCounts.get(pid)'),
    'multi-page items show the part count badge');
  const render = extractFunction('renderPages');
  assert.ok(render.includes('pagePartCountsFromPhysicalPages(') && render.includes('[...pageThumbnailPartCounts]'),
    'renderPages must build the part counts and redraw when they change');
  assert.ok(extractFunction('renderSpreadBoardIfVisible').includes("pageBoardView==='thumbnail'&&!draggingRow"),
    'readiness reloads redraw thumbnails, but not in the middle of a drag');
}

// ---- PDFイメージ表示（見るだけ） ----
// 紙は物理ページ1枚ずつ。並びの器（data-volume）は出力先ごとに1つで、各ページを1回だけ持つ。
{
  const stubs = `
    const selectedPages=new Set(['B']),highlightedPageIds=new Set();
    let activePageVolume=hooks.activePageVolume||'ja-main';
    const asArray=value=>value==null?[]:(Array.isArray(value)?value:[value]);
    const escapeAttr=value=>escapeHtml(value).replace(/\`/g,'&#96;');
    const resolvedPageId=page=>String(page?.pageId||'');
    const workbooks={fresh:{displayName:'営業.xlsx',latest:true},old:{displayName:'旧.xlsx',latest:false}};
    const getWorkbook=id=>workbooks[id]||null,sourceUnitReference=(wb,sheet)=>String(sheet||'');
    const isLatestPdfWorkbook=wb=>!!wb?.latest;
    const pagePreviewAvailable=page=>!!(page?.contentPdf&&(!getWorkbook(page.workbookId)||isLatestPdfWorkbook(getWorkbook(page.workbookId))));
    const pageThumbnailCacheKey=(page,n)=>page.pageId+'|p'+n,pdfImagePaperRatios=new Map([['A|p1','620 / 440']]);
    const readiness={'ja-main':{blockers:[],physicalPagesPending:!!hooks.pending,snapshot:{outputFileName:'本体.pdf',physicalPages:hooks.physicalPages}},'ja-opt':{snapshot:{physicalPages:[]}}};
    const volumeReadiness=volume=>readiness[volume]||{};
    const activeTargetDefinitions=()=>[{targetId:'main',required:true},{targetId:'opt',required:false}];
    const targetVolume=id=>'ja-'+id;`;
  const render = hooks => new Function('hooks', `${stubs}
    ${extractFunction('escapeHtml')}
    ${extractFunction('filteredOutPageHtml')}
    ${extractFunction('pageFirstSourcePage')}
    ${extractFunction('pageLastSourcePage')}
    ${extractFunction('collectPdfImagePageCounts')}
    ${extractFunction('buildPdfImageLayout')}
    ${extractFunction('pdfImageSheetHtml')}
    ${extractFunction('pdfImageVolumeContentHtml')}
    ${extractFunction('pdfImageUnassignedStripHtml')}
    const pdfImagePageCounts=collectPdfImagePageCounts([readiness['ja-main']],hooks.observed||new Map(),hooks.pages);
    return {
      main: pdfImageVolumeContentHtml(hooks.panel,hooks.pages,new Set(hooks.visible||hooks.pages.map(page=>page.pageId))),
      unassigned: pdfImageUnassignedStripHtml({volume:'none',title:'未振り分け'},hooks.unassigned||[]),
      optional: pdfImageVolumeContentHtml({volume:'ja-opt',title:'任意'},[],new Set()),
    };
  `)(hooks);
  const base = {
    panel: { volume: 'ja-main', title: '本体' },
    pages: [
      { pageId: 'A', workbookId: 'fresh', sheetName: 'A', contentPdf: 'A.pdf' },
      { pageId: 'B', workbookId: 'fresh', sheetName: 'B', contentPdf: 'B.pdf', pageRange: { start: 2, end: 3 } },
      { pageId: 'C', workbookId: 'fresh', sheetName: 'C', contentPdf: 'C.pdf' },
    ],
    physicalPages: [...phys('A', 1, 1), ...phys('B', 2, 2).map((entry, i) => ({ ...entry, sourcePdfPageNumber: 2 + i })), ...phys('C', 4, 3)],
  };
  const { main: html, unassigned, optional } = render(base);
  assert.equal((html.match(/data-volume=/g) || []).length, 1, 'only the hidden lane index may carry data-volume');
  const index = html.match(/<div class="spread-lane-index" data-volume="ja-main" hidden>(.*?)<\/div>/);
  assert.ok(index, 'the PDF image panel keeps a hidden lane index for collectBoardVolumes');
  assert.deepEqual([...index[1].matchAll(/data-page-id="([^"]+)"/g)].map(match => match[1]), ['A', 'B', 'C']);
  assert.ok(!/class="pdf-image-sheet[^"]*page-row/.test(html) && !/class="page-row[^"]*pdf-image-sheet/.test(html),
    'PDF image sheets must not be .page-row, or drag/keyboard reorder would treat them as cards');
  const sheets = [...html.matchAll(/class="pdf-image-sheet [^"]*" data-page-id="([^"]+)" data-pdf-image-position="(\d+)" data-source-page="(\d+)"/g)]
    .map(match => `${match[1]}${match[3]}@${match[2]}`);
  assert.deepEqual(sheets, ['A1@1', 'B2@2', 'B3@3', 'C1@4', 'C2@5', 'C3@6'],
    'one sheet per physical page, in output order, drawing the used page of the converted PDF');
  assert.ok(html.includes('本体（提出用PDF）6 ページ') && html.includes('本体.pdf'), 'the summary shows the page count and file name');
  assert.ok(!/\d+ (左|右)|outputPageNumber/.test(html), 'the PDF image must not show page numbers or sides');
  assert.ok(html.includes('data-set-page-board-view="spread"'), 'a link sends page/side questions to the spread view');
  assert.ok(html.includes('style="aspect-ratio:620 / 440"'), 'a known paper ratio is kept across redraws');
  assert.ok(html.includes('本体 2番目・B（1/2）'), 'the caption names the target, position and part');
  assert.ok(/class="pdf-image-sheet [^"]*selected-row[^"]*" data-page-id="B"/.test(html), 'selection is per pageId: every sheet of B is selected');
  assert.ok(!html.includes('pdf-image-file-break'), 'a single target tab draws no file break');
  assert.ok(optional.includes('このPDFは0ページのため出力されません'), 'an empty optional target says it will not be output');

  // 未振り分けは紙にしない。帯と、collectBoardVolumes 用の器だけ。
  const strip = render({ ...base, unassigned: [{ pageId: 'u1' }, { pageId: 'u2' }] }).unassigned;
  assert.ok(!strip.includes('pdf-image-sheet') && !strip.includes('<canvas'), 'unassigned pages are not drawn as paper');
  assert.ok(strip.includes('未振り分け 2ページ') && strip.includes('data-set-page-board-view="thumbnail"'));
  assert.deepEqual([...strip.matchAll(/data-page-id="([^"]+)"/g)].map(match => match[1]), ['u1', 'u2']);
  assert.equal((strip.match(/data-volume="none"/g) || []).length, 1);

  // 絞り込み中も紙は消さず、薄くする。「全体」では出力先ごとにファイルを区切る。
  const filtered = render({ ...base, visible: ['A'], activePageVolume: 'all' }).main;
  assert.equal((filtered.match(/class="pdf-image-sheet /g) || []).length, 6, 'filtering must not remove sheets');
  assert.equal((filtered.match(/pdf-image-dimmed/g) || []).length, 5);
  assert.ok(filtered.includes('pdf-image-file-break') && filtered.includes('ここから 本体.pdf'));

  // 並べ替え直後（readiness は古い並びのまま・pending）でも、今の並びで即時に組む。頁数は変わらない。
  const reordered = render({ ...base, pages: [base.pages[2], base.pages[0], base.pages[1]], pending: true }).main;
  assert.deepEqual([...reordered.matchAll(/data-page-id="([^"]+)" data-pdf-image-position/g)].map(match => match[1]),
    ['C', 'C', 'C', 'A', 'B', 'B'], 'a reorder shows the new order before readiness is reloaded');
  assert.ok(reordered.includes('約 6 ページ') && reordered.includes('頁数を確認中'), 'a pending readiness only makes the count approximate');

  // 古い変換PDFは描いて斜線と帯を重ねる（allow-stale）。PDF未作成は点線の紙。不明は1枚分。
  const mixed = render({ ...base, pages: [
    { pageId: 'A', workbookId: 'fresh', sheetName: 'A', contentPdf: 'A.pdf' },
    { pageId: 'S', workbookId: 'old', sheetName: 'S', contentPdf: 'S.pdf' },
    { pageId: 'N', workbookId: 'old', sheetName: 'N', status: 'render-error' },
  ], physicalPages: phys('A', 1, 1), observed: new Map([['S', 2]]) }).main;
  assert.ok(/class="pdf-image-sheet pdf-image-stale[^"]*" data-page-id="S"/.test(mixed), 'a stale converted PDF is marked');
  assert.ok(mixed.includes('data-page-thumbnail="S" data-thumbnail-page="1" data-thumbnail-allow-stale="1"'), 'a stale sheet may draw the old PDF');
  assert.ok(!mixed.includes('data-page-thumbnail="A" data-thumbnail-page="1" data-thumbnail-allow-stale'), 'fresh sheets do not need allow-stale');
  assert.ok(mixed.includes('原稿が更新されています（変換PDFが古い）'));
  assert.ok(mixed.includes('変換PDFが古いページが 1件') && mixed.includes('変換PDFの作成が必要なページが 2件'),
    'a page without any converted PDF is counted as missing, not as stale');
  assert.ok(/class="pdf-image-sheet [^"]*pdf-image-unknown[^"]*" data-page-id="N"/.test(mixed) && mixed.includes('PDF未作成（頁数不明）') && mixed.includes('作成エラー'),
    'a page without a converted PDF is a dashed placeholder (with the render error label)');
  assert.ok(!mixed.includes('data-page-thumbnail="N"'), 'a page without a converted PDF draws nothing');
  assert.ok(mixed.includes('約 4 ページ') && mixed.includes('このままでは出力できません') && mixed.includes('data-page-go-source'),
    'stale/unknown pages make the count approximate and point to the source screen');

  // PDFイメージの DOM 形でも collectBoardVolumes の結果は変わらない（重複せず、未振り分けも含む）。
  const sheet = id => new FakeElement({ class: 'pdf-image-sheet', 'data-page-id': id });
  const pdfBoard = new FakeElement({ id: 'page-board' }, [
    new FakeElement({ class: 'volume-panel inbox pdf-image-unassigned' }, [lane('none', ['u1', 'u2'])]),
    new FakeElement({ class: 'pdf-image-wrap' }, [
      lane('ja-main', ['A', 'B', 'C']),
      new FakeElement({ class: 'pdf-image-list' }, [sheet('A'), sheet('B'), sheet('B'), sheet('C'), sheet('C'), sheet('C')]),
    ]),
  ]);
  const pdfFactory = new Function('$', `${extractFunction('pageBoardLanes')}\n${extractFunction('collectBoardVolumes')}\nreturn collectBoardVolumes;`);
  assert.deepEqual(pdfFactory(id => (id === 'page-board' ? pdfBoard : null))(), { none: ['u1', 'u2'], 'ja-main': ['A', 'B', 'C'] },
    'the PDF image view must not change what collectBoardVolumes reads');
}

// 見開き・PDFイメージではドラッグと Alt+矢印の並べ替えを止め、Undo の before は状態から読む。
{
  const { isReadOnlyBoardView } = new Function(`let pageBoardView='thumbnail';${extractFunction('isReadOnlyBoardView')}\nreturn { isReadOnlyBoardView };`)();
  assert.deepEqual(['thumbnail', 'detail', 'spread', 'pdf'].map(view => isReadOnlyBoardView(view)), [false, false, true, true]);
  assert.equal(isReadOnlyBoardView(), false);
}
assert.ok(extractFunction('beginPointerPageDrag').includes('if (isReadOnlyBoardView()) return;'),
  'pointer drag must not start in the spread / PDF image views');
assert.ok(extractFunction('handleBoardRowKeydown').includes('isReadOnlyBoardView()&&e.altKey'),
  'Alt+arrow reorder must be stopped in the spread / PDF image views');
assert.ok(extractFunction('movePageIdsToVolume').includes('isReadOnlyBoardView()?pageVolumesFromState():collectBoardVolumes()'),
  'the undo baseline for a move from a read-only view must come from state');
assert.ok(extractFunction('renderSpreadBoardIfVisible').includes('isReadOnlyBoardView()'),
  'readiness reloads redraw the PDF image view too');
// PDFイメージの紙も見開きのマスと同じ選択・キー操作（Space/Enter/M、↑↓で前後の紙）を使う。
assert.ok(extractFunction('spreadCellsInList').includes('.pdf-image-list') && extractFunction('spreadCellKeyboardTarget').includes("classList.contains('pdf-image-sheet')"),
  'PDF image sheets share selection and arrow navigation');
assert.ok(extractFunction('delegateBoardEvents').includes("closest('.spread-cell,.pdf-image-sheet')"),
  'clicks and keys on PDF image sheets go through the read-only cell handlers');

console.log('page-board lanes regression ok');
