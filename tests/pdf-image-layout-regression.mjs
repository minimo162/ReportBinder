// PDFイメージ表示の土台（頁数表と物理ページの並び）の回帰テスト。
//
// 提出用PDFの物理ページの並びは「ページ項目の並び」と「各項目の使用頁数」だけで決まる。
// 頁数は並べ替えでは変わらないので、クライアントで頁数表を持てば、並べ替えた直後でも
// readiness の再計算を待たずに並びを組める。見開きの頁・左右（buildSpreadLayout の
// サーバー確定値）には触れない。
//
// app.js の実物の collectPdfImagePageCounts / buildPdfImageLayout を取り出して確かめる。
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

// 原稿（workbook）の最新・古いは isLatestPdfWorkbook の代わりに latest で表す。
const workbooks = new Map([
  ['wb-ok', { workbookId: 'wb-ok', latest: true }],
  ['wb-old', { workbookId: 'wb-old', latest: false }],
]);
const { collectPdfImagePageCounts, buildPdfImageLayout, pageLastSourcePage } = new Function('workbooks', `
  const resolvedPageId=page=>String(page?.pageId||'');
  const getWorkbook=id=>workbooks.get(String(id||''))||null;
  const isLatestPdfWorkbook=w=>!!w?.latest;
  ${extractFunction('pageFirstSourcePage')}
  ${extractFunction('pageLastSourcePage')}
  ${extractFunction('collectPdfImagePageCounts')}
  ${extractFunction('buildPdfImageLayout')}
  return { collectPdfImagePageCounts, buildPdfImageLayout, pageLastSourcePage };`)(workbooks);

// readiness の physicalPages（出力先ごとに1頁から）。頁・左右も付けるが、PDFイメージは使わない。
const parts = (pageId, start, count, from = 1) => Array.from({ length: count }, (_, i) => ({
  pageId, outputPageNumber: from + i, side: (from + i) % 2 ? 'right' : 'left', spreadIndex: Math.floor((from + i) / 2),
  pagePartIndex: i + 1, pagePartCount: count, sourcePdfPageNumber: start + i,
}));
const page = (pageId, extra = {}) => ({ pageId, workbookId: 'wb-ok', contentPdf: `content-pdf/${pageId}.pdf`, ...extra });
const sheetsOf = layout => layout.sheets.map(s => `${s.pageId}:${s.sourcePdfPageNumber}${s.unknown ? '?' : ''}${s.stale ? '!' : ''}`);

// 使用ページ範囲の終わり。
assert.equal(pageLastSourcePage({}), 0, 'no range means up to the last page');
assert.equal(pageLastSourcePage({ pageRange: { start: 3, end: 5 } }), 5);
assert.equal(pageLastSourcePage({ pageRange: { start: 4 } }), 4);
assert.equal(pageLastSourcePage({ pageRange: '3-5' }), 5);
assert.equal(pageLastSourcePage({ pageRange: '4' }), 4);

// 単ページだけの並び。
{
  const pages = [page('A'), page('B'), page('C')];
  const counts = collectPdfImagePageCounts([[...parts('A', 1, 1), ...parts('B', 1, 1, 2), ...parts('C', 1, 1, 3)]], new Map(), pages);
  const layout = buildPdfImageLayout(['A', 'B', 'C'], counts);
  assert.deepEqual(sheetsOf(layout), ['A:1', 'B:1', 'C:1']);
  assert.equal(layout.pageCount, 3);
  assert.equal(layout.approximate, false);
  assert.deepEqual(layout.sheets.map(s => s.laneIndex), [0, 1, 2]);
  // 頁・左右は持たない。
  for (const sheet of layout.sheets) {
    for (const banned of ['outputPageNumber', 'side', 'spreadIndex', 'position']) assert.ok(!(banned in sheet), `sheets must not carry ${banned}`);
  }
}

// 複数ページのシートの展開と pageRange（start/end）の反映。
{
  const pages = [page('A'), page('B'), page('C', { pageRange: { start: 3, end: 5 } }), page('D', { pageRange: '2' })];
  const counts = collectPdfImagePageCounts([[...parts('A', 1, 1), ...parts('B', 1, 4, 2), ...parts('C', 3, 3, 6), ...parts('D', 2, 1, 9)]], new Map(), pages);
  assert.deepEqual(counts.get('B'), { start: 1, count: 4, source: 'server', stale: false, clamped: false });
  assert.deepEqual(counts.get('C'), { start: 3, count: 3, source: 'server', stale: false, clamped: false });
  const layout = buildPdfImageLayout(['A', 'B', 'C', 'D'], counts);
  assert.deepEqual(sheetsOf(layout), ['A:1', 'B:1', 'B:2', 'B:3', 'B:4', 'C:3', 'C:4', 'C:5', 'D:2']);
  assert.deepEqual(layout.sheets.filter(s => s.pageId === 'B').map(s => `${s.partIndex}/${s.partCount}`), ['1/4', '2/4', '3/4', '4/4']);
  assert.equal(layout.pageCount, 9);
  assert.equal(layout.approximate, false);
}

// 出力先が複数でも、同じページの値は二重に数えない。readiness の配列（snapshot 付き）も読める。
{
  const pages = [page('A'), page('B')];
  const counts = collectPdfImagePageCounts([
    { snapshot: { physicalPages: parts('A', 1, 3) } },
    { snapshot: { physicalPages: [...parts('B', 1, 2), ...parts('A', 1, 3, 3)] } },
    null,
  ], new Map(), pages);
  assert.equal(counts.get('A').count, 3);
  assert.equal(counts.get('B').count, 2);
}

// server と observed が両方あるときは server を優先し、無いものは observed で補う。
{
  const pages = [page('A'), page('B'), page('C', { pageRange: { start: 2, end: 3 } })];
  const observed = new Map([['A', 9], ['B', 4], ['C', 6]]);
  const counts = collectPdfImagePageCounts([parts('A', 1, 3)], observed, pages);
  assert.deepEqual(counts.get('A'), { start: 1, count: 3, source: 'server', stale: false, clamped: false }, 'server wins over observed');
  assert.deepEqual(counts.get('B'), { start: 1, count: 4, source: 'observed', stale: false, clamped: false }, 'observed fills a page missing from readiness');
  assert.deepEqual(counts.get('C'), { start: 2, count: 2, source: 'observed', stale: false, clamped: false }, 'observed applies the page range');
  assert.deepEqual(sheetsOf(buildPdfImageLayout(['C', 'B'], counts)), ['C:2', 'C:3', 'B:1', 'B:2', 'B:3', 'B:4']);
}

// 範囲を変えた直後（readiness は古い範囲のまま）・移した直後で出力先の値が食い違うときは server を使わない。
{
  const pages = [page('D', { pageRange: { start: 2, end: 3 } }), page('E', { pageRange: { start: 1, end: 2 } }), page('F')];
  const observed = new Map([['D', 5], ['E', 5], ['F', 5]]);
  const counts = collectPdfImagePageCounts([[...parts('D', 1, 2), ...parts('E', 1, 4, 3)], parts('F', 1, 5), parts('F', 1, 3)], observed, pages);
  assert.deepEqual(counts.get('D'), { start: 2, count: 2, source: 'observed', stale: false, clamped: false }, 'a stale range start falls back to observed');
  assert.deepEqual(counts.get('E'), { start: 1, count: 2, source: 'observed', stale: false, clamped: false }, 'more parts than the range allows falls back to observed');
  assert.deepEqual(counts.get('F'), { start: 1, count: 5, source: 'observed', stale: false, clamped: false }, 'conflicting targets fall back to observed');
  const unknown = collectPdfImagePageCounts([parts('D', 1, 2)], new Map(), [pages[0]]);
  assert.equal(unknown.get('D').source, 'unknown', 'without observed counts a stale range is unknown');
}

// 範囲外は存在する頁で頭打ちにする（頁数が分かっているとき）。
{
  const pages = [page('A', { pageRange: { start: 3, end: 7 } }), page('B', { pageRange: { start: 7, end: 8 } }), page('C', { pageRange: { start: 2, end: 6 } })];
  const observed = new Map([['A', 5], ['B', 5], ['C', 4]]);
  // サーバーは範囲どおり（変換PDFの頁数を超えて）physicalPages を出す（page-range-out-of-bounds の blocker 付き）。
  const counts = collectPdfImagePageCounts([[...parts('A', 3, 5), ...parts('C', 2, 5, 6)]], observed, pages);
  assert.deepEqual(counts.get('A'), { start: 3, count: 3, source: 'server', stale: false, clamped: true });
  assert.deepEqual(counts.get('B'), { start: 7, count: 0, source: 'observed', stale: false, clamped: true });
  assert.deepEqual(counts.get('C'), { start: 2, count: 3, source: 'server', stale: false, clamped: true });
  const layout = buildPdfImageLayout(['A', 'B', 'C'], counts);
  assert.deepEqual(sheetsOf(layout), ['A:3', 'A:4', 'A:5', 'C:2', 'C:3', 'C:4']);
  assert.deepEqual(layout.clampedPageIds, ['A', 'B', 'C']);
  // 頁数が分からなければ範囲どおり（サーバー値）のまま。
  assert.equal(collectPdfImagePageCounts([parts('A', 3, 5)], new Map(), [pages[0]]).get('A').count, 5);
}

// 未知のページ（readiness にも observed にも無い・変換PDF未作成）は1枚分の場所を取り、approximate になる。
{
  const pages = [page('A'), page('N'), page('M', { contentPdf: '', pageRange: { start: 3, end: 4 } }), page('Z')];
  const counts = collectPdfImagePageCounts([[...parts('A', 1, 2), ...parts('Z', 1, 1, 3)]], new Map([['M', 9]]), pages);
  assert.deepEqual(counts.get('N'), { start: 1, count: 1, source: 'unknown', stale: false, clamped: false });
  assert.equal(counts.get('M').source, 'unknown', 'a page without a content PDF is unknown even if a count was observed');
  const layout = buildPdfImageLayout(['A', 'N', 'M', 'Z', 'ghost'], counts);
  assert.deepEqual(sheetsOf(layout), ['A:1', 'A:2', 'N:1?', 'M:3?', 'Z:1', 'ghost:1?']);
  assert.equal(layout.pageCount, 6);
  assert.equal(layout.approximate, true);
  assert.deepEqual(layout.unknownPageIds, ['N', 'M', 'ghost']);
  // 未知のページより後ろも普通に並ぶ（見開きのように未確定で止めない）。
  assert.deepEqual(layout.sheets.at(-2), { pageId: 'Z', laneIndex: 3, partIndex: 1, partCount: 1, sourcePdfPageNumber: 1, unknown: false, stale: false });
}

// 古い変換PDF（原稿更新後に未再変換）は stale 付きで並び、approximate になる。
{
  const pages = [page('A'), page('S', { workbookId: 'wb-old' })];
  const counts = collectPdfImagePageCounts([[...parts('A', 1, 1), ...parts('S', 1, 2, 2)]], new Map(), pages);
  assert.deepEqual(counts.get('S'), { start: 1, count: 2, source: 'server', stale: true, clamped: false });
  const layout = buildPdfImageLayout(['A', 'S'], counts);
  assert.deepEqual(sheetsOf(layout), ['A:1', 'S:1!', 'S:2!']);
  assert.deepEqual(layout.stalePageIds, ['S']);
  assert.equal(layout.approximate, true);
  // 古い原稿の変換PDFも頁数が分からなければ未知（stale も付ける）。
  const unknownStale = buildPdfImageLayout(['S'], collectPdfImagePageCounts([], new Map(), [pages[1]]));
  assert.deepEqual(sheetsOf(unknownStale), ['S:1?!']);
}

// readiness 待ち（physicalPagesPending）は「約」。
{
  const pages = [page('A')];
  const counts = collectPdfImagePageCounts([parts('A', 1, 1)], new Map(), pages);
  assert.equal(buildPdfImageLayout(['A'], counts).approximate, false);
  assert.equal(buildPdfImageLayout(['A'], counts, { pending: true }).approximate, true);
}

// 並べ替えても頁数表を引き直すだけで結果が変わる（readiness は古いままでよい）。
{
  const pages = [page('A'), page('B'), page('C', { pageRange: { start: 2, end: 3 } })];
  const staleReadiness = [[...parts('A', 1, 1), ...parts('B', 1, 3, 2), ...parts('C', 2, 2, 5)]];
  const counts = collectPdfImagePageCounts(staleReadiness, new Map(), pages);
  const before = buildPdfImageLayout(['A', 'B', 'C'], counts);
  const after = buildPdfImageLayout(['C', 'A', 'B'], counts);
  assert.deepEqual(sheetsOf(before), ['A:1', 'B:1', 'B:2', 'B:3', 'C:2', 'C:3']);
  assert.deepEqual(sheetsOf(after), ['C:2', 'C:3', 'A:1', 'B:1', 'B:2', 'B:3'], 'a reorder is reflected immediately');
  assert.equal(after.approximate, false, 'a reorder alone does not make the layout uncertain');
  assert.equal(after.pageCount, before.pageCount);
  // 別の出力先から移してきたページ（この出力先の readiness には無い）も、他の出力先の値で数える。
  const moved = collectPdfImagePageCounts([parts('A', 1, 1), parts('X', 1, 2)], new Map(), [page('A'), page('X')]);
  assert.deepEqual(sheetsOf(buildPdfImageLayout(['X', 'A'], moved)), ['X:1', 'X:2', 'A:1']);
}

// 未振り分けの pageId は渡さない前提: 頁数表は全ページ分あっても、並びには渡した pageId だけが出る。
{
  const pages = [page('A'), page('U', { volume: 'none' }), page('B')];
  const counts = collectPdfImagePageCounts([[...parts('A', 1, 1), ...parts('B', 1, 1, 2)]], new Map([['U', 3]]), pages);
  assert.ok(counts.has('U'), 'the count table may know unassigned pages');
  const layout = buildPdfImageLayout(['A', 'B'], counts);
  assert.ok(!layout.sheets.some(s => s.pageId === 'U'), 'unassigned pages never become sheets');
  assert.deepEqual(sheetsOf(layout), ['A:1', 'B:1']);
  assert.deepEqual(buildPdfImageLayout(['', null, undefined], counts).sheets, [], 'empty ids are ignored');
  assert.equal(buildPdfImageLayout(null, null).pageCount, 0);
}

// 見開きの契約（頁・左右はサーバー値）に触れない: PDFイメージの関数は頁・左右を読まず、
// buildSpreadLayout は PDFイメージの頁数表を使わない。
for (const name of ['collectPdfImagePageCounts', 'buildPdfImageLayout']) {
  const body = extractFunction(name);
  for (const banned of ['outputPageNumber', 'side', 'spreadIndex']) {
    assert.ok(!body.includes(banned), `${name} must not depend on ${banned}`);
  }
}
assert.ok(!extractFunction('buildSpreadLayout').includes('PdfImage'), 'buildSpreadLayout must stay on server values');

console.log('pdf image layout regression ok');
