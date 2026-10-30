// 変換PDFの文書キャッシュの回帰テスト。
//
// サムネイル・見開きは描くたびに変換PDFを getDocument で開き、描き終えると destroy していた。
// 見開きで5ページのシートを出すと同じ変換PDFを5回丸ごと取得し、1リクエストずつのサーバーで
// readiness や保存の応答まで待たせていた。
//
// app.js の実物の関数を取り出し、偽の getDocument で取得回数と destroy を数えて、
// 同じ版の文書は1回だけ取得し、使用中の文書は追い出さないことを確かめる。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const appPath = fileURLToPath(new URL('../app/web/app.js', import.meta.url));
const source = readFileSync(appPath, 'utf8');

function extractFunction(name) {
  const start = source.search(new RegExp(`(?:async )?function ${name}\\(`));
  assert.ok(start >= 0, `function ${name} is missing from app.js`);
  const open = source.indexOf('{', source.indexOf(')', start));
  let depth = 0;
  for (let i = open; i < source.length; i++) {
    if (source[i] === '{') depth++;
    else if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(`unterminated function ${name}`);
}

function setup({ limit = 8 } = {}) {
  const calls = [];
  const destroyed = [];
  const failing = new Set();
  const lib = {
    getDocument(params) {
      const url = params.url;
      calls.push(url);
      if (failing.has(url)) return { promise: Promise.reject(new Error('fetch failed')) };
      const doc = {
        url, numPages: 5, destroyed: false,
        destroy() { this.destroyed = true; destroyed.push(url); return Promise.resolve(); },
        async getPage(n) {
          return {
            n,
            getViewport: ({ scale }) => ({ width: 100 * scale, height: 141 * scale }),
            render: () => ({ promise: Promise.resolve() }),
          };
        },
      };
      return { promise: Promise.resolve(doc) };
    },
  };
  const workbooks = { wb: { lastRenderedSnapshotId: 's1', lastRenderedVersionId: 'v1' } };
  const fakeCanvas = () => ({ width: 0, height: 0, getContext: () => ({ fillRect() {} }) });
  const factory = new Function('lib', 'workbooks', 'fakeCanvas', 'LIMIT', `
    const window={devicePixelRatio:1};
    const document={createElement:()=>fakeCanvas()};
    const token='t';
    const ensureDiffPdfJs=async()=>lib;
    const apiUrl=(path,params)=>path+'?'+new URLSearchParams(params).toString();
    const getWorkbook=id=>workbooks[id];
    const resolvedPageId=page=>String(page?.pageId||'');
    const pageThumbnailCache=new Map();
    const PAGE_THUMBNAIL_CACHE_LIMIT=64;
    const pageDocumentCache=new Map();
    const PAGE_DOCUMENT_CACHE_LIMIT=LIMIT;
    let pageDocumentCacheClock=0;
    const pdfImageObservedPageCounts=new Map();
    ${extractFunction('pageDocumentCacheKey')}
    ${extractFunction('pageThumbnailCacheKey')}
    ${extractFunction('acquirePageDocument')}
    ${extractFunction('evictPageDocuments')}
    ${extractFunction('clearPageDocumentCache')}
    ${extractFunction('buildPageThumbnail')}
    return { pageDocumentCacheKey, pageThumbnailCacheKey, acquirePageDocument, clearPageDocumentCache, buildPageThumbnail, pageDocumentCache, pageThumbnailCache, pdfImageObservedPageCounts };`);
  const api = factory(lib, workbooks, fakeCanvas, limit);
  const urlOf = page => `/api/file?${new URLSearchParams({ pageId: page.pageId, workbookId: page.workbookId || '', sheetName: page.sheetName || '' })}`;
  return { ...api, calls, destroyed, failing, workbooks, urlOf };
}

const tick = () => new Promise(resolve => setTimeout(resolve, 0));
const page = id => ({ pageId: id, workbookId: 'wb', sheetName: id, contentPdf: `${id}.pdf` });

// 同じ版の変換PDFを5ページ描いても、getDocument は1回。描き終えても文書は破棄しない。
{
  const t = setup();
  const p = page('A');
  await Promise.all([1, 2, 3, 4, 5].map(n => t.buildPageThumbnail(p, n)));
  assert.equal(t.calls.length, 1, 'one sheet must be fetched once for all of its pages');
  assert.equal(t.destroyed.length, 0, 'a cached document must not be destroyed after each render');
  assert.equal(t.pageDocumentCache.get(t.pageDocumentCacheKey(p)).refs, 0, 'every render must release its document');
  assert.equal(t.pdfImageObservedPageCounts.get('A'), 5, 'the opened page count must be recorded');
  // 描画結果の canvas キャッシュから追い出された後でも、文書は取り直さない。
  t.pageThumbnailCache.clear();
  await t.buildPageThumbnail(p, 2);
  assert.equal(t.calls.length, 1, 'a page dropped from the canvas cache must reuse the cached document');
}

// サムネイルのキーは文書のキーにページ番号を足したもの。版が変われば別の文書になる。
{
  const t = setup();
  const p = page('A');
  assert.equal(t.pageThumbnailCacheKey(p, 3), `${t.pageDocumentCacheKey(p)}|p3`);
  assert.equal(t.pageThumbnailCacheKey(p), t.pageThumbnailCacheKey(p, 1));
  const before = t.pageDocumentCacheKey(p);
  const first = await t.acquirePageDocument(p);
  first.release();
  t.workbooks.wb.lastRenderedVersionId = 'v2';
  assert.notEqual(t.pageDocumentCacheKey(p), before, 'a re-rendered version must use a different key');
  const second = await t.acquirePageDocument(p);
  second.release();
  assert.equal(t.calls.length, 2, 'a new version must be fetched again');
  assert.notEqual(first.doc, second.doc);
}

// 上限を超えたら、使っていない文書を古い順に破棄する。使用中の文書は追い出さない。
{
  const t = setup({ limit: 2 });
  const a = await t.acquirePageDocument(page('A')); // 使用中のまま
  const b = await t.acquirePageDocument(page('B')); b.release();
  const c = await t.acquirePageDocument(page('C')); c.release();
  await tick();
  assert.equal(a.doc.destroyed, false, 'a document in use must never be evicted');
  assert.equal(b.doc.destroyed, true, 'the oldest unused document must be destroyed');
  assert.equal(c.doc.destroyed, false);
  assert.deepEqual([...t.pageDocumentCache.keys()].map(key => key.split('|')[0]).sort(), ['A', 'C']);
  // 全部使用中なら上限を一時的に超えてよい。返したら上限まで戻す。
  const d = await t.acquirePageDocument(page('D'));
  const e = await t.acquirePageDocument(page('E'));
  await tick();
  assert.equal(t.pageDocumentCache.size >= 2, true);
  for (const lease of [a, d, e]) lease.release();
  await tick();
  assert.equal(t.pageDocumentCache.size, 2);
  assert.equal(a.doc.destroyed, true, 'the least recently used document goes first once released');
  // 二重に返しても refs は負にならない。
  e.release(); e.release();
  assert.equal(t.pageDocumentCache.get(t.pageDocumentCacheKey(page('E'))).refs, 0);
}

// 取得に失敗した文書はキャッシュに残さず、次は取り直す。
{
  const t = setup();
  const p = page('X');
  t.failing.add(t.urlOf(p));
  await assert.rejects(t.acquirePageDocument(p));
  await tick();
  assert.equal(t.pageDocumentCache.has(t.pageDocumentCacheKey(p)), false, 'a failed document must not stay cached');
  t.failing.clear();
  const lease = await t.acquirePageDocument(p);
  lease.release();
  assert.equal(t.calls.length, 2, 'the next request must fetch again after a failure');
}

// 全部手放すとき、使用中の文書は返したときに破棄する。
{
  const t = setup();
  const busy = await t.acquirePageDocument(page('A'));
  const idle = await t.acquirePageDocument(page('B')); idle.release();
  t.clearPageDocumentCache();
  await tick();
  assert.equal(t.pageDocumentCache.size, 0);
  assert.equal(idle.doc.destroyed, true);
  assert.equal(busy.doc.destroyed, false, 'a document still rendering must not be destroyed under it');
  busy.release();
  await tick();
  assert.equal(busy.doc.destroyed, true);
}

// 画面外へ出た未着手のサムネイルはキューから外す。描画中・完了済みの扱いは変えない。
{
  const queue = [];
  const { releasePageThumbnailCanvas } = new Function('pageThumbnailRenderQueue', `
    ${extractFunction('dequeuePageThumbnail')}
    ${extractFunction('releasePageThumbnailCanvas')}
    return { releasePageThumbnailCanvas };`)(queue);
  const canvas = state => ({ dataset: state ? { thumbnailState: state } : {}, width: 10, height: 10, parentElement: null });
  const pending = canvas(), loading = canvas('loading'), ready = canvas('ready');
  queue.push(pending, loading);
  releasePageThumbnailCanvas(pending);
  releasePageThumbnailCanvas(loading);
  releasePageThumbnailCanvas(ready);
  assert.deepEqual(queue, [loading], 'only a not-yet-started canvas leaves the queue');
  assert.equal(loading.dataset.thumbnailState, 'loading');
  assert.equal(ready.width, 0);
  assert.equal(ready.dataset.thumbnailState, undefined);
}

console.log('page document cache regression ok');
