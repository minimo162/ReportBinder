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
  const open = source.indexOf('{', start);
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

console.log('page-board lanes regression ok');
