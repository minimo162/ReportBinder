// 未登録原稿の一覧がふつうの名前順に並ぶことの回帰テスト。
//
// サーバー（/api/submission-files）は relativePath の文字列順で返すため、「資料10」が
// 「資料2」より前に来ていた。画面・検索の結果・Shift の範囲選択はどれも
// visibleUnregisteredFiles() の並びを使うので、ここが表示名の名前順（ja・数字は数として・
// 大文字小文字を区別しない）を返すことを、app.js の実物の関数で確かめる。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const appPath = fileURLToPath(new URL('../app/web/app.js', import.meta.url));
const source = readFileSync(appPath, 'utf8');

function extractFunction(name) {
  const start = source.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `function ${name} is missing from app.js`);
  const open = source.indexOf('{', source.indexOf(')', start));
  let depth = 0;
  for (let i = open; i < source.length; i++) {
    if (source[i] === '{') depth++;
    else if (source[i] === '}' && --depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(`unterminated function ${name}`);
}

const collatorLine = source.match(/^const unregisteredFileCollator=.*$/m);
assert.ok(collatorLine, 'unregisteredFileCollator is missing from app.js');

const registered = [{ relativePath: 'registered.xlsx' }];
const factory = new Function('workbooksForActivePreset', 'filesForActivePreset',
  collatorLine[0] + '\n' + ['fileDisplayName', 'compareUnregisteredFiles', 'visibleUnregisteredFiles', 'updateSelectionRange']
    .map(extractFunction).join('\n') + '\nreturn { visibleUnregisteredFiles, updateSelectionRange };');
const { visibleUnregisteredFiles, updateSelectionRange } = factory(() => registered, files => files);

// サーバーの並び（文字列順）を模した入力。
const serverOrder = ['B.xlsx', 'Registered.xlsx', 'a.xlsx', '資料1.xlsx', '資料10.xlsx', '資料2.xlsx', 'あ.docx', 'ア.pdf', 'い.xlsx']
  .map(name => ({ relativePath: name, fileName: name }));
const input = [...serverOrder];
const names = visibleUnregisteredFiles(input).map(f => f.relativePath);

assert.ok(!names.includes('Registered.xlsx'), 'registered files must stay hidden (case-insensitive)');
assert.ok(names.indexOf('資料2.xlsx') < names.indexOf('資料10.xlsx'), `numbers must sort naturally: ${names.join(', ')}`);
assert.ok(names.indexOf('資料1.xlsx') < names.indexOf('資料2.xlsx'), `numbers must sort naturally: ${names.join(', ')}`);
assert.ok(names.indexOf('a.xlsx') < names.indexOf('B.xlsx'), `case must not split the order: ${names.join(', ')}`);
assert.ok(names.indexOf('あ.docx') < names.indexOf('い.xlsx') && names.indexOf('ア.pdf') < names.indexOf('い.xlsx'),
  `kana must sort in Japanese order: ${names.join(', ')}`);
const expected = [...names].sort((x, y) => x.localeCompare(y, 'ja', { numeric: true, sensitivity: 'base' }) || x.localeCompare(y));
assert.deepEqual(names, expected, 'unregistered files must be in display-name order');
assert.deepEqual(input, serverOrder, 'sorting must not reorder availableFiles itself');

// 表示名（fileName）で並べる。relativePath とずれるときも表示名が優先。
const byDisplay = visibleUnregisteredFiles([
  { relativePath: 'z-file.xlsx', fileName: '1.xlsx' },
  { relativePath: 'a-file.xlsx', fileName: '2.xlsx' },
]).map(f => f.relativePath);
assert.deepEqual(byDisplay, ['z-file.xlsx', 'a-file.xlsx'], 'display name decides the order');

// Shift の範囲選択は画面の並び（名前順）で範囲を取る。
const values = visibleUnregisteredFiles(input).map(f => f.relativePath);
const selected = new Set();
const anchor = updateSelectionRange(values, selected, '資料1.xlsx', true, false, '');
updateSelectionRange(values, selected, '資料10.xlsx', true, true, anchor);
assert.deepEqual([...selected].sort(), ['資料1.xlsx', '資料10.xlsx', '資料2.xlsx'].sort(),
  'shift range must follow the on-screen name order');

console.log('unregistered file sort regression passed');
