"""Pure regression checks for the numeric-sheet organizer contract.

The browser implementation is deliberately DOM-bound, so these checks model the
same page partition/slot rules without starting the app or touching user data.
They protect the high-risk ordering cases (anchors, hidden sheets, and very
large digit strings) while app/tools/selfcheck.py protects the server/UI wiring.
"""

from functools import cmp_to_key
from pathlib import Path
import re
import subprocess


STRICT = re.compile(r"^[0-9]+$")


def strict_numeric(name: str) -> bool:
    return bool(name) and "\n" not in name and "\r" not in name and bool(STRICT.fullmatch(name))


def numeric_key(page):
    raw = page["sheet"]
    canonical = raw.lstrip("0") or "0"
    return canonical


def compare_numeric(left, right):
    left_key, right_key = numeric_key(left), numeric_key(right)
    if len(left_key) != len(right_key):
        return -1 if len(left_key) < len(right_key) else 1
    if left_key != right_key:
        return -1 if left_key < right_key else 1
    left_file_order, right_file_order = left.get("file_order", 999999), right.get("file_order", 999999)
    if left_file_order != right_file_order:
        return -1 if left_file_order < right_file_order else 1
    left_file, right_file = left.get("file", "").casefold(), right.get("file", "").casefold()
    if left_file != right_file:
        return -1 if left_file < right_file else 1
    left_index, right_index = left.get("sheet_index", 999999), right.get("sheet_index", 999999)
    if left_index != right_index:
        return -1 if left_index < right_index else 1
    return (left["id"] > right["id"]) - (left["id"] < right["id"])


def organize(pages, target):
    volumes = {}
    for page in pages:
        volumes.setdefault(page["volume"], []).append(page)
    numeric, excluded = [], []
    for page in pages:
        if page["type"] != "excel":
            continue
        if page.get("hidden") or not strict_numeric(page["sheet"]):
            excluded.append(page)
        else:
            numeric.append(page)
    numeric.sort(key=cmp_to_key(compare_numeric))
    def fill_excel_slots(current, replacements):
        slots = [i for i, page in enumerate(current) if page["type"] == "excel"]
        last_slot = slots[-1] if slots else -1
        result, replacement_index = [], 0
        for index, page in enumerate(current):
            if page["type"] == "excel":
                if replacement_index < len(replacements):
                    result.append(replacements[replacement_index])
                    replacement_index += 1
                if index == last_slot:
                    result.extend(replacements[replacement_index:])
                    replacement_index = len(replacements)
            else:
                result.append(page)
        if last_slot < 0:
            result.extend(replacements)
        return result

    target_result = fill_excel_slots(volumes.get(target, []), numeric)
    result = {volume: [] for volume in set(volumes) | {target, "none"}}
    result[target] = target_result
    for volume, current in volumes.items():
        if volume == target:
            continue
        result[volume] = [page for page in current if page["type"] != "excel"]
    result["none"] = fill_excel_slots(volumes.get("none", []), excluded)
    return {volume: [page["id"] for page in current] for volume, current in result.items()}


def test_strict_names():
    assert all(strict_numeric(value) for value in ("0", "1", "0002", "1234567890123456789012345678901"))
    assert not any(strict_numeric(value) for value in ("", " 1", "1 ", "１", "1\n", "1a"))


def test_large_numeric_sort():
    pages = [
        {"id": "long", "sheet": "1000000000000000000000000000001"},
        {"id": "two", "sheet": "2"},
        {"id": "ten", "sheet": "10"},
    ]
    assert [page["id"] for page in sorted(pages, key=cmp_to_key(compare_numeric))] == ["two", "ten", "long"]


def test_equal_numeric_value_uses_source_tie_breakers():
    pages = [
        {"id": "leading", "sheet": "01", "file_order": 20, "file": "b.xlsx", "sheet_index": 1},
        {"id": "plain", "sheet": "1", "file_order": 10, "file": "a.xlsx", "sheet_index": 9},
    ]
    assert [page["id"] for page in sorted(pages, key=cmp_to_key(compare_numeric))] == ["plain", "leading"]


def test_mixed_anchors_and_hidden_exclusion():
    pages = [
        {"id": "doc-a", "type": "word", "sheet": "", "volume": "main"},
        {"id": "memo", "type": "excel", "sheet": "メモ", "volume": "main"},
        {"id": "doc-b", "type": "pdf", "sheet": "", "volume": "main"},
        {"id": "ten", "type": "excel", "sheet": "10", "volume": "none"},
        {"id": "two", "type": "excel", "sheet": "2", "volume": "appendix"},
        {"id": "hidden", "type": "excel", "sheet": "1", "hidden": True, "volume": "appendix"},
    ]
    actual = organize(pages, "main")
    # The one Excel slot in main is filled first, then the extra numeric page is
    # inserted immediately after that slot; the Word/PDF anchors stay put.
    assert actual["main"] == ["doc-a", "two", "ten", "doc-b"]
    assert actual["appendix"] == []
    assert actual["none"] == ["memo", "hidden"]


def test_noop_is_sequence_equality():
    pages = [
        {"id": "one", "type": "excel", "sheet": "1", "volume": "main"},
        {"id": "two", "type": "excel", "sheet": "2", "volume": "main"},
        {"id": "doc", "type": "word", "sheet": "", "volume": "main"},
    ]
    result = organize(pages, "main")
    assert result["main"] == ["one", "two", "doc"]


def test_unassigned_non_excel_anchor_is_preserved():
    pages = [
        {"id": "memo", "type": "excel", "sheet": "メモ", "volume": "none"},
        {"id": "cover", "type": "pdf", "sheet": "", "volume": "none"},
        {"id": "hidden", "type": "excel", "sheet": "3", "hidden": True, "volume": "none"},
        {"id": "one", "type": "excel", "sheet": "1", "volume": "main"},
    ]
    result = organize(pages, "main")
    assert result["none"] == ["memo", "cover", "hidden"]


def extract_js_function(source: str, name: str) -> str:
    marker = f"function {name}("
    start = source.index(marker)
    brace = source.index("{", start)
    depth, quote, escaped = 0, None, False
    for index in range(brace, len(source)):
        char = source[index]
        if quote:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == quote:
                quote = None
            continue
        if char in ("'", '"', "`"):
            quote = char
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return source[start:index + 1]
    raise AssertionError(f"unterminated JavaScript function: {name}")


def extract_ps_function(source: str, name: str) -> str:
    marker = f"function {name}"
    start = source.index(marker)
    match = re.search(r"\nfunction [A-Za-z0-9_-]+", source[start + len(marker):])
    return source[start:] if match is None else source[start:start + len(marker) + match.start()]


def test_production_javascript_helpers():
    root = Path(__file__).resolve().parents[1]
    app = (root / "app/web/app.js").read_text(encoding="utf-8")
    functions = "\n".join(extract_js_function(app, name) for name in (
        "strictNumericSheetName", "numericSheetPageInfo", "compareAsciiNumericText",
        "compareNumericSheetPages", "fillExcelPageSlots",
    ))
    harness = r"""
const workbooks = new Map([
  ['plain', {sourceType:'excel', fileName:'_10_a.xlsx'}],
  ['leading', {sourceType:'excel', fileName:'_20_b.xlsx'}],
  ['file-a', {sourceType:'excel', fileName:'_30_a.xlsx'}],
  ['file-b', {sourceType:'excel', fileName:'_30_b.xlsx'}],
  ['same', {sourceType:'excel', fileName:'_40_same.xlsx'}],
]);
function getWorkbook(id){ return workbooks.get(String(id)) || {sourceType:'excel',fileName:'_99_z.xlsx'}; }
function sourceTypeValue(w){ return String(w?.sourceType || 'excel'); }
function fileOrderValue(name){ const m=String(name||'').match(/_(\d{1,4})_/);return m?Number(m[1]):Number.MAX_SAFE_INTEGER; }
function resolvedPageId(p){ return String(p?.pageId || ''); }
function asArray(v){ return Array.isArray(v)?v:(v==null?[]:[v]); }
function check(ok,message){ if(!ok)throw new Error(message); }
check(strictNumericSheetName('0') && strictNumericSheetName('1234567890123456789012345678901'),'ASCII digits rejected');
for(const value of ['', ' 1', '1 ', '１', '-1', '1.0', '1\n'])check(!strictNumericSheetName(value),'invalid accepted: '+JSON.stringify(value));
const tied=[
  {pageId:'leading',workbookId:'leading',sheetName:'01',sheetIndex:1},
  {pageId:'plain',workbookId:'plain',sheetName:'1',sheetIndex:9},
].sort(compareNumericSheetPages);
check(tied.map(x=>x.pageId).join(',')==='plain,leading','production tie order changed');
const byFile=[
  {pageId:'file-b',workbookId:'file-b',sheetName:'1',sheetIndex:1},
  {pageId:'file-a',workbookId:'file-a',sheetName:'01',sheetIndex:9},
].sort(compareNumericSheetPages);
check(byFile.map(x=>x.pageId).join(',')==='file-a,file-b','production filename tie order changed');
const byIndex=[
  {pageId:'index-9',workbookId:'same',sheetName:'1',sheetIndex:9},
  {pageId:'index-1',workbookId:'same',sheetName:'01',sheetIndex:1},
].sort(compareNumericSheetPages);
check(byIndex.map(x=>x.pageId).join(',')==='index-1,index-9','production sheet-index tie order changed');
const byId=[
  {pageId:'z',workbookId:'same',sheetName:'1',sheetIndex:1},
  {pageId:'a',workbookId:'same',sheetName:'01',sheetIndex:1},
].sort(compareNumericSheetPages);
check(byId.map(x=>x.pageId).join(',')==='a,z','production page-id tie order changed');
const huge=[
  {pageId:'huge',sheetName:'1000000000000000000000000000001'},
  {pageId:'two',sheetName:'2'}, {pageId:'ten',sheetName:'10'},
].sort(compareNumericSheetPages);
check(huge.map(x=>x.pageId).join(',')==='two,ten,huge','production large-number order changed');
const slots=[
  {pageId:'memo',sheetName:'memo'},
  {pageId:'cover',workbookId:'pdf',sheetName:''},
  {pageId:'hidden',sheetName:'3',sheetHidden:true},
];
workbooks.set('pdf',{sourceType:'pdf',fileName:'cover.pdf'});
const replacements=[slots[0],slots[2]];
check(fillExcelPageSlots(slots,replacements).map(x=>x.pageId).join(',')==='memo,cover,hidden','none anchor moved');
"""
    completed = subprocess.run(["node", "-"], input=functions + "\n" + harness, text=True, encoding="utf-8", capture_output=True)
    assert completed.returncode == 0, completed.stdout + completed.stderr


def test_production_server_contracts():
    root = Path(__file__).resolve().parents[1]
    server = (root / "app/server.ps1").read_text(encoding="utf-8-sig")
    excluded = extract_ps_function(server, "Set-ExcludedSheetPagesNotRendered")
    for statement in (
        "Set-NoteProperty $page 'volume' 'none'",
        "Set-NoteProperty $page 'enabled' $false",
        "Set-NoteProperty $page 'contentPdf' $null",
        "Set-NoteProperty $page 'status' 'not-rendered'",
        "Set-NoteProperty $page 'sheetSelectionExcluded' $true",
    ):
        assert statement in excluded, f"production exclusion contract missing: {statement}"
    comparison = extract_ps_function(server, "Render-SnapshotForComparison")
    assert "Get-SnapshotRenderSheetSelection" in comparison
    assert "$BaselineVersionId" in comparison
    assert "$comparisonSheetSelection -eq 'numeric-only'" in comparison
    assert "Test-StrictNumericSheetName $sheetName" in comparison
    resolver = extract_ps_function(server, "Get-SnapshotRenderSheetSelection")
    assert "Get-RenderRecordDir $Language $WorkbookId $SnapshotId $VersionId" in resolver
    compare_visual = extract_ps_function(server, "Compare-SnapshotVisual")
    assert "Render-SnapshotForComparison $Language $WorkbookId $baseSnap $baseVer" in compare_visual
    render = extract_ps_function(server, "Render-Workbook")
    zero = render.index("if ($targetSheetNames.Count -eq 0)")
    export = render.index("Export-WorkbookSheetsToPdfBatch")
    assert zero < export and "Remove-Item -LiteralPath $contentDir" in render[zero:export]
    locked = render.index("$pageSync = Update-StructureLocked $Language")
    assert "Save-LayoutSnapshot" not in render[render.index("$usedBatchOutput"):locked]
    assert "Save-LayoutSnapshot" in render[locked:]
    assert "selectionAffectedVolumesLocked" in render[locked:]


if __name__ == "__main__":
    test_strict_names()
    test_large_numeric_sort()
    test_equal_numeric_value_uses_source_tie_breakers()
    test_mixed_anchors_and_hidden_exclusion()
    test_noop_is_sequence_equality()
    test_unassigned_non_excel_anchor_is_preserved()
    test_production_javascript_helpers()
    test_production_server_contracts()
    print("numeric-sheet-selection regression ok")
