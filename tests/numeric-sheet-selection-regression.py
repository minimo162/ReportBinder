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


def test_page_mutation_state_sync():
    """Exercise the production mutation freshness hand-off, not a mock copy."""
    root = Path(__file__).resolve().parents[1]
    app = (root / "app/web/app.js").read_text(encoding="utf-8")
    server = (root / "app/server.ps1").read_text(encoding="utf-8-sig")

    mutation = extract_js_function(app, "applyPageMutationResult")
    for needed in (
        "const mutationState=payload?.state",
        "state.packProgress=mutationState.packProgress",
        "state.finalReadiness=Object.assign",
        "advanceFinalReadinessEpoch();",
        "renderGlobalHeader()",
        "renderVolumeLinks()",
        "renderDashboardOverview()",
        "markFinalReadinessStaleForPageMutation",
        "reloadFinalReadinessAfterPageMutation",
    ):
        assert needed in mutation, f"page mutation state sync missing: {needed}"

    # The old compatibility routes are still used by numeric sort and page
    # settings. They must carry the same post-mutation state as the V2 route.
    route_markers = (
        "'/api/pages/reorder'",
        "-in @('/api/pages/sort-by-numeric-sheet','/api/pages/sort-by-sheet')",
        "'/api/pages/update'",
    )
    for route in route_markers:
        route_block = server.split(f"$path -eq {route}", 1)[1].split("\n        if ", 1)[0] if route.startswith("'") else server.split(f"$path {route}", 1)[1].split("\n        if ", 1)[0]
        assert "Get-V2StatePayload $language" in route_block, f"legacy mutation state missing: {route}"
    v2_state = extract_ps_function(server, "Get-V2StatePayload")
    assert "Get-FinalBuildReadiness" in v2_state and "$finalReadiness $packId" in v2_state, "custom pack readiness missing from mutation state"

    # Refresh/conflict/restore must not carry an active custom key through an
    # authoritative state replacement, and restore must consume the response
    # state before the follow-up GET.
    refresh = "async function refresh(options" + app.split("async function refresh(options", 1)[1].split("\nasync function refreshAndLoad", 1)[0]
    layout_conflict = app.split("async function handlePageLayoutConflict", 1)[1].split("\nfunction saveBoardOrder", 1)[0]
    settings_conflict = app.split("async function handlePageSettingsConflict", 1)[1].split("\nfunction pageSettingsBody", 1)[0]
    restore = app.split("async function previewLayoutRestore", 1)[1].split("\nfunction volumeLabel", 1)[0]
    pack_settings = app.split("async function savePackSettings", 1)[1].split("\nfunction requiredRenderProfileVersion", 1)[0]
    assert "delete retained[carriedKey]" in refresh
    assert "authoritativeReadiness" in refresh
    assert "authoritativeReadiness:true" in layout_conflict
    assert "authoritativeReadiness:true" in settings_conflict
    assert "applyPageMutationResult(res)" in restore
    assert "authoritativeReadiness:true" in restore
    assert "await loadFinalReadiness({render:false})" in pack_settings

    stale = extract_js_function(app, "advanceFinalReadinessEpoch") + "\n" + extract_js_function(app, "markFinalReadinessStaleForPageMutation")
    stale_harness = r"""
let finalReadinessEpoch = 0;
let state = {finalReadiness:{pack_custom:{volumes:{
  'ja-main': {pageCount:2,builtFingerprint:'old',outputPdf:'out.pdf',status:'built',displayState:'built',staleReasons:[]}
}}}};
function finalReadinessKey(){ return 'pack_custom'; }
function asArray(value){ return value == null ? [] : (Array.isArray(value) ? value : [value]); }
advanceFinalReadinessEpoch();
markFinalReadinessStaleForPageMutation();
const ready=state.finalReadiness.pack_custom.volumes['ja-main'];
if(ready.displayState!=='needs-rebuild' || ready.status!=='needs-rebuild')throw new Error('reorder left final PDF marked latest');
if(ready.pageCount!==2 || ready.outputPdf!=='out.pdf')throw new Error('stale transition discarded output metadata');
if(finalReadinessEpoch!==1)throw new Error('readiness invalidation epoch did not advance');
"""
    completed = subprocess.run(["node", "-"], input=stale + "\n" + stale_harness, text=True, encoding="utf-8", capture_output=True)
    assert completed.returncode == 0, completed.stdout + completed.stderr

    # Production-coupled race: a readiness GET started before a mutation must
    # not overwrite the mutation's server-provided needs-rebuild state.
    load_readiness = "async function loadFinalReadiness" + app.split("async function loadFinalReadiness", 1)[1].split("\nfunction advanceFinalReadinessEpoch", 1)[0]
    race_functions = "\n".join([
        extract_js_function(app, "finalReadinessKey"), load_readiness,
        extract_js_function(app, "advanceFinalReadinessEpoch"),
        extract_js_function(app, "markFinalReadinessStaleForPageMutation"),
        extract_js_function(app, "applyPageMutationResult"),
    ])
    race_harness = r"""
const finalReadinessInFlight = new Map();
let finalReadinessEpoch = 0, finalPreflightCheckedAt = '';
let lastPageBoardRenderSignature = '';
let activePackId = 'pack_custom', activePreset = 'bod', activeView = 'final';
let activePack = {packId:'pack_custom',category:''};
let state = {structure:{pages:[{pageId:'old'}],volumes:{}},packProgress:{},finalReadiness:{
  pack_custom:{volumes:{'ja-main':{pageCount:2,status:'built',displayState:'built',outputPdf:'out.pdf',builtFingerprint:'old'}}}
}};
let requests=[];
function activePackRecord(){return activePack;}
function configured(){return true;}
function api(url){return new Promise(resolve=>requests.push({url,resolve}));}
function asArray(value){return value==null?[]:(Array.isArray(value)?value:[value]);}
function targetVolume(id){return String(id||'');}
function log(){}
function rememberLayoutFingerprint(){}
function resolvedPageId(page){return String(page?.pageId||'');}
function renderGlobalHeader(){} function renderNavBadges(){} function renderStepBar(){}
function renderSummary(){} function renderDashboardOverview(){} function renderFinalOverview(){}
function renderPages(){} function renderPageOverview(){} function renderVolumeLinks(){}
function isModalOpen(){return false;} function syncPreviewOrganizerControls(){}
const oldGet=loadFinalReadiness({render:false});
if(requests.length!==1||!requests[0].url.includes('pack_custom'))throw new Error('initial readiness GET missing');
applyPageMutationResult({result:{pages:[{pageId:'new'}],affectedVolumes:['ja-main']},state:{
  packProgress:{unassigned:0},finalReadiness:{pack_custom:{volumes:{
    'ja-main':{pageCount:2,status:'needs-rebuild',displayState:'needs-rebuild',outputPdf:'out.pdf',builtFingerprint:'new'}
  }}}
}});
requests[0].resolve({volumes:{'ja-main':{pageCount:2,status:'built',displayState:'built',outputPdf:'out.pdf',builtFingerprint:'old'}}});
await oldGet;
const afterMutation=state.finalReadiness.pack_custom.volumes['ja-main'];
if(afterMutation.displayState!=='needs-rebuild'||afterMutation.status!=='needs-rebuild')throw new Error('old GET restored built state after mutation');
if(state.structure.pages[0].pageId!=='new')throw new Error('mutation result structure not applied');

// Pack-key race: A pending request must not be reused for B, and A's late
// response must not commit while B is active.
requests=[]; state.finalReadiness={}; finalReadinessEpoch=0;
activePackId='pack_a'; activePack={packId:'pack_a',category:''};
const aGet=loadFinalReadiness({render:false});
activePackId='pack_b'; activePack={packId:'pack_b',category:''};
const bGet=loadFinalReadiness({render:false});
if(requests.length!==2||!requests[0].url.includes('pack_a')||!requests[1].url.includes('pack_b'))throw new Error('pack switch reused an in-flight GET');
requests[1].resolve({volumes:{'ja-main':{status:'built',displayState:'built',pageCount:2}}});
await bGet;
requests[0].resolve({volumes:{'ja-main':{status:'built',displayState:'built',pageCount:1}}});
await aGet;
if(!state.finalReadiness.pack_b||state.finalReadiness.pack_a)throw new Error('late A readiness clobbered active B');
"""
    completed = subprocess.run(["node", "--input-type=module", "-"], input=race_functions + "\n" + race_harness, text=True, encoding="utf-8", capture_output=True)
    assert completed.returncode == 0, completed.stdout + completed.stderr

    # Conflict/restore custom-key regression: refresh drops the carried active
    # key and waits for the real readiness endpoint before returning.
    refresh_functions = "\n".join([
        extract_js_function(app, "finalReadinessKey"), load_readiness,
        extract_js_function(app, "advanceFinalReadinessEpoch"), refresh,
    ])
    refresh_harness = r"""
const finalReadinessInFlight = new Map();
let finalReadinessEpoch=0, finalPreflightCheckedAt='', activePackId='pack_custom', activePreset='bod', activeView='pages';
let activePack={packId:'pack_custom',category:''};
let state={finalReadiness:{pack_custom:{volumes:{'ja-main':{status:'built',displayState:'built',pageCount:2}}}}};
function activePackRecord(){return activePack;} function configured(){return true;}
function normalizeStatePayload(payload){return payload;}
function targetVolume(id){return String(id||'');} function log(){} function renderAll(){}
function showMessage(){} const snapshotHistoryResponseCache=new Map(); let historyPanelsInitialized=false;
function api(url){
  if(url==='/api/state')return Promise.resolve({structure:{pages:[]},packProgress:{},finalReadiness:{}});
  return Promise.resolve({volumes:{'ja-main':{status:'needs-rebuild',displayState:'needs-rebuild',pageCount:2}}});
}
await refresh({authoritativeReadiness:true,render:false});
const ready=state.finalReadiness.pack_custom.volumes['ja-main'];
if(ready.displayState!=='needs-rebuild'||ready.status!=='needs-rebuild')throw new Error('refresh carried stale custom readiness');
"""
    completed = subprocess.run(["node", "--input-type=module", "-"], input=refresh_functions + "\n" + refresh_harness, text=True, encoding="utf-8", capture_output=True)
    assert completed.returncode == 0, completed.stdout + completed.stderr


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
    new_excel = extract_ps_function(server, "New-WorkbookObject")
    assert "excelSheetSelectionMode = 'numeric-only'" in new_excel
    assert "lastRenderedSheetSelectionMode = 'numeric-only'" in new_excel
    normalize = extract_ps_function(server, "Normalize-ExcelSheetSelection")
    # Blank/legacy records remain all-visible; only a newly created workbook
    # opts into numeric-only explicitly.
    assert "return 'all-visible'" in normalize
    assert "function Get-NumericSheetSortRank" in server
    assert "function Sort-PagesByNumericDefault" in server
    assert "function Sort-NumericPagesWithinAnchors" in server
    sort_endpoint = extract_ps_function(server, "Sort-PagesByNumericSheet")
    assert "Sort-NumericPagesWithinAnchors" in sort_endpoint
    assert "Get-RequestedBaseLayout" in sort_endpoint
    assert "if (-not $inputChanged) { continue }" in sort_endpoint
    assert sort_endpoint.index("if (-not $inputChanged) { continue }") < sort_endpoint.index("Save-LayoutSnapshot")
    assert "/api/pages/sort-by-numeric-sheet" in server
    insertion = extract_ps_function(server, "Insert-PageInSheetOrder")
    assert "numericExisting" in insertion and "numericSorted" in insertion
    assert "Sort-PagesByNumericDefault $withNew" in insertion
    rank_helper = extract_ps_function(server, "Get-NumericSheetSortRank")
    assert "length = $ordinal.Length" in rank_helper and "ordinal = $ordinal" in rank_helper
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


def test_page_composition_ui_contracts():
    root = Path(__file__).resolve().parents[1]
    app = (root / "app/web/app.js").read_text(encoding="utf-8")
    html = (root / "app/web/index.html").read_text(encoding="utf-8")
    css = (root / "app/web/style.css").read_text(encoding="utf-8")
    server = (root / "app/server.ps1").read_text(encoding="utf-8-sig")
    sort_action = extract_js_function(app, "sortPagesByNumericSheet")
    assert "if(numericBefore)" in sort_action
    assert "hasManualFlags" not in sort_action
    assert "pageVolumeSnapshotsEqual(beforeVolumes,collectBoardVolumes())" in sort_action
    assert "if(pageMutationBusy||pageLayoutHistoryBusy)return" in sort_action
    assert "priorSaveSucceeded=await boardSavePromise" in sort_action
    assert "error.code==='structure-conflict'" in sort_action
    assert "await handlePageLayoutConflict(beforeVolumes,requestRevision)" in sort_action
    assert "finally{pageMutationBusy=false" in sort_action
    drag_start = app.split("function beginPointerPageDrag", 1)[1].split("function setPageThumbnailEditor", 1)[0]
    drag_finish = app.split("function finishPointerDrag", 1)[1].split("function beginPointerPageDrag", 1)[0]
    key_handler = extract_js_function(app, "handleBoardRowKeydown")
    history_action = extract_js_function(app, "applyPageLayoutHistory")
    assert "if (pageMutationBusy || !row" in drag_start
    assert "pointerType" in drag_start and "interactive" in drag_start
    assert "if(pointerType==='touch')return" in drag_start
    assert "moved < 7" in drag_start
    assert "lostpointercapture" in drag_start and "window.addEventListener('blur'" in drag_start
    assert "if (!started)" in drag_start and "handlePageCardSelection(row,ev)" in drag_start
    assert "selectedPages.clear();drag.rows.forEach" in drag_finish
    assert "function handlePageCardSelection" in app
    assert "if(event.ctrlKey||event.metaKey)" in app
    assert "window.getSelection?.()?.removeAllRanges()" in app
    assert "syncPageRovingTabIndex(card)" in app
    assert "const visibleSelection=new Set" in app
    assert "activePageVolume=target;pageVolumeAutoPicked=true" in app
    assert "const focusRow=drag.rows.find" in app
    assert "pageVolumesFromState" in app and "const domLanes" not in app
    assert "renderPageOverview({volumes:afterVolumes})" in app
    assert "textContent=manual?'手動順':'半角数字順'" in app
    assert "sortButton.disabled=pageMutationBusy||!manual" in app
    schedule_action = extract_js_function(app, "scheduleBoardSave")
    save_action = app.split("function saveBoardOrder()", 1)[1].split("\nasync function savePageFromRow", 1)[0]
    assert "renderPageOverview({volumes:afterVolumes})" in schedule_action
    assert "pageMutationBusy=false" in save_action and "renderPageOverview()" in save_action
    assert "if(pageMutationBusy&&e.altKey)" in key_handler
    assert "if(pageLayoutHistoryBusy||pageMutationBusy)return" in history_action
    assert "pageLayoutHistoryBusy=true;pageMutationBusy=true" in history_action
    assert "pageLayoutHistoryBusy=false;pageMutationBusy=false" in history_action
    assert "undo.disabled=pageMutationBusy||pageLayoutHistoryBusy" in app
    assert "redo.disabled=pageMutationBusy||pageLayoutHistoryBusy" in app
    assert "classList.toggle('pages-shell', activeView === 'pages')" in app
    assert html.index('id="sort-by-sheet-btn"') < html.index('class="page-tools-details"')
    assert "main.shell.pages-shell{" in css and "max-width:none" in css
    assert (
        "main.shell.pages-shell .overview-board.thumbnail-board .thumbnail-wrap"
        "{height:auto;max-height:none;overflow:visible}" in css
    )
    assert "grid-template-columns:repeat(2,minmax(0,1fr))" in css
    assert "grid-template-columns:1fr" in css
    assert "grid-template-columns:repeat(3,minmax(0,1fr))" in css
    assert (
        "main.shell.pages-shell .page-command-bar"
        "{flex-wrap:nowrap;overflow-x:auto;overflow-y:hidden" in css
    )
    assert ".thumbnail-board .page-thumb-card{cursor:grab;touch-action:pan-y}" in css
    assert ".thumbnail-board .page-thumb-card{cursor:default;touch-action:pan-y}" in css
    assert ".thumbnail-board .page-thumb-check,.thumbnail-board .page-thumb-drag{display:none!important}" in css
    assert ".page-thumb-preview-surface{pointer-events:none;cursor:default}" in css
    assert "touch-action:pan-y" in css
    assert "<th class=\"seq-col\">順</th><th>ページ名</th>" in app
    assert 'type="checkbox" data-page-check' not in app
    assert "pages.length>=12?'high-volume':''" not in app
    assert ".volume-panel.high-volume" not in css
    assert "grid-template-columns:repeat(auto-fill,minmax(min(100%,232px),1fr))" in css
    assert ".volume-panel:not(.inbox) .thumbnail-grid{grid-template-columns:repeat(2,minmax(0,1fr))}" in css
    assert ".volume-panel.inbox .thumbnail-grid{grid-template-columns:1fr}" in css
    assert ".thumbnail-board .page-thumb-card,.detail-board .page-row{-webkit-user-select:none;user-select:none}" in css
    assert '[contenteditable="true"]){-webkit-user-select:text;user-select:text}' in css
    assert "requestBaseLayout=activeLayoutFingerprint(requestPackId)" in app
    collect_action = extract_js_function(app, "collectNumericSheets")
    assert "latestBaseLayout" not in collect_action
    assert "if(requestBaseLayout)body.baseLayout=requestBaseLayout" in collect_action
    assert "waitForQueuedPageLayoutMutation" in app
    assert "pageMutationBusy=true;updateBulkSelectionLabel()" in app
    assert "function pageChangeBadge" in app
    page_row = app.split("function pageRowHtml", 1)[1].split("\nfunction pageThumbnailHtml", 1)[0]
    page_thumbnail = app.split("function pageThumbnailHtml", 1)[1].split("\nfunction pageThumbnailCacheKey", 1)[0]
    assert "pageChangeBadge(p)" in page_row
    assert "pageChangeBadge(" not in page_thumbnail
    assert "function reviewChip" not in app
    assert "変換PDFの更新が必要" in app
    assert "function deletePack" in app
    assert 'data-pack-action="delete"' in app
    assert "method:'DELETE'" in app
    assert "提出フォルダーへ出力済みのPDFは削除しません" in app

    remove_pack = server.split("function Remove-DocumentPack", 1)[1].split("\nfunction ", 1)[0]
    assert "if (-not (Test-PackArchived $pack))" in remove_pack
    assert "sourceFilesDeleted=$false" in remove_pack
    assert "Set-NoteProperty $structure 'packs'" in remove_pack
    assert "Set-NoteProperty $structure 'workbooks'" in remove_pack
    assert "Set-NoteProperty $structure 'pages'" in remove_pack
    assert "Set-NoteProperty $structure 'outputs'" in remove_pack
    assert "if ($method -eq 'DELETE' -and $path -match '^/api/v2/packs/([^/]+)$')" in server
    pack_dashboard = server.split("function Get-PackProgressDashboard", 1)[1].split("\nfunction ", 1)[0]
    assert "Test-WorkbookRenderIsCurrent" in pack_dashboard


if __name__ == "__main__":
    test_strict_names()
    test_large_numeric_sort()
    test_equal_numeric_value_uses_source_tie_breakers()
    test_mixed_anchors_and_hidden_exclusion()
    test_noop_is_sequence_equality()
    test_unassigned_non_excel_anchor_is_preserved()
    test_production_javascript_helpers()
    test_production_server_contracts()
    test_page_mutation_state_sync()
    test_page_composition_ui_contracts()
    print("numeric-sheet-selection regression ok")
