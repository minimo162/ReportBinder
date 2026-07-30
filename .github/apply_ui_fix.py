from pathlib import Path

ROOT = Path('.')


def read(path, enc='utf-8'):
    return (ROOT / path).read_text(encoding=enc)


def write(path, text, enc='utf-8'):
    (ROOT / path).write_text(text, encoding=enc, newline='')


def replace_once(text, old, new, label):
    count = text.count(old)
    if count != 1:
        raise SystemExit(f'{label}: expected 1 match, got {count}')
    return text.replace(old, new, 1)


# index.html
path = Path('app/web/index.html')
html = read(path)
html = replace_once(
    html,
    '<link rel="stylesheet" href="style.css?v=20260729_v49">',
    '<link rel="stylesheet" href="style.css?v=20260730_v51">',
    'style version',
)
refresh = '<symbol id="i-refresh" viewBox="0 0 16 16"><path d="M13 5.5V2.75l-1.35 1.3A5.5 5.5 0 1 0 13.3 9"/></symbol>'
history_icon = refresh + '\n    <symbol id="i-history" viewBox="0 0 16 16"><path d="M8 2.25a5.75 5.75 0 1 1-5.1 3.1"/><path d="M2.25 2.75v3.5h3.5M8 4.75V8l2.25 1.5"/></symbol>'
html = replace_once(html, refresh, history_icon, 'history icon')
pages = '        <button class="nav-item" type="button" data-view-nav="pages"><svg class="icon"><use href="#i-list"/></svg><span>ページ構成</span><span id="nav-pages-count" class="nav-count hidden"></span></button>'
history = '        <button class="nav-item" type="button" data-view-nav="history"><svg class="icon"><use href="#i-history"/></svg><span>履歴・比較</span></button>'
html = replace_once(html, pages, history + '\n' + pages, 'history navigation')
old_unregistered = '            <section class="card table-card unregistered-card"><div class="card-head"><div><h3>未登録Excel</h3><p id="file-selection-summary" class="caption">0件中 0件を選択</p></div><label class="search-field"><svg class="icon"><use href="#i-search"/></svg><input id="file-filter" type="search" placeholder="ファイル名で検索"></label></div><div id="file-list"></div><div class="card-actions"><div class="toolbar"><button id="select-visible-files-btn" class="btn ghost" type="button">表示分を選択</button><button id="clear-selected-files-btn" class="btn ghost" type="button">選択解除</button></div><button id="register-selected-btn" class="btn primary" type="button">選択したExcelを登録</button></div></section>'
new_unregistered = '''            <section class="card table-card unregistered-card">
              <div class="card-head"><div><h3>未登録Excel</h3><p id="file-selection-summary" class="caption">0件中 0件を選択</p></div><label class="search-field"><svg class="icon"><use href="#i-search"/></svg><input id="file-filter" type="search" placeholder="ファイル名で検索"></label></div>
              <div class="list-action-bar" aria-label="未登録Excelの操作"><div class="toolbar"><button id="select-visible-files-btn" class="btn secondary" type="button">表示分を選択</button><button id="clear-selected-files-btn" class="btn ghost" type="button">選択解除</button></div><button id="register-selected-btn" class="btn primary" type="button">選択したExcelを登録</button></div>
              <div id="file-list"></div>
            </section>'''
html = replace_once(html, old_unregistered, new_unregistered, 'unregistered action bar')
history_start = html.index('          <section class="card"><div class="card-head"><h3>変更履歴</h3>')
developer_log = html.index('        <details class="dev-log">', history_start)
new_history_panel = '''        </section>

        <section data-view-panel="history" class="view-panel">
          <div class="page-head"><div><h2>履歴・比較</h2><p>同じExcelの任意の2時点を選び、左右比較・重ね合わせ・差分強調で確認します。</p></div><button id="history-refresh-btn" class="btn secondary" type="button"><svg class="icon"><use href="#i-refresh"/></svg>履歴を更新</button></div>
          <section class="card history-primary-card"><div class="card-head"><div><span class="eyebrow">主要機能</span><h3>任意の2時点を視覚比較</h3><p class="caption">対象Excelを選び、古い版を「比較元」、新しい版を「比較先」に指定してください。</p></div></div>
            <div class="snapshot-history-controls"><label>対象Excel：<select id="snapshot-history-workbook"></select></label></div>
            <div id="snapshot-history" class="history-table"></div>
            <div id="snapshot-history-diff" class="history-table"></div>
          </section>
          <section class="card"><div class="card-head"><h3>変更履歴</h3></div><div id="history-timeline" class="history-table"></div></section>
          <section class="card"><div class="card-head"><h3>ページ構成の履歴</h3></div><div id="layout-history" class="history-table"></div></section>
          <section class="card"><div class="card-head"><h3>保存済みの出力</h3></div><div id="final-archives" class="history-table"></div></section>
        </section>

'''
html = html[:history_start] + new_history_panel + html[developer_log:]
html = replace_once(
    html,
    '<script src="app.js?v=20260730_v50"></script>',
    '<script src="app.js?v=20260730_v51"></script>',
    'app version',
)
write(path, html)

# app.js
path = Path('app/web/app.js')
app_js = read(path, 'utf-8-sig')
app_js = replace_once(
    app_js,
    "  excel: 'Excel登録・PDF作成',\n  pages: 'ページ構成',",
    "  excel: 'Excel登録・PDF作成',\n  history: '履歴・比較',\n  pages: 'ページ構成',",
    'history view name',
)
app_js = replace_once(
    app_js,
    "  if (!options.noScroll) window.scrollTo({top: 0, behavior: options.instant ? 'auto' : 'smooth'});\n}",
    "  if (!options.noScroll) window.scrollTo({top: 0, behavior: options.instant ? 'auto' : 'smooth'});\n  if (activeView === 'history' && state) loadHistoryPanels();\n}",
    'history view loading',
)
app_js = replace_once(
    app_js,
    "  if(activeView==='final'&&state?.inputHistoryEnabled)loadHistoryPanels();",
    "  if(activeView==='history')loadHistoryPanels();",
    'history refresh routing',
)
write(path, app_js, 'utf-8-sig')

# style.css
path = Path('app/web/style.css')
css = read(path)
css_marker = '/* V5.1 history discoverability and list actions */'
if css_marker in css:
    raise SystemExit('CSS patch was already applied')
css += '''

/* V5.1 history discoverability and list actions */
.list-action-bar{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:10px 14px;border-top:1px solid var(--border);border-bottom:1px solid var(--border);background:var(--surface-subtle)}
.list-action-bar .toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.history-primary-card{border-color:var(--accent-border);box-shadow:0 0 0 1px var(--accent-subtle)}
.history-primary-card>.card-head{background:var(--accent-subtle);border-bottom:1px solid var(--accent-border)}
.history-primary-card .snapshot-history-controls{padding:14px 16px 6px}
.history-primary-card .history-table{padding-left:16px;padding-right:16px}
@media(max-width:900px){.list-action-bar{align-items:stretch;flex-direction:column}.list-action-bar>.btn{width:100%}.history-primary-card .snapshot-history-controls label{display:grid;gap:6px}.history-primary-card select{width:100%}}
'''
write(path, css)

# server.ps1
path = Path('app/server.ps1')
server = read(path, 'utf-8-sig')
output_start = server.index('function Get-OutputFileName(')
output_end = server.index('\nfunction Get-SheetOrderNumber', output_start)
output_functions = '''function Get-CategoryProjectId([string]$ProjectId, [string]$Category) {
    $cat = Require-WorkbookCategory $Category
    $categoryLabel = $cat.ToUpperInvariant()
    $base = ([string]$ProjectId).Trim()
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'ReportBinder' }
    if ($base -match '^(?<prefix>.*?)(?:[_-](?:ECM|BOD|DMM))$') {
        $prefix = (([string]$matches['prefix']) -replace '[_-]+$','')
        if ([string]::IsNullOrWhiteSpace($prefix)) { return $categoryLabel }
        return "${prefix}_${categoryLabel}"
    }
    return "${base}_${categoryLabel}"
}

function Get-OutputFileName([string]$Volume, [string]$ProjectId, [string]$Category) {
    $namedProjectId = Get-CategoryProjectId $ProjectId $Category
    switch ($Volume) {
        'ja-main' { return "${namedProjectId}_J_本体.pdf" }
        'ja-appendix' { return "${namedProjectId}_J_補足.pdf" }
        'en-main' { return "${namedProjectId}_E_Main.pdf" }
        'en-appendix' { return "${namedProjectId}_E_Appendix.pdf" }
        default { throw "未知の成果物: $Volume" }
    }
}
'''
server = server[:output_start] + output_functions + server[output_end:]
server = replace_once(
    server,
    '$outName=Get-OutputFileName $Volume $projectId;',
    '$outName=Get-OutputFileName $Volume $projectId $cat;',
    'legacy category filename call',
)
server = replace_once(
    server,
    '$outName = Get-OutputFileName $v ([string]$snapshots[$v].projectId)',
    '$outName = Get-OutputFileName $v ([string]$snapshots[$v].projectId) $Category',
    'transaction category filename call',
)
write(path, server, 'utf-8-sig')

# selfcheck.py
path = Path('app/tools/selfcheck.py')
selfcheck = read(path)
completion_marker = "print('selfcheck ok')"
regression_checks = '''# V5.1: history comparison must be a first-class destination and list actions must be visible before the list.
for needed in ['data-view-nav="history"', 'data-view-panel="history"', '<h2>履歴・比較</h2>', '任意の2時点を視覚比較', 'class="list-action-bar"']:
    if needed not in html: raise SystemExit(f'history/list action UX missing: {needed}')
if html.index('id="register-selected-btn"') > html.index('id="file-list"'):
    raise SystemExit('unregistered Excel actions must appear before the file list')
if "history: '履歴・比較'" not in appjs or "activeView === 'history'" not in appjs:
    raise SystemExit('history navigation is not wired')

# Final PDF names must use the selected category, not a category inherited from an Excel file name.
for needed in ['function Get-CategoryProjectId', 'Get-OutputFileName $Volume $projectId $cat', 'Get-OutputFileName $v ([string]$snapshots[$v].projectId) $Category']:
    if needed not in server: raise SystemExit(f'category-aware final filename missing: {needed}')
if '$outName=Get-OutputFileName $Volume $projectId;' in server:
    raise SystemExit('legacy final output still omits category')

'''
if selfcheck.count(completion_marker) != 1:
    raise SystemExit('selfcheck completion marker mismatch')
selfcheck = selfcheck.replace(completion_marker, regression_checks + completion_marker, 1)
write(path, selfcheck)

# CHANGELOG_V5.md
path = Path('CHANGELOG_V5.md')
changelog = read(path, 'utf-8-sig')
changelog_heading = '# ReportBinder V5 実装履歴\n'
changelog_entry = '''
## UI導線・カテゴリ別出力名の修正（2026-07-30）

- 任意2時点の視覚比較を「最終PDF」画面下部から独立した「履歴・比較」メニューへ移し、主要機能として先頭に表示した。
- 未登録Excelの「表示分を選択」「選択解除」「選択したExcelを登録」を一覧下部から一覧上部の操作バーへ移した。
- BOD/DMMの最終PDF名が元Excel名のECM表記を引き継ぐ問題を修正し、選択中カテゴリをファイル名へ必ず反映するようにした。
- UI配置とカテゴリ別ファイル名の回帰検査をselfcheckへ追加した。
'''
changelog = replace_once(
    changelog,
    changelog_heading,
    changelog_heading + changelog_entry,
    'changelog heading',
)
write(path, changelog, 'utf-8-sig')
