# ReportBinder（ECM資料PDF作成ツール）UI/UX改修 指示書 V4

対象バージョン: `app/web/index.html` `app/web/app.js` `app/web/style.css` `app/server.ps1`（2026-07-23 時点）
作成日: 2026-07-28

## 改訂履歴

| 版 | 変更内容 |
|---|---|
| V2 | 初版 |
| V3 | レビュー指摘を反映。**V2の誤りを訂正**：①volumesは「12キー」ではなく**言語別6キー**（`structure.json` が `dataDir\ja` `dataDir\en` に分かれているため）②A-5の「他言語まで波及する」は誤り（`PSObject.Properties.Name -contains` のガードで実害なし）。実際の欠陥は同一言語の本体・補足を無差別に汚すこと ③C-2の「手動順を維持したまま数字順の途中に挿入」は論理的に両立しないため方針を確定。**追加**：Phase 0（selfcheck更新・structure更新トランザクション・配布物整理）、最終PDFを開くAPIのカテゴリ対応、ビルド成功時のstaleReasonsクリア、readiness検証の共通化 |
| V4 | V3再レビューを反映。**必須修正を5点に拡張**：①manifest／一時PDF名をカテゴリ込みの決定的名称へ変更 ②組版入力fingerprintと`builtFingerprint`で組版中変更・stalenessを検知 ③`not-built / needs-rebuild / output-missing / built`の状態遷移と表示を一対一化 ④最終PDF系APIのcategoryをfail-closedで必須化 ⑤schema移行をstructureロック内へ隔離し、移行前バックアップ・ロック取得順・旧版停止手順を追加。あわせて、直接配置する「すべて再作成」ボタン、ステータスバッジの淡色tint、ローカル設定優先方式、JRE配布方式を確定 |

---

## 0. 改修のゴール

| # | ゴール | 判定基準 |
|---|---|---|
| G1 | 「今、最終PDFは最新か」が常に画面から分かる | Excel更新・PDF作成・ページ構成変更のいずれの後でも、最終PDF画面まで行かずに「再出力が必要」が分かる |
| G2 | 4ステップの導線が止まらない | 登録→PDF作成→ページ構成→最終PDF を、サイドメニューを自分で押し直さずに完走できる |
| G3 | ページ順を安全にリセットできる | ドラッグで崩した順序を「シート名の半角数字順」に1クリックで戻せる |
| G4 | カテゴリ（ECM/BOD/DMM）の状態が混ざらない | ECMで出力した結果がBODの画面に出ない |
| G5 | Apple系の余白・階層 × Linear系の1pxミニマルに刷新 | 装飾色ゼロ、アクセントはindigo1色、影はモーダルのみ |

---

## 1. レビュー結果サマリ

重大度: ★★★=業務上の誤操作・誤解を生む / ★★=作業効率を明確に落とす / ★=品質・信頼感

| ID | 重大度 | 症状 | 根拠（現行コード） |
|---|---|---|---|
| A-1 | ★★★ | **PDF作成しても最終PDFが「出力済み」のまま**。再出力が必要なことが画面に出ない | `Render-Workbook`（server.ps1 L1647〜）は `volumes.*.status` を一切更新しない。`Mark-StructureVolumesNeedRebuild` はシート増減時（L1055）しか呼ばれない |
| A-2 | ★★★ | 30秒スキャンで元Excelの更新を検知しても、最終PDFの状態は変わらない | `Scan-Updates` は workbook の status のみ更新 |
| A-3 | ★★★ | 最終PDF画面に「最新/再出力が必要/未出力」の表示が無い | `renderFinalOverview()`（app.js L735）はページ数と履歴しか描画しない |
| A-4 | ★★ | `needs-rebuild` になると「出力済みPDFを開く」リンクが**消える**。ファイルは存在するのに開けなくなる | `renderVolumeLinks()`（app.js L771）が `status === 'built'` のときだけ表示 |
| A-5 | ★★ | **本体だけを触っても補足まで「要再出力」になる**。タイトル文言を直しただけでも両方が汚れるため、警告が常時点灯して意味を失う | `Reorder-Pages` / `Update-Page` / `Unregister-Workbook` が固定配列を走査して同一言語のmain/appendixを無差別更新する。V2の「他言語へ波及」は誤り |
| A-6 | ★★ | 古いcontent-pdfのまま出力しようとすると、**出力ボタンを押して初めて例外**で止まる | `Build-FinalPdf` の検証がthrow中心で、事前readiness APIがない |
| A-7 | ★★★ | **Java組版中にページ順・numbering・content-pdfが変わっても、旧入力で作ったPDFを`built`にできる** | `Build-FinalPdf` は開始時にstructureを読み、Java終了後に入力の再照合をせず`status='built'`へ更新する |
| A-8 | ★★★ | イベント駆動の`Mark-VolumeNeedsRebuild`呼び出しを1箇所でも忘れると、変更済みなのに「最新」のまま残る | freshnessの正否をイベントの呼び出し網だけに依存している。今後API追加時の漏れを検知できない |
| B-1 | ★★★ | **volume状態がカテゴリ別でない**。ECMで本体PDFを出力すると、BOD/DMMタブでも出力済みと表示される | `volumes` のキーは言語ごとにmain/appendixの2つのみ。一方、ページ抽出はカテゴリで絞られる |
| B-2 | ★★ | 「先頭ページはページ番号なし」の既定がカテゴリ横断で計算され、2番目以降のカテゴリの先頭ページに番号が付く | `Apply-DefaultNumberingPerVolume` はvolumeだけでフィルタしcategoryを見ない |
| B-3 | ★★★ | **前回出力PDFを開く処理がカテゴリを見ない**。BOD画面からECMのPDFを開き得る | `Serve-FinalPdfByVolume`・GET/POST `/api/final/file`・`openFinalVolume()`がvolumeしか送受信しない |
| B-4 | ★★★ | **category省略時に全カテゴリが対象になるfail-open**。ECM/BOD/DMMのページが混在した最終PDFを生成し得る | `Test-WorkbookCategory` はcategoryが空なら`$true`、`Build-FinalPdf`も未指定時は全workbookを対象にする |
| B-5 | ★★★ | カテゴリ別volumeロックへ変更しただけでは、ECM/BOD同時組版時にmanifestが衝突する | 現行manifest名は `exports\manifest_$Volume.json` でcategoryを含まない |
| C-1 | ★★★ | 一度ドラッグすると、シート名の数字順に戻す手段がない | `Reorder-Pages` が対象volume内の全ページを`orderManual=$true`にする |
| C-2 | ★★★ | 手動順と自動順でorderの採番スケールが混在し、新規ページが常に末尾に付く | 手動は10刻み、自動はシート番号×100000 |
| D-1 | ★★ | Excel登録・ページ構成・最終PDF画面からカテゴリを切り替えられない | `data-preset`は一部画面だけに存在 |
| D-2 | ★★ | 各ステップ完了後に次へ進む導線がない | 登録後などに文言だけ表示し、次アクションを実行できない |
| D-3 | ★★ | ダッシュボードの「次にやること」が最終PDFの鮮度・出力ファイル実在を見ない | `renderDashboardOverview()`は`needsPdf`とページ数中心 |
| D-4 | ★★ | 30秒ポーリングのDOM再構築でページ名入力中のフォーカスとキャレットが飛ぶ | `renderPages()`が`innerHTML`を差し替える |
| D-5 | ★ | 通知が自動で消えず、同じエラーが通知とエラーパネルの2箇所に出る | `showMessage()`の整理不足 |
| D-6 | ★ | サイドメニューに未処理件数がなく、どのステップに用事があるか分からない | `.nav-item`に状態表示がない |
| E-1 | ★★ | 押しても何も起きない、または操作可能に見えるUIが複数ある | 対応イベントがないページャ、ヒント、行末メニュー等 |
| E-2 | ★ | `bind()`対象の死にIDがあり、有用な一括処理もUIから到達不能 | `render-all-btn`、`build-all-btn`等 |
| E-3 | ★ | 最近使用した提出フォルダと出力者列が実データを表していない | 現在値1件、固定値`ReportBinder` |
| F-1 | ★ | 記号アイコンがフォント依存で環境により崩れる | index.html全域 |
| F-2 | ★ | 太字・大角丸・色付き影・グラデーションがミニマル方針と逆 | style.css全域 |
| G-1 | ★★★ | **structure.jsonの読込→変更→保存が一体でロックされず、別プロセスの更新を失う** | `Save-Structure`は書込のみを直列化し、読込はロック外 |
| G-2 | ★★ | `selfcheck.py`が現行の印刷設定とUI方針に追随していない | 1.0cm / 水平中央false / 0.3cm、`render-all-btn`禁止を期待 |
| G-3 | ★ | 配布ZIPに実運用config・ログ・キャッシュ等が残る | `app/config.json`、`app/logs`、`thirdparty-cache`等 |
| G-4 | ★★★ | **schema移行を通常の読み取り処理で行うと、複数PCの同時GETで移行書込みが競合する** | 現行`Get-Structure`自身が不足プロパティをrepairして書き戻す構造 |
| G-5 | ★★★ | schemaVersion 2移行後に旧版プロセスが同じdataDirへ書くと、新旧構造が混在・上書きされ得る | 共有フォルダ上で複数利用者が起動し、既に動作中の旧プロセスは新コードのversion guardを認識しない |

## 2. A. 最終PDFの鮮度を可視化する

### A-1. freshnessの正否はfingerprintを主判定にする

V3の`staleReasons`イベントだけで状態を決める方式を改め、**最終PDFを作成したときの入力fingerprintと現在の入力fingerprintの比較をfreshnessの主判定**とする。

`structure.volumes.<key>`は次の形とする。

```jsonc
{
  "status": "built | needs-rebuild | not-built",
  "lastBuiltAt": "2026-07-28T10:12:00+09:00",
  "outputPdf": "...\XXXX_J_本体.pdf",
  "builtFingerprint": "sha256:...",       // 最終PDFへ採用した入力のfingerprint
  "staleReasons": [                       // 人間向けの補助情報。最大10件、新しい順
    { "type": "render", "at": "...", "detail": "3件のExcelをPDF作成しました" },
    { "type": "reorder", "at": "...", "detail": "ページ構成を変更しました" }
  ]
}
```

- `outputPdfExists`は永続化しない。`/api/state`と`/api/final/readiness`で`Test-Path`により毎回算出する。
- `builtFingerprint`と現在fingerprintが異なれば、イベント呼び出し漏れがあっても`needs-rebuild`と判定する。
- `staleReasons`は状態の唯一の根拠にしない。変更操作時の分かりやすい理由表示に使い、理由が取得できないfingerprint差分には`最終PDFの入力が変更されました`を補う。
- 初回出力前は`status='not-built'`、`builtFingerprint=''`、`staleReasons=@()`とする。「再出力が必要」とは表示しない。

**状態遷移**

| 条件 | 内部状態・表示 |
|---|---|
| 対象ページがあり、`builtFingerprint`が空 | `not-built` / `最終PDFは未出力` |
| `builtFingerprint`あり、現在fingerprintと不一致 | `needs-rebuild` / `最終PDFの再出力が必要` |
| `builtFingerprint`一致、`outputPdfExists=false` | `built`の記録は保持するが表示状態は`output-missing` / `前回出力が見つかりません` |
| `builtFingerprint`一致、PDF実在、blockerなし | `built` / `最終PDFは最新です` |
| workbook未レンダリング・content-pdf不整合等のblockerあり | 出力可否`canBuild=false` / `最終PDFを出力できません`。前回PDFが存在する場合は開ける |

`Mark-VolumeNeedsRebuild`は残すが、**freshnessの補助情報を追加する関数**として位置付ける。

```powershell
function Mark-VolumeNeedsRebuild(
    $Structure,
    [string]$Language,
    [string]$Category,
    [string[]]$Volumes,
    [string]$Type,
    [string]$Detail
)
```

- `$Volumes`は必ず明示し、影響したvolumeだけを渡す。
- titleだけの変更では呼ばない。
- 初回出力前は`status`を`not-built`のままにし、理由も原則蓄積しない。
- 出力済みの場合のみ`needs-rebuild`へ遷移させ、同typeの先頭理由は件数を集約する。

### A-2. readiness・fingerprint・Build-FinalPdfを共通化する

`GET /api/final/readiness?category=ecm`を新設する。categoryはB-4のとおり必須。

唯一の検証・スナップショット実装を次に集約する。

```powershell
function Get-FinalBuildInputSnapshot(
    $Structure,
    [string]$Language,
    [string]$Volume,
    [string]$Category
)

function Get-FinalBuildFingerprint($Snapshot)

function Get-FinalBuildReadiness(
    $Structure,
    [string]$Language,
    [string]$Volume,
    [string]$Category
)
```

**fingerprint対象は、実際の最終PDF内容を決める入力だけ**とする。

- 対象ページの`pageId`、順番、`enabled`、`volume`、`numberingMode`
- `contentPdf`の正規化済み相対パス
- `contentPdf`のファイルサイズと`LastWriteTimeUtc.Ticks`
- workbookの`lastRenderedVersionId`
- 組版設定のバージョン（余白・パンチシフト・ページ番号仕様を変更した場合に更新する定数）

**fingerprint対象外**

- `currentExcelHash`

元ExcelがJava組版中に更新されても、Javaが読んだ入力は`lastRenderedVersionId`配下のcontent-pdfであるため、そのPDF自体を破棄しない。代わりに、組版完了後のreadiness再評価でExcel更新をblockerとして検知し、出力済みPDFを保持したまま`needs-rebuild`へ遷移させる。

**Build-FinalPdfの処理順**

```text
1. volume_<volume>_<category>.lock を取得
2. structure.lock内で最新structureを読み直す
3. readiness確認、snapshotBeforeとfingerprintBeforeを作る
4. manifestをsnapshotBeforeから作成してstructure.lockを解放
5. Java組版（structure.lockを保持しない）
6. structure.lock内で最新structureを読み直し、fingerprintAfterを再計算
7-A. fingerprintBefore != fingerprintAfter
     → 一時PDFを削除、既存の最終PDFは差し替えない
     → needs-rebuildのまま「組版中にページ構成またはPDF入力が変更されました」
7-B. fingerprint一致
     → 一時PDFを最終出力へ差し替え
     → builtFingerprint=fingerprintBefore、lastBuiltAt、outputPdfを保存
     → staleReasonsを一旦クリア
     → readinessを再評価し、元Excel更新等のblockerがあればPDFは保持したまま
       status=needs-rebuild、理由を再付与する
```

readinessレスポンス例：

```jsonc
{
  "ok": true,
  "volumes": {
    "ja-main": {
      "canBuild": false,
      "pageCount": 24,
      "status": "needs-rebuild",
      "displayState": "blocked",
      "builtFingerprint": "sha256:...",
      "currentFingerprint": "sha256:...",
      "outputPdfExists": true,
      "blockers": [
        {
          "code": "stale-content",
          "pageTitle": "3 業績サマリー",
          "workbookName": "..._ECM_J_1.xlsx",
          "message": "元Excelが更新されています。先にPDF作成してください。"
        }
      ],
      "staleReasons": []
    }
  }
}
```

### A-3. UI：状態と文言を一対一にする

**グローバルヘッダ・ステップバー・ダッシュボード・最終PDFカードで同じ判定関数を使う。** 表示優先順位は次のとおり。

| 優先 | 条件 | 文言 |
|---:|---|---|
| 1 | 対象volumeにblockerあり | `最終PDFを出力できません`（カード内に具体的理由） |
| 2 | `builtFingerprint`あり・現在fingerprint不一致 | `最終PDFの再出力が必要` |
| 3 | fingerprint一致・`outputPdfExists=false` | `前回出力が見つかりません` |
| 4 | `builtFingerprint`なし | `最終PDFは未出力` |
| 5 | fingerprint一致・PDF実在・blockerなし | `最終PDFは最新です` |

- 集約判定では`pageCount===0`のvolumeを除外する。本体のみの案件で、空の補足が永久に未完了にならないようにする。
- ヘッダの状態表示をクリックすると最終PDF画面へ遷移する。
- `needs-rebuild`でも`outputPdf && outputPdfExists`なら`前回出力を開く`を表示し続ける。
- blockerがあるときは出力ボタンをdisabledにし、`Excel登録・PDF作成へ →`を出す。
- `output-missing`では出力ボタンを有効、前回出力リンクを非表示にする。

**ステップ④の完了条件**

```js
status === 'built'
&& outputPdfExists === true
&& blockers.length === 0
&& builtFingerprint === currentFingerprint
```

### A-4. Excel登録画面・ページ構成画面にも波及表示

- Excel登録画面：PDF作成完了通知に`ページ構成を確認 →`を出す。
- 組版中に元Excelだけが更新された場合、作成済みPDFは破棄せず、通知に`元Excelが更新されたため、PDFを再作成後に最終PDFを再出力してください`と出す。
- ページ構成画面：`ページ構成を変更しました。最終PDFの再出力が必要です`を表示する。
- fingerprint不一致だが`staleReasons`が空の場合は、汎用理由`最終PDFの入力が変更されました`を表示する。

## 3. B. カテゴリ別にボリューム状態を分ける（データ構造）

### B-1. volumesのキーをカテゴリ込みにする（言語別6キー）

`Get-WorkspacePath`は`dataDir\<language>`を返すため、structure.jsonは言語別である。1ファイルが持つ状態キーは2 volume × 3 categoryの**6キー**とする。

```text
ja/structure.json:
  ja-main|ecm      ja-appendix|ecm
  ja-main|bod      ja-appendix|bod
  ja-main|dmm      ja-appendix|dmm

en/structure.json:
  en-main|ecm      en-appendix|ecm
  en-main|bod      en-appendix|bod
  en-main|dmm      en-appendix|dmm
```

```powershell
function Get-VolumeStateKey([string]$Volume, [string]$Category)
```

- 通常の書込経路ではcategoryを必須とし、`|_`へ新規状態を書かない。
- `|_`を扱う場合は、旧データの防御的な読み取り・移行専用とする。
- `New-EmptyStructure`はその言語の6キーだけを初期化する。
- schema移行はG-4の`Initialize-Or-MigrateStructure`内だけで行う。
- 旧schemaの実データは**各言語2キー**である。V2の12キー案は未実装のため、標準移行fixtureは旧2キーを用いる。試作版等で部分的な12キー相当が存在する場合のみ、防御的fixtureを別途追加する。

参照置換対象：`Build-FinalPdf`、readiness、`Mark-VolumeNeedsRebuild`、`Serve-FinalPdfByVolume`、`renderVolumeLinks()`、`renderFinalOverview()`、`buildAllVolumes()`、履歴生成。

### B-2. ページ番号の既定をカテゴリ内で計算する

`Apply-DefaultNumberingPerVolume`に`[string]$Category`を追加し、対象ページをvolume × categoryで絞る。カテゴリ未指定で全カテゴリをまとめて処理する入口は作らず、呼び出し側がecm/bod/dmmを明示的に列挙する。

### B-3. 前回出力PDFを開く経路をカテゴリ対応にする

| 箇所 | 変更 |
|---|---|
| `Serve-FinalPdfByVolume` | `[string]$Category`を追加し、`Get-VolumeStateKey`経由で参照 |
| GET `/api/final/file` | queryの`category`を必須取得 |
| POST `/api/final/file` | bodyの`category`を必須取得 |
| `openFinalVolume()` | `{ volume, category: activePreset }`を送信 |
| `renderVolumeLinks()` | `dataset.category`を保持して渡す |

後方互換のためのcategory省略探索は行わない。旧クライアントからの省略呼び出しはB-4により拒否する。

### B-4. 最終PDF系APIのcategoryを必須化する（fail-closed）

現行`Test-WorkbookCategory`はcategoryが空なら全件マッチするため、最終PDF経路では使用前に必ずcategoryを検証する。

```powershell
function Require-WorkbookCategory([string]$Category) {
    $cat = Normalize-WorkbookCategory $Category ''
    if (@('ecm','bod','dmm') -notcontains $cat) {
        throw [System.ArgumentException]::new(
            'categoryには ecm / bod / dmm のいずれかを指定してください。'
        )
    }
    return $cat
}
```

必須化するAPI・関数：

- POST `/api/final/build`
- GET `/api/final/readiness`
- GET/POST `/api/final/file`
- POST `/api/pages/sort-by-sheet`
- category別volume状態を更新するページ操作
- `Build-FinalPdf`、`Get-FinalBuildReadiness`、`Serve-FinalPdfByVolume`

**HTTPステータス**：現行`Handle-Api`の共通catchはすべて500にするため、`throw`だけではT23を満たさない。`ArgumentException`または専用validation例外を400へ変換するcatchを追加するか、ルート内で検証して`Write-JsonResponse 400`を返す。

```powershell
catch [System.ArgumentException] {
    Write-JsonResponse $Context 400 ([ordered]@{ ok=$false; error=$_.Exception.Message })
}
catch {
    Write-JsonResponse $Context 500 ([ordered]@{ ok=$false; error=$_.Exception.Message; detail=[string]$_ })
}
```

### B-5. manifest・一時PDFをカテゴリ込みの決定的名称にする

カテゴリ別volumeロックへ変更すると、異なるカテゴリの同一volumeが同時に組版できる。現行`manifest_$Volume.json`は衝突するため、次へ変更する。

```powershell
$manifestPath = Join-Path $workspace "exports\manifest_${Volume}_${Category}.json"
$outTmpPath   = Join-Path $paths.outputDir "~building_${Volume}_${Category}.pdf"
$lockPath     = Join-Path $workspace "locks\volume_${Volume}_${Category}.lock"
```

- 同一volume × categoryは上記ロックで直列化されるためGUIDは付けない。
- manifestは削除せず、次回ビルドで上書きする。前回どの構成で組版したかを調査できるよう、`inputFingerprint`、`createdAt`、対象projectIdを含める。
- 一時PDFは成功時にMoveされる。異常終了時に残った一時PDFは次回開始時に削除してから組版する。
- manifest・一時PDF・最終出力名のcategory整合をselfcheckで確認する。

## 4. C. ページ構成の並び順

### C-1. 「シート名順に並べ替え」を追加する

**API**: `POST /api/pages/sort-by-sheet`
```jsonc
{ "category": "ecm", "volumes": ["ja-main", "ja-appendix"] }   // volumes省略時は none を除く全て
```

**ソートキー（安定ソート）**
1. シート名の半角数字（`Get-SheetOrderNumber`。数値化できないものは 999999）
2. ファイル名の `_(\d{1,4})_`（`Get-FileOrderNumber`。無ければ 999999）
3. ファイル名の文字列順（`[StringComparer]::OrdinalIgnoreCase`）
4. `Resolve-PageId`

**動作**
- **本体/補足/出力しない の割り当ては変更しない。** 各表の中だけを並べ替える。
- 対象ページの `orderManual` を `$false` に戻す。
- `order` を C-2 の採番規則で振り直す。
- `Apply-DefaultNumberingPerVolume` を再実行（先頭ページのページ番号なしが正しく移る）。
- `Mark-VolumeNeedsRebuild ... -Type 'reorder'`。

**UI**（ページ構成画面のツールバー右端）
```
並び順： シート名順        [ シート名順に並べ替え ]
        ↑ 手動 のときは attention 色
```
- ボタンは `secondary`。押下時に確認ダイアログ：
  「本体・補足・出力しない の割り当てはそのままで、それぞれの表の中だけをシート名の数字順に並べ替えます。よろしいですか？」
- 状態表示のロジック：対象カテゴリのページに `orderManual=true` が1件でもあれば `手動`、無ければ `シート名順`。
- ページ構成画面が `シート名順` のときはボタンを `disabled` にはせず、押せば冪等に再整列できるようにする（新規追加直後の整列に使うため）。

### C-2. orderの採番スケールを統一し、新規ページを安全に挿入する

`order`はvolume × category内の連番×10（10, 20, 30, …）だけを使用する。`Get-OrderHint`は挿入位置判定にのみ用いる。

```powershell
function Renumber-VolumeOrder($Structure, [string]$Volume, [string]$Category)
function Insert-PageInSheetOrder($Structure, $NewPage, [string]$Volume, [string]$Category)
```

**手動順と自動挿入の優先関係**

- 対象volume × categoryに`orderManual=$true`が1件でもあれば、新規ページは末尾へ追加する。
- 全件`orderManual=$false`なら、シート番号・ファイル順の正しい位置へ挿入する。
- 利用者の手動順をシステムが自動で組み替えない。

`Sync-Pages`の戻り値はboolではなく件数・対象IDを返す。

```jsonc
{
  "addedCount": 5,
  "insertedInOrderCount": 2,
  "insertedAtEndCount": 3,
  "insertedAtEndPageIds": ["...", "...", "..."]
}
```

UIは`insertedAtEndCount > 0`の場合に、`新しい3ページを末尾に追加しました`と`シート名順に並べ替え`ボタンを表示する。`insertedAtEndPageIds`は該当行へのスクロール・一時ハイライトにも利用する。

**既存データ移行**：G-4のschema移行時に、全volume × categoryを現在のorder順で10刻みに正規化する。手動フラグは維持する。

**受入基準**

1. シート名順の状態でシート2を追加すると1と3の間へ入る。
2. 手動順の状態では末尾に入り、既存順は変わらない。
3. `シート名順に並べ替え`で1,2,3…へ戻る。
4. 本体・補足・出力しないの割当は変わらない。

### C-3. ページ構成の細かい操作性

| 項目 | 変更内容 |
|---|---|
| 通し番号 | `順` 列は volume 内の連番のまま。ただし `font-variant-numeric: tabular-nums` を指定して桁ズレを止める |
| ドラッグ | 現行の青い挿入行方式は維持（評価が高い）。ハンドルの当たり判定を 24×24px 以上に |
| キーボード | 行選択中に `Alt+↑/↓` で1つ移動、`Alt+Shift+↑/↓` で先頭/末尾へ。ドラッグが難しい長い表の救済 |
| タイトル編集 | `input` の `blur` 時のみ保存する現行方式を維持しつつ、D-4（再描画によるフォーカス喪失）を修正 |
| プレビュー | 行ダブルクリックに加え、`ページ名` セルのファイル名部分クリックでも開く |

---

## 5. D. 導線をつなぐ

### D-1. カテゴリ切替をグローバルヘッダへ移す

- sticky header 左端に `ECM / BOD / DMM` のセグメントコントロールを常設し、**全画面から切り替え可能**にする。
- ダッシュボードの `カテゴリ切替` カード、提出フォルダ画面の `現在のカテゴリ` カードは**削除**（重複のため）。
- 各画面の `data-current-category` チップも削除（ヘッダに常時出ているため冗長）。
- 切替時の通知（`ECMに切り替えました`）は**出さない**。セグメントの見た目で十分で、通知が流れると本当の警告が埋もれる。

### D-2. ステップバーを常設し、完了後に次を提示する

sticky header直下に4ステップバーを常設する。

```text
① 提出フォルダ ──✓── ② Excel登録・PDF作成 ──●── ③ ページ構成 ──○── ④ 最終PDF
```

- ①：`configured()`
- ②：登録済み>0かつPDF作成必要件数=0
- ③：対象カテゴリに有効ページが1枚以上
- ④：A-3の完成条件を満たす
- blocker、needs-rebuild、output-missingは`!`表示にする。
- pageCount=0のvolumeは④の集約から除外する。
- 各ステップはクリックで遷移できる。

**完了時の次アクション**

| 完了したこと | 通知タイトル | ボタン |
|---|---|---|
| Excel登録 | `3件を登録しました` | `登録した3件をPDF作成 →` |
| PDF作成 | `PDFを作成しました` | `ページ構成を確認 →` |
| ページ構成保存 | `ページ構成を保存しました` | readinessがbuild可能なら`最終PDFを出力 →` |
| 最終PDF出力 | `本体PDFを出力しました` | `出力フォルダを開く` / `PDFを開く` |

登録後は登録したworkbookIdを選択状態にする。

### D-3. ダッシュボードの「次にやること」にreadinessを組み込む

上から順に評価する。

```text
1. !configured()                         → 提出フォルダを設定してください
2. 登録済み===0                           → Excelを登録してください
3. needsPdf>0                             → PDF必要分を作成してください（N件）
4. 対象volumeにblockersあり               → 最終PDFを出力できません（理由）
5. 未割当ページ>0                         → ページ構成を確認してください（Nページ）
6. builtFingerprint不一致                 → 最終PDFの再出力が必要です
7. fingerprint一致・outputPdfExists=false → 前回出力が見つかりません
8. builtFingerprintなし                   → 最終PDFを出力してください
9. 対象の全volumeが完成条件を満たす        → 最新の状態です
```

9は灰色のチェックだけにし、成功状態を過度に強調しない。pageCount=0のvolumeは6〜9の集約から除外する。

### D-4. ポーリング再描画でフォーカスを失わない

`renderAll()` を条件付きにする。

```js
function isEditing() {
  const el = document.activeElement;
  return !!el && el.matches('input:not([readonly]), select, textarea');
}
```
- `scanUpdatesSilently()` は `isEditing()` が true のとき、`renderPages()` と `renderFileList()` と `renderWorkbooks()` をスキップし、サマリー・バッジ・ヘッダのみ更新する。次の周期で再試行。
- ドラッグ中（`draggingRow !== null`）とモーダル表示中も同様にスキップ。
- あわせて `renderPages()` に「再描画前後で `currentBoardSignature()` が同一なら DOM を差し替えない」早期リターンを入れる。

### D-5. 通知の整理

- `showMessage(type)` に `autoHideMs` を追加。`ok` と `''`（処理中）は 4000ms で自動フェードアウト、`warn` は 8000ms、`danger` は自動で消さない。
- `danger` のとき、上部通知と赤エラーパネルの**両方**を出すのをやめ、**エラーパネルのみ**にする（現行は同じ内容が2箇所に出る）。
- 通知は画面上部固定ではなく **右下トースト**に変更（Linear系。作業対象の表を隠さないため）。ただしエラーパネルは従来どおり本文上部に留める。

### D-6. その他の導線

- 最終PDF画面に`本体・補足をまとめて出力`を追加し、`buildAllVolumes()`へ接続する。
- Excel登録画面に`すべて再作成`を**secondaryボタンとして直接配置**する。オーバーフローメニューは新設しない。
- 提出フォルダ画面とExcel登録画面の両方に`更新確認`を置く。
- 最終PDF出力後にfingerprint差分またはblockerが生じている場合は、成功通知の後に状態警告を続けて表示する。

## 6. E. 死んでいるUIの整理

**削除するもの**
| 要素 | 対応 |
|---|---|
| `.collapse-hint`（メニューを閉じる） | 削除。サイドバーは常時表示・幅240px固定 |
| `.pagination-lite`（ページャ5ボタン） | 削除。登録済みExcelは全件表示（実運用で最大90件程度、仮想化不要） |
| `.info-dot`（`i` / `?`） | 削除。必要な説明はラベル直下の12pxキャプションに書く |
| `.menu-col`（`⋮`） | 削除。行操作は行を選択してツールバーで行う |
| 出力先/管理データの「自動設定」ボタン（disabled） | 削除。値は自動決定である旨をキャプションで説明 |
| `最近使用した提出フォルダ` カード | 削除（実データが現在のフォルダ1件しかない）。履歴を本実装する場合は共有フォルダ上の `config.json` ではなく `%LOCALAPPDATA%\ReportBinder\recent.json` に保存すること |
| 出力履歴の `出力者` 列 | 削除（常に固定値） |
| `ページ番号の設定` カード | 静的説明のみのため、最終PDFカード下の1行キャプション（`ページ番号：- 2 - 形式／先頭ページは番号なし`）に縮小 |

**機能を付けるもの**
| 要素 | 対応 |
|---|---|
| 言語切替バッジ（`日本語 ▾`） | クリックで `日本語管理 / 英語管理` の切替を案内するポップオーバーを出す。ワークスペースが別プロセスのため、`英語管理.vbs を起動してください` と導線を示す（自動切替は行わない） |
| `bind()` の死にID 9個 | HTMLに対応要素が無いものは`bind`呼び出しごと削除。`build-all-btn`は最終PDF画面、`render-all-btn`はExcel登録画面の直接表示secondaryボタンとして接続する |

---

## 7. F. デザイン仕様（Apple × Linear）

### F-0. 方針

- ベースは白。カードは1px線で区切り、影はモーダル／ポップオーバーだけに使う。
- アクセントはindigo `#5E6AD2` 1色。
- 状態色は要対応・エラーの2色だけ許可する。
- **淡い面塗りはステータスバッジ／トースト／エラーパネルに限り`--*-subtle`を許可する。** 通常カード、ナビ、装飾背景には状態色の面塗りを使わない。
- 達成状態はグレーで控えめにし、未処理・再出力・エラーだけを視認しやすくする。
- 余白は大きく、表の情報密度は維持する。

### F-1. トークン（`style.css` の `:root` を全面置換）

```css
:root {
  /* surface */
  --bg:              #FFFFFF;
  --surface:         #FFFFFF;
  --surface-subtle:  #FAFAFB;   /* 行hover・入力欄・空状態 */
  --surface-sunken:  #F6F7F8;   /* コードブロック・開発用詳細 */

  /* line — Linear系の1px */
  --border:          #E6E7EA;
  --border-strong:   #D5D7DC;   /* 入力欄・区切りの強調 */

  /* text — 階層は3段のみ */
  --text:            #1B1C1E;
  --text-secondary:  #6B6F76;
  --text-tertiary:   #9A9EA6;
  --text-on-accent:  #FFFFFF;

  /* accent — 1色 */
  --accent:          #5E6AD2;
  --accent-hover:    #515DC4;
  --accent-active:   #4650B0;
  --accent-subtle:   #F1F2FB;   /* 選択行・現在地の背景 */
  --accent-border:   #C9CDF0;

  /* status — 例外の2色。ステータスUIに限り淡いtintを許可 */
  --attention:       #B45309;   /* 要対応：PDF作成が必要 / 再出力が必要 */
  --attention-subtle:#FDF6EC;
  --attention-border:#EBD9BE;
  --danger:          #B42318;   /* エラー */
  --danger-subtle:   #FEF3F2;
  --danger-border:   #F0C2BD;

  /* radius — Linear系。角は控えめ */
  --r-sm: 4px;
  --r-md: 6px;
  --r-lg: 8px;    /* カード */
  --r-xl: 12px;   /* モーダル */

  /* spacing */
  --s-1: 4px;  --s-2: 8px;  --s-3: 12px; --s-4: 16px;
  --s-5: 24px; --s-6: 32px; --s-7: 48px; --s-8: 64px;

  /* elevation — モーダルとポップオーバーのみ */
  --shadow-pop:   0 4px 16px rgba(27,28,30,.10), 0 0 0 1px rgba(27,28,30,.05);
  --shadow-modal: 0 24px 64px rgba(27,28,30,.18), 0 0 0 1px rgba(27,28,30,.06);

  /* layout */
  --sidebar-w: 240px;
  --header-h:  56px;
  --step-h:    44px;
  --content-max: 1160px;
}
```

**削除するトークン**: `--primary`系 / `--green`系 / `--orange`系 / `--purple`系 / `--red`系 / `--shadow-sm` / `--shadow` / `--surface-soft` / `--line`系。`body` の `radial-gradient` と `linear-gradient` も削除して `background: var(--bg)` にする。

### F-2. タイポグラフィ

```css
body {
  font-family: -apple-system, "Segoe UI Variable Text", "Segoe UI",
               "BIZ UDPGothic", "BIZ UDPゴシック", "Yu Gothic UI", Meiryo, sans-serif;
  font-size: 14px;
  line-height: 1.6;
  color: var(--text);
  -webkit-font-smoothing: antialiased;
}
```
> ラテン文字と数字は Segoe UI 系、日本語は BIZ UDPゴシックに落ちる順序。BIZ UDPゴシックは社内の可読性方針として残す。

| 役割 | size / line-height / weight / letter-spacing | color |
|---|---|---|
| 画面タイトル (h2) | 24px / 1.25 / 600 / -0.02em | `--text` |
| セクション見出し (h3) | 15px / 1.4 / 600 / -0.01em | `--text` |
| 本文 | 14px / 1.6 / 400 | `--text` |
| 表ヘッダ | 12px / 1.4 / 500 | `--text-secondary` |
| キャプション | 12px / 1.5 / 400 | `--text-secondary` |
| 数値（件数・ページ数） | 28px / 1.1 / 600 / `tabular-nums` | `--text` |
| バッジ・チップ | 12px / 1 / 500 | 状況による |

- **`font-weight: 900` は全廃。最大 600。**
- 表の数値セル・日時セル・順番セルに `font-variant-numeric: tabular-nums` を必ず指定する。

### F-3. レイアウト

```
┌──────────┬────────────────────────────────────────────────┐
│          │ ┌── sticky header (56px, backdrop-filter) ───┐ │
│ sidebar  │ │ ECM BOD DMM        ● 再出力が必要   日本語 ▾ │ │
│  240px   │ ├── step bar (44px, sticky, 白, 下1px) ──────┤ │
│  白      │ │ ① ─✓─ ② ─●─ ③ ─○─ ④                      │ │
│  右1px   │ └────────────────────────────────────────────┘ │
│          │                                                │
│          │        content  max-width 1160px               │
│          │        padding 40px 48px                        │
│          │        section 間 32px                          │
└──────────┴────────────────────────────────────────────────┘
```

```css
.workspace-top {
  position: sticky; top: 0; z-index: 20;
  height: var(--header-h);
  background: rgba(255,255,255,.72);
  backdrop-filter: saturate(180%) blur(20px);
  -webkit-backdrop-filter: saturate(180%) blur(20px);
  border-bottom: 1px solid var(--border);
}
.step-bar { position: sticky; top: var(--header-h); z-index: 19;
            height: var(--step-h); background: var(--surface);
            border-bottom: 1px solid var(--border); }
```
> `backdrop-filter` は Edge (Chromium) で問題なく動作する。IEモード起動時のフォールバックとして `@supports not (backdrop-filter: blur(1px)) { .workspace-top { background: #fff; } }` を必ず添える。

**サイドバー**
- 背景 `--surface`、右 `1px solid var(--border)`、影なし（現行の `box-shadow: 10px 0 30px` を削除）。
- ブランドは `ReportBinder` のテキストのみ。3本線のグラデーションマークは削除し、6px角丸・`--accent` 単色の8px正方形1つに置換。
- `.nav-item`：高さ32px、`font-size:13px`、`font-weight:500`、`padding:0 8px`、`border-radius: var(--r-md)`。
  - 通常 `color: var(--text-secondary)`、hover `background: var(--surface-subtle)`。
  - active `background: var(--accent-subtle); color: var(--accent); font-weight:600`。
- ナビ間の `gap` は 2px（現行10px は間延びしている）。

### F-4. コンポーネント

**カード**
```css
.card { background: var(--surface); border: 1px solid var(--border);
        border-radius: var(--r-lg); padding: var(--s-5); box-shadow: none; }
```
入れ子のカードは作らない。現行の `dashboard-grid` 内の二重枠を解消する。

**ボタン**

| variant | 背景 | 文字 | 枠 |
|---|---|---|---|
| primary | `--accent` → hover `--accent-hover` | `--text-on-accent` | なし |
| secondary | `--surface` → hover `--surface-subtle` | `--text` | `1px var(--border-strong)` |
| ghost | 透明 → hover `--surface-subtle` | `--text-secondary` | なし |
| danger | `--surface` | `--danger` | `1px var(--danger-border)` |

- 高さ：`sm 28px / md 32px / lg 36px`。`border-radius: var(--r-md)`、`font-size:13px`、`font-weight:500`、`padding: 0 12px`。
- **`.btn.xl` `.btn.purple` `.btn.orange` は廃止。** 一括設定の3ボタンはすべて `secondary` にし、ラベルで区別する（`本体に設定` / `補足に設定` / `出力しない`）。
- フォーカス：`outline: 2px solid var(--accent); outline-offset: 2px;`。
- `runBusy` のラベル差し替え（`処理中`）は、幅が跳ねるので**ラベルを維持したままボタン内に12pxのスピナー**を左付けする方式に変更。

**バッジ**
```css
.badge { display:inline-flex; align-items:center; gap:6px;
         height:20px; padding:0 8px; border-radius: var(--r-sm);
         font-size:12px; font-weight:500; background:transparent;
         border:1px solid var(--border); color: var(--text-secondary); }
.badge.attention { color: var(--attention); border-color: var(--attention-border);
                   background: var(--attention-subtle); }
.badge.danger    { color: var(--danger);    border-color: var(--danger-border);
                   background: var(--danger-subtle); }
.badge::before   { content:""; width:6px; height:6px; border-radius:50%;
                   background: currentColor; }
.badge.neutral::before { display:none; }
```
- `最新PDFあり` → `.badge.neutral`（ドットなし・グレー）。**達成状態は目立たせない。**
- `PDF作成が必要` `再出力が必要` → `.badge.attention`。
- `作成エラー` → `.badge.danger`。
- 現行の `ok / warn / blue / danger` クラスはこの4種にマッピングし直す。

**テーブル**
```css
.data-table { width:100%; border-collapse: collapse; }
.data-table th { height:36px; font-size:12px; font-weight:500;
                 color: var(--text-secondary); text-align:left;
                 border-bottom: 1px solid var(--border); }
.data-table td { height:44px; border-bottom: 1px solid var(--border); }
.data-table tbody tr:hover { background: var(--surface-subtle); }
.data-table tr.selected-row { background: var(--accent-subtle);
                              box-shadow: inset 3px 0 0 var(--accent); }
.data-table tr:last-child td { border-bottom: none; }
```
- ゼブラストライプは使わない。
- テーブルは `.card` の内側に置き、`.card` の左右パディングを打ち消して全幅に伸ばす（`margin: 0 calc(var(--s-5) * -1)` ＋ セルに `padding-left/right: var(--s-5)`）。

**入力欄**
```css
input, select { height:32px; padding:0 10px; font-size:13px;
                border:1px solid var(--border-strong); border-radius: var(--r-md);
                background: var(--surface); }
input:focus, select:focus { border-color: var(--accent);
                            box-shadow: 0 0 0 3px var(--accent-subtle); }
input[readonly] { background: var(--surface-subtle); color: var(--text-secondary); }
```
- チェックボックスは 16px、`accent-color: var(--accent)`。

**進捗表示**
- 現行の円形リング（`.progress-ring`）は廃止。
- 2px の水平バー（`background: var(--accent)`、track は `--border`）＋ その下に `13px` で `12 / 24 件・ECM_1_業績.xlsx / シート 3` を表示。
- 完了後 1.8秒でフェードアウト（現行踏襲）。

**空状態**
```
（中央寄せ、パディング 48px 24px）
   13px --text-secondary で1行の説明
   [ プライマリアクション1つ ]
```
現行の「提出フォルダを選んでください。」等の文言はそのまま使い、ボタンを添える。

**モーダル（PDFプレビュー）**
- オーバーレイ `rgba(27,28,30,.40)` ＋ `backdrop-filter: blur(4px)`。
- カード `border-radius: var(--r-xl)`、`box-shadow: var(--shadow-modal)`、最大 `min(1100px, 92vw) × 88vh`。
- 開いたとき `閉じる` ボタンへフォーカスを移し、`Tab` をモーダル内にトラップ、閉じたら元の行へフォーカスを戻す。

### F-5. アイコン

全角記号（`⌂ □ ▦ ▤ ▣ ✓ △ ⋮ ◉ ⌄ ↻ ⌕ ⊘ ›`）を**すべて廃止**し、インラインSVGスプライトに置き換える。

```html
<svg class="icon" aria-hidden="true"><use href="#i-home"></use></svg>
```
```css
.icon { width:16px; height:16px; stroke: currentColor; stroke-width:1.5;
        fill:none; stroke-linecap:round; stroke-linejoin:round; flex:none; }
```
必要なアイコン（`index.html` 冒頭に `<svg style="display:none">` で `<symbol>` 定義）:
`i-home` `i-folder` `i-grid` `i-list` `i-file-pdf` `i-file-excel` `i-check` `i-alert` `i-chevron-right` `i-chevron-down` `i-refresh` `i-search` `i-close` `i-drag` `i-external`

外部CDNは使えないため、16×16 の手書きパスを直接埋め込むこと（総量 3KB 程度）。

### F-6. 画面別の構成変更まとめ

| 画面 | 変更 |
|---|---|
| ダッシュボード | `次にやること` ＋ サマリー4枚 ＋ `最近の状況` の3ブロックのみに縮小。`カテゴリ切替` `レポート作成の流れ` カードは削除（ヘッダとステップバーへ統合） |
| 提出フォルダ | 3行のパス設定カード1枚 ＋ `フォルダの状況` に集約。`現在のカテゴリ` `最近使用した提出フォルダ` を削除。`出力先` `管理データ` は入力欄をやめ、`--text-secondary` の12pxテキストで表示 |
| Excel登録・PDF作成 | 左右2カラムは維持。ページャ削除、`⋮`列削除。`更新確認` `必要分を選択` `解除` を右上の小さめツールバーに、`PDF作成` のみ primary |
| ページ構成 | サマリーストリップを4項目→`本体 / 補足 / 出力しない` の3項目に縮小。ツールバーに `並び順` 表示と `シート名順に並べ替え` を追加。右サイドの `出力先のサマリー` はストリップと重複するため削除し、その幅を表に回す |
| 最終PDF | 2枚のカードに鮮度セクションを追加（A-3）。`ページ番号の設定` カード削除、`出力サマリー` は履歴と統合。`本体・補足をまとめて出力` を追加 |

---

## 8. G. 土台の修正（Phase 0：UI改修より先に行う）

### G-1. structure更新をトランザクション化する（最重要）

読込→変更→保存を1つのstructureロック内で行う。

```powershell
function Read-StructureUnlocked([string]$Language)
function Write-StructureUnlocked([string]$Language, $Structure)
function Update-StructureLocked([string]$Language, [scriptblock]$Mutation)
```

```powershell
function Update-StructureLocked([string]$Language, [scriptblock]$Mutation) {
    $workspace = Get-WorkspacePath $Language
    $lockPath  = Join-Path $workspace 'locks\structure.lock'
    return Invoke-WithLock $lockPath {
        $structure = Read-StructureUnlocked $Language
        $result = & $Mutation $structure
        Write-StructureUnlocked $Language $structure
        return $result
    }
}
```

- 書込処理の唯一の入口を`Update-StructureLocked`とする。
- `Get-Structure`は表示用の読み取りに限定し、repair・移行・保存を行わない。
- `Save-Structure`の直接呼出しは廃止またはprivate化する。
- Excel COM処理・Java組版中はstructureロックを保持しない。

**ロック取得順を固定する**

```text
workbook_<id>.lock または volume_<volume>_<category>.lock
    → 必要な短時間だけ structure.lock
```

逆順は禁止する。structure.lockを持ったままworkbook lock／volume lockを待たない。

対象：登録、解除、Render結果反映、Scan-Updates、ページ更新・確認・並べ替え、Build状態反映、schema移行。

### G-2. selfcheck.pyを現行仕様とV4仕様へ更新する

現行印刷設定へ合わせる。

| selfcheck旧期待 | 現行 | 更新 |
|---|---|---|
| `CenterHorizontally = $false` | `$true` | `$true`を期待 |
| 左右1.0cm | 1.2cm | 1.2cmを期待 |
| パンチ0.3cm | 0.2cm | 0.2cmを期待 |
| `render-all-btn`禁止 | V4で直接配置 | 禁止チェック削除 |

追加チェック：

- category必須validatorと400応答
- 言語別6キーへの旧2キー移行、冪等性
- `structure.json.v1.bak`作成
- 移行が`Initialize-Or-MigrateStructure`＋structureロック内だけで行われること
- orderの10刻み移行
- manifest／一時PDF／volume lockのcategory込み命名
- fingerprintに`currentExcelHash`を含めず、content-pdfのサイズ・更新時刻を含めること
- `builtFingerprint`比較がfreshness判定に使われること
- 書込関数が`Update-StructureLocked`経由であること
- 最終出力後にinput再照合を行うこと

### G-3. configと配布物を分離する

**設定ファイル**

```text
共有アプリ側:
  default-config.json   読み取り専用の既定値

%LOCALAPPDATA%\ReportBinder\:
  config.json           利用者ごとの現在値
```

起動時：

1. ローカルconfigがあれば使用
2. なければ共有`default-config.json`を読み、ローカルへ初回コピー
3. 以後の保存はローカルconfigだけへ行う

共有側へ現在値を書き戻す経路をコード上から削除し、selfcheckで検査する。

**配布物**

- 実運用パス入りconfig、logs、一時job、ダウンロードZIPキャッシュは除外する。
- JREは配布方式を2種類に分ける。
  - オンライン導入版：展開済みJREを含めず`install-thirdparty.cmd`で導入
  - 社内オフライン完結版：展開済みJREを同梱
- H:共有からの初回Java起動はSMB経由クラスロードで数秒遅くなる可能性をREADMEに記載する。
- 同梱OpenJDKのビルド名・バージョン・ライセンス通知（GPLv2 + Classpath Exception）をREADME／NOTICEに記載する。

### G-4. schema移行を読み取りAPIから分離する

移行は通常の`Get-Structure`で行わず、起動時または明示初期化時の次の関数だけで行う。

```powershell
function Initialize-Or-MigrateStructure([string]$Language) {
    Update-StructureLocked $Language {
        param($structure)
        # schema判定、バックアップ、移行、検証
    }
}
```

要件：

1. structure.lock内で最新ファイルを読み直す。
2. schemaVersion 1の実データ（各言語旧2キー）を検知する。
3. 最初の変更前に同じディレクトリへ`structure.json.v1.bak`を作成する。
4. バックアップは既存なら上書きしない。作成後にサイズまたはSHA-256で元ファイルと一致確認する。
5. volumeを言語別6キーへ移行し、orderをvolume × categoryごとに10刻みへ正規化する。
6. 変換後の必須キーと参照整合性を検証してからschemaVersion 2として一時ファイル経由で保存する。
7. 2回目以降は何も変更しない。
8. 新アプリは対応上限を超えるschemaVersionを検知したら起動を拒否する。

**旧版アプリとの同時利用**

既に起動中の旧プロセスを、新コードだけで確実に停止・拒否することはできない。したがって初回移行は保守手順として次を必須とする。

- 共有アプリをV4へ切り替える前に利用者へ停止を告知する。
- 旧プロセスが終了するまで待つ（現行の30分アイドル終了を使う場合は35分以上の保守枠を確保）。
- 旧ランチャーを退避または置換し、新規起動がV4だけになる状態で移行する。
- 移行後、旧版アプリから同じdataDirを開く運用を禁止する。

ゼロダウンタイムで旧版と併存させる必要がある場合は、今回のスコープを超えるため、workspace全体を`dataDir\v2\<language>`へ分離してcontent-pdfも移行する別設計とする。

### G-5. manifest・fingerprintの診断情報を保持する

manifestは削除せず、同一volume × categoryの次回出力で上書きする。最低限次を含める。

```jsonc
{
  "schemaVersion": 2,
  "language": "ja",
  "category": "ecm",
  "volume": "ja-main",
  "projectId": "FY160-4Q_ECM",
  "inputFingerprint": "sha256:...",
  "createdAt": "...",
  "pages": []
}
```

不具合調査時に「どの入力で前回組んだか」を確認できるようにする。個人情報・元Excelの絶対パスはmanifestへ追加しない。

## 9. 実装順序

| Phase | 内容 | 理由 |
|---|---|---|
| **0A** | G-2 selfcheckの現行化、G-3配布・config分離、category validatorと400処理 | 現行の回帰確認を通し、fail-openを先に塞ぐ |
| **0B** | G-1 structureトランザクション、G-4ロック内移行＋バックアップ＋旧版停止手順 | 以降のすべての書込・移行の土台 |
| **1** | B-1言語別6キー、B-2 numbering、B-3 open API、B-5 manifest／一時PDF命名、C-2 order移行 | データ構造とカテゴリ隔離を確定 |
| **2** | A-2 snapshot／fingerprint／readiness、A-1 `builtFingerprint`・staleReasons補助化、組版完了時の再照合、C-1 sort API | freshnessの真偽判定をサーバー側で完成 |
| **3** | A-3/A-4、D-2/D-3/D-4 | 既存デザインのまま状態表示と導線を確認 |
| **4** | D-1/D-5/D-6、E-1/E-2 | HTML構造と操作導線を整理 |
| **5** | F-1〜F-5、SVG化 | ロジック安定後に外観を刷新 |

各Phase終了時に`app\tools\selfcheck.py`を実行する。Phase 0BとPhase 1は実データのコピーを使った移行ドライランを別フォルダで行い、本番dataDirへ直接試行しない。

## 10. 受入テスト

| # | 手順 | 期待結果 |
|---|---|---|
| T1 | ECMでPDF作成→本体PDF出力→Excelを1件再作成 | 本体だけ`再出力が必要`。前回PDFは開ける |
| T2 | 元Excelを上書きして30秒待つ | blockerに元Excel更新が出て最終PDF出力不可。前回PDFは保持 |
| T3 | 日本語本体を並べ替える／titleだけ変更する | 並べ替えでは本体だけneeds-rebuild。titleだけでは変化しない |
| T4 | ECM本体出力後にBODへ切替 | BODは未出力、ECMリンクは出ない |
| T5 | ECM/BODの両方を登録しBOD本体を出力 | BOD先頭ページは番号なし |
| T6 | 手動順からシート名順へ戻す | 各表内だけ数字順へ戻り、割当は維持 |
| T7a | シート名順でシート2を追加 | 1と3の間に入る |
| T7b | 手動順でシート2を追加 | 末尾へ入り、件数通知と対象行ハイライトが出る |
| T8 | タイトル入力中にポーリング | フォーカスとキャレット維持 |
| T9 | Excelを3件登録 | 3件選択済み、登録した3件をPDF作成ボタン表示 |
| T10 | 各画面へ切替 | カテゴリセグメントが常時表示 |
| T11 | ページャ・行末⋮・info-dot等を探す | 存在しない。`すべて再作成`は直接表示ボタン |
| T12 | Edge表示 | 白＋1px、影はモーダル中心、状態バッジだけ淡いtint |
| T13 | Tabだけで操作 | フォーカス可視、モーダルfocus trap・復帰 |
| T14 | ECM/BODのPDFを別々に開く | 各カテゴリの正しいPDF |
| T15 | stale理由がある状態で再出力 | 入力が変わらなければ理由クリア・最新へ戻る |
| T16 | 出力PDFを手動削除 | `前回出力が見つかりません`、ステップ④未完了、出力ボタン有効 |
| T17 | PDF作成ジョブ中に別PCでページ順変更 | 更新消失・JSON破損がない |
| T18 | 旧2キーstructureを移行 | 言語別6キー、order10刻み。2回目は変更なし |
| T19 | 配布ZIPを展開 | 実運用config、logs、job、ダウンロードキャッシュなし |
| T20 | selfcheck実行 | `selfcheck ok` |
| T21 | ECM本体とBOD本体を同時出力 | manifest・一時PDFが衝突せず、各カテゴリの正しいPDFが完成 |
| T22a | 最終PDF組版中にページ順またはcontent-pdfを変更 | 一時PDFを破棄し、既存最終PDFを差し替えず再実行を案内 |
| T22b | 最終PDF組版中に元Excelだけを更新 | 組版PDFは保持するが完了直後にneeds-rebuild／blockerとなる |
| T23 | `/api/final/build`・readiness・fileをcategory省略で呼ぶ | HTTP 400、混合PDFを生成しない |
| T24 | builtの出力PDFを削除 | ステップ④未完了、ヘッダが`前回出力が見つかりません` |
| T25 | schemaVersion 1の**実際の旧2キー**structureを置き、同時に複数の読み取り要求を行う | `structure.json.v1.bak`を1回だけ作成し、ロック内で6キーへ移行、破損なし |
| T26 | schemaVersion 2移行後に対応外の旧版を起動しようとする | 運用手順で旧ランチャーが利用不能。新アプリは上位schemaを拒否 |
| T27 | content-pdfを同じパスで手動上書き | サイズまたは更新時刻の差でfingerprint不一致を検知 |
| T28 | イベント呼び出しを意図的に省いたテスト用ページ更新を行う | `builtFingerprint`比較でneeds-rebuildを検知し、最新のまま残らない |
| T29 | 初回起動ユーザー／既存ユーザーで設定を保存 | 共有defaultを初回コピーし、その後はLOCALAPPDATAだけが更新される |

## 11. 変更ファイルとスコープ外

**変更対象**

- `app/server.ps1`
  - category必須validatorと400応答
  - structureのUnlocked read/writeとトランザクション
  - ロック内schema移行・バックアップ
  - 言語別6キー、カテゴリ別ロック／manifest／一時PDF
  - input snapshot・fingerprint・`builtFingerprint`
  - readiness共通化、組版完了時の入力再照合
  - state遷移、staleReasons補助化
  - order採番、sort API、category別numbering
- `app/web/app.js`
  - readinessを使った一貫表示
  - category付きopen API
  - ステップ／ダッシュボード／通知／再描画制御
  - `すべて再作成`直接表示
- `app/web/index.html`
  - sticky header、step bar、SVG、不要UI削除
- `app/web/style.css`
  - Apple × Linear仕様へ置換。状態バッジだけ淡いtintを許可
- `app/tools/selfcheck.py`
  - 現行印刷設定、category fail-closed、移行、fingerprint、命名、LOCALAPPDATA書込の静的・fixtureチェック
- `app/default-config.json`（新規）
  - 共有の読み取り専用既定値
- `%LOCALAPPDATA%\ReportBinder\config.json`
  - 実行時に作成する利用者別現在値
- パッケージ手順書／README／NOTICE
  - 配布除外、オンライン／オフラインJRE、SMB初回遅延、OpenJDKライセンス、V4移行停止手順

**スコープ外**

- PDF組版仕様そのもの（左右1.2cm、パンチ0.2cm、水平中央、`- 2 -`）
- Excel COM印刷設定プロファイル
- VBS / Edge / ローカルサーバー / 30分アイドル終了という起動方式
- 旧版とV4をゼロダウンタイムで同一dataDir上に併存させる仕組み
- content-pdfの暗号学的ハッシュ常時計算。V4ではパス・サイズ・更新時刻・versionIdをfingerprintに使い、性能と検知力を両立する
