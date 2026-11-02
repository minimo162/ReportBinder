# ReportBinder Local API V4

すべて`127.0.0.1`だけで待ち受けます。`X-ReportBinder-Token`ヘッダーまたは`?token=`が必要です。

validationエラーはHTTP 400、未処理エラーはHTTP 500です。

管理データの保存形式は汎用資料パック対応の`schemaVersion: 3`です。Local API V4は既存UIとの互換性のため、`structure`を従来の`schemaVersion: 2`、`workbooks / pages / volumes`形式へ変換して返します。保存データ上では`pack / source / unit / item / artifact / output`と互換フィールドが同期されます。

V4の`/api/submission-files`、`/api/workbooks/register*`、`/api/workbooks/render/start`は、内部では共通原稿dispatcherを経由し、Excelは`excel-com-v1`、Wordは`word-com-v1`、PowerPointは`powerpoint-com-v1`、PDFは`pdfbox-import-v1`アダプターへ接続します。レスポンスには共通項目として`sourceId / sourceType / adapterId / packId`が追加される場合があります。従来の`workbookId / category`も維持します。

## V2 汎用資料パックAPI

新UI向けの汎用ドメインAPIは`/api/v2`に置きます。APIの`v2`はHTTP契約、レスポンスの`domainSchemaVersion: 3`は保存ドメインの版を表します。

> **現行UIが使っていない経路について（2026-08-10 時点の実態）**
>
> 書き込み系はほぼV2へ移行済みですが、**読み取り系は現行UIが旧APIを使い続けています**。現行UIは`workbooks / pages`形式を前提としており、`GET /api/v2/state`はそれを返さないためです（`app/web/app.js`の`registerSelectedFiles`付近のコメントを参照）。
>
> 次の6件は**画面から一度も呼ばれていません**。実装は既存関数への別名で、動作確認は行われていない前提で扱ってください。
>
> | 未使用の経路 | 現行UIが代わりに使うもの |
> | --- | --- |
> | `GET /api/v2/state` | `GET /api/state` |
> | `GET /api/v2/pack-templates` | `/api/state`の`packTemplates` |
> | `GET /api/v2/source-candidates` | `GET /api/submission-files` |
> | `POST /api/v2/sources/unregister` | `POST /api/workbooks/unregister` |
> | `POST /api/v2/sources/scan-updates` | `POST /api/scan-updates` |
> | `PATCH /api/v2/items/{itemId}` | `POST /api/pages/update` |
>
> `POST /api/v2/items/reorder`、`PATCH /api/v2/items/{itemId}`、`POST /api/v2/outputs/spread-preview` は `baseLayout` を受け取り、現行レイアウトの指紋と一致しない場合に409 `structure-conflict`を返します。画面は直前のレスポンスに含まれる `layoutFingerprint` を次の変更の `baseLayout` として送ってください。

- `GET /api/v2/state`: `packs / sources / units / items / artifacts / outputs / packProgress`を返す。`packProgress`は未提出必須原稿に加え、`overdueRequiredCount / dueSoonRequiredCount / nearestRequiredDueDate`で期限超過・7日以内・次の期限を集計する
- `GET /api/v2/pack-templates`: 組み込みテンプレートと利用者定義ひな形を返す
- `POST /api/v2/pack-templates`: 利用者定義ひな形を作成する
- `PATCH /api/v2/pack-templates/{templateId}`: 利用者定義ひな形を更新し、版を進める
- `DELETE /api/v2/pack-templates/{templateId}`: 未使用の利用者定義ひな形を削除する
- `GET /api/v2/packs`: 使用中の資料パック一覧を返す。`includeArchived=true`でアーカイブ済みも含める
- `POST /api/v2/packs`: 任意の資料パックを作成する
- `POST /api/v2/packs/{packId}/duplicate`: 設定を引き継いで資料パックを複製する
- `GET /api/v2/packs/{packId}/template-upgrade`: 利用中のひな形と最新版との差分を返す
- `POST /api/v2/packs/{packId}/template-upgrade`: 確認後に最新版のひな形を資料パックへ適用する
- `PATCH /api/v2/packs/{packId}`: 名称または仕上げ設定を変更する
- `POST /api/v2/packs/{packId}/archive`: 任意資料パックを一覧からアーカイブする
- `POST /api/v2/packs/{packId}/restore`: アーカイブした資料パックを復元する
- `GET /api/v2/packs/{packId}/review`: 必須出力の指紋、提出可否、レビュー状態、監査履歴を返す
- `POST /api/v2/packs/{packId}/review`: `submit / approve / request-changes / reopen`でレビュー状態を更新する
- `GET /api/v2/source-candidates`: 登録可能なExcel（`.xlsx` / `.xlsm`）・Word（`.docx`）・PowerPoint（`.pptx`）・PDF原稿を返す
- `POST /api/v2/sources/register-batch`: 複数形式の原稿を一括登録する
- `POST /api/v2/sources/unregister`: `sourceId`で登録解除する
- `POST /api/v2/sources/scan-updates`: 登録済み原稿の更新を検出する
- `POST /api/v2/sources/render/start`: `sourceIds`で変換PDF作成を開始する。Excelは保存済みの`excelSheetSelectionMode`（新規登録は`numeric-only`、既存の明示設定や旧データの欠損フィールドは互換のため維持）をジョブへ固定する。必要ならリクエストにも同名項目を指定できる
- `PATCH /api/v2/sources/{sourceId}`: `ownerDepartment / required / defaultTargetId`を更新する。Excelでは`excelSheetSelectionMode`も指定でき、`numeric-only`はシート名がstrictな`^[0-9]+$`の表示シートだけをPDF化する。対象が0件の場合はPDF・ページ構成を変更せずエラーにする
- `POST /api/v2/items/reorder`: 資料パック内のページを本体・補足・未振り分けへ移動、並べ替えする。ページ画面の「半角数字シートをまとめる」はこのAPIを使い、Excelのstrictな半角数字シートを数字順に指定先へ集め、非該当Excelを未振り分けへ戻す。`baseLayout`を指定した場合は競合を検出する。送ったレーン（`volumes`または`targets`のキー）に今あるページは、送った内容のどこかにちょうど1回ずつ含めてください。空のレーン、ページの欠落、この一式に無いID、同じIDの重複、同じ出力先の二重指定はどれも400を返し、何も保存しない
- `PATCH /api/v2/items/{itemId}`: ページ名、番号表示、ページ範囲、出力先を変更する（現行UIは未使用）。`baseLayout`を指定した場合は競合を検出する
- `GET /api/v2/layout/snapshots?packId={packId}`: 資料パックの保存済みページ構成を返す
- `POST /api/v2/layout/restore/preview`: 保存済み構成を復元した場合の適用数・差異を返す
- `POST /api/v2/layout/restore`: 保存済みのページ構成を同じ資料パックへ復元する
- `GET /api/v2/outputs/readiness?packId={packId}`: 本体・補足の出力可否と再出力理由を返す
- `POST /api/v2/outputs/spread-preview`: 並べ替えの案を保存せずに試算し、出力先ごとの物理ページ（頁・左右・見開き）を before/after で返す。ページ構成の「位置を指定して移動」が確定前の見比べに使う。要求と検証は`POST /api/v2/items/reorder`と同じ
- `POST /api/v2/outputs/build`: 資料パックの提出用PDFを出力するジョブを開始し、`job`（`jobId`・`status`・`total`）を返す。HTTPサーバーはリクエストを直列に処理するため、出力を要求の中で完結させると進捗ポーリングも中止も受け付けられない。実処理は子プロセスで走る
- `GET /api/v2/outputs/build/status?jobId={jobId}`: 出力ジョブの進捗（`percent`・`completed`/`total`・`phase`・`currentTargetName`）と結果（`built`・`skipped`・`errors`）を返す
- `POST /api/v2/outputs/build/cancel`: 出力ジョブの中止を要求する。作成中の1冊は書き上げてから停止し、残りの出力先は作成しない
- `POST /api/v2/outputs/publish`: 最新の提出用PDFを共有発行する。出力を確定したときにPDFのサイズ・更新日時・SHA-256を`outputIntegrity`として記録し、発行の前にハッシュを、共有側へコピーした後にもハッシュを照合する。出力後に変更・上書きされたPDF、記録の無い（以前の版で出力した）PDFは発行せず、再出力を求める。readinessはサイズと更新日時が記録と違えば`needs-rebuild`（`staleReasons[].type`=`output-changed`）を返す。提出用PDFの出力ジョブの実行中は拒否し、出力と同じロックを持ってコピーする（`publish-all`も同じ）。本文に`includeChangeReports: true`を付けると、前回コピーした版から今の提出用PDFまでの変更点PDFを同じ発行フォルダーへ入れる（省略時は false で、今までどおり提出用PDFだけ）。詳しくは「変更点PDFの共有コピーへの同梱」
- `POST /api/v2/outputs/publish-all`: ページの入っている出力先（本体・補足など）の最新の提出用PDFを、1つの発行フォルダーへまとめて共有発行する。1冊でも最新でなければ何もコピーしない。応答の`files`に各PDFの`targetId`・`fileName`・`sharedPath`・`buildId`（コピーした出力のアーカイブ。組み込み一式は空）・`sha256`・`changeReport`（同梱した変更点PDF。無ければ null）を返す。本文の`includeChangeReports`は`publish`と同じ
- `GET /api/v2/outputs/archives?packId={packId}`: 自動保存された最終出力を返す。各行に`publishedAt`・`folderName`（その版を最後に共有用フォルダーへコピーした日時と発行フォルダー名。コピーしていなければ空）、`isLastPublished`（その出力先で最後にコピーした版）、`isCurrent`（今の提出用PDF。出力状態の`buildId`と`outputIntegrity.sha256`がアーカイブと一致し、出力フォルダーのPDFが出力後に変わっていないときだけ true）、`hasPageMap`（manifest に頁・左右の記録`physicalPages`がある）を含む
- `GET /api/v2/outputs/changes?packId={packId}&targetId={targetId}[&fromBuildId=&toBuildId=]`: 出力先ごとに、前回共有用フォルダーへコピーした版（無ければ1つ前の出力）と今の提出用PDFの違いを、ページ番号・左右つきで返す。詳しくは「出力先ごとの変更点」
- `POST /api/v2/outputs/changes/report`: 変更点を確認用のPDF（変更点PDF）にするジョブを始める（Edge のヘッドレス印刷）。`GET /api/v2/outputs/changes/report/status?jobId=`で状態を読み、`POST /api/v2/outputs/changes/report/file`で開く／保存先フォルダーを開く。詳しくは「変更点PDF」

`PATCH`の例：

```json
{
  "ownerDepartment": "経理部",
  "required": true,
  "defaultTargetId": "appendix",
  "excelSheetSelectionMode": "numeric-only"
}
```

`numeric-only`のPDF作成では、非該当または非表示になった既存Excelページを、復元ポイント（`sheet-selection`）を残してページ構成から外す。未振り分けには残さない。半角数字のシートを改名して非該当になった場合も、旧名のページを改名後のシートへ引き継がずに外す。シート名を半角数字へ戻すか再表示して再度PDF化すると、シート名から決まる同じ`pageId`の新しいページとして戻る。以前の版が`volume=none`・`sheetSelectionExcluded=true`で残したページ（変換PDFなし、未振り分けのもの）は、起動時の読み込みで復元ポイント（`sheet-selection-cleanup`）を残して整理する。`all-visible`へ戻すと再び表示シートをPDF化する。Excelから変換PDFへ書き出す処理自体は、シートと分割PDFの対応を保つためタブ順で行う。ページ構成の初期配置と、画面上の「半角数字順に並べ直す」は、半角数字シートを数値順に配置する。

既存のECM/BOD/DMM操作は引き続きV4 APIと互換です。組み込みテンプレートの`acceptedSourceTypes`は`excel / word / pdf / powerpoint`です。

### 利用者定義ひな形

利用者定義ひな形は言語別の`dataDir\<language>\templates`にJSON保存します。`displayName`、1つ以上の`acceptedSourceTypes`、1〜12件の任意出力先、ルール、出力既定値を指定できます。`targetId`は小文字英数字で始まる48文字以内の小文字英数字・`_`・`-`とし、ひな形内で重複できません。

```json
{
  "displayName": "監査レビュー",
  "description": "監査部提出用",
  "acceptedSourceTypes": ["excel", "word", "powerpoint", "pdf"],
  "targets": [
    { "targetId": "main", "displayName": "監査報告", "required": true },
    { "targetId": "appendix", "displayName": "証憑", "required": false }
  ],
  "rules": {
    "newItemDestination": "main",
    "retainManualOrder": true,
    "blockBuildWhenRequiredSourceIsStale": true,
    "blockBuildWhenRequiredSourceFailed": true
  },
  "output": {
    "fileNamePattern": "{packName}_{targetName}_{yyyyMMdd}.pdf"
  }
}
```

更新のたびに`templateVersion`が増えます。資料パックは作成時点のひな形を`templateConfig`として保持するため、後からひな形を編集しても既存パックの出力先名や設定は自動では変わりません。`template-upgrade`で差分を確認してから明示適用できます。新版で削除された出力先のページは未振り分けへ移動します。使用中のひな形と組み込みひな形は削除できません。画面では利用者定義ひな形を複製し、JSONとして書き出し・読み込みできます。

`rules.newItemDestination`は、原稿の`defaultTargetId`が`unassigned`の場合だけ使用します。既存ページの手動配置は再変換後も保持します。必須原稿では、ファイル欠落を`required-source-missing`、未変換を`required-source-not-rendered`、更新を`required-source-stale`、変換失敗を`required-source-failed`としてreadinessに返します。

### 資料パックの作成・複製・保管

新規作成では`displayName`が必須です。`templateId`を省略すると`builtin-generic-department-pack`を使用します。

```json
{
  "displayName": "月次部門報告",
  "templateId": "builtin-generic-department-pack",
  "settings": {
    "documentTitle": "月次部門報告書",
    "outputFileNamePattern": "{packName}_{targetName}_{yyyyMMdd}.pdf"
  }
}
```

複製はテンプレートと仕上げ設定だけを引き継ぎ、登録原稿、ページ構成、出力PDF、履歴は複製しません。`displayName`を省略すると「元の名前 (コピー)」のような重複しない名前を付けます。

名称変更は従来の仕上げ設定と同じ`PATCH`へ`displayName`を渡します。設定を`settings`内にまとめる形式と、従来どおり直下へ置く形式の両方を受け付けます。

```json
{
  "displayName": "四半期部門報告",
  "settings": {
    "documentSubtitle": "2026年度 第1四半期"
  }
}
```

アーカイブは削除ではなく、通常一覧から隠す操作です。データは保持され、`restore`で戻せます。現行の固定UIとの互換性を守るため、組み込みECM / BOD / DMMはアーカイブできません。

### ページ構成の保存履歴と復元

ページの移動、並べ替え、タイトル・番号表示・ページ範囲などの設定変更では、変更直前の構成を資料パック単位で自動保存します。履歴にはレイアウト項目だけを保存し、原稿ファイル、変換PDF、更新検出状態には触れません。

```json
{
  "packId": "pack_0123456789abcdef",
  "snapshotId": "20260807T130000.000_abcd1234"
}
```

`restore/preview`は現在も存在するページへの適用数、過去にだけ存在するページ、現在にだけ存在するページ、出力先が変わるページを返します。`restore`は同じ`packId`のページにだけ適用し、復元の直前にも新しい履歴を保存します。組み込みECM / BOD / DMMは従来の`/api/layout/*`と互換です。

### 差分レビュー状態

`POST /api/history/diff-review`は、表示中の比較に含まれる1つのページ項目を確認済み・未確認へ切り替える。`workbookId`、任意比較の場合は`fromSnapshotId`と`toSnapshotId`（版を指定して開いた比較なら`fromVersionId`と`toVersionId`も）、画面が取得した`baselineVersionId`と`currentVersionId`、`sheetKey`、`confirmed`を送る。確認状態は利用者のローカル設定へ保存され、snapshot、render version、差分アルゴリズム版のいずれかが変わった比較へは引き継がれない。

### 資料パックの提出・承認・監査履歴

提出用PDFを作成した後、`submit`で必須出力の現在の指紋を固定してレビューへ提出します。`approve`は提出後に内容が変わっていない場合だけ成功します。原稿、ページ構成、仕上げ設定などが変わって指紋が一致しなくなると状態は`stale`になり、再出力・再提出が必要です。`request-changes`には差し戻し理由が必須です。`reopen`は下書きへ戻します。

```json
{
  "action": "request-changes",
  "note": "補足資料の数値根拠を追記してください",
  "actor": "reviewer-name"
}
```

状態変更は`events`へ最大100件保存され、操作、操作者、日時、コメント、提出時指紋を追跡できます。ダッシュボードの`packProgress`は、出力完了後も`review-draft / in-review / review-changes / review-stale / complete`を区別します。

## GET /api/state

設定、workbook、page、volume状態、サマリー、PDF.js配置状態を返します。`pdfjsPresent` / `pdfjsMode` はファイル配置の診断値であり、V5の画面プレビューがPDF.jsを使用していることを示すものではありません。

主な追加項目：

`excelPrintProfileVersion`、`wordRenderProfileVersion`、`pdfImportProfileVersion`は、形式ごとに変換PDFの再作成が必要かをUIが判定するためのプロファイル版です。

```jsonc
{
  "structure": {
    "schemaVersion": 2,
    "volumes": {
      "ja-main|ecm": {
        "status": "built",
        "builtFingerprint": "...",
        "lastBuiltAt": "...",
        "outputPdf": "...",
        "outputPdfExists": true,
        "staleReasons": []
      }
    }
  },
  "packTemplates": [
    { "templateId": "builtin-ecm", "packId": "pack_ecm", "acceptedSourceTypes": ["excel", "word", "powerpoint", "pdf"] }
  ],
  "packs": [
    { "packId": "pack_ecm", "displayName": "ECM", "category": "ecm", "workflowAvailable": true }
  ],
  "finalReadiness": {
    "ecm": {
      "volumes": {
        "ja-main": {
          "canBuild": true,
          "pageCount": 24,
          "displayState": "built",
          "builtFingerprint": "...",
          "currentFingerprint": "...",
          "outputPdfExists": true,
          "blockers": [],
          "staleReasons": []
        }
      }
    }
  }
}
```

`outputPdfExists`は永続化せず、レスポンス生成時に算出します。

## POST /api/diagnostics/run

Office、Java、PDFBox、PDF.js、保存先の動作環境を診断します。Excel/Wordはレジストリ登録だけでなくCOM起動と版番号を確認します。Officeを起動できない非対話セッションでもAPI自体は失敗せず、`ready / limited / blocked`と利用者向け対応を返します。

```jsonc
{
  "status": "limited",
  "summary": "PDF原稿は処理できます。Office原稿には対応が必要です。",
  "office": {
    "minimumSupportedMajor": 16,
    "excel": { "registered": true, "available": false, "status": "unavailable", "errorCode": "windows-session-unavailable" },
    "word": { "registered": true, "available": false, "status": "unavailable", "errorCode": "windows-session-unavailable" }
  },
  "runtime": {
    "java": { "ready": true },
    "pdfbox": { "ready": true },
    "pdfjs": { "ready": true }
  },
  "storage": { "configured": true, "dataFreeBytes": 100000000000, "outputFreeBytes": 100000000000 }
}
```

`status`が`blocked`のときは、次の一手を案内する`advice`フィールドが追加されます（旧クライアントは無視できる追加分）。配布形態は配布物の根にある`release-manifest.json`（package-release.ps1が出力）で判別します。`portableJavaIncluded: true`（flavor `shared-folder-offline` / `オフライン完結版`）の配布では壊れた配布として`kind: "redistribute"`を返し、`portableJavaIncluded: false`（flavor `オンライン導入版`）では導入コマンドで修復可能として`kind: "install-online"`を返します。マニフェストを持たない開発ツリーでは`app\lib\java`フォルダーの有無で推定します。マニフェストがあっても読み取れない場合は誤った案内を避けるため`advice`自体を省略し、従来どおりの応答に劣化します。

```jsonc
// kind: "install-online" の例
{
  "status": "blocked",
  "advice": {
    "kind": "install-online",
    "title": "PDFの作成と結合に必要なソフトがこのPCに入っていません。",
    "steps": [
      "インターネット接続が必要です。",
      "ReportBinderフォルダーの中にある app\\tools\\install-thirdparty.cmd を実行してください。",
      "終わったら、もう一度「動作環境を診断」を実行してください。"
    ],
    "toolsDir": "C:\\ReportBinder\\app\\tools",
    "hasToolsDir": true
  }
}

// kind: "redistribute" の例（オフライン完結版。toolsDir / hasToolsDir は持たない）
{ "kind": "redistribute", "title": "...", "steps": ["..."] }
```

## POST /api/diagnostics/open-tools

診断の修復案内（`kind: "install-online"`）で案内する `AppRoot\tools` フォルダーをエクスプローラーで開きます。開く先はサーバー側で固定であり、要求本文は参照しません。

```json
{}
```

応答:

```json
{ "ok": true, "path": "C:\\ReportBinder\\app\\tools" }
```

フォルダーが無い場合は`404`と`error: "tools-folder-missing"`を返します。

## GET /api/version-check

共有フォルダー配布において、配布元により新しい版が公開されているかを返します。サーバーは起動2分後、以後30分ごとにバックグラウンドで配布元の`runtime-version.json`を確認し、結果をメモリに保持します。このAPIは要求時に共有フォルダーへアクセスせず、保持した結果だけを返すため、配布元が応答しなくても即座に応答します。

```jsonc
// 確認済みで、より新しい版がある例
{
  "ok": true,
  "checkedAt": "2026-08-22T17:00:00+09:00",
  "runningVersion": "2026.08.01.1",
  "latestVersion": "2026.08.22.1",
  "updateAvailable": true
}

// 未確認・確認失敗・開発ツリーからの直接起動の例（fail-silent）
{
  "ok": true,
  "checkedAt": null,
  "runningVersion": "2026.08.01.1",
  "latestVersion": "",
  "updateAvailable": false
}
```

`updateAvailable`は、双方の版が`x.y.z.w`の4部構成の数字で、配布元の版が実行中より厳密に新しいときだけ`true`になります。それ以外の形式や、版ファイルの欠落・読み取り失敗では例外を表に出さず、`checkedAt: null`として扱います。

## POST /api/heartbeat

```json
{ "clientId": "..." }
```

## POST /api/client/close

```json
{ "clientId": "..." }
```

## POST /api/submission/select-and-start

提出フォルダ選択ダイアログを開き、`_reportbinder`と`出力`を自動設定します。

```json
{ "initialDir": "C:\\Reports" }
```

現在値は`%LOCALAPPDATA%\ReportBinder\config.json`へ保存します。

## POST /api/dialog/folder

```json
{ "title": "提出フォルダを選択", "initialDir": "C:\\Reports" }
```

## POST /api/paths

```json
{
  "submissionDir": "C:\\Reports\\提出",
  "dataDir": "C:\\Reports\\提出\\_reportbinder",
  "outputDir": "C:\\Reports\\提出\\出力"
}
```

## GET /api/submission-files

提出フォルダ直下の原稿一覧を返します。現在のUIが登録対象として表示するのは`.xlsx`、`.xlsm`、`.docx`、`.pptx`、`.pdf`です。Word/PowerPointの一時ファイル`~$*`、`.docm`、`.pptm`、子フォルダ内のファイルは候補に含めません。

- `scannedAt`: 一覧を取得した日時
- `files[].sourceType`: `excel`、`word`、`powerpoint`または`pdf`
- `files[].adapterId`: 使用する原稿アダプター
- `files[].modifiedAt`: ファイルサーバーから再取得した原稿の最終保存日時
- `files[].modifiedAtUtcTicks`: 更新判定・調査用のUTC ticks

## POST /api/workbooks/register-batch

categoryは必須です。

```json
{
  "category": "ecm",
  "relativePaths": ["FY160-4Q_ECM_J_00_Cover.xlsx", "department-report.docx", "briefing.pptx", "department-appendix.pdf"]
}
```

## POST /api/workbooks/register

```json
{
  "category": "ecm",
  "relativePath": "FY160-4Q_ECM_J_00_Cover.xlsx"
}
```

## POST /api/workbooks/unregister

```json
{ "workbookId": "ecm-fy160-4q-ecm-j-00-cover-..." }
```

## POST /api/workbooks/render/start

選択分：

```json
{
  "category": "ecm",
  "workbookIds": ["..."],
  "excelSheetSelectionMode": "numeric-only"
}
```

PDF必要分：

```json
{
  "category": "ecm",
  "onlyUpdated": true,
  "excelSheetSelectionMode": "all-visible"
}
```

戻り値のjobIdで`/api/jobs/status`をポーリングします。

## GET / POST /api/jobs/status

```json
{ "jobId": "job_20260728_120000_abcdefgh" }
```

## POST /api/jobs/cancel

```json
{ "jobId": "job_20260728_120000_abcdefgh" }
```

実行中のPDF作成へ安全な中止要求を送ります。現在処理中の原稿は完了させ、次の原稿へ進む前に`cancelled`で終了します。完了済みの変換PDF、ページ構成、履歴比較資産は保持します。`/api/jobs/status`では受付後に`cancelRequested: true`を返します。

完了結果の各workbookには、ページ追加情報を含みます。

```jsonc
{
  "sheetSync": {
    "addedCount": 5,
    "insertedInOrderCount": 2,
    "insertedAtEndCount": 3,
    "insertedAtEndPageIds": ["...", "...", "..."]
  }
}
```

## POST /api/scan-updates

登録済みExcel・Word・PowerPoint・PDF原稿の更新を確認します。Word/PowerPoint/PDFの差し替えは`source-updated`となり、再変換後に物理ページ・スライドの追加・削除も同期します。

```json
{ "forceHash": false }
```

## POST /api/pages/reorder

categoryは必須です。

```json
{
  "category": "ecm",
  "volumes": {
    "ja-main": ["page-1", "page-2"],
    "ja-appendix": ["page-3"],
    "none": []
  }
}
```

実際に順序・割当が変わったvolumeだけを更新対象にします。送ったvolumeに今あるページが欠けている場合や、未知・重複のIDがある場合は400を返し、何も保存しません（`POST /api/v2/items/reorder`と同じ検証）。

ページ構成画面の「半角数字シートをまとめる」は、選択した出力先へ`^[0-9]+$`のシート名を文字列として安全に数字順で配置します。Excel以外のページはアンカーとして保持し、非該当Excelと既存の非表示シートページは`none`へ移します。変更がない場合はリクエストも履歴スナップショットも作りません。

## POST /api/pages/sort-by-numeric-sheet

categoryは必須です。

```json
{
  "category": "ecm",
  "volumes": ["ja-main", "ja-appendix", "none"]
}
```

volumeを省略した場合は本体・補足だけを対象にします。割当は変えず、未振り分けを含む指定レーンの中だけを、半角数字シートの数値順（桁数→文字列順、同値はファイル順・ファイル名・シート位置・pageId）で並べ替えます。Excel以外や半角数字以外のページはアンカーとして同じレーン内に残します。既存クライアント向けに`POST /api/pages/sort-by-sheet`も同じ動作の別名として受け付けます。

画面上のページカードはWindows/macOSのファイル一覧と同じ選択モデルです。通常クリックは1枚だけ、Ctrl/Cmdクリックは追加・解除、Shiftクリックは範囲選択、Ctrl+Aは表示中を選択します。マウス・ペンは編集部以外のカード面を直接ドラッグでき、タッチでは縦スクロールを優先します。カードのダブルクリックまたはEnterでプレビューを開き、ドロップ後は移動したカードだけを選択状態にします。

## POST /api/pages/update

categoryは必須です。

```json
{
  "category": "ecm",
  "pageId": "...",
  "title": "CA1 売上実績",
  "numberingMode": "visible",
  "numberingManual": true
}
```

タイトルだけの変更では最終PDFを再出力扱いにしません。

## POST /api/pages/confirm

互換用です。categoryは必須です。

```json
{ "category": "ecm", "pageId": "...", "action": "confirm" }
```

## GET / POST /api/file

content-pdfを取得します。

POST例：

```json
{
  "pageId": "...",
  "workbookId": "...",
  "sheetName": "1",
  "contentPdf": "content-pdf\\...\\1.pdf"
}
```

## GET /api/final/readiness?category=ecm

categoryは必須です。省略・不正値はHTTP 400です。

```jsonc
{
  "ok": true,
  "category": "ecm",
  "volumes": {
    "ja-main": {
      "canBuild": false,
      "pageCount": 24,
      "status": "needs-rebuild",
      "displayState": "blocked",
      "blockers": [
        {
          "code": "stale-content",
          "pageTitle": "3 業績サマリー",
          "workbookName": "...xlsx",
          "message": "元Excelが更新されています。先にPDF作成してください。"
        }
      ]
    }
  }
}
```

`displayState`は`blocked`、`needs-rebuild`、`output-missing`、`not-built`、`built`です。

各出力先の`snapshot.physicalPages`（`GET /api/v2/outputs/readiness`、`/api/state`の`finalReadiness`も同じ）は、提出用PDFの1物理ページを1要素で返します。頁は出力先ごとに1から数えます。変換PDFの無いページは含みません。

| 項目 | 値 |
| --- | --- |
| `outputPageNumber` | 出力先内の物理ページ番号（1から） |
| `side` | `right`（奇数頁）／`left`（偶数頁）。製本のパンチ穴シフト（`ReportPdfComposer`）と同じ規則 |
| `spreadIndex` | 見開きの番号。`floor(outputPageNumber/2)`。1頁は見開き0の右、2・3頁は見開き1 |
| `pagePartIndex`／`pagePartCount` | 同じ`pageId`の中で何枚目か（1から）／使用ページ範囲を反映した全枚数 |
| `sourcePdfPageNumber` | 変換PDF内のページ番号 |

`side`・`spreadIndex`・`pagePartIndex`・`pagePartCount`は表示用で、入力の指紋（`fingerprint`）には含めません。追加しても「再出力が必要」の判定は変わりません。

## POST /api/v2/outputs/spread-preview

並べ替えの案を当てたときの物理ページを、確定前に before/after で見比べるための試算です。**何も保存しません**（`structure.json`、ページ構成の保存履歴、出力先の「再出力が必要」の印のどれも変えません）。要求は`POST /api/v2/items/reorder`と同じ形（`volumes`または`targets`）で、同じ検証をします。空のレーン、ページの欠落、この一式に無いID、同じIDの重複、不正な出力先は400です。`baseLayout`が今の並びの指紋と違えば409 `structure-conflict`を返します（古い並びに対して試算しないため）。

```json
{ "packId": "pack_...", "baseLayout": "…", "volumes": { "ja-main": ["page_a", "page_c", "page_b"], "ja-appendix": [], "none": [] } }
```

```json
{
  "ok": true,
  "apiVersion": 2,
  "preview": {
    "packId": "pack_...",
    "layoutFingerprint": "…",
    "changedVolumes": ["ja-main"],
    "before": { "main": { "targetId": "main", "volume": "ja-main", "pageCount": 6, "itemCount": 3, "physicalPages": [], "unknownPageIds": [] }, "appendix": { } },
    "after":  { "main": { "targetId": "main", "volume": "ja-main", "pageCount": 7, "itemCount": 4, "physicalPages": [], "unknownPageIds": ["page_x"] }, "appendix": { } }
  }
}
```

- `before`・`after`は資料パックの出力先ごと（キーは`targetId`）で、`physicalPages`の各要素は readiness の`snapshot.physicalPages`と同じ形です。`after`は同じ案を`POST /api/v2/items/reorder`で確定したあとの readiness と一致します。
- `before`も、readiness ではなく要求を受けた時点の構成から作り直します。readiness は並べ替えのあと遅れて更新されるためです。
- `unknownPageIds`は、出力先に入っているのに変換PDFが無く頁数がわからないページです。`physicalPages`には出てこないので、それより後ろの頁・左右は確定値ではありません。
- `layoutFingerprint`は試算に使った並びの指紋です。確定の`baseLayout`にそのまま使えます。
- 変換PDFのページ数は、パス・サイズ・更新時刻をキーにしたキャッシュ（`Get-PdfPageCount`）から読みます。

## POST /api/final/build

volumeとcategoryは必須です。

```json
{ "volume": "ja-main", "category": "ecm" }
```

組版前後にfingerprintを比較します。組版中にページ順またはcontent-pdfが変わった場合、既存最終PDFを差し替えません。

manifest、一時PDF、lockはcategory込みです。

```text
manifest_ja-main_ecm.json
~building_ja-main_ecm.pdf
volume_ja-main_ecm.lock
```

## GET /api/final/file

volumeとcategoryは必須です。

```text
/api/final/file?volume=ja-main&category=ecm&token=...
```

## POST /api/final/file

```json
{ "volume": "ja-main", "category": "ecm" }
```

## GET /api/history/diff-detail

最新PDFに紐づく比較メタデータ、原稿内項目一覧、ページ対応、差分領域、生成状態を返します。
`fromSnapshotId`と`toSnapshotId`を両方指定した場合は、同じ登録済み原稿の任意の
保存済み2版を比較します。片方だけの指定、同一版、履歴外のIDは拒否します。

`sheets`はV4互換のコレクション名です。各項目の`beforeSheetName`と`afterSheetName`は
比較元・比較先で実際に参照する項目キー、`matchConfidence`はページ対応の確信度、
`matchMethod`は対応方法です。Word/PDFで途中にページが挿入された場合は、完全一致する
前後ページを基準に`Page 2 → Page 3`のように対応付けます。対応を安全に確定できない区間は
`kind: "unknown"`として返し、差分位置を推測しません。`sheetName`はV4クライアント向けに残します。

```text
/api/history/diff-detail?workbookId=wb_...&token=...
/api/history/diff-detail?workbookId=wb_...&fromSnapshotId=...&toSnapshotId=...&token=...
/api/history/diff-detail?workbookId=wb_...&fromSnapshotId=...&toSnapshotId=...&fromVersionId=...&toVersionId=...&token=...
```

`fromVersionId`と`toVersionId`（任意。両方そろえて、`fromSnapshotId`・`toSnapshotId`と一緒に指定する）は、比べる描画の版です。
指定すると、その版どうしだけを比べ、別の版には代えません（その snapshot に存在しない版、画像ハッシュと
変換PDFがそろっていない版なら`status: "unavailable"`）。省略時は今までどおり、それぞれの snapshot の
いちばん新しい使える版を使います。提出用PDF一式の変更点から開く「くわしく見る」は、アーカイブの
`metadata.sourceWorkbooks[].versionId`（提出用PDFに実際に入った版）を指定し、同じ snapshot を後で
描き直した版と比べないようにします。版を指定したときは、同じ snapshot の別の版どうしも比べられます。
片方だけの指定や、履歴版なしの版の指定は 400 です。応答の`comparison.baselineVersionId`・
`comparison.currentVersionId`は、実際に比べた版です。`POST /api/history/diff-review`も、同じ指定を
本文の`fromVersionId`・`toVersionId`で受け付けます。

`status`は`not-generated`、`generating`、`ready`、`unavailable`、`failed`です。
領域座標は0～1の正規化値で、各領域は`modified`、`added`、`removed`のいずれかです。
判定不能のページでは推測した領域を返しません。

応答の`message`（比較できない理由など）とページ対応の`mappingMessage`は、比較画面で左右のPDF画像に重ねず、
画像の表示欄より上に通常の行として表示します。画面がブラウザで差分を計算したときの
計算方法の説明（文字の照合・位置ずれの補正・描画ノイズの除外など）は表示しません。
このAPIの応答の形は変わりません。

## POST /api/history/diff/prepare

未生成の差分画像をバックグラウンドジョブとして開始します。作成済みのキャッシュが
現在版・比較基準版・アルゴリズム版と一致する場合は再利用します。

```json
{ "workbookId": "wb_..." }
```

履歴2版の差分を生成する場合:

```json
{
  "workbookId": "wb_...",
  "fromSnapshotId": "20260701T090000.000_...",
  "toSnapshotId": "20260730T090000.000_..."
}
```

履歴比較は自動比較の`comparison-baseline.json`を変更せず、選択した2版専用の
キャッシュへ保存します。自動比較の最新状態や変更バッジには影響しません。

返された`jobId`は既存の`GET /api/jobs/status`で確認します。ダイアログを閉じても処理は継続します。

## GET /api/history/diff-page

差分詳細manifestに登録されたページ画像だけを返します。

```text
/api/history/diff-page
  ?workbookId=wb_...
  &currentSnapshotId=...
  &baselineSnapshotId=...
  &sheetKey=sheet-...
  &pageNumber=1
  &asset=before
  &scope=history
```

`asset`は`before`、`after`、`before-mask`、`before-overlay`、`mask`、`overlay`のいずれかです。
`scope`は`automatic`または`history`です。識別子は対応する履歴・比較結果と照合し、
ファイルパスは受け付けません。

## Presets

categoryは次のいずれかです。

```text
ecm / bod / dmm
```

最終PDF、readiness、前回PDF取得、ページ構成のcategory省略は許可しません。

## PATCH /api/v2/packs/{packId}

資料パックごとの最終PDF仕上げ設定を保存します。変更後は本体・補足とも再出力対象になります。

```json
{
  "documentTitle": "月次経営会議資料",
  "documentSubtitle": "2026年8月",
  "outputFileNamePattern": "{projectId}_{targetName}_{yyyyMMdd}.pdf"
}
```

ファイル名には`{projectId}`、`{packName}`、`{targetName}`、`{yyyyMMdd}`を使用できます。`{targetName}`は必須です。含まない型は一式の設定・ひな形の保存時に400で拒否します。保存済みの古い型で出力先どうし（または他の一式の出力PDF）と同じパスになる場合は、出力を開始せずに失敗させます。

## 任意資料パックの原稿・ページ・出力

任意資料パックも、組み込みパックと同じExcel・Word・PowerPoint・PDFアダプターを使用します。Excelは`.xlsx`と`.xlsm`、Wordは`.docx`、PowerPointは`.pptx`を受け付け、Office Automation Securityを強制してマクロを実行しません。v2 APIでは`category`ではなく、作成時に返された`packId`を指定します。

```text
POST  /api/v2/sources/register-batch
POST  /api/v2/sources/render/start
POST  /api/v2/items/reorder
PATCH /api/v2/items/{itemId}
GET   /api/v2/layout/snapshots?packId=pack_...
POST  /api/v2/layout/restore/preview
POST  /api/v2/layout/restore
GET   /api/v2/outputs/readiness?packId=pack_...
POST  /api/v2/outputs/spread-preview
POST  /api/v2/outputs/build
POST  /api/v2/outputs/file
POST  /api/v2/outputs/publish
POST  /api/v2/outputs/publish-all
GET   /api/v2/outputs/archives?packId=pack_...
GET   /api/v2/outputs/changes?packId=pack_...&targetId=main[&fromBuildId=...&toBuildId=...]
POST  /api/v2/outputs/changes/report
GET   /api/v2/outputs/changes/report/status?jobId=job_...
POST  /api/v2/outputs/changes/report/file
```

原稿登録:

```json
{ "packId": "pack_...", "relativePaths": ["部門資料.xlsx", "説明.docx", "説明会.pptx", "添付.pdf"] }
```

ページ構成は`volumes`（例: `ja-main` / `ja-executive-summary` / `none`）または`targets`（例: `main` / `executive-summary` / `unassigned`）で指定できます。使用できるキーは資料パックが保持するひな形の`targets`から決まり、固定の本体・補足には限定されません。

```json
{
  "packId": "pack_...",
  "targets": {
    "main": ["item-1", "item-2"],
    "appendix": ["item-3"],
    "unassigned": []
  }
}
```

最終PDFの出力:

```json
{ "packId": "pack_...", "targetIds": ["main", "appendix"] }
```

共有発行:

```json
{ "packId": "pack_...", "targetId": "main" }
```

最新状態の最終PDFだけを共有発行できます。提出フォルダー直下の一時フォルダーへコピーし、サイズ検証後に日時・言語・利用者名付きフォルダーへ切り替えます。

最終PDFの作成時には、PDF、manifest、SHA-256、使用した原稿版と変換環境を`exports/archive/<packId>/...`へ保存します。`GET /api/v2/outputs/archives`は新しい順にアーカイブ情報を返します。

共有発行のたびに、出力先ごとの発行記録を`exports/publish-records/<packId>/<volume>/<時刻ID>.json`へ1件ずつ書きます（組み込み一式は`<packId>`の代わりに`ecm`など）。中身は`final.published`と同じ項目（`packId`・`targetId`・`volume`・`buildId`・`sha256`・`fileName`・`folderName`・`sharedFolder`・`destination`・`publishedAt`・`userName`・`changeReport`（同梱した変更点PDFの`fromBuildId`・`toBuildId`・`fileName`・`destination`・`sha256`。無ければ null）ほか）です。「前回コピーした版」（`Get-LastPublishedBuildId`）は、その出力先のいちばん新しい記録の`buildId`で、古い記録へはさかのぼりません。履歴イベントの書き込み（`Write-HistoryEvent`）は廃止済みで何も書かないため、この記録が唯一の手がかりです。記録を書けなかったときもコピー自体は成功として返し、`publishRecordError`に理由を入れます。

サーバー内部の`Get-FinalArchiveRecord`はアーカイブ1件を読み、`sha256.txt`・`metadata.json`の`outputPdfSha256`・`final.pdf`の実際のハッシュが一致しないもの（すり替えられた版）は読みません。`buildId`・`volume`は`Assert-SafeStorageSegment`で確かめます。

### 出力先ごとの変更点（`GET /api/v2/outputs/changes`）

`GET /api/v2/outputs/changes?packId={packId}&targetId={targetId}[&fromBuildId={buildId}][&toBuildId={buildId}]`は、同じ出力先の保存済み出力（アーカイブ）2件を比べ、ページごとの違いを返します。応答は`{ ok, apiVersion: 2, changes: {...} }`です。manifest と原稿の見た目のハッシュを読むだけなので、その場で計算して返します（画像は作りません）。

比べる版の決め方:

- 比べる先（`to`）: 省略時は今の提出用PDF（`/api/v2/outputs/archives`の`isCurrent`の行）。無ければ最後に出力した版にし、`notes`に`not-current`を入れます。
- 比べる元（`from`）: 省略時は、その出力先で前回共有用フォルダーへコピーした版（発行記録の`buildId`。最後の記録に`buildId`が無ければ、記録の`sha256`と`outputPdfSha256`が同じアーカイブ）。まだコピーしていない（`not-published`）、前回コピーした版が今の版と同じ（`same-as-published`）、前回コピーした版のアーカイブが無い（`published-missing`）ときは、比べる先の1つ前の出力と比べます。
- `fromBuildId`だけを指定すると、その版と今の提出用PDFを比べます（保存済み出力の「今の版と比べる」）。新しい版を比べる元に選んだときは古い→新しいの向きに入れ替え、`swapped: true`と`notes`の`swapped`を返します。
- 指定した版がこの出力先の保存済み出力に無い（別の出力先の版を含む）、同じ版どうし、パスとして不正な識別子、一式に無い出力先は 400 です。

`changes`の主な項目:

| 項目 | 内容 |
| --- | --- |
| `status` | `ready`（比べた）／`no-output`（まだ出力していない）／`no-baseline`（はじめての出力で比べる元が無い）／`unavailable`（アーカイブのPDFが記録と合わず読めない）／`unsupported`（組み込み一式。出力の版を保存していない） |
| `message` | `status`が`ready`以外のときの説明 |
| `defaultPair` / `fromReason` / `swapped` | 既定の組み合わせか。比べる元の決め方（`published`＝前回コピーした版、`previous`＝1つ前の出力、`selected`＝指定、`none`） |
| `from` / `to` | `buildId`・`builtAt`・`pageCount`（出力ページ数）・`itemCount`・`outputFileName`・`outputPdfSha256`・`hasPageMap`・`publishedAt`・`folderName`・`isLastPublished`・`isCurrent`・`isLatest`（`to`だけ。今の提出用PDFが原稿の最新内容で出力済みか） |
| `hasSides` | 両方の版に頁・左右の記録（`physicalPages`）があり、左右の入れ替わりを比べたか |
| `complete` | 見た目のハッシュがそろわず判定を保留したページ（`maybe-content`・`unavailable`）が無いか |
| `counts` | `content`・`sameLook`・`maybeContent`・`unavailable`・`added`・`removed`・`moved`・`shifted`（`moved`のうち前のページの増減でずれただけのもの）・`side`・`renamed`・`unchanged`・`changed`（変わったページの数。`unchanged`と`sameLook`だけのページは数えない）・`partCountChanged`（ページ数が変わった原稿の数）・`total` |
| `pages[]` | 今回のページ番号順（なくなったページは、前回その直前にあったページの後ろ）。`rowKey`（行の安定したキー。`<項目のキー>#<項目内の順番>`、例`item:it_1#2`。出力ページ番号を含まないので、前のページの増減では変わらない）・`toPage`・`toSide`・`toSpreadIndex`・`fromPage`・`fromSide`・`fromSpreadIndex`・`toSourcePageNumber`/`fromSourcePageNumber`（原稿の変換PDFの中で何ページ目か。`GET /api/history/diff-detail`の`pages[].afterPageNumber`／`beforePageNumber`と同じ数え方。無い側は`null`）・`fromSourceId`/`fromSheetName`（前回のページの原稿IDとシート名。`sourceId`・`sheetName`は今回のページ（無ければ前回）のもので、原稿の差し替え・シート名の変更があると前回と違う。変更点PDFの前回の画像に使う。前回が無い行は空。`summary.json`が作り直されるまでの古い結果には無い）・`itemId`・`pageId`・`sourceId`・`sourceName`・`title`・`fromTitle`・`sheetKey`（単位の`sourceKey`。例`worksheet:売上`）・`sheetName`（原稿単位の比較でのシート名）・`pagePartIndex`・`fromPartCount`・`toPartCount`・`partCountChanged`・`kinds[]`・`shiftOnly`・`changedRatio`/`regionCount`（一覧は画像の比較を行わないため常に`null`。枠は変更点PDFのページごとの詳細に載せる）・`fromSnapshotId`/`fromVersionId`・`toSnapshotId`/`toVersionId`・`detailAvailable`（同じ原稿の別の版どうしで、原稿単位の比較を開ける）・`message` |
| `notes[]` | `{ code, message }`。上記のほか`not-latest`・`no-page-map`（古い形式の版で左右を比べていない）・`no-changes`・`some-unavailable` |
| `report` | 変更点PDF（`{ exists, reportId, fileName, createdAt, attachable, publishable, reason, detailSource }`）。`detailSource`は作成済みのときだけ入り、ページごとの詳細の枠の出どころ（`browser`＝すべてアプリの「くわしく見る」と同じ計算、`server`＝すべてサーバーの画素比較（予備）、`mixed`＝両方、`none`＝内容の変わったページの詳細が無い）。`status`が`ready`のときだけ調べ、`exists`は作成済みで記録と中身が合うとき true。`attachable`はこの組み合わせが「前回コピーした版 → 今の提出用PDF」（共有コピーへ同梱できる組み合わせ）のとき true、`publishable`は作成済みで、今コピーすれば同梱されるとき true（コピーのときと同じ`Resolve-PublishableChangeReport`で決める）。`reason`は`not-created`（まだ作っていない）・`outdated`（作ったものが今の比べた結果や今の提出用PDFと合わない）・`publishable`（作成済みで同梱できる）・`not-attachable`（作成済み。好きな2回の組み合わせや、今の提出用PDFでない版との組み合わせなので同梱しない）・`not-available`（比べられない）。詳しくは「変更点PDF」 |

`kinds[]`の値: `content`（原稿の版が違い、ページの見た目のハッシュも違う）、`same-look`（版は違うがハッシュが同じ）、`maybe-content`（版が違いハッシュが取れない）、`unavailable`（比べる版の記録が消えている）、`added`、`removed`、`moved`（出力ページ番号が違う）、`side`（記録された左右が違う）、`renamed`（ページ名が違う）、`unchanged`。

ページの対応は項目（`itemId`、無ければ`pageId`）ごとに、`pagePartIndex`の小さい方から順に取ります。原稿のページ数が変わったときは余りを`added`／`removed`にします。原稿の版（`metadata.sourceWorkbooks`の`snapshotId`・`versionId`）とページの位置が同じページは、ハッシュを見ずに内容が同じとみなします。頁・左右・台紙はアーカイブに記録された値をそのまま返し、数え直しません。`physicalPages`の無い古い版は`pages[]`の`sourcePageStart`〜`sourcePageEnd`から出力ページ番号だけを数え、左右・台紙は空（推定しない）にします。

画面での使い方: 提出用PDF画面の出力先カードの「前回コピーしたときからの変更点を見る」は`fromBuildId`・`toBuildId`を付けずに呼び（既定の組み合わせ）、保存済み出力の「今の版と比べる」は`fromBuildId`だけ、変更点の画面の「ほかの版と比べる」は両方を付けて呼びます。画面は`pages[]`の`toPage`/`toSide`/`fromPage`/`fromSide`をそのまま表示し、左右が空の行で左右を補いません。「くわしく見る」は`detailAvailable`の行だけで、`GET /api/history/diff-detail?workbookId={sourceId}&fromSnapshotId={fromSnapshotId}&toSnapshotId={toSnapshotId}&fromVersionId={fromVersionId}&toVersionId={toVersionId}`（原稿単位の比較。提出用PDFに入った版どうし）を開きます。組み合わせを切り替えた後に届いた前の応答は表示に使いません。

計算結果は`exports/change-reports/<packId>/<targetId>/<fromBuildId>__<toBuildId>/summary.json`に覚え、`complete`が true で、両方の`outputPdfSha256`が一致し、計算の版（`algorithmVersion`。`rowKey`・`toSourcePageNumber`・`fromSourcePageNumber`を足したので 2）が同じときだけ使い回します。見た目のハッシュがまだ無い結果は、解析が済むと変わるため使い回しません。

### 変更点PDF（`POST /api/v2/outputs/changes/report`）

変更点を、紙やメールで渡せる確認用のPDF（A4 横）にします。提出物ではありません。中身は表紙（一式名・出力先名、比べた2つの版の出力日時・ページ数・共有用フォルダーにコピーした日時と発行フォルダー名・出力ファイル名、まとめの数字、ページごとの詳細の始まるページ、「このPDFは確認用です。提出物ではありません。」）、変更一覧（今回 p.／左右・前回 p.／左右・原稿名 / シート名・変わったこと・見た目の変化）、ページごとの詳細（下記）です。全ページの上に「確認用」、下に「変更点 n / N」を入れます。総ページ数（表紙＋一覧＋詳細）はサーバーが HTML を作る時点で決めます。

**作成の版（`reportVersion`）を 2 に上げました。ページごとの詳細の無い、版 1 で作った変更点PDFは作成済みとみなさず（`report.reason: "outdated"`）、共有用フォルダーへも同梱しません。「変更点PDFを作り直す」（共有コピーの流れでは「変更点PDFを作ってからコピーする」）で作り直す必要があります。**

- 変わっていないページ（`unchanged`・`same-look`だけのページ）は載せません。前のページの増減でずれただけのページ（`shiftOnly`）は、ずれ幅が同じで今回のページ番号が続いている間を「p.15〜p.34 は前のページが増えたため2ページ後ろにずれました」の1行にまとめます。ずれた結果左右が入れ替わったページはまとめません。
- 頁・左右は`pages[]`の値をそのまま書き、左右の記録が無い版の行には頁番号だけを書きます（推定しません）。左右の記録が無い版との比較では、左右の入れ替わりの数を書きません。変更が1つも無いときは「変わったページはありません」と書いた表紙1枚だけです。
- 原稿名・シート名・ページ名・一式名はすべて HTML エスケープします。
- 一覧のセルは行数で切りません。行の高さを中身の長さ（全角1・半角0.6として数え、欄の幅で割った行数）から見積もり、1ページの高さ（12行 × 最小の行の高さ）に収まるようにページを分けます（1ページは12行まで）。1つの欄だけで1ページを超えるほど長い文字は「…（長いため省略しました）」と分かる形で短くします。
- 一覧の「見た目の変化」は、詳細ページがあれば「変更 3か所・追加 1か所（詳細 p.5）」のように枠の数と詳細ページの番号を書きます（表示しきれない枠は「ほか n か所」）。増えた・なくなったページは「今回だけにあるページ（詳細 p.6）」など。詳細が無い内容の変わったページは「変更箇所を特定できませんでした（ページ全体を確認してください）」、上限を超えたページは「詳細は省略しました（…最初の200ページ分だけ…）」です。アプリでの確かめ方（「くわしく見る」）は書きません（受け取った人はアプリを持っていないため）。

ページごとの詳細: 一覧の並び順で、`kinds`に`content`・`maybe-content`を含む行と、増えた（`added`）・なくなった（`removed`）行を、先頭から最大 200 件、1件1ページで載せます（ずれただけ・左右の入れ替わりだけ・名前だけの行には作りません）。201件目からは載せず、表紙に「ページごとの詳細は最初の200ページ分だけ載せています（ほかの n ページは一覧だけです）」と書きます。

- 見出しは「今回 p.12（右） ← 前回 p.10（右）／原稿名 / シート名（シート内 2/3 ページ）」。頁・左右は`pages[]`の値のまま書きます（推定しません）。枠の色の凡例（変更・追加・削除・不明）は見出しと同じ行の右端に置きます。
- 画像の枠（左右それぞれ幅 134mm）の高さはページごとに決め、画像の下の文字の一覧が全部入る高さまで広げます（最大 158mm、最小 110mm。最小まで縮めても入らない分は「（続き）」のページへ）。画像は縦横比を保って枠に収め（縦長のページは高さで、横長のページは幅で決まります）、紙の上の寸法（mm）を`img`の`style`に書きます。A4 縦の原稿は実物の約 53%、A4 横は約 45% です。寸法は`$Script:PackChangeReport*Mm`の変数1か所で決め、CSS とページ割りの見積もり（`Get-PackChangeReportDetailLayout`）の両方がそこから作ります（片方だけ変わって文字の一覧が紙の外で切れることが無いように）。
- 左に前回、右に今回のページの画像を置き、変わった箇所に番号つきの枠を重ねます（CSS の絶対配置。色はアプリの凡例と同じで、変更＝琥珀・追加＝青・削除＝赤の破線・不明＝灰の点線。色だけに頼らず「①変更」「②追加」のように番号と種類を書きます）。アプリの差分画面と同じく、追加の枠は今回の画像だけ、削除の枠は前回の画像だけ、変更・不明は両方に描きます。増えた・なくなったページは片側の画像とページ全体の枠だけです。
- 画像の下に番号つきの一覧「① 変更: 1,234 → 1,240」「② 追加: ○○」「③ 削除: ○○」を書きます（文字が取れた枠だけ。取れなかった枠は番号をまとめて1行）。入りきらないときは「（続き）」のページに回します。
- 枠は画面が送った`details`（アプリの「くわしく見る」と同じ計算）を使います。`details`の無い・使えない行は、ジョブがサーバーの画素比較（`Invoke-DiffImageBatchGeneration`の枠）で埋め、「枠は、前回と今回のページの画像の違いから自動で検出した範囲です。」と書きます。どちらも無いときは「変更箇所を特定できませんでした。前回と今回のページ全体を見比べて確認してください。」と書きます。
- 枠の座標は、前後のページのうち大きい方の寸法（左上を合わせる）に対する割合で受け取り、前後それぞれの画像の寸法に対する割合に換算してから描きます（用紙サイズが前後で違っても位置がずれないため）。画像の外に出る枠は削り、画像に入らない枠は描きません。
- 画像は作成ジョブ（子プロセス）が、アーカイブに記録された版（`metadata.sourceWorkbooks[].versionId`）の原稿の変換PDFから（`Resolve-DiffContentPdfPath`。別の版に代えません）、差分画面のサーバー側の画像化（`tools/diff-image-batch.ps1`・`PdfBatchRasterizer`）で 120dpi 相当（`$Script:PackChangeReportImageDpi`。差分画面の`DiffDetailDpi`と同じ。高さ 158mm に A4 縦を置いて約 225dpi で、PDF を 200% に拡大しても読めます）の PNG にし、TEMP の作業フォルダー（`rb-chg-<guid>\img\`）に置きます。画像を作るあいだは使う版に lease（スナップショットと変換PDF）を取り、作り終えたら外します。画像を作れなかったページは詳細を載せず、表紙に「n ページは画像を用意できなかったため、ページごとの詳細を載せていません」と書きます（作成は失敗にしません。ジョブの`warnings[]`に理由を残します）。保存先の`report.html`は画像を含まない控えです（PDF には画像が入っています）。
- 変更前後の文字・原稿名・シート名はすべて HTML エスケープします。

作り方: サーバーが`report.html`（UTF-8、`@page { size: A4 landscape }`）を作り、Microsoft Edge のヘッドレス印刷（`msedge.exe --headless --disable-gpu --no-first-run --user-data-dir=<一時フォルダー> --no-pdf-header-footer --print-to-pdf=<pdf> <file:///…/report.html>`）で PDF にします。起動中の Edge に処理を取られないよう、毎回 TEMP に作る専用の`--user-data-dir`を渡し、終わったら消します。印刷は TEMP の短い名前の作業フォルダー（`rb-chg-<guid>`）で行い、`file://`のパスに日本語・空白・長いパスが入らないようにします。待ち時間の上限は 60 秒＋一覧の行1件あたり 0.5 秒＋詳細ページの画像1枚あたり 2.5 秒（最大 15 分。`$Script:EdgePrintMaxTimeoutSeconds`）で、時間切れ（`timedOut`）は「時間内に終わりませんでした」として返します。Edge が正常終了しても、PDF が存在し先頭が`%PDF-`でなければ失敗にします（管理ポリシーでヘッドレスが止められている PC など）。

| メソッド・パス | 内容 |
| --- | --- |
| `POST /api/v2/outputs/changes/report` | 本文`{ packId, targetId, fromBuildId, toBuildId, details? }`（`fromBuildId`・`toBuildId`は両方必須。画面に出ている`changes.from.buildId`・`changes.to.buildId`）。`details`は画面が集めた「くわしく見る」の結果（下記）で、比べ直した結果と突き合わせて整えてからジョブの入力に入れます（このときは画像を作りません）。比べられない組み合わせ・組み込み一式・片方の版の省略は 400。作成済みなら作り直さず`{ jobType: "change-report", jobId: "", status: "completed", reportReady: true, report }`を返します。同じ組み合わせの作成中のジョブがあればそれに合流し（`joinedExistingReportJob: true`）、無ければ子プロセス（`server.ps1 -ChangeReportJobPath`）を起動してジョブ（`jobType: "change-report"`）を返します。応答は`{ ok, apiVersion: 2, job }` |
| `GET /api/v2/outputs/changes/report/status?jobId={jobId}` | ジョブの状態（`Read-RenderJobStatus`）。`status`（`queued`・`launching`・`running`・`completed`・`failed`）・`percent`・`message`。完了時は`report`（上の`report`と同じ形）、失敗時は`code`（`edge-missing`＝Edge が見つからない、`timeout`＝時間切れ、`no-pdf`＝PDF ができなかった）と`errors[]`（`error`・`code`・`detail`・`edgeLog`＝Edge の出力。診断用） |
| `POST /api/v2/outputs/changes/report/file` | 本文`{ packId, targetId, fromBuildId, toBuildId, action }`。`action: "open"`（既定）は PDF そのもの（`application/pdf`）を返し、`action: "reveal"`は保存先フォルダーをエクスプローラーで開いて PDF を選択状態にします（応答`{ ok, apiVersion: 2, path }`）。`meta.json`の`reportSha256`と`report.pdf`の中身が合わなければ開きません |

保存先は`exports/change-reports/<packId>/<targetId>/<fromBuildId>__<toBuildId>/`（`report.pdf`・`report.html`・`meta.json`・`summary.json`）です。ローカルの出力フォルダーにはコピーしません（提出用PDFと取り違えないため）。`report.pdf`は一時名で書いてから名前を変え、`meta.json`（`reportVersion`（今は 2）・`fromBuildId`・`toBuildId`・`fromOutputPdfSha256`・`toOutputPdfSha256`・`reportSha256`・`complete`・`changedCount`・`detailPageCount`（詳細ページの件数）・`detailSource`（`browser`・`server`・`mixed`・`none`）・`fileName`・`createdAt`）を最後に書きます。作成済み（`report.exists`）とみなすのは、`meta.json`の両方の`outputPdfSha256`が比べた2つの版と同じで、`reportSha256`が`report.pdf`の中身と合い、作成の版（`reportVersion`）が同じときだけです。見た目のハッシュがそろわないうち（`complete: false`）に作ったものは、そろった後は`outdated`です。`fileName`は`<提出用PDFのファイル名（拡張子なし）>_変更点.pdf`で、共有用フォルダーへ同梱するときの名前です。

画面での使い方: 変更点の画面の「変更点をPDFにする」で作成を始め、1秒ごとに状態を読んで進み具合を出します。終わると「開く」（別タブ）と「フォルダーを開く」が押せます。作成中に組み合わせを切り替えたり画面を閉じたりしたら、そのジョブの状態では画面を書き換えません（ジョブはサーバーで最後まで進み、同じ組み合わせでもう一度押すと合流します）。

`details`（「くわしく見る」の結果）: 作成を始める前に、画面は`pages[]`のうち`kinds`に`content`か`maybe-content`を含み、`detailAvailable`で版（`fromVersionId`・`toVersionId`）と`toSourcePageNumber`のある行を、先頭から最大 200 件選びます。原稿ごとに`GET /api/history/diff-detail`を行の版を指定して1回だけ読み、シート（今回の`sheetName`。無ければ前回の名前）とページ（`toSourcePageNumber`＝`pages[].afterPageNumber`。出力ページ番号では引かない）を引いて、差分画面と同じ計算（PDF.js の描画 → `diff-worker.js`の画素比較 → PDF.js の文字照合）を回します。進み具合は変更点PDFの欄（共有コピーの流れでは通知）に「変わったページを比べています（n / N ページ）」と出します。共有コピーの前に変更点PDFを作るときも同じです。1件の形は次のとおりです。

```json
{
  "rowKey": "item:it_1#2", "fromVersionId": "v_...", "toVersionId": "v_...", "sheetName": "売上",
  "beforePageNumber": 2, "afterPageNumber": 2, "width": 1240, "height": 1754,
  "status": "ready", "changedRatio": 0.012, "regionCount": 3,
  "regions": [
    { "kind": "modified", "before": { "x": 0.1, "y": 0.2, "width": 0.3, "height": 0.02 }, "after": { "x": 0.1, "y": 0.2, "width": 0.3, "height": 0.02 },
      "beforeText": "1,234", "afterText": "1,240" }
  ]
}
```

- 座標は 0〜1 の割合で、前後のページのうち大きい方の寸法（左上を合わせる）に対する値です。`kind`は`modified`・`added`・`removed`・`unknown`だけです。
- `beforeText`・`afterText`は、PDF の文字から取れた変更前後の文字（それぞれ 200 字まで）です。取れないときは空です。原稿の中身そのものなので、使う側で必ずエスケープします。
- 1件の`regions`は 30 件まで（`regionCount`は切る前の数）、`details`全体は UTF-8 で 768KB までに収め、入りきらない行は送りません（要求本文の上限 1MB を超えないため）。
- 比較情報が取れない・行の版と違う版の比較が返った・計算に失敗した行は送りません。集めている間に組み合わせを切り替えたり閉じたりしたら、何も送らず作成も始めません。集めること自体に失敗しても、`details: []`で作成は続けます。
- 計算は差分画面とは別の Worker と PDF のキャッシュで行うので、集めている間に差分画面を開いたり閉じたりしても止まりません。

サーバーでの検証（`ConvertTo-PackChangeReportDetails`。中身は信用せず、表示用の座標と文字だけを使う）:

- `rowKey`が比べ直した結果の`pages[]`に実在し、その行の`kinds`に`content`か`maybe-content`があること（大文字・小文字も区別）。同じ`rowKey`の2件目は捨てます。
- `fromVersionId`・`toVersionId`が行の版と同じで、`afterPageNumber`が行の`toSourcePageNumber`、`beforePageNumber`が行の`fromSourcePageNumber`と同じこと（違う版・違うページの比較は使いません）。
- `kind`は`modified`・`added`・`removed`・`unknown`だけ（それ以外の枠は捨てます）。座標は数値だけを受け、0〜1 に丸めてページからはみ出す分を削ります（文字列の座標・面積の無い枠は捨て、前後とも使えない枠は捨てます）。枠は1行 30 件まで、`beforeText`・`afterText`は制御文字を除いて 200 字まで、件数は 200 件まで、全体の見積もりの大きさは 2MB までです。合わない`details`は捨て、その行はサーバーの予備の枠で埋めます。
- 作成ジョブも、入力ファイルの`details`をもう一度同じように突き合わせてから使います。

画面での表示: 作成済み（または作り終えた）変更点PDFの`detailSource`が`server`なら「ページごとの詳細の枠は、画像の違いから自動で検出した範囲です。」、`mixed`なら「一部のページの詳細の枠は、画像の違いから自動で検出した範囲です。」を変更点PDFの欄に添えます。

### 変更点PDFの共有コピーへの同梱（`includeChangeReports`）

`POST /api/v2/outputs/publish`・`publish-all`の本文に`includeChangeReports: true`を付けると、出力先ごとに変更点PDFを提出用PDFと同じ発行フォルダーへ入れます（名前は`<提出用PDFのファイル名（拡張子なし）>_変更点.pdf`）。発行の経路は1本のままで、作業用の隠しフォルダーでコピーとハッシュの照合をしてから名前を一度だけ切り替え、発行フォルダーと中のファイル（変更点PDFも）の隠し属性を外します。出力と同じロックを持っている間に変更点PDFを作ることはしません（作成済みのものだけを付けます）。

同梱するのは、次をすべて満たすときだけです（`Resolve-PublishableChangeReport`）。

1. 変更点PDFの比べる先（`toBuildId`）が、その出力先の今の`buildId`と同じ
2. `meta.json`の`toOutputPdfSha256`が、今の提出用PDFの`outputIntegrity.sha256`と同じ
3. 比べる元（`fromBuildId`）が「前回コピーした版」（`Get-LastPublishedBuildId`）と同じ。好きな2回で作ったものは付けません（受け取った人が「前回の提出からの変更」と読むため）
4. `report.pdf`の中身が`meta.json`の`reportSha256`と同じで、作成の版（`reportVersion`）も同じ。場所は`exports/change-reports`の下だけ

- 提出用PDFの確認（最新か・ハッシュが合うか・ファイル名の重複）はすべて先に行い、1冊でも通らなければ変更点PDFも含めて何もコピーしません。変更点PDFの名前が別の提出用PDFと重なるときも何もコピーしません。
- 条件に合わない出力先は変更点PDFを付けずに提出用PDFだけをコピーし、応答の`changeReportsSkipped[]`（`targetId`・`targetName`・`reason`・`message`）に理由を返します。`reason`は`not-published`（まだコピーしたことがない）・`same-as-published`（前回コピーしてから出力し直していない）・`no-build`（保存済み出力の記録が無い）・`not-created`（まだ作っていない）・`outdated`（作り直すとコピーできる）・`outside`（場所が不正）。
- 応答の`changeReports[]`に同梱したもの（`targetId`・`targetName`・`fileName`・`sharedPath`・`fromBuildId`・`toBuildId`・`sha256`）、`files[].changeReport`にも同じ内容を返します。履歴イベントは`final.published`に`changeReport`を足し、同梱した変更点PDFごとに`final.change-report.published`を書きます（`Write-HistoryEvent`は今は何も書きません。発行記録の`changeReport`に残ります）。
- 見た目のハッシュがそろわないうちに作ったもの（`complete: false`）も、4つの条件を満たせば付けます（同じ2つの版の比較です）。画面では、そろった後は作り直してからコピーします。

画面での使い方: 「共有用フォルダーにコピー」・「まとめて共有用フォルダーにコピー」を押すと、出力先ごとに`GET /api/v2/outputs/changes`（既定の組み合わせ）を読み、`report.attachable`の出力先があれば確認画面に「変更点PDFもいっしょにコピーする」（最初からチェックあり）を出します。まだ作っていない（`publishable`が false）ものがあれば、ボタンが「変更点PDFを作ってからコピーする」になり、先に変更点PDFを作ってからコピーします。作れなかったときは何もコピーせず、チェックを外せば提出用PDFだけをコピーできると伝えます。付けられるものが無い出力先だけなら、今までどおり確認なしでコピーします。組み込み一式は対象外です。

## 自動PDF作成

- `GET /api/auto/state`: 原稿ごとの待機・実行・再試行・失敗状態と現在の設定を返す
- `PATCH /api/auto/settings`: 自動処理の設定を検証して保存し、スケジューラーを開始または停止する
- `POST /api/auto/render`: 指定原稿を待機時間なしで実行する

```json
{
  "enabled": true,
  "quietPeriodSeconds": 180,
  "maxRetryCount": 3,
  "minFreeMegabytes": 1024,
  "notifyOnCompletion": true,
  "notifyOnFailure": true
}
```

失敗時は指数バックオフで再試行し、上限へ達した同じ原稿版は停止します。原稿が再保存されて新しいハッシュになった場合は、新規更新として静止待ちから自動復帰します。管理データまたは出力先の空き容量が`minFreeMegabytes`未満の場合は`disk-low`として保留し、成果物を変更しません。

## ページ範囲

`POST /api/pages/update`の`pageRange`に`"2"`または`"2-5"`を指定すると、項目が参照する変換PDFの一部だけを最終PDFへ含めます。全ページへ戻す場合は`clearPageRange: true`を送信します。

最終PDFのmanifestはschema 3です。`pages`には論理項目と元PDF内の開始・終了ページ、`physicalPages`には出力後の物理ページ番号を記録します。表紙・目次・区切りページも通常の原稿として登録し、ページ構成で配置します。成功時のmanifestは`exports/manifest_<volume>_<category>.json`に保存されます。
