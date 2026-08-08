# ReportBinder Local API V4

すべて`127.0.0.1`だけで待ち受けます。`X-ReportBinder-Token`ヘッダーまたは`?token=`が必要です。

validationエラーはHTTP 400、未処理エラーはHTTP 500です。

管理データの保存形式は汎用資料パック対応の`schemaVersion: 3`です。Local API V4は既存UIとの互換性のため、`structure`を従来の`schemaVersion: 2`、`workbooks / pages / volumes`形式へ変換して返します。保存データ上では`pack / source / unit / item / artifact / output`と互換フィールドが同期されます。

V4の`/api/submission-files`、`/api/workbooks/register*`、`/api/workbooks/render/start`は、内部では共通原稿dispatcherを経由し、Excelは`excel-com-v1`、Wordは`word-com-v1`、PowerPointは`powerpoint-com-v1`、PDFは`pdfbox-import-v1`アダプターへ接続します。レスポンスには共通項目として`sourceId / sourceType / adapterId / packId`が追加される場合があります。従来の`workbookId / category`も維持します。

## V2 汎用資料パックAPI

新UI向けの汎用ドメインAPIは`/api/v2`に置きます。APIの`v2`はHTTP契約、レスポンスの`domainSchemaVersion: 3`は保存ドメインの版を表します。

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
- `POST /api/v2/sources/render/start`: `sourceIds`で変換PDF作成を開始する
- `PATCH /api/v2/sources/{sourceId}`: `ownerDepartment / required / defaultTargetId`を更新する
- `POST /api/v2/items/reorder`: 資料パック内のページを本体・補足・未振り分けへ移動、並べ替えする
- `PATCH /api/v2/items/{itemId}`: ページ名、番号表示、ページ範囲、出力先を変更する
- `GET /api/v2/layout/snapshots?packId={packId}`: 資料パックの保存済みページ構成を返す
- `POST /api/v2/layout/restore/preview`: 保存済み構成を復元した場合の適用数・差異を返す
- `POST /api/v2/layout/restore`: 保存済みのページ構成を同じ資料パックへ復元する
- `GET /api/v2/outputs/readiness?packId={packId}`: 本体・補足の出力可否と再出力理由を返す
- `POST /api/v2/outputs/build`: 資料パックの提出用PDFを出力する
- `POST /api/v2/outputs/publish`: 最新の提出用PDFを共有発行する
- `GET /api/v2/outputs/archives?packId={packId}`: 自動保存された最終出力を返す

`PATCH`の例：

```json
{
  "ownerDepartment": "経理部",
  "required": true,
  "defaultTargetId": "appendix"
}
```

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
    "outputFileNamePattern": "{packName}_{yyyyMMdd}.pdf"
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

`POST /api/history/diff-review`は、表示中の比較に含まれる1つのページ項目を確認済み・未確認へ切り替える。`workbookId`、任意比較の場合は`fromSnapshotId`と`toSnapshotId`、画面が取得した`baselineVersionId`と`currentVersionId`、`sheetKey`、`confirmed`を送る。確認状態は利用者のローカル設定へ保存され、snapshot、render version、差分アルゴリズム版のいずれかが変わった比較へは引き継がれない。

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
  "workbookIds": ["..."]
}
```

PDF必要分：

```json
{
  "category": "ecm",
  "onlyUpdated": true
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

実際に順序・割当が変わったvolumeだけを更新対象にします。

## POST /api/pages/sort-by-sheet

categoryは必須です。

```json
{
  "category": "ecm",
  "volumes": ["ja-main", "ja-appendix", "none"]
}
```

volumeを省略した場合は本体・補足だけを対象にします。割当は変えず、各表の中だけを半角数字、ファイル順、ファイル名、pageIdの安定順で並べ替えます。

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
```

`status`は`not-generated`、`generating`、`ready`、`unavailable`、`failed`です。
領域座標は0～1の正規化値で、各領域は`modified`、`added`、`removed`のいずれかです。
判定不能のページでは推測した領域を返しません。

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

ファイル名には`{projectId}`、`{packName}`、`{targetName}`、`{yyyyMMdd}`を使用できます。

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
POST  /api/v2/outputs/build
POST  /api/v2/outputs/file
POST  /api/v2/outputs/publish
GET   /api/v2/outputs/archives?packId=pack_...
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
