# ReportBinder Local API V4

すべて`127.0.0.1`だけで待ち受けます。`X-ReportBinder-Token`ヘッダーまたは`?token=`が必要です。

validationエラーはHTTP 400、未処理エラーはHTTP 500です。

管理データの保存形式は汎用資料パック対応の`schemaVersion: 3`です。Local API V4は既存UIとの互換性のため、`structure`を従来の`schemaVersion: 2`、`workbooks / pages / volumes`形式へ変換して返します。保存データ上では`pack / source / unit / item / artifact / output`と互換フィールドが同期されます。

V4の`/api/submission-files`、`/api/workbooks/register*`、`/api/workbooks/render/start`は、内部では共通原稿dispatcherを経由し、Excelは`excel-com-v1`、Wordは`word-com-v1`、PDFは`pdfbox-import-v1`アダプターへ接続します。レスポンスには共通項目として`sourceId / sourceType / adapterId / packId`が追加される場合があります。従来の`workbookId / category`も維持します。

## V2 汎用資料パックAPI

新UI向けの汎用ドメインAPIは`/api/v2`に置きます。APIの`v2`はHTTP契約、レスポンスの`domainSchemaVersion: 3`は保存ドメインの版を表します。

- `GET /api/v2/state`: `packs / sources / units / items / artifacts / outputs`を返す
- `GET /api/v2/pack-templates`: ECM / BOD / DMMの組み込みテンプレートを返す
- `GET /api/v2/packs`: 現在の資料パック一覧を返す
- `GET /api/v2/source-candidates`: 登録可能なExcel・Word・PDF原稿を返す
- `POST /api/v2/sources/register-batch`: 複数形式の原稿を一括登録する
- `POST /api/v2/sources/unregister`: `sourceId`で登録解除する
- `POST /api/v2/sources/scan-updates`: 登録済み原稿の更新を検出する
- `POST /api/v2/sources/render/start`: `sourceIds`で変換PDF作成を開始する
- `PATCH /api/v2/sources/{sourceId}`: `ownerDepartment / required / defaultTargetId`を更新する

`PATCH`の例：

```json
{
  "ownerDepartment": "経理部",
  "required": true,
  "defaultTargetId": "appendix"
}
```

既存のECM/BOD/DMM操作は引き続きV4 APIと互換です。組み込みテンプレートの`acceptedSourceTypes`は`excel / word / pdf`です。

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
    { "templateId": "builtin-ecm", "packId": "pack_ecm", "acceptedSourceTypes": ["excel", "word", "pdf"] }
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

提出フォルダ直下の原稿一覧を返します。現在のUIが登録対象として表示するのは`.xlsx`、`.docx`、`.pdf`です。Wordの一時ファイル`~$*.docx`、`.docm`、子フォルダ内のファイルは候補に含めません。

- `scannedAt`: 一覧を取得した日時
- `files[].sourceType`: `excel`、`word`または`pdf`
- `files[].adapterId`: 使用する原稿アダプター
- `files[].modifiedAt`: ファイルサーバーから再取得した原稿の最終保存日時
- `files[].modifiedAtUtcTicks`: 更新判定・調査用のUTC ticks

## POST /api/workbooks/register-batch

categoryは必須です。

```json
{
  "category": "ecm",
  "relativePaths": ["FY160-4Q_ECM_J_00_Cover.xlsx", "department-report.docx", "department-appendix.pdf"]
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

登録済みExcel・Word・PDF原稿の更新を確認します。Word/PDFの差し替えは`source-updated`となり、再変換後に物理ページの追加・削除も同期します。

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
  "includeCover": true,
  "includeToc": true,
  "includeSectionDividers": true,
  "outputFileNamePattern": "{projectId}_{targetName}_{yyyyMMdd}.pdf"
}
```

ファイル名には`{projectId}`、`{packName}`、`{targetName}`、`{yyyyMMdd}`を使用できます。

## ページ範囲

`POST /api/pages/update`の`pageRange`に`"2"`または`"2-5"`を指定すると、項目が参照する変換PDFの一部だけを最終PDFへ含めます。全ページへ戻す場合は`clearPageRange: true`を送信します。

最終PDFのmanifestはschema 3です。`pages`には論理項目と元PDF内の開始・終了ページ、`physicalPages`には表紙・目次・区切りを含む出力後の物理ページ番号を記録します。成功時のmanifestは`exports/manifest_<volume>_<category>.json`に保存されます。
