# ReportBinder Local API V4

すべて`127.0.0.1`だけで待ち受けます。`X-ReportBinder-Token`ヘッダーまたは`?token=`が必要です。

validationエラーはHTTP 400、未処理エラーはHTTP 500です。

## GET /api/state

設定、workbook、page、volume状態、サマリー、PDF.js配置状態を返します。`pdfjsPresent` / `pdfjsMode` はファイル配置の診断値であり、V5の画面プレビューがPDF.jsを使用していることを示すものではありません。

主な追加項目：

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

提出フォルダ直下のExcel一覧を返します。現在のUIが登録対象として表示するのは`.xlsx`です。

- `scannedAt`: 一覧を取得した日時
- `files[].modifiedAt`: ファイルサーバーから再取得したExcelの最終保存日時
- `files[].modifiedAtUtcTicks`: 更新判定・調査用のUTC ticks

## POST /api/workbooks/register-batch

categoryは必須です。

```json
{
  "category": "ecm",
  "relativePaths": ["FY160-4Q_ECM_J_00_Cover.xlsx"]
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

登録済みExcelの更新を確認します。

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

最新PDFに紐づく比較メタデータ、シート一覧、ページ対応、差分領域、生成状態を返します。
`fromSnapshotId`と`toSnapshotId`を両方指定した場合は、同じ登録済みExcelの任意の
保存済み2版を比較します。片方だけの指定、同一版、履歴外のIDは拒否します。

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
