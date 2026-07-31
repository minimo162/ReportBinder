# 入力履歴・差分・保存設定

入力履歴と差分機能は既定で有効です。旧版で必要だった
`_reportbinder\common\policy.json` は使用しません。ファイルが無い場合も機能制限はなく、
既存の `policy.json` が残っていても値は参照されません。

## 提出Excel現物の保存条件

提出Excelの現物は、管理データ（`dataDir`）が提出フォルダ配下にある場合だけ保存します。
標準構成の `提出フォルダ\_reportbinder` はこの条件を満たします。

`dataDir` が提出フォルダの外にある場合、履歴・差分の記録は続けますが、提出Excelの現物は
履歴フォルダへ複製しません。この制限は設定フラグではなく、実パスの検証で常に適用します。

## 保持の仕組み（世代数と pin）

保持対象は次の設定で管理します。

| 対象 | 設定 | 既定 |
|---|---|---|
| 正式版に使っていない通常の検知版 | `inputHistory.retainSourceVersions` | 5 |
| 正式版に使った検知版（pin付き） | `inputHistory.sourceRetentionDaysAfterBuild` | `null`（期間では削除しない） |
| シート別PDF（content-pdf）の世代 | `inputHistory.retainContentPdfVersions` | 3 |
| 履歴容量の警告基準 | `inputHistory.softCapMegabytes` | 5120 MB |
| 容量警告の割合 | `inputHistory.warnAtPercent` | 80% |

正式版を出力した検知版はアプリが自動的に pin し、
`exports\archive\<カテゴリ>\<巻>\<buildId>` に最終PDF・manifest・SHA-256とともに記録します。
同じ内容のExcelは重複排除されるため、内容が変わらなければ同じ検知版を再利用します。

`retainSourceVersions` の枠から外れ、pin も無い版は、画像ハッシュや検知記録を含む
検知版フォルダごと削除されます。差分を何版さかのぼる必要があるかに合わせて設定してください。

## 設定値の配り方

設定は利用者ごとの `%LOCALAPPDATA%\ReportBinder\config.json` を参照します。
掃除を実行したPCの設定が共有ワークスペースへ適用されるため、値は利用PC間で揃えてください。

- まだ起動していないPC：共有アプリの `app\default-config.json` が初回起動時に複製されます。
- 起動済みのPC：不足キーだけが補完され、既存値は上書きされません。

## 自動処理

`autoRender.enabled` の既定値は `false` です。自動処理を使う場合だけ `true` にします。
最初は1台のPCだけで確認してください。複数台で同時に有効化すると、
ブック単位の所有権ロック（`locks\auto-owner_*.lock`）の取り合いが起き、調査が難しくなります。

## 手動整理

1. ReportBinder を終了します（`app\tools\stop-reportbinder.cmd`）。
2. `_reportbinder\<言語>\input-history\<workbookId>\<snapshotId>` を古い順にフォルダごと削除します。
3. フォルダ内の `pins` や `manifest.json` だけを削除しないでください。

`pins` に `final-pdf_*` がある検知版は正式版の証跡です。正式版が不要になってから削除してください。
容量は最終PDF画面で確認できます。
