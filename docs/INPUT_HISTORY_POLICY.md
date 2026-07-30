# 入力履歴の運用方針（Phase 0 承認事項）

V5の履歴・差分機能は、ワークスペース単位の承認フラグ `_reportbinder\common\policy.json`
が無い間は一切動作せず、V4.1と同じ挙動になります。本書は承認内容と、承認後の運用を記録します。

## 承認事項（2026-07-29）

1. **現物の保存**：他部署提出Excelの現物を `_reportbinder` 配下に保存してよい。
2. **閲覧権限**：`_reportbinder` は提出フォルダ直下にあり、アクセスできる範囲は現行と同一。
   ただし従来は各部署の最新ファイルのみが見えていたのに対し、今後は**過去の提出版が保持される**点が変わる。
3. **保存期間**：期間では区切らない。世代数で管理し、不足分は正式版完成後に手動で整理する。
4. **正式版に使った版の固定**：出力時点でアプリが自動的に固定（pin）し、
   `exports\archive\<カテゴリ>\<巻>\<buildId>` に最終PDF・manifest・SHA-256とともに記録する。
   提出フォルダ側のファイルがその後編集された場合は最終PDFが「再出力が必要」と表示されるため、
   運用ルールで編集を禁止する必要はない。

## 承認フラグ

`_reportbinder\common\policy.json` を次の内容で作成します（`docs\POLICY_SAMPLE.json` が雛形）。

```json
{
  "schemaVersion": 1,
  "inputHistoryApproved": true,
  "sourceRetentionApproved": true,
  "approvedBy": "",
  "approvedAt": "2026-07-29",
  "minimumAppVersion": ""
}
```

| キー | 意味 |
|---|---|
| `inputHistoryApproved` | 検知版の記録と差分機能。`false` の間はV4.1と同じ挙動 |
| `sourceRetentionApproved` | 提出Excel現物の保存。`false` なら一時コピーのみで残さない |

`sourceRetentionApproved` は、`dataDir` が提出フォルダ配下にあることも起動時に確認します。
現行の `提出フォルダ\_reportbinder` 構成はこの条件を満たします。

**policy.json は技術的なアクセス制御ではなく運用フラグです。**
共有 `dataDir` に書き込める利用者は編集できます。

## 保持の仕組み（世代数と pin）

保持の対象は2種類あり、それぞれ別の設定で決まります。

| 対象 | 設定 | 既定 |
|---|---|---|
| 正式版に**使っていない**通常の検知版 | `inputHistory.retainSourceVersions` | 5 |
| 正式版に**使った**検知版（pin付き） | `inputHistory.sourceRetentionDaysAfterBuild` | `null`（＝期間では消さない） |
| シート別PDF（content-pdf）の世代 | `inputHistory.retainContentPdfVersions` | 3 |

承認事項3のとおり `sourceRetentionDaysAfterBuild` は `null` のままにします。
これにより**正式版の証跡は自動では消えず**、それ以外は最新5版だけが残ります。

同じ内容のExcelは重複排除されるため、正式版を何回出しても提出Excelが変わっていなければ
同じ検知版が使い回されます。増えるのは実際に正式版へ使われた版の数だけです。

`retainSourceVersions` は現物だけでなく検知版フォルダごとの保持数です。
枠から外れて pin も無い版は、画像ハッシュや検知記録も含めて丸ごと削除されます。
差分を何回分さかのぼりたいかで決めてください。

## 設定値の配り方（重要）

これらの値は共有側ではなく、利用者ごとの
`%LOCALAPPDATA%\ReportBinder\config.json` を参照します。
掃除を実行した人の設定が共有ワークスペースへ適用されるため、値がPCごとに違うとばらつきます。

- **まだ起動していないPC**：共有アプリの `app\default-config.json` の値が初回起動時に複製されます。
- **既に起動したことがあるPC**：既存キーは上書きされません（不足キーだけが補完されます）。
  値を変えるには `%LOCALAPPDATA%\ReportBinder\config.json` を直接編集してください。

## 手動整理の手順

1. ReportBinder を終了する（`app\tools\stop-reportbinder.cmd`）。
2. `_reportbinder\<言語>\input-history\<workbookId>\<snapshotId>` を**フォルダごと**、古い順に削除する。
3. フォルダ内の `pins` や `manifest.json` だけを消さない（保護情報と完成マーカーが壊れます）。

`pins` に `final-pdf_*` がある検知版は正式版の証跡です。削除は正式版が不要になってから行ってください。

容量は最終PDF画面に表示されます。`softCapMegabytes`（既定5120）の `warnAtPercent`（既定80%）で警告します。

## 段階的な有効化

1. **第1段階**：`inputHistoryApproved` / `sourceRetentionApproved` を `true`。自動処理は無効のまま。
2. **第2段階**：`autoRender.enabled` を `true`。最初は1台のPCだけで試す。
   複数台で同時に有効化すると、ブック単位の所有権ロック（`locks\auto-owner_*.lock`）の
   取り合いが起きて切り分けが難しくなります。
