# ReportBinder 汎用資料パック化 設計書

- 状態: 汎用資料パックMVP完成（PR 1〜8実装・実機E2E・大容量受け入れ完了）
- 作成日: 2026-08-06
- 対象: Local API V5 / structure schemaVersion 3（V4 / schemaVersion 1・2互換移行を継続）
- 目的: ECM 資料作成で培った「原稿更新の検知・PDF化・差分確認・ページ構成・最終出力」を、各部署から集めた Excel / Word / PDF にも適用する

## 1. 結論

ReportBinder を「ECM 用 Excel 結合アプリ」から、**原稿追従型の資料パック作成アプリ**へ拡張する。

既存の ECM / BOD / DMM は廃止せず、最初から用意される「組版テンプレート」として残す。利用者は従来どおり ECM を選んで作業でき、新しい用途では「部門月次報告」「取締役会資料」「監査提出資料」などの資料パックを作成する。

最初の汎用版で扱う原稿は次の3形式とする。

| 形式 | MVP の扱い | 構成項目の単位 | 更新差分 |
|---|---|---|---|
| Excel `.xlsx` | 現行機能をアダプター化 | ワークシート | シート追加・削除・変更 + PDF見た目差分 |
| Word `.docx` | Word COM で一時コピーを PDF 化 | 物理ページ | ページ追加・削除・変更 + PDF見た目差分 |
| PDF `.pdf` | 再変換せず検証・保管 | 物理ページ | ページ追加・削除・変更 + PDF見た目差分 |

画像、メール、OCR、Office 文書の意味的な赤入れは初期範囲に含めない。PowerPointは汎用資料パックMVPの次段階として`.pptx`の安全変換と視覚差分へ対応する。まず、異種原稿を混在させても現在の更新検知・履歴・構成・最終出力が一貫して動くことを完成条件とする。

## 2. 市販製品から採用する考え方

市販製品の機能をそのまま広く模倣するのではなく、ReportBinder の強みにつながる部分だけを採用する。

| 参考製品 | 参考にする点 | ReportBinder での具体化 |
|---|---|---|
| Adobe Acrobat | 異種ファイルを追加して順序変更し、1つの PDF にする | Excel / Word / PDF を同じ構成ボードへ置く |
| Tungsten Power PDF | 作成前に入力順・ブックマーク・出力条件を指定する | 資料パックテンプレートと出力先を分離する |
| Bundledocs | セクション、索引、直前差し替えを含むバインダー作成 | セクション付き構成と、差し替え後も保持される手動順序 |
| Litera Compare | 形式ごとの比較を同じ入口から提供する | 履歴画面は共通化し、比較方法は原稿アダプター側で選ぶ |
| Bluebeam Revu | PDF の重ね合わせとページ対応 | 正規化した PDF 見た目差分を全形式の共通比較基盤にする |

ReportBinder が狙う差別化は、単発の PDF 結合ではない。**原稿フォルダーを継続監視し、どの原稿が変わったか、前回出力のどこが変わったか、構成が最新かを示した上で、同じ資料パックを繰り返し作れること**に置く。

## 3. 製品概念と用語

現行名称は次の概念へ移す。

| 現行 | 新しい内部概念 | 画面上の日本語 | 説明 |
|---|---|---|---|
| category | `pack` | 資料パック | 1つの提出物・会議資料・報告書のまとまり |
| workbook | `source` | 原稿 | 登録した Excel / Word / PDF |
| sheet | `unit` | 原稿内項目 | Excel のシート、Word/PDF の物理ページ |
| page record | `item` | 構成項目 | 順序、出力先、表示名、ページ番号設定を持つ単位 |
| volume | `target` | 出力先 | 本体、補足、非出力、または利用者定義の出力 PDF |
| category preset | `pack template` | テンプレート | ECM、BOD、DMM、汎用資料などの初期設定 |
| input snapshot | `source snapshot` | 原稿の履歴 | 原稿ファイルと検知時情報の不変な版 |
| content-pdf | `render artifact` | 変換PDF | 原稿を最終組版へ渡す中間 PDF |

重要な区別は `unit` と `item` である。

- `unit` は原稿アダプターが発見する論理単位。Excel はワークシート単位、Word/PDF はPDF化後の物理ページ単位。
- `item` は利用者が構成ボードで動かす単位。MVP では1 unitにつき1 itemとする。
- Word/PDFを物理ページ単位にしたことで、利用者は任意のページだけを本体・補足へ移動でき、差し替え時のページ追加・削除も既存のシート同期と同じ規則で扱える。

## 4. 利用者の基本フロー

1. テンプレートを選び、資料パックを作る。
2. 原稿フォルダーを選び、候補一覧から Excel / Word / PDF を登録する。
3. 登録時に担当部署、必須・任意、既定の出力先を指定する。
4. 「変換PDFを作成」を実行する。PDF 原稿は検証して取り込み、Excel / Word は一時コピーから PDF 化する。
5. 構成ボードで項目を複数選択し、出力先や位置を変更する。
6. 原稿更新を検知したら、変更された原稿だけ再変換する。
7. 自動比較または任意の履歴2版を開き、ページごとの見た目差分を確認する。
8. 出力準備状況を確認し、本体・補足などの最終 PDF を作成する。
9. 最終 PDF と manifest、使用した原稿版をアーカイブする。

既存 ECM 利用者には、テンプレート選択後の手順と表示順を現在と同じに見せる。汎用資料パックでだけ追加項目を表示する。

## 5. 情報アーキテクチャと画面設計

### 5.1 全体ナビゲーション

画面上部に現在の資料パックを常時表示し、次の5画面へ整理する。

1. **概要**: 未登録、更新あり、変換失敗、最終 PDF の再出力が必要、最近の履歴
2. **原稿**: 候補の登録、担当部署、必須・任意、更新状態、変換操作
3. **ページ構成**: 出力先ごとの並べ替え、複数選択、未使用項目、元に戻す・やり直す
4. **履歴・比較**: 原稿ごとのタイムライン、任意の2版比較、変更ページ一覧
5. **最終出力**: 必須原稿・変換・未振り分け・仕上げ設定の出力前チェック、出力先別の準備状況、ブロッカー、作成、公開先へのコピー

現行の作業順は維持するが、「Excel登録」「シートPDF」という形式依存のラベルは「原稿登録」「変換PDF」へ変更する。原稿行には形式アイコンと `Excel / Word / PDF` バッジを表示する。

### 5.2 原稿登録

候補一覧は次の列を持つ。

| 列 | 内容 |
|---|---|
| 選択 | Shift 範囲選択、Ctrl 個別選択、全選択 |
| 原稿名 | 相対パスとファイル名 |
| 種類 | Excel / Word / PDF |
| 担当部署 | 任意入力。テンプレートから候補を出せる |
| 必須 | 最終出力の readiness 判定へ使用 |
| 既定の出力先 | 本体、補足、未振り分けなど |
| 状態 | 未登録、最新、更新あり、変換中、変換失敗、暗号化など |

登録操作は複数ファイルをまとめて行い、成功分と失敗分を分けて通知する。1件の失敗で全件をロールバックしない。

### 5.3 ページ構成

構成項目カードには次を表示する。

- サムネイル
- 表示名
- 元原稿名と形式バッジ
- Excel はシート名、Word/PDF はページ数
- 更新あり、変換失敗、差分ありの状態
- ページ番号表示設定

操作は一般的な項目移動 UI に揃える。

- クリックで単一選択、Ctrl+クリックで追加、Shift+クリックで範囲選択
- ドラッグ中は選択中の全項目を1つの束として移動
- 挿入位置を項目間の線で明示
- 別の出力先へドラッグ、またはコンテキストバーの「移動先」で移動
- 移動後も選択を維持し、続けて移動できる
- `Ctrl+Z` / `Ctrl+Y` と画面上の元に戻す・やり直す
- キーボード操作用に「前へ」「次へ」「先頭」「末尾」「出力先へ移動」を用意

新しい unit が見つかった場合は、既存の手動順序を崩さない。テンプレートルールに従い、元原稿の直後か「未振り分け」の末尾へ挿入し、追加位置を通知する。

### 5.4 履歴・比較

左側に原稿の履歴、中央に比較対象、右側に変更箇所を表示する。

- 既定は「前回正式版で使用した版」と「現在版」の比較
- 利用者は履歴から任意の2版を選べる
- Excel はシート一覧、Word/PDF はページ一覧を表示
- 追加、削除、変更、判定不能を共通の表現にする
- 見た目差分は変更領域のマスクと前後表示を提供
- MVP では「文字列が何から何へ変わった」という意味的な断定はしない

## 6. schemaVersion 3

### 6.1 基本形

```jsonc
{
  "schemaVersion": 3,
  "language": "ja",
  "packs": [
    {
      "packId": "pack_ecm",
      "templateId": "builtin-ecm",
      "templateVersion": 1,
      "displayName": "ECM資料",
      "createdAt": "2026-08-06T09:00:00+09:00",
      "settings": {}
    }
  ],
  "sources": [
    {
      "sourceId": "wb_existing_id",
      "packId": "pack_ecm",
      "relativePath": "経理\\月次報告.xlsx",
      "sourceType": "excel",
      "adapterId": "excel-com-v1",
      "displayName": "月次報告",
      "ownerDepartment": "経理部",
      "required": true,
      "defaultTargetId": "main",
      "status": "ready",
      "currentSourceHash": "sha256:...",
      "currentSnapshotId": "20260806T090000.000_...",
      "lastRenderedSnapshotId": "20260806T090000.000_...",
      "lastRenderedVersionId": "render_...",
      "lastRenderedAt": "2026-08-06T09:01:00+09:00",
      "sourceMetadata": {
        "size": 123456,
        "modifiedAtUtcTicks": 638900000000000000
      }
    }
  ],
  "units": [
    {
      "unitId": "unit_...",
      "sourceId": "wb_existing_id",
      "unitKind": "worksheet",
      "sourceKey": "worksheet:1",
      "title": "1 売上実績",
      "sourceIndex": 0,
      "status": "ready",
      "renderVersionId": "render_...",
      "artifactId": "artifact_...",
      "physicalPageCount": 2
    }
  ],
  "items": [
    {
      "itemId": "existing_page_id",
      "packId": "pack_ecm",
      "sourceId": "wb_existing_id",
      "unitId": "unit_...",
      "targetId": "main",
      "sectionId": "body",
      "order": 10,
      "enabled": true,
      "title": "1 売上実績",
      "numberingMode": "visible",
      "numberingManual": false,
      "artifactId": "artifact_...",
      "pageRange": null
    }
  ],
  "artifacts": [
    {
      "artifactId": "artifact_...",
      "sourceId": "wb_existing_id",
      "unitId": "unit_...",
      "snapshotId": "20260806T090000.000_...",
      "adapterId": "excel-com-v1",
      "adapterVersion": 1,
      "relativePdfPath": "content-pdf\\...\\1.pdf",
      "sha256": "...",
      "pageCount": 2,
      "createdAt": "2026-08-06T09:01:00+09:00"
    }
  ],
  "outputs": {
    "pack_ecm|main": {
      "status": "built",
      "builtFingerprint": "...",
      "lastBuiltAt": "2026-08-06T09:10:00+09:00",
      "outputPdf": "出力\\ECM資料_本体.pdf",
      "buildId": "build_..."
    }
  }
}
```

### 6.2 ID と順序の規則

- `packId` は作成時の不変 ID。表示名変更では変えない。
- schema 2 の `workbookId` は移行時にそのまま `sourceId` として保持する。
- `unitId` は `sourceId + sourceKey` から生成し、表示名や並び順の変更では変えない。
- Excel の `sourceKey` は可能ならワークシート CodeName を使い、取得できない場合は既存 sheet key と名称対応履歴を使う。
- Word/PDF の MVP は1文書1 unitなので `sourceKey = document:root` とする。
- schema 2 の `pageId` は移行時に `itemId` として保持する。
- `order` は現在と同じ10刻みを基本とし、必要時だけ対象出力先内を再採番する。
- `artifact` は不変。再変換時は新規作成し、既存 item の参照を transaction 内で差し替える。

### 6.3 テンプレート定義

テンプレートはアプリ配下の読み取り専用 JSON と、dataDir 内の利用者定義 JSON を同じ形式で扱う。

```jsonc
{
  "templateSchemaVersion": 1,
  "templateId": "builtin-generic-department-pack",
  "version": 1,
  "displayName": "部門資料パック",
  "acceptedSourceTypes": ["excel", "word", "powerpoint", "pdf"],
  "targets": [
    {
      "targetId": "main",
      "displayName": { "ja": "本体", "en": "Main" },
      "required": true,
      "sections": [
        { "sectionId": "body", "displayName": { "ja": "本文", "en": "Body" } }
      ]
    },
    {
      "targetId": "appendix",
      "displayName": { "ja": "補足", "en": "Appendix" },
      "required": false,
      "sections": [
        { "sectionId": "appendix", "displayName": { "ja": "補足", "en": "Appendix" } }
      ]
    }
  ],
  "rules": {
    "newItemDestination": "unassigned",
    "retainManualOrder": true,
    "blockBuildWhenRequiredSourceIsStale": true,
    "blockBuildWhenRequiredSourceFailed": true
  },
  "output": {
    "fileNamePattern": "{packName}_{targetName}_{yyyyMMdd}.pdf",
    "pageNumbering": "continuous",
    "bookmarks": "from-items",
    "tableOfContents": false
  }
}
```

ECM / BOD / DMM は組み込みテンプレートとして、現在のファイル名、余白、ページ番号、出力先、カテゴリ判定を再現する。既存案件では `packId` と template の対応を自動生成するため、利用者による移行設定は不要とする。

## 7. 原稿アダプター設計

### 7.1 責務

形式固有の処理を `server.ps1` の主処理から分離する。PowerShell のクラス継承ではなく、共通の dispatcher と結果オブジェクトで統一する。

```powershell
Get-SourceAdapterDescriptor -SourceType <excel|word|pdf>
Test-SourceCandidate         -Context <context>
Inspect-Source               -Context <context>
Render-Source                -Context <context> -Snapshot <snapshot>
Get-SourceChangeSummary      -Context <context> -Before <snapshot> -After <snapshot>
```

`Render-Source` は structure を直接書き換えず、次の結果だけを返す。

```jsonc
{
  "sourceId": "...",
  "snapshotId": "...",
  "adapterId": "word-com-v1",
  "units": [],
  "artifacts": [],
  "warnings": [],
  "temporaryFiles": []
}
```

orchestrator が結果を検証し、短時間だけ `structure.lock` を取得して units / artifacts / items を一括更新する。Office COM や PDF 解析中に `structure.lock` を保持しない。ロック順は現在の `workbook/volume → structure` を、`source/output → structure` として引き継ぐ。

### 7.2 共通パイプライン

```text
候補検出
  → パス・拡張子・サイズ検証
  → 原稿 snapshot 作成または重複排除
  → 形式別 inspect
  → 形式別 render / PDF検証
  → artifact の SHA-256・ページ数検証
  → unit 対応付け
  → 既存の手動構成を保った item 同期
  → 前版との比較メタデータ作成
  → structure transaction
  → 保持ポリシー適用
```

途中失敗時は新 artifact を参照せず、前回成功した artifact を維持する。source の状態だけを `render-error` とし、エラー内容と失敗 stage を job 結果へ記録する。

### 7.3 Excel アダプター `excel-com-v1`

現行 `Render-Workbook` を移設して挙動を維持する。

- 対応: `.xlsx`
- unit: 印刷対象ワークシート
- 現行余白、拡大縮小、印刷範囲、シート別 PDF を維持
- 元ファイルを変更せず、一時コピーで処理
- 非表示シートは現行ルールを継承し、設定可能にするのは後続版
- シート追加は既存の手動順序を保持して同期
- シート名変更は安定キーと前後の位置・内容を用いて対応付けし、確信できない場合は削除+追加として扱う

### 7.4 Word アダプター `word-com-v1`

- 対応: `.docx`
- unit: MVP は文書全体1件
- 元ファイルを変更せず、一時コピーを Word で開く
- 既定の表示は「最終版」。変更履歴のマークアップを出す設定は pack override とする
- 外部リンクの自動更新はしない
- マクロは実行しない。`.docm` は MVP では候補に出さない
- フィールド・目次の更新は既定で行わない。必要な案件だけ `word.updateFieldsOnTemporaryCopy` を明示指定する
- パスワード保護、破損、変換ダイアログが必要な文書は `requires-attention` として失敗させる
- PDF 化後にページ数、空ファイル、SHA-256 を検証する

Word 文書の途中ページを構成ボードで分割する機能は MVP 後とする。初期版では、見出しや改ページを含めて1つの文書項目として扱う。

### 7.5 PDF アダプター `pdf-pass-through-v1`

- 対応: `.pdf`
- unit: 文書全体1件
- 再印刷や再エンコードをせず、snapshot 内の PDF を artifact として登録
- ページ数、暗号化、破損、極端なページサイズを検証
- 閲覧パスワードが必要な PDF は `requires-attention`
- 署名付き PDF は内容を改変しないが、最終結合後の署名効力は引き継げないことを警告
- 将来 `pageRange` によりページ抽出・分割へ拡張できる

## 8. 更新検知・履歴・差分

### 8.1 版管理

原稿形式にかかわらず次を保存する。

- source snapshot: 検知時のファイルハッシュ、サイズ、更新日時、保存可能な場合は原稿現物
- render version: 使用した adapter と版、設定 fingerprint、生成 artifact
- build pin: 最終出力で使用した snapshot / render version / item 順序

重複排除キーは `sourceId + source SHA-256` とする。ファイル名が同じでも内容が違えば別版、更新日時だけ変わって内容が同じなら同じ版を再利用する。

現物保存条件は現在の安全方針を維持し、dataDir が原稿フォルダー配下にあるときだけ snapshot に複製する。外部 dataDir ではメタデータと artifact の履歴は残すが、原稿現物は複製しない。

### 8.2 比較方式

比較を2層に分ける。

1. **構造差分**: unit の追加、削除、対応付け、ページ数変化
2. **見た目差分**: render artifact の物理ページを画像化して比較

| 形式 | 構造差分 | 見た目差分の単位 |
|---|---|---|
| Excel | シート追加・削除・変更 | シート内の各 PDF ページ |
| Word | 文書変更、ページ数変化 | 文書の各 PDF ページ |
| PDF | 文書変更、ページ数変化 | PDF の各ページ |

ページ追加・削除で全ページがずれて見えないよう、既存のページハッシュと類似度を使って前後ページを対応付けする。対応の確信度が閾値未満なら領域を推測せず「判定不能」とする。

比較キャッシュのキーは次で固定する。

```text
sourceId
+ beforeSnapshotId / beforeRenderVersionId
+ afterSnapshotId / afterRenderVersionId
+ comparisonAlgorithmVersion
```

自動比較と任意2版比較は別キャッシュとし、任意比較で自動比較基準を変更しない現行仕様を維持する。

差分詳細は、原稿名、比較元・比較先、項目対応、判定、対応確信度、対応方法、ページ数、説明をCSVへ出力できる。正式な承認ワークフローではなく、部署間レビューで共有する確認資料として扱う。

## 9. Local API V5

新UIは `/api/v2` 名前空間を使う。現在の API V4 は少なくとも汎用版の1メジャーリリース期間、互換変換して残す。

| Method | Endpoint | 用途 |
|---|---|---|
| GET | `/api/v2/state` | packs / sources / units / items / outputs の状態 |
| GET | `/api/v2/pack-templates` | 利用可能なテンプレート一覧 |
| POST | `/api/v2/packs` | 資料パック作成 |
| PATCH | `/api/v2/packs/{packId}` | 表示名・設定変更 |
| GET | `/api/v2/source-candidates?packId=...` | Excel / Word / PDF 候補一覧 |
| POST | `/api/v2/sources/register-batch` | 原稿の一括登録 |
| POST | `/api/v2/sources/unregister` | 原稿の登録解除 |
| POST | `/api/v2/sources/scan-updates` | 更新検知 |
| POST | `/api/v2/sources/render/start` | 選択または更新原稿の変換開始 |
| GET/POST | `/api/v2/jobs/status` | job 状態 |
| POST | `/api/v2/items/reorder` | 出力先別の構成を transaction 更新 |
| PATCH | `/api/v2/items/{itemId}` | 表示名、番号、出力設定 |
| GET | `/api/v2/history/diff-detail` | 自動または任意2版の比較 |
| POST | `/api/v2/history/diff/prepare` | 比較 job 開始 |
| GET | `/api/v2/outputs/readiness?packId=...` | 出力先別 readiness |
| POST | `/api/v2/outputs/build` | 最終 PDF 作成 |
| POST | `/api/v2/outputs/publish` | 公開先へコピー |

一括登録例:

```jsonc
{
  "packId": "pack_monthly",
  "sources": [
    {
      "relativePath": "営業部\\実績.xlsx",
      "ownerDepartment": "営業部",
      "required": true,
      "defaultTargetId": "main"
    },
    {
      "relativePath": "管理部\\説明.docx",
      "ownerDepartment": "管理部",
      "required": true,
      "defaultTargetId": "main"
    },
    {
      "relativePath": "参考\\約款.pdf",
      "required": false,
      "defaultTargetId": "appendix"
    }
  ]
}
```

V4 互換層では次の変換を行う。

- `category=ecm|bod|dmm` → 対応する built-in pack
- `workbookId` → sourceId。ただし sourceType が excel 以外なら V4 endpoint では拒否
- `pageId` → itemId
- `volume=ja-main|en-main` → pack の language + `targetId=main`
- `/api/workbooks/*` → Excel adapter に限定した `/api/v2/sources/*`

## 10. 最終出力と readiness

出力 fingerprint は次から作る。

- packId / targetId
- 有効 item の順序
- item が参照する artifact SHA-256 と pageRange
- ページ番号設定
- 組版プロファイル版
- テンプレート版と出力 override

原稿ファイルの current hash は fingerprint へ直接含めない現行方針を維持する。組版中に原稿だけが更新された場合は完成 PDF を保持し、直後に「変換PDF作成が必要」「最終PDFの再出力が必要」へ遷移させる。

readiness の blocker は形式非依存のコードへ揃える。

| code | 条件 |
|---|---|
| `required-source-missing` | 必須原稿が未登録または見つからない |
| `source-stale` | 現在 snapshot が最新 render に反映されていない |
| `render-failed` | 必須原稿の変換が失敗 |
| `item-unassigned` | 必須 item が未振り分け |
| `artifact-missing` | 参照 PDF が存在しない |
| `output-conflict` | 組版中に構成 fingerprint が変化 |

出力 manifest には各物理ページについて `sourceId / snapshotId / unitId / itemId / artifactId / originalPhysicalPageNumber` を記録し、「どの原稿のどの版から作られたか」を追跡可能にする。

## 11. 保存先

初回移行では既存ディレクトリを改名せず、コード上で意味だけを一般化する。既存データとの共存を優先する。

```text
_reportbinder\<language>\
  structure.json
  input-history\<sourceId>\<snapshotId>\
  content-pdf\<sourceId>\<renderVersionId>\
  comparisons\<sourceId>\...
  exports\archive\<packId>\<targetId>\<buildId>\
  templates\
```

新規環境でも `input-history` と `content-pdf` の名前は当面維持する。保存形式の変更と機能拡張を同時に行わず、運用中データの調査可能性を優先する。

## 12. schema 2 からの移行

### 12.1 自動変換

1. `structure.json` の schemaVersion 2 を検出する。
2. 現在と同じ方式で時刻付きバックアップを作る。
3. category ごとに `pack` を作成し、built-in template を割り当てる。
4. workbook を sourceType=`excel` の source へ1対1変換する。`workbookId` は保持する。
5. sheet ごとに unit を作成する。
6. page record を item へ1対1変換し、`pageId` を保持する。
7. `ja-main / en-main` を targetId=`main`、appendix を `appendix`、none を `unassigned` へ変換する。
8. contentPdf を artifact として登録し、存在・SHA-256・ページ数を検証する。
9. category 付き volumes 状態を packId + targetId の outputs へ変換する。
10. schema 3 を一時ファイルへ保存し、再読込・validation 後に atomic replace する。

変換不能なレコードを黙って捨てない。警告付きで `migrationIssues` に記録し、最終出力をブロックする。元の schema 2 バックアップは削除しない。

### 12.2 互換性の完成条件

- 移行前後で ECM / BOD / DMM の item 数、順序、出力先、番号設定が一致する
- 既存 content-pdf を再変換せず開ける
- 既存履歴2版の比較を開ける
- 移行直後の readiness と最終 PDF fingerprint が合理的に一致する
- schema 2 fixture、混在順 fixture、実案件を匿名化した fixture で自動試験する

## 13. 安全性と障害時の動作

### 13.1 ファイル安全性

- 原稿フォルダーから canonical path を算出し、範囲外・junction 越し・不正相対パスを拒否
- 許可拡張子は `.xlsx`, `.docx`, `.pdf` の明示リスト
- Office は一時コピーだけを開き、元原稿へ保存しない
- Office のマクロ実行を無効化し、外部リンクを自動更新しない
- COM 処理に source 単位の timeout を設け、失敗時に専用プロセスを回収
- PDF は暗号化、破損、ページ数、ファイルサイズ、ページ寸法を事前検査
- ログへ token、原稿本文、ネットワーク資格情報を出さない

### 13.2 同時利用

ロックは次の順序に統一する。

```text
source_<sourceId>.lock または output_<packId>_<targetId>.lock
  → 必要な短時間だけ structure.lock
```

変換や組版の長時間処理中は structure lock を保持しない。開始時と commit 前に source snapshot / output fingerprint を再確認し、競合した結果は公開しない。

### 13.3 診断情報

job には最低限次を記録する。

- jobId, packId, sourceId, sourceType, adapterId
- stage: validate / snapshot / inspect / render / verify / compare / commit
- stage ごとの開始・終了・所要時間
- Office / PDF ツールの終了コードと分類済みエラーコード
- 作成した snapshotId / renderVersionId / artifactId

## 14. 実装順序

各段階を独立した PR にし、常に ECM のリリース可能状態を保つ。

### PR 1: schema 3 と互換モデル（実装済み）

- schema 3 validator、migration、backup、fixture を追加
- server 内に pack/source/unit/item の読み取り helper を追加
- V4 レスポンスは互換 view から生成
- UI と変換処理はまだ変えない

受け入れ条件: 現行 selfcheck と E2E が同じ結果で通り、実案件コピーの移行前後で構成が一致する。

### PR 2: Excel アダプター境界（実装済み）

- `Get-ExcelFilesInSubmission`, `Register-Workbook`, `Render-Workbook` の形式固有処理を adapter へ移す
- orchestrator、共通 job 結果、artifact commit を実装
- Excel の出力と差分を現行同等に保つ

受け入れ条件: 同じ Excel から作った sheet PDF と最終 PDF の視覚比較が許容差内で一致する。

### PR 3: 資料パックと新UI（実装済み）

- `/api/v2/state`, templates, packs, sources, items, outputs を追加
- ECM / BOD / DMM を built-in template 化
- 画面名称を資料パック・原稿・変換PDFへ一般化
- pack 切替、形式バッジ、担当部署、必須設定を追加

受け入れ条件: 従来の ECM 手順に余分な必須入力がなく、既存 E2E を新UIで完走できる。

### PR 4: PDF 原稿（実装済み）

- `.pdf` 候補検出、登録、snapshot、検証、artifact 化
- Excel と PDF の混在構成、更新検知、履歴、最終出力

受け入れ条件: PDF 差し替え後に更新ありとなり、再取り込み後だけ出力可能になり、差分ページを開ける。

### PR 5: Word 原稿（実装済み）

- Word COM の専用プロセス処理
- `.docx` の一時コピー変換、timeout、ダイアログ抑止、エラー分類
- Word/PDF ページ対応と差分

受け入れ条件: 文章のみ、表、ヘッダー/フッター、履歴書型レイアウト、縦横混在、改ページ増減の各 fixture で変換・更新・比較が成立する。

### PR 6: 履歴UIの完全一般化（実装済み）

- sheet 固有名称を unit/page へ変更
- 自動比較と任意2版比較を全形式で統一
- ページ対応の確信度と判定不能表示

受け入れ条件: 3形式すべてで前回正式版との比較と任意2版比較が独立して動く。

### PR 7: 資料らしい仕上げ（実装済み）

- item からの PDF ブックマーク
- 表紙、セクション区切り、目次のテンプレート設定
- pageRange による PDF/Word 文書の一部分割
- 出力ファイル名パターンのUI

受け入れ条件: 目次・ブックマーク・ページ番号が item 順序変更後に再生成され、manifest と一致する。

### PR 8: 配布・実機E2E・運用移行（実装済み）

- Office 有無と対応版の診断
- オフライン/オンライン配布物の更新
- 既存 ECM、汎用混在、障害復旧、共有フォルダー同時利用の実機 E2E
- 利用者向け移行手順と管理者向け容量見積もり

## 15. E2E テストマトリクス

| シナリオ | 期待結果 |
|---|---|
| 既存 ECM schema 2 を開く | 自動バックアップ後に同じ順序・出力先で開く |
| Excel 2件 + Word 2件 + PDF 2件を登録 | 1つの構成ボードへ6項目以上が表示される |
| 複数選択で本体から補足へ移動後、1件だけ戻す | 選択・挿入位置・undo/redo が直感どおり動く |
| Excel にシート追加 | 手動順序を崩さず新項目だけ通知される |
| Word の文章を修正 | source stale → 再変換 → 変更ページ表示となる |
| Word に1ページ追加 | 後続全ページではなく挿入ページとして対応付けされる |
| PDF を差し替え | hash で更新検知し、任意2版比較が可能になる |
| 原稿更新中に最終組版 | 完成物を不正に最新扱いせず再出力必要になる |
| Word 変換が timeout | 前回 artifact を保持し、必須なら build をブロックする |
| PDF が暗号化 | 内容を推測せず利用者対応が必要と表示する |
| dataDir が原稿フォルダー外 | 原稿現物を履歴へ複製しない |
| build 後に履歴整理 | 使用版は pin され、未使用の古い版だけ削除される |
| 日本語/英語 pack | target 表示と言語別出力名が混線しない |

実機 E2E では自動化だけでなく、Chrome から一連の操作を行い、選択状態、通知、エラー後の復帰、ファイル選択ダイアログ、Word/Excel のバックグラウンド終了も確認する。

## 16. MVP の完成定義

次のすべてを満たした時点を「汎用資料パック MVP」とする。

- 既存 ECM / BOD / DMM が移行操作なしで動く
- 1つの pack に Excel / Word / PDF を混在登録できる
- 更新された原稿だけ再変換できる
- 3形式の自動比較と任意2版比較ができる
- 手動順序と出力先が原稿更新後も保持される
- 必須原稿の欠落・更新・変換失敗を readiness が正しくブロックする
- 最終 PDF から使用した原稿版まで manifest で追跡できる
- 原稿を変更せず、ローカルまたは共有フォルダーで完結する
- 配布版の実機 E2E を、文章中心、表中心、印刷帳票型、縦横混在の原稿で完走する

## 17. 初期対象外

- Office ファイルを ReportBinder 内で編集する機能
- クラウド上でのリアルタイム共同編集
- SharePoint / Box / Google Drive との直接同期
- 電子署名、承認ワークフロー、証跡の法的保証
- `.docm`、`.pptm`、画像、メールの取り込み（`.xlsm`と`.pptx`はマクロを実行しないOffice原稿として対応済み）
- スキャン PDF の OCR と意味的な文字差分
- Litera Compare 相当の Word/Excel ネイティブな赤入れ
- DMS のような全社文書検索・アクセス権管理

これらは、混在資料の反復作成という中核が安定した後に需要を見て追加する。

## 18. 現行コードへの変更マップ

| 現行箇所 | 変更方針 |
|---|---|
| `New-EmptyStructure`, `Test-StructureDocument`, `Repair-StructurePages` | schema 3 loader / validator / migration へ分離 |
| `Normalize-WorkbookCategory`, `Require-WorkbookCategory` | V4 互換層へ移し、本体は packId を必須化 |
| `Get-VolumeList`, `Get-OutputFileName` | template の targets / output 設定から解決 |
| `Get-ExcelFilesInSubmission` | 共通 candidate scanner + extension 別 adapter 判定 |
| `Register-Workbook(s)Batch` | `Register-Source(s)Batch` orchestrator と Excel 互換 wrapper |
| `Render-Workbook` | `excel-com-v1` adapter + 共通 artifact commit |
| `Get-StatePayload` | V5 state を正本とし、V4 workbook/page/volume view を派生 |
| history の `workbookId/sheetName` | 内部は sourceId/unitId、V4 query は変換 |
| `app/web/app.js` の workbook/sheet/category/volume | source/unit/pack/target の view model を追加し段階置換 |
| 最終組版 manifest | artifact と snapshot の由来を物理ページ単位で追記 |

## 19. 現在の実装進捗と次に着手する作業

PR 1〜8は実装済みである。schema 3とV4互換viewを維持したまま、Excel・Word・PowerPoint・PDFを同じ原稿一覧、更新検出、物理ページ構成、履歴差分、最終組版、配布・復旧へ接続した。WordとPowerPointは専用COMプロセス、一時コピー、マクロ無効、120秒timeout、所有プロセスだけの終了を実装している。

履歴比較は、Excelではシート名、Word/PDFでは完全一致ページの列をアンカーに項目対応を作る。途中へのページ挿入・削除を追跡し、確定できない区間は確信度付きの判定不能として表示する。自動比較と任意2版比較は同じ比較契約・画面を使用し、任意比較が自動比較基準を変更しない。

PR 7でitemからのPDFブックマーク、任意targetの出力、pageRangeによるWord/PDF部分利用、manifestの物理ページ追跡まで組版を一般化した。PR 8で配布、移行、復旧、実機E2E、大容量・混在原稿の性能受け入れを完成させ、汎用資料パックMVPの完成条件を満たした。

2026-08-07に、固定ECM / BOD / DMM表示から動的な資料パック選択・作成・複製・保管UIへの移行を完了した。続いて利用者定義ひな形の作成・編集・版管理を追加し、作成時点のひな形スナップショット、出力先表示名、新規ページの既定配置、原稿別配置指定、必須原稿の欠落・更新・変換失敗によるreadiness停止まで接続した。組み込みパックは互換動作を維持し、利用者定義ひな形の変更は既存パックへ遡及しない。

同日の次段階で、ひな形の複製・JSON移送・明示更新、1〜12件の動的出力先、必須提出資料リスト、資料パック横断進捗、自動PDF作成の再試行・容量保護・通知、マクロを実行しない`.xlsm`取り込みまで実装した。直接クラウド同期は初期対象外を維持し、まずローカルまたは同期済み共有フォルダーを安全に扱う。

部署横断の進捗は、未提出必須原稿の期限超過と7日以内を集計し、資料パックの要対応順へ反映する。提出済み原稿を必要原稿へ割り当てると期限アラートは自動解消される。

続く形式拡張で`.pptx`を追加した。PPTX内部構造とスライド数を登録時に検査し、専用PowerPoint COMワーカーからPDFへ変換する。スライドを物理ページとして履歴比較・ページ構成・最終組版へ流し、`.pptm`や内部マクロ、暗号化原稿は安全側で拒否する。

PR 1で追加済みの主要fixtureは次のとおり。

- schema 2 の ECM/BOD/DMM 混在データ
- main / appendix / none が混在し手動順序が変更済みのデータ
- content-pdf が一部欠落したデータ
- 履歴と build pin があるデータ
- schema 3 へ変換済みの期待値
- schema 3 を再度読み込んでも変化しない idempotence ケース

## 参考資料

- [Adobe Acrobat: Combine files](https://helpx.adobe.com/acrobat/desktop/edit-documents/combine-files/combine-files.html)
- [Adobe Acrobat: Compare PDFs](https://www.adobe.com/acrobat/how-to/compare-two-pdf-files.html)
- [Tungsten Power PDF: Create Assistant](https://docshield.tungstenautomation.com/PowerPDF/en_US/2025.3-jlrwz2ja2j/help/PowerPDF_help/PowerPDF_help/c_aboutcreateassistant.html)
- [Tungsten Power PDF: Compare documents](https://docshield.tungstenautomation.com/PowerPDF/en_US/2025.3-jlrwz2ja2j/help/PowerPDF_help/PowerPDF_help/t_comparedocuments.html)
- [Bundledocs: Create a PDF binder](https://www.bundledocs.com/blog/2017/3/9/create-a-pdf-binder-in-minutes-with-bundledocs)
- [Bundledocs: Document indexing](https://www.bundledocs.com/document-indexing/)
- [Litera Compare](https://www.litera.com/products/litera-compare)
- [Bluebeam Revu: Compare Documents vs. Overlay Pages](https://support.bluebeam.com/revu/features/compare-documents-vs-overlay-pages.html)

## 20. PR 7 実装結果（2026-08-06）

PR 7「資料らしい仕上げ」を実装した。

- itemタイトルと原稿名からPDFブックマークを生成
- 資料パック別に表紙、目次、原稿ごとの区切りページを設定
- itemの`pageRange`を互換pageへ同期し、PDF/Word変換物の一部ページだけを組版
- `{projectId}`、`{packName}`、`{targetName}`、`{yyyyMMdd}`対応の出力ファイル名パターンUI
- manifest schema 3で、生成ページを含む`physicalPages`と元PDFページを記録
- 実PDFで項目順、目次開始ページ、ブックマーク、部分抽出、物理ページ数を検証

## 21. PR 8 実装状況（2026-08-07）

配布・運用移行の安全網を実装した。

- Excel / Word 16.x以降の登録・COM起動・版、Java、PDFBox、PDF.js、保存先容量を画面から診断
- Officeを起動できないセッションでもPDF原稿を継続できる`limited`判定と対応案を表示
- オフライン、オンライン、共有フォルダー配布から開発・fixture資産を除外
- 配布種別、Java同梱有無、組版jar/PDFBoxのSHA-256を`release-manifest.json`へ記録
- 排他ロック、最終PDFのハッシュベース自動復旧、外部変更時の手動復旧停止を実ファイルで検証
- schema 1/2からの移行、障害復旧、同時利用、容量計画を運用ガイド化

対話Windowsセッションで実機E2Eを完走した。

- Excel 16.0 build 20228、Word 16.0 build 16.0.20228、Java、PDFBox、PDF.jsを画面診断で確認
- 文章中心、複雑な表、履歴書型、縦横混在PDF 4件、複数シートExcel 2件、文章・表・帳票・横向き・追加資料を含むWord 1件を登録・変換
- 19ページを本体12・補足7へ構成し、取りこぼしたページの再選択移動、元に戻す、やり直すを確認
- PDF 2件、Excel 1件、Word 1件だけを更新し、未更新3件を再変換しないことを確認
- Wordの既存割当を維持したまま追加1ページだけを未振り分けへ置き、補足へ追加
- 任意の旧版・新版比較でWordの変更3、追加1、変更なし2を左右表示・強調表示し、追加ページの片側表示を確認
- 更新後の最終出力は本体21ページ、補足15ページ。表紙、目次、区切り、ページ番号、更新Excel、追加Wordページを目視確認
- Office生成PDFを含む組版で発見した`COSStream has been closed`を、入力PDFを保存完了まで保持する修正で解消

大容量・混在原稿の受け入れ試験も完了した。PDF 6件（各20ページ）、Word 3件（各10ページ）、Excel 3件（各8シート）を実機Officeで処理し、原稿174ページを本体155・補足40ページの最終PDFへ組版した。初回変換144.633秒、更新3件の検出32.117秒、再変換・解析50.862秒、更新後出力9.479秒、ピーク作業領域約153 MiBで、定義したlarge基準をすべて満たした。条件と実測値は[`SCALE_BENCHMARK.md`](SCALE_BENCHMARK.md)に記録した。
