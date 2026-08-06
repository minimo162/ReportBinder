# 大容量・混在原稿ベンチマーク

## 目的

Excel、Word、PDFが混在する実案件相当の負荷で、初回変換、更新検出、更新原稿だけの再変換、ページ構成、最終PDF出力に必要な時間と一時容量を測る。合成データは機密情報を含まず、同じ条件を再生成できる。

この試験は対話中のWindowsユーザー資格情報でExcelとWordを実際にCOM起動する。サービス、非対話セッション、Officeを起動できないサンドボックスでは実行しない。

## 負荷

| ティア | PDF | Word | Excel | 想定用途 |
|---|---:|---:|---:|---|
| medium | 4件×12ページ | 2件×6ページ | 2件×5シート×32行 | 日常的な確認 |
| large | 6件×20ページ | 3件×10ページ | 3件×8シート×45行 | リリース受け入れ |

PDFとWordには文章、画像、表、罫線を含める。Excelには数式、表示形式、折り返し、表罫線、複数シートを含める。初回出力後、PDF 1件、Word 1件、Excel 1件だけを更新し、3件だけが再変換されることを検証する。

`sourceLogicalPageCount`は原稿として構成したページ数である。最終PDFの`finalOutputPageCount`には各資料パックの表紙、目次、原稿区切りなども含まれるため一致しない。

## 実行

Python、Node.js、`@oai/artifact-tool`を含む`node_modules`を用意し、サインイン中のWindowsデスクトップから実行する。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\app\tools\scale-benchmark.ps1 `
  -Tier large `
  -PythonPath "C:\path\to\python.exe" `
  -NodePath "C:\path\to\node.exe" `
  -NodeModulesPath "C:\path\to\node_modules" `
  -ReportPath ".\docs\benchmarks\scale-benchmark-windows-YYYYMMDD.json"
```

通常は作業領域を自動削除する。目視確認が必要な調査時だけ`-KeepWorkspace`を付け、確認後に表示された`ReportBinderScale_<GUID>`フォルダーを削除する。

## 合格基準

largeティアの標準受け入れ基準は次のとおり。時間値はこの実機で回帰を検出するための上限であり、異なるPCではベースラインを別に記録する。

- 12原稿をすべて登録し、初回変換失敗が0件
- 更新した3原稿だけを検出し、再変換失敗が0件
- 本体・補足の両PDFが最新状態で生成される
- 初回変換300秒以内、更新検出60秒以内、更新変換・解析120秒以内
- 初回・更新後の最終出力がそれぞれ60秒以内
- ピーク作業領域1 GiB以下、処理中の空き容量10 GiB以上
- 代表ページの文章、画像、表、罫線、ページ番号をPoppler描画で目視確認

## 2026-08-07 実測結果

Windows実機、Excel 16.0 build 20228、Word 16.0 build 16.0.20228、OpenJDK 17.0.20でlargeティアを実行し、全基準に合格した。

| フェーズ | 時間 | 結果 |
|---|---:|---|
| fixture生成 | 15.567秒 | PDF 6、Word 3、Excel 3 |
| 登録 | 4.482秒 | 12/12件 |
| 初回変換 | 144.633秒 | 12/12件、失敗0 |
| 初回最終出力 | 10.725秒 | 原稿174ページを構成 |
| 更新検出 | 32.117秒 | PDF・Word・Excel各1件 |
| 更新変換・解析 | 50.862秒 | 3/3件、失敗0 |
| 更新後最終出力 | 9.479秒 | 本体155、補足40ページ |

ピーク作業領域は160,419,771 bytes（約153 MiB）、管理データは85,883,355 bytes、最終PDFは合計63,549,516 bytesだった。小さな合成原稿に高解像度画像を持たせるため、原稿サイズ比は34.95倍となる。容量計画には倍率だけでなく絶対量と、実際の業務原稿での再測定を使う。

代表として本体の表紙・PDF由来ページ・Excel由来ページ、補足の表紙・Word由来ページ、および更新後Excelプレビューを描画し、欠落、文字切れ、罫線崩れがないことを確認した。目視確認後、一時作業領域は削除した。機械可読の実測値は[`benchmarks/scale-benchmark-windows-20260807.json`](benchmarks/scale-benchmark-windows-20260807.json)に保存する。
