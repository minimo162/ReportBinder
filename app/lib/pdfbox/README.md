# PDFBox CLI

このフォルダは最終PDF組版用の専用Java CLIを置く場所です。

必要ファイル:

- `pdfbox-app.jar` - Apache PDFBox 2.0.x の app jar
- `ReportPdfComposer.jar` - `src/ReportPdfComposer.java` をビルドしたもの

依存ファイルの取得:

```powershell
cd app\tools
.\install-thirdparty.ps1
```

手動で配置する場合は、Apache PDFBox 2.0.x の `pdfbox-app-*.jar` を取得し、このフォルダに `pdfbox-app.jar` という名前で置いてください。

ビルド手順:

```powershell
cd app\lib\pdfbox
.\build.ps1
```

実行例:

```powershell
java -cp "ReportPdfComposer.jar;pdfbox-app.jar" ReportPdfComposer --manifest C:\path\manifest_ja-main.json
```

`ReportPdfComposer` は、content-pdf を結合し、本文を必要に応じて左右へ移動し、`numberingMode=visible` のページだけ Arial 8pt のページ番号を物理ページ下部中央に描画します。ページ番号は `- 2 -` の形式です。番号非表示ページも通番には含めます。


## パンチ穴余白

暫定PDFは左右1.0cmの余白で生成します。最終PDFでは `punchShiftPt` により、奇数ページを右へ0.3cm、偶数ページを左へ0.3cm移動します。これにより、見た目上は奇数ページが左1.3cm/右0.7cm、偶数ページが左0.7cm/右1.3cm相当になります。
