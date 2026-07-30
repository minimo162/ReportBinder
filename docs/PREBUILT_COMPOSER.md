# ReportPdfComposer.jar 同梱版について

この配布版には `app/lib/pdfbox/ReportPdfComposer.jar` を同梱しています。
そのため、最終PDF作成に必要な専用Java CLIを利用するだけであれば JDK / javac は不要です。

必要なもの:

- Java 実行環境の `java` コマンド
- `app/lib/pdfbox/pdfbox-app.jar`
- `app/lib/pdfbox/ReportPdfComposer.jar`

`install-thirdparty.cmd` 実行時に `javac が見つからない` と出ても、
`ReportPdfComposer.jar` が存在する場合は通常利用に支障ありません。

再ビルドしたい場合のみ、JDK を入れて次を実行してください。

```powershell
cd app\lib\pdfbox
.\build.ps1
```


## PDFBox 2.x フォント読込互換性

この同梱 `ReportPdfComposer.jar` は、PDFBox 2.x の `PDType0Font.load(...)` の差異を吸収するため、
Arialフォント読込を単一のオーバーロードへ直接固定しない実装にしています。

これにより、実行時の以下のエラーを避けます。

```text
java.lang.NoSuchMethodError: PDType0Font.load(PDDocument, File, boolean)
```

## この版の組版方式

この版の `ReportPdfComposer.jar` は、最終PDF作成時にPDFBoxの `PDFMergerUtility` でcontent-pdfを結合し、その後に本文の左右移動とページ番号だけを追加します。
旧方式のように各ページをフォーム化して貼り直す処理を減らしているため、ページ数が多い最終PDFで処理時間を短縮しやすくなっています。

ページ番号はArial 8ptの `- n -` 形式で、番号を表示しないページも通番に含めます。
