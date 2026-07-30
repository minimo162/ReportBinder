# Third-party dependency setup

ReportBinderには2種類の配布形態があります。

## オフライン完結版

次が配置済みです。

```text
app\lib\pdfbox\pdfbox-app.jar
app\web\pdfjs\pdf.min.mjs
app\web\pdfjs\pdf.worker.min.mjs
app\lib\java\bin\java.exe
```

共有ドライブ上のJREを初めて起動すると、SMB経由のクラスロードにより数秒余分にかかる場合があります。

## オンライン導入版

Windowsのエクスプローラーから次を実行します。

```text
ReportBinder\app\tools\install-thirdparty.cmd
```

またはPowerShellから実行します。

```powershell
cd ReportBinder\app\tools
.\install-thirdparty.ps1
```

スクリプトは不足している依存物を次へ配置します。PDF.jsはV5の現行プレビューには不要ですが、将来機能との互換性のため導入対象に含めています。

```text
app\lib\pdfbox\pdfbox-app.jar
app\web\pdfjs\pdf.min.mjs
app\web\pdfjs\pdf.worker.min.mjs
app\lib\java\bin\java.exe  ※システムJavaがない場合
```

`ReportPdfComposer.jar`は同梱済みです。通常利用にJDKは不要です。

## 手動配置

### Apache PDFBox

PDFBox 2.0.xの`pdfbox-app-*.jar`を取得し、SHA-512を確認したうえで次へ配置します。

```text
app\lib\pdfbox\pdfbox-app.jar
```

### PDF.js（V5では任意）

V5の画面プレビューは blob URL + iframe によるブラウザ内蔵PDFビューアを使用するため、現在の機能だけならPDF.jsは不要です。将来のPDF.jsベース視覚差分表示を見越して同梱・導入しています。配置する場合はgeneric buildから次を配置します。

```text
app\web\pdfjs\pdf.min.mjs
app\web\pdfjs\pdf.worker.min.mjs
```

### Java Runtime

Javaがない場合は、Windows x64のJREを次の構成で配置します。

```text
app\lib\java\bin\java.exe
```

オフライン版の基準ビルドはEclipse Temurin 17.0.19+10です。

## ライセンス

Apache PDFBox、PDF.js、Eclipse Temurinのライセンス・NOTICEを配布物へ含めてください。ReportBinder付属の情報は`THIRD_PARTY_NOTICES.md`を参照してください。
