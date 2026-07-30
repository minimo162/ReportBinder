# Third-party dependency setup

ReportBinderには2種類の配布形態があります。

## GitHub取得直後

GitHubでは大容量の外部依存物を管理しません。最初に次を実行してください。

```text
ReportBinder\app\tools\install-thirdparty.cmd
```

またはPowerShellから実行します。

```powershell
cd ReportBinder\app\tools
.\install-thirdparty.ps1
```

既定実行では、次の4系統をすべて取得・検証・配置します。

```text
app\lib\pdfbox\pdfbox-app.jar
app\web\pdfjs\pdf.min.mjs
app\web\pdfjs\pdf.worker.min.mjs
app\lib\java\bin\java.exe
```

## 検証内容

### Apache PDFBox

PDFBox 2.0.37の`pdfbox-app.jar`と公式SHA-512ファイルをApache配布元から取得し、
`Get-FileHash -Algorithm SHA512`で一致を確認してから配置します。

### PDF.js

npmレジストリからpdfjs-dist 5.7.284のメタデータを取得し、
`dist.integrity`（SHA-512 SRI）でパッケージtarballを検証します。
検証済みtarballから次を展開します。

```text
build\pdf.min.mjs
build\pdf.worker.min.mjs
LICENSE
```

Windows 10/11に標準搭載される`tar.exe`を使用します。

### Eclipse Temurin JRE

Adoptium APIから最新のTemurin 17 GA / Windows x64 / HotSpot / JRE資産を取得します。
APIが返すSHA-256とZIPを照合し、展開後に`java.exe -version`を実行します。
実際に取得したリリース名、SemVer、SHA-256、取得元URLを次へ記録します。

```text
app\lib\java\JAVA_VERSION.txt
```

## Java取得オプション

既定では、システムJavaが存在してもオフライン配布に使えるポータブルJREを取得します。

```powershell
# 既存ファイルを再取得・更新
.\install-thirdparty.ps1 -Force

# システムJavaがある場合はポータブルJREを取得しない
.\install-thirdparty.ps1 -PreferSystemJava

# Java取得を完全に省略
.\install-thirdparty.ps1 -SkipJava
```

`-PreferSystemJava`または`-SkipJava`を使用した状態では、オフライン完結版を作成できない場合があります。

## 依存物の単独検証

```powershell
# PDFBoxとPDF.jsを確認
.\verify-thirdparty.ps1 -RequirePdfJs

# 実行可能なJavaも確認（ローカルまたはPATH）
.\verify-thirdparty.ps1 -RequirePdfJs -RequireJava

# オフライン完結版に必要なポータブルJREを確認
.\verify-thirdparty.ps1 -RequirePdfJs -RequirePortableJava
```

検証では、必要ファイルの存在・最小サイズ・JAR内の必須クラス・バージョン記録・Java実行可否を確認します。

## 配布ZIPの作成

```powershell
.\package-release.ps1
```

`package-release.ps1`は、オフライン完結版を作成する前に
`app\lib\java\bin\java.exe`、`release`、`NOTICE`、`JAVA_VERSION.txt`が揃っていることを必須確認します。
システムJavaだけではオフライン完結版を作成しません。

## オンライン導入版

オンライン導入版はポータブルJREを除外します。利用PCにJavaがない場合は、
展開後に`install-thirdparty.cmd`を実行してください。

## 手動配置

### Apache PDFBox

PDFBox 2.0.37の`pdfbox-app-2.0.37.jar`を取得し、公式SHA-512を確認したうえで次へ配置します。

```text
app\lib\pdfbox\pdfbox-app.jar
```

### PDF.js（V5では任意）

V5の画面プレビューはblob URL + iframeによるブラウザ内蔵PDFビューアを使用するため、
現行機能だけならPDF.jsは不要です。将来のPDF.jsベース視覚差分表示を見越して導入対象にしています。

```text
app\web\pdfjs\pdf.min.mjs
app\web\pdfjs\pdf.worker.min.mjs
```

### Java Runtime

Windows x64のJREを次の構成で配置します。

```text
app\lib\java\bin\java.exe
```

## ライセンス

Apache PDFBox、PDF.js、Eclipse Temurinのライセンス・NOTICEを配布物へ含めてください。
ReportBinder付属の情報は`THIRD_PARTY_NOTICES.md`を参照してください。
