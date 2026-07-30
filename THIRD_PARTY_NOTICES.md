# Third-party notices

ReportBinderは次の第三者コンポーネントを利用します。

## Apache PDFBox

- 用途: PDF結合、ページ移動、ページ番号付与
- 配置: `app/lib/pdfbox/pdfbox-app.jar`
- 対象系列: PDFBox 2.0.x
- 既定バージョン: 2.0.37
- ライセンス: Apache License 2.0
- バージョン・検証情報: `app/lib/pdfbox/PDFBOX_VERSION.txt`

`install-thirdparty.ps1`はApache配布元のSHA-512ファイルを取得し、JARのハッシュを照合してから配置します。
Apache PDFBoxのLICENSEおよびNOTICEを社内再配布物へ含めてください。

## PDF.js

- 用途: 将来のPDF.jsベース視覚差分表示向けの先行同梱（V5の画面プレビューはブラウザ内蔵ビューアを使用）
- 配置: `app/web/pdfjs/`
- 既定バージョン: pdfjs-dist 5.7.284
- ライセンス: Apache License 2.0
- 同梱ライセンス: `app/web/pdfjs/LICENSE`
- バージョン・検証情報: `app/web/pdfjs/PDFJS_VERSION.txt`

`install-thirdparty.ps1`はnpmレジストリが返す`dist.integrity`（SHA-512 SRI）でパッケージtarballを検証し、
検証済みパッケージから`pdf.min.mjs`、`pdf.worker.min.mjs`、`LICENSE`を取り出します。

## Eclipse Temurin / OpenJDK

オフライン完結版には、Windows x64のポータブルJREを同梱します。

- 実装: Eclipse Adoptium Temurin
- Java系列: 17
- イメージ: Windows x64 JRE / HotSpot
- 配置: `app/lib/java/`
- ライセンス: GNU General Public License version 2 with the Classpath Exception
- 同梱通知: `app/lib/java/NOTICE`
- ビルド情報: `app/lib/java/release`
- 取得版・SHA-256・取得元: `app/lib/java/JAVA_VERSION.txt`

既定では、システムJavaが存在してもポータブルJREを取得します。
`-PreferSystemJava`を指定した場合のみシステムJavaを優先し、`-SkipJava`でJava取得を明示的に省略できます。
取得したZIPはAdoptium APIが返すSHA-256と照合してから展開します。

Classpath Exceptionを含む正確な条件は、同梱JREのNOTICEおよびライセンス文書を優先してください。
