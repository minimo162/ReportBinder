# Third-party notices

ReportBinderは次の第三者コンポーネントを利用します。

## Apache PDFBox

- 用途: PDF結合、ページ移動、ページ番号付与
- 配置: `app/lib/pdfbox/pdfbox-app.jar`
- ライセンス: Apache License 2.0
- バージョン情報: `app/lib/pdfbox/PDFBOX_VERSION.txt`

Apache PDFBoxのLICENSEおよびNOTICEを社内再配布物へ含めてください。

## PDF.js

- 用途: 将来のPDF.jsベース視覚差分表示向けの先行同梱（V5の画面プレビューはブラウザ内蔵ビューアを使用）
- 配置: `app/web/pdfjs/`
- ライセンス: Apache License 2.0
- 同梱ライセンス: `app/web/pdfjs/LICENSE`
- バージョン情報: `app/web/pdfjs/PDFJS_VERSION.txt`

## Eclipse Temurin / OpenJDK

オフライン完結版にのみ同梱します。

- 実装: Eclipse Adoptium Temurin
- バージョン: 17.0.19+10
- イメージ: Windows x64 JRE / HotSpot
- 配置: `app/lib/java/`
- ライセンス: GNU General Public License version 2 with the Classpath Exception
- 同梱通知: `app/lib/java/NOTICE`
- ビルド情報: `app/lib/java/release`

Classpath Exceptionを含む正確な条件は、同梱JREのNOTICEおよびライセンス文書を優先してください。
