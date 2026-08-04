# ReportBinder V5

部署別Excelを、シート単位のcontent-pdfへ変換し、本体・補足の最終PDFへ組版するWindows向けツールです。

V4では、最終PDFの鮮度判定、ECM/BOD/DMMの分離、ページ順の復旧、複数PCからの同時更新、UI導線を全面的に改修しています。

## 起動入口

```text
日本語管理.cmd
英語管理.cmd
```

日本語と英語は別ワークスペースです。画面右上の言語表示は案内用で、自動的に別プロセスへ切り替えません。
VBScriptはWindowsで段階的に廃止されるため、通常の起動入口をCMDへ移行しました。既存ショートカットの移行猶予としてソースと従来ZIPにはVBSも残しますが、新しい共有フォルダー用配布物には含めません。

## 共有フォルダー用の完成フォルダー

リポジトリ直下の次のファイルを実行します。

```text
共有フォルダー用フォルダー作成.cmd
```

開発用selfcheckにはログやキャッシュがないクリーンなリポジトリが必要なため、この作成処理からは分離しています。実際に配布するPDF.js、PDFBox、ポータブルJREと完成フォルダー構成を必ず検証した後、OneDrive同期の影響を受けない
`%LOCALAPPDATA%\ReportBinder\release\ReportBinder_共有フォルダー用_yyyyMMdd_HHmmss`
へ、共有フォルダーへそのままコピーできるオフライン完結版を作成し、作成先をエクスプローラーで開きます。

作成物にはポータブルJRE、PDFBox、PDF.jsを含みます。一方、個人設定、ログ、キャッシュ、管理データ、出力PDF、GitHub設定、テストfixture、Javaビルドソース、パッケージ作成ファイル、移行用VBSは含めません。作成されたフォルダー全体を共有フォルダーへコピーしてください。

PowerShellから作成先を指定する場合は次を実行します。作成先にReportBinderの元フォルダー配下は指定できません。OneDriveや共有フォルダーを指定してディレクトリ移動が拒否された場合は、内容コピー、完成物の再検証、失敗時削除へ自動的に切り替わります。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\app\tools\package-release.ps1 -SharedFolderOnly -OutputDir "D:\配布作業"
```

## 配布形態

### オフライン完結版

`app\lib\java`を含みます。PDFBox、PDF.js、ポータブルJREを配置済みのため、社内ネットワークからの取得なしで利用できます。V5の通常プレビューはブラウザ内蔵PDFビューア、差分詳細はPDF.jsで表示中のページだけをブラウザ内描画・比較します。比較用PNGは作成しません。

共有フォルダの起動入口は小さなバージョン情報だけを確認します。初回または更新時にアプリ一式を
`%LOCALAPPDATA%\ReportBinder\runtime\versions\<version>` へコピーし、通常はローカル版の
PowerShell、Web画面、Java、PDFBox、PDF.jsを使用します。更新がない起動では共有側の大量の
小ファイルを読み直しません。

### オンライン導入版

`app\lib\java`を含みません。初回に次を実行します。

```text
app\tools\install-thirdparty.cmd
```

システムJavaがない場合は、ポータブルJREも自動配置します。プロキシ等で取得できない場合は`docs\THIRD_PARTY_SETUP.md`を参照してください。

## 入力履歴・差分機能

入力履歴と差分機能は既定で有効です。`_reportbinder\common\policy.json` は不要で、既存ファイルがあっても参照しません。
提出Excelの履歴、比較用PDF、ページ構成は利用者ごとのローカルプロジェクトに保持します。
同じ提出フォルダを複数人が選んでも、登録状態・比較基準・ページ構成・出力は互いに影響しません。
保持世代数、容量上限、自動処理の設定は `docs\INPUT_HISTORY_POLICY.md` を参照してください。

登録済みExcelの「変更 Nシート」「追加 Nシート」「削除 Nシート」「変更なし」
バッジを選ぶと、差分詳細を開きます。前回版と現在版の左右比較、同期スクロール・倍率、
重ね合わせ、強調濃度、前後の変更領域への移動を利用できます。前後PDFはそのまま配信し、
表示中の1ページだけをPDF.jsで描画してWeb Workerで差分解析します。比較用PNGの保存や
全シートの事前生成は行いません。画素ハッシュが完全一致するページは差分領域解析も省略します。
行・列の追加や削除、行高・列幅の変更、1ページに収める縮尺の変化がある場合は、縦横それぞれの
縮尺と局所的な位置ずれを補正してから差分領域を抽出します。補正は無補正より十分に一致度が上がる
場合だけ採用し、セル内容や図形の一部変更を位置ずれとして過剰補正しません。部分的な拡大縮小は
変更された範囲として表示し、位置ずれだけでページ全体を変更扱いにしません。

## 初回起動と設定

共有アプリ側には読み取り専用の既定値だけを置きます。

```text
共有アプリ\app\default-config.json
```

利用者ごとの現在値は次へ保存します。

```text
%LOCALAPPDATA%\ReportBinder\config.json
```

ローカル設定がない初回だけ`default-config.json`をコピーし、以後はローカル設定だけを更新します。共有側の既定値へ現在値を書き戻しません。

提出フォルダを選ぶと、共有側は提出Excelの読込元としてだけ使用し、作業場所は自動的に
利用者ローカルへ設定されます。

```text
管理データ: %LOCALAPPDATA%\ReportBinder\projects\<提出フォルダ識別子>\data
通常出力:   %LOCALAPPDATA%\ReportBinder\projects\<提出フォルダ識別子>\output
共有発行（日本語）: 提出フォルダ\MMdd_HHmmss_J_Windowsユーザー名
共有発行（英語）:   提出フォルダ\MMdd_HHmmss_E_Windowsユーザー名
```

既存の `提出フォルダ\_reportbinder` と `提出フォルダ\出力` は参照もコピーもしません。
新しい利用者または移行未完了の利用者は、空のローカル管理領域から開始します。移行完了マーカーが
ある既存利用者のローカルデータはそのまま保持します。旧共有管理フォルダーはアプリから自動削除しません。

## 基本操作

画面上部の4ステップに沿って進めます。

```text
1. 提出フォルダ
2. Excel登録・PDF作成
3. ページ構成
4. 最終PDF
```

作業カテゴリは全画面共通です。

```text
ECM / BOD / DMM
```

カテゴリは登録済みExcel、ページ構成、最終PDFの状態・ファイル・manifestまで完全に分離されます。

### Excel登録・PDF作成

1. 未登録Excelを選び、`選択したExcelを登録`を押します。
2. 登録直後は登録したExcelが選択済みになります。
3. `PDF作成`を押します。未選択の場合はPDF作成が必要なExcelだけを処理します。
4. 印刷設定を更新した後などは`すべて再作成`を使います。

OpenXML形式のExcelは、ローカル一時コピーへ標準印刷設定を先に書き込んでからExcel COMで開きます。
シートごとの重いPageSetup通信を省き、複数シートは1回のPDF出力にまとめます。ブックを開く時間と
Excel自身のPDF変換時間は残ります。内訳は利用者別ローカル管理領域の`logs\render_*.json`に
`timingsMs`として記録します。

登録対象は提出フォルダ直下の`.xlsx`です。子フォルダ、`_reportbinder`、`出力`は対象にしません。

### ページ構成

Excelの表示中ワークシートは、シート名が日本語・英数字・記号のどれでも登録できます。
新しく見つかったページは、まず `未振り分け（出力しない）` に入ります。必要なページだけを
`本体` または `補足` へ移してください。明示的に振り分けたページだけが最終PDFに含まれます。

```text
未振り分け（出力しない） → 本体 / 補足 → 最終PDF
```

- `Alt+↑/↓`で1行移動
- `Alt+Shift+↑/↓`で先頭・末尾へ移動
- 行またはファイル名をダブルクリック／クリックしてプレビュー
- `Excelのシート順に整える`で、各表の割り当てを変えずにブック内のシート順へ戻す
- 既に振り分け済みのページ構成は、アプリ更新後もそのまま維持する

手動で並べた表へ新規ページが追加されても、既存の順番は変更しません。新規ページは
未振り分けの末尾に追加し、画面上の案内で振り分けが必要なことを示します。

### 最終PDF

本体・補足ごとに次の状態を表示します。

| 状態 | 表示 |
|---|---|
| 出力条件を満たさない | 最終PDFを出力できません |
| 出力後に入力が変わった | 最終PDFの再出力が必要 |
| 記録はあるがファイルがない | 前回出力が見つかりません |
| 初回出力前 | 最終PDFは未出力 |
| fingerprint一致・ファイル実在 | 最終PDFは最新です |

前回PDFが実在する場合、再出力が必要な状態でも`前回出力を開く`を利用できます。
通常の最終PDFはローカルへ作成します。`共有発行`を押すと、提出フォルダ直下へ
日本語は `MMdd_HHmmss_J_Windowsユーザー名`、英語は
`MMdd_HHmmss_E_Windowsユーザー名` の発行フォルダを作ります。PDFのファイル名は
ローカル最終PDFと同じままです。同じ秒に同じ利用者が発行した場合は、フォルダ名へ連番を付けて既存フォルダを上書きしません。

## 鮮度判定

最終PDFを作成したときの`builtFingerprint`と現在の入力fingerprintを比較します。fingerprintには次を含めます。

- ページID、順番、出力先、番号表示
- content-pdfの相対パス、サイズ、更新時刻
- `lastRenderedVersionId`
- 最終組版プロファイルのバージョン

元Excelの`currentExcelHash`は組版fingerprintに含めません。組版中に元Excelだけが更新された場合、完成したPDFは保持しつつ、完了直後に`PDF作成が必要`および`最終PDFの再出力が必要`へ遷移します。

カテゴリを省略した最終PDF系APIはHTTP 400で拒否します。空categoryを全カテゴリ一致として扱うfail-open経路はありません。

## 同時利用とロック

structureの更新は、読込→変更→保存を`structure.lock`内で行います。

ロック取得順は次に固定しています。

```text
workbook_<id>.lock または volume_<volume>_<category>.lock
    → 必要な短時間だけ structure.lock
```

逆順は禁止です。Excel COM処理中やJava組版中は`structure.lock`を保持しません。

最終PDFは組版開始時と完了時にfingerprintを再比較します。ページ順またはcontent-pdfが組版中に変わった場合、一時PDFを破棄し、既存の最終PDFを差し替えません。

## manifestと一時PDF

カテゴリ別の決定的な名前を使用します。

```text
manifest_ja-main_ecm.json
~building_ja-main_ecm.pdf
volume_ja-main_ecm.lock
```

manifestは不具合調査用に保持し、同じ言語・volume・categoryの次回組版で上書きします。

## 印刷・余白設定

Excelからcontent-pdfを作る際は、一時コピー上で次へ標準化します。

```text
上・下余白: 0.8 cm
左・右余白: 1.2 cm
ヘッダー・フッター: 空
ヘッダー・フッター余白: 0
水平中央: ON
垂直中央: OFF
拡大縮小: 1ページに収める
印刷範囲: Excel側の印刷範囲を尊重
```

最終PDFでは奇数ページを右へ、偶数ページを左へ0.2cm移動します。ページ番号はArial 8pt、`- 2 -`形式です。番号非表示ページも通番には含めます。

## 診断・停止

```text
app\tools\diagnostic-ja.cmd
app\tools\diagnostic-en.cmd
app\tools\stop-reportbinder.cmd
```

通常は30分の無操作で終了します。PDF作成ジョブが残っている場合は終了を遅らせます。
通常起動ではPowerShell画面と開発用ログを表示しません。起動失敗時の調査用ログは従来どおり
`%LOCALAPPDATA%\ReportBinder\logs` に保存します。

## 回帰確認

```text
python app\tools\selfcheck.py
```

正常時は次で終了します。

```text
selfcheck ok
```

selfcheckは、category必須化、言語別6キー、移行バックアップ、transaction、fingerprint、カテゴリ別manifest、10刻みorder、UIの死に要素削除、現行印刷設定を確認します。

## 配布ZIPの作成

Windows PowerShellで実行します。

```powershell
app\tools\package-release.ps1 -OutputDir C:\Temp\ReportBinderRelease
```

オフライン版とオンライン導入版を作成し、実運用config、過去ログ、一時job、ダウンロードキャッシュを除外します。

## 第三者ソフトウェア

同梱物・ライセンスは`THIRD_PARTY_NOTICES.md`を参照してください。オフライン版に含むJavaはEclipse Temurin 17.0.19+10で、GPLv2 with Classpath Exceptionの条件に従います。
