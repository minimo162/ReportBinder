# ReportBinder V5

部署別Excelを、シート単位のcontent-pdfへ変換し、本体・補足の最終PDFへ組版するWindows向けツールです。

V4では、最終PDFの鮮度判定、ECM/BOD/DMMの分離、ページ順の復旧、複数PCからの同時更新、UI導線を全面的に改修しています。

## 起動入口

```text
日本語管理.vbs
英語管理.vbs
```

日本語と英語は別ワークスペースです。画面右上の言語表示は案内用で、自動的に別プロセスへ切り替えません。

## V3以前から更新する場合

V4の初回起動では、各言語の`structure.json`をschemaVersion 2へ移行します。共有フォルダで複数人が利用している場合は、次の手順を守ってください。

1. 利用者へ停止を告知します。
2. 旧版のブラウザとReportBinderプロセスをすべて終了します。30分アイドル終了を待つ場合は35分以上の保守枠を確保します。
3. 旧ランチャーを退避し、V4だけを起動できる状態にします。
4. V4を起動します。
5. `_reportbinder\ja`と`_reportbinder\en`に`structure.json.v1.bak`が作成されたことを確認します。
6. 移行後は、旧版から同じ`_reportbinder`を開かないでください。

移行は`structure.lock`内で1回だけ行われます。対応上限を超えるschemaVersionを検知した場合、V4は起動を拒否します。

## 配布形態

### オフライン完結版

`app\lib\java`を含みます。PDFBox、PDF.js、ポータブルJREを配置済みのため、社内ネットワークからの取得なしで利用できます。V5の通常プレビューはブラウザ内蔵PDFビューア、差分詳細はPDFBoxで生成したページ画像を使用します。PDF.jsは将来拡張向けの先行同梱です。

H:などの共有ドライブから初めてJavaを起動すると、SMB経由のクラスロードにより数秒余分にかかることがあります。

### オンライン導入版

`app\lib\java`を含みません。初回に次を実行します。

```text
app\tools\install-thirdparty.cmd
```

システムJavaがない場合は、ポータブルJREも自動配置します。プロキシ等で取得できない場合は`docs\THIRD_PARTY_SETUP.md`を参照してください。

## 入力履歴・差分機能

入力履歴と差分機能は既定で有効です。`_reportbinder\common\policy.json` は不要で、既存ファイルがあっても参照しません。
提出Excelの現物は、管理データ（`dataDir`）が提出フォルダ配下にある場合だけ保持します。
保持世代数、容量上限、自動処理の設定は `docs\INPUT_HISTORY_POLICY.md` を参照してください。

登録済みExcelの「変更 Nシート」「追加 Nシート」「削除 Nシート」「変更なし」
バッジを選ぶと、差分詳細を開きます。前回版と現在版の左右比較、同期スクロール・倍率、
重ね合わせ、強調濃度、前後の変更領域への移動を利用できます。差分画像の初回作成は
バックグラウンドで継続し、2回目以降は保存済みキャッシュを表示します。
変更シート内でも画素ハッシュが完全一致するページは差分領域解析を省略し、比較画像の作成時間を短縮します。

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

提出フォルダを選ぶと、次が自動設定されます。

```text
管理データ: 提出フォルダ\_reportbinder
出力先:     提出フォルダ\出力
```

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

登録対象は提出フォルダ直下の`.xlsx`です。子フォルダ、`_reportbinder`、`出力`は対象にしません。

### ページ構成

ページを次の表へドラッグします。

```text
本体 / 補足 / 出力しない
```

- `Alt+↑/↓`で1行移動
- `Alt+Shift+↑/↓`で先頭・末尾へ移動
- 行またはファイル名をダブルクリック／クリックしてプレビュー
- `シート名順に並べ替え`で各表の中だけを半角数字順へ戻す

手動順のvolumeへ新規ページが追加された場合は、既存の手動順を壊さず末尾へ追加し、通知と一時ハイライトで示します。

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
