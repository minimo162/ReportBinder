# ReportBinder V5 実装履歴

## 実機Word多様文書の差分局所化（2026-08-05）

- 実機Microsoft Word 16.0で、表を含まない2ページ業務文書、日本式履歴書、2段組み職務経歴書をPDF化し、文字置換・段落追加・行間変更・文字色変更・連絡先変更・職歴行追加・罫線変更・箇条書き追加を検証した。
- 複数ページの文字行を文書全体で照合し、追加段落だけを1領域へまとめ、改ページで次ページへ移動した既存文章は変更表示しないようにした。
- 2段組み文書では文字配置から列境界を求め、影響を受けた列だけを照合して追加箇条書きを1領域へ局所化した。
- 同じ文字内容の位置変化量を行ごとに比較し、行間や罫線による後続レイアウト移動を変位の開始箇所1領域へ集約した。
- 文書全体・レイアウト比較では上下5%の反復ヘッダー／フッターを除外し、本文の改ページとページ家具の変更を混同しないようにした。
- 差分詳細アルゴリズム版を`20`、RuntimeVersionを`2026.08.05.31`へ更新した。

## 実機Excel複合帳票の差分局所化（2026-08-05）

- 実機Microsoft Excel 16.0で、複数表・説明文・KPI・レビュー表・点線／破線／太線／二重線を含むA4横帳票をPDF化し、ブラウザPDF.js経路で回帰検証した。
- 左右に表が並ぶページでも、画像側が確認した表帯へPDF文字行LCSを限定し、行追加・逆向き削除を対象行1領域へ集約するようにした。
- 行高変更による印刷倍率変化で表上端の罫線を失わないよう、横罫線探索を表の公称行高に合わせて拡張した。
- 複数表が連結して見える帯を単一表の列追加と誤認しないよう、列構造モデルの罫線数を制限した。
- 離れた濃色タイトルを表ヘッダーとして採用せず、列幅変更の強調を対象表・対象列だけに限定した。
- 差分詳細アルゴリズム版を`19`、RuntimeVersionを`2026.08.05.30`へ更新した。

## 比較画像生成の一括・並列化（2026-07-30）

- 変更シートごとに新旧PDF用のJavaプロセスを2回起動する処理を廃止し、対象シート全部を1つのJVMで一括画像化するようにした。
- 独立したPDFの画像化を最大4並列で実行し、複数シート比較の待ち時間を短縮した。
- 変更マスク・枠線・M/A/Dタグを4枚のPNGとして保存する処理を廃止し、領域JSONからブラウザ上で描画するようにした。比較ページごとの保存画像は新旧2枚だけになった。
- 差分キャッシュのアルゴリズム版を8へ更新し、旧方式の未完了キャッシュを再利用しないようにした。

## 比較速度・誤検知・変更表示の改善（2026-07-30）

- 比較用ラスタライズを150dpiから120dpiへ下げ、元のPNGを再圧縮せずキャッシュへ再利用することで、比較画像の作成時間と一時メモリを削減した。
- 全画素SHA-256の完全一致だけでなく、抽出文字ハッシュと32x32知覚ハッシュを組み合わせる判定を追加し、アンチエイリアス等の人の目に見えない差を「変更」から除外した。
- 新旧画像の最大4pxの位置ずれを自動補正してから差分を取り、ページ全体が変更扱いになる誤検知を抑制した。
- 近接する変更領域を統合し、Mタグは1ページ12領域までに限定した。密集時は枠線だけを表示する。
- 広範囲かつ低密度の差分は全ページを囲まず、判定不能として左右比較を案内するようにした。

## UI導線・カテゴリ別出力名の修正（2026-07-30）

- 任意2時点の視覚比較を「最終PDF」画面下部から独立した「履歴・比較」メニューへ移し、主要機能として先頭に表示した。
- 未登録Excelの「表示分を選択」「選択解除」「選択したExcelを登録」を一覧下部から一覧上部の操作バーへ移した。
- BOD/DMMの最終PDF名が元Excel名のECM表記を引き継ぐ問題を修正し、選択中カテゴリをファイル名へ必ず反映するようにした。
- UI配置とカテゴリ別ファイル名の回帰検査をselfcheckへ追加した。

## レビュー修正（2026-07-30・第6回）

- PDFページ画像解析を`-Djava.awt.headless=true`付きで起動するようにした。タスク実行・リモートセッション・画面のない検証環境でもAWTがデスクトップ接続を要求せず、画像ハッシュ生成が安定する。
- 差分画像生成の実行失敗を、比較結果としての「判定不能」と分離した。従来はPDF不足・画像化失敗・書込み失敗なども`unknown`へ変換され、差分詳細全体が`ready`になって再試行できなかった。生成失敗は`failed`として保持し、全体再試行または対象シートの再選択で再生成できるようにした。
- 変更なしシートの遅延生成が失敗した場合も、全体の再試行対象へ含めるようにした。全体再試行時にも既存の差分詳細を識別子検証付きで読み戻し、`failed`状態を失わない。従来の全体処理は既存詳細を捨てたうえで変更なしシートを常に除外するため、一度失敗すると復旧経路がなかった。
- 別シートの生成ジョブへ相乗りした後、対象シートが`failed`または`deferred`のままなら再要求するようにした。自分で起動したジョブが失敗した場合は無駄な連続再実行をせず、その場で失敗表示へ確定する。絞り込みで失敗シートが選択された場合も再試行できる。
- 差分詳細の再試行ボタンが`MouseEvent`オブジェクトを`sheetKey`としてAPIへ送っていた不具合を修正した。クリック時は引数なしで全体再試行を呼ぶ。
- 失敗後に再試行が成功しても古いエラーメッセージが画面へ残る問題を修正した。描画のたびに状態欄を初期化し、現在のシート／ページ状態だけを表示する。
- 保存済み差分詳細を再利用する際、アルゴリズム版だけでなくworkbook・scope・比較元／比較先のsnapshotId・versionIdをすべて検証するようにした。短縮キャッシュキーの偶発衝突や不整合ファイルを別比較として表示しない。
- 差分PNGを返すAPIでも同じ完全一致検証を行い、詳細JSONが別比較・旧形式・破損状態なら画像を配信しないようにした。
- 差分キャッシュキーを32bit相当（8桁）から64bit相当（16桁）へ拡張した。短縮後の実パス長はMAX_PATH内に収まりつつ、任意2版比較を繰り返した際の衝突余地を大幅に下げた。
- キャッシュ形式変更に伴い差分詳細アルゴリズム版を`6`へ更新した。
- 差分ジョブが終端状態なのに`diff-detail.json`が欠けている場合、`generating`を返し続けるのではなく再試行可能な`failed`を返すようにした。
- 単一シートの遅延生成が外側の例外で中断した場合も、詳細全体と対象シートを`failed`へ確定し、「生成中」のまま残さないようにした。

## レビュー修正（2026-07-30・第5回）

- 遅延生成した1シートが差分詳細全体を「作成済み」にしていた問題を修正した。
  全シート生成が途中で失敗したあとに「変更なし」シートを1件開くと、`diff-detail.json` の
  `status` が `ready` へ上書きされ、`pending` のまま残ったシートを二度と生成できなく
  なっていた（`Start-DiffDetailJob` が「差分詳細は作成済みです」を返し、画面にも
  再試行ボタンが出ない）。完了時の状態は未生成シートの残件数から決めるようにした。
- 変更領域を「文字1つごと」ではなく「まとまり」で切り出すようにした。従来は連結領域の
  抽出後に余白を足していたため、数値1つの書き換えで桁ごとに5個の枠が重なって表示され、
  行がずれたページでは1ページあたり数千個の枠になっていた（実測3,772個）。
  変更画素を余白幅で膨張させてから連結判定するよう変更し、同じ検証で22個になった。
- 1ページあたりの強調領域を400個で上限とし、超えた場合は枠を描かずに
  「変更領域が多すぎるため、領域の強調を停止しました」と表示するようにした。
- ページ正規化の `DrawImageUnscaled` を明示的な転送先矩形付き `DrawImage` に置き換えた。
  `DrawImageUnscaled` は名前に反して元画像のDPIメタデータで拡縮するため、
  PDFBoxがPNGの解像度情報を書かない構成では、ページが拡大されて右下が欠けた画像が
  保存されるおそれがあった。
- 差分画像のファイル名を短縮した（`page-0001-before-overlay.png` → `0001-bo.png`）。
  実ワークスペースでのキャッシュパスは253文字で MAX_PATH まで7文字しか余裕が
  なかったため、236文字（余裕24文字）にした。
- キャッシュのレイアウトと領域抽出が変わるため、差分詳細アルゴリズム版を `5` へ更新した。
  適用後の初回表示では差分画像を作り直す。
- 上記すべてを `selfcheck.py` の回帰検査へ追加した。あわせて `DiffImageEngine.cs` が
  C# 5 互換のまま（`=>` / `$"` / `?.` / `out var` を使わない）であることも検査する。
  Windows PowerShell 5.1 の `Add-Type` は C# 5 の CodeDom でコンパイルするため。


## 履歴視覚比較の整合性・同時実行修正（2026-07-30）

- 画像ハッシュの世代と画面表示用content PDFの世代を厳密一致させ、別versionへの
  暗黙フォールバックを廃止した。比較可否はmanifestの古い保持フラグではなく、
  画像ハッシュと同一versionの実PDFが全シート分存在するかで判定する。
- 履歴一覧へ視覚比較可否・推奨version・利用不可理由を返し、比較資産が不足する版は
  選択できないようにした。
- 自動比較・履歴比較の保存結果をscope、比較元／比較先のsnapshotId・versionIdの
  完全一致で検索し、比較メタデータと画像キャッシュの両方を分離した。
- GETの差分詳細取得では比較結果を保存せず、POSTの準備処理でのみ保存するよう変更した。
  保存後の再読込・識別子検証に成功してから履歴イベントを記録する。
- 比較組合せ単位の共有ロック、snapshot/content PDFのleaseとheartbeat、終了時finally解放を追加した。
  履歴整理・content PDF整理とも同じ保守ロックを使い、lease作成との競合を防止する。
- 次回の自動比較基準に加え、直近の変更バッジが参照する比較元・比較先のsnapshotと
  同一versionのcontent PDFを専用pinで保護し、保持世代数を超えても判定対象と表示対象が残るようにした。
  旧基準の同一世代PDFが欠けている場合は、保存済みExcelからハッシュとPDFを一組で再生成し、
  今回版・前回版の両方を実ファイルで検証してから比較結果を保存する。
- 自動比較結果も保存後に再読込し、scopeと比較元／比較先のsnapshotId・versionIdを検証してから
  完了イベントを記録する。保存不整合がある比較を変更バッジへ載せない。
- 変更なしシートは初回の一括画像化から除外し、選択時だけ生成するようにした。
  別シートの生成ジョブが実行中だった場合は完了後に対象シートを再要求する。
- 差分詳細アルゴリズム版を`4`へ更新し、上記の整合性・排他・遅延生成をselfcheckへ追加した。

## 任意の過去2版の視覚比較（2026-07-30）

- 履歴画面で比較元と比較先を選び、保存済みの任意2版を既存の差分詳細ダイアログで
  左右比較・重ね合わせ・領域移動できるようにした。比較結果が無い組合せも、
  保存済みの画像ハッシュとcontent PDFからオンデマンド生成する。
- 履歴比較は短い専用キャッシュ名へ保存し、自動比較の
  `comparison-baseline.json`や最新PDFの変更バッジを変更しない。
- 自動比較と履歴比較で同じ2版を選んでも生成済み詳細が混線しないよう、
  キャッシュキーへ比較scopeを含め、差分詳細アルゴリズム版を`3`へ更新した。
- MAX_PATH対策後の短縮比較ファイル名をJSON内の`baselineSnapshotId`で検索するよう修正し、
  旧ファイル名形式を前提にした履歴差分検索の不整合を解消した。
- 履歴画面に選択中の比較元／比較先日時と入れ替え操作を追加し、日英それぞれで
  同じ期間の変更を担当者が素早く見比べられるようにした。

## セキュリティ強化（2026-07-30）

- `server.ps1`をランチャーを介さず単独起動した場合も、32バイトの
  `RandomNumberGenerator`からURL安全なトークンを生成するようにした。
- クエリ、ヘッダー、Cookieから受け取るトークンは、UTF-8文字列のSHA-256値を
  全バイト走査する固定時間比較で照合するようにした。PowerShellの既定の
  大文字小文字を区別しない`-eq`には依存しない。
- 上記2点を`selfcheck.py`の回帰検査へ追加した。
- 初回PDF作成直後など、まだ比較基準または比較結果が無い状態で差分詳細APIを開いた場合に、
  空の版IDを履歴パスへ渡してHTTP 400になる問題を修正した。正常な`unavailable`状態を返す。
- PDFジョブ完了後の画像解析と、直後に開始された次版PDF作成が重なった場合でも、
  現在版manifestの`previousSnapshotId`と保存済み画像ハッシュから比較基準を回復し、
  比較結果が欠落しないようにした。
- 各レンダー版へ`comparison-analysis.json`を保存し、比較できなかった場合も
  `unavailable`の理由を後から診断できるようにした。
- 深い提出フォルダでは比較JSONのパスが260文字を超え、判定は完了しているのに
  保存だけ失敗する問題を修正した。比較ファイル名、差分キャッシュキー、
  シートキー、画像格納フォルダをハッシュ短縮し、PowerShell 5.1でも扱える長さにした。
  キャッシュレイアウト変更に伴い差分詳細アルゴリズム版を`2`へ更新した。
- JSONの原子的保存で使用するGUID付き一時パスがMAX_PATHを超えた場合も、
  例外を捕捉して短い最終パスへの共有直接書き込みへ縮退するようにした。
- PDFBoxが生成するラスタ作業画像を深い履歴キャッシュ内ではなく、
  短いWindows TEMP配下へ作成し、完成PNGだけを履歴へ保存するようにした。

## 差分詳細・視覚比較（2026-07-29）

- 登録済みExcelの変更バッジをボタン化し、前回版と現在版を同じ画面で確認できる
  差分詳細ダイアログを追加した。
- シート単位の変更・追加・削除・判定不能・変更なしを一覧化し、件数バッジで絞り込めるようにした。
- content PDFを150 DPI・RGBで画像化し、RGB最大差18、微小ノイズ除去、16px以上の
  連結領域、6px余白の条件で変更領域を生成する機能を追加した。
- 変更は琥珀色＋実線＋M、追加は青色＋実線＋A、削除は赤色＋破線＋D、
  判定不能は灰色＋点線＋?で表示する。強調のON/OFFと0%／25%／50%の濃度に対応した。
- 左右のスクロールと倍率を同期し、左右比較／重ね合わせ、全体表示／幅合わせ、
  ページ移動、前後の変更領域への移動に対応した。
- 差分画像はバックグラウンドジョブで作成し、現在版・比較基準版・アルゴリズム版に紐づけて
  履歴内へキャッシュする。シート単位で進捗を保存するため、ダイアログを閉じても処理は継続する。
- `GET /api/history/diff-detail`、`POST /api/history/diff/prepare`、
  `GET /api/history/diff-page`を追加した。画像取得は履歴インデックスとmanifestを照合し、
  任意パスを受け付けない。
- 1100px以下では前回版／現在版をタブ切替にし、Esc、フォーカストラップ、起点への
  フォーカス復帰、色以外の記号・枠線・凡例を実装した。

## Codexレビュー修正（2026-07-29）

- 履歴・差分・自動処理APIから受け取る `workbookId` / `snapshotId` / `versionId` を、
  保存先の単一パス要素として検証するようにした。区切り文字、ドライブ指定、`..`、
  制御不能な長さを拒否し、管理フォルダ外の参照・書き込みを防止する。
- 履歴の保護情報を書き込めなかった場合に、成功応答を返さずエラーとして通知するようにした。
- ローカルHTTPサーバーのリクエスト本文を1 MiBに制限し、不正な `Content-Length`、
  本文が途中で切れたリクエスト、過大なリクエストを HTTP 400 で拒否するようにした。
- 画面側のAPI共通処理で、`FormData` / `Blob` をJSON文字列化して空の本文にしてしまう
  不整合を修正した。
- 「Excel登録・PDF作成」の変更列で、追加シートが表示されず「変更分のみ表示」からも
  漏れていた問題を修正した。ブック単位では `追加 Nシート`、ページ単位では `追加` と表示する。
- 1440px前後の画面では未登録Excelを約1/3、登録済みExcelを約2/3の横配置にし、
  情報量の多い登録済み表へ広く割り当てた。1320px以下では縦配置に戻し、変更列は230px確保する。
- Excel一覧ではファイル名を主情報として省略せず折り返し表示し、更新日時／PDF作成日時は
  ファイル名の下へ補助情報として配置した。横幅を日時列に奪われず、ファイルを識別しやすくした。
- 登録済みExcelの「変更」と「状態」を「PDF状況」へ統合した。PDF作成前は
  「差分は作成後に確認」、作成後だけ「前回PDFとの差分」を表示し、古い差分結果と現在の
  PDF作成要否が矛盾して見える問題を解消した。
- 上記の保存パス検証とHTTP本文制限を自己診断の回帰チェックへ追加した。

## 実運用での発見 その4（2026-07-29・P0）

**日本語サーバーを起動すると、稼働中の英語サーバーが強制終了される**（逆も同様）。
E側で作業中にJ側を立ち上げると画面が「Failed to fetch」になり、PDF作成中なら結果も失われる。

`launch.ps1` の `Stop-StaleUiServerProcesses` は、起動前に古いUIサーバーを掃除する。
その判定が「コマンドラインに server.ps1 を含むか」だけで `-Mode` を見ていなかったため、
別言語の稼働中サーバーまで `Stop-Process -Force` の対象になっていた。
`-AutoSchedulerPath` の子プロセスも除外されておらず、巻き添えで落ちていた。

- 停止対象を「同じ `-Mode` のUIサーバー」に限定
- `-RenderJobPath`（PDF作成の子）に加えて `-AutoSchedulerPath`（自動処理の子）も除外
- 自己診断に回帰チェックを追加

なお `app\tools\stop-reportbinder.cmd` は「ReportBinderを止める」ための道具なので、
従来どおり両言語をまとめて停止する（仕様どおり）。

## 実運用での発見 その3（2026-07-29・P0）

日本語ワークスペースで全17件のPDF作成が「引数の型が一致しません」で失敗する。
英語側は正常。違いは **J側は pages が0件で、ページを新規追加する経路を通る** こと。

`pages` は JSON 由来の `PSCustomObject` で構成されるのに、
新規ページだけ `[ordered]@{}`（`OrderedDictionary`）で作っていたため、
コレクション内に2種類の型が混在していた。
Windows PowerShell 5.1 はこの混在を `Sort-Object` などの型比較で
`ArgumentException`（Argument types do not match）として弾くことがある。
PowerShell 7 では再現しないため、AST解析でも pwsh 実行でも検出できなかった。

- 新規ページを `[pscustomobject][ordered]@{}` に変更し、`pages` の要素型を統一
- `Renumber-VolumeOrder` / `Apply-DefaultNumberingPerVolume` のソートキーを型安全化
  （`[double](Get-DataProperty $_ 'order' 0)` と `[string](Resolve-PageId $_)` に統一。
  後者は `$_.order` を直接参照しており、プロパティ欠損時に型が揺れていた）
- `Get-ErrorDetail` を強化。従来は `Exception.ToString()` が取れないと
  メッセージ1行だけになり原因追跡が不可能だった。
  例外の型・HResult・.NETスタック・内部例外（5段まで）・発生行・errorId を、
  取得できたものから必ず積むようにした

## 実運用での発見 その2（2026-07-29・P0）

`renders` フォルダは作られるのに `visual-hashes.json` が一切書かれない、という状態が残っていた。
手動で `PdfPageAnalyzer` を実行すると exit=0 でハッシュも正常に出るのに、
アプリ経由では必ず失敗扱いになる。

原因は **Windows PowerShell 5.1 の `2>&1` と `$ErrorActionPreference = 'Stop'` の組み合わせ**。
PS5.1 ではネイティブコマンドの stderr を `2>&1` で取り込むと ErrorRecord としてパイプラインに流れ、
`Stop` 設定下では NativeCommandError の終了エラーになる。
PDFBox は日本語フォントを含むPDFで
`Format 14 cmap table is not supported and will be ignored` を stderr に出すため、
解析が成功して `result.json` を書いていても、その手前で例外になり `catch` へ落ちていた。
結果、`render-manifest.json` だけが書かれ `visual-hashes.json` は書かれない。

- `Invoke-NativeCapture` を追加。呼び出しの間だけ `ErrorActionPreference` を `Continue` にし、
  終了コードと出力を文字列として回収する
- Java を呼ぶ5箇所すべてを置き換え
  （`PdfPageAnalyzer` / `BatchPdfSplitter` / `ReportPdfComposer` 2経路 / `java -version`）
- 自己診断に、`Invoke-NativeCapture` 以外の生の `2>&1` を禁止する回帰チェックを追加

Linux の検証環境では PowerShell を経由しないため再現しなかった。
PS7 はこの挙動を廃止しているため、AST解析や pwsh での実行でも検出できない。

## 実運用での発見（2026-07-29・P0）

有効化後の実機確認で、**画像ハッシュの解析が一度も実行されていない**ことが判明した。
`structure.json` の `lastRenderedSnapshotId` は正しく入るのに
`input-history\<workbookId>\<snapshotId>\renders` が作られず、
「変更」列が永久に空欄のままになる。エラーは出ない（解析失敗はPDF作成を止めない設計のため）。

原因は `Render-Workbook` 末尾の `$Script:PendingAnalysis` の扱い。
共有Excelを使う一括ジョブ（`KeepExcelOpen = $true`）では、
`Invoke-RenderJobFromFile` が `$deferredAnalyses` へ回収してから
Excelとロックを解放したあとに解析する設計だが、
`Render-Workbook` が `KeepExcelOpen` に関係なく `$Script:PendingAnalysis = $null` を
先に実行していたため、呼出元が回収する時点では常に `$null` だった。
画面からのPDF作成は必ずこの一括ジョブ経路を通るため、解析が一切走っていなかった。

- `KeepExcelOpen` のときは `$Script:PendingAnalysis` を消さずに呼出元へ渡すよう修正
- 未使用のまま残っていた `Invoke-DeferredAnalysis`（どこからも呼ばれていない死んだ関数）を削除
- 自己診断に、クリアが `KeepExcelOpen` ガードより後にあることの回帰チェックを追加

**適用後、各Excelを1回ずつPDF作成し直して比較基準を作り直す必要がある。**
それまでに作られた検知版には `renders` が無いため比較できない。

## 承認反映（2026-07-29）

Phase 0（入力履歴の業務承認）が確定したため、配布物の既定値と手順書を更新した。コードの変更はなし。

- `app\default-config.json` の `inputHistory.retainSourceVersions` を 2 → **5** に変更。
  正式版に使っていない検知版を最新5世代まで残す。
  正式版に使った版は pin により世代枠の対象外で、`sourceRetentionDaysAfterBuild` は `null`
  （期間では消さない）のままとし、不要になった時点で手動整理する。
- `docs\INPUT_HISTORY_POLICY.md` を追加。承認事項4件、`policy.json` の作り方、
  保持の仕組み（世代数と pin の違い）、設定値の配り方、手動整理の手順、段階的な有効化をまとめた。
- 設定値は共有側ではなく `%LOCALAPPDATA%\ReportBinder\config.json` を参照する。
  既に起動したことのあるPCでは既存キーが上書きされないため、
  `default-config.json` の変更は未起動のPCにだけ反映される点を明記した。

## レビュー修正 第4回（2026-07-29）

第3回修正版の再レビュー。第3回の修正内容はいずれも妥当で、
PowerShell 7.4.6 での構文解析（全7本エラー0）も通過した。以下2点のみ修正した。

- **配布ZIPの日本語ファイル名が文字化けする問題を修正（最優先・配布不能レベル）**。
  第3回のZIPはエントリ名をUTF-8バイトで書きながら
  「言語エンコーディングフラグ（汎用目的ビット11）」を立てていなかったため、
  日本語Windowsのエクスプローラーが名前をCP932と誤解し、
  `日本語管理.vbs` / `英語管理.vbs` / `docs\ReportBinder_UIUX改修指示書_V4.md`
  の3件が文字化けした状態で展開される。起動用VBSが読めなくなる。
  展開後の `selfcheck.py` もこの3件を「missing」として検出する。
  `app\tools\package-release.ps1` の `New-Utf8Zip`（V5-P0(#1)で導入済みの正規パッケージャー）
  と同じ方式で再パッケージした。
  自己診断に、パッケージャーが `New-Utf8Zip` を使い続けているかの回帰チェックを追加。
- 自動スケジューラーの停止確認間隔を 100ms → 500ms に変更。
  `ControlPath` は共有ドライブ上にあるため、100ms間隔だと利用者1人あたり毎秒10回の
  SMB `Test-Path` が常時発生する。親側は `WaitForExit(2500)` で待つため、
  500msでも正常終了と `finally` の後片付けは間に合う。

### 検証（第4回）

- PowerShell 7.4.6 で全7本の `.ps1` を構文解析：**エラー0**（第3回版・第4回版とも）
- `selfcheck.py`：合格。**ZIP再展開後も合格**（第3回のZIPは文字化けにより不合格）
- `app.js` 構文確認：合格 ／ 全 `.ps1` の UTF-8 BOM：確認済み
- ZIPエントリの UTF-8 フラグ（bit11）が非ASCII名で立っていることを確認
- **Windows上のExcel COM・VBSを含むE2E試験は未実施**

## レビュー修正 第3回（2026-07-29）

第2回修正の再レビューで見つかった後処理・キャッシュ境界の不具合を修正。データ形式の変更はなし。

- 履歴容量の60秒キャッシュを `language + input-historyルート` 単位に変更し、dataDir切替直後に旧ワークスペースの容量を表示する問題を修正
- `Reset-ConfigCaches` から履歴容量キャッシュも無効化
- TCP listenerの生成・開始失敗時でも自動スケジューラーを必ず停止するよう、サーバー初期化全体を `try/finally` 内へ移動
- スケジューラー子プロセスの10秒Sleepを100ms単位の停止確認へ変更。正常終了を待ってから最終手段としてKillするため、ロック解放・状態復旧の `finally` が通常は実行される
- スケジューラーの制御JSONと `.stop` ファイルを終了時・起動失敗時に削除し、`schedulerRunning` を実プロセスの生存確認で返すよう修正
- 巻き戻し済みトランザクションの `state\final-backups\<transactionId>` が無期限に残る問題を修正。巻き戻し完了時に削除し、30日ジャーナル掃除時にも安全側で削除
- PDF.jsの説明を `README.md`、`THIRD_PARTY_NOTICES.md`、`docs/THIRD_PARTY_SETUP.md`、`docs/API.md` まで統一
- 自己診断へ上記の回帰チェックを追加

## レビュー修正 第2回（2026-07-29）

ホットパスのディスクI/Oを中心とした性能不具合の修正。機能・データ形式の変更はなし。

- `Get-AppConfig` が呼ばれるたびに `config.json` を書き戻していた問題を修正。
  `Get-Paths` → `Get-WorkspacePath` 経由でほぼ全関数から呼ばれるため、
  検知版を1件参照するごとにローカル設定ファイルへの書き込みが発生していた。
  既定値を実際に補完したときだけ書き、結果は2秒キャッシュする（`Reset-ConfigCaches` で無効化）
- `Get-Paths` が共有ドライブ上の `common\paths.json` を毎回読み直していた問題を修正（同キャッシュ）
- `Get-WorkspacePolicy`（`Test-InputHistoryEnabled` / `Test-SourceRetentionEnabled`）を5秒キャッシュ。
  `Write-HistoryEvent` などのホットパスから多数回呼ばれていた
- `Get-InputHistorySizeMb` が `input-history` 配下を全再帰列挙していた問題を修正。
  `/api/state` のポーリングごとに実行されていたため60秒キャッシュとし、掃除後に強制再計測する
- `Get-LatestComparison` がブック1件ごとに `structure.json` を丸ごと読み直していた問題を修正。
  読込済みのブックを受け取れるようにし、`/api/state` の90回の再読込を1回にした
- `Stop-AutoSchedulerProcess` が定義のみで一度も呼ばれていなかった問題を修正。
  サーバー終了時に呼び出す（従来は親PID監視のみで、子プロセスが最大10秒残り `auto-owner` ロックを保持していた）
- 完了済み・巻き戻し済みの出力トランザクションのジャーナルが無期限に溜まる問題を修正（30日で削除。
  `manual-recovery-required` は担当者が確認するまで残す）
- `app/web/pdfjs/README.md` の記述を実装に合わせて訂正。
  V5時点のプレビューは blob URL + iframe（ブラウザ内蔵ビューア）であり、PDF.js は使用していない。
  同梱ファイルは今後の視覚差分表示に向けた先行配置である旨を明記した
- 自己診断に上記の回帰チェックを追加

### 検証（第2回）

- PowerShell 7.4.6 で `server.ps1` ほか全 `.ps1` の構文解析：**エラー0**
- ロック用スクリプトブロックの動的スコープ衝突をAST走査で検査：実害のある取り違えなし
- 設定・パス・ポリシーのキャッシュを実行して動作確認（初回作成／定常読取で書き込みなし／
  キー欠落時の補完と書き戻し／`Save-AppConfig` での無効化／承認フラグの反映）
- `PdfPageAnalyzer` を実PDFで実行し、同一PDF＝同一ハッシュ・1行変更で別ハッシュを確認
- **Windows上でのExcel COM / VBS起動を含むE2E試験は未実施**

## レビュー修正（2026-07-29）

- カスタムTCPサーバーのクエリ文字列が PowerShell により配列へ展開され、正しいトークンでも API が HTTP 403 になる問題を修正
- レンダリングロックのラッパーが PowerShell の動的スコープで自身を再帰実行し、ブックロックへ自己衝突する問題を修正
- Windows PowerShell 5.1 で `List[object]` を `@()` 変換するとページ追加時に「引数の型が一致しません」になる問題を修正
- レンダリング失敗ログへ例外型・PowerShellスタック・発生行を記録し、原因を追跡できるように改善
- `ReportPdfComposer.jar` を再ビルドし、欠落していた `PdfPageAnalyzer.class` を同梱
- 再ビルド用の一時 `classes` フォルダーを配布物へ残さないように修正
- リリースZIP名を `ReportBinder_V5_*` に修正
- 自己診断に、クエリ文字列コレクション・解析クラス同梱・一時クラスフォルダーの回帰チェックを追加

## Stage 1 — 基盤

Phase 1A 以降が依存する土台のみ。**この時点では動作は V4.1 と同じ**で、履歴・差分・自動化の機能はまだ入っていない。

### 決定事項

- Excel COM は**言語ごとに1ジョブ**に制限する（`locks\render-engine.lock`）。

### server.ps1

| 区分 | 内容 |
|---|---|
| ハッシュ表現 | `Normalize-FileHash` を追加。`New-Sha256`（64文字・大文字・prefixなし）と `Get-Sha256Text`（`sha256:`+小文字）の混同を防ぐ |
| ID規則 | `New-RbId` / `New-RbVersionId` / `New-UniqueDirectory` を追加。`versionId` を秒単位からミリ秒+GUID8へ変更（同一秒の2回処理で世代フォルダが混ざる問題） |
| ロック | `Try-AcquireLockHandle` / `Release-LockHandle` を追加（`Invoke-WithLock` は body 終了でハンドルを閉じるため、tick をまたいだ保持ができない）。`Invoke-WithRenderLock` を追加し、取得順序を `render-engine` → `render_<workbookId>` に統一。旧 `workbook_<id>.lock` は廃止 |
| フィールド所有権 | `Render-Workbook` から `currentExcel*` の書き込みを削除。コミット時のコピー対象からも除外。これらは `Scan-Updates` の専有 |
| レンダリング成功時 | ロック内で最新の `currentExcelHash` と突き合わせ、一致すれば `rendered-unchecked`、不一致なら `excel-updated`（ページは `stale`）。content-pdf は保存する |
| レンダリング失敗時 | `Set-WorkbookRenderError` に `AttemptedSnapshotId` / `AttemptedHash` を追加。試行した版が現在版でなければ `render-error` にせず `excel-updated` のままにし、ページの status も上書きしない。失敗情報は `lastRenderErrorSnapshotId` / `lastRenderErrorHash` に残す。3つの catch 経路すべてに `Get-LastRenderAttemptFor` 経由で試行版を渡す |
| 環境指紋 | `Get-RenderEnvironmentFingerprint` を追加（`excelVersion` / `osVersion` / `activePrinter` / フォント / 印刷プロファイル版。`pcName` は診断情報として主判定から除外）。`Reset-RenderEnvironmentForJob` をレンダリングジョブの開始点2箇所に挿入し、サーバー稼働中1回だけだった取得を毎ジョブに変更 |
| 承認ポリシー | `Get-WorkspacePolicy` / `Test-InputHistoryEnabled` / `Test-SourceRetentionEnabled` を追加。`<dataDir>\common\policy.json` を参照する。`sourceRetentionApproved` は `dataDir` が提出フォルダ配下にあることも確認する |
| 設定マージ | `Merge-ConfigDefaults` を追加し、`Get-AppConfig` がローカル config と `default-config.json` をキー単位で再帰マージするよう変更。既存利用者に新キーが届くようにする |
| 矛盾設定 | `Get-AutoRenderSettings` が `autoRender.enabled=true` かつ履歴未承認のとき自動処理を無効化して警告する |
| スケジューラー | `param` に `-AutoSchedulerPath` / `-ParentProcessId` を追加（Stage 3 で使用） |

### その他

- `app/default-config.json` に `inputHistory` / `autoRender` を追加（`autoRender.enabled` は `false`）
- `docs/POLICY_SAMPLE.json` を追加
- `app/tools/selfcheck.py` に V5 の回帰チェックを追加（フィールド所有権、3つの catch 経路、ハッシュ形式、ID粒度、ロック順序、承認ポリシー、設定マージ、環境の毎ジョブ取得）

## Stage 2 — Phase 1A: 検知版スナップショット

- `input-history\<workbookId>\<snapshotId>\` に manifest / source.xlsx / renders / pins / leases を配置
- `Ensure-SnapshotMetadata`（重複排除）と `Capture-RenderInput`（入力の確保）を分離。同じ内容が再登場して現物が削除済みでも、レンダリング入力を必ず確保する
- 保存手順はハッシュ検証つき（コピー後・移動後の2回）。`manifest.json` を最後に書き、その存在を完成マーカーにする
- コミットは compare-and-set。**PC の時刻ではなく提出ファイルの ticks / size / hash で新旧を判定する**
- pins（`final-pdf_<buildId>` / `comparison-baseline` / `manual`）と leases（期限つき）をファイルの作成・削除で管理
- 掃除は2段階。段階1は `source.xlsx` のみ削除（`sourceRetentionDaysAfterBuild` を適用）、段階2は pin が無いときだけフォルダごと削除
- 常時保護：`currentSnapshotId` / `pendingSnapshotId` / `lastRenderedSnapshotId` / baseline
- `Remove-WorkbookContentPdfs` を保持世代数＋pins/leases 対応に変更。レンダリング中の個別削除は廃止
- 登録解除では履歴を消さない（`Remove-WorkbookContentPdfsAll` は履歴無効時のみ）
- 縮退モード（現物保存が未承認）：`state\jobs\<captureId>\` に一時コピーし、ジョブ完了時と `ephemeralCopyMaxAgeMinutes` 超過で必ず削除

## Stage 3 — Phase 2A: 検知パイプライン

- `server.ps1 -AutoSchedulerPath -ParentProcessId` でスケジューラーモードを起動（独立スクリプトにしない。既存関数を共用するため）
- **親PID監視**で子プロセスを終了。tick は10秒でスリープせず、ブックごとに直列180秒待たない
- ブック単位の所有権 `locks\auto-owner_<workbookId>.lock` を `Try-AcquireLockHandle` で tick をまたいで保持
- 対話型Excelの判定は `MainWindowHandle != 0` と `~$<ファイル名>.xlsx` の存在（単純なプロセス有無では止まりすぎる）
- レンダリングは既存の子ジョブ方式（`Start-RenderJob`）に委譲
- 起動時リカバリー：`rendering` / `ready` を `waiting` へ戻し、消えた `pendingSnapshotId` を破棄

## Stage 4 — Phase 2B: 画像ハッシュと比較

- `PdfPageAnalyzer.java` を追加。**RGB / 150dpi**、ピクセル入力形式を `[幅4][高さ4][RGB行優先]`（alpha 除外）に固定。PNG は保存しない
- 環境指紋は `excelVersion` / `osVersion` / `activePrinter` / フォント / 印刷プロファイルから算出（`pcName` は診断情報で主判定に含めない）
- 比較は3経路：環境一致なら保存済みハッシュ、不一致なら前回 `source.xlsx` を現在環境で再レンダリング、現物が無ければ打ち切って今回版を基準にする
- `Render-SnapshotForComparison` は structure / content-pdf / volumes / 出力履歴を一切変更しない。PDF だけ捨て、`visual-hashes.json` は残す
- baseline は可変ポインタ（`input-history\<workbookId>\comparison-baseline.json`）＋pin。**解析に成功したときだけ**更新する
- 解析は PDF 作成のクリティカルパスの外。失敗しても PDF 作成は成功扱い（判定は `unknown`）

## Stage 5 — Phase 1B / 2C: アーカイブ・レイアウト履歴・出力トランザクション

- 最終PDFアーカイブ（`final.pdf` / `manifest.json` / `metadata.json` / `sha256.txt`）。一時フォルダで完成させてから `buildId` フォルダへ移動する冪等方式
- `metadata.json` の `renderEnvironments` は fingerprint をキーにした辞書（1つの最終PDFが複数PCの出力を含みうるため）。組版したPCは `composerEnvironment`
- レイアウト投影スナップショットと限定復元。適用するのは `title` / `volume` / `enabled` / `order` / `orderManual` / `numberingMode` / `numberingManual` のみ
- 復元前プレビューと `pre-restore` スナップショット（復元の取り消し用）
- `Invoke-FinalBuildTransaction`：**単体出力もまとめて出力も同じ経路**。8段階のジャーナルを副作用の前に書き、対象は `pageCount > 0` の volume のみ
- 復旧は `oldPdfHash` / `newPdfHash` と実ファイルの照合を正とする。どれとも一致しなければ自動で上書きせず `manual-recovery-required`
- 起動時復旧も `final-build_<category>.lock` を取得する

## UI

- 登録済みExcelに「変更」列（`変更 Nシート` / `変更なし` / `判定不能` / `比較不能`）
- ページ構成に変更バッジと「変更分のみ表示」トグル
- Excel画面に自動処理の状態（静止待ち・保留理由・今すぐ実行）と容量警告
- 最終PDF画面に変更履歴タイムライン
- 「本体・補足をまとめて出力」を `/api/final/build-all` の1回呼び出しへ変更（本体だけ成功する状態を作らない）

## 追加API

```
GET  /api/history/timeline      GET  /api/history/snapshots    GET  /api/history/diff
GET  /api/history/content-pdf   POST /api/history/pin          POST /api/history/unpin
GET  /api/layout/snapshots      POST /api/layout/restore/preview  POST /api/layout/restore
GET  /api/final/archives        POST /api/final/build-all
GET  /api/auto/state            POST /api/auto/run-now
```

## 承認と機能ゲート

`<dataDir>\common\policy.json` が未作成、または `inputHistoryApproved` が `false` の間は、
履歴・差分・自動処理はすべて動作せず、V4.1 と同じ挙動になります（サンプル: `docs/POLICY_SAMPLE.json`）。

```
inputHistoryApproved    : 検知版の記録と差分機能
sourceRetentionApproved : 提出Excel現物の保存（false なら一時コピーのみ）
```

`sourceRetentionApproved` は、`dataDir` が提出フォルダ配下にあることも起動時に確認します。
**なお policy.json は技術的なアクセス制御ではなく運用フラグです。**共有 `dataDir` に書き込める利用者は編集できます。

## 検証状況

- `app/tools/selfcheck.py`：**通過**（`PdfPageAnalyzer.class` 未同梱の警告あり）
- 括弧バランス：ベースラインとの差分 0
- **PowerShell の実行・構文チェックは未実施**（Linux 検証環境に `pwsh` が無いため）。Windows での起動確認が必要
- Excel COM / PDF レンダリング / VBS 起動を含むエンドツーエンド試験は未実施
