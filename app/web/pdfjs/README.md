# PDF.js 配置先

このフォルダには Mozilla PDF.js の次の2ファイルを配置します。

- `pdf.min.mjs`
- `pdf.worker.min.mjs`

配置は `app\tools\install-thirdparty.cmd` で自動化できます。オフライン完結版には同梱済みです。
社内ネットワークで利用する前提のため、CDN参照は使わず、ファイルをローカル配置します。

## 現在の利用状況（V5時点）

**画面右側のPDFプレビューは、現時点では PDF.js を使っていません。**
サーバーから取得したPDFを blob URL にしてiframeへ渡し、ブラウザ内蔵のPDFビューアで表示しています
（`web/app.js` の `previewPage`）。このため、これらのファイルが無くてもプレビューは動作します。

同梱している理由は、今後の「PDF.js によるページ画像の視覚差分表示」で使う前提の先行配置です。
サーバー側は配置状況を `/api/state` の `pdfjsPresent` / `pdfjsMode` で返しますが、
V5 時点では画面側にこれを読む処理はありません。

配布サイズ（約1.7MB）を削りたい場合は、このフォルダごと削除してもV5の動作には影響しません。
その場合は `docs/THIRD_PARTY_SETUP.md` と `THIRD_PARTY_NOTICES.md` の記載も合わせて更新してください。
