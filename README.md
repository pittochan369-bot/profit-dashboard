# 粗利ダッシュボード

Googleスプレッドシートの公開CSVを読み込み、本日・今月・店舗別の粗利データを表示する静的Webアプリです。

## 公開方法

GitHub Pagesで無料公開できます。

1. GitHubで新しいリポジトリを作成します。
2. このフォルダのファイルをリポジトリへpushします。
3. GitHubの `Settings` → `Pages` で、`Deploy from a branch` を選びます。
4. `main` ブランチの `/root` を選択して保存します。

## 必要なファイル

- `index.html`
- `styles.css`
- `app.js`

## データ取得

`app.js` 内のGoogleスプレッドシートURLからCSVを取得しています。シートは「リンクを知っている全員が閲覧可」にしてください。
