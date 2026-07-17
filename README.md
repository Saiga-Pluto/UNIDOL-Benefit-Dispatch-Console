# UNIDOL Benefit Dispatch Console

CSV / TSV / Excelをアップロードして、メール本文のプレビュー、送信対象の選択、GAS API経由のメール送信を行うWebアプリです。

## ディレクトリ

| パス | 内容 |
| --- | --- |
| `public/` | GitHub Pagesなどで公開する静的Webアプリ |
| `gas-send-api/` | Webアプリから呼び出す送信用GAS API |
| `spreadsheet-gas/` | 既存のスプレッドシート操作用GASの保存版 |

## 入力ファイル

1行目は見出し、2行目以降が送信対象です。

| 列 | 内容 |
| --- | --- |
| A | メールアドレス |
| B | 宛名 |
| C1 | 1つ目の特典見出し |
| C2以降 | 1つ目の特典のサブ名、担当者名など |
| D2以降 | 1つ目の特典リンク |
| E1 | 2つ目の特典見出し |
| E2以降 | 2つ目の特典のサブ名、担当者名など |
| F2以降 | 2つ目の特典リンク |
| G2以降 | `TRUE` なら初期状態で送信対象 |

1つ目の特典は、D列のリンクが入っている場合だけ本文に入ります。
2つ目の特典は、F列のリンクが入っている場合だけ本文に入ります。
C列とE列のサブ名は任意です。空の場合でも、リンクがあれば見出しとリンクだけ本文に入ります。

状態欄では、`1つ目リンクなし`、`1つ目サブ名なし`、`2つ目リンクなし`、`2つ目サブ名なし` を同時に確認できます。

要確認として件数に数え、行を薄い赤色で表示するのは、`メールなし` と `宛名なし` の2つだけです。

## ローカル確認

`public/index.html` をブラウザで開けば画面確認できます。

ローカルサーバーで確認する場合:

```sh
python3 -m http.server 5173 --directory public
```

## GitHub Pages

GitHub Pagesで公開する場合は、リポジトリの Pages 設定で公開元を対象ブランチの `/public` にしてください。

`main` ブランチのGAS API方式では、`gas-send-api/` のGASを別途デプロイし、Webアプリ画面の「送信用GAS API URL」に貼り付けます。

## Gmail APIログイン送信の実験ブランチ

`experiment-gmail-api-login` ブランチでは、送信用GAS API URLを使わず、Googleログインで取得したアクセストークンを使ってGmail APIから送信します。

この方式では、送信者はログインしたGoogleアカウントのGmailアドレスになります。Webアプリに必要なのは、Google Cloudで作成したOAuthクライアントIDです。

### 必要なGoogle Cloud設定

1. Google Cloud Consoleでプロジェクトを作成する。
2. Gmail APIを有効化する。
3. OAuth同意画面を設定する。
4. OAuthクライアントIDを作成する。
5. アプリケーションの種類は `ウェブ アプリケーション` を選ぶ。
6. `承認済みの JavaScript 生成元` に公開URLを追加する。
   ローカル確認では `http://127.0.0.1:5173` などを追加する。
7. 作成された `クライアントID` をWebアプリの `Google OAuth クライアントID` に貼る。

### 注意点

この方式では `https://www.googleapis.com/auth/gmail.send` の権限を使います。運用対象が広い場合、GoogleのOAuth同意画面やアプリ確認が必要になる可能性があります。

既存のGAS API URL方式は `main` ブランチに残しています。

## 送信用GAS API URL

送信用GAS API URLは、`gas-send-api/` のApps Scriptをウェブアプリとしてデプロイしたときに発行される `/exec` で終わるURLです。

入手手順:

1. Google Apps Scriptで新しいプロジェクトを作成する。
2. `gas-send-api/Code.js` と `gas-send-api/appsscript.json` を配置する。
3. 右上の `デプロイ` から `新しいデプロイ` を選ぶ。
4. 種類で `ウェブアプリ` を選ぶ。
5. 実行ユーザーを `自分` にする。
6. アクセスできるユーザーを運用に合わせて選ぶ。
7. デプロイ後に表示される `ウェブアプリ URL` をコピーする。
8. Webアプリ画面の `送信用GAS API URL` に貼り付ける。

このURLを通じて、Web画面で作った送信内容をGASに渡し、GAS側が `GmailApp.sendEmail` でメールを送信します。
