# Webアプリ送信用 GAS API

このGASは、GitHub Pagesなどで動くWebアプリから送信依頼を受け取り、`GmailApp.sendEmail` でメールを送ります。

## デプロイ

1. Google Apps Scriptで新しいプロジェクトを作成する。
2. `Code.js` と `appsscript.json` を配置する。
3. `デプロイ` → `新しいデプロイ` → 種類を `ウェブアプリ` にする。
4. 実行ユーザーは `自分` を選ぶ。
5. アクセスできるユーザーは、運用方針に合わせて選ぶ。
6. デプロイ後に発行される `/exec` のURLをWebアプリの「送信用GAS API URL」に貼る。

## API

`POST` で以下のJSONを送ります。ブラウザ側ではプリフライトを避けるため、`Content-Type: text/plain` でJSON文字列を送ります。

```json
{
  "senderName": "筑波大学アイドル研究会 Bombs!",
  "replyTo": "info@bombstsukuba.com",
  "messages": [
    {
      "rowNumber": 2,
      "email": "sample@example.com",
      "subject": "件名",
      "body": "本文"
    }
  ]
}
```

1回のリクエストは100件までです。
