# Make Report App

アンケート結果レポート生成アプリケーション

## デプロイ手順（Render）

### 1. 環境変数の設定

Renderのダッシュボードで以下の環境変数を設定してください：

#### 必須の環境変数

- `PORT` - デフォルト: `3000`
- `NODE_VERSION` - 推奨: `18.0.0` 以上

#### Google API認証情報

Google Cloud Consoleでサービスアカウントを作成し、JSON形式の認証情報を取得してください。

**オプション1: Render Secretsを使用（推奨）**

1. Renderダッシュボードで「Secrets」セクションに移動
2. `/etc/secrets/credentials.json` に認証情報JSONファイルをアップロード

**オプション2: 環境変数として設定**

`GOOGLE_CREDENTIALS_JSON` 環境変数に、認証情報JSONの内容を文字列として設定

### 2. Build Command

```bash
npm install
```

### 3. Start Command

```bash
npm start
```

### 4. ヘルスチェック

デプロイ後、以下のエンドポイントで動作確認：

```
GET /health
```

レスポンス例：
```json
{
  "status": "healthy",
  "timestamp": "2025-11-13T15:41:00.000Z",
  "uptime": 123.456,
  "nodeVersion": "v18.0.0"
}
```

## ローカル開発

### 前提条件

- Node.js 18.0.0以上
- Google Cloud Platform サービスアカウント（Google Sheets API、Google Drive API有効化済み）

### セットアップ

1. 依存関係のインストール

```bash
npm install
```

2. 認証情報の配置

`/etc/secrets/credentials.json` にGoogle Cloud Platformのサービスアカウント認証情報を配置

3. サーバー起動

```bash
npm start
```

サーバーは `http://localhost:3000` で起動します。

## トラブルシューティング

### デプロイ失敗: "Cause of failure could not be determined"

**原因:** Google API認証情報が正しく設定されていない可能性があります。

**対処法:**

1. Renderのログを確認：
   - "Failed to initialize Google API clients" エラーがないか確認
   - "Kuromoji tokenizer built successfully" が表示されているか確認

2. 認証情報の確認：
   - `/etc/secrets/credentials.json` が正しくマウントされているか
   - または `GOOGLE_CREDENTIALS_JSON` 環境変数が設定されているか

3. サービスアカウントの権限確認：
   - Google Sheets API が有効化されているか
   - Google Drive API が有効化されているか
   - 対象のスプレッドシートに対する読み取り/書き込み権限があるか

### Kuromojiの初期化失敗

**原因:** 辞書ファイルが見つからない

**対処法:**

ログに以下が表示されていることを確認：
```
Trying dictionary path: /path/to/node_modules/kuromoji/dict
Kuromoji tokenizer built successfully.
```

表示されていない場合は、`node_modules` が正しくインストールされていません。

## 主な機能

- Google Sheetsからアンケートデータ取得
- レポート生成（グラフ、統計、分析）
- AI分析（OpenAI GPT-4）
- ワードクラウド生成（Kuromoji形態素解析）
- PDF出力（印刷プレビュー）
- Google Drive画像プロキシ

## 技術スタック

- **バックエンド:** Node.js, Express
- **フロントエンド:** Vanilla JavaScript, TailwindCSS
- **API:** Google Sheets API, Google Drive API, OpenAI API
- **形態素解析:** Kuromoji
- **グラフ:** Google Charts
