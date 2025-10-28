const express = require('express');
const path = require('path');
const routes = require('./routes'); // ./routes/index.js を読み込む
const { initializeKuromoji } = require('./services/kuromoji'); // Kuromoji初期化関数

const app = express();
const PORT = process.env.PORT || 3000;

// ミドルウェア設定
app.use(express.json({ limit: '10mb' })); // JSONリクエストボディの解析 (サイズ上限UP)
app.use(express.static(path.join(__dirname, 'public'))); // 静的ファイル配信 (publicディレクトリ)

// ルーティング設定
app.use('/', routes);

// サーバー起動関数
async function startServer() {
    try {
        await initializeKuromoji(); // Kuromojiの辞書読み込みを待つ
        app.listen(PORT, () => {
            console.log(`Server listening on port ${PORT}`);
        });
    } catch (error) {
        console.error('Failed to initialize server:', error);
        process.exit(1); // 初期化失敗時は終了
    }
}

// サーバー起動
startServer();
