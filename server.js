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
        console.log('=== Server Startup ===');
        console.log('Node version:', process.version);
        console.log('Working directory:', process.cwd());
        console.log('__dirname:', __dirname);
        console.log('PORT:', PORT);

        console.log('Initializing Kuromoji...');
        await initializeKuromoji(); // Kuromojiの辞書読み込みを待つ
        console.log('Kuromoji initialized successfully.');

        app.listen(PORT, '0.0.0.0', () => {
            console.log(`✓ Server listening on 0.0.0.0:${PORT}`);
            console.log('=== Server Ready ===');
        });
    } catch (error) {
        console.error('✗ Failed to initialize server:', error);
        console.error('Error stack:', error.stack);
        process.exit(1); // 初期化失敗時は終了
    }
}

// サーバー起動
startServer();
