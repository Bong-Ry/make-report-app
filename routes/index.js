// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/routes/index.js

const express = require('express');
const { google } = require('googleapis');
const reportController = require('../controllers/reportController');
const analysisController = require('../controllers/analysisController');

const router = express.Router();

// Google Drive API初期化（画像プロキシ用）
const KEYFILEPATH = '/etc/secrets/credentials.json';
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',  // すべてのDriveファイルへの読み取りアクセス
];
let drive;

try {
  const auth = new google.auth.GoogleAuth({ keyFile: KEYFILEPATH, scopes: SCOPES });
  drive = google.drive({ version: 'v3', auth });
  console.log('Google Drive API client initialized for image proxy');
} catch (err) {
  console.error('Failed to initialize Google Drive API for images:', err);
}

// --- API Routes (レポート・データ関連) ---
router.get('/api/getClinicList', reportController.getClinicList);
// router.post('/generate-pdf', reportController.generatePdf); // [修正] PDF生成ルートを削除

// =================================================================
// === ▼▼▼ 新アーキテクチャ用ルート (変更なし) ▼▼▼ ===
// =================================================================

// 1. (画面1) 集計期間を送信し、集計スプレッドシートIDを取得（または作成）
router.post('/api/findOrCreateSheet', reportController.findOrCreateSheet);

// 2. (画面2) 「レポート発行」で、元データ -> 集計スプシ へのデータ転記(ETL)を実行
router.post('/api/getReportData', reportController.getReportData);

// 3. (画面3以降) グラフ表示用に、集計スプシから集計済みデータを取得
router.post('/api/getChartData', reportController.getChartData);

// 4. (画面2) [変更] 転記済みのタブリストと「完了ステータス」を取得
router.post('/api/getTransferredList', reportController.getTransferredList);

// =================================================================
// === ▼▼▼ 分析系APIルート (変更なし) ▼▼▼ ===
// =================================================================

// Kuromoji (ワードクラウド用 - 変更なし)
router.post('/api/analyzeText', analysisController.analyzeText);

// 市区町村レポート (変更なし: 中身がRead-Onlyに変更されたがURLは同じ)
router.post('/api/generateMunicipalityReport', analysisController.generateMunicipalityReport);

// ▼▼▼ [新規] おすすめ理由(N列)レポート読み取り (新しいRead-Only API) ▼▼▼
router.post('/api/getRecommendationReport', analysisController.getRecommendationReport);


// AI詳細分析 (実行) (変更なし)
router.post('/api/generateDetailedAnalysis', analysisController.generateDetailedAnalysis);

// AI詳細分析 (読み出し) (変更なし: 中身が単一シート参照に変更されたがURLは同じ)
router.post('/api/getDetailedAnalysis', analysisController.getDetailedAnalysis);

// AI詳細分析 (編集・保存) (変更なし: 中身が単一シート参照に変更されたがURLは同じ)
router.post('/api/updateDetailedAnalysis', analysisController.updateDetailedAnalysis);

// ▼▼▼ [ここから変更] ▼▼▼
// (タブ切り替え時の単一セル取得API)
router.post('/api/getSingleAnalysisCell', analysisController.getSingleAnalysisCell);
// ▲▲▲ [変更ここまで] ▲▲▲

// =================================================================
// === ▼▼▼ [変更なし] コメント編集用APIルート ▼▼▼ ===
// =================================================================

// [変更なし] コメントシートからデータを読み込む (コントローラー側で `sheetName` を読むよう変更)
router.post('/api/getCommentData', analysisController.getCommentData);

// [変更なし] コメントシートのセルを更新する (コントローラー側で `sheetName` と `cell` を読むよう変更)
router.post('/api/updateCommentData', analysisController.updateCommentData);

// =================================================================
// === ▼▼▼ 画像プロキシエンドポイント ▼▼▼ ===
// =================================================================

// Google Drive画像をプロキシして配信
router.get('/image/:id', async (req, res) => {
  try {
    const imageId = req.params.id;

    if (!drive) {
      return res.status(500).send('Google Drive API not initialized');
    }

    // Google Drive APIでファイルのメタデータを取得
    const fileMetadata = await drive.files.get({
      fileId: imageId,
      fields: 'mimeType'
    });

    // ファイルをストリームとして取得
    const fileStream = await drive.files.get(
      { fileId: imageId, alt: 'media' },
      { responseType: 'stream' }
    );

    // Content-Typeヘッダーを設定
    res.set('Content-Type', fileMetadata.data.mimeType || 'image/png');
    res.set('Cache-Control', 'public, max-age=86400'); // 24時間キャッシュ

    // ストリームをレスポンスにパイプ
    fileStream.data.pipe(res);

  } catch (error) {
    console.error('Image proxy error:', error);
    if (error.code === 404) {
      res.status(404).send('Image not found');
    } else if (error.code === 403) {
      res.status(403).send('Access denied - check file permissions');
    } else {
      res.status(500).send('Failed to fetch image');
    }
  }
});

module.exports = router;
