const express = require('express');
const reportController = require('../controllers/reportController');
const analysisController = require('../controllers/analysisController');

const router = express.Router();

// ルートパス (index.html表示用 - server.jsのstatic配信で処理されるため省略可だが明示)
// router.get('/', (req, res) => {
//     res.sendFile(path.join(__dirname, '../public', 'index.html'));
// });

// --- API Routes (レポート・データ関連) ---
router.get('/api/getClinicList', reportController.getClinicList);
router.post('/generate-pdf', reportController.generatePdf); // PDF生成ルート

// =================================================================
// === ▼▼▼ 新アーキテクチャ用ルート ▼▼▼ ===
// =================================================================

// 1. (画面1) 集計期間を送信し、集計スプレッドシートIDを取得（または作成）
router.post('/api/findOrCreateSheet', reportController.findOrCreateSheet);

// 2. (画面2) 「レポート発行」で、元データ -> 集計スプシ へのデータ転記(ETL)を実行
router.post('/api/getReportData', reportController.getReportData);

// 3. (画面3以降) グラフ表示用に、集計スプシから集計済みデータを取得
router.post('/api/getChartData', reportController.getChartData);

// 4. (画面2) [新規] 転記済みのタブリストを取得
router.post('/api/getTransferredList', reportController.getTransferredList);


// =================================================================
// === ▼▼▼ 分析系APIルート ▼▼▼ ===
// =================================================================

// Kuromoji (ワードクラウド用 - 変更なし)
router.post('/api/analyzeText', analysisController.analyzeText);

// 市区町村レポート (集計スプシ参照に変更 - URLは変更なし)
router.post('/api/generateMunicipalityReport', analysisController.generateMunicipalityReport);

// おすすめ理由(N列)分類 (変更なし)
router.post('/api/classifyRecommendations', analysisController.classifyRecommendationOthers);

// AI詳細分析 (実行) (集計スプシ参照に変更 - URLは変更なし)
router.post('/api/generateDetailedAnalysis', analysisController.generateDetailedAnalysis);

// AI詳細分析 (読み出し) (新規)
router.post('/api/getDetailedAnalysis', analysisController.getDetailedAnalysis);

// AI詳細分析 (編集・保存) (新規)
router.post('/api/updateDetailedAnalysis', analysisController.updateDetailedAnalysis);


module.exports = router;
