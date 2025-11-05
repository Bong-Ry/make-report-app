// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/routes/index.js

const express = require('express');
const reportController = require('../controllers/reportController');
const analysisController = require('../controllers/analysisController');

const router = express.Router();

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

// 4. (画面2) [変更] 転記済みのタブリストと「完了ステータス」を取得
router.post('/api/getTransferredList', reportController.getTransferredList);

// 5. スライド生成 (DLボタン) (変更なし)
router.post('/api/generateSlide', reportController.generateSlide);


// =================================================================
// === ▼▼▼ 分析系APIルート (変更あり) ▼▼▼ ===
// =================================================================

// Kuromoji (ワードクラウド用 - 変更なし)
router.post('/api/analyzeText', analysisController.analyzeText);

// 市区町村レポート (変更なし: 中身がRead-Onlyに変更されたがURLは同じ)
router.post('/api/generateMunicipalityReport', analysisController.generateMunicipalityReport);

// ▼▼▼ [削除] おすすめ理由(N列)分類 (古いAPI) ▼▼▼
// router.post('/api/classifyRecommendations', analysisController.classifyRecommendationOthers);

// ▼▼▼ [新規] おすすめ理由(N列)レポート読み取り (新しいRead-Only API) ▼▼▼
router.post('/api/getRecommendationReport', analysisController.getRecommendationReport);


// AI詳細分析 (実行) (変更なし)
router.post('/api/generateDetailedAnalysis', analysisController.generateDetailedAnalysis);

// AI詳細分析 (読み出し) (変更なし: 中身が単一シート参照に変更されたがURLは同じ)
router.post('/api/getDetailedAnalysis', analysisController.getDetailedAnalysis);

// AI詳細分析 (編集・保存) (変更なし: 中身が単一シート参照に変更されたがURLは同じ)
router.post('/api/updateDetailedAnalysis', analysisController.updateDetailedAnalysis);


module.exports = router;
