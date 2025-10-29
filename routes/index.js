const express = require('express');
const reportController = require('../controllers/reportController');
const analysisController = require('../controllers/analysisController');

const router = express.Router();

// ルートパス (index.html表示用 - server.jsのstatic配信で処理されるため省略可だが明示)
// router.get('/', (req, res) => {
//     res.sendFile(path.join(__dirname, '../public', 'index.html'));
// });

// API Routes
router.get('/api/getClinicList', reportController.getClinicList);
router.post('/api/getReportData', reportController.getReportData);
router.post('/api/analyzeText', analysisController.analyzeText);
router.post('/api/generateDetailedAnalysis', analysisController.generateDetailedAnalysis); // 詳細分析ルート
router.post('/api/generateMunicipalityReport', analysisController.generateMunicipalityReport); // 市区町村集計ルート
router.post('/generate-pdf', reportController.generatePdf); // PDF生成ルート

// =================================================================
// === ▼▼▼ 新しいAPIルートを追加 ▼▼▼ ===
// =================================================================
router.post('/api/classifyRecommendations', analysisController.classifyRecommendationOthers); // おすすめ理由(N列)分類用ルート
// =================================================================
// === ▲▲▲ 新しいAPIルートを追加 ▲▲▲ ===
// =================================================================

module.exports = router;
