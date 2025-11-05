// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/controllers/analysisController.js

const kuromojiService = require('../services/kuromoji');
// const openaiService = require('../services/openai'); // (不要になったため削除)
// const postalCodeService = require('../services/postalCodeService'); // (不要になったため削除)
const googleSheetsService = require('../services/googleSheets');
const aiAnalysisService = require('../aiAnalysisService'); 
// ▼▼▼ [変更] 必要なヘルパーをインポート ▼▼▼
const { 
    getSystemPromptForDetailAnalysis, 
    getSystemPromptForRecommendationAnalysis,
    getAnalysisTaskTypeForEdit // (AI分析の編集・保存用)
} = require('../utils/helpers');
// ▲▲▲

// --- (変更なし) テキスト分析 (Kuromoji) ---
exports.analyzeText = async (req, res) => {
    console.log('POST /api/analyzeText called');
    const { textList } = req.body;

    if (!textList || !Array.isArray(textList) || textList.length === 0) {
        console.error('[/api/analyzeText] Invalid request: textList is required.', req.body);
        return res.status(400).send('解析対象のテキストリストが必要です。');
    }

    try {
        const analysisResult = kuromojiService.analyzeTextList(textList);
        console.log(`[/api/analyzeText] Kuromoji analysis complete. Found ${analysisResult.results.length} significant words.`);
        res.json(analysisResult);
    } catch (error) {
        console.error("[/api/analyzeText] Error during kuromoji analysis:", error);
        res.status(500).send(`テキスト解析エラー: ${error.message}`);
    }
};

// =================================================================
// === (変更なし) AI詳細分析 (実行・保存) ===
// =================================================================
// (このAPIは、手動で「再実行」ボタンが押された場合にのみ使用される)
exports.generateDetailedAnalysis = async (req, res) => {
    const { centralSheetId, clinicName, columnType } = req.body;
    console.log(`POST /api/generateDetailedAnalysis called for ${clinicName}, type: ${columnType}, SheetID: ${centralSheetId}`);

    if (!centralSheetId || !clinicName || !columnType) {
        console.error('[/api/generateDetailedAnalysis] Invalid request:', req.body);
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType が必要です。');
    }

    try {
        // (aiAnalysisService.js 側で、新しい saveToAnalysisSheet を使うように修正済み)
        const analysisJson = await aiAnalysisService.runAndSaveAnalysis(centralSheetId, clinicName, columnType);
        
        // AIが生成した生のJSONをクライアントに返す
        res.json(analysisJson);

    } catch (error) {
        console.error('[/api/generateDetailedAnalysis] Error generating detailed analysis:', error);
        
        if (error.message && error.message.includes('テキストデータが0件')) {
            return res.status(404).send(error.message);
        }
        
        res.status(500).send(`AI分析中にエラーが発生しました: ${error.message}`);
    }
};

// =================================================================
// === ▼▼▼ [変更] AI詳細分析 (読み出し) (単一シート対応) ▼▼▼ ===
// =================================================================
exports.getDetailedAnalysis = async (req, res) => {
    const { centralSheetId, clinicName, columnType } = req.body;
    console.log(`POST /api/getDetailedAnalysis called for ${clinicName}, type: ${columnType}`);
    
    if (!centralSheetId || !clinicName || !columnType) {
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType が必要です。');
    }
    
    try {
        // ▼▼▼ [変更] 古い getAIAnalysisFromSheet の代わりに、新しい汎用I/O関数を3回呼び出す ▼▼▼
        const [analysisRes, suggestionRes, overallRes] = await Promise.all([
            googleSheetsService.readFromAnalysisSheet(centralSheetId, clinicName, `${columnType}_ANALYSIS`),
            googleSheetsService.readFromAnalysisSheet(centralSheetId, clinicName, `${columnType}_SUGGESTIONS`),
            googleSheetsService.readFromAnalysisSheet(centralSheetId, clinicName, `${columnType}_OVERALL`)
        ]);
        
        const analysisData = {
            analysis: analysisRes || '（データがありません）',
            suggestions: suggestionRes || '（データがありません）',
            overall: overallRes || '（データがありません）'
        };
        // ▲▲▲
        
        res.json(analysisData);
    } catch (err) {
        console.error('[/api/getDetailedAnalysis] Error:', err);
        res.status(500).send(err.message || 'AI分析結果の読み込みに失敗しました。');
    }
};

// =================================================================
// === ▼▼▼ [変更] AI詳細分析 (更新・保存) (単一シート対応) ▼▼▼ ===
// =================================================================
exports.updateDetailedAnalysis = async (req, res) => {
    // ▼▼▼ [変更] sheetName の代わりに、columnType と tabId を受け取る ▼▼▼
    const { centralSheetId, clinicName, columnType, tabId, content } = req.body;
    console.log(`POST /api/updateDetailedAnalysis called for ${clinicName} (${columnType} - ${tabId})`);
    
    if (!centralSheetId || !clinicName || !columnType || !tabId || content == null) {
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType, tabId, content が必要です。');
    }
    
    try {
        // 1. 保存先のタスクタイプ (例: "L_ANALYSIS") をヘルパーから取得
        const taskType = getAnalysisTaskTypeForEdit(columnType, tabId);

        // 2. ▼▼▼ [変更] 新しい汎用I/O関数で保存 ▼▼▼
        await googleSheetsService.saveToAnalysisSheet(
            centralSheetId, 
            clinicName, 
            taskType, 
            content
        );
        
        res.json({ status: 'ok', message: `分析結果 (${taskType}) を更新しました。` });
    } catch (err) {
        console.error('[/api/updateDetailedAnalysis] Error:', err);
        res.status(500).send(err.message || 'AI分析結果の更新に失敗しました。');
    }
};

// =================================================================
// === ▼▼▼ [変更] 市区町村レポート (読み取り専用) ▼▼▼ ===
// =================================================================
exports.generateMunicipalityReport = async (req, res) => {
    const { centralSheetId, clinicName } = req.body;
    console.log(`POST /api/generateMunicipalityReport (Read-Only) called for ${clinicName}`);

    if (!centralSheetId || !clinicName) {
        return res.status(400).send('不正なリクエスト: centralSheetId と clinicName が必要です。');
    }

    try {
        // ▼▼▼ [変更] 分析ロジックを削除し、読み取り専用にする ▼▼▼
        // (分析・生成は aiAnalysisService.runAndSaveMunicipalityAnalysis がバックグラウンドで行う)
        
        console.log(`[/api/generateMunicipalityReport] Reading 'MUNICIPALITY_TABLE' from sheet...`);
        const tableData = await googleSheetsService.readFromAnalysisSheet(
            centralSheetId, 
            clinicName, 
            'MUNICIPALITY_TABLE'
        );
        
        if (!tableData || tableData.length < 2) { // (ヘッダー + データ)
            console.log(`[/api/generateMunicipalityReport] No data found in 'MUNICIPALITY_TABLE'.`);
            return res.json([]); // データが存在しない
        }

        // 2. ヘッダー行を捨てる
        tableData.shift(); 
        
        // 3. フロントエンドが期待する形式 (割合を 12.3 形式) に変換
        const frontendTable = tableData.map(row => ({
            prefecture: row[0] || '',
            municipality: row[1] || '',
            count: parseFloat(row[2]) || 0,
            percentage: (parseFloat(row[3]) || 0) * 100 // (シートには 0.123 形式で保存されている)
        }));
        
        res.json(frontendTable);

    } catch (error) {
        console.error('[/api/generateMunicipalityReport] Error reading municipality report:', error);
        res.status(500).send(`市区町村レポートの読み込み中にエラーが発生しました: ${error.message}`);
    }
};


// =================================================================
// === ▼▼▼ [変更] おすすめ理由 (削除) & (読み取り専用APIを新設) ▼▼▼ ===
// =================================================================

/**
 * [削除] /api/classifyRecommendations
 * (このAPIは不要になったため、routes/index.js からも削除する)
 */
// exports.classifyRecommendationOthers = ... (削除)


/**
 * [新規] /api/getRecommendationReport (読み取り専用)
 * (バックグラウンドで集計済みの「おすすめ理由」テーブルを取得する)
 */
exports.getRecommendationReport = async (req, res) => {
    const { centralSheetId, clinicName } = req.body;
    console.log(`POST /api/getRecommendationReport (Read-Only) called for ${clinicName}`);

    if (!centralSheetId || !clinicName) {
        return res.status(400).send('不正なリクエスト: centralSheetId と clinicName が必要です。');
    }

    try {
        // 1. シートから 'RECOMMENDATION_TABLE' を読み込む
        console.log(`[/api/getRecommendationReport] Reading 'RECOMMENDATION_TABLE' from sheet...`);
        const tableData = await googleSheetsService.readFromAnalysisSheet(
            centralSheetId, 
            clinicName, 
            'RECOMMENDATION_TABLE'
        );
        
        if (!tableData || tableData.length < 2) { // (ヘッダー + データ)
            console.log(`[/api/getRecommendationReport] No data found in 'RECOMMENDATION_TABLE'.`);
             // (フロントエンドが [[カテゴリ, 件数], ...] の形式を期待しているため、空のヘッダーを返す)
            return res.json([['カテゴリ', '件数']]);
        }
        
        // 2. ヘッダー行を捨てる
        tableData.shift(); 
        
        // 3. フロントエンドの円グラフが期待する形式 [ [項目, 件数], ... ] に変換
        // (ご要望: A項目, B件数, C割合 で保存されている)
        const chartData = [['カテゴリ', '件数']]; // ヘッダー
        
        tableData.forEach(row => {
            const item = row[0] || '不明'; // A列: 項目
            const count = parseFloat(row[1]) || 0; // B列: 件数
            // (C列の割合はここでは不要)
            
            if (count > 0) {
                chartData.push([item, count]);
            }
        });
        
        res.json(chartData);

    } catch (error) {
        console.error('[/api/getRecommendationReport] Error reading recommendation report:', error);
        res.status(500).send(`おすすめ理由レポートの読み込み中にエラーが発生しました: ${error.message}`);
    }
};
