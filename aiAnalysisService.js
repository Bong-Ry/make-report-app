// services/aiAnalysisService.js (新規作成)

const openaiService = require('./openai');
const googleSheetsService = require('./googleSheets');
const { getSystemPromptForDetailAnalysis } = require('../utils/helpers');

/**
 * [新規] AI分析を実行し、結果をシートに保存する共通ロジック
 * (analysisController.js の generateDetailedAnalysis の中身をリファクタリング)
 * @param {string} centralSheetId 
 * @param {string} clinicName 
 * @param {string} columnType (例: 'L', 'I_good', 'J'...)
 * @returns {object} AIが生成した生のJSON
 * @throws {Error} テキストデータが0件の場合、またはAI分析・保存に失敗した場合
 */
exports.runAndSaveAnalysis = async (centralSheetId, clinicName, columnType) => {
    console.log(`[aiAnalysisService] Running for ${clinicName}, type: ${columnType}`);
    
    const systemPrompt = getSystemPromptForDetailAnalysis(clinicName, columnType);
    if (!systemPrompt) {
        throw new Error(`[aiAnalysisService] Invalid analysis type (no prompt): ${columnType}`);
    }

    // 1. 集計スプシから分析対象のテキストリストを取得
    // ★ 要求 #4 (ヘッダーなし転記) に対応した getReportDataForCharts を呼び出す
    const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    
    let textList = [];
    switch (columnType) {
        case 'L': textList = reportData.npsData.rawText || []; break;
        case 'I_good': textList = reportData.feedbackData.i_column.results || []; break;
        // ▼▼▼ [修正] I_bad は I_good と同じ元データ (I列) を参照する ▼▼▼
        case 'I_bad': textList = reportData.feedbackData.i_column.results || []; break; 
        case 'J': textList = reportData.feedbackData.j_column.results || []; break;
        case 'M': textList = reportData.feedbackData.m_column.results || []; break;
        default: throw new Error(`[aiAnalysisService] 無効な分析タイプです: ${columnType}`);
    }
    
    if (textList.length === 0) {
        console.log(`[aiAnalysisService] No text data (0 items) found for ${columnType}.`);
        // 404エラーの代わりに例外をスローし、呼び出し元で処理できるようにする
        throw new Error(`分析対象のテキストデータが0件です。 (Type: ${columnType})`);
    }

    // 2. 入力テキストを結合・制限 (変更なし)
    const truncatedList = textList.length > 100 ? textList.slice(0, 100) : textList;
    const combinedText = truncatedList.join('\n\n---\n\n');
    const inputText = combinedText.substring(0, 15000);

    console.log(`[aiAnalysisService] Sending ${truncatedList.length} comments (input text length: ${inputText.length}) to OpenAI...`);

    // 3. OpenAI API 呼び出し (変更なし)
    const analysisJson = await openaiService.generateJsonAnalysis(systemPrompt, inputText);
    
    // 4. 結果をGoogleスプレッドシートに保存 (変更なし)
    console.log(`[aiAnalysisService] Saving analysis results to Google Sheet...`);
    await googleSheetsService.saveAIAnalysisToSheet(centralSheetId, clinicName, columnType, analysisJson);

    // 5. AIが生成した生のJSONを返す
    return analysisJson;
};
