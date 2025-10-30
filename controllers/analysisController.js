const kuromojiService = require('../services/kuromoji');
const openaiService = require('../services/openai');
const postalCodeService = require('../services/postalCodeService');
// ▼▼▼ googleSheetsService を追加 ▼▼▼
const googleSheetsService = require('../services/googleSheets');
const { getSystemPromptForDetailAnalysis, getSystemPromptForRecommendationAnalysis } = require('../utils/helpers');

// メモリキャッシュ (簡易版)
// ▼▼▼ detailedAnalysisResultsCache はGoogle Sheetが担うため削除 ▼▼▼
// let detailedAnalysisResultsCache = {}; 
let municipalityReportCache = {}; // 市区町村レポート用キャッシュ
let recommendationAnalysisCache = {}; // おすすめ理由 分類用キャッシュ

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
// === ▼▼▼ 更新API ▼▼▼ ===
// AI詳細分析の「実行」と「保存」
// =================================================================
exports.generateDetailedAnalysis = async (req, res) => {
    // ▼▼▼ textList の代わりに centralSheetId を受け取る ▼▼▼
    const { centralSheetId, clinicName, columnType } = req.body;
    console.log(`POST /api/generateDetailedAnalysis called for ${clinicName}, type: ${columnType}, SheetID: ${centralSheetId}`);

    if (!centralSheetId || !clinicName || !columnType) {
        console.error('[/api/generateDetailedAnalysis] Invalid request:', req.body);
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType が必要です。');
    }

    // ▼▼▼ キャッシュロジックは削除 ▼▼▼
    // (Google Sheetが永続キャッシュとなるため)

    // AIへの指示（プロンプト）を生成 (変更なし)
    const systemPrompt = getSystemPromptForDetailAnalysis(clinicName, columnType);
    if (!systemPrompt) {
        console.error(`[/api/generateDetailedAnalysis] Invalid analysis type: ${columnType}`);
        return res.status(400).send('無効な分析タイプです。');
    }

    try {
        // 1. ▼▼▼ 集計スプシから分析対象のテキストリストを取得 ▼▼▼
        console.log(`[/api/generateDetailedAnalysis] Fetching chart data to get textList...`);
        const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
        
        let textList = [];
        switch (columnType) {
            case 'L': textList = reportData.npsData.rawText || []; break;
            case 'I_good': textList = reportData.feedbackData.i_column.results || []; break;
            case 'I_bad': textList = reportData.feedbackData.i_column.results || []; break; // I_good と同じソース
            case 'J': textList = reportData.feedbackData.j_column.results || []; break;
            case 'M': textList = reportData.feedbackData.m_column.results || []; break;
            default: throw new Error(`無効な分析タイプです: ${columnType}`);
        }
        
        if (textList.length === 0) {
            console.log(`[/api/generateDetailedAnalysis] No text data (0 items) found for ${columnType}.`);
            return res.status(404).send('分析対象のテキストデータが0件です。');
        }

        // 2. 入力テキストを結合・制限 (変更なし)
        const truncatedList = textList.length > 100 ? textList.slice(0, 100) : textList;
        const combinedText = truncatedList.join('\n\n---\n\n');
        const inputText = combinedText.substring(0, 15000);

        console.log(`[/api/generateDetailedAnalysis] Sending ${truncatedList.length} comments (input text length: ${inputText.length}) to OpenAI...`);

        // 3. OpenAI API 呼び出し (変更なし)
        const analysisJson = await openaiService.generateJsonAnalysis(systemPrompt, inputText);
        
        // 4. ▼▼▼ 結果をGoogleスプレッドシートに保存 ▼▼▼
        console.log(`[/api/generateDetailedAnalysis] Saving analysis results to Google Sheet...`);
        await googleSheetsService.saveAIAnalysisToSheet(centralSheetId, clinicName, columnType, analysisJson);

        // 5. ▼▼▼ AIが生成した生のJSONをクライアントに返す (初回表示用) ▼▼▼
        res.json(analysisJson);

    } catch (error) {
        console.error('[/api/generateDetailedAnalysis] Error generating detailed analysis:', error);
        res.status(500).send(`AI分析中にエラーが発生しました: ${error.message}`);
    }
};

// =================================================================
// === ▼▼▼ 新規API (1/2) ▼▼▼ ===
// AI詳細分析（保存済み）の「読み出し」
// =================================================================
exports.getDetailedAnalysis = async (req, res) => {
    const { centralSheetId, clinicName, columnType } = req.body;
    console.log(`POST /api/getDetailedAnalysis called for ${clinicName}, type: ${columnType}`);
    
    if (!centralSheetId || !clinicName || !columnType) {
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType が必要です。');
    }
    
    try {
        const analysisData = await googleSheetsService.getAIAnalysisFromSheet(centralSheetId, clinicName, columnType);
        
        // analysisData は { analysis: "...", suggestions: "...", overall: "..." } の形式
        res.json(analysisData);
        
    } catch (err) {
        console.error('[/api/getDetailedAnalysis] Error:', err);
        res.status(500).send(err.message || 'AI分析結果の読み込みに失敗しました。');
    }
};

// =================================================================
// === ▼▼▼ 新規API (2/2) ▼▼▼ ===
// AI詳細分析（編集後）の「更新（保存）」
// =================================================================
exports.updateDetailedAnalysis = async (req, res) => {
    const { centralSheetId, sheetName, content } = req.body;
    console.log(`POST /api/updateDetailedAnalysis called for sheet: ${sheetName}`);
    
    if (!centralSheetId || !sheetName || content == null) {
        return res.status(400).send('不正なリクエスト: centralSheetId, sheetName, content が必要です。');
    }
    
    try {
        await googleSheetsService.updateAIAnalysisInSheet(centralSheetId, sheetName, content);
        res.json({ status: 'ok', message: `シート ${sheetName} を更新しました。` });
    } catch (err) {
        console.error('[/api/updateDetailedAnalysis] Error:', err);
        res.status(500).send(err.message || 'AI分析結果の更新に失敗しました。');
    }
};

// =================================================================
// === ▼▼▼ 更新API ▼▼▼ ===
// 市区町村レポート生成
// =================================================================
exports.generateMunicipalityReport = async (req, res) => {
    // ▼▼▼ postalCodeCounts の代わりに centralSheetId と clinicName を受け取る ▼▼▼
    const { centralSheetId, clinicName } = req.body;
    console.log(`POST /api/generateMunicipalityReport called for ${clinicName}, SheetID: ${centralSheetId}`);

    if (!centralSheetId || !clinicName) {
        return res.status(400).send('不正なリクエスト: centralSheetId と clinicName が必要です。');
    }

    // ▼▼▼ キャッシュキーを変更 ▼▼▼
    const cacheKey = `${centralSheetId}-${clinicName}`;
    if (municipalityReportCache[cacheKey]) {
        console.log(`[/api/generateMunicipalityReport] Returning cached municipality report for ${cacheKey}`);
        return res.json(municipalityReportCache[cacheKey]);
    }
    
    const addressAggregates = {}; // { "都道府県-市区町村": count }
    let totalValidCodesCount = 0;

    try {
        // 1. ▼▼▼ 集計スプシから郵便番号データを取得 ▼▼▼
        console.log(`[/api/generateMunicipalityReport] Fetching postal code data...`);
        const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
        
        const postalCodeCounts = reportData.postalCodeData.counts;
        if (!postalCodeCounts || Object.keys(postalCodeCounts).length === 0) {
            console.log(`[/api/generateMunicipalityReport] No postal code data found in sheet.`);
            return res.json([]);
        }

        // 2. 郵便番号APIで住所を検索 (変更なし)
        const uniquePostalCodes = Object.keys(postalCodeCounts);
        console.log(`[/api/generateMunicipalityReport] Looking up ${uniquePostalCodes.length} unique postal codes...`);

        const lookupPromises = uniquePostalCodes.map(async (postalCode) => {
            const count = postalCodeCounts[postalCode];
            const address = await postalCodeService.lookupPostalCode(postalCode); // 既存サービス呼び出し
            
            if (address && address.prefecture && address.municipality) {
                const key = `${address.prefecture}-${address.municipality}`;
                addressAggregates[key] = (addressAggregates[key] || 0) + count;
            } else {
                addressAggregates['不明-不明'] = (addressAggregates['不明-不明'] || 0) + count;
            }
            totalValidCodesCount += count;
        });

        // 3. 集計 (変更なし)
        await Promise.all(lookupPromises);
        
        if (totalValidCodesCount === 0) {
            return res.json([]);
        }

        const finalTable = Object.entries(addressAggregates).map(([key, count]) => {
            const [prefecture, municipality] = key.split('-');
            return {
                prefecture: prefecture,
                municipality: municipality,
                count: count,
                percentage: (count / totalValidCodesCount) * 100
            };
        });
        
        finalTable.sort((a, b) => b.count - a.count);
        
        // サーバーキャッシュに保存
        municipalityReportCache[cacheKey] = finalTable;
        console.log(`[/api/generateMunicipalityReport] Generated and cached report for ${cacheKey}`);
        
        res.json(finalTable);

    } catch (error) {
        console.error('[/api/generateMunicipalityReport] Error during municipality report generation:', error);
        res.status(500).send(`市区町村レポートの生成中にエラーが発生しました: ${error.message}`);
    }
};


// =================================================================
// === (変更なし) おすすめ理由の「その他」分類 ===
// =================================================================
/**
 * N列（おすすめ理由）の「その他」項目をAIで分類するAPI
 */
exports.classifyRecommendationOthers = async (req, res) => {
    // (このAPIはクライアントから textList を受け取るため、変更なし)
    
    const { clinicName, otherList, fixedKeys } = req.body;
    console.log(`POST /api/classifyRecommendations called for ${clinicName}`);

    if (!clinicName || !otherList || !Array.isArray(otherList) || !fixedKeys || !Array.isArray(fixedKeys)) {
        return res.status(400).send('不正なリクエスト: clinicName, otherList, fixedKeys が必要です。');
    }
    
    if (otherList.length === 0) {
        console.log('[/api/classifyRecommendations] otherList is empty, returning empty result.');
        return res.json({ classifiedResults: [] });
    }

    const textHash = otherList.join('|').substring(0, 50);
    const cacheKey = `${clinicName}-rec-${textHash}`;

    if (recommendationAnalysisCache[cacheKey]) {
        console.log(`[/api/classifyRecommendations] Returning cached classification for ${cacheKey}`);
        return res.json(recommendationAnalysisCache[cacheKey]);
    }

    const systemPrompt = getSystemPromptForRecommendationAnalysis(fixedKeys);
    const inputText = otherList.join('\n');
    
    console.log(`[/api/classifyRecommendations] Sending ${otherList.length} "other" texts (length: ${inputText.length}) to OpenAI for classification...`);

    try {
        const analysisJson = await openaiService.generateJsonAnalysis(systemPrompt, inputText);
        
        if (!analysisJson || !Array.isArray(analysisJson.classifiedResults)) {
            console.error('[/api/classifyRecommendations] AI returned invalid JSON structure:', analysisJson);
            throw new Error('AIが予期しない分類結果フォーマットを返しました。');
        }

        recommendationAnalysisCache[cacheKey] = analysisJson;
        console.log(`[/api/classifyRecommendations] Cached classification result for ${cacheKey}`);

        res.json(analysisJson);

    } catch (error) {
        console.error('[/api/classifyRecommendations] Error classifying recommendations:', error);
        res.status(500).send(`AI分類中にエラーが発生しました: ${error.message}`);
    }
};
