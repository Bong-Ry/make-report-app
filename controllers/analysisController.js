const kuromojiService = require('../services/kuromoji');
const openaiService = require('../services/openai');
const postalCodeService = require('../services/postalCodeService');
const googleSheetsService = require('../services/googleSheets');
// ▼▼▼ [新規] AI分析サービスを読み込む ▼▼▼
const aiAnalysisService = require('../services/aiAnalysisService'); 
const { getSystemPromptForDetailAnalysis, getSystemPromptForRecommendationAnalysis } = require('../utils/helpers');

// メモリキャッシュ (簡易版)
// ▼▼▼ [変更点] 市区町村レポートのキャッシュはGoogle Sheetが担うため削除 ▼▼▼
// let municipalityReportCache = {}; 
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
        // ★ 要求 #4 (ヘッダーなし転記) に対応した getReportDataForCharts が返す
        //    rawText を使うため、この関数自体は変更不要
        const analysisResult = kuromojiService.analyzeTextList(textList);
        console.log(`[/api/analyzeText] Kuromoji analysis complete. Found ${analysisResult.results.length} significant words.`);
        res.json(analysisResult);
    } catch (error) {
        console.error("[/api/analyzeText] Error during kuromoji analysis:", error);
        res.status(500).send(`テキスト解析エラー: ${error.message}`);
    }
};

// =================================================================
// === ▼▼▼ [変更] AI詳細分析 (実行・保存) ▼▼▼ ===
// =================================================================
exports.generateDetailedAnalysis = async (req, res) => {
    const { centralSheetId, clinicName, columnType } = req.body;
    console.log(`POST /api/generateDetailedAnalysis called for ${clinicName}, type: ${columnType}, SheetID: ${centralSheetId}`);

    if (!centralSheetId || !clinicName || !columnType) {
        console.error('[/api/generateDetailedAnalysis] Invalid request:', req.body);
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType が必要です。');
    }

    try {
        // ▼▼▼ [変更点] AI分析の実行と保存ロジックをサービスに移譲 ▼▼▼
        const analysisJson = await aiAnalysisService.runAndSaveAnalysis(centralSheetId, clinicName, columnType);
        
        // 5. AIが生成した生のJSONをクライアントに返す
        res.json(analysisJson);

    } catch (error) {
        console.error('[/api/generateDetailedAnalysis] Error generating detailed analysis:', error);
        
        // ▼▼▼ [変更点] サービスからの 404 エラーをハンドル ▼▼▼
        if (error.message && error.message.includes('テキストデータが0件')) {
            return res.status(404).send(error.message);
        }
        
        res.status(500).send(`AI分析中にエラーが発生しました: ${error.message}`);
    }
};

// =================================================================
// === (変更なし) AI詳細分析 (読み出し) ===
// =================================================================
exports.getDetailedAnalysis = async (req, res) => {
    const { centralSheetId, clinicName, columnType } = req.body;
    console.log(`POST /api/getDetailedAnalysis called for ${clinicName}, type: ${columnType}`);
    
    if (!centralSheetId || !clinicName || !columnType) {
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType が必要です。');
    }
    
    try {
        const analysisData = await googleSheetsService.getAIAnalysisFromSheet(centralSheetId, clinicName, columnType);
        res.json(analysisData);
    } catch (err) {
        console.error('[/api/getDetailedAnalysis] Error:', err);
        res.status(500).send(err.message || 'AI分析結果の読み込みに失敗しました。');
    }
};

// =================================================================
// === (変更なし) AI詳細分析 (更新・保存) ===
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
// === ▼▼▼ 更新API (要求 #3 市区町村レポート) ▼▼▼ ===
// =================================================================
exports.generateMunicipalityReport = async (req, res) => {
    const { centralSheetId, clinicName } = req.body;
    console.log(`POST /api/generateMunicipalityReport called for ${clinicName}, SheetID: ${centralSheetId}`);

    if (!centralSheetId || !clinicName) {
        return res.status(400).send('不正なリクエスト: centralSheetId と clinicName が必要です。');
    }

    // ▼▼▼ [変更点] 保存するシート名を定義 ▼▼▼
    const municipalitySheetName = `${clinicName}-地域`;

    try {
        // 1. ▼▼▼ [変更点] まず、保存済みシートの読み取りを試みる ▼▼▼
        console.log(`[/api/generateMunicipalityReport] Trying to read from sheet: "${municipalitySheetName}"`);
        const existingData = await googleSheetsService.readMunicipalityData(centralSheetId, municipalitySheetName);
        
        if (existingData) {
            console.log(`[/api/generateMunicipalityReport] Found existing data in sheet. Returning.`);
            return res.json(existingData); // 既存データを返す
        }

        // 2. ▼▼▼ 既存データがない場合のみ、集計処理を実行 ▼▼▼
        console.log(`[/api/generateMunicipalityReport] No existing sheet. Generating new report...`);
        console.log(`[/api/generateMunicipalityReport] Fetching postal code data...`);
        
        // ★ 要求 #4 (ヘッダーなし転記) に対応した getReportDataForCharts を呼び出す
        const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
        
        const postalCodeCounts = reportData.postalCodeData.counts;
        if (!postalCodeCounts || Object.keys(postalCodeCounts).length === 0) {
            console.log(`[/api/generateMunicipalityReport] No postal code data found in sheet.`);
            return res.json([]);
        }

        // 3. 郵便番号APIで住所を検索 (変更なし)
        const uniquePostalCodes = Object.keys(postalCodeCounts);
        console.log(`[/api/generateMunicipalityReport] Looking up ${uniquePostalCodes.length} unique postal codes...`);
        
        const addressAggregates = {}; // { "都道府県-市区町村": count }
        let totalValidCodesCount = 0;

        const lookupPromises = uniquePostalCodes.map(async (postalCode) => {
            const count = postalCodeCounts[postalCode];
            // ★ 要求 #1 (市区町村の整形) に対応した postalCodeService を呼び出す
            const address = await postalCodeService.lookupPostalCode(postalCode); 
            
            if (address && address.prefecture && address.municipality) {
                const key = `${address.prefecture}-${address.municipality}`;
                addressAggregates[key] = (addressAggregates[key] || 0) + count;
            } else {
                addressAggregates['不明-不明'] = (addressAggregates['不明-不明'] || 0) + count;
            }
            totalValidCodesCount += count;
        });

        // 4. 集計 (変更なし)
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
                percentage: (count / totalValidCodesCount) // 割合 (0.123 形式)
            };
        });
        
        finalTable.sort((a, b) => b.count - a.count);
        
        // 5. ▼▼▼ [変更点] 結果をスプレッドシートに保存 ▼▼▼
        console.log(`[/api/generateMunicipalityReport] Saving ${finalTable.length} rows to sheet: "${municipalitySheetName}"`);
        await googleSheetsService.saveMunicipalityData(centralSheetId, municipalitySheetName, finalTable);

        // 6. ▼▼▼ [変更点] フロントエンドには割合を % に変換して返す ▼▼▼
        const frontendTable = finalTable.map(row => ({
            ...row,
            percentage: row.percentage * 100 // 0.123 -> 12.3
        }));
        
        console.log(`[/api/generateMunicipalityReport] Generated and saved report.`);
        res.json(frontendTable);

    } catch (error) {
        console.error('[/api/generateMunicipalityReport] Error during municipality report generation:', error);
        res.status(500).send(`市区町村レポートの生成中にエラーが発生しました: ${error.message}`);
    }
};


// =================================================================
// === (変更なし) おすすめ理由の「その他」分類 ===
// =================================================================
exports.classifyRecommendationOthers = async (req, res) => {
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
