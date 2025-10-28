const kuromojiService = require('../services/kuromoji');
const openaiService = require('../services/openai');
const { getSystemPromptForDetailAnalysis } = require('../utils/helpers');

// メモリキャッシュ (簡易版)
let detailedAnalysisResultsCache = {};

// --- テキスト分析 (Kuromoji) ---
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

// --- 詳細テキスト分析 (OpenAI) ---
exports.generateDetailedAnalysis = async (req, res) => {
    const { clinicName, columnType, textList } = req.body;
    console.log(`POST /api/generateDetailedAnalysis called for ${clinicName}, type: ${columnType}`);
    console.log(`[/api/generateDetailedAnalysis] Received ${textList ? textList.length : 0} texts.`);

    if (!clinicName || !columnType || !textList || !Array.isArray(textList) || textList.length === 0) {
        console.error('[/api/generateDetailedAnalysis] Invalid request:', req.body);
        return res.status(400).send('不正なリクエスト: クリニック名、列タイプ、テキストリストが必要です。');
    }

    const cacheKey = `${clinicName}-${columnType}`;

    // キャッシュ確認
    if (detailedAnalysisResultsCache[cacheKey]) {
        console.log(`[/api/generateDetailedAnalysis] Returning cached detailed analysis for ${cacheKey}`);
        return res.json(detailedAnalysisResultsCache[cacheKey]);
    }

    // AIへの指示（プロンプト）を生成
    const systemPrompt = getSystemPromptForDetailAnalysis(clinicName, columnType);
    if (!systemPrompt) {
        console.error(`[/api/generateDetailedAnalysis] Invalid analysis type: ${columnType}`);
        return res.status(400).send('無効な分析タイプです。');
    }

    // 入力テキストを結合・制限
    const truncatedList = textList.length > 100 ? textList.slice(0, 100) : textList;
    const combinedText = truncatedList.join('\n\n---\n\n');
    const inputText = combinedText.substring(0, 15000);

    console.log(`[/api/generateDetailedAnalysis] Sending ${truncatedList.length} comments (input text length: ${inputText.length}) to OpenAI model 'gpt-4o-mini' for detailed analysis...`);

    try {
        const analysisJson = await openaiService.generateJsonAnalysis(systemPrompt, inputText);

        // 結果をメモリにキャッシュ
        detailedAnalysisResultsCache[cacheKey] = analysisJson;
        console.log(`[/api/generateDetailedAnalysis] Cached detailed analysis result for ${cacheKey}`);

        res.json(analysisJson);

    } catch (error) {
        console.error('[/api/generateDetailedAnalysis] Error generating detailed analysis:', error);
        res.status(500).send(`AI分析中にエラーが発生しました: ${error.message}`);
    }
};
