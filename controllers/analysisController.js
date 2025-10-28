const kuromojiService = require('../services/kuromoji');
const openaiService = require('../services/openai');
const postalCodeService = require('../services/postalCodeService'); // ★ 郵便番号サービス
const { getSystemPromptForDetailAnalysis } = require('../utils/helpers');

// メモリキャッシュ (簡易版)
let detailedAnalysisResultsCache = {};
let municipalityReportCache = {}; // ★ 市区町村レポート用キャッシュ

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

// --- ★ 市区町村レポート生成 (新規追加) ---
exports.generateMunicipalityReport = async (req, res) => {
    const { postalCodeCounts, clinicName } = req.body;
    console.log(`POST /api/generateMunicipalityReport called for ${clinicName}`);

    if (!postalCodeCounts || !clinicName) {
        return res.status(400).send('不正なリクエスト: 郵便番号データとクリニック名が必要です。');
    }

    // サーバーキャッシュキー (クリニック名)
    const cacheKey = clinicName;
    if (municipalityReportCache[cacheKey]) {
        console.log(`[/api/generateMunicipalityReport] Returning cached municipality report for ${cacheKey}`);
        return res.json(municipalityReportCache[cacheKey]);
    }

    const addressAggregates = {}; // { "都道府県-市区町村": count }
    let totalValidCodesCount = 0;
    
    // ユニークな郵便番号リストに対して並列でAPIを呼び出す
    const uniquePostalCodes = Object.keys(postalCodeCounts);
    console.log(`[/api/generateMunicipalityReport] Looking up ${uniquePostalCodes.length} unique postal codes...`);

    const lookupPromises = uniquePostalCodes.map(async (postalCode) => {
        const count = postalCodeCounts[postalCode];
        const address = await postalCodeService.lookupPostalCode(postalCode); // 新サービス呼び出し
        
        if (address && address.prefecture && address.municipality) {
            // "東京都-新宿区" のようなキーを作成
            const key = `${address.prefecture}-${address.municipality}`;
            addressAggregates[key] = (addressAggregates[key] || 0) + count;
        } else {
            // APIで引けなかった場合
            addressAggregates['不明-不明'] = (addressAggregates['不明-不明'] || 0) + count;
        }
        totalValidCodesCount += count;
    });

    try {
        await Promise.all(lookupPromises); // すべてのAPI呼び出しが完了するのを待つ
        
        if (totalValidCodesCount === 0) {
            return res.json([]);
        }

        // ユーザーが要求したテーブル形式に変換
        const finalTable = Object.entries(addressAggregates).map(([key, count]) => {
            const [prefecture, municipality] = key.split('-');
            return {
                prefecture: prefecture,   // A列
                municipality: municipality, // B列
                count: count,             // C列
                percentage: (count / totalValidCodesCount) * 100 // D列
            };
        });
        
        // 件数(C列)の降順でソート
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
