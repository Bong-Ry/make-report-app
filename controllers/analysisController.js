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
    getAnalysisSheetName, // (タブ名取得用)
    formatAiJsonToMap,     // (AI結果をMapに変換用)
    getAiAnalysisKeys    // (AI分析のキー一覧取得用)
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
// === ▼▼▼ [変更] AI詳細分析 (再実行) ▼▼▼ ===
// =================================================================
// (このAPIは、手動で「再実行」ボタンが押された場合に使用される)
exports.generateDetailedAnalysis = async (req, res) => {
    const { centralSheetId, clinicName, columnType } = req.body;
    console.log(`POST /api/generateDetailedAnalysis called for ${clinicName}, type: ${columnType}`);

    if (!centralSheetId || !clinicName || !columnType) {
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType が必要です。');
    }

    const aiSheetName = getAnalysisSheetName(clinicName, 'AI');

    try {
        // 1. [変更] まず、`_AI分析` タブから現在のデータをすべて読み込む
        console.log(`[/api/generateDetailedAnalysis] Reading existing data from "${aiSheetName}"...`);
        const existingAiDataMap = await googleSheetsService.readAiAnalysisData(centralSheetId, aiSheetName);
        
        // 2. [変更] 依頼された 1種のAI分析のみ実行 (aiAnalysisService.js には無いので、ここで実行)
        console.log(`[/api/generateDetailedAnalysis] Running single AI analysis for ${columnType}...`);
        
        // 2a. 元データを取得
        const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
        let textList = [];
        switch (columnType) {
            case 'L': textList = reportData.npsData.rawText || []; break;
            case 'I_good': textList = reportData.feedbackData.i_column.results || []; break;
            case 'I_bad': textList = reportData.feedbackData.i_column.results || []; break; 
            case 'J': textList = reportData.feedbackData.j_column.results || []; break;
            case 'M': textList = reportData.feedbackData.m_column.results || []; break;
            default: throw new Error(`無効な分析タイプです: ${columnType}`);
        }
        
        if (textList.length === 0) {
            throw new Error(`分析対象のテキストデータが0件です。 (Type: ${columnType})`);
        }
        
        // 2b. AI呼び出し
        const systemPrompt = getSystemPromptForDetailAnalysis(clinicName, columnType);
        const truncatedList = textList.length > 100 ? textList.slice(0, 100) : textList;
        const combinedText = truncatedList.join('\n\n---\n\n');
        const inputText = combinedText.substring(0, 15000);
        const analysisJson = await openaiService.generateJsonAnalysis(systemPrompt, inputText);

        // 3. [変更] 新しい分析結果 (3キー) をMapに変換
        const newAnalysisMap = formatAiJsonToMap(analysisJson, columnType);

        // 4. [変更] 既存のMapに、新しい分析結果(3キー)を上書きマージ
        newAnalysisMap.forEach((value, key) => {
            existingAiDataMap.set(key, value);
        });
        
        // 5. [変更] マージしたMap (15キー) を `_AI分析` タブに丸ごと上書き保存
        console.log(`[/api/generateDetailedAnalysis] Saving merged data back to "${aiSheetName}"...`);
        await googleSheetsService.saveAiAnalysisData(centralSheetId, aiSheetName, existingAiDataMap);
        
        // 6. フロントエンドには、今回生成した生のJSONのみ返す
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
// === ▼▼▼ [変更] AI詳細分析 (読み出し) ▼▼▼ ===
// =================================================================
exports.getDetailedAnalysis = async (req, res) => {
    const { centralSheetId, clinicName, columnType } = req.body;
    console.log(`POST /api/getDetailedAnalysis called for ${clinicName}, type: ${columnType}`);
    
    if (!centralSheetId || !clinicName || !columnType) {
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType が必要です。');
    }
    
    try {
        // ▼▼▼ [変更] `_AI分析` タブから全データを読み込む ▼▼▼
        const aiSheetName = getAnalysisSheetName(clinicName, 'AI');
        const aiDataMap = await googleSheetsService.readAiAnalysisData(centralSheetId, aiSheetName);

        // ▼▼▼ [変更] Mapから要求された3キーを抽出して返す ▼▼▼
        const analysisData = {
            analysis: aiDataMap.get(`${columnType}_ANALYSIS`) || '（データがありません）',
            suggestions: aiDataMap.get(`${columnType}_SUGGESTIONS`) || '（データがありません）',
            overall: aiDataMap.get(`${columnType}_OVERALL`) || '（データがありません）'
        };
        
        res.json(analysisData);
    } catch (err) {
        console.error('[/api/getDetailedAnalysis] Error:', err);
        res.status(500).send(err.message || 'AI分析結果の読み込みに失敗しました。');
    }
};

// =================================================================
// === ▼▼▼ [変更] AI詳細分析 (更新・保存) ▼▼▼ ===
// =================================================================
exports.updateDetailedAnalysis = async (req, res) => {
    // ▼▼▼ [変更] sheetName の代わりに、clinicName, columnType, tabId を受け取る ▼▼▼
    const { centralSheetId, clinicName, columnType, tabId, content } = req.body;
    console.log(`POST /api/updateDetailedAnalysis called for ${clinicName} (${columnType} - ${tabId})`);
    
    if (!centralSheetId || !clinicName || !columnType || !tabId || content == null) {
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType, tabId, content が必要です。');
    }
    
    const aiSheetName = getAnalysisSheetName(clinicName, 'AI');
    
    try {
        // 1. [変更] まず、`_AI分析` タブから現在のデータをすべて読み込む
        console.log(`[/api/updateDetailedAnalysis] Reading existing data from "${aiSheetName}"...`);
        const existingAiDataMap = await googleSheetsService.readAiAnalysisData(centralSheetId, aiSheetName);

        // 2. [変更] 保存先のキー (例: "L_ANALYSIS") を特定
        // (helpers.js にこのための関数は無いので、ここで組み立てる)
        let keyToUpdate;
        switch(tabId) {
            case 'analysis': keyToUpdate = `${columnType}_ANALYSIS`; break;
            case 'suggestions': keyToUpdate = `${columnType}_SUGGESTIONS`; break;
            case 'overall': keyToUpdate = `${columnType}_OVERALL`; break;
            default: throw new Error(`無効なタブIDです: ${tabId}`);
        }
        
        // 3. [変更] 既存のMapの、該当キーの値を上書き
        if (!getAiAnalysisKeys().includes(keyToUpdate)) {
             throw new Error(`無効な更新キーです: ${keyToUpdate}`);
        }
        existingAiDataMap.set(keyToUpdate, content);
        console.log(`[/api/updateDetailedAnalysis] Updating key: "${keyToUpdate}"`);

        // 4. [変更] マージしたMap (15キー) を `_AI分析` タブに丸ごと上書き保存
        await googleSheetsService.saveAiAnalysisData(centralSheetId, aiSheetName, existingAiDataMap);
        
        res.json({ status: 'ok', message: `分析結果 (${keyToUpdate}) を更新しました。` });
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
        const sheetName = getAnalysisSheetName(clinicName, 'MUNICIPALITY');
        console.log(`[/api/generateMunicipalityReport] Reading table from sheet: "${sheetName}"`);

        // (saveTableToSheet が [ヘッダー, [データ], [データ]] 形式で保存)
        const tableData = await readTableFromSheet(centralSheetId, sheetName);
        
        if (!tableData) { // (readTableFromSheet が null を返す)
            console.log(`[/api/generateMunicipalityReport] No data found in "${sheetName}".`);
            return res.json([]); // データが存在しない
        }

        // 2. フロントエンドが期待する形式 (割合を 12.3 形式) に変換
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
        // 1. シートから `_おすすめ理由` タブを読み込む
        const sheetName = getAnalysisSheetName(clinicName, 'RECOMMENDATION');
        console.log(`[/api/getRecommendationReport] Reading table from sheet: "${sheetName}"`);
        
        const tableData = await readTableFromSheet(centralSheetId, sheetName);

        if (!tableData) { // (readTableFromSheet が null を返す)
            console.log(`[/api/getRecommendationReport] No data found in "${sheetName}".`);
            return res.json([['カテゴリ', '件数']]); // 空のヘッダー
        }
        
        // 2. フロントエンドの円グラフが期待する形式 [ [項目, 件数], ... ] に変換
        const chartData = [['カテゴリ', '件数']]; // ヘッダー
        
        tableData.forEach(row => {
            const item = row[0] || '不明'; // A列: 項目
            const count = parseFloat(row[1]) || 0; // B列: 件数
            // (C列の割合はここでは不要)
            
            // (バックグラウンドで0件も保存するようにしたので、ここでフィルタリング)
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


// =================================================================
// === ▼▼▼ [新規] テーブル読み込み専用ヘルパー ▼▼▼ ===
// =================================================================

/**
 * [新規ヘルパー] テーブルデータ（市区町村、おすすめ理由）を読み込む
 * @param {string} centralSheetId
 * @param {string} sheetName (例: "クリニックA_市区町村")
 * @returns {Promise<string[][] | null>} (ヘッダー行を*含まない* 2D配列)
 */
async function readTableFromSheet(centralSheetId, sheetName) {
    // ▼▼▼ [修正] sheets オブジェクトを googleSheetsService から取得
    const sheets = googleSheetsService.sheets; 
    
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    try {
        // A:D (市区町村) または A:C (おすすめ理由) の最大範囲を読み込む
        const range = `'${sheetName}'!A:D`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
            valueRenderOption: 'UNFORMATTED_VALUE' // 割合を 0.123 で取得
        });

        const rows = response.data.values;

        if (!rows || rows.length < 2) { // (ヘッダー + データ)
            return null; // データが存在しない
        }

        rows.shift(); // ヘッダーを捨てる
        return rows; // データ行のみ返す

    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            console.warn(`[readTableFromSheet] Sheet "${sheetName}" not found. Returning null.`);
            return null; // シートが存在しない
        }
        console.error(`[readTableFromSheet] Error reading table data: ${e.message}`, e);
        throw new Error(`分析テーブル(${sheetName})の読み込みに失敗しました: ${e.message}`);
    }
}
