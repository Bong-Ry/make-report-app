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
    getAiAnalysisKeys,    // (AI分析のキー一覧取得用)
    getAiAnalysisCellAndTitle // [変更] 新しいヘルパーを追加
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
             // (aiAnalysisService.js と同様に、404を返す)
            console.warn(`[/api/generateDetailedAnalysis] No text data (0 items) for ${columnType}.`);
            return res.status(404).send(`分析対象のテキストデータが0件です。 (Type: ${columnType})`);
        }
        
        // 2b. AI呼び出し
        const systemPrompt = getSystemPromptForDetailAnalysis(clinicName, columnType);
        
        // ▼▼▼ [修正] ダミーではなく、aiAnalysisService.js 経由で呼び出す ▼▼▼
        // (aiAnalysisService.js が openaiService を require している前提)
        // [注] aiAnalysisService.js には runSingleAiAnalysis が無いため、ここで openaiService を直接呼ぶか、
        // aiAnalysisService.js に新設する必要があります。
        // ここでは、元のファイル(analysisController.js)が aiAnalysisService.js を呼んでいる前提のまま、
        // aiAnalysisService.js に runSingleAiAnalysis が存在すると仮定して進めます。
        
        // [訂正] 元の analysisController.js は aiAnalysisService を呼んでいなかったため、
        // 元のファイルの依存関係 (openaiService を直接呼ぶ) に戻します。
        // [再訂正] ユーザー提供のファイルでは aiAnalysisService.js を呼んでいました。
        // aiAnalysisService.js には openai を直接呼ぶ機能がないため、
        // ここで aiAnalysisService.js が内部で使っている openaiService を呼び出します。
        // （aiAnalysisService.js の中身を確認したところ、openaiService を require していました）
        
        const openaiService = require('../services/openai'); // 不足していた require を追加
        const analysisJson = await openaiService.generateJsonAnalysis(systemPrompt, textList);


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
    // ▼▼▼ [修正] `tabId`, `content` の代わりに `analysis`, `suggestions`, `overall` を受け取る
    const { centralSheetId, clinicName, columnType, analysis, suggestions, overall } = req.body;
    console.log(`POST /api/updateDetailedAnalysis called for ${clinicName} (${columnType})`);
    
    // [修正] バリデーション
    if (!centralSheetId || !clinicName || !columnType || analysis == null || suggestions == null || overall == null) {
        return res.status(400).send('不正なリクエスト: centralSheetId, clinicName, columnType, analysis, suggestions, overall が必要です。');
    }
    
    const aiSheetName = getAnalysisSheetName(clinicName, 'AI');
    
    try {
        // 1. [変更] まず、`_AI分析` タブから現在のデータをすべて読み込む
        console.log(`[/api/updateDetailedAnalysis] Reading existing data from "${aiSheetName}"...`);
        const existingAiDataMap = await googleSheetsService.readAiAnalysisData(centralSheetId, aiSheetName);

        // 2. [修正] 保存先のキー (例: "L_ANALYSIS" など3つ) を特定
        const analysisKey = `${columnType}_ANALYSIS`;
        const suggestionsKey = `${columnType}_SUGGESTIONS`;
        const overallKey = `${columnType}_OVERALL`;
        
        // 3. [修正] 既存のMapの、該当キーの値を3つとも上書き
        const allKeys = getAiAnalysisKeys();
        if (!allKeys.includes(analysisKey) || !allKeys.includes(suggestionsKey) || !allKeys.includes(overallKey)) {
             throw new Error(`無効な更新キーです: ${columnType}`);
        }
        
        existingAiDataMap.set(analysisKey, analysis);
        existingAiDataMap.set(suggestionsKey, suggestions);
        existingAiDataMap.set(overallKey, overall);
        console.log(`[/api/updateDetailedAnalysis] Updating keys for: "${columnType}"`);

        // 4. [変更] マージしたMap (15キー) を `_AI分析` タブに丸ごと上書き保存
        await googleSheetsService.saveAiAnalysisData(centralSheetId, aiSheetName, existingAiDataMap);
        
        res.json({ status: 'ok', message: `分析結果 (${columnType}) を更新しました。` });
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

        // (readTableFromSheet はこのファイルにないので、googleSheetsService を呼ぶ)
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
// === ▼▼▼ [修正] テーブル読み込み専用ヘルパー (googleSheetsService を呼ぶ) ▼▼▼ ===
// =================================================================
/**
 * [修正] googleSheetsService を呼び出すラッパー
 */
async function readTableFromSheet(centralSheetId, sheetName) {
    const sheets = googleSheetsService.sheets; 
    
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    try {
        const range = `'${sheetName}'!A:D`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
            valueRenderOption: 'UNFORMATTED_VALUE'
        });

        const rows = response.data.values;

        if (!rows || rows.length < 2) { 
            return null;
        }

        rows.shift(); 
        return rows; 

    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            console.warn(`[readTableFromSheet] Sheet "${sheetName}" not found. Returning null.`);
            return null; 
        }
        console.error(`[readTableFromSheet] Error reading table data: ${e.message}`, e);
        throw new Error(`分析テーブル(${sheetName})の読み込みに失敗しました: ${e.message}`);
    }
}


// =================================================================
// === ▼▼▼ [大幅修正] コメント編集用コントローラー (新仕様) ▼▼▼ ===
// =================================================================

/**
 * [大幅修正] /api/getCommentData (読み取り)
 * (HTMLがコメントシートからデータを読み込む)
 */
exports.getCommentData = async (req, res) => {
    // ▼▼▼ [変更] `commentType` の代わりに `sheetName` を受け取る
    const { centralSheetId, sheetName } = req.body;
    console.log(`POST /api/getCommentData called for sheet: ${sheetName}`);

    if (!centralSheetId || !sheetName) {
        return res.status(400).send('不正なリクエスト: centralSheetId, sheetName が必要です。');
    }

    try {
        // ▼▼▼ [変更] `readCommentsBySheet` を呼び出す
        const data = await googleSheetsService.readCommentsBySheet(centralSheetId, sheetName);
        res.json(data);
    } catch (error) {
        console.error('[/api/getCommentData] Error reading comment sheet:', error);
        res.status(500).send(`コメントシートの読み込み中にエラーが発生しました: ${error.message}`);
    }
};

/**
 * [大幅修正] /api/updateCommentData (書き込み)
 * (HTMLが編集したコメントをシートに保存する)
 */
exports.updateCommentData = async (req, res) => {
    // ▼▼▼ [変更] `sheetName`, `cell`, `value` を受け取る
    const { centralSheetId, sheetName, cell, value } = req.body;
    console.log(`POST /api/updateCommentData called for sheet: ${sheetName}, cell: ${cell}`);

    // バリデーション
    if (!centralSheetId || !sheetName || !cell || value == null) {
        return res.status(400).send('不正なリクエスト: centralSheetId, sheetName, cell, value が必要です。');
    }
    // (A1, B5, AA10 などの形式を簡易チェック)
    if (!/^[A-Z]+[1-9][0-9]*$/.test(cell)) {
         return res.status(400).send(`不正なリクエスト: 無効なセル指定です (Cell: ${cell})。`);
    }

    try {
        // ▼▼▼ [変更] `updateCommentSheetCell` を呼び出す
        await googleSheetsService.updateCommentSheetCell(centralSheetId, sheetName, cell, value);
        res.json({ status: 'ok', message: `セル ${cell} を更新しました。` });
    } catch (error) {
        console.error('[/api/updateCommentData] Error updating comment cell:', error);
        res.status(500).send(`コメントシートの更新中にエラーが発生しました: ${error.message}`);
    }
};

// ▼▼▼ [ここから変更] ▼▼▼
/**
 * [新規] AI分析タブの単一セルを取得するAPI
 */
exports.getSingleAnalysisCell = async (req, res) => {
    const { centralSheetId, clinicName, columnType, tabId } = req.body;
    // (columnType = 'L', tabId = 'suggestions' など)
    
    console.log(`POST /api/getSingleAnalysisCell called for ${clinicName}, Type: ${columnType}, Tab: ${tabId}`);

    if (!centralSheetId || !clinicName || !columnType || !tabId) {
        return res.status(400).send('不正なリクエスト: 必要なパラメータが不足しています。');
    }

    try {
        // 1. (例: "クリニックA_AI分析")
        const aiSheetName = getAnalysisSheetName(clinicName, 'AI'); 
        
        // 2. (例: { cell: "B3", pentagonText: "改善提案" })
        const { cell, pentagonText } = getAiAnalysisCellAndTitle(columnType, tabId);
        
        if (!cell) {
             throw new Error(`無効なキー(${columnType}, ${tabId})のためセルを特定できませんでした。`);
        }

        // 3. (例: "B3" の値を読み込む)
        const content = await googleSheetsService.readSingleCell(centralSheetId, aiSheetName, cell);
        
        res.json({
            content: content || '（データがありません）',
            pentagonText: pentagonText // (五角形用のテキストも返す)
        });

    } catch (error) {
        console.error('[/api/getSingleAnalysisCell] Error:', error);
        res.status(500).send(error.message || '分析データの単一セル読み込みに失敗しました。');
    }
};
// ▲▲▲ [変更ここまで] ▲▲▲
