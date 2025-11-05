// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/aiAnalysisService.js

const openaiService = require('./services/openai');
const googleSheetsService = require('./services/googleSheets');
// ▼▼▼ [変更] postalCodeService, helpers (おすすめ理由マッピング含む) をインポート ▼▼▼
const postalCodeService = require('./services/postalCodeService');
const { 
    getSystemPromptForDetailAnalysis, 
    getSystemPromptForRecommendationAnalysis,
    RECOMMENDATION_DISPLAY_NAMES // (おすすめ理由の表示名マッピング)
} = require('./utils/helpers');
// ▲▲▲

/**
 * [変更] AI分析(5種)を実行し、結果を *単一分析シート* に保存する
 * @param {string} centralSheetId 
 * @param {string} clinicName 
 * @param {string} columnType (例: 'L', 'I_good', 'J'...)
 * @returns {object} AIが生成した生のJSON
 * @throws {Error} テキストデータが0件の場合、またはAI分析・保存に失敗した場合
 */
exports.runAndSaveAnalysis = async (centralSheetId, clinicName, columnType) => {
    console.log(`[aiAnalysisService-Core] Running for ${clinicName}, type: ${columnType}`);
    
    const systemPrompt = getSystemPromptForDetailAnalysis(clinicName, columnType);
    if (!systemPrompt) {
        throw new Error(`[aiAnalysisService-Core] Invalid analysis type (no prompt): ${columnType}`);
    }

    // 1. 集計スプシから分析対象のテキストリストを取得
    const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    
    let textList = [];
    switch (columnType) {
        case 'L': textList = reportData.npsData.rawText || []; break;
        case 'I_good': textList = reportData.feedbackData.i_column.results || []; break;
        case 'I_bad': textList = reportData.feedbackData.i_column.results || []; break; 
        case 'J': textList = reportData.feedbackData.j_column.results || []; break;
        case 'M': textList = reportData.feedbackData.m_column.results || []; break;
        default: throw new Error(`[aiAnalysisService-Core] 無効な分析タイプです: ${columnType}`);
    }
    
    if (textList.length === 0) {
        console.log(`[aiAnalysisService-Core] No text data (0 items) found for ${columnType}.`);
        throw new Error(`分析対象のテキストデータが0件です。 (Type: ${columnType})`);
    }

    // 2. 入力テキストを結合・制限 (変更なし)
    const truncatedList = textList.length > 100 ? textList.slice(0, 100) : textList;
    const combinedText = truncatedList.join('\n\n---\n\n');
    const inputText = combinedText.substring(0, 15000);

    console.log(`[aiAnalysisService-Core] Sending ${truncatedList.length} comments (input text length: ${inputText.length}) to OpenAI...`);

    // 3. OpenAI API 呼び出し (変更なし)
    const analysisJson = await openaiService.generateJsonAnalysis(systemPrompt, inputText);
    
    // 4. ▼▼▼ [変更] 結果をGoogleスプレッドシートの *単一分析シート* に保存 ▼▼▼
    console.log(`[aiAnalysisService-Core] Saving analysis results to Google Sheet (Single Analysis Sheet)...`);
    
    // 3つのセル (Analysis, Suggestions, Overall) に分けて保存
    const analysisText = (analysisJson.analysis && analysisJson.analysis.themes) ? analysisJson.analysis.themes.map(t => `【${t.title}】\n${t.summary}`).join('\n\n---\n\n') : '（分析データがありません）';
    const suggestionsText = (analysisJson.suggestions && analysisJson.suggestions.items) ? analysisJson.suggestions.items.map(i => `【${i.themeTitle}】\n${i.suggestion}`).join('\n\n---\n\n') : '（改善提案データがありません）';
    const overallText = (analysisJson.overall && analysisJson.overall.summary) ? analysisJson.overall.summary : '（総評データがありません）';

    await Promise.all([
        googleSheetsService.saveToAnalysisSheet(centralSheetId, clinicName, `${columnType}_ANALYSIS`, analysisText),
        googleSheetsService.saveToAnalysisSheet(centralSheetId, clinicName, `${columnType}_SUGGESTIONS`, suggestionsText),
        googleSheetsService.saveToAnalysisSheet(centralSheetId, clinicName, `${columnType}_OVERALL`, overallText)
    ]);

    // 5. AIが生成した生のJSONを返す (変更なし)
    return analysisJson;
};


// =================================================================
// === ▼▼▼ [新規] 市区町村分析 (バックグラウンド実行用) ▼▼▼ ===
// =================================================================
/**
 * [新規] 市区町村レポートを生成し、結果を単一分析シートに保存する
 * (analysisController.js の generateMunicipalityReport の中身を移植)
 * @param {string} centralSheetId 
 * @param {string} clinicName 
 * @throws {Error} データが0件の場合、または分析・保存に失敗した場合
 */
exports.runAndSaveMunicipalityAnalysis = async (centralSheetId, clinicName) => {
    console.log(`[aiAnalysisService-Muni] Running for ${clinicName}`);

    // 1. 集計スプシから郵便番号データを取得
    const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    const postalCodeCounts = reportData.postalCodeData.counts;
    
    if (!postalCodeCounts || Object.keys(postalCodeCounts).length === 0) {
        console.log(`[aiAnalysisService-Muni] No postal code data found for ${clinicName}.`);
        throw new Error('分析対象の郵便番号データが0件です。 (Type: MUNICIPALITY)');
    }

    // 2. 郵便番号APIで住所を検索
    const uniquePostalCodes = Object.keys(postalCodeCounts);
    console.log(`[aiAnalysisService-Muni] Looking up ${uniquePostalCodes.length} unique postal codes...`);
    
    const addressAggregates = {}; // { "都道府県-市区町村": count }
    let totalValidCodesCount = 0;

    const lookupPromises = uniquePostalCodes.map(async (postalCode) => {
        const count = postalCodeCounts[postalCode];
        const address = await postalCodeService.lookupPostalCode(postalCode); 
        
        if (address && address.prefecture && address.municipality) {
            const key = `${address.prefecture}-${address.municipality}`;
            addressAggregates[key] = (addressAggregates[key] || 0) + count;
        } else {
            addressAggregates['不明-不明'] = (addressAggregates['不明-不明'] || 0) + count;
        }
        totalValidCodesCount += count;
    });

    // 3. 集計
    await Promise.all(lookupPromises);
    
    if (totalValidCodesCount === 0) {
        throw new Error('有効な郵便番号データが0件です。 (Type: MUNICIPALITY)');
    }

    const finalTableData = Object.entries(addressAggregates).map(([key, count]) => {
        const [prefecture, municipality] = key.split('-');
        return {
            prefecture: prefecture,
            municipality: municipality,
            count: count,
            percentage: (count / totalValidCodesCount) // 割合 (0.123 形式)
        };
    });
    
    finalTableData.sort((a, b) => b.count - a.count);

    // 4. [変更] シート保存用の2D配列に変換 (ヘッダー含む)
    const header = ['都道府県', '市区町村', '件数', '割合'];
    const dataRows = finalTableData.map(row => [
        row.prefecture,
        row.municipality,
        row.count,
        row.percentage // 0.123 形式 (シート側で%書式設定)
    ]);
    const sheetData = [header, ...dataRows];

    // 5. [変更] 結果を単一分析シートに保存
    console.log(`[aiAnalysisService-Muni] Saving ${dataRows.length} rows to analysis sheet...`);
    await googleSheetsService.saveToAnalysisSheet(centralSheetId, clinicName, 'MUNICIPALITY_TABLE', sheetData);
    
    console.log(`[aiAnalysisService-Muni] SUCCESS for ${clinicName}`);
};


// =================================================================
// === ▼▼▼ [新規] おすすめ理由分析 (バックグラウンド実行用) ▼▼▼ ===
// =================================================================
/**
 * [新規] おすすめ理由(N列)をAI分類・集計し、結果を単一分析シートに保存する
 * (ご要望: A項目, B件数, C割合 | 表示名マッピング)
 * @param {string} centralSheetId 
 * @param {string} clinicName 
 * @throws {Error} データが0件の場合、または分析・保存に失敗した場合
 */
exports.runAndSaveRecommendationAnalysis = async (centralSheetId, clinicName) => {
    console.log(`[aiAnalysisService-Rec] Running for ${clinicName}`);

    // 1. 集計スプシからN列データを取得
    const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    const { fixedCounts, otherList, fixedKeys } = reportData.recommendationData;
    
    // 2. 「その他」をAIで分類
    let finalCounts = { ...fixedCounts };
    finalCounts['その他'] = 0; // AI分類不能な場合の受け皿

    if (otherList && otherList.length > 0) {
        console.log(`[aiAnalysisService-Rec] Calling AI for ${otherList.length} "other" items...`);
        const systemPrompt = getSystemPromptForRecommendationAnalysis(fixedKeys);
        const inputText = otherList.join('\n');
        
        try {
            const analysisJson = await openaiService.generateJsonAnalysis(systemPrompt, inputText);
            
            if (!analysisJson || !Array.isArray(analysisJson.classifiedResults)) {
                throw new Error('AIが予期しない分類結果フォーマットを返しました。');
            }

            // 3. AI分類結果を集計
            analysisJson.classifiedResults.forEach(item => {
                if (!item.matchedCategories || item.matchedCategories.length === 0) {
                    finalCounts['その他']++;
                    return;
                }
                item.matchedCategories.forEach(category => {
                    // category が 'その他' または fixedKeys のどれか
                    if (finalCounts[category] !== undefined) {
                        finalCounts[category]++;
                    } else {
                        // AIが fixedKeys 以外の新カテゴリを返した場合 (プロンプト指示違反)
                        // もしくは、"その他" に分類された場合
                        finalCounts['その他']++;
                    }
                });
            });

        } catch (aiError) {
            console.error(`[aiAnalysisService-Rec] AI classification failed for ${clinicName}: ${aiError.message}`);
            // AI分類が失敗しても、固定キーと未分類の「その他」だけで集計を続行
            finalCounts['その他'] = (finalCounts['その他'] || 0) + otherList.length; // AI分類失敗 = すべて「その他」扱い
        }
    }

    // 4. ご要望の A/B/C テーブル形式に変換
    const totalCount = Object.values(finalCounts).reduce((a, b) => a + b, 0);
    
    if (totalCount === 0) {
        console.log(`[aiAnalysisService-Rec] No recommendation data found for ${clinicName}.`);
        throw new Error('分析対象のおすすめ理由データが0件です。 (Type: RECOMMENDATION)');
    }

    const tableData = [];
    
    // 5. [変更] ご要望の表示名マッピングを適用
    // (fixedKeys: 元のキー, RECOMMENDATION_DISPLAY_NAMES: マッピング)
    
    // まず固定キーをマッピングしながら処理
    // (fixedKeys にないキーが finalCounts にあっても無視される)
    fixedKeys.forEach(originalKey => {
        // ▼▼▼ ご要望の表示名マッピングを適用 ▼▼▼
        const displayName = RECOMMENDATION_DISPLAY_NAMES[originalKey] || originalKey;
        const count = finalCounts[originalKey] || 0;
        
        // (件数が 0 でも項目としては追加する、という仕様もありうるが、ここでは 0 は除外)
        // if (count > 0) {
            tableData.push({
                item: displayName, // (ご要望: 表示名)
                count: count,
                percentage: (count / totalCount)
            });
        // }
    });

    // 次に「その他」を追加
    const otherCount = finalCounts['その他'] || 0;
    // if (otherCount > 0) {
         tableData.push({
            item: 'その他', // 「その他」の表示名はそのまま
            count: otherCount,
            percentage: (otherCount / totalCount)
        });
    // }

    // 件数でソート
    tableData.sort((a, b) => b.count - a.count);

    // 6. シート保存用の2D配列に変換 (ヘッダー含む)
    // (ご要望: A項目, B件数, C割合)
    const header = ['項目', '件数', '割合'];
    const dataRows = tableData.map(row => [
        row.item,
        row.count,
        row.percentage // 0.123 形式
    ]);
    const sheetData = [header, ...dataRows];

    // 7. [変更] 結果を単一分析シートに保存
    console.log(`[aiAnalysisService-Rec] Saving ${dataRows.length} rows to analysis sheet...`);
    await googleSheetsService.saveToAnalysisSheet(centralSheetId, clinicName, 'RECOMMENDATION_TABLE', sheetData);

    console.log(`[aiAnalysisService-Rec] SUCCESS for ${clinicName}`);
};
