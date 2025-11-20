// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/aiAnalysisService.js

const openaiService = require('./services/openai');
const googleSheetsService = require('./services/googleSheets');
const postalCodeService = require('./services/postalCodeService');
// ▼▼▼ [変更] 必要なヘルパーをインポート ▼▼▼
const { 
    getSystemPromptForDetailAnalysis, 
    getSystemPromptForRecommendationAnalysis,
    RECOMMENDATION_DISPLAY_NAMES,
    getAnalysisSheetName, // (タブ名取得用)
    formatAiJsonToMap     // (AI結果をMapに変換用)
} = require('./utils/helpers');
// ▲▲▲

// =================================================================
// === ▼▼▼ [変更] AI分析(5種)をまとめて実行・保存 ▼▼▼ ===
// =================================================================
/**
 * [新規] AI分析(5種)をすべて実行し、結果を `_AI分析` タブに保存する
 * @param {string} centralSheetId 
 * @param {string} clinicName 
 * @throws {Error} データ取得失敗時
 */
exports.runAllAiAnalysesAndSave = async (centralSheetId, clinicName) => {
    console.log(`[aiAnalysisService-All] Running for ${clinicName}`);
    
    // 1. 集計スプシから分析対象のテキストリストを *一度だけ* 取得
    const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    
    // 2. 分析対象のテキストを準備
    const analysisInputs = {
        'L': reportData.npsData.rawText || [],
        'I_good': reportData.feedbackData.i_column.results || [],
        'I_bad': reportData.feedbackData.i_column.results || [], // (I_good と同じ)
        'J': reportData.feedbackData.j_column.results || [],
        'M': reportData.feedbackData.m_column.results || []
    };
    
    const analysisPromises = [];
    const analysisTypes = ['L', 'I_good', 'I_bad', 'J', 'M'];
    
    // 3. 5種類のAI分析を並列実行
    for (const columnType of analysisTypes) {
        const textList = analysisInputs[columnType];
        
        if (textList.length === 0) {
            console.log(`[aiAnalysisService-All] No text data (0 items) for ${columnType}. Skipping AI call.`);
            // (データ0件の場合は、空のJSONに対するフォーマット処理をプッシュ)
            analysisPromises.push(Promise.resolve(formatAiJsonToMap({}, columnType)));
            continue;
        }

        const systemPrompt = getSystemPromptForDetailAnalysis(clinicName, columnType);
        
        // テキストを結合
        const truncatedList = textList.length > 100 ? textList.slice(0, 100) : textList;
        const combinedText = truncatedList.join('\n\n---\n\n');
        const inputText = combinedText.substring(0, 15000);
        
        console.log(`[aiAnalysisService-All] Sending ${truncatedList.length} comments (type: ${columnType}) to OpenAI...`);

        // OpenAI API 呼び出しをプロミス配列に追加
        const promise = openaiService.generateJsonAnalysis(systemPrompt, inputText)
            .then(analysisJson => {
                // 成功時: JSON を Map (L_ANALYSIS => "...") に変換
                return formatAiJsonToMap(analysisJson, columnType);
            })
            .catch(err => {
                console.error(`[aiAnalysisService-All] OpenAI failed for ${columnType}: ${err.message}`);
                // 失敗時: 空のJSON (デフォルト値) を Map に変換
                return formatAiJsonToMap({}, columnType);
            });
            
        analysisPromises.push(promise);
    }
    
    // 4. すべてのAI分析(5件)の完了を待つ
    const resultsMaps = await Promise.all(analysisPromises); // (Map[] が返る)
    
    // 5. 5つのMap (各3キー) を 1つの大きなMap (15キー) に統合
    const finalAiDataMap = new Map();
    resultsMaps.forEach(map => {
        map.forEach((value, key) => {
            finalAiDataMap.set(key, value);
        });
    });

    // 6. ▼▼▼ [変更] 結果をGoogleスプレッドシートの `_AI分析` タブに *一括保存* ▼▼▼
    const aiSheetName = getAnalysisSheetName(clinicName, 'AI');
    console.log(`[aiAnalysisService-All] Saving all 15 AI keys to Sheet: "${aiSheetName}"`);
    
    await googleSheetsService.saveAiAnalysisData(centralSheetId, aiSheetName, finalAiDataMap);
    
    console.log(`[aiAnalysisService-All] SUCCESS for ${clinicName}`);
};

// (古い runAndSaveAnalysis は削除)
// exports.runAndSaveAnalysis = ... (削除)


// =================================================================
// === ▼▼▼ [変更] 市区町村分析 (保存先タブ修正) ▼▼▼ ===
// =================================================================
/**
 * [変更] 市区町村レポートを生成し、`_市区町村` タブに保存する
 * @param {string} centralSheetId 
 * @param {string} clinicName 
 * @throws {Error} データが0件の場合、または分析・保存に失敗した場合
 */
exports.runAndSaveMunicipalityAnalysis = async (centralSheetId, clinicName) => {
    console.log(`[aiAnalysisService-Muni] Running for ${clinicName}`);

    // 1. 集計スプシから郵便番号データを取得 (変更なし)
    const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    const postalCodeCounts = reportData.postalCodeData.counts;
    
    if (!postalCodeCounts || Object.keys(postalCodeCounts).length === 0) {
        console.log(`[aiAnalysisService-Muni] No postal code data found for ${clinicName}.`);
        throw new Error('分析対象の郵便番号データが0件です。 (Type: MUNICIPALITY)');
    }

    // 2. 郵便番号APIで住所を検索 (変更なし)
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

    // 3. 集計 (変更なし)
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

    // 4. シート保存用の2D配列に変換 (ヘッダー含む) (変更なし)
    const header = ['都道府県', '市区町村', '件数', '割合'];
    const dataRows = finalTableData.map(row => [
        row.prefecture,
        row.municipality,
        row.count,
        row.percentage // 0.123 形式 (シート側で%書式設定)
    ]);
    const sheetData = [header, ...dataRows];

    // 5. ▼▼▼ [変更] 結果を `_市区町村` タブに保存 ▼▼▼
    const sheetName = getAnalysisSheetName(clinicName, 'MUNICIPALITY');
    console.log(`[aiAnalysisService-Muni] Saving ${dataRows.length} rows to sheet: "${sheetName}"`);
    await googleSheetsService.saveTableToSheet(centralSheetId, sheetName, sheetData);
    
    console.log(`[aiAnalysisService-Muni] SUCCESS for ${clinicName}`);
};


// =================================================================
// === ▼▼▼ [変更] おすすめ理由分析 (保存先タブ修正) ▼▼▼ ===
// =================================================================
/**
 * [変更] おすすめ理由(N列)をAI分類・集計し、`_おすすめ理由` タブに保存する
 * @param {string} centralSheetId 
 * @param {string} clinicName 
 * @throws {Error} データが0件の場合、または分析・保存に失敗した場合
 */
exports.runAndSaveRecommendationAnalysis = async (centralSheetId, clinicName) => {
    console.log(`[aiAnalysisService-Rec] Running for ${clinicName}`);

    // 1. 集計スプシからN列データを取得 (変更なし)
    const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    const { fixedCounts, otherList, fixedKeys } = reportData.recommendationData;
    
    // 2. 「その他」をAIで分類 (変更なし)
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

            // 3. AI分類結果を集計 (変更なし)
            analysisJson.classifiedResults.forEach(item => {
                if (!item.matchedCategories || item.matchedCategories.length === 0) {
                    finalCounts['その他']++;
                    return;
                }
                item.matchedCategories.forEach(category => {
                    if (finalCounts[category] !== undefined) {
                        finalCounts[category]++;
                    } else {
                        finalCounts['その他']++;
                    }
                });
            });

        } catch (aiError) {
            console.error(`[aiAnalysisService-Rec] AI classification failed for ${clinicName}: ${aiError.message}`);
            finalCounts['その他'] = (finalCounts['その他'] || 0) + otherList.length; 
        }
    }

    // 4. A/B/C テーブル形式に変換 (変更なし)
    const totalCount = Object.values(finalCounts).reduce((a, b) => a + b, 0);
    
    if (totalCount === 0) {
        console.log(`[aiAnalysisService-Rec] No recommendation data found for ${clinicName}.`);
        throw new Error('分析対象のおすすめ理由データが0件です。 (Type: RECOMMENDATION)');
    }

    const tableData = [];
    
    // 5. 表示名マッピングを適用 (変更なし)
    const allUsedKeys = new Set([...fixedKeys, ...Object.keys(finalCounts)]);
    
    allUsedKeys.forEach(originalKey => {
        // (fixedKeys 以外のキー (例: 'その他') も処理対象にする)
        if (originalKey === 'その他') {
             tableData.push({
                item: 'その他',
                count: finalCounts['その他'] || 0,
                percentage: (finalCounts['その他'] || 0) / totalCount
            });
        } else if (fixedKeys.includes(originalKey)) {
             const displayName = RECOMMENDATION_DISPLAY_NAMES[originalKey] || originalKey;
             const count = finalCounts[originalKey] || 0;
             tableData.push({
                item: displayName, // (ご要望: 表示名)
                count: count,
                percentage: (count / totalCount)
            });
        }
        // (fixedKeys にも 'その他' にも該当しないキーは無視)
    });
    
    // 件数でソート
    tableData.sort((a, b) => b.count - a.count);

    // 6. シート保存用の2D配列に変換 (ヘッダー含む) (変更なし)
    const header = ['項目', '件数', '割合']; // (ご要望: A項目, B件数, C割合)
    const dataRows = tableData.map(row => [
        row.item,
        row.count,
        row.percentage // 0.123 形式
    ]);
    const sheetData = [header, ...dataRows];

    // 7. ▼▼▼ [変更] 結果を `_おすすめ理由` タブに保存 ▼▼▼
    const sheetName = getAnalysisSheetName(clinicName, 'RECOMMENDATION');
    console.log(`[aiAnalysisService-Rec] Saving ${dataRows.length} rows to sheet: "${sheetName}"`);
    await googleSheetsService.saveTableToSheet(centralSheetId, sheetName, sheetData);

    // 8. ▼▼▼ [新規] 全体-おすすめ理由シートに集計 ▼▼▼
    console.log(`[aiAnalysisService-Rec] Updating aggregate sheet: "全体-おすすめ理由"`);
    await updateAggregateRecommendationSheet(centralSheetId, tableData);

    console.log(`[aiAnalysisService-Rec] SUCCESS for ${clinicName}`);
};

// =================================================================
// === ▼▼▼ [新規] 全体-おすすめ理由シート更新関数 ▼▼▼ ===
// =================================================================
/**
 * 全体-おすすめ理由シートを読み込み、各項目の件数を加算する
 * @param {string} centralSheetId
 * @param {Array} tableData - [{item, count, percentage}, ...]
 */
async function updateAggregateRecommendationSheet(centralSheetId, tableData) {
    const aggregateSheetName = '全体-おすすめ理由';

    try {
        // 1. 既存データを読み込む
        const response = await googleSheetsService.sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: `'${aggregateSheetName}'!A:B`,
            valueRenderOption: 'FORMATTED_VALUE'
        });

        let existingData = response.data.values || [];
        let aggregateMap = new Map();

        // 2. 既存データをMapに変換（ヘッダー行をスキップ）
        if (existingData.length > 1) {
            for (let i = 1; i < existingData.length; i++) {
                const [item, countStr] = existingData[i];
                if (item) {
                    aggregateMap.set(item, parseInt(countStr) || 0);
                }
            }
        }

        // 3. 新しいデータを加算
        tableData.forEach(row => {
            const currentCount = aggregateMap.get(row.item) || 0;
            aggregateMap.set(row.item, currentCount + row.count);
        });

        // 4. Map を配列に変換してソート
        const updatedRows = Array.from(aggregateMap.entries())
            .map(([item, count]) => [item, count])
            .sort((a, b) => b[1] - a[1]); // 件数降順

        // 5. ヘッダーを追加
        const finalData = [['項目', '件数'], ...updatedRows];

        // 6. シートに書き込む
        await googleSheetsService.saveTableToSheet(centralSheetId, aggregateSheetName, finalData);

        console.log(`[aiAnalysisService-Rec] Aggregate sheet updated: ${updatedRows.length} items`);

    } catch (error) {
        // シートが存在しない場合は新規作成
        if (error.message && error.message.includes('not found')) {
            console.log(`[aiAnalysisService-Rec] Aggregate sheet not found, creating new one...`);

            const newData = tableData
                .map(row => [row.item, row.count])
                .sort((a, b) => b[1] - a[1]);

            const finalData = [['項目', '件数'], ...newData];
            await googleSheetsService.saveTableToSheet(centralSheetId, aggregateSheetName, finalData);

            console.log(`[aiAnalysisService-Rec] New aggregate sheet created with ${newData.length} items`);
        } else {
            console.error(`[aiAnalysisService-Rec] Error updating aggregate sheet: ${error.message}`);
            throw error;
        }
    }
}
