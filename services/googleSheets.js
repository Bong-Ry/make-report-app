const { google } = require('googleapis');
const { getSpreadsheetIdFromUrl } = require('../utils/helpers'); // 既存のヘルパー

// ▼▼▼ GAS WebアプリURL (設定済み) ▼▼▼
const GAS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzn4rNw6NttPPmcJBpSKJifK8-Mb1CatsGhqvYF5G6BIAf6bOUuNS_E72drg0tH9re-qQ/exec';

const KEYFILEPATH = '/etc/secrets/credentials.json';
const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets', // 読み書き
    'https://www.googleapis.com/auth/drive.file'    // ファイルの検索・移動・作成
];
const MASTER_FOLDER_ID = '1_pJQKl5-RRi6h-U3EEooGmPkTrkF1Vbj'; // 集計スプシ作成先フォルダ

let sheets;
let drive; // Drive API クライアント

// --- 初期化 (変更なし) ---
try {
    const auth = new google.auth.GoogleAuth({ keyFile: KEYFILEPATH, scopes: SCOPES });
    sheets = google.sheets({ version: 'v4', auth });
    drive = google.drive({ version: 'v3', auth }); // Drive APIを初期化
    console.log('Google Sheets (R/W) and Drive API clients initialized successfully.');
} catch (err) {
    console.error('Failed to initialize Google API clients:', err);
}

// --- (変更なし) マスターシートからクリニックリストを取得 ---
exports.getMasterClinicList = async () => {
    // (既存のコードのまま - 変更なし)
    const currentMasterSheetId = process.env.MASTER_SHEET_ID;
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    if (!currentMasterSheetId) throw new Error('サーバー設定エラー: マスターシートIDがありません。');
    const MASTER_RANGE = 'シート1!A2:A';
    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: currentMasterSheetId,
            range: MASTER_RANGE,
        });
        const rows = response.data.values;
        return (rows && rows.length > 0) ? rows.map((row) => row[0]).filter(Boolean) : [];
    } catch (err) {
        console.error('[googleSheetsService] Master Sheet API (clinics) returned an error: ' + err);
        throw new Error('マスターシートのクリニック一覧読み込みに失敗しました。');
    }
};

// --- (変更なし) マスターシートからクリニック名とURLのマップを取得 ---
exports.getMasterClinicUrls = async () => {
    // (既存のコードのまま - 変更なし)
    const currentMasterSheetId = process.env.MASTER_SHEET_ID;
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    if (!currentMasterSheetId) throw new Error('サーバー設定エラー: マスターシートIDがありません。');
    const MASTER_CLINIC_URL_RANGE = 'シート1!A2:B';
    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: currentMasterSheetId,
            range: MASTER_CLINIC_URL_RANGE,
        });
        const rows = response.data.values;
        if (!rows || rows.length === 0) return null;
        const urlMap = {};
        rows.forEach(row => {
            if (row[0] && row[1]) {
                urlMap[row[0]] = row[1];
            }
        });
        return urlMap;
    } catch (err) {
        console.error('[googleSheetsService] Master Sheet API (URLs) returned an error: ' + err);
        throw new Error('マスターシートのURL一覧読み込みに失敗しました。');
    }
};

// =================================================================
// === ▼▼▼ 関数 (1/X) (GAS呼び出し・変更なし) ▼▼▼ ===
// =================================================================
exports.findOrCreateCentralSheet = async (periodText) => {
    const fileName = periodText;
    console.log(`[googleSheetsService] Finding or creating central sheet via GAS: "${fileName}"`);
    try {
        const response = await fetch(GAS_WEB_APP_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                periodText: fileName,
                folderId: MASTER_FOLDER_ID
            })
        });
        if (!response.ok) {
            throw new Error(`GAS Web App request failed with status ${response.status}: ${await response.text()}`);
        }
        const result = await response.json();
        if (result.status === 'ok' && result.spreadsheetId) {
            console.log(`[googleSheetsService] GAS operation successful. ID: ${result.spreadsheetId}`);
            return result.spreadsheetId;
        } else {
            console.error('[googleSheetsService] GAS Web App returned an error:', result.message);
            throw new Error(`GAS側でのシート作成に失敗しました: ${result.message || '不明なエラー'}`);
        }
    } catch (err) {
        console.error(`[googleSheetsService] Error in findOrCreateCentralSheet (GAS) for "${fileName}".`);
        console.error(err); 
        throw new Error(`集計スプレッドシートの検索または作成に失敗しました (GAS): ${err.message}`);
    }
};

// =================================================================
// === ▼▼▼ 更新関数 (2/X) (要求 #4 ヘッダーなし転記) ▼▼▼ ===
// =================================================================
exports.fetchAndAggregateReportData = async (clinicUrls, period, centralSheetId) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    if (!centralSheetId) throw new Error('集計スプレッドシートIDが指定されていません。');

    const startDate = new Date(period.start + '-01T00:00:00Z');
    const [endYear, endMonth] = period.end.split('-').map(Number);
    const endDate = new Date(Date.UTC(endYear, endMonth, 0));
    endDate.setUTCHours(23, 59, 59, 999);
    console.log(`[googleSheetsService-ETL] Filtering data between ${startDate.toISOString()} and ${endDate.toISOString()}`);

    const processedClinics = []; 

    for (const clinicName in clinicUrls) {
        const sourceSheetId = getSpreadsheetIdFromUrl(clinicUrls[clinicName]);
        if (!sourceSheetId) {
            console.warn(`[googleSheetsService-ETL] Invalid URL for ${clinicName}. Skipping.`);
            continue;
        }
        
        console.log(`[googleSheetsService-ETL] Processing ${clinicName} (Source ID: ${sourceSheetId})`);

        try {
            // 1. 元データ（フォームの回答 1）を読み取る
            const range = "'フォームの回答 1'!A:R"; // R列(郵便番号)まで
            const clinicDataResponse = await sheets.spreadsheets.values.get({
                spreadsheetId: sourceSheetId, 
                range: range,
                dateTimeRenderOption: 'SERIAL_NUMBER',
                valueRenderOption: 'UNFORMATTED_VALUE'
            });

            const clinicDataRows = clinicDataResponse.data.values;
            if (!clinicDataRows || clinicDataRows.length < 2) {
                console.log(`[googleSheetsService-ETL] No data or only header found for ${clinicName}. Skipping.`);
                continue;
            }

            const header = clinicDataRows.shift();
            const timestampIndex = 0;
            const filteredRows = [];

            // 2. 期間でフィルタリング (変更なし)
            clinicDataRows.forEach((row) => {
                const serialValue = row[timestampIndex];
                if (typeof serialValue !== 'number' || serialValue <= 0 || isNaN(serialValue)) return;
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                const timestamp = new Date(excelEpoch.getTime() + serialValue * 24 * 60 * 60 * 1000);

                if (timestamp.getTime() >= startDate.getTime() && timestamp.getTime() <= endDate.getTime()) {
                    filteredRows.push(row);
                }
            });

            console.log(`[googleSheetsService-ETL] For ${clinicName}: ${filteredRows.length} rows matched period.`);

            if (filteredRows.length > 0) {
                // 3. 集計スプシに「クリニック名」のタブを作成（またはクリア） (変更なし)
                const clinicSheetTitle = clinicName; 
                await findOrCreateSheet(centralSheetId, clinicSheetTitle);
                
                // 4. データを「クリニック名」タブに書き込み
                await clearSheet(centralSheetId, clinicSheetTitle);
                
                // ▼▼▼ [変更点] ヘッダー (header) を除き、データ (filteredRows) のみ書き込む ▼▼▼
                await writeData(centralSheetId, clinicSheetTitle, filteredRows);
                console.log(`[googleSheetsService-ETL] Wrote ${filteredRows.length} rows to sheet: "${clinicSheetTitle}" (HEADERLESS)`);

                // 5. データを「全体」タブに「追記」 (変更なし)
                await writeData(centralSheetId, '全体', filteredRows, true); // append = true
                console.log(`[googleSheetsService-ETL] Appended ${filteredRows.length} rows to sheet: "全体"`);
            }
            
            processedClinics.push(clinicName);

        } catch (e) {
            if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
                console.error(`[googleSheetsService-ETL] Error for ${clinicName}: Sheet 'フォームの回答 1' not found or invalid range. Skipping. Error: ${e.toString()}`);
            } else {
                console.error(`[googleSheetsService-ETL] Error processing sheet for ${clinicName}: ${e.toString()}`, e.stack);
            }
            continue; // 次のクリニックへ
        }
    }
    
    // 処理が成功したクリニック名のリストを返す
    return processedClinics;
};

// =================================================================
// === ▼▼▼ 関数 (3/X) (集計 - 変更なし) ▼▼▼ ===
// =================================================================
exports.getReportDataForCharts = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    console.log(`[googleSheetsService-AGG] Aggregating data from Sheet ID: ${centralSheetId}, Tab: "${sheetName}"`);

    // --- 集計用の定義 (変更なし) ---
    const satisfactionKeys = ['非常に満足', '満足', 'ふつう', '不満', '非常に不満'];
    const ageKeys = ['10代', '20代', '30代', '40代'];
    const childrenKeys = ['1人', '2人', '3人', '4人', '5人以上'];
    const recommendationKeys = [
        'インターネット（Googleの口コミ）', 'インターネット（SNS）', 'インターネット（産院のホームページ）',
        '知人の紹介', '家族の紹介', '自宅からの距離', 'インターネット（情報サイト）'
    ];
    const initializeCounts = (keys) => keys.reduce((acc, key) => { acc[key] = 0; return acc; }, {});
    const createChartData = (counts, keys) => {
        const chartData = [['カテゴリ', '件数']];
        keys.forEach(key => {
            chartData.push([key, counts[key] || 0]);
        });
        return chartData;
    };
    // --- 集計用変数の初期化 (変更なし) ---
    const allNpsReasons = [], allFeedbacks_I = [], allFeedbacks_J = [], allFeedbacks_M = [];
    const satisfactionCounts_B = initializeCounts(satisfactionKeys), satisfactionCounts_C = initializeCounts(satisfactionKeys), satisfactionCounts_D = initializeCounts(satisfactionKeys), satisfactionCounts_E = initializeCounts(satisfactionKeys), satisfactionCounts_F = initializeCounts(satisfactionKeys), satisfactionCounts_G = initializeCounts(satisfactionKeys), satisfactionCounts_H = initializeCounts(satisfactionKeys);
    const childrenCounts_P = initializeCounts(childrenKeys);
    const ageCounts_O = initializeCounts(ageKeys);
    const incomeCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0 };
    const postalCodeCounts = {};
    const npsScoreCounts = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0 };
    const recommendationCounts = initializeCounts(recommendationKeys);
    const recommendationOthers = [];
    
    try {
        // 1. 集計スプレッドシートからデータを読み取る
        const range = `'${sheetName}'!A:R`;
        const clinicDataResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
            dateTimeRenderOption: 'SERIAL_NUMBER', 
            valueRenderOption: 'UNFORMATTED_VALUE'
        });

        const clinicDataRows = clinicDataResponse.data.values;

        // ▼▼▼ [変更点] ヘッダー行がなくなったため、 .length < 1 に変更 ▼▼▼
        if (!clinicDataRows || clinicDataRows.length < 1) { 
            console.log(`[googleSheetsService-AGG] No data found in "${sheetName}".`);
            return buildReportDataObject(null); 
        }

        // ▼▼▼ [変更点] ヘッダー行を shift() しない ▼▼▼
        // const header = clinicDataRows.shift();
        
        // ▼▼▼ [変更点] インデックスはヘッダーがないため、0から固定で指定 ▼▼▼
        // (元データ 'フォームの回答 1'!A:R の列構成が変わらない前提)
        const satBIndex = 1, satCIndex = 2, satDIndex = 3, satEIndex = 4, satFIndex = 5, satGIndex = 6, satHIndex = 7, feedbackIIndex = 8, feedbackJIndex = 9, scoreKIndex = 10, reasonLIndex = 11, feedbackMIndex = 12, recommendationNIndex = 13, ageOIndex = 14, childrenPIndex = 15, incomeQIndex = 16, postalCodeRIndex = 17;

        // 2. データをループして集計 (変更なし)
        clinicDataRows.forEach((row) => {
            // L列 (NPS理由)
            const score = row[scoreKIndex], reason = row[reasonLIndex]; if (reason != null && String(reason).trim() !== '') { const scoreNum = parseInt(score, 10); if (!isNaN(scoreNum)) allNpsReasons.push({ score: scoreNum, reason: String(reason).trim() }); }
            // I, J, M列 (フィードバック)
            const feedbackI = row[feedbackIIndex]; if (feedbackI != null && String(feedbackI).trim() !== '') allFeedbacks_I.push(String(feedbackI).trim());
            const feedbackJ = row[feedbackJIndex]; if (feedbackJ != null && String(feedbackJ).trim() !== '') allFeedbacks_J.push(String(feedbackJ).trim());
            const feedbackM = row[feedbackMIndex]; if (feedbackM != null && String(feedbackM).trim() !== '') allFeedbacks_M.push(String(feedbackM).trim());
            // B-H列 (満足度)
            const satB = row[satBIndex]; if (satB != null && satisfactionKeys.includes(String(satB))) satisfactionCounts_B[String(satB)]++;
            const satC = row[satCIndex]; if (satC != null && satisfactionKeys.includes(String(satC))) satisfactionCounts_C[String(satC)]++;
            const satD = row[satDIndex]; if (satD != null && satisfactionKeys.includes(String(satD))) satisfactionCounts_D[String(satD)]++;
            const satE = row[satEIndex]; if (satE != null && satisfactionKeys.includes(String(satE))) satisfactionCounts_E[String(satE)]++;
            const satF = row[satFIndex]; if (satF != null && satisfactionKeys.includes(String(satF))) satisfactionCounts_F[String(satF)]++;
            const satG = row[satGIndex]; if (satG != null && satisfactionKeys.includes(String(satG))) satisfactionCounts_G[String(satG)]++;
            const satH = row[satHIndex]; if (satH != null && satisfactionKeys.includes(String(satH))) satisfactionCounts_H[String(satH)]++;
            // P, O, Q, R列
            const childrenP = row[childrenPIndex]; if (childrenP != null && childrenKeys.includes(String(childrenP))) childrenCounts_P[String(childrenP)]++;
            const ageO = row[ageOIndex]; if (ageO != null && ageKeys.includes(String(ageO))) ageCounts_O[String(ageO)]++;
            const income = row[incomeQIndex]; if (typeof income === 'number' && income >= 1 && income <= 10 && !isNaN(income)) incomeCounts[income]++;
            const postalCodeRaw = row[postalCodeRIndex];
            if (postalCodeRaw) {
                const postalCode = String(postalCodeRaw).replace(/-/g, '').trim();
                if (/^\d{7}$/.test(postalCode)) {
                    postalCodeCounts[postalCode] = (postalCodeCounts[postalCode] || 0) + 1;
                }
            }
            // K列 (NPSスコア)
            const npsScore = row[scoreKIndex];
            if (npsScore != null && npsScore >= 0 && npsScore <= 10) {
                const scoreNum = parseInt(npsScore, 10);
                if (!isNaN(scoreNum)) npsScoreCounts[scoreNum]++;
            }
            // N列 (おすすめ理由)
            const recommendation = row[recommendationNIndex];
            if (recommendation != null) {
                const recText = String(recommendation).trim();
                if (recommendationKeys.includes(recText)) {
                    recommendationCounts[recText]++;
                } else if (recText !== '') {
                    recommendationOthers.push(recText);
                }
            }
        });
        
        console.log(`[googleSheetsService-AGG] Aggregation complete for "${sheetName}".`);

        // 3. レポートオブジェクトを構築して返す (変更なし)
        const aggregationData = {
            allNpsReasons, allFeedbacks_I, allFeedbacks_J, allFeedbacks_M,
            satisfactionCounts_B, satisfactionCounts_C, satisfactionCounts_D, satisfactionCounts_E, satisfactionCounts_F, satisfactionCounts_G, satisfactionCounts_H,
            childrenCounts_P, ageCounts_O, incomeCounts, postalCodeCounts,
            npsScoreCounts, recommendationCounts, recommendationOthers,
            satisfactionKeys, ageKeys, childrenKeys, recommendationKeys,
            createChartData
        };

        return buildReportDataObject(aggregationData);

    } catch (e) {
        console.error(`[googleSheetsService-AGG] Error aggregating data from "${sheetName}": ${e.toString()}`, e.stack);
        return buildReportDataObject(null);
    }
};

// =================================================================
// === ▼▼▼ 関数 (4/X) (構築 - 変更なし) ▼▼▼ ===
// =================================================================
function buildReportDataObject(data) {
    // (既存のコードのまま - 変更なし)
    if (!data) {
        const emptyChart = [['カテゴリ', '件数']];
        return {
            npsData: { totalCount: 0, results: {}, rawText: [] },
            feedbackData: { i_column: { totalCount: 0, results: [] }, j_column: { totalCount: 0, results: [] }, m_column: { totalCount: 0, results: [] } },
            satisfactionData: { b_column: { results: emptyChart }, c_column: { results: emptyChart }, d_column: { results: emptyChart }, e_column: { results: emptyChart }, f_column: { results: emptyChart }, g_column: { results: emptyChart }, h_column: { results: emptyChart } },
            ageData: { results: emptyChart },
            childrenCountData: { results: emptyChart },
            incomeData: { results: [['評価', '割合', { role: 'annotation' }]], totalCount: 0 },
            postalCodeData: { counts: {} },
            npsScoreData: { counts: {}, totalCount: 0 },
            recommendationData: { fixedCounts: {}, otherList: [], fixedKeys: [] }
        };
    }
    const {
        allNpsReasons, allFeedbacks_I, allFeedbacks_J, allFeedbacks_M,
        satisfactionCounts_B, satisfactionCounts_C, satisfactionCounts_D, satisfactionCounts_E, satisfactionCounts_F, satisfactionCounts_G, satisfactionCounts_H,
        childrenCounts_P, ageCounts_O, incomeCounts, postalCodeCounts,
        npsScoreCounts, recommendationCounts, recommendationOthers,
        satisfactionKeys, ageKeys, childrenKeys, recommendationKeys,
        createChartData
    } = data;
    const groupedByScore = allNpsReasons.reduce((acc, item) => { if (typeof item.score === 'number' && !isNaN(item.score)) { if (!acc[item.score]) acc[item.score] = []; acc[item.score].push(item.reason); } return acc; }, {});
    const incomeChartData = [['評価', '割合', { role: 'annotation' }]];
    const totalIncomeCount = Object.values(incomeCounts).reduce((a, b) => a + b, 0);
    if (totalIncomeCount > 0) { for (let i = 1; i <= 10; i++) { const count = incomeCounts[i] || 0; const percentage = (count / totalIncomeCount) * 100; incomeChartData.push([String(i), percentage, `${Math.round(percentage)}%`]); } }
    return {
        npsData: { totalCount: allNpsReasons.length, results: groupedByScore, rawText: allNpsReasons.map(r => r.reason) },
        feedbackData: { i_column: { totalCount: allFeedbacks_I.length, results: allFeedbacks_I }, j_column: { totalCount: allFeedbacks_J.length, results: allFeedbacks_J }, m_column: { totalCount: allFeedbacks_M.length, results: allFeedbacks_M } },
        satisfactionData: { b_column: { results: createChartData(satisfactionCounts_B, satisfactionKeys) }, c_column: { results: createChartData(satisfactionCounts_C, satisfactionKeys) }, d_column: { results: createChartData(satisfactionCounts_D, satisfactionKeys) }, e_column: { results: createChartData(satisfactionCounts_E, satisfactionKeys) }, f_column: { results: createChartData(satisfactionCounts_F, satisfactionKeys) }, g_column: { results: createChartData(satisfactionCounts_G, satisfactionKeys) }, h_column: { results: createChartData(satisfactionCounts_H, satisfactionKeys) } },
        ageData: { results: createChartData(ageCounts_O, ageKeys) },
        childrenCountData: { results: createChartData(childrenCounts_P, childrenKeys) },
        incomeData: { results: incomeChartData, totalCount: totalIncomeCount },
        postalCodeData: { counts: postalCodeCounts },
        npsScoreData: { counts: npsScoreCounts, totalCount: Object.values(npsScoreCounts).reduce((a, b) => a + b, 0) },
        recommendationData: { fixedCounts: recommendationCounts, otherList: recommendationOthers, fixedKeys: recommendationKeys }
    };
}


// =================================================================
// === ▼▼▼ 関数 (5/X) (AI保存 - 変更なし) ▼▼▼ ===
// =================================================================
exports.saveAIAnalysisToSheet = async (centralSheetId, clinicName, analysisType, jsonData) => {
    // (既存のコードのまま - 変更なし)
    if (!jsonData) throw new Error('AI analysis JSON data is missing.');
    const baseSheetName = `${clinicName}-AI分析-${analysisType}`;
    try {
        const analysisSheetName = `${baseSheetName}-分析`;
        const analysisContent = (jsonData.analysis && jsonData.analysis.themes) ? jsonData.analysis.themes.map(t => `【${t.title}】\n${t.summary}`).join('\n\n---\n\n') : '分析データがありません。';
        await findOrCreateSheet(centralSheetId, analysisSheetName);
        await writeToCell(centralSheetId, analysisSheetName, 'A1', analysisContent);
        console.log(`[googleSheetsService-AI] Saved Analysis to: "${analysisSheetName}"`);
        const suggestionSheetName = `${baseSheetName}-改善案`;
        const suggestionContent = (jsonData.suggestions && jsonData.suggestions.items) ? jsonData.suggestions.items.map(i => `【${i.themeTitle}】\n${i.suggestion}`).join('\n\n---\n\n') : '改善提案データがありません。';
        await findOrCreateSheet(centralSheetId, suggestionSheetName);
        await writeToCell(centralSheetId, suggestionSheetName, 'A1', suggestionContent);
        console.log(`[googleSheetsService-AI] Saved Suggestions to: "${suggestionSheetName}"`);
        const overallSheetName = `${baseSheetName}-総評`;
        const overallContent = (jsonData.overall && jsonData.overall.summary) ? jsonData.overall.summary : '総評データがありません。';
        await findOrCreateSheet(centralSheetId, overallSheetName);
        await writeToCell(centralSheetId, overallSheetName, 'A1', overallContent);
        console.log(`[googleSheetsService-AI] Saved Overall to: "${overallSheetName}"`);
        return true;
    } catch (err) {
        console.error(`[googleSheetsService-AI] Failed to save AI analysis to sheets: ${err.message}`, err);
        throw new Error(`AI分析結果のシート保存に失敗しました: ${err.message}`);
    }
};

// =================================================================
// === ▼▼▼ 関数 (6/X) (AI読込 - 変更なし) ▼▼▼ ===
// =================================================================
exports.getAIAnalysisFromSheet = async (centralSheetId, clinicName, analysisType) => {
    // (既存のコードのまま - 変更なし)
    const baseSheetName = `${clinicName}-AI分析-${analysisType}`;
    try {
        const analysisSheetName = `${baseSheetName}-分析`;
        const suggestionSheetName = `${baseSheetName}-改善案`;
        const overallSheetName = `${baseSheetName}-総評`;
        const [analysisRes, suggestionRes, overallRes] = await Promise.all([
            readCell(centralSheetId, analysisSheetName, 'A1'),
            readCell(centralSheetId, suggestionSheetName, 'A1'),
            readCell(centralSheetId, overallSheetName, 'A1')
        ]);
        console.log(`[googleSheetsService-AI] Read AI analysis from 3 sheets for: "${baseSheetName}"`);
        return {
            analysis: analysisRes || '（分析データがありません）',
            suggestions: suggestionRes || '（改善提案データがありません）',
            overall: overallRes || '（総評データがありません）'
        };
    } catch (err) {
        console.error(`[googleSheetsService-AI] Failed to read AI analysis from sheets: ${err.message}`, err);
        throw new Error(`AI分析結果のシート読み込みに失敗しました: ${err.message}`);
    }
};

// =================================================================
// === ▼▼▼ 関数 (7/X) (AI更新 - 変更なし) ▼▼▼ ===
// =================================================================
exports.updateAIAnalysisInSheet = async (centralSheetId, sheetName, content) => {
    // (既存のコードのまま - 変更なし)
    try {
        await writeToCell(centralSheetId, sheetName, 'A1', content);
        console.log(`[googleSheetsService-AI] Updated content in: "${sheetName}"`);
        return true;
    } catch (err) {
        console.error(`[googleSheetsService-AI] Failed to update AI analysis sheet: ${err.message}`, err);
        throw new Error(`AI分析結果の更新に失敗しました: ${err.message}`);
    }
};

// =================================================================
// === ▼▼▼ 新規関数 (8/X) (要求 #2 市区町村シート保存) ▼▼▼ ===
// =================================================================
exports.saveMunicipalityData = async (centralSheetId, sheetName, tableData) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService-Muni] Saving municipality data to sheet: "${sheetName}"`);
    
    try {
        // 1. ヘッダー行を作成
        const header = ['都道府県', '市区町村', '件数', '割合'];
        
        // 2. データ行を作成
        const dataRows = tableData.map(row => [
            row.prefecture,
            row.municipality,
            row.count,
            row.percentage // 割合 (数値)
        ]);

        // 3. シートを作成（またはクリア）
        await findOrCreateSheet(centralSheetId, sheetName);
        await clearSheet(centralSheetId, sheetName);
        
        // 4. ヘッダー + データを書き込み
        await writeData(centralSheetId, sheetName, [header, ...dataRows]);
        
        // 5. [装飾] 割合(D列)をパーセンテージ表示に設定
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: centralSheetId,
            resource: {
                requests: [
                    {
                        repeatCell: {
                            range: {
                                sheetId: (await getSheetId(centralSheetId, sheetName)),
                                startRowIndex: 1, // 2行目から
                                startColumnIndex: 3, // D列
                                endColumnIndex: 4
                            },
                            cell: {
                                userEnteredFormat: {
                                    numberFormat: {
                                        type: 'PERCENT',
                                        pattern: '0.00%'
                                    }
                                }
                            },
                            fields: 'userEnteredFormat.numberFormat'
                        }
                    }
                ]
            }
        });
        
        console.log(`[googleSheetsService-Muni] Successfully saved and formatted data for "${sheetName}"`);

    } catch (err) {
        console.error(`[googleSheetsService-Muni] Error saving municipality data: ${err.message}`, err);
        throw new Error(`市区町村データのシート保存に失敗しました: ${err.message}`);
    }
};

// =================================================================
// === ▼▼▼ 新規関数 (9/X) (要求 #3 市区町村シート読込) ▼▼▼ ===
// =================================================================
exports.readMunicipalityData = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService-Muni] Reading municipality data from sheet: "${sheetName}"`);

    try {
        // A:D 列を読み込む (ヘッダー含む)
        const range = `'${sheetName}'!A:D`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
            valueRenderOption: 'UNFORMATTED_VALUE' // 割合を 0.123 の形式で取得
        });

        const rows = response.data.values;

        if (!rows || rows.length < 2) {
            console.log(`[googleSheetsService-Muni] No data found in "${sheetName}".`);
            return null; // データが存在しない (まだ生成されていない)
        }

        rows.shift(); // ヘッダーを捨てる

        // フロントエンド用のテーブル形式に変換
        const tableData = rows.map(row => ({
            prefecture: row[0],
            municipality: row[1],
            count: parseFloat(row[2]) || 0,
            percentage: (parseFloat(row[3]) || 0) * 100 // 0.123 -> 12.3
        }));
        
        console.log(`[googleSheetsService-Muni] Successfully read ${tableData.length} rows.`);
        return tableData;

    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            console.warn(`[googleSheetsService-Muni] Sheet "${sheetName}" not found. Returning null.`);
            return null; // シートが存在しない
        }
        console.error(`[googleSheetsService-Muni] Error reading municipality data: ${e.message}`, e);
        throw new Error(`市区町村データのシート読み込みに失敗しました: ${e.message}`);
    }
};


// =================================================================
// === ▼▼▼ 新規関数 (10/X) (要求 #6 転記状況確認) ▼▼▼ ===
// =================================================================
exports.getSheetTitles = async (spreadsheetId) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        const metadata = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            fields: 'sheets(properties(title))'
        });
        
        const titles = metadata.data.sheets.map(sheet => sheet.properties.title);
        console.log(`[getSheetTitles] Found ${titles.length} sheets:`, titles.join(', '));
        return titles;
        
    } catch (e) {
        console.error(`[getSheetTitles] Error getting sheet titles: ${e.message}`);
        throw new Error(`シート一覧の取得に失敗しました: ${e.message}`);
    }
};


// =================================================================
// === ヘルパー関数群 (変更あり) ===
// =================================================================

/**
 * [Helper] スプレッドシートに新しいシート（タブ）を追加する
 */
async function addSheet(spreadsheetId, title) {
    try {
        const request = {
            spreadsheetId: spreadsheetId,
            resource: {
                requests: [{ addSheet: { properties: { title: title } } }]
            }
        };
        const response = await sheets.spreadsheets.batchUpdate(request);
        // ▼▼▼ [変更点] 新しいシートのIDを返すようにする ▼▼▼
        const newSheetId = response.data.replies[0].addSheet.properties.sheetId;
        console.log(`[Helper] Added sheet "${title}" (ID: ${newSheetId})`);
        return newSheetId;
    } catch (e) {
        if (e.message.includes('already exists')) {
            console.warn(`[Helper] Sheet "${title}" already exists.`);
            // 既存のシートIDを検索して返す
            return await getSheetId(spreadsheetId, title);
        } else {
            console.error(`[Helper] Error adding sheet "${title}": ${e.message}`);
            throw e;
        }
    }
}

/**
 * [Helper] シート名からシートIDを取得する
 */
async function getSheetId(spreadsheetId, title) {
    try {
        const metadata = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            fields: 'sheets(properties(sheetId,title))'
        });
        const sheet = metadata.data.sheets.find(s => s.properties.title === title);
        return sheet ? sheet.properties.sheetId : null;
    } catch (e) {
        console.error(`[Helper] Error getting sheet ID for "${title}": ${e.message}`);
        return null;
    }
}


/**
 * [Helper] 指定したシート（タブ）を見つけて作成する（存在すれば何もしない）
 * @returns {number} sheetId
 */
async function findOrCreateSheet(spreadsheetId, title) {
    try {
        const existingSheetId = await getSheetId(spreadsheetId, title);
        
        if (existingSheetId !== null) {
            console.log(`[Helper] Sheet "${title}" already exists (ID: ${existingSheetId}).`);
            return existingSheetId;
        }
        
        // 存在しない場合のみ作成
        console.log(`[Helper] Sheet "${title}" not found. Creating...`);
        const newSheetId = await addSheet(spreadsheetId, title);
        return newSheetId;
        
    } catch (e) {
        console.error(`[Helper] Error in findOrCreateSheet for "${title}": ${e.message}`);
        throw e;
    }
}


/**
 * [Helper] シート（タブ）の全データをクリアする
 */
async function clearSheet(spreadsheetId, range) {
    try {
        await sheets.spreadsheets.values.clear({
            spreadsheetId: spreadsheetId,
            range: range, // シート名全体
        });
    } catch (e) {
        console.error(`[Helper] Error clearing sheet "${range}": ${e.message}`);
        throw e;
    }
}

/**
 * [Helper] シートにデータを書き込む (上書き または 追記)
 */
async function writeData(spreadsheetId, range, values, append = false) {
    if (!values || values.length === 0) {
        console.warn(`[Helper] No data to write for sheet "${range}".`);
        return;
    }
    
    try {
        if (append) {
            // 追記モード
            await sheets.spreadsheets.values.append({
                spreadsheetId: spreadsheetId,
                range: range,
                valueInputOption: 'USER_ENTERED',
                insertDataOption: 'INSERT_ROWS',
                resource: {
                    values: values
                }
            });
        } else {
            // 上書きモード
            await sheets.spreadsheets.values.update({
                spreadsheetId: spreadsheetId,
                range: range,
                valueInputOption: 'USER_ENTERED',
                resource: {
                    values: values
                }
            });
        }
    } catch (e) {
        console.error(`[Helper] Error writing data to sheet "${range}" (Append: ${append}): ${e.message}`);
        throw e;
    }
}

/**
 * [Helper] 指定したシートの特定のセルに値を書き込む (A1など)
 */
async function writeToCell(spreadsheetId, sheetName, cell, content) {
    try {
        const range = `'${sheetName}'!${cell}`;
        await sheets.spreadsheets.values.update({
            spreadsheetId: spreadsheetId,
            range: range,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: [[content]]
            }
        });
    } catch (e) {
        console.error(`[Helper] Error writing to cell "${sheetName}!${cell}": ${e.message}`);
        throw e;
    }
}

/**
 * [Helper] 指定したシートの特定のセルから値を読み取る
 */
async function readCell(spreadsheetId, sheetName, cell) {
    try {
        const range = `'${sheetName}'!${cell}`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: spreadsheetId,
            range: range,
            valueRenderOption: 'FORMATTED_VALUE'
        });
        
        const value = response.data.values ? response.data.values[0][0] : null;
        return value;
    } catch (e) {
         if (e.message.includes('Unable to parse range')) {
             console.warn(`[Helper] Cell not found (or sheet missing) for "${sheetName}!${cell}". Returning null.`);
             return null;
         }
        console.error(`[Helper] Error reading cell "${sheetName}!${cell}": ${e.message}`);
        throw e;
    }
}

// ★ 既存の getSpreadsheetIdFromUrl を googleSheetsService の末尾にもエクスポート
exports.getSpreadsheetIdFromUrl = getSpreadsheetIdFromUrl;
