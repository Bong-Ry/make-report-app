// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/services/googleSheets.js

const { google } = require('googleapis');
const { getSpreadsheetIdFromUrl, getAiAnalysisKeys } = require('../utils/helpers');

// --- (変更なし) 認証・初期化 ---
const GAS_SHEET_FINDER_URL = 'https://script.google.com/macros/s/AKfycbzn4rNw6NttPPmcJBpSKJifK8-Mb1CatsGhqvYF5G6BIAf6bOUuNS_E72drg0tH9re-qQ/exec';
const KEYFILEPATH = '/etc/secrets/credentials.json';
const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
];
const MASTER_FOLDER_ID = '1_pJQKl5-RRi6h-U3EEooGmPkTrkF1Vbj';

let sheets;
let drive;

try {
    const auth = new google.auth.GoogleAuth({ keyFile: KEYFILEPATH, scopes: SCOPES });
    sheets = google.sheets({ version: 'v4', auth });
    drive = google.drive({ version: 'v3', auth });
    console.log('Google Sheets, Drive API clients initialized successfully.');
} catch (err) {
    console.error('Failed to initialize Google API clients:', err);
}

exports.sheets = sheets;
// --- (変更なし) 認証・初期化 終わり ---

// --- (変更なし) マスターシート関数 (2つ) ---
exports.getMasterClinicList = async () => {
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

exports.getMasterClinicUrls = async () => {
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
// --- (変更なし) マスターシート関数 終わり ---

// --- (変更なし) GAS呼び出し (findOrCreateCentralSheet) ---
exports.findOrCreateCentralSheet = async (periodText) => {
    const fileName = periodText;
    console.log(`[googleSheetsService] Finding or creating central sheet via GAS: "${fileName}"`);
    try {
        const response = await fetch(GAS_SHEET_FINDER_URL, {
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
// --- (変更なし) GAS呼び出し 終わり ---

// --- (変更なし) データ転記 (fetchAndAggregateReportData) ---
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
            // [メモ] ユーザー設定により "フォームの回答 1" が参照されます
            const range = "'フォームの回答 1'!A:R"; 
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
                const clinicSheetTitle = clinicName; 
                await findOrCreateSheet(centralSheetId, clinicSheetTitle);
                
                await clearSheet(centralSheetId, clinicSheetTitle);
                await writeData(centralSheetId, clinicSheetTitle, filteredRows);
                console.log(`[googleSheetsService-ETL] Wrote ${filteredRows.length} rows to sheet: "${clinicSheetTitle}" (HEADERLESS)`);

                await findOrCreateSheet(centralSheetId, '全体');
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
            continue; 
        }
    }
    
    return processedClinics;
};
// --- (変更なし) データ転記 終わり ---

// --- (変更なし) データ集計 (getReportDataForCharts) ---
exports.getReportDataForCharts = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService-AGG] Aggregating data from Sheet ID: ${centralSheetId}, Tab: "${sheetName}"`);
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
        const range = `'${sheetName}'!A:R`;
        const clinicDataResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
            dateTimeRenderOption: 'SERIAL_NUMBER', 
            valueRenderOption: 'UNFORMATTED_VALUE'
        });
        const clinicDataRows = clinicDataResponse.data.values;
        if (!clinicDataRows || clinicDataRows.length < 1) { 
            console.log(`[googleSheetsService-AGG] No data found in "${sheetName}".`);
            return buildReportDataObject(null); 
        }
        const satBIndex = 1, satCIndex = 2, satDIndex = 3, satEIndex = 4, satFIndex = 5, satGIndex = 6, satHIndex = 7, feedbackIIndex = 8, feedbackJIndex = 9, scoreKIndex = 10, reasonLIndex = 11, feedbackMIndex = 12, recommendationNIndex = 13, ageOIndex = 14, childrenPIndex = 15, incomeQIndex = 16, postalCodeRIndex = 17;
        clinicDataRows.forEach((row) => {
            const score = row[scoreKIndex], reason = row[reasonLIndex]; if (reason != null && String(reason).trim() !== '') { const scoreNum = parseInt(score, 10); if (!isNaN(scoreNum)) allNpsReasons.push({ score: scoreNum, reason: String(reason).trim() }); }
            const feedbackI = row[feedbackIIndex]; if (feedbackI != null && String(feedbackI).trim() !== '') allFeedbacks_I.push(String(feedbackI).trim());
            const feedbackJ = row[feedbackJIndex]; if (feedbackJ != null && String(feedbackJ).trim() !== '') allFeedbacks_J.push(String(feedbackJ).trim());
            const feedbackM = row[feedbackMIndex]; if (feedbackM != null && String(feedbackM).trim() !== '') allFeedbacks_M.push(String(feedbackM).trim());
            const satB = row[satBIndex]; if (satB != null && satisfactionKeys.includes(String(satB))) satisfactionCounts_B[String(satB)]++;
            const satC = row[satCIndex]; if (satC != null && satisfactionKeys.includes(String(satC))) satisfactionCounts_C[String(satC)]++;
            const satD = row[satDIndex]; if (satD != null && satisfactionKeys.includes(String(satD))) satisfactionCounts_D[String(satD)]++;
            const satE = row[satEIndex]; if (satE != null && satisfactionKeys.includes(String(satE))) satisfactionCounts_E[String(satE)]++;
            const satF = row[satFIndex]; if (satF != null && satisfactionKeys.includes(String(satF))) satisfactionCounts_F[String(satF)]++;
            const satG = row[satGIndex]; if (satG != null && satisfactionKeys.includes(String(satG))) satisfactionCounts_G[String(satG)]++;
            const satH = row[satHIndex]; if (satH != null && satisfactionKeys.includes(String(satH))) satisfactionCounts_H[String(satH)]++;
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
            const npsScore = row[scoreKIndex];
            if (npsScore != null && npsScore >= 0 && npsScore <= 10) {
                const scoreNum = parseInt(npsScore, 10);
                if (!isNaN(scoreNum)) npsScoreCounts[scoreNum]++;
            }
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
// --- (変更なし) データ集計 終わり ---

// --- (変更なし) データ構築 (buildReportDataObject) ---
function buildReportDataObject(data) {
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
// --- (変更なし) データ構築 終わり ---


// --- (変更なし) コメントシート名 取得ヘルパー (ローカル) ---
function getCommentSheetName(clinicName, type) {
    switch(type) {
        case 'L_10': return `${clinicName}_NPS10`;
        case 'L_9': return `${clinicName}_NPS9`;
        case 'L_8': return `${clinicName}_NPS8`;
        case 'L_7': return `${clinicName}_NPS7`;
        case 'L_6_under': return `${clinicName}_NPS6以下`;
        case 'I': return `${clinicName}_よかった点悪かった点`;
        case 'J': return `${clinicName}_印象スタッフ`;
        case 'M': return `${clinicName}_お産意見`;
        default:
            throw new Error(`[helpers] 無効なコメントシートタイプです: ${type}`);
    }
};
exports.getCommentSheetName = getCommentSheetName;
// --- (変更なし) コメントシート名 取得ヘルパー 終わり ---

// --- (変更なし) コメントデータ保存・読み込み・更新 (3関数) ---
const ROWS_PER_COLUMN = 20;

function formatCommentsToColumns(comments) {
    const columns = [];
    for (let i = 0; i < comments.length; i += ROWS_PER_COLUMN) {
        columns.push(comments.slice(i, i + ROWS_PER_COLUMN));
    }
    return columns;
}

function transpose(columns) {
    if (!columns || columns.length === 0) return [];
    const maxRows = Math.max(...columns.map(col => col.length));
    const rows = [];
    for (let i = 0; i < maxRows; i++) {
        const row = [];
        for (let j = 0; j < columns.length; j++) {
            row.push(columns[j][i] || '');
        }
        rows.push(row);
    }
    return rows;
}

exports.saveCommentsToSheet = async (centralSheetId, sheetName, comments) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    if (!comments || comments.length === 0) {
        console.log(`[googleSheetsService-Comments] No comments to save for "${sheetName}". Skipping.`);
        try {
            await findOrCreateSheet(centralSheetId, sheetName);
            await clearSheet(centralSheetId, sheetName);
        } catch (e) { /* (シートが存在しない場合のエラーは無視) */ }
        return;
    }

    console.log(`[googleSheetsService-Comments] Saving ${comments.length} comments to "${sheetName}"...`);
    
    const columnsData = formatCommentsToColumns(comments);
    const rowsData = transpose(columnsData);

    try {
        await findOrCreateSheet(centralSheetId, sheetName);
        await clearSheet(centralSheetId, sheetName);
        await writeData(centralSheetId, sheetName, rowsData);
        console.log(`[googleSheetsService-Comments] Saved comments to "${sheetName}".`);
    } catch (e) {
        console.error(`[googleSheetsService-Comments] Failed to save comments to "${sheetName}": ${e.message}`, e);
        throw e;
    }
};

exports.readCommentsBySheet = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    const range = `'${sheetName}'!A:Z`;
    console.log(`[googleSheetsService-Comments] Reading comments from "${sheetName}" (Range: ${range})`);

    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
            valueRenderOption: 'FORMATTED_VALUE',
            majorDimension: 'COLUMNS'
        });

        const columns = response.data.values;

        if (!columns || columns.length === 0) { 
            console.log(`[googleSheetsService-Comments] No data found in "${sheetName}".`);
            return [];
        }

        return columns;

    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            console.warn(`[googleSheetsService-Comments] Sheet "${sheetName}" not found. Returning empty data.`);
            return [];
        }
        console.error(`[googleSheetsService-Comments] Error reading comment data: ${e.message}`, e);
        throw new Error(`コメントシート(${sheetName})の読み込みに失敗しました: ${e.message}`);
    }
};

exports.updateCommentSheetCell = async (centralSheetId, sheetName, cell, value) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    const range = `'${sheetName}'!${cell}`;
    
    console.log(`[googleSheetsService-Comments] Updating cell "${range}" with new value...`);
    
    try {
        await writeData(centralSheetId, range, [[value]], false);
        console.log(`[googleSheetsService-Comments] Cell update successful.`);
    } catch (e) {
        console.error(`[googleSheetsService-Comments] Error updating cell "${range}": ${e.message}`, e);
        throw new Error(`コメントシート(${sheetName})のセル(${range})更新に失敗しました: ${e.message}`);
    }
};
// --- (変更なし) コメントI/O 終わり ---


// --- (変更なし) シート名一覧取得 (getSheetTitles) ---
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
// --- (変更なし) シート名一覧取得 終わり ---

// --- (変更なし) 分析タブI/O (saveTableToSheet) ---
exports.saveTableToSheet = async (centralSheetId, sheetName, dataRows) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService] Saving Table to Sheet: "${sheetName}"`);
    
    try {
        await findOrCreateSheet(centralSheetId, sheetName);
        await clearSheet(centralSheetId, sheetName);
        await writeData(centralSheetId, sheetName, dataRows);
        
        const sheetId = await getSheetId(centralSheetId, sheetName);
        if (sheetId) {
             await sheets.spreadsheets.batchUpdate({
                spreadsheetId: centralSheetId,
                resource: { requests: [
                    { autoResizeDimensions: {
                        dimensions: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 4 }
                    }}
                ]}
            });
        }
        
        console.log(`[googleSheetsService] Successfully saved Table for "${sheetName}"`);

    } catch (err) {
        console.error(`[googleSheetsService] Error saving Table data: ${err.message}`, err);
        throw new Error(`分析テーブルのシート保存に失敗しました: ${err.message}`);
    }
};

// --- (変更なし) 分析タブI/O (saveAiAnalysisData) ---
// (A1:B1にヘッダー、A2:B16にキーと値を書き込む)
exports.saveAiAnalysisData = async (centralSheetId, sheetName, aiDataMap) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService] Saving AI Key-Value Data to Sheet: "${sheetName}"`);
    
    try {
        await findOrCreateSheet(centralSheetId, sheetName);
        await clearSheet(centralSheetId, sheetName);
        
        const allKeys = getAiAnalysisKeys(); 
        const dataRows = allKeys.map(key => {
            const value = aiDataMap.get(key) || '（データなし）';
            return [key, value]; // [A列, B列]
        });
        
        const header = ['項目キー', '分析文章データ'];
        const finalData = [header, ...dataRows];

        // A1から書き込む (A1:B1, A2:B2, ..., A16:B16)
        await writeData(centralSheetId, `'${sheetName}'!A1`, finalData);
        
        const sheetId = await getSheetId(centralSheetId, sheetName);
        if (sheetId) {
             await sheets.spreadsheets.batchUpdate({
                spreadsheetId: centralSheetId,
                resource: { requests: [
                    { autoResizeDimensions: {
                        dimensions: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }
                    }},
                    { updateDimensionProperties: {
                        range: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 },
                        properties: { pixelSize: 800 },
                        fields: 'pixelSize'
                    }}
                ]}
            });
        }
        
        console.log(`[googleSheetsService] Successfully saved AI Key-Value Data for "${sheetName}"`);

    } catch (err) {
        console.error(`[googleSheetsService] Error saving AI Key-Value data: ${err.message}`, err);
        throw new Error(`AI分析(Key-Value)のシート保存に失敗しました: ${err.message}`);
    }
};

// =================================================================
// === ▼▼▼ [修正] AI詳細分析 (読み出し) B2:B16固定参照 ▼▼▼ ===
// =================================================================
exports.readAiAnalysisData = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    // 1. 15個のキーを定義順に取得
    const allKeys = getAiAnalysisKeys();
    const aiDataMap = new Map();

    console.log(`[googleSheetsService] Reading AI Fixed-Cell Data from Sheet: "${sheetName}" (Range: B2:B16)`);

    try {
        // 2. B2:B16 の範囲を固定で読み込む
        const range = `'${sheetName}'!B2:B16`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
            valueRenderOption: 'FORMATTED_VALUE'
        });

        const rows = response.data.values; // 例: [['B2の値'], ['B3の値'], ..., ['B16の値']]

        // 3. 取得した値(最大15行)を、15個のキーに順番にマッピング
        allKeys.forEach((key, index) => {
            // rows[index] が存在し、かつ rows[index][0] (B列の値) が存在する場合
            const value = (rows && rows[index] && rows[index][0] != null)
                ? rows[index][0]
                : '（データがありません）'; // 該当セルが空か、範囲外の場合
                
            aiDataMap.set(key, value);
        });
        
        console.log(`[googleSheetsService] Successfully read ${allKeys.length} AI values from fixed cells.`);
        return aiDataMap;

    } catch (e) {
        // (シートが存在しない場合など)
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            console.warn(`[googleSheetsService] Sheet "${sheetName}" not found. Returning empty map with defaults.`);
            // キー検索をしないため、全キーにデフォルト値を設定して返す
            allKeys.forEach(key => {
                aiDataMap.set(key, '（データがありません）');
            });
            return aiDataMap;
        }
        console.error(`[googleSheetsService] Error reading AI Key-Value data: ${e.message}`, e);
        throw new Error(`AI分析(Key-Value)のシート読み込みに失敗しました: ${e.message}`);
    }
};
// =================================================================
// === ▲▲▲ 修正完了 ▲▲▲ ===
// =================================================================

// --- (変更なし) 管理マーカーI/O (3関数) ---
const MANAGEMENT_SHEET_NAME = '管理'; 

exports.writeInitialMarker = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService-Marker] Writing Initial Marker for: "${clinicName}"`);
    try {
        await findOrCreateSheet(centralSheetId, MANAGEMENT_SHEET_NAME);
        await writeData(centralSheetId, MANAGEMENT_SHEET_NAME, [[clinicName]], true); // append = true
    } catch (e) {
        console.error(`[googleSheetsService-Marker] Error writing initial marker: ${e.message}`, e);
    }
};

exports.writeCompletionMarker = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService-Marker] Writing Completion Marker for: "${clinicName}"`);
    try {
        await findOrCreateSheet(centralSheetId, MANAGEMENT_SHEET_NAME);
        
        const range = `'${MANAGEMENT_SHEET_NAME}'!A:A`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
        });
        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            throw new Error('「管理」シートが空です。');
        }

        const rowIndex = rows.findIndex(row => row[0] === clinicName);
        
        if (rowIndex === -1) {
            console.warn(`[googleSheetsService-Marker] Clinic name "${clinicName}" not found in management sheet A-column.`);
            await writeData(centralSheetId, MANAGEMENT_SHEET_NAME, [[clinicName, 'Complete']], true); // append
        } else {
            const cellToUpdate = `'${MANAGEMENT_SHEET_NAME}'!B${rowIndex + 1}`;
            await writeData(centralSheetId, cellToUpdate, [['Complete']], false); // (上書き)
        }

    } catch (e) {
        console.error(`[googleSheetsService-Marker] Error writing completion marker: ${e.message}`, e);
    }
};

exports.readCompletionStatusMap = async (centralSheetId) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService-Marker] Reading Completion Status Map...`);
    
    const statusMap = {};
    
    try {
        await findOrCreateSheet(centralSheetId, MANAGEMENT_SHEET_NAME);
        
        const range = `'${MANAGEMENT_SHEET_NAME}'!A:B`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            return statusMap;
        }
        
        rows.forEach(row => {
            const clinicName = row[0];
            const status = row[1];
            if (clinicName) {
                statusMap[clinicName] = (status === 'Complete');
            }
        });
        
        return statusMap;

    } catch (e) {
        console.error(`[googleSheetsService-Marker] Error reading status map: ${e.message}`, e);
        return statusMap;
    }
};
// --- (変更なし) 管理マーカーI/O 終わり ---

// データ行数を取得する関数
exports.getSheetRowCounts = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');

    try {
        const results = {};

        // 「全体」シートのデータ行数（ヘッダー行を除く）
        const overallResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: '全体!A:A',
        });
        results.overallCount = (overallResponse.data.values?.length || 1) - 1; // ヘッダーを除く

        // 「管理」シートのデータ行数（ヘッダー行を除く）
        const managementResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: '管理!A:A',
        });
        results.managementCount = (managementResponse.data.values?.length || 1) - 1; // ヘッダーを除く

        // クリニック名シートのデータ行数（ヘッダー行を除く）
        const clinicResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: `${clinicName}!A:A`,
        });
        results.clinicCount = (clinicResponse.data.values?.length || 1) - 1; // ヘッダーを除く

        return results;
    } catch (error) {
        console.error('[getSheetRowCounts] Error:', error);
        throw error;
    }
};

// ▼▼▼ [ここから変更] ▼▼▼
/**
 * [新規] 指定された単一セルの値を読み込む
 * @param {string} centralSheetId
 * @param {string} sheetName (例: "クリニックA_AI分析")
 * @param {string} cell (例: "B3")
 * @returns {Promise<string|null>} セルの値、または空
 */
exports.readSingleCell = async (centralSheetId, sheetName, cell) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    const range = `'${sheetName}'!${cell}`;
    console.log(`[googleSheetsService] Reading single cell: "${range}"`);
    
    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: range,
            valueRenderOption: 'FORMATTED_VALUE'
        });
        
        // 値が存在すれば [0][0] に入っている
        const value = response.data.values?.[0]?.[0];
        return value || null; // セルが空か存在しない場合は null を返す
        
    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            console.warn(`[readSingleCell] Sheet or cell not found: "${range}". Returning null.`);
            return null; // シートやセルが見つからない
        }
        console.error(`[readSingleCell] Error reading cell "${range}": ${e.message}`, e);
        throw new Error(`セル(${range})の読み込みに失敗しました: ${e.message}`);
    }
};
// ▲▲▲ [変更ここまで] ▲▲▲


// --- (変更なし) ヘルパー関数群 (addSheet, getSheetId, findOrCreateSheet, clearSheet, writeData) ---
async function addSheet(spreadsheetId, title) {
    try {
        const request = {
            spreadsheetId: spreadsheetId,
            resource: {
                requests: [{ addSheet: { properties: { title: title } } }]
            }
        };
        const response = await sheets.spreadsheets.batchUpdate(request);
        const newSheetId = response.data.replies[0].addSheet.properties.sheetId;
        console.log(`[Helper] Added sheet "${title}" (ID: ${newSheetId})`);
        return newSheetId;
    } catch (e) {
        if (e.message.includes('already exists')) {
            return await getSheetId(spreadsheetId, title);
        } else {
            console.error(`[Helper] Error adding sheet "${title}": ${e.message}`);
            throw e;
        }
    }
}

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

async function findOrCreateSheet(spreadsheetId, title) {
    try {
        const existingSheetId = await getSheetId(spreadsheetId, title);
        
        if (existingSheetId !== null) {
            return existingSheetId;
        }
        
        console.log(`[Helper] Sheet "${title}" not found. Creating...`);
        const newSheetId = await addSheet(spreadsheetId, title);
        return newSheetId;
        
    } catch (e) {
        console.error(`[Helper] Error in findOrCreateSheet for "${title}": ${e.message}`);
        throw e;
    }
}

async function clearSheet(spreadsheetId, range) {
    try {
        const formattedRange = range.includes('!') ? range : `'${range}'`;
        await sheets.spreadsheets.values.clear({
            spreadsheetId: spreadsheetId,
            range: formattedRange,
        });
    } catch (e) {
        console.error(`[Helper] Error clearing sheet "${range}": ${e.message}`);
        throw e;
    }
}

async function writeData(spreadsheetId, range, values, append = false) {
    if (!values || values.length === 0) {
        console.warn(`[Helper] No data to write for sheet "${range}".`);
        return;
    }

    try {
        const formattedRange = range.includes('!') ? range : `'${range}'`;
        if (append) {
            await sheets.spreadsheets.values.append({
                spreadsheetId: spreadsheetId,
                range: formattedRange,
                valueInputOption: 'USER_ENTERED',
                insertDataOption: 'INSERT_ROWS',
                resource: {
                    values: values
                }
            });
        } else {
            await sheets.spreadsheets.values.update({
                spreadsheetId: spreadsheetId,
                range: formattedRange,
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
