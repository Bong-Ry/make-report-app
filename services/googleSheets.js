// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/services/googleSheets.js

const { google } = require('googleapis');
// ▼▼▼ [変更] ヘルパーからId取得と、[新規] 分析シート定義をインポート ▼▼▼
const { getSpreadsheetIdFromUrl, getAnalysisSheetConfig } = require('../utils/helpers');

// ▼▼▼ [変更なし] ▼▼▼
exports.GAS_SLIDE_GENERATOR_URL = 'https://script.google.com/macros/s/AKfycby-b31JKqSR5HNLi1fQxK1hePsxkpDL2StBhd1gsP_dKqFvjNRoqcTsca0hLSEzE3x2Xg/exec';
const GAS_SHEET_FINDER_URL = 'https://script.google.com/macros/s/AKfycbzn4rNw6NttPPmcJBpSKJifK8-Mb1CatsGhqvYF5G6BIAf6bOUuNS_E72drg0tH9re-qQ/exec';
const KEYFILEPATH = '/etc/secrets/credentials.json';
const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/presentations'
];
const MASTER_FOLDER_ID = '1_pJQKl5-RRi6h-U3EEooGmPkTrkF1Vbj';

let sheets;
let drive;
let slides;

try {
    const auth = new google.auth.GoogleAuth({ keyFile: KEYFILEPATH, scopes: SCOPES });
    sheets = google.sheets({ version: 'v4', auth });
    drive = google.drive({ version: 'v3', auth });
    slides = google.slides({ version: 'v1', auth });
    console.log('Google Sheets, Drive, and Slides API clients initialized successfully.');
} catch (err) {
    console.error('Failed to initialize Google API clients:', err);
}

exports.slides = slides;
// ▲▲▲ [変更なし] ▲▲▲


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
// === (変更なし) 関数 (1/X) (GAS呼び出し) ===
// =================================================================
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

// =================================================================
// === (変更なし) 更新関数 (2/X) (ヘッダーなし転記) ===
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
            // ▼▼▼ [修正] ユーザーのカスタム指示に基づきシート名を "フォームの回答 1" にする ▼▼▼
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
                
                // ▼▼▼ [変更なし] ヘッダーなしで書き込む ▼▼▼
                await writeData(centralSheetId, clinicSheetTitle, filteredRows);
                console.log(`[googleSheetsService-ETL] Wrote ${filteredRows.length} rows to sheet: "${clinicSheetTitle}" (HEADERLESS)`);

                // 5. データを「全体」タブに「追記」 (変更なし)
                await writeData(centralSheetId, '全体', filteredRows, true); // append = true
                console.log(`[googleSheetsService-ETL] Appended ${filteredRows.length} rows to sheet: "全体"`);
            }
            
            processedClinics.push(clinicName);

        } catch (e) {
            // ▼▼▼ [修正] ユーザーのカスタム指示に基づきシート名を "フォームの回答 1" にする ▼▼▼
            if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
                console.error(`[googleSheetsService-ETL] Error for ${clinicName}: Sheet 'フォームの回答 1' not found or invalid range. Skipping. Error: ${e.toString()}`);
            } else {
                console.error(`[googleSheetsService-ETL] Error processing sheet for ${clinicName}: ${e.toString()}`, e.stack);
            }
            continue; // 次のクリニックへ
        }
    }
    
    return processedClinics;
};

// =================================================================
// === (変更なし) 関数 (3/X) (集計) ===
// =================================================================
exports.getReportDataForCharts = async (centralSheetId, sheetName) => {
    // (既存のコードのまま - 変更なし)
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

// =================================================================
// === (変更なし) 関数 (4/X) (構築) ===
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
// === ▼▼▼ [削除] 古いAI分析関数 (5/X, 6/X, 7/X) ▼▼▼ ===
// =================================================================
// exports.saveAIAnalysisToSheet ... (削除)
// exports.getAIAnalysisFromSheet ... (削除)
// exports.updateAIAnalysisInSheet ... (削除)


// =================================================================
// === ▼▼▼ [削除] 古い市区町村関数 (8/X, 9/X) ▼▼▼ ===
// =================================================================
// exports.saveMunicipalityData ... (削除)
// exports.readMunicipalityData ... (削除)


// =================================================================
// === (変更なし) 関数 (10/X) (シート名一覧取得) ===
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
// === ▼▼▼ [新規] 分析シート用 汎用I/O関数群 (11/X) ▼▼▼ ===
// =================================================================

/**
 * [新規] 分析スプレッドシート（`[クリニック名]-分析`）のIDを取得（なければ作成）
 */
const analysisSheetCache = new Map(); // K: clinicName, V: spreadsheetId
async function findOrCreateAnalysisSheet(centralSheetId, clinicName) {
    const cacheKey = `${centralSheetId}-${clinicName}`;
    if (analysisSheetCache.has(cacheKey)) {
        return analysisSheetCache.get(cacheKey);
    }

    // TODO: 本来は centralSheetId とは別に、クリニックごとの分析スプシを作成・検索する
    //       （GAS App Maker のロジック）
    //       現状は、集計スプレッドシート (centralSheetId) をそのまま分析シートとしても使う
    //       （＝分析タブが集計スプシ内に作られる）
    
    //       もし `[クリニック名]-分析` という名前の *別ファイル* をDriveから検索/作成
    //       する場合は、ここに Drive API のロジックを追加する。
    
    //       **現在の実装**: centralSheetId をそのまま分析用スプレッドシートIDとして扱う
    // console.warn(`[getAnalysisSheetId] TODO: 本実装ではDrive APIで "[${clinicName}]-分析" ファイルを検索・作成する。現在は集計スプシIDを流用。`);
    
    const analysisSheetId = centralSheetId; // 仮
    analysisSheetCache.set(cacheKey, analysisSheetId);
    return analysisSheetId;
}

/**
 * [新規] 分析結果を、定義（CONFIG）に基づき単一分析シートに書き込む
 * @param {string} centralSheetId (集計スプシID)
 * @param {string} clinicName (例: "クリニックA")
 * @param {string} taskType (例: "L_ANALYSIS", "MUNICIPALITY_TABLE")
 * @param {string | string[][] | object} data - 保存するデータ
 */
exports.saveToAnalysisSheet = async (centralSheetId, clinicName, taskType, data) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    try {
        // 1. 保存先の定義を取得
        const { sheetName, range, type } = getAnalysisSheetConfig(taskType);
        
        // 2. 保存先のスプレッドシートIDを取得 (または centralSheetId を流用)
        //    (この関数が「分析用スプレッドシート」と「集計スプレッドシート」が同一だと仮定)
        const analysisSheetId = await findOrCreateAnalysisSheet(centralSheetId, clinicName); 
        
        // 3. 保存先の「タブ」を作成 (例: "AI分析", "データ")
        await findOrCreateSheet(analysisSheetId, sheetName);

        // 4. 保存形式に応じてデータを整形
        let values; // [[...], [...]] の形式
        
        if (type === 'CELL') {
            // 単一セルの場合
            if (typeof data !== 'string' && data != null) {
                console.warn(`[saveToAnalysisSheet] CELLタイプには文字列が必要です (task: ${taskType})。JSONに変換します。`);
                values = [[JSON.stringify(data, null, 2)]];
            } else {
                values = [[data || ""]]; // null や undefined を空文字に
            }
        } 
        else if (type === 'TABLE') {
            // テーブル (2D配列) の場合
            if (!Array.isArray(data) || (data.length > 0 && !Array.isArray(data[0]))) {
                throw new Error(`[saveToAnalysisSheet] TABLEタイプには2D配列データが必要です (task: ${taskType})`);
            }
            values = data.length > 0 ? data : [[]]; // 空配列だとエラーになるため [[]] を送る
        } 
        else {
            throw new Error(`[saveToAnalysisSheet] 未定義の保存タイプです: ${type}`);
        }

        // 5. データを書き込み (update = 上書き)
        const writeRange = `'${sheetName}'!${range}`;
        console.log(`[googleSheetsService] Saving to Analysis Sheet (ID: ${analysisSheetId}, Range: ${writeRange}, Task: ${taskType})`);
        
        // まずクリア (古いデータが残らないように)
        if (type === 'TABLE') {
             // テーブルの場合は、定義された範囲全体をクリア
             // (注意: 'B2' のような単一セル指定の場合、クリア範囲を広げる必要がある)
             // (現状の 'B2' 指定は「B2から書き始める」という意味で使っているため、クリアは書き込み時に上書きされる)
             
             // A1表記 (B2) から R1C1表記に変換して、適切な範囲をクリアする (実装例)
             // ... が、複雑になるため、writeData(append=false)の上書きで代用
             // await clearSheet(analysisSheetId, writeRange); // B2だけクリアしても意味がない
        }
        
        await writeData(analysisSheetId, writeRange, values, false); // append = false (上書き)

    } catch (err) {
        console.error(`[googleSheetsService] Error saving to analysis sheet (Task: ${taskType}): ${err.message}`, err);
        throw new Error(`分析結果(Task: ${taskType})のシート保存に失敗: ${err.message}`);
    }
};

/**
 * [新規] 分析結果を、定義（CONFIG）に基づき単一分析シートから読み込む
 * @param {string} centralSheetId
 * @param {string} clinicName
 * @param {string} taskType (例: "L_ANALYSIS", "MUNICIPALITY_TABLE")
 * @returns {Promise<string | string[][] | null>} - 読み込んだデータ (単一セルの場合は文字列、テーブルの場合は2D配列)
 */
exports.readFromAnalysisSheet = async (centralSheetId, clinicName, taskType) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    try {
        // 1. 読み込み元の定義を取得
        const { sheetName, range, type } = getAnalysisSheetConfig(taskType);
        
        // 2. 読み込み元のスプレッドシートIDを取得
        const analysisSheetId = await findOrCreateAnalysisSheet(centralSheetId, clinicName);
        
        // 3. 読み込み範囲
        let readRange = `'${sheetName}'!${range}`;
        
        if (type === 'TABLE') {
            // 'B2' のような単一セル指定の場合、'B2:ZZ' のように拡張してテーブル全体を読み込む
            const startCell = range.match(/^([A-Z]+)(\d+)$/);
            if (startCell) {
                // 'B2' -> 'B2:Z' (列全体) のように末尾を開放する
                // (注意: これだと G2:I のおすすめ理由を読み込む際に、B列の市区町村まで読んでしまう)
                // -> やはり config (helpers.js) で 'B2:E' や 'G2:I' のように範囲指定すべき
                // helpers.js を 'B2', 'G2' と定義してしまったため、ここで動的に補正
                if (taskType === 'MUNICIPALITY_TABLE') {
                    readRange = `'${sheetName}'!B2:E`; // B:D (件数) -> B:E (割合) に変更
                } else if (taskType === 'RECOMMENDATION_TABLE') {
                    readRange = `'${sheetName}'!G2:I`; // G:H (件数) -> G:I (割合) に変更
                }
            }
        }
        
        console.log(`[googleSheetsService] Reading from Analysis Sheet (ID: ${analysisSheetId}, Range: ${readRange}, Task: ${taskType})`);

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: analysisSheetId,
            range: readRange,
            valueRenderOption: 'FORMATTED_VALUE'
        });

        const values = response.data.values;

        if (!values || values.length === 0) {
            console.log(`[googleSheetsService] No data found for ${taskType}.`);
            return null;
        }

        // 4. 形式に応じて返す
        if (type === 'CELL') {
            return values[0][0] || null; // 単一セルの文字列
        } 
        else if (type === 'TABLE') {
            return values; // 2D配列
        }
        
        return null;

    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            console.warn(`[googleSheetsService] Sheet/Range not found for ${taskType}. Returning null.`);
            return null; // シートまたはセルが存在しない
        }
        console.error(`[googleSheetsService] Error reading from analysis sheet (Task: ${taskType}): ${e.message}`, e);
        throw new Error(`分析結果(Task: ${taskType})のシート読み込みに失敗: ${e.message}`);
    }
};

/**
 * [新規] 完了マーカーを読み込む (getTransferredList用)
 */
exports.readCompletionMarker = async (centralSheetId, clinicName) => {
    try {
        const marker = await exports.readFromAnalysisSheet(centralSheetId, clinicName, 'COMPLETION_MARKER');
        return (marker === 'COMPLETED'); // "COMPLETED" という文字列が書き込まれていれば true
    } catch (e) {
        // (readFromAnalysisSheet がエラー（シートない等）を null で返すため、ここでは false になる)
        console.warn(`[readCompletionMarker] Error reading marker for ${clinicName}: ${e.message}`);
        return false;
    }
};


// =================================================================
// === ヘルパー関数群 (変更なし) ===
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
        const newSheetId = response.data.replies[0].addSheet.properties.sheetId;
        console.log(`[Helper] Added sheet "${title}" (ID: ${newSheetId})`);
        return newSheetId;
    } catch (e) {
        if (e.message.includes('already exists')) {
            console.warn(`[Helper] Sheet "${title}" already exists.`);
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
            // console.log(`[Helper] Sheet "${title}" already exists (ID: ${existingSheetId}).`);
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

// (readCell, writeToCell は汎用関数 save/readFromAnalysisSheet に吸収されたため不要)

// ★ 既存の getSpreadsheetIdFromUrl を googleSheetsService の末尾にもエクスポート
exports.getSpreadsheetIdFromUrl = getSpreadsheetIdFromUrl;
