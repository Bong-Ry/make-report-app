const { google } = require('googleapis');

const KEYFILEPATH = '/etc/secrets/credentials.json';
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];

let sheets;

// --- 初期化 ---
try {
    const auth = new google.auth.GoogleAuth({ keyFile: KEYFILEPATH, scopes: SCOPES });
    sheets = google.sheets({ version: 'v4', auth });
    console.log('Google Sheets API client initialized successfully in googleSheets.js.');
} catch (err) {
    console.error('Failed to initialize Google Sheets API client in googleSheets.js:', err);
    // エラー発生時は sheets が undefined のままになる
}

// --- マスターシートからクリニックリストを取得 ---
exports.getMasterClinicList = async () => {
    const currentMasterSheetId = process.env.MASTER_SHEET_ID;
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    if (!currentMasterSheetId) throw new Error('サーバー設定エラー: マスターシートIDがありません。');

    const MASTER_RANGE = 'シート1!A2:A';
    console.log(`[googleSheetsService] Fetching master sheet clinics: ID=${currentMasterSheetId}, Range=${MASTER_RANGE}`);

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

// --- マスターシートからクリニック名とURLのマップを取得 ---
exports.getMasterClinicUrls = async () => {
    const currentMasterSheetId = process.env.MASTER_SHEET_ID;
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    if (!currentMasterSheetId) throw new Error('サーバー設定エラー: マスターシートIDがありません。');

    const MASTER_CLINIC_URL_RANGE = 'シート1!A2:B';
     console.log(`[googleSheetsService] Fetching master clinic URLs: ID=${currentMasterSheetId}, Range=${MASTER_CLINIC_URL_RANGE}`);

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

// --- 各クリニックのデータを取得し集計 ---
exports.fetchAndAggregateReportData = async (clinicUrls, period) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');

    const startDate = new Date(period.start + '-01T00:00:00Z');
    const [endYear, endMonth] = period.end.split('-').map(Number);
    const endDate = new Date(Date.UTC(endYear, endMonth, 0));
    endDate.setUTCHours(23, 59, 59, 999);
    console.log(`[googleSheetsService] Filtering data between ${startDate.toISOString()} and ${endDate.toISOString()}`);

    const reportData = {};
    const satisfactionKeys = ['非常に満足', '満足', 'ふつう', '不満', '非常に不満'];
    const ageKeys = ['10代', '20代', '30代', '40代'];
    const childrenKeys = ['1人', '2人', '3人', '4人', '5人以上'];
    const initializeCounts = (keys) => keys.reduce((acc, key) => { acc[key] = 0; return acc; }, {});

    // グラフデータ作成関数 (0件データを含む)
    const createChartData = (counts, keys) => {
        console.log(`[googleSheetsService] Creating chart data with counts:`, counts);
        const chartData = [['カテゴリ', '件数']];
        keys.forEach(key => {
            chartData.push([key, counts[key] || 0]);
        });
        console.log(`[googleSheetsService] Generated chart data:`, chartData);
        if (chartData.length <= 1) {
             console.warn('[googleSheetsService] createChartData generated header-only data.');
        }
        return chartData;
    };

    for (const clinicName in clinicUrls) {
        const clinicSheetId = clinicUrls[clinicName];
        console.log(`[googleSheetsService] Processing ${clinicName} (ID: ${clinicSheetId})`);
        const allNpsReasons = [], allFeedbacks_I = [], allFeedbacks_J = [], allFeedbacks_M = [];
        const satisfactionCounts_B = initializeCounts(satisfactionKeys), satisfactionCounts_C = initializeCounts(satisfactionKeys), satisfactionCounts_D = initializeCounts(satisfactionKeys), satisfactionCounts_E = initializeCounts(satisfactionKeys), satisfactionCounts_F = initializeCounts(satisfactionKeys), satisfactionCounts_G = initializeCounts(satisfactionKeys), satisfactionCounts_H = initializeCounts(satisfactionKeys);
        const childrenCounts_P = initializeCounts(childrenKeys);
        const ageCounts_O = initializeCounts(ageKeys);
        const incomeCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0 };
        let processedRowCount = 0;
        let matchedRowCount = 0;

        try {
            // ★ シート名は半角 '1' を使用
            const range = "'フォームの回答 1'!A:Q";
            console.log(`[googleSheetsService] Fetching clinic data: ID=${clinicSheetId}, Range=${range}`);
            const clinicDataResponse = await sheets.spreadsheets.values.get({ spreadsheetId: clinicSheetId, range: range, dateTimeRenderOption: 'SERIAL_NUMBER', valueRenderOption: 'UNFORMATTED_VALUE' });
            const clinicDataRows = clinicDataResponse.data.values;

            if (!clinicDataRows || clinicDataRows.length < 2) {
                console.log(`[googleSheetsService] No data or only header found for ${clinicName}.`);
                continue;
            }
            console.log(`[googleSheetsService] Fetched ${clinicDataRows.length} rows (including header) for ${clinicName}`);

            const header = clinicDataRows.shift();
            const timestampIndex = 0, satBIndex = 1, satCIndex = 2, satDIndex = 3, satEIndex = 4, satFIndex = 5, satGIndex = 6, satHIndex = 7, feedbackIIndex = 8, feedbackJIndex = 9, scoreKIndex = 10, reasonLIndex = 11, feedbackMIndex = 12, ageOIndex = 14, childrenPIndex = 15, incomeQIndex = 16;

            clinicDataRows.forEach((row, rowIndex) => {
                processedRowCount++;
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                const serialValue = row[timestampIndex];
                if (typeof serialValue !== 'number' || serialValue <= 0 || isNaN(serialValue)) return;
                const timestamp = new Date(excelEpoch.getTime() + serialValue * 24 * 60 * 60 * 1000);

                if (timestamp.getTime() >= startDate.getTime() && timestamp.getTime() <= endDate.getTime()) {
                    matchedRowCount++;
                    // 集計ロジック... (変更なし)
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
                }
            });
            console.log(`[googleSheetsService] For ${clinicName}: Processed ${processedRowCount} data rows, ${matchedRowCount} rows matched the period.`);

        } catch (e) {
             if (e.message && e.message.includes('Requested entity was not found')) {
                console.error(`[googleSheetsService] Error for ${clinicName}: Sheet 'フォームの回答 1' not found or spreadsheet ID invalid. Skipping clinic. Error: ${e.toString()}`);
             } else if (e.message && e.message.includes('Unable to parse range')) {
                 console.error(`[googleSheetsService] Error for ${clinicName}: Invalid range specified. Check sheet name and range format. Skipping clinic. Error: ${e.toString()}`);
             } else {
                console.error(`[googleSheetsService] Error processing sheet for ${clinicName}: ${e.toString()}`, e.stack);
             }
            continue; // 次のクリニックへ
        }

        // NPS理由をスコアごとにグループ化 (変更なし)
        const groupedByScore = allNpsReasons.reduce((acc, item) => { if (typeof item.score === 'number' && !isNaN(item.score)) { if (!acc[item.score]) acc[item.score] = []; acc[item.score].push(item.reason); } return acc; }, {});

        // 世帯年収グラフデータ作成 (変更なし)
        const incomeChartData = [['評価', '割合', { role: 'annotation' }]];
        const totalIncomeCount = Object.values(incomeCounts).reduce((a, b) => a + b, 0);
        console.log(`[googleSheetsService] Income counts for ${clinicName}:`, incomeCounts, `Total: ${totalIncomeCount}`);
        if (totalIncomeCount > 0) { for (let i = 1; i <= 10; i++) { const count = incomeCounts[i] || 0; const percentage = (count / totalIncomeCount) * 100; incomeChartData.push([String(i), percentage, `${Math.round(percentage)}%`]); } }
        console.log(`[googleSheetsService] Generated income chart data for ${clinicName}:`, incomeChartData);

        // レポートデータオブジェクト作成
        reportData[clinicName] = {
            npsData: { totalCount: allNpsReasons.length, results: groupedByScore, rawText: allNpsReasons.map(r => r.reason) },
            feedbackData: { i_column: { totalCount: allFeedbacks_I.length, results: allFeedbacks_I }, j_column: { totalCount: allFeedbacks_J.length, results: allFeedbacks_J }, m_column: { totalCount: allFeedbacks_M.length, results: allFeedbacks_M } },
            satisfactionData: { b_column: { results: createChartData(satisfactionCounts_B, satisfactionKeys) }, c_column: { results: createChartData(satisfactionCounts_C, satisfactionKeys) }, d_column: { results: createChartData(satisfactionCounts_D, satisfactionKeys) }, e_column: { results: createChartData(satisfactionCounts_E, satisfactionKeys) }, f_column: { results: createChartData(satisfactionCounts_F, satisfactionKeys) }, g_column: { results: createChartData(satisfactionCounts_G, satisfactionKeys) }, h_column: { results: createChartData(satisfactionCounts_H, satisfactionKeys) } },
            ageData: { results: createChartData(ageCounts_O, ageKeys) },
            childrenCountData: { results: createChartData(childrenCounts_P, childrenKeys) },
            incomeData: { results: incomeChartData, totalCount: totalIncomeCount }
        };
        console.log(`[googleSheetsService] Finished processing data for ${clinicName}`);
    }

    return reportData;
};
