const { google } = require('googleapis');
const { getSpreadsheetIdFromUrl } = require('../utils/helpers'); // 既存のヘルパー

const KEYFILEPATH = '/etc/secrets/credentials.json';
// ▼▼▼ スコープを「読み書き」に変更し、Driveスコープも追加 ▼▼▼
const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets', // 読み書き
    'https://www.googleapis.com/auth/drive.file'    // ファイルの検索・移動
];
const MASTER_FOLDER_ID = '1_pJQKl5-RRi6h-U3EEooGmPkTrkF1Vbj'; // 集計スプシ作成先フォルダ

let sheets;
let drive; // Drive API クライアント

// --- 初期化 ---
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
// === ▼▼▼ 新規関数 (1/7) ▼▼▼ ===
// 集計用スプレッドシートを検索または新規作成する
// =================================================================
exports.findOrCreateCentralSheet = async (periodText) => {
    if (!sheets || !drive) throw new Error('Google APIクライアントが初期化されていません。');

    const fileName = periodText; // 例: "2025-09～2025-10"
    console.log(`[googleSheetsService] Finding or creating central sheet: "${fileName}"`);

    try {
        // 1. フォルダ内をファイル名で検索
        const searchRes = await drive.files.list({
            q: `name='${fileName}' and '${MASTER_FOLDER_ID}' in parents and trashed=false`,
            fields: 'files(id, name)',
            spaces: 'drive',
        });

        if (searchRes.data.files && searchRes.data.files.length > 0) {
            // 2. 見つかった場合: IDを返す
            const fileId = searchRes.data.files[0].id;
            console.log(`[googleSheetsService] Found existing sheet. ID: ${fileId}`);
            return fileId;
        }

        // 3. 見つからない場合: 新規作成
        console.log(`[googleSheetsService] No existing sheet found. Creating new one...`);
        const createRes = await sheets.spreadsheets.create({
            resource: {
                properties: {
                    title: fileName,
                },
                sheets: [
                    {
                        properties: {
                            title: '全体', // 最初のシート名を「全体」に
                        }
                    }
                ]
            },
            fields: 'spreadsheetId,sheets(properties(sheetId))',
        });

        const newSheetId = createRes.data.spreadsheetId;
        console.log(`[googleSheetsService] Created new sheet. ID: ${newSheetId}. Now moving to folder...`);

        // 4. 作成したファイルを指定フォルダに移動
        // (create時にparentを指定できないため、updateで移動する)
        await drive.files.update({
            fileId: newSheetId,
            addParents: MASTER_FOLDER_ID,
            removeParents: 'root', // rootから削除
            fields: 'id, parents',
        });
        
        console.log(`[googleSheetsService] Moved sheet ${newSheetId} to folder ${MASTER_FOLDER_ID}.`);

        // 5. ユーザー定義（フォームの回答 1）のシート名を追加
        // ※ 確実に 'フォームの回答 1' が存在するようにするため
        try {
            await addSheet(newSheetId, 'フォームの回答 1');
            console.log(`[googleSheetsService] Added 'フォームの回答 1' sheet for user preference.`);
        } catch (e) {
             console.warn(`[googleSheetsService] Could not add 'フォームの回答 1' sheet (maybe exists): ${e.message}`);
        }

        return newSheetId;

    } catch (err) {
        console.error(`[googleSheetsService] Error in findOrCreateCentralSheet for "${fileName}":`, err);
        throw new Error(`集計スプレッドシートの検索または作成に失敗しました: ${err.message}`);
    }
};

// =================================================================
// === ▼▼▼ 更新関数 (2/7) ▼▼▼ ===
// 役割変更: 元シートからデータを読み取り、集計スプシに「転記」する
// =================================================================
exports.fetchAndAggregateReportData = async (clinicUrls, period, centralSheetId) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    if (!centralSheetId) throw new Error('集計スプレッドシートIDが指定されていません。');

    const startDate = new Date(period.start + '-01T00:00:00Z');
    const [endYear, endMonth] = period.end.split('-').map(Number);
    const endDate = new Date(Date.UTC(endYear, endMonth, 0));
    endDate.setUTCHours(23, 59, 59, 999);
    console.log(`[googleSheetsService-ETL] Filtering data between ${startDate.toISOString()} and ${endDate.toISOString()}`);

    const processedClinics = []; // 正常に処理できたクリニック名

    for (const clinicName in clinicUrls) {
        const clinicSheetId = clinicUrls[clinicName];
        console.log(`[googleSheetsService-ETL] Processing ${clinicName} (Source ID: ${clinicSheetId})`);

        try {
            // 1. 元データ（フォームの回答 1）を読み取る
            const range = "'フォームの回答 1'!A:R"; // R列(郵便番号)まで
            const clinicDataResponse = await sheets.spreadsheets.values.get({
                spreadsheetId: clinicSheetId,
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

            // 2. 期間でフィルタリング
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
                // 3. 集計スプシに「クリニック名」のタブを作成（またはクリア）
                const clinicSheetTitle = clinicName; // タブ名をクリニック名に
                await findOrCreateSheet(centralSheetId, clinicSheetTitle);
                
                // 4. データを「クリニック名」タブに書き込み (ヘッダー + データ)
                await clearSheet(centralSheetId, clinicSheetTitle);
                await writeData(centralSheetId, clinicSheetTitle, [header, ...filteredRows]);
                console.log(`[googleSheetsService-ETL] Wrote ${filteredRows.length} rows to sheet: "${clinicSheetTitle}"`);

                // 5. データを「全体」タブに「追記」 (データのみ)
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
// === ▼▼▼ 新規関数 (3/7) ▼▼▼ ===
// 集計スプシからデータを読み取り、グラフ用に「集計」する
// (旧 fetchAndAggregateReportData の集計ロジックを移植)
// =================================================================
exports.getReportDataForCharts = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    
    console.log(`[googleSheetsService-AGG] Aggregating data from Sheet ID: ${centralSheetId}, Tab: "${sheetName}"`);

    // --- 集計用の定義 (旧関数からコピー) ---
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
    // --- 集計用変数の初期化 (旧関数からコピー) ---
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
            // 日付はすでにフィルタリング済みだが、念のため元データと同じオプション
            dateTimeRenderOption: 'SERIAL_NUMBER', 
            valueRenderOption: 'UNFORMATTED_VALUE'
        });

        const clinicDataRows = clinicDataResponse.data.values;

        if (!clinicDataRows || clinicDataRows.length < 2) {
            console.log(`[googleSheetsService-AGG] No data or only header found in "${sheetName}".`);
            // データが空でも空の集計結果を返す
            return buildReportDataObject(null); // nullを渡して空のオブジェクトを構築
        }

        const header = clinicDataRows.shift();
        // インデックス定義 (旧関数からコピー)
        const timestampIndex = 0, satBIndex = 1, satCIndex = 2, satDIndex = 3, satEIndex = 4, satFIndex = 5, satGIndex = 6, satHIndex = 7, feedbackIIndex = 8, feedbackJIndex = 9, scoreKIndex = 10, reasonLIndex = 11, feedbackMIndex = 12, recommendationNIndex = 13, ageOIndex = 14, childrenPIndex = 15, incomeQIndex = 16, postalCodeRIndex = 17;

        // 2. データをループして集計 (旧関数のロジックをそのまま使用)
        // ※ 日付フィルタリングは不要 (すでにフィルタリング済みのデータのため)
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

        // 3. レポートオブジェクトを構築して返す
        const aggregationData = {
            allNpsReasons, allFeedbacks_I, allFeedbacks_J, allFeedbacks_M,
            satisfactionCounts_B, satisfactionCounts_C, satisfactionCounts_D, satisfactionCounts_E, satisfactionCounts_F, satisfactionCounts_G, satisfactionCounts_H,
            childrenCounts_P, ageCounts_O, incomeCounts, postalCodeCounts,
            npsScoreCounts, recommendationCounts, recommendationOthers,
            satisfactionKeys, ageKeys, childrenKeys, recommendationKeys,
            createChartData // 関数を渡す
        };

        return buildReportDataObject(aggregationData);

    } catch (e) {
        console.error(`[googleSheetsService-AGG] Error aggregating data from "${sheetName}": ${e.toString()}`, e.stack);
        // 読み取りエラーの場合は空のデータを返す
        return buildReportDataObject(null);
    }
};

// =================================================================
// === ▼▼▼ 新規関数 (4/7) ▼▼▼ ===
// 集計データからレポートオブジェクトを構築する（旧関数の後半部分）
// =================================================================
function buildReportDataObject(data) {
    // data が null の場合（データなし）
    if (!data) {
        // 空の構造を返す
        const emptyChart = [['カテゴリ', '件数']];
        const emptyCounts = {};
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
    
    // data が存在する場合
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
// === ▼▼▼ 新規関数 (5/7) ▼▼▼ ===
// AI分析結果を3枚のシートに分けて保存する
// =================================================================
exports.saveAIAnalysisToSheet = async (centralSheetId, clinicName, analysisType, jsonData) => {
    if (!jsonData) throw new Error('AI analysis JSON data is missing.');

    const baseSheetName = `${clinicName}-AI分析-${analysisType}`;
    
    try {
        // 1. 分析 (Analysis)
        const analysisSheetName = `${baseSheetName}-分析`;
        const analysisContent = (jsonData.analysis && jsonData.analysis.themes)
            ? jsonData.analysis.themes.map(t => `【${t.title}】\n${t.summary}`).join('\n\n---\n\n')
            : '分析データがありません。';
        await findOrCreateSheet(centralSheetId, analysisSheetName);
        await writeToCell(centralSheetId, analysisSheetName, 'A1', analysisContent);
        console.log(`[googleSheetsService-AI] Saved Analysis to: "${analysisSheetName}"`);

        // 2. 改善提案 (Suggestions)
        const suggestionSheetName = `${baseSheetName}-改善案`;
        const suggestionContent = (jsonData.suggestions && jsonData.suggestions.items)
            ? jsonData.suggestions.items.map(i => `【${i.themeTitle}】\n${i.suggestion}`).join('\n\n---\n\n')
            : '改善提案データがありません。';
        await findOrCreateSheet(centralSheetId, suggestionSheetName);
        await writeToCell(centralSheetId, suggestionSheetName, 'A1', suggestionContent);
        console.log(`[googleSheetsService-AI] Saved Suggestions to: "${suggestionSheetName}"`);

        // 3. 総評 (Overall)
        const overallSheetName = `${baseSheetName}-総評`;
        const overallContent = (jsonData.overall && jsonData.overall.summary)
            ? jsonData.overall.summary
            : '総評データがありません。';
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
// === ▼▼▼ 新規関数 (6/7) ▼▼▼ ===
// 3枚のシートからAI分析結果を読み出す
// =================================================================
exports.getAIAnalysisFromSheet = async (centralSheetId, clinicName, analysisType) => {
    const baseSheetName = `${clinicName}-AI分析-${analysisType}`;
    
    try {
        const analysisSheetName = `${baseSheetName}-分析`;
        const suggestionSheetName = `${baseSheetName}-改善案`;
        const overallSheetName = `${baseSheetName}-総評`;

        // 3枚のシートから並行してA1セルを読み取る
        const [analysisRes, suggestionRes, overallRes] = await Promise.all([
            readCell(centralSheetId, analysisSheetName, 'A1'),
            readCell(centralSheetId, suggestionSheetName, 'A1'),
            readCell(centralSheetId, overallSheetName, 'A1')
        ]);
        
        console.log(`[googleSheetsService-AI] Read AI analysis from 3 sheets for: "${baseSheetName}"`);
        
        // フロントエンドが期待する { analysis: { title: ..., themes: [...] }, ... } の
        // 形式に再構築せず、タブ表示用の { analysis: "...", suggestions: "...", overall: "..." } で返す
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
// === ▼▼▼ 新規関数 (7/7) ▼▼▼ ===
// AI分析（編集後）を指定のシートのA1に保存する
// =================================================================
exports.updateAIAnalysisInSheet = async (centralSheetId, sheetName, content) => {
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
// === ヘルパー関数群 ===
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
        await sheets.spreadsheets.batchUpdate(request);
    } catch (e) {
        if (e.message.includes('already exists')) {
            console.warn(`[Helper] Sheet "${title}" already exists.`);
        } else {
            console.error(`[Helper] Error adding sheet "${title}": ${e.message}`);
            throw e;
        }
    }
}

/**
 * [Helper] 指定したシート（タブ）を見つけて作成する（存在すれば何もしない）
 */
async function findOrCreateSheet(spreadsheetId, title) {
    try {
        // 1. シートのメタデータを取得
        const metadata = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            fields: 'sheets(properties(title))'
        });
        
        const exists = metadata.data.sheets.some(sheet => sheet.properties.title === title);
        
        // 2. 存在しない場合のみ作成
        if (!exists) {
            console.log(`[Helper] Sheet "${title}" not found. Creating...`);
            await addSheet(spreadsheetId, title);
        } else {
             console.log(`[Helper] Sheet "${title}" already exists.`);
        }
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
 * @param {string} spreadsheetId
 * @param {string} range (シート名)
 * @param {Array<Array<string>>} values
 * @param {boolean} [append=false] 追記モードにするか
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
 * @returns {string | null} セルの値
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
