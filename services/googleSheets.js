// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/services/googleSheets.js

const { google } = require('googleapis');
const { getSpreadsheetIdFromUrl, getAiAnalysisKeys } = require('../utils/helpers');

// --- 認証・初期化 ---
const GAS_SHEET_FINDER_URL = 'https://script.google.com/macros/s/AKfycbzn4rNw6NttPPmcJBpSKJifK8-Mb1CatsGhqvYF5G6BIAf6bOUuNS_E72drg0tH9re-qQ/exec';
const KEYFILEPATH = '/etc/secrets/credentials.json';
const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive', // Drive API (ファイル名取得用)
];

// ★★★ [設定] フォルダID設定 ★★★
const FOLDER_CONFIG = {
    MAIN:       '1_pJQKl5-RRi6h-U3EEooGmPkTrkF1Vbj', // ① 全体・管理 (既存)
    RAW:        '1baxkwAXMkgFYd6lg4AMeQqnWfx2-uv6A', // ② 元データ (RAW)
    REC:        '1t-rzPW2BiLOXCb_XlMEWT8DvE1iM3yO-', // ③ おすすめ理由 (REC)
    AI:         '1kO9EWERPUO7pbhq51kr9eVG2aJyalIfM', // ④ AI分析 (AI)

    // NPSコメント
    NPS_10:     '1p5uPULplr4jS7LCwKaz3JsmOWwqNbx1V', // ⑤ NPS 10
    NPS_9:      '1KL6IpplS3Uapgja0ku1OQibCtpt-bt1x', // ⑤ NPS 9
    NPS_8:      '13ptWLa5z--keuCIBB-ihrI9bfNG7Fdoc', // ⑤ NPS 8
    NPS_7:      '1A00rQFe9fWu8z70o1vUIy0KZfQ49JPU4', // ⑤ NPS 7
    NPS_6_UNDER:'1YwysnvQn6J7-3JNYEAgU8_4iASv7yx5X', // ⑤ NPS 6以下

    // その他コメント
    GOODBAD:    '1ofRq1uS9hrJ86NFH86cHpheVi4WCm4KI', // ⑥ 良/悪点 (GOODBAD)
    STAFF:      '1x6-f5yEH6KzEIxNznRK2S5Vp6nOHyPXM', // ⑦ スタッフ (STAFF)
    DELIVERY:   '1waeSxj0cCjd4YLDVLCyxDJ8d5JHJ53kt'  // ⑧ お産意見 (DELIVERY)
};
// ★★★ 設定ここまで ★★★

const MASTER_FOLDER_ID = FOLDER_CONFIG.MAIN;

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

// =================================================================
// === [新規] シンプルファイル管理 (同名ファイル検索方式) ===
// =================================================================

// IDキャッシュ (MainSheetId + FolderKey -> SubFileId)
const fileIdCache = new Map();
// ファイル名キャッシュ (MainSheetId -> "2025-04～2025-09")
const fileNameCache = new Map();

// ファイル種別の定義
const FILE_TYPES = {
    MAIN: 'MAIN', RAW: 'RAW', REC: 'REC', AI: 'AI',
    NPS_10: 'NPS_10', NPS_9: 'NPS_9', NPS_8: 'NPS_8', NPS_7: 'NPS_7', NPS_6_UNDER: 'NPS_6',
    GOODBAD: 'GOODBAD', STAFF: 'STAFF', DELIVERY: 'DELIVERY'
};

/**
 * [変更] 期間選択時に呼ばれる関数 (オーケストレーター)
 * 指定された期間名のファイルが各フォルダにあるか確認し、なければ一斉に作成します。
 */
exports.findOrCreateCentralSheet = async (periodText) => {
    console.log(`[Orchestrator] Initializing ALL files for period: "${periodText}"`);

    // 作成・確認対象のファイル定義一覧
    const fileDefinitions = [
        { type: FILE_TYPES.MAIN,        folderId: FOLDER_CONFIG.MAIN },
        { type: FILE_TYPES.RAW,         folderId: FOLDER_CONFIG.RAW },
        { type: FILE_TYPES.REC,         folderId: FOLDER_CONFIG.REC },
        { type: FILE_TYPES.AI,          folderId: FOLDER_CONFIG.AI }, // 市区町村含む
        { type: FILE_TYPES.NPS_10,      folderId: FOLDER_CONFIG.NPS_10 },
        { type: FILE_TYPES.NPS_9,       folderId: FOLDER_CONFIG.NPS_9 },
        { type: FILE_TYPES.NPS_8,       folderId: FOLDER_CONFIG.NPS_8 },
        { type: FILE_TYPES.NPS_7,       folderId: FOLDER_CONFIG.NPS_7 },
        { type: FILE_TYPES.NPS_6_UNDER, folderId: FOLDER_CONFIG.NPS_6_UNDER },
        { type: FILE_TYPES.GOODBAD,     folderId: FOLDER_CONFIG.GOODBAD },
        { type: FILE_TYPES.STAFF,       folderId: FOLDER_CONFIG.STAFF },
        { type: FILE_TYPES.DELIVERY,    folderId: FOLDER_CONFIG.DELIVERY },
    ];

    // 1. すべてのファイルについて GAS API を呼び出し (並列実行)
    // GAS側は「フォルダ内に同名ファイルがあればそのIDを返し、なければ作成してIDを返す」挙動
    const results = await Promise.all(fileDefinitions.map(async (def) => {
        // API負荷分散のためわずかにウェイトを入れる
        await new Promise(r => setTimeout(r, Math.random() * 2000));

        const id = await callGasToCreateFile(periodText, def.folderId);
        console.log(`[Orchestrator] Verified/Created ${def.type}: ${id}`);
        return { type: def.type, id: id };
    }));

    // 2. メインファイルのIDを特定
    const mainFileEntry = results.find(r => r.type === FILE_TYPES.MAIN);
    if (!mainFileEntry) throw new Error("メインファイルの作成に失敗しました。");
    const mainSheetId = mainFileEntry.id;

    // 3. 結果をすべてキャッシュに保存 (後続の処理でAPIを呼ばなくて済むようにする)
    results.forEach(res => {
        const cacheKey = `${mainSheetId}_${res.type}`;
        fileIdCache.set(cacheKey, res.id);
    });

    // メインファイルの期間名もキャッシュ
    fileNameCache.set(mainSheetId, periodText);

    // (オプション) メインファイルの管理シートにも一応記録しておく（バックアップとして）
    saveFileMapToManagementSheet(mainSheetId, results).catch(e => console.warn("管理シートへの記録に失敗(非致命的):", e));

    console.log(`[Orchestrator] All 12 files are ready. Main ID: ${mainSheetId}`);
    return mainSheetId;
};

/**
 * [重要] 司令塔関数: シート名から、書き込むべきファイルのIDを返す
 */
async function getTargetSpreadsheetId(mainSheetId, sheetName, clinicName) {
    // 1. フォルダIDの特定
    let targetFolderId = FOLDER_CONFIG.MAIN;
    let typeKey = 'MAIN';

    if (['全体', '管理', '全体-おすすめ理由'].includes(sheetName)) {
        return mainSheetId; // メインファイルそのもの
    } else if (sheetName.endsWith('_AI分析') || sheetName.endsWith('_市区町村')) {
        targetFolderId = FOLDER_CONFIG.AI;
        typeKey = 'AI';
    } else if (sheetName.endsWith('_おすすめ理由')) {
        targetFolderId = FOLDER_CONFIG.REC;
        typeKey = 'REC';
    } else if (sheetName.endsWith('_よかった点悪かった点')) {
        targetFolderId = FOLDER_CONFIG.GOODBAD;
        typeKey = 'GOODBAD';
    } else if (sheetName.endsWith('_印象スタッフ')) {
        targetFolderId = FOLDER_CONFIG.STAFF;
        typeKey = 'STAFF';
    } else if (sheetName.endsWith('_お産意見')) {
        targetFolderId = FOLDER_CONFIG.DELIVERY;
        typeKey = 'DELIVERY';
    } else if (sheetName.endsWith('_NPS10')) {
        targetFolderId = FOLDER_CONFIG.NPS_10;
        typeKey = 'NPS_10';
    } else if (sheetName.endsWith('_NPS9')) {
        targetFolderId = FOLDER_CONFIG.NPS_9;
        typeKey = 'NPS_9';
    } else if (sheetName.endsWith('_NPS8')) {
        targetFolderId = FOLDER_CONFIG.NPS_8;
        typeKey = 'NPS_8';
    } else if (sheetName.endsWith('_NPS7')) {
        targetFolderId = FOLDER_CONFIG.NPS_7;
        typeKey = 'NPS_7';
    } else if (sheetName.endsWith('_NPS6以下')) {
        targetFolderId = FOLDER_CONFIG.NPS_6_UNDER;
        typeKey = 'NPS_6_UNDER';
    } else if (sheetName === clinicName) {
        targetFolderId = FOLDER_CONFIG.RAW; // 元データ
        typeKey = 'RAW';
    }

    // メインフォルダの場合はそのまま返す
    if (targetFolderId === FOLDER_CONFIG.MAIN) {
        return mainSheetId;
    }

    // 2. キャッシュ確認 (findOrCreateCentralSheetで作成済みならここヒットする)
    const cacheKey = `${mainSheetId}_${typeKey}`;
    if (fileIdCache.has(cacheKey)) {
        return fileIdCache.get(cacheKey);
    }

    // 3. キャッシュにない場合（サーバー再起動後など）は、ファイル名から再取得を試みる
    let periodFileName = fileNameCache.get(mainSheetId);
    if (!periodFileName) {
        try {
            const fileMeta = await drive.files.get({
                fileId: mainSheetId,
                fields: 'name'
            });
            periodFileName = fileMeta.data.name;
            fileNameCache.set(mainSheetId, periodFileName);
        } catch (e) {
            console.error(`[Orchestrator] Failed to get filename for ID ${mainSheetId}`, e);
            throw new Error('メインファイル名の取得に失敗しました');
        }
    }

    // 4. GASを呼んでIDを取得（作成済みならIDだけ返ってくる）
    const targetFileId = await callGasToCreateFile(periodFileName, targetFolderId);

    // 5. 結果をキャッシュして返す
    fileIdCache.set(cacheKey, targetFileId);
    return targetFileId;
}

/**
 * GAS APIを呼んでファイルを作成/取得する
 */
async function callGasToCreateFile(fileName, folderId) {
    try {
        const response = await fetch(GAS_SHEET_FINDER_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                periodText: fileName,
                folderId: folderId
            })
        });
        const result = await response.json();
        if (result.status === 'ok' && result.spreadsheetId) {
            return result.spreadsheetId;
        }
        throw new Error(result.message || 'Unknown GAS error');
    } catch (e) {
        console.error(`[GAS] Failed to create/find file "${fileName}" in folder "${folderId}":`, e.message);
        throw e;
    }
}

// (管理シートへのバックアップ記録用ヘルパー)
async function saveFileMapToManagementSheet(mainSheetId, results) {
    const rows = results.map(r => [r.type, r.id]);
    try {
        await findOrCreateSheet(mainSheetId, '管理');
        await writeData(mainSheetId, "'管理'!D:E", rows, true);
    } catch (e) {}
}


// =================================================================
// === 既存関数の改修（getTargetSpreadsheetId を適用） ===
// =================================================================

// --- データ転記 (ETL) ---
exports.fetchAndAggregateReportData = async (clinicUrls, period, centralSheetId) => {
    const startDate = new Date(period.start + '-01T00:00:00Z');
    const [endYear, endMonth] = period.end.split('-').map(Number);
    const endDate = new Date(Date.UTC(endYear, endMonth, 0));
    endDate.setUTCHours(23, 59, 59, 999);

    const processedClinics = [];

    for (const clinicName in clinicUrls) {
        const sourceSheetId = getSpreadsheetIdFromUrl(clinicUrls[clinicName]);
        if (!sourceSheetId) continue;

        try {
            const range = "'フォームの回答 1'!A:R";
            const clinicDataResponse = await sheets.spreadsheets.values.get({
                spreadsheetId: sourceSheetId, range, dateTimeRenderOption: 'SERIAL_NUMBER', valueRenderOption: 'UNFORMATTED_VALUE'
            });
            const clinicDataRows = clinicDataResponse.data.values;
            if (!clinicDataRows || clinicDataRows.length < 2) continue;

            const dataBody = clinicDataRows.slice(1);
            const timestampIndex = 0;
            const filteredRows = [];

            dataBody.forEach((row) => {
                const serialValue = row[timestampIndex];
                if (typeof serialValue !== 'number' || serialValue <= 0) return;
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                const timestamp = new Date(excelEpoch.getTime() + serialValue * 24 * 60 * 60 * 1000);
                if (timestamp.getTime() >= startDate.getTime() && timestamp.getTime() <= endDate.getTime()) {
                    filteredRows.push(row);
                }
            });

            if (filteredRows.length > 0) {
                // 1. 個別データシート (RAWフォルダのファイルへ)
                const clinicSheetTitle = clinicName;
                const targetId = await getTargetSpreadsheetId(centralSheetId, clinicSheetTitle, clinicName);

                const sheetId = await findOrCreateSheet(targetId, clinicSheetTitle);
                await clearSheet(targetId, clinicSheetTitle);
                await writeData(targetId, clinicSheetTitle, filteredRows);

                const rowCount = filteredRows.length;
                const colCount = filteredRows[0] ? filteredRows[0].length : 18;
                await resizeSheetToFitData(targetId, sheetId, rowCount, colCount);

                // 2. 全体シート (MAINファイルへ)
                await findOrCreateSheet(centralSheetId, '全体');
                await writeData(centralSheetId, '全体', filteredRows, true);
            }

            processedClinics.push(clinicName);

        } catch (e) {
            console.error(`Error for ${clinicName}: ${e.message}`);
            continue;
        }
    }
    return processedClinics;
};

// --- データ集計 (読み込み) ---
exports.getReportDataForCharts = async (centralSheetId, sheetName) => {
    let targetId = centralSheetId;

    if (sheetName !== '全体') {
        // クリニック名の場合は RAWフォルダのファイルID を取得
        targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName);
    }

    console.log(`[AGG] Reading data from ID: ${targetId}, Sheet: "${sheetName}"`);

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
            spreadsheetId: targetId, // ★切り替えたIDを使用
            range: range,
            dateTimeRenderOption: 'SERIAL_NUMBER',
            valueRenderOption: 'UNFORMATTED_VALUE'
        });
        const clinicDataRows = clinicDataResponse.data.values;
        if (!clinicDataRows || clinicDataRows.length < 1) {
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

// --- buildReportDataObject (変更なし) ---
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

// --- マスターシート関数 (変更なし) ---
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

// --- コメントシート名ヘルパー (変更なし) ---
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

// --- コメントI/O (変更なし: 12件設定) ---
const ROWS_PER_COLUMN = 12;

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
        return;
    }

    console.log(`[googleSheetsService-Comments] Saving ${comments.length} comments to "${sheetName}"...`);

    const columnsData = formatCommentsToColumns(comments);
    const rowsData = transpose(columnsData);

    try {
        const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);

        const sheetId = await findOrCreateSheet(targetId, sheetName);
        await clearSheet(targetId, sheetName);
        await writeData(targetId, sheetName, rowsData);

        const rowCount = rowsData.length;
        const colCount = rowsData[0] ? rowsData[0].length : 1;
        await resizeSheetToFitData(targetId, sheetId, rowCount, colCount);

        console.log(`[googleSheetsService-Comments] Saved comments to "${sheetName}".`);
    } catch (e) {
        console.error(`[googleSheetsService-Comments] Failed to save comments to "${sheetName}": ${e.message}`, e);
        throw e;
    }
};

exports.readCommentsBySheet = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');

    const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);
    const range = `'${sheetName}'!A:Z`;
    console.log(`[googleSheetsService-Comments] Reading from ID: ${targetId}, Sheet: "${sheetName}"`);

    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: targetId,
            range: range,
            valueRenderOption: 'FORMATTED_VALUE',
            majorDimension: 'COLUMNS'
        });

        const columns = response.data.values;

        if (!columns || columns.length === 0) {
            return [];
        }

        return columns;

    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            return [];
        }
        console.error(`[googleSheetsService-Comments] Error reading comment data: ${e.message}`, e);
        throw new Error(`コメントシート(${sheetName})の読み込みに失敗しました: ${e.message}`);
    }
};

exports.updateCommentSheetCell = async (centralSheetId, sheetName, cell, value) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');

    const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);
    const range = `'${sheetName}'!${cell}`;

    console.log(`[googleSheetsService-Comments] Updating ID: ${targetId}, cell "${range}"`);

    try {
        await writeData(targetId, range, [[value]], false);
        console.log(`[googleSheetsService-Comments] Cell update successful.`);
    } catch (e) {
        console.error(`[googleSheetsService-Comments] Error updating cell "${range}": ${e.message}`, e);
        throw new Error(`コメントシート(${sheetName})のセル(${range})更新に失敗しました: ${e.message}`);
    }
};

exports.getSheetTitles = async (spreadsheetId) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        const metadata = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            fields: 'sheets(properties(title))'
        });

        const titles = metadata.data.sheets.map(sheet => sheet.properties.title);
        console.log(`[getSheetTitles] Found ${titles.length} sheets in MAIN file.`);
        return titles;

    } catch (e) {
        console.error(`[getSheetTitles] Error getting sheet titles: ${e.message}`);
        throw new Error(`シート一覧の取得に失敗しました: ${e.message}`);
    }
};

exports.saveTableToSheet = async (centralSheetId, sheetName, dataRows) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService] Saving Table to Sheet: "${sheetName}"`);

    try {
        const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);

        const sheetId = await findOrCreateSheet(targetId, sheetName);
        await clearSheet(targetId, sheetName);
        await writeData(targetId, sheetName, dataRows);

        if (sheetId) {
             await sheets.spreadsheets.batchUpdate({
                spreadsheetId: targetId,
                resource: { requests: [
                    { autoResizeDimensions: {
                        dimensions: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 4 }
                    }}
                ]}
            });
        }

        const rowCount = dataRows.length;
        const colCount = dataRows[0] ? dataRows[0].length : 4;
        await resizeSheetToFitData(targetId, sheetId, rowCount, colCount);

        console.log(`[googleSheetsService] Successfully saved Table for "${sheetName}"`);

    } catch (err) {
        console.error(`[googleSheetsService] Error saving Table data: ${err.message}`, err);
        throw new Error(`分析テーブルのシート保存に失敗しました: ${err.message}`);
    }
};

exports.saveAiAnalysisData = async (centralSheetId, sheetName, aiDataMap) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    console.log(`[googleSheetsService] Saving AI Key-Value Data to Sheet: "${sheetName}"`);

    try {
        const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);

        const sheetId = await findOrCreateSheet(targetId, sheetName);
        await clearSheet(targetId, sheetName);

        const allKeys = getAiAnalysisKeys();
        const dataRows = allKeys.map(key => {
            const value = aiDataMap.get(key) || '（データなし）';
            return [key, value];
        });

        const header = ['項目キー', '分析文章データ'];
        const finalData = [header, ...dataRows];

        await writeData(targetId, `'${sheetName}'!A1`, finalData);

        if (sheetId) {
             await sheets.spreadsheets.batchUpdate({
                spreadsheetId: targetId,
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

        await resizeSheetToFitData(targetId, sheetId, 16, 2);

        console.log(`[googleSheetsService] Successfully saved AI Key-Value Data for "${sheetName}"`);

    } catch (err) {
        console.error(`[googleSheetsService] Error saving AI Key-Value data: ${err.message}`, err);
        throw new Error(`AI分析(Key-Value)のシート保存に失敗しました: ${err.message}`);
    }
};

exports.readAiAnalysisData = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');

    const allKeys = getAiAnalysisKeys();
    const aiDataMap = new Map();

    try {
        const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);
        console.log(`[googleSheetsService] Reading AI Data from ID: ${targetId}, Sheet: "${sheetName}"`);

        const range = `'${sheetName}'!B2:B16`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: targetId,
            range: range,
            valueRenderOption: 'FORMATTED_VALUE'
        });

        const rows = response.data.values;

        allKeys.forEach((key, index) => {
            const value = (rows && rows[index] && rows[index][0] != null)
                ? rows[index][0]
                : '（データがありません）';
            aiDataMap.set(key, value);
        });

        return aiDataMap;

    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            allKeys.forEach(key => {
                aiDataMap.set(key, '（データがありません）');
            });
            return aiDataMap;
        }
        console.error(`[googleSheetsService] Error reading AI Key-Value data: ${e.message}`, e);
        throw new Error(`AI分析(Key-Value)のシート読み込みに失敗しました: ${e.message}`);
    }
};

const MANAGEMENT_SHEET_NAME = '管理';

exports.writeInitialMarker = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        await findOrCreateSheet(centralSheetId, MANAGEMENT_SHEET_NAME);
        await writeData(centralSheetId, MANAGEMENT_SHEET_NAME, [[clinicName]], true);
    } catch (e) {
        console.error(`[googleSheetsService-Marker] Error writing initial marker: ${e.message}`, e);
    }
};

exports.writeCompletionMarker = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
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
            await writeData(centralSheetId, MANAGEMENT_SHEET_NAME, [[clinicName, 'Complete']], true);
        } else {
            const cellToUpdate = `'${MANAGEMENT_SHEET_NAME}'!B${rowIndex + 1}`;
            await writeData(centralSheetId, cellToUpdate, [['Complete']], false);
        }

    } catch (e) {
        console.error(`[googleSheetsService-Marker] Error writing completion marker: ${e.message}`, e);
    }
};

exports.readCompletionStatusMap = async (centralSheetId) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');

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

exports.getSheetRowCounts = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');

    try {
        const results = {};

        const overallResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: '全体!A:A',
        });
        results.overallCount = (overallResponse.data.values?.length || 1) - 1;

        const managementResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: centralSheetId,
            range: '管理!A:A',
        });
        results.managementCount = (managementResponse.data.values?.length || 1) - 1;

        const targetId = await getTargetSpreadsheetId(centralSheetId, clinicName, clinicName);
        const clinicResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: targetId,
            range: `${clinicName}!A:A`,
        });
        results.clinicCount = (clinicResponse.data.values?.length || 1) - 1;

        return results;
    } catch (error) {
        console.error('[getSheetRowCounts] Error:', error);
        throw error;
    }
};

exports.readSingleCell = async (centralSheetId, sheetName, cell) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');

    try {
        const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);
        const range = `'${sheetName}'!${cell}`;
        console.log(`[googleSheetsService] Reading single cell ID: ${targetId}, "${range}"`);

        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: targetId,
            range: range,
            valueRenderOption: 'FORMATTED_VALUE'
        });

        const value = response.data.values?.[0]?.[0];
        return value || null;

    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            return null;
        }
        console.error(`[readSingleCell] Error reading cell: ${e.message}`, e);
        throw new Error(`セル読み込みに失敗しました: ${e.message}`);
    }
};

async function resizeSheetToFitData(spreadsheetId, sheetId, dataRowsCount, dataColsCount) {
    const targetRowCount = Math.max(dataRowsCount + 2, 5);
    const targetColCount = Math.max(dataColsCount + 1, 2);

    try {
        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: spreadsheetId,
            resource: {
                requests: [{
                    updateSheetProperties: {
                        properties: {
                            sheetId: sheetId,
                            gridProperties: {
                                rowCount: targetRowCount,
                                columnCount: targetColCount
                            }
                        },
                        fields: 'gridProperties(rowCount,columnCount)'
                    }
                }]
            }
        });
    } catch (e) {
        console.warn(`[Helper] Failed to resize sheet ${sheetId}: ${e.message}`);
    }
}

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

exports.getSpreadsheetIdFromUrl = getSpreadsheetIdFromUrl;
