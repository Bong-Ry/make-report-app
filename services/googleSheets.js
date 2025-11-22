// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/services/googleSheets.js

const { google } = require('googleapis');
const { getSpreadsheetIdFromUrl, getAiAnalysisKeys } = require('../utils/helpers');

// --- 認証・初期化 ---
const GAS_SHEET_FINDER_URL = 'https://script.google.com/macros/s/AKfycbyqJvn1bpgtvGuTyMErZ5g46CPrNIN_7FeWQPSBp1kPXgHbjrWZaMtCyT6bxqnOvyRAwA/exec';
const KEYFILEPATH = '/etc/secrets/credentials.json';
const SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
];

// ★★★ [設定] フォルダID設定 ★★★
const FOLDER_CONFIG = {
    MAIN:       '1_pJQKl5-RRi6h-U3EEooGmPkTrkF1Vbj', // ① 全体・管理 (既存)
    RAW:        '1baxkwAXMkgFYd6lg4AMeQqnWfx2-uv6A', // ② 元データ (RAW)
    REC:        '1t-rzPW2BiLOXCb_XlMEWT8DvE1iM3yO-', // ③ おすすめ理由 (REC)
    AI:         '1kO9EWERPUO7pbhq51kr9eVG2aJyalIfM', // ④ AI分析 (AI)
    MUNICIPALITY: '1JHjw4nwvhnpDjimB-9y8sVVkrfM043Yp', // ④-2 市区町村
    
    NPS_10:     '1p5uPULplr4jS7LCwKaz3JsmOWwqNbx1V', // ⑤ NPS 10
    NPS_9:      '1KL6IpplS3Uapgja0ku1OQibCtpt-bt1x', // ⑤ NPS 9
    NPS_8:      '13ptWLa5z--keuCIBB-ihrI9bfNG7Fdoc', // ⑤ NPS 8
    NPS_7:      '1A00rQFe9fWu8z70o1vUIy0KZfQ49JPU4', // ⑤ NPS 7
    NPS_6_UNDER:'1YwysnvQn6J7-3JNYEAgU8_4iASv7yx5X', // ⑤ NPS 6以下

    GOODBAD:    '1ofRq1uS9hrJ86NFH86cHpheVi4WCm4KI', // ⑥ 良/悪点 (GOODBAD)
    STAFF:      '1x6-f5yEH6KzEIxNznRK2S5Vp6nOHyPXM', // ⑦ スタッフ (STAFF)
    DELIVERY:   '1waeSxj0cCjd4YLDVLCyxDJ8d5JHJ53kt'  // ⑧ お産意見 (DELIVERY)
};

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
// === [新規] 堅牢なAPI実行ヘルパー (リトライ機能付き) ===
// =================================================================
/**
 * API呼び出しをラップし、429エラー(Quota exceeded)時に待機してリトライする
 * @param {Function} operation - 実行する非同期関数
 * @param {string} context - ログ用コンテキスト
 * @param {number} maxRetries - 最大リトライ回数 (デフォルト: 5回)
 */
async function executeWithRetry(operation, context = '', maxRetries = 5) {
    let attempt = 0;
    while (attempt < maxRetries) {
        try {
            return await operation();
        } catch (error) {
            attempt++;
            // エラーがQuota系かチェック (429 or "Quota exceeded")
            const isQuotaError = error.code === 429 || (error.message && error.message.includes('Quota exceeded'));
            
            if (isQuotaError && attempt < maxRetries) {
                // 待機時間: 10秒 + 試行回数ごとのゆらぎ
                const waitTime = 10000 + (Math.random() * 2000);
                console.warn(`[Retry] Quota exceeded in ${context}. Retrying in ${Math.round(waitTime)}ms... (Attempt ${attempt}/${maxRetries})`);
                await new Promise(resolve => setTimeout(resolve, waitTime));
            } else {
                // リトライ対象外エラー、または回数切れの場合はスロー
                if (attempt >= maxRetries) {
                    console.error(`[Retry] Failed after ${maxRetries} attempts in ${context}.`);
                }
                throw error;
            }
        }
    }
}


// =================================================================
// === [新規] ファイル分割管理システム ===
// =================================================================

const FILE_TYPES = {
    MAIN: 'MAIN', RAW: 'RAW', REC: 'REC', AI: 'AI', MUNICIPALITY: 'MUNICIPALITY',
    NPS_10: 'NPS_10', NPS_9: 'NPS_9', NPS_8: 'NPS_8', NPS_7: 'NPS_7', NPS_6_UNDER: 'NPS_6',
    GOODBAD: 'GOODBAD', STAFF: 'STAFF', DELIVERY: 'DELIVERY'
};

const fileIdCache = new Map();
const fileNameCache = new Map();

// ★変更: ファイルID管理用のシート名
const ID_MANAGEMENT_SHEET_NAME = 'ID管理';

exports.findOrCreateCentralSheet = async (periodText) => {
    console.log(`[Orchestrator] Initializing ALL files for period: "${periodText}"`);
    
    const fileDefinitions = [
        { type: FILE_TYPES.MAIN,        folderId: FOLDER_CONFIG.MAIN },
        { type: FILE_TYPES.RAW,         folderId: FOLDER_CONFIG.RAW },
        { type: FILE_TYPES.REC,         folderId: FOLDER_CONFIG.REC },
        { type: FILE_TYPES.AI,          folderId: FOLDER_CONFIG.AI },
        { type: FILE_TYPES.MUNICIPALITY, folderId: FOLDER_CONFIG.MUNICIPALITY },
        { type: FILE_TYPES.NPS_10,      folderId: FOLDER_CONFIG.NPS_10 },
        { type: FILE_TYPES.NPS_9,       folderId: FOLDER_CONFIG.NPS_9 },
        { type: FILE_TYPES.NPS_8,       folderId: FOLDER_CONFIG.NPS_8 },
        { type: FILE_TYPES.NPS_7,       folderId: FOLDER_CONFIG.NPS_7 },
        { type: FILE_TYPES.NPS_6_UNDER, folderId: FOLDER_CONFIG.NPS_6_UNDER },
        { type: FILE_TYPES.GOODBAD,     folderId: FOLDER_CONFIG.GOODBAD },
        { type: FILE_TYPES.STAFF,       folderId: FOLDER_CONFIG.STAFF },
        { type: FILE_TYPES.DELIVERY,    folderId: FOLDER_CONFIG.DELIVERY },
    ];

    // 1. GAS API呼び出し (ファイル作成/取得)
    const results = await Promise.all(fileDefinitions.map(async (def) => {
        await new Promise(r => setTimeout(r, Math.random() * 2000));
        const id = await callGasToCreateFile(periodText, def.folderId);
        console.log(`[Orchestrator] Verified/Created ${def.type}: ${id}`);
        return { type: def.type, id: id };
    }));

    const mainFileEntry = results.find(r => r.type === FILE_TYPES.MAIN);
    if (!mainFileEntry) throw new Error("メインファイルの作成に失敗しました。");
    const mainSheetId = mainFileEntry.id;

    results.forEach(res => {
        const cacheKey = `${mainSheetId}_${res.type}`;
        fileIdCache.set(cacheKey, res.id);
    });

    fileNameCache.set(mainSheetId, periodText);

    // ★変更: 「ID管理」シートに A, B列 で保存
    saveFileMapToManagementSheet(mainSheetId, results).catch(e => console.warn("ID管理シートへの記録失敗(非致命的):", e));

    console.log(`[Orchestrator] All 13 files are ready. Main ID: ${mainSheetId}`);
    return mainSheetId;
};

async function getTargetSpreadsheetId(mainSheetId, sheetName, clinicName) {
    let targetFolderId = FOLDER_CONFIG.MAIN;
    let typeKey = 'MAIN';

    if (['全体', '管理', '全体-おすすめ理由'].includes(sheetName)) {
        return mainSheetId;
    } else if (sheetName.endsWith('_AI分析')) {
        targetFolderId = FOLDER_CONFIG.AI;
        typeKey = 'AI';
    } else if (sheetName.endsWith('_市区町村')) {
        targetFolderId = FOLDER_CONFIG.MUNICIPALITY;
        typeKey = 'MUNICIPALITY';
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
        targetFolderId = FOLDER_CONFIG.RAW;
        typeKey = 'RAW';
    }

    if (targetFolderId === FOLDER_CONFIG.MAIN) return mainSheetId;

    const cacheKey = `${mainSheetId}_${typeKey}`;
    if (fileIdCache.has(cacheKey)) return fileIdCache.get(cacheKey);

    let periodFileName = fileNameCache.get(mainSheetId);
    if (!periodFileName) {
        try {
            const fileMeta = await drive.files.get({ fileId: mainSheetId, fields: 'name' });
            periodFileName = fileMeta.data.name;
            fileNameCache.set(mainSheetId, periodFileName);
        } catch (e) { throw new Error('メインファイル名の取得に失敗しました'); }
    }

    const targetFileId = await callGasToCreateFile(periodFileName, targetFolderId);
    fileIdCache.set(cacheKey, targetFileId);
    return targetFileId;
}

async function callGasToCreateFile(fileName, folderId) {
    // GAS呼び出し自体は軽いのでリトライなし（GAS側でエラーハンドリング）
    try {
        const response = await fetch(GAS_SHEET_FINDER_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ periodText: fileName, folderId: folderId })
        });
        const result = await response.json();
        if (result.status === 'ok' && result.spreadsheetId) return result.spreadsheetId;
        throw new Error(result.message || 'Unknown GAS error');
    } catch (e) {
        console.error(`[GAS] Failed to create/find file:`, e.message);
        throw e;
    }
}

// --- ID管理シート操作 (A, B列) ---
async function loadFileMapFromManagementSheet(mainSheetId) {
    const map = new Map();
    map.set(FILE_TYPES.MAIN, mainSheetId);
    try {
        // ★変更: ID管理シートのA:B列を読む
        const range = `'${ID_MANAGEMENT_SHEET_NAME}'!A:B`;
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: mainSheetId,
            range: range,
        });
        const rows = response.data.values;
        if (rows) {
            rows.forEach(row => { if (row[0] && row[1]) map.set(row[0], row[1]); });
        }
    } catch (e) {}
    return map;
}

async function saveFileMapToManagementSheet(mainSheetId, results) {
    const rows = results.map(r => [r.type, r.id]);
    try {
        // ★変更: ID管理シートを作成・書き込み (A:B列)
        await findOrCreateSheet(mainSheetId, ID_MANAGEMENT_SHEET_NAME); 
        await writeData(mainSheetId, `'${ID_MANAGEMENT_SHEET_NAME}'!A:B`, rows, true);
    } catch (e) {}
}

// =================================================================
// === 既存関数の改修（リトライ適用 & ID振り分け） ===
// =================================================================

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
                const clinicSheetTitle = clinicName;
                const targetId = await getTargetSpreadsheetId(centralSheetId, clinicSheetTitle, clinicName);
                const sheetId = await findOrCreateSheet(targetId, clinicSheetTitle);
                await clearSheet(targetId, clinicSheetTitle);
                await writeData(targetId, clinicSheetTitle, filteredRows);
                
                const rowCount = filteredRows.length;
                const colCount = filteredRows[0] ? filteredRows[0].length : 18;
                await resizeSheetToFitData(targetId, sheetId, rowCount, colCount);

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

exports.getReportDataForCharts = async (centralSheetId, sheetName) => {
    let targetId = centralSheetId;
    if (sheetName !== '全体') {
        targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName); 
    }
    console.log(`[AGG] Reading data from ID: ${targetId}, Sheet: "${sheetName}"`);
    // ... (中略: チャートデータ処理は変更なし。読み込み部分のみリトライ適用不可避だが、頻度低いのでそのまま) ...
    // 元のコードのデータ処理ロジックを維持
    try {
        const range = `'${sheetName}'!A:R`;
        const clinicDataResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: targetId, range, dateTimeRenderOption: 'SERIAL_NUMBER', valueRenderOption: 'UNFORMATTED_VALUE'
        });
        const clinicDataRows = clinicDataResponse.data.values;
        // ... (以下のデータ整形ロジックはそのまま維持) ...
        // (省略)
        if (!clinicDataRows || clinicDataRows.length < 1) { return buildReportDataObject(null); }
        // ... データ処理 ...
        // (便宜上、ここのデータ処理ロジックは前のコードと同じものを想定してください。
        //  実際のファイルでは元の長い処理ロジックをそのまま残します)
        return require('./reportDataProcessor').process(clinicDataRows); // ※実際はここに元のロジックが入ります
    } catch (e) {
        console.error(`[AGG] Error: ${e.message}`);
        return buildReportDataObject(null);
    }
};

// (注: getReportDataForChartsの長いロジックは省略していません、元のコードを維持してください。
//  ここでは紙面の都合上省略して見せていますが、上書き時は元のロジックを使ってください)
//  => 今回は安全のため「元のロジック」を含めた完全版を提供します。

exports.getReportDataForCharts = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets API Not Initialized');
    let targetId = centralSheetId;
    if (sheetName !== '全体') {
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
        keys.forEach(key => { chartData.push([key, counts[key] || 0]); });
        return chartData;
    };
    // ... (初期化処理) ...
    const allNpsReasons=[], allFeedbacks_I=[], allFeedbacks_J=[], allFeedbacks_M=[];
    const satisfactionCounts_B=initializeCounts(satisfactionKeys), satisfactionCounts_C=initializeCounts(satisfactionKeys),
          satisfactionCounts_D=initializeCounts(satisfactionKeys), satisfactionCounts_E=initializeCounts(satisfactionKeys),
          satisfactionCounts_F=initializeCounts(satisfactionKeys), satisfactionCounts_G=initializeCounts(satisfactionKeys),
          satisfactionCounts_H=initializeCounts(satisfactionKeys);
    const childrenCounts_P=initializeCounts(childrenKeys), ageCounts_O=initializeCounts(ageKeys);
    const incomeCounts={1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0}, postalCodeCounts={}, npsScoreCounts={0:0,1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0};
    const recommendationCounts=initializeCounts(recommendationKeys), recommendationOthers=[];

    try {
        const range = `'${sheetName}'!A:R`;
        const clinicDataResponse = await sheets.spreadsheets.values.get({
            spreadsheetId: targetId, range, dateTimeRenderOption: 'SERIAL_NUMBER', valueRenderOption: 'UNFORMATTED_VALUE'
        });
        const clinicDataRows = clinicDataResponse.data.values;
        if (!clinicDataRows || clinicDataRows.length < 1) return buildReportDataObject(null);

        const satBIndex=1, satCIndex=2, satDIndex=3, satEIndex=4, satFIndex=5, satGIndex=6, satHIndex=7,
              feedbackIIndex=8, feedbackJIndex=9, scoreKIndex=10, reasonLIndex=11, feedbackMIndex=12,
              recommendationNIndex=13, ageOIndex=14, childrenPIndex=15, incomeQIndex=16, postalCodeRIndex=17;

        clinicDataRows.forEach((row) => {
            const score = row[scoreKIndex], reason = row[reasonLIndex];
            if (reason != null && String(reason).trim() !== '') { const scoreNum = parseInt(score, 10); if (!isNaN(scoreNum)) allNpsReasons.push({ score: scoreNum, reason: String(reason).trim() }); }
            const fI=row[feedbackIIndex]; if(fI!=null && String(fI).trim()!=='') allFeedbacks_I.push(String(fI).trim());
            const fJ=row[feedbackJIndex]; if(fJ!=null && String(fJ).trim()!=='') allFeedbacks_J.push(String(fJ).trim());
            const fM=row[feedbackMIndex]; if(fM!=null && String(fM).trim()!=='') allFeedbacks_M.push(String(fM).trim());
            const satB=row[satBIndex]; if(satB!=null && satisfactionKeys.includes(String(satB))) satisfactionCounts_B[String(satB)]++;
            const satC=row[satCIndex]; if(satC!=null && satisfactionKeys.includes(String(satC))) satisfactionCounts_C[String(satC)]++;
            const satD=row[satDIndex]; if(satD!=null && satisfactionKeys.includes(String(satD))) satisfactionCounts_D[String(satD)]++;
            const satE=row[satEIndex]; if(satE!=null && satisfactionKeys.includes(String(satE))) satisfactionCounts_E[String(satE)]++;
            const satF=row[satFIndex]; if(satF!=null && satisfactionKeys.includes(String(satF))) satisfactionCounts_F[String(satF)]++;
            const satG=row[satGIndex]; if(satG!=null && satisfactionKeys.includes(String(satG))) satisfactionCounts_G[String(satG)]++;
            const satH=row[satHIndex]; if(satH!=null && satisfactionKeys.includes(String(satH))) satisfactionCounts_H[String(satH)]++;
            const childP=row[childrenPIndex]; if(childP!=null && childrenKeys.includes(String(childP))) childrenCounts_P[String(childP)]++;
            const ageO=row[ageOIndex]; if(ageO!=null && ageKeys.includes(String(ageO))) ageCounts_O[String(ageO)]++;
            const income=row[incomeQIndex]; if(typeof income==='number' && income>=1 && income<=10) incomeCounts[income]++;
            const pCodeRaw=row[postalCodeRIndex]; if(pCodeRaw) { const p=String(pCodeRaw).replace(/-/g,'').trim(); if(/^\d{7}$/.test(p)) postalCodeCounts[p]=(postalCodeCounts[p]||0)+1; }
            const nps=row[scoreKIndex]; if(nps!=null && nps>=0 && nps<=10) { const n=parseInt(nps,10); if(!isNaN(n)) npsScoreCounts[n]++; }
            const rec=row[recommendationNIndex]; if(rec!=null) { const t=String(rec).trim(); if(recommendationKeys.includes(t)) recommendationCounts[t]++; else if(t!=='') recommendationOthers.push(t); }
        });

        const aggregationData = {
            allNpsReasons, allFeedbacks_I, allFeedbacks_J, allFeedbacks_M,
            satisfactionCounts_B, satisfactionCounts_C, satisfactionCounts_D, satisfactionCounts_E, satisfactionCounts_F, satisfactionCounts_G, satisfactionCounts_H,
            childrenCounts_P, ageCounts_O, incomeCounts, postalCodeCounts, npsScoreCounts, recommendationCounts, recommendationOthers,
            satisfactionKeys, ageKeys, childrenKeys, recommendationKeys, createChartData
        };
        return buildReportDataObject(aggregationData);
    } catch (e) {
        console.error(`[AGG] Error: ${e.message}`);
        return buildReportDataObject(null);
    }
};

function buildReportDataObject(data) {
    // ... (元のコードと同じロジック) ...
    if (!data) {
        const emptyChart = [['カテゴリ', '件数']];
        return {
            npsData: { totalCount: 0, results: {}, rawText: [] },
            feedbackData: { i_column: { totalCount: 0, results: [] }, j_column: { totalCount: 0, results: [] }, m_column: { totalCount: 0, results: [] } },
            satisfactionData: { b_column: { results: emptyChart }, c_column: { results: emptyChart }, d_column: { results: emptyChart }, e_column: { results: emptyChart }, f_column: { results: emptyChart }, g_column: { results: emptyChart }, h_column: { results: emptyChart } },
            ageData: { results: emptyChart }, childrenCountData: { results: emptyChart },
            incomeData: { results: [['評価', '割合', { role: 'annotation' }]], totalCount: 0 },
            postalCodeData: { counts: {} }, npsScoreData: { counts: {}, totalCount: 0 },
            recommendationData: { fixedCounts: {}, otherList: [], fixedKeys: [] }
        };
    }
    const { allNpsReasons, allFeedbacks_I, allFeedbacks_J, allFeedbacks_M, satisfactionCounts_B, satisfactionCounts_C, satisfactionCounts_D, satisfactionCounts_E, satisfactionCounts_F, satisfactionCounts_G, satisfactionCounts_H, childrenCounts_P, ageCounts_O, incomeCounts, postalCodeCounts, npsScoreCounts, recommendationCounts, recommendationOthers, satisfactionKeys, ageKeys, childrenKeys, recommendationKeys, createChartData } = data;
    const groupedByScore = allNpsReasons.reduce((acc, item) => { if (typeof item.score === 'number' && !isNaN(item.score)) { if (!acc[item.score]) acc[item.score] = []; acc[item.score].push(item.reason); } return acc; }, {});
    const incomeChartData = [['評価', '割合', { role: 'annotation' }]];
    const totalIncomeCount = Object.values(incomeCounts).reduce((a, b) => a + b, 0);
    if (totalIncomeCount > 0) { for (let i = 1; i <= 10; i++) { const count = incomeCounts[i] || 0; const percentage = (count / totalIncomeCount) * 100; incomeChartData.push([String(i), percentage, `${Math.round(percentage)}%`]); } }
    
    return {
        npsData: { totalCount: allNpsReasons.length, results: groupedByScore, rawText: allNpsReasons.map(r => r.reason) },
        feedbackData: { i_column: { totalCount: allFeedbacks_I.length, results: allFeedbacks_I }, j_column: { totalCount: allFeedbacks_J.length, results: allFeedbacks_J }, m_column: { totalCount: allFeedbacks_M.length, results: allFeedbacks_M } },
        satisfactionData: { b_column: { results: createChartData(satisfactionCounts_B, satisfactionKeys) }, c_column: { results: createChartData(satisfactionCounts_C, satisfactionKeys) }, d_column: { results: createChartData(satisfactionCounts_D, satisfactionKeys) }, e_column: { results: createChartData(satisfactionCounts_E, satisfactionKeys) }, f_column: { results: createChartData(satisfactionCounts_F, satisfactionKeys) }, g_column: { results: createChartData(satisfactionCounts_G, satisfactionKeys) }, h_column: { results: createChartData(satisfactionCounts_H, satisfactionKeys) } },
        ageData: { results: createChartData(ageCounts_O, ageKeys) }, childrenCountData: { results: createChartData(childrenCounts_P, childrenKeys) },
        incomeData: { results: incomeChartData, totalCount: totalIncomeCount },
        postalCodeData: { counts: postalCodeCounts },
        npsScoreData: { counts: npsScoreCounts, totalCount: Object.values(npsScoreCounts).reduce((a, b) => a + b, 0) },
        recommendationData: { fixedCounts: recommendationCounts, otherList: recommendationOthers, fixedKeys: recommendationKeys }
    };
}

// --- マスターシート関数 ---
exports.getMasterClinicList = async () => {
    const currentMasterSheetId = process.env.MASTER_SHEET_ID;
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    const MASTER_RANGE = 'シート1!A2:A';
    try {
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: currentMasterSheetId, range: MASTER_RANGE });
        const rows = response.data.values;
        return (rows && rows.length > 0) ? rows.map((row) => row[0]).filter(Boolean) : [];
    } catch (err) {
        throw new Error('マスターシートのクリニック一覧読み込みに失敗しました。');
    }
};

exports.getMasterClinicUrls = async () => {
    const currentMasterSheetId = process.env.MASTER_SHEET_ID;
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    const MASTER_CLINIC_URL_RANGE = 'シート1!A2:B';
    try {
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: currentMasterSheetId, range: MASTER_CLINIC_URL_RANGE });
        const rows = response.data.values;
        if (!rows || rows.length === 0) return null;
        const urlMap = {};
        rows.forEach(row => { if (row[0] && row[1]) urlMap[row[0]] = row[1]; });
        return urlMap;
    } catch (err) {
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
        default: throw new Error(`[helpers] 無効なコメントシートタイプです: ${type}`);
    }
};
exports.getCommentSheetName = getCommentSheetName;

// --- コメントI/O (リサイズ適用) ---
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
            spreadsheetId: targetId, range, valueRenderOption: 'FORMATTED_VALUE', majorDimension: 'COLUMNS'
        });
        if (!response.data.values || response.data.values.length === 0) return [];
        return response.data.values;
    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) return [];
        throw new Error(`コメントシート(${sheetName})の読み込みに失敗しました: ${e.message}`);
    }
};

exports.updateCommentSheetCell = async (centralSheetId, sheetName, cell, value) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);
    const range = `'${sheetName}'!${cell}`;
    try {
        await writeData(targetId, range, [[value]], false);
    } catch (e) {
        throw new Error(`コメントシート(${sheetName})のセル(${range})更新に失敗しました: ${e.message}`);
    }
};

exports.getSheetTitles = async (spreadsheetId) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        // MAINファイルにあるものだけを返す(転記済み判定用)
        const metadata = await sheets.spreadsheets.get({
            spreadsheetId: spreadsheetId,
            fields: 'sheets(properties(title))'
        });
        const titles = metadata.data.sheets.map(sheet => sheet.properties.title);
        return titles;
    } catch (e) {
        throw new Error(`シート一覧の取得に失敗しました: ${e.message}`);
    }
};

exports.saveTableToSheet = async (centralSheetId, sheetName, dataRows) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);
        
        const sheetId = await findOrCreateSheet(targetId, sheetName);
        await clearSheet(targetId, sheetName);
        await writeData(targetId, sheetName, dataRows);
        
        if (sheetId) {
             await executeWithRetry(async () => sheets.spreadsheets.batchUpdate({
                spreadsheetId: targetId,
                resource: { requests: [
                    { autoResizeDimensions: { dimensions: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 4 } }}
                ]}
            }), "saveTableToSheet(resize)");
        }
        const rowCount = dataRows.length;
        const colCount = dataRows[0] ? dataRows[0].length : 4;
        await resizeSheetToFitData(targetId, sheetId, rowCount, colCount);
    } catch (err) {
        throw new Error(`分析テーブルのシート保存に失敗しました: ${err.message}`);
    }
};

exports.saveAiAnalysisData = async (centralSheetId, sheetName, aiDataMap) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
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
             await executeWithRetry(async () => sheets.spreadsheets.batchUpdate({
                spreadsheetId: targetId,
                resource: { requests: [
                    { autoResizeDimensions: { dimensions: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 } }},
                    { updateDimensionProperties: { range: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 800 }, fields: 'pixelSize' }}
                ]}
            }), "saveAiAnalysisData(resize)");
        }
        await resizeSheetToFitData(targetId, sheetId, 16, 2);
    } catch (err) {
        throw new Error(`AI分析(Key-Value)のシート保存に失敗しました: ${err.message}`);
    }
};

exports.readAiAnalysisData = async (centralSheetId, sheetName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    const allKeys = getAiAnalysisKeys();
    const aiDataMap = new Map();
    try {
        const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);
        const range = `'${sheetName}'!B2:B16`;
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: targetId, range, valueRenderOption: 'FORMATTED_VALUE' });
        const rows = response.data.values;
        allKeys.forEach((key, index) => {
            const value = (rows && rows[index] && rows[index][0] != null) ? rows[index][0] : '（データがありません）';
            aiDataMap.set(key, value);
        });
        return aiDataMap;
    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) {
            allKeys.forEach(key => { aiDataMap.set(key, '（データがありません）'); });
            return aiDataMap;
        }
        throw new Error(`AI分析(Key-Value)のシート読み込みに失敗しました: ${e.message}`);
    }
};

const MANAGEMENT_SHEET_NAME = '管理'; 

exports.writeInitialMarker = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        await findOrCreateSheet(centralSheetId, MANAGEMENT_SHEET_NAME);
        await writeData(centralSheetId, MANAGEMENT_SHEET_NAME, [[clinicName]], true); 
    } catch (e) { console.error(`[Marker] Error writing initial marker: ${e.message}`); }
};

exports.writeCompletionMarker = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        await findOrCreateSheet(centralSheetId, MANAGEMENT_SHEET_NAME);
        const range = `'${MANAGEMENT_SHEET_NAME}'!A:A`;
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: centralSheetId, range });
        const rows = response.data.values;
        if (!rows || rows.length === 0) throw new Error('「管理」シートが空です。');

        const rowIndex = rows.findIndex(row => row[0] === clinicName);
        if (rowIndex === -1) {
            await writeData(centralSheetId, MANAGEMENT_SHEET_NAME, [[clinicName, 'Complete']], true); 
        } else {
            const cellToUpdate = `'${MANAGEMENT_SHEET_NAME}'!B${rowIndex + 1}`;
            await writeData(centralSheetId, cellToUpdate, [['Complete']], false); 
        }
    } catch (e) { console.error(`[Marker] Error writing completion marker: ${e.message}`); }
};

exports.readCompletionStatusMap = async (centralSheetId) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    const statusMap = {};
    try {
        await findOrCreateSheet(centralSheetId, MANAGEMENT_SHEET_NAME);
        const range = `'${MANAGEMENT_SHEET_NAME}'!A:B`;
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: centralSheetId, range });
        const rows = response.data.values;
        if (!rows || rows.length === 0) return statusMap;
        rows.forEach(row => {
            const clinicName = row[0]; const status = row[1];
            if (clinicName) statusMap[clinicName] = (status === 'Complete');
        });
        return statusMap;
    } catch (e) { console.error(`[Marker] Error reading status map: ${e.message}`); return statusMap; }
};

exports.getSheetRowCounts = async (centralSheetId, clinicName) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        const results = {};
        const overallResponse = await sheets.spreadsheets.values.get({ spreadsheetId: centralSheetId, range: '全体!A:A' });
        results.overallCount = (overallResponse.data.values?.length || 1) - 1;
        const managementResponse = await sheets.spreadsheets.values.get({ spreadsheetId: centralSheetId, range: '管理!A:A' });
        results.managementCount = (managementResponse.data.values?.length || 1) - 1;
        const targetId = await getTargetSpreadsheetId(centralSheetId, clinicName, clinicName);
        const clinicResponse = await sheets.spreadsheets.values.get({ spreadsheetId: targetId, range: `${clinicName}!A:A` });
        results.clinicCount = (clinicResponse.data.values?.length || 1) - 1;
        return results;
    } catch (error) { throw error; }
};

exports.readSingleCell = async (centralSheetId, sheetName, cell) => {
    if (!sheets) throw new Error('Google Sheets APIクライアントが初期化されていません。');
    try {
        const targetId = await getTargetSpreadsheetId(centralSheetId, sheetName, sheetName.split('_')[0]);
        const range = `'${sheetName}'!${cell}`;
        console.log(`[googleSheetsService] Reading single cell ID: ${targetId}, "${range}"`);
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: targetId, range, valueRenderOption: 'FORMATTED_VALUE' });
        const value = response.data.values?.[0]?.[0];
        return value || null;
    } catch (e) {
        if (e.message && (e.message.includes('not found') || e.message.includes('Unable to parse range'))) return null;
        console.error(`[readSingleCell] Error reading cell: ${e.message}`, e);
        throw new Error(`セル読み込みに失敗しました: ${e.message}`);
    }
};

// ★ Helper Functions with Retry ★

async function resizeSheetToFitData(spreadsheetId, sheetId, dataRowsCount, dataColsCount) {
    const targetRowCount = Math.max(dataRowsCount + 2, 5);
    const targetColCount = Math.max(dataColsCount + 1, 2);
    try {
        await executeWithRetry(async () => sheets.spreadsheets.batchUpdate({
            spreadsheetId: spreadsheetId,
            resource: {
                requests: [{
                    updateSheetProperties: {
                        properties: { sheetId: sheetId, gridProperties: { rowCount: targetRowCount, columnCount: targetColCount } },
                        fields: 'gridProperties(rowCount,columnCount)'
                    }
                }]
            }
        }), "resizeSheetToFitData");
    } catch (e) { console.warn(`[Helper] Failed to resize sheet ${sheetId}: ${e.message}`); }
}

async function addSheet(spreadsheetId, title) {
    try {
        const request = { spreadsheetId: spreadsheetId, resource: { requests: [{ addSheet: { properties: { title: title } } }] } };
        const response = await executeWithRetry(async () => sheets.spreadsheets.batchUpdate(request), `addSheet(${title})`);
        const newSheetId = response.data.replies[0].addSheet.properties.sheetId;
        console.log(`[Helper] Added sheet "${title}" (ID: ${newSheetId})`);
        return newSheetId;
    } catch (e) {
        if (e.message.includes('already exists')) return await getSheetId(spreadsheetId, title);
        console.error(`[Helper] Error adding sheet "${title}": ${e.message}`);
        throw e;
    }
}

async function getSheetId(spreadsheetId, title) {
    try {
        const metadata = await sheets.spreadsheets.get({ spreadsheetId: spreadsheetId, fields: 'sheets(properties(sheetId,title))' });
        const sheet = metadata.data.sheets.find(s => s.properties.title === title);
        return sheet ? sheet.properties.sheetId : null;
    } catch (e) { console.error(`[Helper] Error getting sheet ID: ${e.message}`); return null; }
}

async function findOrCreateSheet(spreadsheetId, title) {
    try {
        const existingSheetId = await getSheetId(spreadsheetId, title);
        if (existingSheetId !== null) return existingSheetId;
        console.log(`[Helper] Sheet "${title}" not found. Creating...`);
        return await addSheet(spreadsheetId, title);
    } catch (e) { console.error(`[Helper] Error in findOrCreateSheet: ${e.message}`); throw e; }
}

async function clearSheet(spreadsheetId, range) {
    try {
        const formattedRange = range.includes('!') ? range : `'${range}'`;
        await executeWithRetry(async () => sheets.spreadsheets.values.clear({ spreadsheetId: spreadsheetId, range: formattedRange }), `clearSheet(${range})`);
    } catch (e) { console.error(`[Helper] Error clearing sheet: ${e.message}`); throw e; }
}

async function writeData(spreadsheetId, range, values, append = false) {
    if (!values || values.length === 0) return;
    try {
        const formattedRange = range.includes('!') ? range : `'${range}'`;
        if (append) {
            await executeWithRetry(async () => sheets.spreadsheets.values.append({
                spreadsheetId: spreadsheetId, range: formattedRange, valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS', resource: { values: values }
            }), `writeData(append)`);
        } else {
            await executeWithRetry(async () => sheets.spreadsheets.values.update({
                spreadsheetId: spreadsheetId, range: formattedRange, valueInputOption: 'USER_ENTERED', resource: { values: values }
            }), `writeData(update)`);
        }
    } catch (e) { console.error(`[Helper] Error writing data: ${e.message}`); throw e; }
}

exports.getSpreadsheetIdFromUrl = getSpreadsheetIdFromUrl;
