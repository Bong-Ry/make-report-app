// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/controllers/reportController.js

const googleSheetsService = require('../services/googleSheets');
const aiAnalysisService = require('../aiAnalysisService');
const { getCommentSheetName } = require('../services/googleSheets');

// p-limit (並列実行数制限) の読み込み
const pLimit = async () => (await import('p-limit')).default;
let limit; 
(async () => {
    try {
        const PLimit = await pLimit();
        limit = PLimit(2); 
    } catch (e) {
        console.error("p-limit の動的インポートに失敗しました。", e);
        const PLimitSync = require('p-limit');
        limit = PLimitSync(2);
    }
})();


// --- クリニック一覧取得 ---
exports.getClinicList = async (req, res) => {
    console.log('GET /api/getClinicList called');
    try {
        const clinics = await googleSheetsService.getMasterClinicList();
        console.log('[/api/getClinicList] Fetched clinics:', clinics);
        res.json(clinics);
    } catch (err) {
        console.error('[/api/getClinicList] Error getting clinic list:', err);
        res.status(500).send(err.message || 'マスターシートの読み込みに失敗しました。');
    }
};

// --- ▼▼▼ [修正] 転記状況の確認 (完了したものだけを返す) ▼▼▼ ---
exports.getTransferredList = async (req, res) => {
    const { centralSheetId } = req.body;
    console.log(`POST /api/getTransferredList called for: ${centralSheetId}`);

    if (!centralSheetId) {
        return res.status(400).send('Invalid request: centralSheetId required.');
    }
    
    try {
        // 1. 「管理」シートの状態マップを取得
        const completionMap = await googleSheetsService.readCompletionStatusMap(centralSheetId);
        
        // 2. ★変更: ステータスが true (完了) のクリニックだけをリスト化する
        const transferredClinics = Object.keys(completionMap).filter(name => completionMap[name] === true);
        
        // AI完了状況用（こちらは全件チェック）
        const masterClinics = await googleSheetsService.getMasterClinicList();
        const aiCompletionStatus = {};

        for (const clinicName of masterClinics) {
            if (completionMap.hasOwnProperty(clinicName)) {
                aiCompletionStatus[clinicName] = (completionMap[clinicName] === true);
            } else {
                aiCompletionStatus[clinicName] = false;
            }
        }
        
        console.log(`[/api/getTransferredList] Completed (Transferred) count: ${transferredClinics.length}`);
        
        res.json({ 
            sheetTitles: transferredClinics, // 完了済みリストのみ返す
            aiCompletionStatus: aiCompletionStatus
        });

    } catch (err) {
        console.error('[/api/getTransferredList] Error:', err);
        res.status(500).send(err.message || '転記済みシート一覧の取得に失敗しました。');
    }
};

// --- シート行数を取得 ---
exports.getSheetRowCounts = async (req, res) => {
    const { centralSheetId, clinicName } = req.body;
    console.log(`POST /api/getSheetRowCounts called for: ${clinicName} in ${centralSheetId}`);

    if (!centralSheetId || !clinicName) {
        return res.status(400).send('Invalid request: centralSheetId and clinicName required.');
    }

    try {
        const rowCounts = await googleSheetsService.getSheetRowCounts(centralSheetId, clinicName);
        res.json(rowCounts);
    } catch (err) {
        console.error('[/api/getSheetRowCounts] Error:', err);
        res.status(500).send(err.message || 'シート行数の取得に失敗しました。');
    }
};

// --- ファイル作成/特定 ---
exports.findOrCreateSheet = async (req, res) => {
    const { periodText } = req.body;
    console.log(`POST /api/findOrCreateSheet called for: ${periodText}`);
    
    if (!periodText) {
        return res.status(400).send('Invalid request: periodText required.');
    }
    
    try {
        const centralSheetId = await googleSheetsService.findOrCreateCentralSheet(periodText);
        res.json({ centralSheetId: centralSheetId, periodText: periodText });
    } catch (err) {
        console.error('[/api/findOrCreateSheet] Error:', err);
        res.status(500).send(err.message || '集計スプレッドシートの検索または作成に失敗しました。');
    }
};

// --- データ転記トリガー ---
exports.getReportData = async (req, res) => {
    const { period, selectedClinics, centralSheetId } = req.body;
    console.log('POST /api/getReportData (ETL Trigger) called');

    if (!period || !selectedClinics || !Array.isArray(selectedClinics) || selectedClinics.length === 0 || !centralSheetId) {
        return res.status(400).send('Invalid request: period, selectedClinics, and centralSheetId required.');
    }
    
    if (selectedClinics.length > 10) {
        return res.status(400).send('Invalid request: 一度に処理できるのは10件までです。');
    }

    try {
        const masterClinicUrls = await googleSheetsService.getMasterClinicUrls();
        if (!masterClinicUrls) {
            return res.json({});
        }
        
        const clinicUrls = {};
        Object.entries(masterClinicUrls).forEach(([clinicName, sheetUrl]) => {
            if (selectedClinics.includes(clinicName) && sheetUrl) {
                clinicUrls[clinicName] = sheetUrl;
            }
        });
        
        if (Object.keys(clinicUrls).length === 0) {
            return res.json({ status: 'ok', processed: [] });
        }

        const processedClinics = await googleSheetsService.fetchAndAggregateReportData(clinicUrls, period, centralSheetId);
        
        if (processedClinics.length > 0) {
            console.log(`[/api/getReportData] Triggering background AI analysis for ${processedClinics.length} clinics...`);
            runBackgroundAiAnalysis(centralSheetId, processedClinics);
        }
        
        res.json({ status: 'ok', processed: processedClinics });

    } catch (err) {
        console.error('[/api/getReportData] Error in getReportData (ETL) controller:', err);
        res.status(500).send(err.message || 'レポートデータ転記中にエラーが発生しました。');
    }
};

/**
 * AI分析 + コメント保存 をバックグラウンドで実行
 */
async function runBackgroundAiAnalysis(centralSheetId, clinicNames) {
    if (!limit) {
        const PLimitSync = require('p-limit');
        limit = PLimitSync(2);
    }
    
    console.log(`[BG-AI] Background task started for ${clinicNames.join(', ')}`);
    
    const analysisTasksDefinition = [
        { type: 'AI_ALL', func: aiAnalysisService.runAllAiAnalysesAndSave },
        { type: 'MUNICIPALITY', func: aiAnalysisService.runAndSaveMunicipalityAnalysis },
        { type: 'RECOMMENDATION', func: aiAnalysisService.runAndSaveRecommendationAnalysis }
    ];
    
    for (const clinicName of clinicNames) {
        console.log(`[BG-AI] Starting all analyses for ${clinicName} (Concurrency: 2)...`);
        
        try {
            // 1. 集計データを取得
            console.log(`[BG-AI] Fetching aggregated data for ${clinicName}...`);
            const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);

            // 2. 初期マーカー書き込み
            await googleSheetsService.writeInitialMarker(centralSheetId, clinicName);

            // 3. 分析タスク(3種)
            const analysisPromises = analysisTasksDefinition.map(task => {
                return limit(async () => {
                    try {
                        console.log(`[BG-AI] Running ${clinicName} - ${task.type}...`);
                        await task.func(centralSheetId, clinicName);
                        console.log(`[BG-AI] SUCCESS: ${clinicName} - ${task.type}`);
                        return { type: task.type, status: 'success' };
                    } catch (e) {
                        if (e.message && e.message.includes('データが0件')) {
                            console.log(`[BG-AI] SKIP: ${clinicName} - ${task.type} (No data)`);
                            return { type: task.type, status: 'skipped' };
                        }
                        console.error(`[BG-AI] FAILED: ${clinicName} - ${task.type}: ${e.message}`);
                        return { type: task.type, status: 'failed', error: e.message };
                    }
                });
            }); 
            
            // 4. コメント保存タスク
            const commentTasks = [];
            
            // NPS
            const npsResults = reportData.npsData.results;
            if (npsResults) {
                const nps10 = npsResults['10'] || [];
                const nps9 = npsResults['9'] || [];
                const nps8 = npsResults['8'] || [];
                const nps7 = npsResults['7'] || [];
                const nps6_under = Object.keys(npsResults)
                    .map(Number)
                    .filter(score => score <= 6)
                    .flatMap(score => npsResults[score] || []);

                if (nps10.length > 0) commentTasks.push({ type: 'L_10', comments: nps10 });
                if (nps9.length > 0) commentTasks.push({ type: 'L_9', comments: nps9 });
                if (nps8.length > 0) commentTasks.push({ type: 'L_8', comments: nps8 });
                if (nps7.length > 0) commentTasks.push({ type: 'L_7', comments: nps7 });
                if (nps6_under.length > 0) commentTasks.push({ type: 'L_6_under', comments: nps6_under });
            }
            
            // Feedback
            const feedbackI = reportData.feedbackData.i_column.results;
            const feedbackJ = reportData.feedbackData.j_column.results;
            const feedbackM = reportData.feedbackData.m_column.results;
            
            if (feedbackI.length > 0) commentTasks.push({ type: 'I', comments: feedbackI });
            if (feedbackJ.length > 0) commentTasks.push({ type: 'J', comments: feedbackJ });
            if (feedbackM.length > 0) commentTasks.push({ type: 'M', comments: feedbackM });

            const commentPromises = commentTasks.map(task => {
                return limit(async () => {
                    const sheetName = getCommentSheetName(clinicName, task.type);
                    try {
                        // ★★★ [修正] 待機時間を 3000ms に延長してAPI制限を回避 ★★★
                        await new Promise(resolve => setTimeout(resolve, 3000));

                        console.log(`[BG-AI] Running ${clinicName} - COMMENTS (${task.type}) -> ${sheetName}`);
                        await googleSheetsService.saveCommentsToSheet(centralSheetId, sheetName, task.comments);
                        console.log(`[BG-AI] SUCCESS: ${clinicName} - COMMENTS (${task.type})`);
                        return { type: `COMMENTS_${task.type}`, status: 'success' };
                    } catch (e) {
                        console.error(`[BG-AI] FAILED: ${clinicName} - COMMENTS (${task.type}): ${e.message}`);
                        return { type: `COMMENTS_${task.type}`, status: 'failed', error: e.message };
                    }
                });
            });
            
            // 5. 完了待ち
            const allPromises = [...analysisPromises, ...commentPromises];
            const results = await Promise.all(allPromises);
            
            const hasFailed = results.some(r => r.status === 'failed');
            
            if (hasFailed) {
                 console.error(`[BG-AI] COMPLETED (WITH FAILURES) for ${clinicName}. Completion marker will NOT be set.`);
            } else {
                console.log(`[BG-AI] COMPLETED (SUCCESS/SKIP) all tasks for ${clinicName}.`);
                await googleSheetsService.writeCompletionMarker(centralSheetId, clinicName);
                console.log(`[BG-AI] Set COMPLETION MARKER for ${clinicName}.`);
            }

        } catch (e) {
            console.error(`[BG-AI] FATAL ERROR during pre-fetch or execution for ${clinicName}: ${e.message}`);
        }
    }
    console.log('[BG-AI] All background tasks finished.');
}


// --- チャートデータ取得 ---
exports.getChartData = async (req, res) => {
    const { centralSheetId, sheetName } = req.body;
    console.log(`POST /api/getChartData called for SheetID: ${centralSheetId}, Tab: "${sheetName}"`);
    
    if (!centralSheetId || !sheetName) {
        return res.status(400).send('Invalid request: centralSheetId and sheetName required.');
    }
    
    try {
        const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, sheetName);
        res.json(reportData);
    } catch (err) {
        console.error('[/api/getChartData] Error:', err);
        res.status(500).send(err.message || '集計データ(グラフ用)の取得に失敗しました。');
    }
};
