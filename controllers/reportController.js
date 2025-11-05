// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/controllers/reportController.js

const googleSheetsService = require('../services/googleSheets');
const pdfGeneratorService = require('../services/pdfGenerator');
const aiAnalysisService = require('../aiAnalysisService');
const googleSlidesService = require('../services/googleSlidesService');

// ▼▼▼ [新規] p-limit をインポート (package.json に "p-limit": "^5.0.0" を追加し、npm install が必要) ▼▼▼
const pLimit = async () => (await import('p-limit')).default;
let limit; // p-limit のインスタンス
(async () => {
    try {
        const PLimit = await pLimit();
        // ▼▼▼ [変更] ご要望に基づき、並列実行数を 2 に制限 ▼▼▼
        limit = PLimit(2); 
    } catch (e) {
        console.error("p-limit の動的インポートに失敗しました。", e);
        // フォールバックとして、逐次実行(1)に設定
        const PLimitSync = require('p-limit');
        limit = PLimitSync(1);
    }
})();
// ▲▲▲


// --- (変更なし) クリニック一覧取得 ---
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

// =================================================================
// === ▼▼▼ [変更] 転記状況の確認 (「押せない」問題の修正) ▼▼▼ ===
// =================================================================
/**
 * [変更] 集計スプレッドシート内の全タブ名と、*バックグラウンド分析* の完了状況を取得する
 */
exports.getTransferredList = async (req, res) => {
    const { centralSheetId } = req.body;
    console.log(`POST /api/getTransferredList called for: ${centralSheetId}`);

    if (!centralSheetId) {
        return res.status(400).send('Invalid request: centralSheetId required.');
    }
    
    try {
        // 1. 全シート名を取得 (既存)
        const sheetTitles = await googleSheetsService.getSheetTitles(centralSheetId);
        const sheetTitlesSet = new Set(sheetTitles);

        // 2. マスターからクリニック一覧を取得
        const masterClinics = await googleSheetsService.getMasterClinicList();

        // 3. ▼▼▼ [変更] 分析完了ステータスを計算 ▼▼▼
        const aiCompletionStatus = {};

        // (Promise.all で全クリニックのマーカー読み取りを並列化)
        await Promise.all(masterClinics.map(async (clinicName) => {
            // このクリニックが転記済み (タブが存在する) か？
            if (sheetTitlesSet.has(clinicName)) {
                
                // [変更] 多数の分析シートの有無をチェックする代わりに、
                //       単一分析シートの「完了マーカー」を読み取る (ご要望: 揃ってい流のに押せない)
                const isComplete = await googleSheetsService.readCompletionMarker(centralSheetId, clinicName);
                aiCompletionStatus[clinicName] = isComplete;
                
            } else {
                aiCompletionStatus[clinicName] = false; // 転記自体されていない
            }
        }));
        
        console.log(`[/api/getTransferredList] AI Status:`, aiCompletionStatus);
        
        // 4. シート名リストとAI完了ステータスの両方を返す
        res.json({ 
            sheetTitles: sheetTitles,
            aiCompletionStatus: aiCompletionStatus // { "クリニックA": true, "クリニックB": false, ... }
        });

    } catch (err) {
        console.error('[/api/getTransferredList] Error:', err);
        res.status(500).send(err.message || '転記済みシート一覧の取得に失敗しました。');
    }
};


// =================================================================
// === (変更なし) findOrCreateSheet (1/3) ===
// =================================================================
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

// =================================================================
// === (変更なし) getReportData (ETL Trigger) (2/3) ===
// =================================================================
exports.getReportData = async (req, res) => {
    const { period, selectedClinics, centralSheetId } = req.body;
    console.log('POST /api/getReportData (ETL Trigger) called');
    console.log('[/api/getReportData] Request Body:', req.body);

    if (!period || !selectedClinics || !Array.isArray(selectedClinics) || selectedClinics.length === 0 || !centralSheetId) {
        console.error('[/api/getReportData] Invalid request body:', req.body);
        return res.status(400).send('Invalid request: period, selectedClinics, and centralSheetId required.');
    }
    
    if (selectedClinics.length > 10) {
         console.error('[/api/getReportData] Too many clinics selected:', selectedClinics.length);
        return res.status(400).send('Invalid request: 一度に処理できるのは10件までです。');
    }

    try {
        // 1. マスターからURLマップを取得 (変更なし)
        const masterClinicUrls = await googleSheetsService.getMasterClinicUrls();
        if (!masterClinicUrls) {
            console.log('[/api/getReportData] No clinic/URL data found in master sheet.');
            return res.json({});
        }
        
        const clinicUrls = {};
        Object.entries(masterClinicUrls).forEach(([clinicName, sheetUrl]) => {
            if (selectedClinics.includes(clinicName) && sheetUrl) {
                clinicUrls[clinicName] = sheetUrl;
            }
        });
        
        console.log('[/api/getReportData] Target Source Sheet URLs:', clinicUrls);
        if (Object.keys(clinicUrls).length === 0) {
            console.warn('[/api/getReportData] No valid source sheet URLs found for selected clinics.');
            return res.json({ status: 'ok', processed: [] });
        }

        // 2. データを集計スプシに「転記」する (ETL実行)
        const processedClinics = await googleSheetsService.fetchAndAggregateReportData(clinicUrls, period, centralSheetId);
        
        // 3. [変更なし] AI分析をバックグラウンドで実行 (await しない)
        if (processedClinics.length > 0) {
            console.log(`[/api/getReportData] Triggering background AI analysis for ${processedClinics.length} clinics...`);
            
            // AI分析の実行をトリガーする。
            runBackgroundAiAnalysis(centralSheetId, processedClinics);
            
        }
        
        console.log('[/api/getReportData] Finished ETL process. Responding to client.');
        // 4. AI分析の完了を待たずに、すぐにフロントエンドに応答を返す
        res.json({ status: 'ok', processed: processedClinics });

    } catch (err) {
        console.error('[/api/getReportData] Error in getReportData (ETL) controller:', err);
        res.status(500).send(err.message || 'レポートデータ転記中にエラーが発生しました。');
    }
};

/**
 * [変更] AI分析をバックグラウンドで実行（並列数 2）
 * (ご要望: 市区町村・おすすめ理由 も / 最大2つ)
 * @param {string} centralSheetId 
 * @param {string[]} clinicNames - 処理対象のクリニック名リスト
 */
async function runBackgroundAiAnalysis(centralSheetId, clinicNames) {
    if (!limit) {
        console.error('[BG-AI] p-limit is not initialized! Running sequentially.');
        // p-limit の動的インポートがまだ完了していない場合のフォールバック
        const PLimit = (await pLimit()).default;
        limit = PLimit(2); // ここで再設定
    }
    
    console.log(`[BG-AI] Background task started for ${clinicNames.join(', ')}`);
    
    // ▼▼▼ [変更] タスクを7種類に増やす (ご要望: 市区町村・おすすめ理由 も) ▼▼▼
    const analysisTasks = [
        { type: 'L', func: aiAnalysisService.runAndSaveAnalysis },
        { type: 'I_good', func: aiAnalysisService.runAndSaveAnalysis },
        { type: 'I_bad', func: aiAnalysisService.runAndSaveAnalysis },
        { type: 'J', func: aiAnalysisService.runAndSaveAnalysis },
        { type: 'M', func: aiAnalysisService.runAndSaveAnalysis },
        { type: 'MUNICIPALITY', func: aiAnalysisService.runAndSaveMunicipalityAnalysis },
        { type: 'RECOMMENDATION', func: aiAnalysisService.runAndSaveRecommendationAnalysis }
    ];
    
    // 転記されたクリニックごとにループ
    for (const clinicName of clinicNames) {
        console.log(`[BG-AI] Starting all 7 analyses for ${clinicName} (Concurrency: 2)...`);
        
        try {
            // ▼▼▼ [変更] 7種類のタスクを並列数 2 で実行 (ご要望: 最大2つ) ▼▼▼
            const promises = analysisTasks.map(task => {
                // p-limit (limit) でラップして呼び出す
                return limit(async () => {
                    try {
                        console.log(`[BG-AI] Running ${clinicName} - ${task.type}...`);
                        
                        if (task.type === 'MUNICIPALITY' || task.type === 'RECOMMENDATION') {
                            // 市区町村・おすすめ理由
                            await task.func(centralSheetId, clinicName);
                        } else {
                            // AI 5種
                            await task.func(centralSheetId, clinicName, task.type);
                        }
                        
                        console.log(`[BG-AI] SUCCESS: ${clinicName} - ${task.type}`);
                        return { type: task.type, status: 'success' };
                    } catch (e) {
                        if (e.message && e.message.includes('テキストデータが0件')) {
                            console.log(`[BG-AI] SKIP: ${clinicName} - ${task.type} (No data)`);
                            return { type: task.type, status: 'skipped' };
                        } else if (e.message && e.message.includes('郵便番号データが0件')) {
                             console.log(`[BG-AI] SKIP: ${clinicName} - ${task.type} (No data)`);
                            return { type: task.type, status: 'skipped' };
                        } else if (e.message && e.message.includes('おすすめ理由データが0件')) {
                             console.log(`[BG-AI] SKIP: ${clinicName} - ${task.type} (No data)`);
                            return { type: task.type, status: 'skipped' };
                        } else {
                            console.error(`[BG-AI] FAILED: ${clinicName} - ${task.type}: ${e.message}`);
                            return { type: task.type, status: 'failed', error: e.message };
                        }
                    }
                });
            }); // promises.map 終了
            
            // このクリニックの全タスク(7件)の完了を待つ
            const results = await Promise.all(promises);
            
            // 1件でも失敗(failed)があったか？ (skipped は除く)
            const hasFailed = results.some(r => r.status === 'failed');
            
            if (hasFailed) {
                 console.error(`[BG-AI] COMPLETED (WITH FAILURES) for ${clinicName}. Completion marker will NOT be set.`);
            } else {
                // 7件すべてが success または skipped の場合
                console.log(`[BG-AI] COMPLETED (SUCCESS/SKIP) all 7 analyses for ${clinicName}.`);
                // ▼▼▼ [新規] 完了マーカーを書き込む (ご要望: 揃ってい流のに押せない) ▼▼▼
                await googleSheetsService.saveToAnalysisSheet(centralSheetId, clinicName, 'COMPLETION_MARKER', 'COMPLETED');
                console.log(`[BG-AI] Set COMPLETION MARKER for ${clinicName}.`);
            }

        } catch (e) {
            // (p-limit 自体のエラーなど、予期せぬエラー)
            console.error(`[BG-AI] FATAL ERROR during parallel execution for ${clinicName}: ${e.message}`);
        }
    }
    console.log('[BG-AI] All background tasks finished.');
}


// =================================================================
// === (変更なし) getChartData (3/3) ===
// =================================================================
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

// =================================================================
// === (変更なし) PDF生成 ===
// =================================================================
exports.generatePdf = async (req, res) => {
    console.log("POST /generate-pdf called");
    const { clinicName, periodText, centralSheetId } = req.body;
    
    if (!clinicName || !periodText || !centralSheetId) {
        console.error('[/generate-pdf] Missing data:', { clinicName: !!clinicName, periodText: !!periodText, centralSheetId: !!centralSheetId });
        return res.status(400).send('PDF生成に必要なデータ(clinicName, periodText, centralSheetId)が不足');
    }

    try {
        console.log(`[/generate-pdf] Fetching data for PDF from sheet: "${clinicName}"`);
        const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
        
        const pdfBuffer = await pdfGeneratorService.generatePdfFromData(clinicName, periodText, reportData);

        res.contentType('application/pdf');
        const fileName = `${clinicName}_${periodText.replace(/～/g, '-')}_レポート.pdf`;
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
        res.send(pdfBuffer);
        console.log("[/generate-pdf] PDF sent successfully.");

    } catch (error) {
        console.error('[/generate-pdf] PDF generation failed:', error);
        res.status(500).send(`PDF生成失敗: ${error.message}`);
    }
};

// =================================================================
// === (変更なし) スライド生成 ===
// =================================================================
exports.generateSlide = async (req, res) => {
    console.log("POST /api/generateSlide called");
    const { clinicName, centralSheetId, period, periodText } = req.body;

    if (!clinicName || !centralSheetId || !period || !periodText) {
        console.error('[/api/generateSlide] Missing data:', { 
            clinicName: !!clinicName, 
            centralSheetId: !!centralSheetId,
            period: !!period,
            periodText: !!periodText
        });
        return res.status(400).send('スライド生成に必要なデータ(clinicName, centralSheetId, period, periodText)が不足しています。');
    }

    try {
        console.log(`[/api/generateSlide] Starting slide generation for: ${clinicName}`);
        
        // ▼▼▼ [変更] googleSheetsService から AI データを読み込むロジックを修正 ▼▼▼
        const newSlideUrl = await googleSlidesService.generateSlideReport(
            clinicName,
            centralSheetId,
            period,
            periodText
        );
        
        console.log(`[/api/generateSlide] Successfully generated slide. URL: ${newSlideUrl}`);
        
        // 成功したら、クライアントに新しいスライドのURLを返す
        res.json({
            status: 'ok',
            newSlideUrl: newSlideUrl
        });

    } catch (error) {
        console.error(`[/api/generateSlide] Slide generation failed for ${clinicName}:`, error);
        res.status(500).send(`スライド生成失敗: ${error.message}`);
    }
};
