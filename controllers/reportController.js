const googleSheetsService = require('../services/googleSheets');
const pdfGeneratorService = require('../services/pdfGenerator');
const aiAnalysisService = require('../services/aiAnalysisService');
// ▼▼▼ [新規] スライド生成サービスを読み込む ▼▼▼
const googleSlidesService = require('../services/googleSlidesService');

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
// === ▼▼▼ [変更] 転記状況の確認 (AI完了ステータスを追加) ▼▼▼ ===
// =================================================================
/**
 * [変更] 集計スプレッドシート内の全タブ名と、AI分析の完了状況を取得する
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
        const sheetTitlesSet = new Set(sheetTitles); // 高速検索用のSet

        // 2. マスターからクリニック一覧を取得
        // (どのクリニックのAI分析状況をチェックすべきか知るため)
        const masterClinics = await googleSheetsService.getMasterClinicList();

        // 3. AI分析の完了ステータスを計算
        const aiCompletionStatus = {};
        // AI分析は5種類
        const requiredTypes = ['L', 'I_good', 'I_bad', 'J', 'M'];

        for (const clinicName of masterClinics) {
            // このクリニックが転記済み (タブが存在する) か？
            if (sheetTitlesSet.has(clinicName)) {
                let isAllComplete = true;
                
                // 5種類すべてのAI分析シート (例: "クリニック名-AI分析-L-分析") が存在するかチェック
                for (const type of requiredTypes) {
                    //
                    const aiSheetName = `${clinicName}-AI分析-${type}-分析`;
                    if (!sheetTitlesSet.has(aiSheetName)) {
                        isAllComplete = false;
                        break; // 1つでも欠けていたらチェック終了
                    }
                }
                aiCompletionStatus[clinicName] = isAllComplete;
            } else {
                aiCompletionStatus[clinicName] = false; // 転記自体されていない
            }
        }
        
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
// === ▼▼▼ [変更] getReportData (AIバックグラウンド実行) (2/3) ▼▼▼ ===
// =================================================================
exports.getReportData = async (req, res) => {
    const { period, selectedClinics, centralSheetId } = req.body;
    console.log('POST /api/getReportData (ETL Trigger) called');
    console.log('[/api/getReportData] Request Body:', req.body);

    if (!period || !selectedClinics || !Array.isArray(selectedClinics) || selectedClinics.length === 0 || !centralSheetId) {
        console.error('[/api/getReportData] Invalid request body:', req.body);
        return res.status(400).send('Invalid request: period, selectedClinics, and centralSheetId required.');
    }
    
    // ▼▼▼ [変更] 10件制限をバックエンドでも強制 (任意) ▼▼▼
    if (selectedClinics.length > 10) {
         console.error('[/api/getReportData] Too many clinics selected:', selectedClinics.length);
        return res.status(400).send('Invalid request: 一度に処理できるのは10件までです。');
    }
    // ▲▲▲

    try {
        // 1. マスターからURLマップを取得 (変更なし)
        const masterClinicUrls = await googleSheetsService.getMasterClinicUrls();
        if (!masterClinicUrls) {
            console.log('[/api/getReportData] No clinic/URL data found in master sheet.');
            return res.json({});
        }
        
        // ▼▼▼ [修正点] clinicUrls のマップには ID ではなく「元URL」を格納する ▼▼▼
        const clinicUrls = {};
        Object.entries(masterClinicUrls).forEach(([clinicName, sheetUrl]) => {
            if (selectedClinics.includes(clinicName) && sheetUrl) {
                clinicUrls[clinicName] = sheetUrl; // URL (https://...) をそのまま渡す
            }
        });
        
        // ▼▼▼ [修正点] ログに表示するのはURL ▼▼▼
        console.log('[/api/getReportData] Target Source Sheet URLs:', clinicUrls);
        if (Object.keys(clinicUrls).length === 0) {
            console.warn('[/api/getReportData] No valid source sheet URLs found for selected clinics.');
            return res.json({ status: 'ok', processed: [] });
        }
        // (IDへの変換は services/googleSheets.js の fetchAndAggregateReportData 側で行われる)

        // 2. データを集計スプシに「転記」する (ETL実行)
        //
        const processedClinics = await googleSheetsService.fetchAndAggregateReportData(clinicUrls, period, centralSheetId);
        
        // 3. ▼▼▼ [新規] AI分析をバックグラウンドで実行 (await しない) ▼▼▼
        if (processedClinics.length > 0) {
            console.log(`[/api/getReportData] Triggering background AI analysis for ${processedClinics.length} clinics...`);
            
            // AI分析の実行をトリガーする。
            // `await` を付けないことで、この関数の完了を待たずに
            // すぐにクライアントに応答を返す (非同期実行)
            runBackgroundAiAnalysis(centralSheetId, processedClinics);
            
        }
        // ▲▲▲
        
        console.log('[/api/getReportData] Finished ETL process. Responding to client.');
        // 4. AI分析の完了を待たずに、すぐにフロントエンドに応答を返す
        res.json({ status: 'ok', processed: processedClinics });

    } catch (err) {
        console.error('[/api/getReportData] Error in getReportData (ETL) controller:', err);
        res.status(500).send(err.message || 'レポートデータ転記中にエラーが発生しました。');
    }
};

/**
 * [新規] AI分析をバックグラウンドで実行するラッパー関数
 * @param {string} centralSheetId 
 * @param {string[]} clinicNames - 処理対象のクリニック名リスト
 */
async function runBackgroundAiAnalysis(centralSheetId, clinicNames) {
    console.log(`[BG-AI] Background task started for ${clinicNames.join(', ')}`);
    
    // AI分析は5種類
    const analysisTypes = ['L', 'I_good', 'I_bad', 'J', 'M'];
    
    // 転記されたクリニックごとにループ
    for (const clinicName of clinicNames) {
        console.log(`[BG-AI] Starting all 5 analyses for ${clinicName}...`);
        try {
            // 5種類のAI分析を順番に実行
            for (const type of analysisTypes) {
                try {
                    console.log(`[BG-AI] Running ${clinicName} - ${type}...`);
                    // 新しいサービス関数を呼び出す
                    await aiAnalysisService.runAndSaveAnalysis(centralSheetId, clinicName, type);
                    console.log(`[BG-AI] SUCCESS: ${clinicName} - ${type}`);
                } catch (e) {
                    if (e.message && e.message.includes('テキストデータが0件')) {
                        console.log(`[BG-AI] SKIP: ${clinicName} - ${type} (No data)`);
                    } else {
                        console.error(`[BG-AI] FAILED: ${clinicName} - ${type}: ${e.message}`);
                    }
                    // 1つのタイプが失敗しても、次のタイプの分析に進む
                }
            }
            console.log(`[BG-AI] COMPLETED all 5 analyses for ${clinicName}`);
        } catch (e) {
            console.error(`[BG-AI] FATAL ERROR for ${clinicName} (e.g., getReportData failed): ${e.message}`);
            // 1つのクリニックが(データ取得などで)失敗しても、次のクリニックに進む
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
        // ★ 要求 #4 (ヘッダーなし転記) に対応した getReportDataForCharts を呼び出す
        //
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
        // ★ 要求 #4 (ヘッダーなし転記) に対応した getReportDataForCharts を呼び出す
        console.log(`[/generate-pdf] Fetching data for PDF from sheet: "${clinicName}"`);
        //
        const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
        
        //
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
// === ▼▼▼ [新規] スライド生成 ▼▼▼ ===
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
        
        // 新しいスライドサービスを呼び出す
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
