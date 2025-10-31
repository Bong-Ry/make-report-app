const googleSheetsService = require('../services/googleSheets');
const pdfGeneratorService = require('../services/pdfGenerator');

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
// === ▼▼▼ 新規API (要求 #6 転記状況の確認) ▼▼▼ ===
// =================================================================
/**
 * [新規] 集計スプレッドシート内の全タブ名を取得する
 * (フロントエンドが「転記済みか」を判断するために使う)
 */
exports.getTransferredList = async (req, res) => {
    const { centralSheetId } = req.body;
    console.log(`POST /api/getTransferredList called for: ${centralSheetId}`);

    if (!centralSheetId) {
        return res.status(400).send('Invalid request: centralSheetId required.');
    }
    
    try {
        const sheetTitles = await googleSheetsService.getSheetTitles(centralSheetId);
        // sheetTitles = ["全体", "フォームの回答 1", "あらかわレディースクリニック", "サンプルクリニック", ...]
        res.json({ sheetTitles: sheetTitles });
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
// === ▼▼▼ 更新API (2/3) (バグ修正) ▼▼▼ ===
// (前回修正した getSpreadsheetIdFromUrl のロジックが間違っていたため修正)
// =================================================================
exports.getReportData = async (req, res) => {
    const { period, selectedClinics, centralSheetId } = req.body;
    console.log('POST /api/getReportData (ETL Trigger) called');
    console.log('[/api/getReportData] Request Body:', req.body);

    if (!period || !selectedClinics || !Array.isArray(selectedClinics) || selectedClinics.length === 0 || !centralSheetId) {
        console.error('[/api/getReportData] Invalid request body:', req.body);
        return res.status(400).send('Invalid request: period, selectedClinics, and centralSheetId required.');
    }

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
        const processedClinics = await googleSheetsService.fetchAndAggregateReportData(clinicUrls, period, centralSheetId);
        
        console.log('[/api/getReportData] Finished ETL process. Processed clinics:', processedClinics);
        res.json({ status: 'ok', processed: processedClinics });

    } catch (err) {
        console.error('[/api/getReportData] Error in getReportData (ETL) controller:', err);
        res.status(500).send(err.message || 'レポートデータ転記中にエラーが発生しました。');
    }
};

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
