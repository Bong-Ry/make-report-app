const googleSheetsService = require('../services/googleSheets');
const pdfGeneratorService = require('../services/pdfGenerator');
// ▼▼▼ getSpreadsheetIdFromUrl は不要になったため削除 ▼▼▼
// const { getSpreadsheetIdFromUrl } = require('../utils/helpers');

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
// === ▼▼▼ 新規API (1/3) ▼▼▼ ===
// 集計用スプレッドシートの検索または新規作成
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
// === ▼▼▼ 更新API (2/3) ▼▼▼ ===
// 役割変更: 「レポート発行」ボタンで呼ばれ、データ転記(ETL)を実行する
// =================================================================
exports.getReportData = async (req, res) => {
    // ▼▼▼ centralSheetId をリクエストから受け取る ▼▼▼
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
        const clinicUrls = {};
        Object.entries(masterClinicUrls).forEach(([clinicName, sheetUrl]) => {
            if (selectedClinics.includes(clinicName) && sheetUrl) {
                // getSpreadsheetIdFromUrl は googleSheetsService 側で実行される
                // ここでは元URLのマップではなく、IDのマップを渡す必要がある
                // ※※ services/googleSheets.js の getMasterClinicUrls がIDを返すと仮定 ※※
                // → services/googleSheets.js の getMasterClinicUrls はURLを返すままだった
                // → → services/googleSheets.js (前回の) を修正する必要がある
                
                // --- services/googleSheets.js の getMasterClinicUrls を修正する前提で進めます ---
                // ※※※
                // 変更案: 
                // services/googleSheets.js の getMasterClinicUrls 内で getSpreadsheetIdFromUrl を呼ぶようにする
                // (前回のファイル出力を修正)
                // 
                // [前回]
                // if (row[0] && row[1]) { urlMap[row[0]] = row[1]; }
                // [修正案]
                // if (row[0] && row[1]) { 
                //   const sheetId = getSpreadsheetIdFromUrl(row[1]);
                //   if (sheetId) { urlMap[row[0]] = sheetId; }
                // }
                // ※※※
                
                // (上記の修正が services/googleSheets.js に適用されたと仮定して)
                // 
                // [修正後のロジック]
                const sourceSheetId = googleSheetsService.getSpreadsheetIdFromUrl(sheetUrl);
                if (sourceSheetId) {
                    clinicUrls[clinicName] = sourceSheetId;
                } else {
                     console.warn(`[/api/getReportData] Invalid URL found for ${clinicName}: ${sheetUrl}`);
                }
            }
        });
        
        console.log('[/api/getReportData] Target Source Sheet IDs:', clinicUrls);
        if (Object.keys(clinicUrls).length === 0) {
            console.warn('[/api/getReportData] No valid source sheet IDs found for selected clinics.');
            return res.json({ status: 'ok', processed: [] });
        }

        // 2. ▼▼▼ データを集計スプシに「転記」する (ETL実行) ▼▼▼
        const processedClinics = await googleSheetsService.fetchAndAggregateReportData(clinicUrls, period, centralSheetId);
        
        console.log('[/api/getReportData] Finished ETL process. Processed clinics:', processedClinics);
        // ▼▼▼ 成功ステータスと処理済みクリニック名を返す ▼▼▼
        res.json({ status: 'ok', processed: processedClinics });

    } catch (err) {
        console.error('[/api/getReportData] Error in getReportData (ETL) controller:', err);
        res.status(500).send(err.message || 'レポートデータ転記中にエラーが発生しました。');
    }
};

// =================================================================
// === ▼▼▼ 新規API (3/3) ▼▼▼ ===
// 集計スプシから「集計済みのグラフ用データ」を取得する
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
// === ▼▼▼ 更新API ▼▼▼ ===
// PDF生成ロジックを、集計スプシからのデータ取得に変更
// =================================================================
exports.generatePdf = async (req, res) => {
    console.log("POST /generate-pdf called");
    // ▼▼▼ reportData の代わりに centralSheetId と clinicName を受け取る ▼▼▼
    const { clinicName, periodText, centralSheetId } = req.body;
    
    if (!clinicName || !periodText || !centralSheetId) {
        console.error('[/generate-pdf] Missing data:', { clinicName: !!clinicName, periodText: !!periodText, centralSheetId: !!centralSheetId });
        return res.status(400).send('PDF生成に必要なデータ(clinicName, periodText, centralSheetId)が不足');
    }

    try {
        // ▼▼▼ PDF生成に必要なデータを集計スプシから取得 ▼▼▼
        console.log(`[/generate-pdf] Fetching data for PDF from sheet: "${clinicName}"`);
        const reportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
        
        // (pdfGeneratorService は npsData しか使っていないため、これだけで動くはず)
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
