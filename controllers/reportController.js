const googleSheetsService = require('../services/googleSheets');
const pdfGeneratorService = require('../services/pdfGenerator');
const { getSpreadsheetIdFromUrl } = require('../utils/helpers');

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

// --- レポートデータ取得 ---
exports.getReportData = async (req, res) => {
    const { period, selectedClinics } = req.body;
    console.log('POST /api/getReportData called');
    console.log('[/api/getReportData] Request Body:', req.body);

    if (!period || !selectedClinics || !Array.isArray(selectedClinics) || selectedClinics.length === 0) {
        console.error('[/api/getReportData] Invalid request body:', req.body);
        return res.status(400).send('Invalid request: period and selectedClinics array required.');
    }

    try {
        const masterClinicUrls = await googleSheetsService.getMasterClinicUrls();
        if (!masterClinicUrls) {
            console.log('[/api/getReportData] No clinic/URL data found in master sheet.');
            return res.json({});
        }

        const clinicUrls = {};
        Object.entries(masterClinicUrls).forEach(([clinicName, sheetUrl]) => {
            if (selectedClinics.includes(clinicName) && sheetUrl) {
                const sheetId = getSpreadsheetIdFromUrl(sheetUrl);
                if (sheetId) {
                    clinicUrls[clinicName] = sheetId;
                } else {
                    console.warn(`[/api/getReportData] Invalid URL found for ${clinicName}: ${sheetUrl}`);
                }
            }
        });

        console.log('[/api/getReportData] Target Clinic Sheet IDs:', clinicUrls);
        if (Object.keys(clinicUrls).length === 0) {
            console.warn('[/api/getReportData] No valid sheet IDs found for selected clinics.');
            return res.json({});
        }

        const reportData = await googleSheetsService.fetchAndAggregateReportData(clinicUrls, period);

        console.log('[/api/getReportData] Finished processing all selected clinics. Sending report data.');
        res.json(reportData);

    } catch (err) {
        console.error('[/api/getReportData] Error in getReportData controller:', err);
        res.status(500).send(err.message || 'レポートデータ取得中にエラーが発生しました。');
    }
};

// --- PDF生成 ---
exports.generatePdf = async (req, res) => {
    console.log("POST /generate-pdf called");
    const { clinicName, periodText, reportData } = req.body;
    if (!clinicName || !periodText || !reportData) {
        console.error('[/generate-pdf] Missing data:', { clinicName: !!clinicName, periodText: !!periodText, reportData: !!reportData });
        return res.status(400).send('PDF生成に必要なデータが不足');
    }

    try {
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
