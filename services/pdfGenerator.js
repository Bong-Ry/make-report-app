const puppeteer = require('puppeteer');

exports.generatePdfFromData = async (clinicName, periodText, reportData) => {
    console.log("[pdfGeneratorService] Starting PDF generation...");

    // HTML生成 (簡略版)
    let pdfHtml = `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><title>レポート:${clinicName}</title><style>body{font-family:'Noto Sans JP',sans-serif;padding:20px;} h1,h2{border-bottom:1px solid #ccc;padding-bottom:5px;} .comment-list{margin-top:10px;padding-left:20px;white-space:pre-wrap;font-size:10pt;}</style></head><body><h1>レポート:${clinicName}</h1><p>集計期間:${periodText}</p><hr>`;
    pdfHtml += `<h2>NPS理由(全${reportData.npsData?.totalCount || 0}件)</h2>`;
    if (reportData.npsData?.results) {
        const scores = Object.keys(reportData.npsData.results).map(Number).sort((a, b) => b - a);
        scores.forEach(score => {
            const reasons = reportData.npsData.results[score];
            if (reasons && reasons.length > 0) {
                pdfHtml += `<h3>推奨度 ${score}(${reasons.length}人)</h3><ul class="comment-list">`;
                reasons.forEach(reason => {
                    const escapedReason = reason.replace(/</g, "&lt;").replace(/>/g, "&gt;");
                    pdfHtml += `<li>${escapedReason}</li>`;
                });
                pdfHtml += `</ul>`;
            }
        });
    } else {
        pdfHtml += `<p>データなし</p>`;
    }
    pdfHtml += `<hr></body></html>`;

    let browser;
    try {
        console.log("[pdfGeneratorService] Launching Puppeteer...");
        // Render環境での推奨設定を含める
        browser = await puppeteer.launch({
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage', // /dev/shm の使用を減らす
                '--single-process', // シングルプロセスモード（メモリ節約）
                '--no-zygote' // Zygoteプロセス不使用（メモリ節約）
            ],
             // 実行可能パスを指定 (RenderのChromeビルドパックを使う場合)
             // 必要に応じて /usr/bin/google-chrome などに変更
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || puppeteer.executablePath()
        });
        const page = await browser.newPage();
        console.log("[pdfGeneratorService] Setting HTML content...");
        await page.setContent(pdfHtml, { waitUntil: 'networkidle0' });
        console.log("[pdfGeneratorService] Generating PDF buffer...");
        const pdfBuffer = await page.pdf({
            format: 'A4',
            printBackground: true,
            margin: { top: '20mm', right: '10mm', bottom: '20mm', left: '10mm' }
        });
        console.log("[pdfGeneratorService] PDF buffer generated successfully.");
        return pdfBuffer;
    } catch (error) {
        console.error('[pdfGeneratorService] Error during PDF generation:', error);
        throw error; // エラーを再スローしてコントローラーで処理
    } finally {
        if (browser) {
            console.log("[pdfGeneratorService] Closing Puppeteer browser...");
            await browser.close();
            console.log("[pdfGeneratorService] Puppeteer browser closed.");
        }
    }
};
