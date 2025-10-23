const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path'); // Node.jsの標準モジュール
const app = express();

// JSONリクエストボディを解析するためのミドルウェア
app.use(express.json());

// Renderが使用するPORT環境変数、またはローカル用の3000番ポート
const PORT = process.env.PORT || 3000;

// ★★★ 修正点 ★★★
// ルートURL ('/') へのアクセスを先に明示的に処理する
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ★★★ 修正点 ★★★
// / 以外の静的ファイル（CSSやJSなど）を 'public' フォルダから配信する
app.use(express.static(path.join(__dirname, 'public')));


/**
 * PDF生成エンドポイント (ダミー)
 * 今後、GASアプリのロジックをここに移植します。
 */
app.get('/generate-pdf', async (req, res) => {
  try {
    const browser = await puppeteer.launch({
      // Render環境で必要となるChrome起動オプション
      args: [
        '--no-sandbox', 
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage' // メモリ不足対策
      ]
    });
    
    const page = await browser.newPage();
    
    // ダミーとしてGoogleのページをPDF化
    await page.goto('https://www.google.com'); 
    const pdf = await page.pdf({ format: 'A4' });

    await browser.close();

    res.contentType('application/pdf');
    res.send(pdf);

  } catch (error) {
    console.error('PDF generation failed:', error);
    res.status(500).send('PDFの生成に失敗しました。');
  }
});

// (ここに今後、/api/getClinicList や /api/getReportData などのAPIを実装していきます)


app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
