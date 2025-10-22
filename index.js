const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path'); // Node.jsの標準モジュール
const app = express();

// JSONリクエストボディを解析するためのミドルウェア
app.use(express.json());
// 'public' フォルダ内の静的ファイル（index.htmlなど）を配信する設定
app.use(express.static(path.join(__dirname, 'public')));

// Renderが使用するPORT環境変数、またはローカル用の3000番ポート
const PORT = process.env.PORT || 3000;

// ルートURL ('/') にアクセスがあったら、public/index.html を送信
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

/**
 * PDF生成エンドポイント (ダミー)
 * 今後、GASアプリのロジックをここに移植します。
 */
app.get('/generate-pdf', async (req, res) => {
  try {
    const browser = await puppeteer.launch({
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    const page = await browser.newPage();
    
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
