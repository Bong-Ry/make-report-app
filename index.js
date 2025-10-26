const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const { google } = require('googleapis'); // Google APIライブラリ

const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;

// --- ↓↓ Renderの設定に合わせて修正 ↓↓ ---

// 1. シークレットファイル (credentials.json) のパス
// Renderは /etc/secrets/ にファイルを配置するため、パスをそこに変更
const KEYFILEPATH = '/etc/secrets/credentials.json'; 

// 2. スプレッドシートID
// Renderの環境変数 MASTER_SHEET_ID から読み込む
const SPREADSHEET_ID = process.env.MASTER_SHEET_ID;

// ------------------------------------

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];

const auth = new google.auth.GoogleAuth({
  keyFile: KEYFILEPATH,
  scopes: SCOPES,
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.use(express.static(path.join(__dirname, 'public')));

app.get('/generate-pdf', async (req, res) => {
  try {
    const browser = await puppeteer.launch({
      // Render環境で必須のオプション
      args: [
        '--no-sandbox', 
        '--disable-setuid-sandbox', 
        '--disable-dev-shm-usage',
        // ★ Chrome実行ファイルのパスを明示的に指定 (Renderの標準パス)
        '--executablePath=/usr/bin/google-chrome'
      ]
    });
    const page = await browser.newPage();
    await page.goto('https://www.google.com'); 
    const pdf = await page.pdf({ format: 'A4' });
    await browser.close();
    res.contentType('application/pdf');
    res.send(pdf);
  
  // --- ▼▼▼ 修正点 1/2 ▼▼▼ ---
  // 閉じ括弧 ')' を追加
  } catch (error) { 
  // --- ▲▲▲ 修正点 1/2 ▲▲▲ ---
    console.error('PDF generation failed:', error);
    res.status(500).send('PDFの生成に失敗しました。');
  }
});

app.get('/api/getClinicList', async (req, res) => {
  console.log('GET /api/getClinicList called (from Google Sheets)');
  
  // --- ★ 要確認 (1/1) ★ ---
  // クリニック名が記載されているシート名とセル範囲
  // 現在: 'シート1' の 'A2'セルからA列の最後まで
  const RANGE = 'シート1!A2:A'; 
  // --------------------------

  // SPREADSHEET_ID が環境変数から読み込めているかチェック
  if (!SPREADSHEET_ID) {
    console.error('MASTER_SHEET_ID (スプレッドシートID) が環境変数に設定されていません。');
    return res.status(500).send('サーバー設定エラー: スプレッドシートIDがありません。');
  }

  try {
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: RANGE,
    });

    const rows = response.data.values;
    if (rows && rows.length > 0) {
      const clinics = rows.map((row) => row[0]).filter(Boolean);
      console.log('Fetched clinics:', clinics);
      res.json(clinics);
    } else {
      console.log('No data found.');
      res.json([]);
    }
  } catch (err) {
    console.error('Google Sheets API returned an error: ' + err);
    // 認証エラーや共有設定ミスの場合は、ここでエラーが発生します
    res.status(500).send('スプレッドシートの読み込みに失敗しました。Renderのシークレットファイルやスプレッドシートの共有設定を確認してください。');
  }
});

// --- ▼▼▼ 修正点 2/2 ▼▼▼ ---
// /api/getReportData を呼び出すPOSTエンドポイント（ダミー）を追加
// (index.html からの呼び出しに対応するため)
app.post('/api/getReportData', async (req, res) => {
  const { period, selectedClinics } = req.body;
  console.log('POST /api/getReportData called');
  console.log('Period:', period);
  console.log('Clinics:', selectedClinics);

  // TODO: 本来はここで period と selectedClinics を使って
  // スプレッドシートから集計データを取得するロジックを実装します。

  // (現在はダミーデータを返します)
  // (index.html が期待する形式: { "クリニック名": { ...データ... } })
  const dummyReportData = {};
  selectedClinics.forEach(name => {
    dummyReportData[name] = {
      npsData: { totalCount: 2, results: { 10: ['素晴らしい (ダミー)'], 1: ['良くない (ダミー)'] } },
      feedbackData: { i_column: { totalCount: 1, results: ['良かった (ダミー)'] }, j_column: { totalCount: 0, results: [] }, m_column: { totalCount: 0, results: [] } },
      satisfactionData: { b_column: { results: [['評価', '件数'], ['5', 10], ['4', 5]] } },
      ageData: { results: [['年代', '件数'], ['30代', 15]] },
      childrenCountData: { results: [['人数', '件数'], ['1人', 15]] },
      incomeData: { totalCount: 1, results: [['年収', '割合', { role: 'annotation' }], ['1000万', 100, '100%']] }
    };
  });

  res.json(dummyReportData);
});
// --- ▲▲▲ 修正点 2/2 ▲▲▲ ---

app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
