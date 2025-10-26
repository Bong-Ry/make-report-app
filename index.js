const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const { google } = require('googleapis'); // Google APIライブラリ

const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;

// --- Render環境設定読み込み ---
const KEYFILEPATH = '/etc/secrets/credentials.json';
const MASTER_SPREADSHEET_ID = process.env.MASTER_SHEET_ID; // マスターシートのID
// ------------------------------

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];

// Google APIクライアントの初期化
let sheets;
try {
  const auth = new google.auth.GoogleAuth({
    keyFile: KEYFILEPATH,
    scopes: SCOPES,
  });
  sheets = google.sheets({ version: 'v4', auth });
  console.log('Google Sheets API client initialized successfully.');
} catch (err) {
  console.error('Failed to initialize Google Sheets API client:', err);
  // エラーが発生してもサーバーは起動させるが、APIは機能しない
}

// --- ヘルパー関数: URLからSpreadsheet IDを抽出 ---
function getSpreadsheetIdFromUrl(url) {
  if (!url || typeof url !== 'string') return null;
  const match = url.match(/\/d\/(.+?)\//);
  return match ? match[1] : null;
}
// --------------------------------------------------

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.use(express.static(path.join(__dirname, 'public')));

app.get('/generate-pdf', async (req, res) => {
  // (PDF生成ロジックはまだダミーのまま)
  try {
    const browser = await puppeteer.launch({
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--executablePath=/usr/bin/google-chrome']
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

app.get('/api/getClinicList', async (req, res) => {
  console.log('GET /api/getClinicList called');
  
  if (!sheets) {
      return res.status(500).send('Google Sheets APIクライアントが初期化されていません。');
  }
  if (!MASTER_SPREADSHEET_ID) {
    console.error('MASTER_SHEET_ID が環境変数に設定されていません。');
    return res.status(500).send('サーバー設定エラー: マスターシートIDがありません。');
  }

  const MASTER_RANGE = 'シート1!A2:A'; // クリニック名リストの範囲

  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: MASTER_SPREADSHEET_ID,
      range: MASTER_RANGE,
    });

    const rows = response.data.values;
    if (rows && rows.length > 0) {
      const clinics = rows.map((row) => row[0]).filter(Boolean);
      console.log('Fetched clinics:', clinics);
      res.json(clinics);
    } else {
      console.log('No clinic data found in master sheet.');
      res.json([]);
    }
  } catch (err) {
    console.error('Master Sheet API returned an error: ' + err);
    res.status(500).send('マスターシートの読み込みに失敗しました。');
  }
});

// --- ▼▼▼ 実際の集計ロジックを実装 ▼▼▼ ---
app.post('/api/getReportData', async (req, res) => {
  const { period, selectedClinics } = req.body;
  console.log('POST /api/getReportData called');
  console.log('Period:', period);
  console.log('Selected Clinics:', selectedClinics);

  if (!sheets) {
      return res.status(500).send('Google Sheets APIクライアントが初期化されていません。');
  }
  if (!MASTER_SPREADSHEET_ID) {
    console.error('MASTER_SHEET_ID が環境変数に設定されていません。');
    return res.status(500).send('サーバー設定エラー: マスターシートIDがありません。');
  }
  if (!period || !selectedClinics || !Array.isArray(selectedClinics)) {
      return res.status(400).send('不正なリクエスト: 期間またはクリニックリストが不足しています。');
  }

  const MASTER_CLINIC_URL_RANGE = 'シート1!A2:B'; // マスターシートのクリニック名とURLの範囲

  try {
    // 1. マスターシートからクリニック名とシートURLを取得
    const masterResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: MASTER_SPREADSHEET_ID,
      range: MASTER_CLINIC_URL_RANGE,
    });

    const masterRows = masterResponse.data.values;
    if (!masterRows || masterRows.length === 0) {
      console.log('No clinic/URL data found in master sheet.');
      return res.json({});
    }

    const clinicUrls = {};
    masterRows.forEach(row => {
      const clinicName = row[0];
      const sheetUrl = row[1];
      // URLからスプレッドシートIDを抽出して保存
      if (selectedClinics.includes(clinicName) && sheetUrl) {
        const sheetId = getSpreadsheetIdFromUrl(sheetUrl);
        if (sheetId) {
            clinicUrls[clinicName] = sheetId;
        } else {
            console.warn(`Invalid URL found for ${clinicName}: ${sheetUrl}`);
        }
      }
    });
    console.log('Target Clinic Sheet IDs:', clinicUrls);

    // 2. 期間設定
    const startDate = new Date(period.start + '-01T00:00:00Z');
    const endDate = new Date(period.end.split('-')[0], period.end.split('-')[1], 0); // 月末日を取得
    endDate.setHours(23, 59, 59, 999); // 月末日の23:59:59.999 に設定
    console.log(`Filtering data between ${startDate.toISOString()} and ${endDate.toISOString()}`);


    const reportData = {};

    // 集計用のキー定義 (GAS版からコピー)
    const satisfactionKeys = ['非常に満足', '満足', 'ふつう', '不満', '非常に不満'];
    const ageKeys = ['10代', '20代', '30代', '40代'];
    const childrenKeys = ['1人', '2人', '3人', '4人', '5人以上'];
    const initializeCounts = (keys) => keys.reduce((acc, key) => { acc[key] = 0; return acc; }, {});
    const createChartData = (counts, keys) => {
        const chartData = [['カテゴリ', '件数']];
        keys.forEach(key => {
          // カウントが0より大きい場合のみ追加（GAS版のロジックに合わせる）
          if (counts[key] > 0) {
            chartData.push([key, counts[key]]);
          }
        });
        // データがない場合でもヘッダー行は返す
        return chartData;
    };


    // 3. クリニックごとにデータを取得して集計
    for (const clinicName in clinicUrls) {
      const clinicSheetId = clinicUrls[clinicName];
      console.log(`Processing data for ${clinicName} (Sheet ID: ${clinicSheetId})`);

      const allNpsReasons = [];
      const allFeedbacks_I = [];
      const allFeedbacks_J = [];
      const allFeedbacks_M = [];
      
      const satisfactionCounts_B = initializeCounts(satisfactionKeys);
      const satisfactionCounts_C = initializeCounts(satisfactionKeys);
      const satisfactionCounts_D = initializeCounts(satisfactionKeys);
      const satisfactionCounts_E = initializeCounts(satisfactionKeys);
      const satisfactionCounts_F = initializeCounts(satisfactionKeys);
      const satisfactionCounts_G = initializeCounts(satisfactionKeys);
      const satisfactionCounts_H = initializeCounts(satisfactionKeys);
      const childrenCounts_P = initializeCounts(childrenKeys);
      const ageCounts_O = initializeCounts(ageKeys);
      const incomeCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0 };

      try {
        // データシートからデータを取得 (A列からQ列まで)
        const clinicDataResponse = await sheets.spreadsheets.values.get({
          spreadsheetId: clinicSheetId,
          range: "'フォームの回答 １'!A:Q", // GAS版と同じ範囲を指定
          dateTimeRenderOption: 'SERIAL_NUMBER', // 日付をシリアル値で取得
          valueRenderOption: 'UNFORMATTED_VALUE' // 書式設定されていない値を取得 (特に数値)
        });

        const clinicDataRows = clinicDataResponse.data.values;
        if (!clinicDataRows || clinicDataRows.length < 2) { // ヘッダー行+データ1行以上必要
            console.log(`No data or only header found in sheet for ${clinicName}.`);
            continue; // 次のクリニックへ
        }

        const header = clinicDataRows.shift(); // ヘッダー行を削除

        // 列インデックス (GAS版に合わせる)
        const timestampIndex = 0;    // A
        const satBIndex = 1;         // B
        const satCIndex = 2;         // C
        const satDIndex = 3;         // D
        const satEIndex = 4;         // E
        const satFIndex = 5;         // F
        const satGIndex = 6;         // G
        const satHIndex = 7;         // H
        const feedbackIIndex = 8;    // I
        const feedbackJIndex = 9;    // J
        const scoreKIndex = 10;      // K
        const reasonLIndex = 11;     // L
        const feedbackMIndex = 12;   // M
        // N列はスキップ
        const ageOIndex = 14;        // O
        const childrenPIndex = 15;   // P
        const incomeQIndex = 16;     // Q

        // 行ごとに集計
        clinicDataRows.forEach(row => {
          // スプレッドシートのシリアル値をJavaScriptのDateに変換
          // シリアル値 1 = 1899年12月31日 (Excel/Sheetsの基準)
          // JavaScriptのDate基準 (1970年1月1日) との差を考慮
          const excelEpoch = new Date(Date.UTC(1899, 11, 30)); // Month is 0-indexed
          const serialValue = row[timestampIndex];
          if (typeof serialValue !== 'number' || serialValue <= 0) return; // 不正な日付はスキップ

          // 日付のみを考慮し、時間を無視する (日単位で比較するため)
          // シリアル値からミリ秒を計算し、JavaScriptのタイムスタンプに変換
          const timestamp = new Date(excelEpoch.getTime() + serialValue * 24 * 60 * 60 * 1000);
          
          // 日付が期間内かチェック (UTC基準で比較)
          if (timestamp.getTime() >= startDate.getTime() && timestamp.getTime() <= endDate.getTime()) {
            
            // NPS (K, L)
            const score = row[scoreKIndex];
            const reason = row[reasonLIndex];
            if (reason != null && String(reason).trim() !== '') {
              allNpsReasons.push({ score: parseInt(score, 10), reason: String(reason).trim() });
            }
            
            // Feedbacks (I, J, M)
            const feedbackI = row[feedbackIIndex]; if (feedbackI != null && String(feedbackI).trim() !== '') allFeedbacks_I.push(String(feedbackI).trim());
            const feedbackJ = row[feedbackJIndex]; if (feedbackJ != null && String(feedbackJ).trim() !== '') allFeedbacks_J.push(String(feedbackJ).trim());
            const feedbackM = row[feedbackMIndex]; if (feedbackM != null && String(feedbackM).trim() !== '') allFeedbacks_M.push(String(feedbackM).trim());

            // Satisfactions (B-H) - GAS版と同様のキーでカウント
            const satB = row[satBIndex]; if (satB != null && satisfactionKeys.includes(String(satB))) satisfactionCounts_B[String(satB)]++;
            const satC = row[satCIndex]; if (satC != null && satisfactionKeys.includes(String(satC))) satisfactionCounts_C[String(satC)]++;
            const satD = row[satDIndex]; if (satD != null && satisfactionKeys.includes(String(satD))) satisfactionCounts_D[String(satD)]++;
            const satE = row[satEIndex]; if (satE != null && satisfactionKeys.includes(String(satE))) satisfactionCounts_E[String(satE)]++;
            const satF = row[satFIndex]; if (satF != null && satisfactionKeys.includes(String(satF))) satisfactionCounts_F[String(satF)]++;
            const satG = row[satGIndex]; if (satG != null && satisfactionKeys.includes(String(satG))) satisfactionCounts_G[String(satG)]++;
            const satH = row[satHIndex]; if (satH != null && satisfactionKeys.includes(String(satH))) satisfactionCounts_H[String(satH)]++;

            // Children (P)
            const childrenP = row[childrenPIndex]; if (childrenP != null && childrenKeys.includes(String(childrenP))) childrenCounts_P[String(childrenP)]++;

            // Age (O)
            const ageO = row[ageOIndex]; if (ageO != null && ageKeys.includes(String(ageO))) ageCounts_O[String(ageO)]++;
            
            // Income (Q) - 数値として扱う
            const income = row[incomeQIndex];
            if (typeof income === 'number' && income >= 1 && income <= 10) {
              incomeCounts[income]++;
            }
          }
        });

      } catch (e) {
        console.error(`Error processing sheet for ${clinicName} (ID: ${clinicSheetId}): ${e.toString()}`);
        // エラーが発生しても他のクリニックの処理は続ける
        continue;
      }
      
      // 集計結果を整形 (GAS版と同じロジック)
      const groupedByScore = allNpsReasons.reduce((acc, item) => {
        if (!acc[item.score]) acc[item.score] = [];
        acc[item.score].push(item.reason);
        return acc;
      }, {});
      
      const incomeChartData = [['評価', '割合', { role: 'annotation' }]];
      const totalIncomeCount = Object.values(incomeCounts).reduce((a, b) => a + b, 0);
      if (totalIncomeCount > 0) {
        for (let i = 1; i <= 10; i++) {
          const percentage = (incomeCounts[i] / totalIncomeCount) * 100;
          incomeChartData.push([String(i), percentage, `${Math.round(percentage)}%`]);
        }
      }

      // レポートデータを構築
      reportData[clinicName] = {
        npsData: { totalCount: allNpsReasons.length, results: groupedByScore },
        feedbackData: {
          i_column: { totalCount: allFeedbacks_I.length, results: allFeedbacks_I },
          j_column: { totalCount: allFeedbacks_J.length, results: allFeedbacks_J },
          m_column: { totalCount: allFeedbacks_M.length, results: allFeedbacks_M }
        },
        satisfactionData: {
          b_column: { results: createChartData(satisfactionCounts_B, satisfactionKeys) },
          c_column: { results: createChartData(satisfactionCounts_C, satisfactionKeys) },
          d_column: { results: createChartData(satisfactionCounts_D, satisfactionKeys) },
          e_column: { results: createChartData(satisfactionCounts_E, satisfactionKeys) },
          f_column: { results: createChartData(satisfactionCounts_F, satisfactionKeys) },
          g_column: { results: createChartData(satisfactionCounts_G, satisfactionKeys) },
          h_column: { results: createChartData(satisfactionCounts_H, satisfactionKeys) }
        },
        ageData: { results: createChartData(ageCounts_O, ageKeys) },
        childrenCountData: { results: createChartData(childrenCounts_P, childrenKeys) },
        incomeData: { results: incomeChartData, totalCount: totalIncomeCount }
      };
      console.log(`Finished processing data for ${clinicName}`);
    } // End of clinic loop
    
    console.log('Finished all clinics. Sending report data.');
    res.json(reportData);

  } catch (err) {
    console.error('Error in /api/getReportData:', err);
    res.status(500).send('レポートデータの取得中にエラーが発生しました。');
  }
});
// --- ▲▲▲ 修正点 ▲▲▲ ---

app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
