const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const { google } = require('googleapis');
const OpenAI = require('openai'); // ★ OpenAIライブラリを追加

const app = express();
app.use(express.json({ limit: '10mb' })); // ★ JSONリクエストのサイズ上限を上げる (大量テキスト対策)
const PORT = process.env.PORT || 3000;

// --- Render環境設定読み込み ---
const KEYFILEPATH = '/etc/secrets/credentials.json';
const MASTER_SPREADSHEET_ID = process.env.MASTER_SHEET_ID;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY; // ★ OpenAI APIキーを読み込む
// ------------------------------

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];

// --- Google APIクライアント初期化 ---
let sheets;
try {
  const auth = new google.auth.GoogleAuth({ keyFile: KEYFILEPATH, scopes: SCOPES });
  sheets = google.sheets({ version: 'v4', auth });
  console.log('Google Sheets API client initialized successfully.');
} catch (err) {
  console.error('Failed to initialize Google Sheets API client:', err);
}
// ------------------------------------

// --- ★ OpenAIクライアント初期化 ---
let openai;
if (OPENAI_API_KEY) {
  try {
    openai = new OpenAI({ apiKey: OPENAI_API_KEY });
    console.log('OpenAI client initialized successfully.');
  } catch (err) {
    console.error('Failed to initialize OpenAI client:', err);
  }
} else {
  console.warn('OPENAI_API_KEY environment variable is not set. Text analysis API will not work.');
}
// ------------------------------------

// --- ヘルパー関数: URLからSpreadsheet IDを抽出 ---
function getSpreadsheetIdFromUrl(url) { /* (変更なし) */
    if (!url || typeof url !== 'string') return null;
    const match = url.match(/\/d\/(.+?)\//);
    return match ? match[1] : null;
}
// --------------------------------------------------

app.get('/', (req, res) => { /* (変更なし) */
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});
app.use(express.static(path.join(__dirname, 'public')));
app.get('/generate-pdf', async (req, res) => { /* (変更なし - ダミーのまま) */
    try {
        const browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--executablePath=/usr/bin/google-chrome'] });
        const page = await browser.newPage();
        await page.goto('https://www.google.com');
        const pdf = await page.pdf({ format: 'A4' });
        await browser.close();
        res.contentType('application/pdf');
        res.send(pdf);
    } catch (error) { console.error('PDF generation failed:', error); res.status(500).send('PDFの生成に失敗しました。'); }
});
app.get('/api/getClinicList', async (req, res) => { /* (変更なし) */
    console.log('GET /api/getClinicList called');
    if (!sheets) return res.status(500).send('Google Sheets APIクライアントが初期化されていません。');
    if (!MASTER_SPREADSHEET_ID) { console.error('MASTER_SHEET_ID 環境変数が未設定'); return res.status(500).send('サーバー設定エラー: マスターシートIDがありません。'); }
    const MASTER_RANGE = 'シート1!A2:A';
    try {
        const response = await sheets.spreadsheets.values.get({ spreadsheetId: MASTER_SPREADSHEET_ID, range: MASTER_RANGE });
        const rows = response.data.values;
        const clinics = (rows && rows.length > 0) ? rows.map((row) => row[0]).filter(Boolean) : [];
        console.log('Fetched clinics:', clinics);
        res.json(clinics);
    } catch (err) { console.error('Master Sheet API error:', err); res.status(500).send('マスターシートの読み込みに失敗しました。'); }
});
app.post('/api/getReportData', async (req, res) => { /* (変更なし - 実際の集計ロジック) */
    const { period, selectedClinics } = req.body;
    console.log('POST /api/getReportData called'); console.log('Period:', period); console.log('Clinics:', selectedClinics);
    if (!sheets) return res.status(500).send('Google Sheets APIクライアントが初期化されていません。');
    if (!MASTER_SPREADSHEET_ID) { console.error('MASTER_SHEET_ID 環境変数が未設定'); return res.status(500).send('サーバー設定エラー: マスターシートIDがありません。'); }
    if (!period || !selectedClinics || !Array.isArray(selectedClinics)) return res.status(400).send('不正なリクエスト');
    const MASTER_CLINIC_URL_RANGE = 'シート1!A2:B';
    try {
        const masterResponse = await sheets.spreadsheets.values.get({ spreadsheetId: MASTER_SPREADSHEET_ID, range: MASTER_CLINIC_URL_RANGE });
        const masterRows = masterResponse.data.values;
        if (!masterRows || masterRows.length === 0) { console.log('No clinic/URL data in master sheet.'); return res.json({}); }
        const clinicUrls = {}; masterRows.forEach(row => { const clinicName = row[0], sheetUrl = row[1]; if (selectedClinics.includes(clinicName) && sheetUrl) { const sheetId = getSpreadsheetIdFromUrl(sheetUrl); if (sheetId) clinicUrls[clinicName] = sheetId; else console.warn(`Invalid URL for ${clinicName}`); } });
        console.log('Target Clinic Sheet IDs:', clinicUrls);
        const startDate = new Date(period.start + '-01T00:00:00Z'); const endDate = new Date(period.end.split('-')[0], period.end.split('-')[1], 0); endDate.setHours(23, 59, 59, 999);
        console.log(`Filtering data between ${startDate.toISOString()} and ${endDate.toISOString()}`);
        const reportData = {};
        const satisfactionKeys = ['非常に満足', '満足', 'ふつう', '不満', '非常に不満']; const ageKeys = ['10代', '20代', '30代', '40代']; const childrenKeys = ['1人', '2人', '3人', '4人', '5人以上'];
        const initializeCounts = (keys) => keys.reduce((acc, key) => { acc[key] = 0; return acc; }, {});
        const createChartData = (counts, keys) => { const chartData = [['カテゴリ', '件数']]; keys.forEach(key => { if (counts[key] > 0) chartData.push([key, counts[key]]); }); return chartData; };
        for (const clinicName in clinicUrls) {
            const clinicSheetId = clinicUrls[clinicName]; console.log(`Processing ${clinicName} (ID: ${clinicSheetId})`);
            const allNpsReasons = [], allFeedbacks_I = [], allFeedbacks_J = [], allFeedbacks_M = [];
            const satisfactionCounts_B = initializeCounts(satisfactionKeys), satisfactionCounts_C = initializeCounts(satisfactionKeys), satisfactionCounts_D = initializeCounts(satisfactionKeys), satisfactionCounts_E = initializeCounts(satisfactionKeys), satisfactionCounts_F = initializeCounts(satisfactionKeys), satisfactionCounts_G = initializeCounts(satisfactionKeys), satisfactionCounts_H = initializeCounts(satisfactionKeys);
            const childrenCounts_P = initializeCounts(childrenKeys); const ageCounts_O = initializeCounts(ageKeys); const incomeCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0, 7: 0, 8: 0, 9: 0, 10: 0 };
            try {
                const clinicDataResponse = await sheets.spreadsheets.values.get({ spreadsheetId: clinicSheetId, range: "'フォームの回答 １'!A:Q", dateTimeRenderOption: 'SERIAL_NUMBER', valueRenderOption: 'UNFORMATTED_VALUE' });
                const clinicDataRows = clinicDataResponse.data.values; if (!clinicDataRows || clinicDataRows.length < 2) { console.log(`No data for ${clinicName}.`); continue; }
                const header = clinicDataRows.shift();
                const timestampIndex = 0, satBIndex = 1, satCIndex = 2, satDIndex = 3, satEIndex = 4, satFIndex = 5, satGIndex = 6, satHIndex = 7, feedbackIIndex = 8, feedbackJIndex = 9, scoreKIndex = 10, reasonLIndex = 11, feedbackMIndex = 12, ageOIndex = 14, childrenPIndex = 15, incomeQIndex = 16;
                clinicDataRows.forEach(row => {
                    const excelEpoch = new Date(Date.UTC(1899, 11, 30)); const serialValue = row[timestampIndex]; if (typeof serialValue !== 'number' || serialValue <= 0) return;
                    const timestamp = new Date(excelEpoch.getTime() + serialValue * 24 * 60 * 60 * 1000);
                    if (timestamp.getTime() >= startDate.getTime() && timestamp.getTime() <= endDate.getTime()) {
                        const score = row[scoreKIndex], reason = row[reasonLIndex]; if (reason != null && String(reason).trim() !== '') allNpsReasons.push({ score: parseInt(score, 10), reason: String(reason).trim() });
                        const feedbackI = row[feedbackIIndex]; if (feedbackI != null && String(feedbackI).trim() !== '') allFeedbacks_I.push(String(feedbackI).trim());
                        const feedbackJ = row[feedbackJIndex]; if (feedbackJ != null && String(feedbackJ).trim() !== '') allFeedbacks_J.push(String(feedbackJ).trim());
                        const feedbackM = row[feedbackMIndex]; if (feedbackM != null && String(feedbackM).trim() !== '') allFeedbacks_M.push(String(feedbackM).trim());
                        const satB = row[satBIndex]; if (satB != null && satisfactionKeys.includes(String(satB))) satisfactionCounts_B[String(satB)]++;
                        const satC = row[satCIndex]; if (satC != null && satisfactionKeys.includes(String(satC))) satisfactionCounts_C[String(satC)]++;
                        const satD = row[satDIndex]; if (satD != null && satisfactionKeys.includes(String(satD))) satisfactionCounts_D[String(satD)]++;
                        const satE = row[satEIndex]; if (satE != null && satisfactionKeys.includes(String(satE))) satisfactionCounts_E[String(satE)]++;
                        const satF = row[satFIndex]; if (satF != null && satisfactionKeys.includes(String(satF))) satisfactionCounts_F[String(satF)]++;
                        const satG = row[satGIndex]; if (satG != null && satisfactionKeys.includes(String(satG))) satisfactionCounts_G[String(satG)]++;
                        const satH = row[satHIndex]; if (satH != null && satisfactionKeys.includes(String(satH))) satisfactionCounts_H[String(satH)]++;
                        const childrenP = row[childrenPIndex]; if (childrenP != null && childrenKeys.includes(String(childrenP))) childrenCounts_P[String(childrenP)]++;
                        const ageO = row[ageOIndex]; if (ageO != null && ageKeys.includes(String(ageO))) ageCounts_O[String(ageO)]++;
                        const income = row[incomeQIndex]; if (typeof income === 'number' && income >= 1 && income <= 10) incomeCounts[income]++;
                    }
                });
            } catch (e) { console.error(`Error processing sheet for ${clinicName}: ${e.toString()}`); continue; }
            const groupedByScore = allNpsReasons.reduce((acc, item) => { if (!acc[item.score]) acc[item.score] = []; acc[item.score].push(item.reason); return acc; }, {});
            const incomeChartData = [['評価', '割合', { role: 'annotation' }]]; const totalIncomeCount = Object.values(incomeCounts).reduce((a, b) => a + b, 0);
            if (totalIncomeCount > 0) { for (let i = 1; i <= 10; i++) { const percentage = (incomeCounts[i] / totalIncomeCount) * 100; incomeChartData.push([String(i), percentage, `${Math.round(percentage)}%`]); } }
            reportData[clinicName] = {
                npsData: { totalCount: allNpsReasons.length, results: groupedByScore, rawText: allNpsReasons.map(r => r.reason) }, // ★ 元テキストも追加
                feedbackData: { i_column: { totalCount: allFeedbacks_I.length, results: allFeedbacks_I }, j_column: { totalCount: allFeedbacks_J.length, results: allFeedbacks_J }, m_column: { totalCount: allFeedbacks_M.length, results: allFeedbacks_M } },
                satisfactionData: { b_column: { results: createChartData(satisfactionCounts_B, satisfactionKeys) }, c_column: { results: createChartData(satisfactionCounts_C, satisfactionKeys) }, d_column: { results: createChartData(satisfactionCounts_D, satisfactionKeys) }, e_column: { results: createChartData(satisfactionCounts_E, satisfactionKeys) }, f_column: { results: createChartData(satisfactionCounts_F, satisfactionKeys) }, g_column: { results: createChartData(satisfactionCounts_G, satisfactionKeys) }, h_column: { results: createChartData(satisfactionCounts_H, satisfactionKeys) } },
                ageData: { results: createChartData(ageCounts_O, ageKeys) }, childrenCountData: { results: createChartData(childrenCounts_P, childrenKeys) }, incomeData: { results: incomeChartData, totalCount: totalIncomeCount }
            };
            console.log(`Finished processing data for ${clinicName}`);
        }
        console.log('Finished all clinics. Sending report data.');
        res.json(reportData);
    } catch (err) { console.error('Error in /api/getReportData:', err); res.status(500).send('レポートデータの取得中にエラーが発生しました。'); }
});


// --- ★★★ 新しいAPIエンドポイント: テキスト解析 ★★★ ---
app.post('/api/analyzeText', async (req, res) => {
    console.log('POST /api/analyzeText called');
    const { textList } = req.body; // フロントからテキストの配列を受け取る

    if (!openai) {
        return res.status(500).send('OpenAIクライアントが初期化されていません。APIキーを確認してください。');
    }
    if (!textList || !Array.isArray(textList) || textList.length === 0) {
        return res.status(400).send('解析対象のテキストリストが必要です。');
    }

    // 簡単な単語分割とフィルタリング（日本語の基本的な助詞、助動詞、記号などを除外）
    // より正確な処理にはMeCabやSudachiPyなどが必要だが、まずは簡易的に実装
    const unwantedChars = /[、「」。、？！・（）()\[\]{}【】『』<>]+/g;
    const stopWords = new Set([
        'の', 'に', 'は', 'を', 'が', 'と', 'へ', 'や', 'も', 'で', 'から', 'まで', 'より', 'です', 'ます', 'でした', 'ました', 'する', 'し', 'さ', 'れ', 'いる', 'い', 'あり', 'ある', 'ない', 'なく', 'なる', 'なっ', '思う', '思い', '感じ', '感じた', 'いう', '言っ', 'こと', 'もの', 'よう', 'ため', 'とき', '中', '等', '的', '的', 'まし', 'ので', 'から', 'けど', 'また', 'そして', 'しかし', 'とても', 'すごく', '特に', '非常', '私', '方', 'これ', 'それ', 'あれ', 'ここ', 'そこ', 'あそこ', 'どの', 'この', 'その', 'あの'
        // 必要に応じてストップワードを追加
    ]);

    // OpenAIに投げるテキストは結合する（APIコール回数を減らすため）
    const combinedText = textList.join('\n');
    console.log(`Analyzing ${textList.length} texts (combined length: ${combinedText.length})`);

    try {
        const completion = await openai.chat.completions.create({
            model: "gpt-4o-mini",
            messages: [
                {
                    role: "system",
                    content: `以下の日本語テキストから、重要な「名詞」「動詞」「形容詞」「感動詞」を抽出し、それぞれの単語と品詞のペアをJSON形式でリストアップしてください。一般的な単語（例：「こと」「もの」「する」「思う」「です」「ます」）や助詞、助動詞、記号は除外してください。同じ単語が複数回出現しても、リストには一度だけ含めてください。
例: [{"word": "病院", "pos": "名詞"}, {"word": "綺麗", "pos": "形容詞"}, {"word": "感謝", "pos": "名詞"}, {"word": "ありがとう", "pos": "感動詞"}, {"word": "診る", "pos": "動詞"}]`
                },
                {
                    role: "user",
                    content: combinedText.substring(0, 15000) // ★ トークン数制限のため、長すぎる場合は切り詰める
                }
            ],
            response_format: { type: "json_object" }, // ★ JSONモードを指示
        });

        console.log("OpenAI API response received.");
        let wordsWithPos = [];
        try {
            // contentがJSON文字列として返ってくることを想定
            const jsonResponse = JSON.parse(completion.choices[0].message.content);
            // 想定する形式は {"words": [{"word": "単語", "pos": "品詞"}, ...]} のようなものかもしれないし、
            // プロンプト例のように直接配列かもしれないので、両方に対応してみる
            if (Array.isArray(jsonResponse)) {
                wordsWithPos = jsonResponse;
            } else if (jsonResponse && Array.isArray(jsonResponse.words)) {
                 wordsWithPos = jsonResponse.words;
            } else {
                 console.error("Unexpected JSON structure from OpenAI:", jsonResponse);
                 throw new Error("OpenAIからの応答が予期しない形式です。");
            }

            // 念のため、形式チェック
            if (!wordsWithPos.every(item => item && typeof item.word === 'string' && typeof item.pos === 'string')) {
                 console.error("Invalid item format in OpenAI response:", wordsWithPos);
                 throw new Error("OpenAIからの応答データ形式が不正です。");
            }

        } catch (parseError) {
            console.error("Failed to parse OpenAI response:", parseError);
            console.error("Raw content:", completion.choices[0].message.content); // 生の応答内容をログに出力
            throw new Error("OpenAIからの応答の解析に失敗しました。");
        }


        // 出現頻度をカウント
        const wordCounts = {};
        const posMap = {}; // 単語と品詞のマップ

        textList.forEach(text => {
            // 簡易的な単語分割（より精度が必要なら形態素解析ライブラリを使う）
             const tokens = text.toLowerCase().replace(unwantedChars, ' ').split(/\s+/).filter(Boolean);
            tokens.forEach(token => {
                if (!stopWords.has(token) && token.length > 1) { // ストップワードと1文字を除外
                    // OpenAIのリストにある単語かチェック
                    const matchedWord = wordsWithPos.find(w => w.word === token);
                    if (matchedWord) {
                        wordCounts[token] = (wordCounts[token] || 0) + 1;
                        posMap[token] = matchedWord.pos; // 品詞情報を保存
                    }
                }
            });
        });

        // 結果を整形 (word, score(frequency), pos)
        const analysisResult = Object.entries(wordCounts).map(([word, count]) => ({
            word: word,
            score: count, // スコア = 出現頻度
            pos: posMap[word] || '不明' // 品詞情報
        }));

        // スコア（頻度）で降順ソート
        analysisResult.sort((a, b) => b.score - a.score);

        console.log(`Analysis complete. Found ${analysisResult.length} significant words.`);
        res.json({
            totalDocs: textList.length,
            results: analysisResult
        });

    } catch (error) {
        console.error("Error calling OpenAI API or processing results:", error);
        res.status(500).send(`テキスト解析中にエラーが発生しました: ${error.message}`);
    }
});
// --- ▲▲▲ APIエンドポイント ▲▲▲ ---


app.listen(PORT, () => {
  console.log(`Server listening on port ${PORT}`);
});
