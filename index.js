const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const { google } = require('googleapis');
const kuromoji = require('kuromoji'); // kuromoji ライブラリ

const app = express();
app.use(express.json({ limit: '10mb' }));
const PORT = process.env.PORT || 3000;

// --- Render環境設定読み込み ---
const KEYFILEPATH = '/etc/secrets/credentials.json';
const MASTER_SPREADSHEET_ID = process.env.MASTER_SPREADSHEET_ID;
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

// --- kuromoji Tokenizer 初期化 ---
let tokenizer;
kuromoji.builder({ dicPath: path.join(__dirname, 'node_modules', 'kuromoji', 'dict') })
    .build((err, _tokenizer) => {
        if (err) {
            console.error('Failed to build kuromoji tokenizer:', err);
        } else {
            tokenizer = _tokenizer;
            console.log('kuromoji tokenizer built successfully.');
        }
    });
// ---------------------------------

// --- ▼▼▼ ストップワードリスト更新 ▼▼▼ ---
const stopWords = new Set([
  'こと', 'もの', 'ため', 'よう', 'とき', 'ところ', 'うち', 'わけ', 'はず', 'つもり', '上', '下', '方',
  '私', 'あなた', '彼', '彼女', 'これ', 'それ', 'あれ', 'ここ', 'そこ', 'あそこ', 'どの', 'この', 'その', 'あの',
  'する', 'いる', 'ある', 'なる', 'できる', 'いく', 'くる', 'いう', '思う', '感じる', '見る', '聞く', '言う', '行く', '来る', '与える', '受ける', '持つ', '取る', 'かける', 'くれる', '下さる', 'いただく', 'おる', 'くださる', '頂く', '下さい', '致す', '参る', '申す', '存じる', '拝見', '伺う', 'おります', 'ございます', 'お願', '感じ', '思い', '考え', // 「お願い」の「い」を除く
  'いい', 'よい', '悪い', '高い', '低い', '多い', '少ない', '大きい', '小さい', '長い', '短い', '早い', '遅い', '新しい', '古い',
  'ない', 'なく', 'です', 'ます', 'でし', 'まし', 'ませ', 'ん', 'たい', 'たがる', 'らしい', 'そうだ', 'ようだ', 'みたいだ',
  'の', 'が', 'を', 'に', 'へ', 'と', 'より', 'から', 'で', 'や', 'も', 'は',
  'など', '等', '的', '性', '化', '感', 'さ', 'み', 'ながら', 'つつ', 'て', 'ば', 'たり', 'のみ', 'だけ', 'しか', 'さえ', 'まで', 'こそ', 'でも', 'だの', 'なり', 'やら', 'か', 'のやら', 'とか', 'だって', 'とも', 'ても', 'けれど', 'けれども', 'けど', 'のに', 'ので', 'し', 'ものの', 'くせに', 'ところが', 'それでも', 'だから', 'すると', 'そこで', 'さて', 'では', 'もし', 'たとえ', 'いくら', 'どんなに', 'よしんば', 'かりに',
  'いや', 'はい', 'ええ', 'ああ', 'うん', 'まあ', 'さあ', 'ねえ', 'おい', 'こら', 'もしもし', 'おお', 'あら', 'わあ', 'いやはや', 'それでは', 'じゃあ', 'ところで', 'そして', 'それから', 'ならびに', 'および', 'また', 'かつ', 'あるいは', 'または', 'ないしは', 'もしくは', 'それとも', 'しかし', 'だが', 'だって', 'なぜなら', 'というのは', 'もっとも', 'ただし',
  'おはよう', 'こんにちは', 'こんばんは', 'さようなら', 'ありがとう', 'すみません', 'ごめんなさい', 'ください',
  '様', 'さん', '君', 'ちゃん',
  '非常', '大変', '本当', 'とても', 'すごく', '特に', '少し', 'ちょっと', 'いつも', '全て', '皆様', 'みなさま', '皆さん', 'みなさん', '方々'
]);
// --- ▲▲▲ ストップワードリスト更新 ▲▲▲ ---

function getSpreadsheetIdFromUrl(url) { if (!url || typeof url !== 'string') return null; const match = url.match(/\/d\/(.+?)\//); return match ? match[1] : null; }

app.get('/', (req, res) => { res.sendFile(path.join(__dirname, 'public', 'index.html')); });
app.use(express.static(path.join(__dirname, 'public')));

app.post('/generate-pdf', async (req, res) => { /* (変更なし) */ console.log("POST /generate-pdf called");const{clinicName,periodText,reportData}=req.body;if(!clinicName||!periodText||!reportData){return res.status(400).send('PDF生成に必要なデータが不足');} let pdfHtml=`<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8"><title>レポート:${clinicName}</title><style>body{font-family:'Noto Sans JP',sans-serif;padding:20px;} h1,h2{border-bottom:1px solid #ccc;padding-bottom:5px;} .chart-container{text-align:center;margin-bottom:20px;} .comment-list{margin-top:10px;padding-left:20px;white-space:pre-wrap;font-size:10pt;}</style></head><body><h1>レポート:${clinicName}</h1><p>集計期間:${periodText}</p><hr>`;pdfHtml+=`<h2>NPS理由(全${reportData.npsData?.totalCount||0}件)</h2>`;if(reportData.npsData?.results){const scores=Object.keys(reportData.npsData.results).map(Number).sort((a,b)=>b-a);scores.forEach(score=>{const reasons=reportData.npsData.results[score];if(reasons&&reasons.length>0){pdfHtml+=`<h3>推奨度 ${score}(${reasons.length}人)</h3><ul class="comment-list">`;reasons.forEach(reason=>{const escapedReason=reason.replace(/</g,"&lt;").replace(/>/g,"&gt;");pdfHtml+=`<li>${escapedReason}</li>`;});pdfHtml+=`</ul>`;}});}else{pdfHtml+=`<p>データなし</p>`;} pdfHtml+=`<hr>`;/* TODO: 他レポート追加 */ pdfHtml+=`</body></html>`;try{console.log("Launching Puppeteer...");const browser=await puppeteer.launch({args:['--no-sandbox','--disable-setuid-sandbox','--disable-dev-shm-usage','--executablePath=/usr/bin/google-chrome']});const page=await browser.newPage();console.log("Setting HTML...");await page.setContent(pdfHtml,{waitUntil:'networkidle0'});console.log("Generating PDF...");const pdf=await page.pdf({format:'A4',printBackground:true,margin:{top:'20mm',right:'10mm',bottom:'20mm',left:'10mm'}});await browser.close();console.log("Puppeteer closed.");res.contentType('application/pdf');const fileName=`${clinicName}_${periodText.replace(/～/g,'-')}_レポート.pdf`;res.setHeader('Content-Disposition',`attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);res.send(pdf);}catch(error){console.error('PDF generation failed:',error);res.status(500).send(`PDF生成失敗: ${error.message}`);} });
app.get('/api/getClinicList', async (req, res) => { /* (変更なし) */ console.log('GET /api/getClinicList called'); if(!sheets)return res.status(500).send('Sheets API client not initialized.');if(!MASTER_SPREADSHEET_ID){console.error('MASTER_SHEET_ID env var not set.');return res.status(500).send('Server config error: MASTER_SHEET_ID missing.');} const MASTER_RANGE='シート1!A2:A';try{const response=await sheets.spreadsheets.values.get({spreadsheetId:MASTER_SPREADSHEET_ID,range:MASTER_RANGE});const rows=response.data.values;const clinics=(rows&&rows.length>0)?rows.map(row=>row[0]).filter(Boolean):[];console.log('Fetched clinics:',clinics);res.json(clinics);}catch(err){console.error('Master Sheet API error:'+err);res.status(500).send('マスターシート読込失敗');} });
app.post('/api/getReportData', async (req, res) => { /* (変更なし) */ const{period,selectedClinics}=req.body;console.log('POST /api/getReportData called');console.log('Period:',period);console.log('Clinics:',selectedClinics);if(!sheets)return res.status(500).send('Sheets API client not initialized.');if(!MASTER_SPREADSHEET_ID){console.error('MASTER_SHEET_ID env var not set.');return res.status(500).send('Server config error: MASTER_SHEET_ID missing.');} if(!period||!selectedClinics||!Array.isArray(selectedClinics)){return res.status(400).send('Invalid request');} const MASTER_CLINIC_URL_RANGE='シート1!A2:B';try{const masterResponse=await sheets.spreadsheets.values.get({spreadsheetId:MASTER_SPREADSHEET_ID,range:MASTER_CLINIC_URL_RANGE});const masterRows=masterResponse.data.values;if(!masterRows||masterRows.length===0){console.log('No clinic/URL data in master sheet.');return res.json({});} const clinicUrls={};masterRows.forEach(row=>{const clinicName=row[0],sheetUrl=row[1];if(selectedClinics.includes(clinicName)&&sheetUrl){const sheetId=getSpreadsheetIdFromUrl(sheetUrl);if(sheetId)clinicUrls[clinicName]=sheetId;else console.warn(`Invalid URL for ${clinicName}`);}});console.log('Target Clinic Sheet IDs:',clinicUrls);const startDate=new Date(period.start+'-01T00:00:00Z');const endDate=new Date(period.end.split('-')[0],period.end.split('-')[1],0);endDate.setHours(23,59,59,999);console.log(`Filtering data between ${startDate.toISOString()} and ${endDate.toISOString()}`);const reportData={};const satisfactionKeys=['非常に満足','満足','ふつう','不満','非常に不満'];const ageKeys=['10代','20代','30代','40代'];const childrenKeys=['1人','2人','3人','4人','5人以上'];const initializeCounts=(keys)=>keys.reduce((acc,key)=>{acc[key]=0;return acc;},{});const createChartData=(counts,keys)=>{const chartData=[['カテゴリ','件数']];keys.forEach(key=>{if(counts[key]>0)chartData.push([key,counts[key]]);});return chartData;};for(const clinicName in clinicUrls){const clinicSheetId=clinicUrls[clinicName];console.log(`Processing ${clinicName} (ID: ${clinicSheetId})`);const allNpsReasons=[],allFeedbacks_I=[],allFeedbacks_J=[],allFeedbacks_M=[];const satisfactionCounts_B=initializeCounts(satisfactionKeys),satisfactionCounts_C=initializeCounts(satisfactionKeys),satisfactionCounts_D=initializeCounts(satisfactionKeys),satisfactionCounts_E=initializeCounts(satisfactionKeys),satisfactionCounts_F=initializeCounts(satisfactionKeys),satisfactionCounts_G=initializeCounts(satisfactionKeys),satisfactionCounts_H=initializeCounts(satisfactionKeys);const childrenCounts_P=initializeCounts(childrenKeys);const ageCounts_O=initializeCounts(ageKeys);const incomeCounts={1:0,2:0,3:0,4:0,5:0,6:0,7:0,8:0,9:0,10:0};try{const clinicDataResponse=await sheets.spreadsheets.values.get({spreadsheetId:clinicSheetId,range:"'フォームの回答 １'!A:Q",dateTimeRenderOption:'SERIAL_NUMBER',valueRenderOption:'UNFORMATTED_VALUE'});const clinicDataRows=clinicDataResponse.data.values;if(!clinicDataRows||clinicDataRows.length<2){console.log(`No data for ${clinicName}.`);continue;} const header=clinicDataRows.shift();const timestampIndex=0,satBIndex=1,satCIndex=2,satDIndex=3,satEIndex=4,satFIndex=5,satGIndex=6,satHIndex=7,feedbackIIndex=8,feedbackJIndex=9,scoreKIndex=10,reasonLIndex=11,feedbackMIndex=12,ageOIndex=14,childrenPIndex=15,incomeQIndex=16;clinicDataRows.forEach(row=>{const excelEpoch=new Date(Date.UTC(1899,11,30));const serialValue=row[timestampIndex];if(typeof serialValue!=='number'||serialValue<=0)return;const timestamp=new Date(excelEpoch.getTime()+serialValue*24*60*60*1000);if(timestamp.getTime()>=startDate.getTime()&&timestamp.getTime()<=endDate.getTime()){const score=row[scoreKIndex],reason=row[reasonLIndex];if(reason!=null&&String(reason).trim()!=='')allNpsReasons.push({score:parseInt(score,10),reason:String(reason).trim()});const feedbackI=row[feedbackIIndex];if(feedbackI!=null&&String(feedbackI).trim()!=='')allFeedbacks_I.push(String(feedbackI).trim());const feedbackJ=row[feedbackJIndex];if(feedbackJ!=null&&String(feedbackJ).trim()!=='')allFeedbacks_J.push(String(feedbackJ).trim());const feedbackM=row[feedbackMIndex];if(feedbackM!=null&&String(feedbackM).trim()!=='')allFeedbacks_M.push(String(feedbackM).trim());const satB=row[satBIndex];if(satB!=null&&satisfactionKeys.includes(String(satB)))satisfactionCounts_B[String(satB)]++;const satC=row[satCIndex];if(satC!=null&&satisfactionKeys.includes(String(satC)))satisfactionCounts_C[String(satC)]++;const satD=row[satDIndex];if(satD!=null&&satisfactionKeys.includes(String(satD)))satisfactionCounts_D[String(satD)]++;const satE=row[satEIndex];if(satE!=null&&satisfactionKeys.includes(String(satE)))satisfactionCounts_E[String(satE)]++;const satF=row[satFIndex];if(satF!=null&&satisfactionKeys.includes(String(satF)))satisfactionCounts_F[String(satF)]++;const satG=row[satGIndex];if(satG!=null&&satisfactionKeys.includes(String(satG)))satisfactionCounts_G[String(satG)]++;const satH=row[satHIndex];if(satH!=null&&satisfactionKeys.includes(String(satH)))satisfactionCounts_H[String(satH)]++;const childrenP=row[childrenPIndex];if(childrenP!=null&&childrenKeys.includes(String(childrenP)))childrenCounts_P[String(childrenP)]++;const ageO=row[ageOIndex];if(ageO!=null&&ageKeys.includes(String(ageO)))ageCounts_O[String(ageO)]++;const income=row[incomeQIndex];if(typeof income==='number'&&income>=1&&income<=10)incomeCounts[income]++;}});}catch(e){console.error(`Error processing sheet for ${clinicName}: ${e.toString()}`);continue;} const groupedByScore=allNpsReasons.reduce((acc,item)=>{if(!acc[item.score])acc[item.score]=[];acc[item.score].push(item.reason);return acc;},{});const incomeChartData=[['評価','割合',{role:'annotation'}]];const totalIncomeCount=Object.values(incomeCounts).reduce((a,b)=>a+b,0);if(totalIncomeCount>0){for(let i=1;i<=10;i++){const percentage=(incomeCounts[i]/totalIncomeCount)*100;incomeChartData.push([String(i),percentage,`${Math.round(percentage)}%`]);}} reportData[clinicName]={npsData:{totalCount:allNpsReasons.length,results:groupedByScore,rawText:allNpsReasons.map(r=>r.reason)},feedbackData:{i_column:{totalCount:allFeedbacks_I.length,results:allFeedbacks_I},j_column:{totalCount:allFeedbacks_J.length,results:allFeedbacks_J},m_column:{totalCount:allFeedbacks_M.length,results:allFeedbacks_M}},satisfactionData:{b_column:{results:createChartData(satisfactionCounts_B,satisfactionKeys)},c_column:{results:createChartData(satisfactionCounts_C,satisfactionKeys)},d_column:{results:createChartData(satisfactionCounts_D,satisfactionKeys)},e_column:{results:createChartData(satisfactionCounts_E,satisfactionKeys)},f_column:{results:createChartData(satisfactionCounts_F,satisfactionKeys)},g_column:{results:createChartData(satisfactionCounts_G,satisfactionKeys)},h_column:{results:createChartData(satisfactionCounts_H,satisfactionKeys)}},ageData:{results:createChartData(ageCounts_O,ageKeys)},childrenCountData:{results:createChartData(childrenCounts_P,childrenKeys)},incomeData:{results:incomeChartData,totalCount:totalIncomeCount}};console.log(`Finished processing data for ${clinicName}`);} console.log('Finished all clinics. Sending report data.');res.json(reportData);}catch(err){console.error('Error in /api/getReportData:',err);res.status(500).send('レポートデータ取得エラー');} });

// --- テキスト分析 API (kuromoji + ストップワード) ---
app.post('/api/analyzeText', async (req, res) => {
    console.log('POST /api/analyzeText called (using kuromoji + stop words)');
    const { textList } = req.body;

    if (!tokenizer) { return res.status(500).send('形態素解析器が初期化されていません。'); }
    if (!textList || !Array.isArray(textList) || textList.length === 0) { return res.status(400).send('解析対象のテキストリストが必要です。'); }

    console.log(`Analyzing ${textList.length} texts using kuromoji...`);
    try {
        const wordCounts = {};
        let processedTokensCount = 0;

        for (const text of textList) {
            if (!text || typeof text !== 'string') continue;
            const tokens = tokenizer.tokenize(text);
            processedTokensCount += tokens.length;

            tokens.forEach(token => {
                const pos = token.pos; const posDetail1 = token.pos_detail_1;
                // ★ 基本形があればそれを、なければ表層形を使う
                const wordBase = token.basic_form !== '*' ? token.basic_form : token.surface_form;

                let targetPosCategory = null;
                // 品詞フィルタリング (前回と同じ)
                if (pos === '名詞' && ['一般', '固有名詞', 'サ変接続', '形容動詞語幹'].includes(posDetail1)) { targetPosCategory = '名詞'; }
                else if (pos === '動詞' && (posDetail1 === '自立' || posDetail1 === 'サ変接続')) { targetPosCategory = '動詞'; }
                else if (pos === '形容詞' && posDetail1 === '自立') { targetPosCategory = '形容詞'; }
                else if (pos === '感動詞') { targetPosCategory = '感動詞'; }

                // 対象品詞で、かつ1文字より長く、かつ数字でなく、かつストップワードでない
                if (targetPosCategory && wordBase.length > 1 && !/^[0-9]+$/.test(wordBase) && !stopWords.has(wordBase)) {
                    if (!wordCounts[wordBase]) {
                        wordCounts[wordBase] = { count: 0, pos: targetPosCategory };
                    }
                    wordCounts[wordBase].count++;
                }
            });
        }
        console.log(`Processed ${processedTokensCount} tokens. Filtered by POS and stop words.`);

        const analysisResult = Object.entries(wordCounts).map(([word, data]) => ({
            word: word,
            score: data.count,
            pos: data.pos
        }));
        analysisResult.sort((a, b) => b.score - a.score);

        console.log(`kuromoji analysis complete. Found ${analysisResult.length} significant words.`);

        res.json({
            totalDocs: textList.length,
            results: analysisResult
        });
    } catch (error) { console.error("Error during kuromoji analysis:", error); res.status(500).send(`テキスト解析エラー: ${error.message}`); }
});
// ------------------------------------------

// --- サーバー起動 ---
app.listen(PORT, () => { console.log(`Server listening on port ${PORT}`); });
// ------------------
