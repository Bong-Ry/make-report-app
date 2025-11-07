const puppeteer = require('puppeteer');
const googleSheetsService = require('./googleSheets');
const kuromojiService = require('./kuromoji');

// --- Helper Functions and Constants (Replicated from front-end/old logic for server-side PDF) ---

const MAX_COMMENT_LINES_PER_PAGE = 20; // 1ページあたりのコメント行数上限
const AI_ANALYSIS_TYPES = ['L', 'I_bad', 'I_good', 'J', 'M'];

function getColumnName(columnType) { 
    switch(columnType){
        case'L':return'NPS推奨度 理由';
        case'I':case'I_good':case'I_bad':return'良かった点や悪かった点など';
        case'J':return'印象に残ったスタッフへのコメント';
        case'M':return'お産にかかわるご意見・ご感想';
        default:return'不明';
    } 
}
function getAnalysisTitle(columnType, count) { 
    const bt=`アンケート結果　ー${getColumnName(columnType)}ー`;
    return`${bt}　※全回答数${count}件ー`; 
}

function calculateNps(counts, totalCount) { 
    if (totalCount === 0) return 0;
    let promoters = 0, passives = 0, detractors = 0;
    for (let i = 0; i <= 10; i++) {
        const count = counts[i] || 0;
        if (i >= 9) promoters += count;
        else if (i >= 7) passives += count;
        else detractors += count;
    } 
    return ((promoters / totalCount) - (detractors / totalCount)) * 100; 
}

function chunkComments(commentsArray, title, score, totalCount) {
    const pages = [];
    if (!commentsArray || commentsArray.length === 0) {
        return pages;
    }
    
    let currentChunkLines = 0;
    let currentChunkText = '';
    let isFirstPage = true;
    
    commentsArray.forEach(comment => {
        const text = `・ ${String(comment).trim()}`;
        const linesInComment = (text.match(/\n/g) || []).length + 1;

        if (currentChunkLines > 0 && (currentChunkLines + linesInComment) > MAX_COMMENT_LINES_PER_PAGE) {
            pages.push({ title: '', body: currentChunkText, isCommentPage: true });
            currentChunkText = text;
            currentChunkLines = linesInComment;
            isFirstPage = false;
        } else {
            currentChunkText += (currentChunkLines === 0 ? '' : '\n') + text;
            currentChunkLines += linesInComment;
        }
    });
    
    if (currentChunkText) {
        let pageTitle = title;
        if (score !== undefined) {
             let imageTag = '';
             if (score === 10 || score === 9) imageTag = '画像3';
             else if (score === 8 || score === 7) imageTag = '画像4';
             else if (score <= 6 && score >= 0) imageTag = '画像5';
             pageTitle = `${imageTag}　推奨度${score} ${commentsArray.length}人`;
        }
        
        pages.push({ 
            title: isFirstPage ? pageTitle : '', 
            body: currentChunkText, 
            isCommentPage: true,
            totalCount: totalCount 
        });
    }
    
    return pages;
}

function getDetailedAnalysisTitleFull(analysisType) {
    switch (analysisType) {
        case 'L': return '知人に病院を紹介したいと思う理由の分析';
        case 'I_bad': return '「悪かった点」の分析と改善策';
        case 'I_good': return '「良かった点」の分析と改善策';
        case 'J': return '印象に残ったスタッフへのコメント分析とスタッフ評価';
        case 'M': return 'お産に関わるご意見の分析と改善策の提案';
        default: return 'AI詳細分析レポート';
    }
}
function getDetailedAnalysisSubtitle(analysisType, tabId) {
    const base = '※コメントでいただいたフィードバックを元に分析しています';
    if (tabId === 'analysis') {
        switch (analysisType) {
            case 'L': return base + '\n患者から寄せられたお産に関するご意見を分析すると、以下の主要なテーマが浮かび上がります。';
            case 'I_bad': return base + '\nフィードバックの中で挙げられた「悪かった点」を分析すると、患者にとって以下の要素が特に課題として感じられていることが分かります。';
            case 'I_good': return base + '\nフィードバックの中で挙げられた「良かった点」を分析すると、以下の要素が患者にとって特に高く評価されていることが分かります。';
            case 'J': return base + '\n印象に残ったスタッフに対するコメントから、いくつかの重要なテーマが浮かび上がります。\nこれらのテーマは、スタッフの評価においても重要なポイントとなります。';
            case 'M': return base + '\n患者から寄せられたお産に関するご意見を分析すると、以下の主要なテーマが浮かび上がります。';
            default: return base;
        }
    } else if (tabId === 'suggestions') {
        return base + (analysisType === 'L' ? '\n患者から寄せられたお産に関するご意見を分析すると、以下の主要なテーマが浮かび上がります。' : 
                       analysisType === 'I_bad' ? '\nフィードバックの中で挙げられた「悪かった点」を分析すると、患者にとって以下の要素が特に課題として感じられていることが分かります。' :
                       analysisType === 'I_good' ? '\nフィードバックの中で挙げられた「良かった点」を分析すると、以下の要素が患者にとって特に高く評価されていることが分かります。' : 
                       analysisType === 'J' ? '\n印象に残ったスタッフに対するコメントから、いくつかの重要なテーマが浮かび上がります。\nこれらのテーマは、スタッフの評価においても重要なポイントとなります。' : 
                       analysisType === 'M' ? '\n患者から寄せられたお産に関するご意見を分析すると、以下の主要なテーマが浮かび上がります。' : '');
    } else if (tabId === 'overall') {
        return base + '\n患者から寄せられたお産に関するご意見の分析と改善策を基にした、総評は以下のとおりです.';
    }
    return base;
}


// --- HTML Generation Functions ---

// 共通レイアウトのラッパーHTML
function getPageWrapper(title, subtitle, bodyHtml, separator=true) {
    const titleClass = 'report-title' + (title.includes('表紙') ? ' text-center' : ' text-left');
    const subTitleClass = 'report-subtitle text-right';
    const separatorHtml = separator ? '<hr class="report-separator">' : '';

    return `
        <div class="report-page">
            <h1 class="${titleClass}">${title}</h1>
            <p class="${subTitleClass}">${subtitle}</p>
            ${separatorHtml}
            <div class="report-content-pdf">
                ${bodyHtml}
            </div>
        </div>
    `;
}

// グラフページのHTMLシェル
function generateChartPageHtml(title, subtitle, type, isBar=false, clinicData, overallData) {
    const isNPSScore = type === 'nps_score';
    const cid = isBar ? 'bar-chart' : 'pie-chart';
    const hc = isBar ? 'h-[350px]' : 'h-[400px]';
    
    // NPSスコア計算
    const clinicNpsScore = calculateNps(clinicData.npsScoreData.counts, clinicData.npsScoreData.totalCount);
    const overallNpsScore = calculateNps(overallData.npsScoreData.counts, overallData.npsScoreData.totalCount);

    let bodyHtml;

    if (isNPSScore) {
        bodyHtml = `
            <div class="grid grid-cols-1 md:grid-cols-2 gap-8 items-start h-1/2">
                <div class="flex flex-col h-full">
                    <h3 class="font-bold text-lg mb-4 text-center">貴院の結果 (全 ${clinicData.npsScoreData.totalCount || 0} 件)</h3>
                    <div id="clinic-bar-chart" class="w-full h-[300px] bg-gray-50 border border-gray-200 flex items-center justify-center">
                        [グラフ描画エリア - スコア分布]
                    </div>
                </div>
                <div id="nps-summary-area" class="flex flex-col justify-center items-center space-y-6 pt-12">
                     <div class="text-left text-3xl space-y-5 p-6 border rounded-lg bg-gray-50 shadow-inner w-full max-w-xs"> 
                        <p>全体：<span class="font-bold text-gray-800">${overallNpsScore.toFixed(1)}</span></p> 
                        <p>貴院：<span class="font-bold text-red-600">${clinicNpsScore.toFixed(1)}</span></p> 
                    </div>
                </div>
            </div>
            <div class="w-full h-1/2 flex flex-col justify-center items-center pt-4">
                <p class="text-sm text-gray-500 mb-2">【画像入力エリア】</p>
                <div class="w-full h-full border border-dashed border-gray-300 flex items-center justify-center text-gray-400">
                    [画像を入力する]
                </div>
            </div>
        `;
    } else {
        bodyHtml = `
            <div class="grid grid-cols-1 md:grid-cols-2 gap-8 h-full">
                <div class="flex flex-col items-center h-full">
                    <h3 class="font-bold text-lg mb-4 text-center">貴院の結果</h3>
                    <div id="clinic-${cid}" class="w-full ${hc} bg-gray-50 border border-gray-200 flex items-center justify-center">
                        [グラフ描画エリア - 貴院]
                    </div>
                </div>
                <div class="flex flex-col items-center h-full">
                    <h3 class="font-bold text-lg mb-4 text-center">（参照）全体平均</h3>
                    <div id="average-${cid}" class="w-full ${hc} bg-gray-50 border border-gray-200 flex items-center justify-center">
                        [グラフ描画エリア - 全体]
                    </div>
                </div>
            </div>
        `;
    }

    return getPageWrapper(title, subtitle, bodyHtml);
}

// コメントページのHTMLシェル
function generateCommentPageHtml(baseTitle, pageTitle, commentsHtml) {
    let title = baseTitle.replace(/　※.+ー$/, ''); 
    let subtitle = baseTitle.includes('NPS推奨度 理由') ? '' : 'データ一覧（20データずつ）';

    return getPageWrapper(
        title,
        subtitle, 
        (pageTitle ? `<h3 class="font-bold text-lg mb-4">${pageTitle}</h3>` : '') + `<div class="comment-list-pdf">${commentsHtml}</div>`,
        false 
    );
}

// WCページのHTMLシェル
function generateWCPageHtml(title, totalCount, wordCloudData, kuromojiResults) {
    const wordCloudDataString = JSON.stringify(wordCloudData);
    
    // POSマップを生成
    const posMap = kuromojiResults.reduce((map, item) => { map[item.word] = item.pos; return map; }, {});
    const posMapString = JSON.stringify(posMap);
    
    const wcBody = `
        <div class="grid grid-cols-2 gap-4 h-full">
            <div class="space-y-2 overflow-y-auto h-full pr-2 flex flex-col">
                <div id="noun-chart" class="w-full h-1/4 bg-gray-50 border flex items-center justify-center">[グラフ描画エリア - 名詞]</div>
                <div id="verb-chart" class="w-full h-1/4 bg-gray-50 border flex items-center justify-center">[グラフ描画エリア - 動詞]</div>
                <div id="adj-chart" class="w-full h-1/4 bg-gray-50 border flex items-center justify-center">[グラフ描画エリア - 形容詞]</div>
                <div id="int-chart" class="w-full h-1/4 bg-gray-50 border flex items-center justify-center">[グラフ描画エリア - 感動詞]</div>
            </div>
            <div class="space-y-4 flex flex-col h-full">
                <p class="text-xs text-gray-600">
                    スコアが高い単語を複数選び出し、その値に応じた大きさで図示しています。<br>
                    単語の色は品詞の種類で異なります。<br>
                    <span class="text-blue-600 font-semibold">青色=名詞</span>、
                    <span class="text-red-600 font-semibold">赤色=動詞</span>、
                    <span class="text-green-600 font-semibold">緑色=形容詞</span>、
                    <span class="text-gray-600 font-semibold">灰色=感動詞</span>
                </p>
                <div id="word-cloud-container" class="h-full border border-gray-200 bg-gray-50">
                    <canvas id="word-cloud-canvas-pdf" data-words='${wordCloudDataString}' data-pos-map='${posMapString}' class="w-full h-full"></canvas>
                    <div class="flex items-center justify-center h-full text-gray-500">[ワードクラウド描画エリア]</div>
                </div>
            </div>
        </div>
    `;

    return getPageWrapper(
        getAnalysisTitle(title, totalCount),
        '章中に出現する単語の頻出度を表にしています。単語ごとに表示されている「スコア」の大きさは、その単語がどれだけ特徴的であるかを表しています。通常はその単語の出現回数が多いほどスコアが高くなるが、「言う」や「思う」など、どの文書にもよく現れる単語についてはスコアが低めになります。',
        wcBody
    );
}

// AI分析ページのHTMLシェル
function generateAiAnalysisPageHtml(type, tabId, content) {
    const aiTitle = tabId === 'analysis' ? '分析と考察' : (tabId === 'suggestions' ? '改善点' : '総評');
    
    const mainTitle = getDetailedAnalysisTitleFull(type);
    const subtitle = getDetailedAnalysisSubtitle(type, tabId);
    
    const bodyHtml = `
        <div class="ai-analysis-container">
            <div class="ai-analysis-sidebar">
                <div class="ai-analysis-shape">${aiTitle}</div>
            </div>
            <div class="ai-analysis-content text-sm whitespace-pre-wrap">${content}</div>
        </div>
    `;

    return getPageWrapper(mainTitle, subtitle, bodyHtml);
}

/**
 * PDF生成のメイン関数
 */
exports.generatePdfFromData = async (clinicName, periodText, clinicReportData) => {
    console.log("[pdfGeneratorService] Starting PDF generation...");

    // centralSheetId は reportController から渡される reportData には含まれないため、
    // clinicReportData が持つと仮定するか、別途渡す必要があります。
    // ここでは、一旦呼び出し元から取得できると仮定します。
    // (通常、reportDataにはcentralSheetIdは含まれないため、ここではダミーで設定)
    const centralSheetId = 'DUMMY_SHEET_ID'; // 実際には呼び出し元で修正が必要

    // --- 1. 全データ収集 ---
    let overallReportData, municipalityData, recommendationData, aiAnalysisData;
    let overallCount = 0;
    let totalClinics = 0;
    
    try {
        overallReportData = await googleSheetsService.getReportDataForCharts(clinicReportData.centralSheetId || centralSheetId, "全体");
        overallCount = overallReportData.npsScoreData.totalCount || 0;
        
        // 取り込み済みクリニック数を取得
        const sheetTitles = await googleSheetsService.getSheetTitles(clinicReportData.centralSheetId || centralSheetId);
        totalClinics = sheetTitles.length - (sheetTitles.includes('全体') ? 1 : 0) - (sheetTitles.includes('管理') ? 1 : 0);
        
        // 市区町村データを取得 (API呼び出しを模倣)
        const muniRes = await googleSheetsService.sheets.spreadsheets.values.get({ 
            spreadsheetId: clinicReportData.centralSheetId || centralSheetId, 
            range: `'${clinicName}_市区町村'!A:D`,
            valueRenderOption: 'UNFORMATTED_VALUE'
        });
        municipalityData = muniRes.data.values;
        if (municipalityData && municipalityData.length > 1) municipalityData.shift(); // ヘッダーを削除

        // おすすめ理由データを取得 (API呼び出しを模倣)
        const recRes = await googleSheetsService.sheets.spreadsheets.values.get({
            spreadsheetId: clinicReportData.centralSheetId || centralSheetId,
            range: `'${clinicName}_おすすめ理由'!A:C`,
            valueRenderOption: 'UNFORMATTED_VALUE'
        });
        recommendationData = recRes.data.values;
        if (recommendationData && recommendationData.length > 1) recommendationData.shift();

        // AI分析データを取得
        aiAnalysisData = await googleSheetsService.readAiAnalysisData(clinicReportData.centralSheetId || centralSheetId, `${clinicName}_AI分析`);

    } catch (e) {
        console.warn("[pdfGeneratorService] Failed to fetch auxiliary data (Overall/Muni/Rec/AI). Using partial data.", e.message);
        // エラーが発生しても、可能な限りレポートを生成する
    }


    // --- 2. 全ページHTML生成 ---
    const allPagesHtml = [];
    const [sy, sm] = periodText.split('～')[0].split('-').map(s => s.replace('年', '').replace('月', ''));
    const [ey, em] = periodText.split('～')[1].split('-').map(s => s.replace('年', '').replace('月', ''));
    const startDay = new Date(sy, sm - 1, 1).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });
    const endDay = new Date(ey, em, 0).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric' });
    
    const clinicCount = clinicReportData.npsScoreData.totalCount || 0;

    // --- 2a. 例外構成 (表紙, 目次, 概要) ---
    allPagesHtml.push(getPageWrapper(
        clinicName,
        '',
        '<div class="flex items-center justify-center h-full"><h2 class="text-4xl font-bold">アンケートレポート</h2></div>',
        false
    )); // 表紙

    allPagesHtml.push(getPageWrapper(
        '目次',
        '',
        `
        <div class="p-8">
            <ul class="text-2xl font-semibold space-y-4">
                <li>１．アンケート概要</li>
                <li>２．アンケート結果</li>
                <ul class="pl-8 space-y-2 font-normal">
                    <li>―１　顧客属性</li>
                    <li>―２　病院への満足度（施設・ハード面）</li>
                    <li>―３　病院への満足度（質・スタッフ面）</li>
                    <li>―４　NPS推奨度・理由</li>
                </ul>
                <li>３．アンケート結果からの考察</li>
            </ul>
        </div>
        `,
        false
    )); // 目次

    allPagesHtml.push(getPageWrapper(
        'アンケート概要',
        '',
        `
        <div class="p-8">
            <ul class="text-lg font-normal space-y-4">
                <li><span class="font-bold text-gray-800 w-32 inline-block">調査目的</span>｜貴院に対する満足度調査</li>
                <li><span class="font-bold text-gray-800 w-32 inline-block">調査方法</span>｜スマホ利用してのアンケートフォームによるインターネット調査</li>
                <li><span class="font-bold text-gray-800 w-32 inline-block">調査対象</span>｜貴院で出産された方（退院後～１か月健診までの期間）</li>
                <li><span class="font-bold text-gray-800 w-32 inline-block">調査期間</span>｜${startDay}〜${endDay}</li>
                <li><span class="font-bold text-gray-800 w-32 inline-block">回答件数</span>｜全体：${overallCount}件（${totalClinics}病院）　貴院：${clinicCount}件</li>
            </ul>
        </div>
        `,
        false
    )); // 概要
    
    // --- 2b. 基本構成（グラフ/テーブル） ---
    const chartPages = [
        { type: 'age', title: 'アンケート結果　ーご回答者さまの年代ー', subtitle: 'ご出産された方の年代について教えてください。' },
        { type: 'children', title: 'アンケート結果　ーご回答者さまのお子様の人数ー', subtitle: 'ご出産された方のお子様の人数について教えてください。' },
        { type: 'income', title: 'アンケート結果　ーご回答者さまの世帯年収ー', subtitle: 'ご出産された方の世帯年収について教えてください。', isBar: true },
        { type: 'municipality', title: 'アンケート結果　ーご回答者さまの市町村ー', subtitle: 'ご出産された方の住所（市町村）について教えてください。', isTable: true, data: municipalityData },
        { type: 'satisfaction_b', title: 'アンケート結果　ー満足度ー', subtitle: 'ご出産された産婦人科医院への満足度について、教えてください\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満' },
        { type: 'satisfaction_c', title: 'アンケート結果　ー施設の充実度・快適さー', subtitle: 'ご出産された産婦人科医院への施設の充実度・快適さについて、教えてください\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満' },
        { type: 'satisfaction_d', title: 'アンケート結果　ーアクセスの良さー', subtitle: 'ご出産された産婦人科医院へのアクセスの良さについて、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満' },
        { type: 'satisfaction_e', title: 'アンケート結果　ー費用ー', subtitle: 'ご出産された産婦人科医院への費用について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満' },
        { type: 'satisfaction_f', title: 'アンケート結果　ー病院の雰囲気ー', subtitle: 'ご出産された産婦人科医院への病院の雰囲気について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満' },
        { type: 'satisfaction_g', title: 'アンケート結果　ースタッフの対応ー', subtitle: 'ご出産された産婦人科医院へのスタッフの対応について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満' },
        { type: 'satisfaction_h', title: 'アンケート結果　ー先生の診断・説明ー', subtitle: 'ご出産された産婦人科医院への先生の診断・説明について、教えてください。\n＜5段階評価＞ 5:非常に満足〜 1:非常に不満' },
        { type: 'recommendation', title: 'アンケート結果　ー本病院を選ぶ上で最も参考にしたものー', subtitle: 'ご出産された産婦人科医院への本病院を選ぶ上で最も参考にしたものについて、教えてください。', isRec: true, data: recommendationData },
        { type: 'nps_score', title: 'アンケート結果　ーNPS(ネットプロモータースコア)＝推奨度ー', subtitle: 'これから初めてお産を迎える友人知人がいた場合、\nご出産された産婦人科医院をどのくらいお勧めしたいですか。\n友人知人への推奨度を教えてください。＜推奨度＞ 10:強くお勧めする〜 0:全くお勧めしない', isNPSScore: true }
    ];

    for (const page of chartPages) {
        if (page.isTable) {
             // 市区町村テーブル
             let tableHtml = `<div class="overflow-y-auto max-h-full border border-gray-200 rounded-lg"><table class="min-w-full divide-y divide-gray-200"><thead class="bg-gray-50 sticky top-0 z-10"><tr><th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">都道府県</th><th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">市区町村</th><th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">件数</th><th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">割合</th></tr></thead><tbody class="bg-white divide-y divide-gray-200">`;
             if (page.data && page.data.length > 0) {
                 page.data.forEach(row => {
                     const percentage = (parseFloat(row[3]) || 0) * 100;
                     tableHtml += `<tr><td class="px-6 py-4 whitespace-nowrap text-sm font-medium text-gray-900">${row[0] || ''}</td><td class="px-6 py-4 whitespace-nowrap text-sm text-gray-700">${row[1] || ''}</td><td class="px-6 py-4 whitespace-nowrap text-sm text-gray-700 text-right">${row[2] || 0}</td><td class="px-6 py-4 whitespace-nowrap text-sm text-gray-700 text-right">${percentage.toFixed(2)}%</td></tr>`;
                 });
             } else {
                 tableHtml += '<tr><td colspan="4" class="text-center text-gray-500 py-16">集計データがありません。</td></tr>';
             }
             tableHtml += '</tbody></table></div>';
             allPagesHtml.push(getPageWrapper(page.title, page.subtitle, tableHtml));
        } else if (page.isRec) {
             // おすすめ理由（円グラフ）- シェル生成を再利用
             allPagesHtml.push(generateChartPageHtml(page.title, page.subtitle, page.type, false, clinicReportData, overallReportData));
        } else {
            // その他のグラフページ
            allPagesHtml.push(generateChartPageHtml(page.title, page.subtitle, page.type, page.isBar, clinicReportData, overallReportData));
        }
    }

    // --- 2c. コメントページとWCページ ---
    const commentSections = [
        { type: 'L', data: clinicReportData.npsData.results, totalCount: clinicReportData.npsData.totalCount, isNPS: true, rawText: clinicReportData.npsData.rawText },
        { type: 'I', data: { 0: clinicReportData.feedbackData.i_column.results }, totalCount: clinicReportData.feedbackData.i_column.totalCount, isNPS: false, rawText: clinicReportData.feedbackData.i_column.results },
        { type: 'J', data: { 0: clinicReportData.feedbackData.j_column.results }, totalCount: clinicReportData.feedbackData.j_column.totalCount, isNPS: false, rawText: clinicReportData.feedbackData.j_column.results },
        { type: 'M', data: { 0: clinicReportData.feedbackData.m_column.results }, totalCount: clinicReportData.feedbackData.m_column.totalCount, isNPS: false, rawText: clinicReportData.feedbackData.m_column.results }
    ];
    
    // コメントリストページ
    for (const section of commentSections) {
        if (section.totalCount === 0) continue;
        
        const isNPS = section.isNPS;
        const baseTitle = `アンケート結果　ー${getColumnName(section.type)}　※全回答数${section.totalCount}件ー`;
        
        // NPSはスコア別にチャンク化
        if (isNPS) {
            const scores = Object.keys(section.data).map(Number).sort((a, b) => b - a);
            for (const score of scores) {
                const comments = section.data[score];
                const pages = chunkComments(comments, baseTitle, score, section.totalCount);
                for (const page of pages) {
                    allPagesHtml.push(generateCommentPageHtml(baseTitle, page.title, page.body));
                }
            }
        } else {
            // 一般コメントは一括チャンク化
            const allComments = section.data[0];
            const pages = chunkComments(allComments, baseTitle, undefined, section.totalCount);
            for (const page of pages) {
                allPagesHtml.push(generateCommentPageHtml(baseTitle, '', page.body));
            }
        }
    }
    
    // Word Cloudページ
    for (const section of commentSections) {
        if (section.totalCount === 0) continue;
        const textList = section.rawText;
        
        // Word Cloud分析の実行 (サーバーサイドで実行)
        let kuromojiResults = [];
        let wordCloudData = [];
        try {
            // Kuromojiの初期化が必要な場合があるため、ここではサービス呼び出しを試みる
            await kuromojiService.initializeKuromoji();
            const analysisResult = kuromojiService.analyzeTextList(textList);
            kuromojiResults = analysisResult.results;
            wordCloudData = kuromojiResults.map(r => [r.word, r.score]).slice(0, 100);
        } catch (e) {
            console.error(`[pdfGeneratorService] Kuromoji analysis failed for ${section.type}: ${e.message}`);
        }
        
        allPagesHtml.push(generateWCPageHtml(section.type, section.totalCount, wordCloudData, kuromojiResults));
    }


    // --- 2d. AI分析ページ (考察) ---
    allPagesHtml.push(getPageWrapper('アンケート結果からの考察', '', '<div class="flex items-center justify-center h-full"><h2 class="text-4xl font-bold">アンケート結果からの考察</h2></div>', false)); // 考察の区切りページ
    
    for (const type of AI_ANALYSIS_TYPES) {
        allPagesHtml.push(generateAiAnalysisPageHtml(type, 'analysis', aiAnalysisData.get(`${type}_ANALYSIS`) || '（データがありません）'));
        allPagesHtml.push(generateAiAnalysisPageHtml(type, 'suggestions', aiAnalysisData.get(`${type}_SUGGESTIONS`) || '（データがありません）'));
        allPagesHtml.push(generateAiAnalysisPageHtml(type, 'overall', aiAnalysisData.get(`${type}_OVERALL`) || '（データがありません）'));
    }


    // --- 3. HTMLラッパー生成 ---
    const finalHtml = `
        <!DOCTYPE html>
        <html lang="ja">
        <head>
          <meta charset="UTF-8">
          <title>レポート:${clinicName}</title>
          <style>
            @page {
                size: A4 landscape; /* A4横向き */
                margin: 0;
            }
            body { 
                font-family:'Noto Sans JP', sans-serif;
                margin: 0;
                padding: 0;
                width: 100vw;
                height: 100vh;
                background-color: #ffffff;
            }
            /* 固定レポートボディスタイル */
            .report-page {
                box-sizing: border-box;
                width: 297mm; /* A4横幅 */
                height: 210mm; /* A4縦幅 */
                margin: 0;
                padding: 40px; 
                border: 6px solid #fcf1ed; 
                border-radius: 0; 
                overflow: hidden;
                page-break-after: always;
            }
            .report-page:last-child {
                page-break-after: avoid;
            }
            
            .report-title { font-size: 24pt; font-weight: bold; margin-bottom: 10px; text-align: left; }
            .report-title.text-center { text-align: center; }
            .report-subtitle { font-size: 10pt; color: #6b7280; white-space: pre-wrap; text-align: right; }
            .report-separator { border-top: 1px dashed #9ca3af; margin-bottom: 20px; }
            
            .report-content-pdf { 
                height: calc(100% - 110px); 
                overflow: hidden; 
            }
            .comment-list-pdf { font-size: 10pt; white-space: pre-wrap; max-height: 100%; overflow: hidden; line-height: 1.5; }
            
            /* AI Analysis specific styles */
            .ai-analysis-container { display: flex; height: 100%; }
            .ai-analysis-sidebar { width: 20%; padding-right: 15px; display: flex; flex-direction: column; justify-content: flex-start; align-items: flex-start; height: 100%; }
            .ai-analysis-content { width: 80%; padding-left: 15px; font-size: 10pt; line-height: 1.5; white-space: pre-wrap; overflow-y: auto; max-height: 100%; }
            .ai-analysis-shape { 
                background-color: #fcf1ed; border: 1px solid #f9d8d1; /* 薄いピンク */
                color: #333; font-weight: bold; font-size: 14pt; padding: 10px; border-radius: 8px;
                clip-path: polygon(0% 0%, 100% 0%, 100% 75%, 75% 75%, 75% 100%, 50% 75%, 0% 75%);
                width: 100%; text-align: center; margin-bottom: 15px;
            }

            /* Chart/Table styles */
            .grid { display: grid; }
            .grid-cols-2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            .gap-8 { gap: 32px; }
            .h-full { height: 100%; }
            .h-1\\/2 { height: 50%; }
            .w-full { width: 100%; }
            .flex { display: flex; }
            .items-center { align-items: center; }
            .justify-center { justify-content: center; }
            .text-center { text-align: center; }
            .text-lg { font-size: 1.125rem; }
            .font-bold { font-weight: bold; }
            .mb-4 { margin-bottom: 16px; }
            .bg-gray-50 { background-color: #f9fafb; }
            .border { border: 1px solid #e5e7eb; }

            /* Table styles */
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 8px; border-bottom: 1px solid #e5e7eb; text-align: left; font-size: 8pt; }
            th { background-color: #f3f4f6; font-weight: bold; }

          </style>
        </head>
        <body>
            ${allPagesHtml.join('')}
        </body>
        <script>
            // PDF生成時のWordCloud描画シミュレーション（WordCloudライブラリはサーバー側で描画困難）
            function drawWordCloudPDF() {
                const canvas = document.getElementById('word-cloud-canvas-pdf');
                if (!canvas) return;

                const wordsString = canvas.getAttribute('data-words');
                const posMapString = canvas.getAttribute('data-pos-map');
                
                let list = [];
                let posMap = {};
                try {
                    list = JSON.parse(wordsString || '[]');
                    posMap = JSON.parse(posMapString || '{}');
                } catch (e) {
                    console.error('Failed to parse WordCloud data:', e);
                    return;
                }

                if (list.length === 0) {
                    canvas.parentNode.innerHTML = '<p class="text-center text-gray-500 py-16">分析結果なし</p>';
                    return;
                }
                
                // Puppeteerの環境ではWordCloudの描画が難しい場合があるため、
                // 簡易的な処理やメッセージを残す
                canvas.parentNode.innerHTML = '<div class="flex items-center justify-center h-full text-gray-500">[ワードクラウド描画エリア]</div>';

            }

            drawWordCloudPDF();
        </script>
        </html>
    `;

    // --- 4. PuppeteerでPDF生成 ---
    let browser;
    try {
        console.log("[pdfGeneratorService] Launching Puppeteer...");
        browser = await puppeteer.launch({
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--single-process',
                '--no-zygote'
            ],
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || puppeteer.executablePath()
        });
        const page = await browser.newPage();

        // A4横向きのビューポートに設定
        await page.setViewport({ width: 1122, height: 793 });

        console.log("[pdfGeneratorService] Setting HTML content...");
        await page.setContent(finalHtml, { waitUntil: 'networkidle0' });

        console.log("[pdfGeneratorService] Generating PDF buffer...");
        const pdfBuffer = await page.pdf({
            format: 'A4',
            landscape: true, 
            printBackground: true,
            margin: { top: '0', right: '0', bottom: '0', left: '0' } 
        });
        
        console.log("[pdfGeneratorService] PDF buffer generated successfully.");
        return pdfBuffer;

    } catch (error) {
        console.error('[pdfGeneratorService] Error during PDF generation:', error);
        throw error; 
    } finally {
        if (browser) {
            console.log("[pdfGeneratorService] Closing Puppeteer browser...");
            await browser.close();
        }
    }
};
