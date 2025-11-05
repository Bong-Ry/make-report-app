// services/googleSlidesService.js (新規作成)

const googleSheetsService = require('./googleSheets');
const { slides: slidesApi, GAS_SLIDE_GENERATOR_URL } = require('./googleSheets');
const { getSystemPromptForDetailAnalysis } = require('../utils/helpers'); // AI分析テキスト取得用

/**
 * =================================================================
 * メイン関数: スライド生成の全プロセスを実行
 * =================================================================
 * 1. GASを呼び出してスライドを複製＆要素マップ取得
 * 2. シートから集計データを取得
 * 3. AI分析データを取得
 * 4. スライドAPI (batchUpdate) で全データを挿入
 * @param {string} clinicName - クリニック名
 * @param {string} centralSheetId - 集計スプレッドシートID
 * @param {object} period - { start: 'YYYY-MM', end: 'YYYY-MM' }
 * @param {string} periodText - 'YYYY-MM～YYYY-MM'
 * @returns {string} - 生成された新しいスライドのURL
 */
exports.generateSlideReport = async (clinicName, centralSheetId, period, periodText) => {
    
    // --- 1. GASを呼び出し、スライド複製と要素マップ取得 ---
    console.log(`[googleSlidesService] Calling GAS to clone slide for: ${clinicName}`);
    const { newSlideId, newSlideUrl, analysisData } = await callGasToCloneSlide(clinicName);
    console.log(`[googleSlidesService] Slide cloned. New ID: ${newSlideId}`);

    // --- 2. 必要な全データを集計シートから取得 ---
    // (グラフ用データとAI分析の元テキストデータを一括で取得)
    console.log(`[googleSlidesService] Fetching aggregation data for: ${clinicName}`);
    const clinicReportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    const overallReportData = await googleSheetsService.getReportDataForCharts(centralSheetId, "全体");
    
    // --- 3. AI分析の結果をシートから取得 ---
    // (これはバックグラウンドで完了している前提)
    console.log(`[googleSlidesService] Fetching AI analysis text results...`);
    const aiTypes = ['L', 'I_bad', 'I_good', 'J', 'M'];
    const aiAnalysisResults = {};
    for (const type of aiTypes) {
        // (並列実行も可能だが、APIリミットを考慮し直列で取得)
        aiAnalysisResults[type] = await googleSheetsService.getAIAnalysisFromSheet(centralSheetId, clinicName, type);
    }
    console.log(`[googleSlidesService] All data fetched. Building update requests...`);

    // --- 4. スライドAPI (batchUpdate) で全データを挿入 ---
    
    // (A) プレースホルダーとObject IDのマッピングを作成
    // analysisData = [ [slideIdx, elemIdx, id, type, text], ... ]
    // を { "C8": "g123_45_6", "C17": "g123_45_7", ... } の形式に変換
    const placeholderMap = {};
    analysisData.forEach(row => {
        const objectId = row[2]; // ID
        const textContent = row[4]; // 内容 (e.g., "C8", "C134")
        if (textContent && textContent.startsWith('C')) {
            placeholderMap[textContent] = objectId;
        }
    });

    // (B) 挿入リクエストの配列を作成
    const requests = [];

    // (C) テキストデータを挿入 (ご指定のIDリスト)
    addTextRequests(requests, placeholderMap, period, clinicReportData, overallReportData, aiAnalysisResults);

    // (D) グラフ・画像を挿入 (TODO: ご指定のIDリスト)
    // TODO: addChartRequests(requests, placeholderMap, clinicReportData, overallReportData);

    // (E) コメントスライドを複製・挿入 (ご指定のIDリスト)
    // TODO: const npsRequests = await addNpsCommentSlides(requests, placeholderMap, clinicReportData, newSlideId);
    // TODO: requests.push(...npsRequests);

    // (F) batchUpdate を実行
    if (requests.length > 0) {
        console.log(`[googleSlidesService] Sending ${requests.length} update requests to Slides API...`);
        await slidesApi.presentations.batchUpdate({
            presentationId: newSlideId,
            resource: {
                requests: requests
            }
        });
        console.log(`[googleSlidesService] Slides API update complete.`);
    } else {
        console.log(`[googleSlidesService] No update requests to send.`);
    }

    // --- 5. 完了後、新しいスライドのURLを返す ---
    return newSlideUrl;
};


/**
 * =================================================================
 * ヘルパー (1/4): GAS Web App 呼び出し
 * =================================================================
 * @param {string} clinicName 
 * @returns {Promise<object>} - { newSlideId, newSlideUrl, analysisData }
 */
async function callGasToCloneSlide(clinicName) {
    try {
        const response = await fetch(GAS_SLIDE_GENERATOR_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ clinicName: clinicName })
        });

        if (!response.ok) {
            throw new Error(`GAS Web App (Slide Clone) request failed with status ${response.status}: ${await response.text()}`);
        }
        const result = await response.json();
        
        if (result.status === 'ok' && result.newSlideId && result.analysisData) {
            return {
                newSlideId: result.newSlideId,
                newSlideUrl: result.newSlideUrl,
                analysisData: result.analysisData // [ [slideIdx, elemIdx, id, type, text], ... ]
            };
        } else {
            console.error('[googleSlidesService] GAS Web App (Slide Clone) returned an error:', result.message);
            throw new Error(`GAS側でのスライド複製・分析に失敗: ${result.message || '不明なエラー'}`);
        }
    } catch (err) {
        console.error(`[googleSlidesService] Error in callGasToCloneSlide for "${clinicName}".`);
        console.error(err); 
        throw new Error(`スライド複製の呼び出しに失敗 (GAS): ${err.message}`);
    }
}

/**
 * =================================================================
 * ヘルパー (2/4): テキスト挿入リクエスト作成
 * =================================================================
 */
function addTextRequests(requests, placeholderMap, period, clinicData, overallData, aiData) {
    
    // --- 期間テキスト ---
    const [startY, startM] = period.start.split('-');
    const [endY, endM] = period.end.split('-');
    const periodFullText = `${startY}年${startM}月1日〜${endY}年${endM}月末日`;
    const periodShortText = `${startY}年${startM}月〜${endY}年${endM}月`;
    
    addTextUpdateRequest(requests, placeholderMap, 'C8', periodShortText);
    addTextUpdateRequest(requests, placeholderMap, 'C17', periodFullText);

    // --- NPS値 ---
    const clinicNpsScore = calculateNps(clinicData.npsScoreData.counts, clinicData.npsScoreData.totalCount);
    const overallNpsScore = calculateNps(overallData.npsScoreData.counts, overallData.npsScoreData.totalCount);
    addTextUpdateRequest(requests, placeholderMap, 'C124', clinicNpsScore.toFixed(1));
    addTextUpdateRequest(requests, placeholderMap, 'C125', overallNpsScore.toFixed(1));

    // --- NPSコメント人数 (スライド 17, 18, 19, 20) ---
    const npsResults = clinicData.npsData.results || {};
    addTextUpdateRequest(requests, placeholderMap, 'C135', `${(npsResults[10] || []).length}人`);
    addTextUpdateRequest(requests, placeholderMap, 'C143', `${(npsResults[9] || []).length}人`);
    addTextUpdateRequest(requests, placeholderMap, 'C152', `${(npsResults[8] || []).length}人`);
    addTextUpdateRequest(requests, placeholderMap, 'C164', `${(npsResults[7] || []).length}人`);
    const nps6BelowCount = (npsResults[6] || []).length + (npsResults[5] || []).length + (npsResults[4] || []).length + (npsResults[3] || []).length + (npsResults[2] || []).length + (npsResults[1] || []).length + (npsResults[0] || []).length;
    addTextUpdateRequest(requests, placeholderMap, 'C165', `${nps6BelowCount}人`);
    
    // --- NPSコメント本文 (スライド 17, 18, 19, 20) ---
    // (1ページ目のみ。スライド複製ロジックは別途実装)
    addTextUpdateRequest(requests, placeholderMap, 'C134', formatComments(npsResults[10]));
    addTextUpdateRequest(requests, placeholderMap, 'C142', formatComments(npsResults[9]));
    addTextUpdateRequest(requests, placeholderMap, 'C151', formatComments(npsResults[8]));
    addTextUpdateRequest(requests, placeholderMap, 'C162', formatComments(npsResults[7]));
    const nps6BelowComments = [
        ...(npsResults[6] || []), ...(npsResults[5] || []), ...(npsResults[4] || []),
        ...(npsResults[3] || []), ...(npsResults[2] || []), ...(npsResults[1] || []), ...(npsResults[0] || [])
    ];
    addTextUpdateRequest(requests, placeholderMap, 'C163', formatComments(nps6BelowComments));


    // --- AI分析テキスト ---
    // (aiData[type] は { analysis, suggestions, overall } というオブジェクト)
    addTextUpdateRequest(requests, placeholderMap, 'C215', aiData['L']?.analysis || '（データなし）');
    addTextUpdateRequest(requests, placeholderMap, 'C222', aiData['L']?.suggestions || '（データなし）');
    addTextUpdateRequest(requests, placeholderMap, 'C229', aiData['L']?.overall || '（データなし）');
    
    addTextUpdateRequest(requests, placeholderMap, 'C236', aiData['I_bad']?.analysis || '（データなし）');
    addTextUpdateRequest(requests, placeholderMap, 'C243', aiData['I_bad']?.suggestions || '（データなし）');
    // (I_bad には Overall がない)

    addTextUpdateRequest(requests, placeholderMap, 'C250', aiData['I_good']?.analysis || '（データなし）');
    addTextUpdateRequest(requests, placeholderMap, 'C257', aiData['I_good']?.suggestions || '（データなし）'); // モアポイント
    addTextUpdateRequest(requests, placeholderMap, 'C264', aiData['I_good']?.overall || '（データなし）');
    
    addTextUpdateRequest(requests, placeholderMap, 'C271', aiData['J']?.analysis || '（データなし）');
    addTextUpdateRequest(requests, placeholderMap, 'C278', aiData['J']?.suggestions || '（データなし）');
    addTextUpdateRequest(requests, placeholderMap, 'C286', aiData['J']?.overall || '（データなし）');

    addTextUpdateRequest(requests, placeholderMap, 'C293', aiData['M']?.analysis || '（データなし）');
    addTextUpdateRequest(requests, placeholderMap, 'C300', aiData['M']?.suggestions || '（データなし）');
    addTextUpdateRequest(requests, placeholderMap, 'C307', aiData['M']?.overall || '（データなし）');
}

/**
 * =================================================================
 * ヘルパー (3/4): グラフ・画像挿入リクエスト作成 (TODO)
 * =================================================================
 */
// TODO
// function addChartRequests(requests, placeholderMap, clinicData, overallData) {
//     // C25: 年代グラフ（貴院）
//     // C26: 年代グラフ（全体）
//     // ...
//     // C114: おすすめ理由（貴院）
//     // C115: おすすめ理由（全体）
//     // ...
//     // C128: NPSグラフ
//     // ...
//     // C171, C172 (NPS WC, Bar)
//     // C183, C182 (I WC, Bar)
//     // C193, C194 (J WC, Bar)
//     // C204, C205 (M WC, Bar)
// }

/**
 * =================================================================
 * ヘルパー (4/4): その他ユーティリティ
 * =================================================================
 */

/**
 * [Util] Google Slides API へのテキスト挿入リクエストを生成
 * @param {Array} requests - 変更リクエストの配列 (ここに追記される)
 * @param {object} placeholderMap - { "C8": "objectId_...", ... }
 * @param {string} placeholder - "C8" などのプレースホルダー
 * @param {string} newText - 挿入する新しいテキスト
 */
function addTextUpdateRequest(requests, placeholderMap, placeholder, newText) {
    const objectId = placeholderMap[placeholder];
    if (!objectId) {
        console.warn(`[googleSlidesService] Placeholder "${placeholder}" not found in slide map. Skipping update.`);
        return;
    }
    
    // 既存のテキストをすべて削除するリクエスト
    requests.push({
        deleteText: {
            objectId: objectId,
            textRange: { type: 'ALL' }
        }
    });
    
    // 新しいテキストを挿入するリクエスト
    requests.push({
        insertText: {
            objectId: objectId,
            insertionIndex: 0,
            text: newText || '（データなし）' // null や undefined を防ぐ
        }
    });
}

/**
 * [Util] NPSスコアを計算 (フロントエンドから移植)
 */
function calculateNps(counts, totalCount) {
    if (!counts || totalCount === 0) return 0;
    let promoters = 0, detractors = 0;
    for (let i = 0; i <= 10; i++) {
        const count = counts[i] || 0;
        if (i >= 9) promoters += count;
        else if (i <= 6) detractors += count; // 7,8はパッシブ
    }
    return ((promoters / totalCount) - (detractors / totalCount)) * 100;
}

/**
 * [Util] コメント配列をスライド挿入用の単一文字列にフォーマット
 */
function formatComments(commentsArray) {
    if (!commentsArray || commentsArray.length === 0) {
        return '（該当コメントなし）';
    }
    // スライドのテキストボックスは \n で改行できる
    return commentsArray.map(c => `・ ${c.trim()}`).join('\n');
}
