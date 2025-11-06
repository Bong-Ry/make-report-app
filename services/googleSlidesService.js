// bong-ry/make-report-app/make-report-app-760b11a6c865c8524de321636e02a3a77d24e0d4/services/googleSlidesService.js
// (構文エラーを修正し、画像/グラフ挿入ロジックを *完全に* 削除した安定版)

// ▼▼▼ [修正] 必要なモジュールのみインポート ▼▼▼
const googleSheetsService = require('./googleSheets');
// (aiAnalysisService はここでは不要)
// (kuromojiService はここでは不要)
const { slides: slidesApi, GAS_SLIDE_GENERATOR_URL } = require('./googleSheets');
const { getAnalysisSheetName } = require('../utils/helpers');
// ▲▲▲

// ▼▼▼ [修正] L12 の孤立した } catch (err) { を削除 ▼▼▼
// (ここにあった構文エラーを削除)
// ▲▲▲

// ▼▼▼ [削除] 画像生成ヘルパー関数群 (すべて削除) ▼▼▼
// async function generateChartImageBuffer(...) { ... }
// async function generateWcImageBuffer(...) { ... }
// async function generateWordCloudImageBuffer(...) { ... }
// function addImageUpdateRequest(...) { ... }
// ▲▲▲

/**
 * =================================================================
 * メイン関数: スライド生成の全プロセスを実行
 * =================================================================
 */
exports.generateSlideReport = async (clinicName, centralSheetId, period, periodText) => {
    
    // --- 1. GASを呼び出し、スライド複製と要素マップ取得 ---
    console.log(`[googleSlidesService] Calling GAS to clone slide for: ${clinicName}`);
    
    const { newSlideId, newSlideUrl, analysisData } = await callGasToCloneSlide(clinicName, centralSheetId);
    console.log(`[googleSlidesService] Slide cloned. New ID: ${newSlideId}`);

    if (!analysisData || analysisData.length === 0) {
        throw new Error('GAS analysis data (analysisData) is empty or undefined. (GASがグループ内を検索してもプレースホルダーテキストを見つけられなかった可能性があります)');
    }

    // --- 2. 必要な「集計」データをシートから取得 ---
    console.log(`[googleSlidesService] Fetching aggregation data for: ${clinicName}`);
    const clinicReportData = await googleSheetsService.getReportDataForCharts(centralSheetId, clinicName);
    const overallReportData = await googleSheetsService.getReportDataForCharts(centralSheetId, "全体");
    
    // --- 3. 必要な「AI分析」データをシートから取得 ---
    console.log(`[googleSlidesService] Fetching AI analysis text results...`);
    const aiTypes = ['L', 'I_bad', 'I_good', 'J', 'M'];
    const aiAnalysisResults = {}; 
    
    const aiSheetName = getAnalysisSheetName(clinicName, 'AI');
    // ▼▼▼ [修正] AIデータのみ取得 (行数カウント、市区町村テキスト取得を削除) ▼▼▼
    const aiDataMap = await googleSheetsService.readAiAnalysisData(centralSheetId, aiSheetName);
    // ▲▲▲
    
    console.log(`[googleSlidesService] Fetched AI Data Map from "${aiSheetName}".`);
    
    for (const type of aiTypes) {
         aiAnalysisResults[type] = {
            analysis: aiDataMap.get(`${type}_ANALYSIS`) || '（データなし）',
            suggestions: aiDataMap.get(`${type}_SUGGESTIONS`) || '（データなし）',
            overall: aiDataMap.get(`${type}_OVERALL`) || '（データなし）'
        };
    }
    
    // ▼▼▼ [削除] BH (カウント) 用のデータ準備を削除 ▼▼▼
    // const reportCounts = { ... };
    // ▲▲▲
    
    // --- 4. スライドAPI (batchUpdate) で全データを挿入 ---
    
    console.log(`[googleSlidesService] Building placeholder map from GAS analysisData (Sheet) by *searching* placeholder text...`);
    
    const placeholderMap = {}; // "A" -> "g120" (elementId)
    const slideTemplateMap = {}; // "AC" -> "g_slide_22" (slideId)

    /**
     * [修正] getIdsFromCellRef (空白・大文字小文字無視)
     */
    const getIdsFromCellRef = (ref) => {
        const searchKey = ref.trim().toLowerCase();
        const row = analysisData.find(r => {
            const valueInSheet = r[4]; // E列 (index 4) の値
            if (typeof valueInSheet === 'string') {
                return valueInSheet.trim().toLowerCase() === searchKey;
            }
            return false;
        });

        if (!row) {
            console.warn(`[googleSlidesService] Placeholder key "${ref}" (searching as "${searchKey}") not found in E-column (index 4) of analysisData sheet.`);
            return { elementId: null, slideId: null };
        }
        
        const elementId = row[2] || null; // C列 (index 2) が elementId
        const slideId = row[5] || null; // F列 (index 5) が slideId
        
        console.log(`[googleSlidesService] Found mapping: ${ref} -> elementId: ${elementId}, slideId: ${slideId}`);
        return { elementId, slideId };
    };

    // ▼▼▼ [修正] プレースホルダーのキーを再分類 (画像キーを削除) ▼▼▼
    
    // (1) 複製なし・テキスト / 値 / AI分析
    const allPlaceholders = [
        "A", "B", "I", // 基本情報・表
        "AA", "AB", // NPS値
        "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", // AI分析
        "BH", // 集計値
        "BI", "BJ", "BK", "BL", "BM" // NPS人数
    ];
    
    // (2) 複製あり・コメント
    const commentTemplatePlaceholders = [
        "AC", "AD", "AE", "AF", "AG", // NPSコメント
        "AK", // 良かった点コメント
        "AN", // スタッフコメント
        "AQ" // お産意見コメント
    ];

    // (3) ▼▼▼ [削除] グラフ / 画像キー (すべて削除) ▼▼▼
    // const imagePlaceholders = [ ... ];
    // ▲▲▲
    
    // 1. 通常のテキスト/値/AI用ID (A, B, AA, AT...) を取得
    for (const ref of allPlaceholders) {
         const { elementId } = getIdsFromCellRef(ref); 
         if (elementId) {
             placeholderMap[ref] = elementId;
         }
    }
    
    // 2. コメント複製用テンプレートのID (AC, AD...) を取得
    for (const ref of commentTemplatePlaceholders) {
         const { elementId, slideId } = getIdsFromCellRef(ref); 
         if (elementId) {
             placeholderMap[ref] = elementId;
         }
         if (slideId) {
             slideTemplateMap[ref] = slideId;
         }
    }

    // 3. ▼▼▼ [削除] グラフ/画像用のID取得ループ (削除) ▼▼▼
    // for (const ref of imagePlaceholders) { ... }
    // ▲▲▲
    
    console.log(`[googleSlidesService] Placeholder Map (Text/Value/AI):\n`, placeholderMap);
    console.log(`[googleSlidesService] Slide Template Map (Comments):\n`, slideTemplateMap);


    // (B) 挿入リクエストの配列を作成
    const requests = [];

    // (C) テキストデータを挿入 (A, B, AA, AT... など)
    // ▼▼▼ [修正] addTextRequests から市区町村テキスト(I)とカウント(BH)の引数を削除 ▼▼▼
    addTextRequests(requests, placeholderMap, period, clinicReportData, overallReportData, aiAnalysisResults);

    // (D) コメントスライドを複製・挿入
    console.log(`[googleSlidesService] Generating comment slide duplication requests...`);
    
    // ▼▼▼ [修正なし] コメントキー (AC〜AQ) はそのまま ▼▼▼
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[10] || []), "AC");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[9] || []), "AD");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[8] || []), "AE");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[7] || []), "AF");
    const nps6BelowComments = [
        ...(clinicReportData.npsData.results[6] || []), ...(clinicReportData.npsData.results[5] || []), ...(clinicReportData.npsData.results[4] || []),
        ...(clinicReportData.npsData.results[3] || []), ...(clinicReportData.npsData.results[2] || []), ...(clinicReportData.npsData.results[1] || []), ...(clinicReportData.npsData.results[0] || [])
    ];
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, nps6BelowComments, "AG");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, clinicReportData.feedbackData.i_column.results, "AK");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, clinicReportData.feedbackData.j_column.results, "AN"); 
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, clinicReportData.feedbackData.m_column.results, "AQ");
    
    // (E) ▼▼▼ [削除] グラフ・画像挿入ロジック (すべて削除) ▼▼▼
    // console.log(`[googleSlidesService] TODO: Generating chart/image insertion requests...`);
    // const imagePromises = imagePlaceholders.map(async (placeholderKey) => { ... });
    // await Promise.all(imagePromises);
    // ▲▲▲


    // (F) batchUpdate を実行
    if (requests.length > 0) {
        console.log(`[googleSlidesService] Sending ${requests.length} update requests (Text and Comments ONLY) to Slides API...`);
        
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
 * ヘルパー (1/X): GAS Web App 呼び出し
 * =================================================================
 */
async function callGasToCloneSlide(clinicName, centralSheetId) { 
    try {
        const response = await fetch(GAS_SLIDE_GENERATOR_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ 
                clinicName: clinicName,
                centralSheetId: centralSheetId 
            })
        });

        if (!response.ok) {
            throw new Error(`GAS Web App (Slide Clone) request failed with status ${response.status}: ${await response.text()}`);
        }
        const result = await response.json();
        
        if (result.status === 'ok' && result.newSlideId) {
            return {
                newSlideId: result.newSlideId,
                newSlideUrl: result.newSlideUrl,
                analysisData: result.analysisData // ★ GASからの分析データ (2D配列)
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
 * ヘルパー (2/X): テキスト挿入リクエスト作成 (複製なし)
 * =================================================================
 */
// ▼▼▼ [修正] municipalityText と reportCounts の引数を削除 ▼▼▼
function addTextRequests(requests, placeholderMap, period, clinicData, overallData, aiData) {
    
    // --- 期間テキスト (A, B) ---
    const [startY, startM] = period.start.split('-');
    const [endY, endM] = period.end.split('-');
    const periodFullText = `${startY}年${startM}月1日〜${endY}年${endM}月末日`;
    const periodShortText = `${startY}年${startM}月〜${endY}年${endM}月`;
    
    // A, B (deleteText: true)
    addTextUpdateRequest(requests, placeholderMap, 'A', periodShortText, true); 
    addTextUpdateRequest(requests, placeholderMap, 'B', periodFullText, true); 

    // --- 市区町村 (I) ---
    // ▼▼▼ [修正] 市区町村テキストの挿入を一時的にコメントアウト (データがないため) ▼▼▼
    // const municipalityTableText = "TODO: 市区町村の表テキスト (I)";
    // addTextUpdateRequest(requests, placeholderMap, 'I', municipalityTableText, true);
    // ▲▲▲

    // --- NPS値 (AA, AB) ---
    const clinicNpsScore = calculateNps(clinicData.npsScoreData.counts, clinicData.npsScoreData.totalCount);
    const overallNpsScore = calculateNps(overallData.npsScoreData.counts, overallData.npsScoreData.totalCount);
    addTextUpdateRequest(requests, placeholderMap, 'AA', clinicNpsScore.toFixed(1), false); // (deleteText: false)
    addTextUpdateRequest(requests, placeholderMap, 'AB', overallNpsScore.toFixed(1), false); // (deleteText: false)

    // --- NPSコメント人数 (BI, BJ, BK, BL, BM) ---
    const npsResults = clinicData.npsData.results || {};
    addTextUpdateRequest(requests, placeholderMap, 'BI', `${(npsResults[10] || []).length}人`, false);
    addTextUpdateRequest(requests, placeholderMap, 'BJ', `${(npsResults[9] || []).length}人`, false); 
    addTextUpdateRequest(requests, placeholderMap, 'BK', `${(npsResults[8] || []).length}人`, false);
    addTextUpdateRequest(requests, placeholderMap, 'BL', `${(npsResults[7] || []).length}人`, false);
    const nps6BelowCount = (npsResults[6] || []).length + (npsResults[5] || []).length + (npsResults[4] || []).length + (npsResults[3] || []).length + (npsResults[2] || []).length + (npsResults[1] || []).length + (npsResults[0] || []).length;
    addTextUpdateRequest(requests, placeholderMap, 'BM', `${nps6BelowCount}人`, false);
    
    // --- AI分析テキスト (AT 〜 BG) ---
    // (deleteText: true)
    addTextUpdateRequest(requests, placeholderMap, 'AT', aiData['L']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'AU', aiData['L']?.suggestions || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'AV', aiData['L']?.overall || '（データなし）', true);
    
    addTextUpdateRequest(requests, placeholderMap, 'AW', aiData['I_bad']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'AX', aiData['I_bad']?.suggestions || '（データなし）', true);

    addTextUpdateRequest(requests, placeholderMap, 'AY', aiData['I_good']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'AZ', aiData['I_good']?.suggestions || '（データなし）', true); 
    addTextUpdateRequest(requests, placeholderMap, 'BA', aiData['I_good']?.overall || '（データなし）', true);
    
    addTextUpdateRequest(requests, placeholderMap, 'BB', aiData['J']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'BC', aiData['J']?.suggestions || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'BD', aiData['J']?.overall || '（データなし）', true);

    addTextUpdateRequest(requests, placeholderMap, 'BE', aiData['M']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'BF', aiData['M']?.suggestions || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'BG', aiData['M']?.overall || '（データなし）', true);

    // --- 集計値 (BH) ---
    // ▼▼▼ [修正] 行数カウントの挿入を一時的にコメントアウト (データがないため) ▼▼▼
    // const { clinicCount, overallCount, clinicListCount } = reportCounts;
    // const bhText = `全体：${overallCount}件（${clinicListCount}病院） 貴院：${clinicCount}件`;
    // addTextUpdateRequest(requests, placeholderMap, 'BH', bhText, true);
    // ▲▲▲
}


/**
 * =================================================================
 * ヘルパー (3/X): コメントスライド複製リクエスト作成
 * =================================================================
 */
function addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, comments, templatePlaceholder) {
    // (templatePlaceholder は "AC", "AK" など)
    const templateSlideId = slideTemplateMap[templatePlaceholder];
    const templateElementId = placeholderMap[templatePlaceholder];

    if (!templateSlideId || !templateElementId) {
        console.warn(`[googleSlidesService] Template slide or element not found for placeholder ${templatePlaceholder}. Skipping duplication.`);
        addTextUpdateRequest(requests, placeholderMap, templatePlaceholder, "（該当コメントなし）", true); 
        return;
    }

    const commentChunks = chunkComments(comments);

    if (commentChunks.length === 0) {
        addTextUpdateRequest(requests, placeholderMap, templatePlaceholder, "（該当コメントなし）", true);
        return;
    }

    // 4. 1ページ目 (既存のテンプレートスライド) に1チャンク目を挿入
    console.log(`[googleSlidesService] Filling template slide ${templateSlideId} (placeholder ${templatePlaceholder}) with chunk 1...`);
    addTextUpdateRequest(requests, placeholderMap, templatePlaceholder, commentChunks[0], true); // deleteText: true

    // 5. 2ページ目以降のチャンクがあれば、スライドを複製して挿入
    for (let i = 1; i < commentChunks.length; i++) {
        const chunk = commentChunks[i];
        const newSlideId = `new_slide_${templatePlaceholder}_${i}`;
        const newElementId = `new_element_${templatePlaceholder}_${i}`;
        
        const idMap = {};
        idMap[templateSlideId] = newSlideId;
        idMap[templateElementId] = newElementId;
        
        console.log(`[googleSlidesService] Duplicating slide ${templateSlideId} -> ${newSlideId} (for ${templatePlaceholder} chunk ${i+1})`);

        requests.push({
            duplicateObject: {
                objectId: templateSlideId,
                objectIds: idMap
            }
        });
        
        requests.push({
            deleteText: {
                objectId: newElementId, 
                textRange: { type: 'ALL' }
            }
        });
        
        requests.push({
            insertText: {
                objectId: newElementId, 
                insertionIndex: 0,
                text: chunk
            }
        });
    }
}


/**
 * =================================================================
 * (変更) ヘルパー (4/X): その他ユーティリティ
 * =================================================================
 */

/**
 * [Util] Google Slides API へのテキスト挿入リクエストを生成
 */
function addTextUpdateRequest(requests, placeholderMap, placeholder, newText, doDeleteText = true) {
    // (placeholder は "A", "AT" など)
    const objectId = placeholderMap[placeholder];
    if (!objectId) {
        console.warn(`[googleSlidesService] Placeholder key "${placeholder}" for TEXT not found in placeholderMap (was it missing from the GAS analysis sheet E-column?). Skipping update.`);
        return;
    }
    
    if (doDeleteText) {
        requests.push({
            deleteText: {
                objectId: objectId,
                textRange: { type: 'ALL' }
            }
        });
    }

    requests.push({
        insertText: {
            objectId: objectId,
            insertionIndex: 0,
            text: newText || '（データなし）' // null や undefined を防ぐ
        }
    });
}

/**
 * [Util] NPSスコアを計算 (変更なし)
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
 * [Util] コメント配列をスライド挿入用の単一文字列にフォーマット (変更なし)
 */
function formatComments(commentsArray) {
    if (!commentsArray || commentsArray.length === 0) {
        return '（該当コメントなし）';
    }
    // スライドのテキストボックスは \n で改行できる
    return commentsArray.map(c => `・ ${String(c).trim()}`).join('\n');
}

/**
 * [Util/新規] コメント配列を、スライド1枚あたりの上限(MAX_LINES)で分割する
 */
function chunkComments(commentsArray) {
    const MAX_LINES_PER_SLIDE = 20; // (20行)
    
    if (!commentsArray || commentsArray.length === 0) {
        return [];
    }
    
    const chunks = [];
    let currentChunkLines = 0;
    let currentChunkText = '';
    
    commentsArray.forEach(comment => {
        const text = `・ ${String(comment).trim()}`;
        const linesInComment = (text.match(/\n/g) || []).length + 1;

        if (currentChunkLines > 0 && (currentChunkLines + linesInComment) > MAX_LINES_PER_SLIDE) {
            chunks.push(currentChunkText);
            currentChunkText = text;
            currentChunkLines = linesInComment;
        } else {
            currentChunkText += (currentChunkLines === 0 ? '' : '\n') + text;
            currentChunkLines += linesInComment;
        }
    });
    
    // 最後のチャンクを保存
    if (currentChunkText) {
        chunks.push(currentChunkText);
    }
    
    return chunks;
}
