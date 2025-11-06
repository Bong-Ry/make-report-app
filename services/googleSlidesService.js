// bong-ry/make-report-app/make-report-app-c8b1854bb39f1d12bc0a586aaf23ca01a5592bcb/services/googleSlidesService.js
// (エラー修正版)

const googleSheetsService = require('./googleSheets');
const { slides: slidesApi, GAS_SLIDE_GENERATOR_URL } = require('./googleSheets');
const { getAnalysisSheetName } = require('../utils/helpers');

/**
 * =================================================================
 * メイン関数: スライド生成の全プロセスを実行
 * =================================================================
 */
exports.generateSlideReport = async (clinicName, centralSheetId, period, periodText) => {
    
    // --- 1. GASを呼び出し、スライド複製と要素マップ取得 ---
    console.log(`[googleSlidesService] Calling GAS to clone slide for: ${clinicName}`);
    
    // ★ 修正: analysisData (GASが生成したシート内容) を受け取る
    const { newSlideId, newSlideUrl, analysisData } = await callGasToCloneSlide(clinicName, centralSheetId);
    console.log(`[googleSlidesService] Slide cloned. New ID: ${newSlideId}`);

    if (!analysisData || analysisData.length === 0) {
        // ★ 修正: GAS側がグループ内を検索するようになったため、このエラーは「GASがC8などを見つけられなかった」ことを示す
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
    const aiDataMap = await googleSheetsService.readAiAnalysisData(centralSheetId, aiSheetName);
    
    console.log(`[googleSlidesService] Fetched AI Data Map from "${aiSheetName}".`);
    
    for (const type of aiTypes) {
         aiAnalysisResults[type] = {
            analysis: aiDataMap.get(`${type}_ANALYSIS`) || '（データなし）',
            suggestions: aiDataMap.get(`${type}_SUGGESTIONS`) || '（データなし）',
            overall: aiDataMap.get(`${type}_OVERALL`) || '（データなし）'
        };
    }
    
    // --- 4. スライドAPI (batchUpdate) で全データを挿入 ---
    
    // (A) ▼▼▼ [根本修正] GASから受け取った analysisData (シート配列) から ID を抽出 ▼▼▼
    
    console.log(`[googleSlidesService] Building placeholder map from GAS analysisData (Sheet) by *searching* placeholder text...`);
    
    const placeholderMap = {}; // "C8" -> "g123" (elementId)
    const slideTemplateMap = {}; // "C176" -> "g_slide_22" (slideId)

    /**
     * ★★★ [根本修正] ★★★
     * getIdsFromCellRef 関数を、インデックス参照(C143 -> 143行目)から、
     * analysisData 配列を 'find' (検索) するロジックに変更。
     * * GASが書き出す 'analysisData' (2D配列) の形式:
     * A(0): スライド番号, B(1): 要素番号, C(2): 要素ID, 
     * D(3): 種類, E(4): 内容(C8, C143など), F(5): スライドID
     */
    const getIdsFromCellRef = (ref) => {
        // 'ref' は "C8" や "C143" などのプレースホルダーテキスト
        
        // analysisData (全要素の配列) から、
        // E列(index 4) のテキスト内容が 'ref' と一致する行を探す
        const row = analysisData.find(r => r[4] === ref);

        if (!row) {
            // (GASがグループ内を検索しても "C143" などのテキストを見つけられなかった場合)
            console.warn(`[googleSlidesService] Placeholder text "${ref}" not found in analysisData sheet.`);
            return { elementId: null, slideId: null };
        }
        
        // 見つかった行から C列(index 2) と F列(index 5) のIDを取得
        const elementId = row[2] || null; // C列 (index 2) が elementId
        const slideId = row[5] || null; // F列 (index 5) が slideId
        
        // ★ ログ追加: どのテキスト(ref)がどのID(elementId)に対応したか確認
        // console.log(`[googleSlidesService] Found mapping: ${ref} -> elementId: ${elementId}, slideId: ${slideId}`);

        return { elementId, slideId };
    };

    // (プレースホルダーのリストは変更なし)
    const allPlaceholders = [
        "C8", "C17", "C124", "C125", "C135", "C143", "C152", "C164", "C165", 
        "C215", "C222", "C229", "C236", "C243", "C250", "C257", "C264", 
        "C271", "C278", "C286", "C293", "C300"
    ];
    const commentTemplatePlaceholders = ["C134", "C142", "C151", "C162", "C163", "C176", "C187", "C198"];
    
    // 1. 通常のテキスト置換用ID (C8, C17, C124...) を取得
    for (const ref of allPlaceholders) {
         const { elementId } = getIdsFromCellRef(ref); // ★ 修正後の検索関数をコール
         if (elementId) {
             placeholderMap[ref] = elementId;
         }
    }
    
    // 2. コメント複製用テンプレートのID (C134, C176...) を取得
    for (const ref of commentTemplatePlaceholders) {
         const { elementId, slideId } = getIdsFromCellRef(ref); // ★ 修正後の検索関数をコール
         if (elementId) {
             placeholderMap[ref] = elementId;
         }
         if (slideId) {
             slideTemplateMap[ref] = slideId;
         }
    }
    
    console.log(`[googleSlidesService] Placeholder Map (e.g., C8 -> elementId):\n`, placeholderMap);
    console.log(`[googleSlidesService] Slide Template Map (e.g., C176 -> slideId):\n`, slideTemplateMap);


    // (B) 挿入リクエストの配列を作成
    const requests = [];

    // (C) テキストデータを挿入 (C8, C17 など、複製が不要なもの)
    addTextRequests(requests, placeholderMap, period, clinicReportData, overallReportData, aiAnalysisResults);

    // (D) コメントスライドを複製・挿入
    console.log(`[googleSlidesService] Generating comment slide duplication requests...`);
    
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[10] || []), "C134");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[9] || []), "C142");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[8] || []), "C151");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[7] || []), "C162");
    const nps6BelowComments = [
        ...(clinicReportData.npsData.results[6] || []), ...(clinicReportData.npsData.results[5] || []), ...(clinicReportData.npsData.results[4] || []),
        ...(clinicReportData.npsData.results[3] || []), ...(clinicReportData.npsData.results[2] || []), ...(clinicReportData.npsData.results[1] || []), ...(clinicReportData.npsData.results[0] || [])
    ];
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, nps6BelowComments, "C163");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, clinicReportData.feedbackData.i_column.results, "C176");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, clinicReportData.feedbackData.j_column.results, "C187");
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, clinicReportData.feedbackData.m_column.results, "C198");
    
    // (E) グラフ・画像を挿入 (TODO)
    // TODO: addChartRequests(requests, placeholderMap, clinicReportData, overallReportData);

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
function addTextRequests(requests, placeholderMap, period, clinicData, overallData, aiData) {
    
    // --- 期間テキスト ---
    const [startY, startM] = period.start.split('-');
    const [endY, endM] = period.end.split('-');
    const periodFullText = `${startY}年${startM}月1日〜${endY}年${endM}月末日`;
    const periodShortText = `${startY}年${startM}月〜${endY}年${endM}月`;
    
    // (C8, C17 はプレースホルダーテキストが入っているので deleteText: true)
    addTextUpdateRequest(requests, placeholderMap, 'C8', periodShortText, true); 
    addTextUpdateRequest(requests, placeholderMap, 'C17', periodFullText, true); 

    // --- NPS値 ---
    const clinicNpsScore = calculateNps(clinicData.npsScoreData.counts, clinicData.npsScoreData.totalCount);
    const overallNpsScore = calculateNps(overallData.npsScoreData.counts, overallData.npsScoreData.totalCount);
    // (C124, C125 は空欄のテキストボックスを想定し deleteText: false)
    addTextUpdateRequest(requests, placeholderMap, 'C124', clinicNpsScore.toFixed(1), false); 
    addTextUpdateRequest(requests, placeholderMap, 'C125', overallNpsScore.toFixed(1), false); 

    // --- NPSコメント人数 (スライド 17, 18, 19, 20) ---
    const npsResults = clinicData.npsData.results || {};
    // (C135, C143 なども空欄のテキストボックスを想定し deleteText: false)
    addTextUpdateRequest(requests, placeholderMap, 'C135', `${(npsResults[10] || []).length}人`, false);
    addTextUpdateRequest(requests, placeholderMap, 'C143', `${(npsResults[9] || []).length}人`, false); 
    addTextUpdateRequest(requests, placeholderMap, 'C152', `${(npsResults[8] || []).length}人`, false);
    addTextUpdateRequest(requests, placeholderMap, 'C164', `${(npsResults[7] || []).length}人`, false);
    const nps6BelowCount = (npsResults[6] || []).length + (npsResults[5] || []).length + (npsResults[4] || []).length + (npsResults[3] || []).length + (npsResults[2] || []).length + (npsResults[1] || []).length + (npsResults[0] || []).length;
    addTextUpdateRequest(requests, placeholderMap, 'C165', `${nps6BelowCount}人`, false);
    
    // --- AI分析テキスト ---
    // (これらは {{...}} などのプレースホルダーテキストが入っているので deleteText: true)
    addTextUpdateRequest(requests, placeholderMap, 'C215', aiData['L']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'C222', aiData['L']?.suggestions || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'C229', aiData['L']?.overall || '（データなし）', true);
    
    addTextUpdateRequest(requests, placeholderMap, 'C236', aiData['I_bad']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'C243', aiData['I_bad']?.suggestions || '（データなし）', true);

    addTextUpdateRequest(requests, placeholderMap, 'C250', aiData['I_good']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'C257', aiData['I_good']?.suggestions || '（データなし）', true); 
    addTextUpdateRequest(requests, placeholderMap, 'C264', aiData['I_good']?.overall || '（データなし）', true);
    
    addTextUpdateRequest(requests, placeholderMap, 'C271', aiData['J']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'C278', aiData['J']?.suggestions || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'C286', aiData['J']?.overall || '（データなし）', true);

    addTextUpdateRequest(requests, placeholderMap, 'C293', aiData['M']?.analysis || '（データなし）', true);
    addTextUpdateRequest(requests, placeholderMap, 'C300', aiData['M']?.suggestions || '（データなし）', true);
}


/**
 * =================================================================
 * ヘルパー (3/X): コメントスライド複製リクエスト作成
 * =================================================================
 */
function addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, comments, templatePlaceholder) {
    const templateSlideId = slideTemplateMap[templatePlaceholder];
    const templateElementId = placeholderMap[templatePlaceholder];

    if (!templateSlideId || !templateElementId) {
        console.warn(`[googleSlidesService] Template slide or element not found for placeholder ${templatePlaceholder}. Skipping duplication.`);
        // (テンプレート自体が見つからない場合でも、プレースホルダーに「該当なし」と書き込む)
        // (コメントテンプレートはプレースホルダーテキスト(C134など)が入っている前提のため deleteText: true)
        addTextUpdateRequest(requests, placeholderMap, templatePlaceholder, "（該当コメントなし）", true); 
        return;
    }

    const commentChunks = chunkComments(comments);

    if (commentChunks.length === 0) {
        // (コメントテンプレートはプレースホルダーテキスト(C134など)が入っている前提のため deleteText: true)
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
 * (doDeleteText フラグは、呼び出し元で制御するよう変更)
 */
function addTextUpdateRequest(requests, placeholderMap, placeholder, newText, doDeleteText = true) {
    const objectId = placeholderMap[placeholder];
    if (!objectId) {
        // ★ 修正: この警告は、GASが "C143" などのプレースホルダーを見つけられなかったことを意味する
        console.warn(`[googleSlidesService] Placeholder "${placeholder}" not found in slide map (was it missing from the GAS analysis sheet?). Skipping update.`);
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
