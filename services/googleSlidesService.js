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
    
    const { newSlideId, newSlideUrl, analysisData } = await callGasToCloneSlide(clinicName, centralSheetId);
    console.log(`[googleSlidesService] Slide cloned. New ID: ${newSlideId}`);

    if (!analysisData || analysisData.length === 0) {
        throw new Error('GAS analysis data (analysisData) is empty or undefined.');
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
    
    // (A) ▼▼▼ [修正] GASから受け取った analysisData (シート配列) からセル参照でIDを抽出 ▼▼▼
    
    console.log(`[googleSlidesService] Building placeholder map from GAS analysisData (Sheet) by cell reference...`);
    
    const placeholderMap = {}; // "C8" -> "g123" (elementId)
    const slideTemplateMap = {}; // "C176" -> "g_slide_22" (slideId)

    /**
     * [修正ヘルパー] "C8" などのセル参照文字列を 
     * analysisData (0-based配列) の [行index] に変換し、
     * C列 (elemId) と F列 (slideId) の値を取得する。
     */
    const getIdsFromCellRef = (ref) => {
        const match = ref.match(/^C(\d+)$/);
        if (!match) {
            console.warn(`[googleSlidesService] Invalid cell reference format: ${ref}. Only 'C' column supported.`);
            return { elementId: null, slideId: null };
        }
        
        const rowIndex = parseInt(match[1], 10) - 1; // "C8" -> rowIndex 7
        
        if (rowIndex < 0 || rowIndex >= analysisData.length) {
            console.warn(`[googleSlidesService] Cell reference ${ref} (index ${rowIndex}) is out of bounds for analysisData (length ${analysisData.length})`);
            return { elementId: null, slideId: null };
        }
        
        const row = analysisData[rowIndex];
        if (!row) {
             console.warn(`[googleSlidesService] Row data not found for ${ref} (index ${rowIndex}).`);
             return { elementId: null, slideId: null };
        }
        
        // C列 (index 2) が elementId
        // ▼▼▼ [修正] F列 (index 5) が slideId ▼▼▼
        const elementId = row[2] || null;
        const slideId = row[5] || null; // GASコードでF列(index 5)に出力
        
        return { elementId, slideId };
    };

    // ▼▼▼ [修正] "C307" をリストから削除 (analysisData 範囲外のため) ▼▼▼
    const allPlaceholders = [
        "C8", "C17", "C124", "C125", "C135", "C143", "C152", "C164", "C165", 
        "C215", "C222", "C229", "C236", "C243", "C250", "C257", "C264", 
        "C271", "C278", "C286", "C293", "C300" // "C307" を削除
    ];
    // (コメント複製用テンプレートの「セル位置」)
    const commentTemplatePlaceholders = ["C134", "C142", "C151", "C162", "C163", "C176", "C187", "C198"];
    
    // 1. 通常のテキスト置換用ID (C8, C17, C124...) を取得
    for (const ref of allPlaceholders) {
         const { elementId } = getIdsFromCellRef(ref);
         if (elementId) {
             placeholderMap[ref] = elementId;
         }
    }
    
    // 2. コメント複製用テンプレートのID (C134, C176...) を取得
    for (const ref of commentTemplatePlaceholders) {
         const { elementId, slideId } = getIdsFromCellRef(ref);
         if (elementId) {
             placeholderMap[ref] = elementId;
         }
         // ▼▼▼ [修正] GASのF列から slideId を正しく取得 ▼▼▼
         if (slideId) {
             slideTemplateMap[ref] = slideId;
         }
         // (フォールバックロジックは削除)
    }
    
    console.log(`[googleSlidesService] Placeholder Map (e.g., C8 -> elementId):\n`, placeholderMap);
    console.log(`[googleSlidesService] Slide Template Map (e.g., C176 -> slideId):\n`, slideTemplateMap);


    // (B) 挿入リクエストの配列を作成
    const requests = [];

    // (C) テキストデータを挿入 (C8, C17 など、複製が不要なもの)
    addTextRequests(requests, placeholderMap, period, clinicReportData, overallReportData, aiAnalysisResults);

    // (D) コメントスライドを複製・挿入
    console.log(`[googleSlidesService] Generating comment slide duplication requests...`);
    
    // NPS 10 (C134)
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[10] || []), "C134");
    // NPS 9 (C142)
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[9] || []), "C142");
    // NPS 8 (C151)
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[8] || []), "C151");
    // NPS 7 (C162)
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, (clinicReportData.npsData.results[7] || []), "C162");
    // NPS 6以下 (C163)
    const nps6BelowComments = [
        ...(clinicReportData.npsData.results[6] || []), ...(clinicReportData.npsData.results[5] || []), ...(clinicReportData.npsData.results[4] || []),
        ...(clinicReportData.npsData.results[3] || []), ...(clinicReportData.npsData.results[2] || []), ...(clinicReportData.npsData.results[1] || []), ...(clinicReportData.npsData.results[0] || [])
    ];
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, nps6BelowComments, "C163");

    // 良かった点/悪かった点 (C176)
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, clinicReportData.feedbackData.i_column.results, "C176");
    // スタッフ (C187)
    addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, clinicReportData.feedbackData.j_column.results, "C187");
    // お産 (C198)
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
    
    addTextUpdateRequest(requests, placeholderMap, 'C8', periodShortText, false); // "false" = deleteText しない
    addTextUpdateRequest(requests, placeholderMap, 'C17', periodFullText, false); // "false" = deleteText しない

    // --- NPS値 ---
    const clinicNpsScore = calculateNps(clinicData.npsScoreData.counts, clinicData.npsScoreData.totalCount);
    const overallNpsScore = calculateNps(overallData.npsScoreData.counts, overallData.npsScoreData.totalCount);
    // ▼▼▼ [エラー修正] "true" -> "false" に変更 ▼▼▼
    addTextUpdateRequest(requests, placeholderMap, 'C124', clinicNpsScore.toFixed(1), false); // "false" = deleteText しない
    addTextUpdateRequest(requests, placeholderMap, 'C125', overallNpsScore.toFixed(1), false); // "false" = deleteText しない

    // --- NPSコメント人数 (スライド 17, 18, 19, 20) ---
    const npsResults = clinicData.npsData.results || {};
    // ▼▼▼ [エラー修正] "true" -> "false" に変更 (C143を含む、関連する数値すべて) ▼▼▼
    addTextUpdateRequest(requests, placeholderMap, 'C135', `${(npsResults[10] || []).length}人`, false); // "false" = deleteText しない
    addTextUpdateRequest(requests, placeholderMap, 'C143', `${(npsResults[9] || []).length}人`, false); // "false" = deleteText しない
    addTextUpdateRequest(requests, placeholderMap, 'C152', `${(npsResults[8] || []).length}人`, false); // "false" = deleteText しない
    addTextUpdateRequest(requests, placeholderMap, 'C164', `${(npsResults[7] || []).length}人`, false); // "false" = deleteText しない
    const nps6BelowCount = (npsResults[6] || []).length + (npsResults[5] || []).length + (npsResults[4] || []).length + (npsResults[3] || []).length + (npsResults[2] || []).length + (npsResults[1] || []).length + (npsResults[0] || []).length;
    addTextUpdateRequest(requests, placeholderMap, 'C165', `${nps6BelowCount}人`, false); // "false" = deleteText しない
    
    // --- AI分析テキスト ---
    // (これらはテンプレートに {{...}} が入っているので "true" のまま)
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
    
    // (C307 は範囲外エラーのため、addTextUpdateRequest を呼ばない)
}


/**
 * =================================================================
 * ヘルパー (3/X): コメントスライド複製リクエスト作成
 * =================================================================
 */
function addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, comments, templatePlaceholder) {
    // 1. テンプレートスライドID (例: "g_slide_22") を取得
    const templateSlideId = slideTemplateMap[templatePlaceholder];
    // 2. テンプレートスライド上のテキストボックスID (例: "g_element_176") を取得
    const templateElementId = placeholderMap[templatePlaceholder];

    // ▼▼▼ [修正] F列から slideId を取得できたため、このチェックが動作するはず ▼▼▼
    if (!templateSlideId || !templateElementId) {
        console.warn(`[googleSlidesService] Template slide or element not found for placeholder ${templatePlaceholder}. Skipping duplication.`);
        // (テンプレート自体が見つからない場合でも、プレースホルダーに「該当なし」と書き込む)
        // (deleteText しない)
        addTextUpdateRequest(requests, placeholderMap, templatePlaceholder, "（該当コメントなし）", false); 
        return;
    }

    // 3. コメントをチャンク (スライド1枚分の文字列) に分割
    const commentChunks = chunkComments(comments);

    if (commentChunks.length === 0) {
        // 該当コメントが0件の場合
        // (deleteText しない)
        addTextUpdateRequest(requests, placeholderMap, templatePlaceholder, "（該当コメントなし）", false);
        return;
    }

    // 4. 1ページ目 (既存のテンプレートスライド) に1チャンク目を挿入
    console.log(`[googleSlidesService] Filling template slide ${templateSlideId} (placeholder ${templatePlaceholder}) with chunk 1...`);
    // (コメントテンプレートは "C134" などのプレースホルダーテキストが入っているので deleteText が必要)
    addTextUpdateRequest(requests, placeholderMap, templatePlaceholder, commentChunks[0], true);

    // 5. 2ページ目以降のチャンクがあれば、スライドを複製して挿入
    for (let i = 1; i < commentChunks.length; i++) {
        const chunk = commentChunks[i];
        const newSlideId = `new_slide_${templatePlaceholder}_${i}`;
        const newElementId = `new_element_${templatePlaceholder}_${i}`;
        
        const idMap = {};
        idMap[templateSlideId] = newSlideId;      // (スライドIDのマッピング)
        idMap[templateElementId] = newElementId;  // (テキストボックスIDのマッピング)
        
        console.log(`[googleSlidesService] Duplicating slide ${templateSlideId} -> ${newSlideId} (for ${templatePlaceholder} chunk ${i+1})`);

        // Request 1: スライド自体を複製
        requests.push({
            duplicateObject: {
                objectId: templateSlideId,
                objectIds: idMap
            }
        });
        
        // Request 2: 複製したスライドのテキストボックスから、「1チャンク目」のテキストを削除
        requests.push({
            deleteText: {
                objectId: newElementId, 
                textRange: { type: 'ALL' }
            }
        });
        
        // Request 3: 複製したスライドのテキストボックスに、次のコメント群を挿入
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
 * (★ 修正: GaxiosError: Invalid requests[...].deleteText のエラー回避)
 *
 * @param {string} placeholder - "C8", "C124" など
 * @param {boolean} doDeleteText - true の場合のみ deleteText を発行する
 */
function addTextUpdateRequest(requests, placeholderMap, placeholder, newText, doDeleteText = true) {
    const objectId = placeholderMap[placeholder];
    if (!objectId) {
        console.warn(`[googleSlidesService] Placeholder "${placeholder}" not found in slide map. Skipping update.`);
        return;
    }
    
    // ▼▼▼ [修正] doDeleteText が true の場合のみ、deleteText リクエストを発行 ▼▼▼
    // (テンプレートで空欄が想定される C143 などは false で呼び出す)
    if (doDeleteText) {
        requests.push({
            deleteText: {
                objectId: objectId,
                textRange: { type: 'ALL' }
            }
        });
    }

    // 新しいテキストを挿入するリクエスト
    requests.push({
        insertText: {
            objectId: objectId,
            insertionIndex: 0,
            text: newText || '（データなし）' // null や undefined を防ぐ
        }
    });
}

// ▼▼▼ [削除] addTextUpdateSimple は addTextUpdateRequest(..., false) に統合されたため削除
// function addTextUpdateSimple(requests, placeholderMap, placeholder, newText) { ... }


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
