        
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
    
    // ▼▼▼ [修正] キーを "C8" 形式から "A", "B"... 形式に変更 ▼▼▼
    
    // (A, B はプレースホルダーテキストが入っているので deleteText: true)
    addTextUpdateRequest(requests, placeholderMap, 'A', periodShortText, true); // (旧: C8)
    addTextUpdateRequest(requests, placeholderMap, 'B', periodFullText, true); // (旧: C17)

    // --- NPS値 ---
    const clinicNpsScore = calculateNps(clinicData.npsScoreData.counts, clinicData.npsScoreData.totalCount);
    const overallNpsScore = calculateNps(overallData.npsScoreData.counts, overallData.npsScoreData.totalCount);
    // (C, D は空欄のテキストボックスを想定し deleteText: false)
    addTextUpdateRequest(requests, placeholderMap, 'C', clinicNpsScore.toFixed(1), false); // (旧: C124)
    addTextUpdateRequest(requests, placeholderMap, 'D', overallNpsScore.toFixed(1), false); // (旧: C125)

    // --- NPSコメント人数 (スライド 17, 18, 19, 20) ---
    const npsResults = clinicData.npsData.results || {};
    // (E, F なども空欄のテキストボックスを想定し deleteText: false)
    addTextUpdateRequest(requests, placeholderMap, 'E', `${(npsResults[10] || []).length}人`, false); // (旧: C135)
    addTextUpdateRequest(requests, placeholderMap, 'F', `${(npsResults[9] || []).length}人`, false); // (旧: C143)
    addTextUpdateRequest(requests, placeholderMap, 'G', `${(npsResults[8] || []).length}人`, false); // (旧: C152)
    addTextUpdateRequest(requests, placeholderMap, 'H', `${(npsResults[7] || []).length}人`, false); // (旧: C164)
    const nps6BelowCount = (npsResults[6] || []).length + (npsResults[5] || []).length + (npsResults[4] || []).length + (npsResults[3] || []).length + (npsResults[2] || []).length + (npsResults[1] || []).length + (npsResults[0] || []).length;
    addTextUpdateRequest(requests, placeholderMap, 'I', `${nps6BelowCount}人`, false); // (旧: C165)
    
    // --- AI分析テキスト ---
    // (これらは {{...}} などのプレースホルダーテキストが入っているので deleteText: true)
    addTextUpdateRequest(requests, placeholderMap, 'J', aiData['L']?.analysis || '（データなし）', true); // (旧: C215)
    addTextUpdateRequest(requests, placeholderMap, 'K', aiData['L']?.suggestions || '（データなし）', true); // (旧: C222)
    addTextUpdateRequest(requests, placeholderMap, 'L', aiData['L']?.overall || '（データなし）', true); // (旧: C229)
    
    addTextUpdateRequest(requests, placeholderMap, 'M', aiData['I_bad']?.analysis || '（データなし）', true); // (旧: C236)
    addTextUpdateRequest(requests, placeholderMap, 'N', aiData['I_bad']?.suggestions || '（データなし）', true); // (旧: C243)

    addTextUpdateRequest(requests, placeholderMap, 'O', aiData['I_good']?.analysis || '（データなし）', true); // (旧: C250)
    addTextUpdateRequest(requests, placeholderMap, 'P', aiData['I_good']?.suggestions || '（データなし）', true); // (旧: C257)
    addTextUpdateRequest(requests, placeholderMap, 'Q', aiData['I_good']?.overall || '（データなし）', true); // (旧: C264)
    
    addTextUpdateRequest(requests, placeholderMap, 'R', aiData['J']?.analysis || '（データなし）', true); // (旧: C271)
    addTextUpdateRequest(requests, placeholderMap, 'S', aiData['J']?.suggestions || '（データなし）', true); // (旧: C278)
    addTextUpdateRequest(requests, placeholderMap, 'T', aiData['J']?.overall || '（データなし）', true); // (旧: C286)

    addTextUpdateRequest(requests, placeholderMap, 'U', aiData['M']?.analysis || '（データなし）', true); // (旧: C293)
    addTextUpdateRequest(requests, placeholderMap, 'V', aiData['M']?.suggestions || '（データなし）', true); // (旧: C300)
    
    // ▲▲▲ [修正完了] ▲▲▲
}


/**
 * =================================================================
 * ヘルパー (3/X): コメントスライド複製リクエスト作成
 * =================================================================
 */
function addCommentSlidesRequests(requests, placeholderMap, slideTemplateMap, comments, templatePlaceholder) {
    // (templatePlaceholder は "W", "X" など)
    const templateSlideId = slideTemplateMap[templatePlaceholder];
    const templateElementId = placeholderMap[templatePlaceholder];

    if (!templateSlideId || !templateElementId) {
        console.warn(`[googleSlidesService] Template slide or element not found for placeholder ${templatePlaceholder}. Skipping duplication.`);
        // (テンプレート自体が見つからない場合でも、プレースホルダーに「該当なし」と書き込む)
        // (コメントテンプレートはプレースホルダーテキスト("W"など)が入っている前提のため deleteText: true)
        addTextUpdateRequest(requests, placeholderMap, templatePlaceholder, "（該当コメントなし）", true); 
        return;
    }

    const commentChunks = chunkComments(comments);

    if (commentChunks.length === 0) {
        // (コメントテンプレートはプレースホルダーテキスト("W"など)が入っている前提のため deleteText: true)
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
    // (placeholder は "A", "W" など)
    const objectId = placeholderMap[placeholder];
    if (!objectId) {
        // ★ 修正: この警告は、GASが "A" などのプレースホルダーを見つけられなかったことを意味する
        console.warn(`[googleSlidesService] Placeholder key "${placeholder}" not found in placeholderMap (was it missing from the GAS analysis sheet E-column?). Skipping update.`);
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
