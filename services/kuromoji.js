const kuromoji = require('kuromoji');
const path = require('path');
const { stopWords } = require('../config/stopwords'); // ストップワードを読み込む

let tokenizer;

// --- Tokenizer 初期化 ---
exports.initializeKuromoji = () => {
    return new Promise((resolve, reject) => {
        if (tokenizer) {
            console.log('Kuromoji tokenizer already initialized.');
            return resolve();
        }
        console.log('Initializing Kuromoji tokenizer...');
        kuromoji.builder({ dicPath: path.join(__dirname, '..', 'node_modules', 'kuromoji', 'dict') })
            .build((err, _tokenizer) => {
                if (err) {
                    console.error('Failed to build kuromoji tokenizer:', err);
                    reject(err);
                } else {
                    tokenizer = _tokenizer;
                    console.log('Kuromoji tokenizer built successfully.');
                    resolve();
                }
            });
    });
};

// --- テキストリスト分析 ---
exports.analyzeTextList = (textList) => {
    if (!tokenizer) {
        throw new Error('Kuromoji tokenizer is not initialized.');
    }

    console.log(`[kuromojiService] Analyzing ${textList.length} texts...`);
    const wordCounts = {};
    let processedTokensCount = 0;
    let validTextCount = 0;

    for (const text of textList) {
        if (!text || typeof text !== 'string' || text.trim() === '') continue;
        validTextCount++;
        const tokens = tokenizer.tokenize(text);
        processedTokensCount += tokens.length;

        tokens.forEach(token => {
            const pos = token.pos; const posDetail1 = token.pos_detail_1;
            const wordBase = token.basic_form !== '*' ? token.basic_form : token.surface_form;

            let targetPosCategory = null;
            if (pos === '名詞' && ['一般', '固有名詞', 'サ変接続', '形容動詞語幹'].includes(posDetail1)) { targetPosCategory = '名詞'; }
            else if (pos === '動詞' && (posDetail1 === '自立' || posDetail1 === 'サ変接続')) { targetPosCategory = '動詞'; }
            else if (pos === '形容詞' && posDetail1 === '自立') { targetPosCategory = '形容詞'; }
            else if (pos === '感動詞') { targetPosCategory = '感動詞'; }

            if (targetPosCategory && wordBase.length > 1 && !/^[0-9]+$/.test(wordBase) && !stopWords.has(wordBase)) {
                if (!wordCounts[wordBase]) {
                    wordCounts[wordBase] = { count: 0, pos: targetPosCategory };
                }
                wordCounts[wordBase].count++;
            }
        });
    }
    console.log(`[kuromojiService] Processed ${validTextCount} valid texts, ${processedTokensCount} tokens.`);

    const analysisResult = Object.entries(wordCounts).map(([word, data]) => ({
        word: word,
        score: data.count,
        pos: data.pos
    }));
    analysisResult.sort((a, b) => b.score - a.score);

    return {
        totalDocs: validTextCount,
        results: analysisResult
    };
};
