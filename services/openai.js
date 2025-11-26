const OpenAI = require('openai');

const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
let openai;

// --- 初期化 ---
if (OPENAI_API_KEY) {
    try {
        openai = new OpenAI({ apiKey: OPENAI_API_KEY });
        console.log('OpenAI client initialized successfully in openai.js.');
    } catch (err) {
        console.error('Failed to initialize OpenAI client in openai.js:', err);
    }
} else {
    console.warn('OPENAI_API_KEY is not set. Detailed analysis will not work.');
}

// --- JSON形式で分析結果を生成 ---
exports.generateJsonAnalysis = async (systemPrompt, inputText) => {
    if (!openai) {
        throw new Error('OpenAI client is not initialized. Check OPENAI_API_KEY.');
    }

    console.log("[openaiService] Calling OpenAI chat completions API...");
    try {
        const completion = await openai.chat.completions.create({
            model: "gpt-4o-mini",
            messages: [
                { role: "system", content: systemPrompt },
                { role: "user", content: `# 分析対象の生データ:\n${inputText}` }
            ],
            response_format: { type: "json_object" },
            temperature: 0.3,
        });

        console.log("[openaiService] OpenAI API response received.");
        let content = completion.choices[0]?.message?.content;

        if (!content) {
            console.error("[openaiService] OpenAI response content is empty.");
            throw new Error('AIからの応答が空でした。');
        }

        // ▼▼▼ AI応答のクリーニング処理 ▼▼▼
        // 1. Markdown記法 (```json ... ```) が含まれていた場合に除去
        content = content.replace(/```json/g, '').replace(/```/g, '');

        // 2. 最初の中括弧 { から 最後の中括弧 } までを切り出す（前後の余計な空白や文章を削除）
        const firstBrace = content.indexOf('{');
        const lastBrace = content.lastIndexOf('}');
        if (firstBrace !== -1 && lastBrace !== -1) {
            content = content.substring(firstBrace, lastBrace + 1);
        }
        // ▲▲▲ ここまで ▲▲▲

        try {
            const analysisJson = JSON.parse(content);
            return analysisJson;
        } catch (parseError) {
            console.error('[openaiService] Failed to parse AI JSON response:', parseError);
            console.error('[openaiService] AI Raw Text Response:', content);
            throw new Error(`AIが予期しない形式(JSON以外)で応答しました。`);
        }
    } catch (error) {
        console.error('[openaiService] Error calling OpenAI API:', error);
        if (error.response) {
            console.error('[openaiService] OpenAI Error Status:', error.response.status);
            console.error('[openaiService] OpenAI Error Data:', error.response.data);
        }
        // エラーをそのままスローしてコントローラーで処理
        throw new Error(`OpenAI API呼び出しエラー: ${error.message}`);
    }
};
