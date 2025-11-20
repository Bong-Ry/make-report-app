// bong-ry/make-report-app/make-report-app-2d48cdbeaa4329b4b6cca765878faab9eaea94af/utils/helpers.js

exports.getSpreadsheetIdFromUrl = (url) => {
    if (!url || typeof url !== 'string') return null;
    const match = url.match(/\/d\/(.+?)\//);
    return match ? match[1] : null;
};

// =================================================================
// === ▼▼▼ [変更] 分析タブ名 生成関数 ▼▼▼ ===
// =================================================================

/**
 * [新規] 保存先の分析タブ名（シート名）を取得する
 * @param {string} clinicName (例: "クリニックA")
 * @param {string} type (例: "AI", "MUNICIPALITY", "RECOMMENDATION")
 * @returns {string} (例: "クリニックA_AI分析")
 */
exports.getAnalysisSheetName = (clinicName, type) => {
    switch(type) {
        case 'AI':
            return `${clinicName}_AI分析`;
        case 'MUNICIPALITY':
            return `${clinicName}_市区町村`;
        case 'RECOMMENDATION':
            return `${clinicName}_おすすめ理由`;
        default:
            throw new Error(`[helpers] 無効な分析タイプです: ${type}`);
    }
};

/**
 * [新規] 「_AI分析」タブのA列に記載するキー（項目名）の一覧
 * (AI分析5種 * 3項目 = 15キー)
 */
const AI_ANALYSIS_KEYS = [
    // L (NPS)
    'L_ANALYSIS', 'L_SUGGESTIONS', 'L_OVERALL',
    // I_bad (悪い点)
    'I_bad_ANALYSIS', 'I_bad_SUGGESTIONS', 'I_bad_OVERALL',
    // I_good (良い点)
    'I_good_ANALYSIS', 'I_good_SUGGESTIONS', 'I_good_OVERALL',
    // J (スタッフ)
    'J_ANALYSIS', 'J_SUGGESTIONS', 'J_OVERALL',
    // M (お産意見)
    'M_ANALYSIS', 'M_SUGGESTIONS', 'M_OVERALL'
];

exports.getAiAnalysisKeys = () => {
    return AI_ANALYSIS_KEYS;
};

/**
 * [新規] AI分析の生JSONを指定のキーを持つMapに変換する
 * @param {object} analysisJson (OpenAIからの生JSON)
 * @param {string} columnType (例: "L", "I_good")
 * @returns {Map<string, string>} (例: Map { "L_ANALYSIS" => "分析文...", "L_SUGGESTIONS" => "改善案文..." })
 */
exports.formatAiJsonToMap = (analysisJson, columnType) => {
    const map = new Map();
    
    const analysisText = (analysisJson.analysis && analysisJson.analysis.themes) ? analysisJson.analysis.themes.map(t => `【${t.title}】\n${t.summary}`).join('\n\n---\n\n') : '（分析データがありません）';
    const suggestionsText = (analysisJson.suggestions && analysisJson.suggestions.items) ? analysisJson.suggestions.items.map(i => `【${i.themeTitle}】\n${i.suggestion}`).join('\n\n---\n\n') : '（改善提案データがありません）';
    const overallText = (analysisJson.overall && analysisJson.overall.summary) ? analysisJson.overall.summary : '（総評データがありません）';

    map.set(`${columnType}_ANALYSIS`, analysisText);
    map.set(`${columnType}_SUGGESTIONS`, suggestionsText);
    map.set(`${columnType}_OVERALL`, overallText);
    
    return map;
};


// =================================================================
// === ▼▼▼ [変更なし] おすすめ理由 表示名マッピング ▼▼▼ ===
// =================================================================
/**
 * [変更なし] おすすめ理由の「元カテゴリ名」から「表示名」へのマッピング
 * (ご要望: 口コミ、ホームページ など)
 */
exports.RECOMMENDATION_DISPLAY_NAMES = {
    // --- 表示名を変更する項目 ---
    'インターネット（情報サイト）': '情報サイト',
    'インターネット（Googleの口コミ）': '口コミ',
    'インターネット（産院のホームページ）': '公式サイト',

    // --- 表示名を変更しない固定項目 ---
    '知人の紹介': '知人の紹介',
    '家族の紹介': '家族の紹介',
    '自宅からの距離': '自宅からの距離',
    'インターネット（SNS）': 'インターネット（SNS）'
};


// =================================================================
// === ▼▼▼ [変更なし] OpenAI プロンプト ▼▼▼ ===
// =================================================================

// --- OpenAI (GPT) に投げるプロンプトの定義 (詳細分析用) ---
// (変更なし)
exports.getSystemPromptForDetailAnalysis = (clinicName, columnType) => {
    // 求めるJSONの厳格なフォーマット定義
    const jsonFormatString = `
# 出力フォーマット (厳格なJSONオブジェクトのみ):
{
  "analysis": {
    "title": "（ここに「分析と考察」のタイトル）",
    "themes": [
      {"title": "(抽出したテーマ1)", "summary": "(テーマ1の具体的な分析内容。100～150字程度)"},
      {"title": "(抽出したテーマ2)", "summary": "(テーマ2の具体的な分析内容。100～150字程度)"}
    ]
  },
  "suggestions": {
    "title": "（ここに「改善提案」のタイトル）",
    "items": [
      {"themeTitle": "(テーマ1に対応)", "suggestion": "(テーマ1への具体的な改善提案。100～150字程度)"},
      {"themeTitle": "(テーマ2に対応)", "suggestion": "(テーマ2への具体的な改善提案。100～150字程度)"}
    ]
  },
  "overall": {
    "title": "（ここに「総評」のタイトル）",
    "summary": "(全体を総括する詳細なレポート。約400～600文字程度で、改行(\\n)を含めて読みやすく記述してください)"
  }
}

# 指示:
- あなたは ${clinicName} の経営コンサルタントです。
- 以下の「タスク」に基づき、ユーザーから提供される「分析対象の生データ」を分析してください。
- 主要テーマ(themes)は3〜5個に絞り、具体的で分かりやすいタイトルを付けてください。
- 各summary, suggestion, overall.summaryは指定された文字数を目安に、プロフェッショナルなトーンで具体的に記述してください。
- ${clinicName} という病院名を適宜文章に含めてください。
- 出力は必ず指定されたJSONフォーマットに従い、それ以外のテキスト（例: "承知いたしました。以下が分析結果です。"など）は一切含めないでください。`;

    let contextTitle = "";
    let systemInstruction = "";

    // 分析タイプに応じてタスクを変更
    switch(columnType) {
        case 'L':
            contextTitle = "知人に病院を紹介したいと思う理由 (NPS推奨理由)";
            systemInstruction = `# タスク:\n「${contextTitle}」のアンケート回答を分析してください。\n1. **analysis**: "分析と考察" (title)。生データから主要な評価テーマ（例：スタッフ対応、院内環境）を3〜5個抽出し、それぞれ(title)と(summary)を記述。\n2. **suggestions**: "強みを伸ばす提案" (title)。(analysis)の各テーマに基づき、${clinicName}がさらに評価を高めるための具体的な提案(suggestion)を(themeTitle)ごとに記述。\n3. **overall**: "総評" (title)。(analysis)と(suggestions)を総括し、${clinicName}が取るべき戦略を詳細に記述(summary)。\n\n${jsonFormatString}`;
            break;
        case 'I_bad':
             contextTitle = "悪かった点";
            systemInstruction = `# タスク:\n「${contextTitle}」のアンケート回答を分析してください。\n1. **analysis**: "「悪かった点」の分析" (title)。生データから主要な不満のテーマ（例：待ち時間、情報の不足）を3〜5個抽出し、(title)と(summary)を記述。\n2. **suggestions**: "改善策の提案" (title)。(analysis)の各不満テーマに基づき、問題を解決する具体的な改善策(suggestion)を(themeTitle)ごとに記述。\n3. **overall**: "総評（緊急度と対策）" (title)。(analysis)と(suggestions)を総括し、${clinicName}が優先的に取り組むべき課題と対策を詳細に記述(summary)。\n\n${jsonFormatString}`;
            break;
        case 'I_good':
            contextTitle = "良かった点";
            systemInstruction = `# タスク:\n「${contextTitle}」のアンケート回答を分析してください。\n1. **analysis**: "「良かった点」の分析" (title)。生データから主要な評価テーマ（例：スタッフの優しさ、食事）を3〜5個抽出し、(title)と(summary)を記述。\n2. **suggestions**: "強みを伸ばす提案" (title)。(analysis)の各評価テーマに基づき、その強みを維持・強化する施策(suggestion)を(themeTitle)ごとに記述。\n3. **overall**: "総評（アピール戦略）" (title)。(analysis)と(suggestions)を総括し、${clinicName}が強みをどう活用すべきか詳細に記述(summary)。\n\n${jsonFormatString}`;
            break;
        case 'J':
             contextTitle = "印象に残ったスタッフへのコメント";
             systemInstruction = `# タスク:\n「${contextTitle}」のアンケート回答を分析してください。\n1. **analysis**: "コメント分析とスタッフ評価" (title)。生データから患者がスタッフを評価している主要ポイント（例：具体的な声かけ、技術）を3〜5個抽出し、(title)と(summary)を記述。\n2. **suggestions**: "接遇改善と評価への活用" (title)。(analysis)の評価ポイントに基づき、病院全体の接遇レベル向上策や、コメントをスタッフ評価に繋げる施策(suggestion)を(themeTitle)ごとに記述。\n3. **overall**: "総評（スタッフマネジメント）" (title)。(analysis)と(suggestions)を総括し、${clinicName}がスタッフの強みを活かすマネジメント戦略を詳細に記述(summary)。\n\n${jsonFormatString}`;
            break;
        case 'M':
             contextTitle = "お産にかかわるご意見・ご感想";
             systemInstruction = `# タスク:\n「${contextTitle}」のアンケート回答を分析してください。\n1. **analysis**: "ご意見の分析" (title)。生データから患者が重視しているテーマ（例：設備の快適さ、情報提供）を3〜5個抽出し、(title)と(summary)を記述。\n2. **suggestions**: "改善策の提案" (title)。(analysis)の各テーマに基づき、具体的な改善策（設備、サービス、情報提供方法など）(suggestion)を(themeTitle)ごとに記述。\n3. **overall**: "総評（サービス改善戦略）" (title)。(analysis)と(suggestions)を総括し、${clinicName}が「お産」体験の質を向上させるための優先施策を詳細に記述(summary)。\n\n${jsonFormatString}`;
            break;
        default:
             console.warn(`[helpers] Unknown columnType for prompt generation: ${columnType}`);
            return null; // 無効なタイプ
    }
    return systemInstruction;
};

// --- N列（おすすめ理由）分類プロンプト ---
// (変更なし)
exports.getSystemPromptForRecommendationAnalysis = (fixedKeys) => {
    
    // AIに認識させるカテゴリ名のリストを文字列として生成
    const categoriesListString = fixedKeys.map(key => `- "${key}"`).join('\n');

    // AIに出力してほしいJSONの形式を定義
    const jsonFormatString = `
# 出力フォーマット (厳格なJSONオブジェクトのみ):
{
  "classifiedResults": [
    {
      "originalText": "(分類対象の原文1)",
      "matchedCategories": ["(合致したカテゴリ名1)", "(合致したカテゴリ名2)"]
    },
    {
      "originalText": "(分類対象の原文2)",
      "matchedCategories": ["(合致したカテゴリ名1)"]
    },
    {
      "originalText": "(分類対象の原文3)",
      "matchedCategories": ["その他"]
    }
  ]
}
`;

    // AIへの指示本体
    const systemInstruction = `
# あなたの役割
あなたは、アンケートの自由回答を分析するデータ分類アシスタントです。

# タスク
ユーザーから「分類対象の生データ」として自由回答のリストが提供されます。
以下の「分類カテゴリ」を参照し、各回答がどのカテゴリに該当するかを判断してください。

# 分類カテゴリ
${categoriesListString}
- "その他"

# 指示
1.  **多重分類**: 1つの回答が複数のカテゴリに該当する内容を含んでいる場合（例：「家から近いし、口コミも良かった」）、該当する**すべて**のカテゴリ名を \`matchedCategories\` 配列に含めてください。
2.  **意図の汲み取り**: 表記が完全一致しなくても（例：「グーグルの評価」→「インターネット（Googleの口コミ）」）、意図が合致している場合はそのカテゴリとして分類してください。
3.  **「その他」の扱い**: どのカテゴリにも明確に合致しない回答のみ、\`matchedCategories\` 配列に \`"その他"\` という文字列を**一つだけ**入れてください。
4.  **出力**: 分析結果を、必ず指定されたJSONフォーマット（\`classifiedResults\` 配列）でのみ出力してください。JSON以外のテキスト（例: "承知いたしました。"など）は絶対に含めないでください。

${jsonFormatString}
`;

    return systemInstruction;
};

// ▼▼▼ [ここから変更] ▼▼▼
/**
 * [新規] AI分析の項目タイプとタブIDから、
 * 参照すべきセル名と五角形のタイトルを返す
 * @param {string} columnType (例: "L", "I_bad")
 * @param {string} tabId (例: "analysis", "suggestions", "overall")
 * @returns {{cell: string|null, pentagonText: string}}
 */
exports.getAiAnalysisCellAndTitle = (columnType, tabId) => {
    // 内部キー名 (例: "_ANALYSIS") へのマッピング
    const tabKeyMap = {
        'analysis': '_ANALYSIS',
        'suggestions': '_SUGGESTIONS',
        'overall': '_OVERALL'
    };
    
    // 五角形のテキスト (ご要望)
    const pentagonTitleMap = {
        'analysis': '分析と考察',
        'suggestions': '改善提案',
        'overall': '総評'
    };

    // 1. 探したいキー名を生成 (例: "L_SUGGESTIONS")
    const keyToFind = `${columnType}${tabKeyMap[tabId]}`;
    
    // 2. 五角形のテキストを取得
    const pentagonText = pentagonTitleMap[tabId] || '...';

    // 3. AI_ANALYSIS_KEYS (B2から始まるリスト) の中で、そのキーが何番目か検索
    // (getAiAnalysisKeys() はこのファイルの上部で定義されている前提)
    const index = AI_ANALYSIS_KEYS.indexOf(keyToFind);

    if (index === -1) {
        // 万が一キーが見つからない場合
        console.error(`[helpers] getAiAnalysisCell: Key not found: ${keyToFind}`);
        return { cell: null, pentagonText: 'エラー' };
    }

    // 4. セル名を計算
    // index 0 = B2
    // index 1 = B3
    // index 2 = B4
    const cellRow = index + 2; 
    const cell = `B${cellRow}`; // 例: "B3"
    
    return { cell, pentagonText };
};
// ▲▲▲ [変更ここまで] ▲▲▲
