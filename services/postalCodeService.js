const postalCodeCache = new Map();

/**
 * [新規] 市区町村名から「〜〜」以降を削除するヘルパー
 * (例: "横浜市港北区新横浜" -> "横浜市港北区")
 * (例: "大阪市" -> "大阪市")
 */
function truncateMunicipality(municipality) {
    if (!municipality) return '';
    // "市" または "区" が含まれる場合
    if (municipality.includes('市') || municipality.includes('区')) {
        // 正規表現で "市" または "区" までの部分を抽出
        // (例: "横浜市港北区" や "千葉市" にマッチ)
        const match = municipality.match(/^.*?(市|区)/);
        if (match) {
            return match[0]; // "横浜市港北区" や "千葉市" を返す
        }
    }
    // "郡" の場合 (例: "愛知郡東郷町") - 郡の後は「町」や「村」まで含める
    if (municipality.includes('郡')) {
         const match = municipality.match(/^.*?郡(.*?)(町|村)/);
         if (match) {
             return match[0]; // "愛知郡東郷町" を返す
         }
    }
    // どれにも当てはまらない場合はそのまま返す (例: "利尻富士町")
    return municipality;
}

/**
 * 外部API(zipcloud)から郵便番号の住所を非同期で取得
 * @param {string} postalCode 7桁の郵便番号 (例: "1000001")
 * @returns {Promise<object|null>} {prefecture, municipality} または null
 */
async function fetchFromApi(postalCode) {
    console.log(`[postalCodeService] Fetching from API: ${postalCode}`);
    try {
        // Node.js 18+ 標準の fetch を使用
        const response = await fetch(`https://zipcloud.ibsnet.co.jp/api/search?zipcode=${postalCode}`);
        if (!response.ok) {
            console.warn(`[postalCodeService] API response not OK for ${postalCode}. Status: ${response.status}`);
            return null;
        }
        
        const data = await response.json();
        
        if (data.status === 200 && data.results && data.results.length > 0) {
            const { address1, address2, address3 } = data.results[0];
            
            // ▼▼▼ 市区町村(address2)とそれ以降(address3)を結合 ▼▼▼
            const fullMunicipality = (address2 || '') + (address3 || '');
            
            // ▼▼▼ [変更点] 整形関数を呼び出す ▼▼▼
            const truncated = truncateMunicipality(fullMunicipality);

            const result = { prefecture: address1, municipality: truncated };
            postalCodeCache.set(postalCode, result); // 成功結果をキャッシュ
            return result;
        } else {
            console.warn(`[postalCodeService] API returned no results for ${postalCode}. Status: ${data.status}, Message: ${data.message}`);
            return null;
        }
    } catch (e) {
        console.error(`[postalCodeService] API call failed for ${postalCode}:`, e.message);
        return null;
    }
}

/**
 * 郵便番号から住所を取得 (キャッシュ優先)
 * @param {string} postalCode 7桁の郵便番号
 */
exports.lookupPostalCode = async (postalCode) => {
    if (postalCodeCache.has(postalCode)) {
        return postalCodeCache.get(postalCode);
    }
    return await fetchFromApi(postalCode);
};
