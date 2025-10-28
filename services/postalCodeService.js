const postalCodeCache = new Map();

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
            // 市区町村(address2)とそれ以降(address3)を結合
            const result = { prefecture: address1, municipality: (address2 || '') + (address3 || '') };
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
