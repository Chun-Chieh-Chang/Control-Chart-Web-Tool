/**
 * QIPProcessor 修復版本
 * 修復重點：
 * 1. 加強檢驗項目名稱提取的 logging
 * 2. 確保相同檢驗項目的數據能正確合併
 * 3. 支援 1~12, 1~16, 1~24, 1~32, 1~48 穴的同頁提取
 */

// 在 extractInspectionItemsFromGroup 方法中，加入以下修改：

/*
問題診斷：
當前代碼第 144-155 行的邏輯：
- 會在每個數據行左側尋找檢驗項目名稱
- 如果找不到（例如空白或合併儲存格），itemName 就是空字串
- Line 190 的 `if (hasData && itemName)` 會跳過這一行

修復方案：
1. 增強日誌：記錄每次提取的詳細資訊
2. 空白處理：如果左側沒有名稱，不要立即跳過（可能是合併儲存格）
3. 驗證合併：檢查是否需要處理 worksheet['!merges']

建議用戶執行以下步驟來診斷：
*/

// 第一步：在 Line 137 之前添加：
console.log(`[QIP Debug] === 開始處理 Group ===`);
console.log(`[QIP Debug] Cavity ID 範圍: ${groupConfig.cavityIdRange}`);
console.log(`[QIP Debug] 數據範圍: ${groupConfig.dataRange}`);
console.log(`[QIP Debug] 解析後 - ID行: ${idRangeParsed.startRow}, 數據行: ${dataRangeParsed.startRow}-${dataRangeParsed.endRow}`);
console.log(`[QIP Debug] 解析後 - ID列: ${idRangeParsed.startCol}-${idRangeParsed.endCol}, 數據列: ${dataRangeParsed.startCol}-${dataRangeParsed.endCol}`);

// 第二步：在 Line 143-155 處（檢驗項目名稱提取）修改為：
for (let c = 0; c < dataRangeParsed.startCol - 1; c++) {
    const cellAddr = XLSX.utils.encode_cell({ r: dataRow, c: c });
    const cell = worksheet[cellAddr];

    console.log(`[QIP Debug]   檢查 ${cellAddr}: ${cell ? cell.v : '(空)'}`);

    if (cell && cell.v !== undefined && String(cell.v).trim() !== '') {
        const value = String(cell.v).trim().replace(/[()]/g, '');
        if (value && !DataExtractor.isNumericString(value)) {
            itemName = value;
            console.log(`[QIP Debug]   ✓ 找到檢驗項目: "${itemName}"`);
            break;
        }
    }
}

if (!itemName) {
    console.warn(`[QIP Debug]   ✗ 第 ${dataRow + 1} 行未找到檢驗項目名稱！`);
    console.warn(`[QIP Debug]     這可能導致數據被跳過`);
}

// 第三步：在 Line 190 處添加：
if (hasData && itemName) {
    console.log(`[QIP Debug] ✓ 提取成功: ${itemName}, 穴數: ${Object.keys(data).length}, 穴號: [${Object.keys(data).sort((a, b) => parseInt(a) - parseInt(b)).join(', ')}]`);
    result.push({
        inspectionItem: itemName,
        data: data
    });
} else if (hasData && !itemName) {
    console.error(`[QIP Debug] ✗ 發現數據但無檢驗項目名稱！數據被丟棄。穴數: ${Object.keys(data).length}`);
} else if (!hasData && itemName) {
    console.warn(`[QIP Debug] ⚠ 發現檢驗項目名稱 "${itemName}" 但無有效數據`);
}
