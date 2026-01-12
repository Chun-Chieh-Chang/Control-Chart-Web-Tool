// 核心修復：處理合併儲存格的邏輯

// 替換 extractInspectionItemsFromGroup 方法（Line 126-211）

extractInspectionItemsFromGroup(worksheet, groupConfig) {
    const result = [];

    try {
        const idRangeParsed = DataValidator.parseRangeString(groupConfig.cavityIdRange);
        const dataRangeParsed = DataValidator.parseRangeString(groupConfig.dataRange);

        if (!idRangeParsed || !dataRangeParsed) {
            return result;
        }

        console.log(`[QIP] 提取穴組數據 - Cavity ID範圍: ${groupConfig.cavityIdRange}, 數據範圍: ${groupConfig.dataRange}`);

        // 獲取合併儲存格資訊
        const merges = worksheet['!merges'] || [];

        // 輔助函數：從可能的合併儲存格中獲取值
        const getValueFromMergedCell = (row, col) => {
            let cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
            let cell = worksheet[cellAddr];
            let value = cell ? cell.v : null;

            // 如果沒有值，檢查是否是合併儲存格的一部分
            if (value === undefined || value === null || value === '') {
                const merge = merges.find(m =>
                    row >= m.s.r && row <= m.e.r &&
                    col >= m.s.c && col <= m.e.c
                );
                if (merge) {
                    // 從合併區域的左上角讀取
                    cellAddr = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
                    cell = worksheet[cellAddr];
                    value = cell ? cell.v : null;
                }
            }

            return value;
        };

        // 遍歷數據範圍的每一行
        for (let rowOffset = 0; rowOffset <= dataRangeParsed.endRow - dataRangeParsed.startRow; rowOffset++) {
            const dataRow = dataRangeParsed.startRow - 1 + rowOffset; // 0-indexed
            const data = {};
            let itemName = '';

            // 提取檢驗項目名稱 - 模仿 VBA 邏輯
            // VBA: 先嘗試 A 列（col 0），再嘗試 B 列（col 1）
            let tempValue = getValueFromMergedCell(dataRow, 0);
            if (!tempValue || String(tempValue).trim() === '') {
                tempValue = getValueFromMergedCell(dataRow, 1);
            }

            if (tempValue) {
                // 移除括號並清理（VBA: Replace(Replace(tempValue, "(", ""), ")", "")）
                tempValue = String(tempValue).trim().replace(/[()]/g, '');
                // 只接受非數字的文字作為檢驗項目名稱
                if (tempValue && !DataExtractor.isNumericString(tempValue)) {
                    itemName = tempValue;
                    console.log(`[QIP][Row${dataRow + 1}] ✓ 找到檢驗項目: "${itemName}"`);
                }
            }

            // 如果沒找到項目名稱，跳過此行
            if (!itemName) {
                console.warn(`[QIP][Row${dataRow + 1}] ✗ 未找到有效的檢驗項目名稱，跳過此行`);
                continue;
            }

            // 提取穴號數據
            const idRow = idRangeParsed.startRow - 1; // 0-indexed
            let hasData = false;

            for (let colOffset = 0; colOffset <= idRangeParsed.endCol - idRangeParsed.startCol; colOffset++) {
                const col = idRangeParsed.startCol - 1 + colOffset;

                // 獲取穴號 ID（處理合併儲存格）
                let cavityId = getValueFromMergedCell(idRow, col);
                if (!cavityId) continue;

                cavityId = String(cavityId).trim();

                // 提取穴號數字（支援 "1號穴"、"CAV1"、"1" 等格式）
                const numMatch = cavityId.match(/\d+/);
                if (numMatch) {
                    cavityId = numMatch[0];
                }

                // 獲取數據（處理合併儲存格）
                const cellValue = getValueFromMergedCell(dataRow, col);

                if (cellValue !== undefined && cellValue !== null) {
                    const cleanValue = DataExtractor.cleanCellValue(cellValue);
                    if (cleanValue !== '' && !isNaN(parseFloat(cleanValue))) {
                        data[cavityId] = parseFloat(cleanValue);
                        hasData = true;
                    }
                }
            }

            // 記錄結果
            if (hasData && itemName) {
                const cavityIds = Object.keys(data).sort((a, b) => parseInt(a) - parseInt(b));
                result.push({
                    inspectionItem: itemName,
                    data: data
                });
                console.log(`[QIP][Row${dataRow + 1}] ✓ 提取成功: "${itemName}", 穴號: [${cavityIds.join(', ')}], 數量: ${cavityIds.length}`);
            } else if (hasData && !itemName) {
                console.error(`[QIP][Row${dataRow + 1}] ✗ 有數據但無檢驗項目名稱！數據被丟棄，穴數: ${Object.keys(data).length}`);
            } else if (!hasData && itemName) {
                console.warn(`[QIP][Row${dataRow + 1}] ⚠ 檢驗項目 "${itemName}" 無有效數據`);
            }
        }

    } catch (error) {
        console.error('[QIP] extractInspectionItemsFromGroup error:', error);
        this.errorLogger.logError('數據提取', error.message);
    }

    return result;
}
