/**
 * SpecificationExtractor - 規格數據提取模組
 * 對應 VBA SpecificationExtractor.bas
 */
class SpecificationExtractor {

    /**
     * 規格數據結構
     */
    static createSpecificationData() {
        return {
            symbol: '',
            nominalValue: 0,
            upperTolerance: 0,
            lowerTolerance: 0,
            usl: 0,
            lsl: 0,
            target: 0,
            isValid: false
        };
    }

    /**
     * 尋找包含規格數據的工作表
     * @param {Object} workbook - SheetJS workbook
     * @returns {Object|null} worksheet
     */
    static findSpecificationWorksheet(workbook) {
        try {
            const specKeywords = ['規格', 'spec', 'specification', '檢驗標準', '檢驗規格'];

            // 優先尋找包含規格關鍵字的工作表
            for (const sheetName of workbook.SheetNames) {
                const lowerName = sheetName.toLowerCase();
                for (const keyword of specKeywords) {
                    if (lowerName.includes(keyword.toLowerCase())) {
                        return workbook.Sheets[sheetName];
                    }
                }
            }

            // 尋找包含規格數據的工作表
            for (const sheetName of workbook.SheetNames) {
                const ws = workbook.Sheets[sheetName];
                if (this.hasSpecificationData(ws)) {
                    return ws;
                }
            }

            // 回退到第一個工作表
            if (workbook.SheetNames.length > 0) {
                return workbook.Sheets[workbook.SheetNames[0]];
            }

            return null;
        } catch (error) {
            console.error('SpecificationExtractor.findSpecificationWorksheet error:', error);
            return null;
        }
    }

    /**
     * 檢查工作表是否包含規格數據
     * @param {Object} worksheet 
     * @returns {boolean}
     */
    static hasSpecificationData(worksheet) {
        try {
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const maxRow = Math.min(range.e.r, 99); // 只檢查前100行

            for (let r = 0; r <= maxRow; r++) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: 0 });
                const cell = worksheet[cellAddr];

                if (cell && cell.v !== undefined) {
                    const value = String(cell.v).trim().toLowerCase();
                    if (value.includes('檢驗項目') || value.includes('inspection') ||
                        value.includes('規格') || value.includes('spec')) {
                        return true;
                    }
                }
            }

            return false;
        } catch (error) {
            console.error('SpecificationExtractor.hasSpecificationData error:', error);
            return false;
        }
    }

    /**
     * 根據檢驗項目查找規格數據
     * @param {Object} worksheet 
     * @param {string} itemName - 檢驗項目名稱
     * @returns {Object} SpecificationData
     */
    static findSpecificationByItem(worksheet, itemName) {
        try {
            const spec = this.createSpecificationData();

            if (!itemName || !worksheet) return spec;

            // 清理檢驗項目名稱
            const cleanItemName = itemName.replace(/[()]/g, '').trim();
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            const maxRow = Math.min(range.e.r, 99);

            for (let r = 0; r <= maxRow; r++) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: 0 });
                const cell = worksheet[cellAddr];

                if (cell && cell.v !== undefined) {
                    const cellValue = String(cell.v).trim().replace(/[()]/g, '');

                    // 檢查是否匹配
                    if (cellValue === cleanItemName ||
                        cellValue.includes(cleanItemName) ||
                        cleanItemName.includes(cellValue)) {

                        const rowSpec = this.readSpecificationFromRow(worksheet, r);
                        if (rowSpec.isValid) {
                            return rowSpec;
                        }
                    }
                }
            }

            return spec;
        } catch (error) {
            console.error('SpecificationExtractor.findSpecificationByItem error:', error);
            return this.createSpecificationData();
        }
    }

    /**
     * 安全地獲取合併儲存格的值
     * @param {Object} worksheet 
     * @param {number} row 
     * @param {number} col 
     * @returns {string|number|null}
     */
    static getMergedCellValue(worksheet, row, col) {
        try {
            const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellAddr];
            let value = cell && cell.v !== undefined ? cell.v : null;

            // 如果沒有值，檢查是否是合併儲存格的一部分
            if (value === null || value === undefined || value === '') {
                const merges = worksheet['!merges'] || [];
                const merge = merges.find(m =>
                    m && m.s && m.e &&
                    row >= m.s.r && row <= m.e.r &&
                    col >= m.s.c && col <= m.e.c
                );
                if (merge) {
                    const mergedAddr = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
                    const mergedCell = worksheet[mergedAddr];
                    value = mergedCell && mergedCell.v !== undefined ? mergedCell.v : null;
                }
            }
            return value;
        } catch (err) {
            console.error(`[SpecExtractor] 讀取儲存格 (${row},${col}) 時發生錯誤:`, err);
            return null;
        }
    }

    /**
     * 從指定行讀取規格數據
     * @param {Object} worksheet 
     * @param {number} row - 0-indexed row
     * @returns {Object} SpecificationData
     */
    static readSpecificationFromRow(worksheet, row) {
        const spec = this.createSpecificationData();

        try {
            // DEBUG: 決定性的列索引診斷
            let rowDump = [];
            for (let c = 0; c < 15; c++) { // 掃描前15列
                const val = this.getMergedCellValue(worksheet, row, c);
                if (val !== null && String(val).trim() !== '') {
                    // 轉換列索引為字母 (0->A, 1->B...)
                    const colLetter = String.fromCharCode(65 + c);
                    rowDump.push(`${colLetter}[${c}]:"${val}"`);
                }
            }
            console.log(`[SpecExtract][Row${row + 1}] 發現數據: ${rowDump.join(', ')}`);

            // 下一行的數據也打印一下（為了下公差）
            let nextRowDump = [];
            for (let c = 0; c < 15; c++) {
                const val = this.getMergedCellValue(worksheet, row + 1, c);
                if (val !== null && String(val).trim() !== '') {
                    const colLetter = String.fromCharCode(65 + c);
                    nextRowDump.push(`${colLetter}[${c}]:"${val}"`);
                }
            }
            console.log(`[SpecExtract][Row${row + 2}] 下一行數據: ${nextRowDump.join(', ')}`);

            // 定義欄位索引（回歸 VBA 定義，因為截圖可能有隱藏列）
            // VBA: Tool=C(2), Symbol=D(3), Nominal=E(4)/F(5), Sign=G(6), Tol=H(7)
            const nominalCol1 = 4; // E列
            const nominalCol2 = 5; // F列
            const signCol = 6;     // G列
            const tolCol = 7;      // H列

            // 讀取基準值（嘗試 E 列和 F 列）
            let nominalValue = null;
            let nominalVal = this.getMergedCellValue(worksheet, row, nominalCol1);

            if (nominalVal !== null && !isNaN(parseFloat(nominalVal))) {
                nominalValue = parseFloat(nominalVal);
            } else {
                // 嘗試 F 列
                nominalVal = this.getMergedCellValue(worksheet, row, nominalCol2);
                if (nominalVal !== null && !isNaN(parseFloat(nominalVal))) {
                    nominalValue = parseFloat(nominalVal);
                }
            }

            if (nominalValue === null) return spec;

            spec.nominalValue = nominalValue;
            spec.target = nominalValue;

            // --- 符號感知公差計算 (Sign-Aware Tolerance Calculation) ---
            // 讀取上公差 (本行: Row)
            const upperSignVal = this.getMergedCellValue(worksheet, row, signCol);
            const upperTolValRaw = this.getMergedCellValue(worksheet, row, tolCol);

            let upperTolVal = 0;
            if (upperTolValRaw !== null && !isNaN(parseFloat(upperTolValRaw))) {
                upperTolVal = Math.abs(parseFloat(upperTolValRaw));
            }

            const upperSign = upperSignVal ? String(upperSignVal).trim() : '+';

            // 讀取下公差 (下一行: Row + 1)
            const lowerSignVal = this.getMergedCellValue(worksheet, row + 1, signCol);
            const lowerTolValRaw = this.getMergedCellValue(worksheet, row + 1, tolCol);

            let lowerTolVal = 0;
            if (lowerTolValRaw !== null && !isNaN(parseFloat(lowerTolValRaw))) {
                lowerTolVal = Math.abs(parseFloat(lowerTolValRaw));
            }

            const lowerSign = lowerSignVal ? String(lowerSignVal).trim() : '-';

            // 符號感知計算：根據符號決定偏移方向
            let boundary1, boundary2;

            if (upperSign === '±') {
                // ± 符號：雙向對稱偏移
                boundary1 = nominalValue + upperTolVal;
                boundary2 = nominalValue - upperTolVal;
                console.log(`[SpecExtract] ± 對稱公差: ±${upperTolVal}`);
            } else {
                // 根據符號計算兩個邊界
                // + 符號：正偏移 (加上公差值)
                // - 符號：負偏移 (減去公差值)
                if (upperSign === '+') {
                    boundary1 = nominalValue + upperTolVal;
                } else if (upperSign === '-') {
                    boundary1 = nominalValue - upperTolVal;
                } else {
                    // 未知符號，默認為正偏移
                    boundary1 = nominalValue + upperTolVal;
                }

                if (lowerSign === '+') {
                    boundary2 = nominalValue + lowerTolVal;
                } else if (lowerSign === '-') {
                    boundary2 = nominalValue - lowerTolVal;
                } else {
                    // 未知符號，默認為負偏移
                    boundary2 = nominalValue - lowerTolVal;
                }

                console.log(`[SpecExtract] 符號感知計算: ${upperSign}${upperTolVal} / ${lowerSign}${lowerTolVal}`);
            }

            // 自動邊界校準：確保 USL > LSL
            spec.usl = Math.max(boundary1, boundary2);
            spec.lsl = Math.min(boundary1, boundary2);
            spec.upperTolerance = spec.usl - nominalValue;
            spec.lowerTolerance = spec.lsl - nominalValue;

            console.log(`[SpecExtract] 最終規格: Nominal=${nominalValue}, USL=${spec.usl}, LSL=${spec.lsl}`);

            spec.isValid = true;

            return spec;
        } catch (error) {
            console.error('SpecificationExtractor.readSpecificationFromRow error:', error);
            return spec;
        }
    }

    /**
     * 嘗試從工作表自動提取規格
     * @param {Object} workbook 
     * @param {string} inspectionItem 
     * @returns {Object} SpecificationData
     */
    static extractSpecification(workbook, inspectionItem) {
        try {
            const specWs = this.findSpecificationWorksheet(workbook);
            if (!specWs) {
                console.log('未找到規格工作表');
                return this.createSpecificationData();
            }

            const spec = this.findSpecificationByItem(specWs, inspectionItem);
            if (spec.isValid) {
                console.log(`成功提取規格: ${inspectionItem}`, spec);
            } else {
                console.log(`未找到規格: ${inspectionItem}`);
            }

            return spec;
        } catch (error) {
            console.error('SpecificationExtractor.extractSpecification error:', error);
            return this.createSpecificationData();
        }
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = SpecificationExtractor;
}
