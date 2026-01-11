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
            // 定義欄位索引（根據用戶 QIP 報表截圖調整）
            // Row 26: (1)[A] Tool[B] 1.10[C] +[D] 0.13[E]
            // Col A = 0, Col B = 1, Col C = 2, Col D = 3, Col E = 4
            // Row 27:                        -[D] 0.13[E]
            const nominalCol = 2;   // C欄 (Col index 2)
            const signCol = 3;      // D欄 (Col index 3)
            const tolCol = 4;       // E欄 (Col index 4)

            // 讀取基準值（C欄）
            let nominalValue = null;
            const nominalVal = this.getMergedCellValue(worksheet, row, nominalCol);

            if (nominalVal !== null && !isNaN(parseFloat(nominalVal))) {
                nominalValue = parseFloat(nominalVal);
            } else {
                // 如果本行沒有，嘗試從合併儲存格獲取或假設與上一行相同（不太可能，通常是itemName行有target）
                // 這裡我們假設Target值必須存在
            }

            if (nominalValue === null) return spec;

            spec.nominalValue = nominalValue;
            spec.target = nominalValue;

            // --- 讀取上公差 (本行: Row) ---
            const upperSignVal = this.getMergedCellValue(worksheet, row, signCol);
            const upperTolValRaw = this.getMergedCellValue(worksheet, row, tolCol);

            let upperTolVal = 0;
            if (upperTolValRaw !== null && !isNaN(parseFloat(upperTolValRaw))) {
                upperTolVal = Math.abs(parseFloat(upperTolValRaw));
            }

            const upperSign = upperSignVal ? String(upperSignVal).trim() : '+';
            // 如果符號是 '-'，則為負；否則為正
            const upperActualTol = (upperSign === '-') ? -upperTolVal : upperTolVal;

            spec.upperTolerance = upperActualTol;
            spec.usl = spec.nominalValue + spec.upperTolerance;

            // --- 讀取下公差 (下一行: Row + 1) ---
            // 注意：下公差通常在下一行，與 Target 相同的列通常是空的
            const lowerSignVal = this.getMergedCellValue(worksheet, row + 1, signCol);
            const lowerTolValRaw = this.getMergedCellValue(worksheet, row + 1, tolCol);

            let lowerTolVal = 0;
            if (lowerTolValRaw !== null && !isNaN(parseFloat(lowerTolValRaw))) {
                lowerTolVal = Math.abs(parseFloat(lowerTolValRaw));
            }

            const lowerSign = lowerSignVal ? String(lowerSignVal).trim() : '-'; // 預設可能是負
            // 下公差計算：如果是 '-' 則減去，如果是 '+' 則加上
            // LSL = Target + (Sign * Value)
            const lowerActualTol = (lowerSign === '-') ? -lowerTolVal : lowerTolVal;

            spec.lowerTolerance = lowerActualTol; // 這裡存儲的是相對於 Target 的偏移量
            spec.lsl = spec.nominalValue + spec.lowerTolerance;

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
