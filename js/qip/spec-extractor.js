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
     * 從指定行讀取規格數據
     * @param {Object} worksheet 
     * @param {number} row - 0-indexed row
     * @returns {Object} SpecificationData
     */
    static readSpecificationFromRow(worksheet, row) {
        const spec = this.createSpecificationData();

        try {
            // 定義欄位索引（根據VBA中的定義）
            const symbolCol = 2;    // C欄
            const nominalCol1 = 4;  // E欄
            const nominalCol2 = 5;  // F欄
            const upperSignCol = 6; // G欄
            const upperTolCol = 7;  // H欄
            const lowerSignCol = 8; // I欄
            const lowerTolCol = 9;  // J欄

            // 讀取符號
            const symbolCell = worksheet[XLSX.utils.encode_cell({ r: row, c: symbolCol })];
            spec.symbol = symbolCell ? String(symbolCell.v || '').trim() : '';

            // 讀取基準值（嘗試E、F欄）
            let nominalValue = null;
            for (const col of [nominalCol1, nominalCol2]) {
                const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
                if (cell && !isNaN(parseFloat(cell.v))) {
                    nominalValue = parseFloat(cell.v);
                    break;
                }
                // 嘗試下一行
                const nextCell = worksheet[XLSX.utils.encode_cell({ r: row + 1, c: col })];
                if (nextCell && !isNaN(parseFloat(nextCell.v))) {
                    nominalValue = parseFloat(nextCell.v);
                    break;
                }
            }

            if (nominalValue === null) return spec;

            spec.nominalValue = nominalValue;
            spec.target = nominalValue;

            // 讀取上公差
            const upperSignCell = worksheet[XLSX.utils.encode_cell({ r: row, c: upperSignCol })];
            const upperTolCell = worksheet[XLSX.utils.encode_cell({ r: row, c: upperTolCol })];
            let upperTol = 0;

            if (upperTolCell && !isNaN(parseFloat(upperTolCell.v))) {
                upperTol = Math.abs(parseFloat(upperTolCell.v));
            }

            // 根據符號確定正負
            const upperSign = upperSignCell ? String(upperSignCell.v).trim() : '+';
            spec.upperTolerance = upperTol;

            // 讀取下公差
            const lowerSignCell = worksheet[XLSX.utils.encode_cell({ r: row + 1, c: lowerSignCol })];
            const lowerTolCell = worksheet[XLSX.utils.encode_cell({ r: row + 1, c: lowerTolCol })];
            let lowerTol = 0;

            if (lowerTolCell && !isNaN(parseFloat(lowerTolCell.v))) {
                lowerTol = Math.abs(parseFloat(lowerTolCell.v));
            }

            spec.lowerTolerance = lowerTol;

            // 計算 USL 和 LSL
            spec.usl = spec.nominalValue + spec.upperTolerance;
            spec.lsl = spec.nominalValue - spec.lowerTolerance;
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
