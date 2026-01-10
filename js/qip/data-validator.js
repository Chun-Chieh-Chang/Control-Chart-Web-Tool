/**
 * DataValidator - 數據驗證模組
 * 對應 VBA DataValidator.bas
 */
class DataValidator {
    // 常量定義
    static BATCH_NUMBER_HEADER = '生產批號';
    static EXCLUDED_SHEETS = ['處理異常紀錄', '參數配置', '配置歷史', '圖表生成異常紀錄'];
    static DATA_START_ROW = 3; // 數據從第3行開始（1-indexed: 第1行標題，第2行規格，第3行開始是數據）
    static CAVITY_START_COLUMN = 2; // 穴號從第2列（B列）開始

    /**
     * 驗證工作表是否為有效的數據工作表
     * @param {Object} worksheet - SheetJS worksheet object
     * @param {string} sheetName - 工作表名稱
     * @returns {boolean}
     */
    static isValidDataSheet(worksheet, sheetName) {
        try {
            // 檢查工作表名稱是否在排除列表中
            if (this.isExcludedSheet(sheetName)) {
                return false;
            }

            // 檢查 A1 是否包含「生產批號」
            const a1Value = this.getCellValue(worksheet, 'A1');
            if (a1Value !== this.BATCH_NUMBER_HEADER) {
                return false;
            }

            // 檢查第二行是否有數據
            const a2Value = this.getCellValue(worksheet, 'A2');
            if (!a2Value || String(a2Value).trim() === '') {
                return false;
            }

            // 檢查是否有穴號列（至少要有B列）
            const b1Value = this.getCellValue(worksheet, 'B1');
            const b2Value = this.getCellValue(worksheet, 'B2');
            if ((!b1Value || String(b1Value).trim() === '') && 
                (!b2Value || String(b2Value).trim() === '')) {
                return false;
            }

            return true;
        } catch (error) {
            console.error('DataValidator.isValidDataSheet error:', error);
            return false;
        }
    }

    /**
     * 檢查工作表名稱是否在排除列表中
     * @param {string} sheetName 
     * @returns {boolean}
     */
    static isExcludedSheet(sheetName) {
        return this.EXCLUDED_SHEETS.includes(sheetName);
    }

    /**
     * 獲取工作表中的有效數據範圍
     * @param {Object} worksheet 
     * @returns {Object|null} { startRow, endRow, startCol, endCol }
     */
    static getValidDataRange(worksheet) {
        try {
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            
            // 獲取最後一行（A列）
            let lastRow = range.e.r + 1; // 1-indexed
            for (let r = range.e.r; r >= 0; r--) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: 0 });
                if (worksheet[cellAddr] && worksheet[cellAddr].v !== undefined && worksheet[cellAddr].v !== '') {
                    lastRow = r + 1;
                    break;
                }
            }

            if (lastRow < 2) return null;

            // 獲取最後一列
            let lastCol = range.e.c + 1; // 1-indexed
            
            return {
                startRow: 1,
                endRow: lastRow,
                startCol: 1,
                endCol: lastCol
            };
        } catch (error) {
            console.error('DataValidator.getValidDataRange error:', error);
            return null;
        }
    }

    /**
     * 獲取穴號列的數量
     * @param {Object} worksheet 
     * @returns {number}
     */
    static getCavityColumnCount(worksheet) {
        try {
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
            let cavityCount = 0;
            const statsKeywords = ['最大值', '最小值', '平均值', '標準差', '範圍', 
                                   'Max', 'Min', 'Average', 'Avg', 'StdDev', 'Range'];

            // 從第2列開始計數，直到遇到統計列
            for (let c = 1; c <= range.e.c; c++) {
                const cellAddr = XLSX.utils.encode_cell({ r: 0, c: c });
                const cellValue = worksheet[cellAddr] ? String(worksheet[cellAddr].v || '').trim() : '';
                
                // 檢查是否為統計列
                if (this.isStatisticsColumn(cellValue, statsKeywords)) {
                    break;
                }
                
                cavityCount++;
            }

            return cavityCount;
        } catch (error) {
            console.error('DataValidator.getCavityColumnCount error:', error);
            return 0;
        }
    }

    /**
     * 檢查是否為統計列
     * @param {string} columnHeader 
     * @param {Array} statsKeywords 
     * @returns {boolean}
     */
    static isStatisticsColumn(columnHeader, statsKeywords) {
        const header = columnHeader.toLowerCase();
        return statsKeywords.some(keyword => header.includes(keyword.toLowerCase()));
    }

    /**
     * 獲取穴號列的起始列號
     * @returns {number}
     */
    static getCavityStartColumn() {
        return this.CAVITY_START_COLUMN;
    }

    /**
     * 獲取數據起始行號
     * @returns {number}
     */
    static getDataStartRow() {
        return this.DATA_START_ROW;
    }

    /**
     * 輔助函數：從 worksheet 獲取儲存格值
     * @param {Object} worksheet 
     * @param {string} cellAddress - 如 "A1"
     * @returns {*}
     */
    static getCellValue(worksheet, cellAddress) {
        const cell = worksheet[cellAddress];
        if (!cell) return null;
        return cell.v;
    }

    /**
     * 解析範圍字串並返回範圍物件
     * @param {string} rangeStr - 如 "K3:R10"
     * @returns {Object} { startRow, endRow, startCol, endCol } (1-indexed)
     */
    static parseRangeString(rangeStr) {
        try {
            const range = XLSX.utils.decode_range(rangeStr);
            return {
                startRow: range.s.r + 1,
                endRow: range.e.r + 1,
                startCol: range.s.c + 1,
                endCol: range.e.c + 1
            };
        } catch (error) {
            console.error('DataValidator.parseRangeString error:', error);
            return null;
        }
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = DataValidator;
}
