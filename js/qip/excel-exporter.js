/**
 * ExcelExporter - Excel 輸出模組
 * 使用 SheetJS 生成 Excel 檔案，輸出格式與 VBA 版本相同或更優
 */
class ExcelExporter {
    constructor() {
        this.workbook = XLSX.utils.book_new();
        this.workbook.Props = {
            Title: 'QIP 數據提取結果',
            Author: 'QIP Data Extract Tool',
            CreatedDate: new Date()
        };
    }

    /**
     * 從處理結果創建 Excel
     * @param {Object} results - QIPProcessor 的處理結果
     * @param {string} productCode - 產品品號
     * @returns {ExcelExporter}
     */
    createFromResults(results, productCode = '') {
        console.log('開始創建 Excel...', results);

        for (const [itemName, itemData] of Object.entries(results.inspectionItems)) {
            this.addInspectionSheet(itemName, itemData, productCode, results.productInfo);
        }

        return this;
    }

    /**
     * 添加檢驗項目工作表
     * @param {string} sheetName - 工作表名稱（檢驗項目）
     * @param {Object} itemData - 項目數據 { batches, allCavities, specification }
     * @param {string} productCode - 產品品號
     */
    /**
     * 添加檢驗項目工作表
     * @param {string} sheetName - 工作表名稱（檢驗項目）
     * @param {Object} itemData - 項目數據 { batches, allCavities, specification }
     * @param {string} productCode - 產品品號
     * @param {Object} productInfo - 產品資訊 { productName, measurementUnit }
     */
    addInspectionSheet(sheetName, itemData, productCode = '', productInfo = null) {
        // 清理工作表名稱
        const cleanName = this.cleanSheetName(sheetName);

        // 獲取並排序所有穴號
        const cavities = Array.from(itemData.allCavities)
            .map(Number)
            .sort((a, b) => a - b);

        // 獲取批號列表
        const batches = Object.keys(itemData.batches);

        // 構建數據陣列
        const data = [];

        // 第1行：標題行
        const headerRow = ['生產批號', 'Target', 'USL', 'LSL'];
        for (const cavityNum of cavities) {
            headerRow.push(`${cavityNum}號穴`);
        }
        data.push(headerRow);

        // 第2行：規格數據
        const specRow = [''];
        if (itemData.specification && itemData.specification.isValid) {
            specRow.push(itemData.specification.target);
            specRow.push(itemData.specification.usl);
            specRow.push(itemData.specification.lsl);
        } else {
            specRow.push('未設定', '未設定', '未設定');
        }
        // 規格行的穴號欄位留空
        for (let i = 0; i < cavities.length; i++) {
            specRow.push('');
        }
        data.push(specRow);

        // 數據行（從第3行開始）
        for (const batchName of batches) {
            const batchData = itemData.batches[batchName];
            const row = [batchName, '', '', '']; // 批號 + 3個空欄（規格欄）

            for (const cavityNum of cavities) {
                const value = batchData[String(cavityNum)];
                row.push(value !== undefined ? value : '');
            }

            data.push(row);
        }

        // 創建工作表
        const worksheet = XLSX.utils.aoa_to_sheet(data);

        // 設置列寬
        const colWidths = [
            { wch: 15 },  // 生產批號
            { wch: 12 },  // Target
            { wch: 12 },  // USL
            { wch: 12 }   // LSL
        ];
        for (let i = 0; i < cavities.length; i++) {
            colWidths.push({ wch: 10 });
        }
        worksheet['!cols'] = colWidths;

        // 設置儲存格樣式（SheetJS 基礎版不支援完整樣式，但我們記錄意圖）
        // 標題行樣式
        this.setHeaderStyles(worksheet, headerRow.length);

        // 規格行樣式
        this.setSpecificationStyles(worksheet, itemData.specification);

        // 添加到工作簿
        XLSX.utils.book_append_sheet(this.workbook, worksheet, cleanName);

        // 寫入產品資訊 (比照 VBA: B5/C5 標題, B6/C6 內容)
        if (productInfo) {
            // 修正：從 B 欄開始寫入，避免覆蓋 A 欄的生產批號
            const infoHeader = [['ProductName', 'MeasurementUnit']]; // Row 5 (index 4), starting at Col B
            const infoData = [[productInfo.productName || '', productInfo.measurementUnit || '']]; // Row 6 (index 5), starting at Col B

            // 使用 sheet_add_aoa 寫入 (origin: B5 means start at Row 5, Col B)
            XLSX.utils.sheet_add_aoa(worksheet, infoHeader, { origin: 'B5' });
            XLSX.utils.sheet_add_aoa(worksheet, infoData, { origin: 'B6' });
        }

        console.log(`創建工作表: ${cleanName}, 批次數: ${batches.length}, 穴數: ${cavities.length}`);
    }

    /**
     * 設置標題行樣式
     * @param {Object} worksheet 
     * @param {number} colCount 
     */
    setHeaderStyles(worksheet, colCount) {
        // SheetJS 免費版樣式支援有限，這裡主要設置數字格式
        // 使用 xlsx-style 或 exceljs 可以獲得完整樣式支援
        for (let c = 0; c < colCount; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r: 0, c: c });
            if (worksheet[cellAddr]) {
                // 標記為標題（供後續處理）
                worksheet[cellAddr].s = {
                    font: { bold: true },
                    fill: { fgColor: { rgb: '92D050' } },
                    alignment: { horizontal: 'center' }
                };
            }
        }
    }

    /**
     * 設置規格行樣式
     * @param {Object} worksheet 
     * @param {Object} specification 
     */
    setSpecificationStyles(worksheet, specification) {
        // 設置規格數字格式
        for (let c = 1; c <= 3; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r: 1, c: c });
            if (worksheet[cellAddr] && typeof worksheet[cellAddr].v === 'number') {
                worksheet[cellAddr].z = '0.0000';
            }
        }
    }

    /**
     * 清理工作表名稱
     * @param {string} name 
     * @returns {string}
     */
    cleanSheetName(name) {
        // 移除不允許的字符
        let result = name.replace(/[\\/:*?"<>|]/g, '_');
        // 限制長度（Excel 最多31字符）
        if (result.length > 31) {
            result = result.substring(0, 31);
        }
        return result.trim() || '未命名項目';
    }

    /**
     * 導出 Excel 檔案
     * @param {string} filename - 檔案名稱（不含副檔名）
     */
    export(filename = 'QIP_數據提取結果') {
        const fullFilename = `${filename}.xlsx`;
        XLSX.writeFile(this.workbook, fullFilename);
        console.log(`Excel 檔案已導出: ${fullFilename}`);
    }

    /**
     * 獲取 Excel 二進制數據
     * @returns {ArrayBuffer}
     */
    getArrayBuffer() {
        return XLSX.write(this.workbook, {
            bookType: 'xlsx',
            type: 'array'
        });
    }

    /**
     * 獲取 Blob 物件（用於下載）
     * @returns {Blob}
     */
    getBlob() {
        const buffer = this.getArrayBuffer();
        return new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
    }

    /**
     * 觸發下載
     * @param {string} filename 
     */
    download(filename = 'QIP_數據提取結果') {
        const blob = this.getBlob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${filename}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        console.log(`開始下載: ${filename}.xlsx`);
    }

    /**
     * 獲取工作表數量
     * @returns {number}
     */
    getSheetCount() {
        return this.workbook.SheetNames.length;
    }

    /**
     * 重置（創建新工作簿）
     */
    reset() {
        this.workbook = XLSX.utils.book_new();
        this.workbook.Props = {
            Title: 'QIP 數據提取結果',
            Author: 'QIP Data Extract Tool',
            CreatedDate: new Date()
        };
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ExcelExporter;
}
