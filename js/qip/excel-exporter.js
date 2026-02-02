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
     * @param {Object} productInfo - 產品資訊 { productName, measurementUnit }
     */
    addInspectionSheet(sheetName, itemData, productCode = '', productInfo = null) {
        // 清理工作表名稱
        const cleanName = this.cleanSheetName(sheetName);

        // 獲取並排序所有穴號
        const cavities = Array.from(itemData.allCavities)
            .map(Number)
            .sort((a, b) => a - b);

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
        for (const batch of itemData.batches) {
            const batchName = batch.name;
            const batchData = batch.data;
            const row = [batchName, '', '', '']; // 批號 + 3個空欄（規格欄）

            for (const cavityNum of cavities) {
                const value = batchData[String(cavityNum)];
                row.push(value !== undefined ? value : '');
            }

            data.push(row);
        }

        // 寫入產品資訊到固定位置 (B5/B6)，以滿足其他工具的讀寫需求
        if (productInfo) {
            const productLabels = ['', 'ProductName', 'MeasurementUnit'];
            const productValues = ['', productInfo.productName || 'N/A', productInfo.measurementUnit || 'mm'];

            // 補充空行直到至少有 4 行 (Row 1-4)，確保能放入第 5, 6 行
            while (data.length < 4) {
                data.push(['', '', '', '']);
            }

            // 插入到 index 4 (第 5 行) 和 index 5 (第 6 行)
            // 這樣原本的 Batch 3, 4... 會被推到後面，符合 VBA 工具讀取 metadata 的結構
            data.splice(4, 0, productLabels, productValues);
        }

        // 創建工作表
        const worksheet = XLSX.utils.aoa_to_sheet(data);

        // 套用數據行字體 (After aoa_to_sheet)
        for (let r = 2; r < data.length; r++) {
            // 判斷是否為產品資訊列 (固定在 Row 5, 6)
            const isProductInfoRow = (r === 4 || r === 5);

            for (let c = 0; c < (data[r] ? data[r].length : 0); c++) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: c });
                if (worksheet[cellAddr]) {
                    if (!worksheet[cellAddr].s) worksheet[cellAddr].s = {};
                    worksheet[cellAddr].s.font = {
                        name: (c === 0 || isProductInfoRow) ? 'Microsoft JhengHei' : 'Times New Roman'
                    };
                }
            }
        }

        // 設置列寬
        const colWidths = [
            { wch: 20 },  // 生產批號（拉寬以防檔名較長）
            { wch: 12 },  // Target
            { wch: 12 },  // USL
            { wch: 12 }   // LSL
        ];
        for (let i = 0; i < cavities.length; i++) {
            colWidths.push({ wch: 10 });
        }
        worksheet['!cols'] = colWidths;

        // 設置儲存格樣式
        this.setHeaderStyles(worksheet, headerRow.length);
        this.setSpecificationStyles(worksheet, itemData.specification);

        // 添加到工作簿
        XLSX.utils.book_append_sheet(this.workbook, worksheet, cleanName);

        console.log(`創建工作表: ${cleanName}, 批次數: ${itemData.batches.length}, 穴數: ${cavities.length}`);
    }

    /**
     * 設置標題行樣式
     * @param {Object} worksheet 
     * @param {number} colCount 
     */
    setHeaderStyles(worksheet, colCount) {
        for (let c = 0; c < colCount; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r: 0, c: c });
            if (worksheet[cellAddr]) {
                worksheet[cellAddr].s = {
                    font: { bold: true, name: 'Microsoft JhengHei' },
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
        for (let c = 1; c <= 3; c++) {
            const cellAddr = XLSX.utils.encode_cell({ r: 1, c: c });
            if (worksheet[cellAddr]) {
                worksheet[cellAddr].z = '0.0000';
                if (!worksheet[cellAddr].s) worksheet[cellAddr].s = {};
                worksheet[cellAddr].s.font = { name: 'Times New Roman' };
            }
        }
    }

    /**
     * 清理工作表名稱
     * @param {string} name 
     * @returns {string}
     */
    cleanSheetName(name) {
        let result = name.replace(/[\\/:*?"<>|]/g, '_');
        if (result.length > 31) {
            result = result.substring(0, 31);
        }
        return result.trim() || '未命名項目';
    }

    /**
     * 導出 Excel 檔案
     */
    export(filename = 'QIP_數據提取結果') {
        const fullFilename = `${filename}.xlsx`;
        XLSX.writeFile(this.workbook, fullFilename);
        console.log(`Excel 檔案已導出: ${fullFilename}`);
    }

    getArrayBuffer() {
        return XLSX.write(this.workbook, {
            bookType: 'xlsx',
            type: 'array'
        });
    }

    getBlob() {
        const buffer = this.getArrayBuffer();
        return new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
    }

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

    getSheetCount() {
        return this.workbook.SheetNames.length;
    }

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
