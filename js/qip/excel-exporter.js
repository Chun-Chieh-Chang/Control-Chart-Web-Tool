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
        // A: Target, B: USL, C: LSL, D: 生產批號, E+: 穴號
        const headerRow = ['Target', 'USL', 'LSL', '生產批號'];
        for (const cavityNum of cavities) {
            headerRow.push(`${cavityNum}號穴`);
        }
        data.push(headerRow);

        // 填寫數據行 (從第2行開始，對應 AOA 索引 1)
        itemData.batches.forEach((batch, index) => {
            const rowIndex = index + 1; // 在 AOA 中的索引
            const batchData = batch.data;

            // 基礎行：[A, B, C, D]
            const row = ['', '', '', batch.name];

            // A, B, C 欄只在第2行 (index 1) 填寫規格
            if (rowIndex === 1) {
                if (itemData.specification && itemData.specification.isValid) {
                    row[0] = itemData.specification.target;
                    row[1] = itemData.specification.usl;
                    row[2] = itemData.specification.lsl;
                } else {
                    row[0] = row[1] = row[2] = '未設定';
                }
            }

            // 修改：A5/B5 為產品名稱，A6/B6 為單位
            if (rowIndex === 4 && productInfo) { // Row 5
                row[0] = 'ProductName';
                row[1] = productInfo.productName || 'N/A';
            }
            if (rowIndex === 5 && productInfo) { // Row 6
                row[0] = 'MeasurementUnit';
                row[1] = productInfo.measurementUnit || 'mm';
            }

            // E 欄開始填寫量測數據
            for (const cavityNum of cavities) {
                const value = batchData[String(cavityNum)];
                row.push(value !== undefined ? value : '');
            }

            data.push(row);
        });

        // 確保至少有 6 行以完整顯示 A5/B5, A6/B6 (即使批次很少)
        if (productInfo) {
            while (data.length < 6) {
                const emptyRow = ['', '', '', ''];
                if (data.length === 4) {
                    emptyRow[0] = 'ProductName';
                    emptyRow[1] = productInfo.productName || 'N/A';
                } else if (data.length === 5) {
                    emptyRow[0] = 'MeasurementUnit';
                    emptyRow[1] = productInfo.measurementUnit || 'mm';
                }
                data.push(emptyRow);
            }
        }

        // 創建工作表
        const worksheet = XLSX.utils.aoa_to_sheet(data);

        // 套用樣式 (After aoa_to_sheet)
        for (let r = 0; r < data.length; r++) {
            for (let c = 0; c < (data[r] ? data[r].length : 0); c++) {
                const cellAddr = XLSX.utils.encode_cell({ r: r, c: c });
                if (!worksheet[cellAddr]) continue;

                if (!worksheet[cellAddr].s) worksheet[cellAddr].s = {};

                // 字體設定
                const isMetadataCell = (r === 4 || r === 5) && (c === 0 || c === 1);
                const isBatchColumn = (c === 3);
                const isHeader = (r === 0);

                if (isHeader || isBatchColumn || isMetadataCell) {
                    worksheet[cellAddr].s.font = { name: 'Microsoft JhengHei', bold: isHeader };
                } else {
                    worksheet[cellAddr].s.font = { name: 'Times New Roman' };
                }

                // 數值格式
                if (r > 0 && ((c >= 0 && c <= 2) || (c >= 4))) {
                    const val = worksheet[cellAddr].v;
                    if (typeof val === 'number') {
                        worksheet[cellAddr].z = '0.0000';
                    }
                }
            }
        }

        // 設置列寬
        const colWidths = [
            { wch: 12 },  // A: Target
            { wch: 12 },  // B: USL
            { wch: 12 },  // C: LSL
            { wch: 20 }   // D: 生產批號
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
