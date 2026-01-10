/**
 * QIPProcessor - 核心處理邏輯
 * 對應 VBA theCode.bas 中的處理核心
 */
class QIPProcessor {
    constructor(config) {
        this.config = config;
        this.errorLogger = new ErrorLogger();
        this.results = {
            inspectionItems: {},
            totalBatches: 0,
            totalCavities: 0,
            totalCavities: 0,
            processedSheets: 0,
            productInfo: { productName: '', measurementUnit: '' }
        };
    }

    /**
     * 處理工作簿
     * @param {Object} workbook - SheetJS workbook
     * @param {Function} progressCallback - 進度回調函數
     * @returns {Object} 處理結果
     */
    async processWorkbook(workbook, progressCallback = () => { }) {
        console.log('開始處理工作簿...');
        console.log('配置:', this.config);

        const sheetCount = workbook.SheetNames.length;
        let processedCount = 0;

        // 計算最大頁面偏移量，決定步長
        let maxOffset = 0;
        if (this.config.cavityGroups) {
            for (let g = 1; g <= 6; g++) {
                if (this.config.cavityGroups[g]) {
                    maxOffset = Math.max(maxOffset, this.config.cavityGroups[g].pageOffset || 0);
                }
            }
        }
        const step = maxOffset + 1;
        console.log(`處理步長: ${step} (最大偏移: ${maxOffset})`);

        // 遍歷所有工作表
        for (let i = 0; i < sheetCount; i += step) {
            const sheetName = workbook.SheetNames[i];
            const worksheet = workbook.Sheets[sheetName];

            try {
                // 更新進度
                progressCallback({
                    current: i + 1,
                    total: sheetCount,
                    message: `處理批次: ${sheetName}`,
                    percent: Math.round((i + 1) / sheetCount * 100)
                });

                // 提取並彙整數據
                await this.processWorksheet(workbook, worksheet, sheetName, i);
                processedCount++;

            } catch (error) {
                console.error(`處理工作表 ${sheetName} 時發生錯誤:`, error);
                this.errorLogger.logError(sheetName, error.message);
            }

            // 讓 UI 有機會更新
            await this.sleep(10);
        }

        this.results.processedSheets = processedCount;

        // 嘗試提取規格數據
        await this.extractSpecifications(workbook, progressCallback);

        // 提取產品資訊 (比照 VBA)
        this.results.productInfo = this.extractProductInfo(workbook);

        console.log('處理完成', this.results);
        return this.getResults();
    }

    /**
     * 處理單個工作表
     * @param {Object} workbook 
     * @param {Object} worksheet 
     * @param {string} sheetName 
     * @param {number} sheetIndex 
     */
    async processWorksheet(workbook, worksheet, sheetName, sheetIndex) {
        // 獲取批號（使用工作表名稱作為批號）
        const batchName = sheetName;

        // 處理每個穴組
        for (let groupIndex = 1; groupIndex <= 6; groupIndex++) {
            const groupConfig = this.config.cavityGroups[groupIndex];

            if (!groupConfig || !groupConfig.cavityIdRange || !groupConfig.dataRange) {
                continue;
            }

            // 計算目標工作表索引
            const targetSheetIndex = sheetIndex + (groupConfig.pageOffset || 0);

            if (targetSheetIndex < 0 || targetSheetIndex >= workbook.SheetNames.length) {
                continue;
            }

            const targetWs = workbook.Sheets[workbook.SheetNames[targetSheetIndex]];

            // 從該穴組提取檢驗項目數據
            const items = this.extractInspectionItemsFromGroup(targetWs, groupConfig);

            for (const item of items) {
                this.addToResults(item.inspectionItem, batchName, item.data);
            }
        }
    }

    /**
     * 從穴組提取檢驗項目數據
     * @param {Object} worksheet 
     * @param {Object} groupConfig 
     * @returns {Array}
     */
    extractInspectionItemsFromGroup(worksheet, groupConfig) {
        const result = [];

        try {
            const idRangeParsed = DataValidator.parseRangeString(groupConfig.cavityIdRange);
            const dataRangeParsed = DataValidator.parseRangeString(groupConfig.dataRange);

            if (!idRangeParsed || !dataRangeParsed) {
                return result;
            }

            // 遍歷數據範圍的每一行（每行可能是不同的檢驗項目）
            for (let rowOffset = 0; rowOffset <= dataRangeParsed.endRow - dataRangeParsed.startRow; rowOffset++) {
                const dataRow = dataRangeParsed.startRow - 1 + rowOffset; // 0-indexed
                const data = {};
                let itemName = '';

                // 提取檢驗項目名稱（從 A 或 B 欄）
                for (let c = 0; c < dataRangeParsed.startCol - 1; c++) {
                    const cellAddr = XLSX.utils.encode_cell({ r: dataRow, c: c });
                    const cell = worksheet[cellAddr];

                    if (cell && cell.v !== undefined) {
                        const value = String(cell.v).trim().replace(/[()]/g, '');
                        if (value && !DataExtractor.isNumericString(value)) {
                            itemName = value;
                            break;
                        }
                    }
                }

                // 提取穴號數據
                const idRow = idRangeParsed.startRow - 1; // 0-indexed
                let hasData = false;

                for (let colOffset = 0; colOffset < idRangeParsed.endCol - idRangeParsed.startCol + 1; colOffset++) {
                    const col = idRangeParsed.startCol - 1 + colOffset;

                    // 獲取穴號 ID
                    const idCellAddr = XLSX.utils.encode_cell({ r: idRow, c: col });
                    const idCell = worksheet[idCellAddr];
                    let cavityId = idCell ? String(idCell.v || '').trim() : '';

                    if (!cavityId) continue;

                    // 提取穴號數字
                    const numMatch = cavityId.match(/\d+/);
                    if (numMatch) {
                        cavityId = numMatch[0];
                    }

                    // 獲取數據
                    const dataCellAddr = XLSX.utils.encode_cell({ r: dataRow, c: col });
                    const dataCell = worksheet[dataCellAddr];

                    if (dataCell && dataCell.v !== undefined) {
                        const cleanValue = DataExtractor.cleanCellValue(dataCell.v);
                        if (cleanValue !== '' && !isNaN(parseFloat(cleanValue))) {
                            data[cavityId] = parseFloat(cleanValue);
                            hasData = true;
                        }
                    }
                }

                if (hasData && itemName) {
                    result.push({
                        inspectionItem: itemName,
                        data: data
                    });
                    console.log(`提取檢驗項目: ${itemName}, 穴數: ${Object.keys(data).length}`);
                }
            }

        } catch (error) {
            console.error('extractInspectionItemsFromGroup error:', error);
            this.errorLogger.logError('數據提取', error.message);
        }

        return result;
    }

    /**
     * 將數據添加到結果
     * @param {string} inspectionItem 
     * @param {string} batchName 
     * @param {Object} data 
     */
    addToResults(inspectionItem, batchName, data) {
        if (!inspectionItem || Object.keys(data).length === 0) return;

        if (!this.results.inspectionItems[inspectionItem]) {
            this.results.inspectionItems[inspectionItem] = {
                batches: {},
                allCavities: new Set(),
                specification: null
            };
        }

        const item = this.results.inspectionItems[inspectionItem];

        // 如果批次已存在，合併數據
        if (item.batches[batchName]) {
            Object.assign(item.batches[batchName], data);
        } else {
            item.batches[batchName] = { ...data };
            this.results.totalBatches++;
        }

        // 記錄所有穴號
        for (const cavityId of Object.keys(data)) {
            item.allCavities.add(cavityId);
        }

        this.results.totalCavities = Math.max(
            this.results.totalCavities,
            item.allCavities.size
        );
    }

    /**
     * 提取規格數據
     * @param {Object} workbook 
     * @param {Function} progressCallback 
     */
    async extractSpecifications(workbook, progressCallback) {
        progressCallback({
            current: 100,
            total: 100,
            message: '提取規格數據...',
            percent: 95
        });

        for (const itemName of Object.keys(this.results.inspectionItems)) {
            const spec = SpecificationExtractor.extractSpecification(workbook, itemName);
            this.results.inspectionItems[itemName].specification = spec;
        }
    }

    /**
     * 提取產品資訊 (Product Name & Measurement Unit)
     * @param {Object} workbook 
     * @returns {Object}
     */
    extractProductInfo(workbook) {
        let productName = '';
        let measurementUnit = '';

        for (const sheetName of workbook.SheetNames) {
            const ws = workbook.Sheets[sheetName];

            // 提取產品名稱 (P2 或 P3)
            // P2 = Index {c:15, r:1}, P3 = {c:15, r:2}
            if (!productName) {
                const cellP2 = ws['P2'];
                if (cellP2 && cellP2.v) productName = String(cellP2.v).trim();

                if (!productName) {
                    const cellP3 = ws['P3'];
                    if (cellP3 && cellP3.v) productName = String(cellP3.v).trim();
                }

                // 掃描 P2:V3 區域
                if (!productName) {
                    for (let r = 1; r <= 2; r++) {
                        for (let c = 15; c <= 21; c++) {
                            const addr = XLSX.utils.encode_cell({ r, c });
                            const cell = ws[addr];
                            if (cell && cell.v && String(cell.v).trim() !== '0' && String(cell.v).trim() !== 'False') {
                                productName = String(cell.v).trim();
                                break;
                            }
                        }
                        if (productName) break;
                    }
                }
            }

            // 提取測量單位 (W23) -> Index {c:22, r:22}
            if (!measurementUnit) {
                const cellW23 = ws['W23'];
                if (cellW23 && cellW23.v) {
                    let val = String(cellW23.v).trim();
                    // 移除 "單位：" 前綴
                    val = val.replace(/單位[:：]/g, '').trim();
                    measurementUnit = val;
                }
            }

            if (productName && measurementUnit) break;
        }

        console.log('提取到的產品資訊:', { productName, measurementUnit });
        return { productName, measurementUnit };
    }

    /**
     * 獲取處理結果
     * @returns {Object}
     */
    getResults() {
        return {
            inspectionItems: this.results.inspectionItems,
            totalBatches: this.results.totalBatches,
            totalCavities: this.results.totalCavities,
            totalBatches: this.results.totalBatches,
            totalCavities: this.results.totalCavities,
            processedSheets: this.results.processedSheets,
            productInfo: this.results.productInfo,
            itemCount: Object.keys(this.results.inspectionItems).length,
            errors: this.errorLogger.getErrors(),
            hasErrors: this.errorLogger.hasErrors()
        };
    }

    /**
     * 輔助函數：延遲
     * @param {number} ms 
     */
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = QIPProcessor;
}
