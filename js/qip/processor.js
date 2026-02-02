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
     * 處理多個工作簿
     * @param {Array} workbooks - Array of SheetJS workbooks
     * @param {Function} progressCallback - 進度回調函數
     * @returns {Object} 處理結果
     */
    async processWorkbooks(workbooks, progressCallback = () => { }) {
        console.log(`開始處理 ${workbooks.length} 個工作簿...`);

        for (let i = 0; i < workbooks.length; i++) {
            const wb = workbooks[i];
            const currentRatio = i / workbooks.length;
            const nextRatio = (i + 1) / workbooks.length;

            await this.processWorkbook(wb, (p) => {
                // 調整進度回調，使其反應整體的進度
                progressCallback({
                    current: p.current,
                    total: p.total,
                    message: `[檔案 ${i + 1}/${workbooks.length}] ${p.message}`,
                    percent: Math.round((currentRatio + (p.percent / 100) * (nextRatio - currentRatio)) * 100)
                });
            }, true); // pass true to indicate it's part of a multi-batch
        }

        // 最後統一提取規格與產品資訊 (從第一個有效的工作簿提取即可，假設格式相同)
        if (workbooks.length > 0) {
            await this.extractSpecifications(workbooks[0], progressCallback);
            const info = this.extractProductInfo(workbooks[0]);
            // 如果 results.productInfo 尚未設定或為空，則更新
            if (!this.results.productInfo.productName) {
                this.results.productInfo = info;
            }
        }

        console.log('所有工作簿處理完成', this.results);
        return this.getResults();
    }

    /**
     * 處理工作簿
     * @param {Object} workbook - SheetJS workbook
     * @param {Function} progressCallback - 進度回調函數
     * @param {boolean} cumulative - 是否為累加模式（不重置統計）
     * @returns {Object} 處理結果
     */
    async processWorkbook(workbook, progressCallback = () => { }, cumulative = false) {
        console.log('開始處理工作簿...');

        const sheetCount = workbook.SheetNames.length;
        let processedCount = 0;

        // 計算最小與最大頁面偏移量，決定步長與起始點
        let minOffset = Infinity;
        let maxOffset = -Infinity;
        let hasActiveGroups = false;
        if (this.config.cavityGroups) {
            for (let g = 1; g <= 6; g++) {
                if (this.config.cavityGroups[g] && this.config.cavityGroups[g].cavityIdRange) {
                    const offset = this.config.cavityGroups[g].pageOffset || 0;
                    minOffset = Math.min(minOffset, offset);
                    maxOffset = Math.max(maxOffset, offset);
                    hasActiveGroups = true;
                }
            }
        }

        if (!hasActiveGroups) {
            minOffset = 0;
            maxOffset = 0;
        }

        const step = maxOffset - minOffset + 1;

        // 遍歷所有工作表
        // i 代表批次的基準偏移量（從 0 開始）
        for (let i = 0; i < sheetCount; i += step) {
            // 如果基準點 + 最小偏移量已經超出範圍，說明已無完整批次
            if (i + minOffset >= sheetCount) break;

            const sheetName = workbook.SheetNames[i + minOffset];
            const worksheet = workbook.Sheets[sheetName];

            try {
                // 更新進度
                progressCallback({
                    current: i + minOffset + 1,
                    total: sheetCount,
                    message: `處理批次: ${sheetName}`,
                    percent: Math.round((i + minOffset + 1) / sheetCount * 100)
                });

                // 提取並彙整數據 (傳入 i 作為基準偏移)
                await this.processWorksheet(workbook, worksheet, sheetName, i);
                processedCount++;

            } catch (error) {
                console.error(`處理工作表 ${sheetName} 時發生錯誤:`, error);
                this.errorLogger.logError(sheetName, error.message);
            }

            await this.sleep(5);
        }

        if (cumulative) {
            this.results.processedSheets += processedCount;
        } else {
            this.results.processedSheets = processedCount;
            // 非累加模式下才執行最後的提取
            await this.extractSpecifications(workbook, progressCallback);
            this.results.productInfo = this.extractProductInfo(workbook);
        }

        return cumulative ? this.results : this.getResults();
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

            // 獲取合併儲存格資訊（用於處理 VBA MergeArea 邏輯）
            const merges = worksheet['!merges'] || [];

            // 輔助函數：安全地從可能的合併儲存格中獲取值
            const safeGetMergedValue = (row, col) => {
                try {
                    const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
                    const cell = worksheet[cellAddr];
                    let value = cell && cell.v !== undefined ? cell.v : null;

                    // 如果沒有值，檢查是否是合併儲存格的一部分
                    if (!value || value === '') {
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
                    console.error(`[QIP] 讀取儲存格 (${row},${col}) 時發生錯誤:`, err);
                    return null;
                }
            };

            // 遍歷數據範圍的每一行（每行可能是不同的檢驗項目）
            for (let rowOffset = 0; rowOffset <= dataRangeParsed.endRow - dataRangeParsed.startRow; rowOffset++) {
                const dataRow = dataRangeParsed.startRow - 1 + rowOffset; // 0-indexed
                const data = {};
                let itemName = '';

                // 提取检验项目名称 - 只从 A 列读取
                // 注意：检验项目名称可以是数字（如 "1", "2"），必须作为文字处理
                let tempValue = safeGetMergedValue(dataRow, 0);  // A 列

                if (tempValue !== null && tempValue !== undefined) {
                    itemName = String(tempValue).trim();
                }

                // 如果 A 列没有，尝试 B 列
                if (!itemName || itemName === '') {
                    tempValue = safeGetMergedValue(dataRow, 1);  // B 列
                    if (tempValue !== null && tempValue !== undefined) {
                        itemName = String(tempValue).trim();
                    }
                }

                if (itemName && itemName !== '') {
                    console.log(`[QIP][Row${dataRow + 1}] ✓ 找到检验项目: "${itemName}"`);
                } else {
                    console.warn(`[QIP][Row${dataRow + 1}] ✗ A/B 列都没有内容，跳过此行`);
                    continue;
                }


                // 提取穴號數據
                const idRow = idRangeParsed.startRow - 1; // 0-indexed
                let hasData = false;

                for (let colOffset = 0; colOffset < idRangeParsed.endCol - idRangeParsed.startCol + 1; colOffset++) {
                    const col = idRangeParsed.startCol - 1 + colOffset;

                    // 獲取穴號 ID（處理合併儲存格）
                    let cavityId = safeGetMergedValue(idRow, col);
                    if (!cavityId) continue;

                    cavityId = String(cavityId).trim();

                    // 提取穴號數字（支援 "1號穴"、"CAV1"、"1" 等格式）
                    const numMatch = cavityId.match(/\d+/);
                    if (numMatch) {
                        cavityId = numMatch[0];
                    }

                    // 獲取數據（處理合併儲存格）
                    const cellValue = safeGetMergedValue(dataRow, col);

                    if (cellValue !== undefined && cellValue !== null) {
                        const cleanValue = DataExtractor.cleanCellValue(cellValue);
                        if (cleanValue !== '' && !isNaN(parseFloat(cleanValue))) {
                            data[cavityId] = parseFloat(cleanValue);
                            hasData = true;
                        }
                    }
                }

                if (hasData && itemName) {
                    const cavityIds = Object.keys(data).sort((a, b) => parseInt(a) - parseInt(b));
                    result.push({
                        inspectionItem: itemName,
                        data: data
                    });
                    console.log(`[QIP][Row${dataRow + 1}] ✓ 提取成功: "${itemName}", 穴號: [${cavityIds.join(', ')}], 數量: ${cavityIds.length}`);
                } else if (!hasData && itemName) {
                    console.warn(`[QIP][Row${dataRow + 1}] ⚠ 檢驗項目 "${itemName}" 無有效數據`);
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
                batches: [],
                allCavities: new Set(),
                specification: null
            };
            console.log(`[QIP] 新增檢驗項目: ${inspectionItem}`);
        }

        const item = this.results.inspectionItems[inspectionItem];
        const cavityIds = Object.keys(data);

        // Always add a new batch entry instead of merging by name
        // This supports multiple batches with the same name (e.g. "Setup") from different files
        item.batches.push({ name: batchName, data: { ...data } });
        console.log(`[QIP] 新增批次 - 項目: ${inspectionItem}, 批次: ${batchName}, 穴號: [${cavityIds.join(', ')}]`);

        // 記錄所有穴號
        for (const cavityId of cavityIds) {
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
        // 計算實際的唯一批次數（從所有檢驗項目中收集所有唯一批次名稱）
        const allBatchOccurrencesCount = [];
        const uniqueBatchNames = new Set();

        for (const itemName of Object.keys(this.results.inspectionItems)) {
            const item = this.results.inspectionItems[itemName];
            allBatchOccurrencesCount.push(item.batches.length);
            for (const b of item.batches) {
                uniqueBatchNames.add(b.name);
            }
        }
        const maxBatchesInAnyItem = Math.max(...allBatchOccurrencesCount, 0);
        const actualUniqueBatchNamesCount = uniqueBatchNames.size;

        console.log(`[QIP] 處理結果統計:`);
        console.log(`  ├─ 檢驗項目數: ${Object.keys(this.results.inspectionItems).length}`);
        console.log(`  ├─ 最大單項批次數: ${maxBatchesInAnyItem}`);
        console.log(`  ├─ 唯一批次名稱數: ${actualUniqueBatchNamesCount}`);
        console.log(`  ├─ 最大穴數: ${this.results.totalCavities}`);
        console.log(`  └─ 處理工作表數: ${this.results.processedSheets}`);

        return {
            inspectionItems: this.results.inspectionItems,
            totalBatches: maxBatchesInAnyItem,
            uniqueBatchNamesCount: actualUniqueBatchNamesCount,
            totalCavities: this.results.totalCavities,
            processedSheets: this.results.processedSheets,
            productInfo: this.results.productInfo,
            productCode: this.config.productCode || '',
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
