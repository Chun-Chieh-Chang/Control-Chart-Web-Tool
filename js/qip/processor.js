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
            processedSheets: 0,
            productInfo: { productName: '', measurementUnit: '' }
        };
        this.currentWorkbookId = 0;
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

            this.currentWorkbookId = i;
            await this.processWorkbook(wb, (p) => {
                // 調整進度回調，使其反應整體的進度
                progressCallback({
                    current: p.current,
                    total: p.total,
                    message: `[檔案 ${i + 1}/${workbooks.length}] ${p.message}`,
                    percent: Math.round((currentRatio + (p.percent / 100) * (nextRatio - currentRatio)) * 100)
                });
            }, true, i); // pass workbook index as ID
        }

        // 從所有工作簿中尋找規格與產品資訊
        await this.extractSpecifications(workbooks, progressCallback);
        const info = this.extractProductInfo(workbooks);
        this.results.productInfo = info;

        console.log('所有工作簿處理完成', this.results);
        return this.getResults();
    }

    /**
     * 處理工作簿
     * @param {Object} workbook - SheetJS workbook
     * @param {Function} progressCallback - 進度回調函數
     * @param {boolean} cumulative - 是否為累加模式（不重置統計）
     * @param {number} workbookId - 當前工作簿 ID
     * @returns {Object} 處理結果
     */
    async processWorkbook(workbook, progressCallback = () => { }, cumulative = false, workbookId = 0) {
        console.log('開始處理工作簿...');

        const sheetCount = workbook.SheetNames.length;
        let processedCount = 0;

        // 計算最小與最大頁面偏移量，決定步長與起始點
        let minOffset = Infinity;
        let maxOffset = -Infinity;
        let hasActiveGroups = false;
        let refGroup = null;

        if (this.config.cavityGroups) {
            for (let g = 1; g <= 6; g++) {
                if (this.config.cavityGroups[g] && this.config.cavityGroups[g].cavityIdRange) {
                    const offset = this.config.cavityGroups[g].pageOffset || 0;
                    minOffset = Math.min(minOffset, offset);
                    maxOffset = Math.max(maxOffset, offset);
                    hasActiveGroups = true;

                    if (!refGroup || offset < refGroup.pageOffset) {
                        refGroup = this.config.cavityGroups[g];
                    }
                }
            }
        }

        if (!hasActiveGroups) {
            minOffset = 0;
            maxOffset = 0;
        }

        // Alignment logic: Find where the "Setup" sheet (reference sheet) is in THIS workbook
        let baseIndex = 0;
        if (refGroup && refGroup.sheetName) {
            const index = workbook.SheetNames.indexOf(refGroup.sheetName);
            if (index !== -1) {
                // If refGroup was at index 2 during setup, but is at index 0 here,
                // then baseIndex should be -2, such that baseIndex + offset(2) = 0.
                baseIndex = index - refGroup.pageOffset;
            }
        }

        const step = maxOffset - minOffset + 1;

        // 遍歷所有工作表
        for (let i = baseIndex; i < sheetCount; i += step) {
            if (i + minOffset < 0) continue;
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
                await this.processWorksheet(workbook, worksheet, sheetName, i, workbookId);
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
     * 獲取基準批號名稱（移除字尾，如 -1, -2, (1) 等）
     * @param {string} name 
     * @returns {string}
     */
    getBaseBatchName(name) {
        if (!name) return '';
        // 移除常見的字尾，如 -1, -2, (1), (2), _1, _2 等
        return name.replace(/[-_(\s](\d+)\)?$/g, '').trim();
    }

    /**
     * 處理單個工作表
     * @param {Object} worksheet 
     * @param {string} sheetName 
     * @param {number} sheetIndex 
     * @param {number} workbookId
     */
    async processWorksheet(workbook, worksheet, sheetName, sheetIndex, workbookId) {
        // 獲取批號（使用工作表名稱作為批號，並提取基準名稱以利合併）
        const batchName = this.getBaseBatchName(sheetName);
        // 用於彙整該工作表（批次）中所有穴組的數據
        // itemName -> merged data object
        const batchItemData = {};

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
                if (!batchItemData[item.inspectionItem]) {
                    batchItemData[item.inspectionItem] = {};
                }
                // 合併數據（如果穴號重複，後面的穴組會覆蓋前面的，通常穴組之間穴號不應重複）
                Object.assign(batchItemData[item.inspectionItem], item.data);
            }
        }

        // 將該批次合併後的數據一次性加入結果
        for (const [itemName, data] of Object.entries(batchItemData)) {
            this.addToResults(itemName, batchName, data, workbookId);
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
                // 注意：检验项目名称可以是數字（如 "1", "2"），必須作為文字處理
                let tempValue = safeGetMergedValue(dataRow, 0);  // A 列

                if (tempValue !== null && tempValue !== undefined) {
                    itemName = String(tempValue).trim();
                }

                // 如果 A 列没有，嘗試 B 列
                if (!itemName || itemName === '') {
                    tempValue = safeGetMergedValue(dataRow, 1);  // B 列
                    if (tempValue !== null && tempValue !== undefined) {
                        itemName = String(tempValue).trim();
                    }
                }

                if (itemName && itemName !== '') {
                    itemName = itemName.trim(); // Normalize to prevent duplicates due to spaces
                    console.log(`[QIP][Row${dataRow + 1}] ✓ 找到檢驗項目: "${itemName}"`);
                } else {
                    console.warn(`[QIP][Row${dataRow + 1}] ✗ A/B 列都沒有內容，跳過此行`);
                    continue;
                }


                // 提取穴號數據
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
     * @param {number} workbookId
     */
    addToResults(inspectionItem, batchName, data, workbookId = 0) {
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

        // 檢查是否在當前工作簿中已有同名的批次 (使用基準名稱)
        // 如果有，則合併數據（解決單個批次跨多個工作表的問題，例如 BatchA-1, BatchA-2）
        // 如果沒有，或者工作簿不同，則新增批次（解決不同檔案中同名批號如 Setup 的問題）
        const existingBatch = item.batches.find(b => b.name === batchName && b.workbookId === workbookId);

        if (existingBatch) {
            console.log(`[QIP] 合併批次數據 - 項目: ${inspectionItem}, 批次: ${batchName}, 原始穴數: ${Object.keys(existingBatch.data).length}, 新增穴數: ${cavityIds.length}`);
            Object.assign(existingBatch.data, data);
        } else {
            item.batches.push({
                name: batchName,
                data: { ...data },
                workbookId: workbookId
            });
            console.log(`[QIP] 新增批次 - 項目: ${inspectionItem}, 批次: ${batchName}, 穴號: [${cavityIds.join(', ')}]`);
        }

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
     * @param {Array|Object} workbooks - 單個或多個工作簿
     * @param {Function} progressCallback 
     */
    async extractSpecifications(workbooks, progressCallback) {
        progressCallback({
            current: 100,
            total: 100,
            message: '提取規格數據...',
            percent: 95
        });

        const wbList = Array.isArray(workbooks) ? workbooks : [workbooks];

        for (const itemName of Object.keys(this.results.inspectionItems)) {
            let foundSpec = null;
            // 遍歷所有工作簿，直到找到有效的規格
            for (const wb of wbList) {
                const spec = SpecificationExtractor.extractSpecification(wb, itemName);
                if (spec && spec.isValid) {
                    foundSpec = spec;
                    break;
                }
                // 如果沒找到有效的，先暫存第一個結果
                if (!foundSpec) foundSpec = spec;
            }
            this.results.inspectionItems[itemName].specification = foundSpec;
        }
    }

    /**
     * 提取產品資訊 (Product Name & Measurement Unit)
     * @param {Array|Object} workbooks - 單個或多個工作簿
     * @returns {Object}
     */
    extractProductInfo(workbooks) {
        let productName = '';
        let measurementUnit = '';

        const wbList = Array.isArray(workbooks) ? workbooks : [workbooks];

        for (const workbook of wbList) {
            for (const sheetName of workbook.SheetNames) {
                const ws = workbook.Sheets[sheetName];

                // 提取產品名稱 (P2 或 P3)
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

                // 提取測量單位 (W23)
                if (!measurementUnit) {
                    const cellW23 = ws['W23'];
                    if (cellW23 && cellW23.v) {
                        let val = String(cellW23.v).trim();
                        val = val.replace(/單位[:：]/g, '').trim();
                        measurementUnit = val;
                    }
                }

                if (productName && measurementUnit) break;
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

    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
}

// 導出供其他模組使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = QIPProcessor;
}
