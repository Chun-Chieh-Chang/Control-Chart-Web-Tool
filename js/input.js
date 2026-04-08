/**
 * DATA INPUT - Excel File Parsing and Preprocessing
 */

// Clean batch name by removing suffixes like "-1", "(2)", " (2)", etc.
function cleanBatchName(name) {
    if (!name) return '';
    return String(name).trim();
}

/**
 * DataInput Class
 * @param {Object} worksheet - SheetJS worksheet object
 */
function DataInput(worksheet) {
    this.ws = worksheet;
    this.data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    this.parse();
}

DataInput.prototype.parse = function () {
    var self = this;

    // 提取中繼資料 (Metadata)
    // 根據新格式：A5 為 'ProductName'，B5 為其值；A6 為 'MeasurementUnit'，B6 為其值
    var row5 = this.data[4] || [];
    var row6 = this.data[5] || [];

    if (row5[0] === 'ProductName' || row5[0] === '產品名稱') {
        this.productInfo = {
            name: row5[1] || '',
            item: '', // 品號通常在檔名
            unit: row6[1] || 'Inch',
            char: '平均值/全距',
            dept: '品管部',
            inspector: '品管組',
            batchRange: '',
            chartNo: ''
        };
    } else {
        // 回退機制 (Fallback)
        this.productInfo = {
            name: this.data[0] && this.data[0][1] ? this.data[0][1] : '',
            item: this.data[0] && this.data[0][2] ? this.data[0][2] : '',
            unit: 'Inch',
            char: '平均值/全距',
            dept: '品管部',
            inspector: '品管組',
            batchRange: '',
            chartNo: ''
        };
    }

    // 3. 提取規格資訊 (Specification Mapping Detection)
    // 優先檢查「舊格式」(Golden Format): A=Target, B=USL, C=LSL, D=生產批號
    var row1 = this.data[0] || [];
    var row2 = this.data[1] || [];
    var row1A = row1[0] ? row1[0].toString().trim() : '';
    var row1B = row1[1] ? row1[1].toString().trim() : '';
    var row1C = row1[2] ? row1[2].toString().trim() : '';
    var row1D = row1[3] ? row1[3].toString().trim() : '';

    var specIndices = { target: 0, usl: 1, lsl: 2 }; // 默認 A, B, C
    var isBenchmark = false;

    // 判斷是否為「標竿格式」(Benchmark): A=批號, B=USL, C=Target, D=LSL
    if (row1A === '生產批號' || row1A === 'BatchNo' || row1C === 'Target') {
        if (row1B === 'USL' && row1D === 'LSL') {
           isBenchmark = true;
           specIndices = { usl: 1, target: 2, lsl: 3 };
        }
    }

    // 依照偵測到的索引提取 (Using detected indices)
    this.specs = {
        target: parseFloat(row2[specIndices.target]) || 0,
        usl: parseFloat(row2[specIndices.usl]) || 0,
        lsl: parseFloat(row2[specIndices.lsl]) || 0
    };

    // 4. 判斷批號位置與數據起點 (Batch Column & Data Start Detection)
    var batchColIdx = isBenchmark ? 0 : 3; // 標竿 A (0), 舊格式 D (3)
    var dataStartIdx = 2; // 默認 Row 3 (有 Header)
    
    // 偵測數據列 (Cavity Columns)
    // 根據標竿格式：Column E (index 4) 之後為穴號量測值
    this.headers = this.data[0] || [];
    this.cavityColumns = [];
    for (var i = 4; i < this.headers.length; i++) {
        var header = this.headers[i];
        if (header && (header.toString().indexOf('穴') >= 0 || !isNaN(parseFloat(header)))) {
            this.cavityColumns.push({ index: i, name: header });
        }
    }
    
    // 精確偵測數據起點：檢查 Row 2 (index 1) 是否包含量測數據
    var row2HasData = false;
    if (this.cavityColumns.length > 0) {
        var firstCavIdx = this.cavityColumns[0].index;
        var firstVal = parseFloat(row2[firstCavIdx]);
        if (!isNaN(firstVal)) {
            row2HasData = true;
        }
    }
    
    if (row2HasData) {
        dataStartIdx = 1; // 從 Row 2 開始讀取
    }

    // 提取批號與同步數據行
    var rawRows = this.data.slice(dataStartIdx);
    this.dataRows = [];
    this.batchNames = [];
    for (var j = 0; j < rawRows.length; j++) {
        var name = rawRows[j][batchColIdx];
        if (name && name !== '') {
            this.batchNames.push(cleanBatchName(name));
            this.dataRows.push(rawRows[j]);
        }
    }
};

DataInput.prototype.getSpecs = function () { return this.specs; };
DataInput.prototype.getProductInfo = function () { return this.productInfo; };
DataInput.prototype.getCavityNames = function () { return this.cavityColumns.map(function (c) { return c.name; }); };
DataInput.prototype.getCavityCount = function () { return this.cavityColumns.length; };

DataInput.prototype.getDataMatrix = function () {
    var matrix = [];
    for (var i = 0; i < this.dataRows.length; i++) {
        var batchData = [];
        for (var j = 0; j < this.cavityColumns.length; j++) {
            var value = parseFloat(this.dataRows[i][this.cavityColumns[j].index]);
            batchData.push(isNaN(value) ? null : value);
        }
        matrix.push(batchData);
    }
    return matrix;
};

DataInput.prototype.getCavityBatchData = function (cavityIndex) {
    var column = this.cavityColumns[cavityIndex];
    if (!column) return [];
    var result = [];
    for (var i = 0; i < this.dataRows.length; i++) {
        var value = parseFloat(this.dataRows[i][column.index]);
        if (!isNaN(value)) result.push(value);
    }
    return result;
};
