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

    // 提取規格 (Specifications)
    // 根據新格式：Row 2 (index 1) 的 A, B, C 欄為 Target, USL, LSL
    var row2 = this.data[1] || [];
    this.specs = {
        target: parseFloat(row2[0]) || 0,
        usl: parseFloat(row2[1]) || 0,
        lsl: parseFloat(row2[2]) || 0
    };

    // 偵測數據列 (Cavity Columns)
    // 根據新格式：Column E (index 4) 之後為穴號量測值
    this.headers = this.data[0] || [];
    this.cavityColumns = [];
    for (var i = 4; i < this.headers.length; i++) {
        var header = this.headers[i];
        if (typeof header === 'string' && (header.indexOf('穴') >= 0 || !isNaN(parseFloat(header)))) {
            this.cavityColumns.push({ index: i, name: header });
        }
    }

    // 提取批號與實際數據行
    // 根據新格式：從 Row 2 (index 1) 開始，批號在 D 欄 (index 3)
    this.dataRows = this.data.slice(1);
    this.batchNames = [];
    for (var j = 0; j < this.dataRows.length; j++) {
        var name = this.dataRows[j][3]; // D 欄
        if (name && name !== '') {
            this.batchNames.push(cleanBatchName(name));
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
