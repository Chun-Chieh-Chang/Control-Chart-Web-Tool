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

    // Extract metadata from headers or fixed positions
    // Supports both old format (Row 1) and new VBA-compatible format (Row 5/6)
    var row5 = this.data[4] || [];
    var row6 = this.data[5] || [];

    if (row5[1] === 'ProductName') {
        // VBA-compatible format
        this.productInfo = {
            name: row6[1] || '',
            item: '', // Item code usually in filename
            unit: row6[2] || 'Inch',
            char: '平均值/全距',
            dept: '品管部',
            inspector: '品管組',
            batchRange: '',
            chartNo: ''
        };
    } else {
        // Old format / fallback
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

    this.headers = this.data[0] || [];
    this.specs = {
        target: parseFloat(this.data[1] && this.data[1][1]) || 0,
        usl: parseFloat(this.data[1] && this.data[1][2]) || 0,
        lsl: parseFloat(this.data[1] && this.data[1][3]) || 0
    };

    this.cavityColumns = [];
    for (var i = 0; i < this.headers.length; i++) {
        var header = this.headers[i];
        if (typeof header === 'string' && header.indexOf('穴') >= 0) {
            this.cavityColumns.push({ index: i, name: header });
        }
    }

    this.dataRows = this.data.slice(2);
    this.batchNames = [];
    for (var j = 0; j < this.dataRows.length; j++) {
        var name = this.dataRows[j][0];
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
