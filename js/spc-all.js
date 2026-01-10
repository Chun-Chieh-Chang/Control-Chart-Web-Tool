// ============================================================================
// SPC Analysis Tool - All-in-One JavaScript (No ES6 Modules)
// Using SheetJS for Excel generation (best compatibility)
// ============================================================================

// ============================================================================
// SPC ENGINE - Core Statistical Calculations
// ============================================================================

var SPCEngine = {
    SPC_CONSTANTS: {
        2: { A2: 1.88, D3: 0, D4: 3.267 },
        3: { A2: 1.023, D3: 0, D4: 2.373 },
        4: { A2: 0.729, D3: 0, D4: 2.282 },
        5: { A2: 0.577, D3: 0, D4: 2.115 },
        6: { A2: 0.483, D3: 0, D4: 2.004 },
        7: { A2: 0.419, D3: 0.076, D4: 1.924 },
        8: { A2: 0.373, D3: 0.136, D4: 1.864 },
        9: { A2: 0.337, D3: 0.184, D4: 1.816 },
        10: { A2: 0.308, D3: 0.223, D4: 1.777 },
        11: { A2: 0.285, D3: 0.256, D4: 1.744 },
        12: { A2: 0.266, D3: 0.283, D4: 1.717 },
        13: { A2: 0.249, D3: 0.307, D4: 1.693 },
        14: { A2: 0.235, D3: 0.328, D4: 1.672 },
        15: { A2: 0.223, D3: 0.347, D4: 1.653 },
        16: { A2: 0.212, D3: 0.363, D4: 1.637 },
        17: { A2: 0.203, D3: 0.378, D4: 1.622 },
        18: { A2: 0.194, D3: 0.391, D4: 1.608 },
        19: { A2: 0.187, D3: 0.403, D4: 1.597 },
        20: { A2: 0.18, D3: 0.415, D4: 1.585 },
        21: { A2: 0.173, D3: 0.425, D4: 1.575 },
        22: { A2: 0.167, D3: 0.434, D4: 1.566 },
        23: { A2: 0.162, D3: 0.443, D4: 1.557 },
        24: { A2: 0.157, D3: 0.451, D4: 1.548 },
        25: { A2: 0.153, D3: 0.459, D4: 1.541 }
    },

    getConstants: function (n) {
        if (n < 2) return { A2: 0, D3: 0, D4: 0 };
        if (n > 25) return { A2: 3 / Math.sqrt(n), D3: 0.5, D4: 1.5 };
        return this.SPC_CONSTANTS[n];
    },

    mean: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length === 0) return 0;
        var sum = 0;
        for (var i = 0; i < filtered.length; i++) sum += filtered[i];
        return sum / filtered.length;
    },

    min: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        return filtered.length > 0 ? Math.min.apply(null, filtered) : 0;
    },

    max: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        return filtered.length > 0 ? Math.max.apply(null, filtered) : 0;
    },

    range: function (data) {
        return this.max(data) - this.min(data);
    },

    stdDev: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length < 2) return 0;
        var avg = this.mean(filtered);
        var sumSq = 0;
        for (var i = 0; i < filtered.length; i++) {
            sumSq += Math.pow(filtered[i] - avg, 2);
        }
        return Math.sqrt(sumSq / (filtered.length - 1));
    },

    withinStdDev: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length < 2) return 0;
        var mrSum = 0;
        for (var i = 1; i < filtered.length; i++) {
            mrSum += Math.abs(filtered[i] - filtered[i - 1]);
        }
        return (mrSum / (filtered.length - 1)) / 1.128;
    },

    calculateXBarRLimits: function (dataMatrix) {
        var self = this;
        var n = dataMatrix[0] ? dataMatrix[0].length : 0;
        var constants = this.getConstants(n);
        var xBars = [];
        var ranges = [];

        for (var i = 0; i < dataMatrix.length; i++) {
            var subgroup = dataMatrix[i];
            var filtered = subgroup.filter(function (v) { return v !== null && !isNaN(v); });
            if (filtered.length > 0) {
                xBars.push(self.mean(filtered));
                ranges.push(self.range(filtered));
            }
        }

        var xDoubleBar = this.mean(xBars);
        var rBar = this.mean(ranges);

        var xUCL = xDoubleBar + constants.A2 * rBar;
        var xLCL = xDoubleBar - constants.A2 * rBar;
        var rUCL = constants.D4 * rBar;
        var rLCL = constants.D3 * rBar;

        var results = {
            xBar: {
                data: xBars,
                UCL: xUCL,
                CL: xDoubleBar,
                LCL: xLCL,
                sigma: (xUCL - xDoubleBar) / 3
            },
            R: {
                data: ranges,
                UCL: rUCL,
                CL: rBar,
                LCL: rLCL
            },
            summary: { n: n, k: xBars.length, xDoubleBar: xDoubleBar, rBar: rBar }
        };

        results.xBar.violations = this.checkNelsonRules(xBars, xDoubleBar, results.xBar.sigma);
        return results;
    },

    checkNelsonRules: function (data, cl, sigma) {
        var violations = [];
        if (data.length === 0 || sigma === 0) return violations;

        for (var i = 0; i < data.length; i++) {
            var rules = [];
            if (Math.abs(data[i] - cl) > 3 * sigma) rules.push(1);
            if (i >= 8) {
                var sameSide = true, side = data[i] > cl;
                for (var j = i - 8; j <= i; j++) { if ((data[j] > cl) !== side || data[j] === cl) { sameSide = false; break; } }
                if (sameSide) rules.push(2);
            }
            if (i >= 5) {
                var inc = true, dec = true;
                for (var j = i - 5; j < i; j++) { if (data[j + 1] <= data[j]) inc = false; if (data[j + 1] >= data[j]) dec = false; }
                if (inc || dec) rules.push(3);
            }
            if (i >= 13) {
                var isAlt = true;
                for (var j = i - 13; j < i; j++) { if ((data[j + 1] >= data[j] && data[j] >= data[j - 1]) || (data[j + 1] <= data[j] && data[j] <= data[j - 1])) { isAlt = false; break; } }
                if (isAlt) rules.push(4);
            }
            if (i >= 2) {
                var up = 0, lo = 0;
                for (var j = i - 2; j <= i; j++) { if (data[j] > cl + 2 * sigma) up++; if (data[j] < cl - 2 * sigma) lo++; }
                if (up >= 2 || lo >= 2) rules.push(5);
            }
            if (i >= 4) {
                var up = 0, lo = 0;
                for (var j = i - 4; j <= i; j++) { if (data[j] > cl + 1 * sigma) up++; if (data[j] < cl - 1 * sigma) lo++; }
                if (up >= 4 || lo >= 4) rules.push(6);
            }
            if (rules.length > 0) violations.push({ index: i, rules: rules });
        }
        return violations;
    },

    calculateProcessCapability: function (data, usl, lsl) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length < 2) return { Cp: 0, Cpk: 0, Pp: 0, Ppk: 0, mean: 0, withinStdDev: 0, overallStdDev: 0, count: 0 };

        var mean = this.mean(filtered);
        var withinStdDev = this.withinStdDev(filtered);
        var overallStdDev = this.stdDev(filtered);
        var tolerance = usl - lsl;

        var Cp = withinStdDev > 0 ? tolerance / (6 * withinStdDev) : 0;
        var Cpu = withinStdDev > 0 ? (usl - mean) / (3 * withinStdDev) : 0;
        var Cpl = withinStdDev > 0 ? (mean - lsl) / (3 * withinStdDev) : 0;
        var Cpk = Math.min(Cpu, Cpl);

        var Pp = overallStdDev > 0 ? tolerance / (6 * overallStdDev) : 0;
        var Ppu = overallStdDev > 0 ? (usl - mean) / (3 * overallStdDev) : 0;
        var Ppl = overallStdDev > 0 ? (mean - lsl) / (3 * overallStdDev) : 0;
        var Ppk = Math.min(Ppu, Ppl);

        return {
            Cp: Cp, Cpk: Cpk, Pp: Pp, Ppk: Ppk,
            mean: mean, withinStdDev: withinStdDev, overallStdDev: overallStdDev,
            min: this.min(filtered), max: this.max(filtered),
            range: this.range(filtered), count: filtered.length
        };
    },

    getCapabilityColor: function (cpk) {
        if (cpk >= 1.67) return { bg: '#c6efce', text: '#006100' };
        if (cpk >= 1.33) return { bg: '#c6efce', text: '#006100' };
        if (cpk >= 1.0) return { bg: '#ffeb9c', text: '#9c5700' };
        return { bg: '#ffc7ce', text: '#9c0006' };
    },

    round: function (value, decimals) {
        decimals = decimals || 4;
        return Math.round(value * Math.pow(10, decimals)) / Math.pow(10, decimals);
    }
};

// ============================================================================
// DATA INPUT - Excel File Parsing
// ============================================================================

// Clean batch name by removing suffixes like "-1", "(2)", " (2)", etc.
function cleanBatchName(name) {
    if (!name) return '';
    var str = String(name).trim();
    // Remove patterns like: -1, -2, (1), (2), _1, _2, etc.
    // Pattern: trailing dash/underscore + number, or parentheses with numbers
    str = str.replace(/[\-_]\d+$/, '');        // Remove -1, -2, _1, _2 at end
    str = str.replace(/\s*\(\d+\)$/, '');      // Remove (1), (2), " (1)" at end
    str = str.replace(/\s*\[\d+\]$/, '');      // Remove [1], [2] at end
    return str.trim();
}

function DataInput(worksheet) {
    this.ws = worksheet;
    this.data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    this.parse();
}

DataInput.prototype.parse = function () {
    var self = this;

    // Extract metadata from headers or fixed positions
    this.productInfo = {
        name: this.data[0] && this.data[0][1] ? this.data[0][1] : '',
        item: this.data[0] && this.data[0][2] ? this.data[0][2] : '',
        unit: 'Inch', // Default or extracted
        char: '平均值/全距',
        dept: '品管部',
        inspector: '品管組',
        batchRange: '',
        chartNo: ''
    };

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

DataInput.prototype.getDataMatrix = function () {
    var self = this;
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

DataInput.prototype.getCavityCount = function () { return this.cavityColumns.length; };
DataInput.prototype.getCavityNames = function () {
    var names = [];
    for (var i = 0; i < this.cavityColumns.length; i++) {
        names.push(this.cavityColumns[i].name);
    }
    return names;
};
DataInput.prototype.getSpecs = function () { return this.specs; };

// ============================================================================
// MAIN APPLICATION
// ============================================================================

var SPCApp = {
    currentLanguage: 'zh',
    workbook: null,
    selectedItem: null,
    analysisResults: null,
    chartInstances: [],

    init: function () {
        this.setupLanguageToggle();
        this.setupFileUpload();
        this.setupEventListeners();
        this.updateLanguage();
        console.log('SPC Analysis Tool initialized');
    },

    t: function (zh, en) {
        return this.currentLanguage === 'zh' ? zh : en;
    },

    setupLanguageToggle: function () {
        var self = this;
        document.getElementById('langBtn').addEventListener('click', function () {
            self.currentLanguage = self.currentLanguage === 'zh' ? 'en' : 'zh';
            document.getElementById('langText').textContent = self.currentLanguage === 'zh' ? 'EN' : '中文';
            self.updateLanguage();
        });
    },

    updateLanguage: function () {
        var self = this;
        var elements = document.querySelectorAll('[data-en][data-zh]');
        for (var i = 0; i < elements.length; i++) {
            var el = elements[i];
            el.textContent = self.currentLanguage === 'zh' ? el.dataset.zh : el.dataset.en;
        }
    },

    setupFileUpload: function () {
        var self = this;
        var uploadZone = document.getElementById('uploadZone');
        var fileInput = document.getElementById('fileInput');

        uploadZone.addEventListener('click', function () { fileInput.click(); });
        uploadZone.addEventListener('dragover', function (e) { e.preventDefault(); uploadZone.classList.add('dragover'); });
        uploadZone.addEventListener('dragleave', function () { uploadZone.classList.remove('dragover'); });
        uploadZone.addEventListener('drop', function (e) {
            e.preventDefault();
            uploadZone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) self.handleFile(e.dataTransfer.files[0]);
        });
        fileInput.addEventListener('change', function (e) {
            if (e.target.files.length > 0) self.handleFile(e.target.files[0]);
        });
        document.getElementById('resetBtn').addEventListener('click', function () { self.resetApp(); });
    },

    handleFile: function (file) {
        var self = this;
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            alert(this.t('請選擇 Excel 檔案', 'Please select an Excel file'));
            return;
        }
        this.showLoading(this.t('讀取檔案中...', 'Reading file...'));

        var reader = new FileReader();
        reader.onload = function (e) {
            try {
                var data = new Uint8Array(e.target.result);
                self.workbook = XLSX.read(data, { type: 'array' });
                document.getElementById('uploadZone').style.display = 'none';
                document.getElementById('fileInfo').style.display = 'flex';
                document.getElementById('fileName').textContent = file.name;
                self.showInspectionItems();
                self.hideLoading();
            } catch (error) {
                self.hideLoading();
                alert(self.t('檔案讀取失敗', 'File reading failed') + ': ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    },

    showInspectionItems: function () {
        var self = this;
        var itemList = document.getElementById('itemList');
        itemList.innerHTML = '';

        var sheets = this.workbook.SheetNames;
        for (var i = 0; i < sheets.length; i++) {
            var sheetName = sheets[i];
            var ws = self.workbook.Sheets[sheetName];
            var data = XLSX.utils.sheet_to_json(ws, { header: 1 });
            var target = data[1] && data[1][1] ? data[1][1] : 'N/A';

            var card = document.createElement('div');
            card.className = 'item-card';
            card.innerHTML = '<div class="item-info"><h3>' + sheetName + '</h3><p>Target: ' + target + '</p></div><div class="item-badge">' + self.t('選擇', 'Select') + '</div>';
            card.dataset.sheet = sheetName;
            card.addEventListener('click', function () {
                self.selectedItem = this.dataset.sheet;
                self.showAnalysisOptions();
            });
            itemList.appendChild(card);
        }

        document.getElementById('step2').style.display = 'block';
        document.getElementById('step2').scrollIntoView({ behavior: 'smooth' });
    },

    showAnalysisOptions: function () {
        document.getElementById('step3').style.display = 'block';
        document.getElementById('step3').scrollIntoView({ behavior: 'smooth' });
    },

    setupEventListeners: function () {
        var self = this;
        var buttons = document.querySelectorAll('[data-analysis]');
        for (var i = 0; i < buttons.length; i++) {
            buttons[i].addEventListener('click', function () {
                self.executeAnalysis(this.dataset.analysis);
            });
        }
        document.getElementById('downloadExcel').addEventListener('click', function () { self.downloadExcel(); });
    },

    executeAnalysis: function (type) {
        var self = this;
        this.showLoading(this.t('分析中...', 'Analyzing...'));

        for (var i = 0; i < this.chartInstances.length; i++) {
            this.chartInstances[i].destroy();
        }
        this.chartInstances = [];

        setTimeout(function () {
            try {
                var ws = self.workbook.Sheets[self.selectedItem];
                var dataInput = new DataInput(ws);
                var results;

                if (type === 'batch') {
                    var dataMatrix = dataInput.getDataMatrix();
                    var xbarR = SPCEngine.calculateXBarRLimits(dataMatrix);
                    results = {
                        type: 'batch',
                        xbarR: xbarR,
                        batchNames: dataInput.batchNames,
                        specs: dataInput.getSpecs(),
                        dataMatrix: dataMatrix,
                        cavityNames: dataInput.getCavityNames(),
                        productInfo: dataInput.getProductInfo()
                    };
                } else if (type === 'cavity') {
                    var specs = dataInput.getSpecs();
                    var cavityStats = [];
                    for (var i = 0; i < dataInput.getCavityCount(); i++) {
                        var cavData = dataInput.getCavityBatchData(i);
                        var cap = SPCEngine.calculateProcessCapability(cavData, specs.usl, specs.lsl);
                        cap.name = dataInput.getCavityNames()[i];
                        cavityStats.push(cap);
                    }
                    results = { type: 'cavity', cavityStats: cavityStats, specs: specs };
                } else if (type === 'group') {
                    var specs = dataInput.getSpecs();
                    var dataMatrix = dataInput.getDataMatrix();
                    var groupStats = [];
                    for (var i = 0; i < dataMatrix.length; i++) {
                        var filtered = dataMatrix[i].filter(function (v) { return v !== null && !isNaN(v); });
                        groupStats.push({
                            batch: dataInput.batchNames[i] || 'Batch ' + (i + 1),
                            avg: SPCEngine.mean(filtered),
                            max: SPCEngine.max(filtered),
                            min: SPCEngine.min(filtered),
                            range: SPCEngine.range(filtered),
                            count: filtered.length
                        });
                    }
                    results = { type: 'group', groupStats: groupStats, specs: specs };
                }

                self.analysisResults = results;
                self.displayResults();
                self.hideLoading();
            } catch (error) {
                self.hideLoading();
                alert(self.t('分析失敗', 'Analysis failed') + ': ' + error.message);
                console.error(error);
            }
        }, 100);
    },

    displayResults: function () {
        var resultsContent = document.getElementById('resultsContent');
        var html = '';
        var data = this.analysisResults;
        var self = this;

        if (data.type === 'batch') {
            var totalBatches = Math.min(data.batchNames.length, data.xbarR.xBar.data.length);
            var maxPerPage = 25;
            var totalPages = Math.ceil(totalBatches / maxPerPage);

            this.batchPagination = {
                currentPage: 1,
                totalPages: totalPages,
                maxPerPage: maxPerPage,
                totalBatches: totalBatches
            };

            html = '<div class="results-summary">' +
                '<div class="stat-card"><div class="stat-label">' + this.t('模穴數 (n)', 'Cavity Count') + '</div><div class="stat-value">' + data.xbarR.summary.n + '</div></div>' +
                '<div class="stat-card"><div class="stat-label">' + this.t('總批號數', 'Total Batches') + '</div><div class="stat-value">' + totalBatches + '</div></div>' +
                '</div>';

            if (totalPages > 1) {
                html += '<div class="pagination-controls" style="display:flex;justify-content:center;align-items:center;gap:15px;margin:20px 0;">' +
                    '<button id="prevPageBtn" class="btn-secondary" style="padding:8px 16px;">' + this.t('上一頁', 'Prev') + '</button>' +
                    '<span id="pageInfo" style="font-weight:bold;">' + this.t('第 ', 'Page ') + '1 / ' + totalPages + this.t(' 頁', '') + '</span>' +
                    '<button id="nextPageBtn" class="btn-secondary" style="padding:8px 16px;">' + this.t('下一頁', 'Next') + '</button>' +
                    '</div>';
            }

            // Detailed Data Table Container (Excel Style)
            html += '<div id="detailedTableContainer" style="margin-bottom:30px; overflow-x:auto;"></div>';

            html += '<div id="pageLimitsContainer"></div>';
            html += '<div id="diagnosticContainer" style="margin-top:20px;"></div>';

            html += '<div id="batchChartsContainer">' +
                '<div class="chart-container"><h3 class="chart-title">' + this.t('X̄ 管制圖', 'X-Bar Chart') + '</h3><canvas id="xbarChart"></canvas></div>' +
                '<div class="chart-container"><h3 class="chart-title">' + this.t('R 管制圖', 'R Chart') + '</h3><canvas id="rChart"></canvas></div>' +
                '</div>';

        } else if (data.type === 'cavity') {
            var rows = '';
            for (var i = 0; i < data.cavityStats.length; i++) {
                var s = data.cavityStats[i];
                var c = SPCEngine.getCapabilityColor(s.Cpk);
                rows += '<tr><td>' + s.name + '</td><td>' + SPCEngine.round(s.mean, 4) + '</td><td>' + SPCEngine.round(s.withinStdDev, 4) + '</td><td>' + SPCEngine.round(s.overallStdDev, 4) + '</td><td>' + SPCEngine.round(s.Cp, 3) + '</td><td style="background:' + c.bg + ';color:' + c.text + ';font-weight:bold;">' + SPCEngine.round(s.Cpk, 3) + '</td><td>' + SPCEngine.round(s.Pp, 3) + '</td><td>' + SPCEngine.round(s.Ppk, 3) + '</td><td>' + s.count + '</td></tr>';
            }
            html = '<div class="chart-container"><h3 class="chart-title">' + this.t('Cpk 比較', 'Cpk Comparison') + '</h3><canvas id="cpkChart"></canvas></div>' +
                '<h3 class="chart-title">' + this.t('模穴統計', 'Cavity Statistics') + '</h3><div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>Cavity</th><th>Mean</th><th>σ_within</th><th>σ_overall</th><th>Cp</th><th>Cpk</th><th>Pp</th><th>Ppk</th><th>n</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
        } else if (data.type === 'group') {
            var rows = '';
            for (var i = 0; i < data.groupStats.length; i++) {
                var s = data.groupStats[i];
                rows += '<tr><td>' + s.batch + '</td><td>' + SPCEngine.round(s.avg, 4) + '</td><td>' + SPCEngine.round(s.max, 4) + '</td><td>' + SPCEngine.round(s.min, 4) + '</td><td>' + SPCEngine.round(s.range, 4) + '</td><td>' + s.count + '</td></tr>';
            }
            html = '<div class="chart-container"><h3 class="chart-title">' + this.t('Min-Max-Avg 圖', 'Min-Max-Avg Chart') + '</h3><canvas id="groupChart"></canvas></div>' +
                '<h3 class="chart-title">' + this.t('群組統計', 'Group Statistics') + '</h3><div style="overflow-x:auto;"><table class="data-table"><thead><tr><th>Batch</th><th>Avg</th><th>Max</th><th>Min</th><th>Range</th><th>n</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
        }

        resultsContent.innerHTML = html;
        document.getElementById('results').style.display = 'block';

        // Setup pagination event listeners
        if (data.type === 'batch' && this.batchPagination.totalPages > 1) {
            document.getElementById('prevPageBtn').addEventListener('click', function () { self.changeBatchPage(-1); });
            document.getElementById('nextPageBtn').addEventListener('click', function () { self.changeBatchPage(1); });
            this.updatePaginationButtons();
        }

        setTimeout(function () { self.renderCharts(); document.getElementById('results').scrollIntoView({ behavior: 'smooth' }); }, 100);
    },

    changeBatchPage: function (delta) {
        var p = this.batchPagination;
        var newPage = p.currentPage + delta;
        if (newPage >= 1 && newPage <= p.totalPages) {
            p.currentPage = newPage;
            this.updatePaginationButtons();
            this.renderCharts();
        }
    },

    updatePaginationButtons: function () {
        var p = this.batchPagination;
        document.getElementById('pageInfo').textContent = this.t('第 ', 'Page ') + p.currentPage + ' / ' + p.totalPages + this.t(' 頁', '');
        document.getElementById('prevPageBtn').disabled = (p.currentPage <= 1);
        document.getElementById('nextPageBtn').disabled = (p.currentPage >= p.totalPages);
    },

    renderDetailedDataTable: function (pageLabels, pageDataMatrix, pageXbarR) {
        var data = this.analysisResults;
        var info = data.productInfo;
        var specs = data.specs;
        var cavityCount = data.xbarR.summary.n;

        var colWidths = {
            label: 60,
            batch: 60,     // 生產批號欄位寬度
            summary: 30    // 彙總基礎欄寬 (總寬度 = 30 * 4 = 120px)
        };
        // Calculate total width explicitly to force horizontal scrolling
        // Total = Label(1) + Batch(25) + Summary(4)
        var totalWidth = colWidths.label + (25 * colWidths.batch) + (4 * colWidths.summary);

        var html = '<table class="excel-table" style="width:' + totalWidth + 'px; border-collapse:collapse; font-size:10px; font-family:sans-serif; border:2px solid #000; table-layout:fixed;">';

        html += '<colgroup>';
        html += '<col style="width:' + colWidths.label + 'px;">';
        for (var c = 0; c < 25; c++) html += '<col style="width:' + colWidths.batch + 'px;">';
        html += '<col style="width:' + colWidths.summary + 'px;"><col style="width:' + colWidths.summary + 'px;"><col style="width:' + colWidths.summary + 'px;"><col style="width:' + colWidths.summary + 'px;">';
        html += '</colgroup>';


        // --- Row 1: Header ---
        html += '<tr style="background:#f3f4f6;"><td colspan="30" style="border:1px solid #000; text-align:center; font-weight:bold; font-size:14px; padding:3px;">X̄ - R 管制圖</td></tr>';

        // --- Row 2-5: Metadata & Limits (Re-distributed colspans out of 30) ---
        var rows = [
            { l1: '商品名稱', v1: info.name, l2: '規格', v2: '標準', l3: '管制圖', v3: 'X̄', v4: 'R', l4: '製造部門', v4_val: info.dept },
            { l1: '商品料號', v1: info.item, l2: '最大值', v2: SPCEngine.round(specs.usl, 4), l3: '上限', v3: SPCEngine.round(pageXbarR.xBar.UCL, 4), v4: SPCEngine.round(pageXbarR.R.UCL, 4), l4: '檢驗人員', v4_val: info.inspector },
            { l1: '測量單位', v1: info.unit, l2: '目標值', v2: SPCEngine.round(specs.target, 4), l3: '中心值', v3: SPCEngine.round(pageXbarR.xBar.CL, 4), v4: SPCEngine.round(pageXbarR.R.CL, 4), l4: '管制特性', v4_val: info.char },
            { l1: '檢驗日期', v1: info.batchRange || '-', l2: '最小值', v2: SPCEngine.round(specs.lsl, 4), l3: '下限', v3: SPCEngine.round(pageXbarR.xBar.LCL, 4), v4: '-', l4: '圖表編號', v4_val: info.chartNo || '-' }
        ];

        rows.forEach(function (r) {
            html += '<tr>' +
                '<td colspan="2" style="border:1px solid #000; padding:1px 1px; font-weight:bold; background:#f9fafb;">' + r.l1 + '</td>' +
                '<td colspan="7" style="border:1px solid #000; padding:1px 1px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + r.v1 + '</td>' +
                '<td colspan="2" style="border:1px solid #000; padding:1px 1px; font-weight:bold; background:#f9fafb;">' + r.l2 + '</td>' +
                '<td colspan="2" style="border:1px solid #000; padding:1px 1px;">' + r.v2 + '</td>' +
                '<td colspan="2" style="border:1px solid #000; padding:1px 1px; font-weight:bold; background:#f9fafb;">' + r.l3 + '</td>' +
                '<td colspan="3" style="border:1px solid #000; padding:1px 1px;">' + r.v3 + '</td>' +
                '<td colspan="3" style="border:1px solid #000; padding:1px 1px;">' + (r.v4 || '') + '</td>' +
                '<td colspan="2" style="border:1px solid #000; padding:1px 1px; font-weight:bold; background:#f9fafb;">' + (r.l4 || '') + '</td>' +
                '<td colspan="7" style="border:1px solid #000; padding:1px 1px;">' + (r.v4_val || '') + '</td>' +
                '</tr>';
        });

        // --- Data Header: Batch Names ---
        html += '<tr style="background:#e5e7eb; font-weight:bold;">' +
            '<td style="border:1px solid #000; text-align:center;">批號</td>';
        pageLabels.forEach(function (name) {
            html += '<td style="border:1px solid #000; text-align:center; height:35px; overflow:hidden; font-size:10px; white-space:nowrap; text-overflow:ellipsis;">' + name + '</td>';
        });
        // Fill empty if less than 25 batches
        for (var f = pageLabels.length; f < 25; f++) html += '<td style="border:1px solid #000;"></td>';

        html += '<td colspan="4" style="border:1px solid #000; text-align:center;">彙總</td></tr>';

        // --- Main Data Rows: Cavities ---
        for (var i = 0; i < cavityCount; i++) {
            html += '<tr><td style="border:1px solid #000; text-align:center; font-weight:bold;">X' + (i + 1) + '</td>';
            for (var j = 0; j < 25; j++) {
                var val = (pageDataMatrix[j] && pageDataMatrix[j][i] !== undefined) ? pageDataMatrix[j][i] : null;
                html += '<td style="border:1px solid #000; text-align:center;">' + (val !== null ? val : '') + '</td>';
            }

            // Sidebar summary on first few rows
            if (i === 0) html += '<td colspan="4" rowspan="2" style="border:1px solid #000; text-align:center; font-weight:bold; background:#fefefe; font-size:10px;">ΣX̄ = ' + SPCEngine.round(pageXbarR.summary.xBarSum, 4) + '</td>';
            else if (i === 2) html += '<td colspan="4" rowspan="2" style="border:1px solid #000; text-align:center; font-weight:bold; background:#fefefe; font-size:10px;">X̿ = ' + SPCEngine.round(pageXbarR.summary.xDoubleBar, 4) + '</td>';
            else if (i === 4) html += '<td colspan="4" rowspan="2" style="border:1px solid #000; text-align:center; font-weight:bold; background:#fefefe; font-size:10px;">ΣR = ' + SPCEngine.round(pageXbarR.summary.rSum, 4) + '</td>';
            else if (i === 6) html += '<td colspan="4" rowspan="2" style="border:1px solid #000; text-align:center; font-weight:bold; background:#fefefe; font-size:10px;">R̄ = ' + SPCEngine.round(pageXbarR.summary.rBar, 4) + '</td>';
            else if (i >= 8) html += '<td colspan="4" style="border:1px solid #000; background:#fcfcfc;"></td>';
            html += '</tr>';
        }

        // --- Footer Rows: ΣX, X̄, R ---
        // ΣX Row
        html += '<tr style="background:#f9fafb;"><td style="border:1px solid #000; text-align:center; font-weight:bold;">ΣX</td>';
        for (var b = 0; b < 25; b++) {
            var val = '';
            if (pageDataMatrix[b]) {
                var sum = pageDataMatrix[b].reduce(function (a, b) { return a + (b || 0); }, 0);
                val = SPCEngine.round(sum, 4);
            }
            html += '<td style="border:1px solid #000; text-align:center;">' + val + '</td>';
        }
        html += '<td colspan="4" style="border:1px solid #000; background:#f9fafb;"></td></tr>';

        // X̄ Row (with yellow highlighting)
        html += '<tr style="background:#f9fafb;"><td style="border:1px solid #000; text-align:center; font-weight:bold;">X̄</td>';
        for (var k = 0; k < 25; k++) {
            var val = '', style = '';
            if (pageXbarR.xBar.data[k] !== undefined) {
                var v = pageXbarR.xBar.data[k];
                val = SPCEngine.round(v, 4);
                if (v > pageXbarR.xBar.UCL || v < pageXbarR.xBar.LCL) style = 'background:yellow;';
            }
            html += '<td style="border:1px solid #000; text-align:center;' + style + '">' + val + '</td>';
        }
        html += '<td colspan="4" style="border:1px solid #000; background:#f9fafb;"></td></tr>';

        // R Row
        html += '<tr style="background:#f9fafb;"><td style="border:1px solid #000; text-align:center; font-weight:bold;">R</td>';
        for (var k = 0; k < 25; k++) {
            var val = '', style = '';
            if (pageXbarR.R.data[k] !== undefined) {
                var v = pageXbarR.R.data[k];
                val = SPCEngine.round(v, 4);
                if (v > pageXbarR.R.UCL) style = 'background:yellow;';
            }
            html += '<td style="border:1px solid #000; text-align:center;' + style + '">' + val + '</td>';
        }
        html += '<td colspan="4" style="border:1px solid #000; background:#f9fafb;"></td></tr>';

        html += '</table>';
        document.getElementById('detailedTableContainer').innerHTML = html;
    },




    renderCharts: function () {
        var data = this.analysisResults;

        // Destroy existing charts
        for (var i = 0; i < this.chartInstances.length; i++) {
            this.chartInstances[i].destroy();
        }
        this.chartInstances = [];

        if (data.type === 'batch') {
            // Get pagination info
            var p = this.batchPagination || { currentPage: 1, maxPerPage: 25, totalBatches: data.batchNames.length };
            var startIdx = (p.currentPage - 1) * p.maxPerPage;
            var endIdx = Math.min(startIdx + p.maxPerPage, p.totalBatches);

            // Slice data for current page
            var pageLabels = data.batchNames.slice(startIdx, endIdx);
            var pageDataMatrix = data.dataMatrix.slice(startIdx, endIdx);

            // Calculate control limits for this page's data (VBA style - each page has its own limits)
            var pageXbarR = SPCEngine.calculateXBarRLimits(pageDataMatrix);

            // Add sums for detailed table display
            pageXbarR.summary.xBarSum = pageXbarR.xBar.data.reduce(function (a, b) { return a + b; }, 0);
            pageXbarR.summary.rSum = pageXbarR.R.data.reduce(function (a, b) { return a + b; }, 0);

            // Render detailed data table
            this.renderDetailedDataTable(pageLabels, pageDataMatrix, pageXbarR);

            // Update page limits display
            var limitsHtml = '<div class="results-summary" style="margin-top:15px;">' +
                '<div class="stat-card"><div class="stat-label">' + this.t('本頁批號數 (k)', 'Page Batches') + '</div><div class="stat-value">' + pageXbarR.summary.k + '</div></div>' +
                '<div class="stat-card"><div class="stat-label">X̿</div><div class="stat-value">' + SPCEngine.round(pageXbarR.summary.xDoubleBar, 4) + '</div></div>' +
                '<div class="stat-card"><div class="stat-label">R̄</div><div class="stat-value">' + SPCEngine.round(pageXbarR.summary.rBar, 4) + '</div></div>' +
                '</div>' +
                '<h3 class="chart-title">' + this.t('管制界限 (本頁)', 'Control Limits (This Page)') + '</h3>' +
                '<table class="data-table"><thead><tr><th>Chart</th><th>UCL</th><th>CL</th><th>LCL</th></tr></thead><tbody>' +
                '<tr><td>X̄</td><td>' + SPCEngine.round(pageXbarR.xBar.UCL, 4) + '</td><td>' + SPCEngine.round(pageXbarR.xBar.CL, 4) + '</td><td>' + SPCEngine.round(pageXbarR.xBar.LCL, 4) + '</td></tr>' +
                '<tr><td>R</td><td>' + SPCEngine.round(pageXbarR.R.UCL, 4) + '</td><td>' + SPCEngine.round(pageXbarR.R.CL, 4) + '</td><td>' + SPCEngine.round(pageXbarR.R.LCL, 4) + '</td></tr></tbody></table>';
            document.getElementById('pageLimitsContainer').innerHTML = limitsHtml;

            var diagnosticHtml = '<div class="diagnostic-panel" style="background:#fff; border-radius:8px; padding:15px; margin-top:20px; border:1px solid #e2e8f0;">' +
                '<h3 style="margin-top:0; color:#1e293b; border-bottom:2px solid #3b82f6; padding-bottom:8px; margin-bottom:12px;">' +
                this.t('異常診斷 (Nelson Rules)', 'Abnormality Diagnostic') + '</h3>';

            if (pageXbarR.xBar.violations.length === 0) {
                diagnosticHtml += '<p style="color:#10b981; font-weight:bold;">✅ ' + this.t('本頁數據未發現明顯異常趨勢。', 'No obvious abnormal trends found.') + '</p>';
            } else {
                diagnosticHtml += '<ul style="padding-left:20px; color:#475569;">';
                var ruleDescs = {
                    1: this.t('法則 1: 超出管制界限 (3σ)', 'Rule 1: Outside 3σ'),
                    2: this.t('法則 2: 連續 9 點在中心線同一側', 'Rule 2: 9 pts on one side'),
                    3: this.t('法則 3: 連續 6 點持續上升或下降', 'Rule 3: 6 pts trending'),
                    4: this.t('法則 4: 連續 14 點交互升降', 'Rule 4: 14 pts alternating'),
                    5: this.t('法則 5: 3 點中有 2 點在 2σ 外', 'Rule 5: 2/3 outside 2σ'),
                    6: this.t('法則 6: 5 點中有 4 點在 1σ 外', 'Rule 6: 4/5 outside 1σ')
                };
                pageXbarR.xBar.violations.forEach(function (v) {
                    diagnosticHtml += '<li style="margin-bottom:8px;"><strong style="color:#ef4444;">[' + pageLabels[v.index] + ']</strong>: ' +
                        v.rules.map(function (r) { return ruleDescs[r]; }).join(', ') + '</li>';
                });
                diagnosticHtml += '</ul>';
            }
            diagnosticHtml += '</div>';

            // Ensure diagnosticContainer exists or create it
            var diagDiv = document.getElementById('diagnosticContainer');
            if (!diagDiv) {
                diagDiv = document.createElement('div');
                diagDiv.id = 'diagnosticContainer';
                document.getElementById('batchChartsContainer').parentNode.insertBefore(diagDiv, document.getElementById('batchChartsContainer'));
            }
            diagDiv.innerHTML = diagnosticHtml;

            var fillUCL = [], fillCL = [], fillLCL = [];
            var xbarColors = [], rColors = [];

            var vMap = {};
            pageXbarR.xBar.violations.forEach(function (v) { vMap[v.index] = true; });

            for (var i = 0; i < pageLabels.length; i++) {
                fillUCL.push(pageXbarR.xBar.UCL);
                fillCL.push(pageXbarR.xBar.CL);
                fillLCL.push(pageXbarR.xBar.LCL);

                if (vMap[i]) {
                    xbarColors.push('#ef4444');
                } else {
                    xbarColors.push('#3b82f6');
                }

                var rVal = pageXbarR.R.data[i];
                if (rVal > pageXbarR.R.UCL) {
                    rColors.push('#ef4444');
                } else {
                    rColors.push('#3b82f6');
                }
            }

            this.chartInstances.push(new Chart(document.getElementById('xbarChart'), {
                type: 'line',
                data: {
                    labels: pageLabels, datasets: [
                        {
                            label: 'X̄',
                            data: pageXbarR.xBar.data,
                            borderColor: '#3b82f6',
                            borderWidth: 2,
                            pointBackgroundColor: xbarColors,
                            pointBorderColor: xbarColors,
                            pointRadius: 4,
                            pointHoverRadius: 6,
                            tension: 0,
                            fill: false
                        },
                        { label: 'UCL', data: fillUCL, borderColor: '#ef4444', borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, fill: false },
                        { label: 'CL', data: fillCL, borderColor: '#10b981', borderWidth: 1.5, pointRadius: 0, fill: false },
                        { label: 'LCL', data: fillLCL, borderColor: '#ef4444', borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, fill: false }
                    ]
                },
                options: {
                    responsive: true,
                    aspectRatio: 4, // More compact vertical
                    plugins: {
                        legend: { position: 'top', labels: { boxWidth: 10, padding: 10, font: { size: 12 } } }
                    },
                    scales: {
                        x: { grid: { display: false } },
                        y: { grid: { color: 'rgba(0,0,0,0.08)' } }
                    }
                }
            }));

            var rUCL = [], rCL = [];
            for (var i = 0; i < pageLabels.length; i++) { rUCL.push(pageXbarR.R.UCL); rCL.push(pageXbarR.R.CL); }
            this.chartInstances.push(new Chart(document.getElementById('rChart'), {
                type: 'line',
                data: {
                    labels: pageLabels, datasets: [
                        {
                            label: 'R',
                            data: pageXbarR.R.data,
                            borderColor: '#3b82f6',
                            borderWidth: 2,
                            pointBackgroundColor: rColors,
                            pointBorderColor: rColors,
                            pointRadius: 4,
                            pointHoverRadius: 6,
                            tension: 0,
                            fill: false
                        },
                        { label: 'UCL', data: rUCL, borderColor: '#ef4444', borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, fill: false },
                        { label: 'CL', data: rCL, borderColor: '#10b981', borderWidth: 1.5, pointRadius: 0, fill: false }
                    ]
                },
                options: {
                    responsive: true,
                    aspectRatio: 4,
                    plugins: {
                        legend: { position: 'top', labels: { boxWidth: 10, padding: 10, font: { size: 12 } } }
                    },
                    scales: {
                        x: { grid: { display: false } },
                        y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.08)' } }
                    }
                }
            }));
        } else if (data.type === 'cavity') {
            var labels = [], cpkValues = [], colors = [];
            for (var i = 0; i < data.cavityStats.length; i++) {
                labels.push(data.cavityStats[i].name);
                cpkValues.push(data.cavityStats[i].Cpk);
                colors.push(SPCEngine.getCapabilityColor(data.cavityStats[i].Cpk).bg);
            }
            var line167 = [], line133 = [], line100 = [];
            for (var i = 0; i < labels.length; i++) { line167.push(1.67); line133.push(1.33); line100.push(1.0); }

            this.chartInstances.push(new Chart(document.getElementById('cpkChart'), {
                type: 'bar',
                data: {
                    labels: labels, datasets: [
                        { label: 'Cpk', data: cpkValues, backgroundColor: colors, borderWidth: 0, barPercentage: 0.7 },
                        { label: '1.67', data: line167, type: 'line', borderColor: '#10b981', borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, fill: false },
                        { label: '1.33', data: line133, type: 'line', borderColor: '#f59e0b', borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, fill: false },
                        { label: '1.0', data: line100, type: 'line', borderColor: '#ef4444', borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, fill: false }
                    ]
                },
                options: {
                    responsive: true,
                    aspectRatio: 3.5,
                    plugins: { legend: { position: 'top', labels: { boxWidth: 12, padding: 20, font: { size: 14 } } } },
                    scales: { x: { grid: { display: false } }, y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.08)' } } }
                }
            }));
        } else if (data.type === 'group') {
            var labels = [], maxVals = [], avgVals = [], minVals = [];
            var uslLine = [], tgtLine = [], lslLine = [];
            for (var i = 0; i < data.groupStats.length; i++) {
                labels.push(data.groupStats[i].batch);
                maxVals.push(data.groupStats[i].max);
                avgVals.push(data.groupStats[i].avg);
                minVals.push(data.groupStats[i].min);
                uslLine.push(data.specs.usl);
                tgtLine.push(data.specs.target);
                lslLine.push(data.specs.lsl);
            }

            this.chartInstances.push(new Chart(document.getElementById('groupChart'), {
                type: 'line',
                data: {
                    labels: labels, datasets: [
                        { label: 'Max', data: maxVals, borderColor: '#f87171', borderWidth: 1.5, pointRadius: 0, fill: false },
                        { label: 'Avg', data: avgVals, borderColor: '#3b82f6', borderWidth: 2, pointRadius: 3, pointHoverRadius: 5, fill: false },
                        { label: 'Min', data: minVals, borderColor: '#f87171', borderWidth: 1.5, pointRadius: 0, fill: false },
                        { label: 'USL', data: uslLine, borderColor: '#ff9800', borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, fill: false },
                        { label: 'Target', data: tgtLine, borderColor: '#10b981', borderWidth: 1.5, pointRadius: 0, fill: false },
                        { label: 'LSL', data: lslLine, borderColor: '#ff9800', borderWidth: 1.5, borderDash: [6, 4], pointRadius: 0, fill: false }
                    ]
                },
                options: {
                    responsive: true,
                    aspectRatio: 3.5,
                    plugins: { legend: { position: 'top', labels: { boxWidth: 12, padding: 20, font: { size: 14 } } } },
                    scales: { x: { grid: { display: false } }, y: { grid: { color: 'rgba(0,0,0,0.08)' } } }
                }
            }));
        }
    },

    downloadExcel: function () {
        var self = this;
        var data = this.analysisResults;
        var wb = XLSX.utils.book_new();
        var maxBatchesPerSheet = 25;

        if (data.type === 'batch') {
            // VBA style: Horizontal layout with batches as columns
            var totalBatches = Math.min(data.batchNames.length, data.xbarR.xBar.data.length);
            var totalSheets = Math.ceil(totalBatches / maxBatchesPerSheet);
            var cavityCount = data.cavityNames ? data.cavityNames.length : data.xbarR.summary.n;

            for (var sheetIdx = 0; sheetIdx < totalSheets; sheetIdx++) {
                var startBatch = sheetIdx * maxBatchesPerSheet;
                var endBatch = Math.min((sheetIdx + 1) * maxBatchesPerSheet, totalBatches);
                var batchCount = endBatch - startBatch;
                var wsData = [];

                // Row 1: Title
                var row1 = [this.t('X̄ - R 管制圖', 'X-Bar R Control Chart')];
                for (var c = 0; c < batchCount + 5; c++) row1.push('');
                wsData.push(row1);

                // Row 2: Product info header
                var row2 = [this.t('產品名稱', 'Product'), '', '', '', this.t('規格', 'Spec'), '', '', this.t('管制界限', 'Limits'), '', '', '', this.t('彙總', 'Summary')];
                wsData.push(row2);

                // Row 3: Specs and limits
                var row3 = [this.t('檢驗項目', 'Item'), this.selectedItem, '', '',
                    'Target', data.specs.target, '',
                    'UCL', SPCEngine.round(data.xbarR.xBar.UCL, 4), '', '',
                    'ΣX̄', SPCEngine.round(data.xbarR.summary.xDoubleBar * data.xbarR.summary.k, 4)];
                wsData.push(row3);

                // Row 4: More specs
                var row4 = [this.t('模穴數', 'Cavities'), cavityCount, '', '',
                    'USL', data.specs.usl, '',
                    'CL', SPCEngine.round(data.xbarR.xBar.CL, 4), '', '',
                    'X̿', SPCEngine.round(data.xbarR.summary.xDoubleBar, 4)];
                wsData.push(row4);

                // Row 5: More info
                var row5 = [this.t('批號數', 'Batches'), data.xbarR.summary.k, '', '',
                    'LSL', data.specs.lsl, '',
                    'LCL', SPCEngine.round(data.xbarR.xBar.LCL, 4), '', '',
                    'ΣR', SPCEngine.round(data.xbarR.summary.rBar * data.xbarR.summary.k, 4)];
                wsData.push(row5);

                // Row 6: R chart limits
                var row6 = ['', '', '', '', '', '', '',
                    'R_UCL', SPCEngine.round(data.xbarR.R.UCL, 4), '', '',
                    'R̄', SPCEngine.round(data.xbarR.summary.rBar, 4)];
                wsData.push(row6);

                // Empty row
                wsData.push([]);

                // Row 8: Data table header - batch names
                var headerRow = [this.t('檢驗批號', 'Batch No.')];
                for (var b = startBatch; b < endBatch; b++) {
                    headerRow.push(data.batchNames[b] || 'B' + (b + 1));
                }
                wsData.push(headerRow);

                // Rows for each cavity (X1, X2, X3...)
                for (var cav = 0; cav < cavityCount; cav++) {
                    var cavRow = ['X' + (cav + 1)];
                    for (var b = startBatch; b < endBatch; b++) {
                        var value = data.dataMatrix && data.dataMatrix[b] ? data.dataMatrix[b][cav] : null;
                        cavRow.push(value !== null ? SPCEngine.round(value, 4) : '');
                    }
                    wsData.push(cavRow);
                }

                // X-bar row
                var xbarRow = ['X̄'];
                for (var b = startBatch; b < endBatch; b++) {
                    xbarRow.push(SPCEngine.round(data.xbarR.xBar.data[b], 4));
                }
                wsData.push(xbarRow);

                // R row
                var rRow = ['R'];
                for (var b = startBatch; b < endBatch; b++) {
                    rRow.push(SPCEngine.round(data.xbarR.R.data[b], 4));
                }
                wsData.push(rRow);

                // Create worksheet
                var ws = XLSX.utils.aoa_to_sheet(wsData);

                // Set column widths
                var cols = [{ wch: 12 }];
                for (var c = 0; c < batchCount + 12; c++) cols.push({ wch: 10 });
                ws['!cols'] = cols;

                var sheetName = this.selectedItem.substring(0, 25) + '-' + String(sheetIdx + 1).padStart(3, '0');
                XLSX.utils.book_append_sheet(wb, ws, sheetName);
            }


        } else if (data.type === 'cavity') {
            var wsData = [];
            wsData.push([this.t('模穴分析', 'Cavity Analysis'), '', '', '', '', this.t('檢驗項目', 'Inspection Item') + ': ' + this.selectedItem]);
            wsData.push([]);
            wsData.push(['Target', data.specs.target, 'USL', data.specs.usl, 'LSL', data.specs.lsl]);
            wsData.push([]);
            wsData.push([this.t('模穴', 'Cavity'), this.t('平均', 'Mean'), 'σ_within', 'σ_overall', 'Cp', 'Cpk', 'Pp', 'Ppk', 'n']);

            for (var i = 0; i < data.cavityStats.length; i++) {
                var s = data.cavityStats[i];
                wsData.push([s.name, SPCEngine.round(s.mean, 4), SPCEngine.round(s.withinStdDev, 4), SPCEngine.round(s.overallStdDev, 4), SPCEngine.round(s.Cp, 3), SPCEngine.round(s.Cpk, 3), SPCEngine.round(s.Pp, 3), SPCEngine.round(s.Ppk, 3), s.count]);
            }

            var ws = XLSX.utils.aoa_to_sheet(wsData);
            ws['!cols'] = [{ wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 8 }];
            XLSX.utils.book_append_sheet(wb, ws, this.t('模穴分析', 'CavityAnalysis'));

        } else if (data.type === 'group') {
            var wsData = [];
            wsData.push([this.t('群組分析 (Min-Max-Avg)', 'Group Analysis (Min-Max-Avg)'), '', '', '', '', this.t('檢驗項目', 'Inspection Item') + ': ' + this.selectedItem]);
            wsData.push([]);
            wsData.push(['Target', data.specs.target, 'USL', data.specs.usl, 'LSL', data.specs.lsl]);
            wsData.push([]);
            wsData.push([this.t('批號', 'Batch'), this.t('平均', 'Avg'), this.t('最大', 'Max'), this.t('最小', 'Min'), this.t('全距', 'Range'), 'n']);

            for (var i = 0; i < data.groupStats.length; i++) {
                var s = data.groupStats[i];
                wsData.push([s.batch, SPCEngine.round(s.avg, 4), SPCEngine.round(s.max, 4), SPCEngine.round(s.min, 4), SPCEngine.round(s.range, 4), s.count]);
            }

            var ws = XLSX.utils.aoa_to_sheet(wsData);
            ws['!cols'] = [{ wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 8 }];
            XLSX.utils.book_append_sheet(wb, ws, this.t('群組分析', 'GroupAnalysis'));
        }

        var filename = 'SPC_' + data.type + '_' + this.selectedItem + '.xlsx';
        XLSX.writeFile(wb, filename);
    },

    showLoading: function (text) {
        document.getElementById('loadingText').textContent = text;
        document.getElementById('loadingOverlay').classList.add('active');
    },

    hideLoading: function () {
        document.getElementById('loadingOverlay').classList.remove('active');
    },

    resetApp: function () {
        this.workbook = null;
        this.selectedItem = null;
        this.analysisResults = null;
        for (var i = 0; i < this.chartInstances.length; i++) this.chartInstances[i].destroy();
        this.chartInstances = [];
        document.getElementById('fileInput').value = '';
        document.getElementById('uploadZone').style.display = 'block';
        document.getElementById('fileInfo').style.display = 'none';
        document.getElementById('step2').style.display = 'none';
        document.getElementById('step3').style.display = 'none';
        document.getElementById('results').style.display = 'none';
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }
};

document.addEventListener('DOMContentLoaded', function () { SPCApp.init(); });
