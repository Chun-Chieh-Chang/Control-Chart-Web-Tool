// ============================================================================
// SPC Analysis Tool - Professional Final Version
// Layout: Hyper-Compact 1-Screen / Precision: 4 Decimals / Export: VBA Style
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

    range: function (data) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length === 0) return 0;
        return Math.max.apply(null, filtered) - Math.min.apply(null, filtered);
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
            xBar: { data: xBars, UCL: xUCL, CL: xDoubleBar, LCL: xLCL, sigma: (xUCL - xDoubleBar) / 3 },
            R: { data: ranges, UCL: rUCL, CL: rBar, LCL: rLCL },
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
            // Rule 1: Outside 3 Sigma
            if (Math.abs(data[i] - cl) > 3 * sigma) rules.push(1);
            // Rule 2: 9 points on one side
            if (i >= 8) {
                var side = data[i] > cl;
                var count = 0;
                for (var j = i - 8; j <= i; j++) { if ((data[j] > cl) === side && data[j] !== cl) count++; }
                if (count === 9) rules.push(2);
            }
            // Rule 3: 6 points trending
            if (i >= 5) {
                var inc = true, dec = true;
                for (var j = i - 5; j < i; j++) { if (data[j + 1] <= data[j]) inc = false; if (data[j + 1] >= data[j]) dec = false; }
                if (inc || dec) rules.push(3);
            }
            // Rule 5: 2/3 points outside 2 Sigma
            if (i >= 2) {
                var up = 0, lo = 0;
                for (var j = i - 2; j <= i; j++) { if (data[j] > cl + 2 * sigma) up++; if (data[j] < cl - 2 * sigma) lo++; }
                if (up >= 2 || lo >= 2) rules.push(5);
            }
            if (rules.length > 0) violations.push({ index: i, rules: rules });
        }
        return violations;
    },

    calculateProcessCapability: function (data, usl, lsl) {
        var filtered = data.filter(function (v) { return v !== null && !isNaN(v); });
        if (filtered.length < 2) return { Cp: 0, Cpk: 0, Pp: 0, Ppk: 0, mean: 0, count: 0 };
        var mean = this.mean(filtered);
        var sumSq = 0;
        for (var i = 0; i < filtered.length; i++) sumSq += Math.pow(filtered[i] - mean, 2);
        var overallStdDev = Math.sqrt(sumSq / (filtered.length - 1));
        var tolerance = usl - lsl;
        var Pp = overallStdDev > 0 ? tolerance / (6 * overallStdDev) : 0;
        var Ppk = overallStdDev > 0 ? Math.min((usl - mean) / (3 * overallStdDev), (mean - lsl) / (3 * overallStdDev)) : 0;
        return { Cp: Pp, Cpk: Ppk, Pp: Pp, Ppk: Ppk, mean: mean, overallStdDev: overallStdDev, count: filtered.length };
    },

    getCapabilityColor: function (cpk) {
        if (cpk >= 1.33) return { bg: '#c6efce', text: '#006100' };
        if (cpk >= 1.0) return { bg: '#ffeb9c', text: '#9c5700' };
        return { bg: '#ffc7ce', text: '#9c0006' };
    },

    round: function (v, d) { return Math.round(v * Math.pow(10, d || 4)) / Math.pow(10, d || 4); }
};

function cleanBatchName(name) {
    if (!name) return '';
    return String(name).trim().replace(/[\-_]\d+$/, '').replace(/\s*\(\d+\)$/, '').trim();
}

function DataInput(worksheet) {
    this.ws = worksheet;
    this.data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    this.parse();
}

DataInput.prototype.parse = function () {
    this.productInfo = {
        name: this.data[0] && this.data[0][1] ? this.data[0][1] : '',
        item: this.data[0] && this.data[0][2] ? this.data[0][2] : '',
        unit: 'Inch', char: '平均值/全距', dept: '品管部', inspector: '品管組', batchRange: '', chartNo: ''
    };
    this.headers = this.data[0] || [];
    this.specs = {
        target: parseFloat(this.data[1] && this.data[1][1]) || 0,
        usl: parseFloat(this.data[1] && this.data[1][2]) || 0,
        lsl: parseFloat(this.data[1] && this.data[1][3]) || 0
    };
    this.cavityColumns = [];
    for (var i = 0; i < this.headers.length; i++) {
        if (typeof this.headers[i] === 'string' && this.headers[i].indexOf('穴') >= 0) {
            this.cavityColumns.push({ index: i, name: this.headers[i] });
        }
    }
    this.dataRows = this.data.slice(2);
    this.batchNames = [];
    for (var j = 0; j < this.dataRows.length; j++) {
        if (this.dataRows[j][0]) this.batchNames.push(cleanBatchName(this.dataRows[j][0]));
    }
    if (this.batchNames.length > 0) {
        this.productInfo.batchRange = this.batchNames[0] + ' ~ ' + this.batchNames[this.batchNames.length - 1];
    }
};

DataInput.prototype.getDataMatrix = function () {
    var matrix = [];
    for (var i = 0; i < this.dataRows.length; i++) {
        var row = [];
        for (var j = 0; j < this.cavityColumns.length; j++) {
            var v = parseFloat(this.dataRows[i][this.cavityColumns[j].index]);
            row.push(isNaN(v) ? null : v);
        }
        matrix.push(row);
    }
    return matrix;
};

// --- APP CORE ---
var SPCApp = {
    currentLanguage: 'zh', workbook: null, selectedItem: null, analysisResults: null, chartInstances: [],
    init: function () { this.setupLanguageToggle(); this.setupFileUpload(); this.setupEventListeners(); this.updateLanguage(); },
    t: function (zh, en) { return this.currentLanguage === 'zh' ? zh : en; },
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
        document.querySelectorAll('[data-en][data-zh]').forEach(function (el) {
            el.textContent = self.currentLanguage === 'zh' ? el.dataset.zh : el.dataset.en;
        });
    },
    setupFileUpload: function () {
        var self = this;
        var zone = document.getElementById('uploadZone');
        var input = document.getElementById('fileInput');
        zone.addEventListener('click', function () { input.click(); });
        input.addEventListener('change', function (e) { if (e.target.files.length) self.handleFile(e.target.files[0]); });
        document.getElementById('resetBtn').addEventListener('click', function () { location.reload(); });
    },
    handleFile: function (file) {
        var self = this;
        var reader = new FileReader();
        reader.onload = function (e) {
            var data = new Uint8Array(e.target.result);
            self.workbook = XLSX.read(data, { type: 'array' });
            document.getElementById('uploadZone').style.display = 'none';
            document.getElementById('fileInfo').style.display = 'flex';
            document.getElementById('fileName').textContent = file.name;
            self.showInspectionItems();
        };
        reader.readAsArrayBuffer(file);
    },
    showInspectionItems: function () {
        var self = this;
        var list = document.getElementById('itemList');
        list.innerHTML = '';
        this.workbook.SheetNames.forEach(function (name) {
            var card = document.createElement('div');
            card.className = 'item-card';
            card.innerHTML = '<h3>' + name + '</h3><div class="item-badge">' + self.t('選擇', 'Select') + '</div>';
            card.onclick = function () { self.selectedItem = name; document.getElementById('step3').style.display = 'block'; window.scrollTo(0, document.body.scrollHeight); };
            list.appendChild(card);
        });
        document.getElementById('step2').style.display = 'block';
    },
    setupEventListeners: function () {
        var self = this;
        document.querySelectorAll('[data-analysis]').forEach(function (btn) {
            btn.onclick = function () { self.executeAnalysis(this.dataset.analysis); };
        });
        document.getElementById('downloadExcel').onclick = function () { self.downloadExcel(); };
    },
    executeAnalysis: function (type) {
        var self = this;
        var ws = this.workbook.Sheets[this.selectedItem];
        var input = new DataInput(ws);
        var matrix = input.getDataMatrix();
        if (type === 'batch') {
            this.analysisResults = { type: 'batch', xbarR: SPCEngine.calculateXBarRLimits(matrix), batchNames: input.batchNames, specs: input.specs, dataMatrix: matrix, cavityNames: input.getCavityNames(), productInfo: input.productInfo };
        } else if (type === 'cavity') {
            var stats = input.getCavityNames().map(function (n, i) {
                var cap = SPCEngine.calculateProcessCapability(matrix.map(function (r) { return r[i]; }), input.specs.usl, input.specs.lsl);
                cap.name = n; return cap;
            });
            this.analysisResults = { type: 'cavity', cavityStats: stats, specs: input.specs };
        }
        this.displayResults();
    },
    displayResults: function () {
        var self = this;
        var html = '';
        var data = this.analysisResults;
        if (data.type === 'batch') {
            var total = data.batchNames.length;
            this.batchPagination = { currentPage: 1, totalPages: Math.ceil(total / 25), maxPerPage: 25, totalBatches: total };
            html = '<div class="results-summary"><div class="stat-card"><div class="stat-label">n</div><div class="stat-value">' + data.xbarR.summary.n + '</div></div><div class="stat-card"><div class="stat-label">k</div><div class="stat-value">' + total + '</div></div></div>';
            if (this.batchPagination.totalPages > 1) {
                html += '<div class="pagination-controls" style="display:flex;justify-content:center;gap:15px;margin:15px 0;"><button id="prevPageBtn" class="btn-secondary">Prev</button><span id="pageInfo"></span><button id="nextPageBtn" class="btn-secondary">Next</button></div>';
            }
            html += '<div id="detailedTableContainer" style="overflow-x:auto; margin-bottom:20px;"></div><div id="pageLimitsContainer"></div><div id="diagnosticContainer"></div><div class="chart-container"><canvas id="xbarChart"></canvas></div><div class="chart-container"><canvas id="rChart"></canvas></div>';
        } else {
            html = '<div class="chart-container"><canvas id="cpkChart"></canvas></div>';
        }
        document.getElementById('resultsContent').innerHTML = html;
        document.getElementById('results').style.display = 'block';
        if (data.type === 'batch' && this.batchPagination.totalPages > 1) {
            document.getElementById('prevPageBtn').onclick = function () { self.changePage(-1); };
            document.getElementById('nextPageBtn').onclick = function () { self.changePage(1); };
        }
        this.renderCharts();
    },
    changePage: function (d) {
        var p = this.batchPagination;
        if (p.currentPage + d >= 1 && p.currentPage + d <= p.totalPages) { p.currentPage += d; this.renderCharts(); }
    },
    renderDetailedDataTable: function (labels, matrix, xbarR) {
        var info = this.analysisResults.productInfo;
        var specs = this.analysisResults.specs;
        var html = '<table class="excel-table" style="width:100\%; border-collapse:collapse; font-size:11px; font-family:sans-serif; border:2px solid #000; table-layout:fixed;">';
        html += '<colgroup><col style="width:4.5\%;">';
        for (var i = 0; i < 25; i++) html += '<col style="width:3.58\%;">';
        html += '<col style="width:1.5\%;" span="4"></colgroup>';
        html += '<tr style="background:#f3f4f6;"><td colspan="30" style="border:1px solid #000; text-align:center; font-weight:bold; font-size:14px; padding:3px;">X̄ - R 管制圖</td></tr>';
        var rows = [
            { l1: '商品名稱', v1: info.name, l2: '規格', v2: '標準', l3: '管制圖', v3: 'X̄', v4: 'R', l4: '製造部門', v4_val: info.dept },
            { l1: '商品料號', v1: info.item, l2: '最大值', v2: specs.usl, l3: '上限', v3: SPCEngine.round(xbarR.xBar.UCL, 4), v4: SPCEngine.round(xbarR.R.UCL, 4), l4: '檢驗人員', v4_val: info.inspector },
            { l1: '測量單位', v1: info.unit, l2: '目標值', v2: specs.target, l3: '中心值', v3: SPCEngine.round(xbarR.xBar.CL, 4), v4: SPCEngine.round(xbarR.R.CL, 4), l4: '管制特性', v4_val: info.char },
            { l1: '檢驗日期', v1: info.batchRange || '-', l2: '最小值', v2: specs.lsl, l3: '下限', v3: SPCEngine.round(xbarR.xBar.LCL, 4), v4: '-', l4: '圖表編號', v4_val: info.chartNo || '-' }
        ];
        rows.forEach(function (r) {
            html += '<tr><td colspan="3" style="border:1px solid #000; padding:1px 4px; font-weight:bold; background:#f9fafb;">' + r.l1 + '</td><td colspan="11" style="border:1px solid #000; padding:1px 4px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + r.v1 + '</td><td colspan="2" style="border:1px solid #000; padding:1px 4px; font-weight:bold; background:#f9fafb;">' + r.l2 + '</td><td colspan="2" style="border:1px solid #000; padding:1px 4px;">' + r.v2 + '</td><td colspan="2" style="border:1px solid #000; padding:1px 4px; font-weight:bold; background:#f9fafb;">' + r.l3 + '</td><td colspan="2" style="border:1px solid #000; padding:1px 4px;">' + r.v3 + '</td><td colspan="2" style="border:1px solid #000; padding:1px 4px;">' + (r.v4 || '') + '</td><td colspan="2" style="border:1px solid #000; padding:1px 4px; font-weight:bold; background:#f9fafb;">' + (r.l4 || '') + '</td><td colspan="4" style="border:1px solid #000; padding:1px 4px;">' + r.v4_val + '</td></tr>';
        });
        html += '<tr style="background:#e5e7eb; font-weight:bold;"><td style="border:1px solid #000; text-align:center;">批號</td>';
        for (var i = 0; i < 25; i++) html += '<td style="border:1px solid #000; text-align:center; font-size:9px; height:35px; word-break:break-all;">' + (labels[i] || '') + '</td>';
        html += '<td colspan="4" style="border:1px solid #000; text-align:center;">彙總</td></tr>';
        for (var i = 0; i < xbarR.summary.n; i++) {
            html += '<tr><td style="border:1px solid #000; text-align:center; font-weight:bold;">X' + (i + 1) + '</td>';
            for (var j = 0; j < 25; j++) html += '<td style="border:1px solid #000; text-align:center;">' + (matrix[j] && matrix[j][i] !== null ? matrix[j][i] : '') + '</td>';
            if (i === 0) html += '<td colspan="4" rowspan="2" style="border:1px solid #000; padding:0 2px; font-weight:bold; font-size:10px; background:#fefefe; vertical-align:middle; line-height:1;">ΣX̄=' + SPCEngine.round(xbarR.xBar.data.reduce(function (a, b) { return a + b; }, 0), 4) + '</td>';
            else if (i === 2) html += '<td colspan="4" rowspan="2" style="border:1px solid #000; padding:0 2px; font-weight:bold; font-size:10px; background:#fefefe; vertical-align:middle; line-height:1;">X̿=' + SPCEngine.round(xbarR.summary.xDoubleBar, 4) + '</td>';
            else if (i === 4) html += '<td colspan="4" rowspan="2" style="border:1px solid #000; padding:0 2px; font-weight:bold; font-size:10px; background:#fefefe; vertical-align:middle; line-height:1;">ΣR=' + SPCEngine.round(xbarR.R.data.reduce(function (a, b) { return a + b; }, 0), 4) + '</td>';
            else if (i === 6) html += '<td colspan="4" rowspan="2" style="border:1px solid #000; padding:0 2px; font-weight:bold; font-size:10px; background:#fefefe; vertical-align:middle; line-height:1;">R̄=' + SPCEngine.round(xbarR.summary.rBar, 4) + '</td>';
            else if (i >= 8) html += '<td colspan="4" style="border:1px solid #000; background:#fcfcfc;"></td>';
            html += '</tr>';
        }
        // Footer: ΣX, X̄, R
        ['ΣX', 'X̄', 'R'].forEach(function (rowLabel) {
            html += '<tr style="background:#f9fafb;"><td style="border:1px solid #000; text-align:center; font-weight:bold;">' + rowLabel + '</td>';
            for (var k = 0; k < 25; k++) {
                var val = '', style = '';
                if (matrix[k]) {
                    if (rowLabel === 'ΣX') val = SPCEngine.round(matrix[k].reduce(function (a, b) { return a + (b || 0); }, 0), 4);
                    if (rowLabel === 'X̄') { val = SPCEngine.round(xbarR.xBar.data[k], 4); if (val > xbarR.xBar.UCL || val < xbarR.xBar.LCL) style = 'background:yellow;'; }
                    if (rowLabel === 'R') { val = SPCEngine.round(xbarR.R.data[k], 4); if (val > xbarR.R.UCL) style = 'background:yellow;'; }
                }
                html += '<td style="border:1px solid #000; text-align:center;' + style + '">' + val + '</td>';
            }
            html += '<td colspan="4" style="border:1px solid #000; background:#f9fafb;"></td></tr>';
        });
        html += '</table>';
        document.getElementById('detailedTableContainer').innerHTML = html;
    },
    renderCharts: function () {
        var self = this; var data = this.analysisResults;
        this.chartInstances.forEach(function (c) { c.destroy(); }); this.chartInstances = [];
        if (data.type === 'batch') {
            var p = this.batchPagination; var start = (p.currentPage - 1) * p.maxPerPage; var end = Math.min(start + p.maxPerPage, p.totalBatches);
            var pageLabels = data.batchNames.slice(start, end); var pageMatrix = data.dataMatrix.slice(start, end);
            var pageXbarR = SPCEngine.calculateXBarRLimits(pageMatrix);
            if (p.totalPages > 1) document.getElementById('pageInfo').textContent = p.currentPage + ' / ' + p.totalPages;
            this.renderDetailedDataTable(pageLabels, pageMatrix, pageXbarR);
            // Diagnostic
            var diagHtml = '<div class="diagnostic-panel" style="padding:15px; margin-top:15px; border:1px solid #e2e8f0; border-radius:8px;"><h3>' + this.t('異常診斷', 'Diagnostics') + '</h3>';
            if (pageXbarR.xBar.violations.length === 0) diagHtml += '<p style="color:#10b981;">✅ No violations found.</p>';
            else {
                diagHtml += '<ul>';
                pageXbarR.xBar.violations.forEach(function (v) {
                    diagHtml += '<li><strong>[' + pageLabels[v.index] + ']</strong>: Rule ' + v.rules.join(', ') + '</li>';
                });
                diagHtml += '</ul>';
            }
            document.getElementById('diagnosticContainer').innerHTML = diagHtml + '</div>';
            // Charts
            var commonOpt = { responsive: true, aspectRatio: 4, maintainAspectRatio: true, animation: { duration: 0 } };
            this.chartInstances.push(new Chart(document.getElementById('xbarChart'), {
                type: 'line',
                data: {
                    labels: pageLabels, datasets: [
                        { label: 'X̄', data: pageXbarR.xBar.data, borderColor: '#3b82f6', pointBackgroundColor: pageXbarR.xBar.violations.map(function (v) { return '#f00'; }), fill: false },
                        { label: 'UCL', data: pageLabels.map(function () { return pageXbarR.xBar.UCL; }), borderColor: '#ef4444', borderDash: [5, 5], pointRadius: 0, fill: false },
                        { label: 'CL', data: pageLabels.map(function () { return pageXbarR.xBar.CL; }), borderColor: '#64748b', pointRadius: 0, fill: false },
                        { label: 'LCL', data: pageLabels.map(function () { return pageXbarR.xBar.LCL; }), borderColor: '#ef4444', borderDash: [5, 5], pointRadius: 0, fill: false }
                    ]
                }, options: commonOpt
            }));
            this.chartInstances.push(new Chart(document.getElementById('rChart'), {
                type: 'line',
                data: {
                    labels: pageLabels, datasets: [
                        { label: 'R', data: pageXbarR.R.data, borderColor: '#10b981', fill: false },
                        { label: 'UCL', data: pageLabels.map(function () { return pageXbarR.R.UCL; }), borderColor: '#ef4444', borderDash: [5, 5], pointRadius: 0, fill: false }
                    ]
                }, options: commonOpt
            }));
        } else if (data.type === 'cavity') {
            this.chartInstances.push(new Chart(document.getElementById('cpkChart'), {
                type: 'bar',
                data: { labels: data.cavityStats.map(function (s) { return s.name; }), datasets: [{ label: 'Cpk', data: data.cavityStats.map(function (s) { return s.Cpk; }), backgroundColor: data.cavityStats.map(function (s) { return SPCEngine.getCapabilityColor(s.Cpk).bg; }) }] },
                options: { responsive: true, aspectRatio: 3 }
            }));
        }
    },
    downloadExcel: function () {
        var data = this.analysisResults; if (!data || data.type !== 'batch') return;
        var wb = XLSX.utils.book_new();
        // 1. Summary Sheet
        var summaryData = [["SPC Analysis Report"], ["Product", data.productInfo.name], ["Item", data.productInfo.item], ["Target", data.specs.target], ["USL", data.specs.usl], ["LSL", data.specs.lsl]];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), "Report Summary");
        // 2. Batches (VBA Format)
        for (var p = 0; p < Math.ceil(data.batchNames.length / 25); p++) {
            var start = p * 25; var end = Math.min(start + 25, data.batchNames.length);
            var pageLabels = data.batchNames.slice(start, end); var pageMatrix = data.dataMatrix.slice(start, end);
            var pageXbarR = SPCEngine.calculateXBarRLimits(pageMatrix);
            // Construct AOAs to match VBA table...
            var aoa = [["X-Bar / R Control Chart (Page " + (p + 1) + ")"], ["Item", data.productInfo.item, "USL", data.specs.usl, "X-UCL", pageXbarR.xBar.UCL], ["Name", data.productInfo.name, "Target", data.specs.target, "X-CL", pageXbarR.xBar.CL]];
            aoa.push(["Batch"].concat(pageLabels));
            for (var i = 0; i < data.xbarR.summary.n; i++) aoa.push(["X" + (i + 1)].concat(pageMatrix.map(function (r) { return r[i]; })));
            aoa.push(["X-Bar"].concat(pageXbarR.xBar.data));
            aoa.push(["R"].concat(pageXbarR.R.data));
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "Batches_" + (p + 1));
        }
        XLSX.writeFile(wb, "SPC_Analysis_Report.xlsx");
    }
};

document.addEventListener('DOMContentLoaded', function () { SPCApp.init(); });
