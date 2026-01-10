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
        this.loadFromHistory();
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

        // Allow clicking the zone to upload, BUT ignore if clicking the button
        // because the button has its own onclick="...click()" in HTML or acts as a trigger.
        uploadZone.addEventListener('click', function (e) {
            // STOP if the click originated from a button or inside a button
            if (e.target.tagName === 'BUTTON' || e.target.closest('button')) {
                return;
            }
            fileInput.click();
        });

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

        var resetBtn = document.getElementById('resetBtn');
        if (resetBtn) {
            resetBtn.addEventListener('click', function (e) {
                e.stopPropagation();
                self.resetApp();
            });
        }
    },




    history: [],

    getChartTheme: function () {
        var isDark = document.documentElement.classList.contains('dark');
        return {
            mode: isDark ? 'dark' : 'light',
            text: isDark ? '#94a3b8' : '#64748b',
            grid: isDark ? '#334155' : '#f1f5f9'
        };
    },

    saveToHistory: function (file, analysisType, item) {
        if (!file) return;
        var entry = {
            name: file.name,
            size: (file.size / 1024).toFixed(1) + ' KB',
            type: analysisType,
            item: item,
            time: new Date().toISOString(),
            id: Date.now()
        };
        this.history.unshift(entry);
        if (this.history.length > 20) this.history.pop();

        localStorage.setItem('spc_history', JSON.stringify(this.history));
        this.renderRecentFiles();
    },

    loadFromHistory: function () {
        var saved = localStorage.getItem('spc_history');
        if (saved) {
            try {
                this.history = JSON.parse(saved);
                this.renderRecentFiles();
            } catch (e) {
                console.error('History load error', e);
            }
        }
    },

    clearHistory: function () {
        this.history = [];
        localStorage.removeItem('spc_history');
        this.renderRecentFiles();
        this.renderHistoryView();
    },

    renderRecentFiles: function () {
        var container = document.getElementById('recentFilesContainer');
        if (!container) return;

        if (this.history.length === 0) {
            container.innerHTML = '<div class="text-[10px] text-slate-400 py-4 text-center italic" data-en="No recent activities" data-zh="尚無近期活動">' +
                this.t('尚無近期活動', 'No recent activities') + '</div>';
            return;
        }

        var html = this.history.slice(0, 5).map(function (h) {
            var d = new Date(h.time);
            var timeStr = d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            return '<div class="flex items-center group cursor-pointer hover:bg-slate-50 dark:hover:bg-slate-800/50 p-2 rounded-lg transition-all">' +
                '<div class="w-8 h-8 rounded-lg bg-slate-100 dark:bg-slate-700 flex items-center justify-center mr-3 group-hover:bg-primary/10 transition-colors">' +
                '<span class="material-icons-outlined text-sm text-slate-400 group-hover:text-primary transition-colors">description</span>' +
                '</div>' +
                '<div class="flex-1 min-w-0">' +
                '<div class="text-xs font-bold text-slate-800 dark:text-slate-200 truncate">' + h.name + '</div>' +
                '<div class="text-[10px] text-slate-400">' + h.size + ' • ' + timeStr + '</div>' +
                '</div>' +
                '</div>';
        }).join('');
        container.innerHTML = html;
    },

    renderHistoryView: function () {
        var body = document.getElementById('historyTableBody');
        if (!body) return;

        if (this.history.length === 0) {
            body.innerHTML = '<tr><td colspan="5" class="px-6 py-12 text-center text-slate-400 italic">No history records found</td></tr>';
            return;
        }

        var self = this;
        body.innerHTML = this.history.map(function (h) {
            var d = new Date(h.time);
            var dateStr = d.toLocaleDateString() + ' ' + d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
            return '<tr class="hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-colors">' +
                '<td class="px-6 py-4 font-medium text-slate-900 dark:text-white">' + h.name + '</td>' +
                '<td class="px-6 py-4">' +
                '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold uppercase ' + (h.type === 'batch' ? 'bg-indigo-50 text-indigo-600' : h.type === 'cavity' ? 'bg-rose-50 text-rose-600' : 'bg-amber-50 text-amber-600') + '">' + h.type + '</span>' +
                '</td>' +
                '<td class="px-6 py-4 text-xs font-mono text-slate-500">' + h.item + '</td>' +
                '<td class="px-6 py-4 text-xs text-slate-500">' + dateStr + '</td>' +
                '<td class="px-6 py-4 text-center">' +
                '<button class="text-primary hover:text-indigo-700 font-bold text-xs">View Log</button>' +
                '</td>' +
                '</tr>';
        }).join('');
    },

    selectedFile: null,
    handleFile: function (file) {
        var self = this;
        this.selectedFile = file;
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

        // Add a clear instruction header if not present
        // (Actually header is in HTML "Select Inspection Item")

        var sheets = this.workbook.SheetNames;
        for (var i = 0; i < sheets.length; i++) {
            var sheetName = sheets[i];
            var ws = self.workbook.Sheets[sheetName];
            var data = XLSX.utils.sheet_to_json(ws, { header: 1 });
            var target = data[1] && data[1][1] ? data[1][1] : 'N/A';

            var card = document.createElement('div');
            // Material SaaS Selection Card
            card.className = 'saas-card p-6 cursor-pointer hover:border-primary hover:shadow-xl hover:shadow-indigo-50 dark:hover:shadow-none transition-all group relative overflow-hidden';

            card.innerHTML =
                '<div class="relative z-10">' +
                '<div class="flex justify-between items-start mb-4">' +
                '<h3 class="text-base font-bold text-slate-900 dark:text-white group-hover:text-primary transition-colors">' + sheetName + '</h3>' +
                '<span class="material-icons-outlined text-slate-400 group-hover:text-primary transition-colors">check_circle_outline</span>' +
                '</div>' +
                '<div class="flex flex-col space-y-2">' +
                '<div class="text-[10px] text-slate-400 font-bold uppercase tracking-widest">' + this.t('目標值', 'Target Value') + '</div>' +
                '<div class="text-sm font-mono font-bold text-slate-700 dark:text-slate-300">' + target + '</div>' +
                '</div>' +
                '</div>' +
                '<div class="absolute -right-2 -bottom-2 opacity-[0.05] group-hover:opacity-[0.1] transition-opacity">' +
                '<span class="material-icons-outlined text-7xl">table_view</span>' +
                '</div>';

            card.dataset.sheet = sheetName;
            card.addEventListener('click', function () {
                self.selectedItem = this.dataset.sheet;
                self.showAnalysisOptions();
            });
            itemList.appendChild(card);
        }

        document.getElementById('step2').style.display = 'block';

        // Use a slight timeout to ensure visibility before scrolling
        setTimeout(function () {
            document.getElementById('step2').scrollIntoView({ behavior: 'smooth', block: 'center' });
        }, 100);
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

        // Recent Files Clear All
        var clearBtn = document.getElementById('clearRecentBtn');
        if (clearBtn) {
            clearBtn.addEventListener('click', function () {
                self.clearHistory();
            });
        }

        // Sidebar Navigation
        var navIds = ['nav-dashboard', 'nav-import', 'nav-analysis', 'nav-history', 'nav-settings'];
        navIds.forEach(function (id) {
            var el = document.getElementById(id);
            if (el) {
                el.addEventListener('click', function (e) {
                    e.preventDefault();
                    self.switchView(id.replace('nav-', ''));
                });
            }
        });

        this.setupSearch();
        this.setupDarkModeToggle();
    },

    setupSearch: function () {
        var self = this;
        var input = document.getElementById('globalSearch');
        if (input) {
            input.addEventListener('input', function (e) {
                var term = e.target.value.toLowerCase();
                self.performSearch(term);
            });
        }
    },

    performSearch: function (term) {
        // Filter inspection items if visible (Step 2 cards)
        var cards = document.querySelectorAll('#itemList .saas-card');
        cards.forEach(function (card) {
            var text = card.textContent.toLowerCase();
            card.style.display = text.includes(term) ? 'block' : 'none';
        });

        // Filter history rows if visible
        var rows = document.querySelectorAll('#historyTableBody tr');
        rows.forEach(function (row) {
            var text = row.textContent.toLowerCase();
            row.style.display = text.includes(term) ? 'table-row' : 'none';
        });

        // Inform user if no results in current context
        console.log('Searching for: ' + term);
    },

    setupDarkModeToggle: function () {
        var self = this;
        var btn = document.getElementById('darkModeBtn');
        if (btn) {
            btn.addEventListener('click', function () {
                // Ensure chart theme updates because ApexCharts uses specific theme modes
                setTimeout(function () {
                    if (self.analysisResults) {
                        self.renderCharts();
                    }
                }, 100);
            });
        }
    },

    switchView: function (viewId) {
        var self = this;
        var viewMap = {
            'dashboard': 'view-import',
            'import': 'view-import',
            'analysis': 'view-analysis',
            'history': 'view-history',
            'settings': 'view-import'
        };

        var targetId = viewMap[viewId] || 'view-import';
        var vImport = document.getElementById('view-import');
        var vAnalysis = document.getElementById('view-analysis');
        var vHistory = document.getElementById('view-history');

        if (vImport) vImport.classList.add('hidden');
        if (vAnalysis) vAnalysis.classList.add('hidden');
        if (vHistory) vHistory.classList.add('hidden');

        var targetEl = document.getElementById(targetId);
        if (targetEl) targetEl.classList.remove('hidden');

        if (viewId === 'history') {
            this.renderHistoryView();
        }

        // Update sidebar links
        var navLinks = document.querySelectorAll('#main-nav .nav-link');
        for (var i = 0; i < navLinks.length; i++) {
            navLinks[i].className = 'nav-link flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-slate-400 hover:text-white hover:bg-slate-800 transition-all';
        }

        var activeLink = document.getElementById('nav-' + viewId);
        if (activeLink) {
            activeLink.className = 'nav-link flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium bg-primary text-white shadow-lg shadow-primary/20';
        }

        var breadcrumb = document.querySelector('header .text-slate-900');
        if (breadcrumb) {
            var titles = {
                'dashboard': self.t('數據總覽', 'Dashboard Overview'),
                'import': self.t('QIP 數據導入', 'QIP Import'),
                'analysis': self.t('統計分析結果', 'Statistical Analysis'),
                'history': self.t('歷史分析紀錄', 'History Records'),
                'settings': self.t('系統設定', 'System Settings')
            };
            breadcrumb.innerText = titles[viewId] || 'QIP 統計分析';
        }

        if (viewId === 'import') {
            var side = document.getElementById('anomalySidebar');
            if (side) side.classList.add('hidden');
        } else if (viewId === 'analysis' && this.analysisResults) {
            var side = document.getElementById('anomalySidebar');
            if (side) side.classList.remove('hidden');
        }
    },

    executeAnalysis: function (type) {
        var self = this;
        if (!this.selectedItem) {
            alert(this.t('請先選擇分析項目', 'Please select an item first'));
            return;
        }
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

                    // Flatten data matrix for global Cpk calculation
                    var allValues = [];
                    for (var i = 0; i < dataMatrix.length; i++) {
                        for (var j = 0; j < dataMatrix[i].length; j++) {
                            if (dataMatrix[i][j] !== null) allValues.push(dataMatrix[i][j]);
                        }
                    }
                    var specs = dataInput.getSpecs();
                    var capability = SPCEngine.calculateProcessCapability(allValues, specs.usl, specs.lsl);
                    xbarR.summary.Cpk = capability.Cpk;
                    xbarR.summary.Ppk = capability.Ppk;

                    results = {
                        type: 'batch',
                        xbarR: xbarR,
                        batchNames: dataInput.batchNames,
                        specs: specs,
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
                self.saveToHistory(self.selectedFile, type, self.selectedItem);
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

            // Compact Dashboard metrics
            html = '<div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">' +
                '<div class="saas-card p-4 flex flex-col justify-center">' +
                '<div class="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">' + this.t('模穴數 (n)', 'Cavity Count') + '</div>' +
                '<div class="text-xl font-bold text-slate-900 dark:text-white">' + data.xbarR.summary.n + '</div>' +
                '</div>' +
                '<div class="saas-card p-4 flex flex-col justify-center">' +
                '<div class="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">' + this.t('總記錄數', 'Total Batches') + '</div>' +
                '<div class="text-xl font-bold text-slate-900 dark:text-white">' + totalBatches + '</div>' +
                '</div>' +
                '<div class="saas-card p-4 flex flex-col justify-center">' +
                '<div class="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1">' + this.t('製程能力 (Cpk)', 'Process Cpk') + '</div>' +
                '<div class="text-xl font-bold text-primary">' + SPCEngine.round(data.xbarR.summary.Cpk, 3) + '</div>' +
                '</div>' +
                '<div class="saas-card p-4 flex flex-col justify-center dark:bg-slate-800">' +
                '<div class="text-[10px] font-bold text-slate-500 dark:text-slate-400 uppercase tracking-widest mb-1">' + this.t('異常狀態', 'Status') + '</div>' +
                '<div class="flex items-center space-x-2">' +
                '<span class="material-icons-outlined text-sm ' + (data.xbarR.xBar.violations.length > 0 ? 'text-rose-500' : 'text-emerald-500') + '">' + (data.xbarR.xBar.violations.length > 0 ? 'warning' : 'check_circle') + '</span>' +
                '<span class="text-sm font-semibold ' + (data.xbarR.xBar.violations.length > 0 ? 'text-rose-600 dark:text-rose-400' : 'text-emerald-600 dark:text-emerald-400') + '">' +
                (data.xbarR.xBar.violations.length > 0 ? this.t('偵測到異常', 'Alert') : this.t('系統受控', 'Normal')) + '</span>' +
                '</div>' +
                '</div>' +
                '</div>';

            if (totalPages > 1) {
                html += '<div class="flex items-center justify-between mb-4 bg-white dark:bg-slate-800 p-3 rounded-lg border border-slate-100 dark:border-slate-700 shadow-sm">' +
                    '<div class="text-xs font-medium text-slate-600 dark:text-slate-400" id="pageInfo">' + this.t('頁次 ', 'Page ') + '1 / ' + totalPages + '</div>' +
                    '<div class="flex gap-2">' +
                    '<button id="prevPageBtn" class="px-4 py-1.5 text-xs font-bold border border-slate-200 dark:border-slate-600 dark:text-white rounded-md hover:bg-slate-50 dark:hover:bg-slate-700 transition-all disabled:opacity-30 disabled:cursor-not-allowed">' + this.t('上一頁', 'Prev') + '</button>' +
                    '<button id="nextPageBtn" class="px-4 py-1.5 text-xs font-bold text-white bg-primary rounded-md hover:bg-indigo-700 transition-all disabled:opacity-30 disabled:cursor-not-allowed">' + this.t('下一頁', 'Next') + '</button>' +
                    '</div>' +
                    '</div>';
            }

            html += '<div id="detailedTableContainer" class="mb-10 overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm"></div>';
            html += '<div id="pageLimitsContainer"></div>';
            html += '<div id="diagnosticContainer"></div>';

            html += '<div class="grid grid-cols-1 gap-8">' +
                '<div class="saas-card p-8">' +
                '<div class="flex justify-between items-center mb-6">' +
                '<h3 class="text-base font-bold text-slate-800 dark:text-white flex items-center gap-2"><span class="material-icons-outlined text-primary text-lg">bar_chart</span>' + this.t('X̄ 管制圖 (均值)', 'X-Bar Control Chart') + '</h3>' +
                '<span class="px-2 py-0.5 bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 text-[10px] font-bold rounded uppercase tracking-wider">SVG Vector</span>' +
                '</div>' +
                '<div id="xbarChart" class="h-96"></div>' +
                '</div>' +
                '<div class="saas-card p-8">' +
                '<div class="flex justify-between items-center mb-6">' +
                '<h3 class="text-base font-bold text-slate-800 dark:text-white flex items-center gap-2"><span class="material-icons-outlined text-primary text-lg">show_chart</span>' + this.t('R 管制圖 (全距)', 'R Control Chart') + '</h3>' +
                '<span class="px-2 py-0.5 bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 text-[10px] font-bold rounded uppercase tracking-wider">SVG Vector</span>' +
                '</div>' +
                '<div id="rChart" class="h-80"></div>' +
                '</div>' +
                '</div>';

        } else if (data.type === 'cavity') {
            html = '<div class="grid grid-cols-1 gap-8">' +
                '<div class="saas-card p-8">' +
                '<h3 class="text-base font-bold text-slate-800 dark:text-white mb-6 flex items-center gap-2"><span class="material-icons-outlined text-primary text-lg">pie_chart</span>' + this.t('模穴 Cpk 效能比較', 'Cavity Cpk Performance') + '</h3>' +
                '<div id="cpkChart" class="h-96"></div>' +
                '</div>' +
                '<div class="grid grid-cols-1 md:grid-cols-2 gap-6">' +
                '<div class="saas-card p-8">' +
                '<h3 class="text-base font-bold text-slate-800 dark:text-white mb-6 flex items-center gap-2"><span class="material-icons-outlined text-primary text-lg">insights</span>' + this.t('模穴平均值比較', 'Cavity Mean Comparison') + '</h3>' +
                '<div id="meanChart" class="h-80"></div>' +
                '</div>' +
                '<div class="saas-card p-8">' +
                '<h3 class="text-base font-bold text-slate-800 dark:text-white mb-6 flex items-center gap-2"><span class="material-icons-outlined text-primary text-lg">scatter_plot</span>' + this.t('模穴標準差比較', 'Cavity StdDev Comparison') + '</h3>' +
                '<div id="stdDevChart" class="h-80"></div>' +
                '</div>' +
                '</div>' +
                '<div class="saas-card overflow-hidden">' +
                '<div class="p-6 border-b border-slate-50 dark:border-slate-700 flex justify-between items-center">' +
                '<h3 class="text-base font-bold text-slate-800 dark:text-white flex items-center gap-2"><span class="material-icons-outlined text-primary text-lg">table_chart</span>' + this.t('模穴細節統計', 'Cavity Details') + '</h3>' +
                '</div>' +
                '<div class="overflow-x-auto">' +
                '<table class="w-full text-sm text-left">' +
                '<thead class="text-[11px] text-slate-500 dark:text-slate-400 uppercase bg-slate-50/50 dark:bg-slate-700/50">' +
                '<tr>' +
                '<th class="px-8 py-4 font-bold">Cavity Name</th>' +
                '<th class="px-8 py-4 text-center">Mean</th>' +
                '<th class="px-8 py-4 text-center">Cpk</th>' +
                '<th class="px-8 py-4 text-center">n</th>' +
                '</tr>' +
                '</thead>' +
                '<tbody class="divide-y divide-slate-100 dark:divide-slate-700">' +
                data.cavityStats.map(function (s) {
                    return '<tr>' +
                        '<td class="px-8 py-4 font-bold text-slate-700 dark:text-slate-300">' + s.name + '</td>' +
                        '<td class="px-8 py-4 text-center font-mono text-slate-700 dark:text-slate-300">' + SPCEngine.round(s.mean, 4) + '</td>' +
                        '<td class="px-8 py-4 text-center">' +
                        '<span class="inline-block px-3 py-1 rounded-full text-xs font-bold ' +
                        (s.Cpk < 1.0 ? 'bg-rose-50 text-rose-600 dark:bg-rose-900/30 dark:text-rose-400' : 'bg-emerald-50 text-emerald-600 dark:bg-emerald-900/30 dark:text-emerald-400') + '">' +
                        SPCEngine.round(s.Cpk, 3) + '</span>' +
                        '</td>' +
                        '<td class="px-8 py-4 text-center text-slate-600 dark:text-slate-400 font-mono">' + s.count + '</td>' +
                        '</tr>';
                }).join('') +
                '</tbody>' +
                '</table>' +
                '</div>' +
                '</div>' +
                '</div>';

        } else if (data.type === 'group') {
            html = '<div class="grid grid-cols-1 gap-12">' +
                '<div class="saas-card p-8">' +
                '<h3 class="text-base font-bold text-slate-800 dark:text-white mb-6 flex items-center gap-2"><span class="material-icons-outlined text-primary text-lg">trending_up</span>' + this.t('群組數據趨勢 (Min-Max-Avg)', 'Group Trend Analysis') + '</h3>' +
                '<div id="groupChart" class="h-96"></div>' +
                '</div>' +
                '<div class="saas-card p-8">' +
                '<h3 class="text-base font-bold text-slate-800 dark:text-white mb-6 flex items-center gap-2"><span class="material-icons-outlined text-primary text-lg">compare_arrows</span>' + this.t('模穴間變異 (Range)', 'Inter-Cavity Variation') + '</h3>' +
                '<div id="groupVarChart" class="h-96"></div>' +
                '</div>' +
                '</div>';
        }

        resultsContent.innerHTML = html;
        document.getElementById('results').style.display = 'block';

        // Re-attach download listener
        var downloadBtn = document.getElementById('downloadExcel');
        if (downloadBtn) {
            // Remove old listeners by replacing the element or just adding a new one (delegation is better but this is explicit)
            var newBtn = downloadBtn.cloneNode(true);
            downloadBtn.parentNode.replaceChild(newBtn, downloadBtn);
            newBtn.addEventListener('click', function () { self.downloadExcel(); });
        }

        if (data.type === 'batch') {
            if (this.batchPagination.totalPages > 1) {
                document.getElementById('prevPageBtn').addEventListener('click', function () { self.changeBatchPage(-1); });
                document.getElementById('nextPageBtn').addEventListener('click', function () { self.changeBatchPage(1); });
                this.updatePaginationButtons();
            }
        }

        setTimeout(function () {
            self.renderCharts();
            self.switchView('analysis');
            document.getElementById('results').scrollIntoView({ behavior: 'smooth' });
        }, 100);
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
            batch: 58,     // 生產批號欄位寬度
            summary: 30    // 彙總基礎欄寬 (總寬度 = 30 * 4 = 120px)
        };
        // Calculate total width explicitly to force horizontal scrolling
        // Total = Label(1) + Batch(25) + Summary(4)
        var totalWidth = colWidths.label + (25 * colWidths.batch) + (4 * colWidths.summary);

        var html = '<table class="excel-table" style="width:' + totalWidth + 'px; border-collapse:collapse; font-size:10px; font-family:sans-serif; border:2px solid var(--table-border); table-layout:fixed;">';

        html += '<colgroup>';
        html += '<col style="width:' + colWidths.label + 'px;">';
        for (var c = 0; c < 25; c++) html += '<col style="width:' + colWidths.batch + 'px;">';
        html += '<col style="width:' + colWidths.summary + 'px;"><col style="width:' + colWidths.summary + 'px;"><col style="width:' + colWidths.summary + 'px;"><col style="width:' + colWidths.summary + 'px;">';
        html += '</colgroup>';


        // --- Row 1: Header ---
        html += '<tr style="background:var(--table-header-bg);"><td colspan="30" style="border:1px solid var(--table-border); text-align:center; font-weight:bold; font-size:14px; padding:3px;">X̄ - R 管制圖</td></tr>';

        // --- Row 2-5: Metadata & Limits (Re-distributed colspans out of 30) ---
        var rows = [
            { l1: '商品名稱', v1: info.name, l2: '規格', v2: '標準', l3: '管制圖', v3: 'X̄', v4: 'R', l4: '製造部門', v4_val: info.dept },
            { l1: '商品料號', v1: info.item, l2: '最大值', v2: SPCEngine.round(specs.usl, 4), l3: '上限', v3: SPCEngine.round(pageXbarR.xBar.UCL, 4), v4: SPCEngine.round(pageXbarR.R.UCL, 4), l4: '檢驗人員', v4_val: info.inspector },
            { l1: '測量單位', v1: info.unit, l2: '目標值', v2: SPCEngine.round(specs.target, 4), l3: '中心值', v3: SPCEngine.round(pageXbarR.xBar.CL, 4), v4: SPCEngine.round(pageXbarR.R.CL, 4), l4: '管制特性', v4_val: info.char },
            { l1: '檢驗日期', v1: info.batchRange || '-', l2: '最小值', v2: SPCEngine.round(specs.lsl, 4), l3: '下限', v3: SPCEngine.round(pageXbarR.xBar.LCL, 4), v4: '-', l4: '圖表編號', v4_val: info.chartNo || '-' }
        ];

        rows.forEach(function (r) {
            html += '<tr>' +
                '<td colspan="2" style="border:1px solid var(--table-border); padding:1px 1px; font-weight:bold; background:var(--table-label-bg); overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + r.l1 + '</td>' +
                '<td colspan="7" style="border:1px solid var(--table-border); padding:1px 1px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + r.v1 + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); padding:1px 1px; font-weight:bold; background:var(--table-label-bg); overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + r.l2 + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); padding:1px 1px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + r.v2 + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); padding:1px 1px; font-weight:bold; background:var(--table-label-bg); overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + r.l3 + '</td>' +
                '<td colspan="3" style="border:1px solid var(--table-border); padding:1px 1px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + r.v3 + '</td>' +
                '<td colspan="3" style="border:1px solid var(--table-border); padding:1px 1px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + (r.v4 || '') + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); padding:1px 1px; font-weight:bold; background:var(--table-label-bg); overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + (r.l4 || '') + '</td>' +
                '<td colspan="7" style="border:1px solid var(--table-border); padding:1px 1px; overflow:hidden; white-space:nowrap; text-overflow:ellipsis;">' + (r.v4_val || '') + '</td>' +
                '</tr>';
        });

        // --- Data Header: Batch Names ---
        html += '<tr style="background:var(--table-data-header-bg); font-weight:bold;">' +
            '<td style="border:1px solid var(--table-border); text-align:center;">批號</td>';
        pageLabels.forEach(function (name) {
            html += '<td style="border:1px solid var(--table-border); text-align:center; height:35px; overflow:hidden; font-size:10px; white-space:nowrap; text-overflow:ellipsis; width:' + colWidths.batch + 'px; max-width:' + colWidths.batch + 'px; min-width:' + colWidths.batch + 'px; padding:0;">' + name + '</td>';
        });
        // Fill empty if less than 25 batches
        for (var f = pageLabels.length; f < 25; f++) html += '<td style="border:1px solid var(--table-border); width:' + colWidths.batch + 'px; max-width:' + colWidths.batch + 'px; min-width:' + colWidths.batch + 'px;"></td>';

        html += '<td colspan="4" style="border:1px solid var(--table-border); text-align:center;">彙總</td></tr>';

        // --- Main Data Rows: Cavities ---
        for (var i = 0; i < cavityCount; i++) {
            html += '<tr><td style="border:1px solid var(--table-border); text-align:center; font-weight:bold; background:var(--table-label-bg);">X' + (i + 1) + '</td>';
            for (var j = 0; j < 25; j++) {
                var val = (pageDataMatrix[j] && pageDataMatrix[j][i] !== undefined) ? pageDataMatrix[j][i] : null;
                html += '<td style="border:1px solid var(--table-border); text-align:center; padding:0; overflow:hidden; white-space:nowrap; width:' + colWidths.batch + 'px; max-width:' + colWidths.batch + 'px; min-width:' + colWidths.batch + 'px;">' + (val !== null ? val : '') + '</td>';
            }

            // Sidebar summary on first few rows
            if (i === 0) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); text-align:center !important; font-weight:bold; background:var(--table-bg); font-size:10px;">ΣX̄ = ' + SPCEngine.round(pageXbarR.summary.xDoubleBar * pageXbarR.xBar.data.length, 4) + '</td>';
            else if (i === 2) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); text-align:center !important; font-weight:bold; background:var(--table-bg); font-size:10px;">X̿ = ' + SPCEngine.round(pageXbarR.summary.xDoubleBar, 4) + '</td>';
            else if (i === 4) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); text-align:center !important; font-weight:bold; background:var(--table-bg); font-size:10px;">ΣR = ' + SPCEngine.round(pageXbarR.summary.rBar * pageXbarR.R.data.length, 4) + '</td>';
            else if (i === 6) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); text-align:center !important; font-weight:bold; background:var(--table-bg); font-size:10px;">R̄ = ' + SPCEngine.round(pageXbarR.summary.rBar, 4) + '</td>';
            else if (i >= 8) html += '<td colspan="4" style="border:1px solid var(--table-border); background:transparent;"></td>';
            html += '</tr>';
        }

        // --- Footer Rows: ΣX, X̄, R ---
        // ΣX Row
        html += '<tr style="background:var(--table-header-bg);"><td style="border:1px solid var(--table-border); text-align:center; font-weight:bold;">ΣX</td>';
        for (var b = 0; b < 25; b++) {
            var val = '';
            if (pageDataMatrix[b]) {
                var sum = pageDataMatrix[b].reduce(function (a, b) { return a + (b || 0); }, 0);
                val = SPCEngine.round(sum, 4);
            }
            html += '<td style="border:1px solid var(--table-border); text-align:center;">' + val + '</td>';
        }
        html += '<td colspan="4" style="border:1px solid var(--table-border);"></td></tr>';

        // X̄ Row (with yellow highlighting)
        html += '<tr style="background:var(--table-header-bg);"><td style="border:1px solid var(--table-border); text-align:center; font-weight:bold;">X̄</td>';
        for (var k = 0; k < 25; k++) {
            var val = '', style = '';
            if (pageXbarR.xBar.data[k] !== undefined) {
                var v = pageXbarR.xBar.data[k];
                val = SPCEngine.round(v, 4);
                if (v > pageXbarR.xBar.UCL || v < pageXbarR.xBar.LCL) style = 'background:var(--table-ooc-bg); color:var(--table-ooc-text);';
            }
            html += '<td style="border:1px solid var(--table-border); text-align:center;' + style + '">' + val + '</td>';
        }
        html += '<td colspan="4" style="border:1px solid var(--table-border);"></td></tr>';

        // R Row
        html += '<tr style="background:var(--table-header-bg);"><td style="border:1px solid var(--table-border); text-align:center; font-weight:bold;">R</td>';
        for (var k = 0; k < 25; k++) {
            var val = '', style = '';
            if (pageXbarR.R.data[k] !== undefined) {
                var v = pageXbarR.R.data[k];
                val = SPCEngine.round(v, 4);
                if (v > pageXbarR.R.UCL) style = 'background:var(--table-ooc-bg); color:var(--table-ooc-text);';
            }
            html += '<td style="border:1px solid var(--table-border); text-align:center;' + style + '">' + val + '</td>';
        }
        html += '<td colspan="4" style="border:1px solid var(--table-border);"></td></tr>';

        html += '</table>';
        document.getElementById('detailedTableContainer').innerHTML = html;
    },




    renderCharts: function () {
        var self = this;
        var data = this.analysisResults;
        this.chartInstances.forEach(function (c) { if (c && c.destroy) c.destroy(); });
        this.chartInstances = [];

        if (data.type === 'batch') {
            var p = this.batchPagination || { currentPage: 1, maxPerPage: 25, totalBatches: data.batchNames.length };
            var startIdx = (p.currentPage - 1) * p.maxPerPage;
            var endIdx = Math.min(startIdx + p.maxPerPage, p.totalBatches);

            var pageLabels = data.batchNames.slice(startIdx, endIdx);
            var pageDataMatrix = data.dataMatrix.slice(startIdx, endIdx);
            var pageXbarR = SPCEngine.calculateXBarRLimits(pageDataMatrix);
            pageXbarR.summary.xBarSum = pageXbarR.xBar.data.reduce(function (a, b) { return a + b; }, 0);
            pageXbarR.summary.rSum = pageXbarR.R.data.reduce(function (a, b) { return a + b; }, 0);

            this.renderDetailedDataTable(pageLabels, pageDataMatrix, pageXbarR);

            var theme = this.getChartTheme();

            // X-Bar Apex Chart
            var xOptions = {
                chart: {
                    type: 'line',
                    height: 380,
                    toolbar: { show: false },
                    background: 'transparent',
                    animations: { enabled: true }
                },
                theme: { mode: theme.mode },
                series: [
                    { name: 'X-Bar', data: pageXbarR.xBar.data },
                    { name: 'UCL', data: new Array(pageLabels.length).fill(pageXbarR.xBar.UCL) },
                    { name: 'CL', data: new Array(pageLabels.length).fill(pageXbarR.xBar.CL) },
                    { name: 'LCL', data: new Array(pageLabels.length).fill(pageXbarR.xBar.LCL) }
                ],
                colors: ['#4f46e5', '#f43f5e', '#10b981', '#f43f5e'],
                stroke: { width: [3, 1.5, 1.5, 1.5], dashArray: [0, 6, 0, 6], curve: 'straight' },
                xaxis: { categories: pageLabels, labels: { style: { colors: theme.text, fontSize: '10px' } }, axisBorder: { color: theme.grid }, axisTicks: { color: theme.grid } },
                yaxis: {
                    labels: {
                        formatter: function (val) { return val.toFixed(4); },
                        style: { colors: theme.text }
                    }
                },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode, y: { formatter: function (val) { return val.toFixed(4); } } },
                markers: {
                    size: 4,
                    discrete: pageXbarR.xBar.violations.map(function (v) {
                        return { seriesIndex: 0, dataPointIndex: v.index, fillColor: '#f43f5e', strokeColor: '#fff', size: 6 };
                    })
                },
                annotations: {
                    points: pageXbarR.xBar.violations.map(function (v) {
                        return {
                            x: pageLabels[v.index], y: pageXbarR.xBar.data[v.index],
                            marker: { size: 0 },
                            label: { borderColor: '#f43f5e', style: { color: '#fff', background: '#f43f5e' }, text: 'OOC' }
                        };
                    })
                },
            };
            var chartX = new ApexCharts(document.querySelector("#xbarChart"), xOptions);
            chartX.render();
            this.chartInstances.push(chartX);
            document.querySelector("#xbarChart").addEventListener('dblclick', function () { chartX.resetSeries(); });

            // R Apex Chart
            var rOptions = {
                chart: { type: 'line', height: 300, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [
                    { name: 'Range', data: pageXbarR.R.data },
                    { name: 'UCL', data: new Array(pageLabels.length).fill(pageXbarR.R.UCL) },
                    { name: 'CL', data: new Array(pageLabels.length).fill(pageXbarR.R.CL) }
                ],
                colors: ['#64748b', '#f43f5e', '#10b981'],
                stroke: { width: [2.5, 1, 1], dashArray: [0, 6, 0] },
                xaxis: { categories: pageLabels, labels: { style: { colors: theme.text, fontSize: '10px' } }, axisBorder: { color: theme.grid }, axisTicks: { color: theme.grid } },
                yaxis: {
                    labels: {
                        formatter: function (val) { return val.toFixed(4); },
                        style: { colors: theme.text }
                    }
                },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode, y: { formatter: function (val) { return val.toFixed(4); } } }
            };
            var chartR = new ApexCharts(document.querySelector("#rChart"), rOptions);
            chartR.render();
            this.chartInstances.push(chartR);
            document.querySelector("#rChart").addEventListener('dblclick', function () { chartR.resetSeries(); });

            this.renderAnomalySidebar(pageXbarR, pageLabels);

        } else if (data.type === 'cavity') {
            var labels = data.cavityStats.map(function (s) { return s.name; });
            var theme = this.getChartTheme(); labels.map(function (s) { return s.name; });
            // Cpk Performance Chart
            var cpkOptions = {
                chart: { type: 'bar', height: 350, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [{ name: 'Cpk', data: cpkVals }],
                xaxis: { categories: labels, labels: { style: { colors: theme.text, fontSize: '11px' } }, axisBorder: { color: theme.grid }, axisTicks: { color: theme.grid } },
                yaxis: { labels: { formatter: function (v) { return v.toFixed(3); }, style: { colors: theme.text } } },
                colors: ['#4f46e5'],
                plotOptions: { bar: { borderRadius: 6, columnWidth: '55%' } },
                dataLabels: { enabled: false },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode, y: { formatter: function (val) { return val.toFixed(3); } } }
            };
            var chartCpk = new ApexCharts(document.querySelector("#cpkChart"), cpkOptions);
            chartCpk.render();
            this.chartInstances.push(chartCpk);
            document.querySelector("#cpkChart").addEventListener('dblclick', function () { chartCpk.resetSeries(); });

            // Mean Comparison Chart (Line)
            var meanOpt = {
                chart: { type: 'line', height: 350, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [
                    { name: this.t('平均值', 'Mean'), data: meanVals },
                    { name: this.t('目標值', 'Target'), data: new Array(labels.length).fill(data.specs.target) },
                    { name: 'USL', data: new Array(labels.length).fill(data.specs.usl) },
                    { name: 'LSL', data: new Array(labels.length).fill(data.specs.lsl) }
                ],
                stroke: { width: [3, 2, 1.5, 1.5], dashArray: [0, 0, 8, 8], curve: 'straight' },
                colors: ['#007bff', '#28a745', '#dc3545', '#dc3545'], // Match image colors
                xaxis: { categories: labels, labels: { style: { colors: theme.text, fontSize: '10px' } }, axisBorder: { color: theme.grid }, axisTicks: { color: theme.grid } },
                yaxis: { labels: { formatter: function (v) { return v.toFixed(4); }, style: { colors: theme.text } } },
                markers: { size: [5, 0, 0, 0] },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode, y: { formatter: function (val) { return val.toFixed(4); } } }
            };
            var chartMean = new ApexCharts(document.querySelector("#meanChart"), meanOpt);
            chartMean.render();
            this.chartInstances.push(chartMean);
            document.querySelector("#meanChart").addEventListener('dblclick', function () { chartMean.resetSeries(); });

            // StdDev Comparison Chart (Line)
            var stdOpt = {
                chart: { type: 'line', height: 350, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [
                    { name: this.t('整體標準差 (s)', 'Overall StdDev'), data: stdOverallVals },
                    { name: this.t('組內標準差 (σ)', 'Within StdDev'), data: stdWithinVals }
                ],
                stroke: { width: 3, curve: 'straight' },
                colors: ['#dc3545', '#007bff'], // Red square-ish vs Blue circle-ish mapping
                xaxis: { categories: labels, labels: { style: { colors: theme.text, fontSize: '10px' } }, axisBorder: { color: theme.grid }, axisTicks: { color: theme.grid } },
                yaxis: { labels: { formatter: function (v) { return v.toFixed(5); }, style: { colors: theme.text } } },
                markers: { size: 5, shape: ['square', 'circle'] },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode, y: { formatter: function (val) { return val.toFixed(5); } } }
            };
            var chartStd = new ApexCharts(document.querySelector("#stdDevChart"), stdOpt);
            chartStd.render();
            this.chartInstances.push(chartStd);
            document.querySelector("#stdDevChart").addEventListener('dblclick', function () { chartStd.resetSeries(); });

        } else if (data.type === 'group') {
            var labels = data.groupStats.map(function (s) { return s.batch; });
            var avgVals = data.groupStats.map(function (s) { return s.avg; });
            var rangeVals = data.groupStats.map(function (s) { return s.range; });
            var minVals = data.groupStats.map(function (s) { return s.min; });
            var maxVals = data.groupStats.map(function (s) { return s.max; });

            var theme = this.getChartTheme();
            // Group Trend Chart - Use Line Chart as requested
            var gOptions = {
                chart: { type: 'line', height: 380, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [
                    { name: 'Max', data: maxVals },
                    { name: 'Average', data: avgVals },
                    { name: 'Min', data: minVals }
                ],
                colors: ['#f43f5e', '#4f46e5', '#10b981'],
                stroke: { width: [1, 3, 1], curve: 'straight' },
                xaxis: { categories: labels, labels: { style: { colors: theme.text, fontSize: '10px' } }, axisBorder: { color: theme.grid }, axisTicks: { color: theme.grid } },
                yaxis: { labels: { formatter: function (v) { return v.toFixed(4); }, style: { colors: theme.text } } },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode, y: { formatter: function (val) { return val.toFixed(4); } } }
            };
            var chartG = new ApexCharts(document.querySelector("#groupChart"), gOptions);
            chartG.render();
            this.chartInstances.push(chartG);
            document.querySelector("#groupChart").addEventListener('dblclick', function () { chartG.resetSeries(); });

            // Inter-Cavity Variation (Range) - Use Line Chart as requested
            var vOptions = {
                chart: { type: 'line', height: 380, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [{ name: 'Range', data: rangeVals }],
                colors: ['#8b5cf6'],
                stroke: { width: 3, curve: 'straight' },
                xaxis: { categories: labels, labels: { style: { colors: theme.text, fontSize: '10px' } }, axisBorder: { color: theme.grid }, axisTicks: { color: theme.grid } },
                yaxis: { labels: { formatter: function (v) { return v.toFixed(4); }, style: { colors: theme.text } } },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode, y: { formatter: function (val) { return val.toFixed(4); } } }
            };
            var chartV = new ApexCharts(document.querySelector("#groupVarChart"), vOptions);
            chartV.render();
            this.chartInstances.push(chartV);
            document.querySelector("#groupVarChart").addEventListener('dblclick', function () { chartV.resetSeries(); });
        }
    },

    downloadExcel: async function () {
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
                        cavRow.push(value !== null ? value : '');
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
            ws['!cols'] = [{ wch: 18 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 8 }];
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

        var now = new Date();
        var ts = now.getFullYear() + String(now.getMonth() + 1).padStart(2, '0') + String(now.getDate()).padStart(2, '0') + '_' +
            String(now.getHours()).padStart(2, '0') + String(now.getMinutes()).padStart(2, '0');

        // Sanitize filename: remove / \ : * ? " < > | and replace spaces with underscore
        var safeItem = String(this.selectedItem || 'Analysis').replace(/[\\\/\:\*\?\"\<\>\|]/g, '').replace(/\s+/g, '_');
        var filename = 'SPC_Report_' + data.type + '_' + safeItem + '_' + ts + '.xlsx';

        console.log('Attempting export: ' + filename);

        if (window.showSaveFilePicker) {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: filename,
                    types: [{
                        description: 'Excel Workbooks',
                        accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
                    }]
                });
                const writable = await handle.createWritable();
                const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
                await writable.write(out);
                await writable.close();
                console.log('Export successful via File System Access API');
            } catch (err) {
                if (err.name !== 'AbortError') {
                    console.error('File picker error:', err);
                    alert(this.t('由於系統權限或瀏覽器限制，將改用預設下載方式。', 'Folder selection failed. Using default download instead.'));
                    XLSX.writeFile(wb, filename);
                }
            }
        } else {
            console.log('showSaveFilePicker not supported, using XLSX.writeFile');
            // Inform user if they are on a non-secure context or incompatible browser
            if (window.location.protocol === 'file:') {
                alert(this.t('提示：偵測到本地執行 (file://)，瀏覽器限制無法主動選擇資料夾。檔案將下載至您的預設下載預設路徑。', 'Note: Local file detected. Browser security prevents folder selection. File will go to your default Downloads.'));
            }
            XLSX.writeFile(wb, filename);
        }
    },

    showLoading: function (text) {
        document.getElementById('loadingText').textContent = text;
        document.getElementById('loadingOverlay').classList.add('active');
    },

    hideLoading: function () {
        document.getElementById('loadingOverlay').classList.remove('active');
    },

    renderAnomalySidebar: function (pageXbarR, pageLabels) {
        var sidebar = document.getElementById('anomalySidebar');
        var list = document.getElementById('anomalyList');
        var self = this;

        list.innerHTML = '';
        var violations = pageXbarR.xBar.violations;

        if (violations.length === 0) {
            list.innerHTML = '<div class="text-center text-slate-500 dark:text-slate-400 mt-10 px-6">' +
                '<p class="text-sm font-medium">' + this.t('本頁數據全數在管制界限內。', 'All points within control limits.') + '</p>' +
                '</div>';
            return;
        }

        violations.forEach(function (v) {
            var batchName = pageLabels[v.index] || 'Batch ' + (v.index + 1);
            var card = document.createElement('div');
            card.className = 'bg-white dark:bg-slate-800 border border-slate-100 dark:border-slate-700 p-5 rounded-xl shadow-sm hover:border-rose-200 dark:hover:border-rose-700 hover:shadow-md transition-all cursor-pointer group mb-3 mx-4';

            var rulesText = v.rules.map(function (r) { return 'Rule ' + r; }).join(', ');

            card.innerHTML =
                '<div class="flex items-start gap-4">' +
                '<div class="w-8 h-8 rounded-lg bg-rose-50 dark:bg-rose-900/30 flex items-center justify-center flex-shrink-0">' +
                '<span class="material-icons-outlined text-rose-500 text-lg">error_outline</span>' +
                '</div>' +
                '<div class="flex-1">' +
                '<div class="text-sm font-bold text-slate-900 dark:text-white mb-0.5">' + batchName + '</div>' +
                '<div class="text-[10px] text-rose-500 font-bold uppercase tracking-wider mb-2">' + rulesText + '</div>' +
                '<div class="text-[11px] text-slate-500 dark:text-slate-400 font-medium leading-relaxed">' +
                self.t('偵測到異常偏離，建議檢查該批次生產條件。', 'Deviation detected. Verify production settings for this batch.') +
                '</div>' +
                '</div>' +
                '</div>';

            list.appendChild(card);
        });
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
