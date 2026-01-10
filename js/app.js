/**
 * SPC Analysis Tool - Main Application Logic
 * Orchestrates UI, Events, and Analysis Flow.
 */

var SPCApp = {
    currentLanguage: 'zh',
    workbook: null,
    selectedItem: null,
    analysisResults: null,
    chartInstances: [],
    batchPagination: { currentPage: 1, totalPages: 1, maxPerPage: 25, totalBatches: 0 },

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
        var langBtn = document.getElementById('langBtn');
        if (langBtn) {
            langBtn.addEventListener('click', function () {
                self.currentLanguage = self.currentLanguage === 'zh' ? 'en' : 'zh';
                var langText = document.getElementById('langText');
                if (langText) langText.textContent = self.currentLanguage === 'zh' ? 'EN' : '中文';
                self.updateLanguage();
            });
        }
    },

    updateLanguage: function () {
        var self = this;
        var elements = document.querySelectorAll('[data-en][data-zh]');
        elements.forEach(function (el) {
            el.textContent = self.currentLanguage === 'zh' ? el.dataset.zh : el.dataset.en;
        });
    },

    setupFileUpload: function () {
        var self = this;
        var uploadZone = document.getElementById('uploadZone');
        var fileInput = document.getElementById('fileInput');

        if (!uploadZone || !fileInput) return;

        uploadZone.addEventListener('click', function (e) {
            if (e.target.tagName === 'BUTTON' || e.target.closest('button')) return;
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

    getChartTheme: function () {
        var isDark = document.documentElement.classList.contains('dark');
        return {
            mode: isDark ? 'dark' : 'light',
            text: isDark ? '#94a3b8' : '#64748b',
            grid: isDark ? '#334155' : '#f1f5f9'
        };
    },

    history: [],

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
        if (!itemList) return;
        itemList.innerHTML = '';

        var sheets = this.workbook.SheetNames;
        sheets.forEach(function (sheetName) {
            var ws = self.workbook.Sheets[sheetName];
            var data = XLSX.utils.sheet_to_json(ws, { header: 1 });
            var target = data[1] && data[1][1] ? data[1][1] : 'N/A';

            var card = document.createElement('div');
            card.className = 'saas-card p-6 cursor-pointer hover:border-primary hover:shadow-xl hover:shadow-indigo-50 dark:hover:shadow-none transition-all group relative overflow-hidden';
            card.innerHTML =
                '<div class="relative z-10">' +
                '<div class="flex justify-between items-start mb-4">' +
                '<h3 class="text-base font-bold text-slate-900 dark:text-white group-hover:text-primary transition-colors">' + sheetName + '</h3>' +
                '<span class="material-icons-outlined text-slate-400 group-hover:text-primary transition-colors">check_circle_outline</span>' +
                '</div>' +
                '<div class="flex flex-col space-y-2">' +
                '<div class="text-[10px] text-slate-400 font-bold uppercase tracking-widest">' + self.t('目標值', 'Target Value') + '</div>' +
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
        });

        document.getElementById('step2').style.display = 'block';
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
        document.querySelectorAll('[data-analysis]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                self.executeAnalysis(this.dataset.analysis);
            });
        });

        var clearRecentBtn = document.getElementById('clearRecentBtn');
        if (clearRecentBtn) clearRecentBtn.addEventListener('click', function () { self.clearHistory(); });

        var navIds = ['nav-dashboard', 'nav-import', 'nav-analysis', 'nav-history', 'nav-qip-extract', 'nav-settings'];
        navIds.forEach(function (id) {
            var el = document.getElementById(id);
            if (el) el.addEventListener('click', function (e) { e.preventDefault(); self.switchView(id.replace('nav-', '')); });
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
                document.querySelectorAll('#itemList .saas-card').forEach(function (card) {
                    card.style.display = card.textContent.toLowerCase().includes(term) ? 'block' : 'none';
                });
                document.querySelectorAll('#historyTableBody tr').forEach(function (row) {
                    row.style.display = row.textContent.toLowerCase().includes(term) ? 'table-row' : 'none';
                });
            });
        }
    },

    setupDarkModeToggle: function () {
        var self = this;
        var btn = document.getElementById('darkModeBtn');
        if (btn) {
            btn.addEventListener('click', function () {
                setTimeout(function () {
                    if (self.analysisResults) self.renderCharts();
                }, 100);
            });
        }
    },

    switchView: function (viewId) {
        var self = this;
        var viewMap = { 'dashboard': 'view-import', 'import': 'view-import', 'analysis': 'view-analysis', 'history': 'view-history', 'qip-extract': 'view-qip-extract', 'settings': 'view-import' };
        var targetId = viewMap[viewId] || 'view-import';

        ['view-import', 'view-analysis', 'view-history', 'view-qip-extract'].forEach(function (id) {
            var el = document.getElementById(id);
            if (el) el.classList.add('hidden');
        });

        var targetEl = document.getElementById(targetId);
        if (targetEl) targetEl.classList.remove('hidden');

        if (viewId === 'history') this.renderHistoryView();

        // Check for extracted QIP data when switching to import view
        if ((viewId === 'import' || viewId === 'dashboard') && window.qipExtractedData) {
            try {
                this.loadExtractedData(window.qipExtractedData);
                window.qipExtractedData = null; // Consume data
            } catch (e) {
                console.error('Failed to load extracted data', e);
                alert('載入提取數據失敗: ' + e.message);
            }
        }

        document.querySelectorAll('#main-nav .nav-link').forEach(function (link) {
            link.className = 'nav-link flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium text-slate-400 hover:text-white hover:bg-slate-800 transition-all';
        });

        var activeLink = document.getElementById('nav-' + viewId);
        if (activeLink) activeLink.className = 'nav-link flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium bg-primary text-white shadow-lg shadow-primary/20';

        var titleEl = document.querySelector('header .text-slate-900');
        if (titleEl) {
            var titles = { 'dashboard': this.t('數據總覽', 'Dashboard'), 'import': this.t('QIP 數據導入', 'QIP Import'), 'analysis': this.t('統計分析結果', 'Analysis Result'), 'history': this.t('歷史分析紀錄', 'History'), 'qip-extract': this.t('QIP 數據提取', 'QIP Data Extract'), 'settings': this.t('系統設定', 'Settings') };
            titleEl.innerText = titles[viewId] || 'SPC Analysis';
        }

        var side = document.getElementById('anomalySidebar');
        if (side) {
            if (viewId === 'analysis' && this.analysisResults) side.classList.remove('hidden');
            else side.classList.add('hidden');
        }
    },

    // Convert QIP Extracted JSON to Virtual Workbook for SPC
    loadExtractedData: function (extracted) {
        console.log('Loading extracted data into SPC...', extracted);
        var self = this;
        var wb = XLSX.utils.book_new();

        // Group by batch (sheet name)
        var batches = {};
        extracted.data.forEach(function (item) {
            if (!batches[item.batch]) batches[item.batch] = [];
            batches[item.batch].push(item);
        });

        // Create a sheet for each batch
        Object.keys(batches).forEach(function (batchName) {
            var items = batches[batchName];
            // Sort by cavity ID (assuming numeric if possible)
            items.sort(function (a, b) {
                var na = parseFloat(a.cavity);
                var nb = parseFloat(b.cavity);
                if (!isNaN(na) && !isNaN(nb)) return na - nb;
                return a.cavity.localeCompare(b.cavity);
            });

            // Construct AOA (Array of Arrays)
            // Row 1: Headers (Cavity IDs)
            // Row 2: Values
            var headers = items.map(function (i) { return i.cavity; });
            var values = items.map(function (i) { return i.value; });

            // Add some metadata headers if standard logic expects specific row?
            // Existing logic in DataInput parses headers (cavity ID) and data. 
            // Usually assumes Layout: Cavity IDs in one row, Data in subsequent rows.
            // Let's create a standard clean sheet.

            var aoa = [headers, values];
            var ws = XLSX.utils.aoa_to_sheet(aoa);
            XLSX.utils.book_append_sheet(wb, ws, batchName);
        });

        this.workbook = wb;
        this.renderItemList();

        // Update UI state to show loaded
        var fileInfo = document.getElementById('file-info');
        var uploadArea = document.getElementById('upload-area');
        if (fileInfo && uploadArea) {
            fileInfo.classList.remove('hidden');
            uploadArea.classList.add('hidden');
            document.getElementById('filename').textContent = 'QIP_Extracted_' + (extracted.productCode || 'Data');
        }

        alert(this.t('提取數據已載入，共 ' + extracted.data.length + ' 筆。請選擇項目進行分析。', 'Extracted data loaded. Please select an item to analyze.'));
    },

    executeAnalysis: function (type) {
        var self = this;
        if (!this.selectedItem) { alert(this.t('請先選擇分析項目', 'Please select an item first')); return; }
        this.showLoading(this.t('分析中...', 'Analyzing...'));

        this.chartInstances.forEach(function (c) { if (c.destroy) c.destroy(); });
        this.chartInstances = [];

        setTimeout(function () {
            try {
                var ws = self.workbook.Sheets[self.selectedItem];
                var dataInput = new DataInput(ws);
                var results;

                if (type === 'batch') {
                    var dataMatrix = dataInput.getDataMatrix();
                    var xbarR = SPCEngine.calculateXBarRLimits(dataMatrix);
                    var allValues = dataMatrix.flat().filter(function (v) { return v !== null; });
                    var specs = dataInput.specs;
                    var cap = SPCEngine.calculateProcessCapability(allValues, specs.usl, specs.lsl);
                    xbarR.summary.Cpk = cap.Cpk;
                    xbarR.summary.Ppk = cap.Ppk;

                    results = { type: 'batch', xbarR: xbarR, batchNames: dataInput.batchNames, specs: specs, dataMatrix: dataMatrix, cavityNames: dataInput.getCavityNames(), productInfo: dataInput.productInfo };
                } else if (type === 'cavity') {
                    var specs = dataInput.specs;
                    var cavityStats = [];
                    for (var i = 0; i < dataInput.getCavityCount(); i++) {
                        var cap = SPCEngine.calculateProcessCapability(dataInput.getCavityBatchData(i), specs.usl, specs.lsl);
                        cap.name = dataInput.getCavityNames()[i];
                        cavityStats.push(cap);
                    }
                    results = { type: 'cavity', cavityStats: cavityStats, specs: specs, productInfo: dataInput.productInfo };
                } else if (type === 'group') {
                    var specs = dataInput.specs;
                    var dataMatrix = dataInput.getDataMatrix();
                    var groupStats = dataMatrix.map(function (row, i) {
                        var filtered = row.filter(function (v) { return v !== null && !isNaN(v); });
                        return { batch: dataInput.batchNames[i] || 'B' + (i + 1), avg: SPCEngine.mean(filtered), max: SPCEngine.max(filtered), min: SPCEngine.min(filtered), range: SPCEngine.range(filtered), count: filtered.length };
                    });
                    results = { type: 'group', groupStats: groupStats, specs: specs, productInfo: dataInput.productInfo };
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
        var data = this.analysisResults;
        var self = this;
        var html = '';

        if (data.type === 'batch') {
            var totalBatches = Math.min(data.batchNames.length, data.xbarR.xBar.data.length);
            this.batchPagination = { currentPage: 1, totalPages: Math.ceil(totalBatches / 25), maxPerPage: 25, totalBatches: totalBatches };

            html = '<div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('模穴數', 'Cavities') + '</div> <div class="text-xl font-bold dark:text-white">' + data.xbarR.summary.n + '</div> </div>' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('總記錄', 'Total') + '</div> <div class="text-xl font-bold dark:text-white">' + totalBatches + '</div> </div>' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('製程能力 (Cpk)', 'Cpk') + '</div> <div class="text-xl font-bold text-primary">' + SPCEngine.round(data.xbarR.summary.Cpk, 3) + '</div> </div>' +
                '<div class="saas-card p-4"> <div class="text-[10px] font-bold text-slate-500 uppercase">' + this.t('狀態', 'Status') + '</div> <div class="flex items-center gap-2">' +
                '<span class="material-icons-outlined text-sm ' + (data.xbarR.xBar.violations.length > 0 ? 'text-rose-500' : 'text-emerald-500') + '">' + (data.xbarR.xBar.violations.length > 0 ? 'warning' : 'check_circle') + '</span>' +
                '<span class="text-sm font-semibold ' + (data.xbarR.xBar.violations.length > 0 ? 'text-rose-600' : 'text-emerald-600') + '">' + (data.xbarR.xBar.violations.length > 0 ? this.t('異常', 'Alert') : this.t('正常', 'Normal')) + '</span> </div> </div> </div>';

            if (this.batchPagination.totalPages > 1) {
                html += '<div class="flex items-center justify-between mb-4 bg-white dark:bg-slate-800 p-3 rounded-lg border border-slate-100 dark:border-slate-700 shadow-sm">' +
                    '<div class="text-xs font-medium dark:text-slate-400" id="pageInfo">' + this.t('第 1 / ', 'Page 1 / ') + this.batchPagination.totalPages + '</div>' +
                    '<div class="flex gap-2">' +
                    '<button id="prevPageBtn" class="px-4 py-1 pb-1.5 text-xs font-bold border rounded-md disabled:opacity-30 dark:text-white">' + this.t('上頁', 'Prev') + '</button>' +
                    '<button id="nextPageBtn" class="px-4 py-1 pb-1.5 text-xs font-bold text-white bg-primary rounded-md disabled:opacity-30">' + this.t('下頁', 'Next') + '</button> </div> </div>';
            }
            html += '<div id="detailedTableContainer" class="mb-10 overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm"></div>' +
                '<div class="grid grid-cols-1 gap-8">' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('X̄ 管制圖', 'X-Bar Chart') + '</h3> <div id="xbarChart" class="h-96"></div> </div>' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('R 管制圖', 'R Chart') + '</h3> <div id="rChart" class="h-80"></div> </div> </div>';
        } else if (data.type === 'cavity') {
            html = '<div class="grid grid-cols-1 gap-8">' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('模穴 Cpk 效能比較', 'Cavity Cpk') + '</h3> <div id="cpkChart" class="h-96"></div> </div>' +
                '<div class="grid grid-cols-1 md:grid-cols-2 gap-6">' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('均值比較', 'Mean Comp') + '</h3> <div id="meanChart" class="h-80"></div> </div>' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('標差比較', 'StdDev Comp') + '</h3> <div id="stdDevChart" class="h-80"></div> </div> </div>' +
                '<div class="saas-card overflow-hidden"> <div class="p-6 border-b dark:border-slate-700"> <h3 class="text-base font-bold dark:text-white">' + this.t('數據明細', 'Details') + '</h3> </div>' +
                '<table class="w-full text-sm text-left"> <thead class="text-xs text-slate-400 bg-slate-50/50 dark:bg-slate-700/50 uppercase"> <tr><th class="px-6 py-3">Name</th><th class="px-6 py-3 text-center">Mean</th><th class="px-6 py-3 text-center">Cpk</th><th class="px-6 py-3 text-center">n</th></tr> </thead>' +
                '<tbody class="divide-y dark:divide-slate-700">' + data.cavityStats.map(function (s) {
                    return '<tr> <td class="px-6 py-4 font-bold dark:text-slate-300">' + s.name + '</td> <td class="px-6 py-4 text-center font-mono dark:text-slate-300">' + SPCEngine.round(s.mean, 4) + '</td> <td class="px-6 py-4 text-center">' +
                        '<span class="px-2 py-0.5 rounded-full text-[10px] font-bold ' + (s.Cpk < 1.33 ? 'bg-rose-50 text-rose-600' : 'bg-emerald-50 text-emerald-600') + '">' + SPCEngine.round(s.Cpk, 3) + '</span></td> <td class="px-6 py-4 text-center dark:text-slate-400">' + s.count + '</td> </tr>';
                }).join('') + '</tbody> </table> </div> </div>';
        } else if (data.type === 'group') {
            html = '<div class="grid grid-cols-1 gap-8">' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('群組趨勢 (Min-Max-Avg)', 'Trend') + '</h3> <div id="groupChart" class="h-96"></div> </div>' +
                '<div class="saas-card p-8"> <h3 class="text-base font-bold mb-6 dark:text-white">' + this.t('組間全距 (Variation)', 'Variation') + '</h3> <div id="groupVarChart" class="h-96"></div> </div> </div>';
        }

        resultsContent.innerHTML = html;
        document.getElementById('results').style.display = 'block';

        var downloadBtn = document.getElementById('downloadExcel');
        if (downloadBtn) {
            var newBtn = downloadBtn.cloneNode(true);
            downloadBtn.parentNode.replaceChild(newBtn, downloadBtn);
            newBtn.addEventListener('click', function () { self.downloadExcel(); });
        }

        if (data.type === 'batch' && this.batchPagination.totalPages > 1) {
            document.getElementById('prevPageBtn').addEventListener('click', function () { self.changeBatchPage(-1); });
            document.getElementById('nextPageBtn').addEventListener('click', function () { self.changeBatchPage(1); });
            this.updatePaginationButtons();
        }

        setTimeout(function () { self.renderCharts(); self.switchView('analysis'); document.getElementById('results').scrollIntoView({ behavior: 'smooth' }); }, 100);
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
        var info = document.getElementById('pageInfo');
        if (info) info.textContent = this.t('第 ', 'Page ') + p.currentPage + ' / ' + p.totalPages + this.t(' 頁', '');
        var prev = document.getElementById('prevPageBtn');
        var next = document.getElementById('nextPageBtn');
        if (prev) prev.disabled = (p.currentPage <= 1);
        if (next) next.disabled = (p.currentPage >= p.totalPages);
    },

    renderDetailedDataTable: function (pageLabels, pageDataMatrix, pageXbarR) {
        var data = this.analysisResults;
        var info = data.productInfo;
        var specs = data.specs;
        var cavityCount = data.xbarR.summary.n;
        var totalWidth = 60 + (25 * 58) + 120;

        var html = '<table class="excel-table" style="width:' + totalWidth + 'px; border-collapse:collapse; font-size:10px; font-family:sans-serif; border:2px solid var(--table-border); table-layout:fixed;">';
        html += '<colgroup> <col style="width:60px;"> ';
        for (var c = 0; c < 25; c++) html += '<col style="width:58px;">';
        html += '<col style="width:30px;"><col style="width:30px;"><col style="width:30px;"><col style="width:30px;"> </colgroup>';

        html += '<tr style="background:var(--table-header-bg); text-align:center;"><td colspan="30" style="border:1px solid var(--table-border); font-weight:bold; font-size:14px; padding:3px;">X̄ - R 管制圖</td></tr>';

        var meta = [
            { l1: '商品名稱', v1: info.name, l2: '規格', v2: '標準', l3: '管制圖', v3: 'X̄', v4: 'R', l4: '製造部門', v4_val: info.dept },
            { l1: '商品料號', v1: info.item, l2: '最大值', v2: SPCEngine.round(specs.usl, 4), l3: '上限', v3: SPCEngine.round(pageXbarR.xBar.UCL, 4), v4: SPCEngine.round(pageXbarR.R.UCL, 4), l4: '檢驗人員', v4_val: info.inspector },
            { l1: '測量單位', v1: info.unit, l2: '目標值', v2: SPCEngine.round(specs.target, 4), l3: '中心值', v3: SPCEngine.round(pageXbarR.xBar.CL, 4), v4: SPCEngine.round(pageXbarR.R.CL, 4), l4: '管制特性', v4_val: info.char },
            { l1: '檢驗日期', v1: info.batchRange || '-', l2: '最小值', v2: SPCEngine.round(specs.lsl, 4), l3: '下限', v3: SPCEngine.round(pageXbarR.xBar.LCL, 4), v4: '-', l4: '圖表編號', v4_val: info.chartNo || '-' }
        ];

        meta.forEach(function (r) {
            html += '<tr>' +
                '<td colspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); padding:1px;">' + r.l1 + '</td>' +
                '<td colspan="7" style="border:1px solid var(--table-border); padding:1px;">' + (r.v1 || '') + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); padding:1px;">' + r.l2 + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); padding:1px;">' + (r.v2 || '') + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); padding:1px;">' + r.l3 + '</td>' +
                '<td colspan="3" style="border:1px solid var(--table-border); padding:1px;">' + (r.v3 || '') + '</td>' +
                '<td colspan="3" style="border:1px solid var(--table-border); padding:1px;">' + (r.v4 || '') + '</td>' +
                '<td colspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg); padding:1px;">' + r.l4 + '</td>' +
                '<td colspan="7" style="border:1px solid var(--table-border); padding:1px;">' + (r.v4_val || '') + '</td>' +
                '</tr>';
        });

        html += '<tr style="background:var(--table-data-header-bg); font-weight:bold; text-align:center;">' +
            '<td style="border:1px solid var(--table-border);">批號</td>';
        for (var b = 0; b < 25; b++) {
            var name = pageLabels[b] || '';
            html += '<td style="border:1px solid var(--table-border); height:35px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">' + name + '</td>';
        }
        html += '<td colspan="4" style="border:1px solid var(--table-border);">彙總</td></tr>';

        for (var i = 0; i < cavityCount; i++) {
            html += '<tr style="text-align:center;"><td style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-label-bg);">X' + (i + 1) + '</td>';
            for (var j = 0; j < 25; j++) {
                var val = (pageDataMatrix[j] && pageDataMatrix[j][i] !== undefined) ? pageDataMatrix[j][i] : '';
                html += '<td style="border:1px solid var(--table-border); background:var(--table-bg);">' + val + '</td>';
            }
            if (i === 0) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-bg);">ΣX̄ = ' + SPCEngine.round(pageXbarR.summary.xDoubleBar * pageXbarR.summary.k, 4) + '</td>';
            else if (i === 2) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-bg);">X̿ = ' + SPCEngine.round(pageXbarR.summary.xDoubleBar, 4) + '</td>';
            else if (i === 4) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-bg);">ΣR = ' + SPCEngine.round(pageXbarR.summary.rBar * pageXbarR.summary.k, 4) + '</td>';
            else if (i === 6) html += '<td colspan="4" rowspan="2" style="border:1px solid var(--table-border); font-weight:bold; background:var(--table-bg);">R̄ = ' + SPCEngine.round(pageXbarR.summary.rBar, 4) + '</td>';
            else if (i >= 8) html += '<td colspan="4" style="border:1px solid var(--table-border); background:transparent;"></td>';
            html += '</tr>';
        }

        // Footer Rows
        ['ΣX', 'X̄', 'R'].forEach(function (type) {
            html += '<tr style="background:var(--table-header-bg); text-align:center;"><td style="border:1px solid var(--table-border); font-weight:bold;">' + type + '</td>';
            for (var b = 0; b < 25; b++) {
                var val = '', style = '';
                if (pageDataMatrix[b]) {
                    if (type === 'ΣX') val = SPCEngine.round(pageDataMatrix[b].reduce(function (a, b) { return a + (b || 0); }, 0), 4);
                    else if (type === 'X̄' && pageXbarR.xBar.data[b] !== undefined) {
                        var v = pageXbarR.xBar.data[b]; val = SPCEngine.round(v, 4);
                        if (v > pageXbarR.xBar.UCL || v < pageXbarR.xBar.LCL) style = 'background:var(--table-ooc-bg); color:var(--table-ooc-text);';
                    }
                    else if (type === 'R' && pageXbarR.R.data[b] !== undefined) {
                        var v = pageXbarR.R.data[b]; val = SPCEngine.round(v, 4);
                        if (v > pageXbarR.R.UCL) style = 'background:var(--table-ooc-bg); color:var(--table-ooc-text);';
                    }
                }
                html += '<td style="border:1px solid var(--table-border); ' + style + '">' + val + '</td>';
            }
            html += '<td colspan="4" style="border:1px solid var(--table-border);"></td></tr>';
        });

        html += '</table>';
        document.getElementById('detailedTableContainer').innerHTML = html;
    },

    renderCharts: function () {
        var self = this;
        var data = this.analysisResults;
        this.chartInstances.forEach(function (c) { if (c.destroy) c.destroy(); });
        this.chartInstances = [];

        var theme = this.getChartTheme();

        if (data.type === 'batch') {
            var p = this.batchPagination;
            var start = (p.currentPage - 1) * p.maxPerPage;
            var end = Math.min(start + p.maxPerPage, p.totalBatches);
            var pageLabels = data.batchNames.slice(start, end);
            var pageData = data.dataMatrix.slice(start, end);
            var pageXbarR = SPCEngine.calculateXBarRLimits(pageData);

            this.renderDetailedDataTable(pageLabels, pageData, pageXbarR);

            var xOpt = {
                chart: { type: 'line', height: 380, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [
                    { name: 'X-Bar', data: pageXbarR.xBar.data },
                    { name: 'UCL', data: new Array(pageLabels.length).fill(pageXbarR.xBar.UCL) },
                    { name: 'CL', data: new Array(pageLabels.length).fill(pageXbarR.xBar.CL) },
                    { name: 'LCL', data: new Array(pageLabels.length).fill(pageXbarR.xBar.LCL) }
                ],
                colors: ['#4f46e5', '#f43f5e', '#10b981', '#f43f5e'],
                stroke: { width: [3, 1.5, 1.5, 1.5], dashArray: [0, 6, 0, 6] },
                xaxis: { categories: pageLabels, labels: { style: { colors: theme.text, fontSize: '10px' } } },
                yaxis: { labels: { formatter: function (v) { return v.toFixed(4); }, style: { colors: theme.text } } },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode },
                markers: { size: 4, discrete: pageXbarR.xBar.violations.map(function (v) { return { seriesIndex: 0, dataPointIndex: v.index, fillColor: '#f43f5e', strokeColor: '#fff', size: 6 }; }) }
            };
            var chartX = new ApexCharts(document.querySelector("#xbarChart"), xOpt);
            chartX.render(); this.chartInstances.push(chartX);

            var rOpt = {
                chart: { type: 'line', height: 300, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [
                    { name: 'Range', data: pageXbarR.R.data },
                    { name: 'UCL', data: new Array(pageLabels.length).fill(pageXbarR.R.UCL) },
                    { name: 'CL', data: new Array(pageLabels.length).fill(pageXbarR.R.CL) }
                ],
                colors: ['#64748b', '#f43f5e', '#10b981'],
                stroke: { width: [2.5, 1, 1], dashArray: [0, 6, 0] },
                xaxis: { categories: pageLabels, labels: { style: { colors: theme.text, fontSize: '10px' } } },
                yaxis: { labels: { formatter: function (v) { return v.toFixed(4); }, style: { colors: theme.text } } },
                grid: { borderColor: theme.grid },
                tooltip: { theme: theme.mode }
            };
            var chartR = new ApexCharts(document.querySelector("#rChart"), rOpt);
            chartR.render(); this.chartInstances.push(chartR);

            this.renderAnomalySidebar(pageXbarR, pageLabels);

        } else if (data.type === 'cavity') {
            var labels = data.cavityStats.map(s => s.name);
            var cpkVal = data.cavityStats.map(s => s.Cpk);

            var cpkOpt = {
                chart: { type: 'bar', height: 350, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [{ name: 'Cpk', data: cpkVal }],
                xaxis: { categories: labels, labels: { style: { colors: theme.text } } },
                colors: ['#4f46e5'],
                plotOptions: { bar: { borderRadius: 6 } },
                grid: { borderColor: theme.grid }
            };
            var chartCpk = new ApexCharts(document.querySelector("#cpkChart"), cpkOpt);
            chartCpk.render(); this.chartInstances.push(chartCpk);

            var meanOpt = {
                chart: { type: 'line', height: 350, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [
                    { name: this.t('平均值', 'Mean'), data: data.cavityStats.map(s => s.mean) },
                    { name: 'USL', data: new Array(labels.length).fill(data.specs.usl) },
                    { name: 'LSL', data: new Array(labels.length).fill(data.specs.lsl) }
                ],
                colors: ['#007bff', '#dc3545', '#dc3545'],
                xaxis: { categories: labels, labels: { style: { colors: theme.text } } },
                grid: { borderColor: theme.grid }
            };
            var chartMean = new ApexCharts(document.querySelector("#meanChart"), meanOpt);
            chartMean.render(); this.chartInstances.push(chartMean);

            var stdOpt = {
                chart: { type: 'line', height: 350, toolbar: { show: false }, background: 'transparent' },
                theme: { mode: theme.mode },
                series: [
                    { name: 'Overall σ', data: data.cavityStats.map(s => s.overallStdDev) },
                    { name: 'Within σ', data: data.cavityStats.map(s => s.withinStdDev) }
                ],
                colors: ['#dc3545', '#007bff'],
                xaxis: { categories: labels, labels: { style: { colors: theme.text } } },
                grid: { borderColor: theme.grid }
            };
            var chartStd = new ApexCharts(document.querySelector("#stdDevChart"), stdOpt);
            chartStd.render(); this.chartInstances.push(chartStd);

        } else if (data.type === 'group') {
            var labels = data.groupStats.map(s => s.batch);
            var gOpt = {
                chart: { type: 'line', height: 380, toolbar: { show: false } },
                theme: { mode: theme.mode },
                series: [
                    { name: 'Max', data: data.groupStats.map(s => s.max) },
                    { name: 'Avg', data: data.groupStats.map(s => s.avg) },
                    { name: 'Min', data: data.groupStats.map(s => s.min) }
                ],
                colors: ['#f43f5e', '#4f46e5', '#10b981'],
                xaxis: { categories: labels, labels: { style: { colors: theme.text } } },
                grid: { borderColor: theme.grid }
            };
            var chartG = new ApexCharts(document.querySelector("#groupChart"), gOpt);
            chartG.render(); this.chartInstances.push(chartG);

            var vOpt = {
                chart: { type: 'line', height: 380, toolbar: { show: false } },
                theme: { mode: theme.mode },
                series: [{ name: 'Range', data: data.groupStats.map(s => s.range) }],
                colors: ['#8b5cf6'],
                xaxis: { categories: labels, labels: { style: { colors: theme.text } } },
                grid: { borderColor: theme.grid }
            };
            var chartV = new ApexCharts(document.querySelector("#groupVarChart"), vOpt);
            chartV.render(); this.chartInstances.push(chartV);
        }
    },

    downloadExcel: async function () {
        var self = this;
        var data = this.analysisResults;
        var wb = XLSX.utils.book_new();
        var maxBatchesPerSheet = 25;

        if (data.type === 'batch') {
            var totalBatches = Math.min(data.batchNames.length, data.xbarR.xBar.data.length);
            var totalSheets = Math.ceil(totalBatches / maxBatchesPerSheet);
            var cavityCount = data.productInfo.n || data.xbarR.summary.n;

            for (var sheetIdx = 0; sheetIdx < totalSheets; sheetIdx++) {
                var startBatch = sheetIdx * maxBatchesPerSheet;
                var endBatch = Math.min((sheetIdx + 1) * maxBatchesPerSheet, totalBatches);
                var batchCount = endBatch - startBatch;
                var wsData = [];

                wsData.push([this.t('X̄ - R 管制圖', 'X-Bar R Chart')]);
                wsData.push([this.t('產品名稱', 'Product'), data.productInfo.name, '', '', this.t('規格', 'Spec'), '', '', this.t('管制界限', 'Limits')]);
                wsData.push([this.t('檢驗項目', 'Item'), this.selectedItem, '', '', 'Target', data.specs.target, '', 'UCL', SPCEngine.round(data.xbarR.xBar.UCL, 4)]);
                wsData.push([this.t('模穴數', 'Cavities'), cavityCount, '', '', 'USL', data.specs.usl, '', 'CL', SPCEngine.round(data.xbarR.xBar.CL, 4)]);
                wsData.push([this.t('批號數', 'Batches'), data.xbarR.summary.k, '', '', 'LSL', data.specs.lsl, '', 'LCL', SPCEngine.round(data.xbarR.xBar.LCL, 4)]);
                wsData.push([]);

                var headerRow = [this.t('檢驗批號', 'Batch No.')];
                for (var b = startBatch; b < endBatch; b++) headerRow.push(data.batchNames[b]);
                wsData.push(headerRow);

                for (var cav = 0; cav < cavityCount; cav++) {
                    var cavRow = ['X' + (cav + 1)];
                    for (var b = startBatch; b < endBatch; b++) cavRow.push(data.dataMatrix[b][cav]);
                    wsData.push(cavRow);
                }
                var xbarRow = ['X̄']; for (var b = startBatch; b < endBatch; b++) xbarRow.push(data.xbarR.xBar.data[b]); wsData.push(xbarRow);
                var rRow = ['R']; for (var b = startBatch; b < endBatch; b++) rRow.push(data.xbarR.R.data[b]); wsData.push(rRow);

                var ws = XLSX.utils.aoa_to_sheet(wsData);
                XLSX.utils.book_append_sheet(wb, ws, 'P' + (sheetIdx + 1));
            }
        } else if (data.type === 'cavity') {
            var wsData = [[this.t('模穴分析', 'Cavity Analysis')], ['Cavity', 'Mean', 'Cpk', 'n']];
            data.cavityStats.forEach(s => wsData.push([s.name, s.mean, s.Cpk, s.count]));
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), 'Cavity');
        }

        var filename = 'SPC_Report_' + data.type + '.xlsx';
        XLSX.writeFile(wb, filename);
    },

    showLoading: function (text) {
        var overlay = document.getElementById('loadingOverlay');
        var textEl = document.getElementById('loadingText');
        if (textEl) textEl.textContent = text;
        if (overlay) overlay.style.display = 'flex';
    },

    hideLoading: function () {
        var overlay = document.getElementById('loadingOverlay');
        if (overlay) overlay.style.display = 'none';
    },

    renderAnomalySidebar: function (pageXbarR, pageLabels) {
        var list = document.getElementById('anomalyList');
        if (!list) return;
        list.innerHTML = '';
        var violations = pageXbarR.xBar.violations;
        if (violations.length === 0) {
            list.innerHTML = '<div class="text-center text-slate-500 mt-10 px-6">' + this.t('本頁數據全數受控', 'All points within control limits.') + '</div>';
            return;
        }
        violations.forEach(function (v) {
            var card = document.createElement('div');
            card.className = 'bg-white dark:bg-slate-800 border p-4 rounded-xl mb-3 mx-4';
            card.innerHTML = '<div class="text-sm font-bold dark:text-white">' + (pageLabels[v.index] || 'Batch') + '</div>' +
                '<div class="text-[10px] text-rose-500 uppercase font-bold">' + v.rules.map(r => 'Rule ' + r).join(', ') + '</div>';
            list.appendChild(card);
        });
    },

    resetApp: function () {
        this.workbook = null; this.selectedItem = null; this.analysisResults = null;
        this.chartInstances.forEach(c => { if (c.destroy) c.destroy(); });
        this.chartInstances = [];
        document.getElementById('fileInput').value = '';
        document.getElementById('uploadZone').style.display = 'block';
        document.getElementById('fileInfo').style.display = 'none';
        ['step2', 'step3', 'results'].forEach(id => { var el = document.getElementById(id); if (el) el.style.display = 'none'; });
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }
};

document.addEventListener('DOMContentLoaded', function () { SPCApp.init(); });
