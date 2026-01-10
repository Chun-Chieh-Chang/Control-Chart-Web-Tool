/**
 * QIP Data Extract - Application Controller
 * Integrated into ControlChart with SPC data flow.
 */

var QIPExtractApp = {
    workbook: null,
    fileName: '',
    processingResults: null,
    selectionMode: null,
    selectionTarget: null,
    selectionStart: null,
    selectionEnd: null,

    init: function () {
        console.log('QIPExtractApp initializing...');
        this.cacheElements();
        this.bindEvents();
        this.loadSavedConfigs();
    },

    cacheElements: function () {
        this.els = {
            uploadZone: document.getElementById('qip-upload-zone'),
            fileInput: document.getElementById('qip-file-input'),
            fileInfo: document.getElementById('qip-file-info'),
            fileName: document.getElementById('qip-file-name'),
            removeFile: document.getElementById('qip-remove-file'),
            workbookInfo: document.getElementById('qip-workbook-info'),

            productCode: document.getElementById('qip-product-code'),
            cavityCount: document.getElementById('qip-cavity-count'),
            worksheetSelectGroup: document.getElementById('qip-worksheet-select-group'),
            worksheetSelect: document.getElementById('qip-worksheet-select'),
            previewBtn: document.getElementById('qip-preview-btn'),
            cavityGroups: document.getElementById('qip-cavity-groups'),

            configName: document.getElementById('qip-config-name'),
            saveConfig: document.getElementById('qip-save-config'),
            loadConfig: document.getElementById('qip-load-config'),

            startProcess: document.getElementById('qip-start-process'),
            progress: document.getElementById('qip-progress'),
            progressBar: document.getElementById('qip-progress-bar'),
            progressText: document.getElementById('qip-progress-text'),

            resultSection: document.getElementById('qip-result-section'),
            resultSummary: document.getElementById('qip-result-summary'),
            downloadExcel: document.getElementById('qip-download-excel'),
            sendToSpc: document.getElementById('qip-send-to-spc'),
            errorLog: document.getElementById('qip-error-log'),
            errorList: document.getElementById('qip-error-list'),

            previewPanel: document.getElementById('qip-preview-panel'),
            previewContent: document.getElementById('qip-preview-content'),
            prevSheet: document.getElementById('qip-prev-sheet'),
            nextSheet: document.getElementById('qip-next-sheet')
        };
    },

    bindEvents: function () {
        var self = this;
        if (!this.els.uploadZone) return;

        // File upload
        this.els.uploadZone.addEventListener('click', function () { self.els.fileInput.click(); });
        this.els.fileInput.addEventListener('change', function (e) { if (e.target.files[0]) self.loadFile(e.target.files[0]); });
        this.els.uploadZone.addEventListener('dragover', function (e) { e.preventDefault(); e.currentTarget.classList.add('border-primary'); });
        this.els.uploadZone.addEventListener('dragleave', function (e) { e.preventDefault(); e.currentTarget.classList.remove('border-primary'); });
        this.els.uploadZone.addEventListener('drop', function (e) {
            e.preventDefault();
            e.currentTarget.classList.remove('border-primary');
            if (e.dataTransfer.files[0]) self.loadFile(e.dataTransfer.files[0]);
        });

        if (this.els.removeFile) this.els.removeFile.addEventListener('click', function () { self.removeFile(); });

        // Cavity count change
        if (this.els.cavityCount) this.els.cavityCount.addEventListener('change', function () { self.handleCavityCountChange(); });

        // Worksheet selection
        if (this.els.worksheetSelect) this.els.worksheetSelect.addEventListener('change', function () { self.previewWorksheet(); });
        if (this.els.previewBtn) this.els.previewBtn.addEventListener('click', function () { self.previewWorksheet(); });

        // Config management
        if (this.els.saveConfig) this.els.saveConfig.addEventListener('click', function () { self.saveConfiguration(); });
        if (this.els.loadConfig) this.els.loadConfig.addEventListener('click', function () { self.showConfigDialog(); });

        // Processing
        if (this.els.startProcess) this.els.startProcess.addEventListener('click', function () { self.startProcessing(); });
        if (this.els.downloadExcel) this.els.downloadExcel.addEventListener('click', function () { self.downloadResults(); });
        if (this.els.sendToSpc) this.els.sendToSpc.addEventListener('click', function () { self.sendToSPC(); });

        // Sheet navigation
        if (this.els.prevSheet) this.els.prevSheet.addEventListener('click', function () { self.switchSheet(-1); });
        if (this.els.nextSheet) this.els.nextSheet.addEventListener('click', function () { self.switchSheet(1); });
    },

    loadFile: function (file) {
        var self = this;
        if (!file.name.match(/\.(xlsx|xls|xlsm)$/i)) {
            alert('請選擇 Excel 檔案');
            return;
        }

        var reader = new FileReader();
        reader.onload = function (e) {
            try {
                var data = new Uint8Array(e.target.result);
                self.workbook = XLSX.read(data, { type: 'array' });
                self.fileName = file.name;

                self.els.uploadZone.classList.add('hidden');
                self.els.fileInfo.classList.remove('hidden');
                self.els.fileName.textContent = file.name;
                self.els.workbookInfo.textContent = self.workbook.SheetNames.length + ' 個工作表';

                // Auto-fill product code
                if (!self.els.productCode.value) {
                    self.els.productCode.value = file.name.replace(/\.[^/.]+$/, '');
                }

                self.updateWorksheetSelector();
                self.els.worksheetSelectGroup.classList.remove('hidden');
                self.updateStartButton();
            } catch (error) {
                alert('檔案讀取失敗: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    },

    removeFile: function () {
        this.workbook = null;
        this.fileName = '';
        this.els.fileInput.value = '';
        this.els.uploadZone.classList.remove('hidden');
        this.els.fileInfo.classList.add('hidden');
        this.els.worksheetSelectGroup.classList.add('hidden');
        this.els.cavityGroups.classList.add('hidden');
        this.els.resultSection.classList.add('hidden');
        this.updateStartButton();
    },

    updateWorksheetSelector: function () {
        var select = this.els.worksheetSelect;
        select.innerHTML = '<option value="">-- 請選擇 --</option>';
        if (this.workbook) {
            this.workbook.SheetNames.forEach(function (name) {
                var opt = document.createElement('option');
                opt.value = name;
                opt.textContent = name;
                select.appendChild(opt);
            });
        }
    },

    handleCavityCountChange: function () {
        var count = parseInt(this.els.cavityCount.value) || 0;
        if (count > 0) {
            this.renderCavityGroups(count);
            this.els.cavityGroups.classList.remove('hidden');
        } else {
            this.els.cavityGroups.classList.add('hidden');
        }
        this.updateStartButton();
    },

    renderCavityGroups: function (cavityCount) {
        var groupCount = Math.ceil(cavityCount / 8);
        var html = '';
        for (var i = 1; i <= groupCount; i++) {
            var start = (i - 1) * 8 + 1;
            var end = Math.min(i * 8, cavityCount);
            html += '<div class="p-4 bg-slate-50 dark:bg-slate-800 rounded-lg space-y-3">' +
                '<div class="text-xs font-bold text-slate-600 dark:text-slate-400">第 ' + start + '-' + end + ' 穴</div>' +
                '<div class="grid grid-cols-2 gap-3">' +
                '<div><label class="text-[10px] text-slate-500 block">穴號範圍</label>' +
                '<input type="text" id="qip-cavity-id-' + i + '" class="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-700 dark:border-slate-600" placeholder="如: K3:R3"></div>' +
                '<div><label class="text-[10px] text-slate-500 block">數據範圍</label>' +
                '<input type="text" id="qip-data-range-' + i + '" class="w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-700 dark:border-slate-600" placeholder="如: K4:R4"></div>' +
                '</div>';
            if (i > 1) {
                html += '<div><label class="text-[10px] text-slate-500 block">頁面偏移</label>' +
                    '<input type="number" id="qip-offset-' + i + '" class="w-20 px-2 py-1 text-xs border rounded bg-white dark:bg-slate-700 dark:border-slate-600" value="1" min="1" max="10"></div>';
            }
            html += '</div>';
        }
        this.els.cavityGroups.innerHTML = html;
    },

    previewWorksheet: function () {
        var sheetName = this.els.worksheetSelect.value;
        if (!sheetName || !this.workbook) return;

        var ws = this.workbook.Sheets[sheetName];
        var html = XLSX.utils.sheet_to_html(ws, { header: '', editable: false });
        this.els.previewContent.innerHTML = html;
        this.els.previewPanel.classList.remove('hidden');
    },

    switchSheet: function (offset) {
        var select = this.els.worksheetSelect;
        var newIdx = select.selectedIndex + offset;
        if (newIdx >= 1 && newIdx < select.options.length) {
            select.selectedIndex = newIdx;
            this.previewWorksheet();
        }
    },

    gatherConfiguration: function () {
        var config = {
            productCode: this.els.productCode.value,
            cavityCount: this.els.cavityCount.value,
            cavityGroups: {}
        };
        var groupCount = Math.ceil(parseInt(config.cavityCount) / 8) || 1;
        for (var i = 1; i <= groupCount; i++) {
            config.cavityGroups[i] = {
                cavityIdRange: (document.getElementById('qip-cavity-id-' + i) || {}).value || '',
                dataRange: (document.getElementById('qip-data-range-' + i) || {}).value || '',
                pageOffset: i === 1 ? 0 : parseInt((document.getElementById('qip-offset-' + i) || {}).value || '1') - 1
            };
        }
        return config;
    },

    saveConfiguration: function () {
        var name = this.els.configName.value.trim();
        if (!name) { alert('請輸入配置名稱'); return; }

        var config = this.gatherConfiguration();
        config.name = name;
        config.savedAt = new Date().toISOString();

        var configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
        var idx = configs.findIndex(function (c) { return c.name === name; });
        if (idx >= 0) configs[idx] = config;
        else configs.push(config);

        localStorage.setItem('qip_configs', JSON.stringify(configs));
        alert('配置已保存');
    },

    showConfigDialog: function () {
        var configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
        if (configs.length === 0) { alert('尚未保存任何配置'); return; }

        var list = configs.map(function (c, i) {
            return c.name + ' (' + c.cavityCount + '穴)';
        }).join('\n');
        var choice = prompt('請選擇配置編號 (0-' + (configs.length - 1) + '):\n' + list);
        if (choice !== null) {
            this.loadConfiguration(parseInt(choice));
        }
    },

    loadConfiguration: function (index) {
        var configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
        var config = configs[index];
        if (!config) return;

        this.els.productCode.value = config.productCode || '';
        this.els.cavityCount.value = config.cavityCount || '';
        this.els.configName.value = config.name || '';
        this.handleCavityCountChange();

        for (var i = 1; i <= 6; i++) {
            var g = config.cavityGroups[i];
            if (g) {
                var cavId = document.getElementById('qip-cavity-id-' + i);
                var dataR = document.getElementById('qip-data-range-' + i);
                var offset = document.getElementById('qip-offset-' + i);
                if (cavId) cavId.value = g.cavityIdRange || '';
                if (dataR) dataR.value = g.dataRange || '';
                if (offset && i > 1) offset.value = (g.pageOffset || 0) + 1;
            }
        }
        this.updateStartButton();
    },

    loadSavedConfigs: function () {
        var configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
        console.log('QIP: Found ' + configs.length + ' saved configs');
    },

    updateStartButton: function () {
        var hasFile = this.workbook !== null;
        var hasCavity = parseInt(this.els.cavityCount.value) > 0;
        var cavId = document.getElementById('qip-cavity-id-1');
        var dataR = document.getElementById('qip-data-range-1');
        var hasRanges = cavId && cavId.value && dataR && dataR.value;

        this.els.startProcess.disabled = !(hasFile && hasCavity && hasRanges);
    },

    startProcessing: function () {
        var self = this;
        if (!this.workbook) { alert('請先上傳檔案'); return; }

        var config = this.gatherConfiguration();
        console.log('QIP Processing config:', config);

        this.els.progress.classList.remove('hidden');
        this.els.startProcess.disabled = true;
        this.els.resultSection.classList.add('hidden');

        // Simulate processing (replace with actual QIPProcessor call)
        setTimeout(function () {
            try {
                if (typeof QIPProcessor !== 'undefined') {
                    var processor = new QIPProcessor(config);
                    processor.processWorkbook(self.workbook, function (p) {
                        self.els.progressBar.style.width = p.percent + '%';
                        self.els.progressText.textContent = p.message;
                    }).then(function (results) {
                        self.processingResults = results;
                        self.showResults(results);
                    }).catch(function (err) {
                        alert('處理失敗: ' + err.message);
                        self.els.startProcess.disabled = false;
                    });
                } else {
                    // Fallback: simple data extraction
                    self.processingResults = self.simpleExtract(config);
                    self.showResults(self.processingResults);
                }
            } catch (error) {
                alert('處理失敗: ' + error.message);
                self.els.startProcess.disabled = false;
            }
        }, 100);
    },

    simpleExtract: function (config) {
        // Basic extraction for testing
        var results = {
            productCode: config.productCode,
            cavityCount: config.cavityCount,
            sheets: [],
            data: [],
            errors: []
        };

        var sheetNames = this.workbook.SheetNames;
        sheetNames.forEach(function (name) {
            results.sheets.push(name);
        });

        return results;
    },

    showResults: function (results) {
        this.els.progressBar.style.width = '100%';
        this.els.progressText.textContent = '處理完成！';
        this.els.resultSection.classList.remove('hidden');

        var summary = '<div class="text-sm"><strong>產品品號:</strong> ' + (results.productCode || '-') + '</div>' +
            '<div class="text-sm"><strong>處理工作表:</strong> ' + (results.sheets ? results.sheets.length : 0) + ' 個</div>' +
            '<div class="text-sm"><strong>提取數據:</strong> ' + (results.data ? results.data.length : 0) + ' 筆</div>';
        this.els.resultSummary.innerHTML = summary;

        if (results.errors && results.errors.length > 0) {
            this.els.errorLog.classList.remove('hidden');
            this.els.errorList.innerHTML = results.errors.join('<br>');
        } else {
            this.els.errorLog.classList.add('hidden');
        }

        this.els.startProcess.disabled = false;
    },

    downloadResults: function () {
        if (!this.processingResults) { alert('沒有可下載的結果'); return; }

        try {
            if (typeof ExcelExporter !== 'undefined') {
                var exporter = new ExcelExporter();
                exporter.exportResults(this.processingResults, this.fileName.replace(/\.[^/.]+$/, '') + '_extracted.xlsx');
            } else {
                // Fallback
                var wb = XLSX.utils.book_new();
                var ws = XLSX.utils.json_to_sheet(this.processingResults.data || []);
                XLSX.utils.book_append_sheet(wb, ws, 'Data');
                XLSX.writeFile(wb, this.fileName.replace(/\.[^/.]+$/, '') + '_extracted.xlsx');
            }
        } catch (error) {
            alert('下載失敗: ' + error.message);
        }
    },

    sendToSPC: function () {
        if (!this.processingResults || !this.processingResults.data) {
            alert('沒有可傳送的數據');
            return;
        }

        // Store extracted data for SPC analysis
        window.qipExtractedData = this.processingResults;

        // Switch to import view and trigger file processing
        if (typeof SPCApp !== 'undefined') {
            SPCApp.switchView('import');
            alert('數據已準備就緒，請在 QIP 匯入頁面選擇檢驗項目進行分析。');
        } else {
            alert('無法連接 SPC 分析模組');
        }
    }
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', function () {
    QIPExtractApp.init();
});
