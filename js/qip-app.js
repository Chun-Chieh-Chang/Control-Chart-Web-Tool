/**
 * QIP Data Extract - Application Controller
 * Integrated into ControlChart with SPC data flow.
 */

var QIPExtractApp = {
    workbooks: [],
    fileNames: [],
    processingResults: null,
    selectionMode: null,
    selectionTarget: null,
    selectionStart: null,
    selectionEnd: null,
    groupSheetIndices: {}, // Track which sheet index each group uses
    activeSelectionGroup: null, // Track which group is currently being configured
    firstClickCell: null, // Track the first point in two-click selection



    t: function (zh, en) {
        return (window.currentLang === 'zh') ? zh : en;
    },

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
            resetFileBtn: document.getElementById('qip-reset-file-btn'), // New button cache
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
            nextSheet: document.getElementById('qip-next-sheet'),
            currentSheetName: document.getElementById('qip-current-sheet-name'),
            selectionModeIndic: document.getElementById('qip-selection-mode'),
            selectionTypeText: document.getElementById('qip-selection-type')
        };
    },

    bindEvents: function () {
        var self = this;
        if (!this.els.uploadZone) return;

        // File upload
        this.els.uploadZone.addEventListener('click', function () { self.els.fileInput.click(); });
        this.els.fileInput.addEventListener('change', function (e) { if (e.target.files.length > 0) self.loadFiles(e.target.files); });
        this.els.uploadZone.addEventListener('dragover', function (e) { e.preventDefault(); e.currentTarget.classList.add('border-primary'); });
        this.els.uploadZone.addEventListener('dragleave', function (e) { e.preventDefault(); e.currentTarget.classList.remove('border-primary'); });
        this.els.uploadZone.addEventListener('drop', function (e) {
            e.preventDefault();
            e.currentTarget.classList.remove('border-primary');
            if (e.dataTransfer.files.length > 0) self.loadFiles(e.dataTransfer.files);
        });

        if (this.els.removeFile) this.els.removeFile.addEventListener('click', function () { self.removeFile(); });
        if (this.els.resetFileBtn) this.els.resetFileBtn.addEventListener('click', function () { self.removeFile(); });

        // Cavity count change
        if (this.els.cavityCount) this.els.cavityCount.addEventListener('change', function () { self.handleCavityCountChange(); });

        // Worksheet selection
        if (this.els.worksheetSelect) this.els.worksheetSelect.addEventListener('change', function () { self.previewWorksheet(); });
        if (this.els.previewBtn) this.els.previewBtn.addEventListener('click', function () { self.previewWorksheet(); });

        // Sheet navigation
        if (this.els.prevSheet) this.els.prevSheet.addEventListener('click', function () { self.switchSheet(-1); });
        if (this.els.nextSheet) this.els.nextSheet.addEventListener('click', function () { self.switchSheet(1); });

        // Config management
        if (this.els.saveConfig) this.els.saveConfig.addEventListener('click', function () { self.saveConfiguration(); });
        if (this.els.loadConfig) this.els.loadConfig.addEventListener('click', function () { self.showConfigDialog(); });

        // Processing
        if (this.els.startProcess) this.els.startProcess.addEventListener('click', function () { self.startProcessing(); });
        if (this.els.downloadExcel) this.els.downloadExcel.addEventListener('click', function () { self.downloadResults(); });
        if (this.els.sendToSpc) this.els.sendToSpc.addEventListener('click', function () { self.sendToSPC(); });


    },

    loadFiles: function (files) {
        var self = this;
        var validFiles = Array.from(files).filter(function (file) {
            return file.name.match(/\.(xlsx|xls|xlsm)$/i);
        });

        if (validFiles.length === 0) {
            alert(this.t('請選擇有效的 Excel 檔案', 'Please select valid Excel files'));
            return;
        }

        this.workbooks = [];
        this.fileNames = [];

        var promises = validFiles.map(function (file) {
            return new Promise(function (resolve, reject) {
                var reader = new FileReader();
                reader.onload = function (e) {
                    try {
                        var data = new Uint8Array(e.target.result);
                        var workbook = XLSX.read(data, { type: 'array', cellStyles: true });
                        resolve({ workbook: workbook, fileName: file.name });
                    } catch (error) {
                        reject(error);
                    }
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            });
        });

        Promise.all(promises).then(function (results) {
            self.workbooks = results.map(r => r.workbook);
            self.fileNames = results.map(r => r.fileName);

            self.els.uploadZone.classList.add('hidden');
            self.els.fileInfo.classList.remove('hidden');

            // UI 顯示
            if (self.fileNames.length === 1) {
                self.els.fileName.textContent = self.fileNames[0];
            } else {
                self.els.fileName.textContent = self.fileNames.length + ' ' + self.t('個檔案...', 'files...');
                self.els.fileName.title = self.fileNames.join('\n');
            }

            var totalSheets = self.workbooks.reduce((acc, wb) => acc + wb.SheetNames.length, 0);
            self.els.workbookInfo.textContent = totalSheets + ' ' + self.t('個工作表 (來自 ', 'Sheets (from ') + self.fileNames.length + ' ' + self.t('個檔案)', 'files)');

            // 產品料號自動填入 (以第一個檔案為準)
            if (!self.els.productCode.value) {
                self.els.productCode.value = self.fileNames[0].replace(/\.[^/.]+$/, '');
            }

            self.updateWorksheetSelector();
            // self.els.worksheetSelectGroup.classList.remove('hidden'); // Now hidden by default

            // Automatically preview the first sheet
            if (self.els.worksheetSelect.options.length > 1) {
                self.els.worksheetSelect.selectedIndex = 1;
                self.previewWorksheet();
            }

            self.updateStartButton();

        }).catch(function (error) {
            alert(self.t('檔案讀取失敗: ', 'File read failed: ') + error.message);
        });
    },

    loadFile: function (file) { // 保留舊方法供參考或單檔調用
        this.loadFiles([file]);
    },

    removeFile: function () {
        this.workbooks = [];
        this.fileNames = [];
        this.groupSheetIndices = {};
        this.els.fileInput.value = '';

        this.els.uploadZone.classList.remove('hidden');
        this.els.fileInfo.classList.add('hidden');

        // Reset inputs
        this.els.productCode.value = '';
        this.els.cavityCount.value = '';

        this.els.worksheetSelectGroup.classList.add('hidden');
        this.els.previewPanel.classList.add('hidden');
        this.els.cavityGroups.classList.add('hidden');
        this.els.resultSection.classList.add('hidden');
        this.updateStartButton();
    },

    updateWorksheetSelector: function () {
        var select = this.els.worksheetSelect;
        select.innerHTML = '<option value="">' + this.t('-- 請選擇工作表 --', '-- Select Sheet --') + '</option>';
        if (this.workbooks && this.workbooks.length > 0) {
            // 使用第一個工作表進行預覽與配置設定
            this.workbooks[0].SheetNames.forEach(function (name) {
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
        var self = this;
        var groupCount = Math.ceil(cavityCount / 8);
        var html = '';
        for (var i = 1; i <= groupCount; i++) {
            var start = (i - 1) * 8 + 1;
            var end = Math.min(i * 8, cavityCount);
            html += '<div class="p-4 bg-white dark:bg-slate-800/50 rounded-2xl border border-slate-100 dark:border-slate-700/50 space-y-3 shadow-sm">' +
                '<div class="flex items-center justify-between border-b border-slate-50 dark:border-slate-700/50 pb-2">' +
                '<div class="text-sm font-bold text-slate-400 uppercase tracking-widest">' + this.t('模穴 ', 'Cavities ') + start + '-' + end + '</div>' +
                '<div id="qip-group-sheet-info-' + i + '" class="text-[10px] font-mono text-indigo-400 bg-indigo-50/30 dark:bg-indigo-900/10 px-2 py-0.5 rounded-full border border-indigo-100/50 dark:border-indigo-800/20">Sheet: -</div>' +
                '</div>' +


                '<div class="grid grid-cols-1 gap-3">' +
                // 穴號範圍 (Cavity ID)
                '<div class="flex items-center gap-3">' +
                '<div class="w-7 h-7 rounded-lg bg-emerald-50 dark:bg-emerald-900/30 text-emerald-600 flex items-center justify-center flex-shrink-0"><span class="material-icons-outlined text-sm">tag</span></div>' +
                '<div class="flex-1">' +
                '<input type="text" id="qip-cavity-id-' + i + '" class="qip-range-input w-full bg-transparent text-sm font-mono font-bold text-slate-700 dark:text-slate-300 border-b border-slate-100 dark:border-slate-700 focus:border-emerald-500 outline-none pb-1" placeholder="' + this.t('ID 範圍 (例如 K3:R3)', 'ID Range (e.g. K3:R3)') + '">' +
                '</div>' +
                '<button class="qip-select-btn p-1.5 text-violet-500 hover:bg-violet-50 dark:hover:bg-violet-900/20 rounded-lg transition-all" data-target="qip-cavity-id-' + i + '" data-group="' + i + '" data-type="cavity" title="' + this.t('選取範圍', 'Select Range') + '">' +

                '<span class="material-icons-outlined text-base">ads_click</span></button>' +
                '</div>' +

                // 數據範圍 (Data Values)
                '<div class="flex items-center gap-3">' +
                '<div class="w-7 h-7 rounded-lg bg-blue-50 dark:bg-blue-900/30 text-blue-600 flex items-center justify-center flex-shrink-0"><span class="material-icons-outlined text-sm">bar_chart</span></div>' +
                '<div class="flex-1">' +
                '<input type="text" id="qip-data-range-' + i + '" class="qip-range-input w-full bg-transparent text-sm font-mono font-bold text-slate-700 dark:text-slate-300 border-b border-slate-100 dark:border-slate-700 focus:border-blue-500 outline-none pb-1" placeholder="' + this.t('數據範圍 (例如 K4:R4)', 'Data Range (e.g. K4:R4)') + '">' +
                '</div>' +
                '<button class="qip-select-btn p-1.5 text-violet-500 hover:bg-violet-50 dark:hover:bg-violet-900/20 rounded-lg transition-all" data-target="qip-data-range-' + i + '" data-group="' + i + '" data-type="data" title="' + this.t('選取範圍', 'Select Range') + '">' +

                '<span class="material-icons-outlined text-base">ads_click</span></button>' +
                '</div>' +

                '</div>' +
                '</div>';
        }
        this.els.cavityGroups.innerHTML = html;

        // 綁定選擇按鈕事件
        setTimeout(function () {
            document.querySelectorAll('.qip-select-btn').forEach(function (btn) {
                btn.addEventListener('click', function (e) {
                    e.preventDefault();
                    self.startRangeSelection(this.dataset.target, this.dataset.type, this.dataset.group);

                });
            });
            document.querySelectorAll('.qip-range-input').forEach(function (input) {
                input.addEventListener('input', function () {
                    self.updateStartButton();
                });
            });
        }, 100);
    },

    // Removed redundant updateWorksheetSelector definition


    previewWorksheet: function () {
        var sheetName = this.els.worksheetSelect.value;
        if (!sheetName || !this.workbooks || this.workbooks.length === 0) return;

        if (this.els.currentSheetName) {
            this.els.currentSheetName.textContent = sheetName;
        }

        var ws = this.workbooks[0].Sheets[sheetName];
        this.renderPreviewTable(ws);
        this.els.previewPanel.classList.remove('hidden');
    },

    renderPreviewTable: function (ws) {
        var range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
        var maxRows = Math.min(range.e.r + 1, 100);
        var maxCols = Math.min(range.e.c + 1, 50);
        var merges = ws['!merges'] || [];

        // Debug info in UI
        var selectionTypeEl = document.getElementById('qip-selection-type');
        if (selectionTypeEl) {
            var currentText = selectionTypeEl.innerText.split(' (')[0]; // Reset prev info
            if (merges.length > 0) {
                selectionTypeEl.innerText = currentText + ' (' + this.t('已偵測到 ', 'Detected ') + merges.length + ' ' + this.t('個合併區域', 'merges') + ')';
            } else {
                selectionTypeEl.innerText = currentText;
            }
        }
        console.log('Rendering preview. Merged areas count:', merges.length);

        // Pre-calculate skip map for merged cells
        var skipMap = {}; // key: "r,c", val: true if skipped
        var mergeMap = {}; // key: "r,c", val: {rowspan, colspan} for top-left cell

        merges.forEach(function (m) {
            // Mark top-left cell
            var key = m.s.r + ',' + m.s.c;
            mergeMap[key] = {
                rowspan: m.e.r - m.s.r + 1,
                colspan: m.e.c - m.s.c + 1
            };

            // Mark other cells to skip
            for (var r = m.s.r; r <= m.e.r; r++) {
                for (var c = m.s.c; c <= m.e.c; c++) {
                    if (r === m.s.r && c === m.s.c) continue;
                    skipMap[r + ',' + c] = true;
                }
            }
        });

        var html = '<table class="qip-preview-table" style="border-collapse: collapse; font-size: 12px; width: 100%; color: var(--table-text); background: var(--table-bg);">';

        // Header row
        html += '<thead><tr><th style="border: 1px solid var(--table-border); padding: 4px; background: var(--table-header-bg); width: 40px;"></th>';
        for (var c = 0; c < maxCols; c++) {
            html += '<th style="border: 1px solid var(--table-border); padding: 4px; background: var(--table-header-bg); font-weight: bold; min-width: 60px;">' +
                XLSX.utils.encode_col(c) + '</th>';
        }
        html += '</tr></thead><tbody>';

        // Data rows
        for (var r = 0; r < maxRows; r++) {
            html += '<tr><th style="border: 1px solid var(--table-border); padding: 4px; background: var(--table-header-bg); font-weight: bold;">' + (r + 1) + '</th>';
            for (var c = 0; c < maxCols; c++) {
                if (skipMap[r + ',' + c]) continue;

                var cellAddr = XLSX.utils.encode_cell({ r: r, c: c });
                var cell = ws[cellAddr];
                var value = cell ? (cell.w || cell.v || '') : '';

                var mergeAttr = '';
                var m = mergeMap[r + ',' + c];
                if (m) {
                    mergeAttr = ' rowspan="' + m.rowspan + '" colspan="' + m.colspan + '"';
                }

                html += '<td class="qip-cell" data-row="' + r + '" data-col="' + c + '" data-addr="' + cellAddr + '"' + mergeAttr +
                    ' style="border: 1px solid var(--table-border); padding: 4px; cursor: pointer; background: var(--table-bg); user-select: none;">' +
                    value + '</td>';
            }
            html += '</tr>';
        }
        html += '</tbody></table>';

        this.els.previewContent.innerHTML = html;
        this.bindCellEvents();
    },

    bindCellEvents: function () {
        var self = this;
        var cells = this.els.previewContent.querySelectorAll('.qip-cell');

        cells.forEach(function (cell) {
            cell.addEventListener('click', function (e) {
                if (!self.selectionTarget) return;
                e.preventDefault();

                var currentCell = {
                    row: parseInt(this.dataset.row),
                    col: parseInt(this.dataset.col),
                    addr: this.dataset.addr
                };

                if (!self.firstClickCell) {
                    // First click: set the start point
                    self.firstClickCell = currentCell;
                    self.clearSelection();
                    this.classList.add('qip-selected-first');
                    this.classList.add('qip-selected');

                    // Update instruction text if possible
                    if (self.els.selectionTypeText) {
                        var originalText = self.selectionType === 'cavity' ? self.t('穴號範圍', 'Cavity ID Range') : self.t('數據範圍', 'Data Range');
                        self.els.selectionTypeText.textContent = originalText + ' (' + self.t('請點選結束位置', 'Select end point') + ')';
                    }
                } else {
                    // Second click: set the end point and confirm
                    self.highlightRange(self.firstClickCell.row, self.firstClickCell.col, currentCell.row, currentCell.col);

                    // Small delay to let user see the final selection before confirmed/closed
                    setTimeout(function () {
                        self.confirmRangeSelection(self.firstClickCell, currentCell);
                    }, 50);
                }
            });

            cell.addEventListener('mouseover', function () {
                if (!self.selectionTarget || !self.firstClickCell) return;
                var hoverCell = {
                    row: parseInt(this.dataset.row),
                    col: parseInt(this.dataset.col)
                };
                self.highlightRange(self.firstClickCell.row, self.firstClickCell.col, hoverCell.row, hoverCell.col);
            });
        });
    },


    startRangeSelection: function (targetId, type, groupIndex) {
        this.selectionTarget = targetId;
        this.selectionType = type;
        this.activeSelectionGroup = groupIndex;
        this.firstClickCell = null; // Reset selection state
        this.clearSelection();



        if (this.els.selectionModeIndic && this.els.selectionTypeText) {
            this.els.selectionModeIndic.classList.remove('hidden');
            this.els.selectionTypeText.textContent = type === 'cavity' ? this.t('穴號範圍', 'Cavity ID Range') : this.t('數據範圍', 'Data Range');
        }

        if (this.els.previewPanel.classList.contains('hidden')) {
            this.previewWorksheet();
        }

        this.els.previewPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    },

    highlightRange: function (r1, c1, r2, c2) {
        this.clearSelection();
        var minRow = Math.min(r1, r2);
        var maxRow = Math.max(r1, r2);
        var minCol = Math.min(c1, c2);
        var maxCol = Math.max(c1, c2);

        // Auto-expand selection to include full merged cells
        // Using SheetJS workbook if available, or we need to access the cached merge info
        // Since we don't have easy access to ws here, we can infer from DOM or re-fetch
        // Best way: use the current sheet from workbook
        var sheetName = this.els.worksheetSelect.value;
        if (sheetName && this.workbooks && this.workbooks.length > 0) {
            var ws = this.workbooks[0].Sheets[sheetName];
            var merges = ws['!merges'] || [];

            var changed = true;
            while (changed) {
                changed = false;
                for (var i = 0; i < merges.length; i++) {
                    var m = merges[i];
                    // Check if merge block overlaps with current selection
                    // Overlap logic: not (m.e.c < minCol || m.s.c > maxCol || m.e.r < minRow || m.s.r > maxRow)
                    if (!(m.e.c < minCol || m.s.c > maxCol || m.e.r < minRow || m.s.r > maxRow)) {
                        // Check if we need to expand
                        if (m.s.r < minRow) { minRow = m.s.r; changed = true; }
                        if (m.e.r > maxRow) { maxRow = m.e.r; changed = true; }
                        if (m.s.c < minCol) { minCol = m.s.c; changed = true; }
                        if (m.e.c > maxCol) { maxCol = m.e.c; changed = true; }
                    }
                }
            }
        }

        // Store the effective selection range for confirmation
        this.selectionStart = { row: minRow, col: minCol };
        this.selectionEnd = { row: maxRow, col: maxCol };

        var cells = this.els.previewContent.querySelectorAll('.qip-cell');
        cells.forEach(function (cell) {
            var r = parseInt(cell.dataset.row);
            var c = parseInt(cell.dataset.col);
            // Check if cell is within range
            if (r >= minRow && r <= maxRow && c >= minCol && c <= maxCol) {
                cell.classList.add('qip-selected');
            }
        });
    },

    clearSelection: function () {
        var cells = this.els.previewContent.querySelectorAll('.qip-cell');
        cells.forEach(function (cell) {
            cell.classList.remove('qip-selected');
            cell.classList.remove('qip-selected-first');
        });
    },


    confirmRangeSelection: function (start, end) {
        // Use the calculated expanded range from highlightRange if available
        // otherwise fall back to raw start/end

        var sRow, sCol, eRow, eCol;

        if (this.selectionStart && this.selectionEnd) {
            sRow = this.selectionStart.row;
            sCol = this.selectionStart.col;
            eRow = this.selectionEnd.row;
            eCol = this.selectionEnd.col;
        } else {
            sRow = start.row;
            sCol = start.col;
            eRow = end.row;
            eCol = end.col;
        }

        var startAddr = XLSX.utils.encode_cell({ r: sRow, c: sCol });
        var endAddr = XLSX.utils.encode_cell({ r: eRow, c: eCol });

        var rangeStr = startAddr === endAddr ? startAddr : startAddr + ':' + endAddr;

        var input = document.getElementById(this.selectionTarget);
        if (input) {
            input.value = rangeStr;
            input.dispatchEvent(new Event('input', { bubbles: true }));
        }

        // Save the sheet index for this group
        if (this.activeSelectionGroup) {
            var sheetIdx = this.els.worksheetSelect.selectedIndex - 1; // 0-based index of worksheet
            var sheetName = this.els.worksheetSelect.value;
            this.groupSheetIndices[this.activeSelectionGroup] = sheetIdx;

            // Update UI to show which sheet this group is tied to
            var infoEl = document.getElementById('qip-group-sheet-info-' + this.activeSelectionGroup);
            if (infoEl) {
                infoEl.textContent = 'Sheet: ' + sheetName;
            }
        }

        this.selectionTarget = null;
        this.selectionType = null;
        this.selectionStart = null;
        this.selectionEnd = null;
        this.activeSelectionGroup = null;
        this.firstClickCell = null;



        if (this.els.selectionModeIndic) {
            this.els.selectionModeIndic.classList.add('hidden');
        }
        this.clearSelection();
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
                pageOffset: this.groupSheetIndices[i] !== undefined ? this.groupSheetIndices[i] : (i === 1 ? 0 : 0)
            };
        }

        return config;
    },

    saveConfiguration: function () {
        var name = this.els.configName.value.trim();
        if (!name) { alert(this.t('請輸入配置名稱', 'Please enter configuration name')); return; }

        var config = this.gatherConfiguration();
        config.name = name;
        config.savedAt = new Date().toISOString();

        var configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
        var idx = configs.findIndex(function (c) { return c.name === name; });
        if (idx >= 0) configs[idx] = config;
        else configs.push(config);

        localStorage.setItem('qip_configs', JSON.stringify(configs));
        alert(this.t('配置已保存', 'Configuration saved'));
    },

    showConfigDialog: function () {
        var configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
        if (configs.length === 0) {
            alert(this.t('尚未保存任何配置', 'No saved configurations found.'));
            return;
        }

        var list = configs.map(function (c, i) {
            var date = c.savedAt ? new Date(c.savedAt).toLocaleDateString() : 'N/A';
            return i + ': ' + c.name + ' [' + c.cavityCount + ' ' + this.t('穴', 'Cavities') + ', ' + date + ']';
        }).join('\n');

        var choice = prompt(this.t('請輸入欲讀取的配置編號 (0-', 'Enter layout ID to load (0-') + (configs.length - 1) + '):' + '\n\n' + list);
        if (choice !== null && choice !== '' && !isNaN(choice)) {
            var idx = parseInt(choice);
            if (idx >= 0 && idx < configs.length) {
                this.loadConfiguration(idx);
            } else {
                alert(this.t('編號無效', 'Invalid ID'));
            }
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
                if (cavId) cavId.value = g.cavityIdRange || '';
                if (dataR) dataR.value = g.dataRange || '';

                // Handle new sheet index logic
                if (g.pageOffset !== undefined) {
                    this.groupSheetIndices[i] = g.pageOffset;
                    var infoEl = document.getElementById('qip-group-sheet-info-' + i);
                    if (infoEl && this.workbooks && this.workbooks.length > 0 && this.workbooks[0].SheetNames) {
                        var sheetName = this.workbooks[0].SheetNames[g.pageOffset];
                        if (sheetName) infoEl.textContent = 'Sheet: ' + sheetName;
                    }
                }
            }
        }

        this.updateStartButton();
    },

    loadSavedConfigs: function () {
        var configs = JSON.parse(localStorage.getItem('qip_configs') || '[]');
        console.log('QIP: Found ' + configs.length + ' saved configs');
    },

    updateStartButton: function () {
        var hasFile = this.workbooks && this.workbooks.length > 0;
        var hasCavity = parseInt(this.els.cavityCount.value) > 0;
        var cavId = document.getElementById('qip-cavity-id-1');
        var dataR = document.getElementById('qip-data-range-1');
        var hasRanges = cavId && cavId.value && dataR && dataR.value;

        this.els.startProcess.disabled = !(hasFile && hasCavity && hasRanges);
    },

    startProcessing: async function () {
        var self = this;
        if (!this.workbooks || this.workbooks.length === 0) { alert(this.t('請先上傳檔案', 'Please upload files first')); return; }

        var config = this.gatherConfiguration();

        // Basic validation
        if (!config.cavityGroups[1] || !config.cavityGroups[1].cavityIdRange || !config.cavityGroups[1].dataRange) {
            alert(this.t('請至少設定第一組的穴號範圍和數據範圍', 'Please set Cavity ID and Data range for Group 1'));
            return;
        }

        this.els.progress.classList.remove('hidden');
        this.els.startProcess.disabled = true;
        this.els.resultSection.classList.add('hidden');
        this.els.progressBar.style.width = '5%';
        this.els.progressText.textContent = this.t('初始化處理器...', 'Initializing processor...');

        try {
            // 使用新版 QIPProcessor
            const processor = new QIPProcessor(config);

            // 處理多個工作簿
            this.processingResults = await processor.processWorkbooks(this.workbooks, (progress) => {
                self.els.progressBar.style.width = progress.percent + '%';
                var pctEl = document.getElementById('qip-progress-percent');
                if (pctEl) pctEl.textContent = Math.round(progress.percent) + '%';
                self.els.progressText.textContent = progress.message;
            });

            this.showResults(this.processingResults);
        } catch (error) {
            console.error(error);
            alert(this.t('處理失敗: ', 'Processing failed: ') + error.message);
            this.els.progressBar.style.width = '0%';
            this.els.progressText.textContent = this.t('處理失敗', 'Extraction Failed');
        } finally {
            this.els.startProcess.disabled = false;
        }
    },


    showResults: function (results) {
        var self = this;
        // 使用 setTimeout 確保這是在任何懸掛的進度更新之後執行
        setTimeout(function () {
            if (self.els.progressBar) self.els.progressBar.style.width = '100%';
            if (self.els.progressText) self.els.progressText.textContent = 'Extraction Complete';
        }, 100);

        this.els.resultSection.classList.remove('hidden');

        var summary = '<div class="space-y-3">' +
            '<div class="flex justify-between items-center"><span class="text-sm text-slate-500 uppercase font-bold">' + this.t('產品資料', 'Product') + '</span><span class="text-sm font-bold text-white truncate max-w-[150px]">' + (results.productInfo.productName || '-') + '</span></div>' +
            '<div class="flex justify-between items-center"><span class="text-sm text-slate-500 uppercase font-bold">' + this.t('提取項目數', 'Extracted Items') + '</span><span class="text-sm font-bold text-emerald-400">' + results.itemCount + '</span></div>' +
            '<div class="flex justify-between items-center"><span class="text-sm text-slate-500 uppercase font-bold">' + this.t('總批次數', 'Total Batches') + '</span><span class="text-sm font-bold text-indigo-400">' + results.totalBatches + '</span></div>' +
            '</div>';

        this.els.resultSummary.innerHTML = summary;

        if (results.errors && results.errors.length > 0) {
            this.els.errorLog.classList.remove('hidden');
            this.els.errorList.innerHTML = results.errors.join('<br>');
        } else {
            this.els.errorLog.classList.add('hidden');
        }

        // Render detailed extraction results
        this.renderExtractionDetail(results);

        this.els.startProcess.disabled = false;
    },

    renderExtractionDetail: function (results) {
        var detailPanel = document.getElementById('qip-extraction-detail-panel');
        var detailContent = document.getElementById('qip-extraction-detail-content');

        if (!detailPanel || !detailContent) return;

        var html = '';
        var items = results.inspectionItems || {};
        var itemNames = Object.keys(items);

        if (itemNames.length === 0) {
            html = '<div class="text-center py-8 text-slate-400">' + this.t('無提取數據', 'No extracted data') + '</div>';
        } else {
            itemNames.forEach(function (itemName, idx) {
                var item = items[itemName];
                var batchNames = Object.keys(item.batches);
                var totalBatches = batchNames.length;
                var allCavities = Array.from(item.allCavities).sort(function (a, b) {
                    return parseInt(a) - parseInt(b);
                });
                var cavityCount = allCavities.length;

                // Item Card (Collapsible)
                html += '<div class="border border-slate-200 dark:border-slate-700 rounded-xl overflow-hidden">';

                // Header
                html += '<div class="flex items-center justify-between p-4 bg-slate-50 dark:bg-slate-800/50 cursor-pointer hover:bg-slate-100 dark:hover:bg-slate-800" onclick="this.nextElementSibling.classList.toggle(\'hidden\')">';
                html += '<div class="flex items-center gap-3 flex-1 min-w-0">';
                html += '<div class="w-8 h-8 rounded-lg bg-indigo-50 dark:bg-indigo-900/30 flex items-center justify-center flex-shrink-0">';
                html += '<span class="material-icons-outlined text-sm text-indigo-600">' + (idx % 3 === 0 ? 'straighten' : idx % 3 === 1 ? 'speed' : 'widgets') + '</span>';
                html += '</div>';
                html += '<div class="flex-1 min-w-0">';
                html += '<h5 class="text-sm font-bold text-slate-900 dark:text-white truncate">' + itemName + '</h5>';
                html += '<div class="flex items-center gap-3 mt-0.5">';
                html += '<span class="text-xs text-slate-500"><span class="font-mono font-bold text-indigo-600">' + totalBatches + '</span> ' + (totalBatches > 1 ? 'batches' : 'batch') + '</span>';
                html += '<span class="text-xs text-slate-400">•</span>';
                html += '<span class="text-xs text-slate-500"><span class="font-mono font-bold text-emerald-600">' + cavityCount + '</span> ' + (cavityCount > 1 ? 'cavities' : 'cavity') + '</span>';
                html += '</div>';
                html += '</div>';
                html += '</div>';
                html += '<span class="material-icons-outlined text-slate-400 text-sm">expand_more</span>';
                html += '</div>';

                // Content (Collapsible)
                html += '<div class="hidden border-t border-slate-200 dark:border-slate-700">';

                // Cavity Summary
                html += '<div class="p-4 bg-slate-50/50 dark:bg-slate-900/20 border-b border-slate-200 dark:border-slate-700">';
                html += '<div class="text-xs font-bold text-slate-500 uppercase mb-2">' + (window.currentLang === 'zh' ? '穴號範圍' : 'Cavity Range') + '</div>';
                html += '<div class="flex flex-wrap gap-1">';
                allCavities.forEach(function (cav) {
                    html += '<span class="px-1.5 py-0.5 bg-emerald-100 dark:bg-emerald-900/20 text-emerald-700 dark:text-emerald-400 text-xs font-mono rounded">' + cav + '</span>';
                });
                html += '</div>';
                html += '</div>';

                // Sample Batches (Show first 5)
                html += '<div class="p-4 space-y-2">';
                html += '<div class="text-xs font-bold text-slate-500 uppercase mb-2">' + (window.currentLang === 'zh' ? '批次樣本（前5筆）' : 'Sample Batches (First 5)') + '</div>';
                batchNames.slice(0, 5).forEach(function (batchName) {
                    var batchData = item.batches[batchName];
                    var cavIds = Object.keys(batchData).sort(function (a, b) { return parseInt(a) - parseInt(b); });

                    html += '<div class="flex items-start gap-2 p-2 rounded-lg bg-white dark:bg-slate-800 border border-slate-100 dark:border-slate-700">';
                    html += '<div class="text-xs font-mono font-bold text-slate-700 dark:text-slate-300 min-w-[80px]">' + batchName + '</div>';
                    html += '<div class="flex-1 flex flex-wrap gap-1">';
                    cavIds.forEach(function (cavId) {
                        html += '<span class="text-xs px-1 py-0.5 bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-400 rounded font-mono">' + cavId + ':' + batchData[cavId].toFixed(3) + '</span>';
                    });
                    html += '</div>';
                    html += '</div>';
                });
                if (totalBatches > 5) {
                    html += '<div class="text-xs text-slate-400 text-center pt-2">... ' + (window.currentLang === 'zh' ? '及其他 ' : 'and ') + (totalBatches - 5) + (window.currentLang === 'zh' ? ' 筆批次' : ' more batches') + '</div>';
                }
                html += '</div>';

                html += '</div>'; // End collapsible content
                html += '</div>'; // End item card
            });
        }

        detailContent.innerHTML = html;
        detailPanel.classList.remove('hidden');
    },


    downloadResults: function () {
        if (!this.processingResults) { alert(this.t('沒有可下載的結果', 'No results to download')); return; }

        try {
            const exporter = new ExcelExporter();
            exporter.createFromResults(this.processingResults, this.processingResults.productCode);

            const filename = (this.els.productCode.value || 'QIP_Extract').trim();
            exporter.download(filename);

            console.log('Excel 導出成功');
        } catch (error) {
            console.error(error);
            alert(this.t('下載失敗: ', 'Download failed: ') + error.message);
        }
    },

    sendToSPC: function () {
        console.log('QIP: Preparing to send data to SPC analysis...', this.processingResults);

        if (!this.processingResults) {
            alert(this.t('沒有可傳送的數據：請先點擊「開始提取數據」按鈕。', 'No data to send: Please click "Run Extraction" first.'));
            return;
        }

        const items = this.processingResults.inspectionItems || {};
        const itemCount = Object.keys(items).length;

        console.log('QIP: Inspection items found:', itemCount, items);

        if (itemCount === 0) {
            alert(this.t('沒有可傳送的數據：提取結果中不包含任何有效的檢驗項目。', 'No data to send: No valid inspection items found.'));
            return;
        }

        // Store extracted data for SPC analysis (global window object)
        window.qipExtractedData = this.processingResults;
        console.log('QIP: Data stored in window.qipExtractedData');

        // Switch to import view and trigger file processing
        if (typeof SPCApp !== 'undefined') {
            SPCApp.switchView('import');
            alert(this.t('數據已準備就緒，請在 QIP 匯入頁面選擇檢驗項目進行分析。', 'Data ready. Please select an item on the Import page to analyze.'));
        } else {
            console.error('QIP: SPCApp not found!');
            alert(this.t('無法連接 SPC 分析模組', 'Cannot connect to SPC analysis module'));
        }
    }
};

// Initialize when DOM is ready
document.addEventListener('DOMContentLoaded', function () {
    QIPExtractApp.init();
});
