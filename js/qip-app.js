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
                // Specifically verify that merges are parsed
                self.workbook = XLSX.read(data, { type: 'array', cellStyles: true });
                console.log('Workbook loaded. Sheets:', self.workbook.SheetNames);
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
        this.els.previewPanel.classList.add('hidden');
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
        var self = this;
        var groupCount = Math.ceil(cavityCount / 8);
        var html = '';
        for (var i = 1; i <= groupCount; i++) {
            var start = (i - 1) * 8 + 1;
            var end = Math.min(i * 8, cavityCount);
            html += '<div class="p-4 bg-slate-50 dark:bg-slate-800 rounded-lg space-y-3">' +
                '<div class="text-xs font-bold text-slate-600 dark:text-slate-400">第 ' + start + '-' + end + ' 穴</div>' +
                '<div class="space-y-2">' +
                // 穴號範圍
                '<div class="flex gap-2 items-end">' +
                '<div class="flex-1"><label class="text-[10px] text-slate-500 block">穴號範圍</label>' +
                '<input type="text" id="qip-cavity-id-' + i + '" class="qip-range-input w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-700 dark:border-slate-600" placeholder="如: K3:R3"></div>' +
                '<button class="qip-select-btn px-3 py-1.5 text-xs bg-primary text-white rounded hover:bg-indigo-700 flex items-center gap-1" data-target="qip-cavity-id-' + i + '" data-type="cavity">' +
                '<span>📍</span> 選擇</button>' +
                '</div>' +
                // 數據範圍
                '<div class="flex gap-2 items-end">' +
                '<div class="flex-1"><label class="text-[10px] text-slate-500 block">數據範圍</label>' +
                '<input type="text" id="qip-data-range-' + i + '" class="qip-range-input w-full px-2 py-1 text-xs border rounded bg-white dark:bg-slate-700 dark:border-slate-600" placeholder="如: K4:R4"></div>' +
                '<button class="qip-select-btn px-3 py-1.5 text-xs bg-primary text-white rounded hover:bg-indigo-700 flex items-center gap-1" data-target="qip-data-range-' + i + '" data-type="data">' +
                '<span>📍</span> 選擇</button>' +
                '</div>' +
                '</div>';
            if (i > 1) {
                html += '<div><label class="text-[10px] text-slate-500 block">頁面偏移</label>' +
                    '<input type="number" id="qip-offset-' + i + '" class="w-20 px-2 py-1 text-xs border rounded bg-white dark:bg-slate-700 dark:border-slate-600" value="1" min="1" max="10"></div>';
            }
            html += '</div>';
        }
        this.els.cavityGroups.innerHTML = html;

        // 綁定選擇按鈕事件
        setTimeout(function () {
            document.querySelectorAll('.qip-select-btn').forEach(function (btn) {
                btn.addEventListener('click', function (e) {
                    e.preventDefault();
                    self.startRangeSelection(this.dataset.target, this.dataset.type);
                });
            });
            document.querySelectorAll('.qip-range-input').forEach(function (input) {
                input.addEventListener('input', function () {
                    self.updateStartButton();
                });
            });
        }, 100);
    },

    updateWorksheetSelector: function () {
        var select = this.els.worksheetSelect;
        select.innerHTML = '<option value="">-- 請選擇工作表 --</option>';
        if (this.workbook) {
            this.workbook.SheetNames.forEach(function (name) {
                var opt = document.createElement('option');
                opt.value = name;
                opt.textContent = name;
                select.appendChild(opt);
            });
        }
    },

    previewWorksheet: function () {
        var sheetName = this.els.worksheetSelect.value;
        if (!sheetName || !this.workbook) return;

        var ws = this.workbook.Sheets[sheetName];
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
                selectionTypeEl.innerText = currentText + ' (已偵測到 ' + merges.length + ' 個合併區域)';
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

        var html = '<table class="qip-preview-table" style="border-collapse: collapse; font-size: 10px; width: 100%;">';

        // Header row
        html += '<thead><tr><th style="border: 1px solid #cbd5e1; padding: 4px; background: #f1f5f9; width: 40px;"></th>';
        for (var c = 0; c < maxCols; c++) {
            html += '<th style="border: 1px solid #cbd5e1; padding: 4px; background: #f1f5f9; font-weight: bold; min-width: 60px;">' +
                XLSX.utils.encode_col(c) + '</th>';
        }
        html += '</tr></thead><tbody>';

        // Data rows
        for (var r = 0; r < maxRows; r++) {
            html += '<tr><th style="border: 1px solid #cbd5e1; padding: 4px; background: #f1f5f9; font-weight: bold;">' + (r + 1) + '</th>';
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
                    ' style="border: 1px solid #cbd5e1; padding: 4px; cursor: pointer; background: white; user-select: none;">' +
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
        var isSelecting = false;
        var startCell = null;

        cells.forEach(function (cell) {
            cell.addEventListener('mousedown', function (e) {
                if (!self.selectionTarget) return;
                e.preventDefault();
                isSelecting = true;
                startCell = { row: parseInt(this.dataset.row), col: parseInt(this.dataset.col), addr: this.dataset.addr };
                self.clearSelection();
                this.classList.add('qip-selected');
            });

            cell.addEventListener('mouseover', function () {
                if (!isSelecting || !self.selectionTarget) return;
                var endRow = parseInt(this.dataset.row);
                var endCol = parseInt(this.dataset.col);
                self.highlightRange(startCell.row, startCell.col, endRow, endCol);
            });

            cell.addEventListener('mouseup', function () {
                if (!isSelecting || !self.selectionTarget) return;
                isSelecting = false;
                var endRow = parseInt(this.dataset.row);
                var endCol = parseInt(this.dataset.col);
                self.confirmRangeSelection(startCell, { row: endRow, col: endCol, addr: this.dataset.addr });
            });
        });

        document.addEventListener('mouseup', function () {
            isSelecting = false;
        });
    },

    startRangeSelection: function (targetId, type) {
        this.selectionTarget = targetId;
        this.selectionType = type;

        if (this.els.selectionModeIndic && this.els.selectionTypeText) {
            this.els.selectionModeIndic.classList.remove('hidden');
            this.els.selectionTypeText.textContent = type === 'cavity' ? '穴號範圍' : '數據範圍';
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
        if (sheetName && this.workbook) {
            var ws = this.workbook.Sheets[sheetName];
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

        this.selectionTarget = null;
        this.selectionType = null;
        this.selectionStart = null;
        this.selectionEnd = null;

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

    startProcessing: async function () {
        var self = this;
        if (!this.workbook) { alert('請先上傳檔案'); return; }

        var config = this.gatherConfiguration();

        // Basic validation
        if (!config.cavityGroups[1] || !config.cavityGroups[1].cavityIdRange || !config.cavityGroups[1].dataRange) {
            alert('請至少設定第一組的穴號範圍和數據範圍');
            return;
        }

        this.els.progress.classList.remove('hidden');
        this.els.startProcess.disabled = true;
        this.els.resultSection.classList.add('hidden');
        this.els.progressBar.style.width = '10%';
        this.els.progressText.textContent = '初始化處理器...';

        try {
            // 使用新版 QIPProcessor
            const processor = new QIPProcessor(config);

            this.processingResults = await processor.processWorkbook(this.workbook, (progress) => {
                self.els.progressBar.style.width = progress.percent + '%';
                self.els.progressText.textContent = progress.message;
            });

            this.showResults(this.processingResults);
        } catch (error) {
            console.error(error);
            alert('處理失敗: ' + error.message);
            this.els.progressBar.style.width = '0%';
            this.els.progressText.textContent = '處理失敗';
        } finally {
            this.els.startProcess.disabled = false;
        }
    },

    showResults: function (results) {
        this.els.progressBar.style.width = '100%';
        this.els.progressText.textContent = '處理完成！';
        this.els.resultSection.classList.remove('hidden');

        var summary = '<div class="text-sm"><strong>產品名稱:</strong> ' + (results.productInfo.productName || '-') + '</div>' +
            '<div class="text-sm"><strong>產品品號:</strong> ' + (results.productCode || '-') + '</div>' +
            '<div class="text-sm"><strong>檢驗項目:</strong> ' + results.itemCount + ' 個</div>' +
            '<div class="text-sm"><strong>有效批次:</strong> ' + results.totalBatches + ' 批</div>';

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
            const exporter = new ExcelExporter();
            exporter.createFromResults(this.processingResults, this.processingResults.productCode);

            const filename = (this.els.productCode.value || 'QIP_Extract').trim();
            exporter.download(filename);

            console.log('Excel 導出成功');
        } catch (error) {
            console.error(error);
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
