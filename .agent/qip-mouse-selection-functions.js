// QIP Mouse Selection Functions Addition
// 添加到qip-app.js的補充代碼

// ===== 在 bindEvents 函數中添加（約第80行後） =====
/*
// Worksheet preview
if (this.els.worksheetSelect) this.els.worksheetSelect.addEventListener('change', function () { self.previewWorksheet(); });
if (this.els.previewBtn) this.els.previewBtn.addEventListener('click', function () { self.previewWorksheet(); });

// Sheet navigation
if (this.els.prevSheet) this.els.prevSheet.addEventListener('click', function () { self.switchSheet(-1); });
if (this.els.nextSheet) this.els.nextSheet.addEventListener('click', function () { self.switchSheet(1); });
*/

// ===== 完整的滑鼠選擇功能函數 (添加在文件末尾，deleteExcel之前) =====

// 更新工作表選擇器
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

// 預覽工作表
previewWorksheet: function () {
    var sheetName = this.els.worksheetSelect.value;
    if (!sheetName || !this.workbook) return;

    var ws = this.workbook.Sheets[sheetName];
    this.renderPreviewTable(ws);
    this.els.previewPanel.classList.remove('hidden');
},

// 渲染預覽表格（Excel�格式）
renderPreviewTable: function(ws) {
    var range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    var maxRows = Math.min(range.e.r + 1, 100);
    var maxCols = Math.min(range.e.c + 1, 50);

    var html = '<table class="qip-preview-table" style="border-collapse: collapse; font-size: 10px; width: 100%;">';

    // Header row with column letters
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
            var cellAddr = XLSX.utils.encode_cell({ r: r, c: c });
            var cell = ws[cellAddr];
            var value = cell ? (cell.w || cell.v || '') : '';
            html += '<td class="qip-cell" data-row="' + r + '" data-col="' + c + '" data-addr="' + cellAddr + '" ' +
                'style="border: 1px solid #cbd5e1; padding: 4px; cursor: pointer; background: white; user-select: none;">' +
                value + '</td>';
        }
        html += '</tr>';
    }
    html += '</tbody></table>';

    this.els.previewContent.innerHTML = html;
    this.bindCellEvents();
},

// 綁定儲存格事件（滑鼠拖曳選擇）
bindCellEvents: function() {
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

    // 防止拖出表格範圍
    document.addEventListener('mouseup', function () {
        isSelecting = false;
    });
},

// 啟動範圍選擇模式
startRangeSelection: function(targetId, type) {
    this.selectionTarget = targetId;
    this.selectionType = type;

    // 更新UI提示
    if (this.els.selectionModeIndic && this.els.selectionTypeText) {
        this.els.selectionModeIndic.classList.remove('hidden');
        this.els.selectionTypeText.textContent = type === 'cavity' ? '穴號範圍' : '數據範圍';
    }

    // 確保預覽已打開
    if (this.els.previewPanel.classList.contains('hidden')) {
        this.previewWorksheet();
    }

    // 滾動到預覽區
    this.els.previewPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
},

// 高亮範圍
highlightRange: function(r1, c1, r2, c2) {
    this.clearSelection();
    var minRow = Math.min(r1, r2);
    var maxRow = Math.max(r1, r2);
    var minCol = Math.min(c1, c2);
    var maxCol = Math.max(c1, c2);

    var cells = this.els.previewContent.querySelectorAll('.qip-cell');
    cells.forEach(function (cell) {
        var r = parseInt(cell.dataset.row);
        var c = parseInt(cell.dataset.col);
        if (r >= minRow && r <= maxRow && c >= minCol && c <= maxCol) {
            cell.classList.add('qip-selected');
        }
    });
},

// 清除選擇
clearSelection: function() {
    var cells = this.els.previewContent.querySelectorAll('.qip-cell');
    cells.forEach(function (cell) {
        cell.classList.remove('qip-selected');
    });
},

// 確認範圍選擇
confirmRangeSelection: function(start, end) {
    var startAddr = start.addr;
    var endAddr = end.addr;
    var rangeStr = startAddr === endAddr ? startAddr : startAddr + ':' + endAddr;

    // 填入目標輸入框
    var input = document.getElementById(this.selectionTarget);
    if (input) {
        input.value = rangeStr;
        input.dispatchEvent(new Event('input', { bubbles: true }));
    }

    // 清除選擇模式
    this.selectionTarget = null;
    this.selectionType = null;
    if (this.els.selectionModeIndic) {
        this.els.selectionModeIndic.classList.add('hidden');
    }
},

// 切換工作表
switchSheet: function (offset) {
    var select = this.els.worksheetSelect;
    var newIdx = select.selectedIndex + offset;
    if (newIdx >= 1 && newIdx < select.options.length) {
        select.selectedIndex = newIdx;
        this.previewWorksheet();
    }
},
