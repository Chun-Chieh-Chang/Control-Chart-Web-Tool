# QIP 滑鼠範圍選擇功能恢復指南

## 概述
實現類似VBA `Application.InputBox(Type:=8)` 的Web版本滑鼠範圍選擇功能

## 需要恢復的功能

### 1. HTML 更改

#### 在 index.html 的工作表選擇區塊（~495行）添加：

```html
<!-- Worksheet Selection for Preview -->
<div id="qip-worksheet-select-group" class="hidden">
    <label class="text-xs font-bold text-slate-500 uppercase mb-1 block">工作表</label>
    <div class="flex gap-2">
        <select id="qip-worksheet-select" class="flex-1 px-3 py-2 border rounded-lg text-sm">
            <option value="">-- 請選擇工作表 --</option>
        </select>
        <button id="qip-preview-btn" class="px-4 py-2 bg-emerald-600 text-white rounded-lg text-sm hover:bg-emerald-700">
            <span class="material-icons-outlined text-base">visibility</span>
            預覽
        </button>
    </div>
</div>
```

#### 在配置區塊後添加全寬預覽面板（~620行）：

```html
<!-- Worksheet Preview Panel (Full Width) -->
<div id="qip-preview-panel" class="saas-card p-6 hidden mt-6">
    <div class="flex items-center justify-between mb-4">
        <div class="flex items-center gap-3">
            <h4 class="text-sm font-bold">工作表預覽</h4>
            <span id="qip-selection-mode" class="text-xs px-2 py-1 bg-primary/10 text-primary rounded hidden">
                選擇模式：<strong id="qip-selection-type">無</strong>
            </span>
        </div>
        <div class="flex gap-2">
            <button id="qip-prev-sheet" class="p-1.5 rounded hover:bg-slate-100">
                <span class="material-icons-outlined text-base">chevron_left</span>
            </button>
            <button id="qip-next-sheet" class="p-1.5 rounded hover:bg-slate-100">
                <span class="material-icons-outlined text-base">chevron_right</span>
            </button>
        </div>
    </div>
    <div id="qip-preview-content" 
         class="overflow-auto border rounded-lg" 
         style="max-height: 600px;"></div>
</div>
```

### 2. renderCavityGroups 更改

在 `js/qip-app.js` 的 renderCavityGroups 函數中添加選擇按鈕：

```javascript
renderCavityGroups: function (cavityCount) {
    var self = this;
    var groupCount = Math.ceil(cavityCount / 8);
    var html = '';
    for (var i = 1; i <= groupCount; i++) {
        var start = (i - 1) * 8 + 1;
        var end = Math.min(i * 8, cavityCount);
        html += '<div class="p-4 bg-slate-50 dark:bg-slate-800 rounded-lg space-y-3">' +
            '<div class="text-xs font-bold">第 ' + start + '-' + end + ' 穴</div>' +
            '<div class="space-y-2">' +
            // 穴號範圍
            '<div class="flex gap-2 items-end">' +
            '<div class="flex-1"><label class="text-[10px] block">穴號範圍</label>' +
            '<input type="text" id="qip-cavity-id-' + i + '" class="qip-range-input w-full px-2 py-1 text-xs border rounded" placeholder="如: K3:R3"></div>' +
            '<button class="qip-select-btn px-3 py-1.5 text-xs bg-primary text-white rounded hover:bg-indigo-700 flex items-center gap-1" data-target="qip-cavity-id-' + i + '" data-type="cavity">' +
            '<span>📍</span> 選擇' +
            '</button>' +
            '</div>' +
            // 數據範圍
            '<div class="flex gap-2 items-end">' +
            '<div class="flex-1"><label class="text-[10px] block">數據範圍</label>' +
            '<input type="text" id="qip-data-range-' + i + '" class="qip-range-input w-full px-2 py-1 text-xs border rounded" placeholder="如: K4:R4"></div>' +
            '<button class="qip-select-btn px-3 py-1.5 text-xs bg-primary text-white rounded hover:bg-indigo-700 flex items-center gap-1" data-target="qip-data-range-' + i + '" data-type="data">' +
            '<span>📍</span> 選擇' +
            '</button>' +
            '</div>' +
            '</div>';
        if (i > 1) {
            html += '<div><label class="text-[10px] block">頁面偏移</label>' +
                '<input type="number" id="qip-offset-' + i + '" class="w-20 px-2 py-1 text-xs border rounded" value="1" min="1" max="10"></div>';
        }
        html += '</div>';
    }
    this.els.cavityGroups.innerHTML = html;
    
    // 綁定選擇按鈕
    setTimeout(function() {
        document.querySelectorAll('.qip-select-btn').forEach(function(btn) {
            btn.addEventListener('click', function(e) {
                e.preventDefault();
                self.startRangeSelection(this.dataset.target, this.dataset.type);
            });
        });
        document.querySelectorAll('.qip-range-input').forEach(function(input) {
            input.addEventListener('input', function() {
                self.updateStartButton();
            });
        });
    }, 100);
},
```

### 3. 核心滑鼠選擇函數

需要添加以下函數到 `qip-app.js`：

- `previewWorksheet()` - 渲染預覽表格
- `renderPreviewTable(ws)` - 創建Excel風格表格
- `bindCellEvents()` - 綁定滑鼠拖曳事件
- `startRangeSelection(targetId, type)` - 啟動選擇模式
- `highlightRange(r1, c1, r2, c2)` - 高亮選中儲存格  
- `confirmRangeSelection(start, end)` - 確認並填入範圍
- `switchSheet(offset)` - 切換工作表
- `updateWorksheetSelector()` - 更新工作表列表

## 操作流程

```
1. 上傳Excel檔案 ✅
2. 選擇工作表
3. 點擊「預覽」按鈕 → 顯示Excel表格
4. 點擊「📍 選擇」按鈕 → 進入選擇模式
5. 在預覽表格中拖曳滑鼠 → 選擇範圍
6. 鬆開滑鼠 → 自動填入範圍（如 K3:R3）
```

## 技術實現

- 使用 `XLSX.utils.sheet_to_html` 渲染原始表格
- 改用自定義渲染創建帶data屬性的儲存格
- mousedown/mouseover/mouseup 實現拖曳選擇
- CSS `.qip-selected` class 高亮顯示
- 計算並格式化Excel地址（如 A1, K3:R3）
