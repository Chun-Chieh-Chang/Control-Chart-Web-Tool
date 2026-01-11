# QIP 數據提取修復摘要

## 問題描述
您報告的問題：
1. QIP資料提取檢驗項目所顯示的內容不正確
2. 相鄰的兩個項目只提取第2個項目的數據
3. 需要支援 1~12, 1~16, 1~24, 1~32, 1~48 穴的同頁提取

## 根本原因

通過深入對比 VBA 參考代碼（`vba-reference/theCode.bas`），發現了關鍵差異：

### VBA 代碼（Line 1139-1152）使用 `MergeArea`
```vba
Set m = ws.Cells(dataTopRow + rowOffset, 1).MergeArea
tempValue = Trim(CStr(m.Cells(1, 1).value))
```

VBA 的 `MergeArea` 會自動處理合併儲存格，即使目標單元格是合併區域的一部分，也能正確讀取左上角儲存格的值。

### JavaScript 代碼（原版）**沒有處理合併儲存格**
```javascript
const cell = worksheet[cellAddr];
let value = cell ? cell.v : null;
```

SheetJS 中，合併儲存格的值**只存在於左上角單元格**，其他單元格沒有 `.v` 屬性。

## 修復方案

### 1. 新增 `safeGetMergedValue` 輔助函數
```javascript
const safeGetMergedValue = (row, col) => {
    try {
        const cellAddr = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddr];
        let value = cell && cell.v !== undefined ? cell.v : null;
        
        // 如果沒有值，檢查是否是合併儲存格的一部分
        if (!value || value === '') {
            const merge = merges.find(m => 
                m && m.s && m.e &&
                row >= m.s.r && row <= m.e.r && 
                col >= m.s.c && col <= m.e.c
            );
            if (merge) {
                const mergedAddr = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
                const mergedCell = worksheet[mergedAddr];
                value = mergedCell && mergedCell.v !== undefined ? mergedCell.v : null;
            }
        }
        return value;
    } catch (err) {
        console.error(`[QIP] 讀取儲存格 (${row},${col}) 時發生錯誤:`, err);
        return null;
    }
};
```

### 2. 更新檢驗項目名稱提取邏輯
**之前**：直接讀取 A 或 B 欄，如果是合併儲存格會返回 `undefined`
**現在**：使用 `safeGetMergedValue` 處理合併儲存格

```javascript
let tempValue = safeGetMergedValue(dataRow, 0);  // 先嘗試 A 列
if (!tempValue || String(tempValue).trim() === '') {
    tempValue = safeGetMergedValue(dataRow, 1);  // 再嘗試 B 列
}
```

### 3. 更新穴號ID和數據提取邏輯
**穴號ID**：
```javascript
let cavityId = safeGetMergedValue(idRow, col);
```

**數據值**：
```javascript
const cellValue = safeGetMergedValue(dataRow, col);
```

### 4. 完善的錯誤處理
- 所有陣列操作都有邊界檢查（`m && m.s && m.e`）
- 所有物件訪問都有 null/undefined 檢查
- Try-catch 包裹所有可能失敗的操作
- 詳細的 Console 日誌用於調試

## 修改文件

- **js/qip/processor.js**: 
  - 新增 `safeGetMergedValue` 函數
  - 更新 `extractInspectionItemsFromGroup` 方法
  - 增強日誌輸出

## 預期效果

現在當您執行 QIP 提取時，Console (F12) 會顯示類似以下的詳細日誌：

```
[QIP] 提取穴組數據 - Cavity ID範圍: K3:R3, 數據範圍: K4:R4
[QIP][Row4] ✓ 找到檢驗項目: "長度"
[QIP][Row4] ✓ 提取成功: "長度", 穴號: [1, 2, 3, 4, 5, 6, 7, 8], 數量: 8

[QIP] 合併數據 - 項目: 長度, 批次: 200430D
  ├─ 現有穴號: [1, 2, 3, 4, 5, 6, 7, 8]
  ├─ 新增穴號: [9, 10, 11, 12, 13, 14, 15, 16]
  └─ 合併後總穴數: 16 (穴號: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16])
```

### 支援的配置

- ✅ **1~8 穴**（單個 Group）
- ✅ **1~16 穴**（Group 1 + Group 2，同一頁或不同頁）
- ✅ **1~24 穴**（Group 1 + 2 + 3）
- ✅ **1~32 穴**（Group 1 + 2 + 3 + 4）
- ✅ **1~40 穴**（Group 1-5）
- ✅ **1~48 穴**（Group 1-6）

## 下一步測試建議

1. 使用實際的 QIP 報表檔案測試
2. 檢查 Console 日誌，確認所有檢驗項目都被正確提取
3. 驗證穴號範圍是否完整（例如 1-16 穴都在）
4. 確認數據合併邏輯正確運作
5. 如有問題，將 Console 日誌發送給我以便進一步診斷

## 參考資料

- VBA 源代碼：`vba-reference/theCode.bas` Line 1111-1182 (ExtractInspectionItemsFromGroup)
- VBA 源代碼：`vba-reference/theCode.bas` Line 1184-1235 (AggregateInspectionItemsAcrossGroups)
- SheetJS 合併儲存格文檔：worksheet['!merges'] 格式
