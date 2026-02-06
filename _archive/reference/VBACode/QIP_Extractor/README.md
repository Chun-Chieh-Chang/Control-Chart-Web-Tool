# VBA 參考代碼

## 📌 說明
此目錄包含原始的 VBA 巨集代碼,**僅供參考與備份使用**,不參與 Web 應用程式的運行。

---

## 📁 檔案說明

### 1. `theCode.bas` - 原始完整 VBA 代碼
- **類型**: 單體檔案 (Monolithic)
- **內容**: 包含所有功能的完整 VBA 巨集
- **用途**: 
  - 可直接匯入 Excel VBA 編輯器使用
  - 作為 Web 版本的邏輯參考
  - 歷史版本備份

### 2. 模組化 VBA 檔案
為了更好的代碼組織,將 `theCode.bas` 拆分為以下模組:

| 檔案 | 對應 Web 模組 | 職責 |
|------|---------------|------|
| `DataExtractor.bas` | `data-extractor.js` | 數據提取 (批號、穴號、數值) |
| `DataValidator.bas` | `data-validator.js` | 數據驗證 (格式檢查、完整性) |
| `SpecificationExtractor.bas` | `spec-extractor.js` | 規格提取 (USL/LSL、公差) |
| `ErrorLogger.bas` | `error-logger.js` | 錯誤日誌記錄 |

---

## 🔄 VBA 與 Web 版本對應關係

### 核心邏輯對應
```
VBA (theCode.bas)                    Web (docs/js/)
├── ProcessInspectionReports_SheetInput → processor.js::processWorkbook()
├── ExtractBatchNumbers              → data-extractor.js::extractBatchNumbers()
├── ExtractCavityData                → data-extractor.js::extractCavityData()
├── FindSpecificationByItem          → spec-extractor.js::findSpecificationByItem()
├── ValidateDataSheet                → data-validator.js::isValidDataSheet()
└── LogError                         → error-logger.js::logError()
```

### 關鍵功能實現差異

| 功能 | VBA 實現 | Web 實現 | 備註 |
|------|----------|----------|------|
| Excel 讀取 | Excel Object Model | SheetJS (xlsx.js) | Web 使用純 JS 函式庫 |
| 批號合併 | VBA Array + Loop | JavaScript Array Methods | 邏輯相同,語法不同 |
| 跨頁處理 | Worksheet.Copy | 虛擬工作表索引 | Web 使用索引計算 |
| 規格提取 | Range.Find | 陣列遍歷 | Web 使用迭代查找 |
| Excel 輸出 | Workbook.SaveAs | SheetJS Write | Web 生成 Blob 下載 |

---

## 🛠 如何使用 VBA 版本

### 方法 1: 匯入完整代碼
1. 開啟 Excel
2. 按 `Alt + F11` 開啟 VBA 編輯器
3. 插入 → 模組
4. 複製 `theCode.bas` 的內容貼上
5. 執行 `ProcessInspectionReports_SheetInput`

### 方法 2: 匯入模組化代碼
1. 開啟 Excel VBA 編輯器
2. 檔案 → 匯入檔案
3. 依序匯入:
   - `DataExtractor.bas`
   - `DataValidator.bas`
   - `SpecificationExtractor.bas`
   - `ErrorLogger.bas`
4. 從 `theCode.bas` 複製主執行函數

### 快速操作流程
1. 首次執行 `ProcessInspectionReports_SheetInput` 會建立「參數配置」工作表並加入操作按鈕
2. 點擊「開啟樣本」載入樣本檔以選取各群組的範圍
3. 點擊「選擇檔案」設定要處理的 Excel 檔之「來源檔案路徑」
4. 設定「模穴數」後,可按右側「套用」按鈕自動展開對應群組;若已啟用事件,變更值也會自動更新
5. 為 1–6 組依序設定:
   - 「穴號範圍」(例: 1號穴…8號穴)
   - 「數據範圍」(例: 5.86, 5.85… 等數值區域)
   - 「頁面偏移」(1=同頁, 2=下一頁, 3=下下頁)
6. 如需保存配置,輸入「配置名稱」並點擊「保存配置」; 之後在「載入歷史」選擇配置
   - 可捲動彈窗: 使用滑鼠上下捲動瀏覽; 雙擊某一行或選取後按「確定」
   - 分頁彈窗(未啟用信任中心時): 在輸入框鍵入「>」切到下一頁、「<」回上一頁; 直接輸入顯示的「編號」可選取; 留空或取消退出。到最後一頁再輸入「>」會回到第 1 頁,在第 1 頁輸入「<」會跳到最後一頁
7. 點擊「開始處理」,系統會生成輸出工作簿,為各檢驗項目建立工作表並添加統計列與錯誤日誌

### 主要輸入欄位
- 來源檔案路徑: `rngSourceFile`
- 產品品號(可選): `rngFileKeyword`
- 模穴數: `rngCavityCount`(右側含「套用」按鈕)
- 樣本檔案路徑: `rngSampleFilePath`
- 配置名稱: `rngConfigName`
- 各群組設定: `rngCavityId_1..6`, `rngData_1..6`, `rngOffset_2..6`
- 群組展開/收合: 各組「頁面偏移」列右側的「展開/收合」按鈕

---

## ⚠️ 注意事項

1. **VBA 代碼不會自動更新**: 
   - Web 版本的功能更新不會同步到 VBA 代碼
   - VBA 代碼僅作為參考與備份

2. **環境需求**:
   - VBA 版本需要 Microsoft Excel (Windows/Mac)
   - Web 版本可在任何現代瀏覽器運行
   - 若需自動安裝工作表事件處理器,請在 Excel 信任中心啟用「信任對 VBA 工程物件模型的存取」
   - 若要顯示側邊的大綱加減符號,請至 Excel → 檔案 → 選項 → 進階 → 本工作表的顯示選項,啟用「若已套用大綱則顯示大綱符號」

3. **功能差異**:
   - VBA 版本可能包含一些 Web 版本未實現的進階功能
   - Web 版本專注於核心數據提取功能

---

## 📝 維護建議

- **不建議修改**: 此目錄的代碼應保持穩定,作為歷史參考
- **新功能開發**: 請在 `docs/js/` 目錄的 Web 模組中進行
- **Bug 修復**: 如發現邏輯錯誤,應同時更新 VBA 和 Web 版本

---

*最後更新: 2026-01-08*
