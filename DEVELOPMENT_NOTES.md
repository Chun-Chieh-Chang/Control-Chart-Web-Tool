# 開發技術筆記 (Development Notes)

本文件記錄專案開發過程中遇到的關鍵技術挑戰、解決方案及參數設定，供未來維護參考。

## 2026-01-10: 詳細管制表 (Detailed Table) 版面調整技術總結

### 1. 調整目標
在有限的螢幕空間內，完整顯示「生產批號」與「匯總數據」，並確保版面不會因為內容過長而跑版。

### 2. 核心問題點 (The Root Causes)
在調整欄寬的過程中，主要遭遇了三個層次的阻礙，導致 CSS `width` 設定看似無效：

1.  **Metadata 撐開表格 (Metadata Overflow)**：
    *   **現象**：表格上方的資訊欄（如檢驗人員、圖表編號）內容若是過長，即使使用了 `colspan` 跨欄，瀏覽器為了展示完整文字，仍會強行撐開下方的「生產批號」欄位。
    *   **解決**：必須對 **所有** Metadata 儲存格 (`l1`~`v4_val`) 嚴格執行 `overflow: hidden; white-space: nowrap;`，斬斷它對下方欄位的影響。

2.  **HTML Table 的自動佈局特性**：
    *   **現象**：僅在 `<colgroup>` 或 `<col>` 設定 `width` 往往不夠，當內容擁擠時，瀏覽器會啟用「最小內容寬度 (Min-Content)」保護機制，自動忽視寬度設定。
    *   **解決**：必須採用 **「三重鎖定」** 策略 (`width`, `max-width`, `min-width`)。

3.  **總寬度 vs. 單欄寬度**：
    *   **現象**：右側「彙總」區域雖然視覺上是一大欄，但程式邏輯上是 **4 個小欄位** 合併 (`colspan=4`)。
    *   **解決**：設定變數時需注意，`summary` 變數控制的是 **單一小欄** 的寬度。例如設定 `30px`，實際視覺寬度會是 `30px * 4 = 120px`。

### 3. 關鍵解決方案 (The "Nuclear" Solution)

若未來需要再次調整欄寬，請務必遵循以下 **黃金法則**，修改 `js/spc-all.js` 中的 `renderDetailedDataTable` 函式：

#### 法則一：三重鎖定機制 (Triple Lock)
在生成 HTML 字串時，對每一個 `<td>` (包括標題和數據) 直接寫入 Inline Style，同時設定三個屬性：
```css
width: 58px;      /* 期望寬度 */
max-width: 58px;  /* (1) 限制最大不超過 */
min-width: 58px;  /* (2) 強制最小不縮水 (這是強制瀏覽器服從的關鍵) */
```

#### 法則二：零內距 (Zero Padding)
為了在極小的寬度 (如 50-60px) 內顯示內容，必須犧牲內距：
```css
padding: 0;
```

#### 法則三：全域變數控制
目前所有寬度已集中在 `renderDetailedDataTable` 函式開頭的 `colWidths` 物件中管理，不應在其他地方寫死數字：
```javascript
var colWidths = {
    label: 60,
    batch: 58,    // 生產批號 (目前設定)
    summary: 30   // 彙總基礎寬 (目前總寬 = 30 * 4 = 120px)
};
```

### 4. 當前定案參數 (Current Configuration)

*   **生產批號欄寬**：`58px`
*   **彙總區域總寬**：`120px` (基礎 `30px`)
*   **字體大小**：`10px` (表頭與數據一致)
*   **對齊方式**：彙總欄位強制置中 (`text-align: center !important`)

### 5. 其他注意事項
*   CSS `overflow: hidden` 和 `white-space: nowrap` 必須同時存在，否則破版問題會再次出現。
*   彙總欄位的 `text-align: center` 建議加上 `!important`，以防止外部 CSS (如 `style.css`) 的設定覆蓋。
