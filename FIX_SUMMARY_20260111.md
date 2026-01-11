# 專案修復與優化摘要 - 2026/01/11

本次協作主要解決了以下幾個關鍵問題，提升了 QIP 數據提取的準確性與 SPC 分析的使用體驗。

## 1. QIP 檢驗項目提取 (Inspection Item Extraction)

### 問題
- 檢驗項目名稱（如 `(1)`, `(2)`）因含括號或被誤判為純數字而被跳過。
- 合併儲存格導致部分行無法讀取到項目名稱，造成數據遺漏。
- 相鄰的兩個項目只提取了第二個。

### 修復方案
- **支援合併儲存格**：實現了 `safeGetMergedValue` 函數，模仿 VBA 的 `MergeArea` 邏輯，能正確讀取合併區域左上角的值。
- **允许數字名稱**：移除了 `isNumericString` 的限制，現在像 `(1)`、`1` 這樣的名稱都能被正確提取。
- **簡化提取邏輯**：根據實際報表格式，僅從 A 列（或作為備用的 B 列）讀取檢驗項目名稱，不再嘗試錯誤的組合多列。

## 2. QIP 規格與公差提取 (Spec & Tolerance Extraction)

### 問題
- Cpk 計算錯誤，原因是未正確處理公差的正負號（`+` / `-`）。
- 代碼預設的列索引與實際報表不符。

### 修復方案
- **修正列索引**：更新為符合實際報表的索引：
  - Nominal (目標值): **C 列**
  - Sign (公差符號): **D 列**
  - Tolerance (公差值): **E 列**
- **正確的上下公差邏輯**：
  - **上公差 (USL)**：讀取 Target 所在行，計算 `Target + (Sign * Value)`。
  - **下公差 (LSL)**：讀取 Target 的下一行，計算 `Target + (Sign * Value)`。
  - 這通用且準確地處理了對稱、非對稱及同側公差。

## 3. 異常偵測專家意見 (Expert Interpretation)

### 問題
- 當同一個數據點違反多個 Nelson Rules 時，UI 只顯示第一個規則的專家建議，遺漏了其他規則的建議。

### 修復方案
- **完整顯示建議**：修改 `renderAnomalySidebar` 邏輯，遍歷所有觸發的規則。
- **智能組合**：收集所有相關的「成型專家」與「品管專家」建議，並進行去重處理，確保資訊完整且不冗餘。

## 版本資訊
- 最新 Commit: `fix: correct spec extraction column indices and tolerance sign handling`
- 狀態: 已推送到 GitHub `main` 分支

## 下一步建議
- 建議使用者刷新瀏覽器緩存 (Ctrl+F5) 以確保加載最新的 JavaScript 文件。
- 使用實際的 QIP 報表進行端到端測試，驗證提取的 USL/LSL 數值是否與 Excel 顯示一致。

## �̲ת��A�T�{ (Final Verification)
- �Ҧ��״_�N�X�w�� 2026-01-11 17:46 �T�{�ô���C
