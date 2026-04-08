# Development Log (DEV_LOG.md)

## [2026-04-08] - Data Format Restoration & Golden Alignment

### 1. 執行動作 (Actions Taken)

- **審計 (Audit)**: 根據使用者提供的正確格式截圖，確認 Commit `41291df` (Old Format) 為正確標準。
- **還原 (Restoration)**: 
  - **`js/input.js`**: 重構偵測邏輯，強勢優先匹配「舊格式」(Specs A-C, Batch D)。
  - **`js/qip/excel-exporter.js`**: 精確對齊 Golden Template，修正 A5:B6 的元數據 (Metadata) 嵌入位置，並強化置中對齊與字體樣式。
- **相容性 (Robustness)**: 強化偵測算法，確保即使檔案中包含 A5:B6 的產品資訊，也能精確識別每一行的批號與量測數據。
- **驗證 (Verification)**: 成功驗證歷史紀錄刪除與「系統重置」功能，確保 LocalStorage 清理機制運作正常，無數據殘留。

### 2. 狀態確認 (Status Verification)

- **Golden Alignment**: 匯出格式與使用者截圖完全吻合。
## [2026-04-08] - SPC Engine Expansion & Statistical Risk Mitigation

### 1. 執行動作 (Actions Taken)

- **數據格式還原 (Format Restoration)**:
  - **`js/input.js`**: 重構偵測邏輯，優先匹配「舊格式」(Specs A-C, Batch D)，提升對 Legacy QIP 檔案的魯棒性。
  - **`js/qip/excel-exporter.js`**: 精確對齊 Golden Template，修正元數據嵌入位置並強化排版樣式。
- **SPC 引擎升級 (Engine Upgrade)**:
  - **`js/engine.js`**: 寫入 $n=11 \sim 48$ 的精確統計常數，取代 Rough Fallback，確保高穴數模具分析的精確性。
  - **`docs/specs/SPC_Calculation_Logic.md`**: 同步更新統計維護文檔，補全 Xbar-R 管制界限公式。
- **統計風險控管 (Statistical Risk Refinement)**:
  - **UI 優化**: 在 `index.html` 的指標說明彈窗中新增「統計警示 (Statistical Caution)」區塊。
  - **風險說明**: 針對高穴數 ($n > 10$) 提示「過敏風險」與「系統性變異落差」，引導使用者專注於趨勢分析。
- **專案審計 (Project Audit)**:
  - **清理冗餘**: 正式刪除 `live_app.js`, `live_engine.js`, `live_index.html` 與 `.bak` 檔案，確保專案 MECE (獨立窮盡)。

### 2. 狀態確認 (Status Verification)

- **Golden Alignment**: 匯出格式與標準 QIP 模板完全吻合。
- **統計精確度**: 已驗證 $n=48$ 下的常數調用與界限計算邏輯。
- **UI/UX**: 完成雙語在地化審計，彈窗內容排版正確且具備專業警示效果。
- **環境穩定**: 本地與 `main` 分支完成同步，所有核心邏輯已模組化至 `js/`。

---

## [2026-04-07] - Repository Cleanup & Synchronization

### 1. 執行動作 (Actions Taken)

- **同步**: 本地與 `origin/main` 完成同步 (`git pull`)。
- **清理 (Cleanup)**: 移除專案中過時且重複的檔案，符合 MECE 原則。
  - [同步刪除] `_archive/` 下 1.x 版本舊程式: `app_old_clean.js`, `app_original.js`, `cavity-analysis_old.js`, `group-analysis_old.js`。
- **修正**: 修正 `js/engine.js` 中 Nelson Rule 8 的重複判斷邏輯。
- **修正**: 修正 `js/input.js` 的 Excel 數據解析邏輯，解決數據錯位與第一個批次空值問題。
  - 將 Specs 映射調整為 B2, C2, D2。
  - 將批號來源調整為 Column A。
  - 將數據起點調整為 Row 3 (E3)，並確保 `batchNames` 與 `dataRows` 同步。

### 2. 狀態確認 (Status Verification)

- **架構穩定**: 所有 UI 引用 (index.html) 已確認指向 `js/` 目錄下之模組。
- **邏輯校準**: 數據解析索引已與 QIP 標準格式完全對齊。

---

## [2026-03-14] - Large-scale Design Rollback

### 1. 決策背景 (Decision Rationale)

- **狀態回歸**: 由於激進的「現代扁平化」UI 優化未能達到預期，產生「空洞感」，決定執行深度回滾。
- **目標版本**: 回退至 `e357359` (修復 QIP 空欄位穩定版本)。

### 2. 修復歷程 (Repair History Summary)

- **[修復]**: 解決了早期 JS 執行異常與編碼損毀問題。
- **[改進]**: 本地數據導出邏輯優化。
- **[回滾]**: 移除所有實驗性的 Modern Flat Design 相關 CSS 與 HTML 修改。

---

_Cleaned up, synchronized, and optimized by Antigravity AI._
