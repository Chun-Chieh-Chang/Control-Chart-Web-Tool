# Development Log (DEV_LOG.md)

## [2026-03-14] - Large-scale Design Rollback

### 1. 決策背景 (Decision Rationale)

- **狀態回歸**: 由於激進的「現代扁平化」UI 優化未能達到使用者對於美觀與結構合理性的預期，且產生了「空洞感」問題，最終決定執行深度回滾。
- **目標版本**: 回退至 `e357359` (修復 QIP 空欄位穩定版本)。此版本被視為視覺與功能平衡的最佳基準點。

### 2. 修復歷程 (Repair History Summary)

- **[修復]**: 解決了早期 JS 執行異常與編碼損毀問題。
- **[改進]**: 本地數據導出邏輯優化。
- **[回滾]**: 移除所有實驗性的 Modern Flat Design 相關 CSS 與 HTML 修改。

---

## [2026-04-07] - Repository Cleanup & Synchronization

# Development Log (DEV_LOG.md)

## [2026-03-14] - Large-scale Design Rollback

### 1. 決策背景 (Decision Rationale)

- **狀態回歸**: 由於激進的「現代扁平化」UI 優化未能達到使用者對於美觀與結構合理性的預期，且產生了「空洞感」問題，最終決定執行深度回滾。
- **目標版本**: 回退至 `e357359` (修復 QIP 空欄位穩定版本)。此版本被視為視覺與功能平衡的最佳基準點。

### 2. 修復歷程 (Repair History Summary)

- **[修復]**: 解決了早期 JS 執行異常與編碼損毀問題。
- **[改進]**: 本地數據導出邏輯優化。
- **[回滾]**: 移除所有實驗性的 Modern Flat Design 相關 CSS 與 HTML 修改。

---

## [2026-04-07] - Repository Cleanup & Synchronization

### 1. 執行動作 (Actions Taken)

- **同步**: 本地與 `origin/main` 完成同步 (`git pull`)。
- **清理 (Cleanup)**: 移除專案中過時且重複的檔案，符合 MECE 原則。
  - [刪除] 根目錄冗餘檔案: `live_app.js`, `live_engine.js`, `live_index.html`。
  - [刪除] `_archive/` 下 1.x 版本舊程式: `app_old_clean.js`, `app_original.js`, `cavity-analysis_old.js`, `group-analysis_old.js`。
- **修正**: 修正 `js/engine.js` 中 Nelson Rule 8 的重複判斷邏輯。
- **修正**: 修正 `js/input.js` 的 Excel 數據解析邏輯，解決數據錯位與第一個批次空值問題。
  - 將 Specs 映射調整為 B2, C2, D2。
  - 將批號來源調整為 Column A。
  - 將數據起點調整為 Row 3 (E3)，並確保 `batchNames` 與 `dataRows` 同步。

### 2. 狀態確認 (Status Verification)

- **架構穩定**: 所有 UI 引用 (index.html) 已確認指向 `js/` 目錄下之模組，刪除冗餘檔案不影響運行。
- **邏輯校準**: 數據解析索引已與 QIP 標準格式（A 欄批號、B-D 欄規格、E 欄起穴、Row 3 起數據）完全對齊。

---

*Cleaned up, synchronized, and optimized by Antigravity AI.*
