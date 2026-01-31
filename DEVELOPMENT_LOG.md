# 專案開發與修訂日誌 (Development & Revision Log)

## [2026-01-31] UI 介面文字優化
**專案**: Control Chart Web Tool (SPC Analysis Tool)
**類別**: UI/UX 改進

### 變更內容
- **導航選單更新**: 將左側欄「QIP 提取資料匯入」修正為「匯入QIP數據」
  - 位置: `index.html` 第 110 行
  - 影響範圍: 側邊欄導航選單顯示文字
  - 目的: 使介面文字更簡潔明確，符合使用者習慣

### 技術細節
- 同步更新 `data-zh` 屬性以保持雙語系統一致性
- 保持英文版本 `data-en="Import Extracted Data"` 不變

---

**日期**: 2026-01-29
**專案**: MouldexControlChartBuilder (SPC Analysis Tool)
**目標**: 強化多模穴射出成型之專家級統計診斷能力與全局分析邏輯。

---

## 1. 核心算法增強 (Core Algorithm Enhancements)

### 1.1 多模穴平衡性分析 (Cavity Balance Analysis)
- **指標**: 新增「不平衡率 (Imbalance Ratio)」計算。
- **邏輯**: `(模穴均值全距 / 總公差) * 100%`。
- **診斷**: 
  - < 10%: Excellent (良好)
  - 10-25%: Fair (輕微失衡)
  - > 25%: Poor (嚴重失衡 -> 觸發修模建議)

### 1.2 變異源診斷 (Variance Source Diagnosis)
- **指標**: 引入「穩定度指數 (Stability Index, St)」= `Ppk / Cpk`。
- **AI 診斷邏輯**:
  - `St < 0.8`: 判定為 **Shot-to-Shot (批次間)** 變異，主因為機台、原料或環境。
  - `St ≥ 0.8`: 判定為 **Within-Shot (組內)** 變異，主因為模穴差異或單次注射波動。

### 1.3 數據健康度檢核 (Distribution Health Check)
- **指標**: 新增 **偏度 (Skewness)** 與 **峰度 (Kurtosis)** 計算。
- **用途**: 自動判定數據是否符合常態假設。若偏度過大 ($|Skewness| > 1$)，系統會發出統計偏誤警告。

### 1.4 群組穩定度分析 (Group Stability Analysis)
- **指標**: 引入「變異一致性得分 (Consistency Score)」= `(全距之標準差 / 平均全距) * 100%`。
- **用途**: 評估不同批次間的品質波動是否穩定一致。

---

## 2. 全局診斷邏輯 (Global Diagnosis Logic)

### 2.1 跨模式交及驗證 (Cross-Model Triangulation)
- **解決矛盾**: 當批次 Cpk 差但模穴 Cpk 好時，優先判定為「模穴不平衡」。
- **決策優先級**: 物理空間維度 (Cavity) > 穩定度維度 (Group) > 最終績效 (Ppk)。

### 2.2 跨尺寸全局診斷 (Cross-Item Global Diagnosis)
- **邏輯**: 掃描同一模具產出的多個測項（如長度、重量）。
- **判定**:
  - 若所有測項同步失衡 -> **全局模具結構問題**。
  - 若僅特定測項失衡 -> **局部特徵失效 (例如單穴澆口阻塞)**。
  - 敏感度分析 -> **長度看模穴，重量看機台**。

---

## 3. UI/UX 與 系統優化 (UI/UX & System Optimization)

### 3.1 一鍵功能與導航
- **全局分析按鈕**: 在項目選擇頁面加入「全局 AI 診斷」報告生成功能。
- **變更模型按鈕**: 在結果頁面加入跳轉回模型選擇區塊的平滑導引功能。

### 3.2 精度與緩存控制
- **精度修復**: 解決 JavaScript 浮點數轉換造成的長小數位數問題（強制 round 至小數 6 位）。
- **版本控制 (Cache Busting)**：在 `index.html` 中為核心 JS 掛載版本號 (`?v=1.0.2`)，確保 UI 與運算引擎版本同步，防止 `TypeError` 錯誤。

### 3.3 文檔更新
- 更新 `README.md` 與 `SPC_Calculation_Logic.md`，將上述所有專家邏輯納入標準說明文件。

---
**紀錄人**: Antigravity AI Assistant
**適用範圍**: 其他同類 SPC 分析工具之算法移植與邏輯參考。
