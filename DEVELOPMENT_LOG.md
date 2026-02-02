# 專案開發與修訂日誌 (Development & Revision Log)

## [2026-02-02] UI 風格重塑 - 國際藝術大師動畫風格 (Global Art Master Animation Style)
**專案**: Control Chart Web Tool (SPC Analysis Tool)
**類別**: UI/UX, Aesthetic Overhaul

### 變更內容
- **風格轉型**: 將介面從標準 SaaS 風格轉型為「吉卜力/宮崎駿 (Ghibli/Miyazaki)」啟發的自然系動畫風格。
- **色彩系統**: 
  - 導入高飽和度的自然色系：森林綠 (`#569C5B`)、天空藍 (`#4FA3C4`)、夕陽橘 (`#E88D67`)。
  - 背景改為動態漸層光暈，模擬自然光影流動。
- **字體更新**: 
  - 將主要字體更換為 `Nunito` (圓潤無襯線體)，提升親和力與童趣感，符合動畫風格設定。
- **動畫特效**:
  - 新增 `.animate-float` 懸浮動畫，應用於卡片與重要元素，創造輕盈感。
  - 增強 `.pulse-red` 異常警示動畫，使其更具動態張力。
  - 卡片互動增加彈性縮放 (`scale-105`) 與光澤流動效果。

### 技術實作細節
- **CSS 架構**: 
  - 全面重寫 `css/style.css`，採用 CSS Variables 定義新的色彩語意。
  - 實作「牛奶玻璃 (Milky Glass)」效果：高模糊度 (`blur(16px)`) 搭配半透明白底 (`rgba(255,255,255,0.85)`)。
  - 優化 Scrollbar 樣式，使其更圓潤並融入整體色調。
- **Tailwind 配置**:
  - 更新 `index.html` 中的 `tailwind.config`，擴展自定義顏色與動畫 `keyframes` (float, cloud)。

### 測試紀錄
- **自動化測試**: 嘗試使用 `browser_subagent` 進行畫面渲染驗證，但因系統環境變數問題 (`$HOME not set`) 導致 Playwright 無法啟動。
- **靜態分析**: 人工檢核 HTML 與 CSS 代碼，確認：
  - Google Fonts 連結正確 (`Nunito`).
  - CSS 變數引用無誤.
  - Tailwind Config 語法正確.
- **待辦**: 需在 User 端實際部署後確認視覺效果是否符合預期。

### [2026-02-02] 版面配置優化
- **寬度擴展**: 應使用者需求，將主要工作區容器寬度由 `max-w-7xl` (約 1280px) 放寬為 `max-w-[96%]`，以充分利用寬螢幕空間，優化寬版報表與圖表的閱讀體驗。
  - 影響範圍: Dashboard, Import, Analysis, History, Nelson Rules 等所有主要視圖。

### [2026-02-02] 報表字體標準化
- **Excel 導出優化**: 將導出報表中的數字與統計摘要字體從 `Times New Roman` 統一更換為 `Microsoft JhengHei` (微軟正黑體)，以確保中英文混排時的視覺一致性與美觀。
  - 影響範圍: 數據矩陣 (Data Matrix)、統計摘要 (Summary Stats) 等所有自動生成的單元格。
- **介面細節修正 (UI Polish)**:
  - **表格滿版**: 將詳細數據表格 (Detailed Data Table) 的寬度由固定像素改為 `100%`，以填滿右側空間，解決視窗放大時的留白問題。
  - **批號顯示優化**: 改採 **垂直顯示 (Vertical Text)** 模式 (`writing-mode: vertical-rl`)，配合 120px 的行高，確保超長批號 (如 `R1-3509_250123D...`) 能在有限的欄寬下完整顯示且不影響排版，此風格亦更接近工業報表的閱讀習慣。

---

## [2026-02-01] UI/UX 現代化改版 (UI/UX Modernization)

**專案**: Control Chart Web Tool
**類別**: UI/UX, Style Update

### 變更內容
- **視覺風格同步**: 參考目標網站 (`chun-chieh-chang.github.io`) 的設計語言進行改版。
  - **背景特效**: 引入動態光暈背景 (Gradient Orbs)，替換原本單調的背景色。
  - **玻璃擬態 (Glassmorphism)**: 全面導入 `.glass-nav` 與 `.saas-card` 樣式，應用於導航欄、側邊欄與卡片元件，提升介面通透感。
  - **字體升級**: 引入 `Material Icons Round` 圓角圖標，使介面更加圓潤現代。
  - **側邊欄優化**: 修正側邊欄在淺色背景下的文字對比度問題，調整選單項目的 hover 狀態配色。

### 技術實作細節
- **CSS 架構**: 
  - 新增 `.hover-lift` 動畫類別，增加互動反饋。
  - 重構 `.saas-card` 樣式，移除舊版邊框，改用 `backdrop-filter: blur(24px)` 與多層陰影。
  - 更新 Scrollbar 樣式以符合整體設計。
- **HTML 結構**:
  - 調整 `<body>` 與容器層級，以支援 `fixed` 定位的背景動畫層。
  - 更新 Header 與 Aside 的 class list，移除舊版背景色，改用玻璃擬態背景。

### 遭遇問題與排除
- **瀏覽器工具限制**: 嘗試使用 `browser_subagent` 進行自動化 UI 分析時遭遇環境變數錯誤 (`$HOME environment variable is not set`)，導致無法啟動 Playwright。
  - **解決方案**: 改用 `curl` 與 `Invoke-WebRequest` 抓取目標網頁源碼 (`temp_target.html`) 進行靜態分析，成功提取關鍵 CSS 變數與結構設計。

---

## [2026-01-31] 部署自動化 - GitHub Actions 整合
**專案**: Control Chart Web Tool (SPC Analysis Tool)
**類別**: DevOps / CI/CD

### 變更內容
- **部署方式升級**: 將 GitHub Pages 部署來源從靜態分支改為 GitHub Actions 自動化工作流程
  - 新增配置檔案: `.github/workflows/deploy.yml`
  - 新增文檔: `docs/guides/DEPLOYMENT.md`
  - 觸發條件: 推送到 `main` 分支時自動部署
  - 支援手動觸發 (workflow_dispatch)

### 技術優勢
1. **自動化部署**: 每次推送自動觸發部署，無需手動操作
2. **版本控制**: 清晰的部署歷史記錄與狀態追蹤
3. **並發保護**: 智能管理並發部署，避免衝突
4. **靈活性**: 未來可擴展建置步驟（壓縮、測試等）

### 檔案結構 (MECE 原則)
```
.github/
└── workflows/
    └── deploy.yml        # 部署工作流程
docs/
└── guides/
    └── DEPLOYMENT.md     # 部署文檔
```

### 配置細節
- **權限設定**: 
  - `contents: read` - 讀取儲存庫
  - `pages: write` - 部署到 Pages
  - `id-token: write` - 身份驗證
- **部署環境**: `github-pages`
- **並發策略**: 不取消進行中的部署

---

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
**紀錄人**: Wesley Chang
**適用範圍**: 其他同類 SPC 分析工具之算法移植與邏輯參考。
