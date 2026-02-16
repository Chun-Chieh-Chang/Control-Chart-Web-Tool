# 專案開發與修訂日誌 (Development & Revision Log)

## [2026-02-16] 智慧選取邏輯優化與 UI 重構 (Smart Selection Logic & UI Refactor)

**專案**: Control Chart Web Tool (QIP Data Extractor)
**類別**: UI/UX, Selection Logic, Interaction Optimization
**背景**: 提升 QIP 數據提取器的操作效率，將繁瑣的分步選取進化為「兩點式智慧自動分配」模式。

### 變更內容

- 實現自動拆分邏輯：當使用者選取多行區域時，系統自動將 Pt 1 所在行定義為 ID，其餘區域定義為數據範圍。
- **垂直合併儲存格支援**: 優化算法以自動偵測 Pt 1 所在的垂直合併區域，確保 ID 範圍完整覆蓋（如 I24:X25），並使數據範圍正確銜接。
- **引擎邏輯還原 (Crucial)**: 徹底撤銷對 `processor.js` 與 `data-extractor.js` 的所有邏輯修改，將「智慧選取」嚴格定位為 UI 填表助手，確保底層資料處理規則與原始版本（v2.4）完全一致。
- **資料顯示順序優化**: 修正了結果清單自動按字母排序的問題，現在會嚴格按照 Excel 原始工作表與數據行的出現順序顯示項目。
- **修補紀錄**: 修復了開發過程中因誤刪變數導致的 `idRow is not defined` 錯誤。
- **UI 介面重構與美化**:
  - **按鈕整合**: 移除輸入框旁的個人選取按鈕，改在卡片標頭配置「智慧選取」漸變藥丸形按鈕。
  - **視覺增強**: 加入 Hover 浮起動畫、漸變色彩與詳細的 Pt 1 / Pt 2 懸停工具提示 (Tooltip)。
  - **焦點連動 (Focus Linking)**: 點擊或聚焦於 ID 或數據輸入框時，自動觸發卡片的智慧選取功能，減少操作路徑。
- **佈局優化**:
  - 調整卡片樣式為 `qip-cavity-card`，引入懸停投影與更好的邊界對齊。
  - 新增 `.qip-smart-select-btn` 專屬動畫 CSS。

### 技術實作

- 修改 `js/qip-app.js`:
  - 重構 `renderCavityGroups` 模板與事件綁定邏輯。
  - 增強 `confirmRangeSelection` 以包含智慧座標拆分演算法。
- 修改 `index.html`: 新增按鈕動畫 CSS 與視覺樣式。

---

## [2026-02-07] 多模穴專業分析模型 (Multi-Cavity Professional Analysis)

**專案**: Control Chart Web Tool
**類別**: SPC Algorithm, High-Density Data Analysis
**背景**: 針對 16/32/48 等高模穴數之複雜製程，傳統批次分析因無法區分幾何位移噪音而失效。

### 變更內容

- **輪替抽樣 (Rotational Sampling)**:
  - 實現 $n=5$ 輪替抽樣計畫。系統自動將扁平化後的模穴數據進行分組。
  - 數據點現在具備完整的模穴標籤追蹤，支援在圖表懸停時查看具體構成穴號。
- **擴展 Shewhart 管制界限 (Extended Limits)**:
  - 引入 `Extended Shewhart Limits` 計算公式。
  - 使用總體變異 (Total Sigma) 取代組內變異，容許 Model C 製程中的模穴間系統性位置位移。
- **UI 整合與專屬標記**:
  - 新增第四種分析模型卡片「多模穴專業分析」。
  - 結果頁面新增動態勳章與專業診斷建議標記。
  - 圖表標籤自動更新為「Extended Shewhart X̄ Chart」。
- **統計引擎功能擴展**:
  - `engine.js` 新增 `calculateExtendedBatchLimits` 方法。
  - 優化 `calculateProcessCapability` 以支援外部傳入 $n$ 值與 R-bar。

### 技術實作

- 修改 `js/engine.js`: 納入擴展界限模型。
- 修改 `js/app.js`: 實作數據輪替扁平化邏輯與 Tooltip 擴充。
- 更新 `index.html`: UI 模型選項。
- 更新文檔: `README.md`, `docs/specs/SPC_Calculation_Logic.md`。

---

## [2026-02-06] UI 分佈平衡與專案目錄結構化重整 (UI Balance & Project REORG)

**專案**: Control Chart Web Tool
**類別**: UI/UX, File Management, MECE Reorg

### 變更內容

- **上傳區域佈局重整**:
  - 將原本過大的單一上傳卡片（LG 2/3 寬度）改為 **三欄式對稱佈局 (3-Column Grid)**。
  - **左欄**: 匯入數據報表（上傳區間比例縮減，資訊密度提升）。
  - **中欄**: 近期分析活動（將原本隱藏在 Dashboard 的歷史紀錄帶入導入頁面，減少跳轉）。
  - **右欄**: 作業規範指引與模式說明（整合原本零散的側邊資訊）。
- **頁面寬度標準化**: 將「QIP 提取」與「系統設定」頁面容器寬度統一調整為 `max-w-[96%]`，解決切換視圖時的跳動感，確保水平視覺對齊。
- **專案目錄 MECE 清理**:
  - **目錄重整**: 根據 Mutually Exclusive, Collectively Exhaustive 原則，將根目錄下的舊有資料夾（Stitch, templates, VBACode）全數移至 `_archive/reference/` 下進行歸檔。
  - **資源歸位**: 將根目錄多餘的圖片 (`數字人_225x300.png`) 移入 `assets/`。
  - **程式碼瘦身**: 檢查並維持 `js/` 目錄的純淨度，將開發中產生的臨時檔案清除或存檔。
- **失效分析回溯 (Failure Analysis)**:
  - 更新 `docs/logs/FAILURE_ANALYSIS.md`，納入「Minitab 風格轉型失敗」與「初期佈局比例失調」的分析紀錄，作為未來新案設計的參考依據。

### 技術實作

- 修改 `index.html` 中的 Step 1 結構。
- 更新 `js/app.js` 中的 `renderRecentFiles` 模板，微調其邊距與字體大小以適應新版三欄佈局。
- 透過指令執行實體檔案搬移與目錄清理。

## [2026-02-06] 群組趨勢圖表優化 (Group Trend Chart Optimization)

**專案**: Control Chart Web Tool
**類別**: UI Fix, Data Visualization

### 變更內容

- **X 軸標籤優化**: 實施 -45 度旋轉、自動隱藏重疊標籤及動態字體縮放，解決長批號重疊問題。
- **Y 軸智能縮放**: 根據數據極值與規格界限 (USL/LSL) 自動計算最優顯示範圍，消除底部空白，提升數據解析度。
- **數據點標記控制**: 針對大數據集 (>50點) 自動隱藏標記，僅保留趨勢線，減少視覺干擾。
- **交互增強**: 開放 ApexCharts 全功能工具列 (Zoom, Pan, Reset)，支持局部縮放分析。
- **統計數據一致性修正**: 修復詳細數據表底部「匯總行 (ΣX, X̄, R)」在無效數據列上卻顯示 0 或計算值的問題。新增有效數據檢核機制，確保全空批次保持空白顯示。

### 技術實作

- 修改 `js/app.js` 中的 `renderCharts` 函數，優化 `gOpt` 配置。
- 定義 `yMin` / `yMax` 計算邏輯，確保規格線始終在可見範圍內。

---

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
  - **批號顯示優化**: 改採 **垂直顯示 (Vertical Text)** 模式 (`writing-mode: vertical-rl`)，並將字體放大至 **14px 粗體** 且增加字距，確保超長批號 (如 `R1-3509_250123D...`) 清晰可讀且具備工業報表的專業感。
- **數據邏輯修正 (Data Logic Fix)**:
  - **支持重複批號顯示 (Support Duplicate Batches)**: 重構數據結構，將批次儲存由「物件」改為「陣列」。
    - **效果**: 解決多個檔案中相同名稱批號（如 `Setup`）會相互覆蓋的問題。現在所有批號將依序獨立顯示，即使名稱相同也會保留個別數據，滿足多檔案分析需求。
  - **同名工作表處理**: 優化多檔案合併邏輯，支援同名檢驗項目在單一分析中並列顯示。
- **QIP 提取邏輯深度優化 (QIP Extraction Core Fix)**:
  - **支持多檔案規格提取 (Multi-Workbook Spec/Info)**: 重構提取引擎，使其能掃描所有上傳的 Excel 檔案。
  - **智慧批次合併 (Smart Batch Merging)**:
    - **內部合併**: 同一檔案中基準名稱相同的工作表自動合併。
    - **外部隔離**: 不同檔案中的同名工作表視為獨立批次。
  - **Excel 輸出格式重塑**: 將規格固定 A-C、批號固定 D、產品元數據固定 B5/B6，移除空行。
- **專案管理與文檔重整 (REORG & MECE)**:
  - **失敗紀錄文檔化**: 建立 `docs/logs/FAILURE_ANALYSIS.md`。
  - **MECE 檔案結構化**: 依照 MECE 原則重整目錄，收納計算邏輯與日誌至子資料夾，清理根目錄。

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
- **矯正措施**:
  1. 引入動態 Y 軸範圍計算邏輯。
  2. 設置 X 軸標籤旋轉 -45 度並開啟 `hideOverlappingLabels`。
  3. 實施標記點閾值控制（>50點隱藏）。
  4. 增強前端表格渲染檢核，若單一批次無任何數值，則不顯示統計匯總值，避免顯示幽靈數值。
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
