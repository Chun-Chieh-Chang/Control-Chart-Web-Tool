# Control Chart Web Tool

> **SPC 統計製程管制分析工具 / SPC Statistical Process Control Analysis Tool**  
> Web-based QIP (Quality Inspection Program) analysis system

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](../LICENSE)
![Version](https://img.shields.io/badge/version-2.1.0-blue)
[![Deploy to GitHub Pages](https://github.com/Chun-Chieh-Chang/Control-Chart-Web-Tool/actions/workflows/deploy.yml/badge.svg)](https://github.com/Chun-Chieh-Chang/Control-Chart-Web-Tool/actions/workflows/deploy.yml)

> 🚀 **自動部署**: 使用 GitHub Actions 自動部署至 GitHub Pages | [部署文檔](docs/guides/DEPLOYMENT.md)

[中文](#中文說明) | [English](#english-description)

---

## 中文說明

### 📊 功能特色

- **四種專業分析模式**
  - 📈 **批號分析**: X̄-R 管制圖，每頁 25 批號獨立計算管制界限
  - 🔍 **模穴分析**: 模穴比較 + Cpk/Ppk 製程能力評估
    - **獨家功能**: 內建 AI 模穴平衡性分析，自動判定流道平衡狀況。
  - 📊 **多模穴專業分析**: 針對高度複雜製程 (16/32 穴) 的解決方案
    - **輪替抽樣**: 自動實施 $n=5$ 的輪替抽樣計畫 (Rotational Sampling)
    - **擴展管制界限**: 採用擴展 Shewhart 管制界限，容許模穴間系統性位置差異
   - 📊 **群組分析**: Min-Max-Avg 管制圖
  - 🧠 **專家解讀系統 (Nelson Rules 1-8)**: 
    - 內建完整的 **Nelson Rules 1-8** 異常偵測
    - 提供「成型現場」與「品管專家」雙視角實務建議
    - 每條規則包含統計原理說明、診斷檢查表、即時處理項目
    - 支援異常列表懸停即時輔助 (Tooltip)

#### Nelson Rules 1-8 快速參考

| 規則 | 條件 | 嚴重程度 |
|------|------|----------|
| Rule 1 | 單點超出 3σ | 🔴 失控 |
| Rule 2 | 連續 9 點在中心線同側 | 🟠 偏移 |
| Rule 3 | 連續 6 點趨勢 | 🟡 磨損 |
| Rule 4 | 連續 14 點交替 | 🟡 震盪 |
| Rule 5 | 2/3 點 > 2σ | 🟠 中度偏移 |
| Rule 6 | 4/5 點 > 1σ | 🔵 輕度偏移 |
| Rule 7 | 15 點在 ±1σ 內 | 🔵 分層 |
| Rule 8 | 8 點避開中心線 | 🟠 混合分佈 |

- **100% 本地端處理**
  - ✅ 無數據上傳，完全保護隱私
  - ✅ 所有計算在瀏覽器中完成
  - ✅ 支援離線使用

- **專業輸出**
  - 📁 VBA 相容格式 Excel 輸出
  - 📊 互動式網頁圖表顯示
  - 🔴 超出管制界限點紅色標示


### 🚀 使用方法

1. **開啟網頁**
   - 訪問: https://chun-chieh-chang.github.io/Control-Chart-Web-Tool/

2. **選擇數據檔案**
   - 點擊或拖曳 Excel 檔案（.xlsx 或 .xls）

3. **選擇檢驗項目**
   - 從列表中選擇要分析的檢驗項目（工作表）

4. **選擇分析類型**
   - 批號分析、模穴分析或群組分析

5. **查看結果**
   - 網頁即時顯示統計結果與圖表
   - 下載 Excel 檔案

6. **搜尋與自訂**
   - 使用頂部搜尋框快速篩選檢驗項目
   - 點擊右下角按鈕切換深色/淺色模式以符合您的視覺偏好

### 📋 輸入檔案格式

Excel 檔案需符合以下格式：

| Row | A (批號) | B (Target) | C (USL) | D (LSL) | E (穴1) | F (穴2) | ... |
|-----|----------|------------|---------|---------|---------|---------|-----|
| 1   | 生產批號 | Target     | USL     | LSL     | 穴1     | 穴2     | ... |
| 2   | (空白)   | 10.5       | 10.8    | 10.2    | (空白)  | (空白)  | ... |
| 3   | B001     |            |         |         | 10.42   | 10.51   | ... |
| 4   | B002     |            |         |         | 10.45   | 10.48   | ... |

**重點**：
- 第 1 行：標題列
- 第 2 行：規格值 (Target, USL, LSL)
- 第 3 行起：數據
- 模穴欄位標題需包含「穴」字

### 🎯 規格限值計算邏輯

**規格上限 (USL) 與下限 (LSL) 的正確性，與上下公差讀取的正確與否有關**。本系統採用先進的**符號感知計算**邏輯，確保各種公差場景都能正確處理。

#### 符號感知計算 (Sign-Aware Calculation)

系統不再僅僅讀取公差的絕對值，而是結合左側符號欄位來決定偏移方向：

| 符號 | 意義 | 計算方式 |
|------|------|----------|
| **`+`** | 正偏移 | 相對於基準值的**正向**偏移 |
| **`-`** | 負偏移 | 相對於基準值的**負向**偏移 |
| **`±`** | 對稱偏移 | 雙向對稱偏移 |

#### 支持複雜公差場景

透過符號感知邏輯，系統可以正確處理非標準的公差組合：

- ✅ **單向正公差**: `+0.1 / +0.05` (兩者皆高於基準)
- ✅ **單向負公差**: `-0.05 / -0.1` (兩者皆低於基準)
- ✅ **傳統對稱公差**: `+0.1 / -0.1`
- ✅ **不對稱公差**: `+0.15 / -0.05`
- ✅ **± 符號公差**: `±0.1`

#### 自動邊界校準

在計算出兩個偏移邊界後，系統會自動比對並將：
- **較大值** 設為 **USL (上規格限)**
- **較小值** 設為 **LSL (下規格限)**

這確保了數據輸出的邏輯一致性，**避免因輸入順序導致的上限小於下限的情形**。

> 📘 **詳細說明**: 請參閱 [`docs/SPECIFICATION_LIMIT_CALCULATION.md`](docs/SPECIFICATION_LIMIT_CALCULATION.md) 了解完整的計算邏輯與範例

### 🛠️ 技術架構

- **前端**: HTML5, Vanilla JavaScript
- **樣式框架**: Tailwind CSS (CDN)
- **Excel 處理**: SheetJS (讀取與生成)
- **圖表**: ApexCharts (SVG rendering for high resolution)
- **計算引擎**: 自訂 SPC 統計引擎

### 📁 專案結構 (MECE 架構)

```
root/                       # 核心進入點
├── index.html              # 主頁面 (Entry Point)
├── README.md               # 專案說明文件
├── AI數字人.png             # UI 作者頭像
├── docs/                   # 文件與歷程系統
│   ├── guides/             # SOP, 操作指南, 維護指南
│   └── logs/               # 開發日誌, 修正總結, 歷史記錄
├── js/                     # 程式邏輯核心
│   ├── app.js              # UI 控製與流程管理
│   ├── engine.js           # SPC 統計運算核心
│   ├── excel-builder.js    # Excel 報表生成 (手動模式)
│   └── qip/                # QIP 定製化解析模組
├── _archive/               # 封存檔案與測試樣板
└── templates/              # 預設 Excel 範本存儲
```

### 📊 SPC 計算公式

#### X̄-R 管制圖
- UCL(X̄) = X̿ + A₂ × R̄
- CL(X̄) = X̿
- LCL(X̄) = X̿ - A₂ × R̄
- UCL(R) = D₄ × R̄
- CL(R) = R̄
- LCL(R) = D₃ × R̄

#### 製程能力指標
- Cp = (USL - LSL) / (6σ)
- Cpk = min[(USL - μ) / (3σ), (μ - LSL) / (3σ)]
- Pp = (USL - LSL) / (6σ_overall)
- Ppk = min[(USL - μ) / (3σ_overall), (μ - LSL) / (3σ_overall)]

### 📝 開發歷程

基於 VBA 程式碼轉換為 Web 應用，保持完全相同的：
- 輸入檔案格式
- 計算邏輯（每頁 25 批號獨立計算管制界限）
- 輸出結構（VBA 相容格式）

### 📄 授權

MIT License

---

## English Description

### 📊 Features

- **Three Analysis Modes**
  - 📈 **Batch Analysis**: X̄-R Control Charts with 25 batches per page
  - 🔍 **Cavity Analysis**: Cavity Comparison + Cpk/Ppk Assessment
    - **Exclusive**: Built-in AI Cavity Balance Analysis for automatic runner balance diagnostic.
  - 📊 **Multi-Cavity Professional**: Solution for highly complex processes (16/32 cavities)
    - **Rotational Sampling**: Automatic implementation of $n=5$ rotational sampling plans
    - **Extended Limits**: Uses Extended Shewhart Control Limits to account for systematic between-cavity position differences
   - 📊 **Group Analysis**: Min-Max-Avg Control Charts
  - 🧠 **Expert Interpretation System (Nelson Rules 1-8)**: 
    - Built-in complete **Nelson Rules 1-8** anomaly detection
    - Dual-perspective practical advice from "Molding Expert" & "QC Expert"
    - Each rule includes: statistical basis, diagnostic checklist, immediate actions
    - Supports hover tooltips on anomaly list for instant guidance

#### Nelson Rules 1-8 Quick Reference

| Rule | Condition | Severity |
|------|-----------|----------|
| Rule 1 | One point beyond 3σ | 🔴 Out of Control |
| Rule 2 | 9 points on same side of CL | 🟠 Mean Shift |
| Rule 3 | 6 consecutive trending points | 🟡 Drift/Wear |
| Rule 4 | 14 alternating points | 🟡 Oscillation |
| Rule 5 | 2 of 3 points > 2σ | 🟠 Medium Shift |
| Rule 6 | 4 of 5 points > 1σ | 🔵 Small Shift |
| Rule 7 | 15 points within ±1σ | 🔵 Stratification |
| Rule 8 | 8 points avoiding center | 🟠 Mixture/Bimodal |

- **100% Client-Side Processing**
  - ✅ No data upload, complete privacy protection
  - ✅ All calculations performed in browser
  - ✅ Offline support

- **Professional Output**
  - 📁 VBA-compatible Excel file output
  - 📊 Interactive web chart display
  - 🔴 Out-of-control points highlighted in red

- **Bilingual Interface**
  - 🌐 Traditional Chinese and English support
  - 📱 Responsive design for desktop and mobile
  - 🌓 **Dark Mode**: One-click toggle for dark/light themes with auto-adjusting charts/tables
  - 🔍 **Global Search**: Instantly find inspection items and history records via keywords

### 🚀 Usage

1. **Open the Web App**
   - Visit: https://chun-chieh-chang.github.io/Control-Chart-Web-Tool/

2. **Select Data File**
   - Click or drag Excel file (.xlsx or .xls)

3. **Select Inspection Item**
   - Choose the inspection item (worksheet) to analyze

4. **Select Analysis Type**
   - Batch, Cavity, or Group analysis

5. **View Results**
   - Real-time statistics and charts on web page
   - Download Excel file

6. **Search & Customization**
   - Use the top search bar to filter items quickly
   - Toggle Dark/Light mode via the floating button for personalized visual experience

### 📋 Input File Format

Excel file must follow this format:

| Row | A (Batch) | B (Target) | C (USL) | D (LSL) | E (穴1) | F (穴2) | ... |
|-----|-----------|------------|---------|---------|---------|---------|-----|
| 1   | Batch No. | Target     | USL     | LSL     | Cavity1 | Cavity2 | ... |
| 2   | (empty)   | 10.5       | 10.8    | 10.2    | (empty) | (empty) | ... |
| 3   | B001      |            |         |         | 10.42   | 10.51   | ... |
| 4   | B002      |            |         |         | 10.45   | 10.48   | ... |

**Key Points**:
- Row 1: Headers
- Row 2: Specifications (Target, USL, LSL)
- Row 3+: Data
- Cavity column headers must contain "穴"

### 🎯 Specification Limit Calculation Logic

**The correctness of USL (Upper Specification Limit) and LSL (Lower Specification Limit) depends on proper tolerance reading**. This system employs advanced **sign-aware calculation** logic to ensure accurate handling of various tolerance scenarios.

#### Sign-Aware Calculation

The system no longer simply reads absolute tolerance values, but combines them with sign symbols to determine offset direction:

| Symbol | Meaning | Calculation Method |
|--------|---------|-------------------|
| **`+`** | Positive Offset | **Positive** offset from nominal value |
| **`-`** | Negative Offset | **Negative** offset from nominal value |
| **`±`** | Symmetric Offset | Bidirectional symmetric offset |

#### Support for Complex Tolerance Scenarios

Through sign-aware logic, the system can correctly handle non-standard tolerance combinations:

- ✅ **Single-sided positive tolerance**: `+0.1 / +0.05` (both above nominal)
- ✅ **Single-sided negative tolerance**: `-0.05 / -0.1` (both below nominal)
- ✅ **Traditional symmetric tolerance**: `+0.1 / -0.1`
- ✅ **Asymmetric tolerance**: `+0.15 / -0.05`
- ✅ **± symbol tolerance**: `±0.1`

#### Automatic Boundary Calibration

After calculating two offset boundaries, the system automatically compares and assigns:
- **Larger value** as **USL (Upper Specification Limit)**
- **Smaller value** as **LSL (Lower Specification Limit)**

This ensures logical consistency in data output and **prevents scenarios where the upper limit is less than the lower limit due to input order**.

> 📘 **Detailed Documentation**: See [`docs/SPECIFICATION_LIMIT_CALCULATION.md`](docs/SPECIFICATION_LIMIT_CALCULATION.md) for complete calculation logic and examples

### 🛠️ Tech Stack
...
### 📁 Project Structure (MECE)

```
root/                       # Entry point
├── index.html              # Main application page
├── README.md               # Documentation
├── AI數字人.png             # UI Author Image
├── docs/                   # Documentation system
│   ├── guides/             # SOPs, Guides, Maintenance
│   └── logs/               # Dev logs, Summaries, History
├── js/                     # Application logic
│   ├── app.js              # UI controller
│   ├── engine.js           # Statistical engine (SPC)
│   ├── excel-builder.js    # Excel report builder (Manual)
│   └── qip/                # QIP parsing modules
├── _archive/               # Archived legacy files
└── templates/              # Built-in Excel templates
```

#### X̄-R Control Charts
- UCL(X̄) = X̿ + A₂ × R̄
- CL(X̄) = X̿
- LCL(X̄) = X̿ - A₂ × R̄
- UCL(R) = D₄ × R̄
- CL(R) = R̄
- LCL(R) = D₃ × R̄

#### Process Capability Indices
- Cp = (USL - LSL) / (6σ)
- Cpk = min[(USL - μ) / (3σ), (μ - LSL) / (3σ)]
- Pp = (USL - LSL) / (6σ_overall)
- Ppk = min[(USL - μ) / (3σ_overall), (μ - LSL) / (3σ_overall)]

### 📄 License

MIT License

---

© 2026 Wesley Chang @ Mouldex. All rights reserved.
