# Control Chart Web Tool

> **SPC 統計製程管制分析工具 / SPC Statistical Process Control Analysis Tool**  
> Web-based QIP (Quality Inspection Program) analysis system

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](../LICENSE)
![Version](https://img.shields.io/badge/version-1.0.0-green)

[中文](#中文說明) | [English](#english-description)

---

## 中文說明

### 📊 功能特色

- **三種分析模式**
  - 📈 **批號分析**: X̄-R 管制圖，每頁 25 批號獨立計算管制界限
  - 🔍 **模穴分析**: 模穴比較 + Cpk/Ppk 製程能力評估
  - 📊 **群組分析**: Min-Max-Avg 管制圖

- **100% 本地端處理**
  - ✅ 無數據上傳，完全保護隱私
  - ✅ 所有計算在瀏覽器中完成
  - ✅ 支援離線使用

- **專業輸出**
  - 📁 VBA 相容格式 Excel 輸出
  - 📊 互動式網頁圖表顯示
  - 🔴 超出管制界限點紅色標示

- **中英雙語介面**
  - 🌐 支援繁體中文與英文切換
  - 📱 響應式設計，支援桌面與行動裝置

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

### 🛠️ 技術架構

- **前端**: HTML5, CSS3, Vanilla JavaScript
- **Excel 處理**: SheetJS (讀取與生成)
- **圖表**: Chart.js
- **計算引擎**: 自訂 SPC 統計引擎

### 📁 專案結構

```
web/
├── index.html              # 主頁面
├── css/
│   └── style.css           # 樣式表
└── js/
    └── spc-all.js          # 整合 JavaScript（包含所有功能）
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
  - 📊 **Group Analysis**: Min-Max-Avg Control Charts

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

### 🛠️ Tech Stack

- **Frontend**: HTML5, CSS3, Vanilla JavaScript
- **Excel Processing**: SheetJS (reading and writing)
- **Charts**: Chart.js
- **Calculation Engine**: Custom SPC statistical engine

### 📊 SPC Formulas

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

© 2026 Control Chart Web Tool. All rights reserved.
