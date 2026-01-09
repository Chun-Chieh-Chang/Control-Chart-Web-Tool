# Mouldex Control Chart

> **SPC Statistical Process Control Analysis Tool**  
> Web-based QIP (Quality Inspection Program) analysis system

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Version](https://img.shields.io/badge/version-1.0.0-green)

[中文](#中文說明) | [English](#english-description)

---

## 中文說明

### 📊 功能特色

- **三種分析模式**
  - 📈 **批號分析**: X-Bar R 管制圖 + 製程能力分析
  - 🔍 **模穴分析**: 模穴比較 + 能力評估 (Cp/Cpk/Pp/Ppk)
  - 📊 **群組分析**: Min-Max-Avg 管制圖

- **100% 本地端處理**
  - ✅ 無數據上傳，完全保護隱私
  - ✅ 所有計算在瀏覽器中完成
  - ✅ 支援離線使用

- **專業輸出**
  - 📁 Excel 檔案輸出（含數據表格與圖表圖片）
  - 📊 互動式網頁圖表顯示
  - 🖼️ 圖表匯出為 PNG 圖片

- **中英雙語介面**
  - 🌐 支援繁體中文與英文切換
  - 📱 響應式設計，支援桌面與行動裝置

### 🚀 使用方法

1. **開啟網頁**
   - 在瀏覽器中開啟 `index.html`
   - 或訪問 GitHub Pages: [https://YOUR_USERNAME.github.io/Mouldex-Control-Chart](https://YOUR_USERNAME.github.io/Mouldex-Control-Chart)

2. **選擇數據檔案**
   - 點擊或拖曳 Excel 檔案（.xlsx 或 .xls）

3. **選擇檢驗項目**
   - 從列表中選擇要分析的檢驗項目（工作表）

4. **選擇分析類型**
   - 批號分析、模穴分析或群組分析

5. **查看結果**
   - 網頁即時顯示統計結果與圖表
   - 下載 Excel 檔案或匯出圖表

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

- **前端**: HTML5, CSS3, JavaScript (ES6 Modules)
- **Excel 處理**: 
  - SheetJS (讀取)
  - ExcelJS (生成)
- **圖表**: Chart.js
- **計算引擎**: 自訂 SPC 統計引擎

### 📁 專案結構

```
web/
├── index.html              # 主頁面
├── css/
│   └── style.css           # 樣式表
├── js/
│   ├── app.js              # 應用主控制器
│   ├── data-input.js       # Excel 檔案解析
│   ├── spc-engine.js       # SPC 計算引擎
│   ├── batch-analysis.js   # 批號分析
│   ├── cavity-analysis.js  # 模穴分析
│   ├── group-analysis.js   # 群組分析
│   └── excel-export.js     # Excel 匯出
└── README.md               # 說明文件
```

### 📊 SPC 計算公式

#### X-Bar R 管制圖
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
- 計算邏輯
- 輸出結構

### 📄 授權

MIT License

---

## English Description

### 📊 Features

- **Three Analysis Modes**
  - 📈 **Batch Analysis**: X-Bar R Control Charts + Process Capability
  - 🔍 **Cavity Analysis**: Cavity Comparison + Capability Assessment (Cp/Cpk/Pp/Ppk)
  - 📊 **Group Analysis**: Min-Max-Avg Control Charts

- **100% Client-Side Processing**
  - ✅ No data upload, complete privacy protection
  - ✅ All calculations performed in browser
  - ✅ Offline support

- **Professional Output**
  - 📁 Excel file output (with data tables and chart images)
  - 📊 Interactive web chart display
  - 🖼️ Chart export as PNG images

- **Bilingual Interface**
  - 🌐 Traditional Chinese and English support
  - 📱 Responsive design for desktop and mobile

### 🚀 Usage

1. **Open the Web App**
   - Open `index.html` in browser
   - Or visit GitHub Pages: [https://YOUR_USERNAME.github.io/Mouldex-Control-Chart](https://YOUR_USERNAME.github.io/Mouldex-Control-Chart)

2. **Select Data File**
   - Click or drag Excel file (.xlsx or .xls)

3. **Select Inspection Item**
   - Choose the inspection item (worksheet) to analyze

4. **Select Analysis Type**
   - Batch, Cavity, or Group analysis

5. **View Results**
   - Real-time statistics and charts on web page
   - Download Excel file or export charts

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

- **Frontend**: HTML5, CSS3, JavaScript (ES6 Modules)
- **Excel Processing**: 
  - SheetJS (reading)
  - ExcelJS (generation)
- **Charts**: Chart.js
- **Calculation Engine**: Custom SPC statistical engine

### 📊 SPC Formulas

#### X-Bar R Control Charts
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

© 2026 Mouldex Control Chart. All rights reserved.
