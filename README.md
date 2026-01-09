# Control Chart Web Tool

> **SPC 統計製程管制分析工具 / SPC Statistical Process Control Analysis Tool**  
> Web-based QIP (Quality Inspection Program) analysis system

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
![Version](https://img.shields.io/badge/version-1.0.0-green)

## 🌐 Live Demo

**👉 https://chun-chieh-chang.github.io/Control-Chart-Web-Tool/**

## 📊 Features / 功能特色

### 三種分析模式 / Three Analysis Modes
- **批號分析 (Batch Analysis)**: X̄-R 管制圖，每頁 25 批號獨立計算管制界限
- **模穴分析 (Cavity Analysis)**: Cpk/Ppk 計算，色彩標示製程能力
- **群組分析 (Group Analysis)**: Min-Max-Avg 趨勢圖

### 核心特點 / Key Features
- ✅ 100% 客戶端處理，數據不上傳
- ✅ VBA 相容格式 Excel 輸出
- ✅ 中英雙語介面
- ✅ 分頁導航（每頁 25 批號）
- ✅ 超限點紅色標示
- ✅ 響應式設計

## 🚀 Usage / 使用方法

1. 開啟網頁 / Open the web app
2. 上傳 Excel 檔案 / Upload Excel file
3. 選擇檢驗項目 / Select inspection item
4. 選擇分析類型 / Choose analysis type
5. 查看結果並下載 / View results and download

## 📁 Project Structure / 專案結構

```
Control-Chart-Web-Tool/
├── docs/                   # GitHub Pages 部署目錄
│   ├── index.html          # 主頁面
│   ├── css/style.css       # 樣式表
│   └── js/spc-all.js       # 整合 JavaScript
├── web/                    # 開發目錄
├── *.bas                   # 原始 VBA 程式碼
├── LICENSE                 # MIT 授權
└── README.md               # 本文件
```

## 🛠️ Tech Stack / 技術架構

- **前端 / Frontend**: HTML5, CSS3, Vanilla JavaScript
- **Excel 處理 / Excel**: SheetJS (讀寫)
- **圖表 / Charts**: Chart.js
- **計算引擎 / Engine**: 自訂 SPC 統計引擎

## 📊 SPC Formulas / SPC 計算公式

### X̄-R Control Charts
- UCL(X̄) = X̿ + A₂ × R̄
- CL(X̄) = X̿
- LCL(X̄) = X̿ - A₂ × R̄
- UCL(R) = D₄ × R̄
- LCL(R) = D₃ × R̄

### Process Capability Indices
- Cp = (USL - LSL) / (6σ)
- Cpk = min[(USL - μ) / (3σ), (μ - LSL) / (3σ)]

## 📝 Update / 更新方法

```bash
git add -A
git commit -m "Update description"
git push
```

## 📄 License

MIT License

---

© 2026 Control Chart Web Tool. All rights reserved.
