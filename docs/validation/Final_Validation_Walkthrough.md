# SPC 分析工具 - 軟體確效結案報告 (Walkthrough)

本專案已完成符合 **ISO 80002-2:2017** 標準的軟體確效體系建置。

## ✅ 已完成項目

### 1. 確效文件體系
- [Validation_Plan.md](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/Validation_Plan.md): 定義確效策略與範圍。
- [URS_SRS.md](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/URS_SRS.md): 確立需求與軟體規格。
- [Risk_Assessment.md](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/Risk_Assessment.md): 完成基於 FMEA 的風險分析。
- [IQ_OQ_PQ_Protocols.md](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/IQ_OQ_PQ_Protocols.md): 提供具體的測試腳本。

### 2. 穩定性優化 (解決 bug)
- **RCA**: 導因於異常 QIP 匯入觸發的 `null` 對象處理。
- **FIX**: 推送 `2abc193` 修復 `toFixed` 與 `AI 診斷` 渲染報錯。

### 3. 黃金數據集 (Gold Standard)
- 生成 [Validation_Gold_Standard_QIP.xlsx](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/Validation_Gold_Standard_QIP.xlsx)。
- 支援「拖入即用」的 QIP 批次匯入模式。

## 🛠️ 下一步：如何執行 PQ (性能確認)

1. 打開軟體，進入 **「匯入 QIP 數據」** 頁面。
2. 拖入 [Validation_Gold_Standard_QIP.xlsx](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/Validation_Gold_Standard_QIP.xlsx)。
3. 在識讀設定中配置：
   - **檢驗項目名稱**: B 列
   - **穴號範圍 (Cavity ID)**: G8:V8
   - **數據讀值 (Data)**: G12:V13
4. 點擊開始處理，將數據發送至 SPC 頁面。
5. 比對生成的 **Cpk/Ppk** 是否與 Excel 中的「EXPECTED RESULTS」相符。

> [!TIP]
> 建議將此過程截圖存檔，作為確效報告的客觀證據 (Objective Evidence)。
