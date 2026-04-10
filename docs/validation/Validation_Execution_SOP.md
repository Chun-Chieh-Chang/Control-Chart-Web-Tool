# 軟體確效執行標準作業程序 (SOP)
**文件代號：** SOP-VAL-SPC-001  
**版本：** V2.0  
**生效日期：** 2026-04-10  
**相關標準：** ISO 80002-2:2017, ISO 22514-2  

---

## 1. 目的 (Purpose)
本程序旨在規範「SPC 分析工具」的性能確效（PQ）執行步驟，確保軟體計算精度、管制圖渲染及異常偵測邏輯始終符合預期要求，並留下客觀證據以供內部/外部稽核。

## 2. 適用範圍 (Scope)
適用於以下時機：
- 軟體初次部署時。
- 管制圖算法（engine.js）或 UI 邏輯發生重大更新後。
- 定期年度再驗證（Re-validation）。

## 3. 權責 (Responsibilities)
- **確效執行人**：負責按照本 SOP 執行測試，並收集截圖證據。
- **品質主管/QA**：負責審核測試結果並對確效報告進行最終簽署。

## 4. 準備工作 (Preparation)
1. **清理環境**：開啟瀏覽器，清理快取或進入無痕模式，確保無舊數據干擾。
2. **獲取金標準數據**：準備 [PQ_Gold_Standard_Data.md](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/PQ_Gold_Standard_Data.md) 中的測試數據與對應之 [Validation_Gold_Standard.xlsx](file:///c:/Users/USER/Downloads/QIP-to-ControlChart/docs/validation/Validation_Gold_Standard.xlsx)。
3. **確認版本**：在系統畫面右下角確認當前軟體版本號。

## 5. 執行步驟 (Procedure)

### 5.1 基礎功能驗證 (OQ)
1. **語法切換測試**：切換繁中/英文介面，確認所有標籤與診斷訊息動態更新。
2. **數學原理通達測試**：點擊「指標說明」->「解構數學原理」，確認 KaTeX 公式渲染正常（無 raw 代碼）。

### 5.2 統計精度確效 (PQ)
1. **導入標準數據**：
    - 點擊「導入 QIP 數據」或拖放 `Validation_Gold_Standard.xlsx`。
    - 依照檔案內頁籤 `Subgroup_n5` 進行解析。
2. **數據比對**：
    - 將介面上顯示的 $CL (X\text{-double-bar})$、$UCL$、$LCL$、$Cpk$ 與 $Ppk$ 數值與 SOP 附錄中的「標竿解答」進行核對。
3. **常數碼表核對**：
    - 確認 $n=5$ 時，系統調用的 $A_2$ 是否為 $0.577$，$D_4$ 是否為 $2.114$。
4. **異常模式偵測驗證**：
    - 輸入故意偏離的數據（Dataset 3），確認系統正確標註為 **Nelson Rule 2 (Shift)**。

### 5.3 證據收集與存檔 (Documentation)
1. **導出報告**：確效完成後，點擊「導出技術報告」。
2. **截圖記錄**：對關鍵偏差點或計算結果進行全螢幕截圖（需包含系統時間）。
3. **存檔路徑**：將導出的 HTML 報告更名為 `Validation_Report_YYYYMMDD_V[ver].html` 並存入品質系統。

## 6. 判定標準 (Acceptance Criteria)
- **零偏差原則 (Zero Deviation)**：所有計算數值與標竿解答之絕對誤差必須 $\le 0.0001$。
- **邏輯一致性**：Nelson Rules 警示位置與預期完全一致。
- **符合以上條件，判定為 PASS**。

---

## 7. 附錄：標竿解答對照表 (n=5)
| 參數名稱 | 期待值 (Gold Standard) |
| :--- | :--- |
| **X-Double-Bar** | 1.5240 |
| **R-bar** | 0.0300 |
| **X-UCL** | 1.5413 |
| **X-LCL** | 1.5067 |
| **Cpk** | 1.9125 (基於 $\sigma_{within}$) |
| **Ppk** | 1.8241 (基於 $\sigma_{overall}$) |

---
**文件結束**
