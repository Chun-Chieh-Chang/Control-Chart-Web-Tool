# 數據格式規格定義 (Golden Data Format Specification)

**版本**: v2.1.0 (2026-04-08 修訂)  
**狀態**: 正式發佈 (Restored from Commitment `41291df`)

---

## 1. 核心結構 (Core Column Mapping)

系統優先識別以下欄位排列順序：

| 欄位索引 (0-based) | 欄位代號 | 標題文字 (Header) | 內容說明 |
| :--- | :--- | :--- | :--- |
| 0 | Column A | **Target** | 產品標準值 (規格中心) |
| 1 | Column B | **USL** | 規格上限 |
| 2 | Column C | **LSL** | 規格下限 |
| 3 | Column D | **生產批號** | 批次名稱 (Batch ID) |
| 4+ | Column E+ | **(穴號內容)** | 量測數據 (如：1號穴, 2號穴...) |

---

## 2. 行級邏輯 (Row Logic)

### 2.1 標題行 (Row 1)
- 必須包含 `Target`, `USL`, `LSL`, `生產批號`。
- 穴號欄位標題必須包含字元「**穴**」。

### 2.2 規格行 (Row 2)
- 規格數值建議寫在 **Row 2**。
- 系統僅在數據讀取起始點附近的 A, B, C 欄尋找規格值。

### 2.3 元數據嵌入 (Embedded Metadata)
為了兼容舊版 VBA 匯出格式，Row 5 與 Row 6 被定義為「混合行」：

- **Row 5**:
  - `A5`: `ProductName` (標籤)
  - `B5`: 產品名稱數值
  - `D5`: 第 4 批數據的批號
- **Row 6**:
  - `A6`: `MeasurementUnit` (標籤)
  - `B6`: 單位 (如 Inch, mm)
  - `D6`: 第 5 批數據的批號

> [!IMPORTANT]
> 數據提取引擎 (`DataInput.parse`) 已針對此嵌入邏輯進行優化。在讀取 Row 5/6 時，系統會自動忽略 A, B 欄的標籤，並正確從 D 欄獲取批號，從 E 欄之後獲取量測值，確保數據不會因元數據而移位或遺失。

---

## 3. 自動偵測邏輯 (Robustness Logic)

雖然系統鎖定上述 Golden Format 為優先標準，但讀取引擎 (`js/input.js`) 具備以下相容能力：

1. **優先權偵測**: 優先檢查 A=Target, B=USL, C=LSL。
2. **標竿格式相容**: 若偵測到 A=生產批號 且 C=Target，則自動切換至「標竿格式」(Benchmark Format) 讀取邏輯。
3. **起點自適應**: 自動檢測數據是從 Row 2 (無 Header 中間行) 還是 Row 3 (有標柱) 開始，確保規格讀取的完整性。

---

## 4. 變更紀錄 (Historical Context)

- **2026-04-08**: 正式回退至 `41291df` 並還原 D 欄位批號結構，解決了因標竿格式推行導致的舊版檔案不相容問題。
- **2026-04-08**: 新增 A5/B6 元數據嵌入邏輯，完美復刻舊版 VBA 工具之所有視角。

---
*文件編訂: Antigravity AI Architecture Team*
