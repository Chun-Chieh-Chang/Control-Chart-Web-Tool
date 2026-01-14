# 規格計算邏輯更新摘要
## Specification Calculation Logic Update Summary

**日期**: 2026-01-14

---

## 🔄 變更對比 (Change Comparison)

### ❌ 舊版邏輯 (Old Logic)

**VBA 實現**:
```vba
' 驗證公差符號並設定公差值
If upperSign = "+" And lowerSign = "-" Then
    spec.UpperTolerance = upperTol
    spec.LowerTolerance = lowerTol
ElseIf upperSign = "±" Then
    spec.UpperTolerance = upperTol
    spec.LowerTolerance = upperTol
Else
    ' 默認處理：假設為對稱公差
    spec.UpperTolerance = upperTol
    spec.LowerTolerance = lowerTol
End If

' 計算USL和LSL
spec.usl = spec.NominalValue + spec.UpperTolerance
spec.lsl = spec.NominalValue - spec.LowerTolerance
```

**問題**:
1. ❌ 只能處理 `+/-` 和 `±` 兩種組合
2. ❌ 無法處理單向公差 (`+/+` 或 `-/-`)
3. ❌ 固定計算方式：`USL = Nominal + Upper`, `LSL = Nominal - Lower`
4. ❌ 可能導致 USL < LSL 的邏輯錯誤

---

### ✅ 新版邏輯 (New Logic)

**核心改進**:
```vba
' 符號感知計算
If upperSign = "±" Then
    boundary1 = spec.NominalValue + upperTol
    boundary2 = spec.NominalValue - upperTol
Else
    ' 根據符號決定偏移方向
    If upperSign = "+" Then
        boundary1 = spec.NominalValue + upperTol
    ElseIf upperSign = "-" Then
        boundary1 = spec.NominalValue - upperTol
    End If
    
    If lowerSign = "+" Then
        boundary2 = spec.NominalValue + lowerTol
    ElseIf lowerSign = "-" Then
        boundary2 = spec.NominalValue - lowerTol
    End If
End If

' 自動邊界校準
If boundary1 > boundary2 Then
    spec.usl = boundary1
    spec.lsl = boundary2
Else
    spec.usl = boundary2
    spec.lsl = boundary1
End If
```

**優勢**:
1. ✅ 支持所有符號組合 (`+/+`, `-/-`, `+/-`, `±`)
2. ✅ 符號感知：根據符號決定加減方向
3. ✅ 自動邊界校準：確保 USL ≥ LSL
4. ✅ 處理複雜場景：單向公差、不對稱公差

---

## 📊 實際案例對比 (Real-World Comparison)

### 案例 1: 單向正公差 (+0.10 / +0.05)

| 項目 | 舊版邏輯 | 新版邏輯 |
|------|----------|----------|
| 基準值 | 10.00 | 10.00 |
| 計算方式 | `USL = 10 + 0.10 = 10.10`<br>`LSL = 10 - 0.05 = 9.95` | `boundary1 = 10 + 0.10 = 10.10`<br>`boundary2 = 10 + 0.05 = 10.05`<br>`USL = max(10.10, 10.05) = 10.10`<br>`LSL = min(10.10, 10.05) = 10.05` |
| **結果** | ❌ **USL=10.10, LSL=9.95**<br>(錯誤：範圍過大) | ✅ **USL=10.10, LSL=10.05**<br>(正確：兩者皆高於基準) |

### 案例 2: 單向負公差 (-0.05 / -0.10)

| 項目 | 舊版邏輯 | 新版邏輯 |
|------|----------|----------|
| 基準值 | 10.00 | 10.00 |
| 計算方式 | `USL = 10 + 0.05 = 10.05`<br>`LSL = 10 - 0.10 = 9.90` | `boundary1 = 10 - 0.05 = 9.95`<br>`boundary2 = 10 - 0.10 = 9.90`<br>`USL = max(9.95, 9.90) = 9.95`<br>`LSL = min(9.95, 9.90) = 9.90` |
| **結果** | ❌ **USL=10.05, LSL=9.90**<br>(錯誤：上限高於基準) | ✅ **USL=9.95, LSL=9.90**<br>(正確：兩者皆低於基準) |

### 案例 3: 傳統對稱公差 (+0.10 / -0.10)

| 項目 | 舊版邏輯 | 新版邏輯 |
|------|----------|----------|
| 基準值 | 10.00 | 10.00 |
| 計算方式 | `USL = 10 + 0.10 = 10.10`<br>`LSL = 10 - 0.10 = 9.90` | `boundary1 = 10 + 0.10 = 10.10`<br>`boundary2 = 10 - 0.10 = 9.90`<br>`USL = max(10.10, 9.90) = 10.10`<br>`LSL = min(10.10, 9.90) = 9.90` |
| **結果** | ✅ **USL=10.10, LSL=9.90**<br>(正確) | ✅ **USL=10.10, LSL=9.90**<br>(正確，向後兼容) |

---

## 🎯 關鍵改進點 (Key Improvements)

### 1. **符號感知 (Sign-Aware)**
- **舊版**: 忽略符號，固定使用加減法
- **新版**: 根據符號 (`+`, `-`, `±`) 決定偏移方向

### 2. **邊界校準 (Boundary Calibration)**
- **舊版**: 無校準，可能出現 USL < LSL
- **新版**: 自動比對，確保 USL ≥ LSL

### 3. **場景支持 (Scenario Support)**
- **舊版**: 僅支持 `+/-` 和 `±`
- **新版**: 支持所有組合 (`+/+`, `-/-`, `+/-`, `±`)

### 4. **邏輯一致性 (Logical Consistency)**
- **舊版**: 依賴輸入順序，可能產生邏輯錯誤
- **新版**: 自動校準，保證邏輯正確性

---

## 📁 修改的檔案 (Modified Files)

1. **JavaScript 實現**:
   - `js/qip/spec-extractor.js` (Lines 229-258)
   - 方法: `readSpecificationFromRow()`

2. **VBA 實現**:
   - `VBACode_MouldexSingleFile/QIP_Extractor/SpecificationExtractor.bas` (Lines 234-291)
   - 函數: `ReadSpecificationFromRow()`

3. **文件**:
   - `docs/SPECIFICATION_LIMIT_CALCULATION.md` (新增)
   - `docs/SPECIFICATION_UPDATE_SUMMARY.md` (本文件)

---

## ✅ 測試建議 (Testing Recommendations)

### 必測場景 (Must-Test Scenarios)

1. **傳統對稱公差**: `+0.10 / -0.10`
2. **± 對稱公差**: `±0.10`
3. **單向正公差**: `+0.10 / +0.05`
4. **單向負公差**: `-0.05 / -0.10`
5. **不對稱公差**: `+0.15 / -0.05`
6. **反向輸入**: `-0.10 / +0.10` (測試自動校準)

### 驗證方法 (Validation Method)

1. 準備包含上述場景的 Excel 規格表
2. 使用更新後的系統讀取規格
3. 檢查控制台日誌輸出
4. 驗證 USL 和 LSL 的正確性
5. 確認 USL ≥ LSL

---

## 🔍 除錯支援 (Debugging Support)

### JavaScript 控制台日誌

新版實現包含詳細的日誌輸出：

```javascript
console.log(`[SpecExtract] 符號感知計算: ${upperSign}${upperTolVal} / ${lowerSign}${lowerTolVal}`);
console.log(`[SpecExtract] 最終規格: Nominal=${nominalValue}, USL=${spec.usl}, LSL=${spec.lsl}`);
```

**範例輸出**:
```
[SpecExtract] 符號感知計算: +0.1 / +0.05
[SpecExtract] 最終規格: Nominal=10, USL=10.1, LSL=10.05
```

---

## 📌 向後兼容性 (Backward Compatibility)

✅ **完全向後兼容**

- 傳統 `+/-` 公差格式仍然正確處理
- `±` 符號公差正常運作
- 現有的規格表無需修改
- 輸出結果與舊版一致（對於標準場景）

---

## 🚀 下一步 (Next Steps)

1. ✅ 代碼已更新 (JavaScript + VBA)
2. ✅ 文件已建立
3. ⏳ **建議**: 執行完整測試
4. ⏳ **建議**: 更新使用者手冊
5. ⏳ **建議**: 通知相關使用者

---

**更新完成**: 2026-01-14  
**維護者**: Antigravity AI
