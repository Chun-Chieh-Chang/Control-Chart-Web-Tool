# 規格上下限計算邏輯說明
## Specification Limit Calculation Logic Documentation

**更新日期**: 2026-01-14  
**版本**: 2.0 - 符號感知計算 (Sign-Aware Calculation)

---

## 📋 概述 (Overview)

本專案已升級規格上下限 (USL/LSL) 的計算邏輯，從簡單的對稱公差處理升級為**完整的符號感知計算系統**，能夠正確處理各種複雜的公差場景。

---

## 🎯 核心特性 (Core Features)

### 1. **符號感知計算 (Sign-Aware Calculation)**

系統不再僅僅讀取公差的絕對值，而是結合左側符號欄位來決定偏移方向：

| 符號 | 意義 | 計算方式 |
|------|------|----------|
| `+` | 正偏移 | `基準值 + 公差值` |
| `-` | 負偏移 | `基準值 - 公差值` |
| `±` | 雙向對稱偏移 | `基準值 ± 公差值` |

### 2. **支持複雜公差場景 (Complex Tolerance Scenarios)**

透過符號感知邏輯，系統現在可以正確處理非標準的公差組合：

#### 場景 A：單向正公差 (Single-Sided Positive Tolerance)
```
基準值: 10.00
上公差: +0.10
下公差: +0.05

計算結果:
  boundary1 = 10.00 + 0.10 = 10.10
  boundary2 = 10.00 + 0.05 = 10.05
  USL = 10.10 (較大值)
  LSL = 10.05 (較小值)
```

#### 場景 B：單向負公差 (Single-Sided Negative Tolerance)
```
基準值: 10.00
上公差: -0.05
下公差: -0.10

計算結果:
  boundary1 = 10.00 - 0.05 = 9.95
  boundary2 = 10.00 - 0.10 = 9.90
  USL = 9.95 (較大值)
  LSL = 9.90 (較小值)
```

#### 場景 C：傳統對稱公差 (Traditional Symmetric Tolerance)
```
基準值: 10.00
上公差: +0.10
下公差: -0.10

計算結果:
  boundary1 = 10.00 + 0.10 = 10.10
  boundary2 = 10.00 - 0.10 = 9.90
  USL = 10.10 (較大值)
  LSL = 9.90 (較小值)
```

#### 場景 D：± 符號對稱公差 (± Symbol Symmetric Tolerance)
```
基準值: 10.00
公差: ±0.10

計算結果:
  boundary1 = 10.00 + 0.10 = 10.10
  boundary2 = 10.00 - 0.10 = 9.90
  USL = 10.10 (較大值)
  LSL = 9.90 (較小值)
```

#### 場景 E：不對稱公差 (Asymmetric Tolerance)
```
基準值: 10.00
上公差: +0.15
下公差: -0.05

計算結果:
  boundary1 = 10.00 + 0.15 = 10.15
  boundary2 = 10.00 - 0.05 = 9.95
  USL = 10.15 (較大值)
  LSL = 9.95 (較小值)
```

### 3. **自動邊界校準 (Automatic Boundary Calibration)**

在計算出兩個偏移邊界後，系統會自動比對並將：
- **較大值** 設為 **USL (上規格限)**
- **較小值** 設為 **LSL (下規格限)**

這確保了數據輸出的邏輯一致性，**避免因輸入順序導致的上限小於下限的情形**。

---

## 🔧 技術實現 (Technical Implementation)

### JavaScript 實現 (`js/qip/spec-extractor.js`)

```javascript
// 符號感知計算：根據符號決定偏移方向
let boundary1, boundary2;

if (upperSign === '±') {
    // ± 符號：雙向對稱偏移
    boundary1 = nominalValue + upperTolVal;
    boundary2 = nominalValue - upperTolVal;
} else {
    // 根據符號計算兩個邊界
    if (upperSign === '+') {
        boundary1 = nominalValue + upperTolVal;
    } else if (upperSign === '-') {
        boundary1 = nominalValue - upperTolVal;
    } else {
        boundary1 = nominalValue + upperTolVal; // 默認正偏移
    }

    if (lowerSign === '+') {
        boundary2 = nominalValue + lowerTolVal;
    } else if (lowerSign === '-') {
        boundary2 = nominalValue - lowerTolVal;
    } else {
        boundary2 = nominalValue - lowerTolVal; // 默認負偏移
    }
}

// 自動邊界校準：確保 USL > LSL
spec.usl = Math.max(boundary1, boundary2);
spec.lsl = Math.min(boundary1, boundary2);
spec.upperTolerance = spec.usl - nominalValue;
spec.lowerTolerance = spec.lsl - nominalValue;
```

### VBA 實現 (`VBACode_MouldexSingleFile/QIP_Extractor/SpecificationExtractor.bas`)

```vba
Dim boundary1 As Double, boundary2 As Double

' 處理 ± 對稱公差
If upperSign = "±" Then
    boundary1 = spec.NominalValue + upperTol
    boundary2 = spec.NominalValue - upperTol
Else
    ' 根據符號計算第一個邊界（上公差行）
    If upperSign = "+" Then
        boundary1 = spec.NominalValue + upperTol
    ElseIf upperSign = "-" Then
        boundary1 = spec.NominalValue - upperTol
    Else
        boundary1 = spec.NominalValue + upperTol ' 默認正偏移
    End If
    
    ' 根據符號計算第二個邊界（下公差行）
    If lowerSign = "+" Then
        boundary2 = spec.NominalValue + lowerTol
    ElseIf lowerSign = "-" Then
        boundary2 = spec.NominalValue - lowerTol
    Else
        boundary2 = spec.NominalValue - lowerTol ' 默認負偏移
    End If
End If

' 自動邊界校準：確保 USL > LSL
If boundary1 > boundary2 Then
    spec.usl = boundary1
    spec.lsl = boundary2
Else
    spec.usl = boundary2
    spec.lsl = boundary1
End If

' 計算最終的公差值
spec.UpperTolerance = spec.usl - spec.NominalValue
spec.LowerTolerance = spec.lsl - spec.NominalValue
```

---

## 📊 Excel 規格表格式 (Excel Specification Format)

系統從 Excel 規格表讀取以下欄位：

| 欄位 | 說明 | 範例 |
|------|------|------|
| E/F 列 | 基準值 (Nominal Value) | 10.00 |
| G 列 (第一行) | 上公差符號 | `+`, `-`, `±` |
| H 列 (第一行) | 上公差數值 | 0.10 |
| G 列 (第二行) | 下公差符號 | `+`, `-` |
| H 列 (第二行) | 下公差數值 | 0.05 |

**注意**: 公差數值欄位 (H列) 中的數值會被取絕對值，符號由 G 列決定。

---

## ✅ 驗證與測試 (Validation & Testing)

### 測試案例 (Test Cases)

| 測試案例 | 基準值 | 上公差 | 下公差 | 預期 USL | 預期 LSL |
|----------|--------|--------|--------|----------|----------|
| 傳統對稱 | 10.00 | +0.10 | -0.10 | 10.10 | 9.90 |
| ± 對稱 | 10.00 | ±0.10 | - | 10.10 | 9.90 |
| 單向正公差 | 10.00 | +0.10 | +0.05 | 10.10 | 10.05 |
| 單向負公差 | 10.00 | -0.05 | -0.10 | 9.95 | 9.90 |
| 不對稱 | 10.00 | +0.15 | -0.05 | 10.15 | 9.95 |
| 反向輸入 | 10.00 | -0.10 | +0.10 | 10.10 | 9.90 |

### 控制台日誌 (Console Logging)

JavaScript 實現包含詳細的控制台日誌，方便除錯：

```
[SpecExtract] 符號感知計算: +0.1 / -0.1
[SpecExtract] 最終規格: Nominal=10, USL=10.1, LSL=9.9
```

---

## 🔄 版本歷史 (Version History)

### v2.0 (2026-01-14) - 符號感知計算
- ✅ 實現完整的符號感知計算邏輯
- ✅ 支持單向公差 (+/+, -/-)
- ✅ 支持 ± 對稱公差
- ✅ 自動邊界校準 (USL > LSL)
- ✅ JavaScript 和 VBA 雙重實現

### v1.0 (之前) - 基礎公差計算
- 僅支持傳統 +/- 公差
- 有限的 ± 支持
- 無自動邊界校準

---

## 📝 注意事項 (Notes)

1. **符號優先級**: 系統優先讀取 G 列的符號，如果符號欄位為空，則使用默認值（上公差默認 `+`，下公差默認 `-`）

2. **數值處理**: H 列的公差數值會被取絕對值，確保計算的一致性

3. **邊界校準**: 無論輸入順序如何，系統都會自動確保 USL ≥ LSL

4. **向後兼容**: 新邏輯完全向後兼容傳統的 +/- 公差格式

---

## 🎓 參考資料 (References)

- ISO 286-1: Geometrical product specifications (GPS) — ISO code system for tolerances on linear sizes
- ASME Y14.5: Dimensioning and Tolerancing
- GD&T (Geometric Dimensioning and Tolerancing) Standards

---

**文件維護**: Antigravity AI  
**最後更新**: 2026-01-14
