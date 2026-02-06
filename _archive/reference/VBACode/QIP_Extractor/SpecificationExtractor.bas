Option Explicit

' ============================================================================
' 規格數據結構定義
' ============================================================================
Public Type SpecificationData
    Symbol As String
    NominalValue As Double
    UpperTolerance As Double
    LowerTolerance As Double
    usl As Double
    lsl As Double
    target As Double
    IsValid As Boolean
End Type

' ============================================================================
' 規格數據自動提取輔助函數
' ============================================================================

' 尋找包含規格數據的工作表
Public Function FindSpecificationWorksheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' 優先尋找包含規格關鍵字的工作表
    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, "規格", vbTextCompare) > 0 Or _
           InStr(1, ws.Name, "spec", vbTextCompare) > 0 Or _
           InStr(1, ws.Name, "標準", vbTextCompare) > 0 Then
            Set FindSpecificationWorksheet = ws
            Exit Function
        End If
    Next ws
    
    ' 如果沒有找到專門的規格工作表，檢查是否有包含規格數據的工作表
    For Each ws In wb.Worksheets
        ' 跳過特殊工作表
        If ws.Name <> "處理異常紀錄" And ws.Name <> "參數配置" And ws.Name <> "配置歷史" Then
            ' 檢查工作表是否包含規格數據格式
            If HasSpecificationData(ws) Then
                Set FindSpecificationWorksheet = ws
                Exit Function
            End If
        End If
    Next ws
    
    ' 如果都沒找到，返回第一個工作表
    If wb.Worksheets.count > 0 Then
        Set FindSpecificationWorksheet = wb.Worksheets(1)
    End If
    
    Exit Function
    
ErrorHandler:
    Set FindSpecificationWorksheet = Nothing
End Function

' 檢查工作表是否包含規格數據
Private Function HasSpecificationData(ws As Worksheet) As Boolean
    Dim i As Long
    Dim cellValue As String
    
    On Error GoTo ErrorHandler
    
    ' 檢查前100行是否有檢驗項目標識
    For i = 1 To 100
        ' 使用 .Text 而非 .Value 以避免 Excel 格式代碼干擾
        On Error Resume Next
        cellValue = Trim(ws.Cells(i, 1).Text)
        If cellValue = "" Then cellValue = Trim(CStr(ws.Cells(i, 1).value))
        On Error GoTo ErrorHandler
        
        If Len(cellValue) > 0 And Left(cellValue, 1) = "(" And Right(cellValue, 1) = ")" Then
            ' 檢查是否有對應的規格數據格式
            If ws.Cells(i, 4).value <> "" And IsNumeric(ws.Cells(i, 5).value) Then
                HasSpecificationData = True
                Exit Function
            End If
        End If
    Next i
    
    HasSpecificationData = False
    Exit Function
    
ErrorHandler:
    HasSpecificationData = False
End Function

' 根據檢驗項目查找規格數據
Public Function FindSpecificationByItem(ws As Worksheet, itemName As String) As SpecificationData
    Dim spec As SpecificationData
    Dim i As Long
    Dim cellValue As String
    Dim cleanItemName As String
    
    On Error GoTo ErrorHandler
    
    ' 初始化無效規格
    spec.IsValid = False
    
    ' 清理檢驗項目名稱（移除括號和空格）
    cleanItemName = Replace(Replace(itemName, "(", ""), ")", "")
    cleanItemName = Trim(cleanItemName)
    
    ' 在工作表中尋找對應的檢驗項目
    For i = 1 To 100
        ' 使用 .Text 而非 .Value 以避免 Excel 格式代碼干擾（如 "J" 被識別為日期格式）
        On Error Resume Next
        cellValue = Trim(ws.Cells(i, 1).Text)
        If cellValue = "" Then cellValue = Trim(CStr(ws.Cells(i, 1).value))
        On Error GoTo ErrorHandler
        
        ' 清理規格表中的檢驗項目名稱（移除括號和空格）
        Dim CleanCellValue As String
        CleanCellValue = Replace(Replace(cellValue, "(", ""), ")", "")
        CleanCellValue = Trim(CleanCellValue)
        
        ' 檢查是否匹配檢驗項目（比較清理後的名稱）
        If CleanCellValue = cleanItemName Then
            ' 找到匹配的檢驗項目，讀取該行的規格數據
            ' 注意：如果該行的C列（檢測工具代碼）為空，ReadSpecificationFromRow會返回無效規格
            spec = ReadSpecificationFromRow(ws, i)
            If spec.IsValid Then
                FindSpecificationByItem = spec
                Exit Function
            End If
        End If
    Next i
    
    FindSpecificationByItem = spec
    Exit Function
    
ErrorHandler:
    spec.IsValid = False
    FindSpecificationByItem = spec
End Function

' 從指定行讀取規格數據
Private Function ReadSpecificationFromRow(ws As Worksheet, row As Long) As SpecificationData
    Dim spec As SpecificationData
    Dim toolCol As Long, symbolCol As Long, nominalCol1 As Long, nominalCol2 As Long
    Dim upperSignCol As Long, upperTolCol As Long, lowerSignCol As Long, lowerTolCol As Long
    Dim upperTol As Double, lowerTol As Double
    Dim upperSign As String, lowerSign As String
    
    On Error GoTo ErrorHandler
    
    ' 初始化
    spec.IsValid = False
    
    ' 定位規格數據的列位置
    toolCol = 3        ' C列：檢測工具代碼
    symbolCol = 4      ' D列：規格符號（O、R等）
    nominalCol1 = 5    ' E列：基準值（合併欄位第一部分）
    nominalCol2 = 6    ' F列：基準值（合併欄位第二部分）
    upperSignCol = 7   ' G列：上公差符號（第一行）
    upperTolCol = 8    ' H列：上公差數值（第一行）
    lowerSignCol = 7   ' G列：下公差符號（第二行）
    lowerTolCol = 8    ' H列：下公差數值（第二行）
    
    ' 檢查是否有檢測工具代碼（用於識別規格行）
    If Trim(CStr(ws.Cells(row, toolCol).value)) = "" Then
        Exit Function ' 沒有檢測工具代碼，跳過此行
    End If
    
    ' 讀取規格符號
    spec.Symbol = Trim(CStr(ws.Cells(row, symbolCol).value))
    
    ' 讀取基準值（從E欄，基準值跨E、F欄，且可能是合併儲存格跨兩行）
    ' 使用 Text 屬性避免 Excel 將 "J" 等字符誤判為日期格式
    ' 搜尋順序：E(row) -> F(row) -> E(row+1) -> F(row+1)
    Dim tempValue As Variant
    On Error Resume Next
    
    ' 嘗試從 E(row) 讀取
    tempValue = ws.Cells(row, nominalCol1).Text
    If IsNumeric(tempValue) And tempValue <> "" Then
        spec.NominalValue = CDbl(tempValue)
        spec.target = spec.NominalValue
    ' 嘗試從 F(row) 讀取
    ElseIf IsNumeric(ws.Cells(row, nominalCol2).Text) And ws.Cells(row, nominalCol2).Text <> "" Then
        spec.NominalValue = CDbl(ws.Cells(row, nominalCol2).Text)
        spec.target = spec.NominalValue
    ' 嘗試從 E(row+1) 讀取
    ElseIf IsNumeric(ws.Cells(row + 1, nominalCol1).Text) And ws.Cells(row + 1, nominalCol1).Text <> "" Then
        spec.NominalValue = CDbl(ws.Cells(row + 1, nominalCol1).Text)
        spec.target = spec.NominalValue
    ' 嘗試從 F(row+1) 讀取
    ElseIf IsNumeric(ws.Cells(row + 1, nominalCol2).Text) And ws.Cells(row + 1, nominalCol2).Text <> "" Then
        spec.NominalValue = CDbl(ws.Cells(row + 1, nominalCol2).Text)
        spec.target = spec.NominalValue
    Else
        ' 如果 Text 都失敗，回退到 Value
        If IsNumeric(ws.Cells(row, nominalCol1).value) And Not IsEmpty(ws.Cells(row, nominalCol1).value) Then
            spec.NominalValue = CDbl(ws.Cells(row, nominalCol1).value)
            spec.target = spec.NominalValue
        ElseIf IsNumeric(ws.Cells(row, nominalCol2).value) And Not IsEmpty(ws.Cells(row, nominalCol2).value) Then
            spec.NominalValue = CDbl(ws.Cells(row, nominalCol2).value)
            spec.target = spec.NominalValue
        ElseIf IsNumeric(ws.Cells(row + 1, nominalCol1).value) And Not IsEmpty(ws.Cells(row + 1, nominalCol1).value) Then
            spec.NominalValue = CDbl(ws.Cells(row + 1, nominalCol1).value)
            spec.target = spec.NominalValue
        ElseIf IsNumeric(ws.Cells(row + 1, nominalCol2).value) And Not IsEmpty(ws.Cells(row + 1, nominalCol2).value) Then
            spec.NominalValue = CDbl(ws.Cells(row + 1, nominalCol2).value)
            spec.target = spec.NominalValue
        Else
            On Error GoTo ErrorHandler
            Exit Function ' 基準值無效，跳過此規格
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' 讀取上公差（第一行的G、H欄）
    upperSign = Trim(CStr(ws.Cells(row, upperSignCol).value))
    If IsNumeric(ws.Cells(row, upperTolCol).value) Then
        upperTol = CDbl(ws.Cells(row, upperTolCol).value)
        upperTol = Abs(upperTol) ' 確保為正數
    Else
        upperTol = 0
    End If
    
    ' 讀取下公差（第二行的G、H欄）
    lowerSign = Trim(CStr(ws.Cells(row + 1, lowerSignCol).value))
    If IsNumeric(ws.Cells(row + 1, lowerTolCol).value) Then
        lowerTol = CDbl(ws.Cells(row + 1, lowerTolCol).value)
        lowerTol = Abs(lowerTol) ' 確保為正數
    Else
        lowerTol = 0
    End If
    
    ' ============================================================================
    ' 符號感知公差計算 (Sign-Aware Tolerance Calculation)
    ' ============================================================================
    ' 系統根據符號欄位決定偏移方向：
    ' + 符號：視為相對於基準值的 正偏移 (加上公差值)
    ' - 符號：視為相對於基準值的 負偏移 (減去公差值)
    ' ± 符號：視為 雙向對稱偏移
    ' ============================================================================
    
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
            ' 未知符號，默認為正偏移
            boundary1 = spec.NominalValue + upperTol
        End If
        
        ' 根據符號計算第二個邊界（下公差行）
        If lowerSign = "+" Then
            boundary2 = spec.NominalValue + lowerTol
        ElseIf lowerSign = "-" Then
            boundary2 = spec.NominalValue - lowerTol
        Else
            ' 未知符號，默認為負偏移
            boundary2 = spec.NominalValue - lowerTol
        End If
    End If
    
    ' ============================================================================
    ' 自動邊界校準：確保 USL > LSL
    ' ============================================================================
    ' 比對兩個邊界值，較大值設為 USL，較小值設為 LSL
    ' 這樣可以正確處理各種公差組合，包括：
    ' - 單向公差：+0.1 / +0.05 (兩者皆高於基準)
    ' - 單向公差：-0.05 / -0.1 (兩者皆低於基準)
    ' - 傳統公差：+0.1 / -0.1
    ' ============================================================================
    
    If boundary1 > boundary2 Then
        spec.usl = boundary1
        spec.lsl = boundary2
    Else
        spec.usl = boundary2
        spec.lsl = boundary1
    End If
    
    ' 計算最終的公差值（相對於基準值的偏移量）
    spec.UpperTolerance = spec.usl - spec.NominalValue
    spec.LowerTolerance = spec.lsl - spec.NominalValue
    spec.IsValid = True
    
    ReadSpecificationFromRow = spec
    Exit Function
    
ErrorHandler:
    spec.IsValid = False
    ReadSpecificationFromRow = spec
End Function

