Option Explicit

' ============================================================================
' Module: DataExtractor
' Description: 數據提取模組，負責從工作表中提取數據
' ============================================================================

' ============================================================================
' 提取所有生產批號
' ============================================================================
Public Function ExtractBatchNumbers(ws As Worksheet) As Variant
    On Error GoTo ErrorHandler
    Dim lastRow As Long
    Dim rng As Range
    Dim arr As Variant
    Dim i As Long
    Dim rowIndex As Long
    Dim result() As Variant
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < DataValidator.GetDataStartRow() Then
        ExtractBatchNumbers = Array()
        Exit Function
    End If
    Set rng = ws.Range(ws.Cells(DataValidator.GetDataStartRow(), 1), ws.Cells(lastRow, 1))
    arr = rng.value
    ReDim result(1 To rng.Rows.count)
    rowIndex = 1
    For i = 1 To rng.Rows.count
        If Not IsEmpty(arr(i, 1)) Then
            result(rowIndex) = CStr(arr(i, 1))
            rowIndex = rowIndex + 1
        End If
    Next i
    If rowIndex > 1 Then
        ReDim Preserve result(1 To rowIndex - 1)
        ExtractBatchNumbers = result
    Else
        ExtractBatchNumbers = Array()
    End If
    Exit Function
ErrorHandler:
    ExtractBatchNumbers = Array()
End Function

' ============================================================================
' 提取每個批次的平均值
' ============================================================================
Public Function ExtractAverageValues(ws As Worksheet) As Variant
    On Error GoTo ErrorHandler
    Dim lastRow As Long
    Dim cavityCount As Long
    Dim startCol As Long
    Dim rng As Range
    Dim arr As Variant
    Dim r As Long
    Dim c As Long
    Dim sum As Double
    Dim cnt As Long
    Dim result() As Variant
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    cavityCount = DataValidator.GetCavityColumnCount(ws)
    If lastRow < DataValidator.GetDataStartRow() Or cavityCount = 0 Then
        ExtractAverageValues = Array()
        Exit Function
    End If
    startCol = DataValidator.GetCavityStartColumn()
    Set rng = ws.Range(ws.Cells(DataValidator.GetDataStartRow(), startCol), ws.Cells(lastRow, startCol + cavityCount - 1))
    arr = rng.value
    ReDim result(1 To rng.Rows.count)
    For r = 1 To rng.Rows.count
        sum = 0
        cnt = 0
        For c = 1 To rng.Columns.count
            If IsNumeric(arr(r, c)) And Not IsEmpty(arr(r, c)) Then
                sum = sum + CDbl(arr(r, c))
                cnt = cnt + 1
            End If
        Next c
        If cnt > 0 Then
            result(r) = sum / cnt
        Else
            result(r) = 0
        End If
    Next r
    ExtractAverageValues = result
    Exit Function
ErrorHandler:
    ExtractAverageValues = Array()
End Function

' ============================================================================
' 計算單行的平均值
' ============================================================================
Private Function CalculateRowAverage(ws As Worksheet, rowNum As Long, cavityCount As Long) As Double
    Dim col As Long
    Dim sum As Double
    Dim count As Long
    Dim cellValue As Variant
    Dim startCol As Long
    
    sum = 0
    count = 0
    startCol = DataValidator.GetCavityStartColumn()
    
    ' 遍歷穴號列
    For col = startCol To startCol + cavityCount - 1
        cellValue = ws.Cells(rowNum, col).value
        
        If IsNumeric(cellValue) And Not IsEmpty(cellValue) Then
            sum = sum + CDbl(cellValue)
            count = count + 1
        End If
    Next col
    
    ' 計算平均值
    If count > 0 Then
        CalculateRowAverage = sum / count
    Else
        CalculateRowAverage = 0
    End If
End Function

' ============================================================================
' 提取特定穴號的所有數據
' ============================================================================
Public Function ExtractCavityData(ws As Worksheet, cavityIndex As Long) As Variant
    On Error GoTo ErrorHandler
    Dim lastRow As Long
    Dim startRow As Long
    Dim colNum As Long
    Dim rng As Range
    Dim arr As Variant
    Dim r As Long
    Dim result() As Variant
    colNum = DataValidator.GetCavityStartColumn() + cavityIndex - 1
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    startRow = DataValidator.GetDataStartRow()
    If lastRow < startRow Then
        ExtractCavityData = Array()
        Exit Function
    End If
    Set rng = ws.Range(ws.Cells(startRow, colNum), ws.Cells(lastRow, colNum))
    arr = rng.value
    ReDim result(1 To rng.Rows.count)
    For r = 1 To rng.Rows.count
        If IsNumeric(arr(r, 1)) And Not IsEmpty(arr(r, 1)) Then
            result(r) = CDbl(arr(r, 1))
        Else
            result(r) = Empty
        End If
    Next r
    ExtractCavityData = result
    Exit Function
ErrorHandler:
    ExtractCavityData = Array()
End Function

' ============================================================================
' 獲取所有穴號的標題
' ============================================================================
Public Function GetCavityHeaders(ws As Worksheet) As Variant
    On Error GoTo ErrorHandler
    
    Dim cavityCount As Long
    Dim headers() As String
    Dim i As Long
    Dim colNum As Long
    Dim headerValue As Variant
    
    cavityCount = DataValidator.GetCavityColumnCount(ws)
    
    If cavityCount = 0 Then
        GetCavityHeaders = Array()
        Exit Function
    End If
    
    ' 初始化數組
    ReDim headers(1 To cavityCount)
    
    ' 提取標題
    For i = 1 To cavityCount
        colNum = DataValidator.GetCavityStartColumn() + i - 1
        headerValue = ws.Cells(1, colNum).value
        
        ' 如果標題為空，使用列號作為標識
        If IsEmpty(headerValue) Or Trim(CStr(headerValue)) = "" Then
            headers(i) = "穴" & i
        Else
            headers(i) = CStr(headerValue)
        End If
    Next i
    
    GetCavityHeaders = headers
    Exit Function
    
ErrorHandler:
    GetCavityHeaders = Array()
End Function


