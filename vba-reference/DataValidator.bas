Option Explicit

' ============================================================================
' Module: DataValidator
' Description: 數據驗證模組，負責工作表格式驗證
' ============================================================================

' 常量定義
Private Const BATCH_NUMBER_HEADER As String = "生產批號"
Private Const EXCLUDED_SHEETS As String = "處理異常紀錄,參數配置,配置歷史,圖表生成異常紀錄"

' ============================================================================
' 驗證工作表是否為有效的數據工作表
' ============================================================================
Public Function IsValidDataSheet(ws As Worksheet) As Boolean
    On Error GoTo ErrorHandler
    
    IsValidDataSheet = False
    
    ' 檢查工作表名稱是否在排除列表中
    If IsExcludedSheet(ws.Name) Then Exit Function
    
    ' 檢查A1儲存格是否包含「生產批號」
    If ws.Range("A1").value <> BATCH_NUMBER_HEADER Then Exit Function
    
    ' 檢查第二行是否有數據
    If Len(Trim(CStr(ws.Range("A2").value))) = 0 Then Exit Function
    
    ' 檢查是否有穴號列（至少要有B列）
    If (Len(Trim(CStr(ws.Range("B1").value))) = 0) And (Len(Trim(CStr(ws.Range("B2").value))) = 0) Then Exit Function
    
    IsValidDataSheet = True
    Exit Function
    
ErrorHandler:
    IsValidDataSheet = False
End Function

' ============================================================================
' 檢查工作表名稱是否在排除列表中
' ============================================================================
Private Function IsExcludedSheet(sheetName As String) As Boolean
    Dim excludedList() As String
    Dim i As Integer
    
    IsExcludedSheet = False
    excludedList = Split(EXCLUDED_SHEETS, ",")
    
    For i = LBound(excludedList) To UBound(excludedList)
        If Trim(excludedList(i)) = sheetName Then
            IsExcludedSheet = True
            Exit Function
        End If
    Next i
End Function

' ============================================================================
' 獲取工作表中的有效數據範圍
' ============================================================================
Public Function GetValidDataRange(ws As Worksheet) As Range
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set GetValidDataRange = Nothing
    
    ' 獲取最後一行（A列）
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Function
    
    ' 獲取最後一列
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 2 Then Exit Function
    
    ' 返回數據範圍（從A1到最後一行最後一列）
    Set GetValidDataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    Exit Function
    
ErrorHandler:
    Set GetValidDataRange = Nothing
End Function

' ============================================================================
' 獲取穴號列的數量
' ============================================================================
Public Function GetCavityColumnCount(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    
    Dim col As Long
    Dim cavityCount As Long
    Dim lastCol As Long
    Dim cellValue As String
    
    GetCavityColumnCount = 0
    cavityCount = 0
    
    ' 獲取最後一列
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    ' 從第2列開始計數，直到遇到統計列（如「最大值」、「最小值」等）
    For col = 2 To lastCol
        cellValue = Trim(ws.Cells(1, col).value)
        
        ' 檢查是否為統計列
        If IsStatisticsColumn(cellValue) Then
            Exit For
        End If
        
        cavityCount = cavityCount + 1
    Next col
    
    GetCavityColumnCount = cavityCount
    Exit Function
    
ErrorHandler:
    GetCavityColumnCount = 0
End Function

' ============================================================================
' 檢查是否為統計列
' ============================================================================
Private Function IsStatisticsColumn(columnHeader As String) As Boolean
    Dim statsKeywords As Variant
    Dim i As Integer
    
    IsStatisticsColumn = False
    statsKeywords = Array("最大值", "最小值", "平均值", "標準差", "範圍", _
                          "Max", "Min", "Average", "Avg", "StdDev", "Range")
    
    For i = LBound(statsKeywords) To UBound(statsKeywords)
        If InStr(1, columnHeader, statsKeywords(i), vbTextCompare) > 0 Then
            IsStatisticsColumn = True
            Exit Function
        End If
    Next i
End Function

' ============================================================================
' 獲取穴號列的起始列號
' ============================================================================
Public Function GetCavityStartColumn() As Long
    GetCavityStartColumn = 2  ' 穴號從第2列（B列）開始
End Function

' ============================================================================
' 獲取數據起始行號
' ============================================================================
Public Function GetDataStartRow() As Long
    GetDataStartRow = 3  ' 數據從第3行開始（第1行標題，第2行空白，第3行開始是數據）
End Function




