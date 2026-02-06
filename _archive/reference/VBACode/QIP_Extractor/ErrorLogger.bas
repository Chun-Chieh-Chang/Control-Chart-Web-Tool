Option Explicit

' ============================================================================
' Module: ErrorLogger
' Description: 錯誤日誌模組，負責記錄處理異常
' ============================================================================

' 常量定義
Private Const ERROR_LOG_SHEET_NAME As String = "圖表生成異常紀錄"
Private Const HEADER_SHEET_NAME As String = "工作表名稱"
Private Const HEADER_ERROR_TYPE As String = "錯誤類型"
Private Const HEADER_ERROR_MESSAGE As String = "錯誤訊息"
Private Const HEADER_TIMESTAMP As String = "發生時間"

' ============================================================================
' 初始化錯誤日誌工作表
' ============================================================================
Public Sub InitializeErrorLog(wb As Workbook)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim exists As Boolean
    
    ' 檢查錯誤日誌工作表是否已存在
    exists = False
    On Error Resume Next
    Set ws = wb.Worksheets(ERROR_LOG_SHEET_NAME)
    exists = (Err.Number = 0)
    On Error GoTo ErrorHandler
    
    ' 如果不存在，創建新工作表
    If Not exists Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = ERROR_LOG_SHEET_NAME
        
        ' 設置表頭
        With ws
            .Cells(1, 1).value = HEADER_SHEET_NAME
            .Cells(1, 2).value = HEADER_ERROR_TYPE
            .Cells(1, 3).value = HEADER_ERROR_MESSAGE
            .Cells(1, 4).value = HEADER_TIMESTAMP
            
            ' 格式化表頭
            With .Range("A1:D1")
                .Font.Bold = True
                .Interior.color = RGB(217, 217, 217)
                .HorizontalAlignment = xlCenter
            End With
            
            ' 設置列寬
            .Columns("A:A").ColumnWidth = 20
            .Columns("B:B").ColumnWidth = 15
            .Columns("C:C").ColumnWidth = 50
            .Columns("D:D").ColumnWidth = 20
        End With
    End If
    
    Exit Sub
    
ErrorHandler:
    ' 如果無法創建錯誤日誌，靜默失敗
End Sub

' ============================================================================
' 記錄錯誤信息
' ============================================================================
Public Sub LogError(wb As Workbook, sheetName As String, ErrorMessage As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim errorType As String
    
    ' 初始化錯誤日誌（如果尚未初始化）
    InitializeErrorLog wb
    
    ' 獲取錯誤日誌工作表
    Set ws = wb.Worksheets(ERROR_LOG_SHEET_NAME)
    
    ' 獲取最後一行
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow = 1 And ws.Cells(1, 1).value = HEADER_SHEET_NAME Then
        lastRow = 1
    End If
    
    ' 確定錯誤類型
    errorType = DetermineErrorType(ErrorMessage)
    
    ' 寫入錯誤信息
    With ws
        .Cells(lastRow + 1, 1).value = sheetName
        .Cells(lastRow + 1, 2).value = errorType
        .Cells(lastRow + 1, 3).value = ErrorMessage
        .Cells(lastRow + 1, 4).value = Now
        .Cells(lastRow + 1, 4).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End With
    
    Exit Sub
    
ErrorHandler:
    ' 如果無法記錄錯誤，靜默失敗
End Sub

' ============================================================================
' 檢查是否有錯誤記錄
' ============================================================================
Public Function hasErrors(wb As Workbook) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    hasErrors = False
    
    ' 檢查錯誤日誌工作表是否存在
    On Error Resume Next
    Set ws = wb.Worksheets(ERROR_LOG_SHEET_NAME)
    If Err.Number <> 0 Then
        hasErrors = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' 檢查是否有錯誤記錄（除了表頭）
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    hasErrors = (lastRow > 1)
    
    Exit Function
    
ErrorHandler:
    hasErrors = False
End Function

' ============================================================================
' 確定錯誤類型
' ============================================================================
Private Function DetermineErrorType(ErrorMessage As String) As String
    Dim msg As String
    msg = LCase(ErrorMessage)
    
    If InStr(msg, "格式") > 0 Or InStr(msg, "format") > 0 Then
        DetermineErrorType = "格式錯誤"
    ElseIf InStr(msg, "數據") > 0 Or InStr(msg, "data") > 0 Then
        DetermineErrorType = "數據錯誤"
    ElseIf InStr(msg, "圖表") > 0 Or InStr(msg, "chart") > 0 Then
        DetermineErrorType = "圖表錯誤"
    ElseIf InStr(msg, "工作表") > 0 Or InStr(msg, "worksheet") > 0 Then
        DetermineErrorType = "工作表錯誤"
    Else
        DetermineErrorType = "一般錯誤"
    End If
End Function

' ============================================================================
' 清除錯誤日誌
' ============================================================================
Public Sub ClearErrorLog(wb As Workbook)
    On Error Resume Next
    
    Dim ws As Worksheet
    
    ' 刪除錯誤日誌工作表
    Application.DisplayAlerts = False
    Set ws = wb.Worksheets(ERROR_LOG_SHEET_NAME)
    If Not ws Is Nothing Then
        ws.Delete
    End If
    Application.DisplayAlerts = True
End Sub

' ============================================================================
' 獲取錯誤數量
' ============================================================================
Public Function GetErrorCount(wb As Workbook) As Long
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    
    GetErrorCount = 0
    
    ' 檢查錯誤日誌工作表是否存在
    On Error Resume Next
    Set ws = wb.Worksheets(ERROR_LOG_SHEET_NAME)
    If Err.Number <> 0 Then
        GetErrorCount = 0
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' 計算錯誤數量（除了表頭）
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow > 1 Then
        GetErrorCount = lastRow - 1
    Else
        GetErrorCount = 0
    End If
    
    Exit Function
    
ErrorHandler:
    GetErrorCount = 0
End Function


