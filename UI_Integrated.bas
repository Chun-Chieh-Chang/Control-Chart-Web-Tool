Option Explicit

' ============================================================================
' 整合分析 UI 介面
' 提供批號分析和模穴分析的統一操作介面
' 流程：選擇檔案 -> 選擇檢驗項目 -> 選擇分析類型
' ============================================================================

' 全局變數：儲存用戶選擇
Private g_SelectedWorkbook As Workbook
Private g_SelectedItem As String
Private g_SelectedFilePath As String

' ============================================================================
' 創建整合分析 UI
' ============================================================================
Public Sub 創建整合分析介面()
    Dim ws As Worksheet
    Dim wsName As String
    
    wsName = "整合分析"
    
    ' 檢查工作表是否已存在
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' 創建新工作表
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = wsName
    Else
        ' 清除現有內容
        ws.Cells.Clear
    End If
    
    ' 設定工作表
    Call 設計UI介面(ws)
    
    ' 激活工作表
    ws.Activate
    ws.Range("A1").Select
End Sub

' ============================================================================
' 設計 UI 介面
' ============================================================================
Private Sub 設計UI介面(ws As Worksheet)
    ' 標題區域
    ws.Cells(2, 2).Value = "QIP SPC 分析系統 - 整合版"
    ws.Cells(2, 2).Font.Size = 18
    ws.Cells(2, 2).Font.Bold = True
    ws.Cells(2, 2).Font.Color = RGB(0, 112, 192)
    
    ws.Cells(2, 2).Value = "批號分析 + 模穴分析"
    ws.Cells(3, 2).Font.Size = 12
    ws.Cells(3, 2).Font.Color = RGB(128, 128, 128)
    
    ' 分隔線
    ws.Range("B5:H5").Borders(xlEdgeBottom).Weight = xlMedium
    ws.Range("B5:H5").Borders(xlEdgeBottom).Color = RGB(0, 112, 192)
    
    ' 步驟 1：選擇數據檔案
    ws.Cells(7, 2).Value = "步驟 1: 選擇數據檔案"
    ws.Cells(7, 2).Font.Size = 14
    ws.Cells(7, 2).Font.Bold = True
    ws.Cells(7, 2).Font.Color = RGB(0, 112, 192)
    
    ws.Cells(8, 3).Value = "選擇要分析的數據檔案 (第一檔案 QIP_DataExtract 輸出)"
    ws.Cells(8, 3).Font.Size = 10
    ws.Cells(8, 3).Font.Color = RGB(100, 100, 100)
    
    Call 創建按鈕(ws, 10, 3, "選擇檔案", "SelectFile")
    
    ' 顯示已選擇的檔案（放在 E 欄）
    ws.Cells(10, 5).Value = "未選擇"
    ws.Cells(10, 5).Name = "SelectedFile"
    ws.Cells(10, 5).Font.Size = 10
    ws.Cells(10, 5).Font.Color = RGB(150, 150, 150)
    
    ' 重置按鈕 (G欄)
    Call 創建按鈕(ws, 10, 7, "重置系統", "ResetSystem")
    
    ' 步驟 2：選擇檢驗項目
    ws.Cells(13, 2).Value = "步驟 2: 選擇檢驗項目"
    ws.Cells(13, 2).Font.Size = 14
    ws.Cells(13, 2).Font.Bold = True
    ws.Cells(13, 2).Font.Color = RGB(0, 112, 192)
    
    ws.Cells(14, 3).Value = "從已選擇的檔案中選擇要分析的檢驗項目"
    ws.Cells(14, 3).Font.Size = 10
    ws.Cells(14, 3).Font.Color = RGB(100, 100, 100)
    
    Call 創建按鈕(ws, 16, 3, "選擇檢驗項目", "SelectItem")
    
    ' 顯示已選擇的檢驗項目（放在 E 欄）
    ws.Cells(16, 5).Value = "未選擇"
    ws.Cells(16, 5).Name = "SelectedItem"
    ws.Cells(16, 5).Font.Size = 10
    ws.Cells(16, 5).Font.Color = RGB(150, 150, 150)
    
    ' 步驟 3：選擇分析類型
    ws.Cells(19, 2).Value = "步驟 3: 選擇分析類型"
    ws.Cells(19, 2).Font.Size = 14
    ws.Cells(19, 2).Font.Bold = True
    ws.Cells(19, 2).Font.Color = RGB(0, 112, 192)
    
    ' 批號分析 (C欄)
    ws.Cells(21, 3).Value = "批號分析"
    ws.Cells(21, 3).Font.Bold = True
    ws.Cells(21, 3).Font.Size = 11
    ws.Cells(22, 3).Value = "生成 I-MR 管制圖 + 製程能力分析"
    ws.Cells(22, 3).Font.Size = 9
    ws.Cells(22, 3).Font.Color = RGB(100, 100, 100)
    
    Call 創建按鈕(ws, 23, 3, "執行批號分析", "ExecuteBatchAnalysis")
    
    ' 模穴分析 (E欄)
    ws.Cells(21, 5).Value = "模穴分析"
    ws.Cells(21, 5).Font.Bold = True
    ws.Cells(21, 5).Font.Size = 11
    ws.Cells(22, 5).Value = "生成模穴比較圖 + 能力評估"
    ws.Cells(22, 5).Font.Size = 9
    ws.Cells(22, 5).Font.Color = RGB(100, 100, 100)
    
    Call 創建按鈕(ws, 23, 5, "執行模穴分析", "ExecuteCavityAnalysis")
    
    ' 群組分析 (G欄)
    ws.Cells(21, 7).Value = "群組分析"
    ws.Cells(21, 7).Font.Bold = True
    ws.Cells(21, 7).Font.Size = 11
    ws.Cells(22, 7).Value = "生成 Min-Max-Avg 管制圖"
    ws.Cells(22, 7).Font.Size = 9
    ws.Cells(22, 7).Font.Color = RGB(100, 100, 100)
    
    Call 創建按鈕(ws, 23, 7, "執行群組分析", "ExecuteGroupAnalysis")
    
    ' 功能說明區域
    ws.Cells(27, 2).Value = "功能說明"
    ws.Cells(27, 2).Font.Size = 12
    ws.Cells(27, 2).Font.Bold = True
    ws.Cells(27, 2).Font.Color = RGB(0, 112, 192)
    
    ws.Cells(28, 3).Value = "批號分析: Individual-X 管制圖, MR 管制圖, 製程能力分析, 異常檢測"
    ws.Cells(29, 3).Value = "模穴分析: 模穴統計摘要, 平均值/Cpk/標準差比較圖, 能力評估"
    ws.Cells(30, 3).Value = "群組分析: Min-Max-Avg 管制圖, 模穴間變異監控"
    
    ws.Cells(28, 3).Font.Size = 9
    ws.Cells(29, 3).Font.Size = 9
    
    ' 版本資訊
    ws.Cells(32, 2).Value = "版本: v1.0 | 更新日期: 2024-11-20"
    ws.Cells(32, 2).Font.Size = 8
    ws.Cells(32, 2).Font.Color = RGB(150, 150, 150)
    ws.Cells(32, 2).Font.Italic = True
    
    ' 設定列寬 - 調整佈局以獲得更好的間距
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 15
    
    ' 第一組：批號分析 (C欄)
    ws.Columns("C").ColumnWidth = 25
    
    ' 間隔 (D欄)
    ws.Columns("D").ColumnWidth = 5
    
    ' 第二組：模穴分析 (E欄)
    ws.Columns("E").ColumnWidth = 25
    
    ' 間隔 (F欄)
    ws.Columns("F").ColumnWidth = 5
    
    ' 第三組：群組分析 (G欄)
    ws.Columns("G").ColumnWidth = 25
    
    ' 調整背景色範圍
    Dim headerRange As String
    headerRange = "B7:G7,B13:G13,B19:G19,B27:G27"
    ws.Range(headerRange).Interior.Color = RGB(217, 225, 242)
    
    ' 凍結窗格
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

' ============================================================================
' 重置系統（清除記憶體和選擇）
' ============================================================================
Public Sub ResetSystem()
    Dim ws As Worksheet
    
    ' 檢查工作表是否存在
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("整合分析")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    ' 確認是否重置
    If MsgBox("確定要重置所有選擇並清除記憶體嗎？", vbYesNo + vbQuestion, "確認重置") = vbNo Then
        Exit Sub
    End If
    
    ' 清除全局變數
    If Not g_SelectedWorkbook Is Nothing Then
        ' 如果檔案是我們打開的，可能需要關閉？
        ' 這裡我們只清除引用，不關閉檔案，以免影響用戶其他操作
        ' 或者可以選擇關閉： g_SelectedWorkbook.Close False
        Set g_SelectedWorkbook = Nothing
    End If
    
    g_SelectedItem = ""
    g_SelectedFilePath = ""
    
    ' 重置 UI 顯示
    ws.Range("SelectedFile").Value = "未選擇"
    ws.Range("SelectedFile").Font.Color = RGB(150, 150, 150)
    ws.Range("SelectedFile").Font.Bold = False
    
    ws.Range("SelectedItem").Value = "未選擇"
    ws.Range("SelectedItem").Font.Color = RGB(150, 150, 150)
    ws.Range("SelectedItem").Font.Bold = False
    
    MsgBox "系統已重置！", vbInformation
End Sub

' ============================================================================
' 創建按鈕
' ============================================================================
Private Sub 創建按鈕(ws As Worksheet, row As Long, col As Long, _
                    buttonText As String, macroName As String)
    
    Dim btn As Button
    Dim btnLeft As Double
    Dim btnTop As Double
    Dim btnWidth As Double
    Dim btnHeight As Double
    
    ' 計算按鈕位置（按鈕寬度設為與欄寬一致）
    btnLeft = ws.Cells(row, col).Left
    btnTop = ws.Cells(row, col).Top
    btnWidth = ws.Cells(row, col).Width
    btnHeight = 35
    
    ' 刪除現有按鈕（如果存在）
    On Error Resume Next
    ws.Buttons(buttonText).Delete
    On Error GoTo 0
    
    ' 創建按鈕
    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Name = buttonText
    btn.Caption = buttonText
    btn.OnAction = "UI_Integrated." & macroName
    
    ' 設定按鈕樣式
    btn.Font.Size = 11
    btn.Font.Bold = True
End Sub

' ============================================================================
' 步驟 1：選擇檔案
' ============================================================================
Public Sub SelectFile()
    Dim ws As Worksheet
    
    ' 檢查工作表是否存在
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("整合分析")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "請先執行「創建整合分析介面」！", vbExclamation
        Exit Sub
    End If
    
    ' 選擇檔案
    Set g_SelectedWorkbook = DataInput.SelectDataFile()
    
    If Not g_SelectedWorkbook Is Nothing Then
        g_SelectedFilePath = g_SelectedWorkbook.FullName
        
        ' 更新顯示
        ws.Range("SelectedFile").Value = g_SelectedWorkbook.Name
        ws.Range("SelectedFile").Font.Color = RGB(0, 112, 192)
        ws.Range("SelectedFile").Font.Bold = True
        
        ' 清除檢驗項目選擇
        ws.Range("SelectedItem").Value = "未選擇"
        ws.Range("SelectedItem").Font.Color = RGB(150, 150, 150)
        ws.Range("SelectedItem").Font.Bold = False
        g_SelectedItem = ""
        
        MsgBox "檔案已選擇！" & vbCrLf & "請繼續選擇檢驗項目", vbInformation
    End If
End Sub

' ============================================================================
' 步驟 2：選擇檢驗項目
' ============================================================================
Public Sub SelectItem()
    Dim ws As Worksheet
    
    ' 檢查工作表是否存在
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("整合分析")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "請先執行「創建整合分析介面」！", vbExclamation
        Exit Sub
    End If
    
    ' 檢查是否已選擇檔案
    If g_SelectedWorkbook Is Nothing Then
        MsgBox "請先選擇數據檔案！", vbExclamation
        Exit Sub
    End If
    
    ' 選擇檢驗項目
    g_SelectedItem = DataInput.SelectInspectionItem(g_SelectedWorkbook)
    
    If g_SelectedItem <> "" Then
        ' 更新顯示
        ws.Range("SelectedItem").Value = g_SelectedItem
        ws.Range("SelectedItem").Font.Color = RGB(0, 112, 192)
        ws.Range("SelectedItem").Font.Bold = True
        
        ' MsgBox "檢驗項目已選擇！" & vbCrLf & "請選擇要執行的分析類型", vbInformation
    End If
End Sub

' ============================================================================
' 步驟 3：執行批號分析
' ============================================================================
Public Sub ExecuteBatchAnalysis()
    ' 檢查是否已完成前置步驟
    If g_SelectedWorkbook Is Nothing Then
        MsgBox "請先選擇數據檔案！", vbExclamation
        Exit Sub
    End If
    
    If g_SelectedItem = "" Then
        MsgBox "請先選擇檢驗項目！", vbExclamation
        Exit Sub
    End If
    
    ' 執行批號分析（使用已選擇的檔案和項目）
    Call 執行批號分析_已選擇(g_SelectedWorkbook, g_SelectedItem)
End Sub

' ============================================================================
' 步驟 3：執行模穴分析
' ============================================================================
Public Sub ExecuteCavityAnalysis()
    ' 檢查是否已完成前置步驟
    If g_SelectedWorkbook Is Nothing Then
        MsgBox "請先選擇數據檔案！", vbExclamation
        Exit Sub
    End If
    
    If g_SelectedItem = "" Then
        MsgBox "請先選擇檢驗項目！", vbExclamation
        Exit Sub
    End If
    
    ' 執行模穴分析（使用已選擇的檔案和項目）
    Call 執行模穴分析_已選擇(g_SelectedWorkbook, g_SelectedItem)
End Sub

' ============================================================================
' 步驟 3：執行群組分析
' ============================================================================
Public Sub ExecuteGroupAnalysis()
    ' 檢查是否已完成前置步驟
    If g_SelectedWorkbook Is Nothing Then
        MsgBox "請先選擇數據檔案！", vbExclamation
        Exit Sub
    End If
    
    If g_SelectedItem = "" Then
        MsgBox "請先選擇檢驗項目！", vbExclamation
        Exit Sub
    End If
    
    ' 執行群組分析（使用已選擇的檔案和項目）
    Call 執行群組分析_已選擇(g_SelectedWorkbook, g_SelectedItem)
End Sub

' ============================================================================
' 執行批號分析（使用已選擇的檔案）
' ============================================================================
Private Sub 執行批號分析_已選擇(wbSource As Workbook, itemName As String)
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim productCode As String
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrorHandler
    
    ' 設定來源工作表
    Set wsSource = wbSource.Worksheets(itemName)
    
    ' 提取產品品號
    productCode = DataInput.ExtractProductCode(wbSource.FullName)
    
    ' 創建分析工作表
    Set wsTarget = ThisWorkbook.Worksheets.Add
    Dim wsName As String
    wsName = DataInput.GenerateWorksheetName(productCode, itemName, "批號分析")
    DataInput.SetWorksheetName wsTarget, wsName
    
    ' 清空新工作表的內容以避免與舊數據重疊
    wsTarget.Cells.Clear
    
    ' 添加標題資訊
    wsTarget.Cells(1, 1).Value = "來源檔案:"
    wsTarget.Cells(1, 2).Value = wbSource.Name
    wsTarget.Cells(1, 3).Value = "檢驗項目:"
    wsTarget.Cells(1, 4).Value = itemName
    wsTarget.Cells(1, 5).Value = "分析時間:"
    wsTarget.Cells(1, 6).Value = Now
    
    With wsTarget.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' 生成多工作表管制圖
    On Error Resume Next
    Call BatchAnalysis.GenerateMultipleXBarRCharts(wsSource, itemName, productCode)
    On Error GoTo ErrorHandler
    
    ' 刪除臨時創建的工作表（因為新函數會自己創建工作表）
    Application.DisplayAlerts = False
    wsTarget.Delete
    Application.DisplayAlerts = True
    
    MsgBox "批號分析完成！" & vbCrLf & "每個工作表最多包含25組數據", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "執行批號分析時發生錯誤：" & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================================
' 執行模穴分析（使用已選擇的檔案）
' ============================================================================
Private Sub 執行模穴分析_已選擇(wbSource As Workbook, itemName As String)
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim productCode As String
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrorHandler
    
    ' 設定來源工作表
    Set wsSource = wbSource.Worksheets(itemName)
    
    ' 提取產品品號
    productCode = DataInput.ExtractProductCode(wbSource.FullName)
    
    ' 創建分析工作表
    Set wsTarget = ThisWorkbook.Worksheets.Add
    Dim wsName As String
    wsName = DataInput.GenerateWorksheetName(productCode, itemName, "模穴分析")
    DataInput.SetWorksheetName wsTarget, wsName
    
    ' 清空新工作表的內容以避免與舊數據重疊
    wsTarget.Cells.Clear
    
    ' 添加標題資訊
    wsTarget.Cells(1, 1).Value = "來源檔案:"
    wsTarget.Cells(1, 2).Value = wbSource.Name
    wsTarget.Cells(1, 3).Value = "檢驗項目:"
    wsTarget.Cells(1, 4).Value = itemName
    wsTarget.Cells(1, 5).Value = "分析時間:"
    wsTarget.Cells(1, 6).Value = Now
    
    ' 合并标题行前先清除内容以避免警告
    wsTarget.Range("A1:F1").ClearContents
    
    With wsTarget.Range("A1:F1")
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' 讀取規格界限
    Dim target As Double, usl As Double, lsl As Double
    ' 規格數據在第2行（第1行標題，第2行規格+數據）
    ' 使用 Text 屬性避免 Excel 將 "J" 等字符誤判為日期格式
    On Error Resume Next
    target = CDbl(wsSource.Cells(2, 2).Text)
    If Err.Number <> 0 Then target = wsSource.Cells(2, 2).Value
    Err.Clear
    
    usl = CDbl(wsSource.Cells(2, 3).Text)
    If Err.Number <> 0 Then usl = wsSource.Cells(2, 3).Value
    Err.Clear
    
    lsl = CDbl(wsSource.Cells(2, 4).Text)
    If Err.Number <> 0 Then lsl = wsSource.Cells(2, 4).Value
    On Error GoTo ErrorHandler
    
    ' 分析所有模穴
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.UsedRange.Columns.Count + wsSource.UsedRange.Column - 1
    
    Call CavityAnalysis.分析所有模穴(wsTarget, wsSource, lastRow, lastCol, target, usl, lsl)
    
    ' 激活目標工作表
    wsTarget.Activate
    
    MsgBox "模穴分析完成！", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "執行模穴分析時發生錯誤：" & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================================
' 執行群組分析（使用已選擇的檔案）
' ============================================================================
Private Sub 執行群組分析_已選擇(wbSource As Workbook, itemName As String)
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim productCode As String
    Dim lastRow As Long, lastCol As Long
    Dim cavityCol As Long
    Dim i As Long
    Dim values As Collection
    Dim value As Variant
    Dim rowSum As Double, rowMean As Double
    Dim rowMax As Double, rowMin As Double
    Dim rowRange As Double
    Dim target As Double, usl As Double, lsl As Double
    Dim batchLabel As String
    Dim targetRow As Long
    Dim cavityCols As Collection
    
    On Error GoTo ErrorHandler
    
    ' 設定來源工作表
    Set wsSource = wbSource.Worksheets(itemName)
    
    ' 提取產品品號
    productCode = DataInput.ExtractProductCode(wbSource.FullName)
    
    ' 創建分析工作表
    Set wsTarget = ThisWorkbook.Worksheets.Add
    Dim wsName As String
    wsName = DataInput.GenerateWorksheetName(productCode, itemName, "群組分析")
    DataInput.SetWorksheetName wsTarget, wsName
    
    ' 清空新工作表的內容以避免與舊數據重疊
    wsTarget.Cells.Clear
    
    ' 添加標題資訊
    wsTarget.Cells(1, 1).Value = "來源檔案:"
    wsTarget.Cells(1, 2).Value = wbSource.Name
    wsTarget.Cells(1, 3).Value = "檢驗項目:"
    wsTarget.Cells(1, 4).Value = itemName
    wsTarget.Cells(1, 5).Value = "分析時間:"
    wsTarget.Cells(1, 6).Value = Now
    
    ' 合并标题行前先清除内容以避免警告
    wsTarget.Range("A1:F1").ClearContents
    
    With wsTarget.Range("A1:F1")
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With
    
    ' 讀取規格
    ' 規格數據在第2行（第1行標題，第2行規格+數據）
    ' 使用 Text 屬性避免 Excel 將 "J" 等字符誤判為日期格式
    On Error Resume Next
    target = CDbl(wsSource.Cells(2, 2).Text)
    If Err.Number <> 0 Then target = wsSource.Cells(2, 2).Value
    Err.Clear
    
    usl = CDbl(wsSource.Cells(2, 3).Text)
    If Err.Number <> 0 Then usl = wsSource.Cells(2, 3).Value
    Err.Clear
    
    lsl = CDbl(wsSource.Cells(2, 4).Text)
    If Err.Number <> 0 Then lsl = wsSource.Cells(2, 4).Value
    On Error GoTo ErrorHandler
    
    ' 設置數據表標題
    wsTarget.Cells(3, 1).Value = "批號"
    wsTarget.Cells(3, 2).Value = "平均值 (Avg)"
    wsTarget.Cells(3, 3).Value = "最大值 (Max)"
    wsTarget.Cells(3, 4).Value = "最小值 (Min)"
    wsTarget.Cells(3, 5).Value = "全距 (Range)"
    wsTarget.Cells(3, 6).Value = "樣本數 (n)"
    wsTarget.Cells(3, 7).Value = "USL"
    wsTarget.Cells(3, 8).Value = "Target"
    wsTarget.Cells(3, 9).Value = "LSL"
    
    ' 清除内容以避免合并单元格警告
    wsTarget.Range("A3:I3").ClearContents
    
    With wsTarget.Range("A3:I3")
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 找到數據範圍
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.UsedRange.Columns.Count + wsSource.UsedRange.Column - 1
    
    ' 識別模穴列
    Set cavityCols = New Collection
    For cavityCol = 5 To lastCol
        If InStr(CStr(wsSource.Cells(1, cavityCol).Value), "穴") > 0 Then
            cavityCols.Add cavityCol
        End If
    Next cavityCol
    
    If cavityCols.Count = 0 Then
        MsgBox "未找到模穴數據列（標題需包含'穴'字）", vbExclamation
        Exit Sub
    End If
    
    ' 分析每一批次
    targetRow = 4
    For i = 3 To lastRow
        batchLabel = CStr(wsSource.Cells(i, 1).Value)
        
        ' 收集該批次所有模穴的數據
        Set values = New Collection
        For Each value In cavityCols
            If IsNumeric(wsSource.Cells(i, value).Value) And wsSource.Cells(i, value).Value <> "" Then
                values.Add CDbl(wsSource.Cells(i, value).Value)
            End If
        Next value
        
        If values.Count > 0 Then
            ' 計算統計量
            rowSum = 0
            rowMax = values(1)
            rowMin = values(1)
            
            For Each value In values
                rowSum = rowSum + value
                If value > rowMax Then rowMax = value
                If value < rowMin Then rowMin = value
            Next value
            
            rowMean = rowSum / values.Count
            rowRange = rowMax - rowMin
            
            ' 寫入數據
            wsTarget.Cells(targetRow, 1).NumberFormat = "@"
            wsTarget.Cells(targetRow, 1).Value = batchLabel
            wsTarget.Cells(targetRow, 2).Value = Round(rowMean, 4)
            wsTarget.Cells(targetRow, 3).Value = Round(rowMax, 4)
            wsTarget.Cells(targetRow, 4).Value = Round(rowMin, 4)
            wsTarget.Cells(targetRow, 5).Value = Round(rowRange, 4)
            wsTarget.Cells(targetRow, 6).Value = values.Count
            wsTarget.Cells(targetRow, 7).Value = usl
            wsTarget.Cells(targetRow, 8).Value = target
            wsTarget.Cells(targetRow, 9).Value = lsl
            
            targetRow = targetRow + 1
        End If
    Next i
    
    ' 格式化
    wsTarget.Columns.AutoFit
    
    ' 生成群組管制圖
    If targetRow > 4 Then
        ' 調用 GroupAnalysis 模組中的生成圖表函數
        Call GroupAnalysis.生成群組管制圖(wsTarget, targetRow - 1, itemName)
    End If
    
    ' 激活目標工作表
    wsTarget.Activate
    
    MsgBox "群組分析完成！", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "執行群組分析時發生錯誤：" & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================================
' 顯示關於資訊
' ============================================================================
Public Sub 顯示關於()
    MsgBox "QIP SPC 分析系統 - 整合版" & vbCrLf & vbCrLf & _
           "版本: v1.0" & vbCrLf & _
           "更新日期: 2024-11-20" & vbCrLf & vbCrLf & _
           "功能:" & vbCrLf & _
           "- 批號分析 (I-MR 管制圖 + 製程能力分析)" & vbCrLf & _
           "- 模穴分析 (模穴比較 + 能力評估)" & vbCrLf & vbCrLf & _
           "整合第二檔案和第三檔案的完整功能", _
           vbInformation, "關於"
End Sub
