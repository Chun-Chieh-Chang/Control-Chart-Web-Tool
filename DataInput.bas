Option Explicit

' ============================================================================
' 統一數據輸入模組
' 整合第二檔案和第三檔案的檔案選擇和檢驗項目選擇功能
' ============================================================================

' ============================================================================
' 選擇數據檔案（統一的檔案選擇邏輯）
' ============================================================================
Public Function SelectDataFile() As Workbook
    Dim fileDialog As FileDialog
    Dim filePath As String
    Dim defaultPath As String
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler
    
    ' 創建檔案選擇對話框
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "選擇第一檔案（QIP_DataExtract）輸出的數據檔案"
        .Filters.Clear
        .Filters.Add "Excel 檔案", "*.xlsx; *.xlsm; *.xls"
        .AllowMultiSelect = False
        
        ' 設定初始路徑為 ResultData 資料夾（如果存在）
        On Error Resume Next
        defaultPath = ThisWorkbook.Path & "\ResultData\"
        If Dir(defaultPath, vbDirectory) <> "" Then
            .InitialFileName = defaultPath
        Else
            .InitialFileName = ThisWorkbook.Path
        End If
        On Error GoTo ErrorHandler
        
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Set SelectDataFile = Nothing
            Exit Function
        End If
    End With
    
    ' 打開檔案
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set SelectDataFile = wb
    Exit Function
    
ErrorHandler:
    MsgBox "選擇檔案時發生錯誤：" & Err.Description, vbCritical
    Set SelectDataFile = Nothing
End Function

' ============================================================================
' 選擇檢驗項目（使用自定義對話框）
' ============================================================================
Public Function SelectInspectionItem(wbSource As Workbook) As String
    Dim ws As Worksheet
    Dim items As Collection
    Dim itemArray() As String
    Dim i As Long
    Dim selectedItem As Variant
    Dim targetVal As Variant
    Dim displayStr As String
    
    ' 收集所有工作表名稱（排除特殊工作表）
    Set items = New Collection
    
    For Each ws In wbSource.Worksheets
        ' 排除可能的摘要或系統工作表
        If ws.Name <> "摘要" And ws.Name <> "Summary" And _
           ws.Name <> "統計" And ws.Name <> "說明" And _
           Not (ws.Name Like "*分析*") And _
           Not (ws.Name Like "*配置*") Then
            
            ' 讀取 Target 數值 (假設在 B2，第1行標題，第2行規格+數據)
            ' 使用 Text 屬性避免 Excel 將 "J" 等字符誤判為日期格式
            On Error Resume Next
            targetVal = ws.Cells(2, 2).Text
            If targetVal = "" Or Err.Number <> 0 Then
                targetVal = ws.Cells(2, 2).Value
            End If
            On Error GoTo 0
            If IsEmpty(targetVal) Or targetVal = "" Then targetVal = "N/A"
            
            ' 格式化顯示字串：項目名稱 [Target: 數值]
            displayStr = ws.Name & "   [Target: " & targetVal & "]"
            items.Add displayStr
        End If
    Next ws
    
    If items.Count = 0 Then
        MsgBox "未找到檢驗項目工作表！", vbExclamation
        SelectInspectionItem = ""
        Exit Function
    End If
    
    ' 轉換為陣列
    ReDim itemArray(1 To items.Count)
    For i = 1 To items.Count
        itemArray(i) = items(i)
    Next i
    
    ' 顯示選擇對話框 (使用單選模式，因為目前流程只支援單一項目)
    selectedItem = ShowSelectionDialog("選擇檢驗項目", "請選擇要分析的檢驗項目：", itemArray, False)
    
    If VarType(selectedItem) = vbString And selectedItem <> "" Then
        ' 解析出原始工作表名稱 (移除 Target 資訊)
        Dim rawStr As String
        rawStr = selectedItem
        
        If InStr(rawStr, "   [Target:") > 0 Then
            SelectInspectionItem = Left(rawStr, InStr(rawStr, "   [Target:") - 1)
        Else
            SelectInspectionItem = rawStr
        End If
    Else
        SelectInspectionItem = ""
    End If
End Function

' ============================================================================
' 選擇模穴（多選）
' ============================================================================
Public Function SelectCavities(wsSource As Worksheet, Optional headerRow As Long = 1) As Variant
    Dim cavityCols As Collection
    Dim cavityNames() As String
    Dim lastCol As Long, c As Long
    Dim headerVal As String
    Dim i As Long
    Dim selectedVariant As Variant
    
    ' 1. 識別所有模穴
    Set cavityCols = New Collection
    lastCol = wsSource.UsedRange.Columns.Count + wsSource.UsedRange.Column - 1
    
    For c = 1 To lastCol
        headerVal = CStr(wsSource.Cells(headerRow, c).Value)
        ' 簡單判斷：包含 "穴" 字
        If InStr(headerVal, "穴") > 0 Then
            cavityCols.Add headerVal
        End If
    Next c
    
    If cavityCols.Count = 0 Then
        MsgBox "在工作表 " & wsSource.Name & " 中找不到模穴欄位（標題需包含'穴'）" & vbCrLf & _
               "檢查行號：" & headerRow, vbExclamation
        SelectCavities = False
        Exit Function
    End If
    
    ' 2. 轉換為陣列供對話框使用
    ReDim cavityNames(1 To cavityCols.Count)
    For i = 1 To cavityCols.Count
        cavityNames(i) = cavityCols(i)
    Next i
    
    ' 3. 顯示多選對話框
    selectedVariant = ShowSelectionDialog("選擇模穴", "請勾選要分析的模穴（可多選）：", cavityNames, True)
    
    SelectCavities = selectedVariant
End Function

Public Function ShowSelectionDialog(title As String, prompt As String, items As Variant, Optional allowMultiSelect As Boolean = False) As Variant
    Dim wsDialog As Worksheet
    Dim activeWs As Worksheet
    Dim btnConfirm As Button, btnCancel As Button
    Dim lstBox As ListBox
    Dim chkBox As CheckBox
    Dim i As Long, col As Long, row As Long
    Dim dialogWidth As Double, dialogHeight As Double
    Dim leftPos As Double, topPos As Double
    Dim itemsPerCol As Long
    Dim chkAll As CheckBox
    
    itemsPerCol = 8 ' 每列顯示 8 個項目
    
    On Error GoTo ErrorHandler
    
    Debug.Print "=== ShowSelectionDialog 開始 ==="
    Debug.Print "allowMultiSelect: " & allowMultiSelect
    
    Application.ScreenUpdating = False
    Set activeWs = ActiveSheet
    
    Debug.Print "創建臨時工作表..."
    ' 創建臨時工作表
    Set wsDialog = ThisWorkbook.Worksheets.Add
    wsDialog.Name = "Dialog_" & Format(Now, "hhmmss")
    Debug.Print "工作表已創建: " & wsDialog.Name
    
    ' 設定樣式
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
    End With
    wsDialog.Cells.Interior.Color = RGB(240, 240, 240)
    
    Debug.Print "添加標題..."
    ' 標題
    With wsDialog.Shapes.AddLabel(msoTextOrientationHorizontal, 20, 10, 400, 30)
        .TextFrame.Characters.Text = title
        .TextFrame.Characters.Font.Size = 14
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(0, 51, 102)
    End With
    
    Debug.Print "添加提示..."
    ' 提示
    With wsDialog.Shapes.AddLabel(msoTextOrientationHorizontal, 20, 45, 400, 20)
        .TextFrame.Characters.Text = prompt
        .TextFrame.Characters.Font.Size = 10
    End With
    
    If allowMultiSelect Then
        Debug.Print "進入多選模式..."
        ' === 多選模式：使用 ListBox 多選 ===
        
        Set lstBox = wsDialog.ListBoxes.Add(20, 70, 400, 250)
        With lstBox
            .Name = "lstItems"
            .MultiSelect = xlExtended  ' 允許多選（Ctrl+點擊 或 Shift+點擊）
            .RemoveAllItems
            
            If IsArray(items) Then
                For i = LBound(items) To UBound(items)
                    .AddItem items(i)
                Next i
            End If
        End With
        
        ' 添加說明文字
        With wsDialog.Shapes.AddLabel(msoTextOrientationHorizontal, 20, 330, 400, 40)
            .TextFrame.Characters.Text = "提示：" & vbCrLf & _
                "• 按住 Ctrl 鍵點擊可選擇多個項目" & vbCrLf & _
                "• 按住 Shift 鍵點擊可選擇連續範圍"
            .TextFrame.Characters.Font.Size = 9
            .TextFrame.Characters.Font.Color = RGB(100, 100, 100)
        End With
        
        topPos = 380
        
    Else
        Debug.Print "進入單選模式..."
        ' === 單選模式：維持使用 ListBox ===
        Set lstBox = wsDialog.ListBoxes.Add(20, 70, 250, 200)
        With lstBox
            .Name = "lstItems"
            .MultiSelect = xlNone
            .RemoveAllItems
            If IsArray(items) Then
                For i = LBound(items) To UBound(items)
                    .AddItem items(i)
                Next i
            End If
        End With
        topPos = 290
    End If
    
    Debug.Print "創建確認和取消按鈕..."
    ' 確認按鈕
    Set btnConfirm = wsDialog.Buttons.Add(150, topPos, 80, 30)
    With btnConfirm
        .Caption = "確定"
        .OnAction = "'DataInput.DialogConfirm'"
        .Font.Bold = True
    End With
    
    ' 取消按鈕
    Set btnCancel = wsDialog.Buttons.Add(50, topPos, 80, 30)
    With btnCancel
        .Caption = "取消"
        .OnAction = "'DataInput.DialogCancel'"
    End With
    
    Debug.Print "設置狀態標記..."
    ' 狀態標記
    wsDialog.Range("A1").Value = "WAIT"
    wsDialog.Range("A2").Value = ""
    wsDialog.Range("A3").Value = IIf(allowMultiSelect, "MULTI", "SINGLE") ' 記錄模式
    wsDialog.Columns("A").Hidden = True
    
    Application.ScreenUpdating = True
    
    ' 激活對話框工作表，確保用戶可以看到
    wsDialog.Activate
    Application.Goto wsDialog.Range("B10"), True
    
    Debug.Print "等待用戶操作..."
    ' 等待循環
    Do While wsDialog.Range("A1").Value = "WAIT"
        DoEvents
        If wsDialog Is Nothing Then Exit Do
        On Error Resume Next
        If wsDialog.Name = "" Then Exit Do
        On Error GoTo ErrorHandler
    Loop
    
    Debug.Print "處理結果..."
    ' 處理結果
    If wsDialog.Range("A1").Value = "OK" Then
        Dim rawResult As String
        rawResult = wsDialog.Range("A2").Value
        
        If allowMultiSelect Then
            If InStr(rawResult, "|") > 0 Then
                ShowSelectionDialog = Split(rawResult, "|")
            ElseIf rawResult <> "" Then
                Dim singleResult(0 To 0) As String
                singleResult(0) = rawResult
                ShowSelectionDialog = singleResult
            Else
                ShowSelectionDialog = False
            End If
        Else
            ShowSelectionDialog = rawResult
        End If
    Else
        ShowSelectionDialog = False
    End If
    
    Debug.Print "清理對話框..."
    ' 清理
    Application.DisplayAlerts = False
    wsDialog.Delete
    Application.DisplayAlerts = True
    activeWs.Activate
    
    Debug.Print "=== ShowSelectionDialog 完成 ==="
    Exit Function
    
ErrorHandler:
    Debug.Print "!!! 發生錯誤 !!!"
    Debug.Print "錯誤編號: " & Err.Number
    Debug.Print "錯誤描述: " & Err.Description
    Debug.Print "錯誤來源: " & Err.Source
    
    If Not wsDialog Is Nothing Then
        Application.DisplayAlerts = False
        wsDialog.Delete
        Application.DisplayAlerts = True
    End If
    If Not activeWs Is Nothing Then activeWs.Activate
    MsgBox "對話框錯誤：" & Err.Description & vbCrLf & "錯誤號：" & Err.Number & vbCrLf & vbCrLf & "請查看即時視窗以獲取詳細資訊", vbCritical
    ShowSelectionDialog = False
End Function

' ============================================================================
' 全選/取消全選
' ============================================================================
Public Sub ToggleSelectAll()
    Dim ws As Worksheet
    Dim obj As Object
    Dim isSelected As Boolean
    
    Set ws = ActiveSheet
    
    ' 獲取全選按鈕的狀態
    For Each obj In ws.OLEObjects
        If obj.Name = "chkSelectAll" Then
            isSelected = obj.Object.Value
            Exit For
        End If
    Next obj
    
    ' 設置所有項目的狀態
    For Each obj In ws.OLEObjects
        If obj.Name Like "chkItem_*" Then
            obj.Object.Value = isSelected
        End If
    Next obj
End Sub

' ============================================================================
' 對話框確認按鈕回調
' ============================================================================
Public Sub DialogConfirm()
    Dim ws As Worksheet
    Dim lst As ListBox
    Dim resultArray() As String
    Dim count As Long
    Dim mode As String
    Dim i As Long
    Dim idx As Long
    
    Set ws = ActiveSheet
    mode = ws.Range("A3").Value
    
    Debug.Print "DialogConfirm 開始，模式：" & mode
    
    Set lst = ws.ListBoxes("lstItems")
    
    If mode = "MULTI" Then
        ' === 處理 ListBox 多選 ===
        count = 0
        For i = 1 To lst.ListCount
            If lst.Selected(i) Then
                count = count + 1
            End If
        Next i
        
        Debug.Print "總共選擇：" & count & " 個項目"
        
        If count = 0 Then
            MsgBox "請至少選擇一個項目！", vbExclamation
            Exit Sub
        End If
        
        ReDim resultArray(1 To count)
        idx = 0
        
        For i = 1 To lst.ListCount
            If lst.Selected(i) Then
                idx = idx + 1
                resultArray(idx) = lst.List(i)
                Debug.Print "選擇項目 " & idx & ": " & lst.List(i)
            End If
        Next i
        
        ws.Range("A2").Value = Join(resultArray, "|")
        Debug.Print "結果：" & ws.Range("A2").Value
        
    Else
        ' === 處理 ListBox 單選 ===
        If lst.ListIndex = 0 Then
            MsgBox "請選擇一個項目！", vbExclamation
            Exit Sub
        End If
        ws.Range("A2").Value = lst.List(lst.ListIndex)
    End If
    
    ws.Range("A1").Value = "OK"
    Debug.Print "DialogConfirm 完成"
End Sub

' ============================================================================
' 對話框取消按鈕回調
' ============================================================================
Public Sub DialogCancel()
    ActiveSheet.Range("A1").Value = "Cancel"
End Sub

' ============================================================================
' 從檔案名稱提取產品品號
' ============================================================================
Public Function ExtractProductCode(filePath As String) As String
    Dim fileName As String
    Dim productCode As String
    
    fileName = Dir(filePath)
    
    ' 移除副檔名
    If InStr(fileName, ".") > 0 Then
        productCode = Left(fileName, InStr(fileName, ".") - 1)
    Else
        productCode = fileName
    End If
    
    ExtractProductCode = productCode
End Function

' ============================================================================
' 生成工作表名稱（處理長度限制）
' ============================================================================
Public Function GenerateWorksheetName(productCode As String, itemName As String, _
                                     analysisType As String) As String
    Dim wsName As String
    Dim maxItemLen As Integer
    
    ' 組合工作表名稱：產品品號_檢驗項目_分析類型
    wsName = productCode & "_" & itemName & "_" & analysisType
    
    ' 如果名稱超過 31 字元，進行截斷
    If Len(wsName) > 31 Then
        maxItemLen = 31 - Len(productCode) - Len(analysisType) - 2
        If maxItemLen > 0 Then
            wsName = productCode & "_" & Left(itemName, maxItemLen) & "_" & analysisType
        Else
            wsName = Left(productCode, 31 - Len(analysisType) - 1) & "_" & analysisType
        End If
    End If
    
    GenerateWorksheetName = wsName
End Function

' ============================================================================
' 設定工作表名稱（處理重複名稱）
' ============================================================================
Public Sub SetWorksheetName(ws As Worksheet, baseName As String)
    Dim counter As Integer
    Dim newName As String
    
    On Error Resume Next
    
    ' 首先尝试使用基础名称
    ws.Name = baseName
    
    ' 如果命名失败（可能名称已存在），添加序号
    If Err.Number <> 0 Then
        Err.Clear
        counter = 1
        newName = baseName & "_" & counter
        ws.Name = newName
        
        ' 继续尝试直到成功
        Do While Err.Number <> 0
            Err.Clear
            counter = counter + 1
            newName = baseName & "_" & counter
            ws.Name = newName
        Loop
    End If
    
    On Error GoTo 0
End Sub
