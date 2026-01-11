Option Explicit

' === 數據結構定義 ===
Public Type CavityGroupConfig
    pageOffset As Long
    dataRange As String
    cavityIdRange As String
End Type

Public Type UserInputParams
    folderPath As String
    fileKeyword As String
    cavitiesForUI As Long
End Type


' === 全局變量 ===
Private g_SampleWorkbook As Workbook  ' 樣本檔案工作簿
Private gPageOffset(1 To 6) As Long
Private gDataRange(1 To 6) As String
Private gCavityIdRange(1 To 6) As String

' ===================================================================
' 主執行函數
' ===================================================================
Sub ProcessInspectionReports_SheetInput()
    Dim params As UserInputParams
    Dim configSheet As Worksheet
    
    On Error Resume Next
    Set configSheet = ThisWorkbook.Worksheets("參數配置")
    On Error GoTo 0
    
    If configSheet Is Nothing Then
        Set configSheet = CreateConfigSheet()
        If configSheet Is Nothing Then Exit Sub ' 如果創建失敗則退出
        
        MsgBox "已成功創建「參數配置」工作表。" & vbCrLf & vbCrLf & _
               "使用步驟：" & vbCrLf & _
               "1. 點擊 [瀏覽] 按鈕選擇資料夾" & vbCrLf & _
               "2. 點擊 [開啟樣本] 按鈕開啟樣本檔案" & vbCrLf & _
               "3. 點擊 [選擇] 按鈕選擇範圍" & vbCrLf & _
               "4. 點擊 [保存配置] 按鈕保存配置" & vbCrLf & _
               "5. 點擊 [開始處理] 按鈕開始處理", vbInformation, "配置工作表已創建"
        configSheet.Activate
        Exit Sub
    End If
    
    If Not ReadParametersFromSheet(configSheet, params) Then Exit Sub
    If Not ValidateCavityGroupInput(params) Then Exit Sub
    
    Dim confirmMsg As String
    confirmMsg = "請確認以下參數：" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "來源檔案：" & params.folderPath & vbCrLf
    confirmMsg = confirmMsg & "產品品號：" & params.fileKeyword & vbCrLf
    confirmMsg = confirmMsg & "模穴數：" & params.cavitiesForUI & " 穴" & vbCrLf & vbCrLf
    confirmMsg = confirmMsg & "確認無誤請按「是」繼續。"
    
    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "參數確認") <> vbYes Then
        MsgBox "操作已取消。", vbInformation
        Exit Sub
    End If
    
    Call ExecuteProcessingCore(params)
End Sub

Private Function ShowPagedConfigPicker(configWs As Worksheet) As Long
    Dim lastRow As Long, total As Long, pageSize As Long, page As Long, maxPage As Long
    Dim i As Long, startIdx As Long, endIdx As Long
    Dim s As String, r As String
    lastRow = configWs.Cells(configWs.Rows.count, 1).End(xlUp).row
    total = lastRow - 1
    If total <= 0 Then
        ShowPagedConfigPicker = 0
        Exit Function
    End If
    pageSize = 10
    maxPage = (total + pageSize - 1) \ pageSize
    page = 1
    Do
        startIdx = (page - 1) * pageSize + 1
        If startIdx > total Then startIdx = total
        endIdx = page * pageSize
        If endIdx > total Then endIdx = total
        s = "請輸入編號選擇（< 上一頁、> 下一頁、取消留空）" & vbCrLf
        s = s & "第 " & page & "/" & maxPage & " 頁，共 " & total & " 組" & vbCrLf & vbCrLf
        For i = startIdx To endIdx
            s = s & i & ". " & configWs.Cells(i + 1, 1).value & " (" & configWs.Cells(i + 1, 4).value & "穴)" & vbCrLf
            s = s & "   路徑: " & configWs.Cells(i + 1, 2).value & vbCrLf
        Next i
        r = InputBox(s, "載入配置", CStr(startIdx))
        If r = "" Then
            ShowPagedConfigPicker = 0
            Exit Function
        ElseIf r = ">" Then
            If page < maxPage Then
                page = page + 1
            Else
                page = 1
            End If
        ElseIf r = "<" Then
            If page > 1 Then
                page = page - 1
            Else
                page = maxPage
            End If
        ElseIf IsNumeric(r) Then
            i = CLng(r)
            If i >= 1 And i <= total Then
                ShowPagedConfigPicker = i + 1
                Exit Function
            End If
        End If
    Loop
End Function

Private Function ShowScrollableConfigPicker(configWs As Worksheet) As Long
    On Error GoTo Fallback
    Dim vbProj As Object, vbComp As Object, codeMod As Object, existing As String
    Set vbProj = ThisWorkbook.VBProject
    On Error Resume Next
    Set vbComp = vbProj.VBComponents("frmConfigPicker")
    On Error GoTo 0
    If vbComp Is Nothing Then
        Set vbComp = vbProj.VBComponents.Add(3)
        vbComp.Name = "frmConfigPicker"
        With vbComp.Designer
            .Caption = "載入配置"
            .Width = 700
            .Height = 500
            Dim lbl As Object, lst As Object, okb As Object, cab As Object
            Set lbl = .Controls.Add("Forms.Label.1", "lblInfo", True)
            lbl.Left = 12: lbl.Top = 12: lbl.Width = 660: lbl.Height = 24
            Set lst = .Controls.Add("Forms.ListBox.1", "lstConfigs", True)
            lst.Left = 12: lst.Top = 40: lst.Width = 660: lst.Height = 400
            lst.ColumnCount = 5
            lst.ColumnWidths = "40 pt;200 pt;50 pt;360 pt;0 pt"
            lst.IntegralHeight = False
            Set okb = .Controls.Add("Forms.CommandButton.1", "cmdOK", True)
            okb.Caption = "確定": okb.Left = 480: okb.Top = 450: okb.Width = 90: okb.Height = 26
            Set cab = .Controls.Add("Forms.CommandButton.1", "cmdCancel", True)
            cab.Caption = "取消": cab.Left = 582: cab.Top = 450: cab.Width = 90: cab.Height = 26
        End With
        Set codeMod = vbComp.CodeModule
        Dim s As String
        s = s & "Option Explicit" & vbCrLf
        s = s & "Private Sub cmdOK_Click()" & vbCrLf
        s = s & "    If Me.lstConfigs.ListIndex >= 0 Then" & vbCrLf
        s = s & "        Me.Tag = Me.lstConfigs.List(Me.lstConfigs.ListIndex, 4)" & vbCrLf
        s = s & "    Else" & vbCrLf
        s = s & "        Me.Tag = """"" & vbCrLf
        s = s & "    End If" & vbCrLf
        s = s & "    Unload Me" & vbCrLf
        s = s & "End Sub" & vbCrLf
        s = s & "Private Sub cmdCancel_Click()" & vbCrLf
        s = s & "    Me.Tag = """"" & vbCrLf
        s = s & "    Unload Me" & vbCrLf
        s = s & "End Sub" & vbCrLf
        s = s & "Private Sub lstConfigs_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
        s = s & "    If Me.lstConfigs.ListIndex >= 0 Then" & vbCrLf
        s = s & "        Me.Tag = Me.lstConfigs.List(Me.lstConfigs.ListIndex, 4)" & vbCrLf
        s = s & "        Unload Me" & vbCrLf
        s = s & "    End If" & vbCrLf
        s = s & "End Sub" & vbCrLf
        codeMod.AddFromString s
    End If
    Dim lastRow As Long, i As Long, frm As Object
    lastRow = configWs.Cells(configWs.Rows.count, 1).End(xlUp).row
    Set frm = VBA.UserForms.Add("frmConfigPicker")
    Dim lblC As Object, lstC As Object, c As Object
    For Each c In frm.Controls
        If lblC Is Nothing And TypeName(c) = "Label" Then Set lblC = c
        If lstC Is Nothing And TypeName(c) = "ListBox" Then Set lstC = c
    Next
    If Not lblC Is Nothing Then lblC.Caption = "請選擇配置，支援捲動與雙擊選取"
    If lstC Is Nothing Then GoTo Fallback
    On Error Resume Next
    lstC.ColumnCount = 5
    lstC.ColumnWidths = "40 pt;200 pt;50 pt;360 pt;0 pt"
    lstC.IntegralHeight = False
    On Error GoTo 0
    With lstC
        .Clear
        For i = 2 To lastRow
            .AddItem CStr(i - 1)
            .List(.ListCount - 1, 1) = configWs.Cells(i, 1).value
            .List(.ListCount - 1, 2) = configWs.Cells(i, 4).value & "穴"
            .List(.ListCount - 1, 3) = configWs.Cells(i, 2).value
            .List(.ListCount - 1, 4) = CStr(i)
        Next i
    End With
    frm.Show vbModal
    If Len(frm.Tag) = 0 Then
        ShowScrollableConfigPicker = 0
    Else
        ShowScrollableConfigPicker = CLng(frm.Tag)
    End If
    Exit Function
Fallback:
    ShowScrollableConfigPicker = ShowPagedConfigPicker(configWs)
End Function

' ===================================================================
' 創建配置工作表
' ===================================================================
Function CreateConfigSheet() As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Application.EnableEvents = False
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = "參數配置"
    
    ' 標題
    With ws.Range("A1:D1")
        .Merge
        .value = "Excel 檢驗報告數據分析工具 - 參數配置"
        .Font.Size = 16: .Font.Bold = True: .HorizontalAlignment = xlCenter
        .Interior.color = RGB(68, 114, 196): .Font.color = RGB(255, 255, 255): .RowHeight = 30
    End With
    
    ws.Range("B2").value = "別眨眼，奇妙之旅開始了"
    With ws.Range("A2").Font
        .Italic = True
        .color = RGB(0, 112, 192)
        .Bold = True
    End With
    
    lastRow = 3
    
    ' === 基本設定 ===
    Call AddSectionHeader(ws, lastRow, "基本設定", RGB(91, 155, 213))
    Call AddInputRowWithButton(ws, lastRow, "來源檔案路徑", "", True, "點擊按鈕選擇要處理的Excel檔案", "選擇檔案", "BrowseSourceFile"): ws.Range("B" & lastRow - 1).Name = "rngSourceFile"
    Call AddInputRowWithButton(ws, lastRow, "產品品號", "", True, "用於識別此檔案的品號(可選)", "載入歷史", "LoadHistoryConfig"): ws.Range("B" & lastRow - 1).Name = "rngFileKeyword"
    Call AddDropdownRow(ws, lastRow, "模穴數", Array("8", "16", "24", "32", "40", "48"), "", True, "選擇後會自動顯示/隱藏對應的設定行"): ws.Range("B" & lastRow - 1).Name = "rngCavityCount"
    Dim rCav As Long
    rCav = ws.Range("rngCavityCount").Row
    Dim btnCav As Object
    Set btnCav = ws.Buttons.Add(ws.Range("D" & rCav).Left, ws.Range("D" & rCav).Top, ws.Range("D" & rCav).Width - 2, ws.Range("D" & rCav).Height - 2)
    btnCav.OnAction = "ApplyCavityCountChange"
    btnCav.Caption = "套用"
    btnCav.Font.Size = 9
    lastRow = lastRow + 1
    
    ' === 樣本檔案 ===
    Call AddSectionHeader(ws, lastRow, "樣本檔案（用於選擇範圍）", RGB(255, 192, 0))
    Call AddInputRowWithButton(ws, lastRow, "樣本檔案路徑", "", False, "點擊按鈕開啟一個樣本檔案，用於選擇範圍", "開啟樣本", "OpenSampleFile"): ws.Range("B" & lastRow - 1).Name = "rngSampleFilePath"
    lastRow = lastRow + 1
    
    ' === 讀取範圍設定 ===
    Call AddSectionHeader(ws, lastRow, "讀取範圍設定", RGB(112, 173, 71))
    With ws.Range("A" & lastRow & ":D" & lastRow)
        .Merge
        .value = "重要：請先開啟樣本檔案，再點擊「選擇」按鈕選擇範圍"
        .HorizontalAlignment = xlCenter: .Font.Bold = True
        .Interior.color = RGB(255, 242, 204): .Font.color = RGB(255, 0, 0): .RowHeight = 25
    End With
    lastRow = lastRow + 1
    
    ' 第1-8穴
    Call AddInputRowWithButton(ws, lastRow, "第1-8穴 穴號範圍", "", True, "穴號所在的儲存格範圍", "選擇", "SelectRange_1_Cavity"): ws.Range("B" & lastRow - 1).Name = "rngCavityId_1"
    Call AddInputRowWithButton(ws, lastRow, "第1-8穴 數據範圍", "", True, "數據所在的儲存格範圍", "選擇", "SelectRange_1_Data"): ws.Range("B" & lastRow - 1).Name = "rngData_1"
    
    ' 第9-16穴
    Call AddInputRowWithButton(ws, lastRow, "第9-16穴 穴號範圍", "", False, "第9-16號穴的穴號範圍", "選擇", "SelectRange_2_Cavity"): ws.Range("B" & lastRow - 1).Name = "rngCavityId_2"
    Call AddInputRowWithButton(ws, lastRow, "第9-16穴 數據範圍", "", False, "第9-16號穴的數據範圍", "選擇", "SelectRange_2_Data"): ws.Range("B" & lastRow - 1).Name = "rngData_2"
    Call AddInputRowWithButton(ws, lastRow, "9-16穴 頁面偏移", "", False, "1=同一頁, 2=下一頁...", "複製上方", "CopyFromGroup1"): ws.Range("B" & lastRow - 1).Name = "rngOffset_2"
    Dim btnT2 As Object
    Set btnT2 = ws.Buttons.Add(ws.Range("D" & ws.Range("rngOffset_2").Row).Left, ws.Range("D" & ws.Range("rngOffset_2").Row).Top, ws.Range("D" & ws.Range("rngOffset_2").Row).Width - 2, ws.Range("D" & ws.Range("rngOffset_2").Row).Height - 2)
    btnT2.OnAction = "ToggleGroup_2"
    btnT2.Caption = "展開/收合"
    btnT2.Font.Size = 9
    
    ' 第17-24穴
    Call AddInputRowWithButton(ws, lastRow, "第17-24穴 穴號範圍", "", False, "第17-24號穴的穴號範圍", "選擇", "SelectRange_3_Cavity"): ws.Range("B" & lastRow - 1).Name = "rngCavityId_3"
    Call AddInputRowWithButton(ws, lastRow, "第17-24穴 數據範圍", "", False, "第17-24號穴的數據範圍", "選擇", "SelectRange_3_Data"): ws.Range("B" & lastRow - 1).Name = "rngData_3"
    Call AddInputRowWithButton(ws, lastRow, "17-24穴 頁面偏移", "", False, "1=同一頁, 2=下一頁...", "複製上方", "CopyFromGroup1"): ws.Range("B" & lastRow - 1).Name = "rngOffset_3"
    Dim btnT3 As Object
    Set btnT3 = ws.Buttons.Add(ws.Range("D" & ws.Range("rngOffset_3").Row).Left, ws.Range("D" & ws.Range("rngOffset_3").Row).Top, ws.Range("D" & ws.Range("rngOffset_3").Row).Width - 2, ws.Range("D" & ws.Range("rngOffset_3").Row).Height - 2)
    btnT3.OnAction = "ToggleGroup_3"
    btnT3.Caption = "展開/收合"
    btnT3.Font.Size = 9
    
    ' 第25-32穴
    Call AddInputRowWithButton(ws, lastRow, "第25-32穴 穴號範圍", "", False, "第25-32號穴的穴號範圍", "選擇", "SelectRange_4_Cavity"): ws.Range("B" & lastRow - 1).Name = "rngCavityId_4"
    Call AddInputRowWithButton(ws, lastRow, "第25-32穴 數據範圍", "", False, "第25-32號穴的數據範圍", "選擇", "SelectRange_4_Data"): ws.Range("B" & lastRow - 1).Name = "rngData_4"
    Call AddInputRowWithButton(ws, lastRow, "25-32穴 頁面偏移", "", False, "1=同一頁, 2=下一頁...", "複製上方", "CopyFromGroup1"): ws.Range("B" & lastRow - 1).Name = "rngOffset_4"
    Dim btnT4 As Object
    Set btnT4 = ws.Buttons.Add(ws.Range("D" & ws.Range("rngOffset_4").Row).Left, ws.Range("D" & ws.Range("rngOffset_4").Row).Top, ws.Range("D" & ws.Range("rngOffset_4").Row).Width - 2, ws.Range("D" & ws.Range("rngOffset_4").Row).Height - 2)
    btnT4.OnAction = "ToggleGroup_4"
    btnT4.Caption = "展開/收合"
    btnT4.Font.Size = 9
    
    ' 第33-40穴
    Call AddInputRowWithButton(ws, lastRow, "第33-40穴 穴號範圍", "", False, "第33-40號穴的穴號範圍", "選擇", "SelectRange_5_Cavity"): ws.Range("B" & lastRow - 1).Name = "rngCavityId_5"
    Call AddInputRowWithButton(ws, lastRow, "第33-40穴 數據範圍", "", False, "第33-40號穴的數據範圍", "選擇", "SelectRange_5_Data"): ws.Range("B" & lastRow - 1).Name = "rngData_5"
    Call AddInputRowWithButton(ws, lastRow, "33-40穴 頁面偏移", "", False, "1=同一頁, 2=下一頁...", "複製上方", "CopyFromGroup1"): ws.Range("B" & lastRow - 1).Name = "rngOffset_5"
    Dim btnT5 As Object
    Set btnT5 = ws.Buttons.Add(ws.Range("D" & ws.Range("rngOffset_5").Row).Left, ws.Range("D" & ws.Range("rngOffset_5").Row).Top, ws.Range("D" & ws.Range("rngOffset_5").Row).Width - 2, ws.Range("D" & ws.Range("rngOffset_5").Row).Height - 2)
    btnT5.OnAction = "ToggleGroup_5"
    btnT5.Caption = "展開/收合"
    btnT5.Font.Size = 9
    
    ' 第41-48穴
    Call AddInputRowWithButton(ws, lastRow, "第41-48穴 穴號範圍", "", False, "第41-48號穴的穴號範圍", "選擇", "SelectRange_6_Cavity"): ws.Range("B" & lastRow - 1).Name = "rngCavityId_6"
    Call AddInputRowWithButton(ws, lastRow, "第41-48穴 數據範圍", "", False, "第41-48號穴的數據範圍", "選擇", "SelectRange_6_Data"): ws.Range("B" & lastRow - 1).Name = "rngData_6"
    Call AddInputRowWithButton(ws, lastRow, "41-48穴 頁面偏移", "", False, "1=同一頁, 2=下一頁...", "複製上方", "CopyFromGroup1"): ws.Range("B" & lastRow - 1).Name = "rngOffset_6"
    Dim btnT6 As Object
    Set btnT6 = ws.Buttons.Add(ws.Range("D" & ws.Range("rngOffset_6").Row).Left, ws.Range("D" & ws.Range("rngOffset_6").Row).Top, ws.Range("D" & ws.Range("rngOffset_6").Row).Width - 2, ws.Range("D" & ws.Range("rngOffset_6").Row).Height - 2)
    btnT6.OnAction = "ToggleGroup_6"
    btnT6.Caption = "展開/收合"
    btnT6.Font.Size = 9
    lastRow = lastRow + 1
    
    ' === 配置管理 ===
    Call AddSectionHeader(ws, lastRow, "配置管理", RGB(155, 194, 230))
    Call AddInputRowWithButton(ws, lastRow, "配置名稱", "", False, "為當前配置命名（用於保存和載入）", "保存配置", "SaveCurrentConfig"): ws.Range("B" & lastRow - 1).Name = "rngConfigName"
    Call AddInputRowWithButton(ws, lastRow, "重置設置", "", False, "恢復到初始預設值", "重置", "ResetConfigToDefaults")
    lastRow = lastRow + 1
    
    ' === 開始處理按鈕 ===
    Dim btn As Button
    Dim btnWidth As Double
    btnWidth = ws.Columns("A").Width + ws.Columns("B").Width + ws.Columns("C").Width + ws.Columns("D").Width
    Set btn = ws.Buttons.Add(ws.Range("A" & lastRow).Left, ws.Range("A" & lastRow).Top, btnWidth, 35)
    btn.OnAction = "ProcessInspectionReports_SheetInput"
    btn.Caption = "開始處理"
    btn.Font.Bold = True
    btn.Font.Size = 14
    lastRow = lastRow + 2
    
    ' === 使用說明 ===
    With ws.Range("A" & lastRow & ":D" & (lastRow + 8))
        .Merge
        .value = "詳細使用說明：" & vbCrLf & vbCrLf & _
                "【穴號範圍】是什麼？" & vbCrLf & _
                "-> 檔案中顯示「1號穴」「2號穴」...「8號穴」的儲存格範圍" & vbCrLf & vbCrLf & _
                "【數據範圍】是什麼？" & vbCrLf & _
                "-> 檔案中顯示實際數據（如 5.86, 5.85...）的儲存格範圍" & vbCrLf & vbCrLf & _
                "【頁面偏移】是什麼？" & vbCrLf & _
                "-> 1 = 在同一個工作表" & vbCrLf & _
                "-> 2 = 在下一個工作表" & vbCrLf & _
                "-> 3 = 在下下個工作表"
        .WrapText = True: .VerticalAlignment = xlTop: .Interior.color = RGB(217, 234, 211)
        .Borders.LineStyle = xlContinuous: .Font.Size = 10
    End With
    
    ' 調整欄寬
    ws.Columns("A:A").ColumnWidth = 22: ws.Columns("B:B").ColumnWidth = 35
    ws.Columns("C:C").ColumnWidth = 35: ws.Columns("D:D").ColumnWidth = 12
    ws.Columns("B:B").HorizontalAlignment = xlLeft
    
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
    
    ' 手動調用一次，以根據預設值設定好介面
    Call SetupRangeOutlines(ws)
    Call UpdateRowVisibility(ws)
    Call SetAllInputsNoColor(ws)
    Call InstallConfigSheetChangeHandler(ws)
    
    ' 確保顏色規則正確應用
    Application.EnableEvents = True
    Call UpdateInputColors(ws)
    
    Application.EnableEvents = True
    Set CreateConfigSheet = ws
    Exit Function

ErrorHandler:
    MsgBox "建立配置工作表時發生錯誤：" & vbCrLf & Err.description, vbCritical, "錯誤"
    Application.EnableEvents = True
    Set CreateConfigSheet = Nothing
End Function



' ===================================================================
' UI 輔助函數
' ===================================================================
Private Sub AddSectionHeader(ws As Worksheet, ByRef row As Long, title As String, color As Long)
    With ws.Range("A" & row & ":D" & row)
        .Merge: .value = title: .Font.Bold = True: .Font.Size = 12
        .Interior.color = color: .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 25
    End With
    row = row + 1
End Sub

Private Sub AddInputRowWithButton(ws As Worksheet, ByRef row As Long, label As String, defaultValue As String, required As Boolean, description As String, buttonText As String, buttonAction As String)
    ws.Range("A" & row).value = label
    ws.Range("A" & row).Font.Bold = True
    
    ws.Range("B" & row).value = defaultValue
    ws.Range("B" & row).Interior.ColorIndex = xlColorIndexNone
    ws.Range("B" & row).Borders.LineStyle = xlContinuous
    
    ws.Range("C" & row).value = description
    ws.Range("C" & row).Font.Size = 9
    ws.Range("C" & row).Font.Italic = True
    ws.Range("C" & row).Font.color = RGB(100, 100, 100)
    
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("D" & row).Left, ws.Range("D" & row).Top, ws.Range("D" & row).Width - 2, ws.Range("D" & row).Height - 2)
    btn.OnAction = buttonAction
    btn.Caption = buttonText
    btn.Font.Size = 9
    
    ws.Range("A" & row & ":D" & row).VerticalAlignment = xlCenter
    ws.Range("A" & row & ":D" & row).RowHeight = 25
    row = row + 1
End Sub

Private Sub AddDropdownRow(ws As Worksheet, ByRef row As Long, label As String, options As Variant, defaultValue As String, required As Boolean, description As String)
    ws.Range("A" & row).value = label
    ws.Range("A" & row).Font.Bold = True
    
    With ws.Range("B" & row)
        .value = defaultValue
        .Interior.ColorIndex = xlColorIndexNone
        .Borders.LineStyle = xlContinuous
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(options, ",")
        End With
    End With
    
    ws.Range("C" & row).value = description
    ws.Range("C" & row).Font.Size = 9
    ws.Range("C" & row).Font.Italic = True
    ws.Range("C" & row).Font.color = RGB(100, 100, 100)
    
    ws.Range("A" & row & ":D" & row).VerticalAlignment = xlCenter
    ws.Range("A" & row & ":D" & row).RowHeight = 25
    row = row + 1
End Sub

' ===================================================================
' 範圍設定折疊
' ===================================================================
Private Sub SetupRangeOutlines(ws As Worksheet)
    Dim r1 As Long, r2 As Long, i As Long
    On Error Resume Next
    For i = 2 To 6
        r1 = ws.Range("rngCavityId_" & i).Row
        If i >= 2 Then
            r2 = ws.Range("rngOffset_" & i).Row
        Else
            r2 = ws.Range("rngData_" & i).Row
        End If
        ws.Rows(r1 & ":" & r2).Ungroup
        ws.Rows(r1 & ":" & r2).Group
    Next i
    On Error GoTo 0
End Sub

' ===================================================================
' 工作表事件：自動隱藏/顯示行（v23 重構版）
' ===================================================================
Public Sub UpdateRowVisibility(ByVal ws As Worksheet)
    Dim selectedCavities As Long
    Dim i As Long, r1 As Long, r2 As Long, needShow As Boolean, thr As Long
    On Error Resume Next
    selectedCavities = CLng(ws.Range("rngCavityCount").value)
    On Error GoTo 0
    
    ' 如果是 0 或無法解析，預設為 8 穴（隱藏所有額外組）
    If selectedCavities = 0 Then selectedCavities = 8
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    For i = 2 To 6
        r1 = ws.Range("rngCavityId_" & i).Row
        r2 = ws.Range("rngOffset_" & i).Row
        If ws.Rows(r1).OutlineLevel = 0 Or ws.Rows(r2).OutlineLevel = 0 Then
            ws.Rows(r1 & ":" & r2).Ungroup
            ws.Rows(r1 & ":" & r2).Group
        End If
        thr = i * 8
        needShow = (selectedCavities >= thr)
        ws.Rows(r1 & ":" & r2).Hidden = Not needShow
        ws.Rows(r1 & ":" & r2).ShowDetail = needShow
    Next i
    On Error GoTo 0
    
    Application.ScreenUpdating = True
End Sub

' ===================================================================
' 按鈕功能
' ===================================================================
Sub OpenSampleFile()
    Dim ws As Worksheet, filePath As String
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If Not g_SampleWorkbook Is Nothing Then
        On Error Resume Next
        g_SampleWorkbook.Close SaveChanges:=False
        Set g_SampleWorkbook = Nothing
        On Error GoTo 0
    End If
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "選擇一個樣本檔案（用於選擇範圍）"
        .Filters.Clear
        .Filters.Add "Excel 檔案", "*.xls; *.xlsx; *.xlsm"
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            On Error GoTo ErrorHandler
            Set g_SampleWorkbook = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
            ws.Range("rngSampleFilePath").value = filePath
            Call SetCellColorByValue(ws.Range("rngSampleFilePath"))
            MsgBox "已成功開啟樣本檔案。", vbInformation, "操作成功"
        End If
    End With
    Exit Sub
ErrorHandler:
    MsgBox "錯誤：無法開啟檔案。" & vbCrLf & Err.description, vbCritical, "操作失敗"
End Sub

Sub BrowseSourceFile()
    Dim ws As Worksheet, filePath As String
    Set ws = ThisWorkbook.Worksheets("參數配置")
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "選擇要處理的Excel檔案"
        .Filters.Clear
        .Filters.Add "Excel 檔案", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            ws.Range("rngSourceFile").value = filePath
            Call SetCellColorByValue(ws.Range("rngSourceFile"))
            
            ' 自動從檔名提取品號(如果尚未填寫)
            If Trim(ws.Range("rngFileKeyword").value) = "" Then
                Dim fso As Object
                Set fso = CreateObject("Scripting.FileSystemObject")
                Dim fileName As String
                fileName = fso.GetBaseName(filePath)
                ws.Range("rngFileKeyword").value = fileName
                Call SetCellColorByValue(ws.Range("rngFileKeyword"))
            End If
        End If
    End With
End Sub

' ===================================================================
' 智能範圍選擇器
' ===================================================================
Sub SelectRange_1_Cavity(): Call SelectRangeHelper("rngCavityId_1", "第1-8穴 穴號範圍", "請選擇「1號穴」...「8號穴」的範圍"): End Sub
Sub SelectRange_1_Data(): Call SelectRangeHelper("rngData_1", "第1-8穴 數據範圍", "請選擇第1-8穴的數據範圍"): End Sub
Sub SelectRange_2_Cavity(): Call SelectRangeHelper("rngCavityId_2", "第9-16穴 穴號範圍", "請選擇「9號穴」...「16號穴」的範圍"): End Sub
Sub SelectRange_2_Data(): Call SelectRangeHelper("rngData_2", "第9-16穴 數據範圍", "請選擇第9-16穴的數據範圍"): End Sub
Sub SelectRange_3_Cavity(): Call SelectRangeHelper("rngCavityId_3", "第17-24穴 穴號範圍", "請選擇第17-24穴的穴號範圍"): End Sub
Sub SelectRange_3_Data(): Call SelectRangeHelper("rngData_3", "第17-24穴 數據範圍", "請選擇第17-24穴的數據範圍"): End Sub
Sub SelectRange_4_Cavity(): Call SelectRangeHelper("rngCavityId_4", "第25-32穴 穴號範圍", "請選擇第25-32穴的穴號範圍"): End Sub
Sub SelectRange_4_Data(): Call SelectRangeHelper("rngData_4", "第25-32穴 數據範圍", "請選擇第25-32穴的數據範圍"): End Sub
Sub SelectRange_5_Cavity(): Call SelectRangeHelper("rngCavityId_5", "第33-40穴 穴號範圍", "請選擇第33-40穴的穴號範圍"): End Sub
Sub SelectRange_5_Data(): Call SelectRangeHelper("rngData_5", "第33-40穴 數據範圍", "請選擇第33-40穴的數據範圍"): End Sub
Sub SelectRange_6_Cavity(): Call SelectRangeHelper("rngCavityId_6", "第41-48穴 穴號範圍", "請選擇第41-48穴的穴號範圍"): End Sub
Sub SelectRange_6_Data(): Call SelectRangeHelper("rngData_6", "第41-48穴 數據範圍", "請選擇第41-48穴的數據範圍"): End Sub

Private Sub SelectRangeHelper(targetRangeName As String, rangeName As String, promptText As String)
    Dim ws As Worksheet, selectedRange As Range
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If g_SampleWorkbook Is Nothing Then
        If MsgBox("錯誤：尚未開啟樣本檔案！" & vbCrLf & "是否現在開啟？", vbYesNo + vbQuestion, "需要樣本檔案") = vbYes Then
            Call OpenSampleFile
            If g_SampleWorkbook Is Nothing Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    g_SampleWorkbook.Activate
    On Error Resume Next
    Set selectedRange = Application.InputBox(Prompt:=promptText, title:="選擇範圍 - " & rangeName, Type:=8)
    On Error GoTo 0
    ThisWorkbook.Activate
    ws.Activate
    If Not selectedRange Is Nothing Then
        ws.Range(targetRangeName).value = selectedRange.Address(False, False)
        Call SetCellColorByValue(ws.Range(targetRangeName))
    End If
End Sub

Sub ToggleRangeGroup(ws As Worksheet, groupIndex As Long)
    Dim r1 As Long, r2 As Long, isHidden As Boolean
    r1 = ws.Range("rngCavityId_" & groupIndex).Row
    r2 = ws.Range("rngOffset_" & groupIndex).Row
    isHidden = ws.Rows(r1).Hidden Or ws.Rows(r2).Hidden
    ws.Rows(r1 & ":" & r2).Hidden = Not isHidden
End Sub

Sub ToggleGroup_2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If ws Is Nothing Then Exit Sub
    Call ToggleRangeGroup(ws, 2)
End Sub

Sub ToggleGroup_3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If ws Is Nothing Then Exit Sub
    Call ToggleRangeGroup(ws, 3)
End Sub

Sub ToggleGroup_4()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If ws Is Nothing Then Exit Sub
    Call ToggleRangeGroup(ws, 4)
End Sub

Sub ToggleGroup_5()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If ws Is Nothing Then Exit Sub
    Call ToggleRangeGroup(ws, 5)
End Sub

Sub ToggleGroup_6()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If ws Is Nothing Then Exit Sub
    Call ToggleRangeGroup(ws, 6)
End Sub

Sub ApplyCavityCountChange()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If ws Is Nothing Then Exit Sub
    Call SetupRangeOutlines(ws)
    Call UpdateRowVisibility(ws)
    Call SetCellColorByValue(ws.Range("rngCavityCount"))
End Sub

' 手動刷新所有輸入欄位的顏色
Sub RefreshAllInputColors()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("參數配置")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "找不到「參數配置」工作表。", vbExclamation
        Exit Sub
    End If
    Application.EnableEvents = True
    Call UpdateInputColors(ws)
    MsgBox "已刷新所有輸入欄位的顏色。", vbInformation
End Sub

Private Sub InstallConfigSheetChangeHandler(ws As Worksheet)
    On Error GoTo Handler
    Dim vbProj As Object, vbComp As Object, codeMod As Object, existing As String
    Set vbProj = ThisWorkbook.VBProject
    Set vbComp = vbProj.VBComponents(ws.CodeName)
    Set codeMod = vbComp.CodeModule
    Dim lineCount As Long
    lineCount = codeMod.CountOfLines
    existing = codeMod.Lines(1, lineCount)
    If InStr(1, existing, "Private Sub Worksheet_Change(ByVal Target As Range)", vbTextCompare) = 0 Then
        Dim s As String
        s = s & "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
        s = s & "    On Error Resume Next" & vbCrLf
        s = s & "    If Not Intersect(Target, Me.Range(""rngCavityCount"")) Is Nothing Then" & vbCrLf
        s = s & "        Call UpdateRowVisibility(Me)" & vbCrLf
        s = s & "        Call SetCellColorByValue(Me.Range(""rngCavityCount""))" & vbCrLf
        s = s & "    ElseIf Not Intersect(Target, Me.Columns(""B"")) Is Nothing Then" & vbCrLf
        s = s & "        Call SetCellColorByValue(Target)" & vbCrLf
        s = s & "    End If" & vbCrLf
        s = s & "    On Error GoTo 0" & vbCrLf
        s = s & "End Sub" & vbCrLf
        codeMod.AddFromString s
    End If
    Exit Sub
Handler:
    MsgBox "提示：若要在鍵入後即刻變色，請於Excel選項中啟用『信任對VBA工程的存取』。", vbInformation, "需要權限"
End Sub

Sub ResetConfigToDefaults()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("參數配置")
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    ws.Range("rngSourceFile").value = ""
    ws.Range("rngFileKeyword").value = ""
    ws.Range("rngCavityCount").value = ""
    ws.Range("rngSampleFilePath").value = ""
    ws.Range("rngConfigName").value = ""
    ws.Range("rngCavityId_1").value = ""
    ws.Range("rngData_1").value = ""
    ws.Range("rngCavityId_2").value = ""
    ws.Range("rngData_2").value = ""
    ws.Range("rngOffset_2").value = ""
    ws.Range("rngCavityId_3").value = ""
    ws.Range("rngData_3").value = ""
    ws.Range("rngOffset_3").value = ""
    ws.Range("rngCavityId_4").value = ""
    ws.Range("rngData_4").value = ""
    ws.Range("rngOffset_4").value = ""
    ws.Range("rngCavityId_5").value = ""
    ws.Range("rngData_5").value = ""
    ws.Range("rngOffset_5").value = ""
    ws.Range("rngCavityId_6").value = ""
    ws.Range("rngData_6").value = ""
    ws.Range("rngOffset_6").value = ""
    Call UpdateRowVisibility(ws)
    Call SetAllInputsNoColor(ws)
    Application.ScreenUpdating = True
End Sub

' ===================================================================
' 快速複製
' ===================================================================
Sub CopyFromGroup1()
    Dim ws As Worksheet, cavityRange1 As String, dataRange1 As String
    Dim targetRow As Long, buttonCell As Range, changedGroupIndex As Long
    Set ws = ThisWorkbook.Worksheets("參數配置")
    cavityRange1 = ws.Range("rngCavityId_1").value
    dataRange1 = ws.Range("rngData_1").value
    If cavityRange1 = "" Or dataRange1 = "" Then
        MsgBox "警告：請先設定第1-8穴的範圍！", vbExclamation, "無法複製"
        Exit Sub
    End If
    On Error Resume Next
    Set buttonCell = ws.Buttons(Application.Caller).TopLeftCell
    On Error GoTo 0
    If buttonCell Is Nothing Then Exit Sub
    targetRow = buttonCell.row
    changedGroupIndex = 0
    Select Case targetRow
        Case ws.Range("rngOffset_2").row: ws.Range("rngCavityId_2").value = cavityRange1: ws.Range("rngData_2").value = dataRange1: changedGroupIndex = 2
        Case ws.Range("rngOffset_3").row: ws.Range("rngCavityId_3").value = cavityRange1: ws.Range("rngData_3").value = dataRange1: changedGroupIndex = 3
        Case ws.Range("rngOffset_4").row: ws.Range("rngCavityId_4").value = cavityRange1: ws.Range("rngData_4").value = dataRange1: changedGroupIndex = 4
        Case ws.Range("rngOffset_5").row: ws.Range("rngCavityId_5").value = cavityRange1: ws.Range("rngData_5").value = dataRange1: changedGroupIndex = 5
        Case ws.Range("rngOffset_6").row: ws.Range("rngCavityId_6").value = cavityRange1: ws.Range("rngData_6").value = dataRange1: changedGroupIndex = 6
    End Select
    If changedGroupIndex > 0 Then
        Call SetCellColorByValue(ws.Range("rngCavityId_" & changedGroupIndex))
        Call SetCellColorByValue(ws.Range("rngData_" & changedGroupIndex))
    End If
End Sub

' ===================================================================
' 保存和載入配置
' ===================================================================
Private Sub EnsureConfigHistoryHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("配置名稱", "來源檔案", "產品品號", "模穴數", _
                    "1-8穴號", "1-8數據", "1-8偏移", _
                    "9-16穴號", "9-16數據", "9-16偏移", _
                    "17-24穴號", "17-24數據", "17-24偏移", _
                    "25-32穴號", "25-32數據", "25-32偏移", _
                    "33-40穴號", "33-40數據", "33-40偏移", _
                    "41-48穴號", "41-48數據", "41-48偏移", _
                    "樣本檔路徑", "保存時間")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
    Next i
End Sub

Sub SaveCurrentConfig()
    Dim ws As Worksheet, configWs As Worksheet, configName As String, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("參數配置")
    configName = Trim(ws.Range("rngConfigName").value)
    If configName = "" Then
        MsgBox "錯誤：請先輸入配置名稱！", vbExclamation, "無法保存"
        Exit Sub
    End If
    On Error Resume Next
    Set configWs = ThisWorkbook.Worksheets("配置歷史")
    On Error GoTo 0
    If configWs Is Nothing Then
        Set configWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        configWs.Name = "配置歷史"
        Call EnsureConfigHistoryHeaders(configWs)
    Else
        Call EnsureConfigHistoryHeaders(configWs)
    End If
    lastRow = configWs.Cells(configWs.Rows.count, 1).End(xlUp).row + 1
    With configWs.Rows(lastRow)
        .Cells(1).value = configName
        .Cells(2).value = ws.Range("rngSourceFile").value
        .Cells(3).value = ws.Range("rngFileKeyword").value
        .Cells(4).value = ws.Range("rngCavityCount").value
        .Cells(5).value = ws.Range("rngCavityId_1").value
        .Cells(6).value = ws.Range("rngData_1").value
        .Cells(7).value = ""
        .Cells(8).value = ws.Range("rngCavityId_2").value
        .Cells(9).value = ws.Range("rngData_2").value
        .Cells(10).value = ws.Range("rngOffset_2").value
        .Cells(11).value = ws.Range("rngCavityId_3").value
        .Cells(12).value = ws.Range("rngData_3").value
        .Cells(13).value = ws.Range("rngOffset_3").value
        .Cells(14).value = ws.Range("rngCavityId_4").value
        .Cells(15).value = ws.Range("rngData_4").value
        .Cells(16).value = ws.Range("rngOffset_4").value
        .Cells(17).value = ws.Range("rngCavityId_5").value
        .Cells(18).value = ws.Range("rngData_5").value
        .Cells(19).value = ws.Range("rngOffset_5").value
        .Cells(20).value = ws.Range("rngCavityId_6").value
        .Cells(21).value = ws.Range("rngData_6").value
        .Cells(22).value = ws.Range("rngOffset_6").value
        .Cells(23).value = ws.Range("rngSampleFilePath").value
        .Cells(24).value = Now
    End With
    MsgBox "配置已成功保存。", vbInformation, "保存成功"
    Call SetCellColorByValue(ws.Range("rngConfigName"))
End Sub

Sub LoadHistoryConfig()
    Dim ws As Worksheet, configWs As Worksheet, lastRow As Long
    Dim selectedRow As Long
    Set ws = ThisWorkbook.Worksheets("參數配置")
    On Error Resume Next
    Set configWs = ThisWorkbook.Worksheets("配置歷史")
    On Error GoTo 0
    If configWs Is Nothing Or configWs.Cells(configWs.Rows.count, 1).End(xlUp).row < 2 Then
        MsgBox "提示：尚未保存任何配置。", vbInformation, "無配置記錄"
        Exit Sub
    End If
    selectedRow = ShowScrollableConfigPicker(configWs)
    If selectedRow < 2 Then Exit Sub
    With ws
        .Range("rngSourceFile").value = configWs.Cells(selectedRow, 2).value
        Call SetCellColorByValue(.Range("rngSourceFile"))
        .Range("rngFileKeyword").value = configWs.Cells(selectedRow, 3).value
        Call SetCellColorByValue(.Range("rngFileKeyword"))
        .Range("rngCavityCount").value = configWs.Cells(selectedRow, 4).value
        Call SetCellColorByValue(.Range("rngCavityCount"))
        .Range("rngCavityId_1").value = configWs.Cells(selectedRow, 5).value
        Call SetCellColorByValue(.Range("rngCavityId_1"))
        .Range("rngData_1").value = configWs.Cells(selectedRow, 6).value
        Call SetCellColorByValue(.Range("rngData_1"))
        .Range("rngCavityId_2").value = configWs.Cells(selectedRow, 8).value
        Call SetCellColorByValue(.Range("rngCavityId_2"))
        .Range("rngData_2").value = configWs.Cells(selectedRow, 9).value
        Call SetCellColorByValue(.Range("rngData_2"))
        .Range("rngOffset_2").value = configWs.Cells(selectedRow, 10).value
        Call SetCellColorByValue(.Range("rngOffset_2"))
        .Range("rngCavityId_3").value = configWs.Cells(selectedRow, 11).value
        Call SetCellColorByValue(.Range("rngCavityId_3"))
        .Range("rngData_3").value = configWs.Cells(selectedRow, 12).value
        Call SetCellColorByValue(.Range("rngData_3"))
        .Range("rngOffset_3").value = configWs.Cells(selectedRow, 13).value
        Call SetCellColorByValue(.Range("rngOffset_3"))
        .Range("rngCavityId_4").value = configWs.Cells(selectedRow, 14).value
        Call SetCellColorByValue(.Range("rngCavityId_4"))
        .Range("rngData_4").value = configWs.Cells(selectedRow, 15).value
        Call SetCellColorByValue(.Range("rngData_4"))
        .Range("rngOffset_4").value = configWs.Cells(selectedRow, 16).value
        Call SetCellColorByValue(.Range("rngOffset_4"))
        .Range("rngCavityId_5").value = configWs.Cells(selectedRow, 17).value
        Call SetCellColorByValue(.Range("rngCavityId_5"))
        .Range("rngData_5").value = configWs.Cells(selectedRow, 18).value
        Call SetCellColorByValue(.Range("rngData_5"))
        .Range("rngOffset_5").value = configWs.Cells(selectedRow, 19).value
        Call SetCellColorByValue(.Range("rngOffset_5"))
        .Range("rngCavityId_6").value = configWs.Cells(selectedRow, 20).value
        Call SetCellColorByValue(.Range("rngCavityId_6"))
        .Range("rngData_6").value = configWs.Cells(selectedRow, 21).value
        Call SetCellColorByValue(.Range("rngData_6"))
        .Range("rngOffset_6").value = configWs.Cells(selectedRow, 22).value
        Call SetCellColorByValue(.Range("rngOffset_6"))
        .Range("rngSampleFilePath").value = configWs.Cells(selectedRow, 23).value
        Call SetCellColorByValue(.Range("rngSampleFilePath"))
        .Range("rngConfigName").value = configWs.Cells(selectedRow, 1).value
        Call SetCellColorByValue(.Range("rngConfigName"))
    End With
    Call UpdateRowVisibility(ws)
    Application.EnableEvents = True
    Call UpdateInputColors(ws)
    ws.Activate
    MsgBox "配置已成功載入。", vbInformation, "載入成功"
End Sub

Sub ConfirmLoadSelectedConfig()
    Dim ws As Worksheet, configWs As Worksheet, pickWs As Worksheet
    Dim r As Long, lastRow As Long, selectedRow As Long
    Set ws = ThisWorkbook.Worksheets("參數配置")
    Set pickWs = ThisWorkbook.Worksheets("配置選擇")
    Set configWs = ThisWorkbook.Worksheets("配置歷史")
    r = pickWs.ActiveCell.Row
    lastRow = configWs.Cells(configWs.Rows.count, 1).End(xlUp).row
    If r < 2 Or r > lastRow Then
        MsgBox "請選取有效的配置行。", vbExclamation
        Exit Sub
    End If
    selectedRow = r
    With ws
        .Range("rngSourceFile").value = configWs.Cells(selectedRow, 2).value
        Call SetCellColorByValue(.Range("rngSourceFile"))
        .Range("rngFileKeyword").value = configWs.Cells(selectedRow, 3).value
        Call SetCellColorByValue(.Range("rngFileKeyword"))
        .Range("rngCavityCount").value = configWs.Cells(selectedRow, 4).value
        Call SetCellColorByValue(.Range("rngCavityCount"))
        .Range("rngCavityId_1").value = configWs.Cells(selectedRow, 5).value
        Call SetCellColorByValue(.Range("rngCavityId_1"))
        .Range("rngData_1").value = configWs.Cells(selectedRow, 6).value
        Call SetCellColorByValue(.Range("rngData_1"))
        .Range("rngCavityId_2").value = configWs.Cells(selectedRow, 8).value
        Call SetCellColorByValue(.Range("rngCavityId_2"))
        .Range("rngData_2").value = configWs.Cells(selectedRow, 9).value
        Call SetCellColorByValue(.Range("rngData_2"))
        .Range("rngOffset_2").value = configWs.Cells(selectedRow, 10).value
        Call SetCellColorByValue(.Range("rngOffset_2"))
        .Range("rngCavityId_3").value = configWs.Cells(selectedRow, 11).value
        Call SetCellColorByValue(.Range("rngCavityId_3"))
        .Range("rngData_3").value = configWs.Cells(selectedRow, 12).value
        Call SetCellColorByValue(.Range("rngData_3"))
        .Range("rngOffset_3").value = configWs.Cells(selectedRow, 13).value
        Call SetCellColorByValue(.Range("rngOffset_3"))
        .Range("rngCavityId_4").value = configWs.Cells(selectedRow, 14).value
        Call SetCellColorByValue(.Range("rngCavityId_4"))
        .Range("rngData_4").value = configWs.Cells(selectedRow, 15).value
        Call SetCellColorByValue(.Range("rngData_4"))
        .Range("rngOffset_4").value = configWs.Cells(selectedRow, 16).value
        Call SetCellColorByValue(.Range("rngOffset_4"))
        .Range("rngCavityId_5").value = configWs.Cells(selectedRow, 17).value
        Call SetCellColorByValue(.Range("rngCavityId_5"))
        .Range("rngData_5").value = configWs.Cells(selectedRow, 18).value
        Call SetCellColorByValue(.Range("rngData_5"))
        .Range("rngOffset_5").value = configWs.Cells(selectedRow, 19).value
        Call SetCellColorByValue(.Range("rngOffset_5"))
        .Range("rngCavityId_6").value = configWs.Cells(selectedRow, 20).value
        Call SetCellColorByValue(.Range("rngCavityId_6"))
        .Range("rngData_6").value = configWs.Cells(selectedRow, 21).value
        Call SetCellColorByValue(.Range("rngData_6"))
        .Range("rngOffset_6").value = configWs.Cells(selectedRow, 22).value
        Call SetCellColorByValue(.Range("rngOffset_6"))
        .Range("rngSampleFilePath").value = configWs.Cells(selectedRow, 23).value
        Call SetCellColorByValue(.Range("rngSampleFilePath"))
        .Range("rngConfigName").value = configWs.Cells(selectedRow, 1).value
        Call SetCellColorByValue(.Range("rngConfigName"))
    End With
    Call UpdateRowVisibility(ws)
    Application.EnableEvents = True
    Call UpdateInputColors(ws)
    ws.Activate
    MsgBox "配置已成功載入。", vbInformation, "載入成功"
End Sub

' ===================================================================
' 從工作表讀取參數
' ===================================================================
Function ReadParametersFromSheet(ws As Worksheet, ByRef params As UserInputParams) As Boolean
    On Error GoTo ErrorHandler
    params.folderPath = Trim(ws.Range("rngSourceFile").value)
    params.fileKeyword = Trim(ws.Range("rngFileKeyword").value)
    params.cavitiesForUI = CLng(ws.Range("rngCavityCount").value)
    
    ' 驗證檔案是否存在
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(params.folderPath) Then
        MsgBox "錯誤：找不到指定的檔案：" & vbCrLf & params.folderPath, vbCritical, "檔案不存在"
        ReadParametersFromSheet = False
        Exit Function
    End If
    
    Dim ranges(1 To 6) As String, cavityRanges(1 To 6) As String, offsets(1 To 6) As String
    cavityRanges(1) = Trim(ws.Range("rngCavityId_1").value): ranges(1) = Trim(ws.Range("rngData_1").value)
    cavityRanges(2) = Trim(ws.Range("rngCavityId_2").value): ranges(2) = Trim(ws.Range("rngData_2").value): offsets(2) = Trim(ws.Range("rngOffset_2").value)
    cavityRanges(3) = Trim(ws.Range("rngCavityId_3").value): ranges(3) = Trim(ws.Range("rngData_3").value): offsets(3) = Trim(ws.Range("rngOffset_3").value)
    cavityRanges(4) = Trim(ws.Range("rngCavityId_4").value): ranges(4) = Trim(ws.Range("rngData_4").value): offsets(4) = Trim(ws.Range("rngOffset_4").value)
    cavityRanges(5) = Trim(ws.Range("rngCavityId_5").value): ranges(5) = Trim(ws.Range("rngData_5").value): offsets(5) = Trim(ws.Range("rngOffset_5").value)
    cavityRanges(6) = Trim(ws.Range("rngCavityId_6").value): ranges(6) = Trim(ws.Range("rngData_6").value): offsets(6) = Trim(ws.Range("rngOffset_6").value)
    
    If params.folderPath = "" Or ranges(1) = "" Or cavityRanges(1) = "" Then
        MsgBox "錯誤：請填寫所有必填欄位！", vbExclamation, "參數不完整"
        ReadParametersFromSheet = False
        Exit Function
    End If
    
    Call AutoConfigureCavityGroups(params, params.cavitiesForUI, ranges, cavityRanges, offsets)
    ReadParametersFromSheet = True
    Exit Function
    
ErrorHandler:
    MsgBox "錯誤：讀取參數時發生錯誤。" & vbCrLf & Err.description, vbCritical, "讀取錯誤"
    ReadParametersFromSheet = False
End Function

Private Sub AutoConfigureCavityGroups(ByRef params As UserInputParams, cavitiesForUI As Long, ranges() As String, cavityRanges() As String, offsets() As String)
    Dim i As Long
    For i = 1 To 6
        gPageOffset(i) = 0
        gDataRange(i) = ""
        gCavityIdRange(i) = ""
    Next i
    
    gPageOffset(1) = 0
    gDataRange(1) = ranges(1)
    gCavityIdRange(1) = cavityRanges(1)
    
    For i = 2 To 6
        If cavitiesForUI >= i * 8 Then
            If IsNumeric(offsets(i)) Then
                Dim v As Long
                v = CLng(offsets(i))
                If v <= 0 Then
                    gPageOffset(i) = 0
                Else
                    gPageOffset(i) = v - 1
                End If
            Else
                gPageOffset(i) = 0
            End If
            gDataRange(i) = ranges(i)
            gCavityIdRange(i) = cavityRanges(i)
        End If
    Next i
End Sub

Function ValidateCavityGroupInput(params As UserInputParams) As Boolean
    If gDataRange(1) <> "" And Trim(gCavityIdRange(1)) <> "" Then
        ValidateCavityGroupInput = True
    Else
        MsgBox "錯誤：第1-8穴的數據範圍與穴號範圍為必填項。", vbCritical, "輸入錯誤"
        ValidateCavityGroupInput = False
    End If
End Function

' ===================================================================
' 核心處理邏輯
' ===================================================================
' 提取數據範圍的數據，只提取第一行，並返回檢驗項目名稱
Private Function ExtractCavityDataWithInspectionItem(ws As Worksheet, grp As CavityGroupConfig, ByRef dataDict As Object, ByRef inspectionItem As String) As Boolean
    Dim dataRng As Range, idRng As Range
    Dim dataTopRow As Long, dataBottomRow As Long, dataStartCol As Long
    Dim idTopRow As Long, idStartCol As Long, colOffset As Long, taken As Long
    Dim cellValue As Variant, cleanValue As String, cavityId As String
    Dim searchCol As Long
    Dim tempValue As String
    
    taken = 0
    On Error Resume Next
    Set dataRng = ws.Range(grp.dataRange)
    Set idRng = ws.Range(grp.cavityIdRange)
    If dataRng Is Nothing Or idRng Is Nothing Then
        ExtractCavityDataWithInspectionItem = False
        Exit Function
    End If
    On Error GoTo 0
    
    dataTopRow = dataRng.row
    dataBottomRow = dataRng.Rows(dataRng.Rows.count).row
    dataStartCol = dataRng.Column
    idTopRow = idRng.row
    idStartCol = idRng.Column
    
    If inspectionItem = "" Then
        On Error Resume Next
        Dim m As Range
        Set m = ws.Cells(dataTopRow, 1).MergeArea
        tempValue = Trim(m.Cells(1, 1).Text)
        If tempValue = "" Then tempValue = Trim(CStr(m.Cells(1, 1).value))
        
        If tempValue = "" Then
            Set m = ws.Cells(dataTopRow, 2).MergeArea
            tempValue = Trim(m.Cells(1, 1).Text)
            If tempValue = "" Then tempValue = Trim(CStr(m.Cells(1, 1).value))
        End If
        tempValue = Replace(Replace(tempValue, "(", ""), ")", "")
        tempValue = Trim(tempValue)
        If tempValue <> "" Then inspectionItem = tempValue
        If inspectionItem = "" Then
            For searchCol = dataStartCol - 1 To 1 Step -1
                ' 使用 .Text 避免 J 字母衝突
                tempValue = Trim(ws.Cells(dataTopRow, searchCol).Text)
                If tempValue = "" Then tempValue = Trim(CStr(ws.Cells(dataTopRow, searchCol).value))
                
                If tempValue <> "" And Not IsNumeric(tempValue) Then
                    inspectionItem = tempValue
                    Exit For
                End If
            Next searchCol
        End If
        If inspectionItem = "" Then inspectionItem = ws.Name
        On Error GoTo 0
    End If
    
    ' 提取穴號數據（只取第一行）
    For colOffset = 0 To idRng.Columns.count - 1
        cavityId = CStr(ws.Cells(idTopRow, idStartCol + colOffset).MergeArea.Cells(1, 1).value)
        cavityId = Trim(cavityId)
        
        If cavityId <> "" Then
            cellValue = ws.Cells(dataTopRow, dataStartCol + colOffset).MergeArea.Cells(1, 1).Value2
            If (IsEmpty(cellValue) Or CStr(cellValue) = "") And (dataBottomRow > dataTopRow) Then
                cellValue = ws.Cells(dataBottomRow, dataStartCol + colOffset).MergeArea.Cells(1, 1).Value2
            End If
            
            cleanValue = CleanCellValue(CStr(cellValue))
            If IsNumeric(cleanValue) And cleanValue <> "" Then
                If Not dataDict.exists(cavityId) Then
                    dataDict.Add cavityId, CDbl(cleanValue)
                End If
                taken = taken + 1
            End If
        End If
    Next colOffset
    
    ExtractCavityDataWithInspectionItem = (taken > 0)
End Function

Private Function ExtractInspectionItemsFromGroup(ws As Worksheet, grp As CavityGroupConfig) As Collection
    Dim dataRng As Range, idRng As Range
    Dim dataTopRow As Long, dataBottomRow As Long, dataStartCol As Long
    Dim idTopRow As Long, idStartCol As Long
    Dim rowOffset As Long, colOffset As Long
    Dim cavityId As String, cellValue As Variant, cleanValue As String
    Dim itemName As String
    Dim rec As Object, dataDict As Object
    Dim m As Range, tempValue As String
    Dim result As Collection
    Set result = New Collection
    On Error Resume Next
    Set dataRng = ws.Range(grp.dataRange)
    Set idRng = ws.Range(grp.cavityIdRange)
    If dataRng Is Nothing Or idRng Is Nothing Then
        Set ExtractInspectionItemsFromGroup = Nothing
        Exit Function
    End If
    On Error GoTo 0
    dataTopRow = dataRng.row
    dataBottomRow = dataRng.Rows(dataRng.Rows.count).row
    dataStartCol = dataRng.Column
    idTopRow = idRng.row
    idStartCol = idRng.Column
    For rowOffset = 0 To dataRng.Rows.count - 1
        Set dataDict = CreateObject("Scripting.Dictionary")
        itemName = ""
        On Error Resume Next
        Set m = ws.Cells(dataTopRow + rowOffset, 1).MergeArea
        tempValue = Trim(CStr(m.Cells(1, 1).value))
        If tempValue = "" Then
            Set m = ws.Cells(dataTopRow + rowOffset, 2).MergeArea
            tempValue = Trim(CStr(m.Cells(1, 1).value))
        End If
        tempValue = Replace(Replace(tempValue, "(", ""), ")", "")
        itemName = Trim(tempValue)
        On Error GoTo 0
        For colOffset = 0 To idRng.Columns.count - 1
            cavityId = CStr(ws.Cells(idTopRow, idStartCol + colOffset).MergeArea.Cells(1, 1).value)
            cavityId = Trim(cavityId)
            If cavityId <> "" Then
                cellValue = ws.Cells(dataTopRow + rowOffset, dataStartCol + colOffset).MergeArea.Cells(1, 1).Value2
                cleanValue = CleanCellValue(CStr(cellValue))
                If IsNumeric(cleanValue) And cleanValue <> "" Then
                    If Not dataDict.exists(cavityId) Then
                        dataDict.Add cavityId, CDbl(cleanValue)
                        ' 調試信息：記錄提取的模穴數據
                        Debug.Print "提取模穴數據 - 檢驗項目: " & itemName & ", 模穴ID: " & cavityId & ", 數值: " & cleanValue
                    End If
                Else
                    ' 調試信息：記錄跳過的數據
                    Debug.Print "跳過無效數據 - 檢驗項目: " & itemName & ", 模穴ID: " & cavityId & ", 原始值: " & cellValue & ", 清理後: " & cleanValue
                End If
            Else
                ' 調試信息：記錄空的模穴ID
                Debug.Print "跳過空模穴ID - 檢驗項目: " & itemName & ", 欄位偏移: " & colOffset
            End If
        Next colOffset
        If dataDict.count > 0 And itemName <> "" Then
            Set rec = CreateObject("Scripting.Dictionary")
            rec.Add "inspectionItem", itemName
            rec.Add "data", dataDict
            result.Add rec
            ' 調試信息：記錄檢驗項目的模穴數量
            Debug.Print "檢驗項目 '" & itemName & "' 包含 " & dataDict.count & " 個模穴數據"
        Else
            ' 調試信息：記錄被跳過的檢驗項目
            Debug.Print "跳過檢驗項目 - 名稱: '" & itemName & "', 數據數量: " & dataDict.count
        End If
    Next rowOffset
    Set ExtractInspectionItemsFromGroup = result
End Function

Private Function AggregateInspectionItemsAcrossGroups(wbSource As Workbook, baseIndex As Long, params As UserInputParams) As Collection
    Dim agg As Object
    Dim result As Collection
    Dim groupIndex As Long
    Dim targetSheetIndex As Long
    Dim targetSheet As Worksheet
    Dim grpCfg As CavityGroupConfig
    Dim items As Collection
    Dim rec As Variant
    Dim itemName As Variant
    Dim dict As Object, src As Object, k As Variant
    Set agg = CreateObject("Scripting.Dictionary")
    For groupIndex = 1 To 6
        If gDataRange(groupIndex) <> "" And gCavityIdRange(groupIndex) <> "" Then
            targetSheetIndex = baseIndex + gPageOffset(groupIndex)
            If targetSheetIndex >= 1 And targetSheetIndex <= wbSource.Worksheets.count Then
                Set targetSheet = wbSource.Worksheets(targetSheetIndex)
                grpCfg.pageOffset = gPageOffset(groupIndex)
                grpCfg.dataRange = gDataRange(groupIndex)
                grpCfg.cavityIdRange = gCavityIdRange(groupIndex)
                Set items = ExtractInspectionItemsFromGroup(targetSheet, grpCfg)
                If Not items Is Nothing Then
                    For Each rec In items
                        itemName = CStr(rec("inspectionItem"))
                        Set src = rec("data")
                        If agg.exists(itemName) Then
                            Set dict = agg(itemName)
                        Else
                            Set dict = CreateObject("Scripting.Dictionary")
                            agg.Add itemName, dict
                        End If
                        For Each k In src.Keys
                            If Not dict.exists(k) Then
                                dict.Add k, src(k)
                                ' 調試信息：記錄合併的模穴數據
                                Debug.Print "合併模穴數據 - 檢驗項目: " & itemName & ", 模穴: " & k & ", 數值: " & src(k)
                            End If
                        Next k
                    Next rec
                End If
            End If
        End If
    Next groupIndex
    Set result = New Collection
    For Each itemName In agg.Keys
        Set rec = CreateObject("Scripting.Dictionary")
        rec.Add "inspectionItem", CStr(itemName)
        rec.Add "data", agg(itemName)
        result.Add rec
    Next itemName
    Set AggregateInspectionItemsAcrossGroups = result
End Function

Private Sub WriteDataRow(ws As Worksheet, row As Long, batchName As String, dataDict As Object, ByRef masterCavityList As Object)
    ' 寫入批號
    ws.Cells(row, 1).value = batchName
    
    ' 處理所有模穴數據
    Dim cavityKey As Variant, colIndex As Long, cavityNum As Long
    Dim processedCount As Long
    processedCount = 0
    
    ' 記錄處理過程以便調試
    Debug.Print "WriteDataRow - 批號: " & batchName & ", 數據字典大小: " & dataDict.count
    
    For Each cavityKey In dataDict.Keys
        If IsNumeric(cavityKey) Then
            cavityNum = CLng(cavityKey)
            colIndex = EnsureCavityColumn(ws, masterCavityList, cavityNum)
            ws.Cells(row, colIndex).value = dataDict(cavityKey)
            processedCount = processedCount + 1
            
            ' 調試信息
            Debug.Print "  處理模穴 " & cavityNum & " -> 欄位 " & colIndex & " = " & dataDict(cavityKey)
        Else
            ' 記錄非數字鍵值
            Debug.Print "  跳過非數字鍵值: " & cavityKey
        End If
    Next cavityKey
    
    Debug.Print "WriteDataRow - 完成，處理了 " & processedCount & " 個模穴"
    Debug.Print "WriteDataRow - 主模穴列表大小: " & masterCavityList.count
End Sub

Private Function EnsureCavityColumn(ws As Worksheet, ByRef masterCavityList As Object, cavityNum As Long) As Long
    Dim k As Variant, insertCol As Long, targetLabel As String
    
    ' 檢查模穴是否已經存在
    If masterCavityList.exists(cavityNum) Then
        EnsureCavityColumn = masterCavityList(cavityNum)
        ' 確保標題已設置（防止標題遺失）
        targetLabel = CStr(cavityNum) & "號穴"
        If ws.Cells(1, EnsureCavityColumn).value = "" Then
            ws.Cells(1, EnsureCavityColumn).value = targetLabel
            With ws.Cells(1, EnsureCavityColumn)
                .Font.Bold = True
                .Interior.color = RGB(146, 208, 80)
                .HorizontalAlignment = xlCenter
            End With
        End If
        Exit Function
    End If
    
    ' 尋找插入位置（保持模穴號碼順序）
    insertCol = 0
    For Each k In masterCavityList.Keys
        If CLng(k) > cavityNum Then
            If insertCol = 0 Or masterCavityList(k) < insertCol Then
                insertCol = masterCavityList(k)
            End If
        End If
    Next k
    
    ' 決定新欄位位置
    If insertCol = 0 Then
        ' 添加到最後（數據欄從第5欄開始，前4欄為規格資訊）
        EnsureCavityColumn = masterCavityList.count + 5
    Else
        ' 插入到指定位置
        ws.Columns(insertCol).Insert
        ' 更新所有受影響的欄位索引
        For Each k In masterCavityList.Keys
            If masterCavityList(k) >= insertCol Then
                masterCavityList(k) = masterCavityList(k) + 1
            End If
        Next k
        EnsureCavityColumn = insertCol
    End If
    
    ' 添加到主列表並設置標題
    masterCavityList.Add cavityNum, EnsureCavityColumn
    targetLabel = CStr(cavityNum) & "號穴"
    ws.Cells(1, EnsureCavityColumn).value = targetLabel
    With ws.Cells(1, EnsureCavityColumn)
        .Font.Bold = True
        .Interior.color = RGB(146, 208, 80)
        .HorizontalAlignment = xlCenter
    End With
End Function

' 將數據添加到對應的檢驗項目工作表（包含規格資訊）
Private Sub AddDataToInspectionSheet(wbOutput As Workbook, ByVal inspectionItem As String, ByVal batchName As String, ByVal dataDict As Object, ByRef sheetDataDict As Object, Optional sourceWb As Workbook = Nothing)
    Dim ws As Worksheet
    Dim sheetName As String
    Dim outputRow As Long
    Dim masterCavityList As Object
    
    ' 清理檢驗項目名稱，使其適合作為工作表名稱
    sheetName = CleanSheetName(inspectionItem)
    If sheetName = "" Then sheetName = "未命名項目"
    
    ' 檢查該檢驗項目的工作表是否已存在
    If Not sheetDataDict.exists(sheetName) Then
        ' 創建新工作表
        Set ws = wbOutput.Worksheets.Add(After:=wbOutput.Worksheets(wbOutput.Worksheets.count))
        ws.Name = sheetName
        
        ' 設置標題
        Call SetupInitialHeaders(ws)
        
        ' 創建該工作表的穴號列表
        Set masterCavityList = CreateObject("Scripting.Dictionary")
        
        ' 保存到字典
        Dim sheetInfo As Object
        Set sheetInfo = CreateObject("Scripting.Dictionary")
        sheetInfo.Add "worksheet", ws
        sheetInfo.Add "nextRow", 3  ' 數據從第3行開始（第1行標題，第2行空白，第3行開始是數據）
        sheetInfo.Add "cavityList", masterCavityList
        sheetDataDict.Add sheetName, sheetInfo
    Else
        ' 獲取現有工作表信息
        Set sheetInfo = sheetDataDict(sheetName)
        Set ws = sheetInfo("worksheet")
        Set masterCavityList = sheetInfo("cavityList")
    End If
    
    ' 寫入數據
    outputRow = sheetInfo("nextRow")
    Call WriteDataRow(ws, outputRow, batchName, dataDict, masterCavityList)
    
    ' 如果是第一次寫入數據，嘗試提取並設定規格資訊和產品資訊
    ' 第3行是第一筆數據（第1行標題，第2行空白）
    If outputRow = 3 And Not sheetInfo.exists("specificationSet") Then
        On Error Resume Next
        Call SetSpecificationData(ws, inspectionItem, sourceWb)
        ' 新增：提取產品名稱和測量單位
        Call SetProductInformation(ws, sourceWb)
        If Err.Number <> 0 Then
            ' 如果設定規格時發生錯誤，記錄但不中斷處理
            Debug.Print "設定規格數據時發生錯誤：" & Err.description
            Err.Clear
        End If
        On Error GoTo 0
        sheetInfo.Add "specificationSet", True
    End If
    
    ' 更新下一行位置
    sheetInfo("nextRow") = outputRow + 1
End Sub

' 清理工作表名稱（移除不允許的字符）
Private Function CleanSheetName(sheetName As String) As String
    Dim result As String
    Dim invalidChars As String
    Dim i As Long
    
    result = Trim(sheetName)
    invalidChars = "\/:*?""<>|"
    
    ' 移除不允許的字符
    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i
    
    ' 限制長度（Excel工作表名稱最多31個字符）
    If Len(result) > 31 Then
        result = Left(result, 31)
    End If
    
    CleanSheetName = result
End Function

' ============================================================================
' 設定規格資訊到工作表（自動提取版本）
' ============================================================================
Private Sub SetSpecificationData(ws As Worksheet, inspectionItem As String, Optional sourceWb As Workbook = Nothing)
    On Error GoTo ErrorHandler
    
    Dim specWs As Worksheet
    Dim spec As SpecificationData
    Dim autoExtracted As Boolean
    
    autoExtracted = False
    
    Debug.Print "=========================================="
    Debug.Print "開始設定規格 - 項目: " & inspectionItem
    
    ' 步驟1：嘗試從源工作簿自動提取規格
    If Not sourceWb Is Nothing Then
        Debug.Print "源工作簿: " & sourceWb.Name
        
        ' 尋找規格工作表
        Set specWs = FindSpecificationWorksheet(sourceWb)
        
        If Not specWs Is Nothing Then
            Debug.Print "找到規格工作表: " & specWs.Name
            
            ' 根據檢驗項目查找規格
            spec = FindSpecificationByItem(specWs, inspectionItem)
            
            If spec.IsValid Then
                ' 自動提取成功
                ' 規格數據寫入第2行
                ws.Cells(2, 2).value = spec.target
                ws.Cells(2, 3).value = spec.usl
                ws.Cells(2, 4).value = spec.lsl
                
                ' 設定成功的顏色（自動提取）
                With ws.Range("B2:D2")
                    .Interior.color = RGB(198, 224, 180) ' 淺綠色
                    .Font.color = RGB(0, 97, 0) ' 深綠色字體
                    .NumberFormat = "0.0000"
                End With
                
                autoExtracted = True
                Debug.Print "? 自動提取規格成功"
                Debug.Print "  目標值: " & spec.target
                Debug.Print "  USL: " & spec.usl
                Debug.Print "  LSL: " & spec.lsl
            Else
                Debug.Print "? 未找到有效的規格數據"
            End If
        Else
            Debug.Print "? 未找到規格工作表"
        End If
    Else
        Debug.Print "? 未提供源工作簿"
    End If
    
    ' 步驟2：如果自動提取失敗，標記為未設定
    If Not autoExtracted Then
        ws.Cells(2, 2).value = "未設定"
        ws.Cells(2, 3).value = "未設定"
        ws.Cells(2, 4).value = "未設定"
        
        ' 設定警告顏色
        With ws.Range("B2:D2")
            .Interior.color = RGB(255, 235, 156) ' 淺黃色警告
            .Font.color = RGB(156, 87, 0) ' 深橙色字體
        End With
        
        Debug.Print "? 無法自動提取規格，已標記為未設定"
    End If
    
    Debug.Print "=========================================="
    Exit Sub
    
ErrorHandler:
    Debug.Print "? 設定規格數據時發生錯誤：" & Err.description
    Debug.Print "=========================================="
    
    ' 設置錯誤狀態
    ws.Cells(2, 2).value = "錯誤"
    ws.Cells(2, 3).value = "錯誤"
    ws.Cells(2, 4).value = "錯誤"
    
    With ws.Range("B2:D2")
        .Interior.color = RGB(255, 199, 206) ' 淺紅色
        .Font.color = RGB(156, 0, 6) ' 深紅色字體
    End With
End Sub

' ============================================================================
' 提示用戶輸入規格資訊
' ============================================================================
' 暫時註釋掉自動提取功能，避免引起其他問題
' Private Sub AutoExtractSpecificationData(ws As Worksheet, inspectionItem As String, sourceWb As Workbook)
'     ' 此功能暫時停用，直到系統穩定後再重新啟用
' End Sub

Private Sub PromptForSpecificationInputManual(ws As Worksheet, inspectionItem As String)
    On Error GoTo ErrorHandler
    
    Dim target As String, usl As String, lsl As String
    Dim targetValue As Double, uslValue As Double, lslValue As Double
    
    ' 輸入目標值
    target = InputBox("請輸入 '" & inspectionItem & "' 的目標值：", "目標值", "0.245")
    If target = "" Then Exit Sub
    
    ' 輸入上規格限
    usl = InputBox("請輸入 '" & inspectionItem & "' 的上規格限：", "上規格限", "0.250")
    If usl = "" Then Exit Sub
    
    ' 輸入下規格限
    lsl = InputBox("請輸入 '" & inspectionItem & "' 的下規格限：", "下規格限", "0.240")
    If lsl = "" Then Exit Sub
    
    ' 驗證輸入的數值
    If IsNumeric(target) And IsNumeric(usl) And IsNumeric(lsl) Then
        targetValue = CDbl(target)
        uslValue = CDbl(usl)
        lslValue = CDbl(lsl)
        
        ' 驗證規格的合理性
        If uslValue > lslValue And targetValue >= lslValue And targetValue <= uslValue Then
            ' 設定規格資訊到工作表（第2行）
            ' 注意：此處不再檢查 Cells(3,1)
            ws.Cells(2, 2).value = targetValue
            ws.Cells(2, 3).value = uslValue
            ws.Cells(2, 4).value = lslValue
            
            ' 設定成功的顏色（手動輸入）
            With ws.Range("B2:D2")
                .Interior.color = RGB(255, 235, 156) ' 淺黃色
                .Font.color = RGB(156, 101, 0) ' 深黃色字體
            End With
        Else
            MsgBox "規格數據不合理！" & vbCrLf & _
                   "上規格限應大於下規格限，目標值應在規格範圍內。", _
                   vbCritical, "規格數據錯誤"
        End If
    Else
        MsgBox "請輸入有效的數值！", vbCritical, "輸入錯誤"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "設定規格數據時發生錯誤：" & Err.description, vbCritical, "錯誤"
End Sub

' ============================================================================


Private Sub LogError(wb As Workbook, ByRef wsLog As Worksheet, ByRef row As Long, fileName As String, sheetName As String, message As String)
    If wsLog Is Nothing Then
        Set wsLog = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        wsLog.Name = "處理異常紀錄"
        Call SetupErrorLogHeaders(wsLog)
    End If
    
    wsLog.Cells(row, 1).value = fileName
    wsLog.Cells(row, 2).value = sheetName
    wsLog.Cells(row, 3).value = message
    wsLog.Cells(row, 4).value = Now
    row = row + 1
End Sub

Private Sub SetupErrorLogHeaders(ws As Worksheet)
    With ws
        .Cells(1, 1).value = "檔案名稱"
        .Cells(1, 2).value = "工作表名稱"
        .Cells(1, 3).value = "異常訊息"
        .Cells(1, 4).value = "紀錄時間"
        
        With .Range("A1:D1")
            .Font.Bold = True
            .Interior.color = RGB(255, 199, 206)
            .Font.color = RGB(156, 0, 6)
        End With
        
        .Columns("A:A").ColumnWidth = 30
        .Columns("B:B").ColumnWidth = 20
        .Columns("C:C").ColumnWidth = 60
        .Columns("D:D").ColumnWidth = 20
    End With
End Sub

Function CleanCellValue(value As String) As String
    value = Trim(value)
    value = Replace(value, ",", "")
    value = Replace(value, "，", "")
    value = Replace(value, " ", "")
    value = Replace(value, Chr(160), "")
    value = Replace(value, "．", ".")
    If value = "-" Or value = "--" Then value = ""
    value = Replace(value, "–", "-")
    value = Replace(value, "—", "-")
    CleanCellValue = value
End Function

Function CleanFileName(fileName As String) As String
    Dim invalidChars As String, i As Long, result As String
    invalidChars = "\/:*?""<>|"
    result = fileName
    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i
    CleanFileName = result
End Function


Private Sub ExecuteProcessingCore(params As UserInputParams)
    Dim wbOutput As Workbook, startTime As Date
    Dim fso As Object, fileCount As Long
    Dim wsErrorLog As Worksheet, errorLogRow As Long
    Dim sheetDataDict As Object
    Dim ws As Worksheet
    Dim filePath As String
    
    Set sheetDataDict = CreateObject("Scripting.Dictionary")
    startTime = Now
    fileCount = 0
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    If Not g_SampleWorkbook Is Nothing Then
        On Error Resume Next
        g_SampleWorkbook.Close SaveChanges:=False
        Set g_SampleWorkbook = Nothing
        On Error GoTo 0
    End If
    
    Set wbOutput = Workbooks.Add
    
    Set wsErrorLog = Nothing
    errorLogRow = 2
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 驗證檔案存在
    filePath = params.folderPath  ' 在單檔案模式下,folderPath實際存儲的是檔案路徑
    If Not fso.FileExists(filePath) Then
        MsgBox "錯誤：找不到指定的檔案：" & vbCrLf & filePath, vbCritical
        GoTo CleanUp
    End If
    
    ' 處理單一檔案
    Call ProcessSingleFile(filePath, params, wbOutput, wsErrorLog, errorLogRow, sheetDataDict)
    fileCount = 1
    
    ' 為每個工作表添加統計列
    Dim sheetKey As Variant
    For Each sheetKey In sheetDataDict.Keys
        Dim sheetInfo As Object
        Set sheetInfo = sheetDataDict(sheetKey)
        Set ws = sheetInfo("worksheet")
        Dim lastRow As Long
        lastRow = sheetInfo("nextRow") - 1
        Dim cavityCount As Long
        cavityCount = sheetInfo("cavityList").count
        Call AddStatisticsColumns(ws, lastRow + 1, cavityCount)
    Next sheetKey
    
    ' 如果有創建新工作表，刪除預設的空白工作表
    If sheetDataDict.count > 0 Then
        Application.DisplayAlerts = False
        On Error Resume Next
        ' 找到並刪除名為 "Sheet1" 或 "工作表1" 的預設工作表
        Dim defaultSheet As Worksheet
        For Each defaultSheet In wbOutput.Worksheets
            If defaultSheet.Name = "Sheet1" Or defaultSheet.Name = "工作表1" Or _
               defaultSheet.Name = "Sheet" & wbOutput.Worksheets.count Then
                ' 檢查是否為空白工作表（沒有數據）
                If Application.WorksheetFunction.CountA(defaultSheet.UsedRange) = 0 Then
                    defaultSheet.Delete
                    Exit For
                End If
            End If
        Next defaultSheet
        On Error GoTo 0
        Application.DisplayAlerts = True
    End If
    
    ' 移動錯誤日誌工作表到最後
    If Not wsErrorLog Is Nothing Then wsErrorLog.Move After:=wbOutput.Worksheets(wbOutput.Worksheets.count)
    
    ' 激活第一個工作表
    If wbOutput.Worksheets.count > 0 Then wbOutput.Worksheets(1).Activate
    
CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    Dim totalTime As Double
    totalTime = (Now - startTime) * 24 * 60 * 60
    
    Dim finalMsg As String
    Dim sheetCount As Long
    If wbOutput Is Nothing Then
        sheetCount = 0
    Else
        sheetCount = wbOutput.Worksheets.count
        If Not wsErrorLog Is Nothing Then sheetCount = sheetCount - 1
    End If
    
    finalMsg = "資料處理完成！" & vbCrLf & vbCrLf & _
               "處理檔案數：" & fileCount & vbCrLf & _
               "生成工作表數：" & sheetCount & vbCrLf & _
               "處理時間：" & Format(totalTime, "0.0") & " 秒" & vbCrLf & vbCrLf
    
    If Not wsErrorLog Is Nothing Then
        finalMsg = finalMsg & "注意：有部分頁面格式不符或設定錯誤，詳情請見「處理異常紀錄」工作表。" & vbCrLf & vbCrLf
    End If
    
    finalMsg = finalMsg & "處理結果已按檢驗項目分類到不同工作表中，包含規格資訊。" & vbCrLf & _
               "建議保存到 ResultData 資料夾以便後續SPC分析使用。"
    
    MsgBox finalMsg, vbInformation, "處理完成"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "處理錯誤：" & Err.description, vbCritical, "錯誤"
End Sub

Private Sub SetupInitialHeaders(ws As Worksheet)
    ws.Columns("A:A").NumberFormat = "@"
    
    ' 設定標準化標題（包含規格資訊）
    ws.Cells(1, 1).value = "生產批號"
    ws.Cells(1, 2).value = "目標值"
    ws.Cells(1, 3).value = "上規格限"
    ws.Cells(1, 4).value = "下規格限"
    
    ' 設定標題格式
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.color = RGB(68, 114, 196)
        .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 設定欄寬
    ws.Columns("A:A").ColumnWidth = 20
    ws.Columns("B:D").ColumnWidth = 12
End Sub

Private Function GetPagesPerDataset(params As UserInputParams) As Long
    Dim i As Long, maxOffset As Long
    maxOffset = 0
    For i = 1 To 6
        If gDataRange(i) <> "" And gCavityIdRange(i) <> "" Then
            If gPageOffset(i) > maxOffset Then maxOffset = gPageOffset(i)
        End If
    Next i
    GetPagesPerDataset = maxOffset + 1
    If GetPagesPerDataset < 1 Then GetPagesPerDataset = 1
End Function

Private Sub ClearYellowInputs(params As UserInputParams)
    Dim ws As Worksheet, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("參數配置")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Range("rngSourceFile").Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngFileKeyword").Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngCavityCount").Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngSampleFilePath").Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngConfigName").Interior.ColorIndex = xlColorIndexNone
    For i = 1 To 6
        If gDataRange(i) <> "" And gCavityIdRange(i) <> "" Then
            ws.Range("rngCavityId_" & i).Interior.ColorIndex = xlColorIndexNone
            ws.Range("rngData_" & i).Interior.ColorIndex = xlColorIndexNone
            If i >= 2 Then ws.Range("rngOffset_" & i).Interior.ColorIndex = xlColorIndexNone
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub UpdateInputColors(ws As Worksheet)
    On Error Resume Next
    Call SetCellColorByValue(ws.Range("rngSourceFile"))
    Call SetCellColorByValue(ws.Range("rngFileKeyword"))
    Call SetCellColorByValue(ws.Range("rngCavityCount"))
    Call SetCellColorByValue(ws.Range("rngSampleFilePath"))
    Call SetCellColorByValue(ws.Range("rngConfigName"))
    Dim i As Long
    For i = 1 To 6
        Call SetCellColorByValue(ws.Range("rngCavityId_" & i))
        Call SetCellColorByValue(ws.Range("rngData_" & i))
        If i >= 2 Then Call SetCellColorByValue(ws.Range("rngOffset_" & i))
    Next i
    On Error GoTo 0
End Sub

Public Sub SetCellColorByValue(ByVal r As Range)
    If r Is Nothing Then Exit Sub
    If Trim(CStr(r.value)) = "" Then
        r.Interior.ColorIndex = xlColorIndexNone
    Else
        r.Interior.color = RGB(255, 255, 153)
    End If
End Sub

Private Sub ApplyInputColorRules(ws As Worksheet)
    Call UpdateInputColors(ws)
End Sub

Private Sub SetAllInputsYellow(ws As Worksheet)
    On Error Resume Next
    ws.Range("rngSourceFile").Interior.color = RGB(255, 255, 153)
    ws.Range("rngFileKeyword").Interior.color = RGB(255, 255, 153)
    ws.Range("rngCavityCount").Interior.color = RGB(255, 255, 153)
    ws.Range("rngSampleFilePath").Interior.color = RGB(255, 255, 153)
    ws.Range("rngConfigName").Interior.color = RGB(255, 255, 153)
    Dim i As Long
    For i = 1 To 6
        ws.Range("rngCavityId_" & i).Interior.color = RGB(255, 255, 153)
        ws.Range("rngData_" & i).Interior.color = RGB(255, 255, 153)
        If i >= 2 Then ws.Range("rngOffset_" & i).Interior.color = RGB(255, 255, 153)
    Next i
    On Error GoTo 0
End Sub

Private Sub SetAllInputsNoColor(ws As Worksheet)
    On Error Resume Next
    ws.Range("rngSourceFile").Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngFileKeyword").Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngCavityCount").Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngSampleFilePath").Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngConfigName").Interior.ColorIndex = xlColorIndexNone
    Dim j As Long
    For j = 1 To 6
        ws.Range("rngCavityId_" & j).Interior.ColorIndex = xlColorIndexNone
        ws.Range("rngData_" & j).Interior.ColorIndex = xlColorIndexNone
        If j >= 2 Then ws.Range("rngOffset_" & j).Interior.ColorIndex = xlColorIndexNone
    Next j
    On Error GoTo 0
End Sub

Private Sub ClearGroupInputs(ByVal groupIndex As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("參數配置")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Range("rngCavityId_" & groupIndex).Interior.ColorIndex = xlColorIndexNone
    ws.Range("rngData_" & groupIndex).Interior.ColorIndex = xlColorIndexNone
    If groupIndex >= 2 Then ws.Range("rngOffset_" & groupIndex).Interior.ColorIndex = xlColorIndexNone
    On Error GoTo 0
End Sub

Private Sub RestoreGroupInputs(params As UserInputParams)
    Dim ws As Worksheet, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("參數配置")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To 6
        If gDataRange(i) <> "" And gCavityIdRange(i) <> "" Then
            ws.Range("rngCavityId_" & i).Interior.color = RGB(255, 255, 153)
            ws.Range("rngData_" & i).Interior.color = RGB(255, 255, 153)
            If i >= 2 Then ws.Range("rngOffset_" & i).Interior.color = RGB(255, 255, 153)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub RestoreBasicInputs()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("參數配置")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Range("rngSourceFolder").Interior.color = RGB(255, 255, 153)
    ws.Range("rngFileKeyword").Interior.color = RGB(255, 255, 153)
    ws.Range("rngCavityCount").Interior.color = RGB(255, 255, 153)
    ws.Range("rngSampleFilePath").Interior.color = RGB(255, 255, 153)
    ws.Range("rngConfigName").Interior.color = RGB(255, 255, 153)
    On Error GoTo 0
End Sub

Private Sub ProcessSingleFile(ByVal filePath As String, params As UserInputParams, wbOutput As Workbook, ByRef wsErrorLog As Worksheet, ByRef errorLogRow As Long, ByRef sheetDataDict As Object)
    Dim wbSource As Workbook, wsBase As Worksheet, targetSheet As Worksheet
    Dim groupIndex As Long, dataFoundInSheet As Boolean
    Dim fileNameOnly As String
    Dim pagesPerDataset As Long, baseIndex As Long
    Dim items As Collection
    
    fileNameOnly = CreateObject("Scripting.FileSystemObject").GetFileName(filePath)
    
    On Error GoTo ErrorHandler
    
    Set wbSource = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
    If wbSource Is Nothing Then Exit Sub
    
    pagesPerDataset = GetPagesPerDataset(params)
    For baseIndex = 1 To wbSource.Worksheets.count Step pagesPerDataset
        Set wsBase = wbSource.Worksheets(baseIndex)
        dataFoundInSheet = False
        
        Set items = AggregateInspectionItemsAcrossGroups(wbSource, baseIndex, params)
        If Not items Is Nothing Then
            Dim rec As Variant
            For Each rec In items
                Call AddDataToInspectionSheet(wbOutput, CStr(rec("inspectionItem")), wsBase.Name, rec("data"), sheetDataDict, wbSource)
                dataFoundInSheet = True
            Next rec
        End If
        
        If Not dataFoundInSheet Then
            Call LogError(wbOutput, wsErrorLog, errorLogRow, fileNameOnly, wsBase.Name, "在此工作表及其關聯頁面中，根據有效的設定找不到任何數據。")
        End If
    Next baseIndex
    
ErrorHandler:
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
End Sub

Private Sub AddStatisticsColumns(ws As Worksheet, lastDataRow As Long, cavityColCount As Long)
    If lastDataRow <= 2 Then Exit Sub
    
    Dim statsColStart As Long, row As Long, dataRange As Range, cell As Range
    Dim actualLastCol As Long, col As Long
    
    ' 找到實際的最後一個模穴列（掃描標題行）
    actualLastCol = 1  ' 從批號列開始
    For col = 2 To 100
        If ws.Cells(1, col).value <> "" And InStr(ws.Cells(1, col).value, "號穴") > 0 Then
            actualLastCol = col
        ElseIf ws.Cells(1, col).value = "" And actualLastCol > 1 Then
            ' 遇到空白列且已找到模穴列，停止搜索
            Exit For
        End If
    Next col
    
    ' 統計列從最後一個模穴列之後開始
    statsColStart = actualLastCol + 1
    
    Debug.Print "AddStatisticsColumns - 模穴數量: " & cavityColCount & ", 實際最後列: " & actualLastCol & ", 統計列起始: " & statsColStart
    
    ws.Cells(1, statsColStart).value = "最大值"
    ws.Cells(1, statsColStart + 1).value = "最小值"
    ws.Cells(1, statsColStart + 2).value = "平均值"
    ws.Cells(1, statsColStart + 3).value = "標準差"
    
    With ws.Range(ws.Cells(1, statsColStart), ws.Cells(1, statsColStart + 3))
        .Font.Bold = True
        .Interior.color = RGB(237, 125, 49)
        .Font.color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    For row = 2 To lastDataRow - 1
        Set dataRange = Nothing
        
        ' 從第5列（第一個模穴列）到統計列之前
        For Each cell In ws.Range(ws.Cells(row, 5), ws.Cells(row, statsColStart - 1))
            If IsNumeric(cell.value) And cell.value <> "" Then
                If dataRange Is Nothing Then
                    Set dataRange = cell
                Else
                    Set dataRange = Application.Union(dataRange, cell)
                End If
            End If
        Next cell
        
        If Not dataRange Is Nothing Then
            If Application.WorksheetFunction.count(dataRange) > 0 Then
                ws.Cells(row, statsColStart).value = Application.WorksheetFunction.Max(dataRange)
                ws.Cells(row, statsColStart + 1).value = Application.WorksheetFunction.Min(dataRange)
                ws.Cells(row, statsColStart + 2).value = Application.WorksheetFunction.Average(dataRange)
                
                If Application.WorksheetFunction.count(dataRange) > 1 Then
                    ws.Cells(row, statsColStart + 3).value = Application.WorksheetFunction.StDev_S(dataRange)
                End If
            End If
        End If
    Next row
End Sub

' ============================================================================
' 新增：設定產品資訊（產品名稱和測量單位）
' ============================================================================
Private Sub SetProductInformation(ws As Worksheet, Optional sourceWb As Workbook = Nothing)
    On Error GoTo ErrorHandler
    
    Dim productName As String
    Dim measurementUnit As String
    
    ' 如果有源工作簿，嘗試提取產品資訊
    If Not sourceWb Is Nothing Then
        ' 嘗試從各個工作表中找到產品資訊
        Dim wsSource As Worksheet
        For Each wsSource In sourceWb.Worksheets
            On Error Resume Next
            
            Debug.Print "正在檢查工作表: " & wsSource.Name
            
            ' 提取產品名稱 (P2:V3 合併儲存格) - 更正範圍並嘗試多種方式讀取
            If productName = "" Then
                Dim tempProductName As String
                
                ' 方式1: 直接讀取 P2 儲存格
                tempProductName = Trim(CStr(wsSource.Cells(2, 16).value))  ' P2 = 第2行第16欄
                Debug.Print "P2 儲存格內容: '" & tempProductName & "'"
                
                ' 方式2: 如果P2沒有內容，嘗試讀取整個合併範圍 P2:V3
                If tempProductName = "" Or tempProductName = "0" Or tempProductName = "False" Then
                    tempProductName = Trim(CStr(wsSource.Range("P2:V3").value))
                    Debug.Print "P2:V3 範圍內容: '" & tempProductName & "'"
                End If
                
                ' 方式3: 如果還是沒有，嘗試讀取P3
                If tempProductName = "" Or tempProductName = "0" Or tempProductName = "False" Then
                    tempProductName = Trim(CStr(wsSource.Cells(3, 16).value))  ' P3
                    Debug.Print "P3 儲存格內容: '" & tempProductName & "'"
                End If
                
                ' 方式4: 嘗試P2:V3範圍內的其他位置
                If tempProductName = "" Or tempProductName = "0" Or tempProductName = "False" Then
                    Dim col As Integer, row As Integer
                    For row = 2 To 3  ' 第2行到第3行
                        For col = 16 To 22  ' P到V欄
                            tempProductName = Trim(CStr(wsSource.Cells(row, col).value))
                            If tempProductName <> "" And tempProductName <> "0" And tempProductName <> "False" Then
                                Debug.Print "在第" & row & "行第" & col & "欄找到產品名稱: '" & tempProductName & "'"
                                Exit For
                            End If
                        Next col
                        If tempProductName <> "" And tempProductName <> "0" And tempProductName <> "False" Then Exit For
                    Next row
                End If
                
                If tempProductName <> "" And tempProductName <> "0" And tempProductName <> "False" Then
                    productName = tempProductName
                    Debug.Print "從工作表 '" & wsSource.Name & "' 找到產品名稱: '" & productName & "'"
                End If
            End If
            
            ' 提取測量單位 (W23:X23 合併儲存格) - 使用合併儲存格的第一個儲存格
            If measurementUnit = "" Then
                Dim tempMeasurementUnit As String
                ' 對於合併儲存格，直接讀取第一個儲存格 W23
                tempMeasurementUnit = Trim(CStr(wsSource.Cells(23, 23).value))  ' W23 = 第23行第23欄
                Debug.Print "W23 儲存格內容: '" & tempMeasurementUnit & "'"
                If tempMeasurementUnit <> "" And tempMeasurementUnit <> "0" And tempMeasurementUnit <> "False" Then
                    measurementUnit = tempMeasurementUnit
                    Debug.Print "從工作表 '" & wsSource.Name & "' 找到測量單位: '" & measurementUnit & "'"
                End If
            End If
            
            On Error GoTo ErrorHandler
            
            ' 如果兩個都找到了就退出
            If productName <> "" And measurementUnit <> "" Then Exit For
        Next wsSource
    End If
    
    ' 將產品資訊寫入工作表的固定位置
    ' 使用第4-5行的B、C欄存儲產品資訊（避免與既有資訊衝突）
    ' 先解除工作表保護，寫入後再重新保護
    Dim wasProtected As Boolean
    Dim protectionPassword As String
    
    On Error Resume Next
    wasProtected = ws.ProtectContents
    If wasProtected Then
        ws.Unprotect  ' 嘗試解除保護（無密碼）
        If Err.Number <> 0 Then
            Debug.Print "無法解除工作表保護，嘗試寫入資料..."
            Err.Clear
        Else
            Debug.Print "已解除工作表保護"
        End If
    End If
    
    ' 寫入產品名稱資訊
    If productName <> "" Then
        ws.Cells(5, 2).value = "ProductName"  ' B5存儲標題
        ws.Cells(6, 2).value = productName    ' B6存儲產品名稱
        Debug.Print "寫入產品名稱 - B5: ProductName, B6: " & productName
    End If
    
    ' 寫入測量單位資訊（移除前綴文字"單位："）
    If measurementUnit <> "" Then
        Dim cleanUnit As String
        cleanUnit = measurementUnit
        ' 移除前綴文字"單位："
        If InStr(cleanUnit, "單位：") > 0 Then
            cleanUnit = Replace(cleanUnit, "單位：", "")
        End If
        cleanUnit = Trim(cleanUnit)  ' 移除前後空白
        
        ws.Cells(5, 3).value = "MeasurementUnit"  ' C5存儲標題
        ws.Cells(6, 3).value = cleanUnit          ' C6存儲清理後的測量單位
        Debug.Print "寫入測量單位 - C5: MeasurementUnit, C6: " & cleanUnit
        Debug.Print "原始測量單位: '" & measurementUnit & "' -> 清理後: '" & cleanUnit & "'"
    End If
    
    ' 如果原本有保護，重新保護工作表
    If wasProtected Then
        ws.Protect
        Debug.Print "已重新保護工作表"
    End If
    
    On Error GoTo ErrorHandler
    
    Debug.Print "產品資訊已設定 - 產品名稱: '" & productName & "', 測量單位: '" & measurementUnit & "'"
    Debug.Print "存儲位置 - B5: '" & ws.Cells(5, 2).value & "', C5: '" & ws.Cells(5, 3).value & "'"
    Debug.Print "源工作簿是否存在: " & (Not sourceWb Is Nothing)
    If Not sourceWb Is Nothing Then
        Debug.Print "源工作簿名稱: " & sourceWb.Name
        Debug.Print "源工作簿工作表數量: " & sourceWb.Worksheets.count
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "設定產品資訊時發生錯誤: " & Err.description
    Err.Clear
End Sub
