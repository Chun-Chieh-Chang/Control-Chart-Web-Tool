Option Explicit

' ============================================================================
' 整合批號分析模組 (修訂版 V2：批號為子組，模穴為樣本)
' ============================================================================

' ============================================================================
' 執行整合批號分析 (支持多工作表，每表最多25組數據)
' ============================================================================
Public Sub 執行批號分析()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim itemName As String
    Dim productCode As String
    
    On Error GoTo ErrorHandler
    
    ' 1. 選擇數據檔案
    Set wbSource = DataInput.SelectDataFile()
    If wbSource Is Nothing Then Exit Sub
    
    ' 2. 選擇檢驗項目 (這就是管制圖的標題)
    itemName = DataInput.SelectInspectionItem(wbSource)
    If itemName = "" Then
        wbSource.Close False
        Exit Sub
    End If
    
    ' 3. 設定來源工作表
    Set wsSource = wbSource.Worksheets(itemName)
    
    ' 4. 提取產品品號
    productCode = DataInput.ExtractProductCode(wbSource.FullName)
    
    ' 5. 執行多工作表分析與繪圖
    Call GenerateMultipleXBarRCharts(wsSource, itemName, productCode)
    
    ' 6. 關閉來源檔案
    wbSource.Close False
    
    ' 7. 優化檢視
    ActiveWindow.DisplayGridlines = False
    
    MsgBox "X-Bar R 管制圖生成完成！" & vbCrLf & "樣本分組依據：生產批號" & vbCrLf & "每個工作表最多包含25組數據", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not wbSource Is Nothing Then wbSource.Close False
    MsgBox "執行分析時發生錯誤：" & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================================
' 新增：生成多個 X-Bar R 管制圖工作表 (每表最多25組數據)
' ============================================================================
Public Sub GenerateMultipleXBarRCharts(wsSource As Worksheet, itemName As String, productCode As String)
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim batchCol As Long
    Dim cavityCols As Collection
    Dim headerVal As String
    
    ' 檢查參數有效性
    If wsSource Is Nothing Then
        MsgBox "錯誤：來源工作表無效。", vbExclamation
        Exit Sub
    End If
    
    ' 檢查是否可以訪問來源工作表
    On Error Resume Next
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.UsedRange.Columns.Count + wsSource.UsedRange.Column - 1
    On Error GoTo ErrorHandler
    
    ' 檢查是否有數據
    If lastRow < 2 Then
        MsgBox "錯誤：來源工作表沒有足夠的數據。", vbExclamation
        Exit Sub
    End If
    
    ' --- 步驟 1: 識別欄位 ---
    batchCol = 1
    
    ' 收集所有模穴欄位
    Set cavityCols = New Collection
    For c = 1 To lastCol
        headerVal = Trim(CStr(wsSource.Cells(1, c).Value))
        If InStr(headerVal, "穴") > 0 Then
            cavityCols.Add c
        End If
    Next c
    
    If cavityCols.Count < 2 Then
        MsgBox "錯誤：找不到足夠的模穴欄位 (標題需包含'穴')。", vbExclamation
        Exit Sub
    End If
    
    ' --- 步驟 2: 讀取所有批號數據 ---
    Dim startRowIndex As Long
    startRowIndex = 3  ' 從第3行開始（跳過標題行和規格行）
    
    Dim dataArr() As Double
    Dim batchNames() As String
    Dim n As Integer ' 樣本數 (模穴數)
    Dim k As Integer ' 總批號數
    
    n = cavityCols.Count
    k = lastRow - startRowIndex + 1
    
    If k <= 0 Or k > 1000 Then
        MsgBox "錯誤：批號數量不合理 (" & k & ")。", vbExclamation
        Exit Sub
    End If
    
    ReDim dataArr(1 To k, 1 To n)
    ReDim batchNames(1 To k)
    
    ' 讀取所有數據
    Dim batchIdx As Long
    Dim cavIdx As Long
    Dim colIndex As Variant
    
    batchIdx = 0
    For r = startRowIndex To lastRow
        batchIdx = batchIdx + 1
        batchNames(batchIdx) = CStr(wsSource.Cells(r, batchCol).Value)
        
        cavIdx = 0
        For Each colIndex In cavityCols
            cavIdx = cavIdx + 1
            If IsNumeric(wsSource.Cells(r, colIndex).Value) And wsSource.Cells(r, colIndex).Value <> "" Then
                dataArr(batchIdx, cavIdx) = CDbl(wsSource.Cells(r, colIndex).Value)
            Else
                dataArr(batchIdx, cavIdx) = -999999
            End If
        Next colIndex
    Next r
    
    ' --- 步驟 3: 讀取規格 ---
    Dim usl As Variant, lsl As Variant, target As Variant
    On Error Resume Next
    ' 使用 Text 屬性以保留來源格式 (如小數點位數)
    target = wsSource.Cells(2, 2).Text
    usl = wsSource.Cells(2, 3).Text
    lsl = wsSource.Cells(2, 4).Text
    
    ' 如果 Text 讀取失敗或顯示為 ####，則回退到 Value
    If Not IsNumeric(target) Then target = wsSource.Cells(2, 2).Value
    If Not IsNumeric(usl) Then usl = wsSource.Cells(2, 3).Value
    If Not IsNumeric(lsl) Then lsl = wsSource.Cells(2, 4).Value
    On Error GoTo ErrorHandler
    
    ' --- 步驟 3.1: 從第一檔案輸出的數據檔案中讀取產品資訊 ---
    Dim productName As String, measurementUnit As String
    productName = ""
    measurementUnit = ""
    
    On Error Resume Next
    ' 從數據檔案的固定位置讀取產品資訊
    Dim productNameTitle As String, measurementUnitTitle As String
    productNameTitle = CStr(wsSource.Cells(5, 2).Value)      ' B5存儲標題
    measurementUnitTitle = CStr(wsSource.Cells(5, 3).Value)  ' C5存儲標題
    
    ' 讀取實際的產品資訊，如果是空值則保持空值
    If productNameTitle = "ProductName" Then
        Dim tempProductName As Variant
        tempProductName = wsSource.Cells(6, 2).Value
        If IsEmpty(tempProductName) Or tempProductName = "" Then
            productName = ""
        Else
            productName = CStr(tempProductName)
        End If
    End If
    
    If measurementUnitTitle = "MeasurementUnit" Then
        Dim tempMeasurementUnit As Variant
        tempMeasurementUnit = wsSource.Cells(6, 3).Value
        If IsEmpty(tempMeasurementUnit) Or tempMeasurementUnit = "" Then
            measurementUnit = ""
        Else
            measurementUnit = CStr(tempMeasurementUnit)
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' --- 步驟 4: 計算需要的工作表數量 ---
    Dim maxBatchesPerSheet As Integer
    maxBatchesPerSheet = 25
    Dim totalSheets As Integer
    totalSheets = Application.WorksheetFunction.RoundUp(k / maxBatchesPerSheet, 0)
    
    ' --- 步驟 5: 創建多個工作表並填入數據 ---
    Dim sheetIndex As Integer
    Dim startBatch As Integer
    Dim endBatch As Integer
    Dim batchesInSheet As Integer
    Dim wsTarget As Worksheet
    Dim wsName As String
    
    For sheetIndex = 1 To totalSheets
        ' 計算當前工作表的批號範圍
        startBatch = (sheetIndex - 1) * maxBatchesPerSheet + 1
        endBatch = Application.WorksheetFunction.Min(sheetIndex * maxBatchesPerSheet, k)
        batchesInSheet = endBatch - startBatch + 1
        
        ' 創建新工作表
        Set wsTarget = ThisWorkbook.Worksheets.Add
        wsName = productCode & "-" & Format(sheetIndex, "000")
        DataInput.SetWorksheetName wsTarget, wsName
        wsTarget.Cells.Clear
        
        ' 提取當前工作表的數據
        Dim currentDataArr() As Double
        Dim currentBatchNames() As String
        ReDim currentDataArr(1 To batchesInSheet, 1 To n)
        ReDim currentBatchNames(1 To batchesInSheet)
        
        Dim i As Integer, j As Integer
        For i = 1 To batchesInSheet
            currentBatchNames(i) = batchNames(startBatch + i - 1)
            For j = 1 To n
                currentDataArr(i, j) = dataArr(startBatch + i - 1, j)
            Next j
        Next i
        
        ' 繪製當前工作表
        Call DrawLayoutAndFillData(wsTarget, currentDataArr, currentBatchNames, n, batchesInSheet, _
                                 productCode, itemName, usl, target, lsl, sheetIndex, totalSheets, productName, measurementUnit)
    Next sheetIndex
    
    ' 激活第一個工作表
    If totalSheets > 0 Then
        ThisWorkbook.Worksheets(productCode & "-001").Activate
    End If
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "在生成多工作表管制圖時發生錯誤: " & Err.Description, vbExclamation
    Resume Next
End Sub

' ============================================================================
' 核心邏輯：生成 X-Bar R 管制圖 (保留原函數以維持兼容性)
' ============================================================================
Public Sub GenerateXBarRChart(ws As Worksheet, wsSource As Worksheet, _
                              itemName As String, productCode As String)
    
    ' 啟用錯誤處理
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim batchCol As Long
    Dim cavityCols As Collection
    Dim headerVal As String
    
    ' 檢查參數有效性
    If ws Is Nothing Then
        MsgBox "錯誤：目標工作表無效。", vbExclamation
        Exit Sub
    End If
    
    If wsSource Is Nothing Then
        MsgBox "錯誤：來源工作表無效。", vbExclamation
        Exit Sub
    End If
    
    ' 檢查是否可以訪問來源工作表
    On Error Resume Next
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.UsedRange.Columns.Count + wsSource.UsedRange.Column - 1
    On Error GoTo ErrorHandler
    
    ' 檢查是否有數據
    If lastRow < 2 Then
        MsgBox "錯誤：來源工作表沒有足夠的數據。", vbExclamation
        Exit Sub
    End If
    
    ' --- 步驟 1: 識別欄位 ---
    ' 假設第一欄是批號 (通常是 "生產批號" 或 "Batch")
    batchCol = 1
    
    ' 收集所有模穴欄位
    Set cavityCols = New Collection
    For c = 1 To lastCol
        headerVal = Trim(CStr(wsSource.Cells(1, c).Value)) ' 假設標題在第1行
        ' 判斷邏輯：包含 "穴" 字的欄位視為樣本數據
        If InStr(headerVal, "穴") > 0 Then
            cavityCols.Add c
        End If
    Next c
    
    If cavityCols.Count < 2 Then
        MsgBox "錯誤：找不到足夠的模穴欄位 (標題需包含'穴')。", vbExclamation
        Exit Sub
    End If
    
    ' --- 步驟 2: 讀取所有批號數據 ---
    Dim startRowIndex As Long
    startRowIndex = 3 ' 從第3行開始（跳過標題行和規格行）
    
    ' 準備數據陣列
    ' dataArr(批次索引, 模穴索引)
    Dim dataArr() As Double
    Dim batchNames() As String
    Dim n As Integer ' 樣本數 (模穴數)
    Dim k As Integer ' 組數 (批號數)
    
    n = cavityCols.Count
    k = lastRow - startRowIndex + 1 ' 計算所有批號數量
    
    ' 檢查是否超出合理範圍
    If k <= 0 Or k > 1000 Then
        MsgBox "錯誤：批號數量不合理 (" & k & ")。", vbExclamation
        Exit Sub
    End If
    
    ReDim dataArr(1 To k, 1 To n)
    ReDim batchNames(1 To k)
    
    Dim batchIdx As Long
    Dim cavIdx As Long
    Dim colIndex As Variant
    
    batchIdx = 0
    For r = startRowIndex To lastRow
        batchIdx = batchIdx + 1
        batchNames(batchIdx) = CStr(wsSource.Cells(r, batchCol).Value)
        
        cavIdx = 0
        For Each colIndex In cavityCols
            cavIdx = cavIdx + 1
            ' 處理空值或非數值：保持原值以便後續判斷
            If IsNumeric(wsSource.Cells(r, colIndex).Value) And wsSource.Cells(r, colIndex).Value <> "" Then
                dataArr(batchIdx, cavIdx) = CDbl(wsSource.Cells(r, colIndex).Value)
            Else
                ' 空值或非數值設為特殊值（用於後續排除）
                dataArr(batchIdx, cavIdx) = -999999
            End If
        Next colIndex
    Next r
    
    ' --- 步驟 3: 讀取規格 ---
    Dim usl As Variant, lsl As Variant, target As Variant
    On Error Resume Next
    ' 使用 Text 屬性以保留來源格式 (如小數點位數)
    target = wsSource.Cells(2, 2).Text
    usl = wsSource.Cells(2, 3).Text
    lsl = wsSource.Cells(2, 4).Text
    
    ' 如果 Text 讀取失敗或顯示為 ####，則回退到 Value
    If Not IsNumeric(target) Then target = wsSource.Cells(2, 2).Value
    if Not IsNumeric(usl) Then usl = wsSource.Cells(2, 3).Value
    If Not IsNumeric(lsl) Then lsl = wsSource.Cells(2, 4).Value
    On Error GoTo ErrorHandler
    
    ' --- 步驟 3.1: 從第一檔案輸出的數據檔案中讀取產品資訊 ---
    Dim productName As String, measurementUnit As String
    productName = ""
    measurementUnit = ""
    
    On Error Resume Next
    ' 從數據檔案的固定位置讀取產品資訊
    Dim productNameTitle As String, measurementUnitTitle As String
    productNameTitle = CStr(wsSource.Cells(5, 2).Value)      ' B5存儲標題
    measurementUnitTitle = CStr(wsSource.Cells(5, 3).Value)  ' C5存儲標題
    
    ' 讀取實際的產品資訊，如果是空值則保持空值
    If productNameTitle = "ProductName" Then
        Dim tempProductName As Variant
        tempProductName = wsSource.Cells(6, 2).Value
        If IsEmpty(tempProductName) Or tempProductName = "" Then
            productName = ""
        Else                                       
            productName = CStr(tempProductName)
        End If
    End If
    
    If measurementUnitTitle = "MeasurementUnit" Then
        Dim tempMeasurementUnit As Variant
        tempMeasurementUnit = wsSource.Cells(6, 3).Value
        If IsEmpty(tempMeasurementUnit) Or tempMeasurementUnit = "" Then
            measurementUnit = ""
        Else
            measurementUnit = CStr(tempMeasurementUnit)
        End If
    End If
    
    
    
    On Error GoTo ErrorHandler
    
    ' --- 步驟 4: 繪製表單 ---
    Call DrawLayoutAndFillData(ws, dataArr, batchNames, n, k, _
                             productCode, itemName, usl, target, lsl, 1, 1, productName, measurementUnit)
                             
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    ' 忽略所有錯誤，讓程序繼續執行
    Application.ScreenUpdating = True
    MsgBox "在生成管制圖時發生錯誤: " & Err.Description, vbExclamation
    Resume Next
End Sub

' ============================================================================
' 繪製表單佈局並填入數據 (支持多工作表)
' ============================================================================
Private Sub DrawLayoutAndFillData(ws As Worksheet, dataArr() As Double, batchNames() As String, _
                                 n As Integer, k As Integer, _
                                 partNo As String, featureName As String, _
                                 usl As Variant, target As Variant, lsl As Variant, _
                                 Optional sheetIndex As Integer = 1, Optional totalSheets As Integer = 1, _
                                 Optional productName As String = "", Optional measurementUnit As String = "")
    
    ' 啟用錯誤處理
    On Error GoTo ErrorHandler
    
    ' 檢查參數
    If ws Is Nothing Then
        MsgBox "錯誤：目標工作表無效。", vbExclamation
        Exit Sub
    End If
    
    If n <= 0 Or k <= 0 Then
        MsgBox "錯誤：樣本數或批號數無效。", vbExclamation
        Exit Sub
    End If
    
    Dim i As Long, j As Long
    
    ' 設定欄寬
    On Error Resume Next
    ' 全局字型設定
    With ws.Cells.Font
        .Name = "Microsoft JhengHei"  ' 微軟正黑體
        .Size = 10
    End With
    
    ws.Columns("A:B").ColumnWidth = 4
    ws.Columns("C:C").ColumnWidth = 5
    ws.Columns("D:D").AutoFit ' 批號列自動調整寬度
    ws.Columns("AB:AB").ColumnWidth = 16 ' AB欄寬度設為16
    On Error GoTo ErrorHandler
    
    ' 取得 SPC 常數
    Dim A2 As Double, D3 As Double, D4 As Double
    Call GetSPCConstants(n, A2, D3, D4)
    
    ' ============================
    ' 1. 表頭 (Header)
    ' ============================
    ' 清除内容以避免合并单元格警告
    On Error Resume Next
    ws.Range("A1:AB1").ClearContents
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    With ws.Range("A1:AB1")
        .Merge
        .Value = ChrW(88) & ChrW(773) & " - R 管制圖" ' X-bar R
        .Font.Name = "Microsoft JhengHei" ' 微軟正黑體
        .Font.Size = 20
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 35
    End With
    On Error GoTo ErrorHandler
    
    Dim rngHeader As Range
    On Error Resume Next
    Set rngHeader = ws.Range("A2:AB5")
    DrawBorders rngHeader
    On Error GoTo ErrorHandler
    
    ' 產品資訊 - 如果是空值則填入空字串
    Dim displayProductName As String, displayMeasurementUnit As String
    displayProductName = IIf(productName = "", "", productName)
    displayMeasurementUnit = IIf(measurementUnit = "", "", measurementUnit)
    
    FillMergedCell ws, "A2:B2", "產品名稱", True: FillMergedCell ws, "C2:F2", displayProductName, False
    FillMergedCell ws, "A3:B3", "產品料號", True: FillMergedCell ws, "C3:F3", partNo, False
    FillMergedCell ws, "A4:B4", "測量單位", True: FillMergedCell ws, "C4:F4", displayMeasurementUnit, False
    FillMergedCell ws, "A5:B5", "管制特性", True: FillMergedCell ws, "C5:F5", "平均值/全距", False
    
    ' 規格
    FillMergedCell ws, "G2:H2", "規 格", True
    FillMergedCell ws, "I2:J2", "標準", False, True
    FillMergedCell ws, "K2:L2", "管制圖", False, True
    On Error Resume Next
    ws.Range("M2").Value = "X圖": ws.Range("N2").Value = "R圖"
    On Error GoTo ErrorHandler
    
    FillMergedCell ws, "G3:H3", "最大值", True: FillMergedCell ws, "I3:J3", usl, False
    FillMergedCell ws, "G4:H4", "目標值", True: FillMergedCell ws, "I4:J4", target, False
    FillMergedCell ws, "G5:H5", "最小值", True: FillMergedCell ws, "I5:J5", lsl, False
    
    ' 填入標籤 (上限/中心/下限)
    FillMergedCell ws, "K3:L3", "上 限", False
    FillMergedCell ws, "K4:L4", "中心值", False
    FillMergedCell ws, "K5:L5", "下 限", False
    
    ' 日期與批號資訊
    FillMergedCell ws, "O2:P2", "製造部門", True: FillMergedCell ws, "Q2:R2", "射出課", False
    FillMergedCell ws, "O3:P3", "檢驗單位", True: FillMergedCell ws, "Q3:R3", "品管組", False
    FillMergedCell ws, "O4:P5", "檢驗人員", True: FillMergedCell ws, "Q4:R5", "", False
    
    FillMergedCell ws, "S2:T5", "期限", False
    FillMergedCell ws, "U2:X5", "   月   日  到   月   日", False
    On Error Resume Next
    On Error GoTo ErrorHandler
    
    FillMergedCell ws, "Y2:AB2", "管制圖編號", True
    FillMergedCell ws, "Y3:AB5", ws.Name, False
    
    ' ============================
    ' 2. 數據表格 (Grid) - 動態調整高度
    ' ============================
    Dim startRow As Long: startRow = 6
    Dim sampleRows As Integer: sampleRows = n ' 根據模穴數決定行數
    
    ' 左側標題列 - 移除日期行，只保留檢驗日期
    On Error Resume Next
    FillMergedCell ws, ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2)).Address, "檢驗日期", True
    On Error GoTo ErrorHandler
    
    ' 樣本測定值標題 (垂直合併)
    ' 清除内容以避免合并单元格警告
    On Error Resume Next
    ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + sampleRows, 1)).ClearContents
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    With ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + sampleRows, 1))
        .Merge
        .Value = "樣本測定值"
        .Orientation = 90
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = RGB(0, 0, 255)
        .Font.Bold = True
    End With
    On Error GoTo ErrorHandler
    
    ' 樣本編號 (X1, X2... Xn)
    For i = 1 To sampleRows
        If startRow + i <= ws.Rows.Count And 2 <= ws.Columns.Count Then
            With ws.Cells(startRow + i, 2)
                .Value = "X" & i
                .Font.Color = RGB(0, 0, 255)
                .HorizontalAlignment = xlCenter
            End With
        End If
    Next i
    
    ' 統計列位置 - 因為移除了一行，所以位置要調整
    Dim sumRow As Long: sumRow = startRow + 1 + sampleRows
    Dim xBarRow As Long: xBarRow = sumRow + 1
    Dim rangeRow As Long: rangeRow = sumRow + 2
    Dim causeRow As Long: causeRow = rangeRow + 1
    
    ' 統計列標題
    On Error Resume Next
    FillMergedCell ws, ws.Range(ws.Cells(sumRow, 1), ws.Cells(sumRow, 2)).Address, "ΣX", False, True
    FillMergedCell ws, ws.Range(ws.Cells(xBarRow, 1), ws.Cells(xBarRow, 2)).Address, ChrW(88) & ChrW(773), False, True ' X-bar
    FillMergedCell ws, ws.Range(ws.Cells(rangeRow, 1), ws.Cells(rangeRow, 2)).Address, "R", False, True
    On Error GoTo ErrorHandler
    
    ' 注意：在統計列添加合計值的代碼移到數據計算完成後
    
    ' 原因追查
    ' 清除内容以避免合并单元格警告
    On Error Resume Next
    ws.Range(ws.Cells(causeRow, 1), ws.Cells(causeRow + 4, 2)).ClearContents
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    With ws.Range(ws.Cells(causeRow, 1), ws.Cells(causeRow + 4, 2))
        .Merge
        .Value = "原因追查"
        .Orientation = 90
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    On Error GoTo ErrorHandler
    
    ' 填入期限欄位 - 使用第一個和最後一個生產批號
    Dim firstBatch As String, lastBatch As String
    If k > 0 Then
        firstBatch = batchNames(1)
        lastBatch = batchNames(k)
        FillMergedCell ws, "U2:X2", firstBatch & " 到 " & lastBatch, False
    End If
    
    ' 畫格子線 - 固定到AB欄（第28欄）
    Dim borderEndColumn As Long
    borderEndColumn = 28  ' AB欄是第28欄，固定表格寬度
    
    On Error Resume Next
    DrawBorders ws.Range(ws.Cells(startRow, 1), ws.Cells(causeRow + 4, borderEndColumn))
    On Error GoTo ErrorHandler
    
    ' ============================
    ' 3. 填入數據與計算
    ' ============================
    Dim batchX() As Double, batchR() As Double
    ' 修正：陣列大小應該是25，而不是k
    ReDim batchX(1 To 25)
    ReDim batchR(1 To 25)
    
    Dim sumX As Double, maxVal As Double, minVal As Double
    Dim grandTotalX As Double, totalR As Double
    
    ' 固定顯示25個欄位，從第3欄(C)開始到第27欄(AA)
    Dim maxDisplayColumns As Integer
    maxDisplayColumns = 25
    
    For i = 1 To maxDisplayColumns ' 固定25個欄位
        sumX = 0
        maxVal = -9999999
        minVal = 9999999
        
        ' 只有在有實際數據時才處理
        If i <= k Then
        
        ' 填入檢驗日期 (使用文字格式避免科學記號)
            On Error Resume Next
            If ws.Cells(startRow, 2 + i).Address <> "" Then
                ws.Cells(startRow, 2 + i).NumberFormat = "@"  ' 設定為文字格式
                ws.Cells(startRow, 2 + i).Value = batchNames(i)
            End If
            On Error GoTo ErrorHandler
            
            ' 填入樣本數據並排除空值和零值
            Dim validCount As Integer
            validCount = 0
            
            For j = 1 To n ' 模穴迴圈 (Row)
                Dim val As Double
                val = dataArr(i, j)
                
                ' 填入數據到對應位置 - 因為移除了一行，位置要調整
                On Error Resume Next
                If ws.Cells(startRow + j, 2 + i).Address <> "" Then
                    If val = -999999 Then
                        ws.Cells(startRow + j, 2 + i).Value = ""
                    Else
                        ws.Cells(startRow + j, 2 + i).Value = val ' 填入原始數值
                        ws.Cells(startRow + j, 2 + i).NumberFormat = "0.0000" ' 設定格式
                    End If
                End If
                On Error GoTo ErrorHandler
                
                ' 只有非零、非空值且有效的數值才納入計算
                If val <> 0 And val <> -999999 And IsNumeric(val) Then
                    If validCount = 0 Then
                        ' 第一個有效值，初始化 max 和 min
                        maxVal = val
                        minVal = val
                    Else
                        ' 後續有效值，比較 max 和 min
                        If val > maxVal Then maxVal = val
                        If val < minVal Then minVal = val
                    End If
                    sumX = sumX + val
                    validCount = validCount + 1
                End If
            Next j
            
            ' 計算單批統計（只使用有效數據）
            Dim avg As Double, rng As Double
            If validCount > 0 Then
                avg = sumX / validCount
                rng = maxVal - minVal
            Else
                ' 如果沒有有效數據，設為 0
                avg = 0
                rng = 0
            End If
            
            batchX(i) = avg
            batchR(i) = rng
            
            ' 填入統計 - 只有當有有效數據時才填入和累加
            On Error Resume Next
            If validCount > 0 Then
                ' 有有效數據，填入統計值
                If ws.Cells(sumRow, 2 + i).Address <> "" Then
                    ws.Cells(sumRow, 2 + i).Value = sumX
                    ws.Cells(sumRow, 2 + i).NumberFormat = "0.0000"
                End If
                
                If ws.Cells(xBarRow, 2 + i).Address <> "" Then
                    ws.Cells(xBarRow, 2 + i).Value = avg
                    ws.Cells(xBarRow, 2 + i).NumberFormat = "0.0000"
                End If
                
                If ws.Cells(rangeRow, 2 + i).Address <> "" Then
                    ws.Cells(rangeRow, 2 + i).Value = rng
                    ws.Cells(rangeRow, 2 + i).NumberFormat = "0.0000"
                End If
                
                ' 累加到總計（只有有效數據才累加）
                grandTotalX = grandTotalX + sumX
                totalR = totalR + rng
            Else
                ' 沒有有效數據，填入空值或0
                If ws.Cells(sumRow, 2 + i).Address <> "" Then
                    ws.Cells(sumRow, 2 + i).Value = ""  ' 空值
                End If
                
                If ws.Cells(xBarRow, 2 + i).Address <> "" Then
                    ws.Cells(xBarRow, 2 + i).Value = ""  ' 空值
                End If
                
                If ws.Cells(rangeRow, 2 + i).Address <> "" Then
                    ws.Cells(rangeRow, 2 + i).Value = ""  ' 空值
                End If
                
                ' 不累加到總計
            End If
            On Error GoTo ErrorHandler
        Else
            ' 沒有數據的欄位，設置空的統計值
            batchX(i) = 0
            batchR(i) = 0
        End If
    Next i
    
    ' 全局統計 - 只計算有有效數據的批次
    Dim grandAvg As Double, avgR As Double
    Dim validBatchCount As Integer
    Dim sumOfAvg As Double
    validBatchCount = 0
    sumOfAvg = 0
    
    ' 計算有效批次數量並累加子組平均值
    For i = 1 To k
        If batchX(i) <> 0 Or batchR(i) <> 0 Then
            validBatchCount = validBatchCount + 1
            sumOfAvg = sumOfAvg + batchX(i)  ' 累加每個子組的平均值
        End If
    Next i
    
    If validBatchCount > 0 Then
        ' 正確的 X̿ 計算：所有子組平均值(X̅)的總和 ÷ 子組個數
        grandAvg = sumOfAvg / validBatchCount
        avgR = totalR / validBatchCount
    Else
        ' 沒有有效數據，設為0
        grandAvg = 0
        avgR = 0
    End If
    
    ' 計算管制界限 - 只有當有有效數據時才計算
    Dim uclX As Double, lclX As Double
    Dim uclR As Double, lclR As Double
    
    If validBatchCount > 0 Then
        uclX = grandAvg + (A2 * avgR)
        lclX = grandAvg - (A2 * avgR)
        uclR = D4 * avgR
        lclR = D3 * avgR
    Else
        ' 沒有有效數據，管制界限設為0
        uclX = 0
        lclX = 0
        uclR = 0
        lclR = 0
    End If
    
    ' 填入界限值到表頭
    On Error Resume Next
    If ws.Range("M3").Address <> "" Then
        ws.Range("M3").Value = Round(uclX, 4)
    End If
    
    If ws.Range("M4").Address <> "" Then
        ws.Range("M4").Value = Round(grandAvg, 4)
    End If
    
    If ws.Range("M5").Address <> "" Then
        ws.Range("M5").Value = Round(lclX, 4)
    End If
    
    If ws.Range("N3").Address <> "" Then
        ws.Range("N3").Value = Round(uclR, 4)
    End If
    
    If ws.Range("N4").Address <> "" Then
        ws.Range("N4").Value = Round(avgR, 4)
    End If
    
    If ws.Range("N5").Address <> "" Then
        ' R 圖不顯示 LCL，所以這裡留空或顯示 "-"
        ws.Range("N5").Value = "-"
    End If
    On Error GoTo ErrorHandler
    
    ' 檢查並標記超出管制界限的數據欄位（只檢查有實際數據的欄位）
    For i = 1 To k
        If i <= maxDisplayColumns Then ' 只檢查顯示範圍內的欄位
            ' 檢查 X̄ 是否超出管制界限
            If xBarRow <= ws.Rows.Count And 2 + i <= ws.Columns.Count Then
                If ws.Cells(xBarRow, 2 + i).Address <> "" Then
                    Dim xBarValue As Double
                    xBarValue = ws.Cells(xBarRow, 2 + i).Value
                    If xBarValue > uclX Or xBarValue < lclX Then
                        ' 超出管制界限，設定黃色背景
                        ws.Cells(xBarRow, 2 + i).Interior.Color = RGB(255, 255, 0)
                    End If
                End If
            End If
            
            ' 檢查 R 是否超出管制界限
            If rangeRow <= ws.Rows.Count And 2 + i <= ws.Columns.Count Then
                If ws.Cells(rangeRow, 2 + i).Address <> "" Then
                    Dim rValue As Double
                    rValue = ws.Cells(rangeRow, 2 + i).Value
                    ' R 圖只檢查上管制限，不檢查下管制限
                    If rValue > uclR Then
                        ' 超出管制界限，設定黃色背景
                        ws.Cells(rangeRow, 2 + i).Interior.Color = RGB(255, 255, 0)
                    End If
                End If
            End If
        End If
    Next i
    
    ' 在統計列的最後一欄添加合計值（移到這裡，確保變數已定義）
    Dim totalColumn As Long
    totalColumn = 2 + k + 1
    If totalColumn <= 16384 Then
        ' 清除統計列的合計顯示，改為在表格右下角創建簡單的合計區域
        ' 不在統計列顯示任何合計數值
        On Error GoTo ErrorHandler
    End If
    
    ' 在AB欄顯示統計數據
    Dim summaryCol As Long
    summaryCol = 28     ' AB欄是第28欄
    
    On Error Resume Next
    
    ' 在 ΣX 統計列的AB欄顯示 ΣX̅= 數值 (固定在 AB7:AB8)
    With ws.Range(ws.Cells(7, summaryCol), ws.Cells(8, summaryCol))
        .Merge
        .ClearContents
        .Interior.Color = RGB(220, 230, 241)
        ' Type 1: ΣX bar
        AddStatisticalLabel ws, ws.Range(ws.Cells(7, summaryCol), ws.Cells(8, summaryCol)), 1, Round(sumOfAvg, 4)
    End With
    
    ' 在 X̄ 統計列的AB欄顯示 X̿= 數值 (固定在 AB9:AB10)
    With ws.Range(ws.Cells(9, summaryCol), ws.Cells(10, summaryCol))
        .Merge
        .ClearContents
        .Interior.Color = RGB(220, 230, 241)
        ' Type 2: X double bar
        AddStatisticalLabel ws, ws.Range(ws.Cells(9, summaryCol), ws.Cells(10, summaryCol)), 2, Round(grandAvg, 4)
    End With
    
    ' 在 R 統計列的AB欄顯示 ΣR= 數值 (固定在 AB11:AB12)
    With ws.Range(ws.Cells(11, summaryCol), ws.Cells(12, summaryCol))
        .Merge
        .ClearContents
        .Interior.Color = RGB(220, 230, 241)
        ' Type 3: ΣR
        AddStatisticalLabel ws, ws.Range(ws.Cells(11, summaryCol), ws.Cells(12, summaryCol)), 3, Round(totalR, 4)
    End With
    
    ' 在原因追查行的AB欄顯示 R̅= 數值 (固定在 AB13:AB14)
    With ws.Range(ws.Cells(13, summaryCol), ws.Cells(14, summaryCol))
        .Merge
        .ClearContents
        .Interior.Color = RGB(220, 230, 241)
        ' Type 4: R bar
        AddStatisticalLabel ws, ws.Range(ws.Cells(13, summaryCol), ws.Cells(14, summaryCol)), 4, Round(avgR, 4)
    End With
    
    On Error GoTo ErrorHandler
    
    ' ============================
    ' 4. 繪製圖表 (Charts)
    ' ============================
    Dim chartTop As Double, chartLeft As Double, chartW As Double, chartH As Double
    
    ' 圖表位置設定在數據表格下方，固定寬度到AB欄
    On Error Resume Next
    chartTop = ws.Cells(causeRow + 5, 1).Top  ' 圖表放在數據表格下方
    chartLeft = ws.Cells(causeRow + 5, 1).Left
    
    ' 固定圖表寬度到AC欄的範圍（比AB欄再多一個欄位）
    Dim maxChartWidth As Double
    maxChartWidth = ws.Range("A1:AC1").Width
    chartW = maxChartWidth - 50  ' 留一些邊距
    
    chartH = 250
    On Error GoTo ErrorHandler    
    ' X-Bar Chart - 傳入實際數據數量，但不超過25
    Dim actualDataCount As Integer
    actualDataCount = Application.WorksheetFunction.Min(k, 25)
    
    On Error Resume Next
    Call DrawLineChart(ws, chartLeft, chartTop, chartW, chartH, _
                      batchX, batchNames, actualDataCount, uclX, grandAvg, lclX, CDbl(VBA.Val(usl)), CDbl(VBA.Val(lsl)), "X-Bar 管制圖", 1)
    On Error GoTo ErrorHandler
                      
    ' R Chart - 與第一張圖表緊鄰，減少間距
    On Error Resume Next
    Call DrawLineChart(ws, chartLeft, chartTop + chartH + 5, chartW, chartH, _
                      batchR, batchNames, actualDataCount, uclR, avgR, lclR, 0, 0, "R 管制圖", 2)
    On Error GoTo ErrorHandler
    
    ' 在第二張圖表下方加入表單編號和頁數資訊
    Dim formNumberRow As Long
    ' 動態計算第二張圖表的底部位置
    ' 第二張圖表起始位置：causeRow + 5 + chartH/15（第一張圖表高度轉換為行數）+ 1（間距）
    ' 第二張圖表結束位置：再加上 chartH/15（第二張圖表高度轉換為行數）
    ' 表單資訊位置：圖表結束位置 + 1行
    formNumberRow = causeRow + 5 + Int(chartH / 15) + 1 + Int(chartH / 15) + 1
    
    On Error Resume Next
    ' 左側：表單編號
    ws.Cells(formNumberRow, 1).Value = "MOULDEX QE20002-R01.A"
    ws.Cells(formNumberRow, 1).Font.Size = 10
    ws.Cells(formNumberRow, 1).Font.Bold = True
    ws.Cells(formNumberRow, 1).HorizontalAlignment = xlLeft
    
    ' 右側：頁數資訊（在AB欄）
    ws.Cells(formNumberRow, 28).Value = "第 " & sheetIndex & " 頁，共 " & totalSheets & " 頁"
    ws.Cells(formNumberRow, 28).Font.Size = 10
    ws.Cells(formNumberRow, 28).Font.Bold = True
    ws.Cells(formNumberRow, 28).HorizontalAlignment = xlRight
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "在生成管制圖時發生錯誤: " & Err.Description, vbExclamation
    Resume Next
End Sub

' ============================================================================
' 輔助：填入合併儲存格並設定樣式
' ============================================================================
Private Sub FillMergedCell(ws As Worksheet, rngAddr As String, val As Variant, _
                          Optional isBlue As Boolean = False, Optional isBold As Boolean = False)
    On Error GoTo ErrorHandler
    
    ' 先清除范围内的所有内容，避免合并单元格警告
    ws.Range(rngAddr).ClearContents
    
    With ws.Range(rngAddr)
        .Merge
        .Value = val
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        If isBlue Then
            .Font.Color = RGB(0, 0, 255)
            .Font.Bold = True
        End If
        If isBold Then .Font.Bold = True
    End With
    
    Exit Sub
    
ErrorHandler:
    ' 忽略所有錯誤，讓程序繼續執行
    Resume Next
End Sub

' ============================================================================
' 輔助：繪製邊框
' ============================================================================
Private Sub DrawBorders(rng As Range)
    On Error GoTo ErrorHandler
    
    ' 內部框線 - 淺灰色細線
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With

    ' 外框 - 黑色中粗線
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    
    Exit Sub
    
ErrorHandler:
    ' 忽略所有錯誤，讓程序繼續執行
    Resume Next
End Sub

' ============================================================================
' 輔助：繪製折線圖
' ============================================================================
Private Sub DrawLineChart(ws As Worksheet, l As Double, t As Double, w As Double, h As Double, _
                         dataArr() As Double, batchNames() As String, count As Integer, _
                         ucl As Double, cl As Double, lcl As Double, _
                         usl As Double, lsl As Double, title As String, typeIdx As Integer)
    
    ' 啟用錯誤處理
    On Error GoTo ErrorHandler
    
    ' 檢查參數有效性
    If count <= 0 Or ws Is Nothing Then Exit Sub
    
    Dim chartObj As ChartObject
    On Error Resume Next
    Set chartObj = ws.ChartObjects.Add(l, t, w, h)
    On Error GoTo ErrorHandler
    
    ' 檢查圖表物件是否成功建立
    If chartObj Is Nothing Then Exit Sub
    
    ' 數據暫存區 (放在很遠的地方: AZ列 + 偏移)
    Dim dataRow As Long
    dataRow = 100 + (typeIdx * 50)
    
    ' 檢查是否超出 Excel 的限制
    On Error Resume Next
    Dim i As Long
    For i = 1 To count
        ' 確保不超出 Excel 最大行列限制和陣列範圍
        If dataRow + 5 <= ws.Rows.Count And 52 + i <= ws.Columns.Count And i <= UBound(dataArr) Then
            If ws.Cells(dataRow, 52 + i).Address <> "" Then
                ws.Cells(dataRow, 52 + i).Value = dataArr(i)
            End If
            
            If ws.Cells(dataRow + 1, 52 + i).Address <> "" Then
                ws.Cells(dataRow + 1, 52 + i).Value = ucl
            End If
            
            If ws.Cells(dataRow + 2, 52 + i).Address <> "" Then
                ws.Cells(dataRow + 2, 52 + i).Value = cl
            End If
            
            If ws.Cells(dataRow + 3, 52 + i).Address <> "" Then
                ws.Cells(dataRow + 3, 52 + i).Value = lcl
            End If
            
            If ws.Cells(dataRow + 4, 52 + i).Address <> "" Then
                ws.Cells(dataRow + 4, 52 + i).Value = usl  ' USL
            End If
            
            If ws.Cells(dataRow + 5, 52 + i).Address <> "" Then
                ws.Cells(dataRow + 5, 52 + i).Value = lsl  ' LSL
            End If
        End If
    Next i
    On Error GoTo ErrorHandler
    
    ' 檢查數據是否正確寫入
    If Not ws.Cells(dataRow, 53).Value = dataArr(1) Then
        ' 可能發生錯誤，但我們繼續執行
    End If
    
    With chartObj.Chart
        On Error Resume Next
        .ChartArea.Font.Name = "Microsoft JhengHei"
        .ChartArea.Font.Size = 10
        .ChartType = xlLine
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        On Error GoTo ErrorHandler
        
        ' 檢查數據範圍是否有效
        Dim dataStartCell As Range, dataEndCell As Range
        Dim isValidRange As Boolean
        isValidRange = True
        
        ' 檢查數據範圍
        On Error Resume Next
        Set dataStartCell = ws.Cells(dataRow, 53)
        Set dataEndCell = ws.Cells(dataRow, 52 + count)
        If dataStartCell Is Nothing Or dataEndCell Is Nothing Then
            isValidRange = False
        End If
        On Error GoTo ErrorHandler
        
        ' 再次檢查範圍是否有效
        If Not (dataRow + 5 <= ws.Rows.Count And 52 + count <= ws.Columns.Count) Then
            isValidRange = False
        End If
        
        If isValidRange And count > 0 Then
            ' 根據圖表類型決定是否顯示規格線
            Dim showSpecLimits As Boolean
            ' showSpecLimits = (usl <> 0) Or (lsl <> 0)
            ' 修正：取消在 X-Bar 圖上顯示規格界限 (避免與管制界限混淆)
            showSpecLimits = False
            
            ' 定義美化配色
            Dim colorData As Long, colorCL As Long, colorLimit As Long, colorAlert As Long
            colorData = RGB(31, 119, 180)   ' 專業藍 (Data)
            colorCL = RGB(44, 160, 44)      ' 綠色 (Center Line)
            colorLimit = RGB(214, 39, 40)   ' 紅色 (Control Limits)
            colorAlert = RGB(255, 0, 0)     ' 亮紅 (Alert)
            
            If showSpecLimits Then
                ' USL (Red Dash) - 規格上限
                On Error Resume Next
                With .SeriesCollection.NewSeries
                    If ws.Cells(dataRow + 4, 52 + count).Address <> "" And ws.Cells(dataRow + 4, 53).Address <> "" Then
                        .Values = ws.Range(ws.Cells(dataRow + 4, 53), ws.Cells(dataRow + 4, 52 + count))
                    End If
                    .Name = "USL"
                    .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                    .Format.Line.DashStyle = msoLineDash
                    .Format.Line.Weight = 1.25
                End With
                On Error GoTo ErrorHandler
                
                ' LSL (Orange Dash) - 規格下限，使用橙色區分
                On Error Resume Next
                With .SeriesCollection.NewSeries
                    If ws.Cells(dataRow + 5, 52 + count).Address <> "" And ws.Cells(dataRow + 5, 53).Address <> "" Then
                        .Values = ws.Range(ws.Cells(dataRow + 5, 53), ws.Cells(dataRow + 5, 52 + count))
                    End If
                    .Name = "LSL"
                    .Format.Line.ForeColor.RGB = RGB(255, 165, 0)  ' 橙色
                    .Format.Line.DashStyle = msoLineDash
                    .Format.Line.Weight = 1.25
                End With
                On Error GoTo ErrorHandler
            End If
            
            ' UCL (Blue Dash)
            On Error Resume Next
            With .SeriesCollection.NewSeries
                If ws.Cells(dataRow + 1, 52 + count).Address <> "" And ws.Cells(dataRow + 1, 53).Address <> "" Then
                    .Values = ws.Range(ws.Cells(dataRow + 1, 53), ws.Cells(dataRow + 1, 52 + count))
                End If
                .Name = "UCL"
                .Format.Line.ForeColor.RGB = colorLimit
                .Format.Line.DashStyle = msoLineDash
                .Format.Line.Weight = 1.25
            End With
            On Error GoTo ErrorHandler
            
            ' CL (Blue Solid)
            On Error Resume Next
            With .SeriesCollection.NewSeries
                If ws.Cells(dataRow + 2, 52 + count).Address <> "" And ws.Cells(dataRow + 2, 53).Address <> "" Then
                    .Values = ws.Range(ws.Cells(dataRow + 2, 53), ws.Cells(dataRow + 2, 52 + count))
                End If
                .Name = "CL"
                .Format.Line.ForeColor.RGB = colorCL
                .Format.Line.Weight = 1.5
            End With
            On Error GoTo ErrorHandler
            
            ' LCL (Purple Dash) - 僅在 X-Bar 圖顯示，R 圖不顯示，使用紫色區分
            If typeIdx = 1 Then ' 只有 X-Bar 圖才顯示 LCL
                On Error Resume Next
                With .SeriesCollection.NewSeries
                    If ws.Cells(dataRow + 3, 52 + count).Address <> "" And ws.Cells(dataRow + 3, 53).Address <> "" Then
                        .Values = ws.Range(ws.Cells(dataRow + 3, 53), ws.Cells(dataRow + 3, 52 + count))
                    End If
                    .Name = "LCL"
                    .Format.Line.ForeColor.RGB = colorLimit  ' 紅色
                    .Format.Line.DashStyle = msoLineDash
                    .Format.Line.Weight = 1.25
                End With
                On Error GoTo ErrorHandler
            End If
            
            ' Data (Red Solid with Markers) - 根據是否超出管制界限設定不同顏色
            On Error Resume Next
            With .SeriesCollection.NewSeries
                If ws.Cells(dataRow, 52 + count).Address <> "" And ws.Cells(dataRow, 53).Address <> "" Then
                    .Values = ws.Range(ws.Cells(dataRow, 53), ws.Cells(dataRow, 52 + count))
                End If
                ' 根據圖表類型設定不同的標籤名稱
                If typeIdx = 1 Then ' X-Bar 圖
                    .Name = ChrW(88) & ChrW(773) ' X̄ (X-bar)
                ElseIf typeIdx = 2 Then ' R 圖
                    .Name = "R"
                Else
                    .Name = "Data"
                End If
                .Format.Line.ForeColor.RGB = colorData
                .Format.Line.Weight = 1.5
                .MarkerStyle = xlMarkerStyleCircle
                .MarkerSize = 5
                .MarkerForegroundColor = colorData
                .MarkerBackgroundColor = colorData
                
                ' 為超出管制界限的點設定不同顏色（紅色）
                Dim j As Long
                For j = 1 To count
                    ' 檢查陣列範圍
                    If j <= UBound(dataArr) Then
                        Dim pointValue As Double
                        pointValue = dataArr(j)
                    
                    ' 根據圖表類型檢查不同的管制界限
                    Dim isOutOfControl As Boolean
                    isOutOfControl = False
                    
                    If typeIdx = 1 Then ' X-Bar 圖
                        If pointValue > ucl Or pointValue < lcl Then
                            isOutOfControl = True
                        End If
                    ElseIf typeIdx = 2 Then ' R 圖
                        ' R 圖只檢查上管制限，不檢查下管制限
                        If pointValue > ucl Then
                            isOutOfControl = True
                        End If
                    End If
                    
                    ' 如果超出管制界限，設定該點為紅色
                        If isOutOfControl Then
                            On Error Resume Next
                            .Points(j).MarkerForegroundColor = colorAlert
                            .Points(j).MarkerBackgroundColor = colorAlert
                            On Error GoTo ErrorHandler
                        End If
                    End If
                Next j
            End With
            On Error GoTo ErrorHandler
        End If
        
        On Error Resume Next
        .HasTitle = True
        .ChartTitle.Text = title
        .ChartTitle.Font.Name = "Microsoft JhengHei"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Bold = True
        .Legend.Position = xlLegendPositionTop
        
        ' 設定 X 軸和 Y 軸標題
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "樣本號數"
        .Axes(xlCategory).AxisTitle.Font.Bold = True
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Font.Bold = True
        
        ' 根據圖表類型設定 Y 軸標題
        If typeIdx = 1 Then ' X-Bar 圖
            .Axes(xlValue).AxisTitle.Text = "平均數"
        ElseIf typeIdx = 2 Then ' R 圖
            .Axes(xlValue).AxisTitle.Text = "全距"
        End If
        
        On Error GoTo ErrorHandler
        
        ' 設定X軸標籤為批號
        On Error Resume Next
        If count > 0 Then
            ' 建立包含批號的陣列
            Dim xLabels As Variant
            ReDim xLabels(1 To count)
            For i = 1 To count
                ' 檢查批號陣列範圍
                If i <= UBound(batchNames) Then
                    xLabels(i) = batchNames(i)
                Else
                    xLabels(i) = "批號" & i
                End If
            Next i
            
            ' 設定X軸標籤 - 根據圖表類型使用對應的系列名稱
            Dim seriesName As String
            If typeIdx = 1 Then ' X-Bar 圖
                seriesName = ChrW(88) & ChrW(773) ' X̄
            ElseIf typeIdx = 2 Then ' R 圖
                seriesName = "R"
            Else
                seriesName = "Data"
            End If
            .SeriesCollection(seriesName).XValues = xLabels
        End If
        On Error GoTo ErrorHandler
        
        ' 座標軸設定
        ' 座標軸設定與專業風格優化
        On Error Resume Next
        
        ' 1. 格線優化：使用淺灰色虛線，輔助閱讀但不搶眼
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = msoTrue
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(230, 230, 230)
        .Axes(xlValue).MajorGridlines.Format.Line.DashStyle = msoLineSysDot
        .Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.5
        
        ' 2. 背景優化：使用純白背景取代透明，便於列印與報告展示
        .ChartArea.Format.Fill.Visible = msoTrue
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Fill.Transparency = 0
        .PlotArea.Format.Fill.Visible = msoTrue
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.Transparency = 0
        
        ' 3. 圖表邊框：細緻的深灰色邊框
        .ChartArea.Format.Line.Visible = msoTrue
        .ChartArea.Format.Line.ForeColor.RGB = RGB(160, 160, 160)
        .ChartArea.Format.Line.Weight = 0.75
        
        ' 4. 座標軸線條加強
        .Axes(xlCategory).Format.Line.Visible = msoTrue
        .Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(89, 89, 89) ' 深灰色
        .Axes(xlCategory).Format.Line.Weight = 1
        
        .Axes(xlValue).Format.Line.Visible = msoTrue
        .Axes(xlValue).Format.Line.ForeColor.RGB = RGB(89, 89, 89) ' 深灰色
        .Axes(xlValue).Format.Line.Weight = 1
        On Error GoTo ErrorHandler
        
        ' 調整繪圖區以對齊表格
        On Error Resume Next
        .PlotArea.Top = 40 ' 增加頂部邊距以容納上方圖例
        .PlotArea.Height = h - 60
        .PlotArea.Left = 20
        .PlotArea.Width = w - 40
        ' 自動調整X軸刻度
        .Axes(xlCategory).TickLabelSpacing = 1
        .Axes(xlCategory).TickMarkSpacing = 1
        
        ' 當數據點太多時旋轉標籤以避免重疊
        If count > 15 Then
            .Axes(xlCategory).TickLabels.Orientation = 45
        End If
        On Error GoTo ErrorHandler
    End With
    
    ' 隱藏數據源
    On Error Resume Next
    ws.Range(ws.Cells(dataRow, 52), ws.Cells(dataRow + 5, 52 + 50)).Font.Color = RGB(255, 255, 255)
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    ' 忽略所有圖表相關錯誤，讓程序繼續執行
    Resume Next
End Sub

' ============================================================================
' SPC 常數查詢表 (n=2 to 32)
' ============================================================================
Private Sub GetSPCConstants(n As Integer, ByRef A2 As Double, ByRef D3 As Double, ByRef D4 As Double)
    On Error GoTo ErrorHandler
    
    ' 預設值 (避免 n<2 錯誤)
    If n < 2 Then
        A2 = 0: D3 = 0: D4 = 0
        Exit Sub
    End If
    
    ' 簡易查找表 (根據標準 SPC 係數表)
    Select Case n
        Case 2: A2 = 1.88: D3 = 0: D4 = 3.267
        Case 3: A2 = 1.023: D3 = 0: D4 = 2.373  ' Modified from 2.574 to match QE20002-C
        Case 4: A2 = 0.729: D3 = 0: D4 = 2.282
        Case 5: A2 = 0.577: D3 = 0: D4 = 2.115  ' Modified from 2.114 to match QE20002-C
        Case 6: A2 = 0.483: D3 = 0: D4 = 2.004
        Case 7: A2 = 0.419: D3 = 0.076: D4 = 1.924
        Case 8: A2 = 0.373: D3 = 0.136: D4 = 1.864
        Case 9: A2 = 0.337: D3 = 0.184: D4 = 1.816
        Case 10: A2 = 0.308: D3 = 0.223: D4 = 1.777
        Case 11: A2 = 0.285: D3 = 0.256: D4 = 1.744
        Case 12: A2 = 0.266: D3 = 0.283: D4 = 1.717
        Case 13: A2 = 0.249: D3 = 0.307: D4 = 1.693
        Case 14: A2 = 0.235: D3 = 0.328: D4 = 1.672
        Case 15: A2 = 0.223: D3 = 0.347: D4 = 1.653
        Case 16: A2 = 0.212: D3 = 0.363: D4 = 1.637
        Case 17: A2 = 0.203: D3 = 0.378: D4 = 1.622
        Case 18: A2 = 0.194: D3 = 0.391: D4 = 1.608
        Case 19: A2 = 0.187: D3 = 0.403: D4 = 1.597
        Case 20: A2 = 0.18: D3 = 0.415: D4 = 1.585
        Case 21: A2 = 0.173: D3 = 0.425: D4 = 1.575
        Case 22: A2 = 0.167: D3 = 0.434: D4 = 1.566
        Case 23: A2 = 0.162: D3 = 0.443: D4 = 1.557
        Case 24: A2 = 0.157: D3 = 0.451: D4 = 1.548
        Case 25: A2 = 0.153: D3 = 0.459: D4 = 1.541
        
        ' 大樣本數估計 (n > 25)
        ' 當 n 很大時，R 管制圖效率變低，通常改用 s 管制圖
        ' 但為了程式不報錯，這裡提供 n=26~32 的近似值
        Case 26: A2 = 0.149: D3 = 0.466: D4 = 1.534
        Case 27: A2 = 0.145: D3 = 0.473: D4 = 1.527
        Case 28: A2 = 0.141: D3 = 0.48: D4 = 1.52
        Case 29: A2 = 0.138: D3 = 0.486: D4 = 1.514
        Case 30: A2 = 0.135: D3 = 0.492: D4 = 1.508
        Case 31: A2 = 0.132: D3 = 0.498: D4 = 1.502
        Case 32: A2 = 0.129: D3 = 0.504: D4 = 1.496
            
        Case Else
            ' 超出範圍時使用 32 的值或公式估算 (避免崩潰)
            A2 = 3 / Sqr(n) ' 粗略估計
            D3 = 0.5
            D4 = 1.5
    End Select
    
    Exit Sub
    
ErrorHandler:
    ' 忽略所有錯誤，讓程序繼續執行
    A2 = 0: D3 = 0: D4 = 0
    Resume Next
End Sub

' ============================================================================
' 輔助：添加統計標籤 (使用線條繪製解決符號顯示問題)
' ============================================================================
Private Sub AddStatisticalLabel(ws As Worksheet, rng As Range, typeCode As Integer, valTxt As String)
    Dim txt As String
    Dim shp As Shape
    Dim line1 As Shape, line2 As Shape
    Dim startX As Double, startY As Double
    Dim charWidth As Double
    Dim baseName As String
    
    ' 根據類型設定文字
    Select Case typeCode
        Case 1: txt = "ΣX  = " & valTxt  ' ΣX bar
        Case 2: txt = "X  = " & valTxt   ' X double bar
        Case 3: txt = "ΣR  = " & valTxt  ' ΣR
        Case 4: txt = "R  = " & valTxt   ' R bar
    End Select
    
    ' 創建基礎名稱 (移除特殊字符以避免引用錯誤)
    baseName = "Label_" & Replace(Replace(rng.Address, "$", ""), ":", "")
    
    ' 如果已存在同名物件，嘗試刪除 (避免與舊物件衝突導致群組失敗)
    On Error Resume Next
    ws.Shapes(baseName).Delete
    ws.Shapes(baseName & "_L1").Delete
    ws.Shapes(baseName & "_L2").Delete
    On Error GoTo 0
    
    ' 創建文字方塊
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                   rng.Left + 2, rng.Top + 4, _
                                   rng.Width - 4, rng.Height - 4)
    
    With shp
        .Name = baseName
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse ' 禁用自動換行
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            With .TextRange
                .Text = txt
                .Font.Name = "Times New Roman" ' 使用 Times New Roman 字體較為標準
                .Font.Size = 12
                .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
                .ParagraphFormat.Alignment = msoAlignLeft
            End With
        End With
    End With
    
    ' 繪製線條 (Bar)
    charWidth = 9 ' 字元寬度估計值 (12pt Bold)
    startX = shp.Left
    startY = shp.Top + 4 ' 基準高度 (向下移動5個單位: 原本+2 -> 改為+4)
    
    Select Case typeCode
        Case 1 ' ΣX bar (線條在 X 上)
            ' Σ 寬度約 8，X 從 8 開始
            Set line1 = DrawLine(ws, startX + 8, startY, charWidth)
            line1.Name = baseName & "_L1"
            ' 群組
            ws.Shapes.Range(Array(baseName, line1.Name)).Group
            
        Case 2 ' X double bar (兩條線在 X 上)
            ' 線條 1 (下)
            Set line1 = DrawLine(ws, startX, startY, charWidth)
            line1.Name = baseName & "_L1"
            ' 線條 2 (上)
            Set line2 = DrawLine(ws, startX, startY - 2.5, charWidth)
            line2.Name = baseName & "_L2"
            ' 群組
            ws.Shapes.Range(Array(baseName, line1.Name, line2.Name)).Group
            
        Case 3 ' ΣR (無標線)
            ' Do nothing, no grouping needed as there is only one shape
            
        Case 4 ' R bar (線條在 R 上)
            Set line1 = DrawLine(ws, startX, startY, charWidth)
            line1.Name = baseName & "_L1"
            ' 群組
            ws.Shapes.Range(Array(baseName, line1.Name)).Group
            
    End Select
    
End Sub

Private Function DrawLine(ws As Worksheet, x As Double, y As Double, w As Double) As Shape
    Dim ln As Shape
    Set ln = ws.Shapes.AddLine(x, y, x + w, y)
    With ln.Line
        .ForeColor.RGB = RGB(0, 0, 0)
        .Weight = 1.25
        .Visible = msoTrue
    End With
    Set DrawLine = ln
End Function