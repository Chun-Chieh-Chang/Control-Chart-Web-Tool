Option Explicit

' ============================================================================
' 執行群組分析（Min-Max-Avg 管制圖）
' ============================================================================
Public Sub 執行群組分析()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim itemName As String
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
    
    ' 1. 選擇數據檔案
    Set wbSource = DataInput.SelectDataFile()
    If wbSource Is Nothing Then Exit Sub
    
    ' 2. 選擇檢驗項目
    itemName = DataInput.SelectInspectionItem(wbSource)
    If itemName = "" Then
        wbSource.Close False
        Exit Sub
    End If
    
    ' 3. 設定來源工作表
    Set wsSource = wbSource.Worksheets(itemName)
    
    ' 4. 提取產品品號
    productCode = DataInput.ExtractProductCode(wbSource.FullName)
    
    ' 5. 創建分析工作表
    Set wsTarget = ThisWorkbook.Worksheets.Add
    Dim wsName As String
    wsName = DataInput.GenerateWorksheetName(productCode, itemName, "群組分析")
    DataInput.SetWorksheetName wsTarget, wsName
    
    ' 6. 添加標題資訊
    wsTarget.Cells(1, 1).Value = "來源檔案："
    wsTarget.Cells(1, 2).Value = wbSource.Name
    wsTarget.Cells(1, 3).Value = "檢驗項目："
    wsTarget.Cells(1, 4).Value = itemName
    wsTarget.Cells(1, 5).Value = "分析時間："
    wsTarget.Cells(1, 6).Value = Now
    With wsTarget.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With
    
    ' 7. 讀取規格
    ' 規格數據在第2行（第1行標題，第2行規格+數據）
    ' 使用 Text 屬性避免 Excel 將 "J" 等字符誤判為日期格式
    Dim target As Double, usl As Double, lsl As Double
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
    
    ' 8. 設置數據表標題
    wsTarget.Cells(3, 1).Value = "批號"
    wsTarget.Cells(3, 2).Value = "平均值 (Avg)"
    wsTarget.Cells(3, 3).Value = "最大值 (Max)"
    wsTarget.Cells(3, 4).Value = "最小值 (Min)"
    wsTarget.Cells(3, 5).Value = "全距 (Range)"
    wsTarget.Cells(3, 6).Value = "樣本數 (n)"
    wsTarget.Cells(3, 7).Value = "USL"
    wsTarget.Cells(3, 8).Value = "Target"
    wsTarget.Cells(3, 9).Value = "LSL"
    
    With wsTarget.Range("A3:I3")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 9. 找到數據範圍
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.UsedRange.Columns.Count + wsSource.UsedRange.Column - 1
    
    ' 10. 識別模穴列
    Set cavityCols = New Collection
    Debug.Print "開始識別模穴列..."
    For cavityCol = 5 To lastCol
        If InStr(CStr(wsSource.Cells(1, cavityCol).Value), "穴") > 0 Then
            cavityCols.Add cavityCol
            Debug.Print "  找到模穴列: " & wsSource.Cells(1, cavityCol).Value & " (Col " & cavityCol & ")"
        End If
    Next cavityCol
    
    Debug.Print "總共找到 " & cavityCols.Count & " 個模穴列"
    
    If cavityCols.Count = 0 Then
        MsgBox "未找到模穴數據列（標題需包含'穴'字）", vbExclamation
        wbSource.Close False
        Exit Sub
    End If
    
    ' 11. 分析每一批次
    targetRow = 4
    Debug.Print "開始分析批次數據 (Rows 3 to " & lastRow & ")..."
    
    For i = 3 To lastRow  ' 從第3行開始讀取數據（第1行標題，第2行空白）
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
    
    Debug.Print "分析完成，TargetRow = " & targetRow
    
    ' 12. 格式化
    wsTarget.Columns.AutoFit
    
    ' 13. 生成群組管制圖
    If targetRow > 4 Then
        Debug.Print "準備生成群組管制圖..."
        Call 生成群組管制圖(wsTarget, targetRow - 1, itemName)
    Else
        Debug.Print "數據不足 (TargetRow <= 4)，跳過圖表生成"
        MsgBox "數據不足，無法生成管制圖", vbExclamation
    End If
    
    ' 14. 關閉來源檔案
    wbSource.Close False
    
    wsTarget.Activate
    MsgBox "群組分析完成！" & vbCrLf & "已生成 Min-Max-Avg 管制圖", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not wbSource Is Nothing Then wbSource.Close False
    MsgBox "執行群組分析時發生錯誤：" & vbCrLf & Err.Description, vbCritical
End Sub

' ============================================================================
' 生成群組管制圖：繪製 Min-Max-Avg 圖表
' ============================================================================
Public Sub 生成群組管制圖(ws As Worksheet, lastRow As Long, itemName As String)
    Dim chartObj As ChartObject
    
    ' 1. Min-Max-Avg 趨勢圖
    On Error Resume Next
    Set chartObj = ws.ChartObjects.Add(Left:=50, Top:=50, Width:=900, Height:=450)
    If Err.Number <> 0 Then
        MsgBox "創建圖表物件失敗：" & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    With chartObj.Chart
        .ChartType = xlLine
        
        ' 清除自動生成的系列
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        
        ' Series 1: Max (最大值) - 紅色細線
        With .SeriesCollection.NewSeries
            .Name = "Max (最大值)"
            .Values = ws.Range(ws.Cells(4, 3), ws.Cells(lastRow, 3))
            .XValues = ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, 1))
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 1.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        ' Series 2: Avg (平均值) - 藍色粗線
        With .SeriesCollection.NewSeries
            .Name = "Avg (平均值)"
            .Values = ws.Range(ws.Cells(4, 2), ws.Cells(lastRow, 2))
            .Format.Line.ForeColor.RGB = RGB(0, 112, 192)
            .Format.Line.Weight = 2.5
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 5
            .MarkerForegroundColor = RGB(0, 112, 192)
            .MarkerBackgroundColor = RGB(255, 255, 255)
        End With
        
        ' Series 3: Min (最小值) - 紅色細線
        With .SeriesCollection.NewSeries
            .Name = "Min (最小值)"
            .Values = ws.Range(ws.Cells(4, 4), ws.Cells(lastRow, 4))
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.Weight = 1.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        ' Series 4: USL - 橘色虛線
        With .SeriesCollection.NewSeries
            .Name = "USL"
            .Values = ws.Range(ws.Cells(4, 7), ws.Cells(lastRow, 7))
            .Format.Line.ForeColor.RGB = RGB(255, 192, 0)
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 2
        End With
        
        ' Series 5: Target - 綠色實線
        With .SeriesCollection.NewSeries
            .Name = "Target"
            .Values = ws.Range(ws.Cells(4, 8), ws.Cells(lastRow, 8))
            .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
            .Format.Line.Weight = 1.5
        End With
        
        ' Series 6: LSL - 橘色虛線
        With .SeriesCollection.NewSeries
            .Name = "LSL"
            .Values = ws.Range(ws.Cells(4, 9), ws.Cells(lastRow, 9))
            .Format.Line.ForeColor.RGB = RGB(255, 192, 0)
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 2
        End With
        
        ' 設定標題和軸
        .HasTitle = True
        .ChartTitle.Text = itemName & " - 群組管制圖 (Min-Max-Avg)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "生產批號"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "測量值"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    ' 2. 全距 (Range) 圖 - 監控模穴間變異
    On Error Resume Next
    Set chartObj = ws.ChartObjects.Add(Left:=50, Top:=520, Width:=900, Height:=300)
    If Err.Number <> 0 Then
        MsgBox "創建 Range 圖表物件失敗：" & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    With chartObj.Chart
        .ChartType = xlLine
        
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        
        ' Series 1: Range
        With .SeriesCollection.NewSeries
            .Name = "Range (模穴間變異)"
            .Values = ws.Range(ws.Cells(4, 5), ws.Cells(lastRow, 5))
            .XValues = ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, 1))
            .Format.Line.ForeColor.RGB = RGB(112, 48, 160)
            .Format.Line.Weight = 2
            .MarkerStyle = xlMarkerStyleSquare
            .MarkerSize = 5
        End With
        
        .HasTitle = True
        .ChartTitle.Text = itemName & " - 模穴間變異 (Max - Min)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "生產批號"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "全距 (Range)"
    End With
End Sub
