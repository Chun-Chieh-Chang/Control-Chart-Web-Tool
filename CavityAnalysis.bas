Option Explicit

' ============================================================================
' 整合模穴分析模組
' 結合第二檔案的模穴分析 + 第三檔案的增強能力分析
' ============================================================================

' ============================================================================
' 執行整合模穴分析
' ============================================================================
Public Sub 執行模穴分析()
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim itemName As String
    Dim productCode As String
    Dim lastRow As Long, lastCol As Long
    
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
    wsName = DataInput.GenerateWorksheetName(productCode, itemName, "模穴分析")
    DataInput.SetWorksheetName wsTarget, wsName
    
    ' 6. 添加標題資訊
    Call AddHeaderInfo(wsTarget, wbSource.Name, itemName)
    
    ' 7. 讀取規格界限
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
    
    ' 8. 分析所有模穴
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.UsedRange.Columns.Count + wsSource.UsedRange.Column - 1
    
    Call 分析所有模穴(wsTarget, wsSource, lastRow, lastCol, target, usl, lsl)
    
    ' 9. 關閉來源檔案
    wbSource.Close False
    
    ' 10. 激活目標工作表
    wsTarget.Activate
    
    MsgBox "模穴分析完成！" & vbCrLf & _
           "已生成模穴統計摘要和比較圖", vbInformation
    Exit Sub
    
ErrorHandler:
    If Not wbSource Is Nothing Then wbSource.Close False
    MsgBox "執行模穴分析時發生錯誤：" & vbCrLf & vbCrLf & _
           "錯誤描述：" & Err.Description & vbCrLf & _
           "錯誤編號：" & Err.Number, vbCritical, "模穴分析錯誤"
End Sub

' ============================================================================
' 添加標題資訊
' ============================================================================
Private Sub AddHeaderInfo(ws As Worksheet, sourceName As String, itemName As String)
    ws.Cells(1, 1).Value = "來源檔案："
    ws.Cells(1, 2).Value = sourceName
    ws.Cells(1, 3).Value = "檢驗項目："
    ws.Cells(1, 4).Value = itemName
    ws.Cells(1, 5).Value = "分析時間："
    ws.Cells(1, 6).Value = Now
    
    ' 合并标题行前先清除内容以避免警告
    ws.Range("A1:F1").ClearContents
    
    With ws.Range("A1:F1")
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

' ============================================================================
' 分析所有模穴
' ============================================================================
Public Sub 分析所有模穴(wsTarget As Worksheet, wsSource As Worksheet, _
                        lastRow As Long, lastCol As Long, _
                        target As Double, usl As Double, lsl As Double)
    
    Dim cavityCol As Long
    Dim targetRow As Long
    Dim i As Long
    Dim values As Collection
    Dim value As Variant
    Dim cavityCount As Long
    
    ' 創建互動式篩選區域（第3-4行）
    Call 創建指標篩選區(wsTarget)
    
    ' 設置統計表標題（從第6行開始）
    wsTarget.Cells(6, 1).Value = "模穴號"
    wsTarget.Cells(6, 2).Value = "平均值"
    wsTarget.Cells(6, 3).Value = "組內標準差"
    wsTarget.Cells(6, 4).Value = "整體標準差"
    wsTarget.Cells(6, 5).Value = "範圍"
    wsTarget.Cells(6, 6).Value = "最大值"
    wsTarget.Cells(6, 7).Value = "最小值"
    wsTarget.Cells(6, 8).Value = "Cp"
    wsTarget.Cells(6, 9).Value = "Cpk"
    wsTarget.Cells(6, 10).Value = "Pp"
    wsTarget.Cells(6, 11).Value = "Ppk"
    wsTarget.Cells(6, 12).Value = "樣本數"
    
    ' 格式化標題
    ' 清除内容以避免合并单元格警告
    wsTarget.Range("A6:L6").ClearContents
    
    With wsTarget.Range("A6:L6")
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
    
    ' 分析每個模穴（從第7行開始）
    targetRow = 7
    cavityCount = 0
    
    For cavityCol = 5 To lastCol
        Dim headerValue As String
        headerValue = CStr(wsSource.Cells(1, cavityCol).Value)
        
        ' 檢查是否是模穴列
        If InStr(headerValue, "穴") > 0 Then
            cavityCount = cavityCount + 1
            Set values = New Collection
            
            ' 收集該模穴的所有數據（從第3行開始，第1行標題，第2行空白）
            For i = 3 To lastRow
                If IsNumeric(wsSource.Cells(i, cavityCol).Value) And _
                   wsSource.Cells(i, cavityCol).Value <> "" Then
                    values.Add CDbl(wsSource.Cells(i, cavityCol).Value)
                End If
            Next i
            
            If values.Count > 1 Then
                ' 計算並輸出統計量
                Call 計算模穴統計(wsTarget, targetRow, headerValue, values, usl, lsl)
                targetRow = targetRow + 1
            End If
        End If
    Next cavityCol
    
    ' 格式化
    wsTarget.Columns.AutoFit
    
    ' 生成模穴比較圖
    If targetRow > 4 Then
        On Error Resume Next
        Call 生成模穴比較圖(wsTarget, targetRow - 1, target, usl, lsl)
        If Err.Number <> 0 Then
            Debug.Print "圖表生成錯誤: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    Debug.Print "模穴分析完成，共分析 " & cavityCount & " 個模穴"
End Sub

' ============================================================================
' 計算模穴統計（整合版 - 包含 Pp, Ppk）
' ============================================================================
Private Sub 計算模穴統計(ws As Worksheet, row As Long, cavityName As String, _
                        values As Collection, usl As Double, lsl As Double)
    
    Dim value As Variant
    Dim sum As Double, mean As Double
    Dim withinStdDev As Double, overallStdDev As Double
    Dim maxVal As Double, minVal As Double
    Dim sumSq As Double
    Dim cp As Double, cpk As Double, cpu As Double, cpl As Double
    Dim pp As Double, ppk As Double
    Dim tolerance As Double
    Dim i As Long
    
    ' 計算基本統計量
    sum = 0
    maxVal = values(1)
    minVal = values(1)
    
    For Each value In values
        sum = sum + value
        If value > maxVal Then maxVal = value
        If value < minVal Then minVal = value
    Next value
    
    mean = sum / values.Count
    
    ' 計算整體標準差
    sumSq = 0
    For Each value In values
        sumSq = sumSq + (value - mean) ^ 2
    Next value
    overallStdDev = Sqr(sumSq / (values.Count - 1))
    
    ' 計算組內標準差（使用移動極差法）
    Dim mrSum As Double, mrCount As Long
    Dim d2 As Double
    mrSum = 0
    mrCount = 0
    
    For i = 2 To values.Count
        mrSum = mrSum + Abs(values(i) - values(i - 1))
        mrCount = mrCount + 1
    Next i
    
    d2 = 1.128  ' d2 常數（n=2）
    withinStdDev = (mrSum / mrCount) / d2
    
    ' 計算製程能力指標
    tolerance = usl - lsl
    
    ' 短期能力（使用組內標準差）
    cp = tolerance / (6 * withinStdDev)
    cpu = (usl - mean) / (3 * withinStdDev)
    cpl = (mean - lsl) / (3 * withinStdDev)
    cpk = Application.WorksheetFunction.Min(cpu, cpl)
    
    ' 長期性能（使用整體標準差）
    pp = tolerance / (6 * overallStdDev)
    ppk = Application.WorksheetFunction.Min((usl - mean) / (3 * overallStdDev), _
                                           (mean - lsl) / (3 * overallStdDev))
    
    ' 輸出結果
    ws.Cells(row, 1).Value = cavityName
    ws.Cells(row, 2).Value = Round(mean, 4)
    ws.Cells(row, 3).Value = Round(withinStdDev, 4)
    ws.Cells(row, 4).Value = Round(overallStdDev, 4)
    ws.Cells(row, 5).Value = Round(maxVal - minVal, 4)
    ws.Cells(row, 6).Value = Round(maxVal, 4)
    ws.Cells(row, 7).Value = Round(minVal, 4)
    ws.Cells(row, 8).Value = Round(cp, 3)
    ws.Cells(row, 9).Value = Round(cpk, 3)
    ws.Cells(row, 10).Value = Round(pp, 3)
    ws.Cells(row, 11).Value = Round(ppk, 3)
    ws.Cells(row, 12).Value = values.Count
    
    ' Cpk 顏色編碼
    If cpk >= 1.67 Then
        ws.Cells(row, 9).Interior.Color = RGB(198, 239, 206)  ' 綠色 - 優異
    ElseIf cpk >= 1.33 Then
        ws.Cells(row, 9).Interior.Color = RGB(198, 239, 206)  ' 綠色 - 良好
    ElseIf cpk >= 1.0 Then
        ws.Cells(row, 9).Interior.Color = RGB(255, 235, 156)  ' 黃色 - 可接受
    Else
        ws.Cells(row, 9).Interior.Color = RGB(255, 199, 206)  ' 紅色 - 不足
    End If
    
    ' Ppk 顏色編碼
    If ppk >= 1.67 Then
        ws.Cells(row, 11).Interior.Color = RGB(198, 239, 206)
    ElseIf ppk >= 1.33 Then
        ws.Cells(row, 11).Interior.Color = RGB(198, 239, 206)
    ElseIf ppk >= 1.0 Then
        ws.Cells(row, 11).Interior.Color = RGB(255, 235, 156)
    Else
        ws.Cells(row, 11).Interior.Color = RGB(255, 199, 206)
    End If
End Sub

' ============================================================================
' 生成模穴比較圖（含規格線和基準線）
' ============================================================================
Private Sub 生成模穴比較圖(ws As Worksheet, lastRow As Long, _
                          target As Double, usl As Double, lsl As Double)
    
    On Error GoTo ChartError
    
    Dim chartObj As ChartObject
    Dim chartTop As Long
    Dim dataCount As Long
    Dim i As Long
    
    dataCount = lastRow - 6  ' 數據行數（從第7行開始）
    chartTop = (lastRow + 3) * 15  ' 圖表位置在數據表下方
    
    Debug.Print "開始生成圖表，數據行數：" & dataCount & "，lastRow：" & lastRow
    
    ' ========== 平均值比較圖 ==========
    ' 先準備規格線數據（放在數據表右側）
    Dim specCol As Long
    specCol = 20  ' 從 T 列開始（遠離數據表）
    
    ws.Cells(6, specCol).Value = "目標值"
    ws.Cells(6, specCol + 1).Value = "USL"
    ws.Cells(6, specCol + 2).Value = "LSL"
    
    For i = 7 To lastRow
        ws.Cells(i, specCol).Value = target
        ws.Cells(i, specCol + 1).Value = usl
        ws.Cells(i, specCol + 2).Value = lsl
    Next i
    
    ' 創建圖表（改為折線圖）
    On Error Resume Next
    Set chartObj = ws.ChartObjects.Add(Left:=50, Top:=chartTop, Width:=600, Height:=300)
    If Err.Number <> 0 Then
        Debug.Print "創建平均值圖表失敗：" & Err.Description
        Err.Clear
        GoTo SkipAvgChart
    End If
    
    With chartObj.Chart
        .ChartType = xlLineMarkers
        
        ' 只添加平均值系列
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "平均值"
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(1).Values = ws.Range(ws.Cells(7, 2), ws.Cells(lastRow, 2))
        
        On Error Resume Next
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192)
        .SeriesCollection(1).Format.Line.Weight = 2.5
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).MarkerSize = 6
        On Error GoTo 0
        
        ' 添加目標值線
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "目標值"
        .SeriesCollection(2).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(2).Values = ws.Range(ws.Cells(7, specCol), ws.Cells(lastRow, specCol))
        .SeriesCollection(2).ChartType = xlLine
        
        On Error Resume Next
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(0, 176, 80)
        .SeriesCollection(2).Format.Line.Weight = 2
        .SeriesCollection(2).MarkerStyle = xlMarkerStyleNone
        On Error GoTo 0
        
        ' 添加 USL 線
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Name = "USL"
        .SeriesCollection(3).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(3).Values = ws.Range(ws.Cells(7, specCol + 1), ws.Cells(lastRow, specCol + 1))
        .SeriesCollection(3).ChartType = xlLine
        
        On Error Resume Next
        .SeriesCollection(3).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(3).Format.Line.DashStyle = msoLineDash
        .SeriesCollection(3).Format.Line.Weight = 1.5
        .SeriesCollection(3).MarkerStyle = xlMarkerStyleNone
        On Error GoTo 0
        
        ' 添加 LSL 線
        .SeriesCollection.NewSeries
        .SeriesCollection(4).Name = "LSL"
        .SeriesCollection(4).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(4).Values = ws.Range(ws.Cells(7, specCol + 2), ws.Cells(lastRow, specCol + 2))
        .SeriesCollection(4).ChartType = xlLine
        
        On Error Resume Next
        .SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(4).Format.Line.DashStyle = msoLineDash
        .SeriesCollection(4).Format.Line.Weight = 1.5
        .SeriesCollection(4).MarkerStyle = xlMarkerStyleNone
        On Error GoTo 0
        
        .HasTitle = True
        .ChartTitle.Text = "模穴平均值比較"
        
        On Error Resume Next
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "模穴"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "平均值"
        On Error GoTo 0
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
SkipAvgChart:
    On Error GoTo 0
    
    ' ========== Cpk 比較圖 ==========
    ' 先準備基準線數據（放在數據表右側）
    Dim benchCol As Long
    benchCol = 16  ' 從 P 列開始
    
    ws.Cells(6, benchCol).Value = "優異(1.67)"
    ws.Cells(6, benchCol + 1).Value = "良好(1.33)"
    ws.Cells(6, benchCol + 2).Value = "可接受(1.0)"
    
    For i = 7 To lastRow
        ws.Cells(i, benchCol).Value = 1.67
        ws.Cells(i, benchCol + 1).Value = 1.33
        ws.Cells(i, benchCol + 2).Value = 1
    Next i
    
    ' 創建圖表（包含 Cpk 和基準線數據）
    Set chartObj = ws.ChartObjects.Add(Left:=700, Top:=chartTop, Width:=600, Height:=300)
    With chartObj.Chart
        ' 先創建基本圖表
        .ChartType = xlColumnClustered
        
        ' 添加 Cpk 系列
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Cpk"
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(1).Values = ws.Range(ws.Cells(7, 9), ws.Cells(lastRow, 9))
        
        ' 添加基準線系列
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = ws.Cells(6, benchCol).Value
        .SeriesCollection(2).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(2).Values = ws.Range(ws.Cells(7, benchCol), ws.Cells(lastRow, benchCol))
        
        .SeriesCollection.NewSeries
        .SeriesCollection(3).Name = ws.Cells(6, benchCol + 1).Value
        .SeriesCollection(3).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(3).Values = ws.Range(ws.Cells(7, benchCol + 1), ws.Cells(lastRow, benchCol + 1))
        
        .SeriesCollection.NewSeries
        .SeriesCollection(4).Name = ws.Cells(6, benchCol + 2).Value
        .SeriesCollection(4).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(4).Values = ws.Range(ws.Cells(7, benchCol + 2), ws.Cells(lastRow, benchCol + 2))
        
        .HasTitle = True
        .ChartTitle.Text = "模穴 Cpk 比較"
        
        On Error Resume Next
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "模穴"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Cpk"
        On Error GoTo 0
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        
        ' 將基準線改為線條圖並設定樣式
        On Error Resume Next
        With .SeriesCollection(2)
            .ChartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(0, 176, 80)
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 1.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        With .SeriesCollection(3)
            .ChartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(255, 192, 0)
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 1.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        With .SeriesCollection(4)
            .ChartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
            .Format.Line.DashStyle = msoLineDash
            .Format.Line.Weight = 1.5
            .MarkerStyle = xlMarkerStyleNone
        End With
        
        ' 為 Cpk 柱子設定顏色
        Dim pointIndex As Long
        Dim cpkValue As Double
        
        For pointIndex = 1 To dataCount
            cpkValue = ws.Cells(6 + pointIndex, 9).Value  ' 數據從第7行開始
            
            ' 檢查是否為數值
            If IsNumeric(cpkValue) Then
                With .SeriesCollection(1).Points(pointIndex)
                    If cpkValue >= 1.67 Then
                        .Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
                    ElseIf cpkValue >= 1.33 Then
                        .Format.Fill.ForeColor.RGB = RGB(146, 208, 80)
                    ElseIf cpkValue >= 1.0 Then
                        .Format.Fill.ForeColor.RGB = RGB(255, 192, 0)
                    Else
                        .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
                    End If
                End With
            End If
        Next pointIndex
        
        If Err.Number <> 0 Then
            Debug.Print "設定圖表樣式錯誤：" & Err.Description & " (錯誤號：" & Err.Number & ")"
            Err.Clear
        End If
        On Error GoTo 0
    End With
    
    On Error GoTo 0
    
    ' ========== 標準差比較圖 ==========
    Set chartObj = ws.ChartObjects.Add(Left:=50, Top:=chartTop + 350, Width:=600, Height:=300)
    With chartObj.Chart
        .ChartType = xlLineMarkers
        
        ' 添加組內標準差系列
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "組內標準差"
        .SeriesCollection(1).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(1).Values = ws.Range(ws.Cells(7, 3), ws.Cells(lastRow, 3))
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192)
        .SeriesCollection(1).Format.Line.Weight = 2.5
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).MarkerSize = 6
        
        ' 添加整體標準差系列
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "整體標準差"
        .SeriesCollection(2).XValues = ws.Range(ws.Cells(7, 1), ws.Cells(lastRow, 1))
        .SeriesCollection(2).Values = ws.Range(ws.Cells(7, 4), ws.Cells(lastRow, 4))
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        .SeriesCollection(2).Format.Line.Weight = 2.5
        .SeriesCollection(2).MarkerStyle = xlMarkerStyleSquare
        .SeriesCollection(2).MarkerSize = 6
        
        .HasTitle = True
        .ChartTitle.Text = "模穴標準差比較"
        
        On Error Resume Next
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "模穴"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "標準差"
        On Error GoTo 0
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    If Err.Number <> 0 Then
        Debug.Print "標準差圖表錯誤：" & Err.Description
        Err.Clear
    End If
    
    Debug.Print "圖表生成完成"
    Exit Sub
    
ChartError:
    Debug.Print "圖表生成發生錯誤："
    Debug.Print "  錯誤編號：" & Err.Number
    Debug.Print "  錯誤描述：" & Err.Description
    Debug.Print "  錯誤來源：" & Err.Source
    MsgBox "圖表生成時發生錯誤：" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "統計數據已生成，但部分圖表可能未顯示。", vbExclamation, "圖表生成警告"
    Resume Next
End Sub

' ============================================================================
' 創建指標篩選區
' ============================================================================
Private Sub 創建指標篩選區(ws As Worksheet)
    ' 標題
    ws.Cells(3, 1).Value = "圖表說明"
    ws.Cells(3, 1).Font.Bold = True
    ws.Cells(3, 1).Font.Size = 11
    ws.Cells(3, 1).Font.Color = RGB(0, 112, 192)
    
    ' 圖表說明
    ws.Cells(4, 1).Value = "平均值比較圖: 顯示各模穴平均值與規格界限"
    ws.Cells(4, 7).Value = "Cpk 比較圖: 顯示各模穴能力指標 (彩色柱子表示能力等級)"
    ws.Cells(5, 1).Value = "標準差比較圖: 顯示各模穴的組內標準差和整體標準差"
    
    ws.Cells(4, 1).Font.Size = 9
    ws.Cells(4, 7).Font.Size = 9
    ws.Cells(5, 1).Font.Size = 9
    
    ' 設定背景色
    ws.Range("A3:N5").Interior.Color = RGB(217, 225, 242)
End Sub


