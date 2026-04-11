Sub CreateProfessionalQCAnalysis_WithSheetNameInTitle()
    Dim targetSheet As Worksheet
    Dim sheetIndex As Variant, colLetter As String
    Dim lastRow As Long, dataRng As Range
    Dim avg As Double, stdDev As Double, i As Integer, j As Integer
    Dim graphMin As Double, graphMax As Double
    Dim binWidth As Double, numBins As Integer: numBins = 20
    Dim chartObj As ChartObject
    Dim seriesHist As Series, seriesCurve As Series, seriesMarkers As Series, seriesSpec As Series
    Dim xFormat As String
    Dim uslInput As Variant, lslInput As Variant
    Dim hasUSL As Boolean, hasLSL As Boolean
    
    ' --- 1. 選取工作表與數據 ---
    Dim sheetList As String: For i = 1 To Sheets.Count: sheetList = sheetList & i & ". " & Sheets(i).Name & vbCrLf: Next i
    sheetIndex = InputBox(sheetList, "第一步：選取工作表編號")
    If sheetIndex = "" Or Not IsNumeric(sheetIndex) Then Exit Sub
    Set targetSheet = Sheets(CInt(sheetIndex))
    targetSheet.Activate
    colLetter = InputBox("請輸入數據欄位字母（例如：E）:", "第二步：數據欄位", "E")
    If colLetter = "" Then Exit Sub
    
    ' --- 2. 獲取規格界限 (USL/LSL) - 可選項 ---
    uslInput = Application.InputBox("請輸入 USL (規格上限)，若無則直接按確定:", "規格設定", Type:=1 + 2)
    lslInput = Application.InputBox("請輸入 LSL (規格下限)，若無則直接按確定:", "規格設定", Type:=1 + 2)
    hasUSL = IsNumeric(uslInput) And uslInput <> ""
    hasLSL = IsNumeric(lslInput) And lslInput <> ""

    ' --- 3. 計算統計量 ---
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, colLetter).End(xlUp).Row
    Set dataRng = targetSheet.Range(targetSheet.Cells(1, colLetter), targetSheet.Cells(lastRow, colLetter))
    avg = Application.WorksheetFunction.Average(dataRng)
    stdDev = Application.WorksheetFunction.StDev_P(dataRng)
    
    ' 設定 X 軸繪圖範圍 (±4個標準差)
    graphMin = avg - 4 * stdDev: graphMax = avg + 4 * stdDev
    If hasUSL Then If uslInput > graphMax Then graphMax = uslInput + stdDev
    If hasLSL Then If lslInput < graphMin Then graphMin = lslInput - stdDev
    
    ' 動態精度 (針對微小數據如 0.83)
    If stdDev < 0.01 Then xFormat = "0.0000" Else xFormat = "0.00"
    
    ' --- 4. 輔助數據區 (AA:AJ) ---
    targetSheet.Columns("AA:AJ").Clear
    targetSheet.Range("AA1:AJ1").Value = Array("組界", "次數", "", "曲線X", "曲線Y", "", "標記X", "標記Y", "SpecX", "SpecY")
    binWidth = (graphMax - graphMin) / numBins
    For i = 1 To numBins
        Dim bStart As Double: bStart = graphMin + (i - 1) * binWidth
        targetSheet.Cells(i + 1, "AA").Value = bStart + (binWidth / 2)
        targetSheet.Cells(i + 1, "AB").Value = Application.WorksheetFunction.CountIfs(dataRng, ">=" & bStart, dataRng, "<" & (bStart + binWidth))
    Next i
    For i = 0 To 100
        Dim curX As Double: curX = graphMin + (i * (graphMax - graphMin) / 100)
        targetSheet.Cells(i + 2, "AD").Value = curX
        targetSheet.Cells(i + 2, "AE").Value = Application.WorksheetFunction.Norm_Dist(curX, avg, stdDev, False)
    Next i
    For i = 0 To 6
        targetSheet.Cells(i + 2, "AG").Value = avg + (i - 3) * stdDev
        targetSheet.Cells(i + 2, "AH").Value = Application.WorksheetFunction.Norm_Dist(targetSheet.Cells(i + 2, "AG").Value, avg, stdDev, False)
    Next i
    Dim yPeak As Double: yPeak = Application.WorksheetFunction.Norm_Dist(avg, avg, stdDev, False)
    If hasLSL Then targetSheet.Cells(2, "AI").Value = lslInput: targetSheet.Cells(2, "AJ").Value = yPeak * 0.7
    If hasUSL Then targetSheet.Cells(3, "AI").Value = uslInput: targetSheet.Cells(3, "AJ").Value = yPeak * 0.7
    
    ' --- 5. 繪製圖表 ---
    On Error Resume Next: targetSheet.ChartObjects("QC_Final_Analysis").Delete: On Error GoTo 0
    Set chartObj = targetSheet.ChartObjects.Add(Left:=350, Width:=850, Top:=10, Height:=480)
    chartObj.Name = "QC_Final_Analysis"
    
    With chartObj.Chart
        ' 強制 Scatter 初始化
        .ChartType = xlXYScatter: Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        
        Set seriesHist = .SeriesCollection.NewSeries
        seriesHist.XValues = targetSheet.Range("AA2:AA" & numBins + 1)
        seriesHist.Values = targetSheet.Range("AB2:AB" & numBins + 1)
        seriesHist.Name = "次數"

        Set seriesCurve = .SeriesCollection.NewSeries
        seriesCurve.ChartType = xlXYScatterSmoothNoMarkers
        seriesCurve.XValues = targetSheet.Range("AD2:AD102"): seriesCurve.Values = targetSheet.Range("AE2:AE102")
        seriesCurve.AxisGroup = 2: seriesCurve.Name = "正態分佈"

        Set seriesMarkers = .SeriesCollection.NewSeries
        seriesMarkers.ChartType = xlXYScatter: seriesMarkers.XValues = targetSheet.Range("AG2:AG8")
        seriesMarkers.Values = targetSheet.Range("AH2:AH8"): seriesMarkers.AxisGroup = 2: seriesMarkers.Name = "SD標記"

        If hasLSL Or hasUSL Then
            Set seriesSpec = .SeriesCollection.NewSeries
            seriesSpec.ChartType = xlXYScatter
            Dim rx As String: rx = IIf(hasLSL And hasUSL, "AI2:AI3", IIf(hasLSL, "AI2", "AI3"))
            Dim ry As String: ry = IIf(hasLSL And hasUSL, "AJ2:AJ3", IIf(hasLSL, "AJ2", "AJ3"))
            seriesSpec.XValues = targetSheet.Range(rx): seriesSpec.Values = targetSheet.Range(ry)
            seriesSpec.AxisGroup = 2: seriesSpec.Name = "規格界限"
        End If

        ' --- 直方圖轉換與邊框細節 ---
        seriesHist.ChartType = xlColumnClustered
        .ChartGroups(1).GapWidth = 0
        With seriesHist
            .Format.Fill.ForeColor.RGB = RGB(210, 210, 210)
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(140, 140, 140)
            .Format.Line.Weight = 0.75 ' 精緻細線
            .ApplyDataLabels: .DataLabels.Position = xlLabelPositionOutsideEnd
            .DataLabels.Font.Size = 8: .DataLabels.Font.Color = RGB(100, 100, 100)
        End With

        ' --- 6. 座標軸潔淨化 ---
        .Axes(xlValue, 1).HasMajorGridlines = False: .Axes(xlValue, 2).HasMajorGridlines = False
        .Axes(xlCategory, 1).HasMajorGridlines = False: .Axes(xlCategory, 1).TickLabels.NumberFormat = xFormat
        
        .HasAxis(xlCategory, 2) = True
        On Error Resume Next
        With .Axes(xlCategory, 2)
            .MinimumScale = graphMin: .MaximumScale = graphMax
            .TickLabelPosition = xlTickLabelPositionNone: .MajorTickMark = xlTickMarkNone: .AxisLine.Visible = msoFalse
        End With
        .Axes(xlValue, 2).MinimumScale = 0: On Error GoTo 0

        ' --- 7. 線條樣式 ---
        seriesCurve.Format.Line.ForeColor.RGB = RGB(0, 112, 192): seriesCurve.Format.Line.Weight = 2.5
        With seriesMarkers
            .MarkerStyle = xlMarkerStyleCircle: .MarkerSize = 5
            .Format.Fill.ForeColor.RGB = RGB(255, 0, 0): .Format.Line.Visible = False
            .HasErrorBars = True: .ErrorBar xlY, xlMinusValues, xlPercent, 100
            .ErrorBars.EndStyle = xlNoCap: .ErrorBars.Format.Line.ForeColor.RGB = RGB(140, 140, 140)
            .ErrorBars.Format.Line.DashStyle = msoLineDash: .ErrorBar xlX, xlNone, xlFixedValue, 0
        End With
        If Not seriesSpec Is Nothing Then
            With seriesSpec
                .MarkerStyle = xlMarkerStyleNone: .HasErrorBars = True
                .ErrorBar xlY, xlBoth, xlPercent, 100
                .ErrorBars.Format.Line.ForeColor.RGB = RGB(255, 0, 0): .ErrorBars.Format.Line.Weight = 2
                .ErrorBars.Format.Line.DashStyle = msoLineDash: .ErrorBar xlX, xlNone, xlFixedValue, 0
            End With
        End If

        ' --- 8. 標籤優化 (白色背框) ---
        Dim lblNames: lblNames = Array("-3σ", "-2σ", "-1σ", "平均值", "+1σ", "+2σ", "+3σ")
        seriesMarkers.ApplyDataLabels
        For j = 1 To 7
            With seriesMarkers.DataLabels(j)
                .Text = lblNames(j - 1) & vbCrLf & "(" & Format(targetSheet.Cells(j + 1, "AG"), xFormat) & ", " & Format(targetSheet.Cells(j + 1, "AH"), "0.1") & ")"
                .Position = xlLabelPositionAbove: .Font.Size = 8
                .Format.Fill.ForeColor.RGB = RGB(255, 255, 255): .Format.Fill.Transparency = 0.8
                .Format.Line.Visible = msoTrue: .Format.Line.ForeColor.RGB = RGB(220, 220, 220)
            End With
        Next j

        If Not seriesSpec Is Nothing Then
            seriesSpec.ApplyDataLabels
            Dim specIdx As Integer: specIdx = 1
            If hasLSL Then
                seriesSpec.DataLabels(specIdx).Text = "LSL: " & lslInput
                specIdx = specIdx + 1
            End If
            If hasUSL Then
                seriesSpec.DataLabels(specIdx).Text = "USL: " & uslInput
            End If
            For j = 1 To seriesSpec.DataLabels.Count
                With seriesSpec.DataLabels(j)
                    .Font.Color = RGB(255, 0, 0): .Font.Bold = True: .Position = xlLabelPositionAbove
                    .Format.Fill.ForeColor.RGB = RGB(255, 255, 255): .Format.Fill.Transparency = 0.2
                    .Format.Line.Visible = msoTrue: .Format.Line.ForeColor.RGB = RGB(255, 0, 0)
                End With
            Next j
        End If

        ' --- 9. 標題與圖例 (包含工作表名稱) ---
        .ChartArea.Border.LineStyle = -4142
        .Legend.Position = xlLegendPositionBottom
        
        ' *** 關鍵：將工作表名稱加入標題 ***
        .HasTitle = True
        .ChartTitle.Text = targetSheet.Name & " 數據常態分佈分析 - 欄位 " & colLetter
        
        On Error Resume Next
        .Legend.LegendEntries("SD標記").Delete: .Legend.LegendEntries("規格界限").Delete
        On Error GoTo 0
    End With
    
    MsgBox "繪製完成！圖表標題已包含工作表名稱：" & targetSheet.Name, vbInformation
End Sub