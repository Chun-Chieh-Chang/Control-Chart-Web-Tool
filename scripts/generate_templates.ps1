$ErrorActionPreference = "Stop"

function Create-Template {
    param (
        [int]$Cavities
    )
    
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    try {
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = "Control Chart"
        
        # --- Layout Constants (Matching JS/VBA Logic) ---
        $StartRow = 7
        $EndRow = $StartRow + $Cavities - 1
        $SumRow = $EndRow + 1
        $MeanRow = $SumRow + 1
        $RangeRow = $MeanRow + 1
        $ChartRow = $RangeRow + 5
        
        # --- Create Dummy Data for Chart Setup ---
        # 25 Batches (Cols C to AA -> 3 to 27)
        for ($c = 3; $c -le 27; $c++) {
            $ws.Cells.Item($MeanRow, $c).Value2 = 10 # Dummy Mean
            $ws.Cells.Item($RangeRow, $c).Value2 = 2  # Dummy Range
        }
        
        # --- Setup Limit Hidden Data Area (Rows 100-102) ---
        # Row 100: UCL, 101: CL, 102: LCL (XBar)
        # Row 103: UCL, 104: CL (R)
        for ($c = 3; $c -le 27; $c++) {
            $ws.Cells.Item(100, $c).Value2 = 12 # UCL
            $ws.Cells.Item(101, $c).Value2 = 10 # CL
            $ws.Cells.Item(102, $c).Value2 = 8  # LCL
            $ws.Cells.Item(103, $c).Value2 = 5  # R-UCL
            $ws.Cells.Item(104, $c).Value2 = 2  # R-CL
        }
        
        # --- Generate X-Bar Chart ---
        $chartObj = $ws.ChartObjects().Add(0, $chartObj.Top + 100, 700, 250)
        $chart = $chartObj.Chart
        $chart.ChartType = 65 # xlLineMarkers
        
        # Data Series (Mean)
        $s1 = $chart.SeriesCollection().NewSeries()
        $s1.Values = $ws.Range($ws.Cells.Item($MeanRow, 3), $ws.Cells.Item($MeanRow, 27))
        $s1.Name = "X-Bar"
        
        # Limit Series (UCL/CL/LCL)
        $s2 = $chart.SeriesCollection().NewSeries()
        $s2.Values = $ws.Range($ws.Cells.Item(100, 3), $ws.Cells.Item(100, 27))
        $s2.Name = "UCL"
        
        $s3 = $chart.SeriesCollection().NewSeries()
        $s3.Values = $ws.Range($ws.Cells.Item(101, 3), $ws.Cells.Item(101, 27))
        $s3.Name = "CL"
        
        $s4 = $chart.SeriesCollection().NewSeries()
        $s4.Values = $ws.Range($ws.Cells.Item(102, 3), $ws.Cells.Item(102, 27))
        $s4.Name = "LCL"
        
        # Format X-Bar Chart
        $chart.HasTitle = $true
        $chart.ChartTitle.Text = "X-Bar Chart"
        $chartObj.Top = $ws.Cells.Item($ChartRow, 1).Top
        $chartObj.Left = $ws.Cells.Item($ChartRow, 1).Left
        
        # --- Generate R Chart ---
        $rChartObj = $ws.ChartObjects().Add(0, 0, 700, 250)
        $rChart = $rChartObj.Chart
        $rChart.ChartType = 65
        
        # Data
        $rs1 = $rChart.SeriesCollection().NewSeries()
        $rs1.Values = $ws.Range($ws.Cells.Item($RangeRow, 3), $ws.Cells.Item($RangeRow, 27))
        $rs1.Name = "Range"
        
        # Limits
        $rs2 = $rChart.SeriesCollection().NewSeries()
        $rs2.Values = $ws.Range($ws.Cells.Item(103, 3), $ws.Cells.Item(103, 27))
        $rs2.Name = "UCL"
        
        $rs3 = $rChart.SeriesCollection().NewSeries()
        $rs3.Values = $ws.Range($ws.Cells.Item(104, 3), $ws.Cells.Item(104, 27))
        $rs3.Name = "CL"
        
        $rChart.HasTitle = $true
        $rChart.ChartTitle.Text = "R Chart"
        $rChartObj.Top = $ws.Cells.Item($ChartRow + 15, 1).Top
        $rChartObj.Left = $ws.Cells.Item($ChartRow, 1).Left
        
        # Save
        $path = Join-Path (Get-Location) "templates\template_$Cavities.xlsx"
        $wb.SaveAs($path)
        $wb.Close($false)
        
        Write-Host "Generated: $path"
        
    } catch {
        Write-Error "Failed to generate $Cavities : $_"
    } finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

# Generate for common cavity counts
Create-Template -Cavities 8
Create-Template -Cavities 16
Create-Template -Cavities 24
Create-Template -Cavities 32
Create-Template -Cavities 48
