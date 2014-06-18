
'This sub adds a chart
'Params:
'srcSheetName - source data sheet name
'srcRange As String - source range for chart
'ChartSheetName - In which sheet to create the chart
'leftPos - left position for chart
'topPos -top position for chart
'chartWidth - chart width
'chartHeight - chart height
'chartType - excel chart type(xlPie\xlLine etc.)
'chartName - Chart name
'titleChart - chart title
'...

'Example:
'Call MakeChart("Report", "A1:D8", "Report", Cells(11, "C").Left, Cells(11, "C").Top, 240, 245, xlPie, "TopChart1", "Conversions Per Search Engine", msoElementDataLabelBestFit, Array("#,##0.00"), False)
'


Public Sub Make3DPieChart(srcSheetName As String, srcRange As String, ChartSheetName As String, leftPos As Double, topPos As Double, chartWidth As Double, _
    chartHeight As Double, chartName As String, titleChart As String, Optional colorChartArea As Boolean)
'创建3D饼图
    Sheets(ChartSheetName).Select
    Dim chObj As Object
    Range("A1").Select
    
    Set chObj = ActiveSheet.ChartObjects.add(leftPos, topPos, chartWidth, chartHeight)
    chObj.name = chartName
    
    With chObj.Chart
        .SetSourceData Source:=Sheets(srcSheetName).Range(srcRange)
    End With
    
    ActiveSheet.ChartObjects(chartName).Activate
    ActiveChart.ChartStyle = 26
    If titleChart <> "" Then
        ActiveChart.HasTitle = True
        ActiveChart.ChartTitle.text = titleChart
    End If
    
    ActiveChart.ApplyCustomType chartType:=xl3DPie
    ActiveChart.SeriesCollection(1).ApplyDataLabels
    ActiveChart.SeriesCollection(1).ApplyDataLabels ShowPercentage:=True
    ActiveChart.SeriesCollection(1).format.ThreeD.RotationY = 30
        
    ActiveChart.Legend.top = 5
    ActiveChart.Legend.height = chartHeight - 2 * ActiveChart.Legend.top
    
    If colorChartArea Then Call ColorChart(ChartSheetName, chartName)

    
End Sub

Public Sub MakeColumnClusteredChart(srcSheetName As String, srcRange As String, ChartSheetName As String, leftPos As Double, topPos As Double, chartWidth As Double, _
    chartHeight As Double, chartName As String, titleChart As String, Optional colorChartArea As Boolean)
'创建柱状图
    
    Sheets(ChartSheetName).Select
    Dim chObj As Object
    Range("A1").Select
    
    Set chObj = ActiveSheet.ChartObjects.add(leftPos, topPos, chartWidth, chartHeight)
    chObj.name = chartName
    With chObj.Chart
        .ApplyCustomType chartType:=xlColumnClustered
        .SetSourceData Source:=Sheets(srcSheetName).Range(srcRange)
    End With
    ActiveSheet.ChartObjects(chartName).Activate
    ActiveChart.ChartStyle = 26
    If titleChart <> "" Then
        ActiveChart.HasTitle = True
        ActiveChart.ChartTitle.text = titleChart
    End If

    
    If LegendTop = True Then
        ActiveChart.SetElement (msoElementLegendTop)
    End If
    
    Call SecondaryAxis(ChartSheetName, chartName)

    If colorChartArea Then Call ColorChart(ChartSheetName, chartName)
    
End Sub


Sub SecondaryAxis(chartSheet As String, chartName As String)
    Sheets(chartSheet).Select
    ActiveSheet.ChartObjects(chartName).Activate
    With ActiveChart
      If .SeriesCollection.Count > 1 Then
       If ActiveSheet.ChartObjects(chartName).Chart.chartType = xlColumnClustered Then
         .SeriesCollection(2).chartType = xlLine
       End If
       .SetElement (msoElementPrimaryCategoryAxisShow)
       .SeriesCollection(2).AxisGroup = 2
       .SeriesCollection(1).AxisGroup = 1
      End If
    End With
End Sub
Sub ColorChart(chartSheet As String, chartName As String)
    Sheets(chartSheet).Select
    ActiveSheet.ChartObjects(chartName).Interior.color = RGB(242, 242, 242)
    ActiveSheet.ChartObjects(chartName).Activate
    ActiveChart.ChartArea.format.Line.Visible = msoFalse
    With ActiveChart.ChartArea.format.Fill
    
'      .ForeColor.RGB = RGB(195, 214, 155) '(239, 251, 197)
     .OneColorGradient Style:=msoGradientHorizontal, Variant:=3, Degree:=0.9
    End With
    ActiveChart.PlotArea.format.Fill.Visible = False
End Sub


