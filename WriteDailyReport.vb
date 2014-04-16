Sub WriteDailyReport()

    Dim totalColCount As Integer, totalRowCount As Integer
    Dim dataSheetName As String, reportName As String, channelName As String
    Dim strChannel As String, strDate As String, strMonth As String, strYear As String, strYearMonth As String
    Dim rowIndex As Integer, colIndex As Integer, i As Integer, j As Integer
    '是否符合channel条件的数据
    Dim filterFlag As Boolean
    '是否要显示年份的对比报表
    Dim isDisplayYear As Boolean
    '定义渠道数组
    Dim arrChannels As Variant
    '定义日期数组
    Dim arrDate As Variant
    '定义多维数组，存放对比的两个月的数据
    Dim twoMonthDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两周的数据
    Dim twoWeekDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两天的数据
    Dim twoDayDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两年的数据
    Dim twoYearDataArr(1 To 28, 1 To 2)
    '月份数据，数据中最小的两个月
    Dim twoMonthArr(1 To 2) As Variant, dicMonth As Variant, arrMonth As Variant
    '年份数据，数据中最小的两年
    Dim twoYearArr(1 To 2) As Variant, dicYear As Variant, arrYear As Variant
    '指标数据
    Dim fImpression, fClick, fCost, fConversion, fConversion1, fConversion2, fConversion3, fConversion4, fRevenue
    '1元人民币对应的外币
    Dim rate As Double, currencyFormat As Variant
        
    Application.ScreenUpdating = False
    dataSheetName = "Daily元数据"
    reportName = "Daily"
    
    totalRowCount = Sheets(dataSheetName).UsedRange.Rows.Count
    totalColCount = Sheets(dataSheetName).UsedRange.Columns.Count
    
    Set dicYear = CreateObject("Scripting.Dictionary")
    Set dicMonth = CreateObject("Scripting.Dictionary")
    
    '汇率及货币格式
    rate = Sheets(dataSheetName).Cells(1, 4)
    currencyFormat = Sheets(dataSheetName).Cells(1, 4).NumberFormatLocal
    
    '得到数据中最小的两个月份和年份，第3行起始
    For rowIndex = 3 To totalRowCount
        strDate = Sheets(dataSheetName).Cells(rowIndex, 2)
        If strDate = "" Then Exit For
        strMonth = CStr(Month(strDate))
        strYear = CStr(Year(strDate))
        strYearMonth = strYear & "/" & strMonth
        
        If twoMonthArr(1) = "" Then
            twoMonthArr(1) = strYearMonth
        ElseIf YearMonthCompare(twoMonthArr(1), strYearMonth) Then
            twoMonthArr(2) = twoMonthArr(1)
            twoMonthArr(1) = strYearMonth
        ElseIf YearMonthCompare(strYearMonth, twoMonthArr(1)) Then
            If twoMonthArr(2) = "" Or YearMonthCompare(twoMonthArr(2), strYearMonth) Then
                twoMonthArr(2) = strYearMonth
            End If
        End If
        
        If twoYearArr(1) = "" Then
            twoYearArr(1) = strYear
        ElseIf twoYearArr(1) > strYear Then
            twoYearArr(2) = twoYearArr(1)
            twoYearArr(1) = strYear
        ElseIf twoYearArr(1) < strYear Then
            If twoYearArr(2) = "" Or twoYearArr(2) > strYear Then
                twoYearArr(2) = strYear
            End If
        End If
        
        If dicYear(strYear) = "" Then dicYear(strYear) = 1
        If dicMonth(strYearMonth) = "" Then dicMonth(strYearMonth) = 1
        
    Next rowIndex

    
    If twoMonthArr(2) = "" Then twoMonthArr(2) = twoMonthArr(1)
    If twoYearArr(2) = "" Then twoYearArr(2) = twoYearArr(1)
    '对月份和年份数据排序，数组下标0开始
    arrYear = dicYear.keys
    arrMonth = dicMonth.keys
    Call StringSort(arrYear)
    Call YearMonthSort(arrMonth)
    
    Set dicYear = Nothing
    Set dicMonth = Nothing
    
    arrChannels = DDLSourceFromDataColumn("A", dataSheetName, False, 3)
    
    'Channle下拉列表清空，再赋值
    Sheets(reportName).Shapes("ddlDailyChannel").ControlFormat.List = "all"
    Sheets(reportName).Shapes("ddlDailyChannel").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlDailyChannel").ControlFormat.AddItem arrChannels
    Sheets(reportName).Shapes("ddlDailyChannel").OnAction = "DDLDailyChanged"
    
    Sheets(reportName).Shapes("ddlDailyChartChannel").ControlFormat.List = "all"
    Sheets(reportName).Shapes("ddlDailyChartChannel").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlDailyChartChannel").ControlFormat.AddItem arrChannels
    Sheets(reportName).Shapes("ddlDailyChartChannel").OnAction = "DailyChartDataTypeChange"
    
    '向报表中写入前后对比的两个月份列表
    Sheets(reportName).Shapes("ddlMonthStart").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlMonthStart").ControlFormat.AddItem arrMonth
    Sheets(reportName).Shapes("ddlMonthStart").ControlFormat.ListIndex = ArrayDataIndex(arrMonth, twoMonthArr(1)) + 1
    Sheets(reportName).Shapes("ddlMonthStart").OnAction = "DDLDailyChanged"

    Sheets(reportName).Shapes("ddlMonthEnd").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlMonthEnd").ControlFormat.AddItem arrMonth
    Sheets(reportName).Shapes("ddlMonthEnd").ControlFormat.ListIndex = ArrayDataIndex(arrMonth, twoMonthArr(2)) + 1
    Sheets(reportName).Shapes("ddlMonthEnd").OnAction = "DDLDailyChanged"
    
    '图形的选项
    Sheets(reportName).Shapes("chartDataType").ControlFormat.List = "By Month"
    Sheets(reportName).Shapes("chartDataType").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("chartDataType").ControlFormat.AddItem arrMonth
    Sheets(reportName).Shapes("chartDataType").OnAction = "DailyChartDataTypeChange"

    arrDate = DDLSourceFromDataColumn("B", dataSheetName, False, 3)
    '一天开始日期列表
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlDayStart").OnAction = "DDLDailyChanged"
    '一天结束日期列表
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlDayEnd").OnAction = "DDLDailyChanged"
    '一周开始日期列表
    Sheets(reportName).Shapes("ddlWeekStart").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlWeekStart").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlWeekStart").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlWeekStart").OnAction = "DDLDailyChanged"
    '一周结束日期列表
    Sheets(reportName).Shapes("ddlWeekEnd").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlWeekEnd").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlWeekEnd").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlWeekEnd").OnAction = "DDLDailyChanged"
    
    If IsShapeExists(reportName, "ddlYearStart") Then
        isDisplayYear = True
        '一年开始日期列表
        Sheets(reportName).Shapes("ddlYearStart").ControlFormat.RemoveAllItems
        Sheets(reportName).Shapes("ddlYearStart").ControlFormat.AddItem arrYear
        Sheets(reportName).Shapes("ddlYearStart").ControlFormat.ListIndex = ArrayDataIndex(arrYear, CInt(twoYearArr(1))) + 1
        Sheets(reportName).Shapes("ddlYearStart").OnAction = "DDLDailyChanged"
        '一年结束日期列表
        Sheets(reportName).Shapes("ddlYearEnd").ControlFormat.RemoveAllItems
        Sheets(reportName).Shapes("ddlYearEnd").ControlFormat.AddItem arrYear
        Sheets(reportName).Shapes("ddlYearEnd").ControlFormat.ListIndex = ArrayDataIndex(arrYear, CInt(twoYearArr(2))) + 1
        Sheets(reportName).Shapes("ddlYearEnd").OnAction = "DDLDailyChanged"
    Else
        isDisplayYear = False
    End If
    
    channelName = "all"
    
    '汇总数据
    For rowIndex = 3 To totalRowCount
    
        strChannel = Sheets(dataSheetName).Cells(rowIndex, 1)
        strDate = CStr(Sheets(dataSheetName).Cells(rowIndex, 2))
        If strDate = "" Then Exit For
        fImpression = Sheets(dataSheetName).Cells(rowIndex, 3)
        fClick = Sheets(dataSheetName).Cells(rowIndex, 4)
        fCost = Sheets(dataSheetName).Cells(rowIndex, 5) * rate
        fConversion1 = Sheets(dataSheetName).Cells(rowIndex, 6)
        fConversion2 = Sheets(dataSheetName).Cells(rowIndex, 7)
        fConversion3 = Sheets(dataSheetName).Cells(rowIndex, 8)
        fConversion4 = Sheets(dataSheetName).Cells(rowIndex, 9)
        fConversion = fConversion1 + fConversion2 + fConversion3 + fConversion4
        fRevenue = Sheets(dataSheetName).Cells(rowIndex, 10) * rate

        If (channelName = strChannel Or channelName = "" Or channelName = "all") Then
            filterFlag = True
        Else
            filterFlag = False
        End If

        If filterFlag Then
            '两月对比数据
            strMonth = CStr(Year(strDate)) & "/" & CStr(Month(strDate))
            If strMonth = twoMonthArr(1) Then
                Call FillArrayData(twoMonthDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If
            If strMonth = twoMonthArr(2) Then
                Call FillArrayData(twoMonthDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If

            '两周对比数据
            If IsInWeek(CStr(arrDate(1, 1)), strDate) Then
                Call FillArrayData(twoWeekDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
                Call FillArrayData(twoWeekDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If

            '两天对比数据
            If CStr(arrDate(1, 1)) = strDate Then
                Call FillArrayData(twoDayDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
                Call FillArrayData(twoDayDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If
            
            '显示年份的对比数据
            If isDisplayYear Then
                strYear = CStr(Year(strDate))
                If strYear = twoYearArr(1) Then
                    Call FillArrayData(twoYearDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
                End If
                If strYear = twoYearArr(2) Then
                    Call FillArrayData(twoYearDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
                End If
            End If
            
        End If
    
    
    Next rowIndex
    
    
    '向报表中写入前后对比的两个月份数据
    Call ComMetricArrayData(twoMonthDataArr)
    Call FillCellData(reportName, 9, 4, twoMonthDataArr)

    '向报表中写入前后对比的两天数据
    Call ComMetricArrayData(twoDayDataArr)
    Call FillCellData(reportName, 40, 4, twoDayDataArr)

    '向报表中写入前后对比的两周数据
    Call ComMetricArrayData(twoWeekDataArr)
    Call FillCellData(reportName, 40, 9, twoWeekDataArr)
    
    '向报表中写入前后对比的两年数据
    If isDisplayYear Then
        Call ComMetricArrayData(twoYearDataArr)
        Call FillCellData(reportName, 40, 14, twoYearDataArr)
    End If
    
    '设置单个格货币格式
    Call SetCurrencyFormat(reportName, 9, 4, currencyFormat)
    Call SetCurrencyFormat(reportName, 40, 4, currencyFormat)
    Call SetCurrencyFormat(reportName, 40, 9, currencyFormat)
    Call SetCurrencyFormat(reportName, 40, 14, currencyFormat)
    '绘制饼图
    Call PieChartForDaily
    '绘制柱形图
    Call ColumnClusteredChartForDaily

End Sub

Sub DDLDailyChanged()
'渠道列表改变或日期改变
    Dim fImpression, fClick, fCost, fConversion, fConversion1, fConversion2, fConversion3, fConversion4, fRevenue
    Dim strChannel As String, strDate As String, strMonth As String, strYear As String
    Dim i As Integer, j As Integer, totalRowCount As Integer, rowIndex As Integer
    Dim channelName As String, dataSheetName As String, reportName As String
    Dim monthStart As String, monthEnd As String, yearStart As String, yearEnd As String
    Dim weekStart As Date, weekEnd As Date, dayStart As Date, dayEnd As Date
    '是否要显示年份的对比报表
    Dim isDisplayYear As Boolean
    '是否符合campaign条件的数据
    Dim filterFlag As Boolean
    Dim rate As Double
    '定义多维数组，存放对比的两个月的数据
    Dim twoMonthDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两周的数据
    Dim twoWeekDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两天的数据
    Dim twoDayDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两年的数据
    Dim twoYearDataArr(1 To 28, 1 To 2)
    
    Application.ScreenUpdating = False
    

    filterFlag = True
    dataSheetName = "Daily元数据"
    reportName = "Daily"
    
    '汇率
    rate = Sheets(dataSheetName).Cells(1, 4)
    
    'Chanle被改变
    channelName = ValueDDL("ddlDailyChannel", reportName)
    monthStart = ValueDDL("ddlMonthStart", reportName)
    monthEnd = ValueDDL("ddlMonthEnd", reportName)
    If IsShapeExists(reportName, "ddlYearStart") Then
        yearStart = ValueDDL("ddlYearStart", reportName)
        yearEnd = ValueDDL("ddlYearEnd", reportName)
        isDisplayYear = True
    Else
        isDisplayYear = False
    End If
    weekStart = ValueDDL("ddlWeekStart", reportName)
    weekEnd = ValueDDL("ddlWeekEnd", reportName)
    dayStart = ValueDDL("ddlDayStart", reportName)
    dayEnd = ValueDDL("ddlDayEnd", reportName)
    totalRowCount = Sheets(dataSheetName).UsedRange.Rows.Count
    
    For rowIndex = 3 To totalRowCount
        
        strChannel = Sheets(dataSheetName).Cells(rowIndex, 1)
        strDate = CStr(Sheets(dataSheetName).Cells(rowIndex, 2))
        If strDate = "" Then Exit For
        fImpression = Sheets(dataSheetName).Cells(rowIndex, 3)
        fClick = Sheets(dataSheetName).Cells(rowIndex, 4)
        fCost = Sheets(dataSheetName).Cells(rowIndex, 5) * rate
        fConversion1 = Sheets(dataSheetName).Cells(rowIndex, 6)
        fConversion2 = Sheets(dataSheetName).Cells(rowIndex, 7)
        fConversion3 = Sheets(dataSheetName).Cells(rowIndex, 8)
        fConversion4 = Sheets(dataSheetName).Cells(rowIndex, 9)
        fConversion = fConversion1 + fConversion2 + fConversion3 + fConversion4
        fRevenue = Sheets(dataSheetName).Cells(rowIndex, 10) * rate
    
    
        If (channelName = strChannel Or channelName = "" Or channelName = "all") Then
            filterFlag = True
        Else
            filterFlag = False
        End If

        If filterFlag Then
            
            '两月对比数据
            strMonth = CStr(Year(strDate)) & "/" & CStr(Month(strDate))
            If strMonth = monthStart Then
                Call FillArrayData(twoMonthDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If

            If strMonth = monthEnd Then
                Call FillArrayData(twoMonthDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If
            If isDisplayYear Then
                '年对比数据
                strYear = CStr(Year(strDate))
                If strYear = yearStart Then
                    Call FillArrayData(twoYearDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
                End If
                If strYear = yearEnd Then
                    Call FillArrayData(twoYearDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
                End If
            End If
            
            '两周对比数据
            If IsInWeek(weekStart, strDate) Then
                Call FillArrayData(twoWeekDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If
            If IsInWeek(weekEnd, strDate) Then
                Call FillArrayData(twoWeekDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If
            
            '两天对比数据
            If dayStart = strDate Then
                Call FillArrayData(twoDayDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If
            If dayEnd = strDate Then
                Call FillArrayData(twoDayDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If
        End If
        
    Next rowIndex
    
    
    '向报表中写入前后对比的两个月份数据
    Call ComMetricArrayData(twoMonthDataArr)
    Call FillCellData(reportName, 9, 4, twoMonthDataArr)

    '向报表中写入前后对比的两天数据
    Call ComMetricArrayData(twoDayDataArr)
    Call FillCellData(reportName, 40, 4, twoDayDataArr)

    '向报表中写入前后对比的两周数据
    Call ComMetricArrayData(twoWeekDataArr)
    Call FillCellData(reportName, 40, 9, twoWeekDataArr)
    
    '向报表中写入前后对比的两年数据
    If isDisplayYear Then
        Call ComMetricArrayData(twoYearDataArr)
        Call FillCellData(reportName, 40, 14, twoYearDataArr)
    End If

End Sub

Sub ColumnClusteredChartForDaily()
'生成簇状柱形图，默认By Month按月汇总

    Dim metricFields() As Variant
    Dim i As Long, n As Long, j As Long, rowCount As Integer
    Dim dicMetricData, dicDate, arr, arrKeys
    Dim dataSheetName As String, reportName As String
    Dim cellDate As Variant, cellYearMonth As String
    '1元人民币对应的外币
    Dim rate As Double
    
    Application.ScreenUpdating = False
    dataSheetName = "Daily元数据"
    reportName = "Daily"
    
    '汇率
    rate = Sheets(dataSheetName).Cells(1, 4)
    
    '得到有数据的行，需要清理
    n = Sheets(reportName).[AA65536].End(xlUp).Row
    If n > 1 Then
        Sheets(reportName).Range("AA2:AQ" & n).ClearContents
    End If
    
    
    '指标数据字典
    Set dicMetricData = CreateObject("Scripting.Dictionary")
    '日期数据字典
    Set dicDate = CreateObject("Scripting.Dictionary")
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A3:J" & n)
    
    For i = 1 To UBound(arr)
        '取得数据中的日期
        cellDate = arr(i, 2)
        cellYearMonth = CStr(Year(cellDate)) + "/" + CStr(Month(cellDate))

        If dicDate(cellYearMonth) = "" Then
            dicDate(cellYearMonth) = i
            rowCount = rowCount + 1
        End If
        dicMetricData(cellYearMonth & "Imp") = dicMetricData(cellYearMonth & "Imp") + arr(i, 3)
        dicMetricData(cellYearMonth & "Click") = dicMetricData(cellYearMonth & "Click") + arr(i, 4)
        dicMetricData(cellYearMonth & "Cost") = dicMetricData(cellYearMonth & "Cost") + arr(i, 5) * rate
        dicMetricData(cellYearMonth & "Conversion") = dicMetricData(cellYearMonth & "Conversion") + arr(i, 6) + arr(i, 7) + arr(i, 8) + arr(i, 9)
        dicMetricData(cellYearMonth & "Revenue") = dicMetricData(cellYearMonth & "Revenue") + arr(i, 10) * rate

    Next
    
    Erase arr

    arr = Sheets(reportName).Range("AA2", Sheets(reportName).Cells(rowCount + 1, 32))
    
    arrKeys = dicDate.keys
    
    For i = 1 To rowCount
        arr(i, 1) = arrKeys(i - 1)
        arr(i, 2) = dicMetricData(arr(i, 1) & "Imp")
        arr(i, 3) = dicMetricData(arr(i, 1) & "Click")
        arr(i, 4) = dicMetricData(arr(i, 1) & "Cost")
        arr(i, 5) = dicMetricData(arr(i, 1) & "Conversion")
        arr(i, 6) = dicMetricData(arr(i, 1) & "Revenue")
    Next
    
    Sheets(reportName).[AA2].Resize(rowCount, UBound(arr, 2)) = arr
    Erase arr
    Set dicDate = Nothing
    Set dicMetricData = Nothing
    
    '图形的数据源，默认两列imp,clicks
    Call CopyPasteWithDesign(reportName, "AB2:AC" & rowCount + 1, reportName, "ap2")
    
    '没有柱状图重新创建，有的话使用老格式，默认By Day汇总
    If Not IsShapeExists(reportName, "dailyChart") Then
        Call MakeColumnClusteredChart(reportName, "AA2:AA" & (rowCount + 1) & ",AP2:AQ" & (rowCount + 1), reportName, Range("H7").left, Range("H7").top, 810, 220, "dailyChart", "By Day", True)
    Else
        Sheets(reportName).ChartObjects("dailyChart").Chart.SetSourceData (Sheets(reportName).Range("AA2:AA" & (rowCount + 1) & ",AP2:AQ" & (rowCount + 1)))
    End If
    
    '设置在图表中显示隐藏数据
    Sheets(reportName).ChartObjects("dailyChart").Chart.PlotVisibleOnly = False
    
    '设置系列值
    ActiveSheet.ChartObjects("dailyChart").Activate
    ActiveChart.SeriesCollection(1).name = "Imp"
    ActiveChart.SeriesCollection(2).name = "Clicks"
    
    ActiveChart.ChartTitle.text = "By Month"
    
    If Not IsShapeExists(reportName, "chartMetricOne") Then
        metricFields = Array("Imp", "Clicks", "Cost", "Conversion", "Revenue")
        Call AddDdlInOneCell(reportName, "K5", "DailyChartMetricOneChange", metricFields, "chartMetricOne", 1)
    Else
        Sheets(reportName).Shapes("chartMetricOne").ControlFormat.ListIndex = 1
    End If
    
    If Not IsShapeExists(reportName, "chartMetricTwo") Then
        metricFields = Array("Imp", "Clicks", "Cost", "Conversion", "Revenue")
        Call AddDdlInOneCell(reportName, "M5", "DailyChartMetricTwoChange", metricFields, "chartMetricTwo", 2)
    Else
        Sheets(reportName).Shapes("chartMetricTwo").ControlFormat.ListIndex = 2
    End If
    
    Sheets(reportName).Shapes("ddlDailyChartChannel").ZOrder msoBringToFront
    Sheets(reportName).Shapes("chartDataType").ZOrder msoBringToFront
    Sheets(reportName).Shapes("chartMetricOne").ZOrder msoBringToFront
    Sheets(reportName).Shapes("chartMetricTwo").ZOrder msoBringToFront
    
End Sub

Sub PieChartForDaily()
'生成饼图
    Dim metricFields() As Variant
    Dim i As Long, n As Long, j As Long, rowCount As Integer
    Dim dicMetricData, dicChannel, arrKeys, arr, cellChannel
    Dim dataSheetName As String, reportName As String
    Dim rate As Double
    Application.ScreenUpdating = False
    
    dataSheetName = "Daily元数据"
    reportName = "Daily"
    
    '汇率
    rate = Sheets(dataSheetName).Cells(1, 4)
    rowCount = 0
    '得到有数据的行，需要清理
    n = Sheets(reportName).[S65536].End(xlUp).Row
    If n > 1 Then
        Sheets(reportName).Range("S2:Y" & n).ClearContents
    End If
    '指标数据字典
    Set dicMetricData = CreateObject("Scripting.Dictionary")
    '渠道名称字典
    Set dicChannel = CreateObject("Scripting.Dictionary")
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A3:J" & n)
    
    For i = 1 To UBound(arr)
        '取得数据中的渠道
        cellChannel = arr(i, 1)
        If dicChannel(cellChannel) = "" Then
            dicChannel(cellChannel) = i
            rowCount = rowCount + 1
        End If
        dicMetricData(cellChannel & "Imp") = dicMetricData(cellChannel & "Imp") + arr(i, 3)
        dicMetricData(cellChannel & "Click") = dicMetricData(cellChannel & "Click") + arr(i, 4)
        dicMetricData(cellChannel & "Cost") = dicMetricData(cellChannel & "Cost") + arr(i, 5) * rate
        dicMetricData(cellChannel & "Conversion") = dicMetricData(cellChannel & "Conversion") + arr(i, 6) + arr(i, 7) + arr(i, 8) + arr(i, 9)
        dicMetricData(cellChannel & "Revenue") = dicMetricData(cellChannel & "Revenue") + arr(i, 10) * rate
    Next

    Erase arr
    '25为第Y列索引号
    arr = Sheets(reportName).Range("S2", Sheets(reportName).Cells(rowCount + 1, 25))

    arrKeys = dicChannel.keys

    For i = 1 To rowCount
        arr(i, 1) = arrKeys(i - 1)
        arr(i, 2) = dicMetricData(arr(i, 1) & "Imp")
        arr(i, 3) = dicMetricData(arr(i, 1) & "Click")
        arr(i, 4) = dicMetricData(arr(i, 1) & "Cost")
        arr(i, 5) = dicMetricData(arr(i, 1) & "Conversion")
        arr(i, 6) = dicMetricData(arr(i, 1) & "Revenue")
    Next

    Sheets(reportName).[S2].Resize(rowCount, UBound(arr, 2)) = arr
    Erase arr
    Set dicMetricData = Nothing
    Set dicChannel = Nothing

    '图形的数据源，一列
    Call CopyPasteWithDesign(reportName, "V2:V" & rowCount + 1, reportName, "Y2")

    '没有饼图重新创建，有的话使用老格式
    If Not IsShapeExists(reportName, "PieDailyChart") Then
        Call Make3DPieChart(reportName, "S2:S" & (rowCount + 1) & ",Y2:Y" & (rowCount + 1), reportName, Range("C69").left, Range("C69").top, 450, 250, "PieDailyChart", "All Channel", True)
    Else
        Sheets(reportName).ChartObjects("PieDailyChart").Chart.SetSourceData (Sheets(reportName).Range("S2:S" & (rowCount + 1) & ",Y2:Y" & (rowCount + 1)))
    End If

    If Not IsShapeExists(reportName, "chartMetricThree") Then
        metricFields = Array("Cost", "Conversion", "Revenue")
        Call AddDdlInOneCell(reportName, "C72", "DailyChartMetricThreeChange", metricFields, "chartMetricThree", 1)
    Else
        Sheets(reportName).Shapes("chartMetricThree").ControlFormat.ListIndex = 1
    End If

    Sheets(reportName).Shapes("chartMetricThree").ZOrder msoBringToFront
    
End Sub


Sub DailyChartDataTypeChange()
'图形选项By Month，By 某个月 或Channel被选择

    Dim metricFields() As Variant
    Dim i As Long, n As Long, j As Long, rowCount As Integer
    Dim dicMetricData, dicDate, arr, arrKeys
    Dim dataSheetName As String, reportName As String, strChartDataType As String, channelName As String, cellChannel As String
    Dim cellDate As Variant, cellYearMonth As String
    Dim rate As Double
    Dim filterFlag As Boolean
    
    Application.ScreenUpdating = False
    dataSheetName = "Daily元数据"
    reportName = "Daily"
    
    '汇率
    rate = Sheets(dataSheetName).Cells(1, 4)
    
    '得到有数据的行，需要清理
    n = Sheets(reportName).[AA65536].End(xlUp).Row
    If n > 1 Then
        Sheets(reportName).Range("AA2:AQ" & n).ClearContents
    End If
    
    strChartDataType = ValueDDL("chartDataType", reportName)
    channelName = ValueDDL("ddlDailyChartChannel", reportName)
    
    '指标数据字典
    Set dicMetricData = CreateObject("Scripting.Dictionary")
    '日期数据字典
    Set dicDate = CreateObject("Scripting.Dictionary")
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A3:J" & n)
    
    For i = 1 To UBound(arr)
        '取得数据中的渠道和日期
        cellChannel = arr(i, 1)
        cellDate = arr(i, 2)
        
        If (channelName = cellChannel Or channelName = "all") Then
            filterFlag = True
        Else
            filterFlag = False
        End If
        
        If filterFlag Then
        
            cellYearMonth = CStr(Year(cellDate)) + "/" + CStr(Month(cellDate))
            'By Month将数据按月份汇总，By 某个具体的月份则按天汇总
            If strChartDataType = "By Month" Then
                If dicDate(cellYearMonth) = "" Then
                    dicDate(cellYearMonth) = i
                    rowCount = rowCount + 1
                End If
                dicMetricData(cellYearMonth & "Imp") = dicMetricData(cellYearMonth & "Imp") + arr(i, 3)
                dicMetricData(cellYearMonth & "Click") = dicMetricData(cellYearMonth & "Click") + arr(i, 4)
                dicMetricData(cellYearMonth & "Cost") = dicMetricData(cellYearMonth & "Cost") + arr(i, 5) * rate
                dicMetricData(cellYearMonth & "Conversion") = dicMetricData(cellYearMonth & "Conversion") + arr(i, 6) + arr(i, 7) + arr(i, 8) + arr(i, 9)
                dicMetricData(cellYearMonth & "Revenue") = dicMetricData(cellYearMonth & "Revenue") + arr(i, 10) * rate
            Else
                If strChartDataType = cellYearMonth Then
                    If dicDate(cellDate) = "" Then
                        dicDate(cellDate) = i
                        rowCount = rowCount + 1
                    End If
                    dicMetricData(cellDate & "Imp") = dicMetricData(cellDate & "Imp") + arr(i, 3)
                    dicMetricData(cellDate & "Click") = dicMetricData(cellDate & "Click") + arr(i, 4)
                    dicMetricData(cellDate & "Cost") = dicMetricData(cellDate & "Cost") + arr(i, 5) * rate
                    dicMetricData(cellDate & "Conversion") = dicMetricData(cellDate & "Conversion") + arr(i, 6) + arr(i, 7) + arr(i, 8) + arr(i, 9)
                    dicMetricData(cellDate & "Revenue") = dicMetricData(cellDate & "Revenue") + arr(i, 10) * rate
                End If
            End If
        
        End If
    Next
    
    Erase arr
    arr = Sheets(reportName).Range("AA2", Sheets(reportName).Cells(rowCount + 1, 32))
    
    arrKeys = dicDate.keys
    
    For i = 1 To rowCount
        arr(i, 1) = arrKeys(i - 1)
        arr(i, 2) = dicMetricData(arr(i, 1) & "Imp")
        arr(i, 3) = dicMetricData(arr(i, 1) & "Click")
        arr(i, 4) = dicMetricData(arr(i, 1) & "Cost")
        arr(i, 5) = dicMetricData(arr(i, 1) & "Conversion")
        arr(i, 6) = dicMetricData(arr(i, 1) & "Revenue")
    Next
    
    If rowCount = 0 Then Exit Sub
    
    Sheets(reportName).[AA2].Resize(rowCount, UBound(arr, 2)) = arr
    Erase arr
    Set dicDate = Nothing
    Set dicMetricData = Nothing
    
    '重置数据源，报表标题
    Sheets(reportName).ChartObjects("dailyChart").Activate
    If strChartDataType = "By Day" Then
        ActiveChart.ChartTitle.text = "By Day"
    Else
        ActiveChart.ChartTitle.text = strChartDataType
    End If
    
    '图形的数据源，两列
    Call CopyPasteWithDesign(reportName, "AB2:AC" & rowCount + 1, reportName, "ap2")
    '绑定图形数据源
    'Sheets(reportName).ChartObjects("dailyChart").Chart.SetSourceData (Sheets(reportName).Range("AA2:AA" & (rowCount + 1) & ",AP2:AQ" & (rowCount + 1)))
    
    '设置系列值，当数据只有一行时，系列数量会变成1，需要调整
    ActiveSheet.ChartObjects("dailyChart").Activate
    ActiveChart.SeriesCollection(1).name = "Imp"
    
    If ActiveChart.SeriesCollection.Count = 1 Then
        ActiveChart.SeriesCollection.NewSeries
    End If
    ActiveChart.SeriesCollection(2).name = "Clicks"
    
    ActiveChart.SeriesCollection(1).XValues = "=" & reportName & "!$AA$2:$AA$" & (rowCount + 1)
    ActiveChart.SeriesCollection(1).Values = "=" & reportName & "!$AP$2:$AP$" & (rowCount + 1)
    
    ActiveChart.SeriesCollection(2).XValues = "=" & reportName & "!$AA$2:$AA$" & (rowCount + 1)
    ActiveChart.SeriesCollection(2).Values = "=" & reportName & "!$AQ$2:$AQ$" & (rowCount + 1)
    '设置第二个指标使用折线显示
    Call SecondaryAxis(reportName, "dailyChart")
    
    
    '指标选择项重置为imp,clicks
    If Not IsShapeExists(reportName, "chartMetricOne") Then
        metricFields = Array("Imp", "Clicks", "Cost", "Conversion", "Revenue")
        Call AddDdlInOneCell(reportName, "K5", "DailyChartMetricOneChange", metricFields, "chartMetricOne", 1)
    Else
        Sheets(reportName).Shapes("chartMetricOne").ControlFormat.ListIndex = 1
    End If
    
    If Not IsShapeExists(reportName, "chartMetricTwo") Then
        metricFields = Array("Imp", "Clicks", "Cost", "Conversion", "Revenue")
        Call AddDdlInOneCell(reportName, "M5", "DailyChartMetricTwoChange", metricFields, "chartMetricTwo", 2)
    Else
        Sheets(reportName).Shapes("chartMetricTwo").ControlFormat.ListIndex = 2
    End If

End Sub

Sub DailyChartMetricOneChange()
'柱状图的第一个指标项被改变
    Dim valDDL As String
    Dim reportName As String
    Dim srcRng As String
    Dim jCol As Integer
    
    Application.ScreenUpdating = False
    reportName = "Daily"

    valDDL = ValueDDL("chartMetricOne", reportName)
    
    'source
    For jCol = 28 To 40
        If (Cells(1, jCol) = valDDL) Then
            srcRng = Cells(2, jCol).Address & ":" & Cells(Sheets(reportName).[AA65536].End(xlUp).Row, jCol).Address
            Exit For
        End If
    Next

    If (srcRng <> "") Then  'change data source
        Call CopyPasteWithDesign(reportName, srcRng, reportName, "AP2")
    End If

    ActiveSheet.ChartObjects("dailyChart").Activate
    ActiveChart.SeriesCollection(1).name = valDDL
    Sheets(reportName).Select
    ActiveSheet.Range("I8").Select
End Sub

Sub DailyChartMetricTwoChange()
'柱状图的第二个指标项被改变
    Dim valDDL As String
    Dim reportName As String
    Dim srcRng As String
    Dim jCol As Integer
    
    Application.ScreenUpdating = False
    reportName = "Daily"

    valDDL = ValueDDL("chartMetricTwo", reportName)
    
    'source
    For jCol = 28 To 40
        If (Cells(1, jCol) = valDDL) Then
            srcRng = Cells(2, jCol).Address & ":" & Cells(Sheets(reportName).[AA65536].End(xlUp).Row, jCol).Address
            Exit For
        End If
    Next

    If (srcRng <> "") Then  'change data source
        Call CopyPasteWithDesign(reportName, srcRng, reportName, "AQ2")
    End If

    ActiveSheet.ChartObjects("dailyChart").Activate
    ActiveChart.SeriesCollection(2).name = valDDL
    Sheets(reportName).Select
    ActiveSheet.Range("I8").Select
End Sub

Sub DailyChartMetricThreeChange()
'饼图的指标项被改变
    Dim valDDL As String
    Dim reportName As String
    Dim srcRng As String
    Dim jCol As Integer
    Application.ScreenUpdating = False
    reportName = "Daily"

    valDDL = ValueDDL("chartMetricThree", reportName)
    
    'source
    For jCol = 20 To 24
        If (Cells(1, jCol) = valDDL) Then
            srcRng = Cells(2, jCol).Address & ":" & Cells(Sheets(reportName).[S65536].End(xlUp).Row, jCol).Address
            Exit For
        End If
    Next

    If (srcRng <> "") Then
        Call CopyPasteWithDesign(reportName, srcRng, reportName, "Y2")
    End If

    ActiveSheet.ChartObjects("PieDailyChart").Activate
    ActiveChart.SeriesCollection(1).name = valDDL
    Sheets(reportName).Select
    '选择框重置一个
    ActiveSheet.Range("D40").Select

End Sub


Sub TestDailySort()
    


End Sub


