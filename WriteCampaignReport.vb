Sub WriteCampaignReport()
   
    Dim strChannel As String, strDate As String, strCampaign As String, strYearMonth As String
    Dim i As Integer, j As Integer, rowIndex As Integer, colIndex As Integer, totalColCount As Integer, totalRowCount As Integer
    Dim dataSheetName As String, reportName As String, channelName As String, campaignName As String
    '是否符合campaign,channel条件的数据
    Dim filterFlag As Boolean
    '定义渠道数组
    Dim arrChannels As Variant
    '定义计划数据
    Dim arrCampaigns As Variant
    '定义日期数组
    Dim arrDate As Variant, firstDate As Variant
    '定义多维数组，存放对比的两个月的汇总数据
    Dim twoMonthDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两周的汇总数据
    Dim twoWeekDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两天的数据
    Dim twoDayDataArr(1 To 28, 1 To 2)
    '月份数据，数据中最小的两个月
    Dim twoMonthArr(1 To 2) As Variant, dicMonth As Variant, arrMonth As Variant
    '指标数据
    Dim fImpression, fClick, fCost, fConversion, fConversion1, fConversion2, fConversion3, fConversion4, fRevenue
    '1元人民币对应的外币
    Dim rate As Double, currencyFormat As Variant
    
    Application.ScreenUpdating = False
    
    dataSheetName = "Campaign元数据"
    reportName = "Campaign"
    totalRowCount = Sheets(dataSheetName).UsedRange.Rows.Count
    totalColCount = Sheets(dataSheetName).UsedRange.Columns.Count
    
    Set dicMonth = CreateObject("Scripting.Dictionary")
    
    '汇率及货币格式
    rate = Sheets(dataSheetName).Cells(1, 4)
    currencyFormat = Sheets(dataSheetName).Cells(1, 4).NumberFormatLocal

    '得到数据中最小的两个月份
    For rowIndex = 3 To totalRowCount
        strDate = Sheets(dataSheetName).Cells(rowIndex, 2)
        If strDate = "" Then Exit For

        If Not IsDate(strDate) Then
            MsgBox "请检查""" & dataSheetName & """第" & rowIndex & "行数据，必须是日期格式，修正后保存退出，再打开"
            Exit Sub
        End If
        
        strYearMonth = CStr(Year(strDate)) & "/" & CStr(Month(strDate))
        
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
        
        If dicMonth(strYearMonth) = "" Then dicMonth(strYearMonth) = 1
    Next
    
    '没有数据，清空报表
    If dicMonth.Count = 0 Then
        Call ClearReportData(reportName)
        If IsShapeExists(reportName, "campaignChart") Then
            Sheets(reportName).ChartObjects("campaignChart").Delete
        End If
        Sheets(reportName).Shapes("ddlCampaignChannel").ControlFormat.List = ""
        Sheets(reportName).Shapes("ddlCampaign").ControlFormat.List = ""
        Sheets(reportName).Shapes("ddlCampaignChartChannel").ControlFormat.List = ""
        Sheets(reportName).Shapes("chartMetric1").ControlFormat.List = ""
        Sheets(reportName).Shapes("ddlMonthStart").ControlFormat.List = ""
        Sheets(reportName).Shapes("ddlMonthEnd").ControlFormat.List = ""
        Sheets(reportName).Shapes("ddlWeekStart").ControlFormat.List = ""
        Sheets(reportName).Shapes("ddlWeekEnd").ControlFormat.List = ""
        Sheets(reportName).Shapes("ddlDayStart").ControlFormat.List = ""
        Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.List = ""
        Exit Sub
    End If
    
    If twoMonthArr(2) = "" Then twoMonthArr(2) = twoMonthArr(1)
    
    '对月份数据排序，数组下标0开始
    arrMonth = dicMonth.keys
    Call YearMonthSort(arrMonth)
    Set dicMonth = Nothing
    
    'Chanle下拉列表清空，再赋值
    arrChannels = DDLSourceFromDataColumn("A", dataSheetName, False, 3)
    Sheets(reportName).Shapes("ddlCampaignChannel").ControlFormat.List = "all"
    Sheets(reportName).Shapes("ddlCampaignChannel").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlCampaignChannel").ControlFormat.AddItem arrChannels
    Sheets(reportName).Shapes("ddlCampaignChannel").OnAction = "DDLChannelChanged"
    
    Sheets(reportName).Shapes("ddlCampaignChartChannel").ControlFormat.List = "all"
    Sheets(reportName).Shapes("ddlCampaignChartChannel").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlCampaignChartChannel").ControlFormat.AddItem arrChannels
    Sheets(reportName).Shapes("ddlCampaignChartChannel").OnAction = "CampaignChartDataTypeChange"
    
    'Campaign下拉列表清空，再赋值
    arrCampaigns = DDLSourceFromDataColumn("C", dataSheetName, False, 3)
    Sheets(reportName).Shapes("ddlCampaign").ControlFormat.List = "all"
    Sheets(reportName).Shapes("ddlCampaign").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlCampaign").ControlFormat.AddItem arrCampaigns
    Sheets(reportName).Shapes("ddlCampaign").OnAction = "DDLCampaignChannelChanged"

    '向报表中写入前后对比的两个月份列表
    Sheets(reportName).Shapes("ddlMonthStart").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlMonthStart").ControlFormat.AddItem arrMonth
    Sheets(reportName).Shapes("ddlMonthStart").ControlFormat.ListIndex = ArrayDataIndex(arrMonth, twoMonthArr(1)) + 1
    Sheets(reportName).Shapes("ddlMonthStart").OnAction = "DDLCampaignChannelChanged"

    Sheets(reportName).Shapes("ddlMonthEnd").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlMonthEnd").ControlFormat.AddItem arrMonth
    Sheets(reportName).Shapes("ddlMonthEnd").ControlFormat.ListIndex = ArrayDataIndex(arrMonth, twoMonthArr(2)) + 1
    Sheets(reportName).Shapes("ddlMonthEnd").OnAction = "DDLCampaignChannelChanged"

    arrDate = DDLSourceFromDataColumn("B", dataSheetName, False, 3)
    
    If IsDate(arrDate) Then
    '只有一行数据，构造数组用于绑定列表
        firstDate = arrDate
        arrDate = Array(Array(arrDate))
    ElseIf IsArray(arrDate) Then
        firstDate = arrDate(1, 1)
    Else
        MsgBox "元数据日期列格式不对"
        Exit Sub
    End If
    
    '一周开始日期列表
    Sheets(reportName).Shapes("ddlWeekStart").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlWeekStart").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlWeekStart").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlWeekStart").OnAction = "DDLCampaignChannelChanged"
    '一周结束日期列表
    Sheets(reportName).Shapes("ddlWeekEnd").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlWeekEnd").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlWeekEnd").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlWeekEnd").OnAction = "DDLCampaignChannelChanged"
    '一天开始日期列表
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlDayStart").OnAction = "DDLCampaignChannelChanged"
    '一天结束日期列表
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlDayEnd").OnAction = "DDLCampaignChannelChanged"

    '初始条件
    channelName = "all"
    campaignName = "all"

    '汇总数据
    For rowIndex = 3 To totalRowCount

        strChannel = Sheets(dataSheetName).Cells(rowIndex, 1)
        strDate = CStr(Sheets(dataSheetName).Cells(rowIndex, 2))
        
        If strDate = "" Then Exit For
        
        strCampaign = Sheets(dataSheetName).Cells(rowIndex, 3)
        fImpression = Sheets(dataSheetName).Cells(rowIndex, 4)
        fClick = Sheets(dataSheetName).Cells(rowIndex, 5)
        fCost = Sheets(dataSheetName).Cells(rowIndex, 6) * rate
        fConversion1 = Sheets(dataSheetName).Cells(rowIndex, 7)
        fConversion2 = Sheets(dataSheetName).Cells(rowIndex, 8)
        fConversion3 = Sheets(dataSheetName).Cells(rowIndex, 9)
        fConversion4 = Sheets(dataSheetName).Cells(rowIndex, 10)
        fConversion = fConversion1 + fConversion2 + fConversion3 + fConversion4
        fRevenue = Sheets(dataSheetName).Cells(rowIndex, 11) * rate

        If Not IsDate(strDate) Then
            MsgBox "请检查""" & dataSheetName & """第" & rowIndex & "行数据，必须是日期格式，修正后保存退出，再打开"
            Exit Sub
        End If

        If (Not IsNumeric(fImpression)) Or (Not IsNumeric(fClick)) Or (Not IsNumeric(fCost)) Or (Not IsNumeric(fConversion1)) Or (Not IsNumeric(fConversion2)) Or (Not IsNumeric(fConversion3)) Or (Not IsNumeric(fConversion4)) Or (Not IsNumeric(fRevenue)) Then
            MsgBox "请检查""" & dataSheetName & """第" & rowIndex & "行数据，必须是数字类型格式，修正后保存退出，再打开"
            Exit Sub
        End If

        If (channelName = strChannel Or channelName = "" Or channelName = "all") And (campaignName = strCampaign Or campaignName = "" Or campaignName = "all") Then
            filterFlag = True
        Else
            filterFlag = False
        End If

        If filterFlag Then
            '两月对比数据
            strYearMonth = CStr(Year(strDate)) & "/" & CStr(Month(strDate))
            If strYearMonth = twoMonthArr(1) Then
                Call FillArrayData(twoMonthDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If
            If strYearMonth = twoMonthArr(2) Then
                Call FillArrayData(twoMonthDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If

            '两周对比数据
            If IsInWeek(CStr(firstDate), strDate) Then
                Call FillArrayData(twoWeekDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
                Call FillArrayData(twoWeekDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If

            '两天对比数据
            If CStr(firstDate) = strDate Then
                Call FillArrayData(twoDayDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
                Call FillArrayData(twoDayDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If

        End If

    Next

    '向报表中写入前后对比的两个月份数据
    Call ComMetricArrayData(twoMonthDataArr)
    Call FillCellData(reportName, 9, 4, twoMonthDataArr)

    '向报表中写入前后对比的两周数据
    Call ComMetricArrayData(twoWeekDataArr)
    Call FillCellData(reportName, 40, 4, twoWeekDataArr)

    '向报表中写入前后对比的两天数据
    Call ComMetricArrayData(twoDayDataArr)
    Call FillCellData(reportName, 40, 9, twoDayDataArr)
    
    '设置单个格货币格式
    Call SetCurrencyFormat(reportName, 9, 4, currencyFormat)
    Call SetCurrencyFormat(reportName, 40, 4, currencyFormat)
    Call SetCurrencyFormat(reportName, 40, 9, currencyFormat)

    '绘制饼图
    Call ChartForCampaign
End Sub

Sub DDLCampaignChannelChanged()
        
    Dim fImpression, fClick, fCost, fConversion, fConversion1, fConversion2, fConversion3, fConversion4, fRevenue
    Dim strChannel As String, strDate As String, strCampaign As String, strYearMonth As String
    Dim channelName As String, campaignName As String, dataSheetName As String, reportName As String
    Dim monthStart As String, monthEnd As String
    Dim weekStart As Date, weekEnd As Date, dayStart As Date, dayEnd As Date
    Dim i As Integer, j As Integer, totalRowCount As Integer, rowIndex As Integer
    '是否符合campaign,channel条件的数据
    Dim filterFlag As Boolean
    '定义多维数组，存放对比的两个月的汇总数据
    Dim twoMonthDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两周的汇总数据
    Dim twoWeekDataArr(1 To 28, 1 To 2)
    '定义多维数组，存放对比的两天的数据
    Dim twoDayDataArr(1 To 28, 1 To 2)
    '1元人民币对应的外币
    Dim rate As Double
    
    Application.ScreenUpdating = False
    filterFlag = True
    dataSheetName = "Campaign元数据"
    reportName = "Campaign"
    '汇率
    rate = Sheets(dataSheetName).Cells(1, 4)
    
    'Chanle，Campaign，月份被改变
    channelName = ValueDDL("ddlCampaignChannel", reportName)
    campaignName = ValueDDL("ddlCampaign", reportName)
    monthStart = ValueDDL("ddlMonthStart", reportName)
    monthEnd = ValueDDL("ddlMonthEnd", reportName)
    weekStart = ValueDDL("ddlWeekStart", reportName)
    weekEnd = ValueDDL("ddlWeekEnd", reportName)
    dayStart = ValueDDL("ddlDayStart", reportName)
    dayEnd = ValueDDL("ddlDayEnd", reportName)
    totalRowCount = Sheets(dataSheetName).UsedRange.Rows.Count
    
    For rowIndex = 3 To totalRowCount
        
        strChannel = Sheets(dataSheetName).Cells(rowIndex, 1)
        strDate = CStr(Sheets(dataSheetName).Cells(rowIndex, 2))
        If strDate = "" Then Exit For
        strCampaign = Sheets(dataSheetName).Cells(rowIndex, 3)
        fImpression = Sheets(dataSheetName).Cells(rowIndex, 4)
        fClick = Sheets(dataSheetName).Cells(rowIndex, 5)
        fCost = Sheets(dataSheetName).Cells(rowIndex, 6) * rate
        fConversion1 = Sheets(dataSheetName).Cells(rowIndex, 7)
        fConversion2 = Sheets(dataSheetName).Cells(rowIndex, 8)
        fConversion3 = Sheets(dataSheetName).Cells(rowIndex, 9)
        fConversion4 = Sheets(dataSheetName).Cells(rowIndex, 10)
        fConversion = fConversion1 + fConversion2 + fConversion3 + fConversion4
        fRevenue = Sheets(dataSheetName).Cells(rowIndex, 11) * rate
    
        If (channelName = strChannel Or channelName = "" Or channelName = "all") And (campaignName = strCampaign Or campaignName = "" Or campaignName = "all") Then
            filterFlag = True
        Else
            filterFlag = False
        End If

        If filterFlag = True Then
            
            '两月对比数据
            strYearMonth = CStr(Year(strDate)) & "/" & CStr(Month(strDate))
            If strYearMonth = monthStart Then
                Call FillArrayData(twoMonthDataArr, 1, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
            End If

            If strYearMonth = monthEnd Then
                Call FillArrayData(twoMonthDataArr, 2, fImpression, fClick, fConversion1, fConversion2, fConversion3, fConversion4, fConversion, fCost, fRevenue)
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
        
    Next
    
    
    '向报表中写入前后对比的两个月份数据
    Call ComMetricArrayData(twoMonthDataArr)
    Call FillCellData(reportName, 9, 4, twoMonthDataArr)

    '向报表中写入前后对比的两周数据
    Call ComMetricArrayData(twoWeekDataArr)
    Call FillCellData(reportName, 40, 4, twoWeekDataArr)

    '向报表中写入前后对比的两天数据
    Call ComMetricArrayData(twoDayDataArr)
    Call FillCellData(reportName, 40, 9, twoDayDataArr)
    
End Sub

Sub ChartForCampaign()
'生成饼图
    Dim metricFields() As Variant
    Dim i As Long, n As Long, j As Long, rowCount As Integer
    Dim dicMetricData, dicCampaign, arrKeys, arr
    Dim dataSheetName As String, reportName As String
    '1元人民币对应的外币
    Dim rate As Double
    
    
    Application.ScreenUpdating = False
    
    dataSheetName = "Campaign元数据"
    reportName = "Campaign"
    
    '汇率
    rate = Sheets(dataSheetName).Cells(1, 4)
    rowCount = 0
    
    '得到有数据的行，需要清理
    n = Sheets(reportName).[AA65536].End(xlUp).Row
    If n > 1 Then
        Sheets(reportName).Range("AA2:AQ" & n).ClearContents
    End If
    '指标数据字典
    Set dicMetricData = CreateObject("Scripting.Dictionary")
    '计划名称字典
    Set dicCampaign = CreateObject("Scripting.Dictionary")
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A3:K" & n)
    
    For i = 1 To UBound(arr)
        If dicCampaign(arr(i, 3)) = "" Then
            dicCampaign(arr(i, 3)) = i
            rowCount = rowCount + 1
        End If
        dicMetricData(arr(i, 3) & "Imp") = dicMetricData(arr(i, 3) & "Imp") + arr(i, 4)
        dicMetricData(arr(i, 3) & "Click") = dicMetricData(arr(i, 3) & "Click") + arr(i, 5)
        dicMetricData(arr(i, 3) & "Cost") = dicMetricData(arr(i, 3) & "Cost") + arr(i, 6) * rate
        dicMetricData(arr(i, 3) & "Conversion") = dicMetricData(arr(i, 3) & "Conversion") + arr(i, 7) + arr(i, 8) + arr(i, 9) + arr(i, 10)
        dicMetricData(arr(i, 3) & "Revenue") = dicMetricData(arr(i, 3) & "Revenue") + arr(i, 11) * rate
    Next
    
    Erase arr
    '构造饼图的数据源
    arr = Sheets(reportName).Range("AA2", Sheets(reportName).Cells(rowCount + 1, 32))
    
    arrKeys = dicCampaign.keys
    
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
    Set dicMetricData = Nothing
    Set dicCampaign = Nothing
    
    '图形的数据源，一列
    Call CopyPasteWithDesign(reportName, "AB2:AB" & rowCount + 1, reportName, "ap2")
    
    '没有饼图重新创建，有的话使用老格式
    If Not IsShapeExists(reportName, "campaignChart") Then
        Call Make3DPieChart(reportName, "AA2:AA" & (rowCount + 1) & ",AP2:AP" & (rowCount + 1), reportName, Range("H5").left, Range("H5").top, 500, 250, "campaignChart", "All Campaign", True)
    Else
        Sheets(reportName).ChartObjects("campaignChart").Chart.SetSourceData (Sheets(reportName).Range("AA2:AA" & (rowCount + 1) & ",AP2:AP" & (rowCount + 1)))
    End If
    
    '设置在图表中显示隐藏数据
    Sheets(reportName).ChartObjects("campaignChart").Chart.PlotVisibleOnly = False
    
    If Not IsShapeExists(reportName, "chartMetric1") Then
        metricFields = Array("Imp", "Clicks", "Cost", "Conversion", "Revenue")
        Call AddDdlInOneCell(reportName, "H5", "CampaignChartMetricChange", metricFields, "chartMetric1", 1)
    Else
        If Sheets(reportName).Shapes("chartMetric1").ControlFormat.ListCount = 1 Then
            Sheets(reportName).Shapes("chartMetric1").ControlFormat.RemoveAllItems
            Sheets(reportName).Shapes("chartMetric1").ControlFormat.AddItem Array("Imp", "Clicks", "Cost", "Conversion", "Revenue")
        End If
        Sheets(reportName).Shapes("chartMetric1").ControlFormat.ListIndex = 1
    End If
    
    Sheets(reportName).Shapes("chartMetric1").ZOrder msoBringToFront
    Sheets(reportName).Shapes("ddlCampaignChartChannel").ZOrder msoBringToFront
    ActiveSheet.Range("I8").Select
End Sub

Sub CampaignChartDataTypeChange()
'图形选项Channel被选择

    Dim metricFields() As Variant
    Dim i As Long, n As Long, j As Long, rowCount As Integer
    Dim dicMetricData, dicDate, arr, arrKeys
    Dim dataSheetName As String, reportName As String, channelName As String, cellChannel As String
    Dim cellDate As Variant, cellYearMonth As String
    Dim rate As Double
    Dim filterFlag As Boolean
    
    Application.ScreenUpdating = False
    dataSheetName = "Campaign元数据"
    reportName = "Campaign"
    
    '汇率
    rate = Sheets(dataSheetName).Cells(1, 4)
    
    '得到有数据的行，需要清理
    n = Sheets(reportName).[AA65536].End(xlUp).Row
    If n > 1 Then
        Sheets(reportName).Range("AA2:AQ" & n).ClearContents
    End If
    
    channelName = ValueDDL("ddlCampaignChartChannel", reportName)
    
    '指标数据字典
    Set dicMetricData = CreateObject("Scripting.Dictionary")
    '计划名称字典
    Set dicCampaign = CreateObject("Scripting.Dictionary")
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A3:K" & n)
    
    For i = 3 To UBound(arr)
        '取得数据中的渠道
        cellChannel = arr(i, 1)
        
        If (channelName = cellChannel Or channelName = "all") Then
            filterFlag = True
        Else
            filterFlag = False
        End If
        
        If filterFlag Then
        
            If dicCampaign(arr(i, 3)) = "" Then
                dicCampaign(arr(i, 3)) = i
                rowCount = rowCount + 1
            End If
            dicMetricData(arr(i, 3) & "Imp") = dicMetricData(arr(i, 3) & "Imp") + arr(i, 4)
            dicMetricData(arr(i, 3) & "Click") = dicMetricData(arr(i, 3) & "Click") + arr(i, 5)
            dicMetricData(arr(i, 3) & "Cost") = dicMetricData(arr(i, 3) & "Cost") + arr(i, 6) * rate
            dicMetricData(arr(i, 3) & "Conversion") = dicMetricData(arr(i, 3) & "Conversion") + arr(i, 7) + arr(i, 8) + arr(i, 9) + arr(i, 10)
            dicMetricData(arr(i, 3) & "Revenue") = dicMetricData(arr(i, 3) & "Revenue") + arr(i, 11) * rate
            
        End If
    Next
    
    Erase arr
    arr = Sheets(reportName).Range("AA2", Sheets(reportName).Cells(rowCount + 1, 32))
    
    arrKeys = dicCampaign.keys
    
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
    Set dicCampaign = Nothing
    Set dicMetricData = Nothing
    
    '图形的数据源，一列
    Call CopyPasteWithDesign(reportName, "AB2:AB" & rowCount + 1, reportName, "ap2")
    
    '没有饼图重新创建，有的话使用老格式
    If Not IsShapeExists(reportName, "campaignChart") Then
        Call Make3DPieChart(reportName, "AA2:AA" & (rowCount + 1) & ",AP2:AP" & (rowCount + 1), reportName, Range("H5").left, Range("H5").top, 500, 250, "campaignChart", "All Campaign", True)
    Else
        Sheets(reportName).ChartObjects("campaignChart").Chart.SetSourceData (Sheets(reportName).Range("AA2:AA" & (rowCount + 1) & ",AP2:AP" & (rowCount + 1)))
    End If
    
    '设置在图表中显示隐藏数据
    Sheets(reportName).ChartObjects("campaignChart").Chart.PlotVisibleOnly = False
    
    If Not IsShapeExists(reportName, "chartMetric1") Then
        metricFields = Array("Imp", "Clicks", "Cost", "Conversion", "Revenue")
        Call AddDdlInOneCell(reportName, "H5", "CampaignChartMetricChange", metricFields, "chartMetric1", 1)
    Else
        Sheets(reportName).Shapes("chartMetric1").ControlFormat.ListIndex = 1
    End If
    
    Sheets(reportName).Shapes("chartMetric1").ZOrder msoBringToFront
    Sheets(reportName).Shapes("ddlCampaignChartChannel").ZOrder msoBringToFront
    ActiveSheet.Range("I8").Select

End Sub

Sub CampaignChartMetricChange()
    Dim valDDL As String, reportName As String
    Dim srcRng As String
    Dim jCol As Integer
    
    
    Application.ScreenUpdating = False
    reportName = "Campaign"
    valDDL = ValueDDL("chartMetric1", reportName)
    
    For jCol = 28 To 40
        If (Cells(1, jCol) = valDDL) Then
            srcRng = Cells(2, jCol).Address & ":" & Cells(Sheets(reportName).[AA65536].End(xlUp).Row, jCol).Address
            Exit For
        End If
    Next

    If (srcRng <> "") Then
        Call CopyPasteWithDesign(reportName, srcRng, reportName, "AP2")
    End If

    ActiveSheet.ChartObjects("campaignChart").Activate
    ActiveChart.SeriesCollection(1).name = valDDL
    ActiveSheet.Range("I8").Select
End Sub


Sub Test()
    

End Sub

Sub DDLChannelChanged()
'Channel改变后，Campaign联动
    Dim reportName As String, dataSheetName As String, channelName As String, campaignName As String
    Dim strChannel As String, strCampaign As String
    Dim dicCampaign, arrCampaigns
    Dim rowIndex As Integer, totalRowCount As Integer
    
    reportName = "Campaign"
    dataSheetName = "Campaign元数据"
    totalRowCount = Sheets(dataSheetName).UsedRange.Rows.Count
    
    channelName = ValueDDL("ddlCampaignChannel", reportName)
    Set dicCampaign = CreateObject("Scripting.Dictionary")
    
    '选择all时，Campaign不做任何筛选

    For rowIndex = 3 To totalRowCount
        strChannel = Sheets(dataSheetName).Cells(rowIndex, 1)
        strCampaign = Sheets(dataSheetName).Cells(rowIndex, 3)
        If strChannel = "" Or strCampaign = "" Then Exit For
        If (channelName = "all") Or (channelName <> "all" And channelName = strChannel) Then
            If dicCampaign(strCampaign) = "" Then
                dicCampaign(strCampaign) = 1
            End If
        End If
    Next
    
    '没有数据，退出
    If dicCampaign.Count = 0 Then
        Exit Sub
    End If
    
    arrCampaigns = dicCampaign.keys
    
    Sheets(reportName).Shapes("ddlCampaign").ControlFormat.List = "all"
    Sheets(reportName).Shapes("ddlCampaign").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlCampaign").ControlFormat.AddItem arrCampaigns

    
    Set dicCampaign = Nothing

    '更新数据
    Call DDLCampaignChannelChanged
    
End Sub
