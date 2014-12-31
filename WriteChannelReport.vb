Sub WriteChannelReport()
'生成Channel报表

    Dim dataSheetName As String, reportName As String
    Dim cellChannel As String, cellDate As String
    Dim rowIndex As Integer, colIndex As Integer, i As Integer, j As Integer, rowCount As Integer, dataRowBeginIndex As Integer, dataRowEndIndex As Integer, groupRowCount As Integer
    '定义字典，Channel和Metirc
    Dim dicChannel As Variant, dicMetricData As Variant
    '是否符合channel条件的数据
    Dim filterFlag As Boolean
    '定义渠道数组，日期数组，表格数据数组，多聚到数据数组，多渠道汇总数据数组
    Dim arrChannel As Variant, arrDate As Variant, arr As Variant, arrMutilChannelData As Variant, sumMutilChannelData(1, 1 To 23)
    '指标数据
    Dim fImpression, fClick, fCost, fConversion, fConversion1, fConversion2, fConversion3, fConversion4, fRevenue
    '1元人民币对应的外币
    Dim rate As Double, currencyFormat As Variant
        
    Application.ScreenUpdating = False
    dataSheetName = "Daily元数据"
    reportName = "Channel"
    
    Set st = GetSheet(reportName)
    If st Is Nothing Then
        Exit Sub
    End If
    
    '汇率及货币格式
    rate = Sheets(dataSheetName).Cells(1, 4)
    currencyFormat = Sheets(dataSheetName).Cells(1, 4).NumberFormatLocal
    
    '设置报表数据正常的开始和结束行，不包括渠道汇总
    dataRowBeginIndex = 12
    dataRowEndIndex = 32
    
    '设置单元格所在列格式
    Sheets(reportName).Range("E" & dataRowBeginIndex & ":E" & dataRowEndIndex).NumberFormatLocal = currencyFormat
    Sheets(reportName).Range("G" & dataRowBeginIndex & ":G" & dataRowEndIndex).NumberFormatLocal = currencyFormat
    Sheets(reportName).Range("M" & dataRowBeginIndex & ":Q" & dataRowEndIndex).NumberFormatLocal = currencyFormat
    Sheets(reportName).Range("W" & dataRowBeginIndex & ":W" & dataRowEndIndex).NumberFormatLocal = currencyFormat
    
    arrDate = DDLSourceFromDataColumn("B", dataSheetName, False, 3)
    
    If IsDate(arrDate) Then
    '只有一行数据，构造数组用于绑定列表
        arrDate = Array(Array(arrDate))
    ElseIf IsArray(arrDate) Then
    '格式正确
    Else
        MsgBox "元数据日期列格式不对"
        Exit Sub
    End If

    
    '一天开始日期列表
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlDayStart").ControlFormat.ListIndex = 1
    Sheets(reportName).Shapes("ddlDayStart").OnAction = "DDLChannelDayChanged"
    '一天结束日期列表
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.RemoveAllItems
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.AddItem arrDate
    Sheets(reportName).Shapes("ddlDayEnd").ControlFormat.ListIndex = UBound(arrDate)
    Sheets(reportName).Shapes("ddlDayEnd").OnAction = "DDLChannelDayChanged"
    
    '删除由于渠道汇总所插入的行
    n = Sheets(reportName).UsedRange.Rows.Count
    groupRowCount = n - dataRowEndIndex
    For i = 1 To groupRowCount
        Sheets(reportName).Rows(dataRowBeginIndex).Delete
    Next
    
    '清理原始内容
    Sheets(reportName).Range("B" & dataRowBeginIndex & ":B" & CStr(dataRowEndIndex - 1)).ClearContents
    Sheets(reportName).Range("C" & dataRowBeginIndex & ":X" & dataRowEndIndex).ClearContents
    For rowIndex = dataRowBeginIndex To dataRowEndIndex - 1
        Sheets(reportName).Rows(rowIndex).EntireRow.Hidden = False
        Sheets(reportName).Rows(rowIndex).EntireRow.Interior.ColorIndex = xlNone
        Sheets(reportName).Rows(rowIndex).EntireRow.font.color = vbBlack
        Sheets(reportName).Rows(rowIndex).EntireRow.font.Bold = False
    Next
    
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A3:J" & n)
    
    '指标数据字典
    Set dicMetricData = CreateObject("Scripting.Dictionary")
    '渠道名称字典
    Set dicChannel = CreateObject("Scripting.Dictionary")
    '源表中得到有数据的行
    For i = 1 To UBound(arr)
        cellDate = arr(i, 2)
        If cellDate = "" Then Exit For
        
        If Not IsDate(cellDate) Then
            MsgBox "请检查""" & dataSheetName & """第" & rowIndex & "行数据，必须是日期格式，修正后保存退出，再打开"
            Exit Sub
        End If

        If (Not IsNumeric(arr(i, 3))) Or (Not IsNumeric(arr(i, 4))) Or (Not IsNumeric(arr(i, 5))) Or (Not IsNumeric(arr(i, 6))) Or (Not IsNumeric(arr(i, 7))) Or (Not IsNumeric(arr(i, 8))) Or (Not IsNumeric(arr(i, 9))) Or (Not IsNumeric(arr(i, 10))) Then
            MsgBox "请检查""" & dataSheetName & """第" & rowIndex & "行数据，必须是数字类型格式，修正后保存退出，再打开"
            Exit Sub
        End If
        
        '取得数据中的渠道
        cellChannel = arr(i, 1)
        If dicChannel(cellChannel) = "" Then
            dicChannel(cellChannel) = i
            rowCount = rowCount + 1
        End If
        dicMetricData(cellChannel & "Imp") = dicMetricData(cellChannel & "Imp") + arr(i, 3)
        dicMetricData(cellChannel & "Click") = dicMetricData(cellChannel & "Click") + arr(i, 4)
        dicMetricData(cellChannel & "Cost") = dicMetricData(cellChannel & "Cost") + arr(i, 5) * rate
        dicMetricData(cellChannel & "Conversion1") = dicMetricData(cellChannel & "Conversion1") + arr(i, 6)
        dicMetricData(cellChannel & "Conversion2") = dicMetricData(cellChannel & "Conversion2") + arr(i, 7)
        dicMetricData(cellChannel & "Conversion3") = dicMetricData(cellChannel & "Conversion3") + arr(i, 8)
        dicMetricData(cellChannel & "Conversion4") = dicMetricData(cellChannel & "Conversion4") + arr(i, 9)
        dicMetricData(cellChannel & "Conversion") = dicMetricData(cellChannel & "Conversion") + arr(i, 6) + arr(i, 7) + arr(i, 8) + arr(i, 9)
        dicMetricData(cellChannel & "Revenue") = dicMetricData(cellChannel & "Revenue") + arr(i, 10) * rate
        
    Next

    arrChannel = dicChannel.Keys
    
    '按照Channel对应表汇总排序arrChannel
    Call SortBySubChannel(arrChannel)
    
    '重定义多渠道数据数组
    ReDim arrMutilChannelData(1 To rowCount, 1 To 23)
    
    For i = 1 To rowCount
        arrMutilChannelData(i, 1) = arrChannel(i - 1)
        arrMutilChannelData(i, 2) = dicMetricData(arrMutilChannelData(i, 1) & "Imp")
        arrMutilChannelData(i, 3) = dicMetricData(arrMutilChannelData(i, 1) & "Click")
        arrMutilChannelData(i, 4) = dicMetricData(arrMutilChannelData(i, 1) & "Cost")
        arrMutilChannelData(i, 7) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion1")
        arrMutilChannelData(i, 8) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion2")
        arrMutilChannelData(i, 9) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion3")
        arrMutilChannelData(i, 10) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion4")
        arrMutilChannelData(i, 11) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion")
        arrMutilChannelData(i, 22) = dicMetricData(arrMutilChannelData(i, 1) & "Revenue")
        
        '汇总数据
        sumMutilChannelData(1, 2) = sumMutilChannelData(1, 2) + arrMutilChannelData(i, 2)
        sumMutilChannelData(1, 3) = sumMutilChannelData(1, 3) + arrMutilChannelData(i, 3)
        sumMutilChannelData(1, 4) = sumMutilChannelData(1, 4) + arrMutilChannelData(i, 4)
        sumMutilChannelData(1, 7) = sumMutilChannelData(1, 7) + arrMutilChannelData(i, 7)
        sumMutilChannelData(1, 8) = sumMutilChannelData(1, 8) + arrMutilChannelData(i, 8)
        sumMutilChannelData(1, 9) = sumMutilChannelData(1, 9) + arrMutilChannelData(i, 9)
        sumMutilChannelData(1, 10) = sumMutilChannelData(1, 10) + arrMutilChannelData(i, 10)
        sumMutilChannelData(1, 11) = sumMutilChannelData(1, 11) + arrMutilChannelData(i, 11)
        sumMutilChannelData(1, 22) = sumMutilChannelData(1, 22) + arrMutilChannelData(i, 22)
    Next
    
    Set dicCampaign = Nothing
    Set dicMetricData = Nothing
    
    '计算指标数据
    Call ComMutilChannelData(arrMutilChannelData)
    
    '填充明细数据
    Call FillCellData(reportName, dataRowBeginIndex, 2, arrMutilChannelData)
    
    '计算汇总指标数据
    Call ComMutilChannelData(sumMutilChannelData)
    
    '填充汇总数据
    sumMutilChannelData(1, 1) = Sheets(reportName).Cells(dataRowEndIndex, 2)
    Call FillCellData(reportName, dataRowEndIndex, 2, sumMutilChannelData)
    
    '隐藏没有数据的行
    For rowIndex = dataRowBeginIndex + rowCount To dataRowEndIndex - 1
        Sheets(reportName).Rows(rowIndex).EntireRow.Hidden = True
    Next
    
    '分组汇总
    Call InsertGroupSubChannel
    
    '报表页选中
    Sheets(reportName).Activate

End Sub

Sub DDLChannelDayChanged()
'开始日期或结束日期改变
    Dim startDate As Variant, endDate As Variant
    Dim dataSheetName As String, reportName As String
    Dim cellChannel As String, cellDate As String
    Dim rowIndex As Integer, colIndex As Integer, i As Integer, j As Integer, rowCount As Integer, dataRowBeginIndex As Integer, dataRowEndIndex As Integer, groupRowCount As Integer
    '定义字典，Channel和Metirc
    Dim dicChannel As Variant, dicMetricData As Variant
    '是否符合channel条件的数据
    Dim filterFlag As Boolean
    '定义渠道数组，日期数组，表格数据数组，多聚到数据数组，多渠道汇总数据数组
    Dim arrChannel As Variant, arrDate As Variant, arr As Variant, arrMutilChannelData As Variant, sumMutilChannelData(1, 1 To 23)
    '指标数据
    Dim fImpression, fClick, fCost, fConversion, fConversion1, fConversion2, fConversion3, fConversion4, fRevenue
    '1元人民币对应的外币
    Dim rate As Double, currencyFormat As Variant
        
    Application.ScreenUpdating = False
    dataSheetName = "Daily元数据"
    reportName = "Channel"
    
    '设置报表数据正常的开始和结束行，不包括渠道汇总
    dataRowBeginIndex = 12
    dataRowEndIndex = 32
    
    '汇率及货币格式
    rate = Sheets(dataSheetName).Cells(1, 4)
    
    startDate = ValueDDL("ddlDayStart", reportName)
    endDate = ValueDDL("ddlDayEnd", reportName)
    
    If startDate > endDate Then
        MsgBox "开始日期必须小于结束日期"
        Exit Sub
    End If
    
    '删除由于渠道汇总所插入的行
    n = Sheets(reportName).UsedRange.Rows.Count
    groupRowCount = n - dataRowEndIndex
    For i = 1 To groupRowCount
        Sheets(reportName).Rows(dataRowBeginIndex).Delete
    Next
    
    '清理原始内容，背景色，隐藏取消
    Sheets(reportName).Range("B" & dataRowBeginIndex & ":B" & CStr(dataRowEndIndex - 1)).ClearContents
    Sheets(reportName).Range("C" & dataRowBeginIndex & ":X" & dataRowEndIndex).ClearContents
    For rowIndex = dataRowBeginIndex To dataRowEndIndex - 1
        Sheets(reportName).Rows(rowIndex).EntireRow.Hidden = False
        Sheets(reportName).Rows(rowIndex).EntireRow.Interior.ColorIndex = xlNone
        Sheets(reportName).Rows(rowIndex).EntireRow.font.color = vbBlack
        Sheets(reportName).Rows(rowIndex).EntireRow.font.Bold = False
    Next
    
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A3:J" & n)
    
    '指标数据字典
    Set dicMetricData = CreateObject("Scripting.Dictionary")
    '渠道名称字典
    Set dicChannel = CreateObject("Scripting.Dictionary")
    '源表中得到有数据的行
    For i = 1 To UBound(arr)
        cellDate = arr(i, 2)
        If cellDate = "" Then Exit For
        
        If cellDate >= startDate And cellDate <= endDate Then
            filterFlag = True
        Else
            filterFlag = False
        End If
        
        If filterFlag Then
            '取得数据中的渠道
            cellChannel = arr(i, 1)
            If dicChannel(cellChannel) = "" Then
                dicChannel(cellChannel) = i
                rowCount = rowCount + 1
            End If
            dicMetricData(cellChannel & "Imp") = dicMetricData(cellChannel & "Imp") + arr(i, 3)
            dicMetricData(cellChannel & "Click") = dicMetricData(cellChannel & "Click") + arr(i, 4)
            dicMetricData(cellChannel & "Cost") = dicMetricData(cellChannel & "Cost") + arr(i, 5) * rate
            dicMetricData(cellChannel & "Conversion1") = dicMetricData(cellChannel & "Conversion1") + arr(i, 6)
            dicMetricData(cellChannel & "Conversion2") = dicMetricData(cellChannel & "Conversion2") + arr(i, 7)
            dicMetricData(cellChannel & "Conversion3") = dicMetricData(cellChannel & "Conversion3") + arr(i, 8)
            dicMetricData(cellChannel & "Conversion4") = dicMetricData(cellChannel & "Conversion4") + arr(i, 9)
            dicMetricData(cellChannel & "Conversion") = dicMetricData(cellChannel & "Conversion") + arr(i, 6) + arr(i, 7) + arr(i, 8) + arr(i, 9)
            dicMetricData(cellChannel & "Revenue") = dicMetricData(cellChannel & "Revenue") + arr(i, 10) * rate
        End If
    Next

    arrChannel = dicChannel.Keys
    
    '按照Channel对应表汇总排序arrChannel
    Call SortBySubChannel(arrChannel)
    
    '重定义多渠道数据数组
    ReDim arrMutilChannelData(1 To rowCount, 1 To 23)
    
    For i = 1 To rowCount
        arrMutilChannelData(i, 1) = arrChannel(i - 1)
        arrMutilChannelData(i, 2) = dicMetricData(arrMutilChannelData(i, 1) & "Imp")
        arrMutilChannelData(i, 3) = dicMetricData(arrMutilChannelData(i, 1) & "Click")
        arrMutilChannelData(i, 4) = dicMetricData(arrMutilChannelData(i, 1) & "Cost")
        arrMutilChannelData(i, 7) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion1")
        arrMutilChannelData(i, 8) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion2")
        arrMutilChannelData(i, 9) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion3")
        arrMutilChannelData(i, 10) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion4")
        arrMutilChannelData(i, 11) = dicMetricData(arrMutilChannelData(i, 1) & "Conversion")
        arrMutilChannelData(i, 22) = dicMetricData(arrMutilChannelData(i, 1) & "Revenue")
        
        '汇总数据
        sumMutilChannelData(1, 2) = sumMutilChannelData(1, 2) + arrMutilChannelData(i, 2)
        sumMutilChannelData(1, 3) = sumMutilChannelData(1, 3) + arrMutilChannelData(i, 3)
        sumMutilChannelData(1, 4) = sumMutilChannelData(1, 4) + arrMutilChannelData(i, 4)
        sumMutilChannelData(1, 7) = sumMutilChannelData(1, 7) + arrMutilChannelData(i, 7)
        sumMutilChannelData(1, 8) = sumMutilChannelData(1, 8) + arrMutilChannelData(i, 8)
        sumMutilChannelData(1, 9) = sumMutilChannelData(1, 9) + arrMutilChannelData(i, 9)
        sumMutilChannelData(1, 10) = sumMutilChannelData(1, 10) + arrMutilChannelData(i, 10)
        sumMutilChannelData(1, 11) = sumMutilChannelData(1, 11) + arrMutilChannelData(i, 11)
        sumMutilChannelData(1, 22) = sumMutilChannelData(1, 22) + arrMutilChannelData(i, 22)
    Next
    
    Set dicCampaign = Nothing
    Set dicMetricData = Nothing
    
    '计算指标数据
    Call ComMutilChannelData(arrMutilChannelData)
    
    '填充明细数据
    Call FillCellData(reportName, 12, 2, arrMutilChannelData)
    
    '计算汇总指标数据
    Call ComMutilChannelData(sumMutilChannelData)
    
    '填充汇总数据
    sumMutilChannelData(1, 1) = Sheets(reportName).Cells(dataRowEndIndex, 2)
    Call FillCellData(reportName, dataRowEndIndex, 2, sumMutilChannelData)

    '隐藏没有数据的行
    For rowIndex = dataRowBeginIndex + rowCount To dataRowEndIndex - 1
        Sheets(reportName).Rows(rowIndex).EntireRow.Hidden = True
    Next
    
    '分组汇总
    Call InsertGroupSubChannel
    
    '报表页选中
    Sheets(reportName).Activate

End Sub


Sub ComMutilChannelData(arrData As Variant)
'根据指标元数据，生成相关的计算指标，如点击率，ROI
    Dim rowIndex As Integer, rowCount As Integer
    rowCount = UBound(arrData)
    For rowIndex = 1 To rowCount
    
        'CTR=click/impression
        If arrData(rowIndex, 2) <> "" And arrData(rowIndex, 2) <> 0 Then
            arrData(rowIndex, 5) = arrData(rowIndex, 3) / arrData(rowIndex, 2)
        End If
        
        'CPC=cost/click
        If arrData(rowIndex, 3) <> "" And arrData(rowIndex, 3) <> 0 Then
            arrData(rowIndex, 6) = arrData(rowIndex, 4) / arrData(rowIndex, 3)
        End If
    
        'CPA1=cost/conversion1
        If arrData(rowIndex, 7) <> "" And arrData(rowIndex, 7) <> 0 Then
            arrData(rowIndex, 12) = arrData(rowIndex, 4) / arrData(rowIndex, 7)
        End If
        
        'CPA2=cost/conversion2
        If arrData(rowIndex, 8) <> "" And arrData(rowIndex, 8) <> 0 Then
            arrData(rowIndex, 13) = arrData(rowIndex, 4) / arrData(rowIndex, 8)
        End If
        
        'CPA3=cost/conversion3
        If arrData(rowIndex, 9) <> "" And arrData(rowIndex, 9) <> 0 Then
            arrData(rowIndex, 14) = arrData(rowIndex, 4) / arrData(rowIndex, 9)
        End If
        
        'CPA4=cost/conversion4
        If arrData(rowIndex, 10) <> "" And arrData(rowIndex, 10) <> 0 Then
            arrData(rowIndex, 15) = arrData(rowIndex, 4) / arrData(rowIndex, 10)
        End If
        
        'CPA=cost/conversion
        If arrData(rowIndex, 11) <> "" And arrData(rowIndex, 11) <> 0 Then
            arrData(rowIndex, 16) = arrData(rowIndex, 4) / arrData(rowIndex, 11)
        End If
        
        'Conv.Rate=conversion/click
        If arrData(rowIndex, 3) <> "" And arrData(rowIndex, 3) <> 0 Then
            arrData(rowIndex, 17) = arrData(rowIndex, 7) / arrData(rowIndex, 3)
            arrData(rowIndex, 18) = arrData(rowIndex, 8) / arrData(rowIndex, 3)
            arrData(rowIndex, 19) = arrData(rowIndex, 9) / arrData(rowIndex, 3)
            arrData(rowIndex, 20) = arrData(rowIndex, 10) / arrData(rowIndex, 3)
            arrData(rowIndex, 21) = arrData(rowIndex, 11) / arrData(rowIndex, 3)
        End If
        
        'ROI=Revenue/cost
        If arrData(rowIndex, 4) <> "" And arrData(rowIndex, 4) <> 0 Then
            arrData(rowIndex, 23) = arrData(rowIndex, 22) / arrData(rowIndex, 4)
        End If
        
        
    Next


End Sub

Sub SortBySubChannel(ByRef arrChannel As Variant)
'排序
    Dim dicChannel As Object, dicSubChannel As Object, i As Integer, j As Integer, arrSortChannel As Variant
    Dim dataSheetName As String, n As Integer, arr As Variant, rowCount As Integer, colIndex As Integer
    '大渠道下的小渠道
    Dim subChannelArr As Variant, subChannelCount As Integer, v As Variant, mainChannelCount As Integer
    '小渠道集合
    Dim collSubChannel As New Collection
    
    dataSheetName = "Channel对应表"
    '渠道名称字典
    Set dicChannel = CreateObject("Scripting.Dictionary")
    Set dicSubChannel = CreateObject("Scripting.Dictionary")
    
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A2:B" & n)

    '源表中得到有数据的行
    For i = 1 To UBound(arr)
        '取得数据中的渠道
        cellChannel = arr(i, 1)
        If dicChannel(cellChannel) = "" Then
            dicChannel(cellChannel) = arr(i, 2)
            rowCount = rowCount + 1
        Else
            subChannelArr = Split(dicChannel(cellChannel), ",")
            subChannelIndex = ArrayDataIndex(subChannelArr, arr(i, 2))
            If subChannelIndex = -1 Then
                dicChannel(cellChannel) = dicChannel(cellChannel) & "," & arr(i, 2)
            End If
        End If
    Next
    
    If dicChannel.Count = 0 Then Exit Sub
    
    '获取主渠道下的所有子渠道
    v = dicChannel.Items
    
    For i = 0 To dicChannel.Count - 1
        subChannelArr = Split(v(i), ",")
        For j = 0 To UBound(subChannelArr)
            dicSubChannel(subChannelArr(j)) = 1
            collSubChannel.add subChannelArr(j)
        Next
    Next
    
    If collSubChannel.Count = 0 Then Exit Sub
    
    '渠道的数量，数组从0开始所以加1
    subChannelCount = UBound(arrChannel) + 1
    
    '重定义一个二维数组，用于排序
    ReDim arrSortChannel(1 To subChannelCount, 1 To 2)
    
    For i = 1 To subChannelCount
        arrSortChannel(i, 1) = arrChannel(i - 1)
        colIndex = CollectionDataIndex(collSubChannel, arrChannel(i - 1))
        If colIndex <> -1 Then
            arrSortChannel(i, 2) = colIndex
        Else
            arrSortChannel(i, 2) = 999
        End If
    Next
    
    '渠道按Channel对应表的渠道排序
    Call ArraySort(arrSortChannel, 2, 1)
    
    For i = 1 To subChannelCount
        arrChannel(i - 1) = arrSortChannel(i, 1)
    Next
    
    Set dicChannel = Nothing
    Set dicSubChannel = Nothing
End Sub


Sub InsertGroupSubChannel()
'插入子渠道分组汇总
    Dim dicChannel As Object, dicSubChannel As Object, i As Integer, j As Integer, arrSortChannel As Variant
    Dim reportName As String, dataSheetName As String, n As Integer, arr As Variant, rowCount As Integer, colIndex As Integer
    '大渠道下的小渠道
    Dim subChannelArr As Variant, subChannelCount As Integer, k As Variant, v As Variant, mainChannelCount As Integer
    '小渠道集合
    Dim collSubChannel As New Collection
    '大渠道下的最后一个子渠道
    Dim lastSubChannel As String, lastSubChannelIndex As Integer
    '汇总数据开始行
    Dim dataRowBeginIndex As Integer, dataRowEndIndex As Integer, groupDataRowIndex As Integer, dataRowCount As Integer, groupChannelData As Variant
    
    reportName = "Channel"
    dataSheetName = "Channel对应表"
    '渠道名称字典
    Set dicChannel = CreateObject("Scripting.Dictionary")
    Set dicSubChannel = CreateObject("Scripting.Dictionary")
    
    '源表中得到有数据的行
    n = Sheets(dataSheetName).[A65536].End(xlUp).Row
    arr = Sheets(dataSheetName).Range("A2:B" & n)

    '源表中得到有数据的行
    For i = 1 To UBound(arr)
        '取得数据中的渠道
        cellChannel = arr(i, 1)
        If dicChannel(cellChannel) = "" Then
            dicChannel(cellChannel) = arr(i, 2)
            rowCount = rowCount + 1
        Else
            subChannelArr = Split(dicChannel(cellChannel), ",")
            subChannelIndex = ArrayDataIndex(subChannelArr, arr(i, 2))
            If subChannelIndex = -1 Then
                dicChannel(cellChannel) = dicChannel(cellChannel) & "," & arr(i, 2)
            End If
        End If
    Next
    
    If dicChannel.Count = 0 Then Exit Sub

    '获取主渠道下的所有子渠道
    v = dicChannel.Items
    k = dicChannel.Keys
    dataRowBeginIndex = 12
    dataRowEndIndex = 32
    '报表页选中
    Sheets(reportName).Activate
    
    For i = 0 To dicChannel.Count - 1
        subChannelArr = Split(v(i), ",")
        dataRowCount = dataRowCount + UBound(subChannelArr) + 1
        groupDataRowIndex = dataRowBeginIndex + dataRowCount

        '插入汇总行
        ActiveSheet.Rows(groupDataRowIndex).Insert Shift:=xlDown
        
        
        '设置字体，颜色，背景色
        ActiveSheet.Range("A" & groupDataRowIndex & ":Y" & groupDataRowIndex).Interior.ColorIndex = Sheets(dataSheetName).Cells(1, 1).Interior.ColorIndex
        ActiveSheet.Range("A" & groupDataRowIndex & ":Y" & groupDataRowIndex).font.ColorIndex = Sheets(dataSheetName).Cells(1, 1).font.ColorIndex
        ActiveSheet.Range("A" & groupDataRowIndex & ":Y" & groupDataRowIndex).font.Bold = True
        '计算汇总值
        ReDim groupChannelData(1, 1 To 23)
        
        '分组汇总标题，由组名+Total
        groupChannelData(1, 1) = k(i) & ActiveSheet.Cells(dataRowEndIndex + i + 1, 2)
        For j = groupDataRowIndex - 1 - UBound(subChannelArr) To groupDataRowIndex - 1
            groupChannelData(1, 2) = groupChannelData(1, 2) + ActiveSheet.Cells(j, 3)
            groupChannelData(1, 3) = groupChannelData(1, 3) + ActiveSheet.Cells(j, 4)
            groupChannelData(1, 4) = groupChannelData(1, 4) + ActiveSheet.Cells(j, 5)
            groupChannelData(1, 7) = groupChannelData(1, 7) + ActiveSheet.Cells(j, 8)
            groupChannelData(1, 8) = groupChannelData(1, 8) + ActiveSheet.Cells(j, 9)
            groupChannelData(1, 9) = groupChannelData(1, 9) + ActiveSheet.Cells(j, 10)
            groupChannelData(1, 10) = groupChannelData(1, 10) + ActiveSheet.Cells(j, 11)
            groupChannelData(1, 11) = groupChannelData(1, 11) + ActiveSheet.Cells(j, 12)
            groupChannelData(1, 22) = groupChannelData(1, 22) + ActiveSheet.Cells(j, 23)
        Next
        
        '计算分组汇总指标数据
        Call ComMutilChannelData(groupChannelData)
        
        '填充分组汇总
        Call FillCellData(reportName, groupDataRowIndex, 2, groupChannelData)
        
        '已经增加的数据
        dataRowCount = dataRowCount + 1
    Next
    

    
    Set dicChannel = Nothing
    Set dicSubChannel = Nothing
End Sub
