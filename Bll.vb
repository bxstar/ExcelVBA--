Sub ComMetricArrayData(arrData As Variant)
'根据指标元数据，生成相关的计算指标，如点击率
    'CTR
    If arrData(1, 1) <> "" And arrData(1, 1) <> 0 Then
        arrData(3, 1) = arrData(2, 1) / arrData(1, 1)
    End If
    If arrData(1, 2) <> "" And arrData(1, 2) <> 0 Then
        arrData(3, 2) = arrData(2, 2) / arrData(1, 2)
    End If
    'Cost/Conv1.
    If arrData(4, 1) <> "" And arrData(4, 1) <> 0 Then
        arrData(10, 1) = arrData(9, 1) / arrData(4, 1)
    End If
    If arrData(4, 2) <> "" And arrData(4, 2) <> 0 Then
        arrData(10, 2) = arrData(9, 2) / arrData(4, 2)
    End If
    'Cost/Conv2.
    If arrData(5, 1) <> "" And arrData(5, 1) <> 0 Then
        arrData(11, 1) = arrData(9, 1) / arrData(5, 1)
    End If
    If arrData(5, 2) <> "" And arrData(5, 2) <> 0 Then
        arrData(11, 2) = arrData(9, 2) / arrData(5, 2)
    End If
    'Cost/Conv3.
    If arrData(6, 1) <> "" And arrData(6, 1) <> 0 Then
        arrData(12, 1) = arrData(9, 1) / arrData(6, 1)
    End If
    If arrData(6, 2) <> "" And arrData(6, 2) <> 0 Then
        arrData(12, 2) = arrData(9, 2) / arrData(6, 2)
    End If
    'Cost/Conv4.
    If arrData(7, 1) <> "" And arrData(7, 1) <> 0 Then
        arrData(13, 1) = arrData(9, 1) / arrData(7, 1)
    End If
    If arrData(7, 2) <> "" And arrData(7, 2) <> 0 Then
        arrData(13, 2) = arrData(9, 2) / arrData(7, 2)
    End If
    'Cost/Conv.
    If arrData(8, 1) <> "" And arrData(8, 1) <> 0 Then
        arrData(14, 1) = arrData(9, 1) / arrData(8, 1)
    End If
    If arrData(8, 2) <> "" And arrData(8, 2) <> 0 Then
        arrData(14, 2) = arrData(9, 2) / arrData(8, 2)
    End If
    'Conv1. Rate
    If arrData(2, 1) <> "" And arrData(2, 1) <> 0 Then
        arrData(15, 1) = arrData(4, 1) / arrData(2, 1)
    End If
    If arrData(2, 2) <> "" And arrData(2, 2) <> 0 Then
        arrData(15, 2) = arrData(4, 2) / arrData(2, 2)
    End If
    'Conv2. Rate
    If arrData(2, 1) <> "" And arrData(2, 1) <> 0 Then
        arrData(16, 1) = arrData(5, 1) / arrData(2, 1)
    End If
    If arrData(2, 2) <> "" And arrData(2, 2) <> 0 Then
        arrData(16, 2) = arrData(5, 2) / arrData(2, 2)
    End If
    'Conv3. Rate
    If arrData(2, 1) <> "" And arrData(2, 1) <> 0 Then
        arrData(17, 1) = arrData(6, 1) / arrData(2, 1)
    End If
    If arrData(2, 2) <> "" And arrData(2, 2) <> 0 Then
        arrData(17, 2) = arrData(6, 2) / arrData(2, 2)
    End If
    'Conv4. Rate
    If arrData(2, 1) <> "" And arrData(2, 1) <> 0 Then
        arrData(18, 1) = arrData(7, 1) / arrData(2, 1)
    End If
    If arrData(2, 2) <> "" And arrData(2, 2) <> 0 Then
        arrData(18, 2) = arrData(7, 2) / arrData(2, 2)
    End If
    'Conv. Rate
    If arrData(2, 1) <> "" And arrData(2, 1) <> 0 Then
        arrData(19, 1) = arrData(8, 1) / arrData(2, 1)
    End If
    If arrData(2, 2) <> "" And arrData(2, 2) <> 0 Then
        arrData(19, 2) = arrData(8, 2) / arrData(2, 2)
    End If
    'Rev./Conv1.
    If arrData(4, 1) <> "" And arrData(4, 1) <> 0 Then
        arrData(21, 1) = arrData(20, 1) / arrData(4, 1)
    End If
    If arrData(4, 2) <> "" And arrData(4, 2) <> 0 Then
        arrData(21, 2) = arrData(20, 2) / arrData(4, 2)
    End If
    'Rev./Conv2.
    If arrData(5, 1) <> "" And arrData(5, 1) <> 0 Then
        arrData(22, 1) = arrData(20, 1) / arrData(5, 1)
    End If
    If arrData(5, 2) <> "" And arrData(5, 2) <> 0 Then
        arrData(22, 2) = arrData(20, 2) / arrData(5, 2)
    End If
    'Rev./Conv3.
    If arrData(6, 1) <> "" And arrData(6, 1) <> 0 Then
        arrData(23, 1) = arrData(20, 1) / arrData(6, 1)
    End If
    If arrData(6, 2) <> "" And arrData(6, 2) <> 0 Then
        arrData(23, 2) = arrData(20, 2) / arrData(6, 2)
    End If
    'Rev./Conv4.
    If arrData(7, 1) <> "" And arrData(7, 1) <> 0 Then
        arrData(24, 1) = arrData(20, 1) / arrData(7, 1)
    End If
    If arrData(7, 2) <> "" And arrData(7, 2) <> 0 Then
        arrData(24, 2) = arrData(20, 2) / arrData(7, 2)
    End If
    'Rev./Conv.
    If arrData(8, 1) <> "" And arrData(8, 1) <> 0 Then
        arrData(25, 1) = arrData(20, 1) / arrData(8, 1)
    End If
    If arrData(8, 2) <> "" And arrData(8, 2) <> 0 Then
        arrData(25, 2) = arrData(20, 2) / arrData(8, 2)
    End If
    'ROI
    If arrData(9, 1) <> "" And arrData(9, 1) <> 0 Then
        arrData(26, 1) = arrData(20, 1) / arrData(9, 1)
    End If
    If arrData(9, 2) <> "" And arrData(9, 2) <> 0 Then
        arrData(26, 2) = arrData(20, 2) / arrData(9, 2)
    End If
    'CPC
    If arrData(2, 1) <> "" And arrData(2, 1) <> 0 Then
        arrData(28, 1) = arrData(9, 1) / arrData(2, 1)
    End If
    If arrData(2, 2) <> "" And arrData(2, 2) <> 0 Then
        arrData(28, 2) = arrData(9, 2) / arrData(2, 2)
    End If
End Sub


Sub FillArrayData(arrData As Variant, dataColumnIndex As Integer, impValue As Variant, clickValue As Variant, conv1Value As Variant, conv2Value As Variant, conv3Value As Variant, conv4Value As Variant, convValue As Variant, costValue As Variant, revValue As Variant)
'将指标数据填充进多维数组，dataColumnIndex表示写入多维数组的第几列
    arrData(1, dataColumnIndex) = AddValue(arrData(1, dataColumnIndex), impValue)
    arrData(2, dataColumnIndex) = AddValue(arrData(2, dataColumnIndex), clickValue)
    arrData(4, dataColumnIndex) = AddValue(arrData(4, dataColumnIndex), conv1Value)
    arrData(5, dataColumnIndex) = AddValue(arrData(5, dataColumnIndex), conv2Value)
    arrData(6, dataColumnIndex) = AddValue(arrData(6, dataColumnIndex), conv3Value)
    arrData(7, dataColumnIndex) = AddValue(arrData(7, dataColumnIndex), conv4Value)
    arrData(8, dataColumnIndex) = AddValue(arrData(8, dataColumnIndex), convValue)
    arrData(9, dataColumnIndex) = AddValue(arrData(9, dataColumnIndex), costValue)
    arrData(20, dataColumnIndex) = AddValue(arrData(20, dataColumnIndex), revValue)
End Sub

Sub FillCellData(reportName As String, startRowIndex As Integer, startColumnIndex As Integer, arrData As Variant)
'将多维数组的数据填充进单元格
    Dim rowCount As Integer, columnCount As Integer, i As Integer, j As Integer
    
    rowCount = UBound(arrData, 1)
    columnCount = UBound(arrData, 2)
    
    For i = 0 To rowCount - 1
        For j = 0 To columnCount - 1
            Sheets(reportName).Cells(startRowIndex + i, startColumnIndex + j) = arrData(i + 1, j + 1)
        Next
    Next

End Sub

Sub SetCurrencyFormat(sheetName As String, startRowIndex As Integer, startColumnIndex As Integer, format As Variant)
'按相对位置设置货币单元格的格式

    'Cost
    Sheets(sheetName).Cells(startRowIndex + 8, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 8, startColumnIndex + 1).NumberFormatLocal = format
    
    'Cost/Conv
    Sheets(sheetName).Cells(startRowIndex + 9, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 9, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 10, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 10, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 11, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 11, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 12, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 12, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 13, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 13, startColumnIndex + 1).NumberFormatLocal = format
    
    'Rev.
    Sheets(sheetName).Cells(startRowIndex + 19, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 19, startColumnIndex + 1).NumberFormatLocal = format
    
    'Rev./Conv.
    Sheets(sheetName).Cells(startRowIndex + 20, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 20, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 21, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 21, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 22, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 22, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 23, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 23, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 24, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 24, startColumnIndex + 1).NumberFormatLocal = format
    
    'Avg. CPC
    Sheets(sheetName).Cells(startRowIndex + 27, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 27, startColumnIndex + 1).NumberFormatLocal = format

End Sub

Sub ClearReportData(sheetName As String)
'清除报表数据
    If sheetName = "Campaign" Then
        Sheets(sheetName).Range("D9:E36").ClearContents
        Sheets(sheetName).Range("D40:E67").ClearContents
        Sheets(sheetName).Range("I40:J67").ClearContents
        
        If IsShapeExists(sheetName, "campaignChart") Then
            Sheets(sheetName).ChartObjects("campaignChart").Delete
        End If
        Sheets(sheetName).Shapes("ddlCampaignChannel").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlCampaign").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlCampaignChartChannel").ControlFormat.List = ""
        Sheets(sheetName).Shapes("chartMetric1").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlMonthStart").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlMonthEnd").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlWeekStart").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlWeekEnd").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlDayStart").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlDayEnd").ControlFormat.List = ""
        
    ElseIf sheetName = "Daily" Then
        Sheets(sheetName).Range("D9:E36").ClearContents
        Sheets(sheetName).Range("D40:E67").ClearContents
        Sheets(sheetName).Range("I40:J67").ClearContents
        Sheets(sheetName).Range("N40:O67").ClearContents
        
        If IsShapeExists(sheetName, "dailyChart") Then
            Sheets(sheetName).ChartObjects("dailyChart").Delete
        End If
        
        If IsShapeExists(sheetName, "PieDailyChart") Then
            Sheets(sheetName).ChartObjects("PieDailyChart").Delete
        End If
        
        Sheets(sheetName).Shapes("ddlDailyChannel").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlDailyChartChannel").ControlFormat.List = ""
        Sheets(sheetName).Shapes("chartDataType").ControlFormat.List = ""
        Sheets(sheetName).Shapes("chartMetricOne").ControlFormat.List = ""
        Sheets(sheetName).Shapes("chartMetricTwo").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlMonthStart").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlMonthEnd").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlWeekStart").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlWeekEnd").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlDayStart").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlDayEnd").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlYearStart").ControlFormat.List = ""
        Sheets(sheetName).Shapes("ddlYearEnd").ControlFormat.List = ""
        Sheets(sheetName).Shapes("chartMetricThree").ControlFormat.List = ""
    End If
End Sub


Function GetCellCurrencyFormat(cellText As Variant)
'获取单元格的货币格式
    Dim format As String, strText As String, strChar As String, charIndex As Integer, charStartIndex As Integer, charEndIndex As Integer

    '获取货币格式开始索引
    For charIndex = 1 To Len(cellText)
        strChar = Right(left(cellText, charIndex), 1)
        If strChar <> " " Then
            charStartIndex = charIndex
            Exit For
        End If
    Next
    
    '获取货币格式结束索引
    For charIndex = (charStartIndex + 1) To Len(cellText)
        strChar = Right(left(cellText, charIndex), 1)
        If strChar = " " Or IsNumeric(strChar) Then
            charEndIndex = charIndex
            Exit For
        End If
    Next
    
    '货币格式
    format = Mid(cellText, charStartIndex, charEndIndex - charStartIndex)
    
    '自定义格式
    'Selection.NumberFormatLocal = format & "#,##0.00;-" & format & "#,##0.00"


End Function