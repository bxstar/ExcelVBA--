Sub ComMetricArrayData(arrData As Variant)
'根据指标元数据，生成相关的计算指标，如点击率
    'CTR
    If arrData(1, 1) <> "" And arrData(1, 1) <> 0 Then
        arrData(3, 1) = arrData(2, 1) / arrData(1, 1)
    End If
    If arrData(1, 2) <> "" And arrData(1, 2) <> 0 Then
        arrData(3, 2) = arrData(2, 2) / arrData(1, 2)
    End If
    'Cost/Conv.
    If arrData(4, 1) <> "" And arrData(4, 1) <> 0 Then
        arrData(6, 1) = arrData(5, 1) / arrData(4, 1)
    End If
    If arrData(4, 2) <> "" And arrData(4, 2) <> 0 Then
        arrData(6, 2) = arrData(5, 2) / arrData(4, 2)
    End If
    'Conv. Rate
    If arrData(2, 1) <> "" And arrData(2, 1) <> 0 Then
        arrData(7, 1) = arrData(4, 1) / arrData(2, 1)
    End If
    If arrData(2, 2) <> "" And arrData(2, 2) <> 0 Then
        arrData(7, 2) = arrData(4, 2) / arrData(2, 2)
    End If
    'Rev./Conv.
    If arrData(4, 1) <> "" And arrData(4, 1) <> 0 Then
        arrData(9, 1) = arrData(8, 1) / arrData(4, 1)
    End If
    If arrData(4, 2) <> "" And arrData(4, 2) <> 0 Then
        arrData(9, 2) = arrData(8, 2) / arrData(4, 2)
    End If
    'ROI
    If arrData(5, 1) <> "" And arrData(5, 1) <> 0 Then
        arrData(10, 1) = arrData(8, 1) / arrData(5, 1)
    End If
    If arrData(5, 2) <> "" And arrData(5, 2) <> 0 Then
        arrData(10, 2) = arrData(8, 2) / arrData(5, 2)
    End If
    'CPC
    If arrData(2, 1) <> "" And arrData(2, 1) <> 0 Then
        arrData(12, 1) = arrData(5, 1) / arrData(2, 1)
    End If
    If arrData(2, 2) <> "" And arrData(2, 2) <> 0 Then
        arrData(12, 2) = arrData(5, 2) / arrData(2, 2)
    End If
End Sub


Sub FillArrayData(arrData As Variant, dataColumnIndex As Integer, impValue As Variant, clickValue As Variant, convValue As Variant, costValue As Variant, revValue As Variant)
'将指标数据填充进多维数组，dataColumnIndex表示写入多维数组的第几列
    arrData(1, dataColumnIndex) = AddValue(arrData(1, dataColumnIndex), impValue)
    arrData(2, dataColumnIndex) = AddValue(arrData(2, dataColumnIndex), clickValue)
    arrData(4, dataColumnIndex) = AddValue(arrData(4, dataColumnIndex), convValue)
    arrData(5, dataColumnIndex) = AddValue(arrData(5, dataColumnIndex), costValue)
    arrData(8, dataColumnIndex) = AddValue(arrData(8, dataColumnIndex), revValue)
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

    
    Sheets(sheetName).Cells(startRowIndex + 4, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 4, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 5, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 5, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 7, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 7, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 8, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 8, startColumnIndex + 1).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 11, startColumnIndex).NumberFormatLocal = format
    Sheets(sheetName).Cells(startRowIndex + 11, startColumnIndex + 1).NumberFormatLocal = format

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
