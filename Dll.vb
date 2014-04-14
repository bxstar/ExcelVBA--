Sub AddDdl(rng As String, onact As String, itemsArr() As Variant, name As String, DefaultValue As Long)
''''example:
''''Call AddDdle("D25","Ddl_change",array("firstVal","secondVal"),"MyDDL",2)
''''
ZoomingPage = ActiveWindow.Zoom
ActiveWindow.Zoom = 100
        ActiveSheet.DropDowns.add(Range(rng).left + 20, Range(rng).top - 7, 140, 20).Select
        With Selection
            .OnAction = onact
            .AddItem itemsArr
            .name = name
            .Value = DefaultValue
        End With
ActiveWindow.Zoom = ZoomingPage

End Sub

Sub AddDdlInOneCell(sheetName As String, rng As String, onact As String, itemsArr() As Variant, name As String, DefaultValue As Long, Optional width As Long = 5, Optional height As Long = 5)
''''example:
''''Call AddDdle("D25","Ddl_change",array("firstVal","secondVal"),"MyDDL",2)
''''
Sheets(sheetName).Select
Sheets(sheetName).Activate
ZoomingPage = ActiveWindow.Zoom
If (width = 5) Then
    width = Range(rng).width
End If
If (height = 5) Then
    height = Range(rng).height
End If
ActiveWindow.Zoom = 100
        ActiveSheet.DropDowns.add(Range(rng).left, Range(rng).top, width, height).Select
        With Selection
            .OnAction = onact
            .AddItem itemsArr
            .name = name
            .Value = DefaultValue
        End With
ActiveWindow.Zoom = ZoomingPage

End Sub

Function ValueDDL(NameOfDDL As String, NameOfSheet As String) As String
''''example:
''' Xvalue=ValueDDL("MyDDL","Report")
    
    With Sheets(NameOfSheet).Shapes(NameOfDDL).ControlFormat
    If (UBound(.List) = 1 And .List(1) = "") Then
        ValueDDL = ""
    Else: ValueDDL = .List(.Value)
    End If
    End With

End Function

Function DDLSourceFromDataColumn(column As String, dataSheet As Variant, sort As Boolean, Optional dataRowIndex As Integer = 2) As Variant

    Call AddSheet("tmp")
    Sheets(dataSheet).Select
    Columns(column).Select
    Selection.Copy
    Sheets("tmp").Select
    Range("A1").Select
    ActiveSheet.Paste
    If (Sheets(dataSheet).Range("A2") <> "No Data Available") Then
        Application.CutCopyMode = False
        ActiveSheet.Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
        If sort Then
            ActiveWorkbook.Worksheets("tmp").sort.SortFields.add Key:=Range("A1"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With ActiveWorkbook.Worksheets("tmp").sort
                .SetRange Range("A:A")
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
        Dim lRow As Long
        Dim src As Variant
        lRow = Range("A65536").End(xlUp).Row
        src = Range("A" & dataRowIndex & ":A" & lRow).Value
        If (VarType(src) = vbString) Then
            DDLSourceFromDataColumn = Array(src)
        Else
        DDLSourceFromDataColumn = src
        End If
    Else: DDLSourceFromDataColumn = Array("")
    End If
'    Else: ddlSourceFromDataColumn = Array()
'    End If
'    Worksheets(reportSheet).Activate
'    Shapes(ddlName).ControlFormat.List = src
    'Application.DisplayAlerts = False
    'Sheets("tmp").Delete
    Sheets("tmp").Visible = False
    Sheets(dataSheet).Select
    Range("A1").Select
    'Application.Dis
End Function

Function ArrayDataIndex(arr As Variant, data As Variant) As Integer
'获取数据在数组中的索引号
Dim i As Integer
    For i = 1 To UBound(arr)
        If arr(i) = data Then
            ArrayDataIndex = i
            Exit For
        End If
    Next i
End Function

Function AddValue(data As Variant, add As Variant) As Variant
'在可能为空的数字类型上增加
    If data = "" Then
        AddValue = add
    Else
        AddValue = data + add
    End If

End Function

Function IsInWeek(dateWeekStart As Variant, day As Variant) As Boolean
'判断某天是否在以dateWeekStart为第一天的一周内
    Dim dayStart As Date, dayEnd As Date
    
    dayStart = CDate(dateWeekStart)
    dayEnd = DateAdd("d", 7, dayStart)
    
    If day >= dayStart And day <= dayEnd Then
        IsInWeek = True
    Else
        IsInWeek = False
    End If

End Function


Function IsShapeExists(sheetName As String, shapeName As String) As Boolean
'判断图形是否存在
    On Error GoTo Handler
    Debug.Print Sheets(sheetName).Shapes(shapeName).name
    IsShapeExists = True
    Exit Function
Handler:
        IsShapeExists = False
    

End Function


Sub NumberSort(ByRef a() As Double, Optional sort As String = "up")
'数字排序
    Dim min As Long, max As Long, num As Long, first As Long, last As Long, temp As Long, all As New Collection, steps As Long
    min = LBound(a)
    max = UBound(a)
    all.add a(min)
    steps = 1
    For num = min + 1 To max
        last = all.Count
        If a(num) < CDbl(all(1)) Then all.add a(num), BEFORE:=1: GoTo nextnum '加到第一项
        If a(num) > CDbl(all(last)) Then all.add a(num), AFTER:=last: GoTo nextnum '加到最后一项


        first = 1
        Do While last > first + 1 '利用DO循环减少循环次数
        temp = (last + first) / 2
        If a(num) > CDbl(all(temp)) Then
        first = temp
        Else
        last = temp
        steps = steps + 1
        End If
        Loop
        all.add a(num), BEFORE:=last '加到指定的索引

nextnum:
        steps = steps + 1
    Next
    For num = min To max
        If sort = "UP" Or sort = "up" Then a(num) = CDbl(all(num - min + 1)): steps = steps + 1 '升序
        If sort = "DOWN" Or sort = "down" Then a(num) = CDbl(all(max - num + 1)): steps = steps + 1 '降序
    Next
    Set all = Nothing
End Sub


Sub StringSort(ByRef a As Variant, Optional sort As String = "UP")
'字符串排序
    Dim min As Long, max As Long, num As Long, first As Long, last As Long, temp As Long, all As New Collection, steps As Long
    min = LBound(a)
    max = UBound(a)
    all.add a(min)
    steps = 1
    For num = min + 1 To max
        first = 1
        last = all.Count
        If a(num) < all(1) Then all.add a(num), BEFORE:=1: GoTo nextnum '加到第一项
        If a(num) > all(last) Then all.add a(num), AFTER:=last: GoTo nextnum '加到最后一项


        Do While last > first + 1 '利用DO循环减少循环次数
        temp = (last + first) / 2
        If a(num) > all(temp) Then
        first = temp
        Else
        last = temp
        steps = steps + 1
        End If
        Loop
        all.add a(num), BEFORE:=last '加到指定的索引

nextnum:
        steps = steps + 1
    Next
    For num = min To max
        If sort = "UP" Or sort = "up" Then a(num) = all(num - min + 1): steps = steps + 1 '升序
        If sort = "DOWN" Or sort = "down" Then a(num) = all(max - num + 1): steps = steps + 1 '降序
    Next
    Set all = Nothing
End Sub

Sub YearMonthSort(ByRef a As Variant, Optional sort As String = "UP")
'字符串排序
    Dim min As Long, max As Long, num As Long, first As Long, last As Long, temp As Long, all As New Collection, steps As Long
    min = LBound(a)
    max = UBound(a)
    all.add a(min)
    steps = 1
    For num = min + 1 To max
        first = 1
        last = all.Count
        If YearMonthCompare(all(1), a(num)) Then all.add a(num), BEFORE:=1: GoTo nextnum  '加到第一项
        If YearMonthCompare(a(num), all(last)) Then all.add a(num), AFTER:=last: GoTo nextnum '加到最后一项


        Do While last > first + 1 '利用DO循环减少循环次数
        temp = (last + first) / 2
        If YearMonthCompare(a(num), all(temp)) Then
        first = temp
        Else
        last = temp
        steps = steps + 1
        End If
        Loop
        all.add a(num), BEFORE:=last '加到指定的索引

nextnum:
        steps = steps + 1
    Next
    For num = min To max
        If sort = "UP" Or sort = "up" Then a(num) = all(num - min + 1): steps = steps + 1 '升序
        If sort = "DOWN" Or sort = "down" Then a(num) = all(max - num + 1): steps = steps + 1 '降序
    Next
    Set all = Nothing
End Sub


Function YearMonthCompare(yearMonth1 As Variant, yearMonth2 As Variant)
'年月的比较，比如：2014/2和2014/10，2014/2是要小于2014/10，返回值：yearMonth1>yearMonth2返回true
    Dim ymFull1 As String, ymFull2 As String
    
    If Len(yearMonth1) = 6 Then
        ymFull1 = left(yearMonth1, 4) & Replace(yearMonth1, Right(yearMonth1, 1), "0" & Right(yearMonth1, 1), 5)
    Else
        ymFull1 = yearMonth1
    End If
    
    If Len(yearMonth2) = 6 Then
        ymFull2 = left(yearMonth2, 4) & Replace(yearMonth2, Right(yearMonth2, 1), "0" & Right(yearMonth2, 1), 5)
    Else
        ymFull2 = yearMonth2
    End If
    
    YearMonthCompare = ymFull1 > ymFull2

End Function


Function CollectionToArray(coll As Collection)
'集合转换成数组
    Dim Ndx As Long
    
    ReDim arr(1 To coll.Count)
    
    For Ndx = 1 To coll.Count
        If IsObject(coll(Ndx)) = True Then
            Set arr(Ndx) = coll(Ndx)
        Else
            arr(Ndx) = coll(Ndx)
        End If
    Next Ndx
    
    CollectionToArray = arr

End Function
