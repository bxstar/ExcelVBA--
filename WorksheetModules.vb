Sub AddSheet(sheetName As String, Optional AfterS As String = "", Optional ColorTab As Long = xlAutomatic)
    
    Dim sh As Worksheet, flg As Boolean
    flg = False
    For Each sh In Worksheets
        If sh.name Like sheetName Then flg = True: Exit For
    Next
    If flg = False Then
        Worksheets.add().name = sheetName
    Else
        Worksheets(sheetName).Visible = True
        Worksheets(sheetName).Select
        Cells.Select
        Selection.Delete Shift:=xlUp
    End If
   ' Worksheets(SheetName).Columns("a:a").ColumnWidth = 20
   ' Worksheets(SheetName).Columns("b:b").ColumnWidth = 20
    If ColorTab <> xlAutomatic Then Call ColorTabSheet(sheetName, ColorTab)
     ActiveSheet.DisplayRightToLeft = False
End Sub

Sub DuplicateSheet(srcSheetName As String, sheetName As String, Optional AfterS As String = "", Optional ColorTab As Long = xlAutomatic)
        
        flg = False
        For Each sh In Worksheets
            If sh.name Like sheetName Then flg = True: Exit For
        Next
        
        If flg = False Then
            If AfterS = "" Then AfterS = Worksheets(Worksheets.Count).name
            Worksheets(srcSheetName).Copy AFTER:=Worksheets(AfterS)
            ActiveSheet.name = sheetName
        End If
    If ColorTab <> xlAutomatic Then Call ColorTabSheet(sheetName, ColorTab)
            
            
End Sub

Sub ColorTabSheet(sheetName As String, MyColor As Long)
    Worksheets(sheetName).Tab.color = MyColor
End Sub

Sub AddColorToLife(dataSheet As String, rang As String, ColorGet As Long, R As Long, G As Long, b As Long)
If ColorGet <> 0 Then
Sheets(dataSheet).Range(rang).Interior.color = ColorGet
Else


Sheets(dataSheet).Range(rang).Interior.color = RGB(R, G, b)
 End If

    
    
    Sheets(dataSheet).Activate
     Range(rang).Select
       
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = -0.249977111117893
        .Weight = xlThin
    End With
   
    

End Sub
Sub AddTable(dataSheet As String, StartRowT As Long, LastRowT As Long, StartColumnT As String, EndColumnT As String, StartPivotT As Long)
    Dim rang As String
    Dim rang2 As String
    Dim rang3 As String
    Dim flag As Boolean
    For i = StartRowT To StartRowT + LastRowT - StartPivotT - 1 Step 2
    Dim Row As Long
    flag = False
    rang = StartColumnT & i & ":" & EndColumnT & i
    Call AddColorToLife(dataSheet, rang, 0, 242, 242, 242)
    
    k = i + 1
    If k < LastRowT + StartRowT - StartPivotT Then
        rang2 = StartColumnT & k & ":" & EndColumnT & k
        Call AddBorder(dataSheet, rang2)
         flag = True
    End If
    Next
    If flag = True Then
    k = k + 1
    
    End If
    
    rang3 = StartColumnT & k & ":" & EndColumnT & k
    Call AddBorder(dataSheet, rang3)
End Sub
Sub AddBorder(dataSheet As String, rang As String)
    Sheets(dataSheet).Activate
         Range(rang).Select
           
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.249977111117893
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.249977111117893
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.249977111117893
            .Weight = xlThin
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.249977111117893
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.249977111117893
            .Weight = xlThin
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ThemeColor = 1
            .TintAndShade = -0.249977111117893
            .Weight = xlThin
        End With
       
End Sub
Sub FormatColumnDollar(dataSheet As String)
   Dim LastRow As Integer
   LastRow = 1000
  Sheets(dataSheet).Range("i16:i" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"
   Sheets(dataSheet).Range("j16:j" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"
   Sheets(dataSheet).Range("l16:l" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"
   Sheets(dataSheet).Range("m16:m" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"
   Sheets(dataSheet).Range("p16:p" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"

End Sub
Sub FormatColumnPercent(dataSheet As String)
 LastRow = 1000
   Sheets(dataSheet).Range("g16:g" & LastRow).NumberFormat = "0.00%"
   Sheets(dataSheet).Range("k16:k" & LastRow).NumberFormat = "0.00%"
End Sub
Sub FormatColumnDollarAll(dataSheet As String)
   Dim LastRow As Integer
   LastRow = 1000
   Sheets(dataSheet).Range("k16:k" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"
   Sheets(dataSheet).Range("j16:j" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"
   Sheets(dataSheet).Range("m16:m" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"
   Sheets(dataSheet).Range("n16:n" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"
   Sheets(dataSheet).Range("q16:q" & LastRow).NumberFormat = "[$$-409]#,##0.00;[Red][$$-409]#,##0.00"

End Sub
Sub FormatColumnPercentAll(dataSheet As String)
 LastRow = 1000
   Sheets(dataSheet).Range("h16:h" & LastRow).NumberFormat = "0.00%"
   Sheets(dataSheet).Range("l16:l" & LastRow).NumberFormat = "0.00%"
End Sub
Sub ChangeDirection(dataSheet As String)
Dim LastRow As Integer
LastRow = 1000
 With Sheets(dataSheet).Range("A1:v" & LastRow)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub ChangeAllDirec()
Call ChangeDirection("Campaigns(All)")
Call ChangeDirection("Campaigns(Google)")
Call ChangeDirection("Campaigns(MSN)")
Call ChangeDirection("Ads")
Call ChangeDirection("Keywords")
End Sub
Sub CheckAll(PivotName As String, field As String)
    With Sheets(PivotName).PivotTables(PivotName).PivotFields(field)
            For i = 1 To .PivotItems.Count
                .PivotItems(i).Visible = True
            Next i
   End With
End Sub
'add uniqe column
Sub AddColumnConcat(dataSheet As String)
        Dim LastRow As Long
        Dim LastCol As Long
        Dim LastCell As String
        LastRow = Sheets(dataSheet).Range("A65536").End(xlUp).Row
        LastCol = Sheets(dataSheet).Range("A1").End(xlToRight).column
        LastCell = Cells(LastRow, LastCol).Address
        Sheets(dataSheet).Activate
        Sheets(dataSheet).Cells(1, LastCol + 1).Value = "uniqe"
        
           Sheets(dataSheet).Cells(2, LastCol + 1).Formula = "=CONCATENATE(D" & 2 & ",E" & 2 & ")"
            Sheets(dataSheet).Cells(3, LastCol + 1).Formula = "=CONCATENATE(D" & 3 & ",E" & 3 & ")"
         Sheets(dataSheet).Range(Cells(2, LastCol + 1), Cells(3, LastCol + 1)).AutoFill Destination:=Range(Cells(2, LastCol + 1), Cells(LastRow, LastCol + 1))
End Sub
Sub FillZeros(dataSheet As String, StartColumn As Integer, EndColumn As Integer, StartRow As Integer, EndRow As Integer)
 Dim i As Integer
 Dim j As Integer
 For i = StartColumn To EndColumn
     For j = StartRow To EndRow
        If Sheets(dataSheet).Cells(j, i).Value = "" Then
           Sheets(dataSheet).Cells(j, i).Value = 0
        End If
        
        
        
        
        Next
     Next
        
End Sub
Sub FillZerosAll(dataSheet As String, StartColumn As Integer, EndColumn As Integer, StartRow As Integer, EndRow As Integer)
 Dim i As Integer
 Dim j As Integer
 For i = StartColumn To EndColumn
     For j = StartRow To EndRow
      Sheets(dataSheet).Cells(j, i).Value = 0
     
        
        
        
        
        Next
     Next
        
End Sub
Sub AddIcons()
Sheets("Overall").Activate
Range("F51:F54,F57:F60").Select
    
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 0.01
        .Operator = 7
    End With
    Range("F55:F56,F61:F62").Select
   
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = True
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 0.01
        .Operator = 7
    End With
    
    
    Range("K51:K54,K57:K60").Select
   
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 0.01
        .Operator = 7
    End With
     Range("K55:K56,K61:K62").Select
    
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = True
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 0.01
        .Operator = 7
    End With
    
    
    Range("P51:P54,P57:P60").Select
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 0.01
        .Operator = 7
    End With
    Range("P55:P56,P61:P62").Select
   
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = True
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 0.01
        .Operator = 7
    End With

End Sub
Sub AddColumnConcatS(dataSheet As String)
        Dim LastRow As Long
        Dim LastCol As Long
        Dim LastCell As String
        LastRow = Sheets(dataSheet).Range("A65536").End(xlUp).Row
        LastCol = Sheets(dataSheet).Range("A1").End(xlToRight).column
        LastCell = Cells(LastRow, LastCol).Address
        Sheets(dataSheet).Activate
        Sheets(dataSheet).Cells(1, LastCol + 1).Value = "uniqe"
        
           Sheets(dataSheet).Cells(2, LastCol + 1).Formula = "=CONCATENATE(A" & 2 & ",C" & 2 & ")"
            Sheets(dataSheet).Cells(3, LastCol + 1).Formula = "=CONCATENATE(A" & 3 & ",C" & 3 & ")"
         Sheets(dataSheet).Range(Cells(2, LastCol + 1), Cells(3, LastCol + 1)).AutoFill Destination:=Range(Cells(2, LastCol + 1), Cells(LastRow, LastCol + 1))
End Sub
Sub AddIconsSocial()


Sheets("Social Performance").Activate
Range("F21:F24,F27:F30").Select
    
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = False
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 0.01
        .Operator = 7
    End With
    Range("F25:F26,F31:F32").Select
   
    Selection.FormatConditions.AddIconSetCondition
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .ReverseOrder = True
        .ShowIconOnly = False
        .IconSet = ActiveWorkbook.IconSets(xl3TrafficLights1)
    End With
    With Selection.FormatConditions(1).IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = 7
    End With
    With Selection.FormatConditions(1).IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 0.01
        .Operator = 7
    End With

End Sub

Sub PaintSheets()
Sheets("Keywords").Range("c29:s29, c17:s17, c19:s19, c21:s21, c23:s23, c25:s25, c27:s27").Interior.color = RGB(242, 242, 242)
Sheets("Keywords").Range("c35:s35, c37:s37, c39:s39, c41:s41, c43:s43, c45:s45, c47:s47").Interior.color = RGB(242, 242, 242)
Sheets("Ads").Range("c17:v17, c19:v19, c21:v21, c23:v23, c25:v25, c27:v27, c29:v29").Interior.color = RGB(242, 242, 242)
Sheets("Keywords").Range("c15:s15,c33:s33").Interior.color = RGB(90, 90, 90)
Sheets("Ads").Range("c15:v15,c34:v34").Interior.color = RGB(90, 90, 90)
Sheets("Ads").Range("c36:v36, c38:v38, c40:v40, c42:v42, c44:v44, c46:v46, c48:v48").Interior.color = RGB(242, 242, 242)

Sheets("Overall").Range("c68:o68, c70:o70, c72:o72, c74:o74, c76:o76, c78:o78, c86:o86, c88:o88, c90:o90, c92:o92, c94:o94").Interior.color = RGB(242, 242, 242)
Sheets("Overall").Range("c96:o96, c104:o104, c106:o106, c108:o108, c110:o110, c112:o112, c114:o114").Interior.color = RGB(242, 242, 242)
Sheets("Overall").Range("c52:f52, c54:f54, c56:f56, c58:f58, c60:f60, c62:f62").Interior.color = RGB(242, 242, 242)
Sheets("Overall").Range("h52:k52, h54:k54, h56:k56, h58:k58, h60:k60, h62:k62").Interior.color = RGB(242, 242, 242)
Sheets("Overall").Range("m52:p52, m54:p54, m56:p56, m58:p58, m60:p60, m62:p62,c49:f49,h49:k49,m49:p49").Interior.color = RGB(242, 242, 242)
Sheets("Overall").Range("c49:f49,h49:k49,m49:p49,c66:o66,c84:o84,c102:o102").Interior.color = RGB(90, 90, 90)
Sheets("Campaigns(All)").Range("d15:q15").Interior.color = RGB(90, 90, 90)
Sheets("Campaigns(Google)").Range("d15:p15").Interior.color = RGB(90, 90, 90)
Sheets("Campaigns(MSN)").Range("d15:p15").Interior.color = RGB(90, 90, 90)
End Sub
Function checkFB()
    Dim LastCol As Long
    LastCol = Sheets("DataSocialPerfo").Range("A1").End(xlToRight).column
    For i = 1 To LastCol
       If Sheets("DataSocialPerfo").Cells(1, i) = "Social_Imp" Then
          checkFB = True
        End If
    Next
End Function
Sub FormatColumnComma()
   Dim LastRow As Integer
   LastRow = 1000
   Sheets("Campaigns(All)").Range("f16:f" & LastRow).NumberFormat = "#,##0"
   Sheets("Campaigns(All)").Range("i16:i" & LastRow).NumberFormat = "#,##0"
   Sheets("Campaigns(Google)").Range("e16:e" & LastRow).NumberFormat = "#,##0"
    Sheets("Campaigns(Google)").Range("h16:h" & LastRow).NumberFormat = "#,##0"
   Sheets("Campaigns(Google)").Range("f16:f" & LastRow).NumberFormat = "#,##0"
   Sheets("Campaigns(All)").Range("g16:g" & LastRow).NumberFormat = "#,##0"
   Sheets("Campaigns(MSN)").Range("e16:e" & LastRow).NumberFormat = "#,##0"
   Sheets("Campaigns(MSN)").Range("f16:f" & LastRow).NumberFormat = "#,##0"
   Sheets("Campaigns(MSN)").Range("h16:h" & LastRow).NumberFormat = "#,##0"
End Sub
Sub FormatToolTip(valDDL As String, Range As String)
  
     Select Case valDDL
        'comma
        Case "Impressions", "Clicks"
            Sheets("TempSheet").Range(Range).NumberFormat = "#,##0"
        'percent
        Case "CTR.", "Conv.Rate."
            Sheets("TempSheet").Range(Range).NumberFormat = "0.00%"
        Case "Cost/Conv", "Rev/Conv", "AvgCpc", "Cost"
          Sheets("TempSheet").Range(Range).NumberFormat = "[$$-C09]#,##0.00"
        Case "Conversions", "Rev.", "ROI.", "AvgPos"
            Sheets("TempSheet").Range(Range).NumberFormat = "#,##0.00"
        
             
    End Select
   
             
   
   
End Sub
Sub FormatSocial()
  Dim LastRow As Integer
   LastRow = 1000
   Sheets("Social Performance").Range("d38:d" & LastRow).NumberFormat = "#,##0"
   Sheets("Social Performance").Range("e38:e" & LastRow).NumberFormat = "#,##0"
   
    Sheets("Social Performance").Range("f38:f" & LastRow).NumberFormat = "0.00%"
   Sheets("Social Performance").Range("g38:g" & LastRow).NumberFormat = "#,##0"
   Sheets("Social Performance").Range("h38:h" & LastRow).NumberFormat = "[$$-C09]#,##0.00"
   Sheets("Social Performance").Range("i38:i" & LastRow).NumberFormat = "[$$-C09]#,##0.00"
   Sheets("Social Performance").Range("j38:j" & LastRow).NumberFormat = "0.00%"
   Sheets("Social Performance").Range("k38:k" & LastRow).NumberFormat = "#,##0.00"
   Sheets("Social Performance").Range("l38:l" & LastRow).NumberFormat = "[$$-C09]#,##0.00"
   Sheets("Social Performance").Range("m38:m" & LastRow).NumberFormat = "#,##0.00"
   Sheets("Social Performance").Range("n38:n" & LastRow).NumberFormat = "#,##0.00"
   Sheets("Social Performance").Range("o38:o" & LastRow).NumberFormat = "[$$-C09]#,##0.00"
End Sub
