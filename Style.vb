Sub BasicDesign(sheetName As String, EndColumn As String, Optional EndRow As Integer)
'
' Gray in the sides and undisplay gridlines
''the function get paramater :column :from where to start the gray color and optionaly paramater row:from whigh row to start the gray color

    Sheets(sheetName).Activate
    Range("A:A").Select
    Selection.Interior.color = RGB(127, 127, 127)
    Selection.ColumnWidth = 2
    Columns(EndColumn & ":" & EndColumn).Select
    Range(Selection, Selection.End(xlToRight)).Select
'    range(EndColumn & ":XFD").Select
    Selection.Interior.color = RGB(127, 127, 127)
    Range("1:1").Select
    Selection.RowHeight = 12.75
    Selection.Interior.color = RGB(127, 127, 127)
    If EndRow <> 0 Then
       Rows(EndRow & ":" & EndRow).Select
       Range(Selection, Selection.End(xlDown)).Select
       Selection.Interior.color = RGB(127, 127, 127)
    End If
    Range("B:B").Select
    Selection.ColumnWidth = 2
    ActiveWindow.DisplayGridlines = False
    Sheets(sheetName).Range("A1").Select
End Sub
Sub AddBorderWithColor(sheetName As String, Range As String, R As Integer, G As Integer, B As Integer)
'
'the function get range that you want to add border and RGB paramatrs of the color
'
     Sheets(sheetName).Range(Range).Select
     Selection.Borders(xlEdgeLeft).color = RGB(R, G, B)
     Selection.Borders(xlEdgeTop).color = RGB(R, G, B)
     Selection.Borders(xlEdgeBottom).color = RGB(R, G, B)
     Selection.Borders(xlEdgeRight).color = RGB(R, G, B)
     Selection.Borders(xlInsideVertical).color = RGB(R, G, B)
     Selection.Borders(xlInsideHorizontal).color = RGB(R, G, B)
End Sub
Sub AddColor(sheetName As String, Range As String, R As Integer, G As Integer, B As Integer)
    ' the function get range to paint in RGB colors
    Sheets(sheetName).Range(Range).Select
    Selection.Interior.color = RGB(R, G, B)
End Sub
Sub AddProperties(sheetName As String)
    'the function give the properties of the report
    Dim CountRowMap As Integer
    CountRowMap = 20
    Sheets(sheetName).Range("D5:D9").Select
    Selection.Value = "=IFERROR(VLOOKUP($C5,'Headers'!$A$1:$B$" & CountRowMap & ",2,0),$C5 )"
End Sub
Sub AddPropertiesDesign(sheetName As String, RangeToBold As String)
   Sheets(sheetName).Activate
   Call AddBorder(sheetName, "C8:D11")
   Call AddColor(sheetName, "C8:D8", 242, 242, 242)
   Call AddColor(sheetName, "C10:D10", 242, 242, 242)
   Sheets(sheetName).Range("C8") = "Report Type"
   Sheets(sheetName).Range("C9") = "Profile Name"
   Sheets(sheetName).Range("C10") = "Creation Date"
   Sheets(sheetName).Range("C11") = "Time Range"
   Range("C:C").ColumnWidth = 20
   Range("D:D").ColumnWidth = 28
   Call MYCenterFont(sheetName, "D8:D11")
   Call AddProperties(sheetName)
   Call MYBoldFont(sheetName, RangeToBold)
End Sub
Sub MYCenterFont(sheetName As String, Getrange As String)
    Sheets(sheetName).Activate
    Range(Getrange).HorizontalAlignment = xlCenter
End Sub
Sub AddTotalLine(sheetName As String, StartRow As Integer, EndRow As Integer) '17116
    Sheets(sheetName).Range("c41") = "Total"
    Sheets(sheetName).Range("i41") = "=SUM(i" & StartRow & ":i" & EndRow & ")"
    Sheets(sheetName).Range("j41") = "=SUM(j" & StartRow & ":j" & EndRow & ")"
    Sheets(sheetName).Range("k41") = "=IF(i41=0,0,j41/i41)"
    Sheets(sheetName).Range("l41") = "=SUM(l42:l" & EndRow & ")"
    Sheets(sheetName).Range("m41") = "=SUM(m42:m" & EndRow & ")"
    Sheets(sheetName).Range("n41") = "=IF(j41=0,0,m41/j41)"
    Sheets(sheetName).Range("o41") = "=IF(l41=0,0,m41/l41)"
    Sheets(sheetName).Range("p41") = "=IF(j41=0,0,l41/j41)"
    Sheets(sheetName).Range("q41") = "=IF((SUM(i42:i" & EndRow & "))=0,0,(SUMPRODUCT(q42:q" & EndRow & ",i42:i" & EndRow & "))/(SUM(i42:i" & EndRow & ")))"
    Sheets(sheetName).Range("r41") = "=SUM(r42:r" & EndRow & ")"
    Sheets(sheetName).Range("s41") = "=IF(m41=0,0,r41/m41)"
    Sheets(sheetName).Range("t41") = "=IF(l41=0,0,r41/l41)"
End Sub
Sub MYBoldFont(sheetName As String, rng As String)
    Sheets(sheetName).Range(rng).Select
    Selection.font.Bold = True
End Sub
Sub SizeFont(sheetName As String, Size As String, rng As String)
    Range(rng).Select
    With Selection.font
        .name = "Calibri"
        .FontStyle = "Regular"
        .Size = 10
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub

Sub SizeFont2(sheetName As String, Size As Integer, rng As String)
    Range(rng).Select
    With Selection.font
        .name = "Calibri"
        .FontStyle = "Regular"
        .Size = Size
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
End Sub
Sub AddPicture(sheetName As String)
    Dim DrObj
    Dim Pict
    Sheets("Sheet1").Activate
    Set DrObj = ActiveSheet.DrawingObjects
    For Each Pict In DrObj
       ' If Left(Pict.name, 7) = "Picture" Then
        Pict.Select
        Pict.Copy
       ' Pict.Delete
       ' End If
    Next
    Sheets(sheetName).Activate
    Sheets(sheetName).Range("C3").Select
    Sheets(sheetName).Paste
    
End Sub
Sub CopyPaste(sourceSheetName As String, sourceRange As String, destSheetName As String, destRange As String)
    Sheets(sourceSheetName).Select
    Range(sourceRange).Copy
    Sheets(destSheetName).Select
    Range(destRange).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub CopyPasteWithDesign(sourceSheetName As String, sourceRange As String, destSheetName As String, destRange As String)
    Sheets(sourceSheetName).Select
    Range(sourceRange).Copy
    Sheets(destSheetName).Select
    Range(destRange).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
End Sub
Sub CopyPaseFilteredData(sourceSheetName As String, sourceRange As String, destSheetName As String, destRange As String)
    Sheets(sourceSheetName).Select
    Range(sourceRange).SpecialCells(xlCellTypeVisible).Copy
    Sheets(destSheetName).Select
    Range(destRange).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Sub MergeCells(sheetName As String, rng As String)
    Sheets(sheetName).Activate
    Range(rng).Merge
End Sub

Sub ColorToFont(sheetNm As String, rng As String, R As Integer, G As Integer, B As Integer)
    Sheets(sheetNm).Activate
    Range(rng).Select
    Selection.font.color = RGB(R, G, B)
End Sub
Sub AddPropertiesGenericDesign(sheetName As String, RangeToBold As String)
   Sheets(sheetName).Range("C5") = "Report Name"
   Sheets(sheetName).Range("C6") = "Report Type"
   Sheets(sheetName).Range("C7") = "Time Range"
   Sheets(sheetName).Range("C8") = "Profile Name"
   Sheets(sheetName).Range("C9") = "Creation Date"
   
   Sheets(sheetName).Range("D5:D9").Select
     With Selection.font
        .name = "Calibri"
        .FontStyle = "Bold"
        .Size = 12
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
   
    Range("C:C").ColumnWidth = 20
    Range("D:D").ColumnWidth = 28
    Sheets(sheetName).Range("A3").Select
    Selection.RowHeight = 37.5
    Sheets(sheetName).Range("A5:A9").Select
    Selection.RowHeight = 15
    Sheets(sheetName).Range("C5:C9").Select
    With Selection.font
        .name = "Calibri"
        .FontStyle = "Regular"
        .Size = 12
        .color = 8421504
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
   Call MYRightFont(sheetName, "C5:C9")
   Call AddProperties(sheetName)
  
   Call MYBoldFont(sheetName, RangeToBold)
End Sub
Sub MYRightFont(sheetName As String, Getrange As String)
    Sheets(sheetName).Activate
    Range(Getrange).HorizontalAlignment = xlRight
End Sub
Sub AddTextBox(sheetName As String, text As String, left As Single, top As Single, width As Single, height As Single, fontSize As Variant, color As Variant, Optional font As String = "Arial")
Dim s As shape
Sheets(sheetName).Activate
Set s = ActiveSheet.Shapes.AddTextBox(msoTextOrientationHorizontal, left, top, _
        width, height)
        s.name = "txtHeader"
   s.TextFrame.Characters.text = text
    With s.TextFrame.Characters(Start:=1, Length:=216).font
        .name = font
        .FontStyle = "Regular"
        .Size = fontSize
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .color = color
        .TintAndShade = 0
    End With
    s.TextFrame.AutoSize = msoAutoSizeShapeToFitText
    s.Line.Visible = msoFalse
End Sub
Sub AddHeader(sheetName As String, text As String)
    Call AddTextBox(sheetName, text, 210, 30, 400, 50, 28, -14766175)
End Sub
Sub UpdateColorToHeader(sheetName As String)
    Sheets(sheetName).Select
    ActiveSheet.Shapes("txtHeader").Select
    Selection.font.color = RGB(161, 175, 30)
    Range("a1").Select
End Sub
Sub DesignTable(sheetName As String, firstRow As Long, LastRow As Long, firstCol As String, LastCol As String)
'
' DesignTable Macro
'

'
Sheets(sheetName).Select
Dim i As Long
For i = firstRow + 1 To LastRow Step 2
    Range(firstCol & i & ":" & LastCol & i).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
Next

    Range(firstCol & firstRow - 1 & ":" & LastCol & LastRow).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 15
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 15
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 15
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 15
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range(firstCol & firstRow & ":" & LastCol & LastRow).Select

    Selection.font.Size = 10
    Range("A1").Select
End Sub

Sub FormatCells(dataSheetName, sheetName As String, rng As String, metric As String) 'Format cells
    Sheets(sheetName).Select
    Select Case metric
        Case "Imp.": Range(rng).NumberFormat = "#,##0"
        Case "Impressions": Range(rng).NumberFormat = "#,##0"
        Case "Clicks": Range(rng).NumberFormat = "#,##0"
        Case "Conv.": Range(rng).NumberFormat = "#,##0.00"
        Case "Conversions": Range(rng).NumberFormat = "#,##0.00"
        Case "ROI": Range(rng).NumberFormat = "#,##0.00"
        Case "Avg Pos.": Range(rng).NumberFormat = "#,##0.00"
        Case "CTR": Range(rng).NumberFormat = "#,##0.00%"
        Case "Conv. Rate": Range(rng).NumberFormat = "#,##0.00%"
        Case "Cost": Range(rng).NumberFormat = Sheets(dataSheetName).Range("G2").NumberFormat
        Case "Cost/Conv.": Range(rng).NumberFormat = Sheets(dataSheetName).Range("G2").NumberFormat
        Case "Rev.": Range(rng).NumberFormat = Sheets(dataSheetName).Range("G2").NumberFormat
        Case "Rev./Conv.": Range(rng).NumberFormat = Sheets(dataSheetName).Range("G2").NumberFormat
        Case "Avg. CPC": Range(rng).NumberFormat = Sheets(dataSheetName).Range("G2").NumberFormat
        Case Else: Range(rng).NumberFormat = ""
    End Select
End Sub


