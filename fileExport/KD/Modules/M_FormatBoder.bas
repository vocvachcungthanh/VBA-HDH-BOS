Attribute VB_Name = "M_FormatBoder"
'-- Khi dung chi can goi sub Format. VD: Format_ hoac là Call Format_

Sub Format_(dongCuoi, dongCuoiHang, rangeSelect, columSelect, Kieu, nameTable)
'    If StyleExists("Net dam", ActiveWorkbook) = False Then Create_Style
    ActiveSheet.ListObjects(nameTable).DataBodyRange.Borders.LineStyle = xlNone
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim icol As Long
    Dim col As String
    Dim DongCuoiFR As Long

    DongCuoiFR = dongCuoi
    icol = dongCuoiHang
    col = LayTenCotTuSoCot(icol)
    ws.Cells.Select
    Selection.RowHeight = 15
    Clear_border
    Clear_BoderTable nameTable
    If DongCuoiFR > 0 Then
         bovien ws.Range("" & rangeSelect & ":" & col & "" & DongCuoiFR), xlMedium
    End If


    Cells(11, 11).RowHeight = 40
    Cells(10, 10).RowHeight = 5
    Cells(4, 4).RowHeight = 5
    
    Range("" & rangeSelect & ":" & col & "11 ").Select
    Selection.Style = "Net dam"

    Columns("" & columSelect & "").Select
    Selection.ColumnWidth = 2
    Columns("" & col & ":" & col & "").Select
    Selection.ColumnWidth = 15

    icol = icol + 1
    col = LayTenCotTuSoCot(icol)
    Columns("" & col & ":" & col & "").Select
    Selection.ColumnWidth = 2

   If Kieu = 0 Then
        bovien ws.Range("A1:" & col & "" & DongCuoiFR + 2), xlThick
    ElseIf Kieu = 2 Then
        bovien ws.Range("C1:" & col & "" & DongCuoiFR + 2), xlThick
   Else
         bovien ws.Range("B1:" & col & "" & DongCuoiFR + 2), xlThick
   End If


    icol = icol + 1
    col = LayTenCotTuSoCot(icol)
    Columns("" & col & ":" & col & "").Select
    Selection.ColumnWidth = 0.1
    icol = icol + 1
    col = LayTenCotTuSoCot(icol)
    Columns("" & col & ":" & col & "").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ColumnWidth = 0

    Rows(DongCuoiFR + 3 & ":" & DongCuoiFR + 3).Select
    Selection.RowHeight = 0.1
    Rows(DongCuoiFR + 4 & ":" & DongCuoiFR + 4).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RowHeight = 0

    Set ws = Nothing
    Range("A12").Select
End Sub

Sub DinhDang(rg As Range, header As String, table As String)
    rg.Style = table
    Dim i As Long, wg As Variant
    For i = 1 To rg.Columns.Count
    rg(1, i).Style = header
    rg(1, i).HorizontalAlignment = xlCenter
    Next i
End Sub

Sub bovien(rg As Range, tp As Variant)
    rg.Borders(xlDiagonalDown).LineStyle = xlNone
    rg.Borders(xlDiagonalUp).LineStyle = xlNone
    With rg.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = tp
    End With
    With rg.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = tp
    End With
    With rg.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = tp
    End With
    With rg.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = -0.249946592608417
        .Weight = tp
    End With
End Sub

Function Clear_BoderTable(nameTable)
    
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects("" & nameTable & "")
    
    With tbl.Range.Borders
        .LineStyle = xlNone
    End With

End Function
Sub Clear_border()
'    Cells.Select
'    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
'    Selection.Borders(xlEdgeTop).LineStyle = xlNone
'    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
'    Selection.Borders(xlEdgeRight).LineStyle = xlNone
'    Selection.Borders(xlInsideVertical).LineStyle = xlNone
'    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Function LayTenCotTuSoCot(icol As Long)
    Dim num As Long, buf As String
    buf = Cells(10, icol).Address(False, False)
    buf = Left(buf, Len(buf) - 1)
    If Len(buf) > 1 Then
        buf = Left(buf, Len(buf) - 1)
    End If
    LayTenCotTuSoCot = buf
End Function

Public Function StyleExists(ByVal styleName As String, ByVal Target As Workbook) As Boolean
' Returns TRUE if the named style exists in the target workbook.
    On Error Resume Next
    StyleExists = Len(Target.Styles(styleName).Name) > 0
    On Error GoTo 0
End Function

Sub Create_Style()
    ActiveWorkbook.Styles.Add Name:="Net dam"
    With ActiveWorkbook.Styles("Net dam")
        .IncludeNumber = True
        .IncludeFont = True
        .IncludeAlignment = True
        .IncludeBorder = True
        .IncludePatterns = True
        .IncludeProtection = True
    End With
    ActiveWorkbook.Styles("Net dam").NumberFormat = "@"
    With ActiveWorkbook.Styles("Net dam").Font
        .Name = "Arial Narrow"
        .Size = 12
        .Bold = True
        .Italic = False
        .Underline = xlUnderlineStyleNone
        .Strikethrough = False
        .ThemeColor = 10
        .TintAndShade = 0.799951170384838
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveWorkbook.Styles("Net dam")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
    End With
    With ActiveWorkbook.Styles("Net dam").Borders(xlLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = 0.799981688894314
        .Weight = xlThin
    End With
    With ActiveWorkbook.Styles("Net dam").Borders(xlRight)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = 0.799981688894314
        .Weight = xlThin
    End With
    ActiveWorkbook.Styles("Net dam").Borders(xlTop).LineStyle = xlNone
    With ActiveWorkbook.Styles("Net dam").Borders(xlBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 10
        .TintAndShade = 0.799981688894314
        .Weight = xlThin
    End With
    ActiveWorkbook.Styles("Net dam").Borders(xlDiagonalDown).LineStyle = xlNone
    ActiveWorkbook.Styles("Net dam").Borders(xlDiagonalUp).LineStyle = xlNone
    With ActiveWorkbook.Styles("Net dam").Interior
        .Pattern = xlSolid
        .PatternColorIndex = 0
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With ActiveWorkbook.Styles("Net dam")
        .Locked = True
        .FormulaHidden = False
    End With
End Sub




