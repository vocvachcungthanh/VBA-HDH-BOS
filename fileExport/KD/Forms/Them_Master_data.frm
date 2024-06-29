VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Them_Master_data 
   Caption         =   "UserForm2"
   ClientHeight    =   6300
   ClientLeft      =   -60
   ClientTop       =   -150
   ClientWidth     =   8505.001
   OleObjectBlob   =   "Them_Master_data.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Them_Master_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Mst_data_col As Long

Private Sub Label2_Click()

End Sub

Private Sub Mst_data_list_truong_CBB_Change()
    Mst_data_col = WorksheetFunction.Match(Mst_data_list_truong_CBB.Value, ThisWorkbook.Sheets("Master data").Range("A10:AW10"), 0)
End Sub


Private Sub MST_DT_ADD_Finish_OnClick()
    Them_Master_data.Hide
End Sub

Private Sub Them_Masterdata_OK_OnClick()
Dim i As Long, lr As Long, sh As Worksheet, col As Long
Set sh = ThisWorkbook.Sheets("Master data")
BatLimit
    If Mst_data_list_truong_CBB.Value <> "" And TextBox1.Value <> "" Then
    lr = sh.Cells(10, Mst_data_col + 49).End(xlDown).Row + 1
    If lr > 104857600 Or sh.Cells(11, Mst_data_col + 49).Value = "" Then lr = 11
    sh.Cells(lr, Mst_data_col + 49) = TextBox1.Value
    Cells(lr + 1, Mst_data_col + 49).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Cells(11, Mst_data_col).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Cells(lr + 2, Mst_data_col + 49).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Else
    MsgBox "You need to add correct value to these items"
    End If
TatLimit
End Sub


Private Sub UserForm_Initialize()
    Dim a As Object
    Dim LastCol As Integer
            LastCol = ThisWorkbook.Sheets("Master Data").Cells(10, 2).End(xlToRight).Column
    
    For Each a In ThisWorkbook.Sheets("Master data").Range(ThisWorkbook.Sheets("Master data").Cells(10, 2), ThisWorkbook.Sheets("Master data").Cells(10, LastCol))
    Mst_data_list_truong_CBB.AddItem a
    Next
End Sub

Private Sub Bot_Masterdata_OK_OnClick()
Dim i As Long, lr As Long, sh As Worksheet, col As Long
Set sh = ThisWorkbook.Sheets("Master data")
BatLimit
    If Mst_data_list_truong_CBB.Value <> "" And TextBox1.Value <> "" Then
    lr = sh.Cells(10, Mst_data_col + 49).End(xlDown).Row + 1
    If lr > 104857600 Or sh.Cells(11, Mst_data_col + 49).Value = "" Then lr = 11
    sh.Cells(lr, Mst_data_col + 49) = TextBox1.Value
    Cells(lr + 1, Mst_data_col + 49).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Cells(11, Mst_data_col).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Cells(lr + 2, Mst_data_col + 49).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Else
    MsgBox "You need to add correct value to these items"
    End If
TatLimit
End Sub

