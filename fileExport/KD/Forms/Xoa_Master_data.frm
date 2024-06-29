VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Xoa_Master_data 
   Caption         =   "Delete master data"
   ClientHeight    =   1425
   ClientLeft      =   -60
   ClientTop       =   -210
   ClientWidth     =   1590
   OleObjectBlob   =   "Xoa_Master_data.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Xoa_Master_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public selected_index As Long, Mst_data_col_xoa As Long
Private Sub ListBox1_Change()
    selected_index = ListBox1.ListIndex
    Mst_data_col_xoa = WorksheetFunction.Match(Mst_data_list_truong_CBB.Value, ThisWorkbook.Sheets("Master data").Range("A10:AW10"), 0)
End Sub

Private Sub Mst_data_list_truong_CBB_Change()
    Dim lr As Long, sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Master Data")
    Mst_data_col = WorksheetFunction.Match(Mst_data_list_truong_CBB.Value, ThisWorkbook.Sheets("Master data").Range("A10:AW10"), 0)
    lr = sh.Cells(10, Mst_data_col + 49).End(xlDown).Row
    If sh.Cells(11, Mst_data_col + 49).Value <> "" Then
        ListBox1.Clear
        For Each a In ThisWorkbook.Sheets("Master data").Range(sh.Cells(11, Mst_data_col + 49), sh.Cells(lr, Mst_data_col + 49))
            ListBox1.AddItem a 'RowSource = ThisWorkbook.Sheets("Master Data").Range("BA11:BA13")
            Next
    Else
        ListBox1.Clear
    End If
    selected_index = 1048576
End Sub

Private Sub MST_DT_ADD_Finish_OnClick()
    Xoa_Master_data.Hide
End Sub

Private Sub UserForm_Initialize()
    Dim a As Object
    Dim LastCol As Integer
            LastCol = ThisWorkbook.Sheets("Master Data").Cells(10, 2).End(xlToRight).Column
    
    For Each a In ThisWorkbook.Sheets("Master data").Range(ThisWorkbook.Sheets("Master data").Cells(10, 2), ThisWorkbook.Sheets("Master data").Cells(10, LastCol))
    Mst_data_list_truong_CBB.AddItem a
    Next
End Sub

Private Sub Xoa_Masterdata_OK_OnClick()
    If Mst_data_list_truong_CBB.Value <> "" And selected_index <> 1048576 Then
        ThisWorkbook.Sheets("Master Data").Cells(11 + selected_index, Mst_data_col_xoa + 49).Select
        Selection.Delete Shift:=xlUp
        Mst_data_list_truong_CBB_Change
    Else
        MsgBox "You need to choose one of items from combobox"
    End If
End Sub
