VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
   Call ScrollToTop
   Call hideall
    Rows("6:10").Select
    Selection.RowHeight = 0

    Call F_FormatNgayLamViec(Sheet23)
End Sub


