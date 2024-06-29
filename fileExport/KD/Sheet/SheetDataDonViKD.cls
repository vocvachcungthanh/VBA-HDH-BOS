VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetDataDonViKD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Active As Boolean

Private Sub Worksheet_Activate()
    Call ScrollToTop
    Call hideall
    
    If tt = Active Then
        Call LamMoiDuLieuDV
    End If
    Active = True
End Sub
