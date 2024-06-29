VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetTheoDoiBienLoiNhuan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private isActive As Boolean

Private Sub Worksheet_Activate()
    Call ScrollToTop
    Call ChonNam
    Call hideall
    
    If isActive = False Then
        Call HienThiBienLoiNhuan
    End If
    
    isActive = True
End Sub


