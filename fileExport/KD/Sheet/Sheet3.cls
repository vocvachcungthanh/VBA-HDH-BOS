VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private isActive As Boolean

Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall
    
    If isActive = False Then
'        Call VeBieuDoThoiGian
    End If
    
    isActive = True
End Sub

