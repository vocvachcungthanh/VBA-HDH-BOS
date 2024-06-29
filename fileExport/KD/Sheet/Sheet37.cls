VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private isActive As Boolean

Private Sub cbbKH_Change()
    With cbbKH
        Sheet37.Range("J7") = .Value
    End With
End Sub

Private Sub cbbNam_Change()
    With cbbNam
        Sheet37.Range("L7") = .Value
    End With
End Sub

Private Sub Worksheet_Activate()
   Call ScrollToTop
   Call hideall
   
   If isActive = False Then
        Call HienThiDoanhThuNgayTheo_KH
   End If
   
   isActive = True
   
  Call F_FormatNgayLamViec(Sheet37)
End Sub

