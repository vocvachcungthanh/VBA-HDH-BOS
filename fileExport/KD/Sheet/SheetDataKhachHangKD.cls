VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetDataKhachHangKD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim isActive As Boolean

Private Sub Worksheet_Activate()
    Call ScrollToTop
    
    If isActive = False Then
        Call ChayNeuActiveLanDau
        Call LamMoiDuLieuKH
    End If
    
    isActive = True
    Call hideall
End Sub

Sub ChayNeuActiveLanDau()
    BatLimit
    Set wSheet = SheetDataKhachHangKD
    With wSheet
        .Select
        If .Range("C12") <> "" Then
            ReSizeTableKH
            CongThucTinhKH
        End If
    End With
    Set wSheet = Nothing
    
    TatLimit
End Sub
