VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1611"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim isActive As Boolean

Public Sub Worksheet_Activate()
    Call ScrollToTop
    
    If isActive = False Then
        ChayLanDauKhiActive
        Call cmdLoadData
    End If

    isActive = True
    Call F_KhoiTaoNam(Sheet1611.cbbNam)
    Call F_KhoiTaoNam(Sheet1611.cbbNamLichSu)
    Sheet1611.cbbNam.Text = Sheet11.cbbSheetNam.Value
     Call hideall
End Sub

Sub ChayLanDauKhiActive()
    BatLimit
    Set wSheet = Sheet1611
    With wSheet
        .Select
        dongCuoi = tinhdongcuoi("B12:B1048576")
        TinhDoanhThu dongCuoi
        If .Range("B12") <> "" Then
            ActiveSheet.ListObjects("TableSanPham").Resize Range("B11:J" & dongCuoi)
        End If
    End With
    Set wSheet = Nothing
    TatLimit
End Sub

