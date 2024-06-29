Attribute VB_Name = "M_BaoCaoDoanhThuTheo_NVKD"
Sub VeBieuDo_BaoCaoDoanhThuTheo_NVKD()
    BatLimit
    Call F_R_DATA
    Call F_RESET_PIVOT_NVKD
    ActiveWorkbook.RefreshAll
    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String
    ThisWorkbook.Sheets("DT theo NVKD").Select

    'Ban Do nhiet doanh thu theo NVKD
    dongCuoi = Sheet9.Range("F9").Value
    
    PhamViResize = "P11:Q" & dongCuoi + 11
    
    Call UpdateChartDataRange(Sheet4, "Chart 50", Sheet9, PhamViResize)
    Call DinhDangBdNhiet("Chart 50")

    'Ban Do nhiet so luong theo NVKD
    Call select_data("Chart 49", "Table1517[#All]", "Pivot NVKD")
    dongCuoi = Sheet9.Range("W9").Value
    
    PhamViResize = "AC11:AD" & dongCuoi + 11
    
    Call UpdateChartDataRange(Sheet4, "Chart 49", Sheet9, PhamViResize)
    Call DinhDangBdNhiet("Chart 49")

    With Sheet4
        .txtNhom1.Value = 1
        .txtNhom2.Value = 2
        Call .ResetNhom1
        Call .ResetNhom2
    End With
    TatLimit
    
    ThongBao_ThanhCong
End Sub

Public Sub F_RESET_PIVOT_NVKD()
    Dim dongCuoi As Long
    Set wSheet = Sheet9
    With wSheet
        .Select
'        DongCuoi = tinhdongcuoi("B12:B1048576")
'        .ListObjects("Table15").Resize Range("$L$11:$M$" & DongCuoi)
'
'        DongCuoi = tinhdongcuoi("O12:O1048576")
'        .ListObjects("Table1517").Resize Range("$Y$11:$Z$" & DongCuoi)

    End With
    Set wSheet = Nothing
End Sub

Function KhoiTaoCbbBoxPageNVKD()
    Dim TongDuLieu As Double
    TongDuLieu = Sheet9.Range("F9").Value
    Call cbbPage(TongDuLieu, 10, Sheet4.cbbPhanTrangDTNVKD)

    TongDuLieu = Sheet9.Range("W9").Value
    Call cbbPage(TongDuLieu, 10, Sheet4.cbbPhanTrangSoLuongBanTNVKD)
End Function
