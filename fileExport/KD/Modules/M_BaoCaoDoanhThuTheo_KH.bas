Attribute VB_Name = "M_BaoCaoDoanhThuTheo_KH"
Sub VeBieuDoBaoCaoDoanhThuTheo_KH()
    BatLimit
    Call F_R_DATA
    Call F_RESET_PIVOT_KH

    ActiveWorkbook.RefreshAll

    Sheet20.Select
    'Ban Do nhiet doanh thu theo NVKD

    Call select_data("Chart 47", "Table17[#All]", "Pivot KH")
    Call DinhDangBdNhiet("Chart 47")

    'Ban Do nhiet so luong theo NVKD
    Call select_data("Chart 48", "Table1719[#All]", "Pivot KH")
    Call DinhDangBdNhiet("Chart 48")
    Sheet20.Select
    
    With Sheet20
        .txtNhom1.Value = 1
        .txtNhom2.Value = 1
        
        Call .ResizeNhom1
        Call .ResizeNhom2
    End With
    TatLimit
    ThongBao_ThanhCong
End Sub

Public Sub F_RESET_PIVOT_KH()
    Dim dongCuoi As Long
    Set wSheet = Sheet21
    With wSheet
        .Select

        dongCuoi = tinhdongcuoi("B12:B1048576")
        .ListObjects("Table17").Resize Range("$L$11:$M$" & dongCuoi)

        dongCuoi = tinhdongcuoi("O12:O1048576")
        .ListObjects("Table1719").Resize Range("$Y$11:$Z$" & dongCuoi)
    End With
    Set wSheet = Nothing
End Sub

Public Function KhoiTaoCbbBoxPage()
    'tao cbbDanhThuTheoKH
    Dim TongDuLieu As Double

    TongDuLieu = Sheet21.Range("F9")
    Call cbbPage(TongDuLieu, 10, Sheet20.cbbDoanhThuTheoKH)
    
    TongDuLieu = Sheet21.Range("S9")
    Call cbbPage(TongDuLieu, 10, Sheet20.cbbSoLuongTheoKH)
End Function
