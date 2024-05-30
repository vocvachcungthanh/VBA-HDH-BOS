
Public Sub VeBieuDo_BaoCaoLoiNhuanTheo_SP()
    BatLimit
    Call KhoiTaoCbbBoxPageLNTSP
    TatLimit

    ThongBao_ThanhCong
End Sub

Public Function KhoiTaoCbbBoxPageLNTSP()
    'tao cbbDanhThuTheoSLSP
    Dim TongDuLieu As Double

    ' Cbb Nhom_TY_LE_LN_1
    TongDuLieu = Sheet26.Range("EZ6")
    Call cbbPage(TongDuLieu, 10, Sheet10.cbbPhanTrangLNNhom1)

    ' Cbb Nhom_TY_LE_LN_2
    TongDuLieu = Sheet26.Range("FJ6")
    Call cbbPage(TongDuLieu, 10, Sheet10.cbbPhanTrangLNNhom2)

    ' Cbb Nhom_TY_LE_LN_3
    TongDuLieu = Sheet26.Range("FT6")
    Call cbbPage(TongDuLieu, 10, Sheet10.cbbPhanTrangLNNhom3)

    ' Cbb Nhom_TY_LE_LN_4
    TongDuLieu = Sheet26.Range("GD6")
    Call cbbPage(TongDuLieu, 10, Sheet10.cbbPhanTrangLNNhom4)

    ' Cbb Nhom_TY_LE_LN_5
    TongDuLieu = Sheet26.Range("GN6")
    Call cbbPage(TongDuLieu, 10, Sheet10.cbbPhanTrangLNNhom5)

    'Cbb Nhom_TY_LE_LN_6
    TongDuLieu = Sheet26.Range("GX6")
    Call cbbPage(TongDuLieu, 10, Sheet10.cbbPhanTrangLNNhom6)

End Function
