Sub VeBieuDo_BaoCaoSoLuongTheo_SP()
    BatLimit
    F_R_DATA
    Set wSheet = Sheet8
    With Sheet8
        .Select
        ActiveWorkbook.RefreshAll
    End With
    Set wSheet = Nothing
    Call KhoiTaoCbbBoxPageSLSP()
    TatLimit

    ThongBao_ThanhCong
End Sub

Public Function KhoiTaoCbbBoxPageSLSP()
    'tao cbbDanhThuTheoSLSP
    Dim TongDuLieu As Double

    ' Cbb Nhom_VTHH_1_TTT_SL
    TongDuLieu = Sheet26.Range("CC6")
    Call cbbPage(TongDuLieu, 10, Sheet8.cbbDoanhThuTheoSPN1)

    ' Cbb Nhom_VTHH_2_TTT_SL
    TongDuLieu = Sheet26.Range("CL6")
    Call cbbPage(TongDuLieu, 10, Sheet8.cbbDoanhThuTheoSPN2)

    ' Cbb Nhom_VTHH_3_TTT_SL
    TongDuLieu = Sheet26.Range("CU6")
    Call cbbPage(TongDuLieu, 10, Sheet8.cbbDoanhThuTheoSPN3)

    ' Cbb Nhom_VTHH_4_TTT_SL
    TongDuLieu = Sheet26.Range("DD6")
    Call cbbPage(TongDuLieu, 10, Sheet8.cbbDoanhThuTheoSPN4)

    ' Cbb Nhom_VTHH_5_TTT_SL
    TongDuLieu = Sheet26.Range("DM6")
    Call cbbPage(TongDuLieu, 10, Sheet8.cbbDoanhThuTheoSPN5)

    ' Cbb Nhom_VTHH_6_TTT_SL
    TongDuLieu = Sheet26.Range("DV6")
    Call cbbPage(TongDuLieu, 10, Sheet8.cbbDoanhThuTheoSPN6)
End Function