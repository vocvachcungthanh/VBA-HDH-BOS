Public Function KhoiTaoCbbBoxPageBcDtCacDonViKD()
    Dim TongDuLieu As Double

    TongDuLieu = Sheet26.Range("E6").value
    Call cbbPage(TongDuLieu, 10, Sheet2.cbbPhanTrangNhom1)

    TongDuLieu = Sheet26.Range("R6").value
    Call cbbPage(TongDuLieu, 10, Sheet2.cbbPhanTrangNhom2)

    TongDuLieu = Sheet26.Range("AE6").value
    Call cbbPage(TongDuLieu, 10, Sheet2.cbbPhanTrangNhom3)

    TongDuLieu = Sheet26.Range("AO6").value
    Call cbbPage(TongDuLieu, 10, Sheet2.cbbPhanTrangNhom4)

    TongDuLieu = Sheet26.Range("AB6").value
    Call cbbPage(TongDuLieu, 10, Sheet2.cbbPhanTrangNhom5)

    TongDuLieu = Sheet26.Range("BM6").value
    Call cbbPage(TongDuLieu, 10, Sheet2.cbbPhanTrangNhom6)
End Function