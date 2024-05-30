Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall
    Call KhoiTaoCbbBoxPage
End Sub

Function KhoiTaoCbbBoxPage()
    'tao cbbDanhThuTheoKH
    Dim TongDuLieu As Double

    TongDuLieu = Sheet21.Range("F9")
    Call cbbPage(TongDuLieu, 10, Sheet20.cbbDoanhThuTheoKH)

    TongDuLieu = Sheet21.Range("S9")
    Call cbbPage(TongDuLieu, 10, Sheet20.cbbSoLuongTheoKH)
End Function

Private Sub cbbDoanhThuTheoKH_Click()
    Sheet17.Range("B9") = StartRecord(Sheet20.cbbDoanhThuTheoKH.value, 10)
End Sub

Private Sub cbbSoLuongTheoKH_Click()
    Sheet17.Range("G9") = StartRecord(Sheet20.cbbSoLuongTheoKH.value, 10)
End Sub