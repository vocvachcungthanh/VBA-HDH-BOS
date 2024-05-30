Private Sub Worksheet_Activate()
    Call ScrollToTop
    Call hideall
    Call KhoiTaoCbbBoxPageSLSP
End Sub

Public  Sub cbbDoanhThuTheoSPN1_Click()
    Sheet18.Range("B9") = StartRecord(Sheet8.cbbDoanhThuTheoSPN1.value, 10)
End Sub

Public  Sub cbbDoanhThuTheoSPN2_Click()
    Sheet18.Range("G9") = StartRecord(Sheet8.cbbDoanhThuTheoSPN2.value, 10)
End Sub

Public  Sub cbbDoanhThuTheoSPN3_Click()
    Sheet18.Range("B27") = StartRecord(Sheet8.cbbDoanhThuTheoSPN3.value, 10)
End Sub

Public  Sub cbbDoanhThuTheoSPN4_Click()
    Sheet18.Range("G27") = StartRecord(Sheet8.cbbDoanhThuTheoSPN4.value, 10)
End Sub

Public  Sub cbbDoanhThuTheoSPN5_Click()
    Sheet18.Range("B45") = StartRecord(Sheet8.cbbDoanhThuTheoSPN5.value, 10)
End Sub

Public  Sub cbbDoanhThuTheoSPN6_Click()
    Sheet18.Range("G45") = StartRecord(Sheet8.cbbDoanhThuTheoSPN6.value, 10)
End Sub