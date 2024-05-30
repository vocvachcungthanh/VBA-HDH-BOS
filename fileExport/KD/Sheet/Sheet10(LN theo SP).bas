Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall

    Call KhoiTaoCbbBoxPageLNTSP
End Sub

Private Sub cbbPhanTrangLNNhom1_Click()
    Dim dongCuoi As Integer
    Dim PhamViResize As String

    Sheet19.Range("B9") = StartRecord(cbbPhanTrangLNNhom1.value, 10)

    dongCuoi = Sheet19.Range("F9").value

    PhamViResize = "B11:F" & 11 + dongCuoi
    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom1", Sheet19, PhamViResize)

End Sub

Private Sub cbbPhanTrangLNNhom2_Click()
    Dim dongCuoi As Integer
    Dim PhamViResize As String

    Sheet19.Range("H9") = StartRecord(cbbPhanTrangLNNhom2.value, 10)

    dongCuoi = Sheet19.Range("L9").value

    PhamViResize = "H11:L" & 11 + dongCuoi
    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom2", Sheet19, PhamViResize)

End Sub

Private Sub cbbPhanTrangLNNhom3_Click()
    Dim dongCuoi As Integer
    Dim PhamViResize As String

    Sheet19.Range("B27") = StartRecord(cbbPhanTrangLNNhom3.value, 10)

    dongCuoi = Sheet19.Range("F27").value
    PhamViResize = "B29:F" & 29 + dongCuoi

    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom3", Sheet19, PhamViResize)

End Sub

Private Sub cbbPhanTrangLNNhom4_Click()
    Dim dongCuoi As Integer
    Dim PhamViResize As String

    Sheet19.Range("H27") = StartRecord(cbbPhanTrangLNNhom4.value, 10)

    dongCuoi = Sheet19.Range("L27").value
    PhamViResize = "H29:L" & 29 + dongCuoi

    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom4", Sheet19, PhamViResize)
End Sub

Private Sub cbbPhanTrangLNNhom5_Click()
    Dim dongCuoi As Integer
    Dim PhamViResize As String

    Sheet19.Range("B45") = StartRecord(cbbPhanTrangLNNhom5.value, 10)

    dongCuoi = Sheet19.Range("F45").value
    PhamViResize = "B47:F" & 47 + dongCuoi

    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom5", Sheet19, PhamViResize)
End Sub

Private Sub cbbPhanTrangLNNhom6_Click()
    Dim dongCuoi As Integer
    Dim PhamViResize As String

    Sheet19.Range("H45") = StartRecord(cbbPhanTrangLNNhom6.value, 10)

    dongCuoi = Sheet19.Range("L45").value
    PhamViResize = "H47:L" & 47 + dongCuoi

    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom6", Sheet19, PhamViResize)
End Sub
