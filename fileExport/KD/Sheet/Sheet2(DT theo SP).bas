Private isActive As Boolean

Private Sub Worksheet_Activate()

    Call KhoiTaoCbbBoxPageBcDtTheoSP

    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall

    If isActive = False Then
        '        Call VeBieuDo_BaoCao_DoanhThuTheo_SP
    End If

    isActive = True
End Sub

'Nhom 1
Private Sub cbbPhanTrangNhom1_Click()
    Sheet12.Range("B9") = StartRecore(Sheet2.cbbPhanTrangNhom1.value, 10)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("E9").value
    PhamViResize = "B11:H" & 11 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTTL", Sheet12, PhamViResize)
End Sub

'Nhom 2
Private Sub cbbPhanTrangNhom2_Click()
    Sheet12.Range("G9") = StartRecore(Sheet2.cbbPhanTrangNhom1.value, 10)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("J9").value
    PhamViResize = "G11:J" & 11 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTHeoSeries", Sheet12, PhamViResize)
End Sub

'Nhom 3
Private Sub cbbPhanTrangNhom3_Click()
    Sheet12.Range("B27") = StartRecore(Sheet2.cbbPhanTrangNhom3.value, 10)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("E27").value
    PhamViResize = "B29:E" & 29 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartDanhThuTheoNhomSP", Sheet12, PhamViResize)
End Sub

'Nhom 4
Private Sub cbbPhanTrangNhom4_Click()
    Sheet12.Range("G27") = StartRecore(Sheet2.cbbPhanTrangNhom4.value, 10)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("J27").value
    PhamViResize = "G29:J" & 29 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTTheoNhomLN", Sheet12, PhamViResize)
End Sub

'Nhom 5
Private Sub cbbPhanTrangNhom5_Click()
    Sheet12.Range("B45") = StartRecore(Sheet2.cbbPhanTrangNhom5.value, 10)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("E45").value
    PhamViResize = "B247:E" & 47 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTTheoNH", Sheet12, PhamViResize)
End Sub

'Nhom 6
Private Sub cbbPhanTrangNhom6_Click()
    Sheet12.Range("G45") = StartRecore(Sheet2.cbbPhanTrangNhom6.value, 10)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("J45").value
    PhamViResize = "G47:J" & 247 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTTheoSX/NK", Sheet12, PhamViResize)
End Sub