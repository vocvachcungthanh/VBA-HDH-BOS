Private Sub txtNhom1_Change()

End Sub

Private Sub Worksheet_Activate()
    Call ScrollToTop
    Call hideall
End Sub

'Nhom1
Private Sub NextPage1()
    With txtNhom1
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage1()
    With txtNhom1
        .Value = .Value - 1

    End With
End Sub

Public Sub txtNhom1_Click()

    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "CC6")
    With txtNhom1
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With
    Sheet18.Range("B9") = StartRecord(Sheet8.txtNhom1.Value, 10)

    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet18.Range("F9").Value
    CotCuoi = Sheet18.Range("J9").Value

    PhamViResize = "B11:" & CotCuoi & 11 + dongCuoi
    Call UpdateChartDataRange(Sheet8, "ChartBaoCaoSLTheoNhomVTHH1", Sheet18, PhamViResize)
End Sub

'Nhom2
Private Sub NextPage2()
    With txtNhom2
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage2()
    With txtNhom2
        .Value = .Value - 1

    End With
End Sub

Public Sub txtNhom2_Click()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "CC6")
    With txtNhom2
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
        Sheet18.Range("L9") = StartRecord(Sheet8.txtNhom2.Value, 10)

        Dim dongCuoi As Integer
        Dim CotCuoi As String
        Dim PhamViResize As String

        dongCuoi = Sheet18.Range("P9").Value
        CotCuoi = Sheet18.Range("T9").Value

        PhamViResize = "L11:" & CotCuoi & 11 + dongCuoi

        Call UpdateChartDataRange(Sheet8, "ChartBaoCaoSLTheoNhomVTHH2", Sheet18, PhamViResize)
End Sub

'Nhom3
Private Sub NextPage3()
    With txtNhom3
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage3()
    With txtNhom3
        .Value = .Value - 1
    End With
End Sub

Public Sub txtNhom3_Click()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "CU6")
    With txtNhom3
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
        Sheet18.Range("B27") = StartRecord(Sheet8.txtNhom3.Value, 10)

        Dim dongCuoi As Integer
        Dim CotCuoi As String
        Dim PhamViResize As String

        dongCuoi = Sheet18.Range("F27").Value
        CotCuoi = Sheet18.Range("J27").Value

        PhamViResize = "B29:" & CotCuoi & 29 + dongCuoi

        Call UpdateChartDataRange(Sheet8, "ChartBaoCaoSLTheoNhomVTHH3", Sheet18, PhamViResize)
End Sub

'Nhom 4
Private Sub NextPage4()
    With txtNhom4
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage4()
    With txtNhom4
        .Value = .Value - 1
    End With
End Sub

Public Sub txtNhom4_Click()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "DD6")
    With txtNhom4
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If

        Sheet18.Range("L27") = StartRecord(Sheet8.txtNhom4.Value, 10)

        Dim dongCuoi As Integer
        Dim CotCuoi As String
        Dim PhamViResize As String

        dongCuoi = Sheet18.Range("P27").Value
        CotCuoi = Sheet18.Range("T27").Value

        PhamViResize = "L29:" & CotCuoi & 29 + dongCuoi

        Call UpdateChartDataRange(Sheet8, "ChartBaoCaoSLTheoNhomVTHH4", Sheet18, PhamViResize)
End Sub

'Nhom 5
Private Sub NextPage5()
    With txtNhom5
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage5()
    With txtNhom5
        .Value = .Value - 1
    End With
End Sub

Public Sub txtNhom5_Click()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "DM6")
    With txtNhom5
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
        Sheet18.Range("B45") = StartRecord(Sheet8.txtNhom5.Value, 10)

        Dim dongCuoi As Integer
        Dim CotCuoi As String
        Dim PhamViResize As String

        dongCuoi = Sheet18.Range("F45").Value
        CotCuoi = Sheet18.Range("J45").Value

        PhamViResize = "B47:" & CotCuoi & 47 + dongCuoi

        Call UpdateChartDataRange(Sheet8, "ChartBaoCaoSLTheoNhomVTHH5", Sheet18, PhamViResize)
End Sub

'Nhom 6
Private Sub NextPage6()
    With txtNhom6
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage6()
    With txtNhom6
        .Value = .Value - 1
    End With
End Sub

Public Sub txtNhom6_Click()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "DV6")
    With txtNhom6
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
        Sheet18.Range("L45") = StartRecord(Sheet8.txtNhom6.Value, 10)

        Dim dongCuoi As Integer
        Dim CotCuoi As String
        Dim PhamViResize As String

        dongCuoi = Sheet18.Range("P45").Value
        CotCuoi = Sheet18.Range("T45").Value

        PhamViResize = "L47:" & CotCuoi & 47 + dongCuoi

        Call UpdateChartDataRange(Sheet8, "ChartBaoCaoSLTheoNhomVTHH6", Sheet18, PhamViResize)
End Sub
