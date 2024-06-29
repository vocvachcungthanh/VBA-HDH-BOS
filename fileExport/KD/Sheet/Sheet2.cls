VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall
    Call VeBieuDo_BaoCao_DoanhThuTheo_SP
End Sub

'Nhom 1
'KD.xlsb!Sheet8.NextPage1
Private Sub NextPage1()
    With txtBoxPhanTrangNhom1
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage1()
    With txtBoxPhanTrangNhom1
        If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1
    End With
End Sub

Public Function ResizeNhom1()
    Sheet12.Range("B9") = StartRecord(Sheet2.txtBoxPhanTrangNhom1.Value, 10)

    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("F9").Value
    CotCuoi = Sheet12.Range("J9").Value

    PhamViResize = "B11:" & CotCuoi & 11 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTTL", Sheet12, PhamViResize)
End Function

Private Sub txtBoxPhanTrangNhom1_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "E6")
    With txtBoxPhanTrangNhom1
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With
    
    Call ResizeNhom1
End Sub

'Nhom 2
Private Sub NextPage2()
    With txtBoxPhanTrangNhom2
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage2()
    With txtBoxPhanTrangNhom2
          If .Value = 1 Then
            Exit Sub
          End If
        .Value = .Value - 1
    End With
End Sub

Public Function ResizeNhom2()
    Sheet12.Range("L9") = StartRecord(Sheet2.txtBoxPhanTrangNhom2.Value, 10)

    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("P9").Value
    CotCuoi = Sheet12.Range("T9").Value

    PhamViResize = "L11:" & CotCuoi & 11 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTHeoSeries", Sheet12, PhamViResize)
End Function

Private Sub txtBoxPhanTrangNhom2_Change()

    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "R6")
    With txtBoxPhanTrangNhom2
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With

    Call ResizeNhom2
End Sub

'Nhom 3
'KD.xlsb!Sheet8.NextPage3
Private Sub NextPage3()
    With txtBoxPhanTrangNhom3
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage3()
    With txtBoxPhanTrangNhom3
        If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1
    End With
End Sub

Public Function ResizeNhom3()
   Sheet12.Range("B27") = StartRecord(Sheet2.txtBoxPhanTrangNhom3.Value, 10)

    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("F27").Value
    CotCuoi = Sheet12.Range("J27").Value

    PhamViResize = "B29:" & CotCuoi & 29 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartDanhThuTheoNhomSP", Sheet12, PhamViResize)
End Function

Private Sub txtBoxPhanTrangNhom3_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "AE6")
    With txtBoxPhanTrangNhom3
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With

    Call ResizeNhom3
End Sub

'Nhom 4

'KD.xlsb!Sheet2.NextPage4
Private Sub NextPage4()
    With txtBoxPhanTrangNhom4
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage4()
    With txtBoxPhanTrangNhom4
        If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1
    End With
End Sub

Public Function ResizeNhom4()
   Sheet12.Range("L27") = StartRecord(Sheet2.txtBoxPhanTrangNhom4.Value, 10)

    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("P27").Value
    CotCuoi = Sheet12.Range("T27").Value

    PhamViResize = "L29:" & CotCuoi & 29 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTTheoNhomLN", Sheet12, PhamViResize)
End Function

Private Sub txtBoxPhanTrangNhom4_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "AQ6")
    With txtBoxPhanTrangNhom4
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With

  Call ResizeNhom4
End Sub

'Nhom 5
'KD.xlsb!Sheet2.NextPage5
Private Sub NextPage5()
    With txtBoxPhanTrangNhom5
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage5()
    With txtBoxPhanTrangNhom5
        If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1
    End With
End Sub

Public Function ResizeNhom5()
   Sheet12.Range("L27") = StartRecord(Sheet2.txtBoxPhanTrangNhom4.Value, 10)

    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("P27").Value
    CotCuoi = Sheet12.Range("T27").Value

    PhamViResize = "L29:" & CotCuoi & 29 + dongCuoi

    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTTheoNhomLN", Sheet12, PhamViResize)
End Function


Private Sub txtBoxPhanTrangNhom5_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "BD6")
    With txtBoxPhanTrangNhom5
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With
    
    Call ResizeNhom5
End Sub

'Nhom 6

Private Sub NextPage6()
    With txtBoxPhanTrangNhom6
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage6()
    With txtBoxPhanTrangNhom6
        If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1
    End With
End Sub

Public Function ResizeNhom6()
   Sheet12.Range("L45") = StartRecord(Sheet2.txtBoxPhanTrangNhom6.Value, 10)

    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet12.Range("P45").Value
    CotCuoi = Sheet12.Range("T45").Value

    PhamViResize = "L47:" & CotCuoi & 47 + dongCuoi
    Call UpdateChartDataRange(Sheet2, "ChartBaoCaoDTTheoSX/NK", Sheet12, PhamViResize)
End Function

Private Sub txtBoxPhanTrangNhom6_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "BQ6")
    With txtBoxPhanTrangNhom6
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With

   Call ResizeNhom6
End Sub

