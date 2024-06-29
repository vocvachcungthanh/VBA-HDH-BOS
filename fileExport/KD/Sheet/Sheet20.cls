VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall
    
    Call VeBieuDoBaoCaoDoanhThuTheo_KH
End Sub

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

Public Function ResizeNhom1()
    Sheet17.Range("B9") = StartRecord(Sheet20.txtNhom1.Value, 10)
    
    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet17.Range("F9").Value
    CotCuoi = Sheet17.Range("J9").Value

    PhamViResize = "B11:" & CotCuoi & 11 + dongCuoi

    
    Call UpdateChartDataRange(Sheet20, "ChartDoanhThuTheoKH", Sheet17, PhamViResize)
End Function

Private Sub txtNhom1_Change()
      Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet21, "F9")
    With txtNhom1
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With

    Call ResizeNhom1
End Sub

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

Public Function ResizeNhom2()
     Sheet17.Range("L9") = StartRecord(Sheet20.txtNhom2.Value, 10)
    
    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet17.Range("P9").Value
    CotCuoi = Sheet17.Range("T9").Value

    PhamViResize = "L11:" & CotCuoi & 11 + dongCuoi
    
    Call UpdateChartDataRange(Sheet20, "ChartSoLuongTheoKH", Sheet17, PhamViResize)
End Function

Private Sub txtNhom2_Change()
   Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet21, "S9")
    With txtNhom2
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With

   Call ResizeNhom2
End Sub
