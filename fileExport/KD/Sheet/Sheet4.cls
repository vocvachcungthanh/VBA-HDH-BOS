VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
  
   Application.ScreenUpdating = False
   Call ScrollToTop
   Call hideall
   Call VeBieuDo_BaoCaoDoanhThuTheo_NVKD
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

Public Function ResetNhom1()
    Sheet15.Range("B9").Value = StartRecord(txtNhom1.Value, 10)
   
    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet15.Range("F9").Value
    CotCuoi = Sheet15.Range("J9").Value

    PhamViResize = "B11:" & CotCuoi & 11 + dongCuoi
    
    Call UpdateChartDataRange(Sheet4, "ChartDoanhThuTheoNVKD", Sheet15, PhamViResize)
End Function

Private Sub txtNhom1_Change()

    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet9, "F9")
    With txtNhom1
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With

    Call ResetNhom1
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

Public Function ResetNhom2()
    Sheet15.Range("L9").Value = StartRecord(txtNhom2.Value, 10)
   Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String

    dongCuoi = Sheet15.Range("P9").Value
    CotCuoi = Sheet15.Range("T9").Value

    PhamViResize = "L11:" & CotCuoi & 11 + dongCuoi
    
    Call UpdateChartDataRange(Sheet4, "ChartSoLuongBanTheoNVKD", Sheet15, PhamViResize)
End Function

Private Sub txtNhom2_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet9, "W9")
    With txtNhom2
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With
   
   Call ResetNhom2
End Sub

