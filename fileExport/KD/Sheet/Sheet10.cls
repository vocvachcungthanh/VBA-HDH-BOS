VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall
    Call VeLaiDoThi
    
 
End Sub

Sub VeLaiDoThi()
    txtNhom1.Value = 1
    txtNhom1.Value = 1
    txtNhom1.Value = 1
    txtNhom1.Value = 1
    txtNhom1.Value = 1
    txtNhom1.Value = 1
    
    Call ReszieNhom1
    Call ReszieNhom2
    Call ReszieNhom3
    Call ReszieNhom4
    Call ReszieNhom5
    Call ReszieNhom6
    
    Call ThongBao_ThanhCong
End Sub

'Nhom 1
'KD.xlsb!Sheet10.NextPage1
Private Sub NextPage1()
    With txtNhom1
        .Value = .Value + 1
    End With
End Sub

'KD.xlsb!Sheet10.PrevPage1
Private Sub PrevPage1()
    With txtNhom1
         If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1

    End With
End Sub

Public Function ReszieNhom1()
    Sheet19.Range("B9") = StartRecord(txtNhom1.Value, 10)
    Dim dongCuoi As Integer
    Dim PhamViResize As String
    dongCuoi = Sheet19.Range("F9").Value

    PhamViResize = "B11:F" & 11 + dongCuoi
    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom1", Sheet19, PhamViResize)
End Function

Private Sub txtNhom1_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "FD6")
    With txtNhom1
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With
    
   ReszieNhom1

End Sub


'Nhom 2
'KD.xlsb!Sheet10.NextPage2
Private Sub NextPage2()
    With txtNhom2
        .Value = .Value + 1
    End With
End Sub

'KD.xlsb!Sheet10.PrevPage2
Private Sub PrevPage2()
    With txtNhom2
         If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1

    End With
End Sub

Public Function ReszieNhom2()
    Sheet19.Range("H9") = StartRecord(txtNhom2.Value, 10)
        
    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet19.Range("L9").Value

    PhamViResize = "H11:L" & 11 + dongCuoi
    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom2", Sheet19, PhamViResize)
End Function

Private Sub txtNhom2_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "FN6")
    With txtNhom1
        If CInt(.Value) > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With
    
    Call ReszieNhom2
End Sub

'Nhom 3
'KD.xlsb!Sheet10.NextPage3
Private Sub NextPage3()
    With txtNhom3
        .Value = .Value + 1
    End With
End Sub

'KD.xlsb!Sheet10.PrevPage3
Private Sub PrevPage3()
    With txtNhom3
         If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1

    End With
End Sub

Public Function ReszieNhom3()
   Sheet19.Range("B27") = StartRecord(txtNhom3.Value, 10)
    Dim dongCuoi As Integer
    Dim PhamViResize As String
    
    dongCuoi = Sheet19.Range("F27").Value
    PhamViResize = "B29:F" & 29 + dongCuoi
    
    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom3", Sheet19, PhamViResize)
End Function

Private Sub txtNhom3_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "FX6")
    With txtNhom1
        If CInt(.Value) > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If CInt(.Value) < 1 Then
            .Value = 1
        End If
    End With
    
    Call ReszieNhom3

End Sub


'Nhom 4
'KD.xlsb!Sheet10.NextPage4
Private Sub NextPage4()
    With txtNhom4
        .Value = .Value + 1
    End With
End Sub

'KD.xlsb!Sheet10.PrevPage4
Private Sub PrevPage4()
    With txtNhom4
        If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1

    End With
End Sub

Public Function ReszieNhom4()
   Sheet19.Range("H27") = StartRecord(txtNhom4.Value, 10)
    
    Dim dongCuoi As Integer
    Dim PhamViResize As String
    
    dongCuoi = Sheet19.Range("L27").Value
    PhamViResize = "H29:L" & 29 + dongCuoi
    
    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom4", Sheet19, PhamViResize)
End Function

Private Sub txtNhom4_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "GH6")
    With txtNhom1
        If CInt(.Value) > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If CInt(.Value) < 1 Then
            .Value = 1
        End If
    End With
    Call ReszieNhom4
End Sub


'Nhom 5
'KD.xlsb!Sheet10.NextPage5
Private Sub NextPage5()
    With txtNhom5
        .Value = .Value + 1
    End With
End Sub

'KD.xlsb!Sheet10.PrevPage5
Private Sub PrevPage5()
    With txtNhom5
         If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1

    End With
End Sub


Public Function ReszieNhom5()
    Dim dongCuoi As Integer
    Dim PhamViResize As String
    
    Sheet19.Range("B45") = StartRecord(txtNhom5.Value, 10)
    
    dongCuoi = Sheet19.Range("F45").Value
    PhamViResize = "B47:F" & 47 + dongCuoi
    
    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom5", Sheet19, PhamViResize)
End Function

Private Sub txtNhom5_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "GR6")
    With txtNhom5
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With
    Call ReszieNhom5
End Sub


'Nhom 6
'KD.xlsb!Sheet10.NextPage6
Private Sub NextPage6()
    With txtNhom6
        .Value = .Value + 1
    End With
End Sub

'KD.xlsb!Sheet10.PrevPage6
Private Sub PrevPage6()
    With txtNhom6
        If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1
    End With
End Sub

Public Function ReszieNhom6()
    Dim dongCuoi As Integer
    Dim PhamViResize As String
    
    Sheet19.Range("H45") = StartRecord(txtNhom6.Value, 10)
    
    dongCuoi = Sheet19.Range("L45").Value
    PhamViResize = "H47:L" & 47 + dongCuoi
    
    Call UpdateChartDataRange(Sheet10, "Chart_LoiNhuan_Nhom6", Sheet19, PhamViResize)
End Function

Private Sub txtNhom6_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "HB6")
    With txtNhom5
        If .Value > TongDuLieu Then
            .Value = TongDuLieu
        End If

        If .Value < 1 Then
            .Value = 1
        End If
    End With
   Call ReszieNhom6
End Sub
