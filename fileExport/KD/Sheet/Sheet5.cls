VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub cbbChonKieuHienThi_Click()
    Sheet5.Range("E5") = cbbChonKieuHienThi.Text
    Call VeBieuDoDTDV
End Sub

Private Sub Worksheet_Activate()
    BatLimit
    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall
    Call F_KhoiTaoBoLoc
    TatLimit
End Sub

Function F_KhoiTaoBoLoc()
    If cbbChonKieuHienThi.ListCount > 0 Then
        Exit Function
    End If
    
    Dim i As Integer

    With cbbChonKieuHienThi
        .Clear
        .AddItem "N" & ChrW(259) & "m"

        For i = 1 To 12
            .AddItem "Tháng " & i
        Next i
        'Set default

        .Text = .List(0, 0)
    End With
End Function

Function F_LaySoTuChuoi(chuoi As String) As Integer
    If chuoi = "" Then
     Exit Function
    End If

    Dim i As Integer
    Dim kyTu As String

    For i = 1 To Len(chuoi)
        kyTu = Mid(chuoi, i, 1)
        If IsNumeric(kyTu) Then
         Exit For
        End If
    Next i

    F_LaySoTuChuoi = Mid(chuoi, i)

End Function

Function F_FindLastColName() As String
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim cell As Range
    Dim LastCol As Long
    Dim lastColName As String


    Set ws = ThisWorkbook.Sheets("Pivot SP")

    Set searchRange = ws.Range("HZ3:IV3")

    For Each cell In searchRange

        If cell.Value = 1 Then

            LastCol = cell.Column

            lastColName = Split(Cells(1, LastCol).Address, "$")(1)
        End If
    Next cell

    If LastCol > 0 Then
        F_FindLastColName = lastColName
    End If

End Function

Sub VeBieuDoDTDV()
    BatLimit
    Call F_R_DATA

    If cbbChonKieuHienThi.ColumnCount = 0 Then
        F_KhoiTaoBoLoc
    End If

    If cbbChonKieuHienThi.Text = "N" & ChrW(259) & "m" Then
        Call F_taiDuLieu(0)
    Else
        Call F_taiDuLieu(F_LaySoTuChuoi(cbbChonKieuHienThi.Text))
    End If
    
    ActiveWorkbook.RefreshAll
    Sheet5.Select
    TatLimit
    ThongBao_ThanhCong
End Sub

Function F_taiDuLieu(Value As Integer)
    BatLimit
    Dim dbConn As Object
    Dim dongCuoi As LongLong
    Dim Query As String
    Set dbConn = ConnectToDatabase

    Query = "exec KD_BAO_CAO_KINH_DOANH_THEO_NAM " & Value

    With Sheet26
        .Select
        .Range("HZ5:IV49").Clear

        Call viewSheetHeader(Query, Sheet26, "HZ5", "HZ4", dbConn)

        Call CloseDatabaseConnection(dbConn)

        dongCuoi = tinhdongcuoi("HZ5:HZ49")
        ActiveSheet.ListObjects("Table42").Resize Range("$HZ$4:$JA$" & dongCuoi)
    End With
    TatLimit
End Function

'KD.xlsb!Sheet5.NextPage1
Private Sub NextPage1()
    With txtNhom1
        .Value = .Value + 1
    End With
End Sub

Private Sub PrevPage1()
    With txtNhom1
        If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1

    End With
End Sub

Public Function ResizeNhom1()
    Sheet22.Range("B9") = StartRecord(Sheet5.txtNhom1.Value, 11)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet22.Range("H9").Value
    PhamViResize = "B11:H" & 11 + dongCuoi

    Call UpdateChartDataRange(Sheet5, "ChartBaocaoDoanhThuCacDonViKD", Sheet22, PhamViResize)
End Function

Private Sub txtNhom1_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "HB6")
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
         If .Value = 1 Then
            Exit Sub
        End If
        .Value = .Value - 1

    End With
End Sub

Public Function ResizeNhom2()
    Sheet22.Range("J9") = StartRecord(Sheet5.txtNhom2.Value, 11)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet22.Range("O9").Value
    PhamViResize = "J11:O" & 11 + dongCuoi

    Call UpdateChartDataRange(Sheet5, "ChartBaocaoDoanhThuCacDonViKD2", Sheet22, PhamViResize)
End Function

Private Sub txtNhom2_Change()
    Dim TongDuLieu As Integer
    TongDuLieu = totalPage(Sheet26, "II49")
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
