
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
    Call KhoiTaoCbbBoxPageBcDtCacDonViKD
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
    Dim lastCol As Long
    Dim lastColName As String


    Set ws = ThisWorkbook.Sheets("Pivot SP")

    Set searchRange = ws.Range("HZ3:IV3")

    For Each cell In searchRange

        If cell.value = 1 Then

            lastCol = cell.Column

            lastColName = Split(Cells(1, lastCol).Address, "$")(1)
        End If
    Next cell

    If lastCol > 0 Then
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

    Dim dongCuoi As Integer
    With Sheet26
        .Select
        dongCuoi = tinhdongcuoi("HZ5:HZ49")
        .ListObjects("Table40").Resize Range("$IX$4:$JD$" & dongCuoi)

        .ListObjects("Table42").Resize Range("$HZ$4:$IE$" & dongCuoi)
        Dim i As Integer
        Dim formulaIY As String
        Dim formulaIZ As String
        For i = 5 To dongCuoi

            formulaIY = "=IFERROR(VLOOKUP($IX" & i & ",$HZ$4:$IV$30,MATCH(IY$3,$HZ$4:$IV$4,0),0),"""")"

            .Range("IY" & i).Formula = formulaIY

            formulaIZ = "=IFERROR(VLOOKUP($IX" & i & ",BaoCao_KeHoachLuyKe!$D$12:$J$200,3,0),0)"

            .Range("IZ" & i).Formula = formulaIZ
        Next i

        Dim a As String
        a = "$HZ$4:$" & F_FindLastColName & "$" & dongCuoi


        .ListObjects("Table42").Resize Range(a)

    End With

    Sheet5.Select

    ActiveWorkbook.RefreshAll
    TatLimit
    ThongBao_ThanhCong
End Sub

Function F_taiDuLieu(value As Integer)
    BatLimit
    Dim dbConn As Object
    Dim dongCuoi As LongLong
    Dim Query As String
    Set dbConn = ConnectToDatabase

    Query = "exec KD_BAO_CAO_KINH_DOANH_THEO_NAM " & value

    With Sheet26
        .Select
        .Range("HZ5:IV49").Clear

        Call viewSheetHeader(Query, Sheet26, "HZ5", "HZ4", dbConn)

        Call CloseDatabaseConnection(dbConn)

        dongCuoi = tinhdongcuoi("HZ5:HZ49")
        ActiveSheet.ListObjects("Table42").Resize Range("$HZ$4:$IV$" & dongCuoi)
    End With
    TatLimit
End Function

Private Sub cbbBaoCaoDtCacDvKD_Click()
    Sheet22.Range("B9") = StartRecord(Sheet5.cbbBaoCaoDtCacDvKD.value, 11)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet22.Range("H9").value
    PhamViResize = "B11:H" & 11 + dongCuoi

    Call UpdateChartDataRange(Sheet5, "ChartBaocaoDoanhThuCacDonViKD", Sheet22, PhamViResize)
End Sub

Private Sub cbbBaoCaoDtCacDvKD2_Click()
    Sheet22.Range("B9") = StartRecord(Sheet5.cbbBaoCaoDtCacDvKD2.value, 11)

    Dim dongCuoi As Integer
    Dim PhamViResize As String

    dongCuoi = Sheet22.Range("O9").value
    PhamViResize = "J11:O" & 11 + dongCuoi

    Call UpdateChartDataRange(Sheet5, "ChartBaocaoDoanhThuCacDonViKD2", Sheet22, PhamViResize)
End Sub
