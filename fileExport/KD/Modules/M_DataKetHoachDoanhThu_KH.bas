Attribute VB_Name = "M_DataKetHoachDoanhThu_KH"

Function NamKh() As Long
    With Sheet11
        .Select
        NamKh = .Range("C5")
    End With
    
End Function

Sub DTTKH(i)
    Dim key As Variant
    Dim MangCot As Object
    Set MangCot = CreateObject("Scripting.Dictionary")

    MangCot("W") = "K"
    MangCot("X") = "L"
    MangCot("Y") = "M"
    MangCot("Z") = "N"
    MangCot("AA") = "O"
    MangCot("AB") = "P"
    MangCot("AC") = "Q"
    MangCot("AD") = "R"
    MangCot("AE") = "S"
    MangCot("AF") = "T"
    MangCot("AG") = "U"
    MangCot("AH") = "V"

    Set wSheet = SheetDataKhachHangKD
    With wSheet
        For Each key In MangCot.keys
            .Range(key & i) = "=IFERROR(" & MangCot(key) & i & "*$I" & i & ",0)"

        Next key
    End With
    Set wSheet = Nothing
    Set MangCot = Nothing
End Sub

Sub CongThucTinhKH()
    Dim i As Integer
    Dim dongCuoi
    Set wSheet = SheetDataKhachHangKD
    With wSheet
        dongCuoi = tinhdongcuoi("C12:C1048576")
    End With

    For i = 12 To dongCuoi
        DTTKH i

    Next i

    Set wSheet = Nothing
End Sub

Sub CapNhatDuLieuKH()
    On Error Resume Next
    BatLimit
    Dim DongCuoiCoDuLieu As Long
    Dim i As Integer
    Dim PhongBanID As String
    Dim NhanVienIDKH As String
    Dim NamLapKeHoach As String
    Dim KhachHangID As String
    Dim KeHoachDoanhThu As String
    Dim NguoiDungID As Long

    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim SQL_DELETE_KHDT As String
    Dim SQL_DELETE_KHPB As String
    Dim SQL_INSERT_KHDT As String
    Dim SQL_INSERT_KHPB As String

    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    Dim PhanTramT1
    Dim PhanTramT2
    Dim PhanTramT3
    Dim PhanTramT4
    Dim PhanTramT5
    Dim PhanTramT6
    Dim PhanTramT7
    Dim PhanTramT8
    Dim PhanTramT9
    Dim PhanTramT10
    Dim PhanTramT11
    Dim PhanTramT12
    Dim TienT1
    Dim TienT2
    Dim TienT3
    Dim TienT4
    Dim TienT5
    Dim TienT6
    Dim TienT7
    Dim TienT8
    Dim TienT9
    Dim TienT10
    Dim TienT11
    Dim TienT12

    StrCnn = KetNoiMayChu_KhachHang
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Set wSheet = SheetDataKhachHangKD
    With SheetDataKhachHangKD
        DongCuoiCoDuLieu = tinhdongcuoi("C12:C1048576")

        For i = 12 To DongCuoiCoDuLieu

            KhachHangID = .Range("A" & i)
            NhanVienIDKH = .Range("AL" & i)
            PhongBanID = .Range("AK" & i)
            NamLapKeHoach = .Range("H" & i)

            If .Range("I" & i) = "" Then

                KeHoachDoanhThu = 0

            Else
                KeHoachDoanhThu = .Range("I" & i)
            End If

            PhanTramT1 = .Range("K" & i)
            PhanTramT2 = .Range("L" & i)
            PhanTramT3 = .Range("M" & i)
            PhanTramT4 = .Range("N" & i)
            PhanTramT5 = .Range("O" & i)
            PhanTramT6 = .Range("P" & i)
            PhanTramT7 = .Range("Q" & i)
            PhanTramT8 = .Range("R" & i)
            PhanTramT9 = .Range("S" & i)
            PhanTramT10 = .Range("T" & i)
            PhanTramT11 = .Range("U" & i)
            PhanTramT12 = .Range("V" & i)

            TienT1 = .Range("W" & i)
            TienT2 = .Range("X" & i)
            TienT3 = .Range("Y" & i)
            TienT4 = .Range("Z" & i)
            TienT5 = .Range("AA" & i)
            TienT6 = .Range("AB" & i)
            TienT7 = .Range("AC" & i)
            TienT8 = .Range("AD" & i)
            TienT9 = .Range("AE" & i)
            TienT10 = .Range("AF" & i)
            TienT11 = .Range("AG" & i)
            TienT12 = .Range("AH" & i)

            If PhanTramT1 = "" Then
                PhanTramT1 = 0
            End If

            If PhanTramT2 = "" Then
                PhanTramT2 = 0
            End If
            If PhanTramT3 = "" Then
                PhanTramT3 = 0
            End If
            If PhanTramT4 = "" Then
                PhanTramT4 = 0
            End If
            If PhanTramT5 = "" Then
                PhanTramT5 = 0
            End If
            If PhanTramT6 = "" Then
                PhanTramT6 = 0
            End If
            If PhanTramT7 = "" Then
                PhanTramT7 = 0
            End If
            If PhanTramT8 = "" Then
                PhanTramT8 = 0
            End If
            If PhanTramT9 = "" Then
                PhanTramT9 = 0
            End If
            If PhanTramT10 = "" Then
                PhanTramT10 = 0
            End If
            If PhanTramT11 = "" Then
                PhanTramT11 = 0
            End If
            If PhanTramT12 = "" Then
                PhanTramT12 = 0
            End If

            If KhachHangID <> "" And KeHoachDoanhThu <> "" And PhongBanID <> "" Then
                If PhongBanID <> "" Then
                    SQL_DELETE_KHDT = "DELETE FROM KeHoachDoanhThuKh WHERE Nam = " & NamLapKeHoach & " And PhongBanID = " & PhongBanID & " And NhanVienID = " & NhanVienIDKH & " And KhachHangID = " & KhachHangID
                    SQL_DELETE_KHPB = "DELETE FROM KeHoachPhanBoKh WHERE Nam = " & NamLapKeHoach & " And PhongBanID = " & PhongBanID & " And NhanVienID = " & NhanVienIDKH & " And KhachHangID = " & KhachHangID
                    SQL_INSERT_KHDT = "INSERT INTO KeHoachDoanhThuKh(PhongBanID,NhanVienID,KhachHangID, Nam, KeHoachDoanhThuKh) VALUES(" & PhongBanID & "," & NhanVienIDKH & "," & KhachHangID & "," & NamLapKeHoach & "," & KeHoachDoanhThu & ")"
                    SQL_INSERT_KHPB = "INSERT INTO KeHoachPhanBoKh(PhongBanID,NhanVienID,KhachHangID,Nam,PhanTramThang1, PhanTramThang2,PhanTramThang3, PhanTramThang4, PhanTramThang5, PhanTramThang6, PhanTramThang7, PhanTramThang8, PhanTramThang9, PhanTramThang10, PhanTramThang11, PhanTramThang12, TienThang1, TienThang2,TienThang3, TienThang4,TienThang5, TienThang6, TienThang7, TienThang8, TienThang9, TienThang10, TienThang11, TienThang12) VALUES(" & PhongBanID & "," & NhanVienIDKH & "," & KhachHangID & "," & NamLapKeHoach & "," & PhanTramT1 & "," & PhanTramT2 & "," & PhanTramT3 & "," & PhanTramT4 & "," & PhanTramT5 & "," & PhanTramT6 & "," & PhanTramT7 & "," & PhanTramT8 & "," & PhanTramT9 & "," & PhanTramT10 & "," & PhanTramT11 & "," & PhanTramT12 & "," & TienT1 & "," & TienT2 & "," & TienT3 & "," & TienT4 & "," & TienT5 & "," & TienT6 & "," & TienT7 & "," & TienT8 & "," & TienT9 & "," & TienT10 & "," & TienT11 & "," & TienT12 & ") "

                    SQLStr = SQL_DELETE_KHDT & ";" & SQL_DELETE_KHPB & ";" & SQL_INSERT_KHDT & ";" & SQL_INSERT_KHPB
                    Rs.Open SQLStr, Cn, adOpenStatic

                ElseIf PhongBanID = "" Then

                    If NguoiDungID <> 0 Then
                        NhanVienIDKH = "(Select NhanVienID FROM NS_NhanVien WHERE NhanVienIDKH  in(Select NhanVienIDKH FROM PQ_NguoiDung WHERE NguoiDungID = " & NguoiDungID & "))"
                        PhongBanID = "(Select PhongBanID FROM NS_NhanVien WHERE NhanVienID  in(Select NhanVienID FROM PQ_NguoiDung WHERE NguoiDungID = " & NguoiDungID & "))"

                        SQLStr = "INSERT INTO KeHoachDoanhThuKh(PhongBanID,NhanVienID,KhachHangID, Nam, KeHoachDoanhThuKh) VALUES(" & PhongBanID & "," & NhanVienIDKH & "," & KhachHangID & "," & NamLapKeHoach & "," & KeHoachDoanhThu & ")"

                        Rs.Open SQLStr, Cn, adOpenStatic
                    End If
                End If
            End If

        Next i

        Cn.Close

        Set Rs = Nothing
        Set Cn = Nothing

    End With
    Set wSheet = Nothing
    TatLimit
    ThongBao_ThanhCong
End Sub

Sub ReSizeTableKH()
    Dim dongCuoi As Long
    Set wSheet = SheetDataKhachHangKD
    With SheetDataKhachHangKD
        .Select
        .Range("AU1") = "=COUNTA(C12:C1048576)"

        dongCuoi = tinhdongcuoi("C12:C1048576")
    End With
    Set wSheet = Nothing
    Sheets("Data KHDT KH").ListObjects("TableKhachHang").Resize Range("$C$11:$AN$" & dongCuoi)
    '    ActiveSheet.Shapes.Range(Array("Slicer_KhachHang")).Select
    '    ActiveWorkbook.SlicerCaches("Slicer_Nam_KH").SortItems = xlSlicerSortDescending
End Sub

Sub layDuLieuKH(Nam, NguoiDungID, NhanVienIDKH)
    Dim kq As Variant
    Dim dongCuoi
    Set wSheet = SheetDataKhachHangKD
    With wSheet
        .Select
        Call hienThiDuLieuKH(Nam, NguoiDungID, NhanVienIDKH)

        If .Range("C12") <> "" Then

            ReSizeTableKH
            CongThucTinhKH
        End If

        .Range("A1").Select
    End With
    Set wSheet = Nothing
End Sub

Sub LoadDuLieu_KH()
    If NguoiDungID = 0 Then
     Exit Sub
    End If

    If NamKh <> 0 Then

        layDuLieuKH NamKh, NguoiDungID, 0
    Else
        layDuLieuKH Year(Now), NguoiDungID, 0
    End If
End Sub

Sub LamMoiDuLieuKH()
    BatLimit
    Set wSheet = SheetDataKhachHangKD
    Set wSheet = Nothing
    LoadDuLieu_KH
    F_Style_KH
    TatLimit
    ThongBao_ThanhCong
End Sub

Sub hienThiDuLieuKH(NamKh, NguoiDungID, NhanVienIDKH)
    Dim dongCuoi
    Dim Query As String
    Dim DangNhapID

    DangNhapID = NguoiDungID
    Set wSheet = SheetDataKhachHangKD
    With wSheet
        .Select

        If .Range("C12") <> "" Then

            dongCuoi = tinhdongcuoi("C12:C1048576")
            .Range("A12:AN" & dongCuoi).Clear
        End If

        '        Mo ket noi csdl
        Dim dbConn As Object
        Set dbConn = ConnectToDatabase
        Query = "exec dataKHDT_KH_KD_V2 '" & NamKh & "'," & DangNhapID & "," & NhanVienIDKH

        Call viewSheet(Query, SheetDataKhachHangKD, "A12", dbConn)

        'Tong Hop theo KH
        Query = "exec KD_TK_TongHopTheo_KH " & NamKh & ", " & DangNhapID & ", 0"
        Call viewSheet(Query, SheetDataKhachHangKD, "I5", dbConn)

        'Dong Ket noi
        Call CloseDatabaseConnection(dbConn)
    End With
    Set wSheet = Nothing
End Sub

Sub ToMauPhanCapKH()
    Dim i As Integer
    Dim arrayColum As Variant
    Dim item
    Dim dongCuoi As Long
    Set wSheet = SheetDataKhachHangKD
    With wSheet

        dongCuoi = tinhdongcuoi("C12:C1048576")

        For i = 12 To dongCuoi
            Call ToMauKH(i, .Range("A" & i), SheetDataKhachHangKD)
        Next i

    End With

    Set wSheet = Nothing
End Sub

Sub F_Style_KH()

    Dim kq  As String
    Dim dongCuoi As Long
    Set wSheet = SheetDataKhachHangKD
    With wSheet
        .Select
        dongCuoi = tinhdongcuoi("C12:C1048576")

        Columns("B:B").Select

        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

        Format_ dongCuoi, 40, "A11", "B:B", 2, "TableKhachHang"

        ToMauPhanCapKH

        F_BoderStyle .Range("C12:AN" & dongCuoi), "TableKhachHang"

        Columns("B:B").Select

        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With

        F_Width "C:C", 5
        F_Width "AK:AL", 0

        Columns("I:J").Select
        FormatMoney
        Columns("W:AH").Select
        FormatMoney
        Columns("K:V").Select
        FormatPercent
        Range("H12:H" & dongCuoi + 11).Select
        Selection.NumberFormat = "@"
        .Range("K5").Select
        FormatMoney

        If .Range("I5") < 0 Then
            .Range("I5").Select
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        Else
            .Range("I5").Select
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If

        If .Range("J5") < 0 Then
            .Range("J5").Select
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        Else
            .Range("J5").Select
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If

        If .Range("K5") < 0 Then
            .Range("K5").Select
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        Else
            .Range("K5").Select
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If
        'Co dinh tieu de
        FixTop .Range("A12")
    End With

    Set wSheet = Nothing
End Sub

Sub Load_All_KH()
    BatLimit
    LoadDuLieu_KH
    With Worksheets("KHDT theo KH")
        .Select
        .PivotTables("PivotTable5").PivotCache.Refresh
        a.PivotTables("PivotTable6").PivotCache.Refresh
        Range("R107").Select
    End With
    With ActiveSheet.PivotTables("PivotTable6").PivotFields("Mã NV")
        .PivotItems("(blank)").Visible = False
    End With


    TatLimit
    ThongBao_ThanhCong
End Sub

Sub ToMauKH(i, cap, wSheet As Worksheet)
    With wSheet
        .Select
        If cap = 0 Then
            .Range("C" & i & ":AN" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = -0.249977111117893
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
        Else
            .Range("C" & i & ":AN" & i).Select
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Selection.Font.Bold = False

            .Range("i" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If

    End With

End Sub


