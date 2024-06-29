Attribute VB_Name = "M_DataKeHoachDoanhThu_DV"
Function NamDV() As Long
    Dim valueNam As Long

    valueNam = Workbooks("KD.xlsb").Sheets("Data KHDT DVKD").cbNamHienThiDuLieu.Value

    If valueNam <> 0 Then
        NamDV = valueNam
    Else
        NamDV = Year(Now)
    End If
End Function

Sub DoanhThuThangCha(i, rl)
    Dim arrayColum
    Dim item
    arrayColum = Array("Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ")

    With SheetDataDonViKD
        For Each item In arrayColum
            .Range(item & i) = "=SUMIFS(" & item & "$12:" & item & "$" & rl & ",A$12:A$" & rl & ",B" & i & ",F$12:F$" & rl & ",F" & i & ")"

        Next item
    End With

End Sub

Sub DoanhThuTheoThangCon(i)
    Dim key As Variant
    Dim MangCot As Object
    Set MangCot = CreateObject("Scripting.Dictionary")

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
    MangCot("AI") = "W"
    MangCot("AJ") = "X"
    
    Set wSheet = SheetDataDonViKD
    With wSheet
        For Each key In MangCot.keys
            .Range(key & i) = "=" & MangCot(key) & i & " * $G" & i

        Next key
    End With
    
    Set wSheet = Nothing
    Set MangCot = Nothing
End Sub

Sub TiLeDoanhThuTheoThang(i)
    Dim key As Variant
    Dim MangCot As Object
    Set MangCot = CreateObject("Scripting.Dictionary")

    MangCot("M") = "Y"
    MangCot("N") = "Z"
    MangCot("O") = "AA"
    MangCot("P") = "AB"
    MangCot("Q") = "AC"
    MangCot("R") = "AD"
    MangCot("S") = "AE"
    MangCot("T") = "AF"
    MangCot("U") = "AG"
    MangCot("V") = "AH"
    MangCot("W") = "AH"
    
    Set wSheet = SheetDataDonViKD
    With wSheet
        For Each key In MangCot.keys
            .Range(key & i) = "=If(G" & i & "," & MangCot(key) & i & " / G" & i & ", 0)"

        Next key
    End With
    
    Set wSheet = Nothing
    Set MangCot = Nothing
End Sub

Sub CongThucTinh(dongCuoi)
    Dim i As Integer, k As Integer, Con As Integer
    Set wSheet = SheetDataDonViKD
    With wSheet
        For i = 12 To val(dongCuoi)
            .Range("AO" & i) = "=If(COUNTIF(A" & i & ":A" & dongCuoi & ",B" & i & ") > 0,True,False)"

            .Range("L" & i) = "=K" & i & " - G" & i & ""
            .Range("X" & i) = "=100%-SUM(M" & i & ":W" & i & ")"
            .Range("J" & i) = "=I" & i & " - G" & i & ""

            If .Range("AO" & i) = "False" Then
                DoanhThuTheoThangCon i
            ElseIf .Range("AO" & i) = "True" Then

                'Ke hoach DT DVKD

                .Range("G" & i) = "=SUMIFS(G$12:G$" & dongCuoi & ",A$12:A$" & dongCuoi & ",B" & i & ",F$12:F$" & dongCuoi & ",F" & i & ")"

                'DT Thang
                DoanhThuThangCha i, dongCuoi

                'TLDT
                TiLeDoanhThuTheoThang i
            End If

        Next i

    End With
    
    Set wSheet = Nothing
End Sub

Sub ReSizeTableDV(dongCuoi)
    Set wSheet = SheetDataDonViKD
    With wSheet
        .Select
        .ListObjects("Table_Data_DV").Resize Range("$C$11:$AL$" & dongCuoi)
    End With
    
    Set wSheet = Nothing
End Sub

'Auth: NguyenHuuThanh
'Date By: 28/06/2024
'Descript: Cap nhat du lieu ke hoach don vi " Toi uu code"
Sub CapNhatDuLieuDV()
    On Error Resume Next
    BatLimit
    Dim DongCuoiCoDuLieu As Long
    Dim i As Long
    Dim PhongBanID As String
    Dim NamLapKeHoach As String
    Dim KeHoachDoanhThu As String
    Dim DoanhThuThucDat As String

    Dim PhanTram(1 To 12) As Variant
    Dim Tien(1 To 12) As Variant
    
    Dim dbConn As Object
    Dim SQLStr As String
    
    Set dbConn = ConnectToDatabase
    Set wSheet = SheetDataDonViKD
    With wSheet
        DongCuoiCoDuLieu = tinhdongcuoi("C12:C1048576")
        For i = 12 To DongCuoiCoDuLieu

            PhongBanID = .Range("B" & i)
            NamLapKeHoach = .Range("F" & i)
            KeHoachDoanhThu = .Range("G" & i)
            If KeHoachDoanhThu = "" Then KeHoachDoanhThu = 0
            DoanhThuThucDat = .Range("H" & i)

            ' Lay du lieu phan tram va tien
            For j = 1 To 12
                PhanTram(j) = .Cells(i, 12 + j)
                If PhanTram(j) = "" Then PhanTram(j) = 0
                Tien(j) = .Cells(i, 24 + j)
            Next j

            SQLStr = "DELETE FROM KeHoachDoanhThu WHERE Nam = " & NamLapKeHoach & " And PhongBanID = " & PhongBanID & ";"
            SQLStr = SQLStr & "DELETE FROM DoanhThuThucDat WHERE Nam = " & NamLapKeHoach & " And PhongBanID = " & PhongBanID & ";"
            SQLStr = SQLStr & "DELETE FROM KeHoachPhanBoDv WHERE Nam = " & NamLapKeHoach & " And PhongBanID = " & PhongBanID & ";"
            SQLStr = SQLStr & "INSERT INTO KeHoachDoanhThu(PhongBanID, Nam, KeHoachDoanhThu) VALUES(" & PhongBanID & "," & NamLapKeHoach & "," & KeHoachDoanhThu & ");"
            SQLStr = SQLStr & "INSERT INTO DoanhThuThucDat(PhongBanID, Nam, DoanhThuThucDat) VALUES(" & PhongBanID & "," & NamLapKeHoach & "," & DoanhThuThucDat & ");"
            SQLStr = SQLStr & "INSERT INTO KeHoachPhanBoDv(PhongBanID, Nam, PhanTramThang1, PhanTramThang2, PhanTramThang3, PhanTramThang4, PhanTramThang5, PhanTramThang6, PhanTramThang7, PhanTramThang8, PhanTramThang9, PhanTramThang10, PhanTramThang11, PhanTramThang12, TienThang1, TienThang2, TienThang3, TienThang4, TienThang5, TienThang6, TienThang7, TienThang8, TienThang9, TienThang10, TienThang11, TienThang12) VALUES(" & PhongBanID & "," & NamLapKeHoach & "," & PhanTram(1) & "," & PhanTram(2) & "," & PhanTram(3) & "," & PhanTram(4) & "," & PhanTram(5) & "," & PhanTram(6) & "," & PhanTram(7) & "," & PhanTram(8) & "," & PhanTram(9) & "," & PhanTram(10) & "," & PhanTram(11) & "," & PhanTram(12) & "," & Tien(1) & "," & Tien(2) & "," & Tien(3) & "," & Tien(4) & "," & Tien(5) & "," & Tien(6) & "," & Tien(7) & "," & Tien(8) & "," & Tien(9) & "," & Tien(10) & "," & Tien(11) & "," & Tien(12) & ");"

            dbConn.Execute SQLStr
        Next i

        Call CloseDatabaseConnection(dbConn)
    End With
    Set wSheet = Nothing
    TatLimit
    ThongBao_ThanhCong
End Sub

Sub hienThiDuLieuSheetDataDonViKD()
    Dim dongCuoi As Long
    Dim NguoiDangNhap As Variant
    Dim Query As String
    Dim Nam As Integer
    Dim dbConn As Object

    Nam = Sheet11.Range("C5").Value
    NguoiDangNhap = NguoiDungID

    ' Mo ket noi du lieu
    Set dbConn = ConnectToDatabase
    
    ' Xu ly lenh
    If IsEmpty(NamLapKeHoach) Then
        Query = "exec dataKHDT_DV_KD_V2 '" & Nam & "'," & NguoiDangNhap
    End If

    ' Kiem tra co du lieu khong
    With SheetDataDonViKD
        If .Range("C12").Value <> "" Then
            dongCuoi = tinhdongcuoi("C12:C1048576") + 100
            Workbooks("KD.xlsb").Sheets("Data KHDT DVKD").Range("A12:BU" & dongCuoi).Clear
        End If
    End With

    Call viewSheet(Query, SheetDataDonViKD, "A12", dbConn)

    ' Tong ke h?ch don vi
    Query = "exec KD_TK_TongHopTheo_DV " & Nam & ", " & NguoiDangNhap
    Call viewSheet(Query, SheetDataDonViKD, "J5", dbConn)

    With Sheet11
        .Select
        If .Range("G340").Value <> "" Then
            dongCuoi = tinhdongcuoi("G340:G399")
            If .Range("G340").Value <> "" And dongCuoi > 339 Then
                .Range("G340:I" & dongCuoi).Clear
                .Range("G402:I432").Clear
            End If
        End If
    End With

    Query = "EXEC KD_KeHoachDoanhThu_NamTruocNamSau " & Nam
    Call viewSheet(Query, Sheet11, "G340", dbConn)

    ' Dong ket noi du lieu
    Call CloseDatabaseConnection(dbConn)
End Sub

Sub layDuLieuDV(Nam)
    Dim dongCuoi
    Dim kq As Variant
    Set wSheet = SheetDataDonViKD
    
    With wSheet
        .Select
        Call hienThiDuLieuSheetDataDonViKD
        .Select
        dongCuoi = tinhdongcuoi("C12:C1048576")

        If .Range("D12") <> "" Then
            CongThucTinh dongCuoi
            .ListObjects("Table_Data_DV").Resize Range("$C$11:$AL$" & dongCuoi)
        End If
        .Range("A1").Select
    End With
    Set wSheet = Nothing
End Sub


'Auth: NguyenHuuThanh
'Date By: 28/06/2024
'Descript: ToMauPhanCap " Toi uu code"
Sub ToMauPhanCap(dongCuoi As Long, cap As Variant)
    Dim i As Integer
    Dim arrayColum As Variant
    Dim item As Variant
    Dim ws As Worksheet
    
    arrayColum = Array("G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ")
    
    Set ws = SheetDataDonViKD
    With ws
        For i = 12 To dongCuoi
            Call ToMau_DV(i, .Range("D" & i))

            For Each item In arrayColum
                If .Range(item & i).Value < 0 Then
                    With .Range(item & i).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next item

            If Not .Range("G" & i).HasFormula Then
                With .Range("G" & i).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark2
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With .Range("G" & i).Font
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                    .Bold = False
                End With

                With .Range("M" & i & ":W" & i).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark2
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
                With .Range("M" & i & ":W" & i).Font
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                    .Bold = False
                End With
            End If
        Next i

        ' Format J5, K5, L5
        FormatCell .Range("J5")
        FormatCell .Range("K5")
        FormatCell .Range("L5")
    End With

    Set ws = Nothing
End Sub

Sub FormatCell(rng As Range)
    With rng.Font
        If rng.Value < 0 Then
            .Color = -16776961
        Else
            .ThemeColor = xlThemeColorDark1
        End If
        .TintAndShade = 0
    End With
End Sub

'Auth: NguyenHuuThanh
'Date By: 28/06/2024
'Descript: DinhDangDVKD " Toi uu code"
Sub DinhDangDVKD(dongCuoi As Long)
    Dim DongCuoiCuaHang As Long
    Set wSheet = SheetDataDonViKD
    
    With wSheet
        ' Ð?nh d?ng du?ng vi?n "Theo công th?c c?a phòng"
        DongCuoiCuaHang = .Range("C11:AL11").End(xlToRight).Column
        
        Format_ dongCuoi, DongCuoiCuaHang, "C11", "B:B", 1, "Table_Data_DV"
        
        If dongCuoi > 0 Then
            F_BoderStyle .Range("C12:AL" & dongCuoi), "Table_Data_DV"
        End If
        
        With .Columns("B:B").Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With

        ToMauPhanCap dongCuoi, 2
        
        FormatColumns "G:L", "Y:AJ", "M:X", "C:C", "D:D", dongCuoi
        
        FixTop .Range("A12")
    End With
    
    Set wSheet = Nothing
End Sub

Sub FormatColumns(cols1 As String, cols2 As String, cols3 As String, col4 As String, col5 As String, dongCuoi As Long)
    With SheetDataDonViKD
        .Columns(cols1).Select
        FormatMoney
        .Columns(cols2).Select
        FormatMoney
        .Columns(cols3).Select
        FormatPercent
        F_TextCenter col4, ""
        F_Width col4, 5
        F_Width col5, 0
        .Range("F12:F" & dongCuoi).NumberFormat = "@"
    End With
End Sub

'Auth: NguyenHuuThanh
'Date By: 28/06/2024
'Descript: LamMoiDuLieuDV " Toi uu code"
Public Sub LamMoiDuLieuDV()
    BatLimit
    Dim Nam As Long

    With SheetDataDonViKD
        Nam = .Range("C5")
        
        If Nam = 0 Then
            Nam = Year(Now)
        End If
        
        layDuLieuDV Nam
        .Range("A1").Select
    End With

    F_StyleDV
    TatLimit
    ThongBao_ThanhCong
End Sub


Public Sub F_StyleDV()
    With SheetDataDonViKD
        .Select
        Dim dongCuoi As Long
        dongCuoi = tinhdongcuoi("C12:C1048576")
        DinhDangDVKD dongCuoi
    End With
End Sub

Public Sub VeBieuDoDonVi()
    
    On Error Resume Next
    Range("B100").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    
   BatLimit
     
    Dim PhamViResize As String
    Dim lr As Long
    With Sheet11
        .Select
        lr = .Range("I338").Value
        
        PhamViResize = "G339:I" & lr + 339
    
        Call UpdateChartDataRange(Sheet11, "Chart 26", Sheet11, PhamViResize)

        
        'Resize DB_KHDTDVKD_TB
        ThisWorkbook.Sheets("KHDT theo DVKD").Select
        lr = .Range("I100").Value
        PhamViResize = "F101:I" & lr + 101
        Call UpdateChartDataRange(Sheet11, "Chart 7", Sheet11, PhamViResize)
        ActiveWorkbook.RefreshAll
    End With
    Call ScrollToTop
   TatLimit
   ThongBao_ThanhCong
End Sub

Sub ToMau_DV(i, cap)
     Set wSheet = SheetDataDonViKD
    With wSheet
        .Select
        If cap = 2 Then
            .Range("C" & i & ":AL" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = -0.499984740745262
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
        End If

        If cap = 3 Then
            .Range("C" & i & ":AL" & i).Select
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
        End If

        If cap = 4 Then
            .Range("C" & i & ":AL" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With

        End If

        If cap = 5 Then
            .Range("C" & i & ":AL" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    End With
Set wSheet = Nothing
End Sub


