Attribute VB_Name = "M_DataKeHoachDoanhThu_NV"
Function NamNV() As Long
    Dim valueNam As Long

    valueNam = Sheet11.Range("C5")

    If valueNam <> 0 Then
        NamNV = valueNam
    Else
        NamNV = Year(Now)
    End If

End Function

'Auth: NguyenHuuThanh
'Date By: 28/06/2024
'Descript: DTT " Toi uu code"
Sub DTT(i)
    Dim key As Variant
    Dim MangCot As Object
    Set MangCot = CreateObject("Scripting.Dictionary")

    ' Tao tu dien vo cap khoa
    With MangCot
        .Add "Z", "N"
        .Add "AA", "O"
        .Add "AB", "P"
        .Add "AC", "Q"
        .Add "AD", "R"
        .Add "AE", "S"
        .Add "AF", "T"
        .Add "AG", "U"
        .Add "AH", "V"
        .Add "AI", "W"
        .Add "AJ", "X"
        .Add "AK", "Y"
    End With

    With SheetDataNhanVienKD
        For Each key In MangCot.keys
            .Range(key & i).Formula = "=IFERROR(" & MangCot(key) & i & "*$J" & i & ",0)"
        Next key
    End With

    ' Giai phong doi tuong
    Set MangCot = Nothing
End Sub

Sub CongThucTinhNv()
    Dim i As Integer
    Dim dongCuoi As Long
    Set wSheet = SheetDataNhanVienKD
    With SheetDataNhanVienKD
        .Select
        dongCuoi = tinhdongcuoi("C12:C1048576")
        For i = 12 To dongCuoi
            ' Thua thieu KH Theo KH
            .Range("M" & i) = "=L" & i & " - J" & i & ""
            
            .Range("Y" & i) = "=100%-SUM(N" & i & ":X" & i & ")"
            DTT i

        Next i
    End With
    Set wSheet = Nothing
End Sub


'Auth: NguyenHuuThanh
'Date By: 28/06/2024
'Descript: ToMauPhanCapNV " Toi uu code"
Sub ToMauPhanCapNV(cap, Optional parent As Variant = 1)
    Dim i As Long
    Dim dongCuoi As Long
    Dim arrayColum As Variant
    Dim item As Variant

    dongCuoi = tinhdongcuoi("C12:C1048576")

    arrayColum = Array("J", "K", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM")
    
    With SheetDataNhanVienKD
        For i = 12 To dongCuoi
            Call ToMauNV(i, .Range("E" & i), SheetDataNhanVienKD)

            For Each item In arrayColum
                If .Range(item & i).Value < 0 Then
                    With .Range(item & i).Font
                        .Color = -16776961
                        .TintAndShade = 0
                    End With
                End If
            Next item
        Next i

        ' Dinh dang cac o J5, K5, L5
        Dim cellsToCheck As Variant
        cellsToCheck = Array("J5", "K5", "L5")

        For Each cell In cellsToCheck
            If .Range(cell).Value < 0 Then
                With .Range(cell).Font
                    .Color = -16776961
                    .TintAndShade = 0
                End With
            Else
                With .Range(cell).Font
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
            End If
        Next cell
    End With
End Sub

Sub ReSizeTableNV()
    Dim dongCuoi As Long
    SheetDataNhanVienKD.Select
    dongCuoi = tinhdongcuoi("I12:I1048576")
    If dongCuoi <= 11 Then
        SheetDataNhanVienKD.ListObjects("TableNhanVienKD").Resize Range("D11:AM12")
    Else
        SheetDataNhanVienKD.ListObjects("TableNhanVienKD").Resize Range("D11:AM" & dongCuoi)
    End If

    'Resite DB_KHDTNVKD_TB
    
End Sub

Sub F_Style_NV()

    Dim kq As Variant
    Dim dongCuoi

    dongCuoi = tinhdongcuoi("C12:C1048576")
    Set wSheet = SheetDataNhanVienKD
    With SheetDataNhanVienKD
        .Select

        ' Dinh Dang tien
        .Columns("J:O").Select
        FormatMoney

        .Columns("Z:AK").Select
        FormatMoney

        ' Dinh dang Phan %
        .Columns("N:Y").Select
        FormatPercent

        'To mau chu phan Cot B
        Columns("C:C").Select
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With

        Format_ dongCuoi, 39, "D11", "C:C", 1, "TableNhanVienKD"

        ToMauPhanCapNV 2, 1

        F_BoderStyle .Range("D11:AK" & dongCuoi), "TableNhanVienKD"

        F_Width "E:E", 0
        F_Width "AL:AL", 0
        F_Width "AM:AM", 0
        F_Width "D:D", 5
        F_TextCenter "D:D", ""

        Range("I12:I" & dongCuoi).Select
        Selection.NumberFormat = "@"
        Columns("AN").Select
        With Selection.Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With

        FixTop Range("A12")

    End With
    Set wSheet = Nothing
End Sub


'Auth: NguyenHuuThanh
'Date By: 28/06/2024
'Descript: CapNhatDuLieuNV " Toi uu code"
Sub CapNhatDuLieuNV()
    On Error GoTo ErrorHandler
    BatLimit

    Dim i As Long
    Dim dongCuoi As Long
    Dim PhongBanID As String
    Dim NhanVienID As String
    Dim NamLapKeHoach As String
    Dim KeHoachDoanhThu As String
    Dim dbConn As Object
    Dim Rs As Object

    Dim SQLStr As String
    Dim SQL_DELETE_KHDT As String
    Dim SQL_DELETE_KHPB As String
    Dim SQL_INSERT_KHDT As String
    Dim SQL_INSERT_KHPB As String

    Dim PhanTram(1 To 12) As Variant
    Dim Tien(1 To 12) As Variant

    dongCuoi = tinhdongcuoi("C12:C1048576")

    Set dbConn = ConnectToDatabase
    Set wSheet = SheetDataNhanVienKD

    With wSheet
        For i = 12 To dongCuoi
            PhongBanID = .Range("B" & i).Value
            NhanVienID = .Range("C" & i).Value
            NamLapKeHoach = .Range("I" & i).Value
            KeHoachDoanhThu = .Range("J" & i).Value
            If KeHoachDoanhThu = "" Then KeHoachDoanhThu = 0

            ' L?y giá tr? PhanTram
            For j = 1 To 12
                PhanTram(j) = .Cells(i, j + 13).Value
                If PhanTram(j) = "" Then PhanTram(j) = 0
            Next j

            ' L?y giá tr? Tien
            For j = 1 To 12
                Tien(j) = .Cells(i, j + 25).Value
                If Tien(j) = "" Then Tien(j) = 0
            Next j

            If NhanVienID <> "" Then
                SQL_DELETE_KHDT = "DELETE FROM KeHoachDoanhThuNv WHERE Nam = " & NamLapKeHoach & " And PhongBanID = " & PhongBanID & " And NhanVienID = " & NhanVienID
                SQL_DELETE_KHPB = "DELETE FROM KeHoachPhanBoNv WHERE Nam = " & NamLapKeHoach & " And PhongBanID = " & PhongBanID & " And NhanVienID = " & NhanVienID
                SQL_INSERT_KHDT = "INSERT INTO KeHoachDoanhThuNv(PhongBanID, NhanVienID, Nam, KeHoachDoanhThuNv) VALUES(" & PhongBanID & "," & NhanVienID & "," & NamLapKeHoach & "," & KeHoachDoanhThu & ")"
                SQL_INSERT_KHPB = "INSERT INTO KeHoachPhanBoNv(PhongBanID, NhanVienID, Nam, PhanTramThang1, PhanTramThang2, PhanTramThang3, PhanTramThang4, PhanTramThang5, PhanTramThang6, PhanTramThang7, PhanTramThang8, PhanTramThang9, PhanTramThang10, PhanTramThang11, PhanTramThang12, TienThang1, TienThang2, TienThang3, TienThang4, TienThang5, TienThang6, TienThang7, TienThang8, TienThang9, TienThang10, TienThang11, TienThang12) VALUES(" & PhongBanID & "," & NhanVienID & "," & NamLapKeHoach & "," & PhanTram(1) & "," & PhanTram(2) & "," & PhanTram(3) & "," & PhanTram(4) & "," & PhanTram(5) & "," & PhanTram(6) & "," & PhanTram(7) & "," & PhanTram(8) & "," & PhanTram(9) & "," & PhanTram(10) & "," & PhanTram(11) & "," & PhanTram(12) & "," & Tien(1) & "," & Tien(2) & "," & Tien(3) & "," & Tien(4) & "," & Tien(5) & "," & Tien(6) & "," & Tien(7) & "," & Tien(8) & "," & Tien(9) & "," & Tien(10) & "," & Tien(11) & "," & Tien(12) & ")"
                SQLStr = SQL_DELETE_KHDT & ";" & SQL_DELETE_KHPB & ";" & SQL_INSERT_KHDT & ";" & SQL_INSERT_KHPB
                dbConn.Execute SQLStr
            End If
        Next i

        Call CloseDatabaseConnection(dbConn)
    End With

    Set wSheet = Nothing
    TatLimit
    ThongBao_ThanhCong
    Exit Sub

ErrorHandler:
    MsgBox "X" & ChrW(7843) & "y ra l" & ChrW(7895) & "i trong qu" & ChrW(225) & " tr" & ChrW(236) & "nh c" & ChrW(7853) & "p nh" & ChrW(7853) & "t th" & ChrW(7917) & " l" & ChrW(7841) & "i sau"
    On Error GoTo 0
End Sub


Function hienThiSheetDataNhanVienKD(NamNV)
    Dim Query As String

    Dim dongCuoi As Integer
    Dim NguoiDangNhap

    NguoiDangNhap = NguoiDungID
    Set wSheet = SheetDataNhanVienKD
    With SheetDataNhanVienKD
        .Select
        If .Range("C12") <> "" Then
            dongCuoi = tinhdongcuoi("C12:C1048576")
            Workbooks("KD.xlsb").Sheets("Data KHDT NVKD").Range("A12:AO" & dongCuoi).Clear
        End If
        
'        Mo Ket noi csdl
        Dim dbConn As Object
        Set dbConn = ConnectToDatabase
        
        Query = "exec dataKHDT_NV_KD_V2 " & stk & NamNV & stk & "," & NguoiDangNhap & ",0 "

        Call viewSheet(Query, SheetDataNhanVienKD, "A12", dbConn)
        
      
        Query = "exec KD_TK_TongHopTheo_NV " & NamNV & ", " & NguoiDangNhap & ",0"
        Call viewSheet(Query, SheetDataNhanVienKD, "J5", dbConn)
        
'        Dong Ket noi csdl
        Call CloseDatabaseConnection(dbConn)
    End With
     Set wSheet = Nothing
End Function

Sub layDuLieuNV(NamNV)
    Dim kq As Variant
 Set wSheet = SheetDataNhanVienKD
    With SheetDataNhanVienKD
        .Select
        Call hienThiSheetDataNhanVienKD(NamNV)
        .Select

        If .Range("D12") <> "" Then

            ReSizeTableNV
            CongThucTinhNv
        End If
        .Select
        .Range("A1").Select
    End With
     Set wSheet = Nothing
End Sub

Public Sub LamMoiDuLieuNV()
    BatLimit
    Dim Nam As Long
    With Workbooks("KD.xlsb").Sheets("Data KHDT NVKD")
        
        Nam = Sheet11.Range("C5")

        If Nam <> 0 Then
            layDuLieuNV Nam
        Else
            layDuLieuNV Year(Now)
        End If

    End With

    Range("A1").Select
    F_Style_NV
    TatLimit

    ThongBao_ThanhCong
End Sub

Sub VeBieuDoNhanVien()
    BatLimit
    With Worksheets("KHDT theo NVKD")
        .Select
        .Range("B101").Select
        .PivotTables("PivotTable2").PivotCache.Refresh
        ReSizeTableNV
       
        .ChartObjects("Chart 11").Activate
        ActiveChart.PlotArea.Select
        Application.CutCopyMode = False
        ActiveChart.SetSourceData Source:=Range("DB_KHDTNVKD_TB[#All]")

    End With
    TatLimit
    ThongBao_ThanhCong
End Sub

Sub ToMauNV(i, cap, wSheet As Worksheet)
    With wSheet
        .Select
        If cap = 2 Then
            .Range("D" & i & ":AO" & i).Select
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
            .Range("D" & i & ":AO" & i).Select
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
            .Range("D" & i & ":AO" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
            End With

        End If

        If cap = 5 Then
            .Range("D" & i & ":AO" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent5
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
        End If

        If .Range("G" & i) <> "" Then
            .Range("D" & i & ":i" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
            .Range("K" & i & ":AO" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
            .Range("J" & i).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Selection.Font.Bold = False
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
        End If
    End With

End Sub

'Anh Tuan them cac ham nay de resize cho cac do thi loi nhuan


Sub ReSizeTableLN1()
    Dim dongCuoi As Long
    Sheet26.Select
    dongCuoi = tinhdongcuoi("FB7:FB10000")
    If dongCuoi <= 1 Then
        Sheet26.ListObjects("Table_LNTSP_1").Resize Range("FB7:FF8")
    Else
        Sheet26.ListObjects("Table_LNTSP_1").Resize Range("FB7:FF" & dongCuoi)
    End If
    
End Sub
Sub ReSizeTableLN2()
    Dim dongCuoi As Long
    Sheet26.Select
    dongCuoi = tinhdongcuoi("FL7:FL10000")
    If dongCuoi <= 1 Then
        Sheet26.ListObjects("Table_LNTSP_2").Resize Range("FL7:FP8")
    Else
        Sheet26.ListObjects("Table_LNTSP_2").Resize Range("FL7:FP" & dongCuoi)
    End If
    
End Sub

Sub ReSizeTableLN3()
    Dim dongCuoi As Long
    Sheet26.Select
    dongCuoi = tinhdongcuoi("FV7:FV10000")
    If dongCuoi <= 1 Then
        Sheet26.ListObjects("Table_LNTSP_3").Resize Range("FV7:FZ8")
    Else
        Sheet26.ListObjects("Table_LNTSP_3").Resize Range("FV7:FZ" & dongCuoi)
    End If
    
End Sub

Sub ReSizeTableLN4()
    Dim dongCuoi As Long
    Sheet26.Select
    dongCuoi = tinhdongcuoi("GF7:GF10000")
    If dongCuoi <= 1 Then
        Sheet26.ListObjects("Table_LNTSP_4").Resize Range("GF7:GJ8")
    Else
        Sheet26.ListObjects("Table_LNTSP_4").Resize Range("GF7:GJ" & dongCuoi)
    End If
    
End Sub

Sub ReSizeTableLN5()
    Dim dongCuoi As Long
    Sheet26.Select
    dongCuoi = tinhdongcuoi("GP7:GP10000")
    If dongCuoi <= 1 Then
        Sheet26.ListObjects("Table_LNTSP_5").Resize Range("GP7:GT8")
    Else
        Sheet26.ListObjects("Table_LNTSP_5").Resize Range("GP7:GT" & dongCuoi)
    End If
    
End Sub

Sub ReSizeTableLN6()
    Dim dongCuoi As Long
    Sheet26.Select
    dongCuoi = tinhdongcuoi("GZ7:GZ10000")
    If dongCuoi <= 1 Then
        Sheet26.ListObjects("Table_LNTSP_6").Resize Range("GZ7:HD8")
    Else
        Sheet26.ListObjects("Table_LNTSP_6").Resize Range("GZ7:HD" & dongCuoi)
    End If
    
End Sub

