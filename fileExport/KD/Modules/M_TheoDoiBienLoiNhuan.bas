Attribute VB_Name = "M_TheoDoiBienLoiNhuan"
Option Explicit

Sub HienThiBienLoiNhuan()
    BatLimit
    On Error Resume Next
    Dim Nam As Integer
    Dim dongCuoi
    Dim Query As String

    With SheetTheoDoiBienLoiNhuan
        .Select
        If .cbNam.Value <> "" Then
            Nam = .cbNam.Value
        Else
            Nam = Year(Now)
        End If


        ' Check Có du lieu hay khong
        If .Range("B12") <> "" Then
            dongCuoi = tinhdongcuoi("J12:J1048576") + 100
            .Range("A12:V" & dongCuoi).Clear
        End If
        'Xu ly lenh
        If Nam <> 0 Then
            'Viet query  Select
            Query = "Select SP_SanPham.MaSanPham, " & _
            "SP_SanPham.TenSanPham,ISNULL(DL_NT.SoLuongBanNam, 0) As SoLuongBanNamTruoc,ISNULL(DL_NT.DonGiaVon, 0) As DonGiaVonTruoc,ISNULL(DL_NT.GiaVonNam, 0) As GiaVonNamTruoc, " & _
            "ISNULL(DL_NT.GiaBanBinhQuanNam, 0) As GiaBanBinhQuanNam,ISNULL(DL_NT.DoanhThuNam, 0) As DoanhThuNamTruoc,ISNULL(DL_NT.TiLeGiaVonNam, 0) As TiLeGiaVonNamTruoc, " & _
            " isNull(TDBLN.TiLeTangTruongNamNay,0) As TiLeTangTruongNamNay,isNull(TDBLN.SoLuongBan,0) As SoLuongBanNamNay,isNull(TDBLN.GiaVon,0) As DonGiaVonNamNay,isNull(TDBLN.GiaBanBinhQuan,0) As GiaBanBinhQuanNamNay,isNull(TDBLN.DoanhThuDuKien,0) As DoanhThuNamNay,isNull(TDBLN.TiLeTangTruongNamNay,0) As TiLeGiaVonNamNay, " & _
            "SP_SanPham.NhomVTHH1 , SP_SanPham.NhomVTHH2, SP_SanPham.NhomVTHH3, SP_SanPham.NhomVTHH4, SP_SanPham.NhomVTHH5, SP_SanPham.NhomVTHH6 " & _
            "FROM SP_SanPham " & _
            "LEFT JOIN (Select * from KD_layDuLieuNamTruoc('" & Nam - 1 & "')) As DL_NT on SP_SanPham.MaSanPham = DL_NT.MaSanPham  " & _
            "LEFT JOIN (Select * from KD_TheoDoiBienLoiNhuan WHERE Nam = " & Nam & ") TDBLN on SP_SanPham.MaSanPham = TDBLN.MaSanPham"

            '        Mo ket noi csdl
            Dim dbConn As Object
            Set dbConn = ConnectToDatabase
            Call viewSheet(Query, SheetTheoDoiBienLoiNhuan, "B12", dbConn)

            'LayTongDoanhThuNamTruoc
            Dim Rs As Object
            Dim DT As Variant
            Query = "Select Sum(DoanhThuNam) from KD_layDuLieuNamTruoc('" & Nam - 1 & "')"

            Set Rs = dbConn.Execute(Query)

            If Not Rs.EOF And Not Rs.BOF Then
                DT = Rs.GetRows()

                .Range("H10").Value = DT(0, 0)

            End If
            'Dong Ket noi
            Call CloseDatabaseConnection(dbConn)
            dongCuoi = tinhdongcuoi("J12:J1048576")

            Dim i

            For i = 12 To dongCuoi
                .Range("J" & i) = 0
                .Range("K" & i) = "=ROUND(IFERROR(D" & i & "* (J" & i & " + 1),0),0)"
                .Range("L" & i) = "=IFERROR(F" & i & "/ D" & i & ",0)"
                .Range("M" & i) = "=IFERROR(K" & i & "* L" & i & ",0)"
                .Range("N" & i) = "=G" & i
                .Range("O" & i) = "=IFERROR(K" & i & "* N" & i & ",0)"
                .Range("P" & i) = "=IFERROR(L" & i & "/ M" & i & ",0)"
            Next i

            If dongCuoi <= 11 Then
                .ListObjects("Table_TheoDoiBienLoiNhuan").Resize Range("$B$11:$V$12")
            Else
                .ListObjects("Table_TheoDoiBienLoiNhuan").Resize Range("$B$11:$V$" & dongCuoi)
            End If
        Else
            MsgBox "Chon Nam"
        End If


    End With
    DinhDangBienLoiNhuan
    TatLimit
    ThongBao_ThanhCong
End Sub

Sub DinhDangBienLoiNhuan()

    On Error Resume Next
    With SheetTheoDoiBienLoiNhuan
        .Select

        Dim dongCuoi
        dongCuoi = tinhdongcuoi("J12:J1048576")

        If dongCuoi <= 11 Then
            Format_ dongCuoi, 22, "B11", "A:A", 0, "Table_TheoDoiBienLoiNhuan"
            F_BoderStyle .Range("B12:V12"), "Table_TheoDoiBienLoiNhuan"
        Else
            Format_ dongCuoi, 22, "B11", "A:A", 0, "Table_TheoDoiBienLoiNhuan"
            F_BoderStyle .Range("B12:V" & dongCuoi + 2), "Table_TheoDoiBienLoiNhuan"
        End If

        Columns("E:H").Select
        FormatMoney
        Columns("L:O").Select
        FormatMoney

        'To mau phan duoc nhap

        If dongCuoi >= 12 Then
            .Range("J12:J" & dongCuoi).Select
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark2
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
        End If

        Columns("J:J").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"

        Columns("P:P").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
        Selection.NumberFormat = "0.00%"

        Columns("I:I").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"

        Range("F9").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
        ScrollToTop
    End With

    FixTop Range("A12")
End Sub

Sub CapNhatBienDoiLoiNhuan()
    BatLimit
    On Error Resume Next
    Dim Nam As Integer
    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim StrCnn As String
    Dim Query As String

    Set Rs = New ADODB.Recordset
    Set Cn = New ADODB.Connection
    StrCnn = KetNoiMayChu_KhachHang

    Nam = Year(Now)
    Cn.Open StrCnn
    With SheetTheoDoiBienLoiNhuan
        .Select
        If .cbNam.Value <> "" Then
            Nam = .cbNam.Value
        Else
            Nam = Year(Now)
        End If

        If Nam <> 0 Then
            'Xoa du lieu truoc khi cap nhat

            Query = "DELETE FROM KD_TheoDoiBienLoiNhuan WHERE Nam = " & Nam
            Rs.Open Query, Cn, adOpenStatic

            'Cap Nhat
            Set Rs = New ADODB.Recordset

            Dim SoLuongBan As String
            Dim GiaVon As String
            Dim DoanhThuDuKien As String
            Dim MaSanPham As String
            Dim TenSanPham As String
            Dim GiaBanBinhQuan As String
            Dim TiLeTangTruongNamNay As String
            Dim dongCuoi
            Dim i

            dongCuoi = tinhdongcuoi("B12:B1048576")
            For i = 12 To dongCuoi

                MaSanPham = .Range("B" & i)
                TenSanPham = .Range("C" & i)
                SoLuongBan = .Range("K" & i)
                GiaVon = .Range("M" & i)
                DoanhThuDuKien = .Range("O" & i)
                GiaBanBinhQuan = .Range("N" & i)
                TiLeTangTruongNamNay = .Range("J" & i)

                If MaSanPham <> "" Then

                    Query = "INSERT INTO KD_TheoDoiBienLoiNhuan(MaSanPham,TenSanPham,SoLuongBan,GiaVon,DoanhThuDuKien,Nam, GiaBanBinhQuan, TiLeTangTruongNamNay) " & _
                    "VALUES('" & MaSanPham & "',N'" & TenSanPham & "'," & SoLuongBan & "," & GiaVon & "," & DoanhThuDuKien & ",'" & Nam & "', '" & GiaBanBinhQuan & "', '" & TiLeTangTruongNamNay & "')"

                    Rs.Open Query, Cn, adOpenStatic
                Else
                    MsgBox "Nhap day du thong tin"

                    Cn.Close
                    Set Rs = Nothing
                    Set Cn = Nothing
                 Exit Sub
                End If

            Next i

        Else

            MsgBox "Chon Nam"
        End If

        'Dong ket noi
        Cn.Close
        Set Rs = Nothing
        Set Cn = Nothing

    End With
    TatLimit
    ThongBao_ThanhCong
End Sub

Sub ChonNam()
    Call F_KhoiTaoNam(SheetTheoDoiBienLoiNhuan.cbNam)

    SheetTheoDoiBienLoiNhuan.cbNam.Text = Sheet11.cbbSheetNam.Value
End Sub




