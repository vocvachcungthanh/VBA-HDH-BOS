Attribute VB_Name = "M_DataKeHoachDoanhThu_SP"

Sub Load_KHKD_TheoSanPham()
    Dim ThangKh As String
    Dim ThangLs As String
    Dim dongCuoi As Long

    ThangKh = layDanhSachThangLapKh
    ThangLs = layDanhSachThangLichSu
    Set wSheet = Worksheets("Data KHDT SP")
    With wSheet
        .Select
        dongCuoi = tinhdongcuoi("B12:B1048576")
        If .Range("B12") <> 0 Then
            .Range("A12:Y" & dongCuoi).Clear
        End If

        On Error Resume Next
        Dim NamKh As Integer

        If .cbbNam.Text = "" Then
            NamKh = Year(Now)
        Else
            NamKh = .cbbNam.Text
        End If

        Dim NamLs As Integer

        If .cbbNamLichSu.Text = "" Then
            NamLs = Year(Now)
        Else
            NamLs = .cbbNamLichSu.Text
        End If

        If ThangKh = "" Then
            ThangKh = 1
            Sheet1611.lbChonThangLapKH.Selected(0) = True
        End If

        If ThangLs = "" Then
            ThangLs = 1
            Sheet1611.lbThangLichSu.Selected(0) = True
        End If

        On Error Resume Next

        Dim Query As String
        
        '        Mo ket noi csdl
        Dim dbConn As Object
        Set dbConn = ConnectToDatabase
        
        Query = " exec DataKHDT_SP_KD_V2 " & NamKh & ",0,'" & ThangKh & "'," & NamLs & ", '" & ThangLs & "'"
        Call viewSheet(Query, Sheet1611, "B12", dbConn)

        'Tong doanh thu theo san pham
        Query = "exec KD_TK_TongHopTheo_SP " & NamKh & ", " & NguoiDungID & ", 0,'" & ThangKh & "'"
        Call viewSheet(Query, Sheet1611, "H5", dbConn)
        
        'Dong Ket noi
        Call CloseDatabaseConnection(dbConn)

        'Tinh Doanh thu

        TinhDoanhThu dongCuoi
        If Range("B12") <> "" Then
            ActiveSheet.ListObjects("TableSanPham").Resize Range("B11:J" & dongCuoi)
        End If

    End With
    
     Set wSheet = Nothing
End Sub

Sub F_Style_SP()
  
    Dim i As Integer
    Dim dongCuoi As Long
    Set wSheet = Sheet1611
    With wSheet
        .Select
        dongCuoi = tinhdongcuoi("B12:B1048576")
        ' Dinh dan duong vien "Theo cong thuc cua phong"

        Format_ dongCuoi, 11, "B11", "A:A", 0, "TableSanPham"

        Columns("H").Select
        FormatMoney
        Columns("F").Select
        FormatMoney
        Columns("J").Select
        FormatMoney

        F_Width "B:B", 5
        F_Width "K:K", 0

        'To mau phan duoc nhap
        If dongCuoi >= 12 Then
              .Range("I12:I" & dongCuoi).Select
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
      

        If .Range("H5") < 0 Then
            .Range("H5").Select
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        Else
            .Range("H5").Select
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
        End If

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

        F_BoderStyle .Range("B12:J" & dongCuoi), "TableSanPham"

        FixTop .Range("A12")
    End With
  Set wSheet = Nothing
End Sub

Sub TinhDoanhThu(dongCuoi)
    Dim i As Long
    Set wSheet = Sheet1611
    With wSheet
        For i = 12 To dongCuoi
            .Range("J" & i) = "=H" & i & " * I" & i & ""
        Next i
    End With
  Set wSheet = Nothing
End Sub

Function loadDataSanPham()
    If NguoiDungID <> 0 Then
        Load_KHKD_TheoSanPham
    End If
End Function

Public Sub ComboBox_KhoiTaoThangNam_SP()
    On Error Resume Next
    Set wSheet = Sheet1611
    With wSheet
        .Select
      
        'Nam Lap Ke Hoach va cbNamLichSu
        .cbbNam.Clear
        .cbbNamLichSu.Clear
        
        Call F_KhoiTaoNam(Sheet1611.cbbNam)
        .cbbNam.Text = Sheet11.cbbSheetNam.Value
        .cbbNamLichSu.Value = Year(Date)

        'Thang lap ke hoach, thang lich su
        Dim Thang As Long
        .lbChonThangLapKH.Clear
        For Thang = 1 To 12
            .lbChonThangLapKH.AddItem "Tháng " & Thang
            .lbThangLichSu.AddItem "Tháng " & Thang
        Next Thang

        .lbChonThangLapKH.Selected(0) = True
        .lbThangLichSu.Selected(0) = True
    End With
    
    Set wSheet = Nothing
End Sub

Sub cmdLoadData()
    BatLimit
    Set wSheet = Sheet1611
    With wSheet
        .Select
        
        If .cbbNamLichSu.ListCount <= 0 Or _
           .lbThangLichSu.ListCount <= 0 Or _
           .cbbNam.ListCount <= 0 Or _
           .lbChonThangLapKH.ListCount <= 0 _
        Then
            Call ComboBox_KhoiTaoThangNam_SP
        End If
    End With
    
    loadDataSanPham
    Set wSheet = Nothing
    F_Style_SP
    TatLimit
ThongBao_ThanhCong
End Sub

Function XoaBoKyTuThang(dauVao) As Variant

    Dim strArray() As String
    Dim i As Integer

    strArray = Split(dauVao, ",")

    For i = 0 To UBound(strArray)
        strArray(i) = Replace(strArray(i), "Tháng ", "")
    Next i
    XoaBoKyTuThang = Join(strArray, ",")

End Function

Function layDanhSachThangLapKh() As String
    Dim Thang As Integer
    Dim DsThangApDung As String

    DsThangApDung = ""
    Set wSheet = Sheet1611
    With wSheet
        .Select
        'Thang lap ke hoach
        For Thang = 0 To .lbChonThangLapKH.ListCount - 1
            If .lbChonThangLapKH.Selected(Thang) = True Then
                DsThangApDung = DsThangApDung & .lbChonThangLapKH.List(Thang, 0) & ","
            End If
        Next Thang
    End With
    Set wSheet = Nothing
    If DsThangApDung <> "" Then
        layDanhSachThangLapKh = XoaBoKyTuThang(Left(DsThangApDung, Len(DsThangApDung) - 1))
    Else
        layDanhSachThangLapKh = XoaBoKyTuThang("")
    End If

End Function

Function layDanhSachThangLichSu() As String
    Dim Thang As Integer
    Dim DsThangApDung As String

    DsThangApDung = ""
    Set wSheet = Sheet1611
    With wSheet
        .Select
        ' Thang lich su

        For Thang = 0 To .lbThangLichSu.ListCount - 1
            If .lbThangLichSu.Selected(Thang) = True Then
                DsThangApDung = DsThangApDung & .lbThangLichSu.List(Thang, 0) & ","
            End If
        Next Thang
    End With
    Set wSheet = Nothing
    If DsThangApDung <> "" Then
        layDanhSachThangLichSu = XoaBoKyTuThang(Left(DsThangApDung, Len(DsThangApDung) - 1))
    Else
        layDanhSachThangLichSu = XoaBoKyTuThang("")
    End If

End Function

Sub CopySoLuongKyNamTruoc()
    BatLimit
    Dim dongCuoi

   Set wSheet = Sheet1611
    With wSheet

        .Select
         dongCuoi = tinhdongcuoi("B12:B1048576")
        .Range("E12:E" & dongCuoi).Select
        Selection.Copy
        .Range("I12:I" & dongCuoi).Select
        ActiveSheet.Paste
    End With
    Set wSheet = Nothing
    TatLimit
    
    ThongBao_ThanhCong
End Sub

Sub Save_KHKD_TheoSanPham()
    On Error Resume Next
    BatLimit
    Dim Nam As String
    Dim SanPhamID
    Dim SoLuong As Double
    Dim KyLapKeHoach As String

    Dim connect As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim StrCnn As String

    Set Rs = New ADODB.Recordset
    Set connect = New ADODB.Connection
    StrCnn = KetNoiMayChu_KhachHang
     Set wSheet = Sheet1611
    With wSheet
        .Select
        Nam = .cbbNam.Value
        KyLapKeHoach = layDanhSachThangLapKh

        connect.Open StrCnn

        If KyLapKeHoach = "" Or Nam = "" Then
            CreateObject("WScript.Shell").Popup "Ch" & ChrW(7885) & "n n" & ChrW(259) & "m v" & ChrW(224) & " th" & ChrW(225) & "ng l" & ChrW(7853) & "p k" & ChrW(7871) & " ho" & ChrW(7841) & "ch", , "Bos th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 16
            connect.Close
            Set connect = Nothing
         Exit Sub
        Else
            Dim i
            Dim dongCuoi
            Dim Query As String

            Set Rs = New ADODB.Recordset
            dongCuoi = tinhdongcuoi("B12:B1048576")

            For i = 12 To dongCuoi
                SanPhamID = .Range("K" & i).Value
                SoLuong = .Range("I" & i).Value

                Query = " delete from KeHoachDTSanPham where Nam = " & Nam & " And KyLapKeHoach = '" & KyLapKeHoach & "' And SanPhamID = " & SanPhamID & ";" & _
                "INSERT INTO KeHoachDTSanPham (Nam, NhanVienID, SanPhamID, SoLuong, KyLapKeHoach)" & _
                " VALUES(" & Nam & ",0," & SanPhamID & "," & SoLuong & ",'" & KyLapKeHoach & "')"
                Rs.Open Query, connect, adOpenStatic
            Next i
            connect.Close
            Set connect = Nothing
        End If

    End With
    Set wSheet = Nothing
    TatLimit
    
    ThongBao_ThanhCong
End Sub




