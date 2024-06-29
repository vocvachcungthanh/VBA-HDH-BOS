Attribute VB_Name = "M_Hai"
Function Run_script(SQLStr As String)
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset
    StrCnn = KetNoiMayChu_KhachHang
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Rs.Open SQLStr, Cn, adOpenStatic
    
    Cn.Close
    Set Cn = Nothing
End Function

Function UniConvert(Text As String, InputMethod As String) As String
  Dim VNI_Type, Telex_Type, CharCode, Temp, i As Long
  UniConvert = Text
  VNI_Type = Array("a81", "a82", "a83", "a84", "a85", "a61", "a62", "a63", "a64", "a65", "e61", _
      "e62", "e63", "e64", "e65", "o61", "o62", "o63", "o64", "o65", "o71", "o72", "o73", "o74", _
      "o75", "u71", "u72", "u73", "u74", "u75", "a1", "a2", "a3", "a4", "a5", "a8", "a6", "d9", _
      "e1", "e2", "e3", "e4", "e5", "e6", "i1", "i2", "i3", "i4", "i5", "o1", "o2", "o3", "o4", _
      "o5", "o6", "o7", "u1", "u2", "u3", "u4", "u5", "u7", "y1", "y2", "y3", "y4", "y5")
  Telex_Type = Array("aws", "awf", "awr", "awx", "awj", "aas", "aaf", "aar", "aax", "aaj", "ees", _
      "eef", "eer", "eex", "eej", "oos", "oof", "oor", "oox", "ooj", "ows", "owf", "owr", "owx", _
      "owj", "uws", "uwf", "uwr", "uwx", "uwj", "as", "af", "ar", "ax", "aj", "aw", "aa", "dd", _
      "es", "ef", "er", "ex", "ej", "ee", "is", "if", "ir", "ix", "ij", "os", "of", "or", "ox", _
      "oj", "oo", "ow", "us", "uf", "ur", "ux", "uj", "uw", "ys", "yf", "yr", "yx", "yj")
  CharCode = Array(ChrW(7855), ChrW(7857), ChrW(7859), ChrW(7861), ChrW(7863), ChrW(7845), ChrW(7847), _
      ChrW(7849), ChrW(7851), ChrW(7853), ChrW(7871), ChrW(7873), ChrW(7875), ChrW(7877), ChrW(7879), _
      ChrW(7889), ChrW(7891), ChrW(7893), ChrW(7895), ChrW(7897), ChrW(7899), ChrW(7901), ChrW(7903), _
      ChrW(7905), ChrW(7907), ChrW(7913), ChrW(7915), ChrW(7917), ChrW(7919), ChrW(7921), ChrW(225), _
      ChrW(224), ChrW(7843), ChrW(227), ChrW(7841), ChrW(259), ChrW(226), ChrW(273), ChrW(233), ChrW(232), _
      ChrW(7867), ChrW(7869), ChrW(7865), ChrW(234), ChrW(237), ChrW(236), ChrW(7881), ChrW(297), ChrW(7883), _
      ChrW(243), ChrW(242), ChrW(7887), ChrW(245), ChrW(7885), ChrW(244), ChrW(417), ChrW(250), ChrW(249), _
      ChrW(7911), ChrW(361), ChrW(7909), ChrW(432), ChrW(253), ChrW(7923), ChrW(7927), ChrW(7929), ChrW(7925))
  Select Case InputMethod
    Case Is = "VNI": Temp = VNI_Type
    Case Is = "Telex": Temp = Telex_Type
  End Select
  For i = 0 To UBound(CharCode)
    UniConvert = Replace(UniConvert, Temp(i), CharCode(i))
    UniConvert = Replace(UniConvert, UCase(Temp(i)), UCase(CharCode(i)))
  Next i
End Function

Function UNItoVBA(ByVal MyStr As String) As String
     'Chuyen chuoi tu UNICODE sang Code VBA
     Dim str As String, i As Integer, CStart As Integer, CCount As Integer, Status As Boolean
     str = "-225-224-7843-227-7841-259-7855-7857-7859-7861-7863-226-7845-7847-7849-7851-7853-273-233-232-7867-7869-7865-234-7871-7873-7875-7877-7879-237-236-7881-297-7883-243-242-7887-245-7885-244-7889-7891-7893-7895-7897-417-7899-7901-7903-7905-7907-250-249-7911-361-7909-432-7913-7915-7917-7919-7921-253-7923-7927-7929-7925-193-192-7842-195-7840-258-7854-7856-7858-7860-7862-194-7844-7846-7848-7850-7852-272-201-200-7866-7868-7864-202-7870-7872-7874-7876-7878-205-204-7880-296-7882-211-210-7886-213-7884-212-7888-7890-7892-7894-7896-416-7898-7900-7902-7904-7906-218-217-7910-360-7908-431-7912-7914-7916-7918-7920-221-7922-7926-7928-7924-10-"
     For i = 1 To Len(MyStr)
          If InStr(str, "-" & AscW(Mid(MyStr, i, 1)) & "-") = 0 Then
               If Not Status Then
                    CStart = i:        Status = True
               End If
               CCount = CCount + 1
          Else
               If Status Then UNItoVBA = UNItoVBA & IIf(UNItoVBA = "", "", " & ") & """" & Replace(Mid(MyStr, CStart, CCount), """", """""") & """"
               Status = False
               CCount = 0
               UNItoVBA = UNItoVBA & IIf(UNItoVBA = "", "", " & ") & "ChrW(" & AscW(Mid(MyStr, i, 1)) & ")"
          End If
     Next
     If Status Then UNItoVBA = UNItoVBA & IIf(UNItoVBA = "", "", " & ") & """" & Replace(Mid(MyStr, CStart, CCount), """", """""") & """"
End Function

Function ThongBao_ThanhCong()

    TieuDe = "BOS xin thông báo"
    noidung = "Th" & ChrW(7921) & "c hi" & ChrW(7879) & _
        "n thành công!"
    Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
End Function

Function ThongBao_DangNhap_ThanhCong()

    TieuDe = "BOS xin thông báo"
    noidung = "K" & ChrW(7871) & "t n" & ChrW(7889) & _
        "i thành công."
    Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
End Function


Function ThongBao_SaiThongTinDangNhap()

    TieuDe = "BOS xin thông báo"
    noidung = "Thông tin " & ChrW(273) & ChrW(259) & "ng nh" & ChrW(7853) & "p b" & ChrW(7883) & " sai. Vui lòng ki" & ChrW(7875) & "m tra l" & ChrW(7841) & "i."""
    Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
End Function

Function ThongBao_ChucNangChuaCo()

    TieuDe = "BOS xin thông báo"
    noidung = "Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng " & _
        ChrW(273) & "ang " & ChrW(273) & ChrW(432) & ChrW(7907) & "c c" & ChrW( _
        7853) & "p nh" & ChrW(7853) & "t. Vui lòng th" & ChrW(7917) & " l" & ChrW _
        (7841) & "i sau."
    Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
End Function

Function ThongBao_SaiKieuDuLieu()

    TieuDe = "BOS xin thông báo"
    noidung = "Ki" & ChrW(7875) & "u d" & ChrW(7919) & " li" & _
        ChrW(7879) & "u nh" & ChrW(7853) & "p vào không " & ChrW(273) & "úng " & _
        ChrW(273) & ChrW(7883) & "nh d" & ChrW(7841) & "ng yêu c" & ChrW(7847) & _
        "u. Vui lòng ki" & ChrW(7875) & "m tra l" & ChrW(7841) & _
        "i."
    Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
End Function


Function ThongBao_DuLieuQuanTrong()

    TieuDe = "BOS xin c" & ChrW(7843) & "nh báo"
    noidung = ChrW(272) & "ây là d" & ChrW(7919) & " li" & ChrW _
        (7879) & "u r" & ChrW(7845) & "t quan tr" & ChrW(7885) & "ng. B" & ChrW(7841) & "n có ch" & ChrW(7855) & "c ch" & ChrW(7855) & "n th" & ChrW(7921) & "c hi" & ChrW(7879) & "n c" & ChrW(7853) & "p nh" & ChrW(7853) & "t?" & _
    " Có th" & ChrW(7875) & " " & ChrW(7843) & "nh h" & ChrW(432) & ChrW(7903) & "ng " & ChrW(273) & ChrW(7871) & "n nhi" & _
        ChrW(7873) & "u ph" & ChrW(7847) & "n khác trong h" & ChrW(7879) & " th" _
        & ChrW(7889) & "ng."
   Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOKCancel, msoAlertIconWarning, 0, 0, 0
    
End Function

Function ThongBao_LoiKetNoiMayChu()

    TieuDe = "BOS xin thông báo"
    noidung = "K" & ChrW(7871) & "t n" & ChrW(7889) & "i " & _
        ChrW(273) & ChrW(7871) & "n máy ch" & ChrW(7911) & " không th" & ChrW( _
        7921) & "c hi" & ChrW(7879) & "n " & ChrW(273) & ChrW(432) & ChrW(7907) & _
         "c. Vui lòng ki" & ChrW(7875) & "m tra " & "l" & ChrW(7841) & "i m" & ChrW(7841) & "ng ho" & _
         ChrW(7863) & "c thông tin máy ch" & ChrW(7911) & " r" & ChrW(7891) & _
        "i th" & ChrW(7917) & " l" & ChrW(7841) & "i. Xin c" & ChrW(7843) & "m " _
        & ChrW(417) & "n."
    Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
End Function

Function ThongBao_NhapThieuDuLieu()

    TieuDe = "BOS xin thông báo"
    noidung = "D" & ChrW(7919) & " li" & ChrW(7879) & "u nh" & _
        ChrW(7853) & "p vào ch" & ChrW(432) & "a " & ChrW(273) & ChrW(7847) & _
        "y " & ChrW(273) & ChrW(7911) & ". Yêu  c" & ChrW(7847) & "u b" & ChrW( _
        7893) & " sung và th" & ChrW(7917) & " l" & ChrW(7841) & "i. Xin c" & ChrW(7843) & "m " & ChrW(417) & "n."
    Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
End Function

Sub Load_Theme_TuDangNhap()
    Dim Theme As String
    
    On Error Resume Next
    Theme = "Integral"
    
    On Error Resume Next
    Theme = Workbooks("Core.xlsb").Sheets("PhanQuyen").Range("AB1").Value
    
    If Theme <> "" Then
    
    On Error Resume Next
    ActiveWorkbook.ApplyTheme ( _
        "C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\" & Theme & " .thmx" _
        )
    
    On Error Resume Next
    ActiveWorkbook.ApplyTheme ( _
        "C:\Program Files\Microsoft Office\Document Themes 16\" & Theme & ".thmx")

    End If
End Sub

Sub MoPhanQuyen()
    PhanQuyen.Show
End Sub

Sub MoDangNhap()
    DangNhap.Show
End Sub

Sub Moform_LoTrinhTangBac()
    frmLoTrinhTangBac.Show
End Sub

Sub Moform_CaiDat()
    FrmCaiDat.Show
End Sub

Sub Moform_DienBienLuong()
    FrmDienBienLuong.Show
End Sub

Sub MoSheets_KeHoachChiPhi()
    Sheets("KeHoachChiPhi").Select
End Sub

Sub Moform_ImportDuLieu()
    frmQuanLyDuLieu.Show
End Sub

Sub CheckFile_TonTai()
    If Dir("G:\My Drive\Product 2023\Core.xlsb") = "" Then
        MsgBox "File does not exist."
    Else
        MsgBox "OK ok"
    End If

End Sub

Sub TenFile_DangMo()
    Dim wk As Workbook
        For Each wk In Application.Workbooks
    Next wk
End Sub


Sub Load_DuLieu()

    Dim lr As Long
    lr = Worksheets("Data").Range("B150000").End(xlUp).Row
    If lr >= 12 Then
        Worksheets("Data").Range("B12:S" & lr).Clear
    End If
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    StrCnn = KetNoiMayChu_KhachHang
    Dim SQLStr As String
    
    Dim Nam As Integer, Thang As Integer
    Dim MaNhanVien As String
    Nam = cbbNam.Text
    Thang = cbbThang.Text
    TenDangNhap = "tuannx"
    
    SQLStr = " Select NgayHoaDon, SoHoaDon, MaKhachHang,  " & _
                " MaSanPham, HangKhuyenMai, DonViTinh, SoLuongKhuyenMai, SoLuong, DonGia, DoanhSo, ChietKhau, " & _
                " SoLuongTraLai , GiaTriTraLai, GiaTriGiamGia, TongThanhToan, DonGiaVon, GiaVon, NguoiBan " & _
            " from KD_DonHang Left join Ns_Nhanvien on KD_DonHang.NguoiBan  = NS_NhanVien.MaNhanVien " & _
            " Where right(NgayHoaDon,4) = " & Nam & " " & _
                " and (convert(int, SubString(NgayHoaDon, 4,2)) = " & Thang & " or " & Thang & " = 0 ) " & _
                " and NS_NhanVien.PhongBanID  in (Select PQ_NguoiDung_PhongBan.PhongBanID " & _
                                                " from PQ_NguoiDung_PhongBan inner join PQ_NguoiDung on PQ_NguoiDung_PhongBan.NguoiDungID  = PQ_NguoiDung.NguoiDungID " & _
                                                " where PQ_NguoiDung.TenDangNhap  = N'" & TenDangNhap & "')"

    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Rs.Open SQLStr, Cn, adOpenStatic

    Worksheets("Data").Range("B12").CopyFromRecordset Rs

    
    Cn.Close
    
    Set Rs = Nothing
    Set Cn = Nothing

End Sub

Sub Load_Master_Data()

    Worksheets("TestDuLieu").Range("B11:AY1000").Clear
    Dim DsHienThi As Variant
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    StrCnn = KetNoiMayChu_KhachHang
    Dim SQLStr As String

    SQLStr = "Select TenHienThi, LenhSQL, CotExcel from HT_Hienthi_MasterData"
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Rs.Open SQLStr, Cn, adOpenStatic

    If Not Rs.EOF Then
        DsHienThi = Rs.GetRows()
    End If

    Dim i As Integer
    Dim Ten As String
    Dim LenhSQL As String
    Dim CotExcel As String

    For i = 0 To UBound(DsHienThi, 2)
        Ten = DsHienThi(0, i)
        LenhSQL = DsHienThi(1, i)
        CotExcel = DsHienThi(2, i)
        'MsgBox i & " >> " & Ten & " >> " & LenhSQL
        
        Set Rs = New ADODB.Recordset
        Rs.Open LenhSQL, Cn, adOpenStatic
        Worksheets("TestDuLieu").Range(CotExcel & "11").CopyFromRecordset Rs
    
    Next i

    
    Cn.Close
    
    Set Rs = Nothing
    Set Cn = Nothing
End Sub


Function ThongTinDangNhap() As Variant
    Dim TimKiem As String
    TimKiem = Sheets("PhanQuyen").Range("I1").Value
    
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset
    Dim Field As Field
    Set Rs = New ADODB.Recordset

    StrCnn = KetNoiMayChu_KhachHang
    Dim SQLStr As String
    SQLStr = "exec ChaoMung_DangNhap " & "'" & TimKiem & "'"
    
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Rs.Open SQLStr, Cn, adOpenStatic

    Do While Not Rs.EOF
    ThongTinDangNhap = Rs.GetRows()
    Loop

    
    Cn.Close
    
    Set Rs = Nothing
    Set Cn = Nothing

End Function


Function TinhThue_TNCN(ThuNhapChiuThue As Long)
        Select Case ThuNhapChiuThue
        Case Is <= 5000000
                TinhThue_TNCN = 0
        Case 5000001 To 1048576000
                TinhThue_TNCN = ThuNhapChiuThue * 0.1 - 250000
        Case 1048576001 To 18000000
                TinhThue_TNCN = ThuNhapChiuThue * 0.15 - 750000
        Case 18000001 To 32000000
                TinhThue_TNCN = ThuNhapChiuThue * 0.2 - 1650000
        Case 32000001 To 52000000
                TinhThue_TNCN = ThuNhapChiuThue * 0.25 - 3250000
         Case 52000001 To 80000000
                TinhThue_TNCN = ThuNhapChiuThue * 0.3 - 5850000
        Case Else
                TinhThue_TNCN = ThuNhapChiuThue * 0.35 - 9850000
    End Select
End Function

Sub Test_Thue()
    Dim ThuNhap As Long
    ThuNhap = InputBox("Nhap muc thu nhap:")
    Dim Thue As Long
    Thue = TinhThue_TNCN(ThuNhap)
    MsgBox Thue
End Sub

Sub Xoa_Dulieu_SoBanHang()
     Workbooks("KD.xlsb").Worksheets("Data").Range("A4:u104857600").Clear
End Sub
