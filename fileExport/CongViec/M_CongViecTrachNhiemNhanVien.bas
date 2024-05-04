Option Explicit

Private SelectPhongBan As Object
Private SelectPhuongThucTinh As Object
Private SelectNam As Object

Private InputTenCongViec As Object
Private InputDonViLuongHoa As Object
Private InputDinhMucToiThieu As Object
Private InputDinhMucYeuCau As Object
Private InputHeSo As Object
Private InputGhiChu As Object
Private FrameTTCT As Object

Private TableDanhSachNhanVienViTri As Object
Private TableCongViecTrachNhiemTheoViTri As Object

Private Nam As Object
Private Thang_1 As Object
Private Thang_2 As Object
Private Thang_3 As Object
Private Thang_4 As Object
Private Thang_5 As Object
Private Thang_6 As Object
Private Thang_7 As Object
Private Thang_8 As Object
Private Thang_9 As Object
Private Thang_10 As Object
Private Thang_11 As Object
Private Thang_12 As Object

Private addedItemsNhanVienID As Collection
Private addedItemsCongViecID As Collection
Private addedItemsMucTieuCv As Collection
Private Thang As Integer
Private ValueNam As Boolean
Private IndexTrachNhiemViTri As Integer
Private TrachNhiemTheoViTriID As Integer
Private LabeViTri As Object
Private LabelNhanVien As Object
Public IndexNhanvienViTri As Integer

Public bgDefault As String
Public bgActive As String
Public bgCheck As String
Public PhongBanID_Global As Integer
Public NhanVienID_Global As Integer
Public CongViecID_Global As Integer
Private TieuDe As String
Private NoiDung As String
Private CongThucTinhID As Integer
Public IndexThang As Integer

Function F_KhoiTaoCongViecTrachNhiemNhanVien()
   With Form_CongViecTrachNhiemNhanVien
      Set SelectPhongBan = .cbbPhongBan
      Set SelectPhuongThucTinh = .cbbPhuongThucTinh
      Set SelectNam = .cbNam

      Set InputTenCongViec = .txtMaCongViec
      Set InputDonViLuongHoa = .txtDonViLuongHoa
      Set InputDinhMucToiThieu = .txtDinhMucToiThieu
      Set InputDinhMucYeuCau = .txtDinhMucYeuCau
      Set InputHeSo = .txtHeSo
      Set InputGhiChu = .txtGhiChu

      Set TableDanhSachNhanVienViTri = .lvDanhSachNhanVienViTri
      Set TableCongViecTrachNhiemTheoViTri = .lvCongViecTrachNhiemTheoViTri
      Set FrameTTCT = .FrameThongTinChiTiet
      Set LabeViTri = .LabeViTri
      Set LabelNhanVien = .LabelNhanVien

      Set Nam = .btnCaNam
      Set Thang_1 = .btnT1
      Set Thang_2 = .btnT2
      Set Thang_3 = .btnT3
      Set Thang_4 = .btnT4
      Set Thang_5 = .btnT5
      Set Thang_6 = .btnT6
      Set Thang_7 = .btnT7
      Set Thang_8 = .btnT8
      Set Thang_9 = .btnT9
      Set Thang_10 = .btnT10
      Set Thang_11 = .btnT11
      Set Thang_12 = .btnT12

      Set addedItemsNhanVienID = New Collection
      Set addedItemsCongViecID = New Collection

      TieuDe = "BOS xin thông báo"

      bgDefault = &H80000005
      bgActive = &H80FF&
      bgCheck = &HC000&
   End With

   TableDanhSachNhanVienViTri.Font.Charset = VIETNAMESE_CHARSET
   TableCongViecTrachNhiemTheoViTri.Font.Charset = VIETNAMESE_CHARSET

   Call F_HeaderTableDanhSachNhanVienVitri
   Call F_HeaderTableCongViecTrachNhiemTheoViTri

   Call F_DanhSachPhongBan(SelectPhongBan)
   Call F_KhoiTaoNam(SelectNam)

   SelectNam.value = Sheet2.cbbNamSheetCongViec.value
End Function

Function F_HeaderTableDanhSachNhanVienVitri()
   Dim TenPhongBan As String
   Dim TenNhanVien As String
   Dim ViTriCongViec As String
   TenPhongBan = "Tên phòng ban"
   TenNhanVien = "Tên nhân viên"
   ViTriCongViec = "V" & ChrW(7883) & " trí công vi" & ChrW(7879) & _
   "c"

   With TableDanhSachNhanVienViTri.ColumnHeaders
      .Add , , UniToWindows1258(TenPhongBan), 0
      .Add , , UniToWindows1258(TenNhanVien), 100
      .Add , , "NhanVienID", 0
      .Add , , "ViTriID", 0
      .Add , , UniToWindows1258(ViTriCongViec), 115
   End With

End Function

Function F_HeaderTableCongViecTrachNhiemTheoViTri()
   Dim TenCongViec As String
   Dim DonViLuongHoa As String
   Dim DinhMucToiThieu As String
   Dim DinhMucYeuCau As String
   Dim PhuongThucTinh As String
   Dim GhiChu As String
   Dim HeSo As String
   Dim CongViecID As Integer
   Dim DonViLuongHoaID As Integer
   Dim DaLap As String

   TenCongViec = "Tên công vi" & ChrW(7879) & "c"

   DonViLuongHoa = ChrW(272) & ChrW(417) & "n v" & ChrW(7883) & " l" _
   & ChrW(432) & ChrW(7907) & "ng hóa"

   DinhMucToiThieu = ChrW(272) & ChrW(7883) & "nh m" & ChrW(7913) & _
   "c t" & ChrW(7889) & "i thi" & ChrW(7875) & "u"

   DinhMucYeuCau = ChrW(272) & ChrW(7883) & "nh m" & ChrW(7913) & _
   "c yêu c" & ChrW(7847) & "u"

   PhuongThucTinh = "Ph" & ChrW(432) & ChrW(417) & "ng th" & ChrW( _
   7913) & "c tính"

   HeSo = "H" & ChrW(7879) & " s" & ChrW(7889)

   GhiChu = "Ghi chú"

   DaLap = ChrW(272) & "ã l" & ChrW(7853) & "p"

   Dim i As Integer
   With TableCongViecTrachNhiemTheoViTri.ColumnHeaders
      For i = .Count To 1 Step -1
         .Remove i
      Next i
   End With

   With TableCongViecTrachNhiemTheoViTri.ColumnHeaders
      .Add , , UniToWindows1258(TenCongViec), 100
      .Add , , UniToWindows1258(DonViLuongHoa), 100
      .Add , , UniToWindows1258(DinhMucToiThieu), 100
      .Add , , UniToWindows1258(DinhMucYeuCau), 100
      If ValueNam Then
         .Add , , UniToWindows1258(DaLap), 100
      Else
         .Add , , UniToWindows1258(DaLap), 0
      End If
      .Add , , UniToWindows1258(PhuongThucTinh), 100
      .Add , , UniToWindows1258(HeSo), 70
      .Add , , UniToWindows1258(GhiChu), 100
      .Add , , "CongViecID", 120
      .Add , , "NhanVienID", 120
      .Add , , "ViTriID", 120
      .Add , , "TrachNhiemTheoViTriID", 120
      .Add , , "Thang", 120
      .Add , , "CongThucTinh", 120
   End With
End Function

Function F_KiemTraDaChonThangChua() As Boolean

   If Thang = 0 And Nam = False Then
      NoiDung = "Ch" & ChrW(7885) & "n ít nh" & ChrW(7845) & _
      "t 1 tháng tr" & ChrW(432) & ChrW(7899) & "c khi tich ch" & ChrW(7885) & _
      "n"
      Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
      F_KiemTraDaChonThangChua = True
    Exit Function
   Else
      F_KiemTraDaChonThangChua = False
   End If
End Function

Function F_KiemTraChonPB_NV() As Boolean
   TieuDe = "BOS xin thông báo"

   If PhongBanID_Global = 0 Then

      NoiDung = "Ch" & ChrW(7885) & "n phòng ban tr" & ChrW(432) _
      & ChrW(7899) & "c khi ch" & ChrW(7885) & "n tháng"
      Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
      F_KiemTraChonPB_NV = True
    Exit Function
   End If

   If NhanVienID_Global = 0 Then
      NoiDung = "Ch" & ChrW(7885) & "n nhân viên "
      Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
      F_KiemTraChonPB_NV = True
    Exit Function
   End If

   F_KiemTraChonPB_NV = False
End Function

Function F_KiemTraTruocKhiCapNhat(DinhMucToiThieu As Currency, DinhMucYeuCau As Currency, HeSo As Currency, PhuongThucTinh As Variant) As Boolean
   If DinhMucToiThieu = 0 Then

      NoiDung = "Nh" & ChrW(7853) & "p " & ChrW(273) & ChrW(7883) _
      & "nh m" & ChrW(7913) & "c t" & ChrW(7889) & "i thi" & ChrW(7875) & "u"
      Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

      F_KiemTraTruocKhiCapNhat = True
      InputDinhMucToiThieu.BorderColor = colorError
    Exit Function
   Else
      InputDinhMucToiThieu.BorderColor = colorDefault
   End If

   If DinhMucYeuCau = 0 Then
      NoiDung = "Nh" & ChrW(7853) & "p " & ChrW(273) & ChrW(7883) _
      & "nh m" & ChrW(7913) & "c yêu c" & ChrW(7847) & "u"
      Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

      F_KiemTraTruocKhiCapNhat = True
      InputDinhMucYeuCau.BorderColor = colorError
    Exit Function
   Else
      InputDinhMucYeuCau.BorderColor = colorDefault
   End If

   If HeSo = 0 Then
      NoiDung = "Nh" & ChrW(7853) & "p h" & ChrW(7879) & " s" & _
      ChrW(7889)
      Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

      F_KiemTraTruocKhiCapNhat = True
      InputHeSo.BorderColor = colorError
    Exit Function
   Else
      InputHeSo.BorderColor = colorDefault
   End If

   If IsNull(PhuongThucTinh) Then
      NoiDung = "Ch" & ChrW(7885) & "n ph" & ChrW(432) & ChrW(417 _
      ) & "ng th" & ChrW(7913) & "c tính"
      Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

      F_KiemTraTruocKhiCapNhat = True
      SelectPhuongThucTinh.BorderColor = colorError
    Exit Function
   Else
      SelectPhuongThucTinh.BorderColor = colorDefault
   End If
   F_KiemTraTruocKhiCapNhat = False

End Function

' Xoa du lieu
Function F_XoaDuLieuKhiChonThang()
   Dim i
   With TableCongViecTrachNhiemTheoViTri

      For i = .ListItems.Count To 1 Step -1
         .ListItems.Remove i
      Next i

      Set addedItemsCongViecID = New Collection

   End With

   Thang = 0
   ValueNam = False
   IndexTrachNhiemViTri = 0
   TrachNhiemTheoViTriID = 0
   SelectPhuongThucTinh.Clear

   For i = 1 To 12
      With Form_CongViecTrachNhiemNhanVien.Controls("btnT" & i)
         .BackColor = bgDefault
      End With
   Next i

   Call F_ResetFrom
End Function

'Kiem Tra
Function F_KiemTraChonNamChua(Nam As String) As Boolean
   If Nam = "" Then
      colorError = &HFF&
      colorDefault = &H80000000

      NoiDung = "Ch" & ChrW(7885) & "n n" & ChrW(259) & "m"
      Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
      SelectNam.BorderColor = colorError
      F_KiemTraChonNamChua = True
    Exit Function
   Else
      SelectNam.BorderColor = colorDefault
   End If

   F_KiemTraChonNamChua = False
End Function

Function F_Clear()
   Dim i As Integer
   With TableCongViecTrachNhiemTheoViTri
      For i = .ListItems.Count To 1 Step -1
         .ListItems.Remove i
      Next i
      Set addedItemsCongViecID = New Collection
   End With

   With TableDanhSachNhanVienViTri

      For i = .ListItems.Count To 1 Step -1
         .ListItems.Remove i
      Next i

      Set addedItemsNhanVienID = New Collection

   End With

   Thang = 0
   ValueNam = False
   IndexTrachNhiemViTri = 0
   TrachNhiemTheoViTriID = 0
   IndexNhanvienViTri = 0
   SelectPhuongThucTinh = ""
   NhanVienID_Global = 0
   PhongBanID_Global = 0

   For i = 1 To 12
      With Form_CongViecTrachNhiemNhanVien.Controls("btnT" & i)
         .BackColor = bgDefault
      End With
   Next i

   Call F_ResetFrom

End Function

Function F_ResetFrom()
   InputTenCongViec.value = ""
   InputDonViLuongHoa.value = ""
   InputDinhMucToiThieu.value = ""
   InputDinhMucYeuCau.value = ""
   SelectPhuongThucTinh = ""
   InputHeSo.value = ""
   InputGhiChu.value = ""
   IndexTrachNhiemViTri = 0
   TrachNhiemTheoViTriID = 0
End Function

Function F_DanhSachNhanVienTheoPhongBan(PhongBanID As Integer)
   On Error Resume Next
   If PhongBanID = 0 Then
    Exit Function
   End If

   Call F_Clear

   PhongBanID_Global = PhongBanID
   Dim Query As String
   Dim dbConn As Object

   Query = "Select TenPhongBan, Ho + ' ' + Ten As TenNhanVien, NhanVienID, NS_NhanVien.ViTriID, TenViTri FROM PhongBan " & _
   "JOIN NS_NhanVien on NS_NhanVien.PhongBanID = PhongBan.PhongBanID " & _
   "JOIN DM_ViTri VT on VT.ViTriID = NS_NhanVien.ViTriID " & _
   "WHERE PhongBan.PhongBanID in(Select PhongBanID FROM DM_ViTri WHERE PhongBanID in (Select valu from LayDonViCon(" & PhongBanID & "))) And VT.ViTriID in(Select ViTriID   FROM CV_PhongBan_CongViec PB_CV  inner join  CV_CongViec CV on PB_CV.CongViecID= CV.CongViecID  WHERE CV.CongViecChinhID = 1 And CV.Nam = " & SelectNam.value & ")"
   Set dbConn = ConnectToDatabase

   Call F_ViewTableDanhSachNhanvien(Query, dbConn)
   Call CloseDatabaseConnection(dbConn)
End Function

Function F_ViewTableDanhSachNhanvien(Query, dbConn)
   On Error Resume Next
   If Not dbConn Is Nothing Then
      Dim Rs As Object
      Dim i As Integer
      Set Rs = dbConn.Execute(Query)
      With TableDanhSachNhanVienViTri

         For i = .ListItems.Count To 1 Step -1
            .ListItems.Remove i
         Next i
         Set addedItemsMucTieuCv = New Collection
      End With

      Do Until Rs.EOF
         Dim foundItem As Boolean
         foundItem = False
         Dim existingItem As Variant

         For Each existingItem In addedItemsNhanVienID
            If existingItem = Rs.Fields("NhanVienID").value Then
               foundItem = True
             Exit For
            End If
         Next existingItem

         If Not foundItem Then
            Dim ListItem As ListItem
            Set ListItem = TableDanhSachNhanVienViTri.ListItems.Add(, , UniToWindows1258(Rs.Fields("TenPhongBan").value))
            ListItem.SubItems(1) = UniToWindows1258(Rs.Fields("TenNhanVien").value)
            ListItem.SubItems(2) = Rs.Fields("NhanVienID").value
            ListItem.SubItems(3) = Rs.Fields("ViTriID").value
            ListItem.SubItems(4) = UniToWindows1258(Rs.Fields("TenViTri").value)
            addedItemsNhanVienID.Add Rs.Fields("NhanVienID").value
         End If
         Rs.MoveNext
      Loop
   Else
      MsgBox "Mat Ket noi csdl"
   End If
End Function

Function ThangDaLapKeHoach(index As Integer)
   On Error Resume Next
   Dim i As Integer

   Call F_XoaThang
   With TableCongViecTrachNhiemTheoViTri
      For i = .ListItems.Count To 1 Step -1
         .ListItems.Remove i
      Next i
      Set addedItemsCongViecID = New Collection
   End With
   Call F_ResetFrom

   Dim ViTriID
   Dim NhanVienID
   IndexNhanvienViTri = index
   With TableDanhSachNhanVienViTri
      NhanVienID = .ListItems(index).SubItems(2)

      ViTriID = .ListItems(index).SubItems(3)
      NhanVienID_Global = NhanVienID
      LabeViTri = FontConverter(.ListItems(index).SubItems(4), 2, 1)
      LabelNhanVien = FontConverter(.ListItems(index).SubItems(1), 2, 1)
   End With

   If ViTriID = 0 Or NhanVienID = 0 Then
    Exit Function
   End If

   Dim Query As String
   Dim dbConn As Object
   Dim Rs As Object
   Dim Months As Variant

   Query = "Select DISTINCT Thang From CV_TrachNhiem_TheoViTri TTVT " & _
   "WHERE TTVT.NhanVienID = " & NhanVienID

   Set dbConn = ConnectToDatabase

   If Not dbConn Is Nothing Then
      Set Rs = dbConn.Execute(Query)

      If Not Rs.EOF And Not Rs.BOF Then
         Months = Rs.GetRows()

         For i = LBound(Months, 2) + 1 To UBound(Months, 2)
            With Form_CongViecTrachNhiemNhanVien.Controls("btnT" & i)
               .BackColor = bgCheck
            End With
         Next i
      End If
   Else
      MsgBox "Mat Ket noi csdl"
   End If
   Call CloseDatabaseConnection(dbConn)
End Function

Function F_CapNhatThietLapLanDau()
   If F_Count > 0 Then
    Exit Function
   End If

   Dim dbConn As Object
   Dim Query As String
   Dim Rs As Object
   Dim Item As MSComctlLib.ListItem
   Dim i As Integer

   Set dbConn = ConnectToDatabase

   For Each Item In TableCongViecTrachNhiemTheoViTri.ListItems
      If Item.ListSubItems(8) > 0 And Item.Checked Then
         For i = 0 To 12
            Query = "INSERT INTO CV_TrachNhiem_TheoViTri(CongViecID,NhanVienID,Thang) " & _
            "VALUES(" & Item.ListSubItems(8) & ", " & NhanVienID_Global & ", " & i & ")"

            Set Rs = dbConn.Execute(Query)
         Next i

      End If
   Next Item
   Call CloseDatabaseConnection(dbConn)
   Set Rs = Nothing
End Function

Function F_DanhCvTracNhiemTheoVt()
   If IndexNhanvienViTri = 0 Then
    Exit Function
   End If

   Dim ViTriID
   Dim NhanVienID

   With TableDanhSachNhanVienViTri
      NhanVienID = .ListItems(IndexNhanvienViTri).SubItems(2)
      ViTriID = .ListItems(IndexNhanvienViTri).SubItems(3)
   End With

   If ViTriID = 0 Or NhanVienID = 0 Then
    Exit Function
   End If
   Dim Query As String
   Dim dbConn As Object

   Query = "Select " & _
   "TenCongViec, " & _
   "isNUll(CachLuongHoa,'') As CachLuongHoa, " & _
   "ISNULL(TNTV.DinhMucToiThieu,0) As DinhMucToiThieu, " & _
   "Isnull(TNTV.DinhMucYeuCau,0) As DinhMucYeuCau," & _
   "ISNULL(Case " & _
   "WHEN CV.CongThucTinhID = 1 Then (Select Sum(DinhMucYeuCau) from CV_TrachNhiem_TheoViTri  GV where Nam = " & SelectNam.value & " And Thang > 0 And GV.CongViecID = CV.CongViecID And Gv.NhanvienID = NV.NhanVienID) " & _
   "WHEN CV.CongThucTinhID = 2 Then (Select avg(DinhMucYeuCau) from CV_TrachNhiem_TheoViTri  GV where Nam = " & SelectNam.value & " And Thang > 0 And GV.CongViecID = CV.CongViecID And Gv.NhanvienID = NV.NhanVienID) " & _
   "WHEN CV.CongThucTinhID = 3 Then (Select max(DinhMucYeuCau) from CV_TrachNhiem_TheoViTri  GV where Nam = " & SelectNam.value & " And Thang > 0 And GV.CongViecID = CV.CongViecID And Gv.NhanvienID = NV.NhanVienID) " & _
   "WHEN CV.CongThucTinhID = 4 Then (Select min(DinhMucYeuCau) from CV_TrachNhiem_TheoViTri  GV where Nam = " & SelectNam.value & " And Thang > 0 And GV.CongViecID = CV.CongViecID And Gv.NhanvienID = NV.NhanVienID) " & _
   "END,0) As DaLap, " & _
   "ISNULL((Select TOP 1 TenPhuongThucTinh   from CV_TrachNhiem_TheoViTri PTT1 INNER JOIN DM_PhuongThucTinh PTT on PTT1.PhuongThucTinhID = PTT.PhuongThucTinhID  where  CongViecID = CV.CongViecID And NhanVienID = " & NhanVienID & " And Thang = " & Thang & "  ),'')  As PhuongThucTinh,  " & _
   "ISNULL( (Select TOP 1 HeSo   from CV_TrachNhiem_TheoViTri where  CongViecID = CV.CongViecID And NhanVienID =" & NhanVienID & " And Thang = " & Thang & " And Nam = " & SelectNam.value & "),0) As HeSo, " & _
   "ISNULL(TNTV.GhiChu, '') As GhiChu,  " & _
   "CV.CongViecID, NV.NhanVienID, " & _
   "NV.ViTriID, " & _
   "ISNULL((Select TOP 1 TrachNhiemTheoViTriID  from CV_TrachNhiem_TheoViTri where CongViecID = CV.CongViecID And NhanVienID = " & NhanVienID & " And Thang = " & Thang & " And Nam = " & SelectNam.value & " ),0)  TrachNhiemTheoViTriID," & Thang & "  As Thang, isnull(CV.CongThucTinhID,0) As CongThucTinhID " & _
   "FROM CV_PhongBan_CongViec CVPBCV " & _
   "JOIN NS_NhanVien NV on NV.ViTriID = CVPBCV.ViTriID JOIN CV_CongViec CV on CV.CongViecID = CVPBCV.CongViecID  " & _
   "LEFT JOIN (Select * FROM CV_TrachNhiem_TheoViTri WHERE Nam = " & SelectNam.value & " And Thang = " & Thang & ") TNTV on TNTV.CongViecID = CV.CongViecID And TNTV.NhanVienID = NV.NhanVienID " & _
   "WHERE NV.ViTriID = " & ViTriID & " And NV.NhanVienID = " & NhanVienID & " And Cv.CongViecChinhID = 1"

   Set dbConn = ConnectToDatabase
   Call F_ViewTableCvTracNhiemTheoVt(Query, dbConn)

   Call CloseDatabaseConnection(dbConn)
End Function

Function F_ViewTableCvTracNhiemTheoVt(Query, dbConn)
   On Error Resume Next
   If Not dbConn Is Nothing Then
      Dim Rs As Object
      Dim i As Integer
      Set Rs = dbConn.Execute(Query)

      With TableCongViecTrachNhiemTheoViTri

         For i = .ListItems.Count To 1 Step -1
            .ListItems.Remove i
         Next i
         Set addedItemsCongViecID = New Collection
      End With

      If Not Rs.EOF And Not Rs.BOF Then
         Do Until Rs.EOF
            Dim foundItem As Boolean
            foundItem = False
            Dim existingItem As Variant

            For Each existingItem In addedItemsCongViecID
               If existingItem = Rs.Fields("CongViecID").value Then
                  foundItem = True
                Exit For
               End If
            Next existingItem

            If Not foundItem Then
               Dim ListItem As ListItem
               Set ListItem = TableCongViecTrachNhiemTheoViTri.ListItems.Add(, , Trim(UniToWindows1258(Rs.Fields("TenCongViec").value)))
               ListItem.SubItems(1) = Trim(UniToWindows1258(Rs.Fields("CachLuongHoa").value))
               ListItem.SubItems(2) = UniToWindows1258(Rs.Fields("DinhMucToiThieu").value)
               ListItem.SubItems(3) = UniToWindows1258(Rs.Fields("DinhMucYeuCau").value)
               ListItem.SubItems(4) = UniToWindows1258(Rs.Fields("DaLap").value)
               ListItem.SubItems(5) = Trim(UniToWindows1258(Rs.Fields("PhuongThucTinh").value))
               ListItem.SubItems(6) = UniToWindows1258(Rs.Fields("HeSo").value)
               ListItem.SubItems(7) = Trim(UniToWindows1258(Rs.Fields("GhiChu").value))
               ListItem.SubItems(8) = Rs.Fields("CongViecID").value
               ListItem.SubItems(9) = Rs.Fields("NhanVienID").value
               ListItem.SubItems(10) = Rs.Fields("ViTriID").value
               ListItem.SubItems(11) = Rs.Fields("TrachNhiemTheoViTriID").value
               ListItem.SubItems(12) = Rs.Fields("Thang").value
               ListItem.SubItems(13) = Rs.Fields("CongThucTinhID").value
               addedItemsCongViecID.Add Rs.Fields("CongViecID").value

               If F_Count > 0 Then
                  If Rs.Fields("TrachNhiemTheoViTriID").value > 0 Then
                     ListItem.Checked = True
                  End If
               Else
                  ListItem.Checked = True
               End If
            End If
            Rs.MoveNext
         Loop
      End If
   Else
      MsgBox "Mat Ket noi csdl"
   End If
End Function

Function F_Count() As Integer
   Dim dbConn As Object
   Dim Query As String
   Dim Rs As Object
   Dim CountCV As Variant

   Set dbConn = ConnectToDatabase

   Query = "Select Count(TrachNhiemTheoViTriID) As count FROM CV_TrachNhiem_TheoViTri WHERE NhanVienID = " & NhanVienID_Global

   Set Rs = dbConn.Execute(Query)


   CountCV = Rs.GetRows()

   F_Count = CountCV(0, 0)
   Call CloseDatabaseConnection(dbConn)
   Set Rs = Nothing
End Function

Function F_DanhSachPhuongThucTinh()
   Dim Query As String
   Dim dbConn As Object

   Query = "Select PhuongThucTinhID, TenPhuongThucTinh From DM_PhuongThucTinh"
   Set dbConn = ConnectToDatabase

   Call ViewListBox(Query, SelectPhuongThucTinh, dbConn)
   Call CloseDatabaseConnection(dbConn)

End Function

Function F_LayDuLieuXuongFromTrachNhiemNV(index As Integer)
   On Error Resume Next
   Call F_ResetFrom
   IndexTrachNhiemViTri = index

   Call F_DanhSachPhuongThucTinh

   With TableCongViecTrachNhiemTheoViTri
      Call F_ResetFrom

      InputTenCongViec.value = FontConverter(.ListItems(index), 2, 1)
      InputDonViLuongHoa.value = FontConverter(.ListItems(index).SubItems(1), 2, 1)
      InputDinhMucToiThieu.value = FontConverter(.ListItems(index).SubItems(2), 2, 1)
      InputDinhMucYeuCau.value = FontConverter(.ListItems(index).SubItems(3), 2, 1)
      SelectPhuongThucTinh.Text = FontConverter(.ListItems(index).SubItems(5), 2, 1)
      InputHeSo.value = FontConverter(.ListItems(index).SubItems(6), 2, 1)
      InputGhiChu.value = FontConverter(.ListItems(index).SubItems(7), 2, 1)
      TrachNhiemTheoViTriID = .ListItems(index).SubItems(11)
      CongThucTinhID = .ListItems(index).SubItems(13)

      NhanVienID_Global = .ListItems(index).SubItems(9)
      CongViecID_Global = .ListItems(index).SubItems(8)

   End With

   FrameTTCT.Enabled = True
End Function

Function F_CheckCvTrachNhiemNv(index As Integer)
   On Error Resume Next
   Dim NhanVienID As Integer
   Dim ViTriID As Integer
   Dim CongViecID As Integer
   Dim NamLap As String

   NamLap = SelectNam.value
   With TableCongViecTrachNhiemTheoViTri
      CongViecID = .ListItems(index).SubItems(8)
      NhanVienID = .ListItems(index).SubItems(9)
      ViTriID = .ListItems(index).SubItems(10)
   End With

   If F_KiemTraChonNamChua(NamLap) Then
      TableCongViecTrachNhiemTheoViTri.ListItems(index).Checked = False
    Exit Function

   End If

   If NhanVienID = 0 And ViTriID = 0 And CongViecID And Thang = 0 Then
    Exit Function
   End If

   CongViecID_Global = CongViecID

   Dim Query As String
   Dim dbConn As Object
   Dim Rs As Object

   Set dbConn = ConnectToDatabase

   If ValueNam = True Then
      Query = "INSERT INTO CV_TrachNhiem_TheoViTri(CongViecID,NhanVienID,Thang, Nam) VALUES(" & CongViecID & ", " & NhanVienID & ",0 ," & NamLap & ")"
   Else
      Query = "INSERT INTO CV_TrachNhiem_TheoViTri(CongViecID,NhanVienID,Thang, Nam) VALUES(" & CongViecID & ", " & NhanVienID & ", " & Thang & "," & NamLap & ")"
   End If

   Set Rs = dbConn.Execute(Query)
   Set Rs = Nothing

   Call F_DanhCvTracNhiemTheoVt

   Call CloseDatabaseConnection(dbConn)
End Function

Function F_SetBtnThang(i As Integer)
   Call F_CapNhatThietLapLanDau
   Call ThangDaLapKeHoach(IndexThang)

   If F_KiemTraChonPB_NV Then
    Exit Function
   End If

   ValueNam = False
   Nam.BackColor = bgDefault
   '    F_XoaDuLieuKhiChonThang

   With Form_CongViecTrachNhiemNhanVien.Controls("btnT" & i)
      .BackColor = bgActive
      Thang = i
   End With

   Dim j As Integer

   For j = 1 To 12
      If j <> i Then
         With Form_CongViecTrachNhiemNhanVien.Controls("btnT" & j)
            If .BackColor <> bgCheck Then
               .BackColor = bgDefault
            End If

         End With
      End If
   Next j

   ' Chi hien lai khi la cap nhat, them moi khong hien

   Call F_HeaderTableCongViecTrachNhiemTheoViTri
   Call F_DanhCvTracNhiemTheoVt

   ValueNam = False
End Function

Function F_btnTatCa()

   ValueNam = True
   Nam.BackColor = bgActive

   Dim i As Integer
   For i = 1 To 12
      With Form_CongViecTrachNhiemNhanVien.Controls("btnT" & i)
         .BackColor = bgDefault
      End With
   Next i
   Call ThangDaLapKeHoach(IndexNhanvienViTri)
   Call F_HeaderTableCongViecTrachNhiemTheoViTri
   Thang = 0
   Call F_DanhCvTracNhiemTheoVt
End Function

Function F_FormCapNhatThongTinChiTiet()

   Call F_CapNhatThietLapLanDau

   Dim DinhMucToiThieu As Currency
   Dim DinhMucYeuCau As Currency

   Dim DinhMucToiThieuCT As Currency
   Dim DinhMucYeuCauCT As Currency

   Dim PhuongThucTinhID As Variant
   Dim HeSo As Currency
   Dim GhiChu As String
   Dim NamLap As Integer

   NamLap = SelectNam.value

   DinhMucToiThieu = InputDinhMucToiThieu.value
   DinhMucYeuCau = InputDinhMucYeuCau.value

   DinhMucToiThieuCT = F_TinhDinhMucTheoDieuKien(InputDinhMucToiThieu.value)
   DinhMucYeuCauCT = F_TinhDinhMucTheoDieuKien(InputDinhMucYeuCau.value)
   HeSo = InputHeSo.value
   GhiChu = InputGhiChu.value
   PhuongThucTinhID = SelectPhuongThucTinh.value

   If F_KiemTraTruocKhiCapNhat(DinhMucToiThieu, DinhMucYeuCau, HeSo, PhuongThucTinhID) Then
    Exit Function
   End If

   Dim Query As String
   Dim dbConn As Object
   Dim Rs As Object
   Set dbConn = ConnectToDatabase

   If ValueNam = True Then
      Dim i As Integer
      Query = "DELETE FROM CV_TrachNhiem_TheoViTri WHERE TrachNhiemTheoViTriID =" & TrachNhiemTheoViTriID
      On Error Resume Next
      Set Rs = dbConn.Execute(Query)
      Set Rs = Nothing

      For i = 0 To 12

         ' i cap nhat du lieu ca nam
         If i = 0 Then
            Query = "INSERT INTO CV_TrachNhiem_TheoViTri(CongViecID,NhanVienID,DinhMucToiThieu,DinhMucYeuCau,HeSo,GhiChu,Thang, Nam,PhuongThucTinhID) " & _
            "VALUES(" & CongViecID_Global & ", " & NhanVienID_Global & "," & DinhMucToiThieu & ", " & DinhMucYeuCau & ", " & HeSo & ", N'" & GhiChu & "', " & i & " ," & NamLap & ", " & PhuongThucTinhID & ")"
         Else
            ' Du lieu cap nhat cho 12 thang giua vao cong thu tinh F_TinhDinhMucTheoDieuKien
            Query = "INSERT INTO CV_TrachNhiem_TheoViTri(CongViecID,NhanVienID,DinhMucToiThieu,DinhMucYeuCau,HeSo,GhiChu,Thang, Nam,PhuongThucTinhID) " & _
            "VALUES(" & CongViecID_Global & ", " & NhanVienID_Global & "," & DinhMucToiThieuCT & ", " & DinhMucYeuCauCT & ", " & HeSo & ", N'" & GhiChu & "', " & i & " ," & NamLap & ", " & PhuongThucTinhID & ")"
         End If

         Set Rs = dbConn.Execute(Query)
         Set Rs = Nothing
      Next i
   Else


      Query = "UPDATE CV_TrachNhiem_TheoViTri Set DinhMucToiThieu = " & DinhMucToiThieu & "," & _
      "DinhMucYeuCau = '" & DinhMucYeuCau & "', " & _
      "HeSo = '" & HeSo & "', GhiChu = N'" & GhiChu & "', " & _
      "PhuongThucTinhID = " & PhuongThucTinhID & _
      " WHERE  TrachNhiemTheoViTriID = " & TrachNhiemTheoViTriID

      Set Rs = dbConn.Execute(Query)
      Set Rs = Nothing

      CreateObject("WScript.Shell").Popup "Th" & ChrW(225) & "ng 12 s" & ChrW(7869) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c c" & ChrW(7853) & "p nh" & ChrW(7853) & "t l" & ChrW(7841) & "i ", , "BOS Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 0

      Query = "Update CV_TrachNhiem_TheoViTri Set DinhMucYeuCau = " & _
      "(Select top 1 DinhMucYeuCau from CV_TrachNhiem_TheoViTri KHN WHERE Nam = CV_TrachNhiem_TheoViTri.Nam And CongViecID = CV_TrachNhiem_TheoViTri.CongViecID And NhanVienID = CV_TrachNhiem_TheoViTri.NhanVienID  And Thang = 0) - " & _
      "(Select Sum(DinhMucYeuCau) from CV_TrachNhiem_TheoViTri KHT WHERE Nam = CV_TrachNhiem_TheoViTri.Nam And CongViecID = CV_TrachNhiem_TheoViTri.CongViecID And NhanVienID = CV_TrachNhiem_TheoViTri.NhanVienID And Thang between 1 And 11) " & _
      "WHERE Nam = " & NamLap & " And CongViecID = " & CongViecID_Global & " And NhanVienID = " & NhanVienID_Global & " And Thang = 12"

      Set Rs = dbConn.Execute(Query)
      Set Rs = Nothing
   End If

   Call CloseDatabaseConnection(dbConn)
   Call F_DanhCvTracNhiemTheoVt
   Call ThangDaLapKeHoach(IndexNhanvienViTri)
   ThongBao_ThanhCong
End Function

' Tinh dinh muc giua vao cong thuc tinh o phan cong viec
Function F_TinhDinhMucTheoDieuKien(value As Currency) As Currency
   If CongThucTinhID = 1 Then
      F_TinhDinhMucTheoDieuKien = value / 12
   Elseif CongThucTinhID = 2 Or CongThucTinhID = 3 Or CongThucTinhID = 4 Then
      F_TinhDinhMucTheoDieuKien = value
   End If
End Function

Function F_XoaThang()
   Dim i As Integer

   For i = 1 To 12
      With Form_CongViecTrachNhiemNhanVien.Controls("btnT" & i)
         .BackColor = bgDefault
      End With
   Next i
End Function

Function F_XoaCongViecTN()
   On Error Resume Next
   If TrachNhiemTheoViTriID = 0 Then
    Exit Function
   End If

   Dim msgValue As VbMsgBoxResult
   msgValue = Application.Assistant.DoAlert(UniConvert("Carnh baso", "Telex"), UniConvert("Bajn muoosn xosa duwx lieeju?", "Telex"), msoAlertButtonYesNo, msoAlertIconWarning, 0, 0, 0)

   If msgValue = vbYes Then
      Dim dbConn As Object
      Dim Rs As Object
      Dim Query As String
      Set dbConn = ConnectToDatabase

      If ValueNam Then
         Dim i As Integer

         For i = 1 To 12
            Query = "DELETE FROM CV_TrachNhiem_TheoViTri  WHERE NhanVienID = " & NhanVienID_Global & " And CongViecID = " & CongViecID_Global
            Set Rs = dbConn.Execute(Query)
         Next i

      Else
         Query = "DELETE FROM CV_TrachNhiem_TheoViTri WHERE TrachNhiemTheoViTriID = " & TrachNhiemTheoViTriID
         Set Rs = dbConn.Execute(Query)
      End If

      Set Rs = Nothing
      Call CloseDatabaseConnection(dbConn)
      Call F_ResetFrom
      Call F_DanhCvTracNhiemTheoVt

      ThongBao_ThanhCong
   End If

End Function

Function F_BoCheckTranhNhiemNv(index As Integer)
   With TableCongViecTrachNhiemTheoViTri
      CongViecID_Global = .ListItems(index).SubItems(8)
      NhanVienID_Global = .ListItems(index).SubItems(9)
      TrachNhiemTheoViTriID = .ListItems(index).SubItems(11)
   End With

   If TrachNhiemTheoViTriID = 0 Then
    Exit Function
   End If

   Call F_XoaCongViecTN

End Function

Sub HienThiFromCongViecTrachNhiemNhanVien()
   Form_CongViecTrachNhiemNhanVien.Show
End Sub

Function F_DongBoVoiKheHoachNV()
   If NhanVienID_Global < 0 Then
    Exit Function
   End If

   Dim dbConn As Object
   Dim Rs As Object
   Dim KD As Variant
   Dim Query As String

   Set dbConn = ConnectToDatabase

   Query = "Select isNull(KeHoachDoanhThuNv,0) As KeHoachDoanhThuNv FROM KeHoachDoanhThuNv WHERE PhongBanID = " & SelectPhongBan.value & " And Nam = " & SelectNam.value & " And NhanVienID = " & NhanVienID_Global


   Set Rs = dbConn.Execute(Query)

   If Rs.EOF And Rs.BOF Then
    Exit Function
   Else
      KD = Rs.GetRows()

      If IsArrayEmpty(KD) Or UBound(KD, 2) < 1 Then
       Exit Function
      Else
         InputDinhMucYeuCau.value = KD(0, 0)
      End If
   End If

   Call CloseDatabaseConnection(dbConn)
End Function

Function IsArrayEmpty(arr As Variant) As Boolean
   On Error Resume Next
   IsArrayEmpty = (LBound(arr, 1) > UBound(arr, 1)) Or (LBound(arr, 2) > UBound(arr, 2))
   On Error Goto 0
End Function





