
Private addedItems As Collection
Private IndexChiTieuNVDL
Private ChiTieuNhiemVuID As Integer
Private bgDefault As String
Private bgActive As String
Private bgCheck As String
Private ValueThang As Integer
Private ValueNam As Boolean
Private colorError, colorDefault As String
Private NhiemVuID As Integer
Private CongThucTinh As Integer
Public TienThang1 As Currency
Public TienThang2 As Currency
Public TienThang3 As Currency
Public TienThang4 As Currency
Public TienThang5 As Currency
Public TienThang6 As Currency
Public TienThang7 As Currency
Public TienThang8 As Currency
Public TienThang9 As Currency
Public TienThang10 As Currency
Public TienThang11 As Currency
Public TienThang12 As Currency
Public DongBo As Boolean

Dim NoiDung As String
Dim TieuDe As String

Private Sub btnHuy_Click()
    Unload Me
End Sub

Private Sub btnXoa_Click()
    Call F_XoaChiTieuNhiemVuDinhLuong
End Sub

Private Sub cbbDongBoVoiKHDV_Click()
    DongBo = True
    Dim dbConn As Object
    Dim Rs As Object
    Dim Rs2 As Object
    Dim KD As Variant
    Dim KD2 As Variant
    Dim Query As String
    Dim Query2 As String
    Dim i As Integer

    Set dbConn = ConnectToDatabase

    If ValueNam Then
        Query = "Select isNull(KeHoachDoanhThu,0) As KeHoachDoanhThu FROM KeHoachDoanhThu WHERE PhongBanID = " & cbbPhongBan.value & " And Nam = " & cbbNam.value

        For i = 1 To 12
            Query2 = "Select isNull(TienThang" & i & ",0) As KeHoachPhanBoDV FROM KeHoachPhanBoDV WHERE PhongBanID = " & cbbPhongBan.value & " And Nam = " & cbbNam.value & ""

            Set Rs2 = dbConn.Execute(Query2)
            If Rs2.EOF And Rs2.BOF Then
             Exit Sub
            Else
                KD2 = Rs2.GetRows()

                If i = 1 Then
                    TiengThang1 = KD2(0, 0)
                End If

                If i = 2 Then
                    TiengThang2 = KD2(0, 0)
                End If

                If i = 3 Then
                    TiengThang3 = KD2(0, 0)
                End If

                If i = 4 Then
                    TiengThang4 = KD2(0, 0)
                End If

                If i = 5 Then
                    TiengThang5 = KD2(0, 0)
                End If

                If i = 6 Then
                    TiengThang6 = KD2(0, 0)
                End If

                If i = 7 Then
                    TiengThang7 = KD2(0, 0)
                End If

                If i = 8 Then
                    TiengThang8 = KD2(0, 0)
                End If

                If i = 9 Then
                    TiengThang9 = KD2(0, 0)
                End If

                If i = 10 Then
                    TiengThang10 = KD2(0, 0)
                End If

                If i = 11 Then
                    TiengThang11 = KD2(0, 0)
                End If

                If i = 12 Then
                    TiengThang12 = KD2(0, 0)
                End If
            End If
        Next i
    Else
        Query = "Select isNull(TienThang" & ValueThang & ",0) As KeHoachPhanBoDV FROM KeHoachPhanBoDV WHERE PhongBanID = " & cbbPhongBan.value & " And Nam = " & cbbNam.value & ""
    End If

    Set Rs = dbConn.Execute(Query)

    If Rs.EOF And Rs.BOF Then
     Exit Sub
    Else
        KD = Rs.GetRows()
        If IsArrayEmpty(KD) Then
         Exit Sub
        Else
            txtDinhMucYeuCau.value = KD(0, 0)
        End If
    End If

    Call CloseDatabaseConnection(dbConn)

End Sub

Function IsArrayEmpty(arr As Variant) As Boolean
    On Error Resume Next
    IsArrayEmpty = (LBound(arr, 1) > UBound(arr, 1)) Or (LBound(arr, 2) > UBound(arr, 2))
    On Error Goto 0
End Function

Private Sub txtDinhMucToiThieu_Change()
    If Not IsNumeric(txtDinhMucToiThieu) Then
        txtDinhMucToiThieu.value = ""
    End If

    txtDinhMucToiThieu = txtDinhMucToiThieu.value

End Sub

Private Sub txtDinhMucYeuCau_Change()
    If Not IsNumeric(txtDinhMucYeuCau) Then
        txtDinhMucYeuCau.value = ""
    End If

    txtDinhMucYeuCau = txtDinhMucYeuCau.value
End Sub

Private Sub txtTrongSo_Change()
    If Not IsNumeric(txtTrongSo) Then
        txtTrongSo.value = ""
    End If

    txtTrongSo = txtTrongSo.value

End Sub

Private Sub UserForm_Initialize()
    TieuDe = "Bos Xin thông báo"
    bgDefault = &H80000005
    bgActive = &H80FF&
    bgCheck = &HC000&
    colorError = &HFF&
    colorDefault = &H80000000
    Call F_KhoiTaoNam(cbbNam)
    cbbNam.value = Sheet2.cbbNamSheetCongViec.value
    Call F_TableChiTieuNhiemVuDinh
    Set addedItems = New Collection

    Call F_KhoiTaoPhongBan
    Call F_DonViTinh
    lvNhiemVuCuaPhongBan.Font.Charset = VIETNAMESE_CHARSET
    Call F_PhuongThucTinh

    F_ToMauItem
End Sub

Function F_DonViTinh()
    Dim Query As String
    Dim dbConn As Object

    Set dbConn = ConnectToDatabase

    Query = "Select isNULL(DonViTinhID, 0), isNull(TenDonViTinh,'') FROM DM_DonViTinh"

    Call ViewListBox(Query, cbbDonViTinh, dbConn)

    Call CloseDatabaseConnection(dbConn)

End Function

Private Function F_ToMauItem()
    On Error Resume Next
    If NhiemVuIDGlobal <> 0 Then
        With lvNhiemVuCuaPhongBan
            Dim colIndex As Integer
            For colIndex = 1 To .ListItems(1).ListSubItems.Count
                If NhiemVuIDGlobal = .ListItems(colIndex).SubItems(8) Then
                    .ListItems(colIndex).ForeColor = &H8000000D
                    .ListItems(colIndex).Bold = True
                    For colIndex2 = 1 To .ListItems(1).ListSubItems.Count
                        .ListItems(colIndex).ListSubItems(colIndex2).ForeColor = &H8000000D
                        .ListItems(colIndex).ListSubItems(colIndex2).Bold = True

                    Next colIndex2
                 Exit For
                End If
            Next colIndex

        End With
    End If

End Function

Private Sub btnCaNam_Click()
    Call F_btnTatCa
    Call F_ToMauItem
End Sub

Private Sub btnT1_Click()
    Call F_SetBtnThang(1)
End Sub

Private Sub btnT2_Click()
    Call F_SetBtnThang(2)
End Sub

Private Sub btnT3_Click()
    Call F_SetBtnThang(3)
End Sub

Private Sub btnT4_Click()
    Call F_SetBtnThang(4)
End Sub

Private Sub btnT5_Click()
    Call F_SetBtnThang(5)
End Sub

Private Sub btnT6_Click()
    Call F_SetBtnThang(6)
End Sub

Private Sub btnT7_Click()
    Call F_SetBtnThang(7)
End Sub

Private Sub btnT8_Click()
    Call F_SetBtnThang(8)
End Sub

Private Sub btnT9_Click()
    Call F_SetBtnThang(9)
End Sub

Private Sub btnT10_Click()
    Call F_SetBtnThang(10)
End Sub

Private Sub btnT11_Click()
    Call F_SetBtnThang(11)
End Sub

Private Sub btnT12_Click()
    Call F_SetBtnThang(12)
End Sub

Private Sub cbbPhongBan_Click()
    F_ClearForm
    Dim i
    For i = 1 To 12
        With Form_ChiTieuNhiemVuDinhluong.Controls("btnT" & i)
            .BackColor = bgDefault
        End With
    Next i

    If cbbPhongBan.value <> "" Then
        Call F_Thang_DaLapKh
        Call F_ViewListViewNhiemVuPhongBan
        Call F_btnTatCa
        Call F_ToMauItem
    End If
End Sub

Private Sub lvNhiemVuCuaPhongBan_ItemClick(Byval Item As MSComctlLib.ListItem)
    On Error Resume Next
    IndexChiTieuNVDL = Item.index

    With lvNhiemVuCuaPhongBan
        txtTenNhiemVu.value = FontConverter(.ListItems(Item.index).Text, 2, 1)
        txtTenMucTieu.value = FontConverter(.ListItems(Item.index).SubItems(1), 2, 1)
        txtDinhMucToiThieu.value = .ListItems(Item.index).SubItems(2)
        txtDinhMucYeuCau.value = .ListItems(Item.index).SubItems(3)

        cbbPhuongThucTinh.Text = FontConverter(.ListItems(Item.index).SubItems(5), 2, 1)
        txtTrongSo.value = .ListItems(Item.index).SubItems(6)
        ChiTieuNhiemVuID = .ListItems(Item.index).SubItems(7)
        NhiemVuID = .ListItems(Item.index).SubItems(8)
        txtGhiChu.value = FontConverter(.ListItems(Item.index).SubItems(10), 2, 1)
        CongThucTinh = SelectCongThucTH.value
        cbbDonViTinh.Text = FontConverter(.ListItems(Item.index).SubItems(12), 2, 1)

        If .ListItems(Item.index).SubItems(5) = "" Then
            cbbPhuongThucTinh = ""
        End If

    End With
End Sub

Private Sub btnCapNhat_Click()
    On Error Resume Next
    Dim dbConn As Object
    Dim Rs As Object
    Dim Query As String
    Dim valueDinhMucToiThieu As Currency
    Dim ValueDinhMucYeuCau As Currency
    Dim ValuePhuongThucTinhID As Integer
    Dim valueTrongSo As Currency
    Dim valueGhiChu As String
    Dim valuePhongBan As Integer

    valueDinhMucToiThieu = txtDinhMucToiThieu.value
    ValueDinhMucYeuCau = txtDinhMucYeuCau.value

    If Not IsNull(cbbPhuongThucTinh.value) Then
        ValuePhuongThucTinhID = cbbPhuongThucTinh.value
    End If

    If Not IsNull(cbbPhongBan.value) Then
        valuePhongBan = cbbPhongBan.value
    End If

    Dim i As Integer

    valueTrongSo = txtTrongSo.value
    valueGhiChu = txtGhiChu.value

    If F_KiemTraDuLieuTruocKhiCapNhat(valuePhongBan, valueDinhMucToiThieu, ValueDinhMucYeuCau, ValuePhuongThucTinhID, valueTrongSo) Then
     Exit Sub
    End If

    Set dbConn = ConnectToDatabase

    'ChiTieuNhiemVuID = 0 thêm moi
    If ChiTieuNhiemVuID = 0 Then
        Call F_ThemMoiChiTieuNhiemVu(dbConn, valueDinhMucToiThieu, ValueDinhMucYeuCau, ValuePhuongThucTinhID, valueTrongSo, valueGhiChu, valuePhongBan)
    Else
        ' ChiTieuNhiemVuID > 0 C?p nh?t

        Call F_CapNhatChiTieuNhiemVu(dbConn, valueDinhMucToiThieu, ValueDinhMucYeuCau, ValuePhuongThucTinhID, valueTrongSo, valueGhiChu, valuePhongBan)
    End If

    Set Rs = Nothing

    Call CloseDatabaseConnection(dbConn)
    Call F_ViewListViewNhiemVuPhongBan
    Call F_ClearForm
    Call F_ToMauItem
    Call ThongBao_ThanhCong
End Sub

Private Function F_ValueDinhMucTheoCongThuc(value As Currency) As Currency
    'CongThucTinh (Cong thuc tong hop)
    ' CongthucTinh = 1 Tong  dinh muc toi thieu va dinh muc yeu cau chia 12

    If CongThucTinh = 1 Then
        F_ValueDinhMucTheoCongThuc = value / 12
    Elseif CongThucTinh = 2 Or CongThucTinh = 3 Or CongThucTinh = 4 Then ' Trung binh, Max, Min Giu nguyen gia tri nhap vao
        F_ValueDinhMucTheoCongThuc = value
    End If

End Function

Private Function F_ThemMoiChiTieuNhiemVu(dbConn As Object, _
    valueDinhMucToiThieu As Currency, _
    ValueDinhMucYeuCau As Currency, _
    ValuePhuongThucTinhID As Integer, _
    valueTrongSo As Currency, _
    valueGhiChu As String, _
    valuePhongBan As Integer _
    )

    Dim Rs As Object
    Dim Query As String
    Dim i As Integer


    ' ValueNam = True Tuc la da chon ca cam
    If ValueNam Then

        For i = 0 To 12
            If i = 0 Then
                ' C? nam
                Query = "INSERT INTO CV_ChiTieuNhiemVu(NhiemVuID, DinhMucToiThieu,DinhMucYeuCau,PhuongThucTinhID, TrongSo, GhiChu, PhongBanID, Thang, Nam, DonViTinhID) " & _
                "VALUES(" & NhiemVuID & ", " & valueDinhMucToiThieu & ", " & ValueDinhMucYeuCau & ", " & ValuePhuongThucTinhID & ", " & valueTrongSo & ", N'" & valueGhiChu & "', " & valuePhongBan & ",0, " & cbbNam.value & ", " & cbbDonViTinh.value & ")"
            Else
                ' i =1 -> 12  cac thang
                If DongBo Then
                    Query = "INSERT INTO CV_ChiTieuNhiemVu(NhiemVuID, DinhMucToiThieu,DinhMucYeuCau,PhuongThucTinhID, TrongSo, GhiChu, PhongBanID, Thang, Nam, DonViTinhID) " & _
                    "VALUES(" & NhiemVuID & ", " & F_ValueDinhMucTheoCongThuc(valueDinhMucToiThieu) & ", " & TienThang & i & ", " & ValuePhuongThucTinhID & ", " & valueTrongSo & ", N'" & valueGhiChu & "', " & valuePhongBan & "," & i & ", " & cbbNam.value & ", " & cbbDonViTinh.value & ")"
                Else
                    Query = "INSERT INTO CV_ChiTieuNhiemVu(NhiemVuID, DinhMucToiThieu,DinhMucYeuCau,PhuongThucTinhID, TrongSo, GhiChu, PhongBanID, Thang, Nam, DonViTinhID) " & _
                    "VALUES(" & NhiemVuID & ", " & F_ValueDinhMucTheoCongThuc(valueDinhMucToiThieu) & ", " & F_ValueDinhMucTheoCongThuc(ValueDinhMucYeuCau) & ", " & ValuePhuongThucTinhID & ", " & valueTrongSo & ", N'" & valueGhiChu & "', " & valuePhongBan & "," & i & ", " & cbbNam.value & ", " & cbbDonViTinh.value & ")"
                End If

            End If

            ' thuc hien lu dung lieu
            Set Rs = dbConn.Execute(Query)
        Next i
    Else
        ' Nhap theo tung thang
        Query = "INSERT INTO CV_ChiTieuNhiemVu(NhiemVuID, DinhMucToiThieu,DinhMucYeuCau,PhuongThucTinhID, TrongSo, GhiChu, PhongBanID, Thang, Nam, DonViTinhID) " & _
        "VALUES(" & NhiemVuID & ", " & valueDinhMucToiThieu & ", " & ValueDinhMucYeuCau & ", " & ValuePhuongThucTinhID & ", " & valueTrongSo & ", N'" & valueGhiChu & "', " & valuePhongBan & ", " & ValueThang & ", " & cbbNam.value & ", " & cbbDonViTinh.value & ")"
        Set Rs = dbConn.Execute(Query)
    End If

    Set Rs = Nothing
End Function

Private Function F_CapNhatChiTieuNhiemVu(dbConn As Object, _
    valueDinhMucToiThieu As Currency, _
    ValueDinhMucYeuCau As Currency, _
    ValuePhuongThucTinhID As Integer, _
    valueTrongSo As Currency, _
    valueGhiChu As String, _
    valuePhongBan As Integer _
    )

    Dim i As Integer
    Dim Query As String
    Dim Rs As Object
    ' ValueNam = True Tuc la da chon ca cam
    If ValueNam Then
        Dim msgValue As VbMsgBoxResult
        Dim Mes As String

        Mes = "B" & ChrW(7841) & "n có mu" & ChrW(7889) & "n c" _
        & ChrW(7853) & "p nh" & ChrW(7853) & "t d" & ChrW(7919) & " li" & ChrW( _
        7879) & "u cho các tháng còn l" & ChrW(7841) & "i không ?"

        msgValue = Application.Assistant.DoAlert(UniConvert("Carnh baso", "Telex"), UniConvert(Mes, "Telex"), msoAlertButtonYesNo, msoAlertIconWarning, 0, 0, 1)

        If msgValue = vbYes Then
            For i = 0 To 12

                If i = 0 Then
                    ' Ca nam
                    Query = "UPDATE CV_ChiTieuNhiemVu Set DinhMucToiThieu = " & valueDinhMucToiThieu & ", DinhMucYeuCau = " & ValueDinhMucYeuCau & ", PhuongThucTinhID = " & ValuePhuongThucTinhID & ", TrongSo = " & valueTrongSo & ", GhiChu = N'" & valueGhiChu & "' WHERE Thang = " & i & " And Nam = " & cbbNam.value & " And NhiemVuID = " & NhiemVuID & " And DonViTinhID = " & cbbDonViTinh.value
                Else
                    ' i =1 -> 12  cac thang
                    If DongBo Then
                        Query = "UPDATE CV_ChiTieuNhiemVu Set DinhMucToiThieu = " & F_ValueDinhMucTheoCongThuc(valueDinhMucToiThieu) & ", DinhMucYeuCau = " & F_ValueDinhMucTheoCongThuc(ValueDinhMucYeuCau) & ", PhuongThucTinhID = " & ValuePhuongThucTinhID & ", TrongSo = " & valueTrongSo & ", GhiChu = N'" & valueGhiChu & "' WHERE Thang = " & i & " And Nam = " & cbbNam.value & " And NhiemVuID = " & NhiemVuID & " And DonViTinhID = " & cbbDonViTinh.value
                    Else
                        Query = "UPDATE CV_ChiTieuNhiemVu Set DinhMucToiThieu = " & F_ValueDinhMucTheoCongThuc(valueDinhMucToiThieu) & ", DinhMucYeuCau = " & F_ValueDinhMucTheoCongThuc(ValueDinhMucYeuCau) & ", PhuongThucTinhID = " & ValuePhuongThucTinhID & ", TrongSo = " & valueTrongSo & ", GhiChu = N'" & valueGhiChu & "' WHERE Thang = " & i & " And Nam = " & cbbNam.value & " And NhiemVuID = " & NhiemVuID & " And DonViTinhID = " & cbbDonViTinh.value
                    End If

                End If

                Set Rs = dbConn.Execute(Query)
            Next i
        Else
            ' Khi khong muon cap nhat cho cac thang
            Query = "UPDATE CV_ChiTieuNhiemVu Set DinhMucToiThieu = " & valueDinhMucToiThieu & ", DinhMucYeuCau = " & ValueDinhMucYeuCau & ", PhuongThucTinhID = " & ValuePhuongThucTinhID & ", TrongSo = " & valueTrongSo & ", GhiChu = N'" & valueGhiChu & "' WHERE Thang =  0  And Nam = " & cbbNam.value & " And NhiemVuID = " & NhiemVuID & " And DonViTinhID = " & cbbDonViTinh.value & ""
            Set Rs = dbConn.Execute(Query)
        End If
    Else
        ' Nhap theo tung thang
        Query = "UPDATE CV_ChiTieuNhiemVu Set DinhMucToiThieu = " & valueDinhMucToiThieu & ", DinhMucYeuCau = " & ValueDinhMucYeuCau & ", PhuongThucTinhID = " & ValuePhuongThucTinhID & ", TrongSo = " & valueTrongSo & ", GhiChu = N'" & valueGhiChu & "' WHERE Thang = " & ValueThang & " And Nam = " & cbbNam.value & " And NhiemVuID = " & NhiemVuID & " And DonViTinhID = " & cbbDonViTinh.value

        Set Rs = dbConn.Execute(Query)

        Set Rs = Nothing

        Dim Message As String
        Message = "Th" & ChrW(225) & "ng 12 s" & ChrW(7869) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c c" & ChrW(7853) & "p nh" & ChrW(7853) & "t l" & ChrW(7841) & "i "
        msgValue = Application.Assistant.DoAlert(UniConvert("Carnh baso", "Telex"), UniConvert(Message, "Telex"), msoAlertButtonYesNo, msoAlertIconWarning, 0, 0, 0)

        If msgValue = vbYes Then
            ' Dinh muc Yeu cau
            Query = "Update CV_ChiTieuNhiemVu Set DinhMucYeuCau = " & _
            "(Select top 1 DinhMucYeuCau from CV_ChiTieuNhiemVu KHN WHERE Nam = CV_ChiTieuNhiemVu.Nam And NhiemVuID = CV_ChiTieuNhiemVu.NhiemVuID And PhongBanID = CV_ChiTieuNhiemVu.PhongBanID And Thang = 0) - " & _
            "(Select Sum(DinhMucYeuCau) from CV_ChiTieuNhiemVu KHT WHERE Nam = CV_ChiTieuNhiemVu.Nam And NhiemVuID = CV_ChiTieuNhiemVu.NhiemVuID And PhongBanID = CV_ChiTieuNhiemVu.PhongBanID And Thang between 1 And 11) " & _
            "WHERE Nam = " & cbbNam.value & " And NhiemVuID = " & NhiemVuID & " And PhongBanID = " & cbbPhongBan.value & " And Thang = 12"

            Set Rs = dbConn.Execute(Query)
            Set Rs = Nothing

            '            Dinh muc toi thieu
            Query = "Update CV_ChiTieuNhiemVu Set DinhMucToiThieu = " & _
            "(Select top 1 DinhMucToiThieu from CV_ChiTieuNhiemVu KHN WHERE Nam = CV_ChiTieuNhiemVu.Nam And NhiemVuID = CV_ChiTieuNhiemVu.NhiemVuID And PhongBanID = CV_ChiTieuNhiemVu.PhongBanID And Thang = 0) - " & _
            "(Select Sum(DinhMucToiThieu) from CV_ChiTieuNhiemVu KHT WHERE Nam = CV_ChiTieuNhiemVu.Nam And NhiemVuID = CV_ChiTieuNhiemVu.NhiemVuID And PhongBanID = CV_ChiTieuNhiemVu.PhongBanID And Thang between 1 And 11) " & _
            "WHERE Nam = " & cbbNam.value & " And NhiemVuID = " & NhiemVuID & " And PhongBanID = " & cbbPhongBan.value & " And Thang = 12"

            Set Rs = dbConn.Execute(Query)
            Set Rs = Nothing
        End If
    End If

    Set Rs = Nothing
End Function

Private Function F_KiemTraDuLieuTruocKhiCapNhat(PhongBanID As Integer, DinhMucToiThieu As Currency, DinhMucYeuCau As Currency, PhuongThucTinh As Integer, TrongSo As Currency) As Boolean
    If PhongBanID = 0 Or IsNull(cbbPhongBan.value) Then
        NoiDung = "Ch" & ChrW(7885) & "n ph?ng ban"
        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
        FrameGoupPhongBan.BorderColor = colorError
        F_KiemTraDuLieuTruocKhiCapNhat = True
     Exit Function
    Else
        FrameGoupPhongBan.BorderColor = colorDefault
    End If

    If DinhMucToiThieu = 0 Then
        NoiDung = "Nh" & ChrW(7853) & "p " & ChrW(273) & ChrW(7883) _
        & "nh m" & ChrW(7913) & "c t" & ChrW(7889) & "i thi" & ChrW(7875) & "u"
        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
        txtDinhMucToiThieu.BorderColor = colorError
        F_KiemTraDuLieuTruocKhiCapNhat = True
     Exit Function
    Else
        txtDinhMucToiThieu.BorderColor = colorDefault
    End If

    If DinhMucYeuCau = 0 Then
        NoiDung = "Nh" & ChrW(7853) & "p " & ChrW(273) & ChrW(7883) _
        & "nh m" & ChrW(7913) & "c yêu c" & ChrW(7847) & "u"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
        txtDinhMucYeuCau.BorderColor = colorError
        F_KiemTraDuLieuTruocKhiCapNhat = True
     Exit Function
    Else
        txtDinhMucYeuCau.BorderColor = colorDefault

    End If

    If PhuongThucTinh = 0 Then
        NoiDung = "Ch" & ChrW(7885) & "n ph" & ChrW(432) & ChrW(417 _
        ) & "ng th" & ChrW(7913) & "c tính"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
        cbbPhuongThucTinh.BorderColor = colorError
        F_KiemTraDuLieuTruocKhiCapNhat = True

     Exit Function
    Else
        cbbPhuongThucTinh.BorderColor = colorDefault
    End If

    If TrongSo = 0 Then
        NoiDung = "Nh" & ChrW(7853) & "p tr" & ChrW(7885) & "ng s" _
        & ChrW(7889)

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
        txtTrongSo.BorderColor = colorError
        F_KiemTraDuLieuTruocKhiCapNhat = True
     Exit Function

    Else
        txtTrongSo.BorderColor = colorDefault
    End If

    If ValueThang = 0 And ValueNam = False Then
        NoiDung = "Ch" & ChrW(7885) & "n tháng l" & ChrW(7853) & _
        "p k" & ChrW(7871) & " ho" & ChrW(7841) & "ch"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
        F_KiemTraDuLieuTruocKhiCapNhat = True
    End If
End Function

Private Function F_PhuongThucTinh()
    Dim dbConn As Object
    Dim Rs As Object
    Dim Query As String

    Set dbConn = ConnectToDatabase
    Query = "Select PhuongThucTinhID, TenPhuongThucTinh FROM DM_PhuongThucTinh"

    Call ViewListBox(Query, cbbPhuongThucTinh, dbConn)

    Call CloseDatabaseConnection(dbConn)
End Function

Private Function F_KhoiTaoPhongBan()
    On Error Resume Next
    Dim Query As String
    Dim dbConn As Object

    Set dbConn = ConnectToDatabase

    If ValueCacDonViPhuTrachID = "" Or ValueCacDonViPhuTrachID = 0 Then
        Query = "Select PhongBanID, TenPhongBan FROM PhongBan"
        Call ViewListBox(Query, cbbPhongBan, dbConn)
    Else
        Query = "Select PhongBanID, TenPhongBan FROM PhongBan WHERE PhongbanID IN (" & ValueCacDonViPhuTrachID & ")"
        Call ViewListBox(Query, cbbPhongBan, dbConn)

        With cbbPhongBan
            .Text = .List(0, 1)
        End With

        Call F_btnTatCa
        Call F_ToMauItem
    End If

    Call CloseDatabaseConnection(dbConn)
End Function

Private Function F_TableChiTieuNhiemVuDinh()
    Dim TenNhiemVu, TenMucTieu, DinhMucToiThieu, DinhMucYeuCau, PhuongThucTinh, TrongSo, DaLap As String

    TenNhiemVu = "Tên nhi" & ChrW(7879) & "m v" & ChrW(7909)
    TenMucTieu = "T" & ChrW(234) & "n ch" & ChrW(7881) & " ti" & ChrW(234) & "u KGI"
    DinhMucToiThieu = ChrW(272) & ChrW(7883) & "nh m" & ChrW(7913) & _
    "c t" & ChrW(7889) & "i thi" & ChrW(7875) & "u"
    DinhMucYeuCau = ChrW(272) & ChrW(7883) & "nh m" & ChrW(7913) & _
    "c yêu c" & ChrW(7847) & "u"
    PhuongThucTinh = "Ph" & ChrW(432) & ChrW(417) & "ng th" & ChrW( _
    7913) & "c tính"
    TrongSo = "Tr" & ChrW(7885) & "ng s" & ChrW(7889)
    DaLap = ChrW(272) & "ã l" & ChrW(7853) & "p"


    Dim i As Integer
    With lvNhiemVuCuaPhongBan.ColumnHeaders
        For i = .Count To 1 Step -1
            .Remove i
        Next i
    End With

    With lvNhiemVuCuaPhongBan.ColumnHeaders
        .Add , , UniToWindows1258(TenNhiemVu), 138
        .Add , , UniToWindows1258(TenMucTieu), 140
        .Add , , UniToWindows1258(DinhMucToiThieu), 140
        .Add , , UniToWindows1258(DinhMucYeuCau), 140
        .Add , , UniToWindows1258(DaLap), 0
        .Add , , UniToWindows1258(PhuongThucTinh), 140
        .Add , , UniToWindows1258(TrongSo), 80
        .Add , , "ChiTieuNhiemVuID", 0
        .Add , , "NhiemVuID", 0
        .Add , , "Thang", 0
        .Add , , "GhiChu", 0
        .Add , , "CongThucTinhID", 0
        .Add , , "TenDonViTinh", 100
    End With
End Function

Private Function F_ViewListViewNhiemVuPhongBan()

    On Error Resume Next
    Dim dbConn As Object
    Dim Rs As Object
    Dim Query As String

    Query = "Select TVNV.TenNhiemVu,TVNV.TenMucTieu,ISNULL(Data_CTNV.DinhMucToiThieu,0) As DinhMucToiThieu,ISNULL(Data_CTNV.DinhMucYeuCau,0) As DinhMucYeuCau, " & _
    "ISNULL(Case WHEN CTTH.CongThucTinhID = 1 Then (Select sum(DinhMucYeuCau) from CV_ChiTieuNhiemVu where Nam = " & cbbNam.value & " And PhongBanID = " & cbbPhongBan.value & " And Thang > 0 And NhiemVuID = NV.NhiemVuID ) " & _
    "WHEN CTTH.CongThucTinhID = 2 Then (Select avg(DinhMucYeuCau) from CV_ChiTieuNhiemVu where Nam = " & cbbNam.value & " And PhongBanID = " & cbbPhongBan.value & " And Thang > 0 And NhiemVuID = NV.NhiemVuID ) " & _
    "WHEN CTTH.CongThucTinhID = 3 Then (Select max(DinhMucYeuCau) from CV_ChiTieuNhiemVu where Nam = " & cbbNam.value & " And PhongBanID = " & cbbPhongBan.value & " And Thang > 0 And NhiemVuID = NV.NhiemVuID ) " & _
    "WHEN CTTH.CongThucTinhID = 4 Then (Select min(DinhMucYeuCau) from CV_ChiTieuNhiemVu where Nam = " & cbbNam.value & " And PhongBanID = " & cbbPhongBan.value & " And Thang > 0 And NhiemVuID = NV.NhiemVuID ) " & _
    "END,0) As DaLap, ISNULL(Data_CTNV.TenPhuongThucTinh,'') As TenPhuongThucTinh,ISNULL(Data_CTNV.TrongSo,0) As TrongSo, " & _
    "ISNULL(Data_CTNV.ChiTieuNhiemVuID,0) As ChiTieuNhiemVuID,ISNULL(Data_CTNV.NhiemVuID,NV.NhiemVuID) As NhiemVuID, " & _
    "ISNULL(Data_CTNV.Thang,0) As Thang, ISNULL(Data_CTNV.GhiChu,'') As GhiChu, NV.CongThucTinhID,ISNULL(Data_CTNV.TenDonViTinh,'') As DonViTinh " & _
    "FROM CV_NhiemVu NV INNER JOIN CV_ThuVienNhiemVu TVNV ON NV.ThuVienNhiemVuID = TVNV.ThuVienNhiemVuID LEFT JOIN DM_CongThucTinh CTTH ON CTTH.CongThucTinhID = NV.CongThucTinhID " & _
    "Left Join ( Select CTNV.DinhMucToiThieu,CTNV.DinhMucYeuCau,PTT.TenPhuongThucTinh, CTNV.TrongSo,CTNV.ChiTieuNhiemVuID,NV.NhiemVuID, CTNV.Thang, CtNV.GhiChu,DVT.TenDonViTinh " & _
    "From CV_NhiemVu NV LEFT JOIN CV_ChiTieuNhiemVu CTNV ON NV.NhiemVuID = CTNV.NhiemVuID INNER JOIN CV_ThuVienNhiemVu TVNV ON NV.ThuVienNhiemVuID = TVNV.ThuVienNhiemVuID " & _
    "LEFT JOIN DM_PhuongThucTinh PTT ON PTT.PhuongThucTinhID = CTNV.PhuongThucTinhID LEFT JOIN DM_DonViTinh DVT ON CTNV.DonViTinhID = DVT.DonViTinhID " & _
    "Where CTNV.Nam =  " & cbbNam.value & " And CtNV.Thang =  " & ValueThang & " And PhongBanID = " & cbbPhongBan.value & ") Data_CTNV on NV.NhiemVuID = Data_CTNV.NhiemVuID " & _
    "where NV.Nam =  " & cbbNam.value & " And EXISTS ( Select 1 FROM STRING_SPLIT(NV.CacDonViPhuTrachID, ',') As SplitValues " & _
    "WHERE SplitValues.value = Cast(" & cbbPhongBan.value & " As nvarchar(100))) And NV.ChucNangID = " & SelectChucNang.value & " And NV.BSC_ID > 0"

    Set dbConn = ConnectToDatabase


    If Not dbConn Is Nothing Then
        Set Rs = dbConn.Execute(Query)
        With lvNhiemVuCuaPhongBan
            Dim i As Integer
            For i = .ListItems.Count To 1 Step -1
                .ListItems.Remove i
            Next i
            Set addedItems = New Collection
        End With

        Do Until Rs.EOF
            Dim foundItem As Boolean
            foundItem = False
            Dim existingItem As Variant

            For Each existingItem In addedItems
                If existingItem = Rs.Fields("NhiemVuID").value Then
                    foundItem = True
                 Exit For
                End If
            Next existingItem

            If Not foundItem Then
                Dim ListItem As ListItem
                Set ListItem = lvNhiemVuCuaPhongBan.ListItems.Add(, , UniToWindows1258(Rs.Fields("TenNhiemVu").value))
                ListItem.SubItems(1) = UniToWindows1258(Rs.Fields("TenMucTieu").value)
                ListItem.SubItems(2) = Rs.Fields("DinhMucToiThieu").value
                ListItem.SubItems(3) = Rs.Fields("DinhMucYeuCau").value
                ListItem.SubItems(4) = Rs.Fields("DaLap").value
                ListItem.SubItems(5) = UniToWindows1258(Rs.Fields("TenPhuongThucTinh").value)
                ListItem.SubItems(6) = Rs.Fields("TrongSo").value
                ListItem.SubItems(7) = Rs.Fields("ChiTieuNhiemVuID").value
                ListItem.SubItems(8) = Rs.Fields("NhiemVuID").value
                ListItem.SubItems(9) = Rs.Fields("Thang").value
                ListItem.SubItems(10) = UniToWindows1258(Rs.Fields("GhiChu").value)
                ListItem.SubItems(11) = Rs.Fields("CongThucTinhID").value
                ListItem.SubItems(12) = UniToWindows1258(Rs.Fields("DonViTinh").value)
                addedItems.Add Rs.Fields("NhiemVuID").value
            End If
            Rs.MoveNext
        Loop
    Else
        MsgBox "Mat Ket noi csdl"
    End If
End Function

Private Function F_SetBtnThang(i As Integer)
    If F_KiemTraChonPhongBanChua Then
     Exit Function
    End If
    ValueNam = False
    Call F_TableChiTieuNhiemVuDinh
    Dim j As Integer

    For j = 1 To 12
        If j <> i Then
            With Form_ChiTieuNhiemVuDinhluong.Controls("btnT" & j)
                If .BackColor <> bgCheck Then
                    .BackColor = bgDefault
                End If
            End With
        End If
    Next j

    Call F_Thang_DaLapKh
    With Form_ChiTieuNhiemVuDinhluong.Controls("btnT" & i)
        .BackColor = bgActive
        ValueThang = i
    End With

    btnCaNam.BackColor = bgDefault

    If cbbPhongBan.value <> "" Then
        Call F_ViewListViewNhiemVuPhongBan
    End If
    Call F_ClearForm
    Call F_ToMauItem
    ValueNam = False
End Function

Public Function F_KiemTraChonPhongBanChua() As Boolean
    If cbbPhongBan.Text = "" Then
        NoiDung = "Ch" & ChrW(7885) & "n phòng ban tr" & ChrW(432) _
        & ChrW(7899) & "c !"
        FrameGoupPhongBan.BorderColor = colorError
        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
        F_KiemTraChonPhongBanChua = True
     Exit Function
    Else
        FrameGoupPhongBan.BorderColor = bgDefault
        F_KiemTraChonPhongBanChua = False
    End If
End Function

Public Function F_btnTatCa()
    F_ClearForm

    If F_KiemTraChonPhongBanChua Then
     Exit Function
    End If

    Dim i As Integer

    For i = 1 To 12
        With Form_ChiTieuNhiemVuDinhluong.Controls("btnT" & i)
            .BackColor = bgDefault
        End With
    Next i
    F_Thang_DaLapKh

    ValueThang = 0
    ValueNam = True
    btnCaNam.BackColor = bgActive
    Call F_TableChiTieuNhiemVuDinh
    Call F_ViewListViewNhiemVuPhongBan
End Function

Public Function F_Thang_DaLapKh()
    Dim Query As String
    Dim dbConn As Object
    Dim Rs As Object
    Dim Months As Variant
    Dim i As Integer
    Call F_ClearForm

    Query = "Select DISTINCT CTNV.Thang FROM CV_NhiemVu NV " & _
    "LEFT JOIN CV_ChiTieuNhiemVu CTNV ON NV.NhiemVuID = CTNV.NhiemVuID " & _
    "WHERE NV.Nam = " & cbbNam.value & " And CTNV.Thang IS Not NULL And EXISTS (Select 1 FROM STRING_SPLIT(NV.CacDonViPhuTrachID, ',') As SplitValues WHERE SplitValues.value = '" & cbbPhongBan.value & "') And CTNV.PhongBanID =  " & cbbPhongBan.value & ""

    Set dbConn = ConnectToDatabase

    If Not dbConn Is Nothing Then
        Set Rs = dbConn.Execute(Query)

        If Not Rs.EOF And Not Rs.BOF Then
            Months = Rs.GetRows()

            With Form_ChiTieuNhiemVuDinhluong

                For i = LBound(Months, 2) To UBound(Months, 2)

                    If Months(0, i) = 0 Then
                        btnCaNam.BackColor = bgCheck
                    Else
                        With .Controls("btnT" & Months(0, i))
                            .BackColor = bgCheck
                        End With
                    End If

                Next i

                ValueNam = False
            End With
            btnCaNam.BackColor = bgDefault
        Else
            btnCaNam.BackColor = bgActive

            ValueNam = True
            With Form_ChiTieuNhiemVuDinhluong
                For i = 1 To 12
                    With .Controls("btnT" & i)
                        .BackColor = bgDefault
                    End With
                Next i
            End With
        End If

    Else
        MsgBox "Mat Ket noi csdl"
    End If
    Call CloseDatabaseConnection(dbConn)
End Function

Public Function F_XoaChiTieuNhiemVuDinhLuong()
    If ChiTieuNhiemVuID = 0 Then
        TieuDe = "BOS xin thông báo"
        NoiDung = "Ch" & ChrW(7885) & "n ch" & ChrW(7881) & _
        " tiêu c" & ChrW(7847) & "n xóa"
        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
     Exit Function
    End If

    Dim msgValue As VbMsgBoxResult
    msgValue = Application.Assistant.DoAlert(UniConvert("Carnh baso", "Telex"), UniConvert("Bajn muoosn xosa duwx lieeju?", "Telex"), msoAlertButtonYesNo, msoAlertIconWarning, 0, 0, 0)

    If msgValue = vbYes Then
        Dim Query As String
        Dim dbConn As Object
        Dim Rs As Object

        Query = "DELETE CV_ChiTieuNhiemVu WHERE ChiTieuNhiemVuID = " & ChiTieuNhiemVuID & ""
        Set dbConn = ConnectToDatabase
        Set Rs = dbConn.Execute(Query)
        Call CloseDatabaseConnection(dbConn)
        Set Rs = Nothing
    End If
    Call F_ViewListViewNhiemVuPhongBan
    Call F_ClearForm
    Call ThongBao_ThanhCong
End Function

Private Function F_ClearForm()
    txtTenNhiemVu = ""
    txtTenMucTieu = ""
    txtDinhMucToiThieu = ""
    txtDinhMucYeuCau = ""
    cbbPhuongThucTinh = ""
    txtTrongSo = ""
    txtGhiChu = ""
    ChiTieuNhiemVuID = 0
    cbbDonViTinh = ""
    DongBo = False

    Dim i As Integer

    For i = 1 To 12
        CallByName Me, "TienThang" & i, VbLet, 0
    Next i
End Function





