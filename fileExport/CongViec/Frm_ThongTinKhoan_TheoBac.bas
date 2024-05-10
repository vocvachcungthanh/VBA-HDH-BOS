

Dim ThietLapKhoanID As Integer
Dim ViTriID As Integer
Dim CongViecID As Integer
Dim ThietLapKhoan_TheoBacID As Integer
Public TenBacKhoan As String
Public HeSoKhoan As Double
Public GiaiKhoanTu As Double
Public GhiChu As String
Dim tbl_Data As Variant
Dim SQLStr As String
Dim Cn As ADODB.Connection
Dim StrCnn As String
Dim Rs As ADODB.Recordset

Private Sub cmd_CapNhat_Click()
    DinhDang_BatDau
    Check_DuLieu
    If txt_HeSo.BorderColor <> vbBlack Or txt_GiaiKhoanTu.BorderColor <> vbBlack Then
        Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), UniConvert("Bajn caafn ddieefn ddaafy ddur vaf ddusng ddijnh dajng thoong tin.", "Telex"), msoAlertButtonOK, 0, 0, 0, 0
     Exit Sub
    End If

    TenBacKhoan = txt_TenBacKhoan.Text
    HeSoKhoan = txt_HeSo.Text
    GiaiKhoanTu = txt_GiaiKhoanTu.Text
    GhiChu = txt_GhiChu.Text

End Sub

Private Function ThemBacLenCSDL()
    DinhDang_BatDau
    Check_DuLieu
    If txt_HeSo.BorderColor <> vbBlack Or txt_GiaiKhoanTu.BorderColor <> vbBlack Then
        Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), UniConvert("Bajn caafn ddieefn ddaafy ddur vaf ddusng ddijnh dajng thoong tin.", "Telex"), msoAlertButtonOK, 0, 0, 0, 0
     Exit Sub
    End If

    TenBacKhoan = txt_TenBacKhoan.Text
    HeSoKhoan = txt_HeSo.Text
    GiaiKhoanTu = txt_GiaiKhoanTu.Text
    GhiChu = txt_GhiChu.Text
    If ThietLapKhoan_TheoBacID = 0 Then
        SQLStr = "Insert into CV_ThietLapKhoan_TheoBac (ThietLapKhoanID, TenBac, HeSo, GiaiKhoanTu, GhiChu) values ((Select top 1 ThietLapKhoanID from CV_thietLapKhoan where ViTriID = " & ViTriID & " And CongViecID = " & CongViecID & "), N'" & TenBacKhoan & "', " & HeSoKhoan & ", " & GiaiKhoanTu & ", N'" & GhiChu & "')"
    Else
        SQLStr = "update CV_ThietLapKhoan_TheoBac Set TenBac = N'" & TenBacKhoan & "', HeSo = N'" & HeSoKhoan & "', GiaiKhoanTu = N'" & GiaiKhoanTu & "', GhiChu = N'" & GhiChu & "' where ThietLapKhoan_TheoBacID = " & ThietLapKhoan_TheoBacID & " ;   "
    End If

    StrCnn = KetNoiMayChu_KhachHang
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Set Rs = New ADODB.Recordset
    Rs.Open SQLStr, Cn, adOpenStatic

    Cn.Close
    Set Cn = Nothing
    Form_ThietLapLuongKhoan.Load_Data_Listview_ThongTinKhoan_TheoBac
    Call F_ClearDuLieu
    ThongBao_ThanhCong
End Function

Private Function F_ClearDuLieu()
    txt_TenBacKhoan.Text = ""
    txt_HeSo.Text = ""
    txt_GiaiKhoanTu.Text = ""
    txt_GhiChu.Text = ""
End Function

Private Sub cmd_Dong_Click()
    'Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), UniConvert("Bajn cos muoosn giuwx laji thoong tin cux ?", "Telex"), msoAlertButtonOK, msoAlertIconInfo, 0, 1, 0
    Form_ThietLapLuongKhoan.Load_Data_Listview_ThongTinKhoan_TheoBac
    Unload Me
End Sub

Private Sub cmd_LamMoi_Click()
    Dim XacNhan As VbMsgBoxResult
    XacNhan = Application.Assistant.DoAlert(UniConvert("Thoong baso", "Telex"), UniConvert("Bajn cos muoosn giuwx laji thoong tin cux ?", "Telex"), msoAlertButtonYesNo, msoAlertIconQuery, 0, 1, 0)
    If XacNhan = 7 Then
        ThietLapKhoan_TheoBacID = 0
        Call F_ClearDuLieu
    Else
        ThietLapKhoan_TheoBacID = 0
    End If

End Sub


Private Sub cmd_Xoa_Click()
    Dim XacNhan As VbMsgBoxResult
    XacNhan = Application.Assistant.DoAlert(UniConvert("Thoong baso", "Telex"), UniConvert("Bajn cos chawsc chawsn muoosn xosa baajc nafy ?", "Telex"), msoAlertButtonYesNo, msoAlertIconQuery, 0, 1, 0)
    If XacNhan = 7 Then
     Exit Sub

    End If

    SQLStr = "delete from CV_ThietLapKhoan_TheoBac where ThietLapKhoan_TheoBacID = " & ThietLapKhoan_TheoBacID & " ;   "

    StrCnn = KetNoiMayChu_KhachHang
    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Set Rs = New ADODB.Recordset
    Rs.Open SQLStr, Cn, adOpenStatic


    Cn.Close
    Set Cn = Nothing
    Form_ThietLapLuongKhoan.Load_Data_Listview_ThongTinKhoan_TheoBac
    Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), UniConvert("Xosa duwx lieeju thafnh coong.", "Telex"), 0, 0, 0, 0, 0

End Sub


Private Sub UserForm_Initialize()
    ViTriID = 0
    CongViecID = 0
    ThietLapKhoan_TheoBacID = 0
    On Error Resume Next
    ViTriID = Form_ThietLapLuongKhoan.cbbJobTitle.List(Form_ThietLapLuongKhoan.cbbJobTitle.ListIndex, 0)
    On Error Resume Next
    CongViecID = Form_ThietLapLuongKhoan.cbbKPIKhoan.List(Form_ThietLapLuongKhoan.cbbKPIKhoan.ListIndex, 0)
    On Error Resume Next
    ThietLapKhoan_TheoBacID = Form_ThietLapLuongKhoan.ListViewThongTinKhoan.SelectedItem.ListSubItems(4).Text

    If ViTriID = 0 Or CongViecID = 0 Then
        Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), UniConvert("Yeeu caafu chojn thieest laajp khoasn truwowsc.", "Telex"), 0, 0, 0, 0, 0
     Exit Sub
        Unload (Frm_ThongTinKhoan_TheoBac)
    End If

    DinhDang_BatDau
    HienDuLieu
End Sub

Sub HienDuLieu()
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

    StrCnn = KetNoiMayChu_KhachHang
    Dim SQLStr As String
    SQLStr = "Select top 1 TenBac, HeSo, GiaiKhoanTu, GhiChu from CV_ThietLapKhoan_TheoBac " & _
    "where ThietLapKhoanID in (Select top 1 ThietLapKhoanID from CV_thietLapKhoan where ViTriID = " & ViTriID & " And CongViecID = " & CongViecID & ")  " & _
    "And ThietLapKhoan_TheoBacID = " & ThietLapKhoan_TheoBacID & " ; "

    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Rs.Open SQLStr, Cn, adOpenStatic

    If Not Rs.EOF Then
        tbl_Data = Rs.GetRows()

        txt_TenBacKhoan.Text = tbl_Data(0, 0)
        txt_HeSo.Text = tbl_Data(1, 0)
        txt_GiaiKhoanTu.Text = tbl_Data(2, 0)
        txt_GhiChu.Text = tbl_Data(3, 0)
    End If

    Cn.Close
    Set Cn = Nothing

End Sub

Sub Check_DuLieu()
    If IsNumeric(txt_HeSo.Text) = False Then
        txt_HeSo.BorderColor = vbRed
    End If
    If IsNumeric(txt_GiaiKhoanTu.Text) = False Then
        txt_GiaiKhoanTu.BorderColor = vbRed
    End If

End Sub
Sub DinhDang_BatDau()
    txt_HeSo.BorderColor = vbBlack
    txt_GiaiKhoanTu.BorderColor = vbBlack
End Sub
