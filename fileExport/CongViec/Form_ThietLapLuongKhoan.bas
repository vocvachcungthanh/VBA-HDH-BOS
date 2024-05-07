Private NhanVienApDung As Collection
Private NhanVienID As Variant
Private colorError As String
Private colorDefault As String
Private ChiTieuKhoan As Double
Private ThuongKhoan As Double
Private ThietLapKhoanID As Integer
Private ThietLapKhoanNhanVienID As Integer
Private ThietLapKhoanID_TheoBac As Integer

Private Sub btLamMoi_Click()
    Call F_Reset
End Sub

Private Sub btnCreateGrade_Click()
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
    End If

    On Error Resume Next
    Frm_ThongTinKhoan_TheoBac.Show
End Sub

Private Sub btXoa_Click()
    Call F_DELETE
End Sub

Function DeleteThietLapKhoanTheoBac(ID)
    Dim XacNhan As VbMsgBoxResult
    XacNhan = Application.Assistant.DoAlert(UniConvert("Thoong baso", "Telex"), UniConvert("Bajn cos chawsc chawsn muoosn xosa baajc nafy ?", "Telex"), msoAlertButtonYesNo, msoAlertIconQuery, 0, 1, 0)
    If XacNhan = 7 Then
     Exit Function

    End If

    Dim Query As String

    Dim dbConn As Object

    Dim Rs As Object

    Set dbConn = ConnectToDatabase

    Query = "delete from CV_ThietLapKhoan_TheoBac where ThietLapKhoan_TheoBacID = " & ID & " ;   "
    Set Rs = dbConn.Execute(Query)

    Set Rs = Nothing
    Call CloseDatabaseConnection(dbConn)
    Query = ""
    Set Rs = Nothing

    Form_ThietLapLuongKhoan.Load_Data_Listview_ThongTinKhoan_TheoBac
    Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), UniConvert("Xosa duwx lieeju thafnh coong.", "Telex"), 0, 0, 0, 0, 0

End Function

Private Sub cbbDeletBac_Click()
    If ThietLapKhoanID_TheoBac = 0 Then
     Exit Sub
    End If

    Call DeleteThietLapKhoanTheoBac(ThietLapKhoanID_TheoBac)
    ThietLapKhoanID_TheoBac = 0
End Sub

Private Sub cbbNhanVienKhoan_Click()
    NhanVienID = cbbNhanVienKhoan.value
End Sub

Private Sub ListViewThongTinKhoan_Click()
    ThietLapKhoanID_TheoBac = ListViewThongTinKhoan.SelectedItem.ListSubItems(4).Text
End Sub

Private Sub TableObjective_Click()
    Call F_Reset
    With TableObjective
        cbbJobTitle.Text = FontConverter(.List(.ListIndex, 1), 0, 1)
        inputEntity.Caption = FontConverter(.List(.ListIndex, 3), 0, 1)
        InputDepartment.Caption = FontConverter(.List(.ListIndex, 4), 0, 1)
        txtNgayApDung.Text = .List(.ListIndex, 5)
        txtNgayHetHan.Text = .List(.ListIndex, 6)
        ValueTreeViewDepartment = .List(.ListIndex, 7)
        ValueTreeEtityId = .List(.ListIndex, 8)
        If ValueTreeViewDepartment <> "" Then
            Call F_GetKPI
        End If

        ThietLapKhoanID = .List(.ListIndex, 9)

        cbbKPIKhoan.Text = FontConverter(.List(.ListIndex, 10), 0, 1)
        ThietLapKhoanNhanVienID = .List(.ListIndex, 11)
        NhanVienID = .List(.ListIndex, 12)
        Call F_CommissionTarget
    End With
    Load_Data_Listview_ThongTinKhoan_TheoBac

    If ThietLapKhoanID > 0 Then
        btnCreateGrade.Visible = True
        cbbDeletBac.Visible = True
    Else
        btnCreateGrade.Visible = False
        cbbDeletBac.Visible = False
    End If
End Sub

Private Sub UserForm_Initialize()
    Call F_GetAllDepartment
    Call F_ViewHeaderEmployee
    lvApplicableStaff.Font.Charset = VIETNAMESE_CHARSET

    ListViewThongTinKhoan.Font.Charset = VIETNAMESE_CHARSET
    Call F_HeaderCInfoBy
    Set NhanVienApDung = New Collection

    colorError = &HFF&
    colorDefault = &H80000000
End Sub

' Cap nhat
Private Sub btnCapNhat_Click()
    Dim PhongBanID As Integer
    Dim ViTriID As Variant
    Dim TinhTheoPhongBanID As String
    Dim CongViecID As Variant
    Dim NgayApDung As String
    Dim NgayHetHan As String
    Dim DoiTuongID As String
    Dim DoiTuong As String
    Dim TinhTheoPhongBan As String

    PhongBanID = cbbDepartment.value
    If cbbJobTitle.Text <> "" Then
        ViTriID = cbbJobTitle.value
    End If

    TinhTheoPhongBanID = ValueTreeViewDepartment
    TinhTheoPhongBan = InputDepartment.Caption
    If cbbKPIKhoan.Text <> "" Then
        CongViecID = cbbKPIKhoan.value
    End If

    NgayApDung = Format(txtNgayApDung.Text, "yyyy-mm-dd")
    NgayHetHan = Format(txtNgayHetHan.Text, "yyyy-mm-dd")

    DoiTuongID = ValueTreeEtityId
    DoiTuong = inputEntity.Caption

    If F_CheckInfoEmpty(PhongBanID, ViTriID, TinhTheoPhongBanID, CongViecID, NgayApDung, DoiTuongID) = True Then
     Exit Sub
    End If

    Dim dbConn As Object
    Dim Rs As Object

    Set dbConn = ConnectToDatabase

    If ThietLapKhoanID > 0 Then
        Query = "UPDATE CV_ThietLapKhoan Set PhongBanID = " & PhongBanID & ", ViTriID = " & ViTriID & ", TinhTheoPhongBanID = " & TinhTheoPhongBanID & ", " & _
        "TinhTheoPhongBan = N'" & TinhTheoPhongBan & "', CongViecID = " & CongViecID & ", NgayApDung = '" & NgayApDung & "', NgayHetHan = '" & NgayHetHan & "' WHERE ThietLapKhoanID = " & ThietLapKhoanID

        Set Rs = dbConn.Execute(Query)

        Query = "UPDATE CV_ThietLapKhoan_NhanVien Set ThietLapKhoanID = " & ThietLapKhoanID & ", NhanVienID = " & NhanVienID & ", DoiTuongID = '" & DoiTuongID & "', DoiTuong = N'" & DoiTuong & "', " & _
        "ChiTieuKhoan = '" & ChiTieuKhoan & "', LuongThuongDuKien = " & ThuongKhoan & " WHERE ThietLapKhoanNhanVienID = " & ThietLapKhoanNhanVienID

        Set Rs = dbConn.Execute(Query)
    Else
        Query = "INSERT INTO CV_ThietLapKhoan(PhongBanID, ViTriID, TinhTheoPhongBanID,TinhTheoPhongBan, CongViecID, NgayApDung, NgayHetHan) " & _
        "Select " & PhongBanID & ", " & ViTriID & ", '" & TinhTheoPhongBanID & "',N'" & TinhTheoPhongBan & "', " & CongViecID & ", '" & NgayApDung & "', '" & NgayHetHan & "'" & _
        "WHERE Not EXISTS(Select ThietLapKhoanID from CV_ThietLapKhoan where ViTriID = " & ViTriID & " And CongViecID = " & CongViecID & ")"

        Set Rs = dbConn.Execute(Query)
        Query = "INSERT INTO CV_ThietLapKhoan_NhanVien(ThietLapKhoanID, NhanVienID, DoiTuongID,DoiTuong,ChiTieuKhoan,LuongThuongDuKien) " & _
        "Select ThietLapKhoanID, " & NhanVienID & ", '" & DoiTuongID & "',N'" & DoiTuong & "', '" & ChiTieuKhoan & "', '" & ThuongKhoan & "' FROM CV_ThietLapKhoan WHERE ViTriID = " & ViTriID & " And CongViecID = " & CongViecID & "" & _
        " And Not EXISTS(Select ThietLapKhoanNhanVienID FROM CV_ThietLapKhoan_NhanVien WHERE NhanVienID = " & NhanVienID & " And ThietLapKhoanID in(Select ThietLapKhoanID FROM CV_ThietLapKhoan WHERE ViTriID = " & ViTriID & " And CongViecID = " & CongViecID & ") )"
    End If

    Set Rs = dbConn.Execute(Query)

    Set Rs = Nothing
    Query = ""
    Call F_Reset
    Call F_GetObjectiveID
    Call F_GetPostionID
    Call CloseDatabaseConnection(dbConn)
    ThongBao_ThanhCong
End Sub

' Xu ly khi click vao phong ban
Private Sub cbbDepartment_Click()
    Call F_Reset
    If F_CheckDepartmentID Then
     Exit Sub
    End If

    Call F_GetObjectiveID
    Call F_GetPostionID
End Sub

'Dong form
Private Sub btDong_Click()
    Unload Me
End Sub

' Xu ly kich vao vi tri nhan khoan
Private Sub cbbJobTitle_Click()
    Call F_ViewBodyEmployye
    Call F_CommissionTarget
End Sub

' Xu ly khi click kpi Khoan
Private Sub cbbKPKhoan_Click()
    Call F_CommissionTarget
End Sub

' Xu ly chon ngay ap dung
Private Sub cmdChonNgay_Click()
    nameCalendar = "NgayApDung"
    KeyDate = 0
    Call AdvancedCalendar

End Sub

' Xu ly chon ngay het han
Private Sub cmdChonNgayHetHan_Click()
    nameCalendar = "NgayHetHan"
    KeyDate = 0

    Call AdvancedCalendar

End Sub

'Xuy ly chon doi tuong
Private Sub cbbKhoanTheoNhanSu_Click()
    Call F_OpenEtity
End Sub

'Xu ly khi click vao tinh theo phong ban
Private Sub btnDepartment_Click()
    F_GetTreviewDepartment

    If ValueTreeViewDepartment <> "" Then
        Call F_GetKPI
    End If
End Sub

' Kiem tra da nhap day du thong tin chua
Private Function F_CheckInfoEmpty(PhongBanID As Variant, ViTriID As Variant, TinhTheoPhongBanID As String, CongViecID As Variant, NgayApDung As String, DoiTuongID As String) As Boolean
    TieuDe = "Bos Xin thông báo"

    If IsNull(PhongBanID) Or IsEmpty(PhongBanID) Then
        NoiDung = "Ch" & ChrW(7885) & "n ph" & ChrW(242) & "ng ban"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

        cbbDepartment.BorderColor = colorError
        F_CheckInfoEmpty = True

     Exit Function
    Else
        cbbDepartment.BorderColor = colorDefault
    End If

    If IsNull(ViTriID) Or IsEmpty(ViTriID) Then
        NoiDung = "Ch" & ChrW(7885) & "n v" & ChrW(7883) & " tr" & ChrW(237)

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

        cbbJobTitle.BorderColor = colorError
        F_CheckInfoEmpty = True

     Exit Function
    Else
        cbbJobTitle.BorderColor = colorDefault
    End If

    If TinhTheoPhongBanID = "" Then
        NoiDung = "Ch" & ChrW(7885) & "n ph" & ChrW(242) & "ng ban t" & ChrW(237) & "nh theo"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

        InputDepartment.BorderColor = colorError
        F_CheckInfoEmpty = True
     Exit Function
    Else
        InputDepartment.BorderColor = colorDefault
    End If

    If IsNull(CongViecID) Or IsEmpty(CongViecID) Then
        NoiDung = "Ch" & ChrW(7885) & "n KPI kho" & ChrW(225) & "n"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

        cbbKPIKhoan.BorderColor = colorError
        F_CheckInfoEmpty = True

     Exit Function
    Else
        cbbKPIKhoan.BorderColor = colorDefault
    End If

    If IsNull(NhanVienID) Or IsEmpty(NhanVienID) Then
        NoiDung = "Ch" & ChrW(7885) & "n nh" & ChrW(226) & "n vi" & ChrW(234) & "n " & ChrW(225) & "p d" & ChrW(7909) & "ng"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

        F_CheckInfoEmpty = True
     Exit Function
    End If

    If NgayApDung = "" Then
        NoiDung = "Ch" & ChrW(7885) & "n ng" & ChrW(224) & "y " & ChrW(225) & "p d" & ChrW(7909) & "ng"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

        txtNgayApDung.BorderColor = colorError
        F_CheckInfoEmpty = True

     Exit Function
    Else
        txtNgayApDung.BorderColor = colorDefault
    End If

    If DoiTuongID = "" Then
        NoiDung = "Ch" & ChrW(7885) & "n " & ChrW(273) & ChrW(7889) & "i t" & ChrW(432) & ChrW(7907) & "ng " & ChrW(225) & "p d" & ChrW(7909) & "ng"

        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0

        inputEntity.BorderColor = colorError
        F_CheckInfoEmpty = True

     Exit Function
    Else
        inputEntity.BorderColor = colorDefault
    End If

    F_CheckInfoEmpty = False
End Function

' Kiem xem da chon phong ban chua
Private Function F_CheckDepartmentID() As Boolean
    Dim DepartmentID As Integer
    DepartmentID = cbbDepartment.value

    If (DepartmentID <= 0) Then
        F_CheckDepartmentID = True
        cbbDepartment.BorderColor = &HFF&
        CreateObject("WScript.Shell").Popup "Ch" & ChrW(7885) & "n Phong ban", , "BOS Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 16
     Exit Function
    Else
        cbbDepartment.BorderColor = &H80000000
        F_CheckDepartmentID = False
    End If
End Function

' Lay danh sách chi tieu theo phong ban
Private Function F_GetObjectiveID()

    Dim DepartmentID As Integer
    Dim Query As String
    Dim dbConn As Object

    DepartmentID = cbbDepartment.value
    Set dbConn = ConnectToDatabase

    Query = "Select VT.Mavitri,VT.TenViTri,isnull(TLK_NV.ChiTieuKhoan,0) ChiTieuKhoan , isnull(TLK_NV.DoiTuong,'') DoiTuong,TLK.TinhTheoPhongBan,TLK.NgayApDung,TLK.NgayHetHan,  " & _
    " TLK.TinhTheoPhongBanID,isnull(TLK_NV.DoiTuongID,'') DoiTuongID,TLK.ThietLapKhoanID, isnull(CV.TenCongViec,'') As TenCongViec, isnull(TLK_NV.ThietLapKhoanNhanVienID,0) ThietLapKhoanNhanVienID, isnull(TLK_NV.NhanVienID ,0) As NhanVienID " & _
    "FROM CV_ThietLapKhoan TLK " & _
    "LEFT JOIN CV_ThietLapKhoan_NhanVien TLK_NV on  TLK.ThietLapKhoanID = TLK_NV.ThietLapKhoanID " & _
    "LEFT JOIN DM_ViTri VT on VT.ViTriID = TLK.ViTriID " & _
    "LEFT JOIN CV_CongViec CV on CV.CongViecID = TLK.CongViecID " & _
    "WHERE TLK.PhongBanID = " & DepartmentID & ""

    Call ViewListBox(Query, TableObjective, dbConn)
End Function

' lay tat ca danh sach phong ban
Private Function F_GetAllDepartment()
    Dim dbConn As Object
    Dim Query As String

    Set dbConn = ConnectToDatabase

    Query = "Select PhongBanID, TenPhongBan From PhongBan"

    Call ViewListBox(Query, cbbDepartment, dbConn)

    Call CloseDatabaseConnection(dbConn)

    'Set mac dinh khi chon phong ban la cong ty

    With cbbDepartment
        .Text = .List(0, 1)

        If .Text <> "" Then
            ' Goi danh sach chi tieu, danh vi tri theo ID
            Call F_GetObjectiveID
            Call F_GetPostionID
        End If
    End With

End Function

' Lay vi tri theo PhongBanID
Private Function F_GetPostionID()
    Dim DepartmentID As Integer
    Dim Query As String
    Dim dbConn As Object

    DepartmentID = cbbDepartment.value
    Set dbConn = ConnectToDatabase

    Query = "Select isNull(ViTriID,0) As ViTriID, isNull(TenViTri,'') As TenViTri FROM DM_ViTri WHERE PhongBanID in(Select valu FROM LayDonViCon(" & DepartmentID & "))"

    Call ViewListBox(Query, cbbJobTitle, dbConn)
    Call CloseDatabaseConnection(dbConn)
End Function

' Hien thi danh sach phong ban theo treeview
Private Function F_GetTreviewDepartment()
    Dim dbConn As Object
    Dim Query As String
    Dim Rs As Object

    Set dbConn = ConnectToDatabase

    Query = "Select PhongBanID As id,TenPhongBan As name, PhongBanChaID As parent_id from PhongBan"
    Set Rs = dbConn.Execute(Query)

    If Not Rs.EOF And Not Rs.BOF Then
        DataTreeView = Rs.GetRows()
        NameTreeView = "TreeViewDepartemnt"
        Form_TreeView.Show
    End If
    Call CloseDatabaseConnection(dbConn)
    Set Rs = Nothing

End Function

' Lay danh sach KPI khoan theo phong ban da lua chon o tinh theo phong ban
Private Function F_GetKPI()
    Dim dbConn As Object
    Dim Query As String

    Set dbConn = ConnectToDatabase

    Query = "Select CongViecID, TenCongViec from ( " & _
    "Select distinct CV_PhongBan_CongViec.CongViecID, DM_ViTri.PhongBanID ,CV_CongViec.TenCongViec " & _
    "FROM CV_PhongBan_CongViec inner join DM_ViTri on CV_PhongBan_CongViec.ViTriID = DM_ViTri.ViTriID inner Join CV_CongViec On CV_PhongBan_CongViec.CongViecID = CV_CongViec.CongViecID " & _
    "WHERE PhanLoaiCongViec = 'CV' And DM_ViTri.PhongBanID in (Select Value from string_split('" & ValueTreeViewDepartment & "',',')) ) AA Group by CongViecID,TenCongViec " & _
    "having Count(PhongBanID) = (Select Count(Value) from string_split('" & ValueTreeViewDepartment & "',','))"

    Call ViewListBox(Query, cbbKPIKhoan, dbConn)
    Call CloseDatabaseConnection(dbConn)
End Function

' tao header nhan vien ap dung
Private Function F_ViewHeaderEmployee()
    Dim EmployeeName As String
    Dim PerformanceTarget As String
    Dim AnticipatedBonusCommission As String

    EmployeeName = "T" & ChrW(234) & "n nh" & ChrW(226) & "n vi" & ChrW(234) & "n"
    PerformanceTarget = "Ch" & ChrW(7881) & " tiêu khoán"
    AnticipatedBonusCommission = "L" & ChrW(432) & ChrW(417) & "ng - th" & ChrW( _
    432) & ChrW(7903) & "ng khoán d" & ChrW(7921) & " ki" & ChrW(7871) & "n"

    With lvApplicableStaff.ColumnHeaders
        .Add , , UniToWindows1258(EmployeeName), 150
        .Add , , UniToWindows1258(PerformanceTarget), 150
        .Add , , UniToWindows1258(AnticipatedBonusCommission), 160
        .Add , , NhanVienID, 0
    End With
End Function

'hien thi body nhan vien ap dung
Private Function F_ViewBodyEmployye()
    Dim dbConn As Object
    Dim Query, Query2 As String

    Set dbConn = ConnectToDatabase

    Query = "Select ho + ' ' + Ten As HoTen, ISNULL(TLKNV.ChiTieuKhoan,0) As ChiTieuKhoan, ISNULL(TLKNV.LuongThuongDuKien,0) As LuongThuongKhoan, NV.NhanVienID FROM NS_NhanVien NV " & _
    "LEFT JOIN CV_ThietLapKhoan_NhanVien TLKNV ON NV.NhanVienID = TLKNV.NhanVienID " & _
    "WHERE NV.ViTriID = " & cbbJobTitle.value

    Query2 = "Select NV.NhanVienID, ho + ' ' + Ten As HoTen FROM NS_NhanVien NV " & _
    "LEFT JOIN CV_ThietLapKhoan_NhanVien TLKNV ON NV.NhanVienID = TLKNV.NhanVienID " & _
    "WHERE NV.ViTriID = " & cbbJobTitle.value

    Call ViewListBox(Query2, cbbNhanVienKhoan, dbConn)
    Call F_ViewTableEmployye(Query, dbConn)

    Call CloseDatabaseConnection(dbConn)
End Function

Function F_ViewTableEmployye(Query, dbConn)
    On Error Resume Next
    If Not dbConn Is Nothing Then
        Dim Rs As Object
        Dim i As Integer
        Set Rs = dbConn.Execute(Query)
        With lvApplicableStaff
            For i = .ListItems.Count To 1 Step -1
                .ListItems.Remove i
            Next i
            Set NhanVienApDung = New Collection
        End With

        Do Until Rs.EOF
            Dim foundItem As Boolean
            foundItem = False
            Dim existingItem As Variant

            For Each existingItem In NhanVienApDung
                If existingItem = Rs.Fields("NhanVienID").value Then
                    foundItem = True
                 Exit For
                End If
            Next existingItem

            If Not foundItem Then
                Dim ListItem As ListItem
                Set ListItem = lvApplicableStaff.ListItems.Add(, , UniToWindows1258(Rs.Fields("HoTen").value))
                ListItem.SubItems(1) = UniToWindows1258(Rs.Fields("ChiTieuKhoan").value)
                ListItem.SubItems(2) = FormatNumber(Rs.Fields("LuongThuongKhoan").value)
                ListItem.SubItems(3) = Rs.Fields("NhanVienID").value
                NhanVienApDung.Add Rs.Fields("NhanVienID").value

                If Rs.Fields("LuongThuongKhoan").value > 0 Then
                    ListItem.Checked = True
                End If
            End If
            Rs.MoveNext
        Loop
    Else
        MsgBox "Mat Ket noi csdl"
    End If
End Function

'Hien thi treeview doi tuong
Private Function F_ViewEtity()

    Dim dbConn As Object
    Dim Query As String

    Set dbConn = ConnectToDatabase

    Query = "Select ID, Ten, ChaID, TenCha, PL from BOS_DataTree('PB_VT_NV','" & ValueTreeViewDepartment & "') order by STT"
    Set Rs = dbConn.Execute(Query)

    If Not Rs.EOF And Not Rs.BOF Then
        DataTreeView = Rs.GetRows()
        NameTreeView = "view_etity"
        Form_TreeView.Show
    End If
    Call CloseDatabaseConnection(dbConn)
    Set Rs = Nothing
End Function

Private Function F_OpenEtity()
    NameTreeView = "view_etity"
    Call F_ViewEtity
End Function

Private Function F_HeaderCInfoBy()
    Dim BK As String
    Dim HS As String
    Dim GKT As String
    Dim Note As String

    BK = "B" & ChrW(7853) & "c kho" & ChrW(225) & "n"
    HS = "H" & ChrW(7879) & " s" & ChrW(7889)
    GKT = "Gi" & ChrW(7843) & "i kho" & ChrW(225) & "n t" & ChrW(7915)
    Note = "Ghi ch" & ChrW(250)

    With ListViewThongTinKhoan
        With .ColumnHeaders
            .Add , , UniToWindows1258(BK), 150
            .Add , , UniToWindows1258(HS), 150
            .Add , , UniToWindows1258(GKT), 160
            .Add , , UniToWindows1258(Note), 160
            .Add , , "ThongTinGiaoKhoanID", 100
        End With
    End With
End Function

' Tinh Chi Tieu khoan theo thiet lap
Public Function F_CommissionTarget()
    Dim CongViecID As Integer
    Dim DoiTuong As String

    If cbbKPIKhoan.Text <> "" Then
        CongViecID = cbbKPIKhoan.value
    End If

    DoiTuong = ValueTreeEtityId

    If NhanVienID = 0 Or CongViecID = 0 Or DoiTuong = "" Then
        ChiTieuKhoan = 0
        ThuongKhoan = 0
     Exit Function
    End If

    Dim dbConn As Object
    Dim Query As String
    Dim Rs As Object
    Dim Data As Variant

    Set dbConn = ConnectToDatabase

    Query = "Select isNull(ChiTieuKhoan,0) As ChiTieuKhoan,isNull(ThuongKhoan, 0) As ThuongKhoan FROM CV_TinhThuongkhoan_TheoThietLap(" & NhanVienID & ",'" & DoiTuong & "','" & CongViecID & "')"

    Set Rs = dbConn.Execute(Query)

    If Not Rs.EOF And Not Rs.BOF Then
        Data = Rs.GetRows()

        ChiTieuKhoan = Data(0, 0)
        ThuongKhoan = Data(1, 0)
    End If
End Function

' Xoa du lieu dua bien ve ban dau
Private Function F_Reset()

    DataTreeView = ""
    NameTreeView = ""
    ValueTreeViewDepartment = ""
    ValueTreeEtityId = ""
    nameCalendar = ""

    Set NhanVienApDung = New Collection
    NhanVienID = 0
    ChiTieuKhoan = 0
    ThuongKhoan = 0
    cbbJobTitle = ""
    InputDepartment = ""
    cbbKPIKhoan = ""
    txtNgayApDung = ""
    txtNgayHetHan = ""
    inputEntity = ""
    lvApplicableStaff.ListItems.Clear

    ThietLapKhoanID = 0
    ThietLapKhoanNhanVienID = 0
    ListViewThongTinKhoan.ListItems.Clear
    ThietLapKhoanID_TheoBac = 0
    cbbNhanVienKhoan = ""
End Function

Public Function F_DELETE()
    If ThietLapKhoanID = 0 Then
        NoiDung = "Ch" & ChrW(7885) & "n thi" & ChrW(7871) & "t l" & ChrW(7853) & "p kho" & ChrW(225) & "n c" & ChrW(7845) & "n x" & ChrW(243) & "a"
        Application.Assistant.DoAlert TieuDe, NoiDung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
     Exit Function
    End If

    Dim msgValue As VbMsgBoxResult
    msgValue = Application.Assistant.DoAlert(UniConvert("Carnh baso", "Telex"), UniConvert("Bajn muoosn xosa duwx lieeju?", "Telex"), msoAlertButtonYesNo, msoAlertIconWarning, 0, 0, 1)

    If msgValue = vbYes Then

        Query = "DELETE FROM CV_ThietLapKhoan WHERE ThietLapKhoanID = " & ThietLapKhoanID & "; " & _
        "DELETE FROM CV_ThietLapKhoan_NhanVien WHERE ThietLapKhoanID = " & ThietLapKhoanID & "; " & _
        "DELETE FROM CV_ThietLapKhoan_TheoBac WHERE ThietLapKhoanID = " & ThietLapKhoanID

        Dim dbConn As Object

        Dim Rs As Object

        Set dbConn = ConnectToDatabase

        Set Rs = dbConn.Execute(Query)

        Set Rs = Nothing
        Call CloseDatabaseConnection(dbConn)
        Query = ""
        Set Rs = Nothing

        F_Reset
        Call F_GetObjectiveID
        Call F_GetPostionID
    End If
End Function

Sub Load_Data_Listview_ThongTinKhoan_TheoBac()
    ListViewThongTinKhoan.ListItems.Clear
    Dim TT_BacKhoan As Variant
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset

    StrCnn = KetNoiMayChu_KhachHang
    Dim SQLStr As String
    Dim CongViecID As Integer
    Dim ViTriID As Integer
    'ThietLapKhoanID = 5
    ViTriID = cbbJobTitle.List(cbbJobTitle.ListIndex, 0)
    CongViecID = cbbKPIKhoan.List(cbbKPIKhoan.ListIndex, 0)
    SQLStr = "Select TenBac, Heso,GiaiKhoanTu, GhiChu, ThietLapKhoan_TheoBacID  " & _
    "from CV_ThietLapKhoan_TheoBac where ThietLapKhoanID in (Select top 1 ThietLapKhoanID from CV_thietLapKhoan where CongViecID = " & CongViecID & " And VitriID = " & ViTriID & ") " & _
    " order by GiaiKhoanTu asc ;"

    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Set Rs = New ADODB.Recordset
    Rs.Open SQLStr, Cn, adOpenStatic

    If Not Rs.EOF Then
        TT_BacKhoan = Rs.GetRows()
    Else
     Exit Sub
    End If

    Dim i As Integer
    For i = 0 To UBound(TT_BacKhoan, 2)
        With ListViewThongTinKhoan
            With .ListItems.Add
                .Text = UniToWindows1258(TT_BacKhoan(0, i))
                .SubItems(1) = TT_BacKhoan(1, i)
                .SubItems(2) = FormatNumber(TT_BacKhoan(2, i))
                .SubItems(3) = UniToWindows1258(TT_BacKhoan(3, i))
                .SubItems(4) = TT_BacKhoan(4, i)
            End With
        End With
    Next i

    Cn.Close
    Set Cn = Nothing

    F_CommissionTarget
End Sub

Sub CapNhat_ThietLapKhoan()
    Dim PhongBanID As Integer
    Dim ViTriID As Variant
    Dim TinhTheoPhongBanID As String
    Dim CongViecID As Variant
    Dim NgayApDung As String
    Dim NgayHetHan As String
    Dim DoiTuongID As String
    Dim DoiTuong As String
    Dim TinhTheoPhongBan As String

    PhongBanID = cbbDepartment.value
    If cbbJobTitle.Text <> "" Then
        ViTriID = cbbJobTitle.value
    End If

    TinhTheoPhongBanID = ValueTreeViewDepartment
    TinhTheoPhongBan = InputDepartment.Caption
    If cbbKPIKhoan.Text <> "" Then
        CongViecID = cbbKPIKhoan.value
    End If

    NgayApDung = Format(txtNgayApDung.Text, "yyyy-mm-dd")
    NgayHetHan = Format(txtNgayHetHan.Text, "yyyy-mm-dd")

    DoiTuongID = ValueTreeEtityId
    DoiTuong = inputEntity.Caption

    If F_CheckInfoEmpty(PhongBanID, ViTriID, TinhTheoPhongBanID, CongViecID, NgayApDung, DoiTuongID) = True Then
     Exit Sub
    End If

    Dim dbConn As Object
    Dim Query As String
    Dim Rs As Object

    Set dbConn = ConnectToDatabase

    Query = "INSERT INTO CV_ThietLapKhoan(PhongBanID, ViTriID, TinhTheoPhongBanID,TinhTheoPhongBan, CongViecID, NgayApDung, NgayHetHan) " & _
    "Select " & PhongBanID & ", " & ViTriID & ", '" & TinhTheoPhongBanID & "',N'" & TinhTheoPhongBan & "', " & CongViecID & ", '" & NgayApDung & "', '" & NgayHetHan & "'" & _
    "WHERE Not EXISTS(Select ThietLapKhoanID from CV_ThietLapKhoan where ViTriID = " & ViTriID & " And CongViecID = " & CongViecID & ")"

    Set Rs = dbConn.Execute(Query)
    Set Rs = Nothing
    Call CloseDatabaseConnection(dbConn)

End Sub



