Attribute VB_Name = "M_DB_SS_KH_Cua_NVKD"


Sub HienDuLieu_NVKD()
    Dim MaNhanVien As String
    Dim Nam As Integer
    Dim Query As String
    With Sheet34
        MaNhanVien = .Range("E5").Value
        Nam = .Range("G5").Value
    End With

    Query = "Select Isnull((Select top 1 NhanvienID from Ns_NhanVien where MaNhanVien = N'" & MaNhanVien & "'),9999 )"
    Call GenerateQueryAndCallViewSheet("BaoCaoDoanhThu_NhanVienKD_TheoNgay", Nam, Sheet35, Query)
    
    Sheet34.Select
End Sub

Sub VeBieuDo_SoSanhCua_NVKD()
    BatLimit
    Set wSheet = Sheet34
    With wSheet
        .Select
        
        If .cbbDVKD.ListCount <= 0 Or _
           .cbbNam.ListCount <= 0 _
        Then
            Call ComboBox_DonVi_Nam
        End If
        
        If .cbbNV.ListCount <= 0 Or _
           .cbbNam.ListCount <= 0 _
        Then
            Call ComboBox_SoSanhKeHoachcua_NVKD
        End If
        
    End With
    Set wSheet = Nothing
    F_R_DATA
    HienDuLieu_NVKD
    
    Set wSheet = Sheet35
    With wSheet
       .Select
        
       Dim dongCuoi As Integer
      
       dongCuoi = tinhdongcuoi("BJ12:BJ510")
        If dongCuoi > 31 Then
           .ListObjects("Table58").Resize Range("$BJ$30:$BO$" & dongCuoi)
        End If
        
   End With
    
    With Sheet34
        .Select
        
        ActiveSheet.ChartObjects("Chart 6").Activate
        ActiveChart.PlotArea.Select
        ActiveChart.SetSourceData Source:=Sheet35.Range("$AS$153:$AW$154")
 
   End With
   
    Call VeBieuDoLuyKeDoanhThuNVKD
    ActiveWorkbook.RefreshAll

    TatLimit
    ThongBao_ThanhCong
End Sub

Function VeBieuDoLuyKeDoanhThuNVKD()
    Dim PhamViResize As String

    dongCuoi = Sheet35.Range("BO29").Value

    PhamViResize = "BJ30:BO" & 30 + dongCuoi

    Call UpdateChartDataRange(Sheet34, "Chart 16", Sheet35, PhamViResize)
    ActiveWorkbook.RefreshAll
End Function

Sub ComboBox_SoSanhKeHoachcua_NVKD()
    With Sheet34
        Dim dbConn As Object
        Dim Query As String
        
        'Mo ket noi csdl
        Set dbConn = ConnectToDatabase
        
        'Hien thi danh sach Ma nhan vien
        'Voi KhoiID = 2 là khoi kinh doanh
        'LinhVucID = 1 là nhân viên bán hang
        Query = "Select Ho + ' ' + Ten As HoTen, MaNhanVien from NS_NhanVien inner join PhongBan on NS_NhanVien.PhongBanID = PhongBan.PhongBanID " & _
        "where PhongBan.KhoiID = 2 And PhongBan.LinhVucID = 1"
        
        Call ViewListBox(Query, .cbbNV, dbConn)
        
        'Set default Ma Nhan vien dau tien
        With .cbbNV
            .Text = .List(0, 1)
        End With
        
        .Range("E5") = .cbbNV.Value
        
        'Hien thi danh sach nam bao cao
        Query = "Select Distinct Year(COnvert(date, NgayHachToan)) Nam from KD_DonHang where NgayHachToan is Not null order by Year(COnvert(date, NgayHachToan))"
    
        Call ViewListBox(Query, .cbbNam, dbConn)
        
        'Set default Nam bao cao dau tien
        
        With .cbbNam
            .Text = .List(.ListCount - 1, 0)
        End With
        
        .Range("G5") = .cbbNam.Value

        'Dong ket noi csdl
        Call CloseDatabaseConnection(dbConn)
    End With
End Sub

Sub SS_KH_NVKD_Chon_dashboard_SSKHNVKD()
    Sheet34.Select
End Sub

Sub SpinDown_SSKHNVKD()
     Set wSheet = Sheet34
    With wSheet
        .Select

        If .Range("X102") < 40 Then
            .Range("X102") = .Range("X102") + 1
        End If
        .TextBox2.Value = .Range("X102")
    End With
    
     Set wSheet = Nothing
End Sub

Sub SpinUp_SSKHNVKD()
   Set wSheet = Sheet34
    With wSheet
        .Select

        If .Range("X102") > 0 Then
            .Range("X102") = .Range("X102") - 1
        End If

        .TextBox2.Value = .Range("X102")
    End With
       Set wSheet = Nothing
End Sub

Sub chon_tab_ngay_Data_SSKHNVKD()
    Set wSheet = Sheet35
    
    With wSheet
        .Select
        .Range("A1").Select
    End With
    
    Set wSheet = Nothing
End Sub

Sub chon_tab_tuan_Data_SSKHNVKD()
    Set wSheet = Sheet35
    
    With wSheet
        .Select
        .Range("A1").Select
        ActiveWindow.SmallScroll ToRight:=22
    End With
    
    Set wSheet = Nothing
End Sub

Sub chon_tab_thang_Data_SSKHNVKD()
    Set wSheet = Sheet35
    
    With wSheet
    .Select
    .Range("A1").Select
     ActiveWindow.SmallScroll ToRight:=41
    End With
    
End Sub

Sub chon_tab_nam_Data_SSKHNVKD()
    Set wSheet = Sheet35
    
    With wSheet
    .Select
    .Range("A1").Select
    ActiveWindow.SmallScroll ToRight:=59
    
    End With
End Sub





