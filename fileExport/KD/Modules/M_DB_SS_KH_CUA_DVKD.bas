Attribute VB_Name = "M_DB_SS_KH_CUA_DVKD"
Sub HienDuLieu_DVKD()
    Dim TenPhongBan As String
    Dim Nam As Integer
    Dim Query As String

    With Sheet32
        .Select
        TenPhongBan = .Range("E5").Value
        Nam = .Range("F5").Value
    End With
    
    Query = "Select Isnull((Select top 1 PhongBanID from PhongBan where TenPhongBan = N'" & TenPhongBan & "'),9999 )"
    Call GenerateQueryAndCallViewSheet("BaoCaoDoanhThu_DVKD_TheoNgay", Nam, Sheet23, Query)
    With Sheet32
        .Select
        .Range("A1").Select
    End With
End Sub

Sub VeBieuDo_SoSanhKhachHang_DVKD()
    BatLimit
    Set wSheet = Sheet32
    With wSheet
    .Select
    
    If .cbbDVKD.ListCount <= 0 Or _
       .cbbNam.ListCount <= 0 _
    Then
        Call ComboBox_DonVi_Nam
    End If
    End With
     Set wSheet = Nothing
    
    F_R_DATA
    HienDuLieu_DVKD
    
    With Sheet23
       .Select
        
       Dim dongCuoi As Integer
      
       dongCuoi = Abs(tinhdongcuoi("BJ31:BJ110"))
          If dongCuoi < 31 Then
            Sheet32.Select
            Exit Sub
        End If
        .ListObjects("Table57").Resize Range("$BJ$30:$BO$" & dongCuoi)
   End With
    
   Call VeBieuDoLuyKeDoanhThuDV
   Sheet32.Select
    ActiveWorkbook.RefreshAll

   
    TatLimit
     ThongBao_ThanhCong
End Sub

Function VeBieuDoLuyKeDoanhThuDV()
    Dim PhamViResize As String

    dongCuoi = Sheet23.Range("BP29").Value

    PhamViResize = "BJ30:BO" & 30 + dongCuoi

    Call UpdateChartDataRange(Sheet32, "Chart 6", Sheet23, PhamViResize)
    ActiveWorkbook.RefreshAll
End Function

Public Sub ComboBox_DonVi_Nam()
    Dim dbConn As Object
    Dim Query As String

    Set dbConn = ConnectToDatabase
    With Sheet32
       ' hien thi don vi bao cao
       Query = "SELECT 'Công ty' AS TenPhongBan UNION Select TenPhongBan  from PhongBan where KhoiID = 2 And LinhVucID = 1"
       Call ViewListBox(Query, .cbbDVKD, dbConn)
    
       Call ViewListBox(Query, Sheet34.cbbDVKD, dbConn)
       With .cbbDVKD
            .Text = .List(0, 0)
       End With
    
       ' Nam bao cao
       Query = "Select Distinct Year(COnvert(date, NgayHachToan)) Nam from KD_DonHang where NgayHachToan is Not null order by Year(COnvert(date, NgayHachToan))"
       Call ViewListBox(Query, .cbbNam, dbConn)
       
       With .cbbNam
            .Text = .List(.ListCount - 1, 0)
       End With
    End With
    
    Call CloseDatabaseConnection(dbConn)
End Sub

Sub SS_KH_DVKD_Chon_data()
    ThisWorkbook.Sheets("Data SS KH DVKD").Select
End Sub

Sub SS_KH_DVKD_Chon_dashboard()
    ThisWorkbook.Sheets("DB SS KH cua DVKD").Select
End Sub

Public Sub SpinDown_SoSanhKhachHang_DVKD()
    With Sheet23
        .Select

        If .Range("X102") < 40 Then
            .Range("X102") = .Range("X102") + 1
        End If

        .TextBox1.Value = .Range("X102")
    End With
End Sub

Public Sub SpinUp_SoSanhKhachHang_DVKD()
    With Sheet23
        .Select

        If .Range("X102") > 0 Then
            .Range("X102") = .Range("X102") - 1
        End If

        .TextBox1.Value = .Range("X102")
    End With

End Sub

Sub chon_tab_ngay()
    With Sheet23
        .Select
        .Range("A1").Select
    End With
End Sub

Sub chon_tab_tuan()
    With Sheet23
        .Select
        .Range("A1").Select
        ActiveWindow.SmallScroll ToRight:=22
    End With
End Sub

Sub chon_tab_thang()
    With Sheet23
        .Select
        .Range("A1").Select
        ActiveWindow.SmallScroll ToRight:=40
    End With
End Sub

Sub chon_tab_nam()
    With Sheet23
        .Select
        .Range("A1").Select
        ActiveWindow.SmallScroll ToRight:=59
    End With
End Sub






