Attribute VB_Name = "M_DB_SS_KH_Cua_KH"
Sub HienDuLieu_KhachHang()
    Dim Nam As Integer
    Dim Query As String
    
    With Sheet37
        .Select
        MaKhachHang = .Range("J7").Value
        Nam = .Range("L7").Value
        
    End With
    
    Query = "Select Isnull((Select top 1 KhachHangID from KH_KhachHang where MaKhachHang = N'" & MaKhachHang & "'),9999 )"
    Call GenerateQueryAndCallViewSheet("BaoCaoDoanhThu_KhachHang_TheoNgay", Nam, Sheet37, Query)
    
'    Sheet36.Select
End Sub

Public Sub HienThiDoanhThuNgayTheo_KH()
    BatLimit
    Set wSheet = Sheet37
    With wSheet
    .Select
    
    If .cbbKH.ListCount <= 0 Or _
       .cbbNam.ListCount <= 0 _
    Then
        ComboBox_KhachHang_Nam
    End If
    End With
     Set wSheet = Nothing
    Call HienDuLieu_KhachHang
    
    TatLimit
    ThongBao_ThanhCong
End Sub

Sub ComboBox_KhachHang_Nam()

    With Sheet37
        Dim dbConn As Object
        Dim Query As String
        Dim NgungTheoDoi As String
 
        ' Mo Ket noi csdl
        Set dbConn = ConnectToDatabase
        
        ' Hien thi danh sach ma khach hang dang xem
        
        NgungTheoDoi = "False" ' True La ngung theo doi, False La con theo doi
        Query = "Select MaKhachHang from KH_KhachHang WHERE NgungTheoDoi = '" & NgungTheoDoi & "'"
        
        Call ViewListBox(Query, .cbbKH, dbConn)
        
        'Set default khach hang dau tien
        With .cbbKH
            .Text = .List(0, 0)
        End With
        
        ' Hien thi nam bao cao
        
        Query = "Select Distinct Year(Convert(date, NgayHachToan)) Nam from KD_DonHang where NgayHachToan is Not null order by Year(Convert(date, NgayHachToan))"
    
        Call ViewListBox(Query, .cbbNam, dbConn)
        
        'Set default nam bao cao dau tien
        
        With .cbbNam
            .Text = .List(.ListCount - 1, 0)
        End With
        
        'Dong ket noi csdl
        Call CloseDatabaseConnection(dbConn)
    End With


End Sub

Sub chon_tab_ngay_Data_SSKHKH()
    Set wSheet = Sheet37
    
    With wSheet
    .Select
    .Range("A1").Select
    End With
    
    Set wSheet = Nothing
End Sub

Sub chon_tab_tuan_Data_SSKHKH()
    Set wSheet = Sheet37
     
    With wSheet
    .Select
    .Range("A1").Select
     ActiveWindow.SmallScroll ToRight:=22
    End With
    
   Set wSheet = Nothing
End Sub

Sub chon_tab_thang_Data_SSKHKH()
   Set wSheet = Sheet37
     
    With wSheet
    .Select
    .Range("A1").Select
    ActiveWindow.SmallScroll ToRight:=41
    End With
    
   Set wSheet = Nothing
   
End Sub

Sub chon_tab_nam_Data_SSKHKH()
    Set wSheet = Sheet37
     
    With wSheet
    .Select
    .Range("A1").Select
     ActiveWindow.SmallScroll ToRight:=59
    End With
End Sub



