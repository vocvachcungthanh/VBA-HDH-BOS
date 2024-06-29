Attribute VB_Name = "M_MasterSanPham"
Private MangA As New Collection
Private MangB As New Collection

Sub Master_DanhSachSP()
 
        On Error Resume Next
        Dim dongCuoi As Long
        Dim Run As Variant
        Dim wSheet As Worksheet
        Dim DongCuoiCuaHang As Long
        Dim i
        
        Set wSheet = Sheet14
        
        With wSheet
            .Select
            dongCuoi = tinhdongcuoi("B12:B1048576")
            If .Range("B12") <> "" Then
               .Range("B12:V" & dongCuoi).Clear
            End If
            
            Dim Query As String
             '        Mo ket noi csdl
            Dim dbConn As Object
            Set dbConn = ConnectToDatabase
            Query = "SELECT MaSanPham, TenSanPham, NhomVTHH1, NhomVTHH2, NhomVTHH3, NhomVTHH4, NhomVTHH5, NhomVTHH6, NgungTheoDoi, ISNULL(GiaNiemYet, 0) As GiaNiemYet, ISNULL(TiLeChietKhau,0) As TiLeChietKhau, ISNULL(GiaBanBinhQuan,0) as GiaBanBinhQuan, SanPhamID from SP_SanPham ORDER BY SanPhamID"
            Call viewSheet(Query, Sheet14, "B12", dbConn)
            
             'Dong Ket noi
            Call CloseDatabaseConnection(dbConn)
           
            dongCuoi = tinhdongcuoi("B12:B1048576")
            
            For i = 12 To dongCuoi
               MangA.Add .Range("N" & i).Value
            Next i
            
            If dongCuoi <= 11 Then
                ActiveSheet.ListObjects("TableMasterDataSanPham").Resize Range("$B$11:$N$12")
            Else
                 ActiveSheet.ListObjects("TableMasterDataSanPham").Resize Range("$B$11:$N$" & dongCuoi)
                 F_BoderStyle .Range("$B$11:$N$" & dongCuoi), "TableMasterDataSanPham"
            End If
              
              .Range("A1").Select
              
              'Tinh gia ban binh quan
              'Cong thuc gia ban binh quan = Gia niem yet * 1 - Ti le chiet khau / 100
              
              For i = 12 To dongCuoi
                .Range("M" & i) = "=K" & i & " * (1 - L" & i & " / 100)"
              Next i
         
        End With
        
        Set wSheet = Nothing
End Sub

Sub DinhDangMasterDataSanPham()
  
        Dim dongCuoi
        dongCuoi = tinhdongcuoi("B12:B1048576")
        
        If dongCuoi > 12 Then
            Format_ dongCuoi, 14, "B11", "A:A", 0, "TableMasterDataSanPham"
        End If
        
        F_Width "N:N", 0
        FixTop Range("A12")
End Sub

Sub LamMoiMasterSanPham()
    BatLimit
        Master_DanhSachSP
        DinhDangMasterDataSanPham
    TatLimit
       ThongBao_ThanhCong
End Sub

Sub CapNhatSanPham_mt()
    BatLimit
        Dim MaSanPham As String
        Dim TenSanPham As String
        Dim NhomVTHH1 As String
        Dim NhomVTHH2 As String
        Dim NhomVTHH3 As String
        Dim NhomVTHH4 As String
        Dim NhomVTHH5 As String
        Dim NhomVTHH6 As String
        Dim dongCuoi As Long
        Dim NgungTheoDoi As String
        Dim GiaNiemYet As Double
        Dim TiLeChietKhau As Double
        Dim GiaBanBinhQuan As Double
        Dim SanPhamID
        Dim dbConn As Object
        Dim Rs As Object
        Dim Query As String
        
        With Sheet14
             dongCuoi = tinhdongcuoi("B12:B1048576")

             Set dbConn = ConnectToDatabase
                
            For i = 12 To dongCuoi
                MaSanPham = .Range("B" & i)
                TenSanPham = .Range("C" & i)
                NhomVTHH1 = .Range("D" & i)
                NhomVTHH2 = .Range("E" & i)
                NhomVTHH3 = .Range("F" & i)
                NhomVTHH4 = .Range("G" & i)
                NhomVTHH5 = .Range("H" & i)
                NhomVTHH6 = .Range("I" & i)
                NgungTheoDoi = .Range("J" & i)
                GiaNiemYet = .Range("K" & i)
                TiLeChietKhau = .Range("L" & i)
                GiaBanBinhQuan = .Range("M" & i)
                SanPhamID = .Range("N" & i)
                
                MaSanPham = Replace(MaSanPham, "'", "''")
                TenSanPham = Replace(TenSanPham, "'", "''")
                NhomVTHH1 = Replace(NhomVTHH1, "'", "''")
                NhomVTHH2 = Replace(NhomVTHH2, "'", "''")
                NhomVTHH3 = Replace(NhomVTHH3, "'", "''")
                NhomVTHH4 = Replace(NhomVTHH4, "'", "''")
                NhomVTHH5 = Replace(NhomVTHH5, "'", "''")
                NhomVTHH6 = Replace(NhomVTHH6, "'", "''")
                
                MangB.Add .Range("N" & i)
            
                If MaSanPham <> "" And SanPhamID > 0 Then
                    Query = "UPDATE SP_SanPham Set MaSanPham = N'" & MaSanPham & "', TenSanPham = N'" & TenSanPham & "', NhomVTHH1 = N'" & NhomVTHH1 & "', NhomVTHH2 = N'" & NhomVTHH2 & "', NhomVTHH3 = N'" & NhomVTHH3 & "', NhomVTHH4 = N'" & NhomVTHH4 & "', NhomVTHH5 = N'" & NhomVTHH5 & "', NhomVTHH6 = N'" & NhomVTHH6 & "', NgungTheoDoi = '" & NgungTheoDoi & "', GiaNiemYet = " & GiaNiemYet & ", TiLeChietKhau = " & TiLeChietKhau & ", GiaBanBinhQuan = " & GiaBanBinhQuan & " WHERE SanPhamID = " & SanPhamID
                    Set Rs = dbConn.Execute(Query)
                    Set Rs = Nothing
                ElseIf MaSanPham <> "" And (SanPhamID = 0 Or SanPhamID = "") Then
                    Query = "INSERT INTO SP_SanPham(MaSanPham,TenSanPham,NhomVTHH1,NhomVTHH2,NhomVTHH3,NhomVTHH4,NhomVTHH5,NhomVTHH6,NgungTheoDoi,GiaNiemYet,TiLeChietKhau,GiaBanBinhQuan) VALUES(N'" & MaSanPham & "', N'" & TenSanPham & "',N'" & NhomVTHH1 & "', N'" & NhomVTHH2 & "',N'" & NhomVTHH3 & "', N'" & NhomVTHH4 & "', N'" & NhomVTHH5 & "', N'" & NhomVTHH6 & "', '" & NgungTheoDoi & "', " & GiaNiemYet & ", " & TiLeChietKhau & ", " & GiaBanBinhQuan & ")"
               
                    Set Rs = dbConn.Execute(Query)
                    Set Rs = Nothing
                Else
                    Dim TieuDe As String
                    Dim noidung As String
                          
                    TieuDe = "BOS xin thông báo"
                    noidung = "Tên s" & ChrW(7843) & "n ph" & ChrW(7849) & _
                        "m ho" & ChrW(7863) & "c mã s" & ChrW(7843) & "n ph" & ChrW(7849) & "m " _
                        & ChrW(273) & "ang b" & ChrW(7883) & " tr" & ChrW(7889) & "ng t" & ChrW( _
                        7841) & "i dòng" & i & ""
                            Application.Assistant.DoAlert TieuDe, noidung, msoAlertButtonOK, msoAlertIconWarning, 0, 0, 0
                            ActiveWindow.ScrollRow = i
                    Exit Sub
                End If
                
            Next i
            
            Call XoaSanPham(dbConn)
            Set MangA = New Collection
            Set MangB = New Collection
 
            Call CloseDatabaseConnection(dbConn)
           End With
    TatLimit
         ThongBao_ThanhCong
End Sub

Sub XoaSanPham(dbConn As Object)
    On Error Resume Next
    Dim itemA As Variant
    Dim itemB As Variant
    Dim found As Boolean
    Dim Query As String
    Dim Rs As Object
    
    For Each itemA In MangA
        found = False
        For Each itemB In MangB
            If itemA = itemB Then
                found = True
                Exit For
            End If
        Next itemB
        
        If Not found Then
            Query = "DELETE SP_SanPham WHERE SanPhamID = " & itemA
            Set Rs = dbConn.Execute(Query)
            
            Set Rs = Nothing
            
        End If
    Next itemA
End Sub



