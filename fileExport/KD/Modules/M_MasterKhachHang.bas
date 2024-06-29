Attribute VB_Name = "M_MasterKhachHang"
Sub Master_KhachHang()
        On Error Resume Next
        Dim dongCuoi As Long
        Dim Run As Variant
        Dim wSheet As Worksheet
        Dim DongCuoiCuaHang As Long
        Set wSheet = Sheet13
        
        With wSheet
         .Select
           dongCuoi = tinhdongcuoi("B12:B1048576")
           
            If .Range("B12") <> "" Then
               .Range("B12:T" & dongCuoi).Clear
            End If
            
            '        Mo ket noi csdl
            Dim dbConn As Object
            Set dbConn = ConnectToDatabase
                
            Call viewSheet("SELECT * FROM KH_KhachHang ORDER BY KhachHangID", Sheet13, "B12", dbConn)
            
            'Dong Ket noi
            Call CloseDatabaseConnection(dbConn)
        
            'MsgBox SQLStr
            
            dongCuoi = tinhdongcuoi("B12:B1048576")
            
            If dongCuoi <= 11 Then
                ActiveSheet.ListObjects("TableMasterDataKH").Resize Range("$B$11:$T$12")
                F_BoderStyle .Range("$B$11:$T12$"), "TableMasterDataKH"
            Else
                ActiveSheet.ListObjects("TableMasterDataKH").Resize Range("$B$11:$T$" & dongCuoi)
                F_BoderStyle .Range("$B$11:$T$" & dongCuoi), "TableMasterDataKH"
            End If
            
            Set wSheet = Nothing
            .Range("A1").Select
        End With
       
End Sub

Sub DinhDangMasterDataKhachHang()

        Dim dongCuoi
        
        dongCuoi = tinhdongcuoi("B12:B1048576")
        
        Format_ dongCuoi, 20, "B11", "A:A", 0, "TableMasterDataKH"
        F_Width "B:B", 0
         FixTop Range("A12")
End Sub

Sub LamMoiMasterKhachHang()
    BatLimit
        Master_KhachHang
      DinhDangMasterDataKhachHang
    TatLimit
      ThongBao_ThanhCong
End Sub

Sub CapNhatMasterKhachHang()
    BatLimit
        Dim MaKhachHang As String
        Dim NguonHoSoID As LongLong
        Dim TenKhachHang As String
        Dim DiaChi As String
        Dim DienThoai As String
        Dim Email As String
        Dim Website As String
        Dim DaiDienPhapLy As String
        Dim MaSoThue As String
        Dim TrangThai As Integer
        Dim ChiNhanh As String
        Dim LaCNTTNuocNgoai As String
        Dim NgungTheoDoi As String
        Dim NhomKHNCC As String
        Dim TinhTP As String
        Dim QuanHuyen As String
        Dim PhuongXa As String
        Dim MaNhanVien As String
        Dim dongCuoi
        Dim i

        Dim Cn As ADODB.Connection
        Dim StrCnn As String
        Dim SQLStr As String
        Set Rs = New ADODB.Recordset
        
        StrCnn = KetNoiMayChu_KhachHang
        Set Cn = New ADODB.Connection
        Cn.Open StrCnn
        
        ' Xoa Khach hang
        SQLStr = "DELETE FROM KH_KhachHang"
        With Sheet13
            .Select

            dongCuoi = tinhdongcuoi("B12:B1048576")
            
            Rs.Open SQLStr, Cn, adOpenStatic
             
            Set Rs = New ADODB.Recordset

            For i = 12 To dongCuoi
                MaKhachHang = .Range("C" & i)
                NguonHoSoID = .Range("D" & i)
                TenKhachHang = .Range("E" & i)
                DiaChi = .Range("F" & i)
                DienThoai = .Range("G" & i)
                Email = .Range("H" & i)
                Website = .Range("I" & i)
                DaiDienPhapLy = .Range("J" & i)
                MaSoThue = .Range("K" & i)
                TrangThai = .Range("L" & i)
                ChiNhanh = .Range("M" & i)
                LaCNTTNuocNgoai = .Range("N" & i)
                NgungTheoDoi = .Range("O" & i)
                NhomKHNCC = .Range("P" & i)
                TinhTP = .Range("Q" & i)
                QuanHuyen = .Range("R" & i)
                PhuongXa = .Range("S" & i)
                MaNhanVien = .Range("T" & i)
            
                If MaKhachHang <> "" And _
                   NgungTheoDoi <> "" And _
                   MaNhanVien <> "" _
                Then
                    SQLStr = "INSERT INTO KH_KhachHang(MaKhachHang, NguonHoSoID,TenKhachHang, DiaChi,DienThoai,email,Website,DaiDienPhapLy,MaSoThue, TrangThai, ChiNhanh, LaCNCTNuocNgoai, NgungTheoDoi, NhomKHNCC,TinhTP,QuanHuyen,PhuongXa,MaNhanVien) VALUES(N'" & MaKhachHang & "',N'" & NguonHoSoID & "', N'" & TenKhachHang & "', N'" & DiaChi & "', N'" & DienThoai & "', N'" & Email & "', N'" & Website & "', '" & DaiDienPhapLy & "', N'" & MaSoThue & "'," & TrangThai & ", N'" & ChiNhanh & "', N'" & LaCNCTNuocNgoai & "', N'" & NgungTheoDoi & "', N'" & NhomKHNCC & "', N'" & TinhTP & "',N'" & QuanHuyen & "', N'" & PhuongXa & "', N'" & MaNhanVien & "')"
                    
                    Rs.Open SQLStr, Cn, adOpenStatic
                    
                    
                   
                Else
                    MsgBox "Cac truong Ma khach hang, Ngung theo doi, Ma nhan vien khong duoc de trong Dong " & i
                    .Range("C" & i).Select
                    Exit Sub
                End If
            Next i
             
        End With
        Cn.Close
        Set Rs = Nothing
        Set Cn = Nothing
   
    TatLimit
    ThongBao_ThanhCong
End Sub




