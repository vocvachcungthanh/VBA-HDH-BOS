Attribute VB_Name = "M_SoChiTietBanHang"
Sub TaiDuLieuSoChiTietBanHang()
    BatLimit
         Call HienSoBanHang_ChiTiet
'         Call Refresh_all_pivot_TB
         Sheet24.Select
       ActiveWorkbook.RefreshAll
    TatLimit
      ThongBao_ThanhCong
End Sub

Sub HienSoBanHang_ChiTiet()
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset
    Dim SQLStr As String
    Dim dongCuoi
    On Error Resume Next
    With Sheet24
        .Select
        StrCnn = KetNoiMayChu_KhachHang
        Set Cn = New ADODB.Connection
        Cn.Open StrCnn
        
        '------Ngay data moi nhat
        .Range("G1").Clear
        
        Set Rs = New ADODB.Recordset
        SQLStr = "select top 1  Convert(date, NgayHachToan) from KD_DonHang order by  Convert(date, NgayHachToan) desc"
        Rs.Open SQLStr, Cn, adOpenStatic
        .Range("G1").CopyFromRecordset Rs
        
        ' Don hang chi tiet
        Set Rs = New ADODB.Recordset
        
        dongCuoi = Abs(tinhdongcuoi("E4:E1048576"))
        
        If dongCuoi > 3 Then
            .Range("A4:BA" & dongCuoi).Clear
        End If
        
        SQLStr = "exec KD_DonHang_ChiTiet"
       
         Rs.Open SQLStr, Cn, adOpenStatic
        Dim k As Integer
        For Each Field In Rs.Fields
            .Range("a3").Offset(0, k).Value = Field.Name
            k = k + 1
        Next Field
    
        .Range("A4").CopyFromRecordset Rs
        Cn.Close
        
        Set Rs = Nothing
        Set Cn = Nothing
        dongCuoi = Abs(tinhdongcuoi("E4:E1048576"))
        
        If dongCuoi <= 3 Then
            .ListObjects("DataSCTBH").Resize Range("A3:AE4")
        Else
            .ListObjects("DataSCTBH").Resize Range("A3:AE" & dongCuoi)
        End If
        
       
            Columns("V:V").Select
            Selection.NumberFormat = "@"
            
        dongCuoi = tinhdongcuoi("A3:A1000000")
        .Range("DataSCTBH[[#Headers],[Mã khách hàng]]").Select
        .ListObjects("DataSCTBH").Resize Range("A3:AE" & dongCuoi)
    End With
    
End Sub
