VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private isActive As Boolean

Private Sub cbbDVKD_Change()
    HienThiDuLieu_BaoCaoKeHoachLuyKe_NhanVien
End Sub

Private Sub cbbNam_Change()
    With cbbNam
        Sheet34.Range("G5") = .Value
        
    End With
End Sub

Private Sub cbbNV_Change()
    With cbbNV
        Sheet34.Range("E5") = .Text
        Sheet34.Range("B5") = .Value
    End With
End Sub

Private Sub TextBox2_Change()
    If TextBox2.Value >= 1 And TextBox2.Value <= 41 Then
    ThisWorkbook.Sheets("Data SS KH NVKD").Range("X102") = TextBox2.Value - 1
    Else
    MsgBox "Nhap so tu 1 den 41", vbOKOnly
    End If
End Sub

Private Sub Worksheet_Activate()
   Call ScrollToTop
   Call hideall
   
   If isActive = False Then
        Call VeBieuDo_SoSanhCua_NVKD
   End If
   
   isActive = True
End Sub

Sub HienThiDuLieu_BaoCaoKeHoachLuyKe_NhanVien()
    BatLimit
    Dim NgayBaoCao As Date
    Sheets("Data SS KH NVKD").Select
    Sheets("Data SS KH NVKD").Range("BJ29").Formula = "=Data!$G$1"
    Dim TenPB As String
    On Error Resume Next

    TenPB = cbbDVKD.Text
    On Error Resume Next
    NgayBaoCao = Sheets("Data SS KH NVKD").Range("BJ29")
    
    Sheets("Data SS KH NVKD").Range("BJ31:BQ900").Clear
    
    If TenPB = "" Or NgayBaoCao = "1900-01-01" Then
        Exit Sub
    End If
    
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset
    Set Cn = New ADODB.Connection
    StrCnn = KetNoiMayChu_KhachHang
    Dim SQLStr As String
    Dim k As Integer

    
    
'' Hien thi ke hoach phong ban
    SQLStr = "exec KD_BaoCaoKinhDoanh_LuyKeKeHoach 'NV', N'" & TenPB & "', '" & Format(NgayBaoCao, "yyyy-MM-dd") & "' "
    'MsgBox SQLStr
    Set Rs = New ADODB.Recordset
    Cn.Open StrCnn
    Rs.Open SQLStr, Cn, adOpenStatic

    k = 0
    For Each Field In Rs.Fields
        Sheets("Data SS KH NVKD").Range("BJ30").Offset(0, k).Value = Field.Name
        k = k + 1
    Next Field
 
    Sheets("Data SS KH NVKD").Range("BJ31").CopyFromRecordset Rs

    Cn.Close
    Set Cn = Nothing
    
    ' Dinh dang
    'Range("BK31:BM900").Select
    Sheets("Data SS KH NVKD").Range("BK31:BM900").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Sheets("Data SS KH NVKD").Range("BN31:BP900").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    Sheets("DB SS KH cua NVKD").Select
    
    Call VeBieuDoLuyKeDoanhThuNVKD
    TatLimit
End Sub

