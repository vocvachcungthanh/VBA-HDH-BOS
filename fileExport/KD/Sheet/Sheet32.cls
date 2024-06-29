VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private isActive As Boolean

Private Sub cbbDVKD_Change()
    With cbbDVKD
        Sheet32.Range("E5") = .Value
        
    End With

    HienThiDuLieu_BaoCaoKeHoachLuyKe_PhongBan
    Call VeBieuDoLuyKeDoanhThuDV
End Sub

Private Sub cbbNam_Change()
    With cbbNam
        Sheet32.Range("F5") = .Value
    End With
    
    Call VeBieuDoLuyKeDoanhThuDV
End Sub

Private Sub TextBox1_Change()
    If TextBox1.Value >= 1 And TextBox1.Value <= 41 Then
        ThisWorkbook.Sheets("Data SS KH DVKD").Range("X102") = TextBox1.Value - 1
    Else
        MsgBox "Nhap so tu 1 den 41", vbOKOnly
    End If
End Sub

Private Sub Worksheet_Activate()
    Call ComboBox_DonVi_Nam
    Application.ScreenUpdating = False
    Call ScrollToTop
    Call hideall
    
    If isActive = False Then
        Call VeBieuDo_SoSanhKhachHang_DVKD
    End If
    
    isActive = True
End Sub

Sub HienThiDuLieu_BaoCaoKeHoachLuyKe_PhongBan()
   
    Dim NgayBaoCao As Date
    BatLimit
    Sheets("Data SS KH DVKD").Select
    Sheets("Data SS KH DVKD").Range("BJ29").Formula = "=Data!$G$1"
    Dim TenPB As String
    On Error Resume Next

    TenPB = cbbDVKD.Text
    On Error Resume Next
    NgayBaoCao = Sheets("Data SS KH DVKD").Range("BJ29")
    
    Sheets("Data SS KH DVKD").Range("BJ31:BQ900").Clear
    
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
    SQLStr = "exec KD_BaoCaoKinhDoanh_LuyKeKeHoach 'PB', N'" & TenPB & "', '" & Format(NgayBaoCao, "yyyy-MM-dd") & "' "
    'MsgBox SQLStr
    Set Rs = New ADODB.Recordset
    Cn.Open StrCnn
    Rs.Open SQLStr, Cn, adOpenStatic

    k = 0
    For Each Field In Rs.Fields
        Sheets("Data SS KH DVKD").Range("BJ30").Offset(0, k).Value = Field.Name
        k = k + 1
    Next Field
 
    Sheets("Data SS KH DVKD").Range("BJ31").CopyFromRecordset Rs

    Cn.Close
    Set Cn = Nothing
    
    ' Dinh dang
    'Range("BK31:BM900").Select
    Sheets("Data SS KH DVKD").Range("BK31:BM900").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    Sheets("Data SS KH DVKD").Range("BN31:BP900").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    Sheets("DB SS KH cua DVKD").Select
    TatLimit
End Sub


