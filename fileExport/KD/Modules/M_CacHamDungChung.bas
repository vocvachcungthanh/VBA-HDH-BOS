Sub BatLimit()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.PrintCommunication = False
End Sub

Sub TatLimit()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.CutCopyMode = False
End Sub

Sub FormatMoney()
    Selection.NumberFormat = "#,##0;-#,##0"
End Sub

Sub FormatPercent()
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
End Sub

Sub RefreshPivotTables()
    Range("E11").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
End Sub

Sub FixTop(r As Range)
    r.Select
    ActiveWindow.FreezePanes = True
End Sub

Sub AutoWidth(r As Range)
    r.Select
    Selection.Columns.AutoFit
End Sub

Function ThongBaoLamMoi()
    CreateObject("WScript.Shell").Popup "L" & ChrW(224) & "m m" & ChrW(7899) & "i d" & ChrW(7919) & " li" & ChrW(7879) & "u k" & ChrW(7871) & "t th" & ChrW(250) & "c", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 0
End Function

Function ThongBaoCapNhat()
    CreateObject("WScript.Shell").Popup "C" & ChrW(7853) & "p nh" & ChrW(7853) & "t d" & ChrW(7919) & " li" & ChrW(7879) & "u th" & ChrW(224) & "nh c" & ChrW(244) & "ng", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o", 0 + 0
End Function

Function ThongBaoEmpty()
    CreateObject("WScript.Shell").Popup "Vui l" & ChrW(242) & "ng nh" & ChrW(7853) & "p " & ChrW(273) & ChrW(7845) & "y " & ChrW(273) & ChrW(7911) & " th" & ChrW(244) & "ng tin", , "Th" & ChrW(244) & "ng b" & ChrW(225) & "o l" & ChrW(7895) & "i", 0 + 48
End Function

Function AutoZoom(z As Long)
    On Error Resume Next
    ActiveWindow.Zoom = z
    ActiveWindow.ScrollColumn = 1
End Function

Sub ruller()
    If ActiveWindow.DisplayHeadings = True Then
        ActiveWindow.DisplayHeadings = False
    Else
        ActiveWindow.DisplayHeadings = True
    End If

End Sub

Sub hideall()
    '    Workbooks("KD.xlsb").Activate
    '    Application.ExecuteExcel4Macro "Show.toolbar(""Ribbon"",False)"
    '    ActiveWindow.DisplayHorizontalScrollBar = True
    '    ActiveWindow.DisplayVerticalScrollBar = True
    '    ActiveWindow.DisplayHeadings = False
    '    ActiveWindow.DisplayWorkbookTabs = False
    '    Application.DisplayFormulaBar = False
    '    ActiveWindow.DisplayGridlines = False
    '    ActiveWindow.DisplayOutline = False
    '    ActiveWindow.DisplayZeros = True
    '    Application.ScreenUpdating = False
End Sub

Sub showall()
    Dim QC As Integer

    QC = 0

    If QC = 0 Then
        Application.ExecuteExcel4Macro "Show.toolbar(""Ribbon"",True)"
        ActiveWindow.DisplayHorizontalScrollBar = True
        ActiveWindow.DisplayVerticalScrollBar = True
        ActiveWindow.DisplayHeadings = True
        ActiveWindow.DisplayWorkbookTabs = True
        Application.DisplayFormulaBar = True
        ActiveWindow.DisplayGridlines = True
        ActiveWindow.DisplayOutline = True
        ActiveWindow.DisplayZeros = True
        Application.ScreenUpdating = True
    End If
End Sub

Function timdong(rg1 As Range, dk1 As Variant, rg2 As Range, dk2 As Variant) As Long
    Dim i As Long, j As Long, k As Long
    k = 0
    For i = 1 To rg1.Count
        If rg1(i) = dk1 Then
            If rg2(i) = dk2 Then
                k = i
            End If

        End If

    Next i
    timdong = k
End Function

Sub DinhDangBdNhiet(chart_name As String)
    ActiveSheet.ChartObjects(chart_name).Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    ActiveChart.FullSeriesCollection(1).DataLabels.Select
    Selection.ShowValue = True
    Selection.ShowCategoryName = True
    Selection.Separator = "" & Chr(13) & ""
    Selection.ShowSeriesName = False
End Sub

Sub select_data(chart_name As String, table_name As String, Pivot As String)
    ActiveSheet.ChartObjects(chart_name).Activate
    ActiveChart.FullSeriesCollection(1).Select
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=Sheets(Pivot).Range(table_name)
End Sub

Sub show_them_master_data()
    Them_Master_data.Show
End Sub

Sub show_xoa_master_data()
    Xoa_Master_data.Show
End Sub

Sub ShowTreeNode_Menu_KD()
    If TreeNote_Menu_KD.Visible Then
        Unload TreeNote_Menu_KD
    Else
        TreeNote_Menu_KD.Show
    End If
End Sub

Sub Home()
    Dim kd As Workbook
    On Error Resume Next
    Set kd = Workbooks.Open("Core.xlsb")
    Set kd = Workbooks("Core.xlsb")
    kd.Activate
    kd.Sheets("Home").Select

    Set kd = Nothing
End Sub

Private Sub ThucHienTruyVan()
    Dim Cn As ADODB.Connection
    Dim StrCnn As String
    Dim Rs As ADODB.Recordset
    Dim Field As Field
    Set Rs = New ADODB.Recordset

    StrCnn = KetNoiMayChu_KhachHang
    Dim SQLStr As String
    SQLStr = Sheets("Danh sách ??n v?").Range("A1").Value

    Set Cn = New ADODB.Connection
    Cn.Open StrCnn
    Rs.Open SQLStr, Cn, adOpenStatic

    Dim k As Integer
    For Each Field In Rs.Fields
        Worksheets("Danh sách ??n v?").Range("a2").Offset(0, k).Value = Field.Name
        k = k + 1

    Next Field

    Worksheets("Danh sách ??n v?").Range("a3").CopyFromRecordset Rs

    Cn.Close
    Set Rs = Nothing
    Set Cn = Nothing
End Sub

Function NguoiDungID() As Long
    Dim r As Range
    Set r = Workbooks("Core.xlsb").Sheets("PhanQuyen").Range("H1")
    If r.Value <> "" Then
        NguoiDungID = r.Value
    Else
        NguoiDungID = 0
    End If

    Set r = Nothing
End Function

Public Sub F_TextCenter(col As String, r As String)

    If col <> "" Then
        Columns("" & col & "").Select
    Else
        Range(" & r & ").Select
    End If

    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Public Sub F_TextLeft(col As String, r As String)

    If col <> "" Then
        Columns("" & col & "").Select
    Else
        Range("" & r & "").Select
    End If

    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Public Sub F_Width(col As String, Size As Long)
    Columns("" & col & "").Select
    Selection.ColumnWidth = Size
End Sub

Public Sub F_TextWrap(col As String)
    Columns("" & col & "").Select
    With Selection
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

Public Sub F_ClearBaoCaoKinhDoanh(wSheet As Worksheet, sheetName As String)
    With wSheet
        If .Range("B12") <> 0 Then
            Workbooks("KD.xlsb").Sheets("" & sheetName & "").Range("B12:L377").Clear
        End If

        If .Range("X12") <> 0 Then
            Workbooks("KD.xlsb").Sheets("" & sheetName & "").Range("X12:AC65").ClearContents

        End If

        If .Range("AQ12") <> 0 Then
            Workbooks("KD.xlsb").Sheets("" & sheetName & "").Range("AQ12:AW23").ClearContents
        End If

        If .Range("BI12") <> 0 Then
            Workbooks("KD.xlsb").Sheets("" & sheetName & "").Range("BI12:BN15").ClearContents
            Workbooks("KD.xlsb").Sheets("" & sheetName & "").Range("BI20:BN21").ClearContents
            Workbooks("KD.xlsb").Sheets("" & sheetName & "").Range("BI26:BN26").ClearContents
        End If
    End With
End Sub


Sub TestPhanQuyenClear()
    BatLimit
    F_ClearAll
    TatLimit
End Sub

Public Sub F_ClearAll()
    With SheetDataDonViKD
        .Select
        Range("A1").Value = ""
        .Range("J5:L5") = 0

        If .Range("C12") <> "" Then
            dongCuoi = tinhdongcuoi("C12:C1048576") + 100
            .Range("A12:AO" & dongCuoi).ClearContents
        End If

    End With

    With Sheet11
        .Select
        If .Range("G340") <> "" Then
            .Range("G340:I437").ClearContents
        End If

    End With

    With SheetDataNhanVienKD
        .Select
        Range("A1").Value = ""
        .Range("J5:L5") = 0

        If .Range("A12") <> "" Then
            dongCuoi = tinhdongcuoi("C12:C1048576")
            .Range("A12:AO" & dongCuoi).ClearContents
        End If
    End With

    With SheetDataKhachHangKD
        .Select
        .Range("A1").Value = ""
        .Range("I5:K5") = 0

        If .Range("C12") <> "" Then
            dongCuoi = tinhdongcuoi("C12:C1048576")
            .Range("A12:AN" & dongCuoi).ClearContents
        End If
    End With

    With Sheet1611
        .Select
        .Range("H5:J5") = 0
        .cbbNam.Clear
        .lbChonThangLapKH.Clear
        .cbbNamLichSu.Clear
        .lbThangLichSu.Clear

        If .Range("B12") <> "" Then
            dongCuoi = tinhdongcuoi("B12:B1048576")
            .Range("A12:AC" & dongCuoi).ClearContents
        End If
    End With

    With Sheet14
        .Select
        If .Range("B12") <> 0 Then
            dongCuoi = tinhdongcuoi("B12:B1048576")
            .Range("B12:N" & dongCuoi).ClearContents
        End If
    End With

    With Sheet13
        .Select
        If .Range("B12") <> 0 Then
            dongCuoi = tinhdongcuoi("B12:B1048576")
            .Range("B12:T" & dongCuoi).ClearContents
        End If
    End With

    With SheetTheoDoiBienLoiNhuan
        .Select
        .cbNam.Clear
        .cbNam.Value = ""
        If .Range("B12") <> 0 Then
            dongCuoi = tinhdongcuoi("B12:B1048576")
            .Range("B12:I" & dongCuoi).ClearContents
        End If
    End With

    F_ClearBaoCaoKinhDoanh Sheet33, "Data SS KH DVKD"

    F_ClearBaoCaoKinhDoanh Sheet35, "Data SS KH NVKD"
    F_ClearBaoCaoKinhDoanh Sheet37, "Data SS KH KH"
    With Workbooks("KD.xlsb").Sheets("Data SS KH KH")
        .Select
        .cbbNam.Text = ""
        .cbbNam.Clear
        .cbbKH.Text = ""
        .cbbKH.Clear
        .Range("E5") = ""
    End With
    With Workbooks("KD.xlsb").Sheets("DB SS KH cua NVKD")
        .Select
        .cbbNam.Text = ""
        .cbbNam.Clear
        .cbbNV.Text = ""
        .cbbNV.Clear
        .Range("E5") = ""
    End With

    With Workbooks("KD.xlsb").Sheets("DB SS KH cua DVKD")
        .Select
        .cbbNam.Text = ""
        .cbbNam.Clear

        .cbbDVKD.Text = ""
        .cbbDVKD.Clear
    End With

    Call ClearSCTBH

End Sub

Sub ClearSCTBH()
    On Error Resume Next
    Dim dongCuoi As Long
    With Sheet24
        .Select
        .Range("G1").Clear
        dongCuoi = tinhdongcuoi("A4:B1048576")
        If dongCuoi > 3 Then
            Workbooks("KD.xlsb").Sheets("Data").Range("A4:AE" & dongCuoi).Clear
            ActiveSheet.ListObjects("DataSCTBH").Resize Range("$A$3:$AE$4")

            ActiveWorkbook.RefreshAll
        End If
    End With
End Sub

Sub ClearPivotTable()
    BatLimit
    Dim xPt As PivotTable
    Dim xWs As Worksheet
    Dim xPc As PivotCache
    Application.ScreenUpdating = False
    For Each xWs In ActiveWorkbook.Worksheets
        For Each xPt In xWs.PivotTables
            xPt.PivotCache.MissingItemsLimit = xlMissingItemsNone
        Next xPt
    Next xWs
    For Each xPc In ActiveWorkbook.PivotCaches
        On Error Resume Next
        xPc.Refresh
    Next xPc
    Application.ScreenUpdating = True
    On Error Resume Next
    TatLimit
End Sub

Public Sub F_BoderStyle(r As Range, nameTable)
    r.Select
    ActiveSheet.ListObjects("" & nameTable & "").TableStyle = _
    "BOS_Table_Style_1"
End Sub

Sub AnSheet_TheoPhanQuyen()
    Dim sh As Worksheet
    Dim lr As Integer

    Set sh = Workbooks("Core.xlsb").Sheets("PhanQuyen")
    lr = sh.Range("BT1000").End(xlUp).Row

    Dim i As Integer
    Dim tenChucNang As String, sh_pq As String, Quyen As Integer, Menu As String

    For i = 3 To lr
        tenChucNang = sh.Range("BW" & i).Value
        sh_pq = sh.Range("BT" & i).Value
        Quyen = sh.Range("BX" & i).Value
        ' Menu = sh.Range("I" & i).value
        If Quyen = 0 And Len(sh_pq) > 0 Then
            On Error Resume Next
            ThisWorkbook.Worksheets(sh_pq).Visible = False
        Else
            On Error Resume Next
            ThisWorkbook.Worksheets(sh_pq).Visible = True
        End If

    Next i
    On Error Resume Next
    ThisWorkbook.Worksheets("PhanQuyen").Visible = False

    Set sh = Nothing
End Sub


Sub F_R_DATA()
    If Workbooks("KD.xlsb").Worksheets("Data").Range("A4") = "" Then
        HienSoBanHang_ChiTiet
    End If
End Sub

Function tinhdongcuoi(rg As String) As Long
    Dim Cd As Long, dongdau As String, r As Long, s As Long, dongCuoi As String
    Cd = WorksheetFunction.CountIf(Range(rg), "")
    dongdau = ExtractC12FromString(rg)
    dongCuoi = ExtractdongcuoiFromString(rg)
    r = Range(dongdau).Row
    s = Range(dongCuoi).Row
    tinhdongcuoi = s - Cd
End Function

Function ExtractC12FromString(inputString As String) As String

    Dim outputString As String
    Dim colonPosition As Integer

    ' Tìm v? trí d?u hai ch?m (:) trong chu?i
    colonPosition = InStr(1, inputString, ":")

    ' Trích xu?t ký t? "C12" t? chu?i ban d?u
    If colonPosition > 0 Then
        outputString = Left(inputString, colonPosition - 1)
        ExtractC12FromString = outputString
    End If

End Function
Function ExtractdongcuoiFromString(inputString As String) As String

    Dim outputString As String
    Dim colonPosition As Integer

    ' Tìm v? trí d?u hai ch?m (:) trong chu?i
    colonPosition = InStr(1, inputString, ":")

    ' Trích xu?t ký t? "C12" t? chu?i ban d?u
    If colonPosition > 0 Then
        outputString = Right(inputString, Len(inputString) - colonPosition)
        ExtractdongcuoiFromString = outputString
    End If

End Function

Sub Select_lai_data_chart(chart_name As String, tb_name As String)
    ActiveSheet.ChartObjects(chart_name).Activate
    On Error Resume Next
    ActiveChart.FullSeriesCollection(3).Select
    Application.CutCopyMode = False
    ActiveChart.SetSourceData Source:=Sheets("Pivot SP").Range(tb_name & "[#All]")
End Sub

Sub Refresh_all_pivot_TB()
    ThisWorkbook.Sheets("Pivot SP").Select
    Range("A7").Select
    ActiveWorkbook.RefreshAll
End Sub

Public Sub ScrollToTop()
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Sub laygiatrislicer()
    BatLimit
    Dim slcCache As SlicerCache
    Dim slcItem As SlicerItem
    Dim i As Long, k As Long

    For i = 1 To ActiveWorkbook.SlicerCaches.Count
        Set slcCache = ActiveWorkbook.SlicerCaches(i)
        k = 2

        For Each slcItem In slcCache.SlicerItems
            If slcItem.Selected = True Then
                ThisWorkbook.Sheets("Slicer").Cells(1, i) = slcCache.Name
                ThisWorkbook.Sheets("Slicer").Cells(k, i) = slcItem.Name
                k = k + 1
            End If
        Next slcItem
    Next i
    Set slcCache = Nothing
    TatLimit
End Sub

'---------------------

Sub TatTatCaDulieuTu_CSDL()
    On Error Resume Next
    Dim Nam As String
    Dim Query As String
    Dim NguoiDangNhap
    Dim dongCuoi

    NguoiDangNhap = NguoiDungID

    Nam = Year(Now)

    'MoKetNoi
    Dim dbConn As Object
    Set dbConn = ConnectToDatabase

    ' Xoa du lieu tren sheet neu co
    Set wSheetDV = SheetDataDonViKD
    With wSheetDV
        .Select
        If .Range("C12") <> "" Then
            dongCuoi = tinhdongcuoi("C12:C1048576") + 100
            .Range("A12:BU" & dongCuoi).ClearContents
        End If

        Query = "exec dataKHDT_DV_KD_V2 '" & Nam & "'," & NguoiDangNhap & ""
        Call viewSheet(Query, SheetDataDonViKD, "A12", dbConn)

        Query = "exec KD_TK_TongHopTheo_DV " & Nam & ", " & NguoiDangNhap & ""
        Call viewSheet(Query, SheetDataDonViKD, "J5", dbConn)

    End With

    Set wSheetDV = Nothing

    Set wSheet = Sheet11
    With wSheet
        .Select
        If .Range("G340") <> "" Then

            dongCuoi = tinhdongcuoi("G340:G399")

            If .Range("G340") <> "" And dongCuoi > 339 Then
                .Range("G340:I" & dongCuoi).ClearContents
                .Range("G402:I432").ClearContents
            End If

        End If

        Query = "EXEC KD_KeHoachDoanhThu_NamTruocNamSau " & Nam
        Call viewSheet(Query, Sheet11, "G340", dbConn)
    End With
    Set wSheet = Nothing

    Set wSheetNVKD = SheetDataNhanVienKD
    With wSheetNVKD
        .Select
        If .Range("C12") <> "" Then
            dongCuoi = tinhdongcuoi("C12:C1048576")
            .Range("A12:AO" & dongCuoi).ClearContents
        End If

        Query = "exec dataKHDT_NV_KD_V2 '" & Nam & "'," & NguoiDangNhap & ",0 "
        Call viewSheet(Query, SheetDataNhanVienKD, "A12", dbConn)

        Query = "exec KD_TK_TongHopTheo_NV '" & Nam & "', " & NguoiDangNhap & ",0"
        Call viewSheet(Query, SheetDataNhanVienKD, "J5", dbConn)
    End With
    Set wSheetNVKD = Nothing

    Set wSheet = Sheet1611
    With wSheet
        .Select
        dongCuoi = tinhdongcuoi("B12:B1048576")

        If .Range("B12") <> 0 Then
            .Range("A12:Y" & dongCuoi).ClearContents
        End If
        .Cells.Select
        Selection.RowHeight = 15
        Query = " exec DataKHDT_SP_KD_V2 " & Nam & ",0,'1'," & Nam & ", '1'"
        Call viewSheet(Query, Sheet1611, "B12", dbConn)

        'Tong doanh thu theo san pham
        Query = "exec KD_TK_TongHopTheo_SP " & Nam & ", " & NguoiDangNhap & ", 0,'1'"
        Call viewSheet(Query, Sheet1611, "H5", dbConn)
    End With
    Set wSheet = Nothing

    Set wSheetKH = SheetDataKhachHangKD
    With wSheetKH
        .Select

        If .Range("C12") <> "" Then
            dongCuoi = tinhdongcuoi("C12:C1048576")
            .Range("A12:AN" & dongCuoi).ClearContents
        End If

        Query = "exec dataKHDT_KH_KD_V2 '" & Nam & "'," & NguoiDangNhap & ", 0"
        Call viewSheet(Query, SheetDataKhachHangKD, "A12", dbConn)

        'Tong Hop theo KH
        Query = "exec KD_TK_TongHopTheo_KH " & Nam & ", " & NguoiDangNhap & ", 0"
        Call viewSheet(Query, SheetDataKhachHangKD, "I5", dbConn)
    End With
    Set wSheetKH = Nothing
    'So Chi Tiet ban hang

    Set wSheet = Sheet24
    With wSheet
        .Select
        .Range("G1").Clear

        Query = "Select top 1  Convert(date, NgayHachToan) from KD_DonHang order by  Convert(date, NgayHachToan) desc"

        Call viewSheet(Query, Sheet24, "G1", dbConn)

        dongCuoi = tinhdongcuoi("E4:B1048576")

        If dongCuoi > 3 Then
            .Range("A4:BA" & dongCuoi).Clear
        End If

        Query = "exec KD_DonHang_ChiTiet"
        Call viewSheet(Query, Sheet24, "A4", dbConn)
    End With

    Set wSheet = Nothing
    ComboBox_DonVi_Nam

    Set wSheet = Sheet13
    With wSheet
        .Select
        dongCuoi = tinhdongcuoi("B12:B1048576")

        If .Range("B12") <> "" Then
            .Range("B12:T" & dongCuoi).ClearContents
        End If

        Call viewSheet("Select * FROM KH_KhachHang", Sheet13, "B12", dbConn)

    End With
    Set wSheet = Nothing

    Set wSheet = Sheet26
    With wSheet
        .Select
        If .Range("HJ4") <> "" Then
            .Range("HJ4:HL1000").ClearContents
        End If
        Call viewSheet("exec TiLeChiPhi", Sheet26, "HJ4", dbConn)
    End With
    Set wSheet = Nothing
    Dim TenPhongBan As String
    Dim PhongBanID As Integer


    Call ComboBox_SoSanhKeHoachcua_NVKD
    Set wSheet = Sheet32
    With wSheet
        TenPhongBan = .Range("E5").Value
    End With
    Set wSheet = Nothing

    Set wSheet = Sheet33
    With wSheet
        .Select

        Query = "Select Isnull((Select top 1 PhongBanID from PhongBan where TenPhongBan = N'" & TenPhongBan & "'),9999 )"

        Call viewSheet(Query, Sheet33, "A1", dbConn)

        PhongBanID = .Range("A1").Value

        ' Bao cao ngay >> Tu o B12
        Query = "Select STT, NgayThang, Thu, Tuan, Thang, Quy, KY, KeHoach_Ngay , DoanhSoBan_Ngay , TyLeThucHien, (KeHoach_Ngay - DoanhSoBan_Ngay) As Thieu  " & _
        "from BaoCaoDoanhThu_DVKD_TheoNgay(" & Nam & ", " & PhongBanID & ") BC_Ngay"

        Sheet32.Range("B12:L500").Clear
        Call viewSheet(Query, Sheet33, "B12", dbConn)

        ' Bao cáo tuan >> Tu o X12
        Query = "Select Tuan, (Select top 1 Thang from DM_NgayThang where Year(NgayThang) = " & Nam & " And DM_NgayThang.Tuan = BC_Ngay.Tuan order by NgayThang ) Thang, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan,  " & _
        " Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        "from BaoCaoDoanhThu_DVKD_TheoNgay(" & Nam & ", " & PhongBanID & ") BC_Ngay " & _
        "Group by Tuan "

        Call viewSheet(Query, Sheet33, "X12", dbConn)

        ' Bao cao Thang >> Tu o AQ12

        Query = "Select Thang, Quy, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_DVKD_TheoNgay(" & Nam & ", " & PhongBanID & ") BC_Ngay " & _
        "Group by Thang, Quy "
        Call viewSheet(Query, Sheet33, "AQ12", dbConn)

        ' Bao cao Quy >> Tu o BJ12
        Query = "Select Quy, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_DVKD_TheoNgay(" & Nam & ", " & PhongBanID & ") BC_Ngay " & _
        "Group by Quy"

        Call viewSheet(Query, Sheet33, "BI12", dbConn)

        ' Bao cao Ky >> Tu o BI20

        Query = "Select Ky, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_DVKD_TheoNgay(" & Nam & ", " & PhongBanID & ") BC_Ngay " & _
        "Group by Ky"
        Call viewSheet(Query, Sheet33, "BI20", dbConn)

        ' Bao cao Nam >> Tu o BI26

        Query = "Select '" & Nam & "' As Nam, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_DVKD_TheoNgay(" & Nam & ", " & PhongBanID & ") BC_Ngay "

        Call viewSheet(Query, Sheet33, "BI26", dbConn)
    End With
    Set wSheet = Nothing
    Dim MaNhanVien

    Set wSheet = Sheet34
    With wSheet
        MaNhanVien = .Range("E5").Value
    End With
    Set wSheet = Nothing
    ' Xac dinh NhanVienID
    Set wSheet = Sheet35
    With wSheet
        .Select

        Query = "Select Isnull((Select top 1 NhanvienID from Ns_NhanVien where MaNhanVien = N'" & MaNhanVien & "'),9999 )"
        Call viewSheet(Query, Sheet35, "A1", dbConn)

        NhanVienID = .Range("A1").Value

        ' Bao cao ngay >> Tu o B12

        Query = "Select STT, NgayThang, Thu, Tuan, Thang, Quy, Ky, KeHoach_Ngay , DoanhSoBan_Ngay , TyLeThucHien,(KeHoach_Ngay - DoanhSoBan_Ngay) As Thieu  " & _
        "from BaoCaoDoanhThu_NhanVienKD_TheoNgay(" & Nam & ", " & NhanVienID & ") BC_Ngay"
        Call viewSheet(Query, Sheet35, "B12", dbConn)

        ' Bao cáo tuan >> Tu o X12

        Query = "Select Tuan, (Select top 1 Thang from DM_NgayThang where Year(NgayThang) = " & Nam & " And DM_NgayThang.Tuan = BC_Ngay.Tuan order by NgayThang ) Thang, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan,  " & _
        " Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        "from BaoCaoDoanhThu_NhanVienKD_TheoNgay(" & Nam & ", " & NhanVienID & ") BC_Ngay " & _
        "Group by Tuan "

        Call viewSheet(Query, Sheet35, "X12", dbConn)

        ' Bao cao Thang >> Tu o AQ12

        Query = "Select Thang, Quy, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_NhanVienKD_TheoNgay(" & Nam & ", " & NhanVienID & ") BC_Ngay " & _
        "Group by Thang, Quy "
        Call viewSheet(Query, Sheet35, "AQ12", dbConn)


        Query = "Select Quy, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_NhanVienKD_TheoNgay(" & Nam & ", " & NhanVienID & ") BC_Ngay " & _
        "Group by Quy"
        Call viewSheet(Query, Sheet35, "BI12", dbConn)

        ' Bao cao Ky >> Tu o BI20

        Query = "Select Ky, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        "from BaoCaoDoanhThu_NhanVienKD_TheoNgay(" & Nam & ", " & NhanVienID & ") BC_Ngay " & _
        "Group by Ky"
        Call viewSheet(Query, Sheet35, "BI20", dbConn)


        Query = "Select '" & Nam & "' As Nam, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_NhanVienKD_TheoNgay(" & Nam & ", " & NhanVienID & ") BC_Ngay "

        Call viewSheet(Query, Sheet35, "BI26", dbConn)
    End With
    Set wSheet = Nothing

    Set wSheet = Sheet37
    With wSheet
        .Select
        Dim MaKhachHang As String
        Dim KhachHangID As Long
        Call ComboBox_KhachHang_Nam
        MaKhachHang = .Range("J7").Value

        ' Xac dinh KhachHangID

        Query = "Select Isnull((Select top 1 KhachHangID from KH_KhachHang where MaKhachHang = N'" & MaKhachHang & "'),9999 )"
        Call viewSheet(Query, Sheet37, "A1", dbConn)
        KhachHangID = .Range("A1").Value

        ' MsgBox MaNhanVien & " >>>>> " & KhachHangID
        If KhachHangID = 9999 Then
            ' MsgBox "Khach Hang nay khong ton tai. Vui long kiem tra lai"
        End If

        ' Bao cao ngay >> Tu o B12

        Query = "Select STT, NgayThang, Thu, Tuan, Thang, Quy, Ky, KeHoach_Ngay , DoanhSoBan_Ngay , TyLeThucHien, (KeHoach_Ngay - DoanhSoBan_Ngay) As Thieu  " & _
        "from BaoCaoDoanhThu_KhachHang_TheoNgay(" & Nam & ", " & KhachHangID & ") BC_Ngay"
        Call viewSheet(Query, Sheet37, "B12", dbConn)

        ' Bao cáo tuan >> Tu o X12
        Query = "Select Tuan, (Select top 1 Thang from DM_NgayThang where Year(NgayThang) = " & Nam & " And DM_NgayThang.Tuan = BC_Ngay.Tuan order by NgayThang ) Thang, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan,  " & _
        " Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        "from BaoCaoDoanhThu_KhachHang_TheoNgay(" & Nam & ", " & KhachHangID & ") BC_Ngay " & _
        "Group by Tuan "
        Call viewSheet(Query, Sheet37, "X12", dbConn)

        ' Bao cao Thang >> Tu o AQ12

        Query = "Select Thang, Quy, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_KhachHang_TheoNgay(" & Nam & ", " & KhachHangID & ") BC_Ngay " & _
        "Group by Thang, Quy "
        Call viewSheet(Query, Sheet37, "AQ12", dbConn)

        ' Bao cao Quy >> Tu o BJ12

        Query = "Select Quy, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_KhachHang_TheoNgay(" & Nam & ", " & KhachHangID & ") BC_Ngay " & _
        "Group by Quy"
        Call viewSheet(Query, Sheet37, "BI12", dbConn)

        ' Bao cao Ky >> Tu o BI20
        Query = "Select Ky, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        "from BaoCaoDoanhThu_KhachHang_TheoNgay(" & Nam & ", " & KhachHangID & ") BC_Ngay " & _
        "Group by Ky"
        Call viewSheet(Query, Sheet37, "BI20", dbConn)

        ' Bao cao Nam >> Tu o BI26
        Query = "Select '" & Nam & "' As Nam, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from BaoCaoDoanhThu_KhachHang_TheoNgay(" & Nam & ", " & KhachHangID & ") BC_Ngay "

        Call viewSheet(Query, Sheet37, "BI26", dbConn)
    End With
    Set wSheet = Nothing

    'Dong Ket noi
    Call CloseDatabaseConnection(dbConn)

End Sub

Sub F_TaiTatCaDuLieu()
    On Error Resume Next
    Call TatTatCaDulieuTu_CSDL

    Dim dongCuoi

    Set wSheet = Sheet24
    With wSheet
        .Select
        dongCuoi = tinhdongcuoi("E4:E1048576")
        If dongCuoi <= 3 Then
            .ListObjects("DataSCTBH").Resize Range("A3:AE4")
        Else
            .ListObjects("DataSCTBH").Resize Range("A3:AE" & dongCuoi)
        End If

        Columns("V:V").Select
        Selection.NumberFormat = "@"

        dongCuoi = tinhdongcuoi("A3:A1000000")
        .Range("DataSCTBH[[#Headers],[Mã khách hàng]]").Select
        Rows("2:2").RowHeight = 46.5
    End With
    Set wSheet = Nothing
End Sub

Function F_KhoiTaoNam(InputT As Object)
    Dim Nam As Integer

    For Nam = Year(Date) + 7 To 2020 Step -1
        InputT.AddItem Nam
    Next Nam

    InputT.Value = Year(Date)
End Function

Function ReplaceValue(Value As String) As String
    ReplaceValue = Replace(Value, "'", "''")
End Function

Sub ApDungNhan()
    Call SlicerCaption
    Call ThongBao_ThanhCong
End Sub
Sub SlicerCaption()
    '  Slicer  bao cao doanh thu theo san pham
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_1").Slicers("Nhóm VTHH 31")
        .Caption = Sheet14.Range("D11").Value
    End With

    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_2").Slicers("Nhóm VTHH 32")
        .Caption = Sheet14.Range("E11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_3").Slicers("Nhóm VTHH 33")
        .Caption = Sheet14.Range("F11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_4").Slicers("Nhóm VTHH 34")
        .Caption = Sheet14.Range("G11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_5").Slicers("Nhóm VTHH 35")
        .Caption = Sheet14.Range("H11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_6").Slicers("Nhóm VTHH 36")
        .Caption = Sheet14.Range("I11").Value
    End With

    '  Slicer  bao cao daonh thu theo nhan vien kinh doanh

    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_1").Slicers("Nhóm VTHH 62")
        .Caption = Sheet14.Range("D11").Value
    End With

    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_2").Slicers("Nhóm VTHH 63")
        .Caption = Sheet14.Range("E11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_3").Slicers("Nhóm VTHH 64")
        .Caption = Sheet14.Range("F11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_4").Slicers("Nhóm VTHH 65")
        .Caption = Sheet14.Range("G11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_5").Slicers("Nhóm VTHH 66")
        .Caption = Sheet14.Range("H11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_6").Slicers("Nhóm VTHH 67")
        .Caption = Sheet14.Range("I11").Value
    End With

    ' Slicer Bao cao doanh thu theo khach hang
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_1").Slicers("Nhóm VTHH 25")
        .Caption = Sheet14.Range("D11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_2").Slicers("Nhóm VTHH 26")
        .Caption = Sheet14.Range("E11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_3").Slicers("Nhóm VTHH 27")
        .Caption = Sheet14.Range("F11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_4").Slicers("Nhóm VTHH 28")
        .Caption = Sheet14.Range("G11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_5").Slicers("Nhóm VTHH 29")
        .Caption = Sheet14.Range("H11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_6").Slicers("Nhóm VTHH 30")
        .Caption = Sheet14.Range("I11").Value
    End With

    ' Slicer Bao cao doanh thu theo thoi gian
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_1").Slicers("Nhóm VTHH 13")
        .Caption = Sheet14.Range("D11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_2").Slicers("Nhóm VTHH 14")
        .Caption = Sheet14.Range("E11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_3").Slicers("Nhóm VTHH 15")
        .Caption = Sheet14.Range("F11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_4").Slicers("Nhóm VTHH 16")
        .Caption = Sheet14.Range("G11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_5").Slicers("Nhóm VTHH 17")
        .Caption = Sheet14.Range("H11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_6").Slicers("Nhóm VTHH 18")
        .Caption = Sheet14.Range("I11").Value
    End With

    'Slicer Bao cao san luong ban theo san pham
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_1").Slicers("Nhóm VTHH 37")
        .Caption = Sheet14.Range("D11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_2").Slicers("Nhóm VTHH 38")
        .Caption = Sheet14.Range("E11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_3").Slicers("Nhóm VTHH 39")
        .Caption = Sheet14.Range("F11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_4").Slicers("Nhóm VTHH 40")
        .Caption = Sheet14.Range("G11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_5").Slicers("Nhóm VTHH 41")
        .Caption = Sheet14.Range("H11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_6").Slicers("Nhóm VTHH 42")
        .Caption = Sheet14.Range("I11").Value
    End With

    'Slicer Bao cao bien loi nhuan theo san pham
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_1").Slicers("Nhóm VTHH 43")
        .Caption = Sheet14.Range("D11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_2").Slicers("Nhóm VTHH 44")
        .Caption = Sheet14.Range("E11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_3").Slicers("Nhóm VTHH 45")
        .Caption = Sheet14.Range("F11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_4").Slicers("Nhóm VTHH 46")
        .Caption = Sheet14.Range("G11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_5").Slicers("Nhóm VTHH 47")
        .Caption = Sheet14.Range("H11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_6").Slicers("Nhóm VTHH 48")
        .Caption = Sheet14.Range("I11").Value
    End With

    'Slicer Bao doanh thu theo don voi thuc hien thuc hien
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_1").Slicers("Nhóm VTHH 56")
        .Caption = Sheet14.Range("D11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_2").Slicers("Nhóm VTHH 57")
        .Caption = Sheet14.Range("E11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_3").Slicers("Nhóm VTHH 58")
        .Caption = Sheet14.Range("F11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_4").Slicers("Nhóm VTHH 59")
        .Caption = Sheet14.Range("G11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_5").Slicers("Nhóm VTHH 60")
        .Caption = Sheet14.Range("H11").Value
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_Nhóm_VTHH_6").Slicers("Nhóm VTHH 61")
        .Caption = Sheet14.Range("I11").Value
    End With

    'Ð?i tên tiêu d?

    With SheetTheoDoiBienLoiNhuan
        .Range("Q11").Value = Sheet14.Range("D11").Value
        .Range("R11").Value = Sheet14.Range("E11").Value
        .Range("S11").Value = Sheet14.Range("F11").Value
        .Range("T11").Value = Sheet14.Range("G11").Value
        .Range("U11").Value = Sheet14.Range("H11").Value
        .Range("V11").Value = Sheet14.Range("I11").Value
    End With
End Sub

Function Ceil(Value As Double) As Double
    If Value = Int(Value) Then
        Ceil = Value
    Else
        Ceil = Int(Value) + 1
    End If
End Function

'bat dau Cac Function thiet lap phan trang
Function cbbPage(totalData As Double, limit As Integer, cbbBox As Object)
    Dim i As Integer
    Dim totalPage As Integer

    totalPage = Ceil(totalData / limit)

    With cbbBox
        .Clear
        For i = 1 To totalPage
            .AddItem i
        Next i
        .Value = 1
    End With
End Function

Function StartRecord(currentPage As Integer, limit As Integer) As Integer
    StartRecord = (currentPage - 1) * limit
End Function

Function UpdateChartDataRange(SheetChart As Worksheet, TenChart As String, SheetSourceData As Worksheet, r As String)
    Dim cht As ChartObject
    Set cht = SheetChart.ChartObjects(TenChart)
    cht.Chart.SetSourceData Source:=SheetSourceData.Range(r)
End Function

Function SetValuePage(totalData As Double, txtBox As Object, limit)
    Dim totalPage As Integer
    Dim Value As Integer

    totalPage = Ceil(totalData / limit)

    With txtBox
        If .Value = "" Or .Value < 0 Or IsNumeric(.Value) = False Then
            .Value = 1
        End If

        Value = .Value

        If .Value > totalPage Then
            .Value = totalPage
        Else
            .Value = Value
        End If

    End With
End Function

Function LastColumnWithValidData(rng As Range) As String
    Dim LastCol As Long
    Dim ws As Worksheet
    Dim cell As Range

    ' L?y sheet ch?a vùng d? li?u
    Set ws = rng.Worksheet

    ' Kh?i t?o c?t cu?i cùng
    LastCol = 0

    ' L?p qua t?ng ô trong vùng d? li?u
    For Each cell In rng
        ' Ki?m tra n?u ô không tr?ng và khác 0
        If Not IsEmpty(cell.Value) And cell.Value <> 0 Then
            ' Ki?m tra n?u ô có công th?c nhung k?t qu? là tr?ng
            If cell.HasFormula Then
                If Trim(cell.Text) <> "" Then
                    If cell.Column > LastCol Then
                        LastCol = cell.Column
                    End If
                End If
            Else
                If cell.Column > LastCol Then
                    LastCol = cell.Column
                End If
            End If
        End If
    Next cell

    ' Tr? v? ch? cái c?a c?t cu?i cùng
    If LastCol > 0 Then
        LastColumnWithValidData = Split(ws.Cells(1, LastCol).Address, "$")(1)
    Else
        LastColumnWithValidData = "No valid data"
    End If
End Function

Function totalPage(wSheet As Worksheet, r As String) As Integer
    Dim TongDuLieu As Double
    TongDuLieu = wSheet.Range(r).Value
    totalPage = Ceil(TongDuLieu / 10)
End Function

'Ket thuc Cac Function thiet lap phan trang

'Auth: Nguyen_Huu_Thanh
'Date By: 26/06/2024
'Description: Hàm format ngay lam viec voi ngay lam viec = 0 la ngay nghi le To mau Do

Public Function F_FormatNgayLamViec(wSheet As Worksheet)
    BatLimit
    With wSheet
        Dim dongCuoi As Integer
        Dim i As Integer
        dongCuoi = tinhdongcuoi("B12:B500")


        For i = 12 To dongCuoi
            If .Range("M" & i).Value = 0 Then
                'To ngay nghi le mau Do
                .Range("B" & i & ":L" & i).Select
                With Selection.Font
                    .Color = -16776961
                    .Bold = True
                End With
            Else
                'To ca ngay lam viec thanh mau den
                .Range("B" & i & ":L" & i).Select
                With Selection.Font
                    .ThemeColor = xlThemeColorLight1
                    .TintAndShade = 0
                End With
            End If
        Next i

        ' An Colum M
        Call F_Width("M", 0)
    End With
    TatLimit
End Function

'Auth: Nguyen_Huu_Thanh
'Date By: 26/06/2024
'Description: Sub nay dung de tai su dung code dua du lieu ra  data bao cao ngay, bao cao tuan, bao cao thang, bao cao nam
' Cua cac menu SO SANH KH-TH THEO ÐVKD, SO SANH KH-TH ThEO NVKD, SO SANH KH-TH THEO KHACH HANG

Sub GenerateQueryAndCallViewSheet(table As String, Nam As Integer, wSheet As Worksheet, Query2 As String)
    Dim Query As String
    Dim ID As LongLong
    Dim dbConn As Object

    With wSheet
        .Select

        ' Mo ket noi
        Set dbConn = ConnectToDatabase
        Call viewSheet(Query2, wSheet, "A1", dbConn)
        ID = .Range("A1").Value


        ' Query Bao cao cao ngay B -> M
        Query = "Select STT, NgayThang, Thu, Tuan, Thang, Quy, KY, KeHoach_Ngay , DoanhSoBan_Ngay , TyLeThucHien, (KeHoach_Ngay - DoanhSoBan_Ngay) As Thieu, NgayLamViec " & _
        "from " & table & "(" & Nam & ", " & ID & ") BC_Ngay order by STT"

        Call viewSheet(Query, wSheet, "B12", dbConn)

        'Query bao cao tuan >> X -> AC
        Query = "Select Tuan, (Select top 1 Thang from DM_NgayThang where Year(NgayThang) = " & Nam & " And DM_NgayThang.Tuan = BC_Ngay.Tuan order by NgayThang ) Thang, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan,  " & _
        " Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        "from " & table & "(" & Nam & ", " & ID & ") BC_Ngay " & _
        "Group by Tuan "

        Call viewSheet(Query, wSheet, "X12", dbConn)

        'Query bao cao thang >> AQ12
        Query = "Select Thang, Quy, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from " & table & "(" & Nam & ", " & ID & ") BC_Ngay " & _
        "Group by Thang, Quy "
        Call viewSheet(Query, wSheet, "AQ12", dbConn)

        ' Bao cao Quy >> Tu o BJ12
        Query = "Select Quy, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from " & table & "(" & Nam & ", " & ID & ") BC_Ngay " & _
        "Group by Quy"

        Call viewSheet(Query, wSheet, "BI12", dbConn)

        ' Bao cao Ky >> Tu o BI20

        Query = "Select Ky, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from " & table & "(" & Nam & ", " & ID & ") BC_Ngay " & _
        "Group by Ky"
        Call viewSheet(Query, wSheet, "BI20", dbConn)

        ' Bao cao Nam >> Tu o BI26

        Query = "Select '" & Nam & "' As Nam, " & _
        "Sum(Kehoach_Ngay) As Kehoach, Sum(DoanhSoBan_Ngay) As DoanhSoBan, " & _
        "Case when Sum(KeHoach_Ngay) = 0 Then 0 Else  Sum(DoanhSoBan_Ngay)/Sum(KeHoach_Ngay) end As TyLeThucHien, " & _
        "Sum(Kehoach_Ngay) - Sum(DoanhSoBan_Ngay) ConThieu " & _
        " from " & table & "(" & Nam & ", " & ID & ") BC_Ngay "

        Call viewSheet(Query, wSheet, "BI26", dbConn)

        'Dong Ket noi
        Call CloseDatabaseConnection(dbConn)

        Call formatSyleDataSS(wSheet)

    End With
End Sub

'Auth: Nguyen_Huu_Thanh
'Date By: 26/06/2024
'Description: Sub nay dung de tai su dung code format
'Cua cac menu SO SANH KH-TH THEO ÐVKD, SO SANH KH-TH ThEO NVKD, SO SANH KH-TH THEO KHACH HANG

Sub formatSyleDataSS(ws As Worksheet)
    With ws
        .Select
        .Columns("C:C").Select
        Selection.NumberFormat = "m/d/yyyy"
        .Columns("I:J").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"

        .Columns("K:K").Select
        Selection.NumberFormat = "0.00%"
        .Columns("L:L").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"

        .Columns("Z:AA").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"

        .Columns("AB:AB").Select
        Selection.NumberFormat = "0.00%"

        .Range("AC12:AC101").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"

        .Columns("As:AT").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"
        .Range("AU12:AU24").Select
        Selection.NumberFormat = "0.00%"
        .Range("AV12:AV151").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"
        .Range("AU154:AU156").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"

        .Columns("BJ:BK").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"
        .Columns("BM:BM").Select
        Selection.NumberFormat = "#,##0_);(#,##0)"

        .Columns("BL:BL").Select
        Selection.NumberFormat = "0.00%"
        .Columns("BN:BN").Select
        Selection.NumberFormat = "0%"
    End With
End Sub








