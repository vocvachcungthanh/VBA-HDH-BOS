Attribute VB_Name = "Module_Duong"
Public isFormOpen As Boolean
Public fullScreen As Boolean
Public treeMenu As Boolean

Sub Initialize()
    If InStr(ThisWorkbook.FullName, "KD") > 0 Then
        fullScreen = True
        ttform = True
    Else
        fullScreen = False
    End If
    treeMenu = True
End Sub

'===========================================================================================================

Sub VeMenuCon(sheetName As String)
    Dim sh As Worksheet
    Dim numColumns As Integer
    Dim numRows As Integer
    Dim imageWidth As Integer
    Dim imageHeight As Integer
    Dim imageSpacing As Integer
    Dim leftPos As Integer
    Dim topPos As Integer
    Dim imageIndex, menuHeigt As Integer
    Dim image As Variant
    Dim imgName As Variant
    Dim label As Variant
    Dim path As String
    Dim menuHome, menuChung, Menu, menu1, menu2 As Variant
    Dim lastRow As Long
    
    image = Array("Settings.ico", "import.ico", "pass.ico", "User.ico", "config.ico", "Update.ico", "add.ico", _
            "bell.ico", "filter.ico", "save.ico", "ruler.ico", "export.ico", "load.jpg", "delete.jpg", "delete.jpg", "pass.ico")
            
    imgName = Array("CaiDat", "Import", "MatKhau", "TaiKhoan", "PhanQuyen", "CapNhat", "ThemMoi", _
                    "ThongBao", "Loc", "Luu", "Thuoc", "Export", "Load", "Xoa", "XoaSBH", "HienThi")
                    
    label = Array(UniConvert("Cafi ddawjt", "Telex"), UniConvert("Theem Excel", "Telex"), UniConvert("Maajt Khaaru", "Telex"), UniConvert("Tafi Khoarn", "Telex"), _
                UniConvert("Phaan Quyeefn", "Telex"), UniConvert("Caajp Nhaajt", "Telex"), UniConvert("Theem Mowsi", "Telex"), _
                UniConvert("Thoong baso", "Telex"), UniConvert("Lojc", "Telex"), UniConvert("Luwu", "Telex"), UniConvert("Thuwowsc", "Telex"), UniConvert("Xuaast Excel", "Telex"), _
                UniConvert("Tari taast car duwx lieeju", "Telex"), UniConvert("Xosa duwx lieeju", "Telex"), UniConvert("Xosa duwx lieeju SCTBH", "Telex"), UniConvert("Hieern Thij Menu", "Telex"))
                
    path = ThisWorkbook.path & "\Icon\"
    
    menuHome = Array(0, 7, 10, 11, 12, 13, 15)
    menuChung = Array(0, 7, 10, 11, 12, 13, 14, 15)
'    menu1 = Array(0, 5, 7, 8, 9, 10, 11)
'    menu2 = Array(0, 6, 7, 8, 10, 11, 12)
    
    numColumns = 3
    numRows = 3

    imageWidth = 38
    imageHeight = 38
    imageSpacing = 20

    leftPos = 25
    topPos = 35
    
    Set sh = Workbooks("Core.xlsb").Sheets("PhanQuyen")
    lastRow = sh.Range("B1048576").End(xlUp).Row
    
    If sheetName = "Data" Then
        Menu = menuHome
    Else
        Menu = menuChung
    End If
    
    imageIndex = 0
    
    For i = 0 To UBound(Menu)
'        For j = 1 To numColumns
            If imageIndex < UBound(Menu) + 1 Then

                Dim ctl As Control
                Dim img As MSForms.image
                Dim img1 As MSForms.image
                Dim lbName As MSForms.label
                Set img = frmMenuChucNang.Controls.Add("Forms.Image.1", "Img" & imgName(Menu(imageIndex)))
                Set lbName = frmMenuChucNang.Controls.Add("forms.label.1", "lb" & imgName(Menu(imageIndex)))
                With img
                    .Left = leftPos
                    .Top = topPos
                    .Width = imageWidth
                    .Height = imageHeight
                    .PictureSizeMode = fmPictureSizeModeStretch
                    .Picture = LoadPicture(path & image(Menu(imageIndex)))
                End With
                
                With lbName
                    .Left = leftPos - 2
                    .Top = topPos + img.Height + 8
                    .Width = 45
                    .Height = imageHeight
                    .Caption = label(Menu(imageIndex))
                    .TextAlign = fmTextAlignCenter
                End With
             
                
'                leftPos = leftPos + imageWidth + imageSpacing
                imageIndex = imageIndex + 1
              End If
'        Next j
        leftPos = 25
        topPos = topPos + imageHeight + imageSpacing + 15
    Next i
    
    If UBound(Menu) + 1 > 6 Then
        menuHeigt = UBound(Menu) + 28
    Else
        menuHeigt = UBound(Menu)
    End If
    Set sh = Nothing
'    frmMenuChucNang.Width = (numColumns * imageWidth) + ((numColumns + 1) * imageSpacing) + 2 * 4
'    frmMenuChucNang.Height = (numRows * imageHeight) + ((numRows + 1) * imageSpacing) + 2 * menuHeigt
End Sub

'=============================================================================================================================

' Export
Sub ExportSheet()
'    Application.Sheets(ActiveSheet.Name).Select
'    Application.EnableEvents = False
'    Application.Sheets(ActiveSheet.Name).Copy
    BatLimit
    Dim sourceSheet As Worksheet
    Dim destinationWorkbook As Workbook
    Dim destinationSheet As Worksheet
    Dim sourceWorkbook As Workbook
    Dim shName As String
    
    Set sourceWorkbook = Workbooks.Open(ActiveWorkbook.FullName)
    Set sourceSheet = sourceWorkbook.Worksheets(ActiveSheet.Name)
    
    shName = sourceSheet.Name
    
    Set destinationWorkbook = Workbooks.Add
        
    Set destinationSheet = destinationWorkbook.Worksheets(1)
    
    destinationSheet.Name = shName
    
    Application.CutCopyMode = False
    
    sourceSheet.Cells.Copy
    
    destinationSheet.Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme

    destinationSheet.DrawingObjects.Delete
    Application.ExecuteExcel4Macro "Show.toolbar(""Ribbon"",True)"
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayOutline = True
    ActiveWindow.DisplayZeros = True
    TatLimit
    
    destinationWorkbook.Activate
    
    destinationSheet.Range("D6").Select
    
    Set sourceWorkbook = Nothing
    Set sourceSheet = Nothing
    Set destinationWorkbook = Nothing
    Set destinationSheet = Nothing
End Sub

Function GetSheetZoom() As Double
    Dim ws As Worksheet
    Dim zoomValue As Integer
    
    Set ws = Sheets(ActiveSheet.Name)
    
    zoomValue = ws.parent.Windows(1).Zoom
    GetSheetZoom = zoomValue / 100
    
    Set ws = Nothing
End Function

'===================================================================================

Sub Hien_Thi_Menu_Con()
    If isFormOpen Then
        Unload frmMenuChucNang
        isFormOpen = False
    Else
        frmMenuChucNang.Show
        isFormOpen = True
    End If

End Sub



