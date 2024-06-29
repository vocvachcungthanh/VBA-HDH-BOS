VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim isActive As Boolean
Private Sub Worksheet_Activate()
    Call ScrollToTop
    
    BatLimit
    If isActive = False Then
        Set wSheet = Sheet14
        With wSheet
            .Select
    
            .Range("M12") = "=K12 * (1 - L12 / 100)"
    
            If .Range("B12") <> "" Then
                Range("M12").Select
                Selection.AutoFill Destination:=Range( _
                "TableMasterDataSanPham[Giá bán bình quân]")
                Range("TableMasterDataSanPham[Giá bán bình quân]").Select
            End If
        End With
         Set wSheet = Nothing
    End If
    
    isActive = True
    TatLimit
    Call hideall
    
    Call LamMoiMasterSanPham
End Sub

