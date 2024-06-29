VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenuChucNang 
   Caption         =   "Menu"
   ClientHeight    =   12105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "frmMenuChucNang.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmMenuChucNang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myIMG() As MenuControl

Private Sub UserForm_Initialize()
    Dim bg_color As Variant, kstr As Variant
    Dim obj As Object, ImgPointer As Long
    
    Dim ws As Worksheet
    Dim leftPos  As Double
    Dim topPos  As Double
    Dim shape As shape
    
    Set ws = ActiveSheet
    
    bg_color = RGB(255, 255, 255)
    kstr = RGB(217, 210, 227)
    
    Me.BackColor = bg_color
    
    VeMenuCon ActiveSheet.Name
    
    BSTaskPaneX1.Create Me
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ReDim myIMG(1 To Me.Controls.Count)
    
    For Each obj In Me.Controls
        If TypeName(obj) = "Image" Then
            ImgPointer = ImgPointer + 1
            Set myIMG(ImgPointer) = New MenuControl
            Set myIMG(ImgPointer).aImage = obj
            myIMG(ImgPointer).aImage.SpecialEffect = fmSpecialEffectFlat
            myIMG(ImgPointer).aImage.BorderStyle = fmBorderStyleSingle
        End If
    Next obj
    
    ReDim Preserve myIMG(1 To ImgPointer)

End Sub

