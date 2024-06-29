VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TreeNote_Menu_KD 
   Caption         =   "UserForm1"
   ClientHeight    =   12390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3765
   OleObjectBlob   =   "TreeNote_Menu_KD.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TreeNote_Menu_KD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Author: Nguyen Duy Tuan - duytuan@bluesofts.net
'Website: http://bluesofts.net - http://atoolspro.com
'Support online: https://www.facebook.com/groups/hocexcel/
'View learn BSTreeView: https://www.youtube.com/watch?v=dIczazD5bxM
Option Explicit
Dim tp As BSTaskPane

Private Sub UserForm_Initialize()
    Call AnSheet_TheoPhanQuyen
    'Create Task Pane by BSTaskPanes
    Dim TPs As New BSTaskPanes
    Set tp = TPs.Add(ConvertStr("KINH DOANH", AcTcvn3ToUnicode), Me, False)
    tp.AllowHide = False 'Prevent user from click [X] on task pane
    Set TPs = Nothing
    'AddIcons ' - No need in BSAC 3.0. Icons/Image added in design-time
    DoCreateBSTreeView
    'DoCreateBSTreeView_Easy ' For lean easy
End Sub

Private Sub UserForm_Terminate()
    'BSImageList1.ListImages.Clear 'Use this line if you run method "AddIcons"
    Set tp = Nothing 'Free Task Pane
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next 'Important!
    BSTreeview1.Height = Height - 26
    BSTreeview1.Width = Width
End Sub

Private Sub BSButton1_OnClick()
    Unload Me
End Sub

Private Sub DoCreateBSTreeView()
    Dim i As Integer, Node As BSTreeNode, sh As Worksheet
    Dim key As String, ParentKey As String, Text As String
    Dim PhanHe As String, phan_quyen As String, idx As Long
    Set sh = Workbooks("Core.xlsb").Sheets("PhanQuyen")
    BSTreeview1.ReadOnly = True
  '  BSTreeview1.hImageList = frmResource.BSImageList1.hImageList 'Link icons/images to BSTreeVew
    BSTreeview1.Items.Clear
    BSTreeview1.Font.Name = "Arial"
    BSTreeview1.Font.Size = 10
    BSTreeview1.CheckBoxes = False
    BSTreeview1.Items.BeginUpdate 'For faster. You must run EndUpdate later
    For i = 3 To sh.Range("A1000").End(xlUp).Row 'Dong cuoi du lieu
        'Nhan gia tri cho bien
        key = sh.Range("A" & i).Value
        Text = sh.Range("B" & i)
        ParentKey = sh.Range("D" & i).Value
      '  idx = sh.Range("a" & i).Value
        phan_quyen = sh.Range("H" & i).Value
        PhanHe = sh.Range("I" & i).Value
        'Tao cac node cho BSTreeView
        If PhanHe = "KD" Then
        If phan_quyen = "1" Then
            If ParentKey = "" Then  'Parent in root
                Set Node = BSTreeview1.Items.Add(, key, Text)
    
            Else 'Child
                Set Node = BSTreeview1.Items.Add(ParentKey, key, Text)
            End If
        End If
        End If
    Next i
    
    If BSTreeview1.Items.Count > 0 Then
        BSTreeview1.Items.EndUpdate
        BSTreeview1.FullExpand
        BSTreeview1.Items(0).MakeVisible
    Else
        MsgBox "You have not authorightiy for this function"
    End If
End Sub

Private Sub BSTreeview1_OnNodeAfterEdit(ByVal Node As BSAC.BSTreeNode, Text As String)
    'sh.Cells(Node.Key, 2).Value = Text
End Sub

Private Sub BSTreeview1_OnNodeCheck(ByVal Node As BSAC.BSTreeNode, ByVal IsChecked As Boolean)
    Dim n As BSTreeNode
    On Error GoTo lbEndSub
    If Node.HasChildren Then
        Set n = Node.GetFirstChild
        While Not n Is Nothing
            n.Checked = IsChecked
            Set n = n.GetNextSibling
        Wend
    End If
lbEndSub:
    Debug.Print "ERROR:" & Err.Description
End Sub

Private Sub BSTreeview1_OnNodeClick(ByVal Node As BSAC.BSTreeNode)
    On Error Resume Next
    Dim sh As Worksheet, i As Integer, key As String, Row As Integer
    Set sh = Workbooks("Core.xlsb").Sheets("PhanQuyen")
    For i = 3 To sh.Range("B1000").End(xlUp).Row
        key = sh.Range("A" & i).Value
        If key = Node.key Then
            Row = i
            Exit For
        End If
    Next i
    Dim Link_Sheet As String
    Dim path As String
    Dim file As Workbook
    path = Application.ThisWorkbook.path
    
    Link_Sheet = sh.Range("E" & Row).Value
    
    If Link_Sheet = "KHDT theo NVKD" Then
       Link_Sheet = "Data KHDT NVKD"
    End If
    
    If Link_Sheet = "Lich_Su_Bao_Gia_Khach_Hang" Then
        path = path & "\KD-BAO-GIA.xlsb"
        
        On Error Resume Next
        Set file = Workbooks.Open(path)
        Application.GoTo Workbooks(file.Name).Sheets(Link_Sheet).Range("A1")
    End If
    
    If Link_Sheet = "" Then
        ThongBao_ChucNangChuaCo
    Else
      Application.GoTo Workbooks("KD.xlsb").Sheets(Link_Sheet).Range("A1")
    End If
End Sub

Private Sub Collap_OnClick()
    DoExpendTree False
End Sub

Private Sub ExpendAll_OnClick()
    DoExpendTree True
End Sub

Private Sub DoExpendTree(Expend As Boolean)
    Dim Node As BSTreeNode
    For Each Node In BSTreeview1.Items
        DoExpendNode Node, Expend
    Next
End Sub

Private Sub DoExpendNode(Node As BSTreeNode, Expend As Boolean)
    Dim n As BSTreeNode
    Dim i&
    Set n = Node
    While Not n Is Nothing
        n.Expanded = Expend
        If n.HasChildren > 0 Then  'N.Count > 0
            DoExpendNode n.GetFirstChild, Expend
        End If
        Set n = n.GetNextSibling 'Next node
    Wend
End Sub

Private Sub AddIcons()
    'You can add icons by command bellow:
    Dim i&, sh As Worksheet
    Set sh = ThisWorkbook.Sheets("IMAGES")
    With frmResource
        .BSImageList1.SetSize 16, 16
        .BSImageList1.Transparent = False
        .BSImageList1.TransparentMode = tmAuto
        .BSImageList1.TransparentGraphicBeforeAdd = True
        For i = 4 To sh.Range("A6000").End(xlUp).Row
           .BSImageList1.ListImages.Add ThisWorkbook.path & "\" & sh.Cells(i, 1).Value
        Next i
    End With
End Sub

