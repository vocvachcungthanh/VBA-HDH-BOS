VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MenuControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents aImage As MSForms.image
Attribute aImage.VB_VarHelpID = -1

Private Sub aImage_Click()
    Dim msgValue As VbMsgBoxResult
    
    Unload frmMenuChucNang
   
  
    Select Case aImage.Name
        Case Is = "ImgCaiDat"
            Application.Run "Core.xlsb!Moform_CaiDat"
            Exit Sub
        Case Is = "ImgThuoc"
            ruller
            Exit Sub
        Case Is = "ImgExport"
            ExportSheet
            Exit Sub
        Case Is = "ImgThongBao"
            
            Exit Sub
        Case Is = "ImgLoad"
            BatLimit
                F_TaiTatCaDuLieu
            TatLimit
            Exit Sub
        Case Is = "ImgXoa"
            
            msgValue = Application.Assistant.DoAlert(UniConvert("Carnh baso", "Telex"), UniConvert("Bajn cos chawsc chawsn muoosn xosa duwx lieeju treen file khoong?", "Telex"), _
            msoAlertButtonYesNo, msoAlertIconWarning, 0, 0, 0)
            
            If msgValue = vbYes Then
                BatLimit
                    F_ClearAll
                 TatLimit
                  Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), UniConvert("Thafnh coong!", "Telex"), msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
            End If
            Exit Sub
        Case Is = "ImgXoaSBH"
            
            msgValue = Application.Assistant.DoAlert(UniConvert("Carnh baso", "Telex"), UniConvert("Bajn cos chawsc " & _
            "chawsn muoosn xosa duwx lieeju cura Soor chi tieest basn hafng?", "Telex"), _
            msoAlertButtonYesNo, msoAlertIconWarning, 0, 0, 0)
            
            If msgValue = vbYes Then
                BatLimit
                ClearSCTBH
             TatLimit
                    Application.Assistant.DoAlert UniConvert("Thoong baso", "Telex"), UniConvert("Thafnh coong!", "Telex"), msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
            End If
            Exit Sub
        Case Is = "ImgHienThi"
            CheckExcelFiles
            Exit Sub
    End Select

End Sub

Private Sub aImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    aImage.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub aImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    aImage.SpecialEffect = fmSpecialEffectEtched
End Sub

Private Sub aImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    aImage.SpecialEffect = fmSpecialEffectFlat
End Sub

