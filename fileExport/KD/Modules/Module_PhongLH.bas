Attribute VB_Name = "Module_PhongLH"

Sub CheckExcelFiles()
    Dim wb As Workbook
    Dim fileIsOpen As Boolean
    Dim isListedWorkbook As Boolean
    
    For Each wb In Application.Workbooks
        isListedWorkbook = False
        
        If wb.Name Like "CONG-VIEC.xlsb" Or wb.Name Like "Core.xlsb" Or wb.Name Like "KD.xlsb" Or wb.Name Like "CUNG-UNG.xlsb" Or wb.Name Like "TC.xlsb" Or wb.Name Like "KD-BAO-GIA.xlsb" Then
            isListedWorkbook = True
        End If
        
        If Not isListedWorkbook Then
            Dim ws As Worksheet
            
            wb.Activate
            For Each ws In wb.Worksheets
                With Application
                    .ExecuteExcel4Macro "Show.toolbar(""Ribbon"",True)"
                End With
                Exit For
            Next ws
        Else
            hideall
        End If
    Next wb
End Sub


