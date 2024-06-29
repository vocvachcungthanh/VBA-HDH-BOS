VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim isActive As Boolean

Private Sub cbbSheetNam_Click()
    With Sheet11
        .Range("C5").Value = .cbbSheetNam.Value

        If .cbbSheetNam.Value <> "" Then
            BatLimit
            Dim Nam As Long

            Set wSheet = SheetDataDonViKD
            With wSheet
                .Select
                Nam = SheetDataDonViKD.Range("C5")

                If Nam <> 0 Then
                    layDuLieuDV Nam
                Else
                    layDuLieuDV Year(Now)
                End If
            End With
            Set wSheet = Nothing
            F_StyleDV
            Sheet11.Select
            Sheet11.Range("A1").Select

            Range("B100").Select
            ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh



            Dim lr As Long
            With Sheet11
                .Select
                lr = tinhdongcuoi("G340:G399")

                If lr <= 339 Then
                    .ListObjects("Table13").Resize Range("$G$339:$I$340")
                Else
                    .ListObjects("Table13").Resize Range("$G$339:$I$" & lr)
                End If

                ActiveSheet.ChartObjects("Chart 26").Activate
                ActiveChart.FullSeriesCollection(2).Select
                ActiveChart.SetSourceData Source:=Range("Table13[#All]")

                'Resize DB_KHDTDVKD_TB
                ThisWorkbook.Sheets("KHDT theo DVKD").Select
                lr = tinhdongcuoi("B102:B337")

                If lr <= 101 Then
                    ThisWorkbook.Sheets("KHDT theo DVKD").ListObjects("DB_KHDTDVKD_TB").Resize Range("$F$101:$I$102")
                Else
                    ThisWorkbook.Sheets("KHDT theo DVKD").ListObjects("DB_KHDTDVKD_TB").Resize Range("$F$101:$I$" & lr)
                End If

            End With
            Call ScrollToTop
            TatLimit
        End If
    End With
End Sub

Private Sub Worksheet_Activate()
    If isActive = False Then
        Call ChayNeuActiveLanDau
        Call F_KhoiTaoNam(Sheet11.cbbSheetNam)
    End If

    isActive = True

    Sheet11.Select
    Call ScrollToTop
    Call hideall
End Sub

Sub ChayNeuActiveLanDau()
    On Error Resume Next
    Range("B100").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    BatLimit
    Dim dongCuoi
    Set wSheet = SheetDataDonViKD
    With SheetDataDonViKD
        .Select
        dongCuoi = tinhdongcuoi("C12:C1048576")

        If .Range("D12") <> "" Then
            ReSizeTableDV dongCuoi
            CongThucTinh dongCuoi
        End If
    End With
    Set wSheet = Nothing
End Sub

