Attribute VB_Name = "M_BaoCaoDoanhThuTheo_SP"
Public Sub VeBieuDo_BaoCao_DoanhThuTheo_SP()
    BatLimit
    F_R_DATA
    Call TiLeChiPhi
    Set wSheet = Sheet2
    Dim dongCuoi As Integer
    Dim CotCuoi As String
    Dim PhamViResize As String
    With wSheet
        .Select
        ActiveWorkbook.RefreshAll

        'Nhom 1
       
        dongCuoi = Sheet26.Range("E6").Value
    
        PhamViResize = "K8:L" & dongCuoi + 8
    
        Call UpdateChartDataRange(Sheet2, "Chart 46", Sheet26, PhamViResize)
        Call DinhDangBdNhiet("Chart 46")

        'Nhom 2
        dongCuoi = Sheet26.Range("R6").Value
    
        PhamViResize = "X8:Y" & dongCuoi + 8
        Call UpdateChartDataRange(Sheet2, "Chart 36", Sheet26, PhamViResize)
        Call DinhDangBdNhiet("Chart 36")

        'Nhom 3
         dongCuoi = Sheet26.Range("AE6").Value
    
        PhamViResize = "AK8:AL" & dongCuoi + 8
        Call UpdateChartDataRange(Sheet2, "Chart 13", Sheet26, PhamViResize)
        Call DinhDangBdNhiet("Chart 13")

        'Nhom 4
         dongCuoi = Sheet26.Range("AQ6").Value
    
        PhamViResize = "AW8:AX" & dongCuoi + 8
        Call UpdateChartDataRange(Sheet2, "Chart 41", Sheet26, PhamViResize)
        Call DinhDangBdNhiet("Chart 41")

        'Nhom 5
        dongCuoi = Sheet26.Range("BD6").Value
        PhamViResize = "BJ8:BK" & dongCuoi + 8
        Call UpdateChartDataRange(Sheet2, "Chart 42", Sheet26, PhamViResize)
        Call DinhDangBdNhiet("Chart 42")

        'Nhom 6
        dongCuoi = Sheet26.Range("BQ6").Value
        PhamViResize = "BW8:BX" & dongCuoi + 8
        Call UpdateChartDataRange(Sheet2, "Chart 44", Sheet26, PhamViResize)
        Call DinhDangBdNhiet("Chart 44")
        
    
    End With
    
    With Sheet2
        .txtBoxPhanTrangNhom1.Value = 1
        .txtBoxPhanTrangNhom2.Value = 1
        .txtBoxPhanTrangNhom3.Value = 1
        .txtBoxPhanTrangNhom4.Value = 1
        .txtBoxPhanTrangNhom5.Value = 1
        .txtBoxPhanTrangNhom6.Value = 1
        
        Call .ResizeNhom1
        Call .ResizeNhom2
        Call .ResizeNhom3
        Call .ResizeNhom4
        Call .ResizeNhom5
        Call .ResizeNhom6
    End With
    
    Set wSheet = Nothing
    TatLimit

    ThongBao_ThanhCong
End Sub

Sub TiLeChiPhi()
    Set wSheet = Sheet26
    With wSheet
        .Select

        '        Mo ket noi csdl
        Dim dbConn As Object
        Set dbConn = ConnectToDatabase
        Call viewSheet("exec TiLeChiPhi", Sheet26, "HJ4", dbConn)

        'Dong Ket noi
        Call CloseDatabaseConnection(dbConn)

        Columns("HK:HL").Select
        Selection.Style = "Percent"
        Selection.NumberFormat = "0.00%"
    End With
    Set wSheet = Nothing
    Call F_RESET_TABLE_PIVOT_SP

End Sub

Public Sub F_RESET_TABLE_PIVOT_SP()
    Dim dongCuoi
    Set wSheet = Sheet26
    With wSheet
        .Select

        If dongCuoi > 9 Then
            dongCuoi = tinhdongcuoi("A9:A1048576")


            .ListObjects("Table8").Resize Range("$K$8:$L$" & dongCuoi)

            dongCuoi = tinhdongcuoi("N9:N1048576")

            .ListObjects("Table9").Resize Range("$X$8:$Y$" & dongCuoi)

            dongCuoi = tinhdongcuoi("AA9:AA1048576")

            .ListObjects("Table7").Resize Range("$AI$8:$AJ$" & dongCuoi)

            dongCuoi = tinhdongcuoi("AL9:AL1048576")

            .ListObjects("Table10").Resize Range("$AT$8:$AU$" & dongCuoi)

            dongCuoi = tinhdongcuoi("AW9:AW1048576")

            .ListObjects("Table11").Resize Range("$BF$8:$BG$" & dongCuoi)

            dongCuoi = tinhdongcuoi("BI9:BI1048576")

            .ListObjects("Table12").Resize Range("$BS$8:$BT$" & dongCuoi)

            dongCuoi = tinhdongcuoi("EX8:EX1048576")

            .ListObjects("Table_LNTSP_1").Resize Range("$FB$7:$FF$" & dongCuoi)

            dongCuoi = tinhdongcuoi("FH8:FH1048576")

            .ListObjects("Table_LNTSP_2").Resize Range("$FL$7:$FP$" & dongCuoi)

            dongCuoi = tinhdongcuoi("FR8:FR1048576")

            .ListObjects("Table_LNTSP_3").Resize Range("$FV$7:$FZ$" & dongCuoi)

            dongCuoi = tinhdongcuoi("GB8:GB1048576")

            .ListObjects("Table_LNTSP_4").Resize Range("$GF$7:$GJ$" & dongCuoi)
            dongCuoi = tinhdongcuoi("GL8:GL1048576")

            .ListObjects("Table_LNTSP_5").Resize Range("$GP$7:$GT$" & dongCuoi)

            dongCuoi = tinhdongcuoi("GV8:GV1048576")

            .ListObjects("Table_LNTSP_6").Resize Range("$GZ$7:$HD$" & dongCuoi)
        End If

    End With

    Set wSheet = Nothing
End Sub



