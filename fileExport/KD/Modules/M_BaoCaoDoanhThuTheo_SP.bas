Public Sub VeBieuDo_BaoCao_DoanhThuTheo_SP()
    BatLimit
    F_R_DATA
    Call TiLeChiPhi
    Set wSheet = Sheet2
    With wSheet
        .Select
        ActiveWorkbook.RefreshAll

        'Nhom 1
        Call select_data("Chart 46", "Table8[#All]", "Pivot SP")
        Call DinhDangBdNhiet("Chart 46")

        'Nhom 2
        Call select_data("Chart 36", "Table9[#All]", "Pivot SP")
        Call DinhDangBdNhiet("Chart 36")

        'Nhom 3
        Call select_data("Chart 13", "Table7[#All]", "Pivot SP")
        Call DinhDangBdNhiet("Chart 13")

        'Nhom 4
        Call select_data("Chart 41", "Table10[#All]", "Pivot SP")
        Call DinhDangBdNhiet("Chart 41")

        'Nhom 5
        Call select_data("Chart 42", "Table11[#All]", "Pivot SP")
        Call DinhDangBdNhiet("Chart 42")

        'Nhom 6
        Call select_data("Chart 44", "Table12[#All]", "Pivot SP")
        Call DinhDangBdNhiet("Chart 44")
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
        Call viewSheet("exec TiLeChiPhi", Sheet26, "HF4", dbConn)

        'Dong Ket noi
        Call CloseDatabaseConnection(dbConn)

        Columns("GW:GX").Select
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

Public Function KhoiTaoCbbBoxPageSP()

    Dim TongDuLieu As Double

    TongDuLieu = Sheet26.Range("EZ6")
    Call cbbPage(TongDuLieu, 10, Sheet10.cbbPhanTrangLNNhom1)
End Function 