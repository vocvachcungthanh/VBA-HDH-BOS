Attribute VB_Name = "M_BaoCaoSoLuongTheo_SP"
Sub VeBieuDo_BaoCaoSoLuongTheo_SP()
    BatLimit
    F_R_DATA
    Set wSheet = Sheet8
    With Sheet8
        .Select
        ActiveWorkbook.RefreshAll
    End With
    Set wSheet = Nothing
        With Sheet8
           .txtNhom1.Value = 1
           .txtNhom2.Value = 1
           .txtNhom3.Value = 1
           .txtNhom4.Value = 1
           .txtNhom5.Value = 1
           .txtNhom6.Value = 1
           
           Call .ResizeNhom1
           Call .ResizeNhom2
           Call .ResizeNhom3
           Call .ResizeNhom4
           Call .ResizeNhom5
           Call .ResizeNhom6
        End With
    TatLimit

    ThongBao_ThanhCong
End Sub

