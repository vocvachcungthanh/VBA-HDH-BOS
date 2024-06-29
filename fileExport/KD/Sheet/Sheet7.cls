VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 Sub VeThucHien()
    BatLimit
    With Sheet7
        .Select
        Range("B4").Select
        .PivotTables("PivotTable6").PivotCache.Refresh
    End With
    
    With Sheet6
    .Select
    .Range("A1").Select
    End With
    TatLimit
End Sub

