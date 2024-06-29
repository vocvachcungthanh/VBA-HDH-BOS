Attribute VB_Name = "M_KetNoi"
'Ket noi
Function connect(sqlQuery As String)
  Dim Cn As ADODB.Connection
  Dim StrCnn As String
  Dim Rs As ADODB.Recordset
  Dim SQLStr As String

  Set Cn = New ADODB.Connection
  Set Rs = New ADODB.Recordset

  StrCnn = KetNoiMayChu_KhachHang

  'Xu ly lenh
  SQLStr = sqlQuery

  Cn.Open StrCnn
  Rs.Open SQLStr, Cn, adOpenStatic

  Cn.Close

  Set Rs = Nothing
  Set Cn = Nothing
End Function

Function sqlGetRows(sqlQuery As String) As Variant
  Dim Cn As ADODB.Connection
  Dim StrCnn As String
  Dim Rs As ADODB.Recordset
  Dim SQLStr As String
  Set Rs = New ADODB.Recordset

  StrCnn = KetNoiMayChu_KhachHang
  'Xu ly lenh

  SQLStr = sqlQuery

  Set Cn = New ADODB.Connection
  Cn.Open StrCnn
  Rs.Open SQLStr, Cn, adOpenStatic

  Do While Not Rs.EOF
    sqlGetRows = Rs.GetRows()
  Loop
  Cn.Close

  Set Rs = Nothing
  Set Cn = Nothing
End Function

Sub viewSheet(sqlQuery As String, wSheet As Worksheet, r As String, dbConn As Object)
'  On Error Resume Next
  
  If Not dbConn Is Nothing Then
    Dim Rs As Object
'    dbConn.CommandTimeout = 60
    Set Rs = dbConn.Execute(sqlQuery)

    If Not Rs.EOF And Not Rs.BOF Then
      With wSheet
        .Select
        .Cells.Select
        Selection.RowHeight = 15
        .Range(r).CopyFromRecordset Rs
      End With
    End If

    Set Rs = Nothing
  Else
    MsgBox "Mat Ket noi csdl"
  End If
  'Xu ly lenh

End Sub

Sub viewSheetHeader(sqlQuery As String, wSheet As Worksheet, r As String, r2 As String, dbConn As Object)
'    On Error Resume Next
    
    If Not dbConn Is Nothing Then
        Dim Rs As Object
        Set Rs = dbConn.Execute(sqlQuery)

        If Not Rs.EOF And Not Rs.BOF Then
            With wSheet
                .Select
                .Cells.Select
                Selection.RowHeight = 15

                ' Add headers dynamically based on the field names
                For i = 1 To Rs.Fields.Count
'                    On Error Resume Next
                    .Range(r2).Offset(0, i - 1).Value = Rs.Fields(i - 1).Name
                Next i

                ' Copy data from the recordset starting from the second row (r)
                .Range(r).CopyFromRecordset Rs
            End With
        End If

        Set Rs = Nothing
    Else
        MsgBox "Mat Ket noi csdl"
    End If
    'Xu ly lenh
End Sub

Sub ViewListBox(Query As String, listBox As Object, dbConn As Object)

  If Not dbConn Is Nothing Then
    Dim Rs As Object

    Set Rs = dbConn.Execute(Query)

    If Not Rs.EOF And Not Rs.BOF Then
      Dim Data As Variant
      Data = Rs.GetRows()
      With listBox
        .Clear
        If UBound(Data, 2) <> 0 Then
          .List = Application.Transpose(Data)
        Else
          .Column = Application.Transpose(Data)
        End If

      End With
    End If

    Set Rs = Nothing
  Else
    MsgBox "Mat Ke noi csld"
  End If
End Sub

Function ConnectToDatabase() As Object
'  On Error Resume Next
  Dim isConnect As Boolean

  If Not isConnect Then

    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")

    Dim connectionString As String
    connectionString = KetNoiMayChu_KhachHang

    conn.Open connectionString

    If conn.State = 1 Then
      Set ConnectToDatabase = conn

      isConnect = True
    Else
      isConnect = False
      Set ConnectToDatabase = Nothing
    End If
  End If

End Function

Sub CloseDatabaseConnection(ByRef conn As Object)
  If Not conn Is Nothing Then
    If conn.State = 1 Then
      conn.Close
    End If
    Set conn = Nothing
  End If
End Sub

