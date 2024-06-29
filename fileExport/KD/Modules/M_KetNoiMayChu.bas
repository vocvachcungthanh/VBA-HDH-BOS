Attribute VB_Name = "M_KetNoiMayChu"
Option Explicit


#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hModule As LongPtr) As Long
    Private Declare PtrSafe Function GetConnectionString Lib "DB_Connection.dll" (ByVal str As Variant, _
                                                                    ByVal str2 As Variant) As Variant
#Else
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hModule As Long) As Long
    Private Declare Function GetConnectionString Lib "DB_Connection.dll" (ByVal str As Variant, _
                                                                    ByVal str2 As Variant) As Variant
#End If

Function KetNoiMayChu_KhachHang() As String
   Dim mayChu As String
    Dim csdl As String
    mayChu = Workbooks("Core.xlsb").Sheets("PhanQuyen").Range("AA1").Value
    csdl = Workbooks("Core.xlsb").Sheets("PhanQuyen").Range("AA4").Value

    KetNoiMayChu_KhachHang = "Driver={SQL Server};Server=" & mayChu & ";Database=" & csdl & ";Uid=QuanTriBOS;Pwd=Bos58ToHuu;"
'    KetNoiMayChu_KhachHang = KetNoiMayChu_DLL
End Function


Function KetNoiMayChu_DLL() As String
    Dim mayChu As String
    Dim csdl As String
    Dim hModule As LongPtr
    Dim pFunc As LongPtr
    Dim dllPath As String
    Dim sConnect As String
    Dim functionAddress As LongPtr
    
    mayChu = Workbooks("Core.xlsb").Sheets("PhanQuyen").Range("AA1").Value

    csdl = Workbooks("Core.xlsb").Sheets("PhanQuyen").Range("AA4").Value
    
    #If VBA7 Then
        dllPath = ThisWorkbook.path & "\Win64\DB_Connection.dll"
    #Else
        dllPath = ThisWorkbook.path & "\Win32\DB_Connection.dll"
    #End If
    
    hModule = LoadLibrary(dllPath)
    
    If hModule <> 0 Then
        functionAddress = GetProcAddress(hModule, "GetConnectionString")
        If functionAddress <> 0 Then
            KetNoiMayChu_DLL = GetConnectionString(mayChu, csdl)
'            MsgBox KetNoiMayChu_DLL
            FreeLibrary hModule
            Exit Function
        Else
            Application.Assistant.DoAlert "BOS " & UniConvert(" Thoong baso", "Telex"), UniConvert("Không tifm thaasy hafm trong file DLL.", "Telex"), msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
            FreeLibrary hModule
            Exit Function
        End If
        FreeLibrary hModule
    Else
        Application.Assistant.DoAlert "BOS " & UniConvert("Thoong baso", "Telex"), UniConvert("Không tifm thaasy file DLL.", "Telex"), msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
        FreeLibrary hModule
        Exit Function
    End If
End Function






