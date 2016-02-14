Attribute VB_Name = "RegBas"
'Option Explicit
' Reg Keys
'Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const HKEY_USERS = &H80000003
Public IntrucData(4) As String
'
' For returned values in registry functions
Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NO_MORE_ITEMS = 259&

' Data type
Public Const REG_SZ = 1                         ' Null string
Public redonkey As String
Public rovernet As String
'
' 32 bits Windows API
'
Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Declare Function RegEnumKey Lib "advapi32" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal iSubKey As Long, ByVal lpszName As String, ByVal cchName As Long) As Long
Declare Function RegOpenKey Lib "advapi32" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpszSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Function QueryRegBase(ByVal Entry As String, Optional vKey) As String
    Dim buf As String
    Dim buflen As Long
    Dim hKey As Long
    Dim phkResult As Long
    If IsMissing(vKey) Then
        hKey = HKEY_CURRENT_USER
    Else
        hKey = CLng(vKey)
    End If
    On Local Error Resume Next
    buf = Space$(300)
    buflen = Len(buf)
    If Not RegQueryValue(hKey, Entry, buf, buflen) = 0 Then
      Call RegCreateKey(hKey, "Software\ConsolaSS", phkResult)
    End If
  
  
    If RegQueryValue(hKey, Entry, buf, buflen) = 0 Then
        If buflen > 1 Then
            QueryRegBase = Left$(buf, buflen - 1)
        Else
            QueryRegBase = ""
        End If
    End If
    
    On Local Error GoTo 0
End Function

'Public Sub aclave()
'Call RegSetValue(HKEY_CURRENT_USER, QueryRegBase("Software\ConsolaSS"), REG_DWORD, "luk", 3)
'End Sub

Public Sub aregistro(ruta, archivo)
    Dim sProgId As String
    Dim sDef As String
    Dim hKey As Long
    Dim phkResult As Long
    Dim lRet As Long
    Dim sValue As String
    Dim sKey As String
    sProgId = QueryRegBase("Software\ConsolaSS")
    'If Len(sProgId) Then
        sDef = "Software\ConsolaSS" & ruta
        hKey = HKEY_CURRENT_USER
        lRet = RegCreateKey(hKey, sDef, phkResult)
        If lRet = ERROR_SUCCESS Then
            sKey = ""
            sValue = archivo
            lRet = RegSetValue(phkResult, sKey, REG_SZ, sValue, Len(sValue))
            lRet = RegCloseKey(phkResult)
        End If
    'End If
End Sub

Public Function interp(ctcp, cudp, accion)
    'Mid(string, start[, length]) cut from a to b
    'InStr([start, ]string1, string2[, compare])
    Dim posi1, posi2 As String
    Dim i, z As Integer
    z = 0
    If Len(ctcp) <> 0 Then
        posi1 = 1
        posi2 = 1
        For i = 0 To 4
            posi2 = InStr(posi1, ctcp, ";")
            If posi2 = 0 Then posi2 = Len(ctcp)
            If posi1 >= Len(ctcp) Then Exit For
            If accion = "s" Then IntrucData(i) = "set naptserver tcp " & Mid(ctcp, posi1, posi2 - posi1) & " " & config.ipdft.Text
            If accion = "d" Then IntrucData(i) = "delete naptserver tcp " & Mid(ctcp, posi1, posi2 - posi1)
            posi1 = posi2 + 1
            z = i
        Next
    End If
    If Len(cudp) <> 0 Then
        posi1 = 1
        posi2 = 1
        If i <> 0 Then z = z + 1
        For i = z To 4
            posi2 = InStr(posi1, cudp, ";")
            If posi2 = 0 Then posi2 = Len(cudp)
            If posi1 >= Len(cudp) Then Exit For
            If accion = "s" Then IntrucData(i) = "set naptserver udp " & Mid(cudp, posi1, posi2 - posi1) & " " & config.ipdft.Text
            If accion = "d" Then IntrucData(i) = "delete naptserver udp " & Mid(cudp, posi1, posi2 - posi1)
            posi1 = posi2 + 1
        Next
    End If
End Function
Public Function encrip(dato)
Dim aux(30) As String
For i = 1 To Len(dato)
    aux(i) = Mid(dato, i, 1)
    aux(i) = Chr$(Asc(aux(i)) + 27 - i)
    encrip = encrip & aux(i)
Next
End Function
Public Function dencrip(dato)
Dim aux(30) As String
For i = 1 To Len(dato)
    aux(i) = Mid(dato, i, 1)
    aux(i) = Chr$(Asc(aux(i)) - 27 + i)
    dencrip = dencrip & aux(i)
Next
End Function
