Attribute VB_Name = "modRegistry"
'copyright information for this API module:

'KPD-Team 1998
'URL: http://www.allapi.net/
'E-Mail: KPDTeam@Allapi.net

Public RecentFiles(0 To 7) As String

Private Const REG_SZ = 1 ' Unicode nul terminated string
Private Const REG_BINARY = 3 ' Free form binary
Private Const HKEY_CURRENT_USER = &H80000001

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'retrieve nformation about the key
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                'Remove the unnecessary chr$(0)'s
                If strBuf <> "" Then
                    RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
                End If
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function
Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    'Open the key
    RegOpenKey hKey, strPath, Ret
    'Get the key's content
    GetString = RegQueryStringValue(Ret, strValue)
    'Close the key
    RegCloseKey Ret
End Function
Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Save a string to the key
    RegSetValueEx Ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    'close the key
    RegCloseKey Ret
End Sub
Sub SaveStringLong(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Set the key's value
    RegSetValueEx Ret, strValue, 0, REG_BINARY, CByte(strData), 4
    'close the key
    RegCloseKey Ret
End Sub
Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    Dim Ret
    'Create a new key
    RegCreateKey hKey, strPath, Ret
    'Delete the key's value
    RegDeleteValue Ret, strValue
    'close the key
    RegCloseKey Ret
End Sub

' the following code was designed by SaxxonPike for Recent File use in ZAP editor:

Sub GetRecentFiles()
    Dim x As Long
    For x = 0 To UBound(RecentFiles)
        RecentFiles(x) = GetString(HKEY_CURRENT_USER, "SOFTWARE\SaxxonPike\ZAP editor", "Recent" + CStr(x))
    Next x
End Sub

Sub SaveRecentFiles()
    Dim x As Long
    For x = 0 To UBound(RecentFiles)
        SaveString HKEY_CURRENT_USER, "SOFTWARE\SaxxonPike\ZAP editor", "Recent" + CStr(x), RecentFiles(x)
    Next x
End Sub

Sub AddRecentFile(newfile As String)
    Dim x As Long
    For x = 0 To UBound(RecentFiles)
        If UCase$(RecentFiles(x)) = UCase$(newfile) Then
            Exit Sub
        End If
    Next x
    For x = UBound(RecentFiles) To 1 Step -1
        RecentFiles(x) = RecentFiles(x - 1)
    Next x
    RecentFiles(0) = newfile
End Sub
