Attribute VB_Name = "ModReg"
Option Explicit
Public AppLang As Integer
Private Const Pwd = "65C0B586B7E744BBA4EDBD57E227A66B"
Private Const HKEY_CURRENT_USER = &H80000001

Private Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Private Const REG_BINARY = 3                     ' Binario de formato libre
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "KERNEL32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function SHSetValue Lib "SHLWAPI.DLL" Alias "SHSetValueA" (ByVal hKey As Long, ByVal pszSubKey As String, ByVal pszValue As String, ByVal dwType As Long, pvData As String, ByVal cbData As Long) As Long

Function ReadEntry(FechaRet As String) As Boolean
Dim lenK&, hKey&, nKeys&
Dim elTime As FILETIME, sysTime As SYSTEMTIME
Dim retDate As Date
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Xiao", 0, KEY_READ, hKey) = 0 Then
        RegQueryInfoKey hKey, "B83C1C4F-10F71942", 17, 0, nKeys, lenK, 0, 0, 0, 0, 0, elTime
        FileTimeToSystemTime elTime, sysTime
        retDate = Format(sysTime.wDay & "-" & sysTime.wMonth & "-" & sysTime.wYear, "dd-MM-yyyy")
        RegCloseKey hKey
        FechaRet = retDate
    End If
End Function

Function RegNew(subFolder$, strData)
     Dim Ret, sDt$
     sDt = CStr(strData)
    'Create a new key
    RegCreateKey HKEY_CURRENT_USER, "Software\Xiao", Ret
    'Set the key's value
      Ret = SHSetValue(HKEY_CURRENT_USER, "Software\Xiao", subFolder, REG_SZ, ByVal sDt, CLng(LenB(StrConv(sDt, vbFromUnicode)) + 1))
    'close the key
    RegCloseKey Ret
End Function

Function SetNewValue(OnKey As String, newValue)
Dim Ret&
    Ret = SHSetValue(HKEY_CURRENT_USER, "Software\Xiao", OnKey, REG_SZ, ByVal newValue, CLng(LenB(StrConv(newValue, vbFromUnicode)) + 1))
End Function

Function GetReg(SubKey As String)
    Dim Ret, sRet$
    'Open the key
    RegOpenKey HKEY_CURRENT_USER, "Software\Xiao", Ret
    'Get the key's content
    'Decode sDt
    sRet = RegQueryStringValue(Ret, SubKey) ' "B83C1C4F-10F71942"
    GetReg = sRet
    'Close the key
    RegCloseKey Ret
End Function

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
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            'Dim strData As Integer
            Dim strData&
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function
