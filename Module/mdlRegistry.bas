Attribute VB_Name = "mdlRegistry"
Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDesc As Long
    bInheritHandle As Long
End Type

Private SECURITY_ATT As SECURITY_ATTRIBUTES

Public Const HKEY_CURRENT_USER As Long = &H80000001

Public Const KEY_QUERY_VALUE As Long = &H1
Public Const KEY_SET_VALUE As Long = &H2
Public Const KEY_CREATE_SUB_KEY As Long = &H4
Public Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Public Const KEY_NOTIFY As Long = &H10
Public Const KEY_CREATE_LINK As Long = &H20
Public Const KEY_READ As Long = &H20019
Public Const KEY_WRITE As Long = &H20006
Public Const KEY_ALL_ACCESS As Long = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + &H20000

Public Const REG_SZ As Integer = 1
Public Const REG_DWORD As Integer = 4

Public Const KEYS_SYS_INFO As String = "SOFTWARE\MJSTONE\INVENTORY_COMPANY\"
Public Const KEYS_SYS_INFO_SERVER1 As String = "SOFTWARE\MJSTONE\INVENTORY_SERVER\"
Public Const KEYS_SYS_INFO_SERVER2 As String = "SOFTWARE\MJSTONE\FINANCE_SERVER\"
Public Const KEYS_SYS_INFO_SERVER3 As String = "SOFTWARE\MJSTONE\ACCOUNTING_SERVER\"
Public Const KEYS_SYS_INFO_RUN As String = "Software\Microsoft\Windows\CurrentVersion\Run"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal lKey As Long, ByVal strSubKey As String, ByVal lOptions As Long, ByVal lDesired As Long, ByRef lResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal lKey As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As Any, ByRef lpcbData As Long) As Long

Public Function WriteToRegistry(ByRef lngKeyRoot As Long, ByRef strKeyName As String) As Long
    Dim lngRegKey As Long

    SECURITY_ATT.lpSecurityDesc = 0
    SECURITY_ATT.bInheritHandle = True
    SECURITY_ATT.nLength = Len(SECURITY_ATT)

    WriteToRegistry = RegCreateKeyEx(lngKeyRoot, strKeyName, 0, "", 0, KEY_ALL_ACCESS, SECURITY_ATT, lngRegKey, 0)
End Function

Public Function WriteValueRegistry(ByRef lngKeyRoot As Long, ByRef strKeyName As String, ByRef strSubKeys As String, ByRef strValueData As String) As Long
    Dim lngRegKey As Long

    WriteValueRegistry = RegCreateKey(lngKeyRoot, strKeyName, lngRegKey)
    WriteValueRegistry = RegSetValueEx(lngRegKey, strSubKeys, 0, REG_SZ, ByVal strValueData, Len(strValueData))
End Function

Public Function OpenRegistry( _
    ByRef lngKeyRoot As Long, _
    ByRef strKeyName As String, _
    ByRef lngValueBack As Long) As Long
    Dim lngRegKey As Long

    OpenRegistry = RegOpenKeyEx(lngKeyRoot, strKeyName, 0, KEY_ALL_ACCESS, lngRegKey)
    
    lngValueBack = lngRegKey
End Function

Public Function ReadValueRegistry( _
    ByRef lngKeyValueRegistry As Long, _
    ByRef strSubKeys As String, _
    ByRef lngTypeBack As Long, _
    ByRef strDataBack As String, _
    ByRef lngSizeBack As Long) As Long
    strDataBack = String$(1024, 0)
    lngSizeBack = 1024

    ReadValueRegistry = _
        RegQueryValueEx(lngKeyValueRegistry, strSubKeys, 0, lngTypeBack, strDataBack, lngSizeBack)
End Function

Public Function DeleteKeysRegistry(ByRef lngKeyRoot As Long, ByRef strKeyName As String) As Long
    DeleteKeysRegistry = RegDeleteKey(lngKeyRoot, strKeyName)
End Function

Public Function DeleteSubKeysRegistry( _
    ByRef lngKeyRoot As Long, _
    ByRef strKeyName As String, _
    ByRef strSubKeys As String) As Long
    Dim lngRegKey As Long

    DeleteSubKeysRegistry = RegOpenKeyEx(lngKeyRoot, strKeyName, 0, KEY_ALL_ACCESS, lngRegKey)
    
    If DeleteSubKeysRegistry = 0 Then DeleteSubKeysRegistry = RegDeleteValue(lngRegKey, strSubKeys)
End Function

Public Function CloseRegistry(ByRef lngKeyValueRegistry As Long) As Long
    CloseRegistry = RegCloseKey(lngKeyValueRegistry)
End Function
