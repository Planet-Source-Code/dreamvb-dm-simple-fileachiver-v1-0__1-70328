Attribute VB_Name = "dReg"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const ERROR_SUCCESS As Long = 0&
Private Const REG_EXPAND_SZ = 2
Private Const HKEY_CLASSES_ROOT = &H80000000

Public Function GetFileType(FleExt As String) As String
Dim sType As String

    'Get File Type
    sType = RegGetStrValue(FleExt, vbNullString)
    sType = RegGetStrValue(sType, vbNullString)
    
    If (LenB(sType) <> 0) Then
        'Return FileType found
        GetFileType = sType
    Else
        'Just return the File's Ext
        GetFileType = FleExt & " File"
    End If
    
    sType = vbNullString
    
End Function

Private Function RegGetStrValue(KeyPath As String, KeyName As String) As String
Dim sBuffer As String
Dim lSize As Long
Dim sRegKey As Long

    If RegOpenKeyEx(HKEY_CLASSES_ROOT, KeyPath, 0&, KEY_QUERY_VALUE, sRegKey) <> ERROR_SUCCESS Then
        Exit Function
    ElseIf RegQueryValueEx(sRegKey, KeyName, 0&, REG_EXPAND_SZ, ByVal 0&, lSize) <> ERROR_SUCCESS Then
        Exit Function
    Else
        'Return Found String
        sBuffer = Space(lSize - 1)
        RegQueryValueEx sRegKey, KeyName, 0&, REG_EXPAND_SZ, ByVal sBuffer, lSize
        'Clsoe open key
        RegCloseKey sRegKey
    End If
    
    RegGetStrValue = sBuffer
    'Clear up
    sBuffer = vbNullString
    lSize = 0
    
End Function


