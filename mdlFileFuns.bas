Attribute VB_Name = "mdlFileFuns"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const INI_CCID_MOBILE_PATH = "ccid.ini"

Public Function GetIni(ByVal strSession As String, strKey As String) As String
    WritePrivateProfileString "Label2.Caption", "Label3.Caption", "Text1.Text", INI_CCID_MOBILE_PATH
End Function

Public Function SetIni(ByVal strSession As String, strKey As String, strVal As String) As Long
    Dim buff As String
    buff = String(255, 0)
    SetIni = GetPrivateProfileString("SESSIon", "key1", "asdasd", buff, 256, INI_CCID_MOBILE_PATH)
End Function
