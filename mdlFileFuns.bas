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



Public Sub SaveInfoToFile(ByVal SaveString As String, Optional ByVal InfoFileName As String)
    Dim nFileNo As Long
    Dim strFileName As String
    Dim strAppend As String
    
On Error GoTo ErrorSave
    
    nFileNo = FreeFile()
    If InfoFileName = "" Then InfoFileName = "SendReord.txt"
    strFileName = App.Path & "\" & InfoFileName 'SendReord.txt"
    
    strAppend = "==========================================" & vbCrLf & _
                SaveString ' & vbCrLf

    Open strFileName For Append Access Write Shared As #nFileNo
        Print #nFileNo, strAppend
    Close #nFileNo
    'Form1.Caption = "±£´æÍê±Ï"
    
    Exit Sub
ErrorSave:
    MsgBox "Error:" & Err & "." & vbCrLf & Err.Description
End Sub


Public Sub SaveInitSettings()
    
    Dim iFileNo As Integer
    Dim strFileName As String, strTmp As String, strTmp1 As String
    
'On Error Resume Next

    iFileNo = FreeFile()
    strFileName = App.Path & "\sys.set"
    If Dir(strFileName) <> "" Then Kill (strFileName)
    Open strFileName For Binary Access Write As #iFileNo
        Put #iFileNo, , g_SysInfo
    Close #iFileNo
    
End Sub


Public Function LoadInitSettings() As Boolean
    Dim iFileNo As Integer
    Dim strFileName As String, strTmp As String, strExprmtName As String
    Dim n As Long

On Error GoTo ErrorLoadIni

    LoadInitSettings = False
    iFileNo = FreeFile()
    strFileName = App.Path & "\sys.set"
    If Dir(strFileName) <> "" Then
        If FileLen(strFileName) > 0 Then
            Open strFileName For Binary Access Read As #iFileNo
                '-------
                Get #iFileNo, , g_SysInfo
                LoadInitSettings = True
                Err.Clear
ErrorLoadIni:
            If Err <> 0 Then MsgBox Err.Description
            Close #iFileNo
        End If
    End If
        
End Function


