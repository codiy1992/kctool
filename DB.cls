VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Conn As ADODB.Connection
Private Const DB_HOST = "localhost"
Private Const DB_USER = "root"
Private Const DB_PASS = "root"
Private Const DB_NAME = "codiy"
Private Const TB_BIND = "kc_main"
Private Const TB_SMS_IN = "kc_sms_in"
Private Const TB_SMS_OUT = "kc_sms_out"

Private Sub Class_Initialize()
    Set Conn = New ADODB.Connection
    Conn.Open "DRIVER={MySQL ODBC 5.3 Unicode Driver};" & _
              "SERVER=" & DB_HOST & ";" & _
              "DATABASE=" & DB_NAME & ";" & _
              "UID=" & DB_USER & ";PWD=" & DB_PASS & ";" & _
              "OPTION=3;stmt=SET NAMES UTF8"
    Conn.CursorLocation = adUseClient
End Sub

Private Sub Class_Terminate()
    If Conn.State = True Then
        Conn.Close
    End If
End Sub

Public Function Query(ByRef strSql As String)
    Dim RecS As New ADODB.Recordset
    RecS.Open strSql, Conn
    Query = RecS(0)
    RecS.Close
End Function

Public Function InitCard(ByRef strIccid As String, _
                        Optional strImei As String, _
                        Optional strImsi As String, _
                        Optional strMobile As String)
    Dim Sql As String
    Dim RecS As New ADODB.Recordset
    Sql = "select * from " & TB_BIND & " where iccid = '" & strIccid & "'"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic
    If RecS.RecordCount = 0 Then
        RecS.AddNew
        RecS("iccid") = strIccid
        If strImei <> Empty Then
            RecS("imei") = strImei
        End If
        If strImsi <> Empty Then
            RecS("imsi") = strImsi
        End If
        If strMobile <> Empty Then
            RecS("mobile") = strMobile
        End If
        RecS.Update
    Else
        If strImei <> Empty Then
            RecS("imei") = strImei
        End If
        If strImsi <> Empty Then
            RecS("imsi") = strImsi
        End If
        If strMobile <> Empty Then
            RecS("mobile") = strMobile
        End If
        RecS.Find "iccid = " & strIccid
        RecS.Update
    End If
    RecS.Close
End Function

Public Function SaveSMS(ByVal strIccid As String, strFrom As String, strCnt As String, strTime As String, Optional strMobile As String)
    Dim Sql As String
    Dim RecS As New ADODB.Recordset
    Sql = "select * from " & TB_SMS_IN & " where id = 1"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic ' 按读写方式打开
    RecS.AddNew
    RecS("iccid") = strIccid
    RecS("from") = strFrom
    RecS("cnt") = strCnt
    RecS("time") = strTime
    If strMobile <> Empty Then
        RecS("mobile") = strMobile
    End If
    RecS.Update
    RecS.Close
End Function

'- ------------------------------------------- -
' |   说明: 取出尚未发送的短信
'  ------------------------------------------- -
' |   参数: strActIccids (在线卡号如："'89860115841028567295','89860115841028567296'"）
'  ------------------------------------------- -
' | 返回值: 未发短信列表 (字符串数组)
'- ------------------------------------------- -
Public Function NotSendedSMS(ByVal strActIccids As String)
    Dim i As Integer
    Dim Sql As String
    Dim smsArr() As String
    Dim RecS As New ADODB.Recordset
    Sql = "select * from " & TB_SMS_OUT & " where iccid in (" & strActIccids & ") and status = 0"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic ' 按读写方式打开
    ReDim smsArr(0 To RecS.RecordCount - 1, 0 To 3)
    For i = 0 To RecS.RecordCount - 1
        smsArr(i, 0) = RecS("iccid").Value
        smsArr(i, 1) = RecS("mobile").Value
        smsArr(i, 2) = RecS("to").Value
        smsArr(i, 3) = RecS("cnt").Value
        RecS.MoveNext
    Next i
    NotSendedSMS = smsArr
    RecS.Close
End Function
'- ------------------------------------------- -
'  函数说明：GB2312转换为UTF8
'- ------------------------------------------- -
Public Function GB2312ToUTF8(strIn As String, Optional ByVal ReturnValueType As VbVarType = vbString) As Variant
    Dim adoStream As Object

    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 2 'adTypeText
    adoStream.Open
    adoStream.WriteText strIn
    adoStream.Position = 0
    adoStream.Type = 1 'adTypeBinary
    GB2312ToUTF8 = adoStream.Read()
    adoStream.Close

    If ReturnValueType = vbString Then GB2312ToUTF8 = Mid(GB2312ToUTF8, 1)
End Function

'- ------------------------------------------- -
'  函数说明：UTF8转换为GB2312
'- ------------------------------------------- -
Public Function UTF8ToGB2312(ByVal varIn As Variant) As String
    Dim bytesData() As Byte
    Dim adoStream As Object

    bytesData = varIn
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Charset = "utf-8"
    adoStream.Type = 1 'adTypeBinary
    adoStream.Open
    adoStream.Write bytesData
    adoStream.Position = 0
    adoStream.Type = 2 'adTypeText
    UTF8ToGB2312 = adoStream.ReadText()
    adoStream.Close
End Function
