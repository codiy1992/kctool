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
Private Const DB_HOST = "192.168.5.3"
Private Const DB_USER = "root"
Private Const DB_PASS = "root"
Private Const DB_NAME = "codiy"
Private Const TB_INI = "kc_ini"
Private Const TB_MAIN = "kc_main"
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

Public Function GetComByKc(ByRef kc As Integer)
    Dim I As Integer
    Dim Sql As String
    Dim com() As Integer
    Dim RecS As New ADODB.Recordset
    Sql = "select * from " & TB_INI & " where kc = " & kc
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic ' 按读写方式打开
    If RecS.RecordCount > 0 Then
        ReDim com(0 To RecS.RecordCount - 1)
        For I = 0 To RecS.RecordCount - 1
            com(I) = RecS("com").Value
            RecS.MoveNext
        Next I
    End If
    GetComByKc = com()
    RecS.Close
    Conn.Execute "update " & TB_MAIN & " set status = 0 where status = 1 "
    Conn.Execute "update " & TB_MAIN & " set bind_status = -1 where bind_status >= 1  and bind_status != 9"
End Function

Public Function GetSwitchCard(ByRef strIccid As String) As String
    Dim Sql As String
    Dim RecS As New ADODB.Recordset
    Sql = "select * from " & TB_MAIN & " where use_status = 1 and kid != '' limit 1"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic ' 按读写方式打开
    If RecS.RecordCount > 0 Then
        strIccid = RecS("iccid").Value
        GetSwitchCard = RecS("kid").Value
    Else
        GetSwitchCard = ""
    End If
    RecS.Close
End Function
Public Function SetCardCanUse(ByRef strIccid As String)
    Conn.Execute "update " & TB_MAIN & " set use_status = 2 where use_status = 1 and iccid = '" & strIccid & "'"
    
End Function

Public Function IsUsing(strIccid As String, Optional strTime As String) As Boolean
    Dim Sql As String
    Dim RecS As New ADODB.Recordset
    Sql = "select * from " & TB_MAIN & " where iccid in (" & strIccid & ") and status =1 and use_status >= 1"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic ' 按读写方式打开
    If RecS.RecordCount > 0 Then
        IsUsing = True
        strTime = RecS("last_use_time")
    Else
        IsUsing = False
    End If
    RecS.Close
End Function
Public Function setUseTimeOut()
    Conn.Execute "update " & TB_MAIN & " set use_status = 0 where status = 1 and use_status >=1 and last_use_time <= UNIX_TIMESTAMP(date_add(now(),interval -5 minute))"
End Function
Public Function RegistCard(ByRef strIccid As String, _
                        ByVal strKid As String, _
                        Optional strImei As String, _
                        Optional strImsi As String, _
                        Optional strMobile As String) As String
    Dim tmp As Integer
    Dim Sql As String
    Dim RecS As New ADODB.Recordset
    Sql = "select * from " & TB_MAIN & " where iccid = '" & strIccid & "'"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic
    If RecS.RecordCount = 0 Then
        RecS.AddNew
        RecS("kid") = strKid
        RecS("iccid") = strIccid
        If strImei <> Empty Or strImei <> "" Then
            RecS("imei") = strImei
        End If
        If strImsi <> Empty Or strImsi <> "" Then
            RecS("imsi") = strImsi
        End If
        If strMobile <> Empty Or strMobile <> "" Then
            RecS("mobile") = strMobile
        End If
        RecS.Update
        RegistCard = ""
    Else
        If RecS("mobile") <> "" Then
            RegistCard = RecS("mobile")
        Else
            RegistCard = ""
        End If
        If strImei <> Empty Or strImei <> "" Then
            RecS("imei") = strImei
        End If
        If strImsi <> Empty Or strImsi <> "" Then
            RecS("imsi") = strImsi
        End If
        If strMobile <> Empty Or strMobile <> "" Then
            RecS("mobile") = strMobile
        End If
        tmp = Val(RecS("use_status"))
        'RecS("kid") = strKid
        RecS("status") = 1
        RecS.Update
    End If
    RecS.Close
    If tmp = 1 Then
        Conn.Execute "update " & TB_MAIN & " set status = 1,use_status = 2,last_use_time=UNIX_TIMESTAMP() where iccid = '" & strIccid & "'"
    End If
End Function
Public Function setCardClose(ByVal strIccid As String)
    Dim Sql As String
    Dim RecS As New ADODB.Recordset
    Sql = "update " & TB_MAIN & " set status = 0,use_status = 0 where iccid in (" & strIccid & ")"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic
End Function

Public Function SaveSMS(ByVal strIccid As String, strFrom As String, strCnt As String, strTime As String, Optional strMobile As String)
    'Dim Sql As String
    If strMobile <> Empty Then
        Conn.Execute "insert into " & TB_SMS_IN & " set iccid  = '" & strIccid & "',sms_from = '" & strFrom & "',cnt='" & strCnt & "',sms_time = '" & strTime & "',create_time = now(),mobile='" & strMobile & "'"
    Else
        Conn.Execute "insert into " & TB_SMS_IN & " set iccid  = '" & strIccid & "',sms_from = '" & strFrom & "',cnt='" & strCnt & "',sms_time = '" & strTime & "',create_time = now()"
    End If

'    Conn.Execute "insert into " & TB_MAIN & " set iccid  = '" & strIccid & "',`from` = '" & strFrom & "',cnt='" & strCnt & "',sms_time = '" & strTime & "',create_time = now()"
'    Dim RecS As New ADODB.Recordset
'    Sql = "select * from " & TB_SMS_IN & " where id = 1"
'    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic ' 按读写方式打开
'    RecS.AddNew
'    RecS("iccid") = strIccid
'    RecS("from") = strFrom
'    RecS("cnt") = strCnt
'    RecS("sms_time") = strTime
'    RecS("create_time") = Now()
'    If strMobile <> Empty Then
'        RecS("mobile") = strMobile
'    End If
'    RecS.Update
'    RecS.Close
End Function

Public Function NotExecBind(ByVal strActIccids As String)
    Dim I As Integer
    Dim Sql As String
    Dim execArr() As String
    Dim bind_status As Integer
    Dim RecS As New ADODB.Recordset
    If strActIccids = "''" Then
        NotExecBind = False
        Exit Function
    End If
    Sql = "select * from " & TB_MAIN & " where iccid in (" & strActIccids & ") and bind_status >= 1 and bind_status<=5 limit 5"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic ' 按读写方式打开
    If RecS.RecordCount > 0 Then
        ReDim execArr(0 To RecS.RecordCount - 1, 0 To 2)
        For I = 0 To RecS.RecordCount - 1
            execArr(I, 0) = RecS("iccid").Value
            execArr(I, 1) = RecS("id").Value
            execArr(I, 2) = RecS("bind_mobile").Value
            bind_status = RecS("bind_status")
            RecS("bind_status") = bind_status + 1
            RecS.Update
            RecS.MoveNext
        Next I
        NotExecBind = execArr()
    Else
        NotExecBind = False
    End If
    RecS.Close
End Function

Public Function SetBinded(ByVal strIccid As String, iBindId As Integer)
    Conn.Execute "update " & TB_MAIN & " set ccfc_mobile = bind_mobile,bind_status = 9 where id = " & iBindId & " and iccid = '" & strIccid & "'"
End Function
Public Function SetNotBind(ByVal strIccid As String, iBindId As Integer)
    Conn.Execute "update " & TB_MAIN & " set bind_status = -1 where id = " & iBindId & " and iccid = '" & strIccid & "' and bind_status > 5"
End Function

'- ------------------------------------------- -
' |   说明: 取出尚未发送的短信
'  ------------------------------------------- -
' |   参数: strActIccids (在线卡号如："'89860115841028567295','89860115841028567296'"）
'  ------------------------------------------- -
' | 返回值: 未发短信列表 (字符串数组)
'- ------------------------------------------- -
Public Function NotSendedSMS(ByVal strActIccids As String)
    Dim I As Integer
    Dim Sql As String
    Dim smsArr() As String
    Dim RecS As New ADODB.Recordset
    If strActIccids = "''" Then
        NotSendedSMS = False
        Exit Function
    End If
    Sql = "select * from " & TB_SMS_OUT & " where iccid in (" & strActIccids & ") and status = 1 group by iccid limit 10"
    RecS.Open Sql, Conn, adOpenKeyset, adLockOptimistic ' 按读写方式打开
    If RecS.RecordCount > 0 Then
        ReDim smsArr(0 To RecS.RecordCount - 1, 0 To 3)
        For I = 0 To RecS.RecordCount - 1
            smsArr(I, 0) = RecS("iccid").Value
            smsArr(I, 1) = RecS("id").Value
            smsArr(I, 2) = RecS("to").Value
            smsArr(I, 3) = RecS("cnt").Value
            RecS("status") = 2
            RecS.Update
            RecS.MoveNext
        Next I
        NotSendedSMS = smsArr()
    Else
        NotSendedSMS = False
    End If
    RecS.Close
End Function

Public Function SetSMSSended(ByVal strIccid As String, iSmsId As Integer)
    Conn.Execute "update " & TB_SMS_OUT & " set status = 3 , sms_time = now() where id = " & iSmsId & " and iccid = '" & strIccid & "'"
End Function
Public Function SetSMSNotSend(ByVal strIccid As String, iSmsId As Integer)
    Conn.Execute "update " & TB_SMS_OUT & " set status = -1 where id = " & iSmsId & " and iccid = '" & strIccid & "'"
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
