VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Com"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private blIsWaiting As Boolean ' 是否等待下一帧数据
Public blIsOpen As Boolean     ' 是否进行设备初始化
Public blIsCheck As Boolean    ' 是否进行设备检查
Public blIsPickSms As Boolean  ' 是否拉取短信
Public blIsATExecing As Boolean  ' 当前AT命令是否正在执行

Public iSmsId As Integer     ' 当前正在发送的短信ID
Public iBindId As Integer    ' 当前正在执行的绑定任务ID

Private strAtTemp  As String ' AT响应数据暂存区
Public Iccid As String       ' ICCID卡号
Public Imei As String        ' IMEI号
Public Imsi As String        ' IMSI号
Public Mobile As String      ' 手机号
Public bMobile As String     ' 呼转手机号


Public task As task          ' 将送往该COM口的命令任务队列
Public comPort As String     ' 串口编号

Private Sub Class_Initialize()
    blIsWaiting = False
    blIsOpen = False
    blIsCheck = True
    blIsPickSms = True
    blIsATExecing = False
    strAtTemp = ""
    Iccid = ""
    Imei = ""
    Imsi = ""
    Mobile = ""
    bMobile = ""
    iSmsId = 0
    iBindId = 0
    Set task = New task
End Sub
Public Function ReSet()
    blIsWaiting = False
    blIsOpen = False
    blIsCheck = True
    blIsPickSms = True
    blIsATExecing = False
    strAtTemp = ""
    Iccid = ""
    Imei = ""
    Imsi = ""
    Mobile = ""
    bMobile = ""
    iSmsId = 0
    iBindId = 0
    Set task = New task
End Function
Public Function GetData(ByRef strInput As String) As String
    Dim iLen As Integer
    If blIsWaiting = False Then
        If UCase(Left(strInput, 2)) = "AT" Then         '正常情况下的AT命令响应
            ' 判断是否为完整的AT命令响应数据
            If InStr(strInput, "OK") > 0 Or InStr(strInput, "ERROR") > 0 Then
                GetData = strInput
                strAtTemp = ""
                blIsWaiting = False
            Else
                strAtTemp = strInput
                blIsWaiting = True
            End If
        Else    '不是已AT开头的可能是串口通知 "+CMTI:"(新短信到达)
            GetData = strInput
            strAtTemp = ""
            blIsWaiting = False
        End If
    Else
        If InStr(strInput, "OK") > 0 Or InStr(strInput, "ERROR") > 0 Or (strAtTemp & strInput = "AT+CFUN=1,1" & vbCrLf & "IIII") Then
            GetData = strAtTemp & strInput
            strAtTemp = ""
            blIsWaiting = False
        Else
            strAtTemp = strAtTemp & strInput
            blIsWaiting = True
        End If
    End If
End Function


Public Function AnalysisData(strInput As String, strOut As String) As String
    Dim iLen As Integer
    Dim strAT As String
    If UCase(Left(strInput, 2)) = "AT" Then
        iLen = InStr(strInput, vbCrLf)
        If iLen > 0 Then
            strAT = UCase(Left(strInput, iLen - 1))
        End If
    ElseIf UCase(Left(strInput, 6)) = "+CMTI:" Then '收到新短信
        strAT = "+CMTI:"
    Else
        strAT = UCase(regExpGetAT(strInput))
    End If
    
    Select Case strAT
        Case "AT+CSQ"
            strOut = regExpGetFirst(":\s([^\s]*)", strInput)
        Case "AT+COPS?"
            strOut = regExpGetFirst(":\s([^\s]*)", strInput)
        Case "AT+CGSN"
            strOut = regExpGetFirst("(\d{15})", strInput)
        Case "AT+CIMI"
            strOut = regExpGetFirst("(\d{15})", strInput)
        Case "AT+CCID"
            strOut = regExpGetFirst("(\d{20})", strInput)
            If strOut = "" Then
                strOut = "-RETRY-"
            End If
        Case "AT+CCFC=0,2" 'AT+CCFC=0,2\r\n+CCFC: 1,1,"+8618650346179",145,,,\r\nOK
            strOut = regExpGetFirst("""(?:\+86)?(\d{11})""", strInput)
        Case "AT+CCFC=0,4"
            If InStr(strInput, "OK") Then
                strAT = "-AT-UNBIND-MOBILE-OK-"
            Else
                strAT = "-AT-UNBIND-MOBILE-FAILED-"
            End If
        Case "AT+CFUN=1,1"
            If InStr(strInput, "IIII") Then
                strAT = "-AT-INIT-OK-"
                strOut = ""
            End If
        Case Else
            If InStr(strAT, "AT+CMGL") > 0 Then
                strAT = "AT+CMGL"
                strOut = strInput
            ElseIf InStr(strAT, "AT+CMGS=") > 0 Then
                If InStr(strInput, "+CMGS:") > 0 Then
                    strAT = "-AT-SMS-SEND-OK-"  ' 短信发送成功
                    strOut = strInput
                Else
                    strAT = "-AT-SMS-SEND-FAILED-"  ' 短信发送成功
                    strOut = strInput

                End If
            ElseIf InStr(strAT, "AT+CCFC=0,3") > 0 Then
                If InStr(strInput, "OK") Then
                    strAT = "-AT-BIND-MOBILE-OK-"
                Else
                    strAT = "-AT-BIND-MOBILE-FAILED-"
                End If
            Else
                strOut = ""
            End If
    End Select
    AnalysisData = strAT
End Function


Public Function sendSMS(strTo, strCnt, Optional blIsEN As Boolean)
    strTo = GB2Unicode(strTo)
    strCnt = GB2Unicode(strCnt)
    task.Push ("AT+CMGF=1" & vbCrLf)          ' 短信格式选择 0:PDU 1:文本
    task.Push ("AT+CSCS=""UCS2""" & vbCrLf)   ' 字符集选择
    task.Push ("AT+CSCA?" & vbCrLf)           ' 自动设置短消息中心号码，只做查询即可
    task.Push ("AT+CNMI=2,1" & vbCrLf)        ' 新消息指示格式 "+CMTI:”SM”,3"
    If blIsEN <> True Then
        task.Push ("AT+CSMP=17,167,0,8" & vbCrLf) ' 发送中文短信
    Else
        task.Push ("AT+CSMP=17,167,0,0" & vbCrLf) ' 发送英文短信
    End If
    task.Push ("AT+CMGS=""" & strTo & """" & vbCrLf)
    task.Push (strCnt)
    task.Push ("--TAIL--")  ' 结束符 CTRL+Z
End Function

Public Function delSMS(iSmsIndex As Integer, Optional blDelAll As Boolean)
    If blDelAll = True Then
        task.Push ("AT+CMGD=0,2" & vbCrLf)
    Else
        If IsNumeric(iSmsIndex) And 1 <= iSmsIndex And iSmsIndex <= 50 Then
            task.Push ("AT+CMGD=" & iSmsIndex & vbCrLf)  ' 删除指定位置短信
        End If
    End If
End Function

Public Function bindMobile(strMobile As String)
    task.Push ("AT+CCFC=0,3,""" & strMobile & """,129,1" & vbCrLf)
End Function

Public Function unBindMobile()
    task.Push ("AT+CCFC=0,4" & vbCrLf)
End Function

Public Function regExpGetAT(ByRef strInput As String) As String
   Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
   Dim Matches   As MatchCollection     '声明Matches为匹配结果集合
   
   Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
   objRegExp.Pattern = "(AT[^\s]*)"            '设置该正则对象的正则表达式
   objRegExp.IgnoreCase = False              '是否区分大小写
   objRegExp.Global = True                  '是否全局匹配
   
   If (objRegExp.Test(strInput) = True) Then '测试是否能匹配到我们需要的字符串
        Set Matches = objRegExp.Execute(strInput)   '开始检索
        regExpGetAT = Matches(0).SubMatches(0)
   Else
        regExpGetAT = ""
   End If
End Function

Public Function regExpGetFirst(strPat As String, strInput As String) As String
   Dim objRegExp As RegExp                 '声明objRegExp为一个正则对象
   Dim Matches   As MatchCollection     '声明Matches为匹配结果集合
   
   Set objRegExp = New RegExp               '初始化一个新的正则对象objRegExp
   objRegExp.Pattern = strPat            '设置该正则对象的正则表达式
   objRegExp.IgnoreCase = False              '是否区分大小写
   objRegExp.Global = True                  '是否全局匹配
   
   If (objRegExp.Test(strInput) = True) Then '测试是否能匹配到我们需要的字符串
        Set Matches = objRegExp.Execute(strInput)   '开始检索
        regExpGetFirst = Matches(0).SubMatches(0)
   Else
        regExpGetFirst = ""
   End If
End Function


