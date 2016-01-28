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

' 以下成员与COM口相关
Public blIsOpen As Boolean       ' 串口是否打开
Private blIsWaiting As Boolean   ' 是否需要等待下一帧数据
Public blIsATExecing As Boolean  ' 当前AT命令是否正在执行


Public comPort As Integer    ' 当前串口编号
Public portErr As Integer    ' 串口错误码
Public iWaitCnt As Integer   ' 等待数据次数
Private iTryCnt As Integer   ' 重复尝试次数 查ICCID

Public Imei As String        ' IMEI号
Private strAtTemp  As String ' AT响应数据暂存区
Public task As task          ' 将送往该COM口的命令任务队列

' 以下成员与SIM卡相关
Public blIsPickSms As Boolean  ' 是否拉取短信

Public iSmsId As Integer     ' 当前正在发送的短信ID
Public iBindId As Integer    ' 当前正在执行的绑定任务ID
Public iQccfcCnt As Integer  ' 查询呼转设置

Public Iccid As String       ' ICCID卡号
Public Imsi As String        ' IMSI号
Public Mobile As String      ' 手机号
Public bMobile As String     ' 呼转手机号


Private Sub Class_Initialize()
    blIsOpen = False
    blIsWaiting = False
    blIsATExecing = False
    Imei = ""
    strAtTemp = ""
    Set task = New task
    
    blIsPickSms = True
    Iccid = ""
    Imsi = ""
    Mobile = ""
    bMobile = ""
End Sub
Public Function ReSet()
    ClosePort
    task.Clean
    blIsOpen = False
    blIsWaiting = False
    blIsATExecing = False
    iTryCnt = 0
    iWaitCnt = 0
    Imei = ""
    strAtTemp = ""
    
    blIsPickSms = True
    Iccid = ""
    Imsi = ""
    Mobile = ""
    bMobile = ""
    iSmsId = 0
    iBindId = 0
    iQccfcCnt = 0
End Function
Public Function OpenPort()
    Dim ret As Long
    Dim mode As Long
    Dim Hw  As Long, Sw As Long
    Dim timeout As Long
    
    mode = P_NONE Or BIT_8 Or STOP_1 'n,8,1
    Hw = 0 '校验位
    Sw = 0 '流控制
    blIsOpen = False
    If comPort <= 1 Then
        Exit Function
    End If
    ret = sio_open(comPort)
    If ret <> SIO_OK Then
        portErr = ret
        sio_close (comPort)
        Exit Function
    End If
    
    ret = sio_ioctl(comPort, B115200, mode)
    If ret <> SIO_OK Then
        sio_close (comPort)
        Exit Function
    End If

    ret = sio_flowctrl(comPort, Hw Or Sw)
    If ret <> SIO_OK Then
        sio_close (comPort)
        Exit Function
    End If

    ret = sio_DTR(comPort, 1)
    If ret <> SIO_OK Then
        sio_close (comPort)
        Exit Function
    End If

    ret = sio_RTS(comPort, 1)
    If ret <> SIO_OK Then
        sio_close (comPort)
        Exit Function
    End If
    
    timeout = 1000 / sio_getbaud(comPort)  'ms /byte 获取串口波特率
    If timeout < 1 Then
        timeout = 1
    End If
    timeout = timeout * 1 * 3 ' 1 byte; '*3' is for delay
    If timeout < 100 Then
        timeout = 100
    End If
    Call sio_SetWriteTimeouts(comPort, timeout)
    blIsOpen = True
End Function

Public Function ClosePort()
    sio_close (comPort)
    blIsOpen = False
End Function

Public Function WriteData(outData)
    Dim timeout As Long
    Dim iLen As Long
    Dim HexArr() As Byte
    iWaitCnt = 0
    If LCase(TypeName(outData)) = "string" Then
        iLen = Str2Hex(outData, HexArr())
    ElseIf LCase(TypeName(outData)) = "byte()" Then
        HexArr() = outData
        iLen = UBound(HexArr) + 1
    End If
    If iLen = 0 Then
        Exit Function
    End If
     Call sio_GetWriteTimeouts(comPort, 0)
    timeout = 1000 / sio_getbaud(comPort)  'ms /byte'
    If timeout < 1 Then
        timeout = 1
    End If
    timeout = timeout * iLen * 2             ' 256 byte;*2 for delay
    Call sio_SetWriteTimeouts(comPort, timeout)
    Call sio_write(comPort, HexArr(0), iLen)
    Call sio_SetWriteTimeouts(comPort, 0)
End Function

Public Function ReadData() As String
    Dim I As Long
    Dim rlen As Long
    Dim blEmpty As Boolean
    Dim buf(0 To 511) As Byte
    Dim Top As Long
    blEmpty = True
    rlen = sio_read(comPort, buf(0), 512)
    For I = 0 To 511
        If buf(I) <> 0 Then
            blEmpty = False
            Exit For
        End If
    Next I
    If blEmpty = False Then
        For I = 0 To 511
            If buf(511 - I) <> 0 Then
               Top = 511 - I
               Exit For
            End If
        Next I
        For I = 0 To Top
            ReadData = ReadData & Chr(Val(buf(I)))
        Next I
    Else
        ReadData = ""
        iWaitCnt = iWaitCnt + 1
    End If
End Function

Public Function GetData(ByRef strInput As String) As String
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
        'MsgBox strAtTemp & strInput
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
        strAT = UCase(regExpGetFirst("(AT[^\r]*)", strInput))
    End If
    
    Select Case strAT
        Case "AT+CSQ"
            strOut = regExpGetFirst(":\s([^\r]*)", strInput)
        Case "AT+COPS?"
            strOut = regExpGetFirst("""([^""]*)""", strInput)
        Case "AT+CSCA?"
            strOut = regExpGetFirst("""([^""]*)""", strInput)
        Case "AT+CGSN"
            strOut = regExpGetFirst("(\d{15})", strInput)
        Case "AT+CIMI"
            strOut = regExpGetFirst("(\d{15})", strInput)
        Case "AT+CCID"
            strOut = regExpGetFirst("(\d{20})", strInput)
            If strOut = "" Then
                strOut = "-RETRY-"
            End If
            If InStr(strInput, "ERROR") > 0 Then
                iTryCnt = iTryCnt + 1
                If iTryCnt > 3 Then
                    strAT = "-AT-NO-CCID-"
                End If
            End If
        Case "AT+CCFC=0,2" 'AT+CCFC=0,2\r\n+CCFC: 1,1,"+8618650346179",145,,,\r\nOK
            strOut = regExpGetFirst("""(?:\+86)?(\d{11})""", strInput)
            If strOut = "" Then
                iQccfcCnt = iQccfcCnt + 1
            End If
        Case "AT+CCFC=0,4"
            If InStr(strInput, "OK") Then
                iQccfcCnt = 0
                strAT = "-AT-UNBIND-MOBILE-OK-"
            Else
                strAT = "-AT-UNBIND-MOBILE-FAILED-"
            End If
        Case "AT+CFUN=1,1"
            If InStr(strInput, "IIII") Or InStr(strInput, "OK") Then
                strAT = "-AT-INIT-OK-"
                strOut = ""
            End If
        Case "AT+CFUN=0"
            If InStr(strInput, "OK") Then
                strAT = "-AT-EXIT-OK-"
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
                    iQccfcCnt = 0
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

Private Function regExpGetFirst(strPat As String, strInput As String) As String
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

Private Function Str2Hex(ByVal strInput As String, HexArr() As Byte)
    Dim temp As String                                                            '  发送数据暂存
    Dim dataCount As Long                                                         '  数据个数计数
    Dim I As Long                                                                 '  局部变量
    Dim abytS() As Byte
    Dim bytTemp As Byte
    Dim strTemp As String
    abytS = StrConv(strInput, vbFromUnicode)
    For I = 0 To UBound(abytS)
        bytTemp = abytS(I)
        strTemp = Hex(bytTemp)
        strTemp = Right("00" & strTemp, 2)
        temp = temp & strTemp
    Next I
    dataCount = Len(temp)
    ReDim HexArr(dataCount / 2 - 1)                                                 '  重新定义数组长度
    For I = 0 To dataCount / 2 - 1
        HexArr(I) = Val("&H" + Mid(temp, I * 2 + 1, 2))                          '  取出数据转换成十六进制并
    Next
    Str2Hex = dataCount / 2                                          '  计算总发送数
End Function
