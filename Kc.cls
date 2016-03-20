VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Kc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 以下成员与COM口相关
Public blIsOpen As Boolean       ' 串口是否打开
Private blIsWaiting As Boolean   ' 是否需要等待下一帧数据
Public blCanSwitch As Boolean    ' 是否允许切卡
Public blIsATExecing As Boolean  ' 当前AT命令是否正在执行
Public blIsSwitching As Boolean  ' 是否正在切卡

Public iTmrCnt As Integer    ' cron定时器触发计数
Public rowIndex As Integer   ' 卡位置
Public index As Integer      ' 卡池编号
Public comPort As Integer    ' 当前串口编号
Public portErr As Integer    ' 串口错误码
Public iWaitCnt As Integer   ' 等待数据次数
Public iNullCnt As Integer   ' 无sim卡个数

Public Imei As String        ' IMEI号
Private strAtTemp  As String ' AT响应数据暂存区
Public task As task          ' 将送往该COM口的命令任务队列

Private Sub Class_Initialize()
    blIsOpen = False
    blIsWaiting = False
    blIsATExecing = False
    blCanSwitch = False
    blIsSwitching = False
    rowIndex = 1
    Imei = ""
    strAtTemp = ""
    Set task = New task
End Sub

Public Function ReSet()
    ClosePort
    task.Clean
    rowIndex = 1
    blIsOpen = False
    blIsWaiting = False
    blCanSwitch = False
    blIsATExecing = False
    blIsSwitching = False
    iWaitCnt = 0
    Imei = ""
    strAtTemp = ""
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
    timeout = 1000 / sio_getbaud(comPort)  ' ms /byte'
    If timeout < 1 Then
        timeout = 1
    End If
    timeout = timeout * iLen * 2 ' 256 byte;*2 for delay
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
        ' 判断是否为完整的AT命令响应数据
        If InStr(strInput, "OK") > 0 Then
            GetData = strInput
            strAtTemp = ""
            blIsWaiting = False
        Else
            strAtTemp = strInput
            blIsWaiting = True
        End If
    Else
        'MsgBox strAtTemp & strInput
        If InStr(strInput, "OK") > 0 Then
            GetData = strAtTemp & strInput
            strAtTemp = ""
            blIsWaiting = False
        Else
            strAtTemp = strAtTemp & strInput
            blIsWaiting = True
        End If
    End If
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

