VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SerialPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************
'**模 块 名：SerialPort
'**说    明：YFsoft 版权所有2006 - 2007(C)
'**创 建 人：叶帆
'**日    期：2006-08-17 14:32:29
'**修 改 人：
'**日    期：
'**描    述：串口异步读写(API)
'**版    本：V1.0.0
'*************************************************************************

Private Type ComStat
    fCtsHold As Long
    fDsrHold As Long
    fRlsdHold As Long
    fXoffHold As Long
    fXoffSent As Long
    fEof As Long
    fTxim As Long
    fReserved As Long
    cbInQue As Long
    cbOutQue As Long
End Type

Private Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

Private Type DCB
    DCBlength As Long
    BaudRate As Long
    'DWORD DCBlength;      /* sizeof(DCB)                     */
    'DWORD BaudRate;       /* Baudrate at which running       */
    'DWORD fBinary: 1;     /* Binary Mode (skip EOF check)    */
    'DWORD fParity: 1;     /* Enable parity checking          */
    'DWORD fOutxCtsFlow:1; /* CTS handshaking on output       */
    'DWORD fOutxDsrFlow:1; /* DSR handshaking on output       */
    'DWORD fDtrControl:2;  /* DTR Flow control                */
    'DWORD fDsrSensitivity:1; /* DSR Sensitivity              */
    'DWORD fTXContinueOnXoff: 1; /* Continue TX when Xoff sent */
    'DWORD fOutX: 1;       /* Enable output X-ON/X-OFF        */
    'DWORD fInX: 1;        /* Enable input X-ON/X-OFF         */
    'DWORD fErrorChar: 1;  /* Enable Err Replacement          */
    'DWORD fNull: 1;       /* Enable Null stripping           */
    'DWORD fRtsControl:2;  /* Rts Flow control                */
    'DWORD fAbortOnError:1; /* Abort all reads and writes on Error */
    'DWORD fDummy2:17;      /* Reserved                        */
    fBitFields As Long 'See Comments in Win32API.Txt
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer 'Reserved; Do Not Use
End Type

Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long 'OVERLAPPED
Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function BuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Private Declare Function SetCommState Lib "kernel32" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Private Declare Function GetCommState Lib "kernel32" (ByVal nCid As Long, lpDCB As DCB) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function SetCommMask Lib "kernel32" (ByVal hFile As Long, ByVal dwEvtMask As Long) As Long
Private Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, lpStat As ComStat) As Long
Private Declare Function GetOverlappedResult Lib "kernel32" (ByVal hFile As Long, lpOverlapped As OVERLAPPED, lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function SetupComm Lib "kernel32" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const DTR_CONTROL_DISABLE = &H0
Private Const RTS_CONTROL_ENABLE = &H1
Private Const PURGE_RXABORT = &H2
Private Const PURGE_RXCLEAR = &H8
Private Const PURGE_TXABORT = &H1
Private Const PURGE_TXCLEAR = &H4
Private Const ERROR_IO_PENDING = 997
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = (STATUS_WAIT_0 + 0)
Private Const WAIT_TIMEOUT = 258&

Private m_Handle As Long
Private m_OverlappedRead As OVERLAPPED
Private m_OverlappedWrite As OVERLAPPED

'以上是API声明

'*************************************************************************
'**函 数 名：OpenPort
'**输    入：ComNumber(Long)     - 串口号
'**        ：Comsettings(String) - 配置信息
'**输    出：(Long) - 0 成功 非 0 失败
'**功能描述：打开串口
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2006-08-17 14:40:14
'**修 改 人：
'**日    期：
'**版    本：V1.0.0
'*************************************************************************
Public Function OpenPort(ComNumber As Long, Comsettings As String, Optional lngInSize As Long = 1024, Optional lngOutSize As Long = 512) As Long
    On Error GoTo handelinitcom
    Dim retval As Long
    Dim CtimeOut As COMMTIMEOUTS, dcbs As DCB
    Dim strCOM As String, strConfig As String

    If ComNumber > 9 Then
        strCOM = "\\.\COM" & Format(ComNumber, "00")
    Else
        strCOM = "COM" & Format(ComNumber, "0")
    End If
    m_Handle = CreateFile(strCOM, GENERIC_READ Or GENERIC_WRITE, 0, 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_OVERLAPPED, 0)
'    m_Handle = CreateFile("COM1", &HC0000000, 0, 0&, &H3, 0, 0)
    If m_Handle = -1 Then
        OpenPort = -1
        Exit Function
    End If

    '设置dcb块
    dcbs.DCBlength = Len(dcbs)                           '长度
    Call GetCommState(m_Handle, dcbs)

    '波特率,奇偶校验,数据位,停止位  如：9600,n,8,1
    strConfig = "COM" & Format(ComNumber, "0") & ":" & Comsettings
    Call BuildCommDCB(strConfig, dcbs)

    '------------------------------
    '    dcbs.fBinary = 1                          '二进制方式
    '    dcbs.fOutxCtsFlow = 0                     '不用CTS检测发送流控制
    '    dcbs.fOutxDsrFlow = 0                     '不用DSR检测发送流控制
    '    dcbs.fDtrControl = DTR_CONTROL_DISABLE    '禁止DTR流量控制
    '    dcbs.fDsrSensitivity = 0                  '对DTR信号线不敏感
    '    dcbs.fTXContinueOnXoff = 1                '检测接收缓冲区
    '    dcbs.fOutX = 0                            '不做发送字符控制
    '    dcbs.fInX = 0                             '不做接收控制
    '    dcbs.fErrorChar = 0                       '是否用指定字符替换校验错的字符
    '    dcbs.fNull = 0                            '保留NULL字符
    '    dcbs.fRtsControl = RTS_CONTROL_ENABLE     '允许RTS流量控制
    '    dcbs.fAbortOnError = 0                    '发送错误后，继续进行下面的读写操作
    '    dcbs.fDummy2 = 0                          '保留
    dcbs.fBitFields = 1 * 2 ^ 0 Or DTR_CONTROL_DISABLE * 2 ^ 4 Or 1 * 2 ^ 7 Or RTS_CONTROL_ENABLE * 2 ^ 12

    dcbs.wReserved = 0                        '没有使用，必须为0
    dcbs.XonLim = 0                           '指定在XOFF字符发送之前接收到缓冲区中可允许的最小字节数
    dcbs.XoffLim = 0                          '指定在XOFF字符发送之前缓冲区中可允许的最小可用字节数
    dcbs.XonChar = 0                          '发送和接收的XON字符
    dcbs.XoffChar = 0                         '发送和接收的XOFF字符
    dcbs.ErrorChar = 0                        '代替接收到奇偶校验错误的字符
    dcbs.EofChar = 0                          '用来表示数据的结束
    dcbs.EvtChar = 0                          '事件字符，接收到此字符时，会产生一个事件
    'dcbs.wReserved1 = 0                      '没有使用
    'dcbs.BaudRate =9600                      '波特率
    'dcbs.Parity=0                            '奇偶校验
    'dcbs.ByteSize=8                          '数据位
    'dcbs.StopBits=0                          '停止位
    '------------------------------

    If dcbs.Parity = 0 Then                   ' 0-4=None,Odd,Even,Mark,Space
        dcbs.fBitFields = dcbs.fBitFields And &HFFFD     'dcbs.fParity = 0                      '奇偶校验无效
    Else
        dcbs.fBitFields = dcbs.fBitFields Or &H2         'dcbs.fParity = 1                      '奇偶校验有效
    End If

    '超时设置
    CtimeOut.ReadIntervalTimeout = 20                  '0
    CtimeOut.ReadTotalTimeoutConstant = 1              '2500
    CtimeOut.ReadTotalTimeoutMultiplier = 1            '0
    CtimeOut.WriteTotalTimeoutConstant = 10            '2500
    CtimeOut.WriteTotalTimeoutMultiplier = 1           '0
    
    retval = SetCommTimeouts(m_Handle, CtimeOut)

    If retval = -1 Then
        retval = GetLastError()
        OpenPort = retval
        retval = CloseHandle(m_Handle)
        Exit Function
    End If

    '获取信号句柄
    Dim lpEventAttributes1 As SECURITY_ATTRIBUTES
    Dim lpEventAttributes2 As SECURITY_ATTRIBUTES

    m_OverlappedRead.hEvent = CreateEvent(lpEventAttributes1, 1, 0, 0)
    m_OverlappedWrite.hEvent = CreateEvent(lpEventAttributes2, 1, 0, 0)

    '判断设置参数是否成功   设置输入和输出缓冲区是否成功
    If SetCommState(m_Handle, dcbs) = -1 Or SetupComm(m_Handle, lngInSize, lngOutSize) = -1 Or m_OverlappedRead.hEvent = 0 Or m_OverlappedWrite.hEvent = 0 Then
        retval = GetLastError()
        OpenPort = retval
        If (m_OverlappedRead.hEvent <> 0) Then CloseHandle (m_OverlappedRead.hEvent)
        If (m_OverlappedWrite.hEvent <> 0) Then CloseHandle (m_OverlappedWrite.hEvent)
        Call CloseHandle(m_Handle)
        m_Handle = 0
        Exit Function
    End If

    OpenPort = 0
    Exit Function
handelinitcom:
    Call CloseHandle(m_Handle)
    m_Handle = 0
    OpenPort = -2
    Exit Function
End Function

'*************************************************************************
'**函 数 名：ClosePort
'**输    入：无
'**输    出：(Long) - 0 成功 -1 失败
'**功能描述：关闭串口
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2006-08-17 14:56:13
'**修 改 人：
'**日    期：
'**版    本：V1.0.0
'*************************************************************************
Public Function ClosePort() As Long
    If (m_Handle = 0) Then
        ClosePort = 1
        Exit Function
    End If

    Call SetCommMask(m_Handle, 0)
    Call SetEvent(m_OverlappedRead.hEvent)
    Call SetEvent(m_OverlappedWrite.hEvent)

    If (m_OverlappedRead.hEvent <> 0) Then CloseHandle (m_OverlappedRead.hEvent)
    If (m_OverlappedWrite.hEvent <> 0) Then CloseHandle (m_OverlappedWrite.hEvent)

    If CloseHandle(m_Handle) <> 0 Then
        ClosePort = 0
    Else
        ClosePort = -1
    End If

    m_Handle = 0
End Function

'*************************************************************************
'**函 数 名：ClearInBuf
'**输    入：无
'**输    出：无
'**功能描述：清空输入缓冲区
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2006-08-17 14:57:26
'**修 改 人：
'**日    期：
'**版    本：V1.0.0
'*************************************************************************
Public Function ClearInBuf() As Long
    If (m_Handle = 0) Then
        ClearInBuf = 1
        Exit Function
    End If
    Call PurgeComm(m_Handle, PURGE_RXABORT Or PURGE_RXCLEAR)
    ClearInBuf = 0
End Function

'*************************************************************************
'**函 数 名：ClearOutBuf
'**输    入：无
'**输    出：(Long) -
'**功能描述：清空输出缓冲区
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2006-08-17 15:40:38
'**修 改 人：
'**日    期：
'**版    本：V1.0.0
'*************************************************************************
Public Function ClearOutBuf() As Long
    If (m_Handle = 0) Then
        ClearOutBuf = 1
        Exit Function
    End If
    Call PurgeComm(m_Handle, PURGE_TXABORT Or PURGE_TXCLEAR)
    ClearOutBuf = 0
End Function

'*************************************************************************
'**函 数 名：SendData
'**输    入：bytBuffer()(Byte) - 数据
'**        ：lngSize(Long)     - 数据长度
'**输    出：(Long) -
'**功能描述：发送数据
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2006-08-17 15:43:42
'**修 改 人：
'**日    期：
'**版    本：V1.0.0
'*************************************************************************
Public Function SendData(bytBuffer() As Byte, lngSize As Long) As Long
    On Error GoTo ToExit '打开错误陷阱
    '------------------------------------------------
    If (m_Handle = 0) Then
        SendData = 1
        Exit Function
    End If

    Dim dwBytesWritten As Long
    Dim bWriteStat As Long
    Dim ComStats As ComStat
    Dim dwErrorFlags As Long

    dwBytesWritten = lngSize

    Call ClearCommError(m_Handle, dwErrorFlags, ComStats)
    bWriteStat = WriteFile(m_Handle, bytBuffer(0), lngSize, dwBytesWritten, m_OverlappedWrite)

    If bWriteStat = 0 Then
        If GetLastError() = ERROR_IO_PENDING Then
            Call GetOverlappedResult(m_Handle, m_OverlappedWrite, dwBytesWritten, 1)                   '等待直到发送完毕
        End If
    Else
        dwBytesWritten = 0
    End If

    SendData = dwBytesWritten
    '------------------------------------------------
    Exit Function
    '----------------
ToExit:
    SendData = -1
End Function

'*************************************************************************
'**函 数 名：ReadData
'**输    入：bytBuffer()(Byte) - 数据
'**        ：lngSize(Long)     - 数据长度
'**输    出：(Long) -
'**功能描述：读取数据
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2006-08-17 16:04:38
'**修 改 人：
'**日    期：
'**版    本：V1.0.0
'*************************************************************************
Public Function ReadData(bytBuffer() As Byte, lngSize As Long, Optional Overtime As Long = 3000) As Long
    On Error GoTo ToExit '打开错误陷阱
    '------------------------------------------------
    If (m_Handle = 0) Then
        ReadData = 1
        Exit Function
    End If

    Dim lngBytesRead As Long
    Dim fReadStat As Long
    Dim dwRes  As Long

    lngBytesRead = lngSize

    '读数据
    fReadStat = ReadFile(m_Handle, bytBuffer(0), lngSize, lngBytesRead, m_OverlappedRead)
    If fReadStat = 0 Then
        If GetLastError() = ERROR_IO_PENDING Then                           '重叠 I/O 操作在进行中
            dwRes = WaitForSingleObject(m_OverlappedRead.hEvent, Overtime)  '等待，直到超时
            Select Case dwRes
            Case WAIT_OBJECT_0:   '读完成
                If GetOverlappedResult(m_Handle, m_OverlappedRead, lngBytesRead, 0) = 0 Then
                    '错误
                    ReadData = -2
                    Exit Function
                End If
            Case WAIT_TIMEOUT:    '超时
                ReadData = -1
                Exit Function
            Case Else:                  'WaitForSingleObject 错误
            End Select
        End If
    End If
    ReadData = lngBytesRead
    '------------------------------------------------
    Exit Function
    '----------------
ToExit:
    ReadData = -1
End Function

'*************************************************************************
'**函 数 名：Class_Terminate
'**输    入：无
'**输    出：无
'**功能描述：
'**全局变量：
'**调用模块：
'**作    者：叶帆
'**日    期：2006-08-17 16:36:21
'**修 改 人：
'**日    期：
'**版    本：V1.0.0
'*************************************************************************
Private Sub Class_Terminate()
    Call ClosePort
End Sub

