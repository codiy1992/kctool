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

' ���³�Ա��COM�����
Public blIsOpen As Boolean       ' �����Ƿ��
Private blIsWaiting As Boolean   ' �Ƿ���Ҫ�ȴ���һ֡����
Public blCanSwitch As Boolean    ' �Ƿ������п�
Public blIsATExecing As Boolean  ' ��ǰAT�����Ƿ�����ִ��
Public blIsSwitching As Boolean  ' �Ƿ������п�

Public iTmrCnt As Integer    ' cron��ʱ����������
Public rowIndex As Integer   ' ��λ��
Public index As Integer      ' ���ر��
Public comPort As Integer    ' ��ǰ���ڱ��
Public portErr As Integer    ' ���ڴ�����
Public iWaitCnt As Integer   ' �ȴ����ݴ���
Public iNullCnt As Integer   ' ��sim������

Public Imei As String        ' IMEI��
Private strAtTemp  As String ' AT��Ӧ�����ݴ���
Public task As task          ' ��������COM�ڵ������������

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
    Hw = 0 'У��λ
    Sw = 0 '������
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
    
    timeout = 1000 / sio_getbaud(comPort)  'ms /byte ��ȡ���ڲ�����
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
        ' �ж��Ƿ�Ϊ������AT������Ӧ����
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
   Dim objRegExp As RegExp                 '����objRegExpΪһ���������
   Dim Matches   As MatchCollection     '����MatchesΪƥ��������
   
   Set objRegExp = New RegExp               '��ʼ��һ���µ��������objRegExp
   objRegExp.Pattern = strPat            '���ø����������������ʽ
   objRegExp.IgnoreCase = False              '�Ƿ����ִ�Сд
   objRegExp.Global = True                  '�Ƿ�ȫ��ƥ��
   
   If (objRegExp.Test(strInput) = True) Then '�����Ƿ���ƥ�䵽������Ҫ���ַ���
        Set Matches = objRegExp.Execute(strInput)   '��ʼ����
        regExpGetFirst = Matches(0).SubMatches(0)
   Else
        regExpGetFirst = ""
   End If
End Function

Private Function Str2Hex(ByVal strInput As String, HexArr() As Byte)
    Dim temp As String                                                            '  ���������ݴ�
    Dim dataCount As Long                                                         '  ���ݸ�������
    Dim I As Long                                                                 '  �ֲ�����
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
    ReDim HexArr(dataCount / 2 - 1)                                                 '  ���¶������鳤��
    For I = 0 To dataCount / 2 - 1
        HexArr(I) = Val("&H" + Mid(temp, I * 2 + 1, 2))                          '  ȡ������ת����ʮ�����Ʋ�
    Next
    Str2Hex = dataCount / 2                                          '  �����ܷ�����
End Function
