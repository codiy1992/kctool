Attribute VB_Name = "mdlComFuns"
Option Explicit
'**********************************************************************
' 获取所有在运行的iccid
'**********************************************************************
Public Function GetAllIccid(Optional blIsSkipSendingSMS As Boolean) As String
    Dim cIdx As Integer
    GetAllIccid = "''"
    If IsComEmpty = False Then
        If blIsSkipSendingSMS = True Then
            For cIdx = 0 To UBound(com())
                If com(cIdx).Iccid <> "" And com(cIdx).iSmsId = 0 Then
                    GetAllIccid = GetAllIccid & " , '" & com(cIdx).Iccid & "'"
                End If
            Next cIdx
        Else
            For cIdx = 0 To UBound(com())
                If com(cIdx).Iccid <> "" Then
                    GetAllIccid = GetAllIccid & " , '" & com(cIdx).Iccid & "'"
                End If
            Next cIdx
        End If
    End If
End Function

Public Function IsComEmpty() As Boolean
    On Error GoTo Err
    If UBound(com()) > -1 Then
        IsComEmpty = False
        Exit Function
    End If
Err:
    IsComEmpty = True
End Function

Public Function ComErr(iErrCode As Integer) As String
    Select Case iErrCode
        Case SIO_BADPORT
            ComErr = "端口未打开"
        Case SIO_OUTCONTROL
            ComErr = "OUT CONTROL"
        Case SIO_NODATA
            ComErr = "无数据/缓冲区"
        Case SIO_OPENFAIL
            ComErr = "端口被占用"
        Case SIO_RTS_BY_HW
            ComErr = "代码:-6"
        Case SIO_BADPARM
            ComErr = "参数错误"
        Case SIO_WIN32FAIL
            ComErr = "调用WIN32失败"
        Case SIO_BOARDNOTSUPPORT
            ComErr = "代码:-9"
        Case SIO_FAIL
            ComErr = "代码:-10"
        Case SIO_ABORT_WRITE
            ComErr = "代码:-11"
        Case SIO_WRITETIMEOUT
            ComErr = "写超时"
        Case Else
            ComErr = "未知错误"
    End Select
End Function

'#########################################################
'功能： 生成短信PDU串
'输入： 目标手机号码、短信息内容、[可选的短信服务中心号码]
'输出： 生成的PDU串
'返回： 整个字串的长度
'#########################################################
Public Function SmsPDU(ByVal DestNo As String, _
                        ByVal SMSText As String, _
                        ByRef PDUString As String, _
                        Optional ByVal ServiceNo As String) As Long
    On Error GoTo ErrorPDU
    Dim I As Integer
    Dim iAsc As Integer
    Dim iLen As Integer
    Dim strTmp As String
    Dim strTmp2 As String
    Dim strChar As String
    
    If SMSText = "" Then Exit Function
    
    ' 对消息中心号码进行编码
    If ServiceNo = "" Then
        ServiceNo = "00"
    Else
        If Left(ServiceNo, 3) = "+86" Then
            ServiceNo = Mid(ServiceNo, 4)
        End If
        For I = 1 To Len(ServiceNo)
            strChar = Mid(ServiceNo, I, 1)
            iAsc = Asc(strChar)
            If iAsc > 57 Or iAsc < 48 Then Exit Function
        Next I
        If Len(ServiceNo) Mod 2 = 1 Then
            ServiceNo = ServiceNo & "F"
        End If
        For I = 1 To 12 Step 2
            strTmp2 = Mid(ServiceNo, I, 2)
            strTmp = strTmp & Right(strTmp2, 1) & Left(strTmp2, 1)
        Next I
        ServiceNo = "089168" & strTmp
    End If
    
    ' 对目标号码进行编码 0D9168
    strTmp2 = ""
    strTmp = ""
    If Left(DestNo, 3) = "+86" Then
        DestNo = Mid(DestNo, 4)
    End If
    
    For I = 1 To Len(DestNo)
        strChar = Mid(DestNo, I, 1)
        iAsc = Asc(strChar)
        If iAsc > 57 Or iAsc < 48 Then Exit Function
    Next I
    
    If Len(DestNo) Mod 2 = 1 Then
        DestNo = DestNo & "F"
    End If

    For I = 1 To Len(DestNo) Step 2
        strTmp2 = Mid(DestNo, I, 2)
        strTmp = strTmp & Right(strTmp2, 1) & Left(strTmp2, 1)
    Next I
    
    DestNo = "0" & Hex(Len(strTmp) + 1) & "9168" & strTmp

    ' 对内容进行编码
    SMSText = GB2Unicode(SMSText)
    iLen = Len(SMSText) \ 2
    strChar = Hex(iLen)
    If Len(strChar) < 2 Then strChar = "0" & strChar
    SMSText = strChar & SMSText
    
    
    SmsPDU = Len("1100") / 2 + Len(DestNo) / 2 + Len("0008AA") / 2 + Len(SMSText) / 2
    PDUString = ServiceNo & "1100" & DestNo & "0008AA" & SMSText
    Exit Function
ErrorPDU:
    SmsPDU = 0
    PDUString = ""
End Function

Public Function PickAllSMS(ByRef InputString As String, RetSMS() As SMSDef) As String

    Dim I As Integer, iTmp As Integer, iLen As Integer, iNext As Integer, iCr As Integer
    
    Dim n As Long
    
    Dim strTmp As String, strTmp1 As String, strTmp2 As String
    
    Dim btTmp() As Byte, btTmp2() As Byte
    
    Dim blRet As Boolean
    
On Error Resume Next
    
    strTmp = ""
    btTmp = InputString
    
    '======== 将短消息中的双引号去除 ========
    iTmp = 0
    For I = 0 To UBound(btTmp)
        strTmp1 = Chr(btTmp(I))
        If strTmp1 <> """" And btTmp(I) <> 0 And strTmp1 <> vbLf Then
            ReDim Preserve btTmp2(0 To iTmp + 1)
            btTmp2(iTmp) = btTmp(I)
            btTmp2(iTmp + 1) = 0
            iTmp = iTmp + 2
        End If
    Next I
    InputString = btTmp2
    
    n = 0
    I = 1
    Do
        iTmp = InStr(I, InputString, "+CMGL:")
        iCr = InStr(iTmp, InputString, vbCr)
        
        If iTmp > 0 Then
            If iCr - iTmp + 1 > 0 Then n = n + 1
        ElseIf iTmp = 0 Then
            Exit Do
        End If
        I = iTmp + 7
    Loop
    
    If n > 0 Then
        ReDim RetSMS(1 To n)
    Else
        ReDim RetSMS(0 To 0)
    End If
    
    '======== 逐条保存到数据库中 ========
    For I = 1 To n
        iTmp = InStr(InputString, "+CMGL:")
        iCr = InStr(InputString, vbCr)
        
        If iCr > 0 And iTmp > 0 Then
            InputString = Right(InputString, Len(InputString) - iTmp + 1)
            iTmp = InStr(InputString, "+CMGL:")
            iNext = InStr(iTmp + 7, InputString, "+CMGL:")
            
            If iNext > 0 Then
                strTmp = Mid(InputString, iTmp, iNext - iTmp)
                InputString = Right(InputString, Len(InputString) - iNext + 1)
            Else
                iCr = InStr(iTmp, InputString, vbCr)
                iCr = InStr(iCr + 1, InputString, vbCr)
                strTmp = Mid(InputString, iTmp, iCr - iTmp)
                InputString = Right(InputString, Len(InputString) - iCr + 1)
            End If
            blRet = PickOneSMS(strTmp, RetSMS(I), True)
            If blRet Then
'                RetSMS(i).SmsIndex = i
                On Error GoTo ErrorNode
ErrorNode:
            End If
        End If
    Next I

    PickAllSMS = "共有" & n & "条短信"
End Function

Public Function PickOneSMS(strInputData As String, RetSMS As SMSDef, ByVal blIsList As Boolean) As Boolean
    
    Dim blRetFunc       As Boolean
    
    Dim I As Integer, iLen As Integer, iCr As Integer
    Dim nD As Long, nRet As Long
    
    Dim strTmp As String, strTmp1 As String, strTmp2 As String, strTmp3 As String
    
    Dim MyStr()         As String
    Dim aryTmp()        As String
    Dim DateTime As String
On Error GoTo ErrorSave
    
'+CMGL: 24,"REC READ","+8613811055271",,"04/06/03,22:35:35+32"
'4F608D767D27776189C95427002C621177E590534E86002C665A5B89
    
    '======== 取出短信息头部 ========
    
    iCr = InStr(strInputData, vbCr)
    iLen = Len(strInputData)
    If iCr > 0 And iCr <= iLen Then
        strTmp2 = Left(strInputData, iCr - 1)
        strInputData = Right(strInputData, iLen - iCr)
    End If
    
    '======== 取出短信息内容 ========
    iCr = InStr(strInputData, vbCr)
    
    iLen = Len(strInputData)
    
    If iCr > 0 Then
        If iCr <= iLen Then
            strTmp3 = Left(strInputData, iCr - 1)
            strInputData = Right(strInputData, iLen - iCr)
        End If
    Else
        If iCr < iLen Then
            strTmp3 = strInputData
        End If
    End If
    
On Error GoTo ErrorDecode
    '======== 分解短消息，以逗号(,)作为分隔符 ========
    blRetFunc = False
    blRetFunc = String2Array(strTmp2, ",", nD, aryTmp, True)
    
ErrorDecode:
    'Set myFunc = Nothing
    
    If blRetFunc Then
        '======== 如果传过来的短消息格式是"CMGL" ========
        If blIsList Then
            ReDim MyStr(0 To nD - 1)
            For I = 0 To nD - 2
                MyStr(I) = aryTmp(I + 1)
            Next I
        '======== 否则，传送过来的消息格式是"CMGR"，这两者是有区别的。 ========
        Else
            ReDim MyStr(0 To nD - 1)
            For I = 0 To nD - 1
                MyStr(I) = aryTmp(I)
            Next I
        End If
        
        RetSMS.ListOrRead = blIsList
        
        '======== 短信在SIM卡中的位置 ========
        iLen = InStr(aryTmp(0), ":")
        If iLen > 0 Then
            strTmp = Trim(Right(aryTmp(0), Len(aryTmp(0)) - iLen))
            If IsNumeric(strTmp) Then
                RetSMS.SmsIndex = CLng(strTmp)
            End If
        End If
        
        '======== 如果对方的SIM号码前面有"+86"，则剔除掉 ========
        On Error Resume Next
        If Left(MyStr(1), 3) = "+86" Then
            MyStr(1) = Right(MyStr(1), Len(MyStr(1)) - 3)
        End If
        
        
        '======== 取出短信中的用户数据UD ========
        iCr = InStr(strTmp3, vbCr)
        If iCr > 0 Then
            strTmp3 = Left(strTmp3, iCr - 1)
        End If
        
        '======== 分别提取短消息的详细内容 ========
        If Left(MyStr(1), 2) = "00" Then
            RetSMS.SourceNo = Unicode2GB(MyStr(1))
        Else
            RetSMS.SourceNo = MyStr(1)
        End If
        RetSMS.SmsMain = Unicode2GB(strTmp3)
        
        '======== 如果时间中含有时区，则去除 ========
        DateTime = Format(MyStr(2), "YYYY-MM-DD") & " " & Format(MyStr(3), "HH:MM:SS")
        iLen = InStr(DateTime, "+")
        If iLen > 0 Then DateTime = Left(DateTime, iLen - 1)
        iLen = InStr(DateTime, "-")
        If iLen > 0 Then DateTime = Left(DateTime, iLen - 1)
        
        RetSMS.ReachDate = MyStr(2)
        RetSMS.ReachTime = MyStr(3)
        RetSMS.DateTime = DateTime
        
        If Err = 0 Then PickOneSMS = True
    End If
    
    Exit Function
    
ErrorSave:
    PickOneSMS = False
End Function

