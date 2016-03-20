Attribute VB_Name = "mdlComFuns"
Option Explicit
'**********************************************************************
' ��ȡ���������е�iccid
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
            ComErr = "�˿�δ��"
        Case SIO_OUTCONTROL
            ComErr = "OUT CONTROL"
        Case SIO_NODATA
            ComErr = "������/������"
        Case SIO_OPENFAIL
            ComErr = "�˿ڱ�ռ��"
        Case SIO_RTS_BY_HW
            ComErr = "����:-6"
        Case SIO_BADPARM
            ComErr = "��������"
        Case SIO_WIN32FAIL
            ComErr = "����WIN32ʧ��"
        Case SIO_BOARDNOTSUPPORT
            ComErr = "����:-9"
        Case SIO_FAIL
            ComErr = "����:-10"
        Case SIO_ABORT_WRITE
            ComErr = "����:-11"
        Case SIO_WRITETIMEOUT
            ComErr = "д��ʱ"
        Case Else
            ComErr = "δ֪����"
    End Select
End Function

'#########################################################
'���ܣ� ���ɶ���PDU��
'���룺 Ŀ���ֻ����롢����Ϣ���ݡ�[��ѡ�Ķ��ŷ������ĺ���]
'����� ���ɵ�PDU��
'���أ� �����ִ��ĳ���
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
    
    ' ����Ϣ���ĺ�����б���
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
    
    ' ��Ŀ�������б��� 0D9168
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

    ' �����ݽ��б���
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
    
    '======== ������Ϣ�е�˫����ȥ�� ========
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
    
    '======== �������浽���ݿ��� ========
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

    PickAllSMS = "����" & n & "������"
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
    
    '======== ȡ������Ϣͷ�� ========
    
    iCr = InStr(strInputData, vbCr)
    iLen = Len(strInputData)
    If iCr > 0 And iCr <= iLen Then
        strTmp2 = Left(strInputData, iCr - 1)
        strInputData = Right(strInputData, iLen - iCr)
    End If
    
    '======== ȡ������Ϣ���� ========
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
    '======== �ֽ����Ϣ���Զ���(,)��Ϊ�ָ��� ========
    blRetFunc = False
    blRetFunc = String2Array(strTmp2, ",", nD, aryTmp, True)
    
ErrorDecode:
    'Set myFunc = Nothing
    
    If blRetFunc Then
        '======== ����������Ķ���Ϣ��ʽ��"CMGL" ========
        If blIsList Then
            ReDim MyStr(0 To nD - 1)
            For I = 0 To nD - 2
                MyStr(I) = aryTmp(I + 1)
            Next I
        '======== ���򣬴��͹�������Ϣ��ʽ��"CMGR"����������������ġ� ========
        Else
            ReDim MyStr(0 To nD - 1)
            For I = 0 To nD - 1
                MyStr(I) = aryTmp(I)
            Next I
        End If
        
        RetSMS.ListOrRead = blIsList
        
        '======== ������SIM���е�λ�� ========
        iLen = InStr(aryTmp(0), ":")
        If iLen > 0 Then
            strTmp = Trim(Right(aryTmp(0), Len(aryTmp(0)) - iLen))
            If IsNumeric(strTmp) Then
                RetSMS.SmsIndex = CLng(strTmp)
            End If
        End If
        
        '======== ����Է���SIM����ǰ����"+86"�����޳��� ========
        On Error Resume Next
        If Left(MyStr(1), 3) = "+86" Then
            MyStr(1) = Right(MyStr(1), Len(MyStr(1)) - 3)
        End If
        
        
        '======== ȡ�������е��û�����UD ========
        iCr = InStr(strTmp3, vbCr)
        If iCr > 0 Then
            strTmp3 = Left(strTmp3, iCr - 1)
        End If
        
        '======== �ֱ���ȡ����Ϣ����ϸ���� ========
        If Left(MyStr(1), 2) = "00" Then
            RetSMS.SourceNo = Unicode2GB(MyStr(1))
        Else
            RetSMS.SourceNo = MyStr(1)
        End If
        RetSMS.SmsMain = Unicode2GB(strTmp3)
        
        '======== ���ʱ���к���ʱ������ȥ�� ========
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

