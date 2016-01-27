Attribute VB_Name = "mdlSmsFuns"
Option Explicit
Public Function PickAllSMS(ByRef InputString As String, RetSMS() As SMSDef) As String

    Dim i As Integer, iTmp As Integer, iLen As Integer, iNext As Integer, iCr As Integer
    
    Dim n As Long
    
    Dim strTmp As String, strTmp1 As String, strTmp2 As String
    
    Dim btTmp() As Byte, btTmp2() As Byte
    
    Dim blRet As Boolean
    
On Error Resume Next
    
    strTmp = ""
    btTmp = InputString
    
    '======== ������Ϣ�е�˫����ȥ�� ========
    iTmp = 0
    For i = 0 To UBound(btTmp)
        strTmp1 = Chr(btTmp(i))
        If strTmp1 <> """" And btTmp(i) <> 0 And strTmp1 <> vbLf Then
            ReDim Preserve btTmp2(0 To iTmp + 1)
            btTmp2(iTmp) = btTmp(i)
            btTmp2(iTmp + 1) = 0
            iTmp = iTmp + 2
        End If
    Next i
    InputString = btTmp2
    
    n = 0
    i = 1
    Do
        iTmp = InStr(i, InputString, "+CMGL:")
        iCr = InStr(iTmp, InputString, vbCr)
        
        If iTmp > 0 Then
            If iCr - iTmp + 1 > 0 Then n = n + 1
        ElseIf iTmp = 0 Then
            Exit Do
        End If
        i = iTmp + 7
    Loop
    
    If n > 0 Then
        ReDim RetSMS(1 To n)
    Else
        ReDim RetSMS(0 To 0)
    End If
    
    '======== �������浽���ݿ��� ========
    For i = 1 To n
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
            blRet = PickOneSMS(strTmp, RetSMS(i), True)
            If blRet Then
'                RetSMS(i).SmsIndex = i
                On Error GoTo ErrorNode
ErrorNode:
            End If
        End If
    Next i

    PickAllSMS = "����" & n & "������"
End Function

Public Function PickOneSMS(strInputData As String, RetSMS As SMSDef, ByVal blIsList As Boolean) As Boolean
    
    Dim blRetFunc       As Boolean
    
    Dim i As Integer, iLen As Integer, iCr As Integer
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
            For i = 0 To nD - 2
                MyStr(i) = aryTmp(i + 1)
            Next i
        '======== ���򣬴��͹�������Ϣ��ʽ��"CMGR"����������������ġ� ========
        Else
            ReDim MyStr(0 To nD - 1)
            For i = 0 To nD - 1
                MyStr(i) = aryTmp(i)
            Next i
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

