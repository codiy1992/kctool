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
Private blIsWaiting As Boolean ' �Ƿ�ȴ���һ֡����
Public blIsinit As Boolean     ' �Ƿ�����豸��ʼ��
Public blIsCheck As Boolean    ' �Ƿ�����豸���
Public blIsPickSms As Boolean  ' �Ƿ���ȡ����

Private strAtTemp  As String    ' AT��Ӧ�����ݴ���
Public Iccid As String       ' ICCID����
Public Imei As String        ' IMEI��
Public Imsi As String        ' IMSI��
Public Mobile As String      ' �ֻ���
Public bMobile As String     ' ��ת�ֻ���

Public Task As Task          ' ��������COM�ڵ������������

Private Sub Class_Initialize()
    blIsWaiting = False
    blIsinit = True
    blIsCheck = True
    blIsPickSms = True
    strAtTemp = ""
    Iccid = ""
    Imei = ""
    Imsi = ""
    Mobile = ""
    Set Task = New Task
End Sub

Public Function GetData(ByRef strInput As String) As String
    Dim iLen As Integer
    If blIsWaiting = False Then
        If UCase(Left(strInput, 2)) = "AT" Then         '��������µ�AT������Ӧ
            ' �ж�AT��������,�����ȡ����(AT+CMGL),������COM����ʱ��������ȡ��������
            iLen = InStr(strInput, vbCrLf)
            If iLen > 0 Then
                If InStr(UCase(Left(strInput, iLen - 1)), "AT+CMGL") > 0 Then
                    blIsPickSms = False
                End If
            End If
            
            ' �ж��Ƿ�Ϊ������AT������Ӧ����
            If InStr(strInput, "OK") > 0 Or InStr(strInput, "ERROR") > 0 Then
                GetData = strInput
                strAtTemp = ""
                blIsWaiting = False
            Else
                strAtTemp = strInput
                blIsWaiting = True
            End If
        Else    '������AT��ͷ�Ŀ����Ǵ���֪ͨ "+CMTI:"(�¶��ŵ���)
            GetData = strInput
            strAtTemp = ""
            blIsWaiting = False
        End If
    Else
        If InStr(strInput, "OK") > 0 Or (strAtTemp & strInput = "AT+CFUN=1,1" & vbCrLf & "IIII") Then
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
    'ElseIf UCase(Left(strInput, 6)) = "+CMGS:" Then '���ŷ��ͳɹ�
    '    strAT = "+CMGS:"
    ElseIf UCase(Left(strInput, 6)) = "+CMTI:" Then '�յ��¶���
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
'        Case InStr(strAT, "AT+CMGL") > 0
'            strAT = "AT+CMGL"
'            strOut = strInput
'        Case "AT+CMGL=""ALL"",1"
'            strInput = Right(strInput, Len(strInput) - 17)
'            strOut = Left(strInput, Len(strInput) - 4)
        Case "AT+CFUN=1,1"
            If InStr(strInput, "IIII") Then
                strAT = "-AT-INIT-OK-"
                strOut = ""
            End If
        Case Else
            If InStr(strAT, "AT+CMGL") > 0 Then
                strAT = "AT+CMGL"
                strOut = strInput
            ElseIf InStr(strAT, "AT+CMGS=") > 0 And InStr(strInput, "+CMGS:") > 0 Then
                strAT = "-AT-SMS-SEND-OK"  ' ���ŷ��ͳɹ�
                strOut = strInput
            Else
                strOut = ""
            End If
    End Select
    AnalysisData = strAT
End Function


Public Function sendSMS(strMobile As String, strCnt As String, Optional blIsCH As Boolean)
    strMobile = GB2Unicode(strMobile)
    strCnt = GB2Unicode(strCnt)
    Task.Push ("AT+CMGF=1" & vbCrLf)          ' ���Ÿ�ʽѡ�� 0:PDU 1:�ı�
    Task.Push ("AT+CSCS=""UCS2""" & vbCrLf)   ' �ַ���ѡ��
    Task.Push ("AT+CSCA?" & vbCrLf)           ' �Զ����ö���Ϣ���ĺ��룬ֻ����ѯ����
    Task.Push ("AT+CNMI=2,1" & vbCrLf)        ' ����Ϣָʾ��ʽ "+CMTI:��SM��,3"
    If blIsCH <> True Then
        Task.Push ("AT+CSMP=17,167,0,8" & vbCrLf) ' �������Ķ���
    Else
        Task.Push ("AT+CSMP=17,167,0,0" & vbCrLf) ' ����Ӣ�Ķ���
    End If
    Task.Push ("AT+CMGS=""" & strMobile & """" & vbCrLf)
    Task.Push (strCnt)
    Task.Push ("--SEND--")
End Function

Public Function delSMS(iSmsIndex As Integer, Optional blDelAll As Boolean)
    If blDelAll = True Then
        Task.Push ("AT+CMGD=0,2" & vbCrLf)
    Else
        If IsNumeric(iSmsIndex) And 1 <= iSmsIndex And iSmsIndex <= 50 Then
            Task.Push ("AT+CMGD=" & iSmsIndex & vbCrLf)  ' ɾ��ָ��λ�ö���
        End If
    End If
End Function
Public Function regExpGetAT(ByRef strInput As String) As String
   Dim objRegExp As RegExp                 '����objRegExpΪһ���������
   Dim Matches   As MatchCollection     '����MatchesΪƥ��������
   
   Set objRegExp = New RegExp               '��ʼ��һ���µ��������objRegExp
   objRegExp.Pattern = "(AT[^\s]*)"            '���ø����������������ʽ
   objRegExp.IgnoreCase = False              '�Ƿ����ִ�Сд
   objRegExp.Global = True                  '�Ƿ�ȫ��ƥ��
   
   If (objRegExp.Test(strInput) = True) Then '�����Ƿ���ƥ�䵽������Ҫ���ַ���
        Set Matches = objRegExp.Execute(strInput)   '��ʼ����
        regExpGetAT = Matches(0).SubMatches(0)
   Else
        regExpGetAT = ""
   End If
End Function
Public Function regExpGetFirst(strPat As String, strInput As String) As String
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

