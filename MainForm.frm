VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form MainForm 
   Caption         =   "�������� -- by CODIY"
   ClientHeight    =   12090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   21720
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12090
   ScaleWidth      =   21720
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox TextCom1Imsi 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   9000
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox TextCom1Mobile 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   7560
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "17070594726"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   8895
      Left            =   16800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   600
      Width           =   4335
   End
   Begin VB.Timer TimerTaskCom1 
      Interval        =   300
      Left            =   240
      Top             =   11400
   End
   Begin VB.TextBox TextCom1Iccid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   5520
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox TextCom1Net 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   3840
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox TextCom1Signal 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   270
      Left            =   3240
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   720
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ǰ���Դ���(��)"
      Height          =   9735
      Left            =   11520
      TabIndex        =   9
      Top             =   120
      Width           =   5175
      Begin VB.TextBox TextRec 
         Height          =   8895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label_num_rec 
         Caption         =   "0"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "R:"
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label_num_send 
         Caption         =   "0"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "S:"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "��������"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   14880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      BaudRate        =   115200
   End
   Begin VB.Frame Frame2 
      Caption         =   "������ť"
      Height          =   1575
      Left            =   360
      TabIndex        =   4
      Top             =   8280
      Width           =   9135
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   375
         Left            =   5280
         TabIndex        =   39
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "дINI"
         Height          =   375
         Left            =   3600
         TabIndex        =   38
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "��ʼУ��"
         Height          =   375
         Left            =   1800
         TabIndex        =   37
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ɾ�����ж���"
         Height          =   375
         Left            =   7440
         TabIndex        =   36
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ɾ������"
         Height          =   375
         Left            =   7680
         TabIndex        =   35
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CommandStopCheck 
         Caption         =   "ֹͣ����"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton CommandSendSms 
         Caption         =   "���Ͷ���"
         Height          =   375
         Left            =   3600
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton CommandCloseAll 
         Caption         =   "�ر�����"
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CommandStartAll 
         Caption         =   "��������"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Timer TimerCom1CheckStat 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   10920
   End
   Begin VB.Frame Frame1 
      Caption         =   "����״̬"
      Height          =   8055
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   11055
      Begin VB.TextBox TextCom3Stat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "δ����"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox TextCom2Stat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "δ����"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TextCom1Stat 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "δ����"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "�ֻ���"
         Height          =   255
         Left            =   7560
         TabIndex        =   31
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "IMSI��"
         Height          =   255
         Left            =   8880
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "ICCID��"
         Height          =   255
         Left            =   5760
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "����״̬"
         Height          =   255
         Left            =   3960
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "�ź�"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "����״̬"
         Height          =   255
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "���ں�"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LabelCom3 
         Caption         =   "COM*"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   495
         TabIndex        =   18
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label Label6 
         Caption         =   "03-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "02-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   375
      End
      Begin VB.Label LabelCom2 
         Caption         =   "COM*"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   495
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "01-"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   375
      End
      Begin VB.Label LabelCom1 
         Caption         =   "COM*"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   495
         TabIndex        =   1
         Top             =   600
         Width           =   600
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   10200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Com1 As Com
Dim DB As New DB
Private Sub Command5_Click()
    Dim smsArr
    smsArr = DB.NotSendedSMS("'89860115841028567295'")
    For i = 0 To UBound(smsArr)
        
        'smsArr(i, 0) & vbCrLf & smsArr(i, 1) & vbCrLf & smsArr(i, 2) & vbCrLf & smsArr(i, 3) & vbCrLf
    Next i
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim comport_use() As String
    ' ������ߴ���
    Call comportScan(comport_use())
    For i = 0 To UBound(comport_use()) - 1
        Select Case i
        Case 1
        LabelCom1.Caption = "COM" & comport_use(i)
        'Set taskCom1 = New Task
        Set Com1 = New Com
        End Select
    Next i
End Sub
Private Sub CommandStartAll_Click()
    Dim temp As String
    ' ��������1
    If Val(Right(LabelCom1.Caption, 1)) > 0 And LabelCom1.ForeColor = &H80000012 Then
        MSComm1.CommPort = Val(Right(LabelCom1.Caption, 1))
        On Error Resume Next        '�жϸô����Ƿ񱻴�
        MSComm1.PortOpen = True
        If Err = 8005 Then
            MSComm1.PortOpen = False
            TextCom1Stat.Text = "�˿ڱ�ռ��"
        Else
            With MSComm1
                .Settings = "115200,n,8,1" 'Trim(Combo_rate.Text) & "," & Left(Trim(Combo_jiaoyan.Text), 1) & "," & Trim(Combo_databyte.Text) & "," & Trim(Combo_stopbyte)
                .RThreshold = 1
                .InputMode = comInputModeBinary
            End With
            TextCom1Stat.Text = "�ѿ���"
            LabelCom1.ForeColor = &H8000000D
            Com1.Task.Push ("ATE1" & vbCrLf)
            Com1.Task.Push ("AT+CFUN=1,1" & vbCrLf)
            ' ������⴮��״̬��ʱ��
            'TimerCom1CheckStat.Enabled = True
            TimerTaskCom1.Enabled = True
        End If
    End If
    
End Sub
Private Sub CommandCloseAll_Click()
    ' �رմ���1
    TimerTaskCom1.Enabled = False
    TimerCom1CheckStat.Enabled = False
    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False
        TextCom1Stat.Text = "δ����"
        LabelCom1.ForeColor = &H80000012
        TextCom1Signal.Text = ""
        TextCom1Net.Text = ""
    End If
End Sub

Private Sub CommandSendSms_Click()
    TimerCom1CheckStat.Enabled = False
    Com1.sendSMS "10027", "102", True
    TimerTaskCom1.Enabled = True
End Sub

Private Sub Command3_Click()
    Com1.blIsCheck = True
    Com1.blIsPickSms = True
End Sub

Private Sub CommandStopCheck_Click()
    Com1.blIsCheck = False
    Com1.blIsPickSms = False
End Sub
Private Sub Command1_Click()
    Com1.delSMS 1
End Sub

Private Sub Command2_Click()
    Com1.delSMS 1, True
End Sub

Private Sub TimerCom1CheckStat_Timer()
    If Com1.blIsinit = True Then
        Com1.Task.Push ("ATE1" & vbCrLf)         ' ��������
        'Com1.Task.Push ("AT+CNMI=2,1" & vbCrLf)  '
        Com1.Task.Push ("AT+CMGF=1" & vbCrLf)    ' ���ö��Ÿ�ʽ
        Com1.blIsinit = False
        'Com1.Task.Push ("--SEND--")
    End If
    
    If Com1.blIsCheck = True Then
        If Com1.Iccid = "" Then
            Com1.Task.Push ("AT+CCID" & vbCrLf)      ' ��ѯICCID��
        End If
        If Com1.Imei = "" Then
            Com1.Task.Push ("AT+CGSN" & vbCrLf)      ' Imei
        End If
        If Com1.Imsi = "" Then
            Com1.Task.Push ("AT+CIMI" & vbCrLf)
        End If
        
        Com1.Task.Push ("AT+CSQ" & vbCrLf)
        Com1.Task.Push ("AT+COPS?" & vbCrLf)
    End If
    
    If Com1.blIsPickSms = True Then
        Com1.Task.Push ("AT+CMGL=""ALL""" & vbCrLf)
        Com1.blIsPickSms = False ' ��������ǰ�������ٷ����µ�ȡ��������
    End If
    If TimerTaskCom1.Enabled = False Then TimerTaskCom1.Enabled = True
End Sub

Private Sub TimerTaskCom1_Timer()
    Dim Task As String
    Dim tail(0) As Byte
    Task = Com1.Task.Pop
    If Task <> Empty Then
         'TextRec.Text = TextRec.Text & task
        If MSComm1.PortOpen = True Then
            If Task = "--SEND--" Then
                tail(0) = &H1A
                MSComm1.Output = tail
            Else
                MSComm1.Output = Task
            End If
        End If
    Else
       TimerTaskCom1.Enabled = False
    End If
    
End Sub


Private Sub MSComm1_OnComm()
    Dim strAtData As String
    Dim tmpBuf() As Byte, strTmp As String
    Dim strOut As String
    Dim strAT As String
    Dim smsArr() As SMSDef
On Error Resume Next
    Select Case MSComm1.CommEvent

        '''''''''''''''''''''''''''''''''''''''
        Case comEvReceive
            tmpBuf() = MSComm1.Input
            For i = 0 To UBound(tmpBuf())
                strTmp = strTmp & Chr(Val(tmpBuf(i)))
            Next i
            Text1.Text = Text1.Text & strTmp ' & "------------------"
            Text1.SelStart = Len(Text1.Text)
            strAtData = Com1.GetData(strTmp)
            If strAtData <> Empty Then
                strAT = Com1.AnalysisData(strAtData, strOut)
                If strAT <> "" And strAT <> vbCr And strAtData <> vbCrLf And Not IsEmpty(strAT) Then
                    'TextRec.Text = TextRec.Text & strOut & "------------------" & vbCrLf
                    Select Case strAT
                        Case "AT+CSQ"
                            TextCom1Signal.Text = strOut
                        Case "AT+COPS?"
                            TextCom1Net.Text = strOut
                        Case "AT+CGSN"
                            Com1.Imei = strOut
                        Case "AT+CIMI"
                            TextCom1Imsi.Text = strOut
                            Com1.Imsi = strOut
                        Case "AT+CCID"
                            TextCom1Iccid.Text = strOut
                            Com1.Iccid = strOut
                        Case "AT+CMGL"
                            Com1.blIsPickSms = False    ' �������ǰ,�����ڽ���ȡ��������
                            strOut = PickAllSMS(strOut, smsArr)
                            If UBound(smsArr) > 0 Then
                                For n = 1 To UBound(smsArr)
                                    TextRec.Text = TextRec.Text & vbCrLf & smsArr(n).SmsIndex & "   " _
                                    & Format(smsArr(n).ReachDate, "YYYY-MM-DD") & " " & Format(smsArr(n).ReachTime, "HH:MM:SS") & vbTab & smsArr(n).SourceNo & vbCrLf _
                                     & vbCrLf & smsArr(n).SmsMain & vbCrLf & "-------------------------------------" & vbCrLf
                                     If Com1.Iccid <> "" Then
                                        DB.SaveSMS Com1.Iccid, smsArr(n).SourceNo, smsArr(n).SmsMain, smsArr(n).DateTime, "17070594726"
                                        Com1.Task.Push ("AT+CMGD=" & smsArr(n).SmsIndex & vbCrLf)
                                     End If
                                Next n
                            End If
                            TextRec.SelStart = Len(TextRec.Text)
                            Com1.blIsPickSms = True    ' ������ɺ�,���������µ�ȡ��������
                        Case "+CMTI:" ' �յ��¶���
                        Case "-AT-SMS-SEND-OK" ' ���ŷ��ͳɹ�
                            TimerCom1CheckStat.Enabled = True
                        Case "-AT-INIT-OK-"
                            TimerCom1CheckStat.Enabled = True
                        Case Else
                    End Select
                        
                End If
            End If
            
        '''''''''''''''''''''''''''''''''''''''
        Case comEventBreak
            TextCom1Stat.Text = "Modem�����ж��źţ�ϣ��������ܵȺ����Ժ�."
            MSComm1.PortOpen = False
            MSComm1.PortOpen = True
        Case comEvCTS
            If MSComm1.CTSHolding = True Then 'Modem��ʾ��������Է�������
                TextCom1Stat.Text = "Modem�ܹ����ռ��������"
            Else 'Modem�޷���Ӧ��������ݣ����ܻ���������
                TextCom1Stat.Text = "Modem����������ʱ��Ҫ��������"
                MSComm1.DTREnable = Not MSComm1.DTREnable
                DoEvents
                MSComm1.DTREnable = Not MSComm1.DTREnable
            End If
        Case comEvDSR
            If MSComm1.DSRHolding = True Then '��Modem�յ�������Ѿ�������Modem��ʾ�Լ�Ҳ����
                TextCom1Stat.Text = "Modem���Ը��������������"
            Else '�ڼ��������DTR�źź�Modem���ܻ�û�о���
                TextCom1Stat.Text = "Modem��û�г�ʼ�����"
            End If
        Case comEventFrame
            MSComm1.PortOpen = False
            MSComm1.PortOpen = True
        Case comEvRing
            TextCom1Stat.Text = "��⵽����仯"
        Case comEvCD
            TextCom1Stat.Text = "��⵽�ز��仯"
        Case Else
            MsgBox MSComm1.CommEvent
    End Select

End Sub
'**********************************************************************
' ����ɨ��
'**********************************************************************
Function comportScan(comport() As String)
    Dim i As Integer
    ReDim Preserve comport(0)
    For i = 1 To 255
        On Error Resume Next
        MSComm.CommPort = i
        MSComm.PortOpen = True
        If Err = 0 Or Err = 8005 Then
        comport(UBound(comport())) = i
        ReDim Preserve comport(UBound(comport()) + 1)
        MSComm.PortOpen = False
        End If
    Next i
End Function
'**********************************************************************
' ���Ͷ���
'**********************************************************************
Public Function sendSMS(ByRef strIccid As String, ByRef objTask As Task, strMobile As String, strCnt As String)
    objCOM.sendSMS objTask, strMobile, strCnt
End Function
