VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������� -- by CODIY"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15000
   ForeColor       =   &H80000008&
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   15000
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   17
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   17
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   17
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   16
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   16
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   16
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   15
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   15
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   15
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   14
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   14
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   14
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   13
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   13
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   13
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   12
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   12
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   12
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   3000
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   200
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   200
      Left            =   3600
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   960
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   200
      Left            =   2280
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   300
      Left            =   2880
      Top             =   9840
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   3000
      Left            =   3240
      Top             =   9840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   9360
   End
   Begin VB.Timer TimerPickDBTask 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   960
      Top             =   9360
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   3000
      Left            =   1920
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   300
      Left            =   1560
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   300
      Left            =   240
      Top             =   9840
   End
   Begin VB.Frame Frame3 
      Caption         =   "��ǰ���Դ���(��)"
      Height          =   9015
      Left            =   15000
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.TextBox TextLog 
         Height          =   4455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox TextRec 
         Height          =   3855
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   5040
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "�յ��Ķ���:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4800
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "������ť"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   8160
      Width           =   14775
      Begin VB.CommandButton Command1 
         Caption         =   "�������ݿ�"
         Height          =   375
         Left            =   11400
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton CommandShow 
         Caption         =   "���ݴ���"
         Height          =   375
         Left            =   12840
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton CommandCloseAll 
         Caption         =   "�ر�����"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CommandStartAll 
         Caption         =   "�����豸"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   3000
      Left            =   600
      Top             =   9840
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����б�"
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      Begin MSComctlLib.ListView ListView 
         Height          =   7545
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   14505
         _ExtentX        =   25585
         _ExtentY        =   13309
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483637
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Menu MenuList 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu MENU_OPEN 
         Caption         =   "����������"
         Shortcut        =   ^O
      End
      Begin VB.Menu MENU_CLOSE 
         Caption         =   "�رձ�����"
         Shortcut        =   ^C
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_STOP 
         Caption         =   "��ͣ������"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_START 
         Caption         =   "�ָ�������"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_DEBUG 
         Caption         =   "�鿴������"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MainForm.Width = 15090
    ' ��ʼ������
    g_iDebugIndex = -1
    g_show = False
    ' ��ʼ���б���ͼ�ؼ�
    With ListView 'ListView��ʼ��
         .View = 3 ' �б���ʾ��ʽ
         .ColumnHeaders.Add , , "���", 600         ' 0
         .ColumnHeaders.Add , , "����", 800         ' 1
         .ColumnHeaders.Add , , "״̬", 1600        ' 2
         .ColumnHeaders.Add , , "�ֻ�����", 1500    ' 3
         .ColumnHeaders.Add , , "ICCID��", 2300     ' 4
         .ColumnHeaders.Add , , "IMEI��", 1800      ' 5
         .ColumnHeaders.Add , , "IMSI��", 1800      ' 6
         .ColumnHeaders.Add , , "��ת����", 1500    ' 7
         .ColumnHeaders.Add , , "�ź�", 1000         ' 8
         .ColumnHeaders.Add , , "����", 1600    ' 9
    End With
    
    TimerPickDBTask.Enabled = True
End Sub
Private Sub CommandStartAll_Click()
    Dim cIdx As Integer
    If IsComEmpty = False Then
        For cIdx = 0 To UBound(Com())
            If Com(cIdx).blIsOpen = False Then
                Com(cIdx).OpenPort
                If Com(cIdx).blIsOpen = False Then
                    ListView.ListItems(cIdx + 1).SubItems(2) = ComErr(Com(cIdx).portErr)
                Else
                    ListView.ListItems(cIdx + 1).SubItems(2) = "��ʼ����(0)..."
                    Com(cIdx).task.Push ("ATE1" & vbCrLf)
                    Com(cIdx).task.Push ("AT+CFUN=1,1" & vbCrLf)
                    ' ����������������ʱ��
                    TimerComTask(cIdx).Enabled = True
                    TimerComRead(cIdx).Enabled = True
                End If
            End If
        Next cIdx
    ElseIf Timer1.Enabled = False Then
        ' ������ߴ���
        CommandStartAll.Caption = "������..."
        Timer1.Enabled = True
    End If
End Sub
Private Sub CommandCloseAll_Click()
    ' �رմ���
    If IsComEmpty = False Then
        For cIdx = 0 To UBound(Com())
            'TimerComTask(cIdx).Enabled = False
            DB.setCardClose ("'" & Com(cIdx).Iccid & "'")
            ' ��ֹͣ��������������ִ�ж���
            Com(cIdx).task.Push ("--CLOSE--")
            TimerComTask(cIdx).Enabled = True
            TimerComCheck(cIdx).Enabled = False
        Next cIdx
    End If
End Sub
Private Sub CommandShow_Click()
    If g_show = False Then
        g_show = True
        MainForm.Width = 19400
    Else
        g_show = False
        MainForm.Width = 15090
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case vbFormControlMenu          ' 0 �û��Ӵ����ϵġ��ؼ����˵���ѡ�񡰹رա�ָ�
        Case vbFormCode                 ' 1 Unload ��䱻������á�
        Case vbAppWindows               ' 2 ��ǰ Microsoft Windows ���������Ự������
        Case vbAppTaskManager           ' 3 Microsoft Windows ������������ڹر�Ӧ�ó���
        Case vbFormMDIForm              ' 4 MDI �Ӵ������ڹرգ���Ϊ MDI �������ڹرա�
        Case vbFormOwner                ' 5 ��Ϊ��������������ڹرգ����Դ���Ҳ�ڹرա�
    End Select
    If IsComEmpty = False Then
        If GetAllIccid <> "''" Then
            If MsgBox("���ȹر�ȫ�����ں��ٲ���", vbYesNo + vbDefaultButton1) = vbYes Then
                For cIdx = 0 To UBound(Com())
                    DB.setCardClose ("'" & Com(cIdx).Iccid & "'")
                    ' ��ֹͣ��������������ִ�ж���
                    Com(cIdx).task.Push ("--CLOSE--")
                    TimerComTask(cIdx).Enabled = True
                    TimerComCheck(cIdx).Enabled = False
                Next cIdx
            End If
            Cancel = True
        End If
    End If
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    Dim comport_use() As String
    Call comportScan(comport_use())
    If comport_use(0) <> "" Then
        ReDim Com(UBound(comport_use()))
        For i = 0 To UBound(comport_use())
            Set LV = ListView.ListItems.Add(i + 1, , i + 1)
            LV.SubItems(1) = "COM" & comport_use(i)
            Set Com(i) = New Com
            Com(i).comPort = comport_use(i)
        Next i
        CommandStartAll.Caption = "��������"
    Else
        CommandStartAll.Caption = "�޿����豸"
    End If
    Timer1.Enabled = False
End Sub

Private Sub TimerComCheck_Timer(cIdx As Integer)
    Dim exec As Integer, pick As Integer
    If Com(cIdx).blIsATExecing = True Then
        exec = 1
    End If
    If Com(cIdx).blIsPickSms = True Then
        pick = 1
    End If
    ' ������ڷ����Ż������ú�ת,����������������µ���������
    If Com(cIdx).iSmsId > 0 Or Com(cIdx).iBindId > 0 Then
        ListView.ListItems(cIdx + 1).SubItems(2) = "[" & Com(cIdx).iSmsId & "|" & Com(cIdx).iBindId & "]"
        Exit Sub
    End If
    
    Com(cIdx).task.Push ("AT+CSQ" & vbCrLf)
    Com(cIdx).task.Push ("AT+COPS?" & vbCrLf)
    
    If Com(cIdx).Imei = "" Then
        Com(cIdx).task.Push ("AT+CGSN" & vbCrLf)      ' Imei
    End If
    If Com(cIdx).Imsi = "" Then
        Com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
    End If
    
    If Com(cIdx).blIsPickSms = True Then
        Com(cIdx).task.Push ("AT+CMGL=""ALL""" & vbCrLf)
        Com(cIdx).blIsPickSms = False ' ��������ǰ�������ٷ����µ�ȡ��������
    End If
    
    If Com(cIdx).bMobile = "" Then
        Com(cIdx).task.Push ("AT+CCFC=0,2" & vbCrLf)
        'TimerComCheck(cIdx).Enabled = False
    End If
    
    
    ListView.ListItems(cIdx + 1).SubItems(2) = Com(cIdx).task.wIndex & "-" & Com(cIdx).task.rIndex & "-" & _
                                                exec & "-" & pick & "-" & Com(cIdx).iWaitCnt
    TimerComTask(cIdx).Enabled = True
End Sub

Private Sub TimerComTask_Timer(cIdx As Integer)
    Dim strAT As String
    Dim Top As String
    Dim tail(0) As Byte
    
    Top = Com(cIdx).task.Top
    
    If UCase(Left(Top, 2)) = "AT" And Com(cIdx).blIsATExecing = True Then
        Exit Sub
    End If
    
    strAT = Com(cIdx).task.Pop
    If strAT <> Empty Then
        If Com(cIdx).blIsOpen = True Then
            Select Case strAT
                Case "--TAIL--"
                    tail(0) = &H1A
                    Com(cIdx).WriteData (tail)
                Case "--STOP--"
                    TimerComTask(cIdx).Enabled = False
                    ListView.ListItems(cIdx + 1).SubItems(2) = "����ͣ"
                Case "--CLOSE--"
                    If Com(cIdx).blIsOpen = True Then
                        Com(cIdx).ClosePort
                    End If
                    Com(cIdx).ReSet
                    TimerComTask(cIdx).Enabled = False
                    TimerComRead(cIdx).Enabled = False
                    ListView.ListItems(cIdx + 1).SubItems(2) = "δ����"
                Case Else
                    'Or InStr(strAT, "AT+CMGL") > 0 Or InStr(strAT, "AT+CCFC") > 0
                    If UCase(Left(Top, 2)) = "AT" Then
                        Com(cIdx).blIsATExecing = True
                    End If
                    Com(cIdx).WriteData (strAT)
            End Select
        End If
    Else
       TimerComTask(cIdx).Enabled = False
    End If
End Sub
Private Sub TimerComRead_Timer(cIdx As Integer)
    Dim strAtData As String
    Dim tmpBuf() As Byte, strTmp As String
    Dim strOut As String
    Dim strAT As String
    Dim smsArr() As SMSDef
    strTmp = Com(cIdx).ReadData
    If Com(cIdx).iWaitCnt >= 10 Then
        Com(cIdx).iWaitCnt = 0
        Com(cIdx).blIsATExecing = False
    End If
    If strTmp = "" Then 'Com(cIdx).blIsATExecing = False Or
        Exit Sub
    End If
    If Com(cIdx).blIsOpen = True Then
        If cIdx = g_iDebugIndex Then
            TextLog.Text = TextLog.Text & strTmp ' & vbCrLf & "------------------" & vbCrLf
            TextLog.SelStart = Len(TextLog.Text)
        End If
        strAtData = Com(cIdx).GetData(strTmp)
        If strAtData <> Empty Then
            strAT = Com(cIdx).AnalysisData(strAtData, strOut)
            If strAT <> "" And strAT <> vbCr And strAtData <> vbCrLf And Not IsEmpty(strAT) Then
                'TextRec.Text = TextRec.Text & strOut & "------------------" & vbCrLf
                Select Case strAT
                    Case "AT+CSQ"
                        ListView.ListItems(cIdx + 1).SubItems(8) = strOut
                        Com(cIdx).blIsATExecing = False
                    Case "AT+COPS?"
                        ListView.ListItems(cIdx + 1).SubItems(9) = strOut
                        Com(cIdx).blIsATExecing = False
                    Case "AT+CGSN"
                        Com(cIdx).Imei = strOut
                        ListView.ListItems(cIdx + 1).SubItems(5) = strOut
                        Com(cIdx).blIsATExecing = False
                    Case "AT+CIMI"
                        Com(cIdx).Imsi = strOut
                        ListView.ListItems(cIdx + 1).SubItems(6) = strOut
                        Com(cIdx).blIsATExecing = False
                    Case "AT+CCFC=0,2"
                        Com(cIdx).bMobile = strOut
                        ListView.ListItems(cIdx + 1).SubItems(7) = strOut
                        If Com(cIdx).iQccfcCnt >= 15 And strOut = "" Then
                            Com(cIdx).bMobile = "0"
                            ListView.ListItems(cIdx + 1).SubItems(7) = "(��)"
                        End If
                        Com(cIdx).blIsATExecing = False
                    Case "-AT-BIND-MOBILE-OK-"
                        DB.SetBinded Com(cIdx).Iccid, Com(cIdx).iBindId
                        Com(cIdx).iBindId = 0
                        Com(cIdx).blIsATExecing = False
                    Case "-AT-BIND-MOBILE-FAILED-"
                        DB.SetNotBind Com(cIdx).Iccid, Com(cIdx).iBindId
                        Com(cIdx).iBindId = 0
                        Com(cIdx).blIsATExecing = False
                    Case "-AT-UNBIND-MOBILE-OK-"
                        DB.SetBinded Com(cIdx).Iccid, Com(cIdx).iBindId
                        Com(cIdx).iBindId = 0
                        Com(cIdx).bMobile = ""
                        Com(cIdx).blIsATExecing = False
                        ListView.ListItems(cIdx + 1).SubItems(7) = ""
                    Case "-AT-UNBIND-MOBILE-FAILED-"
                        DB.SetNotBind Com(cIdx).Iccid, Com(cIdx).iBindId
                        Com(cIdx).iBindId = 0
                        Com(cIdx).blIsATExecing = False
                    Case "AT+CMGL"
                        If InStr(strOut, "ERROR") Then
                            Com(cIdx).task.Push ("AT+CMGF=1" & vbCrLf)    ' ���ö��Ÿ�ʽ
                        End If
                        strOut = PickAllSMS(strOut, smsArr)
                        If UBound(smsArr) > 0 Then
                            For n = 1 To UBound(smsArr)
                                 If cIdx = g_iDebugIndex Then
                                     TextRec.Text = TextRec.Text & vbCrLf & smsArr(n).SmsIndex & vbTab _
                                                    & Format(smsArr(n).ReachDate, "YYYY-MM-DD") & vbTab _
                                                    & Format(smsArr(n).ReachTime, "HH:MM:SS") & vbTab _
                                                    & smsArr(n).SourceNo & vbCrLf _
                                                    & smsArr(n).SmsMain & vbCrLf _
                                                    & "-------------------------------------" & vbCrLf
                                     TextRec.SelStart = Len(TextRec.Text)
                                 End If
                                 If Com(cIdx).Iccid <> "" Then
                                    DB.SaveSMS Com(cIdx).Iccid, smsArr(n).SourceNo, smsArr(n).SmsMain, smsArr(n).DateTime, Com(cIdx).Mobile
                                    Com(cIdx).task.Push ("AT+CMGD=" & smsArr(n).SmsIndex & vbCrLf)
                                 End If
                            Next n
                        End If
                        Com(cIdx).blIsPickSms = True    ' ������ɺ�,���������µ�ȡ��������
                        Com(cIdx).blIsATExecing = False
                    Case "+CMTI:" ' �յ��¶���
                    Case "-AT-SMS-SEND-OK-" ' ���ŷ��ͳɹ�
                        DB.SetSMSSended Com(cIdx).Iccid, Com(cIdx).iSmsId
                        Com(cIdx).iSmsId = 0
                        'TimerComCheck(cIdx).Enabled = True
                    Case "-AT-SMS-SEND-FAILED-" '���ŷ���ʧ��
                        DB.SetSMSNotSend Com(cIdx).Iccid, Com(cIdx).iSmsId
                        Com(cIdx).iSmsId = 0
                        'TimerComCheck(cIdx).Enabled = True
                    Case "-AT-INIT-OK-"    ' ��ʼ��������һ���ɹ���
                        Com(cIdx).task.Push ("ATE1" & vbCrLf)         ' ��������
                        Com(cIdx).task.Push ("AT+CIURC=0" & vbCrLf)
                        Com(cIdx).task.Push ("AT+CGSN" & vbCrLf)
                        Com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
                        Com(cIdx).task.Push ("AT+CNMI=2,1" & vbCrLf)  '
                        Com(cIdx).task.Push ("AT+CMGF=1" & vbCrLf)    ' ���ö��Ÿ�ʽ
                        Com(cIdx).task.Push ("AT+CGSN" & vbCrLf)
                        Com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
                        Com(cIdx).task.Push ("AT+CCID" & vbCrLf)      ' ��ѯICCID��
                        ListView.ListItems(cIdx + 1).SubItems(2) = "��ʼ����(1)..."
                        ' ����������������ʱ��
                        TimerComTask(cIdx).Enabled = True
                        Com(cIdx).blIsATExecing = False
                    Case "-AT-NO-CCID-"
                        ListView.ListItems(cIdx + 1).SubItems(2) = "������SIM��"
                        Com(cIdx).blIsATExecing = False
                    Case "AT+CCID"         ' ��ʼ������������ɹ���
                        If strOut = "-RETRY-" Then
                            Com(cIdx).task.Push ("AT+CCID" & vbCrLf)      ' ��ѯICCID��
                        Else
                            Com(cIdx).Iccid = Left(strOut, 19)
                            ListView.ListItems(cIdx + 1).SubItems(4) = Left(strOut, 19)
                            ' �����ݿ�ע�᱾SIM��������ȡ�ֻ�����
                            Com(cIdx).Mobile = DB.RegistCard(Com(cIdx).Iccid, Com(cIdx).Imei, Com(cIdx).Imsi)
                            If Com(cIdx).Mobile = "" Then
                                Com(cIdx).ReSet
                                ListView.ListItems(cIdx + 1).SubItems(2) = "���������ֻ���"
                            Else
                                Com(cIdx).blIsOpen = True
                                ListView.ListItems(cIdx + 1).SubItems(2) = "��������"
                                ListView.ListItems(cIdx + 1).SubItems(3) = Com(cIdx).Mobile
                                TimerComCheck(cIdx).Enabled = True
                            End If
                        End If
                        Com(cIdx).blIsATExecing = False
                    Case Else
                        Com(cIdx).blIsATExecing = False
                End Select
                    
            End If
        End If
    End If
End Sub

Private Sub TimerPickDBTask_Timer()
    Dim i, j As Integer
    Dim cIdx As Integer
    Dim cntA As Integer, cntB As Integer
    Dim strIccids As String
    Dim execArr
    If IsComEmpty = True Then
        Exit Sub
    End If
    ' ִ�а���������ת����
    execArr = DB.NotExecBind(GetAllIccid)
    If LCase(TypeName(execArr)) = "string()" Then
        cntA = UBound(execArr) + 1
        For cIdx = 0 To UBound(Com())
            For i = 0 To UBound(execArr)
                If execArr(i, 0) = Com(cIdx).Iccid Then
                    Com(cIdx).iBindId = Val(execArr(i, 1))
                    If execArr(i, 2) <> "" Then
                        Com(cIdx).bindMobile (execArr(i, 2))
                    Else
                        Com(cIdx).unBindMobile
                    End If
                End If
                TimerComTask(cIdx).Enabled = True
            Next i
        Next cIdx
    End If
    ' ִ�з��Ͷ�������
    'Frame1.Caption = GetAllIccid(True)
    execArr = DB.NotSendedSMS(GetAllIccid(True))
    If LCase(TypeName(execArr)) = "string()" Then
        cntB = UBound(execArr) + 1
        For cIdx = 0 To UBound(Com())
            For i = 0 To UBound(execArr)
                If Com(cIdx).Iccid = execArr(i, 0) Then
                    Com(cIdx).iSmsId = Val(execArr(i, 1))
                    'TimerComCheck(cIdx).Enabled = False
                    If execArr(i, 2) = "10027" Then
                        Com(cIdx).sendSMS execArr(i, 2), execArr(i, 3), True
                    Else
                        Com(cIdx).sendSMS execArr(i, 2), execArr(i, 3)
                    End If
                    TimerComTask(cIdx).Enabled = True
                End If
            Next i
        Next cIdx
    End If
    Frame1.Caption = "�����б� - ִ��������[" & cntA & "][" & cntB & "](" & Now() & ")"
End Sub

Private Sub ListView_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim Index As Integer
    '��������Ҽ�
    If Button = vbRightButton Then
        If ListView.ListItems.Count > 0 Then
            Index = ListView.SelectedItem.Index
            If Com(Index - 1).blIsOpen = True Then
                ' �����رղ˵�
                MENU_CLOSE.Visible = True
                MENU_CLOSE.Enabled = True
                ' �رտ����˵�
                MENU_OPEN.Visible = False
                MENU_OPEN.Enabled = False
                
                If MENU_START.Visible = False Then
                    MENU_STOP.Visible = True
                    MENU_STOP.Enabled = True
                End If
            Else
                ' ���������˵�
                MENU_OPEN.Visible = True
                MENU_OPEN.Enabled = True
                ' �رչرղ˵�
                MENU_CLOSE.Visible = False
                MENU_CLOSE.Enabled = False
                ' �ر���ͣ�˵�
                MENU_STOP.Visible = False
                MENU_STOP.Enabled = False
                ' �رջָ��˵�
                MENU_START.Visible = False
                MENU_START.Enabled = False
            End If
            PopupMenu MenuList
        End If
    End If
End Sub
Private Sub MENU_OPEN_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.Index - 1
    If Com(cIdx).blIsOpen = False Then
        On Error Resume Next        '�жϸô����Ƿ񱻴�
        Com(cIdx).OpenPort
        If Com(cIdx).blIsOpen = False Then
            ListView.ListItems(cIdx + 1).SubItems(2) = ComErr(Com(cIdx).portErr)
        Else
            ListView.ListItems(cIdx + 1).SubItems(2) = "��ʼ����(0)..."
            Com(cIdx).task.Push ("ATE1" & vbCrLf)
            Com(cIdx).task.Push ("AT+CFUN=1,1" & vbCrLf)
            ' ����������������ʱ��
            TimerComTask(cIdx).Enabled = True
            TimerComRead(cIdx).Enabled = True
        End If
    Else
        Com(cIdx).ClosePort
        Com(cIdx).OpenPort
        If Com(cIdx).blIsOpen = False Then
            ListView.ListItems(cIdx + 1).SubItems(2) = ComErr(Com(cIdx).portErr)
        Else
            ListView.ListItems(cIdx + 1).SubItems(2) = "��ʼ����(0)..."
            Com(cIdx).task.Push ("ATE1" & vbCrLf)
            Com(cIdx).task.Push ("AT+CFUN=1,1" & vbCrLf)
            ' ����������������ʱ��
            TimerComTask(cIdx).Enabled = True
            TimerComRead(cIdx).Enabled = True
        End If
    End If
End Sub
Private Sub MENU_CLOSE_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.Index - 1
    ' �޸����ݿ��ֶ�״̬
    DB.setCardClose ("'" & Com(cIdx).Iccid & "'")
    ' ��ֹͣ��������������ִ�ж���
    Com(cIdx).task.Push ("--CLOSE--")
    TimerComTask(cIdx).Enabled = True
    TimerComCheck(cIdx).Enabled = False
End Sub
Private Sub MENU_STOP_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.Index - 1
    Com(cIdx).task.Push ("--STOP--") ' ��ֹͣ��������������ִ�ж���
    TimerComTask(cIdx).Enabled = True
    TimerComCheck(cIdx).Enabled = False
    ' �˵�
    MENU_START.Visible = True
    MENU_START.Enabled = True
    MENU_STOP.Visible = False
    MENU_STOP.Enabled = False
End Sub
Private Sub MENU_START_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.Index - 1
    TimerComTask(cIdx).Enabled = True
    TimerComCheck(cIdx).Enabled = True
    ListView.ListItems(cIdx + 1).SubItems(2) = "��������"
    ' �˵�
    MENU_START.Visible = False
    MENU_START.Enabled = False
    MENU_STOP.Visible = True
    MENU_STOP.Enabled = True
End Sub
Private Sub MENU_DEBUG_Click()
    If g_iDebugIndex = ListView.SelectedItem.Index - 1 Then
        g_iDebugIndex = -1
        Frame3.Caption = "��ǰ���Դ���(��)"
    Else
        g_iDebugIndex = ListView.SelectedItem.Index - 1
        TextLog.Text = ""
        Frame3.Caption = "��ǰ���Դ���(COM" & Com(g_iDebugIndex).comPort & ")"
    End If
End Sub

'**********************************************************************
' ����ɨ��
'**********************************************************************
Function comportScan(comPort() As String)
    Dim i As Integer
    ReDim Preserve comPort(0)
    For i = 2 To 32
        If i <> 19 Then
            ret = sio_open(i)
            If ret = SIO_OK Then
                sio_close (i)
                comPort(UBound(comPort())) = i
                ReDim Preserve comPort(UBound(comPort()) + 1)
            End If
        End If
    Next i
    ReDim Preserve comPort(UBound(comPort()) - 1)
End Function

'**********************************************************************
' ��ȡ���������е�iccid
'**********************************************************************
Public Function GetAllIccid(Optional blIsSkipSendingSMS As Boolean) As String
    Dim cIdx As Integer
    GetAllIccid = "''"
    If IsComEmpty = False Then
        If blIsSkipSendingSMS = True Then
            For cIdx = 0 To UBound(Com())
                If Com(cIdx).Iccid <> "" And Com(cIdx).iSmsId = 0 Then
                    GetAllIccid = GetAllIccid & " , '" & Com(cIdx).Iccid & "'"
                End If
            Next cIdx
        Else
            For cIdx = 0 To UBound(Com())
                If Com(cIdx).Iccid <> "" Then
                    GetAllIccid = GetAllIccid & " , '" & Com(cIdx).Iccid & "'"
                End If
            Next cIdx
        End If
    End If
End Function

Public Function IsComEmpty() As Boolean
    On Error GoTo Err
    If UBound(Com()) > -1 Then
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
