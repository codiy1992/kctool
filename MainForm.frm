VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������� -- by CODIY"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11340
   ForeColor       =   &H80000008&
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   11340
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer TimerKcCron 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2280
      Top             =   8640
   End
   Begin VB.TextBox TextDsp 
      Height          =   3015
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4800
      Width           =   4340
   End
   Begin VB.TextBox TextLog 
      Height          =   3015
      Left            =   7400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4800
      Width           =   3875
   End
   Begin VB.Timer TimerKcTask 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1800
      Top             =   8640
   End
   Begin VB.Timer TimerKcRead 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2760
      Top             =   8640
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   15
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   14
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   13
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   12
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   3000
      Left            =   4680
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   3000
      Left            =   3240
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   3000
      Left            =   1920
      Top             =   9840
   End
   Begin VB.Timer TimerComCron 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   3000
      Left            =   600
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   15
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   15
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   14
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   14
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   13
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   13
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   12
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   12
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   11
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   10
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   9
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   8
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   7
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   6
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   200
      Left            =   5040
      Top             =   9840
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   300
      Left            =   4320
      Top             =   9840
   End
   Begin VB.Timer TimerComRead 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   200
      Left            =   5040
      Top             =   9840
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
   Begin VB.Timer TimerInit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   8640
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
   Begin VB.Frame Frame2 
      Caption         =   "������ť"
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   2775
      Begin VB.CommandButton Command1 
         Caption         =   "�����ѿ�"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton CommandSwitchIndex 
         Caption         =   "ָ���п�"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CommandDspInfo 
         Caption         =   "��ʾ��Ϣ"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton CommandGetCCFC 
         Caption         =   "��ȡ��ת"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton CommandSwitch 
         Caption         =   "�Զ��п�"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton CommandClnTextLog 
         Caption         =   "���AT��¼"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton CommandClnTextDsp 
         Caption         =   "��ջ���"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton CommandKcSwitch 
         Caption         =   "����һ��"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CommandKcReSet 
         Caption         =   "��������"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton CommandCloseAll 
         Caption         =   "�ر�����"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CommandStartAll 
         Caption         =   "�����豸"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�����б�"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11150
      Begin MSComctlLib.ListView ListView 
         Height          =   3945
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   10885
         _ExtentX        =   19209
         _ExtentY        =   6959
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
   Begin VB.Label Label1 
      Caption         =   "ATָ���¼:"
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "���ݻ��Դ���:"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Menu MenuList 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu MENU_OPEN 
         Caption         =   "����������"
      End
      Begin VB.Menu MENU_CLOSE 
         Caption         =   "�رձ�����"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_STOP 
         Caption         =   "��ͣ������"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_START 
         Caption         =   "�ָ�������"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_INFO 
         Caption         =   "�鿴������Ϣ"
      End
      Begin VB.Menu MENU_CCFC 
         Caption         =   "���ת����"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_DEBUG 
         Caption         =   "�鿴������"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' �����쳣�޸������ݿ������޸�

Private Sub Command1_Click()
    If g_blPrint = True Then
        g_blPrint = False
        Command1.Caption = "�����ѹ�"
    Else
        g_blPrint = True
        Command1.Caption = "�����ѿ�"
    End If
End Sub

'1����ʼ��ʱ���жϴ����ǿ��ػ���simͨ��
'2���п�ʱ��У�飬�����п����ϴ��Ƿ�Ϊͬһ�ſ�(�տ������һ��)
'3��������simͨ�������Ϣ
'4���࿨��֧��
'5���Կ����͵�֧��
Private Sub Form_Load()
    ' ��ʼ������
    g_iDebugIndex = -1
    g_blPrint = True
    ' ��ʼ���б���ͼ�ؼ�
    With ListView 'ListView��ʼ��
         .View = 3 ' �б���ʾ��ʽ
         .ColumnHeaders.Add , , "���", 600         ' 0
         .ColumnHeaders.Add , , "����", 800         ' 1
         .ColumnHeaders.Add , , "״̬", 3200        ' 2
         .ColumnHeaders.Add , , "ICCID��", 2300     ' 3
         .ColumnHeaders.Add , , "�ֻ�����", 1400    ' 4
         .ColumnHeaders.Add , , "IMEI��", 1800      ' 5
         '.ColumnHeaders.Add , , "�ź�", 1700         ' 6
         .ColumnHeaders.Add , , "ռ��", 800, lvwColumnCenter   ' 7
    End With
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
                For cIdx = 0 To UBound(com())
                    ' ��ֹͣ��������������ִ�ж���
                    com(cIdx).task.Push ("--CLOSE--")
                    TimerComTask(cIdx).Enabled = True
                Next cIdx
            End If
            Cancel = True
        End If
    End If
End Sub

' ====================================  ������ť  ====================================================
Private Sub CommandStartAll_Click()
    Dim cIdx As Integer
    If IsComEmpty = False Then
        For cIdx = 0 To UBound(com())
            OpenCom (cIdx)
        Next cIdx
        kc.blCanSwitch = True
    ElseIf TimerInit.Enabled = False Then
        ' ������ߴ���
        CommandStartAll.Caption = "������..."
        TimerInit.Enabled = True
    End If
End Sub

Private Sub CommandCloseAll_Click()
    ' �رմ���
    If IsComEmpty = False Then
        For cIdx = 0 To UBound(com())
           ' ��ֹͣ��������������ִ�ж���
            If com(cIdx).task.Top <> "--CLOSE--" Then
                com(cIdx).task.Push ("--CLOSE--")
                TimerComTask(cIdx).Enabled = True
            End If
        Next cIdx
    End If
End Sub

Private Sub CommandClnTextDsp_Click()
    TextDsp.Text = ""
End Sub

Private Sub CommandClnTextLog_Click()
    TextLog.Text = ""
End Sub

Private Sub CommandDspInfo_Click()
    If IsComEmpty = False Then
        For cIdx = 0 To UBound(com())
            If com(cIdx).blIsOpen = False Then
                TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                           "     COM" & com(cIdx).comPort & ": δ����" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
            Else
                TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                           "     COM" & com(cIdx).comPort & ": �ѿ���" & vbCrLf & _
                           "  �ֻ�����: " & com(cIdx).Mobile & vbCrLf & _
                           "   ICCID��: " & com(cIdx).Iccid & vbCrLf & _
                           "  IMEI����: " & com(cIdx).Imei & vbCrLf & _
                           "  IMSI����: " & com(cIdx).Imsi & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
            End If
        Next cIdx
    End If
    TextDsp.SelStart = Len(TextDsp.Text)
End Sub

Private Sub CommandGetCCFC_Click()
    If IsComEmpty = False Then
        For cIdx = 0 To UBound(com())
            com(cIdx).task.Push ("AT+CCFC=0,2" & vbCrLf)
            TimerComTask(cIdx).Enabled = True
        Next cIdx
    End If
End Sub

Private Sub CommandKcReSet_Click()
    If IsComEmpty = False Then
        If DB.IsUsing(GetAllIccid) = False Then
            kc.task.Push ("--KC-RESET--")
            TimerKcTask.Enabled = True
        End If
    End If
End Sub

Private Sub CommandKcSwitch_Click()
    If IsComEmpty = False Then
        If DB.IsUsing(GetAllIccid) = False Then
            kc.task.Push ("--KC-NEXT-A--")
            TimerKcTask.Enabled = True
        End If
    End If
End Sub

Private Sub CommandSwitch_Click()
    If IsComEmpty = True Then
        Exit Sub
    End If
    If kc.blCanSwitch = True Then
        kc.blCanSwitch = False
    Else
        kc.blCanSwitch = True
    End If
End Sub

Private Sub CommandSwitchIndex_Click()
    Dim simIndex
    simIndex = InputBox("sim��λ��(1-16)", "�����л�λ��", 1)
    simIndex = Val(simIndex)
    If IsComEmpty = True Then
        Exit Sub
    End If
    If simIndex < 1 Or simIndex > 16 Then
        MsgBox "������1-16֮�������"
        Exit Sub
    End If
    If DB.IsUsing(GetAllIccid) = False Then
        kc.task.Push ("--KC-INDEX-A-" & Format(simIndex, "00"))
        TimerKcTask.Enabled = True
    End If
End Sub

Private Sub TimerKcCron_Timer()
    Dim I, j As Integer
    Dim cIdx As Integer
    Dim cntA As Integer, cntB As Integer
    Dim strIccid As String
    Dim execArr
    kc.iTmrCnt = kc.iTmrCnt + 1
    If kc.iTmrCnt Mod 60 = 0 Then
            kc.iTmrCnt = 0
    End If
    
    Frame1.Caption = "���� - �Զ��п�[" & kc.iTmrCnt * 3 & "/180]:" & kc.blCanSwitch & "    (" & Now() & ")"
    If IsComEmpty = True Then
        Exit Sub
    End If
    
    ' �Զ��п�����
    If kc.blCanSwitch = True Then
        If kc.iTmrCnt = 0 Then
            If DB.IsUsing(GetAllIccid) = False Then
                ' �п�
                If InStr(kc.task.Top, "--KC-INDEX-AS-") > 0 Then
                Else
                    kc.task.Push ("--KC-INDEX-AS-" & Format((kc.rowIndex Mod 10) + 1, "00"))
                    TimerKcTask.Enabled = True
                End If
            End If
        End If
    End If
    
    If kc.iTmrCnt Mod 2 = 0 Then
        Exit Sub
    End If
    ' 6�����ó�ʱsim��ʹ��״̬
    DB.setUseTimeOut
    ' ִ���п�����
    If kc.blIsSwitching = False Then
        execArr = DB.GetSwitchCard(strIccid)
        If execArr <> "" Then
            Dim simIndex As Integer
            simIndex = Val(Mid(execArr, 5, 2))
            If simIndex = kc.rowIndex Then
                DB.SetCardCanUse (strIccid)
            Else
                kc.blIsSwitching = True
                kc.task.Push ("--KC-INDEX-ASU-" & Format(simIndex, "00"))
                TimerKcTask.Enabled = True
            End If
        End If
    End If
    
    ' ִ�а���������ת����
    execArr = DB.NotExecBind(GetAllIccid)
    If LCase(TypeName(execArr)) = "string()" Then
        cntA = UBound(execArr) + 1
        For cIdx = 0 To UBound(com())
            For I = 0 To UBound(execArr)
                If execArr(I, 0) = com(cIdx).Iccid Then
                    com(cIdx).iBindId = Val(execArr(I, 1))
                    If execArr(I, 2) <> "" Then
                        com(cIdx).bindMobile (execArr(I, 2))
                    Else
                        com(cIdx).unBindMobile
                    End If
                End If
                TimerComTask(cIdx).Enabled = True
            Next I
        Next cIdx
    End If
    
    ' ִ�з��Ͷ�������
    execArr = DB.NotSendedSMS(GetAllIccid(True))
    If LCase(TypeName(execArr)) = "string()" Then
        cntB = UBound(execArr) + 1
        For cIdx = 0 To UBound(com())
            For I = 0 To UBound(execArr)
                If com(cIdx).Iccid = execArr(I, 0) Then
                    com(cIdx).iSmsId = Val(execArr(I, 1))
                    com(cIdx).sendSMS execArr(I, 2), execArr(I, 3)
                    TimerComTask(cIdx).Enabled = True
                End If
            Next I
        Next cIdx
    End If
End Sub
'==============================================  ���ؿ���ģ�� ==================================================================
Private Sub TimerKcTask_Timer()
    Dim strAT As String
    Dim blCanSwitch As Boolean
    Dim cIdx As Long
    blCanSwitch = True
    
    If kc.blIsATExecing = True Then
        Exit Sub
    End If
    
    strAT = kc.task.Top
    If strAT <> Empty Then
        If kc.blIsOpen = True Then
            Select Case strAT
                Case "--KC-NEXT-A--" ' �л�����һ�ſ�(ȫ��ͨ��)
                    If IsComEmpty = False Then
                        For cIdx = 0 To UBound(com())
                            If com(cIdx).blIsOpen = True And com(cIdx).blIsNormal = True Then
                                blCanSwitch = False
                                com(cIdx).task.Push ("--CLOSE--")
                                TimerComTask(cIdx).Enabled = True
                            End If
                        Next cIdx
                        If blCanSwitch = False Then
                            Exit Sub
                        End If
                    End If
                    kc.blIsATExecing = True
                    kc.WriteData ("AT+NEXT11" & vbCrLf)
                Case "--KC-NEXT-AS--" ' �л�����һ�ſ�(ȫ��ͨ��)
                    If IsComEmpty = False Then
                        For cIdx = 0 To UBound(com())
                            If com(cIdx).blIsOpen = True And com(cIdx).blIsNormal = True Then
                                blCanSwitch = False
                                com(cIdx).task.Push ("--CLOSE--")
                                TimerComTask(cIdx).Enabled = True
                            End If
                        Next cIdx
                        If blCanSwitch = False Then
                            Exit Sub
                        End If
                    End If
                    kc.blIsATExecing = True
                    kc.WriteData ("AT+NEXT11" & vbCrLf)
                Case "--KC-RESET--" ' ��������(ȫ��ͨ��)
                    If IsComEmpty = False Then
                        For cIdx = 0 To UBound(com())
                            If com(cIdx).blIsOpen = True And com(cIdx).blIsNormal = True Then
                                blCanSwitch = False
                            End If
                            If com(cIdx).task.Top <> "--CLOSE--" Then
                                com(cIdx).task.Push ("--CLOSE--")
                            End If
                            TimerComTask(cIdx).Enabled = True
                        Next cIdx
                        If blCanSwitch = False Then
                            Exit Sub
                        End If
                    End If
                    kc.blIsATExecing = True
                    kc.WriteData ("AT+NEXT00" & vbCrLf)
                Case Else
                    If InStr(strAT, "--KC-INDEX-A") > 0 Then ' ָ���л���ĳ�ſ�(ȫ��ͨ��)
                        If IsComEmpty = False Then
                            For cIdx = 0 To UBound(com())
                                If com(cIdx).blIsOpen = True And com(cIdx).blIsNormal = True Then
                                    blCanSwitch = False
                                    If com(cIdx).task.Top <> "--CLOSE--" Then
                                        com(cIdx).task.Push ("--CLOSE--")
                                    End If
                                    TimerComTask(cIdx).Enabled = True
                                End If
                            Next cIdx
                            If blCanSwitch = False Then
                                Exit Sub
                            End If
                        End If
                        kc.blIsATExecing = True
                        kc.WriteData ("AT+SWIT00-00" & Right(strAT, 2) & vbCrLf)
                    Else
                        kc.blIsATExecing = True
                        kc.WriteData (strAT)
                    End If
            End Select
        End If
    Else
       TimerKcTask.Enabled = False
    End If
End Sub

Private Sub TimerKcRead_Timer()
    Dim strTmp As String
    Dim strAT As String
    Dim cIdx As Integer
    Dim simIndex As Integer
    
    strTmp = kc.ReadData
    
    If kc.iWaitCnt >= 10 Then
        kc.iWaitCnt = 0
        kc.blIsATExecing = False
    End If
    If strTmp = "" Or kc.blIsATExecing = False Then 'Com(cIdx).blIsATExecing = False Or
        Exit Sub
    End If
    
    strAT = kc.task.Top
    strTmp = kc.GetData(strTmp)
    
    If InStr(strTmp, "OK") > 0 Then
        Select Case strAT
            Case "--KC-NEXT-A--"
                kc.iNullCnt = 0
                If IsComEmpty = False Then
                    For cIdx = 0 To UBound(com())
                        com(cIdx).simIndex = (com(cIdx).simIndex Mod 16) + 1
                    Next cIdx
                    If g_blPrint = True Then
                    TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                           "  �����п�: �ɹ� ��������" & com(0).simIndex & "��" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                End If
                
            Case "--KC-NEXT-AS--"
                kc.iNullCnt = 0
                If IsComEmpty = False Then
                    For cIdx = 0 To UBound(com())
                        com(cIdx).simIndex = (com(cIdx).simIndex Mod 16) + 1
                        OpenCom (cIdx)
                    Next cIdx
                    If g_blPrint = True Then
                    TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                           "  �����п�: �ɹ� ��������" & com(0).simIndex & "��" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                End If
                
            Case "AT+NEXT11" & vbCrLf ' ����ȫ���л��ɹ�
                kc.iNullCnt = 0
                If IsComEmpty = False Then
                    For cIdx = 0 To UBound(com())
                        com(cIdx).simIndex = (com(cIdx).simIndex Mod 16) + 1
                        OpenCom (cIdx)
                    Next cIdx
                End If
                
            Case "--KC-RESET--"
                kc.iNullCnt = 0
                If IsComEmpty = False Then
                    For cIdx = 0 To UBound(com())
                        com(cIdx).simIndex = 1
                    Next cIdx
                    kc.rowIndex = 1
                    If g_blPrint = True Then
                    TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                           "  ��������: �ɹ�" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                    TimerKcCron.Enabled = True
                End If
                
            Case "AT+NEXT00" & vbCrLf ' �������óɹ�
                kc.iNullCnt = 0
                If IsComEmpty = False Then
                    For cIdx = 0 To UBound(com())
                        com(cIdx).simIndex = 1
                        OpenCom (cIdx)
                    Next cIdx
                    
                End If
                kc.rowIndex = 1
            Case "AT+CWSIM" & vbCr   ' �뿨�����ֳɹ�
                'Kc.task.Push ("AT+NEXT00" & vbCrLf)
                kc.task.Push ("--KC-RESET--")
                TimerKcTask.Enabled = True
                
            Case Else
                If InStr(strAT, "AT+NEXT11-") > 0 Then '"AT+NEXT11-06" & vbCrLf
                    cIdx = Int(Left(Right(strAT, 4), 2)) - 1
                    com(cIdx).simIndex = com(cIdx).simIndex + 1
                    OpenCom (cIdx)
                    
                End If
                
                ' �л���ָ��λ��(ȫ��ͨ��)
                If InStr(strAT, "--KC-INDEX-A-") > 0 Then
                    kc.iNullCnt = 0
                    If IsComEmpty = False Then
                        For cIdx = 0 To UBound(com())
                            com(cIdx).simIndex = Val(Right(strAT, 2))
                        Next cIdx
                    End If
                    kc.rowIndex = Val(Right(strAT, 2))
                    kc.blIsSwitching = False
                    If g_blPrint = True Then
                    TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                           "  �����п�: �ɹ� ��������" & Val(Right(strAT, 2)) & "��" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                    
                End If
                ' �л���ָ��λ��(ȫ��ͨ��),
                If InStr(strAT, "--KC-INDEX-AS-") > 0 Then
                    kc.iNullCnt = 0
                    If IsComEmpty = False Then
                        For cIdx = 0 To UBound(com())
                            com(cIdx).simIndex = Val(Right(strAT, 2))
                            OpenCom (cIdx)
                        Next cIdx
                    End If
                    kc.rowIndex = Val(Right(strAT, 2))
                    kc.blIsSwitching = False
                    If g_blPrint = True Then
                    TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                           "  �����п�: �ɹ� ��������" & Val(Right(strAT, 2)) & "��" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                End If
                ' �л���ָ��λ��(ȫ��ͨ��)
                If InStr(strAT, "--KC-INDEX-ASU-") > 0 Then
                    kc.iNullCnt = 0
                    If IsComEmpty = False Then
                        For cIdx = 0 To UBound(com())
                            com(cIdx).simIndex = Val(Right(strAT, 2))
                            OpenCom (cIdx)
                        Next cIdx
                    End If
                    kc.rowIndex = Val(Right(strAT, 2))
                    kc.blIsSwitching = False
                    If g_blPrint = True Then
                    TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                           "  �����п�: �ɹ� ��������" & Val(Right(strAT, 2)) & "��" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                End If
                
                If InStr(strAT, "AT+SWIT00-") > 0 Then
                    kc.iNullCnt = 0
                    If IsComEmpty = False Then
                        For cIdx = 0 To UBound(com())
                            com(cIdx).simIndex = Val(Left(Right(strAT, 4), 2))
                            OpenCom (cIdx)
                        Next cIdx
                    End If
                End If
                
                ' �л���ָ��λ��(ĳͨ��)
                If InStr(strAT, "AT+SWIT") > 0 And Mid(strAT, 8, 2) <> "00" Then
                    kc.iNullCnt = 0
                    cIdx = Val(Mid(strAT, 8, 2)) - 1
                    com(cIdx).simIndex = Val(Left(Right(strAT, 4), 2))
                    'OpenCom (cIdx)
                End If
                
        End Select
        strAT = kc.task.Pop
        kc.blIsATExecing = False
    Else
        kc.blIsATExecing = False
    End If
    
End Sub

Private Sub TimerInit_Timer()
    Dim I As Integer
    Dim comport_use() As String
    
    ' ����
    Set kc = New kc
    kc.index = 1
    kc.comPort = 36
    kc.OpenPort
    If kc.blIsOpen = True Then
        TimerKcTask.Enabled = True
        TimerKcRead.Enabled = True
        kc.task.Push ("AT+CWSIM" & vbCr)
    End If
    
    Call comportScan(36, comport_use())
    If comport_use(0) <> "" Then
        ReDim com(UBound(comport_use()))
        For I = 0 To UBound(comport_use())
            Set LV = ListView.ListItems.Add(I + 1, , I + 1)
            LV.SubItems(1) = "COM" & comport_use(I)
            Set com(I) = New com
            com(I).comPort = comport_use(I)
            com(I).simIndex = 1
            com(I).kc = kc.index
            LV.SubItems(2) = "δ����"
        Next I
        CommandStartAll.Caption = "��������"
    Else
        CommandStartAll.Caption = "�޿����豸"
    End If
    TimerInit.Enabled = False
End Sub

'=========================================================================================================================
Private Sub TimerComCron_Timer(cIdx As Integer)
    
    com(cIdx).iTmrCnt = com(cIdx).iTmrCnt + 1
    If com(cIdx).iTmrCnt Mod 60 = 0 Then
         com(cIdx).iTmrCnt = 0
    End If
    
    If com(cIdx).iTmrCnt Mod 2 = 1 Then
        Dim strTime As String
        If DB.IsUsing("'" & com(cIdx).Iccid & "'", strTime) = True Then
            strTime = 300 - (DateDiff("s", "1970-1-1 8:0:0", DateAdd("s", 0, Now())) - Val(strTime))
            ListView.ListItems(cIdx + 1).SubItems(6) = Format(strTime, "000")
        Else
            ListView.ListItems(cIdx + 1).SubItems(6) = "---"
        End If
    End If
    
    If com(cIdx).blIsShowStat = True Then
        Dim exec As Integer, pick As Integer
        If com(cIdx).blIsATExecing = True Then
            exec = 1
        End If
        If com(cIdx).blIsPickSms = True Then
            pick = 1
        End If
        ListView.ListItems(cIdx + 1).SubItems(2) = "[" & Format(com(cIdx).kc, "000") & "-" & Format(com(cIdx).simIndex, "00") & "-" & Format(cIdx + 1, "00") & "]" & "-" & Format(com(cIdx).task.wIndex, "00") & "-" & Format(com(cIdx).task.rIndex, "00") & "-" & _
                                                        exec & "-" & Format(com(cIdx).iWaitCnt, "00") & "-" & Format(kc.iNullCnt, "00")
    End If
    If com(cIdx).blIsCheck = True And com(cIdx).blIsNormal = True Then
    
        ' ������ڷ����Ż������ú�ת,����������������µ���������
        If com(cIdx).iSmsId > 0 Or com(cIdx).iBindId > 0 Then
            ListView.ListItems(cIdx + 1).SubItems(2) = "[" & com(cIdx).iSmsId & "|" & com(cIdx).iBindId & "]"
            Exit Sub
        End If
        
        'com(cIdx).task.Push ("AT+CSQ" & vbCrLf)
        
        com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
        
        If com(cIdx).SP = "" Then
            com(cIdx).task.Push ("AT+COPS?" & vbCrLf)
        End If
        
        If com(cIdx).Imei = "" Then
            com(cIdx).task.Push ("AT+CGSN" & vbCrLf)      ' Imei
        End If
        
        'If com(cIdx).Imsi = "" Then
        
        'End If
        
        If com(cIdx).blIsPickSms = True Then
            com(cIdx).task.Push ("AT+CMGL=""ALL""" & vbCrLf)
            com(cIdx).blIsPickSms = False ' ��������ǰ�������ٷ����µ�ȡ��������
        End If
        
        TimerComTask(cIdx).Enabled = True
    End If
    
    
End Sub

Private Sub TimerComTask_Timer(cIdx As Integer)
    Dim strAT As String
    strAT = com(cIdx).task.Top
    If strAT = Empty Then
        TimerComTask(cIdx).Enabled = False
        Exit Sub
    End If

    If UCase(Left(strAT, 2)) = "AT" And com(cIdx).blIsATExecing = True Then 'Or strAT = "--SWITCH--")
        Exit Sub
    End If
    'If com(cIdx).blIsNormal = True Then
        Select Case strAT
            Case "--TAIL--"
                Dim tail(0) As Byte
                strAT = com(cIdx).task.Pop
                tail(0) = &H1A
                com(cIdx).WriteData (tail)
            Case "--STOP--"
                strAT = com(cIdx).task.Pop
                TimerComTask(cIdx).Enabled = False
                ListView.ListItems(cIdx + 1).SubItems(2) = "����ͣ"
            Case "--CLOSE--"
                TimerComRead(cIdx).Enabled = True
                If com(cIdx).blIsOpen = True Or com(cIdx).Iccid <> "" Then ' ���ŵĻ������ֻ��ŵ�
                    com(cIdx).blIsCheck = False
                    com(cIdx).WriteData ("AT+CFUN=0" & vbCrLf)
                Else    ' ��sim��
                     strAT = com(cIdx).task.Pop
                     CloseCom (cIdx)
                End If
            Case "--SWITCH--"
                TimerComRead(cIdx).Enabled = True
                If com(cIdx).blIsOpen = True Or com(cIdx).Iccid <> "" Then ' ���ŵĻ������ֻ��ŵ�
                    com(cIdx).blIsCheck = False
                    com(cIdx).blIsATExecing = True
                    com(cIdx).WriteData ("AT+CFUN=0" & vbCrLf)
                Else    ' ��sim��
                     strAT = com(cIdx).task.Pop
                     CloseCom (cIdx)
                     kc.task.Push ("AT+NEXT11-" & Format(cIdx + 1, "00") & vbCrLf)
                     TimerKcTask.Enabled = True
                End If
            Case Else
                strAT = com(cIdx).task.Pop
                If UCase(Left(strAT, 2)) = "AT" Then
                    com(cIdx).blIsATExecing = True
                End If
                com(cIdx).WriteData (strAT)
        End Select
    'End If
End Sub


Private Sub TimerComRead_Timer(cIdx As Integer)
    Dim strAtData As String
    Dim tmpBuf() As Byte, strTmp As String
    Dim strOut As String
    Dim strAT As String
    Dim smsArr() As SMSDef
    strTmp = com(cIdx).ReadData
    If com(cIdx).iWaitCnt >= 40 Then
        If com(cIdx).iWaitCnt >= 98 And (com(cIdx).blIsCheck = True Or com(cIdx).blIsPickSms = True) Then
            com(cIdx).iWaitCnt = 0
            com(cIdx).blIsNormal = False
            ListView.ListItems(cIdx + 1).SubItems(2) = "��������Ӧ"
        End If
        If com(cIdx).iWaitCnt Mod 40 = 0 Then
            com(cIdx).blIsATExecing = False
        End If
    End If
    If strTmp = "" Then
        Exit Sub
    End If
    If com(cIdx).blIsOpen = True Then
        If cIdx = g_iDebugIndex Then
            TextLog.Text = TextLog.Text & strTmp ' & vbCrLf & "------------------" & vbCrLf
            TextLog.SelStart = Len(TextLog.Text)
        End If
        strAtData = com(cIdx).GetData(strTmp)
        If strAtData = Empty Then
            Exit Sub
        End If
        
        strAT = com(cIdx).AnalysisData(strAtData, strOut)
        
        If strAT = "" Then 'And strAT <> vbCr And strAtData <> vbCrLf And Not IsEmpty(strAT)
            Exit Sub
        End If
        
        com(cIdx).blIsATExecing = False
        Select Case strAT
            Case "AT+CSQ"
                ListView.ListItems(cIdx + 1).SubItems(6) = strOut
            Case "AT+CSCA?"
            Case "AT+COPS?"
                com(cIdx).SP = strOut
                'ListView.ListItems(cIdx + 1).SubItems(6) = strOut
            Case "AT+CGSN"
                com(cIdx).Imei = strOut
                ListView.ListItems(cIdx + 1).SubItems(5) = strOut
            Case "AT+CIMI"
                com(cIdx).Imsi = strOut
            Case "AT+CCFC=0,2"
                If com(cIdx).iQccfcCnt < 10 And strOut = "" Then
                    com(cIdx).task.Push ("AT+CCFC=0,2" & vbCrLf)
                Else
                    ' �����ݿ⣬��ȡ���ݿ���ֻ�
                    
                    ' ��������Դ���
                    If strOut <> "" Then
                        TextDsp.Text = TextDsp.Text & Now() & vbCrLf & _
                                            "�ֻ���: " & com(cIdx).Mobile & "(" & com(cIdx).Iccid & ")" & vbCrLf & _
                                            "    ���õĺ�ת����Ϊ ��" & strOut & "��" & vbCrLf & _
                                            "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    Else
                        TextDsp.Text = TextDsp.Text & Now() & vbCrLf & _
                                        "�ֻ���: " & com(cIdx).Mobile & "(" & com(cIdx).Iccid & ")" & vbCrLf & _
                                        "    û�鵽��û���ú�ת����" & vbCrLf & _
                                        "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                End If
            Case "-AT-BIND-MOBILE-OK-"
                DB.SetBinded com(cIdx).Iccid, com(cIdx).iBindId
                com(cIdx).iBindId = 0
            Case "-AT-BIND-MOBILE-FAILED-"
                DB.SetNotBind com(cIdx).Iccid, com(cIdx).iBindId
                com(cIdx).iBindId = 0
            Case "-AT-UNBIND-MOBILE-OK-"
                DB.SetBinded com(cIdx).Iccid, com(cIdx).iBindId
                com(cIdx).iBindId = 0
                com(cIdx).bMobile = ""
            Case "-AT-UNBIND-MOBILE-FAILED-"
                DB.SetNotBind com(cIdx).Iccid, com(cIdx).iBindId
                com(cIdx).iBindId = 0
            Case "AT+CMGL"
                If InStr(strOut, "ERROR") Then
                    com(cIdx).task.Push ("AT+CMGF=1" & vbCrLf)    ' ���ö��Ÿ�ʽ
                    com(cIdx).task.Push ("AT+CSCS=""UCS2""" & vbCrLf)    ' ���ö��ű���
                End If
                strOut = PickAllSMS(strOut, smsArr)
                If UBound(smsArr) > 0 Then
                    For n = 1 To UBound(smsArr)
                         'If cIdx = g_iDebugIndex Then
                         If g_blPrint = True Then
                             TextDsp.Text = TextDsp.Text & vbCrLf & smsArr(n).SmsIndex & vbTab _
                                            & smsArr(n).DateTime & vbTab _
                                            & smsArr(n).SourceNo & vbCrLf _
                                            & smsArr(n).SmsMain & vbCrLf _
                                            & "-------------------------------------" & vbCrLf
                             TextDsp.SelStart = Len(TextDsp.Text)
                         End If
                         'End If
                         If com(cIdx).Iccid <> "" Then
                            DB.SaveSMS com(cIdx).Iccid, smsArr(n).SourceNo, smsArr(n).SmsMain, smsArr(n).DateTime, com(cIdx).Mobile
                            com(cIdx).task.Push ("AT+CMGD=" & smsArr(n).SmsIndex & vbCrLf)
                         End If
                    Next n
                End If
                com(cIdx).blIsPickSms = True    ' ������ɺ�,���������µ�ȡ��������
            Case "+CMTI:" ' �յ��¶���
            Case "-AT-SMS-SEND-OK-" ' ���ŷ��ͳɹ�
                DB.SetSMSSended com(cIdx).Iccid, com(cIdx).iSmsId
                com(cIdx).iSmsId = 0
            Case "-AT-SMS-SEND-FAILED-" '���ŷ���ʧ��
                DB.SetSMSNotSend com(cIdx).Iccid, com(cIdx).iSmsId
                com(cIdx).iSmsId = 0
            Case "-AT-INIT-OK-"    ' ��ʼ��������һ���ɹ���
                com(cIdx).task.Push ("ATE1" & vbCrLf)         ' ��������
                'com(cIdx).task.Push ("AT+CGSN" & vbCrLf)
                com(cIdx).task.Push ("AT+CGSN" & vbCrLf)
                com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
                com(cIdx).task.Push ("AT+CNMI=2,1" & vbCrLf)  '
                com(cIdx).task.Push ("AT+CSCS=""UCS2""" & vbCrLf)    ' ���ö��ű���
                com(cIdx).task.Push ("AT+CCID" & vbCrLf)      ' ��ѯICCID��
                ListView.ListItems(cIdx + 1).SubItems(2) = "��ʼ����(1)..."
                ' ����������������ʱ��
                TimerComTask(cIdx).Interval = 100
                TimerComTask(cIdx).Enabled = True
            Case "AT+CCID"         ' ��ʼ������������ɹ���
                com(cIdx).blIsNormal = True
                If strOut = "-RETRY-" Then
                    If com(cIdx).iTryCnt = 2 Then
                        TimerComTask(cIdx).Interval = 600
                    End If
                    If com(cIdx).iTryCnt = 3 Then
                        TimerComTask(cIdx).Interval = 900
                    End If
                    If com(cIdx).iTryCnt = 5 Then
                        TimerComTask(cIdx).Interval = 1000
                    End If
                    com(cIdx).task.Push ("AT+CCID" & vbCrLf)      ' ��ѯICCID��
                    TimerComTask(cIdx).Enabled = True
                ElseIf strOut = "-NO-CCID-" Then
                    ListView.ListItems(cIdx + 1).SubItems(2) = "������SIM��"
                    com(cIdx).Iccid = ""
                    com(cIdx).blIsShowStat = False
                    TimerComRead(cIdx).Enabled = False
                    kc.iNullCnt = kc.iNullCnt + 1
                    If kc.iNullCnt = 16 Then
                        kc.iNullCnt = 0
                        If kc.blCanSwitch = True Then
                            If DB.IsUsing(GetAllIccid) = False Then
                                ' �п�
                                If InStr(kc.task.Top, "--KC-INDEX-AS-") > 0 Then
                                Else
                                    kc.task.Push ("--KC-INDEX-AS-" & Format((kc.rowIndex Mod 10) + 1, "00"))
                                    TimerKcTask.Enabled = True
                                End If
                            End If
                        End If
                    End If
                Else
                    com(cIdx).Iccid = Left(strOut, 19)
                    ListView.ListItems(cIdx + 1).SubItems(3) = Left(strOut, 19)
                    ' �����ݿ�ע�᱾SIM��������ȡ�ֻ�����
                    com(cIdx).Mobile = DB.RegistCard(com(cIdx).Iccid, _
                                                    Format(com(cIdx).kc, "000") & "-" & Format(com(cIdx).simIndex, "00") & "-" & Format(cIdx + 1, "00"), _
                                                    com(cIdx).Imei, com(cIdx).Imsi)
                    If com(cIdx).Mobile = "" Then
                        com(cIdx).blIsShowStat = False
                        ListView.ListItems(cIdx + 1).SubItems(2) = "���������ֻ���"
                        TimerComRead(cIdx).Enabled = False
                    Else
                        ListView.ListItems(cIdx + 1).SubItems(2) = "��������"
                        ListView.ListItems(cIdx + 1).SubItems(4) = com(cIdx).Mobile
                        com(cIdx).blIsCheck = True
                        'Com(cIdx).blIsSwitch = True
                    End If
                End If
                If strOut <> "-RETRY-" Then
                    com(cIdx).OldIccid = com(cIdx).Iccid
                    TimerComCron(cIdx).Enabled = True
                    TimerComTask(cIdx).Interval = 400
                    TimerComTask(cIdx).Enabled = True
                End If
            Case "-AT-EXIT-OK-"  '�ر�simָ�� OK
                 If com(cIdx).task.Top = "--SWITCH--" Then
                    strAT = com(cIdx).task.Pop
                    CloseCom (cIdx)
                    kc.task.Push ("AT+NEXT11-" & Format(cIdx + 1, "00") & vbCrLf)
                    TimerKcTask.Enabled = True
                 Else
                    CloseCom (cIdx)
                 End If
            Case Else
                com(cIdx).blIsATExecing = False
        End Select
    End If
End Sub
Private Sub ListView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim index As Integer
    '��������Ҽ�
    If Button = vbRightButton Then
        If ListView.ListItems.Count > 0 Then
            index = ListView.SelectedItem.index
            If com(index - 1).blIsOpen = True Then
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
                
                MENU_CCFC.Visible = True
                MENU_CCFC.Enabled = True
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
                
                MENU_CCFC.Visible = False
                MENU_CCFC.Enabled = False
            End If
            PopupMenu MenuList
        End If
    End If
End Sub
Private Sub MENU_OPEN_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.index - 1
    OpenCom (cIdx)
End Sub
Private Sub MENU_CLOSE_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.index - 1
    ' �޸����ݿ��ֶ�״̬
    DB.setCardClose ("'" & com(cIdx).Iccid & "'")
    ' ��ֹͣ��������������ִ�ж���
    If com(cIdx).task.Top <> "--CLOSE--" Then
        com(cIdx).task.Push ("--CLOSE--")
        TimerComTask(cIdx).Enabled = True
    End If
    com(cIdx).blIsCheck = False
End Sub
Private Sub MENU_STOP_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.index - 1
    com(cIdx).task.Push ("--STOP--") ' ��ֹͣ��������������ִ�ж���
    com(cIdx).blIsCheck = False
    TimerComTask(cIdx).Enabled = True
    TimerComRead(cIdx).Enabled = False
    ' �˵�
    MENU_START.Visible = True
    MENU_START.Enabled = True
    MENU_STOP.Visible = False
    MENU_STOP.Enabled = False
End Sub
Private Sub MENU_START_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.index - 1
    com(cIdx).blIsCheck = True
    TimerComTask(cIdx).Enabled = True
    TimerComRead(cIdx).Enabled = True
    ListView.ListItems(cIdx + 1).SubItems(2) = "��������"
    ' �˵�
    MENU_START.Visible = False
    MENU_START.Enabled = False
    MENU_STOP.Visible = True
    MENU_STOP.Enabled = True
End Sub
Private Sub MENU_DEBUG_Click()
    If g_iDebugIndex = ListView.SelectedItem.index - 1 Then
        g_iDebugIndex = -1
        Label1.Caption = "ATָ���¼(��):"
    Else
        g_iDebugIndex = ListView.SelectedItem.index - 1
        TextLog.Text = ""
        Label1.Caption = "ATָ���¼(COM" & com(g_iDebugIndex).comPort & "):"
    End If
End Sub


Private Sub MENU_CCFC_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.index - 1
    com(cIdx).task.Push ("AT+CCFC=0,2" & vbCrLf)
    TimerComTask(cIdx).Enabled = True
End Sub

Private Sub MENU_INFO_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.index - 1 '"    ��ת����:" com(cidx).bMobile & vbcrlf &
    If com(cIdx).blIsOpen = False Then
        TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                   "     COM" & com(cIdx).comPort & ": δ����" & vbCrLf & _
                   "----- ----- ----- ----- ----- ----- -----" & vbCrLf
    Else
        TextDsp.Text = TextDsp.Text & "      ʱ��: " & Now() & vbCrLf & _
                   "     COM" & com(cIdx).comPort & ": �ѿ���" & vbCrLf & _
                   "  �ֻ�����: " & com(cIdx).Mobile & vbCrLf & _
                   "   ICCID��: " & com(cIdx).Iccid & vbCrLf & _
                   "  IMEI����: " & com(cIdx).Imei & vbCrLf & _
                   "  IMSI����: " & com(cIdx).Imsi & vbCrLf & _
                   "----- ----- ----- ----- ----- ----- -----" & vbCrLf
    End If
    TextDsp.SelStart = Len(TextDsp.Text)
End Sub


'**********************************************************************
' ����ɨ��
'**********************************************************************
Function comportScan(kc As Integer, comPort() As String)
    Dim I As Integer
    Dim ret As Long
    ReDim Preserve comPort(0)
    Dim com() As Integer
    com() = DB.GetComByKc(kc)
    For I = 0 To UBound(com())
        ret = sio_open(com(I))
        If ret = SIO_OK Then
            sio_close (com(I))
            comPort(UBound(comPort())) = com(I)
            ReDim Preserve comPort(UBound(comPort()) + 1)
        End If
    Next I
    ReDim Preserve comPort(UBound(comPort()) - 1)
End Function


Public Function OpenCom(ByRef cIdx As Integer)
    If com(cIdx).blIsOpen = False Then
        com(cIdx).OpenPort
        If com(cIdx).blIsOpen = False Then
            ListView.ListItems(cIdx + 1).SubItems(2) = ComErr(com(cIdx).portErr)
        Else
            ListView.ListItems(cIdx + 1).SubItems(2) = "��ʼ����(0)..."
            com(cIdx).task.Clean
            com(cIdx).task.Push ("ATE1" & vbCrLf)
            com(cIdx).task.Push ("AT+CFUN=1,1" & vbCrLf)
            ' ����������������ʱ��
            TimerComTask(cIdx).Enabled = True
            TimerComRead(cIdx).Enabled = True
        End If
    End If
End Function

Public Function CloseCom(ByRef cIdx As Integer)
    TimerComRead(cIdx).Enabled = False
    TimerComTask(cIdx).Enabled = False
    com(cIdx).blIsCheck = False
    TimerComCron(cIdx).Enabled = False
    ListView.ListItems(cIdx + 1).SubItems(3) = ""
    ListView.ListItems(cIdx + 1).SubItems(4) = ""
    ListView.ListItems(cIdx + 1).SubItems(5) = ""
    ListView.ListItems(cIdx + 1).SubItems(6) = ""
    'ListView.ListItems(cIdx + 1).SubItems(6) = ""
    DB.setCardClose ("'" & com(cIdx).Iccid & "'")
    com(cIdx).ReSet
    ListView.ListItems(cIdx + 1).SubItems(2) = "δ����"
End Function
