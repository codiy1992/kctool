VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "卡池助手 -- by CODIY"
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
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "操作按钮"
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   2775
      Begin VB.CommandButton Command1 
         Caption         =   "回显已开"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton CommandSwitchIndex 
         Caption         =   "指定切卡"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CommandDspInfo 
         Caption         =   "显示信息"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton CommandGetCCFC 
         Caption         =   "获取呼转"
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton CommandSwitch 
         Caption         =   "自动切卡"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton CommandClnTextLog 
         Caption         =   "清空AT记录"
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton CommandClnTextDsp 
         Caption         =   "清空回显"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton CommandKcSwitch 
         Caption         =   "切下一排"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CommandKcReSet 
         Caption         =   "卡池重置"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton CommandCloseAll 
         Caption         =   "关闭所有"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CommandStartAll 
         Caption         =   "搜索设备"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "串口列表"
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
            Name            =   "宋体"
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
      Caption         =   "AT指令记录:"
      Height          =   255
      Left            =   7680
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "数据回显窗口:"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Menu MenuList 
      Caption         =   "操作"
      Visible         =   0   'False
      Begin VB.Menu MENU_OPEN 
         Caption         =   "开启本串口"
      End
      Begin VB.Menu MENU_CLOSE 
         Caption         =   "关闭本串口"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_STOP 
         Caption         =   "暂停本串口"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_START 
         Caption         =   "恢复本串口"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_INFO 
         Caption         =   "查看串口信息"
      End
      Begin VB.Menu MENU_CCFC 
         Caption         =   "查呼转号码"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_DEBUG 
         Caption         =   "查看数据流"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 程序异常修复，数据库数据修复

Private Sub Command1_Click()
    If g_blPrint = True Then
        g_blPrint = False
        Command1.Caption = "回显已关"
    Else
        g_blPrint = True
        Command1.Caption = "回显已开"
    End If
End Sub

'1、初始化时，判断串口是卡池还是sim通道
'2、切卡时候校验，本次切卡与上次是否为同一张卡(空卡和最后一张)
'3、卡池与sim通道配对信息
'4、多卡池支持
'5、对卡类型的支持
Private Sub Form_Load()
    ' 初始化数据
    g_iDebugIndex = -1
    g_blPrint = True
    ' 初始化列表视图控件
    With ListView 'ListView初始化
         .View = 3 ' 列表显示方式
         .ColumnHeaders.Add , , "序号", 600         ' 0
         .ColumnHeaders.Add , , "串口", 800         ' 1
         .ColumnHeaders.Add , , "状态", 3200        ' 2
         .ColumnHeaders.Add , , "ICCID号", 2300     ' 3
         .ColumnHeaders.Add , , "手机号码", 1400    ' 4
         .ColumnHeaders.Add , , "IMEI号", 1800      ' 5
         '.ColumnHeaders.Add , , "信号", 1700         ' 6
         .ColumnHeaders.Add , , "占用", 800, lvwColumnCenter   ' 7
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case vbFormControlMenu          ' 0 用户从窗体上的“控件”菜单中选择“关闭”指令。
        Case vbFormCode                 ' 1 Unload 语句被代码调用。
        Case vbAppWindows               ' 2 当前 Microsoft Windows 操作环境会话结束。
        Case vbAppTaskManager           ' 3 Microsoft Windows 任务管理器正在关闭应用程序。
        Case vbFormMDIForm              ' 4 MDI 子窗体正在关闭，因为 MDI 窗体正在关闭。
        Case vbFormOwner                ' 5 因为窗体的所有者正在关闭，所以窗体也在关闭。
    End Select
    If IsComEmpty = False Then
        If GetAllIccid <> "''" Then
            If MsgBox("请先关闭全部串口后再操作", vbYesNo + vbDefaultButton1) = vbYes Then
                For cIdx = 0 To UBound(com())
                    ' 将停止命令推送入命令执行队列
                    com(cIdx).task.Push ("--CLOSE--")
                    TimerComTask(cIdx).Enabled = True
                Next cIdx
            End If
            Cancel = True
        End If
    End If
End Sub

' ====================================  操作按钮  ====================================================
Private Sub CommandStartAll_Click()
    Dim cIdx As Integer
    If IsComEmpty = False Then
        For cIdx = 0 To UBound(com())
            OpenCom (cIdx)
        Next cIdx
        kc.blCanSwitch = True
    ElseIf TimerInit.Enabled = False Then
        ' 检测在线串口
        CommandStartAll.Caption = "搜索中..."
        TimerInit.Enabled = True
    End If
End Sub

Private Sub CommandCloseAll_Click()
    ' 关闭串口
    If IsComEmpty = False Then
        For cIdx = 0 To UBound(com())
           ' 将停止命令推送入命令执行队列
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
                TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                           "     COM" & com(cIdx).comPort & ": 未开启" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
            Else
                TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                           "     COM" & com(cIdx).comPort & ": 已开启" & vbCrLf & _
                           "  手机号码: " & com(cIdx).Mobile & vbCrLf & _
                           "   ICCID号: " & com(cIdx).Iccid & vbCrLf & _
                           "  IMEI串号: " & com(cIdx).Imei & vbCrLf & _
                           "  IMSI串号: " & com(cIdx).Imsi & vbCrLf & _
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
    simIndex = InputBox("sim卡位置(1-16)", "输入切换位置", 1)
    simIndex = Val(simIndex)
    If IsComEmpty = True Then
        Exit Sub
    End If
    If simIndex < 1 Or simIndex > 16 Then
        MsgBox "请输入1-16之间的数字"
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
    
    Frame1.Caption = "卡池 - 自动切卡[" & kc.iTmrCnt * 3 & "/180]:" & kc.blCanSwitch & "    (" & Now() & ")"
    If IsComEmpty = True Then
        Exit Sub
    End If
    
    ' 自动切卡任务
    If kc.blCanSwitch = True Then
        If kc.iTmrCnt = 0 Then
            If DB.IsUsing(GetAllIccid) = False Then
                ' 切卡
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
    ' 6秒重置超时sim卡使用状态
    DB.setUseTimeOut
    ' 执行切卡任务
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
    
    ' 执行绑定无条件呼转任务
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
    
    ' 执行发送短信任务
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
'==============================================  卡池控制模块 ==================================================================
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
                Case "--KC-NEXT-A--" ' 切换到下一张卡(全部通道)
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
                Case "--KC-NEXT-AS--" ' 切换到下一张卡(全部通道)
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
                Case "--KC-RESET--" ' 卡池重置(全部通道)
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
                    If InStr(strAT, "--KC-INDEX-A") > 0 Then ' 指定切换到某张卡(全部通道)
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
                    TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                           "  卡池切卡: 成功 【切至：" & com(0).simIndex & "】" & vbCrLf & _
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
                    TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                           "  卡池切卡: 成功 【切至：" & com(0).simIndex & "】" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                End If
                
            Case "AT+NEXT11" & vbCrLf ' 卡池全部切换成功
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
                    TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                           "  卡池重置: 成功" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                    TimerKcCron.Enabled = True
                End If
                
            Case "AT+NEXT00" & vbCrLf ' 卡池重置成功
                kc.iNullCnt = 0
                If IsComEmpty = False Then
                    For cIdx = 0 To UBound(com())
                        com(cIdx).simIndex = 1
                        OpenCom (cIdx)
                    Next cIdx
                    
                End If
                kc.rowIndex = 1
            Case "AT+CWSIM" & vbCr   ' 与卡池握手成功
                'Kc.task.Push ("AT+NEXT00" & vbCrLf)
                kc.task.Push ("--KC-RESET--")
                TimerKcTask.Enabled = True
                
            Case Else
                If InStr(strAT, "AT+NEXT11-") > 0 Then '"AT+NEXT11-06" & vbCrLf
                    cIdx = Int(Left(Right(strAT, 4), 2)) - 1
                    com(cIdx).simIndex = com(cIdx).simIndex + 1
                    OpenCom (cIdx)
                    
                End If
                
                ' 切换到指定位置(全部通道)
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
                    TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                           "  卡池切卡: 成功 【切至：" & Val(Right(strAT, 2)) & "】" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                    
                End If
                ' 切换到指定位置(全部通道),
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
                    TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                           "  卡池切卡: 成功 【切至：" & Val(Right(strAT, 2)) & "】" & vbCrLf & _
                           "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    End If
                End If
                ' 切换到指定位置(全部通道)
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
                    TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                           "  卡池切卡: 成功 【切至：" & Val(Right(strAT, 2)) & "】" & vbCrLf & _
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
                
                ' 切换到指定位置(某通道)
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
    
    ' 卡池
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
            LV.SubItems(2) = "未开启"
        Next I
        CommandStartAll.Caption = "开启所有"
    Else
        CommandStartAll.Caption = "无可用设备"
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
    
        ' 如果正在发短信或者设置呼转,则不往任务队列推送新的命令任务
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
            com(cIdx).blIsPickSms = False ' 命令处理完成前，将不再发送新的取短信命令
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
                ListView.ListItems(cIdx + 1).SubItems(2) = "已暂停"
            Case "--CLOSE--"
                TimerComRead(cIdx).Enabled = True
                If com(cIdx).blIsOpen = True Or com(cIdx).Iccid <> "" Then ' 开着的或者无手机号的
                    com(cIdx).blIsCheck = False
                    com(cIdx).WriteData ("AT+CFUN=0" & vbCrLf)
                Else    ' 无sim卡
                     strAT = com(cIdx).task.Pop
                     CloseCom (cIdx)
                End If
            Case "--SWITCH--"
                TimerComRead(cIdx).Enabled = True
                If com(cIdx).blIsOpen = True Or com(cIdx).Iccid <> "" Then ' 开着的或者无手机号的
                    com(cIdx).blIsCheck = False
                    com(cIdx).blIsATExecing = True
                    com(cIdx).WriteData ("AT+CFUN=0" & vbCrLf)
                Else    ' 无sim卡
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
            ListView.ListItems(cIdx + 1).SubItems(2) = "串口无响应"
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
                    ' 查数据库，提取数据库绑定手机
                    
                    ' 输出到回显窗口
                    If strOut <> "" Then
                        TextDsp.Text = TextDsp.Text & Now() & vbCrLf & _
                                            "手机号: " & com(cIdx).Mobile & "(" & com(cIdx).Iccid & ")" & vbCrLf & _
                                            "    设置的呼转号码为 【" & strOut & "】" & vbCrLf & _
                                            "----- ----- ----- ----- ----- ----- -----" & vbCrLf
                    Else
                        TextDsp.Text = TextDsp.Text & Now() & vbCrLf & _
                                        "手机号: " & com(cIdx).Mobile & "(" & com(cIdx).Iccid & ")" & vbCrLf & _
                                        "    没查到或没设置呼转号码" & vbCrLf & _
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
                    com(cIdx).task.Push ("AT+CMGF=1" & vbCrLf)    ' 设置短信格式
                    com(cIdx).task.Push ("AT+CSCS=""UCS2""" & vbCrLf)    ' 设置短信编码
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
                com(cIdx).blIsPickSms = True    ' 处理完成后,继续接受新的取短信命令
            Case "+CMTI:" ' 收到新短信
            Case "-AT-SMS-SEND-OK-" ' 短信发送成功
                DB.SetSMSSended com(cIdx).Iccid, com(cIdx).iSmsId
                com(cIdx).iSmsId = 0
            Case "-AT-SMS-SEND-FAILED-" '短信发送失败
                DB.SetSMSNotSend com(cIdx).Iccid, com(cIdx).iSmsId
                com(cIdx).iSmsId = 0
            Case "-AT-INIT-OK-"    ' 初始化【步骤一：成功】
                com(cIdx).task.Push ("ATE1" & vbCrLf)         ' 开启回显
                'com(cIdx).task.Push ("AT+CGSN" & vbCrLf)
                com(cIdx).task.Push ("AT+CGSN" & vbCrLf)
                com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
                com(cIdx).task.Push ("AT+CNMI=2,1" & vbCrLf)  '
                com(cIdx).task.Push ("AT+CSCS=""UCS2""" & vbCrLf)    ' 设置短信编码
                com(cIdx).task.Push ("AT+CCID" & vbCrLf)      ' 查询ICCID号
                ListView.ListItems(cIdx + 1).SubItems(2) = "初始化中(1)..."
                ' 开启串口命令任务定时器
                TimerComTask(cIdx).Interval = 100
                TimerComTask(cIdx).Enabled = True
            Case "AT+CCID"         ' 初始化【步骤二：成功】
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
                    com(cIdx).task.Push ("AT+CCID" & vbCrLf)      ' 查询ICCID号
                    TimerComTask(cIdx).Enabled = True
                ElseIf strOut = "-NO-CCID-" Then
                    ListView.ListItems(cIdx + 1).SubItems(2) = "串口无SIM卡"
                    com(cIdx).Iccid = ""
                    com(cIdx).blIsShowStat = False
                    TimerComRead(cIdx).Enabled = False
                    kc.iNullCnt = kc.iNullCnt + 1
                    If kc.iNullCnt = 16 Then
                        kc.iNullCnt = 0
                        If kc.blCanSwitch = True Then
                            If DB.IsUsing(GetAllIccid) = False Then
                                ' 切卡
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
                    ' 向数据库注册本SIM卡，并获取手机号码
                    com(cIdx).Mobile = DB.RegistCard(com(cIdx).Iccid, _
                                                    Format(com(cIdx).kc, "000") & "-" & Format(com(cIdx).simIndex, "00") & "-" & Format(cIdx + 1, "00"), _
                                                    com(cIdx).Imei, com(cIdx).Imsi)
                    If com(cIdx).Mobile = "" Then
                        com(cIdx).blIsShowStat = False
                        ListView.ListItems(cIdx + 1).SubItems(2) = "请先设置手机号"
                        TimerComRead(cIdx).Enabled = False
                    Else
                        ListView.ListItems(cIdx + 1).SubItems(2) = "正常工作"
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
            Case "-AT-EXIT-OK-"  '关闭sim指令 OK
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
    '按下鼠标右键
    If Button = vbRightButton Then
        If ListView.ListItems.Count > 0 Then
            index = ListView.SelectedItem.index
            If com(index - 1).blIsOpen = True Then
                ' 开启关闭菜单
                MENU_CLOSE.Visible = True
                MENU_CLOSE.Enabled = True
                ' 关闭开启菜单
                MENU_OPEN.Visible = False
                MENU_OPEN.Enabled = False
                
                If MENU_START.Visible = False Then
                    MENU_STOP.Visible = True
                    MENU_STOP.Enabled = True
                End If
                
                MENU_CCFC.Visible = True
                MENU_CCFC.Enabled = True
            Else
                ' 开启开启菜单
                MENU_OPEN.Visible = True
                MENU_OPEN.Enabled = True
                ' 关闭关闭菜单
                MENU_CLOSE.Visible = False
                MENU_CLOSE.Enabled = False
                ' 关闭暂停菜单
                MENU_STOP.Visible = False
                MENU_STOP.Enabled = False
                ' 关闭恢复菜单
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
    ' 修改数据库字段状态
    DB.setCardClose ("'" & com(cIdx).Iccid & "'")
    ' 将停止命令推送入命令执行队列
    If com(cIdx).task.Top <> "--CLOSE--" Then
        com(cIdx).task.Push ("--CLOSE--")
        TimerComTask(cIdx).Enabled = True
    End If
    com(cIdx).blIsCheck = False
End Sub
Private Sub MENU_STOP_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.index - 1
    com(cIdx).task.Push ("--STOP--") ' 将停止命令推送入命令执行队列
    com(cIdx).blIsCheck = False
    TimerComTask(cIdx).Enabled = True
    TimerComRead(cIdx).Enabled = False
    ' 菜单
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
    ListView.ListItems(cIdx + 1).SubItems(2) = "正常工作"
    ' 菜单
    MENU_START.Visible = False
    MENU_START.Enabled = False
    MENU_STOP.Visible = True
    MENU_STOP.Enabled = True
End Sub
Private Sub MENU_DEBUG_Click()
    If g_iDebugIndex = ListView.SelectedItem.index - 1 Then
        g_iDebugIndex = -1
        Label1.Caption = "AT指令记录(无):"
    Else
        g_iDebugIndex = ListView.SelectedItem.index - 1
        TextLog.Text = ""
        Label1.Caption = "AT指令记录(COM" & com(g_iDebugIndex).comPort & "):"
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
    cIdx = ListView.SelectedItem.index - 1 '"    呼转号码:" com(cidx).bMobile & vbcrlf &
    If com(cIdx).blIsOpen = False Then
        TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                   "     COM" & com(cIdx).comPort & ": 未开启" & vbCrLf & _
                   "----- ----- ----- ----- ----- ----- -----" & vbCrLf
    Else
        TextDsp.Text = TextDsp.Text & "      时间: " & Now() & vbCrLf & _
                   "     COM" & com(cIdx).comPort & ": 已开启" & vbCrLf & _
                   "  手机号码: " & com(cIdx).Mobile & vbCrLf & _
                   "   ICCID号: " & com(cIdx).Iccid & vbCrLf & _
                   "  IMEI串号: " & com(cIdx).Imei & vbCrLf & _
                   "  IMSI串号: " & com(cIdx).Imsi & vbCrLf & _
                   "----- ----- ----- ----- ----- ----- -----" & vbCrLf
    End If
    TextDsp.SelStart = Len(TextDsp.Text)
End Sub


'**********************************************************************
' 串口扫描
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
            ListView.ListItems(cIdx + 1).SubItems(2) = "初始化中(0)..."
            com(cIdx).task.Clean
            com(cIdx).task.Push ("ATE1" & vbCrLf)
            com(cIdx).task.Push ("AT+CFUN=1,1" & vbCrLf)
            ' 开启串口命令任务定时器
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
    ListView.ListItems(cIdx + 1).SubItems(2) = "未开启"
End Function
