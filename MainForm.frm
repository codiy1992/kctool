VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form MainForm 
   Caption         =   "卡池助手 -- by CODIY"
   ClientHeight    =   9945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19275
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9945
   ScaleWidth      =   19275
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer TimerComCheck 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   3000
      Left            =   840
      Top             =   10920
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   300
      Left            =   840
      Top             =   10560
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   0
      Left            =   120
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer TimerComTask 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   300
      Left            =   120
      Top             =   10560
   End
   Begin VB.Frame Frame3 
      Caption         =   "当前调试串口(无)"
      Height          =   9735
      Left            =   15000
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Text1 
         Height          =   4455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox TextRec 
         Height          =   4695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   4920
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "操作按钮"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   7560
      Width           =   14775
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "写INI"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "开始校验"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "删除所有短信"
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "删除短信"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton CommandStopCheck 
         Caption         =   "停止检验"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton CommandSendSms 
         Caption         =   "发送短信"
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton CommandCloseAll 
         Caption         =   "关闭所有"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CommandStartAll 
         Caption         =   "开启所有"
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
      Left            =   120
      Top             =   10920
   End
   Begin VB.Frame Frame1 
      Caption         =   "串口状态"
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      Begin MSComctlLib.ListView ListView 
         Height          =   6915
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14500
         _ExtentX        =   25585
         _ExtentY        =   12197
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
   Begin MSCommLib.MSComm MSComm 
      Index           =   1
      Left            =   840
      Top             =   9960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Menu MenuList 
      Caption         =   "操作"
      Visible         =   0   'False
      Begin VB.Menu MENU_OPEN 
         Caption         =   "开启本串口"
         Shortcut        =   ^O
      End
      Begin VB.Menu MENU_CLOSE 
         Caption         =   "关闭本串口"
         Shortcut        =   ^C
      End
      Begin VB.Menu MENU_STOP 
         Caption         =   "暂停本串口"
         Shortcut        =   ^S
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_START 
         Caption         =   "恢复本串口"
         Visible         =   0   'False
      End
      Begin VB.Menu MENU_DEBUG 
         Caption         =   "调试本串口"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    ' 初始化数据
    g_iDebugIndex = -1
    ' 初始化列表视图控件
    With ListView 'ListView初始化
         .View = 3 ' 列表显示方式
         .ColumnHeaders.Add , , "序号", 600         ' 0
         .ColumnHeaders.Add , , "串口", 800         ' 1
         .ColumnHeaders.Add , , "状态", 1600        ' 2
         .ColumnHeaders.Add , , "手机号码", 1500    ' 3
         .ColumnHeaders.Add , , "ICCID号", 2500     ' 4
         .ColumnHeaders.Add , , "IMEI号", 1800      ' 5
         .ColumnHeaders.Add , , "IMSI号", 1800      ' 6
         .ColumnHeaders.Add , , "呼转号码", 1500    ' 7
         .ColumnHeaders.Add , , "信号", 600         ' 8
         .ColumnHeaders.Add , , "网络状态", 1800    ' 9
    End With

    ' 检测在线串口
    Call comportScan(comport_use())
    ReDim Com(UBound(comport_use()))
    'MsgBox comport_use(0) & " || " & comport_use(1)
    For i = 0 To UBound(comport_use()) - 1
        'LabelCom1.Caption = "COM" & comport_use(1)
        Set LV = ListView.ListItems.Add(i + 1, , i + 1)
        LV.SubItems(1) = "COM" & comport_use(i)
        Set Com(i) = New Com
        Com(i).comPort = comport_use(i)
    Next i
End Sub
Private Sub CommandStartAll_Click()
    Dim cIdx As Integer
    For cIdx = 0 To UBound(Com()) - 1
        MSComm(cIdx).CommPort = Com(cIdx).comPort
        On Error Resume Next        '判断该串口是否被打开
        MSComm(cIdx).PortOpen = True
        If Err = 8005 Then
            MSComm(cIdx).PortOpen = False
            ListView.ListItems(cIdx + 1).SubItems(2) = "端口被占用"
        Else
            With MSComm(cIdx)
                .Settings = "115200,n,8,1" 'Trim(Combo_rate.Text) & "," & Left(Trim(Combo_jiaoyan.Text), 1) & "," & Trim(Combo_databyte.Text) & "," & Trim(Combo_stopbyte)
                .RThreshold = 1
                .InputMode = comInputModeBinary
            End With
            ListView.ListItems(cIdx + 1).SubItems(2) = "初始化中(0)..."
            'ListView.ListItems(cIdx + 1).SubItems(3) = Com(cIdx).Mobile
            Com(cIdx).task.Push ("ATE1" & vbCrLf)
            Com(cIdx).task.Push ("AT+CFUN=1,1" & vbCrLf)
            ' 开启串口命令任务定时器
            TimerComTask(cIdx).Enabled = True
        End If
    Next cIdx
End Sub
Private Sub CommandCloseAll_Click()
    ' 关闭串口
    For cIdx = 0 To UBound(Com()) - 1
        'TimerComTask(cIdx).Enabled = False
        TimerComCheck(cIdx).Enabled = False
        If MSComm(cIdx).PortOpen = True Then
            MSComm(cIdx).PortOpen = False
            ListView.ListItems(cIdx + 1).SubItems(2) = "未开启"
        End If
    Next cIdx
End Sub

Private Sub CommandSendSms_Click()
    TimerComCheck(0).Enabled = False
    Com(0).sendSMS "10027", "102", True
    TimerComTask(0).Enabled = True
End Sub

Private Sub Command3_Click()
    Com(0).blIsCheck = True
    Com(0).blIsPickSms = True
End Sub

Private Sub CommandStopCheck_Click()
    Com(0).blIsCheck = False
    Com(0).blIsPickSms = False
End Sub
Private Sub Command1_Click()
    Com(0).delSMS 1
End Sub

Private Sub Command2_Click()
    Com(0).delSMS 1, True
End Sub



Private Sub MSComm_OnComm(cIdx As Integer)
    Dim strAtData As String
    Dim tmpBuf() As Byte, strTmp As String
    Dim strOut As String
    Dim strAT As String
    Dim smsArr() As SMSDef
On Error Resume Next
    Select Case MSComm(cIdx).CommEvent
        Case comEvReceive
            tmpBuf() = MSComm(cIdx).Input
            For i = 0 To UBound(tmpBuf())
                strTmp = strTmp & Chr(Val(tmpBuf(i)))
            Next i
            If cIdx = g_iDebugIndex Then
                Text1.Text = Text1.Text & strTmp ' & "------------------"
                Text1.SelStart = Len(Text1.Text)
            End If
            strAtData = Com(cIdx).GetData(strTmp)
            If strAtData <> Empty Then
                strAT = Com(cIdx).AnalysisData(strAtData, strOut)
                If strAT <> "" And strAT <> vbCr And strAtData <> vbCrLf And Not IsEmpty(strAT) Then
                    'TextRec.Text = TextRec.Text & strOut & "------------------" & vbCrLf
                    Select Case strAT
                        Case "AT+CSQ"
                            ListView.ListItems(cIdx + 1).SubItems(8) = strOut
                        Case "AT+COPS?"
                            ListView.ListItems(cIdx + 1).SubItems(9) = strOut
                        Case "AT+CGSN"
                            Com(cIdx).Imei = strOut
                            ListView.ListItems(cIdx + 1).SubItems(5) = strOut
                        Case "AT+CIMI"
                            Com(cIdx).Imsi = strOut
                            ListView.ListItems(cIdx + 1).SubItems(6) = strOut
                        Case "AT+CMGL"
                            Com(cIdx).blIsPickSms = False    ' 处理完成前,将不在接受取短信命令
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
                                        DB.SaveSMS Com(cIdx).Iccid, smsArr(n).SourceNo, smsArr(n).SmsMain, smsArr(n).DateTime, "17070594726"
                                        Com(cIdx).task.Push ("AT+CMGD=" & smsArr(n).SmsIndex & vbCrLf)
                                     End If
                                Next n
                            End If
                            Com(cIdx).blIsPickSms = True    ' 处理完成后,继续接受新的取短信命令
                        Case "+CMTI:" ' 收到新短信
                        Case "-AT-SMS-SEND-OK" ' 短信发送成功
                            TimerComCheck(cIdx).Enabled = True
                        Case "-AT-INIT-OK-"    ' 初始化【步骤一：成功】
                            Com(cIdx).task.Push ("ATE1" & vbCrLf)         ' 开启回显
                            Com(cIdx).task.Push ("AT+CIURC=0" & vbCrLf)
                            Com(cIdx).task.Push ("AT+CGSN" & vbCrLf)
                            Com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
                            Com(cIdx).task.Push ("AT+CNMI=2,1" & vbCrLf)  '
                            Com(cIdx).task.Push ("AT+CMGF=1" & vbCrLf)    ' 设置短信格式
                            Com(cIdx).task.Push ("AT+CGSN" & vbCrLf)
                            Com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
                            Com(cIdx).task.Push ("AT+CCID" & vbCrLf)      ' 查询ICCID号
                            ListView.ListItems(cIdx + 1).SubItems(2) = "初始化中(1)..."
                            ' 开启串口命令任务定时器
                            TimerComTask(cIdx).Enabled = True
                        Case "AT+CCID"         ' 初始化【步骤二：成功】
                            Com(cIdx).Iccid = strOut
                            ListView.ListItems(cIdx + 1).SubItems(4) = strOut
                            ' 向数据库注册本SIM卡，并获取手机号码
                            Com(cIdx).Mobile = DB.RegistCard(Com(cIdx).Iccid, Com(cIdx).Imei, Com(cIdx).Imsi)
                            If Com(cIdx).Mobile = "" Then
                                ListView.ListItems(cIdx + 1).SubItems(2) = "请先设置手机号"
                            Else
                                Com(cIdx).blIsOpen = True
                                ListView.ListItems(cIdx + 1).SubItems(2) = "正常工作"
                                ListView.ListItems(cIdx + 1).SubItems(3) = Com(cIdx).Mobile
                                TimerComCheck(cIdx).Enabled = True
                            End If
                        Case Else
                    End Select
                        
                End If
            End If
            
        '''''''''''''''''''''''''''''''''''''''
        Case comEventBreak
            TextCom1Stat.Text = "Modem发出中断信号，希望计算机能等候，请稍候."
            MSComm(cIdx).PortOpen = False
            MSComm(cIdx).PortOpen = True
        Case comEvCTS
            If MSComm(cIdx).CTSHolding = True Then 'Modem表示计算机可以发送数据
                TextCom1Stat.Text = "Modem能够接收计算机数据"
            Else 'Modem无法响应计算机数据，可能缓冲区不够
                TextCom1Stat.Text = "Modem请求计算机暂时不要发送数据"
                MSComm(cIdx).DTREnable = Not MSComm(cIdx).DTREnable
                DoEvents
                MSComm(cIdx).DTREnable = Not MSComm(cIdx).DTREnable
            End If
        Case comEvDSR
            If MSComm(cIdx).DSRHolding = True Then '当Modem收到计算机已经就绪，Modem表示自己也就绪
                TextCom1Stat.Text = "Modem可以给计算机发送数据"
            Else '在计算机发出DTR信号后，Modem可能还没有就绪
                TextCom1Stat.Text = "Modem还没有初始化完毕"
            End If
        Case comEventFrame
            MSComm(cIdx).PortOpen = False
            MSComm(cIdx).PortOpen = True
        Case comEvRing
            TextCom1Stat.Text = "检测到振铃变化"
        Case comEvCD
            TextCom1Stat.Text = "检测到载波变化"
        Case Else
            'MsgBox MSComm(cIdx).CommEvent
    End Select
End Sub


Private Sub TimerComCheck_Timer(cIdx As Integer)
    If Com(cIdx).blIsCheck = True Then
        If Com(cIdx).Imei = "" Then
            Com(cIdx).task.Push ("AT+CGSN" & vbCrLf)      ' Imei
        End If
        If Com(cIdx).Imsi = "" Then
            Com(cIdx).task.Push ("AT+CIMI" & vbCrLf)
        End If
        Com(cIdx).task.Push ("AT+CSQ" & vbCrLf)
        Com(cIdx).task.Push ("AT+COPS?" & vbCrLf)
    End If
    If Com(cIdx).blIsPickSms = True Then
        Com(cIdx).task.Push ("AT+CMGL=""ALL""" & vbCrLf)
        Com(cIdx).blIsPickSms = False ' 命令处理完成前，将不再发送新的取短信命令
    End If
    If TimerComTask(cIdx).Enabled = False Then TimerComTask(cIdx).Enabled = True
End Sub

Private Sub TimerComTask_Timer(cIdx As Integer)
    Dim task As String
    Dim tail(0) As Byte
    task = Com(cIdx).task.Pop
    If task <> Empty Then
        If MSComm(cIdx).PortOpen = True Then
            Select Case task
                Case "--TAIL--"
                    tail(0) = &H1A
                    MSComm(cIdx).Output = tail
                Case "--STOP--"
                    TimerComTask(cIdx).Enabled = False
                    ListView.ListItems(cIdx + 1).SubItems(2) = "已暂停"
                Case "--CLOSE--"
                    If MSComm(cIdx).PortOpen = True Then
                        MSComm(cIdx).PortOpen = False
                    End If
                    Com(cIdx).ReSet
                    TimerComTask(cIdx).Enabled = False
                    ListView.ListItems(cIdx + 1).SubItems(2) = "未开启"
                Case Else
                    MSComm(cIdx).Output = task
            End Select
        End If
    Else
       TimerComTask(cIdx).Enabled = False
    End If
End Sub

Private Sub ListView_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim Index As Integer
    '按下鼠标右键
    If Button = vbRightButton Then
        If ListView.ListItems.Count > 0 Then
            Index = ListView.SelectedItem.Index
            If Com(Index - 1).blIsOpen = True Then
                ' 开启关闭菜单
                MENU_CLOSE.Visible = True
                MENU_CLOSE.Enabled = True
                ' 开启调试菜单
                MENU_DEBUG.Visible = True
                MENU_DEBUG.Enabled = True
                ' 关闭开启菜单
                MENU_OPEN.Visible = False
                MENU_OPEN.Enabled = False
                
                If MENU_START.Visible = False Then
                    MENU_STOP.Visible = True
                    MENU_STOP.Enabled = True
                End If
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
                ' 关闭调试菜单
                MENU_DEBUG.Visible = False
                MENU_DEBUG.Enabled = False
                
            End If
            
            PopupMenu MenuList
        End If
    End If
End Sub
Private Sub MENU_OPEN_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.Index - 1
    MSComm(cIdx).CommPort = Com(cIdx).comPort
    On Error Resume Next        '判断该串口是否被打开
    If MSComm(cIdx).PortOpen = False Then
        MSComm(cIdx).PortOpen = True
        If Err = 8005 Then
            MSComm(cIdx).PortOpen = False
            ListView.ListItems(cIdx + 1).SubItems(2) = "端口被占用"
        Else
            With MSComm(cIdx)
                .Settings = "115200,n,8,1"
                .RThreshold = 1
                .InputMode = comInputModeBinary
            End With
            ListView.ListItems(cIdx + 1).SubItems(2) = "初始化中(0)..."
            Com(cIdx).task.Push ("ATE1" & vbCrLf)
            Com(cIdx).task.Push ("AT+CFUN=1,1" & vbCrLf)
            ' 开启串口命令任务定时器
            TimerComTask(cIdx).Enabled = True
        End If
    End If
End Sub
Private Sub MENU_CLOSE_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.Index - 1
    Com(cIdx).task.Push ("--CLOSE--") ' 将停止命令推送入命令执行队列
    TimerComTask(cIdx).Enabled = True
    TimerComCheck(cIdx).Enabled = False
End Sub
Private Sub MENU_STOP_Click()
    Dim cIdx As Integer
    cIdx = ListView.SelectedItem.Index - 1
    Com(cIdx).task.Push ("--STOP--") ' 将停止命令推送入命令执行队列
    TimerComTask(cIdx).Enabled = True
    TimerComCheck(cIdx).Enabled = False
    ' 菜单
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
    ListView.ListItems(cIdx + 1).SubItems(2) = "正常工作"
    ' 菜单
    MENU_START.Visible = False
    MENU_START.Enabled = False
    MENU_STOP.Visible = True
    MENU_STOP.Enabled = True
End Sub
Private Sub MENU_DEBUG_Click()
    g_iDebugIndex = ListView.SelectedItem.Index - 1
End Sub



'**********************************************************************
' 串口扫描
'**********************************************************************
Function comportScan(comPort() As String)
    Dim i As Integer
    ReDim Preserve comPort(0)
    For i = 2 To 255
        On Error Resume Next
        MSComm(0).CommPort = i
        MSComm(0).PortOpen = True
        If Err = 0 Or Err = 8005 Then
        comPort(UBound(comPort())) = i
        ReDim Preserve comPort(UBound(comPort()) + 1)
        End If
        MSComm(0).PortOpen = False
    Next i
End Function
'**********************************************************************
' 发送短信
'**********************************************************************
Public Function sendSMS(ByRef strIccid As String, strTo As String, strCnt As String)
    objCOM.sendSMS objTask, strMobile, strCnt
End Function
