Attribute VB_Name = "mdlCommVars"
Option Explicit

Public Com() As Com         ' 串口对象数组
Public Kc As Com
Public DB As New DB         ' 数据库对象
Public g_show As Boolean
Public g_iDebugIndex As Integer ' 当前调试串口



'+CMGL: 1,"REC UNREAD","+8613181985843",,"04/06/04,15:31:25+00"
'00480069002C4F60597D5417003F
Type SMSDef
    ListOrRead As Boolean       '是否用列举(List)方法读取
    SmsIndex As Long
    SourceNo As String
    ReachDate As String
    ReachTime As String
    SmsMain As String
    DateTime As String
End Type

