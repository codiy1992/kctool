Attribute VB_Name = "mdlCommVars"
Option Explicit

Public com() As com         ' ���ڶ�������
Public kc As kc             ' ���ض���
Public DB As New DB         ' ���ݿ����

Public g_blPrint
Public g_iDebugIndex As Integer ' ��ǰ���Դ���




'+CMGL: 1,"REC UNREAD","+8613181985843",,"04/06/04,15:31:25+00"
'00480069002C4F60597D5417003F
Type SMSDef
    ListOrRead As Boolean       '�Ƿ����о�(List)������ȡ
    SmsIndex As Long
    SourceNo As String
    ReachDate As String
    ReachTime As String
    SmsMain As String
    DateTime As String
End Type

