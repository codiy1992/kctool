Attribute VB_Name = "mdlType"
Option Explicit

Type SysStruct
    CommPort As Integer
    Baud As String
    ServiceNo As String
    DestNo As String
    CallMelody As Integer
    SMSMelody As Integer
    Clock As Boolean
    ClockSet As String
End Type

Type SMSDef
    
'+CMGL: 1,"REC UNREAD","+8613181985843",,"04/06/04,15:31:25+00"
'00480069002C4F60597D5417003F
    ListOrRead As Boolean       '�Ƿ����о�(List)������ȡ
    SmsIndex As Long
    SourceNo As String
    ReachDate As String
    ReachTime As String
    SmsMain As String
    DateTime As String
End Type


