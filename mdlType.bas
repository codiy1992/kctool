Attribute VB_Name = "mdlType"
Option Explicit


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



