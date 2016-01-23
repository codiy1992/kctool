Attribute VB_Name = "mdlCommVars"
Option Explicit

Public Com() As Com         ' 串口对象数组
Public DB As New DB         ' 数据库对象
Public g_iDebugIndex As Integer ' 当前调试串口
