VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* ************************************************************** *
'*    程序名称：Task.cls
'*    程序功能：一个VB循环队列类  -- by codiy
'* ************************************************************** *

Option Explicit
Private m_Task() '队列
Private m_ReadIndex '列首偏移索引
Private m_WriteIndex '元素插入偏移索引
Private m_MaxLen '队列最大尺寸

Private Sub Class_Initialize()
    m_ReadIndex = 0
    m_WriteIndex = 0
    m_MaxLen = 64 '队列默认大小为64
    ReDim m_Task(m_MaxLen - 1)
End Sub
   
 '获得队列大小
Public Property Get MaxLen()
     MaxLen = m_MaxLen
End Property
'设置队列大小
Public Property Let MaxLen(ByVal NewValue)
    If NewValue > 0 Then
        m_MaxLen = NewValue
        ReDim m_Task(m_MaxLen)
    End If
End Property

Public Property Get rIndex()
     rIndex = m_ReadIndex
End Property
Public Property Get wIndex()
     wIndex = m_WriteIndex
End Property

'插入队列元素
Public Sub Push(ByVal v)
    If m_WriteIndex = m_MaxLen Then '队列已满，挤掉最先插入的元素
        m_WriteIndex = 0
    End If
    m_Task(m_WriteIndex) = v
    m_WriteIndex = m_WriteIndex + 1
End Sub

'弹出栈顶元素
Public Function Pop() As String
    If m_ReadIndex = m_MaxLen Then
        m_ReadIndex = 0
    End If
    If IsEmpty(m_Task(m_ReadIndex)) Then
        Pop = Empty
    Else
        Pop = m_Task(m_ReadIndex)
        m_Task(m_ReadIndex) = Empty
        m_ReadIndex = m_ReadIndex + 1
    End If
End Function

' 栈顶元素
Public Function Top() As String
    If m_ReadIndex = m_MaxLen Then
        m_ReadIndex = 0
    End If
    If IsEmpty(m_Task(m_ReadIndex)) Then
        Top = Empty
    Else
        Top = m_Task(m_ReadIndex)
    End If
End Function
 
'清空队列
Public Sub Clean()
    ReDim m_Task(m_MaxLen)
    m_WriteIndex = 0
    m_ReadIndex = 0
End Sub
 
Private Sub Class_Terminate()
    Erase m_Task
 End Sub
