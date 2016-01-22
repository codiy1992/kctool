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
'*    �������ƣ�Task.cls
'*    �����ܣ�һ��VBѭ��������
'* ************************************************************** *
    Option Explicit
    Dim m_Task() '����
    Dim m_ReadIndex '����ƫ������
    Dim m_WriteIndex 'Ԫ�ز���ƫ������
    Dim m_MaxLen '�������ߴ�
   
    Private Sub Class_Initialize()
        m_ReadIndex = 0
        m_WriteIndex = 0
        m_MaxLen = 128 '����Ĭ�ϴ�СΪ128
        ReDim m_Task(m_MaxLen - 1)
    End Sub
   
    Private Sub Class_Terminate()
        Erase m_Task
    End Sub
   
    '��ö��д�С
    Public Property Get MaxLen()
        MaxLen = m_MaxLen
    End Property
   
    '���ö��д�С
    Public Property Let MaxLen(ByVal NewValue)
        If NewValue > 0 Then
            m_MaxLen = NewValue
            ReDim m_Task(m_MaxLen)
        End If
    End Property
   
    '�������Ԫ��
    Public Sub Push(ByVal v)
        If m_WriteIndex = m_MaxLen Then '�����������������Ȳ����Ԫ��
            m_WriteIndex = 0
            'm_ReadIndex = (m_ReadIndex + 1) Mod m_MaxLen
        End If
        m_Task(m_WriteIndex) = v
        m_WriteIndex = m_WriteIndex + 1
    End Sub
   
    'ɾ������Ԫ�أ�������ԭ��ʵ������ɾ�����Ȳ����Ԫ�أ�
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
   
    '��ն���
    Public Sub Clear()
        ReDim m_Task(m_MaxLen)
        m_WriteIndex = 0
        m_ReadIndex = 0
    End Sub