Attribute VB_Name = "mdlCommFuns"
Option Explicit

Public Function QueryRSSI(strInput As String, iSignal As Integer, iBER As Integer) As Boolean
    
    Dim strTmp As String
    Dim i As Integer, iLen As Integer, iTmp As Integer, iCr As Integer
    
    On Error Resume Next
    
    iCr = InStr(strInput, vbCr)
    If iCr > 0 Then
        For i = 7 To Len(strInput)
            strTmp = Mid(strInput, i, 1)
            iLen = iLen + 1
            If strTmp = "," Then
                iTmp = Mid(strInput, 7, iLen)
                iLen = 0
                iSignal = iTmp
                Exit For
            End If
        Next i
        iTmp = i
        For i = iTmp To Len(strInput)
            strTmp = Mid(strInput, i, 1)
            iLen = iLen + 1
            If strTmp = vbCr Then
                strTmp = Mid(strInput, iTmp + 1, iLen)
                If IsNumeric(strTmp) Then
                    iTmp = CInt(strTmp)
                    iLen = 0
                    iBER = iTmp
                    Exit For
                End If
            End If
        Next i
        strInput = Right(strInput, Len(strInput) - iCr)
        QueryRSSI = True
    Else
        QueryRSSI = False
    End If
End Function


Public Function GB2Unicode(ByVal strGB As String) As String

    Dim byteA()         As Byte
    
    Dim i               As Integer
    
    Dim strTmpUnicode   As String
    Dim strA            As String
    Dim strB            As String

    On Error GoTo ErrorUnicode
    
    i = LenB(strGB)
    
    ReDim byteA(1 To i)
    
    For i = 1 To LenB(strGB)
        strA = MidB(strGB, i, 1)
        byteA(i) = AscB(strA)
    Next i
    
    '��ʱ�Ѿ���strGBת��ΪUnicode���룬����������byteA()�С�
    '������Ҫ����˳�����ַ�������ʽ����
    strTmpUnicode = ""
    
    For i = 1 To UBound(byteA) Step 2
        strA = Hex(byteA(i))
        If Len(strA) < 2 Then strA = "0" & strA
        strB = Hex(byteA(i + 1))
        If Len(strB) < 2 Then strB = "0" & strB
        strTmpUnicode = strTmpUnicode & strB & strA
    Next i
    
    GB2Unicode = strTmpUnicode
    Exit Function
ErrorUnicode:
    GB2Unicode = ""
End Function

Public Function Unicode2GB(ByVal strUnicode As String) As String

    Dim byteA()     As Byte
    
    Dim i           As Integer
    
    Dim strTmp      As String
    Dim strTmpGB    As String
    
    
    On Error GoTo ErrUnicode2GB
    
    i = Len(strUnicode) / 2
    ReDim byteA(1 To i)
    
    For i = 1 To Len(strUnicode) / 2 Step 2
        strTmp = Mid(strUnicode, i * 2 - 1, 2)
        strTmp = Hex2Dec(strTmp)
        byteA(i + 1) = strTmp
        strTmp = Mid(strUnicode, i * 2 + 1, 2)
        strTmp = Hex2Dec(strTmp)
        byteA(i) = strTmp
    Next i
    
    strTmpGB = ""
    For i = 1 To UBound(byteA)
        strTmp = byteA(i)
        strTmpGB = strTmpGB & ChrB(strTmp)
    Next i
    
    Unicode2GB = strTmpGB
    Exit Function

ErrUnicode2GB:
    'MsgBox "Err=" & Err.Number & ",ԭ��" & Err.Description
    Unicode2GB = ""
End Function

Public Function Hex2Dec(ByVal strInput As String) As Long
    Dim i       As Integer
    Dim j       As Integer
    Dim iLen    As Integer
    Dim iTmp    As Integer
    
    Dim nRet    As Long
    Dim strTmp  As String
    
    On Error Resume Next
    
    If strInput <> "" Then
        iLen = Len(strInput)
        nRet = 0
        For i = 1 To iLen
            iTmp = Asc(Mid(strInput, i, 1))
            If iTmp >= 48 And iTmp <= 57 Then               '"0" = 48, "9" = 57
                nRet = nRet + (iTmp - 48) * 16 ^ (iLen - i)
            ElseIf iTmp >= 65 And iTmp <= 70 Then           '"A" = 65, "F" = 70
                nRet = nRet + (iTmp - 55) * 16 ^ (iLen - i)
            ElseIf iTmp >= 97 And iTmp <= 102 Then          '"a" = 97, "f" = 102
                nRet = nRet + (iTmp - 87) * 16 ^ (iLen - i)
            Else
                nRet = 0
                Exit For
            End If
        Next i
    End If
    
    Hex2Dec = nRet

End Function


'�˺����ǽ�һ���ַ�������charRefΪ�ָ�����Ԫ�ر��浽����MyStr()��
'*********************************************
'������
'============================================
'|YourStr��  |  ���ָ����ַ���
'+-----------+-------------------------------
'|charRef��  |  �ָ�����
'+-----------+-------------------------------
'|isNormal�� |  ���Ϊ�٣����ʾ�ָ���������
'|           |  ����ո���ɣ�����Tab���š�
'+-----------+-------------------------------
'|nD��       |  ����ֵ����ʾ�ж��ٸ�Ԫ��
'+-----------+-------------------------------
'|MyStr()��  |  ����ֵ������ָ���ĸ���Ԫ�ء�
'============================================
'
'**********************************************
Public Function String2Array(ByVal YourStr As String, _
                             ByVal charRef As String, _
                             ByRef nD As Long, _
                             ByRef MyStr() As String, _
                             ByVal isNormal As Boolean) As Boolean

    Dim i           As Long
    Dim j           As Long
    Dim nUBound     As Long
    
    Dim iAsc        As Integer
    
    Dim strChar     As String
    Dim strTmp      As String
    Dim aryTr()     As String
  
    On Error GoTo ErrorDecode

    strChar = ""
    YourStr = Trim(YourStr)     '����ȥ���ַ������ߵĿո�
    nUBound = 1
    j = 0
    ReDim aryTr(1 To nUBound)

    If Not isNormal Then
        For i = 1 To Len(YourStr)
            strTmp = Mid(YourStr, i, 1)
            iAsc = Asc(strTmp)
            If iAsc > 122 Or iAsc < 33 Then
                strChar = Mid(YourStr, i - j, j)
                If strChar <> "" Then
                    aryTr(nUBound) = strChar
                    nUBound = nUBound + 1
                    ReDim Preserve aryTr(1 To nUBound)
                End If
                strChar = ""
                j = 0
            Else
                j = j + 1
                If i = Len(YourStr) Then
                    strChar = Mid(YourStr, i - j + 1, j)
                    aryTr(nUBound) = strChar
                End If
            End If
        Next i
        nD = nUBound
        ReDim MyStr(0 To nUBound - 1)
        For i = 1 To nUBound
            MyStr(i - 1) = aryTr(i)
        Next i
        String2Array = True
    Else
        For i = 1 To Len(YourStr)
            strTmp = Mid(YourStr, i, 1)
            If strTmp = charRef Then
                strChar = Mid(YourStr, i - j, j)
                If strChar <> "" Then
                    aryTr(nUBound) = strChar
                    nUBound = nUBound + 1
                    ReDim Preserve aryTr(1 To nUBound)
                End If
                strChar = ""
                j = 0
            Else
                j = j + 1
                If i = Len(YourStr) Then
                    strChar = Mid(YourStr, i - j + 1, j)
                    aryTr(nUBound) = strChar
                End If
            End If
        Next i
        nD = nUBound
        ReDim MyStr(0 To nUBound - 1)
        For i = 1 To nUBound
            MyStr(i - 1) = aryTr(i)
        Next i
        String2Array = True
    End If

    Exit Function

ErrorDecode:
    'MsgBox Err.Number & ":" & Err.Description
    String2Array = False
End Function


Public Function ASCII2Char(ByVal strAsc As String) As String

    Dim i       As Integer
    Dim j       As Integer
    
    Dim strTmp  As String
    Dim strTmpA As String
    Dim strTmpB As String

    On Error Resume Next
    j = Len(strAsc)
    strTmpB = ""

    For i = 1 To j
        strTmpA = Mid(strAsc, i, 1)
        If strTmpA <> " " Then strTmpB = strTmpB & strTmpA
    Next i

    j = Len(strTmpB)

    strTmp = ""
    For i = 1 To j Step 2
        strTmpA = Mid(strTmpB, i, 2)
        strTmp = strTmp & ChrB(Hex2Dec(strTmpA))
    Next i

    ASCII2Char = strTmp

End Function

Public Function CharToAscii(ByVal strChar As String) As String
    Dim iAsc As Integer
    
    Dim n1      As Long
    Dim n2      As Long
    
    Dim strTmp  As String
    Dim strTmp1 As String
    Dim strTmp2 As String
    
    On Error Resume Next
    n1 = LenB(strChar)
    strTmp = ""
    
    For n2 = 1 To n1
        iAsc = AscB(MidB(strChar, n2, 1))
        If iAsc <> 0 Then
            strTmp1 = Hex(iAsc)
            If Len(strTmp1) < 2 Then strTmp1 = "0" & strTmp1
            strTmp = strTmp & strTmp1 & " "
        End If
    Next n2
    
    CharToAscii = Trim(strTmp)

End Function

