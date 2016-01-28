Attribute VB_Name = "mdlBaseFuns"
Option Explicit

Public Function QueryRSSI(strInput As String, iSignal As Integer, iBER As Integer) As Boolean
    
    Dim strTmp As String
    Dim I As Integer, iLen As Integer, iTmp As Integer, iCr As Integer
    
    On Error Resume Next
    
    iCr = InStr(strInput, vbCr)
    If iCr > 0 Then
        For I = 7 To Len(strInput)
            strTmp = Mid(strInput, I, 1)
            iLen = iLen + 1
            If strTmp = "," Then
                iTmp = Mid(strInput, 7, iLen)
                iLen = 0
                iSignal = iTmp
                Exit For
            End If
        Next I
        iTmp = I
        For I = iTmp To Len(strInput)
            strTmp = Mid(strInput, I, 1)
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
        Next I
        strInput = Right(strInput, Len(strInput) - iCr)
        QueryRSSI = True
    Else
        QueryRSSI = False
    End If
End Function


Public Function GB2Unicode(ByVal strGB As String) As String

    Dim byteA()         As Byte
    
    Dim I               As Integer
    
    Dim strTmpUnicode   As String
    Dim strA            As String
    Dim strB            As String

    On Error GoTo ErrorUnicode
    
    I = LenB(strGB)
    
    ReDim byteA(1 To I)
    
    For I = 1 To LenB(strGB)
        strA = MidB(strGB, I, 1)
        byteA(I) = AscB(strA)
    Next I
    
    '此时已经将strGB转换为Unicode编码，保存在数组byteA()中。
    '下面需要调整顺序并以字符串的形式返回
    strTmpUnicode = ""
    
    For I = 1 To UBound(byteA) Step 2
        strA = Hex(byteA(I))
        If Len(strA) < 2 Then strA = "0" & strA
        strB = Hex(byteA(I + 1))
        If Len(strB) < 2 Then strB = "0" & strB
        strTmpUnicode = strTmpUnicode & strB & strA
    Next I
    
    GB2Unicode = strTmpUnicode
    Exit Function
ErrorUnicode:
    GB2Unicode = ""
End Function

Public Function Unicode2GB(ByVal strUnicode As String) As String

    Dim byteA()     As Byte
    
    Dim I           As Integer
    
    Dim strTmp      As String
    Dim strTmpGB    As String
    
    
    On Error GoTo ErrUnicode2GB
    
    I = Len(strUnicode) / 2
    ReDim byteA(1 To I)
    
    For I = 1 To Len(strUnicode) / 2 Step 2
        strTmp = Mid(strUnicode, I * 2 - 1, 2)
        strTmp = Hex2Dec(strTmp)
        byteA(I + 1) = strTmp
        strTmp = Mid(strUnicode, I * 2 + 1, 2)
        strTmp = Hex2Dec(strTmp)
        byteA(I) = strTmp
    Next I
    
    strTmpGB = ""
    For I = 1 To UBound(byteA)
        strTmp = byteA(I)
        strTmpGB = strTmpGB & ChrB(strTmp)
    Next I
    
    Unicode2GB = strTmpGB
    Exit Function

ErrUnicode2GB:
    'MsgBox "Err=" & Err.Number & ",原因：" & Err.Description
    Unicode2GB = ""
End Function

Public Function Hex2Dec(ByVal strInput As String) As Long
    Dim I       As Integer
    Dim j       As Integer
    Dim iLen    As Integer
    Dim iTmp    As Integer
    
    Dim nRet    As Long
    Dim strTmp  As String
    
    On Error Resume Next
    
    If strInput <> "" Then
        iLen = Len(strInput)
        nRet = 0
        For I = 1 To iLen
            iTmp = Asc(Mid(strInput, I, 1))
            If iTmp >= 48 And iTmp <= 57 Then               '"0" = 48, "9" = 57
                nRet = nRet + (iTmp - 48) * 16 ^ (iLen - I)
            ElseIf iTmp >= 65 And iTmp <= 70 Then           '"A" = 65, "F" = 70
                nRet = nRet + (iTmp - 55) * 16 ^ (iLen - I)
            ElseIf iTmp >= 97 And iTmp <= 102 Then          '"a" = 97, "f" = 102
                nRet = nRet + (iTmp - 87) * 16 ^ (iLen - I)
            Else
                nRet = 0
                Exit For
            End If
        Next I
    End If
    
    Hex2Dec = nRet

End Function


'此函数是将一个字符串中以charRef为分隔符的元素保存到数组MyStr()中
'*********************************************
'参数：
'============================================
'|YourStr：  |  待分隔的字符串
'+-----------+-------------------------------
'|charRef：  |  分隔符号
'+-----------+-------------------------------
'|isNormal： |  如果为假，则表示分隔符可能由
'|           |  多个空格组成，例如Tab符号。
'+-----------+-------------------------------
'|nD：       |  返回值，表示有多少个元素
'+-----------+-------------------------------
'|MyStr()：  |  返回值，保存分隔后的各个元素。
'============================================
'
'**********************************************
Public Function String2Array(ByVal YourStr As String, _
                             ByVal charRef As String, _
                             ByRef nD As Long, _
                             ByRef MyStr() As String, _
                             ByVal isNormal As Boolean) As Boolean

    Dim I           As Long
    Dim j           As Long
    Dim nUBound     As Long
    
    Dim iAsc        As Integer
    
    Dim strChar     As String
    Dim strTmp      As String
    Dim aryTr()     As String
  
    On Error GoTo ErrorDecode

    strChar = ""
    YourStr = Trim(YourStr)     '首先去掉字符串两边的空格
    nUBound = 1
    j = 0
    ReDim aryTr(1 To nUBound)

    If Not isNormal Then
        For I = 1 To Len(YourStr)
            strTmp = Mid(YourStr, I, 1)
            iAsc = Asc(strTmp)
            If iAsc > 122 Or iAsc < 33 Then
                strChar = Mid(YourStr, I - j, j)
                If strChar <> "" Then
                    aryTr(nUBound) = strChar
                    nUBound = nUBound + 1
                    ReDim Preserve aryTr(1 To nUBound)
                End If
                strChar = ""
                j = 0
            Else
                j = j + 1
                If I = Len(YourStr) Then
                    strChar = Mid(YourStr, I - j + 1, j)
                    aryTr(nUBound) = strChar
                End If
            End If
        Next I
        nD = nUBound
        ReDim MyStr(0 To nUBound - 1)
        For I = 1 To nUBound
            MyStr(I - 1) = aryTr(I)
        Next I
        String2Array = True
    Else
        For I = 1 To Len(YourStr)
            strTmp = Mid(YourStr, I, 1)
            If strTmp = charRef Then
                strChar = Mid(YourStr, I - j, j)
                If strChar <> "" Then
                    aryTr(nUBound) = strChar
                    nUBound = nUBound + 1
                    ReDim Preserve aryTr(1 To nUBound)
                End If
                strChar = ""
                j = 0
            Else
                j = j + 1
                If I = Len(YourStr) Then
                    strChar = Mid(YourStr, I - j + 1, j)
                    aryTr(nUBound) = strChar
                End If
            End If
        Next I
        nD = nUBound
        ReDim MyStr(0 To nUBound - 1)
        For I = 1 To nUBound
            MyStr(I - 1) = aryTr(I)
        Next I
        String2Array = True
    End If

    Exit Function

ErrorDecode:
    'MsgBox Err.Number & ":" & Err.Description
    String2Array = False
End Function


Public Function ASCII2Char(ByVal strAsc As String) As String

    Dim I       As Integer
    Dim j       As Integer
    
    Dim strTmp  As String
    Dim strTmpA As String
    Dim strTmpB As String

    On Error Resume Next
    j = Len(strAsc)
    strTmpB = ""

    For I = 1 To j
        strTmpA = Mid(strAsc, I, 1)
        If strTmpA <> " " Then strTmpB = strTmpB & strTmpA
    Next I

    j = Len(strTmpB)

    strTmp = ""
    For I = 1 To j Step 2
        strTmpA = Mid(strTmpB, I, 2)
        strTmp = strTmp & ChrB(Hex2Dec(strTmpA))
    Next I

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

