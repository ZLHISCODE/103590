Attribute VB_Name = "mdlFun"
Option Explicit

Public Function CharToNumber(ByVal strchar As String) As Integer
        Select Case strchar
            Case "一"
                CharToNumber = 1
            Case "二"
                CharToNumber = 2
            Case "三"
                CharToNumber = 3
            Case "四"
                CharToNumber = 4
            Case "五"
                CharToNumber = 5
            Case "六"
                CharToNumber = 6
            Case "七"
                CharToNumber = 7
            Case "八"
                CharToNumber = 8
            Case "九"
                CharToNumber = 9
            Case "零"
                CharToNumber = 0
            Case Else
                CharToNumber = -1
        End Select
End Function

Public Function NumberToChar(ByVal lngNumber As Long) As String
        Select Case lngNumber
            Case 1
                NumberToChar = "一"
            Case 2
                NumberToChar = "二"
            Case 3
                NumberToChar = "三"
            Case 4
                NumberToChar = "四"
            Case 5
                NumberToChar = "五"
            Case 6
                NumberToChar = "六"
            Case 7
                NumberToChar = "七"
            Case 8
                NumberToChar = "八"
            Case 9
                NumberToChar = "九"
            Case 0
                NumberToChar = "零"
            Case Else
                NumberToChar = ""
        End Select
End Function

Public Function UnitToChar(ByVal strchar As String) As String
        Select Case strchar
            Case 10
                UnitToChar = "十"
            Case 100
                UnitToChar = "百"
            Case 1000
                UnitToChar = "千"
            Case Else
                UnitToChar = ""
        End Select
End Function

Public Function CharToUnit(ByVal strchar As String) As Integer
        Select Case strchar
            Case "十"
                CharToUnit = 10
            Case "百"
                CharToUnit = 100
            Case "千"
                CharToUnit = 1000
            Case Else
                CharToUnit = ""
        End Select
End Function

Public Function StrToNumber(ByVal strNumber As String) As Long
'功能：支持一万以内的汉字数字转化为阿拉伯数字
    Dim arrChar As Variant
    Dim i As Long
    Dim lngResult As Long
    
    arrChar = StrToChar(strNumber)
    For i = LBound(arrChar) To UBound(arrChar)
        If CharToNumber(arrChar(i)) <> -1 Then
            arrChar(i) = CharToNumber(arrChar(i)) & ""
        End If
    Next
    
    For i = LBound(arrChar) To UBound(arrChar)
        If i = 0 And Not IsNumeric(arrChar(i)) Then '如十三
            lngResult = 10
        End If
        If IsNumeric(arrChar(i)) And Val(arrChar(i)) <> 0 Then
            If i + 1 <= UBound(arrChar) Then
                lngResult = lngResult + Val(arrChar(i)) * CharToUnit(arrChar(i + 1))
            Else
                lngResult = lngResult + Val(arrChar(i))
            End If
        End If
    Next
    
    StrToNumber = lngResult
End Function

Public Function NumberToStr(ByVal lngNumber As Long) As String
'功能：支持一万以内的阿拉伯数字转化为汉字数字
    Dim arrChar As Variant
    Dim i As Long, j As Long
    Dim strResult As String
    Dim lng单位 As Long
    
    arrChar = StrToChar(CStr(lngNumber))
    For i = LBound(arrChar) To UBound(arrChar)
        arrChar(i) = NumberToChar(arrChar(i)) & ""
    Next
    
    For i = LBound(arrChar) To UBound(arrChar)
        lng单位 = 10 ^ (UBound(arrChar) - i)
        If arrChar(i) <> "零" Then
            If i + 1 <= UBound(arrChar) Then
                For j = i + 1 To UBound(arrChar)
                    If arrChar(j) <> "零" Then
                        If i = j - 1 Then
                            strResult = strResult & arrChar(i) & UnitToChar(lng单位)
                        Else
                            strResult = strResult & arrChar(i) & UnitToChar(lng单位) & "零"
                        End If
                        i = j - 1
                        Exit For
                    ElseIf i + 1 = UBound(arrChar) And arrChar(j) = "零" Then
                        strResult = strResult & arrChar(i) & UnitToChar(lng单位)
                    End If
                Next
            Else
                strResult = strResult & arrChar(i) & UnitToChar(lng单位)
            End If
        End If
    Next
    If UBound(arrChar) = 1 And arrChar(0) = "一" Then
        strResult = "十" & IIf(arrChar(1) <> "零", arrChar(1), "")
    End If
    NumberToStr = strResult
End Function

Public Function StrToChar(ByVal strInput As String) As Variant
    Dim i As Long
    Dim lngLen As Long
    Dim arrChar() As String
    
    lngLen = Len(strInput)
    ReDim arrChar(0 To lngLen - 1)
    For i = 0 To Len(strInput) - 1
        arrChar(i) = Mid(strInput, i + 1, 1)
    Next
    StrToChar = arrChar
End Function

Public Function RPAD(ByVal strInput As String, ByVal strFill As String, ByVal lngLen As String)
'功能：右填充至制定长度
    Dim strResult As String

    strResult = Trim(Replace(strInput, Chr(13) & Chr(7), ""))
    If Len(strFill) > 0 Then
        Do
            strResult = strResult & strFill
        Loop While Len(strResult) + Len(strFill) < lngLen
    End If
    
    RPAD = strResult
End Function


