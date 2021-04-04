Attribute VB_Name = "mdlFun"
Option Explicit

Public Function CharToNumber(ByVal strchar As String) As Integer
        Select Case strchar
            Case "һ"
                CharToNumber = 1
            Case "��"
                CharToNumber = 2
            Case "��"
                CharToNumber = 3
            Case "��"
                CharToNumber = 4
            Case "��"
                CharToNumber = 5
            Case "��"
                CharToNumber = 6
            Case "��"
                CharToNumber = 7
            Case "��"
                CharToNumber = 8
            Case "��"
                CharToNumber = 9
            Case "��"
                CharToNumber = 0
            Case Else
                CharToNumber = -1
        End Select
End Function

Public Function NumberToChar(ByVal lngNumber As Long) As String
        Select Case lngNumber
            Case 1
                NumberToChar = "һ"
            Case 2
                NumberToChar = "��"
            Case 3
                NumberToChar = "��"
            Case 4
                NumberToChar = "��"
            Case 5
                NumberToChar = "��"
            Case 6
                NumberToChar = "��"
            Case 7
                NumberToChar = "��"
            Case 8
                NumberToChar = "��"
            Case 9
                NumberToChar = "��"
            Case 0
                NumberToChar = "��"
            Case Else
                NumberToChar = ""
        End Select
End Function

Public Function UnitToChar(ByVal strchar As String) As String
        Select Case strchar
            Case 10
                UnitToChar = "ʮ"
            Case 100
                UnitToChar = "��"
            Case 1000
                UnitToChar = "ǧ"
            Case Else
                UnitToChar = ""
        End Select
End Function

Public Function CharToUnit(ByVal strchar As String) As Integer
        Select Case strchar
            Case "ʮ"
                CharToUnit = 10
            Case "��"
                CharToUnit = 100
            Case "ǧ"
                CharToUnit = 1000
            Case Else
                CharToUnit = ""
        End Select
End Function

Public Function StrToNumber(ByVal strNumber As String) As Long
'���ܣ�֧��һ�����ڵĺ�������ת��Ϊ����������
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
        If i = 0 And Not IsNumeric(arrChar(i)) Then '��ʮ��
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
'���ܣ�֧��һ�����ڵİ���������ת��Ϊ��������
    Dim arrChar As Variant
    Dim i As Long, j As Long
    Dim strResult As String
    Dim lng��λ As Long
    
    arrChar = StrToChar(CStr(lngNumber))
    For i = LBound(arrChar) To UBound(arrChar)
        arrChar(i) = NumberToChar(arrChar(i)) & ""
    Next
    
    For i = LBound(arrChar) To UBound(arrChar)
        lng��λ = 10 ^ (UBound(arrChar) - i)
        If arrChar(i) <> "��" Then
            If i + 1 <= UBound(arrChar) Then
                For j = i + 1 To UBound(arrChar)
                    If arrChar(j) <> "��" Then
                        If i = j - 1 Then
                            strResult = strResult & arrChar(i) & UnitToChar(lng��λ)
                        Else
                            strResult = strResult & arrChar(i) & UnitToChar(lng��λ) & "��"
                        End If
                        i = j - 1
                        Exit For
                    ElseIf i + 1 = UBound(arrChar) And arrChar(j) = "��" Then
                        strResult = strResult & arrChar(i) & UnitToChar(lng��λ)
                    End If
                Next
            Else
                strResult = strResult & arrChar(i) & UnitToChar(lng��λ)
            End If
        End If
    Next
    If UBound(arrChar) = 1 And arrChar(0) = "һ" Then
        strResult = "ʮ" & IIf(arrChar(1) <> "��", arrChar(1), "")
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
'���ܣ���������ƶ�����
    Dim strResult As String

    strResult = Trim(Replace(strInput, Chr(13) & Chr(7), ""))
    If Len(strFill) > 0 Then
        Do
            strResult = strResult & strFill
        Loop While Len(strResult) + Len(strFill) < lngLen
    End If
    
    RPAD = strResult
End Function


