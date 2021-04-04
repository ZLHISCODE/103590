Attribute VB_Name = "mdlInptCk"


Public Sub OnlyIntCK(ByRef KeyAscii As Integer)
'功能：仅能输入正整数
'在TEXTBOX的KEYPRESS时间中使用，将KeyAscII作为参数传入即可

    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub


Public Sub OnlyDblCK(ByRef KeyAscii As Integer, txtContent As String)
'功能：仅能输入double类型的正数，需要传入按下的KeyAscII 和 当前输入框内容

    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
    
    If (InStr(1, txtContent, ".") > 0 And KeyAscii = 46) Or (txtContent = "" And KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub


Public Sub OnlyStrCK(ByRef KeyAscii As Integer, ParamArray arrChr() As Variant)
'功能：仅能输入数字和字母,和指定字符
'需要指定的字符在KeyAscII 后通过传参形式依次传入
'支持粘贴复制，快捷键KeyAscii： CRTRL+C = 3 ,CTRL+V  =22
     Dim intIdx As Integer, intFlag As Integer
    
    intFlag = 1
    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) = 0 Then
            intFlag = 0
        End If
    End If
    
    For intIdx = LBound(arrChr) To UBound(arrChr)
        If Chr(KeyAscii) = arrChr(intIdx) Then
            intFlag = 1
        End If
    Next
    
    If intFlag = 0 Then
        KeyAscii = 0
    End If
    
End Sub

Public Sub OnlyStrChnCK(ByRef KeyAscii As Integer, ParamArray arrChr() As Variant)
'功能：仅能输入中文字符、数字和字母,和指定字符
'需要指定的字符在KeyAscII 后通过传参形式依次传入
    Dim intIdx As Integer, intFlag As Integer
    
    intFlag = 1
    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) = 0 Then
            intFlag = 0
        End If
    End If
    
    If KeyAscii < 0 Then
        intFlag = 1
    End If
    
    For intIdx = LBound(arrChr) To UBound(arrChr)
        If Chr(KeyAscii) = arrChr(intIdx) Then
            intFlag = 1
        End If
    Next
    
    If intFlag = 0 Then
        KeyAscii = 0
    End If
    
End Sub

