Attribute VB_Name = "mdlInptCk"


Public Sub OnlyIntCK(ByRef KeyAscii As Integer)
'���ܣ���������������
'��TEXTBOX��KEYPRESSʱ����ʹ�ã���KeyAscII��Ϊ�������뼴��

    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub


Public Sub OnlyDblCK(ByRef KeyAscii As Integer, txtContent As String)
'���ܣ���������double���͵���������Ҫ���밴�µ�KeyAscII �� ��ǰ���������

    If InStr(1, "1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
    
    If (InStr(1, txtContent, ".") > 0 And KeyAscii = 46) Or (txtContent = "" And KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub


Public Sub OnlyStrCK(ByRef KeyAscii As Integer, ParamArray arrChr() As Variant)
'���ܣ������������ֺ���ĸ,��ָ���ַ�
'��Ҫָ�����ַ���KeyAscII ��ͨ��������ʽ���δ���
'֧��ճ�����ƣ���ݼ�KeyAscii�� CRTRL+C = 3 ,CTRL+V  =22
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
'���ܣ��������������ַ������ֺ���ĸ,��ָ���ַ�
'��Ҫָ�����ַ���KeyAscII ��ͨ��������ʽ���δ���
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

