Attribute VB_Name = "mdlScript"
Option Explicit

Public Function CheckRule(ByVal strFormula As String, ByVal strItems As String) As Boolean
    '�ύʱҪ��mdlCISBase��
    '���ܣ���֤��ͨ��ʽ�Ƿ�Ϸ�
    'strFormula : ��ʽ
    'strItems :���õ���Ŀ�����ڼ��鹫ʽ�еķǷ���Ŀ����ʽΪ",[��Ŀ1],[��Ŀ2],...,[��Ŀn],"
     
    Dim strTmp As String     '�������Ĺ�ʽ��
    Dim strLine As String    '������Ĺ�ʽ������һ�Σ�ɾ��һ��,ɾ���꣬�������
    Dim strErrItem As String '����Ч��Ԫ�ش������ڴ�����ʾ��
    'Dim varReturn As Variant '�����ļ�����
    
    Dim i As Integer         '��ʱ������ÿ����һ��Ԫ�����1��ֵ��ΪԪ�ص�ֵ���빫ʽ������ģ����㡣
    Dim strItem As String    '��ʱ����������ȡ���ĵ���Ԫ�ء�
    Dim lngLength As Long    '��ʱ��������Ԫ�صĳ��ȣ�������ȡԪ�ء�
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    strLine = strFormula
    strTmp = ""
    strErrItem = ""
    Do While strLine Like "*[[]*[]]*"
        strTmp = strTmp & Mid(strLine, 1, InStr(strLine, "[") - 1) & "(" & i & "+ 1)"
        lngLength = InStr(strLine, "]") - InStr(strLine, "[") + 1
        strItem = Mid(strLine, InStr(strLine, "["), lngLength)
        If InStr(UCase(strItems), "," & UCase(strItem) & ",") <= 0 Then
            strErrItem = strErrItem & strItem & ","
        End If
        strLine = Mid(strLine, InStr(strLine, "]") + 1)
        i = i + 1
    Loop
    If InStr(UCase(strLine), UCase("null")) > 0 Or InStr(strLine, "''") > 0 Then
        GoTo ErrHandle
    End If
    If strErrItem <> "" Then
        MsgBox "��ӹ�ʽ��ȥ�����´�����Ŀ��" & vbNewLine & Mid(strErrItem, 1, Len(strErrItem) - 1), vbExclamation, gstrSysName
        Exit Function
    End If
    strTmp = strTmp & strLine
    
    Set rsTmp = zlDatabase.OpenSQLRecord("Select 1 As sequel  From Dual Where " & strTmp, "CheckRule")
    CheckRule = True
    Exit Function
ErrHandle:
    CheckRule = False
    MsgBox "����ʧ�ܣ����鹫ʽ��", vbExclamation, gstrSysName
End Function

Public Function CheckEspecial(ByVal strFormula As String, ByVal strItems As String) As Boolean
    '���ܣ���֤��ͨ��ʽ�Ƿ�Ϸ�
    'strFormula : ��ʽ
    'strItems   : ���õ���Ŀ�����ڼ��鹫ʽ�еķǷ���Ŀ����ʽΪ",[��Ŀ1],[��Ŀ2],...,[��Ŀn],"

    Dim strTmp As String     '�������Ĺ�ʽ��
    Dim strLine As String    '������Ĺ�ʽ������һ�Σ�ɾ��һ��,ɾ���꣬�������
    Dim strErrItem As String '����Ч�Ĺ�ʽ�������ڴ�����ʾ��
    'Dim varReturn As Variant '�����ļ�����
    
    Dim i As Integer         '��ʱ������ÿ����һ��Ԫ�����1��ֵ��ΪԪ�ص�ֵ���빫ʽ������ģ����㡣
    Dim str��ʽ As String    '��ʱ����������ȡ���ĵ�����ʽ��
    Dim lngLength As Long    '��ʱ�������湫ʽ�ĳ��ȣ�������ȡ��ʽ��
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    strLine = UCase(strFormula):    strTmp = "":    strErrItem = ""
    
    Do While strLine Like "*{*}*"
        strTmp = strTmp & Mid(strLine, 1, InStr(strLine, "{") - 1)
        lngLength = InStr(strLine, "}") - InStr(strLine, "{") + 1
        str��ʽ = Mid(strLine, InStr(strLine, "{"), lngLength)
        If str��ʽ Like "{A:*|*}" Then
            '��֤A�����
            If Check_EspecialA(str��ʽ) Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str��ʽ & ","
            End If
        ElseIf str��ʽ Like "{B:*[[]*[]]*}" Then
            '��֤B�����
            If Check_EspecialB(str��ʽ, strItems) Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str��ʽ & ","
            End If
        ElseIf str��ʽ Like "{C:*|*}" Then
            '��֤C�����
            If Check_EspecialC(str��ʽ, strItems) Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str��ʽ & ","
            End If
        ElseIf str��ʽ Like "{D:*}" Then
            '��֤D�����
            If GenFormula("D", strItems, Mid(str��ʽ, 4, Len(str��ʽ) - 4)) <> "" Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str��ʽ & ","
            End If
        ElseIf str��ʽ Like "{E:*[[]*[]]*}" Then
            '��֤E�����
            If Check_EspecialE(str��ʽ, strItems) Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str��ʽ & ","
            End If
        Else
            strErrItem = strErrItem & str��ʽ & ","
        End If
        strLine = Mid(strLine, InStr(strLine, "}") + 1)
    Loop
    
    If strErrItem <> "" Then
        MsgBox "��ӹ�ʽ��ȥ�����´�����Ŀ��" & vbNewLine & Mid(strErrItem, 1, Len(strErrItem) - 1), vbExclamation, gstrSysName
        Exit Function
    End If
    
    strTmp = strTmp & strLine
    strLine = Replace(Replace(Replace(Replace(strTmp, " 1=1 ", ""), "OR", ""), "AND", ""), " ", "")
    If strLine <> "" Then
        MsgBox "��ʽ�в��ܳ��������ַ������޸ģ�" & vbNewLine & strLine, vbExclamation, gstrSysName
        Exit Function
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord("Select 1 As sequel  From Dual Where " & strTmp, "CheckEspecial")
    CheckEspecial = True
    Exit Function
ErrHandle:
    CheckEspecial = False
    MsgBox "����ʧ�ܣ����鹫ʽ��", vbExclamation, gstrSysName
End Function


Private Function Check_EspecialA(ByVal str��ʽ As String) As Boolean
    '���ܣ����������� A ���﷨��
    '�ֽ⹫ʽ
    'str��ʽ :����֤��A��ʽ
    
    Dim strP0 As String, strP1 As String, strP2 As String
    
    strP0 = Replace(Split(str��ʽ, "|")(0), "{A:", "")
    strP1 = Replace(Split(str��ʽ, "|")(1), "}", "")
    
    If IsNumeric(Mid(strP1, 2)) Then
        strP2 = Mid(strP1, 2)
        strP1 = Mid(strP1, 1, 1)
    ElseIf IsNumeric(Mid(strP1, 3)) Then
        strP2 = Mid(strP1, 3)
        strP1 = Mid(strP1, 1, 2)
    Else
        strP1 = ""
        strP2 = ""
    End If
    
    If strP0 = "" Or strP1 = "" Or strP2 = "" Then
        Check_EspecialA = False
    ElseIf InStr(",=,>,<,<>,>=,<=,", "," & strP1 & ",") <= 0 Then
        Check_EspecialA = False
    Else
        If GenFormula("A", "", strP0, strP1, strP2) <> "" Then
            Check_EspecialA = True: Exit Function
        Else
            Check_EspecialA = False
        End If
    End If

End Function

Private Function Check_EspecialB(ByVal str��ʽ As String, ByVal str��Ŀ As String) As Boolean
    '���ܣ����������� A ���﷨��
    '�ֽ⹫ʽ
    'str��ʽ :����֤��B��ʽ
    'str��Ŀ :��ʽ�п��õ���Ŀ
    
    Dim strP As String
    strP = Replace(Replace(str��ʽ, "{B:", ""), "}", "")
    If GenFormula("B", str��Ŀ, strP) <> "" Then
        Check_EspecialB = True: Exit Function
    Else
        Check_EspecialB = False
    End If
    
End Function

Private Function Check_EspecialE(ByVal str��ʽ As String, ByVal str��Ŀ As String) As Boolean
    '���ܣ����������� A ���﷨��
    '�ֽ⹫ʽ
    'str��ʽ :����֤��E��ʽ
    'str��Ŀ :��ʽ�п��õ���Ŀ
    
    Dim strP As String
    strP = Replace(Replace(str��ʽ, "{E:", ""), "}", "")
    If GenFormula("E", str��Ŀ, strP) <> "" Then
        Check_EspecialE = True: Exit Function
    Else
        Check_EspecialE = False
    End If
    
End Function
Private Function Check_EspecialC(ByVal str��ʽ As String, ByVal str��Ŀ As String) As Boolean
    '���ܣ����������� C ���﷨��
    '�ֽ⹫ʽ
    'str��ʽ :����֤��B��ʽ
    'str��Ŀ :��ʽ�п��õ���Ŀ
    
    Dim strP0 As String, strP1 As String, strP2 As String
    
    strP0 = Replace(Split(str��ʽ, "|")(0), "{C:", "")
    strP1 = Replace(Split(str��ʽ, "|")(1), "}", "")
    
    If IsNumeric(Mid(strP1, 2)) Then
        strP2 = Mid(strP1, 2)
        strP1 = Mid(strP1, 1, 1)
    ElseIf IsNumeric(Mid(strP1, 3)) Then
        strP2 = Mid(strP1, 3)
        strP1 = Mid(strP1, 1, 2)
    Else
        strP1 = ""
        strP2 = ""
    End If
    
    If strP0 = "" Or strP1 = "" Or strP2 = "" Then
        Check_EspecialC = False
    ElseIf InStr(",=,>,<,<>,>=,<=,", "," & strP1 & ",") <= 0 Then
        Check_EspecialC = False
    Else
        If GenFormula("C", str��Ŀ, strP0, strP1, strP2) <> "" Then
            Check_EspecialC = True: Exit Function
        Else
            Check_EspecialC = False
        End If
    End If
End Function

Public Function GenFormula(ByVal strType As String, ByVal strItem As String, ParamArray Parameter()) As String
    '�����������ʽ
    Dim varReturn As Variant
    On Error GoTo ErrHandle
    Select Case UCase(strType)
        Case "A"
            If UBound(Parameter()) = 2 Then
                If Parameter(0) = "" Or Parameter(2) = "" Then
                    GenFormula = ""
                    MsgBox "��������Ϊ�գ�", vbExclamation, gstrSysName
                    Exit Function
                End If
                
                If Not IsNumeric(Parameter(2)) Then
                    GenFormula = ""
                    MsgBox "��Ŀ������ҪΪ���֣�", vbExclamation, gstrSysName
                    Exit Function
                End If
                
                GenFormula = "{A:" & Parameter(0) & "|" & Parameter(1) & Parameter(2) & "}"
            Else
                GenFormula = ""
                MsgBox "������������ȷ��", vbExclamation, gstrSysName
                Exit Function
            End If
        Case "B"
            If UBound(Parameter()) = 0 Then
                If CheckRule(CStr(Parameter(0)), strItem) Then
                    GenFormula = "{B:" & CStr(Parameter(0)) & "}"
                End If
            Else
                GenFormula = ""
                MsgBox "������������ȷ��", vbExclamation, gstrSysName
                Exit Function
            End If
        Case "C"
            If UBound(Parameter()) = 2 Then
                If Trim(CStr(Parameter(0))) = "" Then
                    GenFormula = ""
                    MsgBox "��������Ϊ�գ�", vbExclamation, gstrSysName
                    Exit Function
                End If
                
                If Not IsNumeric(Parameter(2)) Then
                    GenFormula = ""
                    MsgBox "��ĿֵҪΪ���֣�", vbExclamation, gstrSysName
                    Exit Function
                End If
                
                If CheckRule(Replace(CStr(Parameter(0)), ",", " + ") & " > 0", strItem) Then
                    GenFormula = "{C:" & CStr(Parameter(0)) & "|" & Parameter(1) & Parameter(2) & "}"
                End If
            Else
                GenFormula = ""
                MsgBox "������������ȷ��", vbExclamation, gstrSysName
                Exit Function
            End If
        Case "D"
            If UBound(Parameter()) = 0 Then
                If InStr(",©����,������,©�������,", "," & Trim(CStr(Parameter(0))) & ",") <= 0 Then
                    GenFormula = ""
                    MsgBox "����ֵ�������飡", vbExclamation, gstrSysName
                    Exit Function
                End If
                GenFormula = "{D:" & Trim(CStr(Parameter(0))) & "}"
            Else
                GenFormula = ""
                MsgBox "������������ȷ��", vbExclamation, gstrSysName
                Exit Function
            End If
        Case "E"
            If UBound(Parameter()) = 0 Then
                If CheckRule(CStr(Parameter(0)), strItem) Then
                    GenFormula = "{E:" & CStr(Parameter(0)) & "}"
                End If
            Else
                GenFormula = ""
                MsgBox "������������ȷ��", vbExclamation, gstrSysName
                Exit Function
            End If
        Case Else
            GenFormula = ""
    End Select
    Exit Function
ErrHandle:
    GenFormula = ""
    MsgBox "��ʽ����ȷ�����飡", vbExclamation, gstrSysName
    
End Function
