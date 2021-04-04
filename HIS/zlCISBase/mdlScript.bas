Attribute VB_Name = "mdlScript"
Option Explicit

Public Function CheckRule(ByVal strFormula As String, ByVal strItems As String) As Boolean
    '提交时要放mdlCISBase中
    '功能：检证普通公式是否合法
    'strFormula : 公式
    'strItems :可用的项目，用于检验公式中的非法项目。格式为",[项目1],[项目2],...,[项目n],"
     
    Dim strTmp As String     '存解析后的公式。
    Dim strLine As String    '存输入的公式，分析一段，删除一段,删除完，则分析完
    Dim strErrItem As String '存无效的元素串，用于错误提示。
    'Dim varReturn As Variant '存最后的计算结果
    
    Dim i As Integer         '临时变量，每发现一个元素则加1，值作为元素的值代入公式，用于模拟计算。
    Dim strItem As String    '临时变量，存提取出的单个元素。
    Dim lngLength As Long    '临时变量，存元素的长度，用于提取元素。
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
        MsgBox "请从公式中去掉以下错误项目！" & vbNewLine & Mid(strErrItem, 1, Len(strErrItem) - 1), vbExclamation, gstrSysName
        Exit Function
    End If
    strTmp = strTmp & strLine
    
    Set rsTmp = zlDatabase.OpenSQLRecord("Select 1 As sequel  From Dual Where " & strTmp, "CheckRule")
    CheckRule = True
    Exit Function
ErrHandle:
    CheckRule = False
    MsgBox "计算失败，请检查公式！", vbExclamation, gstrSysName
End Function

Public Function CheckEspecial(ByVal strFormula As String, ByVal strItems As String) As Boolean
    '功能：检证普通公式是否合法
    'strFormula : 公式
    'strItems   : 可用的项目，用于检验公式中的非法项目。格式为",[项目1],[项目2],...,[项目n],"

    Dim strTmp As String     '存解析后的公式。
    Dim strLine As String    '存输入的公式，分析一段，删除一段,删除完，则分析完
    Dim strErrItem As String '存无效的公式串，用于错误提示。
    'Dim varReturn As Variant '存最后的计算结果
    
    Dim i As Integer         '临时变量，每发现一个元素则加1，值作为元素的值代入公式，用于模拟计算。
    Dim str公式 As String    '临时变量，存提取出的单个公式。
    Dim lngLength As Long    '临时变量，存公式的长度，用于提取公式。
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    strLine = UCase(strFormula):    strTmp = "":    strErrItem = ""
    
    Do While strLine Like "*{*}*"
        strTmp = strTmp & Mid(strLine, 1, InStr(strLine, "{") - 1)
        lngLength = InStr(strLine, "}") - InStr(strLine, "{") + 1
        str公式 = Mid(strLine, InStr(strLine, "{"), lngLength)
        If str公式 Like "{A:*|*}" Then
            '验证A类规则
            If Check_EspecialA(str公式) Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str公式 & ","
            End If
        ElseIf str公式 Like "{B:*[[]*[]]*}" Then
            '验证B类规则
            If Check_EspecialB(str公式, strItems) Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str公式 & ","
            End If
        ElseIf str公式 Like "{C:*|*}" Then
            '验证C类规则
            If Check_EspecialC(str公式, strItems) Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str公式 & ","
            End If
        ElseIf str公式 Like "{D:*}" Then
            '验证D类规则
            If GenFormula("D", strItems, Mid(str公式, 4, Len(str公式) - 4)) <> "" Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str公式 & ","
            End If
        ElseIf str公式 Like "{E:*[[]*[]]*}" Then
            '验证E类规则
            If Check_EspecialE(str公式, strItems) Then
                strTmp = strTmp & " 1=1 "
            Else
                strErrItem = strErrItem & str公式 & ","
            End If
        Else
            strErrItem = strErrItem & str公式 & ","
        End If
        strLine = Mid(strLine, InStr(strLine, "}") + 1)
    Loop
    
    If strErrItem <> "" Then
        MsgBox "请从公式中去掉以下错误项目！" & vbNewLine & Mid(strErrItem, 1, Len(strErrItem) - 1), vbExclamation, gstrSysName
        Exit Function
    End If
    
    strTmp = strTmp & strLine
    strLine = Replace(Replace(Replace(Replace(strTmp, " 1=1 ", ""), "OR", ""), "AND", ""), " ", "")
    If strLine <> "" Then
        MsgBox "公式中不能出现以下字符，请修改！" & vbNewLine & strLine, vbExclamation, gstrSysName
        Exit Function
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord("Select 1 As sequel  From Dual Where " & strTmp, "CheckEspecial")
    CheckEspecial = True
    Exit Function
ErrHandle:
    CheckEspecial = False
    MsgBox "计算失败，请检查公式！", vbExclamation, gstrSysName
End Function


Private Function Check_EspecialA(ByVal str公式 As String) As Boolean
    '功能：检查特殊规则 A 的语法。
    '分解公式
    'str公式 :待验证的A公式
    
    Dim strP0 As String, strP1 As String, strP2 As String
    
    strP0 = Replace(Split(str公式, "|")(0), "{A:", "")
    strP1 = Replace(Split(str公式, "|")(1), "}", "")
    
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

Private Function Check_EspecialB(ByVal str公式 As String, ByVal str项目 As String) As Boolean
    '功能：检查特殊规则 A 的语法。
    '分解公式
    'str公式 :待验证的B公式
    'str项目 :公式中可用的项目
    
    Dim strP As String
    strP = Replace(Replace(str公式, "{B:", ""), "}", "")
    If GenFormula("B", str项目, strP) <> "" Then
        Check_EspecialB = True: Exit Function
    Else
        Check_EspecialB = False
    End If
    
End Function

Private Function Check_EspecialE(ByVal str公式 As String, ByVal str项目 As String) As Boolean
    '功能：检查特殊规则 A 的语法。
    '分解公式
    'str公式 :待验证的E公式
    'str项目 :公式中可用的项目
    
    Dim strP As String
    strP = Replace(Replace(str公式, "{E:", ""), "}", "")
    If GenFormula("E", str项目, strP) <> "" Then
        Check_EspecialE = True: Exit Function
    Else
        Check_EspecialE = False
    End If
    
End Function
Private Function Check_EspecialC(ByVal str公式 As String, ByVal str项目 As String) As Boolean
    '功能：检查特殊规则 C 的语法。
    '分解公式
    'str公式 :待验证的B公式
    'str项目 :公式中可用的项目
    
    Dim strP0 As String, strP1 As String, strP2 As String
    
    strP0 = Replace(Split(str公式, "|")(0), "{C:", "")
    strP1 = Replace(Split(str公式, "|")(1), "}", "")
    
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
        If GenFormula("C", str项目, strP0, strP1, strP2) <> "" Then
            Check_EspecialC = True: Exit Function
        Else
            Check_EspecialC = False
        End If
    End If
End Function

Public Function GenFormula(ByVal strType As String, ByVal strItem As String, ParamArray Parameter()) As String
    '生成特殊规则公式
    Dim varReturn As Variant
    On Error GoTo ErrHandle
    Select Case UCase(strType)
        Case "A"
            If UBound(Parameter()) = 2 Then
                If Parameter(0) = "" Or Parameter(2) = "" Then
                    GenFormula = ""
                    MsgBox "参数不能为空！", vbExclamation, gstrSysName
                    Exit Function
                End If
                
                If Not IsNumeric(Parameter(2)) Then
                    GenFormula = ""
                    MsgBox "项目个数的要为数字！", vbExclamation, gstrSysName
                    Exit Function
                End If
                
                GenFormula = "{A:" & Parameter(0) & "|" & Parameter(1) & Parameter(2) & "}"
            Else
                GenFormula = ""
                MsgBox "参数个数不正确！", vbExclamation, gstrSysName
                Exit Function
            End If
        Case "B"
            If UBound(Parameter()) = 0 Then
                If CheckRule(CStr(Parameter(0)), strItem) Then
                    GenFormula = "{B:" & CStr(Parameter(0)) & "}"
                End If
            Else
                GenFormula = ""
                MsgBox "参数个数不正确！", vbExclamation, gstrSysName
                Exit Function
            End If
        Case "C"
            If UBound(Parameter()) = 2 Then
                If Trim(CStr(Parameter(0))) = "" Then
                    GenFormula = ""
                    MsgBox "参数不能为空！", vbExclamation, gstrSysName
                    Exit Function
                End If
                
                If Not IsNumeric(Parameter(2)) Then
                    GenFormula = ""
                    MsgBox "项目值要为数字！", vbExclamation, gstrSysName
                    Exit Function
                End If
                
                If CheckRule(Replace(CStr(Parameter(0)), ",", " + ") & " > 0", strItem) Then
                    GenFormula = "{C:" & CStr(Parameter(0)) & "|" & Parameter(1) & Parameter(2) & "}"
                End If
            Else
                GenFormula = ""
                MsgBox "参数个数不正确！", vbExclamation, gstrSysName
                Exit Function
            End If
        Case "D"
            If UBound(Parameter()) = 0 Then
                If InStr(",漏项检查,多项检查,漏项多项检查,", "," & Trim(CStr(Parameter(0))) & ",") <= 0 Then
                    GenFormula = ""
                    MsgBox "参数值错误，请检查！", vbExclamation, gstrSysName
                    Exit Function
                End If
                GenFormula = "{D:" & Trim(CStr(Parameter(0))) & "}"
            Else
                GenFormula = ""
                MsgBox "参数个数不正确！", vbExclamation, gstrSysName
                Exit Function
            End If
        Case "E"
            If UBound(Parameter()) = 0 Then
                If CheckRule(CStr(Parameter(0)), strItem) Then
                    GenFormula = "{E:" & CStr(Parameter(0)) & "}"
                End If
            Else
                GenFormula = ""
                MsgBox "参数个数不正确！", vbExclamation, gstrSysName
                Exit Function
            End If
        Case Else
            GenFormula = ""
    End Select
    Exit Function
ErrHandle:
    GenFormula = ""
    MsgBox "公式不正确，请检查！", vbExclamation, gstrSysName
    
End Function
