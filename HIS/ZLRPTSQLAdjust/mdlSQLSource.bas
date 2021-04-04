Attribute VB_Name = "mdlSQLSource"

Option Explicit
Public Enum gEditType
     g新增 = 0
     g修改 = 1
     g查看 = 4
End Enum

'------------------------------------------------------------------------------------
Public grsObject As ADODB.Recordset '当前用户所具有Select权限的对象集(用于向导或发布)
'------------------------------------------------------------------------------------

Public gblnRunLog As Boolean '是否记录使用日志
Public gblnErrLog As Boolean '是否记录运行错误


Public glngKeyHook As Long


Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'控制TAB键的函数
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Const WH_KEYBOARD = 2
Public Const HC_ACTION = 0
Public Const HC_NOREMOVE = 3


Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = -4
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息

Public gblnOK As Boolean
Public Type CustomPar
    组名 As String
    值列表 As String
    分类SQL As String
    明细SQL As String
    分类字段 As String
    明细字段 As String
    对象 As String
    格式 As Byte
End Type
Public lngTXTProc As Long '保存默认的消息函数的地址

Public Const GSTR_SBC = "（＋－＊／＝＜＞）！：１２３４５６７８９０ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ；，。？｜％＃"
Public Const GSTR_DBC = "(+-*/=<>)!:1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcedfghijklmnopqrstuvwxyz;,.?|%#"

Public Function TLen(str As String) As Long
'功能：返回字符串的真实长度
    TLen = LenB(StrConv(str, vbFromUnicode))
End Function

Public Function RemoveNote(ByVal strSQL As String) As String
'功能：移除SQL语句中的注释
'说明：只支持移除整行的注释
    Dim strTmp As String, i As Integer
    Dim arrLine() As String
    
    strSQL = Replace(strSQL, vbTab, " ")
    strSQL = Replace(strSQL, vbLf, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr & vbCr, vbCr)
    strSQL = Replace(strSQL, vbCr, vbCrLf)
    arrLine = Split(strSQL, vbCrLf)
    
    For i = 0 To UBound(arrLine)
        If Not Trim(arrLine(i)) Like "--*" Then
            RemoveNote = RemoveNote & vbCrLf & arrLine(i)
        End If
    Next
    RemoveNote = Mid(RemoveNote, 3)
End Function


Public Function TrimChar(str As String) As String
'功能:去除字符串中连续的空格和回车(含两头的空格,回车),不去除TAB字符,哪怕是连续的
    Dim strTmp As String
    Dim i As Long, j As Long
    
    If Trim(str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(str)
    
    strTmp = Replace(strTmp, "  ", " ")
    strTmp = Replace(strTmp, "  ", " ")
    
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    strTmp = Replace(strTmp, vbCrLf & vbCrLf, vbCrLf)
    
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function


Public Function CheckPars(strSQL As String) As Boolean
'功能：检查SQL语句中参数符"[]"是否配对,以及参数号是否正确(非数字,不连续)
    Dim intLeft As Integer, intRight As Integer
    Dim intMin As Integer, intMax As Integer
    Dim strTmp As String, strPar As String, strPars As String
    Dim i As Long
    
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "[" Then intLeft = intLeft + 1
        If Mid(strSQL, i, 1) = "]" Then intRight = intRight + 1
    Next
    
    If intLeft <> intRight Then Exit Function '"["与"]"不配对
    
    If intLeft = 0 And intRight = 0 Then CheckPars = True: Exit Function
    
    strTmp = strSQL
    intMin = 32767
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        strPar = Left(strTmp, InStr(strTmp, "]") - 1)
        If Trim(strPar) = "" Then
            strPar = 0
        ElseIf Not IsNumeric(strPar) Then
            Exit Function '非数字编号
        End If
        If CInt(strPar) < intMin Then intMin = CInt(strPar)
        If CInt(strPar) > intMax Then intMax = CInt(strPar)
        If InStr(strPars, "," & CInt(strPar)) = 0 Then strPars = strPars & "," & CInt(strPar)
    Loop
    If intMin <> 0 Then Exit Function '不是从0开始编号
    If strPars <> "" Then strPars = Mid(strPars, 2)
    If UBound(Split(strPars, ",")) <> intMax Then Exit Function '不是连续编号
    CheckPars = True
End Function


Public Function SQLObject(ByVal strSQL As String) As String
'功能：分析SQL语句所用到的对象名
'参数：strSQL=要分析的原始SQL语句
'返回：SQL语句所访问到的对象名,如"部门表,病人费用记录,ZLHIS.人员表"
'说明：1.与Oracle SELECT语句兼容
'      2.如果SQL语句中的对象名前加有所有者前缀,则该前缀不会被截取
'      3.需要函数TrimChar;TrueObject的支持
    Dim intB As Long, intE As Long, intL As Long, intR As Long
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    '大写化及去除多余的字符
    strAnal = UCase(TrimChar(strSQL))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
    '先分解处理嵌套子查询
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB '匹配的左右括号位置
        intL = 1: intR = 0
        For i = intB + 1 To Len(strAnal)
            If Mid(strAnal, i, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, i, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = i
                If intE - intB - 1 <= 0 Then
                    '对于非子查询,将括号换成其它符号,以使循环继续
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                ElseIf InStr(Mid(strAnal, intB + 1, intE - intB - 1), "SELECT") > 0 _
                    And InStr(Mid(strAnal, intB + 1, intE - intB - 1), "FROM") > 0 Then
                    '子查询语句
                    strSub = Mid(strAnal, intB + 1, intE - intB - 1)
                    '将该子查询部份作为为特殊对象名
                    strAnal = Replace(strAnal, Mid(strAnal, intB, intE - intB + 1), "嵌套查询")
                    '递归分析
                    strObject = strObject & "," & SQLObject(strSub)
                Else
                    strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
                    strAnal = Left(strAnal, intE - 1) & "@" & Mid(strAnal, intE + 1)
                End If
                Exit For
            End If
        Next
        '无匹配右括号
        If intE = intB Then strAnal = Left(strAnal, intB - 1) & "@" & Mid(strAnal, intB + 1)
    Loop
    
    '分解分析(此时strAnal为简单查询,可能带Union等连接)
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '从第一个From后面部份开始
        strCur = arrFrom(i)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "START WITH") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "START WITH") - 1)
        ElseIf InStr(strCur, "CONNECT BY") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "CONNECT BY") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        ElseIf InStr(strCur, "MINUS") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "MINUS") - 1)
        ElseIf InStr(strCur, "INTERSECT") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "INTERSECT") - 1)
        Else
            strMulti = strCur
        End If
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
            If InStr(strObject & ",", "," & strTrue & ",") = 0 And strTrue <> "嵌套查询" Then
                If InStr(strTrue, "'") = 0 And InStr(strTrue, "@") = 0 Then
                    strObject = strObject & "," & strTrue
                End If
            End If
        Next
    Next
    '完成
    SQLObject = Mid(strObject, 2)
    SQLObject = Replace(SQLObject, ",,", ",")
    Exit Function
errH:
    Err.Clear
End Function

Private Function TrueObject(ByVal strObject As String) As String
'功能：SQLObject函数的子函数,用于去除对象名中的无用字符
    Dim i As Integer
    '寻找第一个正常字符位置
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, i)
    '寻找后面第一个非正常字符
    For i = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, i, 1)) > 0 Then Exit For
    Next
    If i <= Len(strObject) Then strObject = Left(strObject, i - 1)
    TrueObject = strObject
End Function



Public Function CheckObjectPriv(strObject As String) As String
'功能：检查当前用户对指定对象是否完全有权限访问
'参数：strObject=对象名串,如"部门表,病人费用记录"
'返回：完全=空,不完全=不能访问的对象名,如"部门表,病人费用记录"
'说明：用于在校验数据源之前检查是否有权限查询SQL语句中的对象
'参考：grsObject
    Dim i As Integer
    For i = 0 To UBound(Split(strObject, ","))
        If Split(strObject, ",")(i) <> "DUAL" Then
            If InStr(Split(strObject, ",")(i), ".") = 0 Then
                grsObject.Filter = "OBJECT_NAME='" & Split(strObject, ",")(i) & "'"
            Else
                '如果本身就加了所有者前缀,则检查该所有者对象权限
                grsObject.Filter = "OWNER='" & Split(Split(strObject, ",")(i), ".")(0) & _
                    "' And OBJECT_NAME='" & Split(Split(strObject, ",")(i), ".")(1) & "'"
            End If
            If grsObject.EOF Then
                If InStr(CheckObjectPriv & ",", "," & Split(strObject, ",")(i) & ",") = 0 Then
                    CheckObjectPriv = CheckObjectPriv & "," & Split(strObject, ",")(i)
                End If
            End If
        End If
    Next
    If CheckObjectPriv <> "" Then CheckObjectPriv = Mid(CheckObjectPriv, 2)
End Function

Public Function ObjectOwner(strObject As String, Optional frmParent As Object) As String
'功能：根据对象名加上当前用户所能访问的所有者前缀(包括对同一对象名有多个所有者要求选其中之一)
'参数：strObject=对象名串,如"部门表,病人费用记录"
'返回：正常=加了所有者前缀的对象串,如"ZLPER.部门表,ZLHIS.病人费用记录",取消="取消"
'参考：grsObject
    Dim i As Integer, j As Integer
    
    For i = 0 To UBound(Split(strObject, ","))
        If Split(strObject, ",")(i) <> "DUAL" Then
            If InStr(Split(strObject, ",")(i), ".") > 0 Then
                '如果本身就加了所有者前缀,则使用其本身不变
                If InStr(ObjectOwner, "," & Split(strObject, ",")(i)) = 0 Then
                    ObjectOwner = ObjectOwner & "," & Split(strObject, ",")(i)
                End If
            Else
                grsObject.Filter = "OBJECT_NAME='" & Split(strObject, ",")(i) & "'"
                If grsObject.RecordCount = 1 Then
                    If InStr(ObjectOwner & ",", "," & grsObject!owner & "." & Split(strObject, ",")(i) & ",") = 0 Then
                        ObjectOwner = ObjectOwner & "," & grsObject!owner & "." & Split(strObject, ",")(i)
                    End If
                ElseIf grsObject.RecordCount > 1 Then
                    '同一对象有多个所有者,则要求选择
                    Set frmSelOwner.rsObject = grsObject
                    If frmParent Is Nothing Then
                        frmSelOwner.Show 1
                    Else
                        frmSelOwner.Show 1, frmParent
                    End If
                    If gblnOK Then
                        With frmSelOwner.lvw.SelectedItem
                            If InStr(ObjectOwner & ",", "," & .Text & "." & Split(strObject, ",")(i) & ",") = 0 Then
                                ObjectOwner = ObjectOwner & "," & .Text & "." & Split(strObject, ",")(i)
                            End If
                        End With
                        Unload frmSelOwner
                    Else
                        '取消选择,也就是取消操作(调用程序),返回空
                        ObjectOwner = "取消": Exit Function
                    End If
                End If
            End If
        End If
    Next
    If ObjectOwner <> "" Then ObjectOwner = Mid(ObjectOwner, 2)
End Function

Public Function SQLReplaceOwner(ByVal strSQL As String, strOwner As String) As String
'功能：将SQL语句替换成带对象所有者的形式
'参数：strSQL=原始SQL语句,strOwner=对象所有者串,如"ZLPER.部门表,ZLHIS.病人费用记录"
'返回：访问对象加了所有者前缀的SQL语句
'说明：1.本函数用于直接执行用户SQL语句,而不需要授权对象的私有同义词。
'      2.对表名与字段名相同且字段名没有带表别名,则会出错
    Dim i As Long, j As Long
    Dim intLoc As Long, blnDo As Boolean
    
    '处理成只用空格间隔
    strSQL = UCase(SpaceSQL(strSQL))
    
    For i = 0 To UBound(Split(strOwner, ","))
        '采用循环确认方式,确保替换的是表名,而不是其它语句部份或被包含在其它表名中的部份
        j = 0 '当前开始查找位置
        Do
            j = j + 1
            intLoc = InStr(j, strSQL, Split(Split(strOwner, ",")(i), ".")(1))
            If intLoc > 12 Then '至少有"SELECT FROM "
                '本身就有所有者前缀的不替换
                blnDo = True
                '右边以空格、","号、右括号结束
                blnDo = blnDo And (InStr(",) ", Mid(strSQL, intLoc + Len(Split(Split(strOwner, ",")(i), ".")(1)), 1)) > 0)
                '左边则为","号或"FROM "
                blnDo = blnDo And (Mid(strSQL, intLoc - 1, 1) = "," Or Mid(strSQL, intLoc - 5, 5) = "FROM ")
                If blnDo Then
                    strSQL = Left(strSQL, intLoc - 1) & _
                        Replace(strSQL, Split(Split(strOwner, ",")(i), ".")(1), Split(strOwner, ",")(i), intLoc, 1)
                    j = intLoc + Len(Split(strOwner, ",")(i))
                End If
            End If
        Loop Until j >= Len(strSQL)
    Next
    SQLReplaceOwner = strSQL
End Function

Public Function SpaceSQL(ByVal strSQL As String) As String
'功能：将SQL语句变换为只为空格间隔的形式,以便于分析
    Dim i As Long, j As Long, lngB As Long, lngE As Long
    Dim arrSeg() As Variant
                
    strSQL = Replace(strSQL, vbCr, " ")
    strSQL = Replace(strSQL, vbLf, " ")
    strSQL = Replace(strSQL, vbTab, " ")
    
    lngB = -1
    arrSeg = Array()
    For i = 1 To Len(strSQL)
        If Mid(strSQL, i, 1) = "'" Then
            If lngB = -1 Then
                lngB = i
            Else
                ReDim Preserve arrSeg(UBound(arrSeg) + 1)
                arrSeg(UBound(arrSeg)) = lngB & "," & i
                lngB = -1
            End If
        End If
    Next
    If lngB = -1 Then
        For i = 0 To UBound(arrSeg)
            lngB = CLng(Split(arrSeg(i), ",")(0)) + 1
            lngE = CLng(Split(arrSeg(i), ",")(1)) - 1
            For j = lngB To lngE
                If Mid(strSQL, j, 1) = " " Then
                    strSQL = Left(strSQL, j - 1) & Chr(250) & Mid(strSQL, j + 1)
                End If
            Next
        Next
    End If
    
    Do While InStr(strSQL, "  ") > 0
        strSQL = Replace(strSQL, "  ", " ")
    Loop
    
    strSQL = Replace(strSQL, Chr(250), " ")
    
    strSQL = Replace(strSQL, " ,", ",")
    strSQL = Replace(strSQL, ", ", ",")
    SpaceSQL = strSQL
End Function


Public Sub CopyPars(ByVal objSPars As RPTPars, ByRef objOPars As RPTPars)
'功能：拷贝参数集对象
    Dim tmpPar As RPTPar
    
    Set objOPars = New RPTPars
    For Each tmpPar In objSPars
        With tmpPar
            objOPars.Add .组名, .序号, .名称, .类型, .缺省值, .格式, .值列表, .分类SQL, .明细SQL, .分类字段, .明细字段, .对象, "_" & .Key, .Reserve
        End With
    Next
End Sub


Public Function GetExecSQL(ByVal strSQL As String, Optional ByVal objPars As RPTPars) As String
'功能：获取可执行的SQL
    Dim rstmp As New ADODB.Recordset, tmpFld As Field
    Dim strCheck As String, strLeft As String, strRight As String
    Dim strPar As String, bytPar As Byte, i As Integer
    
    strCheck = strSQL
    On Error GoTo errH
    If Not objPars Is Nothing Then
        Do While InStr(strCheck, "[") > 0
            strLeft = Left(strCheck, InStr(strCheck, "[") - 1)
            strRight = Mid(strCheck, InStr(strCheck, "]") + 1)
            strPar = Mid(strCheck, InStr(strCheck, "[") + 1, InStr(strCheck, "]") - InStr(strCheck, "[") - 1)
            If Trim(strPar) = "" Then strPar = 0
            bytPar = CByte(strPar)
            
            '按缺省参数值替换
            If objPars("_" & CInt(bytPar)).缺省值 <> "" And Not objPars("_" & CInt(bytPar)).缺省值 Like "*…" Then
                Select Case objPars("_" & CInt(bytPar)).类型
                    Case 0 '字符
                        strPar = "'" & Replace(objPars("_" & CInt(bytPar)).缺省值, "'", "''") & "'"
                    Case 1 '数字
                        strPar = objPars("_" & CInt(bytPar)).缺省值
                    Case 2 '日期
                        If Left(objPars("_" & CInt(bytPar)).缺省值, 1) = "&" Then
                            strPar = GetParSQLMacro(objPars("_" & CInt(bytPar)).缺省值)
                        Else
                            If InStr(objPars("_" & CInt(bytPar)).缺省值, ":") > 0 Then
                                '长时间格式
                                strPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '短时间格式
                                strPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            End If
                        End If
                    Case 3 '无类型
                        strPar = objPars("_" & CInt(bytPar)).缺省值
                End Select
            Else '缺省值为空或为自定义项
                Select Case objPars("_" & CInt(bytPar)).类型
                    Case 0 '字符
                        strPar = "'空串'"
                    Case 1 '数字
                        strPar = 1 '设置为0可能导致除数为0
                    Case 2 '日期
                        strPar = "Sysdate"
                    Case 3 '无类型(直接替换)
                        If objPars("_" & CInt(bytPar)).缺省值 = "固定值列表…" Then
                            '取固定值中的缺省值
                            '不好的分隔符
                            For i = 0 To UBound(Split(objPars("_" & CInt(bytPar)).值列表, "|"))
                                If Left(Split(objPars("_" & CInt(bytPar)).值列表, "|")(i), 1) = "√" Then
                                    strPar = Split(Split(objPars("_" & CInt(bytPar)).值列表, "|")(i), ",")(1)
                                    Exit For
                                End If
                            Next
                            '没有设置缺省值则取第一个
                            If strPar = "" Then
                                strPar = Split(Split(objPars("_" & CInt(bytPar)).值列表, "|")(0), ",")(1)
                            End If
                        ElseIf objPars("_" & CInt(bytPar)).缺省值 = "选择器定义…" Then
                            If objPars("_" & CInt(bytPar)).值列表 <> "" Then
                                '取缺省绑定值
                                strPar = Split(objPars("_" & CInt(bytPar)).值列表, "|")(1)
                            ElseIf objPars("_" & CInt(bytPar)).明细SQL <> "" And objPars("_" & CInt(bytPar)).明细字段 <> "" Then
                                strPar = GetDefaultValue(objPars("_" & CInt(bytPar)).明细SQL, objPars("_" & CInt(bytPar)).明细字段)
                                If strPar <> "" Then strPar = CStr(Split(strPar, "|")(1))
                                
                                If objPars("_" & CInt(bytPar)).格式 = 1 Then
                                    strPar = " In (" & strPar & ") "
                                End If
                            Else
                                strPar = ""
                            End If
                        Else
                            strPar = objPars("_" & CInt(bytPar)).缺省值
                        End If
                End Select
            End If
            strCheck = strLeft & strPar & strRight
        Loop
    End If
    
    If InStr(UCase(strCheck), "WHERE ") > 0 Then
        strCheck = Replace(UCase(strCheck), "WHERE ", "Where Rownum<1 And ")
    End If
    GetExecSQL = strCheck
    Exit Function
errH:
    Err.Clear
    GetExecSQL = ""
End Function

Public Function CheckSQL(ByVal strSQL As String, strErr As String, Optional ByVal objPars As RPTPars) As String
'功能：根据缺省参数检查SQL语句书写是否正确
'返回：
'     成功=SQL的字段串,包含了各个字段的名称及类型,格式如"姓名,111|年龄,111|奖金,123",类型值以ADO.Field.Type为准
'     失败=空
    Dim rstmp As New ADODB.Recordset, tmpFld As Field
    Dim strCheck As String, strLeft As String, strRight As String
    Dim strPar As String, bytPar As Byte, i As Integer
    
    strCheck = strSQL
    
    On Error GoTo errH
    If Not objPars Is Nothing Then
        Do While InStr(strCheck, "[") > 0
            strLeft = Left(strCheck, InStr(strCheck, "[") - 1)
            strRight = Mid(strCheck, InStr(strCheck, "]") + 1)
            strPar = Mid(strCheck, InStr(strCheck, "[") + 1, InStr(strCheck, "]") - InStr(strCheck, "[") - 1)
            If Trim(strPar) = "" Then strPar = 0
            bytPar = CByte(strPar)
            
            '按缺省参数值替换
            If objPars("_" & CInt(bytPar)).缺省值 <> "" And Not objPars("_" & CInt(bytPar)).缺省值 Like "*…" Then
                Select Case objPars("_" & CInt(bytPar)).类型
                    Case 0 '字符
                        strPar = "'" & Replace(objPars("_" & CInt(bytPar)).缺省值, "'", "''") & "'"
                    Case 1 '数字
                        strPar = objPars("_" & CInt(bytPar)).缺省值
                    Case 2 '日期
                        If Left(objPars("_" & CInt(bytPar)).缺省值, 1) = "&" Then
                            strPar = GetParSQLMacro(objPars("_" & CInt(bytPar)).缺省值)
                        Else
                            If InStr(objPars("_" & CInt(bytPar)).缺省值, ":") > 0 Then
                                '长时间格式
                                strPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                            Else
                                '短时间格式
                                strPar = "To_Date('" & Format(objPars("_" & CInt(bytPar)).缺省值, "yyyy-MM-dd") & "','YYYY-MM-DD')"
                            End If
                        End If
                    Case 3 '无类型
                        strPar = objPars("_" & CInt(bytPar)).缺省值
                End Select
            Else '缺省值为空或为自定义项
                Select Case objPars("_" & CInt(bytPar)).类型
                    Case 0 '字符
                        strPar = "'空串'"
                    Case 1 '数字
                        strPar = 1 '设置为0可能导致除数为0
                    Case 2 '日期
                        strPar = "Sysdate"
                    Case 3 '无类型(直接替换)
                        If objPars("_" & CInt(bytPar)).缺省值 = "固定值列表…" Then
                            '取固定值中的缺省值
                            '不好的分隔符
                            For i = 0 To UBound(Split(objPars("_" & CInt(bytPar)).值列表, "|"))
                                If Left(Split(objPars("_" & CInt(bytPar)).值列表, "|")(i), 1) = "√" Then
                                    strPar = Split(Split(objPars("_" & CInt(bytPar)).值列表, "|")(i), ",")(1)
                                    Exit For
                                End If
                            Next
                            '没有设置缺省值则取第一个
                            If strPar = "" Then
                                strPar = Split(Split(objPars("_" & CInt(bytPar)).值列表, "|")(0), ",")(1)
                            End If
                        ElseIf objPars("_" & CInt(bytPar)).缺省值 = "选择器定义…" Then
                            If objPars("_" & CInt(bytPar)).值列表 <> "" Then
                                '取缺省绑定值
                                strPar = Split(objPars("_" & CInt(bytPar)).值列表, "|")(1)
                            ElseIf objPars("_" & CInt(bytPar)).明细SQL <> "" And objPars("_" & CInt(bytPar)).明细字段 <> "" Then
                                strPar = GetDefaultValue(objPars("_" & CInt(bytPar)).明细SQL, objPars("_" & CInt(bytPar)).明细字段)
                                If strPar <> "" Then strPar = CStr(Split(strPar, "|")(1))
                                
                                If objPars("_" & CInt(bytPar)).格式 = 1 Then
                                    strPar = " In (" & strPar & ") "
                                End If
                            Else
                                strPar = ""
                            End If
                        Else
                            strPar = objPars("_" & CInt(bytPar)).缺省值
                        End If
                End Select
            End If
            strCheck = strLeft & strPar & strRight
        Loop
    End If
    
    If InStr(UCase(strCheck), "WHERE ") > 0 Then
        strCheck = Replace(UCase(strCheck), "WHERE ", "Where Rownum<1 And ")
    End If
    
    Err.Clear
    On Error Resume Next
    Call zlDatabase.OpenRecordset(rstmp, strCheck, "mdlPublic_CheckSQL") '替换成的都是固定条件,同一数据源一般不变,测试SQL也不会大量运行
    strSQL = strCheck   '返回可执行SQL
    If Err.Number = 0 Then
        strErr = ""
        For Each tmpFld In rstmp.Fields
            If InStr(tmpFld.Name, "|") > 0 Then
                strErr = "字段""" & tmpFld.Name & """没有别名！"
                CheckSQL = "": Exit Function
            ElseIf InStr(tmpFld.Name, "'") > 0 Or InStr(tmpFld.Name, """") > 0 Then
                strErr = "字段名 " & tmpFld.Name & " 非法！"
                CheckSQL = "": Exit Function
            Else
                If InStr(CheckSQL & "|", "|" & tmpFld.Name & "," & tmpFld.Type & "|") = 0 Then
                    CheckSQL = CheckSQL & "|" & tmpFld.Name & "," & tmpFld.Type
                Else
                    strErr = "在数据源中发现相同的字段项目！"
                    CheckSQL = "": Exit Function
                End If
            End If
        Next
        CheckSQL = Mid(CheckSQL, 2)
    Else
        strErr = Err.Number & ":" & vbCrLf & Err.Description
        Err.Clear
    End If
    Exit Function
errH:
    Err.Clear
    strErr = "处理数据源的参数发生错误，可能是SQL中的参数在数据表中不存在。"
    CheckSQL = ""
End Function


Public Function GetParSQL(ByVal strSQL As String) As String
'功能：将SQL换成带参数的格式
'Select * FRom 部门表 Where ID=/*B1*/413/*E1*/
'Select * FRom 部门表 Where ID=[1]
    Dim strTmp As String, i As Integer
    Dim strL As String, strR As String
    Dim intMax As Integer
    
    On Error Resume Next
    
    strTmp = strSQL: intMax = -1
    Do While InStr(strTmp, "/*B") > 0
        strL = Left(strTmp, InStr(strTmp, "/*B") - 1)
        strR = Mid(strTmp, InStr(strTmp, "/*B") + 3)
        If Val(strR) > intMax Then intMax = Val(strR)
        strTmp = strL & strR
    Loop
    
    For i = 0 To intMax
        Do While InStr(strSQL, "/*B" & i & "*/") > 0
            strL = Left(strSQL, InStr(strSQL, "/*B" & i & "*/") - 1)
            strR = Mid(strSQL, InStr(strSQL, "/*E" & i & "*/") + Len("/*E" & i & "*/"))
            strSQL = strL & "[" & i & "]" & strR
        Loop
    Next
    
    GetParSQL = strSQL
End Function

Public Function InString(strText As String, strChars As String) As Boolean
'功能：检查在strText中是否包含strChars中指定的字符
    Dim i As Integer
    
    For i = 1 To Len(strChars)
        If InStr(strText, Mid(strChars, i, 1)) > 0 Then
            InString = True
            Exit Function
        End If
    Next
End Function

Public Function GetDBVer() As Long
    Dim strSQL As String, rstmp As ADODB.Recordset
    
    strSQL = "Select To_Number(Replace(Substr(Banner, 6, 4), '.', '')) As Dbver From V$version Where Substr(Banner, 1, 4) = 'CORE'"
    On Error GoTo errH
    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    GetDBVer = Val(rstmp!dbver)
    '10G:102,9i:92
    Exit Function
errH:
    Err.Clear
    GetDBVer = 102
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(lngTXTProc, hwnd, msg, wp, lp)
End Function


Public Function GetParSQLMacro(str As String) As String
'功能:分析报表参数宏,并返回转换后的在SQL语句中可用的值
    Dim curDate As Date
    
    If InStr(str, "&") = 0 Then GetParSQLMacro = str: Exit Function
    
    curDate = Currentdate
    
    Select Case str
        Case "&当前日期"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&当前日期时间"
            GetParSQLMacro = "Sysdate"
        Case "&当天开始时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&当天结束时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&前一天开始时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate - 1, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&前一天结束时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate - 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&前一天同时间"
            GetParSQLMacro = "Sysdate-1"
        Case "&后一天同时间"
            GetParSQLMacro = "Sysdate+1"
        Case "&后一天结束时间"
            GetParSQLMacro = "TO_DATE('" & Format(curDate + 1, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&后一天日期"
            GetParSQLMacro = "Trunc(Sysdate+1)"
        Case "&前一周日期"
            GetParSQLMacro = "Trunc(Sysdate - 7)"
        Case "&前一月日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", -1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&前一季日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", -3, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&前一年日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("yyyy", -1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&下一周日期"
            GetParSQLMacro = "Trunc(Sysdate + 7)"
        Case "&下一月日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", 1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&下一季日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("m", 3, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&下一年日期"
            GetParSQLMacro = "TO_DATE('" & Format(DateAdd("yyyy", 1, curDate), "yyyy-MM-dd") & "','YYYY-MM-DD')"
        Case "&本月初时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&本月末时间"
            curDate = DateAdd("m", 1, curDate)
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&上月初时间"
            curDate = DateAdd("m", -1, curDate)
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-" & Month(curDate) & "-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&上月末时间"
            curDate = CDate(Year(curDate) & "-" & Month(curDate) & "-01") - 1
            GetParSQLMacro = "TO_DATE('" & Format(curDate, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&本年初时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-01-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&本年末时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) & "-12-31", "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&上年初时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) - 1 & "-01-01", "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
        Case "&上年末时间"
            GetParSQLMacro = "TO_DATE('" & Format(Year(curDate) - 1 & "-12-31", "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    End Select
End Function
    

Public Function GetDefaultValue(ByVal strSQL As String, ByVal strFld As String, Optional ByVal strDefBand As String) As String
'功能：根据参数选择器SQL定义，返回显示字段及绑定字段的值
'参数：strFld=参数数据源字段说明串
'      strDefBand=程序传入的缺省绑定值,是否按此值过滤
'返回：显示值|绑定值|原始记录数
    Dim rstmp As New ADODB.Recordset
    Dim strTmp As String, i As Long
    Dim strShow As String, strBand As String
        
    '取出显示,绑定字段名
    For i = 0 To UBound(Split(strFld, "|"))
        strTmp = Split(strFld, "|")(i)
        If Split(strTmp, ",")(2) Like "*&D*" Then strShow = CStr(Split(strTmp, ",")(0))
        If Split(strTmp, ",")(2) Like "*&B*" Then strBand = CStr(Split(strTmp, ",")(0))
    Next
    If strShow = "" And strBand = "" Then Exit Function
        
    '打开参数数据源
    On Error GoTo errH
    strSQL = Replace(RemoveNote(strSQL), "[*]", "")
    Call zlDatabase.OpenRecordset(rstmp, strSQL, "mdlPublic_GetDefaultValue")  '[*]在SQL的''中,类型无法处理
    i = rstmp.RecordCount '原始记录个数
        
    '先按指定的绑定值过滤出数据行
    If Not rstmp.EOF And strDefBand <> "" Then
        If IsType(rstmp.Fields(strBand).Type, adVarChar) Then
            rstmp.Filter = strBand & "='" & strDefBand & "'"
        ElseIf IsType(rstmp.Fields(strBand).Type, adNumeric) Then
            If Not IsNumeric(strDefBand) Then Exit Function
            rstmp.Filter = strBand & "=" & strDefBand
        ElseIf IsType(rstmp.Fields(strBand).Type, adDBTimeStamp) Then
            If Not IsDate(strDefBand) Then Exit Function
            rstmp.Filter = strBand & "=#" & strDefBand & "#"
        End If
    End If
    
    '再返回缺省行数据或过滤行数据
    If Not rstmp.EOF Then
        strShow = Nvl(rstmp.Fields(strShow).Value, "")
        strBand = Nvl(rstmp.Fields(strBand).Value, "")
        If strShow <> "" Or strBand <> "" Then
            GetDefaultValue = strShow & "|" & strBand & "|" & i
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'功能：判断某个ADO字段数据类型是否与指定字段类型是同一类(如数字,日期,字符,二进制)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'功能：在下拉式工具按钮中弹出一个菜单
    Dim vRect As RECT, vDot1 As POINTAPI, vDot2 As POINTAPI
    
    Call GetWindowRect(ToolBar.hwnd, vRect)
    vDot1.x = vRect.Left: vDot1.Y = vRect.Top
    vDot2.x = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot2)
    
    vDot1.x = vDot1.x * 15: vDot1.Y = vDot1.Y * 15
    vDot2.x = vDot2.x * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.x + Button.Left, vDot2.Y
End Sub


Public Function CheckFormInput(objForm As Object, Optional bln单引号 As Boolean) As Boolean
    Dim obj As Object, strText As String
    
    On Error Resume Next
    For Each obj In objForm.Controls
        If InStr("TextBox,ComboBox", TypeName(obj)) > 0 Then
            If obj.Visible And obj.Enabled Then
                Select Case TypeName(obj)
                Case "TextBox"
                    strText = obj.Text
                Case "ComboBox"
                    If obj.Style = 0 Then strText = obj.Text
                End Select
                If InStr(strText, "'") > 0 And Not bln单引号 Then
                    MsgBox "输入中存在非法字符！", vbInformation, App.Title
                    obj.SelStart = 0: obj.SelLength = Len(obj.Text)
                    obj.SetFocus: Exit Function
                End If
            End If
        End If
    Next
    CheckFormInput = True
End Function

Public Function GetCboIndex(cbo As ComboBox, strFind As String) As Long
'功能：由字任串查找ComboBox的索引值
'参数：cbo=ComboBox,strFind=查找字符串
    Dim i As Integer
    If strFind = "" Then GetCboIndex = -1: Exit Function
    For i = 0 To cbo.ListCount - 1
        If cbo.List(i) = strFind Then
            GetCboIndex = i
            Exit Function
        End If
    Next
    GetCboIndex = -1
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function
 
'------------------------------------------------------------------------------------------------
'以下函数用于分析处理数据源权限------------------------------------------------------------------
'------------------------------------------------------------------------------------------------
Public Function UserObject() As ADODB.Recordset
'功能：获取当前用户所具有Select 权限的所有表及视图名(包含用户自身对象及被授权对象)
'返回：成功=对象名称列表(以中英顺序排序),失败=空
'说明：！！！对中联公共用户对象,本系统将不允许查询
    Dim rstmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = _
        "Select USER as OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW') And USER<>'ZLSOFT'" & _
        " Union" & _
        " Select OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From All_Objects O," & _
        " (Select TABLE_NAME From All_Tab_Privs Where Privilege='SELECT') G" & _
        " Where O.Object_Type in('TABLE','VIEW')" & _
        " and O.OBJECT_NAME=G.TABLE_NAME and O.Owner Not in('ZLSOFT')" & _
        " Order by Sort Desc,OBJECT_NAME"
    
    strSQL = _
        "Select USER as OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From User_Objects" & _
        " Where Object_Type in ('TABLE','VIEW')" & _
        " Union" & _
        " Select OWNER,OBJECT_NAME,Sign(ASCII(OBJECT_NAME)-256) as Sort" & _
        " From All_Objects O," & _
        " (Select TABLE_NAME From All_Tab_Privs Where Privilege='SELECT') G" & _
        " Where O.Object_Type in('TABLE','VIEW')" & _
        " and O.OBJECT_NAME=G.TABLE_NAME" & _
        " Order by Sort Desc,OBJECT_NAME"
        
    strSQL = _
        "Select Owner, Object_Name, Sign(Ascii(Object_Name) - 256) As Sort" & vbNewLine & _
        "From (Select User As Owner, Object_Name" & vbNewLine & _
        "       From User_Objects" & vbNewLine & _
        "       Where Object_Type In ('TABLE', 'VIEW')" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select Table_Schema, Table_Name" & vbNewLine & _
        "       From All_Tab_Privs" & vbNewLine & _
        "       Where Privilege = 'SELECT' And Table_Name Not Like '%_ID'" & vbNewLine & _
        "       Group By Table_Schema, Table_Name)" & vbNewLine & _
        "Order By Sort Desc, Object_Name"

    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rstmp, strSQL, "mdlPublic_UserObject")
    Set UserObject = rstmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub ExecuteProcedureArrAy(ByVal strArr As Variant, ByVal strCaption As String)
    '执行过程:
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To strArr.Count
        strSQL = strArr(i)
        Debug.Print strSQL
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    gcnOracle.CommitTrans
End Sub

