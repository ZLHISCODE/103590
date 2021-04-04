Attribute VB_Name = "mdlComLib"

Option Explicit

'######################################################################################################################

Private mblnTrans As Boolean

'错误日志处理相关变量
Private mlngErrNum As Long
Private mstrErrInfo As String
Private mbytErrType As Byte
Private mstrRecentSQL As String  '最近执行的SQL语句

Public Function OpenCursor(ByVal strFormCaption As String, _
                           ByVal strOwner As String, _
                           ByVal strPackagesName As String, _
                           ParamArray varParValue() As Variant) As ADODB.Recordset
'-----------------------------------------
'功能：调用存储过程返回记录集
'入参：strPackagesName ，格式为 包.过程名
'-----------------------------------------
    Dim cmdPackage As New ADODB.Command
    Dim parPackage As ADODB.Parameter
    Dim arrPar As Variant, I As Integer
    Dim varValue As Variant, intMax As Integer
    Dim intMaxArr As Integer  '记录参数个数
    Dim varOutPar As Variant
    Dim strNode As String
    '添加所有者
    If strOwner <> "" Then
        strPackagesName = strOwner & "." & strPackagesName
    End If
    '清除原有参数:不然不能重复执行
    cmdPackage.CommandText = "" '不为空有时清除参数出错
    Do While cmdPackage.Parameters.Count > 0
        cmdPackage.Parameters.Delete 0
    Loop
    
    '------ IN 参数
    strNode = ""
    For I = 0 To UBound(varParValue)
        varValue = varParValue(I)
        If IsNull(varValue) Then Exit For
        
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("Parameter" & I, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '字符
                intMax = LenB(StrConv(varValue, vbFromUnicode))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("Parameter" & I, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '日期
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("Parameter" & I, adDBTimeStamp, adParamInput, , varValue)
            
        End Select
        strNode = strNode & CStr(varValue) & ","
    Next

    If cmdPackage.ActiveConnection Is Nothing Then
        Set cmdPackage.ActiveConnection = gcnOracle
    End If
    
    
    cmdPackage.CommandType = adCmdStoredProc
    cmdPackage.CommandText = strPackagesName
    
    cmdPackage.Properties("PLSQLRSet") = True
    Set OpenCursor = cmdPackage.Execute
    
    cmdPackage.Properties("PLSQLRSet") = False

End Function

Public Function OpenSQLRecord(ByVal strSql As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, I As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSql, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            I = CInt(strSeq)
            strPar = strPar & "," & I
            If I > intMax Then intMax = I
        End If
        
        lngLeft = InStr(lngRight + 1, strSql, "[")
    Loop

    '替换为"?"参数
    strLog = strSql
    For I = 1 To intMax
        strSql = Replace(strSql, "[" & I & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(I - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & I & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & I & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & I & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
'    cmdData.CommandText = "" '不为空有时清除参数出错
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For I = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(I) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(I) <> lngRight Then lngLeft = 0
            lngRight = arrPar(I)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & I & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

    '执行返回记录集
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
    'End If
    cmdData.CommandText = strSql

    Set OpenSQLRecord = cmdData.Execute

End Function

Public Sub ExecuteProcedure(strSql As String, ByVal strFormCaption As String)
    '功能：执行过程语句,并自动对过程参数进行绑定变量处理
    '参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
    '说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
    '  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
    '  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
    '  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, I As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSql), 1) = ")" Then
        '清除原有参数:不然不能重复执行
'        cmdData.CommandText = "" '不为空有时清除参数出错
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
        '执行的过程名
        strTemp = Trim(strSql)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For I = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, I, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, I, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, I, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, I, 1) = "," And Not blnStr And intBra = 0 Then
                strPar = Trim(strPar)
                With cmdData
                    If IsNumeric(strPar) Then '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then '字符串
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        
                        'Oracle连接符运算:'ABCD'||CHR(13)||'XXXX'||CHR(39)||'1234'
                        If InStr(Replace(strPar, " ", ""), "'||") > 0 Then GoTo NoneVarLine
                        
                        '双"''"的绑定变量处理
                        If InStr(strPar, "''") > 0 Then strPar = Replace(strPar, "''", "'")
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时第2000个字符不正确
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
                        If intMax = 0 Or intMax < 200 Then intMax = 200
                        If intMax > 1999 Then GoTo NoneVarLine
                        
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then '日期
                        strPar = Split(strPar, "(")(1)
                        strPar = Trim(Split(strPar, ",")(0))
                        strPar = Mid(strPar, 2, Len(strPar) - 2)
                        If strPar = "" Then
                            'NULL值当成数字处理可兼容其他类型
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, , Null)
                        Else
                            If Not IsDate(strPar) Then GoTo NoneVarLine
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , CDate(strPar))
                        End If
                    ElseIf UCase(strPar) = "SYSDATE" Then '日期
                        If datCur = CDate(0) Then datCur = CurrentDate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then 'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else '可能是其他复杂的表达式，无法处理
                        GoTo NoneVarLine
                    End If
                End With
                
                strPar = ""
            Else
                strPar = strPar & Mid(strTemp, I, 1)
            End If
        Next
        
        '程序员调用过程时书写错误
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "调用 Oracle 过程""" & strProc & """时，引号或括号书写不匹配。原始语句如下：" & vbCrLf & vbCrLf & strSql
            Exit Sub
        End If
        
        '补充?号
        strTemp = ""
        For I = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        '执行过程
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = gcnOracle '这句比较慢
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
        Call cmdData.Execute
        
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:

    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSql = "Call " & strSql
    If InStr(strSql, "(") = 0 Then strSql = strSql & "()"
    gcnOracle.Execute strSql, , adCmdText

End Sub

Public Function To_Date(ByVal dat日期 As Date) As String
'功能:将入参中的日期传换成ORACLE需要的日期格式串
    To_Date = "To_Date('" & Format(dat日期, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Function Nvl(ByVal varValue As Variant, Optional defaultvalue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), defaultvalue, varValue)
End Function


Public Function VarNvl(ByVal strValue As String, Optional defaultvalue As Variant = "") As Variant
    VarNvl = IIf(strValue = "", Default, strValue)
End Function


Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：服务器上当前日期时间
    '-------------------------------------------------------------
    
    Dim rsTmp  As ADODB.Recordset
    
    Err = 0
    On Error GoTo errHandle
    
    Set rsTmp = OpenCursor("clsDataBase", "ZLTOOLS", "B_Public.Get_Current_Date")
    
    If rsTmp.RecordCount > 0 Then
        CurrentDate = rsTmp.Fields(0)
    Else
        CurrentDate = 0
    End If
    Exit Function
    
errHandle:
    If ErrCenter() = 1 Then Resume
    CurrentDate = 0
    Err = 0
End Function

Public Function ErrCenter() As Byte
'------------------------------------------------
'功能： 数据事务错误处理中心
'参数：
'返回： cancel      返回 0
'       resume      返回 1
'------------------------------------------------
    Dim strNote As String, strTemp As String
    Dim bytReturnType As Byte
    
    bytReturnType = 1
    If gcnOracle.Errors.Count <> 0 Then
        'PL/SQL存储过程错误(包括嵌套的过程调用)
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            '日志变量
            mbytErrType = 1
            mlngErrNum = gcnOracle.Errors(0).NativeError
            mstrErrInfo = gcnOracle.Errors(0).Description
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Function
        End If

'        If gcnOracle.Errors(0).NativeError >= 20000 And gcnOracle.Errors(0).NativeError <= 20200 Then
'            '日志变量
'            mbytErrType = 1
'            mlngErrNum = gcnOracle.Errors(0).NativeError
'            mstrErrInfo = gcnOracle.Errors(0).Description
'
'            strNote = gcnOracle.Errors(0).Description
'            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
'            Exit Function
'        End If
        
        'ORACLE其它错误
        '日志变量
        mbytErrType = 2
        mlngErrNum = gcnOracle.Errors(0).NativeError
        mstrErrInfo = gcnOracle.Errors(0).Description
        
        Select Case gcnOracle.Errors(0).NativeError
        Case 1
            strNote = "已经存在相同内容的数据（要求唯一的内容[如编号、名称等]有重复）。"
            bytReturnType = 0
        Case 903
            strNote = "表名称错误。"
            If mstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & mstrRecentSQL
        Case 904, 920
            strNote = "列名称错误" & vbCrLf & vbCrLf & "SQL语句中使用了不存在的列或语句错误."
            If mstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & mstrRecentSQL
        Case 942
            strNote = "表或视图不存在，很可能是你不具备使用该部分数据的权限。"
            bytReturnType = 0
            
            strTemp = GetInvalidTable(mstrRecentSQL)
            If strTemp <> "" Then
                mstrErrInfo = "请对下列对象进行检查：" & vbCrLf & vbCrLf & vbTab & strTemp
            Else
                mstrErrInfo = "错误SQL语句为：" & vbCrLf & vbCrLf & mstrRecentSQL
            End If
        Case 1000
            strNote = "打开的数据表太多，必要时请系统管理员修改数据库的Open_Cursors配置。"
        Case 1005
            strNote = "错误的用户名或密码。"
        Case 1017
            strNote = "错误的用户名或密码。"
            bytReturnType = 0
        Case 1031
            strNote = "没有足够的权限。"
            bytReturnType = 0
        Case 1045
            strNote = "没有联结数据库的权限。"
            bytReturnType = 0
        Case 1400
            strNote = "由于给主键或要求非空列赋予了空值，导致增加失败。"
            bytReturnType = 0
        Case 1401
            strNote = "由于赋予的值超过了列宽限制，导致增加或更新失败。"
            bytReturnType = 0
        Case 1402
            strNote = "由于赋予的值不符合视图的条件限制，导致增加或更新失败。"
            bytReturnType = 0
        Case 1403
            strNote = "由于未检索到数据，导致后续处理失败。"
        Case 1404
            strNote = "修改列操作，导致相关的索引太大。"
        Case 1405
            strNote = "取得的列值为空。"
        Case 1406
            strNote = "取得的列值被切断而缩短了。"
        Case 1407
            strNote = "由于给主键或要求非空列赋予了空值，导致更新失败。"
            bytReturnType = 0
        Case 1408
            strNote = "指定的列已经建立了索引。"
        Case 1409
            strNote = "不能进行无顺序操作(NoSort)，因为本身就没排序。"
        Case 1410
            strNote = "错误的行ID(ROWID)，行ID必须是数字和字符组成的16进制格式。"
        Case 1411
            strNote = "当前列不能存储超过64K的数据。"
            bytReturnType = 0
        Case 1412
            strNote = "当前列数据类型不能存储零长度字符串。"
            bytReturnType = 0
        Case 1413
            strNote = "错误的小数位数，导致失败。"
            bytReturnType = 0
        Case 1415
            strNote = "不能对一个标签伪列指定外连接[Outer-Join(+)]"
        Case 1416
            strNote = "两张表不能同时指向一个外连接[Outer-Join(+)]"
        Case 1417
            strNote = "一张表只能指定指向不超过一张表的外连接[Outer-Join(+)]"
        Case 1418
            strNote = "指定的索引不存在。"
        Case 1424
            strNote = "错误或无效的换码字符(通配符中只能是'%'或'_')。"
        Case 1425
            strNote = "换码字符必须是长度为1的字符。"
        Case 1426
            strNote = "数值表达式的数据溢出(太大或太小)。"
        Case 1427
            strNote = "单行子查询返回了多行。"
        Case 1428
            strNote = "函数的参数错误或超界。"
        Case 1429
            strNote = "一个二进制日期格式超界。"
        Case 1430
            strNote = "希望增加的列已经存在。"
        Case 1431
            strNote = "授权命令(GRANT)导致内在的不一致。"
        Case 1432
            strNote = "希望删除的公共同义词已经不存在。"
        Case 1433
            strNote = "希望建立的同义词已经存在。"
        Case 1434
            strNote = "希望删除的同义词已经不存在。"
        Case 1435
            strNote = "指定的用户不存在。"
            bytReturnType = 0
        Case 1438
            strNote = "数值超过了列允许的精确程度。"
        Case 1439, 1440, 1441
            strNote = "只有空值列才能修改数据类型、将精度或尺寸减小"
        Case 1536
            strNote = "某个超出表空间的空间限量。"
        Case 2290
            strNote = "由于项目值超过允许的范围（违背了检查约束），导致增加或更新失败。"
            bytReturnType = 0
        Case 2291
            strNote = "由于未填写相关表中存在的项目值(违背了外键约束)，导致增加或更新失败。"
        Case 2091, 2292
            strNote = "因为该记录已经使用，导致删除或更新失败。"
            bytReturnType = 0
        Case 2391
            strNote = "用户已达到数据库所允许的最大登录数。"
        Case 12203
            strNote = "由于主机串书写、配置或服务器问题，不能正常连接。"
            bytReturnType = 0
        Case 20003
            strNote = "存储过程无效，请对失效的存储过程进行编译。"
            If mstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & mstrRecentSQL
        Case Else
            strTemp = Err.Description
            If InStr(strTemp, "PLS-00201") > 0 And InStr(strTemp, "ZL_") > 0 Then
                Dim lngPos As Long
                
                lngPos = InStr(strTemp, "ZL_")
                strTemp = Mid(strTemp, lngPos)
                strTemp = Mid(strTemp, 1, InStr(strTemp, "'") - 1)
                
                strNote = "请在服务器管理工具的角色管理程序中增加对过程“" & strTemp & "”的授权。"
            Else
                strNote = "未知错误，发生在" & gcnOracle.Errors(0).Source
            End If
            If mstrRecentSQL <> "" Then mstrErrInfo = mstrErrInfo & vbCrLf & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & mstrRecentSQL
        End Select
        
    Else
        'VB标准错误
        '日志变量
        mbytErrType = 3
        mlngErrNum = Err.Number
        mstrErrInfo = Err.Description
        
        Select Case Err.Number
            Case 3, 3 - 2146828288
                strNote = "未采用标准返回过程"
            Case 5, 5 - 2146828288
                strNote = "无效的过程或参数"
            Case 6, 6 - 2146828288
                strNote = "数据溢出"
            Case 7, 7 - 2146828288
                strNote = "内存溢出"
            Case 9, 9 - 2146828288
                strNote = "下标超界"
            Case 10, 10 - 2146828288
                strNote = "数组是固定数组或暂时锁定"
            Case 11, 11 - 2146828288
                strNote = "除数为零太小"
            Case 13, 13 - 2146828288
                strNote = "类型不匹配"
            Case 14, 14 - 2146828288
                strNote = "超过字符串允许长度"
            Case 16, 16 - 2146828288
                strNote = "表达式太复杂"
            Case 17, 17 - 2146828288
                strNote = "不支持要求的操作"
            Case 18, 18 - 2146828288
                strNote = "发生了用户中断"
            Case 20, 20 - 2146828288
                strNote = "无错误返回"
            Case 28, 28 - 2146828288
                strNote = "堆栈空间溢出"
            Case 35, 35 - 2146828288
                strNote = "过程或函数未定义"
            Case 47, 47 - 2146828288
                strNote = " 太多的动态联结库（DLL）应用客户"
            Case 48, 48 - 2146828288
                strNote = " 调用动态联结库（DLL）错误"
            Case 49, 49 - 2146828288
                strNote = " 动态联结库（DLL）约定错误"
            Case 51, 51 - 2146828288
                strNote = "内部错误"
            Case 52, 52 - 2146828288
                strNote = "错误的文件名或文件号"
            Case 53, 53 - 2146828288
                strNote = "文件未找到"
            Case 54, 54 - 2146828288
                strNote = "文件格式错误"
            Case 55, 55 - 2146828288
                strNote = "文件已经打开"
            Case 57, 57 - 2146828288
                strNote = "设备输入 / 输出错误"
            Case 58, 58 - 2146828288
                strNote = "文件已经存在"
            Case 59, 59 - 2146828288
                strNote = "错误的记录长度"
            Case 61, 61 - 2146828288
                strNote = "磁盘满"
            Case 62, 62 - 2146828288
                strNote = "输入超过文件尾"
            Case 63, 63 - 2146828288
                strNote = "错误的记录号"
            Case 67, 67 - 2146828288
                strNote = "文件太多"
            Case 68, 68 - 2146828288
                strNote = "设备无效或不支持"
            Case 70, 70 - 2146828288
                strNote = "拒绝访问"
            Case 71, 71 - 2146828288
                strNote = "磁盘未准备好"
            Case 74, 74 - 2146828288
                strNote = "不能命名为不同的驱动器"
            Case 75, 75 - 2146828288
                strNote = "路径 / 文件访问错误"
            Case 76, 76 - 2146828288
                strNote = "路径未找到"
            Case 91, 91 - 2146828288
                strNote = "对象变量或块变量为定义(未新建实例)"
            Case 92, 92 - 2146828288
                strNote = "循环未初始化"
            Case 93, 93 - 2146828288
                strNote = "错误的模式字符串"
            Case 94, 94 - 2146828288
                strNote = "错误地使用空(Null)"
            Case 96, 96 - 2146828288
                strNote = " 由于已经使用的对象时间超过了其设置的最大元素号，导致不可能进入事件"
            Case 97, 97 - 2146828288
                strNote = "不能调用一个未建立实例的类对象函数"
            Case 98, 98 - 2146828288
                strNote = " 不能使用一个私有对象的属性和方法?参数和返回值"
            Case 321, 321 - 2146828288
                strNote = "错误的文件格式"
            Case 322, 322 - 2146828288
                strNote = "不能创建需要的临时文件"
            Case 325, 325 - 2146828288
                strNote = "资源文件中错误的格式"
            Case 380, 380 - 2146828288
                strNote = "错误的属性值"
            Case 381, 381 - 2146828288
                strNote = "错误的属性数组索引"
            Case 382, 382 - 2146828288
                strNote = "不支持的运行时设置"
            Case 383, 383 - 2146828288
                strNote = "不支持的只读属性设置"
            Case 385, 384 - 2146828288
                strNote = "需要属性数组索引"
            Case 387, 387 - 2146828288
                strNote = "不允许的设置"
            Case 393, 393 - 2146828288
                strNote = "不支持的运行时读取"
            Case 394, 394 - 2146828288
                strNote = "不支持的只写属性读取"
            Case 422, 422 - 2146828288
                strNote = "不存在的属性"
            Case 423, 423 - 2146828288
                strNote = "不存在的属性或方法"
            Case 424, 424 - 2146828288
                strNote = "要求一个对象"
            Case 429, 429 - 2146828288
                strNote = "ActiveX不能创建部件"
            Case 430, 430 - 2146828288
                strNote = "类不支持的自动化操作或不支持的界面"
            Case 432, 432 - 2146828288
                strNote = "在自动操作期间未找到文件名或类名称"
            Case 438, 438 - 2146828288
                strNote = "对象不支持该属性或方法"
            Case 440, 440 - 2146828288
                strNote = "自动化对象错误"
            Case 442, 442 - 2146828288
                strNote = "到远程类库或对象库的联结丢失，按OK进入对话移去参照"
            Case 443, 443 - 2146828288
                strNote = "自动化对象没有缺省值"
            Case 445, 445 - 2146828288
                strNote = "对象不支持这种操作"
            Case 446, 446 - 2146828288
                strNote = "对象不支持命名参数"
            Case 447, 447 - 2146828288
                strNote = "对象不支持当前本地设置"
            Case 448, 448 - 2146828288
                strNote = "命名参数未找到"
            Case 449, 449 - 2146828288
                strNote = "参数不是可选的"
            Case 450, 450 - 2146828288
                strNote = "错误的参数个数和属性分配"
            Case 451, 451 - 2146828288
                strNote = "属性赋值(Let)过程和读取(Get)过程不返回对象"
            Case 452, 452 - 2146828288
                strNote = "无效的序号"
            Case 453, 453 - 2146828288
                strNote = "指定的DLL函数未找到"
            Case 454, 454 - 2146828288
                strNote = "代码资源未找到"
            Case 455, 455 - 2146828288
                strNote = "代码资源锁定错误"
            Case 457, 457 - 2146828288
                strNote = "该关键值已经与集合的另一元素结合"
            Case 458, 458 - 2146828288
                strNote = "VB不支持的可变自动化类型"
            Case 459, 459 - 2146828288
                strNote = "对象和类不支持的事件集"
            Case 460, 460 - 2146828288
                strNote = "错误的剪贴板格式"
            Case 461, 461 - 2146828288
                strNote = "方法或数据成员未找到"
            Case 462, 462 - 2146828288
                strNote = "远程服务器不存在或无效"
            Case 463, 463 - 2146828288
                strNote = "类没有在本地注册"
            Case 481, 481 - 2146828288
                strNote = "无效的图片格式"
            Case 482, 482 - 2146828288
                strNote = "打印机错误"
            Case 735, 735 - 2146828288
                strNote = "不能将存储为临时文件"
            Case 744, 744 - 2146828288
                strNote = "未找到搜索的主题"
            Case 746, 746 - 2146828288
                strNote = "太长的复制内容"
                
            'ADO错误
            Case -2147483647
                strNote = "未实现"
            Case -2147483646
                strNote = "内存不足"
            Case -2147483645
                strNote = "一个或多个参数无效"
            Case -2147483644
                strNote = "不支持这样的接口"
            Case -2147483643
                strNote = "无效指针"
            Case -2147483642
                strNote = "无效句柄"
            Case -2147483641
                strNote = "操作终止"
            Case -2147483640
                strNote = "不确定的错误"
            Case -2147483639
                strNote = "一般访问拒绝错误"
            Case -2147483638
                strNote = "完成操作所必需的数据不再可用"
            Case -2147467263
                strNote = "未实现"
            Case -2147467262
                strNote = "不支持这样的接口"
            Case -2147467261
                strNote = "无效指针"
            Case -2147467260
                strNote = "操作终止"
            Case -2147467259
                strNote = "不确定的错误"
            Case -2147467258
                strNote = "线程本地存储失败"
            Case -2147467257
                strNote = "获取共享的内存分配程序失败"
            Case -2147467256
                strNote = "获取内存分配程序失败"
            Case -2147467255
                strNote = "不能初始化类的高速缓存"
            Case -2147467254
                strNote = "不能初始化RPC服务"
            Case -2147467253
                strNote = "不能设置线程本地存储通道控制"
            Case -2147467252
                strNote = "不能分配线程本地存储通道控制"
            Case -2147467251
                strNote = "用户提供的内存分配程序不可接受"
            Case -2147467250
                strNote = "OLE服务互斥量已存在"
            Case -2147467249
                strNote = "OLE服务文件映射已存在"
            Case -2147467247
                strNote = "试图启动OLE服务失败"
            Case -2147467246
                strNote = "在单线程模型中试图再一次调用CoInitialize"
            Case -2147467245
                strNote = "需要一个远程激活，但是不允许"
            Case -2147467244
                strNote = "需要一个远程激活，但是提供的服务器名称无效"
            Case -2147467243
                strNote = "类运行配置的安全id与调用者不同"
            Case -2147467242
                strNote = "使用OLE1服务所需的DDE窗口被禁止"
            Case -2147467241
                strNote = "RunAs指定的必须是域名\用户名或只是用户名"
            Case -2147467240
                strNote = "服务进程不能启动，可能路径名不正确"
            Case -2147467239
                strNote = "当配置标识时服务进程不能启动，路径名可能不正确或无效"
            Case -2147467238
                strNote = "由于配置标识不正确，服务进程不能启动。检查用户名和口令"
            Case -2147467237
                strNote = "不允许客户启动这个服务器"
            Case -2147467236
                strNote = "提供这个服务的服务器不能启动"
            Case -2147467235
                strNote = "本计算机不能和服务器提供的其他计算机通信"
            Case -2147467234
                strNote = "服务器启动后不响应"
            Case -2147467233
                strNote = "服务器的注册信息不一致或不完整"
            Case -2147467232
                strNote = "这个接口的注册信息不一致或不完整"
            Case -2147467231
                strNote = "不支持试图执行的操作"
            Case -2147418113
                strNote = "灾难性失败"
            Case -2147024891
                strNote = "一般访问拒绝错误"
            Case -2147024890
                strNote = "无效句柄"
            Case -2147024882
                strNote = "内存不足"
            Case -2147024809
                strNote = "一个或多个参数无效"
            Case 3000
                strNote = "提供者执行请求的动作失败"
            Case 3001
                strNote = "参数类型错误，或数值超过范围，或与其他类型互相冲突。"
            Case 3002
                strNote = "当打开请求的文件时，发生错误"
            Case 3003
                strNote = "读指定的文件时出错"
            Case 3004
                strNote = "写文件时有错误"
            Case 3021
                strNote = "BOF和EOF中一个为True，或者当前记录已被删，而应用程序的请求操作需要当前记录"
            Case 3219
                strNote = "上下文环境不允许当前应用操作（可能是处于尚未结束的事务）。"
            Case 3220
                strNote = "不能改变提供者"
            Case 3246
                strNote = "在事务执行中，不能关闭一个联结对象。"
            Case 3251
                strNote = "提供者不支持该应用程序请求的操作。"
            Case 3265
                strNote = "ADO没找到应用程序要求的对应名称或序号（可能是列名称错误）。"
            Case 3367
                strNote = "对象已在集合中，不能追加"
            Case 3420
                strNote = "对象未引用或引用的对象不再有效。"
            Case 3421
                strNote = "当前操作使用了错误的数值类型。"
            Case 3704
                strNote = "如果对象已关闭，不允许应用程序请求的操作"
            Case 3705
                strNote = "如果对象已打开，不允许应用程序请求的操作"
            Case 3706
                strNote = "ADO不能找到指定的提供者"
            Case 3707
                strNote = "不能采用命令对象改变一个记录集的活动连接源等属性。"
            Case 3708
                strNote = "应用程序出现错误的参数定义。"
            Case 3709
                strNote = "应用程序请求对一个对象的操作时使用了一个引用，而该引用指向了一个关闭的或无效的Connection对象"
            Case 3710
                strNote = "操作不能重新执行"
            Case 3711
                strNote = "操作仍然在执行"
            Case 3712
                strNote = "操作被取消"
            Case 3713
                strNote = "操作仍然在连接中"
            Case 3714
                strNote = "事务无效"
            Case 3715
                strNote = "操作不在执行过程中"
            Case 3716
                strNote = "在这种情况下运行不安全"
            Case 3717
                strNote = "操作引出一个安全对话"
            Case 3718
                strNote = "操作引出一个安全对话头"
            Case 3719
                strNote = "违背数据的完整性，操作失败。"
            Case 3720
                strNote = "用户没有足够的权限完成操作，操作失败。"
            Case 3721
                strNote = "数据超出给定的数据类型的范围"
            Case 3722
                strNote = "动作违背了模式"
            Case 3723
                strNote = "表达式包含不匹配的符号"
            Case 3724
                strNote = "不能转换值不能创建资源"
            Case 3726
                strNote = "这一行中不存在指定的列"
            Case 3727
                strNote = "URL不存在"
            Case 3728
                strNote = "没有查看目录树的权限"
            Case 3729
                strNote = "提供的URL无效"
            Case 3730
                strNote = "资源被锁定"
            Case 3731
                strNote = "资源已经存在"
            Case 3732
                strNote = "不能完成动作"
            Case 3733
                strNote = "文件版本信息没找到"
            Case 3734
                strNote = "服务器得不到足够的空间完成操作，操作失败"
            Case 3735
                strNote = "资源超出范围"
            Case 3736
                strNote = "命令不可用"
            Case 3737
                strNote = "在命名的行中的URL不存在"
            Case 3738
                strNote = "不能删除资源，这超出了允许范围"
            Case 3739
                strNote = "对于选择的列，这个属性无效"
            Case 3740
                strNote = "给属性提供了一个无效的选择"
            Case 3741
                strNote = "给属性提供了一个无效的值"
            Case 3742
                strNote = "设置这个属性造成和其他属性冲突"
            Case 3743
                strNote = "不是所有的属性都能被设置"
            Case 3744
                strNote = "属性没有被设置"
            Case 3745
                strNote = "属性不能被设置"
            Case 3746
                strNote = "属性不被支持"
            Case 3747
                strNote = "类别没有设置所以动作不能执行"
            Case 3748
                strNote = "不能改变连接"
            Case 3749
                strNote = "Fields集合的Update方法失败"
            Case 3750
                strNote = "不能设置Deny权限，因为提供者不支持"
            Case 3751
                strNote = "提供者不支持请求的Deny类型"
                
            Case Else
                strNote = "发生未知的界面错误"
        End Select
        bytReturnType = 0
    End If

    If bytReturnType = 1 Then
        ErrCenter = frmErrAsk.ShowEdit(mlngErrNum, strNote, mstrErrInfo)
    Else
        Call frmErrNote.ShowEdit(mlngErrNum, strNote, mstrErrInfo)
        ErrCenter = 0
    End If
    
    '清除错误
    Err.Clear
End Function

Public Function GetInvalidTable(ByVal strRecentSQL As String) As String
'功能：得到在最近使用的SQL语句中不能访问的表或视图
'
    Dim varTables As Variant
    Dim strTable As String, lngCount As Long
    Dim strInvalidTable As String
    
    varTables = Split(SQLObject(strRecentSQL), ",")
    
    On Error Resume Next
    
    For lngCount = 0 To UBound(varTables)
        strTable = varTables(lngCount)
        
        '测试该对象是否可用
        gcnOracle.Execute "select 1 from " & strTable & " where rownum<1"
        If Err <> 0 Then
            Err.Clear
            strInvalidTable = strInvalidTable & "," & strTable
        End If
    Next
    
    If strInvalidTable <> "" Then
        '去掉第一个逗号
        GetInvalidTable = Mid(strInvalidTable, 2)
    End If
End Function

Public Function SQLObject(ByVal strSql As String) As String
'功能：分析SQL语句所用到的对象名
'参数：strSQL=要分析的原始SQL语句
'返回：SQL语句所访问到的对象名,如"部门表,病人费用记录,ZLHIS.人员表"
'说明：1.与Oracle SELECT语句兼容
'      2.如果SQL语句中的对象名前加有所有者前缀,则该前缀不会被截取
'      3.需要函数TrimChar;TrueObject的支持
    Dim intB As Integer, intE As Integer, intL As Integer, intR As Integer
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim I As Integer, J As Integer
    
    On Error GoTo errH
    
    '大写化及去除多余的字符
    strAnal = UCase(TrimChar(strSql))

    If InStr(strAnal, "SELECT") = 0 Or InStr(strAnal, "FROM") = 0 Then Exit Function
    
    '先分解处理嵌套子查询
    Do While InStr(strAnal, "(") > 0
        intB = InStr(strAnal, "("): intE = intB '匹配的左右括号位置
        intL = 1: intR = 0
        For I = intB + 1 To Len(strAnal)
            If Mid(strAnal, I, 1) = "(" Then
                intL = intL + 1
            ElseIf Mid(strAnal, I, 1) = ")" Then
                intR = intR + 1
            End If
            If intL = intR Then
                intE = I
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
    
    '分解分析
    arrFrom = Split(strAnal, "FROM")
    For I = 1 To UBound(arrFrom) '从第一个From后面部份开始
        strCur = arrFrom(I)
        If InStr(strCur, "WHERE") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "WHERE") - 1)
        ElseIf InStr(strCur, "GROUP") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "GROUP") - 1)
        ElseIf InStr(strCur, "HAVING") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "HAVING") - 1)
        ElseIf InStr(strCur, "ORDER") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "ORDER") - 1)
        ElseIf InStr(strCur, "UNION") > 0 Then
            strMulti = Left(strCur, InStr(strCur, "UNION") - 1)
        Else
            strMulti = strCur
        End If
        For J = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(J))
            If InStr(strObject, "," & strTrue) = 0 And strTrue <> "嵌套查询" Then
                strObject = strObject & "," & strTrue
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

Private Function TrimChar(Str As String) As String
'功能:去除字符串中连续的空格和回车(含两头的空格,回车),不去除TAB字符,哪怕是连续的
    Dim strTmp As String
    Dim I As Long, J As Long
    
    If Trim(Str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(Str)
    I = InStr(strTmp, "  ")
    Do While I > 0
        strTmp = Left(strTmp, I) & Mid(strTmp, I + 2)
        I = InStr(strTmp, "  ")
    Loop
    
    I = InStr(1, strTmp, vbCrLf & vbCrLf)
    Do While I > 0
        strTmp = Left(strTmp, I + 1) & Mid(strTmp, I + 4)
        I = InStr(1, strTmp, vbCrLf & vbCrLf)
    Loop
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
End Function

Private Function TrueObject(ByVal strObject As String) As String
'功能：SQLObject函数的子函数,用于去除对象名中的无用字符
    Dim I As Integer
    '寻找第一个正常字符位置
    For I = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, I, 1)) = 0 Then Exit For
    Next
    strObject = Mid(strObject, I)
    '寻找后面第一个非正常字符
    For I = 1 To Len(strObject)
        If InStr(Chr(32) & Chr(13) & Chr(10) & Chr(9), Mid(strObject, I, 1)) > 0 Then Exit For
    Next
    If I <= Len(strObject) Then strObject = Left(strObject, I - 1)
    TrueObject = strObject
End Function

