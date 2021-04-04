Attribute VB_Name = "mdlDataBase"
'@模块 mdlDataBase-2019/9/17
'@编写 lshuo
'@功能
'   数据库的实现
'@引用
'
'@备注
'
Option Explicit

Public Function CallProcedure(cnInput As ADODB.Connection, ByVal strProcName As String, ByVal strFormCaption As String _
    , blnUseLog As Boolean, ParamArray arrProcParas() As Variant) As Variant
    
    Dim arrPars()   As Variant
    Dim arrRet      As Variant
    
    arrPars = arrProcParas
    arrRet = mdlDataBase.CallProcedureByArray(cnInput, strProcName, strFormCaption, blnUseLog, arrPars, True)
    If UBound(arrRet) = 0 Then
        If IsObject(arrRet(0)) Then
            Set CallProcedure = arrRet(0)
        Else
            CallProcedure = arrRet(0)
        End If
    Else
        CallProcedure = arrRet
    End If
End Function

Public Function CallProcedureByArray(cnInput As ADODB.Connection, ByVal strProcName As String _
    , ByVal strFormCaption As String, blnUseLog As Boolean, arrProcParas() As Variant _
    , Optional ByVal blnAsSubCall As Boolean) As Variant
    
    Dim cmdData As New ADODB.Command
    Dim rsReturn As New ADODB.Recordset
    Dim i As Long, lngAdjust As Long, dtSize As Long, lngMax As Long, lngParaUbound As Long
    Dim pdCur As ParameterDirectionEnum, dtCur As DataTypeEnum
    Dim varValue As Variant, arrRet() As Variant
    Dim arrIntOut() As Integer
    Dim blnOELDB As Boolean
    
    If blnAsSubCall Then
        CallProcedureByArray = Array(False)
    Else
        CallProcedureByArray = False
    End If
    Const MAX_STRING_SIZE = 32767
    
    ReDim Preserve arrIntOut(0)
    For i = LBound(arrProcParas) To UBound(arrProcParas)
        '1、解析参数的定义长度、类型、方向、以及参数值
        pdCur = adParamUnknown: dtCur = adEmpty: dtSize = -1: varValue = Empty: lngMax = -1
        '出参：Empty
        If IsEmpty(arrProcParas(i)) Then
            pdCur = adParamOutput: dtCur = adLongVarChar
        ElseIf IsArray(arrProcParas(i)) Then
            '出参：Array(Empty[,字段类型,字段长度])
            '返回值：Array(Empty,Empty[,字段类型,字段长度])
            varValue = arrProcParas(i)(0)
            pdCur = adParamOutput
            lngParaUbound = UBound(arrProcParas(i))
            '返回值：Array(Empty,Empty[,字段类型,字段长度])
            If lngParaUbound > 0 Then
                If IsEmpty(arrProcParas(i)(0)) And IsEmpty(arrProcParas(i)(1)) Then
                    pdCur = adParamReturnValue
                    If i <> 0 Then
                        Err.Raise vbObjectError + 2, strFormCaption, "函数的返回值(位置0)必须在函数参数之前传递，当前返回值位置：" & i
                    End If
                    If lngParaUbound > 1 Then
                        dtCur = arrProcParas(i)(2)
                        If lngParaUbound > 2 Then
                            dtSize = arrProcParas(i)(3)
                        End If
                    End If
                End If
            End If
            '出参：Array(Empty[,字段类型,字段长度])
            '入出参：Array(值[,字段类型,字段长度])
            If pdCur <> adParamReturnValue Then
                If Not IsEmpty(arrProcParas(i)(0)) Then
                    pdCur = adParamInputOutput
                End If
                If lngParaUbound > 0 Then
                    dtCur = arrProcParas(i)(1)
                    If lngParaUbound > 1 Then
                        dtSize = arrProcParas(i)(2)
                    End If
                End If
            End If
        Else
            varValue = arrProcParas(i)
            pdCur = adParamInput
        End If
        '2\收集出参位置，以方便后面收集返回值。
        If pdCur > adParamInput Then
            ReDim Preserve arrIntOut(UBound(arrIntOut) + 1)
            arrIntOut(UBound(arrIntOut)) = i
        End If
        '3、增加参数
        Select Case VarType(varValue)
            Case vbString
                lngMax = dtSize
                If dtSize = -1 Then         '未定义长度，则获取原始长度
                    lngMax = LenB(varValue) 'LenB(StrConv(varValue, vbFromUnicode)) '超长字符串，该转换耗时
                Else
                    lngMax = arrProcParas(i)(1)
                    If lngMax < LenB(varValue) Then     '定义长度小于实际长度，则变更为实际长度
                        lngMax = LenB(varValue)  'LenB(StrConv(varValue, vbFromUnicode)) '超长字符串，该转换耗时
                    End If
                End If
                '取OLEDB的字符阶数
                If lngMax <= 4000 Then
                    If lngMax < 32 Then
                        lngMax = 32
                    ElseIf lngMax < 128 Then
                        lngMax = 128
                    ElseIf lngMax < 2000 Then
                        lngMax = 2000
                    Else
                        lngMax = 4000
                    End If
                ElseIf dtSize = -1 And pdCur > adParamInput Then '原始长度的出参>4000 ,且没有定义长度，统一使用最大长度
                    lngMax = MAX_STRING_SIZE
                End If
                
                 If lngMax <= 2000 Then
                    If dtCur = adEmpty Then dtCur = adVarChar   '小于2000且没有指定类型，使用adVarChar
                Else
                    If dtCur = adEmpty Then dtCur = adLongVarChar   '大于2000且没有指定类型，使用adLongVarChar
                End If
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, lngMax, varValue)
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal '数字
                If dtCur = adEmpty Then dtCur = adVarNumeric
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, 38, varValue) '以前30修改为38
            Case vbDate
                If dtCur = adEmpty Then dtCur = adDBTimeStamp
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, , varValue)
            Case vbNull, vbEmpty    '参数值是NULL或者EMPTY。NULL，说明参数是NULL,EMPTY，说明参数是出参。
                If dtCur = adEmpty Then
                    If dtSize = -1 Then dtSize = MAX_STRING_SIZE    '没有定义长度，统一使用最大长度
                    If dtSize <= 2000 Then
                        dtCur = adVarChar   '小于2000且没有指定类型，使用adVarChar
                    Else
                        dtCur = adLongVarChar
                    End If
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, dtSize, varValue)
                Else
                    Select Case dtCur
                        Case adVarChar, adLongVarChar, adVarWChar, adLongVarWChar, adBSTR, adVarBinary, adLongVarBinary
                            If dtSize = -1 Then dtSize = MAX_STRING_SIZE    '没有定义长度，统一使用最大长度
                        Case adDBTimeStamp, adDBTime, adDBDate, adDate
                        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                            If dtSize = -1 Then dtSize = 38
                    End Select
                    If VarType(varValue) <> vbNull Then
                        If dtSize = -1 Then
                            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, , varValue)
                        Else
                            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, dtSize, varValue)
                        End If
                    Else
                        If dtSize = -1 Then
                            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur)
                        Else
                            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, dtCur, pdCur, dtSize)
                        End If
                    End If
                End If
            Case Else
                Err.Raise vbObjectError + 1, strFormCaption, "存储过程传递的参数类型无法识别，参数位置：" & i
        End Select
    Next
    Set cmdData.ActiveConnection = cnInput   '这句比较慢
    cmdData.CommandType = adCmdStoredProc
    cmdData.CommandText = strProcName
    blnOELDB = IsOLEDBConnection(cnInput)
    If blnOELDB Then
        cmdData.Properties("PLSQLRSet") = True
    End If
    
    'Set rsReturn = cmdData.Execute
    Set rsReturn = mdlDataBase.CommandExecuteStoredProc(cmdData)
    
    If blnOELDB Then
        cmdData.Properties("PLSQLRSet") = False
    End If
    If rsReturn.State = adStateClosed Then
        arrIntOut(0) = -1       '标记无返回游标
    End If
    If UBound(arrIntOut) > 0 Or arrIntOut(0) <> -1 Then
        '只有1个普通出参以及普通返回值
        If UBound(arrIntOut) = 1 And arrIntOut(0) = -1 Then
            If blnAsSubCall Then
                CallProcedureByArray = Array(cmdData.Parameters(arrIntOut(1)).Value & "")
            Else
                CallProcedureByArray = cmdData.Parameters(arrIntOut(1)).Value & ""
            End If
        '只返回了游标
        ElseIf UBound(arrIntOut) = 0 Then
            If blnAsSubCall Then
                CallProcedureByArray = Array(rsReturn)
            Else
                Set CallProcedureByArray = rsReturn
            End If
        Else
            '加上游标，总共返回数量超过1
            If arrIntOut(0) = -1 Then
                ReDim Preserve arrRet(UBound(arrIntOut) - 1)
                lngAdjust = 1
            Else
                ReDim Preserve arrRet(UBound(arrIntOut))
                Set arrRet(0) = rsReturn
                lngAdjust = 0
            End If
            For i = 1 To UBound(arrIntOut)
                arrRet(i - lngAdjust) = cmdData.Parameters(arrIntOut(i)).Value & ""
            Next
            CallProcedureByArray = arrRet
        End If
    Else
        '无任何返回信息
        If blnAsSubCall Then
            CallProcedureByArray = Array(True)
        Else
            CallProcedureByArray = True
        End If
    End If
End Function

Public Function OpenSQLRecord(cnInput As ADODB.Connection, ByVal strSql As String, ByVal strTitle As String _
    , ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant
    arrPars = arrInput
    Set OpenSQLRecord = mdlDataBase.OpenSQLRecordByArray(cnInput, strSql, strTitle, arrPars)
End Function

Public Function OpenSQLRecordByArray(cnInput As ADODB.Connection, ByVal strSql As String, ByVal strTitle As String _
    , arrInput() As Variant, Optional intLobOprate As Integer = 0) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'               因为使用绑定变量,对带"'"的字符参数,不需要使用"''"形式。
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'      intLobOprate=0:普通SQL,1:LOB类型读取SQL,2:LOB保存SQL
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrStr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    Dim strError As String
    Dim lngPos     As Long
    Dim cnOLEDB     As ADODB.Connection
    
    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLTmp = Trim(UCase(strSql))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrStr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrStr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrStr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrStr(i)) - 1)
                strTmp = Replace(FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrStr(i)) + Len(arrStr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrStr) Then
            If Not Replace(strSQLTmp, " ", "") Like "*/[*]+CARDINALITY*[*]/*" Then '可能有多个CARDINALITY，如：/*+cardinality(c,10) cardinality(d,10)*/
                strSql = "Select /*+ RULE*/" & Mid(Trim(strSql), 7)
            End If
        End If
    End If
    
'    If Replace(strSQLTmp, " ", "") Like "*/[*]+DRIVING_SITE*[*]/*" Then
'        If Not CheckDatamoveRemote Then
'            arrStr = Split(strSql, "/*")
'            strSql = arrStr(LBound(arrStr))
'            For i = LBound(arrStr) To UBound(arrStr)
'                If i <> UBound(arrStr) Then
'                    lngPos = InStr(arrStr(i + 1), "*/")
'                    lngLeft = 0
'                    If lngPos <> 0 Then
'                        lngLeft = InStr(1, arrStr(i + 1), "DRIVING_SITE", vbTextCompare)
'                        If lngLeft > 0 Then
'                            If Trim(Mid(arrStr(i + 1), 1, lngLeft - 1)) <> "+" Then
'                                lngLeft = 0
'                            End If
'                        End If
'                    End If
'
'                    If lngLeft > 0 And lngLeft < lngPos Then
'                        strSql = strSql & Mid(arrStr(i + 1), lngPos + 2)
'                    Else
'                        strSql = strSql & "/*" & arrStr(i + 1)
'                    End If
'                End If
'            Next
'        End If
'    End If
    
    Call AdjustSQL(strSql)
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        If lngRight = 0 Then Exit Do
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSql, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSql, "[")
    Loop
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL语句绑定变量不全，调用来源：" & strTitle
    End If

    '替换为"?"参数
    strLog = strSql
    For i = 1 To intMax
        strSql = Replace(strSql, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency", "Decimal" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = LenB(StrConv(varValue, vbFromUnicode))
            If intMax <= 2000 Then
                intMax = IIf(intMax <= 200, 200, 2000)
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarChar, adParamInput, intMax, varValue)
            Else
                If intMax < 4000 Then intMax = 4000
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adLongVarChar, adParamInput, intMax, varValue)
            End If
        Case "Date" '日期
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adDBTimeStamp, adParamInput, , varValue)
        Case "Variant()" '数组
            '这种方式可用于一些IN子句或Union语句
            '表示同一个参数的多个值,参数号不可与其它数组的参数号交叉,且要保证数组的值个数够用
            If arrPar(i) <> lngRight Then lngLeft = 0
            lngRight = arrPar(i)
            Select Case TypeName(varValue(lngLeft))
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarNumeric, adParamInput, 30, varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", varValue(lngLeft), 1, 1)
            Case "String" '字符
                intMax = LenB(StrConv(varValue(lngLeft), vbFromUnicode))
                If intMax <= 2000 Then
                    intMax = IIf(intMax <= 200, 200, 2000)
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adVarChar, adParamInput, intMax, varValue(lngLeft))
                Else
                    If intMax < 4000 Then intMax = 4000
                    cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adLongVarChar, adParamInput, intMax, varValue(lngLeft))
                End If
                
                strLog = Replace(strLog, "[" & lngRight & "]", "'" & Replace(varValue(lngLeft), "'", "''") & "'", 1, 1)
            Case "Date" '日期
                cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i & "_" & lngLeft, adDBTimeStamp, adParamInput, , varValue(lngLeft))
                strLog = Replace(strLog, "[" & lngRight & "]", "To_Date('" & Format(varValue(lngLeft), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", 1, 1)
            End Select
            lngLeft = lngLeft + 1 '该参数在数组中用到第几个值了
        End Select
    Next

'    If intLobOprate = 0 Then
        Set cmdData.ActiveConnection = cnInput '这句比较慢(这句执行1000次约0.5x秒)
'    Else
'        Set cnOLEDB = gcnOracleOLEDB
'        If cnOLEDB Is Nothing Then
'            If Not IsOLEDBConnection(cnInput) Then
'                Set cnOLEDB = gobjRegister.ReGetConnection(Val("1-OraOLEDB"), strError, cnInput)
'            Else
'                Set cnOLEDB = cnInput
'            End If
'            If cnInput Is gcnOracle Then Set gcnOracleOLEDB = cnOLEDB
'        End If
'        Set cmdData.ActiveConnection = cnOLEDB
'    End If

    cmdData.CommandText = strSql
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
'    If intLobOprate > 0 Then '保存LOB,读取LOB也要使用该参数，否则很慢，约10倍差距
'        Set OpenSQLRecordByArray = New ADODB.Recordset
'        OpenSQLRecordByArray.Open cmdData, , adOpenStatic, adLockOptimistic
'    Else
        Set OpenSQLRecordByArray = cmdData.Execute
        On Error Resume Next
        Set OpenSQLRecordByArray.ActiveConnection = Nothing
        On Error GoTo 0
'    End If
'    Call gobjComLib.SQLTest
End Function

Private Sub AdjustSQL(ByRef strSQLIn As String)
'功能：更新SQL书写方式，避免ADO执行异常

    Const STR_SELECT As String = "select ", STR_FROM As String = "from"
    
    Dim i As Long, lngPos As Long
    Dim intLen As Integer
    Dim strSql As String
    Dim blnDo As Boolean
    
    intLen = Len(STR_SELECT)
    lngPos = 1
    
    On Error GoTo hErr
    
    '1.处理“*/'”特殊字串
    If strSQLIn Like "*/[*]*+*[*]/'*" Then
        strSql = strSQLIn
        Do While True
            lngPos = InStr(lngPos, LCase$(strSql), STR_SELECT)
            If lngPos > 0 Then
                If Trim(Mid$(strSql, lngPos + intLen, 20)) Like "/[*]*+*" Then
                    i = InStr(Mid$(strSql, lngPos + intLen), "/*+")
                    If i <= 0 Then i = InStr(Mid$(strSql, lngPos + intLen), "/* ")
                    If i > 0 Then
                        '存在“ */'”特殊字串
                        For i = lngPos + intLen + 3 To Len(strSql)
                            If Mid$(strSql, i, 3) = "*/'" Then
                                strSql = Left$(strSql, i - 1) & "*/ '" & Mid$(strSql, i + 3)
                            ElseIf LCase(Mid$(strSql, i, Len(STR_FROM))) = STR_FROM Then
                                Exit For
                            End If
                        Next
                        lngPos = i
                    End If
                Else
                    lngPos = lngPos + intLen
                End If
            Else
                Exit Do
            End If
        Loop
        blnDo = True
    End If
    
    '2.暂无
    '...
    
    If blnDo Then strSQLIn = strSql
    Exit Sub
    
hErr:
    '...
End Sub

Public Function CommandExecuteStoredProc(ByRef cmdVar As ADODB.Command) As ADODB.Recordset
    If cmdVar Is Nothing Then Exit Function
    
    On Error GoTo hErr
    Set CommandExecuteStoredProc = cmdVar.Execute
    
    If Not cmdVar.ActiveConnection.Errors Is Nothing Then
        If cmdVar.ActiveConnection.Errors.Count > 0 Then
            If VBA.Err.Number = 0 _
                And (cmdVar.ActiveConnection.Errors(0).Number = CLng(&H40EC9) _
                    Or UCase(cmdVar.ActiveConnection.Errors(0).Description) Like "*ORAOLEDB*40EC9*") Then
                '该条件下清空连接对象的Errors，防止后续代码报错后误将该异常抛出
                cmdVar.ActiveConnection.Errors.Clear
            End If
        End If
    End If
    Exit Function
    
hErr:
    Err.Raise Err.Number, , Err.Description
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'参数：strText=处理字符
'         blnCrlf=是否去掉换行符
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function

'Public Function CheckDatamoveRemote(Optional ByVal lngSys As Long = 100) As Boolean
''功能：检查系统的历史库是否是DBLinK
'    Dim rsTmp As ADODB.Recordset, strSql As String
'
'    On Error GoTo ErrH
'    If lngSys <> 100 Or mbytCheckDatamoveRemote = 0 Then
'        strSql = "Select 1 From zlBakSpaces Where 系统 = [1] And 当前 = 1 And Db连接 Is Not Null"
'        Set rsTmp = OpenSQLRecord(strSql, "CheckDatamoveRemote", lngSys)
'        CheckDatamoveRemote = rsTmp.RecordCount > 0
'        If CheckDatamoveRemote Then
'            mbytCheckDatamoveRemote = 1
'        Else
'            mbytCheckDatamoveRemote = 2
'        End If
'    Else
'        CheckDatamoveRemote = mbytCheckDatamoveRemote = 1
'    End If
'    Exit Function
'
'ErrH:
'    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
'End Function
