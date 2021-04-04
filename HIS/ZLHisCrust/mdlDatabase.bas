Attribute VB_Name = "mdlDatabase"
Option Explicit
Public Function ErrCenter() As Byte
'功能： 数据事务错误处理中心
'参数：
'返回： cancel      返回 0
'       resume      返回 1
    Dim strNote As String
    
        '---------------VB标准错误--------------------------
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
            strNote = "太长的复制"
        '------------------ADO错误-------------------
        Case 3001
            strNote = "参数类型错误，或数值超过范围，或互相冲突。"
        Case 3021
            strNote = "记录超界(EOF/BOF)，或者当前记录被删除；当前应用操作需要定位当前记录。"
        Case 3219
            strNote = "上下文环境不允许当前应用操作（可能是处于尚未结束的事务）。"
        Case 3246
            strNote = "在事务执行中，不能关闭一个联结对象。"
        Case 3251
            strNote = "当前基础不支持这一应用操作。"
        Case 3265
            strNote = "ADO没找到应用程序要求的对应名称或序号。"
        Case 3367
            strNote = "对象已经存在，不能添加。"
        Case 3420
            strNote = "对象未引用。"
        Case 3421
            strNote = "当前操作使用了错误的数值类型。"
        Case 3704
            strNote = "对象关闭时，当前操作不能执行。"
        Case 3705
            strNote = "对象开启时，当前操作不能执行。"
        Case 3706
            strNote = "ADO没找到指定的支持。"
        Case 3707
            strNote = "不能采用命令对象改变一个记录集的活动连接源等属性。"
        Case 3708
            strNote = "应用程序出现错误的参数定义。"
        Case 3709
            strNote = "应用程序要求一个关闭的引用对象或无效的联结对象。"
        Case Else
            strNote = Err.Description
    End Select
    
    ErrCenter = frmErrAsk.ShowForm(Err.Number, strNote)
    Err.Clear
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
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

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
'功能：通过Command对象打开带参数SQL的记录集
'参数：strSQL=条件中包含参数的SQL语句,参数形式为"[x]"
'             x>=1为自定义参数号,"[]"之间不能有空格
'             同一个参数可多处使用,程序自动换为ADO支持的"?"号形式
'             实际使用的参数号可不连续,但传入的参数值必须连续(如SQL组合时不一定要用到的参数)
'      arrInput=不定个数的参数值,按参数号顺序依次传入,必须是明确类型
'      strTitle=用于SQLTest识别的调用窗体/模块标题
'返回：记录集，CursorLocation=adUseClient,LockType=adLockReadOnly,CursorType=adOpenStatic
'举例：
'SQL语句为="Select 姓名 From 病人信息 Where (病人ID=[3] Or 门诊号=[3] Or 姓名 Like [4]) And 性别=[5] And 登记时间 Between [1] And [2] And 险类 IN([6],[7])"
'调用方式为：Set rsPati=OpenSQLRecord(strSQL, Me.Caption, CDate(Format(rsMove!转出日期,"yyyy-MM-dd")),dtp时间.Value, lng病人ID, "张%", "男", 20, 21)
    Static cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        If lngRight = 0 Then Exit Do
        '可能是正常的"[编码]名称"
        strSeq = Mid(strSQL, lngLeft + 1, lngRight - lngLeft - 1)
        If IsNumeric(strSeq) Then
            i = CInt(strSeq)
            strPar = strPar & "," & i
            If i > intMax Then intMax = i
        End If
        
        lngLeft = InStr(lngRight + 1, strSQL, "[")
    Loop

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
        '产生用于SQL跟踪的语句
        varValue = arrInput(i - 1)
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            strLog = Replace(strLog, "[" & i & "]", varValue)
        Case "String" '字符
            strLog = Replace(strLog, "[" & i & "]", "'" & Replace(varValue, "'", "''") & "'")
        Case "Date" '日期
            strLog = Replace(strLog, "[" & i & "]", "To_Date('" & Format(varValue, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')")
        End Select
    Next

    '清除原有参数:不然不能重复执行
    cmdData.CommandText = "" '不为空有时清除参数出错
    Do While cmdData.Parameters.Count > 0
        cmdData.Parameters.Delete 0
    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
            cmdData.Parameters.Append cmdData.CreateParameter("PAR" & i, adVarNumeric, adParamInput, 30, varValue)
        Case "String" '字符
            intMax = ActualLen(varValue)
            
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
                intMax = ActualLen(varValue(lngLeft))
                            
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
    '执行返回记录集
    Set cmdData.ActiveConnection = gcnOracle '这句比较慢
    cmdData.CommandText = strSQL
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
End Function

Public Sub ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'功能：执行过程语句,并自动对过程参数进行绑定变量处理
'参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
'说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
'  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
'  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
'  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '执行的过程名
        strTemp = Trim(strSQL)
        strProc = Trim(Left(strTemp, InStr(strTemp, "(") - 1))
        
        '执行过程参数
        datCur = CDate(0)
        strTemp = Mid(strTemp, InStr(strTemp, "(") + 1)
        strTemp = Trim(Left(strTemp, Len(strTemp) - 1)) & ","
        For i = 1 To Len(strTemp)
            '是否在字符串内，以及表达式的括号内
            If Mid(strTemp, i, 1) = "'" Then blnStr = Not blnStr
            If Not blnStr And Mid(strTemp, i, 1) = "(" Then intBra = intBra + 1
            If Not blnStr And Mid(strTemp, i, 1) = ")" Then intBra = intBra - 1
            
            If Mid(strTemp, i, 1) = "," And Not blnStr And intBra = 0 Then
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
                        
                        '电子病历处理LOB时，如果用绑定变量转换为RAW时超过2000个字符要用adLongVarChar
                        intMax = ActualLen(strPar)
                        If intMax <= 2000 Then
                            intMax = IIf(intMax <= 200, 200, 2000)
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, intMax, strPar)
                        Else
                            If intMax < 4000 Then intMax = 4000
                            .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adLongVarChar, adParamInput, intMax, strPar)
                        End If
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
                        If datCur = CDate(0) Then datCur = Currentdate
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
                strPar = strPar & Mid(strTemp, i, 1)
            End If
        Next
        
        '程序员调用过程时书写错误
        If blnStr Or intBra <> 0 Then
            Err.Raise -2147483645, , "调用 Oracle 过程""" & strProc & """时，引号或括号书写不匹配。原始语句如下：" & vbCrLf & vbCrLf & strSQL
            Exit Sub
        End If
        
        '补充?号
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        Set cmdData.ActiveConnection = gcnOracle '这句比较慢
        cmdData.CommandType = adCmdText
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
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    gcnOracle.Execute strSQL, , adCmdText
End Sub


Public Function Currentdate() As Date
'功能：提取服务器上当前日期
'参数：
'返回：由于Oracle日期格式的问题，所以
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrH
    Set rsTemp = OpenSQLRecord("SELECT SYSDATE FROM DUAL", App.Title)
    Currentdate = rsTemp!SYSDATE
    Exit Function
ErrH:
    Currentdate = 0
End Function


Public Function IsWriteRunErrLog() As Boolean
'功能:是否记录运行错误
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strReturn As String
    On Error GoTo ErrH
    '3   是否记录运行错误(是否记录使用过程中发生的各种错误)
    strSQL = "select 参数值 from ZLTOOLS.ZlOptions where 参数号=3"
    Set rsTmp = OpenSQLRecord(strSQL, "参数号")

    If Not rsTmp.EOF Then
         strReturn = NVL(rsTmp!参数值, "0")
    Else
         strReturn = "0"
    End If
    IsWriteRunErrLog = strReturn <> "0"
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetNoteLength() As Long
'功能：获取一个或多个字段定义长度
'参数：strTable=表名
'        strColumn=列名
'返回：返回列长度
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngLength As Long
    
    lngLength = 300
    strSQL = "Select Column_Name,Nvl(Data_Precision, Data_Length) Collen ,Owner" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Table_Name = [1] And Column_Name =[2]"
    On Error GoTo ErrH
    Set rsTmp = OpenSQLRecord(strSQL, "FieldsLength", "ZLCLIENTS", "说明")
    If Not rsTmp.EOF Then
        rsTmp.Filter = "Owner='ZLTOOLS'"
        If Not rsTmp.EOF Then
            lngLength = Val(rsTmp!collen)
        Else
            rsTmp.Filter = ""
            rsTmp.Sort = "Owner"
            lngLength = Val(rsTmp!collen)
        End If
    End If
    GetNoteLength = lngLength
    Exit Function
ErrH:
    GetNoteLength = lngLength
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Public Function IsSampleFTP() As Boolean
'功能:是否使用简易FTP
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrH
    strSQL = "Select Nvl(Max(内容), 0) As 使用简易ftp工具 From Zlreginfo Where 项目 = 'FTP不检查文件存在'"
    Set rsTmp = OpenSQLRecord(strSQL, "使用简易ftp工具")
    IsSampleFTP = Val(rsTmp!使用简易ftp工具 & "") <> 0
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function IsHaveVersion() As Boolean
'功能：获取一个或多个字段定义长度
'参数：strTable=表名
'        strColumn=列名
'返回：返回列长度
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lngLength As Long
    
    lngLength = 300
    strSQL = "Select Column_Name,Owner" & vbNewLine & _
                "From All_Tab_Columns" & vbNewLine & _
                "Where Table_Name = [1] And Column_Name =[2] And Owner='ZLTOOLS'"
    On Error GoTo ErrH
    Set rsTmp = OpenSQLRecord(strSQL, "FieldsLength", "ZLFILESUPGRADE", "文件版本号")
    If Not rsTmp.EOF Then
        IsHaveVersion = True
    End If
    Exit Function
ErrH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
'复制记录集
'参数：strFields=需要复制的记录集的字段的列顺序或字段名组成的字符串
'          如：1 别名1,3 别名2,7 别名3...表示复制记录集的第1,3,7..字段组成记录集并返回
'              ID 别名1,姓名 别名2,....表示复制记录集的ID,姓名...字段组成记录集返回
'              别名*为新的记录集的列名
'              两中类型混搭容易出现列名相同的问题，请注意
'           arrAppFields=追加的字段信息：列名,类型,长度,默认值,没有默认值传Empty,没有指定长度传Empty
'      blnOnlyStructure=是否只复制结构
'在程序中，经常会涉及到相互传递记录集，而使用ADO的Clone复制产生的记录集，当其中一个记录集的数据发生变化的时候，所有副本都将发生相同的变化（通常指修改或删除），而我们往往希望这些记录集相互间保持独立
  
    Dim rsClone As ADODB.Recordset
    Dim rsTarget As ADODB.Recordset
    Dim intFields As Integer
    Dim arrFieldsName As Variant, strFieldName As String, strFieldNameAlias As String
    Dim arrTmp As Variant
    Dim i As Long
    
    On Error GoTo ErrH
    If Not rsSource Is Nothing Then
        Set rsClone = rsSource.Clone
        rsClone.Filter = rsSource.Filter
    End If
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        '产生记录集结构
        If Not rsClone Is Nothing Then
            If strFields = "" Then '记录集全复制模式
                arrFieldsName = Array()
                If rsClone.Fields.Count > 0 Then
                    ReDim arrFieldsName(rsClone.Fields.Count - 1)
                Else
                    arrFieldsName = Array()
                End If
                For intFields = 0 To rsClone.Fields.Count - 1
                    arrFieldsName(intFields) = rsClone.Fields(intFields).Name & ""
                    .Fields.Append rsClone.Fields(intFields).Name, IIf(rsClone.Fields(intFields).Type = adNumeric, adDouble, rsClone.Fields(intFields).Type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
                Next
            Else '记录集部分复制模式
                If rsClone.Fields.Count > 0 Then
                    arrFieldsName = Split(strFields, ",")
                    For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                        '列包含别名
                        arrTmp = Split(arrFieldsName(intFields) & " ", " ")
                        strFieldName = Trim(arrTmp(0)): strFieldNameAlias = Trim(arrTmp(1))
                        If IsNumeric(strFieldName) Then strFieldName = rsClone.Fields(Val(strFieldName)).Name & ""
                        '获取字段原名，存入数组
                        arrFieldsName(intFields) = strFieldName
                        '添加字段,若果存在别名，则新增列的列名为别名
                        .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).Type = adNumeric, adDouble, rsClone.Fields(strFieldName).Type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:表示新增
                    Next
                End If
            End If
        End If
        '追加字段添加
        If TypeName(arrAppFields) = "Variant()" Then
            For i = LBound(arrAppFields) To UBound(arrAppFields) Step 4
                If arrAppFields(i + 2) = Empty Then
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), , adFldIsNullable, arrAppFields(i + 3)
                    End If
                Else
                    If arrAppFields(i + 3) = Empty Then
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable
                    Else
                        .Fields.Append arrAppFields(i), arrAppFields(i + 1), arrAppFields(i + 2), adFldIsNullable, arrAppFields(i + 3)
                    End If
                End If
            Next
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        '复制数据
        If Not blnOnlyStructure And Not rsClone Is Nothing Then
            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
            Do While Not rsClone.EOF
                .AddNew
                For intFields = LBound(arrFieldsName) To UBound(arrFieldsName)
                    '新记录集的列按顺序添加，因此可以这样
                    .Fields(intFields).Value = rsClone.Fields(arrFieldsName(intFields)).Value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function UpdateRec(ByRef rsInput As Recordset, ByVal strFilter As String, ParamArray arrInput() As Variant) As Boolean
'功能：更新指定条件的记录集的记录
'参数：rsInput=记录集
'      strFilter=条件
'      arrInput=输入的字段名以及值，格式：字段名1,值1, 字段名2,值2,....
'返回：是否成功
'      rsInput=经过更新后的记录集
'说明：arrInput的字段值可以用记录集中的其他字段来更新该字段，此时格式为：!字段名
    Dim strFiledName As String, strFileValue As String
    Dim blnFiled As Boolean, i As Long

    On Error GoTo ErrH
    With rsInput
        .Filter = strFilter
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            For i = LBound(arrInput) To UBound(arrInput) Step 2
                strFiledName = arrInput(i)
                If IsNull(arrInput(i + 1)) Then
                    rsInput(strFiledName).Value = Null
                Else
                    If arrInput(i + 1) Like "!?*" Then
                        blnFiled = True
                        On Error Resume Next
                        strFileValue = rsInput(Mid(arrInput(i + 1), 2)).Value & ""
                        If Err.Number <> 0 Then Err.Clear: blnFiled = False
                        On Error GoTo ErrH
                    End If
                    If Not blnFiled Then
                        rsInput(strFiledName).Value = arrInput(i + 1)
                    Else
                        rsInput(strFiledName).Value = rsInput(Mid(arrInput(i + 1), 2)).Value
                    End If
                End If
                blnFiled = False
                Call rsInput.Update
            Next
            .MoveNext
        Loop
    End With
    UpdateRec = True
    Exit Function
ErrH:

End Function
