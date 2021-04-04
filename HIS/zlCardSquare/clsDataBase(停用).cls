VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*******************************************************************************************
'**功能说明:
'**     以下函数主要是Lis系统等独立安装时,同时存在两种以上的数据库连接,
'**     如果使用ComLib,则mcnOracle连接就会冲突,因此,才建立以下模块
'**编制:刘兴洪
'*******************************************************************************************

'错误日志处理相关变量
Private mlngErrNum As Long, mstrErrInfo As String, mbytErrType As Byte
Private mlngADOErrNum As Long
Private mstrADOErrDesc As String
Private mblnADORoll As Boolean
Private mstrRecentSQL As String  '最近执行的SQL语句
Private mrsUserInfo As ADODB.Recordset
Private mrsParas As ADODB.Recordset '系统参数表缓存
Private mrsUserParas As ADODB.Recordset '系统参数表缓存
Private mrsDeptParas As ADODB.Recordset    '系统参数部门缓存
Private mblnRunLog As Boolean '是否记录使用日志
Private mblnErrLog As Boolean '是否记录运行错误
'ADO事务错误状态判断
Private WithEvents mcnOracle As ADODB.Connection
Attribute mcnOracle.VB_VarHelpID = -1

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mrsUserInfo Is Nothing Then Set mrsUserInfo = Nothing
    If Not mrsParas Is Nothing Then Set mrsParas = Nothing
    If Not mrsUserParas Is Nothing Then Set mrsUserParas = Nothing
    If Not mrsDeptParas Is Nothing Then Set mrsDeptParas = Nothing
End Sub

Private Sub mcnOracle_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    mblnADORoll = False: mlngADOErrNum = 0: mstrADOErrDesc = ""
End Sub

Private Sub mcnOracle_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    mblnADORoll = False: mlngADOErrNum = 0: mstrADOErrDesc = ""
End Sub

Private Sub mcnOracle_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    mblnADORoll = True
    If pConnection.Errors.Count > 0 Then
        mlngADOErrNum = pConnection.Errors(0).NativeError
        mstrADOErrDesc = pConnection.Errors(0).Description
    Else
        mlngADOErrNum = 0: mstrADOErrDesc = ""
    End If
End Sub

Private Function GetComputerName() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取计算机名称
    '返回:返回计算机名称
    '编制:刘兴洪
    '日期:2015-06-02 15:51:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp  As ADODB.Recordset, strSQL As String
    If mcnOracle Is Nothing Then Exit Function
    On Error GoTo errHandle
    strSQL = "Select SYS_CONTEXT('USERENV','TERMINAL') as MName From Dual"
    Set rsTmp = OpenSQLRecord(strSQL, "获取计算机名称")
    If rsTmp.EOF Then Exit Function
    GetComputerName = rsTmp!mname
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub InitCommon(cnMain As ADODB.Connection)
    '------------------------------------------------
    '功能:初始化连接
    '参数：
    '   cnMain:主程序的数据库连接
    '------------------------------------------------
    Set mcnOracle = cnMain
    If gstrComputerName = "" Then gstrComputerName = GetComputerName
End Sub
Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant, i As Long
    arrPars = arrInput
    Set OpenSQLRecord = OpenSQLRecordByArray(strSQL, strTitle, arrPars)
End Function

Public Function OpenSQLRecordByArray(ByVal strSQL As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
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
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    Dim strSQLTmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String

    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上
    strSQLTmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                strTmp = Replace(zlstr.FromatSQL(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSQL = "Select /*+ RULE*/" & Mid(Trim(strSQL), 7)
        End If
    End If
    
    
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
    
    If UBound(arrInput) + 1 < intMax Then
        Err.Raise 9527, strTitle, "SQL语句绑定变量不全，调用来源：" & strTitle
    End If

    '替换为"?"参数
    strLog = strSQL
    For i = 1 To intMax
        strSQL = Replace(strSQL, "[" & i & "]", "?")
        
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
    
    Set cmdData.ActiveConnection = mcnOracle '这句比较慢(这句执行1000次约0.5x秒)
    cmdData.CommandText = strSQL
    
    'SQLTEST中数据库连接的处理,所以直接用SQLTEST
    Call SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
    Call SQLTest
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
    Static mstrErrRecentSQL As String
    
    bytReturnType = 1
    If mcnOracle.Errors.Count <> 0 Then
        '处理事务中的错误不允许重试
        If mblnADORoll Then
            If mlngADOErrNum = mcnOracle.Errors(0).NativeError _
                And mstrADOErrDesc = mcnOracle.Errors(0).Description Then
                bytReturnType = 0
            End If
            mblnADORoll = False: mlngADOErrNum = 0: mstrADOErrDesc = ""
        End If
        
        'PL/SQL存储过程错误(包括嵌套的过程调用)
        strNote = mcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            '日志变量
            mbytErrType = 1
            mlngErrNum = mcnOracle.Errors(0).NativeError
            mstrErrInfo = mcnOracle.Errors(0).Description
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Function
        End If
       
        'ORACLE其它错误
        '日志变量
        mbytErrType = 2
        mlngErrNum = mcnOracle.Errors(0).NativeError
        mstrErrInfo = mcnOracle.Errors(0).Description
        
        Select Case mcnOracle.Errors(0).NativeError
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
            strNote = "表或视图不存在，很可能是你不具备使用该部分数据的权限或该部分对象同义词缺失。"
            bytReturnType = 0
            
            strTemp = mGetInvalidTable()
            If strTemp <> "" Then
                mstrErrInfo = "（建议请系统管理员重新授权、修复同义词）" & vbCrLf & "请对下列对象进行检查：" & vbCrLf & vbCrLf & vbTab & strTemp
            Else
                mstrErrInfo = "（建议请系统管理员重新授权、修复同义词）" & vbCrLf & "错误SQL语句为：" & vbCrLf & vbCrLf & mstrRecentSQL
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
                strNote = "未知错误，发生在" & mcnOracle.Errors(0).Source
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
    
    
    gblnAutoConnect = Val(GetSetting("ZLSOFT", "公共全局\网络断网自动重连", "AutoConnect", 0))
    If gblnAutoConnect Then '是否使用网络断开自动连接功能
          Dim blnConnect As Boolean
          Dim blnNumConnect As Boolean '检查次数是否重新连接
          Dim blnStatus As Boolean '是否其他错误引发的网络问题
          '通过过滤错误信息,检查是否为网络问题引发的错误。mbytErrType=2 Oracle提供的错误信息 mbytErrType=3 VB提供的错误信息
          If mbytErrType = 3 Then
              If mlngErrNum = -2147467259 Or mlngErrNum = -2147217900 Or mlngErrNum = 3709 Then
                  '检查VB具体错误信息
                  If CheckErrConnectInfo(mlngErrNum, strNote, mstrErrInfo, 1) Then
                        
                        '判断相同错误,如果2次以上正常错误提示。
                        If mstrErrRecentSQL = mstrRecentSQL And mstrRecentSQL <> "" Then
                            mlngConnectCount = mlngConnectCount + 1
                            If mlngConnectCount > 2 Then
                                blnNumConnect = False  '正常错误提示
                                mlngConnectCount = 0 '还原计数器
                            Else
                                blnNumConnect = True
                            End If
                        Else
                            mstrErrRecentSQL = mstrRecentSQL
                            mlngConnectCount = 1
                            blnNumConnect = True
                        End If
                  Else
                        blnConnect = False '正常错误提示
                  End If
              End If
          Else
              '错误号12543 TNS: 无法连接目标主机
              If mlngErrNum = -2147467259 Or mlngErrNum = -2147217900 Or mlngErrNum = 0 Or mlngErrNum = 12543 Then
                  '检查ORACLE具体错误信息
                  If CheckErrConnectInfo(mlngErrNum, strNote, mstrErrInfo, 2) Then
                        
                        '判断相同错误,如果2次以上正常错误提示。
                        If mstrErrRecentSQL = mstrRecentSQL And mstrRecentSQL <> "" Then
                            mlngConnectCount = mlngConnectCount + 1
                            If mlngConnectCount > 2 Then
                                blnNumConnect = False  '正常错误提示
                                mlngConnectCount = 0 '还原计数器
                            Else
                                blnNumConnect = True
                            End If
                        Else
                            mstrErrRecentSQL = mstrRecentSQL
                            mlngConnectCount = 1
                            blnNumConnect = True
                        End If
                  Else
                        blnConnect = False '正常错误提示
                  End If
              End If
          End If
          
          '自动重新连接一次,检查是否能自动重新连接
          If blnNumConnect Then '与ORACLE连接已经断开
              If CheckAdoConnction(blnStatus) Then
                  If blnStatus Then
                     blnConnect = False '正常错误提示
                  Else
                     blnConnect = True '提示重连
                  End If
              Else
                  '与ORACLE重新连接成功,不需要提示。直接返回重新执行。
                  blnConnect = False
                  ErrCenter = 1
                  Exit Function
              End If
          End If
    End If
    If bytReturnType = 1 Then
        ErrCenter = frmErrAsk.ShowEdit(mlngErrNum, strNote, mstrErrInfo, blnConnect, Me)
    Else
        Call frmErrNote.ShowEdit(mlngErrNum, strNote, mstrErrInfo, blnConnect, Me)
        ErrCenter = 0
    End If
    '清除错误
    Err.Clear
End Function

Public Function CheckErrConnectInfo(ByVal strErrNum As String, ByVal strNote As String, ByVal strErrInfo As String, ByVal intType As Integer) As Boolean
    '------------------------------------------------
    '功能： 按照类型IntType(1,2)检查vb和oralce返回的具体错误信息，来判断是否为网络断开引发的错误或者是其他的错误引发
    '参数： strNote错误信息,strErrInfo错误详细信息,intType 错误类型 1：VB错误 2:ORACLE错误
    '返回： True:网络引发的错误 False:其他错误
    '------------------------------------------------
    Dim strTemp As String
    Dim i As Integer
    If intType = 1 Then
        'VB具体错误
   
        If InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        Else
            If strErrNum = "3709" Then '3709描述：连接无法用于执行此操作。在此上下文中它可能已被关闭或无效。单独处理
                CheckErrConnectInfo = True
            Else
                If strNote = "不确定的错误" Then
                    CheckErrConnectInfo = True
                Else
                    CheckErrConnectInfo = False
                End If
            End If
        End If
    Else
        'ORACLE具体错误
        If InStr(strErrInfo, "SQLSetConnectAttr") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12560") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "E_FAIL") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12571") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-03114") > 0 Then
            CheckErrConnectInfo = True
        ElseIf InStr(strErrInfo, "ORA-12543") > 0 Then
            CheckErrConnectInfo = True
        Else
            CheckErrConnectInfo = False
        End If
    End If
End Function
Public Sub SaveErrLog()
'功能：将刚才的错误信息写入数据库错误日志
    Dim strSQL As String
    
    If mlngErrNum <> 0 And mbytErrType <> 0 And mblnErrLog Then
        On Local Error Resume Next
        
        strSQL = "Insert Into zlErrorLog(会话号,用户名,工作站,时间,类型,错误序号,错误信息)" & _
            " Select SID,USER,MACHINE,SYSDATE," & mbytErrType & "," & mlngErrNum & "," & _
            zlstr.SQLAdjust(mstrErrInfo) & " From v$Session Where AUDSID=UserENV('SessionID')"
        mcnOracle.Execute strSQL
        
        mlngErrNum = 0: mstrErrInfo = "": mbytErrType = 0
    End If
End Sub


Public Function SetPara(ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True, Optional ByVal lngDeptID As Long) As Boolean
'功能：设置指定的参数值
'参数：varPara=参数号或参数名，以数字或字符类型传入区分
'      strValue=要设置的参数值
'      lngSys=使用该参数的系统编号，如100
'      lngModual=使用该参数的模块号，如1230
'      blnSetup=调用模块是否有参数设置权限
'      lngDeptID=部门级参数
'返回：设置是否成功
    Dim strSQL As String
    Dim strResFilter As String
    '检查参数值，如果没有变化则不处理
    strSQL = GetPara(varPara, lngSys, lngModual)
    If strSQL = strValue Then SetPara = True: Exit Function
    
    On Error GoTo errH
    SetPara = True
    strSQL = "zl_Parameters_Update('" & varPara & "','" & strValue & "'," & lngSys & "," & lngModual & "," & IIf(blnSetup, 1, 0) & "," & lngDeptID & ")"
    Call ExecuteProcedure(strSQL, "SetPara")
    
    '更新缓存记录集，逻辑与zl_Parameters_Update保持一致
    '过滤条件
    If TypeName(varPara) = "String" Then
        strResFilter = "参数名='" & CStr(varPara) & "' And 模块=" & lngModual & " And 系统=" & lngSys
    Else
        strResFilter = "参数号=" & Val(varPara) & " And 模块=" & lngModual & " And 系统=" & lngSys
    End If
    
    mrsParas.Filter = strResFilter
    If mrsParas.EOF Then Exit Function
    '权限判断
    If Not blnSetup Then
        If mrsParas!部门 = 1 Then
            Exit Function
        '公共全局参数,固定需要权限
        ElseIf mrsParas!系统 <> 0 And mrsParas!模块 = 0 And mrsParas!私有 = 0 And mrsParas!本机 = 0 Then
            Exit Function
        '公共模块参数,固定需要权限
        ElseIf mrsParas!模块 = 0 And mrsParas!私有 = 0 And mrsParas!本机 = 0 Then
            Exit Function
        '要授权控制的本机公共模块
        ElseIf mrsParas!系统 <> 0 And mrsParas!模块 <> 0 And mrsParas!私有 = 0 And mrsParas!本机 = 1 And mrsParas!授权 = 1 Then
            Exit Function
        End If
    End If
    
    If mrsParas!部门 = 1 Then
        mrsDeptParas!Filter = "参数ID=" & mrsParas!id & " And 部门ID=" & lngDeptID
        If mrsDeptParas.EOF Then
            mrsDeptParas.AddNew
            mrsDeptParas!参数id = mrsParas!id
           mrsDeptParas!部门ID = lngDeptID
            mrsDeptParas!参数值 = strValue
            mrsDeptParas.Update
        Else
            mrsDeptParas!参数值 = strValue
            mrsDeptParas.Update
        End If
    ElseIf mrsParas!私有 = 1 Or mrsParas!本机 = 1 Then
        mrsUserParas.Filter = "参数ID=" & mrsParas!id & _
                    IIf(mrsParas!私有 = 1, " And 用户名='" & mrsParas!用户名 & "'", " And 用户名='NullUser'") & _
                    IIf(mrsParas!本机 = 1, " And 机器名='" & mrsParas!机器名 & "'", " And 机器名='NullMachine'")
        
        If mrsUserParas.EOF Then
            mrsUserParas.AddNew
            mrsUserParas!参数id = mrsParas!id
            mrsUserParas!用户名 = IIf(mrsParas!私有 = 1, mrsParas!用户名, "NullUser")
            mrsUserParas!机器名 = IIf(mrsParas!本机 = 1, mrsParas!机器名, "NullMachine")
            mrsUserParas!参数值 = strValue
            mrsUserParas.Update
        Else
            mrsUserParas!参数值 = strValue
            mrsUserParas.Update
        End If
    Else
        mrsParas!参数值 = strValue
        mrsParas.Update
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SetPara = False
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
    
    On Error GoTo errH
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
                    .Fields.Append rsClone.Fields(intFields).Name, IIf(rsClone.Fields(intFields).type = adNumeric, adDouble, rsClone.Fields(intFields).type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
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
                        .Fields.Append IIf(strFieldNameAlias = "", strFieldName, strFieldNameAlias), IIf(rsClone.Fields(strFieldName).type = adNumeric, adDouble, rsClone.Fields(strFieldName).type), rsClone.Fields(strFieldName).DefinedSize, adFldIsNullable '0:表示新增
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
                    .Fields(intFields).value = rsClone.Fields(arrFieldsName(intFields)).value
                Next
                .Update
                rsClone.MoveNext
            Loop
            If rsClone.RecordCount <> 0 Then .Filter = "": .MoveFirst
        End If
    End With
    
    Set CopyNewRec = rsTarget
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Append(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
    '功能：将指定记录集的数据添加到另一个记录集上
    '参数：rsSource=目标记录集
    '      rsAppend=数据记录集
    '      arrInput=字段对应规则，该参数不传时，默认两记录集结构相同，格式：arrInput(0):[记录集1].字段1,字段2...；arrInput(1)：[记录集2].字段1,字段2...
    '返回：是否成功
    '      rsSource=添加数据后的记录集
    Dim arrSource As Variant, arrAppend As Variant
    Dim i As Long, arrValues() As Variant
    Dim strTmp As String
    
    If rsAppend Is Nothing Then Append = True: Exit Function
    If rsAppend.RecordCount = 0 Then Append = True: Exit Function
    If rsSource Is Nothing Then Exit Function
    On Error GoTo errH
    If LBound(arrInput) = 2 Then
        arrSource = Split(arrInput(LBound(arrInput)), ",")
        arrAppend = Split(arrInput(UBound(arrInput)), ",")
        If UBound(arrSource) <> UBound(arrAppend) Then Exit Function
        ReDim arrValues(UBound(arrAppend)): rsAppend.MoveFirst
        Do While Not rsAppend.EOF
            For i = LBound(arrAppend) To UBound(arrAppend)
                arrValues(i) = rsAppend(arrAppend(i)).value
            Next
            rsSource.AddNew arrSource, arrValues
            Erase arrValues
            rsAppend.MoveNext
        Loop
    ElseIf LBound(arrInput) = 0 Then
        Do While Not rsAppend.EOF
            rsSource.AddNew
            For i = 0 To rsSource.Fields.Count - 1
                rsSource.Fields(i).value = rsAppend.Fields(i).value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
    End If
    
    Append = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer, Optional ByVal lngDeptID As Long) As String
'功能：读取指定的参数值
'参数：varPara=参数号或参数名，以数字或字符类型传入区分
'      lngSys=使用该参数的系统编号，如100
'      lngModual=使用该参数的模块号，如1230
'      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
'      blnNotCache=是否不从缓存中读取
'      arrControl=控件数组，如Array(Me.Text1, Me.CheckBox1)，用于函数内部自动处理对应控件的显示颜色，是否禁止设置。
'      blnSetup=调用模块是否有参数设置权限
'      intType=返回参数，返回参数类型
'      lngDeptID=部门级参数，根据部门读取
'返回：参数值，字符串形式
    Dim strSQL As String, i As Integer
    Dim blnNew As Boolean, blnEnabled As Boolean, blnNewRow As Boolean, blnNotExists As Boolean
    Dim strSqlFilter As String, strResFilter As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    intType = 0
    
    '过滤条件
    If TypeName(varPara) = "String" Then
        strResFilter = "参数名='" & CStr(varPara) & "' And 模块=" & lngModual & " And 系统=" & lngSys
        strSqlFilter = "参数名=[5] And Nvl(模块,0)=[3] And Nvl(系统,0)= [4] "
    Else
        strResFilter = "参数号=" & Val(varPara) & " And 模块=" & lngModual & " And 系统=" & lngSys
        strSqlFilter = "参数号=[6] And Nvl(模块,0)=[3] And Nvl(系统,0)=[4] "
    End If
    
    '参数缓存判断
    If mrsParas Is Nothing Then
        blnNew = True
    ElseIf mrsParas.State = 0 Then
        blnNew = True
    Else
        mrsParas.Filter = strResFilter
        blnNewRow = mrsParas.EOF
    End If
    
    If blnNew Or blnNewRow Then
        '参数表，获取参数特征
        strSQL = "Select ID,Nvl(系统,0) as 系统,Nvl(模块,0) as 模块,Nvl(私有,0) as 私有,Nvl(本机,0) as 本机,NVL(部门,0) as 部门,Nvl(授权,0) as 授权,参数号,参数名," & _
            " Nvl(参数值,缺省值) as 参数值,[1] as 用户名,[2] as 机器名 From zlParameters Where " & strSqlFilter
        Set rsTmp = OpenSQLRecord(strSQL, "GetPara", gstrDBUser, gstrComputerName, lngModual, lngSys, CStr(varPara), Val(varPara))
    
        If rsTmp.EOF Then
            blnNotExists = True
        Else
            If blnNewRow Then
                Call Append(mrsParas, rsTmp)
            Else
                Set mrsParas = CopyNewRec(rsTmp)
            End If
            
            mrsParas.Filter = strResFilter '再次过滤
            '获取部门级参数
            If mrsParas!部门 = 1 Then
                strSQL = "Select 参数id, 部门ID, 参数值" & vbNewLine & _
                        "From Zldeptparas" & vbNewLine & _
                        "Where 参数id = [1]"
                Set rsTmp = OpenSQLRecord(strSQL, "GetPara", Val(mrsParas!id), lngDeptID)
                
                If mrsDeptParas Is Nothing Then
                    Set mrsDeptParas = CopyNewRec(rsTmp)
                ElseIf mrsDeptParas.State = 0 Then
                    Set mrsDeptParas = CopyNewRec(rsTmp)
                Else
                    Call Append(mrsDeptParas, rsTmp)
                End If
            '获取用户或本机参数
            ElseIf mrsParas!私有 = 1 Or mrsParas!本机 = 1 Then
                strSQL = "Select 参数id, Nvl(用户名, 'NullUser') As 用户名, Nvl(机器名, 'NullMachine') As 机器名, 参数值" & vbNewLine & _
                        "From zlUserParas" & vbNewLine & _
                        "Where 参数id = [3]"
                        
                If mrsParas!私有 = 1 And mrsParas!本机 = 1 Then
                    strSQL = strSQL & " And 用户名=[1] And 机器名=[2]"
                ElseIf mrsParas!私有 = 1 Then
                    strSQL = strSQL & " And 用户名=[1] "
                Else
                    strSQL = strSQL & " And 机器名=[2]"
                End If
                
                Set rsTmp = OpenSQLRecord(strSQL, "GetPara", gstrDBUser, gstrComputerName, Val(mrsParas!id))
                
                If mrsUserParas Is Nothing Then
                    Set mrsUserParas = CopyNewRec(rsTmp)
                ElseIf mrsUserParas.State = 0 Then
                    Set mrsUserParas = CopyNewRec(rsTmp)
                Else
                    Call Append(mrsUserParas, rsTmp)
                End If
            End If
        End If
    End If

    If blnNotExists Then
        GetPara = strDefault
    Else
        '获取参数值
        If mrsParas!部门 = 1 Then
            mrsDeptParas.Filter = "参数ID=" & mrsParas!id & _
                 " And 部门ID=" & lngDeptID
            If Not mrsDeptParas.EOF Then
                GetPara = Nvl(mrsDeptParas!参数值, strDefault)
            Else
                GetPara = Nvl(mrsParas!参数值, strDefault)
            End If
        ElseIf mrsParas!私有 = 1 Or mrsParas!本机 = 1 Then
            mrsUserParas.Filter = "参数ID=" & mrsParas!id & _
                IIf(mrsParas!私有 = 1, " And 用户名='" & mrsParas!用户名 & "'", " And 用户名='NullUser'") & _
                IIf(mrsParas!本机 = 1, " And 机器名='" & mrsParas!机器名 & "'", " And 机器名='NullMachine'")
            If Not mrsUserParas.EOF Then
                GetPara = Nvl(mrsUserParas!参数值, strDefault)
            Else
                GetPara = Nvl(mrsParas!参数值, strDefault)
            End If
        Else
            GetPara = Nvl(mrsParas!参数值, strDefault)
        End If
        
        '返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
        If mrsParas!部门 = 1 Then
            intType = 3   '部门级当作模块公共参数处理
        ElseIf mrsParas!系统 <> 0 And mrsParas!模块 = 0 And mrsParas!私有 = 0 And mrsParas!本机 = 0 Then
            intType = 1
        ElseIf mrsParas!模块 = 0 And mrsParas!私有 = 1 And mrsParas!本机 = 0 Then
            intType = 2
        ElseIf mrsParas!系统 <> 0 And mrsParas!模块 <> 0 And mrsParas!私有 = 0 And mrsParas!本机 = 0 Then
            intType = 3
        ElseIf mrsParas!系统 <> 0 And mrsParas!模块 <> 0 And mrsParas!私有 = 1 And mrsParas!本机 = 0 Then
            intType = 4
        ElseIf mrsParas!系统 <> 0 And mrsParas!模块 <> 0 And mrsParas!私有 = 0 And mrsParas!本机 = 1 Then
            intType = IIf(mrsParas!授权 = 1, 15, 5)
        ElseIf mrsParas!系统 <> 0 And mrsParas!模块 <> 0 And mrsParas!私有 = 1 And mrsParas!本机 = 1 Then
            intType = 6
        End If
        
        '处理对应的控件颜色，可控状态
        If IsArray(arrControl) And (intType = 3 Or (intType Mod 10) = 5) Then
            blnEnabled = Not ((intType = 3 Or (intType Mod 10) = 5 And mrsParas!授权 = 1) And Not blnSetup)
            For i = 0 To UBound(arrControl)
                Select Case TypeName(arrControl(i))
                Case "Label"
                    arrControl(i).ForeColor = vbBlue
                Case "TextBox", "MaskEdBox", "CheckBox", "OptionButton", "ComboBox", "ListBox", "Frame", "PictureBox", "ListView"
                    arrControl(i).ForeColor = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "CommandButton", "DTPicker"
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "MSHFlexGrid"
                    arrControl(i).ForeColor = vbBlue
                    arrControl(i).ForeColorFixed = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                Case "VSFlexGrid"
                    arrControl(i).ForeColor = vbBlue
                    arrControl(i).ForeColorFixed = vbBlue
                    If Not blnEnabled Then arrControl(i).Editable = 0
                Case Else
                    On Error Resume Next
                    arrControl(i).ForeColor = vbBlue
                    If Not blnEnabled Then arrControl(i).Enabled = False
                    Err.Clear: On Error GoTo errH
                End Select
            Next
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function mGetInvalidTable() As String
'功能：得到在最近使用的SQL语句中不能访问的表或视图
    Dim varTables As Variant
    Dim strTable As String, lngCount As Long
    Dim strInvalidTable As String
    
    varTables = Split(SQLObject(mstrRecentSQL), ",")
    
    On Error Resume Next
    
    For lngCount = 0 To UBound(varTables)
        strTable = varTables(lngCount)
        
        '测试该对象是否可用
        mcnOracle.Execute "select 1 from " & strTable & " where rownum<1"
        If Err <> 0 Then
            Err.Clear
            strInvalidTable = strInvalidTable & "," & strTable
        End If
    Next
    
    If strInvalidTable <> "" Then
        '去掉第一个逗号
        mGetInvalidTable = Mid(strInvalidTable, 2)
    End If
End Function


Public Function SQLObject(ByVal strSQL As String) As String
'功能：分析SQL语句所用到的对象名
'参数：strSQL=要分析的原始SQL语句
'返回：SQL语句所访问到的对象名,如"部门表,病人费用记录,ZLHIS.人员表"
'说明：1.与Oracle SELECT语句兼容
'      2.如果SQL语句中的对象名前加有所有者前缀,则该前缀不会被截取
'      3.需要函数TrimChar;TrueObject的支持
    Dim intB As Integer, intE As Integer, intL As Integer, intR As Integer
    Dim strAnal As String, strSub As String, strObject As String
    Dim arrFrom() As String, strCur As String, strMulti As String, strTrue As String
    Dim i As Integer, j As Integer
    
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
    
    '分解分析
    arrFrom = Split(strAnal, "FROM")
    For i = 1 To UBound(arrFrom) '从第一个From后面部份开始
        strCur = arrFrom(i)
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
        For j = 0 To UBound(Split(strMulti, ","))
            strTrue = TrueObject(Split(strMulti, ",")(j))
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
    Dim i As Long, j As Long
    
    If Trim(Str) = "" Then TrimChar = "": Exit Function
    
    strTmp = Trim(Str)
    i = InStr(strTmp, "  ")
    Do While i > 0
        strTmp = Left(strTmp, i) & Mid(strTmp, i + 2)
        i = InStr(strTmp, "  ")
    Loop
    
    i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Do While i > 0
        strTmp = Left(strTmp, i + 1) & Mid(strTmp, i + 4)
        i = InStr(1, strTmp, vbCrLf & vbCrLf)
    Loop
    If Left(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 3)
    If Right(strTmp, 2) = vbCrLf Then strTmp = Mid(strTmp, 1, Len(strTmp) - 2)
    TrimChar = strTmp
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


'--------------------------------------------------
'功能：检查是否为网络断开或ADO断开引发的错误!
'返回：True:恢复连接成功 False恢复连接失败
'--------------------------------------------------
Public Function CheckAdoConnction(ByRef blnStatus As Boolean) As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strcnOracle As String
    Dim blnAdoErr As Boolean
    On Error GoTo Errhand
    blnAdoErr = False
    blnStatus = False

    strcnOracle = CStr(mcnOracle) '原有连接串需要保存,否则CLOSE后,原有连接串会被改变
    On Error GoTo Errhand
    Err = 0
    DoEvents
    With mcnOracle
        If .State = adStateOpen Then .Close
        .Open
    End With

    If blnAdoErr Then
        'True '是ORA-12560不能与ORACLE连接引起
        CheckAdoConnction = True
    Else
        'False '可以正常连接
        CheckAdoConnction = False
        On Error Resume Next
        '重连后判断客户端是否被禁止使用，若被禁止，则自动断开连接
        strSQL = "Select NVL(禁止使用,0)  禁止使用 From zlClients Where 工作站=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "CheckAdoConnction", gstrComputerName)
        If Err.Number <> 0 Then Err.Clear
        If Not rsTmp Is Nothing Then
            If Not rsTmp.EOF Then
                If rsTmp!禁止使用 = 1 Then
                    If mcnOracle.State = adStateOpen Then mcnOracle.Close
                    CheckAdoConnction = True
                    Call SaveSetting("ZLSOFT", "公共全局\网络断网自动重连", "AutoConnect", 0)
                    MsgBox "当前工作站已经被管理员禁用，请联系管理员解除禁用并重新登录！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    Exit Function
Errhand:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        If InStr(Err.Description, "ORA-12560") > 0 Then
            mcnOracle = strcnOracle
            blnAdoErr = True
            Resume Next
        ElseIf InStr(Err.Description, "ORA-12543") > 0 Then
            mcnOracle = strcnOracle
            blnAdoErr = True
            Resume Next
        Else
            '其他错误引发的网络问题
            CheckAdoConnction = True
            blnStatus = True
        End If
    Else
        CheckAdoConnction = False
    End If
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
                        intMax = LenB(StrConv(strPar, vbFromUnicode))
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
        
        Set cmdData.ActiveConnection = mcnOracle '这句比较慢
        cmdData.CommandType = adCmdText
        cmdData.CommandText = strProc
        
        Call SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
        Call SQLTest
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    Call SQLTest(App.ProductName, strFormCaption, strSQL)
    
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    mcnOracle.Execute strSQL, , adCmdText
    
    Call SQLTest
End Sub

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0
    On Error GoTo errH
    strSQL = "SELECT SYSDATE FROM DUAL"
    Set rsTemp = OpenSQLRecord(strSQL, "获取当前服务器系统日期")
    Currentdate = rsTemp.Fields(0).value
    rsTemp.Close
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Currentdate = 0
    Err = 0
End Function
 Public Function GetUserInfo() As ADODB.Recordset
'功能：获取当前用户的基本信息
'返回：返回Ado记录集
    Dim strSQL As String, strDefault As String, blnNew As Boolean
    
    If mrsUserInfo Is Nothing Then
        blnNew = True
    ElseIf mrsUserInfo.State = 0 Then
        blnNew = True
    End If
        
    If blnNew Then
        strDefault = " And C.缺省 = 1"
        strSQL = "Select User,A.Id, A.编号, A.简码, A.姓名, A.专业技术职务,B.用户名, C.部门id, D.编码 As 部门码, D.名称 As 部门名" & vbNewLine & _
            "From 人员表 A, 上机人员表 B, 部门人员 C, 部门表 D" & vbNewLine & _
            "Where A.Id = B.人员id And A.Id = C.人员id And C.部门id = D.Id And B.用户名 = [1]"
        On Error GoTo errH
        Set mrsUserInfo = OpenSQLRecord(strSQL & strDefault, "GetUserInfo", gstrDBUser)
        If mrsUserInfo.RecordCount = 0 Then
            strDefault = " And Rownum < 2"
            Set mrsUserInfo = OpenSQLRecord(strSQL & strDefault, "GetUserInfo", gstrDBUser)
        End If
    End If
    Set GetUserInfo = mrsUserInfo
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPrivFunc(lngSys As Long, lngProgId As Long) As String
'功能：返回当前用户具有的指定程序的功能串
'参数：lngSys     如果是固定模块，则为0
'      lngProgId  程序序号
'返回：分号间隔的功能串,为空表示没有权限
    Dim rsTmp As ADODB.Recordset, blnNew As Boolean
    Dim strSQL As String, strPrivs As String
    Dim blnRegCheck As Boolean
        
    If gcolPrivs Is Nothing Then
        Set gcolPrivs = New Collection
        blnNew = True
    Else
        On Error Resume Next
        strPrivs = gcolPrivs("_" & lngSys & "_" & lngProgId)
        If Err.Number > 0 Then blnNew = True: Err.Clear
    End If
    
    If blnNew Then
        strSQL = "Select Text as 功能 From Table(Cast(zltools.f_Reg_Func([1],[2]) as zlTools.t_Reg_Rowset))"
        On Error GoTo errH
Beging:
        Set rsTmp = OpenSQLRecord(strSQL, "GetPrivFunc", lngSys, lngProgId)
        On Error GoTo 0
        
        Do While Not rsTmp.EOF
            strPrivs = strPrivs & ";" & rsTmp!功能
            rsTmp.MoveNext
        Loop
        strPrivs = Mid(strPrivs, 2)
        gcolPrivs.Add strPrivs, "_" & lngSys & "_" & lngProgId
    End If
        
    GetPrivFunc = strPrivs
    Exit Function
errH:
    If Not blnRegCheck Then
        '如果出错,可能是由于没有调用RegCheck造成,自动调用一次,如果再出错,才提示.
        If RegCheck <> "" Then Exit Function
        blnRegCheck = True
        GoTo Beging
    End If
    If ErrCenter() = 1 Then Resume
End Function

Private Function RegCheck() As String
    '功能：注册检查检查
    '说明:不能使用zlRegister或zl9Commlib，原因是存在两个连接(Lis和HIS)，为了避免不能取权限串，所以需要先执行一下
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select zltools.f_Reg_Audit([1]) As Stamp From zltools.zlRegInfo r Where 项目='授权证章'"
    Set rsTemp = OpenSQLRecord(strSQL, "RegCheck", 0)
    If rsTemp.RecordCount > 0 Then
        If Left(rsTemp.Fields(0).value, 6) <> "ERROR-" Then
            RegCheck = ""
        Else
            RegCheck = rsTemp.Fields(0).value
        End If
    Else
        RegCheck = "注册信息丢失,在重新注册前"
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
