VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataOracle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
Private gstrSysName As String
Private mcnOracle As ADODB.Connection
Private mstrSQL As String
Private mobjRegister As Object

'######################################################################################################################

Public Property Get DatabaseState() As Long
    DatabaseState = mcnOracle.State
End Property

Public Property Get DatabaseConnection() As ADODB.Connection
    Set DatabaseConnection = mcnOracle
End Property

Public Sub BeginTrans()
    mcnOracle.BeginTrans
End Sub

Public Sub CommitTrans()
    mcnOracle.CommitTrans
End Sub

Public Sub RollbackTrans()
    mcnOracle.RollbackTrans
End Sub

Public Function UpdateUserPassword(ByVal strUserName As String, ByVal strPasswd As String)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Call CreateRegister
    Call mobjRegister.UpdateUserPassword(mcnOracle, strUserName, strPasswd, True)
End Function

Public Sub SetRunWithRBO()
    '******************************************************************************************************************
    '功能：当前会话以RBO优化器模式运行SQL语句
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    
    mstrSQL = "alter session set optimizer_mode=rule"
    On Error Resume Next
    mcnOracle.Execute strSQL
    If Err.Number = 0 Then
        MsgBox "已设置当前会话以RBO优化器模式运行！", vbInformation, gstrSysName
    End If
End Sub

Public Function SetSQLTrace(ByVal strUserName As String, ByVal strServerName As String) As String
    '******************************************************************************************************************
    '功能：调用100046事件启动SQL Trace功能
    '参数：
    '返回：Trc文件名
    '******************************************************************************************************************
    Dim strSQL As String, strLevel As String, strFile As String
    
    strServerName = UCase(strServerName)
    
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSQL = "alter session set timed_statistics=true"
        mcnOracle.Execute strSQL
        strSQL = "alter session set max_dump_file_size=10M"
        mcnOracle.Execute strSQL
        Err.Clear
        
        '下面这一条语句在8.1.7及以后才支持
        strFile = "ZL_" & strUserName
        strSQL = "alter session set tracefile_identifier='" & strFile & "'"
        mcnOracle.Execute strSQL
        If Err.Number <> 0 Then strFile = "*.trc": Err.Clear
        
        strLevel = "1"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSQL = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        mcnOracle.Execute strSQL
        If Err.Number = 0 Then SetSQLTrace = strFile
    End If
End Function

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional ByVal blnOLEDB As Boolean = False) As Boolean
    '******************************************************************************************************************
    '功能： 打开指定的数据库
    '参数： strServerName：主机字符串
    '       strUserName：用户名
    '       strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '******************************************************************************************************************
    Dim strSQL As String
    Dim strError As String
    Dim blnTransPassword As Boolean
    Call CreateRegister
    On Error Resume Next
    blnTransPassword = Not (UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM" Or UCase(strUserName) = "ZLTOOLS")
    Set mcnOracle = mobjRegister.GetConnection(strServerName, strUserName, strUserPwd, blnTransPassword, IIf(blnOLEDB, 1, 0))
    If mcnOracle.State = adStateClosed Then
        Exit Function
    End If
    Err = 0
    OraDataOpen = True
    Exit Function
    
errHand:
    OraDataOpen = False
    Err = 0
End Function

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
    Dim arrPar As Variant, i As Integer
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
    For i = 0 To UBound(varParValue)
        varValue = varParValue(i)
        Select Case TypeName(varValue)
            Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarNumeric, adParamInput, 30, varValue)
            Case "String" '字符
                intMax = LenB(StrConv(varValue, vbFromUnicode))
                If intMax = 0 Or intMax < 10 Then intMax = 10
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adVarChar, adParamInput, intMax, varValue)
            Case "Date" '日期
                cmdPackage.Parameters.Append cmdPackage.CreateParameter("P" & i, adDBTimeStamp, adParamInput, , varValue)
        End Select
        strNode = strNode & CStr(varValue) & ","
    Next

    If cmdPackage.ActiveConnection Is Nothing Then
        Set cmdPackage.ActiveConnection = mcnOracle
    End If
    
    
    cmdPackage.CommandType = adCmdStoredProc
    cmdPackage.CommandText = strPackagesName
    
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strPackagesName, strNode)
    cmdPackage.Properties("PLSQLRSet") = True
    Set OpenCursor = cmdPackage.Execute
    OpenCursor.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
    cmdPackage.Properties("PLSQLRSet") = False

End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    '******************************************************************************************************************
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
    '******************************************************************************************************************
    
    Dim cmdData As New ADODB.Command
    Dim strPar As String, arrPar As Variant
    Dim lngLeft As Long, lngRight As Long
    Dim strSeq As String, intMax As Integer, i As Integer
    Dim strLog As String, varValue As Variant
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSQL, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSQL, "]")
        
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
'    cmdData.CommandText = "" '不为空有时清除参数出错
'    Do While cmdData.Parameters.Count > 0
'        cmdData.Parameters.Delete 0
'    Loop
    
    '创建新的参数
    lngLeft = 0: lngRight = 0
    arrPar = Split(Mid(strPar, 2), ",")
    For i = 0 To UBound(arrPar)
        varValue = arrInput((arrPar(i) - 1))
        Select Case TypeName(varValue)
        Case "Byte", "Integer", "Long", "Single", "Double", "Currency" '数字
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

    '执行返回记录集
    'If cmdData.ActiveConnection Is Nothing Then
        Set cmdData.ActiveConnection = mcnOracle '这句比较慢(这句执行1000次约0.5x秒)
    'End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
'    Call gobjComLib.SQLTest
End Function

Public Function ExecuteSQL(strSQL As String, ByVal strFormCaption As String, Optional ByVal blnProc As Boolean = False) As Boolean
    
    On Error GoTo errHand
    
    If blnProc Then
        mcnOracle.Execute strSQL, , adCmdStoredProc
    Else
        mcnOracle.Execute strSQL
    End If
    
    ExecuteSQL = True
    
    Exit Function
    
errHand:
    MsgBox Err.Description
End Function

Public Function ExecuteProcedure(strSQL As String, ByVal strFormCaption As String) As Boolean
    '******************************************************************************************************************
    '功能：执行过程语句,并自动对过程参数进行绑定变量处理
    '参数：strSQL=过程语句,可能带参数,形如"过程名(参数1,参数2,...)"。
    '说明：以下几种情况过程参数不使用绑定变量,仍用老的调用方法：
    '  1.参数部份是表达式,这时程序无法处理绑定变量类型和值,如"过程名(参数1,100.12*0.15,...)"
    '  2.中间没有传入明确的可选参数,这时程序无法处理绑定变量类型和值,如"过程名(参数1, , ,参数3,...)"
    '  3.因为该过程是自动处理,不是一定使用绑定变量,对带"'"的字符参数,仍要使用"''"形式。
    '******************************************************************************************************************
    
    Dim cmdData As New ADODB.Command
    Dim strProc As String, strPar As String
    Dim blnStr As Boolean, intBra As Integer
    Dim strTemp As String, i As Long
    Dim intMax As Integer, datCur As Date
    
    If Right(Trim(strSQL), 1) = ")" Then
        '清除原有参数:不然不能重复执行
'        cmdData.CommandText = "" '不为空有时清除参数出错
'        Do While cmdData.Parameters.Count > 0
'            cmdData.Parameters.Delete 0
'        Loop
        
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
            Exit Function
        End If
        
        '补充?号
        strTemp = ""
        For i = 1 To cmdData.Parameters.Count
            strTemp = strTemp & ",?"
        Next
        strProc = "Call " & strProc & "(" & Mid(strTemp, 2) & ")"
        
        '执行过程
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = mcnOracle '这句比较慢
            cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc
        
'        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
'        Call gobjComLib.SQLTest
    Else
        GoTo NoneVarLine
    End If
    
    ExecuteProcedure = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
NoneVarLine:
'    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
    
    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    mcnOracle.Execute strSQL, , adCmdText
    
    ExecuteProcedure = True
    
'    Call gobjComLib.SQLTest
End Function

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", mcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    Currentdate = 0
    Err = 0
End Function

'######################################################################################################################
Private Sub Class_Initialize()
    Set mcnOracle = New ADODB.Connection
End Sub

Private Sub Class_Terminate()
    If Not (mcnOracle Is Nothing) Then
        If mcnOracle.State = adStateOpen Then mcnOracle.Close
        Set mcnOracle = Nothing
    End If
End Sub


Private Sub CreateRegister()
    '启动程序在登录之前已创建，以及医保部件等可能再次调用，避免重复创建部件
    If mobjRegister Is Nothing Then
        On Error Resume Next
        Set mobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
        
        If mobjRegister Is Nothing Then
            '用于支持未通过导航台（启动程序prjMain）调用本部件的情况。
            '这种情况下，由于没有调用GetConnection,无法通过该对象获取服务器名、用户名、切换连接
            Set mobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If mobjRegister Is Nothing Then
                MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
    End If
End Sub
