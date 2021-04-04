Attribute VB_Name = "mdlMain"
Public gcnOracle As New ADODB.Connection    '公共数据库连接

Public gblnDBA As Boolean                   '是否DBA
Public gstrSQL    As String                 '通用的SQL语句变量
Public gstrSysName As String                '系统名称
Public gstrUserName As String               '用户名
Public gstrPassword As String               '用户口令
Public gstrServer As String                 '服务器名
Public gstrFilePath As String           '文件保存的默认路径

'非数据库连接公共参数
Public gblnHadInit As Boolean        '是否已经初始化
Public gblnIsZlhis As Boolean           '是否为ZLHIS环境
Public gstrVerNum As String           '数据库版本号
Public gstrBigVer As String              '数据库大版本
Public gblnRAC As Boolean               '是否为Rac环境
Public gintCpuCount  As Integer, gintCpuAdvise As Integer, gintCpuMax As Integer   'CPU现状以及建议并行度
Public gintInstId As Integer            'RAC环境下,当前实例ID
Public gblnHasBigtables As Boolean '记录是否有Bigtables这张表
Public gblnHasZltables As Boolean '记录是否有zltable这张表

'API相关
Public Const WM_SYSCOMMAND = &H112
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_RESTORE = &HF120&
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

'公用颜色
Public Enum rowColor
    FULL_颜色 = &HB3DEF5
    BackAlterNate_颜色 = &HF0FFF0
    Back_颜色 = &H80000005
    Used_颜色 = &HEEEEE0
    OFF_颜色 = &HB3DEF5
End Enum

'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long


Public Function OpenSQLRecordByArray(ByVal strSql As String, ByVal strTitle As String, arrInput() As Variant) As ADODB.Recordset
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
    strSQLTmp = Trim(UCase(strSql))
    If Mid(Trim(Mid(strSQLTmp, 7)), 1, 2) <> "/*" And Mid(strSQLTmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLTmp
            Do While InStr(strSQLtmp1, arrstr(i)) > 0
                '判断前面是否用了IN 用了则不加Rule
                '先找到最近一个SELECT
                strTmp = Mid(strSQLtmp1, 1, InStr(strSQLtmp1, arrstr(i)) - 1)
                'strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
                If Len(strTmp) > 1 Then strTmp = Mid(strTmp, Len(strTmp) - 2)  '取后面3个字符
                
                If strTmp = "IN(" Then '属于in(select这种情况，则继续循环，看是否存在没有使用这种写法的其他动态内存函数
                   strSQLtmp1 = Mid(strSQLtmp1, InStr(strSQLtmp1, arrstr(i)) + Len(arrstr(i)))
                Else
                    Exit For
                End If
            Loop
        Next
        If i <= UBound(arrstr) Then
            strSql = "Select /*+ RULE*/" & Mid(Trim(strSql), 7)
        End If
    End If
    
    
    '分析自定的[x]参数
    lngLeft = InStr(1, strSql, "[")
    Do While lngLeft > 0
        lngRight = InStr(lngLeft + 1, strSql, "]")
        
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

    '执行返回记录集
    Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
 
    cmdData.CommandText = strSql
    
    
    Set OpenSQLRecordByArray = cmdData.Execute
    Set OpenSQLRecordByArray.ActiveConnection = Nothing
    
End Function


Public Function OpenSQLRecord(ByVal strSql As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    Dim arrPars() As Variant
    arrPars = arrInput
    Set OpenSQLRecord = OpenSQLRecordByArray(strSql, strTitle, arrPars)
End Function
Public Sub InitTable(vsf As VSFlexGrid, strCol As String)
'功能: 初始化表头
    Dim arrHead As Variant
    Dim i As Long
    
    arrHead = Split(strCol, ";")
   
    With vsf
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows
        .Editable = flexEDNone
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0)
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub

Public Function IsInstallExcel() As Boolean
'功能：判断本机上装有EXCEL没有
'参数：
'返回：有则返回True
    Dim objTemp  As Object
    
    On Error GoTo errH
    Set objTemp = CreateObject("Excel.Application") '打开一个EXCEL程序
    Set objTemp = Nothing
    IsInstallExcel = True
    Exit Function
errH:
    Set objTemp = Nothing
    IsInstallExcel = False
    Err.Clear
End Function

Public Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'说明：主要是RunSQLFile的子函数
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
    TrimEx = strText
End Function

Public Sub ErrCenter(Optional strErr As String)
    MsgBox Err.Description & vbCrLf & strErr, vbExclamation, "错误"
End Sub


Public Sub CreateStr2list()
    '功能：检查是否存在F_STR2LIST2函数，若不存在则添加。
    Dim strSql As String, rsData As ADODB.Recordset
    
    If mblnZlhis Then Exit Sub  'zlhis环境直接退出
    
    On Error Resume Next
    '判断是否需要创建函数
    strSql = "select  1 from all_objects where object_name ='F_STR2LIST2' and OWNER='PUBLIC'  and Object_type ='SYNONYM'"
    Set rsData = OpenSQLRecord(strSql, "CreateStr2list")
    If rsData.RecordCount > 0 Then Exit Sub
    
    '创建-1.创建类型
    strSql = "CREATE OR REPLACE Type t_StrObj2 as object (C1 Varchar2(4000),C2 Varchar2(4000))"
    gcnOracle.Execute strSql
    strSql = "CREATE OR REPLACE Type t_StrList2 as table of t_StrObj2"
    gcnOracle.Execute strSql
    
    '创建-2.创建函数
    strSql = "Create Or Replace Function f_Str2list2" & vbNewLine & _
                    "(" & vbNewLine & _
                    "  Str_In      In Varchar2,Split_In    In Varchar2 := ',', Subsplit_In In Varchar2 := ':'" & vbNewLine & _
                    ") Return t_Strlist2" & vbNewLine & _
                    "  Pipelined As" & vbNewLine & _
                    "  v_Str   Long; P       Number; v_Tmp   Varchar2(4000);" & vbNewLine & _
                    "   Out_Rec t_Strobj2 := t_Strobj2(Null, Null);" & vbNewLine & _
                    "Begin" & vbNewLine & _
                    "  If Str_In Is Null Then" & vbNewLine & _
                    "    Return;" & vbNewLine & _
                    "  End If;" & vbNewLine & _
                    "  v_Str := Str_In || Split_In;" & vbNewLine & _
                    "  Loop" & vbNewLine & _
                    "    P := Instr(v_Str, Split_In);Exit When(Nvl(P, 0) = 0);v_Tmp      := Substr(v_Str, 1, P - 1);Out_Rec.C1 := Substr(v_Tmp, 1, Instr(v_Tmp, Subsplit_In) - 1);" & vbNewLine & _
                    "    Out_Rec.C2 := Substr(v_Tmp, Instr(v_Tmp, Subsplit_In) + 1); Pipe Row(Out_Rec);v_Str := Substr(v_Str, P + 1);" & vbNewLine & _
                    "  End Loop;" & vbNewLine & _
                    "  Return;" & vbNewLine & _
                    "End;"
    gcnOracle.Execute strSql
    
    '创建-3.添加同义词
    strSql = "create or replace synonym F_STR2LIST2 for f_Str2list2"
    gcnOracle.Execute strSql
    
    '创建-4.同义词授权
    strSql = " grant execute on  F_STR2LIST2 to public"
    gcnOracle.Execute strSql
    
    If Err.Number > 0 Then
        MsgBox Err.Description
    End If
End Sub

Public Function CheckTblExist(ByVal strTableName As String) As Boolean
    '功能：根据表名判断表是否存在
    '参数：strTableName - 要查询的表名
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error GoTo errH
    strSql = "select 1 from dba_all_tables where table_name =[1] "
    Set rsData = OpenSQLRecord(strSql, "CheckTblExist", strTableName)
    CheckTblExist = (rsData.RecordCount > 0)
    
    Exit Function
errH:
    MsgBox Err.Description
End Function

Public Function GetPrevSQLID(ByRef strChildNum As String) As String
'功能：获取当前会话最近一次执行的SQLID,并将语句的CHILD_NUMBER赋值到传入变量中
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    On Error GoTo errH
    strSql = "select prev_sql_id,PREV_CHILD_NUMBER from V$session where AUDSID=UserENV('SessionID')"
    Set rsTmp = OpenSQLRecord(strSql, "GetPrevSQLID")
 
    If rsTmp.RecordCount = 0 Then Exit Function
    GetPrevSQLID = rsTmp!prev_sql_id & ""
    strChildNum = rsTmp!PREV_CHILD_NUMBER & ""
    
    Exit Function
errH:
    MsgBox Err.Description
End Function

Public Function GetCurrentdate() As Date
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
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    GetCurrentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    GetCurrentdate = 0
    Err = 0
End Function



Public Function GetTimeString(ByVal datBegin As Date, ByVal datEnd As Date) As String
'功能：获取两个时间值差的格式字符串
'   datBegin=起始时间
'   datEnd=中止时间
    Dim intH As Integer, intM As Integer, intS As Integer
    Dim datTmp As Date

    intH = DateDiff("h", datBegin, datEnd)
    datTmp = DateAdd("h", intH, datBegin)
    intM = DateDiff("n", datTmp, datEnd)
    datTmp = DateAdd("n", intM, datTmp)
    intS = DateDiff("s", datTmp, datEnd)
    
    If intS < 0 Then
        intM = intM - 1
        intS = 60 + intS
    End If
    
    If intM < 0 Then
        intH = intH - 1
        intM = 60 + intM
    End If
    GetTimeString = IIf(intH <> 0, intH & "小时", "") & IIf(intM <> 0, intM & "分", "") & intS & "秒"
End Function

Public Function getVersion() As String
'功能：获取数据库的大版本号
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim arrTmp As Variant
    
    On Error GoTo errH
    'CORE    10.2.0.3.0  Production
    strSql = "Select Banner From V$version Where Banner Like  'CORE%'"
    Set rsTmp = OpenSQLRecord(strSql, App.Title)
    If rsTmp.RecordCount > 0 Then
        arrTmp = Split(TrimEx(rsTmp!Banner & ""), " ")
        If UBound(arrTmp) = 2 Then
            getVersion = Mid(arrTmp(1), 1, InStr(1, arrTmp(1), ".") - 1)
        End If
    End If
    
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, "错误"
End Function

Public Function GetOracleVersion(Optional ByVal blnGetVerNum As Boolean = False) As String
    '功能：获取数据库的版本号，默认返回数据库大版本号
    '参数：blnGetVerNum-是否返回数据库完整版本号
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strTmp As String
    Dim arrTmp As Variant
        
    On Error GoTo errH
    'CORE    10.2.0.3.0  Production
    strSql = "Select Banner From V$version Where Banner Like  'CORE%'"
    Set rsTmp = OpenSQLRecord(strSql, App.Title)
    If rsTmp.RecordCount > 0 Then
        arrTmp = Split(TrimEx(rsTmp!Banner & ""), " ")
        If UBound(arrTmp) = 2 Then
            strTmp = arrTmp(1)
        End If
    End If
    
    '10.2.0.3.0
    If Not blnGetVerNum Then
        arrTmp = Split(strTmp, ".")
        strTmp = Val(arrTmp(0))
    End If
    
    GetOracleVersion = strTmp
    Exit Function
errH:
    ErrCenter "获取数据库版本失败，部分功能将无法使用。"
End Function

Public Function CheckRAC(ByRef intInstID As Integer) As Boolean
'功能：检查是否为RAC环境
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    strSql = "select 1 from gv$active_instances"
    On Error GoTo errH
    Set rsTmp = OpenSQLRecord(strSql, "CheckRAC")
    
    If rsTmp.RecordCount > 0 Then
        CheckRAC = True
        
        strSql = "Select UserENV('instance') Inst_ID From dual"
        Set rsTmp = OpenSQLRecord(strSql, "CheckRAC")
        intInstID = "" & rsTmp!Inst_ID
    Else
        CheckRAC = False
    End If
    
    Exit Function
errH:
    ErrCenter
End Function
Public Function GetCpuCount(ByRef intAdvise As Integer, ByRef intMax As Integer) As String
'功能：设置统计信息收集以及并行DDL的并行度
'返回值： 服务器CPU个数， intDefault 建议并行度，inxMax 最大并行度
    Dim strSql As String, rsTmp As ADODB.Recordset
    
     '最大并行为CPU数，防止过高，实际为CPU个数*单个CPU上并行进程
    On Error GoTo errH
    strSql = "Select Nvl(Max(Value),0) CPU From " & IIf(gblnRAC, "G", "") & "V$parameter Where Name = 'cpu_count'" & IIf(gblnRAC, "And INST_ID = " & gintInstId & " ", "") & " "
    Set rsTmp = OpenSQLRecord(strSql, "获取可用CUP数")
    
    If rsTmp!cpu <= 4 Then
        intAdvise = 1
        intMax = IIf(rsTmp!cpu = 0, 1, rsTmp!cpu)
    ElseIf rsTmp!cpu <= 8 Then
        intAdvise = 4
        intMax = rsTmp!cpu
    ElseIf rsTmp!cpu <= 12 Then
        intAdvise = 8
        intMax = rsTmp!cpu
    Else
        intAdvise = 12
        intMax = rsTmp!cpu
    End If
    
    GetCpuCount = rsTmp!cpu
    Exit Function
errH:
    ErrCenter "获取服务器CPU参数失败，部分功能无法使用。"
End Function

Public Function IsCharChinese(ByVal strAsk As String) As Boolean
    '-------------------------------------------------------------
    '功能：判断指定字符串是否含有汉字
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    Dim i As Integer, j As Integer
    
    If Len(Trim(strAsk)) > 0 Then
        For i = 1 To Len(Trim(strAsk))
            j = Asc(Mid(Trim(strAsk), i, 1))
            If j < 0 Then
                IsCharChinese = True
                Exit Function
            End If
        Next
    End If
    IsCharChinese = False
End Function

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = False
End Function

Public Sub ClearVsf(vsfGrid As VSFlexGrid, strNone As String)
'功能：清空表格，并在fixedRow的下一行添加提示信息
    With vsfGrid
        .Redraw = flexRDNone
        .Rows = .FixedRows
        
        If strNone = "" Then
            .Rows = .FixedRows
        Else
            .Rows = .FixedRows + 1
            .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = strNone
        End If
        
        .MergeCells = flexMergeRestrictRows
        .MergeRow(-1) = True
        .AutoResize = True
        .AutoSize 0, .Cols - 1, False
        .Redraw = flexRDDirect
        .Row = .Rows - 1
    End With
End Sub

Public Sub CreateList2str()
    '修改或新建f_List2str函数s
    Dim strSql As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    
    '如果函数存在，则不用创建
    strSql = "Select 1 From " & IIf(gblnIsZlhis, "DBA_ARGUMENTS", "USER_ARGUMENTS") & " Where object_name = 'F_LIST2STR' And argument_name = 'P_MAXLENGTH'"
    Set rsTmp = OpenSQLRecord(strSql, "CreateList2str")
    If rsTmp.RecordCount > 0 Then Exit Sub
    
    '如果不是ZLHIS环境，需要创建类型 create or replace type zltools.t_StrList as Table of Varchar2(4000)
    If Not gblnIsZlhis Then
        strSql = "Select 1 From user_types  Where type_name = 'T_STRLIST'"
        Set rsTmp = OpenSQLRecord(strSql, "CreateList2str")
        If rsTmp.RecordCount = 0 Then
            strSql = "create or replace type t_StrList as Table of Varchar2(4000)"
            gcnOracle.Execute strSql
        End If
    End If
    
    '创建函数
    strSql = "CREATE OR REPLACE Function " & IIf(gblnIsZlhis, "Zltools.", "") & "f_List2str" & vbNewLine & _
                    "(" & vbNewLine & _
                    "  p_Strlist   In t_Strlist,p_Delimiter In Varchar2 Default ',', p_Distinct  In Number Default 1, p_Maxlength In Number Default 0" & vbNewLine & _
                    ") Return Varchar2 Is" & vbNewLine & _
                    "l_String Long;l_Add    Number;" & vbNewLine & _
                    "Begin" & vbNewLine & _
                    "  If p_Strlist.Count > 0 Then" & vbNewLine & _
                    "    For I In p_Strlist.First .. p_Strlist.Last Loop" & vbNewLine & _
                    "      l_Add := 0;" & vbNewLine & _
                    "      If p_Distinct = 1 Then" & vbNewLine & _
                    "        If Instr(',' || l_String || ',', ',' || p_Strlist(I) || ',') = 0 Then l_Add := 1; End If;" & vbNewLine & _
                    "      Else l_Add := 1; End If;" & vbNewLine & _
                    "      If l_Add = 1 Then If I != p_Strlist.First Then  l_String := l_String || p_Delimiter;End If;" & vbNewLine & _
                    "        l_String := l_String || p_Strlist(I);If p_Maxlength <> 0 And Length(l_String) > p_Maxlength Then" & vbNewLine & _
                    "        l_String := Substr(l_String, 1, p_Maxlength); Return l_String; End If;" & vbNewLine & _
                    "      End If;" & vbNewLine & _
                    "    End Loop;" & vbNewLine & _
                    "  End If;" & vbNewLine & _
                    "  Return l_String;" & vbNewLine & _
                    "End f_List2str;"
    gcnOracle.Execute strSql
    
    '如果是ZLHIS，就创建同义词，如果不是ZLHIS，那么就将函数保存在用户下。
    If gblnIsZlhis Then
        '创建-3.添加同义词
        strSql = "create or replace public synonym F_LIST2STR for ZLTOOLS.f_List2str"
        gcnOracle.Execute strSql
        
        '创建-4.同义词授权
        strSql = " grant execute on  F_LIST2STR to public"
        gcnOracle.Execute strSql
    End If
        
    Exit Sub
errH:
    ErrCenter
End Sub

Public Function CurrentDate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo errH
    '不能调用OpenSQLRecord,因为OpenSQLRecord也使用了该方法
    With rsTemp
        .CursorLocation = adUseClient
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    CurrentDate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
errH:
    If MsgBox(Err.Description, vbRetryCancel, gstrSysName) = vbRetry Then
        Resume
    End If
    CurrentDate = 0
    Err = 0
End Function

Public Sub DrawPicture(pctDraw As PictureBox, rsData As ADODB.Recordset, intStart As Integer, intEnd As Integer, Optional blnEscape0 As Boolean)
'功能:根据传入的数据集进行绘图
'参数:pctDraw - 绘图的控件  rsData-绘图的数据集     intStart -数据集起始列   intEnd - 数据集结束列  blnEscape0-是否跳过0
    Dim i As Integer, j As Integer
    Dim intBaseX As Integer, intBaseY As Integer    '原点坐标
    Dim intXwidth As Integer, intYheight As Integer '坐标轴长度
    Dim intCountX As Integer '横坐标列数
    Dim lngMaxY As Long  '纵坐标最大值
    Dim intLastX As Integer, intLastY As Integer
    Dim intNowX As Integer, intNowY As Integer
    Dim 坐标系Color As Long
    
    Dim dateNow As Date
    
    dateNow = CurrentDate
    pctDraw.Cls
    If rsData.State = 0 Then Exit Sub
    If rsData.RecordCount = 0 Then Exit Sub  '如果没有数据，就退出
    rsData.MoveFirst
    
    '颜色定义
    坐标系Color = &H454545
    
    '绘制坐标系
    intBaseX = 720 '原点
    intBaseY = pctDraw.Height - 1000
    intXwidth = pctDraw.ScaleWidth - 240 - intBaseX 'X、Y轴宽高
    intYheight = intBaseY - 600
    
    With pctDraw
        .DrawWidth = 1
        .FontSize = 9
        
        'X轴
        intCountX = intEnd - intStart + 1
        pctDraw.Line (intBaseX, intBaseY)-(intBaseX + intXwidth, intBaseY - 15), 坐标系Color, B
        .CurrentX = intBaseX: .CurrentY = intBaseY + 45
        pctDraw.Print "0"
        For i = intStart To intEnd
            .CurrentX = intBaseX + intXwidth / (intCountX + 1) * (i - intStart + 1):
            .CurrentY = intBaseY + 45
            pctDraw.Print UCase(rsData.Fields(i).Name)
            .CurrentX = intBaseX + intXwidth / (intCountX + 1) * (i - intStart + 1)
            .CurrentY = intBaseY
            pctDraw.Line (.CurrentX, .CurrentY)-(.CurrentX, .CurrentY - 60), 坐标系Color, B
        Next
        
        'Y轴
        pctDraw.Line (intBaseX, intBaseY)-(intBaseX + 15, intBaseY - intYheight), 坐标系Color, B
        
        '取Y轴最大值确定刻度
        rsData.MoveFirst
        lngMaxY = 1
        Do While Not rsData.EOF
            lngMaxY = IIf(lngMaxY < Val(rsData!MaxValue), Val(rsData!MaxValue), lngMaxY)
            rsData.MoveNext
        Loop
        lngMaxY = (lngMaxY + 5) / 10
        lngMaxY = lngMaxY * 10          '四舍五入取整
        
        For i = 0 To 4
            .CurrentX = intBaseX
            .CurrentY = intBaseY - intYheight / 6 * (i + 1)
            pctDraw.Line (.CurrentX, .CurrentY)-(.CurrentX + 60, .CurrentY), 坐标系Color, B
            .CurrentX = intBaseX - 600
            .CurrentY = intBaseY - intYheight / 6 * (i + 1) - 90
            pctDraw.Print (i + 1) * (lngMaxY / 5)
        Next
        
    End With
    
    rsData.MoveFirst
    With rsData
        pctDraw.FillStyle = 0
        i = 1
        Do While Not rsData.EOF
        '绘制图例
            If i <= 5 Then
                pctDraw.FillColor = RGB(255, 255 - i * 51, 255 - i * 51)
            ElseIf i <= 10 And i > 5 Then
                pctDraw.FillColor = RGB(255 - (i - 5) * 51, 255 - (i - 5) * 51, 255 - (i - 5) * 51)
            Else
                pctDraw.FillColor = RGB(255 - (i - 10) * 51, 255, 255 - (i - 10) * 51)
            End If
            
            pctDraw.DrawWidth = 1.5
            pctDraw.FillStyle = 0
            
            pctDraw.Line (intBaseX + (i - 1) * 1600, intBaseY + 500)-(intBaseX + (i - 1) * 1600 + 500, intBaseY + 700), pctDraw.FillColor, BF
            pctDraw.CurrentX = intBaseX + (i - 1) * 1600 + 550
            pctDraw.CurrentY = intBaseY + 500
            pctDraw.Print .Fields(0).Value
            
            '绘制折线
            pctDraw.DrawMode = 3
            pctDraw.DrawWidth = 1
            For j = intStart To intEnd
                intLastX = intNowX: intNowX = intBaseX + intXwidth / 25 * (j - 2)
                intLastY = intNowY: intNowY = intBaseY - IIf(IsNull(.Fields(j).Value), 0, .Fields(j).Value) / lngMaxY * intYheight * 5 / 6
                pctDraw.Circle (intNowX, intNowY), 25, pctDraw.FillColor
                
                If blnEscape0 Then
                    If j <> intStart Then
                        If CDate(Format(.Fields(0).Value, "yyyy/mm/dd")) < CDate(Format(dateNow, "yyyy/mm/dd")) Then
                            pctDraw.Line (intLastX, intLastY)-(intNowX, intNowY), pctDraw.FillColor
                        ElseIf CDate(Format(.Fields(0).Value, "yyyy/mm/dd")) = CDate(Format(dateNow, "yyyy/mm/dd")) _
                        And j - 3 < Val(Format(dateNow, "hh")) Then
                            pctDraw.Line (intLastX, intLastY)-(intNowX, intNowY), pctDraw.FillColor
                        End If
                    End If
                Else
                    If j <> intStart Then
                        pctDraw.Line (intLastX, intLastY)-(intNowX, intNowY), pctDraw.FillColor
                    End If
                End If
            Next
            
            i = i + 1
            rsData.MoveNext
        Loop
    End With
End Sub

Public Function ChangeSQL(ByVal intMod As Integer, ByVal strOldID As String, ByVal strSqlText As String, ByRef strChildNum As String, ByVal strInstID, Optional strOptVersion As String) As String
    '功能：修改SQL文本，产生新的SQLID用于修改执行计划
    '参数说明：intMod为修改参数, 1-添加RULE提示；2-删除RULE提示；3-添加优化器版本提示 ； 4-删除优化器版本提示 ；5-自定义添加提示字
    '                   strOldID-需要修改的SQLID; strOptVersion 优化器参数版本;strSqlText-自定义的SQL语句
    '返回值:返回1说明SQL语句包含RULE提示，返回2说明没有RULE提示，返回3说明优化器版本提取失败 ，返回4说明没有优化器参数，返回5说明语句执行失败 ,否则返回新的SQL语句
    Dim strNewSQL As String, strTemp As String
    Dim arrPar() As Variant, strSql As String, rsData As ADODB.Recordset
    Dim intHintStart As Integer, intHintEnd As Integer, strHints  As String
    
    On Error GoTo errH
    strNewSQL = strSqlText
    '将语句中的空格去掉，用于匹配查找字符
    strTemp = Replace(UCase(strNewSQL), " ", "")
    
    '添加RULE提示
    If intMod = 1 Then
        If InStr(1, strTemp, "/*+RULE*/") > 0 Then ChangeSQL = "1": Exit Function
        strNewSQL = Left(strNewSQL, InStr(1, strNewSQL, "SELECT")) + Replace(strNewSQL, " ", " /*+ RULE*/ ", InStr(1, strNewSQL, "SELECT") + 1, 1)
    '删除RULE提示
    ElseIf intMod = 2 Then
        If InStr(1, strTemp, "/*+RULE*/") = 0 Then ChangeSQL = "2": Exit Function
        
        Do While InStr(1, strTemp, "/*+RULE*/") > 0
            intHintStart = InStr(1, strNewSQL, "/")
            intHintEnd = InStr(intHintStart + 1, strNewSQL, "/")
            strHints = Mid(strNewSQL, intHintStart, intHintEnd - intHintStart + 1)
            strNewSQL = Replace(strNewSQL, strHints, " ")
            strTemp = Replace(strNewSQL, " ", "")
        Loop
    
    '添加优化器版本提示
    ElseIf intMod = 3 Then
        If strOptVersion = "" Then
            ChangeSQL = "3"
            Exit Function
        End If
        strNewSQL = Replace(strNewSQL, "", "/*+ optimizer_features_enable('" & strOptVersion & "') */", 1, 1)
    ' 删除优化器版本提示
    ElseIf intMod = 4 Then
        If InStr(1, strTemp, "/*+OPTIMIZER_FEATURES_ENABLE") = 0 Then ChangeSQL = "4": Exit Function
        
        Do While InStr(1, strTemp, "/*+OPTIMIZER_FEATURES_ENABLE") > 0
            intHintStart = InStr(1, strNewSQL, "/")
            intHintEnd = InStr(intHintStart + 1, strNewSQL, "/")
            strHints = Mid(strNewSQL, intHintStart, intHintEnd - intHintStart + 1)
            strNewSQL = Replace(strNewSQL, strHints, " ")
            strTemp = Replace(strNewSQL, " ", "")
        Loop
        
    '自定义提示
    Else
        strNewSQL = TrimEx(strSqlText)
    End If
    
    '查看语句是否有绑定变量，有绑定变量则修改后执行
    strSql = "select POSITION,NAME,VALUE_STRING ,last_captured ,DataType from " & IIf(gblnRAC, "G", "") & "v$sql_bind_capture where  SQL_ID= [1]  " & _
                 "and CHILD_NUMBER in (select max(CHILD_NUMBER)  from " & IIf(gblnRAC, "G", "") & "v$sql_bind_capture where SQL_ID= [1]  " & IIf(gblnRAC, "And INST_ID = " & strInstID & " ", "") & "  ) order by POSITION"
    Set rsData = OpenSQLRecord(strSql, "ChangeSQL", strOldID)
    
    '如果查询结果为空，尝试从dba_hist_sqlbind视图中查询
    If rsData.RecordCount = 0 Then
        strSql = "Select  Distinct Position, Name, Value_String, Last_Captured, Datatype From dba_hist_sqlbind  A Where Sql_Id = [1] order by POSITION"
        Set rsData = OpenSQLRecord(strSql, "ChangeSQL", strOldID)
    End If
    
    
    If rsData.RecordCount > 0 Then
        rsData.MoveFirst
        ReDim arrPar(rsData.RecordCount - 1)
    End If
    Do While Not rsData.EOF
        
        '替换绑定变量格式为 ：[1]、[2]
        strNewSQL = Replace(strNewSQL, rsData!Name, "[" & rsData!Position & "]", 1, 1)
        
        '添加参数
        If rsData!DataType = 12 Or rsData!DataType = 180 Then '日期型参数
            arrPar(rsData!Position - 1) = CDate(Format(rsData!value_String, "mm-dd-yyyy hh:mm:ss"))
        ElseIf rsData!DataType = 2 Then  '数字型
            arrPar(rsData!Position - 1) = Int(rsData!value_String)
        Else '字符型
            arrPar(rsData!Position - 1) = "" & rsData!value_String
        End If
        rsData.MoveNext
    Loop
    

    If rsData.RecordCount = 0 Then
        Call OpenSQLRecord(strNewSQL, "ChangeSQL")
    Else
        Call OpenSQLRecordByArray(strNewSQL, "ChangeSQL", arrPar)
    End If
    
    '查找上一条执行的SQLID
    ChangeSQL = GetPrevSQLID(strChildNum)
    If ChangeSQL = "" Then
        ChangeSQL = "5"
        Exit Function
    End If

    Exit Function
errH:
    ChangeSQL = "5"
    If 0 = 1 Then
        Resume
    End If
End Function


Public Function CreateSqlProfiles(strBdSqlID As String, strGdSqlID As String, strChildNum As String) As Boolean
    '功能： 根据传入的SQLID创建SQL PROFILES，成功返回True
    '参数说明：strBdSqlID-需要修改执行计划的SQLID；strGdSqlID-好的执行计划的SQLID
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Declare" & vbNewLine & _
                "  Ar_Hint_Table    Sys.Dbms_Debug_Vc2coll; Ar_Profile_Hints Sys.Sqlprof_Attr := Sys.Sqlprof_Attr(); Cl_Sql_Text      Clob;  I                Pls_Integer;" & vbNewLine & _
                "Begin" & vbNewLine & _
                "  With A As(Select Rownum As r_No, a.* From Table( Dbms_Xplan.Display_Cursor('" & strGdSqlID & "', " & strChildNum & ", 'OUTLINE') ) A)," & vbNewLine & _
                "  B As (Select Min(r_No) As Start_r_No From A Where a.Plan_Table_Output = 'Outline Data')," & vbNewLine & _
                "  C As (Select Min(r_No) As End_r_No From A, B Where a.r_No > b.Start_r_No And a.Plan_Table_Output = '  */')," & vbNewLine & _
                "  D As (Select Instr(a.Plan_Table_Output, 'BEGIN_OUTLINE_DATA') As Start_Col From A, B Where r_No = b.Start_r_No + 4)" & vbNewLine & _
                "  Select Substr(a.Plan_Table_Output, d.Start_Col) As Outline_Hints Bulk Collect" & vbNewLine & _
                "  Into Ar_Hint_Table From A, B, C, D Where a.r_No >= b.Start_r_No + 4 And a.r_No <= c.End_r_No - 1 Order By a.r_No;" & vbNewLine & _
                "  Select Sql_FullText Into Cl_Sql_Text From GV$sql Where Sql_Id = '" & strBdSqlID & "' And Rownum<2;" & vbNewLine & _
                "  I := Ar_Hint_Table.First;" & vbNewLine & _
                "  While I Is Not Null Loop" & vbNewLine & _
                "    If Ar_Hint_Table.Exists(I + 1) Then" & vbNewLine & _
                "      If Substr(Ar_Hint_Table(I + 1), 1, 1) = ' ' Then" & vbNewLine & _
                "        Ar_Hint_Table(I) := Ar_Hint_Table(I) || Trim(Ar_Hint_Table(I + 1)); Ar_Hint_Table.Delete(I + 1);" & vbNewLine & _
                "      End If;" & vbNewLine & _
                "    End If;" & vbNewLine & _
                "    I := Ar_Hint_Table.Next(I);" & vbNewLine & _
                "  End Loop;" & vbNewLine & _
                "  I := Ar_Hint_Table.First;" & vbNewLine & _
                "  While I Is Not Null Loop" & vbNewLine & _
                "    Ar_Profile_Hints.Extend;Ar_Profile_Hints(Ar_Profile_Hints.Count) := Ar_Hint_Table(I); I := Ar_Hint_Table.Next(I);" & vbNewLine & _
                "  End Loop;" & vbNewLine & _
                "  Dbms_Sqltune.Import_Sql_Profile(Sql_Text => Cl_Sql_Text, Profile => Ar_Profile_Hints, Name => 'PROFILE_" & strBdSqlID & "'  , Force_Match => True);" & vbNewLine & _
                "End;"
    gcnOracle.Execute strSql
    CreateSqlProfiles = True
    Exit Function
errH:
    MsgBox "添加SQL PROFILES失败，请联系DBA。" & vbNewLine & Err.Description
    CreateSqlProfiles = False
End Function



Public Sub CheckSqlPlan(vsfPlanTbl As VSFlexGrid, ByVal intOptCol As Integer, ByVal intObjCol As Integer, _
                                            rsBigtbl As ADODB.Recordset, rsBigIdx As ADODB.Recordset, rsLowIdx As ADODB.Recordset)
'功能:检查VSF表格中的执行计划
'         1.大表全表扫描zltables+zlbigtable+zlbaktables，
'         2.中型表全表扫描(如果有统计信息，User_tab_statistics:num_rows>3000(药品目录一般是这个值以上) AND num_rows<100 0000百万以内)
'         3.大表上引用基础表(非大表)的外键上的索引
'         4.大表和中型表索引全扫描（inex full scan，INDEX FAST FULL SCAN）
'         5.大表和中型表跳跃式索引扫描（INDEX SKIP SCAN）
'参数:
'vsfPlanTbl - 执行计划表格
'intOptCol - 操作列,如:Index full scan ,intObjCol - 操作涉及的对象列,如: 病人医嘱记录_IX_ID
'rsBigtbl,rsBigIdx,rsLowIdx -涉及的表/索引
    
    Dim strOperation As String, strObject As String
    Dim strTmp() As String, i As Integer, j As Integer
    Dim blnTmp As Boolean
    
    On Error GoTo errH
    With vsfPlanTbl
        If .Redraw = flexRDNone Then Exit Sub
        
        '遍历表格,获取对象
        For i = .FixedRows To .Rows - .FixedRows
            If intOptCol <> intObjCol Then
                '执行计划的操作和对象不在一列中,直接获取
                strOperation = TrimEx(.TextMatrix(i, intOptCol))
                strObject = TrimEx(.TextMatrix(i, intObjCol))
            Else
                '涉及情况:TABLE ACCESS FULL/INDEX FAST FULL SCAN/INDEX FULL SCAN/INDEX SKIP SCAN/INDEX RANGE SCAN
                strTmp = Split("TABLE ACCESS FULL/INDEX FULL SCAN/INDEX SKIP SCAN/INDEX RANGE SCAN/INDEX FAST FULL SCAN", "/")
                
                For j = 0 To UBound(strTmp)
                    If InStr(1, TrimEx(.TextMatrix(i, intOptCol)), strTmp(j)) > 0 Then
                        strOperation = strTmp(j)
                        strObject = Split(Replace(TrimEx(.TextMatrix(i, intOptCol)), strTmp(j), ""), " ")(0)
                        Exit For
                    End If
                Next
            End If
            
            If strOperation <> "" And strObject <> "" Then
                If strOperation = "TABLE ACCESS FULL" Then '获取全表扫描
                    blnTmp = CheckRs(rsBigtbl, "表名 = '" & strObject & "'") Or gcnOracle = ""
                ElseIf InStr(1, "INDEX FULL SCAN/INDEX SKIP SCAN/INDEX FAST FULL SCAN", strOperation) > 0 Then '索引全扫描\索引跳扫描
                    blnTmp = CheckRs(rsBigIdx, "索引名 = '" & strObject & "'") Or gcnOracle = ""
                ElseIf strOperation = "INDEX RANGE SCAN" And gcnOracle <> "" Then  '索引范围扫描:低效索引
                    blnTmp = CheckRs(rsLowIdx, "约束名= '" & GetFkByIdx(strObject) & "'")
                End If
            End If
                
            If blnTmp Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = FULL_颜色
            
            strOperation = "": strObject = ""
            blnTmp = False
        Next

    End With
    Exit Sub
errH:
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub


Public Sub GetMidTabSize(ByRef lngMinSize As Long, ByRef lngMaxSize As Long)
    '功能:获取中型表大小
    
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    lngMinSize = 3000: lngMaxSize = 1000000
    
    On Error GoTo errH
    strSql = "Select A.参数名,Nvl(A.参数值,A.缺省值) As 参数值 " & _
                 "From zlParameters A " & _
                 "Where A.参数名 = '检查中型表' And a.系统 is null And a.模块 is null"
    Set rsTmp = OpenSQLRecord(strSql, "GetMidTabSize")
    
    If rsTmp.EOF Then Exit Sub
    lngMinSize = Split(rsTmp!参数值, ",")(0)
    lngMaxSize = Split(rsTmp!参数值, ",")(1)
    
    Exit Sub
errH:
    ErrCenter
    If 0 = 1 Then
        Resume
    End If
End Sub

Public Function GetCheckObj(ByVal intMod As Integer, Optional ByVal lngMinSize As Long, Optional ByVal lngMaxSize As Long) As ADODB.Recordset
'功能:获取涉及性能问题的表/索引对象,返回一个记录集
'参数intMod: 1-表,2-索引,3-低效索引
'lngMinSize,lngMaxSize - 判定中型表的区间,不传则默认为3000-1000000

    Dim strSql As String
    
    On Error GoTo errH
    
    If gblnHasZltables Then
        strSql = "Union Select Distinct 表名 From Zltables Where 分类 In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3')"
    Else
        strSql = "Union Select Distinct 表名 From Zlbigtables" & vbNewLine & _
                        "Union" & vbNewLine & _
                        "Select Distinct 表名 From zlBakTables"
    End If
    
    Select Case intMod
        Case 1
            strSql = "Select distinct  Table_Name 表名" & vbNewLine & _
                            "From Dba_Tab_Statistics" & vbNewLine & _
                            "Where Num_Rows Between " & IIf(lngMinSize = 0, 3000, lngMinSize) & " And " & IIf(lngMaxSize = 0, 1000000, lngMaxSize) & vbNewLine & _
                            strSql
                            
        Case 2
            strSql = "Select distinct Index_Name 索引名" & vbNewLine & _
                            "From Dba_Indexes" & vbNewLine & _
                            "Where Table_Name In" & vbNewLine & _
                            " ( Select Table_Name 表名 From Dba_Tab_Statistics Where Num_Rows Between " & IIf(lngMinSize = 0, 3000, lngMinSize) & " And " & IIf(lngMaxSize = 0, 1000000, lngMaxSize) & vbNewLine & _
                            strSql & ")"

        Case 3
            strSql = "Select distinct  a.Constraint_Name 约束名" & vbNewLine & _
                            "From Dba_Constraints A, Dba_Indexes B" & vbNewLine & _
                            "Where a.Constraint_Type = 'R' And b.uniqueness='UNIQUE' And a.r_Constraint_Name = b.Index_Name And a.r_Owner = b.Owner And" & vbNewLine & _
                            "      b.Table_Name Not In" & vbNewLine & _
                            "      (Select Distinct 表名 From Zlbigtables" & vbNewLine & _
                            "       Union Select Distinct 表名 From zlBakTables" & vbNewLine & _
                            IIf(gblnHasZltables, "Union Select Distinct 表名 From Zltables Where 分类 In ('B1', 'B2', 'B3', 'C1', 'C2', 'C3')", "") & vbNewLine & _
                            "       )"

    End Select
    
    Set GetCheckObj = OpenSQLRecord(strSql, "GetCheckObj")
    Exit Function
errH:
    Set GetCheckObj = Nothing
    MsgBox Err.Description
End Function

Public Function CheckRs(rsData As ADODB.Recordset, ByVal strFilter As String) As Boolean
'功能:对传入的记录集添加过滤,如果有匹配项则返回True
    
    If rsData Is Nothing Then Exit Function
    rsData.Filter = strFilter
    CheckRs = Not rsData.EOF
    rsData.Filter = 0
End Function

Public Function GetFkByIdx(ByVal strIdxName As String) As String
'功能:根据传入的索引返回对应的外键约束名称
    
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error GoTo errH:
    
    strSql = "Select Distinct a.Constraint_Name" & vbNewLine & _
                    "From Dba_Cons_Columns A, Dba_Ind_Columns B" & vbNewLine & _
                    "Where a.Owner = b.Table_Owner And a.Table_Name = b.Table_Name And a.Column_Name = b.Column_Name And a.Position = b.Column_Position And" & vbNewLine & _
                    "      b.Index_Name = [1]"

    Set rsData = OpenSQLRecord(strSql, "GetFkByIdx", strIdxName)
        
    If Not rsData.EOF Then
        GetFkByIdx = rsData!Constraint_Name & ""
    End If
    Exit Function
errH:
    GetFkByIdx = ""
    MsgBox Err.Description
    If 0 = 1 Then
        Resume
    End If
End Function
