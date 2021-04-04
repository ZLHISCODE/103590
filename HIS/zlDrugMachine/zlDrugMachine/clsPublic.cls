VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_MSG As String = "clsPublic"
Private mcnOracle As ADODB.Connection
Private mrsParas As ADODB.Recordset
Private mrsDeptParas As ADODB.Recordset
Private mrsUserParas As ADODB.Recordset

Public Sub InitCommon(cnMain As ADODB.Connection)
    Set mcnOracle = cnMain
End Sub

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
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
                strTmp = Replace(TrimEx(Mid(strTmp, 1, InStrRev(strTmp, "SELECT") - 1)), " ", "")
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

    '执行返回记录集
    'If cmdData.ActiveConnection Is Nothing Then
'        If gblnSys = True Then
'            Set cmdData.ActiveConnection = gcnSysConn
'        Else
            Set cmdData.ActiveConnection = mcnOracle '这句比较慢(这句执行1000次约0.5x秒)
'        End If
    'End If
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
        
        '执行过程
        'If cmdData.ActiveConnection Is Nothing Then
            Set cmdData.ActiveConnection = mcnOracle    '这句比较慢
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
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"
    mcnOracle.Execute strSQL, , adCmdText
End Sub

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
'    MsgBox Err.Description, vbInformation, MSTR_MSG
    Currentdate = 0
    Err = 0
End Function

Private Function TrimEx(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
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

Public Function GetPara(ByVal varPara As Variant, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, _
    Optional intType As Integer, Optional ByVal lngDeptID As Long) As String
    
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
    Dim blnNewVersion As Boolean
    
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
        '获取版本号
        strSQL = "Select 版本号 From zlSystems Where 编号 = 100 "
        Set rsTmp = OpenSQLRecord(strSQL, "获取ZLHIS数据库版本")
        If rsTmp.EOF = False Then
            blnNewVersion = NVL(rsTmp!版本号) >= "10.35.0"
        End If
        rsTmp.Close
        
        '参数表，获取参数特征
        If blnNewVersion Then
            '10.35.0以上版本
            strSQL = "Select ID,Nvl(系统,0) as 系统,Nvl(模块,0) as 模块,Nvl(私有,0) as 私有,Nvl(本机,0) as 本机,NVL(部门,0) as 部门,Nvl(授权,0) as 授权,参数号,参数名," & _
                " Nvl(参数值,缺省值) as 参数值,[1] as 用户名,[2] as 机器名 " & vbNewLine & _
                "From zlParameters " & vbNewLine & _
                "Where " & strSqlFilter
        Else
            '10.35.0以下版本
            strSQL = "Select ID,Nvl(系统,0) as 系统,Nvl(模块,0) as 模块,Nvl(私有,0) as 私有,Nvl(本机,0) as 本机,0 as 部门,Nvl(授权,0) as 授权,参数号,参数名," & _
                " Nvl(参数值,缺省值) as 参数值,[1] as 用户名,[2] as 机器名 " & vbNewLine & _
                "From zlParameters " & vbNewLine & _
                "Where " & strSqlFilter
        End If
        Set rsTmp = OpenSQLRecord(strSQL, "GetPara", CStr("ZLHIS"), CStr("P"), lngModual, lngSys, CStr(varPara), Val(varPara))
    
        If rsTmp.EOF Then
            blnNotExists = True
        Else
            If blnNewRow Then
                Call Me.Append(mrsParas, rsTmp)
            Else
                Set mrsParas = CopyNewRec(rsTmp)
            End If
            mrsParas.Filter = strResFilter '再次过滤
            '获取部门级参数
            If mrsParas!部门 = 1 Then
                strSQL = "Select 参数id, 部门ID, 参数值" & vbNewLine & _
                        "From Zldeptparas" & vbNewLine & _
                        "Where 参数id = [1]"
                Set rsTmp = OpenSQLRecord(strSQL, "GetPara", Val(mrsParas!ID), lngDeptID)
                
                If mrsDeptParas Is Nothing Then
                    Set mrsDeptParas = CopyNewRec(rsTmp)
                ElseIf mrsDeptParas.State = 0 Then
                    Set mrsDeptParas = CopyNewRec(rsTmp)
                Else
                    Call Me.Append(mrsDeptParas, rsTmp)
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
                
                Set rsTmp = OpenSQLRecord(strSQL, "GetPara", "", "", Val(mrsParas!ID))
                
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
            mrsDeptParas.Filter = "参数ID=" & mrsParas!ID & _
                 " And 部门ID=" & lngDeptID
            If Not mrsDeptParas.EOF Then
                GetPara = NVL(mrsDeptParas!参数值, strDefault)
            Else
                GetPara = NVL(mrsParas!参数值, strDefault)
            End If
        ElseIf mrsParas!私有 = 1 Or mrsParas!本机 = 1 Then
            mrsUserParas.Filter = "参数ID=" & mrsParas!ID & _
                IIf(mrsParas!私有 = 1, " And 用户名='" & mrsParas!用户名 & "'", " And 用户名='NullUser'") & _
                IIf(mrsParas!本机 = 1, " And 机器名='" & mrsParas!机器名 & "'", " And 机器名='NullMachine'")
            If Not mrsUserParas.EOF Then
                GetPara = NVL(mrsUserParas!参数值, strDefault)
            Else
                GetPara = NVL(mrsParas!参数值, strDefault)
            End If
        Else
            GetPara = NVL(mrsParas!参数值, strDefault)
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
End Function

Private Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional blnOnlyStructure As Boolean, Optional ByVal strFields As String, Optional arrAppFields As Variant) As ADODB.Recordset
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
    
errH:
End Function

Function Append(ByRef rsSource As ADODB.Recordset, ByVal rsAppend As ADODB.Recordset, ParamArray arrInput() As Variant) As Boolean
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
                arrValues(i) = rsAppend(arrAppend(i)).Value
            Next
            rsSource.AddNew arrSource, arrValues
            Erase arrValues
            rsAppend.MoveNext
        Loop
    ElseIf LBound(arrInput) = 0 Then
        Do While Not rsAppend.EOF
            rsSource.AddNew
            For i = 0 To rsSource.Fields.Count - 1
                rsSource.Fields(i).Value = rsAppend.Fields(i).Value
            Next
            rsSource.Update
            rsAppend.MoveNext
        Loop
    End If
    
    Append = True
    Exit Function
    
errH:
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'clsCommFun存在该函数
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
