Attribute VB_Name = "mdlPubLisComlib"
'---------------------------------------------------------------------------------------
'创    建:王振涛
'创建时间:2018/9/27
'模块功能:接口中公共方法。
'---------------------------------------------------------------------------------------

Option Explicit

Public Const Sel_Lis_DB As Integer = 1
Public Const Sel_His_DB As Integer = 2

Public intLis_Setup As Integer                                     '判断LIS是否安装 0=未安装  1=已安装
Public intEMR_Setup As Integer                                     '判断EMR是否安装 0=未安装 1=已安装 现在未使用，新版电子病历中没有没有编号概念，判断安装时通过创建部件和初始化连接是否成功确定是否判断安装



Public Function InitDBConn(cnHisOracle As Connection, Optional strErr As String) As Boolean
      '功能       初使LIS和HIS的数据库连接
      '参数       1=lis 2=his 3=tj 4=xk 5=EMR
      '返回       连接成功返回真,连接不成功返回假

          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim strConn As String
          Dim strCode As String
          Dim astrItem() As String

1         On Error GoTo InitDBConn_Error

2         If cnHisOracle.State <> 1 Then
3             strErr = "传入的HIS连接状态不正常！请检查!"
4             InitDBConn = False
5             Exit Function
6         End If

          '从数据库读取配置的HIS连接
7         strSQL = "Select 参数值 From zlOptions Where  参数名 ='LIS系统连接配置'"
8         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "zlGetSymbol")
9         If rsTmp.RecordCount > 0 Then
10            strConn = rsTmp("参数值") & ""
11        End If
          '没有配置连接表示共享安装
12        If strConn = "" Then
13            Set gcnLisOracle = cnHisOracle
14        Else
              '通过配置的连接来连接HIS库
15            strCode = gobjHisComLib.zlStr.Sm4DecryptEcb(strConn)
16            astrItem = Split(strCode, "<SP 1>")
17            If OraDataOpen(gcnLisOracle, astrItem(2), astrItem(0), IIf(UCase(astrItem(1)) = "SYS" Or UCase(astrItem(1)) = "SYSTEM", _
                                                                         astrItem(1), astrItem(1))) = False Then
18                InitDBConn = False
19                Exit Function
20            End If
21        End If

          '---------------------------判断各个系统是否安装-----------------------------------------

          'LIS
22        strSQL = "select count(*) count from zlsystems where 编号 = 2500 "
23        Set rsTmp = OpenSQLRecord(Sel_Lis_DB, strSQL, "初始化")
24        If rsTmp("count") > 0 Then
25            intLis_Setup = 1
26        End If
          
          '判断LIS系统是否安装，未安装则直接退出
27        If intLis_Setup <> 1 Then Exit Function
          
          '--------------------获取安装版本--------------------------------------
           '获取LIS版本
28        strSQL = "select 版本号  from zlsystems where 编号 = 2500 "
29        Set rsTmp = OpenSQLRecord(Sel_Lis_DB, strSQL, "初始化")
30        If rsTmp.RecordCount > 0 Then
31            gSysInfo.VersionLIS = rsTmp("版本号")
32        End If
          '获取HIS版本
33        strSQL = "select 版本号  from zlsystems where 编号 = 100 "
34        Set rsTmp = OpenSQLRecord(Sel_His_DB, strSQL, "初始化")
35        If rsTmp.RecordCount > 0 Then
36            gSysInfo.VersionHIS = rsTmp("版本号")
37        End If
          
          '新版电子病历
          '未采用以上方式判断安装原因
          '1、独立安装,没有zlsystems表不能通过以上方式判断
          '2、判断安装时直接通过创建EMR部件是否成功,和初始化连接是否成功来判断是否安装。这个不用判断，使用时才判断

          '新版电子病历
38        intEMR_Setup = getERPSetupType
          
          '------------------------------------------------------------------------------------------

39        InitDBConn = True

40        Exit Function
InitDBConn_Error:
41        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(InitDBConn)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
42        Err.Clear
End Function

Private Function getERPSetupType() As Integer
    '功能           判断新版电子病历是否安装
    '由于新版电子病历没有编号,所以只能尝试创建部件,如果部件能够创建,则说明已经安装,否则表示没有安装
    '返回           1=已安装
            
    On Error GoTo getERPSetupType_Error

    If gobjEmrInterface Is Nothing Then
        Set gobjEmrInterface = CreateObject("zl9EmrInterface.ClsEmrInterface")
    End If
    
    getERPSetupType = 1
    
    Exit Function
getERPSetupType_Error:
    getERPSetupType = 0
End Function

Public Function ComGetUserInfo(ByRef strErr As String) As Boolean
      '功能：获取登陆用户信息
      '       intType         1=lis 2=his 3=体验 4=血库
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim objLogin As Object

1         On Error GoTo ComGetUserInfo_Error

2         gstrDBUser = GetUserDB(Sel_His_DB)

3         strSQL = "Select SYS_CONTEXT('USERENV','TERMINAL') as MName From Dual"
4         Set rsTmp = OpenSQLRecord(Sel_His_DB, strSQL, "初始化")
5         gUserInfo.ComputerName = rsTmp("MName")


          '获取登陆站点
6         Set objLogin = CreateObject("ZLLogin.clsLogin")
7         gUserInfo.NodeNo = objLogin.NodeNo
8         If gUserInfo.NodeNo = "" Then gUserInfo.NodeNo = "-"


9         Set rsTmp = GetUserInfo(Sel_His_DB)

10        If Not rsTmp.EOF Then
11            gUserInfo.ID = Val("" & rsTmp!ID)
12            gUserInfo.No = Trim("" & rsTmp!编号)
13            gUserInfo.DeptID = Val("" & rsTmp!部门ID)
14            gUserInfo.DeptName = Trim("" & rsTmp!部门名)
15            gUserInfo.Code = Trim("" & rsTmp!简码)
16            gUserInfo.Name = Trim("" & rsTmp!姓名)
17            ComGetUserInfo = True
18        End If


19        Exit Function
ComGetUserInfo_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(ComGetUserInfo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear

End Function
Public Function ComGetSysParameter(strErr As String) As Boolean
          '读取系统参数

1         On Error GoTo ComGetSysParameter_Error

2         ComGetSysParameter = False
          
3         gSysParameter.BuffDir = App.Path & "\Buffer"
4         gSysParameter.InvaidWord = "`#@$%&|\{}[]?;""'"
5         ComGetSysParameter = True


6         Exit Function
ComGetSysParameter_Error:
7         Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(ComGetSysParameter)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
8         Err.Clear

End Function

Public Function OraDataOpen(cnOracle As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strError As String
    Dim strSysName As String
    
    strSysName = "接口配置"
    
    '兼容性设置,如果有zlRegister则采用新的注册方式,否则采用老的注册方式
    On Error GoTo errhandOld
   
    Set cnOracle = FunGetConnection(strServerName, strUserName, strUserPwd, True, , strError, False)
    If strError <> "" Then
        MsgBox strError, vbInformation, strSysName
        Exit Function
    End If
    OraDataOpen = True
    Exit Function
errhandOld:
    
    On Error Resume Next
    Err = 0
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, TranPasswd(strUserPwd)
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, strSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, strSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, strSysName
            Else
                MsgBox strError, vbInformation, strSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    
    OraDataOpen = True

End Function


Public Function TranPasswd(strOld As String) As String

    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, StrBit As String
    Dim strNew As String
    On Error GoTo TranPasswd_Error

    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        StrBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(StrBit = "0", "W", StrBit = "1", "I", StrBit = "2", "N", StrBit = "3", "T", StrBit = "4", "E", StrBit = "5", "R", StrBit = "6", "P", StrBit = "7", "L", StrBit = "8", "U", StrBit = "9", "M", _
                   StrBit = "A", "H", StrBit = "B", "T", StrBit = "C", "I", StrBit = "D", "O", StrBit = "E", "K", StrBit = "F", "V", StrBit = "G", "A", StrBit = "H", "N", StrBit = "I", "F", StrBit = "J", "J", _
                   StrBit = "K", "B", StrBit = "L", "U", StrBit = "M", "Y", StrBit = "N", "G", StrBit = "O", "P", StrBit = "P", "W", StrBit = "Q", "R", StrBit = "R", "M", StrBit = "S", "E", StrBit = "T", "S", _
                   StrBit = "U", "T", StrBit = "V", "Q", StrBit = "W", "L", StrBit = "X", "Z", StrBit = "Y", "C", StrBit = "Z", "X", True, StrBit)
        Case 2
            strNew = strNew & _
                Switch(StrBit = "0", "7", StrBit = "1", "M", StrBit = "2", "3", StrBit = "3", "A", StrBit = "4", "N", StrBit = "5", "F", StrBit = "6", "O", StrBit = "7", "4", StrBit = "8", "K", StrBit = "9", "Y", _
                   StrBit = "A", "6", StrBit = "B", "J", StrBit = "C", "H", StrBit = "D", "9", StrBit = "E", "G", StrBit = "F", "E", StrBit = "G", "Q", StrBit = "H", "1", StrBit = "I", "T", StrBit = "J", "C", _
                   StrBit = "K", "U", StrBit = "L", "P", StrBit = "M", "B", StrBit = "N", "Z", StrBit = "O", "0", StrBit = "P", "V", StrBit = "Q", "I", StrBit = "R", "W", StrBit = "S", "X", StrBit = "T", "L", _
                   StrBit = "U", "5", StrBit = "V", "R", StrBit = "W", "D", StrBit = "X", "2", StrBit = "Y", "S", StrBit = "Z", "8", True, StrBit)
        Case 0
            strNew = strNew & _
                Switch(StrBit = "0", "6", StrBit = "1", "J", StrBit = "2", "H", StrBit = "3", "9", StrBit = "4", "G", StrBit = "5", "E", StrBit = "6", "Q", StrBit = "7", "1", StrBit = "8", "X", StrBit = "9", "L", _
                   StrBit = "A", "S", StrBit = "B", "8", StrBit = "C", "5", StrBit = "D", "R", StrBit = "E", "7", StrBit = "F", "M", StrBit = "G", "3", StrBit = "H", "A", StrBit = "I", "N", StrBit = "J", "F", _
                   StrBit = "K", "O", StrBit = "L", "4", StrBit = "M", "K", StrBit = "N", "Y", StrBit = "O", "D", StrBit = "P", "2", StrBit = "Q", "T", StrBit = "R", "C", StrBit = "S", "U", StrBit = "T", "P", _
                   StrBit = "U", "B", StrBit = "V", "Z", StrBit = "W", "0", StrBit = "X", "V", StrBit = "Y", "I", StrBit = "Z", "W", True, StrBit)
        End Select
    Next
    TranPasswd = strNew


    Exit Function
TranPasswd_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlPubLisComlib", "执行(TranPasswd)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
    Err.Clear

End Function

Public Sub ExecuteProcedure(ByVal selDB As Integer, strSQL As String, ByVal strFormCaption As String)
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
                    If IsNumeric(strPar) Then    '数字
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarNumeric, adParamInput, 30, strPar)
                    ElseIf Left(strPar, 1) = "'" And Right(strPar, 1) = "'" Then    '字符串
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
                    ElseIf UCase(strPar) Like "TO_DATE('*','*')" Then    '日期
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
                    ElseIf UCase(strPar) = "SYSDATE" Then    '日期
                        If datCur = CDate(0) Then datCur = Currentdate
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adDBTimeStamp, adParamInput, , datCur)
                    ElseIf UCase(strPar) = "NULL" Then    'NULL值当成字符处理可兼容其他类型
                        .Parameters.Append .CreateParameter("PAR" & .Parameters.Count, adVarChar, adParamInput, 200, Null)
                    ElseIf strPar = "" Then    '可选参数当成NULL处理可能改变了缺省值:因此可选参数不能写在中间
                        GoTo NoneVarLine
                    Else    '可能是其他复杂的表达式，无法处理
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

        If selDB = Sel_Lis_DB Then
            Set cmdData.ActiveConnection = gcnLisOracle    '这句比较慢(这句执行1000次约0.5x秒)
        ElseIf selDB = Sel_His_DB Then
            Set cmdData.ActiveConnection = gcnHisOracle    '这句比较慢(这句执行1000次约0.5x秒)
        End If
        
        Call ExportLog(selDB, False, "ExecuteProcedure", strFormCaption, strSQL)
        '执行过程
        'If cmdData.ActiveConnection Is Nothing Then
        '            Set cmdData.ActiveConnection = gcnOracle '这句比较慢
        cmdData.CommandType = adCmdText
        'End If
        cmdData.CommandText = strProc

        '        Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)
        Call cmdData.Execute
        '        Call gobjComLib.SQLTest
        Call ExportLog(selDB, True, "ExecuteProcedure", strFormCaption, "")
    Else
        GoTo NoneVarLine
    End If
    Exit Sub
NoneVarLine:
    '    Call gobjComLib.SQLTest(App.ProductName, strFormCaption, strSQL)

    '说明：为了兼容新连接方式
    '1.新连接用adCmdStoredProc方式在8i下面有问题
    '2.新连接如果不使用{},则即使过程没有参数也要加()
    strSQL = "Call " & strSQL
    If InStr(strSQL, "(") = 0 Then strSQL = strSQL & "()"

    If selDB = Sel_Lis_DB Then
        gcnLisOracle.Execute strSQL, , adCmdText
    ElseIf selDB = Sel_His_DB Then
        gcnHisOracle.Execute strSQL, , adCmdText
    End If


    '    Call gobjComLib.SQLTest

End Sub

'为避免因增加接口函数带来编译部件的兼容性影响，专版上定义为Private函数
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


Public Function OpenSQLRecord(ByVal selDB As Integer, ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
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
    Dim strSQLtmp As String, arrstr As Variant
    Dim strTmp As String, strSQLtmp1 As String
    
    '检查如果使用了动态内存表，并且没有使用/*+ XXX*/等提示字时自动加上

    strSQLtmp = Trim(UCase(strSQL))
    If Mid(Trim(Mid(strSQLtmp, 7)), 1, 2) <> "/*" And Mid(strSQLtmp, 1, 6) = "SELECT" Then
        arrstr = Split("F_STR2LIST,F_NUM2LIST,F_NUM2LIST2,F_STR2LIST2", ",")
        For i = 0 To UBound(arrstr)
            strSQLtmp1 = strSQLtmp
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
    Call ExportLog(selDB, False, "OpenSQLRecord", strTitle, strSQL, arrInput)
    '执行返回记录集
    'If cmdData.ActiveConnection Is Nothing Then
    If selDB = Sel_Lis_DB Then
        Set cmdData.ActiveConnection = gcnLisOracle '这句比较慢(这句执行1000次约0.5x秒)
    ElseIf selDB = Sel_His_DB Then
        Set cmdData.ActiveConnection = gcnHisOracle '这句比较慢(这句执行1000次约0.5x秒)
    End If
    
'     Set cmdData.ActiveConnection = gcnOracle '这句比较慢(这句执行1000次约0.5x秒)
    'End If
    cmdData.CommandText = strSQL
    
'    Call gobjComLib.SQLTest(App.ProductName, strTitle, strLog)
    Set OpenSQLRecord = cmdData.Execute
    Set OpenSQLRecord.ActiveConnection = Nothing
    Call ExportLog(selDB, True, "OpenSQLRecord", strTitle, "")
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
        .Open "SELECT SYSDATE FROM DUAL", gcnLisOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).value
    rsTemp.Close
    Exit Function
errH:
'    If gobjComLib.ErrCenter() = 1 Then Resume
    Currentdate = 0
    Err = 0
End Function

Public Function SetPara(ByVal selDB As Integer, ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
      '功能：设置指定的参数值
      '参数：varPara=参数号或参数名，以数字或字符类型传入区分
      '      strValue=要设置的参数值
      '      lngSys=使用该参数的系统编号，如100
      '      lngModual=使用该参数的模块号，如1230
      '      blnSetup=调用模块是否有参数设置权限
      '返回：设置是否成功
          Dim strSQL As String
          Dim strResFilter As String
          '检查参数值，如果没有变化则不处理
1         On Error GoTo SetPara_Error

2         strSQL = GetPara(selDB, varPara, lngSys, lngModual)
3         If strSQL = strValue Then SetPara = True: Exit Function
          
4         SetPara = True
5         strSQL = "zl_Parameters_Update('" & varPara & "','" & strValue & "'," & lngSys & "," & lngModual & "," & IIf(blnSetup, 1, 0) & ")"
6         Call ExecuteProcedure(selDB, strSQL, "SetPara")
          
          '更新缓存记录集，逻辑与zl_Parameters_Update保持一致
          '过滤条件
7         If TypeName(varPara) = "String" Then
8             strResFilter = "参数名='" & CStr(varPara) & "' And 模块=" & lngModual & " And 系统=" & lngSys
9         Else
10            strResFilter = "参数号=" & Val(varPara) & " And 模块=" & lngModual & " And 系统=" & lngSys
11        End If
          
12        grsParas.Filter = strResFilter
13        If grsParas.EOF Then Exit Function
          '权限判断
14        If Not blnSetup Then
              '公共全局参数,固定需要权限
15            If grsParas!系统 <> 0 And grsParas!模块 = 0 And grsParas!私有 = 0 And grsParas!本机 = 0 Then
16                Exit Function
              '公共模块参数,固定需要权限
17            ElseIf grsParas!模块 = 0 And grsParas!私有 = 0 And grsParas!本机 = 0 Then
18                Exit Function
              '要授权控制的本机公共模块
19            ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 0 And grsParas!本机 = 1 And grsParas!授权 = 1 Then
20                Exit Function
21            End If
22        End If
          
23        If grsParas!私有 = 1 Or grsParas!本机 = 1 Then
24            grsUserParas.Filter = "参数ID=" & grsParas!ID & _
                          IIf(grsParas!私有 = 1, " And 用户名='" & grsParas!用户名 & "'", " And 用户名='NullUser'") & _
                          IIf(grsParas!本机 = 1, " And 机器名='" & grsParas!机器名 & "'", " And 机器名='NullMachine'")
              
25            If grsUserParas.EOF Then
26                grsUserParas.AddNew
27                grsUserParas!参数id = grsParas!ID
28                grsUserParas!用户名 = IIf(grsParas!私有 = 1, grsParas!用户名, "NullUser")
29                grsUserParas!机器名 = IIf(grsParas!本机 = 1, grsParas!机器名, "NullMachine")
30                grsUserParas!参数值 = strValue
31                grsUserParas.Update
32            Else
33                grsUserParas!参数值 = strValue
34                grsUserParas.Update
35            End If
36        Else
37            grsParas!参数值 = strValue
38            grsParas.Update
39        End If
          
40        Exit Function
SetPara_Error:
41        SetPara = False
42        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(SetPara)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
43        Err.Clear
End Function


Public Function GetPara(ByVal selDB As Integer, ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
      '功能：读取指定的参数值
      '参数：varPara=参数号或参数名，以数字或字符类型传入区分
      '      lngSys=使用该参数的系统编号，如100
      '      lngModual=使用该参数的模块号，如1230
      '      strDefault=当数据库中没有该参数时使用的缺省值(注意不是为空时)
      '      blnNotCache=是否不从缓存中读取
      '      arrControl=控件数组，如Array(Me.Text1, Me.CheckBox1)，用于函数内部自动处理对应控件的显示颜色，是否禁止设置。
      '      blnSetup=调用模块是否有参数设置权限
      '      intType=返回参数，返回参数类型
      '返回：参数值，字符串形式
          Dim strSQL As String, i As Integer
          Dim blnNew As Boolean, blnEnabled As Boolean, blnNewRow As Boolean, blnNotExists As Boolean
          Dim strSqlFilter As String, strResFilter As String
          Dim rsTmp As ADODB.Recordset
          Dim strDBUser As String

1         On Error GoTo GetPara_Error

2         strDBUser = GetUserDB(selDB)
3         intType = 0
          
          '过滤条件
4         If TypeName(varPara) = "String" Then
5             strResFilter = "参数名='" & CStr(varPara) & "' And 模块=" & lngModual & " And 系统=" & lngSys
6             strSqlFilter = "参数名=[5] And Nvl(模块,0)=[3] And Nvl(系统,0)= [4] "
7         Else
8             strResFilter = "参数号=" & Val(varPara) & " And 模块=" & lngModual & " And 系统=" & lngSys
9             strSqlFilter = "参数号=[6] And Nvl(模块,0)=[3] And Nvl(系统,0)=[4] "
10        End If
          
          '参数缓存判断
11        If grsParas Is Nothing Then
12            blnNew = True
13        ElseIf grsParas.State = 0 Then
14            blnNew = True
15        Else
16            grsParas.Filter = strResFilter
17            blnNewRow = grsParas.EOF
18        End If
          
19        If blnNew Or blnNewRow Then
              '参数表，获取参数特征
20            strSQL = "Select ID,Nvl(系统,0) as 系统,Nvl(模块,0) as 模块,Nvl(私有,0) as 私有,Nvl(本机,0) as 本机,Nvl(授权,0) as 授权,参数号,参数名," & _
                  " Nvl(参数值,缺省值) as 参数值,[1] as 用户名,[2] as 机器名 From zlParameters Where " & strSqlFilter
21            Set rsTmp = OpenSQLRecord(selDB, strSQL, "GetPara", strDBUser, gUserInfo.ComputerName, lngModual, lngSys, CStr(varPara), Val(varPara))
          
22            If rsTmp.EOF Then
23                blnNotExists = True
24            Else
25                If blnNewRow Then
26                    grsParas.AddNew
27                    For i = 0 To rsTmp.Fields.Count - 1
28                        grsParas.Fields(i) = rsTmp.Fields(i).value
29                    Next
30                    grsParas.Update
31                Else
32                    Set grsParas = New ADODB.Recordset
33                    Set grsParas = CopyNewRec(rsTmp)
34                End If
                  '获取用户或本机参数
35                If grsParas!私有 = 1 Or grsParas!本机 = 1 Then
36                    strSQL = "Select 参数id, Nvl(用户名, 'NullUser') As 用户名, Nvl(机器名, 'NullMachine') As 机器名, 参数值" & vbNewLine & _
                              "From zlUserParas" & vbNewLine & _
                              "Where 参数id = [3]"
                              
37                    If grsParas!私有 = 1 And grsParas!本机 = 1 Then
38                        strSQL = strSQL & " And 用户名=[1] And 机器名=[2]"
39                    ElseIf grsParas!私有 = 1 Then
40                        strSQL = strSQL & " And 用户名=[1] "
41                    Else
42                        strSQL = strSQL & " And 机器名=[2]"
43                    End If
                      
44                    Set rsTmp = OpenSQLRecord(selDB, strSQL, "GetPara", strDBUser, gUserInfo.ComputerName, Val(rsTmp!ID))
                      
45                    If grsUserParas Is Nothing Then
46                        Set grsUserParas = New ADODB.Recordset
47                        Set grsUserParas = CopyNewRec(rsTmp)
48                    ElseIf grsUserParas.State = 0 Then
49                        Set grsUserParas = New ADODB.Recordset
50                        Set grsUserParas = CopyNewRec(rsTmp)
51                    End If
                      
52                    Do While Not rsTmp.EOF
53                        grsUserParas.AddNew
54                        For i = 0 To rsTmp.Fields.Count - 1
55                            grsUserParas.Fields(i) = rsTmp.Fields(i).value
56                        Next
57                        grsUserParas.Update
58                        rsTmp.MoveNext
59                    Loop
60                End If
61            End If
62        End If

63        If blnNotExists Then
64            GetPara = strDefault
65        Else
              '获取参数值
66            If grsParas!私有 = 1 Or grsParas!本机 = 1 Then
67                grsUserParas.Filter = "参数ID=" & grsParas!ID & _
                      IIf(grsParas!私有 = 1, " And 用户名='" & grsParas!用户名 & "'", " And 用户名='NullUser'") & _
                      IIf(grsParas!本机 = 1, " And 机器名='" & grsParas!机器名 & "'", " And 机器名='NullMachine'")
68                If Not grsUserParas.EOF Then
69                    GetPara = NVL(grsUserParas!参数值, strDefault)
70                Else
71                    GetPara = NVL(grsParas!参数值, strDefault)
72                End If
73            Else
74                GetPara = NVL(grsParas!参数值, strDefault)
75            End If
              
              '返回参数类型：1.公共全局,2.私有全局,3.公共模块,4.私有模块,5.本机公共模块(不授权控制),6.本机私有模块,15.本机公共模块(要授权控制)
76            If grsParas!系统 <> 0 And grsParas!模块 = 0 And grsParas!私有 = 0 And grsParas!本机 = 0 Then
77                intType = 1
78            ElseIf grsParas!模块 = 0 And grsParas!私有 = 1 And grsParas!本机 = 0 Then
79                intType = 2
80            ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 0 And grsParas!本机 = 0 Then
81                intType = 3
82            ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 1 And grsParas!本机 = 0 Then
83                intType = 4
84            ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 0 And grsParas!本机 = 1 Then
85                intType = IIf(grsParas!授权 = 1, 15, 5)
86            ElseIf grsParas!系统 <> 0 And grsParas!模块 <> 0 And grsParas!私有 = 1 And grsParas!本机 = 1 Then
87                intType = 6
88            End If
              
              '处理对应的控件颜色，可控状态
89            If IsArray(arrControl) And (intType = 3 Or (intType Mod 10) = 5) Then
90                blnEnabled = Not ((intType = 3 Or (intType Mod 10) = 5 And grsParas!授权 = 1) And Not blnSetup)
91                For i = 0 To UBound(arrControl)
92                    Select Case TypeName(arrControl(i))
                      Case "Label"
93                        arrControl(i).ForeColor = vbBlue
94                    Case "TextBox", "MaskEdBox", "CheckBox", "OptionButton", "ComboBox", "ListBox", "Frame", "PictureBox", "ListView"
95                        arrControl(i).ForeColor = vbBlue
96                        If Not blnEnabled Then arrControl(i).Enabled = False
97                    Case "CommandButton", "DTPicker"
98                        If Not blnEnabled Then arrControl(i).Enabled = False
99                    Case "MSHFlexGrid"
100                       arrControl(i).ForeColor = vbBlue
101                       arrControl(i).ForeColorFixed = vbBlue
102                       If Not blnEnabled Then arrControl(i).Enabled = False
103                   Case "VSFlexGrid"
104                       arrControl(i).ForeColor = vbBlue
105                       arrControl(i).ForeColorFixed = vbBlue
106                       If Not blnEnabled Then arrControl(i).Editable = 0
107                   Case Else
108                       On Error Resume Next
109                       arrControl(i).ForeColor = vbBlue
110                       If Not blnEnabled Then arrControl(i).Enabled = False
111                       Err.Clear
112                   End Select
113               Next
114           End If
115       End If
          

116       Exit Function
GetPara_Error:
117       Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(GetPara)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
118       Err.Clear
End Function

Public Function GetPrivFunc(ByVal selDB As Integer, lngSys As Long, lngProgId As Long) As String
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
        
Beging:
        Set rsTmp = OpenSQLRecord(selDB, strSQL, "GetPrivFunc", lngSys, lngProgId)
        On Error GoTo errH
        
        Do While Not rsTmp.EOF
            strPrivs = strPrivs & ";" & rsTmp!功能
            rsTmp.MoveNext
        Loop
        strPrivs = Mid(strPrivs, 2)
        gcolPrivs.Add strPrivs, "_" & lngSys & "_" & lngProgId
    End If
    On Error GoTo 0
    
    GetPrivFunc = strPrivs
    Exit Function
errH:
    If Not blnRegCheck Then
        '如果出错,可能是由于没有调用zlRegCheck造成,自动调用一次,如果再出错,才提示.
        If selDB = 1 Then
            If initRegister = True Then
               zlRegister.zlRegInit gcnLisOracle
            End If
            If FunzlRegCheck(, gcnLisOracle) <> "" Then Exit Function
        Else
'            If initRegister = True Then
'               zlRegister.zlRegInit gcnHisOracle
'            End If
            If FunzlRegCheck(, gcnHisOracle) <> "" Then Exit Function
        End If
'        GetPrivFunc = zlRegister.zlRegFunc(lngSys, lngProgId)
        blnRegCheck = True
        GoTo Beging
    End If
End Function

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset) As ADODB.Recordset
      '编制人:朱玉宝
      '编制日期:2000-11-02
      '复制记录集
      '在程序中，经常会涉及到相互传递记录集，而使用ADO的Clone复制产生的记录集，当其中一个记录集的数据发生变化的时候，所有副本都将发生相同的变化（通常指修改或删除），而我们往往希望这些记录集相互间保持独立
          Dim rsClone As New ADODB.Recordset
          Dim rsTarget As New ADODB.Recordset
          Dim intFields As Integer
          
1         On Error GoTo CopyNewRec_Error

2         Set rsClone = rsSource.Clone
3         rsClone.Filter = rsSource.Filter
4         Set rsTarget = New ADODB.Recordset
5         With rsTarget
6             For intFields = 0 To rsClone.Fields.Count - 1
7                 .Fields.Append rsClone.Fields(intFields).Name, IIf(rsClone.Fields(intFields).Type = adNumeric, adDouble, rsClone.Fields(intFields).Type), rsClone.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
8             Next
              
9             .CursorLocation = adUseClient
10            .CursorType = adOpenStatic
11            .LockType = adLockOptimistic
12            .Open
              
13            If rsClone.RecordCount <> 0 Then rsClone.MoveFirst
14            Do While Not rsClone.EOF
15                .AddNew
16                For intFields = 0 To rsClone.Fields.Count - 1
17                    .Fields(intFields) = rsClone.Fields(intFields).value
18                Next
19                .Update
20                rsClone.MoveNext
21            Loop
22        End With
          
23        Set CopyNewRec = rsTarget


24        Exit Function
CopyNewRec_Error:
25        Call WriteErrLog("zlPublicHisCommLis", "mdlLisHisComm", "执行(CopyNewRec)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
26        Err.Clear
End Function

Public Function ReplaseSpecial(strTmp As String) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能               替换特殊字符
    '参数
    '                   需替换的字符
    '返回               需替换了特殊字符后的字串
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intloop As Integer
    Dim strSpecial As String
    Dim astrtmp() As String
    strSpecial = "'^‘^’^;^；^:^：^?^？^|^,^，^.^。^"""
    astrtmp = Split(strSpecial, "^")
    For intloop = 0 To UBound(astrtmp)
        strTmp = Replace$(strTmp, astrtmp(intloop), "")
    Next
    ReplaseSpecial = strTmp
    
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "Select zlWBcode([1]) From Dual"
    Else
        strSQL = "Select zlSpellcode([1]) From Dual"
    End If
    On Error GoTo Errhand
    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "zlGetSymbol", strInput)
    zlGetSymbol = IIf(IsNull(rsTmp.Fields(0).value), "", rsTmp.Fields(0).value)
    Exit Function
Errhand:
'    If gobjComLib.ErrCenter() = 1 Then Resume
'    Call gobjComLib.SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/9/27
'功    能:提供其他地方调用参数设置，保存，执行过程，查询语句等
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function ComSetPara(ByVal selDB As Integer, ByVal varPara As Variant, ByVal strValue As String, Optional ByVal lngSys As Long, _
    Optional ByVal lngModual As Long, Optional ByVal blnSetup As Boolean = True) As Boolean
    '设置参数
    ComSetPara = SetPara(selDB, varPara, strValue, lngSys, lngModual, blnSetup)
End Function

Public Function ComGetPara(ByVal selDB As Integer, ByVal varPara As Variant, Optional ByVal lngSys As Long, Optional ByVal lngModual As Long, Optional ByVal strDefault As String, _
    Optional ByVal arrControl As Variant, Optional ByVal blnSetup As Boolean, Optional intType As Integer) As String
    '取参数
    
    ComGetPara = GetPara(selDB, varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType)
    
End Function

Public Function ComGetPrivs(ByVal selDB As Integer, ByVal lngSys As Long, ByVal lngModul As Long) As String
    '读取模块权限
   ComGetPrivs = GetPrivFunc(selDB, lngSys, lngModul)
End Function

Public Function StringFormatDate(strDate, Optional MinOrMax As Integer) As String
    '功能               格式化保存到数据库的日期格式
    '参数               strDate 传入日期格式
    '                   MinOrMax 1=最小 2=最大
    '返回               格式化好的日期格式
    Select Case MinOrMax
        Case 0
            StringFormatDate = "TO_DATE('" & Format(strDate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
        Case 1
            StringFormatDate = "TO_DATE('" & Format(strDate, "yyyy-MM-dd 00:00:00") & "','yyyy-mm-dd hh24:mi:ss')"
        Case 2
            StringFormatDate = "TO_DATE('" & Format(strDate, "yyyy-MM-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')"
    End Select
End Function

Public Function GetUserInfo(ByVal intSelDB As Integer) As ADODB.Recordset
      '功能：获取当前用户的基本信息
      '返回：返回Ado记录集
          Dim strSQL As String
          Dim strDefault As String
          Dim strDBUser As String

1         On Error GoTo GetUserInfo_Error
2         strDBUser = GetUserDB(intSelDB)
3         strDefault = " And C.缺省 = 1"
4         strSQL = "Select User,A.Id, A.编号, A.简码, A.姓名, B.用户名, C.部门id, D.编码 As 部门码, D.名称 As 部门名" & vbNewLine & _
                   "From 人员表 A, 上机人员表 B, 部门人员 C, 部门表 D" & vbNewLine & _
                   "Where A.Id = B.人员id And A.Id = C.人员id And C.部门id = D.Id And B.用户名 = [1]"

5         Set GetUserInfo = OpenSQLRecord(intSelDB, strSQL & strDefault, "GetUserInfo", strDBUser)
6         If GetUserInfo.RecordCount = 0 Then
7             strDefault = " And Rownum < 2"
8             Set GetUserInfo = OpenSQLRecord(intSelDB, strSQL & strDefault, "GetUserInfo", strDBUser)
9         End If

10        Exit Function


11        Exit Function
GetUserInfo_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "mdlPubLisComlib", "执行(GetUserInfo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear

End Function

Public Function GetUserDB(ByVal selDB As Integer) As String
          Dim strTmp As String
          Dim strConnStr As String

1         On Error GoTo GetUserDB_Error

2         If selDB = Sel_Lis_DB Then
3             strConnStr = gcnLisOracle.ConnectionString
4         ElseIf selDB = Sel_His_DB Then
5             strConnStr = gcnHisOracle.ConnectionString
6         End If
7         strTmp = Mid(strConnStr, InStr(strConnStr, "User ID="))
8         strTmp = Mid(strTmp, 9, InStr(strTmp, ";") - 9)
9         GetUserDB = UCase(strTmp)


10        Exit Function
GetUserDB_Error:
11        Call WriteErrLog("zlPublicHisCommLis", "mdlPubLisComlib", "执行(GetUserDB)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
12        Err.Clear
End Function

'--------------------------------------------------
'功能：验证系统注册授权的正确性
'参数：blnTemp-是否从未保存的临时注册信息验证
'返回：正确返回"";错误返回错误信息
'--------------------------------------------------
Public Function HiszlRegCheck(ByVal selDB As Integer, Optional blnTemp As Boolean) As String
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim arrMd5_1(5) As String
    Dim arrMd5_2(5) As String
    Dim arrMd5_3(5) As String
    Dim arrMd5_4(5) As String
    Dim arrMd5_5(5) As String
    Dim strMD5 As String
    Dim intLine As Integer
    On Error GoTo Errhand
    
    '---------------------------------Beging 验证 F_Reg_Audit 是否被替换
    '-- 管理工具 9.25 HIS 10.15
    arrMd5_1(0) = "6746B20191FD2AA9B0E08AFB44E80D4B"
    arrMd5_1(1) = "93C94497A547C10EC3B5C95F5188BA5D"
    arrMd5_1(2) = "A5596EA1AB4F6D4939CBD9599CBFBA0F"
    arrMd5_1(3) = "07069FF5FF76C204EEFCC88366F6A495"
    arrMd5_1(4) = "73C7DB3F742EBC654FAC289B4D37A7B0"
        
    '-- 管理工具 9.35 HIS 10.24
    arrMd5_2(0) = "10E1A9794EF861981C7F53D887990B1F"
    arrMd5_2(1) = "C4A92BE1F6882A57564206E9B391A600"
    arrMd5_2(2) = "F4878F9061BFC4357DC4545EAC326CD2"
    arrMd5_2(3) = "4BBF3E2A0D667A50B8CBC443A1110EA2"
    arrMd5_2(4) = "07BC27215593F6ED86C9905C0D215BD9"
        
    '-- 管理工具 9.37 HIS 10.26
    arrMd5_3(0) = "4D1B31CCB39BDCCE4EE61357555DAD9D"
    arrMd5_3(1) = "F544A3A12A833F6EE10CEA514D65782C"
    arrMd5_3(2) = "5CEF0276B15026C1D5546A85F9A3BE1F"
    arrMd5_3(3) = "487CC8AD6D5F2E0DC337677D02EA702F"
    arrMd5_3(4) = "20AD16738F21A228D962E59DAECB0D84"
    
    '-- 管理工具 9.41 HIS 10.30
    arrMd5_4(0) = "01322819F7B38E12BCAA8525895EF288"
    arrMd5_4(1) = "75E62456DB5F6742B9140DFB73D094FE"
    arrMd5_4(2) = "4270A613EA65B66BF4200BA42F205319"
    arrMd5_4(3) = "64FD2D54E72F9F647DD01D14116988AE"
    arrMd5_4(4) = "D7A22AF77FAC34E04086B800570BCB37"
        
    '-- 管理工具 9.45 HIS 10.34
    arrMd5_5(0) = "01322819F7B38E12BCAA8525895EF288"
    arrMd5_5(1) = "02AC74A017BEE67D26051B4BA5DA98E8"
    arrMd5_5(2) = "9D1143BA317F835426BB8ED2F319A8CA"
    arrMd5_5(3) = "E2718B7863EB402205FAC8CDD348D649"
    arrMd5_5(4) = "39A9E549EAB1EDD396230AD61DC559B0"
    '数据启动，第一次执行，RowNum不是和Line排序对应的，第二次执行以后均正常，因此增加子查询
    strSQL = "Select 源码, Rownum As Line" & vbNewLine & _
            "From (Select Substr(Text, 1, 512) As 源码" & vbNewLine & _
            "       From All_Source" & vbNewLine & _
            "       Where Owner = 'ZLTOOLS' And Name = 'F_REG_AUDIT' And Line In (3, 5, 7, 9, 11)" & vbNewLine & _
            "       Order By Line)"

    Set rsTemp = OpenSQLRecord(selDB, strSQL, "zlRegCheck")
    Do Until rsTemp.EOF
        strMD5 = Md5_String_Calc("" & rsTemp!源码)
        intLine = Val("" & rsTemp!Line)
        If Not (arrMd5_1(intLine - 1) = strMD5 Or arrMd5_2(intLine - 1) = strMD5 _
            Or arrMd5_3(intLine - 1) = strMD5 Or arrMd5_4(intLine - 1) = strMD5 _
            Or arrMd5_5(intLine - 1) = strMD5) Then
            HiszlRegCheck = "注册验证程序不正确，请使用正确的注册程序！"
            Exit Do
        End If
        rsTemp.MoveNext
    Loop
    If HiszlRegCheck <> "" Then Exit Function
    '---------------------------------          End  验证 F_Reg_Audit 是否被替换
    
    strSQL = "Select zltools.f_Reg_Audit([1]) As Stamp From zltools.zlRegInfo r Where 项目='授权证章'"
    Set rsTemp = OpenSQLRecord(selDB, strSQL, "zlRegCheck", IIf(blnTemp, 1, 0))
    If rsTemp.RecordCount > 0 Then
        If Left(rsTemp.Fields(0).value, 6) <> "ERROR-" Then
            HiszlRegCheck = ""
        Else
            HiszlRegCheck = rsTemp.Fields(0).value
        End If
    Else
        HiszlRegCheck = "注册信息丢失,在重新注册前"
    End If
    Exit Function
Errhand:
    HiszlRegCheck = Err.Description
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '功能：求取指定字符串的实际长度，用于判断实际包含双字节字符串的
    '       实际数据存储长度
    '参数：
    '       strAsk
    '返回：
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/10/11
'功    能:根据特定规则产生新的号码,本函数规则只适于ZLHIS10，且需要Oracle 8i(8.1.5)以上版本支持
'支持传入不同连接后，生成不同数据库的连接，包含HIS,LIS
'参数：
'int序号=项目序号:
'  1   病人ID 数字
'  2   住院号 数字
'  3   门诊号 数字
'  10  医嘱发送号 数字,顺序递增编号
'  x   其它单据号 字符,根据编号规则顺序递增编号,不自动补缺
'lng科室ID=按科室号码编号规则的项目需要
'返回：最大号码
'说明：
'  编号规则：0-按年顺序编号,1-按日顺序编号,2-按执行科室分月编号(需要读取科室号码表)
'            对门诊号：0-顺序编号,1-年月日(YYMMDD)+顺序号(0000)
'            对住院号：0-顺序编号,1-年月(YYMM)+顺序号(0000),2-年(YYYY)+顺序号(00000)
'  年度位确定：以1990为基数，随年度增长，按“0～9/A～Z”顺序作为年度编码
'  最大号码-10存入号码控制表,用于并发情况下补缺号(取了号,但未使用)
'  For Update在并发情况下锁定行,不用Wait选项以避免向调用者返回空
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Public Function GetNextNo(ByVal selDB As Integer, ByVal int序号 As Integer, Optional ByVal lng科室ID As Long, Optional ByVal strTag As String, Optional ByVal intStep As Integer = 1) As Variant

    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
        
    On Error GoTo GetNextNo_Error

    GetNextNo = Null
    
    
    strSQL = "Select NextNO([1],[2],[3],[4]) as NO From Dual"
    Set rsTmp = OpenSQLRecord(selDB, strSQL, "GetNextNo", int序号, lng科室ID, strTag, intStep)
    
'    If gcnOracle.Errors.Count > 0 Then 'Select中函数出错时,在VB中不自动触发错误
'        Err.Raise gcnOracle.Errors(0).Number, , gcnOracle.Errors(0).Description
'    End If
    
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp!No) Then GetNextNo = rsTmp!No
    End If

    Exit Function
GetNextNo_Error:
    Call WriteErrLog("zlPublicHisCommLis", "mdlPubLisComlib", "执行(GetNextNo)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, False)
    Err.Clear

End Function


'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/10/11
'功能：读取指定表名对应的序列(按规范，其序列名称为“表名称_id”)的下一数值
'参数：
'   strTable：表名称
'返回：
'调整影响:
'---------------------------------------------------------------------------------------
Public Function GetNextId(strTable As String) As Long
    Dim rsTmp           As New ADODB.Recordset
    Dim strSQL As String, strtab As String

    '不能用错误错处理,原因是序列失效和没有序列时,应该返回错误,不然返回零,就有问题!
    'On Error GoTo errH
    strtab = Trim(strTable)
    If strtab = "门诊费用记录" Or strtab = "住院费用记录" Then strtab = "病人费用记录"

    strSQL = "Select " & strtab & "_ID.Nextval From Dual"
    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "生成ID序列")
    If Not rsTmp.EOF Then
        GetNextId = rsTmp.Fields(0).value
    End If
    '    Exit Function
    'errH:
    '    If gobjComLib.ErrCenter() = 1 Then Resume
End Function


'---------------------------------------------------------------------------------------
'编    码:王振涛
'编码时间:2018/10/15
'功    能：比较两个版本号,比当前版本号小，返回1，相等返回0，比当前版本号大返回-1
'参    数：strVerCur=当前版本号
'         strVerCom=对比的版本号
'返    回：对比版本号比当前版本号小，返回1，相等返回0，比当前版本号大返回-1
'调整影响:
'---------------------------------------------------------------------------------------
Public Function VerCompare(ByVal strVerCur As String, Optional ByVal strVerCom As String) As Integer

    If VerFull(strVerCur) < VerFull(strVerCom) Then
        VerCompare = -1
    ElseIf VerFull(strVerCur) > VerFull(strVerCom) Then
        VerCompare = 1
    Else
        VerCompare = 0
    End If
End Function

Public Function VerFull(ByVal strVer As String, Optional ByVal blnMax As Boolean) As String
'功能：返回VB最大支持的版本号形式:9999.9999.9999.9999,最小版本号0000.0000.0000.0000
'参数：strVer=当前版本号
'           blnMax=True,若果为空，则返回最大支持版本，False=若果为空，则返回最小支持版本
    Dim arrVer As Variant
    If Not IsVerSion(strVer) Then
        VerFull = IIf(blnMax, "9999.9999.9999.9999", "0000.0000.0000.0000")
        Exit Function
    End If
    '增加一段，以兼容特殊SP版本号
    arrVer = Split(strVer & ".0", ".")
    VerFull = Format(arrVer(0), "0000") & "." & Format(arrVer(1), "0000") & "." & Format(arrVer(2), "0000") & "." & Format(arrVer(3), "0000")
End Function

Public Function IsVerSion(ByVal strVer As String) As Boolean
'功能：判断字符串是否是版本号
    Dim arrVer As Variant
    Dim i As Integer
    If Not strVer Like "*.*.*" Then Exit Function
    arrVer = Split(strVer, ".")
    If UBound(arrVer) < 2 Or UBound(arrVer) > 3 Then Exit Function
    
    For i = LBound(arrVer) To UBound(arrVer)
        If Not IsNumeric(arrVer(i)) Then Exit Function
        If Val(arrVer(i)) < 0 Or Val(arrVer(i)) > 9999 Then Exit Function
        If i = 3 Then
            If Format(Val(arrVer(i)), "0000") <> Format(Trim(arrVer(i)), "0000") Then Exit Function
        Else
            If Val(arrVer(i)) & "" <> Trim(arrVer(i)) Then Exit Function
        End If
    Next
    
    IsVerSion = True
End Function


'---------三方报告使用
Public Function ReadLob(ByVal lngSys As Long, ByVal Action As Long, ByVal KeyWord As String, _
                        Optional ByVal strFile As String, Optional ByVal bytFunc As Byte = 0, _
                        Optional bytMoved As Byte = 0) As String
'功能：将指定的LOB字段复制为临时文件
'参数：
'lngSys:系统编号
'Action:操作类型（用以区别是操作哪个表）
'---系统100,Zl_Lob_Read
'0-病历标记图形;1-病历文件格式;2-病历文件图形;3-病历范文格式;4-病历范文图形;
'5-电子病历格式;6-电子病历图形;7-病历页面格式(图形)；8-电子病历附件;9-体温重叠标记
'10-临床路径文件,11-临床路径图标;14-人员证书记录;15-人员表;16-人员照片;
'17-药品规格(使用说明);18-药品规格(图片);
'19-部门扩展信息;20-人员扩展信息;22-医嘱报告内容;23-供应商照片;24-自定义申请单文件;25-医嘱申请单文件
'26-门诊路径文件,27-病人照片,28-咨询图片元素,29-咨询段落目录
'--系统600，ZL6_Lob_Read
'0-设备照片
'---系统2400,Zl24_Lob_Read
'手麻常用图形,无Action
'---系统2100,Zl21_Lob_Read
'1-体质类型调养;2-体检体辨结论(该图片只有读取，没有保存);3-体检申报记录;4-体检任务人员,5-体检任务结果
'---系统2500,ZL25_Lob_Read
'0-微生物涂片报告
'---系统2600,Zl26_Lob_Read
'14-导诊控件目录,15-导诊资源目录
'      KeyWord:确定数据记录的关键字，复合关键字以逗号分隔(仅5-电子病历格式为复合)
'      strFile:用户指定存放的文件名；不指定时，自动取临时文件名
'bytFunc-0-BLOB,1-CLOB
'bytMoved=0正常记录,1读取转储后备表记录
'返回：存放内容的文件名，失败则返回零长度""
    Const conChunkSize As Long = 10240
    Dim rsLOB       As ADODB.Recordset
    Dim lngFileNum  As Long, lngCount       As Long, lngBound       As Long
    Dim aryChunk()  As Byte, strText        As String
    Dim strSQL      As String
    Dim objFile     As New FileSystemObject
    Dim lngCurSize  As Long
    
    Err = 0: On Error GoTo Errhand
    Select Case lngSys \ 100
        Case 1
            strSQL = "Select Zl_Lob_Read([1],[2],[3],[4],[5]) as 片段 From Dual"
        Case 6
            strSQL = "Select Zl6_Lob_Read([1],[2],[3],[4],[5]) as 片段 From Dual"
        Case 24
            strSQL = "Select Zl24_Lob_Read([2],[3]) as 片段 From Dual"
        Case 21
            strSQL = "Select Zl21_Lob_Read([1],[2],[3]) as 片段 From Dual"
        Case 25
            strSQL = "Select Zl25_Lob_Read([1],[2],[3],[4],[5]) as 片段 From Dual"
        Case 26
            strSQL = "Select Zl26_Lob_Read([1],[2],[3]) as 片段 From Dual"
    End Select
    If strSQL = "" Then strFile = "": Exit Function
    If bytFunc = 0 Then 'BLOB
        If strFile = "" Then
            strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\" & objFile.GetTempName
        End If
        lngFileNum = FreeFile
        Open strFile For Binary As lngFileNum
        lngCount = 0
        lngCurSize = 0
        Do
            Set rsLOB = OpenSQLRecord(Sel_His_DB, strSQL, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLOB.EOF Then Exit Do
            If IsNull(rsLOB.Fields(0).value) Then Exit Do
            strText = rsLOB.Fields(0).value
            If lngCurSize = 0 Then
                lngCurSize = Len(strText) / 2
                If lngCurSize = 0 Then Exit Do
                ReDim aryChunk(lngCurSize - 1) As Byte
            ElseIf lngCurSize <> Len(strText) / 2 Then '防止重复分配内存
                lngCurSize = Len(strText) / 2
                If lngCurSize = 0 Then Exit Do
                ReDim aryChunk(lngCurSize - 1) As Byte
            End If
            For lngBound = LBound(aryChunk) To UBound(aryChunk)
                aryChunk(lngBound) = CByte("&H" & Mid(strText, lngBound * 2 + 1, 2))
            Next
            Put lngFileNum, , aryChunk()
            lngCount = lngCount + 1
        Loop
        Close lngFileNum
        If lngCount = 0 Then Kill strFile: strFile = ""
    Else  'CLOB
        lngCount = 0
        strFile = ""
        Do
            Set rsLOB = OpenSQLRecord(Sel_His_DB, strSQL, "zllobRead", Action, KeyWord, lngCount, bytMoved, bytFunc)
            If rsLOB.EOF Then Exit Do
            If IsNull(rsLOB.Fields(0).value) Then Exit Do
            strText = rsLOB.Fields(0).value
            strFile = strFile & strText
            lngCount = lngCount + 1
        Loop
    End If
    ReadLob = strFile
    Exit Function
Errhand:
    If bytFunc = 0 Then
        Close lngFileNum
        Kill strFile: ReadLob = ""
    End If
    Err.Clear
End Function
