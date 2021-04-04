Attribute VB_Name = "mdlRunAfter"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2018/12/25
'模块           mdlRunAfter
'说明           延迟脚本执行类
'==================================================================================================
Private Const mstrCurModule     As String = "mdlRunAfter"           '当前模块名称
'说明：
'1、一个数据库可能有多个在线库，如LIS和标准版在一个实例上独立安装。
'2、历史库的延后修正脚本没有需要在在线库执行的脚本。历史库所有者是单独的库。
'3、统计信息收集只需要当前实例的DBA。
'4、延后执行类型的升级脚本暂时不处理，仅保留接口与结构。
'由于以上原因只考虑DBA用户与历史库用户验证。
Private Enum ScriptStruct
    ESS_版本号 = 0
    ESS_并行 = 1
    ESS_服务器 = 2
    ESS_所有者 = 3
    ESS_所有者密码 = 4
    ESS_管理工具密码 = 5
    ESS_DBA用户名 = 6
    ESS_DBA密码 = 7
    ESS_历史库 = 8
    ESS_历史库脚本 = 9
    ESS_统计信息 = 10
    ESS_延迟执行脚本 = 11
End Enum
Private mstrServer              As String                           '当前服务器名称
Private mrsHistoryDeferred      As ADODB.Recordset                  '当前服务器包含历史库的延迟执行脚本。暂时不支持，仅保留兼容
Private mrsAppDeferred          As ADODB.Recordset                  '当前服务器包含的应用系统延迟执行脚本。暂时不支持，仅保留兼容
Private mrsToolDeferred         As ADODB.Recordset                  '当前服务器包含的ZLTOOLS延迟执行脚本。暂时不支持，仅保留兼容
'"ID", adInteger, Empty, Empty, "系统", adInteger, Empty, Empty, "BAKDBName", adVarChar, 100, Empty, _
"BAKUser", adVarChar, 100, Empty, "服务器", adVarChar, 500, Empty, "DBLINK", adVarChar, 200, Empty, _
"SQL", adVarChar, 500, Empty, "ExecOrder", adInteger, Empty, Empty, "FixType", adInteger, Empty, Empty, _
"ExecDB", adInteger, Empty, Empty, "ExecLater", adInteger, Empty, Empty,"DB_ID", adInteger, Empty, Empty,ScriptNO", adInteger, Empty, Empty, "DDLParallel", adInteger, Empty, Empty))
Private mrsHisScript            As ADODB.Recordset                  '当前服务器包含的历史库修正脚本。历史库当前没有需要在在线库执行的SQL
'"ID", adInteger, Empty, Empty, "Owner", adInteger, Empty, Empty, "TableName", adVarChar, 100, Empty, _
"SQL", adVarChar, 500, Empty,"ScriptNO", adInteger, Empty, Empty))
Private mrsStatistics           As ADODB.Recordset                  '当前服务器包含的统计信息收集脚本
'"系统编号", adInteger, Empty, Empty, "系统名称", adVarChar, 50, Empty, "系统版本", adVarChar, 20, Empty, "配置文件", adVarChar, 2000, Empty, _
"编号", adInteger, Empty, Empty, "名称", adVarChar, 30, Empty, "所有者", adVarChar, 50, Empty, _
"当前", adInteger, Empty, Empty, "DB连接", adVarChar, 200, Empty, "密码", adVarChar, 100, Empty, _
"服务器", adVarChar, 500, Empty, "升级", adInteger, Empty, Empty, "当前版本", adVarChar, 20, Empty, _
"目标版本", adVarChar, 20, Empty, "中止信息", adVarChar, 2000, Empty, "可升级", adInteger, 1, 0, "检查结果", adVarChar, 2000, Empty, _
"提前目标版本", adVarChar, 20, Empty, "提前中止信息", adVarChar, 2000, Empty, "可提前升级", adInteger, 1, 0, "提前检查结果", adVarChar, 2000, Empty, _
"验证", adInteger, Empty, Empty))
Private mrsHistory              As ADODB.Recordset                  '历史库信息
Private mlngCurFileLen          As Long                             '当前准备处理的文件长度
Private mlngCurMaxScriptNo      As Long                             '当前准备处理的文件的最大脚本段

Private mstrDBAUser             As String                           'DBA用户名
Private mstrDBAPWD              As String                           'DBA用户密码
Private mblnDBAOK               As Boolean                          'DBA用户是否连接成功
Private mcnDBA                  As ADODB.Connection

Private mblnExecAgain           As Boolean                          '当前服务器由于有新增脚本，因此在此执行。此时不再进行用户验证
Private mcllHistory             As New Collection                   '已经验证的历史库
Private mlngHisID               As Long                             '历史库ID标记，方便选择
'设置服务器
Public Property Get Server() As String
    Server = mstrServer
End Property

Public Property Let Server(strServer As String)
    If mstrServer <> strServer Then
        Set mrsHistory = Nothing
        Set mrsHisScript = Nothing
        Set mrsStatistics = Nothing
        Set mcllHistory = Nothing
        mlngCurFileLen = 0
        mlngCurMaxScriptNo = -1
        mblnDBAOK = False
        mstrDBAUser = ""
        mstrDBAPWD = ""
        mblnExecAgain = False
        Set mcnDBA = Nothing
    Else
        mblnExecAgain = True
        Set mrsHisScript = Nothing
        Set mrsStatistics = Nothing
        Set mrsHistory = Nothing
    End If
    mstrServer = strServer
End Property
'设置DAB
Public Property Get DBAUser() As String
    DBAUser = mstrDBAUser
End Property

Public Property Let DBAUser(strDBAUser As String)
    mstrDBAUser = strDBAUser
End Property

Public Property Get DBAPWD() As String
    DBAPWD = mstrDBAPWD
End Property

Public Property Let DBAPWD(strDBAPWD As String)
    mstrDBAPWD = strDBAPWD
End Property

Public Property Get IsDBAOK() As Boolean
    IsDBAOK = mblnDBAOK
End Property

Public Property Let IsDBAOK(blnDBAOK As Boolean)
    mblnDBAOK = blnDBAOK
End Property
'--------------------------------------------------------------------------------------------------
'接口               RunUpgradeAfter
'功能               执行任务是否完成
'返回值
'入参列表:
'参数名         类型                        说明
'-------------------------------------------------------------------------------------------------
Public Function RunUpgradeAfter() As Boolean
    Dim lngScriptNo         As Long, lngOjbectNo    As Long, lngSQLID     As Long, intIniFileNo   As Integer, blnOk As Boolean
    Dim arrTmp              As Variant
    Dim conTmp              As ADODB.Connection
    Dim strLastCon          As String
    Dim cllHisCon           As New Collection
    Dim comTmp              As New ADODB.Command
    Dim i                   As Long, intLastDDLParallel As Integer
    Dim lngTotal            As Long, lngCurCount        As Long
    
    
    On Error GoTo errH
    '说明有进程正在执行，则退出，不用验证。预先锁定，防止其他进程处理
    If Not SaveOrReadExecuteRunAfterInfo(intIniFileNo, lngScriptNo, lngOjbectNo, lngSQLID) Then
        RunUpgradeAfter = True
        Exit Function
    End If
    Call ShowFlash("正在读取延后执行脚本。", , , Server)
    '读取脚本
    If Not ReadRunAfter(lngScriptNo, lngOjbectNo, lngSQLID) Then
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, lngScriptNo, lngOjbectNo, lngSQLID)
        RunUpgradeAfter = True
        Exit Function
    End If
    Call ShowFlash
    If Not RunAterIdentifyUsers Then
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, lngScriptNo, lngOjbectNo, lngSQLID)
        RunUpgradeAfter = True
        Exit Function
    End If
    Call ShowFlash
    If Not mrsHisScript Is Nothing Then
        mrsHisScript.Filter = ""
        lngTotal = mrsHisScript.RecordCount
    End If
    If Not mrsStatistics Is Nothing And IsDBAOK Then
        mrsStatistics.Filter = ""
        lngTotal = lngTotal + mrsStatistics.RecordCount
    End If
    If lngTotal = 0 Then lngTotal = 1
    
    For i = lngScriptNo To mlngCurMaxScriptNo
        If Not mrsHisScript Is Nothing Then
            mrsHisScript.Filter = "ScriptNO=" & i
            mrsHisScript.Sort = "ID"
            Do While Not mrsHisScript.EOF
                If strLastCon <> "K_" & mrsHisScript!服务器 & "|" & mrsHisScript!BAKUser Then
                    '关闭并行
                    If strLastCon <> "" Then
                        If Not conTmp Is Nothing Then
                            Call SetSessionParallel(conTmp, False, intLastDDLParallel)
                        End If
                    End If
                    strLastCon = "K_" & mrsHisScript!服务器 & "|" & mrsHisScript!BAKUser
                    intLastDDLParallel = Val(mrsHisScript!DDLParallel)
                    If Not InCollection(cllHisCon, strLastCon) Then
                        mrsHistory.Filter = "系统编号=" & mrsHisScript!系统 & " And 名称='" & mrsHisScript!BAKDBName & "'"
                        Set conTmp = gobjRegister.GetConnection(mrsHistory!服务器, mrsHistory!所有者, mrsHistory!密码, False, MSODBC, "", False)
                        If conTmp.State = adStateClosed Then
                            Set conTmp = Nothing
                        End If
                        cllHisCon.Add conTmp, strLastCon
                        '开启并行
                        If Not conTmp Is Nothing Then
                            Call SetSessionParallel(conTmp, True, intLastDDLParallel)
                        End If
                    Else
                        Set conTmp = cllHisCon(strLastCon)
                    End If
                End If
                lngCurCount = lngCurCount + 1
                Call ShowFlash("进  度：" & lngCurCount & "/" & lngTotal & "  历史库索引约束创建", lngCurCount / lngTotal, mrsHisScript!SQL, Server)
                If Not conTmp Is Nothing Then
                    On Error Resume Next
                    If mrsHisScript!ExecDB = 1 Then
                        '当前在在线库执行的脚本未放在延后执行
'                        Set comTmp.ActiveConnection = mcnOracle
                    Else
                        Set comTmp.ActiveConnection = conTmp
                    End If
                    comTmp.CommandText = mrsHisScript!SQL
                    DoEvents
                    comTmp.Execute
                    If Err.Number <> 0 Then
                        Debug.Print Err.Description & "-" & mrsHisScript!SQL
                        Err.Clear
                    End If
                    On Error GoTo errH
                End If
                Call SaveOrReadExecuteRunAfterInfo(intIniFileNo, i, 2, mrsHisScript!Id)
                mrsHisScript.MoveNext
            Loop
        End If
        '关闭并行。放在这里，因为可能所有脚本都是一个数据库，减少执行次数。
        If strLastCon <> "" Then
            If Not conTmp Is Nothing Then
                Call SetSessionParallel(conTmp, False, intLastDDLParallel)
            End If
        End If
        '标记当前脚本序列的历史库以及修正完毕
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo, i, 3, 0)
        If IsDBAOK Then
            If mcnDBA Is Nothing Then
                Set mcnDBA = gobjRegister.GetConnection(mstrServer, DBAUser, DBAPWD, False, MSODBC, "", False)
            End If
            If Not mrsStatistics Is Nothing And mcnDBA.State = adStateOpen Then
                mrsStatistics.Filter = "ScriptNO=" & i
                mrsStatistics.Sort = "ID" '添加DB_ID字段保证排序的唯一性
                Do While Not mrsStatistics.EOF
                    '调用包时指定参数名，仅ODBC连接方式支持
                    '用connection对象，excute方法的Options参数值为这几个都可以：adCmdUnknown 'adCmdStoredProc 'adExecuteNoRecords
                    '用Command对象，必须指定CommandType = adCmdStoredProc
                    On Error Resume Next
                    lngCurCount = lngCurCount + 1
                    Call ShowFlash("进  度：" & lngCurCount & "/" & lngTotal & "  统计信息收集", lngCurCount / lngTotal, mrsStatistics!SQL, Server)
                    DoEvents
                    mcnDBA.Execute mrsStatistics!SQL & "", , adCmdStoredProc
                    If Err.Number <> 0 Then
                        Debug.Print Err.Description & "-" & mrsStatistics!SQL
                        Err.Clear
                    End If
                    On Error GoTo errH
                    Call SaveOrReadExecuteRunAfterInfo(intIniFileNo, i, 3, mrsStatistics!Id)
                    mrsStatistics.MoveNext
                Loop
            End If
        End If
        '标记，当前脚本序列的统计信息已经收集完毕
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo, i + 1, 0, 0)
    Next

    If mlngCurFileLen = FileLen(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".SQL") Then '防止在脚本执行中脚本发生变化
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, mlngCurMaxScriptNo + 1, 0, 0)
        Kill IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".bini"
        Name IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".SQL" As IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & "_" & Format(Now, "YYYYMMDDHHmmss") & ".SQL"
        RunUpgradeAfter = True
    Else
        Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, mlngCurMaxScriptNo + 1, 0, 0)
    End If
    Call ShowFlash
    Exit Function
errH:
    Call SaveOrReadExecuteRunAfterInfo(intIniFileNo * -1, , , , True)
    Call ShowFlash
    RunUpgradeAfter = True
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function


'--------------------------------------------------------------------------------------------------
'接口           ReadRunAfter
'功能           -读取RunAfter.SQL的脚本内容
'返回值         Boolean                    是否读取成功
'入参列表:
'参数名         类型                        说明
'lngScriptNo    Long                        执行到的脚本段位置
'lngOjbectNo    Long                        执行到的脚本段中对象的序号
'lngSQLID       Long                        已经执行的SQLID
'说明：
'历史库修正的SQL以最后一次该历史库的修正为准。
'统计信息收集以最后一次为基准，逐渐递增相比最后一次不存在的表对象。
'-------------------------------------------------------------------------------------------------
Public Function ReadRunAfter(ByVal lngScriptNo As Long, ByVal lngOjbectNo As Long, ByVal lngSQLID As Long) As Boolean
    Dim objTxt          As TextStream, strLine              As String
    Dim lngCurScriptNo  As Long, arrScript()                As Variant, arrLine             As Variant, i               As Long
    Dim cllStatictics    As New Collection
    Dim rsTmpHis        As ADODB.Recordset, rsTmpHisAfter   As ADODB.Recordset, rsTmpSta As ADODB.Recordset
    Dim conTmp          As ADODB.Connection
    Dim strFileter      As String

    On Error GoTo errH
    If gobjFSO.FileExists(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".SQL") Then
        gobjFSO.CopyFile IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".SQL", IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfterTmp" & mstrServer & ".SQL", True '复制一个临时文件来执行脚本，防止文件在执行过程中写入
        '--[SERVER]:Oracle
        '--[SCRIPT]:SerializeMulti(版本标识,并行参数,服务器,应用系统所有者, Sm4EncryptEcb(应用系统所有者密码), Sm4EncryptEcb(管理工具密码), DBA用户名, Sm4EncryptEcb(DBA密码), Sm4EncryptEcb(gclsBase.Serialize(历史库信息记录集), G_APP_KEY), 历史库脚本记录集, 统计信息收集记录集, 延迟执行脚本记录集)
        '--[脚本详情]:
        Set objTxt = gobjFSO.OpenTextFile(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfterTmp" & mstrServer & ".SQL", ForReading)
        Do While Not objTxt.AtEndOfStream
            strLine = objTxt.ReadLine
            If strLine Like "--[[]SCRIPT[]]:*" Then
                arrLine = UnSerializeMulti(Mid(strLine, Len("--[SCRIPT]:*")))
                Set arrLine(ESS_历史库) = UnSerialize(Sm4DecryptEcb(arrLine(ESS_历史库), G_APP_KEY))
                arrLine(ESS_所有者密码) = Sm4DecryptEcb(arrLine(ESS_所有者密码), G_APP_KEY)
                arrLine(ESS_DBA密码) = Sm4DecryptEcb(arrLine(ESS_DBA密码), G_APP_KEY)
                arrLine(ESS_管理工具密码) = Sm4DecryptEcb(arrLine(ESS_管理工具密码), G_APP_KEY)
                ReDim Preserve arrScript(lngCurScriptNo)
                arrScript(lngCurScriptNo) = arrLine
                lngCurScriptNo = lngCurScriptNo + 1
            End If
        Loop

        '历史库用户验证
        For i = UBound(arrScript) To LBound(arrScript) Step -1
            If Not arrScript(i)(ESS_历史库) Is Nothing Then
                Set rsTmpHis = arrScript(i)(ESS_历史库)
                Set rsTmpHisAfter = arrScript(i)(ESS_历史库脚本)
                '存在脚本且脚本未完全执行或未执行
                If Not rsTmpHisAfter Is Nothing And (i > lngScriptNo Or i = lngScriptNo And lngOjbectNo < 3) Then
                    If i = lngScriptNo Then
                        strFileter = "ID>" & lngSQLID
                    Else
                        strFileter = ""
                    End If
                    rsTmpHisAfter.Filter = strFileter
                    '存在未执行的脚本
                    If rsTmpHisAfter.RecordCount <> 0 Then
                        Do While Not rsTmpHis.EOF
                            '倒序查找历史库，第一个遇到的就加入，之后的不再加入
                            If Not InCollection(mcllHistory, "K_" & rsTmpHis!系统编号 & "_" & rsTmpHis!名称) Then
                                rsTmpHisAfter.Filter = strFileter & IIf(strFileter <> "", " And ", "") & "系统=" & rsTmpHis!系统编号 & " And BAKDBName='" & rsTmpHis!名称 & "'"
                                '当前历史库存在未执行的脚本，则加入该脚本以及该历史库
                                If rsTmpHisAfter.RecordCount <> 0 Then
                                    If mrsHistory Is Nothing Then '初始化记录集
                                        Set mrsHistory = CopyNewRec(rsTmpHis, True)
                                        Set mrsHisScript = CopyNewRec(rsTmpHisAfter, True, , Array("ScriptNO", adInteger, Empty, Empty, "DDLParallel", adInteger, Empty, Empty))
                                    End If
                                    Call RecDataAppend(mrsHistory, rsTmpHis, 1, , , True)
                                    Call RecDataAppend(mrsHisScript, rsTmpHisAfter, , "-ScriptNO,DDLParallel", , , Array("ScriptNO", i, "DDLParallel", Val(arrScript(i)(ESS_并行))))
                                    '链接验证
                                    Set conTmp = gobjRegister.GetConnection(rsTmpHis!服务器, rsTmpHis!所有者, rsTmpHis!密码, False, MSODBC, "", False)
                                    If conTmp.State = adStateClosed Then
                                        mrsHistory.Update Array("ID", "当前版本", "目标版本", "验证"), Array(mrsHistory.RecordCount, Null, Null, 0)
                                        mcllHistory.Add 0, "K_" & rsTmpHis!系统编号 & "_" & rsTmpHis!名称
                                    Else
                                        mrsHistory.Update Array("ID", "目标版本", "验证"), Array(mrsHistory.RecordCount, Null, 1)
                                        mcllHistory.Add 1, "K_" & rsTmpHis!系统编号 & "_" & rsTmpHis!名称
                                    End If
                                    If conTmp.State = adStateOpen Then
                                        conTmp.Close
                                    End If
                                    Set conTmp = Nothing
                                End If
                            End If
                            rsTmpHis.MoveNext
                        Loop
                    End If
                End If
            End If
            If Not arrScript(i)(ESS_统计信息) Is Nothing Then
                If Not IsDBAOK Then
                    If DBAUser <> arrScript(i)(ESS_DBA用户名) Or DBAPWD <> arrScript(i)(ESS_DBA密码) Then
                        DBAUser = arrScript(i)(ESS_DBA用户名)
                        DBAPWD = arrScript(i)(ESS_DBA密码)
                        Set conTmp = gobjRegister.GetConnection(Server, DBAUser, DBAPWD, False, MSODBC, "", False)
                        If conTmp.State = adStateOpen Then
                            IsDBAOK = True
                            conTmp.Close
                        End If
                        Set conTmp = Nothing
                    End If
                End If
                Set rsTmpSta = arrScript(i)(ESS_统计信息)
                If mrsStatistics Is Nothing Then
                    Set mrsStatistics = CopyNewRec(rsTmpSta, True, , Array("ScriptNO", adInteger, Empty, Empty))
                End If
                '存在脚本且脚本未完全执行或未执行
                If i >= lngScriptNo Then
                    If i = lngScriptNo And lngOjbectNo = 3 Then
                        rsTmpSta.Filter = "ID>" & lngSQLID
                    Else
                        rsTmpSta.Filter = ""
                    End If
                    '存在未执行的脚本
                    If rsTmpSta.RecordCount <> 0 Then
                        Do While Not rsTmpSta.EOF
                            If Not InCollection(cllStatictics, "K_" & rsTmpSta!Owner & "." & rsTmpSta!TableName) Then
                                '若该统计信息的收集对象未添加，则添加，添加当前行并将游标回滚
                                Call RecDataAppend(mrsStatistics, rsTmpSta, 1, "-ScriptNO", , True, Array("ScriptNO", i))
                                cllStatictics.Add 1, "K_" & rsTmpSta!Owner & "." & rsTmpSta!TableName
                            End If
                            rsTmpSta.MoveNext
                        Loop
                    End If
                End If
            End If
        Next
        mlngCurFileLen = FileLen(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfterTmp" & mstrServer & ".SQL")
        mlngCurMaxScriptNo = UBound(arrScript)
        ReadRunAfter = True
        objTxt.Close
        Set objTxt = Nothing
        Kill IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfterTmp" & mstrServer & ".SQL"
    End If
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

Private Sub SetSessionParallel(ByRef cnInput As ADODB.Connection, Optional ByVal blnEnabled As Boolean, Optional ByVal intDDLParallel As Integer)
'启用或禁用DDL
    Dim strSQL As String, rsTmp As ADODB.Recordset

    On Error GoTo errH
    If intDDLParallel <= 1 Then Exit Sub
    If blnEnabled Then
        strSQL = "Alter Session FORCE PARALLEL DDL PARALLEL " & intDDLParallel
        cnInput.Execute strSQL
    Else
        strSQL = "ALTER Session DISABLE PARALLEL DDL "
        cnInput.Execute strSQL
        strSQL = "Select 'alter index ' || Index_Name || ' noparallel' SQL" & vbNewLine & _
                    "From User_Indexes" & vbNewLine & _
                    "Where Degree Not In ('0', '1') and index_type='NORMAL' And temporary='N'" & vbNewLine & _
                    "Union All" & vbNewLine & _
                    "Select 'alter table ' || Table_Name || ' noparallel' SQL" & vbNewLine & _
                    "From User_Tables" & vbNewLine & _
                    "Where Degree != ('         1')"
        Set rsTmp = gobjRegister.OpenSQLRecord(cnInput, strSQL, App.Title)
        On Error Resume Next
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                cnInput.Execute rsTmp!SQL, , adCmdText
                If Err.Number <> 0 Then
                    Err.Clear
                End If
                rsTmp.MoveNext
            Loop
        End If
    End If
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Sub

'--------------------------------------------------------------------------------------------------
'接口           RunAterIdentifyUsers
'功能           -验证RunAter脚本执行的历史库用户信息，同时验证DBA用户（用于统计信息）
'返回值         ADODB.Recordset            返回的历史库验证信息
'入参列表:
'参数名         类型                        说明
'blnDo          Boolean                     是否执行延迟执行脚本
'-------------------------------------------------------------------------------------------------
Private Function RunAterIdentifyUsers() As Boolean
    Dim rsTmp       As ADODB.Recordset
    '该库的第二次验证
    If mblnExecAgain Then
        RunAterIdentifyUsers = True
        Exit Function
    End If
    On Error GoTo errH
    If Not mrsHistory Is Nothing Then
        mrsHistory.Filter = "验证=0"
        If mrsHistory.RecordCount <> 0 Then
            '将未验证通过的单独独立出来，防止界面出现无用数据
            Set rsTmp = CopyNewRec(mrsHistory)
            Call RecDelete(mrsHistory, "验证=0")
            Call RecUpdate(rsTmp, "", "密码", "")
        End If
    End If
    If Not mrsStatistics Is Nothing And Not IsDBAOK Then
        Call frmUsers.ShowMe(rsTmp, True)
    ElseIf Not rsTmp Is Nothing Then
        Call frmUsers.ShowMe(rsTmp)
    Else
        RunAterIdentifyUsers = True
        Exit Function
    End If
    If Not rsTmp Is Nothing Then
        rsTmp.Filter = ""
        Call RecDataAppend(mrsHistory, rsTmp)
    End If
    RunAterIdentifyUsers = True
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    Err.Clear
End Function

''调用该方法后，注意关闭
Public Function SaveOrReadExecuteRunAfterInfo(intFileNo As Integer, Optional lngScriptNo As Long, Optional lngOjbectNo As Long, Optional lngSQLID As Long, Optional ByVal blnForceClaose As Boolean) As Boolean
'功能：保存RunAfter的脚本执行情况
    '[脚本位置]脚本序号,对象序号,SQL序号
    '字节数：10  4    1   4    1  4
    On Error GoTo errH
    If intFileNo = 0 Then
        intFileNo = FreeFile()
        Open IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".bini" For Binary Access Read Write Lock Read Write As intFileNo
        If FileLen(IIf(gblnInIDE, "C:\APPSOFT", App.Path) & "\RuntimeFile\RunAfter_" & mstrServer & ".bini") < 24 Then
            Put #intFileNo, 1, StrConv("[脚本位置]", vbFromUnicode)
            Put #intFileNo, 11, lngScriptNo
            Put #intFileNo, 15, CByte(44)
            Put #intFileNo, 16, lngOjbectNo
            Put #intFileNo, 20, CByte(44)
            Put #intFileNo, 21, lngSQLID
        Else
            Get #intFileNo, 11, lngScriptNo
            Get #intFileNo, 16, lngOjbectNo
            Get #intFileNo, 21, lngSQLID
        End If
    ElseIf intFileNo > 0 Then
        Put #intFileNo, 11, lngScriptNo
        Put #intFileNo, 16, lngOjbectNo
        Put #intFileNo, 21, lngSQLID
    Else
        If Not blnForceClaose Then
            Put #Abs(intFileNo), 11, lngScriptNo
            Put #Abs(intFileNo), 16, lngOjbectNo
            Put #Abs(intFileNo), 21, lngSQLID
        End If
        Close #Abs(intFileNo)
    End If
    SaveOrReadExecuteRunAfterInfo = True
    Exit Function
errH:
    '锁定读写，防止多进程同时执行。
    Err.Clear
End Function
'




