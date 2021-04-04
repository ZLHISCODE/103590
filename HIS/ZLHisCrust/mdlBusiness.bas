Attribute VB_Name = "mdlBusiness"
Option Explicit
'调用该程序要实现的操作
Public Enum OperateType
    OT_Repair = 0                                       '主动修复，相当于升级,不判断是否预升级完成
    OT_PreUpgrade = 1                                   '提前升级，将升级文件放在临时目录
    OT_OfficialUpgrade = 2                              '从提前升级目录中或者服务器目录中提取文件到安装路径
    OT_Other = 3                                        '暂时只有文件收集，收集APPSOFT目录下的指定类型文件到服务器
End Enum

Public Enum OperateStep
    OS_NotInProcessing = 0                              '未执行
    OS_Completed = 1                                    '执行完成
    OS_Failure = 2                                      '执行失败
    OS_InProcessing = 3                                 '执行中
End Enum

'错误类型
Public Enum MsgType
    MT_MsgHeader = 0                                    '消息头
    MT_InitEnv = 1                                      '该错误类型未标识
    MT_SvrConn = 2                                      '连接服务器错误
    MT_ChcekUpdate = 3                                  '更新检查
    MT_DownAndDec = 4                                   '下载解压部件错误
    MT_SetUp = 5                                        '讲部件放在安装目录出错
    MT_RegCom = 6                                       '部件注册错误
    MT_ExeBat = 7                                       '执行批处理错误
    MT_MsgFoot = 8                                      '消息尾部
End Enum
'部件注册类型
Public Enum RegFileType
    RFT_NotReg = 0                  '不注册的对象
    RFT_NormalReg = 1               '常规注册，自动识别.NET部件，.NET部件通过Regasm注册，其他通过调用DLLRegServer注册
    RFT_NETGAC = 2                  'NET程序集注册，通过gacutil注册到全局程序集缓存
    RFT_NETServer = 3               'NET服务注册，通过installUtil进行安装卸载。
    RFT_NETComReg = 4               '.NET Com部件注册，通过调用Regasm完成
    RFT_VBComReg = 5                '通过手写注册表注册
    RFT_DelphiComReg = 6            'DelphiCom注册，通过DLLRegServer注册
    RFT_PBComReg = 7                'PBCom注册，通过DLLRegServer注册
End Enum
'文件类型
Public Enum FileType
    FT_Public = 0                   '产品公共部件
    FT_Apply = 1                    '产品应用部件
    FT_Help = 2                     '产品帮助文件
    FT_AdditionFile = 3             '产品附加文件
    FT_Other = 4                    '三方产品文件
    FT_System = 5                   '系统文件
End Enum
Public Function SetOperateProcess(ByVal otCurType As OperateType, ByVal osCurStep As OperateStep, Optional ByVal strMsg As String, Optional ByVal lngBeach As Long) As Boolean
'功能：更新操作进度。
'参数：otCurType=当前操作类型
'      osCurStep=当前步骤
'      lngBeach=修正的批次
'      strMsg=操作信息
'返回：是否执行成功
    Dim blnComplete As Boolean, strSQL As String
    Dim strBeach As String
    
    gobjTrace.WriteSection "标记升级进度", SL_LevelThree
    strMsg = MidB(strMsg, 1, glngNoteLength - 30)
    On Error Resume Next
    strSQL = "zlTOOLS.Zl_Zlclients_UpdateProcess('" & gstrComputerName & "'," & otCurType & "," & osCurStep & "," & SQLAdjust(strMsg) & "," & IIf(lngBeach <> 0 And osCurStep = OS_Completed, lngBeach, "Null") & ")"
    Call ExecuteProcedure(strSQL, "SetOperateProcess")
    If Err.Number <> 0 Then
        gobjTrace.WriteInfo "SetOperateProcess", "标记结果", "待定", "标记SQL", Replace(Replace(strSQL, Chr(10), ""), Chr(13), ""), "错误信息", Err.Description
        Err.Clear
        blnComplete = osCurStep = OS_Completed
        Select Case otCurType
            Case OT_OfficialUpgrade '正式升级完成则清除预升级相关信息，主动修复相关信息，并取消预升级标志与升级标志
                strSQL = "Update zlTOOLS.zlClients Set 升级情况=" & osCurStep & " ,升级说明=" & SQLAdjust(strMsg) & "" & IIf(lngBeach <> 0 And osCurStep = OS_Completed, ",批次=" & lngBeach, "") & IIf(blnComplete, ",升级标志=0,是否预升级=0,修复状态=0,预升完成=0,预升级说明=Null,修复说明=Null,是否立即升级=0", "") & " Where 工作站 = '" & gstrComputerName & "'"
            Case OT_PreUpgrade
                strSQL = "Update zlTOOLS.zlClients Set 预升完成=" & osCurStep & " ,预升级说明=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",是否预升级=0", "") & " Where 工作站 = '" & gstrComputerName & "'"
            Case OT_Repair '主动修复完成则清除预升级相关信息，主动修复相关信息，并取消预升级标志与升级标志
                strSQL = "Update zlTOOLS.zlClients Set 修复状态=" & osCurStep & " ,修复说明=" & SQLAdjust(strMsg) & "" & IIf(lngBeach <> 0 And osCurStep = OS_Completed, ",批次=" & lngBeach, "") & IIf(blnComplete, ",升级标志=0,是否预升级=0,升级情况=0,预升完成=0,预升级说明=Null,升级说明=Null,是否立即升级=0", "") & " Where 工作站 = '" & gstrComputerName & "'"
            Case OT_Other
                strSQL = "Update zlTOOLS.zlClients Set 收集状态=" & osCurStep & " ,收集说明=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",收集标志=0", "") & " Where 工作站 = '" & gstrComputerName & "'"
        End Select
        gcnOracle.Execute strSQL, , adCmdText
        If Err.Number <> 0 Then '执行SQL出错，说明结构还没升级上来，则执行老结构修正
            gobjTrace.WriteInfo "SetOperateProcess", "标记结果", "待定", "标记SQL", Replace(Replace(strSQL, Chr(10), ""), Chr(13), ""), "错误信息", Err.Description
            Err.Clear
            Select Case otCurType
                Case OT_OfficialUpgrade '正式升级完成则清除预升级相关信息，主动修复相关信息，并取消预升级标志与升级标志
                    strSQL = "Update zlTOOLS.zlClients Set 升级情况=" & osCurStep & " ,说明=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",升级标志=0,预升完成=0,是否立即升级=0", "") & " Where 工作站 = '" & gstrComputerName & "'"
                Case OT_PreUpgrade
                    strSQL = "Update zlTOOLS.zlClients Set 预升完成=" & osCurStep & " ,说明=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",升级标志=0", "") & " Where 工作站 = '" & gstrComputerName & "'"
                Case OT_Repair '主动修复完成则清除预升级相关信息，主动修复相关信息，并取消预升级标志与升级标志
                    strSQL = "Update zlTOOLS.zlClients Set 升级情况=" & osCurStep & " ,说明=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",升级标志=0,预升完成=0,是否立即升级=0", "") & " Where 工作站 = '" & gstrComputerName & "'"
                Case OT_Other
                    strSQL = "Update zlTOOLS.zlClients Set 说明=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",收集标志=0", "") & " Where 工作站 = '" & gstrComputerName & "'"
            End Select
            gcnOracle.Execute strSQL, , adCmdText
            If Err.Number <> 0 Then
                gobjTrace.WriteInfo "SetOperateProcess", "标记结果", "失败", "标记SQL", Replace(Replace(strSQL, Chr(10), ""), Chr(13), ""), "错误信息", Err.Description
                Call RecordErrMsg(MT_InitEnv, "标记任务执行情况", "请确认管理工具对象与权限完整。" & Err.Description)
                Call RecordErrMsg(MT_MsgFoot, "消息尾", "结果:升级失败 时间:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
                Err.Clear
                MsgBox "无法标记任务执行情况，请联系管理员确认管理工具对象权限完整！", vbInformation, App.Title
                Exit Function
            End If
        End If
    End If
    gobjTrace.WriteInfo "SetOperateProcess", "标记结果", "成功", "标记SQL", Replace(Replace(strSQL, Chr(10), ""), Chr(13), "")
    SetOperateProcess = True
End Function

Public Function CheckJobs() As Boolean
'功能:检查并获取升级程序的任务
    Dim rsTmp       As ADODB.Recordset, strSQL  As String
    Dim datCur      As Date, blnOnlyOfficialUp  As Boolean, blnOnlyPreUp    As Boolean
    Dim blnPreUp    As Boolean, blnOfficialUp   As Boolean, blnPreComplete  As Boolean, blnCollect  As Boolean
    Dim strMsg      As String
    
    On Error GoTo ErrH
    '以下代码一般不可能出错
    datCur = Currentdate
    '判断任务是否合理，获取是否启用了定时升级
    strSQL = "Select Max(内容) 内容 From ZLTOOLS.zlRegInfo Where 项目='客户端升级日期'"
    Set rsTmp = OpenSQLRecord(strSQL, "检查定时升级")
    If rsTmp!内容 & "" <> "" Then
        If CDate(Format(datCur, "YYYY-MM-DD hh:mm:ss")) >= CDate(Format(NVL(rsTmp!内容), "YYYY-MM-DD hh:mm:ss")) Then
            blnOnlyOfficialUp = True '只能正式升级
        Else
            blnOnlyPreUp = True '只能预升级
        End If
    Else
        blnOnlyOfficialUp = True
    End If
    gobjTrace.WriteInfo "CheckJobs", "是否只能正式升级", blnOnlyOfficialUp, "是否只能预升级", blnOnlyPreUp
    On Error Resume Next
    Set rsTmp = Nothing
    '可能没有是否预升级字段(因为预升级时候，数据库还没升级），因此需要错误忽略
    strSQL = "Select Nvl(是否预升级,0) 是否预升级, Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志 From ZLTOOLS.Zlclients Where 工作站 = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "检查当前任务", gstrComputerName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo ErrH
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!是否预升级 = 1
            blnOfficialUp = rsTmp!升级标志 = 1
            blnPreComplete = rsTmp!预升完成 = 1
            blnCollect = rsTmp!收集标志 = 1
        End If
    Else
        '优先新方式读取，失败再使用老方式，增加兼容性
        strSQL = "Select Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志 From ZLTOOLS.Zlclients Where 工作站 = [1]"
        Set rsTmp = OpenSQLRecord(strSQL, "检查当前任务", gstrComputerName)
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!升级标志 = 1
            blnOfficialUp = rsTmp!升级标志 = 1
            blnPreComplete = rsTmp!预升完成 = 1
            blnCollect = rsTmp!收集标志 = 1
        End If
    End If
    gobjTrace.WriteInfo "CheckJobs", "是否需要预升级", blnPreUp, "是否需要正式升级", blnOnlyPreUp, "上次预升级是否完成", blnPreComplete, "是否进行文件收集", blnCollect
    If gotCurType = OT_Repair Then
        If blnOnlyPreUp Then
            gotCurType = OT_PreUpgrade
        End If
    ElseIf (blnOfficialUp Or blnPreUp) And blnOnlyPreUp Then
        gotCurType = OT_PreUpgrade
    ElseIf (blnOfficialUp Or blnPreUp) And blnOnlyOfficialUp Then
        gotCurType = OT_OfficialUpgrade
    ElseIf blnCollect Then
        gotCurType = OT_Other
    Else
        gobjTrace.WriteInfo "CheckJobs", "检测结果", "当前没有任何任务，系统将自动退出"
        Call RecordErrMsg(MT_InitEnv, "任务检测", "当前没有任何任务，系统将自动退出")
        CheckJobs = True
        Exit Function
    End If
    '预升级已经完成
    If blnPreComplete And gotCurType = OT_PreUpgrade Then
        gobjTrace.WriteInfo "CheckJobs", "检测结果", "当前只能预升级，但是预升级已经完成，系统将自动退出。"
        Call RecordErrMsg(MT_InitEnv, "任务检测", "当前只能预升级，但是预升级已经完成，系统将自动退出。")
        CheckJobs = True
        Exit Function
    End If
    gobjTrace.WriteInfo "CheckJobs", "检测结果", Decode(gotCurType, OT_OfficialUpgrade, "正式升级", OT_PreUpgrade, "预升级", OT_Repair, "修复或强制升级", OT_Other, "收集或其他")
    Set gclsConnect = GetFileConnect(strMsg)
    If gclsConnect Is Nothing Then
        gobjTrace.WriteInfo "CheckJobs", "连接失败", strMsg
        Call RecordErrMsg(MT_InitEnv, "任务检测", "无法连接文件服务器,不能继续进行操作。信息：" & strMsg)
        MsgBox "无法连接文件服务器，请联系管理员！信息：" & vbNewLine & strMsg, vbInformation, App.Title
        Exit Function
    End If
    CheckJobs = True
    Exit Function
ErrH:
    strMsg = Err.Description
    gobjTrace.WriteInfo "CheckJobs", "任务检测发生致命错误", strMsg
    MsgBox "任务检测发生致命错误，请联系管理员！信息：" & vbNewLine & strMsg, vbInformation, App.Title
    Err.Clear
End Function

Private Function GetFileConnect(ByRef strMsg As String) As clsConnect
'功能：获取服务器文件连接
    Dim objConn As New clsConnect
    Dim sctConnType As ServerConnectType
    Dim strServerID As String, strServer As String, strUser As String, strPwd As String, strPort As String, strCollectType As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnDefalut As Boolean, blnConnOK As Boolean
    Dim blnOldStype As Boolean
    On Error Resume Next
    If gotCurType = OT_Other Then
        strSQL = "Select 类型, 位置, 用户名, 密码, 端口, 收集类型 From Zltools.Zlupgradeserver Where Nvl(是否收集, 0) = 1"
        Set rsTmp = OpenSQLRecord(strSQL, "获取升级服务器编号", gstrComputerName)
        If Err.Number = 0 Then
            If Not rsTmp.EOF Then
                strServerID = rsTmp!编号 & ""
                sctConnType = IIf(rsTmp!类型 = 0, SCT_Share, SCT_FTP)
                strServer = rsTmp!位置
                strUser = rsTmp!用户名
                strPwd = DeCipher(rsTmp!密码 & "")
                strPort = rsTmp!端口 & ""
                strCollectType = rsTmp!收集类型 & ""
            End If
        Else
            Err.Clear
            blnOldStype = True
        End If
    Else
        strSQL = "Select 升级文件服务器 From ZLTools.zlClients Where 工作站=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "获取升级服务器编号", gstrComputerName)
        If Err.Number = 0 Then
            If Not rsTmp.EOF Then strServerID = rsTmp!升级文件服务器 & ""
        Else
            Err.Clear
            blnOldStype = True
        End If
        If strServerID <> "" Then
            strSQL = "Select 编号,类型, 位置, 用户名, 密码, 端口,Nvl(是否缺省,0) 是否缺省 , 批次 From Zltools.Zlupgradeserver Where 编号 = [1]"
            Set rsTmp = OpenSQLRecord(strSQL, "获取升级服务器", Val(strServerID))
            If Not rsTmp.EOF Then
                strServerID = rsTmp!编号 & ""
                sctConnType = IIf(rsTmp!类型 = 0, SCT_Share, SCT_FTP)
                strServer = rsTmp!位置
                strUser = rsTmp!用户名
                strPwd = DeCipher(rsTmp!密码 & "")
                strPort = rsTmp!端口 & ""
                glngFileBatch = Val(rsTmp!批次 & "")
                blnDefalut = rsTmp!是否缺省 = 1
            Else
                strServerID = ""
            End If
        End If
    End If
    If blnOldStype Then
        Set GetFileConnect = GetFileConnectOld(strMsg)
    Else
        If strServerID <> "" Then
            gobjTrace.WriteInfo "文件服务器", "服务器文件批次", glngFileBatch, "服务器编号", strServerID, "是否默认", blnDefalut
            blnConnOK = objConn.ToConnect(sctConnType, strServer, strUser, strPwd, strPort, strCollectType, strMsg)
        End If
        '连接不成功，升级服务器自动连接默认服务器
        If Not blnConnOK And gotCurType <> OT_Other And Not blnDefalut Then
            strSQL = "Select 编号,类型, 位置, 用户名, 密码, 端口, 批次 From Zltools.Zlupgradeserver Where Nvl(是否缺省,0) = 1"
            Set rsTmp = OpenSQLRecord(strSQL, "获取默认升级服务器")
            If Err.Number = 0 Then
                If Not rsTmp.EOF Then
                    strServerID = rsTmp!编号 & ""
                    sctConnType = IIf(rsTmp!类型 = 0, SCT_Share, SCT_FTP)
                    strServer = rsTmp!位置
                    strUser = rsTmp!用户名
                    strPwd = DeCipher(rsTmp!密码 & "")
                    strPort = rsTmp!端口 & ""
                    glngFileBatch = Val(rsTmp!批次 & "")
                    gobjTrace.WriteInfo "默认服务器", "服务器文件批次", glngFileBatch, "服务器编号", strServerID
                    blnConnOK = objConn.ToConnect(sctConnType, strServer, strUser, strPwd, strPort)
                End If
            Else
                Err.Clear
            End If
        End If
        If blnConnOK Then Set GetFileConnect = objConn
    End If
    Exit Function
ErrH:
    strMsg = Err.Description
End Function

Private Function GetFileConnectOld(ByRef strMsg As String) As clsConnect
'功能：获取文件服务器连接，老方式
'参数：blnUpgrade=True-预升级与升级的连接 ，false-文件收集的连接
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim sctConnType As ServerConnectType, strServerID As String
    Dim objConn As New clsConnect
    Dim arrParas() As Variant, arrValues(4) As String
    Dim strSQLPars As String, i As Integer
    Dim blnReadOk As Boolean, blnConnOK As Boolean, blnGo As Boolean
    
    On Error GoTo ErrH
    '获取连接类型
    sctConnType = SCT_Share
    strSQL = "Select 项目,内容 From ZLTools.zlregInfo where 项目=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, "升级类型", IIf(gotCurType <> OT_Other, "升级类型", "收集方式"))
    If Not rsTmp.EOF Then
        If NVL(rsTmp!内容, 0) = 1 Then sctConnType = SCT_FTP
    End If
    If gotCurType = OT_Other Then
        '文件收集的虚拟ID
        strServerID = IIf(sctConnType = SCT_FTP, "F", "S")
    Else
        '获取服务器ID
        strSQL = "Select 升级服务器,FTP服务器 From ZLTools.zlClients Where 工作站=[1]"
        Set rsTmp = OpenSQLRecord(strSQL, "获取升级服务器编号", gstrComputerName)
        If Not rsTmp.EOF Then strServerID = IIf(sctConnType = SCT_FTP, rsTmp!FTP服务器 & "", rsTmp!升级服务器 & "")
    End If
    '获取参数主信息
    If gotCurType <> OT_Other Then
        If sctConnType = SCT_FTP Then
            arrParas = Array("FTP服务器", "FTP用户", "FTP密码", "FTP端口", "")
        Else
            arrParas = Array("服务器目录", "访问用户", "访问密码", "", "")
        End If
    Else
        arrParas = Array("收集目录", "访问用户", "访问密码", "访问端口", "收集类型")
    End If
ReGetParas:
    '先获取SQL参数
    strSQLPars = ""
    For i = LBound(arrParas) To UBound(arrParas)
        If arrParas(i) <> "" Then
            strSQLPars = strSQLPars & ",'" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
        End If
    Next
    strSQLPars = Mid(strSQLPars, 2)
    strSQL = "Select 项目,内容 From ZLTools.zlregInfo where 项目 in(" & strSQLPars & ")"
    Set rsTmp = OpenSQLRecord(strSQL, "获取服务器")
    If Not rsTmp.EOF Then
        For i = LBound(arrParas) To UBound(arrParas)
            If arrParas(i) <> "" Then
                rsTmp.Filter = "项目='" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
                If Not rsTmp.EOF Then arrValues(i) = rsTmp!内容 & ""
            End If
        Next
    End If
    
    blnReadOk = True
    '服务器，用户，密码为空，则不能进行收集或升级
    If arrValues(0) = "" Or arrValues(1) = "" Or arrValues(2) = "" Then
        blnReadOk = False
    'FTP方式需要一个端口
    ElseIf sctConnType = SCT_FTP And arrValues(3) = "" Then
        blnReadOk = False
    '收集时，收集类型不能为空
    ElseIf gotCurType = OT_Other And arrValues(4) = "" Then
        blnReadOk = False
    End If
    If blnReadOk Then
        gobjTrace.WriteInfo "GetFileConnectOld", "旧方式服务器编号", strServerID
        blnConnOK = objConn.ToConnect(sctConnType, arrValues(0), arrValues(1), arrValues(2), arrValues(3), arrValues(4), strMsg)
    End If
    If (Not blnConnOK Or Not blnReadOk) And gotCurType <> OT_Other Then
        If strServerID <> "" And strServerID <> "0" Then
            strServerID = "0"
            GoTo ReGetParas '重新获取连接服务器的参数
        ElseIf (strServerID = "0" Or strServerID = "") And Not blnGo Then
            blnGo = True '防止循环
            strServerID = IIf(strServerID = "0", "", "0")
            GoTo ReGetParas '重新获取连接服务器的参数
        End If
    End If
    If blnConnOK Then Set GetFileConnectOld = objConn
    Exit Function
ErrH:
    strMsg = Err.Description
End Function

Public Function CheckAndAdjustFolder() As Boolean
'功能：进行安装路径的修复
    Dim strSQL              As String, rsTmp        As ADODB.Recordset
    Dim strPath             As String, arrTmp       As Variant
    Dim i                   As Integer
    Dim strErrInfo          As String
    
    Err.Clear: On Error GoTo ErrH
    strSQL = "Select Distinct Upper(安装路径) 安装路径 From Zlfilesupgrade"
    Set rsTmp = OpenSQLRecord(strSQL, "获取路径文件夹")
    
    Do While Not rsTmp.EOF
        arrTmp = Split(rsTmp!安装路径 & "", "\")
        strPath = ""
        If UBound(arrTmp) <> -1 Then
            arrTmp(0) = Trim(arrTmp(0))
            If arrTmp(0) = "[APPSOFT]" Then
                strPath = gstrSetupPath
            ElseIf arrTmp(0) = "[PUBLIC]" Then
                If Not gobjFSO.FolderExists(gstrSetupPath & "\PUBLIC") Then
                    gobjFSO.CreateFolder (gstrSetupPath & "\PUBLIC")
                End If
                strPath = gstrSetupPath & "\PUBLIC"
            ElseIf arrTmp(0) = "[APPLY]" Then
                strPath = gstrSetupPath & "\APPLY"
            ElseIf arrTmp(0) = "[OS:]" Then '系统盘
                strPath = Left(gstrSystemPath, 2)
            ElseIf arrTmp(0) = "[APP:]" Then  '当前安装盘
                strPath = Left(gstrSetupPath, 2)
            End If
            If strPath <> "" Then
                For i = 1 To UBound(arrTmp)
                    If arrTmp(i) <> "" Then
                        strPath = strPath & "\" & arrTmp(i)
                        If Not gobjFSO.FolderExists(strPath) Then
                            gobjFSO.CreateFolder (strPath)
                        End If
                    End If
                Next
                '缓存安装路径，优化转换速度。
                gcllSetPath.Add strPath, "K_" & rsTmp!安装路径
            End If
        End If
        rsTmp.MoveNext
    Loop
    '缓存基础安装路径，优化转换速度。
    On Error Resume Next
    gcllSetPath.Add gstrSetupPath, "K_[APPSOFT]"
    gcllSetPath.Add gstrSetupPath & "\PUBLIC", "K_[PUBLIC]"
    gcllSetPath.Add gstrSetupPath & "\APPLY", "K_[APPLY]"
    gcllSetPath.Add Left(gstrSystemPath, 2), "K_[OS:]"
    gcllSetPath.Add Left(gstrSetupPath, 2), "K_[APP:]"
    gcllSetPath.Add gstrSystemPath, "K_[SYSTEM]"
    gcllSetPath.Add gobjFSO.GetParentFolderName(gstrSystemPath) & "\Help", "K_[HELP]"
    gcllSetPath.Add gstrSetupPath & "\APPLY", "K_[APPSOFT]\APPLY"
    If Err.Number Then Err.Clear
    On Error Resume Next
    '缓存弃用文件路径
    strSQL = "Select distinct upper(安装路径) 安装路径 From zlFilesExpired"
    Set rsTmp = OpenSQLRecord(strSQL, "获取路径文件夹")
    If Not rsTmp Is Nothing Then
        Err.Clear
        Do While Not rsTmp.EOF
            strPath = gcllSetPath("K_" & rsTmp!安装路径)
            If Err.Number <> 0 Then
                Err.Clear
                arrTmp = Split(rsTmp!安装路径 & "", "\")
                strPath = ""
                If UBound(arrTmp) <> -1 Then
                    arrTmp(0) = Trim(arrTmp(0))
                    If arrTmp(0) = "[APPSOFT]" Then
                        strPath = gstrSetupPath
                    ElseIf arrTmp(0) = "[PUBLIC]" Then
                        If Not gobjFSO.FolderExists(gstrSetupPath & "\PUBLIC") Then
                            gobjFSO.CreateFolder (gstrSetupPath & "\PUBLIC")
                        End If
                        strPath = gstrSetupPath & "\PUBLIC"
                    ElseIf arrTmp(0) = "[APPLY]" Then
                        strPath = gstrSetupPath & "\APPLY"
                    ElseIf arrTmp(0) = "[OS:]" Then '系统盘
                        strPath = Left(gstrSystemPath, 2)
                    ElseIf arrTmp(0) = "[APP:]" Then '当前安装盘
                        strPath = Left(gstrSetupPath, 2)
                    End If
                    If strPath <> "" Then
                        For i = 1 To UBound(arrTmp)
                            If arrTmp(i) <> "" Then
                                strPath = strPath & "\" & arrTmp(i)
                                If Not gobjFSO.FolderExists(strPath) Then
                                    gobjFSO.CreateFolder (strPath)
                                End If
                            End If
                        Next
                        '缓存安装路径，优化转换速度。
                        gcllSetPath.Add strPath, "K_" & rsTmp!安装路径
                    End If
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    If Err.Number Then Err.Clear
    CheckAndAdjustFolder = True
    Exit Function
ErrH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "CheckAndAdjustFolder", "检查修复安装目录失败", strErrInfo
    Call RecordErrMsg(MT_InitEnv, "修复安装目录", strErrInfo)
    MsgBox "检查修复安装目录发生致命错误，请联系管理员！信息：" & vbNewLine & strErrInfo, vbInformation, App.Title
End Function

Public Function UpgradeBase() As Boolean
'功能：下载自动升级所需要的基础部件
    Dim strFile As String, blnAdmin As Boolean
    Dim strErr As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strReturn As String
    Dim strMsg As String
    Dim strCommand As String, strTmp As String
    Dim objText As TextStream, blnMust  As Boolean, blnErr  As Boolean
    
    gobjTrace.WriteSection "基础部件升级", SL_LevelTwo
    On Error GoTo ErrH
    strSQL = "Select 序号, 文件名, Upper(文件名) 标准文件名," & IIf(gblnHaveVersion, "文件版本号", " ") & " 版本号, 修改日期, 文件类型, 业务部件, Upper(安装路径) 安装路径, Md5, 自动注册, 强制覆盖" & vbNewLine & _
            "From ZLTOOLS.Zlfilesupgrade" & vbNewLine & _
            "Where Upper(文件名) In ('ZLRUNAS.EXE','ZLHISCRUST.EXE','7Z.EXE','7Z.DLL','AAMD532.DLL','GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
    Set rsTmp = OpenSQLRecord(strSQL, App.Title)
    '1、优先下载ZLRUNAS.EXE获取管理员权限，由此可以下载MD5计算部件。计算ZlHISCrust部件的MD5
    On Error Resume Next
    strFile = gstrSetupPath & "\zlTestAdmin.txt"
    Call gobjFSO.CreateTextFile(strFile, True)
    Call gobjFSO.CopyFile(strFile, gstrSystemPath & "\zlTestAdmin.txt", True)
    If Err.Number = 75 Then
        blnAdmin = False
    ElseIf Dir(gstrSystemPath & "\zlTestAdmin.txt", vbNormal) <> "" Then
        blnAdmin = True
        Call gobjFSO.DeleteFile(gstrSystemPath & "\zlTestAdmin.txt", True)
    Else
        blnAdmin = False
    End If
    Call gobjFSO.DeleteFile(strFile, True)
    If Err.Number <> 0 Then Err.Clear
    gobjTrace.WriteInfo "UpgradeBase", "System目录写入权限", blnAdmin
    If Not blnAdmin Then
        rsTmp.Filter = "标准文件名='ZLRUNAS.EXE'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!安装路径, Val(rsTmp!文件类型 & ""), rsTmp!文件名)
            If Not gobjFSO.FileExists(strFile) Then
                gobjTrace.WriteInfo "UpgradeBase", "升级基础文件", rsTmp!文件名
                If gclsConnect.IsServerFileExists(rsTmp!标准文件名) Then
                    If Not gclsConnect.DownloadFile(rsTmp!标准文件名, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "服务器文件文件下载失败(ZLRUNAS.EXE(USER权限执行工具))" & strErr
                    End If
                Else
                    strMsg = "服务器文件缺失ZLRUNAS.EXE(USER权限执行工具)"
                End If
            End If
            If gobjFSO.FileExists(strFile) Then
                '先保存命令行，待下次启动使用
                If gobjFSO.FileExists(gstrSetupPath & "\ZLRUNAS.ini") Then
                    gobjFSO.DeleteFile gstrSetupPath & "\ZLRUNAS.ini", True
                End If
                Set objText = gobjFSO.CreateTextFile(gstrSetupPath & "\ZLRUNAS.ini")
                objText.WriteLine Cipher(gstrCommand)
                objText.Close
                Set objText = Nothing
                strMsg = StartZLRunAs(strFile)
            End If
        Else
            strMsg = "服务器目录(ZLfilesUpgrade)中缺失ZLRUNAS.EXE(USER权限执行工具)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "管理员运行工具检查", strMsg
            Call RecordErrMsg(MT_InitEnv, "管理员运行工具检查", strMsg)
            MsgBox strMsg & vbNewLine & "，请联系管理员！", vbInformation, App.Title
            Exit Function
        End If
    End If
    '2、下载AAMD532.dll该部件是用来计算MD5,必须优先ZLHISCrust.exe，否则无法检查ZLHISCrust.exe是否需要升级。
    strMsg = ""
    rsTmp.Filter = "标准文件名='AAMD532.DLL'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!安装路径 & "", Val(rsTmp!文件类型 & ""), rsTmp!文件名)
        If Not gobjFSO.FileExists(strFile) Then
            If gclsConnect.IsServerFileExists(rsTmp!标准文件名) Then
                gobjTrace.WriteInfo "UpgradeBase", "升级基础文件", rsTmp!文件名
                If Not gclsConnect.DownloadFile(rsTmp!标准文件名, gobjFSO.GetParentFolderName(strFile), strErr) Then
                    strMsg = "服务器文件文件下载失败AAMD532.DLL(MD5计算工具)" & strErr
                End If
            Else
                strMsg = "服务器文件缺失AAMD532.DLL(MD5计算工具)"
            End If
        End If
    Else
        strMsg = "服务器目录(ZLfilesUpgrade)中缺失AAMD532.DLL(MD5计算工具)"
    End If
    
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "MD5计算工具检查", strMsg
        Call RecordErrMsg(MT_InitEnv, "MD5计算工具检查", strMsg)
        MsgBox strMsg & vbNewLine & "，请联系管理员！", vbInformation, App.Title
        Exit Function
    End If
    strMsg = ""
    '3、下载ZLHISCrust.exe，该部件可以进行检查升级了
    If Val(GetSetting("ZLSOFT", "公共模块\自动升级", "工具调试", "0")) = 0 Then
        If gintCallTimes = 0 Then '第二次调用升级工具进行升级。不算ZLRUNAS调用的那一次
            rsTmp.Filter = "标准文件名='ZLHISCRUST.EXE'"
            If Not rsTmp.EOF Then
                strFile = GetActualPath(rsTmp!安装路径 & "", Val(rsTmp!文件类型 & ""), rsTmp!文件名)
                If IsFileUpgade(App.Path & "\ZLHISCRUST.EXE", rsTmp!版本号 & "", rsTmp!修改日期 & "", rsTmp!MD5 & "") Then
                    If gclsConnect.IsServerFileExists(rsTmp!标准文件名) Then
                        gobjTrace.WriteInfo "UpgradeBase", "升级基础文件", rsTmp!文件名
                        If Not gclsConnect.DownloadFile(rsTmp!标准文件名, gstrTempPath, strErr) Then
                            strMsg = "服务器文件文件下载失败:ZLHISCRUST.EXE(自动升级主程序)" & strErr
                        Else
                            '文件又变成老部件，则讲文件移动到APPSOft\APPLY下
                            strTmp = UCase(GetVersionInfo(gstrTempPath & "\" & rsTmp!文件名, FVN_ProductName))
                            If strTmp = "" Then strTmp = "ZLHISINSTALLUPDATE"
                            If strTmp <> "ZLHISINSTALLUPDATE" Then 'zlHisInstallUpdate
                                strFile = gstrSetupPath & "\Apply\" & rsTmp!文件名
                                If gobjFSO.FileExists(strFile) Then
                                    If FileSystem.GetAttr(strFile) <> vbNormal Then
                                         Call FileSystem.SetAttr(strFile, vbNormal)
                                    End If
                                    Call gobjFSO.DeleteFile(strFile)
                                End If
                                gobjFSO.CopyFile gstrTempPath & "\" & rsTmp!文件名, strFile, False
                                strCommand = GetHisUpdateCommand(True)
                            Else
                                strFile = gstrTempPath & "\" & rsTmp!文件名
                                strCommand = GetHisUpdateCommand()
                            End If
                            '下载后需要使用新的ZLHISCRUST.EXE来进行升级
                            On Error Resume Next
                            Call gobjTrace.CloseLog
                            If Shell(strFile & " " & strCommand, vbNormalFocus) <> 0 Then
                                Call gclsConnect.CloseConnect
                                End
                            Else
                            End If
                        End If
                    Else
                        strMsg = "服务器文件缺失ZLHISCRUST.EXE(自动升级主程序)"
                    End If
                End If
            Else
                strMsg = "服务器目录(ZLfilesUpgrade)中缺失ZLHISCRUST.EXE(自动升级主程序)"
            End If
        End If
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "自动升级工具检查", strMsg
        Call RecordErrMsg(MT_InitEnv, "自动升级工具检查", strMsg)
        MsgBox strMsg & vbNewLine & "，请联系管理员！", vbInformation, App.Title
        Exit Function
    End If
    strMsg = ""
    '4、下载压缩工具，以便其他常规升级的解压
    rsTmp.Filter = "标准文件名='7Z.DLL'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!安装路径 & "", Val(rsTmp!文件类型 & ""), rsTmp!文件名)
        If Not gobjFSO.FileExists(strFile) Then
            If gclsConnect.IsServerFileExists(rsTmp!标准文件名) Then
                gobjTrace.WriteInfo "UpgradeBase", "升级基础文件", rsTmp!文件名
                If Not gclsConnect.DownloadFile(rsTmp!标准文件名, gobjFSO.GetParentFolderName(strFile), strErr) Then
                    strMsg = "服务器文件文件下载失败(7Z.DLL(解压工具依赖部件))" & strErr
                End If
            Else
                strMsg = "服务器文件缺失7Z.DLL(解压工具依赖部件)"
            End If
        End If
    Else
        strMsg = "服务器目录(ZLfilesUpgrade)中缺失7Z.DLL(解压工具依赖部件)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "解压工具检查", "信息", strMsg
        Call RecordErrMsg(MT_InitEnv, "自动升级工具检查", strMsg)
        MsgBox strMsg & vbNewLine & "，请联系管理员！", vbInformation, App.Title
        Exit Function
    End If
    strMsg = ""
    '4、下载压缩工具，以便其他常规升级的解压
    rsTmp.Filter = "标准文件名='ZL7Z.DLL'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!安装路径 & "", Val(rsTmp!文件类型 & ""), rsTmp!文件名)
        If Not gobjFSO.FileExists(strFile) Then
            If gclsConnect.IsServerFileExists(rsTmp!标准文件名) Then
                gobjTrace.WriteInfo "UpgradeBase", "升级基础文件", rsTmp!文件名
                If Not gclsConnect.DownloadFile(rsTmp!标准文件名, gobjFSO.GetParentFolderName(strFile), strErr) Then
                    strMsg = "服务器文件文件下载失败(ZL7Z.DLL(中联压缩部件))" & strErr
                Else
                    strMsg = ""
                    If Not gclsRegCom.RegCom(strFile, RFT_NormalReg, strMsg) Then
                        gobjTrace.WriteInfo "UpgradeBase", "ZL7Z注册失败", strMsg
                        Call RecordErrMsg(MT_InitEnv, "ZL7Z注册失败", strMsg)
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "ZL7Z注册成功", ""
                    End If
                    strMsg = ""
                End If
            Else
                strMsg = "服务器文件缺失ZL7Z.DLL(中联压缩部件)"
            End If
        End If
    Else
        strMsg = "服务器目录(ZLfilesUpgrade)中缺失ZL7Z.DLL(中联压缩部件)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "解压工具检查", "信息", strMsg
        Call RecordErrMsg(MT_InitEnv, "自动升级工具检查", strMsg)
        MsgBox strMsg & vbNewLine & "，请联系管理员！", vbInformation, App.Title
        Exit Function
    End If
    strMsg = ""
    rsTmp.Filter = "标准文件名='7Z.EXE'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!安装路径 & "", Val(rsTmp!文件类型 & ""), rsTmp!文件名)
        gstr7ZPath = strFile
        If Not gobjFSO.FileExists(strFile) Then
            If gclsConnect.IsServerFileExists(rsTmp!标准文件名) Then
                gobjTrace.WriteInfo "UpgradeBase", "升级基础文件", rsTmp!文件名
                If Not gclsConnect.DownloadFile(rsTmp!标准文件名, gobjFSO.GetParentFolderName(strFile), strErr) Then
                    strMsg = "服务器文件文件下载失败(7Z.EXE(解压工具))" & strErr
                End If
            Else
                strMsg = "服务器文件缺失7Z.EXE(解压工具)"
            End If
        End If
    Else
        strMsg = "服务器目录(ZLfilesUpgrade)中缺失7Z.EXE(解压工具)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "解压工具检查", strMsg
        Call RecordErrMsg(MT_InitEnv, "自动升级工具检查", strMsg)
        MsgBox strMsg & vbNewLine & "，请联系管理员！", vbInformation, App.Title
        Exit Function
    End If
    '5、下载
    strMsg = ""
    blnMust = IsMustGACUTIL(): blnErr = False
    rsTmp.Filter = "标准文件名='GACUTIL.EXE'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!安装路径 & "", Val(rsTmp!文件类型 & ""), rsTmp!文件名)
        gstrGACPath = strFile
        If Not gobjFSO.FileExists(strFile) Then
            If gclsConnect.IsServerFileExists(rsTmp!标准文件名) Then
                gobjTrace.WriteInfo "UpgradeBase", "升级基础文件", rsTmp!文件名
                If Not gclsConnect.DownloadFile(rsTmp!标准文件名, gobjFSO.GetParentFolderName(strFile), strErr) Then
                    strMsg = "服务器文件文件下载失败(GACUTIL.EXE(全局缓存添加工具))" & strErr
                    blnErr = True
                End If
            Else
                strMsg = "服务器文件缺失GACUTIL.EXE(全局缓存添加工具)"
            End If
        End If
    Else
        strMsg = "服务器目录(ZLfilesUpgrade)中缺失GACUTIL.EXE(全局缓存添加工具)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "全局缓存添加工具检查", strMsg
        If blnMust Or blnErr Then
            Call RecordErrMsg(MT_InitEnv, "全局缓存添加工具检查", strMsg)
            MsgBox strMsg & vbNewLine & ",请联系管理员！", vbInformation, App.Title
            Exit Function
        End If
    End If
    strMsg = ""
    blnErr = False
    rsTmp.Filter = "标准文件名='GACUTIL.EXE.CONFIG'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!安装路径 & "", Val(rsTmp!文件类型 & ""), rsTmp!文件名)
        gstrGACPath = strFile
        If Not gobjFSO.FileExists(strFile) Then
            If gclsConnect.IsServerFileExists(rsTmp!标准文件名) Then
                gobjTrace.WriteInfo "UpgradeBase", "升级基础文件", rsTmp!文件名
                If Not gclsConnect.DownloadFile(rsTmp!标准文件名, gobjFSO.GetParentFolderName(strFile), strErr) Then
                    strMsg = "服务器文件文件下载失败(GACUTIL.EXE.CONFIG(全局缓存添加工具配置文件))" & strErr
                    blnErr = True
                End If
            Else
                strMsg = "服务器文件缺失GACUTIL.EXE.CONFIG(全局缓存添加工具配置文件)"
            End If
        End If
    Else
        strMsg = "服务器目录(ZLfilesUpgrade)中缺失GACUTIL.EXE.CONFIG(全局缓存添加工具配置文件)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "全局缓存添加工具检查", strMsg
        If blnMust Or blnErr Then
            Call RecordErrMsg(MT_InitEnv, "全局缓存添加工具检查", strMsg)
            MsgBox strMsg & vbNewLine & ",请联系管理员！", vbInformation, App.Title
            Exit Function
        End If
    End If
    If Not gobj7zZip.Init7zZip Then
        gobjTrace.WriteInfo "UpgradeBase", "7zZip初始化", "无法创建ZL7z部件且没有7z.exe"
        Call RecordErrMsg(MT_InitEnv, "自动升级工具检查", "无法创建ZL7z部件且没有7z.exe")
        MsgBox "无法创建ZL7z部件且没有7z.exe" & vbNewLine & "，请联系管理员！", vbInformation, App.Title
        Exit Function
    End If
    UpgradeBase = True
    Exit Function
ErrH:
    gobjTrace.WriteInfo "UpgradeBase", "升级基础部件发生致命错误", Err.Description
    Call RecordErrMsg(MT_InitEnv, "升级基础部件发生致命错误", Err.Description)
    MsgBox "升级基础部件发生致命错误" & vbNewLine & "，请联系管理员！信息：" & Err.Description, vbInformation, App.Title
    Err.Clear
End Function

Private Function StartZLRunAs(ByVal strPath As String) As String
'功能：启动ZLRunas
    Dim strSQL          As String, rsTmp    As ADODB.Recordset
    Dim strUser         As String, strPwd   As String
    Dim strCommandPara  As String, strMsg   As String, strReturn As String
    Dim blnOk           As Boolean
    
    On Error Resume Next
    strSQL = "Select Max(管理员用户) 管理员, Max(管理员密码)  密码 From ZLTOOLS.zlClients Where 工作站 = [1]"
    Set rsTmp = OpenSQLRecord(strSQL, "获取当前客户端登录许可")
    '兼容模式，低版本没有这两个字段
    If Err.Number = 0 Then
        strUser = NVL(rsTmp!管理员, "Administrator")
        strPwd = Trim(rsTmp!密码 & "")
    Else
        Err.Clear
    End If
    On Error GoTo ErrH
    '密码解密
    If strPwd <> "" And strUser <> "" Then
        strPwd = DeCipher(strPwd)
        strCommandPara = "-u " & strUser & " -p " & strPwd  '用于ZLRunas.EXE命令行
        gobjTrace.WriteInfo "StartZLRunAs", "客户端管理许可", Cipher(strCommandPara)
        '重新启动升级外壳
        strReturn = RunCommand(strPath & " " & strCommandPara & " -ex """ & App.Path & "\ZLHISCRUST.EXE"" -lwp", , True, 60000)
        If InStr(strReturn, (1326)) > 0 Then
            strMsg = "登录失败: 未知的用户名或错误密码。"
        ElseIf InStr(strReturn, (1058)) > 0 Then
            strMsg = "无法启动服务，原因可能是SecLogon服务被禁用。"
        ElseIf InStr(strReturn, (1717)) > 0 Then
            strMsg = "'路径中不能有中文，否则执行不成功"
        Else
            blnOk = True
        End If
    Else
        gobjTrace.WriteInfo "StartZLRunAs", "客户端管理许可", "没有统一管理设置"
    End If
    '使用每个客户端的个人设置
    If Not blnOk Then
        strSQL = "Select Max(Decode(项目, '管理员账号', 内容, '')) As 管理员, Max(Decode(项目, '管理员密码', 内容, '')) As 密码" & vbNewLine & _
                "From Zltools.Zlreginfo" & vbNewLine & _
                "Where 项目 = '管理员账号' Or 项目 = '管理员密码'"
        Set rsTmp = OpenSQLRecord(strSQL, "获取统一许可")
        strUser = NVL(rsTmp!管理员, "Administrator")
        strPwd = Trim(rsTmp!密码 & "")
        If strPwd <> "" And strUser <> "" Then
            strPwd = DeCipher(strPwd)
            strCommandPara = "-u " & strUser & " -p " & strPwd  '用于ZLRunas.EXE命令行
            gobjTrace.WriteInfo "StartZLRunAs", "当前客户端登录许可", Cipher(strCommandPara)
            '重新启动升级外壳
            strReturn = RunCommand(strPath & " " & strCommandPara & " -ex """ & App.Path & "\ZLHISCRUST.EXE"" -lwp", , True, 60000)
            If InStr(strReturn, (1326)) > 0 Then
                strMsg = "登录失败: 未知的用户名或错误密码。"
            ElseIf InStr(strReturn, (1058)) > 0 Then
                strMsg = "无法启动服务，原因可能是SecLogon服务被禁用。"
            ElseIf InStr(strReturn, (1717)) > 0 Then
                strMsg = "'路径中不能有中文，否则执行不成功"
            Else
                blnOk = True
            End If
        Else
            gobjTrace.WriteInfo "StartZLRunAs", "当前客户端登录许可", "没有登录许可设置"
        End If
    End If
    StartZLRunAs = strMsg
    Exit Function
ErrH:
    gobjTrace.WriteInfo "StartZLRunAs", "获取客户端许可发生致命错误", Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetUpgradeFileList() As Boolean
'功能：获取ZLFIleUpgrade
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String, strMsg As String
    
    On Error GoTo ErrH
    '检查同名文件
    strSQL = "Select Upper(a.文件名) 文件名 From Zlfilesupgrade a Group By Upper(a.文件名) Having Count(1) > 1"
    Set rsTmp = OpenSQLRecord(strSQL, "获取文件清单")
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!文件名
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then
        strMsg = "存在同名(大小写区别)部件：" & Mid(Mid(strTmp, 2), 1, 100)
        gobjTrace.WriteInfo "GetUpgradeFileList", "部件清单合法性检查", strMsg
        Call RecordErrMsg(MT_InitEnv, "部件清单合法性检查", strMsg)
        MsgBox "部件清单存在问题，请联系管理员进行处理。" & vbNewLine & strMsg, vbInformation + vbDefaultButton1, App.Title
        Exit Function
    End If
    On Error Resume Next
    strSQL = "Select a.文件名, Upper(a.文件名) 标准文件名," & IIf(gblnHaveVersion, "a.文件版本号 ", " a.") & "版本号, a.修改日期, a.文件类型, a.业务部件, a.安装路径, a.Md5, NVL(a.自动注册,0) 自动注册, NVL(a.强制覆盖,0) 强制覆盖,附加安装路径" & vbNewLine & _
            "From Zltools.Zlfilesupgrade a" & vbNewLine & _
            "Where Upper(a.文件名) Not In ('ZLRUNAS.EXE', 'ZLHISCRUST.EXE', '7Z.EXE', '7Z.DLL', 'AAMD532.DLL', 'GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
    Set rsTmp = OpenSQLRecord(strSQL, "获取文件清单")
    If Err.Number <> 0 Then
        Err.Clear
        strSQL = "Select a.文件名, Upper(a.文件名) 标准文件名, " & IIf(gblnHaveVersion, "a.文件版本号 ", " a.") & "版本号, a.修改日期, a.文件类型, a.业务部件, a.安装路径, a.Md5, NVL(a.自动注册,0) 自动注册, NVL(a.强制覆盖,0) 强制覆盖,Null 附加安装路径" & vbNewLine & _
                "From Zltools.Zlfilesupgrade a" & vbNewLine & _
                "Where Upper(a.文件名) Not In ('ZLRUNAS.EXE', 'ZLHISCRUST.EXE', '7Z.EXE', '7Z.DLL', 'AAMD532.DLL', 'GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
        Set rsTmp = OpenSQLRecord(strSQL, "获取文件清单")
    End If
    '实际路径-安装路径转换为实际路径
    '清理文件路径-错误路径文件
    Set grsFileUpgrade = CopyNewRec(rsTmp, , , Array("更新", adInteger, 1, 0, "实际路径", adVarChar, 500, Empty, "清理文件路径", adVarChar, 1000, Empty, "附加实际路径", adVarChar, 4000, Empty, _
                                                "判断批次", adInteger, 3, 0, "预升级下载", adInteger, 1, 0, "错误信息", adVarChar, 1000, Empty, "检查信息", adVarChar, 1000, Empty, _
                                                "无后缀文件名", adVarChar, 100, Empty, "类型排序", adInteger, 1, 0, "注册错误", adInteger, 1, 0))
    GetUpgradeFileList = True
    Exit Function
ErrH:
    gobjTrace.WriteInfo "GetUpgradeFileList", "部件清单获取失败", Err.Description
    Call RecordErrMsg(MT_InitEnv, "文件清单获取", Err.Description)
    MsgBox "部件清单获取失败，" & vbNewLine & "请联系管理员！信息：" & Err.Description, vbInformation, App.Title
End Function

Public Function GetKILLProcess() As Boolean
'功能：获取要杀掉的进程
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strTmp As String

    On Error Resume Next
    strSQL = "Select 序号, 名称,类型 From Zltools.ZlkillProcess Order By 序号"
    Set rsTmp = OpenSQLRecord(strSQL, "获取文件清单")
    If rsTmp Is Nothing Then
        If Err.Number <> 0 Then Err.Clear
    Else
        Do While Not rsTmp.EOF
            strTmp = strTmp & ";" & UCase(rsTmp!名称)
            rsTmp.MoveNext
        Loop
    End If
    
    If strTmp = "" Then
        strTmp = "zl9LabPrintSvr.exe;zl9LabReceiv.exe;zl9LabTcpSvr.exe;Zl9LISComm.exe;zl9PacsCapture.exe;zl9WizardMain.exe;zl9WizardStart.exe;ZL9Xls.exe;zlActMain.exe;ZLBAExport.exe;zlCDOpen.exe;zlCisAuditPrint.exe;zlDrugMachineManage.exe;zlGetImage.exe;zlGetImageEx.exe;zlHQMSDCollect.exe;zlLisReceiveSend.exe;zlMipClientManage.exe;zlMipClientPoll.exe;zlMipClientShell.exe;zlMsgBuilderStart.exe;zlMsgReceiver.exe;zlMsgSender.exe;ZLNewQuery.exe;zlOrclConfig.exe;ZLPacsBrowserStation.exe;ZlPacsSrv.exe;zlPeisAutoAnalyse.exe;zlQueueShow.exe;ZLRPTSQLAdjust.exe;ZLRUNAS.EXE;zlScreenKeyboard.exe;zlSoftShowArchive.exe;zlSvrNotice.exe;zlSvrStudio.exe;zlUpgradeReader.exe;zlWizardStart.exe;ZLPacsServerCenter.exe"
    Else
        strTmp = Mid(2, strTmp)
    End If
    gobjTrace.WriteInfo "GetKILLProcess", "进程清单", strTmp
    garrKillProcess = Split(UCase(strTmp), ";")
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function IsMustGACUTIL() As Boolean
'功能：是否必须要GACUTIL.EXE与GACUTIL.EXE.CONFIG
    Dim strSQL As String, rsTmp As ADODB.Recordset

    On Error GoTo ErrH
    strSQL = "Select Count(1) 计数 From Zlfilesupgrade a Where a.自动注册 = [1] And a.Md5 Is Not Null"
    Set rsTmp = OpenSQLRecord(strSQL, "获取文件清单", RFT_NETGAC)
    IsMustGACUTIL = rsTmp!计数 > 0
ErrH:
    gobjTrace.WriteInfo "IsMustGACUTIL", "获取GACUTIL注册部件", Err.Description
    Err.Clear
End Function

