VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@模块 clsJob-2019/7/2
'@编写 lshuo
'@功能
'   任务处理类
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------

Private Const M_SHARE_SVR_SEND              As String = "CA4747E9-6A10-4B19-8B95-F2EC3DAFF6C9"              '升级助手发送消息交互的内存共享
Private Const M_SHARE_SVR_RECEIVE           As String = "3892908F-5A80-484C-A031-FA95647E8EBE"              '升级助手接收消息交互的内存共享
Private Const M_HELPERMAIN_STARTJOB         As String = "6F5628B4-69EB-497D-84E7-8D47F1E711F5"              '是否是升级助手启动的任务
Private Const M_LNG_TIMEOUT_JOBRUN          As Long = 600000                        '运行状态的超时限制
Private Const M_LNG_TIMEOUT_WAITCONNECT     As Long = 30000                         '等待连接状态超时限制
Private Const M_LNG_TIMEOUT_CONNECTED       As Long = 20000                         '已经连接状态超时限制
Private Const M_LNG_TIMEOUT_SENDJOB         As Long = 20000                         '已经发送模块超时限制
Private Const M_LNG_TIMEOUT_STARTJOB        As Long = 25000                         '启动模块超时限制

'功能类型
Public Enum JobType
    JT_NoneJob = 0                  '没有任务或者尚未进行检测
    JT_Repaire = 1                  '客户端修复。必须在定时升级时间之后才能执行。修复是客户端主动处理，不通过升级助手中心处理，此处只是预留了框架，后续方便扩展
    JT_ClientPreUpgrade = 2         '客户端预升级，必须在定时升级之前才能执行。
    JT_ClientUpgrade = 3            '客户端正式升级，必须在定时升级时间之后才能执行
    JT_ComponentCheck = 4           '客户端部件检查，必须在没有其他任务的情况下执行。可以考虑和功能验证合并。
    JT_VertifyFuncs = 5             '客户端功能验证，必须当前客户端不在升级状态，才能执行。
End Enum
'数据类型枚举
Private Enum DataStrucType
    DST_VerAuto = 0                 '自动判断处理
    DST_VerNAE40 = 1                '35.40以下的数据结构
    DST_VerNB40 = 2                 '3540-10.35.100的数据结构
    DST_VerNB120 = 3                '>=10.35.120的数据结构
End Enum
'服务器状态控制
Public Enum ServerControl
    SC_Continue = 0
    SC_Finish = 1
    SC_Delay = 2
    SC_Wait = 3
End Enum
'任务异常枚举
Public Enum JobException
    JE_NoneException = 0            '无异常
    JE_MultiException = 1           '同样异常多次出现导致的中止
    JE_ConnectException = 2         '无法连接的异常。
    JE_StateSaveException = 3       '整体任务保存失败异常。
    JE_VertifyFuncUserException = 4 '功能验证没有对应用户异常
    JE_ExeTimeOutNoNameException = 5 '整体执行超时时，没有文件名异常
    JE_ExeTimeOutNoPIDException = 6 '整体执行超时，但是没有找到进程
    JE_ExeTimeOutOldPIDException = 7 '整体任务执行超时异常，进程还是原来的进程。
    JE_ExeTimeOutNewPIDException = 8 '整体任务执行超时异常，进程是新的进程。
    JE_FileException = 9            '进程Exe文件异常。文件不存在。该种异常只需要将该进程对应的验证模块依次保存数据库即可。
    JE_StartUpException = 10        '进程Exe启动相关的异常。该种异常需要重试多次，不成功，则自动将该进程对应的验证模块依次保存数据库即可。
    JE_MsgTimeOutException = 11     '功能验证在发送模块，接收模块验证完毕过程中发生的超时异常
    JE_SubStateSaveException = 12   '子任务状态保存异常
End Enum
'当前的任务状态
Private Enum JobStatus
    JS_WaitExec = 0                 '等待执行
    JS_Successed = 1                '执行完成且成功
    JS_Failed = 2                   '执行完成但是失败
    JS_Running = 3                  '正在执行
End Enum
'进程交互状态
Private Enum ProcessInteractiveState
    PIS_NoneState = 0                '缺省无状态
    PIS_WaitConnect = 1              '已经启动进程，等待回复
    PIS_Connected = 2                '任务进程发送已经连接的消息
    PIS_SendJob = 3                  '主进程发送任务控制消息
    PIS_StartJob = 4                 '任务进程启动任务消息
    PIS_EndJob = 5                   '任务进程完成任务消息
End Enum
'功能验证模块类型
Private Enum VertifyModfyType
    VMT_Default = 0
    VMT_His = 1
    VMT_Lis = 2
    VMT_Peis = 3
End Enum
'功能验证统计信息
Private Enum VertifyTotalType
    VTT_All = 0
    VTT_Successed = 1
    VTT_Failed = 2
End Enum
'更新数据的SQL类型枚举
Private Enum SQLType
    ST_NotJudge = -1
    ST_Proceduer = 0
    ST_UpdateSQL = 1
End Enum
'模块验证与结果枚举
Private Enum ModuleResult
    MR_TypeOrResult = 0 '模块类型或回应是否异常。0-异常，1-正常
    MR_Sys = 1          '验证的系统
    MR_Module = 2       '验证的模块
    MR_ComOrMessge = 3  '发送模块到ZLHIS+时的部件名称信息。或返回本次执行模块的结果描述
    MR_TilTle = 4       '模块标题
End Enum
'清理原因
Private Enum ClearReason
    CR_JobFinish = 0            '任务完成
    CR_SubJobProcessChange = 1  '子任务进程变动
    CR_SubJobStart = 2          '子任务启动
    CR_SubJobFinish = 3         '子任务完成
End Enum

'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
'1.数据库信息
Private mstrUser                            As String                       '进程启动使用的用户
Private mstrProxyUser                       As String                       '代理使用的用户。在升级功能验证时，验证该用户的模块功能。
Private mstrPwd                             As String                       '进程启动使用的密码
'2.任务信息
Private mjtJobType                          As JobType                      '任务类型
Private mjsCurJobSatus                      As JobStatus                    '当前任务的状态
Private mstrExeName                         As String                       '当前任务的启动进程
Private mlngJobStart                        As Long                         '任务初次启动时间
Private mlngLastResponses                   As Long                         '上次回应时间。消息回应、发送。进程启动等操作中最近一次操作时间
Private mlngSubProcess                      As Long                         '业务进程ID
Private mlngCurProcess                      As Long                         '当前进程ID
Private mpisCurSubJob                       As ProcessInteractiveState      '当前子任务交互状态
'3.临时变量
Private mdblTimeSync                        As Double                       '本地时间和服务器时间差值
Private marrVertifyTotal(3, 2)              As Long                         '功能验证的统计信息
Private mdstCurDataStruc                    As DataStrucType                '数据结构类型。0-升级标志,预升完成,收集标志(10.35.40以下） 1-是否预升级,预升完成,升级标志,收集标志(10.35.40) 2-是否预升级,预升完成,升级标志,收集标志,功能验证状态(10.35.130+)
Private mstSQLType                          As SQLType                      '更新数据的方法。0-存储过程。1-SQL
Private mjeLastSubJob                       As JobException                 '上一次子任务异常标志
Private marrModleInfo                       As Variant                      '每个消息分隔后的结果
Private marrLastResponses                   As Variant                      '最后回应消息的分隔
Private mdtStartJob                         As Date                         '开始子任务的时间，本地时间
Private mdtEndJob                           As Date                         '结束子任务的时间，本地时间
Private mblnNotFirstModule                  As Boolean                      '是否不是进程启动的第一个模块

Private mstrLastTry                         As String                       '上一次重试的模块
Private mstrLastError                       As String                       '上次错误信息
Private mlngStartTimes                      As Long                         '启动进程的次数
Private mlngCheckTimes                      As Long                         '监控进程是否存在的标志
Private mobjHelperMainStartJob              As clsMutex
'4.对象变量
Private mobjMsgQueue                        As clsQueue                     '消息队列。
Private mobjSvrSend                         As clsMemoryShare
Private mobjSvrReceive                      As clsMemoryShare
Private mcnOracle                           As ADODB.Connection             '当前使用的连接
Private mobjServer                          As clsServerInfo                '当前的服务器
'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------
Private mstrServer                          As String                       '进程启动使用的服务器。IP:Port/SID
Private mblnJobRestart                      As Boolean                      '是否任务需要重启，当进程启动超时时，需要进行判断
'Server(R):当前服务器
Public Property Get Server() As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.Server")
    Server = mstrServer
    Call Logger.PopMethod("ZLHelperMain.clsJob.Server")
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.Server") = 1 Then
        Resume
    End If
End Property
'Server(R):当前服务器
Public Property Get IsRestart() As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.IsRestart")
    IsRestart = mblnJobRestart
    Call Logger.PopMethod("ZLHelperMain.clsJob.IsRestart", IsRestart)
    Exit Property
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.IsRestart") = 1 Then
        Resume
    End If
End Property
'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------
'@方法    InitJobServer
'   初始化任务的所需要的外部环境
'@返回值
'
'@参数:
'objServer clsServerInfo In
'   当前服务器
'@备注
'
Public Sub InitJobServer(ByRef objServer As clsServerInfo)
    If mstrServer <> objServer.Server Then
        Set mcnOracle = Nothing
    End If
    mstrServer = objServer.Server
    Set mobjServer = objServer
End Sub

'@方法    IsNeedRestartJob
'    是否需要重启任务。当前正在执行的任务和重新判断的任务不相同时，需要重启任务。
'@返回值  Boolean
'
'@参数:
'@备注
'
Public Function IsNeedRestartJob() As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.IsNeedRestartJob")
    If mjtJobType <> JT_NoneJob Then '当前服务器已经进行任务的处理
        If GetConnection Then
            IsNeedRestartJob = NextJob() <> mjtJobType
        Else
            IsNeedRestartJob = True         '当前服务器不能连接，则认为需要重启
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.IsNeedRestartJob", IsNeedRestartJob)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.IsNeedRestartJob") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.IsNeedRestartJob")
End Function
'@方法    FinishJob
'
'@返回值  ServerControl
'   启动任务并判断是否任务全部完成。
'@参数:
'@备注
'
Public Function FinishJob() As ServerControl
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.FinishJob")
    Call RefreshSatus
    FinishJob = JobRun
    Call Logger.PopMethod("ZLHelperMain.clsJob.FinishJob", FinishJob)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.FinishJob") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.FinishJob")
End Function
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------
'@方法    JobRun
'   启动任务进程。
'@返回值  ServerControl
'
'@参数:
'@备注
'
Private Function JobRun() As ServerControl
    Dim blnDelay        As Boolean, blnFirstRun     As Boolean
    Dim strMsg          As String
    Dim jeCurJob        As JobException
    Dim lngProcessHIS   As Long

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.JobRun")
    JobRun = SC_Continue
    jeCurJob = JE_NoneException
    Logger.DebugEx "JobStart", "JobType", mjtJobType, "JobSatus", mjsCurJobSatus, "JobRun", JobRun, "JobException", jeCurJob
    '任务初始处理
    If mjsCurJobSatus = JS_WaitExec Then
        If mjtJobType = JT_NoneJob Then
            blnFirstRun = True
            If GetConnection Then
                mjtJobType = NextJob(blnDelay)                  '判断当前执行任务
                If mjtJobType = JT_NoneJob And Not blnDelay Then
                    JobRun = SC_Finish
                    Logger.Info "任务检查", "结果", "无待处理任务"
                Else
                    If Not blnDelay Then
                        mjsCurJobSatus = JS_Running        '变更状态，在子任务中再启动进程。因为同一个任务的子任务可能会启动不同的进程。
                        Set mobjHelperMainStartJob = New clsMutex
                        If mobjHelperMainStartJob.CheckMutex(M_HELPERMAIN_STARTJOB) Then
                            Logger.Warn "任务启动互斥体创建失败"
                        End If
                        If mjtJobType = JT_VertifyFuncs Then                        '任务启动前的数据准备
                            If CheckZLUAUser(strMsg) Then
                                If Not GetVertifyFuncs() Then   '没有获取模块数据，则任务任务完成
                                    mjsCurJobSatus = JS_Successed
                                    strMsg = "无可以验证的模块。" & GetVertifyTotal() & "耗时：" & GetTickCountDiff(mlngJobStart) \ 1000 & "秒。"
                                End If
                            Else
                                mobjServer.RaiseSameException mjtJobType, "未对应上级人员表用户"
                                jeCurJob = JE_VertifyFuncUserException
                            End If
                        End If
                    Else
                        JobRun = SC_Delay
                    End If
                End If
            Else
                mobjServer.RaiseSameException mjtJobType, "连接异常"
                jeCurJob = JE_ConnectException
            End If
        End If
    End If
    Logger.DebugEx "JobRunning", "JobType", mjtJobType, "JobSatus", mjsCurJobSatus, "JobRun", JobRun, "JobException", jeCurJob
    '任务运行处理
    If mjsCurJobSatus = JS_Running And JobRun = SC_Continue And jeCurJob = JE_NoneException Then
        If blnFirstRun Then
            mlngJobStart = GetTickCount
            mlngLastResponses = mlngJobStart
            mlngCheckTimes = 0
            blnFirstRun = False
            If mjtJobType = JT_VertifyFuncs Then
                If frmMsgBox.ShowMe("将要进行当前客户端的功能验证，可能会影响你的正常使用，请先保存重要工作！") Then
                    JobRun = SC_Wait
                Else
                    If Not SaveJobStates(mjtJobType, mjsCurJobSatus, GetVertifyTotal(True)) Then     '标记任务正在执行
                        jeCurJob = JE_StateSaveException
                    End If
                End If
                Call Logger.PopMethod("ZLHelperMain.clsJob.JobRun", JobRun)
                Exit Function
            ElseIf mjtJobType = JT_ClientUpgrade Then
                lngProcessHIS = 0
                If Not FindOneProcess("ZLHISCRUST.EXE", lngProcessHIS) Then
                    If frmMsgBox.ShowMe("将要进行当前客户端的部件升级，将会关闭所有业务相关进程，请先保存重要工作！") Then
                        JobRun = SC_Wait
                    Else
                        If Not SaveJobStates(mjtJobType, mjsCurJobSatus, GetVertifyTotal(True)) Then     '标记任务正在执行
                            jeCurJob = JE_StateSaveException
                        End If
                    End If
                    Call Logger.PopMethod("ZLHelperMain.clsJob.JobRun", JobRun)
                    Exit Function
                Else
                    If Not SaveJobStates(mjtJobType, mjsCurJobSatus, GetVertifyTotal(True)) Then     '标记任务正在执行
                        jeCurJob = JE_StateSaveException
                    End If
                    Logger.Info "自动升级进程检查", "存在自动升级进程", lngProcessHIS
                End If
            Else
                If Not SaveJobStates(mjtJobType, mjsCurJobSatus, GetVertifyTotal(True)) Then     '标记任务正在执行
                    jeCurJob = JE_StateSaveException
                End If
            End If
            
        End If
        If JobRun = SC_Continue And jeCurJob = JE_NoneException Then
            strMsg = SubJobRun(jeCurJob)
        End If
    End If
    '超时控制，暂时只有运行状态存在超时异常
    If mjsCurJobSatus = JS_Running And JobRun = SC_Continue And jeCurJob = JE_NoneException Then
        If GetTickCountDiff(mlngLastResponses) > M_LNG_TIMEOUT_JOBRUN And mlngLastResponses <> 0 Then
            strMsg = "任务处理超时（>" & M_LNG_TIMEOUT_JOBRUN \ 60000 & "分钟），自动结束该任务。" & GetVertifyTotal() & "耗时：" & GetTickCountDiff(mlngJobStart) \ 1000 & "秒。"
            '可能超过24小时，若是，则自动终止任务
            mblnJobRestart = IsNeedRestartJob()
            If mstrExeName <> "" Then
                If Not IsDesinMode Then
                    If FindOneProcess(UCase(mstrExeName), mlngSubProcess) Then '按进程ID查找
                        jeCurJob = JE_ExeTimeOutOldPIDException
                    Else
                        mlngSubProcess = 0
                        If FindOneProcess(UCase(mstrExeName), mlngSubProcess) Then '按名称查找，可能进程多次重启
                            Logger.DebugEx "按名称查找进程", "ProcessID", mlngSubProcess
                            jeCurJob = JE_ExeTimeOutNewPIDException
                        Else
                            jeCurJob = JE_ExeTimeOutNoPIDException
                        End If
                    End If
                End If
            Else
                jeCurJob = JE_ExeTimeOutNoNameException
            End If
        ElseIf mlngLastResponses <> 0 And mjtJobType <> JT_VertifyFuncs Then
            '非功能验证，每20s监控下进程。
            If GetTickCountDiff(mlngLastResponses) \ 20000 > mlngCheckTimes Then
                mlngCheckTimes = mlngCheckTimes + 1
                If mstrExeName <> "" Then
                    If Not IsDesinMode Then
                        If FindOneProcess(UCase(mstrExeName), mlngSubProcess) Then '按进程ID查找
                            Logger.Info "按进程ID查找到进程"
                        Else
                            mlngSubProcess = 0
                            If FindOneProcess(UCase(mstrExeName), mlngSubProcess) Then '按名称查找，可能进程多次重启
                                Logger.Info "名称查找进程", "ProcessID", mlngSubProcess
                            Else
                                jeCurJob = JE_ExeTimeOutNoPIDException
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    '处理异常状态
    Select Case jeCurJob
        Case JE_VertifyFuncUserException
            mobjServer.RaiseSameException mjtJobType, "系统所有者对应用户检测"
            'strMsg = strMsg
        Case JE_ConnectException
            mobjServer.RaiseSameException mjtJobType, "ZLUA连接异常"
            JobRun = SC_Delay
        Case JE_ExeTimeOutNoPIDException
            mobjServer.RaiseSameException mjtJobType, "任务超时无进程"
            If mjtJobType <> JT_VertifyFuncs Then
                JobRun = SC_Delay
            End If
        Case JE_ExeTimeOutNewPIDException
            mobjServer.RaiseSameException mjtJobType, "任务超时新开进程"
            If mjtJobType <> JT_VertifyFuncs Then
                JobRun = SC_Delay
            End If
        Case JE_ExeTimeOutOldPIDException
            mobjServer.RaiseSameException mjtJobType, "任务超时原有进程"
            If mjtJobType <> JT_VertifyFuncs Then
                JobRun = SC_Delay
            End If
        Case JE_ExeTimeOutNoNameException
            mobjServer.RaiseSameException mjtJobType, "任务超时无进程名"
            If mjtJobType <> JT_VertifyFuncs Then
                JobRun = SC_Delay
            End If
        Case JE_StateSaveException
            mobjServer.RaiseSameException mjtJobType, "保存任务开始状态异常"
            JobRun = SC_Delay
        Case JE_NoneException
        Case JE_StartUpException
            If mjtJobType <> JT_VertifyFuncs Then
                mobjServer.RaiseSameException mjtJobType, "进程启动异常。" & mstrExeName
            End If
            strMsg = "进程无法启动" & mstrExeName
        Case JE_FileException
            If mjtJobType <> JT_VertifyFuncs Then
                mobjServer.RaiseSameException mjtJobType, "文件不存在。" & mstrExeName
            End If
            strMsg = "进程文件：" & mstrExeName & "不存在"
    End Select

    '多次异常判断
    If mobjServer.TryTimes > G_LNG_MAX_JOBTRY Then
        If jeCurJob = JE_ConnectException Or jeCurJob = JE_StateSaveException Then
            JobRun = SC_Finish
        Else
            jeCurJob = JE_MultiException
        End If
    End If
    Logger.DebugEx "JobFinish", "JobType", mjtJobType, "JobSatus", mjsCurJobSatus, "JobRun", JobRun, "JobException", jeCurJob
    '任务完成控制。只处理JS_Successed, JS_Failed，已经判断的无法保存或者不需要保存
    If JobRun = SC_Continue Then
        If jeCurJob <> JE_NoneException Then mjsCurJobSatus = JS_Failed     '存在异常就失败
        If mjsCurJobSatus = JS_Successed Or mjsCurJobSatus = JS_Failed Then
            If jeCurJob = JE_MultiException Then
                JobRun = SC_Finish
                strMsg = "多次尝试（>" & G_LNG_MAX_JOBTRY & "）处理该任务，但是均未成功，自动结束该任务。" & GetVertifyTotal() & "耗时：" & GetTickCountDiff(mlngJobStart) \ 1000 & "秒。"
            ElseIf strMsg = "" Then
                strMsg = GetVertifyTotal() & "耗时：" & GetTickCountDiff(mlngJobStart) \ 1000 & "秒。"
            End If
            '其余异常需要重试
            If JobRun = SC_Continue And jeCurJob <> JE_NoneException Then JobRun = SC_Delay
            '保存结果状态
            If mjtJobType = JT_VertifyFuncs Or strMsg <> "" Then
                If Not SaveJobStates(mjtJobType, mjsCurJobSatus, strMsg) Then
                    '正常失败。失败再一次重试，有些失败需要重新执行任务，有些失败需要不需要重试，由数据库处理
                    mobjServer.RaiseSameException mjtJobType, "保存任务结束状态异常"
                    If mobjServer.TryTimes > G_LNG_MAX_JOBTRY Then
                        JobRun = SC_Finish
                    End If
                End If
            Else
                '非功能验证且没有异常描述，则不保存结束状态，由独立进程处理。
            End If
        End If
    End If
    Logger.DebugEx "JobEnd", "JobType", mjtJobType, "JobSatus", mjsCurJobSatus, "JobRun", JobRun, "JobException", jeCurJob
    '异常状态清理
    If JobRun <> SC_Continue Or mjsCurJobSatus = JS_Successed Or mjsCurJobSatus = JS_Failed Then
        Call ClearJobState(CR_JobFinish)        '清理后继续进行下次轮训，下次轮训判断是否需要继续执行
    End If
    Logger.DebugEx "JobClear", "JobType", mjtJobType, "JobSatus", mjsCurJobSatus, "JobRun", JobRun, "JobException", jeCurJob
    Call Logger.PopMethod("ZLHelperMain.clsJob.JobRun", JobRun)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.JobRun") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.JobRun")
End Function

'@方法    SubJobRun
'   启动子任务进程
'@返回值  String
'   启动的异常信息
'@参数:
'jeCurJob JobException In
'   子任务任务异常
'@备注
'
Private Function SubJobRun(ByRef jeCurJob As JobException) As String
    Dim strExeName          As String
    Dim strCommand          As String
    Dim jeCurSubJob         As JobException
    Dim strMsg              As String
    Dim blnSucess           As Boolean
    Dim blnAutoErr          As Boolean
    Dim blnTimeOut          As Boolean
    Dim blnTimeOutTry       As Boolean
    Dim blnTry              As Boolean

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.SubJobRun")
    strExeName = StepGetExeName             '获取当前模块以及进程名称
    '-------------------------------------
    '1.状态处理
    '-------------------------------------
    '文件存在判断
    If strExeName = mstrExeName Then        '进程还是上一次的进程
        If mjeLastSubJob = JE_FileException Then
            strMsg = "进程文件不存在。" & strExeName
            Logger.Warn "SubJobRun", "Error", strMsg
            jeCurSubJob = mjeLastSubJob
        ElseIf mjeLastSubJob = JE_StartUpException Then
            strMsg = "进程(" & strExeName & ")无法启动，可能没有权限。"
            jeCurSubJob = JE_StartUpException
            Logger.Warn "SubJobRun", "Error", strMsg
        End If
    Else
        If Not gobjFSO.FileExists(AppsoftPath & "\" & strExeName) Then
            strMsg = "进程文件不存在。" & strExeName
            jeCurSubJob = JE_FileException
            Logger.Warn "SubJobRun", "Error", strMsg
        Else
            jeCurSubJob = JE_NoneException
        End If
        If mstrExeName <> "" Then Call ClearJobState(CR_SubJobProcessChange)
    End If
    '未启动状态处理，或重启状态处理
    If mpisCurSubJob = PIS_NoneState Then
        '启动判断
        If jeCurSubJob = JE_NoneException Then
            mlngSubProcess = 0
            If mjtJobType <> JT_VertifyFuncs Then
                If FindOneProcess(UCase(mstrExeName), mlngSubProcess) Then '按名称查找，可能进程多次重启
                    '找到了进程，讲进程纳入监控
                    mlngLastResponses = GetTickCount
                    mlngCheckTimes = 0
                    mpisCurSubJob = PIS_WaitConnect
                    mlngStartTimes = mlngStartTimes + 1
                    mdtStartJob = Now
                    mblnNotFirstModule = False
                End If
            End If
            '没找进程，则重新开辟
            If mlngSubProcess = 0 Then
                strCommand = GetExeCommand(strExeName)
                If Not RunExe(strExeName, strCommand) Then
                    strMsg = "进程(" & strExeName & ")无法启动，可能没有权限。"
                    Logger.Warn "SubJobRun", "Error", strMsg
                    jeCurSubJob = JE_StartUpException
                Else
                    mlngLastResponses = GetTickCount
                    mlngCheckTimes = 0
                    mpisCurSubJob = PIS_WaitConnect
                End If
            End If
        End If
    End If
    '等待连接状态。PIS_WaitConnect，中间状态，等待子进程反馈
    '连接状态
    If mpisCurSubJob = PIS_Connected Then
        If mobjSvrSend.WriteMemory(mobjMsgQueue.Current, mlngCurProcess, mjtJobType, PIS_SendJob) Then
            Logger.Info "发送数据成功。", "Data", mobjMsgQueue.Current, "ProcessID", mlngCurProcess, "Param", mjtJobType, "State", PIS_SendJob
            mpisCurSubJob = PIS_SendJob
            mlngLastResponses = GetTickCount
            mlngCheckTimes = 0
            mdtStartJob = Now
        End If
    End If
    '发送状态。PIS_SendJob，中间状态，等待子进程反馈
    '启动模块状态。PIS_StartJob，中间状态，等待子进程反馈
    '-------------------------------------
    '2.超时控制
    '-------------------------------------
    If jeCurSubJob = JE_NoneException Then
        If mlngLastResponses <> 0 Then
            Select Case mpisCurSubJob
                Case PIS_WaitConnect
                    If mjtJobType = JT_VertifyFuncs Then
                        If GetTickCountDiff(mlngLastResponses) > M_LNG_TIMEOUT_WAITCONNECT Then
                            strMsg = "进程(" & mstrExeName & ")已经启动，但是超时未接收到回应(>" & M_LNG_TIMEOUT_WAITCONNECT \ 1000 & "秒)。耗时：" & GetTickCountDiff(mlngLastResponses) \ 1000 & "秒。"
                            Logger.Warn "SubJobRun", "Error", strMsg
                            blnTimeOut = True
                            jeCurSubJob = JE_StartUpException
                            '可能超过24小时，若是，则自动终止任务
                            mblnJobRestart = IsNeedRestartJob()
                        End If
                    Else
                        '单独进程整体控制在JobRun
                    End If
                Case PIS_Connected
                    If GetTickCountDiff(mlngLastResponses) > M_LNG_TIMEOUT_CONNECTED Then
                        strMsg = "进程(" & mstrExeName & ")已经启动并响应，但是长时间发送任务消息失败（>" & M_LNG_TIMEOUT_WAITCONNECT \ 1000 & "秒）。耗时：" & GetTickCountDiff(mlngLastResponses) \ 1000 & "秒。"
                        Logger.Warn "SubJobRun", "Error", strMsg
                        blnTimeOut = True
                        jeCurSubJob = JE_MsgTimeOutException
                    End If
                Case PIS_SendJob
                    If GetTickCountDiff(mlngLastResponses) > M_LNG_TIMEOUT_SENDJOB Then
                        blnTimeOut = True
                        blnTimeOutTry = True
                        strMsg = "向进程(" & mstrExeName & ")发送启动模块成功，但超时未启动模块（>" & M_LNG_TIMEOUT_SENDJOB \ 1000 & "秒）。耗时：" & GetTickCountDiff(mlngLastResponses) \ 1000 & "秒。"
                        Logger.Warn "SubJobRun", "Error", strMsg
                        jeCurSubJob = JE_MsgTimeOutException
                    End If
                Case PIS_StartJob
                    If GetTickCountDiff(mlngLastResponses) > M_LNG_TIMEOUT_STARTJOB Then
                        blnTimeOut = True
                        blnTimeOutTry = True
                        strMsg = " 进程(" & mstrExeName & ")启动模块成功，但超时未结束模块（>" & M_LNG_TIMEOUT_STARTJOB \ 1000 & "秒）。耗时：" & GetTickCountDiff(mlngLastResponses) \ 1000 & "秒。"
                        Logger.Warn "SubJobRun", "Error", strMsg
                        jeCurSubJob = JE_MsgTimeOutException
                    End If
            End Select
        End If
    End If
    '-------------------------------------
    '3.异常状态处理
    '-------------------------------------
    If jeCurSubJob <> JE_NoneException Then
        '非功能验证的启动异常与文件异常
        If mjtJobType <> JT_VertifyFuncs Then
            jeCurJob = jeCurSubJob         '将该异常传递到外部
        ElseIf mjtJobType = JT_VertifyFuncs Then
            mpisCurSubJob = PIS_EndJob
        End If
    End If
    '-------------------------------------
    '4.任务完成控制
    '-------------------------------------
    mstrExeName = strExeName
    If mpisCurSubJob = PIS_EndJob Then
        If mdtEndJob < mdtStartJob Then mdtEndJob = Now
        mjeLastSubJob = jeCurSubJob
        If mjtJobType = JT_VertifyFuncs Then
            If jeCurSubJob <> JE_NoneException Then
                blnSucess = False
            Else
                blnSucess = Val(marrLastResponses(MR_TypeOrResult)) <> 0
                If marrLastResponses(MR_ComOrMessge) <> "" Or Not blnSucess Then
                    strMsg = marrLastResponses(MR_ComOrMessge)
                Else
                    strMsg = ""
                End If
            End If
            If blnTimeOutTry Then
                blnTry = True
            Else
                If strMsg <> "" Then
                    '此种错误可能是由于VB卸载部件导致，重新打开进程重试一次就正常，手工打开也正常。
                    If strMsg Like "*【错误处理】*" Then
                        blnTry = True
                    Else
                        If strMsg Like "*内存溢出*" Or strMsg Like "*从 * 加载控件 '*' 失败*" Or strMsg Like "*Automation*" Then
                            blnTry = True
                        ElseIf strMsg Like "*【初步检查】*" And Not strMsg Like "*【初步检查】部件*不能正常创建，请检查安装是否正确！信息：ActiveX 部件不能创建对象" Then
                            blnTry = True
                        End If
                    End If
                End If
            End If
            If blnTry Then
                If mstrLastTry = mobjMsgQueue.Current Then
                    blnAutoErr = False
                    mstrLastTry = ""
                    If blnSucess Then
                        strMsg = "重试成功！上次错误：" & mstrLastError
                    Else
                        strMsg = "重试失败！当前错误：" & strMsg
                    End If
                    mstrLastError = ""
                Else
                    blnAutoErr = True
                    mstrLastError = strMsg
                    mstrLastTry = mobjMsgQueue.Current
                End If
            Else
                If mstrLastTry = mobjMsgQueue.Current Then
                    If blnSucess Then
                        strMsg = "重试成功！上次错误：" & mstrLastError
                    Else
                        strMsg = "重试失败！当前错误：" & strMsg
                    End If
                End If
                mstrLastError = ""
                blnAutoErr = False
                mstrLastTry = ""
            End If

            If Not blnAutoErr Then
                If Not blnSucess And Not blnTimeOut Then
                    If InStr(strMsg, "【消息提示】") > 0 And InStr(strMsg, "【错误处理】") = 0 And InStr(strMsg, "【初步检查】") = 0 Then
                        Logger.Info "SubJobRun", "消息", "只存在警告信息"
                        blnSucess = True
                    End If
                End If
                '统计记录
                marrVertifyTotal(VMT_Default, IIf(blnSucess, VTT_Successed, VTT_Failed)) = marrVertifyTotal(VMT_Default, IIf(blnSucess, VTT_Successed, VTT_Failed)) + 1
                marrVertifyTotal(Val(marrModleInfo(MR_TypeOrResult)), IIf(blnSucess, VTT_Successed, VTT_Failed)) = marrVertifyTotal(Val(marrModleInfo(MR_TypeOrResult)), IIf(blnSucess, VTT_Successed, VTT_Failed)) + 1
                '保存任务状态，并清理上次任务信息
                strMsg = "部件：" & marrModleInfo(MR_ComOrMessge) & ",模块：" & marrModleInfo(MR_TilTle) & "。" & strMsg
                Call SaveSubJobStates(Val(marrModleInfo(MR_Sys)), Val(marrModleInfo(MR_Module)), strMsg, IIf(Not blnSucess, IIf(blnTimeOut, 2, 0), 1))
                Call mobjMsgQueue.DeQueue
                If mobjMsgQueue.IsEmpty Then
                    If marrVertifyTotal(VMT_Default, VTT_All) = marrVertifyTotal(VMT_Default, VTT_Successed) Then
                        mjsCurJobSatus = JS_Successed
                    Else
                        mjsCurJobSatus = JS_Failed
                    End If
                End If
            End If
        Else
            If marrLastResponses(MR_TypeOrResult) <> 0 Then
                mjsCurJobSatus = JS_Successed
            Else
                mjsCurJobSatus = JS_Failed
            End If
            SubJobRun = marrLastResponses(MR_ComOrMessge)
        End If
    End If

    '异常状态清理,其他任务异常，外部处理，此处只处理功能验证子任务
    If mjtJobType = JT_VertifyFuncs Then
        If jeCurSubJob <> JE_NoneException Or blnAutoErr Then
            Call ClearJobState(CR_SubJobProcessChange)
        ElseIf mpisCurSubJob = PIS_EndJob Then
            Call ClearJobState(CR_SubJobFinish)
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.SubJobRun", SubJobRun)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.SubJobRun") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.SubJobRun")
End Function

'@方法    GetConnection
'   获取连接
'@返回值  Boolean
'   是否能获取有效连接。
'@参数:
'@备注
'
Private Function GetConnection() As Boolean
    Dim strErrNote      As String
    Dim blnNewConn      As Boolean

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.GetConnection")
    If mcnOracle Is Nothing Then
        blnNewConn = True
    ElseIf mcnOracle.State <> adStateOpen Then
        blnNewConn = True
    End If
    If blnNewConn Then
        mstrUser = "ZLUA"
        mstrPwd = Sm4DecryptEcb("ZLSV2:" & G_UA_PWD, GetGeneralAccountKey(G_UA_KEY))
        Set mcnOracle = OraDBOpen(mstrServer, mstrUser, mstrPwd, MSODBC, strErrNote)
        If mcnOracle.State <> adStateOpen Then
            Call Logger.PopMethod("ZLHelperMain.clsJob.GetConnection", GetConnection)
            Exit Function
        End If
'        Call GetOSUser
    End If
    GetConnection = True
    Call Logger.PopMethod("ZLHelperMain.clsJob.GetConnection", GetConnection)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.GetConnection") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.GetConnection")
End Function

'@方法    NextJob
'   获取下一个任务类型。
'@返回值  JobType
'
'@参数:
'blnDelay Boolean In
'   当前是否只有延迟执行的任务。此时暂时停止对当前服务器的处理，等待合适的机会处理。
'@备注
'
Private Function NextJob(Optional ByRef blnDelay As Boolean) As JobType
    Dim rsTmp           As ADODB.Recordset
    Dim strSQL          As String
    Dim dtNormalUp      As Date
    Dim dtCur           As Date
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.NextJob", blnDelay)
    NextJob = JT_NoneJob
    dtCur = Currentdate(mcnOracle)
    mdblTimeSync = dtCur - Now          '获取时间同步因子
    '判断任务是否合理，获取是否启用了定时升级
    strSQL = "Select Max(内容) 内容 From ZLTOOLS.zlRegInfo Where 项目='客户端升级日期'"
    Set rsTmp = OpenSQLRecord(mcnOracle, strSQL, "检查定时升级")
    dtNormalUp = CDate(0)
    If rsTmp!内容 & "" <> "" Then
        If IsDate(rsTmp!内容 & "") Then
            dtNormalUp = CDate(Format(rsTmp!内容 & "", "YYYY-MM-DD hh:mm:ss"))
        End If
    End If
    Logger.Info "时间判断", "服务器时间", dtCur, "正是升级时间", dtNormalUp
    On Error Resume Next
    Set rsTmp = Nothing
    mdstCurDataStruc = DST_VerNB120
    '可能没有是否预升级字段(因为预升级时候，数据库还没升级），因此需要错误忽略
    strSQL = "Select Nvl(是否预升级,0) 是否预升级, Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志,a.功能验证标志   From ZLTOOLS.Zlclients a Where 工作站 = [1]"
    Set rsTmp = OpenSQLRecord(mcnOracle, strSQL, "检查当前任务", Environment.ComputerName)
    If rsTmp Is Nothing Then
        Err.Clear
        mdstCurDataStruc = DST_VerNB40
        strSQL = "Select Nvl(是否预升级,0) 是否预升级, Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志,Null 功能验证标志  From ZLTOOLS.Zlclients a Where 工作站 = [1]"
        Set rsTmp = OpenSQLRecord(mcnOracle, strSQL, "检查当前任务", Environment.ComputerName)
        If rsTmp Is Nothing Then
            Err.Clear
            mdstCurDataStruc = DST_VerNAE40
            strSQL = "Select Nvl(升级标志, 0) 是否预升级, Nvl(预升完成, 0) 预升完成, Nvl(升级标志, 0) 升级标志, Nvl(收集标志, 0) 收集标志,Null 功能验证标志 From ZLTOOLS.Zlclients Where 工作站 = [1]"
            Set rsTmp = OpenSQLRecord(mcnOracle, strSQL, "检查当前任务", Environment.ComputerName)
        End If
    End If
    On Error GoTo ErrH
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then       '客户端可能没有注册。
            '可以预升级
            If dtCur < dtNormalUp Then
                If rsTmp!是否预升级 = 1 And rsTmp!预升完成 <> 1 Then    '有预升级任务，但是预升级未完成，则预升级
                    NextJob = JT_ClientPreUpgrade
                Else    '预升级阶段，但是不能预升级，则检查部件检查任务
                    If rsTmp!收集标志 = 1 Then
                        NextJob = JT_ComponentCheck
                    ElseIf rsTmp!升级标志 = 1 Or Val(rsTmp!功能验证标志 & "") = 1 Then  '还有当前阶段不能执行的任务
                        blnDelay = True
                    End If
                End If
            Else
                If rsTmp!是否预升级 = 1 Or rsTmp!升级标志 = 1 Then      '正式升级阶段，预升级变升级
                    NextJob = JT_ClientUpgrade
                ElseIf Val(rsTmp!功能验证标志 & "") = 1 Then              '存在功能验证任务
                    NextJob = JT_VertifyFuncs
                Else
                    If rsTmp!收集标志 = 1 Then
                        NextJob = JT_ComponentCheck
                    End If
                End If
            End If
        End If
    Else
        mdstCurDataStruc = DST_VerAuto
    End If
    '预升级与正式升级飞共享才执行，否则延迟
    If NextJob = JT_ClientPreUpgrade Or NextJob = JT_ClientUpgrade Then
        blnDelay = Not IsFtpServer
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.NextJob", NextJob, blnDelay)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.NextJob") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.NextJob")
End Function

'@方法    IsFtpServer
'   当前对应服务器是否是FTP服务器
'@返回值  Boolean
'
'@参数:
'@备注
'
Private Function IsFtpServer() As Boolean
    Dim rsTmp       As ADODB.Recordset, strSQL  As String
    Dim rsServer    As ADODB.Recordset
    Dim strServerID As String, strDefualt       As String, blnDefualtIsFtp  As Boolean

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.IsFtpServer")
    On Error Resume Next
    strSQL = "Select 编号,类型, 位置, 用户名, 密码, 端口,Nvl(是否缺省,0) 是否缺省 , 批次 From Zltools.Zlupgradeserver"
    Set rsServer = OpenSQLRecord(mcnOracle, strSQL, "获取升级服务器", Val(strServerID))

    If Err.Number = 0 Then
        rsServer.Filter = "类型=0"
        If rsServer.EOF Then
            IsFtpServer = True
        Else
            rsServer.Filter = "是否缺省=1"
            If Not rsServer.EOF Then
                strDefualt = rsServer!编号 & ""
                blnDefualtIsFtp = Val(rsServer!类型 & "") <> 0
            End If
            strSQL = "Select 升级文件服务器 From ZLTools.zlClients Where 工作站=[1]"
            Set rsTmp = OpenSQLRecord(mcnOracle, strSQL, "获取升级服务器编号", Environment.ComputerName)
            If Err.Number = 0 Then
                If Not rsTmp.EOF Then strServerID = rsTmp!升级文件服务器 & ""
            Else
                Err.Clear
            End If
            If strServerID <> "" Then
                rsServer.Filter = "编号=" & Val(strServerID)
                If rsServer.EOF Then
                    IsFtpServer = blnDefualtIsFtp
                Else
                    IsFtpServer = Val(rsServer!类型 & "") <> 0
                End If
            Else
                IsFtpServer = blnDefualtIsFtp
            End If
        End If
    Else
        strSQL = "Select 项目,内容 From ZLTools.zlregInfo where 项目=[1]"
        Set rsTmp = OpenSQLRecord(mcnOracle, strSQL, "升级类型", "升级类型")
        If Not rsTmp.EOF Then
            If NvlEx(rsTmp!内容, 0) = 1 Then IsFtpServer = True
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.IsFtpServer", IsFtpServer)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.IsFtpServer") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.IsFtpServer")
End Function

'@方法    CheckZLUAUser
'   检查ZLUA对应人员表是否正常
'@返回值  Boolean
'
'@参数:
'strCheckResult String In
'   返回检查结果
'@备注
'
Private Function CheckZLUAUser(strCheckResult As String) As Boolean
    Dim rsTmp           As ADODB.Recordset
    Dim strSQL          As String

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.CheckZLUAUser")
    '检查ZLUA是否对应了人员，没有对应人员则不允许使用该功能
    strSQL = "Select 1" & vbNewLine & _
            "From 人员表 A" & vbNewLine & _
            "Where a.Id In (Select 人员ID From 上机人员表 B Where 用户名 = 'ZLUA') And Nvl(a.撤档时间, Sysdate + 1) > Sysdate And" & vbNewLine & _
            "      Nvl(a.帐号到期时间, Sysdate + 1) > Sysdate"
    Set rsTmp = OpenSQLRecord(mcnOracle, strSQL, "CheckZLUAUser", "ZLUA")
    If rsTmp.RecordCount = 0 Then
        strCheckResult = "系统所有者尚未对应人员表档案或者对应的人员表档案已经撤档或到期。"
    Else
        CheckZLUAUser = True
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.CheckZLUAUser", CheckZLUAUser, strCheckResult)
    Exit Function
ErrH:
    strCheckResult = "检测系统所有者人员档案失败。错误：" & Err.Description
    If Logger.ErrCenter("ZLHelperMain.clsJob.CheckZLUAUser") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.CheckZLUAUser")
End Function

'@方法    GetVertifyFuncs
'   获取模块功能的数据。True,获取成功。false失败。
'@返回值  Boolean
'
'@参数:
'@备注
'
Private Function GetVertifyFuncs() As Boolean
    Dim rsTmp           As ADODB.Recordset
    Dim blnReRead       As Boolean
    Dim blnSucess       As Boolean
    Dim strLastCom      As String
    Dim blnExist        As Boolean
    Dim strSystems      As String
    
    '获取验证模块
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.GetVertifyFuncs")
    strSystems = CheckSystem
    If strSystems <> "" Then
        If LenB(Environment.SessionUserSID) <> 0 Then
            If Environment.Is64bitOS Then
                mstrProxyUser = Registry.GetRegValue("HKEY_USERS\" & Environment.SessionUserSID & "\Software\WOW6432Node\VB and VBA Program Settings\ZLSOFT\注册信息\登陆信息", "USER", "ZLUA", blnSucess)
            Else
                mstrProxyUser = Registry.GetRegValue("HKEY_USERS\" & Environment.SessionUserSID & "\Software\VB and VBA Program Settings\ZLSOFT\注册信息\登陆信息", "USER", "ZLUA", blnSucess)
            End If
        Else
            mstrProxyUser = GetSetting("ZLSOFT", "注册信息\登陆信息", "USER", "ZLUA")
            blnSucess = True
        End If
        If Not blnSucess Or mstrProxyUser = "" Then
            mstrProxyUser = "ZLUA"
        End If
        mstrProxyUser = UCase(mstrProxyUser)
        '获取验证模块
        Set rsTmp = SubGetVertifyFuncs(mstrProxyUser, strSystems)
        If rsTmp Is Nothing Then
            blnReRead = mstrProxyUser <> "ZLUA"
        Else
            blnReRead = rsTmp.RecordCount = 0 And mstrProxyUser <> "ZLUA"
        End If
        If blnReRead Then
            Set rsTmp = SubGetVertifyFuncs("ZLUA", strSystems)
            If rsTmp Is Nothing Then
                Call Logger.PopMethod("ZLHelperMain.clsJob.GetVertifyFuncs", GetVertifyFuncs)
                Exit Function
            Else
                mstrProxyUser = "ZLUA"
            End If
        ElseIf rsTmp Is Nothing Then
            Call Logger.PopMethod("ZLHelperMain.clsJob.GetVertifyFuncs", GetVertifyFuncs)
            Exit Function
        End If

        If rsTmp.RecordCount > 0 Then
            mobjMsgQueue.QueueSize = rsTmp.RecordCount
            Do While Not rsTmp.EOF
                If strLastCom <> rsTmp!部件 & "" Then
                    strLastCom = rsTmp!部件 & ""
                    blnExist = False
                    '判断部件是否存在，存在则进行验证，只认为部件存在于
                    If gobjFSO.FileExists(AppsoftPath & "\APPLY\" & strLastCom & ".dll") Then
                        blnExist = True
                    ElseIf gobjFSO.FileExists(AppsoftPath & "\APPLY\" & strLastCom & ".EXE") Then
                        blnExist = True
                    Else
                        If gobjFSO.FileExists(AppsoftPath & "\PUBLIC\" & strLastCom & ".dll") Then
                            blnExist = True
                        ElseIf gobjFSO.FileExists(AppsoftPath & "\PUBLIC\" & strLastCom & ".EXE") Then
                            blnExist = True
                        Else
                            If gobjFSO.FileExists(AppsoftPath & "\" & strLastCom & ".dll") Then
                                blnExist = True
                            ElseIf gobjFSO.FileExists(AppsoftPath & "\" & strLastCom & ".EXE") Then
                                blnExist = True
                            End If
                        End If
                    End If
                End If
                If blnExist Then
                    marrVertifyTotal(VMT_Default, VTT_All) = marrVertifyTotal(VMT_Default, VTT_All) + 1
                    marrVertifyTotal(Val(rsTmp!类型), VTT_All) = marrVertifyTotal(Val(rsTmp!类型), VTT_All) + 1
                    Call mobjMsgQueue.EnQueue(rsTmp!类型 & "|" & rsTmp!系统 & "|" & rsTmp!序号 & "|" & rsTmp!部件 & "|" & rsTmp!标题)
                Else
                    Logger.Info "如下模块的部件本地未安装([APPSOFT]\APPLY,[APPSOFT],[APPSOFT]\PUBLIC)不进行功能验证", "系统", rsTmp!系统 & "", "模块", rsTmp!序号 & "", "部件", rsTmp!部件 & ""
                End If
                rsTmp.MoveNext
            Loop
        Else
            mobjMsgQueue.Clear
        End If
        GetVertifyFuncs = rsTmp.RecordCount <> 0
    Else
        Logger.Warn "当前客户端部件环境不支持功能验证"
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.GetVertifyFuncs", GetVertifyFuncs)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.GetVertifyFuncs") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsJob.GetVertifyFuncs")
End Function
'@方法    CheckSystem
'   根据部件检查当前系统能否支持功能验证
'@返回值    String
'
'@参数:
'@备注
'
Private Function CheckSystem() As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.CheckSystem")
    If gobjFSO.FileExists(AppsoftPath & "\APPLY\ZLBRW.DLL") Then
        If VerFull(gobjFSO.GetFileVersion(AppsoftPath & "\APPLY\ZLBRW.DLL")) >= VerFull("10.35.0.130") Then
            CheckSystem = CheckSystem & ",100"
        Else
            Logger.Warn AppsoftPath & "\APPLY\ZLBRW.DLL", "版本号（小于10.35.0.130）", gobjFSO.GetFileVersion(AppsoftPath & "\APPLY\ZLBRW.DLL")
        End If
    End If
    If gobjFSO.FileExists(AppsoftPath & "\APPLY\ZLLISBRW.DLL") Then
        If VerFull(gobjFSO.GetFileVersion(AppsoftPath & "\APPLY\ZLLISBRW.DLL")) >= VerFull("10.35.0.140") Then
            CheckSystem = CheckSystem & ",2500"
        Else
            Logger.Warn AppsoftPath & "\APPLY\ZLLISBRW.DLL", "版本号（小于10.35.0.130）", gobjFSO.GetFileVersion(AppsoftPath & "\APPLY\ZLLISBRW.DLL")
        End If
    End If
    If gobjFSO.FileExists(AppsoftPath & "\ZLHEALTHSTART.EXE") Then
        If VerFull(gobjFSO.GetFileVersion(AppsoftPath & "\ZLHEALTHSTART.EXE")) >= VerFull("10.35.0.130") Then
            CheckSystem = CheckSystem & ",2700"
        Else
            Logger.Warn AppsoftPath & "\ZLHEALTHSTART.EXE", "版本号（小于10.35.0.130）", gobjFSO.GetFileVersion(AppsoftPath & "\ZLHEALTHSTART.EXE")
        End If
    End If
    CheckSystem = Mid(CheckSystem, 2)
    Call Logger.PopMethod("ZLHelperMain.clsJob.CheckSystem", CheckSystem)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.CheckSystem") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.clsJob.CheckSystem")
End Function
'@方法    SubGetVertifyFuncs
'   GetVertifyFuncs的子函数
'@返回值  ADODB.Recordset
'
'@参数:
'strUser String In
'   要获取那个用户的功能模块
'strSystems String In
'   获取的系统
'@备注
'
Private Function SubGetVertifyFuncs(ByVal strUser As String, ByVal strSystems As String) As ADODB.Recordset
    Dim strSQL          As String, rsTmp        As ADODB.Recordset
    Dim blnLis          As Boolean, blnHis      As Boolean, blnHec      As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.SubGetVertifyFuncs", strUser, strSystems)
    blnHis = InStr("," & strSystems & ",", ",100,") > 0
    blnLis = InStr("," & strSystems & ",", ",2500,") > 0
    blnHec = InStr("," & strSystems & ",", ",2700,") > 0
    If blnHis Then
        If blnLis And blnHec Then
            '排序，根据不同启动程序排序
            strSQL = "Select 部件, 系统, 序号, 类型, 标题" & vbNewLine & _
                    "From (Select c.部件, c.系统, c.序号, Decode(c.系统, 2700, 3, 2500, 2, 1) 类型, c.标题" & vbNewLine & _
                    "       From (Select Distinct Upper(g.部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                    "              From zlPrograms G, zlRegFunc R" & vbNewLine & _
                    "              Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统" & vbNewLine & _
                    "              Union All" & vbNewLine & _
                    "              Select Upper(部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                    "              From zlPrograms G" & vbNewLine & _
                    "              Where 序号 Between 10000 And 19999) C," & vbNewLine & _
                    "            (Select Distinct b.系统, b.序号 From zlUserRoles A, zlRoleGrant B Where a.角色 = b.角色 And a.用户 = [1]) D," & vbNewLine & _
                    "            (Select Distinct 系统, 模块 From zlMenus A Where a.组别 In ('缺省', '体检', 'LIS系统')) E" & vbNewLine & _
                    "       Where c.系统 = d.系统 And c.序号 = d.序号 And c.部件 Not in('ZL9REPORT','ZL9EMRINTERFACE','ZL9ACCOUNT','ZL9COST','ZL9HUMAN') And c.部件 Not Like 'ZL9WIZARD%' And c.系统 = e.系统 And" & vbNewLine & _
                    "             c.序号 = e.模块)" & vbNewLine & _
                    "Order By 类型,部件,系统,序号"
        ElseIf blnLis Then
            '排序，根据不同启动程序排序
            strSQL = "Select 部件, 系统, 序号, 类型, 标题" & vbNewLine & _
                    "From (Select c.部件, c.系统, c.序号, Decode(c.系统, 2700, 3, 2500, 2, 1) 类型, c.标题" & vbNewLine & _
                    "       From (Select Distinct Upper(g.部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                    "              From zlPrograms G, zlRegFunc R" & vbNewLine & _
                    "              Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统" & vbNewLine & _
                    "              Union All" & vbNewLine & _
                    "              Select Upper(部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                    "              From zlPrograms G" & vbNewLine & _
                    "              Where 序号 Between 10000 And 19999) C," & vbNewLine & _
                    "            (Select Distinct b.系统, b.序号 From zlUserRoles A, zlRoleGrant B Where a.角色 = b.角色 And a.用户 = [1]) D," & vbNewLine & _
                    "            (Select Distinct 系统, 模块 From zlMenus A Where a.组别 In ('缺省', '体检', 'LIS系统') And a.系统 Not In(2700)) E" & vbNewLine & _
                    "       Where c.系统 = d.系统 And c.序号 = d.序号 And c.部件 Not in('ZL9REPORT','ZL9EMRINTERFACE','ZL9ACCOUNT','ZL9COST','ZL9HUMAN') And c.部件 Not Like 'ZL9WIZARD%' And c.系统 = e.系统 And" & vbNewLine & _
                    "             c.序号 = e.模块)" & vbNewLine & _
                    "Order By 类型,部件,系统,序号"
        ElseIf blnHec Then
            '排序，根据不同启动程序排序
            strSQL = "Select 部件, 系统, 序号, 类型, 标题" & vbNewLine & _
                    "From (Select c.部件, c.系统, c.序号, Decode(c.系统, 2700, 3, 2500, 2, 1) 类型, c.标题" & vbNewLine & _
                    "       From (Select Distinct Upper(g.部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                    "              From zlPrograms G, zlRegFunc R" & vbNewLine & _
                    "              Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统" & vbNewLine & _
                    "              Union All" & vbNewLine & _
                    "              Select Upper(部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                    "              From zlPrograms G" & vbNewLine & _
                    "              Where 序号 Between 10000 And 19999) C," & vbNewLine & _
                    "            (Select Distinct b.系统, b.序号 From zlUserRoles A, zlRoleGrant B Where a.角色 = b.角色 And a.用户 = [1]) D," & vbNewLine & _
                    "            (Select Distinct 系统, 模块 From zlMenus A Where a.组别 In ('缺省', '体检', 'LIS系统') And a.系统 Not In(2500)) E" & vbNewLine & _
                    "       Where c.系统 = d.系统 And c.序号 = d.序号 And c.部件 Not in('ZL9REPORT','ZL9EMRINTERFACE','ZL9ACCOUNT','ZL9COST','ZL9HUMAN') And c.部件 Not Like 'ZL9WIZARD%' And c.系统 = e.系统 And" & vbNewLine & _
                    "             c.序号 = e.模块)" & vbNewLine & _
                    "Order By 类型,部件,系统,序号"

        Else
            '排序，根据不同启动程序排序
            strSQL = "Select 部件, 系统, 序号, 类型, 标题" & vbNewLine & _
                    "From (Select c.部件, c.系统, c.序号, Decode(c.系统, 2700, 3, 2500, 2, 1) 类型, c.标题" & vbNewLine & _
                    "       From (Select Distinct Upper(g.部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                    "              From zlPrograms G, zlRegFunc R" & vbNewLine & _
                    "              Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统" & vbNewLine & _
                    "              Union All" & vbNewLine & _
                    "              Select Upper(部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                    "              From zlPrograms G" & vbNewLine & _
                    "              Where 序号 Between 10000 And 19999) C," & vbNewLine & _
                    "            (Select Distinct b.系统, b.序号 From zlUserRoles A, zlRoleGrant B Where a.角色 = b.角色 And a.用户 = [1]) D," & vbNewLine & _
                    "            (Select Distinct 系统, 模块 From zlMenus A Where a.组别 In ('缺省', '体检', 'LIS系统') And a.系统 Not In(2500,2700)) E" & vbNewLine & _
                    "       Where c.系统 = d.系统 And c.序号 = d.序号 And c.部件 Not in('ZL9REPORT','ZL9EMRINTERFACE','ZL9ACCOUNT','ZL9COST','ZL9HUMAN') And c.部件 Not Like 'ZL9WIZARD%' And c.系统 = e.系统 And" & vbNewLine & _
                    "             c.序号 = e.模块)" & vbNewLine & _
                    "Order By 类型,部件,系统,序号"

        End If
    Else
        '排序，根据不同启动程序排序
        strSQL = "Select 部件, 系统, 序号, 类型, 标题" & vbNewLine & _
                "From (Select c.部件, c.系统, c.序号, Decode(c.系统, 2700, 3, 2500, 2, 1) 类型, c.标题" & vbNewLine & _
                "       From (Select Distinct Upper(g.部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                "              From zlPrograms G, zlRegFunc R" & vbNewLine & _
                "              Where g.序号 = r.序号 And Trunc(g.系统 / 100) = r.系统" & vbNewLine & _
                "              Union All" & vbNewLine & _
                "              Select Upper(部件) As 部件, g.系统, g.序号, g.标题" & vbNewLine & _
                "              From zlPrograms G" & vbNewLine & _
                "              Where 序号 Between 10000 And 19999) C," & vbNewLine & _
                "            (Select Distinct b.系统, b.序号 From zlUserRoles A, zlRoleGrant B Where a.角色 = b.角色 And a.用户 = [1]) D," & vbNewLine & _
                "            (Select Distinct 系统, 模块 From zlMenus A Where a.组别 In ('缺省', '体检', 'LIS系统') And a.系统 in(" & strSystems & ")) E" & vbNewLine & _
                "       Where c.系统 = d.系统 And c.序号 = d.序号 And c.部件 Not in('ZL9REPORT','ZL9EMRINTERFACE','ZL9ACCOUNT','ZL9COST','ZL9HUMAN') And c.部件 Not Like 'ZL9WIZARD%' And c.系统 = e.系统 And" & vbNewLine & _
                "             c.序号 = e.模块)" & vbNewLine & _
                "Order By 类型,部件,系统,序号"
    End If

    Set rsTmp = OpenSQLRecord(mcnOracle, strSQL, "IcanVertifyFunc", strUser)
    Set SubGetVertifyFuncs = rsTmp
    Call Logger.PopMethod("ZLHelperMain.clsJob.SubGetVertifyFuncs", SubGetVertifyFuncs)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.SubGetVertifyFuncs") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.SubGetVertifyFuncs")
End Function

'@方法    GetVertifyTotal
'   获取功能验证的统计信息
'@返回值  String
'
'@参数:
'blnJobStart Boolean In
'   是否是任务开始
'@备注
'
Private Function GetVertifyTotal(Optional ByVal blnJobStart As Boolean) As String
    Dim strReturn           As String

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.GetVertifyTotal")
    If mjtJobType = JT_VertifyFuncs Then
        If blnJobStart Then
            strReturn = "验证用户：" & mstrProxyUser & "。"
            strReturn = strReturn & "共计验证模块：" & marrVertifyTotal(VMT_Default, VTT_All) & "个。"
        Else
            strReturn = "启动进程" & mlngStartTimes & "次。验证用户：" & mstrProxyUser & "。"
            If marrVertifyTotal(VMT_Default, VTT_All) <> 0 Then
                strReturn = strReturn & "共计验证模块：" & marrVertifyTotal(VMT_Default, VTT_All) & "个,验证成功模块：" & marrVertifyTotal(VMT_Default, VTT_Successed) & "个,验证失败模块：" & marrVertifyTotal(VMT_Default, VTT_Failed) & "个。"
            End If
        End If
        If marrVertifyTotal(VMT_His, VTT_All) <> 0 Then
            If blnJobStart Then
                strReturn = strReturn & "ZLHIS+.EXE模块：" & marrVertifyTotal(VMT_His, VTT_All) & "个。"
            Else
                strReturn = strReturn & "ZLHIS+.EXE模块：" & marrVertifyTotal(VMT_His, VTT_All) & "个,验证成功模块：" & marrVertifyTotal(VMT_His, VTT_Successed) & "个,验证失败模块：" & marrVertifyTotal(VMT_His, VTT_Failed) & "个。"
            End If
        End If
        If marrVertifyTotal(VMT_Lis, VTT_All) <> 0 Then
            If blnJobStart Then
                strReturn = strReturn & "ZLLIS+.EXE模块：" & marrVertifyTotal(VMT_Lis, VTT_All) & "个。"
            Else
                strReturn = strReturn & "ZLLIS+.EXE模块：" & marrVertifyTotal(VMT_Lis, VTT_All) & "个,验证成功模块：" & marrVertifyTotal(VMT_Lis, VTT_Successed) & "个,验证失败模块：" & marrVertifyTotal(VMT_Lis, VTT_Failed) & "个。"
            End If
        End If
        If marrVertifyTotal(VMT_Peis, VTT_All) <> 0 Then
            If blnJobStart Then
                strReturn = strReturn & "ZLHEALTHSTART.EXE模块：" & marrVertifyTotal(VMT_Peis, VTT_All) & "个。"
            Else
                strReturn = strReturn & "ZLHEALTHSTART.EXE模块：" & marrVertifyTotal(VMT_Peis, VTT_All) & "个,验证成功模块：" & marrVertifyTotal(VMT_Peis, VTT_Successed) & "个,验证失败模块：" & marrVertifyTotal(VMT_Peis, VTT_Failed) & "个。"
            End If
        End If
    End If
    GetVertifyTotal = strReturn
    Call Logger.PopMethod("ZLHelperMain.clsJob.GetVertifyTotal", GetVertifyTotal)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.GetVertifyTotal") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.GetVertifyTotal")
End Function
'@方法    SaveJobStates
'   保存任务状态。
'@返回值  Boolean
'
'@参数:
'jtCurType JobType In
'   当前任务类型
'jsCurStep JobStatus In
'   当前任务状态
'strMsg String In
'   消息内容
'@备注
'
Private Function SaveJobStates(ByVal jtCurType As JobType, ByVal jsCurStep As JobStatus, Optional ByVal strMsg As String) As Boolean
    Dim strSQL  As String

    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.SaveJobStates", jtCurType, jsCurStep, strMsg)
    Select Case jtCurType
        Case JT_VertifyFuncs
            If jsCurStep = JS_Successed Or jsCurStep = JS_Failed Or jsCurStep = JS_Running Then
                '通过存储过程更新
                If mstSQLType = ST_NotJudge Or mstSQLType = ST_Proceduer Then
                    strSQL = "Zltools.Zl_Zlclients_Updateprocess('" & Environment.ComputerName & "'," & (jtCurType - 1) & "," & jsCurStep & ",'" & strMsg & "')"
                    SaveJobStates = SaveSQL(strSQL, ST_Proceduer)
                    If SaveJobStates Then mstSQLType = ST_Proceduer
                End If
                '直接SQL更新
                If (mstSQLType = ST_NotJudge Or mstSQLType = ST_UpdateSQL) And Not SaveJobStates Then
                    If jsCurStep <> JS_Running Then      '成功\失败
                        strSQL = "Update Zltools.Zlclients Set 功能验证状态 = " & jsCurStep & ",功能验证标志=0, 功能验证说明 = '" & strMsg & "', 验证结束时间 = Sysdate Where 工作站 = '" & Environment.ComputerName & "'"
                    Else
                        strSQL = "Delete Zltools.Zlclientvertify A Where a.客户端 = '" & Environment.ComputerName & "'"
                        SaveJobStates = SaveSQL(strSQL, ST_UpdateSQL)
                        If SaveJobStates Then mstSQLType = ST_UpdateSQL
                        strSQL = "Update Zltools.Zlclients Set 功能验证状态 = 3, 功能验证说明 = '" & strMsg & "', 验证开始时间 = Sysdate Where 工作站 =  '" & Environment.ComputerName & "'"
                    End If
                    SaveJobStates = SaveSQL(strSQL, ST_UpdateSQL)
                    If SaveJobStates Then mstSQLType = ST_UpdateSQL
                End If
            End If
        Case Else   '其余类型超时记录。
            If jsCurStep = JS_Failed Then
                '通过存储过程更新
                If mdstCurDataStruc = DST_VerNB40 Then        '40版本以上使用存储过程
                    If mstSQLType = ST_NotJudge Or mstSQLType = ST_Proceduer Then
                        strSQL = "Zltools.Zl_Zlclients_Updateprocess('" & Environment.ComputerName & "'," & (jtCurType - 1) & "," & jsCurStep & ",'" & strMsg & "')"
                        SaveJobStates = SaveSQL(strSQL, ST_Proceduer)
                        If SaveJobStates Then mstSQLType = ST_Proceduer
                    End If
                Else   '40以下直接更新表
                    If mstSQLType = ST_NotJudge Or mstSQLType = ST_UpdateSQL Then
                        strSQL = "Zltools.Zl_Zlclients_Updateprocess('" & Environment.ComputerName & "'," & (jtCurType - 1) & "," & jsCurStep & ",'" & strMsg & "')"
                        SaveJobStates = SaveSQL(strSQL, ST_Proceduer)
                        If SaveJobStates Then mstSQLType = ST_Proceduer
                    End If
                End If
                If (mstSQLType = ST_NotJudge Or mstSQLType = ST_UpdateSQL) And Not SaveJobStates Then
                    strSQL = ""
                    If mdstCurDataStruc = DST_VerNB40 Then
                        Select Case jtCurType
                              Case JT_ClientUpgrade
                                  strSQL = "Update zlTOOLS.zlClients Set 升级情况=" & jsCurStep & " ,升级说明='" & strMsg & "' Where 工作站 = '" & Environment.ComputerName & "'"
                              Case JT_ClientPreUpgrade
                                  strSQL = "Update zlTOOLS.zlClients Set 预升完成=" & jsCurStep & " ,预升级说明='" & strMsg & "' Where 工作站 = '" & Environment.ComputerName & "'"
                              Case JT_Repaire
                                  strSQL = "Update zlTOOLS.zlClients Set 修复状态=" & jsCurStep & " ,修复说明='" & strMsg & "' Where 工作站 = '" & Environment.ComputerName & "'"
                              Case JT_ComponentCheck
                                  strSQL = "Update zlTOOLS.zlClients Set 收集状态=" & jsCurStep & " ,收集说明='" & strMsg & "' Where 工作站 = '" & Environment.ComputerName & "'"
                          End Select
                    Else
                        Select Case jtCurType
                            Case JT_ClientUpgrade
                                strSQL = "Update zlTOOLS.zlClients Set 升级情况=" & jsCurStep & " ,说明='" & strMsg & "' Where 工作站 = '" & Environment.ComputerName & "'"
                            Case JT_ClientPreUpgrade
                                strSQL = "Update zlTOOLS.zlClients Set 预升完成=" & jsCurStep & " ,说明='" & strMsg & "' Where 工作站 = '" & Environment.ComputerName & "'"
                            Case JT_Repaire
                                strSQL = "Update zlTOOLS.zlClients Set 升级情况=" & jsCurStep & " ,说明='" & strMsg & "' Where 工作站 = '" & Environment.ComputerName & "'"
                            Case JT_ComponentCheck
                                strSQL = "Update zlTOOLS.zlClients Set 收集状态=" & jsCurStep & ",说明='" & strMsg & "' Where 工作站 = '" & Environment.ComputerName & "'"
                        End Select
                    End If
                    SaveJobStates = SaveSQL(strSQL, ST_UpdateSQL)
                    If SaveJobStates Then mstSQLType = ST_UpdateSQL
                End If
            Else
                '非功能验证，非失败的情况一般返回TRUE,由升级程序处理
                SaveJobStates = True
            End If
    End Select
    Call Logger.PopMethod("ZLHelperMain.clsJob.SaveJobStates", SaveJobStates)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.SaveJobStates") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.SaveJobStates")
End Function
'@方法    ClearJobState
'   清理任务状态缓存
'@返回值
'
'@参数:
'lngClearReason ClearReason In (Optional)
'   清理原因。0-任务完成或启动。1-子任务进程变动。2-子任务开始。3-子任务完成
'@备注
'
Private Sub ClearJobState(Optional ByVal lngClearReason As ClearReason)
    Dim i           As Long
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.ClearJobState", lngClearReason)
    '任务完成或启动，清空变量
    If lngClearReason = CR_JobFinish Then
        mstSQLType = ST_NotJudge
        mjtJobType = JT_NoneJob
        mdstCurDataStruc = DST_VerAuto
        mjsCurJobSatus = JS_WaitExec
        mlngJobStart = 0
        For i = VMT_Default To VMT_Lis
            marrVertifyTotal(i, VTT_All) = 0
            marrVertifyTotal(i, VTT_Successed) = 0
            marrVertifyTotal(i, VTT_Failed) = 0
        Next
        mstrExeName = ""
        mjeLastSubJob = JE_NoneException
        mstrLastTry = ""
        mblnJobRestart = False
        mlngStartTimes = 0
        Set mobjHelperMainStartJob = Nothing
    End If
    '子任务进程切换以及
    If lngClearReason = CR_SubJobProcessChange Or lngClearReason = CR_JobFinish Then
        If Not gblnExitProcess Then        '通知服务器退出时，不再杀进程。
            If mlngSubProcess <> 0 Then
                Process.ProcessTerminate mlngSubProcess
                mlngSubProcess = 0
                mblnNotFirstModule = False
            End If
        End If
        mpisCurSubJob = PIS_NoneState
        Call mobjSvrReceive.ReleaseMemory
        Call mobjSvrSend.ReleaseMemory
    ElseIf lngClearReason = CR_SubJobFinish Then
        mpisCurSubJob = PIS_Connected
        mblnNotFirstModule = True
    End If
    '公共信息清理
    mlngLastResponses = 0
    mlngCheckTimes = 0
    If Not IsEmptyArray(marrModleInfo) Then Erase marrModleInfo
    If Not IsEmptyArray(marrLastResponses) Then Erase marrLastResponses
    mdtEndJob = Now
    mdtStartJob = mdtEndJob
    Call mobjSvrSend.ReleaseMemory
    Call mobjSvrReceive.ReleaseMemory
    Call Logger.PopMethod("ZLHelperMain.clsJob.ClearJobState")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.ClearJobState") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.ClearJobState")
End Sub
'@方法    SaveSQL
'   保存SQL.
'@返回值  Boolean
'
'@参数:
'strSQL String In
'   SQL内容
'lngSQLType SQLType In
'   0-存储过程。1-UpdateSQL
'@备注
'
Private Function SaveSQL(strSQL As String, stSQLType As SQLType) As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.SaveSQL", strSQL, stSQLType)
    If stSQLType = ST_Proceduer Then
        Call ExecuteProcedure(mcnOracle, strSQL, "ZLHelperMain.clsJob.SaveSQL")
    Else
        mcnOracle.Execute strSQL, , adCmdText
    End If
    SaveSQL = True
    Call Logger.PopMethod("ZLHelperMain.clsJob.SaveSQL", SaveSQL)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.SaveSQL") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.SaveSQL")
End Function
'@方法    StepGetExeName
'   获取当前步骤的Exe名称
'@返回值  String
'
'@参数:
'@备注
'
Private Function StepGetExeName() As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.StepGetExeName")
    If mjtJobType = JT_VertifyFuncs Then
        If Not mobjMsgQueue.IsEmpty Then
            '获取当前准备处理模块的对应进程。在完成保存时，会清理模块缓存。
            If IsEmptyArray(marrModleInfo) Then
                marrModleInfo = Split(mobjMsgQueue.Current, "|")
            End If
            If Not IsEmptyArray(marrModleInfo) Then
                StepGetExeName = DecodeEx(Val(marrModleInfo(MR_TypeOrResult)), 1, "ZLHIS+.EXE", 2, "ZLLIS+.EXE", 3, "ZLHEALTHSTART.EXE", "ZLHIS+.EXE")
            Else
                StepGetExeName = mstrExeName
            End If
        End If
    Else
        StepGetExeName = "ZLHISCRUST.EXE"
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.StepGetExeName", StepGetExeName, mobjMsgQueue.Current)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.StepGetExeName") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.StepGetExeName")
End Function
'@方法    GetExeCommand
'   获取当前任务EXE的命令行
'@返回值  String
'
'@参数:
'strExePath String In
'   程序路径
'@备注
'
Private Function GetExeCommand(ByVal strExeName As String) As String
    On Error GoTo ErrH
    Dim strCommand      As String, strCheck         As String
    Call Logger.PushMethod("ZLHelperMain.clsJob.GetExeCommand")
    If mjtJobType = JT_VertifyFuncs Then
        '启动程序不管版本兼容问题，不支持，则不启动
        strCommand = Sm4EncryptEcb("USER=ZLUA PASS=" & mstrPwd & " SERVER=" & GetServer(mstrServer) & " HELPERMAIN=" & mjtJobType & " ISDBPASS=1 PARALLELID=0")
    Else
        If UCase(GetVersionInfo(AppsoftPath & "\" & strExeName, FVN_ProductName)) = "ZLHISINSTALLUPDATE" Then
            strCommand = "ZLUPDATE:" & Cipher("USER=ZLTOOLS PASS=ZLTOOLS SERVER=" & GetServer(mstrServer) & " MODE=" & DecodeEx(mjtJobType, JT_Repaire, 0, JT_ClientUpgrade, 1, JT_ClientPreUpgrade, 2, JT_ComponentCheck, 3, 1))
        Else
            strCommand = "Provider=MSDataShape.1;Extended Properties=""Driver={Microsoft ODBC for Oracle};Server=" & GetServer(mstrServer) & _
                                   """;Persist Security Info=True;User ID=ZLUA;Password=" & mstrPwd & ";Data Provider=MSDASQL"
            strCheck = "CMDCHECK:1" & "," & Len(strCommand)
            strCommand = strCommand & "||0"
            strCheck = strCheck & "," & Len(strCommand)
            strCommand = strCommand & "||" & DecodeEx(mjtJobType, JT_Repaire, "Repair", JT_ClientPreUpgrade, "PreUpgrade", JT_ClientUpgrade, "OfficialUpgrade", JT_ComponentCheck, "CheckUpgrade", "OfficialUpgrade")
            strCheck = strCheck & "," & Len(strCommand)
            strCommand = strCommand & "||"
            strCheck = strCheck & "," & Len(strCommand)
            strCommand = strCommand & "||"
            strCheck = strCheck & "," & Len(strCommand)
            strCommand = strCommand & "||" & strCheck
        End If
    End If
    GetExeCommand = strCommand
    Call Logger.PopMethod("ZLHelperMain.clsJob.GetExeCommand", Sm4EncryptEcb(GetExeCommand))
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.GetExeCommand") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.GetExeCommand")
End Function

'@方法    RunExe
'   执行任务进程
'@返回值  Boolean
'
'@参数:
'strExeName String In
'   任务进程Exe名称。
'strCommand String In(Optional)
'   执行任务进程的进程命令行
'@备注
'
Private Function RunExe(ByVal strExeName As String, Optional ByVal strCommand As String) As Boolean
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.RunExe", strExeName, strCommand)
    If IsDesinMode Then mlngSubProcess = 3368
    mlngSubProcess = Shell(AppsoftPath & "\" & strExeName & " " & strCommand)
    RunExe = mlngSubProcess <> 0
    mlngStartTimes = mlngStartTimes + 1
    mdtStartJob = Now
    Logger.DebugEx "RunExe", "ProcessID", mlngSubProcess, "StartTime", mlngJobStart
    mblnNotFirstModule = False
    Call Logger.PopMethod("ZLHelperMain.clsJob.RunExe", RunExe)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.RunExe") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.RunExe", RunExe)
End Function
'@方法    SaveSubJobStates
'   保存子任务状态
'@返回值  Boolean
'
'@参数:
'lngSys Long In
'   系统编号
'lngModule Long In
'   模块号
'strMsg String In
'   消息内容
'lngResult Long In
'   结果值
'@备注
'
Private Function SaveSubJobStates(ByVal lngSys As Long, ByVal lngModule As Long, ByVal strMsg As String, ByVal lngResult As Long) As Boolean
    Dim strSQL          As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.SaveSubJobStates")
    '通过存储过程更新
    If mstSQLType = ST_NotJudge Or mstSQLType = ST_Proceduer Then
        strSQL = "Zltools.Zl_Zlclientvertify_Insert('" & Environment.ComputerName & "'," & lngSys & "," & lngModule & "," & To_DateEx(Format(mdtStartJob + mdblTimeSync, "YYYY-MM-DD hh:mm:ss")) & "," & To_DateEx(Format(mdtEndJob + mdblTimeSync, "YYYY-MM-DD hh:mm:ss")) & "," & lngResult & ",'" & strMsg & "')"
        SaveSubJobStates = SaveSQL(strSQL, ST_Proceduer)
        If SaveSubJobStates Then mstSQLType = ST_Proceduer
    End If
    '直接SQL更新
    If (mstSQLType = ST_NotJudge Or mstSQLType = ST_UpdateSQL) And Not SaveSubJobStates Then
        strSQL = "Insert Into Zltools.Zlclientvertify(客户端, 系统, 模块, 验证开始时间, 验证结束时间, 验证结果, 验证说明) Values ('" & Environment.ComputerName & "'," & lngSys & "," & lngModule & "," & To_DateEx(Format(mdtStartJob + mdblTimeSync, "YYYY-MM-DD hh:mm:ss")) & "," & To_DateEx(Format(mdtEndJob + mdblTimeSync, "YYYY-MM-DD hh:mm:ss")) & "," & lngResult & ",'" & strMsg & "') "
        SaveSubJobStates = SaveSQL(strSQL, ST_Proceduer)
        If SaveSubJobStates Then mstSQLType = ST_UpdateSQL
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.SaveSubJobStates", SaveSubJobStates)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.SaveSubJobStates") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.SaveSubJobStates")
End Function
'@方法    RefreshSatus
'   从共享内存读取数据，并刷新状态
'@返回值
'
'@参数:
'@备注
'
Private Sub RefreshSatus()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.RefreshSatus")
    If mobjSvrReceive.ReadMemoryOnce() Then
        Logger.DebugEx "SvrReceive.ReadMemory", "ProcessID", mobjSvrReceive.ProcessID, "Param", mobjSvrReceive.Param, "State", mobjSvrReceive.State, "Data", mobjSvrReceive.Data
        If mobjSvrReceive.Param = mjtJobType Then
            If IsDesinMode And mobjSvrReceive.ProcessID <> mlngSubProcess Then
                mlngSubProcess = mobjSvrReceive.ProcessID
            End If
            If mobjSvrReceive.ProcessID = mlngSubProcess And mjtJobType = JT_VertifyFuncs Or mjtJobType <> JT_VertifyFuncs Then
                mpisCurSubJob = mobjSvrReceive.State
                marrLastResponses = Split(mobjSvrReceive.Data, "|")
                mlngLastResponses = GetTickCount
                mlngCheckTimes = 0
                Select Case mpisCurSubJob
                    Case PIS_Connected           '
                        If mjtJobType <> JT_VertifyFuncs Then  '自动升级完成后,不再判断进程ID,因为可能多次启动
                            mpisCurSubJob = PIS_EndJob
                        End If
                    Case PIS_StartJob
                    Case PIS_EndJob
                        mblnNotFirstModule = True
                        mdtEndJob = Now
                End Select
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.RefreshSatus")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.RefreshSatus") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.RefreshSatus")
End Sub


'@方法    ClearProcess
'   清理所有升级助手相关的进程
'@返回值
'
'@参数:
'Param 方向
'
'@备注
'
Private Sub ClearProcess()
    Dim arrTmp      As Variant
    Dim i           As Long
    '清理ZLHIS进程
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.ClearProcess")
    arrTmp = Process.ProcessesByProcessName("ZLHIS+.EXE", glngSendProcess)
    For i = LboundEx(arrTmp) To UboundEx(arrTmp)
        If Environment.ProcessUser(arrTmp(i)) = "SYSTEM" Then
            Process.ProcessTerminate arrTmp(i)
        End If
    Next
    '清理ZLLIS.EXE进程
    arrTmp = Process.ProcessesByProcessName("ZLLIS+.EXE", glngSendProcess)
    For i = LboundEx(arrTmp) To UboundEx(arrTmp)
        If Environment.ProcessUser(arrTmp(i)) = "SYSTEM" Then
            Process.ProcessTerminate arrTmp(i)
        End If
    Next
    '清理体检ZLHEALTHSTART.EXE进程
    arrTmp = Process.ProcessesByProcessName("ZLHEALTHSTART.EXE", glngSendProcess)
    For i = LboundEx(arrTmp) To UboundEx(arrTmp)
        If Environment.ProcessUser(arrTmp(i)) = "SYSTEM" Then
            Process.ProcessTerminate arrTmp(i)
        End If
    Next
    '清理ZLHISCrust.EXE进程
    arrTmp = Process.ProcessesByProcessName("ZLHISCRUST.EXE", glngSendProcess, False)
    For i = LboundEx(arrTmp) To UboundEx(arrTmp)
        If Environment.ProcessUser(arrTmp(i)) = "SYSTEM" Then
            Process.ProcessTerminate arrTmp(i)
        End If
    Next
    Call Logger.PopMethod("ZLHelperMain.clsJob.ClearProcess")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.ClearProcess") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsJob.ClearProcess")
End Sub

'@方法    FindOneProcess
'   获取一个指定的进程
'@返回值  Long
'
'@参数:
'Param 方向
'
'@备注
'
Private Function FindOneProcess(ByVal strProcessName As String, Optional ByRef lngCurPID As Long) As Boolean
    Dim arrTmp      As Variant
    Dim i           As Long
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.FindOneProcess", strProcessName, lngCurPID)
    arrTmp = Process.ProcessesByProcessName("ZLHISCRUST.EXE")
    If lngCurPID = 0 Then
        For i = LboundEx(arrTmp) To UboundEx(arrTmp)
            lngCurPID = arrTmp(i)
            FindOneProcess = True
            Exit For
        Next
    Else
        For i = LboundEx(arrTmp) To UboundEx(arrTmp)
            If lngCurPID = arrTmp(i) Then
                FindOneProcess = True
                Exit For
            End If
        Next
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.FindOneProcess", FindOneProcess, lngCurPID)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.FindOneProcess") = 1 Then
        Resume
    End If

    Call Logger.PopMethod("ZLHelperMain.clsJob.FindOneProcess")
End Function

'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.Class_Initialize")
    mlngCurProcess = GetCurrentProcessId
    Set mobjSvrSend = New clsMemoryShare
    Set mobjSvrReceive = New clsMemoryShare
    If Not mobjSvrReceive.CreateMemoryShare(M_SHARE_SVR_RECEIVE, G_MAX_MEMORY_SIZE) Then
        Logger.Warn "CreateMemoryShare", "SVR_RECEIVE", False
    End If
    If Not mobjSvrSend.CreateMemoryShare(M_SHARE_SVR_SEND, G_MAX_MEMORY_SIZE) Then
        Logger.Warn "CreateMemoryShare", "SVR_SEND", False
    End If
    Set mobjMsgQueue = New clsQueue
    mjtJobType = JT_NoneJob
    Call Logger.PopMethod("ZLHelperMain.clsJob.Class_Initialize")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.Class_Initialize") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.Class_Initialize")
End Sub

Private Sub Class_Terminate()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.clsJob.Class_Terminate")
    If Not mcnOracle Is Nothing Then
        If mcnOracle.State = adStateOpen Then Call mcnOracle.Close
        Set mcnOracle = Nothing
    End If
    Call ClearJobState(CR_JobFinish)
    Call ClearProcess
    Set mobjSvrSend = Nothing
    Set mobjSvrReceive = Nothing
    Set mobjMsgQueue = Nothing
    Call Logger.PopMethod("ZLHelperMain.clsJob.Class_Terminate")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.clsJob.Class_Terminate") = 1 Then
        Resume
    End If
    Call Logger.PopMethod("ZLHelperMain.clsJob.Class_Terminate")
End Sub

