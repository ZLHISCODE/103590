Attribute VB_Name = "mdlHelperMain"
'@模块 mdlHelperMain-2019/7/5
'@编写 lshuo
'@功能
'
'@引用
'
'@备注
'
Option Explicit
'---------------------------------------------------------------------------
'                0、API和常量声明
'---------------------------------------------------------------------------
Public Const G_HELPER_RECEIVE              As String = "67332524-C38A-4318-85C4-FA8151C85EDD"          '向升级助手发送命令信息的内存共享
Public Const G_HELPER_SEND                 As String = "012CB54B-017C-4B9B-A0D5-956C0E90EFA6"          '升级助手向进程发送消息的内存共享
Public Enum Helper_Share_Type
    HST_Waitting = 0            '等待状态
    HST_UpgradeToServer = 1     '自动升级向服务器启动的升级助手发送消息。
    HST_ServerToUpgrade = 2     '升级助手向自动升级发送消息
    HST_SelfToServer = 3        '手工启动向服务器启动的升级助手发送消息
End Enum
Public Enum Helper_Share_State
    HST_UnWrited = 0            '未写入
    HST_Writed = 1              '已经写入
End Enum
Public Const G_SINGLE_INSTANCE              As String = "DEF43DC9-722D-48E0-9CBD-73E20E373E86"          '保证单实例运行
Private Const M_MAX_LOG_COUNT               As Long = 600000
'Private Const M_MAX_LOG_COUNT               As Long = 100
Public Const G_MAX_MEMORY_SIZE              As Long = 2048
Public Const G_LNG_MAX_JOBTRY               As Long = 4
Public Const G_LNG_MAX_SUBJOBTRY            As Long = 3


Public Const gstrSysName                    As String = "中联软件"
'---------------------------------------------------------------------------
'                1、常规变量
'---------------------------------------------------------------------------
'全局对象
Public Logger                               As New clsLogger
Public Process                              As New clsProcess
Public Environment                          As New clsEnvironment
Public Registry                             As New clsRegistry

Public gobjFSO                              As New FileSystemObject


'使用SendMesage将会因为服务启动的进程是SYSTEM权限，而自动升级启动的是普通用户权限进程。导致消息发送不成功
Public gobjMetux                            As New clsMutex
Public gobjHelperMainRECEIVE                As New clsMemoryShare               '升级读取消息
Public gobjHelperMainSend                   As New clsMemoryShare               '升级发送消息
Public gstrCommand                          As String

Public gobjServerQueue                      As New clsQueue
Public gcllExecption                        As New Collection               '服务器异常

Public gblnMsgBox                           As Boolean
Public gblnExitProcess                      As Boolean
Public glngSendProcess                      As Long
Public Const G_VESRION                      As Long = 1
'---------------------------------------------------------------------------
'                2、属性变量与定义
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                3、公共方法
'---------------------------------------------------------------------------
Sub Main()
    Dim blnServer           As Boolean
    Dim arrVar              As Variant, i       As Long
    Dim objService          As clsService
    
    Const SE_DEBUG_NAME     As String = "SeDebugPrivilege"
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlHelperMain.Main")
    '设置内部组件的具体错误级别
    Logger.AddComponentLogLevel "ZLHelperMain.clsMutex", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsMemoryShare", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsServerInfo", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.frmMain", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsProcess", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsEnvironment", LogLevel_Info
    Logger.AddComponentLogLevel "ZLHelperMain.clsRegistry", LogLevel_Info
    '服务启动升级助手：标记当前会话ID,操作系统用户，操作系统用户域。SVRSTART SESSIONID=5 USERNAME=Administrator DOMAIN=Win7Base-PC
    '自动升级请求更新操作:此时缓存状态信息：HELPERUPGRADE SAVEANDEXIT
    '轮训一个数据库：EXCFUNC DB=192.168.33.201:1521/TESTBASE35
    '当前进程退出清理：EXIT
    If Environment.IsDesinMode Then
        gstrCommand = "EXCFUNC DB=127.0.0.1:1521/TESTBASE"
'        gstrCommand = "SVRSTART SESSIONID=1 USERNAME=Administrator DOMAIN=Win7Work-PC"
'        gstrCommand = "HELPERUPGRADE SAVEANDEXIT"
'        gstrCommand = "EXIT"
    Else
        gstrCommand = Trim(CStr(Command()))
    End If

    If gstrCommand Like "SVRSTART *" Then
        If Environment.ProcessUser = "SYSTEM" Then
            blnServer = True
        End If
    End If
    If blnServer Then
        Logger.OpenEx , "V" & G_VESRION, LogType_SingleInstance, False, M_MAX_LOG_COUNT, Logger.CurrentLogLevel, Logger.IsLoopTimeFormat
        '清理其他的ZLHelperMain.EXE进程
        arrVar = Process.ProcessesByProcessName("ZLHelperMain.EXE", GetCurrentProcessId())
        For i = LboundEx(arrVar) To UboundEx(arrVar)
            Process.ProcessTerminate arrVar(i)
        Next
        If Not gobjMetux.CheckMutex(G_SINGLE_INSTANCE) Then
            If Not gobjHelperMainRECEIVE.CreateMemoryShare(G_HELPER_RECEIVE) Or Not gobjHelperMainSend.CreateMemoryShare(G_HELPER_SEND) Then
                Set gobjHelperMainRECEIVE = Nothing
                Set gobjHelperMainSend = Nothing
                Set gobjMetux = Nothing
            End If
        Else
            Set gobjMetux = Nothing
        End If
        Call RestoreEnv
        Load frmMain
    Else
        If gobjMetux.CheckMutex(G_SINGLE_INSTANCE) Then
            Logger.OpenEx , "First_V" & G_VESRION, LogType_SingleInstance, False, M_MAX_LOG_COUNT, Logger.CurrentLogLevel, Logger.IsLoopTimeFormat
            If gstrCommand <> "" Then
                If gobjHelperMainRECEIVE.OpenMemoryShare(G_HELPER_RECEIVE) Then
                    '失败也不再处理，简化逻辑，常规来说不会失败
                    Call gobjHelperMainRECEIVE.WriteMemory(gstrCommand, GetCurrentProcessId(), HST_SelfToServer, HST_Writed, False)
                End If
            End If
        Else
            Set gobjMetux = Nothing
            Logger.OpenEx , "Second_V" & G_VESRION, LogType_SingleInstance, False, M_MAX_LOG_COUNT, Logger.CurrentLogLevel, Logger.IsLoopTimeFormat
            'TODO,保存命令行
            Call RestoreEnv
            gobjServerQueue.EnQueue gstrCommand
            Call SaveEnv
            If Not Process.EnablePrivilege(SE_DEBUG_NAME) Then
                Logger.DebugEx SE_DEBUG_NAME, "结果", "失败"
            Else
                Logger.DebugEx SE_DEBUG_NAME, "结果", "成功"
            End If
        End If
        If Not gstrCommand Like "EXCFUNC DB=*" And gstrCommand <> "HELPERUPGRADE SAVEANDEXIT" And gstrCommand <> "EXIT" And Not gstrCommand Like "SVRSTART *" Then
            Set objService = New clsService
            If Not objService.IsInstalled("ZLHelperService") Then
                If gobjFSO.FileExists(AppsoftPath & "\ZLHelperService.exe") Then
                    Call objService.Install("ZLHelperService", "ZLSOFT Upgrade Helper Service", "中联升级助手服务", AppsoftPath & "\ZLHelperService.exe")
                End If
            End If
            If objService.IsInstalled("ZLHelperService") Then
                If Not objService.IsRunning("ZLHelperService") Then
                    Call objService.Start("ZLHelperService")
                End If
            End If
            Set objService = Nothing
        End If
        Call ProcessExit
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.Main")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlHelperMain.Main") = 1 Then
        Resume
    End If
    Call ProcessExit
End Sub

'@方法    RestoreEnv
'   恢复服务器轮训信息
'@返回值
'
'@参数:
'@备注
'
Public Sub RestoreEnv()
    Dim i           As Long, j As Long
    Dim strTmp      As String
    Dim objText     As TextStream
    Dim arrServer   As Variant
    Dim strFile     As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlHelperMain.RestoreEnv")
    strFile = AppsoftPath & "\ZLHelperMain_SessionID_" & Environment.SessionID & ".ini"
    If gobjFSO.FileExists(strFile) Then
        Set objText = gobjFSO.OpenTextFile(strFile, ForReading)
        strTmp = objText.ReadLine
        objText.Close
        Set objText = Nothing
        If strTmp Like "SERVER=*" Then
            strTmp = Mid(strTmp, Len("SERVER=*"))
            arrServer = UnSerialize(strTmp)
            If IsArray(arrServer) Then
                For i = LBound(arrServer) To UBound(arrServer)
                    If IsArray(arrServer(i)) Then
                        For j = LBound(arrServer(i)) To UBound(arrServer(i))
                            Call gobjServerQueue.EnQueue(arrServer(i)(j))
                        Next
                    Else
                        Call gobjServerQueue.EnQueue(arrServer(i))
                    End If
                Next
            Else
                Call gobjServerQueue.EnQueue(arrServer)
            End If
        End If
        gobjFSO.DeleteFile strFile, True
    End If
    Logger.Info "读取到缓存服务器", "Count", gobjServerQueue.Count
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.RestoreEnv")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlHelperMain.RestoreEnv") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.RestoreEnv")
End Sub

'--------------------------------------------------------------------------------------------------
'方法           SaveEnv
'功能           保存服务器轮训信息，只有准备退出时才保存
'返回值
'入参列表:
'参数名         类型                    说明
'strServer      String                  TCP在退出时发送来的新数据。
'-------------------------------------------------------------------------------------------------
Public Sub SaveEnv()
    Dim objQueue    As New clsQueue
    Dim i           As Long
    Dim strTmp      As String
    Dim objText     As TextStream
    Dim strServer   As String
    On Error GoTo ErrH
    Call Logger.PushMethod("mdlHelperMain.mdlHelperMain.SaveEnv")
    objQueue.QueueSize = 5
    Do While Not gobjServerQueue.IsEmpty
        If gobjServerQueue.Current = "HELPERUPGRADE SAVEANDEXIT" Then
            'Do Nothing
        ElseIf Not gobjServerQueue.Current Like "SVRSTART *" Then
            'EXCFUNC DB=192.168.33.201:1521/TESTBASE35
            If gobjServerQueue.Current Like "EXCFUNC DB=*" Then
                strServer = Mid(gobjServerQueue.Current, Len("EXCFUNC DB=*"))
            Else
                strServer = gobjServerQueue.Current
            End If
            objQueue.EnQueue strServer
        End If
        gobjServerQueue.DeQueue
    Loop
    
    For i = gcllExecption.Count To 1 Step -1
        objQueue.EnQueue gcllExecption(i).Server
        gcllExecption.Remove i
    Next
    If Not objQueue.IsEmpty() Then
        strTmp = SerializeEx(objQueue.Data)
        Logger.Info "保存缓存服务器", "Count", objQueue.Count
        Set objText = gobjFSO.OpenTextFile(AppsoftPath & "\ZLHelperMain_SessionID_" & Environment.SessionID & ".ini", ForWriting, True)
        objText.WriteLine "SERVER=" & strTmp
        objText.Close
        Set objText = Nothing
    End If
    
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.SaveEnv")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlHelperMain.SaveEnv") = 1 Then
        Resume
    End If
    
    Call Logger.PopMethod("mdlHelperMain.mdlHelperMain.SaveEnv")
End Sub

Public Sub ProcessExit()
    Dim objService      As New clsService
    Set gobjHelperMainSend = Nothing
    Set gobjHelperMainRECEIVE = Nothing
    Set gobjMetux = Nothing
    Set gcllExecption = Nothing
    '通知退出时关闭服务
    If gblnExitProcess Then
        If objService.IsInstalled("ZLHelperService") Then
            If objService.IsRunning("ZLHelperService") Then
                Call objService.Stopping("ZLHelperService")
            End If
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlHelperMain.Main")
    Set Registry = Nothing
    Set Process = Nothing
    Set Environment = Nothing
    Call Logger.CloseEx
    Set Logger = Nothing
End Sub
'---------------------------------------------------------------------------
'                4、私有方法
'---------------------------------------------------------------------------

'---------------------------------------------------------------------------
'                5、对象方法与事件
'---------------------------------------------------------------------------

