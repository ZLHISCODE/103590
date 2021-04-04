Attribute VB_Name = "mdlMain"
Option Explicit
'==================================================================================================
'编写           lshuo
'日期           2019/1/21
'模块           mdlMain
'说明
'==================================================================================================
Public gstrComputerName                     As String
Public gstrIP                               As String
Public gstrAppSOFT                          As String                       'APPSOFT路径
Public gobjFSO                              As New FileSystemObject         '公共全局文件操作对象
Public gcllExecption                        As New Collection               '服务器异常
Public gobjCurJob                           As clsJob                       '当前的服务器任务
Public Const G_LNG_MAX_JOBTRY               As Long = 4
Public Const G_LNG_MAX_SUBJOBTRY            As Long = 4
Public gblnMsgBox                           As Boolean
Public gstrCommand                          As String
Public glngWinSessionID                     As Long
Public Enum StartType
    ST_Service = 0              '服务启动的升级助手
    ST_SendServer = 1           '向升级助手发送消息
    ST_SaveServer = 2           '保存服务器病退出
    ST_Exit = 3                 '
End Enum
Public gstStartType                         As StartType
'使用SendMesage将会因为服务启动的进程是SYSTEM权限，而自动升级启动的是普通用户权限进程。导致消息发送不成功
Public gobjMetux                            As New clsMutex
Public gobjHelperMain                       As New clsMemoryShareFP         '升级助手间相互通讯
Public gobjSessionSID                       As New clsMemoryShareFP         '当前会话SID
Private Const M_HELPER_MAIN                 As String = "67332524-C38A-4318-85C4-FA8151C85EDD"          '升级助手间进程通讯ID
Private Const M_SINGLE_INSTANCE             As String = "DEF43DC9-722D-48E0-9CBD-73E20E373E86"          '保证单实例运行
Private Const M_CURRENT_SESSION_SID         As String = "DB130B49-92AA-488D-9D58-C1671CD21673"          '存储SID
Public Const G_HISCUST_RUNNINH              As String = "7AA64FD7-9966-46D9-A10C-420AD5CEC766"          '自动升级正在运行
Public gstrCurSID                           As String       '当前会话的sid
Public Logger                               As New clsLogger
Private Const M_MAX_LOG_COUNT               As Long = 600000
Public Const G_MAX_MEMORY_SIZE              As Long = 2048
Private mblnNormalTime                      As Boolean              '日志是否自然时间格式
Public gblnExit                             As Boolean              '是否保存信息并退出
Public gstrSystems                          As String               '检查的系统
Public Const gstrSysName                    As String = "中联软件"
Public gstrZLSOFTRegKey                     As String

Sub Main()
    Dim objService      As clsService
    Dim blnNormalTime   As Boolean
    On Error GoTo ErrH
    Logger.IsUseCache = True        '放在开头，这时日志还没有开启，先使用缓存记录
    Call Logger.PushMethod("ZLHelperMain.mdlMain.Main")
    glngWinSessionID = GetCurrentSessionID()
    '服务启动升级助手：标记当前会话ID,操作系统用户，操作系统用户域。SVRSTART SESSIONID=5 USERNAME=Administrator DOMAIN=Win7Base-PC
    '自动升级请求更新操作:此时缓存状态信息：HELPERUPGRADE SAVEANDEXIT
    '轮训一个数据库：EXCFUNC DB=192.168.33.201:1521/TESTBASE35
    '当前进程退出清理：EXIT
    If IsDesinMode Then
'        gstrCommand = "EXCFUNC DB=192.168.33.201:1521/TESTBASE35"
        gstrCommand = "SVRSTART SESSIONID=1 USERNAME=Administrator DOMAIN=Win7Work-PC"
'        gstrCommand = "HELPERUPGRADE SAVEANDEXIT"
'        gstrCommand = "EXIT"
    Else
        gstrCommand = Trim(CStr(Command()))
    End If
    gstStartType = ST_Service
    If gstrCommand Like "SVRSTART *" Then
        Call gobjMetux.CheckMutex(M_SINGLE_INSTANCE)
        Logger.Info "当前是升级助手服务开辟的升级助手后台进程。", "参数", gstrCommand
        If UCase(GetProcessUserName) <> "SYSTEM" Then
            gstStartType = ST_Exit
            Set gobjMetux = Nothing
        Else
            Call gobjHelperMain.CreateMemoryShare(M_HELPER_MAIN, G_MAX_MEMORY_SIZE)
        End If
        gstrCurSID = GetUserSID(GetSessionUser(gstrCommand))
        If gstrCurSID = "" Then
            gstrCurSID = GetUserSID(GetSessionUser(gstrCommand, True))
        End If
        If gstrCurSID <> "" Then
            Call gobjSessionSID.CreateMemoryShare(M_CURRENT_SESSION_SID, G_MAX_MEMORY_SIZE)
            Call gobjSessionSID.WriteMemory(gstrCurSID)
        End If
        Logger.Info "进程所属会话SID", gstrCurSID
    Else
        If gobjMetux.CheckMutex(M_SINGLE_INSTANCE) Then
            '此时向升级助手进程发送命令行信息
            Logger.Info "当前已经存在升级助手服务开辟的升级助手后台进程。", "参数", gstrCommand
            gstStartType = ST_SendServer
            Call gobjHelperMain.OpenMemoryShare(M_HELPER_MAIN)
        Else
            Logger.Info "当前不是升级助手服务开辟的升级助手后台进程。", "参数", gstrCommand
            '此时没有后台进程运行，此时只需要保存到本地配置文件即可。
            gstStartType = ST_SaveServer
        End If
        Set gobjMetux = Nothing
    End If
    Call InitEnv
    '执行退出保存命令时，不需要安装服务
    If Not (gstrCommand Like "HELPERUPGRADE *" Or gstStartType = ST_SaveServer Or gstStartType = ST_SendServer Or gstrCommand = "EXIT") Then
        If gstStartType > ST_Service Then
          If Not EnablePrivilege(, SE_DEBUG_NAME) Then
                Logger.Info SE_DEBUG_NAME & "失败"
            Else
                Logger.Info SE_DEBUG_NAME & "成功"
            End If
        End If
        Set objService = New clsService
        If Not objService.IsInstalled("ZLHelperService") Then
            If gobjFSO.FileExists(gstrAppSOFT & "\ZLHelperService.exe") Then
                Call objService.Install("ZLHelperService", "ZLSOFT Upgrade Helper Service", "中联升级助手服务", gstrAppSOFT & "\ZLHelperService.exe")
            End If
        End If
        If objService.IsInstalled("ZLHelperService") Then
            If Not objService.IsRunning("ZLHelperService") Then
                Call objService.Start("ZLHelperService")
            End If
        End If
        Set objService = Nothing
    End If
    If (gstrCommand Like "HELPERUPGRADE *" And gstStartType = ST_SaveServer) Or gstStartType = ST_Exit Or gstrCommand = "" Or gstrCommand = "EXIT" Then
        Set gobjHelperMain = Nothing
        Set gobjSessionSID = Nothing
        Set gobjMetux = Nothing
        Set gcllExecption = Nothing
        Call Logger.PopMethod("ZLHelperMain.mdlMain.Main")
        Call Logger.CloseEx
        Set Logger = Nothing
        End
    Else
        Load frmMain
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlMain.Main")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.Main") = 1 Then
        Resume
    End If
    Call Logger.CloseEx
End Sub

'--------------------------------------------------------------------------------------------------
'方法           GetSessionUser
'功能           解析命令行获取当前会话用户
'返回值         String
'入参列表:
'参数名         类型                    说明
'strCommand     String                  命令行
'blnOnlyUser    Boolean                 单单只返回用户名
'-------------------------------------------------------------------------------------------------
Public Function GetSessionUser(ByVal strCommand As String, Optional ByVal blnOnlyUser As Boolean) As String
    Dim arrTmp  As Variant
    Dim strTmp  As String
    Dim strUser As String, strDomain    As String
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlMain.GetSessionUser", strCommand, blnOnlyUser)
    arrTmp = Split(strCommand, "USERNAME=")
    strTmp = arrTmp(1)
    arrTmp = Split(strTmp, "DOMAIN=")
    strUser = Trim(arrTmp(0))
    strDomain = Trim(arrTmp(1))
    If blnOnlyUser Then
        GetSessionUser = strUser
    Else
        GetSessionUser = strDomain & "\" & strUser
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlMain.GetSessionUser", GetSessionUser)
    Exit Function
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.GetSessionUser") = 1 Then
        Resume
    End If
End Function

'--------------------------------------------------------------------------------------------------
'方法           InitEnv
'功能           初始化环境
'返回值
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Private Sub InitEnv()
    Dim objReg          As New clsRegistry
    Dim strRet          As String
    Dim blnLoopTime     As Boolean, llLogLevel       As LogLevel
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlMain.InitEnv")
    gstrIP = IP
    gstrComputerName = ComputerName
    If IsDesinMode Then
        gstrAppSOFT = "C:\APPSOFT"
    Else
        gstrAppSOFT = App.Path
    End If
    If gstrCurSID <> "" Then
        If Is64bit Then
            gstrZLSOFTRegKey = "HKEY_USERS\" & gstrCurSID & "\Software\WOW6432Node\VB and VBA Program Settings\ZLSOFT"
        Else
            gstrZLSOFTRegKey = "HKEY_USERS\" & gstrCurSID & "\Software\VB and VBA Program Settings\ZLSOFT"
        End If
        If objReg.GetRegValue(gstrZLSOFTRegKey & "\公共模块\升级助手", "助手跟踪日志级别", strRet) Then
            If strRet = "" Then
                llLogLevel = LogLevel_Trace
            Else
                llLogLevel = Val(strRet)
            End If
            Logger.Trace "读取" & gstrZLSOFTRegKey & "\公共模块\升级助手", "成功", "助手跟踪日志级别", llLogLevel
        Else
            llLogLevel = LogLevel_Trace
            Logger.Warn "读取" & gstrZLSOFTRegKey & "\公共模块\升级助手", "失败"
        End If
        strRet = ""
        If objReg.GetRegValue(gstrZLSOFTRegKey & "\公共模块\升级助手", "自然时间格式", strRet) Then
            Logger.Trace "读取" & gstrZLSOFTRegKey & "\公共模块\升级助手", "成功", "自然时间格式", strRet
        Else
            Logger.Warn "读取" & gstrZLSOFTRegKey & "\公共模块\升级助手", "失败"
        End If
        If LenB(strRet) = 0 Then
            strRet = "1"
        End If
        blnLoopTime = Val(strRet) = 0
    Else
        llLogLevel = Val(GetSetting("ZLSOFT", "公共模块\升级助手", "助手跟踪日志级别", LogLevel.LogLevel_Trace))
        blnLoopTime = Val(GetSetting("ZLSOFT", "公共模块\升级助手", "自然时间格式", 1)) = 0
    End If
    '设置内部组件的具体错误级别
    Logger.SetComponentLogLevel "ZLHelperMain.clsMutex", LogLevel_Info
    Logger.SetComponentLogLevel "ZLHelperMain.clsMemoryShareFP", LogLevel_Info
    Logger.SetComponentLogLevel "ZLHelperMain.clsException", LogLevel_Info
    Logger.SetComponentLogLevel "ZLHelperMain.frmMain", LogLevel_Info
    Logger.SetComponentLogLevel "ZLHelperMain.mdlRunas", LogLevel_Info
    '加载注册表设置
    Call LoadRegistryLogLevel
    Call Logger.OpenEx("ZLHelperMain_SessionID_" & glngWinSessionID & "_" & Decode(gstStartType, ST_Service, "Service", ST_SendServer, "SendServer", ST_SaveServer, "SaveServer", ST_Exit, "Exit"), , M_MAX_LOG_COUNT, , blnLoopTime)
    Logger.DebugEx "进程所属用户", GetProcessUserName
    Logger.DebugEx "进程是否Admin权限", IsProcessRunAsAdmin
    Logger.DebugEx "进程是否Admin组用户", IsAdministrator

    Call Logger.PopMethod("ZLHelperMain.mdlMain.InitEnv", gstrIP, gstrComputerName, gstrAppSOFT)
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.InitEnv") = 1 Then
        Resume
    End If
End Sub

'@方法    LoadRegistryLogLevel
'   加载注册表中的LogLevel设置
'@返回值
'
'@参数:
'@备注
'   HKEY_USERS\[SID]\Software\WOW6432Node\VB and VBA Program Settings\ZLSOFT\公共模块\升级助手
'       部件名1
'           default=日志级别
'           模块1=日志级别2
Private Sub LoadRegistryLogLevel()
    Dim objReg          As New clsRegistry
    Dim arrSubKey       As Variant
    Dim strParentKey    As String
    Dim i               As Long, j          As Long
    Dim arrSubKeyValue  As Variant
    
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlMain.LoadRegistryLogLevel")
    If LenB(gstrZLSOFTRegKey) <> 0 Then
        arrSubKey = objReg.GetAllSubKey(gstrZLSOFTRegKey & "\公共模块\升级助手")
        If TypeName(arrSubKey) <> "Empty" Then
            For i = LBound(arrSubKey) To UBound(arrSubKey)
                arrSubKeyValue = objReg.GetAllKeyValue(gstrZLSOFTRegKey & "\公共模块\升级助手\" & arrSubKey(i))
                If TypeName(arrSubKeyValue) <> "Empty" Then
                    For j = LBound(arrSubKeyValue) To UBound(arrSubKeyValue) Step 2
                        If arrSubKeyValue(j) = "" Then
                            Logger.SetComponentLogLevel arrSubKey(i) & "", Val(arrSubKeyValue(j + 1))
                            Logger.DebugEx arrSubKey(i) & "", Val(arrSubKeyValue(j + 1))
                        Else
                            Logger.SetComponentLogLevel arrSubKey(i) & "." & arrSubKeyValue(j), Val(arrSubKeyValue(j + 1))
                            Logger.DebugEx arrSubKey(i) & "." & arrSubKeyValue(j), Val(arrSubKeyValue(j + 1))
                        End If
                    Next
                End If
            Next
        End If
    End If
    Call Logger.PopMethod("ZLHelperMain.mdlMain.LoadRegistryLogLevel")
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.LoadRegistryLogLevel") = 1 Then
        Resume
    End If
End Sub

'--------------------------------------------------------------------------------------------------
'方法           CheckSystem
'功能           根据部件检查当前系统能否支持功能验证
'返回值
'入参列表:
'参数名         类型                    说明
'
'-------------------------------------------------------------------------------------------------
Public Sub CheckSystem()
    On Error GoTo ErrH
    Call Logger.PushMethod("ZLHelperMain.mdlMain.CheckSystem")
    If gobjFSO.FileExists(gstrAppSOFT & "\APPLY\ZLBRW.DLL") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppSOFT & "\APPLY\ZLBRW.DLL")) >= VerFull("10.35.0.130") Then
            gstrSystems = gstrSystems & ",100"
        Else
            Logger.Warn gstrAppSOFT & "\APPLY\ZLBRW.DLL", "版本号（小于10.35.0.130）", gobjFSO.GetFileVersion(gstrAppSOFT & "\APPLY\ZLBRW.DLL")
        End If
    End If
    If gobjFSO.FileExists(gstrAppSOFT & "\APPLY\ZLLISBRW.DLL") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppSOFT & "\APPLY\ZLLISBRW.DLL")) >= VerFull("10.35.0.140") Then
            gstrSystems = gstrSystems & ",2500"
        Else
            Logger.Warn gstrAppSOFT & "\APPLY\ZLLISBRW.DLL", "版本号（小于10.35.0.130）", gobjFSO.GetFileVersion(gstrAppSOFT & "\APPLY\ZLLISBRW.DLL")
        End If
    End If
    If gobjFSO.FileExists(gstrAppSOFT & "\ZLHEALTHSTART.EXE") Then
        If VerFull(gobjFSO.GetFileVersion(gstrAppSOFT & "\ZLHEALTHSTART.EXE")) >= VerFull("10.35.0.130") Then
            gstrSystems = gstrSystems & ",2700"
        Else
            Logger.Warn gstrAppSOFT & "\ZLHEALTHSTART.EXE", "版本号（小于10.35.0.130）", gobjFSO.GetFileVersion(gstrAppSOFT & "\ZLHEALTHSTART.EXE")
        End If
    End If
    gstrSystems = Mid(gstrSystems, 2)
    Call Logger.PopMethod("ZLHelperMain.mdlMain.CheckSystem", gstrSystems)
    Exit Sub
ErrH:
    If Logger.ErrCenter("ZLHelperMain.mdlMain.CheckSystem") = 1 Then
        Resume
    End If
End Sub
