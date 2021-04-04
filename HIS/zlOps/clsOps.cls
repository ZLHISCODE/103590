VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################

Public Version As String

'######################################################################################################################

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
    '******************************************************************************************************************
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '返回：
    '******************************************************************************************************************
    Set gcnOracle = cnMain
    
    '获取初始值
    '------------------------------------------------------------------------------------------------------------------
    ParamInfo.系统名称 = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    ParamInfo.产品名称 = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrSysName = ParamInfo.系统名称
    gblnShowInTaskBar = blnShowInTaskBar
    gblnInsure = True
    
    gstrDBUser = strDBUser
    UserInfo.数据库用户 = strDBUser
    
    Call GetUserInfo
    
    UserInfo.模块权限 = GetPrivFunc(lngSys, lngModul)
    UserInfo.单位名称 = zlRegInfo("单位名称")
    ParamInfo.模块号 = lngModul
    ParamInfo.系统号 = lngSys
    
    '初始对象
    '------------------------------------------------------------------------------------------------------------------
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, ParamInfo.系统号, UserInfo.模块权限)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, ParamInfo.系统号, blnShowInTaskBar)
    
    
    '读取系统参数
    '------------------------------------------------------------------------------------------------------------------
    Call InitSysPara
    
    '执行模块
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    Select Case lngModul
    Case 1801               '手术方案设置
        frmOpsScheme.Show , frmMain
    Case 1804               '手术室工作站
        Call gclsInsure.InitOracle(gcnOracle)
        frmOpsStation.Show , frmMain
    End Select
    
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
    '******************************************************************************************************************
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '返回：
    '******************************************************************************************************************
    Set gcnOracle = cnMain
    
    '获取初始值
    '------------------------------------------------------------------------------------------------------------------
    ParamInfo.系统名称 = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    ParamInfo.产品名称 = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrSysName = ParamInfo.系统名称
    gblnShowInTaskBar = blnShowInTaskBar
    gblnInsure = True
    
    gstrDBUser = strDBUser
    UserInfo.数据库用户 = strDBUser
    
    Call GetUserInfo
    
    UserInfo.模块权限 = GetPrivFunc(lngSys, lngModul)
    UserInfo.单位名称 = zlRegInfo("单位名称")
    ParamInfo.模块号 = lngModul
    ParamInfo.系统号 = lngSys
    
    '初始对象
    '------------------------------------------------------------------------------------------------------------------
    Call gobjKernel.InitCISKernel(gcnOracle, Nothing, ParamInfo.系统号, UserInfo.模块权限)
    Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, ParamInfo.系统号, blnShowInTaskBar)
    
    '读取系统参数
    '------------------------------------------------------------------------------------------------------------------
    Call InitSysPara
    
    '执行模块
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    Select Case lngModul
    Case 1801               '手术方案设置
        zlCommFun.ShowChildWindow frmOpsScheme.hWnd, lngMain
    Case 1804               '手术室工作站
        Call gclsInsure.InitOracle(gcnOracle)
        zlCommFun.ShowChildWindow frmOpsStation.hWnd, lngMain, True
    End Select
    
End Sub

Public Function CloseWindows() As Boolean
    '******************************************************************************************************************
    '功能:关闭所有子窗口
    '******************************************************************************************************************
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    
    blnChildren = gobjRichEPR.CloseWindows()
    blnChildren = blnChildren And gobjKernel.CloseWindows()
    CloseWindows = blnChildren And (Forms.Count = 0)

End Function

Private Sub Class_Initialize()
    Version = App.major & "." & App.minor & "." & App.Revision
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
    
End Sub

