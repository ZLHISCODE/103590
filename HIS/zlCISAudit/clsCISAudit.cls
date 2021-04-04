VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################

Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
    '******************************************************************************************************************
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '   strDBUser:当前登录数据库用户名
    '******************************************************************************************************************
    
    Set gcnOracle = cnMain

    gstrInputSeverName = frmMain.objLogin.ServerName
    gstrInputUser = frmMain.objLogin.InputUser
    gstrInputPwd = frmMain.objLogin.InputPwd

    ParamInfo.系统名称 = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    ParamInfo.产品名称 = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    gstrDBUser = strDBUser
    UserInfo.数据库用户 = strDBUser
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, ParamInfo.系统名称
        Exit Sub
    End If

    Set gfrmMain = frmMain
    Set gobjEmr = frmMain.mobjEmr
    
    UserInfo.模块权限 = GetPrivFunc(lngSys, lngModul)
    UserInfo.单位名称 = zlRegInfo("单位名称")
    gstrPrivs = UserInfo.模块权限
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    glngModul = lngModul
    ParamInfo.模块号 = lngModul
    ParamInfo.系统号 = lngSys
    ParamInfo.启用RIS = Val(zlDatabase.GetPara(255, lngSys)) = 1
    glngSys = lngSys
    
    gstrUserName = UserInfo.姓名
    SetDbUser UserInfo.用户名
    gstrDeptName = UserInfo.部门名称
    glngDeptId = UserInfo.部门ID
    gblnInsure = True
    
    '初始对象
    '------------------------------------------------------------------------------------------------------------------
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, ParamInfo.系统号, UserInfo.模块权限)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, ParamInfo.系统号, blnShowInTaskBar)
    Call gobjPath.InitCISPath(gcnOracle, frmMain, ParamInfo.系统号, UserInfo.模块权限)
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    
    Call InitSysPara
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand

    Select Case lngModul
    Case 1550                  '病案评分标准
        frm评分标准维护.Show , frmMain
    Case 1551                   '病案审查标准
        frmAuditItem.Show , frmMain
    Case 1560                   '电子病案审查
        If IsPrivs(gstrPrivs, "查阅审查病案") = False And IsPrivs(gstrPrivs, "查阅归档病案") = False And IsPrivs(gstrPrivs, "查阅抽查病案") = False Then
            ShowSimpleMsg "对不起，您没有查看电子病案审查模块的任何权限，请先到用户授权管理中设置！"
            Exit Sub
        End If
        frmCISAduit.Show , frmMain
    Case 1561                   '电子病案借阅
        frmCISBorrow.Show , frmMain
    Case 1562                   '电子病案评分
        frm病案评分.Show , frmMain
    Case 1563                   '病历质量查阅
        frmEPRAuditMan.Show , frmMain
    Case 1564                   '电子病案接收
        frm病案接收管理.Show , frmMain
    Case 1566
        'frmMedRecPrint.Show , frmMain  '电子病案打印  80880
        If CreateCISJOB() Then
            Call gobjJob.InitCISJob(gcnOracle, frmMain, glngSys, gstrPrivs, gobjEmr)
            gobjJob.ShowArchiveViewAndPrint frmMain, gstrPrivs          '病案查询打印
        End If
    End Select
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
    '******************************************************************************************************************
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '   strDBUser:当前登录数据库用户名
    '******************************************************************************************************************
    
    Set gcnOracle = cnMain
    
    ParamInfo.系统名称 = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    ParamInfo.产品名称 = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    UserInfo.数据库用户 = strDBUser
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, ParamInfo.系统名称
        Exit Sub
    End If

    Set gfrmMain = Nothing
    If Not objLogin Is Nothing Then
        Set gobjEmr = objLogin.mobjEmr
    End If
    
    UserInfo.模块权限 = GetPrivFunc(lngSys, lngModul)
    UserInfo.单位名称 = zlRegInfo("单位名称")
    gstrPrivs = UserInfo.模块权限
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    glngModul = lngModul
    ParamInfo.模块号 = lngModul
    ParamInfo.系统号 = lngSys
    ParamInfo.启用RIS = Val(zlDatabase.GetPara(255, lngSys)) = 1
    glngSys = lngSys
    
    gstrUserName = UserInfo.姓名
    SetDbUser UserInfo.用户名
    gstrDeptName = UserInfo.部门名称
    glngDeptId = UserInfo.部门ID
    gblnInsure = True
    
    '初始对象
    '------------------------------------------------------------------------------------------------------------------
    Call gobjKernel.InitCISKernel(gcnOracle, Nothing, ParamInfo.系统号, UserInfo.模块权限)
    Call gobjRichEPR.InitRichEPR(gcnOracle, Nothing, ParamInfo.系统号, True)
    Call gobjPath.InitCISPath(gcnOracle, Nothing, ParamInfo.系统号, UserInfo.模块权限)
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    Call InitSysPara
    
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHand

    Select Case lngModul
    Case 1550                  '病案评分标准
        zlCommFun.ShowChildWindow frm评分标准维护.hWnd, lngMain
    Case 1551                   '病案审查目录
        zlCommFun.ShowChildWindow frmAuditItem.hWnd, lngMain
    Case 1560                   '电子病案审查
        If IsPrivs(gstrPrivs, "查阅审查病案") = False And IsPrivs(gstrPrivs, "查阅归档病案") = False And IsPrivs(gstrPrivs, "查阅抽查病案") = False Then
            ShowSimpleMsg "对不起，您没有查看电子病案审查模块的任何权限，请先到用户授权管理中设置！"
            Exit Sub
        End If
        
        zlCommFun.ShowChildWindow frmCISAduit.hWnd, lngMain, True
    Case 1561                   '电子病案借阅
        zlCommFun.ShowChildWindow frmCISBorrow.hWnd, lngMain, True
    Case 1562                   '电子病案评分
        zlCommFun.ShowChildWindow frm病案评分.hWnd, lngMain, True
    Case 1563                   '病历质量查阅
        zlCommFun.ShowChildWindow frmEPRAuditMan.hWnd, lngMain, True
    Case 1564                   '电子病案接收
        zlCommFun.ShowChildWindow frm病案接收管理.hWnd, lngMain, True
    Case 1566                   '病案查询打印
        'zlCommFun.ShowChildWindow frmMedRecPrint.hWnd, lngMain, True
        If CreateCISJOB() Then
            zlCommFun.ShowChildWindow gobjJob.GetArchiveViewAndPrintFrom.hWnd, lngMain, True
            Call gobjJob.InitCISJob(gcnOracle, Nothing, glngSys, gstrPrivs, gobjEmr)
            gobjJob.ShowArchiveViewAndPrint Nothing, gstrPrivs
        End If
    End Select
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
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
    blnChildren = blnChildren And gobjPath.CloseWindows()
    CloseWindows = blnChildren And (Forms.count = 0)
    
End Function

Private Sub Class_Initialize()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeBlue
        
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub Class_Terminate()

    On Error Resume Next

'    Unload frmPubIcons
    Unload frmPubResource
End Sub


