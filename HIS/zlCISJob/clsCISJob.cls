VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String
Public Type PatiInfo
    排序 As String
    状态 As Integer '病案主页.状态
    性质 As Integer '0-普通住院病人,1-门诊留观病人,2-住院留观病人
    住院号 As String
    床号 As String
    主页ID As Long
    病区ID As Long
    科室ID As Long
    产科 As Boolean
    入院日期 As Date
    出院日期 As Date
    路径状态 As Integer
    病案状态 As Integer
    数据转出 As Boolean
    险类 As Integer
    结清 As Boolean
End Type

Public Type TYPE_CATE
    strName As String
    lngBeginRow As Long
    lngEndRow As Long
End Type
Private mobjArchiveView As frmArchiveView
Private mobjOneArchiveView As frmArchiveView '显示某一个病人的电子病案
Private mobjArchiveViewAndPrint As frmArchiveViewAndPrint

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名
    Dim objPlugIn As Object
    Dim strMsg As String, intMode As Integer
    On Error GoTo ErrHand

    Set gfrmMain = frmMain
    'If Not CheckValid Then Exit Sub
    On Error Resume Next
    Set gobjEmr = gfrmMain.mobjEMR  '新病历需要部件
    If err.Number <> 0 Then
        Set gobjEmr = Nothing
        err.Clear
    End If
    
    On Error GoTo ErrHand
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDBUser
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub
    gstrUnitName = GetUnitName
    
    If frmMain Is Nothing Then  '医保接口调试调用
        gblnShowInTaskBar = blnShowInTaskBar
    Else
        gblnShowInTaskBar = blnShowInTaskBar Or frmMain.frmHide
    End If
    
    Call gclsInsure.InitOracle(gcnOracle)
    
    '初始化核心部件
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, lngModul, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, blnShowInTaskBar)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, lngModul, gstrPrivs, gobjEmr)
    
    '判断临床工作站必须使用插件
    If lngModul >= 1260 And lngModul <= 1265 Then
        If zlDatabase.GetPara(208, 100, , 0) = 1 Then
            On Error Resume Next
            Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
            err.Clear: On Error GoTo 0
            If objPlugIn Is Nothing Then
                MsgBox "当前系统设置为必须使用插件(zlPlugIn)，本机缺少该插件，请联系管理员。", vbInformation, gstrSysName
                Exit Sub
            Else
                Set objPlugIn = Nothing
            End If
        End If
    End If
    '检查新版护士工作站是否融入整体护理功能
    If lngModul = P新版护士站 Or lngModul = p住院医生站 Then
        '检查是否启用了整体护理(满足条件：接口部件存在、服务器设置正确、该用户登录成功)
        gbln启用整体护理接口 = False
        If InitNurseIntegrate = True Then
            gbln启用整体护理接口 = gobjNurseIntegrate.IsUseIntergrate(strMsg, intMode)
            If gbln启用整体护理接口 = False Then
                If intMode = 1 Then
                    MsgBox "整体护理服务器连接失败，将无法进行整体护理相关操作！" & vbCrLf & "详细信息：" & strMsg, vbInformation, gstrSysName
                ElseIf intMode = 2 Then
                    'IE版本不对不提示（德阳部分电脑未装IE11，也不使用整体护理，此处不再提示）
'                    MsgBox strMsg, vbInformation, gstrSysName
                End If
            Else
                If lngModul = P新版护士站 Then
                    If gobjNurseIntegrate.UserLogin(strMsg) = False Then
                        MsgBox "你登录整理护理服务器失败，将无法进行整体护理相关操作！" & vbCrLf & "详细信息：" & strMsg, vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
    End If
    '调用菜单应用模块
    If lngModul = 1260 Or lngModul = 1261 Or lngModul = 1262 Or lngModul = 1263 Or _
        lngModul = 1265 Or lngModul = 1266 Or lngModul = 1267 Or lngModul = 1268 Then
        Call InitSysPar
    End If
    On Error Resume Next
    Select Case lngModul
        Case 1260 '门诊医生工作站
            Call InitCommunity
            frmOutDoctorStation.Show , frmMain
        Case 1261 '住院医生工作站
            frmInDoctorStation.ShowMe frmMain
        Case 1262 '住院护士工作站
            frmInNurseStation.Show , frmMain
        Case 1263 '医技工作站
            frmTechnicStation.Show , frmMain
        Case 1265 '新版住院护士工作站
            frmSublimeInNurseStation.Show , frmMain
        Case 1071, 1072, 1073, 1074, 1075, 1076, 1273, 1274, 1277
            Call gobjRichEPR.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1270: Call gobjKernel.ShowDiagHelp(vbModeless, frmMain, , blnShowInTaskBar)
        Case 1271: Call gobjKernel.ShowClincHelp(vbModeless, frmMain, , blnShowInTaskBar)
        Case 1078, 1275, 1083, 1272 '临床路径管理,临床路径跟踪,门诊路径管理,门诊路径跟踪
            Call gobjPath.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1266
            frmExamineKSS.Show , frmMain
        Case 1267
            frmExamineOPS.Show , frmMain
        Case 1268
            frmExamineTransfuse.ShowMe frmMain, 2
        Case 1269
            frmKssStatistics.Show , frmMain
        Case 1284 '危急值查询
            Call gobjKernel.zlExecuteCode(glngSys, lngModul, frmMain, gstrPrivs)
        Case 1910  '血液接收登记
            If gbln血库系统 = False Then
                MsgBox "血库系统未正常安装或未勾选系统参数启用血库系统，请与系统管理员联系。", vbInformation, gstrSysName
                Exit Sub
            End If
            If InitObjBlood = True Then
                Call gobjPublicBlood.zlShowBloodReceives(frmMain, 100, lngModul, 0)
            End If
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名
    Dim objPlugIn As Object
    Dim strMsg As String
    Dim intMode As Integer
    
    Set gfrmMain = Nothing
    gblnShowInTaskBar = True
    'If Not CheckValid Then Exit Sub
    
    If Not objLogin Is Nothing Then
        Set gobjEmr = objLogin.mobjEMR  '新病历需要部件
        If err.Number <> 0 Then
            Set gobjEmr = Nothing
            err.Clear
        End If
    End If
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    glngModul = lngModul
    glngSys = lngSys
    
    gstrDBUser = strDBUser
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub
    gstrUnitName = GetUnitName
    
    Call gclsInsure.InitOracle(gcnOracle)
    
    '初始化核心部件
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, Nothing, lngModul, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, objLogin, glngSys, gblnShowInTaskBar)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, Nothing, lngModul, gstrPrivs)
    
    '判断临床工作站必须使用插件
    If lngModul >= 1260 And lngModul <= 1265 Then
        If zlDatabase.GetPara(208, 100, , 0) = 1 Then
            On Error Resume Next
            Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
            err.Clear: On Error GoTo 0
            If objPlugIn Is Nothing Then
                MsgBox "当前系统设置为必须使用插件(zlPlugIn)，本机缺少该插件，请联系管理员。", vbInformation, gstrSysName
                Exit Sub
            Else
                Set objPlugIn = Nothing
            End If
        End If
    End If
    '检查新版护士工作站是否融入整体护理功能
    If lngModul = P新版护士站 Or lngModul = p住院医生站 Then
        '检查是否启用了整体护理(满足条件：接口部件存在、服务器设置正确)
        gbln启用整体护理接口 = False
        If InitNurseIntegrate = True Then
            gbln启用整体护理接口 = gobjNurseIntegrate.IsUseIntergrate(strMsg, intMode)
            If gbln启用整体护理接口 = False Then
                If intMode = 1 Then
                    MsgBox "整体护理服务器连接失败，将无法进行整体护理相关操作！" & vbCrLf & "详细信息：" & strMsg, vbInformation, gstrSysName
                ElseIf intMode = 2 Then
                    'IE版本不对不提示（德阳部分电脑未装IE11，也不使用整体护理，此处不再提示）
'                    MsgBox strMsg, vbInformation, gstrSysName
                End If
            Else
                If lngModul = P新版护士站 Then
                    If gobjNurseIntegrate.UserLogin(strMsg) = False Then
                        MsgBox "你登录整理护理服务器失败，将无法进行整体护理相关操作！" & vbCrLf & "详细信息：" & strMsg, vbInformation, gstrSysName
                    End If
                End If
            End If
        End If
    End If
    '调用菜单应用模块
    If lngModul = 1260 Or lngModul = 1261 Or lngModul = 1262 Or lngModul = 1263 Or _
        lngModul = 1265 Or lngModul = 1266 Or lngModul = 1267 Or lngModul = 1268 Then
        Call InitSysPar
    End If
    On Error Resume Next
    Select Case lngModul
        Case 1260 '门诊医生工作站
            Call InitCommunity
            Call OS.ShowChildWindow(frmOutDoctorStation.hwnd, lngMain, True)
        Case 1261 '住院医生工作站
            Call OS.ShowChildWindow(frmInDoctorStation.hwnd, lngMain, True)
        Case 1262 '住院护士工作站
            Call OS.ShowChildWindow(frmInNurseStation.hwnd, lngMain, True)
        Case 1263 '医技工作站
            Call OS.ShowChildWindow(frmTechnicStation.hwnd, lngMain, True)
        Case 1265 '新版住院护士工作站
            Call OS.ShowChildWindow(frmSublimeInNurseStation.hwnd, lngMain, True)
        Case 1071, 1072, 1073, 1074, 1075, 1076, 1273, 1274
            Call gobjRichEPR.zlExecuteCode(glngSys, lngModul, lngMain, gstrPrivs)
        Case 1270: Call gobjKernel.ShowDiagHelp(vbModeless, Nothing, , gblnShowInTaskBar)
        Case 1271: Call gobjKernel.ShowClincHelp(vbModeless, Nothing, , gblnShowInTaskBar)
        Case 1078, 1275 '临床路径管理,临床路径跟踪
            Call gobjPath.zlExecuteCode(glngSys, lngModul, lngMain, gstrPrivs)
        Case 1266
            Call OS.ShowChildWindow(frmExamineKSS.hwnd, lngMain, True)
        Case 1267
            Call OS.ShowChildWindow(frmExamineOPS.hwnd, lngMain, True)
        Case 1268
            Call OS.ShowChildWindow(frmExamineTransfuse.hwnd, lngMain, True)
        Case 1269
            Call OS.ShowChildWindow(frmKssStatistics.hwnd, lngMain, True)
        Case 1910 '血液接收登记
            If gbln血库系统 = False Then
                MsgBox "血库系统未正常安装或未勾选系统参数启用血库系统，请与系统管理员联系。", vbInformation, gstrSysName
                Exit Sub
            End If
            If InitObjBlood(True) = True Then
                Call gobjPublicBlood.zlShowBloodReceives(lngMain, 100, lngModul, 0)
            End If
    End Select
End Sub

Private Sub InitCommunity()
'功能：入口初始化社区接口对象
    If gobjCommunity Is Nothing Then
        On Error Resume Next
        Set gobjCommunity = CreateObject("zlCommunity.clsCommunity")
        err.Clear: On Error GoTo 0
    End If
    
    If Not gobjCommunity Is Nothing Then
        If Not gobjCommunity.Initialize(gcnOracle) Then
            Set gobjCommunity = Nothing
        End If
    End If
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    Call gobjEmr.CloseForms
    blnChildren = gobjRichEPR.CloseWindows()
    blnChildren = blnChildren And gobjKernel.CloseWindows()
    blnChildren = blnChildren And gobjPath.CloseWindows()
    '调用接口关闭移动控件窗体,负责无法释放进程
    If gbln启用整体护理接口 = True Then
        If InitNurseIntegrate = True Then
            Call gobjNurseIntegrate.UnloadWebKixt
        End If
    End If
    CloseWindows = blnChildren And (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Class_Terminate()
    If Not gobjCommunity Is Nothing Then
        Call gobjCommunity.Terminate
        Set gobjCommunity = Nothing
    End If
    If Not gobjPass Is Nothing Then
        Set gobjPass = Nothing
    End If
    On Error Resume Next
    Unload mobjOneArchiveView
    Unload mobjArchiveView
    Unload mobjArchiveViewAndPrint
    Set mobjOneArchiveView = Nothing
    Set mobjArchiveView = Nothing
    Set mobjArchiveViewAndPrint = Nothing
End Sub

Public Sub ShowArchive(ByRef objForm As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal blnModal As Boolean = False)
'功能：调用电子病案查阅功能，单个病人查阅
    If mobjOneArchiveView Is Nothing Then Set mobjOneArchiveView = New frmArchiveView
    Call mobjOneArchiveView.ShowArchive(objForm, lng病人ID, lng主页ID, blnModal)
End Sub

Public Function GetArchiveOneFrom() As Object
'功能：调用电子病案查阅功能，单个病人查阅的窗体
    Set GetArchiveOneFrom = mobjOneArchiveView
End Function

Public Function GetArchiveFrom() As Object
'功能：调用电子病案查阅功能，嵌入式获取窗体对象
    Set mobjArchiveView = New frmArchiveView
    mobjArchiveView.BorderStyle = FormBorderStyleConstants.vbBSNone '设置为无边框
    mobjArchiveView.Caption = mobjArchiveView.Caption       '重点是这一句
    Set GetArchiveFrom = mobjArchiveView
End Function

Public Sub ShowArchiveViewAndPrint(ByVal frmParent As Object, ByVal strPrivs As String)
'功能：调用病案查阅打印功能
    If mobjArchiveViewAndPrint Is Nothing Then Set mobjArchiveViewAndPrint = New frmArchiveViewAndPrint
    mobjArchiveViewAndPrint.ShowArchive frmParent, strPrivs
End Sub

Public Function GetArchiveViewAndPrintFrom() As Object
'功能：获取病案查阅打印的窗体
    If mobjArchiveViewAndPrint Is Nothing Then Set mobjArchiveViewAndPrint = New frmArchiveViewAndPrint
    Set GetArchiveViewAndPrintFrom = mobjArchiveViewAndPrint
End Function

Public Sub zlArchiveRefresh(ByVal lng病人ID As Long, ByVal lng就诊ID As Long)
'功能：刷新电子病案查阅界面
    Call mobjArchiveView.zlRefresh(lng病人ID, lng就诊ID)
End Sub

Public Sub ViewEPRReport(frmParent As Object, ByVal lng报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(frmParent, lng报告ID, CanPrint)
End Sub

Public Sub ViewPACSImage(frmParent As Object, ByVal lng医嘱ID As Long, ByVal blnMoved As Boolean)
'功能：PACS观片处理
    If CreateObjectPacs(gobjPublicPacs) Then
        Call gobjPublicPacs.ShowImage(lng医嘱ID, frmParent, blnMoved)
    End If
End Sub

Public Sub InitCISJob(cnMain As ADODB.Connection, frmMain As Object, ByVal lngSys As Long, ByVal strMainPrivs As String, Optional ByRef objEmr As Object)
    Set gfrmMain = frmMain
    Set gcnOracle = cnMain
    
    glngSys = lngSys

    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrUnitName = GetUnitName
    gstrPrivs = strMainPrivs
    If Not objEmr Is Nothing Then
        Set gobjEmr = objEmr
    End If
    '初始化操作员信息
    Call GetUserInfo
    If gstrDBUser = "" Then gstrDBUser = UserInfo.用户名   '病案查阅打印
    '初始化系统参数:可分散到具体的GetForm中去初始化
    Call InitSysPar
    '初始化核心部件
    Call gobjKernel.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, 0, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
    Call gobjPath.InitCommon(gcnOracle, glngSys, gstrNodeNo, frmMain, 0, gstrPrivs, gobjEmr)
End Sub

Public Function InitCISJobPar() As Boolean
    InitCISJobPar = InitSysPar
End Function
