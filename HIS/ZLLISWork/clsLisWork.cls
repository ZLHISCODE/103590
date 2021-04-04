VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

'中药配方界面所需要的病人信息
Public Type TYPE_PatiInfoEx
        lng病人ID As Long
        lng主页ID As Long '主页ID
        str挂号单 As String '挂号单号
        str性别 As String  '病人性别
        int婴儿 As Integer
        int险类 As Integer '医保病人的险类
        bln医保 As Boolean '是否医保病人
        lng病人科室ID As Long '用于确定中药配方的缺省药房
End Type

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
    Dim strTmp As String
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    '取就诊卡号长度
'    strTmp = zlDatabase.GetPara("票据号码长度", lngSys, , "7|7|7|7|7")
'    gbytCardNOLen = Val(Split(strTmp, "|")(4))
    
    gstrDBUser = strDBUser
        
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
   
    Set gfrmMain = frmMain
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    glngModul = lngModul
    glngSys = lngSys
    
    gstrUnitName = GetUnitName
    gstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    gblnInsure = (UCase(GetSetting("ZLSOFT", "公共全局", "是否支持医保", "")) = UCase("Yes"))
    gintInsure = Val(GetSetting("ZLSOFT", "公共全局", "医保类别", 0))
    '执行后自动审核
'    gbln执行后审核 = Val(zlDatabase.GetPara(81, glngSys)) <> 0
    '-------------------------------------------------
    
     '初始化核心部件
    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, frmMain, glngSys, False)
    
    On Error Resume Next
    '初始批号处理参数
    gblnManualPH = True
    gintNumberPH = 9999
    
    Select Case lngModul
    Case 1208   '检验技师工作站
        If gblnInsure Then
            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
        Else
            Call gclsInsure.InitOracle(gcnOracle)
        End If
        frmLabMain.Show , frmMain
    Case 1209   '质控数据查询
        frmQCHistory.Show , frmMain
'        FrmQualityDataQuery.Show , frmMain
    Case 1210   '今日质控管理
        frmQCTodayList.Show , frmMain
'        frmPatientHistoryQuery.Show , frmMain
    Case 1211   '检验采集工作站
        frmLabSampling.Show , frmMain
    Case 1212   '检验标本登记
        frmLabSampleRegister.Show , frmMain
    Case 1213   '仪器比对
        frmQCContrast.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String, Optional ByVal objLogin As Object)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
    Dim strTmp As String
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    '取就诊卡号长度
'    strTmp = zlDatabase.GetPara("票据号码长度", lngSys, , "7|7|7|7|7")
'    gbytCardNOLen = Val(Split(strTmp, "|")(4))
    
    gstrDBUser = strDBUser
        
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set gfrmMain = Nothing
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    glngModul = lngModul
    glngSys = lngSys
    
    gstrUnitName = GetUnitName
    gstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    gblnInsure = (UCase(GetSetting("ZLSOFT", "公共全局", "是否支持医保", "")) = UCase("Yes"))
    gintInsure = Val(GetSetting("ZLSOFT", "公共全局", "医保类别", 0))
    '执行后自动审核
'    gbln执行后审核 = Val(zlDatabase.GetPara(81, glngSys)) <> 0
    '-------------------------------------------------
    
     '初始化核心部件
    Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs)
    Call gobjRichEPR.InitRichEPR(gcnOracle, objLogin, glngSys, False)
    Set gobjEmr = objLogin.mobjEMR
    On Error Resume Next
    '初始批号处理参数
    gblnManualPH = True
    gintNumberPH = 9999
    
    Select Case lngModul
    Case 1208   '检验技师工作站
        If gblnInsure Then
            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
        Else
            Call gclsInsure.InitOracle(gcnOracle)
        End If
        zlCommFun.ShowChildWindow frmLabMain.hWnd, lngMain, True
    Case 1209   '质控数据查询
        zlCommFun.ShowChildWindow frmQCHistory.hWnd, lngMain
    Case 1210   '今日质控管理
        zlCommFun.ShowChildWindow frmQCTodayList.hWnd, lngMain
    Case 1211   '检验采集工作站
        zlCommFun.ShowChildWindow frmLabSampling.hWnd, lngMain, True
    Case 1212   '检验标本登记
        zlCommFun.ShowChildWindow frmLabSampleRegister.hWnd, lngMain
    Case 1213   '仪器比对
        zlCommFun.ShowChildWindow frmQCContrast.hWnd, lngMain
    End Select
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    
    Call gclsInsure.CloseWindows
    Call gobjRichEPR.CloseWindows
    Call gobjKernel.CloseWindows
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

