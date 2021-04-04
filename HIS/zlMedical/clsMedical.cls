VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMedical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
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
'    gblnInsure = (UCase(GetSetting("ZLSOFT", "公共全局", "是否支持医保", "")) = UCase("Yes"))
'    gintInsure = Val(GetSetting("ZLSOFT", "公共全局", "医保类别", 0))
    gblnInsure = True
    
    '-------------------------------------------------
    On Error Resume Next
    
    Call InitCISCore(gcnOracle)
    
    Select Case lngModul
    Case 1850               '体检类型
        frmKind.Show , frmMain
    Case 1851               '体检诊断建议
        frmDiagnoseAdvice.Show , frmMain
    Case 1852               '体检项目管理
        frmMedicalItems.Show , frmMain
    Case 1860               '体检预约申请
        frmSchedule.Show , frmMain
    Case 1861               '体检工作管理
'
'        If gblnInsure Then
'            If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
'        Else
'            Call gclsInsure.InitOracle(gcnOracle)
'        End If
        
        frmMedicalStation.Show , frmMain
    Case 1862               '体检团体结算
        frmBalance.Show , frmMain
    Case 1863               '体检随访管理
        frmLaterVisit.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名

    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
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
    gblnInsure = True
    
    '-------------------------------------------------
    On Error Resume Next
    
    Call InitCISCore(gcnOracle)
    
    Select Case lngModul
    Case 1850               '体检类型
        zlCommFun.ShowChildWindow frmKind.hWnd, lngMain
    Case 1851               '体检诊断建议
        zlCommFun.ShowChildWindow frmDiagnoseAdvice.hWnd, lngMain
    Case 1852               '体检项目管理
        zlCommFun.ShowChildWindow frmMedicalItems.hWnd, lngMain
    Case 1860               '体检预约申请
        zlCommFun.ShowChildWindow frmSchedule.hWnd, lngMain
    Case 1861               '体检工作管理
        zlCommFun.ShowChildWindow frmMedicalStation.hWnd, lngMain, True
    Case 1862               '体检团体结算
        zlCommFun.ShowChildWindow frmBalance.hWnd, lngMain
    Case 1863               '体检随访管理
        zlCommFun.ShowChildWindow frmLaterVisit.hWnd, lngMain
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
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Load frmPrint
End Sub


Private Sub Class_Terminate()
    Unload frmPrint
End Sub
