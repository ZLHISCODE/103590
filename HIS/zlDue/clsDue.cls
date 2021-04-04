VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDue"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, FrmMain As Object, strDbUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngSys : 系统编号
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '返回：
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    gstrDBUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    Set gfrmMain = FrmMain
    Call 初始小数位数
    
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    glngSys = lngSys
    gstrUnitName = GetUnitName
    
    '初始系统参数值
    Call zlinitSystemPara
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1025
            frm供应商管理.Show , FrmMain
        Case 1322
            frm应付款管理.Show , FrmMain
        Case 1323
            frm付款管理.Show , FrmMain
        Case 1324
            frm清单管理.Show , FrmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, strDbUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngSys : 系统编号
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '返回：
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    gstrDBUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    Set gfrmMain = Nothing
    Call 初始小数位数
    
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    glngSys = lngSys
    gstrUnitName = GetUnitName
    
    '初始系统参数值
    Call zlinitSystemPara
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1025
            zlCommFun.ShowChildWindow frm供应商管理.hwnd, lngMain
        Case 1322
            zlCommFun.ShowChildWindow frm应付款管理.hwnd, lngMain
        Case 1323
            zlCommFun.ShowChildWindow frm付款管理.hwnd, lngMain
        Case 1324
            zlCommFun.ShowChildWindow frm清单管理.hwnd, lngMain
    End Select
End Sub

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function



