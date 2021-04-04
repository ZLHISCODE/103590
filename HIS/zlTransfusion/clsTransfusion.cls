VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransfusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Public gobjKernel As New clsCISKernel       '临床核心部件
Public Version As String

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名

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
    
    Call InitSysPar
    gstrPrivs = GetPrivFunc(lngSys, 1252)
    If gstrPrivs <> "" Then
        gstrPrivs = GetPrivFunc(lngSys, lngModul) & ";门诊医嘱;"
    Else
        gstrPrivs = GetPrivFunc(lngSys, lngModul)
    End If
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub
    
    gstrPrivs = ";" & gstrPrivs & ";"
    
    gstrUnitName = GetUnitName
'    Call gobjKernel.InitCISKernel(gcnOracle, frmMain, glngSys, gstrPrivs)
    '调用菜单应用模块
    On Error Resume Next
    Select Case lngModul
        Case 1264 '输液排队
            ' 开发时屏蔽 ,frmMain
            frmTransfusion.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名

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
    
    Call InitSysPar
    gstrPrivs = GetPrivFunc(lngSys, 1252)
    If gstrPrivs <> "" Then
        gstrPrivs = GetPrivFunc(lngSys, lngModul) & ";门诊医嘱;"
    Else
        gstrPrivs = GetPrivFunc(lngSys, lngModul)
    End If
    If gstrPrivs = "" Then MsgBox "你不具备使用该模块的权限！", vbExclamation, gstrSysName: Exit Sub
    gstrUnitName = GetUnitName
    
    gstrPrivs = ";" & gstrPrivs & ";"
    
    'Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs)
    
    '调用菜单应用模块
    On Error Resume Next
    Select Case lngModul
        Case 1264 '输液排队
            zlCommFun.ShowChildWindow frmTransfusion.hwnd, lngMain, False
    End Select
End Sub
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
End Sub

