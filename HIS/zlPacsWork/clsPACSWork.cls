VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPACSWork"
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
    
    Call gclsInsure.InitOracle(gcnOracle)
    Call InitCISCore(gcnOracle)
    Call InitTechCore(gcnOracle)
    
    '-------------------------------------------------
    'PASS接口初始化:只初始化一次
    '因为几个模块可能同时使用,且共用变量gblnPass，因此还需要在具体模块中判断权限
    '-------------------------------------------------
    On Error Resume Next
    Select Case lngModul
        Case 1056 '影像检查项目
            frmRadLists.Show , gfrmMain
        Case 1057 '影像设备目录
            frmPacsDev.Show , gfrmMain
        Case 1207 '影像技师工作站
            frmPACStation.Show , gfrmMain
'            frmVideoStation.Show , gfrmMain
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
End Sub

