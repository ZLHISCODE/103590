VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecipeAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------
'说明: 为ZLHIS导航台开放的接口类
'--------------------------------------------------------------------------------
Option Explicit

Private mstrVersion As String

Public Property Get Version() As String
    mstrVersion = App.Major & "." & App.Minor & "." & App.Revision
    Version = mstrVersion
End Property

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, ByRef cnMain As ADODB.Connection, ByRef frmMain As Object, ByRef strDbUser As String)
'功能:ZLHIS导航台调用模块的方法
'参数:
'  lngSys：系统号
'  lngModul：模块号
'  cnMain：Oracle数据库连接对象
'  frmMain：导航台窗体对象
'  strDBUser：数据库登录用户名

    glngSys = lngSys
    glngModule = lngModul
    UserInfo.用户名 = strDbUser
    Set gcnOracle = cnMain
        
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")       '获取软件系统名称
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")   '获取产品名称
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrUnitName = zl9ComLib.GetUnitName    '获取注册的单位名称信息
    
    Select Case glngModule
        Case enuModule.门诊处方审查_1351
            frmRAOutPatient.Show , frmMain
            
        Case enuModule.住院药嘱审查_1352
            frmRAInPatient.Show , frmMain
            
        Case enuModule.处方审查项目_1353
            frmRAItems.Show , frmMain
        
        Case enuModule.处方审查条件_1354
            frmRASetup.Show , frmMain
        
        Case enuModule.处方审查统计_1355
            frmRAStatistics.Show , frmMain
    End Select
    
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, ByRef cnMain As ADODB.Connection, ByRef lngMain As Long, ByRef strDbUser As String)
'功能:BH导航台调用模块的方法
'参数:
'  lngSys：系统号
'  lngModul：模块号
'  cnMain：Oracle数据库连接对象
'  lngMain：导航台窗体对象的句柄号
'  strDBUser：数据库登录用户名

    glngSys = lngSys
    glngModule = lngModul
    UserInfo.用户名 = strDbUser
    Set gcnOracle = cnMain
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrUnitName = zl9ComLib.GetUnitName
    
    Select Case glngModule
        Case enuModule.门诊处方审查_1351
            OS.ShowChildWindow frmRAOutPatient.hwnd, lngMain
        
        Case enuModule.住院药嘱审查_1352
            OS.ShowChildWindow frmRAInPatient.hwnd, lngMain
        
        Case enuModule.处方审查项目_1353
            OS.ShowChildWindow frmRAItems.hwnd, lngMain
        
        Case enuModule.处方审查条件_1354
            OS.ShowChildWindow frmRASetup.hwnd, lngMain
            
        Case enuModule.处方审查统计_1355
            OS.ShowChildWindow frmRAStatistics.hwnd, lngMain
    End Select
End Sub

Public Function CloseWindows() As Boolean
'功能:关闭部件里所有的窗体
'返回:True成功；False失败

    Dim frmThis As Form
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Initialize()
    gintHoursRecipe = 24
End Sub
