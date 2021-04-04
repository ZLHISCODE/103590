VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    
    gstrDBUser = strDBUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Not CheckValid Then Exit Sub

    Set gfrmMain = frmMain
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    glngSys = lngSys
    
    gstrUnitName = GetUnitName
    gblnInsure = (UCase(GetSetting("ZLSOFT", "公共全局", "是否支持医保", "")) = UCase("Yes"))
    gintInsure = Val(GetSetting("ZLSOFT", "公共全局", "医保类别", 0))
    '-------------------------------------------------
    Select Case lngModul
        Case 1530           '用户表格定义
            frmDefTable.Show , frmMain
        Case 1531           '查询图形设置
            frmPicture.Show , frmMain
        Case 1532           '专家介绍清单
            frmDoctor.Show , frmMain
        Case 1533           '播放广告设置
            frmAdvice.Show , frmMain
        Case 1534           '查询页面定义
            frmDefQuery.Show , frmMain
        Case 1535           '查询目录规化
            frmDefTree.Show , frmMain
        Case 1536           '查询内容显示
                        '医保接口
            If gblnInsure Then
                If Not gclsInsure.InitInsure(gcnOracle) Then gblnInsure = False
            End If
            
            frmMainQuery.Show , frmMain
        Case 1537           '咨询参数设置
            frmParameter.Show 1, frmMain
        Case 1538           '自助挂号管理
            Call InitLocPar
            Call InitSysPar
            On Error Resume Next
            frmselectinfo.Show , frmMain
    End Select
End Sub

Private Sub Class_Terminate()
    '
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

