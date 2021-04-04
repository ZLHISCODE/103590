VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCashBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Version As String

Private Sub Class_Initialize()
    Version = App.Major & "." & App.Minor & "." & App.Revision
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

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
    '------------------------------------------------
    '功能： 部件的初始化，并调用执行程序,将主程序的必要内容传递下来
    '参数：
    '   cnMain:主程序的数据库连接
    '返回：
    '------------------------------------------------
    Dim strTitle As String, frmList As Form, frmTemp As Form
    Set gcnOracle = cnMain
    If Not CheckValid Then Exit Sub
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    gstr单位名称 = GetUnitName()
    
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstr单位名称 = GetUnitName()
    gstrDBUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    '-------------------------------------------------
    '64       收费财务监控
    '65       票据使用监控
    '1502   借款管理
    '1504   票据入库管理
    '1506   收款轧帐管理
    '1507   财务组收款管理
    Select Case lngModul
    Case 1500
        strTitle = "收费财务监控"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmFinanceSupervise
        Call frmList.zlShowFinanceSupervise(frmMain, glngModul, gstrPrivs)
    Case 1501
        frmBillSupervise.Show , frmMain
    Case 1502
        strTitle = "人员借款管理"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmPersonLoanMgr
        frmList.ShowList lngModul, strTitle, frmMain
    Case 1504   '票据入库管理
        frmBillIn.Show , frmMain
    Case 1506 '收费轧帐管理
        strTitle = "收费轧帐管理"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmChargeRollingCurtainMgr
        Call frmList.zlShowChargeRollingCourtain(frmMain, 1506, gstrPrivs)
    Case 1507  '财务组收款管理
        frmFeeGroupManage.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
    '------------------------------------------------
    '功能： 部件的初始化，并调用执行程序,将主程序的必要内容传递下来
    '参数：
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '返回：
    '------------------------------------------------
    Dim strTitle As String, frmList As Object, frmTemp As Form
    Set gcnOracle = cnMain
    If Not CheckValid Then Exit Sub
    
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    gstr单位名称 = GetUnitName()
    
    glngSys = lngSys
    glngModul = lngModul
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstr单位名称 = GetUnitName()
    gstrDBUser = strDbUser
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrDec = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    '-------------------------------------------------
    '64       收费财务监控
    '65       票据使用监控
    '1502   借款管理
    '1504   票据入库管理
    Select Case lngModul
    Case 1500
        strTitle = "收费财务监控"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmFinanceSupervise
        Call frmList.BHShowList(lngModul, gstrPrivs, lngMain)
    Case 1501
        zlCommFun.ShowChildWindow frmBillSupervise.hWnd, lngMain
    Case 1502
        strTitle = "人员借款管理"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmPersonLoanMgr
        frmList.BHShowList lngModul, strTitle, lngMain
    Case 1504   '票据入库管理
        zlCommFun.ShowChildWindow frmBillIn.hWnd, lngMain
    Case 1506 '收费轧帐管理
        strTitle = "收费轧帐管理"
        For Each frmTemp In Forms
            If InStr(1, frmTemp.Caption, strTitle) > 0 Then
                frmTemp.Show
                frmTemp.ZOrder 0
                Exit Sub
            End If
        Next
        Set frmList = New frmChargeRollingCurtainMgr
        Call frmList.BHShowList(lngModul, gstrPrivs, lngMain)
    Case 1507  '财务组收款管理
        zlCommFun.ShowChildWindow frmFeeGroupManage.hWnd, lngMain
    End Select
End Sub
