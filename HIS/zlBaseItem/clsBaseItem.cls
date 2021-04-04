VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBaseItem"
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

    '关闭专项记帐单的窗口，因为它在另一个工程中，所以要单独处理
    If Not gobjCustAcc Is Nothing Then
        CloseWindows = CloseWindows And gobjCustAcc.CloseWindows
    End If
End Function

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDbUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '返回：
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gblnMyStyle = zlDatabase.GetPara("使用个性化风格") = "1"
    gstr单位名称 = GetUnitName()
    gstrDbUser = strDbUser
    Call GetUserInfo
    
    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    If Not CheckValid Then Exit Sub
    '-------------------------------------------------
    '1000   数据转移管理
    '1001   部门管理
    '1002   人员管理
    '1003   结算方式设置
    '1004   收入项目设置
    '1005   费别等级设置
    '1009   收费细目管理
    '1010   期间划分
    '1011   基础参数设置
    '1012   部门上班安排
    '1013   疾病编码管理
    '1014   医疗小组管理
    '1015   价格等级管理
    '1021   入出分类
    '1031   专项记帐单管理
    '1292   影像参数设置
    
    
    Select Case lngModul
    Case 1000
        On Error Resume Next
        frmDataMove.Show , frmMain
    Case 1001
        frmDeptManage.Show , frmMain
    Case 1002
        frmPresManage.Show , frmMain
    Case 1003
        frmBalanceManage.Show , frmMain
    Case 1004
        frmEarnManage.Show , frmMain
    Case 1005
        frmChargeSortGrade.Show , frmMain
    Case 1009
        Call InitSystemPara
        frmChargeManage.Show , frmMain
    Case 1010
        frmCourse.Show , frmMain
    
    '参数设置模块
    Case 1011
        If glngSys \ 100 = 8 Then
            '药店运行参数
            frmDrugstorePara.Show vbModal, frmMain
        Else
            frmParPublic.Show vbModal, frmMain  '公共、病人
        End If
    Case 1007
        frmParClinic.Show vbModal, frmMain  '临床
    Case 1006
        frmParFee.Show vbModal, frmMain  '费用
    Case 1022
        frmParMedicine.Show vbModal, frmMain '药品
        
        
        
    Case 1012
        frmDeptTime.Show , frmMain
    Case 1013
        frmIllManage.Show , frmMain
    Case 1014
        frmMedicalTeam.Show , frmMain
    Case 1015 '价格等级管理
        frmPriceGradeManage.Show , frmMain
    Case 1031
        If gobjCustAcc Is Nothing Then
            Set gobjCustAcc = CreateObject("zl9CustAcc.clsCustAcc")
        End If
        If gobjCustAcc Is Nothing Then Exit Sub
        Call gobjCustAcc.CodeMan(glngSys, 1031, gcnOracle, frmMain, strDbUser)
    Case 1021
        frmMedInOutClass.Show , frmMain
    Case 1292
        frmParPacs.Show , frmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '返回：
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gblnMyStyle = zlDatabase.GetPara("使用个性化风格") = "1"
    gstr单位名称 = GetUnitName()
    gstrDbUser = strDbUser
    Call GetUserInfo
    
    glngSys = lngSys
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    glngModul = lngModul
    If Not CheckValid Then Exit Sub
    '-------------------------------------------------
    '1000   数据转移管理
    '1001   部门管理
    '1002   人员管理
    '1003   结算方式设置
    '1004   收入项目设置
    '1005   费别等级设置
    '1009   收费细目管理
    '1010   期间划分
    '1011   基础参数设置
    '1012   部门上班安排
    '1013   疾病编码管理
    '1014   医疗小组管理
    '1015   价格等级管理
    '1031   专项记帐单管理
    '1292   影像参数设置
    
    
    Select Case lngModul
    Case 1000
        On Error Resume Next
        zlCommFun.ShowChildWindow frmDataMove.hwnd, lngMain
    Case 1001
        zlCommFun.ShowChildWindow frmDeptManage.hwnd, lngMain
    Case 1002
        zlCommFun.ShowChildWindow frmPresManage.hwnd, lngMain
    Case 1003
        zlCommFun.ShowChildWindow frmBalanceManage.hwnd, lngMain
    Case 1004
        zlCommFun.ShowChildWindow frmEarnManage.hwnd, lngMain
    Case 1005
        zlCommFun.ShowChildWindow frmChargeSortGrade.hwnd, lngMain
    Case 1009
        Call InitSystemPara
        zlCommFun.ShowChildWindow frmChargeManage.hwnd, lngMain
    Case 1010
        zlCommFun.ShowChildWindow frmCourse.hwnd, lngMain

    '参数设置模块
    Case 1011
        If glngSys \ 100 = 8 Then
            '药店运行参数
            zlCommFun.ShowChildWindow frmDrugstorePara.hwnd, lngMain
        Else
            zlCommFun.ShowChildWindow frmParPublic.hwnd, lngMain '公共、病人
        End If
    Case 1007
        zlCommFun.ShowChildWindow frmParClinic.hwnd, lngMain '临床
    Case 1006
        zlCommFun.ShowChildWindow frmParFee.hwnd, lngMain '费用
    Case 1022
        zlCommFun.ShowChildWindow frmParMedicine.hwnd, lngMain '药品
        
        
    Case 1012
        zlCommFun.ShowChildWindow frmDeptTime.hwnd, lngMain
    Case 1013
        zlCommFun.ShowChildWindow frmIllManage.hwnd, lngMain
    Case 1014
        zlCommFun.ShowChildWindow frmMedicalTeam.hwnd, lngMain
    Case 1015 '价格等级管理
        zlCommFun.ShowChildWindow frmPriceGradeManage.hwnd, lngMain
    Case 1031
        If gobjCustAcc Is Nothing Then
            Set gobjCustAcc = CreateObject("zl9CustAcc.clsCustAcc")
        End If
        If gobjCustAcc Is Nothing Then Exit Sub
        Call gobjCustAcc.BHCodeMan(glngSys, 1031, gcnOracle, lngMain, strDbUser)
    Case 1021
        zlCommFun.ShowChildWindow frmMedInOutClass.hwnd, lngMain
    Case 1292
        zlCommFun.ShowChildWindow frmParPacs.hwnd, lngMain
    End Select
End Sub

Public Function OpenEditWholeSetItem(ByVal frmMain As Object, ByVal cnOracle As ADODB.Connection, _
     ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, ByVal strItems As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从单据中保存成成套项目数据
    '入参:frmMain:调用的主窗体
    '       cnOracle    ADODB.Connection    传入    传入连接对象
    '       lngSys  Long    传入    系统号
    '       lngModule   Long    传入    模块号
    '       strPrivs    String  传入    权限串
    '       strItems    String  传入    序号,父号,收费细目ID,数量,单价,执行科室|序号,父号,收费细目ID,数量,单价,执行科室|…
    '出参:
    '返回:True:保存成功,False:保存失败
    '编制:刘兴洪
    '日期:2010-09-01 14:06:53
    '说明: 目前支持以下模块功能调用:
    '        1.  住院记帐(划价)
    '        2.  医技工作站(记帐)
    '       该接口内部需要检查如下权限:
    '       无"增加成套项目"的权限时,则直接返回false
    '       无"全院成套方案"的权限时,则不能更改使用范围为全院
    '       无"本科成套方案"的权限时,则不能更改为指定科室,只能增加自己的成套项目.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng上次分类ID  As Long
    
    If InStr(1, strPrivs, ";增加成套项目;") = 0 Then Exit Function

    Set gcnOracle = cnOracle
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstr单位名称 = GetUnitName()
    Call GetUserInfo
    glngSys = lngSys
    glngModul = lngModule
    lng上次分类ID = Val(zlDatabase.GetPara("上次成套方案分类", glngSys, lngModule, ""))
    Call InitSystemPara
    OpenEditWholeSetItem = frmChargeWholeSetItemEdit.ShowCard(frmMain, EdI_增加, strPrivs, lngModule, lng上次分类ID, "", strItems)
End Function

Private Sub Class_Terminate()
    Set gobjPlugIn = Nothing
    
    If Not gobjRIS Is Nothing Then
        Set gobjRIS = Nothing
    End If
End Sub
