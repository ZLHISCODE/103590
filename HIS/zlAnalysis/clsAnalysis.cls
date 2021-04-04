VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAnalysis"
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
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstr单位名称 = GetUnitName()
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstr单位名称 = GetUnitName()
    gstrDbUser = strDbUser
    Call GetUserInfo
    If Not CheckValid Then Exit Sub

    '1401    科室资源状况表
    '1402    门诊应诊安排周报
    '1403    收费项目与价目
    '
    '1410    全院现金收入概况
    '1411    全院病人费用汇总
    '1412    分科病人费用情况
    '1413    全院未结费用汇总
    '1414    分科未结费用汇总
    '1415    全院预收款汇总表
    '1416    医院收入分析
    '1417    病人费别汇总表
    '1418    合约病人费用汇总
    '
    '1420    门诊人次汇总分析
    '1421    门诊流量统计分析
    '1422    住院病人流动分析
    '1423    在院病人分布
    '1424    在院危重病人
    '1425    在院特护病人
    '
    '1430    医院工作量统计表
    '1431    工作情况分类统计
    '1432    门诊诊次费用分析
    '1433    住院收治费用分析
    '1434    诊治收费项目分析
    '
    '1440    药品直接收支分析
    '1441    药品超储短缺分析
    '1442    药品效期报警分析
    '1443    药品滞用报警分析
    '1444    临床用药分析
    '1445    流向跟踪分析
    Select Case lngModul
        Case 1401
        Case 1402
        Case 1403
            frmPriceQuery.Show , frmMain
    End Select

End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDbUser As String)
    '------------------------------------------------
    '功能： 部件的初始化，并调用执行程序,将主程序的必要内容传递下来
    '参数：
    '   cnMain:主程序的数据库连接
    '返回：
    '------------------------------------------------
    Set gcnOracle = cnMain
    gstrAviPath = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstr单位名称 = GetUnitName()
    
    gstrPrivs = GetPrivFunc(lngSys, lngModul)
    gstr单位名称 = GetUnitName()
    gstrDbUser = strDbUser
    Call GetUserInfo
    If Not CheckValid Then Exit Sub

    '1401    科室资源状况表
    '1402    门诊应诊安排周报
    '1403    收费项目与价目
    '
    '1410    全院现金收入概况
    '1411    全院病人费用汇总
    '1412    分科病人费用情况
    '1413    全院未结费用汇总
    '1414    分科未结费用汇总
    '1415    全院预收款汇总表
    '1416    医院收入分析
    '1417    病人费别汇总表
    '1418    合约病人费用汇总
    '
    '1420    门诊人次汇总分析
    '1421    门诊流量统计分析
    '1422    住院病人流动分析
    '1423    在院病人分布
    '1424    在院危重病人
    '1425    在院特护病人
    '
    '1430    医院工作量统计表
    '1431    工作情况分类统计
    '1432    门诊诊次费用分析
    '1433    住院收治费用分析
    '1434    诊治收费项目分析
    '
    '1440    药品直接收支分析
    '1441    药品超储短缺分析
    '1442    药品效期报警分析
    '1443    药品滞用报警分析
    '1444    临床用药分析
    '1445    流向跟踪分析
    Select Case lngModul
        Case 1401
        Case 1402
        Case 1403
            Call zlCommFun.ShowChildWindow(frmPriceQuery.hWnd, lngMain)
    End Select

End Sub
