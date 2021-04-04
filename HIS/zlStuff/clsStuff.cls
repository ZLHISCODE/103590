VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStuff"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Version As String

Public Sub TransStuff(FrmMain As Object, cnMain As ADODB.Connection, lngSys As Long, ByVal strDBUser As String, ByVal lng病人id As Long, ByVal strNo As String, ByVal lng库房ID As Long, ByVal strStartDate As String, ByVal strEndDate As String)
'功能：直接调用卫材发放管理
    
    '公共变量
    Set gfrmMain = FrmMain
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    glngSys = lngSys
    glngModul = 1723
    gstrDBUser = strDBUser
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrUnitName = GetUnitName
    gstrPrivs = GetPrivFunc(lngSys, 1723)
                
    '用户信息
    Call GetUserInfo
    
    '初始化相关小数位数
    Call 初始小数位数
    
    If glngModul <> 1720 Then
        '刘兴宏:2007/06/28:初始化系统参数
        If InitSystemPara = False Then Exit Sub
    End If
    
    '调用卫材发放管理
    err.Clear
    On Error Resume Next
    If gfrmMain Is Nothing Then
        frm卫材发放管理_New.Show
    Else
        frm卫材发放管理_New.ShowList FrmMain, lng病人id, strNo, lng库房ID, strStartDate, strEndDate
    End If
End Sub

Public Sub CodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, FrmMain As Object, strDBUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngSys : 系统编号
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   frmMain:主窗体
    '返回：
    '------------------------------------------------
    Dim frmBillMgr As Form
    Dim frmTemp As Form
    Dim strTitle As String
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    gstrDBUser = strDBUser
    
    glngModul = lngModul
    glngSys = lngSys
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    Set gfrmMain = FrmMain
    
    '判读卫材系统是否具有条码管理的权限
    gblnCode = GetCodePrivs
    
    '读取卫材部分要使用的系统参数值
    Call GetSysParms
    
    '初始化相关小数位数
    Call 初始小数位数
    
    If lngModul <> 1720 Then
        '刘兴宏:2007/06/28:初始化系统参数
        If InitSystemPara = False Then Exit Sub
    End If
    
    '取药品价格等级
    Call GetPriceClass
    
    gbytSimpleCodeTrans = Val(zlDatabase.GetPara("简码匹配方式切换", , , 1))
    
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    
    gstrUnitName = GetUnitName
    
    '-------------------------------------------------
    On Error Resume Next
    
    Call setNOtExcetePrice '将到期未执行的价格先执行了
    
    Select Case lngModul
        Case 1710                         '系统参数设置'
'            frmSystemParaSet.Show , frmMain
            frmParStuff.Show , FrmMain
        Case 1711                         '卫材目录管理'
            frmStuffMgr.Show , FrmMain
        Case 1712                       '卫材外购入库管理
            strTitle = "卫材外购入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1713
            strTitle = "卫材自制入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1714                        '卫材其他入库管理
            strTitle = "卫材其他入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1715
            strTitle = "库存差价调整管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1716
            strTitle = "卫材移库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        
        Case 1717
            strTitle = "卫材领用管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1718
            strTitle = "卫材其他出库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1719
           strTitle = "卫材盘点管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmCheckMain
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1720                         '卫材差价计算'
           frmDiffPriceRecalCard.Show , FrmMain
        Case 1721                         '卫材库存查询'
            frmStuffQuery.Show , FrmMain
        Case 1722                         '卫材申领
            frmRequestStuffList.ShowList FrmMain
        Case 1723                         '卫材发放管理
             'frm卫材发放管理.Show , frmMain
             frm卫材发放管理_New.Show , FrmMain
        Case 1724               '卫材计划管理
            strTitle = "卫材计划管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmStuffPlanList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1725               '卫材申购管理
            strTitle = "卫材申购管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmStuffPlanList
            frmBillMgr.ShowList lngModul, strTitle, FrmMain
        Case 1726
            frmStuffPriceList.Show , FrmMain
        Case 1729
            frmClosingAccount.Show , FrmMain
    End Select
End Sub

Public Sub BHCodeMan(ByVal lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, strDBUser As String)
    '------------------------------------------------
    '功能： 根据主程序指定功能，调用执行相关程序
    '参数：
    '   lngSys : 系统编号
    '   lngModul:需要执行的功能序号
    '   cnMain:主程序的数据库连接
    '   lngMain:主窗体句柄
    '返回：
    '------------------------------------------------
    Dim frmBillMgr As Form
    Dim frmTemp As Form
    Dim strTitle As String
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    gstrDBUser = strDBUser
    
    glngModul = lngModul
    glngSys = lngSys
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    Set gfrmMain = Nothing
    
    '判读卫材系统是否具有条码管理的权限
    gblnCode = GetCodePrivs
    
    '初始化相关小数位数
    Call 初始小数位数
    
    If lngModul <> 1720 Then
        '刘兴宏:2007/06/28:初始化系统参数
        If InitSystemPara = False Then Exit Sub
    End If
    
    gbytSimpleCodeTrans = Val(zlDatabase.GetPara("简码匹配方式切换", , , 1))
    
    gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    
    gstrUnitName = GetUnitName
    
    '-------------------------------------------------
    On Error Resume Next
    
    Call setNOtExcetePrice '将到期未执行的价格先执行了
    
    Select Case lngModul
        Case 1710                         '系统参数设置'
            OS.ShowChildWindow frmSystemParaSet.hwnd, lngMain
        Case 1711                         '卫材目录管理'
            OS.ShowChildWindow frmStuffMgr.hwnd, lngMain
        Case 1712                       '卫材外购入库管理
            strTitle = "卫材外购入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1713
            strTitle = "卫材自制入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1714                        '卫材其他入库管理
            strTitle = "卫材其他入库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1715
            strTitle = "库存差价调整管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1716
            strTitle = "卫材移库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        
        Case 1717
            strTitle = "卫材领用管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1718
            strTitle = "卫材其他出库管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmMainList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1719
           strTitle = "卫材盘点管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmCheckMain
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1720                         '卫材差价计算'
            OS.ShowChildWindow frmDiffPriceRecalCard.hwnd, lngMain
        Case 1721                         '卫材库存查询'
            OS.ShowChildWindow frmStuffQuery.hwnd, lngMain
        Case 1722                         '卫材申领
            frmRequestStuffList.ShowList lngMain
        Case 1723                         '卫材发放管理
             OS.ShowChildWindow frm卫材发放管理_New.hwnd, lngMain
        Case 1724               '卫材计划管理
            strTitle = "卫材计划管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmStuffPlanList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1725               '卫材申购管理
            strTitle = "卫材申购管理"
            For Each frmTemp In Forms
                If frmTemp.Caption = strTitle Then
                    frmTemp.Show
                    frmTemp.ZOrder 0
                    Exit Sub
                End If
            Next
            Set frmBillMgr = New frmStuffPlanList
            frmBillMgr.ShowList lngModul, strTitle, lngMain
        Case 1726   '卫材调价管理
            OS.ShowChildWindow frmStuffPriceList.hwnd, lngMain
    End Select
End Sub

Private Sub setNOtExcetePrice()
    '如果已到执行日期而价格未执行，执行计算过程
        Dim rsTemp As ADODB.Recordset
        Dim lng收费细目ID As Long
        
        gstrSQL = "Select a.Id From 收费价目 A, 材料特性 B Where a.收费细目id = b.材料id And a.变动原因 = 0" & GetPriceClassString("A")
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "处理到时间未生效的价格")
        
        With rsTemp
            If Not rsTemp.EOF Then
                rsTemp.MoveFirst
                Do While Not rsTemp.EOF
                    If Not IsNull(rsTemp!Id) Then
                        lng收费细目ID = rsTemp!Id
                        gstrSQL = "zl_材料收发记录_Adjust(" & lng收费细目ID & ")"
                        
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "处理到时间未生效的价格" & "-产生材料价格调整记录")
                    End If
                    rsTemp.MoveNext
                Loop
            End If
        End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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



