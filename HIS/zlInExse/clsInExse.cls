VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInExse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit '要求变量声明
Public Version As String

Public Sub CodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, frmMain As Object, ByVal strDBUser As String, Optional ByVal blnShowInTaskBar As Boolean)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   frmMain:主窗体
'   strDBUser:当前登录数据库用户名
    Set gfrmMain = frmMain
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    gblnShowInTaskBar = blnShowInTaskBar
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitLocPar(lngModul)
    Call CreateSquareCardObject(frmMain, lngModul)
    Call CreatePublicExpenseObject(lngModul)
    Select Case lngModul
        Case 1133 '住院记帐管理
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            Err.Clear
            On Error Resume Next
            If gfrmMain Is Nothing Then
                frmManageBilling.Show
            Else
                frmManageBilling.Show , frmMain
            End If
        Case 1134 '科室分散记帐
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear
            On Error Resume Next
            If gfrmMain Is Nothing Then
                frmDeptBilling.Show
            Else
                frmDeptBilling.Show , frmMain
            End If
        Case 1135 '医技科室记帐
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear
            On Error Resume Next
            If gfrmMain Is Nothing Then
                frmTechnoBilling.Show
            Else
                frmTechnoBilling.Show , frmMain
            End If
        Case 1136 '自动记帐计算
            frmCompute.Show 1, frmMain
        Case 1137 '病人结帐处理
            Call InitSysPar
                        
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            If gfrmMain Is Nothing Then
                frmManageBalance.Show
            Else
                frmManageBalance.Show , frmMain
            End If
        Case 1139 '病人费用查询
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
            
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            '不能取消:gobjKernel,原因是在费用查询中要用到zlCISKernel.clsDockInAdvices
            '因为;zlCISKernel.clsDockInAdvices未开放initCisKernel
            Set gobjKernel = CreateObject("zlCISKernel.clsCISKernel")
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs)
            gblnNurseStation = False
            If gfrmMain Is Nothing Then
                frmPatiFeeQuery.Show
            Else
                frmPatiFeeQuery.Show , frmMain
            End If
        Case 1141 '一日费用清单
            Call InitSysPar
            frmDailyList.Show , frmMain
        Case 1142 '执行登记管理
            Call InitSysPar
            On Error Resume Next
            frmManageExecute.Show , frmMain
        Case 1143 '费用审核管理
            Call InitSysPar
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            On Error Resume Next
            frmFeeVerifyManger.Show , frmMain
    End Select
    
End Sub

Public Sub BHCodeMan(lngSys As Long, ByVal lngModul As Long, cnMain As ADODB.Connection, lngMain As Long, ByVal strDBUser As String)
'功能： 根据主程序指定功能，调用执行相关程序
'参数：
'   lngModul:需要执行的功能序号
'   cnMain:主程序的数据库连接
'   lngMain:主窗体句柄
'   strDBUser:当前登录数据库用户名
    Set gfrmMain = Nothing
    If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrDBUser = strDBUser
    glngSys = lngSys
    glngModul = lngModul
    
    gstrPrivs = ";" & GetPrivFunc(lngSys, lngModul) & ";"
    gstrUnitName = GetUnitName
    gblnShowInTaskBar = True
    
    If Not GetUserInfo Then
        MsgBox "当前用户未设置对应的人员信息,请与系统管理员联系,先到用户授权管理中设置！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call InitLocPar(lngModul)
    Call CreateSquareCardObject(gfrmMain, lngModul)
    Call CreatePublicExpenseObject(lngModul)
    Select Case lngModul
        Case 1133 '住院记帐管理
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmManageBilling.hWnd, lngMain)
        Case 1134 '科室分散记帐
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmDeptBilling.hWnd, lngMain)
        Case 1135 '医技科室记帐
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
                        
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmTechnoBilling.hWnd, lngMain)
        Case 1136 '自动记帐计算
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmCompute.hWnd)
        Case 1137 '病人结帐处理
            Call InitSysPar
                        
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmManageBalance.hWnd, lngMain)
        Case 1139 '病人费用查询
            If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Sub
            Call InitSysPar
            
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Call CreatePlugIn(lngModul)
            '不能取消:gobjKernel,原因是在费用查询中要用到zlCISKernel.clsDockInAdvices
            '因为;zlCISKernel.clsDockInAdvices未开放initCisKernel
            Set gobjKernel = CreateObject("zlCISKernel.clsCISKernel")
            Call gobjKernel.InitCISKernel(gcnOracle, Nothing, glngSys, gstrPrivs)
                         
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmPatiFeeQuery.hWnd, lngMain)
        Case 1141 '一日费用清单
            Call InitSysPar
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmDailyList.hWnd, lngMain)
        Case 1142 '执行登记管理
            Call InitSysPar
            
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmManageExecute.hWnd, lngMain)
        Case 1143 '费用审核管理
            Call InitSysPar
            '医保接口
            Call gclsInsure.InitOracle(gcnOracle)
            Err.Clear: On Error Resume Next
            Call OS.ShowChildWindow(frmFeeVerifyManger.hWnd, lngMain)
    End Select
End Sub

Public Function CallByNurse(frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, lngSys As Long, _
    ByVal lng病区ID As Long, ByVal lng科室ID As Long, ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, ByVal strNO As String, ByVal bln划价 As Boolean) As Object
'功能：由护士工作站调用并自动过滤显示出指定医嘱条件的费用单据，目的是冲销这些单据
'返回：当从非模态窗口返回时,返回记帐管理窗体,用于跟踪关闭事件(非模态显示时的刷新问题)
    Static blnNoFirst As Boolean  '是否第一次调用,初始化全局变量
    
    If Not blnNoFirst Then
        blnNoFirst = True
    
        Set gfrmMain = frmMain
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrDBUser = strDBUser
        glngSys = lngSys
        glngModul = 1133
        
        gstrPrivs = ";" & GetPrivFunc(glngSys, glngModul) & ";"
        gstrUnitName = GetUnitName
        
        If CreatePublicDrug(lngSys, cnMain, strDBUser) = False Then Exit Function
        Call GetUserInfo
        Call InitSysPar
        Call InitLocPar(glngModul)
        Call CreateSquareCardObject(frmMain, glngModul)
        Call CreatePublicExpenseObject(glngModul)
        
        '医保接口
        Call gclsInsure.InitOracle(gcnOracle)
        Call CreatePlugIn(glngModul)
    End If
    
    Set CallByNurse = frmManageBilling.ShowMeByNurse(frmMain, lng病区ID, lng科室ID, lng发送号, lng医嘱ID, strNO, bln划价)
End Function

Public Function CallReCharge(ByVal frmParent As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, ByVal lngSys As Long, _
    ByVal bytFun As Integer, ByVal bytUseType As Byte, ByVal lngDeptID As Long, ByVal strPrivs As String, Optional ByVal lngPatientID As Long, _
    Optional ByVal strNO As String, Optional ByVal lngAdviceID As Long) As Boolean
'功能：调用销帐申请或审核功能(目前有医生站，护士站，医技工作站在调用)
'参数：bytFun=0-申请,1-审核
'      bytUseType=0-护士站调用,1-医技科室调用,2-医生站调用
'      lngDeptID=病区调用时传入当前操作的病人病区ID,医技科室调用时传入医技科室ID
'      strPrivs=调用程序所具有的权限,目前暂时未使用，因为销帐申请窗体是自己取并用1150模块的权限
    Static blnNoFirst As Boolean  '是否第一次调用,初始化全局变量
        
    If Not blnNoFirst Then
        blnNoFirst = True
        
        Set gfrmMain = frmParent
        If Not CheckValid Then Exit Function
        
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrDBUser = strDBUser
        glngSys = lngSys
        glngModul = 1133
                
        '不需要取1133住院记帐管理模块的权限
                
        Call GetUserInfo
        Call InitSysPar
        Call InitLocPar(glngModul)
        Call CreateSquareCardObject(frmParent, glngModul)
        Call gclsInsure.InitOracle(gcnOracle)
        Call CreatePlugIn(glngModul)
        Call CreatePublicExpenseObject(glngModul)
    End If
    
    gblnOK = False
    With frmReCharge
        .mbytFun = bytFun
        .mbytUseType = bytUseType
        .mlngDeptID = lngDeptID
        .mlngPatientID = lngPatientID
        .mstrPrivs = strPrivs
        .mstrInNO = strNO
        .mlngAdviceID = lngAdviceID
        If frmParent Is Nothing Then
            .Show 1
        Else
            .Show 1, frmParent
        End If
    End With
    CallReCharge = gblnOK
End Function

Public Function CallMzFeeTOZyFeeDel(ByVal frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, ByVal lngSys As Long, _
    ByVal lngModule As Long, ByVal strPrivs As String, ByVal int性质 As Integer, Optional ByRef lng病人ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊转住院费用销帐或退费
    '入参:int性质-1-门诊收费;2-门诊记帐
    '返回:只要有一笔以上转出成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-25 11:36:41
    '说明:目前该接口主要门诊收费管理的门诊转住院退费功能中调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static blnNoFirst As Boolean  '是否第一次调用,初始化全局变量
    If Not blnNoFirst Then
        blnNoFirst = True
        Set gfrmMain = frmMain
        Set gcnOracle = cnMain
        gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
        gstrDBUser = strDBUser: glngSys = lngSys: glngModul = lngModule
        Call GetUserInfo
        Call InitSysPar
        Call InitLocPar(glngModul)
        Call gclsInsure.InitOracle(gcnOracle)
        Call CreateSquareCardObject(frmMain, glngModul)
    End If
    If frmFeeRefundment.zlShowEdit(frmMain, int性质, glngModul, strPrivs, lng病人ID) = False Then Exit Function
    CallMzFeeTOZyFeeDel = True
End Function
Public Function FetchUnBalance(cnMain As ADODB.Connection, strPrivs As String, _
    int险类 As Integer, lng病人ID As Long, _
    Optional int业务类型 As Integer = 2) As ADODB.Recordset
    '功能：获取指定病人未结帐细目明细(按收费细目)
    '参数：lng病人ID=病人ID,
    '      int类务类型-1-门诊;2-住院(缺省)
    '返回：成功=记录集,失败=Nothing
    Call InitSysPar
    Set gcnOracle = cnMain
    If int业务类型 = 1 Then
        Set FetchUnBalance = GetVBalance(0, strPrivs, int险类, lng病人ID)
    Else
        Set FetchUnBalance = GetVBalance(1, strPrivs, int险类, lng病人ID)
    End If

End Function

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Dim blnChildren As Boolean
    
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    If Not gobjKernel Is Nothing Then
        blnChildren = gobjKernel.CloseWindows()
    Else
        blnChildren = True
    End If
    CloseWindows = blnChildren And (Forms.Count = 0)
End Function
Private Sub Class_Initialize()
    glngInstanceCount = glngInstanceCount + 1
    Version = App.Major & "." & App.Minor & "." & App.Revision
    Load frmPrint
End Sub
Private Sub Class_Terminate()
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
    
    Set grs医疗付款方式 = Nothing
    Unload frmPrint
    'Call CloseSquareCardObject: 共享时 , 不能关闭, 否则要出错, 在退出时, 程序会自动关闭, 59760
End Sub
