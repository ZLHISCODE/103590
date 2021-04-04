VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFeeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'直接通过类模块激活的事件
Public Event Activate() '子窗体激活
Public Event RequestRefresh() '要求主窗体刷新
Public Event StatusTextUpdate(ByVal strText As String) '要求更新主窗体状态栏文字
Private WithEvents mfrmFeeQuery As frmFeeQuery
Attribute mfrmFeeQuery.VB_VarHelpID = -1
Private Sub Class_Initialize()
    Set mfrmFeeQuery = New frmFeeQuery
    glngInstanceCount = glngInstanceCount + 1
End Sub
Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmFeeQuery Is Nothing Then Unload mfrmFeeQuery
    Set mfrmFeeQuery = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

Private Sub mfrmFeeQuery_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmFeeQuery_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mfrmFeeQuery_StatusTextUpdate(ByVal strText As String)
    RaiseEvent StatusTextUpdate(strText)
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘兴洪
    '日期:2012-06-18 16:50:35
    '问题:50793
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmFeeQuery Is Nothing Then Exit Sub
    mfrmFeeQuery.SetFontSize (bytSize)
End Sub
Public Function zlGetForm() As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取
    
    '初始读取权限,以加快后续速度(不是必要的)
    Call GetInsidePrivs(Enum_Inside_Program.p费用查询, True)
    
    Set zlGetForm = mfrmFeeQuery
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByRef CommandBars As Object, ByVal int场合 As Integer, Optional ByVal blnChildToolBar As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
    '入参:int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
    '       CommandBars=仅用于查看时可以不传(传入Nothing)
    '       blnChildToolBar = True表示工具栏添加在自己的窗体内部
    '出参:
    '返回:
    '说明:
    '   定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)，如果bln内部工具栏为假，则不再主界面上创建工具栏（菜单仍然要创建），
    '   而需要在自己的界面上创建工具栏，因此对于自己界面上已经存在工具栏的程序，应避免关键字重复。
    '注意:
    '         添加工具栏时注意各个功能按钮的主键不要重复
    '         病人事务处理模块无菜单：conMenu_ManagePopup，因此子程序在处理时需要检查，无此对象时添加到自已的菜单中
    '         如果是添加内部工具栏，先删除活动工具栏后再添加
    '         未使用工具栏的模块需要添加初始化工具栏代码
    '         工具条的功能状态的变化，通过主程序调用zlUpdateCommandBars来统一更新
    '编制:刘兴洪
    '日期:2010-10-29 15:14:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmFeeQuery.zlDefCommandBars(frmParent, CommandBars, int场合, blnChildToolBar)
End Sub

Public Sub zlExecuteCommandBarsDirect(ByRef Control As Object, ByRef frmParent As Object, _
    ByRef strPrivs As String, ByVal bln批量 As Boolean, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal byt婴儿 As Byte, ByVal lng病区ID As Long, ByVal lng科室ID As Long, ByVal lng前提ID As Long, _
    ByVal lng界面科室ID As Long, ByVal int场合 As Integer, ByVal bln补费 As Boolean, _
    ParamArray arrPar() As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行相应模块的具体功能，如执行医嘱模块的新开医嘱功能
    '入参:int场合-暂无用(int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS))
    '       bln补费-对转科，转病区的病人进行补费
    '       lng科室ID-当补费为true时,则传入本次需要补费的科室ID
    '       lng病区ID-当补费为true时,则传入本次需要补费的病区ID
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-12 09:57:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objControl As CommandBarControl, bln出院 As Boolean, bln结清 As Boolean, rsTemp As ADODB.Recordset
    Select Case Control.ID
        Case conMenu_File_PrintBedCard: Call zlPrintBedCard(frmParent, lng病人ID, lng主页ID)    '打印床头卡
        Case conMenu_File_PrintSingleBill: Call zlExecPrintSingleBill(frmParent, lng病人ID, strPrivs, "")   '打印单张催款单
        Case conMenu_File_PrintDayDetail: Call zlPrintDayDetail(frmParent, 1, lng病人ID, lng病区ID, False, False) '打印一日清单
        Case conMenu_File_PrintPageSet: Call zlPrintAccountPage(frmParent)   '打印帐页设置
        Case conMenu_Edit_PreBalance    '预结算
            If bln批量 Then
                '批量预结
                Call zlPreBalanceAll(frmParent, lng病区ID)
            Else
                Call zlPreBalance(frmParent, lng病人ID, lng主页ID)
            End If
        Case conMenu_Edit_Billing   '记帐
        
            On Error GoTo errHandle

            gstrSQL = "Select A.出院日期,B.费用余额  From  病案主页 A,病人余额 B Where A.病人ID=[1] and A.主页ID=[2]  and a.病人id=b.病人ID(+) And b.类型(+)=1 And b.性质(+)=1 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息状态", lng病人ID, lng主页ID)
            If rsTemp.EOF Then Exit Sub
            bln出院 = IIf(Not IsNull(rsTemp!出院日期), True, False)
            bln结清 = IIf(Val(Nvl(rsTemp!费用余额)) = 0, True, False)
            '
            Call zlExecBilling(int场合, frmParent, lng病区ID, lng病人ID, bln出院, bln结清, "", lng主页ID, bln补费, lng科室ID)
        Case conMenu_Edit_Balance
            '结帐
            Call ExecBalance(frmParent, lng病人ID, CStr(lng主页ID))
        Case conMenu_Edit_Billing_Mulit    '批量记帐
            On Error GoTo errHandle
            gstrSQL = "Select A.出院日期,B.费用余额  From  病案主页 A,病人余额 B Where A.病人ID=[1] and A.主页ID=[2]  and a.病人id=b.病人ID(+) And b.类型(+)=1 And b.性质(+)=1 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息状态", lng病人ID, lng主页ID)
            If rsTemp.EOF Then Exit Sub
            bln出院 = IIf(Not IsNull(rsTemp!出院日期), True, False)
            bln结清 = IIf(Val(Nvl(rsTemp!费用余额)) = 0, True, False)
            '
            Call zlExecBilling_Mulit(int场合, frmParent, lng病区ID, lng病人ID, bln出院, bln结清, "", lng主页ID, bln补费, lng科室ID)
        Case conMenu_Edit_ReBilling '销帐
        Case conMenu_Edit_CardBackMoney '余额退款
            Call NurseDeposit(frmParent, lng病人ID, lng主页ID)
        Case conMenu_Edit_ReBillingApply, conMenu_Edit_ReBillingAudit   '销帐申请和审核
            Call zlWrite_Off_ApplyAndVerfy(frmParent, lng病区ID, lng病人ID, Control.ID = conMenu_Edit_ReBillingApply)
        Case conMenu_View_TurnToWardFeeQuery '转病区费用变动查询
            If CreatePublicExpenseBillOperation() Then
                Call gobjPublicExpenseBillOperation.zlTurnToWard_Fee_Query(frmParent, 3, lng病人ID, lng主页ID)
            End If
        Case conMenu_View_DateType * 10 + 1, conMenu_View_DateType * 10 + 2 '时间模式
        Case conMenu_View_DetailType * 10 To conMenu_View_DetailType * 10 + 9 '查询方式'
        Case conMenu_View_ReBalance '显示结帐作废
        Case conMenu_View_ZeroFee   '显示零费用
        Case conMenu_View_CheckFee  '显示体检费用
        Case conMenu_View_ToolBar_Button '工具栏
        Case conMenu_View_ToolBar_Text '按钮文字
        Case conMenu_View_ToolBar_Size '大图标
        Case conMenu_View_PatInfor  '查看病人卡片
        Case conMenu_View_Billing   '查看记帐单
        Case conMenu_View_Refresh
        Case conMenu_Tool_Option    '记帐选项
        Case Else
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ExecBalance(frmMain As Object, lng病人ID As Long, str主页ID As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行结帐操作
    '编制:刘兴洪
    '日期:2015-02-05 12:00:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivs As String
    
    strPrivs = ";" & GetInsidePrivs(Enum_Inside_Program.p病人结帐) & ";"
    If Val(zlDatabase.GetPara("结帐界面风格", glngSys, 1137, "1")) = 0 Then
        If frmPatiBalanceTraditional.ShowMe(frmMain, g_Ed_住院结帐, strPrivs, lng病人ID, str主页ID) = False Then Exit Sub
    Else
        If frmPatiBalanceSplit.ShowMe(frmMain, g_Ed_住院结帐, strPrivs, lng病人ID, str主页ID) = False Then Exit Sub
    End If
    
End Sub

Public Sub zlCheckPrivs(ByVal Control As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:由程序检查是否具备该权限
    '入参:Control-功能按钮的ID值
    '出参:
    '返回:存在权限,返回True,否则返回False
    '编制:刘兴洪
    '日期:2010-10-29 15:12:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Static strBillingPrivs As String
    Static strQueryPrivs As String
    Static strBalancePrivs As String
    Dim blnVisible As Boolean
    '权限只需判断一次,已经判断过的命令不用再判断
    If strBillingPrivs = "" Then
        strBillingPrivs = ";" & GetInsidePrivs(Enum_Inside_Program.p记帐操作) & ";"
        If strBillingPrivs = "" Then strBillingPrivs = "-"  '避免重复调用
    End If
    If strQueryPrivs = "" Then
        strQueryPrivs = GetInsidePrivs(Enum_Inside_Program.p费用查询)
        If strQueryPrivs = "" Then strQueryPrivs = "-"  '避免重复调用
    End If
    If strBalancePrivs = "" Then
        strBalancePrivs = GetInsidePrivs(Enum_Inside_Program.p病人结帐)
        If strBalancePrivs = "" Then strBalancePrivs = "-"  '避免重复调用
    End If
    blnVisible = True
    Select Case Control.ID
        Case conMenu_File_PrintPageSet
            blnVisible = InStr(strQueryPrivs, ";病人帐页") > 0
        Case conMenu_File_PrintMultiBill, conMenu_File_PrintSingleBill
            blnVisible = InStr(strQueryPrivs, ";催款单打印") > 0
        Case conMenu_Edit_PreBalanceAll
            blnVisible = InStr(strQueryPrivs, ";预结所有病人") > 0
        Case conMenu_Edit_Billing, conMenu_Edit_Billing_Mulit
            blnVisible = InStr(strBillingPrivs, ";住院记帐;") > 0
        Case conMenu_Edit_ReBilling
            '55380
            blnVisible = InStr(strBillingPrivs, ";药品销帐;") > 0 _
                Or InStr(strBillingPrivs, ";卫材销帐;") > 0 _
                Or InStr(strBillingPrivs, ";诊疗销帐;") > 0
        Case conMenu_Edit_Balance
            blnVisible = InStr(strBalancePrivs, ";允许现金结帐;") = 0
        Case conMenu_Edit_ReBillingApply
            blnVisible = (InStr(strBillingPrivs, ";药品销帐申请;") > 0 _
                Or InStr(strBillingPrivs, ";卫材销帐申请;") > 0 _
                Or InStr(strBillingPrivs, ";诊疗销帐申请;") > 0) And InStr(strBillingPrivs, "部分销帐") > 0
        Case conMenu_Edit_ReBillingAudit
            blnVisible = InStr(strBillingPrivs, "销帐审核") > 0
        Case conMenu_Edit_ReBillingButton
            '55380
            blnVisible = InStr(strBillingPrivs, "销帐审核") > 0 _
                Or ((InStr(strBillingPrivs, ";药品销帐申请;") > 0 _
                    Or InStr(strBillingPrivs, ";卫材销帐申请;") > 0 _
                    Or InStr(strBillingPrivs, ";诊疗销帐申请;") > 0) And InStr(strBillingPrivs, "部分销帐") > 0)
        Case conMenu_Edit_CardBackMoney
            blnVisible = InStr(";" & GetInsidePrivs(9000), ";在院病人余额退款;") > 0 Or InStr(";" & GetInsidePrivs(9000), ";出院病人余额退款;") > 0
    End Select
    Control.Visible = blnVisible
End Sub
Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mfrmFeeQuery.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    
    Call mfrmFeeQuery.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mfrmFeeQuery.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal str住院号 As String, ByVal lng病区ID As Long, _
    ByVal intInsure As Integer, ByVal blnDateMoved As Boolean, ByVal bln出院 As Boolean, _
    ByVal bln结清 As Boolean, Optional blnOnlyRefreshVar As Boolean, _
    Optional bln补费 As Boolean = False, Optional lng科室ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新病人费用数据
    '入参:
    '       bln补费-对转科，转病区的病人进行补费
    '       lng科室ID-当补费为true时,则传入本次需要补费的科室ID
    '       lng病区ID-当补费为true时,则传入本次需要补费的病区ID
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-10 14:24:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmFeeQuery.zlRefresh(lng病人ID, lng主页ID, str住院号, lng病区ID, intInsure, blnDateMoved, bln出院, bln结清, blnOnlyRefreshVar, bln补费, lng科室ID)
End Sub

Public Sub InitCallByNurse(frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, lngSys As Long)
'功能：由护士工作站调用时初始化数据
    
    Set gfrmMain = frmMain
    'If Not CheckValid Then Exit Sub
    
    Set gcnOracle = cnMain
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrDBUser = strDBUser
    glngSys = lngSys
    gblnNurseStation = True
    glngModul = 1139
    
    gstrPrivs = ";" & GetPrivFunc(glngSys, glngModul) & ";"
    gstrUnitName = GetUnitName
    
    Call GetUserInfo
    Call InitSysPar
    Call InitLocPar(glngModul)
    Call InitLocPar(Enum_Inside_Program.p住院记帐)
    Call CreateSquareCardObject(frmMain, glngModul)
    Call CreatePublicDrug(lngSys, cnMain, strDBUser)
    Call CreatePublicExpenseObject(glngModul)
    '医保接口
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreatePlugIn(glngModul)
End Sub
Public Function zlPatiPressMoney(ByVal frmMain As Object, _
    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, _
    ByVal lngMoudle As Long, strDBUser As String, _
    ByVal strPrivs As String, ByVal lng病区ID As Long, ByVal str病区名称 As String, _
    Optional lng病人ID As Long = 0, Optional bytPrintModule As Byte = 2) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进入病人催款管理界面
    '入参:frmMain-调用的窗口
    '       strDBUser-db用户
    '       bytPrintModule-2.打印;1-预览
    '出参:
    '返回:如果打印成功1个以上的病人,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-12-16 10:28:25
    '问题:36025
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glngModul = lngMoudle: glngSys = lngSys
    Set gcnOracle = cnOracle
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    gstrDBUser = strDBUser
    glngSys = lngSys
        
    Call GetUserInfo
    Call InitSysPar
    Call InitLocPar(glngModul)
    Call CreateSquareCardObject(frmMain, glngModul)

    zlPatiPressMoney = frmPatiPressMoney.zlPatiPressMoney(frmMain, lngMoudle, strPrivs, lng病区ID, str病区名称, lng病人ID, bytPrintModule)
End Function
Public Function zlPatiBalance(ByVal frmMain As Object, _
    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, strDBUser As String, _
    ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strDelTime As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人结帐
    '入参:frmMain-调用的窗口
    '       strDBUser-db用户
    '       strDelTime-结帐作废的时间或转出时间(主要是处理冲预交情况)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-03-29 13:51:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    glngModul = 1137: glngSys = lngSys: gstrDBUser = strDBUser
    Set gcnOracle = cnOracle
    gstrPrivs = ";" & GetPrivFunc(lngSys, glngModul) & ";"
    If (InStr(gstrPrivs, ";门诊费用结帐;") = 0) Or _
        (InStr(gstrPrivs, ";普通病人结算;") = 0 And InStr(gstrPrivs, ";保险结算;") = 0) Then
        Exit Function    '没有权限,不能执行
    End If
    If lng主页ID = 0 And InStr(gstrPrivs, ";门诊费用结帐;") = 0 Then
        MsgBox "你不具备门诊费用结帐权限,请与系统管理员联系!", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    Call GetUserInfo: Call InitSysPar
    Call InitLocPar(glngModul)
    '医保接口
    Call gclsInsure.InitOracle(gcnOracle)
    Call CreatePlugIn(glngModul)
    Call CreateSquareCardObject(frmMain, glngModul)
    
    'IIf(lng主页ID = 0, 0, 1) '0-门诊;1-住院
    zlPatiBalance = frmPatiBalanceSplit.ShowMe(frmMain, IIf(lng主页ID = 0, g_Ed_门诊结帐, g_Ed_住院结帐), _
        gstrPrivs, lng病人ID, CStr(lng主页ID), , , True)
End Function

Public Function zlPatiBilling(ByVal frmMain As Object, _
    ByVal lng病人ID As Long, ByVal lng病区ID As Long, lng主页ID As Long, _
    ByVal lng科室ID As Long, ByVal bln补费 As Boolean, ByVal lng医嘱ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人记帐(关联医嘱),目前护士站调用
    '入参:frmMain-调用的窗口
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘尔旋
    '日期:2015-12-17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln出院 As Boolean, bln结清 As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim bln门诊留观病人 As Boolean
    
    On Error GoTo errHandle
    gstrSQL = _
        " Select A.出院日期,B.费用余额,a.病人性质" & _
        " From  病案主页 A,病人余额 B" & _
        " Where A.病人ID=[1] and A.主页ID=[2]  and a.病人id=b.病人ID(+) And b.类型(+)=1 And b.性质(+)=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人信息状态", lng病人ID, lng主页ID)
    If rsTemp.EOF Then Exit Function
    bln出院 = Not IsNull(rsTemp!出院日期)
    bln结清 = Val(Nvl(rsTemp!费用余额)) = 0
    bln门诊留观病人 = Val(Nvl(rsTemp!病人性质)) = 1
    
    zlPatiBilling = zlExecBilling(1, frmMain, lng病区ID, lng病人ID, bln出院, bln结清, "", _
        lng主页ID, bln补费, lng科室ID, lng医嘱ID, bln门诊留观病人)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetFormOperation() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取窗体操作选择，窗体卸载前调用
    '返回:
    '     上次窗体操作条件字符串
    '---------------------------------------------------------------------------------------------------------------------------------------------
    GetFormOperation = mfrmFeeQuery.zlGetFormOperation
End Function

Public Function RestoreFormOperation(ByVal strValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:恢复窗体操作选择，窗体刷新前调用
    '入参:
    '     strValue-上次窗体操作条件字符串
    '返回:
    '     True-窗体恢复成功;False-窗体恢复失败
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strValue = "" Then Exit Function
    If mfrmFeeQuery.zlRestoreFormOperation(strValue) = False Then Exit Function
End Function

