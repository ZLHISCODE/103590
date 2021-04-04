VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event Activate() '子窗体激活
Public Event RequestRefresh() '要求主窗体刷新
Public Event StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String) '要求更新主窗体状态栏文字
'bytType:1-附费执行,2-附费取消

'在附费列表中点击右键菜单时,触发该事件
Public Event zlPopupMenu(lng医嘱ID As Long, lng发送号 As Long, strNO As String, int记录性质 As Integer, X As Single, Y As Single)
Private WithEvents mfrmExpense As frmDockExpense
Attribute mfrmExpense.VB_VarHelpID = -1
Private mlngfrmMainHwnd As Long '父窗口对象
Private Type ty_adviceProperty '医嘱信息
    lng医嘱ID As Long
    lng发送号 As Long
    bln独立执行 As Boolean
    lng病人ID  As Long
    lng主页Id   As Long
    str挂号单   As String
    lng病人科室ID   As Long
    lng病人病区ID   As Long
    lng开嘱科室ID   As Long
    lng执行科室ID As Long
    int记录性质   As Integer
    int审核标志   As Integer
    int结算模式   As Integer
    int病人来源 As Integer
    int执行状态 As Integer

    lng相关ID  As Long
    str诊疗类别  As String
    strNO As String
    dat发送时间  As Date
    str费别   As String
    lng计价性质  As Long
    bln门诊记帐 As Boolean
    str计费状态 As String
    strFeeTab As String
    blnDataMoveed As Boolean '数据是否转出到历史数据空间
End Type

Public Function zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 17:55:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    glngSys = lngSys: Set gcnOracle = cnOracle: gstrDBUser = strDbUser
    If zlGetComLib = False Then Exit Function
    
    Call GetUserInfo
    Call InitVar
    zlInitCommon = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘兴洪
    '日期:2014-04-09 11:46:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmExpense Is Nothing Then Exit Sub
    mfrmExpense.SetFontSize (bytSize)
End Sub
Public Function SetParentHwnd(ByVal lngfrmMainHwnd As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置父窗口句配
    '入参:lngfrmMainHwnd-父窗口句柄
    '返回:设置成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-19 14:52:51
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If lngfrmMainHwnd = 0 Then Exit Function
    mlngfrmMainHwnd = lngfrmMainHwnd
    glngMainHwnd = lngfrmMainHwnd
End Function


Private Sub Class_Initialize()
    Set mfrmExpense = New frmDockExpense
    glngInstanceCount = glngInstanceCount + 1
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    If Not mfrmExpense Is Nothing Then Unload mfrmExpense
    Set mfrmExpense = Nothing
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call zlReleaseResources
End Sub

Private Sub mfrmExpense_Activate()
    RaiseEvent Activate
End Sub

Private Sub mfrmExpense_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mfrmExpense_StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String)
    RaiseEvent StatusTextUpdate(bytType, Text)
End Sub

Private Sub mfrmExpense_zlPopupMenu(lng医嘱ID As Long, lng发送号 As Long, strNO As String, int记录性质 As Integer, X As Single, Y As Single)
    '在附费列表中点击右键菜单时,触发该事件
    RaiseEvent zlPopupMenu(lng医嘱ID, lng发送号, strNO, int记录性质, X, Y)
End Sub


Public Function zlGetForm() As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取子业务窗口,由主程序初始化时只调用一次获取
    '返回:获取附费窗体对象,用于主界面绑定
    '编制:刘兴洪
    '日期:2014-04-09 12:00:14
    '说明:
    '     1.调用该函数,将会初始化医保部件:可分散到具体的GetForm中去初始化
    '     2.初始化医嘱附费管理权限
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call gclsInsure.InitOracle(gcnOracle)
    '初始读取权限,以加快后续速度(不是必要的)
    Call GetInsidePrivs(p医嘱附费管理, True)
    Set zlGetForm = mfrmExpense
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, Optional ByRef objSquareCard As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
    '入参:CommandBars=仅用于查看时可以不传(传入Nothing)
    '     objSquareCard=通过医技工作站传入的一卡通结算部件
    '编制:刘兴洪
    '日期:2014-04-09 12:01:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmExpense.zlDefCommandBars(frmParent, CommandBars, objSquareCard)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据附费内容,设置窗体的菜单和工具条的各项功能的Visible属性或Eanbeld属性
    '入参:Control-菜单控件
    '编制:刘兴洪
    '日期:2014-04-09 13:17:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmExpense.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行附费的相关功能
    '入参:Control-菜单控件
    '编制:刘兴洪
    '日期:2014-04-09 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmExpense.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据附费内容,将相关功能菜单放入单出菜单下
    '入参:CommandBar-CommandBar对象
    '编制:刘兴洪
    '日期:2014-04-09 13:19:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmExpense.zlPopupCommandBars(CommandBar)
End Sub
Public Function zlRefresh(ByVal lng科室id As Long, ByVal strAdviceIdAndPayNums As String, _
    Optional ByVal blnMoved As Boolean = False, Optional ByVal strNos As String, _
    Optional ByVal byt记录性质 As Byte, Optional ByVal byt病人来源 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据重新刷新
    '入参:lng科室id-科室ID
    '     strAdviceIdAndPayNums-医嘱ID和发送号和独立执行标志(医嘱ID1:发送号1:独立执行,医嘱ID2:发送号2:独立执行,...)
    '     strNos:单据号(多个传入时,用逗号分离)
    '     byt记录性质:医嘱ID传空时,才传入,单据性质(1-收费单;2-记帐单)
    '     byt病人来源-1-门诊;2-住院
    '     blnMoved -该病人的数据是否已转出
    '     bln单独执行-用于检验项目，一并采集的一组项目，是否针对其中的某一个单独执行
    '出参:
    '编制:刘兴洪
    '日期:2014-04-10 11:02:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
   zlRefresh = mfrmExpense.zlRefresh(lng科室id, strAdviceIdAndPayNums, blnMoved, strNos, byt记录性质, byt病人来源)
End Function
Public Function BuildMainExpense() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:生成主费用
    '返回:生成成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    BuildMainExpense = mfrmExpense.zlBuildMainExpense()
End Function

Public Function AddChargeExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补收费费用
    '入参:frmMain-调用主窗体
    '     lngModule-模块号
    '     byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '出参:strOutNos-成功保存的单据号
    '返回:补收费费用,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID; 开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    AddChargeExpense = mfrmExpense.zlAddChargeExpense(frmMain, lngModule, byt应用场合, lng病人ID, lng开单科室id, lng病人科室ID, strOutNos, objSaveData)
End Function
Public Function AddBillingExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, ByVal int病人来源 As Integer, _
    Optional ByVal lng病人ID As Long, Optional lng主页Id As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补记帐费用
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '    int病人来源:1-门诊病人,2-住院病人
    '出参:strOutNos-成功保存的单据号
    '返回:补记帐费用,补费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID;开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    AddBillingExpense = mfrmExpense.zlAddBillingExpense(frmMain, lngModule, byt应用场合, int病人来源, lng病人ID, lng主页Id, lng开单科室id, lng病人科室ID, strOutNos, objSaveData)
End Function

Public Function AddZeroExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, ByVal int病人来源 As Integer, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng主页Id As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补零费用
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '    int病人来源:1-门诊病人,2-住院病人
    '    lng主页ID-主页ID可以不传(但住院病人一定要传入)
    '出参:strOutNos-返回零耗费单号,多个用逗号分离
    '返回:补记帐费用,补费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID;开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    AddZeroExpense = mfrmExpense.zlAddZeroExpense(frmMain, lngModule, byt应用场合, int病人来源, lng病人ID, lng主页Id, lng开单科室id, lng病人科室ID, strOutNos, objSaveData)
End Function

Public Function AddStuffChargeExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, ByVal int病人来源 As Integer, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补备货卫材收费费用
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '    int病人来源:1-门诊病人,2-住院病人
    '出参:strNo-返回备货卫材单号,多个用逗号分离
    '返回:补记帐费用,补费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID;开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    AddStuffChargeExpense = mfrmExpense.zlAddStuffChargeExpense(frmMain, lngModule, byt应用场合, int病人来源, lng病人ID, lng开单科室id, lng病人科室ID, strOutNos, objSaveData)
End Function
Public Function AddStuffBillingExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal byt应用场合 As Byte, ByVal int病人来源 As Integer, _
    Optional ByVal lng病人ID As Long, _
    Optional ByVal lng主页Id As Long, _
    Optional ByVal lng开单科室id As Long, _
    Optional ByVal lng病人科室ID As Long, _
    Optional ByRef strOutNos As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:补备货卫材记帐费用
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    byt应用场合:0-医嘱附费;1-体检补费(可选参数)
    '    int病人来源:1-门诊病人,2-住院病人
    '    lng主页ID-主页ID可以不传(但住院病人一定要传入)
    '出参:strNo-返回备货卫材单号,多个用逗号分离
    '返回:补记帐费用,补费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.医嘱附费执行具体的功能(参见:zlCisKernel.dockExpense)
    '       不需要传入病人ID;开单科室及病人科室ID
    '    2.体检补费时,需要传入lng病人ID,开单科室ID,病人科室ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    AddStuffBillingExpense = mfrmExpense.zlAddStuffBillingExpense(frmMain, lngModule, byt应用场合, int病人来源, lng病人ID, lng主页Id, lng开单科室id, lng病人科室ID, strOutNos, objSaveData)
End Function
Public Function IsFunValied(ByVal bytType As Byte, ByVal bytPrivsCheck As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判检查某功能是否有效
    '入参: bytType- 1-修改附费;2-删除附费;3-附费转移;4-附费执行;5-附费取消执行;6-销帐申请;7-销帐审核
    '      bytPrivsCheck -检查权限:0-不检查权限;1-检查数据和权限;2-仅检查权限
    '出参:
    '返回:功能有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 17:00:52
    '说明:
    '   1.根据附费列表中的内容,检查某项功能是否有效
    '   2.根据权限检查某项功能是否有效
    '---------------------------------------------------------------------------------------------------------------------------------------------
     IsFunValied = mfrmExpense.zlIsFunValied(bytType, bytPrivsCheck)
End Function

Public Function ModifyExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, Optional int病人来源 As Integer, _
    Optional ByVal int记录性质 As Integer, _
    Optional ByVal strNO As String, Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改附费
    '入参:frmMain -调用主窗体
    '    int病人来源-1-门诊;2-住院
    '    lngModule -模块号
    '    int记录性质
    '    strNO
    '返回:修改附费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.执行具体的功能(参见:zlCisKernel.dockExpense),不需要传入记录性质和NO
    '    2.体检调用时,需要传入,病人来源,记录性质和NO
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ModifyExpense = mfrmExpense.zlModifyExpense(frmMain, lngModule, int病人来源, int记录性质, strNO, objSaveData)
End Function
Public Function DelExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, Optional int病人来源 As Integer, _
    Optional ByVal int记录性质 As Integer, _
    Optional ByVal strNO As String, _
    Optional ByRef objSaveData As Object) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除附费
    '入参:frmMain -调用主窗体
    '    int病人来源-1-门诊;2-住院
    '    lngModule -模块号
    '    int记录性质
    '    strNO
    '返回:修改附费成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '说明:
    '    1.执行具体的功能(参见:zlCisKernel.dockExpense),不需要传入记录性质和NO
    '    2.体检调用时,需要传入,病人来源,记录性质和NO
    '---------------------------------------------------------------------------------------------------------------------------------------------
    DelExpense = mfrmExpense.zlDelExpense(frmMain, lngModule, int病人来源, int记录性质, strNO, objSaveData)
End Function
Public Function MoveExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:附费转移
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:附费移动成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    MoveExpense = mfrmExpense.zlMoveExpense(frmMain, lngModule)
End Function

Public Function ExcuteExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long, Optional ByVal bln取消执行 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:附费执行
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '    bln取消执行-是否取消执行
    '返回:附费执行成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ExcuteExpense = mfrmExpense.zlExcuteExpense(frmMain, lngModule, bln取消执行)
End Function

Public Function ApplyExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐申请
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:销帐申请成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ApplyExpense = mfrmExpense.zlApplyExpense(frmMain, lngModule)
End Function
 
Public Function AuditExpense(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐申请审核
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:销帐申请审核成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    AuditExpense = mfrmExpense.zlAuditExpense(frmMain, lngModule)
End Function
Public Function ParaOptionSet(ByVal frmMain As Object, _
    ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:销帐申请审核
    '入参:frmMain -调用主窗体
    '    lngModule -模块号
    '返回:销帐申请审核成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-04-10 13:37:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ParaOptionSet = mfrmExpense.zlParaOptionSet(frmMain, lngModule)
End Function



Public Function zlBuldingMainFeeFromAdviceID(ByVal lng医嘱ID As Long, _
    ByVal lng发送号 As Long, ByVal bln独立执行 As Boolean, Optional ByVal lng科室id As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据医嘱ID生成主费用
    '入参:lng医嘱ID-医嘱ID
    '     lng发送号-发送号
    '     lng科室ID -科室ID(不传入，以病人医嘱发送中的执行科室为准)
    '返回:主费用生成成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-07 15:12:55
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyAdviceProperty As ty_adviceProperty, blnMoveed As Boolean
    Dim rsPrice As ADODB.Recordset
    On Error GoTo errHandle
    
    If lng医嘱ID = 0 Then Exit Function
    
    '1.获取医嘱相关信息
    If GetAdviceProperty(lng科室id, lng医嘱ID, lng发送号, bln独立执行, blnMoveed, tyAdviceProperty) = False Then zlBuldingMainFeeFromAdviceID = False: Exit Function
    '2.获取医嘱相关的计价信息
    If GetAdvicePrice(tyAdviceProperty, rsPrice) = False Then Exit Function
    '3.检查合法性
    If CheckBuldingFeeValied(tyAdviceProperty, rsPrice) = False Then Exit Function
    '4.保存数据
    If Not SaveBuildingMainFee(tyAdviceProperty, rsPrice) Then Exit Function
    zlBuldingMainFeeFromAdviceID = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function
Private Function SaveBuildingMainFee(ByRef tyAdviceProperty As ty_adviceProperty, ByRef rsPrice As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存生成的主费用
    '入参:tyAdviceProperty-医嘱相关属性
    '     rsPrice-主费用相关价目
    '返回:保存成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-07 16:14:09
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim cllPro As Collection
    Dim int来源 As Integer, lng医嘱ID As Long
    Dim int价格父号 As Integer, lng项目id As Long, lng执行部门ID As Long
    Dim lng病人病区ID As Long, lng病人科室ID As Long, lng类别ID As Long
    Dim strSql As String, strDate As String, i As Long
    Dim int保险项目否 As Integer, lng保险大类ID As Long, str保险编码 As String, cur统筹金额 As Currency, str费用类型 As String
    Dim lng开嘱科室ID As Long, str开嘱医生 As String, int序号 As Integer, strMsg As String
    Dim int父序号 As Integer
    Dim blnTrans  As Boolean
 
    On Error GoTo errH:
    
    int来源 = IIf(tyAdviceProperty.int病人来源 = 2, 2, 1)
    '获取病人的信息
    strSql = "Select Nvl(Nvl(B.姓名, C.姓名), A.姓名) 姓名, Nvl(Nvl(B.性别, C.性别), A.性别) 性别, Nvl(Nvl(B.年龄, C.年龄), A.年龄) 年龄," & vbNewLine & _
            "       Nvl(B.费别, A.费别) As 费别, A.门诊号, B.住院号, Nvl(A.当前床号, B.出院病床) As 床号, Nvl(A.当前病区id, B.当前病区id) As 病人病区id," & vbNewLine & _
            "       Nvl(A.当前科室id, B.出院科室id) As 病人科室id, Nvl(B.险类, A.险类) As 险类, D.编码 As 付款码" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人挂号记录 C, 医疗付款方式 D" & vbNewLine & _
            "Where A.病人id = B.病人id(+) And B.主页id(+) = [2] And A.病人id = C.病人id(+) And A.门诊号 = C.门诊号(+) And A.医疗付款方式 = D.名称(+) And A.病人id=[1]"

     Set rsPati = gobjDatabase.OpenSQLRecord(strSql, "获取病人信息", tyAdviceProperty.lng病人ID, tyAdviceProperty.lng主页Id)
    
    '可能对照费用为药品费用
    If tyAdviceProperty.int记录性质 = 1 Then
        lng类别ID = ExistIOClass(8) '门诊划价单
    Else
        lng类别ID = ExistIOClass(9) '门诊/住院记帐单
    End If
    
    '可能发送时已自动生成了部份主费用,现在是手工生成剩余部份。
    '1.因为单据号相同,所以要保持序号连续
    '2.如果是生成收费划价单，要保证一张单据中登记时间相同(不然收费无法处理)
    '3.第2点的情况，如果部份主费用已经收费，则不允许再生成主费用
    int序号 = GetBillMax序号(tyAdviceProperty.strNO, tyAdviceProperty.int记录性质, strDate, IIf(tyAdviceProperty.int病人来源 = 2, 2, 1))
    If tyAdviceProperty.int记录性质 = 2 Or strDate = "" Then
        strDate = "To_Date('" & Format(gobjDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strDate = "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    Set cllPro = New Collection
     
    With rsPrice
        .MoveFirst
        For i = 1 To .RecordCount
            '获取对应的医嘱信息
            If lng医嘱ID <> !医嘱ID Then
                strSql = "Select 医嘱期效,病人科室ID,开嘱科室ID,开嘱医生,婴儿,执行频次,计价特性 From 病人医嘱记录 Where ID=[1]"
                Set rsAdvice = gobjDatabase.OpenSQLRecord(strSql, "获取医嘱信息", Val(!医嘱ID))
                
                '将当前这条计费医嘱标记为已计费
                strSql = "ZL_病人医嘱发送_计费(" & !医嘱ID & "," & tyAdviceProperty.lng发送号 & ")"
                zlAddArray cllPro, strSql
                int父序号 = 0
            End If
            lng医嘱ID = !医嘱ID
            
            '病人病区科室
            lng病人病区ID = Nvl(rsPati!病人病区ID, 0)
            lng病人科室ID = Nvl(rsPati!病人科室id, 0)
            If lng病人科室ID = 0 Then
                lng病人病区ID = Nvl(rsAdvice!病人科室id, 0)
                lng病人科室ID = Nvl(rsAdvice!病人科室id, 0)
            End If
            If lng病人科室ID = 0 Then
                lng病人病区ID = UserInfo.部门ID
                lng病人科室ID = UserInfo.部门ID
            End If
            
            '开单科室及开单人
            lng开嘱科室ID = rsAdvice!开嘱科室id
            str开嘱医生 = rsAdvice!开嘱医生
            
            '每个收费项目的处理
            If lng项目id <> !收费细目ID Then
                int价格父号 = int序号 '获取价格父号
                If !从项 = 0 Then int父序号 = int序号  '读取Price时是分医嘱将主项排在前面的
                lng执行部门ID = Get收费执行科室ID(tyAdviceProperty.lng病人ID, tyAdviceProperty.lng主页Id, !类别, !收费细目ID, !执行科室, Nvl(rsAdvice!病人科室id, 0), Nvl(rsAdvice!开嘱科室id, 0), int来源)
                            
                '获取保险项目信息
                If int来源 = 2 And Not IsNull(rsPati!险类) Then
                    strMsg = gclsInsure.GetItemInsure(tyAdviceProperty.lng病人ID, !收费细目ID, !实收, False, rsPati!险类, "||" & !数量)
                    If strMsg <> "" Then
                        int保险项目否 = Val(Split(strMsg, ";")(0))
                        lng保险大类ID = Val(Split(strMsg, ";")(1))
                        cur统筹金额 = Format(Val(Split(strMsg, ";")(2)), gSysPara.Money_Decimal.strFormt_VB)
                        str保险编码 = CStr(Split(strMsg, ";")(3))
                        If UBound(Split(strMsg, ";")) >= 5 Then
                            If Split(strMsg, ";")(5) <> "" Then
                                str费用类型 = Split(strMsg, ";")(5)
                            End If
                        End If
                    End If
                End If
            End If
            lng项目id = !收费细目ID
            If int来源 = 1 Then
                If tyAdviceProperty.int记录性质 = 1 Then
                    '生成门诊划价单据
                    strSql = _
                        "zl_门诊划价记录_Insert('" & tyAdviceProperty.strNO & "'," & int序号 & "," & tyAdviceProperty.lng病人ID & ",NULL," & _
                        IIf(IsNull(rsPati!门诊号), "NULL", "'" & rsPati!门诊号 & "'") & ",'" & Nvl(rsPati!付款码) & "','" & Nvl(rsPati!姓名) & "'," & _
                        "'" & Nvl(rsPati!性别) & "','" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!费别) & "',NULL," & _
                        lng病人科室ID & "," & lng开嘱科室ID & ",'" & str开嘱医生 & "'," & _
                        IIf(Val(Nvl(!从项)) = 1, ZVal(int父序号), "NULL") & "," & lng项目id & ",'" & !类别 & "','" & !计算单位 & "',NULL,1," & !数量 & "," & _
                        !附加手术 & "," & ZVal(lng执行部门ID) & "," & IIf(int价格父号 = int序号, "NULL", int价格父号) & "," & _
                        !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & !应收 & "," & !实收 & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.姓名 & "',NULL," & _
                        !医嘱ID & ",'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                        Nvl(rsAdvice!计价特性, 0) & ",1,'" & str保险编码 & "','" & str费用类型 & "'," & int保险项目否 & "," & ZVal(lng保险大类ID) & "," & _
                        "NULL,0,NULL,NULL," & ZVal(lng病人病区ID) & ")"
                    zlAddArray cllPro, strSql
                        
                Else
                    '生成门诊记帐单据
                    strSql = _
                        "zl_门诊记帐记录_Insert('" & tyAdviceProperty.strNO & "'," & int序号 & "," & tyAdviceProperty.lng病人ID & "," & _
                        IIf(IsNull(rsPati!门诊号), "NULL", "'" & rsPati!门诊号 & "'") & ",'" & Nvl(rsPati!姓名) & "','" & Nvl(rsPati!性别) & "'," & _
                        "'" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!费别) & "',NULL," & ZVal(rsAdvice!婴儿) & "," & _
                        lng病人科室ID & "," & lng开嘱科室ID & "," & _
                        "'" & str开嘱医生 & "'," & IIf(!从项 = 1, ZVal(int父序号), "NULL") & "," & lng项目id & ",'" & !类别 & "'," & _
                        "'" & !计算单位 & "',1," & !数量 & "," & !附加手术 & "," & ZVal(lng执行部门ID) & "," & _
                        IIf(int价格父号 = int序号, "NULL", int价格父号) & "," & !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & _
                        !应收 & "," & !实收 & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.编号 & "'," & _
                        "'" & UserInfo.姓名 & "',NULL,NULL," & !医嘱ID & "," & _
                        "'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                        Nvl(rsAdvice!计价特性, 0) & ",1,NULL,0,NULL," & ZVal(tyAdviceProperty.lng主页Id) & "," & ZVal(lng病人病区ID) & ")"
                    zlAddArray cllPro, strSql
                End If
            Else
                '生成住院记帐单据
                strSql = _
                    "zl_住院记帐记录_Insert('" & tyAdviceProperty.strNO & "'," & int序号 & "," & tyAdviceProperty.lng病人ID & "," & ZVal(tyAdviceProperty.lng主页Id) & "," & _
                    IIf(IsNull(rsPati!住院号), "NULL", "'" & rsPati!住院号 & "'") & ",'" & Nvl(rsPati!姓名) & "','" & Nvl(rsPati!性别) & "'," & _
                    "'" & Nvl(rsPati!年龄) & "','" & Nvl(rsPati!床号) & "','" & Nvl(rsPati!费别) & "'," & _
                    lng病人病区ID & "," & lng病人科室ID & ",NULL," & ZVal(rsAdvice!婴儿) & "," & _
                    lng开嘱科室ID & ",'" & str开嘱医生 & "'," & IIf(!从项 = 1, ZVal(int父序号), "NULL") & "," & lng项目id & ",'" & !类别 & "'," & _
                    "'" & !计算单位 & "'," & int保险项目否 & "," & ZVal(lng保险大类ID) & ",'" & str保险编码 & "'," & _
                    "1," & !数量 & "," & !附加手术 & "," & ZVal(lng执行部门ID) & "," & _
                    IIf(int价格父号 = int序号, "NULL", int价格父号) & "," & !收入项目ID & ",'" & Nvl(!收据费目) & "'," & !单价 & "," & _
                    !应收 & "," & !实收 & "," & cur统筹金额 & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL," & ZVal(lng类别ID) & ",NULL,NULL,NULL," & _
                    !医嘱ID & ",'" & Nvl(rsAdvice!执行频次) & "',NULL,NULL," & Nvl(rsAdvice!医嘱期效, 0) & "," & _
                    Nvl(rsAdvice!计价特性, 0) & ",NULL,'" & str费用类型 & "')"
                zlAddArray cllPro, strSql
            End If
            int序号 = int序号 + 1
            
            .MoveNext
        Next
    End With
    
    On Error GoTo errH
    blnTrans = True
   zlExecuteProcedureArrAy cllPro, "主费用保存", , False
     
    '在提交前进行医保传输
    If int来源 = 2 And Not IsNull(rsPati!险类) Then
        If gclsInsure.GetCapability(support记帐上传, tyAdviceProperty.lng病人ID, rsPati!险类) And Not gclsInsure.GetCapability(support记帐完成后上传, tyAdviceProperty.lng病人ID, rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, tyAdviceProperty.strNO, 2, 1, strMsg, , rsPati!险类) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    '在提交后进行医保传输
    If int来源 = 2 And Not IsNull(rsPati!险类) Then
        If gclsInsure.GetCapability(support记帐上传, tyAdviceProperty.lng病人ID, rsPati!险类) And gclsInsure.GetCapability(support记帐完成后上传, tyAdviceProperty.lng病人ID, rsPati!险类) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, tyAdviceProperty.strNO, 2, 1, strMsg, , rsPati!险类) Then
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                Else
                    MsgBox "单据""" & tyAdviceProperty.strNO & """的数据向医保传送失败,该单据已保存！", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    On Error GoTo 0
    SaveBuildingMainFee = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function


Private Function CheckBuldingFeeValied(ByRef tyAdviceProperty As ty_adviceProperty, ByRef rsPrice As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查主费用的合法性
    '入参:tyAdviceProperty-医嘱相关属性
    '     rsPrice-主费用相关价目
    '返回:生成主费用条件合法，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-07 15:50:34
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivsAnnexFee As String
    On Error GoTo errHandle
    
    If rsPrice Is Nothing Then Exit Function
    If rsPrice.State <> 1 Then Exit Function
    If rsPrice.RecordCount = 0 Then
        MsgBox "该执行项目没有可以计费的主费用。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
        Exit Function
    End If
    If tyAdviceProperty.int执行状态 = 1 Then
        MsgBox "该执行项目已经执行完成，不能生成主费用。", vbInformation, gstrSysName
        Exit Function
    End If
    If tyAdviceProperty.blnDataMoveed Then
        MsgBox "该病人的本次" & IIf(tyAdviceProperty.int病人来源 = 2, "住院", "就诊") & "数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If tyAdviceProperty.int记录性质 = 1 Then
        If BillExistBalance(tyAdviceProperty.strNO) Then
            MsgBox "单据 " & tyAdviceProperty.strNO & " 已经收费，不能再生成这张单据的主费用。" & vbCrLf & "如果需要，你可以手工补充附加费用。", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf tyAdviceProperty.int记录性质 = 2 Then
        '住院出院病人费用限制
        If tyAdviceProperty.int病人来源 = 2 Then
            strPrivsAnnexFee = GetInsidePrivs(p医嘱附费管理)
            If Not PatiCanBilling(tyAdviceProperty.lng病人ID, tyAdviceProperty.lng主页Id, strPrivsAnnexFee, p医嘱附费管理) Then Exit Function
        End If
    End If
    CheckBuldingFeeValied = True
    Exit Function
errHandle:
  If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function


Private Function GetAdvicePrice(ByRef tyAdviceProperty As ty_adviceProperty, _
    ByRef rsPrice As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定医嘱的计价关系到临时记录集
    '入参:tyAdviceProperty-医嘱相关信息
    '出参:rsPrice-返回收费的相关信息
    '返回:读取计价关系成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-07 15:41:41
    '说明:要计算的项目应该不是叮嘱,院外执行,无需计费
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl数量 As Double, bln附加手术 As Boolean
    Dim strSql As String, strIF As String, strPrice As String
    Dim blnHaveSub As Boolean, lng主收入ID As Long
    Dim cur合计 As Currency, i As Long, j As Long
    Dim str药品价格等级 As String, str卫材价格等级 As String, str普通价格等级 As String
    Dim strWherePriceGrade As String
    
    Set rsPrice = New ADODB.Recordset
    rsPrice.Fields.Append "医嘱ID", adBigInt
    rsPrice.Fields.Append "开嘱科室ID", adBigInt
    rsPrice.Fields.Append "类别", adVarChar, 10
    rsPrice.Fields.Append "收费细目ID", adBigInt
    rsPrice.Fields.Append "计算单位", adVarChar, 100, adFldIsNullable
    rsPrice.Fields.Append "附加手术", adInteger
    rsPrice.Fields.Append "执行科室", adInteger
    rsPrice.Fields.Append "收入项目ID", adBigInt
    rsPrice.Fields.Append "收据费目", adVarChar, 50, adFldIsNullable
    rsPrice.Fields.Append "数量", adDouble
    rsPrice.Fields.Append "单价", adDouble
    rsPrice.Fields.Append "应收", adCurrency
    rsPrice.Fields.Append "实收", adCurrency
    rsPrice.Fields.Append "从项", adInteger

    rsPrice.CursorLocation = adUseClient
    rsPrice.LockType = adLockOptimistic
    rsPrice.CursorType = adOpenStatic
    rsPrice.Open

    On Error GoTo errH

    '读取要计算主费用的医嘱记录
    '包含检查部位，附加手术，检验组合的费用
    If tyAdviceProperty.bln独立执行 Then
        strIF = " And B.ID=[1]"
    Else
        'F,D:检查,手术
        'C:检验
        strIF = " And (B.ID=[1] Or (B.相关ID=[1] And B.诊疗类别 IN('F','D')) Or (B.相关ID=[2] And B.诊疗类别='C'))"
    End If

    strSql = _
    " Select B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID,A.执行部门ID," & _
    "        Nvl(A.发送数次,Sum(Nvl(C.本次数次,0))) as 数量,B.标本部位,B.检查方法,B.执行标记,b.病人ID,b.主页ID" & _
    " From 病人医嘱发送 A,病人医嘱记录 B,病人医嘱执行 C" & _
    " Where Nvl(A.计费状态,0)=0 " & strIF & _
    "   And A.医嘱ID=B.ID And A.发送号=[3]" & _
    "   And C.医嘱ID(+)=A.医嘱ID And C.发送号(+)=A.发送号" & _
    " Group by B.序号,A.医嘱ID,B.相关ID,B.诊疗类别,B.诊疗项目ID,B.开嘱科室ID," & _
    "       A.执行部门ID,A.发送数次,B.标本部位,B.检查方法,B.执行标记,b.病人ID,b.主页ID" & _
    " Having Nvl(A.发送数次,Sum(Nvl(C.本次数次,0)))<>0" & _
    " Order by 序号"
    If tyAdviceProperty.blnDataMoveed Then
        strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
        strSql = Replace(strSql, "病人医嘱发送", "H病人医嘱发送")
        strSql = Replace(strSql, "病人医嘱执行", "H病人医嘱执行")
    End If
    Set rsAdvice = gobjDatabase.OpenSQLRecord(strSql, "获取医嘱发送信息", tyAdviceProperty.lng医嘱ID, tyAdviceProperty.lng相关ID, tyAdviceProperty.lng发送号)
    
    If rsAdvice.RecordCount > 0 Then
        '医嘱ID只有一个，肯定是同一个病人
        If GetPriceGradeStartType() > 0 Then
            Call GetPriceGrade(gstrNodeNo, Val(Nvl(rsAdvice!病人ID)), Val(Nvl(rsAdvice!主页ID)), "", str药品价格等级, str卫材价格等级, str普通价格等级)
        End If
        If str药品价格等级 <> "" Or str卫材价格等级 <> "" Or str普通价格等级 <> "" Then
            strWherePriceGrade = _
                "      And ((Instr(';5;6;7;', ';' || c.类别 || ';') > 0 And b.价格等级 = [8])" & vbNewLine & _
                "            Or (Instr(';4;', ';' || c.类别 || ';') > 0 And b.价格等级 = [9])" & vbNewLine & _
                "            Or (Instr(';4;5;6;7;', ';' || c.类别 || ';') = 0 And b.价格等级 = [10])" & vbNewLine & _
                "            Or (b.价格等级 Is Null" & vbNewLine & _
                "                And Not Exists (Select 1" & vbNewLine & _
                "                                From 收费价目" & vbNewLine & _
                "                                Where b.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                "                                      And ((Instr(';5;6;7;', ';' || c.类别 || ';') > 0 And 价格等级 = [8])" & vbNewLine & _
                "                                            Or (Instr(';4;', ';' || c.类别 || ';') > 0 And 价格等级 = [9])" & vbNewLine & _
                "                                            Or (Instr(';4;5;6;7;', ';' || c.类别 || ';') = 0 And 价格等级 = [10])))))"
        Else
            strWherePriceGrade = " And b.价格等级 Is Null "
        End If
    End If
    
    For i = 1 To rsAdvice.RecordCount
        dbl数量 = Nvl(rsAdvice!数量, 0)

        '读取对应的收费价目:只读取固定对照,且不是变价的对照
        bln附加手术 = (rsAdvice!诊疗类别 = "F" And Not IsNull(rsAdvice!相关ID))
        '由于没有加部位等条件，所以要用Distinct
        strPrice = "" & _
        "Select * From (" & _
        "   Select Distinct C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,c.适用科室id" & _
        "               ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
        "   From 诊疗收费关系 C Where C.诊疗项目ID=[1]" & _
        "           And (C.适用科室ID is Null And C.病人来源 = 0 or C.适用科室ID = [6] And C.病人来源 = [7])" & _
        "   ) Where Nvl(适用科室id, 0) = Top"

        strSql = _
        " Select Nvl(A.从属项目,0) as 从项,A.收费项目ID,A.收费数量,B.收入项目ID,D.收据费目," & _
        "       C.类别,C.计算单位,C.执行科室,Decode(C.是否变价,1,B.缺省价格,B.现价) as 单价,C.屏蔽费别," & _
                IIf(bln附加手术, "Nvl(B.附术收费率,100)/100", "1") & " as 附术率" & _
        " From (" & strPrice & ") A,收费价目 B,收费项目目录 C,收入项目 D," & _
        "      (Select [1] as 诊疗项目ID,Decode([2],0,Null,[2]) as 相关ID," & _
        "               Decode([3],'None',Null,[3]) as 标本部位,Decode([4],'None',Null,[4]) as 检查方法,[5] as 执行标记 From Dual " & _
        "       ) X" & _
        " Where A.诊疗项目ID=X.诊疗项目ID" & _
        "       And (   X.相关ID is Null And X.执行标记 IN(1,2) And A.费用性质=1" & _
        "               Or X.标本部位=A.检查部位 And X.检查方法=A.检查方法 And Nvl(A.费用性质,0)=0" & _
        "               Or X.检查方法 is Null And Nvl(A.费用性质,0)=0 And A.检查部位 is Null And A.检查方法 is Null)" & _
        "       And A.收费项目ID=B.收费细目ID And A.收费项目ID=C.ID And B.收入项目ID=D.ID" & _
        "       And (C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                strWherePriceGrade & vbNewLine & _
        "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
        "       And ((Sysdate Between B.执行日期 and B.终止日期) or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        "       And Nvl(A.固有对照,0)=1 And Nvl(C.是否变价,0)=0" & _
        " Order by 从项"
        
        
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取诊疗收费信息", Val(rsAdvice!诊疗项目ID), Val(Nvl(rsAdvice!相关ID, 0)), _
            CStr(Nvl(rsAdvice!标本部位, "None")), CStr(Nvl(rsAdvice!检查方法, "None")), Val(Nvl(rsAdvice!执行标记, 0)), _
            tyAdviceProperty.lng执行科室ID, tyAdviceProperty.int病人来源, str药品价格等级, str卫材价格等级, str普通价格等级)

        blnHaveSub = False: lng主收入ID = 0: cur合计 = 0
        If Not rsTemp.EOF And gSysPara.bln从项汇总折扣 Then
            rsTemp.Filter = "从项=1"
            If Not rsTemp.EOF Then blnHaveSub = True
            rsTemp.Filter = "从项=0"
            If Not rsTemp.EOF Then lng主收入ID = rsTemp!收入项目ID
            rsTemp.Filter = 0
        End If

        For j = 1 To rsTemp.RecordCount
            rsPrice.AddNew
            rsPrice!医嘱ID = rsAdvice!医嘱ID
            rsPrice!开嘱科室id = rsAdvice!开嘱科室id
            rsPrice!类别 = rsTemp!类别
            rsPrice!收费细目ID = rsTemp!收费项目ID
            rsPrice!计算单位 = Nvl(rsTemp!计算单位)
            rsPrice!附加手术 = IIf(bln附加手术, 1, 0)
            rsPrice!执行科室 = Nvl(rsTemp!执行科室, 0)
            rsPrice!收入项目ID = rsTemp!收入项目ID
            rsPrice!收据费目 = rsTemp!收据费目
            rsPrice!单价 = Format(Nvl(rsTemp!单价, 0), gSysPara.Price_Decimal.strFormt_VB)
            rsPrice!数量 = Format(Nvl(rsTemp!收费数量, 0) * dbl数量, "0.00000")
            rsPrice!应收 = Format(rsPrice!数量 * rsPrice!单价 * rsTemp!附术率, gSysPara.Money_Decimal.strFormt_VB)
            rsPrice!从项 = rsTemp!从项
            If gSysPara.bln从项汇总折扣 And blnHaveSub Then
                rsPrice!实收 = rsPrice!应收
                cur合计 = cur合计 + rsPrice!实收
            ElseIf Nvl(rsTemp!屏蔽费别, 0) = 0 Then
                rsPrice!实收 = Format(ActualMoney(tyAdviceProperty.str费别, rsPrice!收入项目ID, rsPrice!应收, _
                    rsTemp!收费项目ID, Nvl(rsAdvice!执行部门ID, 0), Nvl(rsTemp!收费数量, 0) * dbl数量, 0), gSysPara.Money_Decimal.strFormt_VB)
            Else
                rsPrice!实收 = rsPrice!应收
            End If
            rsPrice.Update
            rsTemp.MoveNext
        Next

        If gSysPara.bln从项汇总折扣 And blnHaveSub And lng主收入ID <> 0 Then
            cur合计 = Format(ActualMoney(tyAdviceProperty.str费别, lng主收入ID, cur合计), gSysPara.Money_Decimal.strFormt_VB) - cur合计
            rsPrice.Filter = "从项=0"
            rsPrice!实收 = Nvl(rsPrice!实收, 0) + cur合计
            rsPrice.Update
            rsPrice.Filter = 0
        End If
        rsAdvice.MoveNext
    Next
    If rsPrice.RecordCount > 0 Then rsPrice.MoveFirst
     GetAdvicePrice = True
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Set rsPrice = Nothing
End Function

Private Function GetAdviceProperty(ByVal lng科室id As Long, ByVal lng医嘱ID As Long, _
    ByVal lng发送号 As Long, ByVal bln独立执行 As Boolean, ByVal blnMoveed As Boolean, _
    ByRef tyAdviceProperty As ty_adviceProperty) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医嘱相关属性
    '入参:lng医嘱ID-医嘱ID
    '     lng发送号-发送号
    '出参:tyAdviceProperty-返回医嘱相关属性
    '返回:获取成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-03-07 15:31:13
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tyTempAdviceProperty As ty_adviceProperty
    Dim strSql As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    tyAdviceProperty = tyTempAdviceProperty
    
    strSql = _
    " Select A.病人ID,A.主页ID,A.挂号单,A.病人科室ID,D.当前病区id,A.开嘱科室ID,b.执行部门ID,A.病人来源,C.结算模式,A.诊疗类别,E.计价性质," & _
    "       Decode(A.诊疗类别,'D',Nvl(A.相关ID,A.ID),A.相关ID) as 相关ID,B.NO,B.记录性质,Nvl(B.门诊记帐,0) as 门诊记帐," & _
    "       B.执行状态,B.发送时间,Nvl(D.费别,C.费别) as 费别,d.审核标志 " & _
    " From 病人信息 C,病案主页 D,病人医嘱记录 A,病人医嘱发送 B,诊疗项目目录 E" & _
    " Where A.ID=B.医嘱ID And A.ID=[1] And B.发送号=[2] And A.诊疗项目ID=E.ID" & _
    " And A.病人ID=C.病人ID And A.病人ID=D.病人ID(+) And A.主页ID=D.主页ID(+)"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取医嘱诊疗信息", lng医嘱ID, lng发送号)
    If rsTemp.RecordCount = 0 Then GetAdviceProperty = False: Exit Function
    
    With tyAdviceProperty
        .lng医嘱ID = lng医嘱ID
        .lng发送号 = lng发送号
        .bln独立执行 = bln独立执行
        .lng执行科室ID = IIf(lng科室id = 0, Val(Nvl(rsTemp!执行部门ID)), lng科室id)
        .lng相关ID = Val(Nvl(rsTemp!相关ID))
        .lng病人ID = Val(Nvl(rsTemp!病人ID))
        .lng主页Id = Val(Nvl(rsTemp!主页ID))
        .str挂号单 = Nvl(rsTemp!挂号单)
        .lng病人科室ID = Val(Nvl(rsTemp!病人科室id))
        .lng病人病区ID = Val(Nvl(rsTemp!当前病区ID))
        .lng开嘱科室ID = Val(Nvl(rsTemp!开嘱科室id))
        .int记录性质 = Val(Nvl(rsTemp!记录性质))
        .int审核标志 = Val(Nvl(rsTemp!审核标志))
        .int结算模式 = Val(Nvl(rsTemp!结算模式))
        .str诊疗类别 = Nvl(rsTemp!诊疗类别)
        .strNO = Nvl(rsTemp!NO)
        .int执行状态 = Val(Nvl(rsTemp!执行状态))
        .dat发送时间 = rsTemp!发送时间
        .str费别 = Nvl(rsTemp!费别)
        .lng计价性质 = Val(Nvl(rsTemp!计价性质))
        .str计费状态 = GetSendFeeState(.lng医嘱ID, .lng相关ID, .lng发送号, .bln独立执行, .str诊疗类别, .blnDataMoveed)
        .int病人来源 = Val(Nvl(rsTemp!病人来源))
        .bln门诊记帐 = Val(Nvl(rsTemp!门诊记帐))
         '门诊和住院医生站可发送门诊记帐，存在门诊费用记录中
         '以前的门诊医生站发送为门诊记帐时，rsTemp!门诊记帐的值为空，未修正历史数据
        .strFeeTab = "门诊费用记录"
        If .int病人来源 = 2 Then
            If .bln门诊记帐 Then
                .int病人来源 = 1    '当成门诊病人(这种情况一般是门诊留观病人)
            Else
                .strFeeTab = "住院费用记录"
            End If
        End If
        '检验组合和多部位检查项目的综合执行状态
        If (.str诊疗类别 = "C" Or .str诊疗类别 = "D") And Not .bln独立执行 Then
            strSql = "" & _
            "   Select 执行状态 From 病人医嘱发送 " & _
            "   Where 发送号=[1]  And 医嘱ID IN(Select ID From 病人医嘱记录 Where (ID=[2] Or 相关ID=[2]) And 诊疗类别 In('C','D'))"
            Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取医嘱的执行状态", lng发送号, IIf(.lng相关ID <> 0, .lng相关ID, .lng医嘱ID))
            strSql = ""
            Do While Not rsTemp.EOF
                If InStr(strSql, Nvl(rsTemp!执行状态, 0)) = 0 Then
                    strSql = strSql & Nvl(rsTemp!执行状态, 0)
                End If
                rsTemp.MoveNext
            Loop
            .int执行状态 = IIf(Len(strSql) = 1, Val(strSql), 3)
        End If
    End With
    GetAdviceProperty = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Private Function GetSendFeeState(ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal lng发送号 As Long, ByVal bln独立执行 As Boolean, _
    ByVal str诊疗类别 As String, Optional ByVal blnMoved As Boolean) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定医嘱某次发送之后的计费状态
    '入参:lng医嘱ID-检验组合主项目,手术主项目,第一个检验项目的医嘱ID(即在医技站中显示的项目的)
    '     lng发送号-发送号
    '     bln单独执行-组合项目是否独立执行
    '出参:
    '返回:",-1,0,1,"：其中-1=无需计费,1=已计费,0=未计费,对于门诊单据，2=部分收费,3=全部收费
    '编制:刘兴洪
    '日期:2014-05-27 09:33:50
    '说明:获取指定医嘱某次发送之后的计费状态，主要考虑一些组合医嘱有多种计费的状态
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset, strSql As String

    On Error GoTo errH

    If bln独立执行 Then
        strSql = "Select Distinct 计费状态 From 病人医嘱发送 Where 医嘱ID=[1] And 发送号=[2]"
    Else
        strSql = "Select ID From 病人医嘱记录 Where (ID=[3] Or 相关ID=[3]) And 诊疗类别=[4]"
        strSql = "Select Distinct 计费状态 From 病人医嘱发送 Where 医嘱ID IN(" & strSql & ") And 发送号=[2]"
    End If
    If blnMoved Then
        strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
        strSql = Replace(strSql, "病人医嘱发送", "H病人医嘱发送")
    End If
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "mdlCISKernel", lng医嘱ID, lng发送号, IIf(lng相关ID <> 0, lng相关ID, lng医嘱ID), str诊疗类别)
    strSql = ""
    Do While Not rsTmp.EOF
        strSql = strSql & "," & IIf(Val("" & rsTmp!计费状态) > 1, 1, Val("" & rsTmp!计费状态))
        rsTmp.MoveNext
    Loop
    If strSql <> "" Then GetSendFeeState = strSql & ","
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function
