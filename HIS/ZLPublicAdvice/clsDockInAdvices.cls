VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockInAdvices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event Activate() '子窗体激活
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字
Public Event ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean) '要求查看报告
Public Event PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean) '要求打印报告
Public Event ViewPACSImage(ByVal 医嘱ID As Long) '要求进行观片
Public Event ExecLogNew(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, 完成 As Boolean) '执行情况登记
Public Event ExecLogModi(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, ByVal 执行时间 As String, 完成 As Boolean) '执行情况修改
Public Event EditDiagnose(ParentForm As Object, ByVal 病人ID As Long, ByVal 主页ID As Long, ByVal 科室ID As Long, ByVal str类型 As String, Succeed As Boolean) '编辑住院诊断
Public Event SetEditState(ByVal blnEditState As Boolean)    '编辑状态时禁用菜单和可转移焦点的功能
Public Event DoByAdvice(ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal lngWayID As Long, ByVal strTag As String)

Private WithEvents mclsInAdvice As zlCISKernel.clsDockInAdvices
Attribute mclsInAdvice.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mclsInAdvice = New zlCISKernel.clsDockInAdvices
End Sub

Private Sub Class_Terminate()
    Set mclsInAdvice = Nothing
End Sub

Private Sub mclsInAdvice_Activate()
    RaiseEvent Activate
End Sub

Private Sub mclsInAdvice_SetEditState(ByVal blnEditState As Boolean)
    RaiseEvent SetEditState(blnEditState)
End Sub

Private Sub mclsInAdvice_RequestRefresh(ByVal RefreshNotify As Boolean)
    RaiseEvent RequestRefresh(RefreshNotify)
End Sub

Private Sub mclsInAdvice_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Private Sub mclsInAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(报告ID, CanPrint)
End Sub

Private Sub mclsInAdvice_PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean)
    RaiseEvent PrintEPRReport(报告ID, Preview)
End Sub

Private Sub mclsInAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
    RaiseEvent ViewPACSImage(医嘱ID)
End Sub

Private Sub mclsInAdvice_ExecLogNew(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, 完成 As Boolean)
    RaiseEvent ExecLogNew(医嘱ID, 发送号, 科室ID, 完成)
End Sub

Private Sub mclsInAdvice_ExecLogModi(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, ByVal 执行时间 As String, 完成 As Boolean)
    RaiseEvent ExecLogModi(医嘱ID, 发送号, 科室ID, 执行时间, 完成)
End Sub

Private Sub mclsInAdvice_EditDiagnose(ParentForm As Object, ByVal 病人ID As Long, ByVal 主页ID As Long, ByVal 科室ID As Long, ByVal str类型 As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, 病人ID, 主页ID, 科室ID, str类型, Succeed)
End Sub

Private Sub mclsInAdvice_DoByAdvice(ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal lngWayID As Long, ByVal strTag As String)
    RaiseEvent DoByAdvice(lng医嘱ID, lng相关ID, lngWayID, strTag)
End Sub

Public Function zlGetForm() As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取
    '初始化医保部件:可分散到具体的GetForm中去初始化
    Set zlGetForm = mclsInAdvice.zlGetForm
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int场合 As Integer, _
    Optional ByVal blnInsideTools As Boolean = False, Optional ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
'功能：定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
'参数：int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      CommandBars=仅用于查看时可以不传(传入Nothing)
'      blnInsideTools=是否将按钮显示到窗体内部工具栏上（仅护士站调用时有效）
'      objSquareCard=通过医技工作站传入的一卡通结算部件(住院医生站在发送时再创建并初始化)

    Call mclsInAdvice.zlDefCommandBars(frmParent, CommandBars, int场合, blnInsideTools, objSquareCard, blnModalNew)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mclsInAdvice.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mclsInAdvice.zlExecuteCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBarsDirect(ByVal Control As CommandBarControl, ByRef frmParent As Object, ByRef strPrivs As String, _
    ByVal bln批量 As Boolean, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt婴儿 As Byte, _
    ByVal lng病区ID As Long, ByVal lng科室id As Long, ByVal lng前提ID As Long, ByVal lng界面科室ID As Long, ByVal int场合 As Integer, _
    ParamArray arrPar() As Variant)
'功能：提供单独调用医嘱操作的接口
    Call mclsInAdvice.zlExecuteCommandBarsDirect(Control, frmParent, strPrivs, bln批量, lng病人ID, lng主页ID, byt婴儿, _
        lng病区ID, lng科室id, lng前提ID, lng界面科室ID, int场合, arrPar)
End Sub

Public Sub zlCheckPrivs(ByVal Control As CommandBarControl, ByVal int场合 As Integer)
'功能：检查菜单或按钮的权限，并设置其可见性
    Call mclsInAdvice.zlCheckPrivs(Control, int场合)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mclsInAdvice.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal lng科室id As Long, _
    ByVal int状态 As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal lng前提ID As Long, _
    Optional ByVal int执行状态 As Integer, Optional ByVal lng界面科室ID As Long, Optional ByVal lng路径状态 As Long = -1, _
    Optional ByVal lng医护科室ID As Long, Optional ByRef objMip As Object, Optional ByVal int婴儿 As Integer = -1, Optional ByVal lng前提科室ID As Long, Optional ByVal lng会诊医嘱ID As Long)
'功能：刷新住院医嘱数据
'参数：int状态=0-在院病人,1-预出院病人,2-出院病人,3-待会诊病人(在院),4-已会诊病人,5-最近转科病人(在院),6-转科或转病区待入住病人(不含入院待入住)
'      lng病区ID，lng科室ID=当“5-最近转科病人”时为病人原病区或原科室
'      lng前提ID=当由医技站调用时传入
'      int执行状态=当由医技站调用时传入,医技执行项目的执行状态
'      lng界面科室ID=如果当前医生站是会诊病人，则为会诊科室ID；如果是医技站调用，则为医技科室ID
'      lng路径状态=-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
'      blnMoved=该病人的数据是否已转出
'      lng医护科室ID=医护站界面ID
'      str前提IDs医技站中在当前科室执行的所有医嘱
'      objMip 消息对象 zl9ComLib.clsMipModule
    Call mclsInAdvice.zlRefresh(lng病人ID, lng主页ID, lng病区ID, lng科室id, int状态, blnMoved, lng前提ID, int执行状态, lng界面科室ID, lng路径状态, lng医护科室ID, objMip, int婴儿, lng前提科室ID, lng会诊医嘱ID)
End Sub

Public Sub LocatedAdviceRow(ByVal lng医嘱ID As Long)
'功能：根据工作站传入的医嘱id在选项卡界查找定位医嘱
    Call mclsInAdvice.LocatedAdviceRow(lng医嘱ID)
End Sub

Public Sub zlSeekAndViewEPRReport(ByVal lng报告ID As Long)
'功能：定位到报告对应的医嘱，并打开报告查看
    Call mclsInAdvice.zlSeekAndViewEPRReport(lng报告ID)
End Sub

Public Sub zlAdvicePrint(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long)
'功能：打印医嘱单
    Call mclsInAdvice.zlAdvicePrint(frmParent, lng病人ID, lng主页ID)
End Sub

Public Sub zlInitPath(ByRef clsDockPath As Object)
'功能：初始化临床路径对象，传入医生站或护士站引用的对象
    Call mclsInAdvice.zlInitPath(clsDockPath)
End Sub

Public Sub zlInitPlugIn(ByVal objPlugIn As Object)
'功能:处理化外挂插件对象
'53740:刘鹏飞,2012-09-19
    Call mclsInAdvice.zlInitPlugIn(objPlugIn)
End Sub

Public Sub zlInitMip(ByRef objMip As Object)
'功能:处理消息对象
'参数：objMip zl9ComLib.clsMipModule
    Call mclsInAdvice.zlInitMip(objMip)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置医嘱清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    Call mclsInAdvice.SetFontSize(bytSize)
End Sub

Public Sub zlCompoundUnpack(ByVal frmParent As Object, ByVal lng病区ID As Long, ByVal lng病人ID As Long, Optional ByVal lng医护科室ID As Long, _
    Optional ByVal lng婴儿科室ID As Long, Optional ByVal lng婴儿病区ID As Long)
'功能：批量打包
    Call mclsInAdvice.zlCompoundUnpack(frmParent, lng病区ID, lng病人ID, lng医护科室ID, lng婴儿科室ID, lng婴儿病区ID)
End Sub

Public Function AdviceRisReport(frmParent As Object, ByVal lng病区ID As Long) As Boolean
'功能：批量打印RIS预约单
    AdviceRisReport = gobjKernel.AdviceRisReport(frmParent, lng病区ID)
End Function

Public Function GetFormOperation() As String
    GetFormOperation = mclsInAdvice.GetFormOperation
End Function

Public Function RestoreFormOperation(ByVal strValue As String)
    Call mclsInAdvice.RestoreFormOperation(strValue)
End Function