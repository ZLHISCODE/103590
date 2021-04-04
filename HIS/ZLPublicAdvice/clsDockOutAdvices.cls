VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDockOutAdvices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'窗体相关的事件
Public Event Activate() '子窗体激活
Public Event RequestRefresh() '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字
Public Event ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean) '要求查看报告
Public Event PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean) '要求打印报告
Public Event ViewPACSImage(ByVal 医嘱ID As Long) '要求进行观片
Public Event EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean) '编辑门诊诊断
Public Event CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str疾病ID As String, ByVal str诊断Id As String, ByRef blnNo As Boolean) '根据诊断检查是否书写传染病报告卡
Public Event VSKeyPress(KeyAscii As Integer)
Private WithEvents mclsOutAdvice As zlCISKernel.clsDockOutAdvices
Attribute mclsOutAdvice.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set mclsOutAdvice = New zlCISKernel.clsDockOutAdvices
End Sub

Private Sub Class_Terminate()
    Set mclsOutAdvice = Nothing
End Sub

Private Sub mclsOutAdvice_VSKeyPress(KeyAscii As Integer)
    RaiseEvent VSKeyPress(KeyAscii)
End Sub

Private Sub mclsOutAdvice_Activate()
    RaiseEvent Activate
End Sub

Private Sub mclsOutAdvice_CheckInfectDisease(ByVal blnOnChek As Boolean, ByVal str疾病ID As String, ByVal str诊断Id As String, ByRef blnNo As Boolean)
    RaiseEvent CheckInfectDisease(blnOnChek, str疾病ID, str诊断Id, blnNo)
End Sub

Private Sub mclsOutAdvice_EditDiagnose(ParentForm As Object, ByVal 挂号单 As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, 挂号单, Succeed)
End Sub

Private Sub mclsOutAdvice_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub mclsOutAdvice_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Private Sub mclsOutAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
    RaiseEvent ViewEPRReport(报告ID, CanPrint)
End Sub

Private Sub mclsOutAdvice_PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean)
    RaiseEvent PrintEPRReport(报告ID, Preview)
End Sub

Private Sub mclsOutAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
    RaiseEvent ViewPACSImage(医嘱ID)
End Sub

Public Function zlGetForm() As Object
'功能：获取子业务窗口,由主程序初始化时只调用一次获取
    '初始化医保部件:可分散到具体的GetForm中去初始化
    Set zlGetForm = mclsOutAdvice.zlGetForm
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal CommandBars As Object, ByVal int场合 As Integer, _
    Optional ByRef objPlugIn As Object, Optional ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
'功能：定义子窗体的菜单和工具条(包括主窗体要使用的菜单和工具条)
'参数：int场合=0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      CommandBars=仅用于查看时可以不传(传入Nothing)
'      objPlugIn=通过门诊医生工作站传入的外挂功能对象
'      objSquareCard=通过门诊医生站和医技工作站传入的一卡通结算部件
    Call mclsOutAdvice.zlDefCommandBars(frmParent, CommandBars, int场合, objPlugIn, objSquareCard, blnModalNew)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As Object)
    Call mclsOutAdvice.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As Object)
    Call mclsOutAdvice.zlExecuteCommandBars(Control)
End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As Object)
    Call mclsOutAdvice.zlPopupCommandBars(CommandBar)
End Sub

Public Sub zlInitPath(ByRef clsDockOutPath As Object)
'功能：初始化临床路径对象，传入医生站或护士站引用的对象
    Call mclsOutAdvice.zlInitPath(clsDockOutPath)
End Sub

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal str挂号单 As String, ByVal blnEditable As Boolean, _
    Optional ByVal blnMoved As Boolean, Optional ByVal lng前提ID As Long, Optional ByVal lng界面科室ID As Long, _
    Optional ByRef objMip As Object, Optional ByVal lng前提科室ID As Long, Optional ByVal lng路径状态 As Long = -1, Optional ByVal int就诊类型 As Integer)
'功能：刷新住院医嘱数据
'参数：lng前提ID=当由医技站调用时传入
'      blnMoved=该病人的数据是否已转出
'      blnEditable=可否对病人医嘱进行编辑
'      lng界面科室ID=如果是医技站调用，则为医技科室ID
'      str前提IDs医技站中在当前科室执行的所有医嘱
'      objMip 消息对象 zl9ComLib.clsMipModule
'      lng路径状态=-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
'      int就诊类型= pt候诊 = 0；pt就诊 = 1；pt已诊 = 2；pt转诊 = 3；pt预约 = 4；pt回诊 = 5
    Call mclsOutAdvice.zlRefresh(lng病人ID, str挂号单, blnEditable, blnMoved, lng前提ID, lng界面科室ID, objMip, lng前提科室ID, lng路径状态, int就诊类型)
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置医嘱清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    Call mclsOutAdvice.SetFontSize(bytSize)
End Sub
