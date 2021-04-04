VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockInEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'继承窗体事件
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)

Private WithEvents frmInEPRs As frmDockInEPRs
Attribute frmInEPRs.VB_VarHelpID = -1

Private Sub Class_initialize()
    Set frmInEPRs = New frmDockInEPRs
End Sub

Private Sub Class_Terminate()
    Unload frmInEPRs
    Set frmInEPRs = Nothing
End Sub

Private Sub frmInEPRs_Activate()
    RaiseEvent Activate
End Sub

Private Sub frmInEPRs_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
'-0-小(缺省)，1-大
    Call frmInEPRs.SetFontSize(bytSize)
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = frmInEPRs
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object, Optional ByVal blnInsideTools As Boolean)
'自建工具条
    Call frmInEPRs.zlDefCommandBars(CommandBars, blnInsideTools)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInEPRs.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInEPRs.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long _
    , Optional ByVal bnEdit As Boolean, Optional ByVal blnMoved As Boolean, Optional ByVal lngAdviceID As Long _
    , Optional blnForce As Boolean, Optional ByVal lngUnitID As Long, Optional ByVal intState As Integer) As Long
    '功能:调用刷新指定病人的病历内容，并根据情况提供编辑功能
    '参数:  lngPatiId-病人id;
    '       lngPageId-主页id;
    '       lngDeptId-当前操作部门，注意不是病人当前所在科室或出院科室，需要根据该部门确定可增加的病人病历；
    '       lngUnitID，lngDeptId-当intState=“5-最近转科病人”时为病人原病区或原科室
    '       blnEdit-是否允许编辑，通常当前操作部门不是病人当前所在科室或出院科室，就应该不允许编辑，护士站直接设置为不允许编辑
    '       blnForce-前后连次病人ID,主页ID相同时,是否强制刷新
    '       blnMoved-数据是否被转储
    '       lngAdviceID-医嘱ID,目前为手术模块调用
    '       intState-护士站或医生站传入的病人状态TYPE_PATI_State
    '    ps在院 = 0
'    ps预出 = 1
'    ps出院 = 2
'    ps待诊 = 3          '医生站:待会诊病人(在院)
'    ps已诊 = 4          '医生站:已会诊病人
'    ps最近转出 = 5      '医护站:最近转科或转病区的病人(在院)
'    ps待转入 = 6        '医护站:入科待入住或转病区待入往病人
'转出页面操作intState=5，会诊页面操作如果已会诊完成intState=4,未会诊intState=3，其他页面操作intState=传入病人的当前状态（预出1、在院0、出院2）
'医生站按病区显示 ,intState=5则传入的转科前的科室ID和病区ID，否则传入的当前病人科室ID和病区ID
'医生站按科室显示，不管intstate是多少，固定传入界面操作科室，病区规则同病区显示

    zlRefresh = frmInEPRs.zlRefresh(lngPatiID, lngPageId, lngDeptId, bnEdit, blnForce, blnMoved, lngAdviceID, intState)
End Function

Public Function GetFormOperation() As String
'记录界面选定信息，因为工作站在切换页卡时是释放了对象，换回来时重新初始化刷新的。
    GetFormOperation = frmInEPRs.GetFormOperation
End Function

Public Sub RestoreFormOperation(ByVal strValue As String)
'恢复界面选定信息，工作站在刷新之前调用
    Call frmInEPRs.RestoreFormOperation(strValue)
End Sub
