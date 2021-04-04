VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockInTendEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mfrmInTendEPR As frmDockInTendEPR
Attribute mfrmInTendEPR.VB_VarHelpID = -1

Private Sub Class_initialize()
    Dim strPrivs As String
    
    strPrivs = GetPrivFunc(glngSys, 1255)
    Set mfrmInTendEPR = New frmDockInTendEPR
    Call mfrmInTendEPR.InitData(strPrivs)
End Sub

Private Sub Class_Terminate()
    Unload mfrmInTendEPR
    Set mfrmInTendEPR = Nothing
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
    Call mfrmInTendEPR.SetFontSize(bytSize)
End Sub
Public Function zlGetForm() As Object
    Set zlGetForm = mfrmInTendEPR
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object, Optional ByVal blnInsideTools As Boolean)
'自建工具条
    Call mfrmInTendEPR.zlDefCommandBars(CommandBars, blnInsideTools)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmInTendEPR.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmInTendEPR.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngDeptId As Long, _
                        ByVal bnEdit As Boolean, ByVal blnForce As Boolean, ByVal blnMoved As Boolean, Optional ByVal blnDoctorStation As Boolean = False) As Long
    '功能:调用刷新指定病人的病历内容，并根据情况提供编辑功能
    '参数:  lngPatiId-病人id;
    '       lngPageId-主页id;
    '       lngDeptId-当前操作部门，注意不是病人当前所在科室或出院科室，需要根据该部门确定可增加的病人病历；
    '       blnEdit-是否允许编辑，通常当前操作部门不是病人当前所在科室或出院科室，就应该不允许编辑
    '       blnForce-前后连次病人ID,主页ID相同时,是否强制刷新
    '       blnMoved-数据是否被转储
    zlRefresh = mfrmInTendEPR.RefreshData(lngPatiID, lngPageId, lngDeptId, blnDoctorStation, bnEdit, blnForce, blnMoved)
End Function

Public Function GetFormOperation() As String
'记录界面选定信息，因为工作站在切换页卡时是释放了对象，换回来时重新初始化刷新的。
    GetFormOperation = mfrmInTendEPR.GetFormOperation
End Function
Public Sub RestoreFormOperation(ByVal strValue As String)
'恢复界面选定信息，工作站在刷新之前调用
    Call mfrmInTendEPR.RestoreFormOperation(strValue)
End Sub