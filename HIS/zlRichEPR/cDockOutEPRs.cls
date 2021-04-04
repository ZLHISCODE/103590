VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockOutEPRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'继承窗体事件
Public Event RequestRefresh() '要求主窗体刷新
Public Event Activate()
Public Event ClickDiagRef(DiagnosisID As Long, Modal As Byte)

Private WithEvents frmOutEPRs As frmDockOutEPRs
Attribute frmOutEPRs.VB_VarHelpID = -1

Private Sub Class_initialize()
    Set frmOutEPRs = New frmDockOutEPRs
End Sub

Private Sub Class_Terminate()
    Unload frmOutEPRs
    Set frmOutEPRs = Nothing
End Sub


Private Sub frmOutEPRs_RequestRefresh()
    RaiseEvent RequestRefresh
End Sub

Private Sub frmOutEPRs_Activate()
    RaiseEvent Activate
End Sub

Private Sub frmOutEPRs_ClickDiagRef(DiagnosisID As Long, Modal As Byte)
    RaiseEvent ClickDiagRef(DiagnosisID, Modal)
End Sub
Public Sub SetFontSize(ByVal bytSize As Byte)
    Call frmOutEPRs.SetFontSize(bytSize)
End Sub
Public Function zlGetForm() As Object
    Set zlGetForm = frmOutEPRs
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object)
    Call frmOutEPRs.zlDefCommandBars(CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmOutEPRs.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmOutEPRs.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngBillId As Long, ByVal lngDeptId As Long, Optional ByVal bnEdit As Boolean, _
                            Optional ByVal blnMoved As Boolean, Optional ByVal blnForce As Boolean, Optional ByVal lngAdviceID As Long) As Long
    '功能:调用刷新指定病人的病历内容，并根据情况提供编辑功能
    '参数:  lngPatiId-病人id;
    '       lngBillId-挂号id;
    '       lngDeptId-当前操作部门，注意不是病人本次就诊科室；
    '       blnEdit-是否允许编辑，通常当前操作部门不是病人本次就诊科室，就应该不允许编辑。
    '       blnMoved-数据是否被转储
    '       lngAdviceID 医嘱ID－目前只有手术模块调用传用
    zlRefresh = frmOutEPRs.zlRefresh(lngPatiID, lngBillId, lngDeptId, bnEdit, blnForce, blnMoved, lngAdviceID)
End Function

Public Function zlOpenDefaultEPR(ByVal str挂号单号 As String) As Boolean

    '功能:自动打开新增病历
    '参数:  str挂号单号-挂号单号;
    
    Dim strSQL As String
    Dim bytKind As Byte
    Dim rs As New ADODB.Recordset
        
    strSQL = "Select Nvl(急诊,0) As 急诊,Nvl(复诊,0) As 复诊 From 病人挂号记录 Where No=[1] and 记录性质=1  and 记录状态=1"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "门诊病历", str挂号单号)
    If rs.BOF = False Then
        If rs("复诊").Value = 1 Then
            bytKind = 3
        ElseIf rs("急诊").Value = 1 Then
            bytKind = 2
        Else
            bytKind = 1
        End If
        
        zlOpenDefaultEPR = frmOutEPRs.zlOpenDefaultEPR(bytKind)
        
    End If
    
End Function

Public Function GetFormOperation() As String
'记录界面选定信息，因为工作站在切换页卡时是释放了对象，换回来时重新初始化刷新的。
    GetFormOperation = frmOutEPRs.GetFormOperation
End Function
Public Sub RestoreFormOperation(ByVal strValue As String)
'恢复界面选定信息，工作站在刷新之前调用
    Call frmOutEPRs.RestoreFormOperation(strValue)
End Sub
