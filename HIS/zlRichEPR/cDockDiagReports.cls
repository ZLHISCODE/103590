VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockDiagReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'继承窗体事件
Public Event Activate()

Private WithEvents frmInDiagReport As frmDockDiagReports
Attribute frmInDiagReport.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set frmInDiagReport = New frmDockDiagReports
End Sub

Private Sub Class_Terminate()
    Set frmInDiagReport = Nothing
End Sub

Private Sub frmInTend_Activate()
    RaiseEvent Activate
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = frmInDiagReport
End Function

Public Sub zlDefCommandBars(ByVal frmParent As Object, _
                            ByVal cbsMain As Object, _
                            ByVal int病人来源 As Integer, _
                            Optional ByVal bln护士站 As Boolean)
    Call frmInDiagReport.zlDefCommandBars(frmParent, cbsMain, int病人来源, bln护士站)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInDiagReport.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInDiagReport.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiID As Long, _
                            ByVal strPageID As String, _
                            ByVal lngDeptId As Long, _
                            ByVal int类型 As Integer, _
                            Optional ByVal bnEdit As Boolean, _
                            Optional ByVal bln数据转出 As Boolean) As Long
    '功能:调用刷新指定病人的病历内容，并根据情况提供编辑功能
    '参数:  lngPatiId-病人id;
    '       lngPageId-主页id;
    '       lngDeptId-当前操作部门，注意不是病人当前所在病区或出院病区；
    '       blnEdit-是否允许编辑，通常当前操作部门不是病人所在病区或出院病区，就应该不允许编辑，医生站直接设置为不允许编辑
    zlRefresh = frmInDiagReport.zlRefresh(lngPatiID, strPageID, lngDeptId, bnEdit, int类型, bln数据转出)
    
End Function


