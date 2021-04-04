VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockInAdvices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'¼Ì³Ð´°ÌåÊÂ¼þ
Public Event Activate()

Private WithEvents frmInAdvices As frmDockInAdvices
Attribute frmInAdvices.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set frmInAdvices = New frmDockInAdvices
End Sub

Private Sub Class_Terminate()
    Set frmInAdvices = Nothing
End Sub

Private Sub frmInAdvices_Activate()
    RaiseEvent Activate
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = frmInAdvices
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object)
    Call frmInAdvices.zlDefCommandBars(CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInAdvices.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call frmInAdvices.zlExecuteCommandBars(Control)
End Sub

Public Function zlRefresh(ByVal lngPatiId As Long, ByVal lngPageId As Long) As Long
'    zlRefresh = frmInAdvices.zlRefresh(lngPatiId, lngPageId)
End Function

