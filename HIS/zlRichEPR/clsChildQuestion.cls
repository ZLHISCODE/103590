VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChildQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'继承窗体事件

Private WithEvents mfrmQuestion As frmQuestion
Attribute mfrmQuestion.VB_VarHelpID = -1
Private mfrmMain As Object

'######################################################################################################################

Private Sub Class_initialize()

    Set mfrmQuestion = New frmQuestion

End Sub

Private Sub Class_Terminate()

    If Not (mfrmQuestion Is Nothing) Then Unload mfrmQuestion
    Set mfrmQuestion = Nothing
    Set mfrmMain = Nothing
    
End Sub

Public Sub zlOpenQuestion(ByVal frmMain As Object, ByVal lngPatiID As Long, ByVal lngPageId As Long)
    '功能:  显示指病的电子病历
    '参数:  lngKey-电子病历的ID;
    Set mfrmMain = frmMain
    mfrmQuestion.ShowMe mfrmMain, lngPatiID, lngPageId
End Sub


