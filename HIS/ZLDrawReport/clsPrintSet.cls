VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Activate()             '×Ó´°Ìå¼¤»î
Private WithEvents mfrmChildPrintSet As frmChildPrintSet
Attribute mfrmChildPrintSet.VB_VarHelpID = -1

Public Property Let DataChanged(ByVal blnData As Boolean)
    mfrmChildPrintSet.DataChanged = blnData
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mfrmChildPrintSet.DataChanged
End Property

Private Sub Class_Initialize()
    Set mfrmChildPrintSet = New frmChildPrintSet
End Sub

Private Sub Class_Terminate()
    Set mfrmChildPrintSet = Nothing
End Sub

Private Sub mfrmChildPrintSet_Activate()
    RaiseEvent Activate
End Sub

Public Function Form() As Object
    
    Set Form = mfrmChildPrintSet
    
End Function

Public Function InitData(ByVal frmMain As Object, ByVal strSavePath As String) As Boolean
    
    InitData = mfrmChildPrintSet.InitData(frmMain, strSavePath)
    
End Function

Public Function ValidData() As Boolean
    ValidData = mfrmChildPrintSet.ValidData
End Function

Public Function SaveData() As Boolean
    SaveData = mfrmChildPrintSet.SaveData
End Function

Public Function RefreshData() As Boolean
    
    RefreshData = mfrmChildPrintSet.RefreshData(9, False)

End Function

Public Function GetPaper(ByRef objPaper As USERPAPER, ByVal strSavePath As String) As Boolean
    
    GetPaper = mfrmChildPrintSet.GetPaper(objPaper, strSavePath)
    
End Function


