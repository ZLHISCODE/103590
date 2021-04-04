VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mfrmChild As frmTest

Public Function Initialize(ByVal strDataPath As String) As Boolean
    Set mfrmChild = New frmTest
    
    gstrDataPath = strDataPath
    Initialize = mfrmChild.Initialize
    
End Function

Public Function GetFormObject() As Object
    Set GetFormObject = mfrmChild
End Function

Public Sub ActiveFormObject()
    Call mfrmChild.ActiveForm
End Sub

Public Function CloseFormObject() As Object
    If Not (mfrmChild Is Nothing) Then
        Unload mfrmChild
        Set mfrmChild = Nothing
    End If
End Function

Private Sub Class_Terminate()
    If Not (mfrmChild Is Nothing) Then
        Unload mfrmChild
        Set mfrmChild = Nothing
    End If
End Sub



