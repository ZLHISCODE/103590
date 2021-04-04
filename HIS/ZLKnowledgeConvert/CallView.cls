VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CallView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mobjForm As frmForm
Public Function GetControlHWND() As Object
    Set mobjForm = New frmForm
    Set GetControlHWND = mobjForm
End Function

Public Sub InitDataByString(connectionString As String, str As String)
    Call mobjForm.InitDataByString(connectionString, str)
End Sub

Public Sub InitDataByADODB(conn As ADODB.Connection, str As String)
    Call mobjForm.InitDataByADODB(conn, str)

End Sub

Public Sub LoadContent(drugsid As String)
    Call mobjForm.LoadContent(drugsid)
End Sub
