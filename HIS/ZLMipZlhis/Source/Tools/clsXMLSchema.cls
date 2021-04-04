VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mfrmMakeXMLShema As frmMakeXMLShema

Public Function Initialize(ByVal strDataPath As String) As Boolean
    gstrDataPath = strDataPath
    Initialize = True
End Function

Public Function GetFormObject() As Object
    Set mfrmMakeXMLShema = New frmMakeXMLShema
    Set GetFormObject = mfrmMakeXMLShema
End Function

Public Function ActiveFormObject() As Object
'    mfrmMakeXMLShema.RefreshData
End Function

Public Function CloseFormObject() As Object
    If Not (mfrmMakeXMLShema Is Nothing) Then
        Unload mfrmMakeXMLShema
        Set mfrmMakeXMLShema = Nothing
    End If
End Function

Private Sub Class_Terminate()
    If Not (mfrmMakeXMLShema Is Nothing) Then
        Unload mfrmMakeXMLShema
        Set mfrmMakeXMLShema = Nothing
    End If
End Sub

