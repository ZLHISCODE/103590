VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalVerification"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function IsValidWithActiveVideo() As Boolean
    IsValidWithActiveVideo = False
    
    If Not gobjVideo Is Nothing Then
        If gobjVideo.hwnd = 0 Then Exit Function
    End If
    
    IsValidWithActiveVideo = True
End Function

 
Public Function QuitActiveVideo() As Long
'    QuitActiveVideo = TerminateProcess(GetCurrentProcessId, 0)
     
    ExitProcess 0
    
End Function


Public Function ActiveVideoProcessId() As Long
    ActiveVideoProcessId = GetCurrentProcessId
End Function
