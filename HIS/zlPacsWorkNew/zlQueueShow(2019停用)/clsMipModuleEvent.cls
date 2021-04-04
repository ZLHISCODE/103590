VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMipModuleEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ResponseEvent(ByVal strEventName As String, ByVal objCollEventPara As Collection)
 
Public Sub ActionEvent(ByVal strEventName As String, ByVal objCollEventPara As Collection)
    RaiseEvent ResponseEvent(strEventName, objCollEventPara)
End Sub

