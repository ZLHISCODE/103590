VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICaptureEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'视频采集窗口处理事件
Public Sub OnVideoEvent(ByVal lngEventType As Long, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
End Sub
