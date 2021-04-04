VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotifyEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mblnIsChangePane As Boolean

Public Event OnContainerUnLoad(objOldPane As Object)
Public Event OnContainerLoad(objOldPane As Object)


Property Get IsChangePane() As Boolean
    IsChangePane = mblnIsChangePane
End Property

Property Let IsChangePane(ByVal value As Boolean)
    mblnIsChangePane = value
End Property


'通知对视频的容器pane进行切换
Public Sub NotifyContainerUnLoad(objOldPane As Object)
    mblnIsChangePane = False
    RaiseEvent OnContainerUnLoad(objOldPane)
End Sub

'通知载入视频模板到容器
Public Sub NotifyContainerLoad(objOldPane As Object)
    RaiseEvent OnContainerLoad(objOldPane)
End Sub

