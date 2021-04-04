VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TReportEvtActionType
    ratAfterOpen = 1    '打开后
    ratBeforeEdit = 2   '编辑前
    ratAfterClosed = 3  '关闭后
    ratAfterPrinted = 4 '打印后
    ratAfterDeleted = 5 '删除后
    ratAfterSaved = 6   '保存后
    ratAfterReportStateChanged = 7  '报告状态改变后
End Enum

Public Event OnReportAction(ByVal lngReportEvtAction As TReportEvtActionType, ByVal lngOrderID As Long, ByVal strDocID As String, ByVal lngReportState As Long)
Public Event OnDataRefresh(ByVal strDocID As String)
Public Event OnDocPluginAction(ByVal strPluginID As String, ByVal strPluginName As String, ByVal lngActionType As Long, ByVal strData As String, ByVal strTag As String)
Public Event OnImageSaveToFTP(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal strImageName As String, ByVal strImageBase64 As String, ByVal strFTPDeviceNo As String, ByVal strRptCreateDate As String, ByRef blnResult As Boolean)
Public Event OnImageLoadFromFTP(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal strImageName As String, ByVal strImagePath As String, ByRef blnResult As Boolean)

Public Sub DoReportAction(ByVal lngReportEvtAction As TReportEvtActionType, ByVal lngOrderID As Long, ByVal strDocID As String, ByVal lngReportState As Long)

    On Error GoTo errHandle
    RaiseEvent OnReportAction(lngReportEvtAction, lngOrderID, strDocID, lngReportState)
    Exit Sub
    
errHandle:
    OutputDebugString "zlPacsVbCommon.DoReportAction:" + err.Description
End Sub

Public Sub DoDataRefresh(ByVal strDocID As String)
    On Error GoTo errHandle
    RaiseEvent OnDataRefresh(strDocID)
    Exit Sub
errHandle:
    OutputDebugString "zlPacsVbCommon.DoDataRefresh:" + err.Description
End Sub

Public Sub DoDocPluginAction(ByVal strPluginID As String, ByVal strPluginName As String, ByVal lngActionType As Long, ByVal strData As String, ByVal strTag As String)

    On Error GoTo errHandle
    RaiseEvent OnDocPluginAction(strPluginID, strPluginName, lngActionType, strData, strTag)
    Exit Sub
errHandle:
    OutputDebugString "zlPacsVbCommon.DoDocPluginAction:" + err.Description
End Sub

Public Function DoImageSaveToFTP(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal strImageName As String, ByVal strImageBase64 As String, ByVal strFTPDeviceNo As String, ByVal strRptCreateDate As String) As Boolean
    Dim blnResult As Boolean
    
    On Error GoTo errHandle
    RaiseEvent OnImageSaveToFTP(lngOrderID, strDocID, strImageName, strImageBase64, strFTPDeviceNo, strRptCreateDate, blnResult)
    DoImageSaveToFTP = blnResult
    Exit Function
errHandle:
    OutputDebugString "zlPacsVbCommon.DoImageSaveToFTP:" + err.Description
End Function

Public Function DoImageLoadFromFTP(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal strImageName As String, ByVal strImagePath As String) As Boolean
    Dim blnResult As Boolean
    
    On Error GoTo errHandle
    
    RaiseEvent OnImageLoadFromFTP(lngOrderID, strDocID, strImageName, strImagePath, blnResult)
    DoImageLoadFromFTP = blnResult
    Exit Function
errHandle:
    OutputDebugString "zlPacsVbCommon.DoImageLoadFromFTP:" + err.Description
End Function







