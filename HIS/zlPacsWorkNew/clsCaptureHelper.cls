VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapLinker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'实现ICapHelper接口，传递给视频采集对象进行对应处理
Implements ICapHelper

Private mlngMainAdviceId As Long
Private mlngReportAdviceId As Long
Private mlngLockAdviceId As Long

'Private mobjVideo As Object

Private mobjMainHelper As ucPacsHelper
Private mobjRepPacsHelpers() As ucPacsHelper



Property Get MainAdviceId() As Long
    MainAdviceId = mlngMainAdviceId
End Property
Property Let MainAdviceId(ByVal value As Long)
    mlngMainAdviceId = value
End Property


Property Get ReportAdviceId() As Long
    ReportAdviceId = mlngReportAdviceId
End Property
Property Let ReportAdviceId(ByVal value As Long)
    mlngReportAdviceId = value
End Property


Property Get LockAdviceId() As Long
    LockAdviceId = mlngLockAdviceId
End Property
Property Let LockAdviceId(ByVal value As Long)
    mlngLockAdviceId = value
End Property


'Property Get Video() As Object
'    Set Video = mobjVideo
'End Property
'Property Set Video(value As Object)
'    Set mobjVideo = value
'End Property


Public Sub AddRepPacsHelper(ByVal lngAdviceId As Long, objPacsHelper As Object)

    Dim i As Long
    Dim lngBound As Long
    
    For i = 1 To UBound(mobjRepPacsHelpers)
        If mobjRepPacsHelpers(i).AdviceId = lngAdviceId Then
            Set mobjRepPacsHelpers(i) = Nothing
            Set mobjRepPacsHelpers(i) = objPacsHelper
            
            Exit Sub
        End If
    Next
    
    '新增PacsHelper数组
    lngBound = UBound(mobjRepPacsHelpers) + 1
    ReDim Preserve mobjRepPacsHelpers(lngBound)
    
    Set mobjRepPacsHelpers(lngBound) = objPacsHelper
End Sub

Public Sub RemoveRepPacsHelper(ByVal lngAdviceId As Long)
    Dim i As Long
    Dim blnStartMove As Boolean
    Dim lngBound As Long
    
    blnStartMove = False
    lngBound = UBound(mobjRepPacsHelpers)
    
    For i = 1 To lngBound
        If mobjRepPacsHelpers(i).AdviceId = lngAdviceId Then
            blnStartMove = True
        End If
        
        If blnStartMove And i + 1 <= lngBound Then
        '开始移动数组
            Set mobjRepPacsHelpers(i) = Nothing
            Set mobjRepPacsHelpers(i) = mobjRepPacsHelpers(i + 1)
        End If
    Next
    
    If blnStartMove Then
        Set mobjRepPacsHelpers(lngBound) = Nothing
        
        ReDim Preserve mobjRepPacsHelpers(lngBound - 1)
    End If
    
End Sub



Public Sub Init()

End Sub




Private Sub Class_Initialize()
    ReDim mobjRepPacsHelpers(0)
End Sub

'接口实现********************************************************************************************

Public Function ICapHelper_IsAllowCapture() As Boolean
    '需要判断检查状态
    '需要判断采集数量
    
End Function

Public Function ICapHelper_GetCustomMainID() As Long
'获取专用采集所需的医嘱ID

End Function

Public Function ICapHelper_SaveImg(img As Object, ByVal strDes As String, _
    Optional ByVal blnIsSave As Boolean = True, _
    Optional ByVal strBackTag As String = "") As Boolean
'保存图像

End Function

Public Function ICapHelper_SaveVideo(ByVal strVideoFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsSave As Boolean = True) As Boolean
'保存视频

End Function


Public Function ICapHelper_SaveAudio(ByVal strVideoFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsSave As Boolean = True) As Boolean
'保存音频
End Function


Public Sub ICapHelper_CapLock()
'采集锁定
End Sub


Public Sub ICapHelper_CapUnlock()
'采集解锁
End Sub

Public Sub ICapHelper_OpenLocker()
'打开锁定患者

End Sub

Public Sub ICapHelper_RefreshTag(ByVal strBackTag As String)
'刷新标记

End Sub



'接口实现********************************************************************************************
