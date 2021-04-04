VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICapHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function IsAllowCapture() As Boolean
    '需要判断检查状态
    '需要判断采集数量
End Function

Public Function GetCustomMainID() As Long
'获取专用采集所需的医嘱ID

End Function

Public Function SaveImg(img As Object, ByVal strDes As String, _
    Optional ByVal blnIsNeedSave As Boolean = True, _
    Optional ByVal strBackTag As String = "") As Boolean
'保存图像

End Function

Public Function SaveVideo(ByVal strVideoFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsNeedSave As Boolean = True) As Boolean
'保存视频

End Function


Public Function SaveAudio(ByVal strAudioFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsNeedSave As Boolean = True) As Boolean
'保存音频
End Function


Public Sub CapLock(ByRef strLocker As String)
'采集锁定
End Sub


Public Sub CapUnlock()
'采集解锁
End Sub

'Public Sub OpenLocker()
''打开锁定患者
'
'End Sub

Public Sub AfterTag(ByRef strBackTag As String)
'刷新标记

End Sub

