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
    '��Ҫ�жϼ��״̬
    '��Ҫ�жϲɼ�����
End Function

Public Function GetCustomMainID() As Long
'��ȡר�òɼ������ҽ��ID

End Function

Public Function SaveImg(img As Object, ByVal strDes As String, _
    Optional ByVal blnIsNeedSave As Boolean = True, _
    Optional ByVal strBackTag As String = "") As Boolean
'����ͼ��

End Function

Public Function SaveVideo(ByVal strVideoFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsNeedSave As Boolean = True) As Boolean
'������Ƶ

End Function


Public Function SaveAudio(ByVal strAudioFile As String, ByVal strDes As String, _
    Optional ByVal strEncodeName As String, Optional ByVal lngRecordTimeLen As Long, _
    Optional ByVal blnIsNeedSave As Boolean = True) As Boolean
'������Ƶ
End Function


Public Sub CapLock(ByRef strLocker As String)
'�ɼ�����
End Sub


Public Sub CapUnlock()
'�ɼ�����
End Sub

'Public Sub OpenLocker()
''����������
'
'End Sub

Public Sub AfterTag(ByRef strBackTag As String)
'ˢ�±��

End Sub
