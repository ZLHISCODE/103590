VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWorkModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Property Get AdviceId() As Long

End Property

Property Get DeptId() As Long

End Property


Public Sub zlInitModule(objNotify As IEventNotify, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngDeptId As String)
'��ʼ��ģ��
End Sub


Public Sub zlRefresh(objStudyInfo As clsStudyInfo, ByVal strSelModuleName As String, _
    Optional ByVal blnIsForceRefresh As Boolean = False, Optional ByVal blnIsHistory As Boolean = False)
'ˢ��ģ��
End Sub


Public Function zlGetModule(ByVal strSelModuleName As String, Optional ByVal blnAutoCreate As Boolean = True) As Object
'��ȡģ�����
End Function


'Public Sub zlResetConnection()
''��������
'End Sub