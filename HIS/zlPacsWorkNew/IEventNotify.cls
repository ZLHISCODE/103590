VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEventNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TInfoType
    infNone = -1
    infHint = -2         '��ʾ
    infWaring = -3   '����
    infNormalErr = -4    '��ͨ����
    infDataErr = -5      '���ݿ����
End Enum

Public Function Owner() As Object
'ʵ�ֽӿڵĶ���ӵ����
End Function


Public Function hwnd() As Long
'ʵ�ֽӿڵĶ�����
End Function


'Public Function MainPrivs() As String
''��ȡȨ�޴�
'End Function


Public Sub SendRequest(ByVal lngEventNo As Long, Optional ByVal strTag As String = "", _
    Optional data1, Optional data2, Optional data3, Optional strExPro As String = "")
'��������
End Sub



Public Sub Broadcast(ByVal lngEventNo As Long, Optional ByVal strTag As String = "", _
    Optional data1, Optional data2, Optional data3, Optional strExPro As String = "")
'�㲥�¼�
End Sub


Public Function PrintInfo(ByVal strErr As String, ByVal lngInfoType As Long, _
    Optional ByVal lngHwnd As Long = 0, Optional ByVal strUnitName As String = "", Optional ByVal strMethodName As String = "") As Long
'��ӡ������Ϣ
'strErr:������Ϣ
'lngErrType:�������ͣ�0-���棬1-��ͨ����2-���ݿ����
    
End Function


Public Function StudyInfo() As clsStudyInfo
'��ȡ�����Ϣ

End Function


Public Function PrintErr(objErr As ErrObject, ByVal lngInfoType As Long, _
    Optional ByVal lngHwnd As Long = 0, Optional ByVal strUnitName As String = "", Optional ByVal strMethodName As String = "") As Long

End Function
