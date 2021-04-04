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
    infHint = -2         '提示
    infWaring = -3   '警告
    infNormalErr = -4    '普通错误
    infDataErr = -5      '数据库错误
End Enum

Public Function Owner() As Object
'实现接口的对象拥有者
End Function


Public Function hwnd() As Long
'实现接口的对象句柄
End Function


'Public Function MainPrivs() As String
''获取权限串
'End Function


Public Sub SendRequest(ByVal lngEventNo As Long, Optional ByVal strTag As String = "", _
    Optional data1, Optional data2, Optional data3, Optional strExPro As String = "")
'发送请求
End Sub



Public Sub Broadcast(ByVal lngEventNo As Long, Optional ByVal strTag As String = "", _
    Optional data1, Optional data2, Optional data3, Optional strExPro As String = "")
'广播事件
End Sub


Public Function PrintInfo(ByVal strErr As String, ByVal lngInfoType As Long, _
    Optional ByVal lngHwnd As Long = 0, Optional ByVal strUnitName As String = "", Optional ByVal strMethodName As String = "") As Long
'打印错误消息
'strErr:错误消息
'lngErrType:错误类型，0-警告，1-普通错误，2-数据库错误
    
End Function


Public Function StudyInfo() As clsStudyInfo
'获取检查信息

End Function


Public Function PrintErr(objErr As ErrObject, ByVal lngInfoType As Long, _
    Optional ByVal lngHwnd As Long = 0, Optional ByVal strUnitName As String = "", Optional ByVal strMethodName As String = "") As Long

End Function

