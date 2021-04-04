VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum QueryShowMode
    Query_ShowModal = vbModal
    Query_ShowModeless = vbModeless
End Enum

Public Property Get Index() As Long
    '当前类的索引
End Property

Public Property Let Index(ByVal lngIndex As Long)
    '当前类的索引
End Property

Public Property Get Name() As String
    '名称(只读)
End Property

Public Property Get Description() As String
    '描述(只读)
End Property

Public Sub InitQuery(QueryHost As clsLisQueryHost)
    '初始化 Host类
End Sub

Public Sub DoAction(ByVal ConShowMode As QueryShowMode)
    '执行 插件功能
    '实现此过程时，可以通知 HOS 已完成
    'clsHost.RaiseFinished Me
End Sub

Public Property Get Result() As String
    '
End Property

Public Property Get Value() As Variant
    '
End Property

