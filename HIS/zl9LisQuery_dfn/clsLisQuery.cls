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
    '��ǰ�������
End Property

Public Property Let Index(ByVal lngIndex As Long)
    '��ǰ�������
End Property

Public Property Get Name() As String
    '����(ֻ��)
End Property

Public Property Get Description() As String
    '����(ֻ��)
End Property

Public Sub InitQuery(QueryHost As clsLisQueryHost)
    '��ʼ�� Host��
End Sub

Public Sub DoAction(ByVal ConShowMode As QueryShowMode)
    'ִ�� �������
    'ʵ�ִ˹���ʱ������֪ͨ HOS �����
    'clsHost.RaiseFinished Me
End Sub

Public Property Get Result() As String
    '
End Property

Public Property Get Value() As Variant
    '
End Property
