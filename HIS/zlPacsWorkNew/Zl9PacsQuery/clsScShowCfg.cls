VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScShowCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Name As String
Public Icon As String   '��ͼ��
Public HiddenCol As Boolean '������
Public HiddenTitle As Boolean   '�����б���
Public HiddenData As Boolean    '������������ʾ
Public IsTotal As Boolean       '�Ƿ�ͳ����
Public UseListLocate As Boolean '�Ƿ���б���λ
Public SortContrastCol As String    '���������
Public DataConvert As String    '����ת��
Public IsNumerSort As Boolean    '��������


Private mobjRowRelation() As clsScRowRelation



Property Get RowRelationCount()
    RowRelationCount = UBound(mobjRowRelation)
End Property


Property Get RowRelation(ByVal lngIndex As Long) As clsScRowRelation
    Set RowRelation = mobjRowRelation(lngIndex)
End Property


Public Sub AddRowRelation(objRowRelation As clsScRowRelation)
'����RowRelation
    Dim lngUbound As Long
    
    lngUbound = UBound(mobjRowRelation) + 1
    ReDim Preserve mobjRowRelation(lngUbound)
    
    Set mobjRowRelation(lngUbound) = objRowRelation
End Sub

Private Sub Class_Initialize()
    ReDim mobjRowRelation(0)
End Sub

Private Sub Class_Terminate()
    ReDim mobjRowRelation(0)
    Set mobjRowRelation(0) = Nothing
End Sub