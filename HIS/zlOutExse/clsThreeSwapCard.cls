VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThreeSwapCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mvar�����ID As Long
Private mvar��ҽ�� As Boolean
Private mvar����㷽ʽ As Boolean
Private mvar�������� As Boolean
Private mvarȱʡ���� As Boolean
Private mvarȱʡ���ַ�ʽ As String

Public Property Let �����ID(ByVal vData As Long)
    mvar�����ID = vData
End Property

Public Property Get �����ID() As Long
    �����ID = mvar�����ID
End Property

    
Public Property Let ��ҽ��(ByVal vData As Boolean)
    mvar��ҽ�� = vData
End Property

Public Property Get ��ҽ��() As Boolean
    ��ҽ�� = mvar��ҽ��
End Property

    
Public Property Let ����㷽ʽ(ByVal vData As Boolean)
    mvar����㷽ʽ = vData
End Property

Public Property Get ����㷽ʽ() As Boolean
    ����㷽ʽ = mvar����㷽ʽ
End Property

    
Public Property Let ��������(ByVal vData As Boolean)
    mvar�������� = vData
End Property

Public Property Get ��������() As Boolean
    �������� = mvar��������
End Property

    
Public Property Let ȱʡ����(ByVal vData As Boolean)
    mvarȱʡ���� = vData
End Property

Public Property Get ȱʡ����() As Boolean
    ȱʡ���� = mvarȱʡ����
End Property

    
Public Property Let ȱʡ���ַ�ʽ(ByVal vData As String)
    mvarȱʡ���ַ�ʽ = vData
End Property

Public Property Get ȱʡ���ַ�ʽ() As String
    ȱʡ���ַ�ʽ = mvarȱʡ���ַ�ʽ
End Property