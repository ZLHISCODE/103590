VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BalanceMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'��������ֵ�ľֲ�����
Private mvar���㷽ʽ As String '�ֲ�����
Private mvar�����޸� As Boolean '�ֲ�����
Private mvarԭʼ��� As Double '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar��Ч��� As Double '�ֲ�����
Public Property Let ��Ч���(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ч��� = 5
    mvar��Ч��� = vData
End Property


Public Property Get ��Ч���() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ч���
    ��Ч��� = mvar��Ч���
End Property



Public Property Let ԭʼ���(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ԭʼ��� = 5
    mvarԭʼ��� = vData
End Property


Public Property Get ԭʼ���() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ԭʼ���
    ԭʼ��� = mvarԭʼ���
End Property



Public Property Let �����޸�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�����޸� = 5
    mvar�����޸� = vData
End Property


Public Property Get �����޸�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����޸�
    �����޸� = mvar�����޸�
End Property



Public Property Let ���㷽ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���㷽ʽ = 5
    mvar���㷽ʽ = vData
End Property


Public Property Get ���㷽ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���㷽ʽ
    ���㷽ʽ = mvar���㷽ʽ
End Property


