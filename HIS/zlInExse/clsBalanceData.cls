VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'��������ֵ�ľֲ�����
Private mvarobjBalanceItems As clsBalanceItems '�ֲ�����
Private mvar�����ܶ� As Double '�ֲ�����
Public Property Let �����ܶ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�����ܶ� = 5
    mvar�����ܶ� = vData
End Property

Public Property Get �����ܶ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����ܶ�
    �����ܶ� = mvar�����ܶ�
End Property
Public Property Set objBalanceItems(ByVal vData As clsBalanceItems)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objBalanceItems = Form1
    Set mvarobjBalanceItems = vData
End Property


Public Property Get objBalanceItems() As clsBalanceItems
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objBalanceItems
    Set objBalanceItems = mvarobjBalanceItems
End Property


