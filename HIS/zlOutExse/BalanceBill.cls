VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BalanceBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'��������ֵ�ľֲ�����
Private mvarȫ�Ը� As Currency '�ֲ�����
Private mvar���Ը� As Currency '�ֲ�����
Private mvarԤ���� As BalanceMoneys '�ֲ�����
Private mvarNO As String '�ֲ�����
Private mvar���۵� As Collection '�ֲ�����
Private mvar�շѵ� As Collection '�ֲ�����
Private mvarʵ�պϼ� As Currency '�ֲ�����
Private mvar����ͳ�� As Currency '�ֲ�����
Public Property Let ����ͳ��(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ͳ�� = 5
    mvar����ͳ�� = vData
End Property


Public Property Get ����ͳ��() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ͳ��
    ����ͳ�� = mvar����ͳ��
End Property



Public Property Let ʵ�պϼ�(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ�պϼ� = 5
    mvarʵ�պϼ� = vData
End Property


Public Property Get ʵ�պϼ�() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʵ�պϼ�
    ʵ�պϼ� = mvarʵ�պϼ�
End Property


Public Property Let NO(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.NO = 5
    mvarNO = vData
End Property


Public Property Get NO() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.NO
    NO = mvarNO
End Property



Public Property Set Ԥ����(ByVal vData As BalanceMoneys)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Ԥ���� = Form1
    Set mvarԤ���� = vData
End Property


Public Property Get Ԥ����() As BalanceMoneys
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ����
    Set Ԥ���� = mvarԤ����
End Property



Public Property Let ���Ը�(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���Ը� = 5
    mvar���Ը� = vData
End Property


Public Property Get ���Ը�() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���Ը�
    ���Ը� = mvar���Ը�
End Property



Public Property Let ȫ�Ը�(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ȫ�Ը� = 5
    mvarȫ�Ը� = vData
End Property


Public Property Get ȫ�Ը�() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ȫ�Ը�
    ȫ�Ը� = mvarȫ�Ը�
End Property


Private Sub Class_Initialize()
    Set Me.Ԥ���� = New BalanceMoneys
End Sub

Private Sub Class_Terminate()
    Set Me.Ԥ���� = Nothing
End Sub