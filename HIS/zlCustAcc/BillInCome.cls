VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillInCome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarʵ��Ʊ�� As String
Private mvar������ĿID As Long '�ֲ�����
Private mvar������Ŀ As String '�ֲ�����
Private mvar�վݷ�Ŀ As String '�ֲ�����
Private mvar��׼���� As Currency '�ֲ�����
Private mvar�ּ� As Currency '�ֲ�����
Private mvarԭ�� As Currency '�ֲ�����
Private mvarӦ�ս�� As Currency '�ֲ�����
Private mvarʵ�ս�� As Currency '�ֲ�����
Private mvarͳ���� As Currency '�ֲ�����

Public Property Let ͳ����(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ͳ���� = 5
    mvarͳ���� = vData
End Property

Public Property Get ͳ����() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.ͳ����
    ͳ���� = mvarͳ����
End Property

Public Property Let ʵ��Ʊ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ڼ� = 5
    mvarʵ��Ʊ�� = vData
End Property

Public Property Get ʵ��Ʊ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�ڼ�
    ʵ��Ʊ�� = mvarʵ��Ʊ��
End Property

Public Property Let ʵ�ս��(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ�ս�� = 5
    mvarʵ�ս�� = vData
End Property

Public Property Get ʵ�ս��() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.ʵ�ս��
    ʵ�ս�� = mvarʵ�ս��
End Property

Public Property Let Ӧ�ս��(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ӧ�ս�� = 5
    mvarӦ�ս�� = vData
End Property

Public Property Get Ӧ�ս��() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Ӧ�ս��
    Ӧ�ս�� = mvarӦ�ս��
End Property

Public Property Let ��׼����(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��׼���� = 5
    mvar��׼���� = vData
End Property

Public Property Get ��׼����() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��׼����
    ��׼���� = mvar��׼����
End Property

Public Property Let �վݷ�Ŀ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�վݷ�Ŀ = 5
    mvar�վݷ�Ŀ = vData
End Property

Public Property Get �վݷ�Ŀ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�վݷ�Ŀ
    �վݷ�Ŀ = mvar�վݷ�Ŀ
End Property

Public Property Let ������Ŀ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�վݷ�Ŀ = 5
    mvar������Ŀ = vData
End Property

Public Property Get ������Ŀ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�վݷ�Ŀ
    ������Ŀ = mvar������Ŀ
End Property

Public Property Let ������ĿID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ĿID = 5
    mvar������ĿID = vData
End Property

Public Property Get ������ĿID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.������ĿID
    ������ĿID = mvar������ĿID
End Property

Public Property Let ԭ��(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��׼���� = 5
    mvarԭ�� = vData
End Property

Public Property Get ԭ��() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��׼����
    ԭ�� = mvarԭ��
End Property

Public Property Let �ּ�(ByVal vData As Currency)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��׼���� = 5
    mvar�ּ� = vData
End Property

Public Property Get �ּ�() As Currency
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��׼����
    �ּ� = mvar�ּ�
End Property
