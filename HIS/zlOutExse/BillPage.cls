VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Details"

Option Explicit 'Ҫ���������

Public Key As String
Private mvarDetails As BillDetails '�ֲ�����
Private mvarNO As String '�ֲ�����
Private mvarӦ�ս�� As Double '�ֲ�����
Private mvarʵ�ս�� As Double '�ֲ�����
Private mvar���ս�� As Double '�ֲ�����
Private mvar��Ԥ���� As Double '�ֲ�����
Private mvarӦ�ɽ�� As Double '�ֲ�����
Private mvar����� As Double '�ֲ�����
Private mvar�շѽ��� As String '�ֲ�����
Private mvar����ͳ�� As Double '�ֲ�����
Private mvarȫ�Ը� As Double '�ֲ�����
Private mvar���Ը� As Double '�ֲ�����
Private mvar��������ID As Long '�ֲ�����
Private mvar������ As String '�ֲ�����

Private mvar���ѿ����� As String '�ֲ�����
Private mvar���ѿ�ˢ���� As Double '�ֲ�����
Private mvar�շѵ��� As String

'��������ֵ�ľֲ�����
Private mvar�巨 As String '�ֲ�����
'��������ֵ�ľֲ�����
Private mvarҽ����� As Long '�ֲ�����
Private mvar����ID As Long
Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շѽ��� = 5
    mvar����ID = vData
End Property
Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�շѽ���
    ����ID = mvar����ID
End Property


Public Property Let ҽ�����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ����� = 5
    mvarҽ����� = vData
End Property
Public Property Get ҽ�����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ�����
    ҽ����� = mvarҽ�����
End Property
Public Property Let �巨(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�巨 = 5
    mvar�巨 = vData
End Property


Public Property Get �巨() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�巨
    �巨 = mvar�巨
End Property

Private Sub Class_Initialize()
    Set Me.Details = New BillDetails
End Sub

Private Sub Class_Terminate()
    Set Me.Details = Nothing
End Sub

Public Property Set Details(ByVal vData As BillDetails)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    Set mvarDetails = vData
End Property

Public Property Get Details() As BillDetails
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    Set Details = mvarDetails
End Property
Public Property Let �շѵ���(ByVal vData As String)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar�շѵ��� = vData
End Property

Public Property Get �շѵ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    �շѵ��� = mvar�շѵ���
End Property



Public Property Let NO(ByVal vData As String)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvarNO = vData
End Property

Public Property Get NO() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    NO = mvarNO
End Property

Public Property Let Ӧ�ս��(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvarӦ�ս�� = vData
End Property

Public Property Get Ӧ�ս��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    Ӧ�ս�� = mvarӦ�ս��
End Property

Public Property Let ʵ�ս��(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvarʵ�ս�� = vData
End Property

Public Property Get ʵ�ս��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ʵ�ս�� = mvarʵ�ս��
End Property

Public Property Let ���ս��(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar���ս�� = vData
End Property

Public Property Get ���ս��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ���ս�� = mvar���ս��
End Property

Public Property Let ��Ԥ����(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar��Ԥ���� = vData
End Property

Public Property Get ��Ԥ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ��Ԥ���� = mvar��Ԥ����
End Property

Public Property Let Ӧ�ɽ��(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvarӦ�ɽ�� = vData
End Property

Public Property Get Ӧ�ɽ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    Ӧ�ɽ�� = mvarӦ�ɽ��
End Property

Public Property Let �����(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar����� = vData
End Property

Public Property Get �����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ����� = mvar�����
End Property

Public Property Let �շѽ���(ByVal vData As String)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar�շѽ��� = vData
End Property

Public Property Get �շѽ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    �շѽ��� = mvar�շѽ���
End Property

Public Property Let ����ͳ��(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar����ͳ�� = vData
End Property

Public Property Get ����ͳ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ����ͳ�� = mvar����ͳ��
End Property

Public Property Let ȫ�Ը�(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvarȫ�Ը� = vData
End Property

Public Property Get ȫ�Ը�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ȫ�Ը� = mvarȫ�Ը�
End Property

Public Property Let ���Ը�(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar���Ը� = vData
End Property

Public Property Get ���Ը�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ���Ը� = mvar���Ը�
End Property

Public Property Get ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.������
    ������ = mvar������
End Property

Public Property Let ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property

Public Property Let ��������ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��������ID = 5
    mvar��������ID = vData
End Property

Public Property Get ��������ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��������ID
    ��������ID = mvar��������ID
End Property

Public Property Let ���ѿ�����(ByVal vData As String)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar���ѿ����� = vData
End Property

Public Property Get ���ѿ�����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ���ѿ����� = mvar���ѿ�����
End Property

Public Property Let ���ѿ�ˢ����(ByVal vData As Double)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.Details = vData
    mvar���ѿ�ˢ���� = vData
End Property

Public Property Get ���ѿ�ˢ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Details
    ���ѿ�ˢ���� = mvar���ѿ�ˢ����
End Property