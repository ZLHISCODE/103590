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
Option Explicit 'Ҫ���������

Public Key As String

Private mvar������ĿID As Long '�ֲ�����
Private mvar������Ŀ As String '�ֲ�����
Private mvar�վݷ�Ŀ As String '�ֲ�����
Private mvar��׼���� As Double '�ֲ�����
Private mvar�ּ� As Double '�ֲ�����
Private mvarԭ�� As Double '�ֲ�����
Private mvarӦ�ս�� As Double '�ֲ�����
Private mvarʵ�ս�� As Double '�ֲ�����
Private mvarͳ���� As Double '�ֲ�����
Private mvarTag As Variant '�ֲ�����
Public Property Let Tag(ByVal vData As Variant)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Tag = 5
    mvarTag = vData
End Property
Public Property Set Tag(ByVal vData As Variant)
    '������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
    'Syntax: Set x.Tag = Form1
    Set mvarTag = vData
End Property
Public Property Get Tag() As Variant
    '��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
    'Syntax: Debug.Print X.Tag
    If IsObject(mvarTag) Then
        Set Tag = mvarTag
    Else
        Tag = mvarTag
    End If
End Property

Public Property Let ͳ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ͳ���� = 5
    mvarͳ���� = vData
End Property

Public Property Get ͳ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.ͳ����
    ͳ���� = mvarͳ����
End Property

Public Property Let ʵ�ս��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ�ս�� = 5
    mvarʵ�ս�� = vData
End Property

Public Property Get ʵ�ս��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.ʵ�ս��
    ʵ�ս�� = mvarʵ�ս��
End Property

Public Property Let Ӧ�ս��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ӧ�ս�� = 5
    mvarӦ�ս�� = vData
End Property

Public Property Get Ӧ�ս��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.Ӧ�ս��
    Ӧ�ս�� = mvarӦ�ս��
End Property

Public Property Let ��׼����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��׼���� = 5
    mvar��׼���� = vData
End Property

Public Property Get ��׼����() As Double
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

Public Property Let ԭ��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��׼���� = 5
    mvarԭ�� = vData
End Property

Public Property Get ԭ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��׼����
    ԭ�� = mvarԭ��
End Property

Public Property Let �ּ�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��׼���� = 5
    mvar�ּ� = vData
End Property

Public Property Get �ּ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��׼����
    �ּ� = mvar�ּ�
End Property
