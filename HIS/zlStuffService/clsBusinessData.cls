VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusinessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'��������ֵ�ľֲ�����

Private mvar�������� As gDataType '�ֲ�����
Private mvar�ڵ����� As String '�ֲ�����
Private mvar�������� As String '�ֲ�����
Private mvar�������� As String '�ֲ�����
Private mvar����˵�� As String '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar�Ƿ��ѡ As Boolean '�ֲ�����
Private mvar�Ƿ��ֵ As Boolean '�ֲ�����
Private mvar���� As Integer     '�ֲ�����
Private mvarС�� As Integer     '�ֲ�����
Private mvarTag As String
Public Property Let �ڵ�����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ǽ�ʱ�� = 5
    mvar�ڵ����� = vData
End Property

Public Property Get �ڵ�����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ǽ�ʱ��
    �ڵ����� = mvar�ڵ�����
End Property

Public Property Let Tag(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ǽ�ʱ�� = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ǽ�ʱ��
    Tag = mvarTag
End Property

Public Property Let С��(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.С�� = 5
    mvarС�� = vData
End Property


Public Property Get С��() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.С��
    С�� = mvarС��
End Property

Public Property Let ����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property



Public Property Let �Ƿ��ѡ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��ѡ = 5
    mvar�Ƿ��ѡ = vData
End Property


Public Property Get �Ƿ��ѡ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��ѡ
    �Ƿ��ѡ = mvar�Ƿ��ѡ
End Property

Public Property Let �Ƿ��ֵ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��ֵ = 5
    mvar�Ƿ��ֵ = vData
End Property


Public Property Get �Ƿ��ֵ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��ֵ
    �Ƿ��ֵ = mvar�Ƿ��ֵ
End Property

Public Property Let ����˵��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����˵�� = 5
    mvar����˵�� = vData
End Property


Public Property Get ����˵��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����˵��
    ����˵�� = mvar����˵��
End Property



Public Property Let ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property


Public Property Get ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property



Public Property Let ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property


Public Property Get ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property



Public Property Let ��������(ByVal vData As gDataType)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property


Public Property Get ��������() As gDataType
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property


