VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTColProterty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Key As String

Private mvar�������� As String  '�ֲ�����
Private mvar�����ֶ� As String  '�ֲ�����
Private mvar������ϵ As String '�ֲ�����
Private mvar����ֵ As String  '�ֲ�����
Private mvar������ɫ As Long '�ֲ�����
Private mvar������ɫ As Long '�ֲ�����
Private mvar�Ƿ�Ӵ� As Boolean '�ֲ�����
Private mvar�Ƿ�����Ӧ�� As Boolean '�ֲ�����
Private mint���� As Integer                 '0-�Զ����ֶ�����ȷ�ϣ���1-����2-���У�3-����

Public Property Let ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    �������� = mvar��������
End Property

Public Property Let �����ֶ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ֵ��Դ = 5
    mvar�����ֶ� = vData
End Property

Public Property Get �����ֶ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ֵ��Դ
    �����ֶ� = mvar�����ֶ�
End Property

Public Property Let ������ϵ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar������ϵ = vData
End Property

Public Property Get ������ϵ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    ������ϵ = mvar������ϵ
End Property

Public Property Let ����ֵ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar����ֵ = vData
End Property

Public Property Get ����ֵ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    ����ֵ = mvar����ֵ
End Property


Public Property Let ������ɫ(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar������ɫ = vData
End Property

Public Property Get ������ɫ() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    ������ɫ = mvar������ɫ
End Property

Public Property Let ������ɫ(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar������ɫ = vData
End Property

Public Property Get ������ɫ() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    ������ɫ = mvar������ɫ
End Property

Public Property Let �Ƿ�Ӵ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�Ƿ�Ӵ� = vData
End Property

Public Property Get �Ƿ�Ӵ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �Ƿ�Ӵ� = mvar�Ƿ�Ӵ�
End Property

Public Property Let �Ƿ�����Ӧ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�Ƿ�����Ӧ�� = vData
End Property

Public Property Get �Ƿ�����Ӧ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �Ƿ�����Ӧ�� = mvar�Ƿ�����Ӧ��
End Property

Public Property Let ����(ByVal vData As Integer)
    mint���� = vData
End Property

Public Property Get ����() As Integer
    ���� = mint����
End Property