VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit 'Ҫ���������

Public Key As String

'��������ֵ�ľֲ�����
Private mvarID As Long '�ֲ�����
Private mvarҩ��ID As Long '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��Ʒ�� As String '�ֲ�����
Private mvar������� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��� As String '�ֲ�����
Private mvar���㵥λ As String '�ֲ�����
Private mvar˵�� As String '�ֲ�����
Private mvar���ηѱ� As Boolean '�ֲ�����
Private mvar��� As Boolean '�ֲ�����
Private mvar���� As Boolean '�ֲ�����
Private mvar�Ӱ�Ӽ� As Boolean '�ֲ�����
Private mvar����ժҪ As Boolean '�ֲ�����
Private mvarִ�п��� As Byte '�ֲ�����
Private mvar���д��� As Byte   '�ֲ�����
Private mvar�������� As Double   '�ֲ�����
Private mvar��� As String '�ֲ�����
Private mvar��� As Double   '�ֲ�����
Private mvar���� As String
Private mvar�۸� As Double   '�ֲ�����,����������Ŀ���ƿ���ֵ
Private mvar�������� As Double
Private mvar¼������ As Double
Private mvar����ְ�� As String
Private mvar������� As Byte '�ֲ�����
Private mvarסԺ��װ As Double '�ֲ�����
Private mvarסԺ��λ As String '�ֲ�����
Private mvar�������� As Boolean '�ֲ�����
Private mvarҪ������ As Boolean '�ֲ�����
Private mvar��ҩ��̬ As Integer       '�ֲ�����
Private mvar�������� As String
Private mvar������λ As String
Private mvar����ϵ�� As Double
Private mvar���� As Long         '�ֲ�����



Public Property Let Ҫ������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ҫ������ = 5
    mvarҪ������ = vData
End Property


Public Property Get Ҫ������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ҫ������
    Ҫ������ = mvarҪ������
End Property


Public Property Let ҩ��ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҩ��ID = 5
    mvarҩ��ID = vData
End Property

Public Property Get ҩ��ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.ҩ��ID
    ҩ��ID = mvarҩ��ID
End Property

Public Property Let ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ����ְ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar����ְ�� = vData
End Property

Public Property Get ����ְ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ����ְ�� = mvar����ְ��
End Property

Public Property Let ��������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    �������� = mvar��������
End Property


Public Property Let ¼������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar¼������ = vData
End Property

Public Property Get ¼������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ¼������ = mvar¼������
End Property

Public Property Let �۸�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar�۸� = vData
End Property

Public Property Get �۸�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    �۸� = mvar�۸�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ��������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    �������� = mvar��������
End Property

Public Property Let ���д���(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar���д��� = vData
End Property

Public Property Get ���д���() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ���д��� = mvar���д���
End Property

Public Property Let ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��� = mvar���
End Property

Public Property Let ִ�п���(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ִ�п��� = 5
    mvarִ�п��� = vData
End Property

Public Property Get ִ�п���() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.ִ�п���
    ִ�п��� = mvarִ�п���
End Property

Public Property Let �Ӱ�Ӽ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ӱ�Ӽ� = 5
    mvar�Ӱ�Ӽ� = vData
End Property

Public Property Get �Ӱ�Ӽ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�Ӱ�Ӽ�
    �Ӱ�Ӽ� = mvar�Ӱ�Ӽ�
End Property

Public Property Let ���(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��� = mvar���
End Property

Public Property Let ���ηѱ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ηѱ� = 5
    mvar���ηѱ� = vData
End Property

Public Property Get ���ηѱ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���ηѱ�
    ���ηѱ� = mvar���ηѱ�
End Property

Public Property Let ˵��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.˵�� = 5
    mvar˵�� = vData
End Property

Public Property Get ˵��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.˵��
    ˵�� = mvar˵��
End Property

Public Property Let ���㵥λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���㵥λ = 5
    mvar���㵥λ = vData
End Property

Public Property Get ���㵥λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���㵥λ
    ���㵥λ = mvar���㵥λ
End Property

Public Property Let ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��� = vData
End Property

Public Property Get ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��� = mvar���
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ���� = mvar����
End Property
Public Property Let ��Ʒ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar��Ʒ�� = vData
End Property

Public Property Get ��Ʒ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ��Ʒ�� = mvar��Ʒ��
End Property

Public Property Let ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Property Get ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.ID
    ID = mvarID
End Property

Public Property Let �������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar������� = vData
End Property

Public Property Get �������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����
    ������� = mvar�������
End Property

Public Property Let ���(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ�ϸĿID = 5
    mvar��� = vData
End Property

Public Property Get ���() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�շ�ϸĿID
    ��� = mvar���
End Property

Public Property Let ����ժҪ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ժҪ = 5
    mvar����ժҪ = vData
End Property

Public Property Get ����ժҪ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.����ժҪ
    ����ժҪ = mvar����ժҪ
End Property

Public Property Let �������(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������� = 5
    mvar������� = vData
End Property

Public Property Get �������() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.�������
    ������� = mvar�������
End Property

Public Property Let סԺ��װ(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.סԺ��װ = 5
    mvarסԺ��װ = vData
End Property

Public Property Get סԺ��װ() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.סԺ��װ
    סԺ��װ = mvarסԺ��װ
End Property

Public Property Let סԺ��λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.סԺ��λ = 5
    mvarסԺ��λ = vData
End Property

Public Property Get סԺ��λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.סԺ��λ
    סԺ��λ = mvarסԺ��λ
End Property

Public Property Let ��������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.��������
    �������� = mvar��������
End Property


Public Property Let ��ҩ��̬(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar��ҩ��̬ = vData
End Property

Public Property Get ��ҩ��̬() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ��ҩ��̬ = mvar��ҩ��̬
End Property
Public Property Let ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    �������� = mvar��������
End Property
Public Property Let ������λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar������λ = vData
End Property

Public Property Get ������λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ������λ = mvar������λ
End Property

Public Property Let ����ϵ��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��� = 5
    mvar����ϵ�� = vData
End Property

Public Property Get ����ϵ��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: 'debug.print X.���
    ����ϵ�� = mvar����ϵ��
End Property
