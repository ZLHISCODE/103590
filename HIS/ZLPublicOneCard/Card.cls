VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum gCard_Type
    EM_CardType_Cash = 0    '�ֽ�
    EM_CardType_Consume = 1 '���ѿ�
    EM_CardType_Square = 2  '���㿨
    EM_CardType_HospitalAccount = 3 'Ԥ���
    EM_CardType_Insure = 4  '����
    EM_CardType_other = 5  '����
End Enum
Public Enum gKeyBoard_Type  '���̿���
    EM_KeyBoard_NotKey = 0
    EM_KeyBoard_NumberKey = 1
    EM_KeyBoard_CharKey = 2
End Enum
'��������ֵ�ľֲ�����
Private mvar�ӿ���� As Long     '�ֲ�����
Private mvar�ӿڱ��� As String     '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As Boolean '�ֲ�����
Private mvar�ӿڳ����� As String '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar���� As Integer
Private mvar�Զ���ȡ As Integer
'��������ֵ�ľֲ�����
Private mvarϵͳ As Boolean '�ֲ�����
Private mvar���㷽ʽ As String '�ֲ�����
Private mvarǰ׺�ı� As String '�ֲ�����
Private mvar���ų��� As Long '�ֲ�����
Private mvar���ƿ� As Boolean
Private mvar���ѿ� As Boolean '
Private mvar���� As String
Private mvar�Ƿ��ϸ���� As Boolean
Private mvar�Ƿ�����ʻ� As Boolean
Private mvar�Ƿ�ȫ�� As Boolean
Private mvar��ע As String
Private mvar�ض���Ŀ As String
Private mvar�������Ĺ��� As String
Private mvar�����ظ�ʹ�� As Boolean
Private mvar�Զ���ȡ��� As Integer  '�Զ���ȡ���
Private mvar�Ƿ����� As Boolean
Private mvar���볤�� As Integer
Private mvar���볤������ As Integer
Private mvar������� As Integer
Private mvarͼ���ʶ As String
Private mvar֧������ As Boolean
Private mvar֧��ͼ���ʶ As String
Private mvar������������ As Long
Private mvar�Ƿ�ȱʡ���� As Boolean
Private mvarȱʡ��־ As Boolean
Private mvar���ܼ� As String
Private mvar��� As String
Private mvarģ�������� As Boolean
Private mvar�Ƿ�ģ������ As Boolean
Private mvar�Ƿ�ת�ʼ����� As Boolean
Private mvar�Ƿ�ˢ��  As Boolean
Private mvar�Ƿ�ɨ��  As Boolean
Private mvar�Ƿ�Ӵ�ʽ����  As Boolean
Private mvar�Ƿ�ǽӴ�ʽ���� As Boolean
Private mvar�Ƿ�ֿ����� As Boolean
'��������ֵ�ľֲ�����
Private mvar�Ƿ��ƿ� As Boolean '�ֲ�����
Private mvar�Ƿ񷢿� As Boolean '�ֲ�����
Private mvar�Ƿ�д�� As Boolean '�ֲ�����
Private mvar�������� As Integer   '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar�Ƿ��˿��鿨 As Boolean '�ֲ�����
Private mvar�Ƿ�֤�� As Boolean
Private mvar�豸�Ƿ����ûس� As Boolean
Private mvar�Ƿ�ȱʡ���� As Boolean
Private mvar�Ƿ�������� As Boolean
Private mvarCardObject As Object '�ֲ�����
Private mvarInitCompents As Boolean '�ֲ�����
Private mvar�������� As gKeyBoard_Type
Private mvar������ As gCard_Type  '1-���ѿ�;2-���㿨;3-�ֽ�;4-�ʻ�;5-����
Private mvar���� As Integer
Private mvar�������� As String
Private mvar�Ƿ��͵��ýӿ� As Boolean
Private mvar����ʶ����� As Integer
'��������ֵ�ľֲ�����
Private mvar�Ƿ�֧��ɨ�븶 As Boolean '�ֲ�����
Public Property Let �Ƿ�֧��ɨ�븶(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�֧��ɨ�븶 = 5
    mvar�Ƿ�֧��ɨ�븶 = vData
End Property


Public Property Get �Ƿ�֧��ɨ�븶() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�֧��ɨ�븶
    �Ƿ�֧��ɨ�븶 = mvar�Ƿ�֧��ɨ�븶
End Property



Public Property Let ����ʶ�����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar����ʶ����� = vData
End Property

Public Property Get ����ʶ�����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ʶ�����
    ����ʶ����� = mvar����ʶ�����
End Property



Public Property Let �Ƿ��͵��ýӿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�Ƿ��͵��ýӿ� = vData
End Property

Public Property Get �Ƿ��͵��ýӿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��͵��ýӿ�
    �Ƿ��͵��ýӿ� = mvar�Ƿ��͵��ýӿ�
End Property

Public Property Let ������(ByVal vData As gCard_Type)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������  = 5
    mvar������ = vData
End Property
Public Property Get ������() As gCard_Type
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������
    ������ = mvar������
End Property

Public Property Let ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��������  = 5
    mvar�������� = vData
End Property
Public Property Get ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property
Public Property Let ����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Let ��������(ByVal vData As gKeyBoard_Type)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��������  = 5
    mvar�������� = vData
End Property
Public Property Get ��������() As gKeyBoard_Type
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property


Public Property Let InitCompents(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ѿ� = 5
    mvarInitCompents = vData
End Property
Public Property Get InitCompents() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ѿ�
    InitCompents = mvarInitCompents
End Property

Public Property Set CardObject(ByVal vData As Object)
    '������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
    'Syntax: Set x.CardObject = Form1
    Set mvarCardObject = vData
End Property


Public Property Get CardObject() As Object
    '��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
    'Syntax: Debug.Print X.CardObject
    Set CardObject = mvarCardObject
End Property

Public Property Let �Ƿ��������(ByVal vData As Boolean)
    mvar�Ƿ�������� = vData
End Property
Public Property Get �Ƿ��������() As Boolean
    �Ƿ�������� = mvar�Ƿ��������
End Property

Public Property Let �Ƿ�ȱʡ����(ByVal vData As Boolean)
    mvar�Ƿ�ȱʡ���� = vData
End Property
Public Property Get �Ƿ�ȱʡ����() As Boolean
    �Ƿ�ȱʡ���� = mvar�Ƿ�ȱʡ����
End Property

Public Property Let �豸�Ƿ����ûس�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�д�� = 5
    mvar�豸�Ƿ����ûس� = vData
End Property

Public Property Get �豸�Ƿ����ûس�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�д��
    �豸�Ƿ����ûس� = mvar�豸�Ƿ����ûس�
End Property

Public Property Let �Ƿ�֤��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�д�� = 5
    mvar�Ƿ�֤�� = vData
End Property

Public Property Get �Ƿ�֤��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�д��
    �Ƿ�֤�� = mvar�Ƿ�֤��
End Property

Public Property Let �Ƿ��˿��鿨(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��˿��鿨 = 5
    mvar�Ƿ��˿��鿨 = vData
End Property


Public Property Get �Ƿ��˿��鿨() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��˿��鿨
    �Ƿ��˿��鿨 = mvar�Ƿ��˿��鿨
End Property




Public Property Let �Ƿ�ֿ�����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�д�� = 5
    mvar�Ƿ�ֿ����� = vData
End Property

Public Property Get �Ƿ�ֿ�����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�д��
    �Ƿ�ֿ����� = mvar�Ƿ�ֿ�����
End Property

Public Property Let �Ƿ�ˢ��(ByVal vData As Boolean)
    mvar�Ƿ�ˢ�� = vData
End Property
Public Property Get �Ƿ�ˢ��() As Boolean
    �Ƿ�ˢ�� = mvar�Ƿ�ˢ��
End Property

Public Property Let �Ƿ�ɨ��(ByVal vData As Boolean)
    mvar�Ƿ�ɨ�� = vData
End Property
Public Property Get �Ƿ�ɨ��() As Boolean
    �Ƿ�ɨ�� = mvar�Ƿ�ɨ��
End Property

Public Property Let �Ƿ�Ӵ�ʽ����(ByVal vData As Boolean)
    mvar�Ƿ�Ӵ�ʽ���� = vData
End Property
Public Property Get �Ƿ�Ӵ�ʽ����() As Boolean
    �Ƿ�Ӵ�ʽ���� = mvar�Ƿ�Ӵ�ʽ����
End Property

Public Property Let �Ƿ�ǽӴ�ʽ����(ByVal vData As Boolean)
    mvar�Ƿ�ǽӴ�ʽ���� = vData
End Property
Public Property Get �Ƿ�ǽӴ�ʽ����() As Boolean
    �Ƿ�ǽӴ�ʽ���� = mvar�Ƿ�ǽӴ�ʽ����
End Property

Public Property Let �Ƿ�ת�ʼ�����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�д�� = 5
    mvar�Ƿ�ת�ʼ����� = vData
End Property

Public Property Get �Ƿ�ת�ʼ�����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�д��
    �Ƿ�ת�ʼ����� = mvar�Ƿ�ת�ʼ�����
End Property

Public Property Let ��������(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property

Public Property Let �Ƿ�д��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�д�� = 5
    mvar�Ƿ�д�� = vData
End Property

Public Property Get �Ƿ�д��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�д��
    �Ƿ�д�� = mvar�Ƿ�д��
End Property

Public Property Let �Ƿ񷢿�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ񷢿� = 5
    mvar�Ƿ񷢿� = vData
End Property


Public Property Get �Ƿ񷢿�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ񷢿�
    �Ƿ񷢿� = mvar�Ƿ񷢿�
End Property



Public Property Let �Ƿ��ƿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��ƿ� = 5
    mvar�Ƿ��ƿ� = vData
End Property


Public Property Get �Ƿ��ƿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��ƿ�
    �Ƿ��ƿ� = mvar�Ƿ��ƿ�
End Property

Public Property Let �Ƿ�ģ������(ByVal vData As Boolean)
    mvar�Ƿ�ģ������ = vData
End Property
Public Property Get �Ƿ�ģ������() As Boolean
    �Ƿ�ģ������ = mvar�Ƿ�ģ������
End Property


Public Property Let ģ��������(ByVal vData As Boolean)
    mvarģ�������� = vData
End Property
Public Property Get ģ��������() As Boolean
    ģ�������� = mvarģ��������
End Property

Public Property Let ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar��� = vData
End Property
Public Property Get ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ��� = mvar���
End Property
Public Property Let ���ܼ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar���ܼ� = vData
End Property
Public Property Get ���ܼ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ���ܼ� = mvar���ܼ�
End Property


Public Property Let ȱʡ��־(ByVal vData As Boolean)
    mvarȱʡ��־ = vData
End Property
Public Property Get ȱʡ��־() As Boolean
    ȱʡ��־ = mvarȱʡ��־
End Property

Public Property Let �Ƿ�ȱʡ����(ByVal vData As Boolean)
    mvar�Ƿ�ȱʡ���� = vData
End Property
Public Property Get �Ƿ�ȱʡ����() As Boolean
    �Ƿ�ȱʡ���� = mvar�Ƿ�ȱʡ����
End Property
Public Property Let ������������(ByVal vData As Long)
    mvar������������ = vData
End Property
Public Property Get ������������() As Long
    ������������ = mvar������������
End Property
Public Property Let ֧��ͼ���ʶ(ByVal vData As String)
    mvar֧��ͼ���ʶ = vData
End Property
Public Property Get ֧��ͼ���ʶ() As String
    ֧��ͼ���ʶ = mvar֧��ͼ���ʶ
End Property
Public Property Let ֧������(ByVal vData As Boolean)
    mvar֧������ = vData
End Property
Public Property Get ֧������() As Boolean
    ֧������ = mvar֧������
End Property

Public Property Let ���볤��(ByVal vData As Integer)
    mvar���볤�� = vData
End Property
Public Property Get ���볤��() As Integer
    ���볤�� = mvar���볤��
End Property

Public Property Let ���볤������(ByVal vData As Integer)
    mvar���볤������ = vData
End Property
Public Property Get ���볤������() As Integer
    ���볤������ = mvar���볤������
End Property
Public Property Let �������(ByVal vData As Integer)
    mvar������� = vData
End Property
Public Property Get �������() As Integer
    ������� = mvar�������
End Property

Public Property Let �Ƿ�����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar�Ƿ����� = vData
End Property
Public Property Get �Ƿ�����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    �Ƿ����� = mvar�Ƿ�����
End Property

Public Property Let �Զ���ȡ���(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar�Զ���ȡ��� = vData
End Property
Public Property Get �Զ���ȡ���() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    �Զ���ȡ��� = mvar�Զ���ȡ���
End Property


Public Property Let �����ظ�ʹ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar�����ظ�ʹ�� = vData
End Property
Public Property Get �����ظ�ʹ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    �����ظ�ʹ�� = mvar�����ظ�ʹ��
End Property

Public Property Let ���ѿ�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar���ѿ� = vData
End Property
Public Property Get ���ѿ�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ���ѿ� = mvar���ѿ�
End Property

Public Property Let �������Ĺ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar�������Ĺ��� = vData
End Property
Public Property Get �������Ĺ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    �������Ĺ��� = mvar�������Ĺ���
End Property

Public Property Let ���ų���(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ų��� = 5
    mvar���ų��� = vData
End Property
Public Property Get ���ų���() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ų���
    ���ų��� = mvar���ų���
End Property

Public Property Let ǰ׺�ı�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvarǰ׺�ı� = vData
End Property
Public Property Get ǰ׺�ı�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ǰ׺�ı� = mvarǰ׺�ı�
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
Public Property Let ϵͳ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ϵͳ = 5
    mvarϵͳ = vData
End Property
Public Property Get ϵͳ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ϵͳ
     ϵͳ = mvarϵͳ
End Property

Public Property Let �Ƿ��Զ���ȡ(ByVal vData As Integer)
    mvar�Զ���ȡ = vData
End Property

Public Property Get �Ƿ��Զ���ȡ() As Integer
    �Ƿ��Զ���ȡ = mvar�Զ���ȡ
End Property
Public Property Let �ɷ�����(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property
Public Property Get �ɷ�����() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    �ɷ����� = mvar����
End Property
Public Property Let �ӿڳ�����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ӿڳ����� = 5
    mvar�ӿڳ����� = vData
End Property


Public Property Get �ӿڳ�����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ӿڳ�����
    �ӿڳ����� = mvar�ӿڳ�����
End Property

Public Property Let ����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let �ӿڱ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar�ӿڱ��� = vData
End Property

Public Property Get �ӿڱ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    �ӿڱ��� = mvar�ӿڱ���
End Property
Public Property Let �ӿ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar�ӿ���� = vData
End Property

Public Property Get �ӿ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    �ӿ���� = mvar�ӿ����
End Property
Public Property Let ���ƿ�(ByVal vData As Boolean)
    mvar���ƿ� = vData
End Property

Public Property Get ���ƿ�() As Boolean
    ���ƿ� = mvar���ƿ�
End Property

  
Public Property Let ����(ByVal vData As String)
    mvar���� = vData
End Property
Public Property Get ����() As String
    ���� = mvar����
End Property
Public Property Let ��ע(ByVal vData As String)
    mvar��ע = vData
End Property
Public Property Get ��ע() As String
    ��ע = mvar��ע
End Property
Public Property Let �ض���Ŀ(ByVal vData As String)
    mvar�ض���Ŀ = vData
End Property
Public Property Get �ض���Ŀ() As String
    �ض���Ŀ = mvar�ض���Ŀ
End Property
    
Public Property Let �Ƿ��ϸ����(ByVal vData As Boolean)
    mvar�Ƿ��ϸ���� = vData
End Property
Public Property Get �Ƿ��ϸ����() As Boolean
    �Ƿ��ϸ���� = mvar�Ƿ��ϸ����
End Property
Public Property Let �Ƿ�����ʻ�(ByVal vData As Boolean)
    mvar�Ƿ�����ʻ� = vData
End Property
Public Property Get �Ƿ�����ʻ�() As Boolean
    �Ƿ�����ʻ� = mvar�Ƿ�����ʻ�
End Property
Public Property Let �Ƿ�ȫ��(ByVal vData As Boolean)
    mvar�Ƿ�ȫ�� = vData
End Property
Public Property Get �Ƿ�ȫ��() As Boolean
    �Ƿ�ȫ�� = mvar�Ƿ�ȫ��
End Property
Public Property Let ͼ���ʶ(ByVal vData As String)
    mvarͼ���ʶ = vData
End Property
Public Property Get ͼ���ʶ() As String
    ͼ���ʶ = mvarͼ���ʶ
End Property
Public Function zlCardNOEncrypt(ByVal strCardNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ܵĿ���
    '���:strCardNO-δ���ܿ���
    '����:��**�Ŀ���,�������,���ؿ�:21432**13234
    '����:���˺�
    '����:2012-04-26 11:29:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPass As Variant, strRule As String
    Dim strCardPassText As String, I As Long, J As Long
    strRule = �������Ĺ���
    If ���ѿ� Then
        If Val(strRule) = 1 Then zlCardNOEncrypt = String(Len(strCardNo), "*"): Exit Function
        zlCardNOEncrypt = strCardNo
    End If
    
    strCardPassText = strRule
    If strCardPassText = "" Then zlCardNOEncrypt = strCardNo: Exit Function

    varPass = Split(strCardPassText & "-", "-")
    If Val(varPass(0)) = 0 Or Val(varPass(1)) = 0 Then
        '���λ��ʾ*
        I = IIf(Val(varPass(0)) = 0, Val(varPass(1)), Val(varPass(0)))
        If I = 0 Then zlCardNOEncrypt = strCardNo: Exit Function
        J = Len(strCardNo) - I: J = IIf(J < 0, 0, J)
        zlCardNOEncrypt = Mid(strCardNo, 1, J) & String(I, "*")
        Exit Function
    End If
    I = Val(varPass(0)): J = Val(varPass(1))
    If I > Len(strCardNo) Then zlCardNOEncrypt = strCardNo: Exit Function
    If J > Len(strCardNo) Then J = Len(strCardNo)
    If J < I Then J = I
   zlCardNOEncrypt = Mid(strCardNo, 1, I - 1) & String(J - I + 1, "*") & Mid(strCardNo, J + 1)
End Function
