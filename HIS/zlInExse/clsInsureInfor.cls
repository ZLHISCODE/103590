VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsureInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������ֵ�ľֲ�����
Private mvar�Ƿ���ʽ��� As Boolean '�ֲ�����
Private mvar�ʻ���� As Double '�ֲ�����
Private mvar�����޶� As Double '�ֲ�����
Private mvar����͸֧ As Double '�ֲ�����
Private mvar����֧�� As Double '�ֲ�����
Private mvarͳ��֧�� As Double '�ֲ�����
Private mvar������֤��Ϣ As String '�ֲ�����
Private mvar���� As Integer '�ֲ�����
Private mvar�Ƿ�����ȫ�� As Boolean '�ֲ�����
Private mvarbytMCMode As Byte '�ֲ�����
Private mvar�������� As String '�ֲ�����
Private mvar�Ƿ��Զ���Ժ As Boolean '�ֲ�����
Public Property Let �Ƿ��Զ���Ժ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��Զ���Ժ = 5
    mvar�Ƿ��Զ���Ժ = vData
End Property


Public Property Get �Ƿ��Զ���Ժ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��Զ���Ժ
    �Ƿ��Զ���Ժ = mvar�Ƿ��Զ���Ժ
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



Public Property Let bytMCMode(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.bytMCMode = 5
    mvarbytMCMode = vData
End Property


Public Property Get bytMCMode() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.bytMCMode
    bytMCMode = mvarbytMCMode
End Property



Public Property Let �Ƿ�����ȫ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�����ȫ�� = 5
    mvar�Ƿ�����ȫ�� = vData
End Property


Public Property Get �Ƿ�����ȫ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�����ȫ��
    �Ƿ�����ȫ�� = mvar�Ƿ�����ȫ��
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



Public Property Let ������֤��Ϣ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������֤��Ϣ = 5
    mvar������֤��Ϣ = vData
End Property


Public Property Get ������֤��Ϣ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������֤��Ϣ
    ������֤��Ϣ = mvar������֤��Ϣ
End Property



Public Property Let ͳ��֧��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ͳ��֧�� = 5
    mvarͳ��֧�� = vData
End Property


Public Property Get ͳ��֧��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ͳ��֧��
    ͳ��֧�� = mvarͳ��֧��
End Property



Public Property Let ����֧��(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����֧�� = 5
    mvar����֧�� = vData
End Property


Public Property Get ����֧��() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����֧��
    ����֧�� = mvar����֧��
End Property



Public Property Let ����͸֧(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����͸֧ = 5
    mvar����͸֧ = vData
End Property


Public Property Get ����͸֧() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����͸֧
    ����͸֧ = mvar����͸֧
End Property



Public Property Let �����޶�(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�����޶� = 5
    mvar�����޶� = vData
End Property


Public Property Get �����޶�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����޶�
    �����޶� = mvar�����޶�
End Property



Public Property Let �ʻ����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ʻ���� = 5
    mvar�ʻ���� = vData
End Property


Public Property Get �ʻ����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ʻ����
    �ʻ���� = mvar�ʻ����
End Property



Public Property Let �Ƿ���ʽ���(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ���ʽ��� = 5
    mvar�Ƿ���ʽ��� = vData
End Property


Public Property Get �Ƿ���ʽ���() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ���ʽ���
    �Ƿ���ʽ��� = mvar�Ƿ���ʽ���
End Property


