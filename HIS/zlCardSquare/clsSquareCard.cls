VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSquareCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'���ѿ���Ϣ
Private mvar������ As String
Private mvar���� As String
Private mvar��ֵ�� As Boolean
Private mvar��Ч�� As String
Private mvar����ԭ�� As String
Private mvar������ As String
Private mvar�쿨�� As String
Private mvar����ID As Long
Private mvar����ʱ�� As String
Private mvar�쿨����id As Long
Private mvar�쿨���� As String '����-����
Private mvar��ע As String
Private mvar����� As Double
Private mvar����ֵ As Double
Private mvarʵ������ As Double
Private mvar��ֵ�ۿ��� As Double
Private mvar������ As String
Private mvar����ʱ�� As String
Private mvarͣ���� As String
Private mvarͣ������ As String
Private mvar��ǰ״̬ As String
Private mvar�ѳ�ֵ As Boolean '���ڳ�ֵ��¼�ľͲ������޸ġ�������ֵ��������
Private mvarԭ���� As String
Private mvar������� As String

Public Property Let ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property

Public Property Get ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������
    ������ = mvar������
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
    
    
Public Property Let ��ֵ��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ֵ�� = 5
    mvar��ֵ�� = vData
End Property

Public Property Get ��ֵ��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ֵ��
    ��ֵ�� = mvar��ֵ��
End Property
    

Public Property Let ��Ч��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ч�� = 5
    mvar��Ч�� = vData
End Property

Public Property Get ��Ч��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ч��
    ��Ч�� = mvar��Ч��
End Property
    

Public Property Let ����ԭ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ԭ�� = 5
    mvar����ԭ�� = vData
End Property

Public Property Get ����ԭ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ԭ��
    ����ԭ�� = mvar����ԭ��
End Property
    
 
Public Property Let ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property

Public Property Get ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������
    ������ = mvar������
End Property
    

Public Property Let �쿨��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�쿨�� = 5
    mvar�쿨�� = vData
End Property

Public Property Get �쿨��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�쿨��
    �쿨�� = mvar�쿨��
End Property


Public Property Let ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar����ID = vData
End Property

Public Property Get ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    ����ID = mvar����ID
End Property


Public Property Let ����ʱ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ʱ�� = 5
    mvar����ʱ�� = vData
End Property

Public Property Get ����ʱ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ʱ��
    ����ʱ�� = mvar����ʱ��
End Property


Public Property Let �쿨����id(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�쿨����id = 5
    mvar�쿨����id = vData
End Property

Public Property Get �쿨����id() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�쿨����id
    �쿨����id = mvar�쿨����id
End Property


Public Property Let �쿨����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�쿨���� = 5
    mvar�쿨���� = vData
End Property

Public Property Get �쿨����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�쿨����
    �쿨���� = mvar�쿨����
End Property


Public Property Let ��ע(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ע = 5
    mvar��ע = vData
End Property

Public Property Get ��ע() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ע
    ��ע = mvar��ע
End Property

    
Public Property Let �����(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����� = 5
    mvar����� = vData
End Property

Public Property Get �����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����
    ����� = mvar�����
End Property


Public Property Let ����ֵ(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ֵ = 5
    mvar����ֵ = vData
End Property

Public Property Get ����ֵ() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ֵ
    ����ֵ = mvar����ֵ
End Property


Public Property Let ʵ������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ������ = 5
    mvarʵ������ = vData
End Property

Public Property Get ʵ������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʵ������
    ʵ������ = mvarʵ������
End Property


Public Property Let ��ֵ�ۿ���(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ֵ�ۿ��� = 5
    mvar��ֵ�ۿ��� = vData
End Property

Public Property Get ��ֵ�ۿ���() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ֵ�ۿ���
    ��ֵ�ۿ��� = mvar��ֵ�ۿ���
End Property


Public Property Let ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property

Public Property Get ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������
    ������ = mvar������
End Property


Public Property Let ����ʱ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ʱ�� = 5
    mvar����ʱ�� = vData
End Property

Public Property Get ����ʱ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ʱ��
    ����ʱ�� = mvar����ʱ��
End Property


Public Property Let ͣ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ͣ���� = 5
    mvarͣ���� = vData
End Property

Public Property Get ͣ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ͣ����
    ͣ���� = mvarͣ����
End Property


Public Property Let ͣ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ͣ������ = 5
    mvarͣ������ = vData
End Property

Public Property Get ͣ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ͣ������
    ͣ������ = mvarͣ������
End Property


Public Property Let ��ǰ״̬(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ǰ״̬ = 5
    mvar��ǰ״̬ = vData
End Property

Public Property Get ��ǰ״̬() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ǰ״̬
    ��ǰ״̬ = mvar��ǰ״̬
End Property


Public Property Let �ѳ�ֵ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ѳ�ֵ = 5
    mvar�ѳ�ֵ = vData
End Property

Public Property Get �ѳ�ֵ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ѳ�ֵ
    �ѳ�ֵ = mvar�ѳ�ֵ
End Property


Public Property Let ԭ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ԭ���� = 5
    mvarԭ���� = vData
End Property

Public Property Get ԭ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ԭ����
    ԭ���� = mvarԭ����
End Property


Public Property Let �������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������� = 5
    mvar������� = vData
End Property

Public Property Get �������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������
    ������� = mvar�������
End Property