VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'��������ֵ�ľֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar��ҳID As Long '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvarobjCard As Card '�ֲ�����
Private mvarԤ����� As Double '�ֲ�����
Private mvar������� As Double '�ֲ�����
Private mvarԤ��ʣ��ϼ� As Double '�ֲ�����
Private mvarʵ����� As Double '�ֲ�����
Private mvarδ���ۼ� As Double '�ֲ�����
Private mvar�������� As Integer '�ֲ�����
Private mvarbln�˿� As Boolean '�ֲ�����
Private mvarbln��Ժ As Boolean '�ֲ�����
Private mvarbln�������� As Boolean '�ֲ�����
Private mvar����� As String  '�ֲ�����
Private mvarסԺ�� As String  '�ֲ�����
Private mvarҽ�Ƹ��ʽ As String  '�ֲ�����
Private mvar���� As String '�ֲ�����

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

Public Property Let bln��������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.bln�������� = 5
    mvarbln�������� = vData
End Property


Public Property Get bln��������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.bln��������
    bln�������� = mvarbln��������
End Property



Public Property Let bln��Ժ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.bln��Ժ = 5
    mvarbln��Ժ = vData
End Property


Public Property Get bln��Ժ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.bln��Ժ
    bln��Ժ = mvarbln��Ժ
End Property



Public Property Let bln�˿�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.bln�˿� = 5
    mvarbln�˿� = vData
End Property


Public Property Get bln�˿�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.bln�˿�
    bln�˿� = mvarbln�˿�
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



Public Property Let δ���ۼ�(ByVal vData As Double)
Attribute δ���ۼ�.VB_Description = "�ϴ�δ���ۼƽ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.δ���ۼ� = 5
    mvarδ���ۼ� = vData
End Property


Public Property Get δ���ۼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.δ���ۼ�
      δ���ۼ� = mvarδ���ۼ�
End Property



Public Property Let ʵ�����(ByVal vData As Double)
Attribute ʵ�����.VB_Description = "����Ԥ����ϸ�е�ʵ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ����� = 5
    mvarʵ����� = vData
End Property


Public Property Get ʵ�����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʵ�����
    ʵ����� = mvarʵ�����
End Property



Public Property Let Ԥ��ʣ��ϼ�(ByVal vData As Double)
Attribute Ԥ��ʣ��ϼ�.VB_Description = "Ԥ�����-δ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ��ʣ��ϼ� = 5
    mvarԤ��ʣ��ϼ� = vData
End Property


Public Property Get Ԥ��ʣ��ϼ�() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ��ʣ��ϼ�
    Ԥ��ʣ��ϼ� = mvarԤ��ʣ��ϼ�
End Property



Public Property Let �������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������� = 5
    mvar������� = vData
End Property


Public Property Get �������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�������
    ������� = mvar�������
End Property



Public Property Let Ԥ�����(ByVal vData As Double)
Attribute Ԥ�����.VB_Description = "Ԥ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Ԥ����� = 5
    mvarԤ����� = vData
End Property


Public Property Get Ԥ�����() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Ԥ�����
    Ԥ����� = mvarԤ�����
End Property



Public Property Set objCard(ByVal vData As Card)
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.objCard = Form1
    Set mvarobjCard = vData
End Property


Public Property Get objCard() As Card
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.objCard
    Set objCard = mvarobjCard
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



Public Property Let �Ա�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ա� = 5
    mvar�Ա� = vData
End Property


Public Property Get �Ա�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ա�
    �Ա� = mvar�Ա�
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



Public Property Let ��ҳID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ҳID = 5
    mvar��ҳID = vData
End Property


Public Property Get ��ҳID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ҳID
    ��ҳID = mvar��ҳID
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


Public Property Let �����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar����� = vData
End Property


Public Property Get �����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ����� = mvar�����
End Property

Public Property Let סԺ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvarסԺ�� = vData
End Property


Public Property Get סԺ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    סԺ�� = mvarסԺ��
End Property

Public Property Let ҽ�Ƹ��ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvarҽ�Ƹ��ʽ = vData
End Property


Public Property Get ҽ�Ƹ��ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ҽ�Ƹ��ʽ = mvarҽ�Ƹ��ʽ
End Property
