VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������ֵ�ľֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�Ա� As String '�ֲ�����
Private mvar�������� As String '�ֲ�����
Private mvar����֤�� As String '�ֲ�����
Private mvar����ID As Long '�ֲ�����
Private mvar��ҳID As Long '�ֲ�����
Private mvar����� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��ͥ��ַ As String '�ֲ�����
Private mvar��ϵ�绰 As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�����  As Double   '�ֲ�����
Private mvar�ѱ� As String '�ֲ�����
Private mvarҽ�Ƹ��ʽ As String '�ֲ�����
Private mvarҽ�Ƹ��ʽ���� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar������ַ As String '�ֲ�����
Private mvar����֤�� As String '�ֲ�����
Private mvarְҵ As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvarѧ�� As String '�ֲ�����
Private mvar����״�� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��ͥ�绰 As String '�ֲ�����
Private mvar��ͥ�ʱ� As String '�ֲ�����
Private mvar��ϵ�˹�ϵ As String '�ֲ�����
Private mvar��ϵ�� As String '�ֲ�����
Private mvar��ϵ�˵�ַ As String '�ֲ�����
Private mvar��ϵ�˵绰 As String '�ֲ�����
Private mvar������λ As String '�ֲ�����
Private mvar������λ��ַ As String '�ֲ�����
Private mvar������λ�绰 As String '�ֲ�����
Private mvar������λ�ʱ� As String '�ֲ�����
Private mvar������λ������ As String '�ֲ�����
Private mvar������λ�������ʻ� As String '�ֲ�����
Private mvar���ڵ�ַ As String '�ֲ�����
Private mvar���ڵ�ַ�ʱ� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar�໤�� As String  '�ֲ�����
Private mvar���ѿ�ID As Long    '�ֲ�����
Private mvar���� As Long    '�ֲ�����
Private mvar�������� As String '�ֲ�����
Private mvar���￨�� As String
Private mvar�ֻ��� As String
Private mvar��Ƭ As StdPicture
Private mvar��Ƭ�ļ� As String
Private mvar�������� As String '�ֲ�����
Private mvarTag As String '�ֲ�����
Private mvarPatiXML As String '�ֲ�����
Private mvarҽ���� As String '�ֲ�����
Private mvar����ʱ�� As String '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar��ͬ��λID As Long '�ֲ�����
Private mvar��ϵ������֤�� As String '�ֲ�����
Private mvar�Ǽ�ʱ�� As String '�ֲ�����
Private mvar��Ժ As Boolean '�ֲ�����
Private mvar���� As Boolean '�ֲ�����
Private mvar��ǰ����ID As Long '�ֲ�����
Private mvar��ǰ�������� As String '�ֲ�����
Private mvar��ǰ����ID As Long '�ֲ�����
Private mvar��ǰ�������� As String '�ֲ�����
Private mvar��Ժ����ID As Long '�ֲ�����
Private mvar���� As String '�ֲ�����
Private mvar����� As Boolean '�ֲ�����
Private mvarסԺ�� As String '�ֲ�����
Private mvar�������� As Byte '�ֲ�����
Private mvar��Ժ���� As String '�ֲ�����
Private mvar��Ժ���� As String '�ֲ�����
Private mvarסԺҽʦ As String '�ֲ�����
Private mvar���˱�ע As String '�ֲ�����
Private mvarסԺ״̬ As Byte '�ֲ�����
Private mvar��˱�־ As Byte '�ֲ�����
Private mvar��Ŀ���� As String '�ֲ�����
Private mvar����֤�� As String '�ֲ�����
Private mvar�Һ�ID As Long '�ֲ�����
Private mvarʵ����֤ As Boolean  '�ֲ�����
Private mvar�����ҳID As Long '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar����ʱ�� As String '�ֲ�����
'��������ֵ�ľֲ�����
Private mvarҽ������ As String '�ֲ�����

Public Property Let ҽ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ������ = 5
    mvarҽ������ = vData
End Property


Public Property Get ҽ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ������
    ҽ������ = mvarҽ������
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





Public Property Let �����ҳID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�����ҳID = 5
    mvar�����ҳID = vData
End Property


Public Property Get �����ҳID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����ҳID
    �����ҳID = mvar�����ҳID
End Property


Public Property Let ʵ����֤(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ʵ����֤ = 5
    mvarʵ����֤ = vData
End Property

Public Property Get ʵ����֤() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ʵ����֤
    ʵ����֤ = mvarʵ����֤
End Property


Public Property Let ��ǰ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ǰ�������� = 5
    mvar��ǰ�������� = vData
End Property

Public Property Get ��ǰ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ǰ��������
    ��ǰ�������� = mvar��ǰ��������
End Property

Public Property Let ��ǰ��������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ǰ�������� = 5
    mvar��ǰ�������� = vData
End Property

Public Property Get ��ǰ��������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ǰ��������
    ��ǰ�������� = mvar��ǰ��������
End Property

Public Property Let ��Ŀ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ŀ���� = 5
    mvar��Ŀ���� = vData
End Property

Public Property Get ��Ŀ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ŀ����
    ��Ŀ���� = mvar��Ŀ����
End Property

Public Property Let ��˱�־(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��˱�־ = 5
    mvar��˱�־ = vData
End Property

Public Property Get ��˱�־() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��˱�־
    ��˱�־ = mvar��˱�־
End Property

Public Property Let סԺ״̬(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.סԺ״̬ = 5
    mvarסԺ״̬ = vData
End Property

Public Property Get סԺ״̬() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.סԺ״̬
    סԺ״̬ = mvarסԺ״̬
End Property

Public Property Let ���˱�ע(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���˱�ע = 5
    mvar���˱�ע = vData
End Property

Public Property Get ���˱�ע() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���˱�ע
    ���˱�ע = mvar���˱�ע
End Property

Public Property Let סԺҽʦ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.סԺҽʦ = 5
    mvarסԺҽʦ = vData
End Property

Public Property Get סԺҽʦ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.סԺҽʦ
    סԺҽʦ = mvarסԺҽʦ
End Property

Public Property Let ��Ժ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ժ���� = 5
    mvar��Ժ���� = vData
End Property

Public Property Get ��Ժ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ժ����
    ��Ժ���� = mvar��Ժ����
End Property

Public Property Let ��Ժ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ժ���� = 5
    mvar��Ժ���� = vData
End Property

Public Property Get ��Ժ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ժ����
    ��Ժ���� = mvar��Ժ����
End Property

Public Property Let ��������(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�������� = 5
    mvar�������� = vData
End Property

Public Property Get ��������() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��������
    �������� = mvar��������
End Property

Public Property Let סԺ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.סԺ�� = 5
    mvarסԺ�� = vData
End Property

Public Property Get סԺ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.סԺ��
    סԺ�� = mvarסԺ��
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

Public Property Let �����(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar����� = vData
End Property

Public Property Get �����() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ����� = mvar�����
End Property

Public Property Let ��ǰ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar��ǰ����ID = vData
End Property

Public Property Get ��ǰ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    ��ǰ����ID = mvar��ǰ����ID
End Property

Public Property Let ��ǰ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ǰ����ID = 5
    mvar��ǰ����ID = vData
End Property

Public Property Get ��ǰ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ǰ����ID
    ��ǰ����ID = mvar��ǰ����ID
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

Public Property Let ��Ժ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ժ = 5
    mvar��Ժ = vData
End Property

Public Property Get ��Ժ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ժ
    ��Ժ = mvar��Ժ
End Property

Public Property Let �Ǽ�ʱ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ǽ�ʱ�� = 5
    mvar�Ǽ�ʱ�� = vData
End Property

Public Property Get �Ǽ�ʱ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ǽ�ʱ��
    �Ǽ�ʱ�� = mvar�Ǽ�ʱ��
End Property

Public Property Let ��ϵ������֤��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ������֤�� = 5
    mvar��ϵ������֤�� = vData
End Property

Public Property Get ��ϵ������֤��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ������֤��
    ��ϵ������֤�� = mvar��ϵ������֤��
End Property

Public Property Let ��ͬ��λID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ͬ��λID = 5
    mvar��ͬ��λID = vData
End Property

Public Property Get ��ͬ��λID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ͬ��λID
    ��ͬ��λID = mvar��ͬ��λID
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

Public Property Let ҽ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ҽ���� = 5
    mvarҽ���� = vData
End Property

Public Property Get ҽ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ҽ����
    ҽ���� = mvarҽ����
End Property

Public Property Let PatiXML(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvarPatiXML = vData
End Property

Public Property Get PatiXML() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    PatiXML = mvarPatiXML
End Property

Public Property Let Tag(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Tag = 5
    mvarTag = vData
End Property

Public Property Get Tag() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
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

Public Property Let ��Ƭ�ļ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ļ�Ƭ = 5
    mvar��Ƭ�ļ� = vData
End Property

Public Property Get ��Ƭ�ļ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ƭ�ļ�
    ��Ƭ�ļ� = mvar��Ƭ�ļ�
End Property

Public Property Let ��Ƭ(ByVal vData As StdPicture)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��Ƭ = 5
    Set mvar��Ƭ = vData
End Property

Public Property Get ��Ƭ() As StdPicture
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��Ƭ
    Set ��Ƭ = mvar��Ƭ
End Property

Public Property Let �ֻ���(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�ֻ��� = 5
    mvar�ֻ��� = vData
End Property

Public Property Get �ֻ���() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�ֻ���
    �ֻ��� = mvar�ֻ���
End Property

Public Property Let ���￨��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���￨�� = vData
End Property

Public Property Get ���￨��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���￨�� = mvar���￨��
End Property

Public Property Let ����(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property

Public Property Get ����() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property

Public Property Let ���ѿ�ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ѿ�ID = 5
    mvar���ѿ�ID = vData
End Property

Public Property Get ���ѿ�ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ѿ�ID
    ���ѿ�ID = mvar���ѿ�ID
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

Public Property Let ���ڵ�ַ�ʱ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ڵ�ַ�ʱ� = 5
    mvar���ڵ�ַ�ʱ� = vData
End Property

Public Property Get ���ڵ�ַ�ʱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ڵ�ַ�ʱ�
    ���ڵ�ַ�ʱ� = mvar���ڵ�ַ�ʱ�
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

Public Property Let ���ڵ�ַ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���ڵ�ַ = 5
    mvar���ڵ�ַ = vData
End Property

Public Property Get ���ڵ�ַ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.���ڵ�ַ
    ���ڵ�ַ = mvar���ڵ�ַ
End Property

Public Property Let ������λ�������ʻ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������λ�������ʻ� = 5
    mvar������λ�������ʻ� = vData
End Property

Public Property Get ������λ�������ʻ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������λ�������ʻ�
    ������λ�������ʻ� = mvar������λ�������ʻ�
End Property

Public Property Let ������λ������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������λ������ = 5
    mvar������λ������ = vData
End Property

Public Property Get ������λ������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������λ������
    ������λ������ = mvar������λ������
End Property

Public Property Let ������λ�ʱ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������λ�ʱ� = 5
    mvar������λ�ʱ� = vData
End Property

Public Property Get ������λ�ʱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������λ�ʱ�
    ������λ�ʱ� = mvar������λ�ʱ�
End Property

Public Property Let ������λ�绰(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������λ�绰 = 5
    mvar������λ�绰 = vData
End Property

Public Property Get ������λ�绰() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������λ�绰
    ������λ�绰 = mvar������λ�绰
End Property

Public Property Let ������λ��ַ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������λ��ַ = 5
    mvar������λ��ַ = vData
End Property

Public Property Get ������λ��ַ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������λ��ַ
    ������λ��ַ = mvar������λ��ַ
End Property

Public Property Let ������λ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������λ = 5
    mvar������λ = vData
End Property

Public Property Get ������λ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������λ
    ������λ = mvar������λ
End Property

Public Property Let ��ϵ�˵绰(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�˵绰 = 5
    mvar��ϵ�˵绰 = vData
End Property

Public Property Get ��ϵ�˵绰() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�˵绰
    ��ϵ�˵绰 = mvar��ϵ�˵绰
End Property

Public Property Let ��ϵ�˵�ַ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�˵�ַ = 5
    mvar��ϵ�˵�ַ = vData
End Property

Public Property Get ��ϵ�˵�ַ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�˵�ַ
    ��ϵ�˵�ַ = mvar��ϵ�˵�ַ
End Property

Public Property Let ��ϵ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�� = 5
    mvar��ϵ�� = vData
End Property

Public Property Get ��ϵ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ��
    ��ϵ�� = mvar��ϵ��
End Property

Public Property Let ��ϵ�˹�ϵ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�˹�ϵ = 5
    mvar��ϵ�˹�ϵ = vData
End Property

Public Property Get ��ϵ�˹�ϵ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�˹�ϵ
    ��ϵ�˹�ϵ = mvar��ϵ�˹�ϵ
End Property

Public Property Let �໤��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�໤�� = 5
    mvar�໤�� = vData
End Property

Public Property Get �໤��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�໤��
    �໤�� = mvar�໤��
End Property

Public Property Let ��ͥ�ʱ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ͥ�ʱ� = 5
    mvar��ͥ�ʱ� = vData
End Property

Public Property Get ��ͥ�ʱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ͥ�ʱ�
    ��ͥ�ʱ� = mvar��ͥ�ʱ�
End Property

Public Property Let ��ͥ�绰(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ͥ�绰 = 5
    mvar��ͥ�绰 = vData
End Property

Public Property Get ��ͥ�绰() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ͥ�绰
    ��ͥ�绰 = mvar��ͥ�绰
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

Public Property Let ����״��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����״�� = 5
    mvar����״�� = vData
End Property

Public Property Get ����״��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����״��
    ����״�� = mvar����״��
End Property

Public Property Let ѧ��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ѧ�� = 5
    mvarѧ�� = vData
End Property

Public Property Get ѧ��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ѧ��
    ѧ�� = mvarѧ��
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

Public Property Let ְҵ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ְҵ = 5
    mvarְҵ = vData
End Property

Public Property Get ְҵ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ְҵ
    ְҵ = mvarְҵ
End Property

Public Property Let ����֤��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����֤�� = 5
    mvar����֤�� = vData
End Property

Public Property Get ����֤��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����֤��
    ����֤�� = mvar����֤��
End Property

Public Property Let ������ַ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ַ = 5
    mvar������ַ = vData
End Property

Public Property Get ������ַ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������ַ
    ������ַ = mvar������ַ
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

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�绰 = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�绰
    ���� = mvar����
End Property

Public Property Let ҽ�Ƹ��ʽ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�绰 = 5
    mvarҽ�Ƹ��ʽ = vData
End Property

Public Property Get ҽ�Ƹ��ʽ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�绰
    ҽ�Ƹ��ʽ = mvarҽ�Ƹ��ʽ
End Property

Public Property Let ҽ�Ƹ��ʽ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�绰 = 5
    mvarҽ�Ƹ��ʽ���� = vData
End Property

Public Property Get ҽ�Ƹ��ʽ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�绰
    ҽ�Ƹ��ʽ���� = mvarҽ�Ƹ��ʽ����
End Property

Public Property Let �ѱ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�绰 = 5
    mvar�ѱ� = vData
End Property

Public Property Get �ѱ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�绰
    �ѱ� = mvar�ѱ�
End Property

Public Property Let ����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�绰 = 5
    mvar���� = vData
End Property

Public Property Get ����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�绰
    ���� = mvar����
End Property

Public Property Let ��ϵ�绰(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ϵ�绰 = 5
    mvar��ϵ�绰 = vData
End Property

Public Property Get ��ϵ�绰() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ϵ�绰
    ��ϵ�绰 = mvar��ϵ�绰
End Property

Public Property Let �����(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����� = 5
    mvar����� = vData
End Property

Public Property Get �����() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�����
    ����� = mvar�����
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

Public Property Let ����֤��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����֤�� = 5
    mvar����֤�� = vData
End Property

Public Property Get ����֤��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����֤��
    ����֤�� = mvar����֤��
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

Public Property Let ��ͥ��ַ(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ͥ��ַ = 5
    mvar��ͥ��ַ = vData
End Property

Public Property Get ��ͥ��ַ() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ͥ��ַ
    ��ͥ��ַ = mvar��ͥ��ַ
End Property

Public Property Let ����֤��(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.��ͥ��ַ = 5
    mvar����֤�� = vData
End Property

Public Property Get ����֤��() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.��ͥ��ַ
    ����֤�� = mvar����֤��
End Property

Public Property Let ��Ժ����ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar��Ժ����ID = vData
End Property

Public Property Get ��Ժ����ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    ��Ժ����ID = mvar��Ժ����ID
End Property

Public Property Let �Һ�ID(ByVal vData As Long)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.����ID = 5
    mvar�Һ�ID = vData
End Property

Public Property Get �Һ�ID() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����ID
    �Һ�ID = mvar�Һ�ID
End Property
