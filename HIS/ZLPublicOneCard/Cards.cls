VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsCard"
Attribute VB_Ext_KEY = "Member0" ,"clsCard"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'�ֲ����������漯��
Private mCol As Collection
Private mvar�������ܼ� As String
Private mvar������� As String
Private mvar��ǰ�������ܼ� As String
Private mvar��ǰ������� As String
Private mvar���������ܼ� As String
Private mvar��������� As String
Private mvar��ȱʡ������ As Boolean
Private mvar������ʾ As Boolean
Private mvarģ��������� As String

Public Property Let ģ���������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvarvģ��������� = vData
End Property
Public Property Get ģ���������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ģ��������� = mvarģ���������
End Property

Public Property Let ������ʾ(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar������ʾ = vData
End Property
Public Property Get ������ʾ() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ������ʾ = mvar������ʾ
End Property

Public Property Let ��ȱʡ������(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar��ȱʡ������ = vData
End Property
Public Property Get ��ȱʡ������() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ��ȱʡ������ = mvar��ȱʡ������
End Property


Public Property Let ���������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar��������� = vData
End Property
Public Property Get ���������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ��������� = mvar���������
End Property
Public Property Let ���������ܼ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar���������ܼ� = vData
End Property
Public Property Get ���������ܼ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ���������ܼ� = mvar���������ܼ�
End Property

Public Property Let ��ǰ�������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar��ǰ������� = vData
End Property
Public Property Get ��ǰ�������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ��ǰ������� = mvar��ǰ�������
End Property
Public Property Let ��ǰ�������ܼ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar��ǰ�������ܼ� = vData
End Property
Public Property Get ��ǰ�������ܼ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ��ǰ�������ܼ� = mvar��ǰ�������ܼ�
End Property


Public Property Let �������(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar������� = vData
End Property
Public Property Get �������() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    ������� = mvar�������
End Property
Public Property Let �������ܼ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ǰ׺�ı� = 5
    mvar�������ܼ� = vData
End Property
Public Property Get �������ܼ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ǰ׺�ı�
    �������ܼ� = mvar�������ܼ�
End Property
Public Function AddItem(�ӿ���� As Long, �ӿڱ��� As String, ���� As String, ���� As String, _
    ���� As Boolean, ���ѿ� As Boolean, �ӿڳ����� As String, ���ƿ� As Boolean, �ɷ����� As Integer, _
    �Ƿ��Զ���ȡ As Integer, �Զ���ȡ��� As Integer, ϵͳ As Boolean, ���㷽ʽ As String, ǰ׺�ı� As String, _
    ���ų��� As Long, �Ƿ��ϸ���� As Boolean, �Ƿ�ˢ�� As Boolean, �Ƿ�����ʻ� As Boolean, �Ƿ�ȫ�� As Boolean, _
    �ض���Ŀ As String, ��ע As String, �����ظ�ʹ�� As Boolean, �������Ĺ��� As String, �Ƿ����� As Boolean, _
    ���볤�� As Integer, ���볤������ As Integer, ������� As Integer, _
    Optional sKey As String, Optional ͼ���ʶ As String = "", _
    Optional ֧������ As Boolean = False, Optional ֧��ͼ���ʶ As String = "", _
    Optional ���ܼ� As String, Optional ��� As String = "", _
    Optional �Ƿ��ƿ� As Boolean = False, Optional �Ƿ񷢿� As Boolean = False, _
    Optional �Ƿ�д�� As Boolean = False, Optional �������� As Integer, _
    Optional �Ƿ�ת�ʼ����� As Boolean = False, Optional �Ƿ�ɨ�� As Boolean, _
    Optional �Ƿ�Ӵ�ʽ���� As Boolean, Optional �Ƿ�ǽӴ�ʽ���� As Boolean, _
    Optional �Ƿ�ֿ����� As Boolean = False, Optional �Ƿ��˿��鿨 As Boolean = False, _
    Optional �Ƿ�֤�� As Boolean = False, Optional �豸�Ƿ����ûس� As Boolean = False, _
    Optional �Ƿ�ȱʡ���� As Boolean = False, Optional �Ƿ�������� As Boolean = False, _
    Optional �Ƿ�֧��ɨ�븶 As Boolean) As Card
   
    '�����¶���
    Dim objNewMember As Card
    Set objNewMember = New Card
    
    '���ô��뷽��������
    objNewMember.�ӿ���� = �ӿ����
    objNewMember.�ӿڱ��� = �ӿڱ���
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.�ӿڳ����� = �ӿڳ�����
    objNewMember.�ɷ����� = �ɷ�����
    objNewMember.�Ƿ��Զ���ȡ = �Ƿ��Զ���ȡ
    objNewMember.�Զ���ȡ��� = �Զ���ȡ���
    objNewMember.ϵͳ = ϵͳ
    objNewMember.���㷽ʽ = ���㷽ʽ
    objNewMember.ǰ׺�ı� = ǰ׺�ı�
    objNewMember.���ų��� = ���ų���
    objNewMember.���ƿ� = ���ƿ�
    objNewMember.���ѿ� = ���ѿ�
    objNewMember.�Ƿ��ϸ���� = �Ƿ��ϸ����
    objNewMember.�ض���Ŀ = �ض���Ŀ
    objNewMember.��ע = ��ע
    objNewMember.�Ƿ�����ʻ� = �Ƿ�����ʻ�
    objNewMember.�Ƿ�ȫ�� = �Ƿ�ȫ��
    objNewMember.�����ظ�ʹ�� = �����ظ�ʹ��
    objNewMember.�������Ĺ��� = �������Ĺ���
    objNewMember.�Ƿ����� = �Ƿ�����
    objNewMember.���볤�� = ���볤��
    objNewMember.���볤������ = ���볤������
    objNewMember.������� = �������
    objNewMember.ͼ���ʶ = ͼ���ʶ
    objNewMember.֧������ = ֧������
    objNewMember.֧��ͼ���ʶ = ֧��ͼ���ʶ
    objNewMember.���ܼ� = ���ܼ�
    objNewMember.��� = ���
    objNewMember.�Ƿ��ƿ� = �Ƿ��ƿ�    '56615
    objNewMember.�Ƿ񷢿� = �Ƿ񷢿�
    objNewMember.�Ƿ�д�� = �Ƿ�д��
    objNewMember.�������� = ��������
    objNewMember.�Ƿ�ģ������ = �Ƿ�ģ������
    objNewMember.ȱʡ��־ = ȱʡ��־
    '77872,���ϴ�,2014/9/15:�Ƿ�֧��ת�ʼ�����
    objNewMember.�Ƿ�ת�ʼ����� = �Ƿ�ת�ʼ�����
    objNewMember.�Ƿ�ˢ�� = �Ƿ�ˢ��    '85565,���ϴ�,2015/7/10:��������
    objNewMember.�Ƿ�ɨ�� = �Ƿ�ɨ��
    objNewMember.�Ƿ�Ӵ�ʽ���� = �Ƿ�Ӵ�ʽ����
    objNewMember.�Ƿ�ǽӴ�ʽ���� = �Ƿ�ǽӴ�ʽ����
    objNewMember.�Ƿ�ֿ����� = �Ƿ�ֿ�����
    objNewMember.�Ƿ��˿��鿨 = �Ƿ��˿��鿨
    '90875,���ϴ�,2016/11/8:����ҽ�ƿ�֤������
    objNewMember.�Ƿ�֤�� = �Ƿ�֤��
    objNewMember.�豸�Ƿ����ûس� = �豸�Ƿ����ûس�
    objNewMember.�Ƿ�ȱʡ���� = �Ƿ�ȱʡ����
    objNewMember.�Ƿ�������� = �Ƿ��������
    objNewMember.�Ƿ�֧��ɨ�븶 = �Ƿ�֧��ɨ�븶
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
   '�����Ѵ����Ķ���
    Set AddItem = objNewMember
    Set objNewMember = Nothing
End Function
Public Function Add(objCard As Card, Optional sKey As String) As Card
    If Len(sKey) = 0 Then
       mCol.Add objCard
    Else
        mCol.Add objCard, sKey
    End If
    Set Add = objCard
End Function

Public Property Get Item(vntIndexKey As Variant) As Card
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    count = mCol.count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub
