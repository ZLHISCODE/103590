VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBalanceItem"
Attribute VB_Ext_KEY = "Member0" ,"clsBalanceItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'�ֲ����������漯��
Private mCol As Collection
Public Enum gBalanceType
    gEM_��ͨ���� = 0
    gEM_Ԥ���� = 1
    gEM_ҽ�� = 2
    gEM_һ��ͨ = 3
    gEM_һ��ͨ_�� = 4
    gEm_���ѿ� = 5
End Enum

'��������ֵ�ľֲ�����
Private mvar������ As Double '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar�Ƿ�ת�� As Boolean '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar�˷ѽ���IDs As String '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar�Ƿ��쳣�˷� As Boolean '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar���� As gBalanceType '�ֲ�����
'��������ֵ�ľֲ�����
Private mvar�շ����� As Byte '�ֲ�����
Public Property Let �շ�����(ByVal vData As Byte)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�շ����� = 5
    mvar�շ����� = vData
End Property


Public Property Get �շ�����() As Byte
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�շ�����
    �շ����� = mvar�շ�����
End Property



Public Property Let ����(ByVal vData As gBalanceType)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.���� = 5
    mvar���� = vData
End Property


Public Property Get ����() As gBalanceType
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.����
    ���� = mvar����
End Property



Public Property Let �Ƿ��쳣�˷�(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ��쳣�˷� = 5
    mvar�Ƿ��쳣�˷� = vData
End Property


Public Property Get �Ƿ��쳣�˷�() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ��쳣�˷�
    �Ƿ��쳣�˷� = mvar�Ƿ��쳣�˷�
End Property



Public Property Let �˷ѽ���IDs(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�˷ѽ���IDs = 5
    mvar�˷ѽ���IDs = vData
End Property


Public Property Get �˷ѽ���IDs() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�˷ѽ���IDs
    �˷ѽ���IDs = mvar�˷ѽ���IDs
End Property



Public Property Let �Ƿ�ת��(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ƿ�ת�� = 5
    mvar�Ƿ�ת�� = vData
End Property


Public Property Get �Ƿ�ת��() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ƿ�ת��
    �Ƿ�ת�� = mvar�Ƿ�ת��
End Property



Public Property Let ������(ByVal vData As Double)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.������ = 5
    mvar������ = vData
End Property
Public Property Get ������() As Double
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.������
    ������ = mvar������
End Property

Public Function Add(Key As String, objCard As Card, �����ID As Long, �Ƿ�ת�� As Boolean, Ԥ��ID As Long, ������ˮ�� As String, ����˵�� As String, ������ As Double, ���� As String, ����ID As Long, Tag As String, Optional sKey As String) As clsBalanceItem
    '�����¶���
    Dim objNewMember As clsBalanceItem
    Set objNewMember = New clsBalanceItem


    '���ô��뷽��������
    objNewMember.Key = Key
    If IsObject(objCard) Then
        Set objNewMember.objCard = objCard
    Else
        objNewMember.objCard = objCard
    End If
    objNewMember.�����ID = �����ID
    objNewMember.�Ƿ�ת�� = �Ƿ�ת��
    objNewMember.Ԥ��ID = Ԥ��ID
    objNewMember.������ˮ�� = ������ˮ��
    objNewMember.����˵�� = ����˵��
    objNewMember.������ = ������
    objNewMember.���� = ����
    objNewMember.����ID = ����ID
    objNewMember.Tag = Tag
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Sub AddItem(ByVal objItem As clsBalanceItem, Optional sKey As String)
    If Len(sKey) = 0 Then
        mCol.Add objItem
    Else
        mCol.Add objItem, sKey
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As clsBalanceItem
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
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
