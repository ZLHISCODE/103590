VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusinessDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBusinessData"
Attribute VB_Ext_KEY = "Member0" ,"clsBusinessData"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'�ֲ����������漯��
Private mCol As Collection
Private mvarTag As String
Private mvar�б��ӵ� As String

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
Public Property Let �б��ӵ�(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.�Ǽ�ʱ�� = 5
    mvar�б��ӵ� = vData
End Property

Public Property Get �б��ӵ�() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.�Ǽ�ʱ��
    �б��ӵ� = mvar�б��ӵ�
End Property


Public Function Add(�������� As gDataType, �������� As String, �Ƿ��ֵ As Boolean, �Ƿ��ѡ As Boolean, �������� As String, �ڵ����� As String, ����˵�� As String, Optional sKey As String) As clsBusinessData
    '�����¶���
    Dim objNewMember As clsBusinessData
    Set objNewMember = New clsBusinessData

    '���ô��뷽��������
    objNewMember.�������� = ��������
    objNewMember.�������� = ��������
    objNewMember.�������� = ��������
    objNewMember.����˵�� = ����˵��
    objNewMember.�Ƿ��ֵ = �Ƿ��ֵ
    objNewMember.�Ƿ��ѡ = �Ƿ��ѡ
    objNewMember.�ڵ����� = �ڵ�����
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function
Public Sub AddItem(objBusinessData As clsBusinessData, Optional sKey As String)
   If Len(sKey) = 0 Then
        mCol.Add objBusinessData
    Else
        mCol.Add objBusinessData, sKey
    End If
End Sub
Public Property Get Item(vntIndexKey As Variant) As clsBusinessData
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
