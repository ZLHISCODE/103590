VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCardObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsCardObject"
Attribute VB_Ext_KEY = "Member0" ,"clsCardObject"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'�ֲ����������漯��
Private mCol As Collection

Public Function Add(CardObject As Object, ���ƿ� As Boolean, �ӿ���� As Long, CardPreporty As clsCard, ���ѿ� As Boolean, Optional sKey As String) As clsCardObject
    '�����¶���
    Dim objNewMember As clsCardObject
    Set objNewMember = New clsCardObject


    '���ô��뷽��������
     If IsObject(CardObject) Then
        Set objNewMember.CardObject = CardObject
    Else
        objNewMember.CardObject = CardObject
    End If
    objNewMember.���ƿ� = ���ƿ�
    objNewMember.�ӿ���� = �ӿ����
    If IsObject(CardPreporty) Then
        Set objNewMember.CardPreporty = CardPreporty
    Else
        objNewMember.CardPreporty = CardPreporty
    End If
    objNewMember.���ѿ� = ���ѿ�
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsCardObject
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
