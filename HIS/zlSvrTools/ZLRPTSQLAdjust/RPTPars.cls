VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTPars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"RPTPar"
Attribute VB_Ext_KEY = "Member0" ,"RPTPar"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(���� As String, ��� As Byte, ���� As String, ���� As Byte, ȱʡֵ As String, ��ʽ As Byte, ֵ�б� As String, ����SQL As String, ��ϸSQL As String, �����ֶ� As String, ��ϸ�ֶ� As String, ���� As String, Optional sKey As String, Optional Reserve As String) As RPTPar
    '�����¶���
    Dim objNewMember As RPTPar
    Set objNewMember = New RPTPar

    '���ô��뷽��������
    objNewMember.Key = Mid(sKey, 2)
    objNewMember.��� = ���
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.ȱʡֵ = ȱʡֵ
    objNewMember.��ʽ = ��ʽ
    objNewMember.ֵ�б� = ֵ�б�
    objNewMember.����SQL = ����SQL
    objNewMember.�����ֶ� = �����ֶ�
    objNewMember.��ϸSQL = ��ϸSQL
    objNewMember.��ϸ�ֶ� = ��ϸ�ֶ�
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.Reserve = Reserve
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RPTPar
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