VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Detail"
Attribute VB_Ext_KEY = "Member0" ,"Detail"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(ID As Long, ��� As String, ������� As String, ���� As String, ���� As String, _
    ���� As String, ���� As String, ��� As String, ���㵥λ As String, ˵�� As String, _
    ���ηѱ� As Boolean, ��� As Boolean, �Ӱ�Ӽ� As Boolean, ִ�п��� As Byte, ������� As Integer, _
    ���� As String, Optional sKey As String, Optional �۸� As Currency, Optional Ҫ������ As Boolean, Optional ¼������ As Double) As Detail
    '�����¶���
    Dim objNewMember As Detail
    Set objNewMember = New Detail
    
    '���ô��뷽��������
    objNewMember.Key = sKey
    objNewMember.ID = ID
    objNewMember.��� = ���
    objNewMember.������� = �������
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.��� = ���
    objNewMember.���㵥λ = ���㵥λ
    objNewMember.˵�� = ˵��
    objNewMember.���ηѱ� = ���ηѱ�
    objNewMember.��� = ���
    objNewMember.�Ӱ�Ӽ� = �Ӱ�Ӽ�
    objNewMember.ִ�п��� = ִ�п���
    objNewMember.������� = �������
    objNewMember.¼������ = ¼������
    objNewMember.���� = ����
    objNewMember.�۸� = �۸�
    objNewMember.Ҫ������ = Ҫ������
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As Detail
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��'debug.print x.Count
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
