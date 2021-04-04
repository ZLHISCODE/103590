VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"BillPage"
Attribute VB_Ext_KEY = "Member0" ,"BillPage"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 'Ҫ���������

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(Details As BillDetails, Optional NO As String, _
    Optional ��������ID As Long, Optional ������ As String, _
    Optional Ӧ�ս�� As Double, Optional ʵ�ս�� As Double, _
    Optional ���ս�� As Double, Optional ��Ԥ���� As Double, _
    Optional Ӧ�ɽ�� As Double, Optional ����� As Double, Optional �շѽ��� As String, _
    Optional ȫ�Ը� As Double, Optional ���Ը� As Double, Optional ����ͳ�� As Double, _
    Optional sKey As String, Optional str���ѿ����� As String, _
    Optional cur���ѿ�ˢ���� As Double, Optional ����ID As Long) As BillPage
    '�����¶���
    Dim objNewMember As BillPage
    Set objNewMember = New BillPage
    '���ô��뷽��������
    objNewMember.Key = sKey
    objNewMember.NO = NO
    objNewMember.��������ID = ��������ID
    objNewMember.������ = ������
    objNewMember.Ӧ�ս�� = Ӧ�ս��
    objNewMember.ʵ�ս�� = ʵ�ս��
    objNewMember.���ս�� = ���ս��
    objNewMember.��Ԥ���� = ��Ԥ����
    objNewMember.Ӧ�ɽ�� = Ӧ�ɽ��
    objNewMember.����� = �����
    objNewMember.�շѽ��� = �շѽ���
    objNewMember.ȫ�Ը� = ȫ�Ը�
    objNewMember.���Ը� = ���Ը�
    objNewMember.����ͳ�� = ����ͳ��
    objNewMember.���ѿ����� = str���ѿ�����
    objNewMember.���ѿ�ˢ���� = cur���ѿ�ˢ����
    objNewMember.����ID = ����ID
    Set objNewMember.Details = Details
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As BillPage
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