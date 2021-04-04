VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillInComes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"InCome"
Attribute VB_Ext_KEY = "Member0" ,"InCome"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 'Ҫ���������

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(������ĿID As Long, ������Ŀ As String, �վݷ�Ŀ As String, ��׼���� As Double, _
    Ӧ�ս�� As Double, ʵ�ս�� As Double, Optional ԭ�� As Double, Optional �ּ� As Double, _
    Optional sKey As String, Optional ͳ���� As Double, Optional varTag As Variant) As BillInCome
    '�����¶���
    Dim objNewMember As BillInCome
    Set objNewMember = New BillInCome

    '���ô��뷽��������
    objNewMember.Key = sKey
    objNewMember.������ĿID = ������ĿID
    objNewMember.������Ŀ = ������Ŀ
    objNewMember.�վݷ�Ŀ = �վݷ�Ŀ
    objNewMember.��׼���� = ��׼����
    objNewMember.ԭ�� = ԭ��
    objNewMember.�ּ� = �ּ�
    objNewMember.Ӧ�ս�� = Ӧ�ս��
    objNewMember.ʵ�ս�� = ʵ�ս��
    objNewMember.ͳ���� = ͳ����
    
    If IsObject(varTag) Then
        Set objNewMember.Tag = varTag
    Else
        objNewMember.Tag = varTag
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember ', sKey'��Ϊ�������������,���Բ�����Key
    End If
    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As BillInCome
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