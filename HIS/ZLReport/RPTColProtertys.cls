VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTColProtertys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(�������� As String, �����ֶ� As String, ������ϵ As String, ����ֵ As String, _
    ������ɫ As Long, ������ɫ As Long, �Ƿ�Ӵ� As Boolean, �Ƿ�����Ӧ�� As Boolean, _
    ���� As Integer, Optional sKey As String) As RPTColProterty
            
'�����¶���
    Dim objNewMember As RPTColProterty
    Set objNewMember = New RPTColProterty

    '���ô��뷽��������
    objNewMember.Key = Mid(sKey, 2)
    objNewMember.�������� = ��������
    objNewMember.�����ֶ� = �����ֶ�
    objNewMember.������ϵ = ������ϵ
    objNewMember.����ֵ = ����ֵ
    objNewMember.������ɫ = ������ɫ
    objNewMember.������ɫ = ������ɫ
    objNewMember.�Ƿ�Ӵ� = �Ƿ�Ӵ�
    objNewMember.�Ƿ�����Ӧ�� = �Ƿ�����Ӧ��
    objNewMember.���� = ����
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RPTColProterty
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
