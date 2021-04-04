VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"RPTItem"
Attribute VB_Ext_KEY = "Member0" ,"RPTItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(id As Integer, ��ʽ�� As Byte, ���� As String, �ϼ�ID As Integer, ���� As Byte, _
    ��� As Integer, ���� As String, ���� As Byte, ���� As String, ��ͷ As String, X As Long, _
    Y As Long, W As Long, H As Long, �и� As Long, ���� As Byte, �Ե� As Boolean, ���� As String, _
    �ֺ� As Single, ���� As Boolean, ���� As Boolean, б�� As Boolean, ���� As Long, ǰ�� As Long, _
    ���� As Long, �߿� As Boolean, ���� As Byte, ���� As String, ��ʽ As String, ���� As String, _
    �����߼Ӵ� As Boolean, ����Ӧ�и� As Boolean, _
    Optional ͼƬ As StdPicture, Optional ϵͳ As Boolean, Optional ��ID As Integer, Optional SubIDs As RelatIDs, _
    Optional CopyIDs As RelatIDs, Optional sKey As String, Optional ����Դ As String, Optional ���¼�� As Long, _
    Optional ���Ҽ�� As Long, Optional Դ�к� As Long, Optional ������� As Long, Optional ������� As Long, _
    Optional Relations As RPTRelations, Optional ColProtertys As RPTColProtertys, _
    Optional ˮƽ��ת As Boolean) As RPTItem
    
    '�����¶���
    Dim objNewMember As RPTItem
    Set objNewMember = New RPTItem

    '����Ѿ������ˣ��Ͳ��ٸ�����
    On Error Resume Next
    If Not mCol.Item(sKey) Is Nothing Then
        If Err.Number = 0 Then
            Set objNewMember.CopyIDs = CopyIDs
            Set Add = objNewMember
            Exit Function
        End If
        Err.Clear: On Error GoTo 0
    End If
    '���ô��뷽��������
    objNewMember.Key = Mid(sKey, 2)
    objNewMember.id = id
    objNewMember.��ʽ�� = ��ʽ��
    objNewMember.���� = ����
    objNewMember.�ϼ�ID = �ϼ�ID
    objNewMember.���� = ����
    objNewMember.��� = ���
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.��ͷ = ��ͷ
    objNewMember.X = X
    objNewMember.Y = Y
    objNewMember.W = W
    objNewMember.H = H
    objNewMember.�и� = �и�
    objNewMember.����Ӧ�и� = ����Ӧ�и�
    objNewMember.���� = ����
    objNewMember.�Ե� = �Ե�
    objNewMember.���� = ����
    objNewMember.�ֺ� = �ֺ�
    objNewMember.���� = ����
    objNewMember.�����߼Ӵ� = �����߼Ӵ�
    objNewMember.���� = ����
    objNewMember.б�� = б��
    objNewMember.���� = ����
    objNewMember.ǰ�� = ǰ��
    objNewMember.���� = ����
    objNewMember.�߿� = �߿�
    objNewMember.���� = ����
    objNewMember.���� = ����
    objNewMember.��ʽ = ��ʽ
    objNewMember.���� = ����
    objNewMember.ϵͳ = ϵͳ
    objNewMember.��ID = ��ID
    objNewMember.����Դ = ����Դ
    objNewMember.���¼�� = ���¼��
    objNewMember.���Ҽ�� = ���Ҽ��
    objNewMember.Դ�к� = Դ�к�
    objNewMember.������� = �������
    objNewMember.������� = �������
    objNewMember.ˮƽ��ת = ˮƽ��ת
    
    '�޲���ʱ�Զ�ΪNothing
    Set objNewMember.ͼƬ = ͼƬ
    
    '�޲���ʱ�Զ�ΪNew
    If Not CopyIDs Is Nothing Then Set objNewMember.CopyIDs = CopyIDs
    If Not SubIDs Is Nothing Then Set objNewMember.SubIDs = SubIDs
    If Not Relations Is Nothing Then Set objNewMember.Relations = Relations
    If Not ColProtertys Is Nothing Then Set objNewMember.ColProtertys = ColProtertys
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RPTItem
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    On Error GoTo 0
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