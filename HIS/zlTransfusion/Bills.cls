VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Bill"
Attribute VB_Ext_KEY = "Member0" ,"Bill"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(����ID As Long, ҽ��ID As Long, ҽ������ As String, ���� As Currency, ��λ As String, ���� As Long, ʱ�� As Long, ҽ������ As String, ��� As Integer, ��� As Currency, ��ϸ�Ʒ�״̬ As Integer, Optional sKey As String) As Bill
    '�����¶���
    Dim objNewMember As Bill
    Set objNewMember = New Bill
    
    On Error GoTo errH
    '���ô��뷽��������
    objNewMember.����ID = ����ID
    objNewMember.ҽ��ID = ҽ��ID
    objNewMember.ҽ������ = DelInvalidChar(ҽ������, "'")
    objNewMember.���� = ����
    objNewMember.��λ = ��λ
    objNewMember.���� = ����
    objNewMember.ʱ�� = ʱ��
    objNewMember.ҽ������ = DelInvalidChar(ҽ������, "'")
    objNewMember.��� = ���
    objNewMember.��� = ���
    objNewMember.��ϸ�Ʒ�״̬ = ��ϸ�Ʒ�״̬
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
errH:
    If Err.Number = 457 Then
        Exit Function
    ElseIf ErrCenter = 1 Then
        Resume
    End If

End Function

Public Property Get Item(vntIndexKey As Variant) As Bill
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    ' x.Count
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
