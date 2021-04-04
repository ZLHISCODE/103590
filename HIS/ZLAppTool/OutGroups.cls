VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"OutGroup"
Attribute VB_Ext_KEY = "Member0" ,"OutGroup"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection
Event AddGroup(Key As String, Caption As String)
Event RemoveGroup(vntIndexKey As Variant)

Public Function Add(Optional Key As String, Optional Caption As String) As OutGroup
    '�����¶���
    Dim objNewMember As OutGroup
    Set objNewMember = New OutGroup


    '���ô��뷽��������
    objNewMember.Caption = Caption
    objNewMember.Key = Key
    'objNewMember.
    On Error Resume Next
    If Len(Key) = 0 Then
        objNewMember.Key = "K" & Caption
        mCol.Add objNewMember
    Else
        objNewMember.Key = Key
        mCol.Add objNewMember, Key
    End If
    If Err <> 0 Then Exit Function


    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing
    
    RaiseEvent AddGroup(Key, Caption)
End Function

Public Property Get item(vntIndexKey As Variant) As OutGroup
Attribute item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)

    RaiseEvent RemoveGroup(vntIndexKey)
    mCol.Remove vntIndexKey
End Sub

Public Sub Clear()
    Dim lngIndex As Long
    For lngIndex = 1 To mCol.Count
        Remove 1
    Next
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
