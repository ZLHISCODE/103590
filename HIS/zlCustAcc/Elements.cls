VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Elements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Element"
Attribute VB_Ext_KEY = "Member0" ,"Element"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection

Public Function Add(Key As String, Control As Object, Optional ByVal TabIndex As Long, Optional ByVal Visible As Boolean = True) As Element
    '�����¶���
    Dim objNewMember As Element
    Set objNewMember = New Element


    '���ô��뷽��������
    objNewMember.Key = Key
    Control.Tag = Key
    Set objNewMember.Control = Control
    objNewMember.Visible = Visible
    '������������˫����ʡ������ֶδӼ������ܵõ��ؼ����ӿؼ�Ҳ�ܷ��ʶ�Ӧ�ļ���Ԫ��
    If TabIndex = 0 Then
        objNewMember.TabIndex = mCol.Count + 1 'ʹTabIndexֵ��Ϊ��
    Else
        objNewMember.TabIndex = TabIndex
    End If
    mCol.Add objNewMember, Key
    
    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Element
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
    Dim objNewMember As Element
    Dim lngTabIndex As Long
    
    
    lngTabIndex = mCol(vntIndexKey).TabIndex
    
    For Each objNewMember In mCol
        If objNewMember.TabIndex > lngTabIndex Then
            objNewMember.TabIndex = objNewMember.TabIndex - 1
        End If
    Next
    
    mCol.Remove vntIndexKey
End Sub

Public Sub Clear()
    'ɾ�������е�����Ԫ��
    
    Dim lngCount As Long
    For lngCount = 1 To mCol.Count
        mCol.Remove 1
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
