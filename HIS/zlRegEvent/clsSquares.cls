VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSquares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsSquare"
Attribute VB_Ext_KEY = "Member0" ,"clsSquare"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'�ֲ����������漯��
Private mCol As Collection
Public Function AddItem(NO As String, �������� As Integer, ���㷽ʽ As String, ������ As Double, _
    ������� As String, ��ע As String, �����޸� As Boolean, ���� As Double, _
    ˢ�����ID As Long, ˢ������ As String, ������ˮ�� As String, ����˵�� As String, _
    ����������Ϣ As String, Optional sKey As String) As clsSquare
    '�����¶���
    Dim objNewMember As clsSquare
    Set objNewMember = New clsSquare
    
    '���ô��뷽��������
    objNewMember.NO = NO
    objNewMember.�������� = ��������
    objNewMember.���㷽ʽ = ���㷽ʽ
    objNewMember.������ = ������
    objNewMember.�������� = ��������
    objNewMember.������� = �������
    objNewMember.��ע = ��ע
    objNewMember.�����޸� = �����޸�
    objNewMember.���� = ����
    objNewMember.ˢ�����ID = ˢ�����ID
    objNewMember.ˢ������ = ˢ������
    objNewMember.������ˮ�� = ������ˮ��
    objNewMember.����˵�� = ����˵��
    objNewMember.����������Ϣ = ����������Ϣ
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    '�����Ѵ����Ķ���
    Set AddItem = objNewMember
    Set objNewMember = Nothing
End Function
Public Function Add(objSquare As clsSquare, Optional sKey As String) As clsSquare
    If Len(sKey) = 0 Then
       mCol.Add objSquare
    Else
        mCol.Add objSquare, sKey
    End If
    Set Add = objSquare
End Function

Public Property Get Item(vntIndexKey As Variant) As clsSquare
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
