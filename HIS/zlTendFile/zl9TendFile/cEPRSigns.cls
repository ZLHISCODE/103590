VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRSigns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cEPRSign"
Attribute VB_Ext_KEY = "Member0" ,"cEPRSign"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngKey As Long

Public Property Get Item(vntIndexKey As Variant) As cEPRSign
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_initialize()
    Set mCol = New Collection
    mlngKey = 0
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    mlngKey = 0
End Sub

'################################################################################################################
'## ���ܣ�  ����һ���Ѿ����ڵ�ͼƬ���󵽼�����
'##
'## ������  blnWithKeys :��ʾ�Ƿ�ͬʱ����Keyֵ
'################################################################################################################
Public Function AddExistNode(vData As cEPRSign, Optional blnWithKeys As Boolean = False) As Long
    Dim strKey As String
    If blnWithKeys = False Then
        mlngKey = mlngKey + 1
        strKey = "K" & mlngKey
    Else
        strKey = "K" & vData.Key
    End If
    
    mCol.Add vData.Clone(blnWithKeys), strKey
    
    If blnWithKeys = False Then
        mCol(strKey).Key = mlngKey
        AddExistNode = mlngKey
    Else
        AddExistNode = vData.Key
        Call UpdateMaxKey
    End If
End Function

'################################################################################################################
'## ���ܣ�  �����¶���
'##
'## ������  lngKey :���ڻָ��ɵļ��϶��󣬲��������Ҫ���������ˮ�š�
'################################################################################################################
Public Function Add(Optional lngKey As Long = 0) As Long
    Dim objNewMember As cEPRSign
    Set objNewMember = New cEPRSign
    Dim sKey As String
    
    If lngKey = 0 Then
        mlngKey = mlngKey + 1
        sKey = "K" & mlngKey
        objNewMember.Key = mlngKey
        mCol.Add objNewMember, sKey
        Add = mlngKey
    Else
        objNewMember.Key = lngKey
        mCol.Add objNewMember, "K" & lngKey
        Call UpdateMaxKey
        Add = lngKey
    End If
    Set objNewMember = Nothing
End Function

'################################################################################################################
'## ���ܣ�  ������ж����ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub CheckValidIDs(eEditType As EditTypeEnum)
    Dim i As Long
    Select Case eEditType
    Case cprET_�����ļ�����
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("�����ļ��ṹ")
        Next
    Case cprET_ȫ��ʾ���༭
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("������������")
        Next
    Case cprET_�������༭, cprET_���������
        For i = 1 To mCol.Count
            If mCol(i).ID = 0 Then mCol(i).ID = zlDatabase.GetNextId("���Ӳ�������")
        Next
    End Select
End Sub

'################################################################################################################
'## ���ܣ�  ���������ˮ��
'################################################################################################################
Public Sub UpdateMaxKey()
    Dim i As Long, j As Long
    j = 0
    For i = 1 To mCol.Count
        If mCol(i).Key > j Then j = mCol(i).Key
    Next
    mlngKey = j
End Sub

'################################################################################################################
'## ���ܣ�  ����ͼƬ����
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cEPRSigns
    '���ƶ���
    Dim cNodes As New cEPRSigns, lngKey As Long
    Dim i As Long
    For i = 1 To mCol.Count
        lngKey = cNodes.Add()
        If blnWithKeys Then
            cNodes("K" & lngKey).Key = mCol(i).Key
        End If
        cNodes("K" & lngKey).ID = mCol(i).ID
        cNodes("K" & lngKey).�ļ�ID = mCol(i).�ļ�ID
        cNodes("K" & lngKey).��ID = mCol(i).��ID
        cNodes("K" & lngKey).��ʼ�� = mCol(i).��ʼ��
        cNodes("K" & lngKey).��ֹ�� = mCol(i).��ֹ��
        cNodes("K" & lngKey).������� = mCol(i).�������
        cNodes("K" & lngKey).���� = mCol(i).����
        cNodes("K" & lngKey).ǩ����Ϣ = mCol(i).ǩ����Ϣ
        cNodes("K" & lngKey).ǰ������ = mCol(i).ǰ������
        cNodes("K" & lngKey).ǩ������ = mCol(i).ǩ������
        cNodes("K" & lngKey).�������� = mCol(i).��������
        cNodes("K" & lngKey).ʱ��� = mCol(i).ʱ���
    Next i
    If blnWithKeys Then Call UpdateMaxKey
    Set Clone = cNodes
End Function

'################################################################################################################
'## ���ܣ�  ��ȡָ���汾�����ؼ��֣�����ָ���汾�����һ��ǩ����
'################################################################################################################
Public Function GetMaxKeyByVersion(lngVersion As Long) As Long
    Dim i As Long, j As Long
    For i = 1 To mCol.Count
        If mCol(i).Key > j And mCol(i).��ʼ�� = lngVersion Then j = mCol(i).Key
    Next
    GetMaxKeyByVersion = j
End Function

'################################################################################################################
'## ���ܣ�  ��ȡ���ؼ��֣��������һ��ǩ����
'################################################################################################################
Public Function GetMaxKey() As Long
    Dim i As Long, j As Long
    For i = 1 To mCol.Count
        If mCol(i).Key > j Then j = mCol(i).Key
    Next
    GetMaxKey = j
End Function
