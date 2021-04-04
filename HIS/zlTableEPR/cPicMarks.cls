VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabPicMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cPicMark"
Attribute VB_Ext_KEY = "Member0" ,"cPicMark"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngKey As Long

Public Property Get Item(vntIndexKey As Variant) As cTabPicMark
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub
Public Sub RemoveAll()
    Set mCol = Nothing
    Set mCol = New Collection
    mlngKey = 0
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    mlngKey = 0
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
'################################################################################################################
'## ���ܣ�  �����¶���
'##
'## ������  lngKey :���ڻָ��ɵļ��϶��󣬲��������Ҫ���������ˮ�š�
'################################################################################################################
Public Function Add(Optional lngKey As Long = 0) As Long
    Dim objNewMember As cTabPicMark
    Set objNewMember = New cTabPicMark
    Dim sKey As String

    If lngKey = 0 Then
        mlngKey = mlngKey + 1
        sKey = "K" & mlngKey
        objNewMember.Key = mlngKey
        mCol.Add objNewMember, sKey
    Else
        objNewMember.Key = lngKey
        mCol.Add objNewMember, "K" & lngKey
        Call UpdateMaxKey
    End If
    
    Add = mlngKey
    Set objNewMember = Nothing
End Function

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
'## ���ܣ�  ���Ƽ���
'##
'## ������  blnWithKeys :�Ƿ���ͬKeyֵһ����
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cTabPicMarks
    Dim cNodes As New cTabPicMarks, lngKey As Long, i As Long
    
    For i = 1 To mCol.Count
        lngKey = cNodes.Add()
        If blnWithKeys Then cNodes("K" & lngKey).Key = mCol(i).Key
        cNodes("K" & lngKey).ѡ�� = mCol(i).ѡ��
        cNodes("K" & lngKey).X1 = mCol(i).X1
        cNodes("K" & lngKey).Y1 = mCol(i).Y1
        cNodes("K" & lngKey).X2 = mCol(i).X2
        cNodes("K" & lngKey).Y2 = mCol(i).Y2
        cNodes("K" & lngKey).�㼯 = mCol(i).�㼯
        cNodes("K" & lngKey).���� = mCol(i).����
        cNodes("K" & lngKey).���� = mCol(i).����
        cNodes("K" & lngKey).��䷽ʽ = mCol(i).��䷽ʽ
        cNodes("K" & lngKey).���ɫ = mCol(i).���ɫ
        cNodes("K" & lngKey).�߿� = mCol(i).�߿�
        cNodes("K" & lngKey).����ɫ = mCol(i).����ɫ
        cNodes("K" & lngKey).���� = mCol(i).����
        cNodes("K" & lngKey).���� = mCol(i).����
        cNodes("K" & lngKey).����ɫ = mCol(i).����ɫ
    Next i
    Set Clone = cNodes
End Function
