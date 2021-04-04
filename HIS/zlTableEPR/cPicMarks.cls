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
'## 功能：  创建新对象
'##
'## 参数：  lngKey :用于恢复旧的集合对象，操作完后需要更新最大流水号。
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
'## 功能：  更新最大流水号
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
'## 功能：  复制集合
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cTabPicMarks
    Dim cNodes As New cTabPicMarks, lngKey As Long, i As Long
    
    For i = 1 To mCol.Count
        lngKey = cNodes.Add()
        If blnWithKeys Then cNodes("K" & lngKey).Key = mCol(i).Key
        cNodes("K" & lngKey).选中 = mCol(i).选中
        cNodes("K" & lngKey).X1 = mCol(i).X1
        cNodes("K" & lngKey).Y1 = mCol(i).Y1
        cNodes("K" & lngKey).X2 = mCol(i).X2
        cNodes("K" & lngKey).Y2 = mCol(i).Y2
        cNodes("K" & lngKey).点集 = mCol(i).点集
        cNodes("K" & lngKey).类型 = mCol(i).类型
        cNodes("K" & lngKey).内容 = mCol(i).内容
        cNodes("K" & lngKey).填充方式 = mCol(i).填充方式
        cNodes("K" & lngKey).填充色 = mCol(i).填充色
        cNodes("K" & lngKey).线宽 = mCol(i).线宽
        cNodes("K" & lngKey).线条色 = mCol(i).线条色
        cNodes("K" & lngKey).线型 = mCol(i).线型
        cNodes("K" & lngKey).字体 = mCol(i).字体
        cNodes("K" & lngKey).字体色 = mCol(i).字体色
    Next i
    Set Clone = cNodes
End Function

