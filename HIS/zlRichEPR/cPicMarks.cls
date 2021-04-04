VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cPicMark"
Attribute VB_Ext_KEY = "Member0" ,"cPicMark"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Private mlngKey As Long

Public Property Get Item(vntIndexKey As Variant) As cPicMark
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

Private Sub Class_Initialize()
    Set mCol = New Collection
    mlngKey = 0
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'################################################################################################################
'## 功能：  返回最后添加到集合的 cPicMark 对象
'################################################################################################################
Public Function LastPicMark() As cPicMark
    Set LastPicMark = mCol(mCol.Count)
End Function

'################################################################################################################
'## 功能：  创建新对象
'##
'## 参数：  lngKey :用于恢复旧的集合对象，操作完后需要更新最大流水号。
'################################################################################################################
Public Function Add(Optional lngKey As Long = 0) As Long
    Dim objNewMember As cPicMark
    Set objNewMember = New cPicMark
    Dim sKey As String

    mlngKey = mlngKey + 1
    sKey = "K" & mlngKey
    mCol.Add objNewMember, sKey
    mCol(mCol.Count).Key = "K" & mlngKey        '保存自己在父集合中的关键字！用于删除节点！
    
    Add = mlngKey
    Set objNewMember = Nothing
End Function

'################################################################################################################
'## 功能：  复制集合
'##
'## 参数：  blnWithKeys :是否连同Key值一起复制
'################################################################################################################
Public Function Clone(Optional blnWithKeys As Boolean = False) As cPicMarks
    Dim cNodes As New cPicMarks
    Dim i As Long
    For i = 1 To mCol.Count
        cNodes.Add
        If blnWithKeys Then cNodes.LastPicMark.Key = mCol(i).Key
        cNodes.LastPicMark.Selected = mCol(i).Selected
        cNodes.LastPicMark.X1 = mCol(i).X1
        cNodes.LastPicMark.Y1 = mCol(i).Y1
        cNodes.LastPicMark.X2 = mCol(i).X2
        cNodes.LastPicMark.Y2 = mCol(i).Y2
        cNodes.LastPicMark.点集 = mCol(i).点集
        cNodes.LastPicMark.类型 = mCol(i).类型
        cNodes.LastPicMark.内容 = mCol(i).内容
        cNodes.LastPicMark.填充方式 = mCol(i).填充方式
        cNodes.LastPicMark.填充色 = mCol(i).填充色
        cNodes.LastPicMark.线宽 = mCol(i).线宽
        cNodes.LastPicMark.线条色 = mCol(i).线条色
        cNodes.LastPicMark.线型 = mCol(i).线型
        cNodes.LastPicMark.字体 = mCol(i).字体
        cNodes.LastPicMark.字体色 = mCol(i).字体色
    Next i
    Set Clone = cNodes
End Function

'################################################################################################################
'## 功能：  获取标记的属性字符串
'##
'## 说明：  采用“||”进行连接
'################################################################################################################
Public Property Get 对象属性() As String
    Dim strR As String, i As Long
    For i = 1 To mCol.Count
        If i = 1 Then
            strR = strR & mCol(i).对象属性
        Else
            strR = strR & "||" & mCol(i).对象属性
        End If
    Next
    对象属性 = strR
End Property

'################################################################################################################
'## 功能：  通过一个字符串来设置标记集合内容
'##
'## 说明：  采用“||”进行连接
'################################################################################################################
Public Property Let 对象属性(vData As String)
    Dim T As Variant, i As Long
    T = Split(vData, "||")
    mlngKey = 0
    Set mCol = New Collection
    For i = 0 To UBound(T)
        Call Add
        Me.LastPicMark.对象属性 = T(i)
    Next
End Property

