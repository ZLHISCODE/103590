VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MapItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MapItem"
Attribute VB_Ext_KEY = "Member0" ,"MapItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function Add(类型 As Byte, 内容 As String, 字体 As String, 点集 As String, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, 填充色 As Long, 填充方式 As Integer, 线条色 As Long, 线型 As Byte, 线宽 As Byte, Optional sKey As String, Optional Oper As Byte) As MapItem
    '创建新对象
    Dim objNewMember As MapItem
    Set objNewMember = New MapItem

    '设置传入方法的属性
    objNewMember.类型 = 类型
    
    objNewMember.内容 = 内容
    objNewMember.字体 = 字体
    
    objNewMember.点集 = 点集
    objNewMember.X1 = X1
    objNewMember.Y1 = Y1
    objNewMember.X2 = X2
    objNewMember.Y2 = Y2
    
    objNewMember.填充色 = 填充色
    objNewMember.填充方式 = 填充方式
    objNewMember.线条色 = 线条色
    objNewMember.线宽 = 线宽
    objNewMember.线型 = 线型
    
    objNewMember.Oper = Oper
    objNewMember.Key = sKey
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As MapItem
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub
