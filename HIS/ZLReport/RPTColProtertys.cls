VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTColProtertys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function Add(条件名称 As String, 条件字段 As String, 条件关系 As String, 条件值 As String, _
    字体颜色 As Long, 背景颜色 As Long, 是否加粗 As Boolean, 是否整行应用 As Boolean, _
    对齐 As Integer, Optional sKey As String) As RPTColProterty
            
'创建新对象
    Dim objNewMember As RPTColProterty
    Set objNewMember = New RPTColProterty

    '设置传入方法的属性
    objNewMember.Key = Mid(sKey, 2)
    objNewMember.条件名称 = 条件名称
    objNewMember.条件字段 = 条件字段
    objNewMember.条件关系 = 条件关系
    objNewMember.条件值 = 条件值
    objNewMember.字体颜色 = 字体颜色
    objNewMember.背景颜色 = 背景颜色
    objNewMember.是否加粗 = 是否加粗
    objNewMember.是否整行应用 = 是否整行应用
    objNewMember.对齐 = 对齐
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RPTColProterty
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    count = mCol.count
End Property

Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
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

