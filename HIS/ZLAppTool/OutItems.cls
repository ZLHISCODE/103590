VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OutItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"OutItem"
Attribute VB_Ext_KEY = "Member0" ,"OutItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'局部变量，保存集合
Private mCol As Collection
Event AddItem(Key As String, Caption As String, Icon As String, GroupName As String)
Event RemoveItem(vntIndexKey As Variant)

Public Function Add(Optional Key As String, Optional Caption As String, Optional Icon As String, Optional GroupName As String) As OutItem
    '创建新对象
    Dim objNewMember As OutItem
    Set objNewMember = New OutItem


    '设置传入方法的属性
    objNewMember.Caption = Caption
    objNewMember.Icon = Icon
    objNewMember.GroupName = GroupName
    On Error Resume Next
    If Len(Key) = 0 Then
        objNewMember.Key = "K" & Caption
        mCol.Add objNewMember
    Else
        objNewMember.Key = Key
        mCol.Add objNewMember, Key
    End If
    If Err <> 0 Then Exit Function

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
    RaiseEvent AddItem(Key, Caption, Icon, GroupName)
End Function

Public Property Get item(vntIndexKey As Variant) As OutItem
Attribute item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set item = mCol(vntIndexKey)
End Property


Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)

    RaiseEvent RemoveItem(vntIndexKey)
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

