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

'局部变量，保存集合
Private mCol As Collection

Public Function Add(Key As String, Control As Object, Optional ByVal TabIndex As Long, Optional ByVal Visible As Boolean = True) As Element
    '创建新对象
    Dim objNewMember As Element
    Set objNewMember = New Element


    '设置传入方法的属性
    objNewMember.Key = Key
    Control.Tag = Key
    Set objNewMember.Control = Control
    objNewMember.Visible = Visible
    '这样才能做到双向访问。即按字段从集合中能得到控件，从控件也能访问对应的集合元素
    If TabIndex = 0 Then
        objNewMember.TabIndex = mCol.Count + 1 '使TabIndex值不为零
    Else
        objNewMember.TabIndex = TabIndex
    End If
    mCol.Add objNewMember, Key
    
    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Element
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
    '删除集合中的所有元素
    
    Dim lngCount As Long
    For lngCount = 1 To mCol.Count
        mCol.Remove 1
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

