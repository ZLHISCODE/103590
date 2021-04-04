VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusinessAsks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBusinessAsk"
Attribute VB_Ext_KEY = "Member0" ,"clsBusinessAsk"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'局部变量，保存集合
Private mCol As Collection

Public Function Add(发布ID As String, 过程名 As String, 功能 As String, Optional sKey As String) As clsBusinessAsk
    '创建新对象
    Dim objNewMember As clsBusinessAsk
    Set objNewMember = New clsBusinessAsk


    '设置传入方法的属性
    objNewMember.发布ID = 发布ID
    objNewMember.过程名 = 过程名
    objNewMember.功能 = 功能
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function
Public Sub AddItem(objBusinessAsk As clsBusinessAsk, Optional sKey As String)
    '创建新对象
    If Len(sKey) = 0 Then
        mCol.Add objBusinessAsk
    Else
        mCol.Add objBusinessAsk, sKey
    End If
      
End Sub




Public Property Get Item(vntIndexKey As Variant) As clsBusinessAsk
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

