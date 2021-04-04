VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "停诊记录集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"停诊记录"
Attribute VB_Ext_KEY = "Member0" ,"停诊记录"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function Add(开始时间 As String, 终止时间 As String, 停诊原因 As String, Optional sKey As String) As 停诊记录
    '创建新对象
    Dim objNewMember As 停诊记录
    Set objNewMember = New 停诊记录


    '设置传入方法的属性
    objNewMember.开始时间 = 开始时间
    objNewMember.终止时间 = 终止时间
    objNewMember.停诊原因 = 停诊原因
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As 停诊记录
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

Public Sub AddItem(obj停诊记录 As 停诊记录, Optional sKey As String)
    If obj停诊记录 Is Nothing Then Set obj停诊记录 = New 停诊记录
    If Len(sKey) = 0 Then
        mCol.Add obj停诊记录
    Else
        mCol.Add obj停诊记录, sKey
    End If
End Sub

Public Function Clone() As 停诊记录集
    '创建一个副本
    Dim obj停诊记录集 As New 停诊记录集
    Dim objItem As 停诊记录
    
    With obj停诊记录集
        For Each objItem In mCol
            .AddItem objItem.Clone
        Next
    End With
    Set Clone = obj停诊记录集
End Function

Public Sub RemoveAll()
    '删除集合中的所有元素时使用。
    Dim i As Integer
    
    For i = 1 To mCol.Count
        If i > mCol.Count Then Exit For
        mCol.Remove i
        i = i - 1
    Next
End Sub
