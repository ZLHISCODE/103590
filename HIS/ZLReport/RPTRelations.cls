VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTRelations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function Add(关联报表ID As Long, 参数名 As String, 参数值来源 As String, 关联报表名称 As String, 默认 As Integer) As RPTRelation
'KEy="_" & 参数名
    '创建新对象
    Dim objNewMember As RPTRelation
    Set objNewMember = New RPTRelation

    '设置传入方法的属性
    objNewMember.Key = 关联报表ID & "_" & 参数名
    objNewMember.关联报表ID = 关联报表ID
    objNewMember.参数名 = 参数名
    objNewMember.参数值来源 = 参数值来源
    objNewMember.关联报表名称 = 关联报表名称
    objNewMember.默认 = 默认
    
    mCol.Add objNewMember, objNewMember.Key

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RPTRelation
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    If Err.Number > 0 Then Set Item = New RPTRelation
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



