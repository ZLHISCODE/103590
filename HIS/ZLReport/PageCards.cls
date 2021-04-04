VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'局部变量，保存集合
Private mCol As Collection

Public Function Add(id As Integer, X As Long, Y As Long, W As Long, H As Long, RowB As Long, RowE As Long, Row As Long, Col As Long, Item As Collection, Optional sKey As String) As PageCard
    '创建新对象
    Dim objNewMember As PageCard
    Set objNewMember = New PageCard

    '设置传入方法的属性
    objNewMember.Key = Mid(sKey, 2) 'Key为该问题关键字,应与ID对应
    objNewMember.id = id '数据表格控件索引
    
    '该表格在该页内输出的整体位置、尺寸(含表头)
    objNewMember.X = X
    objNewMember.Y = Y
    objNewMember.W = W
    objNewMember.H = H '数据输出可用高度(含表头),如果=FixH,表示不输出表头
  
    objNewMember.RowB = RowB '表格在该页的开始行(非固定行)
    objNewMember.RowE = RowE '表格在该页的结束行(非固定行)
    objNewMember.Row = Row
    objNewMember.Col = Col
    objNewMember.Item = Item
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As PageCard
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
     On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Items() As Collection
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Items = mCol
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

