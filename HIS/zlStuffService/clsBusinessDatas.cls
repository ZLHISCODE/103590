VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusinessDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBusinessData"
Attribute VB_Ext_KEY = "Member0" ,"clsBusinessData"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'局部变量，保存集合
Private mCol As Collection
Private mvarTag As String
Private mvar列表接点 As String

Public Property Let Tag(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.登记时间
    Tag = mvarTag
End Property
Public Property Let 列表接点(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
    mvar列表接点 = vData
End Property

Public Property Get 列表接点() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.登记时间
    列表接点 = mvar列表接点
End Property


Public Function Add(数据类型 As gDataType, 数据内容 As String, 是否空值 As Boolean, 是否可选 As Boolean, 数据名称 As String, 节点名称 As String, 数据说明 As String, Optional sKey As String) As clsBusinessData
    '创建新对象
    Dim objNewMember As clsBusinessData
    Set objNewMember = New clsBusinessData

    '设置传入方法的属性
    objNewMember.数据类型 = 数据类型
    objNewMember.数据内容 = 数据内容
    objNewMember.数据名称 = 数据名称
    objNewMember.数据说明 = 数据说明
    objNewMember.是否空值 = 是否空值
    objNewMember.是否可选 = 是否可选
    objNewMember.节点名称 = 节点名称
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function
Public Sub AddItem(objBusinessData As clsBusinessData, Optional sKey As String)
   If Len(sKey) = 0 Then
        mCol.Add objBusinessData
    Else
        mCol.Add objBusinessData, sKey
    End If
End Sub
Public Property Get Item(vntIndexKey As Variant) As clsBusinessData
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

