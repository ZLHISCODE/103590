VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTPars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"RPTPar"
Attribute VB_Ext_KEY = "Member0" ,"RPTPar"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function Add(组名 As String, 序号 As Byte, 名称 As String, 类型 As Byte, 缺省值 As String, 格式 As Byte, 值列表 As String, 分类SQL As String, 明细SQL As String, 分类字段 As String, 明细字段 As String, 对象 As String, Optional sKey As String, Optional Reserve As String, Optional 是否锁定 As Boolean) As RPTPar
    '创建新对象
    Dim objNewMember As RPTPar
    Set objNewMember = New RPTPar

    '设置传入方法的属性
    objNewMember.Key = Mid(sKey, 2)
    objNewMember.序号 = 序号
    objNewMember.名称 = 名称
    objNewMember.类型 = 类型
    objNewMember.缺省值 = 缺省值
    objNewMember.格式 = 格式
    objNewMember.值列表 = 值列表
    objNewMember.分类SQL = 分类SQL
    objNewMember.分类字段 = 分类字段
    objNewMember.明细SQL = 明细SQL
    objNewMember.明细字段 = 明细字段
    objNewMember.对象 = 对象
    objNewMember.组名 = 组名
    objNewMember.Reserve = Reserve
    objNewMember.是否锁定 = 是否锁定
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RPTPar
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
