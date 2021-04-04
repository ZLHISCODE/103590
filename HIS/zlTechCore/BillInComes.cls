VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillInComes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"InCome"
Attribute VB_Ext_KEY = "Member0" ,"InCome"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit '要求变量声明

'局部变量，保存集合
Private mCol As Collection

Public Function Add(收入项目ID As Long, 收入项目 As String, 收据费目 As String, 标准单价 As Double, _
    应收金额 As Currency, 实收金额 As Currency, Optional 原价 As Double, Optional 现价 As Double, _
    Optional sKey As String, Optional 统筹金额 As Currency) As BillInCome
    '创建新对象
    Dim objNewMember As BillInCome
    Set objNewMember = New BillInCome

    '设置传入方法的属性
    objNewMember.Key = sKey
    objNewMember.收入项目ID = 收入项目ID
    objNewMember.收入项目 = 收入项目
    objNewMember.收据费目 = 收据费目
    objNewMember.标准单价 = 标准单价
    objNewMember.原价 = 原价
    objNewMember.现价 = 现价
    objNewMember.应收金额 = 应收金额
    objNewMember.实收金额 = 实收金额
    objNewMember.统筹金额 = 统筹金额
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember ', sKey'因为用来存放数据了,所以不用作Key
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As BillInCome
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：'debug.print x.Count
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
