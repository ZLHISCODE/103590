VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillPages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"BillPage"
Attribute VB_Ext_KEY = "Member0" ,"BillPage"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit '要求变量声明

'局部变量，保存集合
Private mCol As Collection

Public Function Add(Details As BillDetails, Optional NO As String, _
    Optional 开单部门ID As Long, Optional 开单人 As String, _
    Optional 应收金额 As Double, Optional 实收金额 As Double, _
    Optional 保险金额 As Double, Optional 冲预交额 As Double, _
    Optional 应缴金额 As Double, Optional 误差金额 As Double, Optional 收费结算 As String, _
    Optional 全自付 As Double, Optional 先自付 As Double, Optional 进入统筹 As Double, _
    Optional sKey As String, Optional str消费卡结算 As String, _
    Optional cur消费卡刷卡额 As Double, Optional 结帐ID As Long) As BillPage
    '创建新对象
    Dim objNewMember As BillPage
    Set objNewMember = New BillPage
    '设置传入方法的属性
    objNewMember.Key = sKey
    objNewMember.NO = NO
    objNewMember.开单部门ID = 开单部门ID
    objNewMember.开单人 = 开单人
    objNewMember.应收金额 = 应收金额
    objNewMember.实收金额 = 实收金额
    objNewMember.保险金额 = 保险金额
    objNewMember.冲预交额 = 冲预交额
    objNewMember.应缴金额 = 应缴金额
    objNewMember.误差金额 = 误差金额
    objNewMember.收费结算 = 收费结算
    objNewMember.全自付 = 全自付
    objNewMember.先自付 = 先自付
    objNewMember.进入统筹 = 进入统筹
    objNewMember.消费卡结算 = str消费卡结算
    objNewMember.消费卡刷卡额 = cur消费卡刷卡额
    objNewMember.结帐ID = 结帐ID
    Set objNewMember.Details = Details
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As BillPage
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
