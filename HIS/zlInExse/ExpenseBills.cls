VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpenseBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"ExpenseBill"
Attribute VB_Ext_KEY = "Member0" ,"ExpenseBill"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit '要求变量声明

'局部变量，保存集合
Private mCol As Collection

Public Function Add(NO As String, 病人ID As Long, 主页ID As Long, 病区ID As Long, 科室ID As Long, _
    姓名 As String, 性别 As String, 年龄 As String, 标识号 As String, 床号 As String, 费别 As String, _
    门诊标志 As Byte, 加班标志 As Byte, 婴儿费 As Byte, 多病人单 As Boolean, 开单部门ID As Long, _
    开单人 As String, 划价人 As String, 操作员姓名 As String, 操作员编号 As String, _
    发生时间 As Date, 登记时间 As Date, Details As BillDetails, Optional sKey As String) As ExpenseBill
    '创建新对象
    Dim objNewMember As ExpenseBill
    Set objNewMember = New ExpenseBill

    '设置传入方法的属性
    objNewMember.Key = sKey
    objNewMember.NO = NO
    objNewMember.病人ID = 病人ID
    objNewMember.主页ID = 主页ID
    objNewMember.病区ID = 病区ID
    objNewMember.科室ID = 科室ID
    objNewMember.姓名 = 姓名
    objNewMember.性别 = 性别
    objNewMember.年龄 = 年龄
    objNewMember.标识号 = 标识号
    objNewMember.床号 = 床号
    objNewMember.费别 = 费别
    objNewMember.门诊标志 = 门诊标志
    objNewMember.加班标志 = 加班标志
    objNewMember.婴儿费 = 婴儿费
    objNewMember.多病人单 = 多病人单
    objNewMember.开单部门ID = 开单部门ID
    objNewMember.开单人 = 开单人
    objNewMember.划价人 = 划价人
    objNewMember.操作员姓名 = 操作员姓名
    objNewMember.操作员编号 = 操作员编号
    objNewMember.发生时间 = 发生时间
    objNewMember.登记时间 = 登记时间
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

Public Property Get Item(vntIndexKey As Variant) As ExpenseBill
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
