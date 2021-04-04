VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"BillDetail"
Attribute VB_Ext_KEY = "Member0" ,"BillDetail"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit '要求变量声明

'局部变量，保存集合
Private mCol As Collection

Public Function Add(费别 As String, Detail As Detail, 收费细目ID As Long, 序号 As Integer, _
    从属父号 As Integer, 收费类别 As String, 计算单位 As String, 发药窗口 As String, 付数 As Integer, _
    数次 As Double, 附加标志 As Byte, 执行部门ID As Long, InComes As BillInComes, _
    Optional sKey As String, Optional 保险项目否 As Boolean, Optional 保险大类ID As Long, _
    Optional 保险编码 As String, Optional 摘要 As String, Optional 原始数量 As Long, Optional 原始执行部门ID As Long) As BillDetail
    '创建新对象
    Dim objNewMember As BillDetail
    Set objNewMember = New BillDetail


    '设置传入方法的属性
    objNewMember.Key = sKey
    Set objNewMember.Detail = Detail
    objNewMember.收费细目ID = 收费细目ID
    objNewMember.序号 = 序号
    objNewMember.从属父号 = 从属父号
    objNewMember.费别 = 费别
    objNewMember.收费类别 = 收费类别
    objNewMember.计算单位 = 计算单位
    objNewMember.发药窗口 = 发药窗口
    objNewMember.付数 = 付数
    objNewMember.数次 = 数次
    objNewMember.附加标志 = 附加标志
    objNewMember.执行部门ID = 执行部门ID
    objNewMember.保险项目否 = 保险项目否
    objNewMember.保险大类ID = 保险大类ID
    objNewMember.保险编码 = 保险编码
    objNewMember.摘要 = 摘要
    objNewMember.原始数量 = 原始数量
    objNewMember.原始执行部门ID = 原始执行部门ID
    
    Set objNewMember.InComes = InComes
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As BillDetail
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
