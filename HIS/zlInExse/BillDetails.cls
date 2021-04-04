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

Public Function Add(Detail As Detail, 收费细目ID As Long, 序号 As Integer, 从属父号 As Integer, _
    病人ID As Long, 主页ID As Long, 病区ID As Long, 科室ID As Long, 姓名 As String, 性别 As String, _
    年龄 As String, 住院号 As String, 床号 As String, 费别 As String, 病人性质 As Integer, 收费类别 As String, 计算单位 As String, _
    发药窗口 As String, 付数 As Integer, 数次 As Double, 附加标志 As Byte, 执行部门ID As Long, _
    InComes As BillInComes, Optional 就诊卡号 As String, Optional sKey As String, Optional 担保额 As Double, _
    Optional 医疗付款 As String, Optional 保险项目否 As Boolean, Optional 保险大类ID As Long, _
    Optional 保险编码 As String, Optional 摘要 As String, Optional 原始数量 As Long, Optional 原始执行部门ID As Long, _
    Optional 婴儿费 As Long, Optional varTag As Variant) As BillDetail
    '创建新对象
    Dim objNewMember As BillDetail
    Set objNewMember = New BillDetail


    '设置传入方法的属性
    objNewMember.Key = sKey
    Set objNewMember.Detail = Detail
    objNewMember.收费细目ID = 收费细目ID
    objNewMember.序号 = 序号
    objNewMember.从属父号 = 从属父号
    objNewMember.病人ID = 病人ID
    objNewMember.主页ID = 主页ID
    objNewMember.病区ID = 病区ID
    objNewMember.科室ID = 科室ID
    objNewMember.姓名 = 姓名
    objNewMember.性别 = 性别
    objNewMember.年龄 = 年龄
    objNewMember.住院号 = 住院号
    objNewMember.床号 = 床号
    objNewMember.费别 = 费别
    objNewMember.病人性质 = 病人性质
    objNewMember.医疗付款 = 医疗付款
    objNewMember.收费类别 = 收费类别
    objNewMember.计算单位 = 计算单位
    objNewMember.发药窗口 = 发药窗口
    objNewMember.付数 = 付数
    objNewMember.数次 = 数次
    objNewMember.附加标志 = 附加标志
    objNewMember.就诊卡号 = 就诊卡号
    objNewMember.执行部门ID = 执行部门ID
    objNewMember.担保额 = 担保额
    
    objNewMember.保险项目否 = 保险项目否
    objNewMember.保险大类ID = 保险大类ID
    objNewMember.保险编码 = 保险编码
    
    objNewMember.摘要 = 摘要
    objNewMember.原始数量 = 原始数量
    objNewMember.原始执行部门ID = 原始执行部门ID
    objNewMember.婴儿费 = 婴儿费
    
    If IsObject(varTag) Then
        Set objNewMember.Tag = varTag
    Else
        objNewMember.Tag = varTag
    End If
    
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
