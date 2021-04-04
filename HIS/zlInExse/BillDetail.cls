VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Detail"
Option Explicit '要求变量声明

Public Key As String
Private mvar收费细目ID As Long '局部复制
Private mvar序号 As Integer  '局部复制
Private mvar从属父号 As Integer  '局部复制
Private mvar病人性质 As Integer  '局部复制
Private mvar病人ID As Long '局部复制
Private mvar主页ID As Long '局部复制
Private mvar病区ID As Long '局部复制
Private mvar科室ID As Long '局部复制
Private mvar姓名 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar年龄 As String '局部复制
Private mvar住院号 As String '局部复制
Private mvar床号 As String '局部复制
Private mvar担保额 As Double   '局部复制
Private mvar费别 As String '局部复制
Private mvar医疗付款 As String '局部复制
Private mvar收费类别 As String '局部复制
Private mvar计算单位 As String '局部复制
Private mvar发药窗口 As String '局部复制
Private mvar付数 As Integer  '局部复制
Private mvar数次 As Double '局部复制
Private mvar附加标志 As Byte '局部复制
Private mvar就诊卡号 As String '局部复制
Private mvar执行部门ID As Long '局部复制
Private mvarInComes As BillInComes '局部复制
Private mvarDetail As Detail '局部复制
Private mvar保险大类ID As Long '局部复制
Private mvar保险项目否 As Boolean '局部复制
Private mvar保险编码 As String '局部复制
Private mvar摘要 As String '局部复制
'保持属性值的局部变量
Private mvar原始数量 As Long '局部复制
Private mvar原始执行部门ID As Long '局部复制
Private mvar婴儿费 As Long '局部复制
Private mvarTag As Variant '局部复制
Public Property Let Tag(ByVal vData As Variant)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Tag = 5
    mvarTag = vData
End Property
Public Property Set Tag(ByVal vData As Variant)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Tag = Form1
    Set mvarTag = vData
End Property
Public Property Get Tag() As Variant
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Tag
    If IsObject(mvarTag) Then
        Set Tag = mvarTag
    Else
        Tag = mvarTag
    End If
End Property


Public Property Let 婴儿费(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.原始执行部门ID = 5
    mvar婴儿费 = vData
End Property


Public Property Get 婴儿费() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.原始执行部门ID
    婴儿费 = mvar婴儿费
End Property



Public Property Let 原始执行部门ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.原始执行部门ID = 5
    mvar原始执行部门ID = vData
End Property


Public Property Get 原始执行部门ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.原始执行部门ID
    原始执行部门ID = mvar原始执行部门ID
End Property



Public Property Let 原始数量(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.原始数量 = 5
    mvar原始数量 = vData
End Property


Public Property Get 原始数量() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.原始数量
    原始数量 = mvar原始数量
End Property




Public Property Let 保险编码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.保险编码 = 5
    mvar保险编码 = vData
End Property

Public Property Get 保险编码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.保险编码
    保险编码 = mvar保险编码
End Property

Public Property Let 保险项目否(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.就诊卡号 = 5
    mvar保险项目否 = vData
End Property

Public Property Get 保险项目否() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.就诊卡号
    保险项目否 = mvar保险项目否
End Property

Public Property Let 保险大类ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.就诊卡号 = 5
    mvar保险大类ID = vData
End Property

Public Property Get 保险大类ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.就诊卡号
    保险大类ID = mvar保险大类ID
End Property

Public Property Set Detail(ByVal vData As Detail)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Detail = Form1
    Set mvarDetail = vData
End Property

Public Property Get Detail() As Detail
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Detail
    Set Detail = mvarDetail
End Property

Public Property Set InComes(ByVal vData As BillInComes)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.InComes = Form1
    Set mvarInComes = vData
End Property

Public Property Get InComes() As BillInComes
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.InComes
    Set InComes = mvarInComes
End Property

Public Property Let 执行部门ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行部门ID = 5
    mvar执行部门ID = vData
End Property

Public Property Get 执行部门ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.执行部门ID
    执行部门ID = mvar执行部门ID
End Property

Public Property Let 就诊卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.就诊卡号 = 5
    mvar就诊卡号 = vData
End Property

Public Property Get 就诊卡号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.就诊卡号
    就诊卡号 = mvar就诊卡号
End Property

Public Property Let 附加标志(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.附加标志 = 5
    mvar附加标志 = vData
End Property

Public Property Get 附加标志() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.附加标志
    附加标志 = mvar附加标志
End Property

Public Property Let 数次(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数次 = 5
    mvar数次 = vData
End Property

Public Property Get 数次() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.数次
    数次 = mvar数次
End Property

Public Property Let 付数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.付数 = 5
    mvar付数 = vData
End Property

Public Property Get 付数() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.付数
    付数 = mvar付数
End Property

Public Property Let 发药窗口(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发药窗口 = 5
    mvar发药窗口 = vData
End Property

Public Property Get 发药窗口() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.发药窗口
    发药窗口 = mvar发药窗口
End Property

Public Property Let 计算单位(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.计算单位 = 5
    mvar计算单位 = vData
End Property

Public Property Get 计算单位() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.计算单位
    计算单位 = mvar计算单位
End Property

Public Property Let 收费类别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费类别 = 5
    mvar收费类别 = vData
End Property

Public Property Get 收费类别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.收费类别
    收费类别 = mvar收费类别
End Property

Public Property Let 费别(ByVal vData As String)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.费别 = Form1
    mvar费别 = vData
End Property

Public Property Get 费别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.费别
    费别 = mvar费别
End Property

Public Property Let 床号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.床号 = 5
    mvar床号 = vData
End Property

Public Property Get 床号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.床号
    床号 = mvar床号
End Property

Public Property Let 住院号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.住院号 = 5
    mvar住院号 = vData
End Property

Public Property Get 住院号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.住院号
    住院号 = mvar住院号
End Property

Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar年龄 = vData
End Property

Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.年龄
    年龄 = mvar年龄
End Property

Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mvar性别 = vData
End Property

Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.性别
    性别 = mvar性别
End Property

Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
End Property

Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.姓名
    姓名 = mvar姓名
End Property

Public Property Let 主页ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.主页ID = 5
    mvar主页ID = vData
End Property

Public Property Get 主页ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.主页ID
    主页ID = mvar主页ID
End Property

Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
End Property

Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.病人ID
    病人ID = mvar病人ID
End Property

Public Property Let 从属父号(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.从属父号 = 5
    mvar从属父号 = vData
End Property

Public Property Get 从属父号() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.从属父号
    从属父号 = mvar从属父号
End Property

Public Property Let 序号(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property

Public Property Get 序号() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.序号
    序号 = mvar序号
End Property

Public Property Let 收费细目ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费细目ID = 5
    mvar收费细目ID = vData
End Property

Public Property Get 收费细目ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.收费细目ID
    收费细目ID = mvar收费细目ID
End Property

Private Sub Class_Initialize()
    Set Me.InComes = New BillInComes
    Set Me.Detail = New Detail
End Sub

Private Sub Class_Terminate()
    Set Me.InComes = Nothing
    Set Me.Detail = Nothing
End Sub

Public Property Let 病区ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病区ID = vData
End Property

Public Property Get 病区ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.病人ID
    病区ID = mvar病区ID
End Property

Public Property Let 科室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar科室ID = vData
End Property

Public Property Get 科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.病人ID
    科室ID = mvar科室ID
End Property

Public Property Let 担保额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数次 = 5
    mvar担保额 = vData
End Property

Public Property Get 担保额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.数次
    担保额 = mvar担保额
End Property

Public Property Let 医疗付款(ByVal vData As String)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.费别 = Form1
    mvar医疗付款 = vData
End Property

Public Property Get 医疗付款() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.费别
    医疗付款 = mvar医疗付款
End Property

Public Property Let 摘要(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.就诊卡号 = 5
    mvar摘要 = vData
End Property

Public Property Get 摘要() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.就诊卡号
    摘要 = mvar摘要
End Property

Public Property Let 病人性质(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人性质 = 5
    mvar病人性质 = vData
End Property

Public Property Get 病人性质() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.病人性质
    病人性质 = mvar病人性质
End Property
