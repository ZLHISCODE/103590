VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatiInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
Private mvar病人ID As Long '局部复制
Private mvar主页ID As Long '局部复制
Private mvar姓名 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar年龄 As String '局部复制
Private mvarobjCard As Card '局部复制
Private mvar预交余额 As Double '局部复制
Private mvar费用余额 As Double '局部复制
Private mvar预交剩余合计 As Double '局部复制
Private mvar实际余额 As Double '局部复制
Private mvar未付累计 As Double '局部复制
Private mvar结帐类型 As Integer '局部复制
Private mvarbln退款 As Boolean '局部复制
Private mvarbln出院 As Boolean '局部复制
Private mvarbln连续结帐 As Boolean '局部复制
Private mvar门诊号 As String  '局部复制
Private mvar住院号 As String  '局部复制
Private mvar医疗付款方式 As String  '局部复制
Private mvar险类 As String '局部复制

Public Property Let 险类(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类 = 5
    mvar险类 = vData
End Property


Public Property Get 险类() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类
     险类 = mvar险类
End Property

Public Property Let bln连续结帐(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bln连续结帐 = 5
    mvarbln连续结帐 = vData
End Property


Public Property Get bln连续结帐() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bln连续结帐
    bln连续结帐 = mvarbln连续结帐
End Property



Public Property Let bln出院(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bln出院 = 5
    mvarbln出院 = vData
End Property


Public Property Get bln出院() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bln出院
    bln出院 = mvarbln出院
End Property



Public Property Let bln退款(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bln退款 = 5
    mvarbln退款 = vData
End Property


Public Property Get bln退款() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bln退款
    bln退款 = mvarbln退款
End Property



Public Property Let 结帐类型(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐类型 = 5
    mvar结帐类型 = vData
End Property


Public Property Get 结帐类型() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐类型
    结帐类型 = mvar结帐类型
End Property



Public Property Let 未付累计(ByVal vData As Double)
Attribute 未付累计.VB_Description = "上次未付累计金额"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.未付累计 = 5
    mvar未付累计 = vData
End Property


Public Property Get 未付累计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.未付累计
      未付累计 = mvar未付累计
End Property



Public Property Let 实际余额(ByVal vData As Double)
Attribute 实际余额.VB_Description = "本次预交明细中的实际余额"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.实际余额 = 5
    mvar实际余额 = vData
End Property


Public Property Get 实际余额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.实际余额
    实际余额 = mvar实际余额
End Property



Public Property Let 预交剩余合计(ByVal vData As Double)
Attribute 预交剩余合计.VB_Description = "预交余额-未结费用"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预交剩余合计 = 5
    mvar预交剩余合计 = vData
End Property


Public Property Get 预交剩余合计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预交剩余合计
    预交剩余合计 = mvar预交剩余合计
End Property



Public Property Let 费用余额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.费用余额 = 5
    mvar费用余额 = vData
End Property


Public Property Get 费用余额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.费用余额
    费用余额 = mvar费用余额
End Property



Public Property Let 预交余额(ByVal vData As Double)
Attribute 预交余额.VB_Description = "预交余额"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预交余额 = 5
    mvar预交余额 = vData
End Property


Public Property Get 预交余额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预交余额
    预交余额 = mvar预交余额
End Property



Public Property Set objCard(ByVal vData As Card)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objCard = Form1
    Set mvarobjCard = vData
End Property


Public Property Get objCard() As Card
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objCard
    Set objCard = mvarobjCard
End Property



Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar年龄 = vData
End Property


Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.年龄
    年龄 = mvar年龄
End Property



Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mvar性别 = vData
End Property


Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.性别
    性别 = mvar性别
End Property



Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
End Property


Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    姓名 = mvar姓名
End Property



Public Property Let 主页ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.主页ID = 5
    mvar主页ID = vData
End Property


Public Property Get 主页ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.主页ID
    主页ID = mvar主页ID
End Property



Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
End Property


Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病人ID
    病人ID = mvar病人ID
End Property


Public Property Let 门诊号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar门诊号 = vData
End Property


Public Property Get 门诊号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    门诊号 = mvar门诊号
End Property

Public Property Let 住院号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar住院号 = vData
End Property


Public Property Get 住院号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    住院号 = mvar住院号
End Property

Public Property Let 医疗付款方式(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar医疗付款方式 = vData
End Property


Public Property Get 医疗付款方式() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    医疗付款方式 = mvar医疗付款方式
End Property

