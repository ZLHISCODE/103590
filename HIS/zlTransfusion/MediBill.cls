VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MediBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "寄存药品"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
Private mvar序号 As Integer '局部复制
Private mvar医嘱ID As Long '局部复制
Private mvar发送号 As Long '
Private mvar药品ID As Long '局部复制
Private mvar药品名称 As String '局部复制
Private mvar规格 As String '局部复制
Private mvar执行分类 As Integer '局部复制
Private mvar使用状态 As Integer '局部复制
Private mvar入出系数 As Integer '局部复制
Private mvar计算单位 As String '局部复制
Private mvar门诊单位 As String '局部复制
Private mvar门诊包装 As Double '局部复制
Private mvar剂量单位 As String '局部复制
Private mvar剂量系数 As Double '局部复制
Private mvar容量 As Double '局部复制
Private mvar数量 As Double '局部复制
Private mvar已用数量 As Double
Private mvar单价 As Double '局部复制
Private mvar金额 As Double '局部复制
Private mvar登记时间 As Date
Private mvar使用摘要 As String
Private mvar填制人 As String

Public Property Let 金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.金额 = 5
    mvar金额 = vData
End Property

Public Property Get 金额() As Double
'检索属性值时使用，位于赋值语句的右边。
' X.金额
    金额 = mvar金额
End Property

Public Property Let 单价(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.单价 = 5
    mvar单价 = vData
End Property


Public Property Get 单价() As Double
'检索属性值时使用，位于赋值语句的右边。
' X.单价
    单价 = mvar单价
End Property

Public Property Let 数量(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数量 = 5
    mvar数量 = vData
End Property


Public Property Get 数量() As Double
'检索属性值时使用，位于赋值语句的右边。
' X.数量
    数量 = mvar数量
End Property

Public Property Let 已用数量(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数量 = 5
    mvar已用数量 = vData
End Property


Public Property Get 已用数量() As Double
'检索属性值时使用，位于赋值语句的右边。
' X.数量
    已用数量 = mvar已用数量
End Property

Public Property Let 容量(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.容量 = 5
    mvar容量 = vData
End Property


Public Property Get 容量() As Double
'检索属性值时使用，位于赋值语句的右边。
' X.容量
    容量 = mvar容量
End Property

Public Property Let 剂量系数(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.剂量系数 = 5
    mvar剂量系数 = vData
End Property


Public Property Get 剂量系数() As Double
'检索属性值时使用，位于赋值语句的右边。
' X.剂量系数
    剂量系数 = mvar剂量系数
End Property

Public Property Let 门诊包装(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.剂量系数 = 5
    mvar门诊包装 = vData
End Property

Public Property Get 门诊包装() As Double
'检索属性值时使用，位于赋值语句的右边。
' X.剂量系数
    门诊包装 = mvar门诊包装
End Property

Public Property Let 剂量单位(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.剂量单位 = 5
    mvar剂量单位 = vData
End Property

Public Property Get 剂量单位() As String
'检索属性值时使用，位于赋值语句的右边。
' X.剂量单位
    剂量单位 = mvar剂量单位
End Property

Public Property Let 门诊单位(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊单位 = 5
    mvar门诊单位 = MidUni(vData, 1, 20)
End Property


Public Property Get 门诊单位() As String
'检索属性值时使用，位于赋值语句的右边。
' X.门诊单位
    门诊单位 = mvar门诊单位
End Property

Public Property Let 计算单位(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊单位 = 5
    mvar计算单位 = MidUni(vData, 1, 20)
End Property


Public Property Get 计算单位() As String
'检索属性值时使用，位于赋值语句的右边。
' X.门诊单位
    计算单位 = mvar计算单位
End Property


Public Property Let 入出系数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.入出系数 = 5
    mvar入出系数 = vData
End Property


Public Property Get 入出系数() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.入出系数
    入出系数 = mvar入出系数
End Property



Public Property Let 使用状态(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.使用状态 = 5
    mvar使用状态 = vData
End Property


Public Property Get 使用状态() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.使用状态
    使用状态 = mvar使用状态
End Property



Public Property Let 执行分类(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行分类 = 5
    mvar执行分类 = vData
End Property


Public Property Get 执行分类() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.执行分类
    执行分类 = mvar执行分类
End Property



Public Property Let 规格(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.规格 = 5
    mvar规格 = MidUni(vData, 1, 40)
End Property


Public Property Get 规格() As String
'检索属性值时使用，位于赋值语句的右边。
' X.规格
    规格 = mvar规格
End Property

Public Property Let 药品名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.药品名称 = 5
    mvar药品名称 = MidUni(vData, 1, 80)
End Property


Public Property Get 药品名称() As String
'检索属性值时使用，位于赋值语句的右边。
' X.药品名称
    药品名称 = mvar药品名称
End Property



Public Property Let 药品ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.药品ID = 5
    mvar药品ID = MidUni(vData, 1, 80)
End Property


Public Property Get 药品ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.药品ID
    药品ID = mvar药品ID
End Property



Public Property Let 医嘱ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医嘱ID = 5
    mvar医嘱ID = vData
End Property


Public Property Get 医嘱ID() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.医嘱ID
    医嘱ID = mvar医嘱ID
End Property


Public Property Let 发送号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医嘱ID = 5
    mvar发送号 = vData
End Property


Public Property Get 发送号() As Long
'检索属性值时使用，位于赋值语句的右边。
' X.医嘱ID
    发送号 = mvar发送号
End Property

Public Property Let 序号(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property


Public Property Get 序号() As Integer
'检索属性值时使用，位于赋值语句的右边。
' X.序号
    序号 = mvar序号
End Property

Public Property Let 登记时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar登记时间 = vData
End Property


Public Property Get 登记时间() As Date
'检索属性值时使用，位于赋值语句的右边。
' X.序号
    登记时间 = mvar登记时间
End Property

Public Property Let 使用摘要(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.药品名称 = 5
    mvar使用摘要 = MidUni(vData, 1, 200)
End Property


Public Property Get 使用摘要() As String
'检索属性值时使用，位于赋值语句的右边。
' X.药品名称
    使用摘要 = mvar使用摘要
End Property

Public Property Let 填制人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.药品名称 = 5
    mvar填制人 = vData
End Property


Public Property Get 填制人() As String
'检索属性值时使用，位于赋值语句的右边。
' X.药品名称
    填制人 = mvar填制人
End Property
