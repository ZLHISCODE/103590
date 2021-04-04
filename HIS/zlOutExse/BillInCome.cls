VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillInCome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit '要求变量声明

Public Key As String

Private mvar收入项目ID As Long '局部复制
Private mvar收入项目 As String '局部复制
Private mvar收据费目 As String '局部复制
Private mvar标准单价 As Double '局部复制
Private mvar现价 As Double '局部复制
Private mvar原价 As Double '局部复制
Private mvar应收金额 As Double '局部复制
Private mvar实收金额 As Double '局部复制
Private mvar统筹金额 As Double '局部复制

Public Property Let 统筹金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.统筹金额 = 5
    mvar统筹金额 = vData
End Property

Public Property Get 统筹金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.统筹金额
    统筹金额 = mvar统筹金额
End Property

Public Property Let 实收金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.实收金额 = 5
    mvar实收金额 = vData
End Property

Public Property Get 实收金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.实收金额
    实收金额 = mvar实收金额
End Property

Public Property Let 应收金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.应收金额 = 5
    mvar应收金额 = vData
End Property

Public Property Get 应收金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.应收金额
    应收金额 = mvar应收金额
End Property

Public Property Let 标准单价(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.标准单价 = 5
    mvar标准单价 = vData
End Property

Public Property Get 标准单价() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.标准单价
    标准单价 = mvar标准单价
End Property

Public Property Let 收据费目(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收据费目 = 5
    mvar收据费目 = vData
End Property

Public Property Get 收据费目() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.收据费目
    收据费目 = mvar收据费目
End Property

Public Property Let 收入项目(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收据费目 = 5
    mvar收入项目 = vData
End Property

Public Property Get 收入项目() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.收据费目
    收入项目 = mvar收入项目
End Property

Public Property Let 收入项目ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收入项目ID = 5
    mvar收入项目ID = vData
End Property

Public Property Get 收入项目ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.收入项目ID
    收入项目ID = mvar收入项目ID
End Property

Public Property Let 原价(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.标准单价 = 5
    mvar原价 = vData
End Property

Public Property Get 原价() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.标准单价
    原价 = mvar原价
End Property

Public Property Let 现价(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.标准单价 = 5
    mvar现价 = vData
End Property

Public Property Get 现价() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.标准单价
    现价 = mvar现价
End Property

