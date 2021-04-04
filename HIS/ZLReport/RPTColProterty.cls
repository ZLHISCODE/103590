VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTColProterty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Key As String

Private mvar条件名称 As String  '局部复制
Private mvar条件字段 As String  '局部复制
Private mvar条件关系 As String '局部复制
Private mvar条件值 As String  '局部复制
Private mvar字体颜色 As Long '局部复制
Private mvar背景颜色 As Long '局部复制
Private mvar是否加粗 As Boolean '局部复制
Private mvar是否整行应用 As Boolean '局部复制
Private mint对齐 As Integer                 '0-自动（字段类型确认）；1-居左；2-居中；3-居右

Public Property Let 条件名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.参数名 = 5
    mvar条件名称 = vData
End Property

Public Property Get 条件名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.名称
    条件名称 = mvar条件名称
End Property

Public Property Let 条件字段(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.参数值来源 = 5
    mvar条件字段 = vData
End Property

Public Property Get 条件字段() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.参数值来源
    条件字段 = mvar条件字段
End Property

Public Property Let 条件关系(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表 = 5
    mvar条件关系 = vData
End Property

Public Property Get 条件关系() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表
    条件关系 = mvar条件关系
End Property

Public Property Let 条件值(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表 = 5
    mvar条件值 = vData
End Property

Public Property Get 条件值() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表
    条件值 = mvar条件值
End Property


Public Property Let 字体颜色(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表 = 5
    mvar字体颜色 = vData
End Property

Public Property Get 字体颜色() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表
    字体颜色 = mvar字体颜色
End Property

Public Property Let 背景颜色(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表 = 5
    mvar背景颜色 = vData
End Property

Public Property Get 背景颜色() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表
    背景颜色 = mvar背景颜色
End Property

Public Property Let 是否加粗(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表 = 5
    mvar是否加粗 = vData
End Property

Public Property Get 是否加粗() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表
    是否加粗 = mvar是否加粗
End Property

Public Property Let 是否整行应用(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表 = 5
    mvar是否整行应用 = vData
End Property

Public Property Get 是否整行应用() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表
    是否整行应用 = mvar是否整行应用
End Property

Public Property Let 对齐(ByVal vData As Integer)
    mint对齐 = vData
End Property

Public Property Get 对齐() As Integer
    对齐 = mint对齐
End Property
