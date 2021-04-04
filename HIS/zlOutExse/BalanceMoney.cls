VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BalanceMoney"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'保持属性值的局部变量
Private mvar结算方式 As String '局部复制
Private mvar允许修改 As Boolean '局部复制
Private mvar原始金额 As Double '局部复制
'保持属性值的局部变量
Private mvar有效金额 As Double '局部复制
Public Property Let 有效金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.有效金额 = 5
    mvar有效金额 = vData
End Property


Public Property Get 有效金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.有效金额
    有效金额 = mvar有效金额
End Property



Public Property Let 原始金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.原始金额 = 5
    mvar原始金额 = vData
End Property


Public Property Get 原始金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.原始金额
    原始金额 = mvar原始金额
End Property



Public Property Let 允许修改(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.允许修改 = 5
    mvar允许修改 = vData
End Property


Public Property Get 允许修改() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.允许修改
    允许修改 = mvar允许修改
End Property



Public Property Let 结算方式(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算方式 = 5
    mvar结算方式 = vData
End Property


Public Property Get 结算方式() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算方式
    结算方式 = mvar结算方式
End Property



