VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubPayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'保持属性值的局部变量
'挂号三方卡支付返回多种结算方式
Private mvar结算方式 As String '局部复制
Private mvar结算金额 As Double '局部复制
Private mvarPayRow As Long '局部复制
Private mvar结算号码 As String '局部复制
Private mvar交易流水号 As String '局部复制
Private mvar交易说明 As String '局部复制
Public Property Let 交易说明(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.交易说明 = 5
    mvar交易说明 = vData
End Property


Public Property Get 交易说明() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.交易说明
    交易说明 = mvar交易说明
End Property



Public Property Let 交易流水号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.交易流水号 = 5
    mvar交易流水号 = vData
End Property


Public Property Get 交易流水号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.交易流水号
    交易流水号 = mvar交易流水号
End Property



Public Property Let 结算号码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算号码 = 5
    mvar结算号码 = vData
End Property


Public Property Get 结算号码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算号码
    结算号码 = mvar结算号码
End Property



Public Property Let PayRow(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.PayRow = 5
    mvarPayRow = vData
End Property


Public Property Get PayRow() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.PayRow
    PayRow = mvarPayRow
End Property



Public Property Let 结算金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算金额 = 5
    mvar结算金额 = vData
End Property


Public Property Get 结算金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算金额
    结算金额 = mvar结算金额
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



