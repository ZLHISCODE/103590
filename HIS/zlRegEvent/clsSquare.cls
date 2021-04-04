VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSquare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'保持属性值的局部变量
Private mvar结算方式 As String '局部复制
Private mvar结算金额 As Double '局部复制
Private mvar结算性质 As Integer '局部复制
Private mvar结算号码 As String '局部复制
Private mvar备注 As String '局部复制
Private mvarNO As String '局部复制
Private mvar允许修改 As Boolean  '局部复制
Private mvar误差费 As Double  '局部复制
Private mvar刷卡类别ID As Long   '局部复制
Private mvar刷卡卡号 As String    '局部复制
Private mvar交易流水号 As String    '局部复制
Private mvar交易说明 As String    '局部复制
Private mvar其他交易信息 As String    '局部复制

Public Property Let 其他交易信息(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.其他交易信息 = 5
    mvar其他交易信息 = vData
End Property
Public Property Get 其他交易信息() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.其他交易信息
    其他交易信息 = mvar其他交易信息
End Property
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


 Public Property Let 刷卡卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.刷卡卡号 = 5
    mvar刷卡卡号 = vData
End Property
Public Property Get 刷卡卡号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.刷卡卡号
    刷卡卡号 = mvar刷卡卡号
End Property

 Public Property Let 刷卡类别ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.刷卡类别ID = 5
    mvar刷卡类别ID = vData
End Property
Public Property Get 刷卡类别ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.刷卡类别ID
    刷卡类别ID = mvar刷卡类别ID
End Property

 Public Property Let 结算性质(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算性质 = 5
    mvar结算性质 = vData
End Property
Public Property Get 结算性质() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算性质
    结算性质 = mvar结算性质
End Property

Public Property Let 误差费(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.误差费 = 5
    mvar误差费 = vData
End Property
Public Property Get 误差费() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.误差费
    误差费 = mvar误差费
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

Public Property Let NO(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.NO = 5
    mvarNO = vData
End Property
Public Property Get NO() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.NO
    NO = mvarNO
End Property
Public Property Let 备注(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.备注 = 5
    mvar备注 = vData
End Property
Public Property Get 备注() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.备注
    备注 = mvar备注
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



