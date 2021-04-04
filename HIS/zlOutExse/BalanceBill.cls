VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BalanceBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'保持属性值的局部变量
Private mvar全自付 As Currency '局部复制
Private mvar先自付 As Currency '局部复制
Private mvar预结算 As BalanceMoneys '局部复制
Private mvarNO As String '局部复制
Private mvar划价单 As Collection '局部复制
Private mvar收费单 As Collection '局部复制
Private mvar实收合计 As Currency '局部复制
Private mvar进入统筹 As Currency '局部复制
Public Property Let 进入统筹(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.进入统筹 = 5
    mvar进入统筹 = vData
End Property


Public Property Get 进入统筹() As Currency
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.进入统筹
    进入统筹 = mvar进入统筹
End Property



Public Property Let 实收合计(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.实收合计 = 5
    mvar实收合计 = vData
End Property


Public Property Get 实收合计() As Currency
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.实收合计
    实收合计 = mvar实收合计
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



Public Property Set 预结算(ByVal vData As BalanceMoneys)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.预结算 = Form1
    Set mvar预结算 = vData
End Property


Public Property Get 预结算() As BalanceMoneys
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预结算
    Set 预结算 = mvar预结算
End Property



Public Property Let 先自付(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.先自付 = 5
    mvar先自付 = vData
End Property


Public Property Get 先自付() As Currency
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.先自付
    先自付 = mvar先自付
End Property



Public Property Let 全自付(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.全自付 = 5
    mvar全自付 = vData
End Property


Public Property Get 全自付() As Currency
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.全自付
    全自付 = mvar全自付
End Property


Private Sub Class_Initialize()
    Set Me.预结算 = New BalanceMoneys
End Sub

Private Sub Class_Terminate()
    Set Me.预结算 = Nothing
End Sub
