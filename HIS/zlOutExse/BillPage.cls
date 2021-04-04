VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Details"

Option Explicit '要求变量声明

Public Key As String
Private mvarDetails As BillDetails '局部复制
Private mvarNO As String '局部复制
Private mvar应收金额 As Double '局部复制
Private mvar实收金额 As Double '局部复制
Private mvar保险金额 As Double '局部复制
Private mvar冲预交额 As Double '局部复制
Private mvar应缴金额 As Double '局部复制
Private mvar误差金额 As Double '局部复制
Private mvar收费结算 As String '局部复制
Private mvar进入统筹 As Double '局部复制
Private mvar全自付 As Double '局部复制
Private mvar先自付 As Double '局部复制
Private mvar开单部门ID As Long '局部复制
Private mvar开单人 As String '局部复制

Private mvar消费卡结算 As String '局部复制
Private mvar消费卡刷卡额 As Double '局部复制
Private mvar收费单号 As String

'保持属性值的局部变量
Private mvar煎法 As String '局部复制
'保持属性值的局部变量
Private mvar医嘱序号 As Long '局部复制
Private mvar结帐ID As Long
Public Property Let 结帐ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费结算 = 5
    mvar结帐ID = vData
End Property
Public Property Get 结帐ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.收费结算
    结帐ID = mvar结帐ID
End Property


Public Property Let 医嘱序号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医嘱序号 = 5
    mvar医嘱序号 = vData
End Property
Public Property Get 医嘱序号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医嘱序号
    医嘱序号 = mvar医嘱序号
End Property
Public Property Let 煎法(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.煎法 = 5
    mvar煎法 = vData
End Property


Public Property Get 煎法() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.煎法
    煎法 = mvar煎法
End Property

Private Sub Class_Initialize()
    Set Me.Details = New BillDetails
End Sub

Private Sub Class_Terminate()
    Set Me.Details = Nothing
End Sub

Public Property Set Details(ByVal vData As BillDetails)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    Set mvarDetails = vData
End Property

Public Property Get Details() As BillDetails
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    Set Details = mvarDetails
End Property
Public Property Let 收费单号(ByVal vData As String)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar收费单号 = vData
End Property

Public Property Get 收费单号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    收费单号 = mvar收费单号
End Property



Public Property Let NO(ByVal vData As String)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvarNO = vData
End Property

Public Property Get NO() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    NO = mvarNO
End Property

Public Property Let 应收金额(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar应收金额 = vData
End Property

Public Property Get 应收金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    应收金额 = mvar应收金额
End Property

Public Property Let 实收金额(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar实收金额 = vData
End Property

Public Property Get 实收金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    实收金额 = mvar实收金额
End Property

Public Property Let 保险金额(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar保险金额 = vData
End Property

Public Property Get 保险金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    保险金额 = mvar保险金额
End Property

Public Property Let 冲预交额(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar冲预交额 = vData
End Property

Public Property Get 冲预交额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    冲预交额 = mvar冲预交额
End Property

Public Property Let 应缴金额(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar应缴金额 = vData
End Property

Public Property Get 应缴金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    应缴金额 = mvar应缴金额
End Property

Public Property Let 误差金额(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar误差金额 = vData
End Property

Public Property Get 误差金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    误差金额 = mvar误差金额
End Property

Public Property Let 收费结算(ByVal vData As String)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar收费结算 = vData
End Property

Public Property Get 收费结算() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    收费结算 = mvar收费结算
End Property

Public Property Let 进入统筹(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar进入统筹 = vData
End Property

Public Property Get 进入统筹() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    进入统筹 = mvar进入统筹
End Property

Public Property Let 全自付(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar全自付 = vData
End Property

Public Property Get 全自付() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    全自付 = mvar全自付
End Property

Public Property Let 先自付(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar先自付 = vData
End Property

Public Property Get 先自付() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    先自付 = mvar先自付
End Property

Public Property Get 开单人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.开单人
    开单人 = mvar开单人
End Property

Public Property Let 开单人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.开单人 = 5
    mvar开单人 = vData
End Property

Public Property Let 开单部门ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.开单部门ID = 5
    mvar开单部门ID = vData
End Property

Public Property Get 开单部门ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.开单部门ID
    开单部门ID = mvar开单部门ID
End Property

Public Property Let 消费卡结算(ByVal vData As String)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar消费卡结算 = vData
End Property

Public Property Get 消费卡结算() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    消费卡结算 = mvar消费卡结算
End Property

Public Property Let 消费卡刷卡额(ByVal vData As Double)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Details = vData
    mvar消费卡刷卡额 = vData
End Property

Public Property Get 消费卡刷卡额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Details
    消费卡刷卡额 = mvar消费卡刷卡额
End Property
