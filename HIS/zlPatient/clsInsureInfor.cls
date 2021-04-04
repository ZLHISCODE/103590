VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsureInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
Private mvar是否个帐结算 As Boolean '局部复制
Private mvar帐户余额 As Double '局部复制
Private mvar个帐限额 As Double '局部复制
Private mvar个帐透支 As Double '局部复制
Private mvar个帐支付 As Double '局部复制
Private mvar统筹支付 As Double '局部复制
Private mvar身份验证信息 As String '局部复制
Private mvar险类 As Integer '局部复制
Private mvar是否作废全退 As Boolean '局部复制
Private mvarbytMCMode As Byte '局部复制
Private mvar结算数据 As String '局部复制
Private mvar是否自动出院 As Boolean '局部复制
Public Property Let 是否自动出院(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否自动出院 = 5
    mvar是否自动出院 = vData
End Property


Public Property Get 是否自动出院() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否自动出院
    是否自动出院 = mvar是否自动出院
End Property



Public Property Let 结算数据(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算数据 = 5
    mvar结算数据 = vData
End Property


Public Property Get 结算数据() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算数据
    结算数据 = mvar结算数据
End Property



Public Property Let bytMCMode(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.bytMCMode = 5
    mvarbytMCMode = vData
End Property


Public Property Get bytMCMode() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.bytMCMode
    bytMCMode = mvarbytMCMode
End Property



Public Property Let 是否作废全退(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否作废全退 = 5
    mvar是否作废全退 = vData
End Property


Public Property Get 是否作废全退() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否作废全退
    是否作废全退 = mvar是否作废全退
End Property



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



Public Property Let 身份验证信息(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.身份验证信息 = 5
    mvar身份验证信息 = vData
End Property


Public Property Get 身份验证信息() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.身份验证信息
    身份验证信息 = mvar身份验证信息
End Property



Public Property Let 统筹支付(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.统筹支付 = 5
    mvar统筹支付 = vData
End Property


Public Property Get 统筹支付() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.统筹支付
    统筹支付 = mvar统筹支付
End Property



Public Property Let 个帐支付(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.个帐支付 = 5
    mvar个帐支付 = vData
End Property


Public Property Get 个帐支付() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.个帐支付
    个帐支付 = mvar个帐支付
End Property



Public Property Let 个帐透支(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.个帐透支 = 5
    mvar个帐透支 = vData
End Property


Public Property Get 个帐透支() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.个帐透支
    个帐透支 = mvar个帐透支
End Property



Public Property Let 个帐限额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.个帐限额 = 5
    mvar个帐限额 = vData
End Property


Public Property Get 个帐限额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.个帐限额
    个帐限额 = mvar个帐限额
End Property



Public Property Let 帐户余额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.帐户余额 = 5
    mvar帐户余额 = vData
End Property


Public Property Get 帐户余额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.帐户余额
    帐户余额 = mvar帐户余额
End Property



Public Property Let 是否个帐结算(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否个帐结算 = 5
    mvar是否个帐结算 = vData
End Property


Public Property Get 是否个帐结算() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否个帐结算
    是否个帐结算 = mvar是否个帐结算
End Property



