VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsurePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'保持属性值的局部变量
Private mvar分币处理 As Boolean '局部复制
Private mvar医保接口打印票据 As Boolean '局部复制
Private mvar门诊病人结算作废 As Boolean '局部复制
Private mvar门诊预结算 As Boolean '局部复制
Private mvar门诊必须传递明细 As Boolean '局部复制
Private mvar门诊结算_结帐设置 As Boolean '局部复制
Private mvar未结清出院 As Boolean '局部复制
Private mvar结算使用个人帐户 As Boolean '局部复制
Private mvar出院结算必须出院 As Boolean '局部复制
Private mvar出院病人结算作废 As Boolean '局部复制
Private mvar中途结帐仅处理已上传部分 As Boolean '局部复制
Private mvar结帐设置后调用接口 As Boolean '局部复制
Private mvar结帐作废后打印回单 As Boolean '局部复制
Private mvar住院结算作废 As Boolean '局部复制
Private mvar允许结多次住院费用 As Boolean '局部复制
Public Property Let 允许结多次住院费用(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.允许结多次住院费用 = 5
    mvar允许结多次住院费用 = vData
End Property


Public Property Get 允许结多次住院费用() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.允许结多次住院费用
    允许结多次住院费用 = mvar允许结多次住院费用
End Property



Public Property Let 住院结算作废(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.住院结算作废 = 5
    mvar住院结算作废 = vData
End Property


Public Property Get 住院结算作废() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.住院结算作废
    住院结算作废 = mvar住院结算作废
End Property



Public Property Let 结帐作废后打印回单(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐作废后打印回单 = 5
    mvar结帐作废后打印回单 = vData
End Property


Public Property Get 结帐作废后打印回单() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐作废后打印回单
    结帐作废后打印回单 = mvar结帐作废后打印回单
End Property



Public Property Let 结帐设置后调用接口(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐设置后调用接口 = 5
    mvar结帐设置后调用接口 = vData
End Property


Public Property Get 结帐设置后调用接口() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐设置后调用接口
    结帐设置后调用接口 = mvar结帐设置后调用接口
End Property



Public Property Let 中途结帐仅处理已上传部分(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.中途结帐仅处理已上传部分 = 5
    mvar中途结帐仅处理已上传部分 = vData
End Property


Public Property Get 中途结帐仅处理已上传部分() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.中途结帐仅处理已上传部分
    中途结帐仅处理已上传部分 = mvar中途结帐仅处理已上传部分
End Property



Public Property Let 出院病人结算作废(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出院病人结算作废 = 5
    mvar出院病人结算作废 = vData
End Property


Public Property Get 出院病人结算作废() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出院病人结算作废
    出院病人结算作废 = mvar出院病人结算作废
End Property



Public Property Let 出院结算必须出院(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出院结算必须出院 = 5
    mvar出院结算必须出院 = vData
End Property


Public Property Get 出院结算必须出院() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出院结算必须出院
    出院结算必须出院 = mvar出院结算必须出院
End Property



Public Property Let 结算使用个人帐户(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算使用个人帐户 = 5
    mvar结算使用个人帐户 = vData
End Property


Public Property Get 结算使用个人帐户() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算使用个人帐户
    结算使用个人帐户 = mvar结算使用个人帐户
End Property



Public Property Let 未结清出院(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.未结清出院 = 5
    mvar未结清出院 = vData
End Property


Public Property Get 未结清出院() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.未结清出院
    未结清出院 = mvar未结清出院
End Property



Public Property Let 门诊结算_结帐设置(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊结算_结帐设置 = 5
    mvar门诊结算_结帐设置 = vData
End Property


Public Property Get 门诊结算_结帐设置() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊结算_结帐设置
    门诊结算_结帐设置 = mvar门诊结算_结帐设置
End Property



Public Property Let 门诊必须传递明细(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊必须传递明细 = 5
    mvar门诊必须传递明细 = vData
End Property


Public Property Get 门诊必须传递明细() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊必须传递明细
    门诊必须传递明细 = mvar门诊必须传递明细
End Property



Public Property Let 门诊预结算(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊预结算 = 5
    mvar门诊预结算 = vData
End Property


Public Property Get 门诊预结算() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊预结算
    门诊预结算 = mvar门诊预结算
End Property



Public Property Let 门诊病人结算作废(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊病人结算作废 = 5
    mvar门诊病人结算作废 = vData
End Property


Public Property Get 门诊病人结算作废() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊病人结算作废
    门诊病人结算作废 = mvar门诊病人结算作废
End Property



Public Property Let 医保接口打印票据(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医保接口打印票据 = 5
    mvar医保接口打印票据 = vData
End Property


Public Property Get 医保接口打印票据() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医保接口打印票据
    医保接口打印票据 = mvar医保接口打印票据
End Property



Public Property Let 分币处理(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.分币处理 = 5
    mvar分币处理 = vData
End Property


Public Property Get 分币处理() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.分币处理
    分币处理 = mvar分币处理
End Property



