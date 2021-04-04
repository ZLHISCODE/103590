VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Enum gCard_Type
    EM_CardType_Cash = 0    '现金
    EM_CardType_Consume = 1 '消费卡
    EM_CardType_Square = 2  '结算卡
    EM_CardType_HospitalAccount = 3 '预存款
    EM_CardType_Insure = 4  '险类
    EM_CardType_other = 5  '其他
End Enum
Public Enum gKeyBoard_Type  '键盘控制
    EM_KeyBoard_NotKey = 0
    EM_KeyBoard_NumberKey = 1
    EM_KeyBoard_CharKey = 2
End Enum
'保持属性值的局部变量
Private mvar接口序号 As Long     '局部复制
Private mvar接口编码 As String     '局部复制
Private mvar名称 As String '局部复制
Private mvar启用 As Boolean '局部复制
Private mvar接口程序名 As String '局部复制
'保持属性值的局部变量
Private mvar设置 As Integer
Private mvar自动读取 As Integer
'保持属性值的局部变量
Private mvar系统 As Boolean '局部复制
Private mvar结算方式 As String '局部复制
Private mvar前缀文本 As String '局部复制
Private mvar卡号长度 As Long '局部复制
Private mvar自制卡 As Boolean
Private mvar消费卡 As Boolean '
Private mvar短名 As String
Private mvar是否严格控制 As Boolean
Private mvar是否存在帐户 As Boolean
Private mvar是否全退 As Boolean
Private mvar备注 As String
Private mvar特定项目 As String
Private mvar卡号密文规则 As String
Private mvar卡号重复使用 As Boolean
Private mvar自动读取间隔 As Integer  '自动读取间隔
Private mvar是否退现 As Boolean
Private mvar密码长度 As Integer
Private mvar密码长度限制 As Integer
Private mvar密码规则 As Integer
Private mvar图像标识 As String
Private mvar支付启用 As Boolean
Private mvar支付图像标识 As String
Private mvar密码输入限制 As Long
Private mvar是否缺省密码 As Boolean
Private mvar缺省标志 As Boolean
Private mvar功能键 As String
Private mvar快键 As String
Private mvar模糊查找项 As Boolean
Private mvar是否模糊查找 As Boolean
Private mvar是否转帐及代扣 As Boolean
Private mvar是否刷卡  As Boolean
Private mvar是否扫描  As Boolean
Private mvar是否接触式读卡  As Boolean
Private mvar是否非接触式读卡 As Boolean
Private mvar是否持卡消费 As Boolean
'保持属性值的局部变量
Private mvar是否制卡 As Boolean '局部复制
Private mvar是否发卡 As Boolean '局部复制
Private mvar是否写卡 As Boolean '局部复制
Private mvar结算性质 As Integer   '局部复制
'保持属性值的局部变量
Private mvar是否退款验卡 As Boolean '局部复制
Private mvar是否证件 As Boolean
Private mvar设备是否启用回车 As Boolean
Private mvar是否缺省退现 As Boolean
Private mvar是否独立结算 As Boolean
Private mvarCardObject As Object '局部复制
Private mvarInitCompents As Boolean '局部复制
Private mvar键盘类型 As gKeyBoard_Type
Private mvar卡种类 As gCard_Type  '1-消费卡;2-结算卡;3-现金;4-帐户;5-险类
Private mvar险类 As Integer
Private mvar险类名称 As String
Private mvar是否发送调用接口 As Boolean
Private mvar卡号识别规则 As Integer
'保持属性值的局部变量
Private mvar是否支持扫码付 As Boolean '局部复制
Public Property Let 是否支持扫码付(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否支持扫码付 = 5
    mvar是否支持扫码付 = vData
End Property


Public Property Get 是否支持扫码付() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否支持扫码付
    是否支持扫码付 = mvar是否支持扫码付
End Property



Public Property Let 卡号识别规则(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算性质 = 5
    mvar卡号识别规则 = vData
End Property

Public Property Get 卡号识别规则() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡号识别规则
    卡号识别规则 = mvar卡号识别规则
End Property



Public Property Let 是否发送调用接口(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算性质 = 5
    mvar是否发送调用接口 = vData
End Property

Public Property Get 是否发送调用接口() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否发送调用接口
    是否发送调用接口 = mvar是否发送调用接口
End Property

Public Property Let 卡种类(ByVal vData As gCard_Type)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡种类  = 5
    mvar卡种类 = vData
End Property
Public Property Get 卡种类() As gCard_Type
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡种类
    卡种类 = mvar卡种类
End Property

Public Property Let 险类名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类名称  = 5
    mvar险类名称 = vData
End Property
Public Property Get 险类名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类名称
    险类名称 = mvar险类名称
End Property
Public Property Let 险类(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类 = 5
    mvar险类 = vData
End Property

Public Property Let 键盘类型(ByVal vData As gKeyBoard_Type)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.键盘类型  = 5
    mvar键盘类型 = vData
End Property
Public Property Get 键盘类型() As gKeyBoard_Type
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.键盘类型
    键盘类型 = mvar键盘类型
End Property


Public Property Let InitCompents(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.消费卡 = 5
    mvarInitCompents = vData
End Property
Public Property Get InitCompents() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.消费卡
    InitCompents = mvarInitCompents
End Property

Public Property Set CardObject(ByVal vData As Object)
    '向属性指派对象时使用，位于 Set 语句的左边。
    'Syntax: Set x.CardObject = Form1
    Set mvarCardObject = vData
End Property


Public Property Get CardObject() As Object
    '检索属性值时使用，位于赋值语句的右边。
    'Syntax: Debug.Print X.CardObject
    Set CardObject = mvarCardObject
End Property

Public Property Let 是否独立结算(ByVal vData As Boolean)
    mvar是否独立结算 = vData
End Property
Public Property Get 是否独立结算() As Boolean
    是否独立结算 = mvar是否独立结算
End Property

Public Property Let 是否缺省退现(ByVal vData As Boolean)
    mvar是否缺省退现 = vData
End Property
Public Property Get 是否缺省退现() As Boolean
    是否缺省退现 = mvar是否缺省退现
End Property

Public Property Let 设备是否启用回车(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否写卡 = 5
    mvar设备是否启用回车 = vData
End Property

Public Property Get 设备是否启用回车() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否写卡
    设备是否启用回车 = mvar设备是否启用回车
End Property

Public Property Let 是否证件(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否写卡 = 5
    mvar是否证件 = vData
End Property

Public Property Get 是否证件() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否写卡
    是否证件 = mvar是否证件
End Property

Public Property Let 是否退款验卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否退款验卡 = 5
    mvar是否退款验卡 = vData
End Property


Public Property Get 是否退款验卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否退款验卡
    是否退款验卡 = mvar是否退款验卡
End Property




Public Property Let 是否持卡消费(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否写卡 = 5
    mvar是否持卡消费 = vData
End Property

Public Property Get 是否持卡消费() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否写卡
    是否持卡消费 = mvar是否持卡消费
End Property

Public Property Let 是否刷卡(ByVal vData As Boolean)
    mvar是否刷卡 = vData
End Property
Public Property Get 是否刷卡() As Boolean
    是否刷卡 = mvar是否刷卡
End Property

Public Property Let 是否扫描(ByVal vData As Boolean)
    mvar是否扫描 = vData
End Property
Public Property Get 是否扫描() As Boolean
    是否扫描 = mvar是否扫描
End Property

Public Property Let 是否接触式读卡(ByVal vData As Boolean)
    mvar是否接触式读卡 = vData
End Property
Public Property Get 是否接触式读卡() As Boolean
    是否接触式读卡 = mvar是否接触式读卡
End Property

Public Property Let 是否非接触式读卡(ByVal vData As Boolean)
    mvar是否非接触式读卡 = vData
End Property
Public Property Get 是否非接触式读卡() As Boolean
    是否非接触式读卡 = mvar是否非接触式读卡
End Property

Public Property Let 是否转帐及代扣(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否写卡 = 5
    mvar是否转帐及代扣 = vData
End Property

Public Property Get 是否转帐及代扣() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否写卡
    是否转帐及代扣 = mvar是否转帐及代扣
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

Public Property Let 是否写卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否写卡 = 5
    mvar是否写卡 = vData
End Property

Public Property Get 是否写卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否写卡
    是否写卡 = mvar是否写卡
End Property

Public Property Let 是否发卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否发卡 = 5
    mvar是否发卡 = vData
End Property


Public Property Get 是否发卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否发卡
    是否发卡 = mvar是否发卡
End Property



Public Property Let 是否制卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否制卡 = 5
    mvar是否制卡 = vData
End Property


Public Property Get 是否制卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否制卡
    是否制卡 = mvar是否制卡
End Property

Public Property Let 是否模糊查找(ByVal vData As Boolean)
    mvar是否模糊查找 = vData
End Property
Public Property Get 是否模糊查找() As Boolean
    是否模糊查找 = mvar是否模糊查找
End Property


Public Property Let 模糊查找项(ByVal vData As Boolean)
    mvar模糊查找项 = vData
End Property
Public Property Get 模糊查找项() As Boolean
    模糊查找项 = mvar模糊查找项
End Property

Public Property Let 快键(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前缀文本 = 5
    mvar快键 = vData
End Property
Public Property Get 快键() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前缀文本
    快键 = mvar快键
End Property
Public Property Let 功能键(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前缀文本 = 5
    mvar功能键 = vData
End Property
Public Property Get 功能键() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前缀文本
    功能键 = mvar功能键
End Property


Public Property Let 缺省标志(ByVal vData As Boolean)
    mvar缺省标志 = vData
End Property
Public Property Get 缺省标志() As Boolean
    缺省标志 = mvar缺省标志
End Property

Public Property Let 是否缺省密码(ByVal vData As Boolean)
    mvar是否缺省密码 = vData
End Property
Public Property Get 是否缺省密码() As Boolean
    是否缺省密码 = mvar是否缺省密码
End Property
Public Property Let 密码输入限制(ByVal vData As Long)
    mvar密码输入限制 = vData
End Property
Public Property Get 密码输入限制() As Long
    密码输入限制 = mvar密码输入限制
End Property
Public Property Let 支付图像标识(ByVal vData As String)
    mvar支付图像标识 = vData
End Property
Public Property Get 支付图像标识() As String
    支付图像标识 = mvar支付图像标识
End Property
Public Property Let 支付启用(ByVal vData As Boolean)
    mvar支付启用 = vData
End Property
Public Property Get 支付启用() As Boolean
    支付启用 = mvar支付启用
End Property

Public Property Let 密码长度(ByVal vData As Integer)
    mvar密码长度 = vData
End Property
Public Property Get 密码长度() As Integer
    密码长度 = mvar密码长度
End Property

Public Property Let 密码长度限制(ByVal vData As Integer)
    mvar密码长度限制 = vData
End Property
Public Property Get 密码长度限制() As Integer
    密码长度限制 = mvar密码长度限制
End Property
Public Property Let 密码规则(ByVal vData As Integer)
    mvar密码规则 = vData
End Property
Public Property Get 密码规则() As Integer
    密码规则 = mvar密码规则
End Property

Public Property Let 是否退现(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前缀文本 = 5
    mvar是否退现 = vData
End Property
Public Property Get 是否退现() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前缀文本
    是否退现 = mvar是否退现
End Property

Public Property Let 自动读取间隔(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前缀文本 = 5
    mvar自动读取间隔 = vData
End Property
Public Property Get 自动读取间隔() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前缀文本
    自动读取间隔 = mvar自动读取间隔
End Property


Public Property Let 卡号重复使用(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前缀文本 = 5
    mvar卡号重复使用 = vData
End Property
Public Property Get 卡号重复使用() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前缀文本
    卡号重复使用 = mvar卡号重复使用
End Property

Public Property Let 消费卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前缀文本 = 5
    mvar消费卡 = vData
End Property
Public Property Get 消费卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前缀文本
    消费卡 = mvar消费卡
End Property

Public Property Let 卡号密文规则(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前缀文本 = 5
    mvar卡号密文规则 = vData
End Property
Public Property Get 卡号密文规则() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前缀文本
    卡号密文规则 = mvar卡号密文规则
End Property

Public Property Let 卡号长度(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡号长度 = 5
    mvar卡号长度 = vData
End Property
Public Property Get 卡号长度() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡号长度
    卡号长度 = mvar卡号长度
End Property

Public Property Let 前缀文本(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前缀文本 = 5
    mvar前缀文本 = vData
End Property
Public Property Get 前缀文本() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前缀文本
    前缀文本 = mvar前缀文本
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
Public Property Let 系统(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.系统 = 5
    mvar系统 = vData
End Property
Public Property Get 系统() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.系统
     系统 = mvar系统
End Property

Public Property Let 是否自动读取(ByVal vData As Integer)
    mvar自动读取 = vData
End Property

Public Property Get 是否自动读取() As Integer
    是否自动读取 = mvar自动读取
End Property
Public Property Let 可否设置(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类 = 5
    mvar设置 = vData
End Property
Public Property Get 可否设置() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类
    可否设置 = mvar设置
End Property
Public Property Let 接口程序名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.接口程序名 = 5
    mvar接口程序名 = vData
End Property


Public Property Get 接口程序名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.接口程序名
    接口程序名 = mvar接口程序名
End Property

Public Property Let 启用(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.启用 = 5
    mvar启用 = vData
End Property


Public Property Get 启用() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.启用
    启用 = mvar启用
End Property

Public Property Let 名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.名称 = 5
    mvar名称 = vData
End Property


Public Property Get 名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.名称
    名称 = mvar名称
End Property

Public Property Let 接口编码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.编码 = 5
    mvar接口编码 = vData
End Property

Public Property Get 接口编码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.编码
    接口编码 = mvar接口编码
End Property
Public Property Let 接口序号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.编码 = 5
    mvar接口序号 = vData
End Property

Public Property Get 接口序号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.编码
    接口序号 = mvar接口序号
End Property
Public Property Let 自制卡(ByVal vData As Boolean)
    mvar自制卡 = vData
End Property

Public Property Get 自制卡() As Boolean
    自制卡 = mvar自制卡
End Property

  
Public Property Let 短名(ByVal vData As String)
    mvar短名 = vData
End Property
Public Property Get 短名() As String
    短名 = mvar短名
End Property
Public Property Let 备注(ByVal vData As String)
    mvar备注 = vData
End Property
Public Property Get 备注() As String
    备注 = mvar备注
End Property
Public Property Let 特定项目(ByVal vData As String)
    mvar特定项目 = vData
End Property
Public Property Get 特定项目() As String
    特定项目 = mvar特定项目
End Property
    
Public Property Let 是否严格控制(ByVal vData As Boolean)
    mvar是否严格控制 = vData
End Property
Public Property Get 是否严格控制() As Boolean
    是否严格控制 = mvar是否严格控制
End Property
Public Property Let 是否存在帐户(ByVal vData As Boolean)
    mvar是否存在帐户 = vData
End Property
Public Property Get 是否存在帐户() As Boolean
    是否存在帐户 = mvar是否存在帐户
End Property
Public Property Let 是否全退(ByVal vData As Boolean)
    mvar是否全退 = vData
End Property
Public Property Get 是否全退() As Boolean
    是否全退 = mvar是否全退
End Property
Public Property Let 图像标识(ByVal vData As String)
    mvar图像标识 = vData
End Property
Public Property Get 图像标识() As String
    图像标识 = mvar图像标识
End Property
Public Function zlCardNOEncrypt(ByVal strCardNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取加密的卡号
    '入参:strCardNO-未加密卡号
    '返回:带**的卡号,如果错误,返回空:21432**13234
    '编制:刘兴洪
    '日期:2012-04-26 11:29:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varPass As Variant, strRule As String
    Dim strCardPassText As String, I As Long, J As Long
    strRule = 卡号密文规则
    If 消费卡 Then
        If Val(strRule) = 1 Then zlCardNOEncrypt = String(Len(strCardNo), "*"): Exit Function
        zlCardNOEncrypt = strCardNo
    End If
    
    strCardPassText = strRule
    If strCardPassText = "" Then zlCardNOEncrypt = strCardNo: Exit Function

    varPass = Split(strCardPassText & "-", "-")
    If Val(varPass(0)) = 0 Or Val(varPass(1)) = 0 Then
        '最后几位显示*
        I = IIf(Val(varPass(0)) = 0, Val(varPass(1)), Val(varPass(0)))
        If I = 0 Then zlCardNOEncrypt = strCardNo: Exit Function
        J = Len(strCardNo) - I: J = IIf(J < 0, 0, J)
        zlCardNOEncrypt = Mid(strCardNo, 1, J) & String(I, "*")
        Exit Function
    End If
    I = Val(varPass(0)): J = Val(varPass(1))
    If I > Len(strCardNo) Then zlCardNOEncrypt = strCardNo: Exit Function
    If J > Len(strCardNo) Then J = Len(strCardNo)
    If J < I Then J = I
   zlCardNOEncrypt = Mid(strCardNo, 1, I - 1) & String(J - I + 1, "*") & Mid(strCardNo, J + 1)
End Function

