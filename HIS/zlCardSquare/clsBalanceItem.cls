VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
Private mvar结算类型 As Integer '局部复制
Private mvar结算ID As Long '局部复制
Private mvarobjCard As Card '局部复制
Private mvar卡类别ID As Long '局部复制
Private mvar预交ID As Long '局部复制
Private mvar交易流水号 As String '局部复制
Private mvar交易说明 As String '局部复制
Private mvar结算金额 As Double '局部复制
Private mvar卡号 As String '局部复制
Private mvar行号 As Long '局部复制
Private mvar结算方式 As String '局部复制
Private mvar结算性质 As Integer
Private mvar校对标志 As Byte '局部复制
Private mvar剩余金额 As Double '局部复制
Private mvar结算号码 As String '局部复制
Private mvar结算摘要 As String '局部复制
Private mvar缴款金额 As Double '局部复制
Private mvar找补 As Double '局部复制
Private mvar误差费 As Double '局部复制
Private mvar是否转帐 As Boolean '局部复制
Private mvar消费卡 As Boolean '局部复制
Private mvar消费卡ID As Long '局部复制
Private mvar是否密文 As Boolean '局部复制
Private mvar限制类别 As String '局部复制
Private mvar帐户余额 As Double '局部复制
Private mvar密码 As String '局部复制
Private mvar是否允许编辑 As Boolean '局部复制
Private mvar是否允许删除 As Boolean '局部复制
Private mvar关联交易ID As Long '局部复制
Private mvar是否结算 As Boolean '局部复制
Private mvar是否预交 As Boolean '局部复制
Private mvar原始金额 As Double '局部复制
Private mvar未退金额 As Double '局部复制
Private mvar是否退款 As Boolean '局部复制
Private mvar是否缺省 As Boolean '局部复制
Private mvar是否允许退现 As Boolean '局部复制
Private mvar是否强制退现 As Boolean '局部复制
Private mvar是否保存 As Boolean '局部复制
Private mvar退款交易流水号 As String '局部复制
Private mvar退款交易说明 As String '局部复制
Private mvar是否脱机医保 As Boolean '局部复制
Private mvar冲销ID As Long '局部复制
Private mvar单据号 As String
Private mvar结算时间 As Date
Private mvarQRCode As String '局部复制
Private mvarTag As String '局部复制
Private mvarobjTag As Object '局部复制
Private mvar单据性质 As Integer ' 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡
Private mvar异常ID As Long

Public Property Let 异常ID(ByVal vData As Long)
    '向属性指派值时使用，位于赋值语句的左边。
    'Syntax: X.异常ID = 5
    mvar异常ID = vData
End Property
Public Property Get 异常ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.异常ID
    异常ID = mvar异常ID
End Property


Public Property Let 单据性质(ByVal vData As Integer)
    '向属性指派值时使用，位于赋值语句的左边。
    'Syntax: X.单据性质 = 5
    mvar单据性质 = vData
End Property
Public Property Get 单据性质() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.单据性质
    单据性质 = mvar单据性质
End Property



Public Property Let 单据号(ByVal vData As String)
    '向属性指派值时使用，位于赋值语句的左边。
    'Syntax: X.单据号 = 5
    mvar单据号 = vData
End Property


Public Property Get 单据号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.单据号
    单据号 = mvar单据号
End Property

Public Property Let QRCode(ByVal vData As String)
    '向属性指派值时使用，位于赋值语句的左边。
    'Syntax: X.QRCode = 5
    mvarQRCode = vData
End Property


Public Property Get QRCode() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.QRCode
    QRCode = mvarQRCode
End Property



Public Property Let 冲销ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.冲销ID = 5
    mvar冲销ID = vData
End Property


Public Property Get 冲销ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.冲销ID
    冲销ID = mvar冲销ID
End Property



Public Property Let 是否脱机医保(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否脱机医保 = 5
    mvar是否脱机医保 = vData
End Property


Public Property Get 是否脱机医保() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否脱机医保
    是否脱机医保 = mvar是否脱机医保
End Property



Public Property Let 退款交易说明(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.退款交易说明 = 5
    mvar退款交易说明 = vData
End Property


Public Property Get 退款交易说明() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.退款交易说明
    退款交易说明 = mvar退款交易说明
End Property



Public Property Let 退款交易流水号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.退款交易流水号 = 5
    mvar退款交易流水号 = vData
End Property


Public Property Get 退款交易流水号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.退款交易流水号
    退款交易流水号 = mvar退款交易流水号
End Property



 
Public Property Let 是否保存(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否保存 = 5
    mvar是否保存 = vData
End Property


Public Property Get 是否保存() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否保存
    是否保存 = mvar是否保存
End Property



Public Property Let 是否强制退现(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否强制退现 = 5
    mvar是否强制退现 = vData
End Property


Public Property Get 是否强制退现() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否强制退现
    是否强制退现 = mvar是否强制退现
End Property



Public Property Let 是否允许退现(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否允许退现 = 5
    mvar是否允许退现 = vData
End Property


Public Property Get 是否允许退现() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否允许退现
    是否允许退现 = mvar是否允许退现
End Property



Public Property Let 是否缺省(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否缺省 = 5
    mvar是否缺省 = vData
End Property


Public Property Get 是否缺省() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否缺省
    是否缺省 = mvar是否缺省
End Property



Public Property Let 是否退款(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否退款 = 5
    mvar是否退款 = vData
End Property


Public Property Get 是否退款() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否退款
    是否退款 = mvar是否退款
End Property



Public Property Let 未退金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.未退金额 = 5
    mvar未退金额 = vData
End Property


Public Property Get 未退金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.未退金额
    未退金额 = mvar未退金额
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



Public Property Let 是否预交(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否预交 = 5
    mvar是否预交 = vData
End Property


Public Property Get 是否预交() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否预交
    是否预交 = mvar是否预交
End Property



Public Property Let 结算类型(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算类型 = 5
    mvar结算类型 = vData
End Property
Public Property Get 结算类型() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算类型
      结算类型 = mvar结算类型
End Property
Public Property Let 是否结算(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否结算 = 5
    mvar是否结算 = vData
End Property
Public Property Get 是否结算() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否结算
    是否结算 = mvar是否结算
End Property
Public Property Let 关联交易ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联交易ID = 5
    mvar关联交易ID = vData
End Property


Public Property Get 关联交易ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联交易ID
    关联交易ID = mvar关联交易ID
End Property

 

Public Property Let 是否允许删除(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否允许删除 = 5
    mvar是否允许删除 = vData
End Property


Public Property Get 是否允许删除() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否允许删除
    是否允许删除 = mvar是否允许删除
End Property



Public Property Let 是否允许编辑(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否允许编辑 = 5
    mvar是否允许编辑 = vData
End Property


Public Property Get 是否允许编辑() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否允许编辑
    是否允许编辑 = mvar是否允许编辑
End Property



Public Property Let 密码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.密码 = 5
    mvar密码 = vData
End Property


Public Property Get 密码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.密码
    密码 = mvar密码
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



Public Property Let 限制类别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.限制类别 = 5
    mvar限制类别 = vData
End Property


Public Property Get 限制类别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.限制类别
    限制类别 = mvar限制类别
End Property



Public Property Let 是否密文(ByVal vData As Boolean)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.是否密文 = Form1
      mvar是否密文 = vData
End Property


Public Property Get 是否密文() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否密文
      是否密文 = mvar是否密文
End Property



Public Property Let 消费卡ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.消费卡ID = 5
    mvar消费卡ID = vData
End Property


Public Property Get 消费卡ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.消费卡ID
    消费卡ID = mvar消费卡ID
End Property



Public Property Let 消费卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.消费卡 = 5
    mvar消费卡 = vData
End Property


Public Property Get 消费卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.消费卡
    消费卡 = mvar消费卡
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



Public Property Let 找补(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.找补 = 5
    mvar找补 = vData
End Property


Public Property Get 找补() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.找补
    找补 = mvar找补
End Property



Public Property Let 缴款金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缴款金额 = 5
    mvar缴款金额 = vData
End Property


Public Property Get 缴款金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缴款金额
     缴款金额 = mvar缴款金额
End Property



Public Property Let 结算摘要(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算摘要 = 5
    mvar结算摘要 = vData
End Property


Public Property Get 结算摘要() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算摘要
    结算摘要 = mvar结算摘要
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



Public Property Set objTag(ByVal vData As Object)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objTag = Form1
    Set mvarobjTag = vData
End Property


Public Property Get objTag() As Object
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objTag
    Set objTag = mvarobjTag
End Property



Public Property Let 剩余金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.剩余金额 = 5
    mvar剩余金额 = vData
End Property


Public Property Get 剩余金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.剩余金额
    剩余金额 = mvar剩余金额
End Property



Public Property Let 结算时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算时间 = 5
    mvar结算时间 = vData
End Property

 


Public Property Get 结算时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算时间
    结算时间 = mvar结算时间
End Property

 

Public Property Let 校对标志(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.校对标志 = 5
    mvar校对标志 = vData
End Property


Public Property Get 校对标志() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.校对标志
    校对标志 = mvar校对标志
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



Public Property Let 行号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.行号 = 5
    mvar行号 = vData
End Property


Public Property Get 行号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.行号
    行号 = mvar行号
End Property

 

Public Property Let Tag(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Tag = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
End Property



Public Property Let 结算ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算ID = 5
    mvar结算ID = vData
End Property


Public Property Get 结算ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算ID
    结算ID = mvar结算ID
End Property



Public Property Let 卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡号 = 5
    mvar卡号 = vData
End Property


Public Property Get 卡号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡号
    卡号 = mvar卡号
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



Public Property Let 预交ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预交ID = 5
    mvar预交ID = vData
End Property


Public Property Get 预交ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预交ID
    预交ID = mvar预交ID
End Property



Public Property Let 是否转帐(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否转帐 = 5
    mvar是否转帐 = vData
End Property


Public Property Get 是否转帐() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否转帐
    是否转帐 = mvar是否转帐
End Property



Public Property Let 卡类别ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.卡类别ID = 5
    mvar卡类别ID = vData
End Property


Public Property Get 卡类别ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.卡类别ID
    卡类别ID = mvar卡类别ID
End Property



Public Property Set objCard(ByVal vData As Card)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objCard = Form1
    Set mvarobjCard = vData
End Property


Public Property Get objCard() As Card
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objCard
    Set objCard = mvarobjCard
End Property

Public Function Clone() As clsBalanceItem
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:克隆一个新的项
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-11-13 10:50:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As clsBalanceItem, objItem As clsBalanceItem
    
    Set objTemp = Me
    Set objItem = New clsBalanceItem
    With objItem
        .Key = objTemp.Key
        Set .objCard = objTemp.objCard.CloneNewCard
        Set .objTag = objTemp.objTag
        .QRCode = objTemp.QRCode
        .Tag = objTemp.Tag
        
        .冲销ID = objTemp.冲销ID
        .单据号 = objTemp.单据号
        .单据性质 = objTemp.单据性质
        .关联交易ID = objTemp.关联交易ID
        .交易流水号 = objTemp.交易流水号
        .交易说明 = objTemp.交易说明
        .缴款金额 = objTemp.缴款金额
        .结算ID = objTemp.结算ID
        .结算方式 = objTemp.结算方式
        .结算号码 = objTemp.结算号码
        .结算金额 = objTemp.结算金额
        .结算类型 = objTemp.结算类型
        .结算时间 = objTemp.结算时间
        .结算性质 = objTemp.结算性质
        .结算摘要 = objTemp.结算摘要
        .卡号 = objTemp.卡号
        .卡类别ID = objTemp.卡类别ID
        .密码 = objTemp.密码
        .剩余金额 = objTemp.剩余金额
        .是否保存 = objTemp.是否保存
        .是否结算 = objTemp.是否结算
        .是否密文 = objTemp.是否密文
        .是否强制退现 = objTemp.是否强制退现
        .是否缺省 = objTemp.是否缺省
        .是否退款 = objTemp.是否退款
        .是否脱机医保 = objTemp.是否脱机医保
        .是否预交 = objTemp.是否预交
        .是否允许编辑 = objTemp.是否允许编辑
        .是否允许删除 = objTemp.是否允许删除
        .是否允许退现 = objTemp.是否允许退现
        .是否转帐 = objTemp.是否转帐
        .退款交易流水号 = objTemp.退款交易流水号
        .退款交易说明 = objTemp.退款交易说明
        .未退金额 = objTemp.未退金额
        .误差费 = objTemp.误差费
        .限制类别 = objTemp.限制类别
        .消费卡 = objTemp.消费卡
        .消费卡ID = objTemp.消费卡ID
        .校对标志 = objTemp.校对标志
        .行号 = objTemp.行号
        .异常ID = objTemp.异常ID
        .预交ID = objTemp.预交ID
        .原始金额 = objTemp.原始金额
        .帐户余额 = objTemp.帐户余额
        .找补 = objTemp.找补
    End With
    Set Clone = objItem
End Function
