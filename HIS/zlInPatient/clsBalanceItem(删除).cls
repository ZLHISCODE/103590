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
'  01:Items对象操作相关
'    0101. zlCopyNewItemFromBalanceItem:复制一个新的Item对象
'    0102. zlCopyNewCardFromCard:根据一个卡对象，复制为新的卡对象
Public Key As String
'保持属性值的局部变量
Private mvarobjCard As Card '局部复制
Private mvar卡类别ID As Long '局部复制
Private mvar是否转帐 As Boolean '局部复制
Private mvar预交ID As Long '局部复制
Private mvar交易流水号 As String '局部复制
Private mvar交易说明 As String '局部复制
Private mvar结算金额 As Double '局部复制
Private mvar卡号 As String '局部复制
Private mvar结帐ID As Long '局部复制
Private mvarTag As String '局部复制
Private mvar是否退款分交易 As Boolean '局部复制
Private mvar行号 As Long '局部复制
Private mvar结算方式 As String '局部复制
Private mvar校对标志 As Byte '局部复制
Private mvar门诊结帐 As Boolean '局部复制
Private mvar结帐时间 As Date '局部复制
Private mvar剩余金额 As Double '局部复制
Private mvarobjTag As Object '局部复制
Private mvar结算号码 As String '局部复制
Private mvar结算摘要 As String '局部复制
Private mvar缴款金额 As Double '局部复制
Private mvar找补 As Double '局部复制
Private mvar误差费 As Double '局部复制
Private mvar消费卡 As Boolean '局部复制
Private mvar消费卡ID As Long '局部复制
Private mvar是否密文 As Boolean '局部复制
Private mvar限制类别 As String '局部复制
Private mvar帐户余额 As Double '局部复制
Private mvar密码 As String '局部复制
Private mvar是否允许编辑 As Boolean '局部复制
Private mvar是否允许删除 As Boolean '局部复制
Private mvar结算IDs As String '局部复制
Private mvar关联交易ID As Long '局部复制
Private mvar是否结算 As Boolean '局部复制
Private mvar结算类型 As Integer '局部复制
Private mvar是否预交 As Boolean '局部复制
Private mvar原始金额 As Double '局部复制
Private mvar结算性质 As Integer '局部复制
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
Private mvar单据号 As String '局部复制

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



Public Property Let 结算IDs(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算IDs = 5
    mvar结算IDs = vData
End Property


Public Property Get 结算IDs() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算IDs
    结算IDs = mvar结算IDs
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



Public Property Let 结帐时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐时间 = 5
    mvar结帐时间 = vData
End Property

 


Public Property Get 结帐时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐时间
    结帐时间 = mvar结帐时间
End Property



Public Property Let 门诊结帐(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.门诊结帐 = 5
    mvar门诊结帐 = vData
End Property


Public Property Get 门诊结帐() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.门诊结帐
    门诊结帐 = mvar门诊结帐
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



Public Property Let 是否退款分交易(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否退款分交易 = 5
    mvar是否退款分交易 = vData
End Property


Public Property Get 是否退款分交易() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否退款分交易
    是否退款分交易 = mvar是否退款分交易
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



Public Property Let 结帐ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐ID = 5
    mvar结帐ID = vData
End Property


Public Property Get 结帐ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐ID
    结帐ID = mvar结帐ID
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




Public Function zlCopyNewItemFromBalanceItem(ByVal objOldItem As clsBalanceItem) As clsBalanceItem
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:复制一个新的Item
    '入参:objOldItem-旧的Item对象
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-04-19 14:14:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As clsBalanceItem
    
    
    On Error GoTo errHandle
    Set objItem = New clsBalanceItem
    If objOldItem Is Nothing Then
        Set objItem.objCard = New Card
        Set zlCopyNewItemFromBalanceItem = objItem: Exit Function
    End If
    
    With objItem
        Set .objCard = zlCopyNewCardFromCard(objOldItem.objCard)
        .Key = objOldItem.Key
        .Tag = objOldItem.Tag
        .关联交易ID = objOldItem.关联交易ID
        .交易流水号 = objOldItem.交易流水号
        .交易说明 = objOldItem.交易说明
        .缴款金额 = objOldItem.缴款金额
        .结算IDs = objOldItem.结算IDs
        
        .结算方式 = objOldItem.结算方式
        .结算号码 = objOldItem.结算号码
        .结算金额 = objOldItem.结算金额
        .结算类型 = objOldItem.结算类型
        .结算性质 = objOldItem.结算性质
        .结算摘要 = objOldItem.结算摘要
        .结帐ID = objOldItem.结帐ID
        .结帐时间 = objOldItem.结帐时间
        .冲销ID = objOldItem.冲销ID
        
        .卡号 = objOldItem.卡号
        .交易流水号 = objOldItem.交易流水号
        .交易说明 = objOldItem.交易说明
        .卡类别ID = objOldItem.卡类别ID
        .密码 = objOldItem.密码
        .是否结算 = objOldItem.是否结算
        .是否密文 = objOldItem.是否密文
        .是否缺省 = objOldItem.是否缺省
        .是否退款 = objOldItem.是否退款
        .是否预交 = objOldItem.是否预交
        .是否允许编辑 = objOldItem.是否允许编辑
        .是否允许删除 = objOldItem.是否允许删除
        .是否允许退现 = objOldItem.是否允许退现
        .是否保存 = objOldItem.是否保存
        .未退金额 = objOldItem.未退金额
        .误差费 = objOldItem.误差费
        .限制类别 = objOldItem.限制类别
        .消费卡 = objOldItem.消费卡
        .消费卡ID = objOldItem.消费卡ID
        .校对标志 = objOldItem.校对标志
        .原始金额 = objOldItem.原始金额
        .帐户余额 = objOldItem.帐户余额
        .退款交易流水号 = objOldItem.退款交易流水号
        .退款交易说明 = objOldItem.退款交易说明
        .找补 = objOldItem.找补
        .行号 = objOldItem.行号
        .预交ID = objOldItem.预交ID
        .是否脱机医保 = objOldItem.是否脱机医保
        .单据号 = objOldItem.单据号
        Set .objTag = Nothing
    End With

    Set zlCopyNewItemFromBalanceItem = objItem
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     Set zlCopyNewItemFromBalanceItem = objItem
End Function

Public Function zlCopyNewCardFromCard(ByVal objOldCard As Card) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一个卡对象，复制为新的卡对象
    '入参:objOldCard-旧卡
    '返回:返回新的Card对象
    '编制:刘兴洪
    '日期:2018-04-19 14:25:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Set objCard = New Card
    If objOldCard Is Nothing Then Set zlCopyNewCardFromCard = Nothing: Exit Function
    
    On Error GoTo errHandle
    With objOldCard
        objCard.备注 = .备注
        objCard.短名 = .短名
        objCard.功能键 = .功能键
        objCard.接口编码 = .接口编码
        objCard.接口程序名 = .接口程序名
        objCard.接口序号 = .接口序号
        objCard.结算方式 = .结算方式
        objCard.结算性质 = .结算性质
        objCard.卡号密文规则 = .卡号密文规则
        objCard.卡号长度 = .卡号长度
        objCard.卡号重复使用 = .卡号重复使用
        objCard.可否设置 = .可否设置
        objCard.快键 = .快键
        objCard.密码规则 = .密码规则
        objCard.密码输入限制 = .密码输入限制
        objCard.密码长度 = .密码长度
        objCard.密码长度限制 = .密码长度限制
        objCard.名称 = .名称
        objCard.模糊查找项 = .模糊查找项
        objCard.启用 = .启用
        objCard.前缀文本 = .前缀文本
        objCard.缺省标志 = .缺省标志
        objCard.设备是否启用回车 = .设备是否启用回车
        objCard.是否持卡消费 = .是否持卡消费
        objCard.是否存在帐户 = .是否存在帐户
        objCard.是否发卡 = .是否发卡
        objCard.是否非接触式读卡 = .是否非接触式读卡
        objCard.是否接触式读卡 = .是否接触式读卡
        objCard.是否模糊查找 = .是否模糊查找
        objCard.是否全退 = .是否全退
        objCard.是否缺省密码 = .是否缺省密码
        objCard.是否扫描 = .是否扫描
        objCard.是否刷卡 = .是否刷卡
        objCard.是否退款验卡 = .是否退款验卡
        objCard.是否退现 = .是否退现
        objCard.是否写卡 = .是否写卡
        objCard.是否严格控制 = .是否严格控制
        objCard.是否证件 = .是否证件
        objCard.是否制卡 = .是否制卡
        objCard.是否转帐及代扣 = .是否转帐及代扣
        objCard.是否自动读取 = .是否自动读取
        objCard.特定项目 = .特定项目
        objCard.图像标识 = .图像标识
        objCard.系统 = .系统
        objCard.消费卡 = .消费卡
        objCard.支付启用 = .支付启用
        objCard.支付图像标识 = .支付图像标识
        objCard.自动读取间隔 = .自动读取间隔
        objCard.自制卡 = .自制卡
    End With
    Set zlCopyNewCardFromCard = objCard
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    zlCopyNewCardFromCard = objCard
End Function

