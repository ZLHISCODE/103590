VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegEventInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsPayInfos"
Option Explicit
'*****************************************************************************************************
'结算数据填写规则及分级
'clsRegEventInfor - |挂号信息：登记时间、就诊id等缓存值
'                   |病人信息：病人基本信息，主要是接口调用时用到
'                   |发卡卡号信息：卡号、卡类别、以及操作类型(发卡、退卡、绑卡、NUll)
'   ∟zlOneCardComLib.clsPatientInfo -|病人基本信息
'   ∟clsPayInfos - |挂号单据信息：挂号单据号、金额、以及是否有三方卡支付(提前判断是否需要调用接口)
'                   |发卡单据信息：发卡单据号以及卡费
'                   |发卡结算信息：支付方式以及接口交易信息
'       ∟clsPayInfo - |挂号结算集合(支付方式、支付类型、交易信息、成功标志)
'                      |交易信息与子集部分重复，只为方便非三方卡或医保结算方式取值
'           ∟clsSubPayInfo - |三方卡或医保交易信息(结算方式、结算金额、交易信息)
'                             |注意：在更新子集的时候clsPayInfo的支付金额自动汇总
'1.挂号费与卡费分开存储，避免结算干扰，三方卡退费时要检查挂号费和卡费是同一个三方卡支付的情况
'2.取挂号结算信息最好判断总金额是否为0，只是现在多用于三方卡检查，可以通过卡类别ID和结算类型区分
'3.取卡费结算信息要判断操作类型
'4.为支持三方卡和医保返回多种结算方式，挂号结算数据分级保存
'*****************************************************************************************************

Public Enum gPagePay
    Pay_CashPay = 0 '现金支付
    Pay_ThreePay = 1 '三方卡支付
    Pay_SquarePay = 2 '消费卡支付
    Pay_AccountPay = 3 '账户支付
    Pay_InsurePay = 4 '医保支付
    Pay_OneCard = 5 '旧一卡通
End Enum

Public Enum gCardPay
    CP_NoCard = 0 '没有发卡
    CP_仅退号
    CP_发卡
    CP_绑卡
    CP_退卡
    CP_解绑
End Enum

'保持属性值的局部变量
Private mvar正式结算 As String '局部复制
Private mvar预结算 As String '局部复制
Private mvarobjPayInfos As clsPayInfos '局部复制
Private mvarobjPati As zlOneCardComLib.clsPatientInfo  '局部复制
'病人信息
Private mvarstrYBPati As String '局部复制
Private mvarintInsure As Integer '局部复制
Private mvar实名认证 As Boolean '局部复制
Private mvar临时卡病人 As Boolean '局部复制
Private mvarcllDrug As Collection '病人过敏药物
Private mvarcllImmune As Collection '病人免疫记录
Private mvarcllCert As Collection '病人证件信息
Private mvarcllOther As Collection '病人从表信息
Private mvarcllContact As Collection '病人联系人信息
'挂号信息
Private mvar计费方式 As Integer '局部复制
Private mvar当前支付 As Double '局部复制
Private mvar就诊ID As Long '局部复制
Private mvar登记时间 As Date '局部复制
Private mvar发生时间 As String '局部复制
Private mvar诊室 As String '局部复制
Private mvar号序 As Long '局部复制
Private mvar是否急诊 As Boolean '局部复制
'发卡信息
Private mvarCard_自制卡 As Boolean '局部复制
Private mvarCard_卡类别ID As Long '局部复制
Private mvarCard_变动类型 As gCardPay '局部复制
Private mvarCard_卡号 As String '局部复制
Private mvarCard_项目ID As Long '局部复制

Public Property Set cllContact(ByVal vData As Collection)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.cllContact = Form1
    Set mvarcllContact = vData
End Property


Public Property Get cllContact() As Collection
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.cllContact
    Set cllContact = mvarcllContact
End Property



Public Property Set cllOther(ByVal vData As Collection)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.cllOther = Form1
    Set mvarcllOther = vData
End Property


Public Property Get cllOther() As Collection
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.cllOther
    Set cllOther = mvarcllOther
End Property



Public Property Set cllCert(ByVal vData As Collection)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.cllCert = Form1
    Set mvarcllCert = vData
End Property


Public Property Get cllCert() As Collection
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.cllCert
    Set cllCert = mvarcllCert
End Property


Public Property Set cllImmune(ByVal vData As Collection)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.cllImmune = Form1
    Set mvarcllImmune = vData
End Property


Public Property Get cllImmune() As Collection
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.cllImmune
    Set cllImmune = mvarcllImmune
End Property



Public Property Set cllDrug(ByVal vData As Collection)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.cllDrug = Form1
    Set mvarcllDrug = vData
End Property


Public Property Get cllDrug() As Collection
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.cllDrug
    Set cllDrug = mvarcllDrug
End Property



Public Property Let Card_项目ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_项目ID = 5
    mvarCard_项目ID = vData
End Property


Public Property Get Card_项目ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_项目ID
    Card_项目ID = mvarCard_项目ID
End Property




Public Property Let 计费方式(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.计费方式 = 5
    mvar计费方式 = vData
End Property


Public Property Get 计费方式() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.计费方式
    计费方式 = mvar计费方式
End Property




Public Property Let 号序(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.号序 = 5
    mvar号序 = vData
End Property


Public Property Get 号序() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.号序
    号序 = mvar号序
End Property



Public Property Let 诊室(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.诊室 = 5
    mvar诊室 = vData
End Property


Public Property Get 诊室() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.诊室
    诊室 = mvar诊室
End Property



Public Property Let 发生时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发生时间 = 5
    mvar发生时间 = vData
End Property


Public Property Get 发生时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.发生时间
    发生时间 = mvar发生时间
End Property

Public Sub Clear(Optional ByVal blnClearPati As Boolean = True, Optional ByVal blnClearInsure As Boolean = True)
    If blnClearPati Then
        mvar实名认证 = False
        mvar临时卡病人 = False
    End If
        
    If blnClearPati Or blnClearInsure Then
        mvarstrYBPati = ""
        mvarintInsure = 0
    End If
    mvar当前支付 = 0
    mvar就诊ID = 0
    mvar登记时间 = 0
    mvar是否急诊 = False
    mvarCard_自制卡 = False
    mvarCard_卡类别ID = 0
    mvarCard_变动类型 = CP_NoCard
    mvarCard_卡号 = ""
    mvar正式结算 = ""
    mvar预结算 = ""
    Set mvarobjPayInfos = New clsPayInfos
    Set mvarobjPati = New zlOneCardComLib.clsPatientInfo
End Sub

Public Sub ClearPati()
    Set mvarobjPati = New zlOneCardComLib.clsPatientInfo
    Set mvarcllDrug = Nothing  '病人过敏药物
    Set mvarcllImmune = Nothing '病人免疫记录
    Set mvarcllCert = Nothing  '病人证件信息
    Set mvarcllOther = Nothing '病人从表信息
    Set mvarcllContact = Nothing '病人联系人信息
    mvarstrYBPati = ""
    mvarintInsure = 0
    mvar实名认证 = False
    mvar临时卡病人 = False
End Sub


Public Property Let 登记时间(ByVal vData As Date)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
    mvar登记时间 = vData
End Property


Public Property Get 登记时间() As Date
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.登记时间
    登记时间 = mvar登记时间
End Property


Public Property Let 当前支付(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.当前支付 = 5
    mvar当前支付 = vData
End Property


Public Property Get 当前支付() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.当前支付
    当前支付 = mvar当前支付
End Property


Public Property Let 是否急诊(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否急诊 = 5
    mvar是否急诊 = vData
End Property


Public Property Get 是否急诊() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否急诊
    是否急诊 = mvar是否急诊
End Property



Public Property Let 就诊ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.就诊ID = 5
    mvar就诊ID = vData
End Property


Public Property Get 就诊ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.就诊ID
    就诊ID = mvar就诊ID
End Property


Public Property Let 临时卡病人(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.临时卡病人 = 5
    mvar临时卡病人 = vData
End Property


Public Property Get 临时卡病人() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.临时卡病人
    临时卡病人 = mvar临时卡病人
End Property



Public Property Let 实名认证(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.实名认证 = 5
    mvar实名认证 = vData
End Property


Public Property Get 实名认证() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.实名认证
    实名认证 = mvar实名认证
End Property


Public Property Let intInsure(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.intInsure = 5
    mvarintInsure = vData
End Property


Public Property Get intInsure() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.intInsure
    intInsure = mvarintInsure
End Property



Public Property Let strYBPati(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.strYBPati = 5
    mvarstrYBPati = vData
End Property


Public Property Get strYBPati() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.strYBPati
    strYBPati = mvarstrYBPati
End Property


'
'Public Property Let 病人ID(ByVal vData As Long)
''向属性指派值时使用，位于赋值语句的左边。
''Syntax: X.病人ID = 5
'    mvar病人ID = vData
'End Property
'
'
'Public Property Get 病人ID() As Long
''检索属性值时使用，位于赋值语句的右边。
''Syntax: Debug.Print X.病人ID
'    病人ID = mvar病人ID
'End Property
'
'
'
'Public Property Let 费别(ByVal vData As String)
''向属性指派值时使用，位于赋值语句的左边。
''Syntax: X.费别 = 5
'    mvar费别 = vData
'End Property
'
'
'Public Property Get 费别() As String
''检索属性值时使用，位于赋值语句的右边。
''Syntax: Debug.Print X.费别
'    费别 = mvar费别
'End Property
'
'
'
'Public Property Let 年龄(ByVal vData As String)
''向属性指派值时使用，位于赋值语句的左边。
''Syntax: X.年龄 = 5
'    mvar年龄 = vData
'End Property
'
'
'Public Property Get 年龄() As String
''检索属性值时使用，位于赋值语句的右边。
''Syntax: Debug.Print X.年龄
'    年龄 = mvar年龄
'End Property
'
'
'
'Public Property Let 性别(ByVal vData As String)
''向属性指派值时使用，位于赋值语句的左边。
''Syntax: X.性别 = 5
'    mvar性别 = vData
'End Property
'
'
'Public Property Get 性别() As String
''检索属性值时使用，位于赋值语句的右边。
''Syntax: Debug.Print X.性别
'    性别 = mvar性别
'End Property
'
'
'
'Public Property Let 姓名(ByVal vData As String)
''向属性指派值时使用，位于赋值语句的左边。
''Syntax: X.姓名 = 5
'    mvar姓名 = vData
'End Property
'
'
'Public Property Get 姓名() As String
''检索属性值时使用，位于赋值语句的右边。
''Syntax: Debug.Print X.姓名
'    姓名 = mvar姓名
'End Property



Public Property Let Card_卡类别ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_卡类别ID = 5
    mvarCard_卡类别ID = vData
End Property


Public Property Get Card_卡类别ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_卡类别ID
    Card_卡类别ID = mvarCard_卡类别ID
End Property



Public Property Let Card_变动类型(ByVal vData As gCardPay)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_变动类型 = 5
    mvarCard_变动类型 = vData
End Property


Public Property Get Card_变动类型() As gCardPay
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_变动类型
    Card_变动类型 = mvarCard_变动类型
End Property



Public Property Let Card_卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_卡号 = 5
    mvarCard_卡号 = vData
End Property


Public Property Get Card_卡号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_卡号
    Card_卡号 = mvarCard_卡号
End Property



Public Property Let Card_自制卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_自制卡 = 5
    mvarCard_自制卡 = vData
End Property


Public Property Get Card_自制卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_自制卡
    Card_自制卡 = mvarCard_自制卡
End Property



Public Property Set objPayInfos(ByVal vData As clsPayInfos)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objPayInfos = Form1
    Set mvarobjPayInfos = vData
End Property


Public Property Get objPayInfos() As clsPayInfos
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objPayInfos
    Set objPayInfos = mvarobjPayInfos
End Property


Public Property Set objPati(ByVal vData As zlOneCardComLib.clsPatientInfo)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objPati = Form1
    Set mvarobjPati = vData
End Property


Public Property Get objPati() As zlOneCardComLib.clsPatientInfo
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objPati
    Set objPati = mvarobjPati
End Property



Public Property Let 预结算(ByVal vData As String)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.预结算 = Form1
    mvar预结算 = vData
End Property


Public Property Get 预结算() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预结算
    预结算 = mvar预结算
End Property



Public Property Let 正式结算(ByVal vData As String)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.正式结算 = Form1
    mvar正式结算 = vData
End Property


Public Property Get 正式结算() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.正式结算
    正式结算 = mvar正式结算
End Property

Private Sub Class_Initialize()
    Set mvarobjPayInfos = New clsPayInfos
    Set mvarobjPati = New zlOneCardComLib.clsPatientInfo
End Sub

Private Sub Class_Terminate()
    Set mvarobjPayInfos = Nothing
    Set mvarobjPati = Nothing
End Sub
