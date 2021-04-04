VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsPayRegInfo"
Attribute VB_Ext_KEY = "Member0" ,"clsPayRegInfo"
Option Explicit
Public Enum gPagePay
    Pay_CashPay = 0 '现金支付
    Pay_ThreePay = 1 '三方卡支付
    Pay_SquarePay = 2 '消费卡支付 -暂时无效
    Pay_AccountPay = 3 '账户支付
    Pay_InsurePay = 4 '医保支付
    Pay_OneCard = 5 '老一卡通
End Enum

'局部变量，保存集合
Private mCol As Collection
'仅挂号支付信息
Private mvar名称 As String '局部复制
Private mvar接口序号 As Long '局部复制
Private mvar消费卡 As Boolean '局部复制
Private mvar校对标志 As Integer '局部复制
Private mvar支付类型 As gPagePay '局部复制
Private mvar卡号 As String '局部复制
Private mvar支付密码 As String '局部复制
Private mvar结算性质 As Integer '局部复制
Private mvar消费卡ID As Long '局部复制
Private mvar结算成功 As Boolean '局部复制
Private mvar关联交易ID As Long '局部复制
Private mvar固定 As Boolean '局部复制
Private mvar独立结算 As Boolean '局部复制
Private mvarQRCode As String '局部复制
Private mvarCardPro As Card '局部复制

'下列信息与子集重复，只为方便其他结算方式取值
Private mvar结算方式 As String '局部复制
Private mvar支付金额 As Double '局部复制
Private mvar结算号码 As String '局部复制
Private mvar交易流水号 As String '局部复制
Private mvar交易说明 As String '局部复制
Private mvarPayRow As Long '局部复制

Public Property Let 独立结算(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.独立结算 = 5
    mvar独立结算 = vData
End Property


Public Property Get 独立结算() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.独立结算
    独立结算 = mvar独立结算
End Property




Public Property Let 固定(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.固定 = 5
    mvar固定 = vData
End Property


Public Property Get 固定() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.固定
    固定 = mvar固定
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



Public Property Let 支付金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.支付金额 = 5
    mvar支付金额 = vData
End Property


Public Property Get 支付金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.支付金额
    If mCol.Count > 1 Then
        Dim i As Integer, dbl金额 As Double, objSubPay As clsSubPayInfo
        For i = 1 To mCol.Count
            Set objSubPay = mCol(i)
            dbl金额 = dbl金额 + objSubPay.结算金额
        Next
        mvar支付金额 = dbl金额
    End If
    支付金额 = mvar支付金额
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




Public Property Let 结算成功(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算成功 = 5
    mvar结算成功 = vData
End Property


Public Property Get 结算成功() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算成功
    结算成功 = mvar结算成功
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



Public Property Let 支付密码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.支付密码 = 5
    mvar支付密码 = vData
End Property


Public Property Get 支付密码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.支付密码
    支付密码 = mvar支付密码
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



Public Property Let 支付类型(ByVal vData As gPagePay)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.支付类型 = Form1
    mvar支付类型 = vData
End Property


Public Property Get 支付类型() As gPagePay
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.支付类型
    支付类型 = mvar支付类型
End Property



Public Property Let 校对标志(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.校对标志 = 5
    mvar校对标志 = vData
End Property


Public Property Get 校对标志() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.校对标志
    校对标志 = mvar校对标志
End Property



Public Property Set CardPro(ByVal vData As Card)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.CardPro = Form1
    Set mvarCardPro = vData
End Property


Public Property Get CardPro() As Card
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CardPro
    Set CardPro = mvarCardPro
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



Public Property Let 接口序号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.接口序号 = 5
    mvar接口序号 = vData
End Property


Public Property Get 接口序号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.接口序号
    接口序号 = mvar接口序号
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


Public Sub AddItem(ByVal objPayRegInfo As clsSubPayInfo, Optional ByVal sKey As String)
    If Len(sKey) = 0 Then
        mCol.Add objPayRegInfo
    Else
        mCol.Add objPayRegInfo, sKey
    End If
End Sub


Public Property Get Item(vntIndexKey As Variant) As clsSubPayInfo
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Sub ReMoveAll()
    Set mCol = New Collection
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub

