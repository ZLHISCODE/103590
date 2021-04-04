VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPayInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsPayInfo"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsPayInfo"
Option Explicit
'当前收费退费结算
'局部变量，保存集合
Private mCol As Collection
'保持属性值的局部变量
Private mvar预交金 As Double

Private mvarReg_单据号 As String '局部复制
Private mvarReg_结算金额 As Double '局部复制
Private mvarReg_结帐ID As Long '局部复制
Private mvarReg_Hav三方卡 As Boolean '挂号费的结算方式是否有三方卡，方便提前判断，不用在子集中查找
Private mvarReg_Hav消费卡 As Boolean
Private mvarReg_电子票据 As Boolean '局部复制

'卡费的支付方式唯一，所以在这里列出来
Private mvarCard_单据号 As String '局部复制
Private mvarCard_结算金额 As Double '局部复制
Private mvarCard_结帐ID As Long '局部复制
Private mvarCard_结算方式 As String '局部复制
Private mvarCard_卡类别ID As Long '局部复制
Private mvarCard_消费卡 As Boolean '局部复制
Private mvarCard_消费卡ID As Long '局部复制
Private mvarCard_支付卡号 As String '局部复制
Private mvarCard_支付密码 As String '局部复制
Private mvarCard_交易流水号 As String '局部复制
Private mvarCard_交易说明 As String '局部复制
Private mvarCard_关联交易ID As Long '局部复制
Private mvarCard_CardPro As Card '局部复制
Private mvarCard_校对标志 As Integer '局部复制
Private mvarCard_结算成功 As Boolean '局部复制
Private mvarCard_结算号码 As String '局部复制
Private mvarCard_电子票据 As Boolean '局部复制
Private mvarCard_PayRow As Long '局部复制

Public Property Let Card_电子票据(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_电子票据 = 5
    mvarCard_电子票据 = vData
End Property


Public Property Get Card_电子票据() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_电子票据
    Card_电子票据 = mvarCard_电子票据
End Property



Public Property Let Reg_电子票据(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Reg_电子票据 = 5
    mvarReg_电子票据 = vData
End Property


Public Property Get Reg_电子票据() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Reg_电子票据
    Reg_电子票据 = mvarReg_电子票据
End Property
Public Property Let Card_PayRow(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_PayRow = 5
    mvarCard_PayRow = vData
End Property


Public Property Get Card_PayRow() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_PayRow
    Card_PayRow = mvarCard_PayRow
End Property




Public Property Let Card_结算号码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_结算号码 = 5
    mvarCard_结算号码 = vData
End Property


Public Property Get Card_结算号码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_结算号码
    Card_结算号码 = mvarCard_结算号码
End Property




Public Property Let Reg_Hav消费卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Reg_Hav消费卡 = 5
    mvarReg_Hav消费卡 = vData
End Property


Public Property Get Reg_Hav消费卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Reg_Hav消费卡
    Reg_Hav消费卡 = mvarReg_Hav消费卡
End Property




Public Property Let Card_消费卡ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_消费卡ID = 5
    mvarCard_消费卡ID = vData
End Property


Public Property Get Card_消费卡ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_消费卡ID
    Card_消费卡ID = mvarCard_消费卡ID
End Property


Public Property Let Card_结算成功(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_结算成功 = 5
    mvarCard_结算成功 = vData
End Property


Public Property Get Card_结算成功() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_结算成功
    Card_结算成功 = mvarCard_结算成功
End Property



Public Property Let Card_校对标志(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_校对标志 = 5
    mvarCard_校对标志 = vData
End Property


Public Property Get Card_校对标志() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_校对标志
    Card_校对标志 = mvarCard_校对标志
End Property





Public Property Let Reg_Hav三方卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Reg_Hav三方卡 = 5
    mvarReg_Hav三方卡 = vData
End Property


Public Property Get Reg_Hav三方卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Reg_Hav三方卡
    Reg_Hav三方卡 = mvarReg_Hav三方卡
End Property



Public Property Set Card_CardPro(ByVal vData As Card)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Card_CardPro = Form1
    Set mvarCard_CardPro = vData
End Property


Public Property Get Card_CardPro() As Card
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_CardPro
    Set Card_CardPro = mvarCard_CardPro
End Property



Public Property Let Card_关联交易ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_关联交易ID = 5
    mvarCard_关联交易ID = vData
End Property


Public Property Get Card_关联交易ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_关联交易ID
    Card_关联交易ID = mvarCard_关联交易ID
End Property



Public Property Let Card_交易说明(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_交易说明 = 5
    mvarCard_交易说明 = vData
End Property


Public Property Get Card_交易说明() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_交易说明
    Card_交易说明 = mvarCard_交易说明
End Property



Public Property Let Card_交易流水号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_交易流水号 = 5
    mvarCard_交易流水号 = vData
End Property


Public Property Get Card_交易流水号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_交易流水号
    Card_交易流水号 = mvarCard_交易流水号
End Property



Public Property Let Card_支付密码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_支付密码 = 5
    mvarCard_支付密码 = vData
End Property


Public Property Get Card_支付密码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_支付密码
    Card_支付密码 = mvarCard_支付密码
End Property



Public Property Let Card_支付卡号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_支付卡号 = 5
    mvarCard_支付卡号 = vData
End Property


Public Property Get Card_支付卡号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_支付卡号
    Card_支付卡号 = mvarCard_支付卡号
End Property



Public Property Let Card_消费卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_消费卡 = 5
    mvarCard_消费卡 = vData
End Property


Public Property Get Card_消费卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_消费卡
    Card_消费卡 = mvarCard_消费卡
End Property



Public Property Let Card_卡类别ID(ByVal vData As Long)
Attribute Card_卡类别ID.VB_Description = "卡费支付卡类别ID"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_卡类别ID = 5
    mvarCard_卡类别ID = vData
End Property


Public Property Get Card_卡类别ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_卡类别ID
    Card_卡类别ID = mvarCard_卡类别ID
End Property




Public Property Let 预交金(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预交金 = 5
    mvar预交金 = vData
End Property


Public Property Get 预交金() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预交金
    预交金 = mvar预交金
End Property


Public Property Let Card_结算方式(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_结算方式 = 5
    mvarCard_结算方式 = vData
End Property


Public Property Get Card_结算方式() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_结算方式
    Card_结算方式 = mvarCard_结算方式
End Property


Public Property Let Card_结帐ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_结帐ID = 5
    mvarCard_结帐ID = vData
End Property


Public Property Get Card_结帐ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_结帐ID
    Card_结帐ID = mvarCard_结帐ID
End Property



Public Property Let Card_结算金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_结算金额 = 5
    mvarCard_结算金额 = vData
End Property


Public Property Get Card_结算金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_结算金额
    Card_结算金额 = mvarCard_结算金额
End Property



Public Property Let Card_单据号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Card_单据号 = 5
    mvarCard_单据号 = vData
End Property


Public Property Get Card_单据号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Card_单据号
    Card_单据号 = mvarCard_单据号
End Property


Public Property Let Reg_结帐ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Reg_结帐ID = 5
    mvarReg_结帐ID = vData
End Property


Public Property Get Reg_结帐ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Reg_结帐ID
    Reg_结帐ID = mvarReg_结帐ID
End Property



Public Property Let Reg_结算金额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Reg_结算金额 = 5
    mvarReg_结算金额 = vData
End Property


Public Property Get Reg_结算金额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Reg_结算金额
    Reg_结算金额 = mvarReg_结算金额
End Property



Public Property Let Reg_单据号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Reg_单据号 = 5
    mvarReg_单据号 = vData
End Property


Public Property Get Reg_单据号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Reg_单据号
    Reg_单据号 = mvarReg_单据号
End Property


Public Sub AddItem(ByVal objPayInfo As clsPayInfo, Optional ByVal sKey As String)
    If Len(sKey) = 0 Then
        mCol.Add objPayInfo
    Else
        mCol.Add objPayInfo, sKey
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As clsPayInfo
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Dim i As Integer
    If UCase(Left(vntIndexKey, 3)) = "PAY" Then
        For i = 1 To mCol.Count
            If mCol(i).支付类型 = Val(Mid(vntIndexKey, 4)) Then
                Set Item = mCol(i): Exit For
            End If
        Next
    Else
        Set Item = mCol(vntIndexKey)
    End If
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
    '删除集合中的所有元素时使用。

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


