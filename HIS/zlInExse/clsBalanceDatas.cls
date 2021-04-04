VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceDatas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBalanceData"
Attribute VB_Ext_KEY = "Member0" ,"clsBalanceData"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'局部变量，保存集合
Private mCol As Collection
'保持属性值的局部变量
Private mvar本次未结 As Double '局部复制
Private mvar当前结帐 As Double '局部复制
Private mvar已付合计 As Double '局部复制
Private mvar未付合计 As Double '局部复制
Private mvar预结算总额 As Double '局部复制
Private mvar医保支付合计 As Double '局部复制
Private mvar冲预交合计 As Double '局部复制
Private mvar结帐ID As Long '局部复制
Private mvar缴款 As Double '局部复制
Private mvar找补 As Double '局部复制
Private mvar误差费 As Double '局部复制
Private mvar现金支付 As Double '局部复制
Private mvar缴预交ID As Long '局部复制
Private mvar缴预交单据号 As String '局部复制
Private mvar冲销ID As Long '局部复制
Private mvar结帐时间 As Date '局部复制
Private mvar是否保存结帐单 As Boolean '局部复制
Private mvar结帐单据号 As String '局部复制
Private mvar是否预交验卡 As Boolean '局部复制
Private mvar病历原因 As String '局部复制
'保持属性值的局部变量
Private mvarobjInsure As clsInsureInfor '局部复制
'保持属性值的局部变量
Private mvar结算类型 As Byte '局部复制
'保持属性值的局部变量
Private mvar是否中途结帐 As Boolean '局部复制
Public Property Let 是否中途结帐(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否中途结帐 = 5
    mvar是否中途结帐 = vData
End Property


Public Property Get 是否中途结帐() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否中途结帐
    是否中途结帐 = mvar是否中途结帐
End Property



Public Property Let 结算类型(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算类型 = 5
    mvar结算类型 = vData
End Property


Public Property Get 结算类型() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算类型
    结算类型 = mvar结算类型
End Property



Public Property Set objInsure(ByVal vData As clsInsureInfor)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.objInsure = Form1
    Set mvarobjInsure = vData
End Property


Public Property Get objInsure() As clsInsureInfor
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.objInsure
    Set objInsure = mvarobjInsure
End Property




Public Property Let 病历原因(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病历原因 = 5
    mvar病历原因 = vData
End Property


Public Property Get 病历原因() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病历原因
    病历原因 = mvar病历原因
End Property



Public Property Let 是否预交验卡(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否预交验卡 = 5
    mvar是否预交验卡 = vData
End Property


Public Property Get 是否预交验卡() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否预交验卡
    是否预交验卡 = mvar是否预交验卡
End Property



Public Property Let 结帐单据号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐单据号 = 5
    mvar结帐单据号 = vData
End Property


Public Property Get 结帐单据号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐单据号
    结帐单据号 = mvar结帐单据号
End Property



Public Property Let 是否保存结帐单(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否保存结帐单 = 5
    mvar是否保存结帐单 = vData
End Property


Public Property Get 是否保存结帐单() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否保存结帐单
    是否保存结帐单 = mvar是否保存结帐单
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



Public Property Let 缴预交单据号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缴预交单据号 = 5
    mvar缴预交单据号 = vData
End Property


Public Property Get 缴预交单据号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缴预交单据号
    缴预交单据号 = mvar缴预交单据号
End Property



Public Property Let 缴预交ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缴预交ID = 5
    mvar缴预交ID = vData
End Property


Public Property Get 缴预交ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缴预交ID
    缴预交ID = mvar缴预交ID
End Property



Public Property Let 现金支付(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.现金支付 = 5
    mvar现金支付 = vData
End Property


Public Property Get 现金支付() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.现金支付
    现金支付 = mvar现金支付
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



Public Property Let 缴款(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缴款 = 5
    mvar缴款 = vData
End Property


Public Property Get 缴款() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缴款
    缴款 = mvar缴款
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



Public Property Let 冲预交合计(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.冲预交合计 = 5
    mvar冲预交合计 = vData
End Property


Public Property Get 冲预交合计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.冲预交合计
    冲预交合计 = mvar冲预交合计
End Property



Public Property Let 医保支付合计(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医保支付合计 = 5
    mvar医保支付合计 = vData
End Property


Public Property Get 医保支付合计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医保支付合计
    医保支付合计 = mvar医保支付合计
End Property



Public Property Let 预结算总额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预结算总额 = 5
    mvar预结算总额 = vData
End Property


Public Property Get 预结算总额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预结算总额
    预结算总额 = mvar预结算总额
End Property



Public Property Let 未付合计(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.未付合计 = 5
    mvar未付合计 = vData
End Property


Public Property Get 未付合计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.未付合计
    未付合计 = mvar未付合计
End Property



Public Property Let 已付合计(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.已付合计 = 5
    mvar已付合计 = vData
End Property


Public Property Get 已付合计() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.已付合计
    已付合计 = mvar已付合计
End Property



Public Property Let 当前结帐(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.当前结帐 = 5
    mvar当前结帐 = vData
End Property


Public Property Get 当前结帐() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.当前结帐
    当前结帐 = mvar当前结帐
End Property



Public Property Let 本次未结(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.本次未结 = 5
    mvar本次未结 = vData
End Property


Public Property Get 本次未结() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.本次未结
     本次未结 = mvar本次未结
End Property
Public Sub AddItem(objBalanceData As clsBalanceData, Optional sKey As String)
    '创建新对象
    Dim objNewMember As clsBalanceData
    Set objNewMember = New clsBalanceData

    '设置传入方法的属性
    Set objNewMember = objBalanceData
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

End Sub




Public Function Add(objBalanceItems As clsBalanceItems, 结算总额 As Double, Optional sKey As String) As clsBalanceData
    '创建新对象
    Dim objNewMember As clsBalanceData
    Set objNewMember = New clsBalanceData

    '设置传入方法的属性
    Set objNewMember.objBalanceItems = objBalanceItems
    objNewMember.结算总额 = 结算总额
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As clsBalanceData
Attribute Item.VB_UserMemId = 0
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


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
    Set mvarobjInsure = New clsInsureInfor
    
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
    Set mvarobjInsure = Nothing
End Sub

