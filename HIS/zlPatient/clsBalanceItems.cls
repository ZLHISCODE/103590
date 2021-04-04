VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBalanceItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBalanceItem"
Attribute VB_Ext_KEY = "Member0" ,"clsBalanceItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'局部变量，保存集合
Private mCol As Collection
Public Enum gBalanceType
    gEM_普通结算 = 0
    gEM_预交款 = 1
    gEM_医保 = 2
    gEM_一卡通 = 3
    gEM_一卡通_老 = 4
    gEm_消费卡 = 5
End Enum

Private mvar结算金额 As Double '局部复制
Private mvar是否转帐 As Boolean '局部复制
Private mvar退费结帐IDs As String '局部复制
Private mvar是否异常退费 As Boolean '局部复制
Private mvar类型 As gBalanceType '局部复制
Private mvar收费类型 As Byte '局部复制
Private mvar结帐ID As String '局部复制
Private mvar结算序号 As String '局部复制
Public Property Let 结算序号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算序号 = 5
    mvar结算序号 = vData
End Property


Public Property Get 结算序号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算序号
    结算序号 = mvar结算序号
End Property



Public Property Let 结帐ID(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结帐ID = 5
    mvar结帐ID = vData
End Property


Public Property Get 结帐ID() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结帐ID
    结帐ID = mvar结帐ID
End Property



Public Property Let 收费类型(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费类型 = 5
    mvar收费类型 = vData
End Property


Public Property Get 收费类型() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.收费类型
    收费类型 = mvar收费类型
End Property



Public Property Let 类型(ByVal vData As gBalanceType)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar类型 = vData
End Property


Public Property Get 类型() As gBalanceType
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.类型
    类型 = mvar类型
End Property



Public Property Let 是否异常退费(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否异常退费 = 5
    mvar是否异常退费 = vData
End Property


Public Property Get 是否异常退费() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否异常退费
    是否异常退费 = mvar是否异常退费
End Property



Public Property Let 退费结帐IDs(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.退费结帐IDs = 5
    mvar退费结帐IDs = vData
End Property


Public Property Get 退费结帐IDs() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.退费结帐IDs
    退费结帐IDs = mvar退费结帐IDs
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

Public Sub AddItem(ByVal objItem As clsBalanceItem, Optional sKey As String)
    If Len(sKey) = 0 Then
        mCol.Add objItem
    Else
        mCol.Add objItem, sKey
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As clsBalanceItem
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
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub
