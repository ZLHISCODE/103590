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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBalanceItem"
Attribute VB_Ext_KEY = "Member0" ,"clsBalanceItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Tag As String

'局部变量，保存集合
Private mCol As Collection
Public Enum gBalanceType
    gEM_普通结算 = 0
    gEM_预交款 = 1
    gEM_医保 = 2
    gEM_一卡通 = 3
    gEM_一卡通_老 = 4
    gEM_消费卡 = 5
    gEM_记帐单 = 6
    gEM_划价单 = 7
End Enum
Private mvarobjTag As Object '局部复制
Private mvar结算金额 As Double '局部复制
Private mvar原始金额 As Double '局部复制
Private mvar未退金额 As Double '局部复制
Private mvar误差费 As Double
Private mvar是否转帐 As Boolean '局部复制
Private mvar退费结帐IDs As String '局部复制
Private mvar是否异常退费 As Boolean '局部复制
Private mvar类型 As gBalanceType '局部复制
Private mvar收费类型 As Byte '局部复制
Private mvar单据号 As String '局部复制
Private mvar是否保存 As Boolean '局部复制
Private mvar同步状态 As Integer '局部复制
Private mvar异常ID As Long '局部复制
Private mvar业务ID As Long '局部复制
Private mvar结算完成 As Boolean '局部复制
Private mvar结算时间 As String '局部复制
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


Public Property Let 结算时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结算时间 = 5
    mvar结算时间 = vData
End Property
Public Property Get 结算时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结算时间
    结算时间 = mvar结算时间
End Property

Public Property Let 结算完成(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否结算完成 = 5
    mvar结算完成 = vData
End Property
Public Property Get 结算完成() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否结算完成
    结算完成 = mvar结算完成
End Property


Public Property Let 业务ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.业务ID = 5
    mvar业务ID = vData
End Property
Public Property Get 业务ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.业务ID
    业务ID = mvar业务ID
End Property


Public Property Let 同步状态(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.同步状态 = 5
    mvar同步状态 = vData
End Property
Public Property Get 同步状态() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.同步状态
    同步状态 = mvar同步状态
End Property

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


Public Function Clone(Optional blnSubItem As Boolean = True) As clsBalanceItems
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:克隆一个新的结算信息集
    '入参:blnSubItem-是否需要复制子项
    '编制:刘兴洪
    '日期:2019-11-13 10:51:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItems As New clsBalanceItems
    Dim objItem As clsBalanceItem, i As Long
    With objItems
        .类型 = Me.类型
        .结算金额 = Me.结算金额
        .未退金额 = Me.未退金额
        
        .是否异常退费 = Me.是否异常退费
        .是否转帐 = Me.是否转帐
        .收费类型 = Me.收费类型
        .退费结帐IDs = Me.退费结帐IDs
        .误差费 = Me.误差费
        .单据号 = Me.单据号
        .业务ID = Me.业务ID
        .异常ID = Me.异常ID
        .同步状态 = Me.同步状态
        .结算完成 = Me.结算完成
        .是否保存 = Me.是否保存
        .结算时间 = Me.结算时间
        .原始金额 = Me.原始金额
        
        .Tag = Me.Tag
        Set .objTag = Me.objTag
    End With
    If blnSubItem = False Then Set Clone = objItems: Exit Function
    For i = 1 To Me.Count
        Set objItem = Me(i).Clone
        objItems.AddItem objItem
    Next
    Set Clone = objItems
End Function
Public Sub CloneItemsPropertyByItems(ByVal objSourceItems As clsBalanceItems)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:克隆Items属性值,其他子项不发生变化
    '入参:
    '编制:刘兴洪
    '日期:2019-11-29 20:08:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    If objSourceItems Is Nothing Then Exit Sub
    With Me
        .类型 = objSourceItems.类型
        .结算金额 = 0
        .是否异常退费 = objSourceItems.是否异常退费
        .是否转帐 = objSourceItems.是否转帐
        .收费类型 = objSourceItems.收费类型
        .退费结帐IDs = objSourceItems.退费结帐IDs
        .误差费 = objSourceItems.误差费
        .单据号 = objSourceItems.单据号
        .业务ID = objSourceItems.业务ID
        .异常ID = objSourceItems.异常ID
        .同步状态 = objSourceItems.同步状态
        .结算完成 = objSourceItems.结算完成
        .是否保存 = objSourceItems.是否保存
        .结算时间 = objSourceItems.结算时间
        .未退金额 = objSourceItems.未退金额
        
        .Tag = objSourceItems.Tag
        Set objTag = objSourceItems.objTag
        For i = 1 To Me.Count
            .结算金额 = .结算金额 + Me(i).结算金额
            .未退金额 = .结算金额 + Me(i).未退金额
            .原始金额 = .原始金额 + Me(i).原始金额
        Next
    End With
End Sub
