VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BillDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Detail"
Option Explicit

Public Key As String
Private mvar工本费 As Boolean '局部复制
Private mvar序号 As Integer  '局部复制
Private mvar收费类别 As String '局部复制
Private mvar收费细目ID As Long '局部复制
Private mvar收费名称 As String '局部复制
Private mvar计算单位 As String '局部复制
Private mvar数次 As Single '局部复制
Private mvar标准单价 As Currency   '局部复制
Private mvar应收金额 As Currency '局部复制
Private mvar实收金额 As Currency '局部复制
Private mvar附加标志 As Byte     '局部复制
Private mvar执行部门ID As Long   '局部复制
Private mvarInComes As BillInComes '局部复制
Private mvarDetail As Detail '局部复制
Private mvar保险大类ID As Long '局部复制
Private mvar保险项目否 As Boolean '局部复制
Private mvar保险编码 As String '局部复制
Private mvar摘要 As String '局部复制


Public Property Let 摘要(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.摘要 = 5
    mvar摘要 = vData
End Property

Public Property Get 摘要() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.摘要
    摘要 = mvar摘要
End Property


Public Property Let 保险编码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.保险编码 = 5
    mvar保险编码 = vData
End Property

Public Property Get 保险编码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.保险编码
    保险编码 = mvar保险编码
End Property

Public Property Let 保险项目否(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.就诊卡号 = 5
    mvar保险项目否 = vData
End Property

Public Property Get 保险项目否() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.就诊卡号
    保险项目否 = mvar保险项目否
End Property

Public Property Let 保险大类ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.就诊卡号 = 5
    mvar保险大类ID = vData
End Property

Public Property Get 保险大类ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.就诊卡号
    保险大类ID = mvar保险大类ID
End Property

Public Property Let 收费名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费名称 = 5
    mvar收费名称 = vData
End Property

Public Property Get 收费名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.收费名称
    收费名称 = mvar收费名称
End Property

Public Property Set Detail(ByVal vData As Detail)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Detail = Form1
    Set mvarDetail = vData
End Property

Public Property Get Detail() As Detail
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.Detail
    Set Detail = mvarDetail
End Property

Public Property Set InComes(ByVal vData As BillInComes)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.InComes = Form1
    Set mvarInComes = vData
End Property

Public Property Get InComes() As BillInComes
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.InComes
    Set InComes = mvarInComes
End Property

Public Property Let 执行部门ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行部门ID = 5
    mvar执行部门ID = vData
End Property

Public Property Get 执行部门ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.执行部门ID
    执行部门ID = mvar执行部门ID
End Property

Public Property Let 附加标志(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.附加标志 = 5
    mvar附加标志 = vData
End Property

Public Property Get 附加标志() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.附加标志
    附加标志 = mvar附加标志
End Property

Public Property Let 数次(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数次 = 5
    mvar数次 = vData
End Property

Public Property Get 数次() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.数次
    数次 = mvar数次
End Property

Public Property Let 计算单位(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.计算单位 = 5
    mvar计算单位 = vData
End Property

Public Property Get 计算单位() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.计算单位
    计算单位 = mvar计算单位
End Property

Public Property Let 收费类别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费类别 = 5
    mvar收费类别 = vData
End Property

Public Property Get 收费类别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.收费类别
    收费类别 = mvar收费类别
End Property

Public Property Let 应收金额(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.应收金额 = 5
    mvar应收金额 = vData
End Property

Public Property Get 应收金额() As Currency
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.应收金额
    应收金额 = mvar应收金额
End Property

Public Property Let 标准单价(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.标准单价 = 5
    mvar标准单价 = vData
End Property

Public Property Get 标准单价() As Currency
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.标准单价
    标准单价 = mvar标准单价
End Property

Public Property Let 序号(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property

Public Property Get 序号() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.序号
    序号 = mvar序号
End Property

Public Property Let 收费细目ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费细目ID = 5
    mvar收费细目ID = vData
End Property

Public Property Get 收费细目ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.收费细目ID
    收费细目ID = mvar收费细目ID
End Property

Private Sub Class_Initialize()
    Set Me.InComes = New BillInComes
    Set Me.Detail = New Detail
End Sub

Private Sub Class_Terminate()
    Set Me.InComes = Nothing
    Set Me.Detail = Nothing
End Sub

Public Property Let 实收金额(ByVal vData As Currency)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.标准单价 = 5
    mvar实收金额 = vData
End Property

Public Property Get 实收金额() As Currency
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.标准单价
    实收金额 = mvar实收金额
End Property

Public Property Let 工本费(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.附加标志 = 5
    mvar工本费 = vData
End Property

Public Property Get 工本费() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.附加标志
    工本费 = mvar工本费
End Property

