VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit '要求变量声明

Public Key As String

'保持属性值的局部变量
Private mvarID As Long '局部复制
Private mvar药名ID As Long '局部复制
Private mvar名称 As String '局部复制
Private mvar商品名 As String '局部复制
Private mvar类别名称 As String '局部复制
Private mvar编码 As String '局部复制
Private mvar简码 As String '局部复制
Private mvar别名 As String '局部复制
Private mvar规格 As String '局部复制
Private mvar计算单位 As String '局部复制
Private mvar说明 As String '局部复制
Private mvar屏蔽费别 As Boolean '局部复制
Private mvar变价 As Boolean '局部复制
Private mvar分批 As Boolean '局部复制
Private mvar加班加价 As Boolean '局部复制
Private mvar补充摘要 As Boolean '局部复制
Private mvar执行科室 As Byte '局部复制
Private mvar固有从属 As Byte   '局部复制
Private mvar从项数次 As Double   '局部复制
Private mvar类别 As String '局部复制
Private mvar库存 As Double   '局部复制
Private mvar类型 As String
Private mvar价格 As Double   '局部复制,所有收入项目不计扣率值
Private mvar处方限量 As Double
Private mvar处方职务 As String
Private mvar服务对象 As Byte '局部复制
Private mvar药房包装 As Double '局部复制
Private mvar药房单位 As String '局部复制
Private mvar跟踪在用 As Boolean '局部复制
Private mvar要求审批 As Boolean '局部复制
Private mvar录入限量 As Double
Private mvar中药形态 As Integer       '局部复制
Private mvar批次 As Long         '局部复制
Private mvar商品条码 As String         '局部复制
Private mvar内部条码 As String         '局部复制

Public Property Let 录入限量(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar录入限量 = vData
End Property

Public Property Get 录入限量() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类型
    录入限量 = mvar录入限量
End Property

Public Property Let 要求审批(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.要求审批 = 5
    mvar要求审批 = vData
End Property


Public Property Get 要求审批() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.要求审批
    要求审批 = mvar要求审批
End Property




Public Property Let 药名ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.药名ID = 5
    mvar药名ID = vData
End Property

Public Property Get 药名ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.药名ID
    药名ID = mvar药名ID
End Property

Public Property Let 分批(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.分批 = 5
    mvar分批 = vData
End Property

Public Property Get 分批() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.分批
    分批 = mvar分批
End Property

Public Property Let 处方职务(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar处方职务 = vData
End Property

Public Property Get 处方职务() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类型
    处方职务 = mvar处方职务
End Property

Public Property Let 处方限量(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar处方限量 = vData
End Property

Public Property Get 处方限量() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类型
    处方限量 = mvar处方限量
End Property

Public Property Let 价格(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar价格 = vData
End Property

Public Property Get 价格() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类型
    价格 = mvar价格
End Property

Public Property Let 类型(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar类型 = vData
End Property

Public Property Get 类型() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类型
    类型 = mvar类型
End Property

Public Property Let 从项数次(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar从项数次 = vData
End Property

Public Property Get 从项数次() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    从项数次 = mvar从项数次
End Property

Public Property Let 固有从属(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar固有从属 = vData
End Property

Public Property Get 固有从属() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    固有从属 = mvar固有从属
End Property

Public Property Let 类别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar类别 = vData
End Property

Public Property Get 类别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    类别 = mvar类别
End Property

Public Property Let 执行科室(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.执行科室 = 5
    mvar执行科室 = vData
End Property

Public Property Get 执行科室() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.执行科室
    执行科室 = mvar执行科室
End Property

Public Property Let 加班加价(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.加班加价 = 5
    mvar加班加价 = vData
End Property

Public Property Get 加班加价() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.加班加价
    加班加价 = mvar加班加价
End Property

Public Property Let 变价(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.变价 = 5
    mvar变价 = vData
End Property

Public Property Get 变价() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.变价
    变价 = mvar变价
End Property

Public Property Let 屏蔽费别(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.屏蔽费别 = 5
    mvar屏蔽费别 = vData
End Property

Public Property Get 屏蔽费别() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.屏蔽费别
    屏蔽费别 = mvar屏蔽费别
End Property

Public Property Let 说明(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.说明 = 5
    mvar说明 = vData
End Property

Public Property Get 说明() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.说明
    说明 = mvar说明
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

Public Property Let 规格(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.规格 = 5
    mvar规格 = vData
End Property

Public Property Get 规格() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.规格
    规格 = mvar规格
End Property

Public Property Let 别名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.别名 = 5
    mvar别名 = vData
End Property

Public Property Get 别名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.别名
    别名 = mvar别名
End Property

Public Property Let 简码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.简码 = 5
    mvar简码 = vData
End Property

Public Property Get 简码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.简码
    简码 = mvar简码
End Property

Public Property Let 编码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.编码 = 5
    mvar编码 = vData
End Property

Public Property Get 编码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.编码
    编码 = mvar编码
End Property

Public Property Let 名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.名称 = 5
    mvar名称 = vData
End Property

Public Property Get 名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.名称
    名称 = mvar名称
End Property
Public Property Let 商品名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.名称 = 5
    mvar商品名 = vData
End Property

Public Property Get 商品名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.名称
    商品名 = mvar商品名
End Property
Public Property Let ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Property Get ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.ID
    ID = mvarID
End Property

Public Property Let 类别名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.名称 = 5
    mvar类别名称 = vData
End Property

Public Property Get 类别名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.名称
    类别名称 = mvar类别名称
End Property

Public Property Let 库存(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.收费细目ID = 5
    mvar库存 = vData
End Property

Public Property Get 库存() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.收费细目ID
    库存 = mvar库存
End Property

Public Property Let 补充摘要(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.补充摘要 = 5
    mvar补充摘要 = vData
End Property

Public Property Get 补充摘要() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.补充摘要
    补充摘要 = mvar补充摘要
End Property

Public Property Let 服务对象(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.服务对象 = 5
    mvar服务对象 = vData
End Property

Public Property Get 服务对象() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.服务对象
    服务对象 = mvar服务对象
End Property

Public Property Let 药房包装(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.药房包装 = 5
    mvar药房包装 = vData
End Property

Public Property Get 药房包装() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.药房包装
    药房包装 = mvar药房包装
End Property

Public Property Let 药房单位(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.药房单位 = 5
    mvar药房单位 = vData
End Property

Public Property Get 药房单位() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.药房单位
    药房单位 = mvar药房单位
End Property

Public Property Let 跟踪在用(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.跟踪在用 = 5
    mvar跟踪在用 = vData
End Property

Public Property Get 跟踪在用() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.跟踪在用
    跟踪在用 = mvar跟踪在用
End Property

Public Property Let 中药形态(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar中药形态 = vData
End Property

Public Property Get 中药形态() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    中药形态 = mvar中药形态
End Property

'------------------------------------------------------------------------
'备用卫材专用属性
Public Property Let 批次(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar批次 = vData
End Property

Public Property Get 批次() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    批次 = mvar批次
End Property

Public Property Let 商品条码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar商品条码 = vData
End Property

Public Property Get 商品条码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    商品条码 = mvar商品条码
End Property
Public Property Let 内部条码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类别 = 5
    mvar内部条码 = vData
End Property

Public Property Get 内部条码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: 'debug.print X.类别
    内部条码 = mvar内部条码
End Property

 

