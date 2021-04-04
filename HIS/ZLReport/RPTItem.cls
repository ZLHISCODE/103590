VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvarID As Integer  '局部复制
Private mvar格式号 As Byte  '局部复制
Private mvar名称 As String
Private mvar上级ID As Integer  '局部复制
Private mvar序号 As Integer '局部复制
Private mvar类型 As Byte '局部复制
Private mvar内容 As String '局部复制
Private mvar图片 As StdPicture
Private mvar参照 As String
Private mvar性质 As Byte
Private mvar表头 As String '局部复制
Private mvarX As Long '局部复制
Private mvarY As Long '局部复制
Private mvarW As Long '局部复制
Private mvarH As Long '局部复制
Private mvar行高 As Long '局部复制
Private mvar对齐 As Byte '局部复制
Private mvar自调 As Boolean
Private mvar字体 As String '局部复制
Private mvar字号 As Single '局部复制
Private mvar粗体 As Boolean '局部复制
Private mvar表格线加粗 As Boolean '局部复制
Private mvar下线 As Boolean '局部复制
Private mvar斜体 As Boolean '局部复制
Private mvar网格 As Long '局部复制
Private mvar前景 As Long '局部复制
Private mvar背景 As Long '局部复制
Private mvar边框 As Boolean '局部复制
Private mvar系统 As Boolean '局部复制
Private mvar分栏 As Byte '局部复制
Private mvar排序 As String '局部复制
Private mvar格式 As String '局部复制
Private mvar汇总 As String '局部复制
Private mvarCopyIDs As RelatIDs '局部复制
Private mvarSubIDs As RelatIDs '局部复制
Private mvar每页数量 As Integer '局部复制
Private mvar父ID As Integer '局部复制
Private mvar数据源 As String '局部复制
Private mvar上下间距 As Long '局部复制
Private mvar左右间距 As Long '局部复制
Private mvar纵向分栏 As Long '局部复制
Private mvar横向分栏 As Long '局部复制
Private mvar源行号 As Long '局部复制
Private mvarRelations As RPTRelations '局部复制
Private mvarColProtertys As RPTColProtertys '局部复制
Private mstrValue As String
Private mbln自适应行高 As Boolean                   '自适应行高
Private mbln水平反转 As Boolean

Public Property Get 水平反转() As Boolean
    水平反转 = mbln水平反转
End Property
Public Property Let 水平反转(ByVal vData As Boolean)
    mbln水平反转 = vData
End Property

Public Property Get 自适应行高() As Boolean
    自适应行高 = mbln自适应行高
End Property

Public Property Let 自适应行高(ByVal vData As Boolean)
    mbln自适应行高 = vData
End Property

Private Sub Class_Initialize()
    Set CopyIDs = New RelatIDs
    Set SubIDs = New RelatIDs
    Set Relations = New RPTRelations
    Set ColProtertys = New RPTColProtertys
    Set 图片 = Nothing
End Sub

Public Property Let Value(ByVal vData As String)
    mstrValue = vData
End Property

Public Property Get Value() As String
    Value = mstrValue
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

Public Property Let 参照(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.参照 = 5
    mvar参照 = vData
End Property

Public Property Get 参照() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.参照
    参照 = mvar参照
End Property

Public Property Let 性质(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性质 = 5
    mvar性质 = vData
End Property

Public Property Get 性质() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.性质
    性质 = mvar性质
End Property

Public Property Let 自调(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.自调 = 5
    mvar自调 = vData
End Property

Public Property Get 自调() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.自调
    自调 = mvar自调
End Property

Public Property Set 图片(ByVal vData As StdPicture)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.图片 = Form1
    Set mvar图片 = vData
End Property

Public Property Get 图片() As StdPicture
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.图片
    Set 图片 = mvar图片
End Property

Public Property Set SubIDs(ByVal vData As RelatIDs)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.SubIDs = Form1
    Set mvarSubIDs = vData
End Property

Public Property Get SubIDs() As RelatIDs
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SubIDs
    Set SubIDs = mvarSubIDs
End Property

Public Property Set CopyIDs(ByVal vData As RelatIDs)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.CopyIDs = Form1
    Set mvarCopyIDs = vData
End Property

Public Property Get CopyIDs() As RelatIDs
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.CopyIDs
    Set CopyIDs = mvarCopyIDs
End Property

Public Property Let 汇总(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.汇总 = 5
    mvar汇总 = vData
End Property

Public Property Get 汇总() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.汇总
    汇总 = mvar汇总
End Property

Public Property Let 分栏(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.分栏 = 5
    mvar分栏 = vData
End Property

Public Property Get 分栏() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.分栏
    分栏 = mvar分栏
End Property

Public Property Let 背景(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.背景 = 5
    mvar背景 = vData
End Property

Public Property Get 背景() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.背景
    背景 = mvar背景
End Property

Public Property Let 前景(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前景 = 5
    mvar前景 = vData
End Property

Public Property Get 前景() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前景
    前景 = mvar前景
End Property

Public Property Let 网格(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.前景 = 5
    mvar网格 = vData
End Property

Public Property Get 网格() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.前景
    网格 = mvar网格
End Property

Public Property Let 斜体(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.斜体 = 5
    mvar斜体 = vData
End Property

Public Property Get 斜体() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.斜体
    斜体 = mvar斜体
End Property

Public Property Let 下线(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.下线 = 5
    mvar下线 = vData
End Property

Public Property Get 下线() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.下线
    下线 = mvar下线
End Property

Public Property Let 粗体(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.粗体 = 5
    mvar粗体 = vData
End Property

Public Property Get 粗体() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.粗体
    粗体 = mvar粗体
End Property

Public Property Let 表格线加粗(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.粗体 = 5
    mvar表格线加粗 = vData
End Property

Public Property Get 表格线加粗() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.粗体
    表格线加粗 = mvar表格线加粗
End Property

Public Property Let 字号(ByVal vData As Single)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.字号 = 5
    mvar字号 = vData
End Property

Public Property Get 字号() As Single
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.字号
    字号 = mvar字号
End Property

Public Property Let 字体(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.字体 = 5
    mvar字体 = vData
End Property

Public Property Get 字体() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.字体
    字体 = mvar字体
End Property

Public Property Let 对齐(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.对齐 = 5
    mvar对齐 = vData
End Property

Public Property Get 对齐() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.对齐
    对齐 = mvar对齐
End Property

Public Property Let 行高(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.行高 = 5
    mvar行高 = vData
End Property

Public Property Get 行高() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.行高
    行高 = mvar行高
End Property

Public Property Let H(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.H = 5
    mvarH = vData
End Property

Public Property Get H() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.H
    H = mvarH
End Property

Public Property Let W(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.W = 5
    mvarW = vData
End Property

Public Property Get W() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.W
    W = mvarW
End Property

Public Property Let Y(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Y = 5
    mvarY = vData
End Property

Public Property Get Y() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Y
    Y = mvarY
End Property

Public Property Let X(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.X = 5
    mvarX = vData
End Property

Public Property Get X() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.X
    X = mvarX
End Property

Public Property Let 表头(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.表头 = 5
    mvar表头 = vData
End Property

Public Property Get 表头() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.表头
    表头 = mvar表头
End Property

Public Property Let 内容(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.内容 = 5
    mvar内容 = vData
End Property

Public Property Get 内容() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.内容
    内容 = mvar内容
End Property

Public Property Let 类型(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'说明：
'1-线条；2-标签/数据源标签；3-未使用；4-自由表格（表头）；6-自由表格（表体）
'5-汇总表格；7-汇总表格（纵向类）；8-汇总表格（横向类）；9-汇总表格（统计项）；10-框线
'11-图片；12-图表；13-条码；14-卡片
    mvar类型 = vData
End Property

Public Property Get 类型() As Byte
'检索属性值时使用，位于赋值语句的右边。
    类型 = mvar类型
End Property

Public Property Let 序号(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property

Public Property Get 序号() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.序号
    序号 = mvar序号
End Property

Public Property Let 上级ID(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.上级ID = 5
    mvar上级ID = vData
End Property

Public Property Get 上级ID() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上级ID
    上级ID = mvar上级ID
End Property

Public Property Let id(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ID = 5
    mvarID = vData
End Property

Public Property Get id() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ID
    id = mvarID
End Property

Public Property Let 边框(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.粗体 = 5
    mvar边框 = vData
End Property

Public Property Get 边框() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.粗体
    边框 = mvar边框
End Property

Public Property Let 格式(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.汇总 = 5
    mvar格式 = vData
End Property

Public Property Get 格式() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.汇总
    格式 = mvar格式
End Property

Public Property Let 排序(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.排序 = 5
    mvar排序 = vData
End Property

Public Property Get 排序() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.排序
    排序 = mvar排序
End Property

Public Property Let 格式号(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.格式号 = 5
    mvar格式号 = vData
End Property

Public Property Get 格式号() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.格式号
    格式号 = mvar格式号
End Property

Public Property Let 系统(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.系统 = 5
    mvar系统 = vData
End Property

Public Property Get 系统() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.系统
    系统 = mvar系统
End Property

Public Property Let 每页数量(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ID = 5
    mvar每页数量 = vData
End Property

Public Property Get 每页数量() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ID
    每页数量 = mvar每页数量
End Property


Public Property Get 父ID() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上级ID
    父ID = mvar父ID
End Property

Public Property Let 父ID(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.上级ID = 5
    mvar父ID = vData
End Property

Public Property Get 数据源() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上级ID
    数据源 = mvar数据源
End Property

Public Property Let 数据源(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.上级ID = 5
    mvar数据源 = vData
End Property

Public Property Get 上下间距() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上级ID
    上下间距 = mvar上下间距
End Property

Public Property Let 上下间距(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.上级ID = 5
    mvar上下间距 = vData
End Property

Public Property Get 左右间距() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上级ID
    左右间距 = mvar左右间距
End Property

Public Property Let 左右间距(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.上级ID = 5
    mvar左右间距 = vData
End Property

Public Property Get 纵向分栏() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上级ID
    纵向分栏 = mvar纵向分栏
End Property

Public Property Let 纵向分栏(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.上级ID = 5
    mvar纵向分栏 = vData
End Property

Public Property Get 横向分栏() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上级ID
    横向分栏 = mvar横向分栏
End Property

Public Property Let 横向分栏(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.上级ID = 5
    mvar横向分栏 = vData
End Property

Public Property Get 源行号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上级ID
    源行号 = mvar源行号
End Property

Public Property Let 源行号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.上级ID = 5
    mvar源行号 = vData
End Property

Public Property Set Relations(ByVal vData As RPTRelations)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Items = Form1
    Set mvarRelations = vData
End Property

Public Property Get Relations() As RPTRelations
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Items
    Set Relations = mvarRelations
End Property

Public Property Set ColProtertys(ByVal vData As RPTColProtertys)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Items = Form1
    Set mvarColProtertys = vData
End Property

Public Property Get ColProtertys() As RPTColProtertys
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Items
    Set ColProtertys = mvarColProtertys
End Property


