VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusinessData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'保持属性值的局部变量

Private mvar数据类型 As gDataType '局部复制
Private mvar节点名称 As String '局部复制
Private mvar数据内容 As String '局部复制
Private mvar数据名称 As String '局部复制
Private mvar数据说明 As String '局部复制
'保持属性值的局部变量
Private mvar是否可选 As Boolean '局部复制
Private mvar是否空值 As Boolean '局部复制
Private mvar长度 As Integer     '局部复制
Private mvar小数 As Integer     '局部复制
Private mvarTag As String
Public Property Let 节点名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
    mvar节点名称 = vData
End Property

Public Property Get 节点名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.登记时间
    节点名称 = mvar节点名称
End Property

Public Property Let Tag(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
    mvarTag = vData
End Property


Public Property Get Tag() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.登记时间
    Tag = mvarTag
End Property

Public Property Let 小数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.小数 = 5
    mvar小数 = vData
End Property


Public Property Get 小数() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.小数
    小数 = mvar小数
End Property

Public Property Let 长度(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.长度 = 5
    mvar长度 = vData
End Property


Public Property Get 长度() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.长度
    长度 = mvar长度
End Property



Public Property Let 是否可选(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否可选 = 5
    mvar是否可选 = vData
End Property


Public Property Get 是否可选() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否可选
    是否可选 = mvar是否可选
End Property

Public Property Let 是否空值(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否空值 = 5
    mvar是否空值 = vData
End Property


Public Property Get 是否空值() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否空值
    是否空值 = mvar是否空值
End Property

Public Property Let 数据说明(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数据说明 = 5
    mvar数据说明 = vData
End Property


Public Property Get 数据说明() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.数据说明
    数据说明 = mvar数据说明
End Property



Public Property Let 数据名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数据名称 = 5
    mvar数据名称 = vData
End Property


Public Property Get 数据名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.数据名称
    数据名称 = mvar数据名称
End Property



Public Property Let 数据内容(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数据内容 = 5
    mvar数据内容 = vData
End Property


Public Property Get 数据内容() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.数据内容
    数据内容 = mvar数据内容
End Property



Public Property Let 数据类型(ByVal vData As gDataType)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数据类型 = 5
    mvar数据类型 = vData
End Property


Public Property Get 数据类型() As gDataType
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.数据类型
    数据类型 = mvar数据类型
End Property



