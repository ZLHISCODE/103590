VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'保持属性值的局部变量
Private mvar编码 As Integer '局部复制
Private mvar名称 As String '局部复制
Private mvar启用 As Boolean '局部复制
Private mvar接口程序名 As String '局部复制
'保持属性值的局部变量
Private mvar险类 As Integer '局部复制
Private mvar设置 As Integer
Private mvar自动读取 As Integer

Public Property Let 是否自动读取(ByVal vdata As Integer)
    mvar自动读取 = vdata
End Property

Public Property Get 是否自动读取() As Integer
    是否自动读取 = mvar自动读取
End Property

Public Property Let 险类(ByVal vdata As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类 = 5
    mvar险类 = vdata
End Property

Public Property Get 险类() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类
    险类 = mvar险类
End Property

Public Property Let 可否设置(ByVal vdata As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.险类 = 5
    mvar设置 = vdata
End Property


Public Property Get 可否设置() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.险类
    可否设置 = mvar设置
End Property



Public Property Let 接口程序名(ByVal vdata As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.接口程序名 = 5
    mvar接口程序名 = vdata
End Property


Public Property Get 接口程序名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.接口程序名
    接口程序名 = mvar接口程序名
End Property



Public Property Let 启用(ByVal vdata As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.启用 = 5
    mvar启用 = vdata
End Property


Public Property Get 启用() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.启用
    启用 = mvar启用
End Property

Public Property Let 名称(ByVal vdata As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.名称 = 5
    mvar名称 = vdata
End Property


Public Property Get 名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.名称
    名称 = mvar名称
End Property

Public Property Let 编码(ByVal vdata As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.编码 = 5
    mvar编码 = vdata
End Property

Public Property Get 编码() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.编码
    编码 = mvar编码
End Property



