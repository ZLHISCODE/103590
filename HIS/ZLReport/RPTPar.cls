VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'保持属性值的局部变量
Private mvar组名 As String '局部复制
Private mvar序号 As Byte '局部复制
Private mvar名称 As String '局部复制
Private mvar类型 As Byte '局部复制
Private mvar缺省值 As String '局部复制
Private mvar格式 As Byte '局部复制
Private mvar值列表 As String '局部复制
Private mvar分类SQL As String '局部复制
Private mvar明细SQL As String '局部复制
Private mvar分类字段 As String '局部复制
Private mvar明细字段 As String '局部复制
Private mvar对象 As String '局部复制
Private mvarReserve As String '局部复制
Private mvar是否锁定 As Boolean '局部复制

Public Property Let 组名(ByVal vData As String)
    mvar组名 = vData
End Property

Public Property Get 组名() As String
    组名 = mvar组名
End Property

Public Property Let 对象(ByVal vData As String)
    mvar对象 = vData
End Property

Public Property Get 对象() As String
    对象 = mvar对象
End Property

Public Property Let 明细字段(ByVal vData As String)
    mvar明细字段 = vData
End Property

Public Property Get 明细字段() As String
    明细字段 = mvar明细字段
End Property

Public Property Let 分类字段(ByVal vData As String)
    mvar分类字段 = vData
End Property

Public Property Get 分类字段() As String
    分类字段 = mvar分类字段
End Property

Public Property Let 明细SQL(ByVal vData As String)
    mvar明细SQL = vData
End Property

Public Property Get 明细SQL() As String
    明细SQL = mvar明细SQL
End Property

Public Property Let 分类SQL(ByVal vData As String)
    mvar分类SQL = vData
End Property

Public Property Get 分类SQL() As String
    分类SQL = mvar分类SQL
End Property

Public Property Let 值列表(ByVal vData As String)
    mvar值列表 = vData
End Property

Public Property Get 值列表() As String
    值列表 = mvar值列表
End Property

Public Property Let 缺省值(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缺省值 = 5
    mvar缺省值 = vData
End Property

Public Property Get 缺省值() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缺省值
    缺省值 = mvar缺省值
End Property

Public Property Let 类型(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.类型 = 5
    mvar类型 = vData
End Property

Public Property Get 类型() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.类型
    类型 = mvar类型
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

Public Property Let 序号(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property

Public Property Get 序号() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.序号
    序号 = mvar序号
End Property

Public Property Let Reserve(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缺省值 = 5
    mvarReserve = vData
End Property

Public Property Get Reserve() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缺省值
    Reserve = mvarReserve
End Property

Public Property Let 格式(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.格式 = 5
    mvar格式 = vData
End Property

Public Property Get 格式() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.格式
    格式 = mvar格式
End Property

Public Property Let 是否锁定(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.格式 = 5
    mvar是否锁定 = vData
End Property

Public Property Get 是否锁定() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.格式
    是否锁定 = mvar是否锁定
End Property


