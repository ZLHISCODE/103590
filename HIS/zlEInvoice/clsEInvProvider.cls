VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEInvProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'保持属性值的局部变量
Private mvar编码 As String '局部复制
Private mvar提供者 As String '局部复制
Private mvar是否启用 As Boolean '局部复制
Private mvar接口程序名称 As String '局部复制
Private mvar允许设置 As Boolean '局部复制
Public Property Let 允许设置(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.允许设置 = 5
    mvar允许设置 = vData
End Property
Public Property Get 允许设置() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.允许设置
    允许设置 = mvar允许设置
End Property
Public Property Let 接口程序名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.接口程序名称 = 5
    mvar接口程序名称 = vData
End Property


Public Property Get 接口程序名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.接口程序名称
    接口程序名称 = mvar接口程序名称
End Property



Public Property Let 是否启用(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否启用 = 5
    mvar是否启用 = vData
End Property


Public Property Get 是否启用() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否启用
    是否启用 = mvar是否启用
End Property



Public Property Let 提供者(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.提供者 = 5
    mvar提供者 = vData
End Property


Public Property Get 提供者() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.提供者
    提供者 = mvar提供者
End Property



Public Property Let 编码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.编码 = 5
    mvar编码 = vData
End Property


Public Property Get 编码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.编码
    编码 = mvar编码
End Property



