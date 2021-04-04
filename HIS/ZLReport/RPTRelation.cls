VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Key As String

Private mvar参数名 As String  '局部复制
Private mvar元素ID As Long '局部复制
Private mvar关联报表ID As Long '局部复制
Private mvar参数值来源 As String  '局部复制
Private mvar关联报表名称 As String '局部复制
Private mvar默认 As Integer '局部复制


Public Property Get 关联报表ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表ID
    关联报表ID = mvar关联报表ID
End Property

Public Property Let 关联报表ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表ID = 5
    mvar关联报表ID = vData
End Property

Public Property Let 参数名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.参数名 = 5
    mvar参数名 = vData
End Property

Public Property Get 参数名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.名称
    参数名 = mvar参数名
End Property

Public Property Let 参数值来源(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.参数值来源 = 5
    mvar参数值来源 = vData
End Property

Public Property Get 参数值来源() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.参数值来源
    参数值来源 = mvar参数值来源
End Property

Public Property Let 关联报表名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表 = 5
    mvar关联报表名称 = vData
End Property

Public Property Get 关联报表名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表
    关联报表名称 = mvar关联报表名称
End Property

Public Property Let 默认(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.关联报表 = 5
    mvar默认 = vData
End Property

Public Property Get 默认() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.关联报表
    默认 = mvar默认
End Property



