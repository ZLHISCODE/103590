VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

Private mvar名称 As String '局部复制
Private mvarSQL As String '局部复制
Private mvar字段 As String '局部复制
Private mvar对象 As String '局部复制
Private mvar类型 As Byte '局部复制
Private mvar说明 As String '局部复制
Private mvarPars As RPTPars '局部复制
Private mvar原名称 As String '局部复制
Private mintDataConnNo As Integer         '数据连接编号

Public Property Set Pars(ByVal vData As RPTPars)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Pars = Form1
    Set mvarPars = vData
End Property

Public Property Get Pars() As RPTPars
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Pars
    Set Pars = mvarPars
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

Public Property Let 说明(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.说明 = 5
    mvar说明 = vData
End Property

Public Property Get 说明() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.说明
    说明 = mvar说明
End Property

Public Property Let SQL(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.SQL = 5
    mvarSQL = vData
End Property

Public Property Get SQL() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.SQL
    SQL = mvarSQL
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

Private Sub Class_Initialize()
    Set Pars = New RPTPars
End Sub

Public Property Let 字段(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.字段 = 5
    mvar字段 = vData
End Property

Public Property Get 字段() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.字段
    字段 = mvar字段
End Property

Public Property Let 对象(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.字段 = 5
    mvar对象 = vData
End Property

Public Property Get 对象() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.字段
    对象 = mvar对象
End Property

Public Property Let 原名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.名称 = 5
    mvar原名称 = vData
End Property

Public Property Get 原名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.名称
    原名称 = mvar原名称
End Property

Public Property Let 数据连接编号(ByVal intDataConnNo As Integer)
    mintDataConnNo = intDataConnNo
End Property

Public Property Get 数据连接编号() As Integer
    数据连接编号 = mintDataConnNo
End Property
