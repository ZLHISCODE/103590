VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPatiInfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'保持属性值的局部变量
Private mvar病人ID As Long '局部复制
Private mvar姓名 As String '局部复制
Private mvar性别 As String '局部复制
Private mvar年龄 As String '局部复制
Private mvar住院号 As String '局部复制
Private mvar床号 As String '局部复制
Private mvarObj费用概况 As clsPatiFeeinfor '局部复制
'保持属性值的局部变量
Private mvar报警方案 As String '局部复制
'保持属性值的局部变量
Private mvar当日记帐总额 As Double '局部复制
Public Property Let 当日记帐总额(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.当日记帐总额 = 5
    mvar当日记帐总额 = vData
End Property


Public Property Get 当日记帐总额() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.当日记帐总额
    当日记帐总额 = mvar当日记帐总额
End Property



Public Property Let 报警方案(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.报警方案 = 5
    mvar报警方案 = vData
End Property


Public Property Get 报警方案() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.报警方案
    报警方案 = mvar报警方案
End Property



Public Property Set Obj费用概况(ByVal vData As clsPatiFeeinfor)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Obj费用概况 = Form1
    Set mvarObj费用概况 = vData
End Property
Public Property Get Obj费用概况() As clsPatiFeeinfor
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Obj费用概况
    Set Obj费用概况 = mvarObj费用概况
End Property



Public Property Let 年龄(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar年龄 = vData
End Property


Public Property Get 年龄() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.年龄
    年龄 = mvar年龄
End Property

Public Property Let 住院号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar住院号 = vData
End Property


Public Property Get 住院号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.年龄
    住院号 = mvar住院号
End Property

Public Property Let 床号(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年龄 = 5
    mvar床号 = vData
End Property


Public Property Get 床号() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.年龄
    床号 = mvar床号
End Property


Public Property Let 性别(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.性别 = 5
    mvar性别 = vData
End Property


Public Property Get 性别() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.性别
    性别 = mvar性别
End Property

Public Property Let 姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.姓名 = 5
    mvar姓名 = vData
End Property


Public Property Get 姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.姓名
    姓名 = mvar姓名
End Property

Public Property Let 病人ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.病人ID = 5
    mvar病人ID = vData
End Property
Public Property Get 病人ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.病人ID
    病人ID = mvar病人ID
End Property
Private Sub Class_Initialize()
    Set mvarObj费用概况 = New clsPatiFeeinfor
End Sub
