VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "号序信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private mvar序号 As Long '局部复制
Private mvar开始时间 As String '局部复制 格式:yyyy-mm-dd HH:MM:SS
Private mvar终止时间 As String '局部复制 格式:yyyy-mm-dd HH:MM:SS
Private mvar数量 As Double '局部复制
Private mvar是否预约 As Boolean '局部复制
Private mvar是否停诊 As Boolean '局部复制

Public Property Let 是否停诊(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否停诊 = 5
    mvar是否停诊 = vData
End Property


Public Property Get 是否停诊() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否停诊
    是否停诊 = mvar是否停诊
End Property

Public Property Let 是否预约(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否预约 = 5
    mvar是否预约 = vData
End Property


Public Property Get 是否预约() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否预约
    是否预约 = mvar是否预约
End Property



Public Property Let 数量(ByVal vData As Double)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.数量 = 5
    mvar数量 = vData
End Property


Public Property Get 数量() As Double
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.数量
    数量 = mvar数量
End Property



Public Property Let 终止时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.终止时间 = 5
    mvar终止时间 = vData
End Property


Public Property Get 终止时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.终止时间
    终止时间 = mvar终止时间
End Property



Public Property Let 开始时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.开始时间 = 5
    mvar开始时间 = vData
End Property


Public Property Get 开始时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.开始时间
    开始时间 = mvar开始时间
End Property



Public Property Let 序号(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.序号 = 5
    mvar序号 = vData
End Property


Public Property Get 序号() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.序号
    序号 = mvar序号
End Property


Public Function Clone() As 号序信息
    '创建一个副本
    Dim obj号序信息 As New 号序信息
    
    With obj号序信息
        .序号 = mvar序号
        .开始时间 = mvar开始时间
        .终止时间 = mvar终止时间
        .数量 = mvar数量
        .是否预约 = mvar是否预约
        .是否停诊 = mvar是否停诊
    End With
    Set Clone = obj号序信息
End Function
