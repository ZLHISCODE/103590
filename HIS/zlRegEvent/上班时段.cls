VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "上班时段"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'保持属性值的局部变量
Private mvar时间段 As String '局部复制
Private mvar开始时间 As String '局部复制
Private mvar结束时间 As String '局部复制
Private mvar缺省预约时间 As String '局部复制
Private mvar提前挂号时间 As String '局部复制
Private mvar出诊预留时间 As Integer '局部复制
Private mvar休息时段 As String '局部复制
Public Property Let 休息时段(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.休息时段 = 5
    mvar休息时段 = vData
End Property


Public Property Get 休息时段() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.休息时段
    休息时段 = mvar休息时段
End Property



Public Property Let 出诊预留时间(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出诊预留时间 = 5
    mvar出诊预留时间 = vData
End Property


Public Property Get 出诊预留时间() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出诊预留时间
    出诊预留时间 = mvar出诊预留时间
End Property


Public Property Let 提前挂号时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.提前挂号时间 = 5
    mvar提前挂号时间 = vData
End Property


Public Property Get 提前挂号时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.提前挂号时间
    提前挂号时间 = mvar提前挂号时间
End Property



Public Property Let 缺省预约时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缺省预约时间 = 5
    mvar缺省预约时间 = vData
End Property


Public Property Get 缺省预约时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缺省预约时间
    缺省预约时间 = mvar缺省预约时间
End Property



Public Property Let 结束时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.结束时间 = 5
    mvar结束时间 = vData
End Property


Public Property Get 结束时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.结束时间
    结束时间 = mvar结束时间
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



Public Property Let 时间段(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.时间段 = 5
    mvar时间段 = vData
End Property


Public Property Get 时间段() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.时间段
    时间段 = mvar时间段
End Property

Public Function Clone() As 上班时段
    '创建一个副本
    Dim obj上班时段  As New 上班时段
    
    With obj上班时段
        .时间段 = mvar时间段
        .结束时间 = mvar结束时间
        .开始时间 = mvar开始时间
        .缺省预约时间 = mvar缺省预约时间
        .提前挂号时间 = mvar提前挂号时间
        .出诊预留时间 = mvar出诊预留时间
        .休息时段 = mvar休息时段
    End With
    Set Clone = obj上班时段
End Function

