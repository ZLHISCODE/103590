VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "出诊号源"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"分诊诊室集"
Option Explicit
'保持属性值的局部变量
Private mvarID As Long '局部复制
Private mvar号类 As String '局部复制
Private mvar号码 As String '局部复制
Private mvar科室ID As Long '局部复制
Private mvar科室名称 As String '局部复制
Private mvar项目ID As Long '局部复制
Private mvar项目名称 As String '局部复制
Private mvar医生ID As Long '局部复制
Private mvar医生姓名 As String '局部复制
Private mvar是否建病案 As Boolean '局部复制
Private mvar预约天数 As Integer '局部复制
Private mvar出诊频次 As Integer '局部复制
Private mvar假日控制状态 As Integer '局部复制
Private mvar是否临床排班 As Boolean '局部复制
Private mvar排班方式 As Integer '局部复制
Private mvar是否删除 As Boolean '局部复制
Private mvar建档时间 As String '局部复制
Private mvar撤档时间 As String '局部复制
Private mvar是否假日换休 As Boolean '局部复制
'保持属性值的局部变量
Private mvar医生职称 As String '局部复制
'保持属性值的局部变量
Private mvar站点 As String '局部复制
Public Property Let 站点(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.站点 = 5
    mvar站点 = vData
End Property


Public Property Get 站点() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.站点
    站点 = mvar站点
End Property



Public Property Let 医生职称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医生职称 = 5
    mvar医生职称 = vData
End Property


Public Property Get 医生职称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医生职称
    医生职称 = mvar医生职称
End Property




Public Property Let 是否假日换休(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否假日换休 = 5
    mvar是否假日换休 = vData
End Property


Public Property Get 是否假日换休() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否假日换休
    是否假日换休 = mvar是否假日换休
End Property



Public Property Let 撤档时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.撤档时间 = 5
    mvar撤档时间 = vData
End Property


Public Property Get 撤档时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.撤档时间
    撤档时间 = mvar撤档时间
End Property



Public Property Let 建档时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.建档时间 = 5
    mvar建档时间 = vData
End Property


Public Property Get 建档时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.建档时间
    建档时间 = mvar建档时间
End Property



Public Property Let 是否删除(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否删除 = 5
    mvar是否删除 = vData
End Property


Public Property Get 是否删除() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否删除
    是否删除 = mvar是否删除
End Property


Public Property Let 排班方式(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.排班方式 = 5
    mvar排班方式 = vData
End Property


Public Property Get 排班方式() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.排班方式
    排班方式 = mvar排班方式
End Property



Public Property Let 是否临床排班(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否临床排班 = 5
    mvar是否临床排班 = vData
End Property


Public Property Get 是否临床排班() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否临床排班
    是否临床排班 = mvar是否临床排班
End Property



Public Property Let 假日控制状态(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.假日控制状态 = 5
    mvar假日控制状态 = vData
End Property


Public Property Get 假日控制状态() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.假日控制状态
    假日控制状态 = mvar假日控制状态
End Property



Public Property Let 出诊频次(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出诊频次 = 5
    mvar出诊频次 = vData
End Property


Public Property Get 出诊频次() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出诊频次
    出诊频次 = mvar出诊频次
End Property



Public Property Let 预约天数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预约天数 = 5
    mvar预约天数 = vData
End Property


Public Property Get 预约天数() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预约天数
    预约天数 = mvar预约天数
End Property

Public Property Let 是否建病案(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否建病案 = 5
    mvar是否建病案 = vData
End Property


Public Property Get 是否建病案() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否建病案
    是否建病案 = mvar是否建病案
End Property



Public Property Let 医生姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医生姓名 = 5
    mvar医生姓名 = vData
End Property


Public Property Get 医生姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医生姓名
    医生姓名 = mvar医生姓名
End Property



Public Property Let 医生ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.医生ID = 5
    mvar医生ID = vData
End Property


Public Property Get 医生ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.医生ID
    医生ID = mvar医生ID
End Property



Public Property Let 项目名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.项目名称 = 5
    mvar项目名称 = vData
End Property


Public Property Get 项目名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.项目名称
    项目名称 = mvar项目名称
End Property



Public Property Let 项目ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.项目ID = 5
    mvar项目ID = vData
End Property


Public Property Get 项目ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.项目ID
    项目ID = mvar项目ID
End Property



Public Property Let 科室名称(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.科室名称 = 5
    mvar科室名称 = vData
End Property


Public Property Get 科室名称() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.科室名称
    科室名称 = mvar科室名称
End Property



Public Property Let 科室ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.科室ID = 5
    mvar科室ID = vData
End Property


Public Property Get 科室ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.科室ID
    科室ID = mvar科室ID
End Property



Public Property Let 号码(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.号码 = 5
    mvar号码 = vData
End Property


Public Property Get 号码() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.号码
    号码 = mvar号码
End Property



Public Property Let 号类(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.号类 = 5
    mvar号类 = vData
End Property


Public Property Get 号类() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.号类
    号类 = mvar号类
End Property



Public Property Let ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ID = 5
    mvarID = vData
End Property


Public Property Get ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.ID
    ID = mvarID
End Property


Public Function Clone()
    '创建一个副本
    Dim obj出诊号源 As New 出诊号源
    
    With obj出诊号源
        .ID = mvarID
        .号类 = mvar号类
        .号码 = mvar号码
        .科室ID = mvar科室ID
        .科室名称 = mvar科室名称
        .项目ID = mvar项目ID
        .项目名称 = mvar项目名称
        .医生ID = mvar医生ID
        .医生姓名 = mvar医生姓名
        .医生职称 = mvar医生职称
        .是否建病案 = mvar是否建病案
        .预约天数 = mvar预约天数
        .出诊频次 = mvar出诊频次
        .假日控制状态 = mvar假日控制状态
        .是否临床排班 = mvar是否临床排班
        .排班方式 = mvar排班方式
        .是否删除 = mvar是否删除
        .建档时间 = mvar建档时间
        .撤档时间 = mvar撤档时间
        .是否假日换休 = mvar是否假日换休
        .站点 = mvar站点
    End With
    Set Clone = obj出诊号源
End Function
