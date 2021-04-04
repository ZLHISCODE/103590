VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "出诊记录"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"上班时段"
Attribute VB_Ext_KEY = "Member1" ,"号序信息集"
Attribute VB_Ext_KEY = "Member2" ,"合作单位控制集"
Option Explicit

'保持属性值的局部变量
Private mvar记录ID As Long '局部复制
Private mvar时间段 As String '局部复制
Private mvar上班时段 As 上班时段 '局部复制
Private mvar是否分时段 As Boolean '局部复制
Private mvar是否序号控制 As Boolean '局部复制
Private mvar限号数 As Long '局部复制
Private mvar限约数 As Long '局部复制
Private mvar预约控制 As Integer '局部复制
Private mvar分诊方式 As Byte '局部复制
Private mvar号序信息集 As 号序信息集 '局部复制
Private mvar合作单位控制集 As 合作单位控制集 '局部复制
Private mvar安排门诊诊室集 As 分诊诊室集 '局部复制
Private mvar出诊日期 As String '局部复制
Private mvar已挂数 As Long '局部复制
Private mvar已约数 As Long '局部复制
Private mvar替诊医生 As String '局部复制
Private mvar开始时间 As String '局部复制
Private mvar终止时间 As String '局部复制
Private mvar是否固定 As Boolean '局部复制
Private mvar科室ID As Long '局部复制
Private mvar是否修改 As Boolean '局部复制
Private mvar是否独占 As Boolean '局部复制
Private mvar医生ID As Long '局部复制
Private mvar医生姓名 As String '局部复制
Private mvar停诊开始时间 As String '局部复制
Private mvar停诊终止时间 As String '局部复制
Private mvar停诊原因 As String '局部复制
'保持属性值的局部变量
Private mvar项目ID As Long '局部复制
Private mvar项目名称 As String '局部复制
'保持属性值的局部变量
Private mvar是否临时出诊 As Boolean '局部复制
Public Property Let 是否临时出诊(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否临时出诊 = 5
    mvar是否临时出诊 = vData
End Property


Public Property Get 是否临时出诊() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否临时出诊
    是否临时出诊 = mvar是否临时出诊
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




Public Property Let 停诊原因(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.停诊原因 = 5
    mvar停诊原因 = vData
End Property


Public Property Get 停诊原因() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.停诊原因
    停诊原因 = mvar停诊原因
End Property



Public Property Let 停诊终止时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.停诊终止时间 = 5
    mvar停诊终止时间 = vData
End Property


Public Property Get 停诊终止时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.停诊终止时间
    停诊终止时间 = mvar停诊终止时间
End Property



Public Property Let 停诊开始时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.停诊开始时间 = 5
    mvar停诊开始时间 = vData
End Property


Public Property Get 停诊开始时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.停诊开始时间
    停诊开始时间 = mvar停诊开始时间
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



Public Property Let 是否独占(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否独占 = 5
    mvar是否独占 = vData
End Property


Public Property Get 是否独占() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否独占
    是否独占 = mvar是否独占
End Property



Public Property Let 是否修改(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否修改 = 5
    mvar是否修改 = vData
End Property


Public Property Get 是否修改() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否修改
    是否修改 = mvar是否修改
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



Public Property Let 是否固定(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否固定 = 5
    mvar是否固定 = vData
End Property


Public Property Get 是否固定() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否固定
    是否固定 = mvar是否固定
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



Public Property Let 替诊医生(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.替诊医生 = 5
    mvar替诊医生 = vData
End Property


Public Property Get 替诊医生() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.替诊医生
    替诊医生 = mvar替诊医生
End Property



Public Property Let 已约数(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.已约数 = 5
    mvar已约数 = vData
End Property


Public Property Get 已约数() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.已约数
    已约数 = mvar已约数
End Property



Public Property Let 已挂数(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.已挂数 = 5
    mvar已挂数 = vData
End Property


Public Property Get 已挂数() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.已挂数
    已挂数 = mvar已挂数
End Property



Public Property Let 出诊日期(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出诊日期 = 5
    mvar出诊日期 = vData
End Property


Public Property Get 出诊日期() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出诊日期
    出诊日期 = mvar出诊日期
End Property





Public Property Let 分诊方式(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.分诊方式 = 5
    mvar分诊方式 = vData
End Property


Public Property Get 分诊方式() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.分诊方式
    分诊方式 = mvar分诊方式
End Property



Public Property Let 记录ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.记录ID = 5
    mvar记录ID = vData
End Property


Public Property Get 记录ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.记录ID
    记录ID = mvar记录ID
End Property



Public Property Get 安排门诊诊室集() As 分诊诊室集
    If mvar安排门诊诊室集 Is Nothing Then
        Set mvar安排门诊诊室集 = New 分诊诊室集
    End If
    Set 安排门诊诊室集 = mvar安排门诊诊室集
End Property


Public Property Set 安排门诊诊室集(vData As 分诊诊室集)
    Set mvar安排门诊诊室集 = vData
End Property



Public Property Let 限约数(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.限约数 = 5
    mvar限约数 = vData
End Property


Public Property Get 限约数() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.限约数
    限约数 = mvar限约数
End Property



Public Property Let 限号数(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.限号数 = 5
    mvar限号数 = vData
End Property


Public Property Get 限号数() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.限号数
    限号数 = mvar限号数
End Property



Public Property Let 是否分时段(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否分时段 = 5
    mvar是否分时段 = vData
End Property


Public Property Get 是否分时段() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否分时段
    是否分时段 = mvar是否分时段
End Property



Public Property Let 是否序号控制(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否序号控制 = 5
    mvar是否序号控制 = vData
End Property


Public Property Get 是否序号控制() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否序号控制
    是否序号控制 = mvar是否序号控制
End Property



Public Property Let 预约控制(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预约控制 = 5
    mvar预约控制 = vData
End Property


Public Property Get 预约控制() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.预约控制
    预约控制 = mvar预约控制
End Property



Public Property Get 合作单位控制集() As 合作单位控制集
    If mvar合作单位控制集 Is Nothing Then
        Set mvar合作单位控制集 = New 合作单位控制集
    End If


    Set 合作单位控制集 = mvar合作单位控制集
End Property


Public Property Set 合作单位控制集(vData As 合作单位控制集)
    Set mvar合作单位控制集 = vData
End Property



Public Property Get 号序信息集() As 号序信息集
    If mvar号序信息集 Is Nothing Then
        Set mvar号序信息集 = New 号序信息集
    End If


    Set 号序信息集 = mvar号序信息集
End Property


Public Property Set 号序信息集(vData As 号序信息集)
    Set mvar号序信息集 = vData
End Property



Private Sub Class_Terminate()
    Set mvar安排门诊诊室集 = Nothing
    Set mvar合作单位控制集 = Nothing
    Set mvar号序信息集 = Nothing
End Sub



Public Property Set 上班时段(ByVal vData As 上班时段)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.上班时段 = Form1
    Set mvar上班时段 = vData
End Property


Public Property Get 上班时段() As 上班时段
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.上班时段
    Set 上班时段 = mvar上班时段
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


Public Function Clone() As 出诊记录
    '创建一个副本
    Dim obj出诊记录  As New 出诊记录
    
    With obj出诊记录
        .记录ID = mvar记录ID
        .时间段 = mvar时间段
        If Not mvar上班时段 Is Nothing Then Set .上班时段 = mvar上班时段.Clone
        .是否分时段 = mvar是否分时段
        .是否序号控制 = mvar是否序号控制
        .限号数 = mvar限号数
        .限约数 = mvar限约数
        .预约控制 = mvar预约控制
        .分诊方式 = mvar分诊方式
        If Not mvar安排门诊诊室集 Is Nothing Then Set .安排门诊诊室集 = mvar安排门诊诊室集.Clone
        If Not mvar号序信息集 Is Nothing Then Set .号序信息集 = mvar号序信息集.Clone
        If Not mvar合作单位控制集 Is Nothing Then Set .合作单位控制集 = mvar合作单位控制集.Clone
        
        .出诊日期 = mvar出诊日期
        .已挂数 = mvar已挂数
        .已约数 = mvar已约数
        .替诊医生 = mvar替诊医生
        .科室ID = mvar科室ID
        .医生ID = mvar医生ID
        .医生姓名 = mvar医生姓名
        If mvar开始时间 = "" And Not mvar上班时段 Is Nothing Then
            .开始时间 = .上班时段.开始时间
        Else
            .开始时间 = mvar开始时间
        End If
        If mvar终止时间 = "" And Not mvar上班时段 Is Nothing Then
            .终止时间 = .上班时段.结束时间
        Else
            .终止时间 = mvar终止时间
        End If
        .是否固定 = mvar是否固定
        .是否修改 = mvar是否修改
        .是否独占 = mvar是否独占
        .停诊开始时间 = mvar停诊开始时间
        .停诊终止时间 = mvar停诊终止时间
        .停诊原因 = mvar停诊原因
        .是否临时出诊 = mvar是否临时出诊
    End With
    Set Clone = obj出诊记录
End Function

