VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "出诊安排"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"出诊记录"
Attribute VB_Ext_KEY = "Member0" ,"出诊号源"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member1" ,"出诊记录"
Attribute VB_Ext_KEY = "Member2" ,"停诊记录"
Option Explicit

'局部变量，保存集合
Private mCol As Collection
'保持属性值的局部变量
Private mvar出诊ID As Long '局部复制
Private mvar出诊表名 As String '局部复制
Private mvar排班方式 As Byte '局部复制
Private mvar年份 As Integer '局部复制
Private mvar月份 As Integer '局部复制
Private mvar周数 As Integer '局部复制
Private mvar应用范围 As Byte '局部复制
Private mvar科室ID As Long '局部复制
Private mvar科室名称 As String '局部复制
Private mvar备注 As String '局部复制
Private mvar发布人 As String '局部复制
Private mvar发布时间 As String '局部复制

Private mvar安排ID As Long '局部复制
Private mvar出诊号源 As 出诊号源 '局部复制
Private mvar排班规则 As Byte '局部复制
Private mvar周六不出诊 As Boolean '局部复制
Private mvar周日不出诊 As Boolean '局部复制
Private mvar开始时间 As String '局部复制
Private mvar终止时间 As String '局部复制
Private mvar医生ID As Long '局部复制
Private mvar医生姓名 As String '局部复制
Private mvar项目ID As Long '局部复制
Private mvar操作员姓名 As String '局部复制
Private mvar登记时间 As String '局部复制

Private mvar所有上班时段 As 上班时段集 '局部复制
Private mvar已保存出诊安排 As 出诊安排 '局部复制
Private mvar所有合作单位 As 合作单位控制集 '局部复制
Private mvar临时出诊 As Boolean '局部复制
Private mvar缺省出诊日期 As String '局部复制
Private mvar更新合作单位 As Boolean '局部复制
Private mvar未保存出诊安排 As 出诊安排 '局部复制
Private mvar号源安排 As 出诊记录集 '局部复制
'保持属性值的局部变量
Private mvar所有分诊诊室 As 分诊诊室集 '局部复制
'保持属性值的局部变量
Private mvar停诊记录 As 停诊记录集 '局部复制
'保持属性值的局部变量
Private mvar项目名称 As String '局部复制
'保持属性值的局部变量
Private mvar医生职称 As String '局部复制
'保持属性值的局部变量
Private mvar是否临时安排 As Boolean '局部复制
'保持属性值的局部变量
Private mvar模板类型 As Byte '局部复制
Public Property Let 模板类型(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.模板类型 = 5
    mvar模板类型 = vData
End Property


Public Property Get 模板类型() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.模板类型
    模板类型 = mvar模板类型
End Property



Public Property Let 是否临时安排(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.是否临时安排 = 5
    mvar是否临时安排 = vData
End Property


Public Property Get 是否临时安排() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.是否临时安排
    是否临时安排 = mvar是否临时安排
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



Public Property Set 停诊记录(ByVal vData As 停诊记录集)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.停诊记录 = Form1
    Set mvar停诊记录 = vData
End Property


Public Property Get 停诊记录() As 停诊记录集
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.停诊记录
    Set 停诊记录 = mvar停诊记录
End Property



Public Property Set 所有分诊诊室(ByVal vData As 分诊诊室集)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.所有分诊诊室 = Form1
    Set mvar所有分诊诊室 = vData
End Property


Public Property Get 所有分诊诊室() As 分诊诊室集
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.所有分诊诊室
    Set 所有分诊诊室 = mvar所有分诊诊室
End Property



Public Property Set 号源安排(ByVal vData As 出诊记录集)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.号源安排 = Form1
    Set mvar号源安排 = vData
End Property


Public Property Get 号源安排() As 出诊记录集
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.号源安排
    Set 号源安排 = mvar号源安排
End Property



Public Property Set 未保存出诊安排(ByVal vData As 出诊安排)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.未保存出诊安排 = Form1
    Set mvar未保存出诊安排 = vData
End Property


Public Property Get 未保存出诊安排() As 出诊安排
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.未保存出诊安排
    Set 未保存出诊安排 = mvar未保存出诊安排
End Property



Public Property Let 更新合作单位(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.更新合作单位 = 5
    mvar更新合作单位 = vData
End Property


Public Property Get 更新合作单位() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.更新合作单位
    更新合作单位 = mvar更新合作单位
End Property



Public Property Let 缺省出诊日期(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.缺省出诊日期 = 5
    mvar缺省出诊日期 = vData
End Property


Public Property Get 缺省出诊日期() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.缺省出诊日期
    缺省出诊日期 = mvar缺省出诊日期
End Property



Public Property Let 临时出诊(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.临时出诊 = 5
    mvar临时出诊 = vData
End Property


Public Property Get 临时出诊() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.临时出诊
    临时出诊 = mvar临时出诊
End Property

Public Property Let 登记时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.登记时间 = 5
    mvar登记时间 = vData
End Property


Public Property Get 登记时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.登记时间
    登记时间 = mvar登记时间
End Property



Public Property Let 操作员姓名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.操作员姓名 = 5
    mvar操作员姓名 = vData
End Property


Public Property Get 操作员姓名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.操作员姓名
    操作员姓名 = mvar操作员姓名
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




Public Property Get Item(vntIndexKey As Variant) As 出诊记录集
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
    Set mvar出诊号源 = Nothing
    Set mvar所有上班时段 = Nothing
    Set mvar已保存出诊安排 = Nothing
    Set mvar所有合作单位 = Nothing
End Sub

Public Property Set 所有合作单位(ByVal vData As 合作单位控制集)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.合作单位控制集 = Form1
    Set mvar所有合作单位 = vData
End Property


Public Property Get 所有合作单位() As 合作单位控制集
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.合作单位控制集
    Set 所有合作单位 = mvar所有合作单位
End Property



Public Property Set 已保存出诊安排(ByVal vData As 出诊安排)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.已保存出诊安排 = Form1
    Set mvar已保存出诊安排 = vData
End Property


Public Property Get 已保存出诊安排() As 出诊安排
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.已保存出诊安排
    Set 已保存出诊安排 = mvar已保存出诊安排
End Property



Public Property Set 所有上班时段(ByVal vData As 上班时段集)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.所有上班时段 = Form1
    Set mvar所有上班时段 = vData
End Property


Public Property Get 所有上班时段() As 上班时段集
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.所有上班时段
    Set 所有上班时段 = mvar所有上班时段
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



Public Property Let 安排ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.安排ID = 5
    mvar安排ID = vData
End Property


Public Property Get 安排ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.安排ID
    安排ID = mvar安排ID
End Property



Public Property Get 出诊号源() As 出诊号源
    If mvar出诊号源 Is Nothing Then
        Set mvar出诊号源 = New 出诊号源
    End If

    Set 出诊号源 = mvar出诊号源
End Property


Public Property Set 出诊号源(vData As 出诊号源)
    Set mvar出诊号源 = vData
End Property


Public Property Let 出诊ID(ByVal vData As Long)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出诊ID = 5
    mvar出诊ID = vData
End Property


Public Property Get 出诊ID() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出诊ID
    出诊ID = mvar出诊ID
End Property


Public Property Let 发布时间(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发布时间 = 5
    mvar发布时间 = vData
End Property


Public Property Get 发布时间() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.发布时间
    发布时间 = mvar发布时间
End Property



Public Property Let 发布人(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.发布人 = 5
    mvar发布人 = vData
End Property


Public Property Get 发布人() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.发布人
    发布人 = mvar发布人
End Property



Public Property Let 备注(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.备注 = 5
    mvar备注 = vData
End Property


Public Property Get 备注() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.备注
    备注 = mvar备注
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



Public Property Let 应用范围(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.应用范围 = 5
    mvar应用范围 = vData
End Property


Public Property Get 应用范围() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.应用范围
    应用范围 = mvar应用范围
End Property



Public Property Let 周数(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.周数 = 5
    mvar周数 = vData
End Property


Public Property Get 周数() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.周数
    周数 = mvar周数
End Property



Public Property Let 月份(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.月份 = 5
    mvar月份 = vData
End Property


Public Property Get 月份() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.月份
    月份 = mvar月份
End Property



Public Property Let 年份(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.年份 = 5
    mvar年份 = vData
End Property


Public Property Get 年份() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.年份
    年份 = mvar年份
End Property



Public Property Let 排班方式(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.排班方式 = 5
    mvar排班方式 = vData
End Property


Public Property Get 排班方式() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.排班方式
    排班方式 = mvar排班方式
End Property



Public Property Let 出诊表名(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.出诊表名 = 5
    mvar出诊表名 = vData
End Property


Public Property Get 出诊表名() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.出诊表名
    出诊表名 = mvar出诊表名
End Property



Public Property Let 周日不出诊(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.周日不出诊 = 5
    mvar周日不出诊 = vData
End Property


Public Property Get 周日不出诊() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.周日不出诊
    周日不出诊 = mvar周日不出诊
End Property



Public Property Let 周六不出诊(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.周六不出诊 = 5
    mvar周六不出诊 = vData
End Property


Public Property Get 周六不出诊() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.周六不出诊
    周六不出诊 = mvar周六不出诊
End Property



Public Property Let 排班规则(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.排班规则 = 5
    mvar排班规则 = vData
End Property


Public Property Get 排班规则() As Byte
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.排班规则
    排班规则 = mvar排班规则
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


Public Sub AddItem(obj出诊记录集 As 出诊记录集, Optional sKey As String)
    If obj出诊记录集 Is Nothing Then Set obj出诊记录集 = New 出诊记录集
    If Len(sKey) = 0 Then
        mCol.Add obj出诊记录集
    Else
        mCol.Add obj出诊记录集, sKey
    End If
End Sub

Public Function Exits(ByVal strKey As String) As Boolean
    '判断元素是否存在
    If Len(strKey) = 0 Then Exit Function
    If mCol Is Nothing Then Exit Function
    
    Exits = True
    Err = 0: On Error Resume Next
    Exits = IsObject(mCol(strKey))
    If Err <> 0 Then Err = 0: Exits = False
End Function

Public Function Clone(Optional blnExceptItem As Boolean) As 出诊安排
    '创建一个副本
    'blnExceptItem 不复制元素
    Dim obj出诊安排 As New 出诊安排
    Dim ObjItem As 出诊记录集, strKey As String
    
    If blnExceptItem = False Then
        For Each ObjItem In mCol
            obj出诊安排.AddItem ObjItem.Clone, GetPlanKey(ObjItem.出诊日期)
        Next
    End If
    With obj出诊安排
        .出诊ID = mvar出诊ID
        .出诊表名 = mvar出诊表名
        .排班方式 = mvar排班方式
        .年份 = mvar年份
        .月份 = mvar月份
        .周数 = mvar周数
        .应用范围 = mvar应用范围
        .科室ID = mvar科室ID
        .科室名称 = mvar科室名称
        .备注 = mvar备注
        .发布人 = mvar发布人
        .发布时间 = mvar发布时间
        .模板类型 = mvar模板类型
        
        .安排ID = mvar安排ID
        If Not mvar出诊号源 Is Nothing Then Set .出诊号源 = mvar出诊号源.Clone
        .排班规则 = mvar排班规则
        .周六不出诊 = mvar周六不出诊
        .周日不出诊 = mvar周日不出诊
        .开始时间 = mvar开始时间
        .终止时间 = mvar终止时间
        .项目ID = mvar项目ID
        .项目名称 = mvar项目名称
        .医生ID = mvar医生ID
        .医生姓名 = mvar医生姓名
        .医生职称 = mvar医生职称
        .操作员姓名 = mvar操作员姓名
        .登记时间 = mvar登记时间
        .是否临时安排 = mvar是否临时安排
        
        If Not mvar所有上班时段 Is Nothing Then Set .所有上班时段 = mvar所有上班时段.Clone
        If Not mvar已保存出诊安排 Is Nothing Then Set .已保存出诊安排 = mvar已保存出诊安排.Clone
        If Not mvar所有合作单位 Is Nothing Then Set .所有合作单位 = mvar所有合作单位.Clone
        
        .临时出诊 = mvar临时出诊
        .缺省出诊日期 = mvar缺省出诊日期
        .更新合作单位 = mvar更新合作单位
        If Not mvar未保存出诊安排 Is Nothing Then Set .未保存出诊安排 = mvar未保存出诊安排.Clone
        If Not mvar号源安排 Is Nothing Then Set .号源安排 = mvar号源安排.Clone
        If Not mvar所有分诊诊室 Is Nothing Then Set .所有分诊诊室 = mvar所有分诊诊室.Clone
        If Not mvar停诊记录 Is Nothing Then Set .停诊记录 = mvar停诊记录.Clone
    End With
    Set Clone = obj出诊安排
End Function

Public Sub RemoveAll()
    '删除集合中的所有元素时使用。
    Dim i As Integer
    
    For i = 1 To mCol.Count
        If i > mCol.Count Then Exit For
        mCol.Remove i
        i = i - 1
    Next
End Sub
