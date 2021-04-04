VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "号序信息集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"号序信息"
Attribute VB_Ext_KEY = "Member0" ,"上班时段"
Option Explicit
'局部变量，保存集合
Private mCol As Collection
'保持属性值的局部变量
Private mvar限号数 As Long '局部复制
Private mvar限约数 As Long '局部复制
Private mvar预约控制 As Integer '局部复制
Private mvar是否序号控制 As Boolean '局部复制
Private mvar是否分时段 As Boolean '局部复制
Private mvar出诊频次 As Integer '局部复制
Private mvar时间段 As String '局部复制
'保持属性值的局部变量
Private mvar是否修改 As Boolean '局部复制
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

Public Function Clone() As 号序信息集
    '创建一个副本
    Dim obj号序信息集 As New 号序信息集
    Dim objItem As 号序信息, strKey As String
    
    With obj号序信息集
        For Each objItem In mCol
            .AddItem objItem.Clone
        Next
        .是否分时段 = mvar是否分时段
        .是否序号控制 = mvar是否序号控制
        .出诊频次 = mvar出诊频次
        .限号数 = mvar限号数
        .限约数 = mvar限约数
        .预约控制 = mvar预约控制
        .时间段 = mvar时间段
        .是否修改 = mvar是否修改
    End With
    Set Clone = obj号序信息集
End Function

Public Function Add(序号 As Long, 开始时间 As String, 终止时间 As String, _
    数量 As Double, 是否预约 As Boolean, 是否停诊 As Boolean, Optional sKey As String) As 号序信息
    '创建新对象
    Dim objNewMember As 号序信息
    Set objNewMember = New 号序信息


    '设置传入方法的属性
    objNewMember.序号 = 序号
    objNewMember.开始时间 = 开始时间
    objNewMember.终止时间 = 终止时间
    objNewMember.数量 = 数量
    objNewMember.是否预约 = 是否预约
    objNewMember.是否停诊 = 是否停诊
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function
Public Sub AddItem(obj号序信息 As 号序信息, Optional sKey As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按号序对象接加
    '编制:刘兴洪
    '日期:2016-01-13 14:00:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If obj号序信息 Is Nothing Then Set obj号序信息 = New 号序信息
    If Len(sKey) = 0 Then
        mCol.Add obj号序信息
    Else
        mCol.Add obj号序信息, sKey
    End If
End Sub
Public Property Get Item(vntIndexKey As Variant) As 号序信息
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
End Sub

Public Sub RemoveAll()
    '删除集合中的所有元素时使用。
    Dim i As Integer
    
    For i = 1 To mCol.Count
        If i > mCol.Count Then Exit For
        mCol.Remove i
        i = i - 1
    Next
End Sub

