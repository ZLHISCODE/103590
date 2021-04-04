VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "出诊记录集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"出诊记录"
Attribute VB_Ext_KEY = "Member0" ,"出诊记录"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'局部变量，保存集合
Private mCol As Collection
'保持属性值的局部变量
Private mvar出诊日期 As String '局部复制
Private mvar是否删除 As Boolean '局部复制
Private mvar是否修改 As Boolean '局部复制

Public Function Add(时间段 As String, 上班时段 As 上班时段, 号序信息集 As 号序信息集, _
    合作单位控制集 As 合作单位控制集, 预约控制 As Integer, 是否序号控制 As Boolean, _
    是否分时段 As Boolean, 限号数 As Long, 限约数 As Long, 安排门诊诊室集 As 分诊诊室集, _
    记录ID As Long, 分诊方式 As Byte, Optional sKey As String) As 出诊记录
    '创建新对象
    Dim objNewMember As 出诊记录
    Set objNewMember = New 出诊记录


    '设置传入方法的属性
    objNewMember.时间段 = 时间段
    Set objNewMember.上班时段 = 上班时段
    Set objNewMember.号序信息集 = 号序信息集
    Set objNewMember.合作单位控制集 = 合作单位控制集
    objNewMember.预约控制 = 预约控制
    objNewMember.是否序号控制 = 是否序号控制
    objNewMember.是否分时段 = 是否分时段
    objNewMember.限号数 = 限号数
    objNewMember.限约数 = 限约数
    Set objNewMember.安排门诊诊室集 = 安排门诊诊室集
    objNewMember.记录ID = 记录ID
    objNewMember.分诊方式 = 分诊方式
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As 出诊记录
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

Public Sub AddItem(obj出诊记录 As 出诊记录, Optional sKey As String)
    If obj出诊记录 Is Nothing Then Set obj出诊记录 = New 出诊记录
    If Len(sKey) = 0 Then
        mCol.Add obj出诊记录
    Else
        mCol.Add obj出诊记录, sKey
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

Public Function Clone() As 出诊记录集
    '创建一个副本
    Dim obj出诊记录集 As New 出诊记录集
    Dim objItem As 出诊记录, strKey As String
    
    With obj出诊记录集
        For Each objItem In mCol
            strKey = "K" & objItem.时间段
            .AddItem objItem.Clone, strKey
        Next
        .出诊日期 = mvar出诊日期
        .是否删除 = mvar是否删除
        .是否修改 = mvar是否修改
    End With
    Set Clone = obj出诊记录集
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
