VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "合作单位控制集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"合作单位控制"
Attribute VB_Ext_KEY = "Member0" ,"号序信息集"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member1" ,"合作单位控制"
Option Explicit
'局部变量，保存集合
Private mCol As Collection
'保持属性值的局部变量
Private mvar是否独占 As Boolean '局部复制
Private mvar预约控制方式 As Byte '局部复制
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

Public Property Let 预约控制方式(ByVal vData As Byte)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.预约控制方式 = 5
    mvar预约控制方式 = vData
End Property

Public Property Get 预约控制方式() As Byte
    '获取合作单位的预约控制方式
    Dim objItem As 合作单位控制
    
    '0-禁止预约(或挂号);1-按比例控制预约(或挂号);2-按总量控制预约(或挂号);3-按序号控制预约(或挂号);4-不作限制
    For Each objItem In mCol
        If objItem.预约控制方式 <> 0 And objItem.预约控制方式 <> 4 Then
            预约控制方式 = objItem.预约控制方式
            Exit Property
        End If
    Next
    If 预约控制方式 = 0 Then
        预约控制方式 = mvar预约控制方式
    End If
End Property

Public Function Clone() As 合作单位控制集
    '创建一个副本
    Dim obj合作单位控制集 As New 合作单位控制集
    Dim objItem As 合作单位控制, strKey As String
    
    With obj合作单位控制集
        For Each objItem In mCol
            strKey = "K" & objItem.合作单位名称
            .AddItem objItem.Clone, strKey
        Next
        .是否独占 = mvar是否独占
        .预约控制方式 = mvar预约控制方式
        .是否修改 = mvar是否修改
    End With
    Set Clone = obj合作单位控制集
End Function

Public Function Add(合作单位名称 As String, 预约控制方式 As Integer, Optional sKey As String) As 合作单位控制
    '创建新对象
    Dim objNewMember As 合作单位控制
    Set objNewMember = New 合作单位控制


    '设置传入方法的属性
    objNewMember.合作单位名称 = 合作单位名称
    objNewMember.预约控制方式 = 预约控制方式
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function
Public Sub AddItem(obj合作单位控制 As 合作单位控制, Optional sKey As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按合作单位控制对象增加
    '编制:刘兴洪
    '日期:2016-01-13 14:00:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If obj合作单位控制 Is Nothing Then Set obj合作单位控制 = New 合作单位控制
    If Len(sKey) = 0 Then
        mCol.Add obj合作单位控制
    Else
        mCol.Add obj合作单位控制, sKey
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As 合作单位控制
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

