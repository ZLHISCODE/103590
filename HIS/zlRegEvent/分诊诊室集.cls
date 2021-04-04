VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "分诊诊室集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"分诊诊室"
Attribute VB_Ext_KEY = "Member0" ,"分诊诊室"
Option Explicit
'局部变量，保存集合
Private mCol As Collection
Private mvar分诊方式 As Byte '局部复制
Private mvar是否修改 As Boolean '局部复制
'保持属性值的局部变量
Private mvar医生姓名 As String '局部复制
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

Public Function Clone() As 分诊诊室集
    '创建一个副本
    Dim obj分诊诊室集 As New 分诊诊室集
    Dim objItem As 分诊诊室, strKey As String
    
    With obj分诊诊室集
        For Each objItem In mCol
            strKey = "K" & objItem.诊室ID
            .AddItem objItem.Clone, strKey
        Next
        .分诊方式 = mvar分诊方式
        .是否修改 = mvar是否修改
        .医生姓名 = mvar医生姓名
    End With
    Set Clone = obj分诊诊室集
End Function

Public Function Add(诊室ID As Long, 诊室名称 As String, Optional sKey As String) As 分诊诊室
    '创建新对象
    Dim objNewMember As 分诊诊室
    Set objNewMember = New 分诊诊室


    '设置传入方法的属性
    objNewMember.诊室ID = 诊室ID
    objNewMember.诊室名称 = 诊室名称
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function
Public Sub AddItem(obj分诊诊室 As 分诊诊室, Optional sKey As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按号序对象接加
    '编制:刘兴洪
    '日期:2016-01-13 14:00:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If obj分诊诊室 Is Nothing Then Set obj分诊诊室 = New 分诊诊室
    If Len(sKey) = 0 Then
        mCol.Add obj分诊诊室
    Else
        mCol.Add obj分诊诊室, sKey
    End If
End Sub
Public Property Get Item(vntIndexKey As Variant) As 分诊诊室
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
