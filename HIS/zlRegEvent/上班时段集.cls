VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "上班时段集"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"上班时段"
Attribute VB_Ext_KEY = "Member0" ,"上班时段"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'局部变量，保存集合
Private mCol As Collection

Public Function Add(时间段 As String, 开始时间 As Date, 结束时间 As String, 缺省预约时间 As String, 提前挂号时间 As String, 出诊预留时间 As Integer, 休息时段 As String, Optional sKey As String) As 上班时段
    '创建新对象
    Dim objNewMember As 上班时段
    Set objNewMember = New 上班时段


    '设置传入方法的属性
    objNewMember.时间段 = 时间段
    objNewMember.开始时间 = 开始时间
    objNewMember.结束时间 = 结束时间
    objNewMember.缺省预约时间 = 缺省预约时间
    objNewMember.提前挂号时间 = 提前挂号时间
    objNewMember.出诊预留时间 = 出诊预留时间
    objNewMember.休息时段 = 休息时段
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Sub AddItem(obj上班时段 As 上班时段, Optional sKey As String)
    If obj上班时段 Is Nothing Then Set obj上班时段 = New 上班时段
    If Len(sKey) = 0 Then
        mCol.Add obj上班时段
    Else
        mCol.Add obj上班时段, sKey
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

Public Property Get Item(vntIndexKey As Variant) As 上班时段
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

Public Function Clone() As 上班时段集
    '创建一个副本
    Dim obj上班时段集 As New 上班时段集
    Dim objItem As 上班时段, strKey As String
    
    For Each objItem In mCol
        strKey = "K" & objItem.时间段
        obj上班时段集.AddItem objItem.Clone, strKey
    Next
    Set Clone = obj上班时段集
End Function

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

