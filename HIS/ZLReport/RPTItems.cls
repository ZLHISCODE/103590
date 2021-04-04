VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPTItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"RPTItem"
Attribute VB_Ext_KEY = "Member0" ,"RPTItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function Add(id As Integer, 格式号 As Byte, 名称 As String, 上级ID As Integer, 类型 As Byte, _
    序号 As Integer, 参照 As String, 性质 As Byte, 内容 As String, 表头 As String, X As Long, _
    Y As Long, W As Long, H As Long, 行高 As Long, 对齐 As Byte, 自调 As Boolean, 字体 As String, _
    字号 As Single, 粗体 As Boolean, 下线 As Boolean, 斜体 As Boolean, 网格 As Long, 前景 As Long, _
    背景 As Long, 边框 As Boolean, 分栏 As Byte, 排序 As String, 格式 As String, 汇总 As String, _
    表格线加粗 As Boolean, 自适应行高 As Boolean, _
    Optional 图片 As StdPicture, Optional 系统 As Boolean, Optional 父ID As Integer, Optional SubIDs As RelatIDs, _
    Optional CopyIDs As RelatIDs, Optional sKey As String, Optional 数据源 As String, Optional 上下间距 As Long, _
    Optional 左右间距 As Long, Optional 源行号 As Long, Optional 横向分栏 As Long, Optional 纵向分栏 As Long, _
    Optional Relations As RPTRelations, Optional ColProtertys As RPTColProtertys, _
    Optional 水平反转 As Boolean) As RPTItem
    
    '创建新对象
    Dim objNewMember As RPTItem
    Set objNewMember = New RPTItem

    '如果已经复制了，就不再复制了
    On Error Resume Next
    If Not mCol.Item(sKey) Is Nothing Then
        If Err.Number = 0 Then
            Set objNewMember.CopyIDs = CopyIDs
            Set Add = objNewMember
            Exit Function
        End If
        Err.Clear: On Error GoTo 0
    End If
    '设置传入方法的属性
    objNewMember.Key = Mid(sKey, 2)
    objNewMember.id = id
    objNewMember.格式号 = 格式号
    objNewMember.名称 = 名称
    objNewMember.上级ID = 上级ID
    objNewMember.类型 = 类型
    objNewMember.序号 = 序号
    objNewMember.参照 = 参照
    objNewMember.性质 = 性质
    objNewMember.内容 = 内容
    objNewMember.表头 = 表头
    objNewMember.X = X
    objNewMember.Y = Y
    objNewMember.W = W
    objNewMember.H = H
    objNewMember.行高 = 行高
    objNewMember.自适应行高 = 自适应行高
    objNewMember.对齐 = 对齐
    objNewMember.自调 = 自调
    objNewMember.字体 = 字体
    objNewMember.字号 = 字号
    objNewMember.粗体 = 粗体
    objNewMember.表格线加粗 = 表格线加粗
    objNewMember.下线 = 下线
    objNewMember.斜体 = 斜体
    objNewMember.网格 = 网格
    objNewMember.前景 = 前景
    objNewMember.背景 = 背景
    objNewMember.边框 = 边框
    objNewMember.分栏 = 分栏
    objNewMember.排序 = 排序
    objNewMember.格式 = 格式
    objNewMember.汇总 = 汇总
    objNewMember.系统 = 系统
    objNewMember.父ID = 父ID
    objNewMember.数据源 = 数据源
    objNewMember.上下间距 = 上下间距
    objNewMember.左右间距 = 左右间距
    objNewMember.源行号 = 源行号
    objNewMember.纵向分栏 = 纵向分栏
    objNewMember.横向分栏 = 横向分栏
    objNewMember.水平反转 = 水平反转
    
    '无参数时自动为Nothing
    Set objNewMember.图片 = 图片
    
    '无参数时自动为New
    If Not CopyIDs Is Nothing Then Set objNewMember.CopyIDs = CopyIDs
    If Not SubIDs Is Nothing Then Set objNewMember.SubIDs = SubIDs
    If Not Relations Is Nothing Then Set objNewMember.Relations = Relations
    If Not ColProtertys Is Nothing Then Set objNewMember.ColProtertys = ColProtertys
    
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As RPTItem
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error Resume Next
    Set Item = mCol(vntIndexKey)
    On Error GoTo 0
End Property

Public Property Get count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    count = mCol.count
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
