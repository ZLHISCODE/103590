VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"PageCell"
Attribute VB_Ext_KEY = "Member0" ,"PageCell"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'局部变量，保存集合
Private mCol As Collection

Public Function Add(ID As Integer, X As Long, Y As Long, W As Long, H As Long, MaxH As Long, RowB As Long, RowE As Long, VRowE As Long, ColB As Long, ColE As Long, FixW As Long, FixH As Long, Copys As Byte, Optional sKey As String) As PageCell
    '创建新对象
    Dim objNewMember As PageCell
    Set objNewMember = New PageCell

    '设置传入方法的属性
    objNewMember.Key = Mid(sKey, 2) 'Key为该问题关键字,应与ID对应
    objNewMember.ID = ID '数据表格控件索引
    
    '该表格在该页内输出的整体位置、尺寸(含表头)
    objNewMember.X = X
    objNewMember.Y = Y
    objNewMember.W = W
    objNewMember.H = H '数据输出可用高度(含表头),如果=FixH,表示不输出表头
    
    '该表格的设计高度(含附加表格)
    objNewMember.MaxH = MaxH
    
    objNewMember.FixW = FixW '表格固定列宽度(如果为0,表示该页没有或不打印固定列)
    objNewMember.FixH = FixH '表格固定行高度(如果为0,表示该页没有或不打印固定行)
    objNewMember.RowB = RowB '表格在该页的开始行(非固定行)
    objNewMember.RowE = RowE '表格在该页的结束行(非固定行)
    objNewMember.VRowE = VRowE '表格在该页的虚拟结束行(0,或者>RowE,可能是超出了表格行范围的)，用于分栏或者票据时处理
    objNewMember.ColB = ColB '表格在该页的开始列(非固定列)
    objNewMember.ColE = ColE '表格在该页的结束列(非固定列)
    objNewMember.Copys = Copys '表格的分栏数,该页的行起止范围已经考虑了分栏后总行
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As PageCell
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
