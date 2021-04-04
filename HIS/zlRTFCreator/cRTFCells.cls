VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFCells"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRTFCell"
Attribute VB_Ext_KEY = "Member0" ,"cRTFCell"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function LastCell() As cRTFCell
Attribute LastCell.VB_Description = "获取最后一个单元格元素。"
    Set LastCell = mCol(mCol.Count)
End Function

Public Function Add(Row As Long, Col As Long, _
Optional DataType As CellContentTypeEnum = cct0_文本, _
Optional Right As Long, _
Optional Width As Long = 0, _
Optional WidthUnit As CellLengthUnitEnum = clt0_默认, _
Optional sKey As String, _
Optional AlignmentV As AlignVTypeEnum = alv0_默认对齐, _
Optional Alignment As AlignTypeEnum = alm0_默认对齐, Optional BorderLeft As Boolean = True, _
Optional BorderRight As Boolean = True, Optional BorderTop As Boolean = True, _
Optional BorderBottom As Boolean = True, Optional BorderLeftDown As Boolean = False, _
Optional BorderLeftUp As Boolean = False, Optional BorderStyle As BorderStyleEnum = cbs00_默认边框样式, _
Optional BorderColor As ColorEnum = clr00_默认, Optional CellPatType As PatTypeEnum = ptt00_无, _
Optional CellPatBackColor As ColorEnum = clr00_默认, Optional CellPatForeColor As ColorEnum = clr00_默认, _
Optional MergeVStart As Boolean = False, Optional MergeV As Boolean = False, _
Optional MergeStart As Boolean = False, Optional Merge As Boolean = False, _
Optional FitText As Boolean = True, Optional NoWrap As Boolean = False, _
Optional Nested As Boolean = False, Optional BorderWidth As Long = 0, _
Optional CellPatPercentage As Long = 100) As Long
Attribute Add.VB_Description = "添加一个单元格元素。"
    '创建新对象
    Dim objNewMember As cRTFCell
    Set objNewMember = New cRTFCell


    '设置传入方法的属性
    objNewMember.AlignmentV = AlignmentV
    objNewMember.Alignment = Alignment
    objNewMember.BorderLeft = BorderLeft
    objNewMember.BorderRight = BorderRight
    objNewMember.BorderTop = BorderTop
    objNewMember.BorderBottom = BorderBottom
    objNewMember.BorderLeftDown = BorderLeftDown
    objNewMember.BorderLeftUp = BorderLeftUp
    objNewMember.BorderStyle = BorderStyle
    objNewMember.BorderColor = BorderColor
    objNewMember.CellPatType = CellPatType
    objNewMember.CellPatBackColor = CellPatBackColor
    objNewMember.CellPatForeColor = CellPatForeColor
    objNewMember.Width = Width
    objNewMember.WidthUnit = WidthUnit
    objNewMember.Right = Right
    objNewMember.MergeVStart = MergeVStart
    objNewMember.MergeV = MergeV
    objNewMember.MergeStart = MergeStart
    objNewMember.Merge = Merge
    objNewMember.Row = Row
    objNewMember.Col = Col
    objNewMember.DataType = DataType
    objNewMember.FitText = FitText
    objNewMember.NoWrap = NoWrap
    objNewMember.Nested = Nested
    objNewMember.BorderWidth = BorderWidth
    objNewMember.CellPatPercentage = CellPatPercentage
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Add = mCol.Count
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cRTFCell
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
Attribute Remove.VB_Description = "删除一个单元格元素。"
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "本属性允许用 For...Each 语法枚举该集合。"
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

