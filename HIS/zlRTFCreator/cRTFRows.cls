VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFRows"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRTFRow"
Attribute VB_Ext_KEY = "Member0" ,"cRTFRow"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'局部变量，保存集合
Private mCol As Collection

Public Function LastRow() As cRTFRow
Attribute LastRow.VB_Description = "获取最后添加的行对象。"
    Set LastRow = mCol(mCol.Count)
End Function

Public Function Add(Optional Row As Long, Optional RowBand As Long, _
Optional sKey As String, Optional IsLastRow As Boolean = False, _
Optional CellSpace As Long = 108, _
Optional Width As Long = 0, Optional WidthUnit As CellLengthUnitEnum = clt3_缇, _
Optional WidthB As Long = 0, Optional WidthBUnit As CellLengthUnitEnum = clt3_缇, _
Optional WidthA As Long = 0, Optional WidthAUnit As CellLengthUnitEnum = clt3_缇, _
Optional PatType As PatTypeEnum = ptt00_无, Optional PatForeColor As ColorEnum = clr00_默认, _
Optional PatBackColor As ColorEnum = clr00_默认, Optional PatPercentage As Long, _
Optional AutoFit As Boolean = True, Optional IsHeadRow As Boolean = False, _
Optional Keep As Boolean = False, Optional KeepFollow As Boolean = False, _
Optional Alignment As RowAlignTypeEnum = ral0_默认对齐, _
Optional Height As Long = 0, Optional PadLeft As Long = 0, _
Optional PadRight As Long = 0, Optional PadTop As Long = 0, _
Optional PadBottom As Long = 0, Optional Nested As Boolean = False, _
Optional BorderLeft As Boolean = True, Optional BorderRight As Boolean = True, _
Optional BorderTop As Boolean = True, Optional BorderBottom As Boolean = True, _
Optional BorderH As Boolean = True, Optional BorderV As Boolean = True, _
Optional BorderStyle As BorderStyleEnum = cbs00_默认边框样式, _
Optional BorderColor As ColorEnum = clr00_默认, Optional BorderWidth As Long = 15) As Long
Attribute Add.VB_Description = "添加一行。"
    '创建新对象
    Dim objNewMember As cRTFRow
    Set objNewMember = New cRTFRow


    '设置传入方法的属性
    objNewMember.Row = Row
    objNewMember.RowBand = RowBand
    objNewMember.IsLastRow = IsLastRow
    objNewMember.CellSpace = CellSpace
    objNewMember.Width = Width
    objNewMember.WidthUnit = WidthUnit
    objNewMember.WidthB = WidthB
    objNewMember.WidthBUnit = WidthBUnit
    objNewMember.WidthA = WidthA
    objNewMember.WidthAUnit = WidthAUnit
    objNewMember.PatType = PatType
    objNewMember.PatForeColor = PatForeColor
    objNewMember.PatBackColor = PatBackColor
    objNewMember.PatPercentage = PatPercentage
    objNewMember.AutoFit = AutoFit
    objNewMember.IsHeadRow = IsHeadRow
    objNewMember.Keep = Keep
    objNewMember.KeepFollow = KeepFollow
    objNewMember.Alignment = Alignment
    objNewMember.Height = Height
    objNewMember.PadLeft = PadLeft
    objNewMember.PadRight = PadRight
    objNewMember.PadTop = PadTop
    objNewMember.PadBottom = PadBottom
    objNewMember.Nested = Nested
    objNewMember.BorderLeft = BorderLeft
    objNewMember.BorderRight = BorderRight
    objNewMember.BorderTop = BorderTop
    objNewMember.BorderBottom = BorderBottom
    objNewMember.BorderH = BorderH
    objNewMember.BorderV = BorderV
    objNewMember.BorderStyle = BorderStyle
    objNewMember.BorderColor = BorderColor
    objNewMember.BorderWidth = BorderWidth
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Add = mCol.Count
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cRTFRow
Attribute Item.VB_Description = "获取一个行对象。"
Attribute Item.VB_UserMemId = 0
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
Attribute Count.VB_Description = "行总数。"
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "删除指定的行元素对象。"
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

