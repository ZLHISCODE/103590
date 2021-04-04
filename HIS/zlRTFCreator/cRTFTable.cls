VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFTable"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private mvarRows As cRTFRows
Private mvarRowCount As Long
Private mvarColCount As Long
Private mvarAlignment As RowAlignTypeEnum
Private mvarNested As Boolean

Public Function LastRow() As cRTFRow
Attribute LastRow.VB_Description = "获取最后添加的行对象。"
    Set LastRow = mvarRows(mvarRows.Count)
End Function

Public Property Let Nested(ByVal vData As Boolean)
Attribute Nested.VB_Description = "是否嵌套。默认:False"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Nested = 5
    Dim i As Long
    For i = 1 To mvarRows.Count
        mvarRows(i).Nested = vData
    Next
    mvarNested = vData
End Property

Public Property Get Nested() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Nested
    Nested = mvarNested
End Property

Public Function GetTextRTF() As String
Attribute GetTextRTF.VB_Description = "获取表格RTF字符串。"
    Dim strR As String
    Dim i As Long
    For i = 1 To mvarRows.Count
        strR = strR & mvarRows(i).GetTextRTF
    Next
    GetTextRTF = strR
End Function

Public Function AddRow(Optional Row As Long, Optional RowBand As Long, _
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
Attribute AddRow.VB_Description = "向表格中添加一行。"
    
    Dim i As Long

    i = mvarRows.Add(Row, RowBand, sKey, IsLastRow, CellSpace, Width, WidthUnit, WidthB, WidthBUnit, _
        WidthA, WidthAUnit, PatType, PatForeColor, PatBackColor, PatPercentage, AutoFit, IsHeadRow, Keep, _
        KeepFollow, Alignment, Height, PadLeft, PadRight, PadTop, PadBottom, Nested, BorderLeft, _
        BorderRight, BorderTop, BorderBottom, BorderH, BorderV, BorderStyle, BorderColor, BorderWidth)
    mvarRows(i).Row = i
    mvarRows(i).RowBand = i
    mvarRowCount = i
    
    AddRow = i
End Function

Public Property Let Alignment(ByVal vData As RowAlignTypeEnum)
Attribute Alignment.VB_Description = "表格对齐方式。"
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Alignment = Form1
    Dim i As Long
    For i = 1 To mvarRows.Count
        mvarRows(i).Alignment = vData
    Next
    mvarAlignment = vData
End Property

Public Property Get Alignment() As RowAlignTypeEnum
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Alignment
    Alignment = mvarAlignment
End Property

Public Property Let ColCount(ByVal vData As Long)
Attribute ColCount.VB_Description = "列数。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.ColCount = 5
    mvarColCount = vData
End Property

Public Property Get ColCount() As Long
    '返回之后一行的Colcount数值：
    ColCount = mvarRows(mvarRows.Count).ColCount
End Property

Public Property Let RowCount(ByVal vData As Long)
Attribute RowCount.VB_Description = "行数。"
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.RowCount = 5
    mvarRowCount = vData
End Property

Public Property Get RowCount() As Long
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.RowCount
    RowCount = mvarRowCount
End Property

Public Property Set Rows(ByVal vData As cRTFRows)
Attribute Rows.VB_Description = "当前表格的所有行的集合。"
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.Rows = Form1
    Set mvarRows = vData
End Property

Public Property Get Rows() As cRTFRows
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Rows
    Set Rows = mvarRows
End Property

Public Function LastCell() As cRTFCell
Attribute LastCell.VB_Description = "获取最后添加的单元格对象。"
    Set LastCell = Rows(mvarRowCount).Cells(Rows(mvarRowCount).Cells.Count)
End Function

Public Function Cell(Optional Row As Long = 0, Optional Col As Long = 0) As cRTFCell
Attribute Cell.VB_Description = "获取指定行、列的单元格对象。"
'默认就是最后一个单元格
    If Row = 0 Then
        Row = Me.RowCount
    End If
    If Col = 0 Then
        Col = Me.ColCount
    End If
    Set Cell = Rows(Row).Cells(Col)
End Function

Private Sub Class_Initialize()
    Set mvarRows = New cRTFRows
End Sub

Private Sub Class_Terminate()
    Set mvarRows = Nothing
End Sub
