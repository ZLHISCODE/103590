VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl EGrid 
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   ScaleHeight     =   4995
   ScaleWidth      =   6750
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton SelCmd 
      Caption         =   "..."
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4020
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox EditBox1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4050
      Visible         =   0   'False
      Width           =   1515
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   7011
      _Version        =   393216
      AllowBigSelection=   0   'False
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
   End
End
Attribute VB_Name = "EGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim DefFont As StdFont
Public AddFlag As Boolean, iCol As Long, iRow As Long
Attribute AddFlag.VB_VarMemberFlags = "40"
Attribute iCol.VB_VarMemberFlags = "40"
Attribute iRow.VB_VarMemberFlags = "40"
Public OldForeColor As OLE_COLOR, OldBackColor As OLE_COLOR, OldResize As flexResize
Attribute OldForeColor.VB_VarMemberFlags = "40"
Attribute OldBackColor.VB_VarMemberFlags = "40"
Attribute OldResize.VB_VarMemberFlags = "40"
Private Columns() As CellAttribute, Cells() As CellAttribute
Dim SelFlag() As Boolean
Attribute SelFlag.VB_VarMemberFlags = "40"
Private EditBox As VB.Control
Private bAllowAddNew As Boolean
Private CellItems() As CellItem
Private ColumnItems() As CellItem

Public Event BeforeRowUpdate(ByVal RowIndex As Long, ByVal ModifyMode As String, Cancel As Boolean)

Public Event Click()
Attribute Click.VB_Description = "点击对象时触发该事件。"

Public Event BeforeColUpdate(ByVal RowIndex As Long, ByVal ColIndex As Long, NewValue As String, ByVal OldValue As String, Cancel As Boolean)

Public Event DblClick()
Attribute DblClick.VB_Description = "双击对象时触发该事件。"

Public Event SelClick(ByVal RowIndex As Long, ByVal ColIndex As Long, ByRef EditText As String)

Public Event KeyPress(KeyAscii As Integer)

Public Event ListClick(ByVal RowIndex As Long, ByVal ColIndex As Long)

Public Event EditEnd()

Private Sub EndEdit()
    Grid.AllowUserResizing = OldResize
    Grid.BackColorSel = OldBackColor
    Grid.ForeColorSel = OldForeColor
    EditBox.Visible = False
    SelCmd.Visible = False
    Grid.SetFocus
    
    RaiseEvent EditEnd
End Sub

Private Sub Combo1_Click()
    On Error Resume Next
    Cells(iCol, iRow).ListIndex = Combo1.ListIndex
    RaiseEvent ListClick(iRow, iCol)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    EditBox1_KeyPress KeyAscii
End Sub

Private Sub EditBox1_KeyPress(KeyAscii As Integer)
    Dim OldValue As String, Cancel As Boolean, NewValue As String
    Dim OldCol As Long
'    DoEvents
    RaiseEvent KeyPress(KeyAscii)
    OldValue = Grid.Text
    NewValue = EditBox
    Cancel = False
    If KeyAscii = 13 Then
        DoEvents
        RaiseEvent BeforeColUpdate(iRow, iCol, NewValue, OldValue, Cancel)
        If Cancel Then Exit Sub
        Grid.Text = NewValue
        
        OldCol = iCol
        'iCol = iCol + 1
        iCol = getNextCol(iCol, iRow)
        If iCol >= Grid.Cols Then
            DoEvents
            If AddFlag Then
                RaiseEvent BeforeRowUpdate(iRow, "Add", Cancel)
            Else
                RaiseEvent BeforeRowUpdate(iRow, "Edit", Cancel)
            End If
            If Cancel Then
                EditBox.SetFocus
                'iCol = iCol - 1
                iCol = OldCol
                Exit Sub
            End If
            iRow = iRow + 1
            'iCol = Grid.FixedCols
            If (AddFlag Or iRow > Grid.Rows - 1) And bAllowAddNew Then
                AddFlag = True
                Grid.AddItem "" & vbTab & "", iRow
                
                AfterAddRow iRow
            ElseIf iRow > Grid.Rows - 1 Then iRow = iRow - 1
            End If
            iCol = getFirstCol(iRow)
        End If
        Grid.AllowUserResizing = OldResize
        Edit iRow, iCol
    ElseIf KeyAscii = 27 Then
        If AddFlag Then
            If Grid.Rows > Grid.FixedRows + 1 Then
                Grid.RemoveItem iRow
                
                AfterDeleteRow iRow
            Else
                Grid.AddItem "" & vbTab & "", iRow
                AfterAddRow iRow
                Grid.RemoveItem iRow + 1
                AfterDeleteRow iRow + 1
            End If
            AddFlag = False
            EndEdit
        Else
'            ValidValue
            EndEdit
        End If
    End If
End Sub

Private Sub Grid_Click()
    DoEvents
    RaiseEvent Click
End Sub

Private Sub Grid_DblClick()
    DoEvents
    RaiseEvent DblClick
End Sub

Private Sub Grid_GotFocus()
    Dim iNewCol As Long, iNewRow As Long, NewValue As String
    Dim OldValue As String, Cancel As Boolean
    If EditBox.Visible Then
        iNewCol = Grid.Col
        iNewRow = Grid.Row
        Grid.Col = iCol
        Grid.Row = iRow
        OldValue = Grid.Text
        NewValue = EditBox
        Cancel = False
        DoEvents
        RaiseEvent BeforeColUpdate(iRow, iCol, NewValue, OldValue, Cancel)
        If Cancel Then
            EditBox.SetFocus
            Exit Sub
        End If
        Grid.Text = NewValue
        If iRow <> iNewRow Then
            DoEvents
            If AddFlag Then
                RaiseEvent BeforeRowUpdate(iRow, "Add", Cancel)
            Else
                RaiseEvent BeforeRowUpdate(iRow, "Edit", Cancel)
            End If
            If Cancel Then
                EditBox.SetFocus
                Exit Sub
            End If
            AddFlag = False
        End If
'        Grid.Text = NewValue
        Grid.AllowUserResizing = OldResize
        Edit iNewRow, iNewCol
        'ValidValue
    End If
End Sub

Private Sub Grid_Scroll()
    If EditBox.Visible Then
        EditBox.Left = Grid.Left + Grid.CellLeft + 30
        EditBox.SetFocus
    End If
End Sub

Private Sub SelCmd_Click()
    Dim EditText As String
    EditBox.SetFocus
    DoEvents
    EditText = EditBox
    RaiseEvent SelClick(iRow, iCol, EditText)
    EditBox = EditText
End Sub

Private Sub UserControl_Initialize()
    bAllowAddNew = True
    
    ReDim CellItems(0)
    ReDim Preserve Columns(Grid.Cols - 1)
    ReDim Preserve Cells(Grid.Cols - 1, Grid.Rows - 1)
    ReDim Preserve ColumnItems(Grid.Cols - 1)
End Sub

Private Sub UserControl_InitProperties()
    Set EditBox = EditBox1
    
    AddFlag = False
    bAllowAddNew = True
    Set DefFont = New StdFont
    DefFont.Name = "宋体"
    DefFont.Size = 9
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ReDim SelFlag(0)
    Set EditBox = EditBox1
    
    Set Grid.Font = PropBag.ReadProperty("Font", DefFont)
    Set Grid.FontFixed = PropBag.ReadProperty("FontFixed", DefFont)
    Set EditBox.Font = Grid.Font
    Grid.BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    Grid.BackColorFixed = PropBag.ReadProperty("BackColorFixed", vbButtonFace)
    Grid.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    Grid.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", vbButtonText)
    EditBox.BackColor = PropBag.ReadProperty("EditBackColor", vbWindowBackground)
    EditBox.ForeColor = PropBag.ReadProperty("EditForeColor", vbButtonText)
    
    Grid.Rows = PropBag.ReadProperty("Rows", 2)
    SetCols PropBag.ReadProperty("Cols", 2)
    
    Grid.FixedCols = PropBag.ReadProperty("FixedCols", 1)
    Grid.FixedRows = PropBag.ReadProperty("FixedRows", 1)
    Grid.WordWrap = PropBag.ReadProperty("WordWrap", False)
    Grid.FocusRect = PropBag.ReadProperty("FocusRect", flexFocusLight)
    Grid.FormatString = PropBag.ReadProperty("FormatString", "")
    Grid.MergeCells = PropBag.ReadProperty("MergeCells", flexMergeNever)
    Grid.RowSizingMode = PropBag.ReadProperty("RowSizingMode", flexRowSizeAll)
    Grid.TextStyle = PropBag.ReadProperty("TextStyle", flexTextFlat)
    Grid.TextStyleFixed = PropBag.ReadProperty("TextStyleFixed", flexTextFlat)
    Grid.AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", flexResizeNone)
    Grid.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 350)
    Grid.GridLines = PropBag.ReadProperty("GridLines", flexGridFlat)
    Grid.GridLinesFixed = PropBag.ReadProperty("GridLinesFixed", flexGridInset)
    Grid.GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
    Grid.GridLineWidthFixed = PropBag.ReadProperty("GridLineWidthFixed", 1)
    Grid.BackColorSel = PropBag.ReadProperty("SelBackColor", vbHighlight)
    Grid.ForeColorSel = PropBag.ReadProperty("SelForeColor", vbHighlightText)
    Grid.HighLight = PropBag.ReadProperty("HighLight", 0)
    OldBackColor = Grid.BackColorSel
    OldForeColor = Grid.ForeColorSel
End Sub

Private Sub UserControl_Resize()
    Grid.Top = 0
    Grid.Left = 0
    Grid.Width = UserControl.Width
    Grid.Height = UserControl.Height
End Sub

Public Property Get Rows() As Long
Attribute Rows.VB_Description = "返回或设置表格总行数。"
Attribute Rows.VB_ProcData.VB_Invoke_Property = ";外观"
    Rows = Grid.Rows
End Property

Public Property Let Rows(ByVal vNewValue As Long)
    Dim iOldRows As Long, i As Long, j As Long, ItemID As Long, iNum As Long
    On Error Resume Next
    
    iOldRows = Grid.Rows
    
    Grid.Rows = vNewValue
    ReDim Preserve Cells(Grid.Cols - 1, Grid.Rows - 1)
    
    For i = iOldRows To Grid.Rows - 1
        For j = 0 To Grid.Cols - 1
            Cells(j, i).Disabled = Columns(j).Disabled
            Cells(j, i).EditMode = Columns(j).EditMode
            
            iNum = -1
            iNum = UBound(ColumnItems(j).List)
            For ItemID = 0 To iNum
                List_AddItem i, j, ColumnItems(j).List(ItemID)
            Next ItemID
        Next j
    Next i
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回或设置对象的背景颜色。"
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";外观"
    BackColor = Grid.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    Grid.BackColor = vNewValue
End Property

Public Property Get Cols() As Long
Attribute Cols.VB_Description = "返回或设置对象的总列数。"
Attribute Cols.VB_ProcData.VB_Invoke_Property = ";外观"
    Cols = Grid.Cols
End Property

Public Property Let Cols(ByVal vNewValue As Long)
    SetCols vNewValue
End Property

Private Sub SetCols(ByVal vNewValue As Long)
    Dim tmpCells() As CellAttribute, iLoop As Long, jLoop As Long
    ReDim tmpCells(Grid.Cols - 1, Grid.Rows - 1)
    For iLoop = 0 To Grid.Cols - 1
        For jLoop = 0 To Grid.Rows - 1
            tmpCells(iLoop, jLoop).Disabled = Cells(iLoop, jLoop).Disabled
            tmpCells(iLoop, jLoop).EditMode = Cells(iLoop, jLoop).EditMode
            tmpCells(iLoop, jLoop).ItemIndex = Cells(iLoop, jLoop).ItemIndex
            tmpCells(iLoop, jLoop).ListIndex = Cells(iLoop, jLoop).ListIndex
        Next
    Next
    Grid.Cols = vNewValue
    ReDim Preserve Columns(Grid.Cols - 1)
    ReDim Preserve ColumnItems(Grid.Cols - 1)
    Erase Cells
    ReDim Cells(Grid.Cols - 1, Grid.Rows - 1)
    For iLoop = 0 To IIf(UBound(tmpCells, 1) > Grid.Cols - 1, Grid.Cols - 1, UBound(tmpCells, 1) > Grid.Cols - 1)
        For jLoop = 0 To Grid.Rows - 1
            Cells(iLoop, jLoop).Disabled = tmpCells(iLoop, jLoop).Disabled
            Cells(iLoop, jLoop).EditMode = tmpCells(iLoop, jLoop).EditMode
            Cells(iLoop, jLoop).ItemIndex = tmpCells(iLoop, jLoop).ItemIndex
            Cells(iLoop, jLoop).ListIndex = tmpCells(iLoop, jLoop).ListIndex
        Next
    Next
End Sub
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "返回或设置单元格文本是否自动换行。"
Attribute WordWrap.VB_ProcData.VB_Invoke_Property = ";行为"
    WordWrap = Grid.WordWrap
End Property

Public Property Let WordWrap(ByVal vNewValue As Boolean)
    Grid.WordWrap = vNewValue
End Property

Public Property Get FixedRows() As Long
Attribute FixedRows.VB_Description = "返回或设置对象中固定行行数。"
Attribute FixedRows.VB_ProcData.VB_Invoke_Property = ";外观"
    FixedRows = Grid.FixedRows
End Property

Public Property Let FixedRows(ByVal vNewValue As Long)
    Dim iOldRows As Long, i As Long, j As Long, ItemID As Long, iNum As Long
    On Error Resume Next
    iOldRows = Grid.Rows
    
    If vNewValue >= Grid.Rows Then Grid.Rows = vNewValue + 1
    Grid.FixedRows = vNewValue
    ReDim Preserve Cells(Grid.Cols - 1, Grid.Rows - 1)

    For i = iOldRows To Grid.Rows - 1
        For j = 0 To Grid.Cols - 1
            Cells(j, i).Disabled = Columns(j).Disabled
            Cells(j, i).EditMode = Columns(j).EditMode
            
            iNum = -1
            iNum = UBound(ColumnItems(j).List)
            For ItemID = 0 To iNum
                List_AddItem i, j, ColumnItems(j).List(ItemID)
            Next ItemID
        Next j
    Next i
End Property

Public Property Get FixedCols() As Long
Attribute FixedCols.VB_Description = "返回或设置对象中固定列列数。"
Attribute FixedCols.VB_ProcData.VB_Invoke_Property = ";外观"
    FixedCols = Grid.FixedCols
End Property

Public Property Let FixedCols(ByVal vNewValue As Long)
    If vNewValue >= Grid.Cols Then SetCols vNewValue + 1
    Grid.FixedCols = vNewValue
End Property

Public Property Get AllowUserResizing() As flexResize
Attribute AllowUserResizing.VB_Description = "返回或设置是否允许用户调整表格行高和列宽。"
Attribute AllowUserResizing.VB_ProcData.VB_Invoke_Property = ";行为"
    AllowUserResizing = Grid.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal vNewValue As flexResize)
    Grid.AllowUserResizing = vNewValue
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "返回或设置对象字体及字体大小。"
Attribute Font.VB_ProcData.VB_Invoke_Property = ";外观"
    Set Font = Grid.Font
End Property

Public Property Set Font(ByVal vNewValue As Font)
    Set Grid.Font = vNewValue
    Set EditBox.Font = vNewValue
    Set Grid.FontFixed = vNewValue
End Property

Public Property Get FontFixed() As Font
Attribute FontFixed.VB_Description = "返回或设置固定单元格的字体及字体大小。"
Attribute FontFixed.VB_ProcData.VB_Invoke_Property = ";外观"
    Set FontFixed = Grid.FontFixed
End Property

Public Property Set FontFixed(ByVal vNewValue As Font)
    Set Grid.FontFixed = vNewValue
End Property

Public Property Get BackColorFixed() As OLE_COLOR
Attribute BackColorFixed.VB_Description = "返回或设置对象中固定单元格的背景颜色。"
Attribute BackColorFixed.VB_ProcData.VB_Invoke_Property = ";外观"
    BackColorFixed = Grid.BackColorFixed
End Property

Public Property Let BackColorFixed(ByVal vNewValue As OLE_COLOR)
    Grid.BackColorFixed = vNewValue
End Property

Public Property Get FocusRect() As flexFocus
Attribute FocusRect.VB_Description = "返回或设置活动单元格模式。"
Attribute FocusRect.VB_ProcData.VB_Invoke_Property = ";外观"
    FocusRect = Grid.FocusRect
End Property

Public Property Let FocusRect(ByVal vNewValue As flexFocus)
    Grid.FocusRect = vNewValue
End Property

Public Property Get EditBackColor() As OLE_COLOR
Attribute EditBackColor.VB_Description = "返回或设置对象的编辑框背景颜色。"
Attribute EditBackColor.VB_ProcData.VB_Invoke_Property = ";外观"
    EditBackColor = EditBox.BackColor
End Property

Public Property Let EditBackColor(ByVal vNewValue As OLE_COLOR)
    EditBox.BackColor = vNewValue
End Property

Public Property Get EditForeColor() As OLE_COLOR
Attribute EditForeColor.VB_Description = "返回或设置对象的编辑框前景颜色。"
Attribute EditForeColor.VB_ProcData.VB_Invoke_Property = ";外观"
    EditForeColor = EditBox.ForeColor
End Property

Public Property Let EditForeColor(ByVal vNewValue As OLE_COLOR)
    EditBox.ForeColor = vNewValue
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回或设置对象中文字颜色。"
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";外观"
    ForeColor = Grid.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    Grid.ForeColor = vNewValue
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
Attribute ForeColorFixed.VB_Description = "返回或设置对象中固定单元的文字颜色。"
Attribute ForeColorFixed.VB_ProcData.VB_Invoke_Property = ";外观"
    ForeColorFixed = Grid.ForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal vNewValue As OLE_COLOR)
    Grid.ForeColorFixed = vNewValue
End Property

Public Property Get FormatString() As String
Attribute FormatString.VB_Description = "设置或返回对象表头字符串。"
Attribute FormatString.VB_ProcData.VB_Invoke_Property = ";外观"
    FormatString = Grid.FormatString
End Property

Public Property Let FormatString(ByVal vNewValue As String)
    Grid.FormatString = vNewValue
End Property

Public Property Get GridLines() As flexGridLine
Attribute GridLines.VB_Description = "设置或返回表格线型。"
Attribute GridLines.VB_ProcData.VB_Invoke_Property = ";外观"
    GridLines = Grid.GridLines
End Property

Public Property Let GridLines(ByVal vNewValue As flexGridLine)
    Grid.GridLines = vNewValue
End Property

Public Property Get GridLinesFixed() As flexGridLine
Attribute GridLinesFixed.VB_Description = "设置或返回表格中固定单元格线型。"
Attribute GridLinesFixed.VB_ProcData.VB_Invoke_Property = ";外观"
    GridLinesFixed = Grid.GridLinesFixed
End Property

Public Property Let GridLinesFixed(ByVal vNewValue As flexGridLine)
    Grid.GridLinesFixed = vNewValue
End Property

Public Property Get GridLineWidth() As Integer
Attribute GridLineWidth.VB_Description = "返回或设置表格线宽度。"
Attribute GridLineWidth.VB_ProcData.VB_Invoke_Property = ";外观"
    GridLineWidth = Grid.GridLineWidth
End Property

Public Property Let GridLineWidth(ByVal vNewValue As Integer)
    Grid.GridLineWidth = vNewValue
End Property

Public Property Get HighLight() As flexHighLight
Attribute HighLight.VB_Description = "决定选定的单元格是否在表格中突出显示"
Attribute HighLight.VB_ProcData.VB_Invoke_Property = ";外观"
    HighLight = Grid.HighLight
End Property

Public Property Let HighLight(ByVal vNewValue As flexHighLight)
    Grid.HighLight = vNewValue
End Property

Public Property Get MergeCells() As flexMerge
Attribute MergeCells.VB_Description = "返回或设置是否合并相同的相邻单元格。"
Attribute MergeCells.VB_ProcData.VB_Invoke_Property = ";行为"
    MergeCells = Grid.MergeCells
End Property

Public Property Let MergeCells(ByVal vNewValue As flexMerge)
    Grid.MergeCells = vNewValue
End Property

Public Property Get RowHeightMin() As Integer
Attribute RowHeightMin.VB_Description = "返回或设置表格的最小行高。"
Attribute RowHeightMin.VB_ProcData.VB_Invoke_Property = ";外观"
    RowHeightMin = Grid.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal vNewValue As Integer)
    If vNewValue < 0 Then
        Grid.RowHeightMin = 0
    Else
        Grid.RowHeightMin = vNewValue
    End If
End Property

Public Property Get RowHeight(RowNum As Long) As Integer
Attribute RowHeight.VB_MemberFlags = "400"
    If RowNum > Grid.Rows - 1 Then
        RowHeight = 0
    Else
        RowHeight = Grid.RowHeight(RowNum)
    End If
End Property

Public Property Let RowHeight(RowNum As Long, ByVal vNewValue As Integer)
    If RowNum < Grid.Rows And RowNum > 0 Then
        If Grid.RowHeightMin <> 0 Then
            If vNewValue < Grid.RowHeightMin Then vNewValue = Grid.RowHeightMin
        End If
        Grid.RowHeight(RowNum) = vNewValue
    End If
End Property

Public Property Get ColWidth(ColNum As Long) As Long
    If ColNum > Grid.Cols - 1 Then
        ColWidth = 0
    Else
        ColWidth = Grid.ColWidth(ColNum)
    End If
End Property

Public Property Let ColWidth(ColNum As Long, ByVal vNewValue As Long)
    Dim vOldValue As Long
    If ColNum < Grid.Cols And ColNum >= 0 And vNewValue >= 0 Then
        vOldValue = Grid.ColWidth(ColNum)
        Grid.ColWidth(ColNum) = vNewValue
        
        EditBox.Width = EditBox.Width + vNewValue - vOldValue
    End If
End Property

Public Property Get RowSizingMode() As flexRowSize
Attribute RowSizingMode.VB_Description = "返回或设置表格行调整模式。"
Attribute RowSizingMode.VB_ProcData.VB_Invoke_Property = ";行为"
    RowSizingMode = Grid.RowSizingMode
End Property

Public Property Let RowSizingMode(ByVal vNewValue As flexRowSize)
    Grid.RowSizingMode = vNewValue
End Property

Public Property Get TextStyle() As flexText
Attribute TextStyle.VB_Description = "返回或设置表格中文本的3D样式。"
Attribute TextStyle.VB_ProcData.VB_Invoke_Property = ";外观"
    TextStyle = Grid.TextStyle
End Property

Public Property Let TextStyle(ByVal vNewValue As flexText)
    Grid.TextStyle = vNewValue
End Property

Public Property Get TextStyleFixed() As flexText
Attribute TextStyleFixed.VB_Description = "返回或设置表格中固定单元格的文本3D样式。"
Attribute TextStyleFixed.VB_ProcData.VB_Invoke_Property = ";外观"
    TextStyleFixed = Grid.TextStyleFixed
End Property

Public Property Let TextStyleFixed(ByVal vNewValue As flexText)
    Grid.TextStyleFixed = vNewValue
End Property

Public Property Get ColAlignment(ColNum As Long) As flexAlign
Attribute ColAlignment.VB_MemberFlags = "400"
    If ColNum >= 0 And ColNum < Grid.Cols Then
        ColAlignment = Grid.ColAlignment(ColNum)
    End If
End Property

Public Property Let ColAlignment(Optional ColNum As Long = -1, ByVal vNewValue As flexAlign)
    If ColNum >= 0 And ColNum < Grid.Cols Then
        Grid.ColAlignment(ColNum) = vNewValue
    Else
        Grid.ColAlignment = vNewValue
    End If
End Property

Public Property Get ColAlignmentFixed(ColNum As Long) As flexAlign
Attribute ColAlignmentFixed.VB_MemberFlags = "400"
    If ColNum >= 0 And ColNum < Grid.Cols Then
        ColAlignmentFixed = Grid.ColAlignmentFixed(ColNum)
    End If
End Property

Public Property Let ColAlignmentFixed(Optional ColNum As Long = -1, ByVal vNewValue As flexAlign)
    If ColNum >= 0 And ColNum < Grid.Cols Then
        Grid.ColAlignmentFixed(ColNum) = vNewValue
    Else
        Grid.ColAlignmentFixed = vNewValue
    End If
End Property

Public Property Get CellHeight() As Integer
Attribute CellHeight.VB_MemberFlags = "400"
    CellHeight = Grid.CellHeight
End Property

Public Property Get CellWidth() As Integer
Attribute CellWidth.VB_MemberFlags = "400"
    CellWidth = Grid.CellWidth
End Property

Public Property Get Col() As Long
Attribute Col.VB_MemberFlags = "400"
    Col = Grid.Col
End Property

Public Property Let Col(ByVal vNewValue As Long)
    If vNewValue >= 0 And vNewValue < Grid.Cols Then
        Grid.Col = vNewValue
    End If
End Property

Public Property Get Row() As Long
Attribute Row.VB_MemberFlags = "400"
    Row = Grid.Row
End Property

Public Property Let Row(ByVal vNewValue As Long)
    If vNewValue >= 0 And vNewValue < Grid.Rows Then
        Grid.Row = vNewValue
    End If
End Property

Public Property Get Text(Optional RowNum As Long = -1, Optional ColNum As Long = -1) As String
Attribute Text.VB_MemberFlags = "400"
    On Error Resume Next
    If RowNum = -1 Then RowNum = Grid.Row
    If ColNum = -1 Then ColNum = Grid.Col
    Text = Grid.TextMatrix(RowNum, ColNum)
End Property

Public Property Let Text(Optional RowNum As Long, Optional ColNum As Long, ByVal vNewValue As String)
    On Error Resume Next
    If RowNum = -1 Then RowNum = Grid.Row
    If ColNum = -1 Then ColNum = Grid.Col
    Grid.TextMatrix(RowNum, ColNum) = vNewValue
    If EditBox.Visible And RowNum = iRow And ColNum = iCol Then EditBox = vNewValue
    
    Set_List_Text RowNum, ColNum, vNewValue
End Property

Public Sub AddItem(ItemStr As String, Optional Index As Long = -1)
Attribute AddItem.VB_Description = "向对象中加入行"
    If Grid.FixedRows - 1 >= Index And Index <> -1 Then Index = Grid.FixedRows
    If Index < 0 Or Index >= Grid.Rows Then
        Grid.AddItem ItemStr
        AfterAddRow
    Else
        Grid.AddItem ItemStr, Index
        AfterAddRow Index
    End If
End Sub

Public Sub Clear()
Attribute Clear.VB_Description = "清除对象中的所有文本。"
    If EditBox.Visible Then EndEdit
    Grid.Clear
End Sub

Public Sub ClearStructure()
    Grid.ClearStructure
End Sub

Public Sub Refresh()
    Grid.Refresh
End Sub

Public Sub RemoveItem(Index As Long)
    If Index <= Grid.FixedRows - 1 Then Index = Grid.FixedRows
    If Index > Grid.Rows - 1 Then Index = Grid.Rows - 1
    Grid.RemoveItem Index
    AfterDeleteRow Index
End Sub

Public Sub Edit(Optional RowNum As Long = -1, Optional ColNum As Long = -1)
    On Error Resume Next
    DoEvents
    OldResize = Grid.AllowUserResizing
    Grid.AllowUserResizing = flexResizeNone
    EditBox.Text = ""
    If Not EditBox.Visible Then
        OldBackColor = Grid.BackColorSel
        OldForeColor = Grid.ForeColorSel
    End If
    Grid.BackColorSel = EditBox.BackColor
    Grid.ForeColorSel = EditBox.ForeColor
    
    ColNum = getNearestCol(ColNum, RowNum)
    If RowNum < 0 Or RowNum > Grid.Rows - 1 Or RowNum < Grid.FixedRows Or ColNum < 0 Or ColNum > Grid.Cols - 1 Or ColNum < Grid.FixedCols Then
        If RowNum < 0 Or RowNum > Grid.Rows - 1 Or RowNum < Grid.FixedRows Then
            RowNum = Grid.FixedRows
        End If
        If ColNum < 0 Or ColNum > Grid.Cols - 1 Or ColNum < Grid.FixedCols Then
            ColNum = getFirstCol(RowNum)
        End If
    End If
    Grid.Col = ColNum
    Grid.Row = RowNum
    
    iCol = Grid.Col
    iRow = Grid.Row
    BEdit
End Sub

Private Sub BEdit()
    Dim Style As Boolean
    Dim i As Long, iNum As Long
    On Error Resume Next
    
    Select Case Cells(iCol, iRow).EditMode
        Case editTextBox
            Combo1.Visible = False
            
            Set EditBox = EditBox1
            EditBox.Height = Screen.TwipsPerPixelY * EditBox.Font.Size
            EditBox.Text = Grid.Text
        Case editComboBox
            EditBox1.Visible = False
            
            Combo1.Clear
            If Cells(iCol, iRow).ItemIndex > 0 Then
                iNum = -1
                iNum = UBound(CellItems(Cells(iCol, iRow).ItemIndex).List)
                For i = 0 To iNum
                    Combo1.AddItem CellItems(Cells(iCol, iRow).ItemIndex).List(i)
                Next
'                Combo1.ListIndex = Cells(iCol, iRow).ListIndex
                Combo1.Text = Grid.TextMatrix(iRow, iCol)
            End If
            Set EditBox = Combo1
            'EditBox.Text = Grid.Text
        Case editDate
    End Select
    If iCol > UBound(SelFlag) Then
        Style = False
    Else
        Style = SelFlag(iCol)
    End If
    If Style And Cells(iCol, iRow).EditMode = editTextBox Then
        SelCmd.Height = Grid.CellHeight
        SelCmd.Width = Grid.CellHeight
        EditBox.Width = Grid.CellWidth - SelCmd.Width - 60
    Else
        EditBox.Width = Grid.CellWidth - 60
    End If
    EditBox.Top = Grid.Top + Grid.CellTop + (Grid.CellHeight - EditBox.Height) / 2 - 15
    EditBox.Left = Grid.Left + Grid.CellLeft + 30
    SelCmd.Left = EditBox.Left + EditBox.Width
    SelCmd.Top = Grid.CellTop
    EditBox.Visible = True
    SelCmd.Visible = Style And Cells(iCol, iRow).EditMode = editTextBox
    EditBox.SetFocus
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Font", Grid.Font, DefFont
    PropBag.WriteProperty "FontFixed", Grid.FontFixed, DefFont
    PropBag.WriteProperty "BackColor", Grid.BackColor, vbWindowBackground
    PropBag.WriteProperty "BackColorFixed", Grid.BackColorFixed, vbButtonFace
    PropBag.WriteProperty "ForeColor", Grid.ForeColor, vbButtonText
    PropBag.WriteProperty "ForeColorFixed", Grid.ForeColorFixed, vbButtonText
    PropBag.WriteProperty "EditBackColor", EditBox.BackColor, vbWindowBackground
    PropBag.WriteProperty "EditForeColor", EditBox.ForeColor, vbButtonText
    PropBag.WriteProperty "Rows", Grid.Rows, 2
    PropBag.WriteProperty "Cols", Grid.Cols, 2
    PropBag.WriteProperty "FixedCols", Grid.FixedCols, 1
    PropBag.WriteProperty "FixedRows", Grid.FixedRows, 1
    PropBag.WriteProperty "WordWrap", Grid.WordWrap, False
    PropBag.WriteProperty "FocusRect", Grid.FocusRect, flexFocusLight
    PropBag.WriteProperty "FormatString", Grid.FormatString, ""
    PropBag.WriteProperty "MergeCells", Grid.MergeCells, flexMergeNever
    PropBag.WriteProperty "RowSizingMode", Grid.RowSizingMode, flexRowSizeAll
    PropBag.WriteProperty "TextStyle", Grid.TextStyle, flexTextFlat
    PropBag.WriteProperty "TextStyleFixed", Grid.TextStyleFixed, flexTextFlat
    PropBag.WriteProperty "AllowUserResizing", Grid.AllowUserResizing, flexResizeNone
    PropBag.WriteProperty "RowHeightMin", Grid.RowHeightMin, 350
    PropBag.WriteProperty "GridLines", Grid.GridLines, flexGridFlat
    PropBag.WriteProperty "GridLinesFixed", Grid.GridLinesFixed, flexGridInset
    PropBag.WriteProperty "GridLineWidth", Grid.GridLineWidth, 1
    PropBag.WriteProperty "GridLineWidthFixed", Grid.GridLineWidthFixed, 1
    PropBag.WriteProperty "SelBackColor", Grid.BackColorSel, vbHighlight
    PropBag.WriteProperty "SelForeColor", Grid.ForeColorSel, vbHighlightText
    PropBag.WriteProperty "HighLight", Grid.HighLight, 0
End Sub

Public Property Get MergeCol(ColNum As Long) As Boolean
Attribute MergeCol.VB_MemberFlags = "400"
    If ColNum < 0 Or ColNum >= Grid.Cols Then
        MergeCol = False
    Else
        MergeCol = Grid.MergeCol(ColNum)
    End If
End Property

Public Property Let MergeCol(ColNum As Long, ByVal vNewValue As Boolean)
    If ColNum >= 0 And ColNum < Grid.Cols Then
        Grid.MergeCol(ColNum) = vNewValue
    End If
End Property

Public Property Get MergeRow(RowNum As Long) As Boolean
Attribute MergeRow.VB_MemberFlags = "400"
    If RowNum < 0 Or RowNum >= Grid.Rows Then
        MergeRow = False
    Else
        MergeRow = Grid.MergeRow(RowNum)
    End If
End Property

Public Property Let MergeRow(RowNum As Long, ByVal vNewValue As Boolean)
    If RowNum >= 0 And RowNum < Grid.Rows Then
        Grid.MergeRow(RowNum) = vNewValue
    End If
End Property

Public Sub AddNew(Optional Index As Long = -1)
Attribute AddNew.VB_Description = "向对象中加入新行并进入编辑模式。"
    AddFlag = True
    If Index < Grid.FixedRows Or Index >= Grid.Rows Then
        Grid.AddItem "" & vbTab & ""
        AfterAddRow
        'Edit Grid.Rows - 1, Grid.FixedCols
        Edit Grid.Rows - 1, getFirstCol(Grid.Rows - 1)
    Else
        Grid.AddItem "" & vbTab & "", Index
        AfterAddRow Index
        'Edit Index, Grid.FixedCols
        Edit Index, getFirstCol(Index)
    End If
End Sub

Public Property Get ModifyMode() As flexMode
Attribute ModifyMode.VB_MemberFlags = "400"
    If AddFlag Then
        ModifyMode = flexAdd
    ElseIf EditBox.Visible Then
        ModifyMode = flexEdit
    Else
        ModifyMode = flexNone
    End If
End Property

Public Property Get SelBackColor() As OLE_COLOR
Attribute SelBackColor.VB_Description = "返回或设置表格中被选定单元的背景颜色。"
Attribute SelBackColor.VB_ProcData.VB_Invoke_Property = ";外观"
    SelBackColor = Grid.BackColorSel
End Property

Public Property Let SelBackColor(ByVal vNewValue As OLE_COLOR)
    Grid.BackColorSel = vNewValue
End Property

Public Property Get SelForeColor() As OLE_COLOR
Attribute SelForeColor.VB_Description = "返回或设置表格中被选定单元的文字颜色。"
Attribute SelForeColor.VB_ProcData.VB_Invoke_Property = ";外观"
    SelForeColor = Grid.ForeColorSel
End Property

Public Property Let SelForeColor(ByVal vNewValue As OLE_COLOR)
    Grid.ForeColorSel = vNewValue
End Property

Public Function ValidValue() As Boolean
    Dim iNewRow As Long, iNewCol As Long, OldValue As String
    ValidValue = False
    On Error Resume Next
    If EditBox.Visible Then
        iNewRow = Grid.Row
        iNewCol = Grid.Col
        Grid.Row = iRow
        Grid.Col = iCol
        OldValue = Grid.Text
        DoEvents
        RaiseEvent BeforeColUpdate(iRow, iCol, EditBox.Text, OldValue, ValidValue)
        If ValidValue Then
            EditBox.SetFocus
            Exit Function
        End If
        Grid.Text = EditBox.Text
        DoEvents
        If Not AddFlag Then
            RaiseEvent BeforeRowUpdate(iRow, "Edit", ValidValue)
        Else
            RaiseEvent BeforeRowUpdate(iRow, "Add", ValidValue)
        End If
        If ValidValue Then
            EditBox.SetFocus
            Exit Function
        End If
        Grid.Row = iNewRow
        Grid.Col = iNewCol
        AddFlag = False
        EndEdit
    End If
End Function

Public Property Get SelCmdFlag(ColNum As Long) As Boolean
Attribute SelCmdFlag.VB_MemberFlags = "400"
    If ColNum > UBound(SelFlag) Then
        SelCmdFlag = False
    Else
        SelCmdFlag = SelFlag(ColNum)
    End If
End Property

Public Property Let SelCmdFlag(ColNum As Long, ByVal vNewValue As Boolean)
    If ColNum > UBound(SelFlag) Then
        ReDim Preserve SelFlag(UBound(SelFlag) + (ColNum - UBound(SelFlag)))
    End If
    SelFlag(ColNum) = vNewValue
End Property

Public Property Get Enabled() As Boolean
    Enabled = Grid.Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    Grid.Enabled = vNewValue
End Property

Public Property Get CellAlignment() As flexAlign
Attribute CellAlignment.VB_MemberFlags = "400"
    CellAlignment = Grid.CellAlignment
End Property

Public Property Let CellAlignment(ByVal vNewValue As flexAlign)
    Grid.CellAlignment = vNewValue
End Property

Private Sub AfterAddRow(Optional ByVal Index As Long = -1)
    Dim i As Long, j As Long, iNum As Long, ItemID As Long
    On Error Resume Next
    
    ReDim Preserve Cells(Grid.Cols - 1, Grid.Rows - 1)
    
    If Index < 0 Or Index > Grid.Rows - 1 Then Index = Grid.Rows - 1
    For i = Grid.Rows - 1 To Index + 1 Step -1
        For j = 0 To Grid.Cols - 1
            Cells(j, i) = Cells(j, i - 1)
        Next j
    Next i
    For j = 0 To Grid.Cols - 1
        Cells(j, Index) = Columns(j)
            
        iNum = -1
        iNum = UBound(ColumnItems(j).List)
        For ItemID = 0 To iNum
            List_AddItem Index, j, ColumnItems(j).List(ItemID)
        Next ItemID
    Next j
End Sub

Private Sub AfterDeleteRow(Optional ByVal Index As Long = -1)
    Dim i As Long, j As Long
    
    If Index < 0 Or Index > Grid.Rows Then Index = Grid.Rows
    For i = Index To Grid.Rows - 1
        For j = 0 To Grid.Cols - 1
            Cells(j, i) = Cells(j, i + 1)
        Next j
    Next i
    
    ReDim Preserve Cells(Grid.Cols - 1, Grid.Rows - 1)
End Sub

Private Function getNextCol(Optional ByVal Col As Long = -1, Optional ByVal Row As Long = -1)
    Dim i As Long
    If Col = -1 Then Col = iCol
    If Row = -1 Then Row = iRow
    
    For i = Col + 1 To Grid.Cols - 1
        If Not Cells(i, Row).Disabled Then Exit For
    Next
    
    getNextCol = i
End Function

Private Function getFirstCol(Optional ByVal Row As Long = -1)
    Dim i As Long
    If Row = -1 Then Row = iRow
    
    For i = 0 To Grid.Cols - 1
        If Not Cells(i, Row).Disabled And i > Grid.FixedCols - 1 Then Exit For
    Next
    
    getFirstCol = i
End Function

Private Function getNearestCol(ByVal Col As Long, Optional ByVal Row As Long = -1)
    Dim i As Long
    If Row = -1 Then Row = iRow
    If Col = -1 Then Col = iCol
    
    For i = Col To Grid.Cols - 1
        If Not Cells(i, Row).Disabled Then Exit For
    Next
    If i > Grid.Cols - 1 Then
        For i = Col - 1 To 0 Step -1
            If Not Cells(i, Row).Disabled Then Exit For
        Next
    End If
    
    getNearestCol = i
End Function

Public Property Get ColDisabled(ByVal Index As Long) As Boolean
    ColDisabled = Columns(Index).Disabled
End Property

Public Property Let ColDisabled(ByVal Index As Long, ByVal vNewValue As Boolean)
    Dim i As Long
    Columns(Index).Disabled = vNewValue
    
    For i = 0 To Grid.Rows - 1
        Cells(Index, i).Disabled = vNewValue
    Next
End Property

Public Property Get ColType(ByVal Index As Long) As EditType
    ColType = Columns(Index).EditMode
End Property

Public Property Let ColType(ByVal Index As Long, ByVal vNewValue As EditType)
    Dim i As Long
    Columns(Index).EditMode = vNewValue
    
    For i = 0 To Grid.Rows - 1
        Cells(Index, i).EditMode = vNewValue
    Next
End Property

Public Property Get CellDisabled(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellDisabled = Cells(Col, Row).Disabled
End Property

Public Property Let CellDisabled(ByVal Row As Long, ByVal Col As Long, ByVal vNewValue As Boolean)
    Cells(Col, Row).Disabled = vNewValue
End Property

Public Property Get CellType(ByVal Row As Long, ByVal Col As Long) As EditType
    CellType = Cells(Col, Row).EditMode
End Property

Public Property Let CellType(ByVal Row As Long, ByVal Col As Long, ByVal vNewValue As EditType)
    Cells(Col, Row).EditMode = vNewValue
End Property

Public Property Get AllowAddNew() As Boolean
    AllowAddNew = bAllowAddNew
End Property

Public Property Let AllowAddNew(ByVal vNewValue As Boolean)
    bAllowAddNew = vNewValue
End Property

Public Property Get List(ByVal Row As Integer, ByVal Col As Integer, ByVal Index As Integer) As String
    On Error Resume Next
    List = ""
    List = CellItems(Cells(Col, Row).ItemIndex).List(Index)
End Property

Public Property Let List(ByVal Row As Integer, ByVal Col As Integer, ByVal Index As Integer, ByVal vNewValue As String)
    On Error Resume Next
    CellItems(Cells(Col, Row).ItemIndex).List(Index) = vNewValue
    
    If Row = iRow And Col = iCol Then Combo1.List(Index) = vNewValue
    If Index = Cells(Col, Row).ListIndex Then Grid.TextMatrix(Row, Col) = vNewValue
End Property

Public Property Get List_Text(ByVal Row As Integer, ByVal Col As Integer) As String
    On Error Resume Next
    List_Text = ""
    List_Text = CellItems(Cells(Col, Row).ItemIndex).List(Cells(Col, Row).ListIndex)
End Property

Public Property Let List_Text(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As String)
    Set_List_Text Row, Col, vNewValue
End Property

Private Sub Set_List_Text(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As String)
    Dim i As Long, iNum As Long
    
    On Error Resume Next
    iNum = -1
    iNum = UBound(CellItems(Cells(Col, Row).ItemIndex).List)
    For i = 0 To iNum
        If CellItems(Cells(Col, Row).ItemIndex).List(i) = Trim(vNewValue) Then
            Cells(Col, Row).ListIndex = i
            Exit For
        End If
    Next
    If Row = iRow And Col = iCol Then Combo1.Text = vNewValue
    Grid.TextMatrix(Row, Col) = vNewValue
End Sub
Public Property Get ListIndex(ByVal Row As Integer, ByVal Col As Integer) As Integer
    On Error Resume Next
    ListIndex = -1
    ListIndex = Cells(Col, Row).ListIndex
End Property

Public Property Let ListIndex(ByVal Row As Integer, ByVal Col As Integer, ByVal vNewValue As Integer)
    On Error Resume Next
    Cells(Col, Row).ListIndex = vNewValue
    If Row = iRow And Col = iCol Then Combo1.ListIndex = vNewValue
    Grid.TextMatrix(Row, Col) = CellItems(Cells(Col, Row).ItemIndex).List(vNewValue)
End Property

Public Sub List_AddItem(ByVal Row As Integer, ByVal Col As Integer, ByVal sItem As String)
    Dim ItemID As Long
    Dim iNum As Long
    
    On Error Resume Next
    ItemID = Cells(Col, Row).ItemIndex
    If ItemID = 0 Then
        ItemID = UBound(CellItems) + 1
        ReDim Preserve CellItems(ItemID)
        Cells(Col, Row).ItemIndex = ItemID
    End If
    
    iNum = -1
    iNum = UBound(CellItems(ItemID).List)
    ReDim Preserve CellItems(ItemID).List(iNum + 1)
    CellItems(ItemID).List(iNum + 1) = sItem
    If Row = iRow And Col = iCol Then Combo1.AddItem sItem
End Sub

Public Sub List_Clear(ByVal Row As Integer, ByVal Col As Integer)
    Dim ItemID As Long
    
    On Error Resume Next
    ItemID = Cells(Col, Row).ItemIndex
'    If ItemID = 0 Then
'        ItemID = UBound(CellItems) + 1
'        ReDim Preserve CellItems(ItemID)
'        Cells(Col, Row).ItemIndex = ItemID
'    End If
    
'    ReDim CellItems(ItemID).List(0)
    If ItemID > 0 Then Erase CellItems(ItemID).List
    Cells(Col, Row).ItemIndex = 0
    
    If Row = iRow And Col = iCol Then Combo1.Clear
End Sub

'Public Sub List_RemoveItem(ByVal Row As Integer, ByVal Col As Integer, ByVal Index As Integer)
'End Sub
Public Property Get ColList(ByVal Col As Integer, ByVal Index As Integer) As String
    On Error Resume Next
    ColList = ""
    ColList = ColumnItems(Col).List(Index)
End Property

Public Property Let ColList(ByVal Col As Integer, ByVal Index As Integer, ByVal vNewValue As String)
    On Error Resume Next
    ColumnItems(Col).List(Index) = vNewValue
End Property

Public Sub ColList_AddItem(ByVal Col As Integer, ByVal sItem As String, Optional ByVal UpdateExistRows As Boolean = False)
    Dim ItemID As Long
    Dim i As Long
    Dim iNum As Long
    
    On Error Resume Next
    
    iNum = -1
    iNum = UBound(ColumnItems(Col).List)
    ReDim Preserve ColumnItems(Col).List(iNum + 1)
    ColumnItems(Col).List(iNum + 1) = sItem

    If UpdateExistRows Then
        For i = 0 To Grid.Rows - 1
            List_AddItem i, Col, sItem
        Next
    End If
End Sub

Public Sub ColList_Clear(ByVal Col As Integer)
    Dim ItemID As Long
    
    On Error Resume Next
    
    Erase ColumnItems(Col).List
End Sub
