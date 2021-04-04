VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.UserControl Table 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin zlSubclass.Subclass Subclass1 
      Left            =   3915
      Top             =   630
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSComctlLib.ImageList imlCursor 
      Left            =   2250
      Top             =   2835
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Table.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Table.ctx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Table.ctx":19B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Table.ctx":268E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   4005
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2250
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   4005
      ScaleHeight     =   435
      ScaleWidth      =   570
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3015
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'################################################################################################################
'## 枚举常量
'################################################################################################################

'字体质量
Public Enum FontQualityEnum
   FQDefault = DEFAULT_QUALITY              '默认质量
   FQDraft = DRAFT_QUALITY                  '草稿质量
   FQProof = PROOF_QUALITY                  '质量优先
   FQNoAntialiased = NONANTIALIASED_QUALITY '取消反锯齿
   FQAntialiased = ANTIALIASED_QUALITY      '反锯齿
   FQClearType = CLEARTYPE_QUALITY          '清晰
End Enum

'选中单元格高亮显示效果
Public Enum HighlightModeEnum
    HMNone = 0
    HMOnlyBorderRectSolid = 1
    HMOnlyBorderRectAlpha = 2
    HMFilledRectSolid = 3
    HMFilledRectAlpha = 4
End Enum

'################################################################################################################
'## 属性变量
'################################################################################################################
Private mvarCells As cCells                     '单元格对象数组
Private RowColInfo() As Long                    '二维动态数组，存储行列与单元格的对应关系

Private mvarRedraw As Boolean                   '是否重绘，默认为True。在数据加载时取消重绘可以避免闪烁
Private mvarSingleLine As Boolean               '是否单行显示，默认为False
Private mvarEnabled As Boolean                  '是否可用
Private mvarRowCount As Long                    '行总数
Private mvarColCount As Long                    '列总数
Private mvarDefaultRowHeight As Long            '默认行高
Private mvarAlternateRowBackColor As OLE_COLOR  '交替色（隔行显示不同颜色）
Private mvarBackColor As OLE_COLOR              '背景色
Private mvarBackgroundPicture As StdPicture     '背景图片
Private mvarGridLineColor As OLE_COLOR          '网格线颜色，默认为黑色
Private mvarGridLineWidth As Long               '网格线宽度，默认为1
Private mvarBorderColor As OLE_COLOR            '边框颜色，默认为黑色
Private mvarBorderWidth As Long                 '边框宽度，默认为0
Private mvarEditable As Boolean                 '是否可以编辑
Private mvarForeColor As OLE_COLOR              '前景色
Private mvarHighlightBackColor As OLE_COLOR     '高亮背景色
Private mvarHighlightForeColor As OLE_COLOR     '高亮前景色
Private mvarHighlightSelectedIcons As Boolean   '是否高亮显示图标（虚线）
Private mvarHighlightMode As HighlightModeEnum  '高亮显示模式
Private mvarDrawFocusRect As Boolean            '是否显示高亮虚框
Private mvarHotTrack As Boolean                 '是否热跟踪
Private mvarSingleClickEdit As Boolean          '单击启动编辑
Private mvarFontQuality As FontQualityEnum      '字体质量
Private mvarAutoHeight As Boolean               '自动高度，默认为True
Private mvarMinRowHeight As Long                '最小行高，默认为0
Private mvarWordEllipsis As Boolean             '是否在文本无法显示完时显示一个省略号
Private mvarhImageList As Long                  '绑定 VB6 ImageList的对象指针
Private mvarCellMargin As Long                  '单元格边距，默认为30
Private mvarCellIndent As Long                  '单元格缩进
Private mvarInnerEdit As Boolean                '启用内部编辑
Private mvarTabKeyMoveNextCell As Boolean       '是否Tab键移动到下一个单元格，否则失去焦点
Private mvarShowToolTipText As Boolean          '是否显示工具提示文本
Private mvarModified As Boolean                 '是否编辑过
Private mvarExtendTag As String                 '扩展Tag属性，用于记录附加信息
Private mvarUserTag As String                   '用户Tag属性，用于用户指定的表格标识

'################################################################################################################
'## 局部变量
'################################################################################################################
Private WithEvents m_tmrHotTrack As cTimer      '热跟踪计时器
Attribute m_tmrHotTrack.VB_VarHelpID = -1

Private m_bDirty As Boolean                     '是否需要重绘整个控件
Private m_DefaultWidth As Long                  '默认宽度，等于控件宽度除以列数
Private m_DefaultHeight As Long                 '默认高度，等于控件宽度除以行数

'选中单元格
Private m_CellKeySelected As Long
Private m_CellKeyHot As Long
Private m_CellKeyEdit As Long
Private m_SelStartRow As Long, m_SelStartCol As Long, m_SelEndRow As Long, m_SelEndCol As Long, m_bMouseDown As Boolean

'内存DC，用于避免闪烁（单行）－同时实现裁剪
Private m_hDC As Long                           '与控件对应的内存设备场景
Private m_hBmp As Long                          '位图句柄
Private m_hBmpOld As Long                       '旧的位图句柄

'调整列宽时用
Private m_bAdjustColWidth As Boolean            '在调整列宽过程中
Private m_bAdjustRowHeight As Boolean           '在调整行高过程中
Private m_ColAdjust As Long                     '在调整列
Private m_RowAdjust As Long                     '在调整行
Private m_OldX As Long                          '
Private m_OldY As Long                          '
Private m_hWndBound As Long                     '绑定的hWnd
Private m_OffsetX As Long                       'X轴偏移量
Private m_OffsetY As Long                       'Y轴偏移量

'背景:
Private m_bBitmap As Boolean
Private m_hDCSrc As Long
Private m_lBitmapW As Long
Private m_lBitmapH As Long
Private m_bTrueColor As Boolean

'编辑标志
Private m_bInEdit As Boolean                        '正在编辑中
Private m_bInEndEditInterlock As Boolean            '
Private m_iRepostMsg As Long
Private m_tRepostPos As POINTAPI
Private m_lRepostShiftState As Long
Private m_bInResize As Boolean                      '手工调整控件尺寸

Private m_hWnd As Long
Private m_hWndParentForm As Long
Private m_bRunningInVBIDE As Boolean                '是否在VB IDE中运行
Private m_tR As RECT
Private m_bInFocus As Boolean                       '是否获得焦点

Private m_bEditHeightChanged As Boolean             '编辑过程中，编辑控件的高度是否已经随内容改变，如果是才重新计算单元格坐标

Private m_IPAOHookStruct As IPAOHookStruct

'################################################################################################################
'## 事件声明
'################################################################################################################

Public Event ColumnWidthStartChange(ByVal lCol As Long, ByRef lNewWidth As Long, ByRef bCancel As Boolean)
Public Event ColumnWidthChanging(ByVal lCol As Long, ByRef lNewWidth As Long, ByRef bCancel As Boolean)
Public Event ColumnWidthChanged(ByVal lCol As Long, ByRef lNewWidth As Long, ByRef bCancel As Boolean)
Public Event ColumnDividerDblClick(ByVal lCol As Long, ByRef bCancel As Boolean)
Public Event SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
Public Event HotItemChange(ByVal lRow As Long, ByVal lCol As Long)
Public Event RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, ByRef bCancel As Boolean)
Public Event PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, ByRef newValue As Variant, ByRef bStayInEditMode As Boolean)
Public Event CancelEdit()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DblClick(ByVal lRow As Long, ByVal lCol As Long)
Public Event Resize(ByVal lWidth As Long, ByVal lHeight As Long)
Public Event ModifyProtected(ByVal lKey As Long)

'################################################################################################################
'## 全局属性
'################################################################################################################
Public Property Get SelStartRow() As Long
    SelStartRow = m_SelStartRow
End Property

Public Property Get SelStartCol() As Long
    SelStartCol = m_SelStartCol
End Property

Public Property Get SelEndRow() As Long
    SelEndRow = m_SelEndRow
End Property

Public Property Get SelEndCol() As Long
    SelEndCol = m_SelEndCol
End Property

Public Property Get Row() As Long
    If m_CellKeySelected > 0 Then Row = mvarCells("K" & m_CellKeySelected).Row
End Property

Public Property Get Col() As Long
    If m_CellKeySelected > 0 Then Col = mvarCells("K" & m_CellKeySelected).Col
End Property

Public Property Get Cell(ByVal lRow As Long, ByVal lCol As Long) As cCell
    If ValidCell(lRow, lCol) Then
        Set Cell = mvarCells("K" & RowColInfo(lRow, lCol))
    Else
        Set Cell = Nothing
    End If
End Property

Public Property Set Cell(ByVal lRow As Long, ByVal lCol As Long, ByRef vData As cCell)
    If ValidCell(lRow, lCol) Then
        With mvarCells("K" & RowColInfo(lRow, lCol))
            .Key = vData.Key
            .Row = vData.Row
            .Col = vData.Col
            .Margin = vData.Margin
            .SingleLine = vData.SingleLine
            .MergeInfo = vData.MergeInfo
            .Selected = vData.Selected
            .Hot = vData.Hot
            .Visibled = vData.Visibled
            .Width = vData.Width
            .Height = vData.Height
            .FixedWidth = vData.FixedWidth
            .AutoHeight = vData.AutoHeight
            .Icon = vData.Icon
            .Text = vData.Text
            .Tag = vData.Tag
            .FormatString = vData.FormatString
            .Indent = vData.Indent
            .HAlignment = vData.HAlignment
            .VAlignment = vData.VAlignment
            .ForeColor = vData.ForeColor
            .BackColor = vData.BackColor
            .GridLineColor = vData.GridLineColor
            .GridLineWidth = vData.GridLineWidth
            .FontName = vData.FontName
            .FontSize = vData.FontSize
            .FontBold = vData.FontBold
            .FontItalic = vData.FontItalic
            .FontStrikeout = vData.FontStrikeout
            .FontUnderline = vData.FontUnderline
            .FontWeight = vData.FontWeight
            .Protected = vData.Protected
            .Dirty = vData.Dirty
            .ToolTipText = vData.ToolTipText
            Set .Picture = vData.Picture
        End With
    End If
End Property

Public Property Get CellKey(ByVal lRow As Long, ByVal lCol As Long) As Long
    If ValidCell(lRow, lCol) Then CellKey = RowColInfo(lRow, lCol)
End Property

Public Property Let Cells(ByVal vData As cCells)
    Set mvarCells = vData
End Property

Public Property Get Cells() As cCells
    Set Cells = mvarCells
End Property

Public Property Let Redraw(ByVal vData As Boolean)
    mvarRedraw = vData
    PropertyChanged "Redraw"
End Property

Public Property Get Redraw() As Boolean
    Redraw = mvarRedraw
End Property

Public Property Let SingleLine(ByVal vData As Boolean)
    mvarSingleLine = vData
    If vData = False Then WordEllipsis = False
    If UserControl.Ambient.UserMode And p_TPPX > 0 And p_TPPY > 0 Then
        Dim i As Long
        For i = 1 To mvarCells.Count
            mvarCells(i).SingleLine = vData
        Next
        Refresh False, True      '不计算宽度，但是要重新计算行高
    End If
    PropertyChanged "SingleLine"
End Property

Public Property Get SingleLine() As Boolean
    SingleLine = mvarSingleLine
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
    m_bDirty = True
    Draw
    PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

Public Property Let RowCount(ByVal vData As Long)
    mvarRowCount = vData
    PropertyChanged "RowCount"
End Property

Public Property Get RowCount() As Long
    RowCount = mvarRowCount
End Property

Public Property Let ColCount(ByVal vData As Long)
    mvarColCount = vData
    PropertyChanged "ColCount"
End Property

Public Property Get ColCount() As Long
    ColCount = mvarColCount
End Property

Public Property Let DefaultRowHeight(ByVal vData As Long)
    mvarDefaultRowHeight = vData
    PropertyChanged "DefaultRowHeight"
End Property

Public Property Get DefaultRowHeight() As Long
    DefaultRowHeight = mvarDefaultRowHeight
End Property

Public Property Let AlternateRowBackColor(ByVal vData As OLE_COLOR)
    mvarAlternateRowBackColor = vData
    PropertyChanged "AlternateRowBackColor"
End Property

Public Property Get AlternateRowBackColor() As OLE_COLOR
    AlternateRowBackColor = mvarAlternateRowBackColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
    UserControl.BackColor = vData
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Set BackgroundPicture(ByVal vData As StdPicture)
    On Error Resume Next
   
    Set picImage.Picture = vData
    Set mvarBackgroundPicture = vData
    picImage.Refresh
    If (Err.Number <> 0) Or (picImage.ScaleWidth = 0) Or (vData Is Nothing) Then
        m_hDCSrc = 0
        m_bBitmap = False
    Else
        '设置成功
        m_bBitmap = True
        m_hDCSrc = picImage.hDC
        m_lBitmapW = picImage.ScaleWidth \ Screen.TwipsPerPixelX
        m_lBitmapH = picImage.ScaleHeight \ Screen.TwipsPerPixelY
    End If
    '重绘
    m_bDirty = True
    Draw
    PropertyChanged "BackgroundPicture"
End Property

Public Property Let BackgroundPicture(ByVal vData As StdPicture)
    On Error Resume Next
   
    Set picImage.Picture = vData
    Set mvarBackgroundPicture = vData
    picImage.Refresh
    If (Err.Number <> 0) Or (picImage.ScaleWidth = 0) Or (vData Is Nothing) Then
        m_hDCSrc = 0
        m_bBitmap = False
    Else
        '设置成功
        m_bBitmap = True
        m_hDCSrc = picImage.hDC
        m_lBitmapW = picImage.ScaleWidth \ Screen.TwipsPerPixelX
        m_lBitmapH = picImage.ScaleHeight \ Screen.TwipsPerPixelY
    End If
    '重绘
    m_bDirty = True
    Draw
    PropertyChanged "BackgroundPicture"
End Property

Public Property Get BackgroundPicture() As StdPicture
    Set BackgroundPicture = mvarBackgroundPicture
End Property

Public Property Let GridLineColor(ByVal vData As OLE_COLOR)
    mvarGridLineColor = vData
    Dim i As Long
    For i = 1 To mvarCells.Count
        mvarCells(i).GridLineColor = vData
    Next
    Draw
    PropertyChanged "GridLineColor"
End Property

Public Property Get GridLineColor() As OLE_COLOR
    GridLineColor = mvarGridLineColor
End Property

Public Property Let GridLineWidth(ByVal vData As Long)
    mvarGridLineWidth = vData
    Dim i As Long
    For i = 1 To mvarCells.Count
        mvarCells(i).GridLineWidth = vData
    Next
    Draw
    PropertyChanged "GridLineWidth"
End Property

Public Property Get GridLineWidth() As Long
    GridLineWidth = mvarGridLineWidth
End Property

Public Property Let BorderColor(ByVal vData As OLE_COLOR)
    mvarBorderColor = vData
    PropertyChanged "BorderColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mvarBorderColor
End Property

Public Property Let BorderWidth(ByVal vData As Long)
    mvarBorderWidth = vData
    PropertyChanged "BorderWidth"
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = mvarBorderWidth
End Property

Public Property Let Editable(ByVal vData As Boolean)
    mvarEditable = vData
    PropertyChanged "Editable"
End Property

Public Property Get Editable() As Boolean
    Editable = mvarEditable
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let HighlightBackColor(ByVal vData As OLE_COLOR)
    mvarHighlightBackColor = vData
    PropertyChanged "HighlightBackColor"
End Property

Public Property Get HighlightBackColor() As OLE_COLOR
    HighlightBackColor = mvarHighlightBackColor
End Property

Public Property Let HighlightForeColor(ByVal vData As OLE_COLOR)
    mvarHighlightForeColor = vData
    PropertyChanged "HighlightForeColor"
End Property

Public Property Get HighlightForeColor() As OLE_COLOR
    HighlightForeColor = mvarHighlightForeColor
End Property

Public Property Let HighlightSelectedIcons(ByVal vData As Boolean)
    mvarHighlightSelectedIcons = vData
    PropertyChanged "HighlightSelectedIcons"
End Property

Public Property Get HighlightSelectedIcons() As Boolean
    HighlightSelectedIcons = mvarHighlightSelectedIcons
End Property

Public Property Let HighlightMode(ByVal vData As HighlightModeEnum)
    mvarHighlightMode = vData
    PropertyChanged "HighlightMode"
End Property

Public Property Get HighlightMode() As HighlightModeEnum
    HighlightMode = mvarHighlightMode
End Property

Public Property Let DrawFocusRect(ByVal vData As Boolean)
    mvarDrawFocusRect = vData
    PropertyChanged "DrawFocusRect"
End Property

Public Property Get DrawFocusRect() As Boolean
    DrawFocusRect = mvarDrawFocusRect
End Property

Public Property Let HotTrack(ByVal vData As Boolean)
    mvarHotTrack = vData
    PropertyChanged "HotTrack"
End Property

Public Property Get HotTrack() As Boolean
    HotTrack = mvarHotTrack
End Property

Public Property Let SingleClickEdit(ByVal vData As Boolean)
    mvarSingleClickEdit = vData
    PropertyChanged "SingleClickEdit"
End Property

Public Property Get SingleClickEdit() As Boolean
    SingleClickEdit = mvarSingleClickEdit
End Property

Public Property Let FontQuality(ByVal vData As FontQualityEnum)
    mvarFontQuality = vData
    PropertyChanged "FontQuality"
End Property

Public Property Get FontQuality() As FontQualityEnum
    FontQuality = mvarFontQuality
End Property

Public Property Let AutoHeight(ByVal vData As Boolean)
    mvarAutoHeight = vData
    PropertyChanged "AutoHeight"
End Property

Public Property Get AutoHeight() As Boolean
    AutoHeight = mvarAutoHeight
End Property

Public Property Let MinRowHeight(ByVal vData As Long)
    mvarMinRowHeight = vData
    PropertyChanged "MinRowHeight"
End Property

Public Property Get MinRowHeight() As Long
    MinRowHeight = mvarMinRowHeight
End Property

Public Property Let WordEllipsis(ByVal vData As Boolean)
    mvarWordEllipsis = vData
    m_bDirty = True
    Draw
    PropertyChanged "MinRowHeight"
End Property

Public Property Get WordEllipsis() As Boolean
    WordEllipsis = mvarWordEllipsis
End Property

Public Property Let ColWidth(ByVal Col As Long, ByVal vData As Long)
    ColInfo(Col).ColWidth = Abs(vData)
    ColInfo(Col).FixedWidth = (vData < 0)
    ColInfo(Col).Visible = (vData <> 0)
End Property

Public Property Get ColWidth(ByVal Col As Long) As Long
    ColWidth = IIf(ColInfo(Col).FixedWidth, -ColInfo(Col).ColWidth, ColInfo(Col).ColWidth)
End Property

Public Property Let ColLeftX(ByVal Col As Long, ByVal vData As Long)
    ColInfo(Col).LeftX = vData
End Property

Public Property Get ColLeftX(ByVal Col As Long) As Long
    ColLeftX = ColInfo(Col).LeftX
End Property

Public Property Let ColFixedWidth(ByVal Col As Long, ByVal vData As Boolean)
    ColInfo(Col).FixedWidth = vData
End Property

Public Property Get ColFixedWidth(ByVal Col As Long) As Boolean
    ColFixedWidth = ColInfo(Col).FixedWidth
End Property

Public Property Let ColVisible(ByVal Col As Long, ByVal vData As Boolean)
    ColInfo(Col).Visible = vData
End Property

Public Property Get ColVisible(ByVal Col As Long) As Boolean
    ColVisible = ColInfo(Col).Visible
End Property

Public Property Let RowHeight(ByVal Row As Long, ByVal vData As Long)
    Dim i As Long
    RowInfo(Row).RowHeight = Abs(vData)
    RowInfo(Row).FixedHeight = (vData < 0)

    If (Row > 0) Then
        For i = 1 To mvarCells.Count
            If mvarCells(i).Row = Row Then mvarCells(i).Height = vData
        Next
    End If
End Property

Public Property Get RowHeight(ByVal Row As Long) As Long
    RowHeight = RowInfo(Row).RowHeight
End Property

Public Property Let RowTopY(ByVal Row As Long, ByVal vData As Long)
    RowInfo(Row).TopY = vData
End Property

Public Property Get RowTopY(ByVal Row As Long) As Long
    RowTopY = RowInfo(Row).TopY
End Property

Public Property Set ImageList(ByVal vData As Object)
    mvarhImageList = 0
    On Error Resume Next
    '必须先初始化该控件
    vData.ListImages(1).Draw 0, 0, 0, 1
    If (Err.Number <> 0) Then
    Else
        mvarhImageList = vData.hImagelist
        p_lIconWidth = vData.ImageWidth * p_TPPX
        p_lIconHeight = vData.ImageHeight * p_TPPY
    End If
    Err = 0: On Error GoTo 0
    PropertyChanged "ImageList"
End Property

Public Property Let ImageList(ByVal vData As Object)
    mvarhImageList = 0
    On Error Resume Next
    '必须先初始化该控件
    vData.ListImages(1).Draw 0, 0, 0, 1
    If (Err.Number <> 0) Then
    Else
        mvarhImageList = vData.hImagelist
        p_lIconWidth = vData.ImageWidth * p_TPPX
        p_lIconHeight = vData.ImageHeight * p_TPPY
    End If
    Err = 0: On Error GoTo 0
    PropertyChanged "ImageList"
End Property

Friend Property Get PtrImageList() As Long
    PtrImageList = mvarhImageList
End Property

Public Property Let CellMargin(ByVal vData As Long)
    mvarCellMargin = vData
End Property

Public Property Get CellMargin() As Long
    CellMargin = mvarCellMargin
End Property

Public Property Let CellIndent(ByVal vData As Long)
    mvarCellIndent = vData
End Property

Public Property Get CellIndent() As Long
    CellIndent = mvarCellIndent
End Property

Public Property Let InnerEdit(ByVal vData As Boolean)
    mvarInnerEdit = vData
End Property

Public Property Get InnerEdit() As Boolean
    InnerEdit = mvarInnerEdit
End Property

Public Property Let ShowToolTipText(ByVal vData As Boolean)
    mvarShowToolTipText = vData
    m_bDirty = True
    Draw
End Property

Public Property Get ShowToolTipText() As Boolean
    ShowToolTipText = mvarShowToolTipText
End Property

Public Property Let Modified(ByVal vData As Boolean)
    mvarModified = vData
End Property

Public Property Get Modified() As Boolean
    Modified = mvarModified
End Property

Public Property Let ExtendTag(ByVal vData As String)
    mvarExtendTag = vData
End Property

Public Property Get ExtendTag() As String
    ExtendTag = mvarExtendTag
End Property

Public Property Let UserTag(ByVal vData As String)
    mvarUserTag = vData
End Property

Public Property Get UserTag() As String
    UserTag = mvarUserTag
End Property

Public Property Let TabKeyMoveNextCell(ByVal vData As Boolean)
    mvarTabKeyMoveNextCell = vData
End Property

Public Property Get TabKeyMoveNextCell() As Boolean
    TabKeyMoveNextCell = mvarTabKeyMoveNextCell
End Property

Public Property Set Font(ByVal vData As StdFont)
    Set UserControl.Font = vData
    PropertyChanged "Font"
End Property

Public Property Let Font(ByVal vData As StdFont)
    Set UserControl.Font = vData
    PropertyChanged "Font"
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Friend Property Get InFocus() As Boolean
    InFocus = m_bInFocus
End Property

Public Property Get InEdit() As Boolean
    InEdit = m_bInEdit
End Property

Public Property Let SelectedCellKey(vData As Long)
    m_CellKeySelected = vData
End Property

Public Property Get SelectedCellKey() As Long
    SelectedCellKey = m_CellKeySelected
End Property

Public Property Get HotedCellKey() As Long
    HotedCellKey = m_CellKeyHot
End Property

Public Property Let hWndBound(vData As Long)
    m_hWndBound = vData
End Property

Friend Property Get hWndBound() As Long
    hWndBound = m_hWndBound
End Property

Public Property Let OffsetX(vData As Long)
    m_OffsetX = vData
End Property

Friend Property Get OffsetX() As Long
    OffsetX = m_OffsetX
End Property

Public Property Let OffsetY(vData As Long)
    m_OffsetY = vData
End Property

Friend Property Get OffsetY() As Long
    OffsetY = m_OffsetY
End Property

'################################################################################################################
'## 全局方法
'################################################################################################################
Public Sub InsertText(ByVal strText As String)
    '如果在编辑过程中，那么在当前位置插入文本
    Dim i As Long, j As Long, S As String
    If txtEdit.Visible Then
        i = txtEdit.SelStart
        j = txtEdit.SelStart + txtEdit.SelLength
        S = txtEdit.Text
        txtEdit.Text = Mid(S, 1, i) & strText & Mid(S, j + 1)
        txtEdit.SelStart = Len(Mid(S, 1, i) & strText)
        txtEdit.SelLength = 0
    End If
End Sub

Public Sub PressDelKey()
    '模拟按下删除键的操作
    Dim i As Long, j As Long, s1 As String, s2 As String, S As String
    If txtEdit.Visible Then
        i = txtEdit.SelStart
        j = txtEdit.SelStart + txtEdit.SelLength
        S = txtEdit.Text
        s1 = Mid(S, 1, i)
        If Mid(S, j + 1, 2) = vbCrLf Then
            s2 = Mid(S, IIf(j = i, j + 3, j + 2), Len(S))
        Else
            s2 = Mid(S, IIf(j = i, j + 2, j + 1), Len(S))
        End If
        txtEdit.Text = s1 & s2
        txtEdit.SelStart = Len(s1)
        txtEdit.SelLength = 0
    End If
End Sub

Private Function ValidCell(ByVal Row As Long, ByVal Col As Long) As Boolean
    If Row > 0 And Col > 0 And Row <= mvarRowCount And Col <= mvarColCount Then
        ValidCell = True
    End If
End Function

Public Function CellDetails(ByVal lRow As Long, ByVal lCol As Long, Optional ByVal sText As String = "", _
    Optional ByVal sFormatString As String = "", Optional ByVal sTag As String = "", _
    Optional ByVal lIcon As Long = -1, Optional ByVal sToolTipText As String, _
    Optional ByVal bVisibled As Boolean = True, _
    Optional ByVal sMergeInfo As String = "", _
    Optional ByVal bProtected As Boolean = False, _
    Optional ByVal eHAlignment As HAlignEnum = HALignLeft, Optional ByVal eVAlignment As VAlignEnum = VALignTop, _
    Optional ByVal sFontName As String = "Arial", Optional ByVal lFontSize As Long = 11, _
    Optional ByVal bFontBold As Boolean = False, Optional ByVal bFontItalic As Boolean = False, _
    Optional ByVal bFontStrikeout As Boolean = False, Optional ByVal bFontUnderline As Boolean = False, _
    Optional ByVal lFontWeight As Long, _
    Optional ByVal oForeColor As OLE_COLOR = -1, _
    Optional ByVal oBackColor As OLE_COLOR = -1 _
    ) As Boolean
    
    Dim lDefaultWidth As Long
    
    '设置单元格内容和格式
    If ValidCell(lRow, lCol) Then
        With Cell(lRow, lCol)
            .Dirty = True
            .Text = sText
            .FormatString = sFormatString
            .Tag = sTag
            .Icon = lIcon
            .Visibled = bVisibled
            .Width = m_DefaultWidth
            .AutoHeight = mvarAutoHeight
            .ToolTipText = sToolTipText
            .GridLineColor = mvarGridLineColor
            .GridLineWidth = mvarGridLineWidth
            .Indent = mvarCellIndent
            .Margin = mvarCellMargin
            .MergeInfo = sMergeInfo
            .Protected = bProtected
            .SingleLine = mvarSingleLine
            .HAlignment = eHAlignment
            .VAlignment = eVAlignment
            .FontBold = IIf(IsMissing(bFontBold), Font.Bold, bFontBold)
            .FontItalic = IIf(IsMissing(bFontItalic), Font.Italic, bFontItalic)
            .FontName = IIf(IsMissing(sFontName), Font.Name, sFontName)
            .FontSize = IIf(IsMissing(lFontSize), Font.Size, lFontSize)
            .FontStrikeout = IIf(IsMissing(bFontStrikeout), Font.Strikethrough, bFontStrikeout)
            .FontUnderline = IIf(IsMissing(bFontUnderline), Font.Underline, bFontUnderline)
            .FontWeight = IIf(IsMissing(lFontWeight), Font.Weight, lFontWeight)
            .ForeColor = IIf(oForeColor = -1, mvarForeColor, oForeColor)
            .BackColor = oBackColor
        End With
        CellDetails = True
    End If
End Function

Public Sub Refresh(Optional ByVal bReCalculateCellWidth As Boolean = True, _
    Optional ByVal bReCalculateCellHeight As Boolean = True, _
    Optional ByVal lKey As Long = 0)
    
    m_bDirty = True
    If bReCalculateCellWidth Then FixCellsWidth
    If bReCalculateCellHeight Then FixCellsHeight lKey
    BuildMemDC      '只有行高、列宽改变时才重建内存位图
    Draw
End Sub

'################################################################################################################
'## 内部方法
'################################################################################################################

Public Sub FixCellsWidth()
    Dim i As Long, j As Long, m As Long, n As Long, lTmp As Long, lRow As Long, lCol As Long, lW As Long
    Dim lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long, sText As String
    
    m_bInResize = True
    
    '修正列X坐标位置和控件宽度
    For i = 1 To mvarColCount
        If i = 1 Then
            lTmp = mvarBorderWidth * p_TPPX
            ColInfo(i).LeftX = lTmp
        Else
            lTmp = lTmp + ColInfo(i - 1).ColWidth
            ColInfo(i).LeftX = lTmp
        End If
    Next i
    UserControl.Width = ColInfo(mvarColCount).LeftX + ColInfo(mvarColCount).ColWidth + (mvarGridLineWidth + mvarBorderWidth) * p_TPPX
    
    '修正单元格宽度，对合并单元格作特殊处理
    For i = 1 To mvarRowCount
        For j = 1 To mvarColCount
            With mvarCells("K" & RowColInfo(i, j))
                .Width = ColInfo(j).ColWidth
                If Len(.MergeInfo) = 16 And .Visibled Then
                    '合并单元格的左上角单元格
                    sText = .MergeInfo
                    lRow1 = Val(Mid(sText, 1, 4))
                    lCol1 = Val(Mid(sText, 5, 4))
                    lRow2 = Val(Mid(sText, 9, 4))
                    lCol2 = Val(Mid(sText, 13, 4))
                    lW = 0
                    For m = lCol1 To lCol2
                        lW = lW + ColInfo(m).ColWidth   '合并单元格宽度
                    Next
                    .Width = lW
                    For m = lRow1 To lRow2
                        For n = lCol1 To lCol2
                            If m <> lRow1 Or n <> lCol1 Then
                                mvarCells("K" & RowColInfo(m, n)).Visibled = False
                            End If
                        Next
                    Next
                End If
            End With
        Next
    Next
    m_bInResize = False
End Sub
 
Public Sub FixCellsHeight(Optional ByVal lCellKey As Long = 0)
    Dim i As Long, j As Long, k As Long, l As Long, m As Long, lW As Long, lH1 As Long, lH2 As Long, lTmp As Long
    Dim lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long, lKey As Long, sText As String
        
    m_bInResize = True
    
    '重新计算每个单元格的高度
    If mvarAutoHeight Then
        If lCellKey > 0 Then
            '改变指定单元格内容，相关行高重新计算！！！
            sText = mvarCells("K" & lCellKey).MergeInfo
            If Len(sText) = 16 Then
                lRow1 = Val(Mid(sText, 1, 4))
                lRow2 = Val(Mid(sText, 9, 4))
                For i = lRow1 To lRow2
                    AutoHeightRow i, mvarMinRowHeight
                Next
            Else
                AutoHeightRow mvarCells("K" & lCellKey).Row, mvarMinRowHeight
            End If
        Else
            '否则全部重新计算！！！
            For i = 1 To mvarRowCount
                AutoHeightRow i, mvarMinRowHeight
            Next
        End If
    End If
    
    '修正行高和单元格高度
    For i = 1 To mvarRowCount
        For j = 1 To mvarColCount
            With mvarCells("K" & RowColInfo(i, j))
                sText = .MergeInfo
                If Len(sText) = 16 Then
                    lRow1 = Val(Mid(sText, 1, 4))
                    lCol1 = Val(Mid(sText, 5, 4))
                    lRow2 = Val(Mid(sText, 9, 4))
                    lCol2 = Val(Mid(sText, 13, 4))
                    If lRow2 = i Then
                        '如果是合并单元格的最末一行，则判断
                        lH1 = 0
                        For k = lRow1 To lRow2
                            lH1 = lH1 + RowInfo(k).RowHeight
                        Next
                        lTmp = mvarCells("K" & RowColInfo(lRow1, lCol1)).EvaluateTextHeight(Me, , mvarAutoHeight)  '合并单元格的实际文本高度
                        If lTmp > lH1 Then
                            '文本高度大于行高总和，那么调整最末行的行高
                            RowInfo(i).RowHeight = RowInfo(i).RowHeight + (lTmp - lH1)
                            For k = 1 To mvarColCount
                                mvarCells("K" & RowColInfo(i, k)).Height = RowInfo(i).RowHeight
                            Next
                            mvarCells("K" & RowColInfo(lRow1, lCol1)).Height = lTmp
                            j = 0
                        Else
                            mvarCells("K" & RowColInfo(lRow1, lCol1)).Height = lH1
                            j = j + (lCol2 - lCol1) '只处理左下角单元格一次！！！！！！！
                        End If
                    End If
                End If
            End With
        Next
    Next
    
    '计算出所有的单元格高度后，最后才能确定每个行的起始Y坐标位置
    For i = 1 To mvarRowCount
        If i = 1 Then
            lTmp = mvarBorderWidth * p_TPPY
            RowInfo(i).TopY = lTmp
        Else
            lTmp = lTmp + RowInfo(i - 1).RowHeight
            RowInfo(i).TopY = lTmp
        End If
    Next i
    
    '修正控件高度
    On Error Resume Next
    UserControl.Height = RowInfo(mvarRowCount).TopY + RowInfo(mvarRowCount).RowHeight + (mvarGridLineWidth + mvarBorderWidth) * p_TPPY
    
    m_bInResize = False
End Sub

Public Sub DrawToDC(ByRef hDC As Long)
    '绘制到指定DC
    Dim tR As RECT
    Dim hPen As Long
    Dim hPenOld As Long
    Dim tJ As POINTAPI

    Dim i As Long
    If hDC = 0 Then Exit Sub
    
    '获取控件尺寸，准备绘图
    GetClientRect UserControl.hWnd, tR
    pFillBackground hDC, tR, 0, 0, False
    
    '绘制边框
    hPen = CreatePen(PS_SOLID, 1, mvarBorderColor)      '设置边框颜色画笔
    hPenOld = SelectObject(hDC, hPen)                   '选入画笔，保存旧画笔
    tR.Right = UserControl.Width / p_TPPX - 1
    tR.Bottom = UserControl.Height / p_TPPY - 1
    For i = 0 To mvarBorderWidth - 1
        MoveToEx hDC, tR.Left + i, tR.Bottom - i, tJ
        LineTo hDC, tR.Right - i, tR.Bottom - i
        LineTo hDC, tR.Right - i, tR.Top + i
        LineTo hDC, tR.Left + i, tR.Top + i
        LineTo hDC, tR.Left + i, tR.Bottom - i
    Next
    SelectObject hDC, hPenOld
    DeleteObject hPen

    '重绘整个控件
    For i = 1 To mvarCells.Count
        '取消选择
        mvarCells(i).Hot = False
        If mvarCells(i).Selected Then
            mvarCells(i).Selected = False
            mvarCells(i).DrawCell Me, hDC
            mvarCells(i).Selected = True
        Else
            mvarCells(i).DrawCell Me, hDC
        End If
    Next
End Sub

Private Sub Draw()
    If UserControl.Width = 0 Then Exit Sub
    '绘制脏单元格
    Dim tR As RECT
    Dim hPen As Long
    Dim hPenOld As Long
    Dim tJ As POINTAPI

    Dim i As Long
    If m_bDirty Then
        If m_hDC = 0 Then
            '设计模式
            GetClientRect UserControl.hWnd, tR
            pFillBackground UserControl.hDC, tR, 0, 0, False
            m_bDirty = False
            UserControl.Picture = UserControl.Image     '刷新显示！！！！！
            UpdateWindow UserControl.hWnd
            Exit Sub
        End If
        
        '获取控件尺寸，准备绘图
        GetClientRect UserControl.hWnd, tR
        pFillBackground m_hDC, tR, 0, 0, False
        
        '绘制边框
        hPen = CreatePen(PS_SOLID, 1, mvarBorderColor)      '设置边框颜色画笔
        hPenOld = SelectObject(m_hDC, hPen)                 '选入画笔，保存旧画笔
        tR.Right = UserControl.Width / p_TPPX - 1
        tR.Bottom = UserControl.Height / p_TPPY - 1
        For i = 0 To mvarBorderWidth - 1
            MoveToEx m_hDC, tR.Left + i, tR.Bottom - i, tJ
            LineTo m_hDC, tR.Right - i, tR.Bottom - i
            LineTo m_hDC, tR.Right - i, tR.Top + i
            LineTo m_hDC, tR.Left + i, tR.Top + i
            LineTo m_hDC, tR.Left + i, tR.Bottom - i
        Next
        SelectObject m_hDC, hPenOld
        DeleteObject hPen

        '重绘整个控件
        For i = 1 To mvarCells.Count
            If mvarCells(i).Dirty Or m_bDirty Then
                mvarCells(i).DrawCell Me
                mvarCells(i).Dirty = False
            End If
        Next
        '绘制完成后，将图片拷贝到控件上显示
        BitBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, m_hDC, 0, 0, vbSrcCopy
        m_bDirty = False
        UserControl.Picture = UserControl.Image     '刷新显示！！！！！
        UpdateWindow UserControl.hWnd
    End If
End Sub

Public Sub MergeSelectedCells()
    Dim i As Long, j As Long, bCanMerge As Boolean
    Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long
    If m_SelStartRow > m_SelEndRow Then
        lRow1 = m_SelEndRow
        lRow2 = m_SelStartRow
    Else
        lRow1 = m_SelStartRow
        lRow2 = m_SelEndRow
    End If
    If m_SelStartCol > m_SelEndCol Then
        lCol1 = m_SelEndCol
        lCol2 = m_SelStartCol
    Else
        lCol1 = m_SelStartCol
        lCol2 = m_SelEndCol
    End If
    If lRow1 = 0 Or lCol1 = 0 Or lRow2 = 0 Or lCol2 = 0 Then Exit Sub
    If lRow1 = lRow2 And lCol1 = lCol2 Then Exit Sub
    bCanMerge = True
    For i = 1 To mvarCells.Count
        If mvarCells(i).Row >= lRow1 And mvarCells(i).Row <= lRow2 And mvarCells(i).Col >= lCol1 And mvarCells(i).Col <= lCol2 Then
            If Len(mvarCells(i).MergeInfo) > 0 Then
                '如果其中某个单元格已经合并，那么不允许再合并！！！
                bCanMerge = False
            End If
        End If
    Next
    If bCanMerge = False Then Exit Sub
    MergeCells lRow1, lCol1, lRow2, lCol2
'    m_SelStartRow = 0
'    m_SelStartCol = 0
'    m_SelEndRow = 0
'    m_SelEndCol = 0
End Sub

Public Sub MergeCells(ByVal lStartRow As Long, ByVal lStartCol As Long, ByVal lEndRow As Long, ByVal lEndCol As Long, _
    Optional ByVal bRefresh As Boolean = True)
    
    '合并单元格
    Dim sText As String, i As Long, j As Long, k As Long, lKey As Long
    If Len(Cell(lStartRow, lStartCol).MergeInfo) = 16 Then Exit Sub
    sText = Format(lStartRow, "0000") & Format(lStartCol, "0000") & Format(lEndRow, "0000") & Format(lEndCol, "0000")
    Cell(lStartRow, lStartCol).MergeInfo = sText
    For i = lStartRow To lEndRow
        For j = lStartCol To lEndCol
            If i <> lStartRow Or j <> lStartCol Then
                lKey = CellKey(i, j)
                mvarCells("K" & lKey).MergeInfo = sText
'                mvarCells("K" &lKey).Text = ""
                mvarCells("K" & lKey).Visibled = False
            End If
        Next
    Next i
    
    '修正单元格高度
    Dim lH As Long, lTmp As Long, lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long, lKey2 As Long
    m_bInResize = True
    lKey = CellKey(lStartRow, lStartCol)
LL:
    For i = lKey To mvarCells.Count
        If mvarCells(i).MergeInfo <> "" And Len(mvarCells(i).MergeInfo) = 16 Then
            lRow1 = mvarCells(i).MergeStartRow
            lCol1 = mvarCells(i).MergeStartCol
            lRow2 = mvarCells(i).MergeEndRow
            lCol2 = mvarCells(i).MergeEndCol
            If mvarCells(i).Row = lRow1 And mvarCells(i).Col = lCol1 Then
                '一次性处理
                lH = 0
                For j = lRow1 To lRow2
                    lH = lH + RowInfo(j).RowHeight
                Next
                '修正末行高度
                lKey = CellKey(lRow1, lCol1)
                lTmp = mvarCells("K" & lKey).EvaluateTextHeight(Me)
                If lH < lTmp Then
                    RowHeight(lRow2) = RowHeight(lRow2) + (lTmp - lH)
                    '同时要影响到其他有关本行的合并行的高度
                    GoTo LL
                    lH = lTmp
                End If
                mvarCells("K" & lKey).Height = lH
                
                For j = lRow1 To lRow2
                    For k = lCol1 To lCol2
                        If j <> lRow1 Or k <> lCol1 Then
                            lKey2 = CellKey(j, k)
                            mvarCells("K" & lKey2).Visibled = False
'                            mvarCells("K" & lKey2).Text = ""
                        End If
                    Next k
                Next j
            End If
        End If
    Next
    m_bInResize = False
    
    m_bDirty = True
    If bRefresh Then Call Refresh   '重新修正高度、宽度
End Sub

Public Sub DisMergeCells(ByVal lRow As Long, lCol As Long)
    '取消合并
    Dim sText As String, i As Long, j As Long, lKey As Long, lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long
    lKey = CellKey(lRow, lCol)
    If lKey > 0 Then
        sText = mvarCells("K" & lKey).MergeInfo
        If Len(sText) = 16 Then
            lRow1 = Val(Mid(sText, 1, 4))
            lCol1 = Val(Mid(sText, 5, 4))
            lRow2 = Val(Mid(sText, 9, 4))
            lCol2 = Val(Mid(sText, 13, 4))
            For i = lRow1 To lRow2
                For j = lCol1 To lCol2
                    lKey = CellKey(i, j)
                    If lKey > 0 Then
                        mvarCells("K" & lKey).MergeInfo = ""
                        mvarCells("K" & lKey).Dirty = True
                        mvarCells("K" & lKey).Visibled = True
                    End If
                Next j
            Next i
        End If
        '恢复第一个单元格的位置
        mvarCells(CellKey(lRow, lCol)).Height = RowInfo(lRow).RowHeight
    End If
    
    m_bDirty = True
    Refresh
End Sub

Public Sub ShowProperty(ByRef frmParent As Object, ByRef oTable As Object, Optional ByVal lStartTab As Long)
    '显示属性对话框
    frmProperty.ShowMe Me, oTable, lStartTab
End Sub

'#########################################################################################################
'## 功能：  设置自动行高
'## 参数：  lRow:           行
'##         lMinimumHeight: 最小高度
'## 返回：  无
'#########################################################################################################
Private Sub AutoHeightRow(ByVal lRow As Long, Optional ByVal lMinimumHeight As Long = 0)
    Dim lCol As Long
    Dim lHeight As Long
    Dim lMaxHeight As Long
    Dim lKey As Long, sText As String
    
    For lCol = 1 To mvarColCount
        With mvarCells("K" & RowColInfo(lRow, lCol))
            If Len(.MergeInfo) = 16 Then
                lHeight = 0                         '暂时不计算合并单元格高度（通过最后修正来确定合并单元格的高度）
            Else
                lHeight = .EvaluateTextHeight(Me)   '文本/图片的实际混合高度
            End If
            If .Icon > 0 Then
                If lHeight < p_lIconHeight Then
                    '保证高度至少大于图标高度
                    lHeight = p_lIconHeight + (.GridLineWidth + .Margin * p_TPPY) * 2
                End If
            End If
            If (lHeight < mvarMinRowHeight) Then
                '保证高度至少大于最小高度
                lHeight = mvarMinRowHeight
            End If
            If (lHeight > lMaxHeight) Then
                lMaxHeight = lHeight        '取最大高度
            End If
        End With
        '设置行高
        RowHeight(lRow) = lMaxHeight
'        RowInfo(lRow).ActualRowHeight = lMaxHeight  '该行的实际内容最大高度！
    Next
End Sub

'#########################################################################################################
'## 功能：  用背景填充指定矩形。可以使背景色或者位图，取决于交替色是否被设置。
'## 参数：  lhDC:       目标设备场景句柄
'##         tR:         需要填充的界矩形
'##         lOffsetX:   绘制位图时网格左上角的水平偏移位置
'##         lOffsetY:   绘制位图时网格左上角的垂直偏移位置
'##         bAlternate: 是否是交替行
'## 返回：  无
'## TOTO：  考虑实现图像半透明处理效果
'#########################################################################################################
Private Sub pFillBackground( _
      ByVal lhDC As Long, _
      ByRef tR As RECT, _
      ByVal lOffsetX As Long, _
      ByVal lOffsetY As Long, _
      ByVal bAlternate As Boolean)
      
    Dim hBr As Long

    If (bAlternate) And Not (mvarAlternateRowBackColor = -1) Then
        '如果是交替行且具有交替色
        hBr = CreateSolidBrush(TranslateColor(mvarAlternateRowBackColor))    '创建交替色的纯色画刷
        FillRect lhDC, tR, hBr  '填充矩形
        DeleteObject hBr        '释放资源
    Else
        '否则
        If (m_bBitmap) Then
            '如果为位图背景，则平铺图片
            TileArea lhDC, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, m_hDCSrc, m_lBitmapW, m_lBitmapH, lOffsetX, lOffsetY
        Else
            '为纯色背景
            If Not (mvarEnabled) Then
                '无效时，创建系统画刷（按钮表面颜色）
                hBr = GetSysColorBrush(vbButtonFace And &H1F&)
            Else
                '有效时
                If (mvarBackColor And &H80000000) = &H80000000 Then
                    '自定义颜色
                    hBr = GetSysColorBrush(mvarBackColor And &H1F&)
                Else
                    '系统颜色
                    hBr = CreateSolidBrush(TranslateColor(mvarBackColor))
                End If
            End If
            '填充矩形
            FillRect lhDC, tR, hBr
            '释放资源
            DeleteObject hBr
        End If
    End If
End Sub

'#########################################################################################################
'## 功能：  创建内存DC，加速图形操作
'## 返回：  无
'#########################################################################################################
Private Sub BuildMemDC()
    Dim tR As RECT
    Dim hBr As Long
   
    If (m_hBmp <> 0) Then
        '如果内存位图存在
        If (m_hBmpOld <> 0) Then
            '直接选入旧的位图
            SelectObject m_hDC, m_hBmpOld
        End If
        '释放资源
        If (m_hBmp <> 0) Then
            DeleteObject m_hBmp
        End If
        m_hBmp = 0
        m_hBmpOld = 0
    End If
    
    If (m_hDC = 0) Then
        '如果m_hDC不存在，则创建控件兼容DC
        m_hDC = CreateCompatibleDC(UserControl.hDC)
    End If
    
    If (m_hDC <> 0) Then
        '如果hDC存在
        m_hBmp = CreateCompatibleBitmap(UserControl.hDC, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY) '创建控件兼容位图
        If (m_hBmp <> 0) Then
            '创建成功
            m_hBmpOld = SelectObject(m_hDC, m_hBmp) '选入创建的新位图，保存旧位图
            If (m_hBmpOld = 0) Then
                '如果旧位图不存在，这清除资源
                DeleteObject m_hBmp
                DeleteObject m_hDC
                m_hBmp = 0
                m_hDC = 0
            Else
                '否则，设置绘图环境，并绘制背景
                SetTextColor m_hDC, TranslateColor(mvarForeColor)               '设置字体颜色
                SetBkColor m_hDC, TranslateColor(mvarBackColor)                 '设置背景色
                SetBkMode m_hDC, TRANSPARENT                                    '设置背景模式为透明
                tR.Right = Screen.Width \ Screen.TwipsPerPixelX                 '计算矩形宽度
                tR.Bottom = UserControl.ScaleHeight                             '计算矩形高度
                hBr = CreateSolidBrush(TranslateColor(mvarBackColor))           '创建单色画刷
                FillRect m_hDC, tR, hBr                                         '填充矩形
                DeleteObject hBr                                                '删除画刷
            End If
        Else
            '创建失败，则释放资源
            DeleteObject m_hDC
            m_hDC = 0
        End If
    End If
End Sub

'#########################################################################################################
'## 功能：  判断鼠标点下面的是哪一个单元格
'## 参数：  X、Y:               相对于控件的坐标
'##         lRow、lCol:         用于返回找到的单元格的行、列，否则为0
'##         lCellKey:           单元格在对象数组中的关键字
'## 返回：  返回找到的单元格的行列
'#########################################################################################################
Public Function CellFromPoint( _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByRef lRow As Long, _
    ByRef lCol As Long, _
    Optional ByRef lCellKey As Long = 0) As Boolean

    
    Dim i As Long, lR As Long, lC As Long

    lCol = 0: lRow = 0
    For i = 1 To mvarCells.Count
        If mvarCells(i).Visibled Then
            With mvarCells(i)
                lR = .Row
                lC = .Col
                If X > ColInfo(lC).LeftX And X < ColInfo(lC).LeftX + .Width And _
                    Y > RowInfo(lR).TopY And Y < RowInfo(lR).TopY + .Height Then
                    lRow = lR
                    lCol = lC
                    lCellKey = mvarCells(i).Key
                    CellFromPoint = True
                    Exit For
                End If
            End With
        End If
    Next
End Function

Private Sub txtEdit_Change()
    Dim lHOld As Long, lHNew As Long
    If m_CellKeyEdit > 0 And mvarAutoHeight Then
        lHOld = txtEdit.Height ' mvarCells("K" & m_CellKeyEdit).EvaluateTextHeight(Me)
        lHNew = mvarCells("K" & m_CellKeyEdit).EvaluateTextHeight(Me, txtEdit.Text) - 2 * mvarCells("K" & m_CellKeyEdit).Margin * p_TPPY - mvarGridLineWidth * p_TPPY
        If lHOld < lHNew And txtEdit.Visible Then
            '高度自动增加
            txtEdit.Height = lHNew
            '更新到控件
            RaiseEvent CancelEdit
            mvarCells("K" & m_CellKeyEdit).Text = txtEdit.Text
            mvarCells("K" & m_CellKeyEdit).Dirty = True
            mvarModified = True
            Refresh False, True, m_CellKeyEdit
            RaiseEvent Resize(UserControl.Width, UserControl.Height)
        End If
    End If
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long, sText As String, lRow As Long, lCol As Long
    Dim i As Long, j As Long, bFinded As Boolean
    Dim cursorPos As POINTAPI
    RaiseEvent KeyDown(KeyCode, Shift)
    
    GetCursorPos cursorPos

    Select Case KeyCode
    Case vbKeyUp
        If m_bInEdit Then
            If GetCurLine(txtEdit.hWnd) = 1 And mvarCells("K" & m_CellKeyEdit).Row > 1 Then
                '在行首按Up
                If mvarCells("K" & m_CellKeySelected).Row > 1 Then
                    '取消编辑
                    pCancelEdit True
                    
                    lRow = mvarCells("K" & m_CellKeySelected).Row - 1
                    lCol = mvarCells("K" & m_CellKeySelected).Col
                    sText = Cell(lRow, lCol).MergeInfo
                    If Len(sText) = 16 Then
                        lRow = Val(Mid(sText, 1, 4))
                        lCol = Val(Mid(sText, 5, 4))
                    End If
                    For i = lRow To 1 Step -1
                        If Cell(i, lCol).Visibled Then bFinded = True: Exit For
                    Next
                    If bFinded Then
                        lRow = i
                        SelectCell , lRow, lCol
                    End If
                    KeyCode = 0
                    SetFocusAPI UserControl.hWnd
                End If
            End If
        End If
    Case vbKeyDown
        If m_bInEdit Then
            If GetLineCount(txtEdit.hWnd) = GetCurLine(txtEdit.hWnd) And mvarCells("K" & m_CellKeyEdit).Row < mvarRowCount Then
                '在末行按Down
                If mvarCells("K" & m_CellKeySelected).Row < mvarRowCount Then
                    '取消编辑
                    pCancelEdit True
                    
                    If Len(mvarCells("K" & m_CellKeySelected).MergeInfo) = 16 Then
                        sText = mvarCells("K" & m_CellKeySelected).MergeInfo
                        lRow = Val(Mid(sText, 9, 4)) + 1
                        lCol = mvarCells("K" & m_CellKeySelected).Col
                    Else
                        lRow = mvarCells("K" & m_CellKeySelected).Row + 1
                        lCol = mvarCells("K" & m_CellKeySelected).Col
                    End If
                    lRow = IIf(lRow > mvarRowCount, mvarRowCount, lRow)
                    lCol = IIf(lCol > mvarColCount, mvarColCount, lCol)
                    sText = Cell(lRow, lCol).MergeInfo
                    If sText <> "" And Len(sText) = 16 Then
                        lRow = Val(Mid(sText, 1, 4))
                        lCol = Val(Mid(sText, 5, 4))
                    End If
                    For i = lRow To mvarRowCount
                        If Cell(i, lCol).Visibled Then bFinded = True: Exit For
                    Next
                    If bFinded Then
                        lRow = i
                        SelectCell , lRow, lCol
                    End If
                    KeyCode = 0
                    SetFocusAPI UserControl.hWnd
                End If
            End If
        End If
    Case vbKeyLeft
        If m_bInEdit Then
            If txtEdit.SelStart = 0 And txtEdit.SelLength = 0 Then
                '在首列按Left
                If mvarCells("K" & m_CellKeySelected).Col > 1 Then
                    '取消编辑
                    pCancelEdit True
                    
                    lRow = mvarCells("K" & m_CellKeySelected).Row
                    lCol = mvarCells("K" & m_CellKeySelected).Col - 1
                    sText = Cell(lRow, lCol).MergeInfo
                    If Len(sText) = 16 Then
                        lRow = Val(Mid(sText, 1, 4))
                        lCol = Val(Mid(sText, 5, 4))
                    End If
                    For j = lCol To 1 Step -1
                        If Cell(lRow, j).Visibled Then bFinded = True: Exit For
                    Next
                    If bFinded Then
                        lCol = j
                        SelectCell , lRow, lCol
                    End If
                    KeyCode = 0
                    SetFocusAPI UserControl.hWnd
                End If
            End If
        End If
    Case vbKeyRight
        If m_bInEdit Then
            If txtEdit.SelStart = Len(txtEdit) Then
                '在末尾按Right
                If mvarCells("K" & m_CellKeySelected).Col < mvarColCount Then
                    '取消编辑
                    pCancelEdit True
                    
                    If Len(mvarCells("K" & m_CellKeySelected).MergeInfo) = 16 Then
                        sText = mvarCells("K" & m_CellKeySelected).MergeInfo
                        lRow = mvarCells("K" & m_CellKeySelected).Row
                        lCol = Val(Mid(sText, 13, 4)) + 1
                    Else
                        lRow = mvarCells("K" & m_CellKeySelected).Row
                        lCol = mvarCells("K" & m_CellKeySelected).Col + 1
                    End If
                    lRow = IIf(lRow > mvarRowCount, mvarRowCount, lRow)
                    lCol = IIf(lCol > mvarColCount, mvarColCount, lCol)
                    sText = Cell(lRow, lCol).MergeInfo
                    If sText <> "" And Len(sText) = 16 Then
                        lRow = Val(Mid(sText, 1, 4))
                        lCol = Val(Mid(sText, 5, 4))
                    End If
                    For j = lCol To mvarColCount
                        If Cell(lRow, j).Visibled Then bFinded = True: Exit For
                    Next
                    If bFinded Then
                        lCol = j
                        SelectCell , lRow, lCol
                    End If
                    KeyCode = 0
                    SetFocusAPI UserControl.hWnd
                End If
            End If
        End If
    Case vbKeyPageUp
        
    Case vbKeyPageDown
        
    Case vbKeyHome
        
    Case vbKeyEnd
        
    Case vbKeySpace
        
    Case vbKeyTab
        '在界面上有其他控件可以获取焦点时：通过 IOLEInPlaceActiveObject 进行捕获！消息由Usercontrol_KeyDown事件统一处理！
        '如果界面上只有本控件，那么Text控件可以捕获Tab按键，此时应该禁止输入Tab按键
        KeyCode = 0
    Case vbKeyReturn
        KeyCode = 0
        If mvarSingleLine Then
            '单行时回车直接保存
            pCancelEdit False       '接受输入值
        ElseIf (Shift And vbCtrlMask) > 0 Then
            '多行时Ctrl＋回车才保存
            pCancelEdit False       '接受输入值
        End If
    Case vbKeyEscape
        KeyCode = 0
        pCancelEdit False, True     '不接受输入值
    End Select
End Sub

Private Function GetCurLine(hWnd As Long) As Long
    Dim l  As Long
    l = SendMessage(hWnd, EM_LINEINDEX, -1, 0)
    GetCurLine = SendMessage(hWnd, EM_LINEFROMCHAR, l, 0) + 1
End Function

Public Function GetLineCount(hWnd As Long) As Long
   GetLineCount = SendMessageLong(hWnd, EM_GETLINECOUNT, 0, 0)
End Function

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyTab
        If mvarTabKeyMoveNextCell Then
            KeyAscii = 0    '不允许输入Tab按键
            UserControl_KeyDown vbKeyTab, 0
        End If
    End Select
End Sub

Private Sub txtEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtEdit.MousePointer = 0
    If mvarHotTrack Then
        If mvarCells("K" & m_CellKeyEdit).Hot = False Then
            If m_CellKeyHot > 0 Then mvarCells("K" & m_CellKeyHot).Hot = False
            mvarCells("K" & m_CellKeyEdit).Hot = True
            m_CellKeyHot = m_CellKeyEdit
            m_bDirty = True
            Call Draw
        End If
    End If
End Sub

Private Sub UserControl_DblClick()
    Dim tP As POINTAPI
    Dim lRow As Long, lCol As Long, lCellKey As Long, bFinded As Boolean, i As Long

    On Error GoTo ErrorHandler
    If (mvarEnabled And mvarInnerEdit) Then
        GetCursorPos tP
        ScreenToClient UserControl.hWnd, tP
        bFinded = CellFromPoint(tP.X * p_TPPX, tP.Y * p_TPPY, lRow, lCol, lCellKey)
        If bFinded Then
            If m_CellKeySelected <> lCellKey Then
                UserControl_MouseDown vbLeftButton, 0, UserControl.ScaleX(tP.X, vbPixels, UserControl.ScaleMode), UserControl.ScaleY(tP.Y, vbPixels, UserControl.ScaleMode)
            End If
            RaiseEvent DblClick(lRow, lCol)
            pRequestEdit 0, True
        End If
    End If
    Exit Sub
ErrorHandler:
    Debug.Assert False
    Exit Sub
    
    Resume 0
End Sub

Private Sub UserControl_GotFocus()
    m_bInFocus = True
    m_bDirty = True
    Call Draw
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long, sText As String, lRow As Long, lCol As Long
    Dim i As Long, j As Long, l As Long, m As Long, bFinded As Boolean
    Dim bCancel As Boolean
    
    RaiseEvent KeyDown(KeyCode, Shift)
    
    Select Case KeyCode
    Case vbKeyReturn
        If Shift = 2 Then
            pRequestEdit KeyCode
        End If
    Case vbKeyUp
        '非编辑状态时，单元格位置移动
        If Not m_bInEdit And m_CellKeySelected > 0 Then
            If mvarCells("K" & m_CellKeySelected).Row > 1 Then
                lRow = mvarCells("K" & m_CellKeySelected).Row - 1
                lCol = mvarCells("K" & m_CellKeySelected).Col
                sText = Cell(lRow, lCol).MergeInfo
                If Len(sText) = 16 Then
                    lRow = Val(Mid(sText, 1, 4))
                    lCol = Val(Mid(sText, 5, 4))
                End If
                For i = lRow To 1 Step -1
                    If Cell(i, lCol).Visibled Then bFinded = True: Exit For
                Next
                If bFinded Then
                    lRow = i
                    SelectCell , lRow, lCol
                End If
            End If
        End If
    Case vbKeyDown
        '非编辑状态时，单元格位置移动
        If Not m_bInEdit And m_CellKeySelected > 0 Then
            If mvarCells("K" & m_CellKeySelected).Row < mvarRowCount Then
                If mvarCells("K" & m_CellKeySelected).MergeInfo <> "" Then
                    sText = mvarCells("K" & m_CellKeySelected).MergeInfo
                    lRow = Val(Mid(sText, 9, 4)) + 1
                    lCol = mvarCells("K" & m_CellKeySelected).Col
                Else
                    lRow = mvarCells("K" & m_CellKeySelected).Row + 1
                    lCol = mvarCells("K" & m_CellKeySelected).Col
                End If
                lRow = IIf(lRow > mvarRowCount, mvarRowCount, lRow)
                lCol = IIf(lCol > mvarColCount, mvarColCount, lCol)
                sText = Cell(lRow, lCol).MergeInfo
                If Len(sText) = 16 Then
                    lRow = Val(Mid(sText, 1, 4))
                    lCol = Val(Mid(sText, 5, 4))
                End If
                For i = lRow To mvarRowCount
                    If Cell(i, lCol).Visibled Then bFinded = True: Exit For
                Next
                If bFinded Then
                    lRow = i
                    SelectCell , lRow, lCol
                End If
            End If
        End If
    Case vbKeyLeft
        '非编辑状态时，单元格位置移动
        If Not m_bInEdit And m_CellKeySelected > 0 Then
            If mvarCells("K" & m_CellKeySelected).Col > 1 Then
                lRow = mvarCells("K" & m_CellKeySelected).Row
                lCol = mvarCells("K" & m_CellKeySelected).Col - 1
                sText = Cell(lRow, lCol).MergeInfo
                If Len(sText) = 16 Then
                    lRow = Val(Mid(sText, 1, 4))
                    lCol = Val(Mid(sText, 5, 4))
                End If
                For j = lCol To 1 Step -1
                    If Cell(lRow, j).Visibled Then bFinded = True: Exit For
                Next
                If bFinded Then
                    lCol = j
                    SelectCell , lRow, lCol
                End If
            End If
        End If
    Case vbKeyRight
        '非编辑状态时，单元格位置移动
        If Not m_bInEdit And m_CellKeySelected > 0 Then
            If mvarCells("K" & m_CellKeySelected).Col < mvarColCount Then
                If mvarCells("K" & m_CellKeySelected).MergeInfo <> "" Then
                    sText = mvarCells("K" & m_CellKeySelected).MergeInfo
                    lRow = mvarCells("K" & m_CellKeySelected).Row
                    lCol = Val(Mid(sText, 13, 4)) + 1
                Else
                    lRow = mvarCells("K" & m_CellKeySelected).Row
                    lCol = mvarCells("K" & m_CellKeySelected).Col + 1
                End If
                lRow = IIf(lRow > mvarRowCount, mvarRowCount, lRow)
                lCol = IIf(lCol > mvarColCount, mvarColCount, lCol)
                sText = Cell(lRow, lCol).MergeInfo
                If Len(sText) = 16 Then
                    lRow = Val(Mid(sText, 1, 4))
                    lCol = Val(Mid(sText, 5, 4))
                End If
                For j = lCol To mvarColCount
                    If Cell(lRow, j).Visibled Then bFinded = True: Exit For
                Next
                If bFinded Then
                    lCol = j
                    SelectCell , lRow, lCol
                End If
            End If
        End If
    Case vbKeyHome
        '非编辑状态时，单元格位置移动
        If Not m_bInEdit And m_CellKeySelected > 0 Then
            lRow = mvarCells("K" & m_CellKeySelected).Row
            lCol = 1
            sText = Cell(lRow, lCol).MergeInfo
            If Len(sText) = 16 Then
                lCol = Val(Mid(sText, 5, 4))
            End If
            SelectCell , lRow, lCol
        End If
    Case vbKeyEnd
        '非编辑状态时，单元格位置移动
        If Not m_bInEdit And m_CellKeySelected > 0 Then
            lRow = mvarCells("K" & m_CellKeySelected).Row
            lCol = mvarColCount
            sText = Cell(lRow, lCol).MergeInfo
            If Len(sText) = 16 Then
                lCol = Val(Mid(sText, 5, 4))
            End If
            SelectCell , lRow, lCol
        End If
    Case vbKeyTab
        '通过 IOLEInPlaceActiveObject 进行捕获！
        KeyCode = 0     '避免界面没有其他控件时Text控件捕获Tab按键！！！
        If mvarTabKeyMoveNextCell Then
            If m_CellKeySelected > 0 Then
                If mvarCells("K" & m_CellKeySelected).Col < mvarColCount Then
                    '相当于后移一格
                    If mvarCells("K" & m_CellKeySelected).MergeInfo <> "" Then
                        sText = mvarCells("K" & m_CellKeySelected).MergeInfo
                        lRow = mvarCells("K" & m_CellKeySelected).Row
                        lCol = Val(Mid(sText, 13, 4)) + 1
                    Else
                        lRow = mvarCells("K" & m_CellKeySelected).Row
                        lCol = mvarCells("K" & m_CellKeySelected).Col + 1
                    End If
                    lRow = IIf(lRow > mvarRowCount, mvarRowCount, lRow)
                    lCol = IIf(lCol > mvarColCount, mvarColCount, lCol)
LL1:
                    If lRow > mvarRowCount Then
                        '说明末行超标
                        lRow = 1
                    End If
                    sText = Cell(lRow, lCol).MergeInfo
                    If Len(sText) = 16 Then
                        If lRow <> Val(Mid(sText, 1, 4)) Or lCol <> Val(Mid(sText, 5, 4)) Then
                            '非合并单元格的第一个单元，则不选择该单元格
                            lCol = Val(Mid(sText, 13, 4))
                            If lCol < mvarColCount Then
                                lCol = lCol + 1
                            Else
                                lRow = lRow + 1
                                lCol = 1
                            End If
                            GoTo LL1
                        End If
                    End If
                    For j = lCol To mvarColCount
                        If Cell(lRow, j).Visibled Then bFinded = True: Exit For
                    Next
                    If bFinded Then
                        If m_bInEdit Then pCancelEdit True
                        lCol = j
                        SelectCell , lRow, lCol
                    End If
                ElseIf mvarCells("K" & m_CellKeySelected).Col = mvarColCount And mvarCells("K" & m_CellKeySelected).Row < mvarRowCount Then
                    '已经在末列，换到下一行
                    If mvarCells("K" & m_CellKeySelected).MergeInfo <> "" Then
                        sText = mvarCells("K" & m_CellKeySelected).MergeInfo
                        lRow = mvarCells("K" & m_CellKeySelected).Row + 1
                        lCol = 1
                    Else
                        lRow = mvarCells("K" & m_CellKeySelected).Row + 1
                        lCol = 1
                    End If
                    lRow = IIf(lRow > mvarRowCount, mvarRowCount, lRow)
                    lCol = IIf(lCol > mvarColCount, mvarColCount, lCol)
LL2:
                    If lRow > mvarRowCount Then
                        '说明末行超标
                        lRow = 1
                    End If
                    sText = Cell(lRow, lCol).MergeInfo
                    If Len(sText) = 16 Then
                        If lRow <> Val(Mid(sText, 1, 4)) Or lCol <> Val(Mid(sText, 5, 4)) Then
                            '非合并单元格的第一个单元，则不选择该单元格
                            lCol = Val(Mid(sText, 13, 4))
                            If lCol < mvarColCount Then
                                lCol = lCol + 1
                            Else
                                lRow = lRow + 1
                                lCol = 1
                            End If
                            GoTo LL2
                        End If
                    End If
                    For j = lCol To mvarColCount
                        If Cell(lRow, j).Visibled Then bFinded = True: Exit For
                    Next
                    If bFinded Then
                        If m_bInEdit Then pCancelEdit True
                        lCol = j
                        SelectCell , lRow, lCol
                    End If
                Else
                    '表示是最后一个单元格，此时跳到第一个单元格
                    For i = 1 To mvarRowCount
                        For j = 1 To mvarRowCount
                            If mvarCells(CellKey(i, j)).Visibled Then
                                bFinded = True
                                lRow = i
                                lCol = j
                                Exit For
                            End If
                        Next j
                        If bFinded Then Exit For
                    Next i
                    If bFinded Then
                        If m_bInEdit Then pCancelEdit True
                        SelectCell , i, j
                    End If
                End If
            End If
        End If
    Case vbKeyDelete, vbKeyBack
        If m_CellKeySelected > 0 Then
            '请求编辑
            RaiseEvent RequestEdit(mvarCells("K" & m_CellKeySelected).Row, mvarCells("K" & m_CellKeySelected).Col, KeyCode, bCancel)
            If (Not bCancel) Then
                '删除当前单元
                If mvarCells("K" & m_CellKeySelected).Text <> "" And mvarCells("K" & m_CellKeySelected).Protected = False Then
                    mvarCells("K" & m_CellKeySelected).Text = ""
                    Refresh False, mvarSingleLine = False And mvarAutoHeight, m_CellKeySelected
                End If
            End If
        End If
    Case Else
        If KeyCode = 229 Then
            '中文输入
            If mvarInnerEdit Then pRequestEdit 0, False
'        Else
'            '英文输入
'            If mvarInnerEdit Then pRequestEdit KeyCode, False
        End If
    End Select
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Dim Shift As Integer, bCancel As Boolean
    If (mvarRowCount = 0) Or (mvarColCount = 0) Then Exit Sub
    If mvarEnabled Then
        Select Case KeyAscii
        Case 3
            '复制
            If m_CellKeySelected > 0 Then
                Clipboard.Clear
                Clipboard.SetText mvarCells("K" & m_CellKeySelected).Text, vbCFText
            End If
        Case 22
            '粘贴
            If m_CellKeySelected > 0 Then
                If mvarCells("K" & m_CellKeySelected).Protected = False Then
                    RaiseEvent RequestEdit(mvarCells("K" & m_CellKeySelected).Row, mvarCells("K" & m_CellKeySelected).Col, KeyAscii, bCancel)
                    If (Not bCancel) Then
                        mvarCells("K" & m_CellKeySelected).Text = Clipboard.GetText(vbCFText)
                        Refresh False, mvarSingleLine = False And mvarAutoHeight, m_CellKeySelected
                    End If
                End If
            End If
        Case vbKeyEscape, vbKeyTab, vbKeyBack ', vbKeyDelete
        
        Case Else
            If mvarInnerEdit Then pRequestEdit KeyAscii, False
        End Select
        RaiseEvent KeyPress(KeyAscii)
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    m_bInFocus = False
    m_bDirty = True
    Call Draw
End Sub

Private Sub SelectCell(Optional ByVal lCellKey As Long, Optional ByVal lRow As Long, Optional ByVal lCol As Long)
    Dim Rct As RECT, i As Long, j As Long
    If lCellKey = 0 Then
        lCellKey = CellKey(lRow, lCol)
        If lCellKey = 0 Then Exit Sub
    Else
        lRow = mvarCells("K" & lCellKey).Row
        lCol = mvarCells("K" & lCellKey).Col
    End If
    If m_CellKeySelected = lCellKey Then Exit Sub
    RaiseEvent SelectionChange(lRow, lCol)
    
    'If m_CellKeySelected > 0 Then mvarCells("K" & m_CellKeySelected).Selected = False
    For i = 1 To mvarCells.Count
        mvarCells(i).Selected = False
    Next
    mvarCells("K" & lCellKey).Selected = True
    m_CellKeySelected = lCellKey
    
    m_bDirty = True
    Call Draw
    
    '显示提示信息
    GetWindowRect m_hWnd, Rct
    If mvarShowToolTipText And mvarCells("K" & lCellKey).ToolTipText <> "" Then
        ShowTipInfor Format(mvarCells("K" & lCellKey).ToolTipText, mvarCells("K" & lCellKey).FormatString), _
            ColInfo(lCol).LeftX / p_TPPX + mvarCells("K" & lCellKey).Width / p_TPPX + Rct.Left, _
            RowInfo(lRow).TopY / p_TPPY + Rct.Top + mvarGridLineWidth, _
            mvarCells("K" & lCellKey).Width / p_TPPX
    Else
        If frmTips.Visible Then frmTips.Hide
    End If
End Sub

Private Sub pDrawLineV(ByVal Pos As Long)
    '模拟Word中的表格网格垂直参考线的绘制
    Dim lhDC As Long
    Dim rClient As RECT
    Static oldPosV As Long
    Dim lCount As Long
    
    '在本控件上绘制线条
    If hWnd = 0 Then Exit Sub
    GetClientRect hWnd, rClient
    InflateRect rClient, 0, -2
    lhDC = GetDC(hWnd)
    
    If Pos = 0 Then
        rClient.Left = oldPosV - 1
        rClient.Right = oldPosV + 1
        InvalidateRect hWnd, rClient, False
    Else
        If Pos <> oldPosV Then
            rClient.Left = Pos
            rClient.Right = Pos
            SetTextColor lhDC, TranslateColor(vbBlack)    '设置字体颜色
            mAPI.DrawFocusRect lhDC, rClient
        End If
    End If
    
    '在指定控件上绘制
    If hWndBound = 0 Then Exit Sub
    GetClientRect hWndBound, rClient
    InflateRect rClient, 0, -2
    lhDC = GetDC(hWndBound)

    If Pos = 0 Then
        rClient.Left = oldPosV + m_OffsetX / p_TPPX - 1
        rClient.Right = oldPosV + m_OffsetX / p_TPPX + 1
        InvalidateRect hWndBound, rClient, True
    Else
        If Pos <> oldPosV Then
            rClient.Left = Pos + m_OffsetX / p_TPPX
            rClient.Right = Pos + m_OffsetX / p_TPPX
            SetTextColor lhDC, TranslateColor(vbBlack)    '设置字体颜色
            mAPI.DrawFocusRect lhDC, rClient
       End If
    End If
    oldPosV = Pos
End Sub

Private Sub pDrawLineH(ByVal Pos As Long)
    '模拟Word中的表格网格水平参考线的绘制
    Dim lhDC As Long
    Dim rClient As RECT
    Static oldPosH As Long
    Dim lCount As Long
    
    '在本控件上绘制线条
    If hWnd = 0 Then Exit Sub
    GetClientRect hWnd, rClient
    InflateRect rClient, 0, -2
    lhDC = GetDC(hWnd)
    
    If Pos = 0 Then
        rClient.Top = oldPosH - 1
        rClient.Bottom = oldPosH + 1
        InvalidateRect hWnd, rClient, False
    Else
        If Pos <> oldPosH Then
            rClient.Top = Pos
            rClient.Bottom = Pos
            SetTextColor lhDC, TranslateColor(vbBlack)    '设置字体颜色
            mAPI.DrawFocusRect lhDC, rClient
        End If
    End If
    
    '在指定控件上绘制
    If hWndBound = 0 Then Exit Sub
    GetClientRect hWndBound, rClient
    InflateRect rClient, 0, -2
    lhDC = GetDC(hWndBound)

    If Pos = 0 Then
        rClient.Top = oldPosH + m_OffsetY / p_TPPY - 1
        rClient.Bottom = oldPosH + m_OffsetY / p_TPPY + 1
        InvalidateRect hWndBound, rClient, True
    Else
        If Pos <> oldPosH Then
            rClient.Top = Pos + m_OffsetY / p_TPPY
            rClient.Bottom = Pos + m_OffsetY / p_TPPY
            SetTextColor lhDC, TranslateColor(vbBlack)    '设置字体颜色
            mAPI.DrawFocusRect lhDC, rClient
       End If
    End If
    oldPosH = Pos
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lCellKey As Long, bFinded As Boolean, i As Long
    Dim hPen As Long, hPenOld As Long, tJ As POINTAPI
    
    If mvarEnabled = False Then Exit Sub
    
    m_bAdjustColWidth = False
    m_bAdjustRowHeight = False
    If Button = vbLeftButton Then
        '先看是否在某列边缘
        For i = mvarColCount To 1 Step -1   '（反向判断，末列优先）
            If Abs(X - (ColInfo(i).LeftX + ColInfo(i).ColWidth)) <= p_TPPX * 1 And ColInfo(i).FixedWidth = False And ColInfo(i).Visible Then
                '点击在列的右边边框上，误差为1个象素
                m_bAdjustColWidth = True
                lCol = i
                Exit For
            End If
        Next i

        '再看是否在行边缘
        If m_bAdjustColWidth = False Then
            For i = mvarRowCount To 1 Step -1    '（反向判断，末行优先）
                If Abs(Y - (RowInfo(i).TopY + RowInfo(i).RowHeight)) <= p_TPPY * 1 And mvarAutoHeight = False And RowInfo(i).FixedHeight = False Then
                    '点击在列的右边边框上，误差为1个象素
                    m_bAdjustRowHeight = True
                    lRow = i
                    Exit For
                End If
            Next i
        End If
    End If
    
        If m_bAdjustColWidth = True And lCol = 0 Then
            MsgBox "请回忆刚才您的前几步操作步骤，请抓图并将信息反馈给中联技术中心。"
        End If
        
    If m_bAdjustColWidth And lCol > 0 Then
        If m_bInEdit Then pCancelEdit True
        '在调整列宽
        UserControl.MousePointer = 99
        UserControl.MouseIcon = imlCursor.ListImages(1).Picture
        '绘制参考线
        pDrawLineV (ColInfo(lCol).LeftX + ColInfo(lCol).ColWidth) / p_TPPX
        
        m_ColAdjust = lCol
        m_OldX = (ColInfo(lCol).LeftX + ColInfo(lCol).ColWidth)
    ElseIf m_bAdjustRowHeight And lRow > 0 Then
        If m_bInEdit Then pCancelEdit True
        '在调整行高
        UserControl.MousePointer = 99
        UserControl.MouseIcon = imlCursor.ListImages(2).Picture
        '绘制参考线
        pDrawLineH (RowInfo(lRow).TopY + RowInfo(lRow).RowHeight) / p_TPPY
        m_RowAdjust = lRow
        m_OldY = (RowInfo(lRow).TopY + RowInfo(lRow).RowHeight)
    ElseIf m_bAdjustColWidth = False And m_bAdjustRowHeight = False Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
        If Button = vbLeftButton Then
            bFinded = CellFromPoint(X, Y, lRow, lCol, lCellKey)
            If bFinded And mvarEnabled Then
                SelectCell lCellKey
                '多选
                m_SelStartRow = lRow
                m_SelStartCol = lCol
                m_SelEndRow = lRow
                m_SelEndCol = lCol
                m_bMouseDown = True
                
                If mvarSingleClickEdit And mvarInnerEdit Then
    '                Debug.Print "UserControl_MouseDown->pRequestEdit"
                    pRequestEdit 0, True
                End If
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lRow As Long, lCol As Long, lCellKey As Long, bFinded As Boolean, i As Long, j As Long
    Dim bAdjustColWidth As Boolean      '是否点击在列边框的网格上
    Dim bAdjustRowHeight As Boolean     '是否点击在列边框的网格上
    
    If mvarEnabled = False Then Exit Sub
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Button = vbRightButton Then Exit Sub
    If m_bAdjustColWidth And m_ColAdjust > 0 Then
        '在调整列宽过程中，绘制参考垂直线
        pDrawLineV 0      '先恢复
        If X < ColInfo(m_ColAdjust).LeftX + p_TPPX * 20 Then X = ColInfo(m_ColAdjust).LeftX + p_TPPX * 20  '至少宽度为8个象素
        pDrawLineV X / p_TPPX
        Exit Sub
    ElseIf m_bAdjustRowHeight And m_RowAdjust > 0 Then
        '在调整行高过程中，绘制参考水平线
        pDrawLineH 0      '先恢复
        If Y < RowInfo(m_RowAdjust).TopY + p_TPPY * 20 Then X = RowInfo(m_RowAdjust).TopY + p_TPPY * 20   '至少高度为8个象素
        pDrawLineH Y / p_TPPY
        Exit Sub
    End If
    
    '判断是否在调整列宽度
    bAdjustColWidth = False
    bAdjustRowHeight = False
    For i = 1 To mvarColCount
        If Abs(X - (ColInfo(i).LeftX + ColInfo(i).ColWidth)) <= p_TPPX * 1 And ColInfo(i).Visible Then
            '鼠标在一列的右边边框上，误差为1个象素
            bAdjustColWidth = True
            lCol = i
            Exit For
        End If
    Next i
    '再看是否在行边缘
    If bAdjustColWidth = False Then
        For i = mvarRowCount To 1 Step -1    '（反向判断，末行优先）
            If Abs(Y - (RowInfo(i).TopY + RowInfo(i).RowHeight)) <= p_TPPY * 1 And mvarAutoHeight = False And RowInfo(i).FixedHeight = False Then
                '点击在列的右边边框上，误差为1个象素
                bAdjustRowHeight = True
                lRow = i
                Exit For
            End If
        Next i
    End If
    If bAdjustColWidth Then
        '在列宽调整位置
        UserControl.MousePointer = 99
        UserControl.MouseIcon = imlCursor.ListImages(1).Picture
    ElseIf bAdjustRowHeight Then
        '在行高调整位置
        UserControl.MousePointer = 99
        UserControl.MouseIcon = imlCursor.ListImages(2).Picture
    Else
        '多选
        Dim lRow1 As Long, lRow2 As Long, lCol1 As Long, lCol2 As Long
        If m_bMouseDown Then
            bFinded = CellFromPoint(X, Y, lRow, lCol, lCellKey)
            If bFinded Then
                m_SelEndRow = lRow
                m_SelEndCol = lCol
                If m_SelStartRow > m_SelEndRow Then
                    lRow1 = m_SelEndRow
                    lRow2 = m_SelStartRow
                Else
                    lRow1 = m_SelStartRow
                    lRow2 = m_SelEndRow
                End If
                If m_SelStartCol > m_SelEndCol Then
                    lCol1 = m_SelEndCol
                    lCol2 = m_SelStartCol
                Else
                    lCol1 = m_SelStartCol
                    lCol2 = m_SelEndCol
                End If
                For i = 1 To mvarCells.Count
                    If mvarCells(i).Row >= lRow1 And mvarCells(i).Row <= lRow2 And _
                        mvarCells(i).Col >= lCol1 And mvarCells(i).Col <= lCol2 Then
                        mvarCells(i).Selected = True
                    Else
                        mvarCells(i).Selected = False
                    End If
                Next
                Call Refresh(False, False)
            End If
        End If
        UserControl.MousePointer = 0
        If mvarHotTrack Then
            bFinded = CellFromPoint(X, Y, lRow, lCol, lCellKey)
            If bFinded Then
                If m_CellKeyHot = lCellKey Then Exit Sub
                If m_CellKeyHot > 0 Then mvarCells("K" & m_CellKeyHot).Hot = False:
                mvarCells("K" & lCellKey).Hot = True
                m_CellKeyHot = lCellKey
                m_bDirty = True
                Call Draw
            End If
        End If
        If m_CellKeyHot > 0 And mvarHotTrack Then
            m_tmrHotTrack.Interval = 50
        Else
            m_tmrHotTrack.Interval = 0
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lH As Long
    m_bMouseDown = False
    If m_bAdjustColWidth And m_ColAdjust > 0 Then
        pDrawLineV 0      '恢复
        If X < ColInfo(m_ColAdjust).LeftX + p_TPPX * 20 Then X = ColInfo(m_ColAdjust).LeftX + p_TPPX * 20  '至少宽度为8个象素
        If Abs(X - m_OldX) > 15 Then ColInfo(m_ColAdjust).ColWidth = ColInfo(m_ColAdjust).ColWidth + (X - m_OldX)
        If ColInfo(m_ColAdjust).ColWidth < p_TPPX * 20 Then ColInfo(m_ColAdjust).ColWidth = p_TPPX * 20
        '刷新结果
        Refresh
        m_bAdjustColWidth = False
        UserControl.MousePointer = 0
        mvarModified = True
        If Abs(X - m_OldX) > 15 Then RaiseEvent Resize(UserControl.Width, UserControl.Height)
    ElseIf m_bAdjustRowHeight And m_RowAdjust > 0 Then
        pDrawLineH 0      '恢复
        If Y < RowInfo(m_RowAdjust).TopY + p_TPPY * 20 Then Y = RowInfo(m_RowAdjust).TopY + p_TPPY * 20   '至少宽度为8个象素
        If Abs(Y - m_OldY) > 15 Then lH = RowInfo(m_RowAdjust).RowHeight + (Y - m_OldY)
        If lH < p_TPPY * 20 Then lH = p_TPPY * 20
        RowHeight(m_RowAdjust) = lH
        '刷新结果
        Refresh False
        m_bAdjustRowHeight = False
        UserControl.MousePointer = 0
        mvarModified = True
        If Abs(Y - m_OldY) > 15 Then RaiseEvent Resize(UserControl.Width, UserControl.Height)
    ElseIf m_bAdjustColWidth = False And m_bAdjustRowHeight = False Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If
End Sub

Public Sub RaiseResizeEvent()
    RaiseEvent Resize(UserControl.Width, UserControl.Height)
End Sub

'#########################################################################################################
'## 功能：  用鼠标钩子来标识取消编辑
'## 参数：  iMsg、hwnd、x、y、hitTest
'## 返回：  无
'#########################################################################################################
Friend Function MouseEvent( _
      ByVal iMsg As Long, _
      ByVal hWnd As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal hitTest As Long _
   ) As Boolean
   
    Dim pt As POINTAPI
    Dim cursorPos As POINTAPI
    Dim lLeft As Long, lTop As Long, lRight As Long, lBottom As Long
    Dim lRow As Long, lCol As Long
    
    If (iMsg = WM_NCLBUTTONDOWN Or iMsg = WM_NCRBUTTONDOWN Or iMsg = WM_NCMBUTTONDOWN) Then '鼠标在非客户区域点击
        EndEdit '终止编辑
        If (m_bInEdit) Then
            ' We have requested to cancel the edit.
            MouseEvent = True
        End If
    ElseIf (iMsg = WM_RBUTTONDOWN Or iMsg = WM_LBUTTONDOWN Or iMsg = WM_MBUTTONDOWN) Then   '鼠标在客户区域点击
        ' 检查当前在哪个类的上面点击
        Dim className As String
        Dim hWndOver As Long
        Dim hWndParent As Long
        Dim hWndDesktop As Long
        Dim hWndChild As Long
        
        hWndDesktop = GetDesktopWindow()    '屏幕
        hWndOver = WindowFromPoint(X, Y)    '当前鼠标位置的窗体
        hWndParent = GetParent(hWndOver)    '鼠标位置窗体的父窗体
        
        ' The owner of a combo is the desktop
        If Not (hWndOver = hWndDesktop) Then
            '如果当前位置不是屏幕
            If (GetProp(hWndOver, MAGIC_END_EDIT_IGNORE_WINDOW_PROP) = 0) Then
                className = WindowClassName(hWndOver)   '获取当前窗体的类名称
                ' Extra check for ComboLBox probably isn't needed, but menus have a parent 0
                If (InStr(className, "ComboLBox") = 0) And (InStr(className, "#32768") = 0) Then ' second check!
                    ' Check if the mouse event is within the boundaries of
                    ' the cell that is being edited:
                    GetCursorPos cursorPos
                    LSet pt = cursorPos
                    ScreenToClient UserControl.hWnd, pt
                    
                    Dim tR As RECT
                    Dim lWidth As Long
                    Dim lHeight As Long
                    Dim ClickedInCell As Boolean
                    Dim lOffsetX As Long
                    
                    '判断是否点击在单元格内部
'                    mvarCells("K" & m_CellKeyEdit).GetCellTextBorder lLeft, lTop, lRight, lBottom
'                    Debug.Print m_CellKeyEdit
                    lRow = mvarCells("K" & m_CellKeyEdit).Row
                    lCol = mvarCells("K" & m_CellKeyEdit).Col
                    tR.Left = ColInfo(lCol).LeftX / p_TPPX
                    tR.Top = RowInfo(lRow).TopY / p_TPPY
                    tR.Right = tR.Left + mvarCells("K" & m_CellKeyEdit).Width / p_TPPX
                    tR.Bottom = tR.Top + mvarCells("K" & m_CellKeyEdit).Height / p_TPPY
                    If (pt.X >= tR.Left And pt.X <= tR.Right) Then
                        If (pt.Y >= tR.Top And pt.Y <= tR.Bottom) Then
                           ClickedInCell = True
                        End If
                    End If
                    
                    If frmTips.Visible Then
                        GetCursorPos cursorPos
                        LSet pt = cursorPos
                        If pt.X >= frmTips.Left / p_TPPX And pt.X <= (frmTips.Left + frmTips.Width) / p_TPPX Then
                            If pt.Y >= frmTips.Top / p_TPPY And pt.Y <= (frmTips.Top + frmTips.Height) / p_TPPY Then
                                ClickedInCell = True
                            End If
                        End If
                    End If
                    
                    If Not (ClickedInCell) Then
                        EndEdit
                        If (m_bInEdit) Then
                            ' We have requested to cancel cancelling the edit.
                            MouseEvent = True
                        Else
                            GetWindowRect Me.hWnd, tR
                            If Not (PtInRect(tR, cursorPos.X, cursorPos.Y) = 0) Then
                                m_iRepostMsg = iMsg
                                LSet m_tRepostPos = cursorPos
                                
                                Dim bShift As Boolean
                                Dim bAlt As Boolean
                                Dim bCtrl As Boolean
                                
                                bShift = (GetAsyncKeyState(vbKeyShift) <> 0)
                                bAlt = (GetAsyncKeyState(vbKeyMenu) <> 0)
                                bCtrl = (GetAsyncKeyState(vbKeyControl) <> 0)
                                m_lRepostShiftState = Abs(bShift * vbShiftMask) Or Abs(bCtrl * vbCtrlMask) Or Abs(bAlt * vbAltMask)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

'#########################################################################################################
'## 功能：  开始编辑
'## 参数：  lRow、lCol: 编辑的行、列
'## 返回：  无
'#########################################################################################################
Public Sub StartEdit(ByVal lRow As Long, ByVal lCol As Long)
    If (mvarEditable) Then
        pCancelEdit False
        If Not (m_bInEdit) Then
            SelectCell , lRow, lCol
            pRequestEdit 0, False
        End If
    End If
End Sub

'#########################################################################################################
'## 功能：  正常结束编辑操作。将触发 PreCancelEdit 事件
'## 参数：  无
'## 返回：  返回选中的单元格数目
'#########################################################################################################
Public Sub EndEdit()
    m_bInEndEditInterlock = True
    If (m_bInEdit) Then
        Dim newValue As Variant
        Dim bStayInEditMode As Boolean
        RaiseEvent PreCancelEdit(mvarCells("K" & m_CellKeyEdit).Row, mvarCells("K" & m_CellKeyEdit).Col, newValue, bStayInEditMode)
        If Not (bStayInEditMode) Then
            CancelEdit
        Else
            If (m_hWndParentForm = 0) Then
                '取消鼠标消息绑定
                AttachMouseHook Me
                m_hWndParentForm = GetParentFormhWnd()
'                AttachMessage Me, m_hWndParentForm, WM_ACTIVATEAPP
            End If
            m_bInEdit = True
        End If
    End If
    m_bInEndEditInterlock = False
End Sub

'#########################################################################################################
'## 功能：  通过hWnd来获取父窗体
'## 返回：  长整型值
'#########################################################################################################
Private Function GetParentFormhWnd() As Long
    Dim lHWnd As Long
    Dim lhWndParent As Long
    
    lHWnd = UserControl.hWnd
    lhWndParent = GetParent(lHWnd)
    Do While Not (lhWndParent = 0) And Not (IsWindowVisible(lhWndParent) = 0)
        lHWnd = lhWndParent
        lhWndParent = GetParent(lHWnd)
    Loop
    GetParentFormhWnd = lHWnd
    
    ' Detect if we're running in the VB IDE - the Message Loop
    ' works in a different way in the IDE compared to as an EXE.
    ' In an EXE, we need to repost end edit mouse events over the
    ' control once it is re-enabled.  In the EXE, we don't.
    ' Bitch!
    Dim sClass As String
    sClass = WindowClassName(lHWnd)
    ' In the IDE, the form's name starts with 'ThunderForm' or 'ThunderMDIForm'.
    ' In EXE, it starts with 'ThunderRT'.  We assume that this message loop
    ' hacking does not occur in other apps, but it may be that it also occurs
    ' in MS Office...
    If InStr(sClass, "ThunderForm") = 1 Or InStr(sClass, "ThunderMDIForm") = 1 Then
        m_bRunningInVBIDE = True
    End If
End Function

'#########################################################################################################
'## 功能：  取消编辑
'#########################################################################################################
Public Sub CancelEdit()
   pCancelEdit False
End Sub

'#########################################################################################################
'## 功能：  取消编辑
'#########################################################################################################
Private Sub pCancelEdit(ByVal bAppDeactivate As Boolean, Optional ByVal bNotAccept As Boolean = False)
    ' 2003-11-24: Otherwise, standard cancel edit mode.
    If (m_bInEdit) Then
        DetachMouseHook Me
'        DetachMessage Me, m_hWndParentForm, WM_ACTIVATEAPP
        m_hWndParentForm = 0
        EnableWindow UserControl.hWnd, 1
        RaiseEvent CancelEdit
        
        If Not (bAppDeactivate) Then
            On Error Resume Next ' Just in case we're not in VB.
            UserControl.Extender.SetFocus
            On Error GoTo 0
        End If
        If mvarInnerEdit Then
            If txtEdit.Text <> txtEdit.Tag And bNotAccept = False Then
                mvarCells("K" & m_CellKeyEdit).Text = txtEdit.Text
                mvarCells("K" & m_CellKeyEdit).Dirty = True
                mvarModified = True
                
                Refresh False, mvarSingleLine = False And mvarAutoHeight, m_CellKeyEdit
            End If
            If txtEdit.Visible Then txtEdit.Visible = False
            txtEdit.Tag = ""
        Else
            txtEdit.Visible = False
            txtEdit.Tag = ""
        End If
        
        If Not (m_bRunningInVBIDE) Then
            '在非VB环境中
            If Not (m_iRepostMsg = 0) Then
                Dim lFlagUp As Long
                Dim lFlagDown As Long
                
                Select Case m_iRepostMsg
                Case WM_LBUTTONDOWN
                    lFlagDown = MOUSEEVENTF_LEFTDOWN
                    lFlagUp = MOUSEEVENTF_LEFTUP
                Case WM_RBUTTONDOWN
                    lFlagDown = MOUSEEVENTF_RIGHTDOWN
                    lFlagUp = MOUSEEVENTF_RIGHTUP
                Case WM_MBUTTONDOWN
                    lFlagDown = MOUSEEVENTF_MIDDLEDOWN
                    lFlagUp = MOUSEEVENTF_MIDDLEUP
                End Select
                mouse_event lFlagDown Or MOUSEEVENTF_ABSOLUTE, 0, 0, 0, 0
                mouse_event lFlagUp Or MOUSEEVENTF_ABSOLUTE, 0, 0, 0, 0
                DoEvents
            End If
        End If
        
        m_iRepostMsg = 0
        m_bInEdit = False
    End If
    Refresh False, False
End Sub

Private Sub pRequestEdit(Optional ByVal iKeyAscii As Integer = 0, _
    Optional ByVal bByMouse As Boolean = False, _
    Optional ByVal bSelStartInEndPos As Boolean = False)
    
    '请求进入编辑
    Dim bCancel As Boolean, lLeft As Long, lTop As Long, lRight As Long, lBottom As Long
    
    If (mvarEnabled) Then
        If (m_CellKeySelected > 0) Then    '选中某个单元格
            If (mvarEditable) Then
                '如果已经在编辑状态中，那么就退出
                If mvarCells("K" & m_CellKeySelected).Protected Then
                    RaiseEvent ModifyProtected(m_CellKeySelected)
                    m_bInEdit = False
                ElseIf Not (m_bInEdit) Then
                    bCancel = False
                    RaiseEvent RequestEdit(mvarCells("K" & m_CellKeySelected).Row, mvarCells("K" & m_CellKeySelected).Col, iKeyAscii, bCancel)
                    m_bInEdit = Not (bCancel)
                    '存储当前请求编辑的单元格
                    If (m_bInEdit) Then
                        '开始编辑
                        m_bEditHeightChanged = False    '记录控件高度是否改变
'                        Debug.Print "m_CellKeyEdit 改变:" & m_CellKeyEdit
                        m_CellKeyEdit = m_CellKeySelected
                        If mvarInnerEdit Then
                            On Error Resume Next
                            PostMessage UserControl.hWnd, WM_LBUTTONUP, 0, 0&
                            ReleaseCapture
                            
                            '绑定消息
                            AttachMouseHook Me
                            m_hWndParentForm = GetParentFormhWnd()
'                            AttachMessage Me, m_hWndParentForm, WM_ACTIVATEAPP
                           
                            mvarCells("K" & m_CellKeyEdit).GetCellTextBorder lLeft, lTop, lRight, lBottom
                            SetEditInfo txtEdit, mvarCells("K" & m_CellKeyEdit)
                            txtEdit.Move (lLeft) * p_TPPX, (lTop) * p_TPPY, Abs(lRight - lLeft) * p_TPPX, Abs(lBottom - lTop) * p_TPPY
                            txtEdit.ZOrder
                            txtEdit.Visible = True
                            DoEvents
'                            If (iKeyAscii <> 0) Then
                                '如果有按键信息
                                If iKeyAscii = vbKeySpace Then
                                    txtEdit.SelStart = 0
                                ElseIf iKeyAscii = vbKeyReturn Or iKeyAscii = vbKeyTab Then
                                    txtEdit.SelStart = 0
                                    txtEdit.SelLength = Len(txtEdit)
                                ElseIf iKeyAscii > 0 Then
                                    txtEdit.Text = Chr$(iKeyAscii) & txtEdit.Text
                                    txtEdit.SelStart = 1
                                Else
                                    txtEdit.SelStart = 0
                                End If
                                txtEdit.SelLength = Len(txtEdit.Text)
'                            End If
                            If bSelStartInEndPos Then txtEdit.SelStart = Len(txtEdit)
                            txtEdit.SetFocus
                            
                            If bByMouse Then    '如果由鼠标激活！
                                '定位关标位置
                                DoEvents
                                SetCapture txtEdit.hWnd
                                If giGetMouseButton = vbLeftButton Then
                                    '左键按下
                                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
                                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
                                Else
                                    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
                                    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
                                End If
                            End If
                        End If
                    End If
                End If
           End If
        End If
    End If
End Sub

Private Sub SetEditInfo(ByRef edtText As TextBox, ByRef oCell As cCell)
    '设置编辑控件格式
    edtText.Font.Name = oCell.FontName
    edtText.Font.Size = oCell.FontSize
    edtText.Font.Bold = oCell.FontBold
    edtText.Font.Strikethrough = oCell.FontStrikeout
    edtText.Font.Underline = oCell.FontUnderline
    edtText.Font.Weight = IIf(oCell.FontBold, FW_BOLD, FW_NORMAL)
    edtText.Font.Italic = oCell.FontItalic
'    edtText.Font.Charset = DEFAULT_CHARSET
    edtText.ForeColor = oCell.ForeColor
    edtText.Alignment = IIf(oCell.HAlignment = HALignLeft, vbLeftJustify, IIf(oCell.HAlignment = HALignRight, vbRightJustify, vbCenter))
    If (mvarAlternateRowBackColor <> -1) And (oCell.Row Mod 2) = 0 Then
        '如果是交替行且具有交替色
        edtText.BackColor = TranslateColor(IIf(oCell.BackColor <> -1, oCell.BackColor, mvarAlternateRowBackColor))
    Else
        edtText.BackColor = TranslateColor(IIf(oCell.BackColor = -1, mvarBackColor, oCell.BackColor))
    End If
    
    edtText.Text = oCell.Text
    edtText.Tag = edtText.Text
End Sub

Private Sub m_tmrHotTrack_ThatTime()
    '绘制热跟踪边框
    If (mvarHotTrack) Then
        If (m_CellKeyHot > 0) Then
            Dim tP As POINTAPI
            Dim tR As RECT
            Dim iGridCol As Long
            GetCursorPos tP                 '获取当前鼠标位置
            ScreenToClient m_hWnd, tP       '获取客户区域
            GetClientRect m_hWnd, tR        '进一步获取客户区域的矩形位置
            If (PtInRect(tR, tP.X, tP.Y) = 0) Then    '如果鼠标不在客户区域以内的话，表示鼠标移出，此时取消Hot
                mvarCells("K" & m_CellKeyHot).Hot = False
                m_CellKeyHot = 0
                m_bDirty = True
                Call Draw
                RaiseEvent HotItemChange(0, 0)
                m_tmrHotTrack.Interval = 0
            End If
        Else
            m_tmrHotTrack.Interval = 0
        End If
    End If
End Sub

'################################################################################################################
'## 控件基本事件处理
'################################################################################################################

Private Sub UserControl_Paint()
'    If mvarRedraw Then Call Draw
End Sub

Private Sub UserControl_Initialize()
    Set mvarCells = New cCells
    
    m_bDirty = True
    mvarEnabled = True
    
   ' Set up information about this control for
   ' IOleInPlaceActiveObject interface:
   Dim IPAO As IOleInPlaceActiveObject

   With m_IPAOHookStruct
      Set IPAO = Me
      CopyMemory .IPAOReal, IPAO, 4
      CopyMemory .TBEx, Me, 4
      .lpVTable = IPAOVTable
      .ThisPointer = VarPtr(m_IPAOHookStruct)
   End With
End Sub

Private Sub UserControl_Resize()
    If UserControl.Ambient.UserMode And m_bInResize = False And mvarRedraw Then
        m_bDirty = True
        Call Draw
    End If
End Sub

Private Sub UserControl_Terminate()
    Set mvarCells = Nothing
    If Not (m_hWnd = 0) Then m_hWnd = 0
    If Not (m_tmrHotTrack Is Nothing) Then
        m_tmrHotTrack.Interval = 0
        Set m_tmrHotTrack = Nothing
    End If
    
    ' Detach the custom IOleInPlaceActiveObject interface
    ' pointers.
    With m_IPAOHookStruct
       CopyMemory .IPAOReal, 0&, 4
       CopyMemory .TBEx, 0&, 4
    End With
    
    '清除内存图片
    If (m_hDC <> 0) Then
        If (m_hBmpOld <> 0) Then
           SelectObject m_hDC, m_hBmpOld
        End If
        If (m_hBmp <> 0) Then
           DeleteObject m_hBmp
        End If
        DeleteDC m_hDC
        m_hDC = 0
    End If
    If Not frmTips Is Nothing Then Unload frmTips
    
    '手动释放内存
    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
End Sub

Private Sub UserControl_InitProperties()
'创建对象新实例时发生，即新属性的最初初始化代码！（即，当用户在窗体上放置一个控件时触发此事件！运行时不再触发！）
    Redraw = True
    SingleLine = False
    Enabled = True
    RowCount = 0
    ColCount = 0
    DefaultRowHeight = 300
    AlternateRowBackColor = -1  '没有交替色
    BackColor = vbWhite
    GridLineColor = vbBlack
    GridLineWidth = 1
    BorderColor = vbBlack
    BorderWidth = 0
    Editable = True
    ForeColor = vbBlack
    BorderColor = vbBlack
    HighlightBackColor = vbHighlight
    HighlightForeColor = vbHighlightText
    HighlightSelectedIcons = False
    HighlightMode = HMFilledRectSolid
    DrawFocusRect = False
    HotTrack = False
    SingleClickEdit = False
    FontQuality = FQProof
    AutoHeight = True
    MinRowHeight = 0
    WordEllipsis = False
    CellMargin = 2
    CellIndent = 0
    InnerEdit = True
    TabKeyMoveNextCell = True
    ShowToolTipText = False
    ExtendTag = ""
    UserTag = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'当加载具有保存状态的对象的旧实例时，发生该事件。
'属性读取（静态属性的读取，从而转化为动态属性，此时调用pInitialise函数初始化句柄！）
    Redraw = PropBag.ReadProperty("Redraw", True)
    SingleLine = PropBag.ReadProperty("SingleLine", True)
    Enabled = PropBag.ReadProperty("Enabled", True)
    RowCount = PropBag.ReadProperty("RowCount", 0)
    ColCount = PropBag.ReadProperty("ColCount", 0)
    DefaultRowHeight = PropBag.ReadProperty("DefaultRowHeight", 300)
    AlternateRowBackColor = PropBag.ReadProperty("AlternateRowBackColor", -1)
    BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    BackgroundPicture = PropBag.ReadProperty("BackgroundPicture", Nothing)
    GridLineColor = PropBag.ReadProperty("GridLineColor", vbBlack)
    GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
    BorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
    BorderWidth = PropBag.ReadProperty("BorderWidth", 0)
    Editable = PropBag.ReadProperty("Editable", True)
    ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    BorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
    HighlightBackColor = PropBag.ReadProperty("HighlightBackColor", vbHighlight)
    HighlightForeColor = PropBag.ReadProperty("HighlightForeColor", vbHighlightText)
    HighlightSelectedIcons = PropBag.ReadProperty("HighlightSelectedIcons", False)
    HighlightMode = PropBag.ReadProperty("HighlightMode", HMFilledRectSolid)
    DrawFocusRect = PropBag.ReadProperty("DrawFocusRect", False)
    HotTrack = PropBag.ReadProperty("HotTrack", False)
    SingleClickEdit = PropBag.ReadProperty("SingleClickEdit", False)
    FontQuality = PropBag.ReadProperty("FontQuality", FQProof)
    AutoHeight = PropBag.ReadProperty("AutoHeight", True)
    MinRowHeight = PropBag.ReadProperty("MinRowHeight", 0)
    WordEllipsis = PropBag.ReadProperty("WordEllipsis", False)
    CellMargin = PropBag.ReadProperty("CellMargin", 2)
    CellIndent = PropBag.ReadProperty("CellIndent", 0)
    InnerEdit = PropBag.ReadProperty("InnerEdit", True)
    TabKeyMoveNextCell = PropBag.ReadProperty("TabKeyMoveNextCell", True)
    ShowToolTipText = PropBag.ReadProperty("ShowToolTipText", False)
    ExtendTag = PropBag.ReadProperty("ExtendTag", "")
    UserTag = PropBag.ReadProperty("UserTag", "")
    
    If Ambient.UserMode Then
        p_TPPX = Screen.TwipsPerPixelX
        p_TPPY = Screen.TwipsPerPixelY
    
        Call BuildMemDC
        
        '子类消息绑定
        m_hWnd = UserControl.hWnd
        Subclass1.hWnd = m_hWnd
        Subclass1.Messages(WM_SETTINGCHANGE) = True
        Subclass1.Messages(WM_DISPLAYCHANGE) = True
        Subclass1.Messages(WM_SETFOCUS) = True
        Subclass1.Messages(WM_ACTIVATEAPP) = True

       ' Set up information about this control for
       ' IOleInPlaceActiveObject interface:
       Dim IPAO As IOleInPlaceActiveObject
    
       With m_IPAOHookStruct
          Set IPAO = Me
          CopyMemory .IPAOReal, IPAO, 4
          CopyMemory .TBEx, Me, 4
          .lpVTable = IPAOVTable
          .ThisPointer = VarPtr(m_IPAOHookStruct)
       End With
            
        Set m_tmrHotTrack = New cTimer
     End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'当保存对象的实例时，发生该事件。该事件通知对象此时需要保存对象的状态，以便将来可恢复该状态。大多数情况下，对象的状态仅包括属性值。
'属性保存（静态属性的保存）
    PropBag.WriteProperty "Redraw", Redraw, True
    PropBag.WriteProperty "SingleLine", SingleLine, True
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "RowCount", RowCount, 0
    PropBag.WriteProperty "ColCount", ColCount, 0
    PropBag.WriteProperty "DefaultRowHeight", DefaultRowHeight, 300
    PropBag.WriteProperty "AlternateRowBackColor", AlternateRowBackColor, -1
    PropBag.WriteProperty "BackColor", BackColor, vbWhite
    PropBag.WriteProperty "BackgroundPicture", BackgroundPicture, Nothing
    PropBag.WriteProperty "GridLineColor", GridLineColor, vbBlack
    PropBag.WriteProperty "GridLineWidth", GridLineWidth, 1
    PropBag.WriteProperty "BorderColor", BorderColor, vbBlack
    PropBag.WriteProperty "BorderWidth", BorderWidth, 0
    PropBag.WriteProperty "Editable", Editable, True
    PropBag.WriteProperty "ForeColor", ForeColor, vbBlack
    PropBag.WriteProperty "BorderColor", BorderColor, vbBlack
    PropBag.WriteProperty "HighlightBackColor", HighlightBackColor, vbHighlight
    PropBag.WriteProperty "HighlightForeColor", HighlightForeColor, vbHighlightText
    PropBag.WriteProperty "HighlightSelectedIcons", HighlightSelectedIcons, False
    PropBag.WriteProperty "HighlightMode", HighlightMode, HMFilledRectSolid
    PropBag.WriteProperty "DrawFocusRect", DrawFocusRect, False
    PropBag.WriteProperty "HotTrack", HotTrack, False
    PropBag.WriteProperty "SingleClickEdit", SingleClickEdit, False
    PropBag.WriteProperty "FontQuality", FontQuality, FQProof
    PropBag.WriteProperty "AutoHeight", AutoHeight, True
    PropBag.WriteProperty "MinRowHeight", MinRowHeight, 0
    PropBag.WriteProperty "WordEllipsis", WordEllipsis, False
    PropBag.WriteProperty "CellMargin", CellMargin, 2
    PropBag.WriteProperty "CellIndent", CellIndent, 0
    PropBag.WriteProperty "InnerEdit", InnerEdit, True
    PropBag.WriteProperty "TabKeyMoveNextCell", TabKeyMoveNextCell, True
    PropBag.WriteProperty "ShowToolTipText", ShowToolTipText, False
    PropBag.WriteProperty "ExtendTag", ExtendTag, ""
    PropBag.WriteProperty "UserTag", UserTag, ""

    PropertyChanged "Redraw"
    PropertyChanged "SingleLine"
    PropertyChanged "Enabled"
    PropertyChanged "RowCount"
    PropertyChanged "ColCount"
    PropertyChanged "DefaultRowHeight"
    PropertyChanged "DefaultColWidth"
    PropertyChanged "AlternateRowBackColor"
    PropertyChanged "BackColor"
    PropertyChanged "GridLineColor"
    PropertyChanged "GridLineWidth"
    PropertyChanged "BorderColor"
    PropertyChanged "BorderWidth"
    PropertyChanged "Editable"
    PropertyChanged "ForeColor"
    PropertyChanged "BorderColor"
    PropertyChanged "HighlightBackColor"
    PropertyChanged "HighlightForeColor"
    PropertyChanged "HighlightSelectedIcons"
    PropertyChanged "HighlightMode"
    PropertyChanged "DrawFocusRect"
    PropertyChanged "HotTrack"
    PropertyChanged "SingleClickEdit"
    PropertyChanged "FontQuality"
    PropertyChanged "AutoHeight"
    PropertyChanged "MinRowHeight"
    PropertyChanged "WordEllipsis"
    PropertyChanged "CellMargin"
    PropertyChanged "CellIndent"
    PropertyChanged "InnerEdit"
    PropertyChanged "TabKeyMoveNextCell"
    PropertyChanged "ShowToolTipText"
    PropertyChanged "ExtendTag"
    PropertyChanged "UserTag"
End Sub

'################################################################################################################
'## 底层消息处理
'################################################################################################################

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, LParam As Long, Result As Long)
    Select Case Msg
    Case WM_SETFOCUS
'        If mvarTabKeyMoveNextCell Then  '捕捉Tab按键
            Dim pOleObject                  As IOleObject
            Dim pOleInPlaceSite             As IOleInPlaceSite
            Dim pOleInPlaceFrame            As IOleInPlaceFrame
            Dim pOleInPlaceUIWindow         As IOleInPlaceUIWindow
            Dim pOleInPlaceActiveObject     As IOleInPlaceActiveObject
            Dim PosRect                     As RECT
            Dim ClipRect                    As RECT
            Dim FrameInfo                   As OLEINPLACEFRAMEINFO
            Dim grfModifiers                As Long
            Dim AcceleratorMsg              As Msg
            
            'Get in-place frame and make sure it is set to our in-between
            'implementation of IOleInPlaceActiveObject in order to catch
            'TranslateAccelerator calls
            Set pOleObject = Me
            Set pOleInPlaceSite = pOleObject.GetClientSite
            pOleInPlaceSite.GetWindowContext pOleInPlaceFrame, pOleInPlaceUIWindow, VarPtr(PosRect), VarPtr(ClipRect), VarPtr(FrameInfo)
            CopyMemory pOleInPlaceActiveObject, m_IPAOHookStruct.ThisPointer, 4
            pOleInPlaceFrame.SetActiveObject pOleInPlaceActiveObject, vbNullString
            If Not pOleInPlaceUIWindow Is Nothing Then
                pOleInPlaceUIWindow.SetActiveObject pOleInPlaceActiveObject, vbNullString
            End If
            ' Clear up the inbetween implementation:
            CopyMemory pOleInPlaceActiveObject, 0&, 4
'        End If
    Case WM_ACTIVATEAPP
        If (wParam = 0) Then
            If (m_bInEdit) Then
                ' Stop editing
                pCancelEdit True
            End If
        End If
    Case WM_SETTINGCHANGE, WM_DISPLAYCHANGE
        m_bTrueColor = (GetDeviceCaps(UserControl.hDC, BITSPIXEL) > 8)
        Call Refresh
    End Select
End Sub

Private Property Get ShiftState() As Integer
    On Error Resume Next
    ShiftState = GetAsyncKeyState(vbKeyShift) * vbShiftMask Or GetAsyncKeyState(vbKeyControl) * vbCtrlMask
End Property

Friend Function TranslateAccelerator(lpMsg As VBOleGuids.Msg) As Long
    TranslateAccelerator = S_FALSE
    If (mvarTabKeyMoveNextCell) Then
        ' Here you can modify the response to the key down
        ' accelerator command using the values in lpMsg.  This
        ' can be used to capture Tabs, Returns, Arrows etc.
        ' Just process the message as required and return S_OK.
        If (lpMsg.wParam And &HFFFF&) = vbKeyTab Then
            Select Case lpMsg.message
            Case WM_KEYDOWN
                UserControl_KeyDown vbKeyTab, ShiftState
                TranslateAccelerator = S_OK
            Case WM_KEYUP
                UserControl_KeyUp vbKeyTab, ShiftState
                TranslateAccelerator = S_OK
            End Select
        End If
    End If
End Function

Public Sub InsertRow(ByVal lRow As Long)
    '插入到指定行之后
    Dim i As Long, lKey As Long
    
    Me.Redraw = False
    mvarRowCount = mvarRowCount + 1
    ReDim Preserve RowInfo(1 To mvarRowCount) As RowInfoType
    
    For i = mvarRowCount To lRow + 2 Step -1
        RowHeight(i) = RowHeight(i - 1)
    Next
    RowHeight(lRow + 1) = mvarMinRowHeight
    
    For i = 1 To mvarCells.Count
        If mvarCells(i).Row > lRow Then
            mvarCells(i).Row = mvarCells(i).Row + 1
        End If
    Next
    For i = 1 To mvarColCount
        lKey = mvarCells.Add(, lRow + 1, i)
    Next
    ReDim RowColInfo(1 To mvarRowCount, 1 To mvarColCount) As Long
    For i = 1 To mvarCells.Count
        RowColInfo(mvarCells(i).Row, mvarCells(i).Col) = mvarCells(i).Key
    Next
    Me.Redraw = True
    Refresh
End Sub

Public Sub DeleteRow(ByVal lRow As Long)
    '删除指定行
    Dim i As Long, lKey As Long, lCol As Long
    Dim lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long, sText As String
    
    Me.Redraw = False
    If m_CellKeySelected > 0 Then lCol = mvarCells("K" & m_CellKeySelected).Col
    
    For i = lRow To mvarRowCount - 1
        RowHeight(i) = RowHeight(i + 1)
    Next
    mvarRowCount = mvarRowCount - 1
    ReDim Preserve RowInfo(1 To mvarRowCount) As RowInfoType
    
    For i = mvarCells.Count To 1 Step -1
        If mvarCells(i).Row = lRow Then
            mvarCells.Remove "K" & mvarCells(i).Key
        End If
    Next
    
    For i = 1 To mvarCells.Count
        If Len(mvarCells(i).MergeInfo) = 16 Then
            sText = mvarCells(i).MergeInfo
            lRow1 = Val(Mid(sText, 1, 4))
            lCol1 = Val(Mid(sText, 5, 4))
            lRow2 = Val(Mid(sText, 9, 4))
            lCol2 = Val(Mid(sText, 13, 4))
            If lRow1 > lRow Then lRow1 = lRow1 - 1
            If lRow2 > lRow Then lRow2 = lRow2 - 1
            sText = Format(lRow1, "0000") & Format(lCol1, "0000") & Format(lRow2, "0000") & Format(lCol2, "0000")
            mvarCells(i).MergeInfo = sText
        End If
        If mvarCells(i).Row > lRow Then
            mvarCells(i).Row = mvarCells(i).Row - 1
        End If
    Next
    
    ReDim RowColInfo(1 To mvarRowCount, 1 To mvarColCount) As Long
    For i = 1 To mvarCells.Count
        RowColInfo(mvarCells(i).Row, mvarCells(i).Col) = mvarCells(i).Key
    Next
    
    If m_CellKeySelected > 0 Then
        m_CellKeySelected = CellKey(IIf(lRow < mvarRowCount, lRow, IIf(lRow - 1 < 1, 1, lRow - 1)), lCol)
        mvarCells("K" & m_CellKeySelected).Selected = True
    End If
    Me.Redraw = True
    Me.Modified = True
    Refresh
End Sub

Public Sub InsertCol(ByVal lCol As Long)
    '插入到指定列之后
    Dim i As Long, lKey As Long
    Me.Redraw = False
    mvarColCount = mvarColCount + 1
    ReDim Preserve ColInfo(1 To mvarColCount) As ColInfoType
    
    For i = mvarColCount To lCol + 2 Step -1
        ColWidth(i) = ColWidth(i - 1)
    Next
    ColWidth(lCol + 1) = 800
    
    For i = 1 To mvarCells.Count
        If mvarCells(i).Col > lCol Then
            mvarCells(i).Col = mvarCells(i).Col + 1
        End If
    Next
    For i = 1 To mvarRowCount
        lKey = mvarCells.Add(, i, lCol + 1)
    Next
    ReDim RowColInfo(1 To mvarRowCount, 1 To mvarColCount) As Long
    For i = 1 To mvarCells.Count
        RowColInfo(mvarCells(i).Row, mvarCells(i).Col) = mvarCells(i).Key
    Next
    Me.Redraw = True
    Refresh
End Sub

Public Sub DeleteCol(ByVal lCol As Long)
    '删除指定列
    Dim i As Long, lKey As Long, lRow As Long
    Dim lRow1 As Long, lCol1 As Long, lRow2 As Long, lCol2 As Long, sText As String
    
    Me.Redraw = False
    If m_CellKeySelected > 0 Then lRow = mvarCells("K" & m_CellKeySelected).Row
    
    For i = lCol To mvarColCount - 1
        ColWidth(i) = ColWidth(i + 1)
    Next
    mvarColCount = mvarColCount - 1
    ReDim Preserve ColInfo(1 To mvarColCount) As ColInfoType
    
    For i = mvarCells.Count To 1 Step -1
        If mvarCells(i).Col = lCol Then
            mvarCells.Remove "K" & mvarCells(i).Key
        End If
    Next
    For i = 1 To mvarCells.Count
        If Len(mvarCells(i).MergeInfo) = 16 Then
            sText = mvarCells(i).MergeInfo
            lRow1 = Val(Mid(sText, 1, 4))
            lCol1 = Val(Mid(sText, 5, 4))
            lRow2 = Val(Mid(sText, 9, 4))
            lCol2 = Val(Mid(sText, 13, 4))
            If lCol1 > lCol Then lCol1 = lCol1 - 1
            If lCol2 > lCol Then lCol2 = lCol2 - 1
            sText = Format(lRow1, "0000") & Format(lCol1, "0000") & Format(lRow2, "0000") & Format(lCol2, "0000")
            mvarCells(i).MergeInfo = sText
        End If
        If mvarCells(i).Col > lCol Then
            mvarCells(i).Col = mvarCells(i).Col - 1
        End If
    Next
    
    ReDim RowColInfo(1 To mvarRowCount, 1 To mvarColCount) As Long
    For i = 1 To mvarCells.Count
        RowColInfo(mvarCells(i).Row, mvarCells(i).Col) = mvarCells(i).Key
    Next
    
    If m_CellKeySelected > 0 Then
        m_CellKeySelected = CellKey(lRow, IIf(lCol < mvarColCount - 1, lCol, IIf(lCol - 1 < 1, 1, lCol - 1)))
        mvarCells("K" & m_CellKeySelected).Selected = True
    End If
    Me.Redraw = True
    Me.Modified = True
    Refresh
End Sub

Public Sub Init(ByVal Rows As Long, ByVal Cols As Long)
    '初始化表格
    Dim i As Long, j As Long, lKey As Long
    mvarRowCount = Rows
    mvarColCount = Cols
    Set mvarCells = New cCells
    ReDim RowColInfo(1 To Rows, 1 To Cols) As Long
    ReDim ColInfo(1 To Cols) As ColInfoType
    ReDim RowInfo(1 To Rows) As RowInfoType
    m_DefaultWidth = UserControl.Width / Cols
    m_DefaultHeight = UserControl.Height / Rows
    mvarExtendTag = ""
    mvarUserTag = ""
    mvarModified = False
    
    For i = 1 To Cols
        ColInfo(i).ColWidth = m_DefaultWidth
    Next
    
    For i = 1 To Rows
        RowInfo(i).RowHeight = m_DefaultHeight
        For j = 1 To Cols
            lKey = mvarCells.Add(, i, j)
            RowColInfo(i, j) = lKey         '存储行列与单元格关键字的对应关系！～
        Next j
    Next i
End Sub

Public Sub AppendRow()
    '添加一行到末尾
    Dim i As Long, j As Long, lKey As Long
    mvarRowCount = mvarRowCount + 1
    ReDim Preserve RowInfo(1 To mvarRowCount) As RowInfoType
    mvarModified = True
    
    i = mvarRowCount
    RowInfo(i).RowHeight = RowInfo(i - 1).RowHeight
    For j = 1 To mvarColCount
        lKey = mvarCells.Add(, i, j)
    Next j
    ReDim RowColInfo(1 To mvarRowCount, 1 To mvarColCount) As Long
    For i = 1 To mvarCells.Count
        RowColInfo(mvarCells(i).Row, mvarCells(i).Col) = mvarCells(i).Key
    Next
End Sub


