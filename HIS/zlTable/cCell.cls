VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'################################################################################################################
'## 枚举常量
'################################################################################################################

'水平对齐方式
Public Enum HAlignEnum
    HALignLeft = DT_LEFT                '靠左
    HALignCentre = DT_CENTER            '居中
    HALignRight = DT_RIGHT              '靠右
End Enum

'垂直对齐方式
Public Enum VAlignEnum
    VALignTop = DT_TOP                  '靠上
    VALignVCentre = DT_VCENTER          '垂直居中
    VALignBottom = DT_BOTTOM            '靠下
End Enum

'################################################################################################################
'## 内部变量
'################################################################################################################

'格式
Private mvarKey As String               '关键字
Private mvarRow As Long                 '行
Private mvarCol As Long                 '列
Private mvarMargin As Long              '单元格边距，默认为30
Private mvarSingleLine As Boolean       '是否单行显示，默认为False
Private mvarForeColor As OLE_COLOR      '前景色，默认为黑色
Private mvarBackColor As OLE_COLOR      '背景色，默认为白色
Private mvarGridLineColor As OLE_COLOR  '边框颜色，默认为黑色
Private mvarGridLineWidth As Long       '边框宽度，默认为1
Private mvarFixedWidth As Boolean       '是否锁定宽度，默认为False
Private mvarAutoHeight As Boolean       '自动高度，默认为True
Private mvarFontName As String          '字体名称
Private mvarFontSize As Long            '字体大小，单位：磅值
Private mvarFontBold As Boolean         '粗体
Private mvarFontItalic As Boolean       '斜体
Private mvarFontStrikeout As Boolean    '删除线
Private mvarFontUnderline As Boolean    '下划线
Private mvarFontWeight As Long          '字体重量
Private mvarFormatString As String      '文本格式
Private mvarIndent As Long              '缩进值
Private mvarHAlignment As HAlignEnum    '水平对齐方式
Private mvarVAlignment As VAlignEnum    '垂直对齐方式
Private mvarProtected As Boolean        '是否保护
Private mvarToolTipText As String       '提示文本
Private mvarPicture As StdPicture       '单元格图片

'数据
Private mvarText As String              '文本内容，默认为空
Private mvarTag As String               '扩展标记
Private mvarIcon As Long                '图标索引值，默认为-1，表示没有图标

'其他属性
Private mvarMergeInfo As String         '合并字符串：分别由"行1:列1:行2:列2"组成，共16位；为空表示没有合并
Private mvarSelected As Boolean         '是否选中
Private mvarHot As Boolean              '是否热跟踪
Private mvarVisibled As Boolean         '是否可见
Private mvarWidth As Long               '宽度
Private mvarHeight As Long              '高度
Private mvarDirty As Boolean            '是否已经被修改过，需要重绘

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let Row(ByVal vData As Long)
    mvarRow = vData
End Property

Public Property Get Row() As Long
    Row = mvarRow
End Property

Public Property Let Col(ByVal vData As Long)
    mvarCol = vData
End Property

Public Property Get Col() As Long
    Col = mvarCol
End Property

Public Property Let Margin(ByVal vData As Long)
    mvarMargin = vData
    mvarDirty = True
End Property

Public Property Get Margin() As Long
    Margin = mvarMargin
End Property

Public Property Let SingleLine(ByVal vData As Boolean)
    mvarSingleLine = vData
    mvarDirty = True
End Property

Public Property Get SingleLine() As Boolean
    SingleLine = mvarSingleLine
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
    mvarDirty = True
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
    mvarDirty = True
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let GridLineColor(ByVal vData As OLE_COLOR)
    mvarGridLineColor = vData
    mvarDirty = True
End Property

Public Property Get GridLineColor() As OLE_COLOR
    GridLineColor = mvarGridLineColor
End Property

Public Property Let GridLineWidth(ByVal vData As Long)
    mvarGridLineWidth = vData
    mvarDirty = True
End Property

Public Property Get GridLineWidth() As Long
    GridLineWidth = mvarGridLineWidth
End Property

Public Property Let FixedWidth(ByVal vData As Boolean)
    mvarFixedWidth = vData
    mvarDirty = True
End Property

Public Property Get FixedWidth() As Boolean
    FixedWidth = mvarFixedWidth
End Property

Public Property Let AutoHeight(ByVal vData As Boolean)
    mvarAutoHeight = vData
    mvarDirty = True
End Property

Public Property Get AutoHeight() As Boolean
    AutoHeight = mvarAutoHeight
End Property

Public Property Let FontName(ByVal vData As String)
    mvarFontName = vData
    mvarDirty = True
End Property

Public Property Get FontName() As String
    FontName = mvarFontName
End Property

Public Property Let FontSize(ByVal vData As Long)
    mvarFontSize = vData
    mvarDirty = True
End Property

Public Property Get FontSize() As Long
    FontSize = mvarFontSize
End Property

Public Property Let FontBold(ByVal vData As Boolean)
    mvarFontBold = vData
    mvarDirty = True
End Property

Public Property Get FontBold() As Boolean
    FontBold = mvarFontBold
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
    mvarFontItalic = vData
    mvarDirty = True
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mvarFontItalic
End Property

Public Property Let FontStrikeout(ByVal vData As Boolean)
    mvarFontStrikeout = vData
    mvarDirty = True
End Property

Public Property Get FontStrikeout() As Boolean
    FontStrikeout = mvarFontStrikeout
End Property

Public Property Let FontUnderline(ByVal vData As Boolean)
    mvarFontUnderline = vData
    mvarDirty = True
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = mvarFontUnderline
End Property

Public Property Let FontWeight(ByVal vData As Long)
    mvarFontWeight = vData
    mvarDirty = True
End Property

Public Property Get FontWeight() As Long
    FontWeight = mvarFontWeight
End Property

Public Property Let FormatString(ByVal vData As String)
    mvarFormatString = vData
    mvarDirty = True
End Property

Public Property Get FormatString() As String
    FormatString = mvarFormatString
End Property

Public Property Let Indent(ByVal vData As Long)
    mvarIndent = vData
    mvarDirty = True
End Property

Public Property Get Indent() As Long
    Indent = mvarIndent
End Property

Public Property Let HAlignment(ByVal vData As HAlignEnum)
    mvarHAlignment = vData
    mvarDirty = True
End Property

Public Property Get HAlignment() As HAlignEnum
    HAlignment = mvarHAlignment
End Property

Public Property Let VAlignment(ByVal vData As VAlignEnum)
    mvarVAlignment = vData
    mvarDirty = True
End Property

Public Property Get VAlignment() As VAlignEnum
    VAlignment = mvarVAlignment
End Property

Public Property Let Protected(ByVal vData As Boolean)
    mvarProtected = vData
    mvarDirty = True
End Property

Public Property Get Protected() As Boolean
    Protected = mvarProtected
End Property

Public Property Let ToolTipText(ByVal vData As String)
    mvarToolTipText = vData
    mvarDirty = True
End Property

Public Property Get ToolTipText() As String
    ToolTipText = mvarToolTipText
End Property

Public Property Set Picture(ByVal vData As StdPicture)
    Set mvarPicture = vData
    mvarDirty = True
End Property

Public Property Let Picture(ByVal vData As StdPicture)
    Set mvarPicture = vData
    mvarDirty = True
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mvarPicture
End Property

Public Property Let Text(ByVal vData As String)
    mvarText = vData
    mvarDirty = True
End Property

Public Property Get Text() As String
    Text = mvarText
End Property

Public Property Let Tag(ByVal vData As String)
    mvarTag = vData
    mvarDirty = True
End Property

Public Property Get Tag() As String
    Tag = mvarTag
End Property

Public Property Let Icon(ByVal vData As Long)
    mvarIcon = vData
    mvarDirty = True
End Property

Public Property Get Icon() As Long
    Icon = mvarIcon
End Property

Public Property Let MergeInfo(ByVal vData As String)
    mvarMergeInfo = vData
    mvarDirty = True
End Property

Public Property Get MergeInfo() As String
    MergeInfo = mvarMergeInfo
End Property

Public Property Let Selected(ByVal vData As Boolean)
    mvarSelected = vData
    mvarDirty = True
End Property

Public Property Get Selected() As Boolean
    Selected = mvarSelected
End Property

Public Property Let Hot(ByVal vData As Boolean)
    mvarHot = vData
    mvarDirty = True
End Property

Public Property Get Hot() As Boolean
    Hot = mvarHot
End Property

Public Property Let Visibled(ByVal vData As Boolean)
    mvarVisibled = vData
    mvarDirty = True
End Property

Public Property Get Visibled() As Boolean
    Visibled = mvarVisibled
End Property

Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
    mvarDirty = True
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
    mvarDirty = True
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let Dirty(ByVal vData As Boolean)
    mvarDirty = vData
End Property

Public Property Get Dirty() As Boolean
    Dirty = mvarDirty
End Property

Public Property Get MergeStartRow() As Long
    If mvarMergeInfo <> "" And Len(mvarMergeInfo) = 16 Then
        MergeStartRow = Val(Mid(mvarMergeInfo, 1, 4))
    End If
End Property

Public Property Get MergeStartCol() As Long
    If mvarMergeInfo <> "" And Len(mvarMergeInfo) = 16 Then
        MergeStartCol = Val(Mid(mvarMergeInfo, 5, 4))
    End If
End Property

Public Property Get MergeEndRow() As Long
    If mvarMergeInfo <> "" And Len(mvarMergeInfo) = 16 Then
        MergeEndRow = Val(Mid(mvarMergeInfo, 9, 4))
    End If
End Property

Public Property Get MergeEndCol() As Long
    If mvarMergeInfo <> "" And Len(mvarMergeInfo) = 16 Then
        MergeEndCol = Val(Mid(mvarMergeInfo, 13, 4))
    End If
End Property

'#########################################################################################################
'## 功能：  对于多行文本，计算文本高度
'## 返回：  返回高度整型值
'#########################################################################################################
Public Property Get EvaluateTextHeight(ByRef oTable As Table, _
    Optional ByVal sNewText As String = "", _
    Optional ByVal bCalculPicHeight As Boolean = True) As Long
    
    Dim tLF As LOGFONT
    Dim hFnt As Long
    Dim hFntOld As Long
    Dim tR As RECT
    Dim sFont As String
    Dim iChar As Integer
    Dim temp() As Byte
    Dim sText As String
    Dim lW As Long, lH As Long
    
    With tLF
        sFont = mvarFontName & vbNullString
        temp = StrConv(sFont, vbFromUnicode)
        For iChar = 0 To UBound(temp)
            .lfFaceName(iChar) = temp(iChar)
        Next iChar
        .lfItalic = mvarFontItalic
        .lfHeight = GetPixcelHeightByPoint(oTable.hDC, mvarFontSize)
        .lfUnderline = mvarFontUnderline
        .lfWeight = IIf(mvarFontBold, FW_BOLD, FW_NORMAL)
        .lfStrikeOut = mvarFontStrikeout
        .lfQuality = IIf(oTable.FontQuality < 0, 0, oTable.FontQuality)
        .lfCharSet = DEFAULT_CHARSET
    End With
    
    tR.Right = (mvarWidth - 2 * mvarMargin * p_TPPX - mvarGridLineWidth * p_TPPX - IIf(mvarIcon > 0, p_lIconWidth + 30, 0)) / p_TPPX
    
    sText = IIf(mvarFormatString = "", IIf(sNewText = "", mvarText, sNewText), Format(IIf(sNewText = "", mvarText, sNewText), mvarFormatString))
    If sText = "" Then
        If p_lIconHeight > 0 Then EvaluateTextHeight = p_lIconHeight + mvarMargin * p_TPPY * 2 + mvarGridLineWidth * p_TPPX
        If oTable.MinRowHeight > EvaluateTextHeight Then EvaluateTextHeight = oTable.MinRowHeight
        GoTo LL
    End If
    
    hFnt = CreateFontIndirect(tLF) '创建字体
    '选入指定字体，保存旧字体
    hFntOld = SelectObject(oTable.hDC, hFnt)
    
    '绘制文本（单行文本非自动宽度时超宽显示省略号；多行文本非自动高度时超宽显示省略号）
    DrawText oTable.hDC, sText, -1, tR, DT_NOPREFIX Or DT_EXPANDTABS Or DT_EDITCONTROL Or mvarVAlignment Or mvarHAlignment Or _
        IIf(mvarSingleLine, DT_SINGLELINE, DT_WORDBREAK) Or _
        IIf(oTable.WordEllipsis, DT_WORD_ELLIPSIS, 0) Or DT_CALCRECT

    EvaluateTextHeight = (tR.Bottom - tR.Top) * Screen.TwipsPerPixelY + mvarMargin * p_TPPY * 2 + 2 * mvarGridLineWidth * p_TPPX
    If EvaluateTextHeight < p_lIconHeight Then
        EvaluateTextHeight = p_lIconHeight + mvarMargin * p_TPPY * 2 + mvarGridLineWidth * p_TPPX
    End If
    If EvaluateTextHeight < oTable.MinRowHeight Then
        EvaluateTextHeight = oTable.MinRowHeight
    End If
   
    '恢复旧字体
    SelectObject oTable.hDC, hFntOld
    '释放临时字体
    DeleteObject hFnt
    
LL:
    If Not mvarPicture Is Nothing And bCalculPicHeight Then
        '图片单元格，高度就是图片缩放高度
        lW = (mvarWidth - 2 * mvarMargin * p_TPPX - mvarGridLineWidth * p_TPPX - IIf(mvarIcon > 0, p_lIconWidth + 30, 0))
        lH = (lW * mvarPicture.Height) / mvarPicture.Width + mvarMargin * p_TPPY * 2 + 2 * mvarGridLineWidth * p_TPPX
        If EvaluateTextHeight < lH Then EvaluateTextHeight = lH
    End If
End Property

'################################################################################################################
'## 功能：  复制单元格对象
'##
'## 参数：  blnWithKey  :是否连同Key值一起复制！
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cCell
    Dim cNode As New cCell
    If blnWithKey Then cNode.Key = Me.Key
    cNode.Row = Me.Row
    cNode.Col = Me.Col
    cNode.Margin = Me.Margin
    cNode.SingleLine = Me.SingleLine
    cNode.MergeInfo = Me.MergeInfo
    cNode.Selected = Me.Selected
    cNode.Hot = Me.Hot
    cNode.Visibled = Me.Visibled
    cNode.Width = Me.Width
    cNode.Height = Me.Height
    cNode.FixedWidth = Me.FixedWidth
    cNode.AutoHeight = Me.AutoHeight
    cNode.Icon = Me.Icon
    cNode.Text = Me.Text
    cNode.Tag = Me.Tag
    cNode.FormatString = Me.FormatString
    cNode.Indent = Me.Indent
    cNode.HAlignment = Me.HAlignment
    cNode.VAlignment = Me.VAlignment
    cNode.ForeColor = Me.ForeColor
    cNode.BackColor = Me.BackColor
    cNode.GridLineColor = Me.GridLineColor
    cNode.GridLineWidth = Me.GridLineWidth
    cNode.FontName = Me.FontName
    cNode.FontSize = Me.FontSize
    cNode.FontBold = Me.FontBold
    cNode.FontItalic = Me.FontItalic
    cNode.FontStrikeout = Me.FontStrikeout
    cNode.FontUnderline = Me.FontUnderline
    cNode.FontWeight = Me.FontWeight
    cNode.Protected = Me.Protected
    cNode.ToolTipText = Me.ToolTipText
    cNode.Dirty = Me.Dirty
    cNode.Picture = Me.Picture
    
    Set Clone = cNode
End Function

Public Sub GetCellTextBorder(ByRef lLeft As Long, ByRef lTop As Long, ByRef lRight As Long, ByRef lBottom As Long)
    lLeft = (ColInfo(mvarCol).LeftX + mvarGridLineWidth * p_TPPX + mvarMargin * p_TPPX + IIf(mvarIcon > 0, p_lIconWidth + 30, 0)) / p_TPPX
    lTop = (RowInfo(mvarRow).TopY + mvarGridLineWidth * p_TPPY + mvarMargin * p_TPPY) / p_TPPY
    lRight = (ColInfo(mvarCol).LeftX + mvarWidth - mvarMargin * p_TPPX) / p_TPPX
    lBottom = (RowInfo(mvarRow).TopY + mvarHeight - mvarMargin * p_TPPY) / p_TPPY
End Sub

Public Sub GetCellPictureBorder(ByRef lLeft As Long, ByRef lTop As Long, ByRef lWidth As Long, ByRef lHeight As Long)
    lLeft = (ColInfo(mvarCol).LeftX)
    lTop = (RowInfo(mvarRow).TopY)
    Dim lW As Long, lH As Long
    lW = mvarWidth + mvarGridLineWidth * p_TPPX - 1 * p_TPPX
    lH = (lW * mvarPicture.Height) / mvarPicture.Width
    If mvarAutoHeight = False Then
        '如果非自动换行模式，图片拉伸
        lH = mvarHeight / p_TPPY + mvarGridLineWidth - 1
    End If
    lWidth = lW
    lHeight = lH
End Sub

'################################################################################################################
'## 功能：  绘制单元格
'##
'## 参数：  oTable:     表格对象
'##         lhDC:       绘制到指定的DC，如果省略，则使用Table的DC
'################################################################################################################
Public Sub DrawCell(ByRef oTable As Table, Optional ByRef lhDC As Long = 0)
    Dim tLF As LOGFONT
    Dim hFnt As Long
    Dim hBr As Long
    Dim hBr2 As Long
    Dim hFntOld As Long
    Dim tR As RECT
    Dim sFont As String
    Dim iChar As Integer
    Dim temp() As Byte
    Dim sText As String
    Dim tRCopy As RECT
    Dim hPen As Long
    Dim hPenOld As Long
    Dim tJ As POINTAPI
    Dim lLeft As Long, lTop As Long, lRight As Long, lBottom As Long
    Dim i As Long, j As Long
    Dim NoFocusHighlightBackColor As OLE_COLOR, NoFocusHighlightForeColor As OLE_COLOR
    Dim bExportMode As Boolean, LL As Long, lT As Long, lW As Long, lH As Long
    
    If mvarVisibled = False Then Exit Sub
    
    If lhDC = 0 Then
        bExportMode = False
        lhDC = oTable.hDC
    Else
        bExportMode = True
    End If
    
    On Error Resume Next
    tR.Left = (ColInfo(mvarCol).LeftX) / p_TPPX
    tR.Top = (RowInfo(mvarRow).TopY) / p_TPPY
    tR.Right = (ColInfo(mvarCol).LeftX + mvarWidth) / p_TPPX + mvarGridLineWidth - 1
    tR.Bottom = (RowInfo(mvarRow).TopY + mvarHeight) / p_TPPY + mvarGridLineWidth - 1
    
    If oTable.HighlightMode = HMFilledRectSolid Then
        If oTable.AlternateRowBackColor <> -1 And (mvarRow Mod 2) = 0 And mvarBackColor = -1 Then
            NoFocusHighlightBackColor = BlendColor(oTable.HighlightBackColor, oTable.AlternateRowBackColor, 92)
        Else
            NoFocusHighlightBackColor = BlendColor(oTable.HighlightBackColor, IIf(mvarBackColor = -1, oTable.BackColor, mvarBackColor), 92)
        End If
    Else
        NoFocusHighlightBackColor = BlendColor(oTable.HighlightBackColor, vbWhite, 120)
    End If
    NoFocusHighlightForeColor = mvarForeColor
    
    If Not mvarPicture Is Nothing Then
        '绘制图片单元格
        LL = tR.Left + 1
        lT = tR.Top + 1
        lW = mvarWidth / p_TPPX + mvarGridLineWidth - 1
        lH = (lW * mvarPicture.Height) / mvarPicture.Width
        If oTable.AutoHeight = False Then
'            '如果非自动换行模式，那么不允许图片高度超出单元格高度
'            If lH > mvarHeight / p_TPPY Then lH = mvarHeight / p_TPPY
            '如果非自动换行模式，图片拉伸
            lH = mvarHeight / p_TPPY + mvarGridLineWidth - 1
        End If
        frmProperty.picTMP.Cls
        frmProperty.picTMP.Width = lW * p_TPPX
        frmProperty.picTMP.Height = lH * p_TPPY
        frmProperty.picTMP.PaintPicture mvarPicture, 0, 0, frmProperty.picTMP.Width, frmProperty.picTMP.Height
        
        '绘制半透明图片
        If mvarSelected Then
            Dim lBlend As Long
            Dim bf As BLENDFUNCTION
            
            ' Draw the first picture:
            bf.BlendOp = AC_SRC_OVER
            bf.BlendFlags = 0
            bf.SourceConstantAlpha = 255
            bf.AlphaFormat = 0
            CopyMemory lBlend, bf, 4
            frmProperty.picBlank.Cls
            frmProperty.picBlank.BackColor = oTable.HighlightBackColor
'            BitBlt frmProperty.picBlank.hDC, 0, 0, lW, lH, frmProperty.picTMP.hDC, 0, 0, vbSrcCopy
            
'            SetStretchBltMode frmProperty.picTMP.hDC, HALFTONE
'            StretchBlt frmProperty.picTMP.hDC, tR.left, tR.top, lW, lH, frmProperty.picBlank.hDC, 0, 0, frmProperty.picBlank.Width, frmProperty.picBlank.Height, SRCCOPY
            
            bf.SourceConstantAlpha = 100
            CopyMemory lBlend, bf, 4
            AlphaBlend frmProperty.picTMP.hDC, 0, 0, _
                frmProperty.picTMP.Width \ Screen.TwipsPerPixelX, _
                frmProperty.picTMP.Height \ Screen.TwipsPerPixelY, _
                frmProperty.picBlank.hDC, 0, 0, _
                frmProperty.picBlank.ScaleWidth \ Screen.TwipsPerPixelX, _
                frmProperty.picBlank.ScaleHeight \ Screen.TwipsPerPixelY, _
                lBlend
            frmProperty.picTMP.Refresh
        End If
        BitBlt lhDC, tR.Left, tR.Top, lW, lH, frmProperty.picTMP.hDC, 0, 0, vbSrcCopy
    Else
        '填充背景
        If (mvarBackColor <> -1 Or ((oTable.AlternateRowBackColor <> -1) And (mvarRow Mod 2) = 0)) Then
            If (oTable.AlternateRowBackColor <> -1) And (mvarRow Mod 2) = 0 Then
                '如果是交替行且具有交替色
                hBr = CreateSolidBrush(TranslateColor(IIf(mvarBackColor <> -1, mvarBackColor, oTable.AlternateRowBackColor))) '创建交替色的纯色画刷
            Else
                hBr = CreateSolidBrush(TranslateColor(mvarBackColor))      '创建背景色的纯色画刷
            End If
            FillRect lhDC, tR, hBr
            DeleteObject hBr
            hBr = 0
        End If
    End If
    
    '绘制高亮背景
    If mvarSelected And oTable.InEdit = False And (mvarPicture Is Nothing) Then
        Select Case oTable.HighlightMode
        Case HMFilledRectSolid, HMFilledRectAlpha
            If oTable.HighlightMode = HMFilledRectSolid Then
                '实心画刷
                hBr = CreateSolidBrush(TranslateColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor)))
                FillRect lhDC, tR, hBr
                DeleteObject hBr
                hBr = 0
            Else
                '实心半透明画刷
                If (oTable.AlternateRowBackColor <> -1) And (mvarRow Mod 2) = 0 Then
                    '如果是交替行且具有交替色
                    hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), IIf(mvarBackColor <> -1, mvarBackColor, oTable.AlternateRowBackColor), 92))
                Else
                    If mvarBackColor = -1 Then
                        hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), oTable.BackColor, 92))
                    Else
                        hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), mvarBackColor, 92))
                    End If
                End If
                FillRect lhDC, tR, hBr
                DeleteObject hBr
                hBr = 0
            End If
        Case HMOnlyBorderRectAlpha, HMOnlyBorderRectSolid
            If oTable.HighlightMode = HMOnlyBorderRectSolid Then
                '实心画刷（仅仅绘制边框）
                hBr = CreateSolidBrush(TranslateColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor)))   '用高亮背景色创建纯色画刷
                LSet tRCopy = tR                        '赋值矩形
                tRCopy.Right = tRCopy.Right + 1
                tRCopy.Bottom = tRCopy.Bottom + 1
                InflateRect tRCopy, -1, -1              '将矩形边界缩小1个单位
                FrameRect lhDC, tRCopy, hBr       '继续绘制边框，宽度为1
                DeleteObject hBr                        '释放资源
                hBr = 0
            Else
                '实心半透明画刷（仅仅绘制边框）
                If (oTable.AlternateRowBackColor <> -1) And (mvarRow Mod 2) = 0 Then
                    '如果是交替行且具有交替色
                    hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), IIf(mvarBackColor <> -1, mvarBackColor, oTable.AlternateRowBackColor), 92))
                Else
                    If mvarBackColor = -1 Then
                        hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), oTable.BackColor, 92))
                    Else
                        hBr = CreateSolidBrush(BlendColor(IIf(oTable.InFocus, oTable.HighlightBackColor, NoFocusHighlightBackColor), mvarBackColor, 92))
                    End If
                End If
                LSet tRCopy = tR                        '赋值矩形
                tRCopy.Right = tRCopy.Right + 1
                tRCopy.Bottom = tRCopy.Bottom + 1
                InflateRect tRCopy, -1, -1              '将矩形边界缩小1个单位
                FrameRect lhDC, tRCopy, hBr       '继续绘制边框，宽度为1
                DeleteObject hBr                        '释放资源
                hBr = 0
            End If
        End Select
        If oTable.DrawFocusRect Then
            '否则，设置焦点边框
            SetTextColor lhDC, TranslateColor(vbBlack)    '设置字体颜色
            LSet tRCopy = tR
            tRCopy.Left = tRCopy.Left + 1
            tRCopy.Top = tRCopy.Top + 1
            InflateRect tRCopy, -1, -1
            DrawFocusRect lhDC, tRCopy
        End If
    End If
    
    '绘制边框
    hPen = CreatePen(PS_SOLID, 1, mvarGridLineColor)                '设置边框颜色画笔
    hPenOld = SelectObject(lhDC, hPen)                        '选入画笔，保存旧画笔
    For i = 0 To mvarGridLineWidth - 1
        MoveToEx lhDC, tR.Left + i, tR.Bottom - i, tJ
        LineTo lhDC, tR.Right - i, tR.Bottom - i
        LineTo lhDC, tR.Right - i, tR.Top + i
        LineTo lhDC, tR.Left + i, tR.Top + i
        LineTo lhDC, tR.Left + i, tR.Bottom - i
    Next
    SelectObject lhDC, hPenOld
    DeleteObject hPen
    hPen = 0
    
'    Debug.Assert Me.Key <> 7
    GetCellTextBorder lLeft, lTop, lRight, lBottom
    tR.Left = lLeft
    tR.Top = lTop
    tR.Right = lRight
    tR.Bottom = lBottom
    '如果设置了垂直对其方式，那么要修正文本高度（因为垂直对其方式只对单行文本有效，要应用到多行文本，需要手工修正）
    If mvarSingleLine = False Then
        If mvarPicture Is Nothing Then
            '纯文本单元格
            If mvarVAlignment = VALignVCentre Then
                tR.Top = lTop + (lBottom - lTop - (Me.EvaluateTextHeight(oTable) - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY) / 2
            ElseIf mvarVAlignment = VALignBottom Then
                tR.Top = lBottom - (Me.EvaluateTextHeight(oTable) - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY
            End If
        Else
            '包含图片的单元格
            Dim lHT As Long
            lHT = Me.EvaluateTextHeight(oTable, , False)    '纯文本高度
            lH = Me.EvaluateTextHeight(oTable)              '考虑图片的高度
            If mvarVAlignment = VALignVCentre Then
                If lHT < lH Then
                    tR.Top = lTop + (lBottom - lTop - (Me.EvaluateTextHeight(oTable) - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY) / 2 + (lH - lHT) / (2 * p_TPPY)
                Else
                    tR.Top = lTop + (lBottom - lTop - (Me.EvaluateTextHeight(oTable) - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY) / 2
                End If
            ElseIf mvarVAlignment = VALignBottom Then
                If lHT < lH Then
                    tR.Top = lBottom - (lH - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY + (lH - lHT) / p_TPPY
                Else
                    tR.Top = lBottom - (lH - mvarMargin * p_TPPY * 2 - 2 * mvarGridLineWidth * p_TPPX) / p_TPPY
                End If
            End If
        End If
    End If
    
    '绘制图标
    If mvarIcon > 0 Then
        DrawImageIcon oTable.PtrImageList, mvarIcon, lhDC, tR.Left - p_lIconWidth / p_TPPX, tR.Top, p_lIconWidth / p_TPPX, _
        p_lIconHeight / p_TPPY, mvarSelected And oTable.HighlightSelectedIcons, Not oTable.Enabled
    End If
    
    '绘制文本
    sText = IIf(mvarFormatString = "", mvarText, Format(mvarText, mvarFormatString))
    With tLF
        sFont = mvarFontName & vbNullString
        temp = StrConv(sFont, vbFromUnicode)
        For iChar = 0 To UBound(temp)
            .lfFaceName(iChar) = temp(iChar)
        Next iChar
        .lfItalic = mvarFontItalic
        .lfHeight = GetPixcelHeightByPoint(lhDC, mvarFontSize)
        .lfUnderline = mvarFontUnderline
        .lfWeight = IIf(mvarFontBold, FW_BOLD, FW_NORMAL)
        .lfStrikeOut = mvarFontStrikeout
        .lfQuality = IIf(oTable.FontQuality < 0, 0, oTable.FontQuality)
        .lfCharSet = DEFAULT_CHARSET
    End With
    hFnt = CreateFontIndirect(tLF) '创建字体
    '设置字体颜色
    If oTable.Enabled Then
        If mvarSelected = False Or oTable.HighlightMode <> HMFilledRectSolid Or (Not mvarPicture Is Nothing) Then
            SetTextColor lhDC, TranslateColor(mvarForeColor)
        Else
            SetTextColor lhDC, TranslateColor(IIf(oTable.InFocus, oTable.HighlightForeColor, NoFocusHighlightForeColor))
        End If
    Else
        SetTextColor lhDC, TranslateColor(vbGrayText)
    End If
    '选入指定字体，保存旧字体
    hFntOld = SelectObject(lhDC, hFnt)
    '绘制文本（单行文本非自动宽度时超宽显示省略号；多行文本非自动高度时超宽显示省略号）
'    tR.bottom = tR.bottom + 9999
    
    DrawText lhDC, sText, -1, tR, DT_NOPREFIX Or DT_EXPANDTABS Or DT_EDITCONTROL Or mvarVAlignment Or mvarHAlignment Or _
        IIf(mvarSingleLine, DT_SINGLELINE, DT_WORDBREAK) Or _
        IIf(oTable.WordEllipsis, DT_WORD_ELLIPSIS, 0)
        
    '恢复旧字体
    SelectObject lhDC, hFntOld
    '释放临时字体
    DeleteObject hFnt
    
    tR.Left = (ColInfo(mvarCol).LeftX) / p_TPPX + mvarGridLineWidth - 1
    tR.Top = (RowInfo(mvarRow).TopY) / p_TPPY + mvarGridLineWidth - 1
    tR.Right = (ColInfo(mvarCol).LeftX + mvarWidth) / p_TPPX
    tR.Bottom = (RowInfo(mvarRow).TopY + mvarHeight) / p_TPPY
    
    '热跟踪
    If oTable.Enabled And (mvarHot Or mvarSelected) Then
        hBr2 = CreateSolidBrush(TranslateColor(oTable.HighlightBackColor))   '用高亮背景色创建纯色画刷
        LSet tRCopy = tR                        '赋值矩形
        tRCopy.Left = tRCopy.Left + 1
        tRCopy.Top = tRCopy.Top + 1
        FrameRect lhDC, tRCopy, hBr2          '填充矩形边框，宽度始终为1
        If mvarHot Then
            'Hot状态绘制粗线！
            InflateRect tRCopy, -1, -1              '将矩形边界缩小1个单位
            FrameRect lhDC, tRCopy, hBr2      '继续绘制边框，宽度为1
        End If
        DeleteObject hBr2                       '释放资源
    End If
    
    '如果有提示文本，那么绘制一个红色三角形到右上角
    If mvarToolTipText <> "" And oTable.ShowToolTipText And bExportMode = False Then
        Dim poly(1 To 3) As POINTAPI, NumCoords As Long, hBrush As Long, hRgn As Long
        NumCoords = 3
        poly(1).X = (ColInfo(mvarCol).LeftX + mvarWidth) / p_TPPX
        poly(1).Y = (RowInfo(mvarRow).TopY) / p_TPPY + mvarGridLineWidth
        poly(2).X = poly(1).X - 5
        poly(2).Y = poly(1).Y
        poly(3).X = poly(1).X
        poly(3).Y = poly(1).Y + 5
        hBrush = CreateSolidBrush(vbRed)
        hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
        If hRgn Then FillRgn lhDC, hRgn, hBrush
        DeleteObject hRgn
        DeleteObject hBrush
    End If
End Sub

Private Sub Class_Initialize()
    '默认属性初始化
    Me.Margin = 2                   '默认边距为2
    Me.SingleLine = False           '默认为多行文本
    Me.MergeInfo = ""               '默认不合并
    Me.Selected = False             '默认未选中
    Me.Hot = False                  '默认False
    Me.Visibled = True              '默认显示
    Me.Width = 1600                 '默认列宽
    Me.Height = 300                 '默认行高
    Me.FixedWidth = False           '默认不锁定列宽，允许手工调整
    Me.AutoHeight = True            '默认为自动行高
    Me.Icon = -1                    '默认没有图标，-1
    Me.Text = ""                    '默认文本为空
    Me.Tag = ""                     '默认标记为空
    Me.FormatString = ""            '默认格式串为空
    Me.Indent = 0                   '默认缩进值为0
    Me.HAlignment = HALignLeft      '默认靠左
    Me.VAlignment = VALignTop       '默认靠上
    Me.ForeColor = vbBlack          '默认前景色为黑色
    Me.BackColor = -1               '默认背景色为透明
    Me.GridLineColor = vbBlack      '默认边框颜色为黑色
    Me.GridLineWidth = 1            '默认边框宽度为1
    Me.FontName = "宋体"            '默认字体为空，取系统默认字体
    Me.FontSize = 11                '默认字体大小为10磅
    Me.FontBold = False             '默认非粗体
    Me.FontItalic = False           '默认非斜体
    Me.FontStrikeout = False        '默认无删除线
    Me.FontUnderline = False        '默认无下划线
    Me.FontWeight = 400             '默认重量为400
    Me.Protected = False            '默认不保护
    Me.Dirty = True                 '默认需要重绘
End Sub
