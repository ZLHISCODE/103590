VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'################################################################################################################
'##模 块 名：cEPRCell.cls
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：表格单元格
'##版    本：
'################################################################################################################
Option Explicit

'################################################################################################################
'## 枚举常量
'################################################################################################################

'单元格类型
Public Enum CellTypeEnum
    cprCTEText = 0          '普通文本
    cprCTEElement = 1       '绑定诊治要素
End Enum

'################################################################################################################
'## 内部变量
'################################################################################################################

Private mvarKey As Long                         '对象标记
Private mvarID As Double                          'ID（单元/文本、诊治要素的ID）
Private mvar文件ID As Long                      '文件ID
Private mvar父ID As Double                       '父ID（表格ID）
Private mvar对象序号 As Double                    '对象序号
Private mvar保留对象 As Boolean                 '保留对象与否
Private mvar开始版 As Long              '开始版，专用于“电子病历记录”
Private mvar终止版 As Long              '终止版，专用于“电子病历记录”
Private mvar内容文本 As String                  '内容文本

Private mvarRow As Long                         '行，从1开始编号
Private mvarCol As Long                         '列，从1开始编号
Private mvarWidth As Long                       '单元格宽度
Private mvarHeight As Long                      '单元格高度
Private mvarMergeNo As String                   '合并号
Private mvarCellType As CellTypeEnum            '节点类型
Private mvarElementKey As Long                 '诊治要素的Key
Private mvarPictureKey As Long                 '诊治要素的Key

'新属性
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

Public Property Let Margin(ByVal vData As Long)
    mvarMargin = vData
End Property

Public Property Get Margin() As Long
    Margin = mvarMargin
End Property

Public Property Let SingleLine(ByVal vData As Boolean)
    mvarSingleLine = vData
End Property

Public Property Get SingleLine() As Boolean
    SingleLine = mvarSingleLine
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let GridLineColor(ByVal vData As OLE_COLOR)
    mvarGridLineColor = vData
End Property

Public Property Get GridLineColor() As OLE_COLOR
    GridLineColor = mvarGridLineColor
End Property

Public Property Let GridLineWidth(ByVal vData As Long)
    mvarGridLineWidth = vData
End Property

Public Property Get GridLineWidth() As Long
    GridLineWidth = mvarGridLineWidth
End Property

Public Property Let FixedWidth(ByVal vData As Boolean)
    mvarFixedWidth = vData
End Property

Public Property Get FixedWidth() As Boolean
    FixedWidth = mvarFixedWidth
End Property

Public Property Let AutoHeight(ByVal vData As Boolean)
    mvarAutoHeight = vData
End Property

Public Property Get AutoHeight() As Boolean
    AutoHeight = mvarAutoHeight
End Property

Public Property Let FontName(ByVal vData As String)
    mvarFontName = vData
End Property

Public Property Get FontName() As String
    FontName = mvarFontName
End Property

Public Property Let FontSize(ByVal vData As Long)
    mvarFontSize = vData
End Property

Public Property Get FontSize() As Long
    FontSize = mvarFontSize
End Property

Public Property Let FontBold(ByVal vData As Boolean)
    mvarFontBold = vData
End Property

Public Property Get FontBold() As Boolean
    FontBold = mvarFontBold
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
    mvarFontItalic = vData
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mvarFontItalic
End Property

Public Property Let FontStrikeout(ByVal vData As Boolean)
    mvarFontStrikeout = vData
End Property

Public Property Get FontStrikeout() As Boolean
    FontStrikeout = mvarFontStrikeout
End Property

Public Property Let FontUnderline(ByVal vData As Boolean)
    mvarFontUnderline = vData
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = mvarFontUnderline
End Property

Public Property Let FontWeight(ByVal vData As Long)
    mvarFontWeight = vData
End Property

Public Property Get FontWeight() As Long
    FontWeight = mvarFontWeight
End Property

Public Property Let FormatString(ByVal vData As String)
    mvarFormatString = vData
End Property

Public Property Get FormatString() As String
    FormatString = mvarFormatString
End Property

Public Property Let Indent(ByVal vData As Long)
    mvarIndent = vData
End Property

Public Property Get Indent() As Long
    Indent = mvarIndent
End Property

Public Property Let HAlignment(ByVal vData As HAlignEnum)
    mvarHAlignment = vData
End Property

Public Property Get HAlignment() As HAlignEnum
    HAlignment = mvarHAlignment
End Property

Public Property Let VAlignment(ByVal vData As VAlignEnum)
    mvarVAlignment = vData
End Property

Public Property Get VAlignment() As VAlignEnum
    VAlignment = mvarVAlignment
End Property

Public Property Let Protected(ByVal vData As Boolean)
    mvarProtected = vData
End Property

Public Property Get Protected() As Boolean
    Protected = mvarProtected
End Property

Public Property Let ToolTipText(ByVal vData As String)
    mvarToolTipText = vData
End Property

Public Property Get ToolTipText() As String
    ToolTipText = mvarToolTipText
End Property

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let 文件ID(ByVal vData As Long)
    mvar文件ID = vData
End Property

Public Property Get 文件ID() As Long
    文件ID = mvar文件ID
End Property

Public Property Let 父ID(ByVal vData As Double)
    mvar父ID = vData
End Property

Public Property Get 父ID() As Double
    父ID = mvar父ID
End Property

Public Property Let 对象序号(ByVal vData As Double)
    mvar对象序号 = vData
End Property

Public Property Get 对象序号() As Double
    对象序号 = mvar对象序号
End Property

Public Property Let 保留对象(ByVal vData As Boolean)
    mvar保留对象 = vData
End Property

Public Property Get 保留对象() As Boolean
    保留对象 = mvar保留对象
End Property

Public Property Let 开始版(ByVal vData As Long)
    mvar开始版 = vData
End Property

Public Property Get 开始版() As Long
    开始版 = IIf(mvar开始版 = 0, 1, mvar开始版)
End Property

Public Property Let 终止版(ByVal vData As Long)
    mvar终止版 = vData
End Property

Public Property Get 终止版() As Long
'    终止版 = IIf(mvar终止版 = 0, 0, mvar终止版)
    终止版 = 0
End Property

Public Property Let 内容文本(ByVal vData As String)
    mvar内容文本 = vData
End Property

Public Property Get 内容文本() As String
    内容文本 = mvar内容文本
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

Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let MergeNo(ByVal vData As String)
    mvarMergeNo = vData
End Property

Public Property Get MergeNo() As String
    MergeNo = mvarMergeNo
End Property

Public Property Let CellType(ByVal vData As CellTypeEnum)
    mvarCellType = vData
End Property

Public Property Get CellType() As CellTypeEnum
    CellType = mvarCellType
End Property

Public Property Let ElementKey(ByVal vData As Long)
    mvarElementKey = vData
End Property

Public Property Get ElementKey() As Long
    ElementKey = mvarElementKey
End Property

Public Property Let PictureKey(ByVal vData As Long)
    mvarPictureKey = vData
End Property

Public Property Get PictureKey() As Long
    PictureKey = mvarPictureKey
End Property

'################################################################################################################
'## 功能：  复制提纲对象
'##
'## 参数：  blnWithKey  :是否连同Key值一起复制！
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRCell
    Dim cNode As New cEPRCell
    If blnWithKey Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.文件ID = Me.文件ID
    cNode.父ID = Me.父ID
    cNode.对象序号 = Me.对象序号
    cNode.内容文本 = Me.内容文本
    cNode.Row = Me.Row
    cNode.Col = Me.Col
    cNode.Width = Me.Width
    cNode.Height = Me.Height
    cNode.MergeNo = Me.MergeNo
    cNode.ElementKey = Me.ElementKey
    cNode.PictureKey = Me.PictureKey
    cNode.Margin = Me.Margin
    cNode.SingleLine = Me.SingleLine
    cNode.ForeColor = Me.ForeColor
    cNode.BackColor = Me.BackColor
    cNode.GridLineColor = Me.GridLineColor
    cNode.GridLineWidth = Me.GridLineWidth
    cNode.FixedWidth = Me.FixedWidth
    cNode.AutoHeight = Me.AutoHeight
    cNode.FontName = Me.FontName
    cNode.FontSize = Me.FontSize
    cNode.FontBold = Me.FontBold
    cNode.FontItalic = Me.FontItalic
    cNode.FontStrikeout = Me.FontStrikeout
    cNode.FontUnderline = Me.FontUnderline
    cNode.FontWeight = Me.FontWeight
    cNode.FormatString = Me.FormatString
    cNode.Indent = Me.Indent
    cNode.HAlignment = Me.HAlignment
    cNode.VAlignment = Me.VAlignment
    cNode.Protected = Me.Protected
    cNode.ToolTipText = Me.ToolTipText
    Set Clone = cNode
End Function

'################################################################################################################
'## 功能：  获取单元格属性字符串
'################################################################################################################
Public Property Get 对象属性() As String
'    '由Format得出格式字符串，供数据库存储用
'    Dim strR As String
'    strR = Me.诊治要素Key & "|"
'    strR = strR & Me.Row & "|"
'    strR = strR & Me.Col & "|"
'    strR = strR & Me.MergeNo & "|"
'    strR = strR & Me.Width & "|"
'    strR = strR & Me.Height & "|"
'    strR = strR + IIf(CellFormat.FontBold, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.FontItalic, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.FontStrikeout, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.FontUnderline, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.MergeCells, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.ProtectionHidden, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.ProtectionLocked, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.WordWrap, "1", "0") & "|"
'    strR = strR & CellFormat.PatternStyle & "|"
'    strR = strR & CellFormat.FontSize & "|"
'    strR = strR & CellFormat.AlignHorizontal & "|"
'    strR = strR & CellFormat.AlignVertical & "|"
'    strR = strR & CellFormat.FontCharSet & "|"
'    strR = strR & CellFormat.FontColor & "|"
'    strR = strR & CellFormat.PatternBG & "|"
'    strR = strR & CellFormat.PatternFG & "|"
'    strR = strR & CellFormat.BorderColor(F1HInsideBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1VInsideBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1TopBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1LeftBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1BottomBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1RightBorder) & "|"
'    strR = strR & CellFormat.BorderStyle(F1HInsideBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1VInsideBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1TopBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1LeftBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1BottomBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1RightBorder) & "|"
'    strR = strR & CellFormat.ValidationRule & "|"
'    strR = strR & CellFormat.ValidationRuleLocal & "|"
'    strR = strR & CellFormat.ValidationText & "|"
'    strR = strR & CellFormat.NumberFormat & "|"
'    strR = strR & CellFormat.NumberFormatLocal & "|"
'    strR = strR & Trim(CellFormat.FontName)
'    对象属性 = strR
    Dim strR As String
    strR = mvarElementKey & "|"
    strR = strR & mvarPictureKey & "|"
    strR = strR & Me.Row & "|"
    strR = strR & Me.Col & "|"
    strR = strR & Me.MergeNo & "|"
    strR = strR & Me.Width & "|"
    strR = strR & Me.Height & "|"
    strR = strR & Me.Margin & "|"
    strR = strR & Me.SingleLine & "|"
    strR = strR & Me.ForeColor & "|"
    strR = strR & Me.BackColor & "|"
    strR = strR & Me.GridLineColor & "|"
    strR = strR & Me.GridLineWidth & "|"
    strR = strR & Me.FixedWidth & "|"
    strR = strR & Me.AutoHeight & "|"
    strR = strR & Me.FontName & "|"
    strR = strR & Me.FontSize & "|"
    strR = strR & Me.FontBold & "|"
    strR = strR & Me.FontItalic & "|"
    strR = strR & Me.FontStrikeout & "|"
    strR = strR & Me.FontUnderline & "|"
    strR = strR & Me.FontWeight & "|"
    strR = strR & Me.FormatString & "|"
    strR = strR & Me.HAlignment & "|"
    strR = strR & Me.VAlignment & "|"
    strR = strR & Me.Protected & "|"
    对象属性 = strR
End Property

'################################################################################################################
'## 功能：  通过一个字符串来设置单元格属性
'################################################################################################################
Public Property Let 对象属性(ByRef vData As String)
'    Dim T As Variant, i As Long, T2 As Variant
'    T = Split(vData, "|")
'    If UBound(T) > 0 Then
'        Me.诊治要素Key = T(0)
'        Me.Row = T(1)
'        Me.Col = T(2)
'        Me.MergeNo = T(3)
'        Me.Width = T(4)
'        Me.Height = T(5)
'
'        T2 = Split(T(6), ";")
'        Me.CellFormat.FontBold = T2(0)
'        Me.CellFormat.FontItalic = T2(1)
'        Me.CellFormat.FontStrikeout = T2(2)
'        Me.CellFormat.FontUnderline = T2(3)
'        Me.CellFormat.MergeCells = T2(4)
'        Me.CellFormat.ProtectionHidden = T2(5)
'        Me.CellFormat.ProtectionLocked = T2(6)
'        Me.CellFormat.WordWrap = T2(7)
'
'        Me.CellFormat.PatternStyle = T(7)
'        Me.CellFormat.FontSize = T(8)
'        Me.CellFormat.AlignHorizontal = T(9)
'        Me.CellFormat.AlignVertical = T(10)
'        Me.CellFormat.FontCharSet = T(11)
'        Me.CellFormat.FontColor = T(12)
'        Me.CellFormat.PatternBG = T(13)
'        Me.CellFormat.PatternFG = T(14)
'
'        T2 = Split(T(15), ";")
'        Me.CellFormat.BorderColor(F1HInsideBorder) = T2(0)
'        Me.CellFormat.BorderColor(F1VInsideBorder) = T2(1)
'        Me.CellFormat.BorderColor(F1TopBorder) = T2(2)
'        Me.CellFormat.BorderColor(F1LeftBorder) = T2(3)
'        Me.CellFormat.BorderColor(F1BottomBorder) = T2(4)
'        Me.CellFormat.BorderColor(F1RightBorder) = T2(5)
'
'        T2 = Split(T(16), ";")
'        Me.CellFormat.BorderStyle(F1HInsideBorder) = T2(0)
'        Me.CellFormat.BorderStyle(F1VInsideBorder) = T2(1)
'        Me.CellFormat.BorderStyle(F1TopBorder) = T2(2)
'        Me.CellFormat.BorderStyle(F1LeftBorder) = T2(3)
'        Me.CellFormat.BorderStyle(F1BottomBorder) = T2(4)
'        Me.CellFormat.BorderStyle(F1RightBorder) = T2(5)
'
'        Me.CellFormat.ValidationRule = T(17)
'        Me.CellFormat.ValidationRuleLocal = T(18)
'        Me.CellFormat.ValidationText = T(19)
'        Me.CellFormat.NumberFormat = T(20)
'        Me.CellFormat.NumberFormatLocal = T(21)
'        Me.CellFormat.FontName = T(22)
'    End If
    On Error Resume Next
    Dim T As Variant, i As Long, T2 As Variant
    T = Split(vData, "|")
    If UBound(T) = 22 Then
        Me.ElementKey = T(0)
        Me.Row = T(1)
        Me.Col = T(2)
        Me.MergeNo = T(3)
        Me.Width = T(4)
        Me.Height = T(5)

        T2 = Split(T(6), ";")
        Me.FontBold = T2(0)
        Me.FontItalic = T2(1)
        Me.FontStrikeout = T2(2)
        Me.FontUnderline = T2(3)
        Me.Protected = (Me.ElementKey > 0)
        Me.SingleLine = False
        Me.GridLineColor = vbBlack
        Me.GridLineWidth = 1

        Me.FontSize = T(8)
        Me.ForeColor = T(12)
        Me.BackColor = vbWhite
        Me.FixedWidth = False
        Me.AutoHeight = True
        Me.FontName = T(22)
    ElseIf UBound(T) = 26 Then
        Me.ElementKey = T(0)
        Me.PictureKey = T(1)
        Me.Row = T(2)
        Me.Col = T(3)
        Me.MergeNo = T(4)
        Me.Width = T(5)
        Me.Height = T(6)
        Me.Margin = T(7)
        Me.SingleLine = T(8)
        Me.ForeColor = T(9)
        Me.BackColor = T(10)
        Me.GridLineColor = T(11)
        Me.GridLineWidth = T(12)
        Me.FixedWidth = T(13)
        Me.AutoHeight = T(14)
        Me.FontName = T(15)
        Me.FontSize = T(16)
        Me.FontBold = T(17)
        Me.FontItalic = T(18)
        Me.FontStrikeout = T(19)
        Me.FontUnderline = T(20)
        Me.FontWeight = T(21)
        Me.FormatString = T(22)
        Me.HAlignment = T(23)
        Me.VAlignment = T(24)
        Me.Protected = T(25)
    End If
End Property

'################################################################################################################
'## 功能：  检查对象的ID是否有效
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。
'################################################################################################################
Public Sub CheckValidID(eEdtType As EditTypeEnum)
    Select Case eEdtType
    Case cprET_病历文件定义
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("病历文件结构")
    Case cprET_全文示范编辑
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("病历范文内容")
    Case cprET_单病历编辑, cprET_单病历审核
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
End Sub

