VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'单元格类型
Public Enum CellTypeEnum
    cprCTFixtext = 0        '0-固定文本(不可编辑)
    cprCTText = 1           '1-文本型(可编辑多行文本)
    cprCTElement = 2        '2-单要素
    cprCTTextElement = 3    '3-文本与多要素混合编辑
    cprCTPicture = 4        '4-参考图
    cprCTReportPic = 5      '5-报告图
    cprCTSign = 6           '6-签名
    cprCTRowSign = 7        '7-行控签名
    cprCTColSign = 8        '8-列控签名
    cprCTPicMark = 9        '9-图片标记
End Enum
'横向对齐
Public Enum HAlignEnum
    HAlignGeneral = 1       '常规
    HALignLeft = 2          '左
    HAlignCenter = 3        '中
    HALignRight = 4         '右
End Enum
'纵向对齐,与F1控件相同
Public Enum VAlignEnum
    VALignTop = 1           '上
    VAlignCenter = 2        '中
    VALignBottom = 3        '下
End Enum

Private mvarIndex As Long                         '对象标记
Private mvarKey As String
Private mvarID As Double                          'ID <!--序列生成-->
Private mvar文件ID As Long                      '文件ID<!--病历文件列表ID，病历范文目录ID，电子病历记录ID-->
Private mvar对象序号 As Long                    '
Private mvar对象类型 As Integer                 '对像类型<!--0-固定文本(不可编辑) 1-文本型(可编辑多行文本) 2-单要素 3-文本与多要素混合编辑 4-参考图 5-报告图 6-签名 7-行控签名 8-列控签名 9-图像标记-->
Private mvar保留对象 As Boolean                 '保留对象与否<!--对象类型=0时固定为1,其它类型为编辑时该单元格是否可操作 0-可操作,1-不可操作.-->
Private mvar对象属性 As String                  '<!--以|为分隔的第0维存储单元格区域；其后，要素存储 自动转换文本|必填;签名存储 签名方式|签名规则|证书ID|显示手签|签名时间|显示时间|签名要素 ;其它为空-->
Private mvar内容行次 As Long                    '<!--仅对象类型=3时 以要素为分隔的行次,从0开始，其它为空-->
Private mvar内容文本 As String                  '<!--文本、要素、签名为显示内容;图像标记录为 选中|类型|内容|点集|X1|Y1|X2|Y2|填充色|填充方式|线条色|字体色|线型|线宽|字体名-->
Private mvar开始版 As Long                      '开始版，专用于“电子病历记录”
Private mvar终止版 As Long                      '终止版，专用于“电子病历记录”

'属性
Private mvarRow As Long                         '行，从1开始编号
Private mvarCol As Long                         '列，从1开始编号
Private mvarWidth As Integer                        '单元格宽度
Private mvarHeight As Integer                      '单元格高度
Private mvarFontName As String                  '字体>宋体</字体>
Private mvarFontSize As Integer                   '<字号>9</字号>
Private mvarFontBold As Boolean                 '<粗体>False</粗体>
Private mvarFontItalic As Boolean               '<斜体>False</斜体>
Private mvarFontUnderline As Boolean            '<下划线>False</下划线>
Private mvarFontStrikeout As Boolean            '<删除线>False</删除线>
Private mvarFontColor   As OLE_COLOR                 '<字体颜色>vbblack</字体颜色>
Private mvarHAlignment As Integer            '<横向对齐>F1HAlignCenter</横向对齐>
Private mvarVAlignment As Integer            '<纵向对齐>F1VAlignCenter</纵向对齐>
Private mvarCellLineTop As F1BorderStyleConstants               '<上边框>1</上边框><!--0-无边框 1-有边框 '4-虚线 5- 粗线-->
Private mvarCellLineBottom As F1BorderStyleConstants           '<下边框>1</下边框>
Private mvarCellLineLeft As F1BorderStyleConstants             '<左边框>1</左边框>
Private mvarCellLineRight As F1BorderStyleConstants           '<右边框>1</右边框>
Private mvarCellLineTopColor As OLE_COLOR            '<上边框颜色>vbBlack</上边框颜色>
Private mvarCellLineBottomColor As OLE_COLOR          '<下边框颜色>vbBlack</下边框颜色>
Private mvarCellLineLeftColor As OLE_COLOR          '左边框颜色>vbBlack</左边框颜色>
Private mvarCellLineRightColor As OLE_COLOR         '<右边框颜色>vbBlack</右边框颜色>
Private mvarMerge As Boolean                   '是否合并                                          if IIf(.Merge, InStr(.MergeRange, ";") > 0, True) Then '只有合并单元格首个或非合并单元格才算有效单元格
Private mvarMergeRange As String               '合并区域左上方单元格(横向,纵向);合并区域右下方单元格(横向,纵向),例:1,2;3,4表示从1行2列合并到3行4列;只有合并区域的首个单元格有合并区域属性,其它单元格没有
Private mvarTextKey As String                     '混合编辑时以|分隔的文本Key,从第一维起计
Private mvarElementKey As String                  '要素的Key,混合编辑时以|分隔的多个要素Key,从第一维起计
Private mvarPictureKey As String                  '图片的Key
Private mvarSignKey As String                     '签名的Key
Private mvarPicMarkKey As String                  '图片标记的Key,以|分隔的多个标记Key,从第一维起计
Private mvarProtected As Boolean        '是否保护
Private mvarToolTipText As String       '提示文本

Public Property Get Index() As Long                       '对象标记
    Index = mvarIndex
End Property
Public Property Let Index(ByVal vData As Long)                        '对象标记
    mvarIndex = vData
    mvar对象序号 = mvarIndex
End Property
Public Property Get Key() As String
    Key = mvarKey
End Property
Public Property Let Key(ByVal vData As String)
    mvarKey = vData
End Property
Public Property Get ID() As Double                        'ID
    ID = mvarID
End Property
Public Property Let ID(ByVal vData As Double)                         'ID
    mvarID = vData
End Property
Public Property Get 文件ID() As Long                    '文件ID
    文件ID = mvar文件ID
End Property
Public Property Let 文件ID(ByVal vData As Long)                    '文件ID
    mvar文件ID = vData
End Property
Public Property Get 对象序号() As Long                    '文件ID
    对象序号 = mvar对象序号
End Property
Public Property Let 对象序号(ByVal vData As Long)                    '文件ID
    mvar对象序号 = vData
End Property
Public Property Get 对象类型() As Integer               '对像类型<!--0-固定文本(不可编辑) 1-文本型(可编辑多行文本) 2-单要素 3-文本与多要素混合编辑 4-参考图 5-报告图 6-签名 7-行控签名 8-列控签名 9-图像标记-->
    对象类型 = mvar对象类型
End Property
Public Property Let 对象类型(ByVal vData As Integer)               '对像类型<!--0-固定文本(不可编辑) 1-文本型(可编辑多行文本) 2-单要素 3-文本与多要素混合编辑 4-参考图 5-报告图 6-签名 7-行控签名 8-列控签名 9-图像标记-->
    mvar对象类型 = vData
End Property
Public Property Get 保留对象() As Boolean               '保留对象与否 <!--对象类型=0时固定为1,其它类型为编辑时该单元格是否可操作 0-可操作,1-不可操作.-->
    保留对象 = mvar保留对象
End Property
Public Property Let 保留对象(ByVal vData As Boolean)               '保留对象与否
    mvar保留对象 = vData
End Property
Public Property Get 对象属性() As String                '<!--以|为分隔的第0维存储单元格区域；其后，要素存储 自动转换文本|必填;签名存储 签名方式|签名规则|证书ID|显示手签|签名时间|显示时间|签名要素 ;其它为空-->
    对象属性 = mvar对象属性
End Property
Public Property Let 对象属性(ByVal vData As String)                '<!--以|为分隔的第0维存储单元格区域；其后，要素存储 自动转换文本|必填;签名存储 签名方式|签名规则|证书ID|显示手签|签名时间|显示时间|签名要素 ;其它为空-->
    mvar对象属性 = vData
End Property
Public Property Get 内容行次() As Long                  '<!--仅对象类型=3时 以要素为分隔的行次,从0开始，其它为空-->
    内容行次 = mvar内容行次
End Property
Public Property Let 内容行次(ByVal vData As Long)                  '<!--仅对象类型=3时 以要素为分隔的行次,从0开始，其它为空-->
    mvar内容行次 = vData
End Property
Public Property Get 内容文本() As String                '<!--文本、要素、签名为显示内容;图像标记录为 选中|类型|内容|点集|X1|Y1|X2|Y2|填充色|填充方式|线条色|字体色|线型|线宽|字体名-->
    内容文本 = mvar内容文本
End Property
Public Property Let 内容文本(ByVal vData As String)                '<!--文本、要素、签名为显示内容;图像标记录为 选中|类型|内容|点集|X1|Y1|X2|Y2|填充色|填充方式|线条色|字体色|线型|线宽|字体名-->
    mvar内容文本 = vData
End Property
Public Property Get 开始版() As Long                    '开始版，专用于“电子病历记录”
    开始版 = mvar开始版
End Property
Public Property Let 开始版(ByVal vData As Long)                    '开始版，专用于“电子病历记录”
    mvar开始版 = vData
End Property
Public Property Get 终止版() As Long                    '终止版，专用于“电子病历记录”
    终止版 = mvar终止版
End Property
Public Property Let 终止版(ByVal vData As Long)                    '终止版，专用于“电子病历记录”
    mvar终止版 = vData
End Property
Public Property Get Row() As Long                       '行，从1开始编号
    Row = mvarRow
End Property
Public Property Let Row(ByVal vData As Long)                       '行，从1开始编号
    mvarRow = vData
End Property
Public Property Get Col() As Long                       '列，从1开始编号
    Col = mvarCol
End Property
Public Property Let Col(ByVal vData As Long)                       '列，从1开始编号
    mvarCol = vData
End Property
Public Property Get Width() As Integer                      '单元格宽度
    Width = mvarWidth
End Property
Public Property Let Width(ByVal vData As Integer)                     '单元格宽度
    mvarWidth = vData
End Property
Public Property Get Height() As Integer                    '单元格高度
    Height = mvarHeight
End Property
Public Property Let Height(ByVal vData As Integer)                    '单元格高度
    mvarHeight = vData
End Property
Public Property Get FontName() As String                '字体>宋体</字体>
    FontName = mvarFontName
End Property
Public Property Let FontName(ByVal vData As String)                '字体>宋体</字体>
     mvarFontName = vData
End Property
Public Property Get FontSize() As Integer                  '<字号>9</字号>
    FontSize = GetFontSizeNumber(mvarFontSize)
End Property
Public Property Let FontSize(ByVal vData As Integer)                '<字号>9</字号>
    mvarFontSize = vData
End Property
Public Property Get FontBold() As Boolean               '<粗体>False</粗体>
    FontBold = mvarFontBold
End Property
Public Property Let FontBold(ByVal vData As Boolean)               '<粗体>False</粗体>
    mvarFontBold = vData
End Property
Public Property Get FontItalic() As Boolean             '<斜体>False</斜体>
    FontItalic = mvarFontItalic
End Property
Public Property Let FontItalic(ByVal vData As Boolean)             '<斜体>False</斜体>
    mvarFontItalic = vData
End Property
Public Property Get FontUnderline() As Boolean          '<下划线>False</下划线>
    FontUnderline = mvarFontUnderline
End Property
Public Property Let FontUnderline(ByVal vData As Boolean)          '<下划线>False</下划线>
    mvarFontUnderline = vData
End Property
Public Property Get FontStrikeout() As Boolean          '<删除线>False</删除线>
    FontStrikeout = mvarFontStrikeout
End Property
Public Property Let FontStrikeout(ByVal vData As Boolean)          '<删除线>False</删除线>
    mvarFontStrikeout = vData
End Property
Public Property Get FontColor() As OLE_COLOR                 '<字体颜色>vbblack</字体颜色>
    FontColor = mvarFontColor
End Property
Public Property Let FontColor(ByVal vData As OLE_COLOR)                 '<字体颜色>vbblack</字体颜色>
    mvarFontColor = vData
End Property
Public Property Get HAlignment() As HAlignEnum          '<横向对齐>F1HAlignCenter</横向对齐>
    HAlignment = mvarHAlignment
End Property
Public Property Let HAlignment(ByVal vData As HAlignEnum)          '<横向对齐>F1HAlignCenter</横向对齐>
    mvarHAlignment = vData
End Property
Public Property Get VAlignment() As VAlignEnum          '<纵向对齐>F1VAlignCenter</纵向对齐>
    VAlignment = mvarVAlignment
End Property
Public Property Let VAlignment(ByVal vData As VAlignEnum)          '<纵向对齐>F1VAlignCenter</纵向对齐>
    mvarVAlignment = vData
End Property
Public Property Get CellLineTop() As F1BorderStyleConstants            '<上边框>1</上边框><!--0-无边框 1-有边框-->
    CellLineTop = mvarCellLineTop
End Property
Public Property Let CellLineTop(ByVal vData As F1BorderStyleConstants)            '<上边框>1</上边框><!--0-无边框 1-有边框-->
    mvarCellLineTop = vData
End Property
Public Property Get CellLineBottom() As F1BorderStyleConstants         '<下边框>1</下边框>
    CellLineBottom = mvarCellLineBottom
End Property
Public Property Let CellLineBottom(ByVal vData As F1BorderStyleConstants)         '<下边框>1</下边框>
    mvarCellLineBottom = vData
End Property
Public Property Get CellLineLeft() As F1BorderStyleConstants           '<左边框>1</左边框>
    CellLineLeft = mvarCellLineLeft
End Property
Public Property Let CellLineLeft(ByVal vData As F1BorderStyleConstants)           '<左边框>1</左边框>
    mvarCellLineLeft = vData
End Property
Public Property Get CellLineRight() As F1BorderStyleConstants         '<右边框>1</右边框>
    CellLineRight = mvarCellLineRight
End Property
Public Property Let CellLineRight(ByVal vData As F1BorderStyleConstants)         '<右边框>1</右边框>
    mvarCellLineRight = vData
End Property

Public Property Get CellLineTopColor() As OLE_COLOR          '<上边框颜色>vbBlack</上边框颜色>
    CellLineTopColor = mvarCellLineTopColor
End Property
Public Property Let CellLineTopColor(ByVal vData As OLE_COLOR)          '<上边框颜色>vbBlack</上边框颜色>
    mvarCellLineTopColor = vData
End Property
Public Property Get CellLineBottomColor() As OLE_COLOR        '<下边框颜色>vbBlack</下边框颜色>
    CellLineBottomColor = mvarCellLineBottomColor
End Property
Public Property Let CellLineBottomColor(ByVal vData As OLE_COLOR)        '<下边框颜色>vbBlack</下边框颜色>
    mvarCellLineBottomColor = vData
End Property
Public Property Get CellLineLeftColor() As OLE_COLOR        '左边框颜色>vbBlack</左边框颜色>
    CellLineLeftColor = mvarCellLineLeftColor
End Property
Public Property Let CellLineLeftColor(ByVal vData As OLE_COLOR)        '左边框颜色>vbBlack</左边框颜色>
    mvarCellLineLeftColor = vData
End Property
Public Property Get CellLineRightColor() As OLE_COLOR       '<右边框颜色>vbBlack</右边框颜色>
    CellLineRightColor = mvarCellLineRightColor
End Property
Public Property Let CellLineRightColor(ByVal vData As OLE_COLOR)       '<右边框颜色>vbBlack</右边框颜色>
    mvarCellLineRightColor = vData
End Property
Public Property Get Merge() As Boolean                 '是否合并
    Merge = mvarMerge
End Property
Public Property Let Merge(ByVal vData As Boolean)                 '是否合并
    mvarMerge = vData
End Property
Public Property Get MergeRange() As String                 '合并区域
    MergeRange = mvarMergeRange
End Property
Public Property Let MergeRange(ByVal vData As String)                 '合并区域
    mvarMergeRange = vData
End Property
Public Property Get TextKey() As String                '要素的Key
    TextKey = mvarTextKey
End Property
Public Property Let TextKey(ByVal vData As String)                '要素的Key
    mvarTextKey = vData
End Property
Public Property Get ElementKey() As String                '要素的Key
    ElementKey = mvarElementKey
End Property
Public Property Let ElementKey(ByVal vData As String)                '要素的Key
    mvarElementKey = vData
End Property
Public Property Get PictureKey() As String                '图片的Key
    PictureKey = mvarPictureKey
End Property
Public Property Let PictureKey(ByVal vData As String)                '图片的Key
    mvarPictureKey = vData
End Property
Public Property Get SignKey() As String                '签名的Key
    SignKey = mvarSignKey
End Property
Public Property Let SignKey(ByVal vData As String)                '签名的Key
    mvarSignKey = vData
End Property
Public Property Get PicMarkKey() As String                '图片标记的Key
    PicMarkKey = mvarPicMarkKey
End Property
Public Property Let PicMarkKey(ByVal vData As String)                '图片标记的Key
    mvarPicMarkKey = vData
End Property
Public Property Get Protected() As Boolean      '是否保护
    Protected = mvarProtected
End Property
Public Property Let Protected(ByVal vData As Boolean)      '是否保护
    mvarProtected = vData
End Property
Public Property Get ToolTipText() As String     '提示文本
    ToolTipText = mvarToolTipText
End Property
Public Property Let ToolTipText(ByVal vData As String)     '提示文本
    mvarToolTipText = vData
End Property
Public Sub Clone(cNode As cTabCell)
Dim i As Integer
    With cNode
        .Width = Me.Width                     '单元格宽度
        .Height = Me.Height                    '单元格高度
        .FontName = Me.FontName                '字体>宋体</字体>
        .FontSize = Me.FontSize                '<字号>9</字号>
        .FontBold = Me.FontBold               '<粗体>False</粗体>
        .FontItalic = Me.FontItalic             '<斜体>False</斜体>
        .FontUnderline = Me.FontUnderline          '<下划线>False</下划线>
        .FontStrikeout = Me.FontStrikeout          '<删除线>False</删除线>
        .FontColor = Me.FontColor                 '<字体颜色>vbblack</字体颜色>
        .HAlignment = Me.HAlignment           '<横向对齐>F1HAlignCenter</横向对齐>
        .VAlignment = Me.VAlignment          '<纵向对齐>F1VAlignCenter</纵向对齐>
        .CellLineTop = Me.CellLineBottom            '<上边框>1</上边框><!--0-无边框 1-有边框-->
        .CellLineBottom = Me.CellLineBottom         '<下边框>1</下边框>
        .CellLineLeft = Me.CellLineLeft           '<左边框>1</左边框>
        .CellLineRight = Me.CellLineRight         '<右边框>1</右边框>
        .CellLineTopColor = Me.CellLineBottomColor          '<上边框颜色>vbBlack</上边框颜色>
        .CellLineBottomColor = Me.CellLineBottomColor        '<下边框颜色>vbBlack</下边框颜色>
        .CellLineLeftColor = Me.CellLineLeftColor        '左边框颜色>vbBlack</左边框颜色>
        .CellLineRightColor = Me.CellLineRightColor       '<右边框颜色>vbBlack</右边框颜色>
        .ElementKey = Me.ElementKey                '要素的Key
        .PictureKey = Me.PictureKey                '图片的Key
        .SignKey = Me.SignKey                   '签名的Key
        .PicMarkKey = Me.PicMarkKey              '图片标记的Key
        .Protected = Me.Protected      '是否保护
        .ToolTipText = Me.ToolTipText     '提示文本
        
        .文件ID = Me.文件ID                  '文件ID<!--病历文件列表ID，病历范文目录ID，电子病历记录ID-->
        .对象类型 = Me.对象类型
        .保留对象 = Me.保留对象
        .对象属性 = Me.对象属性
        .内容行次 = Me.内容行次
        .内容文本 = Me.内容文本
        .开始版 = Me.开始版
        .终止版 = Me.终止版
    End With
End Sub

'################################################################################################################
'## 功能：  获取单元格属性字符串
'################################################################################################################
Public Property Get 单元属性() As String
Dim strR As String
    strR = strR & Me.Row & "|"                        '行，从1开始编号
    strR = strR & Me.Col & "|"                        '列，从1开始编号
    strR = strR & Me.Width & "|"                      '单元格宽度
    strR = strR & Me.Height & "|"                     '单元格高度
    strR = strR & Me.FontName & "|"                 '字体>宋体</字体>
    strR = strR & Me.FontSize & "|"                 '<字号>9</字号>
    strR = strR & Me.FontBold & "|"                 '<粗体>False</粗体>
    strR = strR & Me.FontItalic & "|"               '<斜体>False</斜体>
    strR = strR & Me.FontUnderline & "|"            '<下划线>False</下划线>
    strR = strR & Me.FontStrikeout & "|"            '<删除线>False</删除线>
    strR = strR & Me.FontColor & "|"                  '<字体颜色>vbblack</字体颜色>
    strR = strR & Me.HAlignment & "|"            '<横向对齐>F1HAlignCenter</横向对齐>
    strR = strR & Me.VAlignment & "|"           '<纵向对齐>F1VAlignCenter</纵向对齐>
    strR = strR & Me.CellLineTop & "|"              '<上边框>1</上边框><!--0-无边框 1-有边框-->
    strR = strR & Me.CellLineBottom & "|"           '<下边框>1</下边框>
    strR = strR & Me.CellLineLeft & "|"             '<左边框>1</左边框>
    strR = strR & Me.CellLineRight & "|"           '<右边框>1</右边框>
    strR = strR & Me.CellLineTopColor & "|"           '<上边框颜色>vbBlack</上边框颜色>
    strR = strR & Me.CellLineBottomColor & "|"         '<下边框颜色>vbBlack</下边框颜色>
    strR = strR & Me.CellLineLeftColor & "|"         '左边框颜色>vbBlack</左边框颜色>
    strR = strR & Me.CellLineRightColor & "|"        '<右边框颜色>vbBlack</右边框颜色>
    strR = strR & Me.Merge & "|"                   '是否合并
    strR = strR & Me.MergeRange & "|"             '合并区域左上方单元格(横向,纵向);合并区域右下方单元格(横向,纵向)
    strR = strR & Me.ElementKey & "|"                 '要素的Key
    strR = strR & Me.PictureKey & "|"                 '图片的Key
    strR = strR & Me.SignKey & "|"                   '签名的Key
    strR = strR & Me.PicMarkKey & "|"               '图片标记的Key
    strR = strR & Me.Protected & "|"        '是否保护
    strR = strR & Me.ToolTipText & "|"      '提示文本
    
    单元属性 = strR
End Property

'################################################################################################################
'## 功能：  通过一个字符串来设置单元格属性
'################################################################################################################
Public Property Let 单元属性(ByRef vData As String)
    On Error Resume Next
    Dim T As Variant
    T = Split(vData, "|")
    Me.Row = T(0)                       '行，从1开始编号
    Me.Col = T(1)                       '列，从1开始编号
    Me.Width = T(2)                     '单元格宽度
    Me.Height = T(3)                    '单元格高度
    Me.FontName = T(4)                 '字体>宋体</字体>
    Me.FontSize = T(5)                 '<字号>9</字号>
    Me.FontBold = T(6)               '<粗体>False</粗体>
    Me.FontItalic = T(7)             '<斜体>False</斜体>
    Me.FontUnderline = T(8)          '<下划线>False</下划线>
    Me.FontStrikeout = T(9)          '<删除线>False</删除线>
    Me.FontColor = T(10)                 '<字体颜色>vbblack</字体颜色>
    Me.HAlignment = T(11)           '<横向对齐>F1HAlignCenter</横向对齐>
    Me.VAlignment = T(12)           '<纵向对齐>F1VAlignCenter</纵向对齐>
    Me.CellLineTop = T(13)            '<上边框>1</上边框><!--0-无边框 1-有边框-->
    Me.CellLineBottom = T(14)         '<下边框>1</下边框>
    Me.CellLineLeft = T(15)           '<左边框>1</左边框>
    Me.CellLineRight = T(16)         '<右边框>1</右边框>
    Me.CellLineTopColor = T(17)          '<上边框颜色>vbBlack</上边框颜色>
    Me.CellLineBottomColor = T(18)        '<下边框颜色>vbBlack</下边框颜色>
    Me.CellLineLeftColor = T(19)        '左边框颜色>vbBlack</左边框颜色>
    Me.CellLineRightColor = T(20)       '<右边框颜色>vbBlack</右边框颜色>
    Me.Merge = T(21)                 '是否合并
    Me.MergeRange = T(22)            '合并区域左上方单元格(横向,纵向);合并区域右下方单元格(横向,纵向)
    Me.ElementKey = T(23)                '要素的Key
    Me.PictureKey = T(24)              '图片的Key
    Me.SignKey = T(25)                   '签名的Key
    Me.PicMarkKey = T(26)             '图片标记的Key
    Me.Protected = T(27)      '是否保护
    Me.ToolTipText = T(28)      '提示文本
End Property
Public Property Get CellTypeName() As String
    Select Case Me.对象类型
        Case cprCTFixtext
            CellTypeName = "固定文本"
        Case cprCTText
            CellTypeName = "文本型"
        Case cprCTElement
            CellTypeName = "单要素"
        Case cprCTTextElement
            CellTypeName = "混合编辑"
        Case cprCTPicture
            CellTypeName = "参考图"
        Case cprCTReportPic
            CellTypeName = "报告图"
        Case cprCTSign
            CellTypeName = "签名"
        Case cprCTRowSign
            CellTypeName = "行控签名"
        Case cprCTColSign
            CellTypeName = "列控签名"
    End Select
End Property
Public Function GetSaveSQL(ByVal editType As TabET, arrSQL As Variant) As Boolean
'功能:保存固定文本,单文本内容,混合区域的主项
    On Error GoTo errHand
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    Select Case editType
    Case TabET_病历文件定义
        arrSQL(UBound(arrSQL)) = "Zl_病历文件结构_Update(" & mvarID & "," & mvar文件ID & _
                                    ",Null," & mvar对象序号 & "," & mvar对象类型 & ",Null," & IIf(mvar保留对象, 1, 0) & ",'" & _
                                    mvarRow & "," & mvarCol & "|" & mvar对象属性 & "',Null,'" & Replace(mvar内容文本, "'", "'" & "||Chr(39)||" & "'") & "')"
    Case TabET_全文示范编辑
        arrSQL(UBound(arrSQL)) = "Zl_病历范文内容_Update(" & mvarID & "," & mvar文件ID & _
                                    ",Null," & mvar对象序号 & "," & mvar对象类型 & ",Null," & IIf(mvar保留对象, 1, 0) & ",'" & _
                                    mvarRow & "," & mvarCol & "|" & mvar对象属性 & "',Null,'" & Replace(mvar内容文本, "'", "'" & "||Chr(39)||" & "'") & "')"
    Case TabET_单病历编辑, TabET_单病历审核
        arrSQL(UBound(arrSQL)) = "Zl_电子病历内容_Update(" & mvarID & "," & mvar文件ID & "," & mvar开始版 & "," & mvar终止版 & _
                                    ",Null," & mvar对象序号 & "," & mvar对象类型 & ",Null," & IIf(mvar保留对象, 1, 0) & ",'" & _
                                    mvarRow & "," & mvarCol & "|" & mvar对象属性 & "',Null,'" & Replace(mvar内容文本, "'", "'" & "||Chr(39)||" & "'") & "')"
    End Select
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Public Function clText(ByVal clt As cTabTexts, ByVal lSeq As Long) As cTabText
'跟据内容行次读取内容
Dim i As Integer
    For i = 1 To clt.Count
        If clt(i).内容行次 = lSeq And InStr(Me.TextKey & "|", "|" & clt(i).Key & "|") > 0 Then
            Set clText = clt(i)
            Exit Function
        End If
    Next
    '没找到时
    Set clText = New cTabText
End Function
Public Function clElement(ByVal els As cTabElements, ByVal lSeq As Long) As cTabElement
'跟据内容行次读取内容
Dim i As Integer
    For i = 1 To els.Count
        If els(i).内容行次 = lSeq And InStr(Me.ElementKey & "|", "|" & els(i).Key & "|") > 0 Then
            Set clElement = els(i)
            Exit Function
        End If
    Next
End Function
Private Sub Class_Initialize()
    mvarFontName = "宋体"
    mvarFontSize = 9
    mvarHAlignment = HAlignCenter
    mvarVAlignment = VAlignCenter
    mvarCellLineTop = 1
    mvarCellLineBottom = 1
    mvarCellLineLeft = 1
    mvarCellLineRight = 1
    mvarCellLineTopColor = 0
    mvarCellLineBottomColor = 0
    mvarCellLineLeftColor = 0
    mvarCellLineRightColor = 0
    mvar对象序号 = mvarIndex
End Sub
'################################################################################################################
'## 功能：  检查所有对象的ID是否有效
'##
'## 参数：  eEditType :当前的编辑模式
'##
'## 说明：  如果没有ID，则到数据库中提取一个唯一ID号。父ID=0的情况是单元素由Cell取ID
'################################################################################################################
Public Sub CheckValidID(eEditType As TabET)
    Dim i As Long, strTable As String
    Select Case eEditType
    Case TabET_病历文件定义
        strTable = "病历文件结构"
        If Me.ID = 0 Then
            If IIf(Me.Merge, InStr(Me.MergeRange, ";") > 0, True) Then  '只有非合并单元格和合并单元格的首个单元格才记录ID
                Me.ID = zlDatabase.GetNextId(strTable)
            End If
        End If
    Case TabET_全文示范编辑
        strTable = "病历范文内容"
        If Me.ID = 0 Then
            If IIf(Me.Merge, InStr(Me.MergeRange, ";") > 0, True) Then  '只有非合并单元格和合并单元格的首个单元格才记录ID
                Me.ID = zlDatabase.GetNextId(strTable)
            End If
        End If
    Case TabET_单病历编辑, TabET_单病历审核
        strTable = "电子病历内容"
        If Me.ID = 0 Then
            If IIf(Me.Merge, InStr(Me.MergeRange, ";") > 0, True) Then  '只有非合并单元格和合并单元格的首个单元格才记录ID
                Me.ID = GetEPRContentNextId
            End If
        End If
    End Select
    
End Sub

