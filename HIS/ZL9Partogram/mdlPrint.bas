Attribute VB_Name = "mdlPrint"
Option Explicit

'Window版本函数
Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Const DC_PAPERNAMES = 16 '纸张名称(每64字符为一段,以Chr(0)结束)
Public Const DC_PAPERS = 2 '纸张编号(Array or Word)
Public Const DC_BINNAMES = 12 '进纸方式(每24字符为一段,以Chr(0)结束)
Public Const DC_BINS = 6 '进纸编号(Array or Word)

'打印纸张常量(256=自定义)
Public Const PageSize1 = "信笺， 8 1/2 x 11 英寸"
Public Const PageSize2 = "+A611 小型信笺， 8 1/2 x 11 英寸"
Public Const PageSize3 = "小型报， 11 x 17 英寸"
Public Const PageSize4 = "分类帐， 17 x 11 英寸"
Public Const PageSize5 = "法律文件， 8 1/2 x 14 英寸"
Public Const PageSize6 = "声明书，5 1/2 x 8 1/2 英寸"
Public Const PageSize7 = "行政文件，7 1/2 x 10 1/2 英寸"
Public Const PageSize8 = "A3, 297 x 420 毫米"
Public Const PageSize9 = "A4, 210 x 297 毫米"
Public Const PageSize10 = "A4小号， 210 x 297 毫米"
Public Const PageSize11 = "A5, 148 x 210 毫米"
Public Const PageSize12 = "B4, 250 x 354 毫米"
Public Const PageSize13 = "B5, 182 x 257 毫米"
Public Const PageSize14 = "对开本， 8 1/2 x 13 英寸"
Public Const PageSize15 = "四开本， 215 x 275 毫米"
Public Const PageSize16 = "10 x 14 英寸"
Public Const PageSize17 = "11 x 17 英寸"
Public Const PageSize18 = "便条，8 1/2 x 11 英寸"
Public Const PageSize19 = "#9 信封， 3 7/8 x 8 7/8 英寸"
Public Const PageSize20 = "#10 信封， 4 1/8 x 9 1/2 英寸"
Public Const PageSize21 = "#11 信封， 4 1/2 x 10 3/8 英寸"
Public Const PageSize22 = "#12 信封， 4 1/2 x 11 英寸"
Public Const PageSize23 = "#14 信封， 5 x 11 1/2 英寸"
Public Const PageSize24 = "C 尺寸工作单"
Public Const PageSize25 = "D 尺寸工作单"
Public Const PageSize26 = "E 尺寸工作单"
Public Const PageSize27 = "DL 型信封， 110 x 220 毫米"
Public Const PageSize28 = "C5 型信封， 162 x 229 毫米"
Public Const PageSize29 = "C3 型信封， 324 x 458 毫米"
Public Const PageSize30 = "C4 型信封， 229 x 324 毫米"
Public Const PageSize31 = "C6 型信封， 114 x 162 毫米"
Public Const PageSize32 = "C65 型信封，114 x 229 毫米"
Public Const PageSize33 = "B4 型信封， 250 x 353 毫米"
Public Const PageSize34 = "B5 型信封，176 x 250 毫米"
Public Const PageSize35 = "B6 型信封， 176 x 125 毫米"
Public Const PageSize36 = "信封， 110 x 230 毫米"
Public Const PageSize37 = "信封大王， 3 7/8 x 7 1/2 英寸"
Public Const PageSize38 = "信封， 3 5/8 x 6 1/2 英寸"
Public Const PageSize39 = "U.S. 标准复写簿， 14 7/8 x 11 英寸"
Public Const PageSize40 = "德国标准复写簿， 8 1/2 x 12 英寸"
Public Const PageSize41 = "德国法律复写簿， 8 1/2 x 13 英寸"

Public Const conBin1 = "上层纸盒进纸"
Public Const conBin2 = "下层纸盒进纸"
Public Const conBin3 = "中间纸盒进纸"
Public Const conBin4 = "等待手动插入每页纸"
Public Const conBin5 = "信封进纸器进纸"
Public Const conBin6 = "信封进纸器进纸；但要等待手动插入"
Public Const conBin7 = "当前缺省纸盒进纸"
Public Const conBin8 = "拖拉进纸器进纸"
Public Const conBin9 = "小型进纸器进纸"
Public Const conBin10 = "大型纸盒进纸"
Public Const conBin11 = "大容量进纸器进纸"

'纸张打印边界控制================================================================
Public Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, ByVal lpOutput As String, lpDevMode As Any) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'不同打印机的打印单元精度不同

Public Const PHYSICALHEIGHT = 111  'Physical Height in device units
Public Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Public Const PHYSICALOFFSETY = 113 'Physical Printable Area y margin
Public Const LOGPIXELSX = 88 'Number of pixels per logical inch along the screen width
Public Const LOGPIXELSY = 90
Public Const PHYSICALWIDTH = 110 '  Physical Width in device units
Public Const SCALINGFACTORX = 114  'Scaling factor x
Public Const SCALINGFACTORY = 115  'Scaling factor y
Public Const DRIVERVERSION = 0     'Device driver version

'################################################################################################################
'## 图片缩放模式设置
'######################################################################################
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Public Const BLACKONWHITE = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
Public Const STRETCH_ANDSCANS = BLACKONWHITE
Public Const STRETCH_ORSCANS = WHITEONBLACK
Public Const STRETCH_DELETESCANS = COLORONCOLOR
Public Const STRETCH_HALFTONE = HALFTONE


Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'***************************************************************
'绘画相关API及常量、结构定义
'***************************************************************
'结构定义

'点的坐标
Public Type POINTAPI
    X As Long
    Y As Long
End Type

'文字高度和宽度
Private Type Size
    W   As Long
    H   As Long
End Type

'巨型
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'画笔
Public Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type
'刷子
Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

'字体属性
Public Const LF_FACESIZE = 32

Public Type LogFont
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public T_OldPoint   As POINTAPI
Public T_NewPoint   As POINTAPI
Public T_ClientRect As RECT
Public T_LableRect  As RECT      '待输出文本的有效区域
Public T_Brush      As LOGBRUSH
Public T_Font       As LogFont
Public T_Size       As Size

'创建或得到现有对象
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, _
                            lpBits As Any) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

'创建画笔、刷子
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, _
                            lpStyle As Long) As Long

Public Const PS_SOLID = 0
Public Const PS_DASH = 1                    '  -------
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5                    '不允许画图

Public Const PS_COSMETIC = &H0
Public Const PS_GEOMETRIC = &H10000
Public Const PS_ALTERNATE = 8
Public Const PS_ENDCAP_FLAT = &H200
Public Const PS_ENDCAP_MASK = &HF00
Public Const PS_ENDCAP_ROUND = &H0
Public Const PS_ENDCAP_SQUARE = &H100
Public Const PS_JOIN_BEVEL = &H1000
Public Const PS_JOIN_MASK = &HF000
Public Const PS_JOIN_MITER = &H2000
Public Const PS_JOIN_ROUND = &H0

'CreateSolidBrush 创建纯色画刷
'CreateBrushIndirect 通过 LOGBRUSH 类型创建画刷
'CreateHatchBrush 创建阴影画刷
'CreatePatternBrush 创建图案画刷
'GetSysColorBrush 创建系统标准色画刷
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

'//lbStyle　可选值:
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Const BS_HOLLOW = BS_NULL
Public Const BS_HATCHED = 2
Public Const BS_PATTERN = 3
Public Const BS_INDEXED = 4
Public Const BS_DIBPATTERN = 5
Public Const BS_DIBPATTERNPT = 6
Public Const BS_PATTERN8X8 = 7
Public Const BS_DIBPATTERN8X8 = 8
Public Const BS_MONOPATTERN = 9

'//lbHatch　可选值:
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_VERTICAL = 1                '  |||||
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx

Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
'nIndex,同上面函数的lbHatch
'Public Const HS_HORIZONTAL = 0              '  -----
'Public Const HS_VERTICAL = 1                '  |||||
'Public Const HS_FDIAGONAL = 2               '  \\\\\
'Public Const HS_BDIAGONAL = 3               '  /////
'Public Const HS_CROSS = 4                   '  +++++
'Public Const HS_DIAGCROSS = 5               '  xxxxx

Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'BLACK_BRUSH：黑色画笔
'DKGRAY_BRUSH：暗灰色画笔
'GRAY_BRUSH：灰色画笔
'HOLLOW_BRUSH：空画笔（相当于HOLLOW_BRUSH）
'LTGRAY_BRUSH：亮灰色画笔
'NULL_BRUSH：空画笔（相当于HOLLOW_BRUSH）
'WHITE_BRUSH：白色画笔
'BLACK_PEN：黑色钢笔
'WHITE_PEN：白色钢笔
Public Const WHITE_BRUSH = 0    '白色画笔
Public Const LTGRAY_BRUSH = 1   '亮灰色画笔
Public Const GRAY_BRUSH = 2     '灰色画笔
Public Const DKGRAY_BRUSH = 3   '暗灰色画笔
Public Const BLACK_BRUSH = 4    '黑色画笔
Public Const NULL_BRUSH = 5
Public Const HOLLOW_BRUSH = NULL_BRUSH
Public Const WHITE_PEN = 6      '白色钢笔
Public Const BLACK_PEN = 7      '黑色钢笔

'创建一个区域
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, _
                            ByVal X3 As Long, ByVal Y3 As Long) As Long

'以下是释放对象函数
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'以下是功能函数
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

Public Const ALTERNATE = 1
Public Const WINDING = 2

Public Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetArcDirection Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function GetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hDC As Long, ByVal uObjectType As Long) As Long
Public Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
                            ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Public Const TRANSPARENT = 1

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
                                  lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_CENTER = &H1

Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, _
                                 ByVal nCount As Long) As Long

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, _
                                    ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, _
                                    ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

Public Declare Function GetUpdateRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

'用指定属性创建一种逻辑字体
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
'获取字体的高度,获取汉字的宽度不准
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, _
                                              ByVal cbString As Long, lpSize As Size) As Long
'nNumber*nNumerator/nDenominator 自动四舍五入。无法计算的返回-1
Public Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Public Const OFFSET_LEFT = 10
Public Const OFFSET_TOP = 20
Public Const OFFSET_RIGHT = 10
Public Const OFFSET_BOTTOM = 20

Private Type Type_Printer
    intPage As Integer
    lngWidth As Long
    lngHeight As Long
    lngLeft As Long
    lngTop As Long
    lngRight As Long
    lngBottom As Long
    intOrient As Integer
    intBin As Integer
End Type
Public gPrinter As Type_Printer


'***************************************************************
'产程图绘画相关变量
'***************************************************************

'产程图页数相关信息
'----------------------------------
Public mintMaxPage As Integer '最大页数
Private mintPageCount As Integer
Private mstrTimeRange()      '每一页的时间范围 格式 开始时间;结束时间(数据发生时间)
Private mArrPageTime()       '每一页的开始时间(计算表格列号使用)
'说明：根据当前页号1、可以获取本页数据时间范围和本页开始时间，2、根据时间范围过滤本页的数据。3、根据本页开始时间结算数据所在列
'------------------------------------
'常量
Private Const gintBmpW    As Integer = 12
Private Const gintBmpH    As Integer = 12
Private Const gintRows    As Integer = 11

Private RGB_BLACK         As Long
Private RGB_RED           As Long
Private RGB_WRITE         As Long
Private RGB_BLUE          As Long
Private RGB_GRAY          As Long
Private RGB_FleetGRAY     As Long
Public sinTwipsPerPixelY As Single
Public sinTwipsPerPixelX As Single
Private msngTwips         As Single
Private mobjDraw          As Object '绘图设备对象
Private gstdset           As New StdFont
Private mblnPrint         As Boolean '是打印预览还是展示
Private msnTimeH          As Single  '产程曲线时间的高度
Private mlngHwnd          As Long
Private mlngDC            As Long
Private mlngMemDC         As Long
Private mlngBitmap        As Long
Private mlngOldBitmap     As Long
Private mlngMemBitmap     As Long
Private mlngPen           As Long
Private mlngBrush         As Long
Private mlngOldPen        As Long
Private mlngOldBrush      As Long
Private mlngFont          As Long
Private mlngOldFont       As Long

Private Type DrawClient
    偏移量 As POINTAPI
    刻度区域 As RECT
    刻度单位 As Long
    产程区域 As RECT
    行单位 As Long
    列单位 As Long
    MaxX As Long
    总行数 As Long
    表格区域 As RECT
End Type
Private T_DrawClient As DrawClient

Private Type type_Patient
    lng文件ID As Long
    lng病人ID As Long
    lng主页ID As Long
    lng科室ID As Long
    lng份数 As Long
    lng页数 As Long
End Type
Private T_Info As type_Patient

'--固定项目序号
Private Type type_PartogramItem
    lng宫口扩大 As Long
    lng先露高低 As Long
    lng生产 As Long
    lng处理 As Long
End Type

Private T_Partogram As type_PartogramItem
'----------------------------------------------------------
'产程图相关变量
'----------------------------------------------------------
Private mrsItems As New ADODB.Recordset '护理项目记录
Private mrsPartogram As New ADODB.Recordset '诊治所见项目
Private mrsDrawItems As New ADODB.Recordset '曲线区域记录集
Private mrsSelItems As New ADODB.Recordset
Private gstrFields As String
Private gstrValues As String
Private mstrCatercorner As String           '列对角线集合
Private mbln日期时间合并 As Boolean         '日期与时间合并
Private mblnDateAd As Boolean               '日期缩写?
Private mstr宫缩时间 As String              '产妇有规律宫缩开始时间
Private mstr开始时间 As String              '当前文件的开始时间
Private mstr结束时间 As String              '当前文件的结束时间
Private mlng格式ID As Long                  '文件格式ID
'病历文件格式定义相关
Private mintTabTiers As Integer     '表头层次
Private mintTagFormHour As Integer  '开始时间条件
Private mintTagToHour As Integer    '截止时间条件
Private mobjTagFont As New StdFont  '条件样式字体
Private mobjSubFont As New StdFont  '上下标签字体
Private mobjTitleFont As New StdFont '标题自理
Private mlngTagColor As Long        '条件样式颜色
Private mstrPaperSet As String      '格式
Private mstrPageFoot As String      '页脚
Private mstrTitle As String         '标题内容
Private mstrSubHead As String       '表上标签
Private mstrSubEnd As String        '表下标签
Private mstrOutSubHead As String    '读取数据后表上标签
Private mstrOutSubEnd As String    '读取数据后表下标签
Private mstrTabHead As String       '表头单元
Private mstrColWidth As String      '列宽序列串
Private mstrColumns As String       '当前护理文件各列对应的项目
Private mlngItems As Long '表格项目数
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
Private mTabForeColor As Long '表格文本颜色
Private mTabGridColor As Long '表格内容颜色
'保存产程记录文件的SQL，在其它地方也有使用，不能修改
Private mstrSQL内 As String
Private mstrSQL中 As String
Private mstrSQL列 As String
Private mstrSQL条件 As String
Private mstrSQL As String


Public Sub ShowPrintPartogram(ByVal objParent As Object, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDtpID As Long, ByVal lngFileIndex As Long, ByVal lngFilePage As Long, Optional ByVal blnPrint As Boolean = True, Optional ByVal strPrintDevice As String = "")
'-----------------------------------------------------------------------------------------------------------
'完成产程图预览、打印
'-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strParam As String, lngFileFormat As Long
    Dim objPrint As Object
    
    On Error GoTo errHand
    
    gstrSQL = "select 格式ID from 病人护理文件 where ID=[1] And 病人ID=[2] And 主页ID=[3] And nvl(婴儿,0)=0"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取格式ID", lngFileID, lngPatiID, lngPageId)
    lngFileFormat = Val(NVL(rsTemp!格式ID))
    
    '提取打印设置信息
    If Not PrintState(lngFileFormat, strPrintDevice) Then Exit Sub
    
    If blnPrint = True Then
        Set objPrint = Printer
    Else
        Set objPrint = frmPartogramPrint
    End If
    
    Load frmPartogramRead
    Call frmPartogramRead.InitRechBox(lngFileFormat)
    
    strParam = lngFileID & ";" & lngPatiID & ";" & lngPageId & ";" & lngDtpID & ";" & lngFileIndex & ";" & lngFilePage
    If blnPrint = True Then Printer.Print ""
    If Not PreViewOrPrintPartogram(strParam, objPrint, objParent, True) Then
        MsgBox "未知错误，打印失败！", vbExclamation, gstrSysName
        GoTo ErrEnd
    End If
    
    If blnPrint = True Then
        Printer.EndDoc
    Else
        '显示预览窗体
        Call frmPartogramPrint.Preview(objParent, lngFileID, lngPatiID, lngPageId, lngDtpID, lngFileIndex, lngFilePage)
    End If
ErrEnd:
    Unload frmPartogramRead
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function PreViewOrPrintPartogram(ByVal strParam As String, ObjDraw As Object, ByVal objParent As Object, Optional ByVal blnPrint As Boolean = False) As Boolean
'---------------------------------------------------------------------------------------------------
'完成产程图数据的展示(展示、预览、打印)
'参数：strPram:文件ID;病人ID;主页ID;科室ID;份数;页数
'      objDraw:展现产程图的对象 (Pictrue?Printer)
'      blnPrint:True 表示打印或预览调用，false仅供展示
'说明：blnPrint=true时，份数=-1表示打印所有文件，页数没用;份数>0时，页数=-1表示打印此份文件，页数>0是表示打印此文件某一页
'      blnPrint=False 时，份数和页数都不能小于0，表示展示某分文件的某一页数据
'---------------------------------------------------------------------------------------------------
    Dim arrParam
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset  '读取产程数据，其他地方不要使用
    Dim blnPrinter As Boolean
    Dim stdSet As StdFont
    Dim lngCurX As Long, lngCurY As Long
    Dim lngMaxRows As Long, lngFileIndexCount As Long
    Dim strPartogram As String '产程曲线项目信息（不能用于其他用途）
    Dim strTmp As String
    
    '预览打印相关
    Dim lngFileCount As Long, lngFileIndex As Long, lngPageIndex As Long, lngPicPageIndex As Long, lng原始页数 As Long, lngPrintPage As Long
    Dim i As Long, j As Long
    Dim dblSureW As Double, dblSureH As Double
    Dim blnReadData As Boolean '是否再次读取数据
    Dim lngLeft As Long, lngRight As Long, lngTop As Long, lngBottom As Long
    Dim lngOffsetLeft As Long, lngOffsetTop As Long
    Dim intFine As Integer, intBold As Integer
    Dim lngHeight As Long
    '打印等待窗体变量
    Dim sngCurOpt As Single, sngScale As Single, sngScaleFileIndex As Single, sngScaleFilePage As Single, lngCountOpt As Long
    Dim strInfo As String
    
    '对象大小计算相关变量
    Dim lngObjHeight As Long, lngobjWidth As Long
    On Error GoTo errHand
    '初始化参数
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then
        MsgBox "请检查传入的参数格式串是否正确！", vbInformation, gstrSysName
        Exit Function
    End If
    T_Info.lng文件ID = Val(arrParam(0))
    T_Info.lng病人ID = Val(arrParam(1))
    T_Info.lng主页ID = Val(arrParam(2))
    T_Info.lng科室ID = Val(arrParam(3))
    T_Info.lng份数 = 1: T_Info.lng页数 = 1
    If UBound(arrParam) > 3 Then T_Info.lng份数 = IIf(Val(arrParam(4)) = 0, 1, Val(arrParam(4)))
    If UBound(arrParam) > 4 Then T_Info.lng页数 = IIf(Val(arrParam(5)) = 0, 1, Val(arrParam(5)))
    
    lngPicPageIndex = 0
    Set mobjDraw = ObjDraw
    If blnPrint = True And Not ISobjPrinter Then
        Set mobjDraw = ObjDraw.picPage(lngPicPageIndex)
    Else
        Set mobjDraw = ObjDraw
    End If
    
    mblnPrint = blnPrint
    blnPrinter = ISobjPrinter
    '一、初始数据--------------------------------------------------------
    
    Screen.MousePointer = 11
    Call ShowFlash(blnPrint, "正在初始化数据,请稍后...", 0.1, objParent)
    
    '1、初始化基础数据
    Call InitEnv
    If Not blnPrinter Then
        sinTwipsPerPixelX = Screen.TwipsPerPixelX
        sinTwipsPerPixelY = Screen.TwipsPerPixelY
        msngTwips = 1
        mobjDraw.Width = Printer.Width
        mobjDraw.Height = Printer.Height
        intBold = 2
        intFine = 1
    Else
        sinTwipsPerPixelX = Printer.TwipsPerPixelX
        sinTwipsPerPixelY = Printer.TwipsPerPixelY
        msngTwips = Screen.TwipsPerPixelX / Printer.TwipsPerPixelX
        intBold = 6
        intFine = 2
    End If
    
    '获取产程固定项目信息
    strPartogram = ""
    mrsItems.Filter = 0
    mrsItems.Filter = "项目名称='宫口扩大' And 保留项目=1"
    T_Partogram.lng宫口扩大 = mrsItems!项目序号
    mrsItems.Filter = "项目名称='先露高低' And 保留项目=1"
    T_Partogram.lng先露高低 = mrsItems!项目序号
    mrsItems.Filter = "项目名称='生产' And 保留项目=1"
    T_Partogram.lng生产 = mrsItems!项目序号
    mrsItems.Filter = "项目名称='处理' And 保留项目=1"
    T_Partogram.lng处理 = mrsItems!项目序号
    
    lngMaxRows = 0
    mstrSQL = "SELECT 记录名,记录符,记录色,最大值,最小值,单位值,单位 FROM 体温记录项目 WHERE 项目序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "产程项目", T_Partogram.lng宫口扩大)
    If rsTemp.RecordCount = 0 Then
        strPartogram = "宫口扩大[LPF]Ο[LPF]255[LPF]10[LPF]0[LPF]1[LPF]CM"
        lngMaxRows = 11
    Else
        strPartogram = NVL(rsTemp!记录名, "宫口扩大") & "[LPF]" & NVL(rsTemp!记录符, "Ο") & "[LPF]" & NVL(rsTemp!记录色, 255) & "[LPF]" & _
            NVL(rsTemp!最大值, 10) & "[LPF]" & NVL(rsTemp!最小值, "0") & "[LPF]" & NVL(rsTemp!单位值, 1) & "[LPF]" & NVL(rsTemp!单位, "CM")
        lngMaxRows = Val(NVL(rsTemp!最大值, 10)) - Val(NVL(rsTemp!最小值, "0"))
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "产程项目", T_Partogram.lng先露高低)
    If rsTemp.RecordCount = 0 Then
        strPartogram = strPartogram & "[|LPF|]" & "先露高低[LPF]×[LPF]10485760[LPF]5[LPF]-5[LPF]1[LPF]CM"
        lngMaxRows = 11
    Else
        strPartogram = strPartogram & "[|LPF|]" & NVL(rsTemp!记录名, "先露高低") & "[LPF]" & NVL(rsTemp!记录符, "×") & "[LPF]" & NVL(rsTemp!记录色, 10485760) & "[LPF]" & _
            NVL(rsTemp!最大值, 5) & "[LPF]" & NVL(rsTemp!最小值, "-5") & "[LPF]" & NVL(rsTemp!单位值, 1) & "[LPF]" & NVL(rsTemp!单位, "CM")
        lngMaxRows = Val(NVL(rsTemp!最大值, 10)) - Val(NVL(rsTemp!最小值, "0"))
    End If
    
    If lngMaxRows <= 0 Then lngMaxRows = gintRows
    
    If blnPrint = False Then
        T_DrawClient.偏移量.X = 10 * msngTwips
        T_DrawClient.偏移量.Y = 10 * msngTwips
    Else
        lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips) / sinTwipsPerPixelX
        lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips) / sinTwipsPerPixelY

        '从打印机上边距和左边距
        lngLeft = (gPrinter.lngLeft * conRatemmToTwip) / sinTwipsPerPixelX + lngOffsetLeft
        lngRight = (gPrinter.lngRight * conRatemmToTwip) / sinTwipsPerPixelX
        lngTop = (gPrinter.lngTop * conRatemmToTwip) / sinTwipsPerPixelY + lngOffsetTop
        lngBottom = (gPrinter.lngBottom * conRatemmToTwip) / sinTwipsPerPixelY + lngOffsetTop

        T_DrawClient.偏移量.X = lngLeft
        T_DrawClient.偏移量.Y = lngTop
    End If
    T_DrawClient.刻度单位 = Format(52 * msngTwips, "0")
    T_DrawClient.列单位 = Format(26 * msngTwips, "0")
    T_DrawClient.行单位 = Format(26 * msngTwips, "0")
    T_DrawClient.MaxX = T_DrawClient.列单位 * 24 + T_DrawClient.刻度单位 + T_DrawClient.偏移量.X
    
    msnTimeH = (mobjDraw.TextHeight("1") / sinTwipsPerPixelY) + (2 * msngTwips)
    '---读取文件构造
    If Not ReadStruDef Then GoTo ErrEnd
    '产程图展示时计算PicTrue大小
    If blnPrint = False Then
        lngObjHeight = T_DrawClient.偏移量.Y
        lngobjWidth = T_DrawClient.偏移量.X
        '宽度
        If Val(zlDatabase.GetPara("先露高低显示位置", glngSys, 1255, 0)) = 0 Then lngobjWidth = lngobjWidth + T_DrawClient.刻度单位
        lngobjWidth = lngobjWidth + T_DrawClient.刻度单位 + T_DrawClient.列单位 * 24
        '高度
        mobjDraw.FontSize = mobjTitleFont.Size
        T_Size.H = mobjDraw.TextHeight(mstrTitle) / sinTwipsPerPixelY
        lngObjHeight = lngObjHeight + (T_Size.H * 2) '标题高度
        mobjDraw.FontSize = mobjSubFont.Size
        lngObjHeight = lngObjHeight + (mobjDraw.TextHeight(mstrSubHead) / sinTwipsPerPixelY) + (2 * msngTwips) '上标签
        lngObjHeight = lngObjHeight + lngMaxRows * T_DrawClient.行单位  '产程曲线部分
        If Val(zlDatabase.GetPara("产程图显示产程时间", glngSys, 1255, 0)) = 1 Then lngObjHeight = lngObjHeight + msnTimeH
        lngObjHeight = lngObjHeight + msnTimeH
        mrsSelItems.Filter = "固定=0" '表格部分
        mrsSelItems.Sort = "行"
        Do While Not mrsSelItems.EOF
            lngObjHeight = lngObjHeight + Val(NVL(mrsSelItems!高度))
            mrsSelItems.MoveNext
        Loop
        mobjDraw.FontSize = mobjSubFont.Size
        lngObjHeight = lngObjHeight + (mobjDraw.TextHeight(mstrSubEnd) / sinTwipsPerPixelY) + (4 * msngTwips) '下标签
        
        mobjDraw.Width = (lngobjWidth + 12) * sinTwipsPerPixelX
        mobjDraw.Height = (lngObjHeight + 12) * sinTwipsPerPixelY
    End If
    Call ShowFlash(blnPrint, "正在初始化数据,请稍后...", 0.2, objParent)
    '---------------------------------------------------------------------------------------------------------------------
    '完成产程图展示、预览以及打印操作
    '---------------------------------------------------------------------------------------------------------------------
    lngFileIndex = T_Info.lng份数
    lngFileCount = lngFileIndex
    lngPicPageIndex = 0
    lngFileIndexCount = GetFileCount(T_Info.lng文件ID, T_Info.lng病人ID, T_Info.lng主页ID)
    If blnPrint = True Then
        If T_Info.lng份数 < 1 Then '表示打印所有文件
            lngFileCount = lngFileIndexCount
            lngFileIndex = 1
            lngPageIndex = 1
            T_Info.lng页数 = -1
        End If
    Else
        If T_Info.lng份数 < 1 Then T_Info.lng份数 = 1
        lngFileIndex = T_Info.lng份数
        lngFileCount = lngFileIndex
    End If
    lng原始页数 = T_Info.lng页数
    
    sngScaleFileIndex = Round(0.8 / (lngFileCount - lngFileIndex + 1), 2)
    For i = lngFileIndex To lngFileCount
        T_Info.lng份数 = i
        blnReadData = True
        Set rsData = New ADODB.Recordset
        Call GetFileProperty
        If T_Info.lng页数 > mintPageCount Then T_Info.lng页数 = mintPageCount
        lngPageIndex = T_Info.lng页数
        If blnPrint = True Then
            If lng原始页数 < 1 Then '打印当前文件
                lngPageIndex = 1
            Else
                mintPageCount = lngPageIndex
            End If
        Else
            If T_Info.lng页数 < 1 Then T_Info.lng页数 = 1
            lngPageIndex = T_Info.lng页数
            mintPageCount = lngPageIndex
        End If
        sngScale = (i - lngFileIndex) * sngScaleFileIndex
        sngScaleFilePage = Round(sngScaleFileIndex / (mintPageCount - lngPageIndex + 1), 2)
        For j = lngPageIndex To mintPageCount
            T_Info.lng页数 = j
            If lngPicPageIndex > 0 Then
                If blnPrint = True Then
                    If Not ISobjPrinter Then
                        Load ObjDraw.picPage(lngPicPageIndex)
                        Set mobjDraw = ObjDraw.picPage(lngPicPageIndex)
                        mobjDraw.Cls
                        mobjDraw.Width = Printer.Width
                        mobjDraw.Height = Printer.Height
                    Else
                        Printer.NewPage
                    End If
                    lngPicPageIndex = lngPicPageIndex + 1
                Else
                    Set mobjDraw = ObjDraw
                End If
            Else
                lngPicPageIndex = 1
            End If
            
            sngScale = sngScale + 0.2
            sngCurOpt = sngScaleFilePage * (i - lngFileIndex + 1)
            sngCurOpt = Round(sngCurOpt / 6, 2)
            
            strInfo = "正在打印第[" & i & "]份文件的第[" & j & "]页,请稍后..."
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 1) + sngScale, objParent)
            '二、画一张没有数据的产程图-----------------------------------------
            mlngDC = mobjDraw.hDC
            
            '展示预览首先清空画布对象
            If Not blnPrinter Then
                Call GetClientRect(mobjDraw.Hwnd, T_ClientRect)      '取得屏幕的有效区域
                '创建白色刷子
                mlngBrush = GetStockObject(WHITE_BRUSH)
                '使用该刷子填充背景色（全白）
                mlngOldBrush = SelectObject(mlngDC, mlngBrush)
                Call FillRect(mlngDC, T_ClientRect, mlngBrush)
                '立即销毁临时使用的刷子并还原刷子
                Call SelectObject(mlngDC, mlngOldBrush)
                Call DeleteObject(mlngBrush)
            End If
            '加载页眉信息
            If blnPrint = True Then Call frmPartogramRead.PrintRTBData(mobjDraw, True, lngTop)
             '提取上下标信息
            If blnReadData = True Then Call GetMarkConnect
            '输出标题信息
            Call SetFontIndirect(mobjTitleFont, mlngDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngDC, mlngFont)
            Call GetTextExtentPoint32(mlngDC, mstrTitle, Len(mstrTitle), T_Size)
            lngCurY = T_Size.H + T_DrawClient.偏移量.Y
            Call GetTextRect(mobjDraw, 0, lngCurY, mstrTitle, mobjDraw.Width / sinTwipsPerPixelX, True, T_Size.H)
            Call DrawText(mlngDC, mstrTitle, -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            lngCurY = lngCurY + T_Size.H
            '输出上标信息 mstrOutSubHead
            lngCurX = T_DrawClient.偏移量.X + T_DrawClient.刻度单位
            mstrOutSubHead = Replace(mstrOutSubHead, "[ZLSOFTLPF]", "  ")
            Call DrawMarkConnect(mstrOutSubHead, lngCurX, lngCurY)
            lngCurY = lngCurY + (mobjDraw.TextHeight(mstrOutSubHead) / sinTwipsPerPixelY) + (2 * msngTwips)
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 2) + sngScale, objParent)
            '画曲线信息
            T_DrawClient.刻度区域.Top = lngCurY
            T_DrawClient.刻度区域.Left = T_DrawClient.偏移量.X
            T_DrawClient.刻度区域.Bottom = T_DrawClient.刻度区域.Top + lngMaxRows * T_DrawClient.行单位
            T_DrawClient.刻度区域.Right = T_DrawClient.刻度单位 + T_DrawClient.刻度区域.Left
            T_DrawClient.产程区域.Left = T_DrawClient.刻度区域.Right
            T_DrawClient.产程区域.Top = T_DrawClient.刻度区域.Top
            T_DrawClient.产程区域.Right = T_DrawClient.产程区域.Left + 24 * T_DrawClient.列单位
            T_DrawClient.产程区域.Bottom = T_DrawClient.刻度区域.Bottom
            T_DrawClient.总行数 = lngMaxRows
            
            lngCurY = DrawPartogram(strPartogram) '画产程刻度区域和曲线区域
            T_DrawClient.表格区域.Top = lngCurY
            T_DrawClient.表格区域.Left = T_DrawClient.刻度区域.Left
            T_DrawClient.表格区域.Right = T_DrawClient.刻度区域.Right
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 3) + sngScale, objParent)
            '完成表格区域的绘图
            lngCurY = DrawPartogramTab
            T_DrawClient.表格区域.Bottom = lngCurY
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 4) + sngScale, objParent)
            '输出下标信息 mstrOutSubend
            mstrOutSubEnd = Replace(mstrOutSubEnd, "[ZLSOFTLPF]", "  ")
            If Trim(Replace(mstrOutSubEnd, Chr(1), "")) <> "" Then '只要表下标签内容不为空，表头名称就为"备注"并话边框线
                lngHeight = (mobjDraw.TextHeight(mstrOutSubEnd) / sinTwipsPerPixelY) + (4 * msngTwips)
                Call DrawLine(mlngDC, T_DrawClient.刻度区域.Left, lngCurY, T_DrawClient.刻度区域.Left, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                Call DrawLine(mlngDC, T_DrawClient.刻度区域.Right, lngCurY, T_DrawClient.刻度区域.Right, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                Call DrawLine(mlngDC, T_DrawClient.MaxX, lngCurY, T_DrawClient.MaxX, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                Call DrawLine(mlngDC, T_DrawClient.刻度区域.Left, lngCurY + lngHeight, T_DrawClient.MaxX, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                '输出备注
                strTmp = "备注"
                strTmp = CheckConnect(strTmp, T_DrawClient.刻度单位, lngHeight)
                lngHeight = (lngHeight - mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY) / 2
                Call GetTextRect(mobjDraw, T_DrawClient.刻度区域.Left, lngCurY + lngHeight, strTmp, T_DrawClient.刻度单位, False)
                Call DrawText(mlngDC, strTmp, -1, T_LableRect, DT_CENTER)
            End If
            
            lngCurX = T_DrawClient.产程区域.Left
            lngCurY = lngCurY + (2 * msngTwips)
            Call DrawMarkConnect(mstrOutSubEnd, lngCurX, lngCurY)
            lngCurY = lngCurY + (mobjDraw.TextHeight(mstrOutSubEnd) / sinTwipsPerPixelY) + (2 * msngTwips)
            '三、提取产程曲线和表格数据，并完成数据的展现----------------------------------
            If blnReadData = True Then
                Call SQLCombination
                Call SQLDIY(mstrSQL)
                Set rsData = zlDatabase.OpenSQLRecord(mstrSQL, "提取数据信息", T_Info.lng文件ID, T_Info.lng病人ID, T_Info.lng主页ID, 0, T_Info.lng份数)
            End If
            blnReadData = False
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 5) + sngScale, objParent)
            '完成曲线数据绘画
            Call DrawPartogramCurData(rsData)
            '完成表格数据的绘画
            Call DrawPartogramTabData(rsData)
            
            '页脚图形输出
            If blnPrint = True Then
                Call frmPartogramRead.PrintRTBData(mobjDraw, False, lngBottom, "第 " & j & " 页" & IIf(lngFileIndexCount > 1, "(" & i & ")", ""), mstrPageFoot)
            End If
            
            Call ShowFlash(blnPrint, strInfo, (sngCurOpt * 6) + sngScale, objParent)
            
            If Not blnPrinter Then mobjDraw.Refresh
            
            If blnPrint = True And Not ISobjPrinter Then
                '如果是打印预览,应按打印机的可打印的开始处开始预览
                dblSureW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
                dblSureH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
                On Error Resume Next
                Call DrawRect(mlngDC, (mobjDraw.Width * dblSureW) / sinTwipsPerPixelX, (mobjDraw.Height * (1 - dblSureH)) / sinTwipsPerPixelY, _
                (mobjDraw.Width * (1 - dblSureW)) / sinTwipsPerPixelX, mobjDraw.Height * dblSureH / sinTwipsPerPixelY, PS_DOT, 1, RGB_FleetGRAY)
            End If
        Next j
    Next i
    PreViewOrPrintPartogram = True
ErrEnd:
    Screen.MousePointer = 0
    Call ShowFlash(blnPrint, "", 1, objParent)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
    Call ShowFlash(blnPrint, "", 1, objParent)
    Call SaveErrLog
End Function

Private Function DrawMarkConnect(ByVal strText As String, ByVal lngX As Long, ByVal lngY As Long) As Long
'---------------------------------------------------
'功能:完成上下标签的输出
'---------------------------------------------------
    Dim strTmp As String, strFomart As String
    Dim intIndex As Integer, i As Integer, j As Integer
    Dim ArrMark, ArrCode
    Dim lngCurX As Long, lngCurY As Long, lngY1 As Long
    Dim lngHeight As Long
    Dim intSize As Single, 记录原始字体信息
    
    If Trim(strText) = "" Then DrawMarkConnect = lngY: Exit Function
    strTmp = strText
    
    '孕周格式固定以上下标的方式输出
    Do While True
        If strTmp Like "孕周*+*" Or strTmp Like "*周+*" Then
            intIndex = InStr(1, strTmp, "+")
            i = intIndex + 1
            For i = intIndex + 1 To Len(strTmp)
                If Not IsNumeric(Mid(strTmp, i, 1)) Then Exit For
            Next i
            
            strFomart = strFomart & Mid(strTmp, 1, intIndex - 1) & "['LPF']" & Mid(strTmp, intIndex, i - intIndex) & "['LPF']"
            strTmp = Mid(strTmp, i)
        Else
            strFomart = strFomart & strTmp
            Exit Do
        End If
    Loop
    
    '设置字体和颜色
    intSize = mobjSubFont.Size
    Call SetTextColor(mlngDC, mTabForeColor)
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    lngHeight = mobjDraw.TextHeight("字") / sinTwipsPerPixelY
    ArrMark = Split(strFomart, vbCrLf)
    For i = 0 To UBound(ArrMark)
        lngCurX = lngX
        lngCurY = lngY + (i * lngHeight)
        lngY1 = lngCurY
        strFomart = CStr(ArrMark(i))
        ArrCode = Split(strFomart, "['LPF']")
        For j = 0 To UBound(ArrCode)
            strTmp = CStr(ArrCode(j))
            lngCurY = lngY1
            If j <> 0 And j < UBound(ArrCode) Then '缩小输出
                mobjSubFont.Size = 7
                lngCurY = lngCurY - (2 * msngTwips)
            Else '正常输出
                mobjSubFont.Size = intSize
            End If
            Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
            Call GetTextRect(mobjDraw, lngCurX, lngCurY, strTmp, 0, False)
            Call DrawText(mlngDC, strTmp, -1, T_LableRect, 0)
            lngCurX = lngCurX + (mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX)
        Next j
    Next i
    
    '最后还原字体信息
    mobjSubFont.Size = intSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    DrawMarkConnect = lngY1
End Function

Private Sub DrawPartogramCurData(ByVal rsTemp As ADODB.Recordset)
'----------------------------------------------------------------------
'功能:完成曲线区域数据的展示
'----------------------------------------------------------------------
    Dim rsCurData As New ADODB.Recordset
    Dim arrItemNo
    Dim lngOrder As Long, strOrder As String
    Dim sngX As Single, sngY As Single
    Dim sngOutX As Single, sngOutY As Single
    Dim i As Integer, j As Integer
    Dim strBeginDate As String, strFiled As String, strValue As String, strTime As String, strContent As String
    
    '存放分娩标记内容信息
    Dim rsCurInfo As New ADODB.Recordset
    Dim rsCopyCurInfo As New ADODB.Recordset 'rsCurInfo的副本
    Dim strFields As String, strValues As String, strFiled1 As String
    
    '画图变量
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lng项目序号 As Long, sin原X As Single, sin原Y As Single
    Dim lngRGB As Long, strICon As String
    Dim blnLine As Boolean '输出警戒线（1、连点之间连线通过3CM。2、存在3Cm的点）
    Dim sin3CmY As Single, sin3CmX As Single, intState As Integer
    Dim sin10CmY As Single, sin10CmX As Single
    Dim lngType As Long  '线条样式，是实线还是虚线
    '参数信息
    Dim bln显示警戒线 As Boolean, int警戒线条 As Integer, int异常线条 As Integer
    Dim int宫口曲线标志S As Integer, int先露曲线标志S As Integer '顺产
    Dim int宫口曲线标志Y As Integer, int先露曲线标志Y As Integer '异常产
    Dim int宫口曲线标志 As Integer, int先露曲线标志 As Integer
    
    Dim int标志内容 As Integer, int标志位置 As Integer
    Dim int零点连线 As Integer, blnZeroLine As Boolean  '第一曲线点与0点连线:0-不连线,1-虚线,2-实线
    Dim sinZeroX As Single, sinZeroY As Single
    Dim blnAbnormal As Boolean '异常产
    Dim strTmp As String
    
    On Error GoTo errHand
    
    '曲线数据过滤当天的数据
    strBeginDate = Format(DateAdd("d", IIf((T_Info.lng页数 - 1) < 0, 0, T_Info.lng页数 - 1), mstr宫缩时间), "YYYY-MM-DD HH:mm:ss")
    rsTemp.Filter = "发生时间>='" & strBeginDate & "' And 发生时间<='" & Format(DateAdd("d", 1, CDate(strBeginDate)), "YYYY-MM-DD HH:mm:ss") & "'"
    If rsTemp.RecordCount = 0 Then Exit Sub
    
    '-----读取参数信息
     '产程生产曲线标志
    strTmp = zlDatabase.GetPara("产程生产曲线标志", glngSys, 1255, "1;1", , True)
    int宫口曲线标志S = Val(Split(strTmp, ";")(0))
    int先露曲线标志S = Val(Split(strTmp, ";")(1))
    strTmp = int宫口曲线标志S & ";" & int先露曲线标志S
    strTmp = zlDatabase.GetPara("产程生产曲线标志(异)", glngSys, 1255, strTmp, , True)
    int宫口曲线标志Y = Val(Split(strTmp, ";")(0))
    int先露曲线标志Y = Val(Split(strTmp, ";")(1))
    
    '产程生产措施标志
    strTmp = zlDatabase.GetPara("产程生产措施标志", glngSys, 1255, "1;1", , True)
    int标志内容 = Val(Split(strTmp, ";")(0))
    int标志位置 = Val(Split(strTmp, ";")(1))
    
    '产程警戒线标志
    strTmp = zlDatabase.GetPara("产程警戒异常线标志", glngSys, 1255, "1;1", , True)
    int警戒线条 = Val(Split(strTmp, ";")(0))
    int异常线条 = Val(Split(strTmp, ";")(1))
    bln显示警戒线 = (Val(zlDatabase.GetPara("产程图显示警戒线", glngSys, 1255, "1", , True)) = 1)
    int零点连线 = Val(zlDatabase.GetPara("产程曲线点与0点连线", glngSys, 1255, "0", , True))
    If int零点连线 < 0 Or int零点连线 > 2 Then int零点连线 = 0
    
    '处理曲线记录集
    gstrFields = "项目序号," & adDouble & ",18|内容," & adLongVarChar & ",1000|时间," & adLongVarChar & ",20|X坐标," & adDouble & ",5|Y坐标," & adDouble & ",5"
    Call Record_Init(rsCurData, gstrFields)
    gstrFields = "项目序号|内容|时间|X坐标|Y坐标"
    
    strFields = "项目序号," & adDouble & ",18|数值," & adLongVarChar & ",100|内容," & adLongVarChar & ",1000|时间," & adLongVarChar & ",20|X坐标," & adDouble & ",5|Y坐标," & adDouble & ",5|" & _
        "打印X坐标," & adDouble & ",5|打印Y坐标," & adDouble & ",5|模式," & adInteger & ",1|宽度," & adDouble & ",18|高度," & adDouble & ",18"
    Call Record_Init(rsCurInfo, strFields)
    Call Record_Init(rsCopyCurInfo, strFields)
    strFields = "项目序号|数值|内容|时间|X坐标"
     
    '---先提取宫口扩大、先露下降、生产信息
    strOrder = T_Partogram.lng宫口扩大 & ";" & T_Partogram.lng先露高低 & ";" & T_Partogram.lng生产
    arrItemNo = Split(strOrder, ";")
    For i = 0 To UBound(arrItemNo)
        mrsSelItems.Filter = 0
        mrsSelItems.Filter = "行=" & Val(arrItemNo(i)) & " And 固定=1"
        If mrsSelItems.RecordCount > 0 Then
            lngOrder = Val(mrsSelItems!对象序号)
            strFiled = "C" & Format(lngOrder, "00")
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                strValue = Trim(NVL(rsTemp.Fields(strFiled).Value))
                If strValue <> "" Then
                    If Val(arrItemNo(i)) <> T_Partogram.lng生产 Then
                        '计算X、Y坐标
                        strTime = Format(rsTemp!发生时间, "YYYY-MM-DD HH:mm:ss")
                        sngX = GetXCoordinate(strTime, strBeginDate)
                        sngY = GetYCoordinate(Val(arrItemNo(i)), strValue)
                        '添加记录集
                        If (sngX <= T_DrawClient.MaxX And sngX >= T_DrawClient.产程区域.Left) And (sngY >= T_DrawClient.产程区域.Top And sngY <= T_DrawClient.产程区域.Bottom) Then
                            gstrValues = Val(arrItemNo(i)) & "|" & strValue & "|" & strTime & "|" & sngX & "|" & sngY
                            Call Record_Add(rsCurData, gstrFields, gstrValues)
                        End If
                    Else '存放生产内容和标记
                        '获取处理的字段名
                        If Mid(strValue, 1, 1) = "√" Then
                            strContent = ""
                            Select Case int标志内容
                                Case 1
                                    strContent = "生产"
                                Case 2
                                    mrsSelItems.Filter = 0
                                    mrsSelItems.Filter = "行=" & T_Partogram.lng处理 & " And 固定=1"
                                    If mrsSelItems.RecordCount > 0 Then
                                        strFiled1 = "C" & Format(Val(mrsSelItems!对象序号), "00")
                                        strContent = Trim(NVL(rsTemp.Fields(strFiled1).Value))
                                    End If
                            End Select
                            strTime = Format(rsTemp!发生时间, "YYYY-MM-DD HH:mm:ss")
                            sngX = GetXCoordinate(strTime, strBeginDate)
                            If (sngX <= T_DrawClient.MaxX And sngX >= T_DrawClient.产程区域.Left) Then
                                strValues = Val(arrItemNo(i)) & "|" & strValue & "|" & strContent & "|" & strTime & "|" & sngX
                                Call Record_Add(rsCurInfo, strFields, strValues)
                            End If
                        End If
                    End If
                End If
                rsTemp.MoveNext
            Loop
        End If
    Next i
    
    '开始进行连线操作
    blnPrinter = ISobjPrinter
    If blnPrinter = True Then
        intBold = 4
        intFine = 4
    Else
        intBold = 2
        intFine = 1
    End If
    
    '获取宫口扩大3Cm的坐标
    sin3CmY = GetYCoordinate(T_Partogram.lng宫口扩大, 3)
    
    rsCurData.Filter = ""
    rsCurData.Sort = "项目序号,时间"
    '----完成点与点之间的连线和符号的输出
    blnLine = False
    blnZeroLine = False
    lng项目序号 = -999
    sin3CmX = 0: sin原X = 0: sin原Y = 0
    With rsCurData
        Do While Not .EOF
            blnAbnormal = False
            If NVL(!项目序号) <> lng项目序号 Then
                blnZeroLine = True
                sin原X = 0
                sin原Y = 0
                lng项目序号 = NVL(!项目序号)
                lngRGB = Val(GetDrawItemValue(lng项目序号, "颜色"))
                strICon = GetDrawItemValue(lng项目序号, "记录符")
                intState = GetDrawItemValue(lng项目序号, "显示模式")
            End If
            
            If sin原X <> 0 Then
                '重庆九院需求：异常产先露高低从上一点到生产点画直角虚线
                '情况一：生产是同时录入了先露高低的情况
                If lng项目序号 = T_Partogram.lng先露高低 Then
                    rsCurInfo.Filter = "X坐标=" & !x坐标
                    If rsCurInfo.RecordCount > 0 Then
                        blnAbnormal = (rsCurInfo!数值 = "√(异)")
                    End If
                End If
                If blnAbnormal = True And int先露曲线标志Y = 3 Then
                    Call DrawLine(mlngDC, sin原X, sin原Y, !x坐标, sin原Y, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                    Call DrawLine(mlngDC, !x坐标, sin原Y, !x坐标, !Y坐标, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                Else
                    Call DrawLine(mlngDC, sin原X, sin原Y, !x坐标, !Y坐标, PS_SOLID, intFine, lngRGB)
                End If
                '判断是否空口开大3Cm或两点之间的连线划过3Cm
                If blnLine = False And lng项目序号 = T_Partogram.lng宫口扩大 Then
                    If intState = 0 Then
                        If sin3CmY < sin原Y And sin3CmY > !Y坐标 Then blnLine = True
                    Else
                        If sin3CmY > sin原Y And sin3CmY < !Y坐标 Then blnLine = True
                    End If
                    '计算经过3Cm处的点
                    If blnLine = True Then
                        sin3CmX = ((!x坐标 - sin原X) * (sin3CmY - sin原Y) / (!Y坐标 - sin原Y)) + sin原X
                    End If
                End If
            End If
            sin原X = NVL(!x坐标, 0)
            sin原Y = NVL(!Y坐标, 0)
            If blnLine = False And lng项目序号 = T_Partogram.lng宫口扩大 Then
                If sin3CmY = sin原Y Then blnLine = True: sin3CmX = sin原X
            End If
            '输出图标
            Set gstdset = New StdFont
            gstdset.Name = "宋体"
            gstdset.Size = 9
            gstdset.Underline = False
            gstdset.Italic = False
            Call SetFontIndirect(gstdset, mlngDC, mobjDraw)
            Call SetTextColor(mlngDC, lngRGB)
            Call GetTextRect(mobjDraw, sin原X - (mobjDraw.TextWidth(strICon) / sinTwipsPerPixelX / 2), sin原Y, Trim(strICon))
            Call DrawText(mlngDC, Trim(strICon), -1, T_LableRect, DT_CENTER)
            If int零点连线 > 0 And blnZeroLine = True And sin原X > T_DrawClient.产程区域.Left Then
                sinZeroX = T_DrawClient.产程区域.Left
                mrsDrawItems.Filter = "项目序号=" & lng项目序号
                If mrsDrawItems.RecordCount > 0 Then
                    sinZeroY = GetYCoordinate(lng项目序号, mrsDrawItems!最小值)
                Else
                    sinZeroY = 0
                End If
                If int零点连线 = 1 Then '虚线
                    Call DrawLine(mlngDC, sinZeroX, sinZeroY, sin原X, sin原Y, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                Else '实线
                    Call DrawLine(mlngDC, sinZeroX, sinZeroY, sin原X, sin原Y, PS_SOLID, intFine, lngRGB)
                End If
                blnZeroLine = False
            End If
        .MoveNext
        Loop
    End With
    
    '-----画警戒线和异常线
    If blnLine = True And bln显示警戒线 Then
        If sin3CmX > 0 And sin3CmX < T_DrawClient.MaxX Then
            sin10CmX = sin3CmX + (T_DrawClient.列单位 * 4)
            sin10CmY = GetYCoordinate(T_Partogram.lng宫口扩大, Val(GetDrawItemValue(T_Partogram.lng宫口扩大, "最大值")))
            If sin10CmX > T_DrawClient.MaxX Then GoTo ErrNext
            '画警戒线
            Select Case int警戒线条
                Case 0
                    lngType = IIf(blnPrinter, PS_DASH, PS_DOT)
                Case Else
                    lngType = PS_SOLID
            End Select
            Call DrawLine(mlngDC, sin3CmX, sin3CmY, sin10CmX, sin10CmY, lngType, intFine, RGB_RED)
            '画异常线
            If sin10CmX + (T_DrawClient.列单位 * 4) > T_DrawClient.MaxX Then GoTo ErrNext
             Select Case int异常线条
                Case 0
                    lngType = IIf(blnPrinter, PS_DASH, PS_DOT)
                Case Else
                    lngType = PS_SOLID
            End Select
            Call DrawLine(mlngDC, sin10CmX, sin3CmY, sin10CmX + (T_DrawClient.列单位 * 4), sin10CmY, lngType, intFine, RGB_RED)
        End If
    End If
ErrNext:
    '----进行病人分娩处理
    Dim sinUpY As Single, sinUpX As Single '处理文本坐标使用，其它地方不要使用
    strTmp = ""
    rsCurInfo.Filter = ""
    rsCurInfo.Sort = "时间"
    With rsCurInfo  '"项目序号|内容|时间|X坐标"
        Do While Not .EOF
            sngX = !x坐标
            blnAbnormal = (!数值 = "√(异)")
            If blnAbnormal = True Then
                int宫口曲线标志 = int宫口曲线标志Y
                int先露曲线标志 = int先露曲线标志Y
            Else
                int宫口曲线标志 = int宫口曲线标志S
                int先露曲线标志 = int先露曲线标志S
            End If
            
            If int宫口曲线标志 > 0 Then
                rsCurData.Filter = ""
                rsCurData.Filter = "项目序号=" & T_Partogram.lng宫口扩大 & " And X坐标<=" & sngX
                rsCurData.Sort = "时间 DESC"
                If rsCurData.RecordCount > 0 Then
                    sin原X = rsCurData!x坐标
                    sin原Y = rsCurData!Y坐标
                    lngRGB = Val(GetDrawItemValue(T_Partogram.lng宫口扩大, "颜色"))
                    intState = Val(GetDrawItemValue(T_Partogram.lng宫口扩大, "显示模式"))
                    '显示在宫口扩大时，生产的Y坐标和最近的宫口Y坐标相同
                    Call DrawLine(mlngDC, sin原X, sin原Y, sngX, sin原Y, PS_SOLID, intFine, lngRGB)
                    
                    If intState = 0 Then
                        sin3CmY = sin原Y + T_DrawClient.行单位
                    Else
                        sin3CmY = sin原Y - T_DrawClient.行单位
                    End If
                    If int宫口曲线标志 = 1 Then '显示虚线箭头
                        Call DrawLine(mlngDC, sngX, sin原Y, sngX, sin3CmY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB, True)
                    Else '显示实线箭头
                        Call DrawLine(mlngDC, sngX, sin原Y, sngX, sin3CmY, PS_SOLID, intFine, lngRGB, True)
                    End If
                    If int标志位置 = 0 Then
                       !Y坐标 = sin原Y
                       !打印Y坐标 = sin原Y
                       .Update
                    End If
                End If
            End If
            
            If int先露曲线标志 > 0 Then
                rsCurData.Filter = ""
                rsCurData.Filter = "项目序号=" & T_Partogram.lng先露高低 & " And X坐标<=" & sngX
                rsCurData.Sort = "时间 DESC"
                If rsCurData.RecordCount > 0 Then
                    sin原X = rsCurData!x坐标
                    sin原Y = rsCurData!Y坐标
                    sin10CmY = GetYCoordinate(T_Partogram.lng先露高低, Val(GetDrawItemValue(T_Partogram.lng先露高低, "最大值")))
                    lngRGB = Val(GetDrawItemValue(T_Partogram.lng先露高低, "颜色"))
                    intState = Val(GetDrawItemValue(T_Partogram.lng先露高低, "显示模式"))
                    '显示在宫口扩大时，生产的Y坐标和最近的宫口Y坐标相同
                    If sngX = sin原X Then
                        sin10CmY = sin原Y
                    Else
                        If blnAbnormal = True And int先露曲线标志 = 3 Then
                            Call DrawLine(mlngDC, sin原X, sin原Y, sngX, sin原Y, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                            Call DrawLine(mlngDC, sngX, sin原Y, sngX, sin10CmY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB)
                        Else
                            Call DrawLine(mlngDC, sin原X, sin原Y, sngX, sin10CmY, PS_SOLID, intFine, lngRGB)
                        End If
                    End If
                    If intState = 0 Then
                        sin3CmY = sin10CmY - (T_DrawClient.行单位 / 2)
                        If sin10CmY = T_DrawClient.刻度区域.Top Then
                            sin3CmY = sin10CmY - (T_DrawClient.行单位 / 2)
                        ElseIf sin10CmY - T_DrawClient.行单位 <= T_DrawClient.刻度区域.Top Then
                            sin3CmY = sin10CmY - T_DrawClient.行单位
                        ElseIf sin10CmY - T_DrawClient.刻度区域.Top >= (T_DrawClient.行单位 / 2) Then
                            sin3CmY = T_DrawClient.刻度区域.Top
                        End If
                    Else
                        sin3CmY = sin10CmY + (T_DrawClient.行单位 / 2)
                        If sin10CmY = T_DrawClient.刻度区域.Bottom Then
                            sin3CmY = sin10CmY + (T_DrawClient.行单位 / 2)
                        ElseIf sin10CmY + T_DrawClient.行单位 <= T_DrawClient.刻度区域.Bottom Then
                            sin3CmY = sin10CmY + T_DrawClient.行单位
                        ElseIf T_DrawClient.刻度区域.Bottom - sin10CmY >= (T_DrawClient.行单位 / 2) Then
                            sin3CmY = T_DrawClient.刻度区域.Bottom
                        End If
                    End If
                    If int先露曲线标志 = 1 Then '显示虚线箭头
                        Call DrawLine(mlngDC, sngX, sin10CmY, sngX, sin3CmY, IIf(blnPrinter, PS_DASH, PS_DOT), 1, lngRGB, True)
                    ElseIf int先露曲线标志 = 2 Then '显示实线箭头
                        Call DrawLine(mlngDC, sngX, sin10CmY, sngX, sin3CmY, PS_SOLID, intFine, lngRGB, True)
                    End If
                    '文本标志位置
                    If int标志位置 = 1 Then
                       !Y坐标 = sin原Y
                       !打印Y坐标 = sin原Y
                       .Update
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '----开始计算输出文本描述信息
    If int标志位置 = 0 Then
        lngRGB = Val(GetDrawItemValue(T_Partogram.lng宫口扩大, "颜色"))
        intState = Val(GetDrawItemValue(T_Partogram.lng宫口扩大, "显示模式"))
    Else
        lngRGB = Val(GetDrawItemValue(T_Partogram.lng先露高低, "颜色"))
        intState = Val(GetDrawItemValue(T_Partogram.lng先露高低, "显示模式"))
    End If
    Dim int模式 As Integer, blnInit As Boolean, blnEnd As Boolean
    Dim lngMax As Single, lngMaxY As Single, sng行差 As Single
    Dim lngCurWidth As Long, lngCurHeight As Long
    
    sng行差 = 4 * msngTwips
    
    '复制rsCurInfo
    rsCurInfo.Filter = ""
    Do While Not rsCurInfo.EOF
        rsCopyCurInfo.AddNew
        For j = 0 To rsCurInfo.Fields.Count - 1
            rsCopyCurInfo.Fields(j).Value = IIf(NVL(rsCurInfo.Fields(j).Value) = "", Null, rsCurInfo.Fields(j).Value)
        Next j
        rsCopyCurInfo.Update
    rsCurInfo.MoveNext
    Loop
    
    T_Size.W = mobjDraw.TextWidth("字") / sinTwipsPerPixelX
    T_Size.H = mobjDraw.TextHeight("字") / sinTwipsPerPixelY
    lngMax = T_DrawClient.MaxX
    lngMaxY = T_DrawClient.产程区域.Bottom
    '重新整理内容的X坐标和输出模式(纵向输出还是横向输出)
    rsCopyCurInfo.Filter = ""
    rsCurInfo.Filter = ""
    rsCurInfo.Sort = "X坐标 DESC"
    With rsCurInfo
        Do While Not .EOF
            If Val(NVL(!Y坐标, 0)) >= T_DrawClient.刻度区域.Top And Val(NVL(!Y坐标, 0)) <= T_DrawClient.刻度区域.Bottom Then
                strTmp = NVL(!内容)
                If .AbsolutePosition = 1 Then
                    sngX = Val(NVL(!x坐标))
                Else
                    If sngX < Format(Val(NVL(!x坐标)) + T_Size.W + sng行差, "0.0") Then
                        sngX = Format(sngOutX - T_Size.W - sng行差, "0.0")
                    Else
                        sngX = Val(NVL(!x坐标))
                    End If
                End If
                sngY = Val(NVL(!Y坐标))
                If CInt(lngMax - sngX) < CInt(T_DrawClient.列单位 * 2) Then '纵向输出文本信息
                    int模式 = 1
                    If CInt(sngX + T_Size.W + sng行差) > CInt(lngMax) Then
                        rsCopyCurInfo.Filter = "X坐标<" & sngX
                        rsCopyCurInfo.Sort = "X坐标 DESC"
                        If rsCopyCurInfo.RecordCount > 0 Then
                            If CInt(sngX - Val(NVL(rsCopyCurInfo!x坐标))) < CInt(T_Size.W + sng行差) Then
                                sngX = Format(lngMax - T_Size.W - sng行差, "0.0")
                            Else
                                sngX = Format(sngX - T_Size.W - sng行差, "0.0")
                            End If
                        Else
                            sngX = Format(sngX - T_Size.W - sng行差, "0.0")
                        End If
                    End If
                   
                    If intState = 1 Then
                         If CInt(lngMaxY - sngY) < CInt(mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX) Then
                            sngY = Format(lngMaxY - (mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX) - T_Size.H, "0.0")
                         End If
                    End If
                    If sngY < T_DrawClient.刻度区域.Top Then sngY = T_DrawClient.刻度区域.Top
                    
                    lngCurWidth = lngMax - sngX
                    lngCurHeight = T_DrawClient.产程区域.Bottom - sngY
                    lngMax = sngX
                    sngOutX = sngX
                Else '横向输出文本信息
                    int模式 = 0
                    lngCurWidth = lngMax - sngX
                    strTmp = CheckConnect(strTmp, lngCurWidth, T_DrawClient.产程区域.Bottom - sngY)
                    lngCurHeight = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
                    If CInt(lngCurHeight + sngY + sng行差) > CInt(T_DrawClient.产程区域.Bottom) Then
                        sngY = Format(T_DrawClient.产程区域.Bottom - lngCurHeight - sng行差, "0.0")
                    End If
                    rsCopyCurInfo.Filter = "X坐标<" & sngX
                    rsCopyCurInfo.Sort = "X坐标 DESC"
                    If rsCopyCurInfo.RecordCount > 0 Then
                        If CInt(sngX - Val(NVL(rsCopyCurInfo!x坐标))) < CInt(T_Size.W + sng行差) Then
                            sngOutX = Format(sngX + T_Size.W + sng行差, "0.0")
                            sngX = Format(Val(NVL(rsCopyCurInfo!x坐标)) + T_Size.W + sng行差 - (1 * msngTwips), "0.0")
                        Else
                            lngMax = sngX
                            sngOutX = sngX
                        End If
                    Else
                        lngMax = sngX
                        sngOutX = sngX
                    End If
                End If
                !高度 = lngCurHeight
                !宽度 = lngCurWidth
                !打印X坐标 = Int(sngX)
                !打印Y坐标 = Int(sngY)
                !模式 = int模式
                .Update
            End If
        .MoveNext
        Loop
    End With
    
    '重新复制rsCurInfo记录集
    rsCopyCurInfo.Filter = ""
    Do While Not rsCopyCurInfo.EOF
        rsCopyCurInfo.Delete
        rsCopyCurInfo.Update
    rsCopyCurInfo.MoveNext
    Loop
    
    rsCurInfo.Filter = ""
    Do While Not rsCurInfo.EOF
        rsCopyCurInfo.AddNew
        For j = 0 To rsCurInfo.Fields.Count - 1
            rsCopyCurInfo.Fields(j).Value = IIf(NVL(rsCurInfo.Fields(j).Value) = "", Null, rsCurInfo.Fields(j).Value)
        Next j
        rsCopyCurInfo.Update
    rsCurInfo.MoveNext
    Loop
    
'    rsCurInfo.Filter = ""
'    Call OutputRsData(rsCurInfo, True)
    
    blnInit = False
    blnEnd = False
    '重新整理横向输出内容打印Y坐标
    rsCopyCurInfo.Filter = "模式=0"
    rsCopyCurInfo.Sort = "打印X坐标"
    i = 1
    With rsCopyCurInfo
        Do While Not .EOF
            If blnInit = False Then sngX = Val(NVL(!打印X坐标)): lngCurHeight = 0: blnInit = True
            If Val(NVL(!打印X坐标)) <> sngX Then
ErrEnd:
                rsCurInfo.Filter = "打印X坐标=" & sngX
                rsCurInfo.Sort = "打印Y坐标"
                Do While Not rsCurInfo.EOF
                    sngY = Val(NVL(rsCurInfo!打印Y坐标)) + IIf(rsCurInfo.AbsolutePosition > 1, lngCurHeight + sng行差, 0)
                    sngOutY = lngMaxY - lngCurHeight - sng行差
                    If sngOutY < T_DrawClient.产程区域.Top + sng行差 Then sngOutY = T_DrawClient.产程区域.Top + sng行差
                    If sngY > sngOutY Then sngY = sngOutY
                    If Val(NVL(rsCurInfo!高度)) > lngMaxY - sngY Then rsCurInfo!高度 = Format(lngMaxY - sngY, "0.00")
                    rsCurInfo!打印Y坐标 = sngY
                    rsCurInfo.Update
                    
                    lngCurHeight = lngCurHeight - Val(NVL(rsCurInfo!高度))
                rsCurInfo.MoveNext
                Loop
            
                sngX = Val(NVL(!打印X坐标)): lngCurHeight = Val(NVL(!高度))
                If blnEnd = True Then GoTo ErrOutPut
            Else
                lngCurHeight = lngCurHeight + Val(NVL(!高度))
            End If
            If i = rsCopyCurInfo.RecordCount Then blnEnd = True: GoTo ErrEnd
            i = i + 1
        .MoveNext
        Loop
    End With
ErrOutPut:
    '正式输出文本信息
    Call SetTextColor(mlngDC, lngRGB)
    rsCurInfo.Filter = ""
    rsCurInfo.Sort = "X坐标"
    With rsCurInfo
        Do While Not .EOF
            If Val(NVL(!Y坐标, 0)) >= T_DrawClient.刻度区域.Top And Val(NVL(!Y坐标, 0)) <= T_DrawClient.刻度区域.Bottom Then
                strTmp = NVL(!内容)
                If Val(NVL(!模式)) = 0 Then
                    Call OutBigHConnect(strTmp, !打印X坐标, !打印Y坐标, !高度, !宽度, False, False)
                Else
                    Call OutBigVConnect(strTmp, !打印X坐标, !打印Y坐标, !高度, !宽度, False, False, lngRGB)
                End If
            End If
        .MoveNext
        Loop
    End With

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawPartogramTabData(ByVal rsTemp As ADODB.Recordset)
'----------------------------------------------------------------------
'功能:完成表格区域数据的展示
'----------------------------------------------------------------------
    Dim strBegin As String, strEnd As String, strTmp As String, strPageTime As String
    Dim rsTabData As New ADODB.Recordset
    Dim lngCol As Long, lngColOld As Long, blnInit As Boolean
    Dim intFields As Integer, lngOrder As Long
    '---表格信息
    Dim arrItemOrder, arrItemName, arrItemHeight
    Dim lngCurX As Long, lngCurY As Long, lngHeight As Long
    Dim strName As String
    Dim strLeft As String, strRight As String
    Dim intBold As Integer, intFine As Integer
    On Error GoTo errHand
    
    If ISobjPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    
    '表格数据根据页数获取当前页数据
    rsTemp.Filter = 0
    strTmp = mstrTimeRange(T_Info.lng页数 - 1) '当前页数据范围
    strPageTime = mArrPageTime(T_Info.lng页数 - 1) '当前页开始时间
    strBegin = Format(Split(strTmp, ";")(0), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Split(strTmp, ";")(1), "YYYY-MM-DD HH:mm:ss")
    rsTemp.Filter = "发生时间>='" & strBegin & "' And 发生时间<='" & strEnd & "'"
    If rsTemp.RecordCount = 0 Then Exit Sub
    rsTemp.Sort = "发生时间"
    '复制记录集字段
    Set rsTabData = CopyNewRec(rsTemp)
    gstrFields = ""
    For intFields = 0 To rsTabData.Fields.Count - 1
        gstrFields = gstrFields & "|" & rsTabData.Fields(intFields).Name
    Next intFields
    gstrFields = Mid(gstrFields, 2)
    '计算数据在哪一行
    With rsTemp
        Do While Not .EOF
            gstrValues = ""
            lngCol = Int((CDate(Format(!发生时间, "YYYY-MM-DD HH:mm:ss")) - CDate(Format(strPageTime, "YYYY-MM-DD HH:mm:ss"))) * 24) + 1
            For intFields = 0 To rsTemp.Fields.Count - 1
                gstrValues = gstrValues & "|" & rsTemp.Fields(intFields).Value
            Next intFields
            gstrValues = Mid(gstrValues, 2) & "|" & lngCol
            Call Record_Update(rsTabData, gstrFields, gstrValues, "发生时间|" & Format(!发生时间, "YYYY-MM-DD HH:mm:ss"))
        .MoveNext
        Loop
    End With
    
    '重新整理列号
    blnInit = False: lngCol = 0: lngColOld = 0
    rsTabData.Filter = ""
    rsTabData.Sort = "发生时间"
    With rsTabData
        Do While Not .EOF
            If lngCol <> Val(!列号) Or blnInit = False Then
                lngCol = Val(!列号)
                '73792:刘鹏飞,2014-06-23,表格行位置计算调整
                If lngColOld - lngCol >= 0 Then
                    lngColOld = lngColOld + 1 ' (2 * lngColOld) - lngCol + 1
                Else
                    lngColOld = lngCol
                End If
            Else
                lngColOld = lngColOld + 1
            End If
            blnInit = True
            If lngColOld <> lngCol Then
                Call Record_Update(rsTabData, "发生时间|列号", Format(!发生时间, "YYYY-MM-DD HH:mm:ss") & "|" & lngColOld, "发生时间|" & Format(!发生时间, "YYYY-MM-DD HH:mm:ss"))
            End If
        .MoveNext
        Loop
    End With
    
    arrItemOrder = Array()
    arrItemName = Array()
    arrItemHeight = Array()
    mrsSelItems.Filter = ""
    mrsSelItems.Filter = "固定=0"
    mrsSelItems.Sort = "行"
    Do While Not mrsSelItems.EOF
        ReDim Preserve arrItemOrder(UBound(arrItemOrder) + 1)
        ReDim Preserve arrItemName(UBound(arrItemName) + 1)
        ReDim Preserve arrItemHeight(UBound(arrItemHeight) + 1)
        lngOrder = Val(mrsSelItems!对象序号)
        arrItemOrder(UBound(arrItemOrder)) = lngOrder & ";" & NVL(mrsSelItems!要素名称)
        arrItemName(UBound(arrItemName)) = "C" & Format(lngOrder, "00")
        arrItemHeight(UBound(arrItemHeight)) = Val(mrsSelItems!高度)
    mrsSelItems.MoveNext
    Loop
    '开始完成表格内容输出
    rsTabData.Filter = ""
    rsTabData.Sort = "发生时间"
    With rsTabData
        Call SetTextColor(mlngDC, mTabForeColor)
        Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
        Do While Not .EOF
            lngHeight = 0
            For intFields = 0 To UBound(arrItemOrder)
                strTmp = rsTabData.Fields(CStr(arrItemName(intFields)))
                lngOrder = Val(Split(CStr(arrItemOrder(intFields)), ";")(0))
                strName = Split(CStr(arrItemOrder(intFields)), ";")(1)
                lngCol = rsTabData.Fields("列号")
                lngCurX = T_DrawClient.产程区域.Left + (lngCol - 1) * T_DrawClient.列单位
                lngCurY = T_DrawClient.表格区域.Top + lngHeight
                '有对角线的数据
                If IsDiagonal(lngOrder) And InStr(1, strTmp, "/") <> 0 Then
                    strLeft = Split(strTmp, "/")(0)
                    strRight = Mid(strTmp, InStr(1, strTmp, "/") + 1)
                    '画对角线
                    Call DrawLine(mlngDC, lngCurX, lngCurY + Val(arrItemHeight(intFields)), lngCurX + T_DrawClient.列单位, lngCurY, PS_SOLID, intFine, mTabGridColor)
                    '输出文本
                    Call GetTextRect(mobjDraw, lngCurX, lngCurY, strLeft, 0, False)
                    Call DrawText(mlngDC, strLeft, -1, T_LableRect, 0)
                    T_Size.H = mobjDraw.TextHeight(strRight) / sinTwipsPerPixelY
                    T_Size.W = mobjDraw.TextWidth(strRight) / sinTwipsPerPixelY
                    Call GetTextRect(mobjDraw, IIf(T_DrawClient.列单位 - T_Size.W > 0, lngCurX + T_DrawClient.列单位 - T_Size.W, lngCurX), lngCurY + Val(arrItemHeight(intFields)) - T_Size.H, strRight, 0, False)
                    Call DrawText(mlngDC, strRight, -1, T_LableRect, 0)
                ElseIf isBigConnect(strName, 1) = True Then '对于长度>10的文本项目纵向输出文字信息
                    Call OutBigVConnect(strTmp, lngCurX, lngCurY + (1 * msngTwips), Val(arrItemHeight(intFields)), T_DrawClient.列单位, InStr(1, ",签名人,护士,", "," & strName & ",") <> 0, True, mTabForeColor)
                ElseIf isBigConnect(strName, 0) Then '是否是数值类型的项目
                    Call OutNumConnect(strTmp, lngCurX, lngCurY + (1 * msngTwips), Val(arrItemHeight(intFields)), T_DrawClient.列单位, IIf(strName = "日期", True, False))
                Else
                    Call OutBigHConnect(strTmp, lngCurX, lngCurY + (1 * msngTwips), Val(arrItemHeight(intFields)), T_DrawClient.列单位, True)
                End If
                lngHeight = lngHeight + Val(arrItemHeight(intFields))
            Next intFields
        .MoveNext
        Loop
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub OutNumConnect(ByVal strText As String, ByVal lngX As Long, ByVal lngY As Long, ByVal lngHeight As Long, ByVal lngColWidth As Long, Optional ByVal bln日期 As Boolean = False)
'功能：数字类型项目输出
    Dim i As Integer, j As Integer, sngD As Single, intSize As Integer, intOldSize As Integer
    Dim lngWidth As Long, lngTmp As Long, lngMaxY As Long, lngMaxX As Long
    Dim lngX1 As Long, lngY1 As Long, lngY2 As Long
    Dim strLeft As String, strRight As String
    Dim bln居中 As Boolean
    
    bln居中 = True
    intSize = mobjSubFont.Size
    intOldSize = intSize
    lngX1 = lngX: lngY1 = lngY
    lngTmp = lngColWidth
    If InStr(1, strText, "/") <> 0 Then
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngWidth = mobjDraw.TextWidth(strLeft) / sinTwipsPerPixelX
        If lngWidth > lngTmp Then
            sngD = Round((lngWidth - lngTmp) / lngWidth, 4)
            intSize = Round(Round((1 - sngD), 4) * intSize - 1)
            If intSize < 7 Then intSize = 7
            If intSize > intOldSize Then intSize = intOldSize
        End If
        mobjSubFont.Size = intSize
        Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
        Call GetTextRect(mobjDraw, lngX, lngY, strLeft, T_DrawClient.列单位, False)
        Call DrawText(mlngDC, strLeft, -1, T_LableRect, DT_CENTER)
        T_Size.H = mobjDraw.TextHeight("字") / sinTwipsPerPixelY
        '画对角线
        lngY2 = lngY1 + T_Size.H + (T_Size.H / 2)
        If lngY2 > lngY1 + lngHeight Then
            lngY2 = lngY1 + lngHeight
        End If
        Call DrawLine(mlngDC, lngX1, lngY2, lngX1 + T_DrawClient.列单位, lngY1 + (T_Size.H / 2), PS_SOLID, IIf(ISobjPrinter = True, 2, 1), mTabGridColor)
        lngHeight = lngHeight + lngY1 - lngY2
        lngY = lngY2
        lngY1 = lngY
        strText = Trim(strRight)
        '还原字体
        mobjSubFont.Size = intOldSize
        Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
        bln居中 = False
        GoTo ErrNext
    Else
ErrNext:
        '缩小字体进行输出
        lngWidth = mobjDraw.TextWidth(strText) / sinTwipsPerPixelX
        If lngWidth > lngTmp Then
            sngD = Round((lngWidth - lngTmp) / lngWidth, 4)
            intSize = Round(Round((1 - sngD), 4) * 9 - 1)
            If intSize < 7 Then Call OutBigHConnect(strText, lngX1, lngY1, lngHeight, lngColWidth, bln居中): GoTo ErrEnd
            If intSize > intOldSize Then intSize = intOldSize
        End If
        mobjSubFont.Size = intSize
        Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
        T_Size.H = mobjDraw.TextHeight("字") / sinTwipsPerPixelY
        If bln居中 = True Then
            lngY = lngY1 + (lngHeight - T_Size.H) / 2
            If lngY < lngY1 Then lngY = lngY1
        Else
            lngY = lngY1
        End If
        Call GetTextRect(mobjDraw, lngX, lngY, strText, T_DrawClient.列单位, False)
        Call DrawText(mlngDC, strText, -1, T_LableRect, DT_CENTER)
    End If
    
ErrEnd:
     '还原字体
    mobjSubFont.Size = intOldSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
End Sub

Private Function OutBigVConnect(ByVal strText As String, ByVal lngX As Long, ByVal lngY As Long, ByVal lngHeight As Long, ByVal lngColWidth As Long, _
    Optional ByVal bln居中Y As Boolean = False, Optional ByVal bln居中X As Boolean = True, Optional ByVal lngColor As Long = 0) As Single
'功能：纵向输出文本信息
    Dim i As Integer, j As Integer, sngD As Single, intSize As Integer, intOldSize As Integer
    Dim lngWidth As Long, lngTmp As Long, lngMaxY As Long, lngMaxX As Long
    Dim lngX1 As Long, lngY1 As Long
    Dim strTmp As String, strConnect As String
    Dim arrTmp
    
    '获取合适字体
    intSize = mobjSubFont.Size
    intOldSize = intSize
    lngWidth = mobjDraw.TextWidth(strText) / sinTwipsPerPixelX
    lngTmp = lngHeight * Int(lngColWidth / (mobjDraw.TextWidth("字") / sinTwipsPerPixelX))
    If lngTmp <= 0 Then lngTmp = lngHeight
    If lngWidth > lngTmp Then
        sngD = Round((lngWidth - lngTmp) / lngWidth, 4)
        intSize = Round(Round((1 - sngD), 4) * intSize - 1)
        If intSize < 7 Then intSize = 7
        If intSize > intOldSize Then intSize = intOldSize
    End If
    mobjSubFont.Size = intSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    '开始输出文本信息
    lngMaxY = lngY + lngHeight
    lngMaxX = lngX + lngColWidth
    lngX1 = lngX
    lngY1 = lngY
    strConnect = ""
    arrTmp = Array()
    
    lngX = lngX + (mobjDraw.TextWidth("字") / sinTwipsPerPixelX)
    ReDim arrTmp(UBound(arrTmp) + 1)
    For i = 1 To Len(strText)
        strTmp = Mid(strText, i, 1)
        If Asc(strTmp) > 0 Then
            T_Size.W = mobjDraw.TextWidth("字") / sinTwipsPerPixelX
            T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY / 2
        Else
            T_Size.W = mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX
            T_Size.H = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
        End If
        If lngY + T_Size.H > lngMaxY And strConnect <> "" Then
            If (lngX + T_Size.W - lngMaxX) > (T_Size.W * 0.2) Then GoTo ErrEnd
            lngX = lngX + T_Size.W
            lngY = lngY1
            strConnect = ""
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
           GoTo ErrNext
        Else
ErrNext:
            strConnect = strConnect & strTmp
            arrTmp(UBound(arrTmp)) = strConnect
            lngY = lngY + T_Size.H + (1 * msngTwips)
        End If
    Next i
ErrEnd:
    T_Size.W = mobjDraw.TextWidth("字") / sinTwipsPerPixelX
    If lngColWidth / (UBound(arrTmp) + 1) > T_Size.W Then
        lngWidth = Int(lngColWidth / (UBound(arrTmp) + 1))
    Else
        lngWidth = T_Size.W
    End If
    For i = 0 To UBound(arrTmp)
        lngX = lngX1 + (i * T_Size.W)
        If bln居中Y = True Then
            lngY = 0
            For j = 1 To Len(arrTmp(i))
                strTmp = Mid(arrTmp(i), j, 1)
                If Asc(strTmp) > 0 Then
                    T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY / 2
                Else
                    T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY
                End If
                lngY = lngY + T_Size.H + (1 * msngTwips)
            Next j
            lngY = lngY1 + (lngHeight - lngY) / 2
        Else
            lngY = lngY1
        End If
        
        For j = 1 To Len(arrTmp(i))
            strTmp = Mid(arrTmp(i), j, 1)
            If Asc(strTmp) > 0 Then
                T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY / 2
            Else
                T_Size.H = mobjDraw.TextHeight("1") / sinTwipsPerPixelY
            End If
            Call DrawRotateText(mobjDraw, mlngDC, lngX, lngY, strTmp, IIf(bln居中X = True, lngWidth, 0), lngColor)
            lngY = lngY + T_Size.H + (1 * msngTwips)
        Next j
    Next i
    '还原字体
    mobjSubFont.Size = intOldSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    
    OutBigVConnect = lngX + T_Size.W
End Function

Private Function OutBigHConnect(ByVal strText As String, ByVal lngX As Long, ByVal lngY As Long, ByVal lngHeight As Long, ByVal lngColWidth As Long, _
    Optional ByVal bln居中Y As Boolean = True, Optional ByVal bln居中X As Boolean = True) As Single
'功能：横向输出文本信息
    Dim i As Integer, j As Integer, sngD As Single, intSize As Integer, intOldSize As Integer
    Dim lngWidth As Long, lngTmp As Long, lngMaxY As Long, lngMaxX As Long
    Dim lngX1 As Long, lngY1 As Long
    Dim strTmp As String, strConnect As String
    Dim arrTmp
    
    '获取合适字体
    intSize = mobjSubFont.Size
    intOldSize = intSize
    lngWidth = mobjDraw.TextWidth(strText) / sinTwipsPerPixelX
    lngTmp = lngHeight * Int(lngColWidth / (mobjDraw.TextWidth("字") / sinTwipsPerPixelX))
    If lngTmp <= 0 Then lngTmp = lngHeight
    If lngWidth > lngTmp Then
        sngD = Round((lngWidth - lngTmp) / lngWidth, 4)
        intSize = Round(Round((1 - sngD), 4) * intSize - 1)
        If intSize < 7 Then intSize = 7
        If intSize > intOldSize Then intSize = intOldSize
    End If
    mobjSubFont.Size = intSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    '开始输出文本信息
    lngMaxY = lngY + lngHeight
    lngMaxX = lngX + lngColWidth
    lngX1 = lngX
    lngY1 = lngY
    strConnect = ""
    arrTmp = Array()
    
    lngY = lngY + (mobjDraw.TextHeight("字") / sinTwipsPerPixelY)
    ReDim arrTmp(UBound(arrTmp) + 1)
    For i = 1 To Len(strText)
        strTmp = Mid(strText, i, 1)
        If Asc(strTmp) > 0 Then
            T_Size.W = mobjDraw.TextWidth("字") / sinTwipsPerPixelX / 2
        Else
            T_Size.W = mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX
        End If
        T_Size.H = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
        If lngX + T_Size.W > lngMaxX And strConnect <> "" Then
            If (lngY + T_Size.H - lngMaxY) > (T_Size.W * 0.2) Then GoTo ErrEnd
            lngY = lngY + T_Size.H + (1 * msngTwips)
            lngX = lngX1
            strConnect = ""
            ReDim Preserve arrTmp(UBound(arrTmp) + 1)
           GoTo ErrNext
        Else
ErrNext:
            strConnect = strConnect & strTmp
            arrTmp(UBound(arrTmp)) = strConnect
            lngX = lngX + T_Size.W
        End If
    Next i
ErrEnd:
    If bln居中Y = True Then
        T_Size.H = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
        lngY = lngY1 + (lngHeight - T_Size.H * (UBound(arrTmp) + 1)) / 2
        If lngY < lngY1 Then lngY = lngY1
        lngY1 = lngY
    End If
    lngWidth = lngColWidth
    For i = 0 To UBound(arrTmp)
        lngX = lngX1
        lngY = lngY1 + (i * (T_Size.H + (1 * msngTwips)))
        Call GetTextRect(mobjDraw, lngX, lngY, CStr(arrTmp(i)), IIf(bln居中X = True, lngColWidth, 0), False)
        Call DrawText(mlngDC, CStr(arrTmp(i)), -1, T_LableRect, DT_CENTER)
    Next i
    '还原字体
    mobjSubFont.Size = intOldSize
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    
    OutBigHConnect = lngY + T_Size.H
End Function

Private Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional ByVal blnAddPage As Boolean = True) As ADODB.Recordset
    '只拷贝记录集的结构,同时增加页号,行号字段
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "汇总日期" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:表示新增
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '日期型处理为字符型
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
            End If
        Next
        If blnAddPage Then
            .Fields.Append "列号", adDouble, 18
        End If
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set CopyNewRec = rsTarget
End Function

Public Function GetXCoordinate(ByVal strInput As String, ByVal strBeginDate As String, Optional ByVal bln坐标 As Boolean = True) As String

    '根据时间得到X坐标或根据X坐标转换为时间范围
    Dim sinX   As Single
    Dim sinTime As Single
    Dim strDay As String

    On Error GoTo errHand
    
    If bln坐标 Then
        '计算查多少分钟
        sinTime = Format(DateDiff("n", CDate(strBeginDate), CDate(strInput)) / 60, "#0.0000;-#0.0000;0000")
        
        '计算得到X坐标(每天6列,以列数*列单位得到坐标)
        sinX = Format(T_DrawClient.产程区域.Left + (sinTime * T_DrawClient.列单位), "#0.0")
        GetXCoordinate = sinX
    Else
        '计算得到相差多少个刻度
        sinX = Val(strInput)
        sinTime = Format(sinX - T_DrawClient.产程区域.Left, "#0.0000;-#0.0000;0000") / T_DrawClient.列单位
        sinTime = Format(sinTime * 60, "#0.0")
        strDay = Format(DateAdd("n", sinTime, strBeginDate), "yyyy-MM-dd HH:mm:ss")
        GetXCoordinate = strDay
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function GetYCoordinate(ByVal int项目序号 As Integer, ByVal strInput As String, Optional ByVal bln坐标 As Boolean = True) As Single

    Dim sinCurX As Single, sinCurY As Single, sinScale As Single
    On Error GoTo errHand
    '返回指定曲线数据的Y坐标或根据Y坐标计算数据
    '测试该函数的正确性可以在计算坐标加该代码实现(思想:由该函数自己根据数据计算得到Y坐标,再转换为数据,再转换为坐标后输出字符进行核对,打印无误则说明转换无误):
    
    mrsDrawItems.Filter = "项目序号=" & int项目序号
    If mrsDrawItems.RecordCount = 0 Then
        GetYCoordinate = 0
        Exit Function
    End If
    
    If bln坐标 Then
        '得到有效数据起始坐标
        sinCurX = Split(mrsDrawItems!最大值坐标, ",")(0)
        sinCurY = Split(mrsDrawItems!最大值坐标, ",")(1)
        
        '根据最大值与当前值之间的差额,以及最小值,计算得到相差多少个刻度,再根据单位刻度得到实际坐标
        If Val(mrsDrawItems!显示模式) = 1 Then '是从最小值到最大值
            sinScale = (-1 * (mrsDrawItems!最小值 - Val(strInput)) / mrsDrawItems!单位值) * Val(Split(mrsDrawItems!单位刻度, ",")(0))
        Else
            sinScale = ((mrsDrawItems!最大值 - Val(strInput)) / mrsDrawItems!单位值) * Val(Split(mrsDrawItems!单位刻度, ",")(0))
        End If
        GetYCoordinate = Format(sinCurY + sinScale, "#0.0;-#0.0;0")
    Else
        '得到传入的坐标值
        sinCurY = CDbl(strInput)
        
        '(坐标-最大值坐标)/单位刻度得到相差多少个刻度
        '(最大值-单位刻度*单位值)得到实际数据
        sinScale = (sinCurY - Split(mrsDrawItems!最大值坐标, ",")(1)) / Val(Split(mrsDrawItems!单位刻度, ",")(0))
        If Val(mrsDrawItems!显示模式) = 1 Then  '是从最小值到最大值
            GetYCoordinate = Format(mrsDrawItems!最小值 + sinScale * mrsDrawItems!单位值, "#0.0;-#0.0;0")
        Else
            GetYCoordinate = Format(mrsDrawItems!最大值 - sinScale * mrsDrawItems!单位值, "#0.0;-#0.0;0")
        End If
    End If

    mrsDrawItems.Filter = ""
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetDrawItemValue(ByVal lngItemNo As Long, ByVal strFields As String) As String
'-----------------------------------------------------------------------
'功能:根据字段获取相应信息
'-----------------------------------------------------------------------
    Dim strValue As String
    
    If InStr(1, "|项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|记录符|颜色|", "|" & strFields & "|") = 0 Then Exit Function
    
    mrsDrawItems.Filter = ""
    mrsDrawItems.Filter = "项目序号=" & lngItemNo
    If mrsDrawItems.RecordCount > 0 Then
        strValue = mrsDrawItems.Fields(strFields).Value
    End If
    GetDrawItemValue = strValue
End Function

Private Function ISobjPrinter() As Boolean
'判断释放是打印机对象
    Dim blnPrinter As Boolean
    blnPrinter = (TypeName(mobjDraw) = "Printer")
    ISobjPrinter = blnPrinter
End Function

Private Sub InitEnv()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    RGB_BLACK = RGB(0, 0, 0)
    RGB_RED = RGB(255, 0, 0)
    RGB_WRITE = RGB(255, 255, 255)
    RGB_BLUE = RGB(0, 0, 255)
    RGB_GRAY = &H808080
    RGB_FleetGRAY = &HC0C0C0
    
    '打开现存在的所有护理记录项目
    gstrSQL = " Select   项目序号,项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式,nvl(保留项目,0) 保留项目" & _
              " From 护理记录项目 B" & _
              " Order by 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    '提取所有产程要素信息
    gstrSQL = "Select 中文名,替换域,类型,长度,小数,单位,表示法,数值域,必填" & vbNewLine & _
        "From (Select i.分类id, i.编码, i.中文名, nvl(i.替换域,0) 替换域,i.类型,i.长度,i.小数,i.单位,i.表示法,i.数值域,i.必填" & vbNewLine & _
        "       From 诊治所见项目 I, 诊治所见分类 K" & vbNewLine & _
        "       Where k.Id = i.分类id And k.编码 In ('02', '03', '05', '06') And i.替换域 = 1 And k.性质 = 1" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select i.分类id, i.编码, i.中文名, nvl(i.替换域,0) 替换域,i.类型,i.长度,i.小数,i.单位,i.表示法,i.数值域,i.必填" & vbNewLine & _
        "       From 诊治所见项目 I, 诊治所见分类 K" & vbNewLine & _
        "       Where k.Id = i.分类id And k.编码 In ('04', '05') And k.性质 = 2)" & vbNewLine & _
        "Order By 分类id, 编码, 替换域"

    Set mrsPartogram = zlDatabase.OpenSQLRecord(gstrSQL, "提取产程要素信息")
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function PrintState(ByVal lngFormatID As Long, Optional ByVal strPrintDevice As String = "") As Boolean
'******************************************************************************************************************
'功能:设置打印机属性
'******************************************************************************************************************
    Dim i As Long
    Dim strPaper As String
    Dim strPrintName As String
    Dim blnYesPrinter As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    
    '------------------------------------------------------------------------------------------------------------------
    '打印机恢复及设置
    If Not ExistsPrinter Then
        MsgBox "系统没有安装任何打印机不能继续打印，程序退出！", vbInformation, gstrSysName
        Exit Function
    End If
    gstrSQL = "Select 格式,页脚 From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历页面格式", lngFormatID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!格式: mstrPageFoot = "" & rsTemp!页脚
    End If
    
    gPrinter.lngLeft = OFFSET_LEFT
    gPrinter.lngRight = OFFSET_RIGHT
    gPrinter.lngTop = OFFSET_TOP
    gPrinter.lngBottom = OFFSET_BOTTOM
       
    If UBound(Split(mstrPaperSet, ";")) >= 0 Then
        gPrinter.intPage = Val(Split(mstrPaperSet, ";")(0))
        If UBound(Split(mstrPaperSet, ";")) >= 1 Then gPrinter.intOrient = Val(Split(mstrPaperSet, ";")(1))
        If UBound(Split(mstrPaperSet, ";")) >= 2 Then gPrinter.lngHeight = Val(Split(mstrPaperSet, ";")(2))
        If UBound(Split(mstrPaperSet, ";")) >= 3 Then gPrinter.lngWidth = Val(Split(mstrPaperSet, ";")(3))
        If UBound(Split(mstrPaperSet, ";")) >= 4 Then gPrinter.lngLeft = CLng(Val(Split(mstrPaperSet, ";")(4)) / conRatemmToTwip)
        If UBound(Split(mstrPaperSet, ";")) >= 5 Then gPrinter.lngRight = CLng(Val(Split(mstrPaperSet, ";")(5)) / conRatemmToTwip)
        If UBound(Split(mstrPaperSet, ";")) >= 6 Then gPrinter.lngTop = CLng(Val(Split(mstrPaperSet, ";")(6)) / conRatemmToTwip)
        If UBound(Split(mstrPaperSet, ";")) >= 7 Then gPrinter.lngBottom = CLng(Val(Split(mstrPaperSet, ";")(7)) / conRatemmToTwip)
    End If
    
    If strPrintDevice = "" Then
        If Trim(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", Printers(0).DeviceName)) = "" Then
            MsgBox "没有设置打印机,将使用系统默认打印机设置！", vbInformation, gstrSysName
        Else
            strPrintName = Trim(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "DeviceName", Printers(0).DeviceName))
        End If
    Else
        strPrintName = strPrintDevice
    End If
    
    '打印机
    blnYesPrinter = False
    If Printer.DeviceName <> strPrintName Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrintName Then Set Printer = Printers(i): blnYesPrinter = True: Exit For
        Next
        If blnYesPrinter = False Then
            MsgBox "设置的打印机已不存在,将使用系统默认打印机设置！", vbInformation, gstrSysName
        End If
    End If
            
    gPrinter.intBin = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\Default", "PaperBin", ""))
    
    On Error Resume Next
    '纸张
    If gPrinter.intPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = gPrinter.lngWidth
        Printer.Height = gPrinter.lngHeight
    Else
        Printer.PaperSize = gPrinter.intPage
    End If
    Printer.Orientation = gPrinter.intOrient
    If IsWindowsNT And gPrinter.intPage = 256 Then
        Call SetNTPrinterPaper(frmSample.Hwnd, gPrinter.lngWidth / conRatemmToTwip, gPrinter.lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
        Unload frmSample
    End If
    
    'WinNT自定义纸张处理
    If IsWindowsNT And gPrinter.intPage = 256 Then DelCustomPaper
    
    PrintState = True
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    '读取文件属性
    mblnDateAd = False
    gstrSQL = " Select   格式ID From 病人护理文件 " & _
              " Where 病人ID=[1] And 主页ID=[2] And 婴儿=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件数据", T_Info.lng病人ID, T_Info.lng主页ID, 0, T_Info.lng文件ID)
    If rsTemp.RecordCount > 0 Then mlng格式ID = rsTemp!格式ID

    '提取活动项目并加入列定义(格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...)
    mbln日期时间合并 = False
    
    '读取病历文件格式定义
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历文件格式定义", mlng格式ID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数": mintTabTiers = Val("" & !内容文本)
            Case "总列数":   mlngItems = Val("" & !内容文本)
            Case "最小行高": '
            Case "文本字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = 9 ' Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjSubFont = objFont
                
            Case "文本颜色": mTabForeColor = Val("" & !内容文本)
            Case "表格颜色": mTabGridColor = Val("" & !内容文本)
            
            Case "标题文本": mstrTitle = "" & !内容文本
            Case "标题字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTitleFont = objFont
            
            Case "开始时间": mintTagFormHour = Val("" & !内容文本)
            Case "终止时间": mintTagToHour = Val("" & !内容文本)
            Case "条件字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "条件颜色": mlngTagColor = Val("" & !内容文本)
            Case "有效数据行"
                '
            Case "日期时间合并"
                mbln日期时间合并 = (Val(!内容文本) = 1)
            End Select
            .MoveNext
        Loop
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表上标签定义", mlng格式ID)
    With rsTemp
        mstrSubHead = ""
        Do While Not .EOF
            mstrSubHead = mstrSubHead & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubHead <> "" Then mstrSubHead = Replace(Mid(mstrSubHead, 2), Chr(1), " ")
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表下标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表上标签定义", mlng格式ID)
    With rsTemp
        mstrSubEnd = ""
        Do While Not .EOF
            mstrSubEnd = mstrSubEnd & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubEnd <> "" Then mstrSubEnd = Replace(Mid(mstrSubEnd, 2), Chr(1), " ")
    End With
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元' And d.内容行次=1" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表头单元定义", mlng格式ID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !内容行次 - 1 & "," & !对象序号 & "," & !内容文本
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '查询语句组织
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql外 As String, str格式 As String, strSqlNull As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim bln对角线 As Boolean, bln选择项 As Boolean          '如果上一列是对角线且选择项,则直接提取各项数据,拼列头时在数值间加上/
    Dim lngColumn As Long, blnAddCollect As Boolean
    
    gstrSQL = "Select   d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表列集合定义", mlng格式ID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = "": strSqlNull = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            If lngColumn <> !对象序号 Then
                blnAddCollect = False
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) & "|" & !对象序号 & "'" & !要素名称
                mstrColWidth = mstrColWidth & "," & !对象属性 & "`" & !对象序号 & "`" & !要素表示
                If !要素表示 = 1 Then mstrCatercorner = mstrCatercorner & "," & !对象序号
                str格式 = ""
                If !要素名称 <> "" Then
                    str格式 = "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                    If Mid(strSqlNull, 3) = "" Then
                        strSqlNull = "''"
                    Else
                        strSqlNull = Mid(strSqlNull, 3)
                    End If
                    mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
                    
                Else
                    If strSql外 <> "" Then
                        mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql外 = ""
                strSqlNull = ""
                lngColumn = !对象序号
                bln对角线 = (NVL(!要素表示, 0) = 1)
                bln选择项 = False
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    bln选择项 = (mrsItems!项目表示 = 5)
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !要素名称
                str格式 = str格式 & "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
            End If
            
            Select Case !要素名称
            Case "日期"
                bln日期 = True
                mblnDateAd = (NVL(!要素表示, 0) = 1)
                mstrSQL中 = mstrSQL中 & ",日期"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As 日期"
                strSql外 = strSql外 & "||" & !要素名称
            Case "时间"
                bln时间 = True
                mstrSQL中 = mstrSQL中 & ",时间"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名人"
                bln签名人 = True
                mstrSQL中 = mstrSQL中 & ",签名人"
                mstrSQL内 = mstrSQL内 & ",l.签名人"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名时间"
                bln签名时间 = True
                mstrSQL中 = mstrSQL中 & ",签名时间"
                mstrSQL内 = mstrSQL内 & ",l.签名时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "护士"
                bln护士 = True
                mstrSQL中 = mstrSQL中 & ",护士"
                mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
                strSql外 = strSql外 & "||" & !要素名称
            Case Else
                If !要素名称 <> "" Then
                    mstrSQL中 = mstrSQL中 & ",Max(""" & !要素名称 & """) As """ & !要素名称 & """"
                    'mstrSQL条件 = mstrSQL条件 & " Or """ & !要素名称 & """ Is Not Null"
                    
                    If bln对角线 And bln选择项 Then
                        If strSql外 <> "" Then
                            '第二项
                            strSql外 = strSql外 & "||'/'||""" & !要素名称 & """"
                        Else
                            '第一项
                            strSql外 = strSql外 & "||""" & !要素名称 & """"
                        End If
                    Else
                        strSql外 = strSql外 & "||""" & !要素名称 & """"
                        strSqlNull = strSqlNull & "||" & "'" & !内容文本 & "'||'" & !要素单位 & "'"
                    End If
                    
                    If (Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "") Or (bln对角线 And bln选择项) Then
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & !要素名称 & """"
                        mstrSQL条件 = mstrSQL条件 & " Or Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') Is Not Null"
                    Else
                        'mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')), '') As """ & !要素名称 & """"
                        mstrSQL条件 = mstrSQL条件 & " Or Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')), '') Is Not Null"
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')),  '" & !内容文本 & "'||'" & !要素单位 & "') As """ & !要素名称 & """"
                    End If
                End If
            End Select
            .MoveNext
        Loop
        
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '加入最后一列的格式
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) '& "|" & !对象序号 & "'" & !要素名称
        mstrColumns = Mid(mstrColumns, 2)     '格式如:列号;项目名称1,项目名称2|列号...,实例;1;体温|2;脉搏|3...
        If Mid(strSql外, 3) <> "" Then
            mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL条件 <> "" Then mstrSQL条件 = "(" & Mid(mstrSQL条件, 5) & ")"
        
        '如果没有出现日期，时间，护士，则内层需要补充，以保证中层分组的正常：
        If bln日期 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
        If bln时间 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
        If bln护士 = False Then mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
        If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
        If bln签名时间 = False Then mstrSQL内 = mstrSQL内 & ",l.签名时间"
    End With
    '初始化表格信息
    Call InitTabRecords
    ReadStruDef = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub GetFileProperty()
    '提取文件属性
    Dim rsTemp As New ADODB.Recordset
    Dim strEnd As String
    On Error GoTo errHand
    
    gstrSQL = " Select   开始时间,结束时间,格式ID,科室ID,归档人 From 病人护理文件 " & _
              " Where 病人ID=[1] And 主页ID=[2] And 婴儿=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件数据", T_Info.lng病人ID, T_Info.lng主页ID, 0, T_Info.lng文件ID)
    If rsTemp.RecordCount <> 0 Then
        mlng格式ID = rsTemp!格式ID
        mstr开始时间 = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm")
        mstr结束时间 = Format(rsTemp!结束时间, "yyyy-MM-dd HH:mm")
        mstr宫缩时间 = Format(mstr开始时间, "yyyy-MM-dd HH:mm")
        strEnd = DateAdd("n", -1, CDate(Format(CDate(mstr开始时间) + 1, "yyyy-MM-dd HH:mm:ss")))
        If mstr结束时间 = "" Then
            mstr结束时间 = strEnd
        Else
            If (mstr结束时间 <> "" And CDate(mstr结束时间) > CDate(strEnd)) Then mstr结束时间 = strEnd
        End If
    End If
    '第二份文件重新提取开始时间
    If T_Info.lng份数 > 1 Then
        gstrSQL = "SELECT Max(B.发生时间) 发生时间" & vbNewLine & _
            "FROM 病人护理文件 A,病人护理数据 B,病人护理明细 C,护理记录项目 D" & vbNewLine & _
            "WHERE A.ID=B.文件ID AND B.ID=C.记录ID AND A.ID=[1] And 病人ID=[2] And 主页ID=[3] And 婴儿=[4] AND B.汇总类别<[5] AND C.项目序号=D.项目序号" & vbNewLine & _
            "AND NVL(D.项目名称,'')='生产' AND NVL(D.保留项目,1)=1 ORDER BY B.发生时间"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件数据", T_Info.lng文件ID, T_Info.lng病人ID, T_Info.lng主页ID, 0, T_Info.lng份数)
        If rsTemp.RecordCount <> 0 Then
            mstr开始时间 = DateAdd("n", 1, CDate(Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm")))
        End If
    End If
    
    '获取文件页数
    Call GetPartogramPage(T_Info.lng文件ID, T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng份数)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SQLCombination()
'-提取数据SQL
    mstrSQL = "Select to_char(发生时间,'yyyy-MM-dd hh24:mi:ss') AS 发生时间," & Mid(mstrSQL列, 12) & vbCrLf & _
                " From (Select 汇总类别,时间 as 备用,发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select l.汇总类别,l.发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f " & vbCrLf & _
                "               Where l.Id = c.记录id And l.文件ID=f.ID " & _
                "               And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] And l.汇总类别=[5]" & IIf(mstrSQL条件 <> "", " And (" & mstrSQL条件 & ")", "") & ")" & vbCrLf & _
                "       Group By 日期, 时间, 发生时间,汇总类别,护士,签名人,签名时间" & _
                                "       Order By 汇总类别,发生时间,护士,签名人,签名时间)"
End Sub

Private Sub SQLCombinationPage()
'-获取数据列好SQL
    mstrSQL = " SELECT 发生时间,FLOOR(TO_NUMBER(发生时间-开始时间)*24)+1 AS 行" & vbCrLf & _
                " From (Select 汇总类别,时间 as 备用,发生时间,Max(开始时间) 开始时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select l.汇总类别,l.发生时间,F.开始时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f " & vbCrLf & _
                "               Where l.Id = c.记录id And l.文件ID=f.ID " & _
                "               And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] And l.汇总类别=[5]" & IIf(mstrSQL条件 <> "", " And (" & mstrSQL条件 & ")", "") & ")" & vbCrLf & _
                "       Group By 日期, 时间, 发生时间,汇总类别,护士,签名人,签名时间" & _
                                "       Order By 汇总类别,发生时间,护士,签名人,签名时间)"
End Sub

Private Function GetPeriod() As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    gstrSQL = " Select   入院日期 AS 开始时间 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取入院日期或出生日期", T_Info.lng病人ID, T_Info.lng主页ID)
    GetPeriod = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(mstr结束时间, "yyyy-MM-dd HH:mm:ss")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsDiagonal(ByVal lngOrder As Long) As Boolean
    '判断指定列是否设置了列对角线
    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & lngOrder & ",") <> 0)
End Function

Private Function isBigConnect(ByVal strName As String, ByVal bytMode As Byte) As String
'------------------------------------------------------
'功能：判断项目类型
'bytMode:0-数值;1-大文本(长度>10);2-其他(如：单选;多选;文本长度<10的文本项目)
'------------------------------------------------------
    Dim blnOK As Boolean
    
    mrsItems.Filter = ""
    mrsItems.Filter = "项目名称='" & strName & "'"
    If mrsItems.RecordCount > 0 Then
        Select Case bytMode
        Case 0
            If Val(NVL(mrsItems!项目类型, 0)) = 0 Then
                blnOK = True
            End If
        Case 1
            If Val(NVL(mrsItems!项目类型, 0)) = 1 And Val(NVL(mrsItems!项目表示, 0)) = 0 Then
                blnOK = Val(NVL(mrsItems!项目长度, 0)) > 10
            End If
        Case Else
            blnOK = True
        End Select
    Else
        Select Case bytMode
        Case 0
            If strName = "时间" Or strName = "日期" Then blnOK = True
        Case 1
            If strName = "签名人" Or strName = "护士" Then blnOK = True
        Case Else
            blnOK = True
        End Select
    End If
    isBigConnect = blnOK = True
End Function


Private Sub GetMarkConnect()
'----------------------------------------------------------------------
'功能:提取上下标信息
'----------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim aryRow() As String, aryItem() As String, arrItemEnd() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strSubHend As String, strSubEnd As String
    Dim strTmp As String, str单位 As String
    Dim i As Integer
    
    On Error GoTo errHand
    
    strSubHend = ""
    strSubEnd = ""
    aryPeriod = Split(GetPeriod, "～")
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as 信息 From Dual"
    aryItem = Split(mstrSubHead, "|")
    arrItemEnd = Split(mstrSubEnd, "|")
    For i = 0 To 1
        For lngCount = 0 To IIf(i = 0, UBound(aryItem), UBound(arrItemEnd))
            If i = 0 Then
                strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
                strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
            Else
                strPrefix = Left(arrItemEnd(lngCount), InStr(1, arrItemEnd(lngCount), "{") - 1)
                strItemName = Mid(arrItemEnd(lngCount), InStr(1, arrItemEnd(lngCount), "{") + 1, InStr(1, arrItemEnd(lngCount), "}") - InStr(1, arrItemEnd(lngCount), "{") - 1)
            End If
            mrsPartogram.Filter = 0
            mrsPartogram.Filter = "中文名='" & strItemName & "'"
            '不可能找不到，除非手工修改数据
            If mrsPartogram.RecordCount = 0 Then GoTo ErrNext
            str单位 = Trim(NVL(mrsPartogram!单位))
            If Val(NVL(mrsPartogram!替换域)) = 1 Then
                '产程固定要素信息
                strTmp = strPrefix
                Select Case strItemName
                Case "当前病区"
                
                    strTmpSQL = "Select   b.名称" & vbNewLine & _
                                "From (Select 病区id, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,部门表 b " & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.病区id Is Not Null And b.ID=a.病区id" & vbNewLine & _
                                "Order By a.开始时间"
                                
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前病区", T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    
                Case "当前床号"
                
                    strTmpSQL = "Select   a.床号" & vbNewLine & _
                                "From (Select 床号, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.床号 Is Not Null" & vbNewLine & _
                                "Order By a.开始时间"
        
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前床号", T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "当前科室"
                
                    strTmpSQL = "Select   名称 From 部门表 a Where a.ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前科室", T_Info.lng科室ID)
                    
                Case "住院医师"
                    strTmpSQL = "Select   a.经治医师" & vbNewLine & _
                                "From (Select 经治医师, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.经治医师 Is Not Null" & vbNewLine & _
                                "Order By a.开始时间"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "住院医师", T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "责任护士"
                
                    strTmpSQL = "Select   a.责任护士" & vbNewLine & _
                                "From (Select 责任护士, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.责任护士 Is Not Null" & vbNewLine & _
                                "Order By a.开始时间"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "责任护士", T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "护理等级"
                    strTmpSQL = "Select   b.名称" & vbNewLine & _
                                "From (Select 护理等级ID, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,护理等级 b" & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.护理等级ID Is Not Null And b.序号=a.护理等级ID" & vbNewLine & _
                                "Order By a.开始时间"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "护理等级", T_Info.lng病人ID, T_Info.lng主页ID, T_Info.lng科室ID, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "最后诊断"
                    strTmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as 信息 From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, T_Info.lng病人ID, T_Info.lng主页ID, 0, CDate(aryPeriod(0)))
                Case Else
                    strTmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as 信息 From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, T_Info.lng病人ID, T_Info.lng主页ID, 0)
                End Select
            Else
                '产程录入要素信息
                strTmp = strPrefix
                gstrSQL = "SELECT 内容 From 产程要素内容" & _
                    "   Where 文件ID = [1] And 婴儿 = [2] And 名称 =[3]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", T_Info.lng文件ID, T_Info.lng份数, strItemName)
            End If
            If rsTemp.BOF = False Then
                If i = 0 Then
                    If strTmp <> "" Then
                        strSubHend = strSubHend & "[ZLSOFTLPF]" & strTmp & rsTemp.Fields(0).Value & str单位
                    Else
                        strSubHend = strSubHend & "[ZLSOFTLPF]" & rsTemp.Fields(0).Value & str单位
                    End If
                Else
                    If strTmp <> "" Then
                        strSubEnd = strSubEnd & "[ZLSOFTLPF]" & strTmp & rsTemp.Fields(0).Value & str单位
                    Else
                        strSubEnd = strSubEnd & "[ZLSOFTLPF]" & rsTemp.Fields(0).Value & str单位
                    End If
                End If
            Else
                If i = 0 Then
                    If strTmp <> "" Then
                        strSubHend = strSubHend & "[ZLSOFTLPF]" & strTmp
                    Else
                        strSubHend = strSubHend & "[ZLSOFTLPF]"
                    End If
                Else
                    If strTmp <> "" Then
                        strSubEnd = strSubEnd & "[ZLSOFTLPF]" & strTmp
                    Else
                        strSubEnd = strSubEnd & "[ZLSOFTLPF]"
                    End If
                End If
            End If
ErrNext:
        Next
    Next i
    If Left(strSubHend, 11) = "[ZLSOFTLPF]" Then strSubHend = Mid(strSubHend, 12)
    If Left(strSubEnd, 11) = "[ZLSOFTLPF]" Then strSubEnd = Mid(strSubEnd, 12)
    mstrOutSubHead = strSubHend
    mstrOutSubEnd = strSubEnd
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetFileCount(ByVal lng文件ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long) As Long
'获取文件份数
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    On Error GoTo errHand
    mstrSQL = "SELECT NVL(MAX(NVL(汇总类别,1)),1) 文件" & vbNewLine & _
            "FROM 病人护理文件 A,病人护理数据 B" & vbNewLine & _
            "WHERE A.ID=[1] AND A.病人ID=[2] AND A.主页ID=[3] AND A.ID=B.文件ID"
    Call SQLDIY(mstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "提取文件份数", lng文件ID, lng病人ID, lng主页ID)
    If rsTemp.RecordCount = 0 Then
        lngCount = 1
    Else
        lngCount = Val(NVL(rsTemp!文件, 1))
    End If
    If lngCount < 1 Then lngCount = 1
    GetFileCount = lngCount
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitTabRecords()
    Dim i As Integer, j As Integer, k As Integer
    Dim lngCol As Long, strName As String, lngColHeight As Long, lngOrder As Long, lngItemNo As Long, strName1 As String
    Dim arrColumn, arrItem, arrWidth, strColumns As String, strTabHead As String, strColWidth As String
    On Error GoTo errHand
    
    strColumns = mstrColumns
    strTabHead = mstrTabHead
    strColWidth = mstrColWidth
    '初始化表格格式内容记录集
    gstrFields = "行," & adDouble & ",18|对象序号," & adDouble & ",18|名称," & adLongVarChar & ",20|高度," & adDouble & ",18|固定," & adInteger & ",1|要素名称," & adLongVarChar & ",20"
    Call Record_Init(mrsSelItems, gstrFields)
    gstrFields = "行|对象序号|名称|高度|固定|要素名称"
    
    arrColumn = Split(strColumns, "|")
    arrItem = Split(strTabHead, "|")
    arrWidth = Split(strColWidth, ",")
    j = UBound(arrColumn)
    lngCol = 1
    For i = 0 To j
        lngOrder = Split(arrColumn(i), "'")(0)
        k = UBound(Split(Split(arrColumn(i), "'")(1), ","))
        strName1 = Split(Split(arrColumn(i), "'")(1), ",")(0)
        lngColHeight = Format(CDbl((Split(arrWidth(i), "`")(0)) / sinTwipsPerPixelX), "0.00")
        '固定项目即为：宫口扩大、先露高低、生产、处理
        mrsItems.Filter = "项目名称='" & strName1 & "' And 保留项目=1"
        If mrsItems.RecordCount = 0 Then
ErrAdd:
            strName = Split(arrItem(i), ",")(2)
            gstrValues = lngCol & "|" & lngOrder & "|" & strName & "|" & lngColHeight & "|0|" & strName1
            Call Record_Add(mrsSelItems, gstrFields, gstrValues)
            lngCol = lngCol + 1
        Else
            Select Case strName1
                Case "宫口扩大"
                    lngItemNo = T_Partogram.lng宫口扩大
                Case "先露高低"
                    lngItemNo = T_Partogram.lng先露高低
                Case "生产"
                    lngItemNo = T_Partogram.lng生产
                Case "处理"
                    lngItemNo = T_Partogram.lng处理
                Case Else
                    GoTo ErrAdd
            End Select
            gstrValues = lngItemNo & "|" & lngOrder & "|" & strName1 & "|" & lngColHeight & "|1|" & strName1
            Call Record_Add(mrsSelItems, gstrFields, gstrValues)
            If strName1 = "处理" Or k > 0 Then GoTo ErrAdd
        End If
    Next
    mrsItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function DrawPartogramTab() As Single
'-------------------------------------------------------------------------
'功能:根据构造的产程录入，画产程表格区域
'-------------------------------------------------------------------------
    On Error GoTo errHand
    Dim lngCurX As Long, lngCurY As Long, lngCount As Long
    Dim lngCurMaxY As Long, lngCurMaxX As Long, lngHeight As Long
    Dim intBold As Integer, intFine As Integer
    Dim blnPrinter As Boolean, blnInit As Boolean
    Dim strConnect As String, strConnect1 As String, strConnect2 As String
    Dim i As Integer
    
    If TypeName(mobjDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    
    If blnPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '----首先画表头内容
    lngCurX = T_DrawClient.表格区域.Left
    lngCurY = T_DrawClient.表格区域.Top
    lngCurMaxX = T_DrawClient.MaxX
    lngCurMaxY = lngCurY
    Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurMaxX, lngCurY, PS_SOLID, intFine, mTabGridColor)
    Call SetTextColor(mlngDC, RGB_BLACK)
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    mrsSelItems.Filter = "固定=0"
    mrsSelItems.Sort = "行"
    If mrsSelItems.RecordCount = 0 Then GoTo ErrOver
    lngHeight = 0
    blnInit = False
    strConnect2 = ""
    lngCount = 0
    '画表头内容，表头内容可能存在合并的情况
    With mrsSelItems
        Do While Not .EOF
            strConnect1 = Trim(NVL(mrsSelItems!名称))
            If blnInit = False Then strConnect2 = strConnect1
            If strConnect1 <> strConnect2 Then
ErrEnd:
                strConnect = CheckConnect(strConnect2, T_DrawClient.表格区域.Right - T_DrawClient.表格区域.Left, lngHeight)
                '根据表格的宽度和高度检查字体
                T_Size.H = mobjDraw.TextHeight(strConnect) / sinTwipsPerPixelY
                T_Size.H = (lngHeight - T_Size.H) / 2
                Call GetTextRect(mobjDraw, lngCurX, lngCurY + T_Size.H, strConnect, T_DrawClient.刻度单位, False)
                Call DrawText(mlngDC, strConnect, -1, T_LableRect, DT_CENTER)
                Call DrawLine(mlngDC, lngCurX, lngCurY + lngHeight, T_DrawClient.表格区域.Right, lngCurY + lngHeight, PS_SOLID, intFine, mTabGridColor)
                lngCurY = lngCurY + lngHeight
                lngHeight = 0
                strConnect2 = strConnect1
            End If
            lngHeight = lngHeight + Val(mrsSelItems!高度)
            blnInit = True
            lngCount = lngCount + 1
            If mrsSelItems.RecordCount = lngCount Then GoTo ErrEnd
        .MoveNext
        Loop
    End With
    '划横线
    mrsSelItems.Filter = "固定=0"
    mrsSelItems.Sort = "行"
    lngCurY = T_DrawClient.表格区域.Top
    With mrsSelItems
        Do While Not .EOF
            lngCurY = lngCurY + Val(mrsSelItems!高度)
             Call DrawLine(mlngDC, T_DrawClient.表格区域.Right, lngCurY, lngCurMaxX, lngCurY, PS_SOLID, intFine, mTabGridColor)
        .MoveNext
        Loop
    End With
    lngCurMaxY = lngCurY
    '画左边的竖线
    lngCurX = T_DrawClient.表格区域.Left
    lngCurY = T_DrawClient.表格区域.Top
    Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngCurMaxY, PS_SOLID, intFine, mTabGridColor)
    
    lngCurY = T_DrawClient.表格区域.Top
    lngCurX = T_DrawClient.表格区域.Right
    '画竖线
    For i = 0 To 24
        Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngCurMaxY, PS_SOLID, intFine, mTabGridColor)
        lngCurX = lngCurX + T_DrawClient.列单位
    Next i
ErrOver:
    DrawPartogramTab = lngCurMaxY
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckConnect(ByVal strValue As String, lngWidth As Long, lngHeight As Long) As String
'功能：根据表格的宽度检查字体内容
    Dim strConnect As String, strTmp As String
    Dim i As Integer
    Dim lngX As Long, lngY As Long
    For i = 1 To Len(strValue)
        strTmp = Trim(Mid(strValue, i, 1))
        T_Size.W = mobjDraw.TextWidth(strTmp) / sinTwipsPerPixelX
        T_Size.H = mobjDraw.TextHeight(strTmp) / sinTwipsPerPixelY
        If lngX + T_Size.W > lngWidth Then
            lngY = lngY + T_Size.H
            If lngY > lngHeight Then Exit For
            lngX = 0
            If strConnect = "" Then
                strConnect = strTmp
            Else
                strConnect = strConnect & vbCrLf & strTmp
            End If
        Else
            lngX = lngX + T_Size.W
            strConnect = strConnect & strTmp
        End If
    Next i
    
    CheckConnect = strConnect
End Function

Private Function DrawPartogram(ByVal strPartogram As String) As Single
'----------------------------------------------------------------
'功能：画产程刻度区域和产程曲线区域
'参数：strPartogram :记录名[LPF]记录符[LPF]记录色[LPF]最大值[LPF]最小值[LPF]单位值[LPF]单位[|LPF|]记录名[LPF]记录符...
'------------------------------------------------------------------
    
    Static SlngMaxY As Long                 '记录上一次的最大高度，以决定本次是否需要重画
    Dim lngCurX     As Long, lngCurY As Single  '当前位置
    Dim lngMaxX     As Long, lngMaxY As Single  '边界
    Dim lngRow      As Long
    Dim intLables   As Integer
    '以下都是标准尺度
    Dim intLineMode   As Integer
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim sinRowStep As Single
    Dim arrTemp()     As String, ArrCode() As String, strTmp As String, strConnect As String
    Dim i As Integer
    Dim blnPrinter As Boolean
    Dim intBold As Integer, intFine As Integer
    Dim lngValue As String
    Dim blnInit As Boolean '表示是否第一次显示刻度
    Dim blnDesc As Boolean '用于表示先露高低是否倒序显示
    '以下与绘图区域相关(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    Dim sin刻度 As Single, bln显示刻度 As Boolean
    Dim sin刻度间隔 As Single, sinBegin刻度 As Single, dbl单位值 As Double
    Dim str最大值坐标 As String, str最小值坐标 As String
    
    '参数，产程图模式、先露高低显示位置以及是否显示产程时间
    Dim int模式 As Integer, int先露位置 As Integer
    Dim bln产程时间 As Boolean
    
    On Error GoTo errHand
    If TypeName(mobjDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    
    If blnPrinter = True Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '所有曲线项目的作图区域(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    gstrFields = "项目序号," & adDouble & ",18|最大值," & adDouble & ",18|最小值," & adDouble & ",18|" & "单位值," & adDouble & _
        ",18|最大值坐标," & adLongVarChar & ",20|最小值坐标," & adLongVarChar & ",20|" & "单位刻度," & adLongVarChar & ",20|" & _
        "显示模式," & adDouble & ",5|记录符," & adLongVarChar & ",10|颜色," & adDouble & ",18"
    Call Record_Init(mrsDrawItems, gstrFields)
    '------------------------------------------------------------------------------------------------------------------
    '赋初值
    intLineMode = PS_SOLID
    lngLableStep = T_DrawClient.刻度单位
    lngColStep = T_DrawClient.列单位
    sinRowStep = T_DrawClient.行单位
    
    '产程图模式和先露高低显示位置
    int模式 = Val(zlDatabase.GetPara("产程图模式", glngSys, 1255, 0)) '0-伴行式 1-交叉式
    int先露位置 = Val(zlDatabase.GetPara("先露高低显示位置", glngSys, 1255, 0)) '0-显示在右侧,1-显示在左侧
    bln产程时间 = (Val(zlDatabase.GetPara("产程图显示产程时间", glngSys, 1255, 0)) = 1) '0-不显示,1-显示
    arrTemp = Split(strPartogram, "[|LPF|]")
    '画表格
    intLables = UBound(arrTemp) + 1
    lngCurX = T_DrawClient.偏移量.X
    lngCurY = T_DrawClient.刻度区域.Top
    lngMaxX = T_DrawClient.MaxX '偏移量+刻度区域+24*列单位
    lngMaxY = T_DrawClient.产程区域.Bottom '起始坐标+行数*行单位
        
    SlngMaxY = lngMaxY
    
    If int先露位置 = 1 Then
        Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngMaxY + msnTimeH + IIf(bln产程时间 = True, msnTimeH, 0), PS_SOLID, intFine, RGB_BLACK)
    End If
    lngValue = Val(Split(arrTemp(1), "[LPF]")(3))
    '画产程图所有行
    lngCurX = T_DrawClient.产程区域.Left
    Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngMaxY + msnTimeH + IIf(bln产程时间 = True, msnTimeH, 0), PS_SOLID, intFine, RGB_BLACK)

    For lngRow = 0 To T_DrawClient.总行数
        If lngRow <> 0 Then
            lngCurY = lngCurY + sinRowStep
        End If
        '画产程图的所有行
        Call DrawLine(mlngDC, lngCurX, lngCurY, lngMaxX, lngCurY, PS_SOLID, IIf(lngValue = 0, intBold, intFine), RGB_BLACK)
        lngValue = lngValue - 1
    Next
    
    lngCurY = T_DrawClient.刻度区域.Top
    
    '画产程图所有列
    For lngRow = 1 To 24
        lngCurX = lngCurX + lngColStep
        Call DrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intFine, RGB_BLACK)
    Next

    '画刻度框的标尺（从固定不变的10行开始标识）
'    gstdset.Name = "宋体"
'    gstdset.Size = 9
'    gstdset.Bold = False
'    gstdset.Italic = False
    Call SetFontIndirect(mobjSubFont, mlngDC, mobjDraw)
    mlngFont = CreateFontIndirect(T_Font)
    mlngOldFont = SelectObject(mlngDC, mlngFont)
    
    For lngRow = 0 To UBound(arrTemp)
        ArrCode = Split(arrTemp(lngRow), "[LPF]")
        '显示刻度框项目的名称
        If int先露位置 = 1 Then '先露高低显示在左侧
            strTmp = ArrCode(0) & ArrCode(6)
            lngCurX = T_DrawClient.刻度区域.Left + lngRow * (T_DrawClient.刻度单位 / 2)
            lngCurY = T_DrawClient.刻度区域.Top - Val(Format((Len(strTmp) / 2), "#0")) * mobjDraw.TextHeight(strConnect) / sinTwipsPerPixelY
            For i = 1 To Val(Format((Len(strTmp) / 2), "#0"))
                '产程项目名称
                strConnect = Trim(Mid(strTmp, i * 2 - 1, 2))
                Call SetTextColor(mlngDC, ArrCode(2))
                Call GetTextRect(mobjDraw, lngCurX, lngCurY, Trim(strConnect), T_DrawClient.刻度单位 / 2, False)
                Call DrawText(mlngDC, Trim(strConnect), -1, T_LableRect, DT_CENTER)
                lngCurY = lngCurY + mobjDraw.TextHeight("1") / sinTwipsPerPixelY
            Next i
        Else '先露高低显示在右侧
            strTmp = ArrCode(0) & "(" & ArrCode(6) & ")" '& IIf(int先露位置 = 0, arrCode(1), "")
            If lngRow = 0 Then
                lngCurX = T_DrawClient.刻度区域.Left
            Else
                lngCurX = T_DrawClient.MaxX + T_DrawClient.刻度单位 / 2
            End If
            lngCurY = T_DrawClient.刻度区域.Top + ((T_DrawClient.刻度区域.Bottom - T_DrawClient.刻度区域.Top - ((Len(strTmp) - 1) * mobjDraw.TextHeight("1") / sinTwipsPerPixelY)) / 2)
            For i = 1 To Len(strTmp)
                strConnect = Mid(strTmp, i, 1)
                Call SetTextColor(mlngDC, ArrCode(2))
                Call DrawRotateText(mobjDraw, mlngDC, lngCurX, lngCurY, Trim(strConnect), T_DrawClient.刻度单位 / 2, ArrCode(2))
                If Asc(strConnect) < 0 And strConnect <> "—" Then
                    lngCurY = lngCurY + mobjDraw.TextHeight("1") / sinTwipsPerPixelY
                Else
                    lngCurY = lngCurY + mobjDraw.TextHeight("1") / sinTwipsPerPixelY / 2
                End If
            Next i
        End If
        
        '进行刻度数值X坐标计算
        If int先露位置 = 1 Then
            lngCurX = T_DrawClient.刻度区域.Left + lngRow * (T_DrawClient.刻度单位 / 2)
        Else
            If lngRow = 0 Then
                lngCurX = T_DrawClient.刻度区域.Left + (T_DrawClient.刻度单位 / 2)
            Else
                lngCurX = T_DrawClient.MaxX
            End If
        End If
        lngCurY = T_DrawClient.刻度区域.Top
        dbl单位值 = 1
        sin刻度间隔 = 1
        
        blnDesc = False
        If lngRow = 1 And int模式 = 1 Then blnDesc = True
        blnInit = False
        Do While True
            bln显示刻度 = False
            If blnInit = False Then      '刚进入循环，此时取的最大值
                sin刻度 = IIf(blnDesc = True, Val(ArrCode(4)), Val(ArrCode(3)))
                sinBegin刻度 = sin刻度
                str最大值坐标 = T_DrawClient.产程区域.Left & "," & lngCurY
                blnInit = True
            Else                    '计算得到每个刻度的值
                sin刻度 = sin刻度 - (IIf(blnDesc = True, -1, 1) * dbl单位值)
            End If
            
            '根据设置的刻度间隔显示刻度值
            If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
            If bln显示刻度 = True Then sinBegin刻度 = sinBegin刻度 - (IIf(blnDesc = True, -1, 1) * sin刻度间隔)
            If bln显示刻度 Then
                Call GetTextRect(mobjDraw, lngCurX, lngCurY, Format(sin刻度, "#0"), T_DrawClient.刻度单位 / 2, _
                    IIf(IIf(blnDesc = True, Val(ArrCode(4)), Val(ArrCode(3))) = Val(Format(sin刻度, "#0")), False, True))
                Call DrawText(mlngDC, Format(sin刻度, "#0"), -1, T_LableRect, DT_CENTER)
            End If
            '如果不在有效范围内，或者超出画布则退出
            If Val(Format(sin刻度, "#0.00")) = Val(Format(IIf(blnDesc = True, ArrCode(3), ArrCode(4)), "#0.00")) Or lngCurY > T_DrawClient.刻度区域.Bottom Then
                str最小值坐标 = T_DrawClient.产程区域.Left & "," & lngCurY
                '添加该项目(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
                gstrFields = "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|记录符|颜色"
                gstrValues = IIf(lngRow = 0, T_Partogram.lng宫口扩大, T_Partogram.lng先露高低) & "|" & Val(ArrCode(3)) & "|" & Val(ArrCode(4)) & _
                "|" & dbl单位值 & "|" & str最大值坐标 & "|" & str最小值坐标 & "|" & T_DrawClient.行单位 & "," & T_DrawClient.列单位 & "|" & IIf(blnDesc = True, 1, 0) & "|" & ArrCode(1) & "|" & ArrCode(2)
                Call Record_Add(mrsDrawItems, gstrFields, gstrValues)
                Exit Do
            End If
            lngCurY = lngCurY + T_DrawClient.行单位
        Loop
    Next lngRow
    
    '输出最低下的时间栏
    lngValue = mobjDraw.TextWidth("12") / sinTwipsPerPixelX
    lngCurX = T_DrawClient.刻度区域.Right - (lngValue / 2)
    lngCurY = T_DrawClient.产程区域.Bottom + (msnTimeH / 2)
    Call SetTextColor(mlngDC, RGB_BLACK)
    For lngRow = 1 To 24
        lngCurX = lngCurX + T_DrawClient.列单位
        Call GetTextRect(mobjDraw, lngCurX, lngCurY, lngRow, lngValue)
        Call DrawText(mlngDC, lngRow, -1, T_LableRect, DT_CENTER)
    Next lngRow
    '输出产程时间
    If bln产程时间 = True Then
        lngCurX = T_DrawClient.刻度区域.Left
        lngCurY = T_DrawClient.产程区域.Bottom + msnTimeH + (msnTimeH / 2)
        Call GetTextRect(mobjDraw, lngCurX, lngCurY, "产程时间", T_DrawClient.刻度单位)
        Call DrawText(mlngDC, "产程时间", -1, T_LableRect, DT_CENTER)
        lngCurX = T_DrawClient.产程区域.Left
        Call GetTextRect(mobjDraw, lngCurX, lngCurY, mstr宫缩时间, (mobjDraw.TextWidth(mstr宫缩时间) / sinTwipsPerPixelX) + 2)
        Call DrawText(mlngDC, mstr宫缩时间, -1, T_LableRect, DT_CENTER)
    End If
    Call SelectObject(mlngDC, mlngOldFont)
    Call DeleteObject(mlngFont)

    DrawPartogram = lngCurY + (msnTimeH / 2)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub GetPartogramPage(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFileIndex As Long)
'--------------------------------------------------------------------------
'功能：获取文件页数
'参数;文件ID，病人ID，主页ID，记录组号
'--------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngCol As Long, lngColOld As Long, blnInit As Boolean
    Dim ArrTime, ArrCode() As String
    Dim strTmp As String, intPage As Integer, strEnd As String
    Dim strBeginTime As String
    On Error GoTo errHand
    If Not mblnPrint Then mintMaxPage = 1 '预览打印时不能更新此变量，否则会导致产程图展现页数错误
    intPage = 1
    lngCol = 0
    
    strBeginTime = Format(mstr宫缩时间, "YYYY-MM-DD HH:mm:ss")
    ArrTime = Array()
    mArrPageTime = Array()
    ReDim ArrTime(UBound(ArrTime) + 1)
    ReDim mArrPageTime(UBound(mArrPageTime) + 1)
    ArrTime(UBound(ArrTime)) = strBeginTime & ";" & Format(DateAdd("D", 1, CDate(strBeginTime)), "YYYY-MM-DD HH:mm:ss")
    mArrPageTime(UBound(mArrPageTime)) = Format(strBeginTime, "YYYY-MM-DD HH:mm:ss")
    
'    gstrSQL = _
'        " SELECT 发生时间,FLOOR(TO_NUMBER(发生时间-开始时间)*24)+1 AS 行" & vbNewLine & _
'        " FROM (" & vbNewLine & _
'        " SELECT 发生时间,Max(A.开始时间) 开始时间 FROM 病人护理文件 A,病人护理数据 B,病人护理明细 C" & vbNewLine & _
'        " WHERE A.ID=B.文件ID AND B.ID=C.记录ID AND A.ID=[1] AND A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=0" & vbNewLine & _
'        " AND B.汇总类别=[4]" & vbNewLine & _
'        " GROUP BY B.发生时间 ORDER BY B.发生时间)"
    Call SQLCombinationPage
    Call SQLDIY(mstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "病人护理文件", lngFileID, lngPatiID, lngPageId, 0, lngFileIndex)
    If rsTemp.RecordCount > 0 Then
        With rsTemp
            blnInit = False: lngCol = 0
            Do While Not .EOF
                '73792:刘鹏飞,2014-06-23,表格行位置计算调整
                If blnInit = False Then strEnd = Format(!发生时间, "YYYY-MM-DD HH:mm:ss")
                If lngCol <> Val(!行) Then
                    lngCol = Val(!行)
                    If lngColOld - lngCol >= 0 Then
                        lngColOld = lngColOld + 1 '(2 * lngColOld) - lngCol + 1
                    Else
                        lngColOld = lngCol
                    End If
                Else
                    lngColOld = lngColOld + 1
                End If
                
                blnInit = True
                
                If intPage <> (Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1)) Then
                    intPage = (Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1))
                    ReDim Preserve ArrTime(UBound(ArrTime) + 1)
                    '首先更新上一页的最后时间
                    ArrCode = Split(ArrTime(UBound(ArrTime) - 1), ";")
                    If CDate(strEnd) > CDate(Format(ArrCode(1), "YYYY-MM-DD HH:mm:ss")) Then
                        strEnd = CDate(Format(ArrCode(1), "YYYY-MM-DD HH:mm:ss"))
                    End If
                    ArrTime(UBound(ArrTime) - 1) = ArrCode(0) & ";" & strEnd
                    '更新本页的开始时间
                    strTmp = Format(!发生时间, "YYYY-MM-DD HH:mm:ss")
                    strEnd = Format(DateAdd("H", (intPage * 24) - lngColOld, CDate(strTmp)), "YYYY-MM-DD HH:mm:ss")
                    If CDate(strEnd) > CDate(Format(DateAdd("D", 1, CDate(strBeginTime)), "YYYY-MM-DD HH:mm:ss")) Then
                        strEnd = CDate(Format(DateAdd("D", 1, CDate(strBeginTime)), "YYYY-MM-DD HH:mm:ss"))
                    End If
                    ArrTime(UBound(ArrTime)) = strTmp & ";" & strEnd
                    '更新本页开始时间
                    ReDim Preserve mArrPageTime(UBound(mArrPageTime) + 1)
                    strTmp = DateAdd("n", -1 * ((lngColOld Mod 24) - 1) * 60, CDate(Format(strTmp, "YYYY-MM-DD HH:mm:ss")))
                    mArrPageTime(UBound(mArrPageTime)) = strTmp
                Else
                    strEnd = Format(!发生时间, "YYYY-MM-DD HH:mm:ss")
                End If
            .MoveNext
            Loop
        End With
    End If
    
    If lngColOld <= 0 Then lngColOld = 1
    If Not mblnPrint Then mintMaxPage = Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1)
    mstrTimeRange = ArrTime
    mintPageCount = Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetFilePage(ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal lngFileIndex As Long) As Long
'根据文件号，获取页数
    Dim rsTemp As New ADODB.Recordset
    Dim lngCol As Long, lngColOld As Long
    Dim intPage As Long
    
    On Error GoTo errHand
    
'    gstrSQL = _
'        " SELECT 发生时间,FLOOR(TO_NUMBER(发生时间-开始时间)*24)+1 AS 行" & vbNewLine & _
'        " FROM (" & vbNewLine & _
'        " SELECT 发生时间,MAX(A.开始时间) 开始时间 FROM 病人护理文件 A,病人护理数据 B,病人护理明细 C" & vbNewLine & _
'        " WHERE A.ID=B.文件ID AND B.ID=C.记录ID AND A.ID=[1] AND A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=0" & vbNewLine & _
'        " AND B.汇总类别=[4]" & vbNewLine & _
'        " GROUP BY B.发生时间 ORDER BY B.发生时间)"
    Call SQLCombinationPage
    Call SQLDIY(mstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "病人护理文件", lngFileID, lngPatiID, lngPageId, 0, lngFileIndex)
    If rsTemp.RecordCount > 0 Then
        With rsTemp
            Do While Not .EOF
                If lngCol <> Val(!行) Then
                    lngCol = Val(!行)
                    '73792:刘鹏飞,2014-06-23,表格行位置计算调整
                    If lngColOld - lngCol >= 0 Then
                        lngColOld = lngColOld + 1 ' (2 * lngColOld) - lngCol + 1
                    Else
                        lngColOld = lngCol
                    End If
                Else
                    lngColOld = lngColOld + 1
                End If
            .MoveNext
            Loop
        End With
    End If
    If lngColOld <= 0 Then lngColOld = 1
    intPage = Int(lngColOld / 24) + IIf(lngColOld Mod 24 = 0, 0, 1)
    GetFilePage = intPage
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub SetFontIndirect(ByVal stdSet As StdFont, ByVal lngDc As Long, ByVal ObjDraw As Object)

    '功能:设置字体属性
    Dim BFileName() As Byte
    Dim i As Integer

    On Error GoTo errHand
    
    ObjDraw.Font.Size = stdSet.Size
    BFileName = StrConv(stdSet.Name, vbFromUnicode)
    With T_Font
        For i = 1 To Len(stdSet.Name)
            .lfFaceName(i - 1) = BFileName(i - 1)
        Next i
        .lfHeight = -MulDiv(stdSet.Size, GetDeviceCaps(lngDc, LOGPIXELSY), 71)
        .lfWidth = 0
        .lfWeight = IIf(stdSet.Bold = True, FW_BOLD, FW_NORMAL)
        .lfCharSet = stdSet.Charset
        .lfUnderline = stdSet.Underline
        .lfItalic = stdSet.Italic
        .lfStrikeOut = stdSet.Strikethrough
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub GetTextRect(ByVal ObjDraw As Object, ByVal lngX As Long, ByVal lngY As Long, ByVal strInput As String, _
    Optional ByVal lngWidth As Long = 0, Optional bln居中 As Boolean = True, Optional ByVal lngHeght As Long = 0, Optional ByVal sngScale As Single = 1)
    
    '获取输出字体的有效区域
    
    Dim lngInputW As Long, lng1H As Long
    Dim sngSize As Single
        
    T_LableRect.Left = lngX + 1 '避免与左边界划线重合
    
    If bln居中 = True Then
        T_LableRect.Top = lngY - ObjDraw.TextHeight(strInput) / 2 / sinTwipsPerPixelY
    Else
        T_LableRect.Top = lngY
    End If
    
    T_LableRect.Right = ObjDraw.TextWidth(strInput) / sinTwipsPerPixelY + T_LableRect.Left + 2
    T_LableRect.Bottom = ObjDraw.TextHeight(strInput) / sinTwipsPerPixelY + T_LableRect.Top + 2
    
    If lngWidth <> 0 Then
        '将文本显示在所示宽度的中间区域
        T_LableRect.Left = T_LableRect.Left + (lngWidth - ObjDraw.TextWidth(strInput) / sinTwipsPerPixelY - 1) / 2
        T_LableRect.Right = ObjDraw.TextWidth(strInput) / sinTwipsPerPixelY + T_LableRect.Left + 2
    End If
    
    If lngHeght <> 0 Then
        T_LableRect.Bottom = T_LableRect.Bottom + (lngHeght - ObjDraw.TextHeight(strInput) / sinTwipsPerPixelY)
    End If
    
End Sub


Public Sub DrawLine(ByVal lngDc As Long, ByVal lngSX As Long, ByVal lngSY As Long, ByVal lngDX As Long, ByVal lngDY As Long, _
    Optional ByVal lngType As Long = PS_SOLID, Optional ByVal intWidth As Integer = 1, Optional ByVal lngRGB As Long = 0, _
    Optional ByVal blnEndRow As Boolean = False, Optional ByVal blnPrinter As Boolean = False)
    
    Dim X As Long
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim sngX As Single, sngY As Single
    On Error GoTo errHand
    '创建新画笔进行划线
    
    If msngTwips = 0 Then msngTwips = 1
    sngX = 3 * msngTwips
    sngY = 4 * msngTwips

    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDc, lngPen)
    '绘图
    Call MoveToEx(lngDc, lngSX, lngSY, T_OldPoint)
    Call LineTo(lngDc, lngDX, lngDY)
    '对于物理降温话上下箭头
    If blnEndRow Then
        If lngSY > lngDY Then '向上箭头
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDc, X, lngDY + sngY, T_OldPoint)
                Call LineTo(lngDc, lngDX, lngDY)
            Next X
        Else '向下箭头
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDc, X, lngDY - sngY, T_OldPoint)
                Call LineTo(lngDc, lngDX, lngDY)
            Next X
        End If
    End If
    
    '还原画笔并销毁
    Call SelectObject(lngDc, lngOldPen)
    Call DeleteObject(lngPen)
    lngPen = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub DrawRect(ByVal lngDc As Long, ByVal lngSX As Long, ByVal lngSY As Long, ByVal lngDX As Long, ByVal lngDY As Long, _
    Optional ByVal lngType As Long = PS_SOLID, Optional ByVal intWidth As Integer = 1, Optional ByVal lngRGB As Long = 0)
    
    Dim lngPen As Long, lngOldPen As Long
    On Error GoTo errHand
    '创建新画笔进行画一个矩形
    
    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDc, lngPen)
    '绘图
    Call Rectangle(lngDc, lngSX, lngSY, lngDX, lngDY)
    '还原画笔并销毁
    Call SelectObject(lngDc, lngOldPen)
    Call DeleteObject(lngPen)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub DrawRotateText(ByVal ObjDraw As Object, ByVal lngDc As Long, ByVal X As Single, _
                          ByVal Y As Single, _
                          ByVal strText As String, _
                          ByVal sngWidth As Single, _
                          Optional ByVal ForeColor As Long = 0, _
                          Optional ByVal sglScale As Single = 1)

    '在(X,Y)处输出Text文本
    Dim objFont    As Object
    Dim lngFont    As Long
    Dim lngOldFont As Long
    Dim X1         As Long
    Dim blnPrinter As Boolean
    
    If strText = "" Then Exit Sub
    
    If TypeName(ObjDraw) = "Printer" Then
        blnPrinter = True
    Else
        blnPrinter = False
    End If
    '设置文本颜色
    Call SetTextColor(lngDc, ForeColor)
    
    '正常输出字体
    If Asc(strText) < 0 And strText <> "—" Then
    
        Call GetTextRect(ObjDraw, X, Y, strText, sngWidth, False)
        Call DrawText(lngDc, strText, -1, T_LableRect, DT_CENTER)
        
    Else '反转90度输出字体
        '在打印是反转输出字体 objDraw.TextWidth 必须放在创建字体之前，否则无法发转。
        Call GetTextRect(ObjDraw, X, Y, strText, sngWidth, False)
        X1 = X + (ObjDraw.TextWidth("字") / sinTwipsPerPixelX) + (T_LableRect.Left - X) - (IIf(blnPrinter = True, 2, 1) * msngTwips)
        Set objFont = New clsRotateFont
        Set objFont.LogFont = mobjSubFont
        
        objFont.sngTwpic = msngTwips
        objFont.Rotation = -90
        lngFont = objFont.Handle
        lngOldFont = SelectObject(lngDc, lngFont)
        Call TextOut(lngDc, X1, Y, strText, LenB(StrConv(strText, vbFromUnicode)))
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
    End If
End Sub

Public Sub ShowFlash(ByVal blnPrint As Boolean, Optional strInfo As String, Optional sngPer As Single, Optional frmParent As Object)
    '功能：显示或隐藏等待或进度窗体(strInfo)
    '参数:strInfo=进度提示信息
    '     sngPer=进度
    '     blnPrint=true 表示显示此窗体，false不显示
    If Not blnPrint Then Exit Sub
    Static blnShow As Boolean
    If sngPer > 1 Then sngPer = 1

    If strInfo = "" Then
        Unload frmFlash
        blnShow = False
    Else
        If Not blnShow Then
            On Error Resume Next
            frmFlash.lbl.Top = frmFlash.lbl.Top - frmFlash.lbl.Height / 2
            frmFlash.lblPer.Top = frmFlash.lbl.Top
            frmFlash.lbl.Caption = strInfo
            frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)

            If sngPer > 0 Then
                frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
            Else
                frmFlash.lblPer.Caption = ""
            End If

            If frmParent Is Nothing Then
                SetWindowPos frmFlash.Hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                ShowWindow frmFlash.Hwnd, 5
            Else
                Err.Clear
                frmFlash.Show , frmParent
                If Err.Number <> 0 Then
                    Err.Clear
                    SetWindowPos frmFlash.Hwnd, -1, (Screen.Width - frmFlash.Width) / 2 / Screen.TwipsPerPixelX, (Screen.Height - frmFlash.Height) / 2 / Screen.TwipsPerPixelY, 0, 0, 1
                    ShowWindow frmFlash.Hwnd, 5
                End If
            End If

            frmFlash.Refresh
            blnShow = True
        Else
            frmFlash.lbl.Caption = strInfo
            If sngPer >= 0 Then
                frmFlash.lblDo.Caption = String(25 * sngPer, frmFlash.lblDo.Tag)
                If sngPer > 0 Then
                    frmFlash.lblPer.Caption = Int(sngPer * 100) & "%"
                Else
                    frmFlash.lblPer.Caption = ""
                End If
            End If
            frmFlash.Refresh
        End If
    End If
End Sub

Public Function GetPaperName(intSize As Integer) As String
    '功能： 根据当前打印机的设置，获取纸张名称
    '返回： 纸张名称
    If intSize = 256 Then
        GetPaperName = "用户自定义 ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        GetPaperName = Switch( _
        intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41)
    Else
        GetPaperName = "不可测的纸张 ..."
    End If
End Function

