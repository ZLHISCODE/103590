Attribute VB_Name = "mdlWaveStylePrint"
Option Explicit

'***************************************************************
'绘画相关API及常量、结构定义
'***************************************************************
'结构定义
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type Size
    W   As Long
    H   As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type

'字体属性
Private Const LF_FACESIZE = 32

Private Type LogFont
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

Private T_OldPoint   As POINTAPI
Private T_NewPoint   As POINTAPI
Private T_ClientRect As RECT
Private T_LableRect  As RECT      '待输出文本的有效区域
Private T_ControlRect As RECT     '窗体菜单有效区域
Private T_Brush      As LOGBRUSH
Private T_Font       As LogFont
Private T_Size       As Size

'创建或得到现有对象
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateBitmap _
               Lib "gdi32" (ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            ByVal nPlanes As Long, _
                            ByVal nBitCount As Long, _
                            lpBits As Any) As Long

Private Declare Function CreateCompatibleBitmap _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long) As Long
                            
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

'创建画笔、刷子
Private Declare Function CreatePen _
               Lib "gdi32" (ByVal nPenStyle As Long, _
                            ByVal nWidth As Long, _
                            ByVal crColor As Long) As Long

Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long

Private Declare Function ExtCreatePen _
               Lib "gdi32" (ByVal dwPenStyle As Long, _
                            ByVal dwWidth As Long, _
                            lplb As LOGBRUSH, _
                            ByVal dwStyleCount As Long, _
                            lpStyle As Long) As Long

Private Const PS_SOLID = 0
Private Const PS_DASH = 1                    '  -------
Private Const PS_DOT = 2                     '  .......
Private Const PS_DASHDOT = 3                 '  _._._._
Private Const PS_DASHDOTDOT = 4              '  _.._.._
Private Const PS_NULL = 5                    '不允许画图
Private Const PS_COSMETIC = &H0
Private Const PS_GEOMETRIC = &H10000
Private Const PS_ALTERNATE = 8
Private Const PS_ENDCAP_FLAT = &H200
Private Const PS_ENDCAP_MASK = &HF00
Private Const PS_ENDCAP_ROUND = &H0
Private Const PS_ENDCAP_SQUARE = &H100
Private Const PS_JOIN_BEVEL = &H1000
Private Const PS_JOIN_MASK = &HF000
Private Const PS_JOIN_MITER = &H2000
Private Const PS_JOIN_ROUND = &H0

'CreateSolidBrush 创建纯色画刷
'CreateBrushIndirect 通过 LOGBRUSH 类型创建画刷
'CreateHatchBrush 创建阴影画刷
'CreatePatternBrush 创建图案画刷
'GetSysColorBrush 创建系统标准色画刷
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

'//lbStyle　可选值:
Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8
Private Const BS_MONOPATTERN = 9

'//lbHatch　可选值:
Private Const HS_HORIZONTAL = 0              '  -----
Private Const HS_VERTICAL = 1                '  |||||
Private Const HS_FDIAGONAL = 2               '  \\\\\
Private Const HS_BDIAGONAL = 3               '  /////
Private Const HS_CROSS = 4                   '  +++++
Private Const HS_DIAGCROSS = 5               '  xxxxx

Private Declare Function CreateHatchBrush _
               Lib "gdi32" (ByVal nIndex As Long, _
                            ByVal crColor As Long) As Long
'nIndex,同上面函数的lbHatch
'Private Const HS_HORIZONTAL = 0              '  -----
'Private Const HS_VERTICAL = 1                '  |||||
'Private Const HS_FDIAGONAL = 2               '  \\\\\
'Private Const HS_BDIAGONAL = 3               '  /////
'Private Const HS_CROSS = 4                   '  +++++
'Private Const HS_DIAGCROSS = 5               '  xxxxx

Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

'BLACK_BRUSH：黑色画笔
'DKGRAY_BRUSH：暗灰色画笔
'GRAY_BRUSH：灰色画笔
'HOLLOW_BRUSH：空画笔（相当于HOLLOW_BRUSH）
'LTGRAY_BRUSH：亮灰色画笔
'NULL_BRUSH：空画笔（相当于HOLLOW_BRUSH）
'WHITE_BRUSH：白色画笔
'BLACK_PEN：黑色钢笔
'WHITE_PEN：白色钢笔
Private Const WHITE_BRUSH = 0    '白色画笔
Private Const LTGRAY_BRUSH = 1   '亮灰色画笔
Private Const GRAY_BRUSH = 2     '灰色画笔
Private Const DKGRAY_BRUSH = 3   '暗灰色画笔
Private Const BLACK_BRUSH = 4    '黑色画笔
Private Const NULL_BRUSH = 5
Private Const HOLLOW_BRUSH = NULL_BRUSH
Private Const WHITE_PEN = 6      '白色钢笔
Private Const BLACK_PEN = 7      '黑色钢笔

'创建一个区域
Private Declare Function CreateEllipticRgn _
               Lib "gdi32" (ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long

Private Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Private Declare Function CreateRectRgn _
               Lib "gdi32" (ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long

Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long

Private Declare Function CreateRoundRectRgn _
               Lib "gdi32" (ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long, _
                            ByVal X3 As Long, _
                            ByVal Y3 As Long) As Long

'以下是释放对象函数
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SelectObject _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal hObject As Long) As Long

Private Declare Function ReleaseDC _
               Lib "user32" (ByVal hWnd As Long, _
                             ByVal hDC As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

'以下是功能函数
Private Declare Function DrawFocusRect _
               Lib "user32" (ByVal hDC As Long, _
                             lpRect As RECT) As Long

Private Declare Function Ellipse _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long

Private Declare Function CreatePolygonRgn _
               Lib "gdi32" (lpPoint As POINTAPI, _
                            ByVal nCount As Long, _
                            ByVal nPolyFillMode As Long) As Long

Private Const ALTERNATE = 1

Private Const WINDING = 2

Private Declare Function ExtFloodFill _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal crColor As Long, _
                            ByVal wFillType As Long) As Long

Private Const FLOODFILLBORDER = 0

Private Const FLOODFILLSURFACE = 1

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function Rectangle _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long

Private Declare Function FillRect _
               Lib "user32" (ByVal hDC As Long, _
                             lpRect As RECT, _
                             ByVal hBrush As Long) As Long

Private Declare Function FillRgn _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal hRgn As Long, _
                            ByVal hBrush As Long) As Long

Private Declare Function Polyline _
               Lib "gdi32" (ByVal hDC As Long, _
                            lpPoint As POINTAPI, _
                            ByVal nCount As Long) As Long

Private Declare Function FloodFill _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal crColor As Long) As Long

Private Declare Function GetArcDirection Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function GetBrushOrgEx _
               Lib "gdi32" (ByVal hDC As Long, _
                            lpPoint As POINTAPI) As Long

Private Declare Function GetCurrentObject _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal uObjectType As Long) As Long

Private Declare Function GetCurrentPositionEx _
               Lib "gdi32" (ByVal hDC As Long, _
                            lpPoint As POINTAPI) As Long

Private Declare Function GetPixel _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Long

Private Declare Function InvertRect _
               Lib "user32" (ByVal hDC As Long, _
                             lpRect As RECT) As Long

Private Declare Function LineTo _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Long

Private Declare Function GetClientRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT) As Long

Private Declare Function GetWindowRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT) As Long

Private Declare Function BitBlt _
               Lib "gdi32" (ByVal hDestDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            ByVal hSrcDC As Long, _
                            ByVal xSrc As Long, _
                            ByVal ySrc As Long, _
                            ByVal dwRop As Long) As Long

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Private Declare Function MoveToEx _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            lpPoint As POINTAPI) As Long

Private Declare Function Polygon _
               Lib "gdi32" (ByVal hDC As Long, _
                            lpPoint As POINTAPI, _
                            ByVal nCount As Long) As Long

Private Declare Function PtInRegion _
               Lib "gdi32" (ByVal hRgn As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long) As Long

Private Declare Function SetTextColor _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal crColor As Long) As Long

Private Declare Function SetBkMode _
               Lib "gdi32" (ByVal hDC As Long, _
                            ByVal nBkMode As Long) As Long

Private Const TRANSPARENT = 1

Private Declare Function DrawText _
               Lib "user32" _
               Alias "DrawTextA" (ByVal hDC As Long, _
                                  ByVal lpStr As String, _
                                  ByVal nCount As Long, _
                                  lpRect As RECT, _
                                  ByVal wFormat As Long) As Long

Private Const DT_CENTER = &H1

Private Declare Function TextOut _
               Lib "gdi32" _
               Alias "TextOutA" (ByVal hDC As Long, _
                                 ByVal X As Long, _
                                 ByVal Y As Long, _
                                 ByVal lpString As String, _
                                 ByVal nCount As Long) As Long

Private Declare Function CreateFont _
               Lib "gdi32" _
               Alias "CreateFontA" (ByVal H As Long, _
                                    ByVal W As Long, _
                                    ByVal E As Long, _
                                    ByVal O As Long, _
                                    ByVal W As Long, _
                                    ByVal i As Long, _
                                    ByVal u As Long, _
                                    ByVal s As Long, _
                                    ByVal C As Long, _
                                    ByVal OP As Long, _
                                    ByVal CP As Long, _
                                    ByVal q As Long, _
                                    ByVal PAF As Long, _
                                    ByVal f As String) As Long

Private Declare Function GetUpdateRect _
               Lib "user32" (ByVal hWnd As Long, _
                             lpRect As RECT, _
                             ByVal bErase As Long) As Long

'用指定属性创建一种逻辑字体
Private Declare Function CreateFontIndirect _
                Lib "gdi32" _
                Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700

'获取字体的高度,获取汉字的宽度不准
Private Declare Function GetTextExtentPoint32 _
               Lib "gdi32" _
               Alias "GetTextExtentPoint32A" (ByVal hDC As Long, _
                                              ByVal lpsz As String, _
                                              ByVal cbString As Long, _
                                              lpSize As Size) As Long

'nNumber*nNumerator/nDenominator 自动四舍五入。无法计算的返回-1
Private Declare Function MulDiv _
               Lib "kernel32" (ByVal nNumber As Long, _
                               ByVal nNumerator As Long, _
                               ByVal nDenominator As Long) As Long
'说明
'根据指定设备场景代表的设备的功能返回信息
'参数 类型及说明
'hdc Long，要查询其设备的信息的设备场景
'nIndex Long，根据GetDeviceCaps索引表所示常数确定返回信息的类型

Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90

Private RGB_BLACK          As Long
Private RGB_RED            As Long
Private RGB_WRITE          As Long
Private RGB_BLUE          As Long
Private RGB_GRAY          As Long
Private RGB_FleetGRAY     As Long
'-------------------------------------------------------------------------------------------
Private Const mintCurveNullRow As Integer = 2
Private Const mlngBreathRowHeight As Long = 300 '呼吸表格高度(缇)
Public Const mlngWaveLeft As Long = 180 '体温单展示左边距(缇)
Public Const mlngWaveTop As Long = 180 '体温单展示右边距(缇)

Private mbln呼吸表格 As Boolean '呼吸是否为表格项目
'基本属性
Private Type BasicStyle
    TitleText As String '标题内容
    TitleFont As String '标题字体
    TabRowHeight As Long '一般项目栏高度
    TabDays As Long '监测天数
    TabDayTime As Long '监测次数
    TabBeginTime As Long '开始时点
    TabTimeSplit As Long '时间间隔
    TabTitleName As String '一般项目栏表头名称
    ScaleColWidth As Long '刻度区域总宽度
    CurveColWidth As Long '绘图区域表格宽度
    CurveRowHeight As Long '绘图区域表格高度
    AddCurveNull As Long '绘图区域曲线添加的空行数(只针对曲线，不针对独立曲线)
    DownTabRowHeight As Long '特殊项目栏表格高度
    AddTabNull As Long '特殊项目栏表格添加的空行数
    BlnBaby As Boolean '是否婴儿体温单
End Type
Private TBasicStyle As BasicStyle
'体温单(表格高度、刻度宽度、曲线表格高度宽度等变量)单位像素
Private Type WaveDrawStyle
    上表格高度 As Long
    下表格高度 As Long
    刻度总宽度 As Long
    曲线列宽度 As Long
    曲线列高度 As Long
    曲线总行数 As Long '不包含记录法=3
    表下表格总行数 As Long
    呼吸表格高度 As Long '呼吸为表格时有效
End Type
Private TWaveDrawStyle As WaveDrawStyle
'曲线项目 格式:项目序号|项目序号
Private mstrCurveItem As String
'表格项目 格式:项目序号;记录频次|项目序号;记录频次
Private mstrTabItem As String
'绘图对象的DC
Private mlngDC As Long
Private mlngFont As Long
Private mlngOldFont As Long

'绘图的实际区域(缇)
Private mlngWaveWidth As Long
Private mlngWaveHeight As Long

Public Property Get WaveWidth() As Long
    WaveWidth = mlngWaveWidth
End Property

Public Property Get WaveHeight() As Long
    WaveHeight = mlngWaveHeight
End Property

'-------------------------------------------------------------------------------------------
'体温单样式绘图
'-------------------------------------------------------------------------------------------
Public Function DrawWaveStyle(objPrint As Object, ByVal rsStyle As ADODB.Recordset, Optional ByVal blnExamples As Boolean = False, Optional ByRef mlngHeight As Long) As Boolean
'-----------------------------------------------------------
'功能：根据构造的体温单样式(病历文件结构)完成体温单样板绘图工作
'参数: objPrint 绘图设备对象,
'      rsStyle  体温单样式记录集
'      blnExamples 是否输出"示例"字体,专科体温单默认不用输出
'-----------------------------------------------------------
    Dim strTmp As String, strSQL As String
    Dim lngId As Long
    Dim arrItem() As String, arrCode() As String, lngIndex As Long, lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    Dim objFont As StdFont
    '绘图
    Dim lngWidth As Long, lngHeight As Long, lngLeft As Long, lngTop As Long
    Dim lngBrush As Long, lngOldBrush As Long
    Dim lngX As Long, lngY As Long, lngCurX As Long, lngCurY As Long
    Dim intBold As Integer, intFine As Integer
    '曲线部分
    Dim lngCurveRows As Long, lngMaxValue As Long, lngMinValue As Long
    '曲线表格部分
    Dim lngTabRows As Long
    '体温项目
    Dim strCurveItem As String
    Dim rsCurve As New ADODB.Recordset
    
    On Error GoTo errHand
    mbln呼吸表格 = False
    If rsStyle Is Nothing Or objPrint Is Nothing Then Exit Function
    If rsStyle.State = adStateClosed Then Exit Function
    
    If TypeName(objPrint) = "Printer" Then
        intBold = 4
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    
    RGB_BLACK = RGB(0, 0, 0)
    RGB_RED = RGB(255, 0, 0)
    RGB_WRITE = RGB(255, 255, 255)
    RGB_BLUE = RGB(0, 0, 255)
    RGB_GRAY = &H808080
    RGB_FleetGRAY = &HC0C0C0
    '------------------------------------------------------------------------------------
    '第一步:基础数据准备
    '------------------------------------------------------------------------------------
    With TBasicStyle
        .TitleText = "XX体温单"
        .TitleFont = "宋体,9"
        .TabRowHeight = 225
        .TabDays = 7
        .TabDayTime = 6
        .TabBeginTime = 4
        .TabTimeSplit = 4
        .TabTitleName = "日    期@住院天数@手术后天数&时    间"
        .ScaleColWidth = 1035
        .CurveColWidth = 180
        .CurveRowHeight = 90
        .AddCurveNull = 0
        .DownTabRowHeight = 225
        .AddTabNull = 0
    End With
    
    mstrCurveItem = "1"
    mstrTabItem = "1"
    '获取样式基本属性
    rsStyle.Filter = "父ID=NULL And 对象序号=1 And 内容文本='格式定义'"
    If rsStyle.RecordCount > 0 Then
        lngId = rsStyle!ID
        rsStyle.Filter = "父ID=" & lngId
        Do While Not rsStyle.EOF
            Select Case "" & rsStyle!要素名称
            Case "标题文本"
                TBasicStyle.TitleText = "" & rsStyle!内容文本
            Case "标题字体"
                TBasicStyle.TitleFont = "" & rsStyle!内容文本
                If TBasicStyle.TitleFont = "" Then TBasicStyle.TitleFont = "宋体,9"
            Case "表格高度"
                TBasicStyle.TabRowHeight = Val("" & rsStyle!内容文本)
                If TBasicStyle.TabRowHeight < 225 Or TBasicStyle.TabRowHeight > 600 Then
                    TBasicStyle.TabRowHeight = 225
                End If
            Case "天数"
                TBasicStyle.TabDays = Val("" & rsStyle!内容文本)
                If TBasicStyle.TabDays = 0 Then TBasicStyle.TabDays = 7
            Case "监测次数"
                If InStr(1, ",2,4,6,8,12,24,", "," & Val("" & rsStyle!内容文本) & ",") = 0 Then
                    TBasicStyle.TabDayTime = 6
                Else
                    TBasicStyle.TabDayTime = Val("" & rsStyle!内容文本)
                End If
            Case "开始时点"
                TBasicStyle.TabBeginTime = Val("" & rsStyle!内容文本)
            Case "时间间隔"
                TBasicStyle.TabTimeSplit = Val("" & rsStyle!内容文本)
            Case "列头名称"
                TBasicStyle.TabTitleName = "" & rsStyle!内容文本
            Case "刻度宽度"
                TBasicStyle.ScaleColWidth = Val("" & rsStyle!内容文本)
            Case "曲线列宽"
                TBasicStyle.CurveColWidth = Val("" & rsStyle!内容文本)
            Case "曲线行高"
                TBasicStyle.CurveRowHeight = Val("" & rsStyle!内容文本)
            Case "曲线空行"
                TBasicStyle.AddCurveNull = Val("" & rsStyle!内容文本)
                If TBasicStyle.AddCurveNull < 0 Then TBasicStyle.AddCurveNull = 0
            Case "表格高度1"
                TBasicStyle.DownTabRowHeight = Val("" & rsStyle!内容文本)
                If TBasicStyle.DownTabRowHeight < 225 Or TBasicStyle.TabRowHeight > 600 Then
                    TBasicStyle.DownTabRowHeight = 225
                End If
            Case "表格空行"
                TBasicStyle.AddTabNull = Val("" & rsStyle!内容文本)
                If TBasicStyle.AddTabNull < 0 Then TBasicStyle.AddTabNull = 0
            Case "婴儿体温单"
                TBasicStyle.BlnBaby = "" & rsStyle!内容文本
            End Select
        rsStyle.MoveNext
        Loop
    End If
    With TWaveDrawStyle
        .上表格高度 = Fix(objPrint.ScaleX(TBasicStyle.TabRowHeight, vbTwips, vbPixels))
        .下表格高度 = Fix(objPrint.ScaleX(TBasicStyle.DownTabRowHeight, vbTwips, vbPixels))
        .刻度总宽度 = Fix(objPrint.ScaleX(TBasicStyle.ScaleColWidth, vbTwips, vbPixels))
        .曲线列高度 = Fix(objPrint.ScaleX(TBasicStyle.CurveRowHeight, vbTwips, vbPixels))
        .曲线列宽度 = Fix(objPrint.ScaleX(TBasicStyle.CurveColWidth, vbTwips, vbPixels))
        .呼吸表格高度 = Fix(objPrint.ScaleX(mlngBreathRowHeight, vbTwips, vbPixels))
    End With
    '获取曲线项目信息
    rsStyle.Filter = "父ID=NULL And 对象序号=2 And 内容文本='曲线项目定义'"
    If rsStyle.RecordCount > 0 Then
        lngId = rsStyle!ID
        rsStyle.Filter = "父ID=" & lngId
        strTmp = ""
        Do While Not rsStyle.EOF
            strTmp = strTmp & "|" & "" & rsStyle!内容文本
        rsStyle.MoveNext
        Loop
        If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
        If strTmp <> "" Then mstrCurveItem = strTmp
    End If
    '获取表格项目信息
    rsStyle.Filter = "父ID=NULL And 对象序号=3 And 内容文本='表格项目定义'"
    If rsStyle.RecordCount > 0 Then
        lngId = rsStyle!ID
        rsStyle.Filter = "父ID=" & lngId
        strTmp = ""
        Do While Not rsStyle.EOF
                strTmp = strTmp & "|" & "" & rsStyle!内容文本 & ";" & "" & rsStyle!要素表示
        rsStyle.MoveNext
        Loop
        If Left(strTmp, 1) = "|" Then strTmp = Mid(strTmp, 2)
        If strTmp <> "" Then mstrTabItem = strTmp
    End If
    If mstrCurveItem = "" Then mstrCurveItem = "1"
    '提取绑定的项目
    strCurveItem = Replace(mstrCurveItem, "|", ",")
    If mstrTabItem = "" Then mstrTabItem = ";"
    arrItem = Split(mstrTabItem, "|")
    For lngIndex = 0 To UBound(arrItem)
        strCurveItem = strCurveItem & "," & Split(arrItem(lngIndex), ";")(0)
    Next lngIndex
    strSQL = _
        " SELECT /*+ RULE */" & vbNewLine & _
        "  A.项目序号, A.排列序号,DECODE(A.项目序号,4,'血压',A.记录名) 记录名,A.记录法,A.记录符, A.记录色, NVL(A.最大值, 0) 最大值, NVL(A.最小值, 0) 最小值, NVL(A.单位值, 0) 单位值, A.刻度间隔, A.警示线, B.项目单位 单位," & vbNewLine & _
        "  Decode(A.记录法,3,A.最高行,nvl(A.最高行,2)-2) AS 最高行,DECODE(NVL(C.项目序号,''),'',B.项目表示,4) 项目表示" & vbNewLine & _
        " FROM 护理记录项目 B, 体温记录项目 A,护理波动项目 C" & vbNewLine & _
        " WHERE B.项目序号 = A.项目序号 And B.项目序号=C.项目序号(+) And B.项目序号<>5  AND B.项目性质=1 AND NOT (NVL(B.应用方式,0)=2 And B.项目序号=-1) AND EXISTS" & vbNewLine & _
        "  (SELECT 1 FROM TABLE(CAST(F_NUM2LIST([1]) AS ZLTOOLS.T_NUMLIST)) WHERE COLUMN_VALUE = B.项目序号)" & vbNewLine & _
        " ORDER BY A.排列序号"
    Set rsCurve = zlDatabase.OpenSQLRecord(strSQL, "提取体温曲线项目", strCurveItem)
    '计算曲线表格项目共有多少行
    lngCurveRows = 0
    rsCurve.Filter = "项目序号=1"
    If rsCurve.RecordCount > 0 Then
        lngCurveRows = Val(NVL(rsCurve!最高行, 0))
        If lngCurveRows < 0 Then lngCurveRows = 0
        lngMaxValue = Val(NVL(rsCurve!最大值, 0))
        If lngMaxValue < 42 Then lngMaxValue = 42
        lngMinValue = Val(NVL(rsCurve!最小值, 0))
        If lngMinValue > 35 Then lngMinValue = 35
        '固定加两行用于输入项目名称
        lngCurveRows = lngCurveRows + ((lngMaxValue - lngMinValue) / 0.1) + mintCurveNullRow + TBasicStyle.AddCurveNull
        TWaveDrawStyle.曲线总行数 = lngCurveRows
    End If
    rsCurve.Filter = "记录法=3 And 项目序号<>1"
    rsCurve.Sort = "排列序号"
    Do While Not rsCurve.EOF
        lngMaxValue = Val(zlCommFun.NVL(rsCurve!最大值, 0))
        lngMinValue = Val(zlCommFun.NVL(rsCurve!最小值, 0))
        lngRow = ((lngMaxValue - lngMinValue) / Val(NVL(rsCurve!单位值)))
        If Val(NVL(rsCurve!最高行, 0)) > 0 Then lngRow = lngRow + Val(NVL(rsCurve!最高行, 0))
        If lngRow Mod 2 = 1 Then lngRow = lngRow + 1
        lngCurveRows = lngCurveRows + lngRow
    rsCurve.MoveNext
    Loop
    '计算曲线表格有多少行
    lngTabRows = TBasicStyle.AddTabNull
    rsCurve.Filter = "记录法=2 And 项目序号<>5"
    mstrTabItem = ""
    Do While Not rsCurve.EOF
        For lngIndex = 0 To UBound(arrItem)
            arrCode = Split(arrItem(lngIndex), ";")
            If Val(arrCode(0)) = Val(rsCurve!项目序号) Then
                If Val(rsCurve!项目序号) = 3 Then '说明呼吸为表格项目
                    lngTabRows = lngTabRows + 1
                    mbln呼吸表格 = True
                    arrCode(1) = TBasicStyle.TabDayTime
                Else
                    If InStr(1, GetTabFrequency(Val(TBasicStyle.TabDayTime), Val(rsCurve!项目表示)), Val(arrCode(1))) = 0 Then
                        arrCode(1) = IIf(Val(TBasicStyle.TabDayTime) > 2, 2, 1)
                    End If
                    Select Case Val(arrCode(1))
                    Case 3
                        lngTabRows = lngTabRows + 3
                    Case 4
                        lngTabRows = lngTabRows + 2
                    Case Else
                        lngTabRows = lngTabRows + 1
                    End Select
                End If
                mstrTabItem = mstrTabItem & "|" & Join(arrCode, ";")
                Exit For
            End If
        Next lngIndex
    rsCurve.MoveNext
    Loop
    If Left(mstrTabItem, 1) = "|" Then mstrTabItem = Mid(mstrTabItem, 2)
    If mstrTabItem = "" Then mstrTabItem = ";"
    arrItem = Split(mstrTabItem, "|")
    
    TWaveDrawStyle.表下表格总行数 = lngTabRows
    lngLeft = mlngWaveLeft: lngTop = mlngWaveTop  '单位缇
'    lngLeft=350
    '计算宽度
    lngWidth = objPrint.ScaleX(TWaveDrawStyle.刻度总宽度, vbPixels, vbTwips) + TBasicStyle.TabDays * (TBasicStyle.TabDayTime * objPrint.ScaleX(TWaveDrawStyle.曲线列宽度, vbPixels, vbTwips))
    lngWidth = lngWidth + lngLeft * 2
    '计算高度
    lngHeight = 4 * objPrint.ScaleY(TWaveDrawStyle.上表格高度, vbPixels, vbTwips) + lngCurveRows * objPrint.ScaleY(TWaveDrawStyle.曲线列高度, vbPixels, vbTwips) + lngTabRows * objPrint.ScaleY(TWaveDrawStyle.下表格高度, vbPixels, vbTwips)
    lngHeight = lngHeight - IIf(mbln呼吸表格 = True, (objPrint.ScaleY(TWaveDrawStyle.下表格高度, vbPixels, vbTwips) - objPrint.ScaleY(TWaveDrawStyle.呼吸表格高度, vbPixels, vbTwips)), 0)
    Set objFont = New StdFont
    With objFont
        .Name = "宋体"
        .Size = 9
        .Bold = False: .Italic = False
    End With
    Set objPrint.Font = objFont
    lngHeight = lngHeight + objPrint.TextHeight("刘") * 2
    arrItem = Split(TBasicStyle.TitleFont, ",")
    Set objFont = New StdFont
    With objFont
        .Name = arrItem(0)
        .Size = 9
        If UBound(arrItem) > 0 Then .Size = Val(arrItem(1))
        .Bold = False: .Italic = False
        If InStr(1, TBasicStyle.TitleFont, "粗") > 0 Then .Bold = True
        If InStr(1, TBasicStyle.TitleFont, "斜") > 0 Then .Italic = True
    End With
    Set objPrint.Font = objFont
    lngHeight = lngHeight + objPrint.TextHeight(TBasicStyle.TitleText) + lngTop * 2
    mlngWaveWidth = lngWidth - lngLeft * 2: mlngWaveHeight = lngHeight - lngTop * 2
    objPrint.Width = lngWidth: objPrint.Height = lngHeight
    '获取dc时一定要注意设置完对象的宽度高度后在获取，否则不能绘图成功
    mlngDC = objPrint.hDC
    '------------------------------------------------------------------------------------
    '第一步:开始进行绘图操作
    '------------------------------------------------------------------------------------
    '--ONE:先象绘图对象清空
    T_ClientRect.Left = 0: T_ClientRect.Right = objPrint.Width
    T_ClientRect.Top = 0: T_ClientRect.Bottom = objPrint.Height
    '创建白色刷子
    lngBrush = GetStockObject(WHITE_BRUSH)
    '使用该刷子填充背景色（全白）
    lngOldBrush = SelectObject(mlngDC, lngBrush)
    Call FillRect(mlngDC, T_ClientRect, lngBrush)
    '立即销毁临时使用的刷子并还原刷子
    Call SelectObject(mlngDC, lngOldBrush)
    Call DeleteObject(lngBrush)
    Call SetTextColor(mlngDC, RGB_BLACK)
    '--TWO:一般项目栏信息的输出
    lngX = objPrint.ScaleX(lngLeft, vbTwips, vbPixels)
    lngY = objPrint.ScaleY(lngTop, vbTwips, vbPixels)
    '标题输出
    Call SetFontIndirect(objPrint, TBasicStyle.TitleFont)
    T_Size.W = objPrint.ScaleX(objPrint.TextWidth(TBasicStyle.TitleText), vbTwips, vbPixels)
    T_Size.H = objPrint.ScaleY(objPrint.TextHeight(TBasicStyle.TitleText), vbTwips, vbPixels)
    lngCurX = 0: lngCurY = lngY + T_Size.H \ 2
    Call GetTextRect(objPrint, lngCurX, lngCurY, TBasicStyle.TitleText, objPrint.ScaleX(objPrint.Width, vbTwips, vbPixels), True)
    Call DrawText(mlngDC, TBasicStyle.TitleText, -1, T_LableRect, DT_CENTER)
    Call ReleaseFontIndirect(objPrint)
    '病人基本信息输出
    lngCurX = lngX: lngCurY = T_LableRect.Bottom + objPrint.ScaleY(objPrint.TextHeight("1"), vbTwips, vbPixels) / 2
    strTmp = "姓名:'年龄:'性别:'科别:'床号:'入院日期:'住院病历号:'诊断:"
    arrItem = Split(strTmp, "'")
    Call SetFontIndirect(objPrint, "宋体,9,粗")
    T_Size.W = objPrint.ScaleX(objPrint.TextWidth(strTmp), vbTwips, vbPixels)
    T_Size.H = objPrint.ScaleY(objPrint.TextHeight(strTmp), vbTwips, vbPixels)
    lngY = lngCurY + T_Size.H
    lngWidth = (objPrint.ScaleX(objPrint.Width, vbTwips, vbPixels) - lngX * 2 - T_Size.W) / (UBound(arrItem) + 1)
    If lngWidth < 0 Then lngWidth = 0
    For lngIndex = 0 To UBound(arrItem)
        Call GetTextRect(objPrint, lngCurX, lngCurY, arrItem(lngIndex), 0, False)
        Call DrawText(mlngDC, arrItem(lngIndex), -1, T_LableRect, DT_CENTER)
        lngCurX = lngCurX + lngWidth + objPrint.ScaleX(objPrint.TextWidth(arrItem(lngIndex)), vbTwips, vbPixels)
    Next lngIndex
    Call ReleaseFontIndirect(objPrint)
    '表格输出
    Call SetFontIndirect(objPrint, "宋体,9")
    lngCurX = lngX: lngCurY = lngY + objPrint.ScaleY(objPrint.TextHeight("1"), vbTwips, vbPixels) / 2
    lngY = lngCurY
    arrItem = Split(TBasicStyle.TabTitleName, "@")
    '先输出内容
    For lngIndex = 0 To UBound(arrItem)
        T_Size.W = objPrint.ScaleX(objPrint.TextWidth(arrItem(lngIndex)), vbTwips, vbPixels)
        T_Size.H = objPrint.ScaleY(objPrint.TextHeight(arrItem(lngIndex)), vbTwips, vbPixels)
        lngHeight = (TWaveDrawStyle.上表格高度 - T_Size.H) / 2
        If lngHeight < 0 Then lngHeight = 0
        Call GetTextRect(objPrint, lngCurX, lngCurY + lngHeight, arrItem(lngIndex), objPrint.ScaleY(TBasicStyle.ScaleColWidth, vbTwips, vbPixels), False)
        Call DrawText(mlngDC, arrItem(lngIndex), -1, T_LableRect, DT_CENTER)
        lngCurY = lngCurY + TWaveDrawStyle.上表格高度
    Next lngIndex
    '再画线 (横线)
    lngCurX = lngX: lngCurY = lngY
    lngWidth = TWaveDrawStyle.刻度总宽度 + TBasicStyle.TabDays * (TBasicStyle.TabDayTime * TWaveDrawStyle.曲线列宽度) + lngX
    For lngIndex = 0 To UBound(arrItem) + 1
        Call WaveDrawLine(mlngDC, lngCurX, lngCurY, lngWidth, lngCurY, PS_SOLID, IIf(lngIndex = 0 Or lngIndex = UBound(arrItem) + 1, intBold, intFine), RGB_BLACK)
        lngCurY = lngCurY + TWaveDrawStyle.上表格高度
    Next lngIndex
    '(竖线)
    lngCurX = lngX: lngCurY = lngY
    lngHeight = lngCurY + TWaveDrawStyle.上表格高度 * (UBound(arrItem) + 1)
    lngY = lngHeight
    Call WaveDrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngHeight, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = lngCurX + TWaveDrawStyle.刻度总宽度
    For lngIndex = 0 To TBasicStyle.TabDays
        Call WaveDrawLine(mlngDC, lngCurX, lngCurY, lngCurX, lngHeight, PS_SOLID, intBold, RGB_BLACK)
        lngCurX = lngCurX + TBasicStyle.TabDayTime * TWaveDrawStyle.曲线列宽度
    Next lngIndex
    T_Size.H = objPrint.ScaleY(objPrint.TextHeight("1"), vbTwips, vbPixels)
    lngCurX = lngX + TWaveDrawStyle.刻度总宽度
    If T_Size.H > TWaveDrawStyle.上表格高度 Then
        lngCurY = lngCurY + TWaveDrawStyle.上表格高度 * UBound(arrItem)
    Else
        lngCurY = lngY - T_Size.H
    End If
    '在画出时间
    For lngIndex = 1 To TBasicStyle.TabDays
        For lngRow = 1 To TBasicStyle.TabDayTime
            strTmp = TBasicStyle.TabBeginTime + (lngRow - 1) * TBasicStyle.TabTimeSplit
            Call GetTextRect(objPrint, lngCurX, lngCurY, strTmp, TWaveDrawStyle.曲线列宽度, False)
            Call DrawText(mlngDC, strTmp, -1, T_LableRect, DT_CENTER)
            lngCurX = lngCurX + TWaveDrawStyle.曲线列宽度
            '画线
            If Not lngRow = TBasicStyle.TabDayTime Then
                Call WaveDrawLine(mlngDC, lngCurX, lngY - TWaveDrawStyle.上表格高度, lngCurX, lngY, PS_SOLID, intFine, RGB_BLACK)
            End If
        Next lngRow
    Next lngIndex
    Call ReleaseFontIndirect(objPrint)
    
    lngCurX = lngX: lngCurY = lngY
    '--THREE:生命体征栏的绘制
    Call DrawCanvas(mlngDC, objPrint, rsCurve, lngCurX, lngCurY)
    lngY = lngY + lngCurveRows * TWaveDrawStyle.曲线列高度
    Call SetTextColor(mlngDC, RGB_BLACK)
    '--FOUR:特殊项目栏的绘制
    lngCurX = lngX: lngCurY = lngY
    mlngHeight = lngY
    If TBasicStyle.BlnBaby = False Then
        Call DrawDownTab(mlngDC, objPrint, rsCurve, lngCurX, lngCurY)
    End If
    
    '--Five:标准体温单输出示例样稿
    If blnExamples = True Then
        Call SetTextColor(mlngDC, RGB_RED)
        Call SetFontIndirect(objPrint, "宋体,30")
        strTmp = "示例样稿"
        Call GetTextRect(objPrint, 0, objPrint.ScaleY(mlngWaveHeight, vbTwips, vbPixels) \ 3, strTmp, objPrint.ScaleX(mlngWaveWidth, vbTwips, vbPixels) + 10, True)
        Call DrawText(mlngDC, strTmp, -1, T_LableRect, DT_CENTER)
        Call ReleaseFontIndirect(objPrint)
        Call WaveDrawLine(mlngDC, T_LableRect.Left, T_LableRect.Top, T_LableRect.Right, T_LableRect.Top, PS_SOLID, 2, RGB_RED)
        Call WaveDrawLine(mlngDC, T_LableRect.Left, T_LableRect.Top, T_LableRect.Left, T_LableRect.Bottom, PS_SOLID, 2, RGB_RED)
        Call WaveDrawLine(mlngDC, T_LableRect.Left, T_LableRect.Bottom, T_LableRect.Right, T_LableRect.Bottom, PS_SOLID, 2, RGB_RED)
        Call WaveDrawLine(mlngDC, T_LableRect.Right, T_LableRect.Top, T_LableRect.Right, T_LableRect.Bottom, PS_SOLID, 2, RGB_RED)
        Call SetTextColor(mlngDC, RGB_BLACK)
    End If
    
    DrawWaveStyle = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function DrawCanvas(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, ByVal lngX As Long, ByVal lngY As Long) As String
'------------------------------------------------------------------------------------------------------
'功能:画刻度区域和体温区域并输出刻度值信息
'参数:lngDC 绘图对象的DC，objDraw 绘画对象.rsTemp:体温曲线项目记录集(A.项目序号,A.排列序号,A.记录名,A.记录符,A.记录色,A.最大值,A.最小值,A.单位值,C.项目单位 单位,A.最高行-2 AS 最高行,B.部位)
'出参:返回各个曲线的具体信息包括( "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|颜色")
'返回说明信息(项目的符号)
'-------------------------------------------------------------------------------------------------------
    Dim str说明 As String
    Dim lngMaxX     As Long, lngMaxY As Long  '边界
    Dim lngCurX As Long, lngCurY As Long
    Dim sinCurAlerY As Single '警戒线
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim lngCurveRows As Long '曲线的行数
    Dim bln双行 As Boolean                  '此参数由用户指定,bln双行=TRUE表示只显示五行;否则显示十行
    '以下都是标准尺度
    Dim intLineMode   As Integer
    Dim sinAlertness  As Single              '警戒线,起辅助作用
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim lngRowStep As Long
    Dim arrTemp()     As String
    Dim intBold As Integer, intFine As Integer
    Dim sinY单位 As Single '曲线单位输出的Bottom

    '以下与绘图区域相关(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    Dim sin刻度 As Single, bln显示刻度 As Boolean, blnFirst As Boolean
    Dim sin刻度间隔 As Single, sinBegin刻度 As Single, dbl单位值 As Double

    On Error GoTo errHand
    If TypeName(objDraw) = "Printer" Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    '------------------------------------------------------------------------------------------------------------------
    '赋初值
    str说明 = ""
    intLineMode = PS_SOLID
    bln双行 = True
    
    lngColStep = TWaveDrawStyle.曲线列宽度
    lngRowStep = TWaveDrawStyle.曲线列高度
    '第一步：先完成记录法=1(曲线)的输出
    rsTemp.Filter = "记录法=1"
    intLables = rsTemp.RecordCount
    lngLableStep = TWaveDrawStyle.刻度总宽度 \ intLables
    
    lngCurX = lngX: lngCurY = lngY
    lngMaxX = lngCurX + TWaveDrawStyle.刻度总宽度 + TBasicStyle.TabDays * TBasicStyle.TabDayTime * TWaveDrawStyle.曲线列宽度
    lngMaxY = TWaveDrawStyle.曲线总行数 * lngRowStep + lngCurY
    '先画刻度区域
    For lngRow = 1 To intLables
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow = 1, intBold, intFine), RGB_BLACK)
        lngCurX = lngCurX + lngLableStep
        Call WaveDrawLine(lngDC, lngCurX - lngLableStep, lngMaxY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        If lngRow = intLables Then
            lngCurX = TWaveDrawStyle.刻度总宽度 + lngX
            Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
        End If
    Next
    '默认添加一行用于显示项目名称
    lngCurY = lngCurY + lngRowStep * mintCurveNullRow
    lngCurX = lngX
    Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, PS_SOLID, intFine, RGB_BLACK)
    '画体温单所有行
    lngCurX = lngX + TWaveDrawStyle.刻度总宽度
    lngCurveRows = TWaveDrawStyle.曲线总行数 - mintCurveNullRow
    For lngRow = 1 To lngCurveRows
        lngCurY = lngCurY + lngRowStep
        '画体温单的所有行
        If (bln双行 And lngRow Mod 2 = 0) Or Not bln双行 Or lngRow = lngCurveRows Then
            Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0 Or lngRow = lngCurveRows, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 Or lngRow = lngCurveRows, intBold, intFine), RGB_BLACK)
        End If
    Next
    lngCurY = lngY
    '画体温单所有列
    For lngRow = 1 To TBasicStyle.TabDays * TBasicStyle.TabDayTime
        lngCurX = lngCurX + lngColStep
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod TBasicStyle.TabDayTime = 0, intBold, intFine), IIf(lngRow Mod TBasicStyle.TabDayTime = 0, RGB_RED, RGB_BLACK))
    Next
    '画刻度框的标尺（从固定不变的10行开始标识）
    intLables = 1
    rsTemp.Filter = "记录法=1"
    rsTemp.Sort = "排列序号"
    lngCurX = lngX
    With rsTemp
        Do While Not .EOF
            '显示刻度框项目的名称及符号,如体温×
            If intLables = rsTemp.RecordCount Then
                lngLableStep = TWaveDrawStyle.刻度总宽度 - ((intLables - 1) * lngLableStep)
            End If
            lngCurX = lngX + ((intLables - 1) * lngLableStep)
            lngCurY = lngY
            '输出体温项目的名称
            Call SetFontIndirect(objDraw, "宋体,9")
            Call SetTextColor(lngDC, NVL(!记录色, RGB_BLACK))
            Call GetTextRect(objDraw, lngCurX, lngCurY + objDraw.ScaleY(objDraw.TextHeight(NVL(!记录名)), vbTwips, vbPixels) \ 2, Trim(NVL(!记录名)), lngLableStep)
            Call DrawText(lngDC, Trim(NVL(!记录名)), -1, T_LableRect, DT_CENTER)
            Call ReleaseFontIndirect(objDraw)
            '输出项目单位
            If Trim(NVL(!单位)) <> "" Then
                Call SetFontIndirect(objDraw, "宋体,8")
                Call GetTextRect(objDraw, lngCurX, lngCurY + lngRowStep * mintCurveNullRow + objDraw.ScaleY(objDraw.TextHeight(NVL(!单位)), vbTwips, vbPixels) \ 2, Trim(NVL(!单位)), lngLableStep)
                Call DrawText(lngDC, Trim(NVL(!单位, 0)), -1, T_LableRect, DT_CENTER)
                Call ReleaseFontIndirect(objDraw)
                sinY单位 = T_LableRect.Bottom
            Else
                sinY单位 = lngY + lngRowStep * mintCurveNullRow
            End If
            
            intLables = intLables + 1
            Call SetFontIndirect(objDraw, "宋体,9")
            '强制设定体温曲线项目的显示模式
            Select Case !项目序号
                Case 1  '体温整数时输出刻度
                    sin刻度间隔 = NVL(!刻度间隔, 1)
                    dbl单位值 = 0.1
                    sinAlertness = NVL(!警示线, 37)
                    arrTemp = Split(NVL(!记录符, "·,×,○"), ",")
                    str说明 = str说明 & "、" & NVL(!记录名) & "(口温" & arrTemp(0) & ",腋温" & arrTemp(1) & ",肛温" & arrTemp(2) & ")"
                Case 2, -1  '脉搏/心跳按10的倍数输出刻度
                    sin刻度间隔 = NVL(!刻度间隔, 10)
                    dbl单位值 = 2
                    sinAlertness = NVL(!警示线, 0)
                    If !项目序号 = 2 Then
                        str说明 = str说明 & "、" & NVL(!记录名) & "(缺省记录符" & NVL(!记录符, "+") & ",起搏器H)"
                    Else
                        str说明 = str说明 & "、" & NVL(!记录名) & "(" & NVL(!记录符, "Ο") & ")"
                    End If
                Case 3  '呼吸按5的倍数输出刻度
                    dbl单位值 = 1
                    sin刻度间隔 = NVL(!刻度间隔, 5)
                    sinAlertness = NVL(!警示线, 0)
                    str说明 = str说明 & "、" & NVL(!记录名) & "(自主呼吸" & NVL(!记录符, "*") & ",呼吸机R)"
                Case Else
                    dbl单位值 = Val(NVL(!单位值, 0))
                    sin刻度间隔 = NVL(!刻度间隔, Val(NVL(!单位值, 0)) * 10)
                    If sin刻度间隔 > Val(NVL(!最大值)) - Val(NVL(!最小值)) Then
                        sin刻度间隔 = Val(NVL(!最大值)) - Val(NVL(!最小值))
                    End If
                    sinAlertness = NVL(!警示线, 0)
                    str说明 = str说明 & "、" & NVL(!记录名) & "(" & NVL(!记录符, "*") & ")"
            End Select
            '赋初值
            lngCurY = lngCurY + lngRowStep * mintCurveNullRow '固定前2行的高度不输出刻度
            '根据最高行定位到有效位置
            lngCurY = lngCurY + (TWaveDrawStyle.曲线列高度 * Val(NVL(!最高行, 0)))
            blnFirst = False
            Do While True
                bln显示刻度 = False
                If blnFirst = False Then    '刚进入循环，此时取的最大值
                    sin刻度 = NVL(!最大值, 0)
                    sinBegin刻度 = sin刻度
                    blnFirst = True
                Else                    '计算得到每个刻度的值
                    sin刻度 = sin刻度 - dbl单位值     '如果目前显示模式为双倍，则按双倍累计
                End If

                '根据设置的刻度间隔显示刻度值
                If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
                If bln显示刻度 = True Or sin刻度 < sinBegin刻度 Then sinBegin刻度 = sinBegin刻度 - sin刻度间隔
                If sinBegin刻度 < Val(NVL(!最小值)) Then sinBegin刻度 = Val(NVL(!最小值))

                If bln显示刻度 Then
                    '控制最大值不与曲线单位重复
                    If sin刻度 = Val(NVL(!最大值, 0)) And lngCurY <= sinY单位 Then
                        Call GetTextRect(objDraw, lngCurX, sinY单位 + IIf(lngCurY = sinY单位, (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), 0), Val(Format(sin刻度, "#0.0")), lngLableStep)
                    ElseIf Format(lngCurY, "#0") = lngMaxY Then
                        Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin刻度, "#0.0")), lngLableStep)
                    Else
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Val(Format(sin刻度, "#0.0")), lngLableStep)
                    End If
                    Call DrawText(lngDC, Val(Format(sin刻度, "#0.0")), -1, T_LableRect, DT_CENTER)
                End If
                If Val(Format(sin刻度, "#0.00")) <= Val(Format(NVL(!最小值), "#0.00")) Or Format(lngCurY, "#0") > lngMaxY Then
                    '输出警戒线
                    If sinAlertness > Val(NVL(!最小值)) And sinAlertness < Val(NVL(!最大值)) Then
                        '根据最大值与当前值之间的差额,以及最小值,计算得到相差多少个刻度,再根据单位刻度得到实际坐标
                        sinCurAlerY = Format((Val(NVL(!最大值)) - sinAlertness) / dbl单位值 * TWaveDrawStyle.曲线列高度, "#0.0")
                        sinCurAlerY = Format(sinCurAlerY + lngY + lngRowStep * mintCurveNullRow + (TWaveDrawStyle.曲线列高度 * Val(NVL(!最高行, 0))), "#0")
                        Call WaveDrawLine(lngDC, lngX + TWaveDrawStyle.刻度总宽度, CLng(sinCurAlerY), lngMaxX, CLng(sinCurAlerY), PS_SOLID, intBold, RGB_RED)
                    End If
         
                    Exit Do
                End If
                lngCurY = lngCurY + TWaveDrawStyle.曲线列高度
            Loop
            Call ReleaseFontIndirect(objDraw)
            sinBegin刻度 = 0
            sin刻度 = 0
            .MoveNext
        Loop
    End With
    '完成独立曲线部分的输出
    rsTemp.Filter = "记录法=3"
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            lngY = lngMaxY
            lngCurY = lngY
            lngCurX = lngX
            lngCurveRows = ((Val(NVL(!最大值, 0)) - Val(NVL(!最小值, 0))) / Val(NVL(!单位值)))
            If Val(NVL(!最高行, 0)) > 0 Then lngCurveRows = lngCurveRows + Val(NVL(!最高行, 0))
            If lngCurveRows Mod 2 = 1 Then lngCurveRows = lngCurveRows + 1
            If lngCurveRows > 0 Then
                lngMaxY = lngCurveRows * lngRowStep + lngCurY
                '完成刻度区域的绘制
                Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call WaveDrawLine(lngDC, lngCurX + TWaveDrawStyle.刻度总宽度, lngCurY, lngCurX + TWaveDrawStyle.刻度总宽度, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                Call WaveDrawLine(lngDC, lngCurX, lngMaxY, lngCurX + TWaveDrawStyle.刻度总宽度, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
                '完成所有行的绘制
                lngCurX = lngX + TWaveDrawStyle.刻度总宽度
                For lngRow = 1 To lngCurveRows
                    lngCurY = lngCurY + lngRowStep
                    '画体温单的所有行
                    If (bln双行 And lngRow Mod 2 = 0) Or Not bln双行 Or lngRow = lngCurveRows Then
                        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0 Or lngRow = lngCurveRows, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 Or lngRow = lngCurveRows, intBold, intFine), RGB_BLACK)
                    End If
                Next
                lngCurY = lngY
                '完成所有列的绘制
                For lngRow = 1 To TBasicStyle.TabDays * TBasicStyle.TabDayTime
                    lngCurX = lngCurX + lngColStep
                    Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod TBasicStyle.TabDayTime = 0, intBold, intFine), IIf(lngRow Mod TBasicStyle.TabDayTime = 0, RGB_RED, RGB_BLACK))
                Next
                '完成项目名称和刻度的输出
                lngCurX = lngX: lngCurY = lngY
                '输出体温项目的名称
                Call SetFontIndirect(objDraw, "宋体,9")
                Call SetTextColor(lngDC, NVL(!记录色, RGB_BLACK))
                T_Size.H = objDraw.ScaleY(objDraw.TextHeight("刘"), vbTwips, vbPixels)
                If T_Size.H * Len(NVL(!记录名)) >= lngCurveRows * lngRowStep Then
                    lngCurY = lngY
                Else
                    lngCurY = lngY + ((lngCurveRows * lngRowStep) - (T_Size.H * Len(NVL(!记录名)))) \ 2
                End If
                For lngRow = 1 To Len(NVL(!记录名))
                    Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(NVL(!记录名), lngRow, 1), TWaveDrawStyle.刻度总宽度 \ 2, False)
                    Call DrawText(lngDC, Mid(NVL(!记录名), lngRow, 1), -1, T_LableRect, DT_CENTER)
                    lngCurY = lngCurY + T_Size.H
                Next lngRow
                Call ReleaseFontIndirect(objDraw)
                '输出项目单位
                lngCurY = lngY: If NVL(!记录名) <> "" Then lngCurX = T_LableRect.Right
                If Trim(NVL(!单位)) <> "" And NVL(!记录名) <> "" Then
                    Call SetFontIndirect(objDraw, "宋体,8")
                    T_Size.H = objDraw.ScaleY(objDraw.TextHeight("刘"), vbTwips, vbPixels)
                    If T_Size.H * Len(Trim(NVL(!单位))) >= lngCurveRows * lngRowStep Then
                        lngCurY = lngY
                    Else
                        lngCurY = lngY + ((lngCurveRows * lngRowStep) - (T_Size.H * Len(NVL(!单位)))) \ 2
                    End If
                    For lngRow = 1 To Len(Trim(NVL(!单位)))
                        Call GetTextRect(objDraw, lngCurX, lngCurY, Mid(Trim(NVL(!单位)), lngRow, 1), 0, False)
                        Call DrawText(lngDC, Mid(Trim(NVL(!单位)), lngRow, 1), -1, T_LableRect, DT_CENTER)
                        lngCurY = lngCurY + T_Size.H
                    Next lngRow
                    Call ReleaseFontIndirect(objDraw)
                End If
                Call SetFontIndirect(objDraw, "宋体,9")
                dbl单位值 = Val(NVL(!单位值, 0))
                sin刻度间隔 = NVL(!刻度间隔, Val(NVL(!单位值, 0)) * 10)
                If sin刻度间隔 > Val(NVL(!最大值)) - Val(NVL(!最小值)) Then
                    sin刻度间隔 = Val(NVL(!最大值)) - Val(NVL(!最小值))
                End If
                sinAlertness = NVL(!警示线, 0)
                str说明 = str说明 & "、" & NVL(!记录名) & "(" & NVL(!记录符, "*") & ")"
                lngCurY = lngY + (TWaveDrawStyle.曲线列高度 * Val(NVL(!最高行, 0)))
                blnFirst = False
                Do While True
                    bln显示刻度 = False
                    If blnFirst = False Then     '刚进入循环，此时取的最大值
                        sin刻度 = NVL(!最大值, 0)
                        sinBegin刻度 = sin刻度
                        blnFirst = True
                    Else                    '计算得到每个刻度的值
                        sin刻度 = sin刻度 - dbl单位值     '如果目前显示模式为双倍，则按双倍累计
                    End If
    
                    '根据设置的刻度间隔显示刻度值
                    If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
                    If bln显示刻度 = True Or sin刻度 < sinBegin刻度 Then sinBegin刻度 = sinBegin刻度 - sin刻度间隔
                    If sinBegin刻度 < Val(NVL(!最小值)) Then sinBegin刻度 = Val(NVL(!最小值))
    
                    If bln显示刻度 Then
                        '控制最大值不与曲线单位重复
                        lngCurX = lngX + TWaveDrawStyle.刻度总宽度 - objDraw.ScaleX(objDraw.TextWidth(Val(Format(sin刻度, "#0.0"))), vbTwips, vbPixels)
                        lngCurX = lngCurX - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 3)
                        If sin刻度 = Val(NVL(!最大值, 0)) And lngCurY = lngY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY + (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin刻度, "#0.0")))
                        ElseIf Format(lngCurY, "#0") = lngMaxY Then
                            Call GetTextRect(objDraw, lngCurX, lngCurY - (objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin刻度, "#0.0")))
                        Else
                            Call GetTextRect(objDraw, lngCurX, lngCurY, Val(Format(sin刻度, "#0.0")))
                        End If
                        Call DrawText(lngDC, Val(Format(sin刻度, "#0.0")), -1, T_LableRect, DT_CENTER)
                    End If
                    If Val(Format(sin刻度, "#0.00")) <= Val(Format(NVL(!最小值), "#0.00")) Or Format(lngCurY, "#0") > lngMaxY Then
                        '输出警戒线
                        If sinAlertness > Val(NVL(!最小值)) And sinAlertness < Val(NVL(!最大值)) Then
                            '根据最大值与当前值之间的差额,以及最小值,计算得到相差多少个刻度,再根据单位刻度得到实际坐标
                            sinCurAlerY = Format((Val(NVL(!最大值)) - sinAlertness) / dbl单位值 * TWaveDrawStyle.曲线列高度, "#0.0")
                            sinCurAlerY = Format(sinCurAlerY + lngY + (TWaveDrawStyle.曲线列高度 * Val(NVL(!最高行, 0))), "#0")
                            Call WaveDrawLine(lngDC, lngX + TWaveDrawStyle.刻度总宽度, CLng(sinCurAlerY), lngMaxX, CLng(sinCurAlerY), PS_SOLID, intBold, RGB_RED)
                        End If
                        Exit Do
                    End If
                    lngCurY = lngCurY + TWaveDrawStyle.曲线列高度
                Loop
                Call ReleaseFontIndirect(objDraw)
                sinBegin刻度 = 0
                sin刻度 = 0
            End If
        .MoveNext
        Loop
    End With
    str说明 = "说明:" & Mid(str说明, 2)
    DrawCanvas = str说明
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub DrawDownTab(ByVal lngDC As Long, ByVal objDraw As Object, ByVal rsTemp As ADODB.Recordset, ByVal lngX As Long, ByVal lngY As Long)
'-----------------------------------------------------
'功能: 完成表下表格内容的输出
'-----------------------------------------------------
    Dim lngCurX As Long, lngCurY As Long, lngRowTopY As Long
    Dim lngMaxX As Long, lngMaxY As Long
    Dim lngRow As Long, lngHeight As Long, lngWidth As Long
    Dim lngDayTime As Long, lngRecordTime As Long, lngRowTime As Long
    Dim lngTimeCOlWidth As Long '频次列宽
    Dim strContent As String '表头内容
    Dim intBold As Integer, intFine As Integer '线条属性
    Dim arrItem() As String '绑定的项目属性
    Dim i As Long, j As Long, k As Long
    If TWaveDrawStyle.表下表格总行数 = 0 Then Exit Sub
    
    On Error GoTo errHand
    If TypeName(objDraw) = "Printer" Then
        intBold = 6
        intFine = 2
    Else
        intBold = 2
        intFine = 1
    End If
    
    lngCurX = lngX: lngCurY = lngY
    lngMaxX = lngCurX + TWaveDrawStyle.刻度总宽度 + TBasicStyle.TabDays * TBasicStyle.TabDayTime * TWaveDrawStyle.曲线列宽度
    lngMaxY = lngCurY + TWaveDrawStyle.表下表格总行数 * TWaveDrawStyle.下表格高度 - IIf(mbln呼吸表格 = True, (TWaveDrawStyle.下表格高度 - TWaveDrawStyle.呼吸表格高度), 0)
    
    '先完成表格外边框的绘制
    '画竖线(表头部分)
    Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = lngCurX + TWaveDrawStyle.刻度总宽度
    Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    '画竖线(表体部分)
    For lngRow = 1 To TBasicStyle.TabDays
        lngCurX = lngCurX + TBasicStyle.TabDayTime * TWaveDrawStyle.曲线列宽度
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    Next lngRow
    lngCurX = lngX + TWaveDrawStyle.刻度总宽度
    lngCurY = lngY
    '画横线(表体部分)
    For lngRow = 1 To TWaveDrawStyle.表下表格总行数
        If mbln呼吸表格 = True And lngRow = 1 Then
            lngCurY = lngCurY + TWaveDrawStyle.呼吸表格高度
        Else
            lngCurY = lngCurY + TWaveDrawStyle.下表格高度
        End If
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngMaxX, lngCurY, PS_SOLID, IIf(lngRow = TWaveDrawStyle.表下表格总行数, intBold, intFine), RGB_BLACK)
    Next lngRow
    Call WaveDrawLine(lngDC, lngX, lngMaxY, lngX + TWaveDrawStyle.刻度总宽度, lngMaxY, PS_SOLID, intBold, RGB_BLACK)
    lngCurX = lngX: lngCurY = lngY
    Call SetFontIndirect(objDraw, "宋体,9")
    Call SetTextColor(lngDC, RGB_BLACK)
    '首先完成呼吸表格的输出(如果呼吸为表格项目)
    If mbln呼吸表格 = True Then
        rsTemp.Filter = "项目序号=3"
        lngHeight = TWaveDrawStyle.呼吸表格高度
        lngWidth = TWaveDrawStyle.刻度总宽度
        strContent = NVL(rsTemp!记录名) & IIf(Trim(NVL(rsTemp!单位)) <> "", "(" & Trim(NVL(rsTemp!单位)) & ")", "")
        T_Size.H = objDraw.ScaleY(objDraw.TextHeight(strContent), vbTwips, vbPixels)
        T_Size.W = objDraw.ScaleX(objDraw.TextWidth(strContent), vbTwips, vbPixels)
        If T_Size.H < lngHeight Then
            lngCurY = lngCurY + (lngHeight - T_Size.H) \ 2
        End If
        Call GetTextRect(objDraw, lngCurX, lngCurY, strContent, lngWidth, False)
        '从新设置区域
        If T_LableRect.Left < lngX Then T_LableRect.Left = lngX
        If T_LableRect.Top < lngY Then T_LableRect.Top = lngY
        If T_LableRect.Right > lngWidth + lngX Then T_LableRect.Right = lngWidth + lngX
        If T_LableRect.Bottom > lngHeight + lngY Then T_LableRect.Bottom = lngHeight + lngY
        Call DrawText(lngDC, strContent, -1, T_LableRect, DT_CENTER)
        If TWaveDrawStyle.表下表格总行数 > 1 Then
            lngCurY = lngY + lngHeight
            Call WaveDrawLine(lngDC, lngX, lngCurY, lngX + TWaveDrawStyle.刻度总宽度, lngCurY, PS_SOLID, intFine, RGB_BLACK)
        End If
        '完成频次的绘制
        lngCurY = lngY
        For lngRow = 1 To TBasicStyle.TabDays
            lngCurX = lngX + TWaveDrawStyle.刻度总宽度 + (lngRow - 1) * TBasicStyle.TabDayTime * TWaveDrawStyle.曲线列宽度
            For lngDayTime = 1 To TBasicStyle.TabDayTime - 1
                lngCurX = lngCurX + TWaveDrawStyle.曲线列宽度
                Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngCurY + lngHeight, PS_SOLID, intFine, RGB_BLACK)
            Next lngDayTime
        Next lngRow
    End If
    '其他表格输出
    arrItem = Split(mstrTabItem, "|")
    lngCurY = lngY + IIf(mbln呼吸表格 = True, TWaveDrawStyle.呼吸表格高度, 0)
    lngRowTopY = lngCurY
    lngCurX = lngX
    rsTemp.Filter = "记录法=2 And 项目序号<>3 And 项目序号<>5"
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            For lngRow = 0 To UBound(arrItem)
                If Val(Split(arrItem(lngRow), ";")(0)) = Val(!项目序号) Then
                    lngRecordTime = Val(Split(arrItem(lngRow), ";")(1))
                    Select Case lngRecordTime '记录频次
                    Case 3
                        lngRowTime = 1 '每一行表格列数
                        lngDayTime = 3 '占用的表格行数
                    Case 4
                        lngRowTime = 2 '每一行表格列数
                        lngDayTime = 2 '占用的表格行数
                    Case Else
                        lngRowTime = lngRecordTime '每一行表格列数
                        lngDayTime = 1 '占用的表格行数
                    End Select
                    '输出表头名称
                    lngHeight = TWaveDrawStyle.下表格高度 * lngDayTime
                    lngWidth = TWaveDrawStyle.刻度总宽度
                    strContent = NVL(rsTemp!记录名) & IIf(Trim(NVL(rsTemp!单位)) <> "", "(" & Trim(NVL(rsTemp!单位)) & ")", "")
                    T_Size.H = objDraw.ScaleY(objDraw.TextHeight(strContent), vbTwips, vbPixels)
                    T_Size.W = objDraw.ScaleX(objDraw.TextWidth(strContent), vbTwips, vbPixels)
                    If T_Size.H < lngHeight Then
                        lngCurY = lngCurY + (lngHeight - T_Size.H) \ 2
                    End If
                    Call GetTextRect(objDraw, lngCurX, lngCurY, strContent, lngWidth, False)
                    '重新设置区域
                    If T_LableRect.Left < lngX Then T_LableRect.Left = lngX
                    If T_LableRect.Top < lngRowTopY Then T_LableRect.Top = lngRowTopY
                    If T_LableRect.Right > lngWidth + lngX Then T_LableRect.Right = lngWidth + lngX
                    If T_LableRect.Bottom > lngHeight + lngRowTopY Then T_LableRect.Bottom = lngHeight + lngRowTopY
                    Call DrawText(lngDC, strContent, -1, T_LableRect, DT_CENTER)
                    If Not rsTemp.RecordCount = rsTemp.AbsolutePosition Then
                        lngCurY = lngRowTopY + lngHeight
                        Call WaveDrawLine(lngDC, lngX, lngCurY, lngX + TWaveDrawStyle.刻度总宽度, lngCurY, PS_SOLID, intFine, RGB_BLACK)
                    End If
                    lngCurY = lngRowTopY
                    '计算每一行频次占用的列宽
                    lngTimeCOlWidth = (TBasicStyle.TabDayTime * TWaveDrawStyle.曲线列宽度) \ lngRowTime
                    '绘制表格列
                    For i = 1 To lngDayTime '占用的表格行数
                        lngCurX = lngX + TWaveDrawStyle.刻度总宽度
                        lngCurY = lngRowTopY + (i - 1) * TWaveDrawStyle.下表格高度
                        For j = 1 To TBasicStyle.TabDays '监测的天数
                            lngCurX = lngX + TWaveDrawStyle.刻度总宽度 + (j - 1) * TBasicStyle.TabDayTime * TWaveDrawStyle.曲线列宽度
                            For k = 1 To lngRowTime - 1 '每一行表格列数
                                lngCurX = lngCurX + lngTimeCOlWidth
                                Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngCurX, lngCurY + TWaveDrawStyle.下表格高度, PS_SOLID, intFine, RGB_BLACK)
                            Next k
                        Next j
                    Next i
                    lngCurX = lngX
                    lngRowTopY = lngRowTopY + lngHeight
                    lngCurY = lngRowTopY
                    Exit For
                End If
            Next lngRow
        .MoveNext
        Loop
    End With
    
    '最后完成空行的绘制
    lngRecordTime = TWaveDrawStyle.表下表格总行数 - TBasicStyle.AddTabNull '空行的起始行
    lngCurX = lngX
    lngCurY = lngY + IIf(mbln呼吸表格 = True, TWaveDrawStyle.呼吸表格高度, 0) + (lngRecordTime - IIf(mbln呼吸表格 = True, 1, 0)) * TWaveDrawStyle.下表格高度
    For lngRow = lngRecordTime To TWaveDrawStyle.表下表格总行数 - 1
        Call WaveDrawLine(lngDC, lngCurX, lngCurY, lngX + TWaveDrawStyle.刻度总宽度, lngCurY, PS_SOLID, intFine, RGB_BLACK)
        lngCurY = lngCurY + TWaveDrawStyle.下表格高度
    Next lngRow
    Call ReleaseFontIndirect(objDraw)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'-------------------------------------------------------------------------------------------
Private Sub SetFontIndirect(ByVal objDraw As Object, ByVal strFontInfo As String)
    '功能:设置字体属性,并使用
    Dim BFileName() As Byte
    Dim i As Integer
    Dim stdset As StdFont

    On Error GoTo errHand
    Set stdset = New StdFont
    With stdset
        .Name = Split(strFontInfo, ",")(0)
        .Size = 9
        If UBound(Split(strFontInfo, ",")) > 0 Then .Size = Val(Split(strFontInfo, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strFontInfo, "粗") > 0 Then .Bold = True
        If InStr(1, strFontInfo, "斜") > 0 Then .Italic = True
    End With

    Set objDraw.Font = stdset
    BFileName = StrConv(stdset.Name, vbFromUnicode)
    With T_Font
        For i = 1 To Len(stdset.Name)
            .lfFaceName(i - 1) = BFileName(i - 1)
        Next i
        .lfHeight = -MulDiv(stdset.Size, GetDeviceCaps(mlngDC, LOGPIXELSY), 71)
        .lfWidth = 0
        .lfWeight = IIf(stdset.Bold = True, FW_BOLD, FW_NORMAL)
        .lfCharSet = stdset.Charset
        .lfUnderline = stdset.Underline
        .lfItalic = stdset.Italic
        .lfStrikeOut = stdset.Strikethrough
    End With

    mlngFont = CreateFontIndirect(T_Font)
    mlngOldFont = SelectObject(mlngDC, mlngFont)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ReleaseFontIndirect(ByVal objDraw As Object)
    Dim stdset As StdFont
    Dim strFontInfo As String
    
    strFontInfo = "宋体,9"
    
    Set stdset = New StdFont
    With stdset
        .Name = Split(strFontInfo, ",")(0)
        .Size = 9
        If UBound(Split(strFontInfo, ",")) > 0 Then .Size = Val(Split(strFontInfo, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strFontInfo, "粗") > 0 Then .Bold = True
        If InStr(1, strFontInfo, "斜") > 0 Then .Italic = True
    End With
    Set objDraw.Font = stdset
    Call SelectObject(mlngDC, mlngOldFont)
    Call DeleteObject(mlngFont)
End Sub

Private Sub GetTextRect(ByVal objDraw As Object, ByVal lngX As Long, ByVal lngY As Long, ByVal strInput As String, _
    Optional ByVal lngWidth As Long = 0, Optional bln居中 As Boolean = True, Optional ByVal lngHeght As Long = 0)
    
    Dim lngInputW As Long, lng1H As Long
    Dim sngSize As Single
        
    T_LableRect.Left = lngX + objDraw.ScaleX(15, vbTwips, vbPixels) '避免与左边界划线重合
    
    If bln居中 = True Then
        T_LableRect.Top = lngY - objDraw.ScaleY(objDraw.TextHeight("1") \ 2, vbTwips, vbPixels)
    Else
        T_LableRect.Top = lngY
    End If
    
    T_LableRect.Right = objDraw.ScaleX(objDraw.TextWidth(strInput) + 30, vbTwips, vbPixels) + T_LableRect.Left
    T_LableRect.Bottom = objDraw.ScaleY(objDraw.TextHeight("1") + 30, vbTwips, vbPixels) + T_LableRect.Top
    
    If lngWidth <> 0 And (lngWidth - objDraw.ScaleX(objDraw.TextWidth(strInput) + 30, vbTwips, vbPixels)) \ 2 > 0 Then
        '将文本显示在所示宽度的中间区域
        T_LableRect.Left = T_LableRect.Left + (lngWidth - objDraw.ScaleX(objDraw.TextWidth(strInput) + 30, vbTwips, vbPixels)) \ 2
        T_LableRect.Right = objDraw.ScaleX(objDraw.TextWidth(strInput) + 30, vbTwips, vbPixels) + T_LableRect.Left
    End If
    
    If lngHeght <> 0 And (lngHeght - objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels)) > 0 Then
        T_LableRect.Bottom = T_LableRect.Bottom + (lngHeght - objDraw.ScaleY(objDraw.TextHeight("1"), vbTwips, vbPixels))
    End If
    
End Sub

Public Sub WaveDrawLine(ByVal lngDC As Long, ByVal lngSX As Long, ByVal lngSY As Long, ByVal lngDX As Long, ByVal lngDY As Long, _
    Optional ByVal lngType As Long = PS_SOLID, Optional ByVal intWidth As Integer = 1, Optional ByVal lngRGB As Long = 0, _
    Optional ByVal blnEndRow As Boolean = False)
    
    Dim X As Long
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim sngX As Single, sngY As Single
    On Error GoTo errHand
    '创建新画笔进行划线
    lngPen = CreatePen(lngType, intWidth, lngRGB)
    lngOldPen = SelectObject(lngDC, lngPen)
    '绘图
    Call MoveToEx(lngDC, lngSX, lngSY, T_OldPoint)
    Call LineTo(lngDC, lngDX, lngDY)
    '对于物理降温话上下箭头
    If blnEndRow Then
        If lngSY > lngDY Then '物理降温失败（向上箭头）
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDC, X, lngSY - sngY, T_OldPoint)
                Call LineTo(lngDC, lngSX, lngSY)
            Next X
        Else '物理降温成功(向下箭头)
            For X = lngSX - sngX To lngSX + sngX
                Call MoveToEx(lngDC, X, lngSY + sngY, T_OldPoint)
                Call LineTo(lngDC, lngSX, lngSY)
            Next X
        End If
    End If
    
    '还原画笔并销毁
    Call SelectObject(lngDC, lngOldPen)
    Call DeleteObject(lngPen)
    lngPen = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetTabFrequency(ByVal intFrequency As Integer, ByVal intItemStyle As Integer) As String
'------------------------------------------------------
'功能:根据监测天数确定记录频次
'intFrequency：监测次数
'intItemStyle：项目表示:活动和波动项目记录频次最大为2
'------------------------------------------------------
    Dim strFrequency As String
    Dim arrFrequency() As String
    Dim i As Integer
    
    If intItemStyle = 4 Then
        arrFrequency = Split("1|2", "|")
    Else
        arrFrequency = Split("1|2|3|4|6", "|")
    End If
    For i = 0 To UBound(arrFrequency)
        If intFrequency >= Val(arrFrequency(i)) And Not (intFrequency = 8 And Val(arrFrequency(i)) = 6) Then
            strFrequency = strFrequency & "|" & arrFrequency(i)
        End If
    Next i
    
    If Left(strFrequency, 1) = "|" Then strFrequency = Mid(strFrequency, 2)
    If strFrequency = "" Then strFrequency = "1"
    GetTabFrequency = strFrequency
End Function

Public Function GetPipWaveStyle(ByVal lngFileID As Long) As ADODB.Recordset
'-----------------------------------------------------------------------
'功能:获取标准体温单的样式(自行组装)
'-----------------------------------------------------------------------
    Dim strSQL As String
    Dim rsSource As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngParentId As Long, lngId As Long
    Dim lngRow As Long, lngRowNO As Long
    Dim intFields As Integer
    Dim lngItemNO As Long '项目序号
    On Error GoTo errHand
    strSQL = "SELECT Id, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度," & vbNewLine & _
        "       要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
        " FROM 病历文件结构" & vbNewLine & _
        " WHERE 文件id = 0"
    Call zlDatabase.OpenRecordset(rsSource, strSQL, "病历文件结构")
    '开始复制记录集结构体
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Type = 200 Then       '日期型处理为字符型
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
            End If
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '提取项目信息
    strSQL = _
        " SELECT Decode(b.项目序号, 3, Decode(b.记录法, 2, 1, b.排列序号), b.排列序号) 排列序号, b.项目序号, Decode(b.项目序号, 4, '血压', b.记录名) 项目名称, b.单位," & vbNewLine & _
        "       b.记录法," & vbNewLine & _
        "       Decode(b.记录法," & vbNewLine & _
        "               2," & vbNewLine & _
        "               Decode(b.项目序号," & vbNewLine & _
        "                      3," & vbNewLine & _
        "                      6," & vbNewLine & _
        "                      Decode(Decode(c.项目序号, NULL, a.项目表示, 4)," & vbNewLine & _
        "                             4," & vbNewLine & _
        "                             Decode(Sign(Nvl(b.记录频次, 2) - 2), 1, 2, Nvl(b.记录频次, 2))," & vbNewLine & _
        "                             Nvl(b.记录频次, 2)))," & vbNewLine & _
        "               NULL) 记录频次" & vbNewLine & _
        " FROM 护理记录项目 a, 体温记录项目 b, 护理波动项目 c" & vbNewLine & _
        " WHERE a.项目序号 = b.项目序号 AND a.项目序号 = c.项目序号(+) AND NVL(a.应用方式,0) <> 0 AND a.项目性质 = 1 AND a.项目序号 <> 5" & vbNewLine & _
        " ORDER BY Decode(b.记录法, 2, 2, 1), Decode(b.项目序号, 3, Decode(b.记录法, 2, 1, b.排列序号), b.排列序号)"
    Call zlDatabase.OpenRecordset(rsSource, strSQL, "体温项目")
    '1:体温单的基本样式与属性
    lngParentId = 100
    With rsTemp
        '父对象
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = Null
        .Fields("对象序号").Value = 1: .Fields("对象类型").Value = 1: .Fields("对象属性").Value = "体温单的基本样式与属性"
        .Fields("内容文本").Value = "格式定义"
        .Update
        '子对象
        lngId = 101
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 1: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "标题文本"
        .Fields("内容文本").Value = "体温单": .Fields("要素名称").Value = "标题文本"
        .Update
        lngId = 102
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 2: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "标题字体"
        .Fields("内容文本").Value = "宋体,20,粗": .Fields("要素名称").Value = "标题字体"
        .Update
        lngId = 103
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 3: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "天数"
        .Fields("内容文本").Value = 7: .Fields("要素名称").Value = "天数"
        .Update
        lngId = 104
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 4: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "监测次数"
        .Fields("内容文本").Value = 6: .Fields("要素名称").Value = "监测次数"
        .Update
        lngId = 105
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 5: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "开始时点"
        .Fields("内容文本").Value = 4: .Fields("要素名称").Value = "开始时点"
        .Update
        lngId = 106
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 6: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "时间间隔"
        .Fields("内容文本").Value = 4: .Fields("要素名称").Value = "时间间隔"
        .Update
        lngId = 107
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 7: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "一般项目栏表格高度"
        .Fields("内容文本").Value = 255: .Fields("要素名称").Value = "表格高度"
        .Update
        lngId = 108
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 8: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "一般项目栏列头名称"
        .Fields("内容文本").Value = "日       期@住院天数@手术后天数@时       间"
        .Fields("要素名称").Value = "列头名称"
        .Update
        lngId = 109
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 9: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "刻度区域总宽度(缇)"
        .Fields("内容文本").Value = 1350: .Fields("要素名称").Value = "刻度宽度"
        .Update
        lngId = 110
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 10: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "绘图区域曲线表格列宽(缇)"
        .Fields("内容文本").Value = 225: .Fields("要素名称").Value = "曲线列宽"
        .Update
        lngId = 111
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 11: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "绘图区域曲线表格列高(缇)"
        .Fields("内容文本").Value = 90: .Fields("要素名称").Value = "曲线行高"
        .Update
        lngId = 112
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 12: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "曲线表格添加空行数(不针对独立曲线)"
        .Fields("内容文本").Value = 10: .Fields("要素名称").Value = "曲线空行"
        .Update
        lngId = 113
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 13: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "特殊项目栏表格高度"
        .Fields("内容文本").Value = 255: .Fields("要素名称").Value = "表格高度1"
        .Update
        lngId = 114
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 14: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "特殊项目栏表格添加的空行数"
        .Fields("内容文本").Value = 0: .Fields("要素名称").Value = "表格空行"
        .Update
    End With
    '2:体温单曲线项目定义
    lngParentId = 200
    With rsTemp
        '父对象
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = Null
        .Fields("对象序号").Value = 2: .Fields("对象类型").Value = 1: .Fields("对象属性").Value = "体温单曲线项目定义"
        .Fields("内容文本").Value = "曲线项目定义"
        .Update
        lngRowNO = 1
        rsSource.Filter = "记录法=1 OR 记录法=3"
        rsSource.Sort = "记录法,排列序号"
        Do While Not rsSource.EOF
            '子对象
            If Val(NVL(rsSource!项目序号, 0)) <> 0 Then
                lngId = 200 + lngRowNO
                .AddNew
                .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
                .Fields("对象序号").Value = lngRowNO: .Fields("对象类型").Value = 4
                .Fields("对象属性").Value = NVL(rsSource!项目名称)
                .Fields("内容文本").Value = NVL(rsSource!项目序号)
                .Fields("要素名称").Value = NVL(rsSource!项目名称)
                .Update
                lngRowNO = lngRowNO + 1
            End If
        rsSource.MoveNext
        Loop
    End With
    '2:体温单表格项目定义
    lngParentId = 300
    With rsTemp
        '父对象
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = Null
        .Fields("对象序号").Value = 3: .Fields("对象类型").Value = 1: .Fields("对象属性").Value = "体温单表格项目定义"
        .Fields("内容文本").Value = "表格项目定义"
        .Update
        lngRowNO = 1
        rsSource.Filter = "记录法=2"
        rsSource.Sort = "排列序号"
        Do While Not rsSource.EOF
            If Val(NVL(rsSource!项目序号, 0)) <> 0 Then
                '子对象
                lngId = 300 + lngRowNO
                .AddNew
                .Fields("ID").Value = lngId: .Fields("文件ID").Value = lngFileID: .Fields("父ID").Value = lngParentId
                .Fields("对象序号").Value = lngRowNO: .Fields("对象类型") = 4
                .Fields("对象属性").Value = NVL(rsSource!项目名称)
                .Fields("内容文本").Value = NVL(rsSource!项目序号)
                .Fields("要素名称").Value = NVL(rsSource!项目名称)
                .Fields("要素表示").Value = NVL(rsSource!记录频次)
                .Update
                lngRowNO = lngRowNO + 1
            End If
            rsSource.MoveNext
        Loop
    End With
    rsTemp.Filter = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    rsTemp.Sort = "ID"
    
    Set GetPipWaveStyle = rsTemp
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


