VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'######################################################################################
'##模 块 名：cEditor.cls
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：页面视图下各控件共享的属性类。
'##版    本：
'######################################################################################

Option Explicit
'#########################################################################
'   枚举常量
'#########################################################################

'用途: 视图模式
Public Enum ViewModeEnum
    cprNormal = 0   '普通视图
    cprPaper = 2    '页面视图
End Enum

'用途: 超链接键盘鼠标事件
Public Enum LinkEventTypeEnum
   cprLButtonDblClick = WM_LBUTTONDBLCLK
   cprLButtonDown = WM_LBUTTONDOWN
   cprLButtonUp = WM_LBUTTONUP
   cprMouseMove = WM_MOUSEMOVE
   cprRButtonDblClick = WM_RBUTTONDBLCLK
   cprRButtonDown = WM_RBUTTONDOWN
   cprRBUttonUp = WM_RBUTTONUP
   cprSetCursor = WM_SETCURSOR
End Enum

'用途：纸张种类，按名称、高度、宽度、最小边距(上下左右)、对应打印纸张排列的纸张种类常量
Public Enum PaperKindEnum
    cprPK1 = 1               ' "信笺 8 1/2×11 英寸                        ,15842,12242,482,799,181,181,1"
    cprPK2 = 2              ' "+A611 小型信笺 8 1/2×11 英寸              ,15842,12242,482,799,181,181,2"
    cprPK3 = 3              ' "小型报 11×17 英寸                         ,24477,15842,482,799,181,181,3"
    cprPK4 = 4              ' "分类帐 17×11 英寸                         ,15842,24477,482,799,181,181,4"
    cprPK5 = 5              ' "法律文件 8 1/2×14 英寸                    ,20163,12242,482,799,181,181,5"
    cprPK6 = 6              ' "声明书5 1/2×8 1/2 英寸                    ,12242,7919,482,799,181,181,6"
    cprPK7 = 7              ' "行政文件7 1/2×10 1/2 英寸                 ,15122,10438,482,799,181,181,7"
    cprPK8 = 8              ' "A3 297×420 毫米                           ,23814,16840,482,799,181,193,8"
    cprPKA4 = 9             ' "A4 210×297 毫米                           ,16840,11907,482,805,181,176,9"
    cprPK10 = 10            ' "A4小号 210×297 毫米                      ,16840,11907,482,805,181,176,9"
    cprPK11 = 11            ' "A5 148×210 毫米                          ,11907,8392,482,799,181,176,11"
    cprPK12 = 12            ' "B4 250×354 毫米                          ,20639,14572,482,805,181,181,12"
    cprPK13 = 13            ' "B5 182×257 毫米                          ,14572,10319,482,805,181,176,13"
    cprPK14 = 14            ' "对开本 8 1/2×13 英寸                     ,18722,12242,482,799,181,181,14"
    cprPK15 = 15            ' "四开本 215×275 毫米                      ,15589,12187,482,805,181,176,15"
    cprPK16 = 16            ' "10×14 英寸                               ,20163,14398,482,805,181,176,16"
    cprPK17 = 17            ' "11×17 英寸                               ,24477,15842,482,805,181,176,17"
    cprPK18 = 18            ' "便条8 1/2×11 英寸                        ,15842,12242,482,805,181,176,18"
    cprPK19 = 19            ' "#9 信封 3 7/8×8 7/8 英寸                 ,5579,12780,482,794,181,176,19"
    cprPK20 = 20            ' "#10 信封 4 1/8×9 1/2 英寸                ,5936,13682,482,794,181,181,20"
    cprPK21 = 21            ' "#11 信封 4 1/2×10 3/8 英寸               ,14938,6479,482,794,181,181,21"
    cprPK22 = 22            ' "#12 信封 4 1/2×11 英寸                   ,15842,6479,482,794,181,181,22"
    cprPK23 = 23            ' "#14 信封 5×11 1/2 英寸                   ,16558,7199,482,794,181,181,23"
    cprPK24 = 24            ' "C 尺寸工作单                              ,16558,7199,482,794,181,181,24"
    cprPK25 = 25            ' "D 尺寸工作单                              ,16558,7199,482,794,181,181,25"
    cprPK26 = 26            ' "E 尺寸工作单                              ,16558,7199,482,794,181,181,26"
    cprPK27 = 27            ' "DL 型信封 110×220 毫米                   ,6237,12474,482,805,181,181,27"
    cprPK28 = 28            ' "C5 型信封 162×229 毫米                   ,9185,12984,482,799,181,176,28"
    cprPK29 = 29            ' "C3 型信封 324×458 毫米                   ,25969,18371,482,794,181,176,29"
    cprPK30 = 30            ' "C4 型信封 229×324 毫米                   ,18371,12981,482,794,181,176,30"
    cprPK31 = 31            ' "C6 型信封 114×162 毫米                   ,9183,6462,482,794,181,176,31"
    cprPK32 = 32            ' "C65 型信封114×229 毫米                   ,18371,6462,482,794,181,176,32"
    cprPK33 = 33            ' "B4 型信封 250×353 毫米                   ,20010,14171,482,794,181,176,33"
    cprPK34 = 34            ' "B5 型信封176×250 毫米                    ,9979,14175,482,799,181,193,34"
    cprPK35 = 35            ' "B6 型信封 176×125 毫米                   ,14175,7086,482,799,181,193,35"
    cprPK36 = 36            ' "信封 110×230 毫米                        ,13037,6237,482,799,181,193,36"
    cprPK37 = 37            ' "信封大王 3 7/8×7 1/2 英寸                ,5579,10801,482,794,181,181,37"
    cprPK38 = 38            ' "信封 3 5/8×6 1/2 英寸                    ,9359,5219,482,794,181,181,38"
    cprPK39 = 39            ' "U.S. 标准复写簿 14 7/8×11 英寸           ,15842,21421,0,0,0,1848,39"
    cprPK40 = 40            ' "德国标准复写簿 8 1/2×12 英寸             ,17282,12242,0,0,0,0,40"
    cprPK41 = 41            ' "德国法律复写簿 8 1/2×13 英寸             ,18722,12242,0,0,0,0,41"
    cprPKCustom = 256       ' "自定义纸张                                ,22680,16443,482,0,0,0,256"
End Enum

Public Enum PaperOrientEnum
    cprPOPortrait = vbPRORPortrait          '纵向
    cprPOLandscape = vbPRORLandscape        '横向
End Enum

'#########################################################################
'   属性变量
'#########################################################################

Private mvarAutoDetectURL As Boolean
Private mvarBackColor As OLE_COLOR
Private mvarBorder As Boolean
Private mvarDefaultTabStop As Single
Private mvarDoDefaultURLClick As Boolean
Private mvarEnabled As Boolean
Private mvarFileName As String
Private mvarFoot As String
Private mvarForceEdit As Boolean
Private mvarHead As String
Private mvarMarginBottom As Long
Private mvarMarginLeft As Long
Private mvarMarginRight As Long
Private mvarMarginTop As Long
Private mvarModified As Boolean
Private mvarPaperCount As Long
Private mvarPaperColor As OLE_COLOR
Private mvarPaperHeight As Long
Private mvarPaperWidth As Long
Private mvarPicture As StdPicture
Private mvarReadOnly As Boolean
Private mvarTitle As String
Private mvarTransparent As Boolean
Private mvarViewMode As ViewModeEnum
Private mvarZoomFactor As Double
Private mvarShowPageNumber As Boolean

'#########################################################################
'   属性方法
'#########################################################################

Public Property Let AutoDetectURL(ByVal vData As Boolean)
    mvarAutoDetectURL = vData
End Property

Public Property Get AutoDetectURL() As Boolean
    AutoDetectURL = mvarAutoDetectURL
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let Border(ByVal vData As Boolean)
    mvarBorder = vData
End Property

Public Property Get Border() As Boolean
    Border = mvarBorder
End Property

Public Property Let DefaultTabStop(ByVal vData As Single)
    mvarDefaultTabStop = vData
End Property

Public Property Get DefaultTabStop() As Single
    DefaultTabStop = mvarDefaultTabStop
End Property

Public Property Let DoDefaultURLClick(ByVal vData As Boolean)
    mvarDoDefaultURLClick = vData
End Property

Public Property Get DoDefaultURLClick() As Boolean
    DoDefaultURLClick = mvarDoDefaultURLClick
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property

Public Property Let Foot(ByVal vData As String)
    mvarFoot = vData
End Property

Public Property Get Foot() As String
    Foot = mvarFoot
End Property

Public Property Let ForceEdit(ByVal vData As Boolean)
    mvarForceEdit = vData
End Property

Public Property Get ForceEdit() As Boolean
    ForceEdit = mvarForceEdit
End Property

Public Property Let Head(ByVal vData As String)
    mvarHead = vData
End Property

Public Property Get Head() As String
    Head = mvarHead
End Property

Public Property Let MarginBottom(ByVal vData As Long)
    mvarMarginBottom = vData
End Property

Public Property Get MarginBottom() As Long
    MarginBottom = mvarMarginBottom
End Property

Public Property Let MarginLeft(ByVal vData As Long)
    mvarMarginLeft = vData
End Property

Public Property Get MarginLeft() As Long
    MarginLeft = mvarMarginLeft
End Property

Public Property Let MarginRight(ByVal vData As Long)
    mvarMarginRight = vData
End Property

Public Property Get MarginRight() As Long
    MarginRight = mvarMarginRight
End Property

Public Property Let MarginTop(ByVal vData As Long)
    mvarMarginTop = vData
End Property

Public Property Get MarginTop() As Long
    MarginTop = mvarMarginTop
End Property

Public Property Let Modified(ByVal vData As Boolean)
    mvarModified = vData
End Property

Public Property Get Modified() As Boolean
    Modified = mvarModified
End Property

Public Property Let PaperCount(ByVal vData As Long)
    mvarPaperCount = vData
End Property

Public Property Get PaperCount() As Long
    PaperCount = mvarPaperCount
End Property

Public Property Let PaperColor(ByVal vData As OLE_COLOR)
    mvarPaperColor = vData
End Property

Public Property Get PaperColor() As OLE_COLOR
    PaperColor = mvarPaperColor
End Property

Public Property Let PaperHeight(ByVal vData As Long)
    mvarPaperHeight = vData
End Property

Public Property Get PaperHeight() As Long
    PaperHeight = mvarPaperHeight
End Property

Public Property Let PaperWidth(ByVal vData As Long)
    mvarPaperWidth = vData
End Property

Public Property Get PaperWidth() As Long
    PaperWidth = mvarPaperWidth
End Property

Public Property Set Picture(ByVal vData As StdPicture)
    Set mvarPicture = vData
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mvarPicture
End Property

Public Property Let ReadOnly(ByVal vData As Boolean)
    mvarReadOnly = vData
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mvarReadOnly
End Property

Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
End Property

Public Property Get Title() As String
    Title = mvarTitle
End Property

Public Property Let Transparent(ByVal vData As Boolean)
    mvarTransparent = vData
End Property

Public Property Get Transparent() As Boolean
    Transparent = mvarTransparent
End Property

Public Property Let ViewMode(ByVal vData As ViewModeEnum)
    mvarViewMode = vData
End Property

Public Property Get ViewMode() As ViewModeEnum
    ViewMode = mvarViewMode
End Property

Public Property Let ZoomFactor(ByVal vData As Double)
    mvarZoomFactor = vData
End Property

Public Property Get ZoomFactor() As Double
    ZoomFactor = mvarZoomFactor
End Property

Public Property Let ShowPageNumber(ByVal vData As Boolean)
Attribute ShowPageNumber.VB_Description = "当前选中内容的RTF字符串。"
    mvarShowPageNumber = vData
End Property

Public Property Get ShowPageNumber() As Boolean
    ShowPageNumber = mvarShowPageNumber
End Property

Private Sub Class_Initialize()
    ZoomFactor = 1#
End Sub
