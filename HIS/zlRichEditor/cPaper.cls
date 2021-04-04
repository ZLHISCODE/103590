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
Option Explicit
'#########################################################################
'   枚举常量
'#########################################################################

'用途: 视图模式
Public Enum ViewModeEnum
    cprNormal = 0   '普通视图
    cprOutLine = 1  '大纲视图
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

Public Property Let TRANSPARENT(ByVal vData As Boolean)
    mvarTransparent = vData
End Property

Public Property Get TRANSPARENT() As Boolean
    TRANSPARENT = mvarTransparent
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
