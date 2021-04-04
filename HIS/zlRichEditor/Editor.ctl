VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Editor 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   11.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4755
   ScaleWidth      =   7440
   ToolboxBitmap   =   "Editor.ctx":0000
   Begin VB.VScrollBar VS 
      Height          =   2940
      LargeChange     =   20
      Left            =   6870
      Max             =   0
      TabIndex        =   20
      Top             =   1050
      Width           =   250
   End
   Begin VB.PictureBox picMarginR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   4350
      MouseIcon       =   "Editor.ctx":0532
      ScaleHeight     =   2745
      ScaleWidth      =   255
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1230
      Width           =   250
   End
   Begin VB.PictureBox picMarginL 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   1965
      MouseIcon       =   "Editor.ctx":0684
      ScaleHeight     =   2745
      ScaleWidth      =   255
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1260
      Width           =   250
   End
   Begin zlRichEditor.Document RTBTmp 
      Height          =   210
      Left            =   3960
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   370
      Border          =   0   'False
   End
   Begin zlRichEditor.Progress Progress1 
      Height          =   240
      Left            =   3825
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3915
      Visible         =   0   'False
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   423
   End
   Begin VB.HScrollBar HS 
      Height          =   250
      Left            =   1200
      TabIndex        =   13
      Top             =   4200
      Width           =   735
   End
   Begin zlRichEditor.FButton btnNormal 
      Height          =   255
      Left            =   120
      Top             =   4200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Value           =   -1  'True
      IsOptButton     =   -1  'True
      Picture         =   "Editor.ctx":07D6
      MaskColor       =   16777215
   End
   Begin zlRichEditor.FButton btnPaper 
      Height          =   255
      Left            =   480
      Top             =   4200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      IsOptButton     =   -1  'True
      Picture         =   "Editor.ctx":0832
      MaskColor       =   16777215
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   5085
      Picture         =   "Editor.ctx":08D5
      ScaleHeight     =   120
      ScaleWidth      =   120
      TabIndex        =   12
      Top             =   1530
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picUI 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   585
      ScaleHeight     =   735
      ScaleWidth      =   960
      TabIndex        =   11
      Top             =   3195
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   5715
      ScaleHeight     =   390
      ScaleWidth      =   570
      TabIndex        =   10
      Top             =   3420
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picBuff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   5265
      ScaleHeight     =   510
      ScaleWidth      =   645
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSComctlLib.ImageList ImlScroll 
      Left            =   6660
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   9
      ImageHeight     =   9
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":09D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":0A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":0AA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":0AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":0B53
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Editor.ctx":0BAD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picNull 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   6795
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4230
      Width           =   250
   End
   Begin VB.PictureBox picHRuler 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   165
      ScaleHeight     =   390
      ScaleWidth      =   4215
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   675
      Width           =   4215
      Begin VB.PictureBox picHRulerHead 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   0
         ScaleHeight     =   420
         ScaleWidth      =   360
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   360
         Begin VB.Image Image1 
            Height          =   240
            Left            =   70
            Picture         =   "Editor.ctx":0C0D
            Top             =   90
            Width           =   240
         End
      End
      Begin zlRichEditor.HRuler HRuler 
         Height          =   390
         Left            =   885
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   688
         RulerLength     =   112
         RightMargin     =   1140
         LeftMargin      =   1140
         AllowMargins    =   1
         Quantise        =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin zlRichEditor.Document RTBNormal 
      Height          =   1050
      Left            =   2385
      TabIndex        =   1
      Top             =   1350
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1852
      BackColor       =   0
      Border          =   0   'False
   End
   Begin zlRichEditor.Paper RTBPaper 
      Height          =   1275
      Index           =   1
      Left            =   4905
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1905
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2249
      ShowPageNumber  =   -1  'True
   End
   Begin VB.PictureBox picShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   780
      ScaleHeight     =   375
      ScaleWidth      =   330
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2655
      Visible         =   0   'False
      Width           =   330
   End
   Begin RichTextLib.RichTextBox rtbBuff 
      Height          =   465
      Left            =   2295
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3810
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   820
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"Editor.ctx":0C75
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin zlRichEditor.Document RTBHead 
      Height          =   210
      Left            =   2640
      TabIndex        =   16
      Top             =   90
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   370
   End
   Begin zlRichEditor.Document RTBFoot 
      Height          =   210
      Left            =   3360
      TabIndex        =   17
      Top             =   4380
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   370
   End
   Begin VB.Label lblThis 
      BackStyle       =   0  'Transparent
      Caption         =   "中联图文编辑控件"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   675
      TabIndex        =   5
      Top             =   135
      Width           =   5190
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   45
      Picture         =   "Editor.ctx":0D12
      Top             =   45
      Width           =   480
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'######################################################################################
'##模 块 名：Editor.ctl
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：对外接口的最终编辑控件。封装了普通、页面及大纲视图模式，映射相关消息与属性。
'##版    本：
'######################################################################################

Option Explicit

'#############################################################################################################
'##     局部变量
'#############################################################################################################

Private m_hWnd As Long                  '控件的 hWnd
Private m_hWndParent  As Long           '父窗体的 hWnd
'Private m_TOM As New cTextDocument      'TOM 3.0 模型，核心对象
Private mfrmFindText As New frmFindText '查找替换窗体

Private m_sText As String               '文本内容

Private Const BORDERWIDTH = 15         '边框

'#############################################################################################################
'##     独立属性
'#############################################################################################################

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
Private mvarPageCount As Long               '独有属性
Private mvarCurPage As Long                 '当前页
Private mvarStartPage As Long               '实际显示的起始页
Private mvarEndPage As Long                 '实际显示的终止页
Private mvarWithViewButtonas As Boolean     '是否有切换视图按钮
Private mvarPaperKind As PaperKindEnum      '纸张类型属性
Private mvarPaperOrient As PaperOrientEnum  '纸张方向
Private mvarInProcessing As Boolean         '分页处理中...
Private mvarShowRuler As Boolean            '是否显示标尺

Private mvarHeadFontName As String
Private mvarHeadFontSize As Long
Private mvarHeadFontBold As Boolean
Private mvarHeadFontItalic As Boolean
Private mvarHeadFontUnderline As Boolean
Private mvarHeadFontStrikethrough As Boolean
Private mvarHeadFontColor As OLE_COLOR
Private mvarHeadFile As String

Private mvarFootFontName As String
Private mvarFootFontSize As Long
Private mvarFootFontBold As Boolean
Private mvarFootFontItalic As Boolean
Private mvarFootFontUnderline As Boolean
Private mvarFootFontStrikethrough As Boolean
Private mvarFootFontColor As OLE_COLOR
Private mvarFootFile As String

'#############################################################################################################
'##     事件声明（用于映射事件）
'#############################################################################################################

Public Event Change(ViewMode As ViewModeEnum)    '内容改变！
Public Event MouseWheel(ViewMode As ViewModeEnum, bBackDirection As Boolean, Shift As Integer, x As Single, y As Single, Value As Single)   '鼠标滚轮事件
Public Event Zoom(ViewMode As ViewModeEnum, NewFactor As Double)   '用户通过Ctrl＋鼠标来改变了缩放比例！
Public Event Resize(ViewMode As ViewModeEnum)    '控件尺寸改变
Public Event RequestLine(ViewMode As ViewModeEnum)              '请求行数改变
Public Event SelChange(ViewMode As ViewModeEnum, ByVal lStart As Long, ByVal lEnd As Long)  '选择区域改变
Public Event LinkEvent(ViewMode As ViewModeEnum, ByVal iType As LinkEventTypeEnum, ByVal lStart As Long, ByVal lEnd As Long)     '链接事件
Public Event ModifyProtected(ViewMode As ViewModeEnum, ByRef bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)            '试图编辑受保护区域
Public Event BeforeKeyDown(ViewMode As ViewModeEnum, KeyCode As Integer, Shift As Integer)
Public Event KeyDown(ViewMode As ViewModeEnum, KeyCode As Integer, Shift As Integer)
Public Event KeyPress(ViewMode As ViewModeEnum, KeyAscii As Integer)
Public Event KeyUp(ViewMode As ViewModeEnum, KeyCode As Integer, Shift As Integer)
Public Event MouseDown(ViewMode As ViewModeEnum, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(ViewMode As ViewModeEnum, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(ViewMode As ViewModeEnum, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event RequestRightMenu(ViewMode As ViewModeEnum, Shift As Integer, x As Single, y As Single)
Public Event Click(ViewMode As ViewModeEnum)        '单击
Public Event DblClick(ViewMode As ViewModeEnum)     '双击
Public Event PressTabKey()                          '按下TAB按钮
Public Event GetDelCharColor(ByRef COLOR As OLE_COLOR)     '获取删除字符的颜色
Public Event GetNewCharColor(ByRef COLOR As OLE_COLOR)     '获取新增字符的颜色
Public Event IsDelCharColor(ByVal COLOR As OLE_COLOR, ByRef blnIsDelCharColor As Boolean)   '判断是否是删除字符的颜色
Public Event IsNewCharColor(ByVal COLOR As OLE_COLOR, ByRef blnIsNewCharColor As Boolean)   '判断是否是新增字符的颜色
Public Event UIOpen(ByRef UIhWnd As Long, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long) '打开UI接口
Public Event UIMoved(ByRef UIhWnd As Long, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long) '打开UI接口
Public Event UIClose(ByRef UIhWnd As Long)  '关闭UI接口
Public Event UIClick(ViewMode As ViewModeEnum)   '关闭UI接口

'#############################################################################################################
'##     公共属性（属性映射）
'#############################################################################################################
Public Property Let HeadFontFormat(vData As String)
    On Error Resume Next
    Dim t As Variant
    t = Split(vData, "|")
    mvarHeadFontName = t(0)
    mvarHeadFontSize = t(1)
    mvarHeadFontBold = t(2)
    mvarHeadFontItalic = t(3)
    mvarHeadFontUnderline = t(4)
    mvarHeadFontStrikethrough = t(5)
    mvarHeadFontColor = t(6)
    Err.Clear
End Property

Public Property Get HeadFontFormat() As String
    HeadFontFormat = mvarHeadFontName & "|" & mvarHeadFontSize & "|" & mvarHeadFontBold & "|" & mvarHeadFontItalic & "|" & mvarHeadFontUnderline & "|" & mvarHeadFontStrikethrough & "|" & mvarHeadFontColor
End Property

Public Property Let FootFontFormat(vData As String)
    On Error Resume Next
    Dim t As Variant
    t = Split(vData, "|")
    mvarFootFontName = t(0)
    mvarFootFontSize = t(1)
    mvarFootFontBold = t(2)
    mvarFootFontItalic = t(3)
    mvarFootFontUnderline = t(4)
    mvarFootFontStrikethrough = t(5)
    mvarFootFontColor = t(6)
    Err.Clear
End Property

Public Property Get FootFontFormat() As String
    FootFontFormat = mvarFootFontName & "|" & mvarFootFontSize & "|" & mvarFootFontBold & "|" & mvarFootFontItalic & "|" & mvarFootFontUnderline & "|" & mvarFootFontStrikethrough & "|" & mvarFootFontColor
End Property

Public Property Get UIhWmd() As Long
    UIhWmd = picUI.hwnd
End Property

Public Property Get UIVisibled() As BOOL
    UIVisibled = picUI.Visible
End Property

Public Property Get UILeft() As Long
    UILeft = picUI.Left
End Property

Public Property Get UITop() As Long
    UITop = picUI.Top
End Property

Public Property Let UIWidth(vData As Long)
    picUI.Width = vData
End Property

Public Property Get UIWidth() As Long
    UIWidth = picUI.Width
End Property

Public Property Let UIHeight(vData As Long)
    picUI.Height = vData
End Property

Public Property Get UIHeight() As Long
    UIHeight = picUI.Height
End Property

Public Property Let TargetDC(ByVal vData As Long)
    gTargetDC = vData
End Property

Public Property Get TargetDC() As Long
    TargetDC = gTargetDC
End Property
Public Sub ResetWYSIWYG()
    '重新刷新“所见即所得”显示
    gTargetDC = picBuff.hDC     '解决民航医院预览时右边超出的问题！（只有以屏幕为度量）
    RTBNormal.ResetWYSIWYG
    RTBHead.ResetWYSIWYG
    RTBFoot.ResetWYSIWYG
End Sub

Public Property Let AuditMode(ByVal vData As Boolean)   '审核级别
    RTBNormal.AuditMode = vData
    PropertyChanged "AuditMode"
End Property

Public Property Get AuditMode() As Boolean              '审核级别
    AuditMode = RTBNormal.AuditMode
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Extender.Parent
End Property

Public Property Get OriginRTB() As Object
    Dim Obj As RichTextBox
    Set Obj = RTBNormal.OriginRTB
    Set OriginRTB = Obj
End Property

Public Sub ResetAuditText()
    '重置审核状态下的选中的修订文本
    If Me.AuditMode Then RTBNormal.ResetAuditText
End Sub

Public Sub AcceptAuditText()
    '接受审核状态下的选中的修订文本
    If Me.AuditMode Then RTBNormal.AcceptAuditText
End Sub

Public Property Get TOM_Origin() As cTextDocument
    Set TOM_Origin = RTBNormal.TOM
End Property

Public Property Get TOM() As cTextDocument
    Select Case mvarViewMode
    Case cprNormal
        Set TOM = RTBNormal.TOM
    Case cprPaper
        Set TOM = RTBNormal.TOM
    End Select
End Property

Public Property Let WithViewButtonas(ByVal vData As Boolean)
    mvarWithViewButtonas = vData
    PropertyChanged "WithViewButtonas"
End Property

Public Property Get WithViewButtonas() As Boolean
    WithViewButtonas = mvarWithViewButtonas
End Property

Public Property Let InProcessing(ByVal vData As Boolean)
    mvarInProcessing = vData
    PropertyChanged "InProcessing"
End Property

Public Property Get InProcessing() As Boolean
    InProcessing = mvarInProcessing
End Property

Public Property Let ShowRuler(ByVal vData As Boolean)
    mvarShowRuler = vData
    Call UserControl_Resize
    PropertyChanged "ShowRuler"
End Property

Public Property Get ShowRuler() As Boolean
    ShowRuler = mvarShowRuler
End Property

Public Property Let HeadFontName(ByVal vData As String)
    mvarHeadFontName = vData
    PropertyChanged "HeadFontName"
End Property

Public Property Get HeadFontName() As String
    HeadFontName = mvarHeadFontName
End Property

Public Property Let HeadFontSize(ByVal vData As Long)
    mvarHeadFontSize = vData
    PropertyChanged "HeadFontSize"
End Property

Public Property Get HeadFontSize() As Long
    HeadFontSize = mvarHeadFontSize
End Property

Public Property Let HeadFontBold(ByVal vData As Boolean)
    mvarHeadFontBold = vData
    PropertyChanged "HeadFontBold"
End Property

Public Property Get HeadFontBold() As Boolean
    HeadFontBold = mvarHeadFontBold
End Property

Public Property Let HeadFontItalic(ByVal vData As Boolean)
    mvarHeadFontItalic = vData
    PropertyChanged "HeadFontItalic"
End Property

Public Property Get HeadFontItalic() As Boolean
    HeadFontItalic = mvarHeadFontItalic
End Property

Public Property Let HeadFontUnderline(ByVal vData As Boolean)
    mvarHeadFontUnderline = vData
    PropertyChanged "HeadFontUnderline"
End Property

Public Property Get HeadFontUnderline() As Boolean
    HeadFontUnderline = mvarHeadFontUnderline
End Property

Public Property Let HeadFontStrikethrough(ByVal vData As Boolean)
    mvarHeadFontStrikethrough = vData
    PropertyChanged "HeadFontStrikethrough"
End Property

Public Property Get HeadFontStrikethrough() As Boolean
    HeadFontStrikethrough = mvarHeadFontStrikethrough
End Property

Public Property Let HeadFontColor(ByVal vData As OLE_COLOR)
    mvarHeadFontColor = vData
    PropertyChanged "HeadFontColor"
End Property

Public Property Get HeadFontColor() As OLE_COLOR
    HeadFontColor = mvarHeadFontColor
End Property

Public Property Let FootFontName(ByVal vData As String)
    mvarFootFontName = vData
    PropertyChanged "FootFontName"
End Property

Public Property Get FootFontName() As String
    FootFontName = mvarFootFontName
End Property

Public Property Let FootFontSize(ByVal vData As Long)
    mvarFootFontSize = vData
    PropertyChanged "FootFontSize"
End Property

Public Property Get FootFontSize() As Long
    FootFontSize = mvarFootFontSize
End Property

Public Property Let FootFontBold(ByVal vData As Boolean)
    mvarFootFontBold = vData
    PropertyChanged "FootFontBold"
End Property

Public Property Get FootFontBold() As Boolean
    FootFontBold = mvarFootFontBold
End Property

Public Property Let FootFontItalic(ByVal vData As Boolean)
    mvarFootFontItalic = vData
    PropertyChanged "FootFontItalic"
End Property

Public Property Get FootFontItalic() As Boolean
    FootFontItalic = mvarFootFontItalic
End Property

Public Property Let FootFontUnderline(ByVal vData As Boolean)
    mvarFootFontUnderline = vData
    PropertyChanged "FootFontUnderline"
End Property

Public Property Get FootFontUnderline() As Boolean
    FootFontUnderline = mvarFootFontUnderline
End Property

Public Property Let FootFontStrikethrough(ByVal vData As Boolean)
    mvarFootFontStrikethrough = vData
    PropertyChanged "FootFontStrikethrough"
End Property

Public Property Get FootFontStrikethrough() As Boolean
    FootFontStrikethrough = mvarFootFontStrikethrough
End Property

Public Property Let FootFontColor(ByVal vData As OLE_COLOR)
    mvarFootFontColor = vData
    PropertyChanged "FootFontColor"
End Property

Public Property Get FootFontColor() As OLE_COLOR
    FootFontColor = mvarFootFontColor
End Property
Public Property Let PaperKind(ByVal vData As PaperKindEnum)
    mvarPaperKind = vData
    RTBNormal.PaperKind = vData
    RTBHead.PaperKind = vData
    RTBFoot.PaperKind = vData
    PropertyChanged "PaperKind"
End Property

Public Property Get PaperKind() As PaperKindEnum
    PaperKind = mvarPaperKind
End Property

Public Property Let PaperOrient(ByVal vData As PaperOrientEnum)
    mvarPaperOrient = vData
    RTBNormal.PaperOrient = vData
    PropertyChanged "PaperOrient"
End Property

Public Property Get PaperOrient() As PaperOrientEnum
    PaperOrient = mvarPaperOrient
End Property

Public Property Let AutoDetectURL(ByVal vData As Boolean)
    mvarAutoDetectURL = vData
    RTBNormal.AutoDetectURL = vData
    PropertyChanged "AutoDetectURL"
End Property

Public Property Get AutoDetectURL() As Boolean
    AutoDetectURL = mvarAutoDetectURL
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
    If Ambient.UserMode Then
        UserControl.BackColor = vData
    Else
        UserControl.BackColor = vbWhite
    End If
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Get Border() As Boolean
    Border = mvarBorder
End Property

Public Property Let Border(ByVal vData As Boolean)
    Dim dwStyle As Long
    Dim dwExStyle As Long

    If m_hWnd <> 0 Then
        ' Make sure that the RichEdit never has a border:
        dwStyle = GetWindowLong(m_hWnd, GWL_STYLE)
        dwExStyle = GetWindowLong(m_hWnd, GWL_EXSTYLE)
        dwStyle = dwStyle And Not ES_SUNKEN
        dwExStyle = dwExStyle And Not WS_EX_CLIENTEDGE
        SetWindowLong m_hWnd, GWL_STYLE, dwStyle
        SetWindowLong m_hWnd, GWL_EXSTYLE, dwExStyle
        pStyleChanged
    End If
    UserControl.BorderStyle() = Abs(vData)
    
    mvarBorder = vData
    PropertyChanged "Border"
End Property

Public Property Get CanCopy() As Boolean
    Select Case mvarViewMode
    Case cprNormal
        CanCopy = RTBNormal.CanCopy
    Case cprPaper
        CanCopy = False
    End Select
End Property

Public Property Get CanPaste() As Boolean
    Select Case mvarViewMode
    Case cprNormal
        CanPaste = RTBNormal.CanPaste
    Case cprPaper
        CanPaste = False
    End Select
End Property

Public Property Get CanRedo() As Boolean
    Select Case mvarViewMode
    Case cprNormal
        CanRedo = RTBNormal.CanRedo
    Case cprPaper
        CanRedo = False
    End Select
End Property

Public Property Get CanUndo() As Boolean
    Select Case mvarViewMode
    Case cprNormal
        CanUndo = RTBNormal.CanUndo
    Case cprPaper
        CanUndo = False
    End Select
End Property

Public Property Get CanDelete() As Boolean
    Select Case mvarViewMode
    Case cprNormal
        CanDelete = RTBNormal.CanDelete
    Case cprPaper
        CanDelete = False
    End Select
End Property

Public Property Get CurrentColumn() As Long
    Select Case mvarViewMode
    Case cprNormal
        CurrentColumn = RTBNormal.CurrentColumn
    Case cprPaper
        CurrentColumn = RTBNormal.CurrentColumn
    End Select
End Property

Public Property Get CurrentLine() As Long
    Select Case mvarViewMode
    Case cprNormal
        CurrentLine = RTBNormal.CurrentLine
    Case cprPaper
        CurrentLine = RTBNormal.CurrentLine
    End Select
End Property

Public Property Let DefaultTabStop(ByVal vData As Single)
    mvarDefaultTabStop = vData
    RTBNormal.DefaultTabStop = vData
    PropertyChanged "DefaultTabStop"
End Property

Public Property Get DefaultTabStop() As Single
    DefaultTabStop = mvarDefaultTabStop
End Property

Public Property Let DoDefaultURLClick(ByVal vData As Boolean)
    mvarDoDefaultURLClick = vData
    RTBNormal.DoDefaultURLClick = vData
    PropertyChanged "DoDefaultURLClick"
End Property

Public Property Get DoDefaultURLClick() As Boolean
    DoDefaultURLClick = mvarDoDefaultURLClick
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    mvarEnabled = vData
    RTBNormal.Enabled = vData
    UserControl.Enabled = vData
    PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
    Enabled = mvarEnabled
End Property

Public Property Let FileName(ByVal vData As String)
    Dim strTemp As String
    mvarFileName = vData
    If vData <> "" Then
        strTemp = Mid(vData, InStrRev(vData, "\") + 1)
        Me.Title = Left(strTemp, Len(strTemp) - 4)
    End If
    PubInfo.FileName = vData
    PropertyChanged "FileName"
End Property

Public Property Get FileName() As String
    FileName = mvarFileName
End Property
Public Property Let HeadFile(ByVal vData As String)
'由调用上级自行删除文件
    mvarHeadFile = vData
    RTBHead.OpenDoc vData
    RTBHead.ResetWYSIWYG
    PropertyChanged "HeadFile"
End Property
Public Property Get HeadFileText() As String
    HeadFileText = RTBHead.Text
End Property
Public Property Get HeadFileTextRTF() As String
    HeadFileTextRTF = RTBHead.TextRTF
End Property
Public Property Let HeadFileTextRTF(ByVal vData As String)
    On Error Resume Next
    RTBHead.TextRTF = vData
    RTBHead.ClearEndCrlfChar
    Err.Clear
    PropertyChanged "HeadFileTextRTF"
End Property
Public Property Get HeadFile() As String
    HeadFile = mvarHeadFile
End Property
Public Property Let FootFile(ByVal vData As String)
'由调用上级自行删除文件
    mvarFootFile = vData
    RTBFoot.OpenDoc vData
    RTBFoot.ResetWYSIWYG
    PropertyChanged "FootFile"
End Property
Public Property Get FootFileText() As String
    FootFileText = RTBFoot.Text
End Property
Public Property Get FootFileTextRTF() As String
    FootFileTextRTF = RTBFoot.TextRTF
End Property
Public Property Let FootFileTextRTF(ByVal vData As String)
    On Error Resume Next
    RTBFoot.TextRTF = vData
    RTBFoot.ClearEndCrlfChar
    Err.Clear
    PropertyChanged "FootFileTextRTF"
End Property
Public Property Get FootFile() As String
    FootFile = mvarFootFile
End Property
Public Property Get FirstVisibleLine() As Long
    Select Case mvarViewMode
    Case cprNormal
        FirstVisibleLine = RTBNormal.FirstVisibleLine
    Case cprPaper
        FirstVisibleLine = RTBNormal.FirstVisibleLine
    End Select
End Property

Public Property Let Foot(ByVal vData As String)
    mvarFoot = vData
    PubInfo.Foot = vData
    PropertyChanged "Foot"
End Property

Public Property Get Foot() As String
    Foot = mvarFoot
End Property

Public Property Let ForceEdit(ByVal vData As Boolean)
    mvarForceEdit = vData
    RTBNormal.ForceEdit = vData
    PropertyChanged "ForceEdit"
End Property

Public Property Get ForceEdit() As Boolean
    ForceEdit = RTBNormal.ForceEdit
End Property

Public Property Let Head(ByVal vData As String)
    mvarHead = vData
    PubInfo.Head = vData
    PropertyChanged "Head"
End Property

Public Property Get Head() As String
    Head = mvarHead
End Property

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Public Property Get hWndRTB() As Long
    Select Case mvarViewMode
    Case cprNormal
        hWndRTB = RTBNormal.hWndRTB
    Case cprPaper
        hWndRTB = RTBNormal.hWndRTB
    End Select
End Property

Public Property Get LineCount() As Long
    LineCount = RTBNormal.LineCount
End Property

Public Property Let MarginBottom(ByVal vData As Long)
    mvarMarginBottom = vData
    RTBNormal.MarginBottom = vData
    PubInfo.MarginBottom = vData
    PropertyChanged "MarginBottom"
End Property

Public Property Get MarginBottom() As Long
    MarginBottom = mvarMarginBottom
End Property

Public Property Let MarginLeft(ByVal vData As Long)
    mvarMarginLeft = vData
    RTBNormal.MarginLeft = vData
    RTBHead.MarginLeft = vData
    RTBFoot.MarginLeft = vData
    PubInfo.MarginLeft = vData
    HRuler.LeftMargin = vData
    HRuler.Left = picMarginL.Left
    PropertyChanged "MarginLeft"
End Property

Public Property Get MarginLeft() As Long
    MarginLeft = mvarMarginLeft
End Property

Public Property Let MarginRight(ByVal vData As Long)
    mvarMarginRight = vData
    RTBNormal.MarginRight = vData
    RTBHead.MarginRight = vData
    RTBFoot.MarginRight = vData
    PubInfo.MarginRight = vData
    HRuler.RightMargin = vData
    PropertyChanged "MarginRight"
End Property

Public Property Get MarginRight() As Long
    MarginRight = mvarMarginRight
End Property

Public Property Let MarginTop(ByVal vData As Long)
    mvarMarginTop = vData
    RTBNormal.MarginTop = vData
    PubInfo.MarginTop = vData
    PropertyChanged "MarginTop"
End Property

Public Property Get MarginTop() As Long
    MarginTop = mvarMarginTop
End Property

Public Property Let Modified(ByVal vData As Boolean)
    mvarModified = vData
    RTBNormal.Modified = vData
    PropertyChanged "Modified"
End Property

Public Property Get Modified() As Boolean
    Modified = RTBNormal.Modified
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Let PaperColor(ByVal vData As OLE_COLOR)
    mvarPaperColor = vData
    picMarginR.BackColor = vData
    picMarginL.BackColor = vData
    RTBNormal.PaperColor = vData
    PropertyChanged "PaperColor"
End Property

Public Property Get PaperColor() As OLE_COLOR
    If Ambient.UserMode Then
        Select Case mvarViewMode
        Case cprNormal
            PaperColor = RTBNormal.PaperColor
        Case cprPaper
            PaperColor = RTBNormal.PaperColor
        End Select
    Else
        PaperColor = mvarPaperColor
    End If
End Property

Public Property Let PaperHeight(ByVal vData As Long)
    mvarPaperHeight = vData
    RTBNormal.PaperHeight = vData
    PubInfo.PaperHeight = vData
    PropertyChanged "PaperHeight"
End Property

Public Property Get PaperHeight() As Long
    PaperHeight = mvarPaperHeight
End Property

Public Property Let PaperWidth(ByVal vData As Long)
    mvarPaperWidth = vData
    RTBNormal.PaperWidth = vData
    RTBHead.PaperWidth = vData
    RTBFoot.PaperWidth = vData
    PubInfo.PaperWidth = vData
    HRuler.Width = vData
    PropertyChanged "PaperWidth"
End Property

Public Property Get PaperWidth() As Long
    PaperWidth = mvarPaperWidth
End Property

Public Property Let Picture(ByVal vData As StdPicture)
    Set mvarPicture = vData
    Set PubInfo.Picture = vData
    PropertyChanged "Picture"
End Property

Public Property Set Picture(ByVal vData As StdPicture)
    Set mvarPicture = vData
    Set PubInfo.Picture = vData
    PropertyChanged "Picture"
End Property

Public Property Get Picture() As StdPicture
    Set Picture = mvarPicture
End Property

Public Property Let ReadOnly(ByVal vData As Boolean)
    mvarReadOnly = vData
    RTBNormal.ReadOnly = vData
    PropertyChanged "ReadOnly"
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mvarReadOnly
End Property

Public Property Let SelLength(ByVal vData As Long)
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.SelLength = vData
    Case cprPaper
        RTBNormal.SelLength = vData
    End Select
    PropertyChanged "SelLength"
End Property

Public Property Get SelLength() As Long
    Select Case mvarViewMode
    Case cprNormal
        SelLength = RTBNormal.SelLength
    Case cprPaper
        SelLength = RTBNormal.SelLength
    End Select
End Property

Public Property Let SelRTF(ByVal vData As String)
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.SelRTF = vData
    Case cprPaper
        '
    End Select
    PropertyChanged "SelRTF"
End Property

Public Property Get SelRTF() As String
    Select Case mvarViewMode
    Case cprNormal
        SelRTF = RTBNormal.SelRTF
    Case cprPaper
        SelRTF = RTBNormal.SelRTF
    End Select
End Property

Public Property Let SelStart(ByVal vData As Long)
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.SelStart = vData
    Case cprPaper
        RTBNormal.SelStart = vData
    End Select
    PropertyChanged "SelStart"
End Property

Public Property Get SelStart() As Long
    Select Case mvarViewMode
    Case cprNormal
        SelStart = RTBNormal.SelStart
    Case cprPaper
        SelStart = RTBNormal.SelStart
    End Select
End Property

Public Property Let SelText(ByVal vData As String)
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.SelText = vData
    Case cprPaper
        '
    End Select
    PropertyChanged "SelText"
End Property

Public Property Get SelText() As String
    Select Case mvarViewMode
    Case cprNormal
        SelText = RTBNormal.SelText
    Case cprPaper
        SelText = RTBNormal.SelText
    End Select
End Property

Public Property Get Text() As String
    Text = RTBNormal.Text
End Property

Public Property Let Text(ByRef vData As String)
    RTBNormal.Text = vData
    PropertyChanged "Text"
End Property

Public Property Get TextRTF() As String
    TextRTF = RTBNormal.TextRTF
End Property

Public Property Let TextRTF(ByRef vData As String)
    RTBNormal.TextRTF = vData
    PropertyChanged "TextRTF"
End Property

Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
    RTBNormal.Title = vData
    PubInfo.Title = vData
    PropertyChanged "Title"
End Property

Public Property Get Title() As String
    Select Case mvarViewMode
    Case cprNormal
        Title = RTBNormal.Title
    Case cprPaper
        Title = PubInfo.Title
    End Select
End Property

Public Property Let Transparent(ByVal vData As Boolean)
    mvarTransparent = vData
    RTBNormal.Transparent = vData
    PropertyChanged "Transparent"
End Property

Public Property Get Transparent() As Boolean
    Transparent = mvarTransparent
End Property

Public Property Let ViewMode(ByVal vData As ViewModeEnum)
    Dim lStart As Long, lEnd As Long, lLength As Long
    Dim i As Long, strF As String
    
    On Error Resume Next
    '刷新公共属性（页面模式）
    PubInfo.MarginLeft = Me.MarginLeft
    PubInfo.MarginRight = Me.MarginRight
    PubInfo.MarginTop = Me.MarginTop
    PubInfo.MarginBottom = Me.MarginBottom
    PubInfo.PaperWidth = Me.PaperWidth
    PubInfo.PaperHeight = Me.PaperHeight
    PubInfo.Foot = Me.Foot
    PubInfo.Head = Me.Head
    PubInfo.PaperCount = Me.PageCount
    PubInfo.ShowPageNumber = Me.ShowPageNumber
    
    mvarInProcessing = True
    ForceEdit = True
    
    If Not ExistsPrinter Then
        gTargetDC = picBuff.hDC
    Else
        gTargetDC = Printer.hDC
    End If
    gTargetDC = picBuff.hDC     '解决民航医院预览时右边超出的问题！（只有以屏幕为度量）
    
    mvarViewMode = vData
    Select Case vData
    Case cprNormal
        ResetWYSIWYG
        picHRuler.Visible = True
        picMarginR.BackColor = mvarPaperColor
        picMarginL.BackColor = mvarPaperColor
        HRuler.Left = -mvarMarginLeft * mvarZoomFactor + 390
        HRuler.LeftMargin = mvarMarginLeft * mvarZoomFactor
        HRuler.RightMargin = mvarMarginRight * mvarZoomFactor
        For i = 1 To RTBPaper.UBound
            RTBPaper(i).Visible = False
            picShadow(i).Visible = False
        Next
        VS.Visible = True
        SetVSWithRtb
        RTBNormal.Visible = True
        RTBNormal.SetFocus
    Case cprPaper
        DoVirtualPrint
        
        ShowPages True
        picHRuler.Visible = True
        HRuler.Left = RTBPaper(1).Left
        HRuler.LeftMargin = mvarMarginLeft * mvarZoomFactor
        HRuler.RightMargin = mvarMarginRight * mvarZoomFactor
        For i = 1 To RTBPaper.UBound
            RTBPaper(i).Visible = True
            picShadow(i).Visible = True
        Next
        VS.Visible = True
        VS.Max = GetPrintHeight \ Screen.TwipsPerPixelY
        RTBNormal.Visible = False
        HS.Enabled = True
        RTBPaper(1).SetFocus
    End Select
    mvarInProcessing = False
    ForceEdit = False
    
    PubInfo.ViewMode = vData
    Call UserControl_Resize
    PropertyChanged "ViewMode"
    CloseUIInterface
End Property

Public Property Get ViewMode() As ViewModeEnum
    ViewMode = mvarViewMode
End Property

Public Property Let ZoomFactor(ByVal vData As Double)
    mvarZoomFactor = vData
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.ZoomFactor = mvarZoomFactor
    Case cprPaper
        PubInfo.ZoomFactor = mvarZoomFactor
    End Select
    PropertyChanged "ZoomFactor"
End Property

Public Property Get ZoomFactor() As Double
    Select Case mvarViewMode
    Case cprNormal
        ZoomFactor = RTBNormal.ZoomFactor
    Case cprPaper
        ZoomFactor = PubInfo.ZoomFactor
    End Select
End Property

Public Property Let PageCount(ByVal vData As Long)
    mvarPageCount = vData
    PubInfo.PaperCount = vData
    PropertyChanged "PageCount"
End Property

Public Property Get PageCount() As Long
    PageCount = mvarPageCount
End Property

Public Property Let CurPage(ByVal vData As Long)
    mvarCurPage = vData
    PropertyChanged "CurPage"
End Property

Public Property Get CurPage() As Long
    CurPage = mvarCurPage
End Property

Public Property Get ShowPageNumber() As Boolean
    ShowPageNumber = mvarShowPageNumber
End Property

Public Property Let ShowPageNumber(vData As Boolean)
    mvarShowPageNumber = vData
    PubInfo.ShowPageNumber = vData
    PropertyChanged "ShowPageNumber"
End Property

Public Property Let ProgressVisible(vData As Boolean)
    Progress1.Cls
    Progress1.Visible = vData
End Property

Public Property Get ProgressVisible() As Boolean
    ProgressVisible = Progress1.Visible
End Property

Public Property Let ProgressValue(vData As Single)
    Progress1.Value = vData
End Property

Public Property Get ProgressValue() As Single
    ProgressValue = Progress1.Value
End Property

Private Sub pStyleChanged(Optional ByVal hwnd As Long = 0)
   If hwnd = 0 Then hwnd = m_hWnd
   SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOACTIVATE
End Sub
Public Sub ResizeUIInterface(ByVal lWidth As Long, lHeight As Long)
    '显示UI接口容器（当前内容必须是图片）
    If Me.ViewMode <> cprNormal Then Exit Sub
    Dim mIRichEditOle As olelib.IRichEditOle
    Dim mReObject As olelib.REOBJECT
    Dim mIOleObject As olelib.IOleObject
    Dim pSize As olelib.SIZE
    Dim pt As olelib.Point
    picUI.Cls
    If RTBNormal.Selection.GetType = cprSTPicture Then
        Dim lX As Long, lX1 As Long, lY1 As Long, lX2 As Long, lY2 As Long, lTMP As Long
        Dim lLeft As Long, lTOp As Long
        Dim lngSpaceBefore As Long, lngLinespace As Long
        lngLinespace = RTBNormal.Selection.Para.LineSpacing
        lngSpaceBefore = RTBNormal.Selection.Para.SpaceBefore * Screen.TwipsPerPixelX
        
        '获取精确的高度和宽度（采用OLE接口）
        SendMessage RTBNormal.hWndRTB, EM_GETOLEINTERFACE, 0, mIRichEditOle
        If ObjPtr(mIRichEditOle) = 0 Then
            CloseUIInterface
            Exit Sub
        End If
        '获得oleobject的信息
        mReObject.cbStruct = LenB(mReObject)
        mIRichEditOle.GetObject REO_IOB_SELECTION, mReObject, REO_GETOBJ_ALL_INTERFACES

        SendMessage Me.hWndRTB, EM_POSFROMCHAR, VarPtr(pt), ByVal mReObject.cP
        lLeft = pt.x * Screen.TwipsPerPixelX + RTBNormal.Left
        lTOp = pt.y * Screen.TwipsPerPixelY + RTBNormal.Top + lngSpaceBefore * 1.3333333

        picUI.Move IIf(lLeft <= 0, 0, lLeft) - BORDERWIDTH, IIf(lTOp <= 0, 0, lTOp) - BORDERWIDTH
        picUI.Width = lWidth + 2 * BORDERWIDTH
        picUI.Height = lHeight + 2 * BORDERWIDTH
        Call PaintUIBorder
    End If
End Sub

Public Sub CloseUIInterface()
    If picUI.Visible Then
        picUI.Cls
        RaiseEvent UIClose(picUI.hwnd)
        picUI.Visible = False
    End If
End Sub

Public Sub RefreshUIInterface()
    '刷新UI接口容器的位置（修正偏移的情况）
    '显示UI接口容器（当前内容必须是图片）
    If Me.ViewMode <> cprNormal Then Exit Sub
    Dim mIRichEditOle As olelib.IRichEditOle
    Dim mReObject As olelib.REOBJECT
    Dim mIOleObject As olelib.IOleObject
    Dim pSize As olelib.SIZE
    Dim pt As olelib.Point
    picUI.Cls
    If RTBNormal.Selection.GetType = cprSTPicture Then
        Dim lX As Long, lX1 As Long, lY1 As Long, lX2 As Long, lY2 As Long, lTMP As Long
        Dim lLeft As Long, lTOp As Long, lWidth As Long, lHeight As Long
        Dim lngSpaceBefore As Long, lngLinespace As Long
        lngLinespace = RTBNormal.Selection.Para.LineSpacing
        lngSpaceBefore = RTBNormal.Selection.Para.SpaceBefore * Screen.TwipsPerPixelX
        
        '获取精确的高度和宽度（采用OLE接口）
        SendMessage RTBNormal.hWndRTB, EM_GETOLEINTERFACE, 0, mIRichEditOle
        '获得oleobject的信息
        mReObject.cbStruct = LenB(mReObject)
        mIRichEditOle.GetObject REO_IOB_SELECTION, mReObject, REO_GETOBJ_ALL_INTERFACES
        Set mIOleObject = mReObject.poleobj
        If Not mIOleObject Is Nothing Then
            mIOleObject.GetExtent DVASPECT_CONTENT, pSize
            lWidth = UserControl.ScaleX(pSize.cx, vbHimetric, vbTwips)       '图片原始大小
            lHeight = UserControl.ScaleY(pSize.cy, vbHimetric, vbTwips)
        End If
        SendMessage Me.hWndRTB, EM_POSFROMCHAR, VarPtr(pt), ByVal mReObject.cP
        lLeft = pt.x * Screen.TwipsPerPixelX + RTBNormal.Left
        lTOp = pt.y * Screen.TwipsPerPixelY + RTBNormal.Top + lngSpaceBefore * 1.3333333
        '图片最终大小
        lWidth = mReObject.sizel.cx * 192 / 5080 * Screen.TwipsPerPixelX
        lHeight = mReObject.sizel.cy * 192 / 5080 * Screen.TwipsPerPixelY
        
        picUI.Move IIf(lLeft <= 0, 0, lLeft) - BORDERWIDTH, IIf(lTOp <= 0, 0, lTOp) - BORDERWIDTH, lWidth + BORDERWIDTH * 2, lHeight + BORDERWIDTH * 2
        VS.Tag = VS.Value
        Call PaintUIBorder
        Dim LL As Long, lT As Long, lW As Long, lH As Long
        LL = BORDERWIDTH
        lT = BORDERWIDTH
        lW = picUI.Width - 2 * BORDERWIDTH
        lH = picUI.Height - 2 * BORDERWIDTH
        picUI.Width = picUI.Width
        Call PaintUIBorder
    End If
End Sub

Public Sub ShowUIInterface()
    '显示UI接口容器（当前内容必须是图片）
    If Me.ViewMode <> cprNormal Then Exit Sub
    Dim mIRichEditOle As olelib.IRichEditOle
    Dim mReObject As olelib.REOBJECT
    Dim mIOleObject As olelib.IOleObject
    Dim pSize As olelib.SIZE
    Dim pt As olelib.Point
    
    If RTBNormal.Selection.GetType = cprSTPicture Then
'        Me.Range(Me.Selection.StartPos, Me.Selection.StartPos).ScrollIntoView cprSPStart
        
        Dim lX As Long, lX1 As Long, lY1 As Long, lX2 As Long, lY2 As Long, lTMP As Long
        Dim lLeft As Long, lTOp As Long, lWidth As Long, lHeight As Long
        Dim lngSpaceBefore As Long, lngLinespace As Long
        lngLinespace = RTBNormal.Selection.Para.LineSpacing
        lngSpaceBefore = RTBNormal.Selection.Para.SpaceBefore * Screen.TwipsPerPixelX
        
        '获取精确的高度和宽度（采用OLE接口）
        SendMessage RTBNormal.hWndRTB, EM_GETOLEINTERFACE, 0, mIRichEditOle
        If ObjPtr(mIRichEditOle) = 0 Then
            CloseUIInterface
            Exit Sub
        End If
        '获得oleobject的信息
        mReObject.cbStruct = LenB(mReObject)
        mIRichEditOle.GetObject REO_IOB_SELECTION, mReObject, REO_GETOBJ_ALL_INTERFACES
        Set mIOleObject = mReObject.poleobj
        If Not mIOleObject Is Nothing Then
            mIOleObject.GetExtent DVASPECT_CONTENT, pSize
            lWidth = UserControl.ScaleX(pSize.cx, vbHimetric, vbTwips)       '图片原始大小
            lHeight = UserControl.ScaleY(pSize.cy, vbHimetric, vbTwips)
        Else
            CloseUIInterface
            Exit Sub
        End If
        SendMessage Me.hWndRTB, EM_POSFROMCHAR, VarPtr(pt), ByVal mReObject.cP
        lLeft = pt.x * Screen.TwipsPerPixelX + RTBNormal.Left
        lTOp = pt.y * Screen.TwipsPerPixelY + RTBNormal.Top + lngSpaceBefore * 1.3333333
        '图片最终大小
        lWidth = mReObject.sizel.cx * 192 / 5080 * Screen.TwipsPerPixelX
        lHeight = mReObject.sizel.cy * 192 / 5080 * Screen.TwipsPerPixelY
        
        picUI.Move IIf(lLeft <= 0, 0, lLeft) - BORDERWIDTH, IIf(lTOp <= 0, 0, lTOp) - BORDERWIDTH, lWidth + BORDERWIDTH * 2, lHeight + BORDERWIDTH * 2
        VS.Tag = VS.Value
        Call PaintUIBorder
        Dim LL As Long, lT As Long, lW As Long, lH As Long
        LL = BORDERWIDTH
        lT = BORDERWIDTH
        lW = picUI.Width - 2 * BORDERWIDTH
        lH = picUI.Height - 2 * BORDERWIDTH
        RaiseEvent UIOpen(picUI.hwnd, LL, lT, lW, lH)
        picUI.Width = picUI.Width
        Call PaintUIBorder
        picUI.Visible = True
     Else
        CloseUIInterface
    End If
End Sub

Public Sub GetUIBorder(ByRef lLeft As Long, ByRef lTOp As Long, ByRef lWidth As Long, ByRef lHeight As Long)
    lLeft = BORDERWIDTH
    lTOp = BORDERWIDTH
    lWidth = picUI.Width '- 2 * BORDERWIDTH
    lHeight = picUI.Height '- 2 * BORDERWIDTH
End Sub

Public Sub CopyWithFormat()
    '带格式复制
    RTBNormal.CopyWithFormat
End Sub

Public Sub PasteWithFormat()
    '带格式复制
    RTBNormal.PasteWithFormat
End Sub

Public Sub Copy()
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.Copy
    Case cprPaper
        Clipboard.Clear
    End Select
End Sub

Public Sub Cut()
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.Cut
    Case cprPaper
        Clipboard.Clear
    End Select
End Sub

Public Sub Delete()
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.Delete
    Case cprPaper
        '
    End Select
End Sub

Public Function FindText(sText As String, Optional ByVal iFlag As Long) As Boolean
    '功能：从文档当前位置向后查找指定字符串，查到则选中
    '参数：
    '   sText,要查找的字符串
    '   iFlag,匹配方式,默认为0(不区分大小写、全半角)，可以为以下标志的组合：
    '       tomMatchCase,2-大小写匹配
    '       tomMatchWord,4-完全匹配
    '       实际测试，尚不支持模式匹配等
    Select Case mvarViewMode
    Case cprNormal
        FindText = RTBNormal.FindText(sText, iFlag)
    Case cprPaper
        FindText = RTBNormal.FindText(sText, iFlag)
    End Select
End Function

Public Sub Freeze()
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.Freeze
    Case cprPaper
        RTBNormal.Freeze
    End Select
End Sub

Public Function GetLineString(lLine As Long) As String
    Select Case mvarViewMode
    Case cprNormal
        GetLineString = RTBNormal.GetLineString(lLine)
    Case cprPaper
        GetLineString = RTBNormal.GetLineString(lLine)
    End Select
End Function

Public Sub InsertOLEObject()
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.InsertOLEObject
    Case cprPaper
        '
    End Select
End Sub

Public Sub NewDoc()
    If mvarViewMode <> cprNormal Then Exit Sub
    RTBNormal.NewDoc
    RTBHead.NewDoc
    RTBFoot.NewDoc
    SetVSWithRtb True
End Sub

Public Sub OpenDoc(Optional strFile As String = "")
    If mvarViewMode <> cprNormal Then Exit Sub
    Call SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    If Trim(strFile) <> "" Then FileName = strFile
    RTBNormal.OpenDoc strFile
    SetVSWithRtb
End Sub
Public Sub Paste()
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.Paste
    Case cprPaper
        '
    End Select
End Sub

Public Function Range(lStart As Long, lEnd As Long) As cRange
    Set Range = RTBNormal.Range(lStart, lEnd)
End Function

Public Sub Redo()
    RTBNormal.Redo
End Sub

Public Sub SaveDoc(Optional strFile As String = "")
    Screen.MousePointer = vbHourglass
    If Trim(strFile) <> "" Then FileName = strFile
    
    RTBNormal.SaveDoc strFile
    Screen.MousePointer = vbDefault
End Sub
Public Sub SaveHead(ByVal strFile As String)
    Screen.MousePointer = vbHourglass
    If Trim(strFile) = "" Then strFile = HeadFile
    RTBHead.SaveDoc strFile
    Screen.MousePointer = vbDefault
End Sub
Public Sub SaveFoot(ByVal strFile As String)
    Screen.MousePointer = vbHourglass
    If Trim(strFile) = "" Then strFile = FootFile
    RTBFoot.SaveDoc strFile
    Screen.MousePointer = vbDefault
End Sub
Public Sub SelectAll()
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.SelectAll
    Case cprPaper
        RTBNormal.SelectAll
    End Select
End Sub

Public Function Selection() As cSelection
    Set Selection = RTBNormal.Selection
End Function

Public Sub Undo()
    RTBNormal.Undo
End Sub

Public Sub UnFreeze()
    Select Case mvarViewMode
    Case cprNormal
        RTBNormal.UnFreeze
    Case cprPaper
        RTBNormal.UnFreeze
    End Select
End Sub
Public Function InsertPicture2(objPic As StdPicture, Optional ByRef lWidth As Long, Optional ByRef lHeight As Long, Optional ByVal lStart As Long = -1, Optional ByVal lEnd As Long = -1) As Long
'插入图片到指定位置
    Dim aStr As String, blnForce As Boolean
    If mvarViewMode = cprNormal Then

        If lStart < 0 Then lStart = RTBNormal.Selection.StartPos
        If lEnd < 0 Then lEnd = RTBNormal.Selection.EndPos
        aStr = StdPicAsRTF(objPic, lWidth, lHeight)
        blnForce = RTBNormal.ForceEdit
        RTBNormal.ForceEdit = True
        RTBNormal.Range(lStart, lEnd).Selected
        RTBNormal.SelRTF = aStr
        InsertPicture2 = lStart
        RTBNormal.ForceEdit = blnForce
    End If
End Function

Public Function InsertPicture(objPic As StdPicture, Optional ByRef lWidth As Long, Optional ByRef lHeight As Long, Optional ByVal lStart As Long = -1, Optional ByVal lEnd As Long = -1) As Long
'插入图片到指定位置
    Dim aStr As String, blnForce As Boolean
    If mvarViewMode = cprNormal Then
        
        If lStart < 0 Then lStart = RTBNormal.Selection.StartPos
        If lEnd < 0 Then lEnd = RTBNormal.Selection.EndPos
    
        Call CloseClipboard
        DoEvents
        Clipboard.Clear
        Clipboard.SetData objPic
        
        rtbBuff.Text = ""
        SendMessageLong rtbBuff.hwnd, WM_PASTE, 0, 0
        ResizeReObject rtbBuff, lWidth, lHeight     '调节图片尺寸
        
        Call CloseClipboard
        DoEvents
        Clipboard.Clear
        SetSelection rtbBuff.hwnd, 0, 1
        SendMessageLong rtbBuff.hwnd, WM_COPY, 0, 0
        rtbBuff.Text = ""
        
        blnForce = RTBNormal.ForceEdit
        RTBNormal.ForceEdit = True
        RTBNormal.Range(lStart, lEnd).Font.Protected = False
        RTBNormal.Range(lStart, lEnd).Selected
        RTBNormal.PasteWithFormat
        
        InsertPicture = lStart
        RTBNormal.ForceEdit = blnForce
    End If
End Function

'##################################     内嵌对话框      ##################################

Public Function ShowFontDlg(Optional intFlags As Integer) As Boolean
    '功能：显示字体对话框，可以改变字体,字号,粗体,斜体，根据参数决定是否处理字体效果相关属性；只有普通模式允许
    '参数：
    '   intFlags,是否禁止相关的附加效果选项：
    '       intFlags and (2^0) <> 0,禁止更改删除线属性
    '       intFlags and (2^1) <> 0,禁止更改保护属性
    '       intFlags and (2^2) <> 0,禁止更改隐藏属性
    '       intFlags and (2^3) <> 0,禁止更改下划线属性
    '       intFlags and (2^4) <> 0,禁止更改前景色属性
    '       intFlags and (2^5) <> 0,禁止更改背景色属性
    
    Dim strSample As String
    
    If Me.ViewMode <> cprNormal Then Exit Function
    strSample = Trim(Me.Selection.Text)
    If strSample <> "" Then strSample = Left(Split(strSample, vbCrLf)(0), 10)
    If strSample <> Trim(Me.Selection.Text) Then strSample = strSample & "…"
    
    Me.ForceEdit = True
    ShowFontDlg = frmFontSetup.ShowMe(TOM, intFlags, strSample)
    Me.ForceEdit = False
End Function

Public Function ShowPageSetupDlg(Optional intFlags As Integer) As Boolean
    '功能：显示页面设置对话框
    '参数：
    '   intFlags,是否禁止相关的附加效果选项：
    '       intFlags and (2^0) <> 0,禁止更改页面背景色属性
    '       intFlags and (2^1) <> 0,禁止更改文档背景色属性
    
    If frmPageSetup.ShowMe(Me, intFlags) Then
        ShowPageSetupDlg = True
        Call UserControl_Resize
        
'        '普通模式下所见即所得的重新设置
'        RTBNormal.ResetWYSIWYG
'
'        '页面模式的话需要重新分页
        If mvarViewMode = cprPaper Then ViewMode = cprPaper
        Me.Modified = True
    End If

End Function

Public Function ShowParaDlg(Optional blnHideStyle As Boolean) As Boolean
    '功能：显示段落格式对话框；只有普通模式允许。
    '参数：blnHideStyle-是否禁止大纲样式设置
    
    Dim strText As String, strSample As String, lS As Long, lE As Long, i As Long
    
    If Me.ViewMode <> cprNormal Then Exit Function
        
    '获取段落文字，以便作为示范
    strText = Me.Text
    strSample = ""
    lS = InStrRev(strText, vbCrLf, Me.SelStart + 1) - 1
    lS = IIf(lS <= 0, 0, lS)
    lE = InStr(lS + 1, strText, vbCrLf, vbTextCompare) - 2
    For i = lS To lE
        If Me.Range(i, i + 1).Font.Hidden Then
            i = i + 1
        Else
            strSample = strSample & Me.Range(i, i + 1)
        End If
    Next
    
    ShowParaDlg = frmParagraph.ShowMe(Me.Selection.Para, blnHideStyle, strSample)
    
End Function

Public Function ShowItemNumberDlg() As Boolean
    '功能：显示项目符号和编号对话框
    ShowItemNumberDlg = frmItemNumber.ShowMe(Me.Selection.Para)
End Function

Public Function ShowCharCountDlg() As Boolean
    '功能：显示字数统计对话框
    If mvarViewMode <> cprPaper Then
        Me.InProcessing = True
        DoVirtualPrint
        Me.InProcessing = False
    End If
    ShowCharCountDlg = frmCharCount.ShowMe(Me)
End Function

Public Function ShowInsertDateTimeDlg(Optional blnDelay As Boolean, _
    Optional MinDate As Date, _
    Optional MaxDate As Date, _
    Optional bSaveInEditor As Boolean = True) As String
    '功能：显示插入日期时间对话框
    '参数：
    '   blnDelay,为真时，不直接插入修改编辑器的SelText内容，只返回设置值。
    '   MinDate,允许的最小日期
    '   MaxDate,允许的最大日期
    '返回：设置的日期时间字符串，取消时返回""
    If bSaveInEditor Then
        If Me.Selection.Font.Protected Then ShowInsertDateTimeDlg = False: Exit Function
    End If
    
    Dim strReturn As String
    strReturn = frmDateTime.ShowMe(MinDate, MaxDate)
    ShowInsertDateTimeDlg = strReturn
    If bSaveInEditor Then
        If blnDelay = True Then Exit Function
        If strReturn = "" Then Exit Function
        If Me.AuditMode Then
            Range(Selection.EndPos, Selection.EndPos).Selected
            '保留性属性（便于新增文本）
            ForceEdit = True
            On Error Resume Next
            OriginRTB.SelColor = OriginRTB.GetNewCharColor(tomAutoColor)
    '        OriginRTB.SelUnderline = True
            OriginRTB.SelStrikeThru = False
            ForceEdit = False
        End If
        Me.ForceEdit = True
        Me.SelText = strReturn
        Me.SelStart = Me.SelStart + Len(strReturn)
        Err.Clear
    End If
End Function

Public Function ShowInsertSymbolDlg(ByVal bSaveInEditor As Boolean, ByVal bytSex As Byte, _
                                    ByVal blnReturnStr As Boolean, strInfor As String, objPic As StdPicture) As String
    '功能：显示插入符号和特殊字符对话框
    '参数：
    '   bSaveInEditor,为真时， 直接插入/修改编辑器的SelText内容，否则返回设置值。
    '   bytSex,性别，0-没指定;1-男性;2-女性
    '   blnReturnStr 是否以字符方式返回,表示当前位置不支持图片方式返回.=true 用字符返回
    '   strInfor 编辑图片时传入的文字信息，编辑完后回传
    '            形式为：类型|数据。月经史 1|前辍|分子|分母|后辍|字号; 牙齿 2(恒牙)/3(乳牙)|左上|右上|左下|右下|字号; 胎心位置 4|上方|下方|左方|右方|字号
    '   objPic   编辑窗口生成的图片回传
    '返回：只要执行过插入则返回True，直接关闭返回False
    If bSaveInEditor Then
        If Me.Selection.Font.Protected Then ShowInsertSymbolDlg = False: Exit Function
    End If
    
    Dim strReturn As String, COLOR As OLE_COLOR, lFontSize As Long
    If strInfor <> "" And UBound(Split(strInfor, "|")) > 0 Then
        lFontSize = Val(Split(strInfor, "|")(5))
    Else
        lFontSize = Me.Range(Selection.EndPos, Selection.EndPos).Font.SIZE
    End If
    strReturn = frmInsSymbol.ShowMe(bytSex, blnReturnStr, strInfor, objPic, lFontSize)
    Unload frmInsSymbol
    ShowInsertSymbolDlg = strReturn
    
    If bSaveInEditor And objPic Is Nothing Then
        If Me.AuditMode Then
            Range(Selection.EndPos, Selection.EndPos).Selected
            '保留性属性（便于新增文本）
            ForceEdit = True
            COLOR = vbBlack
            RaiseEvent GetNewCharColor(COLOR)
            OriginRTB.SelColor = COLOR
    '        OriginRTB.SelUnderline = True
            OriginRTB.SelStrikeThru = False
            ForceEdit = False
        End If
        Me.ForceEdit = True
        Me.SelText = strReturn
        Me.SelStart = Me.SelStart + Len(strReturn)
    End If
End Function

Public Function ShowHeadFootDlg() As Boolean
    '功能：显示页眉页脚对话框
    ShowHeadFootDlg = frmHeadFoot.ShowMe(Me)
End Function

Public Function ShowFindReplaceDlg(Optional intShowWhat As Integer) As Boolean
    '功能：显示查找替换对话框，执行查找替换；替换时，不对保护和隐藏的内容进行替换；页面模式不提供。
    '参数：
    '   intShowWhat,显示和禁止的功能:
    '    0,首先显示查找处理
    '    1,首先显示替换处理
    '   -1,显示查找处理并屏蔽替换处理
    If mvarViewMode <> cprNormal Then Exit Function
    ShowFindReplaceDlg = mfrmFindText.ShowMe(Me, intShowWhat)
End Function

Public Sub FindNext()
    '功能： 查找一下个
    If mvarViewMode <> cprNormal Then Exit Sub
    mfrmFindText.FindNext Me
End Sub

Private Sub btnNormal_Click()
    ViewMode = cprNormal
End Sub

Private Sub btnPaper_Click()
    ViewMode = cprPaper
End Sub

Private Sub HRuler_IndentChanged(LeftIndent As Long, FirstLineIndent As Long, RightIndent As Long)
    If mvarViewMode = cprNormal And Me.AuditMode = False Then
        Err = 0: On Error Resume Next
        Dim W As Long
        Const LIMITWIDTH = 3000
        W = (mvarPaperWidth - mvarMarginLeft - mvarMarginRight - LIMITWIDTH) * mvarZoomFactor
        
        '不能超出范围
        If LeftIndent < 0 Then LeftIndent = 0
        If LeftIndent > W Then LeftIndent = W
        
        If FirstLineIndent < 0 Then
            If Abs(FirstLineIndent) > LeftIndent Then FirstLineIndent = -LeftIndent
        Else
            If FirstLineIndent + LeftIndent > W Then FirstLineIndent = W - LeftIndent
        End If
        
        If RightIndent < 0 Then RightIndent = 0
        If RightIndent > W Then RightIndent = W
        
        If RTBNormal.Selection.Font.Protected = False Then
            RTBNormal.Selection.Para.SetIndents FirstLineIndent / 20, LeftIndent / 20, RightIndent / 20
        End If
        Call RTBNormal_SelChange(RTBNormal.Selection.StartPos, RTBNormal.Selection.EndPos)
        If RTBNormal.Enabled And RTBNormal.Visible Then
            RTBNormal.SetFocus
        End If
    End If
    Err.Clear
End Sub

Private Sub HRuler_TabStopChanged(TabCount As Integer, TabPos() As Long, TabAlign() As Byte)
    Err = 0: On Error Resume Next
    If HRuler.Tag <> "" Then Exit Sub
    Dim i As Long, j As Long, k As Long, lS As Long, lE As Long, strText As String, lCur As Long
    If mvarViewMode = cprNormal And Me.AuditMode = False Then
        RTBNormal.ForceEdit = True
        With RTBNormal.TOM.TextDocument.Selection.Para
            If .TabCount = tomUndefined Then
                '选中多个段落
                lS = RTBNormal.Selection.StartPos
                lE = RTBNormal.Selection.EndPos
                strText = RTBNormal.Text
                For i = lS To lE
                    j = InStr(i + 1, strText, vbCrLf)
                    If j = 0 Then
                        '没有发现回车
                        Exit For
                    ElseIf j <= lE Then
                        '范围内发现回车
                        i = j + 1
                        lCur = j - 1
                        RTBNormal.TOM.TextDocument.Range(lCur, lCur).Para.ClearAllTabs
                        For k = 0 To TabCount - 1
                            If TabPos(k) > 0 Then
                                RTBNormal.TOM.TextDocument.Range(lCur, lCur).Para.AddTab TabPos(k) / 20, TabAlign(k), tomSpaces
                            End If
                        Next k
                    Else
                        '范围内没有发现回车，取最末位置设置制表位
                        lCur = lE
                        RTBNormal.TOM.TextDocument.Range(lCur, lCur).Para.ClearAllTabs
                        For k = 0 To TabCount - 1
                            If TabPos(k) > 0 Then
                                RTBNormal.TOM.TextDocument.Range(lCur, lCur).Para.AddTab TabPos(k) / 20, TabAlign(k), tomSpaces
                            End If
                        Next k
                        Exit For
                    End If
                Next
            Else
                '选中单个段落
                .ClearAllTabs
                For i = 0 To TabCount - 1
                    If TabPos(i) > 0 Then .AddTab TabPos(i) / 20, TabAlign(i), tomSpaces
                Next i
            End If
        End With
        RTBNormal.ForceEdit = False
        If RTBNormal.Visible And RTBNormal.Enabled Then
            RTBNormal.SetFocus
        End If
    End If
    Err.Clear
End Sub

Private Sub PaintUIBorder()
    Dim i As Long, j As Long
    picUI.Cls
    For i = 0 To picUI.ScaleWidth Step picBorder.Width
        picUI.PaintPicture picBorder.Picture, i, 0, picBorder.Width, picBorder.Height
        picUI.PaintPicture picBorder.Picture, i, picUI.ScaleHeight - picBorder.Height, picBorder.Width, picBorder.Height
    Next
    For i = 0 To picUI.ScaleHeight Step picBorder.Height
        picUI.PaintPicture picBorder.Picture, 0, i, picBorder.Width, picBorder.Height
        picUI.PaintPicture picBorder.Picture, picUI.ScaleWidth - picBorder.Width, i, picBorder.Width, picBorder.Height
    Next
End Sub


Private Sub picMarginR_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mvarViewMode <> cprNormal Then Exit Sub
    
    Dim R1 As POINTAPI, R2 As POINTAPI
    GetCursorPos R1  '获取当前鼠标位置
    R2.x = R1.x - x / Screen.TwipsPerPixelX - 1
    R2.y = R1.y
    SetCursorPos R2.x, R2.y
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
    SetCursorPos R1.x, R1.y
End Sub

Private Sub picUI_Click()
    RaiseEvent UIClick(mvarViewMode)
End Sub

Private Sub picMarginL_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mvarViewMode <> cprNormal Then Exit Sub
    If picMarginL.Tag = "" Then
        Dim R1 As POINTAPI, R2 As POINTAPI
        GetCursorPos R1  '获取当前鼠标位置
        R2.x = R1.x + (picMarginL.Width - x) / Screen.TwipsPerPixelX
        R2.y = R1.y
        SetCursorPos R2.x, R2.y
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
        Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
        SetCursorPos R1.x, R1.y
        picMarginL.Tag = "Down"
    End If
End Sub

Private Sub picMarginL_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picMarginL.Tag = ""
End Sub

Private Sub RTBNormal_BeforeKeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent BeforeKeyDown(mvarViewMode, KeyCode, Shift)
End Sub

Private Sub RTBNormal_Change()
    RaiseEvent Change(mvarViewMode)
End Sub

Private Sub RTBNormal_Click()
    RaiseEvent Click(mvarViewMode)
End Sub

Private Sub RTBNormal_DblClick()
    RaiseEvent DblClick(mvarViewMode)
End Sub

Private Sub RTBNormal_Focuse()
    CloseUIInterface
End Sub

Private Sub RTBNormal_GetDelCharColor(COLOR As OLE_COLOR)
    RaiseEvent GetDelCharColor(COLOR)
End Sub

Private Sub RTBNormal_GetNewCharColor(COLOR As OLE_COLOR)
    RaiseEvent GetNewCharColor(COLOR)
End Sub

Private Sub RTBNormal_IsDelCharColor(ByVal COLOR As OLE_COLOR, blnIsDelCharColor As Boolean)
    RaiseEvent IsDelCharColor(COLOR, blnIsDelCharColor)
End Sub

Private Sub RTBNormal_IsNewCharColor(ByVal COLOR As OLE_COLOR, blnIsNewCharColor As Boolean)
    RaiseEvent IsNewCharColor(COLOR, blnIsNewCharColor)
End Sub

Private Sub RTBNormal_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(mvarViewMode, KeyCode, Shift)
End Sub

Private Sub RTBNormal_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(mvarViewMode, KeyAscii)
End Sub

Private Sub RTBNormal_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(mvarViewMode, KeyCode, Shift)
End Sub

Private Sub RTBNormal_LinkEvent(ByVal iType As LinkEventTypeEnum, ByVal lStart As Long, ByVal lEnd As Long)
    RaiseEvent LinkEvent(mvarViewMode, iType, lStart, lEnd)
End Sub

Private Sub RTBNormal_ModifyProtected(ByRef bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)
    RaiseEvent ModifyProtected(mvarViewMode, bAllowDoIt, lStart, lEnd, KeyAscii, Shift)
End Sub

Private Sub RTBNormal_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseDown(mvarViewMode, Button, Shift, x, y)
End Sub

Private Sub RTBNormal_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseMove(mvarViewMode, Button, Shift, x, y)
End Sub

Private Sub RTBNormal_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseUp(mvarViewMode, Button, Shift, x, y)
End Sub

Private Sub RTBNormal_MouseWheel(bBackDirection As Boolean, Shift As Integer, x As Single, y As Single, Value As Single)
    SetVSWithRtb
    RaiseEvent MouseWheel(mvarViewMode, bBackDirection, Shift, x, y, Value)
    CloseUIInterface
End Sub

Private Sub RTBNormal_PressTabKey()
    RaiseEvent PressTabKey
End Sub

Private Sub RTBNormal_RequestLine()
    RaiseEvent RequestLine(mvarViewMode)
End Sub

Private Sub RTBNormal_RequestRightMenu(ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent RequestRightMenu(mvarViewMode, Shift, x, y)
End Sub

Private Sub RTBNormal_SelChange(ByVal lStart As Long, ByVal lEnd As Long)
'    SetVSWithRtb
    
    RaiseEvent SelChange(mvarViewMode, lStart, lEnd)
    Dim lF As Long, LL As Long, lR As Long
    lF = RTBNormal.Selection.Para.FirstLineIndent
    LL = RTBNormal.Selection.Para.LeftIndent
    lR = RTBNormal.Selection.Para.RightIndent
    If lF = tomUndefined Then lF = 0
    If LL = tomUndefined Then LL = 0
    If lR = tomUndefined Then lR = 0
    HRuler.FirstLineIndent = lF * 20        '磅值与缇进度为20。
    HRuler.LeftIndent = LL * 20
    HRuler.RightIndent = lR * 20

    Dim i As Long, j As Long
    Dim iT As Single, lA As Long, lLd As Long
    Dim iTabPos() As Long, lAlign() As Byte, lLeader() As Long
    j = RTBNormal.Selection.Para.TabCount

    If j = tomUndefined Then j = 0
    ReDim iTabPos(0 To j) As Long
    ReDim lAlign(0 To j) As Byte
    ReDim lLeader(0 To j) As Long
    HRuler.Tag = "Editing"
    For i = 0 To j - 1
        RTBNormal.TOM.TextDocument.Selection.Para.GetTab i, iT, lA, LL
        iTabPos(i) = iT * 20
        lAlign(i) = lA * 20
        lLeader(i) = lLd * 20
    Next
    HRuler.SetTabs CInt(j), iTabPos, lAlign
    HRuler.Tag = ""
End Sub

Private Sub RTBNormal_Zoom(NewFactor As Double)
    mvarZoomFactor = NewFactor
    Call ResetWYSIWYG
    RaiseEvent Zoom(mvarViewMode, NewFactor)
End Sub

Private Sub RTBPaper_Change(Index As Integer)
    RaiseEvent Change(mvarViewMode)
End Sub

Private Sub RTBPaper_Click(Index As Integer)
    RaiseEvent Click(mvarViewMode)
    mvarCurPage = Index
End Sub

Private Sub RTBPaper_DblClick(Index As Integer)
    RaiseEvent DblClick(mvarViewMode)
End Sub

Private Sub RTBPaper_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(mvarViewMode, KeyCode, Shift)
End Sub

Private Sub RTBPaper_KeyPress(Index As Integer, KeyAscii As Integer)
    RaiseEvent KeyPress(mvarViewMode, KeyAscii)
End Sub

Private Sub RTBPaper_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(mvarViewMode, KeyCode, Shift)
End Sub

Private Sub RTBPaper_LinkEvent(Index As Integer, ByVal iType As LinkEventTypeEnum, ByVal lStart As Long, ByVal lEnd As Long)
    RaiseEvent LinkEvent(mvarViewMode, iType, lStart, lEnd)
End Sub

Private Sub RTBPaper_LostFocus(Index As Integer)
    '检测Tab键！
    Dim iRetVal As Integer
    iRetVal = GetKeyState(VK_SHIFT)
    ' 如果没有按shift，检查tab
    If iRetVal <> -128 And iRetVal <> -127 Then
        iRetVal = GetKeyState(VK_TAB)
        If iRetVal = -128 Or iRetVal = -127 Then ' tab键按下
            If RTBPaper(Index).Visible And RTBPaper(Index).Enabled Then
                RTBPaper(Index).SetFocus
            End If
        End If
    End If
    RTBPaper(Index).Tag = ""
End Sub

Private Sub RTBPaper_ModifyProtected(Index As Integer, bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)
    RaiseEvent ModifyProtected(mvarViewMode, bAllowDoIt, lStart, lEnd, KeyAscii, Shift)
End Sub

Private Sub RTBPaper_MouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    mvarCurPage = Index
    If RTBPaper(Index).Tag = "" Then
        Dim R1 As POINTAPI, R2 As POINTAPI
        GetCursorPos R1  '获取当前鼠标位置
        If x <= mvarMarginLeft * mvarZoomFactor Then
            R2.x = R1.x + (mvarMarginLeft * mvarZoomFactor - x) / Screen.TwipsPerPixelX
            If y <= mvarMarginTop * mvarZoomFactor Then
                '超出上边距
                R2.y = R1.y + (mvarMarginTop * mvarZoomFactor - y) / Screen.TwipsPerPixelY
            ElseIf y >= (mvarPaperHeight - mvarMarginBottom) * mvarZoomFactor Then
                '超出下边距
                R2.y = R1.y + ((mvarPaperHeight - mvarMarginBottom) * mvarZoomFactor - y) / Screen.TwipsPerPixelY
            Else
                R2.y = R1.y
            End If
            SetCursorPos R2.x, R2.y
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
        Else
            If y <= mvarMarginTop * mvarZoomFactor Then
                '超出上边距
                R2.x = R1.x
                R2.y = R1.y + (mvarMarginTop * mvarZoomFactor - y) / Screen.TwipsPerPixelY
            ElseIf y >= (mvarPaperHeight - mvarMarginBottom) * mvarZoomFactor Then
                '超出下边距
                R2.x = R1.x
                R2.y = R1.y + ((mvarPaperHeight - mvarMarginBottom) * mvarZoomFactor - y) / Screen.TwipsPerPixelY
            Else
                R2.x = R1.x + (-x + (mvarPaperWidth - mvarMarginRight) * mvarZoomFactor - 1) / Screen.TwipsPerPixelX
                R2.y = R1.y
            End If
            SetCursorPos R2.x, R2.y
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
        End If
        SetCursorPos R1.x, R1.y
        RTBPaper(Index).Tag = "Down"
    End If
End Sub

Private Sub RTBPaper_MouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RTBPaper(Index).Tag = ""
End Sub

Private Sub RTBPaper_MouseWheel(Index As Integer, bBackDirection As Boolean, Shift As Integer, x As Single, y As Single, Value As Single)
    If mvarViewMode = cprNormal Then Exit Sub
    RaiseEvent MouseWheel(mvarViewMode, bBackDirection, Shift, x, y, Value)
    If VS.Value - IIf(Value < 0, -1, 1) * WHEELNUMBER > VS.Max Then
        VS.Value = VS.Max
    Else
        If VS.Value - IIf(Value < 0, -1, 1) * WHEELNUMBER > 0 Then
            VS.Value = VS.Value - IIf(Value < 0, -1, 1) * WHEELNUMBER
        Else
            VS.Value = 0
        End If
    End If
End Sub

Private Sub RTBPaper_RequestLine(Index As Integer)
    RaiseEvent RequestLine(mvarViewMode)
End Sub

Private Sub RTBPaper_RequestRightMenu(Index As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent RequestRightMenu(mvarViewMode, Shift, x, y)
End Sub

Private Sub RTBPaper_Resize(Index As Integer)
    picShadow(Index).Move RTBPaper(Index).Left + SHADOWOFFSET, RTBPaper(Index).Top + SHADOWOFFSET, RTBPaper(Index).Width, RTBPaper(Index).Height
End Sub

Private Sub RTBPaper_RTBMouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseDown(mvarViewMode, Button, Shift, x, y)
End Sub

Private Sub RTBPaper_RTBMouseMove(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseMove(mvarViewMode, Button, Shift, x, y)
End Sub

Private Sub RTBPaper_RTBMouseUp(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    RaiseEvent MouseUp(mvarViewMode, Button, Shift, x, y)
End Sub

Private Sub RTBPaper_SelChange(Index As Integer, ByVal lStart As Long, ByVal lEnd As Long)
    RaiseEvent SelChange(mvarViewMode, lStart, lEnd)
End Sub

Private Sub RTBPaper_Zoom(Index As Integer, NewFactor As Double)
    RaiseEvent Zoom(mvarViewMode, NewFactor)
End Sub

Private Sub UserControl_GotFocus()
    Dim lngTargetDC As Long
    If Not ExistsPrinter Then
        lngTargetDC = picBuff.hDC
    Else
        lngTargetDC = Printer.hDC
    End If
    gTargetDC = picBuff.hDC     '解决民航医院预览时右边超出的问题！（只有以屏幕为度量）
    If lngTargetDC <> gTargetDC Then ResetWYSIWYG
End Sub

Private Sub UserControl_Initialize()
'在程序创建控件及运行时时发生                                           '换算单位56.6857142857143

    '按名称、高度、宽度、最小边距(上下左右)、对应打印纸张排列的纸张种类常量
    PaperKindConst(1) = "信笺 8 1/2×11 英寸                        ,15842,12242,482,799,181,181,1"
    PaperKindConst(2) = "+A611 小型信笺 8 1/2×11 英寸              ,15842,12242,482,799,181,181,2"
    PaperKindConst(3) = "小型报 11×17 英寸                         ,24477,15842,482,799,181,181,3"
    PaperKindConst(4) = "分类帐 17×11 英寸                         ,15842,24477,482,799,181,181,4"
    PaperKindConst(5) = "法律文件 8 1/2×14 英寸                    ,20163,12242,482,799,181,181,5"
    PaperKindConst(6) = "声明书5 1/2×8 1/2 英寸                    ,12242,7919,482,799,181,181,6"
    PaperKindConst(7) = "行政文件7 1/2×10 1/2 英寸                 ,15122,10438,482,799,181,181,7"
    PaperKindConst(8) = "A3 297×420 毫米                           ,23814,16840,482,799,181,193,8"
    PaperKindConst(9) = "A4 210×297 毫米                           ,16840,11907,482,805,181,176,9"
    PaperKindConst(10) = "A4小号 210×297 毫米                      ,16840,11907,482,805,181,176,9"
    PaperKindConst(11) = "A5 148×210 毫米                          ,11907,8392,482,799,181,176,11"
    PaperKindConst(12) = "B4 250×354 毫米                          ,20067,14171,482,805,181,181,12"
    PaperKindConst(13) = "B5 182×257 毫米                          ,14572,10319,482,805,181,176,13"
    PaperKindConst(14) = "对开本 8 1/2×13 英寸                     ,18722,12242,482,799,181,181,14"
    PaperKindConst(15) = "四开本 215×275 毫米                      ,15589,12187,482,805,181,176,15"
    PaperKindConst(16) = "10×14 英寸                               ,20163,14398,482,805,181,176,16"
    PaperKindConst(17) = "11×17 英寸                               ,24477,15842,482,805,181,176,17"
    PaperKindConst(18) = "便条8 1/2×11 英寸                        ,15842,12242,482,805,181,176,18"
    PaperKindConst(19) = "#9 信封 3 7/8×8 7/8 英寸                 ,5579,12780,482,794,181,176,19"
    PaperKindConst(20) = "#10 信封 4 1/8×9 1/2 英寸                ,5936,13682,482,794,181,181,20"
    PaperKindConst(21) = "#11 信封 4 1/2×10 3/8 英寸               ,14938,6479,482,794,181,181,21"
    PaperKindConst(22) = "#12 信封 4 1/2×11 英寸                   ,15842,6479,482,794,181,181,22"
    PaperKindConst(23) = "#14 信封 5×11 1/2 英寸                   ,16558,7199,482,794,181,181,23"
    PaperKindConst(24) = "C 尺寸工作单                              ,16558,7199,482,794,181,181,24"
    PaperKindConst(25) = "D 尺寸工作单                              ,16558,7199,482,794,181,181,25"
    PaperKindConst(26) = "E 尺寸工作单                              ,16558,7199,482,794,181,181,26"
    PaperKindConst(27) = "DL 型信封 110×220 毫米                   ,6237,12474,482,805,181,181,27"
    PaperKindConst(28) = "C5 型信封 162×229 毫米                   ,9185,12984,482,799,181,176,28"
    PaperKindConst(29) = "C3 型信封 324×458 毫米                   ,25969,18371,482,794,181,176,29"
    PaperKindConst(30) = "C4 型信封 229×324 毫米                   ,18371,12981,482,794,181,176,30"
    PaperKindConst(31) = "C6 型信封 114×162 毫米                   ,9183,6462,482,794,181,176,31"
    PaperKindConst(32) = "C65 型信封114×229 毫米                   ,12981,6462,482,794,181,176,32"
    PaperKindConst(33) = "B4 型信封 250×353 毫米                   ,20010,14171,482,794,181,176,33"
    PaperKindConst(34) = "B5 型信封176×250 毫米                    ,9979,14175,482,799,181,193,34"
    PaperKindConst(35) = "B6 型信封 176×125 毫米                   ,7086,9976,482,799,181,193,35"
    PaperKindConst(36) = "信封 110×230 毫米                        ,13037,6237,482,799,181,193,36"
    PaperKindConst(37) = "信封大王 3 7/8×7 1/2 英寸                ,5579,10801,482,794,181,181,37"
    PaperKindConst(38) = "信封 3 5/8×6 1/2 英寸                    ,9359,5219,482,794,181,181,38"
    PaperKindConst(39) = "U.S. 标准复写簿 14 7/8×11 英寸           ,15842,21421,0,0,0,1848,39"
    PaperKindConst(40) = "德国标准复写簿 8 1/2×12 英寸             ,17282,12242,0,0,0,0,40"
    PaperKindConst(41) = "德国法律复写簿 8 1/2×13 英寸             ,18722,12242,0,0,0,0,41"
    PaperKindConst(42) = "自定义纸张                                ,22680,16443,482,0,0,0,256"
    PageCount = 1
    mvarCurPage = 1

    lblThis.Caption = "中联图文编辑控件 v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & App.LegalCopyright & " " & App.LegalTrademarks
End Sub

Private Sub UserControl_InitProperties()
'创建对象新实例时发生，即新属性的最初初始化代码！（即，当用户在窗体上放置一个控件时触发此事件！运行时不再触发！）
    AutoDetectURL = True
    BackColor = &H99A8AC
    PaperColor = vbWhite
    Border = False
    DefaultTabStop = 21
    DoDefaultURLClick = False
    Enabled = True
    FileName = ""
    ForceEdit = False
    Modified = False
    ReadOnly = False
    Text = ""
    Title = "未命名文档"
    ZoomFactor = 1#
    Foot = ""
    Head = ""
    MarginTop = 1400
    MarginBottom = 1400
    MarginLeft = 1800
    MarginRight = 1800
    PaperHeight = 16840
    PaperWidth = 11907
    Transparent = False
    ShowPageNumber = True
    PageCount = 1
    CurPage = 1
    ViewMode = cprNormal
    WithViewButtonas = True
    PaperKind = cprPKA4
    PaperOrient = cprPOPortrait
    ShowRuler = True
    AuditMode = False
    HeadFontName = "宋体"
    HeadFontSize = 10
    HeadFontBold = False
    HeadFontItalic = False
    HeadFontUnderline = False
    HeadFontStrikethrough = False
    HeadFontColor = vbBlack
    HeadFile = ""
    FootFontName = "宋体"
    FootFontSize = 10
    FootFontBold = False
    FootFontItalic = False
    FootFontUnderline = False
    FootFontStrikethrough = False
    FootFontColor = vbBlack
    FootFile = ""
End Sub

Private Sub UserControl_Paint()
    If Not UserControl.Extender.Visible Then
        Exit Sub
    End If
    If Ambient.UserMode Then
        imgIcon.Visible = False
        lblThis.Visible = False
    Else
        imgIcon.Visible = True
        lblThis.Visible = True
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'当加载具有保存状态的对象的旧实例时，发生该事件。
'属性读取（静态属性的读取，从而转化为动态属性，此时调用pInitialise函数初始化句柄！）
    AutoDetectURL = PropBag.ReadProperty("AutoDetectURL", True)
    BackColor = PropBag.ReadProperty("BackColor", &H99A8AC)
    PaperColor = PropBag.ReadProperty("PaperColor", vbWhite)
    Border = PropBag.ReadProperty("Border", False)
    DefaultTabStop = PropBag.ReadProperty("DefaultTabStop", Me.TOM.TextDocument.Selection.Font.SIZE * 2)
    DoDefaultURLClick = PropBag.ReadProperty("DoDefaultURLClick", False)
    Enabled = PropBag.ReadProperty("Enabled", True)
    FileName = PropBag.ReadProperty("FileName", "")
    ForceEdit = PropBag.ReadProperty("ForceEdit", False)
    ReadOnly = PropBag.ReadProperty("ReadOnly", False)
    Title = PropBag.ReadProperty("Title", "未命名文档")
    ZoomFactor = PropBag.ReadProperty("ZoomFactor", 1#)
    Foot = PropBag.ReadProperty("Foot", "")
    Head = PropBag.ReadProperty("Head", "")
    PaperHeight = PropBag.ReadProperty("PaperHeight", 16840)
    PaperWidth = PropBag.ReadProperty("PaperWidth", 11907)
    MarginTop = PropBag.ReadProperty("MarginTop", 1400)
    MarginBottom = PropBag.ReadProperty("MarginBottom", 1400)
    MarginLeft = PropBag.ReadProperty("MarginLeft", 1800)
    MarginRight = PropBag.ReadProperty("MarginRight", 1800)
    Transparent = PropBag.ReadProperty("Transparent", False)
    ShowPageNumber = PropBag.ReadProperty("ShowPageNumber", True)
    PageCount = PropBag.ReadProperty("PageCount", 1)
    CurPage = PropBag.ReadProperty("CurPage", 1)
    ViewMode = PropBag.ReadProperty("ViewMode", cprNormal)
    WithViewButtonas = PropBag.ReadProperty("WithViewButtonas", True)
    PaperKind = PropBag.ReadProperty("PaperKind", cprPKA4)
    PaperOrient = PropBag.ReadProperty("PaperOrient", cprPOPortrait)
    ShowRuler = PropBag.ReadProperty("ShowRuler", True)
    AuditMode = PropBag.ReadProperty("AuditMode", False)
    HeadFontName = PropBag.ReadProperty("HeadFontName", "宋体")
    HeadFontSize = PropBag.ReadProperty("HeadFontSize", 10)
    HeadFontBold = PropBag.ReadProperty("HeadFontBold", False)
    HeadFontItalic = PropBag.ReadProperty("HeadFontItalic", False)
    HeadFontUnderline = PropBag.ReadProperty("HeadFontUnderline", False)
    HeadFontStrikethrough = PropBag.ReadProperty("HeadFontStrikethrough", False)
    HeadFontColor = PropBag.ReadProperty("HeadFontColor", vbBlack)
    HeadFile = PropBag.ReadProperty("HeadFile", "")
    FootFontName = PropBag.ReadProperty("FootFontName", "宋体")
    FootFontSize = PropBag.ReadProperty("FootFontSize", 10)
    FootFontBold = PropBag.ReadProperty("FootFontBold", False)
    FootFontItalic = PropBag.ReadProperty("FootFontItalic", False)
    FootFontUnderline = PropBag.ReadProperty("FootFontUnderline", False)
    FootFontStrikethrough = PropBag.ReadProperty("FootFontStrikethrough", False)
    FootFontColor = PropBag.ReadProperty("FootFontColor", vbBlack)
    FootFile = PropBag.ReadProperty("FootFile", "")
    If Ambient.UserMode Then
        '获取默认的页面属性
        PaperKind = GetSetting(UCase(App.ProductName), "PAGE", UCase("PaperKind"), cprPKA4)
        PaperOrient = GetSetting(UCase(App.ProductName), "PAGE", UCase("PaperOrient"), cprPOPortrait)
        If PaperKind <> cprPKCustom Then
            If PaperOrient = cprPOPortrait Then
                PaperHeight = Val(Split(PaperKindConst(PaperKind), ",")(1))
                PaperWidth = Val(Split(PaperKindConst(PaperKind), ",")(2))
            Else
                PaperHeight = Val(Split(PaperKindConst(PaperKind), ",")(2))
                PaperWidth = Val(Split(PaperKindConst(PaperKind), ",")(1))
            End If
        Else
            PaperHeight = GetSetting(UCase(App.ProductName), "PAGE", UCase("PaperHeight"), PaperHeight)
            PaperWidth = GetSetting(UCase(App.ProductName), "PAGE", UCase("PaperWidth"), PaperWidth)
        End If
        MarginTop = GetSetting(UCase(App.ProductName), "PAGE", UCase("MarginTop"), MarginTop)
        MarginBottom = GetSetting(UCase(App.ProductName), "PAGE", UCase("MarginBottom"), MarginBottom)
        MarginLeft = GetSetting(UCase(App.ProductName), "PAGE", UCase("MarginLeft"), MarginLeft)
        MarginRight = GetSetting(UCase(App.ProductName), "PAGE", UCase("MarginRight"), MarginRight)
    
        If Not ExistsPrinter Then
            gTargetDC = picBuff.hDC
        Else
            gTargetDC = Printer.hDC
        End If
        gTargetDC = picBuff.hDC     '解决民航医院预览时右边超出的问题！（只有以屏幕为度量）
    End If
    
    ResetWYSIWYG
    
    SetVSWithRtb True
    
    Modified = False    '此属性应该放到最后，避免ViewMode使得内容改变。
    If Ambient.UserMode Then VS_Change
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    If Not UserControl.Extender.Visible Then
        Exit Sub
    End If
    If Ambient.UserMode Then
        If ShowRuler Then
            picHRuler.Move 0, 0, ScaleWidth, 400
        Else
            picHRuler.Move 0, 0, ScaleWidth, 0
        End If
        
        Select Case mvarViewMode
        Case cprNormal
            picHRulerHead.Width = 320
            picMarginL.Visible = True
            picMarginR.Visible = True
            
            If (ScaleWidth - VS.Width) > mvarPaperWidth Then
                picMarginL.Move (ScaleWidth - VS.Width - mvarPaperWidth) / 2, picHRuler.Height, mvarMarginLeft, ScaleHeight - HS.Height - picHRuler.Height
                RTBNormal.Move picMarginL.Left + picMarginL.Width, picHRuler.Height, mvarPaperWidth - mvarMarginLeft - mvarMarginRight + 260, picMarginL.Height
                picMarginR.Move RTBNormal.Left + RTBNormal.Width - 240, picHRuler.Height, mvarMarginRight, picMarginL.Height
                HS.Enabled = False
            Else
                picMarginL.Move 0, picHRuler.Height, mvarMarginLeft, ScaleHeight - HS.Height - picHRuler.Height
                RTBNormal.Move picMarginL.Width, picHRuler.Height, ScaleWidth - picMarginL.Width - VS.Width, picMarginL.Height
                picMarginR.Move RTBNormal.Left + RTBNormal.Width - 240, picHRuler.Height, mvarMarginRight, picMarginL.Height
                '设置水平滚动条的最大值
                Dim Pos As POINTAPI, lngMax As Long
                HS.Max = lngMax
                SendMessage RTBNormal.hwnd, EM_GETSCROLLPOS, 0, Pos
                HS.Value = Pos.x
                HS.Enabled = True
            End If
            VS.LargeChange = WHEELNUMBER
        Case cprPaper
            picMarginL.Visible = False
            picMarginR.Visible = False
            ShowPages False '刷新 VS.MAX 和 VS.Value，不刷新数据
            picHRulerHead.Width = 390
        End Select
        
        HRuler.Width = mvarPaperWidth
        btnNormal.Move 0, UserControl.ScaleHeight - btnNormal.Height
        btnPaper.Move btnNormal.Left + btnNormal.Width, btnNormal.Top
        HS.Move btnPaper.Left + btnPaper.Width, btnNormal.Top, UserControl.ScaleWidth - btnPaper.Width * 2 - picNull.Width
        picNull.Move ScaleWidth - picNull.Width, ScaleHeight - picNull.Height
        Progress1.Move IIf(ScaleWidth > 4500, ScaleWidth - Progress1.Width - 500, 1000), ScaleHeight - HS.Height + 15, IIf(ScaleWidth > 4500, 2000, Abs(ScaleWidth - 1500))   '先修正进度条位置
        VS.Move ScaleWidth - VS.Width, IIf(ShowRuler, picHRuler.Height, 0), 250, ScaleHeight - IIf(ShowRuler, picHRuler.Height, 0) - HS.Height
        
        Call HS_Change
        UpdateWindow UserControl.hwnd
        RaiseEvent Resize(mvarViewMode)
    End If
    Err.Clear
End Sub

Private Sub UserControl_Show()
    If Not UserControl.Extender.Visible Then
        Exit Sub
    End If
    If Ambient.UserMode Then
        imgIcon.Visible = False
        lblThis.Visible = False
        If UserControl.Extender.Visible And UserControl.Extender.Enabled Then
                        On Error Resume Next
            UserControl.Extender.SetFocus
            Err.Clear
        End If
    Else
        imgIcon.Visible = True
        lblThis.Visible = True
        picHRuler.Visible = False
        HS.Visible = False
        VS.Visible = False
        picNull.Visible = False
        RTBNormal.Visible = False
        RTBPaper(1).Visible = False
        picShadow(1).Visible = False
        btnNormal.Visible = False
        btnPaper.Visible = False
        VS.Visible = False
    End If
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    Set Image1.Picture = Nothing
    Set imgIcon.Picture = Nothing
    ImlScroll.ListImages.Clear
    ImageList_Destroy ImlScroll.hImageList
    Set picBlank.Picture = Nothing
    Set picBorder.Picture = Nothing
    Set picBuff.Picture = Nothing
    Set picHRuler.Picture = Nothing
    Set picHRulerHead.Picture = Nothing
    Set picMarginL.Picture = Nothing
    Set picMarginR.Picture = Nothing
    Set picNull.Picture = Nothing
    Set picShadow(1).Picture = Nothing
    Set picUI.Picture = Nothing
    Set Me.MouseIcon = Nothing
    Call SendMessage(RTBTmp.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBHead.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBFoot.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))

    Set PubInfo = Nothing
    If Not mfrmFindText Is Nothing Then Unload mfrmFindText
    Set mfrmFindText = Nothing
    Set mvarPicture = Nothing
    Err.Clear
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'当保存对象的实例时，发生该事件。该事件通知对象此时需要保存对象的状态，以便将来可恢复该状态。大多数情况下，对象的状态仅包括属性值。
'属性保存（静态属性的保存）
    PropBag.WriteProperty "AutoDetectURL", AutoDetectURL, True
    PropBag.WriteProperty "BackColor", BackColor, &H99A8AC
    PropBag.WriteProperty "PaperColor", PaperColor, vbWhite
    PropBag.WriteProperty "Border", Border, False
    PropBag.WriteProperty "DefaultTabStop", DefaultTabStop, 21
    PropBag.WriteProperty "DoDefaultURLClick", DoDefaultURLClick, False
    PropBag.WriteProperty "Enabled", Enabled, True
    PropBag.WriteProperty "FileName", FileName, ""
    PropBag.WriteProperty "ForceEdit", ForceEdit, False
    PropBag.WriteProperty "Modified", Modified, False
    PropBag.WriteProperty "ReadOnly", ReadOnly, False
    PropBag.WriteProperty "Title", Title, "未命名文档"
    PropBag.WriteProperty "ZoomFactor", ZoomFactor, 1#
    PropBag.WriteProperty "Foot", Foot, ""
    PropBag.WriteProperty "Head", Head, ""
    PropBag.WriteProperty "PaperHeight", PaperHeight, 16840
    PropBag.WriteProperty "PaperWidth", PaperWidth, 11907
    PropBag.WriteProperty "MarginTop", MarginTop, 1400
    PropBag.WriteProperty "MarginBottom", MarginBottom, 1400
    PropBag.WriteProperty "MarginLeft", MarginLeft, 1800
    PropBag.WriteProperty "MarginRight", MarginRight, 1800
    PropBag.WriteProperty "Transparent", Transparent, False
    PropBag.WriteProperty "ShowPageNumber", ShowPageNumber, True
    PropBag.WriteProperty "PageCount", PageCount, 1
    PropBag.WriteProperty "CurPage", CurPage, 1
    PropBag.WriteProperty "ViewMode", ViewMode, cprNormal
    PropBag.WriteProperty "WithViewButtonas", WithViewButtonas, True
    PropBag.WriteProperty "PaperKind", PaperKind, cprPKA4
    PropBag.WriteProperty "PaperOrient", PaperOrient, cprPOPortrait
    PropBag.WriteProperty "ShowRuler", ShowRuler, True
    PropBag.WriteProperty "AuditMode", AuditMode, False
    PropBag.WriteProperty "HeadFontName", HeadFontName, "宋体"
    PropBag.WriteProperty "HeadFontSize", HeadFontSize, 10
    PropBag.WriteProperty "HeadFontBold", HeadFontBold, False
    PropBag.WriteProperty "HeadFontItalic", HeadFontItalic, False
    PropBag.WriteProperty "HeadFontUnderline", HeadFontUnderline, False
    PropBag.WriteProperty "HeadFontStrikethrough", HeadFontStrikethrough, False
    PropBag.WriteProperty "HeadFontColor", HeadFontColor, vbBlack
    PropBag.WriteProperty "HeadFile", HeadFile, ""
    PropBag.WriteProperty "FootFontName", FootFontName, "宋体"
    PropBag.WriteProperty "FootFontSize", FootFontSize, 10
    PropBag.WriteProperty "FootFontBold", FootFontBold, False
    PropBag.WriteProperty "FootFontItalic", FootFontItalic, False
    PropBag.WriteProperty "FootFontUnderline", FootFontUnderline, False
    PropBag.WriteProperty "FootFontStrikethrough", FootFontStrikethrough, False
    PropBag.WriteProperty "FootFontColor", FootFontColor, vbBlack
    PropBag.WriteProperty "FootFile", FootFile, ""
    
    PropertyChanged "AutoDetectURL"
    PropertyChanged "BackColor"
    PropertyChanged "PaperColor"
    PropertyChanged "Border"
    PropertyChanged "DefaultTabStop"
    PropertyChanged "DoDefaultURLClick"
    PropertyChanged "Enabled"
    PropertyChanged "FileName"
    PropertyChanged "ForceEdit"
    PropertyChanged "Modified"
    PropertyChanged "ReadOnly"
    PropertyChanged "Title"
    PropertyChanged "ZoomFactor"
    PropertyChanged "Foot"
    PropertyChanged "Head"
    PropertyChanged "PaperHeight"
    PropertyChanged "PaperWidth"
    PropertyChanged "MarginTop"
    PropertyChanged "MarginBottom"
    PropertyChanged "MarginLeft"
    PropertyChanged "MarginRight"
    PropertyChanged "Transparent"
    PropertyChanged "ShowPageNumber"
    PropertyChanged "PageCount"
    PropertyChanged "CurPage"
    PropertyChanged "ViewMode"
    PropertyChanged "WithViewButtonas"
    PropertyChanged "PaperKind"
    PropertyChanged "PaperOrient"
    PropertyChanged "ShowRuler"
    PropertyChanged "AuditMode"
    PropertyChanged "HeadFontName"
    PropertyChanged "HeadFontSize"
    PropertyChanged "HeadFontBold"
    PropertyChanged "HeadFontItalic"
    PropertyChanged "HeadFontUnderline"
    PropertyChanged "HeadFontStrikethrough"
    PropertyChanged "HeadFontColor"
    PropertyChanged "HeadFile"
    PropertyChanged "FootFontName"
    PropertyChanged "FootFontSize"
    PropertyChanged "FootFontBold"
    PropertyChanged "FootFontItalic"
    PropertyChanged "FootFontUnderline"
    PropertyChanged "FootFontStrikethrough"
    PropertyChanged "FootFontColor"
    PropertyChanged "FootFile"
End Sub

Private Sub HS_ButtonClick(ByVal lButton As Long)
    Select Case lButton
    Case 1
        ViewMode = cprNormal
    Case 2
        ViewMode = cprPaper
    End Select
End Sub


Private Sub VS_Change()
    If InProcessing Then Exit Sub
    If mvarViewMode = cprNormal Then
        '普通模式
        Dim Pos As POINTAPI, lngH As Long
        SendMessage hWndRTB, EM_GETSCROLLPOS, 0, Pos
        Pos.x = Pos.x
        Pos.y = CLng(VS.Value) * 15
        SendMessage hWndRTB, EM_SETSCROLLPOS, 0, Pos
        RefreshUIInterface
    ElseIf mvarViewMode = cprPaper Then
        Dim M As Long, H As Long, Hi As Long, k As Long, i As Long, lngVS As Long
        lngVS = VS.Value
        H = ScaleHeight - picHRuler.Height - HS.Height
        Hi = (PAGEMARGIN + mvarPaperHeight) * mvarZoomFactor
        k = Hi / VSTEP
        M = CInt(H / Hi) + 2
        mvarStartPage = CInt(lngVS / k)
        mvarEndPage = mvarStartPage + M
        mvarCurPage = mvarStartPage
        
        Dim lTOp As Long
        For i = 1 To mvarPageCount
            If i < mvarStartPage Or i > mvarEndPage Then
                RTBPaper(i).Visible = False
                picShadow(i).Visible = False
            Else
                lTOp = (Hi * (i - 1) + PAGEMARGIN - lngVS * VSTEP) * mvarZoomFactor + picHRuler.Height
                RTBPaper(i).Top = lTOp
                picShadow(i).Top = lTOp + SHADOWOFFSET
                RTBPaper(i).Visible = True
                picShadow(i).Visible = True
            End If
        Next
        
        Call HS_Change
    End If
End Sub

Private Sub HS_Change()
    If mvarViewMode = cprNormal Then
        '普通模式
        Dim Pos As POINTAPI, lngMax As Long
        lngMax = (mvarPaperWidth - MarginRight - RTBNormal.OriginRTB.Width) / Screen.TwipsPerPixelX
        HS.Max = lngMax
        SendMessage hWndRTB, EM_GETSCROLLPOS, 0, Pos
        Pos.x = HS.Value
        Pos.y = Pos.y
        SendMessage hWndRTB, EM_SETSCROLLPOS, 0, Pos
        HRuler.Left = picMarginL.Left
    ElseIf mvarViewMode = cprPaper Then
        '页面模式
        Dim M As Long, H As Long, Hi As Long, k As Long, i As Long, W As Long, Wi As Long
        H = ScaleHeight - picHRuler.Height - HS.Height
        Hi = (PAGEMARGIN + mvarPaperHeight) * mvarZoomFactor
        k = Hi / VSTEP
        M = CInt(H / Hi) + 2
        mvarStartPage = CInt(VS.Value / k)
        mvarEndPage = mvarStartPage + M
    
        W = ScaleWidth - VS.Width
        Wi = (2 * PAGEMARGIN + mvarPaperWidth) * mvarZoomFactor
        
        Dim lLeft As Long
        If Wi < W Then
            lLeft = (W - Wi) / 2 + PAGEMARGIN
        Else
            lLeft = (PAGEMARGIN - HS.Value * HSTEP) * mvarZoomFactor
        End If
        If mvarViewMode = cprPaper Then HRuler.Left = lLeft
        For i = 1 To mvarPageCount
            If i < mvarStartPage Or i > mvarEndPage Then
                RTBPaper(i).Visible = False
                picShadow(i).Visible = False
            Else
                RTBPaper(i).Left = lLeft
                picShadow(i).Left = lLeft + SHADOWOFFSET
                RTBPaper(i).Visible = True
                picShadow(i).Visible = True
            End If
        Next
    End If
End Sub

'############################################################################################################
'## 功能：  打印单独页面到指定设备（打印机/图片框）
'##
'## 参数：  PageNumber      ：页码
'##         objTarget       ：打印的目标控件（Printer/图片框）
'##         blnPreview      ：是否是预览模式（如果是预览模式，那么页眉页脚颜色为灰色；正式打印是黑色）
'##         lngBlankHeight      ：外部指定的上部留白高度
'############################################################################################################
Public Sub PrintPage(ByVal PageNumber As Long, Optional ByRef objTarget As Object = Nothing, _
    Optional ByVal blnPreview As Boolean = False, Optional ByVal lngBlankHeight As Long = 0, _
    Optional ByVal blnMarginReverse As Boolean)
Dim lngOffsetLeft As Long, lngOffsetTop As Long      '左边缘偏移量'上边缘偏移量
Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
Dim lngPicWidth As Long, lngPicHeight As Long   '图片宽度高度
Dim lngPageCount As Long        '总页数
Dim lngHead As Long, lngFoot As Long          '对于多行页眉的高度'对于多行页脚的高度
Dim fr As FORMATRANGE           '格式化的文本范围
Dim rcDrawTo As RECT            '目标文字区域
Dim rcPage As RECT              '目标页面区域
Dim Rct As RECT                 '打印页眉页脚
Dim lngNextPos As Long          '下一个字符位置
Dim strHead As String, strFoot As String    '页眉页脚

    If objTarget Is Nothing Then Set objTarget = Printer
    objTarget.ScaleMode = vbTwips   '设置打印机单位为缇。
    
    '图片高度和宽度
    lngPicWidth = 0: lngPicHeight = 0
    If Not (mvarPicture Is Nothing) Then
        If mvarPicture.Handle <> 0 Then
            lngPicWidth = objTarget.ScaleX(mvarPicture.Width, vbHimetric, vbTwips)
            lngPicHeight = objTarget.ScaleX(mvarPicture.Height, vbHimetric, vbTwips)
        End If
    End If
    
    If Not ExistsPrinter Then
        gTargetDC = picBuff.hDC
    Else
        gTargetDC = Printer.hDC
    End If
    gTargetDC = picBuff.hDC     '解决民航医院预览时右边超出的问题！（只有以屏幕为度量）
    
    '获取打印机可打印区域的边缘偏移量，单位：Pixel
    lngOffsetLeft = objTarget.ScaleX(GetDeviceCaps(objTarget.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = objTarget.ScaleY(GetDeviceCaps(objTarget.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    If lngPicHeight > 0 Then '打印页眉图片
        If blnMarginReverse Then '要求左右边距反向，支持双面打印
            objTarget.PaintPicture mvarPicture, (mvarMarginRight - lngOffsetLeft), (mvarMarginTop - lngOffsetTop), lngPicWidth, lngPicHeight
        Else
            objTarget.PaintPicture mvarPicture, (mvarMarginLeft - lngOffsetLeft), (mvarMarginTop - lngOffsetTop), lngPicWidth, lngPicHeight
        End If
    End If
    
    With rcPage
        .Left = 0
        .Top = 0
        .Right = mvarPaperWidth
        .Bottom = mvarPaperHeight
    End With
    '计算页眉高度
    RTBTmp.PaperWidth = mvarPaperWidth: RTBTmp.MarginLeft = mvarMarginLeft: RTBTmp.MarginRight = mvarMarginRight: RTBTmp.ResetWYSIWYG
    RTBTmp.TextRTF = HeadFileTextRTF
    Call DocTmpReplaceKey("", "", blnPreview) '需要先替换出其中的关键字
    Call DocTmpReplaceKey("{页码}", PageNumber)       '页码的替换
    With rcDrawTo
        If blnMarginReverse Then '要求左右边距反向，支持双面打印
            .Left = (mvarMarginRight - lngOffsetLeft)
            .Right = (mvarPaperWidth - lngOffsetLeft) - IIf(mvarMarginLeft >= lngOffsetLeft, mvarMarginLeft, lngOffsetLeft)
        Else
            .Left = (mvarMarginLeft - lngOffsetLeft)
            .Right = (mvarPaperWidth - lngOffsetLeft) - IIf(mvarMarginRight >= lngOffsetLeft, mvarMarginRight, lngOffsetLeft)
        End If
        .Top = (mvarMarginTop - lngOffsetTop) + lngPicHeight
        .Bottom = IIf(RTBTmp.Text = "", 0, 99999)
    End With
    With fr
        .hDC = objTarget.hDC
        .hdcTarget = gTargetDC
        .rc = rcDrawTo
        .rcPage = rcPage
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    Call SendMessage(RTBTmp.hWndRTB, EM_FORMATRANGE, 1, fr)  '打印页眉
    lngHead = fr.rc.Bottom - fr.rc.Top
    If lngHead <= 0 Then lngHead = 0
    If RTBHead.Text <> "" Or lngPicHeight > 0 Then
        objTarget.ForeColor = IIf(blnPreview, RGB(149, 149, 149), vbBlack)
        objTarget.Line (fr.rc.Left, fr.rc.Bottom + 50)-(fr.rc.Right, fr.rc.Bottom + 50)
    End If

    RTBTmp.TextRTF = FootFileTextRTF
    Call DocTmpReplaceKey("", "", blnPreview) '需要先替换出其中的关键字
    Call DocTmpReplaceKey("{页码}", PageNumber)       '页码的替换
    Call SendMessage(RTBTmp.hWndRTB, EM_FORMATRANGE, 0, fr)
    lngFoot = fr.rc.Bottom - fr.rc.Top
    If lngFoot <= 0 Then lngFoot = 0
    
    '设置可打印文字区域
    If blnMarginReverse Then '要求左右边距反向，支持双面打印
        lngLeft = (mvarMarginRight - lngOffsetLeft)
        lngRight = (mvarPaperWidth - lngOffsetLeft) - IIf(mvarMarginLeft >= lngOffsetLeft, mvarMarginLeft, lngOffsetLeft)
    Else
        lngLeft = (mvarMarginLeft - lngOffsetLeft) '边距应该已经包含打印偏移
        lngRight = (mvarPaperWidth - lngOffsetLeft) - IIf(mvarMarginRight >= lngOffsetLeft, mvarMarginRight, lngOffsetLeft)
    End If
    lngTop = (mvarMarginTop - lngOffsetTop) + lngPicHeight + lngHead + IIf(lngHead > 0, 350, 0)
    lngBottom = mvarPaperHeight - mvarMarginBottom - lngFoot

    rcDrawTo.Left = lngLeft
    rcDrawTo.Top = lngTop
    rcDrawTo.Right = lngRight
    rcDrawTo.Bottom = lngBottom
    
    '设置打印指令（FormatRange消息需要的打印信息）
    fr.hDC = objTarget.hDC          ' 度量和渲染使用相同的DC
    fr.hdcTarget = gTargetDC        ' 目标控件的DC（关键对象）
    fr.rc = rcDrawTo                ' 文字矩形区域 IN/OUT
    fr.rcPage = rcPage              ' 整个页面矩形区域 IN
    fr.chrg.cpMin = AllPages(PageNumber).Start ' 打印区域的文字开始位置
    fr.chrg.cpMax = AllPages(PageNumber).End   ' 文字结束位置（-1表示直到末尾）
    
    '因为实际打印时文字位置向下有偏移，重新计算遮罩高度
    If lngBlankHeight > lngTop Then  '遮罩到文字部份
        Dim frBlank As FORMATRANGE, rcBlank As RECT, lngblankPos As Long
        With rcBlank
            .Top = lngTop
            .Left = lngLeft
            .Right = lngRight
            .Bottom = lngBlankHeight - lngOffsetTop
        End With
        
        With frBlank
            .hDC = fr.hDC
            .hdcTarget = gTargetDC
            .rc = rcBlank
            .rcPage = fr.rcPage
            .chrg.cpMin = fr.chrg.cpMin
            .chrg.cpMax = -1
        End With
        '以遮罩区域虚拟打印后,Bottom会发生偏移，得出该区域Bottom位置即为遮罩高度
        lngblankPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, frBlank)  '虚拟打印
        '只遮罩1－2行时，rich控件算不准，按实际遮罩计
        lngBlankHeight = IIf((rcBlank.Bottom - rcBlank.Top) <= 600, lngBlankHeight - lngOffsetTop, frBlank.rc.Bottom + 350)
    Else
        lngBlankHeight = lngBlankHeight
    End If
    
    '发送 EM_FORMATRANGE 消息进行打印
    lngNextPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, fr)  '虚拟打印
    fr.rc = rcDrawTo
    If lngNextPos < AllPages(PageNumber).End Then fr.rc.Bottom = fr.rc.Bottom + 99999    '保证一次打印完整页！
    lngNextPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 1, fr)  '实际打印
    

    If lngBlankHeight < 200 Then '有遮罩时不打页脚，因为是在一张纸上打多次
        With rcDrawTo
            If blnMarginReverse Then
                .Left = (mvarMarginRight - lngOffsetLeft)
                .Right = (mvarPaperWidth - lngOffsetLeft) - IIf(mvarMarginLeft >= lngOffsetLeft, mvarMarginLeft, lngOffsetLeft)
            Else
                .Left = (mvarMarginLeft - lngOffsetLeft)
                .Right = (mvarPaperWidth - lngOffsetLeft) - IIf(mvarMarginRight >= lngOffsetLeft, mvarMarginRight, lngOffsetLeft)
            End If
            .Top = IIf(fr.rc.Bottom > lngBottom, fr.rc.Bottom, lngBottom)
            .Bottom = 99999
        End With
        With fr
            .hDC = objTarget.hDC
            .hdcTarget = gTargetDC
            .rcPage = rcPage
            .rc = rcDrawTo
            .chrg.cpMin = 0
            .chrg.cpMax = -1
        End With
        Call SendMessage(RTBTmp.hWndRTB, EM_FORMATRANGE, 1, fr)
    End If
        
    '绘制上部留白矩形
    If lngBlankHeight > 0 Then
        objTarget.PaintPicture picBlank.Image, 0, 0, mvarPaperWidth, lngBlankHeight
    End If
    
    '允许RTF释放内存
    Call SendMessage(RTBTmp.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBHead.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBFoot.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
End Sub

'############################################################################################################
'## 功能：  获取文本高度（通过虚拟打印获取）
'##
'## 参数：  PageNumber      ：页码
'##         objTarget       ：打印的目标控件（Printer/图片框）
'##         blnPreview      ：是否是预览模式（如果是预览模式，那么页眉页脚颜色为灰色；正式打印是黑色）
'##         lngBlankHeight      ：外部指定的上部留白高度
'############################################################################################################
Private Function GetPrintHeight() As Long
    Dim fr As FORMATRANGE           '格式化的文本范围
    Dim rcDrawTo As RECT            '目标文字区域
    Dim rcPage As RECT              '目标页面区域
    Dim lngNextPos As Long          '下一个字符位置
    Dim strHead As String, strFoot As String    '页眉页脚
    Dim r As Long                   '返回值
    
    picBuff.ScaleMode = vbTwips   '设置打印机单位为缇。
    
    '设置可打印页面区域
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = mvarPaperWidth
    rcPage.Bottom = 999999999#
    
    '设置可打印文字区域
    rcDrawTo.Left = mvarMarginLeft
    rcDrawTo.Top = mvarMarginTop
    rcDrawTo.Right = mvarPaperWidth - mvarMarginRight
    rcDrawTo.Bottom = 999999999#
    fr.hDC = picBuff.hDC            ' 度量和渲染使用相同的DC
    fr.hdcTarget = picBuff.hDC      ' 目标控件的DC（关键对象）
    fr.rc = rcDrawTo                ' 文字矩形区域 IN/OUT
    fr.rcPage = rcPage              ' 整个页面矩形区域 IN
    fr.chrg.cpMin = 0               ' 打印区域的文字开始位置
    fr.chrg.cpMax = -1              ' 文字结束位置（-1表示直到末尾）
    
    '发送 EM_FORMATRANGE 消息进行打印
    lngNextPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, fr)  '虚拟打印
    If fr.rc.Bottom = 999999999# Then fr.rc.Bottom = mvarPaperHeight
    GetPrintHeight = (fr.rc.Bottom - fr.rc.Top - UserControl.Height)
    If GetPrintHeight < mvarPaperHeight - mvarMarginTop - mvarMarginBottom Then
        GetPrintHeight = mvarPaperHeight - mvarMarginTop - mvarMarginBottom
    End If
    GetPrintHeight = GetPrintHeight / Screen.TwipsPerPixelY
End Function

'############################################################################################################
'## 功能：  根据rtb的Scroll大小及位置设置VS的大小及位置
'############################################################################################################
Private Sub SetVSWithRtb(Optional ByVal blnInit As Boolean)
    Dim SclInf As SCROLLINFO
    
    If mvarViewMode = cprPaper Then Exit Sub
    If blnInit Then
        SclInf.cbSize = Len(SclInf): SclInf.fMask = SIF_ALL
        SclInf.nMax = 0: SclInf.nPos = 0
        SetScrollInfo hWndRTB, SB_VERT, SclInf, True
        VS.Max = 0
        VS.Value = 0
        SetVSWithRtb
    Else
        SclInf.cbSize = Len(SclInf): SclInf.fMask = SIF_ALL
        GetScrollInfo hWndRTB, SB_VERT, SclInf
        VS.Max = SclInf.nMax \ Screen.TwipsPerPixelY
        VS.Value = SclInf.nPos \ Screen.TwipsPerPixelY
    End If
End Sub



'############################################################################################################
'## 功能：  执行虚拟打印，更新RTF的分页信息
'############################################################################################################
Public Sub DoVirtualPrint()
Dim lngPicWidth As Long, lngPicHeight As Long   '图片宽度高度
Dim lngHead As Long, lngFoot As Long            '页眉页脚高度
Dim lngPageCount As Long        '总页数
Dim fr As FORMATRANGE           '格式化的文本范围
Dim rcDrawTo As RECT            '目标文字区域
Dim rcPage As RECT              '目标页面区域
Dim lngNextPos As Long          '下一个字符位置
Dim r As Long                   '返回值

    On Error Resume Next
    lngPicWidth = 0: lngPicHeight = 0 '图片高度和宽度
    If Not (mvarPicture Is Nothing) Then
        If mvarPicture.Handle <> 0 Then
            lngPicWidth = UserControl.ScaleX(mvarPicture.Width, vbHimetric, vbTwips)
            lngPicHeight = UserControl.ScaleX(mvarPicture.Height, vbHimetric, vbTwips)
        End If
    End If
    
    picBuff.Width = mvarPaperWidth: picBuff.Height = mvarPaperHeight
    picBuff.ScaleMode = vbTwips   '设置打印机单位为缇。
    
    '设置可打印页面区域
    With rcPage
        .Left = 0
        .Top = 0
        .Right = mvarPaperWidth
        .Bottom = mvarPaperHeight
    End With
    
    If Not ExistsPrinter Then
        gTargetDC = picBuff.hDC
    Else
        gTargetDC = Printer.hDC
    End If
    gTargetDC = picBuff.hDC     '解决民航医院预览时右边超出的问题！（只有以屏幕为度量）
    
    '计算页眉高度
    With rcDrawTo
        .Left = mvarMarginLeft
        .Top = mvarMarginTop
        .Right = mvarPaperWidth - mvarMarginRight
        .Bottom = IIf(RTBHead.Text = "", 0, 99999)
    End With
    With fr
        .hDC = picBuff.hDC
        .hdcTarget = gTargetDC
        .rcPage = rcPage
        .rc = rcDrawTo
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    Call SendMessage(RTBHead.hWndRTB, EM_FORMATRANGE, 0, fr)
    lngHead = fr.rc.Bottom - fr.rc.Top
    If lngHead < 0 Then lngHead = 0
    Call SendMessage(RTBFoot.hWndRTB, EM_FORMATRANGE, 0, fr)
    lngFoot = fr.rc.Bottom - fr.rc.Top
    If lngFoot < 0 Then lngFoot = 0
    
    '设置可打印文字区域
    rcDrawTo.Left = mvarMarginLeft
    rcDrawTo.Top = mvarMarginTop + lngPicHeight + lngHead + IIf(lngHead > 0, 350, 0)
    rcDrawTo.Right = mvarPaperWidth - mvarMarginRight
    rcDrawTo.Bottom = mvarPaperHeight - mvarMarginBottom - lngFoot
    
    '设置打印指令（FormatRange消息需要的打印信息）
    fr.hDC = picBuff.hDC            ' 渲染设备
    fr.hdcTarget = gTargetDC        ' 目标设备（关键对象）
    fr.rc = rcDrawTo                ' 文字矩形区域 IN/OUT
    fr.rcPage = rcPage              ' 整个页面矩形区域 IN
    fr.chrg.cpMin = 0               ' 打印区域的文字开始位置
    fr.chrg.cpMax = -1              ' 文字结束位置（-1表示直到末尾）
    
    '获取整个RTF文本长度
    Dim lngTmp As Long              '用于记录单页字符起始位置
    Dim lngLen As Long              '总长度（中英文混合长度）
    lngLen = lstrlen(RTBNormal.Text)
    
    '循环分页打印
    Do
        '发送 EM_FORMATRANGE 消息进行虚拟打印
        lngNextPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, fr)     '只分页，不打印
        
        lngPageCount = lngPageCount + 1             ' 页数＋1
        '记录分页信息
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          '实际打印高度
        AllPages(lngPageCount).Start = lngTmp
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos                      ' 下一页起始字符位置
        fr.hDC = picBuff.hDC
        fr.hdcTarget = gTargetDC
        fr.rc = rcDrawTo                                ' 必须重新设置文字区域，否则有误！
        If lngNextPos <= lngTmp Or lngNextPos >= lngLen Then Exit Do      ' 完成所有页面的分页
        lngTmp = lngNextPos
    Loop
    PageCount = lngPageCount
    AllPages(lngPageCount).End = -1                     ' 最后一页结束位置为最末尾
    
    '允许RTF释放内存
    r = SendMessage(RTBHead.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    r = SendMessage(RTBFoot.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    r = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Err.Clear
End Sub

'############################################################################################################
'## 功能：  打印当前文档到打印机
'##
'## 参数：  blnNoAsk            ：是否显示打印对话框，在打印前进行设置
'##         lngStartPage        ：外部指定的起始页
'##         lngBlankHeight      ：外部指定的上部留白高度
'##         lngCopies           ：指定打印份数，＝0不指定，不控制，否则控制打印份数不能修改，并通过参数返回打印份数
'############################################################################################################
Public Function PrintDoc(Optional ByVal blnNoAsk As Boolean, Optional ByVal lngStartPage As Long, Optional ByVal lngBlankHeight As Long, _
    Optional ByRef strPrinterDeviceName As String, Optional ByRef lngCopies As Long = 0) As Boolean
    
    Dim strOldPrinterName As String
    
    If Not ExistsPrinter Then MsgBox "没有安装打印设备，不能打印！", vbExclamation, App.Title: Exit Function
    If mvarViewMode <> cprPaper Then
        Me.InProcessing = True
        '获取分页信息
        DoVirtualPrint
        Me.InProcessing = False
    End If
    
    Dim intPageFrom As Integer, intPageTo As Integer, bytPageOddEven As Byte
    Dim blnCopyOrder As Boolean, blnDuplex As Boolean, blnCurReverse As Boolean
    Dim t As Variant, aryPage() As String, i As Long, j As Long, k As Long, L As Long, M As Long
    Dim lngPageCount As Long
    Dim Pages() As Long             '打印范围内的所有需打印的页面
    Dim blnRangePrint As Boolean    '是否是页码范围打印
    Dim blnHave As Boolean
    Dim blnFirstPrinted As Boolean
    
    intPageFrom = IIf(lngStartPage > 0, lngStartPage, 1): intPageTo = Me.PageCount: blnCopyOrder = True
    blnRangePrint = False
    ReDim Pages(0 To 0) As Long
    If blnNoAsk = False Then
        strOldPrinterName = Printer.DeviceName
        With frmPrintAsk
            .lngPageStart = intPageFrom
            .lngPageEnd = intPageTo
            .lngCopies = lngCopies
            .txtPageScope.Tag = intPageFrom & "-" & intPageTo
            .txtPageScope.Text = .txtPageScope.Tag
            If strPrinterDeviceName <> "" Then '因为有可能指定的打印机不存在，不能使用直接=方式
                For i = 0 To .cboPrinterName.ListCount - 1
                    If .cboPrinterName.List(i) = strPrinterDeviceName Then
                        .cboPrinterName.ListIndex = i
                        Exit For
                    End If
                Next
            End If
            .Show vbModal, Me.Parent
            If .blnOK = False Then Unload frmPrintAsk: Exit Function
            
            If .optPageScope(2).Value = True Then
                '页码范围
                blnRangePrint = True
                t = Split(.txtPageScope.Tag, ",")
                For i = 0 To UBound(t)
                    aryPage = Split(t(i), "-")
                    If UBound(aryPage) = 0 Then
                        '只有一页
                        lngPageCount = UBound(Pages) + 1
                        ReDim Preserve Pages(0 To lngPageCount) As Long
                        Pages(lngPageCount) = Val(t(i))
                    ElseIf UBound(aryPage) = 1 Then
                        L = Val(Split(t(i), "-")(0))
                        M = Val(Split(t(i), "-")(1))
                        For j = L To M Step IIf(M > L, 1, -1)
                            blnHave = False
                            For k = 1 To UBound(Pages)
                                If Pages(k) = j Then blnHave = True
                            Next
                            If blnHave = False Then
                                lngPageCount = UBound(Pages) + 1
                                ReDim Preserve Pages(0 To lngPageCount) As Long
                                Pages(lngPageCount) = j
                            End If
                        Next
                    End If
                Next
            ElseIf .optPageScope(1).Value = True Then
                '当前页
                intPageFrom = Me.CurPage: intPageTo = Me.CurPage
            Else
                '全部打印
                intPageFrom = IIf(lngStartPage > 0, lngStartPage, 1): intPageTo = Me.PageCount
            End If
            blnDuplex = (.chkDuplex.Value = vbChecked)
            bytPageOddEven = .cboPageOddEven.ListIndex
            lngCopies = Val(.txtCopies.Text)
            blnCopyOrder = IIf(.chkCopyOrder.Value = vbChecked, True, False)
            If Printers(.cboPrinterName.ListIndex).DeviceName <> Printer.DeviceName Then
                Set Printer = Printers(.cboPrinterName.ListIndex)
            End If
            strPrinterDeviceName = Printer.DeviceName
            Unload frmPrintAsk
        End With
    Else
        blnDuplex = True
        lngCopies = 1
        If strPrinterDeviceName <> "" Then
            For i = 0 To Printers.Count - 1
                If Printers(i).DeviceName = strPrinterDeviceName Then
                    Set Printer = Printers(i)
                    Exit For
                End If
            Next
        End If
        strPrinterDeviceName = Printer.DeviceName
    End If
    
    If bytPageOddEven = 1 Then
        '奇数页
        If intPageFrom Mod 2 = 0 Then intPageFrom = intPageFrom + 1
    ElseIf bytPageOddEven = 2 Then
        '偶数页
        If intPageFrom Mod 2 = 1 Then intPageFrom = intPageFrom + 1
    End If
    If intPageFrom > intPageTo Then Exit Function
    
    Dim lngCount As Long, lngNumber As Long
    Err = 0: On Error Resume Next
     
    '设置打印机方向
    If Printer.Orientation <> Me.PaperOrient Then
        Printer.Orientation = Me.PaperOrient
    End If
    '设置纸张，自定义纸张的设置必须放到最后
    If mvarPaperKind = cprPKCustom Then
        Call SetCustomPager(UserControl.hwnd, mvarPaperWidth, mvarPaperHeight)
    Else
        Printer.PaperSize = mvarPaperKind
    End If
    
    If Not ExistsPrinter Then
        gTargetDC = picBuff.hDC
    Else
        gTargetDC = Printer.hDC
    End If
    gTargetDC = picBuff.hDC     '解决民航医院预览时右边超出的问题！（只有以屏幕为度量）
    
    '开始打印
    Printer.Print Space(1)
    
    If blnCopyOrder = True Then
        '逐份打印
        For lngNumber = 1 To lngCopies
            blnCurReverse = True
            For lngCount = intPageFrom To intPageTo Step IIf(bytPageOddEven = 0, 1, 2)
                If blnRangePrint Then
                    '页码范围打印
                    For i = 1 To UBound(Pages)
                        If lngCount = Pages(i) Then
                            If lngNumber > 1 Or blnFirstPrinted Then Printer.NewPage
                            PrintPage lngCount, Printer, , IIf(lngCount = lngStartPage, lngBlankHeight, 0)
                            blnFirstPrinted = True
                            Exit For
                        End If
                    Next
                Else
                    If lngNumber > 1 Or blnFirstPrinted = True Then Printer.NewPage
                    If blnDuplex Then
                        If (intPageFrom = 1 Or (bytPageOddEven = 2 And intPageFrom = 2)) And intPageTo = Me.PageCount Then
                            If bytPageOddEven = 2 Then
                                blnCurReverse = True              '勾选双面打印时，偶数每页 左右边距反向
                            ElseIf bytPageOddEven = 0 Then
                                blnCurReverse = Not blnCurReverse '勾选双面打印，打印全部内容时，每间隔一页，左右页边距反向
                            Else
                                blnCurReverse = False
                            End If
                        Else
                            blnCurReverse = False
                        End If
                    Else
                        blnCurReverse = False '没勾选双面打印时，不反向
                    End If
                    PrintPage lngCount, Printer, , IIf(lngCount = lngStartPage, lngBlankHeight, 0), blnCurReverse
                    blnFirstPrinted = True
                End If
            Next
        Next
    Else
        blnCurReverse = True
        For lngCount = intPageFrom To intPageTo Step IIf(bytPageOddEven = 0, 1, 2)
            For lngNumber = 1 To lngCopies
                If blnRangePrint Then
                    '页码范围打印
                    For i = 1 To UBound(Pages)
                        If lngCount = Pages(i) Then
                            If lngNumber > 1 Or blnFirstPrinted = True Then Printer.NewPage
                            PrintPage lngCount, Printer, , IIf(lngCount = lngStartPage, lngBlankHeight, 0)
                            blnFirstPrinted = True
                            Exit For
                        End If
                    Next
                Else
                    If lngNumber > 1 Or blnFirstPrinted = True Then Printer.NewPage
                    If blnDuplex And lngCopies = 1 Then
                        If (intPageFrom = 1 Or (bytPageOddEven = 2 And intPageFrom = 2)) And intPageTo = Me.PageCount Then
                            If bytPageOddEven = 0 Then
                                blnCurReverse = Not blnCurReverse '勾选双面打印，打印全部内容时，每间隔一页，左右页边距反向
                            ElseIf bytPageOddEven = 2 Then
                                blnCurReverse = True              '勾选双面打印时，偶数每页 左右边距反向
                            Else
                                blnCurReverse = False
                            End If
                        Else
                            blnCurReverse = False
                        End If
                    Else
                        blnCurReverse = False '没勾选双面打印或逐页打印多份时，不反向
                    End If
                    PrintPage lngCount, Printer, , IIf(lngCount = lngStartPage, lngBlankHeight, 0), blnCurReverse
                    blnFirstPrinted = True
                End If
            Next
        Next
    End If
    
    Printer.EndDoc
    
    If blnNoAsk = False Then
        '恢复默认打印机
        If strOldPrinterName <> Printer.DeviceName Then
            For j = 1 To Printers.Count
                If Printers(j).DeviceName = strOldPrinterName Then
                    Set Printer = Printers(j)
                End If
            Next
        End If
    End If
    
    If Not ExistsPrinter Then
        gTargetDC = picBuff.hDC
    Else
        gTargetDC = Printer.hDC
    End If
    gTargetDC = picBuff.hDC     '解决民航医院预览时右边超出的问题！（只有以屏幕为度量）
    
    PrintDoc = True
    Exit Function
    Err.Clear
PrintErr:
    PrintDoc = False
End Function
Public Sub DocTmpReplaceKey(Optional ByVal strSource As String = "", Optional ByVal strTraget As String = "", Optional ByVal blnPreview As Boolean)
'替换掉页眉/页脚中的关键字
Dim strR As String, strUnitName As String, lngS As Long, lngE As Long
Dim strFontName As String, sinFontSize As Single, blnBlod As Boolean, blnItalic As Boolean, blnUnderline As Boolean
    On Error Resume Next
    If strSource <> "" Then
        RTBTmp.Range(0, 0).Selected
        If RTBTmp.FindText(strSource, 4) Then
            With RTBTmp
                '获取原有字体
                lngS = .Selection.StartPos: lngE = .Selection.EndPos
                strFontName = .Range(lngS, lngE).Font.Name
                sinFontSize = .Range(lngS, lngE).Font.SIZE
                blnBlod = .Range(lngS, lngE).Font.Bold
                blnUnderline = .Range(lngS, lngE).Font.Underline
                blnItalic = .Range(lngS, lngE).Font.Italic
                '替换
                .Range(lngS, lngE) = strTraget
                '新字串原有字体
                lngE = lngS + Len(strTraget)
                .Range(lngS, lngE).Font.Name = strFontName
                .Range(lngS, lngE).Font.SIZE = sinFontSize
                .Range(lngS, lngE).Font.Bold = blnBlod
                .Range(lngS, lngE).Font.Underline = blnUnderline
                .Range(lngS, lngE).Font.Italic = blnItalic
            End With
        End If
    Else
        strUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
        Call DocTmpReplaceKey("{单位名称}", strUnitName)
        Call DocTmpReplaceKey("{总页数}", PubInfo.PaperCount)
        Call DocTmpReplaceKey("{标题}", PubInfo.Title)
        Call DocTmpReplaceKey("{路径}", Left(PubInfo.FileName, InStrRev(PubInfo.FileName, "\")))
        Call DocTmpReplaceKey("{文件名}", Mid(PubInfo.FileName, InStrRev(PubInfo.FileName, "\") + 1))
        If Not blnPreview Then
            Call DocTmpReplaceKey("{打印日期}", Format(Now(), "yyyy年mm月dd日"))
            Call DocTmpReplaceKey("{打印时间}", Format(Now(), "hh:MM:ss"))
        End If
    End If
    Err.Clear
End Sub
Public Sub DocHeadCopyWithFormat()
    RTBHead.ClearEndCrlfChar
    RTBHead.Range(0, Len(RTBHead.Text)).Selected
    RTBHead.CopyWithFormat
End Sub
Public Sub DocHeadPasteWithFormat()
    '带格式复制
    RTBHead.ForceEdit = True
    RTBHead.Freeze
    RTBHead.SelectAll
    DoEvents
    RTBHead.PasteWithFormat
    RTBHead.UnFreeze
    RTBHead.ClearEndCrlfChar
End Sub
Public Sub DocHeadReplaceKey(Optional ByVal strSource As String = "", Optional ByVal strTraget As String = "", Optional ByVal blnPreview As Boolean)
'替换掉页眉/页脚中的关键字
Dim strR As String, strUnitName As String, lngS As Long, lngE As Long
Dim strFontName As String, sinFontSize As Single, blnBlod As Boolean, blnItalic As Boolean, blnUnderline As Boolean, lngColor As Long
    On Error Resume Next
    If strSource <> "" Then
        RTBHead.Range(0, 0).Selected
        If RTBHead.FindText(strSource) Then
            With RTBHead
                If Not .ForceEdit Then .ForceEdit = True
                '获取原有字体
                lngS = .Selection.StartPos: lngE = .Selection.EndPos
                strFontName = .Range(lngS, lngE).Font.Name
                sinFontSize = .Range(lngS, lngE).Font.SIZE
                blnBlod = .Range(lngS, lngE).Font.Bold
                blnUnderline = .Range(lngS, lngE).Font.Underline
                blnItalic = .Range(lngS, lngE).Font.Italic
                lngColor = .Range(lngS, lngE).Font.ForeColor
                '替换
                .Range(lngS, lngE) = strTraget
                '新字串原有字体
                lngE = lngS + Len(strTraget)
                .Range(lngS, lngE).Font.Name = strFontName
                .Range(lngS, lngE).Font.SIZE = sinFontSize
                .Range(lngS, lngE).Font.Bold = blnBlod
                .Range(lngS, lngE).Font.Underline = blnUnderline
                .Range(lngS, lngE).Font.Italic = blnItalic
                .Range(lngS, lngE).Font.ForeColor = lngColor
            End With
        End If
    Else
        strUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
        Call DocHeadReplaceKey("{单位名称}", strUnitName)
        Call DocHeadReplaceKey("{标题}", PubInfo.Title)
        Call DocHeadReplaceKey("{路径}", Left(PubInfo.FileName, InStrRev(PubInfo.FileName, "\")))
        Call DocHeadReplaceKey("{文件名}", Mid(PubInfo.FileName, InStrRev(PubInfo.FileName, "\") + 1))
        If Not blnPreview Then
            Call DocHeadReplaceKey("{打印日期}", Format(Now(), "yyyy年mm月dd日"))
            Call DocHeadReplaceKey("{打印时间}", Format(Now(), "hh:MM:ss"))
        End If
    End If
    Err.Clear
End Sub
Public Sub DocFootCopyWithFormat()
    RTBFoot.ClearEndCrlfChar
    RTBFoot.Range(0, Len(RTBFoot.Text)).Selected
    RTBFoot.CopyWithFormat
End Sub
Public Sub DocFootPasteWithFormat()
    '带格式复制
    RTBFoot.ForceEdit = True
    RTBFoot.Freeze
    RTBFoot.SelectAll
    DoEvents
    RTBFoot.PasteWithFormat
    RTBFoot.ClearEndCrlfChar
    RTBFoot.UnFreeze
End Sub
Public Sub DocFootReplaceKey(Optional ByVal strSource As String = "", Optional ByVal strTraget As String = "", Optional ByVal blnPreview As Boolean)
'替换掉页眉/页脚中的关键字
Dim strR As String, strUnitName As String, lngS As Long, lngE As Long
Dim strFontName As String, sinFontSize As Single, blnBlod As Boolean, blnItalic As Boolean, blnUnderline As Boolean, lngColor As Long
    On Error Resume Next
    If strSource <> "" Then
        RTBFoot.Range(0, 0).Selected
        If RTBFoot.FindText(strSource) Then
            With RTBFoot
                If Not .ForceEdit Then .ForceEdit = True
                '获取原有字体
                lngS = .Selection.StartPos: lngE = .Selection.EndPos
                strFontName = .Range(lngS, lngE).Font.Name
                sinFontSize = .Range(lngS, lngE).Font.SIZE
                blnBlod = .Range(lngS, lngE).Font.Bold
                blnUnderline = .Range(lngS, lngE).Font.Underline
                blnItalic = .Range(lngS, lngE).Font.Italic
                lngColor = .Range(lngS, lngE).Font.ForeColor
                '替换
                .Range(lngS, lngE) = strTraget
                '新字串原有字体
                lngE = lngS + Len(strTraget)
                .Range(lngS, lngE).Font.Name = strFontName
                .Range(lngS, lngE).Font.SIZE = sinFontSize
                .Range(lngS, lngE).Font.Bold = blnBlod
                .Range(lngS, lngE).Font.Underline = blnUnderline
                .Range(lngS, lngE).Font.Italic = blnItalic
                .Range(lngS, lngE).Font.ForeColor = lngColor
            End With
        End If
    Else
        strUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
        Call DocFootReplaceKey("{单位名称}", strUnitName)
        Call DocFootReplaceKey("{标题}", PubInfo.Title)
        Call DocFootReplaceKey("{路径}", Left(PubInfo.FileName, InStrRev(PubInfo.FileName, "\")))
        Call DocFootReplaceKey("{文件名}", Mid(PubInfo.FileName, InStrRev(PubInfo.FileName, "\") + 1))
        If Not blnPreview Then
            Call DocFootReplaceKey("{打印日期}", Format(Now(), "yyyy年mm月dd日"))
            Call DocFootReplaceKey("{打印时间}", Format(Now(), "hh:MM:ss"))
        End If
    End If
    Err.Clear
End Sub
Public Sub HeadTextToFile()
'将页眉文字构建成Rtf文件,适用于没有Rtf文件时
Dim strHead As String
    With RTBHead
        strHead = mvarHead
        Do While strHead <> ""
            If Right(strHead, 2) = vbCrLf Then
                strHead = Mid(strHead, 1, Len(strHead) - 2)
            ElseIf Asc(Right(strHead, 1)) = 13 Then
                strHead = Mid(strHead, 1, Len(strHead) - 1)
            ElseIf Asc(Right(strHead, 1)) = 10 Then
                strHead = Mid(strHead, 1, Len(strHead) - 1)
            Else
                Exit Do
            End If
        Loop
            
        .Text = strHead
        If Trim(strHead) = "" Then Exit Sub
        .SelectAll
        .Selection.Font.Name = mvarHeadFontName
        .Selection.Font.SIZE = mvarHeadFontSize
        .Selection.Font.Bold = mvarHeadFontBold
        .Selection.Font.Italic = mvarHeadFontItalic
        .Selection.Font.Underline = mvarHeadFontUnderline
        .Selection.Font.Strikethrough = mvarHeadFontStrikethrough
        .Selection.Font.ForeColor = mvarHeadFontColor
        .SelStart = Len(.Text)
    End With
End Sub
Public Sub FootTextToFile()
'将页脚文字构建成Rtf文件,适用于没有Rtf文件时
Dim strFoot As String
    With RTBFoot
        strFoot = mvarFoot
        Do While strFoot <> ""
            If Right(strFoot, 2) = vbCrLf Then
                strFoot = Mid(strFoot, 1, Len(strFoot) - 2)
            ElseIf Asc(Right(strFoot, 1)) = 13 Then
                strFoot = Mid(strFoot, 1, Len(strFoot) - 1)
            ElseIf Asc(Right(strFoot, 1)) = 10 Then
                strFoot = Mid(strFoot, 1, Len(strFoot) - 1)
            Else
                Exit Do
            End If
        Loop
            
        .Text = strFoot
        If Trim(strFoot) = "" Then Exit Sub
        .SelectAll
        .Selection.Font.Name = mvarFootFontName
        .Selection.Font.SIZE = mvarFootFontSize
        .Selection.Font.Bold = mvarFootFontBold
        .Selection.Font.Italic = mvarFootFontItalic
        .Selection.Font.Underline = mvarFootFontUnderline
        .Selection.Font.Strikethrough = mvarFootFontStrikethrough
        .Selection.Font.ForeColor = mvarFootFontColor
        .SelStart = Len(.Text)
    End With
End Sub
Public Function ShowPages(Optional bFillData As Boolean = True)
'页面初始化
    Dim M As Long, H As Long, Hi As Long, k As Long, i As Long, W As Long, Wi As Long
    H = ScaleHeight - picHRuler.Height - HS.Height
    Hi = (PAGEMARGIN + mvarPaperHeight) * mvarZoomFactor
    k = Hi / VSTEP
    M = CInt(H / Hi) + 2
    If bFillData Then
        For i = 2 To RTBPaper.UBound
            Unload RTBPaper(i)
            Unload picShadow(i)
        Next
        Progress1.Cls
        Progress1.Visible = True
        For i = 2 To mvarPageCount
            If RTBPaper.UBound < i Then Load RTBPaper(i)         '动态生成所有页面
            If picShadow.UBound < i Then Load picShadow(i)
            RTBPaper(i).Visible = False
            RTBPaper(i).PageNumber = i
            picShadow(i).Visible = False
            Progress1.Value = i / (mvarPageCount + 1)        '显示进度条
        Next
        For i = 1 To mvarPageCount
            FillPage (i)
        Next
        Progress1.Visible = False
        PubInfo.PaperCount = mvarPageCount
    End If
    VS.Max = (Hi * mvarPageCount + PAGEMARGIN * mvarZoomFactor - H) / VSTEP
    VS.Tag = (mvarCurPage - 1) * k
    VS.LargeChange = WHEELNUMBER
    VS.Value = (mvarCurPage - 1) * k
    
    W = ScaleWidth - VS.Width
    Wi = (2 * PAGEMARGIN + mvarPaperWidth) * mvarZoomFactor
    k = Wi / HSTEP
    HS.LargeChange = WHEELNUMBER
    HS.Max = (Wi - W) / HSTEP
    HS.Tag = 0
    
    If Wi < W Then
        HS.Value = 0
    Else
        Dim j As Long, lLeft As Long
        lLeft = (W - Wi) / 2
        If lLeft < 0 Then
            lLeft = (mvarMarginLeft - 360) * mvarZoomFactor
        End If
        lLeft = lLeft + 200
        HS.Value = IIf(lLeft / HSTEP >= HS.Max, 0, lLeft / HSTEP)
    End If
    
    Call VS_Change
    Call HS_Change
End Function

Public Function FillPage(Index As Long) As Boolean
    '页面内容填充
    RTBPaper(Index).objPaper.Cls
    RTBPaper(Index).Width = PaperWidth
    RTBPaper(Index).Height = PaperHeight
    PrintPage Index, RTBPaper(Index).objPaper, True
    RTBPaper(Index).DrawBorder
End Function

Public Function LockAllOLEObjectSize() As Boolean
    InProcessing = True
    LockAllOLEObjectSize = RTBNormal.LockAllOLEObjectSize
    InProcessing = False
End Function

Public Function LockOLEObjectSize(ByVal Index As Long) As Boolean
    LockOLEObjectSize = RTBNormal.LockOLEObjectSize(Index)
End Function

Public Function GBtoBIG5(ByVal strText As String) As String
    '简体转繁体
    GBtoBIG5 = J2F(strText)
End Function

Public Function Big5toGB(ByVal strText As String) As String
    '繁体转简体
    Big5toGB = F2J(strText)
End Function

Public Sub RefreshTargetDC()
    '刷新所见即所得与打印机的绑定
    Dim lngTargetDC As Long
    If Not ExistsPrinter Then
        lngTargetDC = picBuff.hDC
    Else
        lngTargetDC = Printer.hDC
    End If
    gTargetDC = lngTargetDC
    ResetWYSIWYG
End Sub


