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
      Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
            Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "����ͼ�ı༭�ؼ�"
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
'##ģ �� ����Editor.ctl
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��1��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ��������ӿڵ����ձ༭�ؼ�����װ����ͨ��ҳ�漰�����ͼģʽ��ӳ�������Ϣ�����ԡ�
'##��    ����
'######################################################################################

Option Explicit

'#############################################################################################################
'##     �ֲ�����
'#############################################################################################################

Private m_hWnd As Long                  '�ؼ��� hWnd
Private m_hWndParent  As Long           '������� hWnd
'Private m_TOM As New cTextDocument      'TOM 3.0 ģ�ͣ����Ķ���
Private mfrmFindText As New frmFindText '�����滻����

Private m_sText As String               '�ı�����

Private Const BORDERWIDTH = 15         '�߿�

'#############################################################################################################
'##     ��������
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
Private mvarPageCount As Long               '��������
Private mvarCurPage As Long                 '��ǰҳ
Private mvarStartPage As Long               'ʵ����ʾ����ʼҳ
Private mvarEndPage As Long                 'ʵ����ʾ����ֹҳ
Private mvarWithViewButtonas As Boolean     '�Ƿ����л���ͼ��ť
Private mvarPaperKind As PaperKindEnum      'ֽ����������
Private mvarPaperOrient As PaperOrientEnum  'ֽ�ŷ���
Private mvarInProcessing As Boolean         '��ҳ������...
Private mvarShowRuler As Boolean            '�Ƿ���ʾ���

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
'##     �¼�����������ӳ���¼���
'#############################################################################################################

Public Event Change(ViewMode As ViewModeEnum)    '���ݸı䣡
Public Event MouseWheel(ViewMode As ViewModeEnum, bBackDirection As Boolean, Shift As Integer, x As Single, y As Single, Value As Single)   '�������¼�
Public Event Zoom(ViewMode As ViewModeEnum, NewFactor As Double)   '�û�ͨ��Ctrl��������ı������ű�����
Public Event Resize(ViewMode As ViewModeEnum)    '�ؼ��ߴ�ı�
Public Event RequestLine(ViewMode As ViewModeEnum)              '���������ı�
Public Event SelChange(ViewMode As ViewModeEnum, ByVal lStart As Long, ByVal lEnd As Long)  'ѡ������ı�
Public Event LinkEvent(ViewMode As ViewModeEnum, ByVal iType As LinkEventTypeEnum, ByVal lStart As Long, ByVal lEnd As Long)     '�����¼�
Public Event ModifyProtected(ViewMode As ViewModeEnum, ByRef bAllowDoIt As Boolean, ByVal lStart As Long, ByVal lEnd As Long, KeyAscii As Integer, Shift As Integer)            '��ͼ�༭�ܱ�������
Public Event BeforeKeyDown(ViewMode As ViewModeEnum, KeyCode As Integer, Shift As Integer)
Public Event KeyDown(ViewMode As ViewModeEnum, KeyCode As Integer, Shift As Integer)
Public Event KeyPress(ViewMode As ViewModeEnum, KeyAscii As Integer)
Public Event KeyUp(ViewMode As ViewModeEnum, KeyCode As Integer, Shift As Integer)
Public Event MouseDown(ViewMode As ViewModeEnum, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(ViewMode As ViewModeEnum, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(ViewMode As ViewModeEnum, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event RequestRightMenu(ViewMode As ViewModeEnum, Shift As Integer, x As Single, y As Single)
Public Event Click(ViewMode As ViewModeEnum)        '����
Public Event DblClick(ViewMode As ViewModeEnum)     '˫��
Public Event PressTabKey()                          '����TAB��ť
Public Event GetDelCharColor(ByRef COLOR As OLE_COLOR)     '��ȡɾ���ַ�����ɫ
Public Event GetNewCharColor(ByRef COLOR As OLE_COLOR)     '��ȡ�����ַ�����ɫ
Public Event IsDelCharColor(ByVal COLOR As OLE_COLOR, ByRef blnIsDelCharColor As Boolean)   '�ж��Ƿ���ɾ���ַ�����ɫ
Public Event IsNewCharColor(ByVal COLOR As OLE_COLOR, ByRef blnIsNewCharColor As Boolean)   '�ж��Ƿ��������ַ�����ɫ
Public Event UIOpen(ByRef UIhWnd As Long, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long) '��UI�ӿ�
Public Event UIMoved(ByRef UIhWnd As Long, lngLeft As Long, lngTop As Long, lngWidth As Long, lngHeight As Long) '��UI�ӿ�
Public Event UIClose(ByRef UIhWnd As Long)  '�ر�UI�ӿ�
Public Event UIClick(ViewMode As ViewModeEnum)   '�ر�UI�ӿ�

'#############################################################################################################
'##     �������ԣ�����ӳ�䣩
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
    '����ˢ�¡����������á���ʾ
    gTargetDC = picBuff.hDC     '�����ҽԺԤ��ʱ�ұ߳��������⣡��ֻ������ĻΪ������
    RTBNormal.ResetWYSIWYG
    RTBHead.ResetWYSIWYG
    RTBFoot.ResetWYSIWYG
End Sub

Public Property Let AuditMode(ByVal vData As Boolean)   '��˼���
    RTBNormal.AuditMode = vData
    PropertyChanged "AuditMode"
End Property

Public Property Get AuditMode() As Boolean              '��˼���
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
    '�������״̬�µ�ѡ�е��޶��ı�
    If Me.AuditMode Then RTBNormal.ResetAuditText
End Sub

Public Sub AcceptAuditText()
    '�������״̬�µ�ѡ�е��޶��ı�
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
'�ɵ����ϼ�����ɾ���ļ�
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
'�ɵ����ϼ�����ɾ���ļ�
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
    'ˢ�¹������ԣ�ҳ��ģʽ��
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
    gTargetDC = picBuff.hDC     '�����ҽԺԤ��ʱ�ұ߳��������⣡��ֻ������ĻΪ������
    
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
    '��ʾUI�ӿ���������ǰ���ݱ�����ͼƬ��
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
        
        '��ȡ��ȷ�ĸ߶ȺͿ�ȣ�����OLE�ӿڣ�
        SendMessage RTBNormal.hWndRTB, EM_GETOLEINTERFACE, 0, mIRichEditOle
        If ObjPtr(mIRichEditOle) = 0 Then
            CloseUIInterface
            Exit Sub
        End If
        '���oleobject����Ϣ
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
    'ˢ��UI�ӿ�������λ�ã�����ƫ�Ƶ������
    '��ʾUI�ӿ���������ǰ���ݱ�����ͼƬ��
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
        
        '��ȡ��ȷ�ĸ߶ȺͿ�ȣ�����OLE�ӿڣ�
        SendMessage RTBNormal.hWndRTB, EM_GETOLEINTERFACE, 0, mIRichEditOle
        '���oleobject����Ϣ
        mReObject.cbStruct = LenB(mReObject)
        mIRichEditOle.GetObject REO_IOB_SELECTION, mReObject, REO_GETOBJ_ALL_INTERFACES
        Set mIOleObject = mReObject.poleobj
        If Not mIOleObject Is Nothing Then
            mIOleObject.GetExtent DVASPECT_CONTENT, pSize
            lWidth = UserControl.ScaleX(pSize.cx, vbHimetric, vbTwips)       'ͼƬԭʼ��С
            lHeight = UserControl.ScaleY(pSize.cy, vbHimetric, vbTwips)
        End If
        SendMessage Me.hWndRTB, EM_POSFROMCHAR, VarPtr(pt), ByVal mReObject.cP
        lLeft = pt.x * Screen.TwipsPerPixelX + RTBNormal.Left
        lTOp = pt.y * Screen.TwipsPerPixelY + RTBNormal.Top + lngSpaceBefore * 1.3333333
        'ͼƬ���մ�С
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
    '��ʾUI�ӿ���������ǰ���ݱ�����ͼƬ��
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
        
        '��ȡ��ȷ�ĸ߶ȺͿ�ȣ�����OLE�ӿڣ�
        SendMessage RTBNormal.hWndRTB, EM_GETOLEINTERFACE, 0, mIRichEditOle
        If ObjPtr(mIRichEditOle) = 0 Then
            CloseUIInterface
            Exit Sub
        End If
        '���oleobject����Ϣ
        mReObject.cbStruct = LenB(mReObject)
        mIRichEditOle.GetObject REO_IOB_SELECTION, mReObject, REO_GETOBJ_ALL_INTERFACES
        Set mIOleObject = mReObject.poleobj
        If Not mIOleObject Is Nothing Then
            mIOleObject.GetExtent DVASPECT_CONTENT, pSize
            lWidth = UserControl.ScaleX(pSize.cx, vbHimetric, vbTwips)       'ͼƬԭʼ��С
            lHeight = UserControl.ScaleY(pSize.cy, vbHimetric, vbTwips)
        Else
            CloseUIInterface
            Exit Sub
        End If
        SendMessage Me.hWndRTB, EM_POSFROMCHAR, VarPtr(pt), ByVal mReObject.cP
        lLeft = pt.x * Screen.TwipsPerPixelX + RTBNormal.Left
        lTOp = pt.y * Screen.TwipsPerPixelY + RTBNormal.Top + lngSpaceBefore * 1.3333333
        'ͼƬ���մ�С
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
    '����ʽ����
    RTBNormal.CopyWithFormat
End Sub

Public Sub PasteWithFormat()
    '����ʽ����
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
    '���ܣ����ĵ���ǰλ��������ָ���ַ������鵽��ѡ��
    '������
    '   sText,Ҫ���ҵ��ַ���
    '   iFlag,ƥ�䷽ʽ,Ĭ��Ϊ0(�����ִ�Сд��ȫ���)������Ϊ���±�־����ϣ�
    '       tomMatchCase,2-��Сдƥ��
    '       tomMatchWord,4-��ȫƥ��
    '       ʵ�ʲ��ԣ��в�֧��ģʽƥ���
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
'����ͼƬ��ָ��λ��
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
'����ͼƬ��ָ��λ��
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
        ResizeReObject rtbBuff, lWidth, lHeight     '����ͼƬ�ߴ�
        
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

'##################################     ��Ƕ�Ի���      ##################################

Public Function ShowFontDlg(Optional intFlags As Integer) As Boolean
    '���ܣ���ʾ����Ի��򣬿��Ըı�����,�ֺ�,����,б�壬���ݲ��������Ƿ�������Ч��������ԣ�ֻ����ͨģʽ����
    '������
    '   intFlags,�Ƿ��ֹ��صĸ���Ч��ѡ�
    '       intFlags and (2^0) <> 0,��ֹ����ɾ��������
    '       intFlags and (2^1) <> 0,��ֹ���ı�������
    '       intFlags and (2^2) <> 0,��ֹ������������
    '       intFlags and (2^3) <> 0,��ֹ�����»�������
    '       intFlags and (2^4) <> 0,��ֹ����ǰ��ɫ����
    '       intFlags and (2^5) <> 0,��ֹ���ı���ɫ����
    
    Dim strSample As String
    
    If Me.ViewMode <> cprNormal Then Exit Function
    strSample = Trim(Me.Selection.Text)
    If strSample <> "" Then strSample = Left(Split(strSample, vbCrLf)(0), 10)
    If strSample <> Trim(Me.Selection.Text) Then strSample = strSample & "��"
    
    Me.ForceEdit = True
    ShowFontDlg = frmFontSetup.ShowMe(TOM, intFlags, strSample)
    Me.ForceEdit = False
End Function

Public Function ShowPageSetupDlg(Optional intFlags As Integer) As Boolean
    '���ܣ���ʾҳ�����öԻ���
    '������
    '   intFlags,�Ƿ��ֹ��صĸ���Ч��ѡ�
    '       intFlags and (2^0) <> 0,��ֹ����ҳ�汳��ɫ����
    '       intFlags and (2^1) <> 0,��ֹ�����ĵ�����ɫ����
    
    If frmPageSetup.ShowMe(Me, intFlags) Then
        ShowPageSetupDlg = True
        Call UserControl_Resize
        
'        '��ͨģʽ�����������õ���������
'        RTBNormal.ResetWYSIWYG
'
'        'ҳ��ģʽ�Ļ���Ҫ���·�ҳ
        If mvarViewMode = cprPaper Then ViewMode = cprPaper
        Me.Modified = True
    End If

End Function

Public Function ShowParaDlg(Optional blnHideStyle As Boolean) As Boolean
    '���ܣ���ʾ�����ʽ�Ի���ֻ����ͨģʽ����
    '������blnHideStyle-�Ƿ��ֹ�����ʽ����
    
    Dim strText As String, strSample As String, lS As Long, lE As Long, i As Long
    
    If Me.ViewMode <> cprNormal Then Exit Function
        
    '��ȡ�������֣��Ա���Ϊʾ��
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
    '���ܣ���ʾ��Ŀ���źͱ�ŶԻ���
    ShowItemNumberDlg = frmItemNumber.ShowMe(Me.Selection.Para)
End Function

Public Function ShowCharCountDlg() As Boolean
    '���ܣ���ʾ����ͳ�ƶԻ���
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
    '���ܣ���ʾ��������ʱ��Ի���
    '������
    '   blnDelay,Ϊ��ʱ����ֱ�Ӳ����޸ı༭����SelText���ݣ�ֻ��������ֵ��
    '   MinDate,�������С����
    '   MaxDate,������������
    '���أ����õ�����ʱ���ַ�����ȡ��ʱ����""
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
            '���������ԣ����������ı���
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
    '���ܣ���ʾ������ź������ַ��Ի���
    '������
    '   bSaveInEditor,Ϊ��ʱ�� ֱ�Ӳ���/�޸ı༭����SelText���ݣ����򷵻�����ֵ��
    '   bytSex,�Ա�0-ûָ��;1-����;2-Ů��
    '   blnReturnStr �Ƿ����ַ���ʽ����,��ʾ��ǰλ�ò�֧��ͼƬ��ʽ����.=true ���ַ�����
    '   strInfor �༭ͼƬʱ�����������Ϣ���༭���ش�
    '            ��ʽΪ������|���ݡ��¾�ʷ 1|ǰ�|����|��ĸ|���|�ֺ�; ���� 2(����)/3(����)|����|����|����|����|�ֺ�; ̥��λ�� 4|�Ϸ�|�·�|��|�ҷ�|�ֺ�
    '   objPic   �༭�������ɵ�ͼƬ�ش�
    '���أ�ֻҪִ�й������򷵻�True��ֱ�ӹرշ���False
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
            '���������ԣ����������ı���
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
    '���ܣ���ʾҳüҳ�ŶԻ���
    ShowHeadFootDlg = frmHeadFoot.ShowMe(Me)
End Function

Public Function ShowFindReplaceDlg(Optional intShowWhat As Integer) As Boolean
    '���ܣ���ʾ�����滻�Ի���ִ�в����滻���滻ʱ�����Ա��������ص����ݽ����滻��ҳ��ģʽ���ṩ��
    '������
    '   intShowWhat,��ʾ�ͽ�ֹ�Ĺ���:
    '    0,������ʾ���Ҵ���
    '    1,������ʾ�滻����
    '   -1,��ʾ���Ҵ��������滻����
    If mvarViewMode <> cprNormal Then Exit Function
    ShowFindReplaceDlg = mfrmFindText.ShowMe(Me, intShowWhat)
End Function

Public Sub FindNext()
    '���ܣ� ����һ�¸�
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
        
        '���ܳ�����Χ
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
                'ѡ�ж������
                lS = RTBNormal.Selection.StartPos
                lE = RTBNormal.Selection.EndPos
                strText = RTBNormal.Text
                For i = lS To lE
                    j = InStr(i + 1, strText, vbCrLf)
                    If j = 0 Then
                        'û�з��ֻس�
                        Exit For
                    ElseIf j <= lE Then
                        '��Χ�ڷ��ֻس�
                        i = j + 1
                        lCur = j - 1
                        RTBNormal.TOM.TextDocument.Range(lCur, lCur).Para.ClearAllTabs
                        For k = 0 To TabCount - 1
                            If TabPos(k) > 0 Then
                                RTBNormal.TOM.TextDocument.Range(lCur, lCur).Para.AddTab TabPos(k) / 20, TabAlign(k), tomSpaces
                            End If
                        Next k
                    Else
                        '��Χ��û�з��ֻس���ȡ��ĩλ�������Ʊ�λ
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
                'ѡ�е�������
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
    GetCursorPos R1  '��ȡ��ǰ���λ��
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
        GetCursorPos R1  '��ȡ��ǰ���λ��
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
    HRuler.FirstLineIndent = lF * 20        '��ֵ��羽���Ϊ20��
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
    '���Tab����
    Dim iRetVal As Integer
    iRetVal = GetKeyState(VK_SHIFT)
    ' ���û�а�shift�����tab
    If iRetVal <> -128 And iRetVal <> -127 Then
        iRetVal = GetKeyState(VK_TAB)
        If iRetVal = -128 Or iRetVal = -127 Then ' tab������
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
        GetCursorPos R1  '��ȡ��ǰ���λ��
        If x <= mvarMarginLeft * mvarZoomFactor Then
            R2.x = R1.x + (mvarMarginLeft * mvarZoomFactor - x) / Screen.TwipsPerPixelX
            If y <= mvarMarginTop * mvarZoomFactor Then
                '�����ϱ߾�
                R2.y = R1.y + (mvarMarginTop * mvarZoomFactor - y) / Screen.TwipsPerPixelY
            ElseIf y >= (mvarPaperHeight - mvarMarginBottom) * mvarZoomFactor Then
                '�����±߾�
                R2.y = R1.y + ((mvarPaperHeight - mvarMarginBottom) * mvarZoomFactor - y) / Screen.TwipsPerPixelY
            Else
                R2.y = R1.y
            End If
            SetCursorPos R2.x, R2.y
            Call mouse_event(MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&)
            Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, 0&)
        Else
            If y <= mvarMarginTop * mvarZoomFactor Then
                '�����ϱ߾�
                R2.x = R1.x
                R2.y = R1.y + (mvarMarginTop * mvarZoomFactor - y) / Screen.TwipsPerPixelY
            ElseIf y >= (mvarPaperHeight - mvarMarginBottom) * mvarZoomFactor Then
                '�����±߾�
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
    gTargetDC = picBuff.hDC     '�����ҽԺԤ��ʱ�ұ߳��������⣡��ֻ������ĻΪ������
    If lngTargetDC <> gTargetDC Then ResetWYSIWYG
End Sub

Private Sub UserControl_Initialize()
'�ڳ��򴴽��ؼ�������ʱʱ����                                           '���㵥λ56.6857142857143

    '�����ơ��߶ȡ���ȡ���С�߾�(��������)����Ӧ��ӡֽ�����е�ֽ�����ೣ��
    PaperKindConst(1) = "�ż� 8 1/2��11 Ӣ��                        ,15842,12242,482,799,181,181,1"
    PaperKindConst(2) = "+A611 С���ż� 8 1/2��11 Ӣ��              ,15842,12242,482,799,181,181,2"
    PaperKindConst(3) = "С�ͱ� 11��17 Ӣ��                         ,24477,15842,482,799,181,181,3"
    PaperKindConst(4) = "������ 17��11 Ӣ��                         ,15842,24477,482,799,181,181,4"
    PaperKindConst(5) = "�����ļ� 8 1/2��14 Ӣ��                    ,20163,12242,482,799,181,181,5"
    PaperKindConst(6) = "������5 1/2��8 1/2 Ӣ��                    ,12242,7919,482,799,181,181,6"
    PaperKindConst(7) = "�����ļ�7 1/2��10 1/2 Ӣ��                 ,15122,10438,482,799,181,181,7"
    PaperKindConst(8) = "A3 297��420 ����                           ,23814,16840,482,799,181,193,8"
    PaperKindConst(9) = "A4 210��297 ����                           ,16840,11907,482,805,181,176,9"
    PaperKindConst(10) = "A4С�� 210��297 ����                      ,16840,11907,482,805,181,176,9"
    PaperKindConst(11) = "A5 148��210 ����                          ,11907,8392,482,799,181,176,11"
    PaperKindConst(12) = "B4 250��354 ����                          ,20067,14171,482,805,181,181,12"
    PaperKindConst(13) = "B5 182��257 ����                          ,14572,10319,482,805,181,176,13"
    PaperKindConst(14) = "�Կ��� 8 1/2��13 Ӣ��                     ,18722,12242,482,799,181,181,14"
    PaperKindConst(15) = "�Ŀ��� 215��275 ����                      ,15589,12187,482,805,181,176,15"
    PaperKindConst(16) = "10��14 Ӣ��                               ,20163,14398,482,805,181,176,16"
    PaperKindConst(17) = "11��17 Ӣ��                               ,24477,15842,482,805,181,176,17"
    PaperKindConst(18) = "����8 1/2��11 Ӣ��                        ,15842,12242,482,805,181,176,18"
    PaperKindConst(19) = "#9 �ŷ� 3 7/8��8 7/8 Ӣ��                 ,5579,12780,482,794,181,176,19"
    PaperKindConst(20) = "#10 �ŷ� 4 1/8��9 1/2 Ӣ��                ,5936,13682,482,794,181,181,20"
    PaperKindConst(21) = "#11 �ŷ� 4 1/2��10 3/8 Ӣ��               ,14938,6479,482,794,181,181,21"
    PaperKindConst(22) = "#12 �ŷ� 4 1/2��11 Ӣ��                   ,15842,6479,482,794,181,181,22"
    PaperKindConst(23) = "#14 �ŷ� 5��11 1/2 Ӣ��                   ,16558,7199,482,794,181,181,23"
    PaperKindConst(24) = "C �ߴ繤����                              ,16558,7199,482,794,181,181,24"
    PaperKindConst(25) = "D �ߴ繤����                              ,16558,7199,482,794,181,181,25"
    PaperKindConst(26) = "E �ߴ繤����                              ,16558,7199,482,794,181,181,26"
    PaperKindConst(27) = "DL ���ŷ� 110��220 ����                   ,6237,12474,482,805,181,181,27"
    PaperKindConst(28) = "C5 ���ŷ� 162��229 ����                   ,9185,12984,482,799,181,176,28"
    PaperKindConst(29) = "C3 ���ŷ� 324��458 ����                   ,25969,18371,482,794,181,176,29"
    PaperKindConst(30) = "C4 ���ŷ� 229��324 ����                   ,18371,12981,482,794,181,176,30"
    PaperKindConst(31) = "C6 ���ŷ� 114��162 ����                   ,9183,6462,482,794,181,176,31"
    PaperKindConst(32) = "C65 ���ŷ�114��229 ����                   ,12981,6462,482,794,181,176,32"
    PaperKindConst(33) = "B4 ���ŷ� 250��353 ����                   ,20010,14171,482,794,181,176,33"
    PaperKindConst(34) = "B5 ���ŷ�176��250 ����                    ,9979,14175,482,799,181,193,34"
    PaperKindConst(35) = "B6 ���ŷ� 176��125 ����                   ,7086,9976,482,799,181,193,35"
    PaperKindConst(36) = "�ŷ� 110��230 ����                        ,13037,6237,482,799,181,193,36"
    PaperKindConst(37) = "�ŷ���� 3 7/8��7 1/2 Ӣ��                ,5579,10801,482,794,181,181,37"
    PaperKindConst(38) = "�ŷ� 3 5/8��6 1/2 Ӣ��                    ,9359,5219,482,794,181,181,38"
    PaperKindConst(39) = "U.S. ��׼��д�� 14 7/8��11 Ӣ��           ,15842,21421,0,0,0,1848,39"
    PaperKindConst(40) = "�¹���׼��д�� 8 1/2��12 Ӣ��             ,17282,12242,0,0,0,0,40"
    PaperKindConst(41) = "�¹����ɸ�д�� 8 1/2��13 Ӣ��             ,18722,12242,0,0,0,0,41"
    PaperKindConst(42) = "�Զ���ֽ��                                ,22680,16443,482,0,0,0,256"
    PageCount = 1
    mvarCurPage = 1

    lblThis.Caption = "����ͼ�ı༭�ؼ� v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & App.LegalCopyright & " " & App.LegalTrademarks
End Sub

Private Sub UserControl_InitProperties()
'����������ʵ��ʱ�������������Ե������ʼ�����룡���������û��ڴ����Ϸ���һ���ؼ�ʱ�������¼�������ʱ���ٴ�������
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
    Title = "δ�����ĵ�"
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
    HeadFontName = "����"
    HeadFontSize = 10
    HeadFontBold = False
    HeadFontItalic = False
    HeadFontUnderline = False
    HeadFontStrikethrough = False
    HeadFontColor = vbBlack
    HeadFile = ""
    FootFontName = "����"
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
'�����ؾ��б���״̬�Ķ���ľ�ʵ��ʱ���������¼���
'���Զ�ȡ����̬���ԵĶ�ȡ���Ӷ�ת��Ϊ��̬���ԣ���ʱ����pInitialise������ʼ���������
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
    Title = PropBag.ReadProperty("Title", "δ�����ĵ�")
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
    HeadFontName = PropBag.ReadProperty("HeadFontName", "����")
    HeadFontSize = PropBag.ReadProperty("HeadFontSize", 10)
    HeadFontBold = PropBag.ReadProperty("HeadFontBold", False)
    HeadFontItalic = PropBag.ReadProperty("HeadFontItalic", False)
    HeadFontUnderline = PropBag.ReadProperty("HeadFontUnderline", False)
    HeadFontStrikethrough = PropBag.ReadProperty("HeadFontStrikethrough", False)
    HeadFontColor = PropBag.ReadProperty("HeadFontColor", vbBlack)
    HeadFile = PropBag.ReadProperty("HeadFile", "")
    FootFontName = PropBag.ReadProperty("FootFontName", "����")
    FootFontSize = PropBag.ReadProperty("FootFontSize", 10)
    FootFontBold = PropBag.ReadProperty("FootFontBold", False)
    FootFontItalic = PropBag.ReadProperty("FootFontItalic", False)
    FootFontUnderline = PropBag.ReadProperty("FootFontUnderline", False)
    FootFontStrikethrough = PropBag.ReadProperty("FootFontStrikethrough", False)
    FootFontColor = PropBag.ReadProperty("FootFontColor", vbBlack)
    FootFile = PropBag.ReadProperty("FootFile", "")
    If Ambient.UserMode Then
        '��ȡĬ�ϵ�ҳ������
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
        gTargetDC = picBuff.hDC     '�����ҽԺԤ��ʱ�ұ߳��������⣡��ֻ������ĻΪ������
    End If
    
    ResetWYSIWYG
    
    SetVSWithRtb True
    
    Modified = False    '������Ӧ�÷ŵ���󣬱���ViewModeʹ�����ݸı䡣
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
                '����ˮƽ�����������ֵ
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
            ShowPages False 'ˢ�� VS.MAX �� VS.Value����ˢ������
            picHRulerHead.Width = 390
        End Select
        
        HRuler.Width = mvarPaperWidth
        btnNormal.Move 0, UserControl.ScaleHeight - btnNormal.Height
        btnPaper.Move btnNormal.Left + btnNormal.Width, btnNormal.Top
        HS.Move btnPaper.Left + btnPaper.Width, btnNormal.Top, UserControl.ScaleWidth - btnPaper.Width * 2 - picNull.Width
        picNull.Move ScaleWidth - picNull.Width, ScaleHeight - picNull.Height
        Progress1.Move IIf(ScaleWidth > 4500, ScaleWidth - Progress1.Width - 500, 1000), ScaleHeight - HS.Height + 15, IIf(ScaleWidth > 4500, 2000, Abs(ScaleWidth - 1500))   '������������λ��
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
'����������ʵ��ʱ���������¼������¼�֪ͨ�����ʱ��Ҫ��������״̬���Ա㽫���ɻָ���״̬�����������£������״̬����������ֵ��
'���Ա��棨��̬���Եı��棩
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
    PropBag.WriteProperty "Title", Title, "δ�����ĵ�"
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
    PropBag.WriteProperty "HeadFontName", HeadFontName, "����"
    PropBag.WriteProperty "HeadFontSize", HeadFontSize, 10
    PropBag.WriteProperty "HeadFontBold", HeadFontBold, False
    PropBag.WriteProperty "HeadFontItalic", HeadFontItalic, False
    PropBag.WriteProperty "HeadFontUnderline", HeadFontUnderline, False
    PropBag.WriteProperty "HeadFontStrikethrough", HeadFontStrikethrough, False
    PropBag.WriteProperty "HeadFontColor", HeadFontColor, vbBlack
    PropBag.WriteProperty "HeadFile", HeadFile, ""
    PropBag.WriteProperty "FootFontName", FootFontName, "����"
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
        '��ͨģʽ
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
        '��ͨģʽ
        Dim Pos As POINTAPI, lngMax As Long
        lngMax = (mvarPaperWidth - MarginRight - RTBNormal.OriginRTB.Width) / Screen.TwipsPerPixelX
        HS.Max = lngMax
        SendMessage hWndRTB, EM_GETSCROLLPOS, 0, Pos
        Pos.x = HS.Value
        Pos.y = Pos.y
        SendMessage hWndRTB, EM_SETSCROLLPOS, 0, Pos
        HRuler.Left = picMarginL.Left
    ElseIf mvarViewMode = cprPaper Then
        'ҳ��ģʽ
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
'## ���ܣ�  ��ӡ����ҳ�浽ָ���豸����ӡ��/ͼƬ��
'##
'## ������  PageNumber      ��ҳ��
'##         objTarget       ����ӡ��Ŀ��ؼ���Printer/ͼƬ��
'##         blnPreview      ���Ƿ���Ԥ��ģʽ�������Ԥ��ģʽ����ôҳüҳ����ɫΪ��ɫ����ʽ��ӡ�Ǻ�ɫ��
'##         lngBlankHeight      ���ⲿָ�����ϲ����׸߶�
'############################################################################################################
Public Sub PrintPage(ByVal PageNumber As Long, Optional ByRef objTarget As Object = Nothing, _
    Optional ByVal blnPreview As Boolean = False, Optional ByVal lngBlankHeight As Long = 0, _
    Optional ByVal blnMarginReverse As Boolean)
Dim lngOffsetLeft As Long, lngOffsetTop As Long      '���Եƫ����'�ϱ�Եƫ����
Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
Dim lngPicWidth As Long, lngPicHeight As Long   'ͼƬ��ȸ߶�
Dim lngPageCount As Long        '��ҳ��
Dim lngHead As Long, lngFoot As Long          '���ڶ���ҳü�ĸ߶�'���ڶ���ҳ�ŵĸ߶�
Dim fr As FORMATRANGE           '��ʽ�����ı���Χ
Dim rcDrawTo As RECT            'Ŀ����������
Dim rcPage As RECT              'Ŀ��ҳ������
Dim Rct As RECT                 '��ӡҳüҳ��
Dim lngNextPos As Long          '��һ���ַ�λ��
Dim strHead As String, strFoot As String    'ҳüҳ��

    If objTarget Is Nothing Then Set objTarget = Printer
    objTarget.ScaleMode = vbTwips   '���ô�ӡ����λΪ羡�
    
    'ͼƬ�߶ȺͿ��
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
    gTargetDC = picBuff.hDC     '�����ҽԺԤ��ʱ�ұ߳��������⣡��ֻ������ĻΪ������
    
    '��ȡ��ӡ���ɴ�ӡ����ı�Եƫ��������λ��Pixel
    lngOffsetLeft = objTarget.ScaleX(GetDeviceCaps(objTarget.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = objTarget.ScaleY(GetDeviceCaps(objTarget.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
    
    If lngPicHeight > 0 Then '��ӡҳüͼƬ
        If blnMarginReverse Then 'Ҫ�����ұ߾෴��֧��˫���ӡ
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
    '����ҳü�߶�
    RTBTmp.PaperWidth = mvarPaperWidth: RTBTmp.MarginLeft = mvarMarginLeft: RTBTmp.MarginRight = mvarMarginRight: RTBTmp.ResetWYSIWYG
    RTBTmp.TextRTF = HeadFileTextRTF
    Call DocTmpReplaceKey("", "", blnPreview) '��Ҫ���滻�����еĹؼ���
    Call DocTmpReplaceKey("{ҳ��}", PageNumber)       'ҳ����滻
    With rcDrawTo
        If blnMarginReverse Then 'Ҫ�����ұ߾෴��֧��˫���ӡ
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
    Call SendMessage(RTBTmp.hWndRTB, EM_FORMATRANGE, 1, fr)  '��ӡҳü
    lngHead = fr.rc.Bottom - fr.rc.Top
    If lngHead <= 0 Then lngHead = 0
    If RTBHead.Text <> "" Or lngPicHeight > 0 Then
        objTarget.ForeColor = IIf(blnPreview, RGB(149, 149, 149), vbBlack)
        objTarget.Line (fr.rc.Left, fr.rc.Bottom + 50)-(fr.rc.Right, fr.rc.Bottom + 50)
    End If

    RTBTmp.TextRTF = FootFileTextRTF
    Call DocTmpReplaceKey("", "", blnPreview) '��Ҫ���滻�����еĹؼ���
    Call DocTmpReplaceKey("{ҳ��}", PageNumber)       'ҳ����滻
    Call SendMessage(RTBTmp.hWndRTB, EM_FORMATRANGE, 0, fr)
    lngFoot = fr.rc.Bottom - fr.rc.Top
    If lngFoot <= 0 Then lngFoot = 0
    
    '���ÿɴ�ӡ��������
    If blnMarginReverse Then 'Ҫ�����ұ߾෴��֧��˫���ӡ
        lngLeft = (mvarMarginRight - lngOffsetLeft)
        lngRight = (mvarPaperWidth - lngOffsetLeft) - IIf(mvarMarginLeft >= lngOffsetLeft, mvarMarginLeft, lngOffsetLeft)
    Else
        lngLeft = (mvarMarginLeft - lngOffsetLeft) '�߾�Ӧ���Ѿ�������ӡƫ��
        lngRight = (mvarPaperWidth - lngOffsetLeft) - IIf(mvarMarginRight >= lngOffsetLeft, mvarMarginRight, lngOffsetLeft)
    End If
    lngTop = (mvarMarginTop - lngOffsetTop) + lngPicHeight + lngHead + IIf(lngHead > 0, 350, 0)
    lngBottom = mvarPaperHeight - mvarMarginBottom - lngFoot

    rcDrawTo.Left = lngLeft
    rcDrawTo.Top = lngTop
    rcDrawTo.Right = lngRight
    rcDrawTo.Bottom = lngBottom
    
    '���ô�ӡָ�FormatRange��Ϣ��Ҫ�Ĵ�ӡ��Ϣ��
    fr.hDC = objTarget.hDC          ' ��������Ⱦʹ����ͬ��DC
    fr.hdcTarget = gTargetDC        ' Ŀ��ؼ���DC���ؼ�����
    fr.rc = rcDrawTo                ' ���־������� IN/OUT
    fr.rcPage = rcPage              ' ����ҳ��������� IN
    fr.chrg.cpMin = AllPages(PageNumber).Start ' ��ӡ��������ֿ�ʼλ��
    fr.chrg.cpMax = AllPages(PageNumber).End   ' ���ֽ���λ�ã�-1��ʾֱ��ĩβ��
    
    '��Ϊʵ�ʴ�ӡʱ����λ��������ƫ�ƣ����¼������ָ߶�
    If lngBlankHeight > lngTop Then  '���ֵ����ֲ���
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
        '���������������ӡ��,Bottom�ᷢ��ƫ�ƣ��ó�������Bottomλ�ü�Ϊ���ָ߶�
        lngblankPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, frBlank)  '�����ӡ
        'ֻ����1��2��ʱ��rich�ؼ��㲻׼����ʵ�����ּ�
        lngBlankHeight = IIf((rcBlank.Bottom - rcBlank.Top) <= 600, lngBlankHeight - lngOffsetTop, frBlank.rc.Bottom + 350)
    Else
        lngBlankHeight = lngBlankHeight
    End If
    
    '���� EM_FORMATRANGE ��Ϣ���д�ӡ
    lngNextPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, fr)  '�����ӡ
    fr.rc = rcDrawTo
    If lngNextPos < AllPages(PageNumber).End Then fr.rc.Bottom = fr.rc.Bottom + 99999    '��֤һ�δ�ӡ����ҳ��
    lngNextPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 1, fr)  'ʵ�ʴ�ӡ
    

    If lngBlankHeight < 200 Then '������ʱ����ҳ�ţ���Ϊ����һ��ֽ�ϴ���
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
        
    '�����ϲ����׾���
    If lngBlankHeight > 0 Then
        objTarget.PaintPicture picBlank.Image, 0, 0, mvarPaperWidth, lngBlankHeight
    End If
    
    '����RTF�ͷ��ڴ�
    Call SendMessage(RTBTmp.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBHead.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Call SendMessage(RTBFoot.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
End Sub

'############################################################################################################
'## ���ܣ�  ��ȡ�ı��߶ȣ�ͨ�������ӡ��ȡ��
'##
'## ������  PageNumber      ��ҳ��
'##         objTarget       ����ӡ��Ŀ��ؼ���Printer/ͼƬ��
'##         blnPreview      ���Ƿ���Ԥ��ģʽ�������Ԥ��ģʽ����ôҳüҳ����ɫΪ��ɫ����ʽ��ӡ�Ǻ�ɫ��
'##         lngBlankHeight      ���ⲿָ�����ϲ����׸߶�
'############################################################################################################
Private Function GetPrintHeight() As Long
    Dim fr As FORMATRANGE           '��ʽ�����ı���Χ
    Dim rcDrawTo As RECT            'Ŀ����������
    Dim rcPage As RECT              'Ŀ��ҳ������
    Dim lngNextPos As Long          '��һ���ַ�λ��
    Dim strHead As String, strFoot As String    'ҳüҳ��
    Dim r As Long                   '����ֵ
    
    picBuff.ScaleMode = vbTwips   '���ô�ӡ����λΪ羡�
    
    '���ÿɴ�ӡҳ������
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = mvarPaperWidth
    rcPage.Bottom = 999999999#
    
    '���ÿɴ�ӡ��������
    rcDrawTo.Left = mvarMarginLeft
    rcDrawTo.Top = mvarMarginTop
    rcDrawTo.Right = mvarPaperWidth - mvarMarginRight
    rcDrawTo.Bottom = 999999999#
    fr.hDC = picBuff.hDC            ' ��������Ⱦʹ����ͬ��DC
    fr.hdcTarget = picBuff.hDC      ' Ŀ��ؼ���DC���ؼ�����
    fr.rc = rcDrawTo                ' ���־������� IN/OUT
    fr.rcPage = rcPage              ' ����ҳ��������� IN
    fr.chrg.cpMin = 0               ' ��ӡ��������ֿ�ʼλ��
    fr.chrg.cpMax = -1              ' ���ֽ���λ�ã�-1��ʾֱ��ĩβ��
    
    '���� EM_FORMATRANGE ��Ϣ���д�ӡ
    lngNextPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, fr)  '�����ӡ
    If fr.rc.Bottom = 999999999# Then fr.rc.Bottom = mvarPaperHeight
    GetPrintHeight = (fr.rc.Bottom - fr.rc.Top - UserControl.Height)
    If GetPrintHeight < mvarPaperHeight - mvarMarginTop - mvarMarginBottom Then
        GetPrintHeight = mvarPaperHeight - mvarMarginTop - mvarMarginBottom
    End If
    GetPrintHeight = GetPrintHeight / Screen.TwipsPerPixelY
End Function

'############################################################################################################
'## ���ܣ�  ����rtb��Scroll��С��λ������VS�Ĵ�С��λ��
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
'## ���ܣ�  ִ�������ӡ������RTF�ķ�ҳ��Ϣ
'############################################################################################################
Public Sub DoVirtualPrint()
Dim lngPicWidth As Long, lngPicHeight As Long   'ͼƬ��ȸ߶�
Dim lngHead As Long, lngFoot As Long            'ҳüҳ�Ÿ߶�
Dim lngPageCount As Long        '��ҳ��
Dim fr As FORMATRANGE           '��ʽ�����ı���Χ
Dim rcDrawTo As RECT            'Ŀ����������
Dim rcPage As RECT              'Ŀ��ҳ������
Dim lngNextPos As Long          '��һ���ַ�λ��
Dim r As Long                   '����ֵ

    On Error Resume Next
    lngPicWidth = 0: lngPicHeight = 0 'ͼƬ�߶ȺͿ��
    If Not (mvarPicture Is Nothing) Then
        If mvarPicture.Handle <> 0 Then
            lngPicWidth = UserControl.ScaleX(mvarPicture.Width, vbHimetric, vbTwips)
            lngPicHeight = UserControl.ScaleX(mvarPicture.Height, vbHimetric, vbTwips)
        End If
    End If
    
    picBuff.Width = mvarPaperWidth: picBuff.Height = mvarPaperHeight
    picBuff.ScaleMode = vbTwips   '���ô�ӡ����λΪ羡�
    
    '���ÿɴ�ӡҳ������
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
    gTargetDC = picBuff.hDC     '�����ҽԺԤ��ʱ�ұ߳��������⣡��ֻ������ĻΪ������
    
    '����ҳü�߶�
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
    
    '���ÿɴ�ӡ��������
    rcDrawTo.Left = mvarMarginLeft
    rcDrawTo.Top = mvarMarginTop + lngPicHeight + lngHead + IIf(lngHead > 0, 350, 0)
    rcDrawTo.Right = mvarPaperWidth - mvarMarginRight
    rcDrawTo.Bottom = mvarPaperHeight - mvarMarginBottom - lngFoot
    
    '���ô�ӡָ�FormatRange��Ϣ��Ҫ�Ĵ�ӡ��Ϣ��
    fr.hDC = picBuff.hDC            ' ��Ⱦ�豸
    fr.hdcTarget = gTargetDC        ' Ŀ���豸���ؼ�����
    fr.rc = rcDrawTo                ' ���־������� IN/OUT
    fr.rcPage = rcPage              ' ����ҳ��������� IN
    fr.chrg.cpMin = 0               ' ��ӡ��������ֿ�ʼλ��
    fr.chrg.cpMax = -1              ' ���ֽ���λ�ã�-1��ʾֱ��ĩβ��
    
    '��ȡ����RTF�ı�����
    Dim lngTmp As Long              '���ڼ�¼��ҳ�ַ���ʼλ��
    Dim lngLen As Long              '�ܳ��ȣ���Ӣ�Ļ�ϳ��ȣ�
    lngLen = lstrlen(RTBNormal.Text)
    
    'ѭ����ҳ��ӡ
    Do
        '���� EM_FORMATRANGE ��Ϣ���������ӡ
        lngNextPos = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, fr)     'ֻ��ҳ������ӡ
        
        lngPageCount = lngPageCount + 1             ' ҳ����1
        '��¼��ҳ��Ϣ
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          'ʵ�ʴ�ӡ�߶�
        AllPages(lngPageCount).Start = lngTmp
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos                      ' ��һҳ��ʼ�ַ�λ��
        fr.hDC = picBuff.hDC
        fr.hdcTarget = gTargetDC
        fr.rc = rcDrawTo                                ' �������������������򣬷�������
        If lngNextPos <= lngTmp Or lngNextPos >= lngLen Then Exit Do      ' �������ҳ��ķ�ҳ
        lngTmp = lngNextPos
    Loop
    PageCount = lngPageCount
    AllPages(lngPageCount).End = -1                     ' ���һҳ����λ��Ϊ��ĩβ
    
    '����RTF�ͷ��ڴ�
    r = SendMessage(RTBHead.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    r = SendMessage(RTBFoot.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    r = SendMessage(RTBNormal.hWndRTB, EM_FORMATRANGE, 0, ByVal CLng(0))
    Err.Clear
End Sub

'############################################################################################################
'## ���ܣ�  ��ӡ��ǰ�ĵ�����ӡ��
'##
'## ������  blnNoAsk            ���Ƿ���ʾ��ӡ�Ի����ڴ�ӡǰ��������
'##         lngStartPage        ���ⲿָ������ʼҳ
'##         lngBlankHeight      ���ⲿָ�����ϲ����׸߶�
'##         lngCopies           ��ָ����ӡ��������0��ָ���������ƣ�������ƴ�ӡ���������޸ģ���ͨ���������ش�ӡ����
'############################################################################################################
Public Function PrintDoc(Optional ByVal blnNoAsk As Boolean, Optional ByVal lngStartPage As Long, Optional ByVal lngBlankHeight As Long, _
    Optional ByRef strPrinterDeviceName As String, Optional ByRef lngCopies As Long = 0) As Boolean
    
    Dim strOldPrinterName As String
    
    If Not ExistsPrinter Then MsgBox "û�а�װ��ӡ�豸�����ܴ�ӡ��", vbExclamation, App.Title: Exit Function
    If mvarViewMode <> cprPaper Then
        Me.InProcessing = True
        '��ȡ��ҳ��Ϣ
        DoVirtualPrint
        Me.InProcessing = False
    End If
    
    Dim intPageFrom As Integer, intPageTo As Integer, bytPageOddEven As Byte
    Dim blnCopyOrder As Boolean, blnDuplex As Boolean, blnCurReverse As Boolean
    Dim t As Variant, aryPage() As String, i As Long, j As Long, k As Long, L As Long, M As Long
    Dim lngPageCount As Long
    Dim Pages() As Long             '��ӡ��Χ�ڵ��������ӡ��ҳ��
    Dim blnRangePrint As Boolean    '�Ƿ���ҳ�뷶Χ��ӡ
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
            If strPrinterDeviceName <> "" Then '��Ϊ�п���ָ���Ĵ�ӡ�������ڣ�����ʹ��ֱ��=��ʽ
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
                'ҳ�뷶Χ
                blnRangePrint = True
                t = Split(.txtPageScope.Tag, ",")
                For i = 0 To UBound(t)
                    aryPage = Split(t(i), "-")
                    If UBound(aryPage) = 0 Then
                        'ֻ��һҳ
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
                '��ǰҳ
                intPageFrom = Me.CurPage: intPageTo = Me.CurPage
            Else
                'ȫ����ӡ
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
        '����ҳ
        If intPageFrom Mod 2 = 0 Then intPageFrom = intPageFrom + 1
    ElseIf bytPageOddEven = 2 Then
        'ż��ҳ
        If intPageFrom Mod 2 = 1 Then intPageFrom = intPageFrom + 1
    End If
    If intPageFrom > intPageTo Then Exit Function
    
    Dim lngCount As Long, lngNumber As Long
    Err = 0: On Error Resume Next
     
    '���ô�ӡ������
    If Printer.Orientation <> Me.PaperOrient Then
        Printer.Orientation = Me.PaperOrient
    End If
    '����ֽ�ţ��Զ���ֽ�ŵ����ñ���ŵ����
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
    gTargetDC = picBuff.hDC     '�����ҽԺԤ��ʱ�ұ߳��������⣡��ֻ������ĻΪ������
    
    '��ʼ��ӡ
    Printer.Print Space(1)
    
    If blnCopyOrder = True Then
        '��ݴ�ӡ
        For lngNumber = 1 To lngCopies
            blnCurReverse = True
            For lngCount = intPageFrom To intPageTo Step IIf(bytPageOddEven = 0, 1, 2)
                If blnRangePrint Then
                    'ҳ�뷶Χ��ӡ
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
                                blnCurReverse = True              '��ѡ˫���ӡʱ��ż��ÿҳ ���ұ߾෴��
                            ElseIf bytPageOddEven = 0 Then
                                blnCurReverse = Not blnCurReverse '��ѡ˫���ӡ����ӡȫ������ʱ��ÿ���һҳ������ҳ�߾෴��
                            Else
                                blnCurReverse = False
                            End If
                        Else
                            blnCurReverse = False
                        End If
                    Else
                        blnCurReverse = False 'û��ѡ˫���ӡʱ��������
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
                    'ҳ�뷶Χ��ӡ
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
                                blnCurReverse = Not blnCurReverse '��ѡ˫���ӡ����ӡȫ������ʱ��ÿ���һҳ������ҳ�߾෴��
                            ElseIf bytPageOddEven = 2 Then
                                blnCurReverse = True              '��ѡ˫���ӡʱ��ż��ÿҳ ���ұ߾෴��
                            Else
                                blnCurReverse = False
                            End If
                        Else
                            blnCurReverse = False
                        End If
                    Else
                        blnCurReverse = False 'û��ѡ˫���ӡ����ҳ��ӡ���ʱ��������
                    End If
                    PrintPage lngCount, Printer, , IIf(lngCount = lngStartPage, lngBlankHeight, 0), blnCurReverse
                    blnFirstPrinted = True
                End If
            Next
        Next
    End If
    
    Printer.EndDoc
    
    If blnNoAsk = False Then
        '�ָ�Ĭ�ϴ�ӡ��
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
    gTargetDC = picBuff.hDC     '�����ҽԺԤ��ʱ�ұ߳��������⣡��ֻ������ĻΪ������
    
    PrintDoc = True
    Exit Function
    Err.Clear
PrintErr:
    PrintDoc = False
End Function
Public Sub DocTmpReplaceKey(Optional ByVal strSource As String = "", Optional ByVal strTraget As String = "", Optional ByVal blnPreview As Boolean)
'�滻��ҳü/ҳ���еĹؼ���
Dim strR As String, strUnitName As String, lngS As Long, lngE As Long
Dim strFontName As String, sinFontSize As Single, blnBlod As Boolean, blnItalic As Boolean, blnUnderline As Boolean
    On Error Resume Next
    If strSource <> "" Then
        RTBTmp.Range(0, 0).Selected
        If RTBTmp.FindText(strSource, 4) Then
            With RTBTmp
                '��ȡԭ������
                lngS = .Selection.StartPos: lngE = .Selection.EndPos
                strFontName = .Range(lngS, lngE).Font.Name
                sinFontSize = .Range(lngS, lngE).Font.SIZE
                blnBlod = .Range(lngS, lngE).Font.Bold
                blnUnderline = .Range(lngS, lngE).Font.Underline
                blnItalic = .Range(lngS, lngE).Font.Italic
                '�滻
                .Range(lngS, lngE) = strTraget
                '���ִ�ԭ������
                lngE = lngS + Len(strTraget)
                .Range(lngS, lngE).Font.Name = strFontName
                .Range(lngS, lngE).Font.SIZE = sinFontSize
                .Range(lngS, lngE).Font.Bold = blnBlod
                .Range(lngS, lngE).Font.Underline = blnUnderline
                .Range(lngS, lngE).Font.Italic = blnItalic
            End With
        End If
    Else
        strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
        Call DocTmpReplaceKey("{��λ����}", strUnitName)
        Call DocTmpReplaceKey("{��ҳ��}", PubInfo.PaperCount)
        Call DocTmpReplaceKey("{����}", PubInfo.Title)
        Call DocTmpReplaceKey("{·��}", Left(PubInfo.FileName, InStrRev(PubInfo.FileName, "\")))
        Call DocTmpReplaceKey("{�ļ���}", Mid(PubInfo.FileName, InStrRev(PubInfo.FileName, "\") + 1))
        If Not blnPreview Then
            Call DocTmpReplaceKey("{��ӡ����}", Format(Now(), "yyyy��mm��dd��"))
            Call DocTmpReplaceKey("{��ӡʱ��}", Format(Now(), "hh:MM:ss"))
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
    '����ʽ����
    RTBHead.ForceEdit = True
    RTBHead.Freeze
    RTBHead.SelectAll
    DoEvents
    RTBHead.PasteWithFormat
    RTBHead.UnFreeze
    RTBHead.ClearEndCrlfChar
End Sub
Public Sub DocHeadReplaceKey(Optional ByVal strSource As String = "", Optional ByVal strTraget As String = "", Optional ByVal blnPreview As Boolean)
'�滻��ҳü/ҳ���еĹؼ���
Dim strR As String, strUnitName As String, lngS As Long, lngE As Long
Dim strFontName As String, sinFontSize As Single, blnBlod As Boolean, blnItalic As Boolean, blnUnderline As Boolean, lngColor As Long
    On Error Resume Next
    If strSource <> "" Then
        RTBHead.Range(0, 0).Selected
        If RTBHead.FindText(strSource) Then
            With RTBHead
                If Not .ForceEdit Then .ForceEdit = True
                '��ȡԭ������
                lngS = .Selection.StartPos: lngE = .Selection.EndPos
                strFontName = .Range(lngS, lngE).Font.Name
                sinFontSize = .Range(lngS, lngE).Font.SIZE
                blnBlod = .Range(lngS, lngE).Font.Bold
                blnUnderline = .Range(lngS, lngE).Font.Underline
                blnItalic = .Range(lngS, lngE).Font.Italic
                lngColor = .Range(lngS, lngE).Font.ForeColor
                '�滻
                .Range(lngS, lngE) = strTraget
                '���ִ�ԭ������
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
        strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
        Call DocHeadReplaceKey("{��λ����}", strUnitName)
        Call DocHeadReplaceKey("{����}", PubInfo.Title)
        Call DocHeadReplaceKey("{·��}", Left(PubInfo.FileName, InStrRev(PubInfo.FileName, "\")))
        Call DocHeadReplaceKey("{�ļ���}", Mid(PubInfo.FileName, InStrRev(PubInfo.FileName, "\") + 1))
        If Not blnPreview Then
            Call DocHeadReplaceKey("{��ӡ����}", Format(Now(), "yyyy��mm��dd��"))
            Call DocHeadReplaceKey("{��ӡʱ��}", Format(Now(), "hh:MM:ss"))
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
    '����ʽ����
    RTBFoot.ForceEdit = True
    RTBFoot.Freeze
    RTBFoot.SelectAll
    DoEvents
    RTBFoot.PasteWithFormat
    RTBFoot.ClearEndCrlfChar
    RTBFoot.UnFreeze
End Sub
Public Sub DocFootReplaceKey(Optional ByVal strSource As String = "", Optional ByVal strTraget As String = "", Optional ByVal blnPreview As Boolean)
'�滻��ҳü/ҳ���еĹؼ���
Dim strR As String, strUnitName As String, lngS As Long, lngE As Long
Dim strFontName As String, sinFontSize As Single, blnBlod As Boolean, blnItalic As Boolean, blnUnderline As Boolean, lngColor As Long
    On Error Resume Next
    If strSource <> "" Then
        RTBFoot.Range(0, 0).Selected
        If RTBFoot.FindText(strSource) Then
            With RTBFoot
                If Not .ForceEdit Then .ForceEdit = True
                '��ȡԭ������
                lngS = .Selection.StartPos: lngE = .Selection.EndPos
                strFontName = .Range(lngS, lngE).Font.Name
                sinFontSize = .Range(lngS, lngE).Font.SIZE
                blnBlod = .Range(lngS, lngE).Font.Bold
                blnUnderline = .Range(lngS, lngE).Font.Underline
                blnItalic = .Range(lngS, lngE).Font.Italic
                lngColor = .Range(lngS, lngE).Font.ForeColor
                '�滻
                .Range(lngS, lngE) = strTraget
                '���ִ�ԭ������
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
        strUnitName = GetSetting("ZLSOFT", "ע����Ϣ", "��λ����", "")
        Call DocFootReplaceKey("{��λ����}", strUnitName)
        Call DocFootReplaceKey("{����}", PubInfo.Title)
        Call DocFootReplaceKey("{·��}", Left(PubInfo.FileName, InStrRev(PubInfo.FileName, "\")))
        Call DocFootReplaceKey("{�ļ���}", Mid(PubInfo.FileName, InStrRev(PubInfo.FileName, "\") + 1))
        If Not blnPreview Then
            Call DocFootReplaceKey("{��ӡ����}", Format(Now(), "yyyy��mm��dd��"))
            Call DocFootReplaceKey("{��ӡʱ��}", Format(Now(), "hh:MM:ss"))
        End If
    End If
    Err.Clear
End Sub
Public Sub HeadTextToFile()
'��ҳü���ֹ�����Rtf�ļ�,������û��Rtf�ļ�ʱ
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
'��ҳ�����ֹ�����Rtf�ļ�,������û��Rtf�ļ�ʱ
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
'ҳ���ʼ��
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
            If RTBPaper.UBound < i Then Load RTBPaper(i)         '��̬��������ҳ��
            If picShadow.UBound < i Then Load picShadow(i)
            RTBPaper(i).Visible = False
            RTBPaper(i).PageNumber = i
            picShadow(i).Visible = False
            Progress1.Value = i / (mvarPageCount + 1)        '��ʾ������
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
    'ҳ���������
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
    '����ת����
    GBtoBIG5 = J2F(strText)
End Function

Public Function Big5toGB(ByVal strText As String) As String
    '����ת����
    Big5toGB = F2J(strText)
End Function

Public Sub RefreshTargetDC()
    'ˢ���������������ӡ���İ�
    Dim lngTargetDC As Long
    If Not ExistsPrinter Then
        lngTargetDC = picBuff.hDC
    Else
        lngTargetDC = Printer.hDC
    End If
    gTargetDC = lngTargetDC
    ResetWYSIWYG
End Sub


