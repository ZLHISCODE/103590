VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.UserControl usrTendFileMutilEditor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   8565
   Begin VB.PictureBox pic过滤条件 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   450
      ScaleHeight     =   315
      ScaleWidth      =   7575
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   30
      Width           =   7575
      Begin VB.CommandButton cmd刷新 
         Caption         =   "刷新(&R)"
         Height          =   315
         Left            =   6660
         TabIndex        =   28
         Top             =   0
         Width           =   885
      End
      Begin VB.ComboBox cbo科室 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
      Begin VB.CheckBox chk出院 
         Caption         =   "出院"
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   5940
         TabIndex        =   26
         ToolTipText     =   "勾选表示提取出院病人"
         Top             =   60
         Width           =   675
      End
      Begin VB.CheckBox chk出科 
         Caption         =   "出科"
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   5190
         TabIndex        =   25
         ToolTipText     =   "勾选表示提取出科病人"
         Top             =   60
         Width           =   675
      End
      Begin VB.ComboBox cbo护理文件格式 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   0
         Width           =   2205
      End
      Begin VB.Label lbl科室 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   3180
         TabIndex        =   23
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lbl文件格式 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "文件格式"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   60
         TabIndex        =   21
         Top             =   60
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   6150
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileMutilEditor.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrTendFileMutilEditor.ctx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3090
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   8385
      TabIndex        =   10
      Top             =   510
      Width           =   8385
      Begin VB.CommandButton cmdWord 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         Picture         =   "usrTendFileMutilEditor.ctx":0734
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1290
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6330
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picDnInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.Label lblDnInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   60
               TabIndex        =   18
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.PictureBox picUpInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.Label lblUpInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   60
               TabIndex        =   17
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.TextBox txtDnInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   525
            MaxLength       =   12
            TabIndex        =   7
            Top             =   30
            Width           =   345
         End
         Begin VB.TextBox txtUpInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   30
            MaxLength       =   12
            TabIndex        =   6
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   435
            TabIndex        =   14
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   1
         ItemData        =   "usrTendFileMutilEditor.ctx":0A76
         Left            =   6660
         List            =   "usrTendFileMutilEditor.ctx":0A8C
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1590
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5790
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   1290
         Visible         =   0   'False
         Width           =   615
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "√"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   12
            Top             =   30
            Width           =   315
         End
      End
      Begin VB.PictureBox picMutilInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   5790
         ScaleHeight     =   435
         ScaleWidth      =   1575
         TabIndex        =   8
         Top             =   3330
         Visible         =   0   'False
         Width           =   1600
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   810
            TabIndex        =   9
            Top             =   90
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "体温体录"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   15
            TabIndex        =   11
            Top             =   112
            Width           =   720
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   0
         ItemData        =   "usrTendFileMutilEditor.ctx":0AC4
         Left            =   5790
         List            =   "usrTendFileMutilEditor.ctx":0ADA
         TabIndex        =   3
         Top             =   1590
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   0
         Top             =   930
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrTendFileMutilEditor.ctx":0B12
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一般护理记录单"
         Height          =   180
         Left            =   3720
         TabIndex        =   27
         Top             =   90
         Width           =   1275
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrTendFileMutilEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private mfrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '是否显示录入框
Private mblnBlowup As Boolean               '放大否？放大1/3，如字体9号放大为12号
Private mblnChange As Boolean               '是否修改数据
Private mblnSaved As Boolean                '是否已保存
Private mblnSigned As Boolean               '是否已签名
Private mstrData As String                  '进入编辑状态前保存之前的数据
Private mintPreDays As Long
Private mstrMaxDate As String

Private mlng文件ID As Long
Private mlng格式ID As Long
Private mlng科室ID As Long
Private mlng病区ID As Long
Private mint页码 As Integer
Private mstrPrivs As String

Private mdtOutEnd As Date
Private mdtOutbegin As Date
Private mintChange As Integer

Private mintSymbol As Integer               '当前控件索引
Private mstrSymbol As String                '特殊字符
Private mstrCollectItems As String          '汇总项目集合
Private mstrColCollect As String            '汇总项目列集合:col;1|col;4,5
Private mstrCOLNothing As String            '未绑定的列集合+活动项目列(不管活动项目列是否绑定)
Private mstrCOLActive As String             '活动列集合
Private mstrCatercorner As String           '列对角线集合
Private mblnEditAssistant As Boolean        '当前选择的项目是否允许进行词句选择
Private mlngPageRows As Long                '此文件格式一页所显示的数据行
Private mlngOverrunRows As Long             '超出数据行
Private mlngRowCount As Long                '当前记录总行数
Private mlngRowCurrent As Long              '当前记录在本页的实际行数
Private mlngDate As Long                    '日期
Private mlngTime As Long                    '时间
Private mlngOperator As Long                '护士
Private mlngSignLevel As Long               '签名级别
Private mlngSigner As Long                  '签名信息
Private mlngSignName As Long                '签名人
Private mlngSignTime As Long                '签名时间
Private mlngRecord As Long                  '记录ID
Private mlngNoEditor As Long                '禁止编辑列,存在护士列则以护士列为准,不存在护士列则以签名列为准

Private mintType As Integer                 '记录当前的编辑模式
Private mblnDateAd As Boolean               '日期缩写?
Private mstr开始时间 As String              '当前文件的开始时间
Private mstr结束时间 As String              '当前文件的结束时间
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '所有护理记录项目清单
Private mrsSelItems As New ADODB.Recordset          '当前录入的护理记录项目清单
Private mrsDataMap As New ADODB.Recordset           '当前操作员录入的数据镜像,与记录单格式一致,相关行数据全部保存以便迅速恢复
Private mrsCellMap As New ADODB.Recordset           '编辑过的数据镜像,字段有:页号,行号,列号,记录ID,数据,部位,删除
Private mrsCopyMap As New ADODB.Recordset           '复制行数据

Private Enum ColIcon
    签名 = 1
    审签 = 2
End Enum
Private Enum SignLevel
    正高 = 1
    副高 = 2
    中级 = 3
    师级 = 4
    员士 = 5
    未定义 = 9
End Enum

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'病历文件格式定义相关
Private mintTabTiers As Integer     '表头层次
Private mintTagFormHour As Integer  '开始时间条件
Private mintTagToHour As Integer    '截止时间条件
Private mobjTagFont As New StdFont  '条件样式字体
Private mlngTagColor As Long        '条件样式颜色
Private mstrPaperSet As String      '格式
Private mstrPageHead As String      '页眉
Private mstrPageFoot As String      '页脚
Private mblnChildForm As Boolean
Private mstrSubhead As String       '表上标签
Private mstrTabHead As String       '表头单元
Private mstrColWidth As String      '列宽序列串
Private mstrColumns As String       '当前护理文件各列对应的项目
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'保存打开护理记录文件的SQL，在其它地方也有使用，不能修改
Private mstrSQL内 As String
Private mstrSQL中 As String
Private mstrSQL列 As String
Private mstrSQL条件 As String
Private mstrSQL As String

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与绘图相关,没事别动
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const WHITE_BRUSH = 0                   '白色画笔
Private Const cdblWidth As Double = 6           '一个英文字符的宽度
Private Const cHideCols = 6                     '前缀列:床号,姓名
Private Const cControlFields = 2                '记录集控制列:页号,行号
Private Const c文件ID As Integer = 1
Private Const c床号 As Integer = 2
Private Const c姓名 As Integer = 3
Private Const c病人ID As Integer = 4
Private Const c主页ID As Integer = 5
Private Const c婴儿 As Integer = 6
Private Const p住院护士站 = 1262

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '将VB的颜色转换为RGB表示
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long

    OleTranslateColor dwOleColour, 0, clrref

    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF

    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    '缺省是宋体9号,按字体大小同比放大
    Dim sinFontSize As Single
    Dim i As Integer, j As Integer

    j = Len(strPara)
    sinFontSize = VsfData.FontSize
    For i = 1 To j
        GetSymbolWidth = GetSymbolWidth + IIf(Asc(Mid(strPara, i, 1)) > 0, 1, 2) * cdblWidth * sinFontSize / 9
    Next
End Function

Private Sub DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '绘图相关
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim t_ClientRect As RECT
    On Error GoTo errHand
    '******************************************
    '在此事件中不能对单元格的任何属性赋值,包括Celldata,否则会引起该事件的死循环,导致工具栏或计时器无法正常工作。
    '******************************************
    '使用匹配的背景色，前景色与字体进行文本输出。
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False

    strText = VsfData.TextMatrix(ROW, COL)
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '赋初值
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        '取字符宽度
        dblWidth = GetSymbolWidth(strRight)
        '设定客户区域大小
        With t_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With

        '1、清空内容
        '创建与背景色相同的刷子
        If ROW < VsfData.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(VsfData.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(VsfData.ForeColorFixed)
        Else
            If ROW = VsfData.RowSel Then
                lngBackColor = GetRBGFromOLEColor(VsfData.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(VsfData.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        '使用该刷子填充背景色
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, t_ClientRect, lngBrush)
        '立即销毁临时使用的刷子并还原刷子
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)

        '2、准备画线
        '创建新画笔
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '画线
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '输出文本
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)

        '还原画笔并销毁
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)

        '已完成作图
        Done = True
    End If
'
'    '3、如果是汇总行，则进行特殊处理
'    If Val(VsfData.TextMatrix(Row, mlngCollectType)) < 0 And Val(VsfData.TextMatrix(Row, mlngCollectStyle)) = 1 _
'        And (Col >= mlngDate And Col < mlngNoEditor) Then
'        Call DrawCollectCell(hDC, Left, Top, Right, Bottom)
'    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与分行相关,没事别动
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long

    GetData = ""
    lngRows = SendMessage(txtLength.hwnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        Call SendMessage(txtLength.hwnd, EM_GETLINE, lngRow - 1, strLine(0))
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|ZYB.ZLSOFT|") & strData & IIf(lngRow < lngRows, vbCrLf, "")
    Next
    GetData = Split(GetData, "|ZYB.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = 0 To intMax
        strLine(intDo) = 0
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long

    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

'**********************************************************************************************************************
'######################################################################################################################

Private Sub ReadStruDef()
    Dim lngCOL As Long
    On Error GoTo errHand

    '读取文件属性
    mblnDateAd = False

    '提取活动项目并加入列定义(格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""

    '读取病历文件格式定义
    gstrSQL = "Select  /*+ RULE */ d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历文件格式定义", mlng格式ID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数": mintTabTiers = Val("" & !内容文本)
            Case "总列数":  VsfData.Cols = Val("" & !内容文本)
            Case "最小行高": VsfData.RowHeightMin = BlowUp(Val("" & !内容文本))
            Case "文本字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set Font = objFont
            Case "文本颜色": VsfData.ForeColor = Val("" & !内容文本)
            Case "表格颜色": VsfData.GridColor = Val("" & !内容文本): VsfData.GridColorFixed = VsfData.GridColor
            
            Case "标题文本"
                lblTitle.Caption = "" & !内容文本
                lblTitle.AutoSize = True
            Case "标题字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False

            Case "开始时间": mintTagFormHour = Val("" & !内容文本)
            Case "终止时间": mintTagToHour = Val("" & !内容文本)
            Case "条件字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "条件颜色": mlngTagColor = Val("" & !内容文本)
            Case "有效数据行"
                mlngOverrunRows = 0
                mlngPageRows = Val(!内容文本)
            End Select
            .MoveNext
        Loop
    End With
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select  /*+ RULE */ d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
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
    Dim strSql外 As String, str格式 As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim bln对角线 As Boolean, bln选择项 As Boolean          '如果上一列是对角线且选择项,则直接提取各项数据,拼列头时在数值间加上/
    Dim lngColumn As Long, blnAddCollect As Boolean

    gstrSQL = "Select  /*+ RULE */ d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表列集合定义", mlng格式ID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = ""
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
                    mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", Mid(strSql外, 3)) & " As C" & Format(lngColumn, "00")
                Else
                    If strSql外 <> "" Then
                        mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql外 = ""
                lngColumn = !对象序号
                bln对角线 = (NVL(!要素表示, 0) = 1)
                bln选择项 = False
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    bln选择项 = (mrsItems!项目表示 = 5)
                    If mrsItems!项目表示 = 4 Then   '汇总项目
                        blnAddCollect = True
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!项目序号
                        mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
                    End If
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !要素名称
                str格式 = str格式 & "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                mrsItems.Filter = "项目名称='" & NVL(!要素名称) & "'"
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!项目表示 = 4 Then   '汇总项目
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!项目序号
                        If blnAddCollect Then
                            mstrColCollect = mstrColCollect & "," & mrsItems!项目序号
                        Else    '有可能一列绑定两个项目,第一个项目不是汇总项目,第二个项目才是汇总项目,因此,下面的代码保证加上列序号
                            mstrColCollect = mstrColCollect & "|" & !对象序号 & ";" & mrsItems!项目序号
                        End If
                    End If
                End If
                mrsItems.Filter = 0
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
                    mstrSQL条件 = mstrSQL条件 & " Or """ & !要素名称 & """ Is Not Null"

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
                    End If

                    If (Trim("" & !内容文本) = "" And Trim("" & !要素单位) = "") Or (bln对角线 And bln选择项) Then
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & !要素名称 & """"
                    Else
                        mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,Decode(c.记录内容,Null,'" & !内容文本 & "'||'" & !要素单位 & "','" & !内容文本 & "'||c.记录内容||'" & !要素单位 & "')), '') As """ & !要素名称 & """"
                    End If
                Else
'                    '为空表示未绑定列,强制加,后面进行替换
                    mstrCOLNothing = mstrCOLNothing & "," & Format(!对象序号, "00")
'                    mstrSQL中 = mstrSQL中 & ",Max(""" & "C" & Format(!对象序号, "00") & """) As C" & Format(!对象序号, "00")
'                    mstrSQL条件 = mstrSQL条件 & " Or """ & "C" & Format(!对象序号, "00") & """ Is Not Null"
'                    mstrSQL内 = mstrSQL内 & ", C" & Format(!对象序号, "00") & " AS C" & Format(!对象序号, "00")
                End If
            End Select
            .MoveNext
        Loop

        If mstrCollectItems <> "" Then
            mstrCollectItems = Mid(mstrCollectItems, 2)
            mstrColCollect = Mid(mstrColCollect, 2)
        End If
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
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

        If Mid(mstrSQL中, 2) = "" Then
            MsgBox "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！", vbInformation, gstrSysName
            Exit Sub
        End If

        '程序内部控制增加固定列
        mstrSQL中 = UCase(mstrSQL中 & ",MAX(签名级别) AS 签名级别,MAX(签名信息) AS 签名信息,MAX(记录ID) AS 记录ID,MAX(行数) AS 行数,MAX(实际行数) AS 实际行数")
        mstrSQL内 = UCase(mstrSQL内 & ",l.签名级别,l.签名人 AS 签名信息,C.记录ID,P.行数||'' AS 行数,1 AS 实际行数")
        mstrSQL列 = UCase(mstrSQL列 & ",签名级别,签名信息,记录ID,行数,实际行数")

        Call SQLCombination
    End With

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SQLCombination(Optional ByVal lng记录ID As Long = 0)
    Dim str条件 As String
    str条件 = mstrSQL条件
    
    mstrSQL = "Select  /*+ RULE */ 0 AS 文件ID,'' AS 床号,'' AS 姓名,0 AS 病人ID,0 AS 主页ID,0 AS 婴儿," & Mid(mstrSQL列, 12) & vbCrLf & _
                " From (Select 记录组号,时间 as 备用,发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select c.记录组号,l.发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f,病人护理打印 p " & vbCrLf & _
                "               Where l.ID=p.记录ID And l.Id = c.记录id And l.文件ID=f.ID And f.ID=p.文件ID " & _
                "               And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And 1=2)" & vbCrLf & _
                IIf(str条件 <> "", "Where " & str条件, "") & _
                "       Group By 日期, 时间, 发生时间,记录组号,护士,签名人,签名时间" & _
                                "       Order By 日期, 时间, 发生时间,记录组号,护士,签名人,签名时间)"
End Sub

Private Sub zlRefresh()
    Err = 0: On Error GoTo errHand
    Call InitCons

    '产生列记录集
    Call InitRecords

    '装入数据
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng文件ID)
    '清除并拷贝记录集结构
    Call DataMap_Init(rsTemp)
    '绑定数据并设置护理记录单的格式,同时实现一行数据分行显示的功能
    Call PreTendFormat(rsTemp)

    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataMap_Init(ByVal rsSource As ADODB.Recordset)
    '初始化内存数据集

    '数据记录集,用于快速恢复
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "页号,行号"
    '修改单元格记录,用于保存
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|页号," & adDouble & ",18|行号," & adDouble & ",18|" & _
            "列号," & adDouble & ",18|记录ID," & adDouble & ",18|数据," & adLongVarChar & ",4000|部位," & adLongVarChar & ",100|" & _
            "汇总," & adDouble & ",1|删除," & adDouble & ",1")
    mrsCellMap.Sort = "页号,行号,列号"
    '复制记录集
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
End Sub

Private Function DataMap_Save() As Boolean
    '将当前页面中用户编辑过的数据保存起来,页面切换或保存前触发
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCOL As Long, lngRows As Long, lngCols As Long
    On Error GoTo errHand
    
    '先删除指定页号的所有数据行
    If mrsDataMap.RecordCount <> 0 Then mrsDataMap.MoveFirst
    Do While True
        If mrsDataMap.EOF Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    
    '复制指定页号的所有数据行
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!页号 = mint页码
        mrsDataMap!行号 = lngRow
        mrsDataMap!删除 = IIf(VsfData.RowHidden(lngRow), 1, 0)
        For lngCOL = 0 To lngCols - VsfData.FixedCols
            mrsDataMap.Fields(cControlFields + lngCOL).Value = IIf(VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols))
        Next
        mrsDataMap.Update
    Next
    
    DataMap_Save = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function DataMap_Restore() As Boolean
    '将指定页面的数据恢复到表格中
    Dim lngRow As Long, lngCOL As Long, lngRows As Long, lngCols As Long
    On Error GoTo errHand
    
    mrsDataMap.MoveFirst
    lngCols = VsfData.Cols - 1
    lngRows = mrsDataMap.RecordCount
    VsfData.Rows = VsfData.FixedRows
    For lngRow = 0 To lngRows - 1
        If lngRow > VsfData.Rows - VsfData.FixedRows - 1 Then VsfData.Rows = VsfData.Rows + 1
        For lngCOL = 0 To lngCols - VsfData.FixedCols
            VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCOL + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCOL).Value)
        Next
        If mrsDataMap!删除 = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        mrsDataMap.MoveNext
    Next
    
    DataMap_Restore = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CellMap_Update(ByVal lngStart As Long, ByVal lngDeff As Long)
    Dim lngPos As Long
    Dim intCol As Integer
    
    '更新当前页面所有大于起始行的行号数据
    With mrsCellMap
        If .RecordCount <> 0 Then .MoveLast
        If .BOF Then Exit Sub
        Do While Not .BOF
            If !页号 = mint页码 And !行号 > lngStart Then
                intCol = !列号
                lngPos = .AbsolutePosition
                !行号 = !行号 + lngDeff
                !ID = mint页码 & "," & !行号 & "," & !列号
                .Update
                .MoveFirst
                .Move lngPos - 2
            Else
                .MovePrevious
            End If
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional ByVal blnAddPage As Boolean = True) As ADODB.Recordset
    '只拷贝记录集的结构,同时增加页号,行号字段
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer

    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If blnAddPage Then
            .Fields.Append "页号", adDouble, 18
            .Fields.Append "行号", adDouble, 18
        End If
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
            .Fields.Append "删除", adDouble, 1
        End If

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    Set CopyNewRec = rsTarget
End Function

Private Sub PreTendMutilRows()
    Dim lngRowCount As Long, lngRowCurrent As Long  '当前记录总行数,当前记录在本页的实际行数
    Dim lngCOL As Long, lngMax As Long
    Dim lngRow As Long
    On Error GoTo errHand

    Dim arrData
    Dim intData As Integer, intDatas As Integer
    '如果一行显示不完则分行显示(根据当前数据占用行数先添加空白行并处理行坐标,然后再依次处理当前行的数据)
    '每页只显示实际的数据行,把'@处取消注释即可

    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If lngRow >= mlngPageRows + mlngOverrunRows + VsfData.FixedRows Then Exit Do
        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
        '@实际数据行
'        lngRowCurrent = Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))

        If lngRowCount > 1 Then
            '先增加空行
            VsfData.Rows = VsfData.Rows + lngRowCount - 1
            '从当前行的下一行开始，每行的位置+所增加的空白行数，保证新增的空白行从当前行的下一行开始
            For intData = VsfData.Rows - lngRowCount To lngRow + 1 Step -1
                VsfData.RowPosition(intData) = intData + lngRowCount - 1
            Next

            '循环处理当前行数据
            For lngCOL = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCOL) And lngCOL <> mlngRowCount Then
                    '循环赋值
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCOL) = VsfData.TextMatrix(lngRow, lngCOL)
                    Next
                ElseIf (lngCOL < mlngNoEditor And lngCOL <> mlngDate And lngCOL <> mlngTime) Then
                    '准备赋值
                    With txtLength
                        .Width = VsfData.ColWidth(lngCOL)
                        .Text = VsfData.TextMatrix(lngRow, lngCOL)
                        .FontName = VsfData.CellFontName
                        .FontSize = VsfData.CellFontSize
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)

                    If intDatas > 0 Then
                        '循环赋值
                        For intData = 0 To intDatas
                            VsfData.TextMatrix(lngRow + intData, lngCOL) = arrData(intData)
                        Next
                    End If
                ElseIf lngCOL = mlngNoEditor Then
                        '将行值改为从1开始,比如有4行数据,就是4|1
                        For intData = 1 To lngRowCount
                            VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                        Next
                    Else
                End If
            Next
            '@实际数据行
'            '如果本页第一行的数据不全,则先将该记录第一行的主数据(日期,时间,签名)信息复制到
'            If lngRow = VsfData.FixedRows And lngRowCount <> lngRowCurrent Then
'                '固定复制显示日期时间与签名列
'                lngMax = lngRowCount - lngRowCurrent
'                If mlngDate > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngDate) = VsfData.TextMatrix(lngRow, mlngDate)
'                If mlngTime > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngTime) = VsfData.TextMatrix(lngRow, mlngTime)
'                if mlngOperator <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngOperator) = VsfData.TextMatrix(lngRow, mlngOperator)
'                if mlngOperator <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngsignname) = VsfData.TextMatrix(lngRow, mlngsignname)
'                '删除多余的行
'                For lngCol = 1 To lngMax
'                    VsfData.RemoveItem lngRow
'                Next
'            End If
'            lngRow = lngRow + lngRowCurrent - 1 '加上该记录在本页实际的行数
            '@实际数据行要注释下面这行代码
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
    Loop
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String
    Dim lngRow As Long, lngCOL As Long, lngCount As Long, strCell As String
    On Error GoTo errHand

    '设置护理记录单的格式
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '表头填写
        .MergeCells = flexMergeFixedOnly
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True

        '程序内部控制列隐藏\
        .ColHidden(c文件ID) = True
        .ColHidden(c病人ID) = True
        .ColHidden(c主页ID) = True
        .ColHidden(c婴儿) = True
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColWidth(0) = 250
        .ColWidth(c姓名) = 1500
        .ColAlignment(c床号) = flexAlignRightCenter

        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&

        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCOL = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCOL + cHideCols + .FixedCols - 1) = strCell
        Next
        '设置固定列及选择列
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, c文件ID) = "文件ID"
        .TextMatrix(1, c文件ID) = "文件ID"
        .TextMatrix(2, c文件ID) = "文件ID"
        .TextMatrix(0, c床号) = "床号"
        .TextMatrix(1, c床号) = "床号"
        .TextMatrix(2, c床号) = "床号"
        .TextMatrix(0, c姓名) = "姓名"
        .TextMatrix(1, c姓名) = "姓名"
        .TextMatrix(2, c姓名) = "姓名"
        .TextMatrix(0, c病人ID) = "病人ID"
        .TextMatrix(1, c病人ID) = "病人ID"
        .TextMatrix(2, c病人ID) = "病人ID"
        .TextMatrix(0, c主页ID) = "主页ID"
        .TextMatrix(1, c主页ID) = "主页ID"
        .TextMatrix(2, c主页ID) = "主页ID"
        .TextMatrix(0, c婴儿) = "婴儿"
        .TextMatrix(1, c婴儿) = "婴儿"
        .TextMatrix(2, c婴儿) = "婴儿"

        '列宽设置
        Dim blnAlign As Boolean
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next

        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next

        If .Rows = .FixedRows Then
            mlngOverrunRows = 0
        Else
            '得到第一行的超出行
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            '加上最后一行的超出行
            mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
        End If

        'Call PreTendMutilRows
        Call FillPage
        Call WriteColor
        
        .ROW = .FixedRows
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long
    '晚班以红色显示，同时将非起始行设置为NoCheckBox，设置图标
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 0) <> "" Then
                '晚班以红色显示
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 0)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 0)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If

            '将非起始行设置为NoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If VsfData.TextMatrix(lngCount, mlngRowCount) Like "*|1" Then
                    '设置图标
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(审签).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(签名).Picture
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    Call SetActiveColColor
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long

    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    VsfData.Move lngScaleLeft + 210, lblTitle.Top + lblTitle.Height + 300, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub InitEnv()
    Dim curDate As Date
    Dim intDay As Integer
    Dim Rs As New ADODB.Recordset
    On Error GoTo errHand
    
    glngHours = Val(zlDatabase.GetPara("数据补录时限", glngSys))
    mintChange = Val(zlDatabase.GetPara("最近转出天数", glngSys, p住院护士站, 7))
    '出院病人时间范围
    curDate = zlDatabase.Currentdate
    intDay = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, p住院护士站, 7))
    mdtOutEnd = Format(curDate + intDay, "yyyy-MM-dd 23:59:59")
    intDay = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, p住院护士站, 30))
    mdtOutbegin = Format(mdtOutEnd - intDay, "yyyy-MM-dd 00:00:00")
    
    '打开现存在的所有护理记录项目
    gstrSQL = " Select  /*+ RULE */ 项目序号,项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式" & _
              " From 护理记录项目 B" & _
              " Where B.应用方式<>0 " & _
              " Order by 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    
    '提取除体温单外的护理文件清单
    gstrSQL = " Select /*+ RULE */ ID,名称 FROM 病历文件列表 WHERE 种类=3 AND 保留<>-1 AND 通用 > 0 ORDER BY 编号 "
    Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取除体温单外的护理文件清单")
    With Rs
        cbo护理文件格式.Clear
        Do While Not .EOF
            cbo护理文件格式.AddItem !名称
            cbo护理文件格式.ItemData(cbo护理文件格式.NewIndex) = !ID
            .MoveNext
        Loop
        If .RecordCount <> 0 Then cbo护理文件格式.ListIndex = 0
    End With
    
    '提取当前病区下的所有科室
    gstrSQL = " Select distinct B.ID,B.编码||'-'||B.名称 AS 科室" & _
              " From 病区科室对应 A,部门表 B,部门人员 C,人员表 D" & _
              " Where A.科室ID = b.ID And A.科室ID=C.部门ID And C.人员ID=D.ID And A.病区ID = [1]" & _
              IIf(InStr(1, mstrPrivs, "当前病区") <> 0, "", " And D.ID=[2]") & _
              " Order by B.编码||'-'||B.名称"
    Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, "提取当前病区下的所有科室", mlng病区ID, glngUserId)
    With cbo科室
        .Clear
        If InStr(1, mstrPrivs, "当前病区") <> 0 Then
            .AddItem "所有科室"
            .ItemData(.NewIndex) = -1
        End If
        Do While Not Rs.EOF
            .AddItem Rs!科室
            .ItemData(.NewIndex) = Rs!ID
            Rs.MoveNext
        Loop
        If Rs.RecordCount <> 0 Then .ListIndex = 0
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCOL As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, strColumns As String
    Dim blnSet As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    strColumns = mstrColumns
    If Not mblnInit Then
        '初始化内存记录集(未对应项目的列为活动项目,其它列均为固定项)
        strFields = "列," & adDouble & ",18|序号," & adDouble & ",2|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",20|固定," & adDouble & ",2|格式," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, strFields)
        strFields = "列|序号|项目序号|项目名称|固定|格式"
    End If

    '加入列定义
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCOL = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '如果已设置以传入值为准'否则找不到项目就是活动项目
            If UBound(Split(arrColumn(i), "'")) > 1 Then
                blnSet = True
                intImmovable = Split(arrColumn(i), "'")(2)
            End If
            If UBound(Split(arrColumn(i), "'")) > 2 Then
                strFormat = Split(arrColumn(i), "'")(3)
            End If

            k = UBound(arrItem)
            For l = 0 To k
                strName = arrItem(l)
                mrsItems.Filter = "项目名称='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!项目序号
                    If Not blnSet Then intImmovable = 1   '固定不允许修改
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0

                    '记录特殊列
                    Select Case strName
                    Case "日期"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "时间"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "护士"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "签名人"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "签名时间"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                strValues = lngCOL & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next

        '加入程序内部控制列(列是在读取数据后绑定时增加的,此时只有预处理下)
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '加上隐藏列
        mlngSigner = mlngSignLevel + 1
        mlngRecord = mlngSigner + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        If mlngOperator <> -1 And mlngSignName <> -1 Then
            mlngNoEditor = IIf(mlngOperator < mlngSignName, mlngOperator, mlngSignName)
        Else
            mlngNoEditor = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        End If
    End If

    mrsItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function SignMe() As Boolean
    Dim blnSign As Boolean          '是否签名成功
    Dim blnRefresh As Boolean
    Dim strTime As String
    Dim strSignTime As String       '保证所有签名的签名时间一致,便于取消签名时按签名时间统一取消
    Dim str状态 As String           '保存签名选项,避免循环签名时不停的弹出签名窗口
    Dim str行错误 As String
    Dim str错误 As String
    Dim intRow As Integer, intRows As Integer
    On Error GoTo errHand
    
    '普签:对当前界面的所有数据进行签名
    '准备签名
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    intRows = VsfData.Rows - 1
    For intRow = VsfData.FixedRows To intRows
        If Val(VsfData.TextMatrix(intRow, mlngRecord)) > 0 And VsfData.TextMatrix(intRow, mlngSigner) = "" Then
            str行错误 = ""
            If InStr(1, VsfData.TextMatrix(intRow, mlngDate), "/") <> 0 Then
                strTime = Format(Now, "yyyy") & "-" & ToStandDate(VsfData.TextMatrix(intRow, mlngDate)) & " " & VsfData.TextMatrix(intRow, mlngTime) & ":00"
            Else
                strTime = VsfData.TextMatrix(intRow, mlngDate) & " " & VsfData.TextMatrix(intRow, mlngTime) & ":00"
            End If
            blnSign = SignName(intRow, strTime, strSignTime, str状态, str行错误)
            If Not blnSign Then Exit For
            If Not blnRefresh Then blnRefresh = blnSign
            If str行错误 <> "" Then
                str错误 = str错误 & vbCrLf & "发生时间=[" & strTime & "]" & str行错误
            End If
        End If
    Next
    
    SignMe = blnRefresh
    mblnSigned = blnRefresh
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe()
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strTime As String
    Dim blnTrans As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim clsSign As Object
    On Error GoTo errHand
    '首先最后一次是本人的签名，根据当前选择数据的签名时间，批量取消签名

    gcnOracle.BeginTrans
    blnTrans = True
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) > 0 And VsfData.TextMatrix(lngRow, mlngSigner) <> "" Then
            If Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) > 0 Then
                '数字签名验证，只验证一次
                Err.Clear
                On Error Resume Next
                If clsSign Is Nothing Then
                    Set clsSign = CreateObject("zl9ESign.clsESign")
                    If Err <> 0 Then Err = 0
    
                    If Not clsSign Is Nothing Then
                        If clsSign.Initialize(gcnOracle, glngSys) Then
                            If Not clsSign.CheckCertificate(gstrDBUser) Then
                                gcnOracle.RollbackTrans
                                Exit Sub
                            End If
                        Else
                            gcnOracle.RollbackTrans
                            RaiseEvent AfterRowColChange("取消签名时需要再次认证，但系统没有设置签名认证中心，不能取消。", True)
                            Exit Sub
                        End If
                    Else
                        gcnOracle.RollbackTrans
                        RaiseEvent AfterRowColChange("签名部件初始化失败！", True)
                        Exit Sub
                    End If
                End If
            End If
        End If
        
        '提取发生时间
        If InStr(1, VsfData.TextMatrix(lngRow, mlngDate), "/") <> 0 Then
            strTime = Format(Now, "yyyy") & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime) & ":00"
        Else
            strTime = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime) & ":00"
        End If

        '取消签名
        gstrSQL = "ZL_病人护理数据_UNSIGNNAME("
        gstrSQL = gstrSQL & VsfData.TextMatrix(lngRow, c文件ID) & ","
        gstrSQL = gstrSQL & "To_Date('" & strTime & "','yyyy-MM-dd hh24:mi:ss'),0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "执行取消签名")
        '更改图标
        VsfData.Cell(flexcpPicture, lngRow, 0) = Nothing
        VsfData.TextMatrix(lngRow, mlngSignLevel) = 0
        VsfData.TextMatrix(lngRow, mlngSigner) = ""
    Next
    gcnOracle.CommitTrans
    blnTrans = False
    mblnSigned = False
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal intRow As Integer, ByVal strStart As String, ByVal strSignTime As String, _
    str状态 As String, Optional str错误 As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cEPRSign
    Dim strSource As String             '审签源数据串
    Dim lngLoop As Long
    Dim Rs As New ADODB.Recordset

    On Error GoTo errHand

    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""

    '获取要签名的内容
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = " Select /*+ RULE */ a.id,a.记录id,a.记录类型,a.项目分组,a.项目id,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位, " & _
              "     a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.开始版本,a.终止版本,a.记录人,a.记录时间  " & _
              " From 病人护理明细 a,病人护理数据 b,病人护理文件 c " & _
              " Where a.记录id=b.ID And B.汇总类别=0 And b.文件ID=c.ID And a.终止版本 Is Null And C.ID=[1] And b.发生时间=[2]"
    Call SQLDIY(gstrSQL)
    Set Rs = zlDatabase.OpenSQLRecord(gstrSQL, "获取要签名的内容", Val(VsfData.TextMatrix(intRow, c文件ID)), CDate(strStart))
    If Rs.BOF = False Then
        Do While Not Rs.EOF
            For lngLoop = 0 To Rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(Rs.Fields(lngLoop).Value, ""))
            Next
            Rs.MoveNext
        Loop
    End If
    If strSource = "" Then
        RaiseEvent AfterRowColChange("当前没有需要签名的信息！", True)
        Exit Function
    End If

    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    Set oSign = frmTendFileSign.ShowMe(Me, mstrPrivs, Val(VsfData.TextMatrix(intRow, c文件ID)), 未定义, strSource, False, str状态, str错误)
    On Error GoTo errHand

    If Not oSign Is Nothing Then
        gstrSQL = "ZL_病人护理数据_SIGNNAME("
        gstrSQL = gstrSQL & Val(VsfData.TextMatrix(intRow, c文件ID)) & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),0,"
        gstrSQL = gstrSQL & "'" & oSign.姓名 & "',"
        gstrSQL = gstrSQL & "'" & oSign.签名信息 & "'," & oSign.签名级别 & ","
        gstrSQL = gstrSQL & oSign.证书ID & ","
        gstrSQL = gstrSQL & oSign.签名方式 & ",'" & strSignTime & "')"
        
        Debug.Print gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, "执行签名")
        SignName = True
        
        VsfData.TextMatrix(intRow, mlngSignLevel) = oSign.证书ID
        VsfData.TextMatrix(intRow, mlngSigner) = "SignName"
        '更新图标
        VsfData.Cell(flexcpPicture, intRow, 0) = imgRow.ListImages(签名).Picture
    End If

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    mblnChange = False
    
    Call InitCons
    
    '内存记录集清空
    mrsCellMap.Filter = 0
    If mrsCellMap.RecordCount <> 0 Then mrsCellMap.MoveFirst
    Do While True
        If mrsCellMap.EOF Then Exit Do
        mrsCellMap.Delete
        mrsCellMap.Update
        mrsCellMap.MoveNext
    Loop
    
    Call DataMap_Restore
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function

    mblnShow = False
    Call InitCons
    SaveME = True
    
    RaiseEvent AfterRowColChange("保存成功！", False)
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngDeptID As Long, Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngDeptID           要显示护理记录的科室
    '返回： 无
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    Err = 0

    mblnInit = False
    mint页码 = 1
    mlng病区ID = lngDeptID
    mstrPrivs = strPrivs
    mblnBlowup = (zlDatabase.GetPara("护理文件显示模式", glngSys, 1255, 0) = 1)
    Set mfrmParent = frmParent

    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd")

    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '初始化环境
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    
    Call InitVariable
    Call InitCons
    
    If cbo科室.ListCount = 0 Then
        MsgBox "您不属于当前病区的任何科室，不能使用该功能！", vbInformation, gstrSysName
        Exit Function
    End If
    
    ShowMe = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckFlip() As Boolean
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCOL As Long, lngRows As Long, lngCols As Long
    
    Dim lng文件ID As Long, str时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '页面切换前检查：日期时间正确才允许继续，这样在保存时就不必再检查其它页面的数据了（其它数据在录入时已经进行了检查，此处略过）

    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow & " And 列号>" & mlngTime
        If mrsCellMap.RecordCount <> 0 Then
            If Not VsfData.RowHidden(lngRow) Then
                blnExit = (VsfData.TextMatrix(lngRow, mlngDate) = "" Or VsfData.TextMatrix(lngRow, mlngTime) = "")
                If blnExit Then
                    mrsCellMap.Filter = 0
                    VsfData.ROW = lngRow
                    If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                    CheckFlip = False
                    RaiseEvent AfterRowColChange("请补充日期时间！", True)
                    Exit Function
                End If
                
                '如果指定文件的录入时间存在数据则不允许录入
                If Val(VsfData.TextMatrix(lngRow, mlngRecord)) = 0 Then
                    lng文件ID = Val(VsfData.TextMatrix(lngRow, c文件ID))
                    If InStr(1, VsfData.TextMatrix(lngRow, mlngDate), "/") <> 0 Then
                        str时间 = Format(Now, "yyyy") & "-" & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate)) & " " & VsfData.TextMatrix(lngRow, mlngTime) & ":00"
                    Else
                        str时间 = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime) & ":00"
                    End If
                    gstrSQL = " Select /*+ RULE */ 1 From 病人护理数据 " & vbNewLine & _
                              " Where 文件ID=[1] And 发生时间=[2]"
                    Call SQLDIY(gstrSQL)
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "如果指定文件的录入时间存在数据则不允许录入", lng文件ID, CDate(str时间))
                    If rsTemp.RecordCount <> 0 Then
                        VsfData.ROW = lngRow
                        If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                        CheckFlip = False
                        mrsCellMap.Filter = 0
                        RaiseEvent AfterRowColChange("录入的发生时间已存在数据，请修改！", True)
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    mrsCellMap.Filter = 0
    CheckFlip = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mrsCellMap.Filter = 0
End Function

Private Function CheckData() As Boolean
    Dim intLevel As Integer
    Dim lngPage As Long
    On Error GoTo errHand
    '检查数据

    '如果修改了数据而日期时间不全则提示（数据合法性在录入时已经检查）
    If Not CheckFlip Then Exit Function
'    Call OutputRsData(mrsCellMap)
'    Call OutputRsData(mrsDataMap)

    CheckData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim arrValue, arrOrder, arrPart, arrCollect
    Dim strSQL() As String
    Dim intAllow As Integer
    Dim lngRecord As Long
    Dim blnTrans As Boolean, blnSaved As Boolean, blnDel As Boolean
    Dim intPos As Integer, intMax As Integer, intPage As Integer, intRow As Integer, intUsedRows As Integer
    Dim strReturn As String, strCellData As String, strPart As String
    Dim strMonth As String, strDay As String
    Dim strDate As String, strTime As String, strTemp As String
    Dim strDatetime As String, strCurrDate As String, strDays As String
    Dim strSaveRows As String

    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '同行多列循环调用：ZL_病人护理数据_UPDATE
    '下一行前调用：
    '   1、ZL_病人护理数据_SYNCHRO，同步数据到体温单与护理记录单中，需要记录删除的明细ID串
    '   2、ZL_病人护理打印_UPDATE，完成打印数据解析
    '删除项目需记录，删除行也需要记录
    '修改数据的同步就将该行数据对应的日期与时间保存到mrsCellMap中

'    objStream.WriteLine (Now & "产生保存SQL")
    intAllow = IIf(InStr(mstrPrivs, "他人护理记录") > 0, 1, 0)
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")

    With mrsCellMap
        '将有效数据过滤出来:记录ID>0的历史数据+新增的有效数据
        .Filter = "记录ID>0 or (记录ID=0 And 删除=0)"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If intRow <> !行号 Then
endWork:
                If intRow > 0 Then
                    blnDel = VsfData.RowHidden(intRow)
                    intUsedRows = Val(Split(VsfData.TextMatrix(intRow, mlngRowCount), "|")(0))
                End If

                If blnSaved Then
                    strSaveRows = strSaveRows & "," & intRow
                    
                    '完成打印数据解析
'                    文件ID_IN IN 病人护理打印.文件ID%TYPE,
'                    发生时间_IN IN 病人护理打印.发生时间%TYPE,
'                    行数_IN IN 病人护理打印.行数%TYPE,
'                    删除_IN Number:=0
                    gstrSQL = "ZL_病人护理打印_UPDATE(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss')," & intUsedRows & "," & IIf(blnDel, "1", "0") & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL

                    '只要修改过数据,必然会执行打印解析,因此在这里进行汇总日期的处理
                    If InStr(1, "," & strDays & ",", "," & Mid(strDatetime, 1, 10) & ",") = 0 Then
                        '同步更新明天的汇总(夜班,全天汇总跨天的处理)
                        strDays = strDays & "," & Mid(strDatetime, 1, 10)
                        gstrSQL = "ZL_汇总数据_UPDATE(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",'" & Mid(strDatetime, 1, 10) & "')"
                        strSQL(ReDimArray(strSQL)) = gstrSQL

                        strTemp = Format(DateAdd("d", 1, CDate(strDatetime)), "yyyy-MM-dd")
                        If InStr(1, "," & strDays & ",", "," & strTemp & ",") = 0 Then
                            strDays = strDays & "," & strTemp
                            gstrSQL = "ZL_汇总数据_UPDATE(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",'" & strTemp & "')"
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                        End If
                    End If

                    blnSaved = False
                    If .EOF Then Exit Do
                End If

                '赋初值
                intPage = !页号
                intRow = !行号
                strDate = ""
                strDatetime = ""
                lngRecord = NVL(!记录ID, 0)
            End If

            If !列号 = mlngDate Then
                If NVL(!汇总, 0) = 1 Then
                    arrCollect = Split(!数据, ";")
                    strDatetime = arrCollect(3)
                '    文件ID_IN IN 病人护理数据.文件ID%TYPE,
                '    发生时间_IN IN 病人护理数据.发生时间%TYPE,
                '    汇总类别_IN IN 病人护理数据.汇总类别%TYPE,
                '    汇总文本_IN IN 病人护理数据.汇总文本%TYPE,
                '    汇总标记_IN IN 病人护理数据.汇总标记%TYPE,
                '    删除_IN Number:=0
                    gstrSQL = "ZL_病人护理数据_COLLECT(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",to_date('" & arrCollect(3) & "','yyyy-MM-dd hh24:mi:ss')," & _
                            Val(arrCollect(1)) & ",'" & arrCollect(0) & "'," & Val(arrCollect(2)) & "," & !删除 & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    blnSaved = True
                Else
                    strDate = NVL(!数据)
                    If strDate <> "" Then
                        If mblnDateAd Then
                            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                        Else
                            strDate = Format(strDate, "yyyy-MM-dd")
                        End If
                    End If
                End If
            ElseIf !列号 = mlngTime Then
                strTime = NVL(!数据)
                If strDatetime = "" Then
                    If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                    strDatetime = strDate & " " & strTime & ":00"
                End If

                If lngRecord <> 0 Then
                    '更新发生时间
                    gstrSQL = "Zl_病人护理数据_发生时间(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'))"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    blnSaved = True
                End If
            Else
                If !列号 > mlngTime Then
                    '取指定单元格的数据
                    strCellData = NVL(!数据)
                    strPart = NVL(!部位)
                    strReturn = ShowInput(!列号, strCellData, True)
                    'strOrders格式：项目序号,项目序号...
                    'strValues格式：值'值'值...
                    arrOrder = Split(Split(strReturn, "||")(0), ",")
                    arrValue = Split(Split(strReturn, "||")(1) & "'", "'")
                    arrPart = Split(strPart & "/////", "/")

                    intMax = UBound(arrOrder)
                    For intPos = 0 To intMax
                        If Not (Val(VsfData.TextMatrix(intRow, mlngRecord)) = 0 And arrValue(intPos) = "") Then
    '                    文件ID_IN IN 病人护理数据.文件ID%TYPE,
    '                    发生时间_IN IN 病人护理数据.发生时间%TYPE,
    '                    记录类型_IN IN 病人护理明细.记录类型%TYPE,          --护理项目=1，上标说明=2，手术日标记=4，签名记录=5，下标说明=6，入出量汇总=9
    '                    项目序号_IN IN 病人护理明细.项目序号%TYPE,          --护理项目的序号，非护理项目固定为0
    '                    记录内容_IN IN 病人护理明细.记录内容%TYPE := NULL,  --记录内容，如果内容为空，即清除以前的内容；37或38/37
    '                    体温部位_IN IN 病人护理明细.体温部位%TYPE := NULL,
    '                    他人记录_IN IN NUMBER := 1,
                        gstrSQL = "ZL_病人护理数据_UPDATE(" & Val(VsfData.TextMatrix(intRow, c文件ID)) & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
                                arrOrder(intPos) & ",'" & arrValue(intPos) & "','" & arrPart(intPos) & "'," & intAllow & ",0,0)"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        blnSaved = True
                        End If
                    Next
                    mrsItems.Filter = 0
                End If
            End If

            .MoveNext
        Loop

        If blnSaved Then GoTo endWork
    End With

    '循环执行SQL保存数据
    On Error Resume Next
    intMax = UBound(strSQL)

    gcnOracle.BeginTrans
    blnTrans = True

    On Error GoTo errHand
    If intMax > 0 Then
        Debug.Print "开始保存数据:" & Now
'        objStream.WriteLine (Now & "准备保存数据")
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                Debug.Print Now & ":" & strSQL(intPos)
    '            objStream.WriteLine (Now & "；SQL：" & strSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "保存护理记录单数据")
            End If
        Next
        Debug.Print Now & ":保存数据完成"
    '    objStream.WriteLine (Now & "保存数据完成")
    End If

    gcnOracle.CommitTrans
    SaveData = True
    blnTrans = False
    mblnSaved = True
    mblnChange = False
    
    '更新数据行的记录ID列,表示该数据已保存
    strSaveRows = strSaveRows & ","
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        If InStr(1, strSaveRows, "," & intRow & ",") <> 0 Then
            VsfData.TextMatrix(intRow, mlngRecord) = 1
        End If
    Next
    
    '内存记录集清空
    mrsCellMap.Filter = 0
    mrsCellMap.MoveFirst
    Do While True
        If mrsCellMap.EOF Then Exit Do
        mrsCellMap.Delete
        mrsCellMap.Update
        mrsCellMap.MoveNext
    Loop
    
    '保存当前数据
    Call DataMap_Save
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDate As String, strTime As String
    Dim strLockItem As String                   '同步过来的数据,不允许修改或删除
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       '同步过来的数据占用的最大行数
    Dim intNULL As Integer, lngStartRow As Long
    Dim lngRow As Long, lngCOL As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String

    Select Case Control.ID
    '粘贴,清除时需要同步mrsCellMap数据
    Case conMenu_Edit_Copy
        '复制指定数据行的数据
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)

        '复制记录集
        Set mrsCopyMap = New ADODB.Recordset
        Set mrsCopyMap = CopyNewRec(mrsDataMap, False)

        '得到指定数据行的起始行,结束行
        lngCols = VsfData.Cols - 1
        lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngRows = lngRow + lngRows - 1
        For lngRow = lngRow To lngRows
            mrsCopyMap.AddNew
            mrsCopyMap!页号 = mint页码
            mrsCopyMap!行号 = lngRow
            For lngCOL = 0 To lngCols - VsfData.FixedCols    '多了一个固定列
                mrsCopyMap.Fields(cControlFields + lngCOL).Value = IIf(VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCOL + VsfData.FixedCols))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        '粘贴时，将目标行整体覆盖，同步过来的数据列，活动列除外
        '活动项目可能不同页面项目不同，部位不同，所以不考虑活动项目
        '同步行所占用的行数不变，如不够再添加空白行，再行粘贴
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If mrsCopyMap.RecordCount = 0 Then Exit Sub

        '得到目标数据行的起始行,结束行
        strField = "ID|页号|行号|列号|记录ID|数据|删除"
        lngCols = VsfData.Cols - 1
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            lngStartRow = lngRow
            If mlngDate > -1 Then strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
        Else
            '删除多余的数据行,仅留一行
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0)) - 1
            For intNULL = 1 To lngRows
                VsfData.RemoveItem lngRow + 1
            Next
        End If

        '往下搜索空行,如果有其它数据行则计算需增加的行数
        intNULL = mrsCopyMap.RecordCount - 1
        For lngRow = 1 To mrsCopyMap.RecordCount - 1
            '保证当前输入的内容在一页中显示全
            If lngRow + VsfData.ROW > VsfData.Rows - 1 Then Exit For

            If Val(VsfData.TextMatrix(lngRow + VsfData.ROW, c病人ID)) = 0 And VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRowCount) = "" Then
                intNULL = intNULL - 1
            Else
                Exit For
            End If
        Next
        '先增加空行
        If intNULL > 0 Then
            VsfData.Rows = VsfData.Rows + intNULL
            '从当前行记录的空白行开始，每行的位置+所增加的空白行数
            For lngRow = 1 To intNULL
                VsfData.RowPosition(VsfData.Rows - 1) = lngStartRow + 1
            Next
        End If

        '还原日期，时间，强制不允许修改
        VsfData.TextMatrix(lngStartRow, mlngDate) = strDate
        VsfData.TextMatrix(lngStartRow, mlngTime) = strTime
        '记录用户修改过的单元格
        If mlngDate <> -1 Then
            strKey = mint页码 & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\时间
        strKey = mint页码 & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '向表格填充数据
        With mrsCopyMap
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For lngCOL = 0 To lngCols - VsfData.FixedCols
                    Select Case lngCOL + VsfData.FixedCols
                    Case 1, c文件ID, c床号, c姓名, c病人ID, c主页ID, c婴儿, _
                         mlngDate, mlngTime, mlngOperator, mlngSigner, mlngSignTime, mlngRecord
                    Case Else
                        If InStr(1, "," & strLockItem & ",", "," & lngCOL - (cHideCols - 1) & ",") = 0 And InStr(1, "," & mstrCOLNothing & ",", "," & lngCOL - (cHideCols - 1) & ",") = 0 Then
                            VsfData.TextMatrix(lngStartRow + .AbsolutePosition - 1, lngCOL + VsfData.FixedCols) = NVL(.Fields(cControlFields + lngCOL).Value)

                            '修改标志
                            If .AbsolutePosition = 1 Then
                                strKey = mint页码 & "," & lngStartRow & "," & lngCOL + VsfData.FixedCols
                                strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCOL + VsfData.FixedCols & "|" & _
                                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & GetMutilData(lngStartRow, lngCOL + VsfData.FixedCols, lngTop, lngHeight) & "|0"
                                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                            End If
                        End If
                    End Select
                Next
                .MoveNext
            Loop
        End With
        '表格上色
        Call SetActiveColColor
        mblnChange = True

    Case conMenu_Edit_Clear
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub

        '准备删除
        strField = "ID|页号|行号|列号|记录ID|数据|汇总|删除"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
        Else
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            '删除所有数据行
            lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            For intNULL = 2 To lngRows
                VsfData.RowHidden(lngRow + intNULL - 1) = True
            Next
        End If

        '记录用户修改过的单元格
        strKey = mint页码 & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '2\时间
        strKey = mint页码 & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        
        '删除启始行中非同步的数据
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            '填写修改标志
            For lngCOL = mlngTime + 1 To mlngNoEditor - 1
                strKey = mint页码 & "," & lngStartRow & "," & lngCOL
                strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCOL & "|" & _
                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            Next
        Else
            '填写修改标志(存在同步数据,日期与时间列不允许清除)``
            For lngCOL = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCOL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCOL <> mlngDate And lngCOL <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCOL) = ""

                    strKey = mint页码 & "," & lngStartRow & "," & lngCOL
                    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCOL & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
            Next
            VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        End If
        mblnChange = True

    Case conMenu_Edit_SPECIALCHAR

        '检查当前录入控件
        On Error Resume Next
        Dim objTXT As TextBox
        Dim strText As String
        Dim intPos As Integer, intLen As Integer

        mstrSymbol = frmInsSymbol.ShowMe(False, 0)
        If mintSymbol = -1 Then
            Set objTXT = txtInput
        Else
            Set objTXT = txt(mintSymbol)
        End If
        strText = objTXT.Text
        intPos = objTXT.SelStart
        intLen = Len(objTXT)
        objTXT.Text = Mid(strText, 1, intPos) & mstrSymbol & Mid(strText, intPos + 1)
    Case conMenu_Edit_Word
        Call cmdWord_Click
    Case conMenu_Edit_NewItem
        '在当前有效数据行（可能当前有效数据行是多行）之后增加一空白行
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        VsfData.Rows = VsfData.Rows + 1
        VsfData.TextMatrix(VsfData.Rows - 1, c文件ID) = VsfData.TextMatrix(lngStartRow, c文件ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c床号) = VsfData.TextMatrix(lngStartRow, c床号)
        VsfData.TextMatrix(VsfData.Rows - 1, c姓名) = VsfData.TextMatrix(lngStartRow, c姓名)
        VsfData.TextMatrix(VsfData.Rows - 1, c病人ID) = VsfData.TextMatrix(lngStartRow, c病人ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c主页ID) = VsfData.TextMatrix(lngStartRow, c主页ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c婴儿) = VsfData.TextMatrix(lngStartRow, c婴儿)
        
        strKey = VsfData.TextMatrix(lngStartRow, mlngRowCount)
        If InStr(1, strKey, "|") <> 0 And strKey <> "1|1" Then
            strKey = Split(strKey, "|")(0)
            strKey = strKey & "|" & strKey
            For lngRow = VsfData.ROW + 1 To VsfData.Rows - 1
                If VsfData.TextMatrix(lngRow, mlngRowCount) = strKey Then
                    lngStartRow = lngRow + 1
                    Exit For
                End If
            Next
        Else
            lngStartRow = VsfData.ROW + 1
        End If
        
        For lngRow = VsfData.Rows - 2 To lngStartRow Step -1    '从倒数第二行开始
            VsfData.RowPosition(lngRow) = lngRow + 1
        Next
        Call SetActiveColColor
        
        mblnChange = True
    Case conMenu_Edit_Save
        Call SaveME
    Case conMenu_Edit_Transf_Cancle
        Call CancelMe
    Case conMenu_Tool_Sign
        Call SignMe
    Case conMenu_Tool_SignEarse
        Call UnSignMe
    End Select
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrData
    Dim blnFind As Boolean
    Dim strItem As String
    Dim intDo  As Integer, intCount As Integer

    Select Case Control.ID
    Case conMenu_Edit_Copy
        Control.Enabled = Not mblnShow
    Case conMenu_Edit_PASTE
        Control.Enabled = False
        If Not mblnInit Then Exit Sub
        If mblnSigned Then Exit Sub
        If mrsCopyMap.State = 0 Then Exit Sub
        Control.Enabled = Not mblnShow And mrsCopyMap.RecordCount
    Case conMenu_Edit_Clear
        Control.Enabled = Not mblnSigned
    Case conMenu_Edit_SPECIALCHAR
        Control.Enabled = mblnShow And (mintType = 0 Or mintType = 6)
    Case conMenu_Edit_Word
        Control.Enabled = mblnEditAssistant And Not mblnSigned
    Case conMenu_Edit_NewItem
        Control.Enabled = Not mblnSigned
    Case conMenu_Edit_Save
        Control.Enabled = mblnChange And Not mblnSigned
    Case conMenu_Edit_Transf_Cancle
        Control.Enabled = mblnChange
    Case conMenu_Tool_Sign
        Control.Enabled = mblnSaved And Not mblnSigned And Not mblnChange
    Case conMenu_Tool_SignEarse
        Control.Enabled = mblnSaved And mblnSigned And Not mblnChange
    End Select
End Sub

Private Function ISActiveUsed(ByVal strTest As String) As Boolean
    Dim arrData, arrCol
    Dim lngCOL As Long
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '检查某个活动项目是否已被其它列绑定
    ISActiveUsed = True

    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        lngCOL = Split(Split(arrData(intDo), "|")(0), ";")(0)
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            If strTest = arrCol(intIn) And VsfData.COL - (cHideCols + VsfData.FixedCols - 1) <> lngCOL Then
                RaiseEvent AfterRowColChange(Split(strTest, ",")(1) & mrsItems!项目名称 & " 已经被绑定到" & lngCOL & "列，不允许重复绑定！", True)
                Exit Function
            End If
        Next
    Next
    ISActiveUsed = False
End Function

Private Function GetActivePart(ByVal intFindCol As Integer, ByVal intItem As Integer) As String
    '获取指定列的活动项目
    Dim arrData
    Dim arrCol
    Dim intCol As Integer, strPart As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '将活动项目加入到查询SQL中，格式：列号;表头名称|项目序号,部位;项目序号,部位||列号;表头名称...
    '绑定多个项目，该列就自动转为对角线列

    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCol = intFindCol - cHideCols Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            strPart = Split(arrCol(intItem), ",")(1)
            Exit For
        End If
    Next
    GetActivePart = strPart
End Function

Private Sub cmdWord_Click()
    Dim strInput As String
    '弹出词句选择器

    If cmdWord.Tag = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, Val(VsfData.TextMatrix(VsfData.ROW, c病人ID)), Val(VsfData.TextMatrix(VsfData.ROW, c主页ID)), Val(VsfData.TextMatrix(VsfData.ROW, c婴儿)), strInput)

    If cmdWord.Tag = -1 Then
        txtInput.Text = strInput
    Else
        txt(Val(cmdWord.Tag)).Text = strInput
    End If
End Sub

Private Sub cmd刷新_Click()
    '读取文件格式
    mblnInit = False
    mlng格式ID = cbo护理文件格式.ItemData(cbo护理文件格式.ListIndex)
    mlng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    
    Call InitVariable
    Call InitCons
    Call ReadStruDef
    Call zlRefresh
    mblnInit = True
    
    '保存当前数据
    Call DataMap_Save
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim lngRow As Long, lngCOL As Long
    Dim dblHeight As Double, dblWidth As Double

    If Not mblnInit Then Exit Sub
    Call InitCons

'    '计算固定行的高度
'    For lngRow = 0 To 2
'        If Not VsfData.RowHidden(lngRow) Then dblHeight = dblHeight + VsfData.ROWHEIGHT(lngRow)
'    Next
'    '从可见行开始向下查找最后一个可见行
'    For lngRow = NewTopRow To VsfData.Rows - 1
'        If Not VsfData.RowIsVisible(lngRow) Then
'            lngRow = lngRow - 1
'            Exit For
'        End If
'    Next
'    '从可见列开始查找最后一个可见列
'    For lngCol = NewLeftCol To VsfData.Cols - 1
'        If Not VsfData.ColIsVisible(lngCol) Then
'            lngCol = lngCol - 1
'            Exit For
'        Else
'            dblWidth = dblWidth + VsfData.ColWidth(lngCol)
'        End If
'    Next
'
'    If Not VsfData.RowIsVisible(VsfData.Row) Then
'        VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'    Else
'        '当前数据行的高度+固定行的高度如果大于表格控件的高度,说明当前选择的数据行存在遮住部分的情况
'        If VsfData.Row >= lngRow - 1 And CellRect.Bottom * (lngRow - NewTopRow + 1) + dblHeight >= VsfData.ClientHeight Then
'            '遮住部分的情况下
'            VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'        End If
'    End If
'
'    If Not VsfData.ColIsVisible(VsfData.Col) Then
'        VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'    Else
'        '当前数据行的高度+固定行的高度如果大于表格控件的高度,说明当前选择的数据行存在遮住部分的情况
'        If VsfData.Col = lngCol And dblWidth >= VsfData.ClientWidth Then
'            '遮住部分的情况下
'            VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'        End If
'    End If
'
'    Call VsfData_EnterCell
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_KeyDown(Asc("A"), 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim strCols As String
    Dim intMax As Integer
    Dim lngStart As Long
    On Error Resume Next

    '隐蔽已显示的录入控件
    cmdWord.Visible = False
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
    Case 4, 5
        picDouble.Visible = False
    Case 6
        picMutilInput.Visible = False
    End Select

    '未定义的列不允许录入数据
    mintType = -1
    If InStr(1, mstrPrivs, "护理记录登记") = 0 Then Exit Sub
    If mblnSigned Then Exit Sub
    If Not mblnShow Then Exit Sub
    
    '如果是活动项目则不允许编辑
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - cHideCols & ",") <> 0 Then Exit Sub
    If VsfData.COL <= c姓名 Then Exit Sub
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
    '让控件获得焦点
    Select Case mintType
    Case 0, 3
        picInput.SetFocus
    Case 1, 2
        lstSelect(mintType - 1).SetFocus
    Case 4, 5
        picDouble.SetFocus
    Case 6
        picMutilInput.SetFocus
    End Select
End Sub

Private Sub vsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    If mblnInit = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub

    '选择列,同步数据列直接退出,避免此处清除提示信息
    '显示当前项目的相关信息
    mrsSelItems.Filter = "列=" & NewCol - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!项目值域) <> "" Then
                If mrsItems!项目类型 = 0 Then
                    strInfo = "有效范围:" & Split(mrsItems!项目值域, ";")(0) & "～" & Split(mrsItems!项目值域, ";")(1)
                Else
                    strInfo = "有效范围:" & mrsItems!项目值域
                End If
            Else
                strInfo = ""
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0

    RaiseEvent AfterRowColChange(strInfo, False)
End Sub

Private Sub vsfData_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCell(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub VsfData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngStart As Long
    Dim intLevel As Integer
    Dim strField As String, strKey As String, strValue As String

    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
End Sub

Private Sub InitVariable()
    '清除常量
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignName = -1
    mlngSignTime = -1
    mlngRecord = -1
    mlngNoEditor = -1

    mblnShow = False
    mblnSigned = False
    mblnSaved = False
    mblnChange = False
    mblnEditAssistant = False
End Sub

Private Sub InitCons()
    '隐藏输入控件
    picInput.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    picDouble.Visible = False
    picMutilInput.Visible = False
    cmdWord.Visible = False
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim Rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar

    On Error GoTo errHand

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False

    cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)

        '------------------------------------------------------------------------------------------------------------------
        '工具栏定义
        Set cbrToolBar = cbsThis.Add("标准工具", xtpBarTop)
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        cbrToolBar.ShowTextBelowIcons = False
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制"): cbrControl.ToolTipText = "复制(Ctrl+C)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "粘贴"):  cbrControl.ToolTipText = "粘贴(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除"):   cbrControl.ToolTipText = "清除"

            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "特殊符号"):  cbrControl.ToolTipText = "插入特殊符号(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "词句选择"):  cbrControl.ToolTipText = "词句选择(Ctrl+W)"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "空行"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "增加空行"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Sign, "签名"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消"): cbrControl.IconId = 229
        End With

        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next

        '------------------------------------------------------------------------------------------------------------------
        '工具栏定义
        Set cbrToolBar = cbsThis.Add("过滤条件", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrCustom = .Add(xtpControlCustom, 0, "")
            cbrCustom.Flags = xtpFlagAlignLeft
            cbrCustom.Handle = pic过滤条件.hwnd
            cbrCustom.ToolTipText = "条件"
        End With

         '快键绑定
        With cbsThis.KeyBindings
            .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
            .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
            .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
            .Add FCONTROL, Asc("W"), conMenu_Edit_Word
        End With

    InitMenuBar = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng病人id As Long, ByVal lng主页id As Long, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '数据发生时间必须在当前科室的有效时间范围内

    blnMsg = (strMsg <> "")

    '检查文件开始,结束时间
    If strTime <= Format(mstr开始时间, "yyyy-MM-dd HH:mm") Then
        strMsg = "发生时间不能小于文件开始时间[" & mstr开始时间 & "]"
        GoTo exitHand
    End If
    If mstr结束时间 <> "" Then
        If strTime <= Format(mstr结束时间, "yyyy-MM-dd HH:mm") Then
            strMsg = "发生时间不能大于文件结束时间[" & mstr结束时间 & "]"
            GoTo exitHand
        End If
    End If

    '根据病人变动记录进行检查
    gstrSQL = " Select  /*+ RULE */ 开始原因,病区ID,to_char(开始时间,'yyyy-MM-dd hh24:mi') AS 开始时间,to_char(NVL(终止时间,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS 终止时间 " & _
              " From 病人变动记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]" & _
              " Order by 开始时间,开始原因"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前科室有效时间范围", lng病人id, lng主页id)
    With rsTemp
        .Filter = "病区ID=" & mlng病区ID
        Do While Not .EOF
            If strTime >= !开始时间 And strTime <= !终止时间 Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '找到了就退出
        If blnExist Then
            If Not IsAllowInput(lng病人id, lng主页id, strTime, strCurTime) Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]"
                GoTo exitHand
            End If

            CheckTime = True
            Exit Function
        End If

        '没找到,就整理原因进行准确性提示
        .Filter = "开始原因=1"
        If .RecordCount <> 0 Then
            If !开始原因 = 1 And strTime < !开始时间 Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入院时间:" & !开始时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=2"
        If .RecordCount <> 0 Then
            If !开始原因 = 2 And strTime < !开始时间 Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能小于病人入科时间:" & !开始时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = "开始原因=10"
        If .RecordCount <> 0 Then
            If !开始原因 = 10 And strTime > !终止时间 Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[发生时间不能大于出院时间:" & !终止时间 & "]"
                GoTo exitHand
            End If
        End If
        .Filter = 0
        '其他情况说明
        strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[不在当前病区的有效时间范围内]"
        GoTo exitHand
    End With

    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Function CheckInput(strReturn As String, strInfo As String) As Boolean
    Dim i As Integer, j As Integer
    Dim strOrders As String, strText As String
    '检查录入数据的合法性(中文也认为是一个字符,考虑到体温项目等存在不升\外出等信息)
    '返回的数据,如果一列绑定多个项目,以单引号做为分隔符

    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定N个项目,手工录入
    Select Case mintType
    Case 0
        strText = txtInput.Text
        strOrders = txtInput.Tag
    Case 1, 2   '免检
        If mintType = 1 Then
            strText = Mid(lstSelect(mintType - 1).Text, 2)
        Else
            j = lstSelect(mintType - 1).ListCount
            For i = 1 To j
                If lstSelect(mintType - 1).Selected(i - 1) Then
                    strText = strText & "," & Mid(lstSelect(mintType - 1).List(i - 1), 2)
                End If
            Next
            If strText <> "" Then strText = Mid(strText, 2)
        End If
        strOrders = lstSelect(mintType - 1).Tag
    Case 4
        strText = txtUpInput.Text & "'" & txtDnInput.Text
        strOrders = txtUpInput.Tag & "'" & txtDnInput.Tag
    Case 6
        j = txt.Count
        For i = 1 To j
            strText = strText & "'" & txt(i - 1).Text
            strOrders = strOrders & "'" & txt(i - 1).Tag
        Next
        If strText <> "" Then
            strText = Mid(strText, 2)
            strOrders = Mid(strOrders, 2)
        End If
    Case 3      '免检
        strText = lblInput.Caption
    Case 5      '免检
        strText = lblUpInput.Caption & "/" & lblDnInput.Caption
    End Select
    If Val(strOrders) <> 0 Then
        If Not CheckValid(strText, strOrders, strInfo) Then Exit Function
    ElseIf VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then
        If Not CheckDateTime(strText, strInfo) Then Exit Function
    End If

    strReturn = strText
    CheckInput = True
End Function

Private Function CheckDateTime(strText As String, strInfo As String) As Boolean
    Dim arrData
    Dim blnCheck As Boolean
    Dim strCurrDate As String
    Dim strDate As String, strMonth As String, strDay As String

    If VsfData.COL = mlngDate Then
        If mblnDateAd Then
            If Trim(strText) = "" Then
                strInfo = "日期不能为空！"
                Exit Function
            End If
            If InStr(1, strText, "/") = 0 Then
                strInfo = "日期格式错误，如1月12日：12/01"
                Exit Function
            End If

            strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
            If Not IsDate(strDate) Then
                strInfo = "录入的数据不是合法的日期，如1月12日：12/01"
                Exit Function
            End If
        Else
            If Trim(strText) = "" Then
                strInfo = "日期不能为空！"
                Exit Function
            End If
            If Not IsDate(strText) Then
                strInfo = "录入的数据不是合法的日期，如1月12日：2011-01-12"
                Exit Function
            End If
            strDate = Format(strText, "yyyy-MM-dd")
        End If
        If strDate > mstrMaxDate Then
            strInfo = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
            Exit Function
        End If

        If VsfData.TextMatrix(VsfData.ROW, mlngTime) <> "" Then
            blnCheck = True
            strDate = strDate & " " & VsfData.TextMatrix(VsfData.ROW, mlngTime)
        End If
    Else
        If Trim(strText) = "" Then
            strInfo = "时间不能为空！"
            Exit Function
        End If
        If Len(strText) <= 2 Then
            strText = String(2 - Len(strText), "0") & strText
            strText = strText & ":00"
        End If
        If Val(Mid(strText, 1, 2)) < 0 Or Val(Mid(strText, 1, 2)) > 23 Then
            strInfo = "录入的时间无效，小时应该在0-23之间！"
            Exit Function
        End If
        If Mid(strText, 3, 1) <> ":" Then
            strInfo = "录入的时间格式错误[09:00]！"
            Exit Function
        End If
        If Len(strText) < 5 Then strText = strText & String(5 - Len(strText), "0")
        If Not (Val(Mid(strText, 4, 2)) >= 0 And Val(Mid(strText, 4, 2)) <= 59) Then
            strInfo = "录入的时间无效，分钟应该在0-59之间！"
            Exit Function
        End If
        If Len(strText) > 5 Then
            strInfo = "录入的时间格式错误[09:00]！"
            Exit Function
        End If

        '进行合法性检查
        If VsfData.TextMatrix(VsfData.ROW, mlngDate) <> "" Then
            strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            strDate = VsfData.TextMatrix(VsfData.ROW, mlngDate)
            If mblnDateAd Then
                strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
            Else
                strDate = Format(VsfData.TextMatrix(VsfData.ROW, mlngDate), "yyyy-MM-dd")
            End If
            strDate = strDate & " " & strText
            blnCheck = True
        End If
    End If

    If blnCheck Then
        '数据发生时间不能在当前操作员所属科室的有效时间以前
'        If Not CheckTime(VsfData.Row, mlng病人id, mlng主页id, strDate, strCurrDate, strInfo) Then
'            Exit Function
'        End If
    End If

    CheckDateTime = True
End Function

Private Function CheckValid(strReturn As String, ByVal strOrders As String, strInfo As String) As Boolean
    Dim arrData, arrOrder
    Dim i As Integer, j As Integer
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String, strFormat As String

    '按列格式组装数据
    mrsSelItems.Filter = "列=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        '有此列但未进行定义
        strFormat = NVL(mrsSelItems!格式)   '{P[体温]C}{...}
    End If
    mrsSelItems.Filter = 0

    '检查数据
    arrData = Split(strReturn, "'")
    arrOrder = Split(strOrders, "'")
    j = UBound(arrData)
    For i = 0 To j
        strText = arrData(i)
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!项目名称
            If strText <> "" Then
                If mrsItems!项目类型 = 0 And mrsItems!项目表示 = 0 Then
                    strText = Val(strText)
                    If NVL(mrsItems!项目小数, 0) <> 0 Then   '等于零是通过控件的MaxLength来控制的
                        If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                        If Len(strText) > mrsItems!项目长度 Then
                            mrsItems.Filter = 0
                            strInfo = "[" & strName & "]录入的数据超过了合法精度！"
                            Exit Function
                        End If

                        strText = Val(arrData(i))
                        If InStr(1, strText, ".") <> 0 Then
                            strText = Mid(strText, InStr(1, strText, ".") + 1)
                            If Len(strText) > mrsItems!项目小数 Then
                                mrsItems.Filter = 0
                                strInfo = "[" & strName & "]录入的小数部分超过了合法精度！"
                                Exit Function
                            End If
                        End If
                        strText = Val(arrData(i))
                    End If
                    If Not IsNull(mrsItems!项目值域) Then
                        dblMin = Split(mrsItems!项目值域, ";")(0)
                        dblMax = Split(mrsItems!项目值域, ";")(1)
                        If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                            mrsItems.Filter = 0
                            strInfo = "[" & strName & "]录入的数据不在" & Format(dblMin, "#0.00") & "～" & Format(dblMax, "#0.00") & "的有效范围！"
                            Exit Function
                        End If
                    End If
                Else
                    If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!项目长度 Then
                        strInfo = "[" & strName & "]录入的数据超过了最大长度：" & mrsItems!项目长度 & "！"
                        mrsItems.Filter = 0
                        Exit Function
                    End If
                End If
                strFormat = Replace(strFormat, "[" & strName & "]", strText)
            Else
                '删除该项目
                Call SubstrPro(strFormat, strName)
            End If
        Else
            strFormat = strReturn
        End If
    Next
    If j = -1 Then
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = mrsItems!项目名称
            strFormat = Replace(strFormat, "[" & strName & "]", strText)
        End If
    End If
    mrsItems.Filter = 0

    strFormat = Replace(strFormat, "{", "")
    strFormat = Replace(strFormat, "}", "")
    strReturn = strFormat
    CheckValid = True
End Function

Public Function SubstrVal(ByVal strData As String, ByVal strFormat As String, ByVal strName As String, intPos As Integer) As String
    Dim i As Integer, j As Integer, l As Integer, r As Integer
    Dim strQZ As String, strHZ As String
    '返回前一个项目的后缀符号+当前项目的前缀符号的位置

    If strData = "" Then Exit Function
    strData = UCase(strData)
    j = Len(strFormat)
    l = InStr(1, strFormat, "[" & strName & "]")
    If l = 0 Then Exit Function
    '得到前缀
    For i = l To 1 Step -1
        If Mid(strFormat, i, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, i + 1, l - i - 1)
    '找到该项目格式串中的结束符号
    i = l + Len(strName) + 2
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    '得到后缀
    strHZ = Mid(strFormat, i, r - i)
    '如果后缀为空,继续向后寻找下一个项目的前缀符号
    If strHZ = "" And r < j Then
        For r = r + 1 To j
            If Mid(strFormat, r, 1) = "[" Then Exit For
        Next
        strHZ = Mid(strFormat, InStr(i, strFormat, "{") + 1, r - InStr(i, strFormat, "{") - 1)
    End If
    '取出指定项目完整的数据串
    If strHZ <> "" Then
        j = InStr(intPos, strData, strHZ) '因为是连续取数,考虑到分隔符可能相同的情况,记录上一次的最后位置,下次从这个位置往后取数据
        If j = 0 Then
            '有可能中间存在回车换行符
            j = InStr(intPos, Replace(strData, vbCrLf, ""), strHZ)
            If j = 0 Then Exit Function
        End If
    End If
    strData = Mid(strData, intPos)
    '前缀为空,继续向前寻找上一个项目的后缀符号
'    If strQZ = "" And i > 1 And intPos > 1 Then
'        For i = i - 1 To 1 Step -1
'            If Mid(strFormat, i, 1) = "]" Then Exit For
'        Next
'        strQZ = Mid(strFormat, i + 1, InStr(i, strFormat, "}") - i - 1)
'    End If

    SubstrVal = SubstrAnaly(strData, strHZ, strQZ)
    intPos = intPos + Len(strQZ & SubstrVal & strHZ)
    '如果是数字型则去掉回车换行符返回,如果是字符型则原样返回
'    If strHZ <> "" Then
'
'        strData = Mid(strData, 1, InStr(1, Replace(strData, vbCrLf, ""), strHZ) - 1) '丢弃该项目后的数据
'        intPOS = i + Len(strHZ)
'    End If
'    If strQZ <> "" Then strData = Mid(strData, InStr(1, strData, strQZ) + Len(strQZ)) '丢弃该项目后的数据
'    SubstrVal = strData ' Replace(strData, vbCrLf, "")
End Function

Private Function SubstrAnaly(ByVal strData As String, ByVal strHZ As String, ByVal strQZ As String) As String
    Dim strText As String
    Dim strCompare As String           '对比串
    Dim intLen As Integer, intActLen As Integer           '前缀/后缀的长度
    Dim intPos As Integer, intEnd As Integer
    Dim lngASC As Long
    Dim blnFind As Boolean
    '遇到回车换行符忽略,空格重新比对

    strText = strData
    If strHZ <> "" Then
        '把后缀去掉
        strHZ = Replace(strHZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strHZ)
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strHZ Then
                        blnFind = True
                        intPos = intPos - intActLen
                    Else
                        strCompare = ""
                        intPos = intPos - intActLen + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        '肯定有
        strText = Mid(strText, 1, intPos)
    End If

    '再去掉前缀
    If strQZ <> "" Then
        If InStr(1, strText, strQZ) = 0 Then strText = strQZ & strText
        strQZ = Replace(strQZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strQZ)
        strCompare = ""
        intActLen = 0
        blnFind = False
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strQZ Then
                        blnFind = True
                        intPos = intPos + 1
                    Else
                        strCompare = ""
                        intPos = intPos + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        strText = Mid(strText, intPos)
    End If

    If IsNumeric(Replace(strText, vbCrLf, "")) Then
        SubstrAnaly = Replace(strText, vbCrLf, "")
    Else
        SubstrAnaly = strText
    End If
End Function

Public Sub SubstrPro(strFormat As String, ByVal strName As String, Optional ByVal intType As Integer = 0)
    Dim i As Integer, j As Integer, l As Integer, r As Integer
    'intType=0-删除指定格式串;1-得到指定格式串
    j = Len(strFormat)
    i = InStr(1, strFormat, "[" & strName & "]")
    If i = 0 Then Exit Sub

    For l = i To 1 Step -1
        If Mid(strFormat, l, 1) = "{" Then Exit For
    Next
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    If intType = 0 Then
        strFormat = Mid(strFormat, 1, l - 1) & Mid(strFormat, r + 1)
    Else
        strFormat = Mid(strFormat, l, r - l + 1)
    End If
End Sub

Private Sub MoveNextCell()
    Dim arrData
    Dim blnNULL As Boolean                      '是否为空行
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngMutilRows As Long, lngDeff As Long
    Dim intRow As Integer, intCount As Integer, intNULL As Integer  '其后有多少空行
    '赋值然后移动到下一个有效单元格

    '检查数据,不合格就再次弹出要求录入
    If mintType >= 0 Then
        If Not CheckInput(strReturn, strMsg) Then
            RaiseEvent AfterRowColChange(strMsg, True)
            Exit Sub
        End If

        lngMutilRows = 1
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
            lngMutilRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        End If
        lngStart = GetStartRow(VsfData.ROW)

        '准备赋值
        With txtLength
            '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
            .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime, 5000, VsfData.CellWidth)
            .Text = strReturn
            .FontName = VsfData.CellFontName
            .FontSize = VsfData.CellFontSize
        End With
        arrData = GetData(txtLength.Text)
        intCount = UBound(arrData)

        If intCount > lngMutilRows - 1 Then
            '往下搜索空行,如果有其它数据行则计算需增加的行数
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '保证当前输入的内容在一页中显示全
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For

                If Val(VsfData.TextMatrix(intRow + lngStart, c病人ID)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                    intNULL = intNULL - 1
                Else
                    Exit For
                End If
            Next
            '先增加空行
            If intNULL > 0 Then
                lngDeff = intNULL
                VsfData.Rows = VsfData.Rows + intNULL
                '从当前行记录的空白行开始，每行的位置+所增加的空白行数
                For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                    VsfData.RowPosition(intRow) = intRow + intNULL
                Next
            End If
            '循环赋值
            intCount = UBound(arrData)
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = arrData(intRow)
                VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intCount + 1 & "|" & intRow + 1
                VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intCount + 1
            Next
            '所有隐蔽列进行赋值
            lngMutilRows = lngStart + intCount
            For intRow = lngStart + 1 To lngMutilRows
                For intCount = 0 To VsfData.Cols - 1
                    VsfData.Cell(flexcpForeColor, intRow, intCount) = VsfData.Cell(flexcpForeColor, lngStart, intCount)
                    If VsfData.ColHidden(intCount) And InStr(1, "," & mlngRowCount & "," & mlngRowCurrent & ",", "," & intCount & ",") = 0 Then
                        VsfData.TextMatrix(intRow, intCount) = VsfData.TextMatrix(lngStart, intCount)
                    End If
                Next
            Next
        Else
            '对该列重新赋值（当只输入一个数字时，不知为何会产生字符ASCII码为1的符号）
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(arrData(intRow), Chr(1), "")
            Next
            For intRow = intCount + 1 To lngMutilRows - 1
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = ""
            Next

            '根据行数据重新填写行序列,intNULL记录最后一条不为空行的行号
            intNULL = lngStart + lngMutilRows - 1
            For intRow = lngMutilRows To 1 Step -1
                blnNULL = True
                For intCount = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(intCount) Then
                        If VsfData.TextMatrix(intRow + lngStart - 1, intCount) <> "" Then
                            blnNULL = False
                            Exit For
                        End If
                    End If
                Next

                If Not blnNULL Then Exit For
                intNULL = intNULL - 1
            Next
            '从新填写行序号
            For intRow = lngStart To intNULL
                VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
            Next
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                VsfData.TextMatrix(intRow, mlngRowCount) = ""
                VsfData.TextMatrix(intRow, mlngRowCurrent) = ""
            Next
        End If
        
        '当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
        If lngDeff <> 0 Then Call CellMap_Update(lngStart, lngDeff)

        If mstrData <> strReturn Then
            mblnChange = True

            '同步保存日期与时间列的数据
            Dim strKey As String, strField As String, strValue As String
            strField = "ID|页号|行号|列号|记录ID|数据|删除"
            '1\日期
            If mlngDate <> -1 Then
                strKey = mint页码 & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, mlngDate) & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\时间
            strKey = mint页码 & "," & lngStart & "," & mlngTime
            strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, mlngTime) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)

            '记录用户修改过的单元格
            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                strPart = GetActivePart(VsfData.COL, 0)
            Else
                strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
            End If

            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
            strKey = mint页码 & "," & lngStart & "," & VsfData.COL
            strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & VsfData.COL & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            
            Call SetActiveColColor
        End If
    End If

toNextCol:
    If VsfData.COL < mlngNoEditor - 1 Then       '护理记录单肯定有护士签名列
        VsfData.COL = VsfData.COL + 1
        If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Or _
            InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - cHideCols & ",") <> 0 Then
            GoTo toNextCol
        End If
    Else
toNextRow:
        '跳到下一行
        intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
        If VsfData.ROW + intRow < VsfData.Rows Then
            VsfData.ROW = VsfData.ROW + intRow
        End If
        If VsfData.RowHidden(VsfData.ROW) Then GoTo toNextRow
        VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
    End If
    If VsfData.ColIsVisible(VsfData.COL) = False Then
        VsfData.LeftCol = VsfData.COL
    End If
    If VsfData.RowIsVisible(VsfData.ROW) = False Then
        VsfData.TopRow = VsfData.ROW
    End If
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '提取数据起始行,超出本页则返回0
    '如果本页未显示全,则说明超出本页,也返回0
    '不允许在连续的数据行中插入新行

    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '总行数
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '当前行
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If

    '寻找起始行
    For lngRow = lngRow To 3 Step -1
        If VsfData.TextMatrix(lngRow, mlngRowCount) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next

    GetStartRow = lngStart
End Function

Private Function GetMutilData(ByVal lngRow As Long, ByVal lngCOL As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '起始行
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '返回第一行的坐标
    '不分行直接取，分行时检查如果当页显示全就拼接，否则从库中读取

    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
        GetMutilData = VsfData.TextMatrix(lngRow, lngCOL)
        Exit Function
    End If
    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
    lngCurRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1))

    If lngCount > 1 Then
        lngStart = GetStartRow(lngRow)
    Else
        lngStart = lngRow
    End If
    For lngRow = lngStart To lngStart + lngCount - 1
        strReturn = strReturn & VsfData.TextMatrix(lngRow, lngCOL)
    Next
    
    '取行高
    VsfData.ROW = lngStart
    dblHeight = lngCount * VsfData.RowHeightMin + 20
    dblTop = VsfData.Top + VsfData.CellTop

    GetMutilData = strReturn
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowInput(Optional ByVal intCol As Integer = -1, Optional ByVal strCellData As String = "", Optional ByVal blnAnalyse As Boolean = False) As String
    Dim arrData, arrValue
    Dim lngOrder As Long
    Dim i As Integer, j As Integer, intPos As Integer, intIndex As Integer
    Dim strFormat As String, strText As String, strValue As String  '格式串,数据串,数值串
    Dim strOrders As String, strTypes As String, strBounds As String, strLen As String, strName As String
    Const txtHeight = 300
    On Error GoTo errHand

    '病历文件构造管理模块需要处理:
    '1、一列绑定一个项目的不用管
    '2、一列绑定两个项目的，血压必须成对，要么都是录入，要么都是选择，不允许交叉出现，也不允许出现单选、复选
    '3、一列绑定多个项目的，只能是录入项目
    '由于以上条件限制，只取第一个项目的性质即可

    '如果是保存处调用则做如下处理
    If intCol = -1 Then intCol = VsfData.COL
    If blnAnalyse Then
        strText = strCellData
    Else
        '取当前单元格的属性
        CellRect.Left = VsfData.CellLeft + VsfData.Left
        CellRect.Top = VsfData.CellTop + VsfData.Top
        CellRect.Bottom = VsfData.CellHeight + 20
        CellRect.Right = VsfData.CellWidth + 20
        strText = GetMutilData(VsfData.ROW, intCol, CellRect.Top, CellRect.Bottom)
    End If
    mstrData = strText
    mintType = 0
    intIndex = 0

    '取当前列的绑定项目
    intPos = 1
    mrsSelItems.Filter = "列=" & intCol - cHideCols
    Do While Not mrsSelItems.EOF
        lngOrder = mrsSelItems!项目序号
        If lngOrder = 0 Then
            strLen = 0
            strValue = strText
            Exit Do
        End If

        '项目表示:2单选;3-多选;4-汇总;5-选择
        '项目值域:项目表示为0-表示最小值;最大值;项目表示为2,3-表示项目A;项目B,前有勾的表示缺省项
        strFormat = UCase(NVL(mrsSelItems!格式))
        strOrders = strOrders & "," & lngOrder
        If lngOrder <> 0 Then
            mrsItems.Filter = "项目序号=" & lngOrder
            strName = strName & "," & mrsItems!项目名称
            strLen = strLen & "," & mrsItems!项目长度 & ";" & NVL(mrsItems!项目小数)
            strTypes = strTypes & "," & mrsItems!项目表示
            strBounds = strBounds & "," & mrsItems!项目值域
            strValue = strValue & "'" & SubstrVal(strText, strFormat, GetActivePart(intCol, intIndex) & mrsItems!项目名称, intPos)

            Select Case mrsItems!项目表示
            Case 0  '文本录入项
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 2  '单选
                mintType = 1
            Case 3  '多选
                mintType = 2
            Case 4  '汇总
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 5  '选择
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 3
                Else
                    mintType = 5
                End If
            End Select
        Else
            strTypes = strTypes & ","
            strBounds = strBounds & ","
            strLen = strLen & ","
            strName = strName & ","
        End If

        intIndex = intIndex + 1
        mrsSelItems.MoveNext
    Loop
    If strOrders <> "" Then
        strOrders = Mid(strOrders, 2)
        strName = Mid(strName, 2)
        strLen = Mid(strLen, 2)
        strTypes = Mid(strTypes, 2)
        strBounds = Mid(strBounds, 2)
        strValue = Mid(strValue, 2)
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0

    If blnAnalyse Then
        ShowInput = strOrders & "||" & strValue
        Exit Function
    End If

    '针对4进行校对,如果表头文本不含/则处理为6
    If mintType = 4 Then
        If Not IsDiagonal(intCol) Then
            mintType = 6
        End If
    End If

    '判断当前列的性质
    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定2个及以上项目,手工录入
    arrValue = Split(strValue & "'", "'")
    Select Case mintType
    Case 0, 3
        With picInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        If mintType = 0 Then
            txtInput.Visible = True
            If Val(strLen) <> 0 And Val(strOrders) <> 10 Then
                txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) '小数位数要加上小数点
            Else
                txtInput.MaxLength = 0
            End If
            txtInput.Tag = lngOrder
        Else
            txtInput.Visible = False
        End If
        With txtInput
            .Top = 0
            .Text = strValue
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '宋体9号时减去90,字体越大扣除的边距越小,以保证文本框分行与实际一致
        End With
        With lblInput
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = CellRect.Bottom
            .Width = CellRect.Right
            .Top = 50
            .Tag = lngOrder
            .Caption = strValue
            .Visible = (mintType = 3)
        End With

        '如果是日期或时间列，设定固定值
        If mintType = 0 And txtInput.Text = "" Then
            If intCol = mlngDate Then
                If mblnDateAd Then
                    txtInput.Text = Format(zlDatabase.Currentdate, "d-M")
                    txtInput.Text = Replace(txtInput.Text, "-", "/")
                Else
                    txtInput.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
                End If
            ElseIf intCol = mlngTime Then
                txtInput.Text = Format(zlDatabase.Currentdate, "HH:mm")
            End If
        End If
    Case 1, 2
        '加载数据
        lstSelect(mintType - 1).Clear
        arrData = Split(strBounds, ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "√" Then
                    lstSelect(mintType - 1).AddItem i + 1 & Mid(arrData(i), 2)
                    If strText = "" Then lstSelect(mintType - 1).ListIndex = i
                Else
                    lstSelect(mintType - 1).AddItem i + 1 & arrData(i)
                End If
            End If
        Next
        '多选且已录入数据的情况下
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(mintType - 1).List(i), 2) & ",") <> 0 Then
                    lstSelect(mintType - 1).Selected(i) = True
                End If
            Next
        End If
        '显示
        With lstSelect(mintType - 1)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Tag = lngOrder
            .Visible = True
        End With
    Case 4, 5
        With picDouble
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDouble.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If

        With txtUpInput
            .Text = arrValue(0)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = (picDouble.Width - lblSplit.Width) * 0.4
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(0)
        End With
        With picUpInput
            .Left = txtUpInput.Left
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(0)
        End With
        With lblUpInput
            .Alignment = 2
            .Caption = arrValue(0)
            .Left = 0
            .Top = 50
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .Tag = Split(strOrders, ",")(0)
        End With
        With txtDnInput
            .Text = arrValue(1)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Left = lblSplit.Left + lblSplit.Width
            .Width = picDouble.Width - .Left
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(1)
        End With
        With picDnInput
            .Left = txtDnInput.Left
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(1)
        End With
        With lblDnInput
            .Alignment = 2
            .Caption = arrValue(1)
            .Left = 0
            .Top = 50
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Tag = Split(strOrders, ",")(1)
        End With

        If mintType = 4 Then
            If strLen <> "" Then txtUpInput.MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
            If strLen <> "" Then txtDnInput.MaxLength = Val(Split(Split(strLen, ",")(1), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(1), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
        End If
    Case 6
        '先删除以前的控件
        j = txt.Count - 1
        For i = 1 To j
            Unload lbl(i)
            Unload txt(i)
        Next
        '设定坐标
        With picMutilInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = IIf(CellRect.Right < 1600, 1600, CellRect.Right)
        End With
        '对缺省控件赋值
        arrData = Split(strOrders, ",")
        j = UBound(arrData)
        lbl(0).Top = 130
        lbl(0).Caption = Split(strName, ",")(0)
        lbl(0).FontName = VsfData.FontName
        lbl(0).FontSize = VsfData.FontSize
        txt(0).Tag = arrData(0)
        txt(0).FontName = VsfData.FontName
        txt(0).FontSize = VsfData.FontSize
        txt(0).Width = picMutilInput.Width - txt(0).Left - 100
        txt(0).MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1)  '小数位数要加上小数点
        txt(0).Text = arrValue(0)
        If Not mblnBlowup Then
            txt(0).Height = 225
        End If

        '加载控件
        For i = 1 To j
            Load lbl(i)
            With lbl(i)
                .Caption = Split(strName, ",")(i)
                .Left = lbl(0).Left + lbl(0).Width - .Width
                .Top = lbl(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Visible = True
            End With
            Load txt(i)
            With txt(i)
                .TabIndex = txt(i - 1).TabIndex + 1
                .Left = txt(0).Left
                .Top = txt(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Tag = arrData(i)
                If strLen <> "" Then
                    .MaxLength = Val(Split(Split(strLen, ",")(i), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(i), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
                End If
                .Text = arrValue(i)
                .Visible = True
            End With
        Next

        With picMutilInput
            .Height = txt(j).Top + txt(j).Height + 120
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
    End Select
    Exit Function

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CheckFormat(ByVal strNames As String, ByVal strFormat As String)
    '如果格式与血压的方式不同,则将样式处理为6

    '去掉前缀后进行对比
    strFormat = Mid(strFormat, InStr(1, strFormat, "["))
    strFormat = Replace(strFormat, "[", "")
    strFormat = Replace(strFormat, "]", "")
    If Not (strFormat Like Split(strNames, ",")(0) & "/}*" Or strFormat Like "{/*" & Split(strNames, ",")(1)) Then
        mintType = 6
    End If
End Sub

Private Function IsDiagonal(ByVal intCol As Integer) As Boolean
    Dim arrCol, arrData
    Dim intDo As Integer, intCount As Integer
    '判断指定列是否设置了列对角线（mstrColWidth的格式：765`11`1`1,765`11`2`1,...，对象属性`对象序号`列对角线）

    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & intCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function

Private Sub ISAssistant(ByVal lngOrder As Long, ByVal objTXT As TextBox)
    Dim intIndex As Integer
    Dim objParent As Object
    '根据项目的长度决定是否允许进行词句选择
    mblnEditAssistant = False
    cmdWord.Visible = mblnEditAssistant

    mrsItems.Filter = "项目序号=" & lngOrder
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        Exit Sub
    End If
    mblnEditAssistant = (mrsItems!项目长度 > 100)
    mrsItems.Filter = 0

    '如果允许词句选择,显示并定位
    If mblnEditAssistant Then
        If UCase(objTXT.Name) = "TXTINPUT" Then
            intIndex = -1 '表示txtInput
            Set objParent = picInput
        Else
            intIndex = objTXT.Index
            Set objParent = picMutilInput
        End If
        With cmdWord
            .Tag = intIndex
            .Top = objParent.Top + objTXT.Top + 25
            .Left = objParent.Left + objTXT.Left + objTXT.Width - .Width + 25
            .Visible = True
        End With
    End If
End Sub

Private Sub FillPage()
    Dim strPatient As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '读取符合条件的病人清单(在院病人+最近几天转科病人+指定时间范围内出院病人),病人清单决定了行数
    
    '在院病人清单
    strPatient = "" & _
        " SELECT 1 AS 性质,B.病人ID, B.主页ID, A.姓名, A.性别, B.住院号, B.出院病床 AS 床号,0 AS 婴儿" & _
        " FROM 病人信息 A,病案主页 B" & _
        " Where A.病人ID = b.病人ID And NVL(b.主页ID, 0) <> 0 And b.当前病区ID + 0 = [3]" & _
        " AND A.在院=1 AND B.出院日期 IS NULL AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL" & _
        IIf(mlng科室ID = -1, "", " And B.出院科室ID+0=[4]")
    If chk出院.Value = 1 Then
        '最近几天出院病人清单
        strPatient = strPatient & _
            " UNION " & _
            " SELECT 3 AS 性质,B.病人ID, B.主页ID, A.姓名, A.性别, B.住院号, B.出院病床 AS 床号,0 AS 婴儿" & _
            " FROM 病人信息 A,病案主页 B" & _
            " Where A.病人ID = b.病人ID And NVL(b.主页ID, 0) <> 0 And b.当前病区ID + 0 = [3]" & _
            " AND B.出院日期 BETWEEN [1] AND [2] AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL" & _
            IIf(mlng科室ID = -1, "", " And B.出院科室ID+0=[4]")
    End If
    If chk出科.Value = 1 Then
        '最近几天转科病人清单
        strPatient = strPatient & _
            " UNION " & _
            " Select 2 AS 性质,B.病人ID, B.主页ID, A.姓名, A.性别, B.住院号, C.床号,0 AS 婴儿" & _
            " From 病人信息 A,病案主页 B,病人变动记录 C" & _
            " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 " & _
            " And Nvl(B.状态,0)<>2 And Nvl(C.附加床位,0)=0 " & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID And C.病区ID+0=[3]" & IIf(mlng科室ID = -1, "", " And B.出院科室ID<>[4] And C.科室ID+0=[4]") & _
            " And C.终止原因=3 And C.终止时间 Between Sysdate-" & mintChange & " And Sysdate" & _
            " And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
    End If
    '提取新生儿列表
    strPatient = strPatient & _
              " UNION " & _
              " Select B.性质,B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,B.住院号,B.床号,A.序号 AS 婴儿" & _
              " From 病人新生儿记录 A,(" & strPatient & ") B" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    
    gstrSQL = " SELECT /*+ RULE */ A.性质,A.病人ID,A.主页ID,A.婴儿,A.姓名,A.床号,MAX(B.ID) AS 文件ID" & _
              " FROM (" & strPatient & ") A,病人护理文件 B" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.婴儿=B.婴儿 " & _
              " And B.归档人 is null And B.结束时间 is null And B.格式ID=[5]" & _
              " GROUP BY A.性质,A.病人ID,A.主页ID,A.婴儿,A.姓名 ,A.床号" & _
              " Order by A.性质,A.床号"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人清单", mdtOutbegin, mdtOutEnd, mlng病区ID, mlng科室ID, mlng格式ID)
    
    '填充数据到表格
    With rsTemp
        Do While Not .EOF
            If .AbsolutePosition > VsfData.Rows - VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
            
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c文件ID) = !文件ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c床号) = NVL(!床号)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c姓名) = IIf(!婴儿 > 0, Space(4), "") & !姓名
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c病人ID) = !病人ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c主页ID) = !主页ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c婴儿) = !婴儿
            .MoveNext
        Loop
    End With
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


'######################################################################################################################
'**********************************************************************************************************************
'以下是基础函数或过程
Private Sub lblDnInput_Click()
    txtDnInput.SetFocus
End Sub

Private Sub lblUpInput_Click()
    txtUpInput.SetFocus
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    mblnEditAssistant = False
End Sub

Private Sub txtDnInput_GotFocus()
    txtDnInput.SelStart = 0
    txtDnInput.SelLength = 100
    Call ISAssistant(Val(txtDnInput.Tag), txtDnInput)
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
    mintSymbol = -1
    Call ISAssistant(Val(txtInput.Tag), txtInput)
End Sub

Private Sub txtUpInput_GotFocus()
    txtUpInput.SelStart = 0
    txtUpInput.SelLength = 100
    Call ISAssistant(Val(txtUpInput.Tag), txtUpInput)
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = 100
    mintSymbol = Index
    Call ISAssistant(Val(txt(Index).Tag), txt(Index))
End Sub

Private Sub lblUpInput_DblClick()
    lblUpInput.Caption = IIf(lblUpInput.Caption = "", "√", "")
    txtUpInput.SetFocus
End Sub

Private Sub lblDnInput_DblClick()
    lblDnInput.Caption = IIf(lblDnInput.Caption = "", "√", "")
    txtDnInput.SetFocus
End Sub

Private Sub lblInput_DblClick()
    lblInput.Caption = IIf(lblInput.Caption = "", "√", "")
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        If txtUpInput.SelStart = Len(txtUpInput.Text) Then txtDnInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtUpInput.Locked Then
        lblUpInput.Caption = IIf(lblUpInput.Caption = "", "√", "")
    End If
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call picDouble_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then txtUpInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtDnInput.Locked Then
        lblDnInput.Caption = IIf(lblDnInput.Caption = "", "√", "")
    End If
End Sub

Private Sub picMutilInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picDouble_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    End If
End Sub

Private Sub picInput_GotFocus()
    If txtInput.Visible Then
        txtInput.SetFocus
    End If
End Sub

Private Sub picInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not txtInput.Visible Then
        If KeyCode = vbKeySpace Then
            Call lblInput_DblClick
        End If
    End If

    If KeyCode = vbKeyReturn Then
        '移动到下一个单元格
        Call MoveNextCell
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picMutilInput_GotFocus()
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < txt.Count - 1 Then
            txt(Index + 1).SetFocus
        Else
            Call picMutilInput_KeyDown(KeyCode, Shift)
        End If
    End If
End Sub

Private Sub picDouble_GotFocus()
    If txtUpInput.Visible Then
        txtUpInput.SetFocus
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    VsfData.Width = picMain.Width
    VsfData.Height = picMain.Height - VsfData.Top
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Exit Sub
    Call picInput_KeyDown(KeyCode, Shift)
End Sub

Private Sub txtUpInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("/") Then
        KeyAscii = 0
        txtDnInput.SetFocus
    End If
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnSigned = False
    mblnSaved = False
    mblnChange = False
    mblnInit = False

'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '以下字符做为数据分隔符或更新记录集的分隔符，因此不允许录入
    If KeyAscii = 39 Or KeyAscii = 13 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call InitCons
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)

    Err = 0: On Error Resume Next
    lblTitle.Move lngScaleLeft, 120, lngScaleRight - lngScaleLeft
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom - lngScaleTop
    VsfData.Move lngScaleLeft + 210, lblTitle.Top + lblTitle.Height + 300, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top

    '表上标签分散处理
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long

    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值

    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '更新记录,如果不存在,则新增
    'strPrimary:字段名|值
    'strFields:字段名|字段名
    'strValues:值|值

    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'strPrimary = "RecordID|5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False

    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate

    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCol As Integer, intCols As Integer
    Dim strValues As String
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strValues = ""
            intCols = .Fields.Count - 1
            For intCol = 0 To intCols
                strValues = strValues & "," & .Fields(intCol).Name & ":" & .Fields(intCol).Value
            Next
            Debug.Print Mid(strValues, 2)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '放大：字体，单元格宽度
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function

Private Sub SetActiveColColor()
    '活动列的背景色设置为灰色,表示不允许编辑
    Dim aryItem, lngRow As Long
    aryItem = Split(mstrCOLNothing, ",")
    For lngRow = 0 To UBound(aryItem)
        VsfData.Cell(flexcpBackColor, VsfData.FixedRows, Val(aryItem(lngRow)) + cHideCols, VsfData.Rows - 1, Val(aryItem(lngRow)) + cHideCols) = &H8000000F
        '.ColHidden(Val(aryItem(lngCount)) + cHideCols) = True
    Next
End Sub
