VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl usrBodyEditor 
   ClientHeight    =   8100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   8100
   ScaleWidth      =   10455
   Begin VB.PictureBox picSerach 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   1515
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1515
      Begin VB.Label lbl查看 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "原始大小"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   420
         TabIndex        =   32
         Top             =   60
         Width           =   960
      End
      Begin VB.Image imgPic 
         Height          =   360
         Left            =   60
         Picture         =   "usrBodyEditor.ctx":0000
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7095
      ScaleWidth      =   10215
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   10215
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   9600
         TabIndex        =   29
         Top             =   4920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1179648
      End
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   7200
         TabIndex        =   28
         Top             =   6120
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1179649
      End
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   240
         ScaleHeight     =   360
         ScaleWidth      =   2730
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   6120
         Width           =   2730
         Begin VB.ComboBox cboBaby 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   0
            Width           =   1920
         End
         Begin VB.Label lblSerach 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "查看"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   9
            Left            =   345
            TabIndex        =   24
            Top             =   105
            Width           =   360
         End
      End
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5895
         Left            =   120
         ScaleHeight     =   5895
         ScaleWidth      =   9375
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   9375
         Begin VB.PictureBox picDraw 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   120
            ScaleHeight     =   1215
            ScaleWidth      =   7335
            TabIndex        =   33
            Top             =   2160
            Width           =   7335
         End
         Begin zl9TemperatureChart.VsfGrid vsf 
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   3480
            Width           =   7215
            _ExtentX        =   12515
            _ExtentY        =   450
         End
         Begin VB.PictureBox picCard 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   810
            Index           =   0
            Left            =   120
            ScaleHeight     =   810
            ScaleWidth      =   8640
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   120
            Width           =   8640
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   7
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   10
               TabStop         =   0   'False
               Text            =   "诊断"
               Top             =   375
               Width           =   2370
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   6
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   9
               TabStop         =   0   'False
               Text            =   "年龄"
               Top             =   60
               Width           =   645
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   5
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   8
               TabStop         =   0   'False
               Text            =   "性别"
               Top             =   60
               Width           =   420
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   4
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   7
               TabStop         =   0   'False
               Text            =   "12"
               Top             =   375
               Width           =   615
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   3
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   6
               TabStop         =   0   'False
               Text            =   "入院日期"
               Top             =   60
               Width           =   1140
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   2
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   5
               TabStop         =   0   'False
               Text            =   "科室"
               Top             =   375
               Width           =   2400
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   1
               Left            =   6645
               Locked          =   -1  'True
               TabIndex        =   4
               TabStop         =   0   'False
               Text            =   "1234567"
               Top             =   60
               Width           =   3825
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   0
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   3
               TabStop         =   0   'False
               Text            =   "姓无名"
               Top             =   60
               Width           =   1425
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "诊    断:"
               Height          =   180
               Index           =   7
               Left            =   4065
               TabIndex        =   18
               Top             =   390
               Width           =   810
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄:"
               Height          =   180
               Index           =   6
               Left            =   2910
               TabIndex        =   17
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别:"
               Height          =   180
               Index           =   4
               Left            =   1980
               TabIndex        =   16
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院日期:"
               Height          =   180
               Index           =   5
               Left            =   4050
               TabIndex        =   15
               Top             =   60
               Width           =   810
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "床号:"
               Height          =   180
               Index           =   3
               Left            =   2910
               TabIndex        =   14
               Top             =   390
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科室:"
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   13
               Top             =   375
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号:"
               Height          =   180
               Index           =   1
               Left            =   6000
               TabIndex        =   12
               Top             =   60
               Width           =   630
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   11
               Top             =   60
               Width           =   450
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid mshDownTab 
            Height          =   975
            Left            =   90
            TabIndex        =   20
            Top             =   3840
            Width           =   7215
            _cx             =   12726
            _cy             =   1720
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   18
            FixedRows       =   0
            FixedCols       =   4
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"usrBodyEditor.ctx":076A
            ScrollTrack     =   0   'False
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   0
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
            AutoSizeMouse   =   -1  'True
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
         Begin VSFlex8Ctl.VSFlexGrid mshUpTab 
            Height          =   1095
            Left            =   120
            TabIndex        =   19
            Top             =   960
            Width           =   7275
            _cx             =   12832
            _cy             =   1931
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   0
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   0
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
            AutoSizeMouse   =   -1  'True
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
            Begin VB.PictureBox picDisplay 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   240
               ScaleHeight     =   165
               ScaleWidth      =   165
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   60
               Width           =   165
               Begin VB.Image imgDisPlay 
                  Appearance      =   0  'Flat
                  Height          =   240
                  Left            =   -30
                  Picture         =   "usrBodyEditor.ctx":08D7
                  Stretch         =   -1  'True
                  Top             =   -30
                  Width           =   240
               End
            End
            Begin VB.Label lblCur 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "△"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   240
               TabIndex        =   27
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VB.Label lblCommText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "说明"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   240
            TabIndex        =   21
            Top             =   4920
            Width           =   360
         End
      End
      Begin VB.PictureBox picBuffer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1785
         Left            =   7200
         ScaleHeight     =   117
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   139
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "临时拷图用,千万别删"
         Top             =   2640
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrBodyEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mstrTitle As String = "体温绘图"
'--菜单
Private mcbrToolBar页面 As CommandBarControl
Private mcbrTools   As CommandBar
Private mcbrToolBar As CommandBar
Private mcbrItem         As CommandBarControl

'--变量
Public mblnResize As Boolean '记录窗体大小是否发生变化
Public mblnMoved As Boolean
Private mlngWidth As Long
Private mlngHeight As Long
Private mintPage      As Integer '记录当前页号
Private mintAllPage As Integer '体温单的总页数
Private mstrParam  As String, mstrParam1 As String, mstrParam2 As String
Private mfrmParent As Object
Private mIntDataEditor As Integer '0 表示调用体温单数据编辑 1表示调用体温单数据显示设置
Private mstrSQL As String
Private mintColMin, mintColMax As Integer
Private mint心率应用 As Double
Private msinVStep As Single      '滚动条的步长
Private msinHStep As Single      '滚动条的步长
Private mblnAutoAdjust As Boolean  '控制体温单格式 ：体温单是否跟随窗体大小自动调整
Private mblnAutoRedraw As Boolean  '控制是否自动重画:是否自动完成重画,内容包括:初始画布,读取数据并整理,绘画,
Private mblnRefresh    As Boolean
Private mintOpDays As Integer '手术标志天数
Private mblnStopFlag As Boolean '手术停止标志
Private mintOpFormat As Integer '手术当天缺省格式 0-不显示;1-显示0;2-显示手术次数
Private mintRepairRows As Integer '体温表格固定输出行数
Private mbln显示皮试 As Boolean
Private mblnKeyDown As Boolean
Private mstrOpdays(1 To 7) As String
Private mstrOpValue(1 To 7) As String
Private mstrNewString() As String '保存皮试结果信息
Private mlng高度 As Single '体温单刻度区域可显示的高度范围
Private mbln出院 As Boolean '病人是否出院
Private mbytSize As Byte '字体大小 0-9号字体 1-12号字体
'--体温单时间
Private mstr开始时间 As String  '一周开始时间
Private mstr结束时间 As String  '一周结束时间
Private mstrEnterDate As String '体温单开始时间
Private mstrEndDate As String   '体温单结束时间
Private mstrComeInDate As String '病人入院时间

'参数
Private mbln灌肠大便分子分母显示 As Boolean

Private WithEvents mfrmCaseTendBodyPrint As frmCaseTendBodyPrint
Attribute mfrmCaseTendBodyPrint.VB_VarHelpID = -1

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2 '浅凹下
Private Const BDR_RAISEDINNER = &H4 '浅凸起
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)


'***************************************************************
'体温单绘画相关变量
'***************************************************************
Private mobjDraw          As Object
Private mobjBuffer        As Object                     '缓冲,额外使用
Private mblnRedraw        As Boolean                    '是否需要重画
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
Private mlngFont           As Long
Private mlngOldFont       As Long


'***************************************************************
'体温单绘画相关内存映射记录集
'***************************************************************
Private mrsItems As New ADODB.Recordset              '所有项目的属性
Private mrsGraph As New ADODB.Recordset              '所有需输出的图形序号(全部提取在picBuffer中,此处保存各项目的部位及其对应的图形序号)
Private mrsDrawItems As New ADODB.Recordset              '所有曲线项目的有效数据区域(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
Private mrsPoint As New ADODB.Recordset              '所有点的表现集合
Private mrsNote  As New ADODB.Recordset              '文本输出集合,可指定颜色

Private Type Type_NO
    血压 As Integer
    舒张压 As Integer
    呼吸 As Integer
End Type

Private mItemNO As Type_NO

Private Type Type_row
    血压 As Integer
    饮入量 As Integer
    排出量 As Integer
End Type
Private mItemRow As Type_row

Private Type T_LPoint
    X As Long
    Y As Long
    W As Single
End Type

'***************************************************************
'病人基本信息
'***************************************************************
Private Type type_Patient
    lng病人ID As Long
    lng主页ID As Long
    lng病区ID As Long
    lng科室ID As Long
    lng出院 As Long
    lng婴儿 As Long
    lng编辑 As Long
    lng护理等级 As Long
    lng文件ID As Long
    lng原始大小 As Long
    lngPage As Long
End Type
Private T_Patient As type_Patient

'--事件定义
Public Event CmdClick(ByVal strParam As String)
Public Event zlAfterPrint()
Public Event DbClickCur(ByVal intDataEditor As Integer)

Public Property Get ParentForm() As Object
    Set ParentForm = mfrmParent
End Property

Public Property Set ParentForm(objParent As Object)
    Set mfrmParent = objParent
End Property

Public Property Get ScrollBarY() As FlatScrollBar
    Set ScrollBarY = vsb
End Property

Public Property Get ScrollBarX() As FlatScrollBar
    Set ScrollBarX = hsb
End Property

Public Property Get DateEditor() As Integer
     DateEditor = mIntDataEditor
End Property

Public Property Let DateEditor(intDataEditor As Integer)
     mIntDataEditor = intDataEditor
End Property

Public Property Let lng病人ID(lng病人ID As Long)
     T_Patient.lng病区ID = lng病人ID
End Property

Public Property Let lng主页ID(lng主页ID As Long)
     T_Patient.lng主页ID = lng主页ID
End Property

Public Property Let lng文件ID(lng文件ID As Long)
     T_Patient.lng文件ID = lng文件ID
End Property

Public Property Let lng科室ID(lng科室ID As Long)
     T_Patient.lng科室ID = lng科室ID
End Property

Public Property Let lng婴儿(lng婴儿 As Long)
     T_Patient.lng婴儿 = lng婴儿
End Property

Public Property Let intPage(intPage1 As Long)
     mintPage = intPage1
End Property

Public Property Get intPage() As Long
    intPage = mintPage + 1
End Property

Public Property Get AllPage() As Integer
    AllPage = mintAllPage
End Property

Public Property Get FontSize() As Byte
    FontSize = mbytSize
End Property

Public Property Let FontSize(bytSize As Byte)
     mbytSize = bytSize
End Property

Private Function InitCommandBar() As Boolean

    '******************************************************************************************************************
    '功能：初始化菜单按钮
    '参数：
    '返回：
    '******************************************************************************************************************

    Dim objCustom  As CommandBarControlCustom
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo Errhand
    '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
'    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003

    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CallPrevious, "上一页")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CallNext, "下一页")
    End With

    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    
    Set mcbrToolBar = cbsMain.Add("婴儿", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    Set objCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Option, "")
    objCustom.flags = xtpFlagAlignLeft
    picTmp.Visible = True
    objCustom.Handle = picTmp.hWnd
    
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyLeft, conMenu_Manage_CallPrevious      '前一条
        .Add FCONTROL, vbKeyRight, conMenu_Manage_CallNext    '后一条
    End With

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function InitBody(ByVal lng文件ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer) As Boolean

    '******************************************************************************************************************
    '功能：提取病人护理时间范围
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSql        As String, strNewSql As String
    Dim strParam1 As String '如果指定了页数 次数记录对应页号的参数值
    Dim RS            As New ADODB.Recordset
    Dim rsTmp         As New ADODB.Recordset
    Dim ArrControlId() As Variant
    Dim cbrPre  As CommandBarButton
    Dim cbrWeek As CommandBarButton
    Dim objCostom As CommandBarControlCustom
    Dim intCount      As Integer
    Dim strDateFrom   As String '每一页 开始时间
    Dim strDateTo     As String '每一页 结束时间
    Dim strEnterDate  As String  '入院时间
    Dim strOutDate  As String   '终止时间
    Dim strMarkDate As String '体温单设置时间
    Dim intCOl        As Integer
    Dim strCaption    As String
    Dim strParameter  As String
    Dim strSvrCaption As String, strSvrCaption1 As String
    Dim strNow        As String
    Dim strCut        As String
    Dim lngLoop       As Long
    Dim strTmp        As String
    Dim lnglast科室id As Long
    
    On Error GoTo Errhand

    If lng病人ID = 0 And lng文件ID = 0 And lng主页ID = 0 Then Exit Function
    mbln出院 = False
    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNow = Format(zldatabase.Currentdate, "yyyy-MM-dd")
    '删除操作页面菜单项

    If Not mcbrToolBar页面 Is Nothing Then mcbrToolBar页面.Delete
    Set mcbrToolBar页面 = mcbrToolBar.Controls.Add(xtpControlPopup, conMenu_Edit_NewItem, "页面"):  mcbrToolBar页面.BeginGroup = True
    mcbrToolBar页面.IconId = conMenu_Edit_Modify
    mcbrToolBar页面.Style = xtpButtonIconAndCaption
    
    
    ArrControlId = Array(conMenu_View_OneWeek, conMenu_View_TwotWeek, conMenu_View_ThreeWeek, conMenu_View_FourWeek, _
        conMenu_View_Forward, conMenu_View_Backward)

    For lngLoop = 0 To UBound(ArrControlId)
        If Not mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop))) Is Nothing Then mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop))).Delete
    Next lngLoop
    
    
    With mcbrToolBar.Controls
        
        '加载上下页
        Set cbrPre = .Add(xtpControlButton, conMenu_View_Forward, "上一页", -1, False)
        Set cbrPre = .Add(xtpControlButton, conMenu_View_Backward, "下一页", -1, False)
        
        '加载 4个周期 此处默认为入院时间开始4个周
        Set cbrPre = .Add(xtpControlButton, conMenu_View_OneWeek, " " & 1 & " ", -1, False)
        cbrPre.ToolTipText = "第一周"
        Set cbrPre = .Add(xtpControlButton, conMenu_View_TwotWeek, " " & 2 & " ", -1, False)
        cbrPre.ToolTipText = "第二周"
        Set cbrPre = .Add(xtpControlButton, conMenu_View_ThreeWeek, " " & 3 & " ", -1, False)
        cbrPre.ToolTipText = "第三周"
        Set cbrPre = .Add(xtpControlButton, conMenu_View_FourWeek, " " & 4 & " ", -1, False)
        cbrPre.ToolTipText = "第四周"
    End With
    
    If Not mcbrToolBar.FindControl(, conMenu_ViewPopup) Is Nothing Then
        mcbrToolBar.FindControl(, conMenu_ViewPopup).Delete
    End If
    
    If mblnAutoAdjust = True Then
        Set objCostom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_ViewPopup, "查看原图")
        objCostom.Handle = picSerach.hWnd
    Else
        picSerach.Visible = False
    End If
    
    '提取用户设置的体温单开始时间(婴儿还是以婴儿出生时间为准)
    strSql = "select 开始时间 from 病人护理文件 where ID=[1] and 病人ID=[2] and 主页id=[3] and nvl(婴儿,0)=[4]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "提取体温单开始时间", lng文件ID, lng病人ID, lng主页ID, int婴儿)
    If rsTmp.RecordCount <> 0 Then
        strEnterDate = Format(rsTmp!开始时间, "YYYY-MM-DD HH:mm:ss")
    End If
    
    '提取婴儿医嘱信息(转科，出院)存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "   (SELECT /*+ RULE */  病人ID,主页ID,婴儿时间,DECODE(nvl(婴儿,0),0, DECODE(NVL(出院日期,''),'',0,1), DECODE(NVL(婴儿时间,''),'',0,1))记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID,A.主页ID,B.开始执行时间 婴儿时间, A.出院日期,B.婴儿" & vbNewLine & _
                "           FROM 病案主页 A," & vbNewLine & _
                "               (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND nvl(B.婴儿,0)<>0 AND C.类别 = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.操作类型 = COLUMN_VALUE) And  B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "           WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "           ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"

    strMarkDate = "to_date('" & strEnterDate & "','yyyy-MM-dd hh24:mi:ss')"
    '------------------------------------------------------------------------------------------------------------------
    '提取体温单页数（婴儿出院时间需检查是否存在医嘱,存在以医嘱时间为准，否则以母亲的为准）
    strSql = "SELECT DECODE(C.出生时间,NULL," & IIf(strEnterDate = "", "B.入院时间", strMarkDate) & ",C.出生时间) AS 入院时间," & vbNewLine & _
                " DECODE(C.出生时间,NULL,B.入院时间,C.出生时间) AS 实际入院时间," & vbNewLine & _
                " DECODE(E.记录,0,DECODE(SIGN(nvl(E.婴儿时间,B.出院时间) - D.发生时间), 1,nvl(E.婴儿时间,B.出院时间) ,D.发生时间),nvl(E.婴儿时间,B.出院时间))  出院时间," & vbNewLine & _
                " 1 + TRUNC((TO_DATE(TO_CHAR(DECODE(E.记录,0,DECODE(SIGN(nvl(E.婴儿时间,B.出院时间) - D.发生时间), 1,nvl(E.婴儿时间,B.出院时间) ,D.发生时间),nvl(E.婴儿时间,B.出院时间)),'yyyy-MM-dd'),'yyyy-MM-dd') - " & vbNewLine & _
                " TO_DATE(TO_CHAR(DECODE(C.出生时间,NULL," & IIf(strEnterDate = "", "B.入院时间", strMarkDate) & ",C.出生时间),'yyyy-MM-dd'),'yyyy-MM-dd')) / 7) AS 页数,D.发生时间,E.记录" & vbNewLine & _
                "    FROM (SELECT 病人ID,主页ID,MIN(开始时间) AS 入院时间,MAX(NVL(终止时间, SYSDATE)) AS 出院时间" & vbNewLine & _
                "    FROM 病人变动记录" & vbNewLine & _
                "    WHERE 开始时间 IS NOT NULL AND 病人ID = [2] AND 主页ID =[3] GROUP BY 病人ID,主页ID) B," & vbNewLine & _
                "    (SELECT 病人ID,主页ID,出生时间 FROM 病人新生儿记录 WHERE 病人ID = [2] AND 主页ID = [3] AND 序号=[4]) C," & vbNewLine & _
                "    (SELECT NVL(发生时间,SYSDATE) 发生时间 FROM (SELECT MAX(发生时间) 发生时间 FROM 病人护理文件 A,病人护理数据 B" & vbNewLine & _
                "           WHERE A.ID=B.文件ID AND A.ID=[1] AND A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=[4])) D," & vbNewLine & _
                strNewSql & vbNewLine & _
                "    WHERE B.病人ID=E.病人ID And B.主页ID=E.主页ID And B.病人ID=C.病人ID(+) AND B.主页ID=C.主页ID(+)"

    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "usrBodyEditor", lng文件ID, lng病人ID, lng主页ID, int婴儿)
    If rsTmp.BOF Then
        MsgBox "无病人本次住院记录！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    mintAllPage = rsTmp("页数").Value
    If T_Patient.lngPage > mintAllPage Then T_Patient.lngPage = 0
    
    If strEnterDate = "" Then strEnterDate = Format(rsTmp!入院时间, "yyyy-MM-dd HH:mm:ss")
    mstrEnterDate = strEnterDate
    mstrComeInDate = Format(rsTmp!实际入院时间, "yyyy-MM-dd HH:mm:ss")
    
    strOutDate = Format(rsTmp!出院时间, "yyyy-MM-dd HH:mm:ss")
    mbln出院 = Not (Val(zlCommFun.Nvl(rsTmp!记录)) = 0)
    
    '------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT 1 + TRUNC((TO_DATE(TO_CHAR(A.开始时间,'yyyy-MM-dd'),'yyyy-MM-dd') - TO_DATE(TO_CHAR(B.入院时间,'yyyy-MM-dd'),'yyyy-MM-dd')) / 7) AS 开始页码," & vbNewLine & _
            "1 + TRUNC((TO_DATE(TO_CHAR(DECODE(A.序号,F.LAST,DECODE(E.记录,0,DECODE(SIGN(nvl(E.婴儿时间,A.终止时间) - D.发生时间), 1,nvl(E.婴儿时间,A.终止时间) ,D.发生时间),nvl(E.婴儿时间,A.终止时间)),nvl(E.婴儿时间,A.终止时间)),'yyyy-MM-dd'),'yyyy-MM-dd') - TO_DATE(TO_CHAR(B.入院时间,'yyyy-MM-dd'),'yyyy-MM-dd')) / 7) AS 结束页码," & vbNewLine & _
            "      B.入院时间,D.发生时间,病区ID,C.名称,A.开始时间,DECODE(A.序号,F.LAST,DECODE(E.记录,0,DECODE(SIGN(nvl(E.婴儿时间,A.终止时间) - D.发生时间), 1,nvl(E.婴儿时间,A.终止时间) ,D.发生时间),nvl(E.婴儿时间,A.终止时间)),nvl(E.婴儿时间,A.终止时间))  终止时间" & vbNewLine & _
            "FROM (SELECT ROWNUM 序号, 病区ID,开始时间,终止时间" & vbNewLine & _
            "      FROM(SELECT  病区ID,MIN(开始时间) AS 开始时间,MAX(NVL(终止时间, SYSDATE)) AS 终止时间" & vbNewLine & _
            "           FROM 病人变动记录" & vbNewLine & _
            "                WHERE 开始时间 IS NOT NULL AND 病人ID =[2] AND 主页ID =[3] GROUP BY 病区ID  ORDER BY 开始时间)) A," & vbNewLine & _
            "      (SELECT DECODE(Y.出生时间,NULL,X.入院时间,Y.出生时间) AS 入院时间,X.病人ID,X.主页ID FROM (SELECT 病人ID,主页ID,MIN(开始时间) AS 入院时间" & vbNewLine & _
            "      FROM 病人变动记录" & vbNewLine & _
            "      WHERE 开始时间 IS NOT NULL AND 病人ID =[2] AND 主页ID =[3] GROUP BY 病人ID,主页ID) X," & vbNewLine & _
            "      (SELECT 病人ID,主页ID,出生时间 FROM 病人新生儿记录 WHERE 病人ID =[2] AND 主页ID =[3] AND 序号=[4]) Y" & vbNewLine & _
            "      WHERE X.病人ID=Y.病人ID(+) AND X.主页ID=Y.主页ID(+) ) B,部门表 C ," & vbNewLine & _
            "      (SELECT NVL(发生时间,SYSDATE) 发生时间 FROM (SELECT MAX(发生时间) 发生时间 FROM 病人护理文件 A,病人护理数据 B" & vbNewLine & _
            "      WHERE A.ID=B.文件ID AND A.ID=[1] AND A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=[4])) D," & vbNewLine & _
            strNewSql & "," & vbNewLine & _
            "      (SELECT  COUNT(*) LAST FROM" & vbNewLine & _
            "      (SELECT 病区ID FROM 病人变动记录" & vbNewLine & _
            "                WHERE 开始时间 IS NOT NULL AND 病人ID =[2] AND 主页ID = [3] GROUP BY 病区ID )) F" & vbNewLine & _
            "WHERE B.病人ID=E.病人ID And B.主页ID=E.主页ID And C.ID(+)=A.病区ID" & vbNewLine & _
            "ORDER BY A.开始时间"
            
    Set RS = zldatabase.OpenSQLRecord(strSql, "usrBodyEditor", lng文件ID, lng病人ID, lng主页ID, int婴儿)
    
    For lngLoop = 0 To rsTmp("页数").Value - 1

        strDateFrom = Format(rsTmp("入院时间").Value + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("入院时间").Value + 7 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"

        If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
        End If

        If strDateFrom < Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then

            If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss")

            RS.Filter = ""
            RS.Filter = "开始页码<=" & lngLoop + 1 & " And 结束页码>=" & lngLoop + 1

            If RS.RecordCount > 0 Then RS.MoveFirst

            For intCOl = 1 To RS.RecordCount

                If strDateFrom < Format(RS("开始时间").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strTmp = Format(RS("开始时间").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strTmp = strDateFrom
                End If

                If strDateTo > Format(RS("终止时间").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strCaption = Format(RS("终止时间").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strCaption = strDateTo
                End If

                strCaption = Format(strTmp, "yyyy-MM-dd") & "～" & Format(strCaption, "yyyy-MM-dd")
                strCaption = "第" & lngLoop + 1 & "页：" & strCaption & "(" & RS("名称").Value & ")"

                '入院时间;科室id;开始时间;结束时间;
                Set mcbrItem = mcbrToolBar页面.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
                mcbrItem.Parameter = strEnterDate & ";" & RS!病区ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop & ";" & strOutDate
                
                If lngLoop + 1 <= 4 Then
                    Set cbrWeek = mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop)))
                    cbrWeek.Parameter = strEnterDate & ";" & RS!病区ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop & ";" & strOutDate
                End If
                 
                lnglast科室id = Val(Nvl(RS("病区ID").Value))

                RS.MoveNext

                strParameter = mcbrItem.Parameter
                
                '指定页号不为0 并且和该页数相等就记录参数值
                If T_Patient.lngPage <> 0 And Val(T_Patient.lngPage - 1) = lngLoop Then
                    strParam1 = strParameter
                    strSvrCaption1 = strCaption
                End If
                
                strSvrCaption = strCaption
            Next
        Else
            mintAllPage = lngLoop
            Exit For
        End If
    Next
    
    '设置入院后固定前四周的状态
    For lngLoop = 0 To 3
        If mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop))).Parameter = "" Then mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop))).Enabled = False
    Next lngLoop
    
    '页号不为空就按指定页号显示
    If strParam1 <> "" Then strParameter = strParam1: strSvrCaption = strSvrCaption1
    
    '设置上一页下一页状态
    Call InitWeekDays(strParameter)

    If strParameter <> "" Then
        mstrParam = strParameter
        mcbrToolBar页面.Caption = strSvrCaption
        Call zlMenuClick("装载数据", mstrParam)
    End If
    
    cbsMain.RecalcLayout
    
    InitBody = True

    Exit Function

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub InitWeekDays(ByVal strParameter As String)

    '-----------------------------------------------------------------------------------
    '功能:设置有多少个上一页和下一页
    '参数:strParameter '入院时间;科室id;开始时间;结束时间;页数;终止时间
    '------------------------------------------------------------------------------------
    Dim ArrCode      As Variant

    Dim strBeginTime As String, strEndTime As String, strDateFrom As String

    Dim cbrMunu      As CommandBarButton
    
    Dim lngLoop As Long
    
    Dim lngPage As Long

    ArrCode = Split(strParameter, ";")
    
    
    On Error GoTo Errhand
    
    If CDate(Format(CStr(ArrCode(0)), "yyyy-MM-dd HH:mm:ss")) > CDate(Format(CStr(ArrCode(5)), "yyyy-MM-dd HH:mm:ss")) Then ArrCode(0) = Format(ArrCode(5), "yyyy-MM-dd HH:mm:ss")
    
    
    For lngLoop = 0 To Round((DateDiff("D", CDate(ArrCode(0)), CDate(ArrCode(5))) + 1) / 7)

        strDateFrom = Format(CDate(ArrCode(0)) + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"

        If strDateFrom < Format(ArrCode(0), "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(ArrCode(0), "yyyy-MM-dd HH:mm:ss")
        End If

        If strDateFrom < Format(ArrCode(5), "yyyy-MM-dd HH:mm:ss") Then
            lngPage = lngLoop
        End If
    Next lngLoop

    With mcbrToolBar.Controls
        
        '设置上下月参数.
        Set cbrMunu = .Find(, conMenu_View_Forward) '上一页
        cbrMunu.Parameter = ArrCode(4)  '存放还有几个上一页
        
        Set cbrMunu = .Find(, conMenu_View_Backward) '下一月
        cbrMunu.Parameter = Val(lngPage - Val(ArrCode(4)))  '还有几个下一月

    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
    Dim RS As New ADODB.Recordset
    Dim varParam   As Variant
    Dim blnRefresh As Boolean
    Dim strStartDate As String
    Dim strEndDate   As String
    Dim intCOl  As Long
    Dim strTime As String, strInput As String
    
    On Error GoTo Errhand
    If Trim(strParam) <> "" Then varParam = Split(strParam, ";")
    
    Select Case strMenuItem
        Case "初始化"
            mstrParam1 = "": mstrParam2 = ""
            mstrParam1 = strParam
            mbln出院 = False
            picMain.Tag = ""
            'strParam 病人ID;主页ID;病区ID;文件ID;出院;编辑;婴儿;护理等级;是否根据具窗体大不自动校正体温单格式(1 否 0 是)
            T_Patient.lng病人ID = varParam(0)
            T_Patient.lng主页ID = varParam(1)
            T_Patient.lng病区ID = varParam(2)
            T_Patient.lng科室ID = varParam(2)
            T_Patient.lng文件ID = varParam(3)
            T_Patient.lngPage = 0
            
            If UBound(varParam) > 3 Then
                T_Patient.lng出院 = varParam(4)
            Else
                T_Patient.lng出院 = 0
            End If
            
            If UBound(varParam) > 4 Then
                T_Patient.lng编辑 = varParam(5)
            Else
                T_Patient.lng编辑 = 0
            End If
            
            If UBound(varParam) > 5 Then
                T_Patient.lng婴儿 = varParam(6)
            Else
                T_Patient.lng婴儿 = 0
            End If
            
            If UBound(varParam) > 6 Then
                T_Patient.lng护理等级 = varParam(7)
            Else
                T_Patient.lng护理等级 = 3
            End If
            
            If UBound(varParam) > 7 Then
                T_Patient.lng原始大小 = Val(varParam(8))
            Else
                T_Patient.lng原始大小 = 0
            End If
            
            If UBound(varParam) > 8 Then
                T_Patient.lngPage = Val(varParam(9))
            End If
            
            mblnAutoAdjust = IIf(T_Patient.lng原始大小 = 1, False, True)
            mblnRedraw = False
            mblnRefresh = True
            
            mstrSQL = "Select 出院科室ID from 病案主页 Where 病人id=[1] And 主页id=[2] "
            Set RS = zldatabase.OpenSQLRecord(mstrSQL, "提取科室ID", T_Patient.lng病人ID, T_Patient.lng主页ID)
            If RS.BOF = False Then
                T_Patient.lng科室ID = Val(zlCommFun.Nvl(RS("出院科室ID").Value))
            End If

            mstrSQL = "SELECT A.序号,A.姓名 FROM(" & vbNewLine & _
                        "SELECT A.序号,A.姓名,A.病人ID,A.主页ID FROM (SELECT 0 序号, B.姓名,A.病人ID,A.主页ID" & vbNewLine & _
                        "            FROM 病案主页 A, 病人信息 B" & vbNewLine & _
                        "            WHERE A.病人ID = B.病人ID AND A.病人ID =[1] AND A.主页ID =[2]" & vbNewLine & _
                        "            UNION ALL" & vbNewLine & _
                        "            SELECT A.序号, DECODE(A.婴儿姓名, NULL, B.姓名 || '之子' || TRIM(TO_CHAR(A.序号, '9')), A.婴儿姓名) AS 姓名,A.病人ID,A.主页ID" & vbNewLine & _
                        "            FROM 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
                        "            WHERE A.病人ID =[1] AND A.主页ID =[2] AND A.病人ID = B.病人ID) A," & vbNewLine & _
                        "            (SELECT A.病人ID,A.主页ID , NVL(A.婴儿,0) 婴儿 FROM 病人护理文件 A,病历文件列表 B" & vbNewLine & _
                        "            WHERE A.格式ID=B.ID AND B.种类=3 AND B.保留=-1) B" & vbNewLine & _
                        "            WHERE A.病人ID=B.病人ID AND A.主页ID=B.主页ID AND A.序号=B.婴儿) A" & vbNewLine & _
                        "ORDER BY A.序号"
            Set RS = zldatabase.OpenSQLRecord(mstrSQL, mstrTitle, T_Patient.lng病人ID, T_Patient.lng主页ID)
            
            cboBaby.Clear
            If RS.BOF = False Then
                Do While Not RS.EOF
                    cboBaby.AddItem RS("姓名").Value
                    cboBaby.ItemData(cboBaby.NewIndex) = RS("序号").Value
                    RS.MoveNext
                    If cboBaby.ListIndex = -1 And T_Patient.lng婴儿 = Val(cboBaby.ItemData(cboBaby.NewIndex)) Then
                        Call zlControl.CboSetIndex(cboBaby.hWnd, cboBaby.NewIndex)
                        T_Patient.lng婴儿 = cboBaby.ItemData(cboBaby.ListIndex)
                    End If
                Loop
            End If
            
            If cboBaby.ListCount > 0 And cboBaby.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cboBaby.hWnd, 0)
                T_Patient.lng婴儿 = cboBaby.ItemData(cboBaby.ListIndex)
            End If
            
            '初始画布
            Call Paint_Init(picDraw, picBuffer)
            
            If Not InitData(T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng出院, T_Patient.lng编辑, T_Patient.lng婴儿) Then Exit Function
            If Not InitBody(T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿) Then Exit Function
            Call ReSetFontSize
        Case "装载数据"
            'strParam格式：起始时间;科室ID;开始时间;结束时间;页号
            
            'Debug.Print Now & ":装载数据"
            mstrParam2 = strParam
            mblnRedraw = True
            mbln呼吸曲线 = True
            mstrEnterDate = Format(varParam(0), "YYYY-MM-DD HH:mm:ss")
            strStartDate = Format(varParam(2), "YYYY-MM-DD HH:mm:ss")
            strEndDate = Format(varParam(3), "YYYY-MM-DD HH:mm:ss")
            mintPage = Val(varParam(4))
            glngCurPage = mintPage + 1
            mstrEndDate = Format(varParam(5), "YYYY-MM-DD HH:mm:ss")
            If mbln出院 = True Then
                '出院时间和入院时间如果在同一列，则将出院时间后移一列（内蒙需求:出院也要录入体温）
                mstrEndDate = Format(RetrunEndTime(CDate(mstrEnterDate), CDate(mstrEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
                strEndDate = Format(RetrunEndTime(CDate(mstrEnterDate), CDate(strEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
            End If
            If strStartDate & ";" & strEndDate = picMain.Tag Then
                mblnRefresh = False
            Else
                mblnRefresh = True
            End If
            
            picMain.Tag = strStartDate & ";" & strEndDate
                        
            mstr开始时间 = strStartDate
            mstr结束时间 = strEndDate
            
            If mstr开始时间 = "" Or mstr结束时间 = "" Then
                Call FaceInitTable(False)
                Call picDraw_Paint '从内存中Copy画布到PIC
            Else
                If mblnRefresh = True Then
                    Call ReadBodyInfo '加载病人基本信息
                    'Debug.Print Now & ":初始化上下表格"
                    Call FaceInitTable '初始化上下表格
                    'Debug.Print Now & ":加载病人体温数据"
                    Call ReadBoyData(mblnAutoAdjust) '加载病人体温数据
                    'Debug.Print Now & ":开始输出图形"
                    Call Paint_Construct   '输出曲线和图形
                    Call Paint_Assistant '输出上标,下标,未记说明,出院，专科，等信息
                    Call picDraw_Paint '从内存中Copy画布到PIC
                    'Debug.Print Now & ":加载表格数据"
                    Call ShowDowntab '加载下表格数据
                End If
            End If
            
            '销毁创建的字体信息
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            mlngOldFont = 0: mlngFont = 0
            
            mlngWidth = UserControl.Width
            mlngHeight = UserControl.Height
            
            'Debug.Print Now & ":装载数据Over"
        Case "显示病人信息"
            If T_Patient.lng婴儿 = 0 Then
                txtCard(0).Text = txtCard(0).Tag
                txtCard(7).Text = txtCard(7).Tag
            Else
                txtCard(5).Text = ""
                txtCard(6).Text = ""
                txtCard(7).Text = ""
                
                mstrSQL = "Select Decode(a.婴儿姓名,Null,b.姓名||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名,a.婴儿性别,a.出生时间 From 病人新生儿记录 a,病人信息 b Where a.病人id=[1] And a.主页id=[2] And a.病人id=b.病人id And a.序号=[3]"
                Set RS = zldatabase.OpenSQLRecord(mstrSQL, "提取婴儿信息", T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿)
                If RS.BOF = False Then
                    txtCard(0).Text = RS("婴儿姓名").Value
                    txtCard(5).Text = RS("婴儿性别").Value
                    txtCard(6).Text = "新生儿"
                End If
            End If
            
        Case "体温数据显示设置"
            If T_Patient.lng编辑 = 0 Then Exit Function
            If mstr开始时间 <> "" Then
                '计算选择的列
                intCOl = (picDisplay.Left - mshUpTab.ColWidth(0) + mshUpTab.ColWidth(1)) / mshUpTab.ColWidth(1)
                intCOl = intCOl - 5
                If intCOl < mintColMin Then intCOl = mintColMin
                
                '计算得到列返回的时间范围
                If Trim(strParam) <> "" Then '在体温编辑界面调用显示是传入时间(因为保存数据体温单刷新后,会定位到第一天)
                    strTime = Format(varParam(0), "YYYY-MM-DD HH:mm:ss")
                Else
                    strTime = Split(GetCurveDate(intCOl, mstr开始时间, gintHourBegin), ";")(0)
                End If
                
                If Format(strTime, "YYYY-MM-DD HH:mm:ss") < Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss") Then
                    strTime = Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")
                End If
                
                strInput = T_Patient.lng病人ID & ";" & T_Patient.lng主页ID & ";" & T_Patient.lng文件ID & ";" & T_Patient.lng婴儿 & ";" & T_Patient.lng科室ID & ";" & T_Patient.lng护理等级
                If frmCaseTendBodySetShowData.ShowEdit(UserControl.Extender.ParentForm, strInput, CDate(strTime), CDate(Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss")), mint心率应用, mblnMoved, FontSize) = True Then
                    '保存成功后刷新体温单显示
                    strParam = mstrParam2
                    picMain.Tag = ""
                    Call zlMenuClick("装载数据", strParam)
                End If
            End If
            
        Case "体温数据编辑"
            If T_Patient.lng编辑 = 0 Then Exit Function
            Dim strCurDate As String, strDay As String
            If mstr开始时间 <> "" Then
                If picMain.Tag = "" Then picMain.Tag = mstr开始时间 & ";" & mstr结束时间
                
               strCurDate = zldatabase.Currentdate
               
    
                
'               intCOl = (lblCur.Left - mshUpTab.ColWidth(0) - mshUpTab.Left - ((mshUpTab.ColWidth(1) - lblCur.Width) / 2)) / mshUpTab.ColWidth(1) + 1
                'intCOl = mshUpTab.Col
                '计算得到列返回的时间范围
                If Trim(strParam) <> "" Then
                    strTime = Format(varParam(0), "YYYY-MM-DD HH:mm:ss") & ";" & Format(varParam(1), "YYYY-MM-DD HH:mm:ss")
                Else
                    '计算选择的列
                    intCOl = (picDisplay.Left - mshUpTab.ColWidth(0) + mshUpTab.ColWidth(1)) / mshUpTab.ColWidth(1)
                    intCOl = intCOl - 5
                    If intCOl < mintColMin Then intCOl = mintColMin
                    strTime = GetCurveDate(intCOl, mstr开始时间, gintHourBegin)
                End If
                
                If Format(Split(strTime, ";")(0), "YYYY-MM-DD HH:mm:ss") < Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss") Then
                    strTime = Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss") & ";" & Format(Split(strTime, ";")(1), "YYYY-MM-DD HH:mm:ss")
                ElseIf Format(Split(strTime, ";")(1), "YYYY-MM-DD HH:mm:ss") > Format(mstr结束时间, "YYYY-MM-DD HH:mm:ss") Then
                    strTime = Format(Split(strTime, ";")(0), "YYYY-MM-DD HH:mm:ss") & ";" & Format(mstr结束时间, "YYYY-MM-DD HH:mm:ss")
                End If
                
                
                
                strInput = T_Patient.lng病人ID & ";" & T_Patient.lng主页ID & ";" & T_Patient.lng文件ID & ";" & T_Patient.lng婴儿 & ";" & T_Patient.lng科室ID & ";" & T_Patient.lng护理等级
                If frmCaseTendBodySetData.ShowEditor(UserControl.Extender.ParentForm, strInput, strTime, mstr开始时间, mint心率应用, mblnMoved, FontSize) = True Then
                    '保存成功后刷新体温单显示
                    mstrParam1 = mstrParam1 & String(9 - UBound(Split(mstrParam1, ";")), ";")
                    varParam = Split(mstrParam1, ";")
                    varParam(3) = T_Patient.lng文件ID
                    varParam(6) = T_Patient.lng婴儿
                    varParam(9) = mintPage + 1
                    mstrParam1 = Join(varParam, ";")
                    strParam = mstrParam1
                    
                    Call zlMenuClick("初始化", strParam)
                End If
            End If
    End Select
    
    zlMenuClick = True
    Exit Function

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function InitData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng出院 As Long, ByVal lng编辑 As Long, ByVal int婴儿 As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    '读取病人数据
    T_Patient.lng病人ID = lng病人ID
    T_Patient.lng主页ID = lng主页ID
    T_Patient.lng出院 = lng出院
    T_Patient.lng编辑 = lng编辑

    '加载初始化参数,设置曲线时间段
    Call InitPara

    '进行必要的检查
    '获取病人当前护理等级
    T_Patient.lng护理等级 = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "获取病人当前护理等级", T_Patient.lng病人ID, T_Patient.lng主页ID)
    If rsTemp.BOF = False Then T_Patient.lng护理等级 = zlCommFun.Nvl(rsTemp("护理等级"), 3)

    '检查是否有曲线体温项目
    gstrSQL = " Select 1 From 体温记录项目 A,诊治所见项目 B,护理记录项目 C " & _
              " Where C.项目序号=A.项目序号 " & _
                        "AND C.项目ID=B.ID(+) " & _
                        "AND C.护理等级>=[1] " & _
                        "And A.记录法=1 And RowNum<2 And C.项目序号<>" & gint心率
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否存在曲线项目", T_Patient.lng护理等级)
    If rsTemp.EOF Then
        MsgBox "至少要有一个曲线项目！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '判断该病人是否已经转出
    If T_Patient.lng病人ID > 0 And T_Patient.lng出院 = 1 Then
        gstrSQL = "select nvl(数据转出,0) 转出 from 病案主页 where 病人ID=[1] and 主页ID=[2]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查病人是否转出", T_Patient.lng病人ID, T_Patient.lng主页ID)
        mblnMoved = (Val(rsTemp("转出")) <> 0)
    End If
    
    vsf.Body.Appearance = flexFlat
    vsf.Body.RowHidden(0) = True
    vsf.Body.ColHidden(0) = True
    vsf.Body.ScrollBars = flexScrollBarNone
    vsf.Body.BorderStyle = flexBorderNone
    vsf.Body.OwnerDraw = flexODOver
    vsf.FixedCols = 1
    vsf.FixedRows = 1
    vsf.Rows = 2
    vsf.Body.RowHeight(vsf.FixedRows) = 400
    vsf.Height = vsf.Body.RowHeight(vsf.FixedRows)
     
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReadBodyInfo()
'功能:获取病人信息
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, strTime As String, strStart As String, strTo As String
    Dim intCOl As Integer
    Dim bln入科显示入院 As Boolean
    On Error GoTo hErr
    
    strStart = mstr开始时间
    strTo = mstr结束时间
    
    If zldatabase.GetPara("体温单显示诊断", glngSys, 1255, 1) = 0 Then
        lblCard(7).Visible = False
        txtCard(7).Visible = False
    Else
        lblCard(7).Visible = True
        txtCard(7).Visible = True
    End If
    
    If CStr(mstrEndDate) < Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And mbln出院 = False Then
        mstrEndDate = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If mintAllPage = mintPage + 1 Then
        If CStr(mstr结束时间) < Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And mbln出院 = False Then
            mstr结束时间 = Format(zldatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    txtCard(3).Text = ""
    
    '如果是新生儿，则重新计算时间，即婴儿体温单的开始时间
    If T_Patient.lng婴儿 > 0 Then
        mstrSQL = " Select  b.出生时间 From 病人新生儿记录 B Where 病人id=[1] And 主页id=[2] And 序号=[3] "
        Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "提取新生儿信息", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID), T_Patient.lng婴儿)
        If rsTmp.BOF = False Then
            mstrEnterDate = Format(zlCommFun.Nvl(rsTmp("出生时间").Value), "yyyy-MM-dd HH:mm:ss")
            txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("出生时间").Value), "yyyy-MM-dd")
            strStart = mstrEnterDate
        End If
    End If
    
    '此处进行时间转换
    intCOl = GetCurveColumn(CDate(strStart), CDate(strStart), gintHourBegin) + mshUpTab.FixedCols - 1
    strStart = Split(GetCurveDate(intCOl - mshUpTab.FixedCols + 1, CDate(strStart), gintHourBegin), ";")(0)
    
    If CDate(strStart) < CDate(mstr开始时间) Then
        strStart = Format(mstr开始时间, "yyyy-MM-dd HH:mm:ss")
    End If
    
    intCOl = GetCurveColumn(CDate(strTo), CDate(strStart), gintHourBegin) + mshUpTab.FixedCols - 1
    strTo = Split(GetCurveDate(intCOl - mshUpTab.FixedCols + 1, CDate(strStart), gintHourBegin), ";")(1)
    If CDate(Format(strTo, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss")) Then
        strTo = Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss")
    End If
    
    mstr开始时间 = Format(strStart, "yyyy-MM-dd HH:mm:ss")
    mstr结束时间 = Format(strTo, "yyyy-MM-dd HH:mm:ss")
    
    picMain.Tag = mstr开始时间 & ";" & mstr结束时间
    
    bln入科显示入院 = False
    If CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrComeInDate, "yyyy-MM-dd HH:mm:ss")) Then
        bln入科显示入院 = True
    ElseIf CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) = CDate(Format(mstrComeInDate, "yyyy-MM-dd HH:mm:ss")) And T_BodyFlag.入院 = 0 Then
        bln入科显示入院 = True
    End If
    
    '入院时间(以入科时间为准)
    mstrSQL = "select 开始时间 from 病人变动记录 where 病人id=[1] And 主页id=[2] and 开始原因=2 order by 开始时间"
    Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "变动记录", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID))
    If rsTmp.BOF = False Then
        If txtCard(3).Text = "" And bln入科显示入院 = True Then txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("开始时间").Value), "yyyy-MM-dd")
    End If
    
    '读取病人基本信息
    mstrSQL = " Select  b.姓名,A.住院号,A.入院日期 入院时间,b.性别,A.年龄 From 病人信息 B,病案主页 A Where A.病人ID=B.病人ID And A.病人id=[1] And A.主页ID=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "提取病人信息", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID))
    If rsTmp.BOF = False Then
        txtCard(0).Text = zlCommFun.Nvl(rsTmp("姓名").Value)
        txtCard(0).Tag = zlCommFun.Nvl(rsTmp("姓名").Value)
        txtCard(1).Text = zlCommFun.Nvl(rsTmp("住院号").Value)
        txtCard(5).Text = zlCommFun.Nvl(rsTmp("性别").Value)
        txtCard(6).Text = zlCommFun.Nvl(rsTmp("年龄").Value)
        If txtCard(3).Text = "" Then txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("入院时间").Value), "yyyy-MM-dd")
    End If
    
    
    '读取病人科室、床号等信息
    
    txtCard(2).Text = ""
    txtCard(4).Text = ""
    
    mstrSQL = " Select  c.名称 As 科室,b.名称 As 病区,a.床号,a.开始原因 " & _
                "From 病人变动记录 a,部门表 b,部门表 c " & _
                "Where a.病人id=[1] And a.主页id=[2] And a.科室id Is Not Null And a.病区id=b.id and a.科室id=c.id And a.开始时间-4/24<=[3] And Nvl(a.终止时间,Sysdate)>=[4] Order By a.开始时间"
    
    Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "读取病人科室、床号等信息", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID), CDate(mstr结束时间), CDate(mstr开始时间))
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            
            If zlCommFun.Nvl(rsTmp("科室").Value) <> strTmp And zlCommFun.Nvl(rsTmp("科室").Value) <> "" Then
            
                strTmp = zlCommFun.Nvl(rsTmp("科室").Value)
                
                If txtCard(2).Text = "" Then
                    txtCard(2).Text = strTmp
                Else
                    txtCard(2).Text = txtCard(2).Text & "->" & strTmp
                End If
                
            End If

            If zlCommFun.Nvl(rsTmp("床号").Value) <> strTime And zlCommFun.Nvl(rsTmp("床号").Value) <> "" Then
                strTime = zlCommFun.Nvl(rsTmp("床号").Value)
                
                If txtCard(4).Text = "" Then
                    txtCard(4).Text = strTime
                Else
                    txtCard(4).Text = txtCard(4).Text & "->" & strTime
                End If
                
            End If
                        
            rsTmp.MoveNext
        Loop
        
        If Left(txtCard(2).Text, 2) = "->" Then txtCard(2).Text = Mid(txtCard(2).Text, 3)
        If Left(txtCard(4).Text, 2) = "->" Then txtCard(4).Text = Mid(txtCard(4).Text, 3)
    End If
    
    '提取病人诊断信息
    mstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,0,[4]) As 最后诊断 From Dual"
    Set rsTmp = zldatabase.OpenSQLRecord(mstrSQL, "最后诊断", "最后诊断", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID), CDate(strStart))
    If rsTmp.BOF = False Then
        If T_Patient.lng婴儿 = 0 Then
            txtCard(7).Text = zlCommFun.Nvl(rsTmp("最后诊断").Value)
        Else
            txtCard(7).Text = ""
        End If
    Else
        txtCard(7).Text = ""
    End If
    txtCard(7).Tag = txtCard(7).Text
    
    Call zlMenuClick("显示病人信息")

    Exit Sub
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FaceInitTable(Optional ByVal blnInitUpdate As Boolean = True)
'---------------------------------------------------------------------
'功能：加载显示表上表下数据
'---------------------------------------------------------------------
    '加载设置上下表格
    Dim rsTemp As New ADODB.Recordset
    Dim intCOl As Integer, intRow As Integer
    Dim lngCount As Long
    Dim lngWith As Long
    Dim strPace As String
    
    On Error GoTo Errhand
    
    If T_DrawClient.列单位 = 0 Then T_DrawClient.列单位 = glngColStep
    T_DrawClient.刻度区域.Left = T_DrawClient.偏移量X
    '得到曲线总数
    lngCount = CurveCount
    
    '结算体温单刻度区域的左右边距
    If lngCount <= 3 Then
        T_DrawClient.刻度区域.Right = T_DrawClient.刻度区域.Left + glngLableWith
    Else
        T_DrawClient.刻度区域.Right = T_DrawClient.刻度区域.Left + lngCount * glngLableStep
    End If
    
    lngWith = T_DrawClient.列单位 * Screen.TwipsPerPixelX
    
    With mshUpTab
        .Cols = 43
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        .Cell(flexcpText, 0, .FixedCols, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 0, .FixedCols, .Rows - 1, .Cols - 1) = ""
    
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeRow(2) = True
        .ColWidth(0) = (T_DrawClient.刻度区域.Right - T_DrawClient.刻度区域.Left) * Screen.TwipsPerPixelX
        .TextMatrix(0, 0) = "日       期"
        .TextMatrix(1, 0) = IIf(T_Patient.lng婴儿 = 0, "住 院 天 数", "出 生 天 数")
        .TextMatrix(2, 0) = "手术后天数"
        .TextMatrix(3, 0) = "时       间"
        
        '.Cell(flexcpWidth, 0, 1, .Rows - 1, .Cols - 1) = lngWith
        For intCOl = 1 To .Cols - 1
            .ColWidth(intCOl) = lngWith
        Next
        .ColWidthMin = lngWith
        .Redraw = flexRDBuffered
    End With
    
    '合并单元格的列
    For intRow = 0 To 2
        Call UniteCellCol(mshUpTab, 6, intRow, mshUpTab.FixedCols)
    Next intRow
    
    If blnInitUpdate = True Then Call ShowUptab
    
    With vsf
        .Cols = 0
        .NewColumn "", 0, 1
        .NewColumn "项目", mshUpTab.ColWidth(0) + 10, 1
    
        For intCOl = 1 To 42
            .NewColumn intCOl, lngWith + 7, 1, , 1
        Next
        
        .Left = T_DrawClient.偏移量X * Screen.TwipsPerPixelX
        .FixedCols = 2
        .Rows = 2
        .Body.Appearance = flexFlat
        .Body.RowHidden(0) = True
        .Body.ColHidden(0) = True
        .Body.ScrollBars = flexScrollBarNone
        .Body.BorderStyle = flexBorderNone
        .Body.OwnerDraw = flexODOver
        .Cell(flexcpAlignment, 1, 1) = flexAlignCenterCenter
        .Cell(flexcpFontName, 1, 2, 1, .Cols - 1) = "Times New Roman"
        .Cell(flexcpFontSize, 1, 2, 1, .Cols - 1) = 7.5
        .Cell(flexcpForeColor, 1, 2, 1, .Cols - 1) = RGB_RED
        .Body.Select 1, 1
        .Body.CellBorder 0, 1, 0, 0, 0, 0, 0
        .Body.Select 1, vsf.Cols - 1
        .Body.CellBorder 0, 0, 0, 1, 0, 0, 0
        .Body.BackColorFixed = .Body.BackColor
        .Visible = False
        For intCOl = 3 To .Cols - 1 Step 2
            .Cell(flexcpBackColor, 1, intCOl, 1, intCOl) = &HF7ECE6
        Next
        For intCOl = 1 To .Cols - 1
            .EditMode(intCOl) = 0
        Next
        .Height = .Body.RowHeight(.FixedRows)
    End With
    
    '加载下表格(护理项目)
    With mshDownTab
        .Cols = 46
        .Rows = 1
        .ColWidth(0) = mshUpTab.ColWidth(0)
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeRow(0) = True
        .Tag = 0
        
        For intCOl = .FixedCols To .Cols - 1
            .ColWidth(intCOl) = mshUpTab.ColWidth(1)
            If (intCOl - .FixedCols + 1) Mod 2 = 0 Then
                .Cell(flexcpBackColor, 0, intCOl, .Rows - 1, intCOl) = &H80000013
            End If
        Next intCOl

        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
    
    mItemNO.呼吸 = 0
    mintRepairRows = zldatabase.GetPara("体温表格行数", glngSys, 1255, 8)
    mbln显示皮试 = (Val(zldatabase.GetPara("体温单显示皮试结果", glngSys, 1255, "0")) = 1)
    
    '检查呼吸是否是表格项目
    gstrSQL = "select 记录法 From 体温记录项目 where 项目序号=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "体温记录项目", gint呼吸)
    If rsTemp.RecordCount > 0 Then
         mintRepairRows = mintRepairRows - IIf(Val(Nvl(rsTemp!记录法)) = 2, 1, 0)
    End If
    If mintRepairRows < 0 Then mintRepairRows = 0

    '加载所有表格项目，包括固定项目和有数据的活动项目
    Set rsTemp = GetAppendGridItem(T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng护理等级, T_Patient.lng婴儿, Int(CDate(mstr开始时间)), CDate(mstr结束时间), IIf(T_Patient.lng婴儿 = 0, 1, 2), T_Patient.lng科室ID, mblnMoved)
    With rsTemp
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            mshDownTab.Rows = 0
            Call AppenGridItem(rsTemp)
        Else
            mshDownTab.Rows = 0
        End If
    End With
    
    mshDownTab.Rows = mintRepairRows
    
    '补充完剩下的空行
    If mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
        For intRow = Val(mshDownTab.Tag) To mshDownTab.Rows - 1

            mshDownTab.MergeRow(intRow) = True
            For intCOl = 0 To mshDownTab.FixedCols
                strPace = " " & String(intCOl, " ") & String(intRow, " ")
                mshDownTab.TextMatrix(intRow, intCOl) = strPace & "" & strPace
            Next intCOl
            
            Call UniteCellCol(mshDownTab, 6, intRow, mshDownTab.FixedCols)
        Next intRow
    End If
    
    If mbln显示皮试 And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
        intRow = Val(mshDownTab.Tag)
        strPace = " " & String(1, " ") & String(intRow, " ")
        mshDownTab.TextMatrix(intRow, 0) = strPace & "皮试结果" & strPace
    End If
    
    '重新整理表格位置
    If mItemNO.呼吸 <> 0 Then
        mbln呼吸曲线 = False
    End If
    
    '设置表格颜色
    For intCOl = mshDownTab.FixedCols To mshDownTab.Cols - 1
        If (intCOl - mshDownTab.FixedCols + 1) Mod 2 = 0 Then
            mshDownTab.Cell(flexcpBackColor, 0, intCOl, mshDownTab.Rows - 1, intCOl) = &HF7ECE6
        End If
    Next intCOl
    mshDownTab.Cell(flexcpAlignment, 0, 0, mshDownTab.Rows - 1, mshDownTab.Cols - 1) = 4
    
    Call picBack_Resize
    
    Call Paint_Canvas(mblnAutoAdjust) '初始化体温数据
    
    Call picBack_Resize
    
    Call SetVisible
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function SetVisible() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------------------
    If T_Patient.lng编辑 = 0 Then
        mshUpTab.Enabled = False
        mshDownTab.Enabled = False
    Else
        mshUpTab.Enabled = True
        mshDownTab.Enabled = True
    End If
End Function


Private Function ShowUptab() As Boolean
'----------------------------------------------------------------
'功能:输出表上日期信息 包括入院日期，住院天数，手术标注
'----------------------------------------------------------------
    Dim lngValue  As Long, intCOl As Long
    Dim lngDays   As Long
    Dim i As Long, j As Long
    Dim lngColor  As Long
    Dim intMinCol As Long, intMaxCol As Long
    Dim strTmp As String
    Dim arrOperDay, strTmp1 As String
    Dim rsTmp  As New ADODB.Recordset
    Dim str时间 As String
    Dim intDays As Integer
    Dim lng次数 As Long
    Dim lngWith As Long

    On Error GoTo Errhand

    With mshUpTab
        
        lngValue = 0
        gstrSQL = "Select zl_CalcInDaysNew([1],[2],[3],[4]) As 开始天数 From Dual"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "提取住院天数", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, Int(CDate(mstr开始时间)))

        If rsTmp.BOF = False Then
            lngValue = rsTmp("开始天数").Value
        End If
        
        '上表格式有单元格合并的，此处需要进行处理
        For intCOl = 1 To 7

            .ColData(intCOl) = 0
            .Row = 0
            .Col = intCOl
            .ColAlignment(intCOl) = 4

            strTmp = Format(CDate(mstr开始时间) + intCOl - 1, "yyyy-MM-dd")

            lngDays = lngValue + (intCOl - 1)
            
            For i = 1 To 6
                .Row = 0
                .Col = (intCOl - 1) * 6 + i
                
                If Right(strTmp, 5) = "01-01" Then
                    '一年的第一天
                    .Text = strTmp
                ElseIf strTmp = Format(mstrEnterDate, "yyyy-MM-dd") Then
                    '入院第一天，写上年份
                    .Text = strTmp
                ElseIf intCOl = 1 Then
                    .Text = Right(strTmp, 5)
                ElseIf Right(strTmp, 2) = "01" Then
                    .Text = Right(strTmp, 5)
                Else
                    .Text = Right(strTmp, 2)
                End If

                .Row = 1
                .Text = lngDays
            Next i
        Next
        
        '输出上表时间点信息
        If picMain.Tag <> "" Then
            Call CalcMinMaxCol(picMain.Tag, intMinCol, intMaxCol)
            mintColMin = intMinCol
            mintColMax = intMaxCol
            
            With picDisplay
                .Left = ((((intMaxCol - 1) \ 6) + 1) * 6 - 1) * mshUpTab.ColWidth(intMinCol) + mshUpTab.ColWidth(0)
                mshUpTab.Row = mshUpTab.FixedRows
                .Top = (mshUpTab.RowHeight(mshUpTab.FixedRows) - .Height) / 2
                .Enabled = IIf(T_Patient.lng编辑 = 1, True, False)
            End With
            
            lblCur.Left = (intMinCol - 1) * .ColWidth(intMinCol) + .ColWidth(0)
            '居中显示
            lblCur.Left = lblCur.Left + (.ColWidth(intMinCol) - lblCur.Width) / 2
            lblCur.Top = .Height - lblCur.Height
            lblCur.Enabled = IIf(T_Patient.lng编辑 = 1, True, False)
        End If
        '改为DrawCell输出 由于列宽太小时间字体显示不完整
'        For i = 1 To 7
'            '输出上午下午时间
'            .Row = 3
'            For j = 1 To 6
'
'                Select Case j
'
'                    Case 1
'                        strTmp = gintHourBegin + 4 * 0
'                        lngColor = &H8080FF
'
'                    Case 2
'                        strTmp = gintHourBegin + 4 * 1
'                        lngColor = &H8080FF
'
'                    Case 3
'                        strTmp = gintHourBegin + 4 * 2
'                        lngColor = &H80000012
'
'                    Case 4
'                        lngColor = &H80000012
'                        strTmp = gintHourBegin + 4 * 3
'
'                    Case 5
'                        lngColor = &H80000012
'                        strTmp = gintHourBegin + 4 * 4
'
'                    Case 6
'                        lngColor = &H8080FF
'                        strTmp = gintHourBegin + 4 * 5
'                End Select
'
'                .Col = j + (i - 1) * 6
'                .ColAlignment(.Col) = 4
'
'                If .Col >= intMinCol And .Col <= intMaxCol Then
'                    lngColor = lngColor
'                Else
'                    lngColor = RGB_FleetGRAY
'                End If
'
'                .CellForeColor = lngColor
'
'                If picMain.Tag <> "" Then
'                    .Text = strTmp
'                End If
'
'            Next j
'        Next i
        
        For i = 1 To 7
            mstrOpValue(i) = .TextMatrix(2, ((i - 1) * 6 + 1))
            mstrOpdays(i) = .TextMatrix(2, ((i - 1) * 6 + 1))
        Next i
        
        '提取输入标志天数和停止手术标志
        mintOpDays = Val(zldatabase.GetPara("手术后标注天数", glngSys, 1255, "10"))
        mblnStopFlag = (Val(zldatabase.GetPara("再次手术停止前次标注", glngSys, 1255, "0")) = 1)
        '51338,刘鹏飞,2012-07-06
        strTmp = zldatabase.GetPara("手术当天缺省格式", glngSys, 1255, "2")
        If Val(strTmp) >= 0 And Val(strTmp) <= 2 Then
            mintOpFormat = Val(strTmp)
        Else
            mintOpFormat = 0
        End If
        
        strTmp = ""
        '显示但前段的手术标记
        gstrSQL = "select B.发生时间 时间" & _
            "   From 病人护理文件 A,病人护理数据 B,病人护理明细 C" & _
            "   where A.ID=B.文件ID And  B.ID=C.记录ID And A.ID=[1] And nvl(A.婴儿,0)=[4]" & _
            "   and A.病人ID=[2] and A.主页ID=[3] and C.记录类型=4 and C.终止版本 is null" & _
            "   and B.发生时间 between [5] and [6] order by B.发生时间"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "提取手术标记", Val(T_Patient.lng文件ID), T_Patient.lng病人ID, T_Patient.lng主页ID, Val(T_Patient.lng婴儿), Int(CDate(mstr开始时间) - 14), CDate(mstr结束时间))
        
        If mblnMoved Then
            gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
            gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
            gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
        End If
        
        Do While Not rsTmp.EOF
            str时间 = Format(rsTmp("时间"), "YYYY-MM-DD")
            For i = 1 To 7
                If DateDiff("d", mstr开始时间, mstr结束时间) + 1 >= i Then
                    intDays = DateDiff("d", str时间, mstr开始时间) + (i - 1)

                    Select Case intDays

                        Case 0 '当前区域内的手术开始时间
                            'Modify 2012-03-05 修改一天可以有多次手术
                            If Trim(mstrOpdays(i)) <> "" Then
                                mstrOpdays(i) = str时间 & "/" & mstrOpdays(i)
                            Else
                                mstrOpdays(i) = str时间
                            End If
                            
                        Case 1 To mintOpDays '手术开始天数

                            If mblnStopFlag Then '手术标注后天数在次手术时停止前一次标注
                                mstrOpValue(i) = intDays
                            Else
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = intDays & "/" & mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = intDays
                                End If
                            End If
                    End Select
                End If
            Next i
            rsTmp.MoveNext
        Loop
        
        
        '提取当前开始日期-14天前的手术记录信息
        gstrSQL = "Select Nvl(Count(B.发生时间),0) 次数" & _
            "   From 病人护理文件 A, 病人护理数据 B,病人护理明细 C" & _
            "   Where A.ID=B.文件ID And B.ID=C.记录ID and A.ID=[1] and nvl(A.婴儿,0)=[4]" & _
            "   And A.病人ID=[2] And A.主页ID=[3] And C.记录类型=4 and C.终止版本 is null" & _
            "   And B.发生时间 <[5] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "提取手术标记", Val(T_Patient.lng文件ID), T_Patient.lng病人ID, T_Patient.lng主页ID, Val(T_Patient.lng婴儿), Int(CDate(mstr开始时间)))
        
        If mblnMoved Then
            gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
            gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
            gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
        End If
        lng次数 = 0
        If rsTmp.BOF = False Then lng次数 = Val(rsTmp("次数"))
        For i = 1 To 7
            If DateDiff("d", mstr开始时间, mstr结束时间) + 1 >= i Then
                '修改一天可能存在多次手术
                If Trim(mstrOpdays(i)) <> "" Then
                    arrOperDay = Split(mstrOpdays(i), "/")
                Else
                    arrOperDay = Split("1", "/")
                End If
                lngValue = lng次数
                If Trim(mstrOpdays(i)) <> "" And lngValue + UBound(arrOperDay) < 12 Then
                    strTmp = "": strTmp1 = ""
                    For j = UBound(arrOperDay) + 1 To 1 Step -1
                        lng次数 = lngValue + j
                        strTmp1 = Switch(lng次数 = 1, "Ⅰ", lng次数 = 2, "Ⅱ", lng次数 = 3, "Ⅲ", lng次数 = 4, "Ⅳ", lng次数 = 5, "Ⅴ", lng次数 = 6, "Ⅵ", lng次数 = 7, "Ⅶ", lng次数 = 8, "Ⅷ", lng次数 = 9, "Ⅸ", lng次数 = 10, "Ⅹ", lng次数 = 11, "Ⅺ", lng次数 = 12, "Ⅻ")
                        If strTmp = "" Then
                            strTmp = strTmp1
                        Else
                            strTmp = strTmp & "/" & strTmp1
                        End If
                        If mblnStopFlag Then Exit For
                    Next j
                    lng次数 = lngValue + UBound(arrOperDay) + 1
                    If mblnStopFlag Then '手术标注后天数在次手术时停止前一次标注
                        Select Case mintOpFormat
                            Case 1 '--显示0
                                mstrOpValue(i) = .TextMatrix(2, ((i - 1) * 6 + 1)) & "0" & .TextMatrix(2, ((i - 1) * 6 + 1))
                            Case 2 '--显示次数
                                If strTmp = "Ⅰ" Then
                                    mstrOpValue(i) = 0
                                Else
                                    mstrOpValue(i) = strTmp & "-0"
                                End If
                            Case Else '--不显示
                                 mstrOpValue(i) = .TextMatrix(2, ((i - 1) * 6 + 1))
                        End Select
                    Else
                        Select Case mintOpFormat
                            Case 1 '--显示0
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = 0 & "/" & mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = 0
                                End If
                            Case 2 '--显示次数
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = strTmp & "/" & mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = strTmp
                                End If
                            Case Else  '--不显示
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = .TextMatrix(2, ((i - 1) * 6 + 1))
                                End If
                        End Select
                    End If
                    .Row = 2
                    For j = 1 To 6
                        .Col = j + (i - 1) * 6
                        .Text = mstrOpValue(i)
                    Next j
                Else
                    .Row = 2
                    For j = 1 To 6
                        .Col = j + (i - 1) * 6
                        .Text = mstrOpValue(i)
                    Next j
                End If
            End If
        Next i
        '设定日期，住院天数文本颜色
        mshUpTab.Cell(flexcpForeColor, 0, mshUpTab.FixedCols, 1, mshUpTab.Cols - 1) = 16711680
        '设定手术 分娩文本颜色
        '51283,刘鹏飞,2012-07-11
        lngColor = Val(zldatabase.GetPara("手术天数显示颜色", glngSys, 1255, "255"))
        mshUpTab.Cell(flexcpForeColor, 2, mshUpTab.FixedCols, 2, mshUpTab.Cols - 1) = lngColor

        lngWith = T_DrawClient.列单位 * Screen.TwipsPerPixelX
        'mshUpTab.Cell(flexcpWidth, 0, 1, mshUpTab.Rows - 1, mshUpTab.Cols - 1) = lngWith
        For intCOl = 1 To mshUpTab.Cols - 1
            mshUpTab.ColWidth(intCOl) = lngWith
        Next intCOl
        mshUpTab.ColWidthMin = lngWith
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
    End With

    ShowUptab = True
    Exit Function
    
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowDowntab() As Boolean
    '输出表下数据（护理项目信息）
    Dim rsTemp   As New ADODB.Recordset
    Dim rsDownTab As New ADODB.Recordset
    Dim intRow As Integer, intRow1 As Integer
    Dim intCOl As Integer, intCol1 As Integer
    Dim intColCount As Integer, intRowCount As Integer
    Dim intDay As Integer
    Dim strItems As String, strItemName As String, strSql As String
    Dim lngItemCode As Long
    Dim strPace As String
    Dim str项目名称 As String, str项目名称1 As String
    Dim int记录频次 As Integer, int项目性质 As Integer, int项目类型 As Integer, int项目表示 As Integer, int入院首测 As Integer
    Dim strBegin As String, str结果 As String, strPart As String
    Dim int舒张压 As Integer, int收缩压 As Integer, Int列号 As Integer
    Dim blnColor As Boolean
    Dim lngColor As Long
    Dim arrTmpString0(1 To 42) As String, arrTmpString1(1 To 42) As String, arrTmpString2(1 To 42) As String
    Dim blnAdd As Boolean, blnValue As Boolean
    Dim SinX As Single
    Dim i As Integer
    Dim int呼吸位置 As Integer, intValue As Integer, int呼吸表格输出格式 As Integer
    Dim bln汇总当天 As Boolean, bln录入小时 As Boolean
    Dim arrTmp() As String
    Dim dtBegin As Date, dtEnd As Date
    
    On Error GoTo Errhand
    
    Call InitPublicData '提取基础数据
    
    ReDim mstrNewString(mintRepairRows, 6)
    'mstrNewString = Split(String(6, ";"), ";")
    int呼吸表格输出格式 = zldatabase.GetPara("呼吸表格输出", glngSys, 1255, 0)
    bln汇总当天 = (Val(zldatabase.GetPara("汇总波动显示当天数据", glngSys, 1255, 0)) = 1)
    mbln灌肠大便分子分母显示 = (Val(zldatabase.GetPara("灌肠后大便显示格式", glngSys, 1255, 0)) = 1)
    '--51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H)
    bln录入小时 = (Val(zldatabase.GetPara("全天汇总显示录入时间", glngSys, 1255, 0)) = 1)
    
    gbln出院 = mbln出院
    dtBegin = Int(CDate(mstr开始时间) - 1)
    dtEnd = CDate(CDate(mstr结束时间) + 1)
    
    If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) Then _
        dtBegin = CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss"))
    If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss")) Then _
        dtEnd = CDate(Format(mstrEndDate, "YYYY-MM-DD HH:mm:ss"))
    
    '提取项目名称(拼接字符串)
    strItems = ""
    For intRow = mshDownTab.FixedRows To Val(mshDownTab.Tag) - 1
        If Val(mshDownTab.RowData(intRow)) <> mItemNO.血压 Then
            i = InStr(1, mshDownTab.TextMatrix(intRow, 0), "(")
            If i > 0 Then
                strItemName = Trim(Left(mshDownTab.TextMatrix(intRow, 0), i - 1))
            Else
                strItemName = Trim(mshDownTab.TextMatrix(intRow, 0))
            End If
            If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                strItems = strItems & ",'" & strItemName & "'"
            End If
        End If
    Next
    
    If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
    If Not mbln呼吸曲线 Then strItems = strItems & ",'呼吸'"
    strItems = strItems & ",'收缩压','舒张压'"
    If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
    'Debug.Print "读取数据开始---" & Now
    '提取病人体温表格记录
    gstrSQL = " SELECT C.ID,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & _
        "   DECODE(E.项目性质,2,C.体温部位 || D.记录名 ,D.记录名) 项目名称,D.项目序号,C.来源ID,C.共用,E.项目性质 " & _
        "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E " & _
        "   Where B.ID=A.文件ID And A.ID = C.记录ID   AND B.ID=[1]  AND Nvl(B.婴儿,0)=[7] " & _
        "   AND B.病人id=[2]  AND B.主页id=[3] AND INSTR([6],decode(E.项目性质,2,C.体温部位 || D.记录名 ,D.记录名))>0 " & _
        "   AND D.项目序号=C.项目序号  AND MOD(c.记录类型,10)=1  AND E.项目序号=D.项目序号 " & _
        "   AND nvl(E.护理等级,0)>=[8]  AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null AND D.记录法=2 "
    
    '提取非体温表格的汇总项目
    strSql = "  SELECT C.ID,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & _
        "   D.项目名称,D.项目序号,C.来源ID,C.共用,D.项目性质" & _
        "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,(SELECT A.项目序号,A.项目名称, 1 项目性质,B.父序号 FROM 护理记录项目 A,护理汇总项目 B" & vbNewLine & _
        "       WHERE A.项目序号=B.序号 AND NOT EXISTS (SELECT C.项目序号 FROM 体温记录项目 C,护理汇总项目 E WHERE C.项目序号=E.序号 AND C.项目序号=A.项目序号)" & vbNewLine & _
        "       AND NVL(A.应用方式,0)=1 AND NVL(A.护理等级,0)>=[8] AND NVL(A.适用病人,0) IN (0,[9])" & vbNewLine & _
        "       AND (A.适用科室=1 OR (A.适用科室=2 AND EXISTS (SELECT 1 FROM 护理适用科室 D WHERE D.项目序号=A.项目序号 AND D.科室ID=[10])))) D" & _
        "   Where B.ID=A.文件ID And A.ID = C.记录ID   AND B.ID=[1]  AND Nvl(B.婴儿,0)=[7] " & _
        "   AND B.病人id=[2]  AND B.主页id=[3]  AND D.项目序号=C.项目序号  AND C.记录类型=1" & _
        "   AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null"

    gstrSQL = "Select ID,时间,记录类型,显示,结果,体温部位,未记说明,数据来源,项目名称,项目序号,来源ID,共用,项目性质 From (" & _
        "   " & gstrSQL & " UNION ALL " & strSql & ")" & _
        "   Order By  Decode(项目名称,'收缩压',0,1)," & strItems & ",时间"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取体温表格数据", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, _
                        CDate(dtBegin), CDate(dtEnd), strItems, T_Patient.lng婴儿, T_Patient.lng护理等级, IIf(T_Patient.lng婴儿 = 0, 1, 2), T_Patient.lng科室ID)
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
        gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
        gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
    End If
    
    'Debug.Print "读取数据结束---" & Now
    '1---输出呼吸表格数据
    vsf.Cell(flexcpText, 1, 2, 1, vsf.Cols - 1) = ""
    vsf.Cell(flexcpData, 1, 2, 1, vsf.Cols - 1) = ""
    vsf.Cell(flexcpForeColor, 1, 2, 1, vsf.Cols - 1) = 200
    
    rsTemp.Filter = "项目序号=" & gint呼吸
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    With rsTemp
        Do While Not .EOF
            If CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")) Then
                blnAdd = False
                intCOl = GetCurveColumn(rsTemp!时间, mstr开始时间, gintHourBegin) + vsf.FixedCols - 1
                str结果 = zlCommFun.Nvl(rsTemp!结果) & ";" & Nvl(rsTemp!体温部位)
                If intCOl < vsf.Cols Then
                    If arrTmpString1(intCOl - vsf.FixedCols + 1) <> "" Then
                        If (Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) <> 1 And Val(zlCommFun.Nvl(!显示, 0)) <> 1) Or _
                            (Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) = 1 And Val(zlCommFun.Nvl(!显示, 0)) = 1) Then
                            
                            '检查那个离重点时间更近
                            SinX = GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"))
                            blnAdd = GetCanvasCenter(CDate(Format(arrTmpString1(intCOl - vsf.FixedCols + 1), "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) = 1 Then
                            blnAdd = False
                        Else
                            blnAdd = True
                        End If
                        
                        If blnAdd = True Then
                            If Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) = 2 Then
                                arrTmpString0(intCOl - vsf.FixedCols + 1) = str结果
                                arrTmpString1(intCOl - vsf.FixedCols + 1) = Format(rsTemp!时间, "YYYY-MM-DD HH:mm:ss")
                                arrTmpString2(intCOl - vsf.FixedCols + 1) = 2
                                GoTo ErrNext
                            End If
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                arrTmpString2(intCOl - vsf.FixedCols + 1) = 2
                                GoTo ErrNext
                            End If
                        End If
                    Else
                        blnAdd = True
                    End If
                    
                    If blnAdd = True Then
                        arrTmpString0(intCOl - vsf.FixedCols + 1) = str结果
                        arrTmpString1(intCOl - vsf.FixedCols + 1) = Format(rsTemp!时间, "YYYY-MM-DD HH:mm:ss")
                        arrTmpString2(intCOl - vsf.FixedCols + 1) = Val(zlCommFun.Nvl(!显示, 0))
                    End If
                End If
            End If
ErrNext:
        .MoveNext
        Loop
    End With
    
    '一次循环呼吸数组,如果检查到显示=2则清空数据
    For i = 1 To 42
        If Val(arrTmpString2(i)) = 2 Then arrTmpString0(i) = ""
    Next i
    
    '2----开始输出呼吸数据 呼吸机为图形输出
    int呼吸位置 = 0
    blnValue = False
     '循环输出呼吸值
    vsf.Cell(flexcpForeColor, 1, vsf.FixedCols, 1, vsf.Cols - 1) = Val(vsf.Tag)
    For i = 1 To 42
        intCOl = i + vsf.FixedCols - 1
        If InStr(1, arrTmpString0(i), ";") > 0 Then
            str结果 = Split(arrTmpString0(i), ";")(0)
            strPart = Split(arrTmpString0(i), ";")(1)
        Else
            str结果 = arrTmpString0(i)
            strPart = ""
        End If
        
        '打印呼吸值（间隔错开打印） 第一行始终在上面
        If IsNumeric(str结果) Then
            vsf.TextMatrix(1, intCOl) = str结果
            If blnValue = False Then
                intValue = IIf(intCOl Mod 2 = 0, 0, 1)
                blnValue = True
                int呼吸位置 = 2
            End If
            
            If int呼吸表格输出格式 = 0 Then '顺序上下显示
                If intCOl Mod 2 = intValue Then
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = flexAlignCenterTop
                    If strPart <> "呼吸机" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = 1
                    End If
                Else
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = flexAlignCenterBottom
                    If strPart <> "呼吸机" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = 2
                    End If
                End If
                
            Else        '有数据时数据之间上下显示
                If int呼吸位置 = 2 Then
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = flexAlignCenterTop
                    If strPart <> "呼吸机" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = 1
                    End If
                Else
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = flexAlignCenterBottom
                    If strPart <> "呼吸机" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = 2
                    End If
                End If
                
                int呼吸位置 = int呼吸位置 + 1
                If int呼吸位置 > 2 Then int呼吸位置 = 1
            End If
        End If
    Next i
       
    'Debug.Print "数据开始---" & Now
    '提取表格项目数据信息
    With mshDownTab
        lngItemCode = 0
        str项目名称 = ""
        For intRow = .FixedRows To .Tag - 1
            i = InStr(1, .TextMatrix(intRow, 0), "(")

            If i > 0 Then
                str项目名称1 = Trim(Mid(.TextMatrix(intRow, 0), 1, i - 1))
            Else
                str项目名称1 = Trim(.TextMatrix(intRow, 0))
            End If
            
            blnColor = False
            If str项目名称1 & ";" & .RowData(intRow) <> str项目名称 & ";" & lngItemCode Then
                
                lngItemCode = .RowData(intRow)
                str项目名称 = str项目名称1
                int项目类型 = Val(Split(.TextMatrix(intRow, 1), ",")(0))
                int记录频次 = Val(Split(.TextMatrix(intRow, 1), ",")(2))
                int项目表示 = Val(Split(.TextMatrix(intRow, 1), ",")(3))
                int项目性质 = Val(Split(.TextMatrix(intRow, 1), ",")(4))
                int入院首测 = Val(Split(.TextMatrix(intRow, 1), ",")(6))
                blnColor = (int项目性质 = 2 And int项目类型 = 1 And int项目表示 = 0)
                
                For intDay = 0 To 6
                    strBegin = DateAdd("D", intDay, CDate(mstr开始时间))
                    If CDate(strBegin) > CDate(mstr结束时间) Then strBegin = mstr结束时间
                    int舒张压 = 0
                    int收缩压 = 0
                    Int列号 = 0
                    '循环得到某个项目某天的数据信息
                    Set rsDownTab = ReturnItemRecord(rsTemp, Int(CDate(strBegin)), CDate(mstrEnterDate), lngItemCode & ";" & str项目名称 & ";" & _
                                int记录频次 & ";" & int项目表示 & ";" & int项目性质 & ";" & int入院首测, bln汇总当天, bln录入小时)
                    If rsDownTab.RecordCount > 0 Then rsDownTab.MoveFirst
                    rsDownTab.Sort = "时间,项目序号,序号"
                    Do While Not rsDownTab.EOF
                        str结果 = zlCommFun.Nvl(rsDownTab!记录内容, "")
                        lngColor = 0
                        If blnColor Then lngColor = Val(zlCommFun.Nvl(rsDownTab!未记说明, 0))
                        intCOl = Val(rsDownTab!序号)
                        intColCount = 0
                        intRow1 = 0
                        strPace = ""
                        
                        Select Case int记录频次
                            Case 1
                                intRow1 = intRow
                                intCOl = intDay * 6 + .FixedCols
                                intColCount = 6
                                strPace = " "
                            Case 2
                                intRow1 = intRow
                                intCOl = (intCOl - 1) * 3 + intDay * 6 + .FixedCols
                                intColCount = 3
                                strPace = String(intCOl, " ")
                            Case 3
                                intRow1 = intRow + (intCOl - 1)
                                intCOl = intDay * 6 + .FixedCols
                                intColCount = 6
                                strPace = " "
                            Case 4
                                intRow1 = intRow + Fix((intCOl - 1) / 2)
                                Select Case intCOl
                                    Case 1, 3
                                        intCOl = 1
                                    Case 2, 4
                                        intCOl = 2
                                End Select
                                intCOl = (intCOl - 1) * 3 + intDay * 6 + .FixedCols
                                intColCount = 3
                                strPace = String(intCOl, " ")
                            Case 6
                                intRow1 = intRow
                                intCOl = (intCOl - 1) + intDay * 6 + .FixedCols
                                intColCount = 1
                                strPace = String(intCOl, " ")
                        End Select
                        
                        '检查本次输出的列是否在输出行数之内
                        If mintRepairRows > 0 And mintRepairRows - 1 >= intRow1 Then
                            strPace = strPace & String(intDay + 1, " ") & String(intRow1, " ")
                            '将数据展示在表格中
                            Select Case rsDownTab!项目序号
                                Case mItemNO.舒张压
                                    If int舒张压 <> Val(rsDownTab!序号) Then
                                        For i = 1 To intColCount
                                            intCol1 = intCOl + (i - 1)
                                            If intCol1 < mshDownTab.Cols Then
                                                If Trim(mshDownTab.TextMatrix(intRow1, intCol1)) <> "" Or str结果 <> "" Then
                                                    If InStr(1, mshDownTab.TextMatrix(intRow1, intCol1), "/") > 0 Then
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & Trim(Split(mshDownTab.TextMatrix(intRow1, intCol1), "/")(0)) & "/" & str结果 & strPace
                                                    Else
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & "/" & str结果 & strPace
                                                    End If
                                                    '--问题号：53505，修改人：李涛，血压显示文字。
                                                    If str结果 = "外出" Or str结果 = "拒测" Or str结果 = "请假" Or str结果 = "未测" Then
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & str结果 & strPace
                                                    End If
                                                End If
                                            End If
                                        Next i
                                        int舒张压 = Val(rsDownTab!序号)
                                    End If
                                Case mItemNO.血压 '收缩压
                                    If int收缩压 <> Val(rsDownTab!序号) Then
                                        For i = 1 To intColCount
                                            intCol1 = intCOl + (i - 1)
                                            If intCol1 < mshDownTab.Cols Then
                                                If Trim(mshDownTab.TextMatrix(intRow1, intCol1)) <> "" Or str结果 <> "" Then
                                                    If InStr(1, mshDownTab.TextMatrix(intRow1, intCol1), "/") > 0 Then
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & str结果 & "/" & Trim(Split(mshDownTab.TextMatrix(intRow1, intCol1), "/")(1)) & strPace
                                                    Else
                                                        mshDownTab.TextMatrix(intRow1, intCol1) = strPace & str结果 & "/" & strPace
                                                    End If
                                                End If
                                            End If
                                        Next i
                                        int收缩压 = Val(rsDownTab!序号)
                                    End If
                                Case Else
                                    If Int列号 <> Val(rsDownTab!序号) Then
                                        For i = 1 To intColCount
                                            intCol1 = intCOl + (i - 1)
                                            If intCol1 < mshDownTab.Cols Then
                                                mshDownTab.TextMatrix(intRow1, intCol1) = strPace & str结果 & strPace
                                                If int项目性质 = 2 And int项目类型 = 1 And int项目表示 = 0 Then
                                                    mshDownTab.Cell(flexcpForeColor, intRow1, intCol1, intRow1, intCol1) = lngColor
                                                End If
                                            End If
                                        Next i
                                        Int列号 = Val(rsDownTab!序号)
                                    End If
                            End Select
                        End If
                    rsDownTab.MoveNext
                    Loop
                    If Format(strBegin, "YYYY-MM-DD") = Format(mstr结束时间, "YYYY-MM-DD") Then
                        Exit For
                    End If
                Next intDay
            End If
        Next intRow
        
        '开始输出皮试结果
        If mbln显示皮试 = True And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
            strSql = _
               "SELECT 时间,F_LIST2STR(CAST(COLLECT(药物名) AS T_STRLIST)) 药物名 FROM (" & vbNewLine & _
                "   SELECT TO_CHAR(开始执行时间,'YYYY-MM-DD') 时间,DECODE(皮试结果,'(+)',255,0) || '-#' || REPLACE(REPLACE(医嘱内容,',',''),'-#','') || 皮试结果  药物名" & vbNewLine & _
                "   FROM 病人医嘱记录" & vbNewLine & _
                "   WHERE  病人ID=[1] AND 主页ID=[2] AND 婴儿=[3] AND 皮试结果 IS NOT NULL" & vbNewLine & _
                "   AND 开始执行时间  BETWEEN [4] AND [5]" & vbNewLine & _
                "   ORDER BY TO_DATE(TO_CHAR(开始执行时间,'YYYY-MM-DD'),'YYYY-MM-DD'),皮试结果" & vbNewLine & _
                ") GROUP BY 时间"

            If mblnMoved Then
                strSql = Replace(strSql, "病人过敏记录", "H病人过敏记录")
            End If

            Set rsDownTab = zldatabase.OpenSQLRecord(strSql, "提取病人过敏记录信息", T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿, CDate(mstr开始时间), CDate(mstr结束时间))

            Do While Not rsDownTab.EOF
                intCOl = DateDiff("D", CDate(Format(mstr开始时间, "YYYY-MM-DD")), CDate(Format(rsDownTab!时间, "YYYY-MM-DD")))
                str结果 = Nvl(rsDownTab!药物名)
                Call ShowTestis(str结果, intCOl)
                rsDownTab.MoveNext
            Loop
        End If
    End With
    'Debug.Print "数据结束---" & Now
    
    
    ShowDowntab = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowTestis(ByVal strValue As String, ByVal intCOl As Integer)
'----------------------------------------------------------------------
'功能:结算皮试结果要输出的行数
'----------------------------------------------------------------------
    Dim intNum As Integer, i As Integer
    Dim lngColor As Long
    Dim strTmp As String, strPart As String, strPic As String
    Dim arrTmp() As String
    Dim LPoint As T_LPoint
    Dim lngDC As Long
    Dim objDraw As Object
    Dim lngH As Long, lngW As Long, lngX1 As Long, lngLen As Long
    Dim intRowCount As Integer
    Dim sngLen As Single
    Dim intRow As Integer
    
    Set objDraw = picBack
    intRowCount = Val(mshDownTab.Tag)
    intNum = 1
    strTmp = strValue
    If strTmp = "" Then Exit Sub
    LPoint.X = 0
    LPoint.W = mshDownTab.ColWidth(mshDownTab.FixedCols) / Screen.TwipsPerPixelX * 6
    lngW = LPoint.W
    lngX1 = 0
    
    '开始计算是否需要换行
    strPart = ""
    arrTmp = Split(strTmp, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        lngColor = Val(Split(arrTmp(i), "-#")(0))
        strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") '皮试结果
        If Trim(strTmp) <> "" Then
            Do While True
                T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                strPic = strTmp
                If T_Size.W - (LPoint.W - (LPoint.X - lngX1)) > 0 Then
                    sngLen = Round((LPoint.W - (LPoint.X - lngX1)) / T_Size.W, 2)
                    lngLen = Len(StrConv(strTmp, vbFromUnicode)) * sngLen
                    '将半角转为全角
                    strTmp = StrConv(strTmp, vbWide)
                    strPart = StrConv(Mid(StrConv(strTmp, vbFromUnicode), lngLen + 1), vbUnicode)
                    strTmp = StrConv(Mid(StrConv(strTmp, vbFromUnicode), 1, lngLen), vbUnicode)
                    '截取原始字符串
                    strPart = Mid(strPic, Len(strTmp) + 1)
                    strTmp = Mid(strPic, 1, Len(strTmp))
                    
                    mstrNewString(intRow, intCOl) = mstrNewString(intRow, intCOl) & "," & lngColor & "-#" & strTmp
                    If Left(mstrNewString(intRow, intCOl), 1) = "," Then mstrNewString(intRow, intCOl) = Mid(mstrNewString(intRow, intCOl), 2)
                    
                    T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                    LPoint.X = LPoint.X + T_Size.W
                    strTmp = strPart
                    T_Size.W = objDraw.TextWidth("字") / T_TwipsPerPixel.X
                    If T_Size.W - (LPoint.W - (LPoint.X - lngX1)) > 0 Then
                        LPoint.X = lngX1
                        intRow = intRow + 1
                        intNum = intNum + 1

                        If intRowCount + intNum > mintRepairRows Then Exit Sub
                    End If
                    If strTmp = "" Then Exit Do
                Else
                    mstrNewString(intRow, intCOl) = mstrNewString(intRow, intCOl) & "," & lngColor & "-#" & strTmp
                    If Left(mstrNewString(intRow, intCOl), 1) = "," Then mstrNewString(intRow, intCOl) = Mid(mstrNewString(intRow, intCOl), 2)
                    If T_Size.W + objDraw.TextWidth("字") / T_TwipsPerPixel.X - LPoint.W > 0 Then
                        LPoint.X = lngX1
                    Else
                        LPoint.X = LPoint.X + T_Size.W
                    End If

                    Exit Do
                End If
            Loop
        End If
    Next i
End Sub

Public Sub AppenGridItem(ByVal rsTemp As ADODB.Recordset)
    '填写表格标题
    Dim intRow  As Integer, intRowStart As Integer
    Dim int频次 As Integer
    Dim intRowNum As Integer, intColNum As Integer
    Dim intRowCount As Integer, intNum As Integer
    Dim i As Integer, j As Integer
    Dim strText As String, str值域 As String

    On Error GoTo Errhand
    
    With rsTemp
        j = 0
        Do While Not .EOF
            intRowCount = mshDownTab.Rows
            Select Case !记录名
                Case "呼吸"
                    mItemNO.呼吸 = !项目序号
                    vsf.TextMatrix(1, 1) = Nvl(!记录名, "呼吸") & IIf(Not IsNull(!单位), "(" & !单位 & ")", "")
                    vsf.Tag = Val(Nvl(!记录色, RGB_RED))
                Case "舒张压"
                    mItemNO.舒张压 = !项目序号
                Case Else
                    If mintRepairRows > 0 And mintRepairRows > intRowCount Then
                        j = j + 1
                        int频次 = zlCommFun.Nvl(!记录频次, 2)
                        
                        '汇总项目或波动项目频次最大为2
                        If Val(zlCommFun.Nvl(!项目表示)) = 4 Or IsWaveItem(Val(zlCommFun.Nvl(!项目序号))) Then
                            If int频次 > 2 Then int频次 = 2
                        End If
                        
                        Select Case int频次
                            'intColNum 要合并的列数
                            'intRowNum 要合并的行
                            Case 1
                                intRowNum = 1
                                intColNum = 6
                            Case 2
                                intRowNum = 1
                                intColNum = 3
                            Case 3
                                intRowNum = 3
                                intColNum = 6
                            Case 4
                                intRowNum = 2
                                intColNum = 3
                            Case 6
                                intRowNum = 1
                                intColNum = 1
                        End Select
                        
                        '计算要添加的列数
                        If mshDownTab.Rows + intRowNum > mintRepairRows Then
                            intNum = mintRepairRows - mshDownTab.Rows
                        Else
                            intNum = intRowNum
                        End If
                        
                        intRowNum = intNum
                        mshDownTab.Rows = mshDownTab.Rows + intRowNum
                        mshDownTab.Tag = mshDownTab.Rows '记录实际输出的表格行数
                        intRowStart = mshDownTab.Rows - intRowNum
                        
                        '合并列并赋值
                        For i = 1 To intRowNum
                            intRow = intRowStart + i - 1
                            
                            mshDownTab.MergeCol(0) = True
                            mshDownTab.MergeRow(intRow) = True
                            
                            If !记录名 = "收缩压" Then
                                mshDownTab.TextMatrix(intRow, 0) = String(j, "　") & "血压" & IIf(Not IsNull(!单位), "(" & !单位 & ")", "") & String(j, "　")
                                mItemNO.血压 = !项目序号
                                mItemRow.血压 = intRowStart
                            Else
                                mshDownTab.TextMatrix(intRow, 0) = String(j, "　") & Replace(Nvl(!记录名), ";", ":") & IIf(Not IsNull(!单位), "(" & !单位 & ")", "") & String(j, "　")
                            End If
                            
                            strText = !项目序号
                            mshDownTab.RowData(intRow) = strText
                            mshDownTab.RowHeight(intRow) = 255
                            
                            mshDownTab.TextMatrix(intRow, 1) = zlCommFun.Nvl(!项目类型) & "," & zlCommFun.Nvl(!项目小数) & "," & _
                                int频次 & "," & zlCommFun.Nvl(!项目表示) & "," & zlCommFun.Nvl(!项目性质) & "," & zlCommFun.Nvl(!项目长度) & "," & zlCommFun.Nvl(!入院首测, 0)
                            mshDownTab.TextMatrix(intRow, 2) = zlCommFun.Nvl(!最大值, "")
                            mshDownTab.TextMatrix(intRow, 3) = zlCommFun.Nvl(!最小值, "")
                            
                            Call UniteCellCol(mshDownTab, intColNum, intRow, mshDownTab.FixedCols)
                        Next i
                    End If
            End Select
            .MoveNext
        Loop
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    On Error Resume Next
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    picBack.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    picBuffer.Move lngLeft, lngTop
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cboBaby.hWnd, KeyAscii)
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_View_Jump '菜单
            mcbrToolBar页面.Caption = Control.Caption
            mstrParam = Control.Parameter
            Call InitWeekDays(mstrParam)
            Call zlMenuClick("装载数据", mstrParam)
            cbsMain.RecalcLayout
        Case conMenu_View_OneWeek To conMenu_View_FourWeek '4个周期按钮
            mstrParam = Control.Parameter
            Call InitWeekDays(mstrParam)
            Call zlMenuClick("装载数据", mstrParam)
            mcbrToolBar页面.Caption = mcbrItem.Controls.Item(mintPage + 1).Caption
        Case conMenu_View_Forward, conMenu_Manage_CallPrevious '上一页
            Call picDraw_KeyDown(vbKeyLeft, vbCtrlMask)
        Case conMenu_View_Backward, conMenu_Manage_CallNext '下一页
            Call picDraw_KeyDown(vbKeyRight, vbCtrlMask)
    End Select

End Sub


Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.Id

        Case conMenu_View_Jump '菜单

            If Control.Parameter = "" Then
                Control.Checked = True
            Else
                Control.Checked = (Val(Split(Control.Parameter, ";")(4)) = mintPage)
            End If

        Case conMenu_View_OneWeek To conMenu_View_FourWeek '4个周期按钮

            If Control.Parameter <> "" Then
                Control.Checked = (Val(Split(Control.Parameter, ";")(4)) = mintPage)
            End If

        Case conMenu_View_Forward, conMenu_View_Backward '上下页
            Control.Enabled = IIf(Val(Control.Parameter) > 0, True, False)
    End Select

End Sub

Private Sub cmdPrimitive_Click()
'查看体温单原始比例
    Dim strParams As String
    
    strParams = ""
    strParams = T_Patient.lng病人ID & ";"
    strParams = strParams & T_Patient.lng主页ID & ";"
    strParams = strParams & T_Patient.lng病区ID & ";"
    strParams = strParams & T_Patient.lng文件ID & ";"
    strParams = strParams & T_Patient.lng出院 & ";"
    strParams = strParams & T_Patient.lng编辑 & ";"
    strParams = strParams & T_Patient.lng婴儿 & ";1;" & mintPage + 1

    RaiseEvent CmdClick(strParams)
End Sub

Private Sub hsb_Change()
    picMain.Left = -1 * hsb.Value * msinHStep
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSerach_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSerach_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lbl查看_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSerach_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl查看_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picSerach_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub mshDownTab_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim i As Integer
    Dim lngColor As Long
    Dim strTmp As String
    Dim arrTmp() As String, arrText() As String
    Dim LPoint As T_LPoint
    Dim T_ClientRect As RECT
    Dim lngBrush As Long, lngOldBrush As Long, lngBackColor As Long
    Dim lngDC As Long, lngFont As Long, lngOldFont As Long
    Dim objDraw As Object, stdset As Object
    Dim lngX1 As Long
    Dim intCOl As Integer, intRow As Integer
    
    On Error GoTo Errhand
    Err = 0
    intRow = UBound(mstrNewString)
Errhand:
    If Err <> 0 Then Exit Sub
    
    lngDC = hDC
    Set objDraw = picBack
    If mbln显示皮试 = True And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 And Col >= mshDownTab.FixedCols And Row >= Val(mshDownTab.Tag) Then
        If (Col - mshDownTab.FixedCols) Mod 6 = 0 And UBound(mstrNewString) >= (Row - Val(mshDownTab.Tag)) Then
            intCOl = (Col - mshDownTab.FixedCols) / 6
            intRow = Row - Val(mshDownTab.Tag)
            strTmp = CStr(mstrNewString(intRow, intCOl))
            If strTmp = "" Then Exit Sub
            
            '设定客户区域大小
            With T_ClientRect
                .Left = Left + 1
                .Top = Top + 1
                .Right = Right - 1
                .Bottom = Bottom - 1
            End With

            LPoint.X = Left
            Call GetTextExtentPoint32(hDC, "字", Len("字"), T_Size)
            LPoint.Y = Top + (Bottom - Top) / 2 '+ T_Size.H / 2
            lngX1 = 0
            
            '1、清空内容
            '创建与背景色相同的刷子
            lngBackColor = GetRBGFromOLEColor(mshDownTab.BackColor)
            lngBrush = CreateSolidBrush(lngBackColor)
            '使用该刷子填充背景色
            lngOldBrush = SelectObject(lngDC, lngBrush)
            Call FillRect(hDC, T_ClientRect, lngBrush)
            '立即销毁临时使用的刷子并还原刷子
            Call SelectObject(lngDC, lngOldBrush)
            Call DeleteObject(lngBrush)
        
'            '创建字体
            Set stdset = New StdFont
            stdset.Name = "宋体"
            stdset.Size = 9
            stdset.Bold = False
            Call SetFontIndirect(stdset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)

            arrTmp = Split(strTmp, ",")
            For i = LBound(arrTmp) To UBound(arrTmp)
                lngColor = Val(Split(arrTmp(i), "-#")(0))
                '设置字体颜色
                Call SetTextColor(lngDC, lngColor)
                strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") '皮试结果
                If i < UBound(arrTmp) Then strTmp = strTmp & ","
                If Trim(strTmp) <> "" Then
                    T_Size.W = objDraw.TextWidth(strTmp) / T_TwipsPerPixel.X
                    Call GetTextRect(objDraw, LPoint.X + lngX1, LPoint.Y, CStr(strTmp), , True)
                    Call DrawText(lngDC, CStr(strTmp), -1, T_LableRect, DT_CENTER)
                    lngX1 = lngX1 + T_Size.W
                End If
            Next i
           Call SelectObject(lngDC, lngOldFont)
           Call DeleteObject(lngFont)
        End If
    End If
    
    '输出大便次数
    If Col >= mshDownTab.FixedCols And Row >= mshDownTab.FixedRows Then
        strTmp = mshDownTab.TextMatrix(Row, Col)
        If AnsyGrade(Val(mshDownTab.RowData(Row)), strTmp, arrText) = True Then
            'lngColor = mshDownTab.Cell(flexcpForeColor, Row, Col, Row, Col)
            Call DrawDownTabAnsyGrade(lngDC, picMain, arrText, Row, Col, Left, Top, Right, Bottom, Done, mbln灌肠大便分子分母显示)
        End If
    End If
End Sub

Private Sub mshUpTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim strTime As String
    If NewRow = 0 And T_Patient.lng编辑 = 1 Then
        strTime = GetCurveDate(NewCol, mstr开始时间, gintHourBegin)
        If Format(Split(strTime, ";")(0), "YYYY-MM-DD") > Format(mstr结束时间, "YYYY-MM-DD") Then
            mshUpTab.FocusRect = flexFocusLight
        Else
            mshUpTab.FocusRect = flexFocusSolid
            If mblnKeyDown = True Then
                picDisplay.Left = ((((NewCol - 1) \ 6) + 1) * 6 - 1) * mshUpTab.ColWidth(NewCol) + mshUpTab.ColWidth(0)
                picDisplay.Top = (mshUpTab.RowHeight(NewRow) - picDisplay.Height) / 2
                picDisplay.Enabled = IIf(T_Patient.lng编辑 = 1, True, False)
            End If
        End If
    Else
        mshUpTab.FocusRect = flexFocusNone
    End If
    mblnKeyDown = False
End Sub

Private Sub mshUpTab_DblClick()
    If T_Patient.lng编辑 = 0 Then Exit Sub
    With mshUpTab
        If .Row = 0 And .FocusRect = flexFocusSolid Then
            RaiseEvent DbClickCur(mIntDataEditor)
        End If
    End With
End Sub

Private Sub mshUpTab_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim intMinCol As Long, intMaxCol As Long
    Dim i As Integer, j As Integer
    Dim strTmp As String
    Dim lngColor As Long, lngDC As Long
    Dim objDraw As Object, stdset As Object
    
    lngDC = hDC
    
    If picMain.Tag = "" Then Exit Sub
    If Row = mshUpTab.Rows - 1 And Col >= mshUpTab.FixedCols Then
        Set objDraw = picBack
        Call CalcMinMaxCol(picMain.Tag, intMinCol, intMaxCol)
        j = (Col - mshUpTab.FixedCols) Mod 6
        Select Case j
            Case 0
                strTmp = gintHourBegin + 4 * 0
                lngColor = &H8080FF
            Case 1
                strTmp = gintHourBegin + 4 * 1
                lngColor = &H8080FF
            Case 2
                strTmp = gintHourBegin + 4 * 2
                lngColor = &H80000012
            Case 3
                lngColor = &H80000012
                strTmp = gintHourBegin + 4 * 3
            Case 4
                lngColor = &H80000012
                strTmp = gintHourBegin + 4 * 4
            Case 5
                lngColor = &H8080FF
                strTmp = gintHourBegin + 4 * 5
        End Select
        '根据参数体温夜班时间范围决定时间颜色
        lngColor = GetTimeColor(Val(strTmp))
        If Col >= intMinCol And Col <= intMaxCol Then
            lngColor = lngColor
        Else
            lngColor = RGB_FleetGRAY
        End If
        
        Call SetTextColor(lngDC, lngColor)
        Call GetTextRect(objDraw, Left, Top + (Bottom - Top) / 2, CStr(strTmp), Right - Left - 3, True)
        Call DrawText(lngDC, CStr(strTmp), -1, T_LableRect, DT_CENTER)
    End If
End Sub

Private Sub mshUpTab_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnKeyDown = False
    With mshUpTab
        If .Row = 0 And .FocusRect = flexFocusSolid Then
            Select Case KeyCode
                Case vbKeyReturn
                    Call mshUpTab_DblClick
                Case vbKeyLeft
                    mblnKeyDown = True
                Case vbKeyRight
                    mblnKeyDown = True
            End Select
        End If
    End With
End Sub

Private Sub picBack_Resize()
    Dim lngLeft As Long
    
    On Error Resume Next
    '设定容器内各个空间的初始位置
    T_DrawClient.偏移量Y = 0
    picMain.Move 0, 0
    picMain.BackColor = &H80000005
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    
    lngLeft = T_DrawClient.偏移量X * T_TwipsPerPixel.X
    
    With vsb
        .Left = picBack.Width - .Width
        .Top = 0
        .Height = picBack.Height - hsb.Height
    End With
    
    With hsb
        .Left = 0
        .Top = picBack.Height - .Height
        .Width = picBack.Width - vsb.Width
    End With
    
    picCard(0).Move lngLeft, 10
    
    mshUpTab.Redraw = False
    mshDownTab.Redraw = False
    
    With mshUpTab
        .ColWidth(0) = (T_DrawClient.刻度区域.Right - T_DrawClient.刻度区域.Left) * 15
        .Left = lngLeft
        .Top = picCard(0).Top + picCard(0).Height
        .RowHeight(3) = 400
        .Height = (3 * mshUpTab.RowHeight(0) + 520)
        .Width = ((T_DrawClient.刻度区域.Right - T_DrawClient.刻度区域.Left) + T_DrawClient.列单位 * 6 * 7 + 1) * T_TwipsPerPixel.X
        .ColWidthMin = T_DrawClient.列单位 * Screen.TwipsPerPixelX
         picCard(0).Width = .Width
         .Refresh
    End With
    
    picDraw.Move 0, mshUpTab.Top + mshUpTab.Height, (T_DrawClient.体温区域.Right + 1) * T_TwipsPerPixel.X, _
        (T_DrawClient.刻度区域.Bottom - T_DrawClient.刻度区域.Top) * Screen.TwipsPerPixelY

    picDisplay.Height = 165
     
    With vsf
        .Top = mshUpTab.Top + mshUpTab.Height + (T_DrawClient.刻度区域.Bottom - T_DrawClient.刻度区域.Top) * Screen.TwipsPerPixelY
        .Left = lngLeft
        .Width = mshUpTab.Width
        .Height = .Body.RowHeight(vsf.FixedRows)
        .Visible = Not mbln呼吸曲线
    End With
        
    With mshDownTab
        .ColWidth(0) = mshUpTab.ColWidth(0)
        .Left = lngLeft
        .Top = mshUpTab.Top + mshUpTab.Height + (IIf(mbln呼吸曲线 = False, vsf.Height, 0)) + (T_DrawClient.刻度区域.Bottom - T_DrawClient.刻度区域.Top) * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        .Width = mshUpTab.Width
        .Height = .Rows * .RowHeight(0)
        .Refresh
    End With
    
    lblCommText.Left = lngLeft
    lblCommText.Top = mshDownTab.Top + mshDownTab.Height
    lblCommText.Visible = True
    
    mshUpTab.Redraw = True
    mshDownTab.Redraw = True
    
    picMain.Width = mshUpTab.Width + mshUpTab.Left
    picMain.Height = lblCommText.Top + lblCommText.Height
    
    '计算滚动条
    Call CalcScrollBarSize
    
    '计算体温单的可画区域大小
    mlng高度 = (picBack.Height - mshUpTab.Top - mshUpTab.Height - mshDownTab.Height - lblCommText.Height - _
        IIf(mbln呼吸曲线 = False, vsf.Height, 0) - IIf(hsb.Visible = True, hsb.Height, 0)) / Screen.TwipsPerPixelY
    
    hsb.Value = 0
    vsb.Value = 0
End Sub

Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回： 调用成功返回TRUE；否则FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    '只根据没显示出来的那部分来计算步长
    msinHStep = (picMain.Width - picBack.Width) / 100
    msinVStep = (picMain.Height - picBack.Height) / 100

    
    hsb.Max = 0 - Int(0 - ((picMain.Width - picBack.Width) / 300)) - 1
    vsb.Max = 0 - Int(0 - ((picMain.Height - picBack.Height) / 300)) - 1
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    
    With vsb
        .Height = picBack.Height - IIf(hsb.Visible = True, hsb.Height, 0)
    End With
    
    With hsb
        .Width = picBack.Width - IIf(vsf.Visible = True, vsb.Width, 0)
    End With
    
    '恒定为100,只是步长发生变化
    If hsb.Enabled Then
        hsb.Max = 100
        hsb.LargeChange = 100 / Int((Round((picMain.Width - picBack.Width) / picBack.Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange / 2
    End If
    
    If vsb.Enabled Then
        vsb.Max = 100
        vsb.LargeChange = 100 / Int((Round((picMain.Height - picBack.Height) / picBack.Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange / 2
    End If
    
    CalcScrollBarSize = True
    
End Function

 Private Sub lblCur_DblClick()
 
    If T_Patient.lng编辑 = 0 Then Exit Sub
    'RaiseEvent DbClickCur
End Sub

Private Sub mshUpTab_BeforeMouseDown(ByVal Button As Integer, _
                                     ByVal Shift As Integer, _
                                     ByVal X As Single, _
                                     ByVal Y As Single, _
                                     Cancel As Boolean)

    
    Dim strTemp   As String
    Dim intMinCol As Long
    Dim intMaxCol As Long
    Dim intCOl As Long
    If Button <> vbLeftButton Then Exit Sub
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    '计算指定的区域才可进行操作
    If T_Patient.lng编辑 = 1 Then
        intCOl = ((mintColMax - 1) \ 6 + 1) * 6
        
        If X > mshUpTab.ColWidth(0) And X < mshUpTab.ColWidth(0) + (intCOl * mshUpTab.ColWidth(intCOl)) Then
            '根据坐标，计算列数的行
            strTemp = GetXCoordinate(X / T_TwipsPerPixel.X + mshUpTab.Left / T_TwipsPerPixel.X - 1, mstr开始时间, False)
            strTemp = mstr开始时间 & ";" & Split(strTemp, ",")(1)
            '根据时间计算列
            Call CalcMinMaxCol(strTemp, intMinCol, intMaxCol)
            picDisplay.Visible = True
            If Y < mshUpTab.RowHeight(0) + 40 Then
                picDisplay.Left = ((((intMaxCol - 1) \ 6) + 1) * 6 - 1) * mshUpTab.ColWidth(intMaxCol) + mshUpTab.ColWidth(0)
                picDisplay.Top = (mshUpTab.RowHeight(mshUpTab.FixedRows) - picDisplay.Height) / 2
                picDisplay.Enabled = IIf(T_Patient.lng编辑 = 1, True, False)
                mshUpTab.Col = intMaxCol
                mshUpTab.Row = mshUpTab.FixedRows
            End If
            
            If X > mshUpTab.ColWidth(0) + ((mintColMin - 1) * mshUpTab.ColWidth(mintColMin)) And X < mshUpTab.ColWidth(0) + ((mintColMax) * mshUpTab.ColWidth(mintColMax)) Then
                If Y > 3 * mshUpTab.RowHeight(0) Then
                    lblCur.Left = (intMaxCol - 1) * mshUpTab.ColWidth(intMaxCol) + mshUpTab.ColWidth(0)
                    '居中显示
                    lblCur.Left = lblCur.Left + (mshUpTab.ColWidth(intMaxCol) - lblCur.Width) / 2
                    lblCur.Top = mshUpTab.Height - lblCur.Height
                End If
            End If
            
        End If
    End If
    
End Sub

Private Sub cboBaby_Click()
    Dim RS As New ADODB.Recordset
    
    If T_Patient.lng婴儿 = cboBaby.ItemData(cboBaby.ListIndex) Then Exit Sub
    T_Patient.lng婴儿 = cboBaby.ItemData(cboBaby.ListIndex)
    
    On Error GoTo Errhand
    '重新提取文件ID
    mstrSQL = "select A.ID from 病人护理文件 A,病历文件列表 B" & _
       "    where A.病人ID=[1] and A.主页Id=[2] and nvl(A.婴儿,0)=[3] and A.格式ID=B.ID and B.种类=3 and B.保留=-1"
    If mblnMoved = True Then
        mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
    End If
    Set RS = zldatabase.OpenSQLRecord(mstrSQL, "提取文件ID", T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿)
    
    If RS.BOF = False Then
        T_Patient.lng文件ID = Val(zlCommFun.Nvl(RS("ID")))
        cboBaby.Enabled = True
    Else
        cboBaby.Enabled = False
        T_Patient.lng婴儿 = 0
        cboBaby.ListIndex = 0
    End If
   
    If Not InitBody(T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿) Then Exit Sub
    Call zlMenuClick("显示病人姓名")
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picDraw_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyRight And Shift = vbCtrlMask Then  '下一月
        If mintPage < mcbrItem.Controls.Count - 1 Then
            mintPage = mintPage + 2
            mstrParam = mcbrItem.Controls.Item(mintPage).Parameter '得到当前页的时间
            Call InitWeekDays(mstrParam)
            mcbrToolBar页面.Caption = mcbrItem.Controls.Item(mintPage).Caption
            cbsMain.RecalcLayout
            Call zlMenuClick("装载数据", mstrParam)
        End If

    ElseIf KeyCode = vbKeyLeft And Shift = vbCtrlMask Then

        If mintPage > 0 Then '上一月
            mstrParam = mcbrItem.Controls.Item(mintPage).Parameter '得到当前页的时间
            Call InitWeekDays(mstrParam)
            mcbrToolBar页面.Caption = mcbrItem.Controls.Item(mintPage).Caption
            cbsMain.RecalcLayout
            Call zlMenuClick("装载数据", mstrParam)
        End If

    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        'mblnAutoRedraw = mblnAutoRedraw Xor True
    End If

End Sub

Private Sub picDraw_Paint()
    '----------------------------------------------------------------------------
    '功能:从内存中Copy图像到PIC上
    '----------------------------------------------------------------------------
    picDraw.Cls
    Call BitBlt(mlngDC, 0, 0, T_ClientRect.Right, T_ClientRect.Bottom, mlngMemDC, 0, 0, SRCCOPY)
End Sub


Private Sub picCard_Paint(Index As Integer)
    Dim intLoop As Integer
    Dim X1 As Single, X2 As Single, Y1 As Single, Y2 As Single
    On Error Resume Next
    
    picCard(Index).Cls
    For intLoop = 0 To txtCard.UBound
        txtCard(intLoop).Height = 180
        If txtCard(intLoop).Visible Then
            X1 = txtCard(intLoop).Left
            Y1 = txtCard(intLoop).Top + txtCard(intLoop).Height + 15
            X2 = txtCard(intLoop).Left + txtCard(intLoop).Width
            Y2 = txtCard(intLoop).Top + txtCard(intLoop).Height + 15
            picCard(Index).ForeColor = &H8000000C
            picCard(Index).DrawStyle = 0
            picCard(Index).DrawWidth = 1
            picCard(Index).Line (X2, Y2)-(X1, Y1)
        End If
    Next
End Sub

Private Sub picCard_Resize(Index As Integer)
    On Error Resume Next
    txtCard(1).Move txtCard(1).Left, txtCard(1).Top, picCard(Index).Width - txtCard(1).Left - 45
    txtCard(7).Move txtCard(7).Left, txtCard(7).Top, picCard(Index).Width - txtCard(7).Left - 45
End Sub


Private Sub picSerach_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Call RaisEffect(picSerach, -2)
End Sub

Private Sub picSerach_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    Call RaisEffect(picSerach, 2)
    Call cmdPrimitive_Click
End Sub

Private Sub txtCard_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtCard(Index))
End Sub

Private Sub UserControl_Initialize()
    picDraw.AutoRedraw = False
    Call InitCommandBar
    picBack.BackColor = &H80000005
    T_DrawClient.列单位 = glngColStep
    T_DrawClient.偏移量X = 5
    T_DrawClient.偏移量Y = 0
    
    Call RaisEffect(picSerach, 2)
End Sub

Public Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = IIf(FontSize = 0, 9, IIf(FontSize = 1, 12, FontSize))
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "宋体"

    Set CtlFont = cbsMain.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsMain.Options.Font = CtlFont

    lblSerach(9).FontSize = bytFontSize
    cboBaby.FontSize = bytFontSize
    cboBaby.Top = (picTmp.Height - cboBaby.Height) \ 2
    lblSerach(9).Top = cboBaby.Top + (cboBaby.Height - lblSerach(9).Height) \ 2
End Sub


Private Sub UserControl_Resize()
    If UserControl.Parent.Visible = False Then Exit Sub
    
    If mblnAutoAdjust = True And Not mblnResize Then
        '检查实际大小是否发生变化
        If Abs(mlngHeight - UserControl.Height) > 20 Then
            'Debug.Print "--大小改变进入--"
            Call zlMenuClick("装载数据", mstrParam)
        End If
    End If
    
    Call RaisEffect(picSerach, 2)
    Call CalcScrollBarSize
End Sub

Private Sub UserControl_Terminate()
    Call ReleaseObj
End Sub

Private Sub vsb_Change()
    picMain.Top = -1 * vsb.Value * msinVStep
End Sub

Private Sub mfrmCaseTendBodyPrint_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub

'------绘图相关函数

Private Sub Paint_Init(ByVal objDraw As Object, ByVal objBuffer As Object)

    On Error GoTo Errhand

    '绘图前的初始化工作
    '入参：主窗体的句柄
    RGB_BLACK = RGB(0, 0, 0)
    RGB_RED = RGB(255, 0, 0)
    RGB_WRITE = RGB(255, 255, 255)
    RGB_BLUE = RGB(0, 0, 255)
    RGB_GRAY = &H808080
    RGB_FleetGRAY = &HC0C0C0
    mblnRedraw = True
    
    mlngHwnd = objDraw.hWnd
    Set mobjDraw = objDraw
    Set mobjBuffer = objBuffer
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    
    '先进性对象释放
    Call Paint_Destory
    
    '得到客户区域
    Call GetClientRect(GetDesktopWindow, T_ClientRect)      '取得屏幕的有效区域
    '得到当前DC句柄
    mlngDC = GetDC(mlngHwnd)
    '创建兼容DC
    mlngMemDC = CreateCompatibleDC(mlngDC)
    '创建兼容位图，将直接在此位图上作画
    mlngMemBitmap = CreateCompatibleBitmap(mlngDC, T_ClientRect.Right, T_ClientRect.Bottom) '必须是源DC才能保证是彩色的位图
    '在兼容DC中使用创建的兼容位图
    mlngOldBitmap = SelectObject(mlngMemDC, mlngMemBitmap)
    
    Call SetBkMode(mlngMemDC, TRANSPARENT)
    
    '创建临时刷子设置背景色
    Dim lngBrush As Long, lngOldBrush As Long

    '创建白色刷子
    lngBrush = GetStockObject(WHITE_BRUSH)
    '使用该刷子填充背景色（全白）
    lngOldBrush = SelectObject(mlngMemDC, lngBrush)
    Call FillRect(mlngMemDC, T_ClientRect, lngBrush)
    '立即销毁临时使用的刷子并还原刷子
    Call SelectObject(mlngMemDC, lngOldBrush)
    Call DeleteObject(lngBrush)
    '将所有体温曲线项目的图形装载入内存
    Call PrepareGraph

    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Reset()

    On Error GoTo Errhand

    '恢复画布到初始状态
    
    '创建临时刷子设置背景色
    Dim lngBrush As Long, lngOldBrush As Long

    '创建白色刷子
    lngBrush = GetStockObject(WHITE_BRUSH)
    '使用该刷子填充背景色（全白）
    lngOldBrush = SelectObject(mlngMemDC, lngBrush)
    Call FillRect(mlngMemDC, T_ClientRect, lngBrush)
    '立即销毁临时使用的刷子并还原刷子
    Call SelectObject(mlngMemDC, lngOldBrush)
    Call DeleteObject(lngBrush)
    
    If Not mrsDrawItems Is Nothing Then If mrsDrawItems.State = 1 Then mrsDrawItems.Close
    '所有曲线项目的作图区域(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式,颜色,警戒线)
    gstrFields = "项目序号," & adDouble & ",18|最大值," & adDouble & ",18|最小值," & adDouble & ",18|" & _
        "单位值," & adDouble & ",18|最大值坐标," & adLongVarChar & ",20|最小值坐标," & adLongVarChar & ",20|" & _
        "单位刻度," & adLongVarChar & ",20|显示模式," & adDouble & ",5|颜色," & adDouble & ",18"
    Call Record_Init(mrsDrawItems, gstrFields)
    
    mblnRedraw = True           '因mrsDrawItems清空,所以强制刷新
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Destory()

    On Error GoTo Errhand

    '销毁所有对象
    If mlngOldBitmap <> 0 Then Call SelectObject(mlngMemDC, mlngOldBitmap)
    If mlngMemBitmap <> 0 Then Call DeleteObject(mlngMemBitmap)
    If mlngMemDC <> 0 Then Call DeleteDC(mlngMemDC)
    If mlngDC <> 0 Then Call ReleaseDC(mlngHwnd, mlngDC)
    mlngOldBitmap = 0
    mlngMemBitmap = 0
    mlngMemDC = 0
    mlngDC = 0
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Canvas(Optional ByVal blnAdjust As Boolean = False)
    '准备画布（完成刻度及表格缩放、三测单表格绘制以及根据设定进行基准线的描绘）
    '最小模式下,不显示表上表格,文本等数据
    'blnAdjust=False表示固定大小，否则跟随主界面进行调整
    
    Static SlngMaxY As Long                 '记录上一次的最大高度，以决定本次是否需要重画
    Dim lngCurX     As Long, lngCurY As Single  '当前位置
    Dim lngMaxX     As Long, lngMaxY As Single  '边界
    Dim lngCurAlerY As Single
    Dim lngRow      As Long
    Dim intLables   As Integer
    Dim bln双行 As Boolean                  '此参数由用户指定,bln双行=TRUE表示只显示五行;否则显示十行
    Dim bln粗线 As Boolean                  '此参数由用户指定,大行分界是粗线还是细线
    Dim rsTemp        As New ADODB.Recordset
    
    '以下都是标准尺度
    Dim intLineMode   As Integer
    Dim blnDoubleRow  As Boolean             '贰行做为一行打印输出
    Dim intTens_digit As Integer            '3：以10的倍数输出；2：以5的倍数输出；1：是个位是整数则输出
    Dim sinAlertness  As Single              '警戒线,起辅助作用
    Dim lngLableStep  As Long
    Dim lngColStep    As Long
    Dim sigRowStepNew As Single, sinRowStep As Single, lngInitRowStep As Long
    Dim lng最高行 As Long, lngMaxRows As Long
    Dim lng体温最小值 As Long
    Dim arrTemp()     As String
    Dim sinY单位 As Single '曲线单位输出的Bottom
    Dim lngCurveRow As Long

    '以下与绘图区域相关(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    Dim sin刻度 As Single, bln显示刻度 As Boolean
    Dim sin刻度间隔 As Single, sinBegin刻度 As Single, dbl单位值 As Double

    Dim str最大值坐标 As String, str最小值坐标 As String

    On Error GoTo Errhand
    
    '实现缩放的原理说明：
    '1、普通模式下所有内容均显示
    '2、最小模式=2，时间刻度不显示，每行10小行改为5小行
    '3、缩小模式<=4，转为虚线显示
    
    '以前是固定以上面有2行来输出数据，所以此处减去2行
    '后面输出时为了对齐好看，再次减2行来输出
    lngCurveRow = Val(zldatabase.GetPara("体温曲线固定添加行数", glngSys, 1255, "0"))
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    T_DrawClient.总列数 = glngMaxRows
    
    gstrSQL = " Select A.项目序号,A.排列序号,A.记录名,A.记录符,A.记录色,nvl(A.最大值,0) 最大值,nvl(A.最小值,0) 最小值," & _
        "nvl(A.单位值,0) 单位值,A.刻度间隔,A.警示线,C.项目单位 单位,nvl(A.最高行,2)-2 AS 最高行,B.部位 " & _
        " From 体温记录项目 A,体温部位 B,护理记录项目 C" & _
        " Where A.项目序号=B.项目序号(+) And B.缺省项(+)=1" & _
        " And A.记录法=1 And A.项目序号=C.项目序号 and nvl(C.应用方式,0)=1 and C.护理等级>=[1]" & _
        " and nvl(C.适用病人,0) in (0,[2]) and (C.适用科室=1 or (C.适用科室=2 and Exists (select 1 from 护理适用科室 D where C.项目序号=D.项目序号 and D.科室ID=[3])))" & _
        " Order by 排列序号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "取开始行", T_Patient.lng护理等级, IIf(T_Patient.lng婴儿 = 0, 1, 2), T_Patient.lng科室ID)
    
    '------------------------------------------------------------------------------------------------------------------
    rsTemp.Filter = "项目序号=" & gint体温
    '计算打印输出的行数
    With rsTemp
        Do While Not .EOF
            lng最高行 = Val(zlCommFun.Nvl(!最高行))
            If lng最高行 < 0 Then lng最高行 = 0
            
             '修改问题51442
            If Val(zlCommFun.Nvl(!最小值, 0)) > 34 Then
                lngMaxRows = lng最高行 + (Val(zlCommFun.Nvl(!最大值, 0)) - 35) / 0.1 + 10
            Else
                lngMaxRows = lng最高行 + (Val(zlCommFun.Nvl(!最大值, 0)) - Val(zlCommFun.Nvl(!最小值, 0))) / 0.1
            End If

            lngMaxRows = lngMaxRows + lngCurveRow
            
            If lngMaxRows > T_DrawClient.总列数 Then
                T_DrawClient.总列数 = lngMaxRows
            End If
        .MoveNext
        Loop
    End With
    
    rsTemp.Filter = 0
    rsTemp.Sort = "排列序号"
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    
    '赋初值
    intLineMode = PS_SOLID
    lngLableStep = glngLableStep
    lngColStep = glngColStep
    lngInitRowStep = glngInitRowStep
    sigRowStepNew = lngInitRowStep
    intTens_digit = 3
    
    '体温单以单格显示(不勾此选项以双格显示，没两个刻度显示一次) 1：单格显示 0：双格显示
    If zldatabase.GetPara("体温单显示格式", glngSys, 1255, 0) = 1 Then
        bln双行 = False
    Else
        bln双行 = True
    End If
    'True表示贰行只输出一行,效果是一个刻度只显示了五行;否则一个刻度显示十行,由用户调整参数决定,与blnDoubleRow无关
    bln粗线 = True
    
    If Not bln粗线 Then intLineMode = PS_DASHDOTDOT
    
    '画表格
    intLables = rsTemp.RecordCount
    lngCurX = T_DrawClient.偏移量X
    lngCurY = T_DrawClient.偏移量Y
    lngMaxX = (intLables * lngLableStep) + (7 * 6 * lngColStep) + T_DrawClient.偏移量X  '刻度+7*宽度 +T_DrawClient.偏移量X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.总列数 * sigRowStepNew + T_DrawClient.偏移量Y '（为表格大小，还需加上起始Y坐标,上面固定留上6行为输出时间信息）
    
    '进行相关数据的校正
    If blnAdjust Then

        '如果小于可见区域大小则进行缩放
        If lngMaxY > mlng高度 Then
            lngMaxY = mlng高度 - 2 * mintNullRow * lngInitRowStep
            sigRowStepNew = Round((lngMaxY) / T_DrawClient.总列数, 1)
        End If

        '如果行高太小，则将贰行做为一行显示
        If sigRowStepNew <= 2 Then
            sinRowStep = 1.5
            blnDoubleRow = True
        End If

        '计算刻度的最大坐标
        lngMaxY = T_DrawClient.总列数 * IIf(blnDoubleRow, sinRowStep, sigRowStepNew) + T_DrawClient.偏移量Y + lngInitRowStep * 2 * mintNullRow

        If Not mblnRedraw Then mblnRedraw = (lngMaxY <> SlngMaxY)
        If sigRowStepNew < 4 Then intLineMode = PS_DOT
    End If
    
    '进行刻度的校正(当曲线项目小于3时)
    If intLables <= 3 Then
        lngLableStep = glngLableWith / intLables
        lngMaxX = (intLables * lngLableStep) + (7 * 6 * lngColStep) + T_DrawClient.偏移量X     '刻度+7*宽度 +偏移量X
    End If
    
    lblCommText.Caption = ""
    
    Call Paint_Reset                                                    '清除画布
    
    SlngMaxY = lngMaxY
    T_DrawClient.刻度单位 = lngLableStep
    T_DrawClient.行单位 = IIf(blnDoubleRow, sinRowStep, sigRowStepNew)
    T_DrawClient.列单位 = lngColStep
    T_DrawClient.双倍 = blnDoubleRow
    
    
    '画刻度区域
'    For lngRow = 1 To intLables
'        Call DrawRect(mlngMemDC, lngCurX - IIf(lngRow = 1, 0, 1), lngCurY, lngCurX + lngLableStep + 1, lngMaxY, PS_SOLID, 1, RGB_BLACK)
'        lngCurX = lngCurX + lngLableStep
'    Next
    
    For lngRow = 1 To intLables
         Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
         lngCurX = lngCurX + lngLableStep
    Next
    Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
    
    '画刻度框
    Call DrawLine(mlngMemDC, T_DrawClient.偏移量X, lngCurY, lngMaxX, lngCurY, PS_SOLID, 1, RGB_BLACK)

    T_DrawClient.刻度区域.Left = T_DrawClient.偏移量X
    T_DrawClient.刻度区域.Top = lngCurY
    T_DrawClient.刻度区域.Right = lngCurX
    T_DrawClient.刻度区域.Bottom = lngMaxY
    
    '默认添加一行用于显示项目名称
    lngCurY = lngCurY + lngInitRowStep * 2
    Call DrawLine(mlngMemDC, T_DrawClient.偏移量X, lngCurY, lngMaxX, lngCurY, PS_SOLID, 1, RGB_BLACK)
    lngCurY = lngCurY + lngInitRowStep * ((mintNullRow - 1) * 2)
    
    '画体温单所有行
    For lngRow = 0 To T_DrawClient.总列数
        If lngRow <> 0 Then
            lngCurY = lngCurY + IIf(blnDoubleRow, sinRowStep, sigRowStepNew)
        End If
        '画体温单的所有行
        If ((blnDoubleRow Or bln双行) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln双行) Then
            Call DrawLine(mlngMemDC, lngCurX + 1, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sigRowStepNew >= 4 And bln粗线, 2, 1), RGB_FleetGRAY)
        End If
    Next
    
    lngCurY = T_DrawClient.刻度区域.Top
    
    '画体温单所有列
    For lngRow = 1 To 6 * 7
        lngCurX = lngCurX + lngColStep
        
        Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod 6 = 0, 2, 1), IIf(lngRow Mod 6 = 0, RGB_RED, RGB_GRAY))
    Next
    
    lngCurX = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Left = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Top = T_DrawClient.刻度区域.Top
    T_DrawClient.体温区域.Right = lngMaxX
    T_DrawClient.体温区域.Bottom = lngMaxY
    
    '画体温区域底线
    Call DrawLine(mlngMemDC, T_DrawClient.偏移量X, lngMaxY - 1, lngMaxX, lngMaxY - 1, PS_SOLID, 1, RGB_BLACK)

    '画刻度框的标尺（从固定不变的10行开始标识）
    With rsTemp
        Do While Not .EOF
            '显示刻度框项目的名称及符号,如体温×
            lngCurX = T_DrawClient.刻度区域.Left + ((.AbsolutePosition - 1) * T_DrawClient.刻度单位)
            lngCurY = T_DrawClient.刻度区域.Top
            
            '设置字体大小
            gstdSet.Name = "宋体"
            gstdSet.Size = 9
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)
            
            '输出体温项目的名称
            Call SetTextColor(mlngMemDC, zlCommFun.Nvl(!记录色, RGB_BLACK))
            Call GetTextRect(mobjDraw, lngCurX, lngCurY + mobjDraw.TextHeight(zlCommFun.Nvl(!记录名)) / Screen.TwipsPerPixelY / 2, Trim(zlCommFun.Nvl(!记录名)), T_DrawClient.刻度单位)
            Call DrawText(mlngMemDC, Trim(zlCommFun.Nvl(!记录名)), -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            
            '设置字体大小
            gstdSet.Name = "宋体"
            gstdSet.Size = 8
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)

            '输出项目单位
            Call GetTextRect(mobjDraw, lngCurX, lngCurY + lngInitRowStep * 2 + mobjDraw.TextHeight(zlCommFun.Nvl(!单位)) / Screen.TwipsPerPixelY / 2, Trim(zlCommFun.Nvl(!单位)), T_DrawClient.刻度单位)
            Call DrawText(mlngMemDC, Trim(zlCommFun.Nvl(!单位)), -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            mobjDraw.Font.Size = 9
            sinY单位 = T_LableRect.Bottom
            '创建字体
            gstdSet.Name = "宋体"
            gstdSet.Size = 9
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)
            '输出体温项目的标识符
            'Call DrawMarker(False, !项目序号, zlCommFun.NVL(!部位, "空"), lngCurX + T_TwipsPerPixel.x / 2, lngCurY + 15, "空", True)
    
            '强制设定体温曲线项目的显示模式
            Select Case !项目序号

                Case gint体温  '体温整数时输出刻度
                    intTens_digit = 1
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 1)
                    dbl单位值 = 0.1
                    sinAlertness = zlCommFun.Nvl(!警示线, 37)
                    arrTemp = Split(zlCommFun.Nvl(!记录符, "·,×,○"), ",")
                    lblCommText.Caption = lblCommText.Caption & "、" & zlCommFun.Nvl(!记录名) & "(口温" & arrTemp(0) & ",腋温" & arrTemp(1) & ",肛温" & arrTemp(2) & ")"

                Case gint脉搏, gint心率  '脉搏/心跳按10的倍数输出刻度
                    intTens_digit = 3
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 10)
                    dbl单位值 = 2
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
                    
                    If !项目序号 = gint脉搏 Then
                        lblCommText.Caption = lblCommText.Caption & "、" & zlCommFun.Nvl(!记录名) & "(缺省记录符" & zlCommFun.Nvl(!记录符, "+") & ",起搏器H)"
                    Else
                        lblCommText.Caption = lblCommText.Caption & "、" & zlCommFun.Nvl(!记录名) & "(" & zlCommFun.Nvl(!记录符, "Ο") & ")"
                    End If

                Case gint呼吸  '呼吸按5的倍数输出刻度
                    mbln呼吸曲线 = True
                    intTens_digit = 2
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 5)
                    dbl单位值 = 1
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
                    lblCommText.Caption = lblCommText.Caption & "、" & zlCommFun.Nvl(!记录名) & "(自主呼吸" & zlCommFun.Nvl(!记录符, "*") & ",呼吸机R)"
                Case Else
                    intTens_digit = 1
                    dbl单位值 = Val(zlCommFun.Nvl(!单位值, 1))
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, Val(zlCommFun.Nvl(!单位值, 0)) * 10)
                    If sin刻度间隔 > Val(zlCommFun.Nvl(!最大值)) - Val(zlCommFun.Nvl(!最小值)) Then
                        sin刻度间隔 = Val(zlCommFun.Nvl(!最大值)) - Val(zlCommFun.Nvl(!最小值))
                    End If
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
                    lblCommText.Caption = lblCommText.Caption & "、" & zlCommFun.Nvl(!记录名) & "(" & zlCommFun.Nvl(!记录符, "*") & ")"
            End Select

            '赋初值
            lngCurY = lngCurY + (lngInitRowStep * 2 * mintNullRow)   '固定前2 * mintNullRow行的高度不输出刻度

            '如果是最小模式,从第30行开始输出标识
            'If blnDoubleRow Then lngCurY = lngCurY + lngInitRowStep * 2 * mintNullRow
            
            '根据最高行定位到有效位置
            lngCurY = lngCurY + (T_DrawClient.行单位 * zlCommFun.Nvl(!最高行, 2))
            
            Do While True
                bln显示刻度 = False
                If sin刻度 = 0 Then     '刚进入循环，此时取的最大值
                    sin刻度 = zlCommFun.Nvl(!最大值, 0)
                    sinBegin刻度 = sin刻度
                    str最大值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                Else                    '计算得到每个刻度的值
                    sin刻度 = sin刻度 - dbl单位值    '如果目前显示模式为双倍，则按双倍累计
                End If
                
                If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
                
                If bln显示刻度 = True Or sin刻度 < sinBegin刻度 Then sinBegin刻度 = sinBegin刻度 - IIf(T_DrawClient.双倍, sin刻度间隔 * 2, sin刻度间隔)
                
                If sinBegin刻度 < 0 Then sinBegin刻度 = 0
                
                If bln显示刻度 Then
                    '控制最大值不与曲线单位重复
                    If sin刻度 = Val(Nvl(!最大值, 0)) And lngCurY < sinY单位 Then
                        Call GetTextRect(mobjDraw, lngCurX, sinY单位, Format(sin刻度, "#0"), T_DrawClient.刻度单位)
                    ElseIf Format(lngCurY, "#0") = T_DrawClient.刻度区域.Bottom Then
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY - (mobjDraw.TextHeight("1") / (T_TwipsPerPixel.Y * 2)), Format(sin刻度, "#0"), T_DrawClient.刻度单位)
                    Else
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY, Format(sin刻度, "#0"), T_DrawClient.刻度单位)
                    End If
                    Call DrawText(mlngMemDC, Format(sin刻度, "#0"), -1, T_LableRect, DT_CENTER)
                End If
                
                '如果不在有效范围内，或者超出画布则退出
                If Val(Format(sin刻度, "#0.00")) <= Val(Format(!最小值, "#0.00")) Or Format(lngCurY, "#0") > T_DrawClient.刻度区域.Bottom Then
                    str最小值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                    '添加该项目(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
                    gstrFields = "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|颜色"
                    gstrValues = zlCommFun.Nvl(!项目序号) & "|" & zlCommFun.Nvl(!最大值) & "|" & zlCommFun.Nvl(!最小值) & "|" & dbl单位值 & "|" & _
                        str最大值坐标 & "|" & str最小值坐标 & "|" & T_DrawClient.行单位 & "," & T_DrawClient.列单位 & "|" & intTens_digit & "|" & !记录色
                    Call Record_Add(mrsDrawItems, gstrFields, gstrValues)
                    
                    '辅助线或警示线
                    If blnDoubleRow = False And (sinAlertness < Val(Nvl(!最大值)) And sinAlertness > Val(Nvl(!最小值))) Then
                        lngCurAlerY = Val(GetYCoordinate(mobjDraw, mrsDrawItems, Val(Nvl(!项目序号)), sinAlertness))
                        Call DrawLine(mlngMemDC, T_DrawClient.体温区域.Left, lngCurAlerY, lngMaxX, lngCurAlerY, PS_SOLID, 1, RGB_RED)
                    End If
                    Exit Do
                End If
                
                lngCurY = lngCurY + T_DrawClient.行单位
            Loop
            
            '还原字体信息
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
                
            sinBegin刻度 = 0
            sin刻度 = 0                 '控制从第一行开始输出
            .MoveNext
        Loop
    End With
        
    '创建字体
    gstdSet.Name = "宋体"
    gstdSet.Size = 9
    Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
    mlngFont = CreateFontIndirect(T_Font)
    mlngOldFont = SelectObject(mlngMemDC, mlngFont)
    
    lblCommText.Caption = "说明:" & Mid(lblCommText.Caption, 2)
    mblnRedraw = False                      '画过一次后就不再画了
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Sub Paint_Date()
'-------------------------------------
'功能:罗列显示出体温单上的时间
'说明:此函数目前未使用
'------------------------------------
    Dim i As Long, j As Long
    Dim lngColor  As Long
    Dim intMinCol As Long, intMaxCol As Long
    Dim strTmp As String
    
    On Error GoTo Errhand
    
    If picMain.Tag <> "" Then
        Call CalcMinMaxCol(picMain.Tag, intMinCol, intMaxCol)
    End If
    
    For i = 1 To 7
        '输出上午 下午信息
        Call SetTextColor(mlngMemDC, RGB_BLACK)
        Call GetTextRect(mobjDraw, T_DrawClient.体温区域.Left + (i - 1) * 6 * T_DrawClient.列单位, T_DrawClient.体温区域.Top + T_DrawClient.时间行单位 * 2, "上午", T_DrawClient.列单位 * 3)
        Call DrawText(mlngMemDC, "上午", -1, T_LableRect, DT_CENTER)

        Call SetTextColor(mlngMemDC, RGB_BLACK)
        Call GetTextRect(mobjDraw, T_DrawClient.体温区域.Left + 3 * T_DrawClient.列单位 + (i - 1) * 6 * T_DrawClient.列单位, T_DrawClient.体温区域.Top + T_DrawClient.时间行单位 * 2, "下午", T_DrawClient.列单位 * 3)
        Call DrawText(mlngMemDC, "下午", -1, T_LableRect, DT_CENTER)
        
        '输出时间信息
        For j = 1 To 6

            Select Case j

                Case 1
                    strTmp = gintHourBegin + 4 * 0
                    lngColor = &H8080FF

                Case 2
                    strTmp = gintHourBegin + 4 * 1
                    lngColor = &H8080FF

                Case 3
                    strTmp = gintHourBegin + 4 * 2
                    lngColor = &H80000012

                Case 4
                    lngColor = &H80000012
                    strTmp = gintHourBegin + 4 * 0

                Case 5
                    lngColor = &H80000012
                    strTmp = gintHourBegin + 4 * 1

                Case 6
                    lngColor = &H8080FF
                    strTmp = gintHourBegin + 4 * 2
            End Select
            
            If j + (i - 1) * 6 >= intMinCol And j + (i - 1) * 6 <= intMaxCol Then
                lngColor = lngColor
            Else
                lngColor = RGB_FleetGRAY
            End If
            
            If picMain.Tag <> "" Then
                Call SetTextColor(mlngMemDC, lngColor)
                Call GetTextRect(mobjDraw, T_DrawClient.体温区域.Left + ((i - 1) * T_DrawClient.列单位 * 6) + ((j - 1) * T_DrawClient.列单位), T_DrawClient.体温区域.Top + T_DrawClient.时间行单位 * 6, Trim(strTmp), T_DrawClient.列单位)
                Call DrawText(mlngMemDC, Trim(strTmp), -1, T_LableRect, DT_CENTER)
            End If

        Next j
    Next i
    
    Exit Sub

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Construct()

    Dim lngRGB  As Long

    Dim blnLine As Boolean              '心率与脉搏共用时,心率不连线

    Dim str心率 As String               '记录所有心率所在的列(X坐标)

    Dim str原值 As String, sinX坐标原 As Single, sinY坐标原 As Single
    
    Dim dbl数值 As Double, dblMinValue As Double, dblMaxValue As Double
    Dim bln不升符号 As Boolean
    Dim lng体温不升显示方式 As Long
    
    
    On Error GoTo Errhand

    '开始作图：完成图形数据的输出（重叠的处理、体温复核、图形标记输出、脉搏短拙）
    '先画线(除心率外)
    '再处理脉搏短拙
    '再输出图形
    
    lng体温不升显示方式 = Val(zldatabase.GetPara("体温不升显示方式", glngSys, 1255, "0"))
    
    With mrsPoint
        .Filter = ""
        '先画线
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "项目序号,时间"
        Do While Not .EOF
            If Val(zlCommFun.Nvl(!状态)) <> 3 Then
                '物理降温的后面处理,不连线
                If Not (!项目序号 = gint体温 And !标记 = 1) Then
                    If str原值 <> !项目序号 Then
                        dblMinValue = GetMinValue(!项目序号)
                        dblMaxValue = GetMaxValue(!项目序号)
                        blnLine = True
                        mrsDrawItems.Filter = "项目序号=" & !项目序号
                        If mrsDrawItems.RecordCount = 0 Then
                            '与脉搏共用则不连线
                            blnLine = False
                            mrsDrawItems.Filter = "项目序号=" & gint脉搏
                        End If

                        lngRGB = mrsDrawItems!颜色
                        mrsDrawItems.Filter = 0
                        
                        sinX坐标原 = 0
                        sinY坐标原 = 0
                        str原值 = !项目序号
                    End If
                    
                    '复查合格
                    If !项目序号 = gint体温 And Val(zlCommFun.Nvl(!复查)) = 1 Then
                        Call SetTextColor(mlngMemDC, lngRGB)
                        Call GetTextRect(mobjDraw, !X坐标, !Y坐标 - Screen.TwipsPerPixelY, "v", T_DrawClient.列单位, False)
                        Call DrawText(mlngMemDC, "v", -1, T_LableRect, DT_CENTER)
                    End If
                    
                    If sinX坐标原 <> 0 And blnLine Then
                        Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, sinX坐标原 + T_DrawClient.列单位 / 2, sinY坐标原, PS_SOLID, 1, lngRGB)
                    End If
                    
                    If !断开 = 0 Then
                        sinX坐标原 = !X坐标
                        sinY坐标原 = !Y坐标
                    Else
                        sinX坐标原 = 0
                    End If
                    
                End If

                '此处处理项目高出项目的最大值 或小于项目最小值
                If !项目序号 = gint体温 And Trim(Nvl(!数值)) = "不升" Then
                    dbl数值 = dblMinValue
                Else
                    dbl数值 = Val(zlCommFun.Nvl(!数值))
                End If
                '重叠时以序号靠前的为准
                If !重叠 = 0 Then
                    If dbl数值 < dblMinValue Then
                        Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 + IIf(T_DrawClient.行单位 < glngInitRowStep, glngInitRowStep, T_DrawClient.行单位) * 2, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, 1, lngRGB, True)
                    End If
                    
                    If dbl数值 > dblMaxValue Then
                        Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 - IIf(T_DrawClient.行单位 < glngInitRowStep, glngInitRowStep, T_DrawClient.行单位) * 2, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, 1, lngRGB, True)
                    End If
                End If
            End If

            .MoveNext
        Loop
        
        '画脉搏短拙连线(脉搏必须连续,心率一个点对应连续的三个脉搏,心率连则区域增大;如果中断或只有点点对应,则只连一线)
        If .RecordCount <> 0 Then .MoveFirst
        .Filter = "项目序号=" & gint心率

        Do While Not .EOF
            str心率 = str心率 & "," & !X坐标 & ";" & !Y坐标
            .MoveNext
        Loop

        If str心率 <> "" Then str心率 = Mid(str心率, 2)
        .Filter = 0

        '形成封闭区域并填充
        If str心率 <> "" Then Call CreatePoly(mrsPoint, mobjDraw, mlngMemDC, mstr开始时间, str心率)

        '输出点或图形
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "项目序号,时间"

        Do While Not .EOF
            If Val(zlCommFun.Nvl(!状态)) <> 3 Then
                If !项目序号 = gint体温 And !标记 = 1 Then
                    '体温的物理降温输出红色的空心圆
                    '字符输出
                    Call SetTextColor(mlngMemDC, RGB_RED)
                    Call GetTextRect(mobjDraw, !X坐标, !Y坐标, "○", T_DrawClient.列单位)
                    Call DrawText(mlngMemDC, "○", -1, T_LableRect, DT_CENTER)
                    T_Size.H = mobjDraw.TextHeight("○") / Screen.TwipsPerPixelY
                    str原值 = Split(!备注, ",")(0)
                    sinX坐标原 = Val(Split(Split(!备注, ",")(1), ";")(0))
                    sinY坐标原 = Val(Split(Split(!备注, ",")(1), ";")(1))
                    
                    If Val(!数值) > Val(str原值) Then
                        '物理降温失败，画带箭头的红色实线，字符固定用○
                        'Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, sinX坐标原 + T_DrawClient.列单位 / 2, sinY坐标原, PS_SOLID, 1, RGB_RED, True)
                        '现在失败也为虚线(医院要求)
                        Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 + (T_Size.H / 4), sinX坐标原 + T_DrawClient.列单位 / 2, sinY坐标原, PS_DOT, 1, RGB_RED, True)
                    ElseIf Val(!数值) < Val(str原值) Then
                        '物理降温成功，画红色虚线，字符固定用○ 不画箭头直接以虚线连接
                        Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 - (T_Size.H / 2), sinX坐标原 + T_DrawClient.列单位 / 2, sinY坐标原, PS_DOT, 1, RGB_RED, False)
                    End If
                Else
                    If !项目序号 = gint体温 And Trim(Nvl(!数值)) = "不升" And (lng体温不升显示方式 = 0 Or lng体温不升显示方式 = 1) Then
                        bln不升符号 = False
                    Else
                        bln不升符号 = True
                    End If
                    
                    If !重叠 = 0 And bln不升符号 Then
                        Call DrawMarker(True, !项目序号, !部位, !X坐标, !Y坐标, !重叠项目, False, !符号)
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub Paint_Assistant()

    '从42度开始打印,如有多项内容允许递延在后面的列打印,最后一列不必考虑能否打全
    Dim rsNote As New ADODB.Recordset
    Dim byt未记说明显示位置 As Byte
    Dim Y As Long, X As Long, Y1 As Long
    Dim bln改变 As Boolean
    Dim lngX          As Long, lngY As Long
    Dim strComment    As String, strTemp As String, strText As String
    Dim intNum        As Integer
    Dim intAscCharNum As Integer
    Dim varNote()     As String
    Dim i  As Integer, j As Integer

    On Error GoTo Errhand

    '辅助作图：完成文字的输出（体温的不升、呼吸次数小于10、上下标与未记说明、手术分娩入出转、特殊用药等）
    
    'mrsNote "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标"
    '类型说明 2-上标;3-入出转;4-手术日;6-下标,99-未记说明
    
    '处理未记说明信息
    mrsNote.Filter = "类型=99"
    mrsNote.Sort = "X坐标"
    With mrsNote
        Do While Not .EOF
            If X = !X坐标 Then
                If InStr(1, "," & strTemp & ",", "," & zlCommFun.Nvl(!内容) & ",") <> 0 Then
                    mrsNote.Delete
                Else
                    strTemp = strTemp & "," & zlCommFun.Nvl(!内容)
                End If
            Else
                X = !X坐标
                strTemp = zlCommFun.Nvl(!内容)
            End If
        .MoveNext
        Loop
    End With
    
    byt未记说明显示位置 = Val(zldatabase.GetPara("未记说明显示位置", glngSys, 1255, "0"))
    
    Y1 = GetYCoordinate(mobjDraw, mrsDrawItems, gint体温, 42, mlngMemDC)
    
    mrsNote.Filter = 0
    mrsNote.Sort = "X坐标,项目序号"
    With mrsNote
        Do While Not .EOF
            If !类型 = 99 Then '根据参数设置检查未记说明显示方式
                varNote = Split(!内容, ";")
                strComment = ""
                strTemp = ""

                For i = 0 To UBound(varNote)
                    '未记说明显示在上方 体温不升作为下标，否则作废上标
                    If Not (varNote(i) = "不升" And byt未记说明显示位置 = 0) And varNote(i) <> "" Then
                        If InStr(1, strTemp, varNote(i)) = 0 Then
                            strTemp = IIf(strTemp = "", varNote(i), strTemp & ";" & varNote(i))
                        End If
                    End If
                Next i
                
                If strTemp <> "" Then
                    strComment = ""
                    varNote = Split(strTemp, ";")

                    For i = 0 To UBound(varNote)

                        If strComment = "" Then
                            strComment = varNote(i)
                        Else
                            strComment = strComment & " " & varNote(i)
                        End If

                    Next i

                End If
                
                '根据参数判断是否直接输未记说明
                If byt未记说明显示位置 = 1 Then
                    Y = GetYCoordinate(mobjDraw, mrsDrawItems, gint体温, 35, mlngMemDC)
                    If lngY <> 0 And X = Val(!X坐标) Then Y = lngY: strComment = " " & strComment
                    X = Val(!X坐标)

                    If strComment <> "" Then
                        strComment = Replace(strComment, ";", " ")
                        '输出信息未记说明
                        For i = 1 To Len(strComment)

                            If Y < T_DrawClient.刻度区域.Bottom Then
                                strText = Mid(strComment, i, 1)
                                Call GetTextExtentPoint32(mlngMemDC, strText, Len(strText), T_Size)
                                '输出字体信息
                                If T_DrawClient.刻度区域.Bottom - Y > T_Size.H Then
                                    Call DrawRotateText(mobjDraw, mlngMemDC, X, Y, strText, !颜色)
                                End If
                                If Asc(strText) < 0 Then
                                    Y = Y + T_Size.H
                                Else
                                    Y = Y + T_Size.H / 2
                                End If
                            End If

                        Next i

                        strComment = " "
                        lngY = Y
                    End If

                    mrsNote!禁用 = 1
                Else
                    mrsNote!内容 = strComment
                    strComment = ""
                    mrsNote!Y坐标 = Y1
                    lngY = 0
                End If

            ElseIf !类型 = 6 Then '输出下标说明
                Y = GetYCoordinate(mobjDraw, mrsDrawItems, gint体温, 35, mlngMemDC)
                strComment = ""
                If lngY <> 0 And X = Val(!X坐标) Then Y = lngY: strComment = " "
                X = Val(!X坐标)
                
                '如果未记说明输出在下方，此处检测如果该下标上存在未记说明,以便保证格式
                If strComment <> "" Then
                    If Asc(strComment) < 0 Then
                        intNum = 0
                    Else
                        intNum = 1
                    End If
                End If
                
                '输出信息未记说明
                strComment = strComment & !内容
                intAscCharNum = 0
                If strComment <> "" Then
                    strComment = Replace(strComment, ";", " ")
                End If
                For i = 1 To Len(strComment)
                    If Y < T_DrawClient.刻度区域.Bottom Then
                        strText = Mid(strComment, i, 1)
                        Call GetTextExtentPoint32(mlngMemDC, strText, Len(strText), T_Size)

                        If Asc(strText) < 0 Then
                            If (intAscCharNum - intNum) Mod 2 = 1 Then Y = Y + T_Size.H / 2
                        End If
                         
                        '输出字体信息
                        If T_DrawClient.刻度区域.Bottom - Y > T_Size.H Then
                            Call DrawRotateText(mobjDraw, mlngMemDC, X, Y, strText, !颜色)
                        End If
                        If Asc(strText) < 0 Then
                            Y = Y + T_Size.H
                            intAscCharNum = 0
                        Else
                            Y = Y + T_Size.H / 2
                            intAscCharNum = intAscCharNum + 1
                        End If
                    End If
                Next i
                mrsNote!禁用 = 1
                lngY = 0
                strComment = ""
            Else
                mrsNote!Y坐标 = Y1
            End If

            .MoveNext
        Loop

    End With
    
    If mrsNote.RecordCount > 0 Then
        mrsNote.MoveFirst
        mrsNote.Update
    End If
    '输出体温标记信息(入出转、手术分娩 上标等信息)
    Call OutPutText(mobjDraw, mrsDrawItems, mlngMemDC, mrsNote, mstr开始时间)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub ReleaseObj()
'----------------------------------------------------------
'关闭掉所有对象
'----------------------------------------------------------
    If Not (mrsNote Is Nothing) Then Set mrsNote = Nothing
    If Not (mrsItems Is Nothing) Then Set mrsItems = Nothing
    If Not (mrsPoint Is Nothing) Then Set mrsPoint = Nothing
    If Not (mrsGraph Is Nothing) Then Set mrsGraph = Nothing
    If Not (mrsDrawItems Is Nothing) Then Set mrsDrawItems = Nothing
    If Not (mrsTabTime Is Nothing) Then Set mrsDrawItems = Nothing
    If Not (mrsCollect Is Nothing) Then Set mrsCollect = Nothing
    Set mobjDraw = Nothing
    Set mobjBuffer = Nothing
    Set gstdSet = Nothing
    
    mintPage = 0
    mintAllPage = 0
    mblnKeyDown = False
    mstrParam = ""
    mstrParam1 = ""
    mstrParam2 = ""
    If Not (mfrmParent Is Nothing) Then Set mfrmParent = Nothing
    
    Call Paint_Destory
End Sub

'-------------------体温单数据提取 函数
Private Function GetMinValue(ByVal lng项目序号 As Long) As Double
    '-------------------------------------------------------------------
    '功能:根据项目序号获取最小值
    '说明:目前根据项目值域确定可录入的范围，最大值最小值确定点的最大坐标和最小坐标，超出以箭头显示
    '-------------------------------------------------------------------
    Dim dblvalue As Double
    
    mrsItems.Filter = "项目序号=" & lng项目序号
    If mrsItems.EOF Then Exit Function
    
'    If InStr(1, Nvl(mrsItems!项目值域), ";") = 0 Then
'        dblvalue = Val(Nvl(mrsItems!最小值, 0))
'    Else
'        dblvalue = Val(Split(mrsItems!项目值域, ";")(0))
'    End If
    dblvalue = Val(Nvl(mrsItems!最小值, 0))
    GetMinValue = dblvalue
End Function

Private Function GetMaxValue(ByVal lng项目序号 As Long) As Double
    '-------------------------------------------------------------------
    '功能:根据项目序号获取最大值
    '说明:目前根据项目值域确定可录入的范围，最大值最小值确定点的最大坐标和最小坐标，超出以箭头显示
    '-------------------------------------------------------------------
    Dim dblvalue As Double
    Dim strValue As String
    
    mrsItems.Filter = "项目序号=" & lng项目序号
    If mrsItems.EOF Then Exit Function
    
'    If InStr(1, Nvl(mrsItems!项目值域), ";") = 0 Then
'        dblvalue = Val(Nvl(mrsItems!最大值, 0))
'    Else
'        dblvalue = Val(Split(mrsItems!项目值域, ";")(1))
'        If dblvalue = 0 Then dblvalue = Val(Nvl(mrsItems!最大值))
'    End If
    dblvalue = Val(Nvl(mrsItems!最大值, 0))
    strValue = Nvl(mrsItems!临界值)
    If strValue <> "" And Val(strValue) <= Val(Nvl(mrsItems!最大值)) And Val(strValue) >= Val(Nvl(mrsItems!最小值)) Then dblvalue = strValue
    GetMaxValue = dblvalue
End Function

Private Sub ReadBoyData(ByVal blnAutoAdjust As Boolean)
    
    On Error GoTo Errhand
    
    mint心率应用 = 0
    If Not (mrsItems Is Nothing) Then If mrsItems.State = 1 Then mrsItems.Close
    '打开现存在适用该病人的护理记录项目
    gstrSQL = " Select C.项目序号,C.项目名称,C.项目类型,C.项目性质,C.项目长度,C.项目小数,C.项目表示,C.项目单位,C.项目值域,A.最大值,A.最小值,A.临界值,C.护理等级,C.应用方式,C.适用病人" & _
              " From 体温记录项目 A,护理记录项目 C" & _
              " where A.项目序号(+)=C.项目序号" & _
              " And nvl(C.应用方式,0)<>0" & _
              " and nvl(C.适用病人,0) in (0,[1])" & _
              " and (C.适用科室=1 or (C.适用科室=2 and Exists (select 1 from 护理适用科室 D where C.项目序号=D.项目序号 and D.科室ID=[2])))" & _
              " Order by C.项目序号"
              
    Set mrsItems = zldatabase.OpenSQLRecord(gstrSQL, "取开始行", IIf(T_Patient.lng婴儿 = 0, 1, 2), T_Patient.lng科室ID)
    mrsItems.Filter = "项目序号=-1"
    If mrsItems.RecordCount > 0 Then mint心率应用 = zlCommFun.Nvl(mrsItems("应用方式").Value, 2): mrsItems.Filter = ""
    
    '所有点的表现集合
    '   重叠是否重叠序号.
    '   重叠项目记录重叠项目
    '   断开的条件:超过一天无数据,存在未记说明
    '   状态:0-未编辑;1-新增;2-修改;3-删除
    '   备注:物理降温时记录原值
    '   符号:用来标注体温不升，或者值小于等于项目最小值大于等于项目最大值是的特殊符号.此外默认为空
    If Not (mrsPoint Is Nothing) Then If mrsPoint.State = 1 Then mrsPoint.Close
    
    gstrFields = "序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|部位," & adLongVarChar & ",200|" & _
                 "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|" & _
                 "状态," & adDouble & ",1|复查," & adDouble & ",1|断开," & adDouble & ",1|重叠项目," & adLongVarChar & ",50|" & _
                 "重叠," & adDouble & ",5|X坐标," & adDouble & ",5|Y坐标," & adDouble & ",5|备注," & adLongVarChar & ",50|" & _
                 "符号," & adLongVarChar & ",10|显示," & adDouble & ",1"
    Call Record_Init(mrsPoint, gstrFields)
    
    '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,13-出生,99-未记说明)
    '禁用表示信息是否输出
    
    If Not mrsNote Is Nothing Then If mrsNote.State = 1 Then mrsNote.Close
    
    gstrFields = "时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|类型," & adDouble & ",2|" & _
        "内容," & adLongVarChar & ",200|颜色," & adLongVarChar & ",20|X坐标," & adDouble & ",20|" & _
        "Y坐标," & adDouble & ",20|高度," & adDouble & ",20|打印X坐标," & adDouble & ",20|" & _
        "禁用," & adInteger & ",1|显示," & adDouble & ",1"
        
    Call Record_Init(mrsNote, gstrFields)
    
    '加载体温数据
    Call SaveMemory(blnAutoAdjust)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SaveMemory(ByVal blnAutoAdjust As Boolean)
    Dim bytShow As Byte
    Dim strPart As String
    Dim strValue As String  '物理降温的原值
    Dim blnAdd As Boolean, bln降温 As Boolean
    Dim SinX As Single, sinY As Single
    Dim str时间 As String, str内容 As String
    Dim rsData As New ADODB.Recordset
    Dim rsPart As New ADODB.Recordset
    Dim strSql As String
    Dim lngColor As Long, lng行号 As Long, lng项目序号  As Long
    Dim str符号 As String
    Dim dbl数值 As Double, dblMinValue As Double, dblMaxValue As Double
    Dim strTmpString0 As String, strTmpString1 As String, strTmpString2 As String
    Dim strTime As String
    Dim blnAllow As Boolean
    Dim arrValues() As String
    Dim arrTmpValue() As Variant, arrTmpNote As Variant
    Dim i As Integer
    Dim int显示 As Integer
    Dim rs脉搏 As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim bln婴儿体温单显示出院 As Boolean, bln入科显示入院 As Boolean
    Dim lng体温不升显示方式 As Long
    Dim int标记 As Integer
    On Error GoTo Errhand
    
    '记录脉搏信息
    strFileds = "项目序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|X坐标," & adDouble & ",5|时间," & adLongVarChar & ",20"
    Call Record_Init(rs脉搏, strFileds)
    
    '提取所有部位信息
    strSql = "Select 项目序号, 部位,缺省项 From 体温部位"
    Call zldatabase.OpenRecordset(rsPart, strSql, "提取体温部位")
    
    
    '体温曲线项目需要增加字段用来决定该数据是否在三测单上显示,目前缺省为显示
    '-----------------------------------------------------------------------
    gstrSQL = "SELECT C.ID 序号,a.发生时间 As 时间,C.显示,C.记录内容 As 数值,C.体温部位,C.复试合格,D.记录名,E.保留项目,D.项目序号,DECODE(D.项目序号,-1,1,C.记录标记) 记录标记,C.未记说明 " & _
                "FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E " & _
                "Where B.ID=A.文件ID " & _
                    "And A.ID = C.记录ID " & _
                    "AND B.ID=[1] " & _
                    "AND Nvl(B.婴儿,0)=[6] " & _
                    "AND B.病人id=[2] " & _
                    "AND B.主页id=[3] " & _
                    "AND D.项目序号=c.项目序号 " & _
                    "AND c.记录类型=1 " & _
                    "AND E.项目序号=D.项目序号 " & _
                    "AND E.护理等级>=[7]  " & _
                    "AND a.发生时间 BETWEEN [4] And [5] And c.终止版本 Is Null " & _
                    "AND D.记录法=1 " & _
                "Order By A.发生时间,DECODE(C.项目序号,-1,1,0),DECODE(D.项目序号,-1,1,C.记录标记)"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "读取曲线项目数据", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, CDate(mstr开始时间), CDate(mstr结束时间), T_Patient.lng婴儿, T_Patient.lng护理等级)
        
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
        gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
        gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
    End If
    
    strTmpString0 = ""
    strTmpString1 = ""
    strTmpString2 = ""
    With rsData
        Do While Not .EOF
            str符号 = ""
            blnAllow = False
            strPart = zlCommFun.Nvl(!体温部位)
            lng项目序号 = Val(zlCommFun.Nvl(!项目序号))
            Select Case lng项目序号
                Case gint心率
                    int标记 = 1
                Case Else
                    int标记 = Val(Nvl(!记录标记))
            End Select
            If strPart = "" Then
                rsPart.Filter = "项目序号=" & lng项目序号 & " and 缺省项=1"
                If rsPart.BOF = False Then
                    strPart = zlCommFun.Nvl(rsPart!部位)
                Else
                    Select Case lng项目序号
                        Case gint体温
                            strPart = "腋温"
                        Case gint呼吸
                            strPart = "自主呼吸"
                        Case Else
                            strPart = ""
                    End Select
                End If
            End If
            
            mrsItems.Filter = "项目序号=" & lng项目序号
            If mrsItems.RecordCount > 0 Then
                SinX = GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"))
                strTime = GetXCoordinate(SinX, Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"), False)
                SinX = GetXCoordinate(Format(Split(strTime, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"))
                
                '记录所有脉搏信息
                If lng项目序号 = gint脉搏 Then
                    strFileds = "项目序号|数值|X坐标|时间"
                    strValues = lng项目序号 & "|" & zlCommFun.Nvl(!数值) & "|" & SinX & "|" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
                    Call Record_Add(rs脉搏, strFileds, strValues)
                End If
                
                If (Not IsNull(!未记说明)) And zlCommFun.Nvl(!数值) <> "不升" Then
                    mrsNote.Filter = "项目序号=" & Val(zlCommFun.Nvl(!项目序号)) & " AND X坐标=" & SinX
                    blnAdd = (mrsNote.RecordCount = 0)
                    '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,99-未记说明)
                    gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用|显示"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
                    gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & !项目序号 & "|99|" & _
                        !未记说明 & "|" & RGB_BLUE & "|" & SinX & "|0|0|0|0|" & Val(zlCommFun.Nvl(!显示))
                   
                    If blnAdd Then
                        '提取接近中间时间点的值做为本列值
                         Call Record_Add(mrsNote, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(mrsNote!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(mrsNote!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                             blnAllow = GetCanvasCenter(CDate(Format(mrsNote!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                        '提取接近中间时间点的值做为本列值
                        If blnAllow = True Then
                            If Val(mrsNote!显示) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(mrsNote, gstrFields, gstrValues, "时间|" & Format(mrsNote!时间, "yyyy-MM-dd HH:mm:ss"))
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                gstrFields = "显示"
                                gstrValues = "2"
                                Call Record_Update(mrsNote, gstrFields, gstrValues, "时间|" & Format(mrsNote!时间, "yyyy-MM-dd HH:mm:ss"))
                            End If
                        End If
                        
                    End If
                Else
                    blnAdd = False
                    
                    mrsPoint.Filter = "项目序号=" & Val(zlCommFun.Nvl(!项目序号, 0)) & " AND X坐标=" & SinX & " And 标记=" & int标记
                    blnAdd = (mrsPoint.RecordCount = 0)
                    
                    dbl数值 = Val(zlCommFun.Nvl(!数值))
                    
                    dblMinValue = GetMinValue(!项目序号)
                    dblMaxValue = GetMaxValue(!项目序号)
                    
                    '不指定符号，项目数值操作最大值和最小值以项目本身符号显示
                    If dbl数值 <= dblMinValue Then
                        dbl数值 = dblMinValue
                        'str符号 = "·"
                    End If
                    
                    If dbl数值 >= dblMaxValue Then
                        dbl数值 = dblMaxValue
                        'str符号 = "·"
                    End If
                    
                     '体温不升是在显示在35刻度
                    If Trim(Nvl(!数值)) = "不升" And lng项目序号 = gint体温 Then dbl数值 = 35
                    sinY = Val(GetYCoordinate(mobjDraw, mrsDrawItems, !项目序号, dbl数值, mlngMemDC, True))
                     
                    gstrFields = "序号|数值|部位|标记|时间|项目序号|状态|复查|断开|重叠项目|重叠|X坐标|Y坐标|备注|符号|显示"
                    gstrValues = Val(zlCommFun.Nvl(!序号)) & "|" & !数值 & "|" & strPart & "|" & int标记 & "|" & _
                                 Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng项目序号 & "|0|" & Val(zlCommFun.Nvl(!复试合格, 0)) & "|" & IIf(zlCommFun.Nvl(!数值, 0) = "不升", 1, 0) & "|空|0|" & _
                                 SinX & "|" & sinY & "||" & str符号 & "|" & Val(zlCommFun.Nvl(!显示, 0))
                    If blnAdd Then '添加
                        Call Record_Add(mrsPoint, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(mrsPoint!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(mrsPoint!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                            blnAllow = GetCanvasCenter(CDate(Format(mrsPoint!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
                        ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                            blnAllow = True
                        End If
                        
                        '提取接近中间时间点的值做为本列值
                        If blnAllow = True Then
                            If Val(mrsPoint!显示) = 2 Then
                                arrValues = Split(gstrValues, "|")
                                arrValues(UBound(arrValues)) = 2
                                gstrValues = Join(arrValues, "|")
                            End If
                            Call Record_Update(mrsPoint, gstrFields, gstrValues, "序号|" & mrsPoint!序号)
                        Else
                            If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                                gstrFields = "显示"
                                gstrValues = "2"
                                Call Record_Update(mrsPoint, gstrFields, gstrValues, "序号|" & mrsPoint!序号)
                            End If
                        End If
                    End If
                End If
            End If
            mrsItems.Filter = 0
        .MoveNext
        Loop
    End With
    
     '上面已经得到了所有项目的数据信息，下来处理物理降温和脉搏和心率数据
    arrTmpValue = Array()
    If mint心率应用 = 2 Then
        mrsPoint.Filter = "项目序号=" & gint心率
        With mrsPoint
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !X坐标 & ";" & Format(!时间, "yyyy-MM-DD HH:mm:ss")
            .MoveNext
            Loop
        End With
    End If
    mrsPoint.Filter = ""
    
    '心率设为脉搏共用时，检查脉搏是否设置为可用
    mrsItems.Filter = "项目序号=" & gint脉搏
    If mrsItems.RecordCount > 0 Then
        For i = 0 To UBound(arrTmpValue)
            '检查心率是否与脉搏相对应
            rs脉搏.Filter = "项目序号=" & gint脉搏 & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
            mrsPoint.Filter = "项目序号=" & gint脉搏 & " and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
            If mrsPoint.RecordCount = 0 Then
                If rs脉搏.RecordCount = 0 Then
                    mrsPoint.Filter = ""
                    gstrFields = "项目序号": gstrValues = gint脉搏
                    Call Record_Update(mrsPoint, gstrFields, gstrValues, "序号|" & Val(Split(CStr(arrTmpValue(i)), ";")(0)))
                Else
                    mrsPoint.Filter = "项目序号=" & gint心率 & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(2)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(3), "yyyy-MM-DD HH:mm:ss") & "'"
                    mrsPoint.Delete
                End If
            End If
        Next i
    End If
    
    If mint心率应用 = 2 Then
        Set rs脉搏 = New ADODB.Recordset
        strFileds = "序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|部位," & adLongVarChar & ",200|" & _
                    "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|" & _
                    "状态," & adDouble & ",1|复查," & adDouble & ",1|断开," & adDouble & ",1|重叠项目," & adLongVarChar & ",50|" & _
                    "重叠," & adDouble & ",5|X坐标," & adDouble & ",5|Y坐标," & adDouble & ",5|备注," & adLongVarChar & ",50|" & _
                    "符号," & adLongVarChar & ",10|显示," & adDouble & ",1"
        Call Record_Init(rs脉搏, strFileds)
        
        mrsPoint.Filter = "项目序号=" & gint脉搏
        With mrsPoint
            Do While Not .EOF
                rs脉搏.AddNew
                For i = 0 To .Fields.Count - 1
                    rs脉搏.Fields(.Fields(i).Name).Value = .Fields(i).Value
                Next i
                rs脉搏.Update
            .MoveNext
            Loop
        End With
        
        mrsPoint.Filter = "项目序号=" & gint脉搏
        Do While Not mrsPoint.EOF
            mrsPoint.Delete
            mrsPoint.MoveNext
        Loop
        
        rs脉搏.Filter = ""
        rs脉搏.Sort = "时间"
        With rs脉搏
            Do While Not .EOF
                blnAdd = False
                blnAllow = False
                
                SinX = Val(zlCommFun.Nvl(!X坐标))
                sinY = Val(zlCommFun.Nvl(!Y坐标))
                mrsPoint.Filter = "项目序号=" & Val(zlCommFun.Nvl(!项目序号, 0)) & " AND X坐标=" & SinX
                blnAdd = IIf(mrsPoint.RecordCount = 0, True, False)
                
                strFileds = "序号|数值|部位|标记|时间|项目序号|状态|复查|断开|重叠项目|重叠|X坐标|Y坐标|备注|符号|显示"
                strValues = Val(zlCommFun.Nvl(!序号)) & "|" & !数值 & "|" & zlCommFun.Nvl(!部位) & "|" & Val(zlCommFun.Nvl(!标记, 0)) & "|" & _
                             Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!项目序号)) & "|0|0|" & Val(zlCommFun.Nvl(!断开)) & "|空|0|" & _
                             SinX & "|" & sinY & "||" & zlCommFun.Nvl(!符号) & "|" & Val(zlCommFun.Nvl(!显示, 0))
                
                If blnAdd Then '添加
                    Call Record_Add(mrsPoint, strFileds, strValues)
                Else
                    If (zlCommFun.Nvl(mrsPoint!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(mrsPoint!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                        blnAllow = GetCanvasCenter(CDate(Format(mrsPoint!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
                    ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                        blnAllow = True
                    End If
                    
                    '提取接近中间时间点的值做为本列值
                    If blnAllow = True Then
                        If Val(mrsPoint!显示) = 2 Then
                            arrValues = Split(strValues, "|")
                            arrValues(UBound(arrValues)) = 2
                            strValues = Join(arrValues, "|")
                        End If
                        Call Record_Update(mrsPoint, strFileds, strValues, "序号|" & mrsPoint!序号)
                    Else
                        If Val(zlCommFun.Nvl(!显示, 0)) = 2 Then
                            strFileds = "显示"
                            strValues = "2"
                            Call Record_Update(mrsPoint, strFileds, strValues, "序号|" & mrsPoint!序号)
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
    End If
    
    '处理物理降温数据
    arrTmpValue = Array()
    mrsPoint.Filter = "项目序号=1 and 标记=0"
    With mrsPoint
        Do While Not .EOF
            ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
            arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !数值 & ";" & !X坐标 & ";" & !Y坐标 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
        .MoveNext
        Loop
    End With
    
    mrsPoint.Filter = "项目序号=1"
    If mrsPoint.RecordCount > 0 Then mrsPoint.MoveFirst
    For i = 0 To UBound(arrTmpValue)
        mrsPoint.Filter = "项目序号=1 and 标记=1 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
        If mrsPoint.RecordCount <> 0 Then
            gstrFields = "备注": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
            Call Record_Update(mrsPoint, gstrFields, gstrValues, "序号|" & zlCommFun.Nvl(mrsPoint!序号))
        End If
    Next i
    
    arrTmpValue = Array()
    mrsPoint.Filter = "项目序号=1 and 标记=1"
    With mrsPoint
        Do While Not .EOF
            ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
            arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !数值 & ";" & !X坐标 & ";" & !Y坐标 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
        .MoveNext
        Loop
    End With
    
    mrsPoint.Filter = "项目序号=1"
    If mrsPoint.RecordCount > 0 Then mrsPoint.MoveFirst
    For i = 0 To UBound(arrTmpValue)
        mrsPoint.Filter = "项目序号=1 and 标记=0 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
        If mrsPoint.RecordCount = 0 Then
            mrsPoint.Filter = "项目序号=1 and 标记=1 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            mrsPoint.Delete
        End If
    Next i
    
    
    '删除显示为2的数据
    mrsPoint.Filter = "显示=2"
    Do While Not mrsPoint.EOF
        mrsPoint.Delete
    mrsPoint.MoveNext
    Loop
    
    mrsNote.Filter = ""
    mrsNote.Filter = "显示=2"
    Do While Not mrsNote.EOF
        mrsNote.Delete
    mrsNote.MoveNext
    Loop

    '处理未记说明和曲线数据该显示那一条
    mrsNote.Filter = ""
    mrsPoint.Filter = ""
    
    arrTmpValue = Array()
    arrTmpNote = Array()
    mrsNote.Sort = "项目序号,X坐标"
    With mrsNote
        Do While Not .EOF
            blnAllow = False
            SinX = Val(!X坐标)
            mrsPoint.Filter = "项目序号=" & Val(!项目序号) & " And X坐标=" & SinX
            If mrsPoint.RecordCount > 0 Then
                If (zlCommFun.Nvl(mrsPoint!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(mrsPoint!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                    blnAllow = GetCanvasCenter(CDate(Format(mrsPoint!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
                ElseIf zlCommFun.Nvl(!显示, 0) = 1 Then
                    blnAllow = True
                End If
                If blnAllow = True Then
                    ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                    arrTmpValue(UBound(arrTmpValue)) = !项目序号 & ";" & SinX
                Else
                    ReDim Preserve arrTmpNote(UBound(arrTmpNote) + 1)
                    arrTmpNote(UBound(arrTmpNote)) = !项目序号 & ";" & SinX
                End If
            End If
        .MoveNext
        Loop
    End With
    
    For i = 0 To UBound(arrTmpValue)
        mrsPoint.Filter = "项目序号=" & Val(Split(CStr(arrTmpValue(i)), ";")(0)) & " And X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(1))
        Do While Not mrsPoint.EOF
            mrsPoint.Delete
        mrsPoint.MoveNext
        Loop
    Next i
    
    For i = 0 To UBound(arrTmpNote)
        mrsNote.Filter = "项目序号=" & Val(Split(CStr(arrTmpNote(i)), ";")(0)) & " And X坐标=" & Val(Split(CStr(arrTmpNote(i)), ";")(1))
        Do While Not mrsNote.EOF
            mrsNote.Delete
        mrsNote.MoveNext
        Loop
    Next i
    
    '处理体温不升 体温为不升需要在35度下纵向输出体温不升二字
    mrsPoint.Filter = "项目序号=" & gint体温 & " and 数值='不升' and 标记<>1"
    mrsPoint.Sort = "时间"
    With mrsPoint
        Do While Not .EOF
            strTmpString0 = strTmpString0 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & Val(zlCommFun.Nvl(!项目序号)) & "|99|" & _
                  "不升|" & RGB_BLUE & "|" & !X坐标 & "|0|0|0|0"
            strTmpString2 = strTmpString2 & ";" & !X坐标
        .MoveNext
        Loop
    End With
    '读取手术、上下标信息
    '-----------------------------------------------------------------------
    gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
    gstrSQL = "" & _
             " Select A.发生时间 AS 时间,C.记录类型,C.项目序号,C.记录内容,C.项目名称,C.未记说明" & _
             " FROM 病人护理文件 B, 病人护理数据 A, 病人护理明细 C" & _
             " Where B.ID=A.文件ID And A.ID = C.记录ID AND B.ID=[1] AND Nvl(B.婴儿, 0)=[6] AND B.病人id=[2] AND B.主页id=[3] And c.终止版本 Is Null" & _
             " AND MOD(C.记录类型,10) <> 1  AND A.发生时间 BETWEEN [4]  And [5]"
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "读取手术、上下标等信息", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, Int(CDate(mstr开始时间)), CDate(mstr结束时间), T_Patient.lng婴儿, T_Patient.lng护理等级)
    
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
        gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
        gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
    End If
        
    With rsData
        Do While Not .EOF
            bytShow = 1
            str内容 = Trim(zlCommFun.Nvl(!记录内容))
            
            lng行号 = IIf(!记录类型 = 2, 10, IIf(!记录类型 = 6, 11, 4))
            
            '对于手术显示需要特殊处理
            If !记录类型 = 4 Then
                str内容 = Trim(zlCommFun.Nvl(!项目名称))
                
                If str内容 = "分娩" Then
                    bytShow = T_BodyFlag.分娩
                Else
                    bytShow = T_BodyFlag.手术
                End If
                
                If bytShow = 2 And Not blnAutoAdjust Then
                    str内容 = str内容 & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                Else
                    str内容 = !项目名称
                End If
                lngColor = RGB_RED
            Else
                lngColor = IIf(Not IsNumeric(Nvl(!未记说明)), RGB_BLUE, Val(Nvl(!未记说明)))
            End If
            
            If bytShow > 0 Then
                SinX = Val(GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), mstr开始时间))
                
                mrsNote.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=" & !记录类型
                If mrsNote.BOF Then
                    gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|" & !记录类型 & "|" & _
                        str内容 & "|" & lngColor & "|" & SinX & "|0|0|0|0"
                    Call Record_Add(mrsNote, gstrFields, gstrValues)
                Else
                    mrsNote!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                    mrsNote!内容 = str内容
                    mrsNote.Update
                End If
            End If
            mrsNote.Filter = 0
            .MoveNext
        Loop
    End With
    
    bln婴儿体温单显示出院 = (zldatabase.GetPara("婴儿体温单显示出院信息", glngSys, 1255, 1) = 1)
    
    bln入科显示入院 = False
    If CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrComeInDate, "yyyy-MM-dd HH:mm:ss")) Then
        bln入科显示入院 = True
    ElseIf CDate(Format(mstrEnterDate, "YYYY-MM-DD HH:mm:ss")) = CDate(Format(mstrComeInDate, "yyyy-MM-dd HH:mm:ss")) And T_BodyFlag.入院 = 0 Then
        bln入科显示入院 = True
    End If
    
    '读取入出转等信息
    '-----------------------------------------------------------------------
    '所有需要输出的文本内容(类型:2-上标;3-入出转;4-手术日;6-下标,99-未记说明)
    '1-入院；2-入科；3-转科；4-换床
    gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
    Set rsData = GetDataFromHis(T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿, CDate(mstr开始时间), CDate(mstr结束时间), 2) ' Int(cdate(mstr开始时间))
    With rsData
        Do While Not .EOF
            If Trim(zlCommFun.Nvl(!内容)) <> "" Then
                bytShow = 0
                lng行号 = Val(!行号)
                str内容 = zlCommFun.Nvl(!内容)
                Select Case lng行号
                Case 5
                    bytShow = T_BodyFlag.入院
                Case 6, 3 '6转入，3转出
                    bytShow = T_BodyFlag.转出
                Case 7
                    bytShow = T_BodyFlag.换床
                Case 8
                    bytShow = T_BodyFlag.出院
                    If T_Patient.lng婴儿 > 0 Then
                        bytShow = IIf(bln婴儿体温单显示出院, bytShow, 0)
                    End If
                Case 9
                    bytShow = T_BodyFlag.入科
                End Select
                 
                If bytShow > 0 Then
                    '目前3，4 针对于转科 3-显示说明和科室 4 显示说明，科室，时间
                    If lng行号 = 9 And bln入科显示入院 = True Then
                        str内容 = "入院"
                    End If
                
                    If bytShow = 2 Then
                        str内容 = str内容 & IIf(blnAutoAdjust = False, gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm")), "")
                    ElseIf bytShow = 3 Then
                        str内容 = str内容 & IIf(blnAutoAdjust = False, gstrCaveSplit & zlCommFun.Nvl(!科室), "")
                    ElseIf bytShow = 4 Then
                        str内容 = str内容 & IIf(blnAutoAdjust = False, gstrCaveSplit & zlCommFun.Nvl(!科室) & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm")), "")
                    ElseIf bytShow = 1 Then
                        str内容 = str内容
                    End If
                    
                    '体温单缩略模式下 换床不显示床号
                    If bytShow = T_BodyFlag.换床 And blnAutoAdjust = True Then
                        If InStr(1, str内容, "(") <> 0 Then
                            str内容 = Split(str内容, "(")(0)
                        End If
                    End If
                    
                    SinX = Val(GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), mstr开始时间))
                    mrsNote.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=3"
                    
                    If mrsNote.BOF Then
                        gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|3|" & _
                            str内容 & "|" & RGB_RED & "|" & SinX & "|0|0|0|0"
                        Call Record_Add(mrsNote, gstrFields, gstrValues)
                    Else
                        mrsNote!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                        mrsNote!内容 = str内容
                        mrsNote.Update
                    End If
                End If
                mrsNote.Filter = 0
            End If
            .MoveNext
        Loop
    End With
    
    '提取婴儿出生信息
    If T_Patient.lng婴儿 > 0 Then
        gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
        Set rsData = GetDataFromHis(T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿, CDate(mstr开始时间), CDate(mstr结束时间), 3)
        With rsData
            Do While Not .EOF
                bytShow = 0
                If Trim(zlCommFun.Nvl(!内容)) <> "" Then
                    lng行号 = 12
                    bytShow = T_BodyFlag.出生
                    If bytShow > 0 Then
                        Select Case bytShow
                            Case 1
                                str内容 = zlCommFun.Nvl(!内容)
                            Case 2
                                If Not blnAutoAdjust Then
                                    str内容 = zlCommFun.Nvl(!内容) & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                                Else
                                    str内容 = zlCommFun.Nvl(!内容)
                                End If
                        End Select
                        
                        SinX = Val(GetXCoordinate(Format(!时间, "YYYY-MM-DD HH:mm:ss"), mstr开始时间))
                        mrsNote.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=13"
                        
                        If mrsNote.BOF Then
                            gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|13|" & _
                                str内容 & "|" & RGB_RED & "|" & SinX & "|0|0|0|0"
                            Call Record_Add(mrsNote, gstrFields, gstrValues)
                        Else
                            mrsNote!时间 = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                            mrsNote!内容 = str内容
                            mrsNote.Update
                        End If
                    End If
                End If
                mrsNote.Filter = 0
            .MoveNext
            Loop
        End With
    End If
    
    str符号 = ""
    Dim bytTag As Byte
    '51512,刘鹏飞,2012-07-11,未记说明显示位置 0-显示在上面,1-显示在下面,2-不显示
    '大医二院要求未记说明不显示，但标注了未记的两边的体温曲线不连接
    bytTag = Abs(Val(zldatabase.GetPara("未记说明显示位置", glngSys, 1255, "0")))
    lng体温不升显示方式 = Val(zldatabase.GetPara("体温不升显示方式", glngSys, 1255, "0"))
    '处理体温不升 体温不升始终显示在 35 度下面，只有未记说明显示在下面的情况，才将不升放入未记说明中，其它情况都放在下标中
    If Left(strTmpString0, 1) = ";" Then
        gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"
        If lng体温不升显示方式 = 0 Or lng体温不升显示方式 = 2 Then
            arrValues = Split(strTmpString0, "|")
            arrValues(3) = "↓ "
            strTmpString0 = Join(arrValues, "|")
        End If
        strTmpString0 = Mid(strTmpString0, "2")
        strTmpString2 = Mid(strTmpString2, 2)
        For i = 0 To UBound(Split(strTmpString0, ";"))
            str符号 = Split(strTmpString0, ";")(i)
            mrsNote.Filter = "类型=" & IIf(bytTag = 1, 99, 6) & " and X坐标=" & Val(Split(strTmpString2, ";")(i))
            mrsNote.Sort = "项目序号"
            If mrsNote.RecordCount > 0 Then
                mrsNote!内容 = IIf(lng体温不升显示方式 = 0 Or lng体温不升显示方式 = 2, "↓ ", "不升") & IIf(bytTag = 1, ";", " ") & zlCommFun.Nvl(mrsNote!内容)
                mrsNote.Update
            Else
                If lng体温不升显示方式 = 0 Or lng体温不升显示方式 = 2 Then
                    str符号 = Replace(str符号, "不升", "↓ ")
                End If
                Call Record_Add(mrsNote, gstrFields, str符号)
                mrsNote!类型 = IIf(bytTag = 1, 99, 6)
                mrsNote.Update
            End If
        Next i
    End If
    
    '更新断开标志(时间超过一天活存在未记说明均不连线)
    Call ProcessPoint(mstr开始时间)
    '计算组织重复的点
    Call GetConverPoint(mrsPoint)
    
    '如果未记说明不显示，将取消记录集mrsNote中类型为99的记录
    If bytTag = 2 Then
        mrsNote.Filter = "类型=99"
        Do While Not mrsNote.EOF
            mrsNote.Delete
            mrsNote.MoveNext
        Loop
        mrsNote.Filter = ""
    End If
    '在立即窗口中输出点信息
    'Call OutputRsData(mrsPoint, True)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PrepareGraph()
    Dim strPic As String, strOverlap As String, strPart As String, lngID As Long    '图片,重叠序号,部位,主项目序号
    Dim lngCurX As Long, sinCurY As Single, lngCount As Long, lngMax As Long        '一行能保存多少个图片?
    Dim rsTemp As New ADODB.Recordset
    Dim rsOverlap As New ADODB.Recordset
    Dim ArrCode() As String, arrChar() As String, arrItem() As String
    Dim strChar As String
    Dim i As Integer
    On Error GoTo Errhand
    
   If Not mrsGraph Is Nothing Then If mrsGraph.State = 1 Then mrsGraph.Close
    
    lngMax = mobjBuffer.ScaleWidth \ gintBmpW      '一行能保存多少个图片?
    '所有需输出的图形序号(包括体温重叠标记),全部提取在picBuffer中,此处保存各项目的部位及其对应的图形序号
    gstrFields = "项目序号," & adDouble & ",18|部位," & adLongVarChar & ",50|记录符," & adLongVarChar & ",50|" & _
                 "记录色," & adDouble & ",18|重叠项目," & adLongVarChar & ",20|行," & adDouble & ",5|列," & adDouble & ",5"    '重叠项目应按项目序号大小排列,如:1,4,5
    Call Record_Init(mrsGraph, gstrFields)
    
    '先根据体温部位装载
    gstrSQL = " Select 项目序号,'' AS 部位, 记录符 标记符号,记录色 标记颜色,1 展现方式,'空' AS 重叠项目 From 体温记录项目 Order by 项目序号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取各曲线项目的展现方式")
    With rsTemp
        Do While Not .EOF
            If !展现方式 = 1 Then
                
                If !项目序号 = 1 Then '体温
                    ArrCode = Split("口温,腋温,肛温", ",")
                    strChar = zlCommFun.Nvl(!标记符号, "·,×,○")
                    strChar = strChar & String(2 - UBound(Split(strChar, ",")), ",")
                    arrChar = Split(strChar, ",")
                    For i = 0 To UBound(ArrCode)
                        gstrFields = "项目序号|部位|重叠项目|记录符|记录色"
                        gstrValues = !项目序号 & "|" & ArrCode(i) & "|" & zlCommFun.Nvl(!重叠项目) & "|" & arrChar(i) & "|" & zlCommFun.Nvl(!标记颜色, 0)
                        Call Record_Add(mrsGraph, gstrFields, gstrValues)
                    Next i
                Else
                    strPart = ""
                    strChar = zlCommFun.Nvl(!标记符号)
                    '产生相应的内存记录数据
                    gstrFields = "项目序号|部位|重叠项目|记录符|记录色"
                    gstrValues = !项目序号 & "|" & strPart & "|" & zlCommFun.Nvl(!重叠项目) & "|" & strChar & "|" & zlCommFun.Nvl(!标记颜色, 0)
                    Call Record_Add(mrsGraph, gstrFields, gstrValues)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    '添加起搏器和呼吸机图形
    arrItem = Split("2,3", ",")
    ArrCode = Split("起搏器,呼吸机", ",")
    arrChar = Split("PACEMAKER,BREATH", ",")
    For i = 0 To UBound(ArrCode)
        strPic = arrChar(i)  '资源文件
        If strPic <> "" Then
            If DrawPicture(mobjBuffer, strPic, lngCurX, sinCurY, lngCurX + gintBmpW, sinCurY + gintBmpH, True) Then
                '产生相应的内存记录数据
                gstrFields = "项目序号|部位|重叠项目|记录符|行|列"
                gstrValues = Val(arrItem(i)) & "|" & ArrCode(i) & "|" & "空" & "||" & sinCurY \ gintBmpH & "|" & lngCount
                Call Record_Add(mrsGraph, gstrFields, gstrValues)
                
                '位移计算
                lngCurX = lngCurX + gintBmpW
                lngCount = lngCount + 1
                If lngCount >= lngMax Then
                    lngCount = 0
                    lngCurX = 0
                    sinCurY = sinCurY + gintBmpH
                End If
            End If
            'If !展现方式 = 2 Then Call FileSystem.Kill(strPic)
        End If
    Next i

    '再根据体温重叠标记装载
    gstrSQL = " Select 序号,标记符号,标记颜色 From 体温重叠标记 Where nvl(重叠数目,0)>0 Order by 序号"
    Set rsOverlap = zldatabase.OpenSQLRecord(gstrSQL, "再根据体温重叠标记装载")
    gstrSQL = " Select 序号,上级序号,项目序号,体温部位 From 体温重叠标记 Where 项目序号 is not null Order by 序号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取重叠从属项目")
    Do While Not rsOverlap.EOF
        strPart = ""
        strOverlap = ""
        With rsTemp
            .Filter = "上级序号=" & rsOverlap!序号
            If rsTemp.RecordCount > 0 Then
                '第一条记录则为主记录
                lngID = zlCommFun.Nvl(!项目序号, 0)
                strPart = rsOverlap!序号 ' zlCommFun.Nvl(!体温部位)
                 
                .Sort = "项目序号"
                Do While Not .EOF
                    If !项目序号 <> lngID Then
                        strOverlap = strOverlap & "," & !项目序号
                    End If
                    .MoveNext
                Loop
                .Sort = "序号"              '此处需要按序号还原,否则取下一个重叠项目的主项目时会取错
                
                strOverlap = Mid(strOverlap, 2)
                If Not IsNull(rsOverlap!标记符号) Then
                    '输出字符
                    '产生相应的内存记录数据
                    gstrFields = "项目序号|部位|重叠项目|记录符|记录色"
                    gstrValues = lngID & "|" & strPart & "|" & strOverlap & "|" & zlCommFun.Nvl(rsOverlap!标记符号) & "|" & rsOverlap!标记颜色
                    Call Record_Add(mrsGraph, gstrFields, gstrValues)
                Else
                    '输出图形文件
                    strPic = zlBlobRead(9, rsOverlap!序号)
                    If strPic <> "" Then
                        If DrawPicture(mobjBuffer, strPic, lngCurX, sinCurY, lngCurX + gintBmpW, sinCurY + gintBmpH, False) Then
                            '产生相应的内存记录数据
                            gstrFields = "项目序号|部位|重叠项目|记录符|行|列"
                            gstrValues = lngID & "|" & strPart & "|" & strOverlap & "||" & sinCurY \ gintBmpH & "|" & lngCount
                            Call Record_Add(mrsGraph, gstrFields, gstrValues)
                            
                            '位移计算
                            lngCurX = lngCurX + gintBmpW
                            lngCount = lngCount + 1
                            If lngCount >= lngMax Then
                                lngCount = 0
                                lngCurX = 0
                                sinCurY = sinCurY + gintBmpH
                            End If
                        End If
                        Call FileSystem.Kill(strPic)
                    End If
                End If
            End If
        End With
        rsOverlap.MoveNext
    Loop
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ProcessPoint(ByVal strBeginDate As String)
    Dim arrData
    Dim lngOrder As Long
    Dim lngCurX As Long                         '记录未记说明所在坐标
    Dim strPrimary As String
    Dim strDate As String
    On Error GoTo Errhand
    '处理所有点的各种标志位，如断开
    
    '先处理未记说明，把未记说明前一个点的断开标志设置为1
    '---------------------------------------------------
    strPrimary = "序号|"        '格式:字段名,值
    gstrFields = "断开"         '格式:字段名|字段名
    gstrValues = 1              '格式:值|值
    
    mrsPoint.Filter = ""

    With mrsNote
        .Filter = "类型=99"
        Do While Not .EOF
            lngCurX = GetXCoordinate(!时间, strBeginDate)
            If mint心率应用 = 2 And !项目序号 = -1 Then
                mrsPoint.Filter = "项目序号=" & gint脉搏 & " And  X坐标<=" & !X坐标
            Else
                If Val(!项目序号) = 1 Then
                    mrsPoint.Filter = "项目序号=" & !项目序号 & " And  标记<>1 And X坐标<" & !X坐标
                Else
                    mrsPoint.Filter = "项目序号=" & !项目序号 & " And X坐标<" & !X坐标
                End If
            End If
            
            mrsPoint.Sort = "时间"
            If mrsPoint.RecordCount <> 0 Then
                mrsPoint.MoveLast
                lngOrder = mrsPoint!序号
                
                Call Record_Update(mrsPoint, gstrFields, gstrValues, strPrimary & lngOrder)
            End If
            mrsPoint.Filter = 0
            
            .MoveNext
        Loop
        
        .Filter = 0
    End With
        
    '有一天未测数据的也设置断开标志
    '---------------------------------------------------
    lngCurX = 0
    lngOrder = 0
    strPrimary = ""
    mrsPoint.Filter = ""
    mrsPoint.Sort = "项目序号,时间,标记"
    
    With mrsPoint
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If Not (Val(!项目序号) = gint体温 And Val(zlCommFun.Nvl(!标记)) = 1) Then
                If lngCurX <> 0 Then
                    If lngCurX <> !项目序号 Then strDate = ""
                End If
                lngCurX = !项目序号
                
                If strDate <> "" Then
                    If DateDiff("d", CDate(strDate), CDate(Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss"))) > 1 Then
                        strPrimary = strPrimary & "," & lngOrder
                    End If
                End If
                '记录当前点的相关信息,供下一个点时检查
                strDate = Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss")
                lngOrder = Val(zlCommFun.Nvl(!序号))
            End If
            .MoveNext
        Loop
    End With
    
    arrData = Split(strPrimary, ",")
    lngOrder = UBound(arrData)
    strPrimary = "序号|"        '格式:字段名,值
    
    For lngCurX = 1 To lngOrder
        Call Record_Update(mrsPoint, gstrFields, gstrValues, strPrimary & arrData(lngCurX))
    Next
    
    '处理体温不升的.把前一个点的断开标志设置为1
    mrsPoint.Filter = ""
    mrsPoint.Filter = "项目序号=" & gint体温 & " and 标记<>1"
    mrsPoint.Sort = "时间,标记"
    With mrsPoint
        Do While Not .EOF
            If !数值 = "不升" And .AbsolutePosition <> 1 Then
                .MovePrevious '更新上一行断开标记
                If Val(zlCommFun.Nvl(!断开)) <> 1 Then
                    lngOrder = !序号
                    Call Record_Update(mrsPoint, gstrFields, gstrValues, strPrimary & lngOrder)
                End If
                .MoveNext
            End If
        .MoveNext
        Loop
    End With
    mrsPoint.Filter = 0
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawMarker(ByVal bln作图区域 As Boolean, ByVal lng项目序号 As Long, ByVal str部位 As String, ByVal lngCurX As Long, ByVal sinCurY As Single, Optional ByVal str重叠项目 As String = "空", Optional ByVal bln项目 As Boolean = False, Optional ByVal str符号 As String = "")

    'bln作图区域=True,体温单作图区域,计算后居中显示;否则,照传入的数据显示
    Dim blnGraph As Boolean
    Dim bln重叠 As Boolean
    Dim str记录符 As String
    Dim lngRGB As Long

    On Error GoTo Errhand

    '输出字符或图形
  
    mrsGraph.Filter = "项目序号=" & lng项目序号 & " And 部位='" & str部位 & "' And 重叠项目='" & str重叠项目 & "'"

    If mrsGraph.RecordCount = 0 Then    '未设置重叠项目的输出方式,则按项目序号+部位输出
        mrsGraph.Filter = "项目序号=" & lng项目序号 & " And 部位='" & str部位 & "'"
    Else
        bln重叠 = True
    End If
    
    If mrsGraph.RecordCount = 0 Then    '未设置该项目按部位的输出方式,则按项目的设置输出
        mrsGraph.Filter = "项目序号=" & lng项目序号
    End If
    
    If mrsGraph.RecordCount = 0 Then Exit Sub
    blnGraph = (zlCommFun.Nvl(mrsGraph!记录符) = "")
    
    If Not blnGraph Then
        If bln重叠 = True And str重叠项目 <> "空" Then
            str记录符 = zlCommFun.Nvl(mrsGraph!记录符)
        Else

            If str符号 <> "" Then
                str记录符 = str符号
            Else
                str记录符 = zlCommFun.Nvl(mrsGraph!记录符)
            End If
        End If
        
        lngRGB = Val(mrsGraph!记录色)
        
        If lng项目序号 = -1 And mint心率应用 = 2 Then lngRGB = RGB_RED
        
        '字符输出
        Call SetTextColor(mlngMemDC, lngRGB)
        Call GetTextRect(mobjDraw, lngCurX - IIf(bln项目 = True, Screen.TwipsPerPixelY / 2, 0), sinCurY + IIf(bln项目 = True, Screen.TwipsPerPixelY / 2, 0), Trim(Split(str记录符 & ",", ",")(0)), IIf(bln作图区域, T_DrawClient.列单位, T_DrawClient.刻度单位))
        Call DrawText(mlngMemDC, Trim(Split(str记录符 & ",", ",")(0)), -1, T_LableRect, DT_CENTER)
    Else

        '输出体温项目的图形
        If bln作图区域 Then
            '体温作图区域居中打印
            Call BitBlt(mlngMemDC, lngCurX + 2, sinCurY - gintBmpH / 2, gintBmpW, gintBmpH, mobjBuffer.hDC, mrsGraph!列 * gintBmpW, mrsGraph!行 * gintBmpH, SRCCOPY)
        Else
            '刻度区域按指定坐标输出
            Call BitBlt(mlngMemDC, lngCurX, sinCurY, gintBmpW, gintBmpH, mobjBuffer.hDC, mrsGraph!列 * gintBmpW, mrsGraph!行 * gintBmpH, SRCCOPY)
        End If
    End If
    
    mrsGraph.Filter = ""

    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CurveCount() As Long
'--------------------------------------------------
'功能:得到体温曲线项目数据
'--------------------------------------------------
    Dim strSql As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    On Error GoTo Errhand
    
    strSql = " Select Count(*) 记录" & _
             " From 体温记录项目 A, 护理记录项目 C" & _
             " Where A.项目序号=C.项目序号 And A.记录法=1" & _
             " And nvl(C.应用方式,0)=1" & _
             " And nvl(C.适用病人,0) in (0,[1]) And  nvl(C.护理等级,3)>=[3] " & _
             " and (C.适用科室=1 OR (C.适用科室=2 and Exists (select 1 from 护理适用科室 D where C.项目序号=D.项目序号 and D.科室ID=[2])))" & _
             " Order by C.项目序号"
    Set rsTemp = zldatabase.OpenSQLRecord(strSql, "取开始行", IIf(T_Patient.lng婴儿 = 0, 1, 2), T_Patient.lng科室ID, T_Patient.lng护理等级)
    
    lngCount = Val(zlCommFun.Nvl(rsTemp!记录))
    
    CurveCount = lngCount
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PrintState(ByVal intPrintRange As Integer, ByVal blnPrint As Boolean, Optional lngBeginY As Long, _
    Optional ByVal intPageNo As Integer = -1, Optional ByVal strPrintDevice As String, Optional strPage As String, Optional strParam As String = "") As Boolean
    '******************************************************************************************************************
    '功能:将当前体温表或当前开始的所有体温表输出到打印机上或预览窗体
    '参数:blnCurState = 是否为只打印当前体温表,否则打印从当前开始的所有体温表
    '     blnPrint    = 是否输出到打印机上否则输出到预览窗体里
    '******************************************************************************************************************
    
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, strNewSql As String
    Dim strPaper As String
    Dim strPrintName As String
    Dim blnYesPrinter As Boolean
    Dim intCOl As Integer
    Dim intBeginPage As Integer
    Dim intEndPage As Integer
    Dim byeReturn As Byte
    Dim strArrFromTo() As String
    Dim intBaby As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim lngIndex As Long, lngIndexEnd As Long
    Dim intCount As Integer
    Dim objPrint As Object
    Dim strMarkDate As String
    Dim arrParam() As String
    
    On Error GoTo ErrHandle
    
    If strParam <> "" Then
        arrParam = Split(strParam, ";")
        If UBound(arrParam) < 2 Then
            MsgBox "strParam参数不为空时,必须传入文件ID;病人ID;主页ID！", vbInformation, gstrSysName
            Exit Function
        End If
        T_Patient.lng文件ID = Val(arrParam(0))
        T_Patient.lng病人ID = Val(arrParam(1))
        T_Patient.lng主页ID = Val(arrParam(2))
        If UBound(arrParam) > 2 Then T_Patient.lng科室ID = Val(arrParam(3))
        If UBound(arrParam) > 3 Then T_Patient.lng婴儿 = Val(arrParam(4))
    End If
    
    intBaby = T_Patient.lng婴儿
    
    '------------------------------------------------------------------------------------------------------------------
    '打印机恢复及设置
    If Not ExistsPrinter Then
        MsgBox "系统没有安装任何打印机不能继续打印，程序退出！", vbInformation, gstrSysName
        Exit Function
    End If
    
    gPrinter.lngLeft = OFFSET_LEFT
    gPrinter.lngRight = OFFSET_RIGHT
    gPrinter.lngTop = OFFSET_TOP
    gPrinter.lngBottom = OFFSET_BOTTOM
    '提取打印数据
    strSql = "Select 格式 From 病历页面格式 Where 种类 = 3 And 编号 In (Select A.页面 From 病历文件列表 A,病人护理文件 B Where A.Id = B.格式ID and B.ID=[1])"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "提取文件打印设置", T_Patient.lng文件ID)
    If Not rsTmp.EOF Then strPaper = "" & rsTmp!格式
    
    If UBound(Split(strPaper, ";")) >= 0 Then
        gPrinter.intPage = Val(Split(strPaper, ";")(0))
        If UBound(Split(strPaper, ";")) >= 1 Then gPrinter.intOrient = Val(Split(strPaper, ";")(1))
        If UBound(Split(strPaper, ";")) >= 2 Then gPrinter.lngHeight = Val(Split(strPaper, ";")(2))
        If UBound(Split(strPaper, ";")) >= 3 Then gPrinter.lngWidth = Val(Split(strPaper, ";")(3))
        If UBound(Split(strPaper, ";")) >= 4 Then gPrinter.lngLeft = CLng(Val(Split(strPaper, ";")(4)) / conRatemmToTwip)
        If UBound(Split(strPaper, ";")) >= 5 Then gPrinter.lngRight = CLng(Val(Split(strPaper, ";")(5)) / conRatemmToTwip)
        If UBound(Split(strPaper, ";")) >= 6 Then gPrinter.lngTop = CLng(Val(Split(strPaper, ";")(6)) / conRatemmToTwip)
        If UBound(Split(strPaper, ";")) >= 7 Then gPrinter.lngBottom = CLng(Val(Split(strPaper, ";")(7)) / conRatemmToTwip)
    End If
    
    If strPrintDevice = "" Then
        If Trim(zldatabase.GetPara("体温单打印机", glngSys, 1255, "")) = "" Then
            MsgBox "没有设置打印机,将使用系统默认打印机设置！", vbInformation, gstrSysName
            strPrintName = Printer.DeviceName
        Else
            strPrintName = Trim(zldatabase.GetPara("体温单打印机", glngSys, 1255, Printer.DeviceName))
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
    '缺省使用打印机默认禁止，此处不再设置(只要设置了禁止方式打印就不正常）
    gPrinter.intBin = Val(zldatabase.GetPara("体温单进纸", glngSys, 1255, Printer.PaperBin))
    
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
        Call SetNTPrinterPaper(frmFlash.hWnd, gPrinter.lngWidth / conRatemmToTwip, gPrinter.lngHeight / conRatemmToTwip, Printer.Orientation, Printer.Copies)
        Unload frmFlash
    End If
    
    On Error GoTo ErrHandle
    
    '------------------------------------------------------------------------------------------------------------------
    lngBeginY = IIf(gPrinter.lngTop > lngBeginY, gPrinter.lngTop, lngBeginY)
    lngIndex = mintPage
    
    '如果只打印当前就只将开始和结束写同一页码
    Set mfrmCaseTendBodyPrint = New frmCaseTendBodyPrint
    Load frmTendFileRead
    Call frmTendFileRead.InitRechBox(T_Patient.lng文件ID)
    strMarkDate = ""
    '提取用户设置的体温单开始时间(婴儿还是以婴儿出生时间为准)
    strSql = "select 开始时间 from 病人护理文件 where ID=[1] and 病人ID=[2] and 主页id=[3] and nvl(婴儿,0)=[4]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSql, "提取体温单开始时间", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿)
    If rsTmp.RecordCount <> 0 Then
        strMarkDate = Format(rsTmp!开始时间, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If strMarkDate <> "" Then strMarkDate = "to_date('" & strMarkDate & "','yyyy-MM-dd hh24:mi:ss')"
    
     '提取婴儿医嘱信息(转科，出院)存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "   (SELECT /*+ RULE */  病人ID,主页ID,婴儿时间,DECODE(nvl(婴儿,0),0, DECODE(NVL(出院日期,''),'',0,1), DECODE(NVL(婴儿时间,''),'',0,1))记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID,A.主页ID,B.开始执行时间 婴儿时间, A.出院日期,B.婴儿" & vbNewLine & _
                "           FROM 病案主页 A," & vbNewLine & _
                "               (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND nvl(B.婴儿,0)<>0 AND C.类别 = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.操作类型 = COLUMN_VALUE) And  B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "           WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "           ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
                
    '读取此病人的体温单总页数
    '------------------------------------------------------------------------------------------------------------------
    strSql = "SELECT DECODE(C.出生时间,NULL," & IIf(strMarkDate = "", "B.入院时间", strMarkDate) & ",C.出生时间) AS 入院时间," & vbNewLine & _
                "    DECODE(E.记录,0,DECODE(SIGN(NVL(E.婴儿时间,B.出院时间) - D.发生时间), 1,NVL(E.婴儿时间,B.出院时间) ,D.发生时间),NVL(E.婴儿时间,B.出院时间))  出院时间," & vbNewLine & _
                "    1 + TRUNC((TO_DATE(TO_CHAR(DECODE(E.记录,0,DECODE(SIGN(NVL(E.婴儿时间,B.出院时间) - D.发生时间), 1,NVL(E.婴儿时间,B.出院时间) ,D.发生时间),NVL(E.婴儿时间,B.出院时间)),'yyyy-MM-dd'),'yyyy-MM-dd') - " & vbNewLine & _
                "    TO_DATE(TO_CHAR(DECODE(C.出生时间,NULL," & IIf(strMarkDate = "", "B.入院时间", strMarkDate) & ",C.出生时间),'yyyy-MM-dd'),'yyyy-MM-dd')) / 7) AS 页数,D.发生时间" & vbNewLine & _
                "    FROM (SELECT 病人ID,主页ID,MIN(开始时间) AS 入院时间," & vbNewLine & _
                "    MAX(NVL(终止时间, SYSDATE)) AS 出院时间" & vbNewLine & _
                "    FROM 病人变动记录" & vbNewLine & _
                "    WHERE 开始时间 IS NOT NULL AND 病人ID = [2] AND 主页ID = [3] GROUP BY 病人ID,主页ID) B," & vbNewLine & _
                "    (SELECT 病人ID,主页ID,出生时间 FROM 病人新生儿记录 WHERE 病人ID =[2] AND 主页ID =[3] AND 序号=[4]) C ," & vbNewLine & _
                "    (SELECT NVL(发生时间,SYSDATE) 发生时间 FROM ( SELECT MAX(发生时间) 发生时间 FROM 病人护理文件 A,病人护理数据 B" & vbNewLine & _
                "    WHERE A.ID=B.文件ID AND A.ID=[1] AND A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=[4])) D," & vbNewLine & _
                strNewSql & vbNewLine & _
                "WHERE B.病人ID=E.病人ID And B.主页ID=E.主页ID And B.病人ID=C.病人ID(+) AND B.主页ID=C.主页ID(+)"

    Set rsTmp = zldatabase.OpenSQLRecord(strSql, mstrTitle, T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿)
    intCount = 0
    For intCOl = 0 To rsTmp("页数").Value - 1
    
        strDateFrom = Format(rsTmp("入院时间").Value + 7 * intCOl, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("入院时间").Value + 7 * (intCOl + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
        If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
        End If
        
        If strDateFrom < Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
        
            If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss")
            
            ReDim Preserve strArrFromTo(intCount)
            strArrFromTo(intCount) = "0;" & intCOl + 1 & ";" & intCOl + 1
            intCount = intCount + 1
        End If
    Next
    
    If blnPrint = True Then
        Set objPrint = Printer
    Else
        Set objPrint = mfrmCaseTendBodyPrint
    End If
    
    Select Case intPrintRange
    Case 0                  '打印当前页
        If InStr(1, strPage, ";") <> 0 Then
            lngIndex = Val(Split(strPage, ";")(1))
        End If
        strPage = lngIndex & ";" & lngIndex
        
        If blnPrint = True Then Printer.Print ""
        If PrintOrPreviewBodyState(objPrint, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng文件ID, intBaby, _
                T_Patient.lng科室ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, False, _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, , mblnMoved) = True Then
                
                If blnPrint = False Then
                    mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, T_Patient.lng病人ID, T_Patient.lng主页ID, _
                        T_Patient.lng文件ID, CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                        CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, strPage, T_Patient.lng科室ID, T_Patient.lng婴儿
                Else
                    'Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                    Printer.EndDoc
                End If
        Else
            MsgBox "未知错误，输出体温单失败！", vbExclamation, gstrSysName
        End If
        
    Case 1              '从当前页连续打印
        If InStr(1, strPage, ";") <> 0 Then
            lngIndex = Val(Split(strPage, ";")(0))
            lngIndexEnd = Val(Split(strPage, ";")(1))
            If lngIndexEnd > UBound(strArrFromTo) Then lngIndexEnd = UBound(strArrFromTo)
            
        Else
            lngIndexEnd = UBound(strArrFromTo)
        End If
        
        strPage = lngIndex & ";" & lngIndexEnd
        
        For intCOl = lngIndex To lngIndexEnd
            If blnPrint = True Then Printer.Print ""
            If PrintOrPreviewBodyState(objPrint, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng文件ID, intBaby, _
                T_Patient.lng科室ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, intCOl <> lngIndex, _
                CInt(Split(strArrFromTo(intCOl), ";")(1)), CInt(Split(strArrFromTo(intCOl), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "未知错误，打印失败！", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                'Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                If intCOl = UBound(strArrFromTo) Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, T_Patient.lng病人ID, T_Patient.lng主页ID, _
            T_Patient.lng文件ID, CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, strPage, T_Patient.lng科室ID, T_Patient.lng婴儿
        Else '连续打印是记录打印的开始页号和结束页号
            strSql = "zl_体温单数据_Printer(" & T_Patient.lng文件ID & "," & lngIndex + 1 & "," & lngIndexEnd + 1 & ")"
            Call zldatabase.ExecuteProcedure(strSql, "zl_体温单数据_Printer")
        End If
        
    Case 2          '从第一页连续打印,即全部打印
        strPage = 0
        For intCOl = 0 To UBound(strArrFromTo)
            If blnPrint = True Then Printer.Print ""
            If PrintOrPreviewBodyState(objPrint, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng文件ID, intBaby, _
                T_Patient.lng科室ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, intCOl <> 0, _
                CInt(Split(strArrFromTo(intCOl), ";")(1)), CInt(Split(strArrFromTo(intCOl), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "未知错误，打印失败！", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                'Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                If intCOl = UBound(strArrFromTo) Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, T_Patient.lng病人ID, T_Patient.lng主页ID, _
            T_Patient.lng文件ID, CInt(Split(strArrFromTo(0), ";")(1)), _
                CInt(Split(strArrFromTo(0), ";")(1)), intPageNo, strArrFromTo, strPage, T_Patient.lng科室ID, T_Patient.lng婴儿
        End If
    End Select
    
    'WinNT自定义纸张处理
    If IsWindowsNT And gPrinter.intPage = 256 Then DelCustomPaper
    
    Unload frmTendFileRead
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'将PictureBox模拟成3D平面按钮
'intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
Private Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .Cls
        .BorderStyle = 0
        
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            
            Select Case IntStyle
                Case 1
                    DrawEdge .hDC, PicRect, CLng(BDR_RAISEDINNER), BF_RECT
                Case 2
                    DrawEdge .hDC, PicRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hDC, PicRect, CLng(BDR_SUNKENOUTER), BF_RECT
                Case -2
                    DrawEdge .hDC, PicRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Private Sub vsf_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call VsfDrawCell(hDC, Row, Col, Left, Top, Right, Bottom, Done)
End Sub

Private Sub VsfDrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'----------------------------------
'功能:完成单元格绘图
'-----------------------------------
    Dim T_ClientRect As RECT
    Dim strText As String, strTmp As String, strPart
    Dim lngBrush As Long, lngOldBrush As Long
    Dim lngBackColor As Long, lngForeColor As Long
    Dim i As Integer
    Dim int列 As Integer, int行 As Integer
    
     
    If mbln呼吸曲线 Or mItemNO.呼吸 = 0 Then Exit Sub
    If vsf.Body.RowHidden(Row) Or vsf.Body.ColHidden(Col) Then Exit Sub
    If Col < vsf.FixedCols Then Exit Sub
             
    On Error GoTo Errhand
    '设定客户区域大小
    With T_ClientRect
        .Left = Left + 1
        .Top = Top + 1
        .Right = Right - 1
        .Bottom = Bottom - 1
    End With
    
    '只花图形
    If Val(vsf.ColData(Col)) <> 0 Then
        '提取图形位置
        mrsGraph.Filter = "项目序号=" & gint呼吸 & " And 部位='呼吸机'"
        If mrsGraph.RecordCount = 0 Then
            int列 = -1: int行 = -1
        Else
            int列 = Val(mrsGraph!列)
            int行 = Val(mrsGraph!行)
        End If
        
        '1、清空内容
        '创建与背景色相同的刷子
        lngBackColor = vsf.Body.Cell(flexcpBackColor, Row, Col, Row, Col)
        If lngBackColor = 0 Then lngBackColor = vsf.Body.BackColor
        lngBackColor = GetRBGFromOLEColor(lngBackColor)
        lngForeColor = 200
        lngBrush = CreateSolidBrush(lngBackColor)
        '使用该刷子填充背景色
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, T_ClientRect, lngBrush)
        '立即销毁临时使用的刷子并还原刷子
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        T_ClientRect.Left = Left + (T_ClientRect.Right - Left - gintBmpW) / 2
        If Val(vsf.ColData(Col)) = 2 Then
            T_ClientRect.Top = Top + (T_ClientRect.Bottom - gintBmpH)
        End If
        '开始进行图形
        Call BitBlt(hDC, T_ClientRect.Left, T_ClientRect.Top, gintBmpW, gintBmpH, mobjBuffer.hDC, int列 * gintBmpW, int行 * gintBmpH, SRCCOPY)
        mrsGraph.Filter = 0
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawDownTabAnsyGrade(ByVal lngDC As Long, ByVal objDraw As Object, arrText() As String, ByVal Row As Long, ByVal Col As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean, Optional ByVal blnFormat As Boolean = False)
'---------------------------------------------------
'功能 大便次数输出
'说明 AnsyGrade=True才能调用此函数
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, intSize As Integer, intOldSize As Integer
    Dim lngBrush As Long, lngOldBrush As Long
    Dim lngBackColor As Long, lngForeColor As Long
    Dim stdset As StdFont, stdOldset As StdFont
    Dim LPoint As T_LPoint, T_ClientRect As RECT
    Dim str1 As String, str2 As String, str3 As String, strTmp As String
    Dim lngX As Long, lngY As Long, sngH As Single, sngW As Single
    
    On Error GoTo Errhand
    
    If UBound(arrText) < 2 Then Exit Sub
    
     '设定客户区域大小
    With T_ClientRect
        .Left = Left + 1
        .Top = Top + 1
        .Right = Right - 1
        .Bottom = Bottom - 1
        LPoint.W = .Right - .Left
        LPoint.X = .Left
        LPoint.Y = .Top + (.Bottom - .Top) / 2
    End With
    
    '1、清空内容
    '创建与背景色相同的刷子
    lngBackColor = mshDownTab.Cell(flexcpBackColor, Row, Col, Row, Col)
    If lngBackColor = 0 Then lngBackColor = objDraw.BackColor
    lngBackColor = GetRBGFromOLEColor(lngBackColor)
    lngForeColor = GetRBGFromOLEColor(mshDownTab.Cell(flexcpForeColor, Row, Col, Row, Col))
    lngBrush = CreateSolidBrush(lngBackColor)
    '使用该刷子填充背景色
    lngOldBrush = SelectObject(lngDC, lngBrush)
    Call FillRect(lngDC, T_ClientRect, lngBrush)
    '立即销毁临时使用的刷子并还原刷子
    Call SelectObject(lngDC, lngOldBrush)
    Call DeleteObject(lngBrush)
    
    str1 = arrText(0): str2 = arrText(1): str3 = arrText(2)
    If blnFormat = True Then
        If Len(str2) > Len(str3) Then
            strTmp = str1 & str2
        Else
            strTmp = str1 & str3
        End If
    Else
        strTmp = str1 & str2 & "/" & str3
    End If
    intSize = objDraw.Font.Size
    intOldSize = intSize
    objDraw.Font.Size = intSize
    Set stdset = New StdFont
    stdset.Name = "宋体"
    stdset.Size = intSize
    stdset.Bold = False
    Set stdOldset = stdset '原始字体
    
    Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , 1)
    '输出左边
    If str1 <> "" Then
        Call SetFontIndirect(stdOldset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngForeColor)
        Call DrawText(lngDC, str1, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        lngX = T_LableRect.Left + (objDraw.TextWidth(str1) / T_TwipsPerPixel.X) - (objDraw.TextWidth("a") / T_TwipsPerPixel.X / 2) + 1
    Else
        lngX = T_LableRect.Left
    End If

    If blnFormat = True Then '分子分母显示
        intSize = 7
        objDraw.Font.Size = intSize
        Set stdset = New StdFont
        stdset.Name = "宋体"
        stdset.Size = intSize
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngForeColor)
        T_LableRect.Left = lngX
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        'If T_LableRect.Top < Top Then T_LableRect.Top = Top - 1
        T_LableRect.Bottom = T_ClientRect.Bottom
        Call DrawText(lngDC, str2, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        '画横线
        objDraw.Font.Size = intOldSize
        Call DrawLine(lngDC, lngX, lngY, lngX + (objDraw.TextWidth("A") / T_TwipsPerPixel.X), lngY)
        '输出分母
        lngY = lngY
        T_LableRect.Left = lngX
        T_LableRect.Top = lngY
        intSize = 7.5
        Set stdset = New StdFont
        stdset.Name = "宋体"
        stdset.Size = intSize
        Call SetFontIndirect(stdset, lngDC, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDC, lngFont)
        Call SetTextColor(lngDC, lngForeColor)
        Call DrawText(lngDC, str3, -1, T_LableRect, 0)
        Call SelectObject(lngDC, lngOldFont)
        Call DeleteObject(lngFont)
    Else
        If str1 <> "" Then
            '输出上标
            intSize = 7
            objDraw.Font.Size = intSize
            Set stdset = New StdFont
            stdset.Name = "宋体"
            stdset.Size = intSize
            Call SetFontIndirect(stdset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngForeColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < T_ClientRect.Top Then T_LableRect.Top = T_ClientRect.Top - 1
            Call DrawText(lngDC, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            '输出后半部分
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngForeColor)
            T_LableRect.Left = lngX
            T_LableRect.Top = lngY
            Call DrawText(lngDC, "/" & str3, -1, T_LableRect, 0)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
        Else
            Call SetFontIndirect(stdOldset, lngDC, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDC, lngFont)
            Call SetTextColor(lngDC, lngForeColor)
            Call DrawText(lngDC, str2 & "/" & str3, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDC, lngOldFont)
            Call DeleteObject(lngFont)
        End If
    End If
    
    objDraw.Font.Size = intOldSize
    Set stdset = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

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

