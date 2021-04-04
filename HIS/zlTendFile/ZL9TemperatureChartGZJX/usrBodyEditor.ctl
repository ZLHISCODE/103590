VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl usrBodyEditor 
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   8205
   ScaleWidth      =   10800
   Begin VB.PictureBox picSerach 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2160
      ScaleHeight     =   375
      ScaleWidth      =   1515
      TabIndex        =   26
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
         TabIndex        =   33
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   10215
      Begin VB.TextBox txtLength 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5505
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   6165
         Visible         =   0   'False
         Width           =   1395
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   9600
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   9375
         Begin VB.PictureBox picCommText 
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
            Height          =   300
            Left            =   75
            ScaleHeight     =   300
            ScaleWidth      =   7260
            TabIndex        =   36
            Top             =   4860
            Width           =   7260
         End
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
            TabIndex        =   34
            Top             =   2160
            Width           =   7335
         End
         Begin VSFlex8Ctl.VSFlexGrid mshUpTab 
            Height          =   1095
            Left            =   120
            TabIndex        =   24
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
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   60
               Width           =   165
               Begin VB.Image imgDisPlay 
                  Appearance      =   0  'Flat
                  Height          =   240
                  Left            =   -30
                  Picture         =   "usrBodyEditor.ctx":076A
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
               TabIndex        =   28
               Top             =   720
               Visible         =   0   'False
               Width           =   180
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid mshDownTab 
            Height          =   975
            Left            =   90
            TabIndex        =   25
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
            FormatString    =   $"usrBodyEditor.ctx":6FBC
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
         Begin zl9TemperatureChartGZJX.VsfGrid vsf 
            Height          =   255
            Left            =   120
            TabIndex        =   27
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
            TabIndex        =   7
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
               TabIndex        =   15
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
               Index           =   4
               Left            =   3360
               Locked          =   -1  'True
               TabIndex        =   12
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
               Index           =   6
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   14
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
               TabIndex        =   13
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
               Index           =   0
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   8
               TabStop         =   0   'False
               Text            =   "姓无名"
               Top             =   60
               Width           =   1425
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   2
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   10
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
               Index           =   3
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   11
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
               Index           =   1
               Left            =   6645
               Locked          =   -1  'True
               TabIndex        =   9
               TabStop         =   0   'False
               Text            =   "1234567"
               Top             =   60
               Width           =   3825
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "诊    断:"
               Height          =   180
               Index           =   7
               Left            =   4065
               TabIndex        =   23
               Top             =   390
               Width           =   810
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "床号:"
               Height          =   180
               Index           =   3
               Left            =   2910
               TabIndex        =   19
               Top             =   390
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄:"
               Height          =   180
               Index           =   6
               Left            =   2910
               TabIndex        =   22
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
               TabIndex        =   21
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   16
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科室:"
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   18
               Top             =   375
               Width           =   450
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院日期:"
               Height          =   180
               Index           =   5
               Left            =   4050
               TabIndex        =   20
               Top             =   60
               Width           =   810
            End
            Begin VB.Label lblCard 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号:"
               Height          =   180
               Index           =   1
               Left            =   6000
               TabIndex        =   17
               Top             =   60
               Width           =   630
            End
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
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "临时拷图用,千万别删"
         Top             =   2640
         Visible         =   0   'False
         Width           =   2115
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
         ScaleWidth      =   5220
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   6120
         Width           =   5220
         Begin VB.ComboBox cboFile 
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
            Left            =   3165
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   0
            Width           =   2085
         End
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
            Left            =   690
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   0
            Width           =   1920
         End
         Begin VB.Label lblSerach 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "文件"
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
            Index           =   0
            Left            =   2730
            TabIndex        =   3
            Top             =   60
            Width           =   360
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
            Left            =   225
            TabIndex        =   1
            Top             =   60
            Width           =   360
         End
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
Private mcbrItem    As CommandBarControl

'--变量
Public mblnResize As Boolean '记录窗体大小是否发生变化
Public mblnMoved As Boolean
Private mlngWidth As Long
Private mlngHeight As Long
Private mintPage      As Integer '记录当前页号
Private mintAllPage As Integer '体温单的总页数
Private mintIndex As Integer '记录页码跳转的索引
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
Private mstrOpdays() As String
Private mstrOpValue() As String
Private mstrNewString() As String '保存皮试结果信息
Private mlngNewHeight() As Long '保存皮试结果行高

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
    lng格式ID As Long
End Type
Private T_Patient As type_Patient

'--事件定义
Public Event CmdClick(ByVal strParam As String)
Public Event zlAfterPrint()
Public Event DbClickCur(ByVal intDataEditor As Integer)
Public Event zlFileChange(ByVal blnRefresh As Boolean, ByVal lngFileID As Long, ByVal lngBaby As Long)
Public Event zlDataChange(ByVal blnChange As Boolean)
Public Event ShowTipInfo(ByVal vsfObj As Object, ByVal strInfo As String, ByVal blnMultiRow As Boolean)

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
    Dim strSQL        As String, strNewSql As String
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
    Dim strMarkDate As String, strFileBeginTime As String, strFileEndTime As String '体温单设置时间
    Dim intCOl        As Integer
    Dim strCaption    As String, strCategory As String, strUnitName As String, blnAddMenu As Boolean
    Dim strParameter  As String
    Dim strSvrCaption As String, strSvrCaption1 As String
    Dim strNow        As String
    Dim strCut        As String
    Dim lngLoop       As Long
    Dim strTmp        As String
    Dim lnglast科室id As Long
    Dim lng天数 As Long
    
    On Error GoTo Errhand

    If lng病人ID = 0 And lng文件ID = 0 And lng主页ID = 0 Then Exit Function
    mbln出院 = False
    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
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
    strSQL = "Select 开始时间,结束时间 From 病人护理文件 where ID=[1] and 病人ID=[2] and 主页id=[3] and nvl(婴儿,0)=[4]"
    If mblnMoved = True Then strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取体温单开始时间", lng文件ID, lng病人ID, lng主页ID, int婴儿)
    If rsTmp.RecordCount <> 0 Then
        strEnterDate = Format(rsTmp!开始时间, "YYYY-MM-DD HH:mm:ss")
        strFileEndTime = Format(rsTmp!结束时间, "YYYY-MM-DD HH:mm:ss")
    End If
    strFileBeginTime = strEnterDate
    strMarkDate = "To_date('" & strEnterDate & "','yyyy-MM-dd hh24:mi:ss')"
    '------------------------------------------------------------------------------------------------------------------
    '提取婴儿医嘱信息(转科，出院)存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "   (SELECT 病人ID,主页ID,婴儿时间,DECODE(nvl(婴儿,0),0, DECODE(NVL(出院日期,''),'',0,1), DECODE(NVL(婴儿时间,''),'',0,1))记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID,A.主页ID,B.开始执行时间 婴儿时间, A.出院日期,B.婴儿" & vbNewLine & _
                "           FROM 病案主页 A," & vbNewLine & _
                "               (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND nvl(B.婴儿,0)<>0 AND B.诊疗类别 = 'Z'" & vbNewLine & _
                "                And Instr(',3,5,11,', ',' || c.操作类型 || ',') > 0 And  B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "           WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "           ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
    '说明:目前有了专科体温单，病人可能同时存在多份体温单。体温单开始时间和终止时间的规则如下:
    '如果文件的开始时间不为空并且大于等于病人入院时间或婴儿出生时间,体温单的开始时间以文件开始时间为准,否则以病人入院时间或婴儿出生时间为准
    '如果文件的终止时间不为空并且小于等于病人或婴儿出院时间（未出院不能不能大于当前时间）,体温单结束时间以文件开始时间为准，否则体温单结束时间以病人或婴儿出院时间为准(未出院为当前时间)
    '如果文件的终止时间为空,保持原有方式,病人如果已经出院，就已出院时间为准,未出院就已当前时间或数据结束时间为准.
    '读取此病人的体温单总页数
    '------------------------------------------------------------------------------------------------------------------
    strSQL = " SELECT  入院时间,实际入院时间,出院时间,1 + TRUNC((TO_DATE(TO_CHAR(出院时间,'yyyy-MM-dd'),'yyyy-MM-dd') -TO_DATE(TO_CHAR(入院时间,'yyyy-MM-dd'),'yyyy-MM-dd')) / " & T_BodyStyle.lng天数 & ") AS 页数,发生时间,记录 " & _
             "  From (" & _
             "      SELECT DECODE(D.开始时间,NULL,DECODE(C.出生时间,NULL," & IIf(strMarkDate = "", "B.入院时间", strMarkDate) & ",C.出生时间)," & vbNewLine & _
             "                 DECODE(SIGN(D.开始时间 - DECODE(C.出生时间,NULL," & IIf(strMarkDate = "", "B.入院时间", strMarkDate) & ",C.出生时间))," & vbNewLine & _
             "                        1," & vbNewLine & _
             "                        D.开始时间," & vbNewLine & _
             "                        DECODE(C.出生时间,NULL," & IIf(strMarkDate = "", "B.入院时间", strMarkDate) & ",C.出生时间))) AS 入院时间," & vbNewLine & _
             "      DECODE(C.出生时间,NULL,B.入院时间,C.出生时间) AS 实际入院时间," & vbNewLine & _
             "      DECODE(D.结束时间,NULL," & vbNewLine & _
             "                 DECODE(E.记录,0," & vbNewLine & _
             "                        DECODE(SIGN(NVL(E.婴儿时间, B.出院时间) - D.发生时间), 1, NVL(E.婴儿时间, B.出院时间), D.发生时间)," & vbNewLine & _
             "                        NVL(E.婴儿时间, B.出院时间))," & vbNewLine & _
             "                 DECODE(SIGN(NVL(E.婴儿时间, B.出院时间) - D.结束时间), 1, D.结束时间, NVL(E.婴儿时间, B.出院时间))) 出院时间," & vbNewLine & _
             "      D.发生时间,DECODE(D.结束时间, NULL, E.记录, 1) 记录" & vbNewLine & _
             "      FROM (SELECT 病人ID,主页ID,MIN(开始时间) AS 入院时间," & vbNewLine & _
             "      MAX(NVL(终止时间, SYSDATE)) AS 出院时间" & vbNewLine & _
             "      FROM 病人变动记录" & vbNewLine & _
             "      WHERE 开始时间 IS NOT NULL AND 病人ID = [2] AND 主页ID = [3] GROUP BY 病人ID,主页ID) B," & vbNewLine & _
             "      (SELECT 病人ID,主页ID,出生时间 FROM 病人新生儿记录 WHERE 病人ID =[2] AND 主页ID =[3] AND 序号=[4]) C ," & vbNewLine & _
             "      (SELECT NVL(发生时间, SYSDATE) 发生时间, 开始时间, 结束时间" & vbNewLine & _
             "         FROM (SELECT MAX(B.发生时间) 发生时间, MAX(A.开始时间) 开始时间, MAX(A.结束时间) 结束时间" & vbNewLine & _
             "                FROM 病人护理文件 A, 病人护理数据 B" & vbNewLine & _
             "                WHERE A.ID = B.文件ID(+) AND A.ID = [1] AND A.病人ID = [2] AND A.主页ID = [3] AND A.婴儿 = [4])) D," & vbNewLine & _
             "  " & strNewSql & vbNewLine & _
             "   WHERE B.病人ID=E.病人ID And B.主页ID=E.主页ID And B.病人ID=C.病人ID(+) AND B.主页ID=C.主页ID(+))"
                
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
        strSQL = Replace(strSQL, "病人护理数据", "H病人护理数据")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng文件ID, lng病人ID, lng主页ID, int婴儿)
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
    strSQL = " SELECT /*+ RULE */ 1 + TRUNC((TO_DATE(TO_CHAR(DECODE(D.开始时间,NULL,A.开始时间, DECODE(SIGN(D.开始时间 - A.开始时间), 1, D.开始时间, A.开始时间)),'YYYY-MM-DD'),'YYYY-MM-DD') -" & vbNewLine & _
            "                  TO_DATE(TO_CHAR(DECODE(D.开始时间, NULL, B.入院时间, D.开始时间), 'YYYY-MM-DD'), 'YYYY-MM-DD')) / " & T_BodyStyle.lng天数 & ") AS 开始页码," & vbNewLine & _
            "       1 + TRUNC((TO_DATE(TO_CHAR(DECODE(A.序号,F.LAST,DECODE(D.结束时间,NULL," & vbNewLine & _
            "                                                 DECODE(E.记录,0,DECODE(SIGN(NVL(E.婴儿时间, A.终止时间) - D.发生时间),1," & vbNewLine & _
            "                                                        NVL(E.婴儿时间, A.终止时间),D.发生时间),NVL(E.婴儿时间, A.终止时间))," & vbNewLine & _
            "                                                 DECODE(SIGN(NVL(E.婴儿时间, A.终止时间) - D.结束时间),1,D.结束时间,NVL(E.婴儿时间, A.终止时间)))," & vbNewLine & _
            "                                          DECODE(D.结束时间, NULL,NVL(E.婴儿时间, A.终止时间)," & vbNewLine & _
            "                                                 DECODE(SIGN(D.结束时间 - NVL(E.婴儿时间, A.终止时间)),1,NVL(E.婴儿时间, A.终止时间),D.结束时间)))," & vbNewLine & _
            "                                   'YYYY-MM-DD')," & vbNewLine & _
            "                           'YYYY-MM-DD') - TO_DATE(TO_CHAR(DECODE(D.开始时间, NULL, B.入院时间, D.开始时间), 'YYYY-MM-DD'), 'YYYY-MM-DD')) / " & T_BodyStyle.lng天数 & ") AS 结束页码," & vbNewLine & _
            "                          D.发生时间, 病区ID, C.名称, DECODE(D.开始时间,NULL,A.开始时间,DECODE(SIGN(D.开始时间 - A.开始时间), 1, D.开始时间, A.开始时间)) 开始时间," & vbNewLine & _
            "      DECODE(A.序号,F.LAST,DECODE(D.结束时间,NULL," & vbNewLine & _
            "                           DECODE(E.记录,0,DECODE(SIGN(NVL(E.婴儿时间, A.终止时间) - D.发生时间),1," & vbNewLine & _
            "                                  NVL(E.婴儿时间, A.终止时间),D.发生时间),NVL(E.婴儿时间, A.终止时间))," & vbNewLine & _
            "                           DECODE(SIGN(NVL(E.婴儿时间, A.终止时间) - D.结束时间),1,D.结束时间,NVL(E.婴儿时间, A.终止时间)))," & vbNewLine & _
            "                    DECODE(D.结束时间, NULL,NVL(E.婴儿时间, A.终止时间)," & vbNewLine & _
            "                           DECODE(SIGN(D.结束时间 - NVL(E.婴儿时间, A.终止时间)),1,NVL(E.婴儿时间, A.终止时间),D.结束时间))) 终止时间"
    strSQL = strSQL & _
            " FROM (SELECT ROWNUM 序号, 病区ID, 开始时间, 终止时间" & vbNewLine & _
            "       FROM (SELECT 病区ID, MIN(开始时间) AS 开始时间, MAX(NVL(终止时间, SYSDATE)) AS 终止时间" & vbNewLine & _
            "              FROM 病人变动记录" & vbNewLine & _
            "              WHERE ((开始时间>=[5]" & IIf(IsDate(strFileEndTime), " And 开始时间<[6]", " ") & ") OR (开始时间<=[5] And (终止时间 IS NULL OR 终止时间>[5]))) AND 病人ID = [2] AND 主页ID = [3]" & vbNewLine & _
            "              GROUP BY 病区ID" & vbNewLine & _
            "              ORDER BY 开始时间,终止时间)) A," & vbNewLine & _
            "     (SELECT DECODE(Y.出生时间, NULL, X.入院时间, Y.出生时间) AS 入院时间, X.病人ID, X.主页ID" & vbNewLine & _
            "       FROM (SELECT 病人ID, 主页ID, MIN(开始时间) AS 入院时间" & vbNewLine & _
            "              FROM 病人变动记录" & vbNewLine & _
            "              WHERE ((开始时间>=[5]" & IIf(IsDate(strFileEndTime), " And 开始时间<[6]", " ") & ") OR (开始时间<=[5] And (终止时间 IS NULL OR 终止时间>[5]))) AND 病人ID = [2] AND 主页ID = [3]" & vbNewLine & _
            "              GROUP BY 病人ID,主页ID) X," & vbNewLine & _
            "            (SELECT 病人ID, 主页ID, 出生时间 FROM 病人新生儿记录 WHERE 病人ID = [2] AND 主页ID = [3] AND 序号 = [4]) Y" & vbNewLine & _
            "       WHERE X.病人ID = Y.病人ID(+) AND X.主页ID = Y.主页ID(+)) B, 部门表 C," & vbNewLine & _
            "     (SELECT NVL(发生时间, SYSDATE) 发生时间, 开始时间, 结束时间" & vbNewLine & _
            "       FROM (SELECT MAX(发生时间) 发生时间, MAX(A.开始时间) 开始时间, MAX(A.结束时间) 结束时间" & vbNewLine & _
            "              FROM 病人护理文件 A, 病人护理数据 B" & vbNewLine & _
            "              WHERE A.ID = B.文件ID(+) AND A.ID = [1] AND A.病人ID = [2] AND A.主页ID = [3]  AND A.婴儿 = [4])) D," & vbNewLine & strNewSql & "," & vbNewLine & _
            "     (SELECT COUNT(*) LAST" & vbNewLine & _
            "       FROM (SELECT 病区ID" & vbNewLine & _
            "              FROM 病人变动记录" & vbNewLine & _
            "              WHERE ((开始时间>=[5]" & IIf(IsDate(strFileEndTime), " And 开始时间<[6]", " ") & ") OR (开始时间<=[5] And (终止时间 IS NULL OR 终止时间>[5]))) AND 病人ID = [2] AND 主页ID = [3]" & vbNewLine & _
            "              GROUP BY 病区ID)) F" & vbNewLine & _
            " WHERE B.病人ID = E.病人ID AND B.主页ID = E.主页ID AND C.ID(+) = A.病区ID" & vbNewLine & _
            " ORDER BY A.开始时间"
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
        strSQL = Replace(strSQL, "病人护理数据", "H病人护理数据")
    End If
    If (IsDate(strFileEndTime)) Then
        Set RS = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng文件ID, lng病人ID, lng主页ID, int婴儿, CDate(strFileBeginTime), CDate(strFileEndTime))
    Else
        Set RS = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng文件ID, lng病人ID, lng主页ID, int婴儿, CDate(strFileBeginTime))
    End If
    
    For lngLoop = 0 To rsTmp("页数").Value - 1

        strDateFrom = Format(rsTmp("入院时间").Value + T_BodyStyle.lng天数 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("入院时间").Value + T_BodyStyle.lng天数 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"

        If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
        End If

        If strDateFrom < Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then

            If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss")
            
            RS.Filter = ""
            RS.Filter = "开始页码<=" & lngLoop + 1 & " And 结束页码>=" & lngLoop + 1
            RS.Sort = "开始时间,终止时间"
            blnAddMenu = (RS.RecordCount = 1)
            strCategory = ""
ToStart:
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
                
                If blnAddMenu = False Then
                    If intCOl = 1 Then
                        strCategory = Format(strTmp, "yyyy-MM-dd")
                        strUnitName = Nvl(RS("名称").Value)
                    ElseIf intCOl = RS.RecordCount Then
                        blnAddMenu = True
                        strCategory = strCategory & "～" & Format(strCaption, "yyyy-MM-dd")
                        If strUnitName <> Nvl(RS("名称").Value) Then
                            strUnitName = strUnitName & "->" & Nvl(RS("名称").Value)
                        End If
                        strCategory = "第" & lngLoop + 1 & "页：" & strCategory & "(" & strUnitName & ")"
                    End If
                    RS.MoveNext
                    If blnAddMenu = True Then GoTo ToStart
                Else
                    strCaption = Format(strTmp, "yyyy-MM-dd") & "～" & Format(strCaption, "yyyy-MM-dd")
                    strCaption = "第" & lngLoop + 1 & "页：" & strCaption & "(" & RS("名称").Value & ")"
                    If strCategory = "" Then strCategory = strCaption
                    '入院时间;科室id;开始时间;结束时间;
                    Set mcbrItem = mcbrToolBar页面.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
                    mcbrItem.Parameter = strEnterDate & ";" & RS!病区ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop & ";" & strOutDate
                    mcbrItem.Category = strCategory
                    
                    If lngLoop + 1 <= 4 Then
                        Set cbrWeek = mcbrToolBar.FindControl(, Val(ArrControlId(lngLoop)))
                        cbrWeek.Parameter = strEnterDate & ";" & RS!病区ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop & ";" & strOutDate
                        cbrWeek.Category = strCategory
                    End If
                     
                    lnglast科室id = Val(Nvl(RS("病区ID").Value))
                    RS.MoveNext
                    strParameter = mcbrItem.Parameter
                    
                    '指定页号不为0 并且和该页数相等就记录参数值
                    If T_Patient.lngPage <> 0 And Val(T_Patient.lngPage - 1) = lngLoop Then
                        strParam1 = strParameter
                        strSvrCaption1 = strCategory
                    End If
                    
                    strSvrCaption = strCategory
                End If
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
    
    
    For lngLoop = 0 To Round((DateDiff("D", CDate(ArrCode(0)), CDate(ArrCode(5))) + 1) / T_BodyStyle.lng天数)

        strDateFrom = Format(CDate(ArrCode(0)) + T_BodyStyle.lng天数 * lngLoop, "yyyy-MM-dd") & " 00:00:00"

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
            mblnMoved = False
            
            mstrSQL = "Select 出院科室ID,nvl(数据转出,0) 转出 from 病案主页 Where 病人id=[1] And 主页id=[2] "
            Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "提取科室ID", T_Patient.lng病人ID, T_Patient.lng主页ID)
            If RS.BOF = False Then
                T_Patient.lng科室ID = Val(zlCommFun.Nvl(RS("出院科室ID").Value))
                If T_Patient.lng出院 = 1 Then mblnMoved = (Val(RS("转出")) <> 0)
            End If
            
            '提取初始体温格式构造数据
            If Not GetStyleBody(T_Patient.lng文件ID, T_Patient.lng护理等级, T_Patient.lng婴儿, T_Patient.lng科室ID) Then Exit Function
    
            mstrSQL = "SELECT A.序号,A.姓名 FROM(" & vbNewLine & _
                        "SELECT A.序号,A.姓名,A.病人ID,A.主页ID FROM (SELECT 0 序号,'病人本人' AS 姓名,A.病人ID,A.主页ID" & vbNewLine & _
                        "            FROM 病案主页 A, 病人信息 B" & vbNewLine & _
                        "            WHERE A.病人ID = B.病人ID AND A.病人ID =[1] AND A.主页ID =[2]" & vbNewLine & _
                        "            UNION ALL" & vbNewLine & _
                        "            SELECT A.序号, DECODE(A.婴儿姓名, NULL, NVL(C.姓名,B.姓名) || '之子' || TRIM(TO_CHAR(A.序号, '9')), A.婴儿姓名) AS 姓名,A.病人ID,A.主页ID" & vbNewLine & _
                        "            FROM 病人信息 B,病案主页 C,病人新生儿记录 A" & vbNewLine & _
                        "            WHERE B.病人ID=C.病人ID And C.病人ID=A.病人ID And C.主页ID=A.主页ID And C.病人ID =[1] AND C.主页ID =[2]) A," & vbNewLine & _
                        "            (SELECT A.病人ID,A.主页ID , NVL(A.婴儿,0) 婴儿 FROM 病人护理文件 A,病历文件列表 B" & vbNewLine & _
                        "            WHERE A.格式ID=B.ID AND B.种类=3 AND B.保留=-1 And A.病人ID =[1] AND A.主页ID =[2] GROUP BY A.病人ID,A.主页ID,NVL(A.婴儿,0)) B" & vbNewLine & _
                        "            WHERE A.病人ID=B.病人ID AND A.主页ID=B.主页ID AND A.序号=B.婴儿) A" & vbNewLine & _
                        "ORDER BY A.序号"
            If mblnMoved = True Then
                mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            End If
            Set RS = zlDatabase.OpenSQLRecord(mstrSQL, mstrTitle, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng格式ID)
            
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
            '提取病人文件列表
            mstrSQL = "select A.ID,A.文件名称 From 病人护理文件 A,病历文件列表 B" & _
               "    where A.病人ID=[1] and A.主页Id=[2] and nvl(A.婴儿,0)=[3] and A.格式ID=B.ID and B.种类=3 and B.保留=-1 Order by A.开始时间"
            If mblnMoved = True Then
                mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            End If
            Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "提取文件", T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿)
            cboFile.Clear
            With RS
                Do While Not .EOF
                    cboFile.AddItem Nvl(!文件名称)
                    cboFile.ItemData(cboFile.NewIndex) = !Id
                    .MoveNext
                    If cboFile.ListIndex = -1 And T_Patient.lng文件ID = Val(cboFile.ItemData(cboFile.NewIndex)) Then
                        Call zlControl.CboSetIndex(cboFile.hWnd, cboFile.NewIndex)
                        T_Patient.lng文件ID = cboFile.ItemData(cboFile.ListIndex)
                    End If
                Loop
            End With
            If cboFile.ListCount > 0 And cboFile.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cboFile.hWnd, 0)
                T_Patient.lng文件ID = cboFile.ItemData(cboFile.ListIndex)
            End If
            
            RaiseEvent zlFileChange(False, T_Patient.lng文件ID, T_Patient.lng婴儿)
            
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
            mstrEnterDate = Format(varParam(0), "YYYY-MM-DD HH:mm:ss")
            strStartDate = Format(varParam(2), "YYYY-MM-DD HH:mm:ss")
            strEndDate = Format(varParam(3), "YYYY-MM-DD HH:mm:ss")
            mintPage = Val(varParam(4))
            glngCurPage = mintPage + 1
            mstrEndDate = Format(varParam(5), "YYYY-MM-DD HH:mm:ss")
            If mbln出院 = True Then
                '出院时间和入院时间如果在同一列，则将出院时间后移一列（内蒙需求:出院也要录入体温）
                mstrEndDate = Format(RetrunEndTimeNew(CDate(mstrEnterDate), CDate(mstrEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
                strEndDate = Format(RetrunEndTimeNew(CDate(mstrEnterDate), CDate(strEndDate), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
            End If
            If strStartDate & ";" & strEndDate = picMain.Tag And mblnResize = True Then
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
                    'Debug.Print Now & ":加载表格数据"
                    Call ShowDowntab '加载下表格数据
                    Call picDraw_Paint '从内存中Copy画布到PIC
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
                
                mstrSQL = "Select Decode(a.婴儿姓名,Null,NVL(C.姓名,B.姓名) ||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名,a.婴儿性别,a.出生时间 " & _
                    " From 病人信息 B,病案主页 C,病人新生儿记录 A " & _
                    " Where B.病人ID=C.病人ID And C.病人ID=A.病人ID And C.主页ID=A.主页ID And C.病人id=[1] And C.主页id=[2] And a.序号=[3]"
                Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "提取婴儿信息", T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿)
                If RS.BOF = False Then
                    txtCard(0).Text = RS("婴儿姓名").Value
                    txtCard(5).Text = RS("婴儿性别").Value
                End If
            End If
            txtCard(6).Text = GetElementValue("年龄", T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿, mstr开始时间)
        Case "体温数据显示设置"
            If T_Patient.lng编辑 = 0 Then Exit Function
            If mstr开始时间 <> "" Then
                '计算选择的列
                intCOl = (picDisplay.Left - mshUpTab.ColWidth(0) + mshUpTab.ColWidth(1)) / mshUpTab.ColWidth(1)
                intCOl = intCOl - T_BodyStyle.lng监测次数 + 1
                If intCOl < mintColMin Then intCOl = mintColMin
                
                '计算得到列返回的时间范围
                If Trim(strParam) <> "" Then '在体温编辑界面调用显示是传入时间(因为保存数据体温单刷新后,会定位到第一天)
                    strTime = Format(varParam(0), "YYYY-MM-DD HH:mm:ss")
                Else
                    strTime = Split(GetCurveDateNew(intCOl, mstr开始时间, gintHourBegin), ";")(0)
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
                    RaiseEvent zlDataChange(True)
                End If
            End If
            
        Case "体温数据编辑"
            If T_Patient.lng编辑 = 0 Then Exit Function
            Dim strCurDate As String, strDay As String
            If mstr开始时间 <> "" Then
            If picMain.Tag = "" Then picMain.Tag = mstr开始时间 & ";" & mstr结束时间
                strCurDate = zlDatabase.Currentdate
                '计算得到列返回的时间范围
                If Trim(strParam) <> "" Then
                    strTime = Format(varParam(0), "YYYY-MM-DD HH:mm:ss") & ";" & Format(varParam(1), "YYYY-MM-DD HH:mm:ss")
                Else
                    '计算选择的列
                    intCOl = (picDisplay.Left - mshUpTab.ColWidth(0) + mshUpTab.ColWidth(1)) / mshUpTab.ColWidth(1)
                    intCOl = intCOl - T_BodyStyle.lng监测次数 + 1
                    If intCOl < mintColMin Then intCOl = mintColMin
                    strTime = GetCurveDateNew(intCOl, mstr开始时间, gintHourBegin)
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
                    RaiseEvent zlDataChange(True)
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

Private Function InitData(ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal lng出院 As Long, ByVal lng编辑 As Long, ByVal int婴儿 As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
            
    '读取病人数据
    T_Patient.lng病人ID = lng病人ID
    T_Patient.lng主页ID = lng主页ID
    T_Patient.lng出院 = lng出院
    T_Patient.lng编辑 = lng编辑
    
    '加载初始化参数,设置曲线时间段
    Call InitPara(T_BodyStyle.bln专科)
    
    '进行必要的检查
    '获取病人当前护理等级
    T_Patient.lng护理等级 = 3
    gstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取病人当前护理等级", T_Patient.lng病人ID, T_Patient.lng主页ID)
    If rsTemp.BOF = False Then T_Patient.lng护理等级 = zlCommFun.Nvl(rsTemp("护理等级"), 3)

    '检查是否有曲线体温项目
    gstrSQL = " Select 1 From 体温记录项目 A,诊治所见项目 B,护理记录项目 C " & _
              " Where C.项目序号=A.项目序号 " & _
                        "AND C.项目ID=B.ID(+) " & _
                        "AND C.护理等级>=[1] " & _
                        "And A.记录法=1 And RowNum<2 And C.项目序号<>" & gint心率
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在曲线项目", T_Patient.lng护理等级)
    If rsTemp.EOF Then
        MsgBox "至少要有一个曲线项目！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '判断该病人是否已经转出
    If T_Patient.lng病人ID > 0 And T_Patient.lng出院 = 1 Then
        gstrSQL = "select nvl(数据转出,0) 转出 from 病案主页 where 病人ID=[1] and 主页ID=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查病人是否转出", T_Patient.lng病人ID, T_Patient.lng主页ID)
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
    vsf.Body.RowHeight(vsf.FixedRows) = 300
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
    Dim bln入科显示入院 As Boolean, bln显示诊断 As Boolean
    On Error GoTo hErr
    
    strStart = mstr开始时间
    strTo = mstr结束时间
    bln显示诊断 = (Val(zlDatabase.GetPara("体温单显示诊断", glngSys, 1255, 1)) = 1)
    If Not bln显示诊断 Then
        lblCard(7).Visible = False
        txtCard(7).Visible = False
    Else
        lblCard(7).Visible = True
        txtCard(7).Visible = True
    End If
    
    If CStr(mstrEndDate) < Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And mbln出院 = False Then
        mstrEndDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    End If
    
    If mintAllPage = mintPage + 1 Then
        If CStr(mstr结束时间) < Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") And mbln出院 = False Then
            mstr结束时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    txtCard(3).Text = ""
    
    '如果是新生儿，则重新计算时间，即婴儿体温单的开始时间
    If T_Patient.lng婴儿 > 0 Then
        mstrSQL = " Select  b.出生时间 From 病人新生儿记录 B Where 病人id=[1] And 主页id=[2] And 序号=[3] "
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "提取新生儿信息", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID), T_Patient.lng婴儿)
        If rsTmp.BOF = False Then
            txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("出生时间").Value), "yyyy-MM-dd")
        End If
    End If
    
    '此处进行时间转换
    intCOl = GetCurveColumnNew(CDate(strStart), CDate(strStart), gintHourBegin) + mshUpTab.FixedCols - 1
    strStart = Split(GetCurveDateNew(intCOl - mshUpTab.FixedCols + 1, CDate(strStart), gintHourBegin), ";")(0)
    
    If CDate(strStart) < CDate(mstr开始时间) Then
        strStart = Format(mstr开始时间, "yyyy-MM-dd HH:mm:ss")
    End If
    
    intCOl = GetCurveColumnNew(CDate(strTo), CDate(strStart), gintHourBegin) + mshUpTab.FixedCols - 1
    strTo = Split(GetCurveDateNew(intCOl - mshUpTab.FixedCols + 1, CDate(strStart), gintHourBegin), ";")(1)
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
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "变动记录", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID))
    If rsTmp.BOF = False Then
        If txtCard(3).Text = "" And bln入科显示入院 = True Then txtCard(3).Text = Format(zlCommFun.Nvl(rsTmp("开始时间").Value), "yyyy-MM-dd")
    End If
    
    '读取病人基本信息
    mstrSQL = " Select  NVL(A.姓名,b.姓名) 姓名,A.住院号,A.入院日期 入院时间,NVL(A.性别,b.性别) 性别,NVL(A.年龄,b.年龄) 年龄" & _
        " From 病人信息 B,病案主页 A Where A.病人ID=B.病人ID And A.病人id=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "提取病人信息", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID))
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
            " From 病人变动记录 a,部门表 b,部门表 c " & _
            " Where a.病人id=[1] And a.主页id=[2] And a.科室id Is Not Null And a.病区id=b.id and a.科室id=c.id  And NVL(A.附加床位,0)=0 " & _
            " And a.开始时间-4/24<=[3] And Nvl(a.终止时间,Sysdate)>=[4] Order By a.开始时间"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "读取病人科室、床号等信息", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID), CDate(mstr结束时间), CDate(mstr开始时间))
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
    If bln显示诊断 = True Then
        '提取诊断的最小时间
        strStart = GetDiagnoseMinTime(T_Patient.lng病人ID, T_Patient.lng主页ID, CDate(strStart), mblnMoved)
        '提取病人诊断信息
        mstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2,NULL,0,[4]) As 最后诊断 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "最后诊断", "最后诊断", Val(T_Patient.lng病人ID), Val(T_Patient.lng主页ID), CDate(strStart))
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
    End If
    
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
    Dim lng列数 As Long
    Dim PicRect  As RECT
    On Error GoTo Errhand
    
    '提取基础数据
    Call InitPublicData
    
    mbln呼吸曲线 = True
    lng列数 = T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数
    T_DrawClient.列单位 = T_BodyStyle.lng曲线列宽 \ Screen.TwipsPerPixelX
    T_DrawClient.刻度区域.Left = T_DrawClient.偏移量X
    '得到曲线总数
    lngCount = CurveCount
    
    '结算体温单刻度区域的左右边距
    If lngCount <= 3 Then
        T_DrawClient.刻度区域.Right = T_DrawClient.刻度区域.Left + T_BodyStyle.lng刻度宽度 \ Screen.TwipsPerPixelX
    Else
        T_DrawClient.刻度区域.Right = T_DrawClient.刻度区域.Left + T_BodyStyle.lng刻度宽度 \ Screen.TwipsPerPixelX
    End If
    
    lngWith = T_DrawClient.列单位 * Screen.TwipsPerPixelX
    
    With mshUpTab
        .Cols = lng列数 + 1
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        .Cell(flexcpText, 0, .FixedCols, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 0, .FixedCols, .Rows - 1, .Cols - 1) = ""
        .ColWidthMin = lngWith
        .RowHeightMin = T_BodyStyle.lng表格高度
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeRow(2) = True
        .ColWidth(0) = (T_DrawClient.刻度区域.Right - T_DrawClient.刻度区域.Left) * Screen.TwipsPerPixelX
        .RowHeight(-1) = .RowHeightMin
        .TextMatrix(0, 0) = Split(T_BodyStyle.str列头名称, "@")(0)
        If UBound(Split(T_BodyStyle.str列头名称, "@")) > 0 Then
            .TextMatrix(1, 0) = Split(T_BodyStyle.str列头名称, "@")(1)
        Else
            .TextMatrix(1, 0) = IIf(T_Patient.lng婴儿 = 0, "住 院 天 数", "出 生 天 数")
        End If
        If UBound(Split(T_BodyStyle.str列头名称, "@")) > 1 Then
            .TextMatrix(2, 0) = Split(T_BodyStyle.str列头名称, "@")(2)
        Else
            .TextMatrix(2, 0) = "手术后天数"
        End If
        If UBound(Split(T_BodyStyle.str列头名称, "@")) > 2 Then
            .TextMatrix(3, 0) = Split(T_BodyStyle.str列头名称, "@")(3)
        Else
            .TextMatrix(3, 0) = "时       间"
        End If
        
        For intCOl = 1 To .Cols - 1
            .ColWidth(intCOl) = lngWith
        Next
        .Redraw = flexRDBuffered
    End With
    
    '合并单元格的列
    For intRow = 0 To 2
        Call UniteCellCol(mshUpTab, T_BodyStyle.lng监测次数, intRow, mshUpTab.FixedCols)
    Next intRow
    
    If blnInitUpdate = True Then Call ShowUptab
    
    With vsf
        .Cols = 0
        .NewColumn "", 0, 1
        .NewColumn "项目", mshUpTab.ColWidth(0) + 10, 1
    
        For intCOl = 1 To lng列数
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
        .Cols = lng列数 + 4
        .Rows = 1
        .ColWidth(0) = mshUpTab.ColWidth(0)
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .MergeCells = flexMergeRestrictRows
        .MergeCol(0) = True
        .MergeRow(0) = True
        .Tag = 0
        .RowHeightMin = T_BodyStyle.lng下表格高度
        .RowHeight(-1) = T_BodyStyle.lng下表格高度
        
        For intCOl = .FixedCols To .Cols - 1
            .ColWidth(intCOl) = mshUpTab.ColWidth(1)
            If (intCOl - .FixedCols + 1) Mod 2 = 0 Then
                .Cell(flexcpBackColor, 0, intCOl, .Rows - 1, intCOl) = &H80000013
            End If
        Next intCOl

        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
    
    mItemNO.呼吸 = 0
    Dim bln呼吸 As Boolean
    Dim intRows  As Integer
    intRows = GetRows(bln呼吸, T_BodyItem.str表格项目)
    mintRepairRows = T_BodyStyle.lng表格空行 + intRows
    mbln显示皮试 = (Val(zlDatabase.GetPara("体温单显示皮试结果", glngSys, 1255, "0")) = 1)
    
    '检查呼吸是否是表格项目
    gstrSQL = "select 记录法 From 体温记录项目 where 项目序号=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "体温记录项目", gint呼吸)
    If rsTemp.RecordCount > 0 Then
         mintRepairRows = mintRepairRows - IIf(Val(Nvl(rsTemp!记录法)) = 2 And bln呼吸 = True, 1, 0)
    End If
    If mintRepairRows < 0 Then mintRepairRows = 0

    '加载所有表格项目，包括固定项目和有数据的活动项目
    Set rsTemp = GetAppendGridItemNew(T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng护理等级, T_Patient.lng婴儿, Int(CDate(mstr开始时间)), CDate(mstr结束时间), IIf(T_Patient.lng婴儿 = 0, 1, 2), T_Patient.lng科室ID, T_BodyItem.str表格项目, mblnMoved)
    With rsTemp
        If rsTemp.RecordCount > 0 Then
            rsTemp.MoveFirst
            mshDownTab.Rows = 0
            Call AppenGridItemNew(rsTemp)
        Else
            mshDownTab.Rows = 0
        End If
    End With
    
    mshDownTab.Rows = mintRepairRows
    mshDownTab.RowHeightMin = T_BodyStyle.lng下表格高度
    mshDownTab.RowHeight(-1) = mshDownTab.RowHeightMin
    
    '补充完剩下的空行
    If mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
        For intRow = Val(mshDownTab.Tag) To mshDownTab.Rows - 1

            mshDownTab.MergeRow(intRow) = True
            For intCOl = 0 To mshDownTab.FixedCols
                strPace = " " & String(intCOl, " ") & String(intRow, " ")
                mshDownTab.TextMatrix(intRow, intCOl) = strPace & "" & strPace
            Next intCOl
            
            Call UniteCellCol(mshDownTab, T_BodyStyle.lng监测次数, intRow, mshDownTab.FixedCols)
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
    If mshDownTab.Rows > mshDownTab.FixedRows Then
        For intCOl = mshDownTab.FixedCols To mshDownTab.Cols - 1
            If (intCOl - mshDownTab.FixedCols + 1) Mod 2 = 0 Then
                mshDownTab.Cell(flexcpBackColor, 0, intCOl, mshDownTab.Rows - 1, intCOl) = &HF7ECE6
            End If
        Next intCOl
        mshDownTab.Cell(flexcpAlignment, 0, 0, mshDownTab.Rows - 1, mshDownTab.Cols - 1) = 4
    End If
    Call picBack_Resize
    Call Paint_CanvasNew(mblnAutoAdjust) '初始化体温数据
    Call picBack_Resize
    
    PicRect.Top = 0
    PicRect.Left = 0
    PicRect.Right = picCommText.Width \ Screen.TwipsPerPixelX
    PicRect.Bottom = picCommText.Height \ Screen.TwipsPerPixelY
    picCommText.Cls
    Call PrintCurveInfo(picCommText, PicRect)
    
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
    Dim lngValue  As Long, intCOl As Integer
    Dim lngDays   As Long
    Dim i As Long, j As Long
    Dim lngColor  As Long
    Dim intMinCol As Integer, intMaxCol As Integer
    Dim strTmp As String
    Dim arrOperDay, strTmp1 As String
    Dim rsTmp  As New ADODB.Recordset
    Dim str时间 As String
    Dim intDays As Integer
    Dim lng次数 As Long
    Dim lngWith As Long
    Dim lng天数 As Long, lng频次 As Long, lng时间间隔 As Long
    Dim bln术后显示 As String
    Dim str结束时间 As String
    

    On Error GoTo Errhand
    
    lng天数 = T_BodyStyle.lng天数
    lng频次 = T_BodyStyle.lng监测次数
    lng时间间隔 = T_BodyStyle.lng时间间隔

    With mshUpTab
        
        lngValue = 0
        gstrSQL = "Select zl_CalcInDaysNew([1],[2],[3],[4]) As 开始天数 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取住院天数", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, Int(CDate(mstr开始时间)))

        If rsTmp.BOF = False Then
            lngValue = rsTmp("开始天数").Value
        End If
        
        '上表格式有单元格合并的，此处需要进行处理
        For intCOl = 1 To lng天数
            .ColData(intCOl) = 0
            .Row = 0
            .Col = intCOl
            .ColAlignment(intCOl) = 4

            strTmp = Format(CDate(mstr开始时间) + intCOl - 1, "yyyy-MM-dd")

            lngDays = lngValue + (intCOl - 1)
            
            For i = 1 To lng频次
                .Row = 0
                .Col = (intCOl - 1) * lng频次 + i
                
                If Right(strTmp, 5) = "01-01" Then
                    '一年的第一天
                    .Text = strTmp
                ElseIf strTmp = Format(mstrEnterDate, "yyyy-MM-dd") Then
                    '入院第一天，写上年份
                    .Text = strTmp
                ElseIf intCOl = 1 Then
                    '70299:刘鹏飞,2014-4-4,每页首列日期显示为年月日(1-年-月-日,0:默认格式:按规则显示)
                    If Val(zlDatabase.GetPara("首列日期格式", glngSys, 1255, "0")) = 1 Then
                        .Text = strTmp
                    Else
                        .Text = Right(strTmp, 5)
                    End If
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
            Call CalcMinMaxColNew(picMain.Tag, intMinCol, intMaxCol)
            mintColMin = intMinCol
            mintColMax = intMaxCol
            
            With picDisplay
                .Left = ((((intMaxCol - 1) \ lng频次) + 1) * lng频次 - 1) * mshUpTab.ColWidth(intMinCol) + mshUpTab.ColWidth(0)
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

        ReDim mstrOpValue(T_BodyStyle.lng天数) As String
        ReDim mstrOpdays(T_BodyStyle.lng天数) As String
        
        For i = 1 To T_BodyStyle.lng天数
            mstrOpValue(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng监测次数 + 1))
            mstrOpdays(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng监测次数 + 1))
        Next i
        
        '提取输入标志天数和停止手术标志
        mintOpDays = Val(zlDatabase.GetPara("手术后标注天数", glngSys, 1255, "10"))
        mblnStopFlag = (Val(zlDatabase.GetPara("再次手术停止前次标注", glngSys, 1255, "0")) = 1)
        bln术后显示 = (Val(zlDatabase.GetPara("病人术后不足14天出院标记显示", glngSys, 1255, "0")) = 1)
        
        '51338,刘鹏飞,2012-07-06
        strTmp = zlDatabase.GetPara("手术当天缺省格式", glngSys, 1255, "2")
        If Val(strTmp) >= 0 And Val(strTmp) <= 3 Then
            mintOpFormat = Val(strTmp)
        Else
            mintOpFormat = 0
        End If
        
        strTmp = ""
        '显示但前段的手术标记
        gstrSQL = "select B.发生时间 时间" & _
            "   From 病人护理文件 A,病人护理数据 B,病人护理明细 C" & _
            "   where A.ID=B.文件ID And  B.ID=C.记录ID And A.ID=[1] And nvl(A.婴儿,0)=[4]" & _
            "   and A.病人ID=[2] and A.主页ID=[3] and C.记录类型=4 And NVL(C.复试合格,0)<>1 and C.终止版本 is null" & _
            "   and B.发生时间 between [5] and [6] order by B.发生时间"
        If mblnMoved Then
            gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
            gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
            gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取手术标记", Val(T_Patient.lng文件ID), T_Patient.lng病人ID, T_Patient.lng主页ID, Val(T_Patient.lng婴儿), Int(CDate(mstr开始时间) - 14), CDate(mstr结束时间))

        str结束时间 = mstr结束时间
        
        Do While Not rsTmp.EOF
            str时间 = Format(rsTmp("时间"), "YYYY-MM-DD")
            
             '问题号:56005,李涛,2013-04-27
            If Not rsTmp.EOF Then
                If bln术后显示 And DateDiff("d", CDate(Format(str时间, "YYYY-MM-DD")), str结束时间) < mintOpDays Then
                    str结束时间 = Format(DateAdd("D", mintOpDays, CDate(Format(str时间, "YYYY-MM-DD"))), "YYYY-MM-DD") & " 23:59:59"
                End If
            End If
            
            For i = 1 To lng天数
                If DateDiff("d", mstr开始时间, str结束时间) + 1 >= i Then
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
                                    If mintOpFormat = 3 Then
                                        mstrOpValue(i) = mstrOpValue(i) & "/" & intDays
                                    Else
                                        mstrOpValue(i) = intDays & "/" & mstrOpValue(i)
                                    End If
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
            "   And A.病人ID=[2] And A.主页ID=[3] And C.记录类型=4 And NVL(C.复试合格,0)<>1 and C.终止版本 is null" & _
            "   And B.发生时间 <[5] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取手术标记", Val(T_Patient.lng文件ID), T_Patient.lng病人ID, T_Patient.lng主页ID, Val(T_Patient.lng婴儿), Int(CDate(mstr开始时间)))
        
        If mblnMoved Then
            gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
            gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
            gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
        End If
        lng次数 = 0
        If rsTmp.BOF = False Then lng次数 = Val(rsTmp("次数"))
        For i = 1 To lng天数
            If DateDiff("d", mstr开始时间, str结束时间) + 1 >= i Then
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
                        '问题号:57771,李涛，2013-05-02
                        If mintOpFormat = 3 Then
                            strTmp1 = Switch(lng次数 = 1, "术日", lng次数 = 2, "术2", lng次数 = 3, "术3", lng次数 = 4, "术4", lng次数 = 5, "术5", lng次数 = 6, "术6", lng次数 = 7, "术7", lng次数 = 8, "术8", lng次数 = 9, "术9", lng次数 = 10, "术10", lng次数 = 11, "术11", lng次数 = 12, "术12")
                        Else
                            strTmp1 = Switch(lng次数 = 1, "Ⅰ", lng次数 = 2, "Ⅱ", lng次数 = 3, "Ⅲ", lng次数 = 4, "Ⅳ", lng次数 = 5, "Ⅴ", lng次数 = 6, "Ⅵ", lng次数 = 7, "Ⅶ", lng次数 = 8, "Ⅷ", lng次数 = 9, "Ⅸ", lng次数 = 10, "Ⅹ", lng次数 = 11, "Ⅺ", lng次数 = 12, "Ⅻ")
                        End If
                       
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
                                mstrOpValue(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng监测次数 + 1)) & "0" & .TextMatrix(2, ((i - 1) * T_BodyStyle.lng监测次数 + 1))
                            Case 2 '--显示次数
                                If strTmp = "Ⅰ" Then
                                    mstrOpValue(i) = 0
                                Else
                                    mstrOpValue(i) = strTmp & "-0"
                                End If
                            Case 3
                                  If strTmp = "术日 1" Then
                                    mstrOpValue(i) = "术日"
                                Else
                                    mstrOpValue(i) = strTmp
                                End If
                            Case Else '--不显示
                                 mstrOpValue(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng监测次数 + 1))
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
                            Case 3
                                  If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = mstrOpValue(i) & "/" & strTmp
                                Else
                                    mstrOpValue(i) = strTmp
                                End If
                            Case Else  '--不显示
                                If Trim(mstrOpValue(i)) <> "" Then
                                    mstrOpValue(i) = mstrOpValue(i)
                                Else
                                    mstrOpValue(i) = .TextMatrix(2, ((i - 1) * T_BodyStyle.lng监测次数 + 1))
                                End If
                        End Select
                    End If
                    .Row = 2
                    For j = 1 To T_BodyStyle.lng监测次数
                        .Col = j + (i - 1) * T_BodyStyle.lng监测次数
                        .Text = mstrOpValue(i)
                    Next j
                Else
                    .Row = 2
                    For j = 1 To T_BodyStyle.lng监测次数
                        .Col = j + (i - 1) * T_BodyStyle.lng监测次数
                        .Text = mstrOpValue(i)
                    Next j
                End If
            End If
        Next i
        '设定日期，住院天数文本颜色
        mshUpTab.Cell(flexcpForeColor, 0, mshUpTab.FixedCols, 1, mshUpTab.Cols - 1) = 16711680
        '设定手术 分娩文本颜色
        '51283,刘鹏飞,2012-07-11
        lngColor = Val(zlDatabase.GetPara("手术天数显示颜色", glngSys, 1255, "255"))
        mshUpTab.Cell(flexcpForeColor, 2, mshUpTab.FixedCols, 2, mshUpTab.Cols - 1) = lngColor

        lngWith = T_DrawClient.列单位 * Screen.TwipsPerPixelX
        mshUpTab.ColWidthMin = lngWith
        'mshUpTab.Cell(flexcpWidth, 0, 1, mshUpTab.Rows - 1, mshUpTab.Cols - 1) = lngWith
        For intCOl = 1 To mshUpTab.Cols - 1
            mshUpTab.ColWidth(intCOl) = lngWith
        Next intCOl
        
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
    Dim strItems As String, strItemName As String, strSQL As String
    Dim lngItemCode As Long
    Dim strPace As String
    Dim str项目名称 As String, str项目名称1 As String
    Dim int记录频次 As Integer, int项目性质 As Integer, int项目类型 As Integer, int项目表示 As Integer, strTabItemTemp As String
    Dim strBegin As String, str结果 As String, strPart As String
    Dim int舒张压 As Integer, int收缩压 As Integer, Int列号 As Integer
    Dim blnColor As Boolean
    Dim lngColor As Long
    Dim arrTmpString0() As String, arrTmpString1() As String, arrTmpString2() As String
    Dim blnAdd As Boolean, blnValue As Boolean
    Dim SinX As Single
    Dim i As Integer, j As Integer
    Dim int呼吸位置 As Integer, intValue As Integer, int呼吸表格输出格式 As Integer
    Dim bln汇总当天 As Boolean, bln录入小时 As Boolean
    Dim arrTmp() As String
    Dim dtBegin As Date, dtEnd As Date
    Dim int监测次数 As Integer
    '73316问题修改相关变量声明
    Dim arrBreathe, blnBreathe As Boolean, intBegin As Integer, intEnd As Integer
    Dim lngX As Long, lngY As Long, lngBottomY As Long
    Dim blnBreatheShowType As Boolean  '参数:呼吸为表格时呼吸机输出方式
    
    On Error GoTo Errhand
    
    ReDim mstrNewString(mintRepairRows, T_BodyStyle.lng天数 - 1)
    ReDim mlngNewHeight(mintRepairRows)
    For i = 0 To UBound(mlngNewHeight)
        mlngNewHeight(i) = mshDownTab.RowHeightMin
    Next i
    ReDim arrTmpString0(1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数) As String
    ReDim arrTmpString1(1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数) As String
    ReDim arrTmpString2(1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数) As String
    
    'mstrNewString = Split(String(T_BodyStyle.lng天数-1, ";"), ";")
    int呼吸表格输出格式 = zlDatabase.GetPara("呼吸表格输出", glngSys, 1255, 0)
    If int呼吸表格输出格式 < 0 Or int呼吸表格输出格式 > 3 Then int呼吸表格输出格式 = 0
    bln汇总当天 = (Val(zlDatabase.GetPara("汇总波动显示当天数据", glngSys, 1255, 0)) = 1)
    mbln灌肠大便分子分母显示 = (Val(zlDatabase.GetPara("灌肠后大便显示格式", glngSys, 1255, 0)) = 1)
    '--51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H)
    bln录入小时 = (Val(zlDatabase.GetPara("全天汇总显示录入时间", glngSys, 1255, 0)) = 1)
    '73316:刘鹏飞,2014-06-26,重庆部分医院要求:
    '（1）呼吸用蓝色笔在呼吸栏相应时间内填写，相邻两次呼吸上下交错填写，先上后下
    '（2）辅助呼吸标识，在起始相应时间用蓝色钢笔在体温单呼吸栏横线上方纵向
    '填写“呼吸机”，用“↑”标识开始，终止以“↓”标识；呼吸机设定频率以数字表示，用蓝
    '色笔在呼吸栏相应时间内填写，相邻两次呼吸上下交错填写，先上后
    '2----开始输出呼吸数据 呼吸机为图形输出
    blnBreatheShowType = (Val(zlDatabase.GetPara("呼吸表格呼吸机输出方式", glngSys, 1255, 0)) = 1)
    
    int监测次数 = T_BodyStyle.lng监测次数
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
            strItemName = mshDownTab.TextMatrix(intRow, 3)
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
        "   Where B.ID=A.文件ID And A.ID = C.记录ID   AND B.ID=[1] AND Nvl(B.婴儿,0)=[7] " & _
        "   AND B.病人id=[2]  AND B.主页id=[3] AND INSTR([6],decode(E.项目性质,2,C.体温部位 || D.记录名 ,D.记录名))>0 " & _
        "   AND D.项目序号=C.项目序号  AND MOD(c.记录类型,10)=1  AND E.项目序号=D.项目序号 " & _
        "   AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null AND 记录法=2"
    
    '提取非体温表格的汇总项目
    strSQL = "  SELECT C.ID,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & _
        "   Decode(d.项目性质, 2, c.体温部位 || d.项目名称, d.项目名称) 项目名称,D.项目序号,C.来源ID,C.共用,D.项目性质" & _
        "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,(SELECT A.项目序号,A.项目名称, A.项目性质,B.父序号 FROM 护理记录项目 A,护理汇总项目 B" & vbNewLine & _
        "       WHERE A.项目序号=B.序号 AND  B.父序号 is not NULL " & vbNewLine & _
        "       AND NVL(A.应用方式,0)=1 AND NVL(A.护理等级,0)>=[8] AND NVL(A.适用病人,0) IN (0,[9])" & vbNewLine & _
        "       AND (A.适用科室=1 OR (A.适用科室=2 AND EXISTS (SELECT 1 FROM 护理适用科室 D WHERE D.项目序号=A.项目序号 AND D.科室ID=[10])))) D" & _
        "   Where B.ID=A.文件ID And A.ID = C.记录ID AND Instr([6], Decode(d.项目性质, 2, c.体温部位 || d.项目名称, d.项目名称)) = 0  AND B.ID=[1]  AND Nvl(B.婴儿,0)=[7] " & _
        "   AND B.病人id=[2]  AND B.主页id=[3]  AND D.项目序号=C.项目序号  AND C.记录类型=1" & _
        "   AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null"
        
    gstrSQL = "Select /*+ Rule*/ ID,时间,记录类型,显示,结果,体温部位,未记说明,数据来源,项目名称,项目序号,来源ID,共用,项目性质 From (" & _
        "   " & gstrSQL & " UNION ALL " & strSQL & ")" & _
        "   Order By  Decode(项目名称,'收缩压',0,1)," & strItems & ",时间"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取体温表格数据", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, _
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
                intCOl = GetCurveColumnNew(rsTemp!时间, mstr开始时间, gintHourBegin) + vsf.FixedCols - 1
                str结果 = zlCommFun.Nvl(rsTemp!结果) & ";" & Nvl(rsTemp!体温部位)
                If intCOl < vsf.Cols Then
                    If arrTmpString1(intCOl - vsf.FixedCols + 1) <> "" Then
                        If (Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) <> 1 And Val(zlCommFun.Nvl(!显示, 0)) <> 1) Or _
                            (Val(arrTmpString2(intCOl - vsf.FixedCols + 1)) = 1 And Val(zlCommFun.Nvl(!显示, 0)) = 1) Then
                            
                            '检查那个离重点时间更近
                            SinX = GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"))
                            blnAdd = GetCanvasCenterNew(CDate(Format(arrTmpString1(intCOl - vsf.FixedCols + 1), "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
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
    For i = 1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数
        If Val(arrTmpString2(i)) = 2 Then arrTmpString0(i) = ""
    Next i
    
    '2----开始输出呼吸数据 呼吸机为图形输出
    int呼吸位置 = 0
    blnValue = False
    arrBreathe = Array(): blnBreathe = False
     '循环输出呼吸值
    vsf.Cell(flexcpForeColor, 1, vsf.FixedCols, 1, vsf.Cols - 1) = Val(vsf.Tag)
    If blnBreatheShowType = True Then
        vsf.Body.OwnerDraw = flexODNone
    Else
        vsf.Body.OwnerDraw = flexODOver
    End If
    
    For i = 1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数
        intCOl = i + vsf.FixedCols - 1
        If InStr(1, arrTmpString0(i), ";") > 0 Then
            str结果 = Split(arrTmpString0(i), ";")(0)
            strPart = Split(arrTmpString0(i), ";")(1)
        Else
            str结果 = arrTmpString0(i)
            strPart = ""
        End If
        '整理每段呼吸机的一个范围
        If mbln呼吸曲线 = False Then
            If strPart = "呼吸机" And IsNumeric(str结果) Then
                If blnBreathe = False Then
                    ReDim Preserve arrBreathe(UBound(arrBreathe) + 1)
                    arrBreathe(UBound(arrBreathe)) = i & ";" & i
                    blnBreathe = True
                Else
                    arrBreathe(UBound(arrBreathe)) = Split(arrBreathe(UBound(arrBreathe)), ";")(0) & ";" & i
                End If
            Else
                blnBreathe = False
            End If
        End If
        '打印呼吸值（间隔错开打印） 第一行始终在上面
        If IsNumeric(str结果) Then
            vsf.TextMatrix(1, intCOl) = str结果
            If blnValue = False Then
                intValue = IIf(intCOl Mod 2 = 0, 0, 1)
                blnValue = True
                int呼吸位置 = 2
            End If
            
            If int呼吸表格输出格式 = 0 Or int呼吸表格输出格式 = 2 Then '顺序上下显示
                If intCOl Mod 2 = intValue Then
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = IIf(int呼吸表格输出格式 = 0, flexAlignCenterTop, flexAlignCenterBottom)
                    If strPart <> "呼吸机" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = IIf(int呼吸表格输出格式 = 0, 1, 2)
                    End If
                Else
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = IIf(int呼吸表格输出格式 = 0, flexAlignCenterBottom, flexAlignCenterTop)
                    If strPart <> "呼吸机" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = IIf(int呼吸表格输出格式 = 0, 2, 1)
                    End If
                End If
                
            Else        '有数据时数据之间上下显示
                If int呼吸位置 = 2 Then
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = IIf(int呼吸表格输出格式 = 1, flexAlignCenterTop, flexAlignCenterBottom)
                    If strPart <> "呼吸机" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = IIf(int呼吸表格输出格式 = 1, 1, 2)
                    End If
                Else
                    vsf.Cell(flexcpAlignment, 1, intCOl, 1, intCOl) = IIf(int呼吸表格输出格式 = 1, flexAlignCenterBottom, flexAlignCenterTop)
                    If strPart <> "呼吸机" Then
                        vsf.ColData(intCOl) = 0
                    Else
                        vsf.ColData(intCOl) = IIf(int呼吸表格输出格式 = 1, 2, 1)
                    End If
                End If
                
                int呼吸位置 = int呼吸位置 + 1
                If int呼吸位置 > 2 Then int呼吸位置 = 1
            End If
        End If
    Next i
    
    '开始在呼吸栏上方输出呼吸
    If blnBreatheShowType = True Then
        lngBottomY = T_DrawClient.曲线总区域.Bottom
        For i = 0 To UBound(arrBreathe)
            intBegin = Split(arrBreathe(i), ";")(0)
            intEnd = Split(arrBreathe(i), ";")(1)
            '输出呼吸机文字
            strPart = "呼吸机"
            Call SetTextColor(mlngMemDC, Val(vsf.Tag))
            T_Size.H = mobjDraw.TextHeight("呼") / T_TwipsPerPixel.Y
            T_Size.W = mobjDraw.TextWidth("呼") / T_TwipsPerPixel.X
            '由于GetTextRect函数默认给X+1所以此处-1
            If intBegin = intEnd Then
                If T_DrawClient.列单位 >= T_Size.W + 6 Then
                    lngX = T_DrawClient.体温区域.Left + (intBegin - 1) * T_DrawClient.列单位 + ((T_DrawClient.列单位 - T_Size.W - 6) \ 2) - 1
                Else
                    lngX = T_DrawClient.体温区域.Left + (intBegin - 1) * T_DrawClient.列单位 - ((T_Size.W + 6 - T_DrawClient.列单位)) - 1
                End If
            Else
                If T_DrawClient.列单位 >= T_Size.W + 3 Then
                    lngX = T_DrawClient.体温区域.Left + (intBegin - 1) * T_DrawClient.列单位 + ((T_DrawClient.列单位 - T_Size.W - 3) \ 2) - 1
                Else
                    lngX = T_DrawClient.体温区域.Left + (intBegin - 1) * T_DrawClient.列单位 - ((T_Size.W + 3 - T_DrawClient.列单位)) - 1
                End If
            End If
            lngY = lngBottomY - T_Size.H * Len(strPart)
            For j = 1 To Len(strPart)
                Call GetTextRect(mobjDraw, lngX, lngY, Mid(strPart, j, 1), 0, False)
                Call DrawText(mlngMemDC, Mid(strPart, j, 1), -1, T_LableRect, DT_CENTER)
                lngY = lngY + T_Size.H
            Next j
            '开始画向上箭头，结束画向下箭头
            If intBegin = intEnd Then
                lngY = T_Size.H * Len(strPart) - T_Size.H
                lngX = lngX + T_Size.W + 3
                Call DrawLine(mlngMemDC, lngX, lngBottomY - lngY - (T_Size.H \ 2), lngX, lngBottomY - (T_Size.H \ 2), PS_SOLID, 1, Val(vsf.Tag), True)
                lngX = lngX + 3
                Call DrawLine(mlngMemDC, lngX, lngBottomY - (T_Size.H \ 2), lngX, lngBottomY - lngY - (T_Size.H \ 2), PS_SOLID, 1, Val(vsf.Tag), True)
            Else
                lngY = T_Size.H * Len(strPart) - T_Size.H
                lngX = lngX + T_Size.W + 3
                Call DrawLine(mlngMemDC, lngX, lngBottomY - lngY - (T_Size.H \ 2), lngX, lngBottomY - (T_Size.H \ 2), PS_SOLID, 1, Val(vsf.Tag), True)
                lngX = T_DrawClient.体温区域.Left + (intEnd - 1) * T_DrawClient.列单位 + T_DrawClient.列单位 \ 2
                Call DrawLine(mlngMemDC, lngX, lngBottomY - (T_Size.H \ 2), lngX, lngBottomY - lngY - (T_Size.H \ 2), PS_SOLID, 1, Val(vsf.Tag), True)
            End If
        Next i
    End If
    
    'Debug.Print "数据开始---" & Now
    '提取表格项目数据信息
    With mshDownTab
        lngItemCode = 0
        str项目名称 = ""
        For intRow = .FixedRows To .Tag - 1
            str项目名称1 = .TextMatrix(intRow, 3)
            blnColor = False
            If str项目名称1 & ";" & .RowData(intRow) <> str项目名称 & ";" & lngItemCode Then
                
                lngItemCode = .RowData(intRow)
                str项目名称 = str项目名称1
                int项目类型 = Val(Split(.TextMatrix(intRow, 1), ",")(0))
                int记录频次 = Val(Split(.TextMatrix(intRow, 1), ",")(2))
                int项目表示 = Val(Split(.TextMatrix(intRow, 1), ",")(3))
                int项目性质 = Val(Split(.TextMatrix(intRow, 1), ",")(4))
                strTabItemTemp = Val(Split(.TextMatrix(intRow, 1), ",")(6)) & ";" & Split(.TextMatrix(intRow, 1), ",")(7)
                blnColor = (int项目性质 = 2 And int项目类型 = 1 And int项目表示 = 0)
                
                For intDay = 0 To T_BodyStyle.lng天数 - 1
                    strBegin = DateAdd("D", intDay, CDate(mstr开始时间))
                    If CDate(strBegin) > CDate(mstr结束时间) Then strBegin = mstr结束时间
                    int舒张压 = 0
                    int收缩压 = 0
                    Int列号 = 0
                    '循环得到某个项目某天的数据信息
                    Set rsDownTab = ReturnItemRecord(rsTemp, Int(CDate(strBegin)), CDate(mstrEnterDate), lngItemCode & ";" & str项目名称 & ";" & _
                                int记录频次 & ";" & int项目表示 & ";" & int项目性质 & ";" & strTabItemTemp, bln汇总当天, bln录入小时)
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
                                intCOl = intDay * int监测次数 + .FixedCols
                                intColCount = int监测次数
                                strPace = " "
                            Case 2
                                intRow1 = intRow
                                intCOl = (intCOl - 1) * (int监测次数 / 2) + intDay * int监测次数 + .FixedCols
                                intColCount = (int监测次数 / 2)
                                strPace = String(intCOl, " ")
                            Case 3
                                intRow1 = intRow + (intCOl - 1)
                                intCOl = intDay * int监测次数 + .FixedCols
                                intColCount = int监测次数
                                strPace = " "
                            Case 4
                                intRow1 = intRow + Fix((intCOl - 1) / 2)
                                Select Case intCOl
                                    Case 1, 3
                                        intCOl = 1
                                    Case 2, 4
                                        intCOl = 2
                                End Select
                                intCOl = (intCOl - 1) * (int监测次数 / 2) + intDay * int监测次数 + .FixedCols
                                intColCount = int监测次数 / 2
                                strPace = String(intCOl, " ")
                            Case 6
                                intRow1 = intRow
                                intCOl = (intCOl - 1) * (int监测次数 / 6) + intDay * int监测次数 + .FixedCols
                                intColCount = int监测次数 / 6
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
                                                    mrsCurInfo.Filter = "名称='" & str结果 & "'"
                                                    If Not mrsCurInfo.EOF Then
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
        lngColor = RGB(0, 0, 255)
        If mbln显示皮试 = True And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 Then
            '83477:LPF,皮试结果提取SQL修正
            strSQL = _
                " Select 时间, f_List2str(Cast(Collect(药物名) As t_Strlist)) 药物名" & vbNewLine & _
                " From (Select To_Char(a.开始执行时间, 'YYYY-MM-DD') 时间," & vbNewLine & _
                "              Decode(皮试结果, '(+)', 255, '(阳性)', 255, " & lngColor & ") || '-#' ||" & vbNewLine & _
                "               Replace(Replace(Replace(Decode(b.试管编码, Null, a.医嘱内容, b.试管编码), ',', ''), '-#', ''), '皮试', '') || a.皮试结果 药物名" & vbNewLine & _
                "       From 病人医嘱记录 a, 诊疗项目目录 b" & vbNewLine & _
                "       Where a.诊疗项目id = b.Id And a.诊疗类别 = 'E' And b.操作类型 = '1' And a.医嘱状态 = 8 And a.皮试结果 Is Not Null And a.皮试结果 <> '免试' And" & vbNewLine & _
                "             a.病人id = [1] And a.主页id = [2] And a.婴儿 = [3] And a.开始执行时间 Between [4] And [5]" & vbNewLine & _
                "       Order By a.开始执行时间, a.皮试结果)" & vbNewLine & _
                " Group By 时间"

            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            End If

            Set rsDownTab = zlDatabase.OpenSQLRecord(strSQL, "提取病人过敏记录信息", T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿, CDate(mstr开始时间), CDate(mstr结束时间))

            Do While Not rsDownTab.EOF
                intCOl = DateDiff("D", CDate(Format(mstr开始时间, "YYYY-MM-DD")), CDate(Format(rsDownTab!时间, "YYYY-MM-DD")))
                str结果 = Nvl(rsDownTab!药物名)
                Call ShowTestis(str结果, intCOl)
                rsDownTab.MoveNext
            Loop
            
            '重新整理皮试结果行高
            For intRow = Val(mshDownTab.Tag) To mintRepairRows - 1
                If mshDownTab.RowHeight(intRow) < mlngNewHeight(intRow - Val(mshDownTab.Tag)) Then
                    mshDownTab.RowHeight(intRow) = mlngNewHeight(intRow - Val(mshDownTab.Tag))
                End If
            Next intRow
            Call picBack_Resize
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
    Dim intNum As Integer, i As Integer, j As Integer
    Dim lngColor As Long
    Dim strTmp As String, strPart As String, strSpace As String
    Dim arrTmp() As String, arrData
    Dim LPoint As T_LPoint
    Dim lngDc As Long
    Dim objDraw As Object
    Dim lngH As Long, lngW As Long, lngX1 As Long, lngLen As Long
    Dim intRowCount As Integer
    Dim sngLen As Single
    Dim intRow As Integer
    Dim sgnSize As Single, strFontName As String
    Dim lngRowHeight As Long
    
    Set objDraw = picBack
    intRowCount = Val(mshDownTab.Tag)
    intNum = 1
    strTmp = strValue
    If strTmp = "" Then Exit Sub
    LPoint.X = 0
    LPoint.W = (mshDownTab.ColWidth(mshDownTab.FixedCols) / Screen.TwipsPerPixelX) * T_BodyStyle.lng监测次数
    lngW = LPoint.W
    lngX1 = 0
    intRow = 0
    
    '开始计算是否需要换行
    strPart = ""
    arrTmp = Split(strTmp, ",")
    For i = LBound(arrTmp) To UBound(arrTmp)
        lngColor = Val(Split(arrTmp(i), "-#")(0))
        strTmp = Replace(CStr(Split(arrTmp(i), "-#")(1)), vbCrLf, "") '皮试结果
        If Trim(strTmp) <> "" Then
            strFontName = "宋体"
            sgnSize = GetFontSize(objDraw, strTmp & "L", LPoint.W)
            '缩小字体后输出要输出的实际行数
            With txtLength
                .Width = LPoint.W * Screen.TwipsPerPixelX
                .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                .FontName = "宋体"
                .FontSize = sgnSize
                .FontBold = False
                .FontItalic = False
            End With
            arrData = GetData(frmTendFileRead.txtLength.Text, txtLength)
            
            '计算每一行皮试结果的最大行高
            If Val(objDraw.TextHeight("刘") * (UBound(arrData) + 1)) > mshDownTab.RowHeightMin Then
                lngRowHeight = objDraw.TextHeight("刘") * (UBound(arrData) + 1)
            Else
                lngRowHeight = mshDownTab.RowHeightMin
            End If
            If mlngNewHeight(intRow) < lngRowHeight Then mlngNewHeight(intRow) = lngRowHeight
            '皮试结果存在多条表头合并显示
            If mshDownTab.Rows > intRow + Val(mshDownTab.Tag) Then
                mshDownTab.MergeRow(intRow) = True
                strSpace = " " & String(1, " ") & String(Val(mshDownTab.Tag), " ")
                mshDownTab.TextMatrix(intRow + Val(mshDownTab.Tag), 0) = strSpace & "皮试结果" & strSpace
                
                For j = 0 To T_BodyStyle.lng监测次数 - 1
                    strSpace = " " & String(intCOl + 1, " ") & String(intRow + Val(mshDownTab.Tag), " ")
                    If intCOl * T_BodyStyle.lng监测次数 + mshDownTab.FixedCols + j < mshDownTab.Cols Then
                        mshDownTab.TextMatrix(intRow + Val(mshDownTab.Tag), intCOl * T_BodyStyle.lng监测次数 + mshDownTab.FixedCols + j) = strSpace & strTmp & strSpace
                    End If
                Next j
            End If
            '开始输出内容
            mstrNewString(intRow, intCOl) = sgnSize & "'" & strFontName & "'" & lngColor & "-#" & strTmp
            
            intRow = intRow + 1
            intNum = intNum + 1
            If intRowCount + intNum > mintRepairRows Then Exit Sub
        End If
    Next i
End Sub

Public Sub AppenGridItemNew(ByVal rsTemp As ADODB.Recordset)
     '填写表格标题
    Dim intRow  As Integer, intRowStart As Integer
    Dim int频次 As Integer
    Dim intRowNum As Integer, intColNum As Integer
    Dim intRowCount As Integer, intNum As Integer
    Dim i As Integer, j As Integer
    Dim strText As String, str值域 As String
    Dim int监测次数 As Long
    Dim strArray() As String


    On Error GoTo Errhand
    int监测次数 = T_BodyStyle.lng监测次数
    
    
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
                        
                        'intColNum 要合并的列数
                        'intRowNum 要合并的行
                        Select Case int频次
                            'intColNum 要合并的列数
                            'intRowNum 要合并的行
                            Case 1
                                intRowNum = 1
                                intColNum = int监测次数
                            Case 2
                                intRowNum = 1
                                intColNum = int监测次数 / 2
                            Case 3
                                intRowNum = 3
                                intColNum = int监测次数
                            Case 4
                                intRowNum = 2
                                intColNum = int监测次数 / 2
                            Case 6
                                intRowNum = 1
                                intColNum = int监测次数 / 6
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
                            
                            mshDownTab.TextMatrix(intRow, 1) = zlCommFun.Nvl(!项目类型) & "," & zlCommFun.Nvl(!项目小数) & "," & _
                                int频次 & "," & zlCommFun.Nvl(!项目表示) & "," & zlCommFun.Nvl(!项目性质) & "," & zlCommFun.Nvl(!项目长度) & "," & zlCommFun.Nvl(!入院首测, 0) & "," & Nvl(!体温部位)
                            mshDownTab.TextMatrix(intRow, 2) = zlCommFun.Nvl(!最小值, "") & ";" & zlCommFun.Nvl(!最大值, "")
                            mshDownTab.TextMatrix(intRow, 3) = Nvl(!记录名)
                            
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
    Dim lngIndex As Long
    Select Case Control.Id
        Case conMenu_View_Jump '菜单
            mcbrToolBar页面.Caption = Control.Category
            mstrParam = Control.Parameter
            Call InitWeekDays(mstrParam)
            Call zlMenuClick("装载数据", mstrParam)
            cbsMain.RecalcLayout
        Case conMenu_View_OneWeek To conMenu_View_FourWeek '4个周期按钮
            mstrParam = Control.Parameter
            Call InitWeekDays(mstrParam)
            Call zlMenuClick("装载数据", mstrParam)
            lngIndex = GetMenuPageIndex(0)
            mcbrToolBar页面.Caption = mcbrItem.Controls.Item(lngIndex).Category
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
    
    If T_Patient.lng文件ID = 0 Then Exit Sub
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
    Dim lngDc As Long, lngFont As Long, lngOldFont As Long
    Dim objDraw As Object, stdSet As Object
    Dim intCOl As Integer, intRow As Integer
    Dim sgnSize As Single, strFontName As String
    Dim arrData
    
    On Error Resume Next
    Err = 0
    intRow = UBound(mstrNewString)
    If Err <> 0 Then Exit Sub
    
    On Error GoTo Errhand
    
    lngDc = hDC
    Set objDraw = picBack
    If mbln显示皮试 = True And mintRepairRows > Val(mshDownTab.Tag) And mintRepairRows > 0 And Col >= mshDownTab.FixedCols And Row >= Val(mshDownTab.Tag) Then
        If (Col - mshDownTab.FixedCols) Mod T_BodyStyle.lng监测次数 = 0 And UBound(mstrNewString) >= (Row - Val(mshDownTab.Tag)) Then
            intCOl = (Col - mshDownTab.FixedCols) / T_BodyStyle.lng监测次数
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
            LPoint.Y = Top
            
            '1、清空内容
            '创建与背景色相同的刷子
            lngBackColor = GetRBGFromOLEColor(mshDownTab.BackColor)
            lngBrush = CreateSolidBrush(lngBackColor)
            '使用该刷子填充背景色
            lngOldBrush = SelectObject(lngDc, lngBrush)
            Call FillRect(hDC, T_ClientRect, lngBrush)
            '立即销毁临时使用的刷子并还原刷子
            Call SelectObject(lngDc, lngOldBrush)
            Call DeleteObject(lngBrush)
        
            sgnSize = 9: strFontName = "宋体"
            If UBound(Split(strTmp, "'")) > 0 Then
                sgnSize = Split(strTmp, "'")(0)
                strFontName = Split(strTmp, "'")(1)
                strTmp = Split(strTmp, "'")(2)
            End If
            
            arrTmp = Split(strTmp, "-#")
            lngColor = Val(arrTmp(0))
            strTmp = arrTmp(1)
            
            With txtLength
                .Width = mshDownTab.ColWidth(mshDownTab.FixedCols) * T_BodyStyle.lng监测次数
                .Text = Replace(Replace(Replace(strTmp, Chr(10), ""), Chr(13), ""), Chr(1), "")
                .FontName = "宋体"
                .FontSize = sgnSize
                .FontBold = False
                .FontItalic = False
            End With
            arrData = GetData(frmTendFileRead.txtLength.Text, txtLength)
            
            '创建字体
            Set stdSet = New StdFont
            stdSet.Name = strFontName
            stdSet.Size = sgnSize
            stdSet.Bold = False
            Call SetFontIndirect(stdSet, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            
            If (UBound(arrData) + 1) * objDraw.TextHeight("刘") < mshDownTab.RowHeight(Row) Then
                LPoint.Y = Top + (Val(mshDownTab.RowHeight(Row)) - ((UBound(arrData) + 1) * objDraw.TextHeight("刘"))) / T_TwipsPerPixel.Y / 2
            Else
                LPoint.Y = Top
            End If
            
            '开始输出内容
            Call SetTextColor(lngDc, lngColor)
            For i = 0 To UBound(arrData)
                Call GetTextRect(objDraw, LPoint.X, LPoint.Y, CStr(arrData(i)), , False)
                Call DrawText(lngDc, CStr(arrData(i)), -1, T_LableRect, DT_CENTER)
                LPoint.Y = LPoint.Y + Format(objDraw.TextHeight("刘") / T_TwipsPerPixel.Y, "#0")
            Next i
                    
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(objDraw)
        End If
    End If
    
    '输出大便次数
    If Col >= mshDownTab.FixedCols And Row >= mshDownTab.FixedRows Then
        strTmp = mshDownTab.TextMatrix(Row, Col)
        If AnsyGrade(Val(mshDownTab.RowData(Row)), strTmp, arrText) = True Then
            'lngColor = mshDownTab.Cell(flexcpForeColor, Row, Col, Row, Col)
            Call DrawDownTabAnsyGrade(lngDc, picMain, arrText, Row, Col, Left, Top, Right, Bottom, Done, mbln灌肠大便分子分母显示)
        End If
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mshDownTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RaiseShowTipInfo(mshDownTab, 3, X, Y)
End Sub

Private Sub mshUpTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim strTime As String
    If NewRow = 0 And T_Patient.lng编辑 = 1 Then
        strTime = GetCurveDateNew(NewCol, mstr开始时间, gintHourBegin)
        If Format(Split(strTime, ";")(0), "YYYY-MM-DD") > Format(mstr结束时间, "YYYY-MM-DD") Then
            mshUpTab.FocusRect = flexFocusLight
        Else
            mshUpTab.FocusRect = flexFocusSolid
            If mblnKeyDown = True Then
                picDisplay.Left = ((((NewCol - 1) \ T_BodyStyle.lng监测次数) + 1) * T_BodyStyle.lng监测次数 - 1) * mshUpTab.ColWidth(NewCol) + mshUpTab.ColWidth(0)
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
    Dim intMinCol As Integer, intMaxCol As Integer
    Dim i As Integer, j As Integer
    Dim strTmp As String
    Dim lngColor As Long, lngDc As Long
    Dim objDraw As Object, stdSet As Object
    Dim lng频次 As Long
    Dim lng时间间隔 As Long
    
    lngDc = hDC
    
    lng频次 = T_BodyStyle.lng监测次数
    lng时间间隔 = T_BodyStyle.lng时间间隔
    
    If picMain.Tag = "" Then Exit Sub
    If Row = mshUpTab.Rows - 1 And Col >= mshUpTab.FixedCols Then
        Set objDraw = picBack
        Call CalcMinMaxColNew(picMain.Tag, intMinCol, intMaxCol)
        j = (Col - mshUpTab.FixedCols) Mod lng频次
        
        strTmp = gintHourBegin + lng时间间隔 * j

        '根据参数体温夜班时间范围决定时间颜色
        lngColor = GetTimeColor(Val(strTmp))
        If Col >= intMinCol And Col <= intMaxCol Then
            lngColor = lngColor
        Else
            lngColor = RGB_FleetGRAY
        End If
        
        Call SetTextColor(lngDc, lngColor)
        Call GetTextRect(objDraw, Left, Top + (Bottom - Top) / 2, CStr(strTmp), Right - Left - 3, True)
        Call DrawText(lngDc, CStr(strTmp), -1, T_LableRect, DT_CENTER)
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

Private Sub mshUpTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call RaiseShowTipInfo(mshUpTab, 1, X, Y)
End Sub

Private Sub RaiseShowTipInfo(ByVal vfgObj As Object, ByVal intType As Byte, ByVal X As Single, ByVal Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim lngHeadWidth As Long
    Dim lngWidth As Long, lngHeight As Long, lngHeight1 As Long
    Dim i As Long
    
    If Not vfgObj.Visible Then Exit Sub
    Select Case intType
    Case 1 '上表格
        lngHeadWidth = vfgObj.ColWidth(0)
        lngWidth = vfgObj.ColWidth(vfgObj.FixedCols)
        lngHeight = vfgObj.RowHeight(vfgObj.FixedRows)
    Case 3 '下表格
        lngHeadWidth = vfgObj.ColWidth(0)
        lngWidth = vfgObj.ColWidth(vfgObj.FixedCols)
        lngHeight = vfgObj.RowHeight(vfgObj.FixedRows)
    Case 2 '呼吸表格
        lngHeadWidth = vfgObj.ColWidth(1)
        lngWidth = vfgObj.ColWidth(vfgObj.FixedCols)
        lngHeight = vfgObj.RowHeight(vfgObj.FixedRows)
    Case Else
        Exit Sub
    End Select
    
    lngHeight = 0
    lngHeight1 = 0
    For i = 0 To vfgObj.Rows - 1
        If vfgObj.RowHidden(i) = False Then
            lngHeight = lngHeight + vfgObj.RowHeight(i)
            If Y > lngHeight1 And Y < lngHeight Then Exit For
            lngHeight1 = lngHeight
        End If
    Next i
    
    If i < vfgObj.Rows Then
        lngRow = i
    Else
        Exit Sub
    End If
    
    If X <= lngHeadWidth Then
        lngCol = IIf(intType = 2, 1, 0)
    Else
        lngCol = (X - lngHeadWidth) \ lngWidth + vfgObj.FixedCols
    End If
    If lngRow >= 0 And lngCol >= 0 And lngRow < vfgObj.Rows - IIf(intType = 1, 1, 0) And lngCol < vfgObj.Cols Then
        RaiseEvent ShowTipInfo(vfgObj, vfgObj.TextMatrix(lngRow, lngCol), True)
    Else
        RaiseEvent ShowTipInfo(vfgObj, "", True)
    End If
End Sub

Private Sub picBack_Resize()
    Dim lngLeft As Long
    Dim lngHeight As Long, lngRow As Long
    
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
        .ColWidth(0) = (T_DrawClient.刻度区域.Right - T_DrawClient.刻度区域.Left) * Screen.TwipsPerPixelX
        .Left = lngLeft
        .Top = picCard(0).Top + picCard(0).Height
        .Height = .Rows * mshUpTab.RowHeight(0)
        .Width = ((T_DrawClient.刻度区域.Right - T_DrawClient.刻度区域.Left) + T_DrawClient.列单位 * T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数 + 1) * T_TwipsPerPixel.X
        .ColWidthMin = T_DrawClient.列单位 * Screen.TwipsPerPixelX
         picCard(0).Width = .Width
         .Refresh
    End With
    
    picDraw.Move 0, mshUpTab.Top + mshUpTab.Height, (T_DrawClient.体温区域.Right + 1) * T_TwipsPerPixel.X, _
        (T_DrawClient.曲线总区域.Bottom - T_DrawClient.曲线总区域.Top + 1) * Screen.TwipsPerPixelY

    picDisplay.Height = 165
     
    With vsf
        .Top = mshUpTab.Top + mshUpTab.Height + (T_DrawClient.曲线总区域.Bottom - T_DrawClient.曲线总区域.Top + 1) * Screen.TwipsPerPixelY
        .Left = lngLeft
        .Width = mshUpTab.Width
        .Height = .Body.RowHeight(vsf.FixedRows)
        .Visible = Not mbln呼吸曲线
    End With
        
    With mshDownTab
        .ColWidth(0) = mshUpTab.ColWidth(0)
        .Left = lngLeft
        .Top = mshUpTab.Top + mshUpTab.Height + (IIf(mbln呼吸曲线 = False, vsf.Height, 0)) + (T_DrawClient.曲线总区域.Bottom - T_DrawClient.曲线总区域.Top + 1) * Screen.TwipsPerPixelY - Screen.TwipsPerPixelY
        .Width = mshUpTab.Width
        lngHeight = 0
        For lngRow = 0 To .Rows - 1
            lngHeight = lngHeight + .RowHeight(lngRow)
        Next lngRow
        .Height = lngHeight
        .Refresh
    End With
    
    picCommText.Left = lngLeft
    picCommText.Top = mshDownTab.Top + mshDownTab.Height
    picCommText.Width = mshDownTab.Width
    picCommText.Visible = True
    
    mshUpTab.Redraw = True
    mshDownTab.Redraw = True
    
    picMain.Width = mshUpTab.Width + mshUpTab.Left
    picMain.Height = picCommText.Top + picCommText.Height
    
    '计算滚动条
    Call CalcScrollBarSize
    
    '计算体温单的可画区域大小
    mlng高度 = (picBack.Height - mshUpTab.Top - mshUpTab.Height - mshDownTab.Height - picCommText.Height - _
        IIf(mbln呼吸曲线 = False, vsf.Height, 0) - IIf(hsb.Max > 0, hsb.Height, 0)) / Screen.TwipsPerPixelY
    
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

    
    hsb.Max = 0 - Int(0 - ((picMain.Width - picBack.Width) / 100)) - 1
    vsb.Max = 0 - Int(0 - ((picMain.Height - picBack.Height) / 100)) - 1
    If vsb.Max > 0 Then
        hsb.Max = 0 - Int(0 - ((picMain.Width - picBack.Width + vsb.Width) / 100)) - 1
    End If
    If hsb.Max > 0 Then
        vsb.Max = 0 - Int(0 - ((picMain.Height - picBack.Height + hsb.Height) / 100)) - 1
    End If
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    If hsb.Visible = True Then hsb.ZOrder 0
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    If vsb.Visible = True Then vsb.ZOrder 0
    
    With vsb
        .Height = picBack.Height
    End With
    
    With hsb
        .Width = picBack.Width - IIf(vsb.Max > 0, vsb.Width, 0)
    End With
    
    '只根据没显示出来的那部分来计算步长
    msinHStep = (picMain.Width - picBack.Width + IIf(vsb.Max > 0, vsb.Width, 0)) / 10
    msinVStep = (picMain.Height - picBack.Height + IIf(hsb.Max > 0, hsb.Height, 0)) / 10
    
    '恒定为100,只是步长发生变化
    If hsb.Enabled Then
        hsb.Max = 10
        hsb.LargeChange = 10 / Int((Round((picMain.Width - picBack.Width + IIf(vsb.Max > 0, vsb.Width, 0)) / picBack.Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange
    End If
    
    If vsb.Enabled Then
        vsb.Max = 10
        vsb.LargeChange = 10 / Int((Round((picMain.Height - picBack.Height + IIf(hsb.Max > 0, hsb.Height, 0)) / picBack.Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange
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
    Dim intMinCol As Integer
    Dim intMaxCol As Integer
    Dim intCOl As Integer
    If Button <> vbLeftButton Then Exit Sub
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    '计算指定的区域才可进行操作
    If T_Patient.lng编辑 = 1 Then
        intCOl = ((mintColMax - 1) \ T_BodyStyle.lng监测次数 + 1) * T_BodyStyle.lng监测次数
        
        If X > mshUpTab.ColWidth(0) And X < mshUpTab.ColWidth(0) + (intCOl * mshUpTab.ColWidth(intCOl)) Then
            '根据坐标，计算列数的行
            strTemp = GetXCoordinateNew(X / T_TwipsPerPixel.X + mshUpTab.Left / T_TwipsPerPixel.X - 1, mstr开始时间, False)
            strTemp = mstr开始时间 & ";" & Split(strTemp, ",")(1)
            '根据时间计算列
            Call CalcMinMaxColNew(strTemp, intMinCol, intMaxCol)
            picDisplay.Visible = True
            If Y < mshUpTab.RowHeight(0) + 40 Then
                picDisplay.Left = ((((intMaxCol - 1) \ T_BodyStyle.lng监测次数) + 1) * T_BodyStyle.lng监测次数 - 1) * mshUpTab.ColWidth(intMaxCol) + mshUpTab.ColWidth(0)
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
    '提取病人文件列表
    mstrSQL = "Select A.ID,A.文件名称 From 病人护理文件 A,病历文件列表 B" & _
       "    where A.病人ID=[1] and A.主页Id=[2] and nvl(A.婴儿,0)=[3] and A.格式ID=B.ID and B.种类=3 and B.保留=-1 Order by A.开始时间"
    If mblnMoved = True Then
        mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
    End If
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "提取文件ID", T_Patient.lng病人ID, T_Patient.lng主页ID, T_Patient.lng婴儿)
    cboFile.Clear
    With RS
        Do While Not .EOF
            cboFile.AddItem Nvl(!文件名称)
            cboFile.ItemData(cboFile.NewIndex) = !Id
        .MoveNext
        Loop
    End With
    
    If cboFile.ListCount > 1 Then
        cboFile.Enabled = True
    Else
        cboFile.Enabled = False
    End If
    
    If cboFile.ListCount > 0 And cboFile.ListIndex = -1 Then cboFile.ListIndex = 0
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboFile_Click()
    Dim strParam As String
    
    If T_Patient.lng文件ID = cboFile.ItemData(cboFile.ListIndex) Then Exit Sub
    T_Patient.lng文件ID = cboFile.ItemData(cboFile.ListIndex)
    If mblnAutoAdjust = False Then '正常模式
        '提取初始体温格式构造数据
        strParam = T_Patient.lng病人ID & ";" & T_Patient.lng主页ID & ";" & T_Patient.lng病区ID & ";" & T_Patient.lng文件ID & ";" & _
        T_Patient.lng出院 & ";" & T_Patient.lng编辑 & ";" & T_Patient.lng婴儿 & ";" & T_Patient.lng护理等级 & ";1"
        Call zlMenuClick("初始化", strParam)
    Else
        RaiseEvent zlFileChange(True, T_Patient.lng文件ID, T_Patient.lng婴儿)
    End If
End Sub

Private Sub picDraw_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIndex As Long
    If KeyCode = vbKeyRight And Shift = vbCtrlMask Then  '下一月
        If mintPage < mintAllPage - 1 Then
            lngIndex = GetMenuPageIndex(1)
            mstrParam = mcbrItem.Controls.Item(lngIndex).Parameter '得到当前页的时间
            Call InitWeekDays(mstrParam)
            mcbrToolBar页面.Caption = mcbrItem.Controls.Item(lngIndex).Category
            cbsMain.RecalcLayout
            Call zlMenuClick("装载数据", mstrParam)
        End If

    ElseIf KeyCode = vbKeyLeft And Shift = vbCtrlMask Then
        If mintPage > 0 Then '上一月
            lngIndex = GetMenuPageIndex(-1)
            mstrParam = mcbrItem.Controls.Item(lngIndex).Parameter '得到当前页的时间
            Call InitWeekDays(mstrParam)
            mcbrToolBar页面.Caption = mcbrItem.Controls.Item(lngIndex).Category
            cbsMain.RecalcLayout
            Call zlMenuClick("装载数据", mstrParam)
        End If
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        'mblnAutoRedraw = mblnAutoRedraw Xor True
    End If
End Sub

Private Function GetMenuPageIndex(ByVal intType As Integer) As Long
    '功能:获取体温单页码对应的菜单索引
    'intType:相对当前页要调转的页数
    '72090:刘鹏飞,2014-07-23
    Dim i As Long, lngIndex As Long, lngPage As Long
    
    lngPage = mintPage + intType
    If lngPage < 0 Then
        lngPage = 0
    ElseIf lngPage > mintAllPage - 1 Then
        lngPage = mintAllPage - 1
    End If
    
    For i = 1 To mcbrItem.Controls.Count
        If Val(Split(mcbrItem.Controls.Item(i).Parameter, ";")(4)) = lngPage Then
            lngIndex = i
            Exit For
        End If
    Next i
    
    GetMenuPageIndex = lngIndex
End Function

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '82732:LPF,光标移动到对应的数据点，显示数据信息
    Dim sgnX As Single, sgnY As Single, sgnTmp As Single
    Dim strInfo As String, strTmp As String
    Dim colPonit As Collection
    Dim arrPoint(0 To 2) As String, i As Integer
    
    If mrsPoint Is Nothing Then Exit Sub
    If mrsPoint.State = adStateClosed Then Exit Sub
    
    sgnX = picDraw.ScaleX(X, vbTwips, vbPixels)
    sgnY = picDraw.ScaleX(Y, vbTwips, vbPixels)
    If sgnX >= T_DrawClient.体温区域.Left And sgnX <= T_DrawClient.曲线总区域.Right And sgnY >= T_DrawClient.曲线总区域.Top And sgnY <= T_DrawClient.曲线总区域.Bottom Then
        '如果按照记录集mrsPoint中的坐标来定位，则鼠标必须移动到准确的点(易用性不好),因此采取区域范围定位
        '1、根据鼠标位置重新计算对应点的实际X坐标
        strTmp = GetXCoordinateNew(sgnX, Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"), False)
        sgnTmp = GetXCoordinateNew(Format(Split(strTmp, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"))
        mrsPoint.Filter = "X坐标=" & sgnTmp
        mrsPoint.Sort = "项目序号"
        '2、循环实际X坐标对应的数据，过滤满足鼠标位置对应的数据(上下左右个移动4个点，便于鼠标浮动显示)
        Set colPonit = New Collection
        Do While Not mrsPoint.EOF
            sgnTmp = Val(mrsPoint!X坐标) + T_DrawClient.列单位 \ 2
            If Val(mrsPoint!Y坐标) > sgnY - 4 And Val(mrsPoint!Y坐标) < sgnY + 4 And sgnTmp > sgnX - 4 And sgnTmp < sgnX + 4 Then
                arrPoint(0) = Val(mrsPoint!项目序号)
                arrPoint(1) = Nvl(mrsPoint!部位)
                arrPoint(2) = Nvl(mrsPoint!数值)
                colPonit.Add arrPoint
            End If
            mrsPoint.MoveNext
        Loop
        mrsPoint.Filter = ""
        '3.完成内容输出，格式:项目名称[(部位)]：数值[(项目单位)]
        For i = 1 To colPonit.Count
            mrsItems.Filter = "项目序号 =" & Val(colPonit.Item(i)(0))
            If mrsItems.RecordCount > 0 Then
                strInfo = IIf(strInfo = "", "", strInfo & vbCrLf) & mrsItems!项目名称 & IIf(colPonit.Item(i)(1) = "", "", "(" & colPonit.Item(i)(1) & ")") & "：" & colPonit.Item(i)(2) & "" & IIf(IsNumeric(colPonit.Item(i)(2)) = True, Nvl(mrsItems!项目单位), "")
            End If
        Next i
        
        RaiseEvent ShowTipInfo(picDraw, strInfo, True)
    Else
        RaiseEvent ShowTipInfo(picDraw, "", False)
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
    T_DrawClient.列单位 = T_BodyStyle.lng曲线列宽 \ Screen.TwipsPerPixelX
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
    
    lblSerach(0).FontSize = bytFontSize
    cboFile.FontSize = bytFontSize
    cboFile.Top = (picTmp.Height - cboFile.Height) \ 2
    lblSerach(0).Top = cboFile.Top + (cboFile.Height - lblSerach(0).Height) \ 2
End Sub


Private Sub UserControl_Resize()
    On Error Resume Next
   
    If UserControl.Parent.Visible = False Then Exit Sub
    If mblnAutoAdjust = True And Not mblnResize Then
        '检查实际大小是否发生变化
        If Abs(mlngHeight - UserControl.Height) > 20 Then
            'Debug.Print "--大小改变进入--"
            Call LockWindowUpdate(UserControl.hWnd)
            Call zlMenuClick("装载数据", mstrParam)
            Call LockWindowUpdate(0)
            mblnResize = True
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
    T_ClientRect.Right = T_ClientRect.Right * 2
    T_ClientRect.Bottom = T_ClientRect.Bottom * 2
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

Public Sub Paint_CanvasNew(Optional ByVal blnAdjust As Boolean = False)
    '准备画布（完成刻度及表格缩放、三测单表格绘制以及根据设定进行基准线的描绘）
    '最小模式下,不显示表上表格,文本等数据
    'blnAdjust=False表示固定大小，否则跟随主界面进行调整
    
    Static SlngMaxY As Long                 '记录上一次的最大高度，以决定本次是否需要重画
    Dim lngCurX     As Long, lngCurY As Long   '当前位置
    Dim lngMaxX     As Long, lngMaxY As Long, lngAllMaxY As Long  '边界
    Dim lngCurAlerY As Long
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
    Dim lngCurveRows As Long '独立曲线总列数
    Dim lngY As Long, lngX As Long
    Dim str说明 As String
    Dim sinCurAlerY As Single
    '以下与绘图区域相关(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
    Dim sin刻度 As Single, bln显示刻度 As Boolean, blnFirst As Boolean
    Dim sin刻度间隔 As Single, sinBegin刻度 As Single, dbl单位值 As Double

    Dim str最大值坐标 As String, str最小值坐标 As String
    Dim lng刻度宽度 As Long
    

    On Error GoTo Errhand
    
    '实现缩放的原理说明：
    '1、普通模式下所有内容均显示
    '2、最小模式=2，时间刻度不显示，每行10小行改为5小行
    '3、缩小模式<=4，转为虚线显示
    
    '以前是固定以上面有2行来输出数据，所以此处减去2行
    '后面输出时为了对齐好看，再次减2行来输出
    lngCurveRow = T_BodyStyle.lng曲线空行
    
    T_TwipsPerPixel.X = Screen.TwipsPerPixelX
    T_TwipsPerPixel.Y = Screen.TwipsPerPixelY
    T_DrawClient.总列数 = glngMaxRows
    
    gstrSQL = " Select /*+ Rule*/ A.项目序号,A.排列序号,A.记录名,A.记录符,A.记录色,nvl(A.最大值,0) 最大值,nvl(A.最小值,0) 最小值," & _
        "nvl(A.单位值,0) 单位值,A.刻度间隔,A.警示线,C.项目单位 单位,Decode(记录法,3,A.最高行,nvl(A.最高行,2)-2) AS 最高行,B.部位,A.记录法" & _
        " From 体温记录项目 A,体温部位 B,护理记录项目 C,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) D" & _
        " Where A.项目序号=B.项目序号(+) And B.缺省项(+)=1" & _
        " And  A.项目序号=C.项目序号 AND A.记录法<>2 AND NOT (NVL(C.应用方式,0)=2 And C.项目序号=-1) and C.项目序号=D.COLUMN_VALUE" & _
        " Order by 排列序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取开始行", T_BodyItem.str曲线项目)
    
    '------------------------------------------------------------------------------------------------------------------
    rsTemp.Filter = "项目序号=" & gint体温
    '计算打印输出的行数
    With rsTemp
        Do While Not .EOF
            lng最高行 = Val(zlCommFun.Nvl(!最高行))
            If lng最高行 < 0 Then lng最高行 = 0
            
             '修改问题51442
            If Val(zlCommFun.Nvl(!最小值, 0)) > 34 Then
                lngMaxRows = lng最高行 + (Val(zlCommFun.Nvl(!最大值, 0)) - 35) / 0.1
            Else
                lngMaxRows = lng最高行 + (Val(zlCommFun.Nvl(!最大值, 0)) - Val(zlCommFun.Nvl(!最小值, 0))) / 0.1
            End If

            lngMaxRows = lngMaxRows + lngCurveRow
            T_DrawClient.总列数 = lngMaxRows
        .MoveNext
        Loop
    End With
    
    T_DrawClient.独立曲线总行数 = 0
    rsTemp.Filter = "记录法=3 And 项目序号<>1"
    rsTemp.Sort = "排列序号"
    Do While Not rsTemp.EOF
        lngRow = ((Val(Nvl(rsTemp!最大值, 0)) - Val(Nvl(rsTemp!最小值, 0))) / Val(Nvl(rsTemp!单位值, 1)))
        If Val(Nvl(rsTemp!最高行, 0)) > 0 Then lngRow = lngRow + Val(Nvl(rsTemp!最高行, 0))
        If lngRow Mod 2 = 1 Then lngRow = lngRow + 1
        T_DrawClient.独立曲线总行数 = T_DrawClient.独立曲线总行数 + lngRow
    rsTemp.MoveNext
    Loop
    
    rsTemp.Filter = "记录法=1"
    rsTemp.Sort = "排列序号"
    intLables = rsTemp.RecordCount
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    
    '赋初值
    intLineMode = PS_SOLID
    
    lngColStep = T_BodyStyle.lng曲线列宽 \ Screen.TwipsPerPixelX
    lngInitRowStep = T_BodyStyle.lng曲线行高 \ Screen.TwipsPerPixelY
    sigRowStepNew = lngInitRowStep
    lng刻度宽度 = (T_BodyStyle.lng刻度宽度 \ Screen.TwipsPerPixelX)
    lngLableStep = Fix(lng刻度宽度 / intLables)
    intTens_digit = 3
    '体温单以单格显示(不勾此选项以双格显示，没两个刻度显示一次) 1：单格显示 0：双格显示
    If zlDatabase.GetPara("体温单显示格式", glngSys, 1255, 0) = 1 Then
        bln双行 = False
    Else
        bln双行 = True
    End If
    'True表示贰行只输出一行,效果是一个刻度只显示了五行;否则一个刻度显示十行,由用户调整参数决定,与blnDoubleRow无关
    bln粗线 = True
    
    If Not bln粗线 Then intLineMode = PS_DASHDOTDOT
    
    '画表格
    lngCurX = T_DrawClient.偏移量X
    lngCurY = T_DrawClient.偏移量Y
    lngMaxX = lng刻度宽度 + (T_BodyStyle.lng天数 * T_BodyStyle.lng监测次数 * lngColStep) + T_DrawClient.偏移量X   '刻度+总列数*宽度 +T_DrawClient.偏移量X
    lngMaxY = 2 * mintNullRow * lngInitRowStep + T_DrawClient.总列数 * sigRowStepNew + T_DrawClient.偏移量Y '非独立曲线部分最大Y坐标
    lngAllMaxY = 2 * mintNullRow * lngInitRowStep + (T_DrawClient.总列数 + T_DrawClient.独立曲线总行数) * sigRowStepNew + T_DrawClient.偏移量Y '所遇曲线部分最大Y坐标
    '进行相关数据的校正
    If blnAdjust Then
        '如果小于可见区域大小则进行缩放
        If lngAllMaxY > mlng高度 Then
            lngAllMaxY = mlng高度 - 2 * mintNullRow * lngInitRowStep
            sigRowStepNew = Round((lngAllMaxY) / (T_DrawClient.总列数 + T_DrawClient.独立曲线总行数), 1)
            sigRowStepNew = Fix(sigRowStepNew + 0.5)
        End If

        '如果行高太小，则将贰行做为一行显示
        If sigRowStepNew <= 2 Then
            sinRowStep = 2
            blnDoubleRow = True
        End If

        If Not mblnRedraw Then mblnRedraw = (lngAllMaxY <> SlngMaxY)
        If sigRowStepNew < 4 Then intLineMode = PS_DOT
    End If
    '计算刻度的最大坐标
    lngMaxY = (lngInitRowStep * 2 * mintNullRow) + T_DrawClient.总列数 * IIf(blnDoubleRow, sinRowStep, sigRowStepNew) + T_DrawClient.偏移量Y
    lngAllMaxY = (lngInitRowStep * 2 * mintNullRow) + (T_DrawClient.总列数 + T_DrawClient.独立曲线总行数) * IIf(blnDoubleRow, sinRowStep, sigRowStepNew) + T_DrawClient.偏移量Y
    
    Call Paint_Reset                                                    '清除画布
    
    SlngMaxY = lngMaxY
    T_DrawClient.刻度单位 = lngLableStep
    T_DrawClient.行单位 = IIf(blnDoubleRow, sinRowStep, sigRowStepNew)
    T_DrawClient.列单位 = lngColStep
    T_DrawClient.双倍 = blnDoubleRow
    
    For lngRow = 1 To intLables
        Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
        '由于可能存在刻度总宽度/项目数除不尽的(如：90/4),处理方式为前3列为Fix(90/4),最后一列的宽度为刻度宽度-前3列的宽度
        '保证表上表格列头、标下表格列头和刻度宽度相同
        If lngRow = intLables Then
            lngCurX = lng刻度宽度 + T_DrawClient.偏移量X
        Else
            lngCurX = lngCurX + lngLableStep
        End If
    Next
    Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
    
    '画刻度框
    Call DrawLine(mlngMemDC, T_DrawClient.偏移量X, lngCurY, lngMaxX, lngCurY, PS_SOLID, 1, RGB_BLACK)

    T_DrawClient.刻度区域.Left = T_DrawClient.偏移量X
    T_DrawClient.刻度区域.Top = lngCurY
    T_DrawClient.刻度区域.Right = lng刻度宽度 + T_DrawClient.偏移量X
    T_DrawClient.刻度区域.Bottom = lngMaxY
    
    '默认添加一行用于显示项目名称
    lngCurY = lngCurY + lngInitRowStep * 2
    Call DrawLine(mlngMemDC, T_DrawClient.偏移量X, lngCurY, lngMaxX, lngCurY, PS_SOLID, 1, RGB_BLACK)
    lngCurY = lngCurY + lngInitRowStep * ((mintNullRow - 1) * 2)
    
    '画体温单所有行
    For lngRow = 0 To T_DrawClient.总列数 - 1
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
    For lngRow = 1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数
        lngCurX = lngCurX + lngColStep
        
        Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng监测次数 = 0, 2, 1), IIf(lngRow Mod T_BodyStyle.lng监测次数 = 0, RGB_RED, RGB_GRAY))
    Next
    
    lngCurX = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Left = T_DrawClient.刻度区域.Right
    T_DrawClient.体温区域.Top = T_DrawClient.刻度区域.Top
    T_DrawClient.体温区域.Right = lngMaxX
    T_DrawClient.体温区域.Bottom = lngMaxY
    
    T_DrawClient.曲线总区域.Left = T_DrawClient.刻度区域.Left
    T_DrawClient.曲线总区域.Top = T_DrawClient.刻度区域.Top
    T_DrawClient.曲线总区域.Right = lngMaxX
    T_DrawClient.曲线总区域.Bottom = lngAllMaxY
    
    '画体温区域底线
    Call DrawLine(mlngMemDC, T_DrawClient.偏移量X, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
    
    Set mobjPart = New Collection
    '画刻度框的标尺（从固定不变的10行开始标识）
    rsTemp.Filter = "记录法=1"
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            '显示刻度框项目的名称及符号,如体温×
            lngCurX = T_DrawClient.刻度区域.Left + ((.AbsolutePosition - 1) * T_DrawClient.刻度单位)
            If .AbsolutePosition = .RecordCount Then
                lngLableStep = (T_DrawClient.刻度区域.Right - T_DrawClient.刻度区域.Left) - ((.AbsolutePosition - 1) * T_DrawClient.刻度单位)
            Else
                lngLableStep = T_DrawClient.刻度单位
            End If
            lngCurY = T_DrawClient.刻度区域.Top
            
            '设置字体大小
            Set gstdSet = New StdFont
            gstdSet.Name = "宋体"
            gstdSet.Size = 9
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)
            
            '输出体温项目的名称
            Call SetTextColor(mlngMemDC, zlCommFun.Nvl(!记录色, RGB_BLACK))
            Call GetTextRect(mobjDraw, lngCurX, lngCurY + mobjDraw.TextHeight(zlCommFun.Nvl(!记录名)) / Screen.TwipsPerPixelY / 2, Trim(zlCommFun.Nvl(!记录名)), lngLableStep)
            Call DrawText(mlngMemDC, Trim(zlCommFun.Nvl(!记录名)), -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            Call ReleaseFontIndirect(mobjDraw)
            '设置字体大小
            Set gstdSet = New StdFont
            gstdSet.Name = "宋体"
            gstdSet.Size = 8
            Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
            mlngFont = CreateFontIndirect(T_Font)
            mlngOldFont = SelectObject(mlngMemDC, mlngFont)

            '输出项目单位
            Call GetTextRect(mobjDraw, lngCurX, lngCurY + lngInitRowStep * 2 + mobjDraw.TextHeight(zlCommFun.Nvl(!单位)) / Screen.TwipsPerPixelY / 2, Trim(zlCommFun.Nvl(!单位)), lngLableStep)
            Call DrawText(mlngMemDC, Trim(zlCommFun.Nvl(!单位)), -1, T_LableRect, DT_CENTER)
            Call SelectObject(mlngMemDC, mlngOldFont)
            Call DeleteObject(mlngFont)
            Call ReleaseFontIndirect(mobjDraw)
            sinY单位 = T_LableRect.Bottom
            '创建字体
            Set gstdSet = New StdFont
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

                Case gint脉搏, gint心率  '脉搏/心跳按10的倍数输出刻度
                    intTens_digit = 3
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 10)
                    dbl单位值 = 2
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)

                Case gint呼吸  '呼吸按5的倍数输出刻度
                    mbln呼吸曲线 = True
                    intTens_digit = 2
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, 5)
                    dbl单位值 = 1
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
                Case Else
                    intTens_digit = 1
                    dbl单位值 = Val(zlCommFun.Nvl(!单位值, 1))
                    sin刻度间隔 = zlCommFun.Nvl(!刻度间隔, Val(zlCommFun.Nvl(!单位值, 0)) * 10)
                    If sin刻度间隔 > Val(zlCommFun.Nvl(!最大值)) - Val(zlCommFun.Nvl(!最小值)) Then
                        sin刻度间隔 = Val(zlCommFun.Nvl(!最大值)) - Val(zlCommFun.Nvl(!最小值))
                    End If
                    sinAlertness = zlCommFun.Nvl(!警示线, 0)
            End Select
            
            If !项目序号 = gint体温 Then
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "口温", arrTemp(0), Nvl(!记录色, RGB_BLACK), "B"), "B" & !项目序号
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "腋温", arrTemp(1), Nvl(!记录色, RGB_BLACK), "A"), "A" & !项目序号
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "肛温", arrTemp(2), Nvl(!记录色, RGB_BLACK), "C"), "C" & !项目序号
            ElseIf !项目序号 = gint脉搏 Then
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "缺省记录符", Nvl(!记录符), Nvl(!记录色, RGB_BLACK), "A"), "A" & !项目序号
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "起搏器", "H", RGB_RED, "B"), "B" & !项目序号
                If mint心率应用 = 2 Then
                    mrsItems.Filter = "项目序号=" & gint心率
                    If mrsItems.RecordCount > 0 Then
                        mobjPart.Add Array("" & gint心率, Nvl(mrsItems!项目名称), "", Nvl(mrsItems!记录符), RGB_RED, "A"), "A" & gint心率
                    End If
                    mrsItems.Filter = ""
                End If
            ElseIf !项目序号 = gint呼吸 Then
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "自主呼吸", Nvl(!记录符), Nvl(!记录色, RGB_BLACK), "A"), "A" & !项目序号
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "呼吸机", "R", RGB_BLACK, "B"), "B" & !项目序号
            Else
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "", Nvl(!记录符), Nvl(!记录色, RGB_BLACK), "A"), "A" & !项目序号
            End If
            
            '赋初值
            lngCurY = lngCurY + (lngInitRowStep * 2 * mintNullRow)   '固定前2 * mintNullRow行的高度不输出刻度

            '如果是最小模式,从第30行开始输出标识
            'If blnDoubleRow Then lngCurY = lngCurY + lngInitRowStep * 2 * mintNullRow
            
            '根据最高行定位到有效位置
            lngCurY = lngCurY + (T_DrawClient.行单位 * zlCommFun.Nvl(!最高行, 2))
            blnFirst = False
            Do While True
                bln显示刻度 = False
                If blnFirst = False Then    '刚进入循环，此时取的最大值
                    sin刻度 = zlCommFun.Nvl(!最大值, 0)
                    sinBegin刻度 = sin刻度
                    str最大值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                    blnFirst = True
                Else                    '计算得到每个刻度的值
                    sin刻度 = sin刻度 - dbl单位值    '如果目前显示模式为双倍，则按双倍累计
                End If
                
                If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
                
                If bln显示刻度 = True Or sin刻度 < sinBegin刻度 Then sinBegin刻度 = sinBegin刻度 - IIf(T_DrawClient.双倍, sin刻度间隔 * 2, sin刻度间隔)
                
                If sinBegin刻度 < Val(Format(!最小值, "#0.00")) Then sinBegin刻度 = Val(Format(!最小值, "#0.00"))
                
                If bln显示刻度 Then
                    '控制最大值不与曲线单位重复
                    If sin刻度 = Val(Nvl(!最大值, 0)) And lngCurY < sinY单位 Then
                        Call GetTextRect(mobjDraw, lngCurX, sinY单位, Format(sin刻度, "#0"), lngLableStep)
                    ElseIf lngCurY = T_DrawClient.刻度区域.Bottom Then
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY - (mobjDraw.TextHeight("1") / (T_TwipsPerPixel.Y * 2)), Format(sin刻度, "#0"), lngLableStep)
                    Else
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY, Format(sin刻度, "#0"), lngLableStep)
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
                    If (sinAlertness < Val(Nvl(!最大值)) And sinAlertness > Val(Nvl(!最小值))) Then
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
            Call ReleaseFontIndirect(mobjDraw)
            sinBegin刻度 = 0
            sin刻度 = 0                 '控制从第一行开始输出
            .MoveNext
        Loop
    End With

    '完成独立曲线部分的输出
    lngMaxY = T_DrawClient.刻度区域.Bottom
    rsTemp.Filter = "记录法=3"
    rsTemp.Sort = "排列序号"
    With rsTemp
        Do While Not .EOF
            lngY = lngMaxY
            lngCurY = lngY
            lngCurX = T_DrawClient.偏移量X
            lngCurveRows = ((Val(Nvl(!最大值, 0)) - Val(Nvl(!最小值, 0))) / Val(Nvl(!单位值)))
            If Val(Nvl(!最高行, 0)) > 0 Then lngCurveRows = lngCurveRows + Val(Nvl(!最高行, 0))
            If lngCurveRows Mod 2 = 1 Then lngCurveRows = lngCurveRows + 1
            If lngCurveRows > 0 Then
                lngMaxY = lngCurveRows * IIf(blnDoubleRow, sinRowStep, sigRowStepNew) + lngCurY
                '完成刻度区域的绘制
                Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
                Call DrawLine(mlngMemDC, lngCurX + lng刻度宽度, lngCurY, lngCurX + lng刻度宽度, lngMaxY, PS_SOLID, 1, RGB_BLACK)
                Call DrawLine(mlngMemDC, lngCurX, lngMaxY, lngCurX + lng刻度宽度, lngMaxY, PS_SOLID, 1, RGB_BLACK)
                '完成所有行的绘制
                lngCurX = lngCurX + lng刻度宽度
                For lngRow = 1 To lngCurveRows
                    '画体温单的所有行
                    If lngRow <> 0 Then
                        lngCurY = lngCurY + IIf(blnDoubleRow, sinRowStep, sigRowStepNew)
                    End If

                    If ((blnDoubleRow Or bln双行) And lngRow Mod 2 = 0) Or (Not blnDoubleRow And Not bln双行) Then
                        Call DrawLine(mlngMemDC, lngCurX + 1, lngCurY, lngMaxX, lngCurY, IIf(lngRow Mod 10 = 0, PS_SOLID, intLineMode), IIf(lngRow Mod 5 = 0 And sigRowStepNew >= 4 And bln粗线, 2, 1), RGB_FleetGRAY)
                    End If
                Next
                lngCurY = lngY

                 '画体温单所有列
                For lngRow = 1 To T_BodyStyle.lng监测次数 * T_BodyStyle.lng天数
                    lngCurX = lngCurX + lngColStep
                    Call DrawLine(mlngMemDC, lngCurX, lngCurY, lngCurX, lngMaxY, PS_SOLID, IIf(lngRow Mod T_BodyStyle.lng监测次数 = 0, 2, 1), IIf(lngRow Mod T_BodyStyle.lng监测次数 = 0, RGB_RED, RGB_GRAY))
                Next
                
                '画体温区域底线
                Call DrawLine(mlngMemDC, T_DrawClient.偏移量X, lngMaxY, lngMaxX, lngMaxY, PS_SOLID, 1, RGB_BLACK)
                '完成项目名称和刻度的输出
                lngCurX = lngX: lngCurY = lngY
                '输出体温项目的名称
                '创建字体
                Set gstdSet = New StdFont
                gstdSet.Name = "宋体"
                gstdSet.Size = 9
                Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
                mlngFont = CreateFontIndirect(T_Font)
                mlngOldFont = SelectObject(mlngMemDC, mlngFont)
                Call SetTextColor(mlngMemDC, Nvl(!记录色, RGB_BLACK))
                T_Size.H = mobjDraw.ScaleY(mobjDraw.TextHeight("刘"), vbTwips, vbPixels)
                If T_Size.H * Len(Nvl(!记录名)) >= lngCurveRows * T_DrawClient.行单位 Then
                    lngCurY = lngY
                Else
                    lngCurY = lngY + ((lngCurveRows * T_DrawClient.行单位) - (T_Size.H * Len(Nvl(!记录名)))) \ 2
                End If
                For lngRow = 1 To Len(Nvl(!记录名))
                    Call GetTextRect(mobjDraw, lngCurX, lngCurY, Mid(Nvl(!记录名), lngRow, 1), lng刻度宽度 \ 2, False)
                    Call DrawText(mlngMemDC, Mid(Nvl(!记录名), lngRow, 1), -1, T_LableRect, DT_CENTER)
                    lngCurY = lngCurY + T_Size.H
                Next lngRow
                Call SelectObject(mlngMemDC, mlngOldFont)
                Call DeleteObject(mlngFont)
                Call ReleaseFontIndirect(mobjDraw)
                '输出项目单位
                lngCurY = lngY: If Nvl(!记录名) <> "" Then lngCurX = T_LableRect.Right
                If Trim(Nvl(!单位)) <> "" And Nvl(!记录名) <> "" Then
                    '设置字体大小
                    Set gstdSet = New StdFont
                    gstdSet.Name = "宋体"
                    gstdSet.Size = 8
                    Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
                    mlngFont = CreateFontIndirect(T_Font)
                    mlngOldFont = SelectObject(mlngMemDC, mlngFont)
                    T_Size.H = mobjDraw.ScaleY(mobjDraw.TextHeight("刘"), vbTwips, vbPixels)
                    If T_Size.H * Len(Trim(Nvl(!单位))) >= lngCurveRows * T_DrawClient.行单位 Then
                        lngCurY = lngY
                    Else
                        lngCurY = lngY + ((lngCurveRows * T_DrawClient.行单位) - (T_Size.H * Len(Nvl(!单位)))) \ 2
                    End If
                    For lngRow = 1 To Len(Trim(Nvl(!单位)))
                        Call GetTextRect(mobjDraw, lngCurX, lngCurY, Mid(Trim(Nvl(!单位)), lngRow, 1), 0, False)
                        Call DrawText(mlngMemDC, Mid(Trim(Nvl(!单位)), lngRow, 1), -1, T_LableRect, DT_CENTER)
                        lngCurY = lngCurY + T_Size.H
                    Next lngRow
                    Call SelectObject(mlngMemDC, mlngOldFont)
                    Call DeleteObject(mlngFont)
                    Call ReleaseFontIndirect(mobjDraw)
                End If
                '设置字体大小
                Set gstdSet = New StdFont
                gstdSet.Name = "宋体"
                gstdSet.Size = 9
                Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
                mlngFont = CreateFontIndirect(T_Font)
                mlngOldFont = SelectObject(mlngMemDC, mlngFont)
                dbl单位值 = Val(Nvl(!单位值, 0))
                sin刻度间隔 = Nvl(!刻度间隔, Val(Nvl(!单位值, 0)) * 10)
                If sin刻度间隔 > Val(Nvl(!最大值)) - Val(Nvl(!最小值)) Then
                    sin刻度间隔 = Val(Nvl(!最大值)) - Val(Nvl(!最小值))
                End If
                sinAlertness = Nvl(!警示线, 0)
                str说明 = str说明 & "、" & Nvl(!记录名) & "(" & Nvl(!记录符, "*") & ")"
                mobjPart.Add Array("" & !项目序号, Nvl(!记录名), "", Nvl(!记录符), Nvl(!记录色, RGB_BLACK), "A"), "A" & !项目序号
                
                intTens_digit = 1
                lngCurY = lngY + (T_DrawClient.行单位 * Val(Nvl(!最高行, 0)))
                blnFirst = False
                Do While True
                    bln显示刻度 = False
                    If blnFirst = False Then     '刚进入循环，此时取的最大值
                        sin刻度 = Nvl(!最大值, 0)
                        sinBegin刻度 = sin刻度
                        str最大值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                        blnFirst = True
                    Else                    '计算得到每个刻度的值
                        sin刻度 = sin刻度 - dbl单位值     '如果目前显示模式为双倍，则按双倍累计
                    End If
    
                    '根据设置的刻度间隔显示刻度值
                    If Val(Format(sin刻度, "#0.00")) = Val(Format(sinBegin刻度, "#0.00")) Then bln显示刻度 = True
                    If bln显示刻度 = True Or sin刻度 < sinBegin刻度 Then sinBegin刻度 = sinBegin刻度 - IIf(T_DrawClient.双倍, sin刻度间隔 * 2, sin刻度间隔)
                    If sinBegin刻度 < Val(Format(Nvl(!最小值), "#0.00")) Then sinBegin刻度 = Val(Format(Nvl(!最小值), "#0.00"))
    
                    If bln显示刻度 Then
                        '控制最大值不与曲线单位重复
                        lngCurX = T_DrawClient.体温区域.Left - mobjDraw.ScaleX(mobjDraw.TextWidth(Val(Format(sin刻度, "#0.0"))), vbTwips, vbPixels)
                        lngCurX = lngCurX - (mobjDraw.ScaleY(mobjDraw.TextHeight("1"), vbTwips, vbPixels) \ 3)
                        If sin刻度 = Val(Nvl(!最大值, 0)) And lngCurY = lngY Then
                            Call GetTextRect(mobjDraw, lngCurX, lngCurY + (mobjDraw.ScaleY(mobjDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin刻度, "#0.0")))
                        ElseIf lngCurY = lngMaxY Then
                            Call GetTextRect(mobjDraw, lngCurX, lngCurY - (mobjDraw.ScaleY(mobjDraw.TextHeight("1"), vbTwips, vbPixels) \ 2), Val(Format(sin刻度, "#0.0")))
                        Else
                            Call GetTextRect(mobjDraw, lngCurX, lngCurY, Val(Format(sin刻度, "#0.0")))
                        End If
                        Call DrawText(mlngMemDC, Val(Format(sin刻度, "#0.0")), -1, T_LableRect, DT_CENTER)
                    End If
                    If Val(Format(sin刻度, "#0.00")) <= Val(Format(Nvl(!最小值), "#0.00")) Or Format(lngCurY, "#0") > lngMaxY Then
                        str最小值坐标 = T_DrawClient.体温区域.Left & "," & lngCurY
                        '添加该项目(项目序号,最大值,最小值,单位值,最大值坐标,最小值坐标,单位刻度,显示模式)
                        gstrFields = "项目序号|最大值|最小值|单位值|最大值坐标|最小值坐标|单位刻度|显示模式|颜色"
                        gstrValues = Nvl(!项目序号) & "|" & Nvl(!最大值) & "|" & Nvl(!最小值) & "|" & dbl单位值 & "|" & _
                            str最大值坐标 & "|" & str最小值坐标 & "|" & T_DrawClient.行单位 & "," & T_DrawClient.列单位 & "|" & intTens_digit & "|" & !记录色
                        Call Record_Add(mrsDrawItems, gstrFields, gstrValues)
                        '输出警戒线
                        If blnDoubleRow = False And sinAlertness > Val(Nvl(!最小值)) And sinAlertness < Val(Nvl(!最大值)) Then
                            '根据最大值与当前值之间的差额,以及最小值,计算得到相差多少个刻度,再根据单位刻度得到实际坐标
                            sinCurAlerY = Val(GetYCoordinate(mobjDraw, mrsDrawItems, Val(Nvl(!项目序号)), sinAlertness))
                            Call DrawLine(mlngMemDC, T_DrawClient.体温区域.Left, CLng(sinCurAlerY), lngMaxX, CLng(sinCurAlerY), PS_SOLID, 1, RGB_RED)
                        End If
                        Exit Do
                    End If
                    lngCurY = lngCurY + T_DrawClient.行单位
                Loop
                '还原字体信息
                Call SelectObject(mlngMemDC, mlngOldFont)
                Call DeleteObject(mlngFont)
                Call ReleaseFontIndirect(mobjDraw)
                sinBegin刻度 = 0
                sin刻度 = 0
            End If
        .MoveNext
        Loop
    End With
        
    '创建字体
    Set gstdSet = New StdFont
    gstdSet.Name = "宋体"
    gstdSet.Size = 9
    Call SetFontIndirect(gstdSet, mlngMemDC, mobjDraw)
    mlngFont = CreateFontIndirect(T_Font)
    mlngOldFont = SelectObject(mlngMemDC, mlngFont)
    
    mblnRedraw = False                      '画过一次后就不再画了
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub Paint_Construct()

    Dim lngRGB  As Long
    Dim blnLine As Boolean              '心率与脉搏共用时,心率不连线
    Dim str心率 As String               '记录所有心率所在的列(X坐标)
    Dim str原值 As String, sinX坐标原 As Single, sinY坐标原 As Single
    Dim dbl数值 As Double, dblMinValue As Double, dblMaxValue As Double
    Dim bln不升符号 As Boolean
    Dim lng体温不升显示方式 As Long
    Dim lng行高 As Long
    Dim lngWith As Long
    Dim bln符号 As Boolean
    Dim strWaveReview As String, lngWaveReviewColor As Long '参数:体温复查的符号及颜色
    On Error GoTo Errhand

    '开始作图：完成图形数据的输出（重叠的处理、体温复核、图形标记输出、脉搏短拙）
    '先画线(除心率外)
    '再处理脉搏短拙
    '再输出图形
    lng行高 = T_BodyStyle.lng曲线行高 \ Screen.TwipsPerPixelY
    
    lng体温不升显示方式 = Val(zlDatabase.GetPara("体温不升显示方式", glngSys, 1255, "0"))
    strWaveReview = zlDatabase.GetPara("体温复试合格符号", glngSys, 1255, "v")
    '75319
    lngWaveReviewColor = Val(zlDatabase.GetPara("体温复试合格颜色", glngSys, 1255, "10485760"))
    
    With mrsPoint
        .Filter = ""
        '先画线
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "项目序号,时间"
        Do While Not .EOF
            If Val(zlCommFun.Nvl(!状态)) <> 3 Then
                '物理降温的后面处理,不连线
                If Not ((!项目序号 = gint体温 Or !项目序号 = gint疼痛强度) And !标记 = 1) Then
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
                        Call SetTextColor(mlngMemDC, lngWaveReviewColor)
                        Call GetTextRect(mobjDraw, !X坐标, !Y坐标 - Screen.TwipsPerPixelY, strWaveReview, T_DrawClient.列单位, False)
                        Call DrawText(mlngMemDC, strWaveReview, -1, T_LableRect, DT_CENTER)
                    End If
                    
                    '问题号:56886,李涛,2013-05-06
                    bln符号 = GetSymbol(!项目序号, !部位, !重叠项目, !符号)
                    lngWith = 0
                    If bln符号 Then
                        lngWith = mobjDraw.TextWidth("○") / 4 / T_TwipsPerPixel.X
                    End If
                    
                    If sinX坐标原 <> 0 And blnLine Then
                        Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2 - lngWith, !Y坐标, sinX坐标原 + T_DrawClient.列单位 / 2, sinY坐标原, PS_SOLID, 1, lngRGB)
                    End If
                    

                    If !断开 = 0 Then
                        sinX坐标原 = !X坐标 + lngWith
                        sinY坐标原 = !Y坐标
                    Else
                        sinX坐标原 = 0
                    End If

                    '此处处理项目高出项目的最大值 或小于项目最小值
                    If Not (!项目序号 = gint体温 And Trim(Nvl(!数值)) = "不升") Then
                        dbl数值 = Val(zlCommFun.Nvl(!数值))
                        '重叠时以序号靠前的为准
                        If !重叠 = 0 Then
                            If dbl数值 < dblMinValue Then
                                Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 + IIf(T_DrawClient.行单位 < lng行高, lng行高, T_DrawClient.行单位) * 2, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, 1, lngRGB, True)
                            End If
                            
                            If dbl数值 > dblMaxValue Then
                                Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 - IIf(T_DrawClient.行单位 < lng行高, lng行高, T_DrawClient.行单位) * 2, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标, PS_SOLID, 1, lngRGB, True)
                            End If
                        End If
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
        If str心率 <> "" Then Call CreatePolyNew(mrsPoint, mobjDraw, mlngMemDC, mstr开始时间, str心率, mint心率应用 = 2)

        '输出点或图形
        If .RecordCount <> 0 Then .MoveFirst
        .Sort = "项目序号,时间"
        
    
        Do While Not .EOF
            If Val(zlCommFun.Nvl(!状态)) <> 3 Then
                If (!项目序号 = gint体温 Or !项目序号 = gint疼痛强度) And !标记 = 1 Then
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
                        Call DrawLine(mlngMemDC, !X坐标 + T_DrawClient.列单位 / 2, !Y坐标 + (T_Size.H / 4), sinX坐标原 + T_DrawClient.列单位 / 2, sinY坐标原, PS_DOT, 1, RGB_RED, False)
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
    Dim blnCurBeginTop As Boolean  '体温标志是否从顶端输出
    Dim Y As Long, X As Long, Y1 As Long
    Dim bln改变 As Boolean
    Dim lngX          As Long, lngY As Long
    Dim strComment    As String, strTemp As String, strText As String
    Dim intNum        As Integer
    Dim intAscCharNum As Integer
    Dim varNote()     As String
    Dim i  As Integer, j As Integer

    On Error GoTo Errhand
    
    '参数
    byt未记说明显示位置 = Val(zlDatabase.GetPara("未记说明显示位置", glngSys, 1255, "0"))
    blnCurBeginTop = (Val(zlDatabase.GetPara("体温标志输出位置", glngSys, 1255, "0")) = 1)
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
    
    If blnCurBeginTop = False Then
        Y1 = GetYCoordinate(mobjDraw, mrsDrawItems, gint体温, 42, mlngMemDC)
    Else
        Y1 = T_DrawClient.体温区域.Top
    End If
    
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

                            If Y < T_DrawClient.体温区域.Bottom Then
                                strText = Mid(strComment, i, 1)
                                Call GetTextExtentPoint32(mlngMemDC, strText, Len(strText), T_Size)
                                '输出字体信息
                                If T_DrawClient.体温区域.Bottom - Y >= T_Size.H - 1 Then
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
                    If Y < T_DrawClient.体温区域.Bottom Then
                        strText = Mid(strComment, i, 1)
                        Call GetTextExtentPoint32(mlngMemDC, strText, Len(strText), T_Size)

                        If Asc(strText) < 0 Then
                            If (intAscCharNum - intNum) Mod 2 = 1 Then Y = Y + T_Size.H / 2
                        End If
                         
                        '输出字体信息
                        If T_DrawClient.体温区域.Bottom - Y >= T_Size.H - 1 Then
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
    Call OutPutTextNew(mobjDraw, mrsDrawItems, mlngMemDC, mrsNote, mstr开始时间, blnCurBeginTop)
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
    Dim strValue As String
    
    mrsItems.Filter = "项目序号=" & lng项目序号
    If mrsItems.EOF Then Exit Function
    
'    If InStr(1, Nvl(mrsItems!项目值域), ";") = 0 Then
'        dblvalue = Val(Nvl(mrsItems!最小值, 0))
'    Else
'        dblvalue = Val(Split(mrsItems!项目值域, ";")(0))
'    End If
    dblvalue = Val(Nvl(mrsItems!最小值, 0))
    strValue = Nvl(mrsItems!临界值)
    If InStr(1, strValue, ";") <> 0 Then
        strValue = Split(strValue, ";")(0)
    Else
        strValue = ""
    End If
    
    If IsNumeric(strValue) = True And Val(strValue) <= Val(Nvl(mrsItems!最大值)) And Val(strValue) >= Val(Nvl(mrsItems!最小值)) Then
        dblvalue = Val(strValue)
    Else
        '体温如果最小值无效，则输出的最小值为35
        If lng项目序号 = gint体温 And dblvalue < 35 Then dblvalue = 35
    End If
    
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
    If InStr(1, strValue, ";") <> 0 Then strValue = Split(strValue, ";")(1)
    If IsNumeric(strValue) = True And Val(strValue) <= Val(Nvl(mrsItems!最大值)) And Val(strValue) >= Val(Nvl(mrsItems!最小值)) Then dblvalue = Val(strValue)
    GetMaxValue = dblvalue
End Function

Private Sub ReadBoyData(ByVal blnAutoAdjust As Boolean)
    
    On Error GoTo Errhand
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
    Dim strSQL As String
    Dim lngColor As Long, lng行号 As Long, lng项目序号  As Long
    Dim str符号 As String
    Dim dbl数值 As Double, dblMinValue As Double, dblMaxValue As Double
    Dim strTmpString0 As String, strTmpString1 As String, strTmpString2 As String
    Dim strTime As String
    Dim blnAllow As Boolean
    Dim arrValues() As String
    Dim arrTmpValue() As Variant, arrTmpNote As Variant
    Dim i As Integer, j As Integer
    Dim int显示 As Integer
    Dim rs脉搏 As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim bln婴儿体温单显示出院 As Boolean, bln入科显示入院 As Boolean
    Dim lng体温不升显示方式 As Long
    Dim int标记 As Integer
    Dim lngSignColor As Long '参数:体温自动标识的颜色
    Dim lngNoRecordColor As Long '参数:未记说明显示颜色
    Dim bln入科不转入院 As Boolean
    
    On Error GoTo Errhand
    
    '71950:刘鹏飞,2014-06-11,体温单未记说明显示颜色
    lngNoRecordColor = Val(zlDatabase.GetPara("未记说明显示颜色", glngSys, 1255, "16711680"))
    '记录脉搏信息
    strFileds = "项目序号," & adDouble & ",18|数值," & adLongVarChar & ",4000|X坐标," & adDouble & ",5|时间," & adLongVarChar & ",20"
    Call Record_Init(rs脉搏, strFileds)
    
    '提取所有部位信息
    strSQL = "Select 项目序号, 部位,缺省项 From 体温部位"
    Call zlDatabase.OpenRecordset(rsPart, strSQL, "提取体温部位")
    
    
    '体温曲线项目需要增加字段用来决定该数据是否在三测单上显示,目前缺省为显示
    '-----------------------------------------------------------------------
    gstrSQL = "SELECT /*+ Rule*/  C.ID 序号,a.发生时间 As 时间,C.显示,C.记录内容 As 数值,C.体温部位,C.复试合格,D.记录名,E.保留项目,D.项目序号,DECODE(D.项目序号,-1,1,C.记录标记) 记录标记,C.未记说明 " & _
                "FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E,Table(Cast(f_num2list([6]) As zlTools.t_Numlist)) F " & _
                "Where B.ID=A.文件ID " & _
                    "And A.ID = C.记录ID " & _
                    "AND B.ID=[1] " & _
                    "AND B.病人id=[2] " & _
                    "AND B.主页id=[3] " & _
                    "AND D.项目序号=c.项目序号 " & _
                    "AND c.记录类型=1 " & _
                    "AND E.项目序号=D.项目序号 " & _
                    "AND F.COLUMN_VALUE=D.项目序号 " & _
                    "AND a.发生时间 BETWEEN [4] And [5] And c.终止版本 Is Null And D.记录法<>2" & _
                "Order By A.发生时间,DECODE(C.项目序号,-1,1,0),DECODE(D.项目序号,-1,1,C.记录标记)"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "读取曲线项目数据", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, CDate(mstr开始时间), CDate(mstr结束时间), T_BodyItem.str曲线项目)
        
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
                SinX = GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"))
                strTime = GetXCoordinateNew(SinX, Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"), False)
                SinX = GetXCoordinateNew(Format(Split(strTime, ",")(0), "YYYY-MM-DD HH:mm:ss"), Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss"))
                
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
                        !未记说明 & "|" & lngNoRecordColor & "|" & SinX & "|0|0|0|0|" & Val(zlCommFun.Nvl(!显示))
                   
                    If blnAdd Then
                        '提取接近中间时间点的值做为本列值
                         Call Record_Add(mrsNote, gstrFields, gstrValues)
                    Else
                        If (zlCommFun.Nvl(mrsNote!显示, 0) = 1 And zlCommFun.Nvl(!显示, 0) = 1) Or (zlCommFun.Nvl(mrsNote!显示, 0) <> 1 And zlCommFun.Nvl(!显示, 0) <> 1) Then
                             blnAllow = GetCanvasCenterNew(CDate(Format(mrsNote!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
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
                            blnAllow = GetCanvasCenterNew(CDate(Format(mrsPoint!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
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
                        blnAllow = GetCanvasCenterNew(CDate(Format(mrsPoint!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
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
    
    '处理物理降温数据,疼痛减痛
    For j = 0 To 1
        lng项目序号 = IIf(j = 0, gint体温, gint疼痛强度)
        arrTmpValue = Array()
        mrsPoint.Filter = "项目序号=" & lng项目序号 & " and 标记=0"
        With mrsPoint
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !数值 & ";" & !X坐标 & ";" & !Y坐标 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        mrsPoint.Filter = "项目序号=" & lng项目序号
        If mrsPoint.RecordCount > 0 Then mrsPoint.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            mrsPoint.Filter = "项目序号=" & lng项目序号 & " and 标记=1 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If mrsPoint.RecordCount <> 0 Then
                gstrFields = "备注": gstrValues = Val(Split(CStr(arrTmpValue(i)), ";")(2)) & "," & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & ";" & Val(Split(CStr(arrTmpValue(i)), ";")(4))
                Call Record_Update(mrsPoint, gstrFields, gstrValues, "序号|" & zlCommFun.Nvl(mrsPoint!序号))
            End If
        Next i
        
        arrTmpValue = Array()
        mrsPoint.Filter = "项目序号=" & lng项目序号 & " and 标记=1"
        With mrsPoint
            Do While Not .EOF
                ReDim Preserve arrTmpValue(UBound(arrTmpValue) + 1)
                arrTmpValue(UBound(arrTmpValue)) = !序号 & ";" & !项目序号 & ";" & !数值 & ";" & !X坐标 & ";" & !Y坐标 & ";" & Format(!时间, "yyyy-MM-dd HH:mm:ss")
            .MoveNext
            Loop
        End With
        
        mrsPoint.Filter = "项目序号=" & lng项目序号
        If mrsPoint.RecordCount > 0 Then mrsPoint.MoveFirst
        For i = 0 To UBound(arrTmpValue)
            mrsPoint.Filter = "项目序号=" & lng项目序号 & " and 标记=0 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
            If mrsPoint.RecordCount = 0 Then
                mrsPoint.Filter = "项目序号=" & lng项目序号 & " and 标记=1 and X坐标=" & Val(Split(CStr(arrTmpValue(i)), ";")(3)) & " And 时间='" & Format(Split(CStr(arrTmpValue(i)), ";")(5), "yyyy-MM-dd HH:mm:ss") & "'"
                mrsPoint.Delete
            End If
        Next i
    Next j
    
    
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
                    blnAllow = GetCanvasCenterNew(CDate(Format(mrsPoint!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstr开始时间, "YYYY-MM-DD HH:mm:ss")), SinX)
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
    
    '62989:刘鹏飞,2013-07-24,体温单标记显示颜色
    lngSignColor = Val(zlDatabase.GetPara("体温单标记显示颜色", glngSys, 1255, "255"))
    
    '读取手术、上下标信息
    '-----------------------------------------------------------------------
    gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"  '入出转缺省是红色,上下标及未记说明缺省是蓝色
    gstrSQL = "" & _
             " Select A.发生时间 AS 时间,C.记录类型,C.项目序号,C.记录内容,C.项目名称,C.未记说明" & _
             " FROM 病人护理文件 B, 病人护理数据 A, 病人护理明细 C" & _
             " Where B.ID=A.文件ID And A.ID = C.记录ID AND B.ID=[1] AND Nvl(B.婴儿, 0)=[6] AND B.病人id=[2] AND B.主页id=[3] And c.终止版本 Is Null" & _
             " AND MOD(C.记录类型,10) <> 1  AND A.发生时间 BETWEEN [4]  And [5]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "读取手术、上下标等信息", T_Patient.lng文件ID, T_Patient.lng病人ID, T_Patient.lng主页ID, Int(CDate(mstr开始时间)), CDate(mstr结束时间), T_Patient.lng婴儿, T_Patient.lng护理等级)
    
    If mblnMoved Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
        gstrSQL = Replace(gstrSQL, "病人护理数据", "H病人护理数据")
        gstrSQL = Replace(gstrSQL, "病人护理明细", "H病人护理明细")
    End If
        
    With rsData
        Do While Not .EOF
            bytShow = 1
            str内容 = Trim(zlCommFun.Nvl(!记录内容))
            
            lng行号 = IIf(!记录类型 = 2, 10, IIf(!记录类型 = 6, 11, 14))
            
            '对于手术显示需要特殊处理
            If !记录类型 = 4 Then
                str内容 = Trim(zlCommFun.Nvl(!项目名称))
                
                If str内容 = "分娩" Then
                    bytShow = T_BodyFlag.分娩
                ElseIf str内容 = "回室" Then
                    bytShow = T_BodyFlag.回室
                Else
                    bytShow = T_BodyFlag.手术
                End If
                
                If bytShow = 2 And Not blnAutoAdjust Then
                    str内容 = str内容 & gstrCaveSplit & ConvertTimeToChinese(Format(!时间, "HH:mm"))
                Else
                    str内容 = !项目名称
                End If
                lngColor = lngSignColor
            Else
                lngColor = IIf(Not IsNumeric(Nvl(!未记说明)), RGB_BLUE, Val(Nvl(!未记说明)))
            End If
            
            If bytShow > 0 Then
                SinX = Val(GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), mstr开始时间))
                
                mrsNote.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=" & !记录类型 & " And 时间='" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "'"
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
    
    bln婴儿体温单显示出院 = (zlDatabase.GetPara("婴儿体温单显示出院信息", glngSys, 1255, 1) = 1)
    '问题号:63525,修改人:李涛,入院标识不显示，入科标识显示时，不自动转为入院。
    bln入科不转入院 = (zlDatabase.GetPara("入科标识不自动转换为入院", glngSys, 1255, 1) = 0)
    
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
                Case 15
                    bytShow = T_BodyFlag.转病区
                End Select
                 
                If bytShow > 0 Then
                    '目前3，4 针对于转科 3-显示说明和科室 4 显示说明，科室，时间
                    If lng行号 = 9 And bln入科显示入院 = True And bln入科不转入院 = True Then
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
                    
                    SinX = Val(GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), mstr开始时间))
                    mrsNote.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=3 And 时间='" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "'"
                    
                    If mrsNote.BOF Then
                        gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|3|" & _
                            str内容 & "|" & lngSignColor & "|" & SinX & "|0|0|0|0"
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
                        
                        SinX = Val(GetXCoordinateNew(Format(!时间, "YYYY-MM-DD HH:mm:ss"), mstr开始时间))
                        mrsNote.Filter = "X坐标=" & SinX & " and 项目序号=" & lng行号 & " and 类型=13 And 时间='" & Format(!时间, "yyyy-MM-dd HH:mm:ss") & "'"
                        
                        If mrsNote.BOF Then
                            gstrValues = Format(!时间, "yyyy-MM-dd HH:mm:ss") & "|" & lng行号 & "|13|" & _
                                str内容 & "|" & lngSignColor & "|" & SinX & "|0|0|0|0"
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
    bytTag = Abs(Val(zlDatabase.GetPara("未记说明显示位置", glngSys, 1255, "0")))
    lng体温不升显示方式 = Val(zlDatabase.GetPara("体温不升显示方式", glngSys, 1255, "0"))
    '处理体温不升 体温不升始终显示在 35 度下面，只有未记说明显示在下面的情况，才将不升放入未记说明中，其它情况都放在下标中
    If Left(strTmpString0, 1) = ";" Then
        gstrFields = "时间|项目序号|类型|内容|颜色|X坐标|Y坐标|高度|打印X坐标|禁用"
        If lng体温不升显示方式 = 0 Or lng体温不升显示方式 = 2 Then
            arrValues = Split(strTmpString0, "|")
            arrValues(3) = "↓ "
            strTmpString0 = Join(arrValues, "|")
        End If
        strTmpString0 = Mid(strTmpString0, 2)
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
    
    mint心率应用 = 0
    If Not (mrsItems Is Nothing) Then If mrsItems.State = 1 Then mrsItems.Close
    '打开现存在适用该病人的护理记录项目
    gstrSQL = " Select C.项目序号,C.项目名称,C.项目类型,C.项目性质,C.项目长度,C.项目小数,C.项目表示,C.项目单位,C.项目值域,A.最大值,A.最小值,A.临界值,A.记录符,A.记录色,C.护理等级,C.应用方式,C.适用病人" & _
              " From 体温记录项目 A,护理记录项目 C,Table(Cast(f_num2list([1]) As zlTools.t_Numlist))  D" & _
              " where A.项目序号(+)=C.项目序号" & _
              " And C.项目序号=D.COLUMN_VALUE " & _
              " Order by C.项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "取开始行", T_BodyItem.str曲线项目 & "," & T_BodyItem.str表格内容)
    mrsItems.Filter = "项目序号=-1"
    If mrsItems.RecordCount > 0 Then mint心率应用 = zlCommFun.Nvl(mrsItems("应用方式").Value, 2): mrsItems.Filter = ""
    
    If Not mrsGraph Is Nothing Then If mrsGraph.State = 1 Then mrsGraph.Close
    
    lngMax = mobjBuffer.ScaleWidth \ gintBmpW      '一行能保存多少个图片?
    '所有需输出的图形序号(包括体温重叠标记),全部提取在picBuffer中,此处保存各项目的部位及其对应的图形序号
    gstrFields = "项目序号," & adDouble & ",18|部位," & adLongVarChar & ",50|记录符," & adLongVarChar & ",50|" & _
                 "记录色," & adDouble & ",18|重叠项目," & adLongVarChar & ",20|行," & adDouble & ",5|列," & adDouble & ",5"    '重叠项目应按项目序号大小排列,如:1,4,5
    Call Record_Init(mrsGraph, gstrFields)
    
    '先根据体温部位装载
    gstrSQL = " Select 项目序号,'' AS 部位, 记录符 标记符号,记录色 标记颜色,1 展现方式,'空' AS 重叠项目 From 体温记录项目 Order by 项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取各曲线项目的展现方式")
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
    Set rsOverlap = zlDatabase.OpenSQLRecord(gstrSQL, "再根据体温重叠标记装载")
    gstrSQL = " Select 序号,上级序号,项目序号,体温部位 From 体温重叠标记 Where 项目序号 is not null Order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取重叠从属项目")
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
            lngCurX = GetXCoordinateNew(!时间, strBeginDate)
            If mint心率应用 = 2 And !项目序号 = -1 Then
                mrsPoint.Filter = "项目序号=" & gint脉搏 & " And  X坐标<=" & !X坐标
            Else
                If Val(!项目序号) = gint体温 Or Val(!项目序号) = gint疼痛强度 Then
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
            If Not ((Val(!项目序号) = gint体温 Or Val(!项目序号) = gint疼痛强度) And Val(zlCommFun.Nvl(!标记)) = 1) Then
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
        T_LableRect.Left = T_LableRect.Left - 1
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
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngCount As Long
    On Error GoTo Errhand
    
    strSQL = " Select /*+ Rule*/ Count(*) 记录" & _
             " From 体温记录项目 A, 护理记录项目 B,Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) C" & _
             " Where A.项目序号=B.项目序号 And B.项目序号=C.COLUMN_VALUE" & _
             " Order by B.项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "取开始行", T_BodyItem.str曲线项目)
    
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
    Dim strSQL As String, strNewSql As String
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
    Dim strMarkDate As String, strFileBeginTime As String
    Dim arrParam() As String
    '----------------------------------------------------
    '病人信息变量
    '----------------------------------------------------
    Dim lng文件ID As Long, lng病人ID As Long, lng主页ID  As Long
    Dim lng科室ID As Long, lng婴儿  As Long, lng护理等级 As Long
    '----------------------------------------------------
    '体温单样式变量
    '----------------------------------------------------
    Dim MT_BodyStyle As type_BodyStyle
    Dim MT_BodyItem As type_BodyItem
    
    On Error GoTo ErrHandle
    
    If strParam <> "" Then
        arrParam = Split(strParam, ";")
        If UBound(arrParam) < 2 Then
            MsgBox "strParam参数不为空时,必须传入文件ID;病人ID;主页ID！", vbInformation, gstrSysName
            Exit Function
        End If
        lng文件ID = Val(arrParam(0))
        lng病人ID = Val(arrParam(1))
        lng主页ID = Val(arrParam(2))
        If UBound(arrParam) > 2 Then lng科室ID = Val(arrParam(3))
        If UBound(arrParam) > 3 Then lng婴儿 = Val(arrParam(4))
        lng护理等级 = 3
        gstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "护理等级", lng病人ID, lng主页ID)
        If rsTmp.BOF = False Then lng护理等级 = Nvl(rsTmp("护理等级"), 3)
    Else
        lng文件ID = T_Patient.lng文件ID
        lng病人ID = T_Patient.lng病人ID
        lng主页ID = T_Patient.lng主页ID
        lng科室ID = T_Patient.lng科室ID
        lng婴儿 = T_Patient.lng婴儿
        lng护理等级 = T_Patient.lng护理等级
    End If
    
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
    strSQL = "Select 格式 From 病历页面格式 Where 种类 = 3 And 编号 In (Select A.页面 From 病历文件列表 A,病人护理文件 B Where A.Id = B.格式ID and B.ID=[1])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取文件打印设置", lng文件ID)
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
        If Trim(zlDatabase.GetPara("体温单打印机", glngSys, 1255, "")) = "" Then
            MsgBox "没有设置打印机,将使用系统默认打印机设置！", vbInformation, gstrSysName
            strPrintName = Printer.DeviceName
        Else
            strPrintName = Trim(zlDatabase.GetPara("体温单打印机", glngSys, 1255, Printer.DeviceName))
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
    gPrinter.intBin = Val(zlDatabase.GetPara("体温单进纸", glngSys, 1255, Printer.PaperBin))
    
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
    '在读取文件之前首先将之前文件的样式保存下来
    With MT_BodyStyle
        .lng开始时点 = T_BodyStyle.lng开始时点
        .lng时间间隔 = T_BodyStyle.lng时间间隔
        .lng监测次数 = T_BodyStyle.lng监测次数
        .lng天数 = T_BodyStyle.lng天数
        .lng刻度宽度 = T_BodyStyle.lng刻度宽度
        .lng曲线列宽 = T_BodyStyle.lng曲线列宽
        .lng曲线行高 = T_BodyStyle.lng曲线行高
        .lng表格高度 = T_BodyStyle.lng表格高度
        .str列头名称 = T_BodyStyle.str列头名称
        .str标题文本 = T_BodyStyle.str标题文本
        .str标题字体 = T_BodyStyle.str标题字体
        .lng曲线空行 = T_BodyStyle.lng曲线空行
        .lng表格空行 = T_BodyStyle.lng表格空行
        .lng下表格高度 = T_BodyStyle.lng下表格高度
        .bln专科 = T_BodyStyle.bln专科
    End With
    With MT_BodyItem
        .str表格内容 = T_BodyItem.str表格内容
        .str表格项目 = T_BodyItem.str表格项目
        .str曲线项目 = T_BodyItem.str曲线项目
    End With
    '提取该文件的样式(不要删除：批量打印需要从新提取)
    If Not GetStyleBody(lng文件ID, lng护理等级, lng婴儿, lng科室ID, blnPrint) Then Exit Function
    intBaby = lng婴儿
    '------------------------------------------------------------------------------------------------------------------
    lngBeginY = gPrinter.lngTop
    lngIndex = mintPage
    
    '如果只打印当前就只将开始和结束写同一页码
    Set mfrmCaseTendBodyPrint = New frmCaseTendBodyPrint
    Load frmTendFileRead
    Call frmTendFileRead.InitRechBox(lng文件ID)
    strMarkDate = ""
    '提取用户设置的体温单开始时间(婴儿还是以婴儿出生时间为准)
    strSQL = "select 开始时间 from 病人护理文件 where ID=[1] and 病人ID=[2] and 主页id=[3] and nvl(婴儿,0)=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "提取体温单开始时间", lng文件ID, lng病人ID, lng主页ID, lng婴儿)
    If rsTmp.RecordCount <> 0 Then
        strMarkDate = Format(rsTmp!开始时间, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If strMarkDate <> "" Then strMarkDate = "to_date('" & strMarkDate & "','yyyy-MM-dd hh24:mi:ss')"
    
    '提取婴儿医嘱信息(转科，出院)存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "   (SELECT 病人ID,主页ID,婴儿时间,DECODE(nvl(婴儿,0),0, DECODE(NVL(出院日期,''),'',0,1), DECODE(NVL(婴儿时间,''),'',0,1))记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID,A.主页ID,B.开始执行时间 婴儿时间, A.出院日期,B.婴儿" & vbNewLine & _
                "           FROM 病案主页 A," & vbNewLine & _
                "               (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND nvl(B.婴儿,0)<>0 AND B.诊疗类别 = 'Z'" & vbNewLine & _
                "                AND Instr(',3,5,11,', ',' || c.操作类型 || ',') > 0 And  B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "           WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "           ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
    '说明:目前有了专科体温单，病人可能同时存在多份体温单。体温单开始时间和终止时间的规则如下:
    '如果文件的开始时间不为空并且大于等于病人入院时间或婴儿出生时间,体温单的开始时间以文件开始时间为准,否则以病人入院时间或婴儿出生时间为准
    '如果文件的终止时间不为空并且小于等于病人或婴儿出院时间（未出院不能不能大于当前时间）,体温单结束时间以文件开始时间为准，否则体温单结束时间以病人或婴儿出院时间为准(未出院为当前时间)
    '如果文件的终止时间为空,保持原有方式,病人如果已经出院，就已出院时间为准,未出院就已当前时间或数据结束时间为准.
    '读取此病人的体温单总页数
    '------------------------------------------------------------------------------------------------------------------
    strSQL = " SELECT  入院时间,出院时间,1 + TRUNC((TO_DATE(TO_CHAR(出院时间,'yyyy-MM-dd'),'yyyy-MM-dd') -TO_DATE(TO_CHAR(入院时间,'yyyy-MM-dd'),'yyyy-MM-dd')) / " & T_BodyStyle.lng天数 & ") AS 页数,发生时间 " & _
            "  From (" & _
                " SELECT DECODE(D.开始时间,NULL,DECODE(C.出生时间,NULL," & IIf(strMarkDate = "", "B.入院时间", strMarkDate) & ",C.出生时间)," & vbNewLine & _
                "               DECODE(SIGN(D.开始时间 - DECODE(C.出生时间,NULL," & IIf(strMarkDate = "", "B.入院时间", strMarkDate) & ",C.出生时间))," & vbNewLine & _
                "                      1," & vbNewLine & _
                "                      D.开始时间," & vbNewLine & _
                "                      DECODE(C.出生时间,NULL," & IIf(strMarkDate = "", "B.入院时间", strMarkDate) & ",C.出生时间))) AS 入院时间," & vbNewLine & _
                "    DECODE(D.结束时间,NULL," & vbNewLine & _
                "               DECODE(E.记录,0," & vbNewLine & _
                "                      DECODE(SIGN(NVL(E.婴儿时间, B.出院时间) - D.发生时间), 1, NVL(E.婴儿时间, B.出院时间), D.发生时间)," & vbNewLine & _
                "                      NVL(E.婴儿时间, B.出院时间))," & vbNewLine & _
                "               DECODE(SIGN(NVL(E.婴儿时间, B.出院时间) - D.结束时间), 1, D.结束时间, NVL(E.婴儿时间, B.出院时间))) 出院时间," & vbNewLine & _
                "    D.发生时间" & vbNewLine & _
                "    FROM (SELECT 病人ID,主页ID,MIN(开始时间) AS 入院时间," & vbNewLine & _
                "    MAX(NVL(终止时间, SYSDATE)) AS 出院时间" & vbNewLine & _
                "    FROM 病人变动记录" & vbNewLine & _
                "    WHERE 开始时间 IS NOT NULL AND 病人ID = [2] AND 主页ID = [3] GROUP BY 病人ID,主页ID) B," & vbNewLine & _
                "    (SELECT 病人ID,主页ID,出生时间 FROM 病人新生儿记录 WHERE 病人ID =[2] AND 主页ID =[3] AND 序号=[4]) C ," & vbNewLine & _
                "    (SELECT NVL(发生时间, SYSDATE) 发生时间, 开始时间, 结束时间" & vbNewLine & _
                "       FROM (SELECT MAX(B.发生时间) 发生时间, MAX(A.开始时间) 开始时间, MAX(A.结束时间) 结束时间" & vbNewLine & _
                "              FROM 病人护理文件 A, 病人护理数据 B" & vbNewLine & _
                "              WHERE A.ID = B.文件ID(+) AND A.ID = [1] AND A.病人ID = [2] AND A.主页ID = [3] AND A.婴儿 = [4])) D," & vbNewLine & _
                strNewSql & vbNewLine & _
                "WHERE B.病人ID=E.病人ID And B.主页ID=E.主页ID And B.病人ID=C.病人ID(+) AND B.主页ID=C.主页ID(+))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, mstrTitle, lng文件ID, lng病人ID, lng主页ID, lng婴儿)
    intCount = 0
    For intCOl = 0 To rsTmp("页数").Value - 1
    
        strDateFrom = Format(rsTmp("入院时间").Value + T_BodyStyle.lng天数 * intCOl, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("入院时间").Value + T_BodyStyle.lng天数 * (intCOl + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
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
        If PrintOrPreviewBodyStateNew(objPrint, lng病人ID, lng主页ID, lng文件ID, intBaby, _
                lng科室ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, False, _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, , mblnMoved) = True Then
                
                If blnPrint = False Then
                    mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, lng病人ID, lng主页ID, _
                        lng文件ID, CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                        CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, strPage, lng科室ID, lng婴儿
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
            If PrintOrPreviewBodyStateNew(objPrint, lng病人ID, lng主页ID, lng文件ID, intBaby, _
                lng科室ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, intCOl <> lngIndex, _
                CInt(Split(strArrFromTo(intCOl), ";")(1)), CInt(Split(strArrFromTo(intCOl), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "未知错误，打印失败！", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                'Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                '68407:刘鹏飞,2013-12-05,修改intCOl = UBound(strArrFromTo)为intCOl=lngIndexEnd,不然会导致打印机挂起
                If intCOl = lngIndexEnd Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, lng病人ID, lng主页ID, _
            lng文件ID, CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, strPage, lng科室ID, lng婴儿
        Else '连续打印是记录打印的开始页号和结束页号
            strSQL = "zl_体温单数据_Printer(" & lng文件ID & "," & lngIndex + 1 & "," & lngIndexEnd + 1 & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, "zl_体温单数据_Printer")
        End If
        
    Case 2          '从第一页连续打印,即全部打印
        strPage = 0
        For intCOl = 0 To UBound(strArrFromTo)
            If blnPrint = True Then Printer.Print ""
            If PrintOrPreviewBodyStateNew(objPrint, lng病人ID, lng主页ID, lng文件ID, intBaby, _
                lng科室ID, lngBeginY * conRatemmToTwip, gPrinter.lngLeft, Me, intCOl <> 0, _
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
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, gPrinter.lngLeft, Me, lng病人ID, lng主页ID, _
            lng文件ID, CInt(Split(strArrFromTo(0), ";")(1)), _
                CInt(Split(strArrFromTo(0), ";")(1)), intPageNo, strArrFromTo, strPage, lng科室ID, lng婴儿
        End If
    End Select
    
    'WinNT自定义纸张处理
    If IsWindowsNT And gPrinter.intPage = 256 Then DelCustomPaper
    
    Unload frmTendFileRead
    
    '------------------------------------------------------------------------------------------------------------------
ReStoreCuve:
    '在预览、打印完成后恢复之前选择文件的样式(批量打印可能导致之前的文件样式发生变化)
    With T_BodyStyle
        .lng开始时点 = MT_BodyStyle.lng开始时点
        .lng时间间隔 = MT_BodyStyle.lng时间间隔
        .lng监测次数 = MT_BodyStyle.lng监测次数
        .lng天数 = MT_BodyStyle.lng天数
        .lng刻度宽度 = MT_BodyStyle.lng刻度宽度
        .lng曲线列宽 = MT_BodyStyle.lng曲线列宽
        .lng曲线行高 = MT_BodyStyle.lng曲线行高
        .lng表格高度 = MT_BodyStyle.lng表格高度
        .str列头名称 = MT_BodyStyle.str列头名称
        .str标题文本 = MT_BodyStyle.str标题文本
        .str标题字体 = MT_BodyStyle.str标题字体
        .lng曲线空行 = MT_BodyStyle.lng曲线空行
        .lng表格空行 = MT_BodyStyle.lng表格空行
        .lng下表格高度 = MT_BodyStyle.lng下表格高度
        .bln专科 = MT_BodyStyle.bln专科
    End With
    With T_BodyItem
        .str表格内容 = MT_BodyItem.str表格内容
        .str表格项目 = MT_BodyItem.str表格项目
        .str曲线项目 = MT_BodyItem.str曲线项目
    End With
    Call InitPara(T_BodyStyle.bln专科)
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    GoTo ReStoreCuve
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

Private Sub vsf_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call RaiseShowTipInfo(vsf.Body, 2, X, Y)
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

Private Sub DrawDownTabAnsyGrade(ByVal lngDc As Long, ByVal objDraw As Object, arrText() As String, ByVal Row As Long, ByVal Col As Long, _
    ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean, Optional ByVal blnFormat As Boolean = False)
'---------------------------------------------------
'功能 大便次数输出
'说明 AnsyGrade=True才能调用此函数
'---------------------------------------------------
    Dim lngFont As Long, lngOldFont As Long, intSize As Integer, intOldSize As Integer
    Dim lngBrush As Long, lngOldBrush As Long
    Dim lngBackColor As Long, lngForeColor As Long
    Dim stdSet As StdFont, stdOldset As StdFont
    Dim LPoint As T_LPoint, T_ClientRect As RECT
    Dim str1 As String, str2 As String, str3 As String, strTmp As String
    Dim lngX As Long, lngY As Long, sngH As Single, sngW As Single
    Dim lngMaxWidth As Long
    
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
    lngOldBrush = SelectObject(lngDc, lngBrush)
    Call FillRect(lngDc, T_ClientRect, lngBrush)
    '立即销毁临时使用的刷子并还原刷子
    Call SelectObject(lngDc, lngOldBrush)
    Call DeleteObject(lngBrush)
    
    str1 = arrText(0): str2 = arrText(1): str3 = arrText(2)
    If blnFormat = True Then
        '60529:刘鹏飞,2013-04-19
        If objDraw.TextWidth(str2) > objDraw.TextWidth(str3) Then
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
    Set stdSet = New StdFont
    stdSet.Name = "宋体"
    stdSet.Size = intSize
    stdSet.Bold = False
    Set stdOldset = stdSet '原始字体
    
    Call GetTextRect(objDraw, LPoint.X, LPoint.Y, strTmp, LPoint.W, True, , 1)
    '输出左边
    If str1 <> "" Then
        Call SetFontIndirect(stdOldset, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngForeColor)
        Call DrawText(lngDc, str1, -1, T_LableRect, 0)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        lngX = T_LableRect.Left + (objDraw.TextWidth(str1) / T_TwipsPerPixel.X) - (objDraw.TextWidth("a") / T_TwipsPerPixel.X / 2) + 1
        Call ReleaseFontIndirect(objDraw)
    Else
        lngX = T_LableRect.Left
    End If

    If blnFormat = True Then '分子分母显示
        intSize = 7
        objDraw.Font.Size = intSize
        '60529:刘鹏飞,2013-04-19
        If objDraw.TextWidth(str2) > objDraw.TextWidth(str3) Then
            lngMaxWidth = objDraw.TextWidth(str2) / T_TwipsPerPixel.X
        Else
            lngMaxWidth = objDraw.TextWidth(str3) / T_TwipsPerPixel.X
        End If
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = intSize
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngForeColor)
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str2) / T_TwipsPerPixel.X) \ 2
        lngY = T_LableRect.Top
        sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.X / 2
        T_LableRect.Top = lngY - sngH
        'If T_LableRect.Top < Top Then T_LableRect.Top = Top - 1
        T_LableRect.Bottom = T_ClientRect.Bottom
        Call DrawText(lngDc, str2, -1, T_LableRect, 0)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        lngY = T_LableRect.Top + (objDraw.TextHeight("A") / T_TwipsPerPixel.Y)
        Call ReleaseFontIndirect(objDraw)
        '画横线
        objDraw.Font.Size = intOldSize
        Call DrawLine(lngDc, lngX, lngY, lngX + lngMaxWidth, lngY)
        '输出分母
        intSize = 7
        objDraw.Font.Size = intSize
        lngY = lngY
        T_LableRect.Left = lngX + (lngMaxWidth - objDraw.TextWidth(str3) / T_TwipsPerPixel.X) \ 2
        T_LableRect.Top = lngY
        Set stdSet = New StdFont
        stdSet.Name = "宋体"
        stdSet.Size = intSize
        Call SetFontIndirect(stdSet, lngDc, objDraw)
        lngFont = CreateFontIndirect(T_Font)
        lngOldFont = SelectObject(lngDc, lngFont)
        Call SetTextColor(lngDc, lngForeColor)
        Call DrawText(lngDc, str3, -1, T_LableRect, 0)
        Call SelectObject(lngDc, lngOldFont)
        Call DeleteObject(lngFont)
        Call ReleaseFontIndirect(mobjDraw)
    Else
        If str1 <> "" Then
            '输出上标
            intSize = 7
            objDraw.Font.Size = intSize
            Set stdSet = New StdFont
            stdSet.Name = "宋体"
            stdSet.Size = intSize
            Call SetFontIndirect(stdSet, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngForeColor)
            T_LableRect.Left = lngX
            lngY = T_LableRect.Top
            sngH = objDraw.TextHeight("A") / T_TwipsPerPixel.Y / 2
            T_LableRect.Top = lngY - sngH
            If T_LableRect.Top < T_ClientRect.Top Then T_LableRect.Top = T_ClientRect.Top - 1
            Call DrawText(lngDc, str2, -1, T_LableRect, 0)
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            lngX = lngX + (objDraw.TextWidth(str2) / T_TwipsPerPixel.X)
            Call ReleaseFontIndirect(mobjDraw)
            '输出后半部分
            Call SetFontIndirect(stdOldset, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngForeColor)
            T_LableRect.Left = lngX
            T_LableRect.Top = lngY
            Call DrawText(lngDc, "/" & str3, -1, T_LableRect, 0)
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(mobjDraw)
        Else
            Call SetFontIndirect(stdOldset, lngDc, objDraw)
            lngFont = CreateFontIndirect(T_Font)
            lngOldFont = SelectObject(lngDc, lngFont)
            Call SetTextColor(lngDc, lngForeColor)
            Call DrawText(lngDc, str2 & "/" & str3, -1, T_LableRect, DT_CENTER)
            Call SelectObject(lngDc, lngOldFont)
            Call DeleteObject(lngFont)
            Call ReleaseFontIndirect(mobjDraw)
        End If
    End If
    
    objDraw.Font.Size = intOldSize
    Set stdSet = Nothing
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


Private Function GetStyleBody(ByVal lng文件ID As Long, ByVal lng护理等级 As Long, lng婴儿 As Long, lng科室ID As Long, Optional ByVal blnPrint As Boolean = False) As Boolean
'-------------------------------------------------------------------------------------------
'功能:获取文件体温单文件样式
'-------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim lng格式ID As Long
    Dim str表格项目 As String
    Dim str体温项目 As String
    Dim str表格内容 As String
    Dim lngTabRows As Long
    Dim i As Integer
    Dim sinTwipsPerPixelX As Single, sinTwipsPerPixelY As Single
    
    On Error GoTo Errhand
    
    If blnPrint = True Then
        sinTwipsPerPixelX = Printer.TwipsPerPixelX
        sinTwipsPerPixelY = Printer.TwipsPerPixelY
    Else
        sinTwipsPerPixelX = Screen.TwipsPerPixelX
        sinTwipsPerPixelY = Screen.TwipsPerPixelY
    End If
    
    gstrSQL = "Select A.格式ID,B.子类 From 病人护理文件 A, 病历文件列表 B Where a.格式id = b.Id And b.种类 = 3 And b.保留 = -1 And A.Id = [1]"
    If mblnMoved = True Then
        gstrSQL = Replace(gstrSQL, "病人护理文件", "H病人护理文件")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "病人护理文件", lng文件ID)
    lng格式ID = CLng(rsTemp!格式ID)
    T_Patient.lng格式ID = lng格式ID
    T_BodyStyle.bln专科 = (Nvl(rsTemp!子类, "0") = "1")
    If T_BodyStyle.bln专科 = True Then
        '表格样式构造数据
        gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, d.要素表示 " & _
                " From 病历文件结构 D, 病历文件结构 P" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '格式定义'" & _
                " Order By d.对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "体温单样式构造", lng格式ID)
        With rsTemp
            Do While Not .EOF
                Select Case "" & !要素名称
                    Case "开始时点"
                        T_BodyStyle.lng开始时点 = Val(Nvl(!内容文本))
                    Case "时间间隔"
                        T_BodyStyle.lng时间间隔 = Val(Nvl(!内容文本))
                    Case "监测次数"
                        T_BodyStyle.lng监测次数 = Val(Nvl(!内容文本))
                    Case "天数"
                        T_BodyStyle.lng天数 = Val(Nvl(!内容文本))
                    Case "刻度宽度"
                        T_BodyStyle.lng刻度宽度 = Fix(Val(Nvl(!内容文本)) / sinTwipsPerPixelX) * sinTwipsPerPixelX
                    Case "曲线列宽"
                        T_BodyStyle.lng曲线列宽 = Fix(Val(Nvl(!内容文本)) / sinTwipsPerPixelX) * sinTwipsPerPixelX
                    Case "曲线行高"
                        T_BodyStyle.lng曲线行高 = Fix(Val(Nvl(!内容文本)) / sinTwipsPerPixelY) * sinTwipsPerPixelY
                    Case "表格高度"
                        T_BodyStyle.lng表格高度 = Fix(Val(Nvl(!内容文本)) / sinTwipsPerPixelY) * sinTwipsPerPixelY
                    Case "列头名称"
                        T_BodyStyle.str列头名称 = Nvl(!内容文本)
                    Case "标题文本"
                        T_BodyStyle.str标题文本 = Nvl(!内容文本)
                    Case "标题字体"
                        T_BodyStyle.str标题字体 = Nvl(!内容文本)
                    Case "曲线空行"
                        T_BodyStyle.lng曲线空行 = Val(Nvl(!内容文本))
                    Case "表格高度1"
                        T_BodyStyle.lng下表格高度 = Fix(Val(Nvl(!内容文本)) / sinTwipsPerPixelY) * sinTwipsPerPixelY
                    Case "表格空行"
                        T_BodyStyle.lng表格空行 = Val(Nvl(!内容文本))
                End Select
                .MoveNext
            Loop
        End With
        
        '曲线项目定义数据
        gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, d.要素表示 " & _
            " From 病历文件结构 D, 病历文件结构 P " & _
            " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '曲线项目定义'" & _
            " Order By d.对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "体温项目定义", lng格式ID)
        i = 0: str体温项目 = ""
        With rsTemp
            Do While Not .EOF
                If i = 0 Then
                    str体温项目 = !内容文本
                Else
                    str体温项目 = str体温项目 & "," & !内容文本
                End If
                i = i + 1
                .MoveNext
            Loop
            T_BodyItem.str曲线项目 = str体温项目
        End With
        
        '表格项目定义数据
        gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, d.要素表示 " & _
            " From 病历文件结构 D, 病历文件结构 P " & _
            " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格项目定义'" & _
            " Order By d.对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "体温项目定义", lng格式ID)
        i = 0: str表格项目 = "": str表格内容 = ""
        With rsTemp
            Do While Not .EOF
                 If i = 0 Then
                    str表格项目 = Nvl(!内容文本) & ":" & Nvl(!要素表示)
                    str表格内容 = Nvl(!内容文本)
                 Else
                    str表格项目 = str表格项目 & "@" & Nvl(!内容文本) & ":" & Nvl(!要素表示)
                    str表格内容 = str表格内容 & "," & Nvl(!内容文本)
                 End If
                 i = i + 1
                .MoveNext
            Loop
            
            T_BodyItem.str表格内容 = str表格内容
            T_BodyItem.str表格项目 = GetString(str表格项目)
        End With
    Else '标准体温单
        '表格样式构造数据
        T_BodyStyle.lng开始时点 = Val(zlDatabase.GetPara("体温开始时间", glngSys, 1255, 4))
        T_BodyStyle.lng时间间隔 = 4
        T_BodyStyle.lng监测次数 = 6
        T_BodyStyle.lng天数 = 7
        T_BodyStyle.lng刻度宽度 = Fix(1350 / sinTwipsPerPixelX) * sinTwipsPerPixelX
        T_BodyStyle.lng曲线列宽 = Fix(225 / sinTwipsPerPixelX) * sinTwipsPerPixelX
        T_BodyStyle.lng曲线行高 = Fix(90 / sinTwipsPerPixelY) * sinTwipsPerPixelY
        T_BodyStyle.lng表格高度 = Fix(255 / sinTwipsPerPixelY) * sinTwipsPerPixelY
        T_BodyStyle.str列头名称 = "日       期@" & IIf(T_Patient.lng婴儿 = 0, "住 院 天 数", "出 生 天 数") & "@手术后天数@时       间"
        T_BodyStyle.str标题文本 = "体温单"
        T_BodyStyle.str标题字体 = "宋体,20"
        T_BodyStyle.lng曲线空行 = Val(zlDatabase.GetPara("体温曲线固定添加行数", glngSys, 1255, "0"))
        T_BodyStyle.lng下表格高度 = Fix(255 / sinTwipsPerPixelY) * sinTwipsPerPixelY
        T_BodyStyle.lng表格空行 = 0
        '提取项目信息
        gstrSQL = _
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
            " WHERE a.项目序号 = b.项目序号 AND a.项目序号 = c.项目序号(+) AND NVL(a.应用方式,0) <> 0 AND a.项目性质 = 1  and A.护理等级>=[1]" & vbNewLine & _
            " And nvl(A.适用病人,0) in (0,[2]) and (A.适用科室=1 or (A.适用科室=2 and Exists (select 1 from 护理适用科室 D where A.项目序号=D.项目序号 and D.科室ID=[3])))" & vbNewLine & _
            " ORDER BY Decode(b.记录法, 2, 2, 1), Decode(b.项目序号, 3, Decode(b.记录法, 2, 1, b.排列序号), b.排列序号)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取体温项目", lng护理等级, IIf(lng婴儿 = 0, 1, 2), lng科室ID)
        '体温曲线项目
        rsTemp.Filter = "记录法=1 OR 记录法=3"
        rsTemp.Sort = "记录法,排列序号"
        i = 0: str体温项目 = ""
        With rsTemp
            Do While Not .EOF
                If i = 0 Then
                    str体温项目 = !项目序号
                Else
                    str体温项目 = str体温项目 & "," & !项目序号
                End If
                i = i + 1
                .MoveNext
            Loop
            T_BodyItem.str曲线项目 = str体温项目
        End With
        
        '体温表格项目
        rsTemp.Filter = "记录法=2 And 项目序号<>5"
        rsTemp.Sort = "排列序号"
        i = 0: str表格项目 = "": str表格内容 = "": lngTabRows = 0
        With rsTemp
            Do While Not .EOF
                If i = 0 Then
                   If Val(!项目序号) = 4 Then
                       str表格项目 = "4,5:" & Nvl(!记录频次)
                       str表格内容 = "4,5"
                   Else
                       str表格项目 = Nvl(!项目序号) & ":" & Nvl(!记录频次)
                       str表格内容 = Nvl(!项目序号)
                   End If
                Else
                   If Val(!项目序号) = 4 Then
                       str表格项目 = str表格项目 & "@" & "4,5" & ":" & Nvl(!记录频次)
                       str表格内容 = str表格内容 & "," & "4,5"
                   Else
                       str表格项目 = str表格项目 & "@" & Nvl(!项目序号) & ":" & Nvl(!记录频次)
                       str表格内容 = str表格内容 & "," & Nvl(!项目序号)
                   End If
                End If
                 '计算表格所占用的总行数
                If Val(!项目序号) = 3 Then '说明呼吸为表格项目
                    lngTabRows = lngTabRows + 1
                Else
                    Select Case Val(Nvl(!记录频次, 2))
                    Case 3
                        lngTabRows = lngTabRows + 3
                    Case 4
                        lngTabRows = lngTabRows + 2
                    Case Else
                        lngTabRows = lngTabRows + 1
                    End Select
                End If
                 i = i + 1
                .MoveNext
            Loop
            T_BodyItem.str表格内容 = str表格内容
            T_BodyItem.str表格项目 = GetString(str表格项目)
        End With
        T_BodyStyle.lng表格空行 = Val(zlDatabase.GetPara("体温表格行数", glngSys, 1255, 8)) - lngTabRows
    End If
    Call GetPainDegreeNO
    GetStyleBody = True
    Exit Function
Errhand:
    If ErrCenter() Then
        Resume
    End If
End Function

Private Function GetString(ByVal strValue As String) As String
    Dim strOld() As String
    Dim strNew As String
    Dim str血压 As String
    Dim i As Integer
    
    strOld = Split(strValue, "@")
    For i = 0 To UBound(strOld)
        If InStr(strOld(i), ",") > 0 Then
            str血压 = Split(strOld(i), ",")(0) & ":" & Split(strOld(i), ":")(1)
            str血压 = str血压 & "," & Split(strOld(i), ",")(1)
        Else
            If i = 0 Then
                strNew = strOld(i)
            Else
                strNew = strNew & "," & strOld(i)
            End If
        End If
    Next
    If str血压 = "" Then
        GetString = strNew
    Else
        GetString = strNew & "," & str血压
    End If
End Function

Private Function GetSymbol(ByVal lng项目序号 As Long, ByVal str部位 As String, Optional ByVal str重叠项目 As String = "空", Optional ByVal str符号 As String = "") As Boolean

    'bln作图区域=True,体温单作图区域,计算后居中显示;否则,照传入的数据显示
    Dim blnGraph As Boolean
    Dim bln重叠 As Boolean
    Dim str记录符 As String

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
    
    If mrsGraph.RecordCount = 0 Then Exit Function
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
    End If

    
    mrsGraph.Filter = ""
    
    If str记录符 <> "○" Then
        GetSymbol = False
    Else
        GetSymbol = True
    End If

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

