VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmNewCheckMain 
   BackColor       =   &H80000005&
   Caption         =   "药品盘点管理"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12690
   Icon            =   "frmNewCheckMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9000
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   2
      Top             =   7560
      Width           =   2175
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "正常"
         Height          =   180
         Left            =   1680
         TabIndex        =   6
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "正常冲销"
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   37
         Width           =   720
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   1605
      Left            =   10320
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2831
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   12615
      Begin VB.CommandButton cmd重置 
         Caption         =   "重置(&F)"
         Height          =   350
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   840
         Width           =   1100
      End
      Begin VB.TextBox Txt审核人 
         Height          =   300
         Left            =   11160
         MaxLength       =   8
         TabIndex        =   31
         Top             =   120
         Width           =   1365
      End
      Begin VB.TextBox Txt填制人 
         Height          =   300
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   29
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton Cmd药品 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   300
         Left            =   12240
         TabIndex        =   37
         Top             =   510
         Width           =   255
      End
      Begin VB.TextBox Txt药品 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8970
         MaxLength       =   50
         ScrollBars      =   3  'Both
         TabIndex        =   36
         Top             =   510
         Width           =   3255
      End
      Begin VB.CheckBox Chk药品 
         BackColor       =   &H80000003&
         Caption         =   "药品"
         Height          =   300
         Left            =   8280
         TabIndex        =   35
         Top             =   510
         Width           =   735
      End
      Begin VB.TextBox txt结束NO 
         Height          =   300
         Left            =   6450
         MaxLength       =   8
         TabIndex        =   27
         Top             =   120
         Width           =   1605
      End
      Begin VB.TextBox txt开始No 
         Height          =   300
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   25
         Top             =   120
         Width           =   1605
      End
      Begin VB.CheckBox chkStrike 
         BackColor       =   &H80000003&
         Caption         =   "包含冲销"
         Enabled         =   0   'False
         Height          =   300
         Left            =   8280
         TabIndex        =   41
         Top             =   907
         Width           =   1095
      End
      Begin VB.CommandButton cmd确认 
         Caption         =   "确认(&S)"
         Height          =   350
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   840
         Width           =   1100
      End
      Begin VB.ComboBox cbo已审核 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   900
         Width           =   1560
      End
      Begin VB.ComboBox cbo未审核 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   510
         Width           =   1560
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1560
         TabIndex        =   23
         Text            =   "cboStock"
         Top             =   120
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Index           =   0
         Left            =   4560
         TabIndex        =   33
         Top             =   503
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   200736771
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Index           =   0
         Left            =   6465
         TabIndex        =   34
         Top             =   503
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   200736771
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Index           =   1
         Left            =   4560
         TabIndex        =   39
         Top             =   900
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   200736771
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Index           =   1
         Left            =   6465
         TabIndex        =   40
         Top             =   900
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   200736771
         CurrentDate     =   36263
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "审核人"
         Height          =   180
         Left            =   10560
         TabIndex        =   30
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "填制人"
         Height          =   180
         Left            =   8280
         TabIndex        =   28
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "～"
         Height          =   180
         Index           =   1
         Left            =   6225
         TabIndex        =   26
         Top             =   180
         Width           =   180
      End
      Begin VB.Label LblNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "No"
         Height          =   180
         Left            =   3660
         TabIndex        =   24
         Top             =   180
         Width           =   180
      End
      Begin VB.Label lbl已审核 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "已审核单据"
         Height          =   180
         Left            =   420
         TabIndex        =   20
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lbl未审核 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "未审核单据"
         Height          =   180
         Left            =   420
         TabIndex        =   19
         Top             =   570
         Width           =   900
      End
      Begin VB.Label lblStock 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "库      房"
         Height          =   180
         Left            =   420
         TabIndex        =   22
         Top             =   180
         Width           =   900
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "～"
         Height          =   180
         Index           =   3
         Left            =   6225
         TabIndex        =   11
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "审核日期"
         Height          =   180
         Index           =   1
         Left            =   3660
         TabIndex        =   10
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl至 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "～"
         Height          =   180
         Index           =   0
         Left            =   6225
         TabIndex        =   9
         Top             =   570
         Width           =   180
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "填制日期"
         Height          =   180
         Index           =   0
         Left            =   3660
         TabIndex        =   8
         Top             =   570
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Height          =   5415
      Left            =   1680
      ScaleHeight     =   5355
      ScaleWidth      =   8475
      TabIndex        =   1
      Top             =   1800
      Width           =   8535
      Begin VB.CommandButton Cmd查阅 
         Caption         =   "查阅(&V)"
         Height          =   350
         Left            =   7320
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1100
      End
      Begin VB.PictureBox picSeparate_s 
         BorderStyle     =   0  'None
         Height          =   370
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   375
         ScaleWidth      =   7935
         TabIndex        =   13
         Top             =   2520
         Width           =   7935
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            Caption         =   "金额差合计："
            Height          =   180
            Left            =   1680
            TabIndex        =   18
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            Caption         =   "盘点金额合计："
            Height          =   180
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label lbl3 
            AutoSize        =   -1  'True
            Caption         =   "账面金额差合计："
            Height          =   180
            Left            =   3000
            TabIndex        =   16
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label lblSum成本金额 
            AutoSize        =   -1  'True
            Caption         =   "盘点成本金额合计："
            Height          =   180
            Left            =   4680
            TabIndex        =   15
            Top             =   120
            Width           =   1620
         End
         Begin VB.Label lbl成本金额差 
            AutoSize        =   -1  'True
            Caption         =   "成本金额差合计："
            Height          =   180
            Left            =   6480
            TabIndex        =   14
            Top             =   120
            Width           =   1440
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1455
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   6255
         _cx             =   11033
         _cy             =   2566
         Appearance      =   1
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1155
         Left            =   0
         TabIndex        =   46
         Top             =   4200
         Width           =   6255
         _cx             =   11033
         _cy             =   2037
         Appearance      =   1
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
         BackColorSel    =   16053482
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   15724527
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   0   'False
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7800
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcDetail 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
      _Version        =   589884
      _ExtentX        =   2566
      _ExtentY        =   1720
      _StockProps     =   64
      Enabled         =   -1  'True
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   1920
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmNewCheckMain.frx":06EA
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmNewCheckMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl
Private mcbrMenuBar As CommandBarPopup
Private mcbrToolBar As CommandBar

'盘点表（单）
Private Const mconTab_CheckCourseCard = 0                 '盘点记录单列表
Private Const mconTab_CheckCard = 1                       '盘点列表

Private mblnLoad As Boolean
Private mbln绑定 As Boolean      '判断是否绑定完成
Private mintLastIndex As Integer '保存上一次点击的Tab
Private mstrSelectTag As String     '当前选择的是填制人还是审核人

Private mlngMode As Long
Private mstrFind As String
Private mblnBootUp As Boolean
Private mlastRow As Long                '上次电击的行
Private mstrTitle As String             '窗体的标题
Private mblnViewCost As Boolean         '查看成本价
'Private Const mstrTitle As String = "药品盘点管理"

Public mstrPrivs As String              '权限

'日期设置
Private mstrStart As Date
Private mstrEnd As Date
Private mstrVerifyStart As Date
Private mstrVerifyEnd As Date

Private mlng库房ID As Long
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private Const mcstComment As String = "黑-盘平;红-盘盈;蓝-盘亏;粗体-停用药品"

'从参数表中取药品价格、数量、金额小数位数（显示精度）
Private mintShowCostDigit As Integer            '成本价小数位数
Private mintShowPriceDigit As Integer           '售价小数位数
Private mintShowNumberDigit As Integer          '数量小数位数
Private mintShowMoneyDigit As Integer           '金额小数位数

Private mintMaxMoneyBit As Integer          '药品库存表中金额小数位数
Private mstrMaxMoneyFormat As String

Private mbln零差价模式 As Boolean

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    lng移入库房 As Long
    str填制人 As String
    str审核人 As String
    lng药品分类 As Long
    str剂型 As String
End Type

Private SQLCondition As Type_SQLCondition



Private Function InitComandBars() As Boolean
    '----------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '----------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl

    Dim panThis As Pane
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    
'    Set cbsThis.Icons = zlCommFun.GetPubIcons
    Set cbsThis.Icons = imgPublic.Icons
    
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.id = mconMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_PrintSet, "打印设置(&S)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Preview, "打印预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_BillPrint, "单据打印(&B)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_BillPreview, "单据预览(&L)")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Excel, "输出到Excel"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Parameter, "参数设置(&R)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.id = mconMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddBill, "增加记录单(&B)")
        
        Set mcbrControl = .Add(xtpControlPopup, mconMenu_Edit_AddTable, "增加盘点表(&T)")
        mcbrControl.id = mconMenu_Edit_AddTable
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableAuto, "自动产生盘点表(&A)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableTotal, "汇总记录单产生盘点表(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableZero, "全部盘为零(&Z)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableHouseAll, "库房全部药品盘点(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableSpecial, "特殊药品盘点(&S)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddModify, "修改(&M)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddDel, "删除(&D)")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddVerify, "审核(&C)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddStrike, "冲销(&K)")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddAffirmant, "月度确认(&O)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddDisplay, "查看单据(&W)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_CheckTable, "盘点表智能检查(&T)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "查看(&V)")
    mcbrMenuBar.id = mconMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls

        Set mcbrControl = .Add(xtpControlButton, mconMenu_View_StatusBar, "状态栏(&S)")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_View_ColSet, "列设置(&C)"): mcbrControl.BeginGroup = True
        
    End With
    

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.id = mconMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Help_Help, "帮助主题(&H)")
        
        Set mcbrControl = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB上的中联")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, "中联主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Forum, "中联论坛(&F)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "发送反馈(&M)…", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
        
    End With

    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), mconMenu_File_Print 'Ctrl+P
        .Add FCONTROL, Asc("B"), mconMenu_File_BillPrint
        .Add FCONTROL, Asc("A"), mconMenu_Edit_AddBill
        .Add 0, VK_DELETE, mconMenu_Edit_AddDel
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help
        .Add 0, VK_ESCAPE, mconMenu_File_Exit
    End With

    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand mconMenu_File_PrintSet
        .AddHiddenCommand mconMenu_File_Excel
    End With

    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Print, "打印")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddBill, "记录单"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlPopup, mconMenu_Edit_AddTable, "盘点表")
        mcbrControl.id = mconMenu_Edit_AddTable
        mcbrControl.IconId = mconMenu_Edit_AddBill
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableAuto, "自动产生盘点表(&A)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableTotal, "汇总记录单产生盘点表(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableZero, "全部盘为零(&Z)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableHouseAll, "库房全部药品盘点(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableSpecial, "特殊药品盘点(&S)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddModify, "修改")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddDel, "删除")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddVerify, "审核"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddStrike, "冲销")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddAffirmant, "月度确认"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_CheckTable, "盘点表智能检查"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    InitComandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume

End Function
Private Sub cboStock_Click()
    If mlng库房ID <> Me.cboStock.ItemData(Me.cboStock.ListIndex) Then
        mlng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
        Call GetDrugDigit(mlng库房ID, mstrTitle, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '重新组织格式化串
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        
        mstrMaxMoneyFormat = "'999999999990." & String(mintMaxMoneyBit, "0") & "'"
    End If
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    str工作性质 = "H,I,J,K,L,M,N"

    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), str工作性质, IIf(zlStr.IsHavePrivs(mstrPrivs, "所有库房"), False, True)) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    If cboStock.ListCount > 0 Then
        If cboStock.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub cbo未审核_Click()
    Dim dateCurrentDate As Date
    
    If cbo未审核.ListIndex = cbo未审核.ListCount - 1 Then '选择自定义日期选择框才可用
        dtp开始时间(0).Enabled = True
        dtp结束时间(0).Enabled = True
    Else
        dtp开始时间(0).Enabled = False
        dtp结束时间(0).Enabled = False
    End If
    
    '根据选择改变时间
    dateCurrentDate = Sys.Currentdate
    Select Case cbo未审核.ListIndex
        Case 1
            dtp开始时间(0).Value = Format(DateAdd("d", 0, dateCurrentDate), "yyyy-MM-dd")
            dtp结束时间(0).Value = dateCurrentDate
        Case 2
            dtp开始时间(0).Value = Format(DateAdd("d", -6, dateCurrentDate), "yyyy-MM-dd")
            dtp结束时间(0).Value = dateCurrentDate
        Case 3
            dtp开始时间(0).Value = Format(DateAdd("m", 0, dateCurrentDate), "yyyy-MM")
            dtp结束时间(0).Value = dateCurrentDate
    End Select
    
End Sub

Private Sub cbo已审核_Click()
    Dim dateCurrentDate As Date

    If cbo已审核.ListIndex = cbo已审核.ListCount - 1 And cbo已审核.Enabled Then    '选择自定义日期选择框才可用
        dtp开始时间(1).Enabled = True
        dtp结束时间(1).Enabled = True
    Else
        dtp开始时间(1).Enabled = False
        dtp结束时间(1).Enabled = False
    End If
    chkStrike.Enabled = cbo已审核.ListIndex <> 0 And cbo已审核.Enabled
    
    '根据选择改变时间
    dateCurrentDate = Sys.Currentdate
    Select Case cbo已审核.ListIndex
        Case 1
            dtp开始时间(1).Value = Format(DateAdd("d", 0, dateCurrentDate), "yyyy-MM-dd")
            dtp结束时间(1).Value = dateCurrentDate
        Case 2
            dtp开始时间(1).Value = Format(DateAdd("d", -6, dateCurrentDate), "yyyy-MM-dd")
            dtp结束时间(1).Value = dateCurrentDate
        Case 3
            dtp开始时间(1).Value = Format(DateAdd("m", 0, dateCurrentDate), "yyyy-MM")
            dtp结束时间(1).Value = dateCurrentDate
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        '文件
        Case mconMenu_File_PrintSet
            cbsFilePrintSet '打印设置
        Case mconMenu_File_Preview
            cbsFilePreView '打印预览
        Case mconMenu_File_Print
            cbsFilePrint '打印
        Case mconMenu_File_BillPrint
            cbsFileBillPrint '单据打印
        Case mconMenu_File_BillPreview
            cbsFileBillPreview '单据预览
        Case mconMenu_File_Excel
            cbsFileExcel '输出到&Excel
        Case mconMenu_File_Parameter
            cbsFileParameter '参数设置
        Case mconMenu_File_Exit
            cbsfileExit '退出
        '编辑
        Case mconMenu_Edit_AddBill
            cbsEditaddBill '增加记录单
        Case mconMenu_Edit_AddTableAuto
            cbsAddTableAuto '自动产生盘点表
        Case mconMenu_Edit_AddTableTotal
            cbsAddTableTotal '汇总记录单产生盘点表
        Case mconMenu_Edit_AddTableZero
            cbsAddTableZero '全部盘为零
        Case mconMenu_Edit_AddTableHouseAll
            cbsAddTableHouseAll '库房全部药品盘点
        Case mconMenu_Edit_AddTableSpecial
            cbsAddTableSpecial '特殊药品盘点
        Case mconMenu_Edit_AddModify
            cbsEditModify '修改
        Case mconMenu_Edit_AddDel
            cbsEditDel '删除
        Case mconMenu_Edit_AddVerify
            cbsVerify '审核
        Case mconMenu_Edit_AddStrike
            cbsEditStrike '冲销
        Case mconMenu_Edit_AddAffirmant
            cbsAffirmant '确认
        Case mconMenu_Edit_AddDisplay
            cbsDisplay '查看单据
        Case mconMenu_Edit_CheckTable
            cbsCheckTable '盘点表智能检查
            
        '查看
        Case mconMenu_View_StatusBar
            cbsViewStatus '状态栏
        Case mconMenu_View_Refresh
            cbsViewRefresh '刷新
        Case mconMenu_View_ColSet
            cbsViewColSet '列设置
        
        '帮助
        Case mconMenu_Help_Help
            cbsHelpTitle '帮助主题
        Case mconMenu_Help_Web_Home
            cbsHelpWebHome '中联主页
        Case mconMenu_Help_Web_Forum
            cbsHelpWebForum '中联论坛
        Case mconMenu_Help_Web_Mail
            cbsHelpWebMail '发送反馈
        Case mconMenu_Help_About
            cbsHelpAbout '关于
        Case Else
            If Control.id > 401 And Control.id < 499 Then
                '执行自定义报表
                Call BillPrint_Custom(Control)
            End If
    End Select
    
End Sub

Private Sub cbsEditaddBill()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmNewCheckCourseCard.ShowCard Me, strNo, 1, , blnSuccess
    
    If blnSuccess Then Call cbsViewRefresh
End Sub

Private Sub cbsViewStatus()
    Dim cbrMenuPop As CommandBarControl
    
    Set cbrMenuPop = Me.cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_StatusBar, , True)
    
    With cbrMenuPop
        .Checked = Not .Checked  ' Xor True
        stbThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '打印自定义报表
    '默认参数：药品=药品id，库房=库房id，开始时间=填制开始时间，结束时间=填制结束时间，盘点单=盘点单NO，盘点表=盘点表NO
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim strNo As String
    Dim strReportName As String

    strReportName = Split(Control.Parameter, ",")(1)

    Select Case strReportName
        Case "ZL1_INSIDE_1307"
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1307", Me, "库房=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)))
        Case "ZL1_INSIDE_1307_1"
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1307_1", Me, "库房=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)), "单位=" & Choose(mintUnit, "售价单位", "门诊单位", "住院单位", "药库单位") & "|" & Choose(mintUnit, 1, 3, 4, 2))
        Case Else
            If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
                strNo = vsfList.TextMatrix(vsfList.Row, 0)
            End If

            str开始时间 = IIf(Format(SQLCondition.date填制时间开始, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间开始, "yyyy-mm-dd"))
            str结束时间 = IIf(Format(SQLCondition.date填制时间结束, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date填制时间结束, "yyyy-mm-dd"))

            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strReportName, Me, _
                "药品=" & IIf(SQLCondition.lng药品 = 0, "", SQLCondition.lng药品), _
                "库房=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
                "开始时间=" & str开始时间, _
                "结束时间=" & str结束时间, _
                "盘点单=" & strNo, _
                "盘点表=" & strNo)
    End Select
End Sub



Private Sub cbsViewRefresh()
    '刷新
    GetList mstrFind
End Sub


Private Sub cbsViewColSet()
    Dim strColsName As String '可以屏蔽的列
    Dim strDefaultColsName As String '默认的列
    Dim i As Integer
    Dim strColName As String
    '列设置
    strDefaultColsName = ":药品来源,0:基本药物,0:库房货位,0:批准文号,0:金额差,0:差价差,0:盘点成本金额差,0:账面金额差,0:成本金额差,0:当前库存,1:" '所有可以隐藏的列
    
    
    strColsName = zlDataBase.GetPara("列设置", glngSys, mlngMode, "") '获取数据库的保存信息
    
    '兼容处理
    If strColsName = "" Then '未提取到列设置信息
        strColsName = strDefaultColsName
    Else
        '判断提取的列与默认列个数，不一致则取默认的
        If UBound(Split(strColsName, ":")) <> UBound(Split(strDefaultColsName, ":")) Then strColsName = strDefaultColsName
        
        '判断提取的列名是否与默认的一致，不一致取默认的
        For i = LBound(Split(strColsName, ":")) + 1 To UBound(Split(strColsName, ":")) - 1
            strColName = Split(Split(strColsName, ":")(i), ",")(0) '获取单个列名
            
            If InStr(1, strDefaultColsName, ":" & strColName) = 0 Then '列名不存在于默认列名中
                strColsName = strDefaultColsName
                Exit For
            End If
        Next
        
    End If
    
    strColsName = frm隐藏列设置.ShowME(Me, strColsName)
    
    If strColsName <> "" Then
        zlDataBase.SetPara "列设置", strColsName, glngSys, mlngMode
    End If
    
    cbsViewRefresh

End Sub

Private Sub cbsHelpAbout()
    '关于
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub cbsHelpTitle()
    '帮助主题
    Dim StrWinName As String
    With vsfList
        StrWinName = "frmMainList8"
    End With
    Call ShowHelp(App.ProductName, Me.hWnd, StrWinName)
End Sub

Private Sub cbsHelpWebHome()
    '中联主页
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub cbsHelpWebForum()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub cbsHelpWebMail()
    '发送反馈
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub cbsFilePreView()
    '打印预览
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub cbsFilePrint()
    '打印
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub cbsFilePrintSet()
    '打印设置
    zlPrintSet
End Sub

Private Sub cbsFileParameter()
    '参数设置
    Dim int查询天数  As Integer
    
    frm参数设置.设置参数 Me, mstrPrivs, mstrTitle
    
    '界面需要变动
    int查询天数 = Val(zlDataBase.GetPara("查询天数", glngSys, mlngMode, 7))
    int查询天数 = IIf(int查询天数 <> 1 And int查询天数 <> 7, 7, int查询天数)
    
    cbo未审核.ListIndex = IIf(int查询天数 = 7, 2, 1)
    
    cmd确认_Click '确认刷新界面
    
End Sub

Private Sub cbsfileExit()
    '退出
    Unload Me
End Sub

Private Sub cbsFileExcel()
    '输出到Excel
    
    If Me.ActiveControl Is vsfList Then
        vsfList.Redraw = flexRDNone
        subPrint 3
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    ElseIf Me.ActiveControl Is vsfDetail Then
        vsfDetail.Redraw = flexRDNone
        subExcel 3
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub cbsFileBillPrint()
    Dim int单位系数 As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case mconint售价单位
                int单位系数 = 4
            Case mconint门诊单位
                int单位系数 = 2
            Case mconint住院单位
                int单位系数 = 1
            Case mconint药库单位
                int单位系数 = 3
        End Select
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 2
    End With
End Sub

Private Sub cbsFileBillPreview()
    Dim int单位系数 As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case mconint售价单位
                int单位系数 = 4
            Case mconint门诊单位
                int单位系数 = 2
            Case mconint住院单位
                int单位系数 = 1
            Case mconint药库单位
                int单位系数 = 3
        End Select
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), Me, "单据编号=" & .TextMatrix(.Row, 0), "记录状态=" & .TextMatrix(.Row, .Cols - 2), "单位系数=" & int单位系数, 1
    End With
End Sub


Private Sub cbsEditModify()
    '修改
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        If tbcDetail.Selected.Index = 0 Then
            frmNewCheckCourseCard.ShowCard Me, strNo, 2, 1, blnSuccess
        Else
            frmNewCheckCard.ShowCard Me, strNo, 2, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
        End If
        
        If blnSuccess Then Call cbsViewRefresh
    End With
End Sub

Private Sub cbsEditStrike()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    
    '如果是外购(blnPurchase为真)，则直接进入冲销
    '询问是否冲销(blnPurchase为提示框返回值)，是则进入冲销
    blnPurchase = (InStr(1, "1300,1302,1304,1305,1306", mlngMode) <> 0)
    With vsfList
        If Not blnPurchase Then
            blnPurchase = (MsgBox("你确实要全部冲销单据号为“" & .TextMatrix(.Row, 0) & "”的单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then cbsViewRefresh
        End If
    End With
End Sub


Private Function StrikeSave() As Boolean
    Dim blnSuccess As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim int库存检查 As Integer
    Dim strMsg As String
    Dim n As Integer
    
    StrikeSave = False
    
    int库存检查 = MediWork_GetCheckStockRule(mlng库房ID)
    
    On Error GoTo ErrHandle
    If int库存检查 <> 0 Then
        gstrSQL = "Select A.药品信息 " & _
            " From (Select Distinct '(' || I.编码 || ')' || Nvl(N.名称, I.名称) As 药品信息, A.实际数量, Nvl(K.实际数量, 0) As 库存数量 " & _
            " From 药品收发记录 A, (Select 药品id, 库房id, 实际数量, Nvl(批次, 0) 批次 From 药品库存 Where 性质 = 1) K, 药品规格 B, 收费项目目录 I, 收费项目别名 N " & _
            " Where A.药品id = K.药品id(+) And A.库房id = K.库房id(+) And Nvl(A.批次, 0) = K.批次(+) And A.药品id = B.药品id And " & _
            " A.药品id = I.ID And A.药品id = N.收费细目id(+) And N.性质(+) = 3 And A.单据 = 12 And A.入出系数 = 1 And A.NO = [1]) A " & _
            " Where A.实际数量 > A.库存数量 "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "检查库存", vsfList.TextMatrix(vsfList.Row, 0))
        
        With rsTemp
            If .RecordCount > 0 Then
                For n = 1 To .RecordCount
                    If n > 5 Then
                        strMsg = strMsg & vbCrLf & "还有其他" & .RecordCount - 5 & "个药品......"
                        Exit For
                    End If
                    strMsg = IIf(strMsg = "", "", strMsg & "," & vbCrLf) & !药品信息
                    .MoveNext
                Next
                
                If int库存检查 = 1 Then
                    If MsgBox("注意，以下药品库存不足：" & vbCrLf & strMsg & vbCrLf & Space(4) & "是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                ElseIf int库存检查 = 2 Then
                    MsgBox "对不起，以下药品库存不足，不能冲销！" & vbCrLf & strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End With
    End If
    
    With vsfList
        gstrSQL = "zl_药品盘点_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.用户姓名 & "')"
        
        Call zlDataBase.ExecuteProcedure(gstrSQL, mstrTitle)
        
        '提示停用药品
        Call CheckStopMedi(单据号.盘点表 & "|" & .TextMatrix(.Row, 0))
    End With
    StrikeSave = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    
    'MsgBox "存盘失败！", vbInformation, gstrSysName
    Call SaveErrLog

End Function

Private Sub cbsEditDel()
    '删除
    Dim strBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With vsfList
        strTitle = IIf(tbcDetail.Selected.Index = 0, "盘点记录单", "盘点表")
        
        On Error GoTo ErrHandle
        intRow = .Row
        strBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("你确实要删除单据号为“" & strBillNo & "”的" & strTitle & "吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            If tbcDetail.Selected.Index = 1 Then
                gstrSQL = "zl_药品盘点_Delete('" & strBillNo & "')"
            Else
                gstrSQL = "zl_药品盘点记录单_Delete('" & strBillNo & "')"
            End If
            Call zlDataBase.ExecuteProcedure(gstrSQL, mstrTitle)
            
            intRecord = intRecord - 1
            mlastRow = 0
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                With vsfDetail
                    .rows = 1
                    .rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                
            End If
                
            '.RowHeight(intRow) = 0
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
            vsfList_EnterCell
        End If
    End With
    stbThis.Panels(2).Text = "当前共有" & intRecord & "张单据"
    cbsViewRefresh
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then Resume 'Resume这种情况不用调用
    Call SaveErrLog
End Sub

Private Sub cbsAffirmant()
    Dim str审核日期 As String       '缺省做为确认记录的结束日期
    '填写月度确认记录
    If tbcDetail.Selected.Index = 1 Then
        str审核日期 = vsfList.TextMatrix(vsfList.Row, 5)
    End If
    With frm月度确认
        Call .ShowEditor(Me.cboStock.ItemData(Me.cboStock.ListIndex), str审核日期)
    End With
End Sub

Private Sub cbsDisplay()
    '查看单据
    
    Dim strNo As String
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        If tbcDetail.Selected.Index = 0 Then
            frmNewCheckCourseCard.ShowCard Me, strNo, 4
        Else
            frmNewCheckCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .Cols - 2)
        End If
    End With
End Sub

Private Sub cbsCheckTable()
    Dim blnSuccess As Boolean
    '智能检查
    frmSmartCheck.ShowME cboStock.ItemData(cboStock.ListIndex), Me, blnSuccess
    
    If blnSuccess Then cbsViewRefresh
End Sub

Private Sub cbsAddTableAuto()
    Dim strNo As String
    Dim blnSuccess As Boolean

    frmNewCheckCard.ShowCard Me, strNo, 1, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsAddTableTotal()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmNewCheckCard.ShowCard Me, strNo, 5, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsAddTableZero()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmNewCheckCard.ShowCard Me, strNo, 6, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsAddTableHouseAll()
    Dim strNo As String
    Dim blnSuccess As Boolean

    frmNewCheckCard.ShowCard Me, strNo, 7, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsAddTableSpecial()
    Dim strNo As String
    Dim blnSuccess As Boolean

    frmNewCheckCard.ShowCard Me, strNo, 8, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsVerify()
    '验收
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmNewCheckCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, .Cols - 2), blnSuccess
    End With
    
    If blnSuccess Then Call cbsViewRefresh
End Sub



Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not mblnLoad Then Exit Sub
    If Not mblnBootUp Then Exit Sub

    '设置控件的相关属性
    Call 权限控制(Control)
    
    '设置菜单和工具按钮的可用属性

    Dim strVerify As String, blnVisible As Boolean
    
    blnVisible = (tbcDetail.Selected.Index = 1)
    If tbcDetail.Selected.Index = 1 Then
        strVerify = vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 8)
    Else
        strVerify = ""
    End If
    
    With vsfList
        .ToolTipText = ""
    
        Select Case Control.id
            Case mconMenu_File_Preview    '预览
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            Case mconMenu_File_Print   '打印
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            Case mconMenu_File_BillPreview    '单据预览
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    Control.Enabled = tbcDetail.Selected.Index = 1
                End If
            Case mconMenu_File_BillPrint    '单据打印
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    Control.Enabled = tbcDetail.Selected.Index = 1
                End If
            Case mconMenu_File_Excel    '输出到Excel
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            Case mconMenu_Edit_AddModify    '修改
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    '未审核单
                        Control.Enabled = True
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '审核单
                        Control.Enabled = False
                    Else '2,3 冲销单
                        Control.Enabled = False
                    End If
                End If
            Case mconMenu_Edit_AddDel    '删除
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    '未审核单
                        Control.Enabled = True
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '审核单
                        Control.Enabled = False
                    Else '2,3 冲销单
                        Control.Enabled = False
                    End If
                End If
            Case mconMenu_Edit_AddVerify   '审核
                Control.Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "审核")
                
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    '未审核单
                        Control.Enabled = True
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '审核单
                        Control.Enabled = False
                    Else '2,3 冲销单
                        Control.Enabled = False
                    End If
                End If
            Case mconMenu_Edit_AddStrike   '冲销
                Control.Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "冲销")
                If Not zlStr.IsHavePrivs(mstrPrivs, "审核") And Control.Visible Then Control.BeginGroup = True
                
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    '未审核单
                        Control.Enabled = False
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '审核单
                        Control.Enabled = True
                    Else '2,3 冲销单
                        If .TextMatrix(.Row, .Cols - 2) Mod 3 = 0 Then
                            .ToolTipText = "冲销单据的原单据"
                            Control.Enabled = True
                        ElseIf .TextMatrix(.Row, .Cols - 2) Mod 3 = 2 Then
                            .ToolTipText = "冲销单据"
                            Control.Enabled = False
                        End If
                    End If
                End If
            Case mconMenu_Edit_AddDisplay    '查看单据
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         '没有单
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    '未审核单
                        Control.Enabled = True
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '审核单
                        Control.Enabled = True
                    Else '2,3 冲销单
                        Control.Enabled = True
                    End If
                End If
                
        End Select
    End With
    
    
End Sub

Private Sub Chk药品_Click()
    Txt药品.Enabled = IIf(Chk药品.Value = 1, True, False)
    Cmd药品.Enabled = IIf(Chk药品.Value = 1, True, False)
End Sub

Private Sub Cmd查阅_Click()
    Call cbsDisplay
End Sub

Private Sub cmd确认_Click()
    Dim strFind As String
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    
    If cbo已审核.Enabled Then
        If cbo未审核.ListIndex = 0 And cbo已审核.ListIndex = 0 Then
            MsgBox "对不起，必须选择一种单据显示（默认显示当天未审核单据）!", vbInformation, gstrSysName
            cbo未审核.ListIndex = 1
            cbo未审核.SetFocus
            Exit Sub
        ElseIf cbo未审核.ListIndex <> 0 And cbo已审核.ListIndex = 0 Then '只有未审核单据
            strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
        ElseIf cbo未审核.ListIndex = 0 And cbo已审核.ListIndex <> 0 Then '只有已审核单据
            If chkStrike.Value = 1 Then '包含冲销
                strFind = " AND  A.审核日期 is not Null And A.审核日期 Between [5] And [6] "
            Else
                strFind = " AND A.记录状态 = 1 And A.审核日期 is not Null And A.审核日期 Between [5] And [6] "
            End If
        Else '包括审核和未审核单据
            If chkStrike.Value = 1 Then  '包含冲销
                strFind = " AND (( A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4]) or ( A.审核日期 is not Null And A.审核日期 Between [5] And [6])) "
            Else
                strFind = " AND (( A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4]) or (A.记录状态 = 1 And A.审核日期 is not Null And A.审核日期 Between [5] And [6])) "
            End If
        End If
    Else
        If cbo未审核.ListIndex = 0 Then
            MsgBox "对不起，必须选择一种单据显示（默认显示当天未审核单据）!", vbInformation, gstrSysName
            cbo未审核.ListIndex = 1
            cbo未审核.SetFocus
            Exit Sub
        ElseIf cbo未审核.ListIndex <> 0 Then '只有未审核单据
            strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
        End If
    End If
    
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, intNO, lng库房ID)
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, intNO, lng库房ID)
    End If

    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then strFind = strFind & " And A.No >= [1] And A.No <=[2] "
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then strFind = strFind & " And A.No >= [1] "
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then strFind = strFind & " And A.No <= [2] "

    SQLCondition.strNO开始 = Me.txt开始No
    SQLCondition.strNO结束 = Me.txt结束NO
    
    SQLCondition.date填制时间开始 = CDate(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date审核时间开始 = CDate(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date审核时间结束 = CDate(Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59")
    
    If Chk药品.Value = 1 Then
        strFind = strFind & " And A.药品ID + 0 =[7] "
    End If
    
    SQLCondition.lng药品 = Val(Txt药品.Tag)
    
    If Me.Txt审核人 <> "" And Txt审核人.Enabled Then strFind = strFind & " And A.审核人 like [10] "
    If Me.Txt填制人 <> "" Then strFind = strFind & " And A.填制人 like [9] "
    
    SQLCondition.str审核人 = Me.Txt审核人 & "%"
    SQLCondition.str填制人 = Me.Txt填制人 & "%"
    
    mstrFind = strFind
    
    GetList (mstrFind)  '列出单据头
End Sub

Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "药品移库管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , True)

    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品id
    
End Sub

Private Sub cmd重置_Click()
    cboStock.ListIndex = 0
    cbo未审核.ListIndex = 1
    cbo已审核.ListIndex = 0
    txt开始No.Text = ""
    txt结束NO.Text = ""
    Txt填制人.Text = ""
    Txt审核人.Text = ""
    Chk药品.Value = 0
    Txt药品.Text = ""
    chkStrike.Value = 0
End Sub

Private Sub Form_Activate()
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = 1
        vsfDetail.Row = 1
    End If
End Sub

Private Sub Form_Load()
    
    mblnLoad = False
    
    mintMaxMoneyBit = gtype_UserDrugDigits.Digit_金额
    mbln零差价模式 = gtype_UserSysParms.P275_零差价管理模式 <> 0
    
    InitComandBars
    InitTabControl
    loadCbo
    
    Me.dtp结束时间(1) = Sys.Currentdate
    Me.dtp开始时间(1) = DateAdd("d", -7, Me.dtp结束时间(1))
    
    Me.Caption = mstrTitle
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    
    Call zlDataBase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    stbThis.Panels(2).Picture = picColor
    
    Dim cbrMenuPop As CommandBarControl
    
    
    mblnLoad = True
End Sub

Private Sub loadCbo()
    '初始化下拉框
    Dim int查询天数  As Integer
    
    int查询天数 = Val(zlDataBase.GetPara("查询天数", glngSys, mlngMode, 7))
    int查询天数 = IIf(int查询天数 <> 1 And int查询天数 <> 7, 7, int查询天数)
    
    cbo未审核.AddItem "0-不显示"
    cbo未审核.AddItem "1-显示今日"
    cbo未审核.AddItem "2-显示7天之内"
    cbo未审核.AddItem "3-显示本月"
    cbo未审核.AddItem "4-自定义"
    cbo未审核.ListIndex = IIf(int查询天数 = 7, 2, 1)
    
    cbo已审核.AddItem "0-不显示"
    cbo已审核.AddItem "1-显示今日"
    cbo已审核.AddItem "2-显示7天之内"
    cbo已审核.AddItem "3-显示本月"
    cbo已审核.AddItem "4-自定义"
    cbo已审核.ListIndex = 0
    
End Sub

Private Sub InitTabControl()
    '初始化分页控件
    
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mconTab_CheckCourseCard, "盘点记录单清单(&1)", Me.picMain.hWnd, 0).Tag = "盘点记录单清单(&1)_"
        .InsertItem(mconTab_CheckCard, "盘点表清单(&2)", Me.picMain.hWnd, 0).Tag = "盘点表清单(&2)_"
        
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 12900 Then Me.Width = 12900
    If Me.Height < 8000 Then Me.Height = 8000
    If picMain.Height < 4000 - picMain.Top Then picMain.Height = 4000 - picMain.Top
    
    
    fraCondition.Move 0, 900, Me.ScaleWidth, 1300
    cmd确认.Left = fraCondition.Width - cmd确认.Width - 100
    cmd确认.Top = dtp结束时间(1).Top - (cmd确认.Height - dtp结束时间(1).Height)
    cmd重置.Left = cmd确认.Left - cmd重置.Width - 50
    cmd重置.Top = cmd确认.Top
    
    With tbcDetail
        .Top = fraCondition.Top + fraCondition.Height + 20
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - fraCondition.Top - fraCondition.Height - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    '状态栏是否勾选
    Me.cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_StatusBar, , True).Checked = stbThis.Visible
    
    picMain.Move 0, 360, tbcDetail.Width, tbcDetail.Height - stbThis.Height
    
'    vsfList.Move 0, 0, picMain.Width, (picMain.Height - picSeparate_s.Height) / 2
    vsfList.Move 0, 0, picMain.Width
    
    With picSeparate_s
        .Left = 0
        .Top = vsfList.Top + vsfList.Height
        .Width = picMain.Width
    End With
    
    
    
'    vsfDetail.Move 0, picSeparate_s.Top + picSeparate_s.Height + 100, picMain.Width, (picMain.Height - picSeparate_s.Height) / 2 - 110
    If picSeparate_s.Top > picMain.Height - 2000 Then
        vsfList.Move 0, 0, picMain.Width, picMain.Height - (2100 + picSeparate_s.Height)
        picSeparate_s.Top = vsfList.Top + vsfList.Height
    End If
    
    With Cmd查阅
        .Left = picMain.Width - .Width - 100
        .Top = vsfList.Top + vsfList.Height + 30
    End With
    
    With vsfDetail
        .Left = 0
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Width = picMain.Width
        .Height = picMain.Height - .Top
    End With
    
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 300
    End With
    
End Sub


'******************************************************************************************************************
'功能：
'参数：
'返回：
'******************************************************************************************************************
Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
On Error GoTo errH
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    DockPannelInit = True
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "ZLSOFT"
    Err.Clear
End Function



Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    Call SaveFlexState(vsfList, tbcDetail.Selected.Caption)
    Call SaveFlexState(vsfDetail, tbcDetail.Selected.Caption)
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        mshSelect.Visible = False
        Exit Sub
    End If
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Booker"
                    Txt填制人 = .TextMatrix(.Row, 2)
                    If tbcDetail.Selected.Index = mconTab_CheckCard Then
                        Txt审核人.SetFocus
                    Else
                        cbo未审核.SetFocus
                    End If
                Case "Verify"
                    Txt审核人 = .TextMatrix(.Row, 2)
                    cbo未审核.SetFocus
            End Select
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub



Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    If vsfList.Height + y <= 1500 Then Exit Sub
    If vsfDetail.Height - y <= 1500 Then Exit Sub

    picSeparate_s.Move 0, picSeparate_s.Top + y
    Cmd查阅.Move Me.ScaleWidth - Cmd查阅.Width - 500, picSeparate_s.Top + 50
    vsfList.Move 0, 0, Me.ScaleWidth, vsfList.Height + y
    vsfDetail.Move 0, vsfList.Height + Cmd查阅.Height + 100, Me.ScaleWidth, vsfDetail.Height - y

    
'    With picSeparate_s
'        If .Top + picMain.Top + y < 2000 Then Exit Sub
'        If .Top + y > picMain.Height - 2000 Then Exit Sub
'        .Move .Left, .Top + y
'    End With
'
'    With vsfList
'        .Height = picSeparate_s.Top - .Top
'    End With
'
'    With Cmd查阅
'        .Top = vsfList.Top + vsfList.Height + 30
'    End With
'
'    With vsfDetail
'        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
'        .Height = picMain.Height - .Top
'    End With
End Sub


Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal FrmMain As Variant)
    Dim strFind As String
    Dim dateCurrentDate As Date
    Dim strTemp As String
    Dim int查询天数 As Integer
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrprivs
    Me.Caption = strTitle
    
    If Not CheckDepend Then Exit Sub            '数据依赖性测试
    
    mlng库房ID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng库房ID, mstrTitle, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        
    mstrVerifyStart = "1901-01-01"
    mstrVerifyEnd = "1901-01-01"
    
    dateCurrentDate = Sys.Currentdate
    mstrStart = Format(DateAdd("d", -6, dateCurrentDate), "yyyy-MM-dd") '默认提取7天数据
    mstrEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strFind = " AND A.记录状态 = 1 And A.审核日期 is Null And A.填制日期 Between [3] And [4] "
    SQLCondition.date填制时间开始 = CDate(Format(mstrStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(mstrEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    GetList (mstrFind)  '列出单据头
    
    RestoreWinState Me, App.ProductName, mstrTitle
        
    mblnBootUp = True
    
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        OS.ShowChildWindow Me.hWnd, FrmMain
    End If
    Me.ZOrder 0
End Sub

'检查数据依赖性
Private Function CheckDepend() As Boolean
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    CheckDepend = False
    
    strStock = "HIJKLMN"
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
             & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 is Null) And c.工作性质 = b.名称 " _
              & "AND Instr([1],b.编码,1) > 0 " _
             & " AND a.id = c.部门id " _
              & "AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'"

    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, strStock) '查看是否有药库性质，药房性质，或者制剂室性质的部门
    
    If rsDepend.EOF Then
        MsgBox "至少应该设置一个具有药库性质，药房性质，或者制剂室性质的部门,请查看部门管理！", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
             & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 is Null) And c.工作性质 = b.名称 " _
              & "AND Instr([1],b.编码,1) > 0 " _
             & " AND a.id = c.部门id " _
              & "AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
              & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有库房"), "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[2])")

    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, strStock, UserInfo.用户ID) '查看用户所属的有药库性质，药房性质，或者制剂室性质的部门
            
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0  '缺省部门不是药库性质，药房性质，或者制剂室性质的部门则默认选择第一个以上性质的部门
        
        If .ListIndex = -1 Then
            If Not zlStr.IsHavePrivs(mstrPrivs, "所有库房") Then
                MsgBox "你不是库房工作人员或不具有所有库房的权限，不能进入！", vbInformation, gstrSysName
                Unload Me
                Exit Function
            End If
            .ListIndex = 0
        End If
    End With

    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'根据权限设置不同的显示项目
Private Sub 权限控制(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '权限控制
  
    Select Case Control.id
        Case mconMenu_Edit_AddBill, mconMenu_Edit_AddTable '记录单、记录表
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "登记")
            If mconMenu_Edit_AddTable = Control.id Then
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "药品盘点表") And zlStr.IsHavePrivs(mstrPrivs, "登记")
                tbcDetail.Item(mconTab_CheckCard).Visible = zlStr.IsHavePrivs(mstrPrivs, "药品盘点表")
            End If
        Case mconMenu_Edit_AddModify '修改
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "修改")
            '判断是否加分界线
            If Not zlStr.IsHavePrivs(mstrPrivs, "登记") And zlStr.IsHavePrivs(mstrPrivs, "修改") Then Control.BeginGroup = True
        Case mconMenu_Edit_AddDel  '删除
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "删除")
            '判断是否加分界线
            If Not zlStr.IsHavePrivs(mstrPrivs, "登记") And Not zlStr.IsHavePrivs(mstrPrivs, "修改") And zlStr.IsHavePrivs(mstrPrivs, "删除") Then Control.BeginGroup = True
        Case mconMenu_Edit_AddVerify  '审核
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "审核")
        Case mconMenu_Edit_AddStrike   '冲销
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "冲销")
            '判断是否加分界线
            If Not zlStr.IsHavePrivs(mstrPrivs, "审核") And zlStr.IsHavePrivs(mstrPrivs, "冲销") Then Control.BeginGroup = True
        Case mconMenu_Edit_AddAffirmant  '确认
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "月度确认")
        Case mconMenu_Edit_AddTableZero   '全部盘为零
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "全部盘为零")
        Case mconMenu_File_BillPrint    '单据打印
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "单据打印")
    End Select
End Sub

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    Dim str包装系数 As String
    Dim strSqlForm As String
    Dim n As Integer
    
    '用于统计合计金额
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim dbl盘点成本金额 As Double
    Dim dbl盘点金额差 As Double

    mlastRow = 0
    On Error GoTo ErrHandle

    Call FS.ShowFlash("正在搜索药品记录,请稍候 ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    strUserPart = " And A.库房ID+0=[11] "
    
    Select Case mintUnit
        Case mconint售价单位
            str包装系数 = "1"
        Case mconint门诊单位
            str包装系数 = "B.门诊包装"
        Case mconint住院单位
            str包装系数 = "B.住院包装"
        Case mconint药库单位
            str包装系数 = "B.药库包装"
    End Select
    
    vsfList.Redraw = flexRDNone
    '频次字段保存的 盘点时间
    If tbcDetail.Selected.Index = 1 Then '选择的是盘点表清单
        If SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 = 0 Then
            strSqlForm = " , 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " And b.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 = "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 诊疗项目目录 G"
            strFind = strFind & " And b.药名id = g.Id And g.分类id + 0=[12] and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " And b.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7') and g.分类id + 0=[12]"
        End If
        
        gstrSQL = "Select NO, 盘点时间, 填制人, 填制日期, 修改人, 修改日期, 审核人, 审核日期, " & _
                "   to_char(Sum(盘点金额), " & mstrMoneyFormat & ") 盘点金额, to_char(Sum(金额差), " & mstrMoneyFormat & ") 金额差,to_char(Sum(账面金额差), " & mstrMoneyFormat & ") 账面金额差,to_char(Sum(盘点成本金额)," & mstrMoneyFormat & ") 盘点成本金额, to_char(Sum(成本金额差)," & mstrMoneyFormat & ") 成本金额差, 记录状态, 摘要" & _
                " from ( SELECT a.no,a.序号, 频次 AS 盘点时间," _
                & "a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期,a.修改人,TO_CHAR (min(a.修改日期), 'yyyy-mm-dd HH24:Mi:SS') AS 修改日期, a.审核人," _
                & "TO_CHAR (min(a.审核日期), 'yyyy-mm-dd HH24:Mi:SS') AS 审核日期, " _
                & "     LTrim(To_Char(to_char(A.扣率 /" & str包装系数 & "," & mstrNumberFormat & ") * TO_CHAR (a.零售价*" & str包装系数 & ", " & mstrPriceFormat & ") , " & mstrMoneyFormat & ")) As 盘点金额," _
                & "ltrim(to_char(零售金额*a.入出系数,decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) 金额差," _
                & "ltrim(to_char(to_char((A.扣率-A.填写数量) /" & str包装系数 & "," & mstrNumberFormat & ") * TO_CHAR (a.零售价* Decode(记录状态, 1, 1, Decode(Mod(记录状态, 3), 0, 1, -1))*" & str包装系数 & ", " & mstrPriceFormat & "),decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," _
                & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) AS 账面金额差," _
                & "ltrim(to_char((a.成本价+to_char(a.零售金额*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1)),decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat _
                & ") ", mstrMoneyFormat) & ")))-(a.成本金额+to_char(a.差价*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1)),decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & ")))," & mstrMoneyFormat & ")) as 盘点成本金额," _
                & "ltrim(to_char(a.零售金额*a.入出系数-a.差价*a.入出系数,decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) as 成本金额差," _
                & " a.记录状态, a.摘要 " _
                & " FROM 药品收发记录 a,药品规格 B " & strSqlForm _
                & " Where A.药品ID=B.药品ID And A.单据 = 12  " & strUserPart & strFind _
                & " Group By a.No,a.序号, 频次, a.填制人, a.修改人, a.审核人, a.成本价, a.入出系数, a.成本价,a.成本金额," & str包装系数 & ", a.零售金额, a.记录状态, a.扣率, a.填写数量, a.零售价,a.扣率, a.单量, a.差价, a.摘要,b.是否零差价管理) " _
                & " Group By NO, 盘点时间, 填制人, 填制日期, 修改人, 修改日期, 审核人, 审核日期, 记录状态, 摘要 ORDER BY no DESC,填制日期 ASC"
    Else '选择的是盘点记录单清单
        If SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 = 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 = "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.分类id + 0=[12] and (g.类别='5' or g.类别='6' or g.类别='7')"
        ElseIf SQLCondition.str剂型 <> "" And SQLCondition.lng药品分类 <> 0 Then
            strSqlForm = " , 药品规格 F, 诊疗项目目录 G, 药品特性 H"
            strFind = strFind & " and a.药品id = f.药品id And f.药名id = g.Id And g.Id = h.药名id(+) and h.药品剂型 in(select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.类别='5' or g.类别='6' or g.类别='7') and g.分类id + 0=[12]"
        End If
        gstrSQL = " SELECT a.no, 频次 AS 盘点时间," _
                    & "a.填制人,TO_CHAR (min(a.填制日期), 'yyyy-mm-dd HH24:Mi:SS') AS 填制日期,a.修改人,TO_CHAR (min(a.修改日期), 'yyyy-mm-dd HH24:Mi:SS') AS 修改日期,a.摘要 " _
                    & " FROM 药品收发记录 a " & strSqlForm _
                    & " Where a.单据 = 14  " & strUserPart & strFind _
                    & " Group by a.no,频次,a.填制人,a.修改人,a.摘要 " _
                    & " ORDER BY no DESC,填制日期 ASC "
    End If
    
    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, _
        SQLCondition.strNO开始, _
        SQLCondition.strNO结束, _
        SQLCondition.date填制时间开始, _
        SQLCondition.date填制时间结束, _
        SQLCondition.date审核时间开始, _
        SQLCondition.date审核时间结束, _
        SQLCondition.lng药品, _
        SQLCondition.lng移入库房, _
        SQLCondition.str填制人, _
        SQLCondition.str审核人, _
        cboStock.ItemData(cboStock.ListIndex), _
        SQLCondition.lng药品分类, _
        SQLCondition.str剂型)
    
    mbln绑定 = False
    Set vsfList.DataSource = rsList
    mbln绑定 = True
    
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = flexRDDirect
            
            .TopRow = 1
            .rows = .rows - 99
        End If
    
        
        .ColAlignment(.ColIndex("盘点成本金额")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("盘点金额")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("金额差")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("账面金额差")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("成本金额差")) = flexAlignRightCenter
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        If tbcDetail.Selected.Index = 1 Then '选择的是盘点表清单
            .colHidden(.Cols - 2) = True '始终隐藏"记录状态"这一列
            .colHidden(.ColIndex("金额差")) = True '默认不显示
            .colHidden(.ColIndex("账面金额差")) = True '默认不显示
            .colHidden(.ColIndex("成本金额差")) = True '默认不显示
            
            vsfHidden vsfList
            
            lbl2.Visible = Not .colHidden(.ColIndex("金额差")) '金额差不显示，则金额差合计不显示
            lbl3.Visible = Not .colHidden(.ColIndex("账面金额差")) '账面金额差不显示，则账面金额差合计不显示
            lbl成本金额差.Visible = Not .colHidden(.ColIndex("成本金额差")) '成本金额差不显示，则成本金额差合计不显示
        End If
        
    
        
        For n = 0 To .Cols - 1
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    
    '统计合计金额
    lbl1.Caption = "盘点金额合计："
    lbl2.Caption = "金额差合计："
    lbl3.Caption = "账面金额差合计："
    
    If tbcDetail.Selected.Index = 1 Then '选择的是盘点表清单
        lbl1.Visible = True
        If mblnViewCost = False Then
            lblSum成本金额.Visible = False
            lbl成本金额差.Visible = False
        Else
            lblSum成本金额.Visible = True
            lbl成本金额差.Visible = Not vsfList.colHidden(vsfList.ColIndex("成本金额差")) '成本金额差不显示，则成本金额差合计不显示
        End If
        If (Not rsList.EOF) And (Not rsList.BOF) Then
            rsList.MoveFirst
            Do While Not rsList.EOF
                dbl1 = dbl1 + IIf(IsNull(rsList!盘点金额), 0, rsList!盘点金额)
                dbl2 = dbl2 + IIf(IsNull(rsList!金额差), 0, rsList!金额差)
                dbl3 = dbl3 + IIf(IsNull(rsList!账面金额差), 0, rsList!账面金额差)
                dbl盘点成本金额 = dbl盘点成本金额 + IIf(IsNull(rsList!盘点成本金额), 0, rsList!盘点成本金额)
                dbl盘点金额差 = dbl盘点金额差 + IIf(IsNull(rsList!成本金额差), 0, rsList!成本金额差)
                rsList.MoveNext
            Loop
            rsList.MoveFirst
            
            lbl1.Caption = "盘点金额合计：" & Format(dbl1, "0." & String(mintShowMoneyDigit, "0"))
            lbl2.Caption = "金额差合计：" & Format(dbl2, "0." & String(mintShowMoneyDigit, "0"))
            lbl3.Caption = "账面金额差合计：" & Format(dbl3, "0." & String(mintShowMoneyDigit, "0"))
            lblSum成本金额.Caption = "盘点成本金额合计：" & Format(dbl盘点成本金额, "0." & String(mintShowMoneyDigit, "0"))
            lbl成本金额差.Caption = "成本金额差：" & Format(dbl盘点金额差, "0." & String(mintShowMoneyDigit, "0"))
        End If
    Else
        lblSum成本金额.Visible = False
        lbl成本金额差.Visible = False
        lbl1.Visible = False
        lbl2.Visible = False
        lbl3.Visible = False
    End If
    
    lbl2.Left = lbl1.Width + lbl1.Left + 200
    lbl3.Left = IIf(lbl2.Visible, lbl2.Width + lbl2.Left + 200, lbl2.Left)
    lblSum成本金额.Left = IIf(lbl3.Visible, lbl3.Width + lbl3.Left + 200, lbl3.Left)
    lbl成本金额差.Left = lblSum成本金额.Width + lblSum成本金额.Left + 200
    
    vsfList_EnterCell    '列出单据体
    
    SetStrikeColor
    With vsfList
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    vsfList.Redraw = flexRDDirect
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    stbThis.Panels(2).Text = "当前共有" & rsList.RecordCount & "张单据"
    rsList.Close
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = 1
    End If
    
    Cmd查阅.Enabled = Not (vsfList.TextMatrix(vsfList.Row, 0) = "" Or vsfList.Row = 0)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'表头列宽初始
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        If tbcDetail.Selected.Index = 1 Then '选择的是盘点表清单
            If mblnBootUp = False Then
                For intCol = 1 To .Cols - 1
                    If intCol = 1 Then
                        .ColWidth(intCol) = 2000
                    ElseIf intCol = .Cols - 2 Then
                        .ColWidth(intCol) = 0
                    Else
                        .ColWidth(intCol) = 1000
                    End If
                Next
            End If
        Else
            If mblnBootUp = False Then
                .ColWidth(1) = 2000
                .ColWidth(4) = 3000
            End If
        End If
        .ColWidth(.ColIndex("盘点成本金额")) = 1500
    End With
    
    Call RestoreFlexState(vsfList, tbcDetail.Selected.Caption)
    If tbcDetail.Selected.Index = 1 And mblnViewCost = False Then
        vsfList.colHidden(vsfList.ColIndex("盘点成本金额")) = True
        vsfList.colHidden(vsfList.ColIndex("成本金额差")) = True
    End If
End Sub



Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '记录单没有审核
    cbo已审核.Enabled = tbcDetail.Selected.Index = mconTab_CheckCard
    cbo已审核_Click
    
    Txt审核人.Enabled = tbcDetail.Selected.Index = mconTab_CheckCard
    If Txt审核人.Enabled Then
        Txt审核人.BackColor = &H80000005
    Else
        Txt审核人.BackColor = &H8000000F
    End If
    
    If Not mblnLoad Then Exit Sub

    Call SaveFlexState(vsfList, tbcDetail.Item(mintLastIndex).Caption)
    Call SaveFlexState(vsfDetail, tbcDetail.Item(mintLastIndex).Caption)
    
    mblnBootUp = False
    If Item.Index = 1 Then
        vsfDetail.ToolTipText = mcstComment
    Else
        vsfDetail.ToolTipText = ""
    End If
    
    If cbo已审核.Enabled Then
        If cbo未审核.ListIndex = 0 And cbo已审核.ListIndex = 0 Then
            MsgBox "对不起，必须选择一种单据显示（默认显示当天未审核单据）!", vbInformation, gstrSysName
            cbo未审核.ListIndex = 1
        End If
    Else
        If cbo未审核.ListIndex = 0 Then
            MsgBox "对不起，必须选择一种单据显示（默认显示当天未审核单据）!", vbInformation, gstrSysName
            cbo未审核.ListIndex = 1
        End If
    End If
    
    cmd确认_Click   '列出单据头
    
    mintLastIndex = Item.Index
    
    mblnBootUp = True
End Sub


Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    If mlngMode = 1307 Then
        If tbcDetail.Selected.Index = 1 Then
            '盘点表
            intNO = 29
        Else
            '盘点记录单
            intNO = 62
        End If
    End If
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, intNO, lng库房ID)
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    If mlngMode = 1307 Then
        If tbcDetail.Selected.Index = 1 Then
            '盘点表
            intNO = 29
        Else
            '盘点记录单
            intNO = 62
        End If
    End If
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, intNO, lng库房ID)
        End If
        Me.txt结束NO.SetFocus
    End If
End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then cmd确定.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            SendKeys vbTab
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)

        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取审核人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", _
                        Me.Txt审核人 & "%", gstrNodeNo)

        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = Txt填制人.Top + fraCondition.Top + Txt填制人.Height
                    .Left = Txt填制人.Left + fraCondition.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt审核人 = IIf(IsNull(!姓名), "", !姓名)
                SendKeys vbTab
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Me.Txt审核人.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt填制人.Text) = "" Then
            If tbcDetail.Selected.Index = mconTab_CheckCard Then
                Txt审核人.SetFocus
            Else
                SendKeys vbTab
            End If
            
            Exit Sub
        End If
        Txt填制人.Text = UCase(Txt填制人.Text)

        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取填制人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", _
                        Me.Txt填制人 & "%", gstrNodeNo)

        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = Txt填制人.Top + fraCondition.Top + Txt填制人.Height
                    .Left = Txt填制人.Left + fraCondition.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt填制人 = IIf(IsNull(!姓名), "", !姓名)
                If tbcDetail.Selected.Index = mconTab_CheckCard Then
                    Me.Txt审核人.SetFocus
                Else
                    SendKeys vbTab
                End If
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt填制人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药品.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fraCondition.Left + Txt药品.Left
    sngTop = Me.Top + fraCondition.Top + Txt药品.Top + Txt药品.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt药品.Height - 3630
    End If
    
    strkey = Trim(Txt药品.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "药品移库管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , True)

'    Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    Txt药品.Tag = RecReturn!药品id
 
    
End Sub

Private Sub Txt药品_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub vsfDetail_EnterCell()
    With vsfDetail
        If .Row = 0 Then Exit Sub
        
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfDetail_GotFocus()
    Call SetGridFocus(vsfDetail, True)
End Sub


Private Sub vsfDetail_LostFocus()
    Call SetGridFocus(vsfDetail, False)
End Sub


Private Sub vsfList_DblClick()
    If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_AddModify, , True).Visible Then Exit Sub
    If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_AddModify, , True).Enabled Then Exit Sub
    If vsfList.MouseRow = 0 Then Exit Sub
    cbsEditModify
End Sub

Private Sub vsfList_GotFocus()
    Call SetGridFocus(vsfList, True)
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_AddModify, , True).Visible Then Exit Sub
    If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_AddModify, , True).Enabled Then Exit Sub
        cbsEditModify
    End If
End Sub

Private Sub vsfList_LostFocus()
    Call SetGridFocus(vsfList, False)
End Sub

Private Sub vsfList_EnterCell()
    Dim rsDetail As New Recordset
    Dim intBill As Integer                      '单据类型  如：1、外购入库；2、
    Dim str包装系数 As String
    Dim str单位字段 As String
    Dim strOrder As String
    Dim strCompare As String
    Dim strSql效期 As String
    Dim lngColor As Long
    Dim n As Long
    Dim i As Integer
    Dim intCol As Integer
    Dim strSql药名 As String
    Dim strSqlOrder As String
    
    If Not mbln绑定 Then Exit Sub
    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    
    On Error GoTo ErrHandle
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品盘点)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "序号"
    
    If strCompare = "0" Then
        '按序号排序
        strSqlOrder = "序号"
    ElseIf strCompare = "1" Then
        '按编码排序
        strSqlOrder = "药品信息"
    ElseIf strCompare = "2" Then
        '按名称排序
        strSqlOrder = "Substr(药品信息, Instr(药品信息, ']') + 1)"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",药品信息,序号"
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        Select Case mintUnit
            Case mconint售价单位
                str包装系数 = "1"
                str单位字段 = "I.计算单位"
            Case mconint门诊单位
                str包装系数 = "B.门诊包装"
                str单位字段 = "B.门诊单位"
            Case mconint住院单位
                str包装系数 = "B.住院包装"
                str单位字段 = "B.住院单位"
            Case mconint药库单位
                str包装系数 = "B.药库包装"
                str单位字段 = "B.药库单位"
        End Select
        
        strSql效期 = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "TO_CHAR(A.效期-1,'YYYY-MM-DD') AS 有效期至", "TO_CHAR(A.效期,'YYYY-MM-DD') AS 失效期")
        
        If gint药品名称显示 = 0 Then
            strSql药名 = ",('['||I.编码||']'||I.名称) AS 药品信息"
        ElseIf gint药品名称显示 = 1 Then
            strSql药名 = ",('['||I.编码||']'||NVL(N.名称,I.名称)) AS 药品信息"
        Else
            strSql药名 = ",('['||I.编码||']'||I.名称) AS 药品信息,N.名称 As 商品名"
        End If
        
        intBill = IIf(tbcDetail.Selected.Index = 1, 12, 14)
        If tbcDetail.Selected.Index = 1 Then '选择的是盘点表清单
            gstrSQL = "Select DISTINCT a.序号" & strSql药名 & "," _
                    & "     B.药品来源,B.基本药物,I.规格,a.产地 as 生产商,a.原产地," & str单位字段 & " as 单位,a.批号," & strSql效期 & ",a.批准文号," _
                    & "     LTRIM(to_char(A.填写数量 /" & str包装系数 & ",decode(a.扣率,0,'999999999990.00000'," & mstrNumberFormat & "))) AS 帐面数," _
                    & "     LTRIM(to_char(A.扣率 /" & str包装系数 & "," & mstrNumberFormat & ")) AS 实盘数," _
                    & "     Decode(Sign(A.扣率-A.填写数量),-1,'亏',1,'盈','平') as 标志," _
                    & "     LTRIM(to_char(A.实际数量 /" & str包装系数 & ",decode(a.扣率,0,'999999999990.00000'," & mstrNumberFormat & "))) AS 数量差," _
                    & "     LTRIM(TO_CHAR (a.单量*" & str包装系数 & ", " & mstrCostFormat & ")) AS 成本价," _
                    & "     LTRIM(TO_CHAR (a.零售价*" & str包装系数 & ", " & mstrPriceFormat & ")) AS 售价," _
                    & "     LTRIM(TO_CHAR (a.零售金额*a.入出系数,decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) AS 金额差," _
                    & "     LTRIM(TO_CHAR (to_char((A.扣率-A.填写数量) /" & str包装系数 & "," & mstrNumberFormat & ") * TO_CHAR (a.零售价* Decode(记录状态, 1, 1, Decode(Mod(记录状态, 3), 0, 1, -1))*" & str包装系数 & ", " & mstrPriceFormat & "),decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," _
                    & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) AS 账面金额差," _
                    & "     LTRIM(TO_CHAR (a.差价*a.入出系数, decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) AS 差价差, " _
                    & "     LTrim(To_Char(to_char(A.扣率 /" & str包装系数 & "," & mstrNumberFormat & ")*TO_CHAR (a.零售价*" & str包装系数 & ", " & mstrPriceFormat & "), " & mstrMoneyFormat & ")) As 盘点金额," _
                    & "     LTrim(To_Char(((a.成本价+to_char(a.零售金额*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1)),decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) _
                    & "     )))-(a.成本金额+to_char(a.差价*a.入出系数*Decode(a.记录状态, 1, 1, Decode(Mod(a.记录状态, 3), 0, 1, -1)),decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))))," & mstrMoneyFormat & ")) as 盘点成本金额, " _
                    & "     ltrim(To_Char((a.零售金额*a.入出系数 - a.差价*a.入出系数 ), decode(nvl(a.扣率,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln零差价模式, " decode(nvl(b.是否零差价管理,0),1,decode(a.单量-a.零售价,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) As 成本金额差," _
                    & " Nvl(I.撤档时间, To_Date('3000-01-01', 'YYYY-MM-DD')) As 撤档时间,e.库房货位 " _
                    & " From (Select a.入出系数,a.记录状态,a.序号,a.药品id,a.产地,a.原产地,a.批号,a.效期,A.填写数量,A.扣率,A.实际数量,a.成本价,a.成本金额,a.零售价,a.零售金额,a.差价,a.单量,a.批准文号,a.库房id" _
                    & "         From 药品收发记录 a" _
                    & "        Where a.记录状态= [2] And a.单据= 12 And a.No=[1]) a," _
                    & "        药品规格 b,收费项目目录 I ,收费项目别名 n,药品储备限额 e" _
                    & " Where a.药品id = b.药品id And a.药品id = i.Id" _
                    & "        And a.药品id=n.收费细目id(+) And n.性质(+)=3 " _
                    & "        And a.药品id = e.药品id(+) and a.库房id = e.库房id(+) " _
                    & " ORDER BY " & strSqlOrder
        Else
            gstrSQL = "Select DISTINCT a.序号" & strSql药名 & "," _
                    & "     B.药品来源,B.基本药物,I.规格,a.产地 as 生产商,a.原产地," & str单位字段 & " as 单位,a.批号," & strSql效期 & ",a.批准文号," _
                    & "     to_char(A.扣率 /" & str包装系数 & "," & mstrNumberFormat & ") AS 实盘数" _
                    & " From (Select a.序号,a.药品id,a.产地,a.原产地,a.批号,a.效期,A.填写数量,A.扣率,A.实际数量,a.零售价,a.零售金额,a.差价,a.批准文号" _
                    & "         From 药品收发记录 a" _
                    & "        Where a.记录状态= 1 And a.单据= 14 And a.No=[1]) a," _
                    & "        药品规格 b,收费项目目录 I ,收费项目别名 n" _
                    & " Where a.药品id = b.药品id And a.药品id = i.Id" _
                    & "        And a.药品id=n.收费细目id(+) And n.性质(+)=3 " _
                    & " ORDER BY " & strSqlOrder
        End If
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, vsfList.TextMatrix(vsfList.Row, 0), vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2))
        
        Set vsfDetail.DataSource = rsDetail
        rsDetail.Close
        
        With vsfDetail
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
        End With

        
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Cols = IIf(tbcDetail.Selected.Index = 1, 25, 12)
            If gint药品名称显示 = 2 Then .Cols = .Cols + 1
            .rows = 2
            .Clear
            
            intCol = 0
            
            .TextMatrix(0, intCol) = "序号": intCol = intCol + 1
            .TextMatrix(0, intCol) = "药品信息": intCol = intCol + 1
            
            If gint药品名称显示 = 2 Then
                .TextMatrix(0, intCol) = "商品名": intCol = intCol + 1
            End If
            
            .TextMatrix(0, intCol) = "药品来源": intCol = intCol + 1
            .TextMatrix(0, intCol) = "基本药物": intCol = intCol + 1
            .TextMatrix(0, intCol) = "规格": intCol = intCol + 1
            .TextMatrix(0, intCol) = "生产商": intCol = intCol + 1
            .TextMatrix(0, intCol) = "原产地": intCol = intCol + 1
            .TextMatrix(0, intCol) = "单位": intCol = intCol + 1
            .TextMatrix(0, intCol) = "批号": intCol = intCol + 1
            .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期"): intCol = intCol + 1
            .TextMatrix(0, intCol) = "批准文号": intCol = intCol + 1
            If tbcDetail.Selected.Index = 0 Then
                .TextMatrix(0, intCol) = "实盘数": intCol = intCol + 1
            Else
                .TextMatrix(0, intCol) = "帐面数": intCol = intCol + 1
                .TextMatrix(0, intCol) = "实盘数": intCol = intCol + 1
                .TextMatrix(0, intCol) = "标志": intCol = intCol + 1
                .TextMatrix(0, intCol) = "数量差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "成本价": intCol = intCol + 1
                .TextMatrix(0, intCol) = "售价": intCol = intCol + 1
                .TextMatrix(0, intCol) = "金额差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "差价差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "盘点成本金额": intCol = intCol + 1
                .TextMatrix(0, intCol) = "盘点金额": intCol = intCol + 1
                .TextMatrix(0, intCol) = "成本金额差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "账面金额差": intCol = intCol + 1
                .TextMatrix(0, intCol) = "撤档时间": intCol = intCol + 1
                .TextMatrix(0, intCol) = "库房货位": intCol = intCol + 1
            End If
            
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
        End With
    End If
    
    With vsfDetail
        .colHidden(.ColIndex("药品来源")) = True  '默认不显示
        .colHidden(.ColIndex("基本药物")) = True  '默认不显示
        .colHidden(.ColIndex("批准文号")) = True  '默认不显示
        If tbcDetail.Selected.Index = 1 Then '只设置盘点表的明细列默认不显示
            .colHidden(.ColIndex("金额差")) = True  '默认不显示
            .colHidden(.ColIndex("差价差")) = True  '默认不显示
            .colHidden(.ColIndex("成本金额差")) = True  '默认不显示
            .colHidden(.ColIndex("账面金额差")) = True  '默认不显示
            .colHidden(.ColIndex("库房货位")) = True  '默认不显示
        End If
    End With
    
    vsfHidden vsfDetail
    SetDetailColWidth
    
    '上色
    If tbcDetail.Selected.Index = 1 Then
        With vsfDetail
            .Redraw = flexRDNone
            For n = 1 To .rows - 1
                If .TextMatrix(n, 0) <> "" Then
                    If .TextMatrix(n, .ColIndex("标志")) = "盈" Then
                        lngColor = vbRed
                    ElseIf .TextMatrix(n, .ColIndex("标志")) = "亏" Then
                        lngColor = vbBlue
                    Else
                        lngColor = vbBlack
                    End If
                    
                    '盘亏盘盈行用颜色区分；
                    If lngColor <> vbBlack Then
                        .Cell(flexcpForeColor, n, 0, n, .Cols - 1) = lngColor
                    End If
                    
                    '如果是停用药品，该行粗体显示
                    If Format(.TextMatrix(n, .ColIndex("撤档时间")), "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, n, 0, n, .Cols - 1) = True
                    End If
                End If
            Next
            .Redraw = flexRDDirect
        End With
    End If
    
    vsfDetail.Row = 1
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    
    On Error GoTo ErrHandle
    
    With vsfDetail
        .ColAlignment(.ColIndex("单位")) = flexAlignCenterCenter    '单位
        .ColAlignment(.ColIndex("实盘数")) = flexAlignRightCenter '实盘数
        If tbcDetail.Selected.Index = 1 Then
            .ColAlignment(.ColIndex("帐面数")) = flexAlignRightCenter     '帐面数
            .ColAlignment(.ColIndex("标志")) = flexAlignCenterCenter    '标志
            .ColAlignment(.ColIndex("数量差")) = flexAlignRightCenter     '数量差
            .ColAlignment(.ColIndex("成本价")) = flexAlignRightCenter    '成本价
            .ColAlignment(.ColIndex("售价")) = flexAlignRightCenter    '售价
            .ColAlignment(.ColIndex("金额差")) = flexAlignRightCenter    '金额差
            .ColAlignment(.ColIndex("差价差")) = flexAlignRightCenter    '差价差
            .ColAlignment(.ColIndex("盘点金额")) = flexAlignRightCenter    '盘点金额
            .ColAlignment(.ColIndex("账面金额差")) = flexAlignRightCenter    '账面金额差
            .ColAlignment(.ColIndex("盘点成本金额")) = flexAlignRightCenter    '盘点成本金额
            .ColAlignment(.ColIndex("成本金额差")) = flexAlignRightCenter    '成本金额差
            
        End If
        
        If tbcDetail.Selected.Index = 1 Then
            If mblnBootUp = False Then
                .ColWidth(0) = 500
                .ColWidth(.ColIndex("药品信息")) = 2500
                For intCol = 2 To .Cols - 1
                    .ColWidth(intCol) = 1000
                Next
                .ColWidth(.ColIndex("撤档时间")) = 0
                .ColWidth(.ColIndex("盘点成本金额")) = 1500
            End If
        Else
            .ColWidth(0) = 500
            .ColWidth(.ColIndex("药品信息")) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        
        Call RestoreFlexState(vsfDetail, tbcDetail.Selected.Caption)
        If tbcDetail.Selected.Index = 1 And mblnViewCost = False Then
            .colHidden(.ColIndex("成本价")) = True
            .colHidden(.ColIndex("差价差")) = True
            .colHidden(.ColIndex("盘点成本金额")) = True
            .colHidden(.ColIndex("成本金额差")) = True
        End If
        
        str库房性质 = ""
        gstrSQL = "Select a.工作性质 From 部门性质说明 A Where a.部门id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断是库房性质", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str库房性质 = str库房性质 & "," & rsDetail!工作性质
            rsDetail.MoveNext
        Loop
        If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
        If bln中药库房 Then
            .colHidden(.ColIndex("原产地")) = False
        Else
            .colHidden(.ColIndex("原产地")) = True
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    With vsfList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            intStatus = IIf(tbcDetail.Selected.Index = 0, 1, Val(.TextMatrix(intRow, .Cols - 2)))
            If intStatus Mod 3 = 0 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001
            End If
            If intStatus Mod 3 = 2 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF
            End If
        Next
    End With
End Sub


Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Format(mstrStart, "yyyy-mm-dd") = "1901-01-01" And Format(mstrVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "审核日期 " & Format(mstrVerifyStart, "yyyy年MM月dd日") & "至" & Format(mstrVerifyEnd, "yyyy年MM月dd日")
    ElseIf Format(mstrVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "填制日期 " & Format(mstrStart, "yyyy年MM月dd日") & "至" & Format(mstrEnd, "yyyy年MM月dd日") & "  审核日期 " & Format(mstrVerifyStart, "yyyy年MM月dd日") & "至" & Format(mstrVerifyEnd, "yyyy年MM月dd日")
    Else
        strRange = "填制日期 " & Format(mstrStart, "yyyy年MM月dd日") & "至" & Format(mstrEnd, "yyyy年MM月dd日")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    objRow.Add "时间：" & strRange
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfList
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub subExcel(bytMode As Byte)
'功能:进行输出到EXCEL
'参数:bytMode3 输出到EXCEL

    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "NO")))
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "盘点库房：" & Trim(cboStock.Text)
    objRow.Add "盘点时间：" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "盘点时间")))
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "摘要:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "摘要"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "填制人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "填制人")) & "  填制日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "填制日期"))
    
    If tbcDetail.Selected.Index = 1 Then
        objRow.Add "审核人:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "审核人")) & "  审核日期:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "审核日期"))
        objPrint.BelowAppRows.Add objRow
    End If
    
    Set objPrint.Body = vsfDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button <> 2 Then Exit Sub
    
    If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_EditPopup, , True).Visible Then Exit Sub '编辑不可见退出
    
    Set objPopup = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_EditPopup)
    If Not objPopup Is Nothing Then
        objPopup.CommandBar.ShowPopup
    End If
    
End Sub

'功能：将vsf表格存在的列并且在列设置中未勾选隐藏的列进行隐藏
Private Sub vsfHidden(ByRef objVSF As Object)
    Dim strColsName As String
    Dim strColName() As String
    Dim i As Integer
    Dim n As Integer
    Dim strDefaultColsName As String '默认的列
    Dim strTempColName As String
    
    strDefaultColsName = ":药品来源,0:基本药物,0:库房货位,0:批准文号,0:金额差,0:差价差,0:盘点成本金额差,0:账面金额差,0:成本金额差,0:当前库存,1:" '所有可以隐藏的列
    
    With objVSF
        strColsName = zlDataBase.GetPara("列设置", glngSys, mlngMode, "")
        
        '兼容处理
        If strColsName = "" Then '未提取到列设置信息
            strColsName = strDefaultColsName
        Else
            '判断提取的列与默认列个数，不一致则取默认的
            If UBound(Split(strColsName, ":")) <> UBound(Split(strDefaultColsName, ":")) Then strColsName = strDefaultColsName
            
            '判断提取的列名是否与默认的一致，不一致取默认的
            For i = LBound(Split(strColsName, ":")) + 1 To UBound(Split(strColsName, ":")) - 1
                strTempColName = Split(Split(strColsName, ":")(i), ",")(0) '获取单个列名
                
                If InStr(1, strDefaultColsName, ":" & strTempColName) = 0 Then '列名不存在于默认列名中
                    strColsName = strDefaultColsName
                    Exit For
                End If
            Next
            
        End If
        
        strColName = Split(strColsName, ":") '格式:C,1
        
        For i = 0 To .Cols - 1
            '判断界面对应列是否是可屏蔽列
            If InStr(1, strColsName, ":" & .TextMatrix(0, i)) > 0 Then
                For n = LBound(strColName) + 1 To UBound(strColName) - 1
                    If Split(strColName(n), ",")(0) = .TextMatrix(0, i) Then .colHidden(i) = Split(strColName(n), ",")(1) <> 1
                Next
            End If
             
        Next
    End With
End Sub
