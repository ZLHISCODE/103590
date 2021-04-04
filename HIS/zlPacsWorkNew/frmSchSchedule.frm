VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "*\A..\ZLIDKind\ZLIDKIND.vbp"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "CODEJOCK.CALENDAR.V16.3.1.OCX"
Begin VB.Form frmSchSchedule 
   Caption         =   "检查项目预约"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13545
   Icon            =   "frmSchSchedule.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   13545
   StartUpPosition =   1  '所有者中心
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   9
      FontName        =   "宋体"
      IDKind          =   -1
      BackColor       =   -2147483633
   End
   Begin VB.PictureBox pictDay 
      BackColor       =   &H8000000E&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picTimeTable 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   4920
      ScaleHeight     =   7095
      ScaleWidth      =   5655
      TabIndex        =   3
      Top             =   240
      Width           =   5655
      Begin VB.Frame frmTimeTable 
         Caption         =   "预约时间表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   8415
         Begin zl9PACSWork.ucScheduleTimetable schTimeTable 
            Height          =   6615
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   11668
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   960
      ScaleHeight     =   9135
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Frame frmInfo 
         Caption         =   "基本信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3615
         Begin VB.ComboBox cboSchDevice 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   330
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   3045
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtNotice 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1635
            Width           =   2055
         End
         Begin VB.TextBox txtPhone 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label lblSchDevice 
            Caption         =   "CT"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   1400
            TabIndex        =   21
            Top             =   3090
            Width           =   1335
         End
         Begin VB.Label lblOrderInfo 
            Caption         =   "医嘱内容："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   2100
            Width           =   1095
         End
         Begin VB.Label lblAddress 
            Caption         =   "检查注意："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1665
            Width           =   1095
         End
         Begin VB.Label lblPhone 
            Caption         =   "电话："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1230
            Width           =   735
         End
         Begin VB.Label lblInfo 
            Caption         =   "预约时间：10:20 - 11:30"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   12
            Top             =   3960
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            Caption         =   "预约日期：2018-2-3"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   11
            Top             =   3525
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            Caption         =   $"frmSchSchedule.frx":0442
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Top             =   2415
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            Caption         =   "预约设备："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   7
            Top             =   3090
            Width           =   1335
         End
         Begin VB.Label lblInfo 
            Caption         =   "年龄：25   来源：门诊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   6
            Top             =   795
            Width           =   2895
         End
         Begin VB.Label lblInfo 
            Caption         =   "姓名：张晓    性别：男"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Width           =   3135
         End
      End
      Begin XtremeCalendarControl.DatePicker dpCalendar 
         Height          =   2895
         Left            =   0
         TabIndex        =   2
         Top             =   4320
         Width           =   3615
         _Version        =   1048579
         _ExtentX        =   6376
         _ExtentY        =   5106
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowNoneButton  =   0   'False
         ShowNonMonthDays=   0   'False
         Show3DBorder    =   0
         AskDayMetrics   =   -1  'True
         TextTodayButton =   "选择今天"
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSchOther 
         Height          =   1815
         Left            =   0
         TabIndex        =   8
         Top             =   7320
         Width           =   3615
         _cx             =   6376
         _cy             =   3201
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         ExplorerBar     =   0
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   8940
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4154
            MinWidth        =   4154
            Picture         =   "frmSchSchedule.frx":0454
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17004
            MinWidth        =   7056
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   240
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":0CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":288C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":365E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":4430
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":5202
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":5FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":6DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":7B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":894A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":971C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":A4EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":B2C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":C092
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":CE64
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":DC36
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":EA08
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":F7DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":105AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1137E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":12150
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":12F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":13CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":14AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":15898
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1666A
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1743C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1820E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":18FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":19DB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSchSchedule.frx":1AB84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   120
      Top             =   1440
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmSchSchedule.frx":1B956
      Left            =   360
      Top             =   360
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSchSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngOrderID As Long                 '医嘱ID
Private mschDate As Date                    '当前的预约日期
Private mlngSchDeviceID As Long             '当前选中的预约设备ID
Private mblnISScheduled As Boolean          '是否已经预约
Private mblnNewSchedule As Boolean          '是否新建预约
Private mstrDefaultPatientType As String    '缺省病人类型
Private mfrmParent As Object                '父窗体
Private mstrDeptIDs As String               '科室ID串
Private mlngDeptID As Long                  '当前科室ID
Private mstrModifiedOrderID As String       '保存过预约信息的医嘱ID串，用“,”连接
Private mblnExecFee As Boolean              '是否预约时执行费用
Private mblnAutoPrint As Boolean            '是否自动打印预约单
Private mstrSchRestDate As String           '当月休息日
Private mblnCheckIn As Boolean              '是否保存预约后报到
Private mblnLoadingDevice As Boolean        '是否正在加载预约设备
Private mlngPatSource As Long               '病人来源
Private mstrPrivs As String                 '调用者的权限
Private mblnIsForceModify As Boolean        '是否强制修改住院门诊信息
Private mblnLoadDone As Boolean             '是否完成窗体的加载

'检查预约设备
Private Enum constScheduleDeviceList
    col_SchDevice_ID = 0
    col_SchDevice_选中 = 1
    col_SchDevice_影像类别 = 2
    col_SchDevice_设备名称 = 3
    col_SchDevice_设备说明 = 4
End Enum

'其他设备上的预约信息
Private Enum constSchOtherList
    col_SchOther_ID = 0
    col_SchOther_预约设备名称 = 1
    col_SchOther_预约日期 = 2
    col_SchOther_医嘱内容 = 3
    col_SchOther_预约开始时间 = 4
    col_SchOther_预约结束时间 = 5
End Enum

Private Sub InitCommandBar()
'------------------------------------------------
'功能：初始化工具栏
'参数： 无
'返回： 无
'------------------------------------------------
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    
    On Error GoTo err
    
    '这部分全局设置，是否必要？
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbrMain.VisualTheme = xtpThemeOffice2003
    Set cbrMain.Icons = zlCommFun.GetPubIcons
        
    With cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True         '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '不显示菜单
    cbrMain.ActiveMenuBar.Visible = False
    
    '显示工具栏
    Set cbrToolBar = cbrMain.Add("预约工具栏", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Save, "保存预约")
        cbrControl.iconid = 6823
        cbrControl.ToolTipText = "保存预约信息"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Print, "打印预约单")
        cbrControl.iconid = 103
        cbrControl.ToolTipText = "打印患者的预约通知单"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_New, "新建预约")
        cbrControl.iconid = 6886
        cbrControl.ToolTipText = "新建一个检查预约"
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Delete, "删除预约")
        cbrControl.iconid = 6822
        cbrControl.ToolTipText = "删除一个检查预约"
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Refresh, "刷新")
        cbrControl.iconid = 791
        cbrControl.ToolTipText = "刷新数据"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_SaveAndCheckin, "保存报到")
        cbrControl.iconid = 744
        cbrControl.ToolTipText = "保存预约，关闭窗口，检查报到"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_SaveAndQuit, "保存退出")
        cbrControl.iconid = 3013
        cbrControl.ToolTipText = "保存预约，关闭窗口"
         
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Quit, "退出")
        cbrControl.iconid = 191
        cbrControl.ToolTipText = "关闭窗口"
        
    End With
    
    cbrToolBar.Position = xtpBarTop
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboSchDevice_Click()
    If cboSchDevice.ListIndex >= 0 And mblnLoadingDevice = False Then
        '修改当前被选中的预约设备ID
        mlngSchDeviceID = cboSchDevice.ItemData(cboSchDevice.ListIndex)
        Call RefreshSchedule(False, True)
        Call RefreshCalendar
    End If
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsSchdule_New        '新建预约
            Call NewSchedule
            
        Case conMenu_PacsSchdule_Delete     '删除预约
            Call DelSchedule(mlngOrderID)
            
        Case conMenu_PacsSchdule_Print      '打印预约单
            Call PrintSchedule
            
        Case conMenu_PacsSchdule_Refresh    '刷新
            Call RefreshForm
            
        Case conMenu_PacsSchdule_ModifyInfo '修改信息
            Call ModifyPatInfo
        
        Case conMenu_PacsSchdule_Save       '保存预约
            Call SaveSchedule
            Call loadTimeTable
            
        Case conMenu_PacsSchdule_SaveAndCheckin '保存退出且打开报到窗口
            If SaveSchedule = True Then
                mblnCheckIn = True
                Unload Me
            End If

        Case conMenu_PacsSchdule_SaveAndQuit '保存退出
            If SaveSchedule = True Then
                Unload Me
            End If
        
        Case conMenu_PacsSchdule_Quit       '退出
            Unload Me
            
    End Select
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsSchdule_Delete     '删除预约
            Control.Enabled = mblnISScheduled
            
        Case conMenu_PacsSchdule_Print      '打印预约单
            Control.Enabled = mblnISScheduled
        
        Case conMenu_PacsSchdule_Save, conMenu_PacsSchdule_SaveAndCheckin, _
             conMenu_PacsSchdule_SaveAndQuit     '保存预约,保存退出且打开报到窗口,保存退出
            Control.Enabled = mblnNewSchedule
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picInfo.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picTimeTable.hwnd
    End If
End Sub

Private Sub dpCalendar_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If InStr(mstrSchRestDate, Format(Day, "YYYY-MM-DD")) > 0 Then
        Metrics.ForeColor = vbRed
        Metrics.Font.Bold = True
    End If
End Sub

Private Sub dpCalendar_MonthChanged()
    If mblnISScheduled = True Or (Format(dpCalendar.FirstVisibleDay, "YYYY-MM") < Format(Now, "YYYY-MM")) Then
        ChangeCalendar (mschDate)
    Else
        Call RefreshCalendar
    End If
End Sub

Private Sub dpCalendar_SelectionChanged()
    Dim dtDate As Date
    
    '更换了日期，重新刷新时间表
    dtDate = dpCalendar.Selection.Blocks(0).DateBegin
    If InStr(mstrSchRestDate, Format(dtDate, "YYYY-MM-DD")) > 0 Or mblnISScheduled = True Then
        If dtDate = Format(Now, "YYYY-MM-DD") And mblnISScheduled = False Then
            Call MsgBox("今天预约已经满了。", vbInformation, "检查预约提示")
        End If
        '是无法预约的日子，不选择
        ChangeCalendar (mschDate)
    Else
        mschDate = dtDate
    End If
    
    Call RefreshSchedule(False, True)
End Sub

Private Sub InitFaceScheme()
'------------------------------------------------
'功能：初始化界面布局
'参数： 无
'返回： 无
'------------------------------------------------
    Dim Pane1 As Pane, Pane2 As Pane
    
    On Error GoTo err
    
    '设置总体显示策略
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .SetCommandBars cbrMain
        .options.HideClient = True
        .options.UseSplitterTracker = False '实时拖动
        .options.ThemedFloatingFrames = True
        .options.AlphaDockingContext = True
        dkpMain.options.DefaultPaneOptions = PaneNoCaption
    End With
    
    '先从注册表读取预先设置好的窗口布局，然后再逐个设置
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    '如果注册表中保存的界面布局Pane数量不对，则加载默认的Pane设置
    If dkpMain.PanesCount <> 2 Then
        dkpMain.DestroyAll
        
        Set Pane1 = dkpMain.CreatePane(1, 350, 150, DockLeftOf)
        Pane1.title = "预约信息"
        Pane1.options = PaneNoCaption
        
        Set Pane2 = dkpMain.CreatePane(2, 650, 300, DockRightOf, Pane1)
        Pane2.title = "预约时间表"
        Pane2.options = PaneNoCaption
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    If mblnLoadDone = True Then
        If mblnNewSchedule = True Then
            If MsgBox("是否保存检查预约信息？", vbYesNo, "检查预约提示") = vbYes Then
                Call SaveSchedule
            End If
        End If
        
        '关闭窗体的时候，保存界面布局
        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        
        Call SaveWinState(Me, App.ProductName)
    End If
    
    Set mfrmParent = Nothing
    ' '关闭窗体时，先释放，会导致VB崩溃
'    '释放DockingPane
'    For i = 1 To dkpMain.PanesCount
'        dkpMain.Panes(i).Handle = 0
'    Next i
'    dkpMain.CloseAll
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    frmInfo.Left = 30
    frmInfo.Top = 50
    frmInfo.Width = picInfo.ScaleWidth - 30
    
    lblInfo(0).Width = frmInfo.Width - lblInfo(0).Left - 50
    lblInfo(1).Width = lblInfo(0).Width
    lblInfo(2).Width = lblInfo(0).Width
    lblInfo(4).Width = lblInfo(0).Width
    lblInfo(5).Width = lblInfo(0).Width
    txtPhone.Width = frmInfo.Width - 1450
    txtNotice.Width = txtPhone.Width
    lblSchDevice.Width = txtPhone.Width
    cboSchDevice.Width = txtPhone.Width
    
    dpCalendar.Left = 0
    dpCalendar.Top = frmInfo.Height + 10
    dpCalendar.Width = frmInfo.Width
    
    vsfSchOther.Left = 0
    vsfSchOther.Top = dpCalendar.Top + dpCalendar.Height + 30
    vsfSchOther.Width = frmInfo.Width
    vsfSchOther.Height = picInfo.ScaleHeight - vsfSchOther.Top - 300
End Sub

Private Sub picTimeTable_Resize()
    On Error Resume Next
    
    frmTimeTable.Left = 0
    frmTimeTable.Top = 0
    frmTimeTable.Width = picTimeTable.ScaleWidth
    frmTimeTable.Height = picTimeTable.ScaleHeight - stbThis.Height
    
    schTimeTable.Left = 0
    schTimeTable.Top = 0
    schTimeTable.Width = frmTimeTable.Width
    schTimeTable.Height = frmTimeTable.Height
End Sub

Public Function ZlShowMe(ByVal strPrivs As String, ByVal lngOrderID As Long, ByVal strDeptIDs As String, _
    ByVal frmParent As Object, Optional ByRef blnCheckin As Boolean = False) As String
'------------------------------------------------
'功能：打开窗口
'参数： lngOrderID -- 医嘱ID
'       strDeptIDs -- 科室ID串
'       frmParent -- 父窗体
'       strPrivs -- 调用者的权限
'返回：保存预约后的医嘱ID
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    mblnLoadDone = False
    mlngOrderID = lngOrderID
    Set mfrmParent = frmParent
    mstrDeptIDs = strDeptIDs
    mstrModifiedOrderID = ""
    mlngSchDeviceID = 0
    mblnCheckIn = False
    mstrPrivs = strPrivs
    mblnIsForceModify = CheckPopedom(mstrPrivs, "强制修改住院门诊信息")
    
    '如果是全部科室，先查执行科室ID，如果没有则取第一个科室
    If InStr(mstrDeptIDs, ",") > 0 Then
        strSQL = "select 执行科室ID from 病人医嘱记录 where id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取当前科室ID", mlngOrderID)
        If rsTemp.EOF = False Then
            mlngDeptID = NVL(rsTemp!执行科室ID)
        Else
            mlngDeptID = Split(mstrDeptIDs, ",")(0)
        End If
    Else
        mlngDeptID = Val(mstrDeptIDs)
    End If
    
    '读取参数
    mblnExecFee = IIf(Val(zlDatabase.GetPara("预约时执行费用", glngSys, 1292)) = 1, True, False)
    mblnAutoPrint = IIf(Val(zlDatabase.GetPara("保存预约后自动打印预约单", glngSys, 1292)) = 1, True, False)
    
    '初始化界面布局
    Call InitFaceScheme
    
    '创建工具栏
    Call InitCommandBar
    
    '先初始化时间表控件
    Call schTimeTable.Init(1)   '检查项目预约
    
    Call RestoreWinState(Me, App.ProductName)
    
    '设置日历参数
    dpCalendar.AskDayMetrics = True
    dpCalendar.ShowNonMonthDays = False
    mschDate = Format(Now, "YYYY-MM-DD")
    
    '加载数据
    If LoadData = False Then
        Unload Me
        Exit Function
    End If
    
    Call RefreshCalendar
    
    mblnLoadDone = True
    
    Me.Show 1, mfrmParent
    
    blnCheckin = mblnCheckIn
    ZlShowMe = mstrModifiedOrderID
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadSchDevice() As Boolean
'------------------------------------------------
'功能：加载预约设备
'参数：
'返回：True -- 成功；False -- 失败
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim iSelRow As Integer
    
    On Error GoTo err
    
    iSelRow = -1
    mblnLoadingDevice = True
    
    strSQL = "Select  distinct a.id, a.设备名称, a.影像类别, a.设备说明, a.是否默认" _
            & " From 影像预约设备 A, 病人医嘱记录 B , 影像预约项目 c " _
            & " Where a.id = c.预约设备id And b.诊疗项目id = c.诊疗项目id And a.是否启用 = 1 " _
            & " And b.ID = [1]  And  a.科室id In (" & mstrDeptIDs & ") order by 是否默认 desc"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约设备", mlngOrderID)
    
    '从数据库加载数据
    If rsTemp.RecordCount = 1 Then
        lblSchDevice.Visible = True
        cboSchDevice.Visible = False
        lblSchDevice.Caption = rsTemp!设备名称
        mlngSchDeviceID = rsTemp!ID
    Else
        lblSchDevice.Visible = False
        cboSchDevice.Visible = True
        
        For i = 1 To rsTemp.RecordCount
            cboSchDevice.AddItem (rsTemp!设备名称)
            cboSchDevice.ItemData(cboSchDevice.NewIndex) = rsTemp!ID
            If rsTemp!ID = mlngSchDeviceID Then
                iSelRow = cboSchDevice.NewIndex
            End If
            rsTemp.MoveNext
        Next i
        If iSelRow <> -1 Then
            cboSchDevice.ListIndex = iSelRow
        ElseIf cboSchDevice.ListCount > 1 Then
            cboSchDevice.ListIndex = 0
            mlngSchDeviceID = cboSchDevice.ItemData(0)
        Else
            mlngSchDeviceID = 0
            Call MsgBoxD(Me, "没有可用于预约的影像设备，请先添加预约设备。", vbOKOnly, "检查预约提示")
            mblnLoadingDevice = False
            Exit Function
        End If
    End If
    mblnLoadingDevice = False
    
    LoadSchDevice = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnLoadingDevice = False
End Function

Private Sub LoadSchOther()
'------------------------------------------------
'功能：加载患者在其他设备上的预约信息
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    strSQL = "Select a.id, a.医嘱id, a.预约设备名称, b.医嘱内容, a.预约开始时间, " _
        & " a.预约结束时间 From 影像预约记录 A, 病人医嘱记录 B, 病人医嘱发送 C " _
        & " Where a.医嘱ID = b.ID And b.ID = c.医嘱ID And c.执行状态 = 0 And b.病人id = " _
        & " (Select f.病人id From 病人医嘱记录 F Where f.ID = [1]) And a.医嘱id <> [1] order by 预约开始时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询在其他设备上的预约", mlngOrderID)
    
    With vsfSchOther
        .Rows = rsTemp.RecordCount + 2
        .Cols = 6
        .FixedRows = 2
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
        .ScrollBars = flexScrollBarBoth
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, 2) = flexAlignCenterCenter
        .ExtendLastCol = True
        
        .ColWidth(col_SchOther_ID) = 50
        .ColWidth(col_SchOther_医嘱内容) = 2000
        .ColWidth(col_SchOther_预约日期) = 1200
        .ColWidth(col_SchOther_预约设备名称) = 1000
        .ColWidth(col_SchOther_预约开始时间) = 1000
        .ColWidth(col_SchOther_预约结束时间) = 1000
        
        '合并第一行
        .RowHeight(0) = 350
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        For i = 0 To 5
            .TextMatrix(0, i) = "其他检查预约信息"
        Next i
        .Cell(flexcpFontBold, 0, 0, 0, 5) = True
        
        '第二行显示标题
        .TextMatrix(1, col_SchOther_医嘱内容) = "医嘱内容"
        .TextMatrix(1, col_SchOther_预约日期) = "预约日期"
        .TextMatrix(1, col_SchOther_预约设备名称) = "预约设备"
        .TextMatrix(1, col_SchOther_预约开始时间) = "开始时间"
        .TextMatrix(1, col_SchOther_预约结束时间) = "结束时间"
        .RowHeight(1) = 300
        
        '从数据库加载数据
        i = 1
        While rsTemp.EOF = False
            If mlngOrderID <> rsTemp!医嘱ID Then
                .TextMatrix(i + 1, col_SchOther_ID) = rsTemp!ID
                .TextMatrix(i + 1, col_SchOther_医嘱内容) = rsTemp!医嘱内容
                .TextMatrix(i + 1, col_SchOther_预约日期) = Format(rsTemp!预约开始时间, "yyyy-mm-dd")
                .TextMatrix(i + 1, col_SchOther_预约设备名称) = rsTemp!预约设备名称
                .TextMatrix(i + 1, col_SchOther_预约开始时间) = Format(rsTemp!预约开始时间, "HH:MM")
                .TextMatrix(i + 1, col_SchOther_预约结束时间) = Format(rsTemp!预约结束时间, "HH:MM")
                i = i + 1
            End If
            rsTemp.MoveNext
        Wend
        
        '隐藏后台数据列
        .ColHidden(col_SchOther_ID) = True
        
    End With

    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function loadTimeTable() As Boolean
'------------------------------------------------
'功能：加载刷新时间表内容，如果已经有预约信息，根据预约信息，重新设置程序界面的预约设备和日期
'参数：
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim strSQL  As String
    Dim rsTemp As ADODB.Recordset
    Dim lngSchDeviceID As Long
    Dim dtSchDate As Date
    
    On Error GoTo err
    
    If mblnISScheduled = True Then
        '已经存在预约信息，直接显示即可
        If schTimeTable.RefreshSchedule(mlngSchDeviceID, mschDate, mlngOrderID) = False Then
            Exit Function
        End If
        mblnNewSchedule = False
    Else
        mblnNewSchedule = True
        dtSchDate = mschDate
        If schTimeTable.NewSchedule(mlngSchDeviceID, mschDate, mlngOrderID, True) = False Then
            Exit Function
        End If
        If dtSchDate <> mschDate Then
            '如果日期被变更了，则修改日历里面的日期
            Call ChangeCalendar(mschDate)
        End If
    End If
    stbThis.Panels(2).Text = schTimeTable.LabelOrderInfo
    loadTimeTable = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadData() As Boolean
'------------------------------------------------
'功能：加载窗体的所有数据
'参数：
'返回：True -- 成功； False -- 失败
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    '有先后顺序
    If LoadSchDevice = False Then
        Exit Function
    End If
    
    mblnISScheduled = False
    
    '刷新患者基本信息
    Call RefreshSchInfo(True)
    
    '设置预约日历
    Call ChangeCalendar(mschDate)
    
    '提取缺省病人类型
    strSQL = "select 名称 from 病人类型 where 缺省标志=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取缺省病人类型")
    If rsTemp.RecordCount > 0 Then mstrDefaultPatientType = NVL(rsTemp!名称)
    
    Call LoadSchOther
    
    If loadTimeTable = False Then
        Exit Function
    End If
    
    LoadData = True

    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub schTimeTable_OnMenuSchedulePrint()
    Call PrintSchedule
End Sub

Private Sub schTimeTable_OnSchLabelModifed(ByVal iIndex As Integer)
    stbThis.Panels(2).Text = schTimeTable.LabelOrderInfo
    mblnNewSchedule = True
End Sub

Private Sub txtNotice_Change()
    If txtNotice.Locked = False Then
        txtNotice.ForeColor = vbRed
        mblnNewSchedule = True
    End If
End Sub

Private Sub txtNotice_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtNotice_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub txtNotice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        txtNotice.Locked = False
    End If
End Sub

Private Sub txtPhone_Change()
    If txtPhone.Locked = False Then
        txtPhone.ForeColor = vbRed
        mblnNewSchedule = True
    End If
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtPhone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And (mlngPatSource = 3 Or mblnIsForceModify = True) Then    '只有外诊的手机号才能修改
        txtPhone.Locked = False
    End If
End Sub

Private Function ValidData() As Boolean
'------------------------------------------------
'功能：检查数据合法性
'参数：
'返回：True -- 合法；False -- 数据不合法，需要修改
'------------------------------------------------
    On Error GoTo err
    
    '手机号合法性检查
    If Trim(txtPhone.Text) <> "" Then
        If Not IDKind.IsMobileNo(Trim(txtPhone.Text)) Then
            MsgBox "[手机号]无效,请重新录入或者删除已录入内容!", vbInformation, gstrSysName
            If txtPhone.Enabled And txtPhone.Visible Then txtPhone.SetFocus: Exit Function
        End If
    End If
    
    ValidData = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveSchedule() As Boolean
'------------------------------------------------
'功能：保存预约时间
'参数：
'返回：True -- 保存成功；False -- 保存失败
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim lngSendNo As Long
    Dim lngState As Long
    Dim strStartTime As String
    Dim strEndTime As String
    
    On Error GoTo err
    
    SaveSchedule = False
    
    If ValidData() = False Then
        Exit Function
    End If
    
    If schTimeTable.Label序号 <> 0 Then
        
        If schTimeTable.funSaveSchedule(schTimeTable.Label开始时间, schTimeTable.Label结束时间, mlngOrderID, _
                schTimeTable.Label姓名, schTimeTable.Label序号, mlngSchDeviceID, schTimeTable.Label开始时间段, _
                schTimeTable.Label结束时间段, txtNotice.Text) = False Then
                
            Exit Function
        End If
        
        '执行费用
        If mblnExecFee = True Then
            strSQL = "select 发送号,执行部门ID,执行过程 from 病人医嘱发送 where 医嘱ID =[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询费用执行信息", mlngOrderID)
            If rsTemp.EOF = False Then
                strSQL = "zl_影像费用执行(" & mlngOrderID & "," & Val(rsTemp!发送号) & "," & Val(NVL(rsTemp!执行过程, 0)) _
                    & ",null,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & Val(rsTemp!执行部门ID) & ")"
                zlDatabase.ExecuteProcedure strSQL, "预约时执行费用"
            End If
        End If
        
        '保存患者联系信息
        If txtPhone.ForeColor = vbRed Then
            strSQL = "select b.病人来源,b.婴儿,a.病人id,a.姓名,a.性别,a.年龄,a.费别," _
                & " a.医疗付款方式,a.民族,a.婚姻状况,a.职业,a.身份证号,a.家庭电话, a.家庭地址邮编," _
                & " a.出生日期,a.主页id,a.家庭地址 from 病人信息 a,病人医嘱记录 b where " _
                & " a.病人id=b.病人id and b.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询病人基本信息", mlngOrderID)
            If rsTemp.EOF = False Then
                If NVL(rsTemp!婴儿, 0) <> 0 Then
                    strSQL = "select  婴儿姓名 ,婴儿性别 , 出生时间  from   病人新生儿记录 " _
                        & " Where 病人ID=[1] And 主页ID=[2] And 序号=[3]"
                    Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "查询婴儿信息", CLng(rsTemp!病人ID), CLng(NVL(rsTemp!主页ID, 0)), CLng(NVL(rsTemp!婴儿, 0)))

                    If rsBaby.EOF = False Then
                        strSQL = "zl_影像病人信息_修改(" & NVL(rsTemp!病人来源) & "," & mlngOrderID & "," _
                        & rsTemp!病人ID & ",'" & NVL(rsBaby!婴儿姓名) & "','" & NVL(rsBaby!婴儿性别) & "','" _
                        & NVL(rsTemp!年龄) & "','" & NVL(rsTemp!费别) & "','" & NVL(rsTemp!医疗付款方式) _
                        & "','" & NVL(rsTemp!民族) & "','" & NVL(rsTemp!婚姻状况) & "','" & NVL(rsTemp!职业) _
                        & "','" & NVL(rsTemp!身份证号) & "','" & NVL(rsTemp!家庭地址) _
                        & "','" & NVL(rsTemp!家庭电话) & "','" & NVL(rsTemp!家庭地址邮编) & "'," _
                        & zlStr.To_Date(CDate(rsBaby!出生时间)) & "," & NVL(rsTemp!主页ID, 0) & "," & NVL(rsTemp!婴儿) _
                        & ",'" & Trim(txtPhone.Text) & "')"
                    zlDatabase.ExecuteProcedure strSQL, "保存患者信息"
                    End If
                End If
                strSQL = "zl_影像病人信息_修改(" & NVL(rsTemp!病人来源) & "," & mlngOrderID & "," _
                    & rsTemp!病人ID & ",'" & NVL(rsTemp!姓名) & "','" & NVL(rsTemp!性别) & "','" _
                    & NVL(rsTemp!年龄) & "','" & NVL(rsTemp!费别) & "','" & NVL(rsTemp!医疗付款方式) _
                    & "','" & NVL(rsTemp!民族) & "','" & NVL(rsTemp!婚姻状况) & "','" & NVL(rsTemp!职业) _
                    & "','" & NVL(rsTemp!身份证号) & "','" & NVL(rsTemp!家庭地址) _
                    & "','" & NVL(rsTemp!家庭电话) & "','" & NVL(rsTemp!家庭地址邮编) & "'," _
                    & zlStr.To_Date(CDate(rsTemp!出生日期)) & "," & NVL(rsTemp!主页ID, 0) & ",0" _
                    & ",'" & Trim(txtPhone.Text) & "')"
                zlDatabase.ExecuteProcedure strSQL, "保存患者信息"
            End If
        End If
        
        '刷新预约基本信息
        Call RefreshSchInfo(True)
        mblnNewSchedule = False
        
        '记录保存了的医嘱ID
        mstrModifiedOrderID = CStr(mlngOrderID)
        SaveSchedule = True
        
        '自动打印预约单
        If mblnAutoPrint = True Then
            Call PrintSchedule
        End If
    End If
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub DelSchedule(lngOrderID As Long)
'------------------------------------------------
'功能：删除预约
'参数： lngOrderID -- 医嘱ID
'返回：无
'------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo err
    
    strSQL = "Zl_影像预约记录_删除(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure strSQL, "删除检查预约"
    
    Call schTimeTable.RefreshSchedule(mlngSchDeviceID, mschDate, lngOrderID)
    mblnNewSchedule = False
    
    '刷新预约基本信息
    Call RefreshSchInfo(False)
    
    '记录保存了的医嘱ID
    mstrModifiedOrderID = CStr(mlngOrderID)
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshSchedule(blnRefreshBaseInfo As Boolean, blnAutoNew As Boolean)
'------------------------------------------------
'功能：刷新预约，如果是新建预约状态，则每次刷新都新建一个预约标签，如果不是则单纯的刷新
'参数： blnRefreshBaseInfo -- 是否刷新患者基本信息
'       blnAutoNew -- 是否自动新增
'返回：无
'------------------------------------------------
    Dim i As Integer
    
    On Error GoTo err
    
    If mblnNewSchedule = True And blnAutoNew = True Then
        Call schTimeTable.NewSchedule(mlngSchDeviceID, mschDate, mlngOrderID, False)
        '刷新预约基本信息
        Call RefreshSchInfo(blnRefreshBaseInfo)
    Else
        Call schTimeTable.RefreshSchedule(mlngSchDeviceID, mschDate, mlngOrderID)
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ChangeCalendar(dtDate As Date)
'------------------------------------------------
'功能：修改预约日历的日期
'参数：dtDate -- 日历的日期
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    dpCalendar.ClearSelection
    Call dpCalendar.Select(dtDate)
    dpCalendar.EnsureVisibleSelection
    If dpCalendar.Visible = True Then
        dpCalendar.SetFocus
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub PrintSchedule()
'------------------------------------------------
'功能：打印当前预约单
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsReports As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnPrinted As Boolean
    Dim lngUniFmt As Long           '通用报表格式序号
    
    On Error GoTo err
    
    If mblnNewSchedule = True Then
        Call MsgBox("请先保存预约后，再打印预约单。", vbInformation, "检查预约提示")
        Exit Sub
    End If
    
    '打印预约单
    If mlngOrderID <> 0 Then
        '首先检查报表是否只有一个格式
        strSQL = "Select a.ID,a.编号,b.序号,b.说明 From zlreports a,zlrptfmts b Where a.Id=b.报表ID And a.编号=[1] Order By 序号"
        Set rsReports = zlDatabase.OpenSQLRecord(strSQL, "查询预约单报表格式", "ZL1_INSIDE_1290_01")

        If rsReports.EOF = True Then
            Call MsgBox("报表“ZL1_INSIDE_1290_01”不存在，请联系管理员添加此报表。", vbInformation, "检查预约提示")
            Exit Sub
        End If
        '如果有多个格式，按照诊疗项目ID，查找对应的报表格式名称
        If rsReports.RecordCount > 1 Then
            strSQL = "Select a.名称 From 病历文件列表 A, 病历单据应用 B, 病人医嘱记录 C " _
                & " Where c.诊疗项目id = b.诊疗项目id And decode(c.病人来源, 3, 1, c.病人来源) = b.应用场合 " _
                & "And b.病历文件id = a.ID And c.ID = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询病历文件名称", mlngOrderID)
            
            If rsTemp.EOF = False Then
            While rsReports.EOF = False And blnPrinted = False
                If NVL(rsReports!说明) = "通用检查预约单" Then
                    lngUniFmt = rsReports!序号
                End If
                
                If NVL(rsReports!说明) = NVL(rsTemp!名称) Then
                    If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "医嘱ID=" & mlngOrderID, "ReportFormat=" & rsReports!序号, 2) = False Then
                        Call MsgBox("报表“ZL1_INSIDE_1290_01”中，格式为：" & NVL(rsReports!说明) & "的报表，打开不成功，请联系管理员修正此报表。", vbInformation, "检查预约提示")
                    Else
                        '打印完退出循环
                        blnPrinted = True
                    End If
                Else
                    rsReports.MoveNext
                End If
            Wend
            End If
            '如果没有，则查找“通用检查预约单”报表来打印
            If blnPrinted = False Then
                If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "医嘱ID=" & mlngOrderID, "ReportFormat=" & lngUniFmt, 2) = False Then
                    Call MsgBox("报表“ZL1_INSIDE_1290_01”中，格式为：“通用检查预约单”的报表，打开不成功，请联系管理员修正此报表。", vbInformation, "检查预约提示")
                Else
                    blnPrinted = True
                End If
            End If
        Else
            If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "医嘱ID=" & mlngOrderID, 2) = False Then
                Call MsgBox("报表“ZL1_INSIDE_1290_01”打开不成功，请联系管理员修正此报表。", vbInformation, "检查预约提示")
            Else
                blnPrinted = True
            End If
        End If
        
        '写入打印记录
        strSQL = "Zl_影像预约记录_打印(" & mlngOrderID & ")"
        zlDatabase.ExecuteProcedure strSQL, "检查预约单打印"
        
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyPatInfo()
'------------------------------------------------
'功能：修改病人基本信息，打开“修改信息窗口”
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "Select  a.姓名, b.执行过程, b.发送号,b.执行部门id as 执行科室ID" _
        & " From 病人医嘱记录 A, 病人医嘱发送 B " _
        & " Where a.id = b.医嘱id  And a.id = [1]"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "修改信息，查询基本信息", mlngOrderID)
    
    With frmRISRequest
        .mstrPrivs = gstrPrivs
        .mlngModul = glngModul
        .mlngSendNo = rsTemp!发送号
        .mlngAdviceId = mlngOrderID
        .mstrPatientName = NVL(rsTemp!姓名)
        .mintEditMode = IIf(NVL(rsTemp!执行过程, 0) > 1, 3, 1) '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = rsTemp!执行科室ID
        .mstrCur科室 = "科室"
        
        Call frmRISRequest.InitMvar(False)
        .ZlShowMe Me, mstrDefaultPatientType, IIf(gbytFontSize <= 9, 0, 1)
            
    End With
    Call RefreshForm
    Call mfrmParent.RefreshList '刷新父窗口
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshForm()
'------------------------------------------------
'功能：刷新窗口内容
'参数：
'返回：无
'------------------------------------------------
    On Error GoTo err
    
    Call LoadData
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshSchInfo(blnRefreshBaseInfo As Boolean)
'------------------------------------------------
'功能：刷新患者的基本信息
'参数： blnRefreshBaseInfo -- 刷新基本信息
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    lblInfo(4).Caption = ""
    lblInfo(5).Caption = ""
        
    If blnRefreshBaseInfo = True Then
        lblInfo(1).Caption = ""
        lblInfo(2).Caption = ""
        txtPhone.Text = ""
        txtPhone.Locked = True
        txtPhone.ForeColor = vbWindowText
        txtNotice.Text = ""
        txtNotice.Locked = True
        txtNotice.ForeColor = vbWindowText
        mlngPatSource = 0
    End If
        
    If blnRefreshBaseInfo = True Then
        strSQL = "Select A.姓名, A.性别, A.年龄, " _
            & " DECODE(A.病人来源, 2, '住院', 1, '门诊', 4, '体检', '外诊') As 病人来源中文,A.病人来源, " _
            & " B.手机号 ,Nvl(B.家庭地址, B.联系人地址) 地址, A.医嘱内容,Nvl(a.婴儿, 0) As 婴儿 " _
            & " From 病人医嘱记录 A, 病人信息 B " _
            & " Where a.ID = [1] And a.病人ID = b.病人ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约患者信息", mlngOrderID)
        
        If Not rsTemp.EOF Then
            '处理婴儿姓名
            If rsTemp!婴儿 <> 0 Then
                strSQL = "Select A.开嘱时间,Nvl(B.婴儿姓名, A.姓名 || '之子' || Trim(To_Char(B.序号, '9'))) As 婴儿姓名, B.婴儿性别, B.出生时间" & vbNewLine & _
                                 "  From 病人医嘱记录 A, 病人新生儿记录 B " & vbNewLine & _
                                 "  Where a.病人ID = b.病人ID  And b.序号 = [2] And a.ID = [1]"
                            
                Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "提取婴儿信息", mlngOrderID, CLng(rsTemp!婴儿))
                
                lblInfo(0).Caption = "婴儿姓名：" & rsBaby!婴儿姓名 & "   性别：" & rsBaby!婴儿性别
                lblInfo(1).Caption = "出生时间：" & rsBaby!出生时间 & "   来源：" & rsTemp!病人来源中文
            Else
                lblInfo(0).Caption = "姓名：" & rsTemp!姓名 & "   性别：" & rsTemp!性别
                lblInfo(1).Caption = "年龄：" & rsTemp!年龄 & "   来源：" & rsTemp!病人来源中文
            End If
            
            lblInfo(2).Caption = rsTemp!医嘱内容
            txtPhone.Text = NVL(rsTemp!手机号)
            mlngPatSource = rsTemp!病人来源
        Else
            lblInfo(0).Caption = "姓名：         性别："
            lblInfo(1).Caption = "年龄：         来源："
            lblInfo(2).Caption = ""
            txtPhone.Text = ""
            mlngPatSource = 0
        End If
    End If
    
    strSQL = "select 预约设备ID,预约设备名称,预约开始时间,预约结束时间,检查注意 from 影像预约记录 where 医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约记录", mlngOrderID)
    If Not rsTemp.EOF Then
        txtNotice.Text = NVL(rsTemp!检查注意)
        lblInfo(4).Caption = "预约日期: " & Format(rsTemp!预约开始时间, "yyyy-mm-dd")
        lblInfo(5).Caption = "预约时间：" & Format(rsTemp!预约开始时间, "HH:MM:SS") _
            & " - " & Format(rsTemp!预约结束时间, "HH:MM:SS")
        mschDate = Format(rsTemp!预约开始时间, "YYYY-MM-DD")
        mlngSchDeviceID = rsTemp!预约设备ID
        
        '设置预约设备
        lblSchDevice.Caption = rsTemp!预约设备名称
        
        mblnISScheduled = True
    Else
        txtNotice.Text = ""
        lblInfo(4).Caption = "预约日期: "
        lblInfo(5).Caption = "预约时间："
        mblnISScheduled = False
    End If
       
    lblSchDevice.Visible = IIf(cboSchDevice.ListCount > 0, False, True)
    cboSchDevice.Visible = IIf(cboSchDevice.ListCount > 0, True, False)
   
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub NewSchedule()
'------------------------------------------------
'功能：新建预约
'参数：
'返回：无
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "select ID from 影像预约记录 where 医嘱id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询预约记录", mlngOrderID)
    
    If rsTemp.EOF = False Then
        If MsgBox("新建预约之前，将自动删除这个检查原有的预约信息，" & vbCrLf & vbCrLf & "是否确认删除原有的预约信息？", vbYesNo, "检查预约提示") = vbNo Then
            Exit Sub
        Else
            Call DelSchedule(mlngOrderID)
        End If
    End If
    
    mblnNewSchedule = True
    Call RefreshSchedule(False, True)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshCalendar()
'------------------------------------------------
'功能：刷新日历
'参数：
'返回：无
'------------------------------------------------
    
    On Error GoTo err
    
    mstrSchRestDate = RefeshSchRestDay(mlngOrderID, mlngSchDeviceID, dpCalendar.LastVisibleDay)
    
    dpCalendar.RedrawControl
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
