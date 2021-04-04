VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmTransfusion 
   BackColor       =   &H8000000C&
   Caption         =   "门诊输液注射管理"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11670
   Icon            =   "frmTransfusion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picReadyReceive 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   3615
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6480
      Width           =   3615
      Begin zlIDKind.PatiIdentify ptiReadyReceive 
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmTransfusion.frx":6852
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindAppearance=   0
         ShowSortName    =   -1  'True
         ShowPropertySet =   -1  'True
         DefaultCardType =   "就诊卡"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAutoCommCard=   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         NotContainFastKey=   "F1;CTRL+F1;F12;CTRL+F12"
      End
   End
   Begin VB.PictureBox picTmp 
      Height          =   255
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5550
      TabIndex        =   30
      Top             =   105
      Width           =   1905
   End
   Begin VB.Timer tmrAutoReady 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1710
      Top             =   225
   End
   Begin VB.PictureBox picRecord 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   45
      ScaleHeight     =   1035
      ScaleWidth      =   3705
      TabIndex        =   17
      Top             =   5070
      Width           =   3705
      Begin XtremeReportControl.ReportControl rptRecord 
         Height          =   780
         Left            =   60
         TabIndex        =   18
         Top             =   75
         Width           =   3555
         _Version        =   589884
         _ExtentX        =   6271
         _ExtentY        =   1376
         _StockProps     =   0
         BorderStyle     =   2
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   5115
      ScaleHeight     =   585
      ScaleWidth      =   7740
      TabIndex        =   0
      Top             =   465
      Width           =   7740
      Begin VB.Frame fraInfo 
         Height          =   645
         Left            =   15
         TabIndex        =   3
         Top             =   -60
         Width           =   7695
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "卡号:"
            Height          =   180
            Index           =   14
            Left            =   90
            TabIndex        =   20
            Top             =   375
            Width           =   450
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   15
            Left            =   570
            TabIndex        =   19
            Top             =   375
            Width           =   1065
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   13
            Left            =   4350
            TabIndex        =   16
            Top             =   375
            Width           =   90
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "诊断:"
            Height          =   180
            Index           =   12
            Left            =   3855
            TabIndex        =   15
            Top             =   375
            Width           =   585
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   11
            Left            =   2175
            TabIndex        =   14
            Top             =   375
            Width           =   1605
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "科室:"
            Height          =   180
            Index           =   10
            Left            =   1695
            TabIndex        =   13
            Top             =   375
            Width           =   585
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   7
            Left            =   4350
            TabIndex        =   12
            Top             =   165
            Width           =   1290
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "费别:"
            Height          =   180
            Index           =   6
            Left            =   3855
            TabIndex        =   11
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   5
            Left            =   3285
            TabIndex        =   10
            Top             =   165
            Width           =   525
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "年龄:"
            Height          =   180
            Index           =   4
            Left            =   2775
            TabIndex        =   9
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   2175
            TabIndex        =   8
            Top             =   165
            Width           =   540
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "性别:"
            Height          =   180
            Index           =   2
            Left            =   1695
            TabIndex        =   7
            Top             =   165
            Width           =   450
         End
         Begin VB.Label lblinfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   1
            Left            =   570
            TabIndex        =   6
            Top             =   165
            Width           =   1020
         End
         Begin VB.Label lblinfo 
            BackStyle       =   0  'Transparent
            Caption         =   "姓名:"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   4
            Top             =   165
            Width           =   450
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7455
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTransfusion.frx":693C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15716
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
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
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   6240
      Left            =   5130
      TabIndex        =   2
      Top             =   1080
      Width           =   6600
      _Version        =   589884
      _ExtentX        =   11642
      _ExtentY        =   11007
      _StockProps     =   64
   End
   Begin VB.PictureBox picLeft 
      BorderStyle     =   0  'None
      Height          =   4485
      Left            =   165
      ScaleHeight     =   4485
      ScaleWidth      =   4890
      TabIndex        =   1
      Top             =   480
      Width           =   4890
      Begin VB.PictureBox picQueue0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   3780
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3555
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue0 
            Height          =   3270
            Left            =   195
            TabIndex        =   42
            Top             =   270
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.Frame fraWhere 
         Height          =   1300
         Left            =   45
         TabIndex        =   33
         Top             =   -45
         Width           =   4815
         Begin zlIDKind.IDKindNew idkSelect 
            Height          =   270
            Left            =   120
            TabIndex        =   39
            Top             =   900
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   476
            IDKindStr       =   "姓|姓名或就诊卡|0|0|0|0|0|;医|医保号|0|0|0|0|0|;身|身份证号|0|0|0|0|0|;IC|IC卡号|1|0|0|0|0|;门|门诊号|0|0|0|0|0|"
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
            ShowPropertySet =   -1  'True
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
            BackColor       =   -2147483633
            SaveRegType     =   4
         End
         Begin VB.ComboBox cboDate 
            Height          =   300
            ItemData        =   "frmTransfusion.frx":71CE
            Left            =   795
            List            =   "frmTransfusion.frx":71E1
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   540
            Width           =   2600
         End
         Begin VB.TextBox txtInfo 
            Height          =   270
            Left            =   1095
            TabIndex        =   40
            Top             =   900
            Width           =   2580
         End
         Begin VB.CommandButton cmdOk 
            Height          =   270
            Left            =   3405
            Picture         =   "frmTransfusion.frx":7213
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   540
            Width           =   315
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   795
            TabIndex        =   35
            Text            =   "cboDept"
            Top             =   195
            Width           =   2910
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "时间(&T)"
            Height          =   180
            Left            =   135
            TabIndex        =   36
            Top             =   600
            Width           =   630
         End
         Begin VB.Label lblB 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "科室(&D)"
            Height          =   180
            Left            =   135
            TabIndex        =   34
            Top             =   255
            Width           =   630
         End
      End
      Begin VB.PictureBox picQueue7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   1080
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3585
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue7 
            Height          =   3270
            Left            =   150
            TabIndex        =   32
            Top             =   255
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.PictureBox picQueue6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   3285
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3255
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue6 
            Height          =   3270
            Left            =   135
            TabIndex        =   29
            Top             =   150
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
      End
      Begin VB.PictureBox picQueue5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   2295
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2865
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue5 
            Height          =   3270
            Left            =   210
            TabIndex        =   27
            Top             =   675
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
         Begin VB.TextBox txtNo5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   45
            Top             =   225
            Width           =   1995
         End
         Begin VB.Label lblNo5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "填入挂号单回车完成穿刺"
            Height          =   180
            Left            =   15
            TabIndex        =   46
            Top             =   0
            Width           =   1980
         End
      End
      Begin VB.PictureBox picQueue1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3120
         Left            =   855
         ScaleHeight     =   3090
         ScaleWidth      =   2835
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2565
         Width           =   2865
         Begin XtremeReportControl.ReportControl rptQueue1 
            Height          =   3270
            Left            =   150
            TabIndex        =   25
            Top             =   630
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
         Begin VB.TextBox txtNo1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   150
            TabIndex        =   43
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label lblNo1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "填入挂号单回车完成配液"
            Height          =   180
            Left            =   165
            TabIndex        =   44
            Top             =   75
            Width           =   1980
         End
      End
      Begin VB.PictureBox picQueueAll 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   225
         ScaleHeight     =   2775
         ScaleWidth      =   3630
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2430
         Width           =   3660
         Begin XtremeReportControl.ReportControl rptPati 
            Height          =   3270
            Left            =   105
            TabIndex        =   23
            Top             =   420
            Width           =   3525
            _Version        =   589884
            _ExtentX        =   6218
            _ExtentY        =   5768
            _StockProps     =   0
            BorderStyle     =   2
            AutoColumnSizing=   0   'False
         End
      End
      Begin XtremeSuiteControls.TabControl tbcList 
         Height          =   960
         Left            =   225
         TabIndex        =   21
         Top             =   1335
         Width           =   3630
         _Version        =   589884
         _ExtentX        =   6403
         _ExtentY        =   1693
         _StockProps     =   64
      End
      Begin VB.Timer timRefresh 
         Interval        =   1000
         Left            =   3045
         Top             =   2745
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   2010
         Top             =   2625
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":DA65
               Key             =   "未执行"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":DFFF
               Key             =   "已执行"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":E599
               Key             =   "拒绝执行"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":EB33
               Key             =   "正在执行"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":F0CD
               Key             =   "已报到"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTransfusion.frx":F667
               Key             =   "Calling"
            EndProperty
         EndProperty
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmTransfusion.frx":15EC9
      Left            =   675
      Top             =   135
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTransfusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private patiList As New cPatients '病人列表类
Private ObjOutNurse As New OutNurses '门诊护士列表
Private mobjPopupInfo As CommandBar

Private mstr挂号单 As String  '当前用户唯一标识
Private mstrPrivs As String
Private mlngModul As Long

Private mlngPreDept As Long '上次科室ID
Private mstr座位 As String
Private mDateBegin As Date '开始时间
Private mdateEnd As Date  '结束时间
Private mintRefresh As Integer
Public mbln皮试验证 As Boolean
'子窗体
Private WithEvents mclsSeating As clsDockSeating
Attribute mclsSeating.VB_VarHelpID = -1
Private mfrmLeaveMedi As frmLeaveMedi '药品寄存
Private mfrmRecord As frmRecord '待执行项目
Public mobjRecord As ExecRecord

Private mcolSubForm As Collection 'subTab子窗体集合
'病人项目列表
Private Enum rptCOL
    rptCOL_执行分类 = 0
    rptCOL_接单时间 = 1
    rptCOL_配药人 = 2
    rptCOL_接单人 = 3

    rptCOL_耗时 = 4
    rptCOL_滴系数 = 5
    rptCOL_组数 = 6
    rptCOL_流水号 = 7
End Enum

'Private mfrmActive As Form
Private mintRow As Integer    '当前行,用于叫号
Private Type SiblingRow
    PrivRowIndex As Integer
    PrivRow挂号单 As String
    PrivRow状态 As String

    curRow挂号单 As String
    curRow状态 As String
    curRowIndex As Integer

    nextRow挂号单 As String
    nextRow状态 As String
    nextRowIndex As Integer
End Type

Private mintPatirow As Integer '刷新时,重新定位
Private mintRecordRow As Integer '刷新时,重新定位

Private mintFindType As Integer '查找类型 0-就诊卡,1-门诊号,2-挂号单,3-姓名,4-身份证,5-IC卡

'Private mstrIDCard As String '最近自动刷出来的身份证号
'Private WithEvents mobjIDCard As clsIDCard '身份证对象
'Private mobjICCard As Object 'IC卡对象
Private mstrQueueTab As String  '当前队列页面
Private mintLastFind As Integer     '查找次数

Private mblnLiquid  As Boolean  '是否有配液流程
'Private mblnPuncture As Boolean '是否有穿刺流程  简单/标准列表中有本科室即认为有穿刺流程
'Private mblnCall    As Boolean  '是否有呼叫流程

Private mblnVisits As Boolean   '是否有巡视流程
    
Private mobjSquareCard   As Object   '一卡通部件 add by 2011-08-23
Private mstrSquareCards As String    '一卡通卡名
Private mintCards As Integer         '一卡能卡的数量
Private mstrPatiKey As String        '过滤病人信息类型
Private mblnReadCard As Boolean
Private mintPatiIdentify As String
Private mfrmTimeCall As Form      '排队叫号的轮循窗体

Private Const MLNG_INFO As Long = 100000
Private Const MSTR_MODE As String = "挂|挂号单|0;就|就诊卡|1;门|门诊号|0;姓|姓名|0;身|身份证号|0;IC|IC卡|1"

Private Sub ShowPage()
    '根据系统参数显示配液页面
    Dim i As Integer
    Dim strPara As String
    
    '85046
    'mblnLiquid = GetDeptInListPara("无线输液_配液科室列表", mlngPreDept)
    strPara = zlDatabase.GetPara("待配液科室列表", glngSys, mlngModul, "")
    mblnLiquid = InStr("," & strPara & ",", "," & mlngPreDept & ",") > 0
    
'    mblnPuncture = True '穿刺功能必须有，好填写开始时间，开始操作员。
    ' GetDeptInListPara("无线输液_标准穿刺列表", mlngPreDept) Or GetDeptInListPara("无线输液_简单穿刺列表", mlngPreDept)
    
    '85046 问题取消该参数的控制
    'mblnCall = GetDeptInListPara("无线输液_呼叫科室列表", mlngPreDept)
    
    With Me.tbcList
        For i = 1 To .ItemCount - 1
            If (Not mblnLiquid And .Item(i).Tag = "待配液") Or .Item(i).Tag = "待执行" Then
                .Item(i).Visible = False
            Else
                .Item(i).Visible = True
            End If
        Next
    End With
End Sub

Private Sub cboDate_Click()
    mdateEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    If Trim(txtInfo.Text) = "" Then
        Select Case cboDate.ListIndex
            Case 1              '昨天
                mDateBegin = Format(mdateEnd - 1, "yyyy-MM-dd 00:00:00")
            Case 2              '近三天
                mDateBegin = Format(mdateEnd - 2, "yyyy-MM-dd 00:00:00")
            Case 3              '近一周
                mDateBegin = Format(mdateEnd - 6, "yyyy-MM-dd 00:00:00")
            Case 4              '近十天
                mDateBegin = Format(mdateEnd - 9, "yyyy-MM-dd 00:00:00")
            Case Else           '当天
                mDateBegin = Format(mdateEnd, "yyyy-MM-dd 00:00:00")
        End Select
    Else
        mDateBegin = Format(mdateEnd - 364, "yyyy-MM-dd 00:00:00")
    End If
End Sub

Private Sub cboDept_Click()
    If cboDept.ListCount <= 0 Then Exit Sub
    If cboDept.ItemData(cboDept.ListIndex) = mlngPreDept Then Exit Sub
    mlngPreDept = cboDept.ItemData(cboDept.ListIndex)
    
    Call ShowPage
    Call ObjOutNurse.getOutNurse(mlngPreDept)  '初始化本科室护士列表
    '初始化patients类
    Call ShowLblInfo("")
    Call ShowReport
    mstr挂号单 = ""
    mintRow = 0
    ShowPatiList
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If cboDept.ListIndex <> -1 Then cboDept.Tag = cboDept.ListIndex

    '支持录入科室简码、名称、编码
    If KeyAscii = vbKeyReturn Then
        If Trim(cboDept.Text) = "" Then Exit Sub
        
        Dim intIndex As Integer
        Dim strText As String, strSQL As String
        Dim rsTmp As ADODB.Recordset
        Dim vRect As RECT
        Dim blnCanel As Boolean
        
        KeyAscii = 0
        intIndex = cboDept.ListIndex
        strText = UCase(cboDept.Text) & "%"
        If Val(zlDatabase.GetPara("输入匹配", , "0")) = 1 Then
            strText = UCase(cboDept.Text) & "%"
        Else
            strText = "%" & UCase(cboDept.Text) & "%"
        End If
        
        If InStr(mstrPrivs, ";所有科室;") > 0 Then
            strSQL = "Select /*+ Rule*/ Distinct a.Id, a.编码, a.名称 " & vbCr & _
                     "From 部门表 A, 部门性质说明 B " & vbCr & _
                     "Where b.部门id = a.Id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And b.服务对象 In (1, 3) " & vbCr & _
                     "  And b.工作性质 In ('治疗', '临床') And (a.站点 = [2] Or a.站点 Is Null) " & vbCr & _
                     "  And (A.编码 Like [3] Or A.简码 Like [3] Or A.名称 Like [3]) " & vbCr & _
                     "Order By a.编码 "
        Else
            strSQL = "Select /*+ Rule*/ Distinct a.Id, a.编码, a.名称 " & vbCr & _
                     "From 部门表 A, 部门性质说明 B, 部门人员 C " & vbCr & _
                     "Where b.部门id = a.Id And a.Id = c.部门id And c.人员id = [1] " & vbCr & _
                     "  And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) And b.服务对象 In (1, 3) " & vbCr & _
                     "  And b.工作性质 In ('治疗', '临床') And (a.站点 = [2] Or a.站点 Is Null) " & vbCr & _
                     "  And (A.编码 Like [3] Or A.简码 Like [3] Or A.名称 Like [3]) " & vbCr & _
                     "Order By a.编码 "
        End If
        On Error GoTo errHandle
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取科室信息", UserInfo.ID, zl9ComLib.gstrNodeNo, strText)
        If Not rsTmp.EOF Then
            If rsTmp.RecordCount = 1 Then
                Call FindCboIndex(cboDept, rsTmp!ID)
            Else
                rsTmp.Close
                
                vRect = zlControl.GetControlRect(cboDept.hwnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "指定科室", False, "", "选择科室", False, False, True, _
                                         vRect.Left, vRect.Top, cboDept.Height, blnCanel, True, True, _
                                         UserInfo.ID, zl9ComLib.gstrNodeNo, strText)
                If blnCanel = False Then
                    Call FindCboIndex(cboDept, rsTmp!ID)
                    rsTmp.Close
                Else
                    cboDept.ListIndex = Val(cboDept.Tag)
                End If
                
            End If
        Else
            cboDept.ListIndex = Val(cboDept.Tag)
            rsTmp.Close
        End If
        
        Call zlCommFun.PressKey(vbKeyTab)
        
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim strStat As String
    Dim objPati As cPatient
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.STYLE = IIf(objControl.STYLE = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Jump '跳转
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
'    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 99 '查找方式
'        mintFindType = Val(Right(Control.ID, 2)) - 1
'
'        Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_View_FindType)
'        objControl.Caption = "↓按" & Split(Control.Caption, "(")(0) & "查找"
'
'        cbsMain.RecalcLayout
'        txtFind.Tag = Control.Parameter
'        txtFind.Text = ""
'        txtFind.SetFocus
        
'    Case MLNG_INFO + 1 To MLNG_INFO + 99     '指定提取病人信息
'        lblBill.Caption = Split(Control.Caption, "(")(0)
'        lblBill.Caption = lblBill.Caption & "↓"
'        lblBill.Tag = Control.ID - MLNG_INFO
        
        
    Case conMenu_View_Find '查找
        If Me.ActiveControl Is txtFind Then
            txtFind.SetFocus '有时需要定位一下
            If txtFind.Text <> "" Then
                Call ExecuteFindPati
            End If
        Else
            txtFind.SetFocus
        End If
        
'    Case conMenu_View_FindNext '查找下一个
'        If txtFind.Text = "" And mstrIDCard = "" Then
'            txtFind.SetFocus
'        Else
'            Call ExecuteFindPati(True, IIf(txtFind.Text = "", mstrIDCard, ""))
'        End If
'    Case conMenu_View_ReadIC '读IC卡
'        If Not mobjICCard Is Nothing Then
'            txtFind.Text = mobjICCard.Read_Card(Me)
'            If txtFind.Text <> "" Then Call ExecuteFindPati
'        End If
    Case conMenu_View_Refresh '刷新
        Call cmdOk_Click
        
    Case conMenu_View_Expend_CurCollapse '折叠当前组
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                rptPati.SelectedRows(0).Expanded = False
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    rptPati.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '因折叠定位到分组上,不会自动激活该事件
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_CurExpend '展开当前组
        If rptPati.SelectedRows.Count > 0 Then
            rptPati.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '折叠所有组
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '因折叠定位到分组上,不会自动激活该事件
        Call rptPati_SelectionChanged
    Case conMenu_View_Expend_AllExpend '展开所有组
        For Each objRow In rptPati.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    
    Case conMenu_Manage_ThingAdd
        '接单
        Call thingAdd
    Case conMenu_File_Parameter
        '参数设置
         Call ParameterSetup
    Case conMenu_File_RoomSet
        '穿刺台设置
        frmPunctureDeskSet.ShowMe mlngPreDept
    Case conMenu_Manage_Call
        '呼叫
        Call Calling(2)
        If mstr挂号单 <> "" Then Call CallOnePlay(mstr挂号单)
'    Case conMenu_Manage_CallNext
'        '下一个
'        Call Calling(1)
'    Case conMenu_Manage_CallPrevious
'        '上一个
'        Call Calling(-1)
    Case conMenu_Manage_Up
        '上移
        Call rptQueueMove(-1)
    Case conMenu_Manage_Down
        '下移
        Call rptQueueMove(1)
        
    Case conMenu_Manage_Discard
        '弃号
        If UpdateState("2-弃号") Then
            Set objPati = patiList.Item(mstr挂号单)
            SaveOperLog mlngPreDept, objPati, QUEUE, "弃号操作"
        End If
    Case conMenu_Manage_Recall
        '召回
        If UpdateState("7-执行中") Then
            Set objPati = patiList.Item(mstr挂号单)
            SaveOperLog mlngPreDept, objPati, QUEUE, "召回操作"
        End If
    Case conMenu_Manage_Untread
        '退号
        If UpdateState("3-退号") Then
            Set objPati = patiList.Item(mstr挂号单)
            SaveOperLog mlngPreDept, objPati, QUEUE, "退号操作"
        End If
    Case conMenu_Manage_TagEnd
        '标为全部
        If UpdateState("4-结束") Then
            Set objPati = patiList.Item(mstr挂号单)
            SaveOperLog mlngPreDept, objPati, QUEUE, "结束操作"
        End If
    Case conMenu_Edit_Transf_Liquid
        '配液
        Call LiquidAndPlay
    Case conMenu_Edit_Transf_Puncture
        '穿刺
        Call Puncture
    Case conMenu_Edit_Bed_Modify
        '调整状态
        strStat = patiList.Item(mstr挂号单).排队状态
        If frmChangeStat.ShowMe(strStat, mblnLiquid) Then
            If UpdateState(strStat) Then
                Set objPati = patiList.Item(mstr挂号单)
                SaveOperLog mlngPreDept, objPati, QUEUE, "调整状态为" & strStat
            End If
        End If
    Case conMenu_Queue_Setup    '呼叫参数设置
        Call QueueSetup(Me)
    Case conMenu_View_Show      '查看日志
        Call frmTransfusionLog.ShowMe(mlngPreDept, mstr挂号单)
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Exit '退出
        Unload Me
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If mstr挂号单 <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "病人ID=" & patiList.Item(mstr挂号单).病人ID)
            Else
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
            End If
        ElseIf Me.tbcSub.Selected.Tag = "座位管理" Then
            Call mclsSeating.zlExecuteCommandBars(Control)
        ElseIf tbcSub.Selected.Tag = "执行项目" Then
            Call mfrmRecord.zlExecuteCommandBars(Control)
        ElseIf tbcSub.Selected.Tag = "药品寄存" Then
            Call mfrmLeaveMedi.zlExecuteCommandBars(Control, Me)
        End If
    End Select
    
    '刷新TabControl
    On Error Resume Next
    Select Case Control.ID
        Case conMenu_Edit_Transf_Liquid, conMenu_Edit_Transf_Puncture, conMenu_Edit_Bed_Modify, _
            conMenu_Manage_Discard, conMenu_Manage_Recall, conMenu_Manage_Untread, conMenu_Manage_TagEnd
            If Not tbcList.Selected Is Nothing Then Call tbcList_SelectedChanged(tbcList.Selected)
    End Select
    Err.Clear
End Sub
Private Sub LiquidAndPlay()
    '配液并呼叫
    Dim strStat As String, strErr As String, i As Integer
    Dim objPati As cPatient
    
    strStat = Liquid(mlngPreDept, mstr挂号单, patiList, strErr)
    If strErr <> "" Then
        MsgBox strErr, vbInformation, Me.Caption
        Exit Sub
    End If
    If strStat <> "" Then
        If Not mobjRecord Is Nothing Then
            For i = 1 To mobjRecord.Count
                Call mobjRecord.Item(i).SaveDispenseUser(1, zlDatabase.Currentdate, UserInfo.姓名)
            Next
            
            Call ShowLblInfo(mstr挂号单)
        End If
        
        Set objPati = patiList.Item(mstr挂号单)
        If UpdateState(strStat) Then
            SaveOperLog mlngPreDept, objPati, QUEUE, "配液后调整状态为" & strStat
        Else
            SaveOperLog mlngPreDept, objPati, QUEUE, "配液后未调整状态"
        End If
    End If
    If strStat = "5-待穿刺" Then
        Call CallPlay(mstr挂号单)
    End If
End Sub
Private Sub Puncture()
    '穿刺操作
    Dim strSQL As String, i As Integer, Y As Integer
    Dim dateS As Date, dateE As Date, blnExitFor As Boolean, strGroupKey As String
    Dim intOneOrTow As Integer, rsTmp As ADODB.Recordset, lngTaiID As Long, strNextNo As String
    Dim objPati As cPatient
    
    On Error GoTo hErr
    
    If mstr挂号单 <> "" Then
        Set objPati = patiList.Item(mstr挂号单)
        If objPati Is Nothing Then
            MsgBox "当前病人无找到，拒绝穿刺！", vbInformation, gstrSysName
            Exit Sub
        End If
                 
        strSQL = "ZL_门诊穿刺台_Puncture(" & mlngPreDept & "," & objPati.病人ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        If UpdateState("7-执行中") Then
            SaveOperLog mlngPreDept, objPati, QUEUE, "穿刺后调整状态为7-执行中"
        Else
            SaveOperLog mlngPreDept, objPati, QUEUE, "穿刺后未调整状态"
        End If
        
        '准备呼叫待穿病人。通过病人科室、穿刺台，找到“待穿病人ID”，再找到病人的挂号单或主页ID
        strSQL = "Select a.病人id, a.挂号单, a.主页id " & vbNewLine & _
                 "From 排队记录 A, 门诊穿刺台 B " & vbNewLine & _
                 "Where a.科室id = b.科室id And a.病人id = b.待穿病人id And a.科室id = [1] And a.穿刺台 = [2] and a.状态 = 5 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询准备呼叫的等待病人", mlngPreDept, objPati.穿刺台)
        If rsTmp.EOF = False Then
            If Val(zlCommFun.NVL(rsTmp!主页id)) <= 0 Then
                '门诊
                strNextNo = zlCommFun.NVL(rsTmp!挂号单)
            Else
                '门诊留观
                strNextNo = zlStr.FormatString("[1]_[2]", rsTmp!病人ID, rsTmp!主页id)
            End If
        Else
            strNextNo = ""
        End If
        rsTmp.Close
        
        '呼叫准备穿刺的病人
        If strNextNo <> "" Then
            Call CallPlay(strNextNo)
        End If
        
        dateS = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
        dateE = Format(dateS, "yyyy-MM-dd 23:59:59")
        If Not mobjRecord Is Nothing Then
            blnExitFor = False
            For i = 1 To mobjRecord.Count
                For Y = 1 To mobjRecord.Item(i).Count
                     If mobjRecord.Item(i).Item(Y).执行分类 = "1-输液" And _
                        mobjRecord.Item(i).Item(Y).组次 = 1 And _
                        mobjRecord.Item(i).执行时间 >= dateS And mobjRecord.Item(i).执行时间 <= dateE And _
                        mobjRecord.Item(i).Item(Y).执行人 = "" Then
                            '今天第一组药，没有填执行人，则填
                            dateS = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
                            strGroupKey = mobjRecord.Item(i).Item(Y).执行医嘱ID & "_" & mobjRecord.Item(i).Item(Y).发送号
                            Call mobjRecord.Item(i).ExecStart(1, strGroupKey, dateS, UserInfo.姓名)
                            
                            '
                            Call ExecComplt(CStr(mobjRecord.Item(i).流水号), strGroupKey)
                            blnExitFor = True
                            Exit For
                     End If
                Next
                If blnExitFor Then Exit For
            Next
        End If
    Else
        MsgBox "请选择一条记录后再执行此操作!", vbQuestion, Me.Caption
    End If
    Exit Sub
    
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CallPlay(ByVal strNO As String)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    '--- 顺呼
     
'    If mblnCall Then Exit Sub                       '本科室不在呼叫列表中
    Call QueueCall(strNO, mlngPreDept, patiList.Item(strNO))
    
End Sub

Private Sub CallOnePlay(ByVal strNO As String)
    '单独的呼叫和显示
    
    '取消本逻辑，因为无论无线输液启用或不启用，台式机端都可以操作呼叫指定病人
    'If Not mblnCall Then Exit Sub                       '本科室不在呼叫列表中
    
    '0-治疗类的医嘱不呼叫窗口内容
    Dim lngNo As Long
    Dim bln治疗 As Boolean
    
    lngNo = Get流水号()
    If lngNo > 0 Then
        If Not mobjRecord.Item(CStr(lngNo)) Is Nothing Then
            bln治疗 = (mobjRecord.Item(CStr(lngNo)).执行分类 Like "0*")
            If bln治疗 Then
                '治疗类的医嘱不呼叫穿刺台窗口内容
                Call QueueOnePlay(strNO, "请、" & patiList.Item(strNO).姓名 & "、、来治疗", lngNo)
            End If
        End If
    End If
    If bln治疗 = False Then
        Call QueueOnePlay(strNO, "请、" & patiList.Item(strNO).姓名 & "、到、" & patiList.Item(strNO).穿刺台 & "号窗口输液", lngNo)
    End If
    
End Sub

Private Sub thingAdd(Optional ByVal bytType As Byte = 0)
'功能：接单过程
'参数：
'  bytType： 0-点击接单按钮方式；1-自动调用接单（待接单病人）

    Dim strJZK As String, strName As String
    Dim lngDeptID As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    '检查当前科室穿刺台设置
    lngDeptID = cboDept.ItemData(cboDept.ListIndex)
    strSQL = "Select Count(1) Rec from 门诊穿刺台 where 科室id = [1] and 有效 = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "门诊穿刺台", lngDeptID)
    lngDeptID = rsTemp!Rec
    rsTemp.Close
    If lngDeptID <= 0 Then
        MsgBox "当前科室未设置穿刺台或无有效的穿刺台！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '接单
    strJZK = lblinfo(15).Caption    '卡号
    strName = lblinfo(1).Caption    '姓名
    
    If bytType = 1 Then
        Call frmReady.ShowIncepBill(bytType, mlngPreDept, cboDept.List(cboDept.ListIndex), mstr座位, mDateBegin, mdateEnd, patiList, _
                    ObjOutNurse, Me, , , , Me.ptiReadyReceive)
    Else
        Call frmReady.ShowIncepBill(bytType, mlngPreDept, cboDept.List(cboDept.ListIndex), mstr座位, mDateBegin, mdateEnd, patiList, _
                    ObjOutNurse, Me, Me.txtInfo.Text, strJZK, strName)
    End If
        
    '数据已变化，刷新显示        '更新座位显示
    mlngPreDept = -1
    mdateEnd = CDate(0)
    Call cboDept_Click

    
'    If mstr挂号单 = "" Then Exit Sub
'    With rptPati.SelectedRows(0)
'        If Not .GroupRow Then
'            If mstr挂号单 <> "" And InStr("1-待配液,0-未接单", .Record(col_排队状态).Value) > 0 Then
'                '-- 科室,科室名称,座位,开始日期,结束日期,patient类,窗体
'                If frmReady.InceptBill(mlngPreDept, cboDept.List(cboDept.ListIndex), mstr座位, mDateBegin, mdateEnd, _
'                                       patiList.Item(mstr挂号单), patiList.mSeatings, ObjOutNurse, Me) Then
'                '刷新座位
'                    mlngPreDept = -1
'                    mdateEnd = CDate(0)
'                    Call cboDept_Click
'                    Call rptPati_SelectionChanged
'                End If
'            End If
'        End If
'    End With

End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    If CommandBar.Parent Is Nothing Then Exit Sub

    '右键弹出菜单
    Select Case CommandBar.Parent.ID
    Case conMenu_View_FindType '查找方式
'        With CommandBar.Controls
'            If .Count = 0 Then
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 1, "就诊卡(&1)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 2, "门诊号(&2)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 3, "单据号(&3)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 4, "姓  名(&4)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 5, "身份证(&5)"
'                .Add xtpControlButton, conMenu_View_FindType * 100# + 6, "ＩＣ卡(&6)"
'            End If
'        End With
        
    Case Else
        If tbcSub.Selected.Tag = "座位管理" Then
           Call mclsSeating.zlPopupCommandBars(CommandBar)
        End If
        
        If tbcSub.Selected.Tag = "执行项目" Then
            Call mfrmRecord.zlPopupCommandBars(CommandBar)
        End If
        
        If tbcSub.Selected.Tag = "药品寄存" Then
            Call mfrmLeaveMedi.zlPopupCommandBars(CommandBar)
        End If
        
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.picInfo
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
    End With

    With Me.tbcSub
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = picInfo.Top + picInfo.Height
        .Height = lngBottom - .Top - stbThis.Height
    End With

End Sub

Private Function GetRowState(objRpt As ReportControl, blnQueue As Boolean) As SiblingRow
    '取指定RPT控件的相临行状态
    'blnQueue = 取呼叫列 ,否则取状态列
    Dim intCurRow As Integer
    If blnQueue Then
        intCurRow = mintRow
        If intCurRow > 0 Then
            GetRowState = SiblingRowState(objRpt, intCurRow)
        End If
    Else
        If objRpt.SelectedRows.Count > 0 Then
            intCurRow = objRpt.SelectedRows(0).Index
            If intCurRow >= 0 Then
                GetRowState = SiblingRowState(objRpt, intCurRow)
            End If
        End If
    End If
End Function
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Dim TCurRowState As SiblingRow  '移动用
    Dim TQueueRowState As SiblingRow
    Dim blnEnabled As Boolean
    Dim lng流水号 As Long, intItem As Integer
    
    If tbcList.Selected.Tag = "未接单" Then
        TCurRowState = GetRowState(rptQueue0, False)
        TQueueRowState = GetRowState(rptQueue0, True)
    ElseIf tbcList.Selected.Tag = "待配液" Then
        '取当前行及相邻行状态
        TCurRowState = GetRowState(rptQueue1, False)
        TQueueRowState = GetRowState(rptQueue1, True)
    ElseIf tbcList.Selected.Tag = "待穿刺" Then
        TCurRowState = GetRowState(rptQueue5, False)
        TQueueRowState = GetRowState(rptQueue5, True)
    ElseIf tbcList.Selected.Tag = "待执行" Then
        TCurRowState = GetRowState(rptQueue6, False)
        TQueueRowState = GetRowState(rptQueue6, True)
    ElseIf tbcList.Selected.Tag = "执行中" Then
        TCurRowState = GetRowState(rptQueue7, False)
        TQueueRowState = GetRowState(rptQueue7, True)
    ElseIf tbcList.Selected.Tag = "已结束" Then
        TCurRowState = GetRowState(rptPati, False)
        TQueueRowState = GetRowState(rptPati, True)
    End If
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).STYLE = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Expend_CurExpend '展开当前组
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptPati.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '折叠当前组
        blnEnabled = False
        If rptPati.SelectedRows.Count > 0 Then
            If rptPati.SelectedRows(0).GroupRow Then
                blnEnabled = rptPati.SelectedRows(0).Expanded
            ElseIf Not rptPati.SelectedRows(0).ParentRow Is Nothing Then
                If rptPati.SelectedRows(0).ParentRow.GroupRow Then
                    blnEnabled = rptPati.SelectedRows(0).ParentRow.Expanded
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend '折叠/展开组
        Control.Enabled = rptPati.GroupsOrder.Count > 0 And rptPati.Rows.Count > 0
'    Case conMenu_View_FindType '查找方式
'        If Control.Parent Is cbsMain.ActiveMenuBar Then
'            If mintFindType <= 5 Then
'                Control.Caption = "按" & Decode(mintFindType, 0, "就诊卡", 1, "门诊号", 2, "挂号单", 3, "姓名", 4, "身份证", 5, "ＩＣ卡") & "查找"
'            Else
'                Control.Caption = ""
'            End If
'        End If
'        txtFind.PasswordChar = IIf(mintFindType = 0 And gblnCardHide, "*", "")
'    Case conMenu_View_FindType * 100# + 1 To conMenu_View_FindType * 100# + 99
'        '查找方式
'        Control.Checked = Val(Right(Control.ID, 2)) - 1 = mintFindType
'    Case conMenu_View_ReadIC '读IC卡
'        Control.Visible = mintFindType = 5
    Case conMenu_Manage_ThingAdd    '接单
        Control.Enabled = InStr(mstrPrivs, ";" & "医嘱接单" & ";") <> 0 'And InStr("3-退号", TCurRowState.curRow状态) <= 0 And TCurRowState.curRow状态 <> ""
    Case conMenu_Edit_Transf_Liquid
        '配液
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And InStr("1-待配液,0-未接单", TCurRowState.curRow状态) > 0
    Case conMenu_Manage_Call, conMenu_Edit_Transf_Puncture
        '呼叫，穿刺
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And (TCurRowState.curRow状态 = "5-待穿刺" Or Val(TCurRowState.curRow状态) = 7) And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
        
'    Case conMenu_Manage_Call
'        '呼叫
'        Control.Enabled = TQueueRowState.curRow挂号单 <> "" And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
'    Case conMenu_Manage_CallNext
'        '下一个
'        Control.Enabled = TQueueRowState.nextRow挂号单 <> "" And TQueueRowState.nextRow状态 = "1-待配液" And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
'    Case conMenu_Manage_CallPrevious
'        '上一个
'        Control.Enabled = TQueueRowState.PrivRow挂号单 <> "" And TQueueRowState.PrivRow状态 = "1-待配液" And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
    Case conMenu_Manage_Reset
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
    Case conMenu_Manage_Up
        '上移
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And TCurRowState.PrivRow状态 = "5-待穿刺" And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
    Case conMenu_Manage_Down
        '下移
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And TCurRowState.nextRow状态 = "5-待穿刺" And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
    Case conMenu_Manage_Discard
        '弃号
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And InStr(",2,3,4,", "," & Val(TCurRowState.curRow状态) & ",") <= 0 And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
        If Control.Enabled Then
            lng流水号 = Get流水号()
            If lng流水号 > 0 Then
                If Not mobjRecord.Item(CStr(lng流水号)) Is Nothing Then
                    For intItem = 1 To mobjRecord.Item(CStr(lng流水号)).Count
                        If mobjRecord.Item(CStr(lng流水号)).Item(intItem).执行状态 = 1 Then
                            Control.Enabled = False
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    Case conMenu_Edit_Bed_Modify
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And InStr(mstrPrivs, ";调整状态;") > 0
    Case conMenu_Manage_Recall
        '召回
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And InStr("2-弃号,4-结束", TCurRowState.curRow状态) > 0 And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
    Case conMenu_Manage_TagEnd
        '标为结束
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And InStr(",2,3,4,", "," & Val(TCurRowState.curRow状态) & ",") <= 0 And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
    Case conMenu_Manage_Untread
        '退号
        Control.Enabled = TCurRowState.curRow挂号单 <> "" And InStr(",2,3,4,", "," & Val(TCurRowState.curRow状态) & ",") <= 0 And InStr(mstrPrivs, ";" & "排队管理" & ";") <> 0
'        TCurRowState.
        If Control.Enabled Then
            lng流水号 = Get流水号()
            If lng流水号 > 0 Then
                If Not mobjRecord.Item(CStr(lng流水号)) Is Nothing Then
                For intItem = 1 To mobjRecord.Item(CStr(lng流水号)).Count
                    If mobjRecord.Item(CStr(lng流水号)).Item(intItem).执行状态 = 1 Then
                        Control.Enabled = False
                        Exit For
                    End If
                Next
                End If
            End If
        End If
    Case conMenu_File_RoomSet   '穿刺台管理
        Control.Enabled = InStr(mstrPrivs, ";座位管理;") > 0
        
'    Case MLNG_INFO + 1 To MLNG_INFO + 99     '指定提取病人信息
'        Control.Checked = (Val(lblBill.Tag) = Control.ID - MLNG_INFO)
    Case Else
        If Me.tbcSub.Selected.Tag = "座位管理" Then
            Call mclsSeating.zlUpdateCommandBars(Control)
        End If
        If Me.tbcSub.Selected.Tag = "执行项目" Then
            Call mfrmRecord.zlUpdateCommandBars(Control)
        End If
        If Me.tbcSub.Selected.Tag = "药品寄存" Then
            Call mfrmLeaveMedi.zlUpdateCommandBars(Control)
        End If

    End Select
End Sub

Private Sub chkInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mobjPopupInfo Is Nothing And Button = vbRightButton Then mobjPopupInfo.ShowPopup
End Sub

Private Sub cmdOk_Click()
    '发送时间
    Dim datBegin As Date, datEnd As Date
    
'    If DateDiff("d", dtpBegin.Value, dtpEnd.Value) > 7 Then
'        If MsgBox("指定的日期间隔超过7天，可能会影响查询速度，是否继续？", vbOKCancel + vbDefaultButton2, Me.Caption) = vbCancel Then Exit Sub
'    End If
    
    mdateEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
    If Trim(txtInfo.Text) = "" Then
        Select Case cboDate.ListIndex
            Case 1              '昨天
                mDateBegin = Format(mdateEnd - 1, "yyyy-MM-dd 00:00:00")
            Case 2              '近三天
                mDateBegin = Format(mdateEnd - 2, "yyyy-MM-dd 00:00:00")
            Case 3              '近一周
                mDateBegin = Format(mdateEnd - 6, "yyyy-MM-dd 00:00:00")
            Case 4              '近十天
                mDateBegin = Format(mdateEnd - 9, "yyyy-MM-dd 00:00:00")
            Case Else           '当天
                mDateBegin = Format(mdateEnd, "yyyy-MM-dd 00:00:00")
        End Select
    Else
        mDateBegin = Format(mdateEnd - 364, "yyyy-MM-dd 00:00:00")
    End If
    
'    If Format(dtpEnd.Value, "yyyy-MM-dd HH:mm") = Format(dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
'        mdateEnd = CDate(0)  '表示取当前时间
'    Else
'        mdateEnd = Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")
'    End If
    
    If Not Me.idkSelect.GetCurCard Is Nothing Then
        If Me.idkSelect.GetCurCard.名称 = "挂号单" Then
            '挂号单自动补齐内容
            txtInfo.Text = zlCommFun.GetFullNO(txtInfo.Text, 12)
        End If
    End If

    '刷新
    mlngPreDept = 0
    Call cboDept_Click

End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.Title Like "病人列表*" Then
        Item.Handle = picLeft.hwnd
    ElseIf Item.Title = "接单内容" Then
        Item.Handle = picRecord.hwnd
    ElseIf Item.ID = 3 Then
        Item.Handle = picReadyReceive.hwnd
    End If
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then
        Bottom = stbThis.Height
    End If
End Sub

Private Function ShowPar() As String
    Dim strPar As String, strType As String, i As Integer
    strPar = zlDatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1")
    For i = 0 To 3
        strType = strType & IIf(Val(Split(strPar, ",")(i)) = 1, "," & i, "")
    Next
    strType = Mid(strType, 2)
    strType = Replace(strType, "0", "治疗")
    strType = Replace(strType, "1", "输液")
    strType = Replace(strType, "2", "注射")
    strType = Replace(strType, "3", "皮试")
    
    ShowPar = "病人列表(" & strType & ")"
End Function

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim arrVal As Variant
    Dim i As Integer
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, gblnShowInTaskBar)

    '创建/初始化一卡通部件
'    mintCards = 0
    Err = 0: On Error Resume Next
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    If Not mobjSquareCard.zlInitComponents(Me, glngModul, glngSys, gstrDBUser, gcnOracle) Then
        Set mobjSquareCard = Nothing
        MsgBox "医疗卡部件（zl9CardSquare）初始化失败！", vbInformation, gstrSysName
    Else
        mstrSquareCards = mobjSquareCard.zlGetIDKindStr(mstrSquareCards)
'        If mstrSquareCards <> "" Then
'            arrVal = Split(mstrSquareCards, ";")
'            mintCards = UBound(arrVal) + 1
'        End If
    End If

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False

    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call initMenus
    
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Set objPane = Me.dkpMain.CreatePane(1, 250, 400, DockLeftOf, Nothing)
    objPane.Title = ShowPar
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    objPane.MinTrackSize.Width = 200
    objPane.MaxTrackSize.Width = 500

    Set objPane = Me.dkpMain.CreatePane(2, 250, 400, DockBottomOf, dkpMain.FindPane(1))
    objPane.Title = "接单内容"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set objPane = Me.dkpMain.CreatePane(3, 250, 60, DockBottomOf, dkpMain.FindPane(2))
    objPane.Title = "待接单病人"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable
    objPane.MaxTrackSize.Height = 60
    objPane.MinTrackSize.Height = 60

    picLeft.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picInfo.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)

    'TabControl
    '-----------------------------------------------------
    Set mclsSeating = New clsDockSeating
    Set mfrmRecord = New frmRecord
    Set mfrmLeaveMedi = New frmLeaveMedi

    Set mcolSubForm = New Collection
    mcolSubForm.Add mclsSeating.zlGetForm, "_座位管理"
    mcolSubForm.Add mfrmRecord, "_执行项目"
    mcolSubForm.Add mfrmLeaveMedi, "_药品寄存"

    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。

        
        '恢复上次选择的卡片
        strTab = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "输液注射", "")
        
        '恢复上次选择的病人信息
        mstrPatiKey = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "提取病人信息", "")
        
        '待接单病人信息
        mintPatiIdentify = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "待接单病人信息", "1"))
        If Val(mintPatiIdentify) <= 0 Then
            mintPatiIdentify = 1
        End If
        
        ''If InStr(mstrPrivs, ";" & "座位安排" & ";") <> 0 Or InStr(mstrPrivs, "座位管理" & ";") <> 0 Then
        '.InsertItem(intIdx, "座位管理", mcolSubForm("_座位管理").hwnd, 0).Tag = "座位管理": intIdx = intIdx + 1
        .InsertItem(intIdx, "座位管理", picTmp.hwnd, 0).Tag = "座位管理": intIdx = intIdx + 1
        ''End If

        If InStr(mstrPrivs, ";" & "医嘱执行" & ";") <> 0 Then
            '.InsertItem(intIdx, "执行项目", mcolSubForm("_执行项目").hwnd, 0).Tag = "执行项目": intIdx = intIdx + 1
            .InsertItem(intIdx, "执行项目", picTmp.hwnd, 0).Tag = "执行项目": intIdx = intIdx + 1
        End If

        If InStr(mstrPrivs, ";" & "药品寄存" & ";") <> 0 Then
            '.InsertItem(intIdx, "药品寄存", mcolSubForm("_药品寄存").hwnd, 0).Tag = "药品寄存": intIdx = intIdx + 1
            .InsertItem(intIdx, "药品寄存", picTmp.hwnd, 0).Tag = "药品寄存": intIdx = intIdx + 1
        End If

        If .ItemCount = 0 Then
            MsgBox "你没有使用输液注射管理的权限。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If

        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = strTab
            .Item(intIdx).Selected = True
            If intIdx = 0 Then tbcSub_SelectedChanged .Item(0)
        Else
            .Item(0).Selected = True
            tbcSub_SelectedChanged .Item(0)
        End If
        Call SubWinDefCommandBar(.Selected) '初始刷新定义一次菜单及按钮
    End With

    '2012-08-23 病人列表分页
    Call TabListInit
    
    '医技科室初始化
    '----------------------------------------------------
    If patiList.DeptToCbo(cboDept, mstrPrivs) = False Then
        MsgBox "初始化医技科室失败,不能使用本系统。", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If

    If cboDept.ListIndex = -1 Then
        If InStr(mstrPrivs, "所有科室") > 0 Then
            MsgBox "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
        Else
            MsgBox "没有发现你所属科室,不能使用本系统。", vbInformation, gstrSysName
        End If
        Unload Me: Exit Sub
    End If

    '
    '发送时间
    '----------------------------
    cboDate.ListIndex = 0
    
    
    '读卡改用IDKindNew控件
    idkSelect.zlInit Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, MSTR_MODE, txtInfo
    idkSelect.IDKind = 1
    For i = 1 To idkSelect.ListCount
        If idkSelect.Cards(i).名称 = mstrPatiKey Then
            idkSelect.IDKind = i
            Exit For
        End If
    Next
    
    ptiReadyReceive.zlInit Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquareCard, MSTR_MODE
    ptiReadyReceive.IDKindIDX = mintPatiIdentify
    
    '病人列表初始
    '--------------------
    Call InitReport
    Call cmdOk_Click
    
    '界面恢复:放在最后执行
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    '初始化监听端口
    'TransUdpSock.SockSend
    
'    Set mobjIDCard = New clsIDCard
'    On Error Resume Next
'    Set mobjICCard = CreateObject("zlICCard.clsICCard")
'    On Error GoTo 0
    
    Call SetTimer '设置自动刷新
    
    
    mbln皮试验证 = Val(zlDatabase.GetPara("皮试验证身份", glngSys, 1264)) <> 0
    
    Call QueueInit
    Call mdlQueueManage.QueueInit
    If Val(zlDatabase.GetPara("移动呼叫", glngSys, 1264)) = 1 Then
        Set mfrmTimeCall = mdlQueueManage.QueueTimeCall
        If Not mfrmTimeCall Is Nothing Then
            mfrmTimeCall.Show , Me
        End If
    End If
    
    '清除35天前的操作日志
    Dim strSQL As String
    strSQL = "zl_门诊输液操作日志_ClearOld"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
End Sub

Private Sub TabListInit()
    Dim intIdx As Integer, strTab As String
    With Me.tbcList
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
                
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        
        rptQueue0.Tag = "0"
        .InsertItem(intIdx, "未接单", picQueue0.hwnd, 0).Tag = "未接单": .Item(intIdx).Visible = False:   intIdx = intIdx + 1   '0-未排队
        rptQueue1.Tag = "1"
        .InsertItem(intIdx, "待配液", picQueue1.hwnd, 0).Tag = "待配液":  intIdx = intIdx + 1  '1 -已接单/待配液（接单后根据参数“有准备/配液流程”决定是否填写）
        rptQueue5.Tag = "5"
        .InsertItem(intIdx, "待穿刺", picQueue5.hwnd, 0).Tag = "待穿刺": intIdx = intIdx + 1  '5 -待穿刺,需要呼叫（输液，注射）
        rptQueue6.Tag = "6"
        .InsertItem(intIdx, "待执行", picQueue6.hwnd, 0).Tag = "待执行": .Item(intIdx).Visible = False:  intIdx = intIdx + 1 '6 -待执行,不需呼叫（皮试，治疗）
        rptQueue7.Tag = "7"
        .InsertItem(intIdx, "执行中", picQueue7.hwnd, 0).Tag = "执行中": intIdx = intIdx + 1  '7 -执行中,
        rptPati.Tag = "2,3,4"
        .InsertItem(intIdx, "已结束", picQueueAll.hwnd, 0).Tag = "已结束": intIdx = intIdx + 1 ' 2 -弃号,3－退号,4－已完成
        
        '恢复上次选择的卡片
        strTab = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "病人列表", "")
        
        For intIdx = 0 To .ItemCount - 1
            If tbcList(intIdx).Visible And tbcList(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= .ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '避免激活事件
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            If mblnLiquid Then
                .Item(1).Selected = True '新建时就自动选中了这个,不会再激活事件
            Else
                .Item(4).Selected = True '新建时就自动选中了这个,不会再激活事件
            End If
        End If
        
    End With
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Height <= 10000 Then Height = 10000
    If Width <= 10000 Then Width = 10000
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strMode As String, strIndex As String

    strMode = idkSelect.GetCurCard.名称
    strIndex = ptiReadyReceive.IDKindIDX
    
    Set patiList = Nothing
    mstrPrivs = ""
    mlngModul = 0
    mstr挂号单 = ""
    mlngPreDept = 0
    mintRow = 0
    mDateBegin = CDate(0)
    mdateEnd = CDate(0)
    
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name _
              , "提取病人信息", strMode)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name _
              , "待接单病人信息", strIndex)

    On Error Resume Next
    
'    mstrIDCard = ""
'    If Not mobjIDCard Is Nothing Then
'        mobjIDCard.SetEnabled False
'        Set mobjIDCard = Nothing
'    End If
'    Set mobjICCard = Nothing
    
    Unload mfrmLeaveMedi
    Unload mfrmRecord
    
    Set mclsSeating = Nothing
    Set mobjSquareCard = Nothing
    Set mfrmTimeCall = Nothing
    
    Call QueueUnload
    
End Sub

Private Sub lblBill_Click()
    If Not mobjPopupInfo Is Nothing Then mobjPopupInfo.ShowPopup
End Sub

Private Sub idkSelect_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtInfo.Enabled And txtInfo.Visible Then
        txtInfo.Text = ""
        txtInfo.SetFocus
    End If
End Sub

Private Sub idkSelect_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtInfo.Text = objPatiInfor.卡号
    mblnReadCard = True
    Call txtInfo_KeyPress(0)
End Sub

Private Sub mclsSeating_RequestRefresh()
    '功能：座位子窗体要求刷新
    mlngPreDept = -1
    Call cboDept_Click
End Sub

Private Sub mclsSeating_StatusTextUpdate(ByVal Text As String)
    '当前选中的座位号
    If InStr(Text, "_") > 0 Then
        If patiList.mSeatings.Item(Text).病人ID = 0 Then
            mstr座位 = Text
        Else
            Dim objPati As cPatient
            For Each objPati In patiList
                If objPati.座位号 = Mid(Text, InStr(Text, "_") + 1) And objPati.病人ID = patiList.mSeatings.Item(Text).病人ID Then
                    mstr挂号单 = objPati.Key
                    Call ShowLblInfo(mstr挂号单)
                End If
            Next
        End If
    End If
End Sub

'Private Sub optDate_Click(Index As Integer)
'
'    Dim curDate As Date
'    curDate = zldatabase.Currentdate
'    dtpEnd.MaxDate = Format(curDate, "yyyy-MM-dd 23:59:59"): dtpBegin.MaxDate = curDate
'
'    dtpEnd.Value = Format(curDate, "yyyy-MM-dd 23:59:59")
'    dtpEnd.Tag = Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")
'
'    Select Case Index
'    Case 1 '昨天
'        curDate = curDate - 1
'    Case 2 '最近三天
'        curDate = curDate - 2
'    Case 3 '最近一周
'        curDate = curDate - 6
'    End Select
'    dtpBegin.Value = CDate(Format(curDate, "yyyy-MM-dd 00:00:00"))
'    optDate(Index).Value = True
'
'End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Left = 0
    fraInfo.Top = -90
    fraInfo.Width = picInfo.ScaleWidth
    fraInfo.Height = picInfo.Height + 90
    
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    
    fraWhere.Top = 0
    fraWhere.Left = 15
    fraWhere.Width = picLeft.ScaleWidth - fraWhere.Left
    
    cboDept.Width = fraWhere.Width - cboDept.Left - 135
    cmdOk.Left = fraWhere.Width - cmdOk.Width - 135
    cboDate.Width = fraWhere.Width - cboDate.Left - cmdOk.Width - 150
    txtInfo.Width = fraWhere.Width - txtInfo.Left - 135
    
    tbcList.Left = picLeft.ScaleLeft
    tbcList.Top = fraWhere.Top + fraWhere.Height
    tbcList.Width = picLeft.ScaleWidth - tbcList.Left
    tbcList.Height = picLeft.ScaleHeight - tbcList.Top
    

End Sub

Private Sub ShowPatiList()
    Dim curDate As Date, objRpt As ReportControl
    Dim datBegin As Date, datEnd As Date
    Dim strOrderCol As String
    Dim arrOrderCol As Variant, arrEle As Variant
    Dim i As Integer
    Dim strInfo As String, strTmp As String, strCard As String
    Dim strReserve As String

    '查询时间段
    curDate = zlDatabase.Currentdate
    If mDateBegin = CDate(0) Then
        mDateBegin = CDate(Format(curDate, "yyyy-mm-dd 00:00:00"))
    End If
    datBegin = mDateBegin

    If mdateEnd = CDate(0) Then
        mdateEnd = CDate(Format(curDate, "yyyy-mm-dd 23:59:59"))
    End If
    datEnd = mdateEnd
    
    strCard = idkSelect.GetCurCard.名称
    
    '指定提取病人信息
    If Trim(txtInfo.Text) <> "" Then
        '一卡通的卡，取卡的类别
        strTmp = GetSquareCardInfo(mstrSquareCards, strCard, enuCardProperty.卡类别ID)
        '准备参数
        Select Case strCard
            Case "就诊卡"
                strInfo = "1"
            Case "门诊号"
                strInfo = "2"
            Case "挂号单"
                strInfo = "3"
            Case "姓名"
                strInfo = "4"
            Case "身份证号", "二代身份证"
                strInfo = "5"
            Case Else
                strInfo = "6"
        End Select
        strInfo = strInfo & "|" & Trim(txtInfo.Text) & "|" & strTmp
    End If

    '显示病人列表
    Call patiList.FetchPatients(mlngPreDept, datBegin, datEnd, , strInfo, , , mobjSquareCard)
    Call PlugInFunc
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        Call gobjPlugIn.TransfusionShowPatiList(glngSys, 1264, mlngPreDept, datBegin, datEnd, strReserve)
        Call zlPlugInErrH(Err, "TransfusionShowPatiList")
        If Err.Number <> 0 Then Err.Clear: On Error GoTo 0
    End If
    If tbcList.Selected.Tag = "未接单" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "待配液" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "待穿刺" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "待执行" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "执行中" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "已结束" Then
        Set objRpt = Me.rptPati
    End If
    
    '保存旧排序
    If objRpt.SortOrder.Count > 0 Then
        For i = 0 To objRpt.SortOrder.Count - 1
            strOrderCol = strOrderCol & objRpt.SortOrder.Column(i).Index & ";" & IIf(objRpt.SortOrder.Column(i).SortAscending, 1, 0)
            If i < objRpt.SortOrder.Count - 1 Then
                strOrderCol = strOrderCol & "|"
            End If
        Next
    End If
    Call patiList.initObjRpt(objRpt, img16)
    '恢复旧排序
    If strOrderCol <> "" Then
        arrOrderCol = Split(strOrderCol, "|")
        objRpt.SortOrder.DeleteAll
        For i = LBound(arrOrderCol) To UBound(arrOrderCol)
            arrEle = Split(arrOrderCol(i), ";")
            objRpt.SortOrder.Add objRpt.Columns(arrEle(0))
            objRpt.SortOrder(i).SortAscending = (arrEle(1) = 1)
        Next
    End If
    Set arrEle = Nothing
    Set arrOrderCol = Nothing
    
    Call patiList.PatiListRefresh(objRpt, objRpt.Tag)
   
    If mintPatirow > 0 And mintPatirow < objRpt.Rows.Count Then
        If Not objRpt.Rows(mintPatirow).GroupRow Then
            Call objRpt.SelectedRows.Add(objRpt.Rows(mintPatirow))
            objRpt.Rows(mintPatirow).Selected = True
            Call RptSelectChanged(objRpt)
        End If
    End If
    '显示当前呼叫行
    Call Calling(0)
    
    Call SubWinRefreshData(tbcSub.Selected)

End Sub

Private Sub picQueue0_Resize()
    Call PicQueueResize(picQueue0, rptQueue0)
End Sub

Private Sub picQueue1_Resize()
    Call PicQueueResize(picQueue1, rptQueue1, lblNo1, txtNo1)
End Sub

Private Sub PicQueueResize(objPic As PictureBox, objRpt As ReportControl, Optional objLbl As Label, Optional objTxt As TextBox)
    Dim panTmp As Pane
    
    On Error Resume Next
    
    With objPic
        objRpt.Left = objPic.ScaleLeft
        objRpt.Top = objPic.ScaleTop
        objRpt.Width = objPic.ScaleWidth
        objRpt.Height = objPic.ScaleHeight
        Set panTmp = dkpMain.FindPane(2)    '得到接单内容Pane
        If Not panTmp Is Nothing Then
            If panTmp.Closed Or panTmp.Hidden Then objRpt.Height = objPic.ScaleHeight - 350
        End If
        Set panTmp = Nothing
    End With
    
    If Not objLbl Is Nothing Then
        With objRpt
            objLbl.Left = .Left + 15
            objLbl.Top = .Top + 15
            objTxt.Left = .Left + 15
            objTxt.Top = objLbl.Top + objLbl.Height + 15
            objTxt.Width = .Width - 30
            
            .Top = objTxt.Top + objTxt.Height + 15
            .Height = .Height - .Top
            
        End With
    End If
End Sub

Private Sub picQueue5_Resize()
    Call PicQueueResize(picQueue5, rptQueue5, lblNo5, txtNo5)
End Sub

Private Sub picQueue6_Resize()
    Call PicQueueResize(picQueue6, rptQueue6)
End Sub

Private Sub picQueue7_Resize()
    Call PicQueueResize(picQueue7, rptQueue7)
End Sub

Private Sub picQueueAll_Resize()
    '原先的列表，作为历史记录的操作界面
    Call PicQueueResize(picQueueAll, rptPati)
    On Error Resume Next
'    lblDate.Left = picQueueAll.ScaleLeft + 25
'    lblDate.Top = picQueueAll.ScaleTop + 15
'
'    cmdOk.Top = lblDate.Top + lblDate.Height + 25
'    cmdOk.Left = picQueueAll.ScaleWidth - cmdOk.Width - 15
'
'    dtpBegin.Top = lblDate.Top + lblDate.Height + 15
'    dtpBegin.Left = picQueueAll.ScaleLeft + 15
'    dtpBegin.Width = (cmdOk.Left - 30) / 2
'    DtpEnd.Top = dtpBegin.Top
'    DtpEnd.Left = dtpBegin.Left + dtpBegin.Width + 15
'    DtpEnd.Width = dtpBegin.Width
'
'    optDate(0).Left = picQueueAll.ScaleLeft + 15
'    optDate(0).Top = dtpBegin.Top + dtpBegin.Height + 15
'
'    optDate(1).Left = optDate(0).Left + optDate(0).Width
'    optDate(1).Top = optDate(0).Top
'
'    optDate(2).Left = optDate(1).Left + optDate(1).Width
'    optDate(2).Top = optDate(0).Top
'
'    optDate(3).Left = optDate(2).Left + optDate(2).Width
'    optDate(3).Top = optDate(0).Top
'
'    optDate(0).Value = 1
    
'    rptPati.Left = picQueueAll.ScaleLeft
'    rptPati.Top = optDate(0).Top + optDate(0).Height + 30
'    rptPati.Width = picQueueAll.ScaleWidth
'    rptPati.Height = picQueueAll.ScaleHeight - rptPati.Top

End Sub

Private Sub picReadyReceive_Resize()
    On Error Resume Next
    ptiReadyReceive.Width = picReadyReceive.ScaleWidth - ptiReadyReceive.Left * 2
End Sub

Private Sub picRecord_Resize()
    On Error Resume Next
    rptRecord.Top = 0
    rptRecord.Left = 0
    rptRecord.Width = picRecord.ScaleWidth
    rptRecord.Height = picRecord.ScaleHeight - Me.stbThis.Height
End Sub

Private Sub ptiReadyReceive_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    If Not objHisPati Is Nothing Then
        ptiReadyReceive.Tag = CLng(objHisPati.病人ID)
    Else
        ptiReadyReceive.Tag = ""
    End If
    blnCancel = True    '录入信息后不改换为病人姓名
End Sub

Private Sub ptiReadyReceive_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    ptiReadyReceive.Text = ""
    If ptiReadyReceive.Visible And ptiReadyReceive.Enabled Then ptiReadyReceive.SetFocus
End Sub

Private Function FindReadyReceivePati(ByVal ptiVar As PatiIdentify) As Boolean
'功能：检查待接单病人的单据
'参数：
'  ptiVar：PatiIdentify控件
'返回：False已接过单或无单可接；True有单未接过

    Dim strSQL As String, strPati As String, strPar As String, strTemp As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim lngPatiId As Long
    
    '病人信息
    Select Case ptiReadyReceive.GetCurCard.名称
    Case "挂号单"
        strPati = " And a.挂号单 = [5] "
    Case "门诊号"
        strPati = " And c.门诊号 = [5] "
    Case "姓名"
        strPati = " And c.姓名 = [5] "
    Case "身份证号", "二代身份证"
        strPati = " And c.身份证号 = [5] "
    Case Else
        If Val(ptiVar.Tag) > 0 Then
            strPati = " And c.病人id = [6] "
            lngPatiId = Val(ptiVar.Tag)
        Else
            strPati = " And c.IC卡号 = [5] "
        End If
    End Select
    
    '单据种类
    strTemp = zlDatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1")
    For i = 0 To 3
        strPar = strPar & IIf(Val(Split(strTemp, ",")(i)) = 1, "," & i, "")
    Next
    
    On Error GoTo hErr
    
    strSQL = "Select a.医嘱id, a.发送号, Max(a.发送数次) 发送数次, Sum(Nvl(b.本次数次, 0)) 已接数次 " & vbNewLine & _
             "From (Select a.病人id, b.医嘱id, b.发送号, b.发送数次 " & vbNewLine & _
             "  From 病人医嘱记录 A, 病人医嘱发送 B, 病人信息 C, 病人挂号记录 D1, 病案主页 D2, 诊疗项目目录 E, 部门表 F " & vbNewLine & _
             "  Where a.Id = b.医嘱id And a.病人id = c.病人id And a.挂号单 = D1.No(+) And a.病人id = D2.病人id(+) And a.主页id = D2.主页id(+) " & vbNewLine & _
             "    And a.诊疗项目id = e.Id And a.执行科室id = f.Id And a.病人来源 In (1, 2) And Decode(D2.病人性质(+), -1, 1, D2.病人性质(+)) = 1 " & vbNewLine & _
             "    And b.执行部门id = [1] And b.发送时间 Between [2] And [3] And D1.记录性质(+) = 1 And D1.记录状态(+) = 1 " & vbNewLine & _
             "    And Instr([4], Nvl(e.执行分类, 0)) > 0 " & vbNewLine & _
             strPati & vbNewLine & _
             ") A, 病人医嘱执行 B " & vbNewLine & _
             "Where a.医嘱id = b.医嘱id(+) And a.发送号 = b.发送号(+) " & vbNewLine & _
             "Group By a.医嘱id, a.发送号 " & vbNewLine & _
             "Having Max(a.发送数次) - Sum(Nvl(b.本次数次, 0)) > 0 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取待接单病人的单据", mlngPreDept, mDateBegin, mdateEnd, strPar, ptiVar.Text, lngPatiId)
    If rsTemp.RecordCount > 0 Then
        '有单未接过
        FindReadyReceivePati = True
    End If
    rsTemp.Close
            
    Exit Function
    
hErr:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ptiReadyReceive_KeyPress(KeyAscii As Integer)
    Dim strCard As String
    Dim lngID As Long
    
    strCard = ptiReadyReceive.Cards(ptiReadyReceive.IDKindIDX).名称

    If KeyAscii = 13 Then
        '获取病人ID
        If Not mobjSquareCard Is Nothing Then
            If ptiReadyReceive.IDKindIDX = 2 Or ptiReadyReceive.IDKindIDX >= 6 Then
                Call mobjSquareCard.zlGetPatiID(ptiReadyReceive.IDKindIDX - 1, ptiReadyReceive.Text, , lngID)
                ptiReadyReceive.Tag = CLng(lngID)
            End If
        End If
        
        '挂号单自动补齐单号
        If strCard = "挂号单" Then
            ptiReadyReceive.Text = zlCommFun.GetFullNO(ptiReadyReceive.Text, 12)
        End If
    
        Call zlControl.TxtSelAll(ptiReadyReceive)
        
        '检查填写的病人待接单数据
        If FindReadyReceivePati(ptiReadyReceive) Then
            '无数据就调用接单界面
            Call thingAdd(Val("1-自动接单"))
        Else
            '有数据
            MsgBox "未找到待接单病人的单据！", vbInformation, gstrSysName
            ptiReadyReceive.SetFocus
        End If

    Else
        Select Case strCard
            Case "门诊号"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "挂号单"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (ptiReadyReceive.Text = "" Or ptiReadyReceive.SelLength = Len(ptiReadyReceive.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "身份证号", "二代身份证"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case Else
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
        End Select
    End If
End Sub

Private Sub rptPati_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptPati, Button, Shift, X, Y)
End Sub

Private Sub rptMouseMove(objRpt As ReportControl, Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = 1 And (Abs(X) > 220 Or Abs(Y) > 220) Then
        If objRpt.SelectedRows.Count > 0 Then
            If Not objRpt.SelectedRows(0).GroupRow Then
                If objRpt.SelectedRows(0).Record(col_图标).Value = "" Then
                    Set objRpt.DragIcon = img16.ListImages("未执行").Picture
                    objRpt.Drag vbBeginDrag
                End If
            End If
        End If
    End If
End Sub

Private Sub rptMouseUp(objRpt As ReportControl, Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    If Button = 2 Then
        If objRpt.Records.Count <= 0 Then Exit Sub
        If Not objRpt.SelectedRows(0).GroupRow Then
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_ManagePopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        Else
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_ViewPopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub
Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptPati, Button, Shift, X, Y)
End Sub


Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Not Row.GroupRow Then
        If Row.Record(col_排队状态).Value = "1-待配液" Then
            Call Calling(0)
        End If
    End If
End Sub

Public Sub rptPati_SelectionChanged()
    Call RptSelectChanged(rptPati)
End Sub

Private Sub RptSelectChanged(objRpt)

    Dim i As Integer
    
    If objRpt.SelectedRows.Count = 0 Then
        If objRpt.Rows.Count > 1 Then
            '有记录,取第个非分组行,做当前行
            For i = 1 To objRpt.Rows.Count - 1
                If Not objRpt.Rows(i).GroupRow Then
                    objRpt.Rows(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If
    
'    Call ShowLblInfo("")
'    Call ShowReport

    If objRpt.SelectedRows.Count = 0 Then Exit Sub  '非正常情况
    mintPatirow = objRpt.SelectedRows(0).Index
    
    With objRpt.SelectedRows(0)
        mfrmRecord.流水号 = Get流水号
        mfrmRecord.编辑 = 0
        mfrmRecord.修改过 = False
        mfrmRecord.组Key = ""

        If Not .GroupRow Then
            mstr挂号单 = .Record(col_key).Value
            Call ShowLblInfo(mstr挂号单)
            Call SubWinRefreshData(tbcSub.Selected)
        Else
            '分组行展开，折叠会改变行数,重新取一次calling所在行数
            Call Calling(0)
            mstr挂号单 = ""
            Call ShowLblInfo(mstr挂号单)
            Call SubWinRefreshData(tbcSub.Selected)
        End If

    End With
End Sub
Private Sub initMenus()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim i As Integer
    Dim strTmp As String, strDefName As String
    Dim arrCard As Variant

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")

        Set objControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&M)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_RoomSet, "穿刺台设置(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Queue_Setup, "呼叫设置(&C)")
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "操作日志(&L)"): objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "执行(&E)", -1, False)
    objMenu.ID = conMenu_ManagePopup
    With objMenu.CommandBar.Controls
    
        'Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Call, "叫号(&J)")
        
'        With objPopup.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Manage_CallNext, "下一位(&N)", -1, False)
'            Set objControl = .Add(xtpControlButton, conMenu_Manage_CallPrevious, "上一位(&P)", -1, False)
'        End With

        'Set objPopup = .Add(xtpControlButtonPopup, conMenu_Manage_Reset, "调整顺序(&R)"): objControl.BeginGroup = True
        'With objPopup.CommandBar.Controls
        '    Set objControl = .Add(xtpControlButton, conMenu_Manage_Up, "上移(&U)", -1, False)
        '    Set objControl = .Add(xtpControlButton, conMenu_Manage_Down, "下移(&D)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Discard, "弃号(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Recall, "召回(&R)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Untread, "退号(&U)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_Manage_TagEnd, "结束(&E)", -1, False): objControl.BeginGroup = True
        'End With
        
         
        
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "接单(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Liquid, "配液")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Call, "呼叫(&J)", -1, False)
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Puncture, "穿刺")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Bed_Modify, "调整状态")
        
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_FindType, "查找方式(&Y)"): objPopup.BeginGroup = True
'        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)")
'        Set objControl = .Add(xtpControlButton, conMenu_View_Filter, "病人过滤(&O)"): objControl.BeginGroup = True
'
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): objControl.BeginGroup = True '固有

    End With

'    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
'    objMenu.ID = conMenu_ToolPopup
'    With objMenu.CommandBar.Controls
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "资料参考(&R)")
'        With objPopup.CommandBar.Controls
'            .Add xtpControlButton, conMenu_Tool_Reference_1, "疾病诊断参考(&D)", -1, False
'            .Add xtpControlButton, conMenu_Tool_Reference_2, "诊疗措施参考(&C)", -1, False
'        End With
'    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)") '固有
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False '固有
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With

'有了精确提取病人，查找取消
    '查找项特殊处理
    '-----------------------------------------------------
'   主菜单右侧的查找 按就诊卡号查找，支持刷卡
'    With cbsMain.ActiveMenuBar.Controls
'        Set objPopup = .Add(xtpControlPopup, conMenu_View_FindType, "查找")
'        objPopup.ID = conMenu_View_FindType
'        objPopup.Flags = xtpFlagRightAlign
'        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
'
'        objCustom.Handle = txtFind.hwnd
'
'        objCustom.Flags = xtpFlagRightAlign
'
'        Set objControl = .Add(xtpControlButton, conMenu_View_ReadIC, "读卡")
'        objControl.Flags = xtpFlagRightAlign
'    End With
    txtFind.Visible = False
    
    
'取消原方式选择病人信息，改用IDKindNew控件代替
'    '弹出式菜单
'    strTmp = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name _
'                         , "提取病人信息", "1"))
'    If Val(strTmp) = 0 Then
'        lblBill.Tag = "1"
'    Else
'        i = Val(strTmp)
'        If i > UBound(Split(mstrSquareCards, ";")) + 1 + 6 Then
'            lblBill.Tag = "1"
'        Else
'            lblBill.Tag = strTmp
'        End If
'    End If
'
'    Set mobjPopupInfo = cbsMain.Add("指定信息", xtpBarPopup)
'    With mobjPopupInfo.Controls
'        .Add xtpControlButton, MLNG_INFO + 1, "就诊卡(&1)"
'        .Add xtpControlButton, MLNG_INFO + 2, "门诊号(&2)"
'        .Add xtpControlButton, MLNG_INFO + 3, "单据号(&3)"
'        .Add xtpControlButton, MLNG_INFO + 4, "姓  名(&4)"
'        .Add xtpControlButton, MLNG_INFO + 5, "身份证(&5)"
'        .Add xtpControlButton, MLNG_INFO + 6, "ＩＣ卡(&6)"
'        '一卡通的卡
'        If mstrSquareCards <> "" Then
'            arrCard = Split(mstrSquareCards, ";")
'            For i = LBound(arrCard) To UBound(arrCard)
'                strTmp = Split(arrCard(i), "|")(enuCardProperty.全名)
'                If Val(lblBill.Tag) = i + 7 Then
'                    strDefName = strTmp
'                End If
'                If InStr(";就诊卡;门诊号;单据号;姓名;身份证;IC卡;ＩＣ卡;", ";" & strTmp & ";") = 0 Then
'                    .Add xtpControlButton, MLNG_INFO + 7 + i, strTmp & "(&" & i + 7 & ")"
'                End If
'            Next
'        End If
'    End With
'    Select Case Val(lblBill.Tag)
'        Case 1
'            lblBill.Caption = "就诊卡"
'        Case 2
'            lblBill.Caption = "门诊号"
'        Case 3
'            lblBill.Caption = "单据号"
'        Case 4
'            lblBill.Caption = "姓  名"
'        Case 5
'            lblBill.Caption = "身份证"
'        Case 6
'            lblBill.Caption = "ＩＣ卡"
'        Case Else
'            If strDefName = "" Then
'                '默认为就诊卡
'                lblBill.Caption = "就诊卡"
'            Else
'                lblBill.Caption = strDefName
'            End If
'    End Select
'    lblBill.Caption = lblBill.Caption & "↓"

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览") '固有

        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "接单"): objControl.BeginGroup = True: objControl.ToolTipText = "接单"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Liquid, "配液")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Call, "呼叫"): objControl.ToolTipText = "呼叫当前人员"
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Puncture, "穿刺")
        
'        Set objControl = .Add(xtpControlButton, conMenu_Manage_CallNext, "下一位"): objControl.BeginGroup = True: objControl.ToolTipText = "呼叫下一位"
'        Set objControl = .Add(xtpControlButton, conMenu_Manage_CallPrevious, "上一位"):: objControl.ToolTipText = "呼叫上一位"

        'Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Manage_Reset, "排号", objControl.Index + 1)
        'objPopup.ID = conMenu_Manage_Reset: objPopup.BeginGroup = True
        'With objPopup.CommandBar.Controls
        '    Set objControl = .Add(xtpControlButton, conMenu_Manage_Up, "上移"): objControl.BeginGroup = True: objControl.ToolTipText = "排队顺序上移"
        '    Set objControl = .Add(xtpControlButton, conMenu_Manage_Down, "下移"): objControl.ToolTipText = "排队顺序下移"
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Discard, "弃号"): objControl.ToolTipText = "暂离排队序列"
            Set objControl = .Add(xtpControlButton, conMenu_Manage_TagEnd, "结束"): objControl.ToolTipText = "标记为结束序列"
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Recall, "召回"): objControl.ToolTipText = "返回排队序列"
            Set objControl = .Add(xtpControlButton, conMenu_Manage_Untread, "退号"): objControl.ToolTipText = "退出排队序列"
        'End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出") '固有
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend          '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse   '折叠所有组
        .Add 0, vbKeyF12, conMenu_File_Parameter            '参数设置
        
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add 0, vbKeyF3, conMenu_Manage_Call                '呼叫
        .Add FCONTROL, vbKeyPageUp, conMenu_Manage_Up       '上移
        .Add FCONTROL, vbKeyPageDown, conMenu_Manage_Down   '下移
        .Add 0, vbKeyF5, conMenu_View_Refresh               '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
    End With

    '设置一些公共的不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet         '打印设置
        .AddHiddenCommand conMenu_File_Excel            '输出到Excel
    End With

    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)

End Sub


Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
'功能：刷新界面和数据
'参数：
'  objItem：tbcTab控件的Item对象

    Dim objPati As cPatient
    Dim strOutNurse As String, objNurse As OutNurse, lng病人ID As Long
    For Each objNurse In ObjOutNurse
        strOutNurse = strOutNurse & "|" & objNurse.姓名
    Next
    If Mid(strOutNurse, 1, 1) = "|" Then strOutNurse = Mid(strOutNurse, 2)
    Call cbsMain_Resize
    Select Case objItem.Caption
    Case "座位管理"
        lng病人ID = 0
        Set objPati = Nothing
        
        If mstr挂号单 <> "" Then
            If Not patiList.Item(mstr挂号单) Is Nothing Then
                If patiList.Item(mstr挂号单).座位号 = "无" Or patiList.Item(mstr挂号单).座位号 = "" Then
                    lng病人ID = patiList.Item(mstr挂号单).病人ID
                    Set objPati = patiList.Item(mstr挂号单)
                End If
            End If
        End If
        Call mclsSeating.zlRefresh(patiList.mSeatings, lng病人ID, objPati)
    Case "执行项目"
        Set objPati = Nothing
        
        If mstr挂号单 <> "" Then
            If Not patiList.Item(mstr挂号单) Is Nothing Then
                lng病人ID = patiList.Item(mstr挂号单).病人ID
                Set objPati = patiList.Item(mstr挂号单)
            End If
        End If
        Call mfrmRecord.zlRefresh(mobjRecord, objPati)
    Case "药品寄存"
        If mstr挂号单 <> "" Then
            If Not patiList.Item(mstr挂号单) Is Nothing Then
                mfrmLeaveMedi.dateBeging = mDateBegin
                mfrmLeaveMedi.DateEnd = mdateEnd
                mfrmLeaveMedi.病人ID = patiList.Item(mstr挂号单).病人ID
                mfrmLeaveMedi.挂号单 = mstr挂号单
                mfrmLeaveMedi.姓名 = lblinfo(1)
                mfrmLeaveMedi.性别 = lblinfo(3)
                mfrmLeaveMedi.年龄 = lblinfo(5)
                mfrmLeaveMedi.科室ID = mlngPreDept
                mfrmLeaveMedi.科室 = cboDept.List(cboDept.ListIndex)
                Call mfrmLeaveMedi.zlRefresh
            End If
        End If
    End Select
End Sub

'Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
''功能：身份证识别成功后激活
'    mstrIDCard = strID
'    If mintFindType = 4 Then
'        txtFind.Text = mstrIDCard
'    Else
'        txtFind.Text = "" '否则清除(目前是在已清除情况下才能激活)。
'    End If
'    Call ExecuteFindPati(False, mstrIDCard)
'End Sub

Private Sub rptQueue0_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue0, Button, Shift, X, Y)
End Sub

Private Sub rptQueue0_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue0, Button, Shift, X, Y)
End Sub

Private Sub rptQueue0_SelectionChanged()
    Call RptSelectChanged(rptQueue0)
End Sub

Private Sub rptQueue1_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue1, Button, Shift, X, Y)
End Sub

Private Sub rptQueue1_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue1, Button, Shift, X, Y)
End Sub

Private Sub rptQueue1_SelectionChanged()
    Call RptSelectChanged(rptQueue1)
End Sub

Private Sub rptQueue5_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue5, Button, Shift, X, Y)
End Sub

Private Sub rptQueue5_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue5, Button, Shift, X, Y)
End Sub

Private Sub rptQueue5_SelectionChanged()
    Call RptSelectChanged(rptQueue5)
End Sub

Private Sub rptQueue6_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue6, Button, Shift, X, Y)
End Sub

Private Sub rptQueue6_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue6, Button, Shift, X, Y)
End Sub

Private Sub rptQueue6_SelectionChanged()
    Call RptSelectChanged(rptQueue6)
End Sub

Private Sub rptQueue7_MouseMove(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseMove(rptQueue7, Button, Shift, X, Y)
End Sub

Private Sub rptQueue7_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Call rptMouseUp(rptQueue7, Button, Shift, X, Y)
End Sub

Private Sub rptQueue7_SelectionChanged()
    Call RptSelectChanged(rptQueue7)
End Sub

Private Sub rptRecord_GotFocus()
    mfrmRecord.流水号 = Get流水号
End Sub

Private Sub rptRecord_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    Dim objRpt As ReportControl
    
    If Button = 2 Then
        If tbcList.Selected.Tag = "未接单" Then
            Set objRpt = Me.rptQueue0
        ElseIf tbcList.Selected.Tag = "待配液" Then
            Set objRpt = Me.rptQueue1
        ElseIf tbcList.Selected.Tag = "待穿刺" Then
            Set objRpt = Me.rptQueue5
        ElseIf tbcList.Selected.Tag = "待执行" Then
            Set objRpt = Me.rptQueue6
        ElseIf tbcList.Selected.Tag = "执行中" Then
            Set objRpt = Me.rptQueue7
        ElseIf tbcList.Selected.Tag = "已结束" Then
            Set objRpt = Me.rptPati
        End If
    
        Call SubWinRefreshData(tbcSub.Selected)     '刷新
    
        If objRpt.SelectedRows.Count <= 0 Then Exit Sub
        If Not objRpt.SelectedRows(0).GroupRow Then
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_ManagePopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        Else
            Set objPopup = cbsMain.ActiveMenuBar.FindControl(, conMenu_ViewPopup)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    Else
        Call SubWinRefreshData(tbcSub.Selected)     '刷新
    End If
    
End Sub

Public Sub rptRecord_SelectionChanged()
    Dim i As Integer
    
    mfrmRecord.流水号 = 0
    
    If rptRecord.SelectedRows.Count = 0 Then
        If mintRecordRow > 0 And mintRecordRow < rptRecord.Rows.Count Then
            If Not rptRecord.Rows(mintRecordRow).GroupRow Then
                Call rptRecord.SelectedRows.Add(rptRecord.Rows(mintRecordRow))
                rptRecord.Rows(mintRecordRow).Selected = True
            End If
        End If
    End If

    If rptRecord.SelectedRows.Count = 0 Then
        If rptRecord.Rows.Count > 1 Then
            '有记录,取第个非分组行,做当前行
            For i = 1 To rptRecord.Rows.Count - 1
                If Not rptRecord.Rows(i).GroupRow Then
                    rptRecord.Rows(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If

    If rptRecord.SelectedRows.Count = 0 Then Exit Sub '还是没有选择的行,则退出
    
    mintRecordRow = rptRecord.SelectedRows(0).Index
    If mfrmRecord.编辑 = 0 Then
        '浏览模式
        mfrmRecord.流水号 = Get流水号
        Call mfrmRecord.ShowVsList(mfrmRecord.流水号)
        Call mfrmRecord.KernalRefresh
    Else
        '修改模式
        mfrmRecord.流水号 = Get流水号
        If mfrmRecord.流水号 <> mfrmRecord.编辑 Then
            MsgBox "请将当前记录的修改完成之后，再做其他操作。", vbExclamation, gstrSysName
            mfrmRecord.流水号 = mfrmRecord.编辑
            Exit Sub
        End If
        If Not mfrmRecord.修改过 Then
            '有过修改，不刷新。
            
            Call mfrmRecord.ShowVsList(mfrmRecord.流水号)
            Call mfrmRecord.KernalRefresh
            
        End If
    End If
    mfrmRecord.组Key = ""
End Sub

Private Sub tbcList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim objRpt As ReportControl
    
    If Item.Tag = "" Then Exit Sub
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "病人列表", Item.Tag
    mstrQueueTab = Item.Tag
    If Item.Tag = "未接单" Then
        Set objRpt = Me.rptQueue0
    ElseIf Item.Tag = "待配液" Then
        Set objRpt = Me.rptQueue1
    ElseIf Item.Tag = "待穿刺" Then
        Set objRpt = Me.rptQueue5
    ElseIf Item.Tag = "待执行" Then
        Set objRpt = Me.rptQueue6
    ElseIf Item.Tag = "执行中" Then
        Set objRpt = Me.rptQueue7
    ElseIf Item.Tag = "已结束" Then
        Set objRpt = Me.rptPati
    End If
    If Not objRpt Is Nothing Then
        Call patiList.PatiListRefresh(objRpt, objRpt.Tag)
        '选中一行
        Call 刷新(1)
    End If
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    'If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    
    If Item.Tag = "" Then Exit Sub
    
    '保存选择
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "输液注射", Item.Tag
   
    On Error GoTo errHandle
    Screen.MousePointer = vbHourglass
    If picTmp.hwnd = Item.Handle Then
        Dim objItem As TabControlItem
        Dim intIndex As Integer
        intIndex = Item.Index
        Select Case Item.Tag
            Case "座位管理"
                Set objItem = tbcSub.InsertItem(intIndex, "座位管理", mcolSubForm("_座位管理").hwnd, 0)
                objItem.Tag = "座位管理"
            Case "执行项目"
                Set objItem = tbcSub.InsertItem(intIndex, "执行项目", mcolSubForm("_执行项目").hwnd, 0)
                objItem.Tag = "执行项目"
            Case "药品寄存"
                Set objItem = tbcSub.InsertItem(intIndex, "药品寄存", mcolSubForm("_药品寄存").hwnd, 0)
                objItem.Tag = "药品寄存"
        End Select
        If Not objItem Is Nothing Then
            objItem.Selected = True
            tbcSub.RemoveItem intIndex + 1
        
            Call SubWinDefCommandBar(objItem)
            '刷新子窗体数据
            Call SubWinRefreshData(objItem)
        
        End If
    Else
        Call SubWinDefCommandBar(Item)
        '刷新子窗体数据
        Call SubWinRefreshData(Item)
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHandle:
    Screen.MousePointer = vbDefault
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ParameterSetup()
    Dim strRoom As String

    frmTransfusionSetup.mstrPrivs = mstrPrivs
    frmTransfusionSetup.mlng科室ID = cboDept.ItemData(cboDept.ListIndex)
    frmTransfusionSetup.Show vbModal, Me
    If frmTransfusionSetup.mblnOk Then
        '皮试验证身份
        mbln皮试验证 = Val(zlDatabase.GetPara("皮试验证身份", glngSys, 1264)) <> 0
        
        '设置自动刷新
        Call SetTimer
        '重启列表
        Me.dkpMain.FindPane(1).Title = ShowPar
        Call cmdOk_Click
    End If
    timRefresh.Enabled = mintRefresh <> 0
    
End Sub

Private Sub ShowLblInfo(ByVal str挂号单 As String)
    Dim objPati As cPatient
    Dim dateS As Date, dateE As Date
    On Error GoTo hNoPati
    
    If str挂号单 = "" Then
        GoTo hNoPati
    Else
        Set objPati = patiList.Item(str挂号单)
        
        If Not objPati Is Nothing Then
            lblinfo(1) = objPati.姓名
            lblinfo(3) = objPati.性别
            lblinfo(5) = objPati.年龄
            lblinfo(7) = objPati.费别
            lblinfo(11) = objPati.病人科室
            lblinfo(13) = objPati.门诊诊断
            lblinfo(15) = objPati.就诊卡号
                       
            Set mobjRecord = New ExecRecord
            dateS = objPati.挂号时间
            dateE = zlDatabase.Currentdate
            Call mobjRecord.GetExecGroups(objPati, mlngPreDept, 1, dateS, dateE)
         
        Else
            GoTo hNoPati
        End If

    End If
    Exit Sub
hNoPati:
    lblinfo(1) = ""
    lblinfo(3) = ""
    lblinfo(5) = ""
    lblinfo(7) = ""
    lblinfo(11) = ""
    lblinfo(13) = ""
    lblinfo(15) = ""
    Set mobjRecord = Nothing '清空项目
        
End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'功能：刷新子窗体菜单及工具条
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long

    '记录现有菜单样式
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        blnShowBar = cbsMain(2).Visible
        bytStyle = cbsMain(2).Controls(1).STYLE
    End If

    '刷新子窗口菜单
    Call LockWindowUpdate(Me.hwnd)

    Me.Caption = "门诊输液注射管理 - " & objItem.Caption

    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.Count To 2 Step -1
        cbsMain(lngCount).Delete
    Next

    '主窗口重新加入
    Call initMenus

    '子窗口重新加入

    Call dkpMain.FindPane(2).Close
    Select Case objItem.Tag
    Case "执行项目"
        Call mfrmRecord.zlDefCommandBars(Me, Me.cbsMain)
        Call dkpMain.ShowPane(2)
    Case "座位管理"
        Call mclsSeating.zlDefCommandBars(Me, Me.cbsMain, 1)
    Case "药品寄存"
        Call mfrmLeaveMedi.zlDefCommandBars(Me, Me.cbsMain)
    End Select

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
        For Each objControl In cbsMain(lngCount).Controls
            objControl.STYLE = bytStyle
        Next
        cbsMain(lngCount).Visible = blnShowBar
    Next

    '如果用了RecalcLayout反而不正常
    Call LockWindowUpdate(0)

    'Set mfrmActive = mcolSubForm("_" & zlCommFun.NVL(objItem.Tag))
End Sub

Private Sub Calling(ByVal intRow As Integer)
    '显示当前呼叫位置
    Dim intFister As Integer
    Dim objRpt As ReportControl
    Dim blnInitRow As Boolean, i As Integer

    '-- 重新取当前呼叫行
    If tbcList.Selected.Tag = "未接单" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "待配液" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "待穿刺" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "待执行" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "执行中" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "已结束" Then
        Set objRpt = Me.rptPati
    End If
    
    blnInitRow = True
    For i = 1 To objRpt.Rows.Count - 1
        If Not objRpt.Rows(i).GroupRow Then
            If objRpt.Rows(i).Record(col_排队状态).Value = "5-待穿刺" Then
                If intFister = 0 Then intFister = i
                If objRpt.Rows(i).Record(col_calling).Icon = 5 Then
                    mintRow = i
                    blnInitRow = False
                    Exit For
                End If
            End If
        End If
    Next
    If blnInitRow = True And intFister > 0 Then
        If mintRow <= 0 Or mintRow > objRpt.Rows.Count Then mintRow = intFister
    ElseIf intFister = 0 Then
        mintRow = 0
    End If
'
'    '-- 更新显示
'    If mintRow + intRow > 0 And mintRow + intRow <= rptPati.Rows.Count Then
'        If Not rptPati.Rows(mintRow + intRow).GroupRow Then
'            If rptPati.Rows(mintRow + intRow).Record(col_排队状态).Value = "1-待配液" Then
'                rptPati.Rows(mintRow).Record(col_calling).Icon = 6
'                rptPati.Rows(mintRow + intRow).Record(col_calling).Icon = 5
'                mintRow = mintRow + intRow
'                rptPati.Redraw
'            End If
'        End If
'    End If

    Dim objPati As cPatient
    With objRpt
        If .SelectedRows.Count <= 0 Then Exit Sub
        If .SelectedRows(0).GroupRow Then Exit Sub
        If .SelectedRows(0).Record(col_排队状态).Value <> "5-待穿刺" Then Exit Sub
 
        If patiList.Item(mstr挂号单).SetCallTag(mlngPreDept) Then
            For Each objPati In patiList
                objPati.呼叫标志 = 0
            Next
            patiList.Item(mstr挂号单).呼叫标志 = 1
            'Call patiList.PatiListRefresh(rptPati)
            .Rows(mintRow).Record(col_calling).Icon = 6
            .SelectedRows(0).Record(col_calling).Icon = 5
            .Redraw
        End If
        
    End With

End Sub

Private Function SiblingRowState(objRpt As ReportControl, ByVal intRow As Integer) As SiblingRow
    '取相邻行的状态
    With SiblingRowState
        If intRow + 1 < objRpt.Rows.Count Then
            If Not objRpt.Rows(intRow + 1).GroupRow Then
                .nextRow挂号单 = objRpt.Rows(intRow + 1).Record(col_key).Value
                .nextRow状态 = objRpt.Rows(intRow + 1).Record(col_排队状态).Value
                .nextRowIndex = intRow + 1
            End If
        End If

        If intRow - 1 >= 0 And intRow <= objRpt.Rows.Count Then
            If Not objRpt.Rows(intRow - 1).GroupRow Then
                .PrivRow挂号单 = objRpt.Rows(intRow - 1).Record(col_key).Value
                .PrivRow状态 = objRpt.Rows(intRow - 1).Record(col_排队状态).Value
                .PrivRowIndex = intRow - 1
            End If
        End If

        If intRow >= 0 And intRow < objRpt.Rows.Count Then
            If Not objRpt.Rows(intRow).GroupRow Then
                .curRow挂号单 = objRpt.Rows(intRow).Record(col_key).Value
                .curRow状态 = objRpt.Rows(intRow).Record(col_排队状态).Value
                .curRowIndex = intRow
            End If
        End If
    End With
End Function

Private Sub rptQueueMove(ByVal intRow As Integer)
    '移动位置
    Dim icurRow As Integer, lngTmp As Long
    Dim TcurrowStat As SiblingRow, TobjRowStat As SiblingRow
    Dim objRpt As ReportControl
    If tbcList.Selected.Tag = "未接单" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "待配液" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "待穿刺" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "待执行" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "执行中" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "已结束" Then
        Set objRpt = Me.rptPati
    End If
    If objRpt.SelectedRows.Count > 0 Then
        icurRow = objRpt.SelectedRows(0).Index

        TcurrowStat = SiblingRowState(objRpt, icurRow)  '取相邻行状态
        TobjRowStat = SiblingRowState(objRpt, icurRow + intRow)

        If (icurRow + intRow > 0) And ((icurRow + intRow) < objRpt.Rows.Count) Then

            lngTmp = Val(patiList.Item(TcurrowStat.curRow挂号单).加权号)
            patiList.Item(TcurrowStat.curRow挂号单).加权号 = Val(patiList.Item(TobjRowStat.curRow挂号单).加权号)
            patiList.Item(TobjRowStat.curRow挂号单).加权号 = lngTmp

            Call patiList.Item(TcurrowStat.curRow挂号单).UpdateSequence(mlngPreDept)
            Call patiList.Item(TobjRowStat.curRow挂号单).UpdateSequence(mlngPreDept)

            Call patiList.PatiListRefresh(objRpt, objRpt.Tag)
            objRpt.Rows(TobjRowStat.curRowIndex).Selected = True
            
            If mintRow = icurRow Then
                If (mintRow + intRow) > 0 And ((mintRow + intRow) < objRpt.Rows.Count) Then
                    mintRow = mintRow + intRow
                End If
            Else
                If icurRow + intRow = mintRow Then mintRow = mintRow - intRow
            End If
            'Call Calling(0)
        End If
    End If
End Sub

Public Function UpdateState(ByVal strState) As Boolean
    '修改排队的状态
    Dim objRpt As ReportControl
    
    UpdateState = False
    If InStr("2-弃号,1-待配液,3-退号,4-结束,5-待穿刺,6-待执行,7-执行中", strState) > 0 Then

        If mstr挂号单 <> "" Then
            If tbcList.Selected.Tag = "未接单" Then
                Set objRpt = Me.rptQueue0
            ElseIf tbcList.Selected.Tag = "待配液" Then
                Set objRpt = Me.rptQueue1
            ElseIf tbcList.Selected.Tag = "待穿刺" Then
                Set objRpt = Me.rptQueue5
            ElseIf tbcList.Selected.Tag = "待执行" Then
                Set objRpt = Me.rptQueue6
            ElseIf tbcList.Selected.Tag = "执行中" Then
                Set objRpt = Me.rptQueue7
            ElseIf tbcList.Selected.Tag = "已结束" Then
                Set objRpt = Me.rptPati
            End If
                    
            If patiList.Item(mstr挂号单).UpdateState(strState, mlngPreDept) Then
                UpdateState = True
                Call patiList.PatiListRefresh(objRpt, objRpt.Tag)
                'Call Calling(0)
                If Val(strState) >= 2 And Val(strState) <= 4 Then ClearSeat
            End If
        End If
    End If
End Function

'Private Sub TransUdpSock_inSockString()
'    '处理本模块收到的消息
'    If TransUdpSock.Infos(TransUdpSock.Infos.Count).发送模块 = con_发送模块 And _
'       TransUdpSock.本机IP = TransUdpSock.Infos(TransUdpSock.Infos.Count).发送IP Then
'        '本机发送的
'        Call MsgBox("收到本模块发送的信息:" & TransUdpSock.Infos)
'
'    End If
'End Sub


Private Sub ClearSeat()
    
    '执行清除座位操作
    If patiList.Item(mstr挂号单).座位号 <> "" Then
        Dim objSeat As Seating
        For Each objSeat In patiList.mSeatings
            If objSeat.病人ID = patiList.Item(mstr挂号单).病人ID And patiList.Item(mstr挂号单).座位号 = objSeat.编号 Then
                Call patiList.mSeatings.Clear(objSeat.类别 & "_" & objSeat.编号)
                '刷新座位
                Call mclsSeating.zlRefresh(patiList.mSeatings, patiList.Item(mstr挂号单).病人ID, patiList.Item(mstr挂号单))
            End If
        Next
    End If
    
End Sub
Private Sub InitReport()
    '窗体LOad时调用一次
    Dim objCol As ReportColumn
    With rptRecord
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行

        Set objCol = .Columns.Add(rptCOL.rptCOL_执行分类, "执行分类", 0, False): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(rptCOL.rptCOL_接单时间, "接单时间", 80, True)
        Set objCol = .Columns.Add(rptCOL.rptCOL_配药人, "配药人", 60, True)
        Set objCol = .Columns.Add(rptCOL.rptCOL_接单人, "接单人", 60, True)

        '隐藏数据列
        Set objCol = .Columns.Add(rptCOL_耗时, "耗时", 0, False)
        Set objCol = .Columns.Add(rptCOL_滴系数, "滴系数", 0, False)
        Set objCol = .Columns.Add(rptCOL_组数, "组数", 0, False)
        Set objCol = .Columns.Add(rptCOL_流水号, "流水号", 0, False)


        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = rptCOL_执行分类
            If objCol.Width = 0 Then objCol.Visible = False
        Next

        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."

        End With

        .PreviewMode = True

        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList img16

        .GroupsOrder.Add .Columns(rptCOL_执行分类)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的

        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(rptCOL_接单时间)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(rptCOL_流水号)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Public Sub ShowReport()
    Dim i As Integer
    '清除数据
    rptRecord.Records.DeleteAll

    If mobjRecord Is Nothing Then
        rptRecord.Populate '更新显示
        Exit Sub
    End If
    '显示数据
    With mobjRecord
        For i = 1 To mobjRecord.Count
            Call AddRecord(mobjRecord.Item(i))
        Next
    End With
    rptRecord.Populate
    Call rptRecord_SelectionChanged
End Sub

Public Function Get流水号() As Long
    '取当前选定行的流水号
    Get流水号 = 0
    On Error GoTo errHandle
    With rptRecord
        If .SelectedRows.Count > 0 Then
            If Not .SelectedRows(0).GroupRow Then
                If .Columns.Count > rptCOL_流水号 Then
                    Get流水号 = .SelectedRows(0).Record(rptCOL_流水号).Value
                End If
            End If
        End If
    End With
    Exit Function
errHandle:
    Get流水号 = 0
    If Err.Number = 5 Then
        Exit Function
    Else
        If ErrCenter = 1 Then
            Resume
        End If
    End If
End Function

Private Sub AddRecord(ByVal objExecRecord As ExecutiveGroup)
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim intIcon As Integer
    With objExecRecord
        Set objRecord = rptRecord.Records.Add
        Call Add_rptItem(objRecord, .执行分类)
        Call Add_rptItem(objRecord, Format(.执行时间, "MM-dd hh:mm"))
        Call Add_rptItem(objRecord, .配药人)
        Call Add_rptItem(objRecord, .接单人)

        Call Add_rptItem(objRecord, IIf(.总耗时 = 0, "", .总耗时))
        Call Add_rptItem(objRecord, IIf(.滴系数 = 0, "", .滴系数))
        Call Add_rptItem(objRecord, .组数)
       
        
        Call Add_rptItem(objRecord, .流水号)
        Select Case Val(Mid(.执行分类, 1, 1))

        Case 1
            '输液
            objRecord.PreviewText = "组数:" & .组数 & _
                                    IIf(.滴系数 = 0, "", " 滴系数:" & .滴系数) & _
                                    IIf(.总耗时 = 0, "", " 耗时:" & .总耗时)
        Case 2
            '注射
            objRecord.PreviewText = "组数:" & .组数
        Case 3
            '皮试
            objRecord.PreviewText = IIf(.总耗时 = 0, "", " 耗时:" & .总耗时)
        Case Else
            '治疗
        End Select
    End With
End Sub

Private Function Add_rptItem(ByRef objRecord As ReportRecord, ByVal strValues As String) As ReportRecordItem
    Set Add_rptItem = objRecord.AddItem(strValues)
    Add_rptItem.Caption = strValues

End Function

Public Sub 撤消接单(ByVal lng流水号)
    '----------
    '子窗体调用
    Dim lngRow As Long, lngDeptID As Long, objRpt As ReportControl
    Dim lngErrNo As Long
    Dim objPati As cPatient
    
    lngDeptID = cboDept.ItemData(cboDept.ListIndex)
    Call mobjRecord.Item(CStr(lng流水号)).Undo(lng流水号, lngDeptID, lngErrNo)
    If lngErrNo <> 0 Then Exit Sub
    
    Set objPati = patiList.Item(mstr挂号单)
    
    SaveOperLog lngDeptID, objPati, MEDICAL, "流水号为" & lng流水号 & "的医嘱执行了撤单操作"
    If tbcList.Selected.Tag = "未接单" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "待配液" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "待穿刺" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "待执行" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "执行中" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "已结束" Then
        Set objRpt = Me.rptPati
    End If
    
    If objRpt.SelectedRows.Count > 0 Then lngRow = objRpt.SelectedRows(0).Index
    
    
    rptRecord.SelectedRows(0).Record.DeleteAll
    Call mobjRecord.Remove(CStr(lng流水号))
    Call 刷新(lngRow)
    rptRecord.Populate
End Sub

Public Sub 刷新(Optional ByVal lngRow As Long)
    '子窗体调用
    Dim objRpt As ReportControl
    mlngPreDept = -1
    mdateEnd = CDate(0)
    Call cmdOk_Click
    If tbcList.Selected.Tag = "未接单" Then
        Set objRpt = Me.rptQueue0
    ElseIf tbcList.Selected.Tag = "待配液" Then
        Set objRpt = Me.rptQueue1
    ElseIf tbcList.Selected.Tag = "待穿刺" Then
        Set objRpt = Me.rptQueue5
    ElseIf tbcList.Selected.Tag = "待执行" Then
        Set objRpt = Me.rptQueue6
    ElseIf tbcList.Selected.Tag = "执行中" Then
        Set objRpt = Me.rptQueue7
    ElseIf tbcList.Selected.Tag = "已结束" Then
        Set objRpt = Me.rptPati
    End If
    
    If lngRow > 0 And lngRow <= objRpt.Rows.Count Then
        objRpt.SetFocus
        Call objRpt.SelectedRows.Add(objRpt.Rows(lngRow))
        objRpt.Rows(lngRow).Selected = True
        'Call RptSelectChanged(objRpt)
    End If
End Sub

Public Sub 更新状态栏(ByVal strText As String)
    Me.stbThis.Panels(2).Text = strText
End Sub

Private Sub ExecuteFindPati(Optional ByVal blnNext As Boolean, Optional ByVal strIDCard As String)
    '功能：查找(下一个)病人
    '参数：blnNext=是否查找下一个
    '      strIDCard=当有值时，表示固定按身份证号查找

    
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    Dim objRpt As ReportControl, strNO As String
'    '按其他方式查找后，自动刷身份证的继续查找则取消
'    If strIDCard = "" And txtFind.Text <> "" Then mstrIDCard = ""
    
    If Not blnNext And mintFindType = 2 Then
        txtFind.Text = GetFullNO(txtFind.Text, 12)  '12－挂号收据号
    End If
    Call zlControl.TxtSelAll(txtFind)
          
    
    '------------------------------------------------------------------------
    Dim objPati As cPatient, intFind As Integer, iCount As Integer
    
    If blnReStart Then mintLastFind = 0
    
    If blnNext Then
        intFind = mintLastFind + 1
    Else
        intFind = 1
    End If
    strNO = ""
    
    For Each objPati In patiList
        If strIDCard <> "" Then '身份证自动识别强制优先
            If UCase(objPati.身份证号) = UCase(strIDCard) Then
                iCount = iCount + 1
                strNO = objPati.挂号单
            End If
        Else
            If mintFindType = 0 Then '就诊卡
                If UCase(objPati.就诊卡号) = UCase(txtFind.Text) Then
                    strNO = objPati.挂号单
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 1 Then '门诊号
                If UCase(objPati.门诊号) = UCase(txtFind.Text) Then
                    strNO = objPati.挂号单
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 2 Then '单据号
                If UCase(objPati.挂号单) = UCase(txtFind.Text) Then
                    strNO = objPati.挂号单
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 3 Then '姓名
                If UCase(objPati.姓名) Like "*" & UCase(txtFind.Text) & "*" Then
                    strNO = objPati.挂号单
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 4 Then '身份证
                If UCase(objPati.身份证号) = UCase(txtFind.Text) Then
                    strNO = objPati.挂号单
                    iCount = iCount + 1
                End If
            End If
            If mintFindType = 5 Then 'IC卡
                If UCase(objPati.IC卡号) = UCase(txtFind.Text) Then
                    strNO = objPati.挂号单
                    iCount = iCount + 1
                End If
            End If
        End If
        
        If iCount = intFind Then Exit For
    Next
    If iCount > 0 And iCount = intFind Then
        mintLastFind = iCount
    Else
        strNO = ""
    End If
    If strNO = "" Then
        blnReStart = True
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
        Exit Sub
    End If
'    If Val(patiList(strNo).排队状态) = 0 Then
'        Set objRpt = Me.rptQueue0
'        tbcList.Item(0).Selected = True
'    Else
    If Val(patiList(strNO).排队状态) = 1 Then
        Set objRpt = Me.rptQueue1
        tbcList.Item(1).Selected = True
    ElseIf Val(patiList(strNO).排队状态) >= 2 And Val(patiList(strNO).排队状态) <= 4 Then
        Set objRpt = Me.rptPati
        tbcList.Item(5).Selected = True
    ElseIf Val(patiList(strNO).排队状态) = 5 Then
        Set objRpt = Me.rptQueue5
        tbcList.Item(2).Selected = True
    ElseIf Val(patiList(strNO).排队状态) = 6 Then
        Set objRpt = Me.rptQueue6
        tbcList.Item(3).Selected = True
    ElseIf Val(patiList(strNO).排队状态) = 7 Or Val(patiList(strNO).排队状态) = 0 Then
        Set objRpt = Me.rptQueue7
        tbcList.Item(4).Selected = True
    End If
    '------------------------------------------------------------------------
            
    '开始查找行
    
    i = 0 'ReportControl的索引从是0开始
    '查找病人
    
    For i = i To objRpt.Rows.Count - 1
        With objRpt.Rows(i)
            If Not .GroupRow Then
                If .Record(col_挂号单).Value = strNO Then Exit For
            End If
        End With
    Next

    If i <= objRpt.Rows.Count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set objRpt.FocusedRow = objRpt.Rows(i)
        objRpt.SetFocus
        tmrAutoReady.Enabled = True
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的病人。", vbInformation, gstrSysName
    End If

End Sub

Private Sub timRefresh_Timer()
    Static lngSecond As Long
    
    If mintRefresh = 0 Then Exit Sub
    lngSecond = lngSecond + 1 '秒数
    If lngSecond Mod mintRefresh = 0 Then
        lngSecond = 0
        Call 刷新
    End If
End Sub

Private Sub SetTimer()
    mintRefresh = Val(zlDatabase.GetPara("医技刷新间隔", glngSys, 1264))
    If mintRefresh <> 0 And mintRefresh < 30 Then mintRefresh = 30
    If mintRefresh = 0 Then
        timRefresh.Enabled = False
    Else
        timRefresh.Interval = 1000 '固定为1秒钟
        timRefresh.Enabled = True
    End If
End Sub

Private Sub tmrAutoReady_Timer()
    '找到病人自动接单 2012-05-14
    
    tmrAutoReady.Enabled = False
    If Val(zlDatabase.GetPara("门诊输液自动接单", glngSys, 1264)) <> 0 Then
        If cbsMain.FindControl(, conMenu_Manage_ThingAdd).Enabled Then
            Call thingAdd           '自动接单
        End If
    End If
End Sub

'Private Sub txtFind_Change()
'    If Not mobjIDCard Is Nothing Then
'        mobjIDCard.SetEnabled txtFind.Text = "" And Me.ActiveControl Is txtFind
'    End If
'End Sub

'Private Sub txtFind_GotFocus()
'    If txtFind.Tag = "" Then
'        Call zlControl.TxtSelAll(txtFind)
'    End If
'    txtFind.Tag = ""
'
'    If Not mobjIDCard Is Nothing Then
'        If txtFind.Text = "" Then mobjIDCard.SetEnabled True
'    End If
'End Sub

'Private Sub txtFind_KeyPress(KeyAscii As Integer)
'    '按回车
'    Dim blnCard As Boolean
'
'    '是否刷卡完成
'    blnCard = mintFindType = 0 And KeyAscii <> 8 And Len(txtFind.Text) = gbytCardLen - 1 And txtFind.SelLength <> Len(txtFind.Text)
'    If blnCard Or KeyAscii = 13 Then
'        If KeyAscii <> 13 Then
'            txtFind.Text = txtFind.Text & Chr(KeyAscii)
'            txtFind.SelStart = Len(txtFind.Text)
'        End If
'        KeyAscii = 0
'        Call ExecuteFindPati
'    Else
'        Select Case mintFindType
'            Case 0 '就诊卡
'                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
'                    KeyAscii = 0
'                Else
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                End If
'            Case 1 '门诊号
'                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'            Case 2 '挂号单
'                KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                If Not (txtFind.Text = "" Or txtFind.SelLength = Len(txtFind.Text)) _
'                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
'                    KeyAscii = 0
'                End If
'            Case 3 '姓名
'            Case 4 '身份证
'                KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'            Case 5 'IC卡
'        End Select
'    End If
'
'End Sub
'
'Private Sub txtFind_LostFocus()
'    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
'End Sub

Public Sub ExecuteTest(ByVal str流水号 As String, strGroupKey As String)
    '填写皮试结果
    Dim strResult As String
    Dim cnNew As ADODB.Connection, strUserName As String, lngDeptID As Long, blnEnd As Boolean, objGroup As Group, i As Integer
    Dim str结果 As String
    Dim lngErrNo As Long
    
    'If InStr(",(+),(-),免试,", "," & str结果 & ",") <= 0 Then Exit Sub
    
    If Val(Me.mobjRecord.Item(str流水号).执行分类) = 3 And Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行状态 <> 1 Then
        '未完成的皮试
        If Me.mobjRecord.Item(str流水号).Item(strGroupKey).皮试结果 <> "" Then
            If MsgBox("该操作将改变原有的皮试结果，是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If mbln皮试验证 And cnNew Is Nothing Then
            Set cnNew = New ADODB.Connection
            strUserName = zlDatabase.UserIdentify(Me, "填写皮试结果前，请您先输入用户名和密码进行身份验证。", glngSys, 1263, "确认执行完成", cnNew)
            If strUserName = "" Then Exit Sub
        End If
        
        lngDeptID = Me.cboDept.ItemData(Me.cboDept.ListIndex)
        Call Me.mobjRecord.Item(str流水号).Update(str流水号, strGroupKey, lngDeptID, lngErrNo)
        
        If lngErrNo <> 0 Then
'            lngErrNo_Out = lngErrNo
            Exit Sub
        End If
        
        Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行人 = IIf(strUserName = "", UserInfo.姓名, strUserName)
        
        Dim objPati As cPatient
        Dim objexecGroup As ExecutiveGroup
        Set objexecGroup = Me.mobjRecord.Item(str流水号)
        If objexecGroup.ExecuteTestFinish(strGroupKey, Me, mobjSquareCard, str结果) Then
            Me.mobjRecord.Item(str流水号).Item(strGroupKey).皮试结果 = str结果
            Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行状态 = 1
            Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行人 = IIf(strUserName = "", UserInfo.姓名, strUserName)
            
            Set objPati = patiList.Item(mstr挂号单)
            SaveOperLog mlngPreDept, objPati, MEDICAL, "流水号" & str流水号 & ",医嘱ID,发送号" & strGroupKey & "的记录,填写皮试结果为" & str结果
            
            '---- 本流水号下的所有医嘱已执行完，则放入结束队列
            blnEnd = True
            For i = 1 To Me.mobjRecord.Count
                For Each objGroup In Me.mobjRecord.Item(i)
                    If objGroup.执行状态 <> 1 Then blnEnd = False
                Next
            Next
            If blnEnd Then
                If Me.UpdateState("4-结束") Then
                    SaveOperLog mlngPreDept, objPati, QUEUE, "填写皮试结果后调整状态为4-结束"
                Else
                    SaveOperLog mlngPreDept, objPati, QUEUE, "填写皮试结果后未调整状态"
                End If
            End If
        End If
        
    End If
End Sub

Public Sub ExecComplt(ByVal str流水号 As String, ByVal strGroupKey As String)
    '输液医嘱完成功能
    
    Dim int执行状态 As Integer, lngDeptID As Long, blnEnd As Boolean, objGroup As Group, i As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim objPati As cPatient
    
    '0-未执行;1-完全执行;2-拒绝执行;3-正在执行

    int执行状态 = Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行状态
    lngDeptID = Me.cboDept.ItemData(Me.cboDept.ListIndex)
    
    If int执行状态 <> 1 And int执行状态 <> 2 And Val(Me.mobjRecord.Item(str流水号).执行分类) <> 3 Then
        
        If Me.mobjRecord.Item(str流水号).Item(strGroupKey).发送数次 = Me.mobjRecord.Item(str流水号).Item(strGroupKey).已执行数次 + Me.mobjRecord.Item(str流水号).Item(strGroupKey).本次数次 Then
            
            '如果当前是最后一次，或者 前几次还有未填写操作员的记录，则不完成。
            If Not CheckComplt(str流水号, strGroupKey) Then Exit Sub
                
            If Me.mobjRecord.Item(str流水号).ExecuteFinish(strGroupKey, lngDeptID, , Me, mobjSquareCard) Then
                Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行状态 = 1
                
                Set objPati = patiList.Item(mstr挂号单)
                SaveOperLog mlngPreDept, objPati, MEDICAL, "医嘱执行完成操作,流水号" & str流水号 & ",医嘱ID,发送号" & strGroupKey
                
                '---- 所有医嘱已执行完，则放入结束队列
                blnEnd = True
                For i = 1 To Me.mobjRecord.Count
                    For Each objGroup In Me.mobjRecord.Item(i)
                        If Not (objGroup.执行医嘱ID = Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行医嘱ID And objGroup.发送号 = Me.mobjRecord.Item(str流水号).Item(strGroupKey).发送号 And objGroup.上次执行时间 = Me.mobjRecord.Item(str流水号).Item(strGroupKey).上次执行时间) Then
                            If objGroup.执行状态 <> 1 Then blnEnd = False
                        End If
                    Next
                Next
                If blnEnd Then
                    If Me.UpdateState("4-结束") Then
                        SaveOperLog mlngPreDept, objPati, QUEUE, "医嘱完成后调整状态为4-结束"
                        
                        '2012-07-17 完成之后，自动清空座位占用 51193问题
                        strSQL = "select 编号 from 座位状况记录 where 病人ID=[1] and 科室ID=[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "执行完成", patiList.Item(mstr挂号单).病人ID, lngDeptID)
                        If Not rsTmp.EOF Then
                            strSQL = "Zl_座位状况记录_Clear(" & lngDeptID & ",'" & rsTmp!编号 & "')"
                            Call zlDatabase.ExecuteProcedure(strSQL, "执行完成")
                            SaveOperLog mlngPreDept, objPati, SEAT, "医嘱完成后清除占用座位" & rsTmp!编号
                        End If
                    Else
                        SaveOperLog mlngPreDept, objPati, QUEUE, "医嘱完成后未调整状态"
                    End If
                End If
            End If
        End If
    End If
    
End Sub

Public Function ExecStart(ByVal str流水号 As String, ByVal strGroupKey As String, Optional ByVal blnUndo As Boolean = False) As Boolean
'开始/撤消开始功能，主要是用于修改执行时间，填写/清空执行人
    
    Dim int执行状态 As Integer, strOper As String, strDate As String
    Dim objOutNur As OutNurse, strOutNurs As String, strSQL As String
    Dim objPati As cPatient
    
    Set objPati = patiList.Item(mstr挂号单)
    If blnUndo Then
        '撤消开始
        Call Me.mobjRecord.Item(str流水号).ExecStart(2, strGroupKey, Now, "")
        SaveOperLog mlngPreDept, objPati, MEDICAL, "医嘱撤消开始操作，流水号" & str流水号 & "，医嘱ID和发送号为" & strGroupKey
        ExecStart = True
    Else
        '开始
        int执行状态 = Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行状态
        ExecStart = False
        If Not (int执行状态 >= 1 And int执行状态 <= 2) Then
            strOper = Me.mobjRecord.Item(str流水号).Item(strGroupKey).执行人
            strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
            
            For Each objOutNur In ObjOutNurse
                strOutNurs = strOutNurs & "|" & objOutNur.姓名
            Next
            If Mid(strOutNurs, 1, 1) = "|" Then strOutNurs = Mid(strOutNurs, 2)
            
            If frmRecordStart.ShowSelect(strOutNurs, strDate, strOper) Then
                Call Me.mobjRecord.Item(str流水号).ExecStart(1, strGroupKey, CDate(strDate), strOper)
                
                SaveOperLog mlngPreDept, objPati, MEDICAL, "医嘱开始操作，执行人填为" & strOper & "，流水号为" & str流水号 & "，医嘱ID和发送号为" & strGroupKey
                
                Set objPati = patiList.Item(mstr挂号单)
                If Not objPati Is Nothing Then
                    'strSQL = "Zl_排队记录_Startend(1," & mlngPreDept & ",'" & mstr挂号单 & "',to_date('" & strDate & "','yyyy-MM-dd HH24:MI:SS'),'" & strOper & "')"
                    strSQL = "Zl_排队记录_Startend(1," & _
                                    mlngPreDept & "," & _
                                    objPati.病人ID & "," & _
                                    IIf(objPati.病人来源 = 1, "Null", "'" & objPati.挂号单 & "'") & "," & _
                                    IIf(objPati.病人来源 = 1, objPati.单据ID, "Null") & "," & _
                                    IIf(strOper = "", "Null", "'" & strOper & "'") & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    ExecStart = True
                End If
            End If
        End If
    End If
End Function

Private Sub txtInfo_Change()
    idkSelect.SetAutoReadCard Trim(txtInfo.Text) = ""
End Sub

'Private Sub txtInfo_Change()
'    If txtInfo.Enabled = False Then Exit Sub
'    If Not mobjIDCard Is Nothing Then
'        mobjIDCard.SetEnabled txtInfo.Text = "" And Me.ActiveControl Is txtInfo
'    End If
'    If txtInfo.Text = "" Then Call cboDate_Click
'End Sub

Private Sub txtInfo_GotFocus()
    Call zlControl.TxtSelAll(txtInfo)
    idkSelect.SetAutoReadCard Trim(txtInfo.Text) = ""
End Sub

Private Sub txtInfo_KeyPress(KeyAscii As Integer)
'按回车
    Dim strCard As String
    
    strCard = idkSelect.Cards(idkSelect.IDKind).名称

    If mblnReadCard Or KeyAscii = 13 Then
        Call cmdOk_Click
        Call zlControl.TxtSelAll(txtInfo)
        
        '定位页签
        If patiList.Count >= 1 Then
            Call SelectTabItem
        End If

    Else
        Select Case strCard
            Case "门诊号"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "挂号单"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtInfo.Text = "" Or txtInfo.SelLength = Len(txtInfo.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "身份证号", "二代身份证"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case Else
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
        End Select
    End If
    mblnReadCard = False
End Sub

Private Sub txtInfo_LostFocus()
    idkSelect.SetAutoReadCard False
End Sub

'Private Sub txtInfo_LostFocus()
'    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
'End Sub

Private Sub txtNo1_KeyPress(KeyAscii As Integer)
    '找到挂号单，执行配液
    If KeyAscii = vbKeyReturn Then
        Call FindAndExe(rptQueue1, txtNo1, 1)
    End If
End Sub

Private Sub FindAndExe(objRpt As ReportControl, ByVal strNoIn As String, ByVal intLiquidOrPut As Integer)

    '查找并执行 配液 或 穿刺
    Dim strNO As String, i As Integer
    strNO = GetFullNO(strNoIn, 12)  '12－挂号收据号
    
    For i = i To objRpt.Rows.Count - 1
        With objRpt.Rows(i)
            If Not .GroupRow Then
                If .Record(col_挂号单).Value = strNO Then Exit For
            End If
        End With
    Next

    If i <= objRpt.Rows.Count - 1 Then
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set objRpt.FocusedRow = objRpt.Rows(i)
        objRpt.SetFocus
        If intLiquidOrPut = 1 Then
            '配液并呼叫
            Call LiquidAndPlay
        Else
            '穿刺
            Call Puncture
        End If
    Else
        MsgBox "找不到符合条件的病人。", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtNo5_KeyPress(KeyAscii As Integer)
    '找到挂号单，执行配液
    If KeyAscii = vbKeyReturn Then
        Call FindAndExe(rptQueue5, txtNo5, 2)
    End If
End Sub

Private Function CheckComplt(ByVal str流水号 As String, ByVal strGroupKey As String) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngID As Long, lngSend As Long  '医嘱ID,发送号
    Dim objPati As cPatient
    
    On Error GoTo hErr
    
    CheckComplt = True
    lngID = Val(Split(strGroupKey, "_")(0))
    lngSend = Val(Split(strGroupKey, "_")(1))
    

    '是否有未填写操作员的记录，有则不能完成。
    strSQL = "select 执行时间 from 病人医嘱执行 where 执行人 Is Null and 医嘱ID=[1] and 发送号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID, lngSend)
    If Not rsTmp.EOF Then
        CheckComplt = False
        Set objPati = patiList.Item(mstr挂号单)
        SaveOperLog mlngPreDept, objPati, MEDICAL, "医嘱发送号为" & strGroupKey & "在" & Format(rsTmp!执行时间, "yyyy-MM-dd HH:mm:ss") & "的记录还未开始"
    End If
    Exit Function
hErr:
    CheckComplt = False
End Function

Private Sub FindCboIndex(ByVal objCbo As Object, ByVal lngData As Long, Optional ByVal blnKeep As Boolean)
'功能：由项目值查找ComboBox的项目索引
'参数：Keep=如果未匹配，是否保持原索引
    Dim i As Integer
    
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not blnKeep Then objCbo.ListIndex = -1
End Sub

Private Sub SelectTabItem()
'功能：查找定位到病人当前状态页签

    Dim i As Integer
    Dim strTag As String

    strTag = patiList.Item(1).排队状态
    If InStr(strTag, "-") > 0 Then strTag = Mid(strTag, InStr(strTag, "-") + 1)
    
    '统一名称处理；排队状态=结束，而页签名称=已结束
    If strTag = "结束" Then strTag = "已结束"

'    Dim strSQL As String
'    Dim lngID As Long
'    Dim rsTmp As ADODB.Recordset
'    Dim strTag As String
'
'    lngID = patiList.Item(1).病人ID
'
    On Error GoTo errHandle
'
'    strSQL = "Select decode(a.状态, 1, '待配液', 5, '待穿刺', 6, '执行中', 7, '已结束', '') 状态 " & _
'             "From 排队记录 A, 病人信息 B Where a.病人id = b.病人id And b.病人id = [1] "
'    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, "获取病人排队记录的状态", lngID)
'    If rsTmp.EOF = False Then strTag = zlcommfun.NVL(rsTmp!状态)
'    rsTmp.Close
'
    For i = 1 To tbcList.ItemCount - 1
        If tbcList.Item(i).Tag = strTag And tbcList.Item(i).Visible Then
            tbcList.Item(i).Selected = True
            Exit For
        End If
    Next

    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Function FuncThingAudit(ByVal str流水号 As String, strGroupKey As String) As Boolean
'功能：核对
    Dim strSQL As String
    Dim str核对人 As String

    With mobjRecord.Item(CStr(str流水号)).Item(strGroupKey)
        If .核对人 <> "" Then
            MsgBox "该医嘱还已经核对，不能再次核对。", vbInformation, gstrSysName
            Exit Function
        End If
        If .执行人 = "" Then
            MsgBox "该医嘱还未登记执行人，不能核对。", vbInformation, gstrSysName
            Exit Function
        End If
        str核对人 = zlDatabase.UserIdentifyByUser(Me, "在核对执行情况前，请您先输入用户名和密码进行身份验证。", glngSys, 1263, "执行情况登记", , True)
        If str核对人 = "" Then Exit Function
        
        If str核对人 = .执行人 Then
            MsgBox "执行人不能和审核人相同，不能核对。", vbInformation, gstrSysName
            Exit Function
        End If

        On Error GoTo errH
        strSQL = "Zl_病人医嘱核对_Insert(" & Val(.执行医嘱ID) & "," & Val(.发送号) & ",'" & str核对人 & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "医嘱核对")
        .核对人 = str核对人
        FuncThingAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function FuncThingDelAudit(ByVal str流水号 As String, strGroupKey As String) As Boolean
'功能：取消核对
    Dim strSQL As String
    Dim str核对人 As String

    With mobjRecord.Item(CStr(str流水号)).Item(strGroupKey)
        If .核对人 = "" Then
            MsgBox "该医嘱还未进行核对，不能取消。", vbInformation, gstrSysName
            Exit Function
        End If

        If .核对人 <> UserInfo.姓名 Then
            str核对人 = zlDatabase.UserIdentifyByUser(Me, "在取消核对前，请您先输入用户名和密码进行身份验证。", glngSys, 1263, "执行情况登记", , True)
            If str核对人 = "" Then Exit Function
            If str核对人 <> .核对人 Then
                MsgBox "只能取消自己核对的医嘱，当前医嘱核对人是""" & .核对人 & """", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If MsgBox("你确定要取消核对吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
        End If
        On Error GoTo errH
        
        strSQL = "Zl_病人医嘱核对_Delete(" & Val(.执行医嘱ID) & "," & Val(.发送号) & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "取消医嘱核对")
        .核对人 = ""
        FuncThingDelAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

