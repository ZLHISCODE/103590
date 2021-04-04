VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "codejock.commandbars.9600.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5C493D4E-FD57-4FF4-9BA4-C6C670BFF9A7}#71.0#0"; "zl9PacsControl.ocx"
Object = "*\A..\..\0_zlhisPro\common vbp\zlIDKind.vbp"
Begin VB.Form frmPacsMain 
   Caption         =   "影像工作站"
   ClientHeight    =   7605
   ClientLeft      =   8595
   ClientTop       =   975
   ClientWidth     =   11400
   Icon            =   "frmPacsMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timerVideoEvent 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   9015
      Top             =   165
   End
   Begin VB.Timer timerCapture 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8505
      Top             =   135
   End
   Begin VB.PictureBox picTemp 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   4815
      ScaleHeight     =   585
      ScaleWidth      =   825
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Timer timerOperHint 
      Interval        =   500
      Left            =   7920
      Tag             =   "0"
      Top             =   120
   End
   Begin VB.PictureBox picWindow 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   5400
      ScaleHeight     =   4575
      ScaleWidth      =   5535
      TabIndex        =   11
      Top             =   2160
      Width           =   5535
      Begin zl9PacsControl.TranControl tcDisable 
         Height          =   975
         Left            =   4560
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1720
         AlphaValue      =   25
      End
      Begin VB.PictureBox picLoadState 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   1320
         ScaleHeight     =   1095
         ScaleWidth      =   3855
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   3855
         Begin VB.PictureBox picSmile 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   360
            Left            =   240
            Picture         =   "frmPacsMain.frx":1CFA
            ScaleHeight     =   360
            ScaleWidth      =   360
            TabIndex        =   25
            Top             =   240
            Width           =   360
         End
         Begin VB.Label labLoadState 
            Caption         =   " 正在加载工作模块，请耐心等待..."
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.PictureBox picReportContainer 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2055
         Left            =   3720
         ScaleHeight     =   2055
         ScaleWidth      =   1815
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin XtremeSuiteControls.TabControl TabWindow 
         Height          =   2415
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   4260
         _StockProps     =   64
      End
   End
   Begin DicomObjects.DicomViewer dcmRelateViewer 
      Height          =   1095
      Left            =   5880
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
      _Version        =   262147
      _ExtentX        =   4471
      _ExtentY        =   1931
      _StockProps     =   35
   End
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   7320
      Top             =   120
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7245
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPacsMain.frx":2771
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Left            =   6675
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":3005
            Key             =   "紧急"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":359F
            Key             =   "住院"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":3E79
            Key             =   "阳性"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":3FD3
            Key             =   "影像"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":474D
            Key             =   "绿色通道"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":48A7
            Key             =   "路径"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":4E41
            Key             =   "无费"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":51DB
            Key             =   "欠费"
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":5575
            Key             =   "收费"
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":590F
            Key             =   "病理"
            Object.Tag             =   "10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":6021
            Key             =   "补费"
            Object.Tag             =   "11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":63BB
            Key             =   "危急"
            Object.Tag             =   "12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":6755
            Key             =   "检查技师"
            Object.Tag             =   "13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":6E4F
            Key             =   "记费"
            Object.Tag             =   "14"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   6060
      Top             =   90
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":71E9
            Key             =   "复选留空"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":7783
            Key             =   "复选选中"
            Object.Tag             =   "90001"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":7D1D
            Key             =   "定位"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPacsMain.frx":80AF
            Key             =   "查找"
            Object.Tag             =   "4"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6540
      Left            =   45
      ScaleHeight     =   6540
      ScaleWidth      =   4500
      TabIndex        =   1
      Top             =   555
      Width           =   4495
      Begin VB.PictureBox picTag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picExeState 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         ScaleHeight     =   375
         ScaleWidth      =   3975
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   3975
         Begin VB.OptionButton optNeed 
            Caption         =   "需执行"
            Height          =   180
            Left            =   120
            TabIndex        =   19
            Top             =   50
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optAccept 
            Caption         =   "已接受"
            Height          =   180
            Left            =   1080
            TabIndex        =   18
            Top             =   50
            Width           =   975
         End
         Begin VB.OptionButton optFinal 
            Caption         =   "已执行"
            Height          =   180
            Left            =   2040
            TabIndex        =   17
            Top             =   50
            Width           =   975
         End
         Begin VB.OptionButton optAll 
            Caption         =   "所有"
            Height          =   180
            Left            =   3000
            TabIndex        =   16
            Top             =   50
            Width           =   975
         End
      End
      Begin VB.PictureBox picAppend 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2775
         Left            =   240
         ScaleHeight     =   2775
         ScaleWidth      =   3945
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3720
         Width           =   3945
         Begin VB.ComboBox cboTimes 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   120
            Width           =   2235
         End
         Begin VB.TextBox txtAppend 
            BackColor       =   &H00FDD6C6&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1500
            Left            =   10
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   5
            Top             =   1260
            Width           =   3920
         End
         Begin VB.Label labStudyNum 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "检查号："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label labHistory 
            Caption         =   "历史检查："
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lblCash 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "收"
            BeginProperty Font 
               Name            =   "黑体"
               Size            =   21.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   465
            Left            =   3360
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbl个人信息 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名：      性别：    年龄：  "
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label lbl检查信息 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   270
         End
      End
      Begin VB.PictureBox PicLine 
         BorderStyle     =   0  'None
         Height          =   90
         Left            =   240
         MousePointer    =   7  'Size N S
         ScaleHeight     =   90
         ScaleWidth      =   3975
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   3600
         Width           =   3975
      End
      Begin zl9PACSWork.ucFlexGrid ufgStudyList 
         Height          =   2415
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4260
         DefaultCols     =   ""
         HeadCheckValue  =   1
         IsCopyAdoMode   =   0   'False
         IsEjectConfig   =   -1  'True
         Editable        =   0
         ReadOnly        =   -1  'True
         IsShowPopupMenu =   0   'False
         HeadFontCharset =   134
         HeadFontWeight  =   400
         HeadColor       =   0
         DataFontCharset =   134
         DataFontWeight  =   400
         DataColor       =   -2147483640
         GridLineColor   =   14737632
         ExtendLastCol   =   -1  'True
      End
      Begin XtremeSuiteControls.TabControl tabFilter 
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   4005
         _Version        =   589884
         _ExtentX        =   7064
         _ExtentY        =   661
         _StockProps     =   64
      End
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   330
         Left            =   360
         TabIndex        =   28
         Top             =   0
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmPacsMain.frx":8441
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
         CaptionAlignment=   0
         ShowPropertySet =   -1  'True
         DefaultCardType =   "就诊卡"
         IDkindBorderStyle=   1
         IDKindWidth     =   1330
         FindPatiShowName=   0   'False
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
      Begin XtremeCommandBars.CommandBars cbrdock 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Bindings        =   "frmPacsMain.frx":84C8
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPacsMain.frx":84DC
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPacsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

#Const DebugImmediately = False

Private Const M_BLN_ALL_FUNCTIONS_OPEN As Boolean = True
Private Const M_STR_MODULE_MENU_TAG As String = "Main"

Private Const mintCur业务类型 As Integer = 1 '当前系统操作的业务类型，用于排队叫号使用



'公共列
'根据系统不同，“[------]”所代表的内容将被有差异的列替换
Private Const M_STR_PUBLIC_COLS = "|路径>路径状态,w400|紧急>紧急标志,headimg1,w300|来源,headimg2,w400" & _
                        "|收费,headimg9,w300|危急,headimg12,w800|阳性,headimg3,w300|姓名,btn,txtleft,w1200,uncfg" & _
                        "|申请单>申请单医嘱,w1100|检查过程>[placecol],w800|执行状态,hide,uncfg|性别,w450|年龄,w450|标识号,w1400|[------]|报告质量,w800|医嘱内容,w2400" & _
                        "|部位方法>[placecol],w1400|报到时间,w1800,shortdatetime|申请时间,w1800,shortdatetime|开嘱医生,w800|身高,hide,w450" & _
                        "|体重,hide,w450|婴儿,w450|登记人,w800|报到人,w800|完成人,w800|报告操作,w800|绿色通道,hide,uncfg" & _
                        "|报告打印,w800|报告人,w800|复核人,w800|采图时间,w1800,shortdatetime|随访描述,w2400|病人ID,hide,uncfg" & _
                        "|主页ID,hide,uncfg|挂号单,hide|病人科室ID,hide,uncfg|医嘱ID,key,hide,w1200|发送号,hide,uncfg" & _
                        "|检查UID,hide,uncfg|检查状态>检查过程,hide,uncfg|NO,hide,uncfg|记录性质,hide,uncfg|转出,hide,uncfg" & _
                        "|床号>当前床号,hide|当前病区ID,hide,uncfg|报告发放,w800|诊断分类,w800|病人科室,w800|关联ID,hide,uncfg" & _
                        "|就诊卡号,w800|单据号,w800|身份证号,w800|采样时间,hide,uncfg,shortdatetime|图像位置,hide,uncfg|是否技师确认,hide,uncfg|"

Private Const M_STR_PUBLIC_COLS_NEW = "|路径>路径状态,w400|紧急>紧急标志,headimg1,w300|来源,headimg2,w400" & _
                        "|收费,headimg9,w300|危急,headimg12,w800|姓名,btn,txtleft,w1200,uncfg" & _
                        "|申请单>申请单医嘱,w1100|检查过程>[placecol],w800|执行状态,hide,uncfg|性别,w450|年龄,w450|标识号,w1400|[------]|医嘱内容,w2400" & _
                        "|部位方法>[placecol],w1400|报到时间,w1800,shortdatetime|申请时间,w1800,shortdatetime|开嘱医生,w800|身高,hide,w450" & _
                        "|体重,hide,w450|婴儿,w450|登记人,w800|报到人,w800|完成人,w800|绿色通道,hide,uncfg" & _
                        "|采图时间,w1800,shortdatetime|随访描述,w2400|病人ID,hide,uncfg" & _
                        "|主页ID,hide,uncfg|挂号单,hide|病人科室ID,hide,uncfg|医嘱ID,key,hide,w1200|发送号,hide,uncfg" & _
                        "|检查UID,hide,uncfg|检查状态>检查过程,hide,uncfg|NO,hide,uncfg|记录性质,hide,uncfg|转出,hide,uncfg" & _
                        "|床号>当前床号,hide|当前病区ID,hide,uncfg|诊断分类,w800|病人科室,w800|关联ID,hide,uncfg" & _
                        "|就诊卡号,w800|单据号,w800|身份证号,w800|采样时间,hide,uncfg,shortdatetime|图像位置,hide,uncfg|是否技师确认,hide,uncfg|"

'病理
Private Const M_STR_PATHOL_COLS = "病理号,w1400|质量>综合质量,w280|病理执行状态,w1400|检查类别|核收情况,w1200|取材过程,hide,uncfg|制片过程,hide,uncfg|免疫过程,hide,uncfg|分子过程,hide,uncfg|特染过程,hide,uncfg|"
'医技
Private Const M_STR_IMAGES_COLS = "检查号,w1400|影像类别|影像质量,w280|符合情况,w280|执行间,w600|电子胶片>是否电子胶片,W600|胶片打印>是否打印,w800|检查技师,w800|检查技师二,w1000|胶片发放>发放胶片,w800|执行科室ID,hide,uncfg"
'采集
Private Const M_STR_CAPTOR_COLS = "检查号,w1400|影像类别|影像质量,w280|符合情况,w280|执行间,w600|检查技师,w800|检查技师二,w1000"


'当没有数据时，使用此提示信息
Private Const M_STR_HINT_NoSelectData As String = "请选择需要执行的检查数据。"

'根据不同系统，“[------]”将被替换为“检查号”或者“病理号”
Private Const CONST_STR_LOCAL_CARD_TYPE As String = "姓名;就诊卡;标识号;单据号;[------];身份证号;健康号;IC卡号;"
Private Const CONST_STR_FIND_CARD_TYPE As String = "姓名;就诊卡;门诊号;住院号;单据号;[------];身份证号;健康号;IC卡号;"

Private Enum TLocateFindType
    lftLocate = 0
    lftFind = 1
End Enum


'当前医嘱信息
Private Type TAdviceInf
    lngPatID As Long                '1 病人ID
    lngPageID As Long               '2 主页ID
    lngAdviceId As Long             '3 医嘱ID
    lngSendNo As Long               '4 发送号
    strPatientName As String        '5 病人姓名
    
    lngPatDept As Long              '6 病人所属科室
    strRegNo As String              '7 挂号单
    lngRegId As Long                '8 挂号id
    intMoved As Integer             '9 是否转出
    intState As Integer             '10 检查状态
    intStep As Integer              '11 检查过程
    lngUnit As Long                 '12 病区ID
    strStudyUID As String           '13 检查UID
    blnCanPrint As Boolean          '14 是否能够打印
    blnIsInsidePatient As Boolean   '15 是否门诊或住院病人
    lngExeDepartmentId As Long      '16 执行部门ID
    strDoDoctor As String           '17 检查技师
    strExeRoom As String            '18 执行间
    lngPatientFrom As Long          '19 病人来源
    
    strStudyNum As String           '20 检查号
    strBedNum As String             '21 床号
    lngMarkNum As Double            '22 标志号
    lngBaby As Long                 '23 婴儿
    strPatientDepartment As String  '24 病人科室名称
    
    strReportDoctor As String       '25 报告人
    strReportOperation As String    '26 报告操作
    lngLinkId As Long               '27 关联ID
    strImgType As String            '28 影像类别
    intImageLocation As Integer     '29 PACS影像所在的位置，0在中联PACS；1在新网PACS
    strPatientType As String        '30 病人类型
End Type


'过滤条件变量
Private Type Type_SQLCondition
    开始时间 As Date
    结束时间 As Date
    时间类型 As Integer                                 '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
    单据号 As String
    门诊号 As Double
    健康号 As String
    住院号 As Double
    就诊卡 As String
    姓名 As String
    性别 As String
    开始年龄 As Long
    结束年龄 As Long
    年龄条件 As String
    检查号 As Variant
    身份证  As String
    IC卡 As String
    病人科室 As Long
    标本部位 As String
    诊断医生 As String
    审核医生 As String
    疾病诊断 As String
    报告内容 As String
    结果阳性 As Integer
    影像质量 As String
    检查技师 As String
    检查过程 As String
    影像类别 As String
    检查所见 As String
    诊断意见 As String
    建议 As String
    随访 As String
    病人ID As Long
End Type

'系统参数类型定义
Private Type TSystemPar

    '本地参数
    strFirstTab As String                               '首次显示的页面
    bln直接检查 As Boolean                              '登记后直接进入检查
    blnWriteCapDoctor As Boolean                        '是否在采集图像后，自动把当前用户填写为检查技师
    blnAutoOpenReport As Boolean                        '开始检查自动打开报告
    blnNoShowCancel As Boolean                          '不显示取消的检查
    blnPatTrack As Boolean                              '是否对进病人进行跟踪
    strLocalRoom As String                              '本机执行间名称
    
    '流程参数
    blnFinishCommit As Boolean                          '无报告完成里,是否无需再次确认
    blnCompleteCommit As Boolean                        '审核后无需再次确认
    blnIgnoreResult As Boolean                          '忽略阴阳性 '=true 忽略
    
    blnReportWithImage As Boolean                       '有图像才能写报告，无图像不可写报告
    blnReportWithResult As Boolean                      '有阴阳性结果才能写报告
    blnLocalizerBackward As Boolean                     '定位片后置
    
    blnPrintCommit As Boolean                           '打印后直接完成
    blnCanPrint As Boolean                              '平诊需要审核才能打印 =true
    lngBeforeDays As Long                               '默认查询的天数
    lngRefreshInterval As Long                          '病人列表自动刷新间隔
    blnUseQueue As Boolean                              '是否启用排队叫号
    blnSynStudylist As Boolean                          '排队叫号时，点击排队列表或呼叫列表数据后，是否同步定位到检查列表
    blnAutoInQueue As Boolean                           '启用排队叫号后，是否自动入队
    
    blnRelatingPatient As Boolean                       '是否启用关联病人
    lngQueueWay As Long                                 '判断方式，0按执行间名称排队，1按科室名称排队
'    lngSameTime As Long                                 '发放方式，0报告和胶片分别发放 1 报告和胶片同时发放
    
    lngCriticalValues As Long                           '危急值
    lngConformDetermine As Long                         '符合情况
    strImageLevel As String                             '影像质量等级串
    strReportLevel As String                            '报告质量等级串
    lngImageLevel As Long                               '影像质量判定
    lngReportLevel As Long                              '报告质量判定
    lngHintType As Long                                 '诊断结果提示类型
    
    blnIsPetitionScan As Boolean                        '是否启用申请单扫描
    blnChangeUser As Boolean                            '是否启用用户交换
    blnSwitchUser As Boolean                            '是否启用用户切换
    
    lngMoneyExeModle As Long                            '费用执行模式
    
    lngListColorMark As Long                            '为0时标记列表前景色，为1时标记列表背景色
    blnNameColColorCfg As Boolean                       '是否根据病人类型设置列表姓名列颜色
    blnOrdinaryNameColColorCfg As Boolean               '缺省类型的病人是否根据病人类型设置姓名颜色
    
    blnAutoSendWorkList As Boolean                      '是否报道时自动发送WorkList
    
    '状态提醒
    lngEnregAfterTimeLen As Long                        '登记后提醒
    lngCheckInAfterTimeLen As Long                      '报到后提醒
    lngStudyAfterTimeLen As Long                        '检查后提醒
    lngReportAfterTimeLen As Long                       '报告后提醒
    lngAuditAfterTimeLen As Long                        '审核后提醒
    
    blnAutoPrint As Boolean    '报到后自动打印申请单
    
    blnShowImgAfterReport As Boolean                    '报告时观片
End Type


'视频采集事件信息
Private Type TVideoEventInf
    vetEventType As TVideoEventType
    lngAdviceId As Long
    lngSendNo As Long
    strOtherInf As String
End Type

'视频采集消息定义
Private Type TCaptureMsgInf
    lngMsg As Long
    lngVirtualKey As Long
    lngScanKey As Long
    lngFlags As Long
End Type


'ID_查找方式+100之后保留7个是作为查找方式选择的
'ID_影像类别之后保留40个号码作为影像类别，从4021-4060
Private Enum FilterID
    ID_来源 = 4000: ID_门诊 = 4001: ID_住院 = 4002: ID_体检 = 4003: ID_外诊 = 4004
    ID_费用 = 4005: ID_已缴 = 4006: ID_未缴 = 4007: ID_补缴 = 4008: ID_无费 = 4009: ID_记账 = 4018
    ID_状态 = 4010: ID_登记 = 4011: ID_报到 = 4012: ID_检查 = 4013: ID_报告 = 4014: ID_审核 = 4015: ID_驳回 = 4016: ID_完成 = 4017
    ID_查找值 = 4020: ID_开始查找 = 4021: ID_本次住院 = 4022: ID_查找方式 = 4023
    
    ID_影像类别 = 4030
    
    ID_病理类别 = 4100
    ID_病理类别_常规 = 4101: ID_病理类别_冰冻 = 4102: ID_病理类别_细胞 = 4103: ID_病理类别_尸检 = 4104: ID_病理类别_会诊 = 4105: ID_病理类别_快速石蜡 = 4106
        
    ID_影像执行间 = 4110
    
    ID_检查部位 = 4210 '4210--4500
End Enum

Private mblncmd门诊 As Boolean, mblncmd住院 As Boolean, mblncmd体检 As Boolean, mblncmd外诊 As Boolean
Private mblncmd已缴 As Boolean, mblncmd未缴 As Boolean, mblncmd补缴 As Boolean, mblncmd无费 As Boolean, mblncmd记账 As Boolean
Private mblncmd登记 As Boolean, mblncmd报到 As Boolean, mblncmd检查 As Boolean, mblncmd报告 As Boolean
Private mblncmd驳回 As Boolean, mblncmd审核 As Boolean, mblncmd完成 As Boolean

Private mblncmd本次 As Boolean


Private mblncmd常规 As Boolean
Private mblncmd快速石蜡 As Boolean
Private mblncmd细胞 As Boolean
Private mblncmd冰冻 As Boolean
Private mblncmd尸检 As Boolean
Private mblncmd会诊 As Boolean


Private mintcmd影像类别 As Integer      '0表示没有选择影像类别，其他数字表示选择的影像类别的数量
Private mblncmd影像类别() As Boolean    '保存当前选择的影像类别是否被选择

Private mintcmd影像执行间 As Integer    '已选择的需要过滤的影像执行间数量，只有为0时，才不需要根据执行间过滤
Private mblncmd影像执行间() As Boolean

Private mstrcmd部位分组 As String
Private mstrcmd部位 As String

Private mstrFirstTab As String '首次显示的页面

Private mintToolBarWriteReg As Integer        '工具栏注册表状态值


Private mstrPrivs As String, mlngModule As Long              '模块号，本模块权限


'子窗体对像
Private WithEvents mobjEvent As clsEvent            '事件处理对象
Attribute mobjEvent.VB_VarHelpID = -1
Private WithEvents mfrmRISRequest As frmRISRequest
Attribute mfrmRISRequest.VB_VarHelpID = -1
'消息处理中心
Private WithEvents mobjMsgCenter As clsPacsMsgProcess
Attribute mobjMsgCenter.VB_VarHelpID = -1

'工作模块的数据刷新模式分三种情况，
'1.工作模块只要存在，强制对其中的数据进行刷新
'2.工作模块在显示时，才对其中的数据进行刷新
'3.工作模块在相关数据变化时且显示的模块是当前模块，才对其中的数据进行刷新

Private mfrmWork_PacsImg As frmWork_Image           '影像子窗体
Attribute mfrmWork_PacsImg.VB_VarHelpID = -1
Private mobjWork_Pathol As clsWorkModule_Pathol     '病理相关模块
Private mobjWork_His As clsWorkModule_His           'HIS相关模块

Private mobjWork_ActiveVideo As Object  ' zl9PacsCapture.clsPacsCapture  '视频采集模块
Attribute mobjWork_ActiveVideo.VB_VarHelpID = -1
Private WithEvents mobjWork_Report As clsWorkModule_Report     '报告模块
Attribute mobjWork_Report.VB_VarHelpID = -1
Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer            '观片站对象
Attribute mobjPacsCore.VB_VarHelpID = -1
Private WithEvents mobjQueue As frmWork_Queue  'zlQueueManage.clsQueueManage          '排队叫号
Attribute mobjQueue.VB_VarHelpID = -1

Private mfrmPatholSpecimen As frmPatholSpecimen              '标本核收

Private mfrmPACSFilter As frmPACSFilter

Private mclsCISKernel As clsCISKernel '只用了查看检查单方法

'窗口变量
Private mlngCur科室ID As Long                               '当前科室ID
Private mstrCur科室 As String                               '当前科室 编码-名称
Private mstrCanUse科室 As String                            '当前可用科室  ID_编码-名称
Private mlngFilterTab As Long                               '过滤tab页
Private mblnInitOk As Boolean, mblnvsRefresh As Boolean     '初始化完成,装载表格
Private mblnLoadSubFrom As Boolean                          '是否正在加载子窗口
Private mblnAllDepts As Boolean                             '是否选择全部科室
Private mstrCanUse科室IDs As String                         '当前可用的科室ID串，用“，”分隔，可以直接作为SQL查询条件
Private mlngSortCol As Long                                 '病人列表中，当前进行排序的列
Private mintSortOrder As Integer                            '病人列表中，当前进行排序的方式
Private mblnMenuDownState As Boolean                        '避免双击工具栏产生错误
Private mblnIsLoadPatholModule As Boolean                   '是否载入了病理模块
Private mblnFormLoadState As Boolean

Private mblnIsPrintMode As Boolean                          '是否是清单打印

Private mstrDefaultPatientType As String                    '缺省病人类型
Private mlngOldAdviceId As Long                             '前一次选择的检查的医嘱ID

Private mstrRPTExecutor As String                           '保存选择的报告人
Private mrtReportType As ReportType

'流程控制变量
Private mSysPar As TSystemPar                               '系统参数

'Private mlngOldSameTime As Long                             '切换科室前当前科室发放方式，0报告和胶片分别发放 1 报告和胶片同时发放
Private mblnObserve As Boolean                              '是否有观片基本权限   true是  false否
Private mblnSetXWParam As Boolean                           '是否有“影像设备目录”权限，如果有，则可以设置新网PACS的参数
Private mintImgCount As Integer                             '已扫描申请单数量

Private mAstr队列名称() As String       '队列名称，执行间的名称

Private WithEvents mobjCaptureHot As zl9PacsControl.clsHookKey
Attribute mobjCaptureHot.VB_VarHelpID = -1
Private mVideoEventInf As TVideoEventInf
Private mblnUseActivexCapture As Boolean                            '是否使用ActivexExe的视频采集方式
Private mstrCaptureHot As String                                    '采集热键定义
Private mstrCaptureAfterHot As String                               '后台采集热键定义
Private mstrCaptureAfterTagHot As String                            '标记更新热键定义
Private mCaptureMsg As TCaptureMsgInf
Private mobjSquareCard As Object

'本机本地参数
Private mstrSelQueueRooms As String                         '只处理执行间内的病人
Private mstrAllQueueRooms As String

Private mblnMoved As Boolean                                '当前时间段内是否被转移过
Private mstrWorkModule As String

Private mblnPopChangGuiWindow As Boolean
Private mblnPopBingDongWindow As Boolean
Private mblnPopXiBaoWindow As Boolean
Private mblnPopHuiZhenWindow As Boolean
Private mblnPopShiJianWindow As Boolean
Private mblnPopKuaiShuWindow As Boolean

Private SQLCondition As Type_SQLCondition

Private mblnAssignment As Boolean
Private mstrFindWay As String
Private mstrLocateWay As String
Private mlngLocateFindType As Long
Private mstrAllExamineRoomCfg As String    '所有科室执行间选择情况
Private mstrCurExamineRoomCfg As String    '当前科室执行间选择情况

Private mbytFontSize As Byte  '字体大小
Private mbyrFontState As Byte '字体状态，用于判断是否调整控件位置


Private mcurAdviceInf As TAdviceInf          '保存从检查列表或者历史列表中选择的当前检查信息
Private mListAdviceInf As TAdviceInf         '只保存从检查列表中选择的检查信息


'历史记录的显示
Private mblnIsHistory As Boolean


'双用户登录
Private mcnOracleHIS As New ADODB.Connection    '记录HIS导航台登陆时使用的数据库联接串
Private mstrUserNameHIS As String               '记录HIS导航台登陆时使用的用户名
Private mstrUserIDHIS As String                 '记录HIS导航台登录时使用的用户ID
Private mstrUserNameNew As String               '记录双用户登陆的第二个用户名
Private mstrUserIDNew As String                 '记录双用户登录的第二个用户ID
Private mblnCnOracleIsHIS As Boolean            '当前数据库联接是否HIS导航台的连接
Private mintChangeUserState As Integer          '记录用户交换的情况。1- 统一；2-交换

'收藏功能
Private mlngShareFatherID As Long
Private mlngCollectionFatherID As Long

Private mblnIsIntegratedQuery As Boolean        '是否在进行综合查询
Private mlngDefQuerySchemeId As Long            '默认查询方案id
Private mlngSysQuerySchemeId As Long            '系统查询方案id
Private blnIsLoading As Boolean

Dim mlngTempCharged As Long

Private mblnIsCallModuleRefresh As Boolean          '是否调用模块刷新操作
Private mblnAutoRefreshList As Boolean          '是否自动刷新检查列表


Public Sub ShowStation(ByVal lngModule As Long, owner As Object)
    
    mblnInitOk = False
    mblnLoadSubFrom = False
    mlngModule = lngModule
    mblnUseActivexCapture = False
    mblnAutoRefreshList = False
    mblnIsIntegratedQuery = False
    
    '创建卡结算部件
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    
    '初始化卡结算部件
    mobjSquareCard.zlInitComponents Me, mlngModule, glngSys, gstrDBUser, gcnOracle
    
    PatiIdentify.zlInit Me, glngSys, mlngModule, gcnOracle, gstrDBUser, mobjSquareCard, InitCardType(Replace(IIf(mlngLocateFindType = TLocateFindType.lftLocate, CONST_STR_LOCAL_CARD_TYPE, CONST_STR_FIND_CARD_TYPE), "[------]", GetStudyNumberDisplayName))

    If lngModule = G_LNG_VIDEOSTATION_MODULE Or lngModule = G_LNG_PATHOLSYS_NUM Then
        mblnUseActivexCapture = GetSetting("ZLSOFT", "公共模块", "UseActiveVideo", "1")
        Call SaveSetting("ZLSOFT", "公共模块", "UseActiveVideo", mblnUseActivexCapture)
    End If
    
    Call WriteLog("ShowStation -> Step 1：进入影像主窗口初始化流程。")
    
    If Not mblnFormLoadState Then Call InitForm
    
    Call WriteLog("ShowStation -> Step 2")
    
    '先显示出当前系统窗体
    Me.Show , owner
    If Me.WindowState = 1 Then Me.WindowState = 0
    
    DoEvents
    
    Call WriteLog("ShowStation -> Step 3：初始化窗口子模块。")
    '加载所需的工作模块
    Call Me.InitSubForm
    

    DoEvents
    
    Call WriteLog("ShowStation -> Step 4：配置显示子模块。")
    
    If Not TabWindow.Selected Is Nothing Then
        Call ConfigSubForm(TabWindow.Selected)
    End If
    
    mblnInitOk = True
    
    Call WriteLog("ShowStation -> Step 5：刷新数据列表。")
    
    '刷新检查数据
    Call Me.RefreshList
    
    DoEvents
    
    Call WriteLog("ShowStation -> Step 6：创建模块菜单。")
    '创建模块菜单
    Call CreateWorkModuleMenu
    
    Call WriteLog("ShowStation -> Step 7：结束影像主窗口初始化流程。")
End Sub


Private Sub Menu_File_Excel_click()
'功能:将数据复制到可打印的对象，调用打印
'参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
'       lngSelectedRow,记录调用打印部件前的选中行，在清单关闭后恢复
On Error GoTo ErrHandle
    Dim bytMode As Byte
    Dim lngSelectedRow As Long
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = ufgStudyList.DataGrid
    objPrint.Title.Text = "检查病人清单"
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & zlDatabase.Currentdate())
    Call objPrint.BelowAppRows.Add(objAppRow)

    '给 是否是打印清单参数赋值
    mblnIsPrintMode = True
    '得到打印清单前的当前选中行
    lngSelectedRow = ufgStudyList.SelectionRow
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    
    '打印货预览结束后 恢复选中行
    ufgStudyList.DataGrid.Row = lngSelectedRow
    mblnIsPrintMode = False
    
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_RichEPR(ByVal cbrID As Long)
'自动打开报告编辑器，同时处理了PACS报告编辑器和电子病历编辑器
On Error GoTo ErrHandle
    Dim cbrControl As CommandBarControl, i As Long
    
    '如果没有选择行数据，则直接退出执行
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    '报告页面不可见时不执行任何操作
    If TabWindow.Selected.tag <> "报告填写" Then
        For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
            If TabWindow(i).tag = "报告填写" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
        Next
        If TabWindow.Selected.tag <> "报告填写" Then Exit Sub
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If
    
    '找到报告页面，再打开这个报告页面
    With ufgStudyList
        '刷新嵌入页面内容
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.zlUpdateAdviceInf(mListAdviceInf.lngAdviceId, mListAdviceInf.lngPatID, mListAdviceInf.lngSendNo, mListAdviceInf.intStep, mListAdviceInf.intMoved = 1)
            Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, ufgStudyList, mblnIsHistory, mListAdviceInf.blnCanPrint, mListAdviceInf.strDoDoctor, mListAdviceInf.strStudyUID)
            
            Call mobjWork_Report.zlRefreshFace
        End If
    End With
    
    '判断按键可用性
    Set cbrControl = Me.cbrMain.FindControl(, conMenu_PacsReport_Open + 1000000)
    
    If cbrControl Is Nothing Then
        Set cbrControl = Me.cbrMain.FindControl(, cbrID + 1000000)
        If cbrControl Is Nothing Then Exit Sub
    End If
    
    Call cbrMain_Update(cbrControl)
    If cbrControl.Enabled = False Then Exit Sub
        
    '处理双击按钮问题的变量，这里要特殊设置成False，因为当“报到时打开报告窗体”时，实际上此变量为True
    mblnMenuDownState = False
    Call cbrMain_Execute(cbrControl)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_File_Parmeter_click()
On Error GoTo ErrHandle
    With frmTechnicSetup
        .mlngModul = mlngModule
        .mlng科室ID = mlngCur科室ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        
        If .mblnOK Then
            InitLocalPars
            
            If Not mobjWork_Report Is Nothing Then
                '重新加载和报告相关的设置参数
                Call mobjWork_Report.InitReportParameter
            End If
            
            Call RefreshList
        End If
    End With
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


'显示快捷方式配置
Private Sub Menu_File_ShortcutSet_click()
    Dim frmShortcut As New frmShortcutConfig
    
On Error GoTo ErrHandle
    Dim lngCount As Long
    
    Call frmShortcut.ShowShortcutConfig(App.ProductName, mlngModule, Me)
        
    If frmShortcut.blnIsOk Then
        '删除现在的工具栏及顶级菜单项
        Call LockWindowUpdate(Me.hWnd)
        
        For lngCount = cbrMain.ActiveMenuBar.Controls.Count To 1 Step -1
            cbrMain.ActiveMenuBar.Controls(lngCount).Delete
        Next
        
        For lngCount = cbrMain.Count To 2 Step -1
            cbrMain(lngCount).Delete
        Next
    
        Call InitCommandBars
        Call CreateWorkModuleMenu
        
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
        
        Call LockWindowUpdate(0)
    End If
    
    
    Call Unload(frmShortcut)
    Set frmShortcut = Nothing
Exit Sub
ErrHandle:
    Call Unload(frmShortcut)
    Set frmShortcut = Nothing
End Sub


Private Sub Menu_Help_About_click()
On Error GoTo ErrHandle
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Help_click()
'功能：调用帮助主题
On Error GoTo ErrHandle
    ShowHelp App.ProductName, Me.hWnd, Me.Name
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Forum_click()
On Error GoTo ErrHandle
    Call zlWebForum(Me.hWnd)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Help_Web_Mail_click()
On Error GoTo ErrHandle
    zlMailTo hWnd
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_取消关联()
'取消关联的最后结果是，每次取消关联后，图象全部按照序列被拆散成N条临时记录
On Error GoTo ErrHandle
    Dim lngResult As Long
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    lngResult = -1
    
    '如果是模块号为1298的RIS工作站，则调用新网的数据库查询已匹配的图像记录
    If mlngModule = G_LNG_PACSSTATION_MODULE And mListAdviceInf.intImageLocation = 1 Then
        lngResult = XWShowMatched(Me, mListAdviceInf.lngAdviceId)
    Else
        frmSelectMuli.ShowImageReleation mlngModule, mListAdviceInf.lngAdviceId, mstrPrivs, mblnMoved, IIf(mlngModule = G_LNG_PACSSTATION_MODULE, False, True), mlngCur科室ID, 1
        
        If frmSelectMuli.mblnOK = True Then lngResult = 0
    End If
    
    If lngResult <> 0 Then Exit Sub
    
    Call AfterReleationImage(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo, mListAdviceInf.intStep, 1, True)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_无报告完成()
'只有进行中的报告可以操作该菜单,因为此时还没有签名
On Error GoTo ErrHandle
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mListAdviceInf.strReportDoctor <> "" Or mListAdviceInf.strReportOperation <> "" Then
        If MsgBoxD(Me, "是否无报告直接完成,直接完成将删除已填写的报告!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    If mSysPar.blnFinishCommit And InStr(mstrPrivs, "检查完成") > 0 Then '无报告完成后无需再次确认完成,但需要有检查完成的权限
        '此过程,传状态=6,并且报告ID不为空将删除电子病历记录
        
        If bln费用未审核(mListAdviceInf.lngPageID, mListAdviceInf.lngPageID, mListAdviceInf.lngAdviceId, mListAdviceInf.lngPatientFrom) Then
            '执行后自动审核划价单有效，并且病人已出院，且有未审核的划价单
            MsgBoxD Me, "该病人已出院，且有未审核的划价单不能完成！", vbExclamation, gstrSysName
        Else
            If mrtReportType = 报告文档编辑器 Then
                gstrSQL = "ZL_影像检查_状态更新(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",'',6,1,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & "," & To_Date(zlDatabase.Currentdate) & ")"
            Else
                gstrSQL = "ZL_影像检查_STATE(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",6,1,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & "," & To_Date(zlDatabase.Currentdate) & ")"
            End If
        End If
    Else
        If mrtReportType = 报告文档编辑器 Then
            gstrSQL = "ZL_影像检查_状态更新(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",'',5,1,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
        Else
            gstrSQL = "ZL_影像检查_STATE(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",5,1,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
        End If
    End If
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, "改变检查过程")
        
    '取消排队信息
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCompletePacsQueue(mListAdviceInf.lngAdviceId)
    End If
    
        
    If mSysPar.blnFinishCommit Then
        Call StateCheck(6)
    Else
        Call StateCheck(5)
    End If
    
    Call NotificationAllModuleRefresh
    Call RefreshTabWindow
    
    '发送状态同步消息
    Call mobjMsgCenter.Send_Msg_StateSync(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Edit_无报告回退()
On Error GoTo ErrHandle
    Dim rsTemp As ADODB.Recordset

    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要回退该项检查吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub

    '如果有图像，则回退到“已检查”，否则回退到“已报到”
    gstrSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否有图像", mListAdviceInf.lngAdviceId)
    
    gstrSQL = "ZL_影像检查_STATE(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & "," & IIf(Nvl(rsTemp!检查uid) = "", 2, 3) & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        
    Call StateCheck(2)

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetAdviceDetailInf(Optional ByVal lngAdviceId As Long = 0) As TAdviceInf
'根据医嘱id获取详细的医嘱信息
'lngAdviceId:如果为0，则获取当前列表选中的检查医嘱信息

    Dim strSql As String
    Dim strSQLBak As String
    Dim rsTemp As ADODB.Recordset
    Dim lngIndex As Long
    Dim i As Long
    
    lngIndex = -1
    
    '设置默认的医嘱信息
    GetAdviceDetailInf = GetNullAdviceInf
    
    
    '如果列表中加载了数据，则从列表中读取医嘱信息
    If ufgStudyList.GridRows > 1 And ufgStudyList.GridCols > 1 Then
        If lngAdviceId <= 0 Then
            lngIndex = ufgStudyList.SelectionRow
        Else
            For i = 1 To ufgStudyList.GridRows - 1
                If Val(ufgStudyList.KeyValue(i)) = lngAdviceId Then
                    lngIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
    
    
    If lngIndex <= 0 And lngAdviceId > 0 Then
    
        '从数据库中查询指定医嘱id的详细信息
        If mlngModule <> G_LNG_PATHOLSYS_NUM Then
            strSql = "Select A.ID,A.姓名, A.病人科室id, A.开嘱医生,A.病人来源, A.医嘱内容, Nvl(A.婴儿, 0) 婴儿,A.病人id,e.当前床号,e.住院号,e.门诊号,decode(A.病人来源,2,F.病人类型,E.病人类型) AS 病人类型, " & vbNewLine & _
                    " A.主页id, A.挂号单, B.检查号,B.影像类别,B.检查技师, B.检查uid,B.图像位置,B.报告人,B.报告操作,B.关联ID, C.名称, D.发送号,D.执行状态,D.执行过程,D.执行间, 0 as 转出,A.执行科室ID " & vbNewLine & _
                    " From 病人医嘱记录 A, 影像检查记录 B, 部门表 C, 病人医嘱发送 D,病人信息 E,病案主页 F " & vbNewLine & _
                    " Where A.ID = B.医嘱id And A.病人科室id = C.ID And A.ID = D.医嘱id and A.病人ID=E.病人ID and A.病人ID = F.病人ID(+) And A.主页ID+0 = F.主页ID(+) and A.ID = [1]"
        Else
            strSql = "Select A.ID,A.姓名, A.病人科室id, A.开嘱医生,A.病人来源, A.医嘱内容, Nvl(A.婴儿, 0) 婴儿, A.病人id,F.当前床号,F.住院号,F.门诊号,decode(A.病人来源,2,G.病人类型,F.病人类型) AS 病人类型, " & vbNewLine & _
                    " A.主页id, A.挂号单, E.病理号,B.影像类别,B.检查技师, B.检查uid,B.图像位置,B.报告人,B.报告操作,B.关联ID, C.名称, D.发送号,D.执行状态,D.执行过程,D.执行间,0 as 转出,A.执行科室ID " & vbNewLine & _
                    " From 病人医嘱记录 A, 影像检查记录 B, 部门表 C, 病人医嘱发送 D, 病理检查信息 E, 病人信息 F,病案主页 G " & vbNewLine & _
                    " Where A.ID = B.医嘱id And A.病人科室id = C.ID And A.ID = D.医嘱id and A.ID=E.医嘱ID and A.病人ID=F.病人ID and A.病人ID = G.病人ID(+) And A.主页ID+0 = G.主页ID(+) and A.ID = [1]"
        End If
                    
        strSQLBak = strSql
        strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
        strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
        strSQLBak = Replace(strSQLBak, "影像检查记录", "H影像检查记录")
'        strSQLBak = Replace(strSQLBak, "病理检查信息", "H病理检查信息")    '病理检查信息在10.32.0之后不参与转储
'        strSQLBak = Replace(strSQLBak, "病人信息", "H病人信息")            '病人信息表并未参与转存
        
        strSQLBak = Replace(strSQLBak, "0 as 转出", "1 as 转出")
        
        strSql = strSql & vbNewLine & " Union ALL " & strSQLBak
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查历次记录信息", lngAdviceId)
        
        If Not rsTemp.EOF Then
            With GetAdviceDetailInf
                .lngPatID = Val(Nvl(rsTemp!病人ID))
                .lngAdviceId = lngAdviceId
                .lngSendNo = Val(Nvl(rsTemp!发送号))
                .lngPageID = Val(Nvl(rsTemp!主页ID))
                .lngPatDept = Val(Nvl(rsTemp!病人科室ID))
                .strPatientName = Nvl(rsTemp!姓名)
                .lngUnit = .lngPatDept
                .blnCanPrint = True
                
                .lngPatientFrom = Val(Nvl(rsTemp!病人来源, 3))
                .strPatientType = Nvl(rsTemp!病人类型)
                
                .blnIsInsidePatient = (.lngPatientFrom = 1) Or (.lngPatientFrom = 2)
                .intMoved = Val(Nvl(rsTemp!转出))
                .intState = Val(rsTemp!执行状态)
                .intStep = Val(Nvl(rsTemp!执行过程))
                .strRegNo = Val(Nvl(rsTemp!挂号单))
                .lngRegId = getRegID(.strRegNo)
                .strStudyUID = Val(Nvl(rsTemp!检查uid))
                .lngExeDepartmentId = Val(Nvl(rsTemp!执行科室ID))
                .strDoDoctor = Nvl(rsTemp!检查技师)
                .strExeRoom = Nvl(rsTemp!执行间)
                .strStudyNum = Nvl(rsTemp(GetStudyNumberDisplayName))
                .strBedNum = Nvl(rsTemp!当前床号)
                .lngBaby = Val(Nvl(rsTemp!婴儿))
                .strPatientDepartment = Nvl(rsTemp!名称)
                .lngMarkNum = IIf(.lngPatientFrom = 1, Val(Nvl(rsTemp!门诊号)), IIf(.lngPatientFrom = 2, Val(Nvl(rsTemp!住院号)), 0))
                
                .strReportDoctor = Nvl(rsTemp!报告人)
                .strReportOperation = Nvl(rsTemp!报告操作)
                
                .lngLinkId = Val(Nvl(rsTemp!关联ID))
                
                .strImgType = Nvl(rsTemp!影像类别)
                .intImageLocation = Nvl(rsTemp!图像位置)
            End With
        End If
        
        Exit Function
    End If
    
    '如果当前列表中没有检查，且医嘱id为0，则退出该函数
    If lngIndex <= 0 And lngAdviceId <= 0 Then Exit Function
    
    
    '从界面中读取医嘱id相关的详细信息
    With GetAdviceDetailInf
        .lngPatID = Val(ufgStudyList.Text(lngIndex, "病人ID"))
        .lngPageID = Val(ufgStudyList.Text(lngIndex, "主页ID"))
        .lngAdviceId = Val(ufgStudyList.KeyValue(lngIndex))
        .lngSendNo = Val(ufgStudyList.Text(lngIndex, "发送号"))
        .lngPatDept = Val(ufgStudyList.Text(lngIndex, "病人科室ID"))
        .strPatientName = ufgStudyList.Text(lngIndex, "姓名")
        .strRegNo = ufgStudyList.Text(lngIndex, "挂号单")
        .lngRegId = getRegID(.strRegNo)
        .intMoved = Val(ufgStudyList.Text(lngIndex, "转出"))
        .intState = IIf(ufgStudyList.Text(lngIndex, "检查过程") = "已拒绝", 2, IIf(ufgStudyList.Text(lngIndex, "检查过程") = "已完成", 1, 3))
        .intStep = Val(ufgStudyList.Text(lngIndex, "检查状态")) '读取执行过程
        .lngUnit = Val(ufgStudyList.Text(lngIndex, "当前病区ID"))
        
        If mrtReportType = 报告文档编辑器 Then
            If ufgStudyList.Text(lngIndex, "紧急") = " " Or Val(ufgStudyList.Text(lngIndex, "绿色通道")) = 1 Then
                .blnCanPrint = True
            Else
                .blnCanPrint = False
            End If
        Else
            .blnCanPrint = IIf(mSysPar.blnCanPrint, IIf(Val(ufgStudyList.Text(lngIndex, "紧急")) = 1, ufgStudyList.Text(lngIndex, "报告人") <> "", ufgStudyList.Text(lngIndex, "复核人") <> ""), True)
        End If
        
        .strStudyUID = ufgStudyList.Text(lngIndex, "检查UID")
        .lngExeDepartmentId = Val(ufgStudyList.Text(lngIndex, "执行科室ID"))
        .strDoDoctor = ufgStudyList.Text(lngIndex, "检查技师")
        
        '当执行刷新操作后，单元格的flexcpdata数据不会立即就被刷新，只能通过对应的显示文本对值进行转换，flexcpdata值的更新由异步事件触发
        .lngPatientFrom = Decode(ufgStudyList.Text(lngIndex, "来源"), "门", 1, "住", 2, "外", 3, 4)
        
        .strPatientType = ufgStudyList.Text(lngIndex, "病人类型")
        
        .blnIsInsidePatient = (.lngPatientFrom = 1) Or (.lngPatientFrom = 2)
        .strExeRoom = ufgStudyList.Text(lngIndex, "执行间")
        .strStudyNum = ufgStudyList.Text(lngIndex, GetStudyNumberDisplayName)
        .strBedNum = ufgStudyList.Text(lngIndex, "床号")
        .lngMarkNum = Val(ufgStudyList.Text(lngIndex, "标识号"))
        .lngBaby = 0
        
        .strReportDoctor = ufgStudyList.Text(lngIndex, "报告人")
        .strReportOperation = ufgStudyList.Text(lngIndex, "报告操作")
        
        .lngLinkId = Val(ufgStudyList.Text(lngIndex, "关联ID"))
        .strImgType = ufgStudyList.Text(lngIndex, "影像类别")
        .intImageLocation = Val(ufgStudyList.Text(lngIndex, "图像位置"))
        
        strSql = "Select 名称 From 部门表 Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取病人科室", .lngPatDept)
        
        .strPatientDepartment = ""
        If rsTemp.RecordCount > 0 Then .strPatientDepartment = Nvl(rsTemp!名称)

    End With
        
End Function

Private Function getRegID(ByVal strRegNo As String) As Long
'功能:获取挂号id
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    getRegID = 0
    
    strSql = "select id from 病人挂号记录 where no=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, strRegNo)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    getRegID = Nvl(rsTemp!ID, 0)
    
    Exit Function

ErrHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Function IsAlreadyInputQuality(ByVal lngAdviceId As Long) As Boolean
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    IsAlreadyInputQuality = False
    
    strSql = "select 综合质量 from 病理检查信息 where 医嘱ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If Nvl(rsData!综合质量) <> "" Then IsAlreadyInputQuality = True
    
End Function

Private Sub Menu_Manage_检查最终完成(Optional lngAdviceId As Long = 0, Optional blnRefresh As Boolean = True, Optional strReportId As String = "")
'可能由其它过程调用，此时传入有医嘱ID，但需要权限判断
On Error GoTo ErrHandle
    Dim strSql As String
    Dim curAdviceInf As TAdviceInf
    Dim intState As Integer
    
    If InStr(mstrPrivs, "检查完成") <= 0 Then Exit Sub
    
    curAdviceInf = GetAdviceDetailInf(lngAdviceId)
    
    If curAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If lngAdviceId = 0 Then
    '如果是还有报告未完成时完成检查
        If mrtReportType = 报告文档编辑器 Then
            intState = getStudyStateRich(curAdviceInf.lngAdviceId, strReportId, False)
        Else
            intState = getStudyState(curAdviceInf.lngAdviceId, False)
        End If
        
        If intState = 4 Then
            If MsgBoxD(Me, "还有报告没有写完，如果此时完成检查，需要有“补录报告”权限的人才能继续书写报告!" & vbCrLf & "确定要继续完成吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    '如果是病理系统，检查完成时，则需要弹出质量控制窗口
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If (mblnPopChangGuiWindow And ufgStudyList.CurText("检查类别") = "常规") _
            Or (mblnPopKuaiShuWindow And ufgStudyList.CurText("检查类别") = "快速石蜡") _
            Or (mblnPopBingDongWindow And ufgStudyList.CurText("检查类别") = "冰冻") _
            Or (mblnPopXiBaoWindow And ufgStudyList.CurText("检查类别") = "细胞") _
            Or (mblnPopHuiZhenWindow And ufgStudyList.CurText("检查类别") = "会诊") _
            Or (mblnPopShiJianWindow And ufgStudyList.CurText("检查类别") = "尸检") Then
            
            If Not IsAlreadyInputQuality(curAdviceInf.lngAdviceId) Then
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.zlMenu.zlExecuteMenu(conMenu_Pathol_Quality_Manage)
            End If
            
            If Not IsAlreadyInputQuality(curAdviceInf.lngAdviceId) Then
                Call MsgBoxD(Me, "未录入检查质量，不能执行完成操作。", vbInformation, GetWindowCaption)
                Exit Sub
            End If
            
        End If
    End If
    
    If bln费用未审核(curAdviceInf.lngPatID, curAdviceInf.lngPageID, Nvl(curAdviceInf.lngAdviceId), curAdviceInf.lngPatientFrom) Then
        '执行后自动审核划价单有效，并且病人已出院，且有未审核的划价单
        MsgBoxD Me, "该病人已出院，且有未审核的划价单，不能完成！", vbExclamation, gstrSysName
        Exit Sub
    Else
        If mrtReportType = 报告文档编辑器 Then
            strSql = "Zl_影像检查_状态更新(" & curAdviceInf.lngAdviceId & "," & curAdviceInf.lngSendNo & ",'',6,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & "," & To_Date(zlDatabase.Currentdate) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "改变检查过程")
        Else
            strSql = "ZL_影像检查_STATE(" & curAdviceInf.lngAdviceId & "," & curAdviceInf.lngSendNo & ",6,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & "," & To_Date(zlDatabase.Currentdate) & ")"
            Call zlDatabase.ExecuteProcedure(strSql, "改变检查过程")
        End If
        
        If mlngModule = G_LNG_PATHOLSYS_NUM Then
            gstrSQL = "Zl_病理检查_完成(" & curAdviceInf.lngAdviceId & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "病理检查完成")
        End If
    End If

    
    '取消排队信息
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCompletePacsQueue(mListAdviceInf.lngAdviceId)
    End If

    If blnRefresh Then Call StateCheck(6)
        
    Call NotificationAllModuleRefresh
    Call RefreshTabWindow
    
    '发送检查完成消息
    Call mobjMsgCenter.Send_Msg_StudyComplete(mListAdviceInf.lngAdviceId, strReportId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_取消检查完成()
On Error GoTo ErrHandle
    Dim strSql As String
    Dim intState As Integer

    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    If mListAdviceInf.intMoved = 1 Then
        MsgBoxD Me, "该病人的本次住院数据已经转出到后备数据库，不允许操作。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If CheckIsArchived(mListAdviceInf.lngAdviceId) Then
            MsgBoxD Me, "该病人的档案已经归档，不允许操作。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If mrtReportType = 报告文档编辑器 Then
        intState = getStudyStateRich(mListAdviceInf.lngAdviceId, "", True)
        strSql = "ZL_影像检查_状态更新(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",''," & intState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & "," & To_Date(zlDatabase.Currentdate) & ")"
    Else
        intState = getStudyState(mListAdviceInf.lngAdviceId, True)
        strSql = "ZL_影像检查_STATE(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & "," & intState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & "," & To_Date(zlDatabase.Currentdate) & ")"
    End If
    
    zlDatabase.ExecuteProcedure strSql, "取消检查完成"
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSql = "Zl_病理检查_取消完成(" & mListAdviceInf.lngAdviceId & ")"
        Call zlDatabase.ExecuteProcedure(strSql, "病理检查取消完成")
    End If
    
    Call StateCheck(intState)
    
    Call NotificationAllModuleRefresh
    Call RefreshTabWindow
    
    '发送检查撤销完成消息
    Call mobjMsgCenter.Send_Msg_CancelComplete(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function CheckIsArchived(lngAdviceId As Long) As Boolean
 '检查该病人档案是否已经归档，已归档的检查，需要撤档才能取消完成  0--未归档  1--已归档
 On Error GoTo ErrHandle
 
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    strSql = "select distinct c.档案状态 as 状态 from 病理检查信息 a,病理归档信息 b,病理档案信息 c where a.病理医嘱ID = b.病理医嘱ID and b.档案id = c.id and a.医嘱ID =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查是否已归档", lngAdviceId)
    
    If rsTemp.RecordCount < 1 Then
        CheckIsArchived = False
        Exit Function
    End If
    
    CheckIsArchived = IIf(Nvl(rsTemp!状态, 0) = 1, True, False)
Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Menu_Manage_CriticalMark(ByVal lngID As Long)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim intCritical As Integer
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_CriticalValues, conMenu_Manage_Critical
            intCritical = 1
        Case conMenu_Manage_Normal
            intCritical = 0
    End Select
    
    With ufgStudyList
        If intCritical = 1 Then
            If lngID = conMenu_Manage_CriticalValues Then
                Call frmCriticalValues.ShowMe(mListAdviceInf.lngAdviceId, Me)
                If Not frmCriticalValues.mblnSave Then Exit Sub
            End If
            
            strSql = "zl_影像检查_危急更新(" & mListAdviceInf.lngAdviceId & ",1)"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("危急")) = imgList.ListImages("危急").Picture
                .CurText("危急") = " "
                
            Menu_Manage_标记阴阳 conMenu_Manage_Negative
            
            '发送危急值消息
            Call mobjMsgCenter.Send_Msg_Critical(mListAdviceInf.lngAdviceId)
        ElseIf intCritical = 0 Then
            If .CurText("危急") = "" Then Exit Sub
            If MsgBoxD(Me, "确定要取消病人【" & .CurText("姓名") & "】的危急状态吗？", vbOKCancel, "危急处理描述") = 2 Then Exit Sub
            
            strSql = "Zl_影像危急值记录_取消(" & mListAdviceInf.lngAdviceId & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)

            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("危急")) = Nothing
                .CurText("危急") = ""
        End If
        
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "危急", intCritical)
    End With

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_标记阴阳(ByVal lngID As Long)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim iResult As Integer
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_Negative
            iResult = 1
        Case conMenu_Manage_Positive
            iResult = 0
    End Select
    
    If mrtReportType = 报告文档编辑器 Then
        Call mobjWork_Report.Menu_Manage_标记阴阳(mListAdviceInf.lngAdviceId, iResult)
        Exit Sub
    End If
    
    strSql = "ZL_影像检查_结果(" & mListAdviceInf.lngAdviceId & "," & iResult & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "结果阴阳性")
    
    With ufgStudyList
        If iResult = 1 Then
            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("阳性")) = imgList.ListImages("阳性").Picture
            .CurText("阳性") = " "
        Else
            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("阳性")) = Nothing
            .CurText("阳性") = ""
        End If
        
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "阳性", iResult)
    End With

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_绿色通道(ByVal lngID As Long)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim intResult As Integer
    Dim blnCanPrint As Boolean
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    
    strSql = "Zl_绿色通道_Update(" & mListAdviceInf.lngAdviceId & ",'" & intResult & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "绿色通道")
    
    With ufgStudyList
        .CurText("绿色通道") = intResult
        
        If intResult = 1 Then
            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("姓名")) = imgList.ListImages("绿色通道").Picture
        Else
            Set .DataGrid.Cell(flexcpPicture, .DataGrid.Row, .GetColIndex("姓名")) = Nothing
        End If
        
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "绿色通道", intResult)
    End With
    
    If mrtReportType = 报告文档编辑器 Then
        blnCanPrint = ufgStudyList.CurText("紧急") = " " Or Val(ufgStudyList.CurText("绿色通道")) = 1
        
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, ufgStudyList, mblnIsHistory, blnCanPrint, mcurAdviceInf.strDoDoctor, mcurAdviceInf.strStudyUID)
            Call mobjWork_Report.zlRefreshFace(True, False, False)
        End If
    End If

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_符合情况(ByVal lngID As Long)
On Error GoTo ErrHandle
    Dim strResult As String
    Dim strSql As String
    Dim lngColIndex As Long

    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    Select Case lngID
        Case conMenu_Manage_FuHe
            strResult = "符合"
        Case conMenu_Manage_JiBenFuHe
            strResult = "基本符合"
        Case conMenu_Manage_BuFuHe
            strResult = "不符合"
    End Select

    strSql = "Zl_符合情况_Update(" & mListAdviceInf.lngAdviceId & ",'" & strResult & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "符合情况")
        
    With ufgStudyList
        .CurText("符合情况") = strResult
        
        lngColIndex = ufgStudyList.GetColIndex("符合情况")
         
        If strResult = "符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, .DataGrid.Row, lngColIndex) = vbGreen
        If strResult = "基本符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, .DataGrid.Row, lngColIndex) = vbYellow
        If strResult = "不符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, .DataGrid.Row, lngColIndex) = vbRed
        
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "符合情况", strResult)
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_修改()
On Error GoTo ErrHandle
    Dim strOldName As String
    Dim strOldRoom As String
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mListAdviceInf.lngSendNo
            .mlngAdviceId = mListAdviceInf.lngAdviceId
            .mstrPatientName = mListAdviceInf.strPatientName
            .mintEditMode = IIf(mListAdviceInf.intStep > 1, 3, 1)  '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = mlngCur科室ID
            .mstrCur科室 = NeedName(mstrCur科室)
'            .mlngQueueWay = mSysPar.lngQueueWay
            
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            
            If .mlngResultState <> 0 Then
                strOldName = mListAdviceInf.strPatientName
                strOldRoom = mListAdviceInf.strExeRoom
                
                Call RefreshList(.mlngAdviceId, True) '成功返回
                
                If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                    '如果是报到后修改，且改变了执行间，则需要重新进行排队
                    If .mintEditMode = 3 And .mlngResultState = 3 Then
                        If .mstrTechnicRoom <> strOldRoom Then
                            If .mstrTechnicRoom = "" Then
                                '如果为空，则需要插入该检查项目对应的项目分组或者科室的队列中
                                Call mobjQueue.zlGetInQueueInf(mListAdviceInf.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                            Else
                                '如果不为空，则写入对应的执行间名称
                                strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                                strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                            End If
                            
                            Call mobjQueue.zlUpdatePacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                        Else
                            '其他方式的修改，则只对排队叫号中的相关信息进行更新
                            If .mstrPatientName <> strOldName Then
                                Call mobjQueue.zlUpdatePacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = Val(ufgStudyList.CurText("发送号"))
            .mlngAdviceId = mListAdviceInf.lngAdviceId
            .mintEditMode = IIf(Val(ufgStudyList.CurText("检查状态")) > 1, 3, 1)  '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = mlngCur科室ID
            .mintImgCount = mintImgCount
            .mblnHasSpecimenAccept = IIf(InStr(mstrWorkModule, ";标本核收模块;") > 0, True, False)
            .InitMvar
            
            If .RefreshPatiInfor(False) = True Then  '刷新病人
                .mblnOK = False
                .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOK Then RefreshList '成功返回
        End With
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'Private Sub Menu_Manage_ModifBaseInfo()
''基本信息调整
'On Error GoTo errHandle
'    Dim int场合 As Integer
'    Dim str就诊ID As String
'
'    With mcurAdviceInf
'        int场合 = Decode(.lngPatientFrom, 1, 1, 2, 2, 3, 3, 4, 4)
'
'        str就诊ID = Decode(.lngPatientFrom, 1, .lngRegId, 2, .lngPageID, 3, .lngAdviceID, 4, .strRegNo)
'
'        If zlDatabase.zlModiPatiBaseInfo(.lngPatID, str就诊ID, mlngModule, int场合) Then RefreshList    '成功返回
'    End With
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub

Private Sub Menu_Manage_复制登记()
On Error GoTo ErrHandle
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = mlngCur科室ID
            .mstrCur科室 = NeedName(mstrCur科室)
            .mlngResultState = 0
            
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1), mblnAllDepts, mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo
            
            If .mlngResultState <> 0 Then '成功返回
                Call StateCheck(2, .mlngAdviceId)
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = mlngCur科室ID
            .mblnOK = False
            .mblnHasSpecimenAccept = IIf(InStr(mstrWorkModule, ";标本核收模块;") > 0, True, False)
            .InitMvar
            
            If .CopyCheck(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo) = True Then  '刷新病人
                .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            End If
            
            If .mblnOK Then '成功返回
                Call StateCheck(2, .mlngAdviceId)
            End If
        End With
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub AutoPrintApplication(lngAdviceId As Long, lngSendNo As Long, lngClinicId As Long, intSourceType As Long)
'功能:根据能数据自动打印申请单
Dim rsTemp As ADODB.Recordset, strBillNo As String, strExseNo As String, intExseKind As Integer

On Error GoTo errHand
    Dim strSql As String
    
    If Not mSysPar.blnAutoPrint Then Exit Sub
    
    strSql = "select NO,记录性质 from 病人医嘱发送 where 医嘱ID=[1] and 发送号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取NO", lngAdviceId, lngSendNo)
    If rsTemp.EOF Then Exit Sub
    
    strExseNo = rsTemp!NO: intExseKind = rsTemp!记录性质
    
    strSql = "Select B.ID, B.编号" & vbNewLine & _
                "From 病历单据应用 A, 病历文件列表 B" & vbNewLine & _
                "Where A.诊疗项目id =[1] And A.应用场合 =[2] And A.病历文件id = B.ID And B.种类 = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取单据编号", lngClinicId, CLng(Decode(intSourceType, 1, 1, 2, 2, 1)))
    
    If rsTemp.EOF Then Exit Sub
    strBillNo = "ZLCISBILL" & Format(rsTemp!编号, "00000") & "-1"
    ReportOpen gcnOracle, glngSys, strBillNo, Me, "NO=" & strExseNo, "性质=" & intExseKind, "医嘱ID=" & lngAdviceId, 2
    Exit Sub

errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub Menu_Manage_登记()
On Error GoTo ErrHandle
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Set mfrmRISRequest = New frmRISRequest
        With mfrmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mstrPatientName = ""
            .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = mlngCur科室ID
            .mstrCur科室 = NeedName(mstrCur科室)
'            .mlngQueueWay = mSysPar.lngQueueWay
            .mlngResultState = 0
            
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1), mblnAllDepts
            
            If .mlngResultState <> 0 Then '成功返回
                Call StateCheck(2, .mlngAdviceId)
                
                If ufgStudyList.DataGrid.Rows = 2 Then
                    Call ufgStudyList.LocateRow(1)
                End If
                
                If .mlngResultState = 2 Then
                
                    '如果启用排队叫号，则报到后需要插入排队叫号队列......
                    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing Then
                        '设置需要插入的队列名称
                        If .mstrTechnicRoom = "" Then
                            '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
                            Call mobjQueue.zlGetInQueueInf(.mlngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                        Else
                            '如果不为空，则写入对应的执行间名称
                            strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                            strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                        End If
                        
                        Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                    End If
                
                    '打印申请单
                    Call AutoPrintApplication(.mlngAdviceId, .mlngSendNo, .mlngClinicID, .mintSourceType)
                End If
                
                '发送新申请消息
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = 0
            .mlngAdviceId = 0
            .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = mlngCur科室ID
            .mintImgCount = 0
            .mblnOK = False
            .mblnHasSpecimenAccept = IIf(InStr(mstrWorkModule, ";标本核收模块;") > 0, True, False)
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            
            If .mblnOK Then '成功返回
    
                Call StateCheck(2, .mlngAdviceId)
    
                
                If ufgStudyList.DataGrid.Rows = 2 Then
                    Call ufgStudyList.LocateRow(1)
                End If
                
                '如果同时勾选“开始检查自动打开报告”和“登记后自动报到”参数那么会自动打开报告界面
                If mSysPar.blnAutoOpenReport And mSysPar.bln直接检查 Then Call Menu_RichEPR(conMenu_Edit_Modify)
                
                '发送新申请消息
                Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
            End If
        End With
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_取消登记()
On Error GoTo ErrHandle
    Dim strSql As String
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确认要取消当前申请吗？" & Chr(10) & Chr(13) & "申请取消后，其对应的医嘱将拒绝执行！", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSql = "ZL_病人医嘱执行_拒绝执行(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, "撤消登记")
    
    Call RefreshList
    
    '发送医嘱撤销消息
    Call mobjMsgCenter.Send_Msg_CancelAdvice(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_召回取消()
'功能：召回被取消的登记
On Error GoTo ErrHandle
    Dim strSql As String
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If MsgBoxD(Me, "确实要召回被取消登记的项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSql = "ZL_病人医嘱执行_取消拒绝(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ")"
    
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    Call RefreshList
    
    '发送状态同步消息
    Call mobjMsgCenter.Send_Msg_StateSync(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Manage_报到()
On Error GoTo ErrHandle
    Dim blnFocusFind As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If mcurAdviceInf.lngPatientFrom = 4 Then '只有体检病人才执行以下过程
        Call zlDatabase.ExecuteProcedure("zl_PeisLockAdviceState(" & mListAdviceInf.lngAdviceId & ")", Me.Caption)
    End If
    
    If Me.ActiveControl Is Nothing Then
        blnFocusFind = False
    Else
        blnFocusFind = (Me.ActiveControl.Name = "txtFilter")
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        With frmRISRequest
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mListAdviceInf.lngSendNo
            .mlngAdviceId = mListAdviceInf.lngAdviceId
            .mstrPatientName = mListAdviceInf.strPatientName
            .mintEditMode = 2 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = mlngCur科室ID
            .mstrCur科室 = NeedName(mstrCur科室)
'            .mlngQueueWay = mSysPar.lngQueueWay
            .mlngResultState = 0
            
            .InitMvar
            .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            
            If .mlngResultState <> 0 Then  '成功返回
                Call StateCheck(2)
                
                If .mblnIsRelationImage = True Then
                    '如果对提前检查的图像进行了自动关联处理，则这里将对影像图像模块进行刷新
                    If Not mfrmWork_PacsImg Is Nothing Then
                        Call mfrmWork_PacsImg.zlUpdateAdviceInf(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo, mListAdviceInf.intStep, mListAdviceInf.intMoved = 1)
                        Call mfrmWork_PacsImg.zlRefreshFace(True)
                    End If
                End If
                
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '开始检查自动打开报告
                
                If .mlngResultState = 2 Then
                    '如果启用排队叫号，并且报到后自动排队，则报到后需要插入排队叫号队列......
                    If mSysPar.blnUseQueue And mSysPar.blnAutoInQueue And Not mobjQueue Is Nothing Then
                        '设置需要插入的队列名称
                        If .mstrTechnicRoom = "" Then
                            '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
                            Call mobjQueue.zlGetInQueueInf(mListAdviceInf.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                        Else
                            '如果不为空，则写入对应的执行间名称
                            strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                            strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                        End If
                        
                        Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
                    End If
                    
                                    '打印申请单
                    Call AutoPrintApplication(.mlngAdviceId, .mlngSendNo, .mlngClinicID, .mintSourceType)
                End If
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(.mlngAdviceId)
                If mcurAdviceInf.lngPatientFrom <> 3 Then
                    Call mobjMsgCenter.Send_Msg_Arrange(.mlngAdviceId)
                End If
            End If

        End With
    Else
        With frmPatholRIS
            .mstrPrivs = mstrPrivs
            .mlngModul = mlngModule
            .mlngSendNo = mListAdviceInf.lngSendNo
            .mlngAdviceId = mListAdviceInf.lngAdviceId
            .mintEditMode = 2 '0－登记、1－登记后修改、2－报到、3－报到后修改
            .mlngCurDeptId = mlngCur科室ID
            .mintImgCount = mintImgCount
            .mblnHasSpecimenAccept = IIf(InStr(mstrWorkModule, ";标本核收模块;") > 0, True, False)
            .InitMvar
            If .RefreshPatiInfor(True) = True Then  '刷新病人
                .mblnOK = False
                .zlShowMe Me, mstrDefaultPatientType, IIf(mbytFontSize <= 9, 0, 1)
            End If
            If .mblnOK Then  '成功返回
                Call StateCheck(2)
                If mSysPar.blnAutoOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '开始检查自动打开报告
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(.mlngAdviceId)
            End If
            
        End With
    End If
    
    If blnFocusFind Then PatiIdentify.SetFocus '自动定位到定位栏
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

'排队叫号入队
Private Sub zlInPacsQueue()
    Dim strQueueName As String
    Dim strCodeNo As String
    
    If mobjQueue Is Nothing Then Exit Sub
    
    '设置需要插入的队列名称
    If Trim(ufgStudyList.CurText("执行间")) = "" Then
        '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
        Call mobjQueue.zlGetInQueueInf(mListAdviceInf.lngAdviceId, mlngCur科室ID, strQueueName, strCodeNo)
    Else
        '如果不为空，则写入对应的执行间名称
        strQueueName = NeedName(mstrCur科室) & "-" & ufgStudyList.CurText("执行间")
        strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(ufgStudyList.CurText("执行间"), mlngCur科室ID)
    End If
    
    Call mobjQueue.zlInQueue(mListAdviceInf.lngAdviceId, ufgStudyList.CurText("姓名"), mlngCur科室ID, strQueueName, ufgStudyList.CurText("执行间"), strCodeNo)
End Sub




Private Sub Menu_Manage_取消报到()
On Error GoTo ErrHandle
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim lngResult As Long

    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
  
    If mListAdviceInf.intStep <= 1 Then Call Menu_Manage_取消登记: Exit Sub  '工具栏调用
    '------------------------------------有签名的需要先回退签名后再撤消
    strSql = "Select Distinct B.完成时间 From 病人医嘱报告 A, 电子病历记录 B Where A.病历ID=B.Id And A.医嘱ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取是否签名", mListAdviceInf.lngAdviceId)
    
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!完成时间, "") <> "" Then '签名保存
            MsgBoxD Me, "当前病人的检查报告已经签名,若需取消检查,请先回退签名!", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    '如果检查已取材或者制片，则不能进行取消
    strSql = "select count(1) as 数量 from 病理检查信息 a, 病理取材信息 b where a.病理医嘱ID=b.病理医嘱ID and a.医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, mListAdviceInf.lngAdviceId)
    If rsTemp.RecordCount > 0 Then
        If Val(Nvl(rsTemp!数量)) > 0 Then
            Call MsgBoxD(Me, "该检查已执行取材操作，不能进行取消。", vbInformation, GetWindowCaption)
            Exit Sub
        End If
    End If

    If MsgBoxD(Me, "取消本次检查将删除相应的检查图像和检查报告，是否继续？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    If mListAdviceInf.strStudyUID <> "" And InStr(mstrPrivs, "清除图像") <= 0 Then
        MsgBoxD Me, "您没有清除检查图像权限,不能请除图像,所以不能取消此项检查!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '取消排队信息
    If mSysPar.blnUseQueue = True And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlCancelPacsQueue(mListAdviceInf.lngAdviceId)
    End If
    
    '如果是RIS工作站，而且图像在新网PACS中，则需要先取消关联，然后再调用ZL_影像检查_CANCEL过程取消报到
    If mlngModule = G_LNG_PACSSTATION_MODULE And mListAdviceInf.intImageLocation = 1 Then
        '取消图像关联
        Call XWUnmatchImage(mListAdviceInf.lngAdviceId, 0)
    End If
    
    '取消报告，修改数据库状态，删除“影像检查记录”
    strSql = "ZL_影像检查_CANCEL(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",0," & mlngCur科室ID & ")"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strSql = "ZL_病理检查_撤销(" & mListAdviceInf.lngAdviceId & ")"
        zlDatabase.ExecuteProcedure strSql, GetWindowCaption
    End If
    
    '如果图像在中联PACS，则删除影像文件和目录
    If mListAdviceInf.intImageLocation = 0 Then
        RemoveCheckImages mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo
    End If
    
    Call StateCheck(1)
    
    '发送状态回退消息
    Call mobjMsgCenter.Send_Msg_StateCancel(mListAdviceInf.lngAdviceId)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_关联影像()
On Error GoTo ErrHandle
    Dim lngResult As Long
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    lngResult = -1
    '如果是模块号为RIS工作站，则调用新网的数据库查询未匹配的图像记录
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        lngResult = XWShowUnMatched(Me, mListAdviceInf.lngAdviceId, mListAdviceInf.strImgType)
        
        If lngResult = 0 Then
            '图像关联成功后,使其值为1
            mListAdviceInf.intImageLocation = 1
            
            ufgStudyList.CurText("图像位置") = "1"
            
            Call ufgStudyList.UpdateSourceData(mListAdviceInf.lngAdviceId, "图像位置", 1)
        End If
    Else
        frmSelectMuli.ShowImageReleation mlngModule, mListAdviceInf.lngAdviceId, mstrPrivs, mListAdviceInf.intMoved = 1, IIf(mlngModule = G_LNG_PACSSTATION_MODULE, False, True), mlngCur科室ID, 2, mListAdviceInf.strImgType
        
        If Not frmSelectMuli.mblnOK Then Exit Sub
        lngResult = 0
    End If
    
    If lngResult <> 0 Then Exit Sub
    
    Call AfterReleationImage(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo, mListAdviceInf.intStep, 2, True)
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer
    Dim objDepartmentMenu As CommandBarControl
    Dim objControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    
    If Not mblnInitOk Then Exit Sub
    
    If mlngCur科室ID <> control.DescriptionText Or (control.DescriptionText <> 0 And mblnAllDepts = True) Then
        mstrRPTExecutor = UserInfo.姓名
        
        If Not mobjWork_Report Is Nothing And mrtReportType = 报告文档编辑器 Then
            Call mobjWork_Report.SetDocCreator(mstrRPTExecutor)
        End If
        
        stbThis.Panels(4).Text = "报告医生：" & mstrRPTExecutor & "   检查医生：" & Split(stbThis.Panels(4).Text, "检查医生：")(1)
        
        mcurAdviceInf = GetNullAdviceInf
        mListAdviceInf = mcurAdviceInf
        
        '科室切换后，由于没有重新创建菜单和工作模块，也没有调用cbrMain.RecalcLayout，因此需要使用该对象设置科室切换后的科室信息
        Set objDepartmentMenu = cbrMain.FindControl(, conMenu_View_Filter * 10#)
        
        If control.DescriptionText = 0 Then
            '选择所有科室
            mblnAllDepts = True
        
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = "当前科室:全部科室"
            
            If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
                Set objControl = cbrdock.FindControl(, ID_影像执行间)
                For i = 1 To objControl.CommandBar.Controls.Count
                    objControl.CommandBar.Controls(1).Delete
                Next
                
                Call InitExamineRoom(objControl, cbrPopControl, 0)
            End If
        Else
            '选择单个科室
            mblnAllDepts = False
            
            mlngCur科室ID = control.DescriptionText
            mstrCur科室 = Split(control.Caption, "(")(0)
            mrtReportType = GetDeptPara(mlngCur科室ID, "报告编辑器", 0)                 '报告编辑器
            
            Call InitStudyList
            
            If Not objDepartmentMenu Is Nothing Then objDepartmentMenu.Caption = "当前科室:" & mstrCur科室

            If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
                Set objControl = cbrdock.FindControl(, ID_影像执行间)
                For i = 1 To objControl.CommandBar.Controls.Count
                    objControl.CommandBar.Controls(1).Delete
                Next
                
                Call InitExamineRoom(objControl, cbrPopControl, mlngCur科室ID)
            End If
            
            Call InitModuleParameter(False)
            
            Call ReadStudyListColor(mlngCur科室ID)
            
            Call RefreshCustomQueryMenu(cbrMain.FindControl(, conMenu_Manage_Query), mlngCur科室ID)

        
            If mblnUseActivexCapture Then
                '使用ActivexExe方式的视频采集方式处理
                If Not mobjWork_ActiveVideo Is Nothing Then
                    Call mobjWork_ActiveVideo.zlInitModule(gcnOracle, glngSys, mlngModule, mstrPrivs, mlngCur科室ID, Me.hWnd, Me, True)
                End If
            End If
            
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.zlInitModule(mlngModule, mstrPrivs, mlngCur科室ID, Me)
            If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.zlInitModule(mlngModule, mstrPrivs, mlngCur科室ID, Me)
            If Not mobjWork_His Is Nothing Then
                If mblnAllDepts Then
                    Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, UserInfo.部门ID, Me)
                Else
                    Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, mlngCur科室ID, Me)
                End If
            End If
            
            '科室切换后，如果启用了排队叫号，则添加排队叫号页面
            If mSysPar.blnUseQueue = True Then
                If mobjQueue Is Nothing Then
                    mstrWorkModule = mstrWorkModule & ";排队叫号模块;"
                    
                    Set mobjQueue = New frmWork_Queue
                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur科室ID, NeedName(mstrCur科室), mstrPrivs)
                    
                    TabWindow.InsertItem 13, "排队叫号", mobjQueue.hWnd, 10011
                    TabWindow.Item(TabWindow.ItemCount - 1).tag = "排队叫号"
                    
                    Call picWindow_Resize
                Else
                    Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur科室ID, NeedName(mstrCur科室), mstrPrivs)
                End If
            Else
                If mSysPar.blnUseQueue = False And Not mobjQueue Is Nothing Then
                    mstrWorkModule = Replace(mstrWorkModule, ";排队叫号模块;", "")
                    
                    For i = 0 To TabWindow.ItemCount - 1
                        If TabWindow.Item(i).tag = "排队叫号" Then
                            Call TabWindow.RemoveItem(i)
                            Exit For
                        End If
                    Next i
                    
                    Set mobjQueue = Nothing
                    
                    Call picWindow_Resize
                End If
            End If
            
            If mlngModule = G_LNG_PACSSTATION_MODULE Then
                If Not mfrmWork_PacsImg Is Nothing And InStr(mstrWorkModule, ";影像图像模块;") > 0 Then
                    '更新影像质量的子菜单和工具栏
                    Call mfrmWork_PacsImg.zlMenu.zlCreateMenu(Me.cbrMain)
                    Call mfrmWork_PacsImg.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
                End If
            End If
            
            '为保持报告菜单能够一直显示，这里需要对报告菜单进行创建
            If Not mobjWork_Report Is Nothing And (InStr(mstrWorkModule, ";影像报告模块;") > 0 Or InStr(mstrWorkModule, ";病理诊断模块;") > 0) Then
                Call mobjWork_Report.zlInitModule(mlngModule, mstrPrivs, mlngCur科室ID, Me)
                
                '创建报告对应菜单和工具栏（报告编辑器使用不同方式的时候，创建的菜单不同）
                Call mobjWork_Report.zlMenu.zlCreateMenu(Me.cbrMain)
                Call mobjWork_Report.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
            End If
            
            '切换消息的接收科室
            Call mobjMsgCenter.ChangeMsgReceiveDept(mlngCur科室ID)
        End If
        
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
        Call cbrMain.RecalcLayout
        
        '科室切换后，重新刷新科室对应的检查数据
        Call RefreshList
        
        '刷新排队叫号模块数据，如果已经启用
        Call RefreshPacsQueueData
    
        Call FillCurAdviceTxtInfor
        Call FillCurAdviceAppend
        
        '科室切换后，恢复操作提醒的定时器
        timerOperHint.Enabled = True
    End If
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur科室ID
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub RefreshCustomQueryMenu(objQueryMenu As Object, ByVal lngDeptID As Long)
'根据科室Id,刷新自定义查询菜单
    Dim objCurQueryMenu As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    Dim i As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If objQueryMenu Is Nothing Then Exit Sub
    
    Set objCurQueryMenu = objQueryMenu
    
    For i = 1 To objCurQueryMenu.CommandBar.Controls.Count
        objCurQueryMenu.CommandBar.Controls(1).Delete
    Next
    
    
    Set rsTemp = zlDatabase.OpenSQLRecord("select Id, 方案名称, 是否默认 from 影像查询方案 where 使用状态=1 and (所属科室=0 or 所属科室 is null or 所属科室=[1]) order by 所属科室 desc, 方案序号", "创建查询菜单", lngDeptID)
    
    With objCurQueryMenu.CommandBar
        If rsTemp.RecordCount > 0 Then
            '创建自定义的查询方案
            i = 65
            While Not rsTemp.EOF
                Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CustomQuery * 1000# + Val(Nvl(rsTemp!ID)), Nvl(rsTemp!方案名称) & "(&" & Chr(i) & ")", "", 0, False)
                
                i = i + 1
                If Chr(i) = "F" Or Chr(i) = "C" Then i = i + 1
                
                If Val(Nvl(rsTemp!是否默认)) = 1 Then
                    cbrControl.IconId = 3558
                End If
                
                Call rsTemp.MoveNext
            Wend
        End If
            
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CustomQuery, "综合查询", "", 721, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ConfigQuery, "查询配置", "", 3965, False)
    End With
    
End Sub

Private Sub Menu_View_Refresh_click()
On Error GoTo ErrHandle
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_Help_Web_Home_click()
On Error GoTo ErrHandle
    zlHomePage hWnd
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer
    
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If
        
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next
    
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cboTimes_Click()
On Error GoTo ErrHandle
    Dim lngAdviceId As Long
    
    If cboTimes.ListCount <= 1 Then Exit Sub
    If cboTimes.tag = "" Then Exit Sub '此时cbotime项目未增加完成，属listindex赋值触发
    
    lngAdviceId = cboTimes.ItemData(cboTimes.ListIndex)
    
    If lngAdviceId = mListAdviceInf.lngAdviceId Then
        Call ufgStudyList_OnSelChange
        Exit Sub  '当次与当前选中医嘱ID相同时不由本函数控制
    End If

    mblnIsHistory = True
    
    '以下三个过程调用有先后顺序，勿调换
    mcurAdviceInf = GetAdviceDetailInf(lngAdviceId)
    
    Call FillCurAdviceTxtInfor    '填充右上方病人基本信息
    Call FillCurAdviceAppend   '填充左下角医嘱附件
    
    '选择了全部科室后，再且切换了科室
    If mlngCur科室ID <> mcurAdviceInf.lngExeDepartmentId And mblnAllDepts = True Then
        mlngCur科室ID = mcurAdviceInf.lngExeDepartmentId
        mstrCur科室 = GetDeptName(mlngCur科室ID, mstrCanUse科室)
    End If
    
    Call ShowTab    '根据病人提供不同选项卡
    
    Call RefreshModuleAdviceInf
    Call RefreshTabWindow   '刷新子窗体

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function GetDeptName(lngDeptID As Long, strDeptStrings As String) As String
'通过可用的科室串，读取指定科室ID的科室名称
On Error GoTo ErrHandle
    Dim strDepts() As String
    Dim i As Integer
    
    strDepts = Split(strDeptStrings, "|")
    For i = 0 To UBound(strDepts)
        If Split(strDepts(i), "_")(0) = lngDeptID Then
            GetDeptName = Split(strDepts(i), "_")(1)
            Exit For
        End If
    Next i
Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
End Function


Private Sub cboTimes_DropDown()
On Error GoTo ErrHandle
    Call SendMessage(cboTimes.hWnd, &H160, 500, 0)
ErrHandle:
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim objControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim objTmpControl As CommandBarControl
    
    Dim i As Integer, j As Integer
    Dim strTemp As String
    Dim strCardName As String
    Dim strCardText As String
    Dim lngPatientID As Long
    Dim strAllRooms As String
    Dim strRoomName As String
    Dim strStudyTypes As String
    
    Select Case control.ID
        Case ID_查找方式
            If control.IconId = 3 Then
                control.IconId = 4
                
                mstrLocateWay = PatiIdentify.GetCurCard.名称
                '在快速工具栏点击定位和查找时，更新刷卡控件IDKindStr时，会出发ItemClick事件，
                '导致无法分别记录定位和查找字段信息，所以用此变量标记，为true时不触发ItemClick事件
                mblnAssignment = True
                PatiIdentify.IDKindStr = InitCardType(Replace(CONST_STR_FIND_CARD_TYPE, "[------]", GetStudyNumberDisplayName))
                PatiIdentify.IDKindIDX = PatiIdentify.GetKindIndex(mstrFindWay)
                mblnAssignment = False
                
                cbrdock.FindControl(, ID_开始查找).Caption = "开始查找"
                
                Call zlDatabase.SetPara("定位查找方式", 1, glngSys, mlngModule)
            Else
                control.IconId = 3
                
                mstrFindWay = PatiIdentify.GetCurCard.名称
                
                Call subRefreshFilterCondition("", "")
                Call RefreshList
                
                PatiIdentify.tag = ""
                mblnAssignment = True
                PatiIdentify.IDKindStr = InitCardType(Replace(CONST_STR_LOCAL_CARD_TYPE, "[------]", GetStudyNumberDisplayName))
                PatiIdentify.IDKindIDX = PatiIdentify.GetKindIndex(mstrLocateWay)
                mblnAssignment = False
                
                cbrdock.FindControl(, ID_开始查找).Caption = "开始定位"
                
                Call zlDatabase.SetPara("定位查找方式", 0, glngSys, mlngModule)
            End If
            
            Exit Sub
            
            
            
        Case ID_门诊
            mblncmd门诊 = Not control.Checked
        Case ID_住院
            mblncmd住院 = Not control.Checked
        Case ID_外诊
            mblncmd外诊 = Not control.Checked
        Case ID_体检
            mblncmd体检 = Not control.Checked
            
            
            
        Case ID_已缴
            mblncmd已缴 = Not control.Checked
            
        Case ID_未缴
            mblncmd未缴 = Not control.Checked
            
        Case ID_记账
            mblncmd记账 = Not control.Checked
            
        Case ID_补缴
            mblncmd补缴 = Not control.Checked
            
        Case ID_无费
            mblncmd无费 = Not control.Checked
            
        Case ID_影像类别 + 1 To ID_影像类别 + 40
            control.Checked = Not control.Checked
            mblncmd影像类别(control.ID - ID_影像类别 - 1) = control.Checked
            
            If control.Checked = True Then
                mintcmd影像类别 = mintcmd影像类别 + 1
            Else
                mintcmd影像类别 = mintcmd影像类别 - 1
            End If
            
            Set objControl = cbrdock.FindControl(, ID_影像类别)
            
            If mintcmd影像类别 = 0 Then
                strTemp = "类别"
            Else
                strTemp = ""
                For i = 1 To objControl.CommandBar.Controls.Count
                    If objControl.CommandBar.FindControl(, ID_影像类别 + i).Checked = True Then
                        strTemp = IIf(strTemp = "", Mid(objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption, 1, InStr(objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption, "(") - 1), strTemp & "," & Mid(objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption, 1, InStr(objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption, "(") - 1))
                        strStudyTypes = strStudyTypes & "," & Mid(objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption, 1, InStr(objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption, "(") - 1)
                    End If
                Next i
            End If
            
            If strStudyTypes <> "" Then strStudyTypes = Mid(strStudyTypes, 2)
            
            Call InitStudyPlace(cbrdock.FindControl(, ID_检查部位), cbrPopControl, strStudyTypes)
            
            If strTemp = "类别" Or strTemp = "" Then
                objControl.ToolTipText = "根据影像类别进行过滤"
            Else
                objControl.ToolTipText = "显示影像类别为[" & strTemp & "]的检查"
            End If
            
            objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
            
        Case ID_影像执行间 + 1 To ID_影像执行间 + 40
            control.Checked = Not control.Checked
            mblncmd影像执行间(control.ID - ID_影像执行间 - 1) = control.Checked
            
            If control.Checked = True Then
                mintcmd影像执行间 = mintcmd影像执行间 + 1
            Else
                mintcmd影像执行间 = mintcmd影像执行间 - 1
            End If
            
                        
            Set objControl = cbrdock.FindControl(, ID_影像执行间)
            
            mstrSelQueueRooms = ""
            
            If mintcmd影像执行间 <= 0 Then
                strTemp = "执行间"
                mintcmd影像执行间 = 0
            Else
                strTemp = ""
                For i = 1 To objControl.CommandBar.Controls.Count
                
                    Set cbrPopControl = objControl.CommandBar.FindControl(, ID_影像执行间 + i)
                    If Not cbrPopControl Is Nothing Then
                        strRoomName = Mid(cbrPopControl.Caption, 1, InStr(cbrPopControl.Caption, "(") - 1)
                        
                        If cbrPopControl.Checked = True Then
                            strTemp = IIf(strTemp = "", strRoomName, strTemp & "," & strRoomName)
                            
                            If mstrSelQueueRooms <> "" Then mstrSelQueueRooms = mstrSelQueueRooms & ","
                            mstrSelQueueRooms = mstrSelQueueRooms & cbrPopControl.ToolTipText
                        End If
                    End If
                    
                Next i
            End If
            
            '如果没有选择任何执行间，则默认显示所有执行间的数据
            If Trim(mstrSelQueueRooms) = "" Then mstrSelQueueRooms = mstrAllQueueRooms
            
            If strTemp = "执行间" Or strTemp = "" Then
                objControl.ToolTipText = "根据影像执行间进行过滤"
            Else
                objControl.ToolTipText = "显示影像执行间为[" & strTemp & "]的检查"
            End If
            
            '当菜单数量大于6个字符时，后面的字符使用省略号显示
            objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
 
            '如果启用了排队叫号，则刷新排队叫号数据
            Call RefreshPacsQueueData
        
        Case ID_检查部位 To ID_检查部位 + 4500
            control.Checked = Not control.Checked

            Set objControl = cbrdock.FindControl(, ID_检查部位)

            strTemp = ""
            mstrcmd部位 = ""
            
            For i = 1 To objControl.CommandBar.Controls.Count
                Set objTmpControl = objControl.CommandBar.Controls(i)

                If Not objTmpControl Is Nothing Then
                    If objTmpControl.Checked = True Then
                        strTemp = IIf(strTemp = "", objTmpControl.Caption, strTemp & "," & objTmpControl.Caption)
                        mstrcmd部位 = mstrcmd部位 & "," & objTmpControl.Category
                    End If
                End If
            Next i
            
            If control.Checked Then
                mstrcmd部位分组 = mstrcmd部位分组 & "," & control.Caption
            Else
                mstrcmd部位分组 = Replace(mstrcmd部位分组, "," & control.Caption, "")
            End If

            If strTemp = "" Then
                strTemp = "部位"
                objControl.ToolTipText = "根据检查部位进行过滤"
            Else
                objControl.ToolTipText = "显示检查部位为[" & strTemp & "]的检查"
            End If

            objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
            
        Case ID_登记
            mblncmd登记 = Not control.Checked
        Case ID_报到
            mblncmd报到 = Not control.Checked
        Case ID_检查
            mblncmd检查 = Not control.Checked
        Case ID_报告
            mblncmd报告 = Not control.Checked
        Case ID_审核
            mblncmd审核 = Not control.Checked
        Case ID_驳回
            mblncmd驳回 = Not control.Checked
        Case ID_完成
            mblncmd完成 = Not control.Checked
            
            
            
        Case ID_病理类别_常规
            mblncmd常规 = Not control.Checked
        Case ID_病理类别_快速石蜡
            mblncmd快速石蜡 = Not control.Checked
        Case ID_病理类别_冰冻
            mblncmd冰冻 = Not control.Checked
        Case ID_病理类别_细胞
            mblncmd细胞 = Not control.Checked
        Case ID_病理类别_尸检
            mblncmd尸检 = Not control.Checked
        Case ID_病理类别_会诊
            mblncmd会诊 = Not control.Checked
            
            
            
        Case ID_本次住院
            control.Checked = Not control.Checked
            mblncmd本次 = Not mblncmd本次
        Case ID_开始查找
            Call StartReadCard
            Call SaveFilterCmd
            
            Exit Sub
    End Select
    
    '保存快速工具栏参数设置
    Call SaveFilterCmd
    
    cbrdock.RecalcLayout
    
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub subRefreshFilterCondition(ByVal strCardName As String, ByVal strFilter As String)
'------------------------------------------------
'功能：用txtFilter控件的内容更新过滤条件
'参数： strFilter --- 过滤条件
'返回：无
'------------------------------------------------

On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strTemp As String
    
    With SQLCondition
        .姓名 = ""
        .就诊卡 = ""
        .门诊号 = 0
        .住院号 = 0
        .健康号 = ""
        .单据号 = ""
        .检查号 = 0
        .身份证 = ""
        .IC卡 = ""
        .结果阳性 = -1
        .病人ID = 0
        
        Select Case strCardName
            Case "姓名", "姓  名", "姓   名"  '保持与以前的方式兼容
                .姓名 = Trim(strFilter)
                
            Case "就诊卡"
                .就诊卡 = Trim(strFilter)
                
            Case "门诊号"   '快捷方式是“*+数字”,VAL提取前，“*”要特殊处理
                If Left(strFilter, 1) = "*" Then
                    strFilter = Mid(strFilter, 2)
                End If
                .门诊号 = Val(strFilter)
                
            Case "住院号"   '快捷方式是“++数字”
                .住院号 = Val(strFilter)
                
            Case "健康号"
                .健康号 = Trim(strFilter)
                
            Case "单据号"
                If Len(Trim(strFilter)) = 0 Then
                     .单据号 = ""
                Else
                    If Len(Trim(strFilter)) < 8 And Not IsNumeric(Trim(strFilter)) Then
                        strTemp = GetFullNO(0, 0)
                        strTemp = Mid(strTemp, 1, Len(strTemp) - Len(strFilter)) & strFilter
                    Else
                        strTemp = GetFullNO(Nvl(strFilter, 0), 0)
                    End If
                    
                    PatiIdentify.Text = strTemp
                    .单据号 = strTemp
                End If
                
            Case GetStudyNumberDisplayName
                If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                    .检查号 = Val(strFilter)
                Else
                    If Trim(strFilter) = "" Then
                        Exit Sub
                    End If
                    
                    If UCase(Mid(strFilter, Len(strFilter), 1)) = UCase("Z") Then       '如果通过扫描枪，扫描出“Z”打头的号码，说明是制片号
                        strSql = "select 病理号 from 病理检查信息 a, 病理制片信息 b where a.病理医嘱ID=b.病理医嘱Id and b.ID=[1]"
                        Set rsData = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, Mid(strFilter, 1, Len(strFilter) - 1))
                        
                        If rsData.RecordCount > 0 Then
                            .检查号 = Nvl(rsData!病理号)
                            
                            PatiIdentify.Text = .检查号
                        End If
                    ElseIf UCase(Mid(strFilter, Len(strFilter), 1)) = UCase("T") Then   '如果通过扫描枪，扫描出“T”打头的号码，说明是特检制片号
                        strSql = "select 病理号 from 病理检查信息 a, 病理特检信息 b where a.病理医嘱ID=b.病理医嘱Id and b.ID=[1]"
                        Set rsData = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, Mid(strFilter, 1, Len(strFilter) - 1))
                        
                        If rsData.RecordCount > 0 Then
                            .检查号 = Nvl(rsData!病理号)
                            
                            PatiIdentify.Text = .检查号
                        End If
                    Else
                        .检查号 = GetPatholNum(Trim(strFilter))
                    End If
                End If
                
            Case "身份证号", "身份证"
                .身份证 = Trim(strFilter)
                
            Case Else
                .病人ID = Val(strFilter)
        End Select
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetPatholNum(ByVal strSureNum As String) As String
'分解确认号码
    Dim lngFindSplitChar As Long
    
    lngFindSplitChar = InStr(1, strSureNum, "-")
    
    If lngFindSplitChar > 0 Then
        GetPatholNum = UCase(Mid(strSureNum, 1, lngFindSplitChar - 1))
    Else
        GetPatholNum = UCase(strSureNum)
    End If
    
End Function

Private Sub cbrdock_Resize()
On Error GoTo ErrHandle
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    If tabFilter.Visible Then
        '只有病理工作站才显示tab过滤页面
        tabFilter.Top = lngTop
        tabFilter.Left = lngLeft
        tabFilter.Width = PicList.Width
        
        picExeState.Left = lngLeft
        picExeState.Top = lngTop + IIf(tabFilter.Visible, tabFilter.Height, 0)
        picExeState.Width = PicList.Width
    End If
    
    ufgStudyList.Top = IIf(tabFilter.Visible, picExeState.Top + picExeState.Height, lngTop)
    ufgStudyList.Left = lngLeft
    ufgStudyList.Width = PicList.Width
    ufgStudyList.Height = Abs(PicList.Height - lngTop - picAppend.Height - IIf(tabFilter.Visible, tabFilter.Height + picExeState.Height, 0))

    PicLine.Top = lngTop + ufgStudyList.Height + IIf(tabFilter.Visible, tabFilter.Height + picExeState.Height, 0)
    PicLine.Left = PicList
    PicLine.Width = PicList.Width
    PicLine.Height = 90

    picAppend.Top = PicLine.Top + PicLine.Height
    picAppend.Left = lngLeft
    picAppend.Width = PicList.Width
    picAppend.Height = PicList.Height - lngTop - ufgStudyList.Height - IIf(tabFilter.Visible, tabFilter.Height + picExeState.Height, 0)

ErrHandle:
End Sub


Private Sub Form_Activate()
On Error GoTo ErrHandle
    '判断当前工作模块是否影像采集模块，如果是，则判断采集模块是否初始化，如果已经初始化，则退出该过程，否则就对其进行初始化，并显示
    '因为在同一导航台中，如果同时打开病理，视频采集模块将被切换，当另一系统退出时，采集模块也将被释放，因此切换回当前系统后，需要判断是否从新初始化采集模块
    If Not mblnInitOk Then Exit Sub
    If TabWindow.Selected Is Nothing Then Exit Sub
    If TabWindow.Selected.tag <> "影像采集" Then Exit Sub
    
    If mblnUseActivexCapture Then
        '使用ActivexExe方式的视频采集处理
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlUpdateStudyInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved)
            Call mobjWork_ActiveVideo.zlRefreshVideoWindow
            Call mobjWork_ActiveVideo.zlRefreshData(True)
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '加载工作模块时，不允许退出窗口
    If Not mblnInitOk Then Cancel = True
    
    If mblnUseActivexCapture Then
        'TODO:等待加载完成后，才允许退出视频
    End If
End Sub


Private Sub labStudyNum_Change()
On Error GoTo ErrHandle
    Call picAppend_Resize
ErrHandle:
End Sub

Private Sub lbl个人信息_Change()
On Error GoTo ErrHandle
    Call picAppend_Resize
ErrHandle:
End Sub

Private Sub mfrmRISRequest_HaveRegist()
    Dim strQueueName As String
    Dim strCodeNo As String
    With mfrmRISRequest
        If .mlngResultState <> 0 Then '成功返回
            '如果启用排队叫号，则报到后需要插入排队叫号队列......
            If mSysPar.blnUseQueue And Not mobjQueue Is Nothing And .mlngResultState = 2 Then
                '设置需要插入的队列名称
                If .mstrTechnicRoom = "" Then
                    '如果未空，则需要插入该检查项目对应的项目分组或者科室的队列中
                    Call mobjQueue.zlGetInQueueInf(mListAdviceInf.lngAdviceId, .mlngCurDeptId, strQueueName, strCodeNo)
                Else
                    '如果不为空，则写入对应的执行间名称
                    strQueueName = .mstrCur科室 & "-" & .mstrTechnicRoom
                    strCodeNo = mobjQueue.zlGetTechnicRoomCodeNo(.mstrTechnicRoom, .mlngCurDeptId)
                End If
                
                Call mobjQueue.zlInPacsQueue(.mlngAdviceId, .mstrPatientName, .mlngCurDeptId, strQueueName, .mstrTechnicRoom, strCodeNo)
            End If
            
            '发送新申请消息
            Call mobjMsgCenter.Send_Msg_Request(.mlngAdviceId)
        End If
    End With
End Sub

Private Sub mobjCaptureHot_OnKeyBoardLHook(ByVal lngMsg As Long, ByVal lngVkCode As Long, ByVal lngScanCode As Long, ByVal lngFlags As Long)
On Error GoTo ErrHandle
    Dim lngWindowPID As Long
    Dim lngVideoPID As Long
    Dim lngCurrentPID As Long

    If lngMsg <> WM_KEYDOWN Then Exit Sub
    If Trim(mstrCaptureHot) = "" And Trim(mstrCaptureAfterHot) = "" And Trim(mstrCaptureAfterTagHot) = "" Then Exit Sub
    
    mCaptureMsg.lngMsg = lngMsg
    mCaptureMsg.lngVirtualKey = lngVkCode
    mCaptureMsg.lngScanKey = lngScanCode
    mCaptureMsg.lngFlags = lngFlags
    
    '不能直接在Hook回调过程中使用ActiveExe对象的相关方法，否则会产生未知界面错误
    timerCapture.Enabled = True
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjEvent_OnWork(objEvent As Object, ByVal lngWorkType As TWorkEventType, ByVal lngAdviceId As Long, ByVal other As Variant)
'相应工作模块执行操作后触发的事件
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strRoom As String
    Dim i As Integer
    Dim j As Integer
    Dim strStudyUID As String
    
    Dim lngcurRow As Long
    Dim lngColIndex As Long
    
    Select Case lngWorkType
        Case TWorkEventType.wetGetImg           '获取图像（QR）***************************************
            Call RefreshList
            
        Case TWorkEventType.wetTechDo           '技师执行***************************************
            If mListAdviceInf.lngAdviceId = lngAdviceId Then
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "检查技师", UserInfo.姓名)
                
                If ufgStudyList.CurText("是否技师确认") = "1" Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, ufgStudyList.DataGrid.RowSel, ufgStudyList.GetColIndex("检查技师")) = imgList.ListImages("检查技师").Picture
                    ufgStudyList.CurText("检查技师") = UserInfo.姓名
                Else
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, ufgStudyList.DataGrid.RowSel, ufgStudyList.GetColIndex("检查技师")) = Nothing
                    ufgStudyList.CurText("检查技师") = IIf(ufgStudyList.CurText("检查技师") = UserInfo.姓名, "", ufgStudyList.CurText("检查技师"))
                End If
            End If
            
        Case TWorkEventType.wetChangeImgType    '改变影像类型***************************************
            Call RefreshList(lngAdviceId)
        
        Case TWorkEventType.wetLockStudy, TWorkEventType.wetUnLockStudy        '锁定检查,解锁检查
            '修改标签页的显示样式和标题
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Caption Like "*影像采集*" Then
                    If lngWorkType = wetLockStudy Then
                        TabWindow(i).Image = 10013
                        TabWindow(i).Caption = "【" & other & "】 影像采集"
                    Else
                        TabWindow(i).Image = conMenu_Cap_Dynamic
                        TabWindow(i).Caption = "影像采集"
                    End If
                    Exit For
                End If
            Next i
            
            '刷新嵌入报告中的缩略图图像或者视频采集的图像
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngWorkType, lngAdviceId, other)
            
        Case TWorkEventType.wetCaptureFirstImg, TWorkEventType.wetDelAllImg, TWorkEventType.wetUpdateImg  '采集第一幅图像***************************************
            '更新检查列表
            
            strStudyUID = other
            
            If lngWorkType = wetCaptureFirstImg Then
                '回写检查执行间
                If mSysPar.lngQueueWay = 1 And mSysPar.blnUseQueue Then
                    strRoom = mSysPar.strLocalRoom
                    If InStr(strRoom, "-") > 0 Then strRoom = Mid(mSysPar.strLocalRoom, 1, InStr(mSysPar.strLocalRoom, "-") - 1)
        
                    strSql = "zl_影像检查_更新执行间(" & lngAdviceId & ",'" & strRoom & "','" & NeedName(mSysPar.strLocalRoom) & "')"
                    Call zlDatabase.ExecuteProcedure(strSql, GetWindowCaption)
                End If
                
                '更新检查列表
                Call UpdateStudyListState(lngAdviceId, strStudyUID, True, True)
                
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateSync(lngAdviceId)
            ElseIf lngWorkType = wetDelAllImg Then
                '更新检查列表
                Call UpdateStudyListState(lngAdviceId, strStudyUID, False, True)
                
                '发送状态同步消息
                Call mobjMsgCenter.Send_Msg_StateCancel(lngAdviceId)
            End If


            If Val(ufgStudyList.CurKeyValue) <> CStr(lngAdviceId) Then Exit Sub
            
            '刷新嵌入报告中的缩略图图像或者视频采集的图像
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngWorkType, lngAdviceId, other)
            
            '刷新嵌入特检报告界面右下角缩略图图像
            If lngWorkType = wetUpdateImg Then If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
        Case wetChangeUser
            '交换用户时，需要先判断报告是否需要保存
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
        
            Call ChangeUser
            
            '交换用户后，需要刷新报告编辑器，因为用户交换后，原有报告的编辑用户或者创建用户需要进行更新
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
        Case wetPatholRequest       '病理申请
            Call RefreshList(lngAdviceId)
            
        Case wetPatholQuality       '病理质量
            lngcurRow = ufgStudyList.FindRowIndex(CStr(lngAdviceId), "医嘱ID", True)
            
            If lngcurRow > 0 Then
                ufgStudyList.Text(lngcurRow, "质量") = other
                
                lngColIndex = ufgStudyList.GetColIndex("质量")
                
                If other = "符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngcurRow, lngColIndex) = vbGreen
                If other = "基本符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngcurRow, lngColIndex) = vbYellow
                If other = "不符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngcurRow, lngColIndex) = vbRed
                
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "综合质量", other)
            End If
        
        Case wetPatholBatSlices     '制片批量处理
            Call RefreshList(lngAdviceId)
            
        Case wetPatholBatSpeExm     '特检批量处理
            Call RefreshList(lngAdviceId)
            
        Case wetSpecimenAccept      '标本核收
            Call RefreshPatholExecuteState(lngAdviceId)
            
            With ufgStudyList
                lngcurRow = .DataGrid.FindRow(CStr(lngAdviceId), , .GetColIndex("医嘱ID"))
                
                If lngcurRow > 0 Then
                    If Trim(.Text(lngcurRow, "病理号")) = "" Then
                        .Text(lngcurRow, "病理号") = other
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "病理号", other)
                        
                        .Text(lngcurRow, "检查状态") = 2
                        
                        .Text(lngcurRow, "检查过程") = "已报到"
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "检查过程", 2)
                        
                        .Text(lngcurRow, "报到时间") = zlDatabase.Currentdate
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "报到时间", zlDatabase.Currentdate)
                        
                        .Text(lngcurRow, "报到人") = UserInfo.姓名
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "报到人", UserInfo.姓名)
                        
                        .Text(lngcurRow, "核收情况") = "已核收"
                        Call ufgStudyList.UpdateSourceData(lngAdviceId, "核收情况", "已核收")
                        
                        labStudyNum.Caption = "[病理号:" & IIf(other <> "", other, "---") & "]"
                        
                        '刷新其他病理模块数据
                        If Not mobjWork_Pathol Is Nothing Then
                            Call mobjWork_Pathol.zlUpdateAdviceInf(lngAdviceId, 0, 2, False)
                            Call mobjWork_Pathol.NotificationRefresh(mtAll)
                        End If
                    End If
                End If
            End With
        
        Case wetSpecimenReject      '标本拒收
        
        Case wetSpecimenSave        '标本保存
            '标本保存后，刷新取材模块数据
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtMaterial)
            
        Case wetMaterialSure        '取材确认
            Call RefreshPatholExecuteState(lngAdviceId)
            
            '刷新制片模块数据
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtSlices)
            
        Case wetMaterialSave        '材块保存
            '刷新制片模块数据
            If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(TModuleType.mtSlices)
            
        Case wetSlicesSure          '制片确认
            Call RefreshPatholExecuteState(lngAdviceId)
    
        Case wetSpeExamSure         '特检确认
            Call RefreshPatholExecuteState(lngAdviceId)
            
        Case wetViewEprReport       '预览电子病历报告
            Dim strRepInf() As String
            
            strRepInf = Split(other & ",,", ",")
            
            If Val(strRepInf(0)) <= 0 Then Exit Sub
            
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.ViewEPRReport(Val(strRepInf(0)), IIf(Val(strRepInf(1)) = 1, True, False))
        
        Case wetViewPacsImage       '预览Pacs图像
            '超过100张图像的序列，默认每隔5张传一张
            Call OpenViewer(2, mobjPacsCore, lngAdviceId, False, Me, , , mSysPar.blnLocalizerBackward)
            
        Case wetRejectReport        '报告被驳回
            lngcurRow = ufgStudyList.DataGrid.FindRow(CStr(lngAdviceId), , ufgStudyList.GetColIndex("医嘱ID"))
            
            If lngcurRow <= 0 Then Exit Sub
                        
            ufgStudyList.Text(lngcurRow, "检查过程") = "已驳回"
            ufgStudyList.DataGrid.Cell(flexcpBackColor, lngcurRow, 1, lngcurRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已驳回
            
            Call ufgStudyList.UpdateSourceData(lngAdviceId, "检查过程", -1)
            
            '发送状态同步消息
            Call mobjMsgCenter.Send_Msg_StateSync(lngAdviceId)
        Case wetPrintFilm
            '处理胶片打印消息
            lngcurRow = ufgStudyList.DataGrid.FindRow(CStr(lngAdviceId), , ufgStudyList.GetColIndex("医嘱ID"))
            
            If lngcurRow <= 0 Then Exit Sub
            
            ufgStudyList.Text(lngcurRow, "胶片打印") = "1"
            Call ufgStudyList.UpdateSourceData(lngAdviceId, "胶片打印", "1")
                        
        End Select
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub RefreshPatholExecuteState(ByVal lngAdviceId As Long)
'更新病理执行状态
    Dim lngcurRow As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select 检查类型,取材过程,制片过程,免疫过程,分子过程,特染过程 from 病理检查信息 where 医嘱Id=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    lngcurRow = ufgStudyList.DataGrid.FindRow(CStr(lngAdviceId), , ufgStudyList.GetColIndex("医嘱ID"))
        
    If lngcurRow > 0 Then
        ufgStudyList.Text(lngcurRow, "病理执行状态") = GetPatholExecuteState(rsData)
        ufgStudyList.Text(lngcurRow, "检查类别") = Decode(Nvl(rsData!检查类型), 1, "冰冻", 2, "细胞", 3, "会诊", 4, "尸检", 5, "快速石蜡", "常规")
        
    End If
End Sub

Private Function GetPatholExecuteState(rsData As ADODB.Recordset) As String
    Dim strState As String

    strState = ""
    
    If Nvl(rsData!取材过程) = 1 Then strState = "需取材"

    If Nvl(rsData!制片过程) = 1 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "需制片"
    End If
    
    If Nvl(rsData!免疫过程) = 1 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "需免疫"
    End If
    
    If Nvl(rsData!分子过程) = 1 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "需分子"
    End If
    
    If Nvl(rsData!特染过程) = 1 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "需特染"
    End If
    
    
    If Nvl(rsData!制片过程) = 2 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "制片接受"
    End If
    
    If Nvl(rsData!免疫过程) = 2 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "免疫接受"
    End If
    
    If Nvl(rsData!分子过程) = 2 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "分子接受"
    End If
    
    If Nvl(rsData!特染过程) = 2 Then
        If strState <> "" Then strState = strState & ","
        strState = strState & "特染接受"
    End If
    
    If Trim(strState) = "" Then strState = ""
    
    GetPatholExecuteState = strState
End Function

Private Sub mobjMsgCenter_OnRecevieMsg(ByVal strMsgItemIdentity As String, ByVal strXmlContext As String, rsData As ADODB.Recordset, objMsgPro As clsMipModule, objXML As clsXML)
'消息接收处理
    Dim lngRowIndex As Long
    Dim lngAdviceId As Long
    Dim lngStudyState As Long
    Dim strHint As String
    Dim strSql As String
    Dim rsReport As ADODB.Recordset
    Dim rsDataMulite As ADODB.Recordset
    Dim rsDataMuliteClone As ADODB.Recordset
    Dim strCurNo As String
    Dim strNodeId As String
    Dim lngChargeState As Long
    
    
    lngAdviceId = 0
    
    '获取消息中对应的医嘱ID数据
    If strMsgItemIdentity = G_STR_MSG_ZLHIS_PACS_003 Then
        rsData.Filter = "node_name='study_order_id'"
    Else
        rsData.Filter = "node_name='order_id'"
    End If
    
    If rsData.RecordCount > 0 Then
        lngAdviceId = Val(Nvl(rsData!node_value))
    End If
    
    
    Select Case strMsgItemIdentity
        Case G_STR_MSG_ZLHIS_CIS_017    '检查申请
            '弹出消息提示@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "患者 " & Nvl(rsData!node_value) & " 需要进行检查，请及时处理。"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
            
            '从数据库中刷新数据
            Call RefreshList(0, True)
            
        Case G_STR_MSG_ZLHIS_CIS_024    '医嘱撤销
            '弹出撤销提示@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "患者 " & Nvl(rsData!node_value) & " 的检查医嘱已被撤销。 "
        
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
        
            '判断医嘱ID是否在列表中存在，如果存在则删除对应的行
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "医嘱ID")
            If lngRowIndex > 0 Then
                Call ufgStudyList.SyncText(lngRowIndex, "检查过程", "已拒绝", False)
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "执行状态", 2)
            End If
            
            '执行UpdateSourceData后将AdoData.Filter清空
            ufgStudyList.AdoData.Filter = ""
            
            '根据当前选择的医嘱判断是否需要刷新列表
            Call RefreshList(IIf(lngAdviceId = mcurAdviceInf.lngAdviceId, lngAdviceId, 0), False)
            
        Case G_STR_MSG_ZLHIS_CIS_025    '危急值阅读
            '由消息平台配置弹出提示
            
        Case G_STR_MSG_ZLHIS_CHARGE_003 '门诊患者划价单据
            '刷新收费状态显示
            '根据单据号查找对应的医嘱ID
            rsData.Filter = "node_name='bill_no'"
            rsData.Sort = "node_name"
            If rsData.RecordCount <= 0 Then
                Exit Sub
            End If
            
            '获取所有单据的信息
            If objXML.GetMultiNodeRecord("charge_bill", rsDataMulite) = False Then Exit Sub
            
            Set rsDataMuliteClone = rsDataMulite.Clone
            
            rsDataMulite.Filter = "node_name='charge_bill'"
            
            If rsDataMulite.RecordCount <= 0 Then Exit Sub
            
            Do While Not rsDataMulite.EOF
                '获取单据charge_bill对应ID，在存在多个单据的情况下，需根据charge_bill的D查找对应的单据信息
                strNodeId = Val(Nvl(rsDataMulite!ID))
                
                '获取charge_bill节点下的单据号，根据charge_bill的id关联
                rsDataMuliteClone.Filter = "parent_id=" & strNodeId & " and node_name='bill_no'"
                If rsDataMuliteClone.RecordCount > 0 Then strCurNo = Nvl(rsDataMuliteClone!node_value)
                
                '获取charge_bill节点下的单据费用状态,根据charge_bill的id关联
                rsDataMuliteClone.Filter = "parent_id=" & strNodeId & " and node_name='charge_state'"
                If rsDataMuliteClone.RecordCount > 0 Then lngChargeState = Val(Nvl(rsDataMuliteClone!node_value))
                
                lngRowIndex = ufgStudyList.FindRowIndex(strCurNo, "单据号")
                
                If lngChargeState = 2 Then  '=2表示已经收费
                    ufgStudyList.Text(lngRowIndex, "收费") = " "
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRowIndex, ufgStudyList.GetColIndex("收费")) = imgList.ListImages("收费").Picture
                    
                    lngAdviceId = Val(ufgStudyList.KeyValue(lngRowIndex))
                    
                    If mListAdviceInf.lngAdviceId = lngAdviceId Then
                        '刷新列表下方的费用状态显示
                        lblCash.Caption = "收"
                        lblCash.ForeColor = &H8000&
                    End If
                End If
                
                rsDataMulite.MoveNext
            Loop
        
        Case G_STR_MSG_ZLHIS_PACS_001   '检查报告完成，检查完成才算检查报告最终完成
            '更新列表中的显示状态
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "医嘱ID")
            If lngRowIndex > 0 Then
                Call ufgStudyList.SyncText(lngRowIndex, "检查过程", "已完成", False)
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "检查过程", 6)
                
                '执行UpdateSourceData后将AdoData.Filter清空
                ufgStudyList.AdoData.Filter = ""
                
                '根据更新后的数据，刷新列表显示
                Call RefreshList(IIf(lngAdviceId = mcurAdviceInf.lngAdviceId, lngAdviceId, 0), False)
            End If
            
        Case G_STR_MSG_ZLHIS_PACS_002, G_STR_MSG_ZLHIS_PACS_003  '检查状态同步与检查状态回退处理
            '如果报告被驳回，需要弹出提醒@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='study_cur_state'"
            If Nvl(rsData!node_value) = -1 Then
                
                
                '需要判断当前用户是否为报告人
                strSql = "select 报告人 from 影像检查记录 where 医嘱ID=[1]"
                Set rsReport = zlDatabase.OpenSQLRecord(strSql, "查询报告人", lngAdviceId)
                If rsReport.RecordCount > 0 Then
                    If Nvl(rsReport!报告人) = UserInfo.姓名 Then
                        '弹出消息
                        rsData.Filter = "node_name='patient_name'"
                        strHint = "患者" & Nvl(rsData!node_value) & "的报告已被驳回，请注意处理。"
                        
                        Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
                    End If
                End If
            End If
            
            
        
            '刷新列表对应显示
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "医嘱ID")
            If lngRowIndex > 0 Then
            
                rsData.Filter = "node_name='study_cur_state'"
                If rsData.RecordCount > 0 Then
                    lngStudyState = Val(Nvl(rsData!node_value))
                    Call ufgStudyList.UpdateSourceData(lngAdviceId, "检查过程", lngStudyState)
                    
                    '根据更新后的数据，刷新列表显示
                    Call RefreshList '(IIf(lngAdviceID = mcurAdviceInf.lngAdviceID, lngAdviceID, 0), False)
                End If
            End If
        
        Case G_STR_MSG_ZLHIS_PACS_004   '检查报告撤销
            '更新列表中的显示状态
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "医嘱ID")
            If lngRowIndex > 0 Then
                rsData.Filter = "node_name='cur_state'"
                If rsData.RecordCount > 0 Then
                    lngStudyState = Val(Nvl(rsData!node_value))
                    Call ufgStudyList.UpdateSourceData(lngAdviceId, "检查过程", lngStudyState)
                    
                    ufgStudyList.AdoData.Filter = ""
                    '根据更新后的数据，刷新列表显示
                    Call RefreshList(IIf(lngAdviceId = mcurAdviceInf.lngAdviceId, lngAdviceId, 0), False)
                End If
            End If
            
        
        Case G_STR_MSG_ZLHIS_PACS_005   '检查危急值通知
            '在科室内弹出危急提醒@@@@@@@@@@@@@@@@@@@@
            rsData.Filter = "node_name='patient_name'"
            strHint = "患者 " & Nvl(rsData!node_value) & "的"
            
            rsData.Filter = "node_name='check_item_title'"
            strHint = strHint & "检查项目 " & Nvl(rsData!node_value) & " 产生危急情况。"
            
            Call objMsgPro.ShowMessage(strMsgItemIdentity, strHint)
        
        
            '更新列表中的显示状态
            lngRowIndex = ufgStudyList.FindRowIndex(lngAdviceId, "医嘱ID")
            If lngRowIndex > 0 Then
                Call ufgStudyList.SyncText(lngRowIndex, "危急", " ", False)
                Call ufgStudyList.UpdateSourceData(lngAdviceId, "危急", 1)
                
                ufgStudyList.AdoData.Filter = ""
                Call RefreshList(IIf(lngAdviceId = mcurAdviceInf.lngAdviceId, lngAdviceId, 0), False)
            End If
            
    End Select
    
    
End Sub

Private Sub mobjPacsCore_AfterSaveOuterImage(strStudyUID As String)
    '保存了外部图像，刷新图像的序列列表
On Error GoTo ErrHandle
    
    '没有记录则退出
    If mListAdviceInf.lngAdviceId = 0 Then Exit Sub
    
    '是当前的检查，才刷新检查的序列列表
    If mListAdviceInf.strStudyUID = strStudyUID Then
        Call mfrmWork_PacsImg.zlRefreshFace(True)
    End If
    
    Exit Sub
ErrHandle:
    '不处理
End Sub


Public Sub OnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
'视频采集操作回调事件
    mVideoEventInf.vetEventType = lngEventType
    mVideoEventInf.lngAdviceId = lngAdviceId
    mVideoEventInf.lngSendNo = lngSendNo
    mVideoEventInf.strOtherInf = strOther

    timerVideoEvent.Enabled = True
End Sub

Public Sub OnImgLoadState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean)
    If blnLoadFinish Then
        Me.stbThis.Panels(5).Text = ""
    Else
        If blnUpLoad Then
            Me.stbThis.Panels(5).Text = "正在上传文件，请勿关闭工作站！"
        Else
            Me.stbThis.Panels(5).Text = "正在下载文件，请勿关闭工作站！"
        End If
    End If
End Sub

Public Sub OnState(ByVal blnLoadFinish As Boolean, ByVal blnUpLoad As Boolean, ByVal lngThreadId As Long)
    Call mdlPACSWork.OnState(blnLoadFinish, blnUpLoad, lngThreadId)
End Sub

Public Sub OnDockClose()
'浮动窗口关闭回调事件
End Sub

Private Sub DoOnStateChange(ByVal lngEventType As TVideoEventType, ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal strOther As String)
'相应工作模块执行操作后触发的事件
On Error GoTo ErrHandle
    Dim strSql As String
    Dim strRoom As String
    Dim strStudyUID As String
    Dim i As Long
    
    Select Case lngEventType
        Case TVideoEventType.vetLockStudy, TVideoEventType.vetUnLockStudy         '锁定检查,解锁检查
            '修改标签页的显示样式和标题
            For i = 0 To TabWindow.ItemCount - 1
                If TabWindow(i).Caption Like "*影像采集*" Then
                    If lngEventType = vetLockStudy Then
                        TabWindow(i).Image = 10013
                        TabWindow(i).Caption = "【" & strOther & "】 影像采集"
                    Else
                        TabWindow(i).Image = conMenu_Cap_Dynamic
                        TabWindow(i).Caption = "影像采集"
                    End If
                    Exit For
                End If
            Next i
            
     
            '刷新嵌入报告中的缩略图图像或者视频采集的图像
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)

            
        Case TVideoEventType.vetCaptureFirstImg, TVideoEventType.vetDelAllImg, TVideoEventType.vetUpdateImg  '采集第一幅图像***************************************
            '更新检查列表
            
            strStudyUID = strOther
            
            If lngEventType = TVideoEventType.vetCaptureFirstImg Then
                '报到时执行费用或不为影像采集系统时执行费用
                If mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngMoneyExeModle = 1 Then
                    strSql = "Zl_影像费用执行(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
                    
                    Call zlDatabase.ExecuteProcedure(strSql, "执行检查费用")
                End If
        
                '回写检查执行间
                If mSysPar.lngQueueWay = 1 And mSysPar.blnUseQueue Then
                    strRoom = mSysPar.strLocalRoom
                    If InStr(strRoom, "-") > 0 Then strRoom = Mid(mSysPar.strLocalRoom, 1, InStr(mSysPar.strLocalRoom, "-") - 1)
        
                    strSql = "zl_影像检查_更新执行间(" & lngAdviceId & ",'" & strRoom & "','" & NeedName(mSysPar.strLocalRoom) & "')"
                    Call zlDatabase.ExecuteProcedure(strSql, GetWindowCaption)
                End If
                
                '更新检查列表
                Call UpdateStudyListState(lngAdviceId, strStudyUID, True, True)
                
                If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
            ElseIf lngEventType = TVideoEventType.vetDelAllImg Then
                '更新检查列表
                Call UpdateStudyListState(lngAdviceId, strStudyUID, False, True)
            End If


            If Val(ufgStudyList.CurKeyValue) <> CStr(lngAdviceId) Then Exit Sub
            
            '刷新嵌入报告中的缩略图图像或者视频采集的图像
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
            
            '刷新嵌入特检报告界面右下角缩略图图像
            If lngEventType = TVideoEventType.vetUpdateImg Then If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtProRep)
        Case TVideoEventType.vetAfterUpdateImg
            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.UpdateVideoCaptureState(lngEventType, lngAdviceId, strOther)
            
    End Select
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub AfterReleationImage(ByVal lngAdviceId As Long, ByVal lngSendNo As Long, ByVal intStep As Integer, ByVal lngReleationType As Long, ByVal blnUseMenuReleation As Boolean)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If lngReleationType = 1 Then
        If InStr("345", intStep) > 0 Then
            gstrSQL = "Select 检查uid From 影像检查记录 Where  医嘱ID=[1] And 发送号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceId, lngSendNo)
            
            If IsNull(rsTemp!检查uid) Then
                '设置影像检查状态，如果当前医嘱已经没有图像，而且检查过程为3，则修改为2
                If intStep = 3 Then
                    gstrSQL = "Zl_影像检查_State(" & lngAdviceId & "," & lngSendNo & ",2,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
                    zlDatabase.ExecuteProcedure gstrSQL, "取消关联"
                End If
            End If
        End If
    Else
        '设置影像检查状态，如果原来的状态是已报到，则修改成已检查，
        If intStep = 2 Then
            '如果病人已经有图像，则修改成已检查
            strSql = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "检查是否有图像", lngAdviceId)
            
            If Not IsNull(rsTemp!检查uid) Then
                strSql = "Zl_影像检查_State(" & lngAdviceId & "," & lngSendNo & ",3,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
                zlDatabase.ExecuteProcedure strSql, "关联影像"
            End If
        End If
    End If
    
    Call RefreshList
    
    If mblnUseActivexCapture Then
        '使用ActivexExe视频采集图像刷新处理
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlRefreshData(True)
            Call mobjWork_ActiveVideo.zlUpdateAfterCaptureInfo
        End If
    End If
    
    If Not mfrmWork_PacsImg Is Nothing Then
        Call mfrmWork_PacsImg.zlRefreshFace(True)
    End If
    
    If Not mobjWork_Report Is Nothing And blnUseMenuReleation Then
        Call mobjWork_Report.UpdateVideoCaptureState(TVideoEventType.vetAfterUpdateImg, lngAdviceId, "")
    End If
End Sub

Private Sub mobjWork_Report_AfterReleationImage(ByVal lngOrderID As Long, ByVal lngSendNo As Long, ByVal intStep As Integer, ByVal lngReleationType As Long)
On Error GoTo ErrHandle
    Call AfterReleationImage(lngOrderID, lngSendNo, intStep, lngReleationType, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjWork_Report_DocPluginAction(ByVal actionType As Long, ByVal data As String, ByVal tag As String)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
On Error GoTo ErrHandle
    If actionType = 5 And Trim(data) <> "" And (Trim(tag) = "关联图像" Or Trim(tag) = "取消关联") Then
        '根据医嘱ID获取发送号和检查过程
        strSql = "select b.发送号,b.执行过程 from  影像检查记录 a,病人医嘱发送 b where a.医嘱ID =[1] and a.医嘱id = b.医嘱id"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", Val(data))
        
        If rsTemp.RecordCount > 0 Then
            Call AfterReleationImage(data, Val(Nvl(rsTemp!发送号)), Val(Nvl(rsTemp!执行过程)), IIf(Trim(tag) = "关联图像", 2, 1), False)
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optAccept_Click()
On Error GoTo ErrHandle
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optAll_Click()
On Error GoTo ErrHandle
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optFinal_Click()
On Error GoTo ErrHandle
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub optNeed_Click()
On Error GoTo ErrHandle
    Call RefreshList(, False)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub PatiIdentify_KeyPress(KeyAscii As Integer)
'录入事件
On Error GoTo ErrHandle
    Dim blnCard As Boolean
    Dim lngPatientID As Long
    
    If Trim(PatiIdentify.GetCurCard.名称) = "住院号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    End If
    
    If KeyAscii = 13 Then
        Call StartReadCard
        
        Exit Sub
    End If
    
    If PatiIdentify.GetCurCard.是否刷卡 Then
        blnCard = PatiIdentify.zlIsBrushCard(PatiIdentify.objTxtInput, KeyAscii)
            
        If blnCard And Len(PatiIdentify.Text) = PatiIdentify.GetCardNoLen - 1 And KeyAscii <> 8 Then  '刷卡完毕处理
            PatiIdentify.Text = PatiIdentify.Text & Chr(KeyAscii)
    
            KeyAscii = 0
            
            If PatiIdentify.GetCurCard.接口序号 > 0 Then
                Call mobjSquareCard.zlGetPatiID(PatiIdentify.GetCurCard.接口序号, PatiIdentify.Text, , lngPatientID)
                
                Call OnFilterRead(PatiIdentify.GetCurCard.名称, PatiIdentify.Text, IIf(lngPatientID > 0, lngPatientID, ""))
            Else
                Call OnFilterRead(PatiIdentify.GetCurCard.名称, PatiIdentify.Text, "")
            End If
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picAppend_Resize()
On Error GoTo ErrHandle
    labHistory.Left = 120
    labHistory.Top = 120
    
    cboTimes.Left = labHistory.Left + labHistory.Width
    cboTimes.Top = 60
    cboTimes.Width = picAppend.Width - labHistory.Width - lblCash.Width - 360
    
    lblCash.Left = cboTimes.Left + cboTimes.Width + 120
    lblCash.Top = 0
    
    labStudyNum.Left = 120
    labStudyNum.Top = cboTimes.Top + cboTimes.Height + 90
    labStudyNum.Width = picAppend.Width - 240
    
    lbl个人信息.Left = 120
    lbl个人信息.Top = labStudyNum.Top + labStudyNum.Height + 30
    
    If picAppend.Width > lbl检查信息.Width + lbl个人信息.Width + 360 Then
        lbl检查信息.Left = lbl个人信息.Left + lbl个人信息.Width + 240
        lbl检查信息.Top = lbl个人信息.Top
    Else
        lbl检查信息.Left = 120
        lbl检查信息.Top = lbl个人信息.Top + lbl个人信息.Height + 60
    End If
    
    txtAppend.Top = lbl检查信息.Top + lbl检查信息.Height + 120
    txtAppend.Left = 60
    txtAppend.Width = picAppend.Width - 70
    txtAppend.Height = picAppend.Height - cboTimes.Height - lbl个人信息.Height - lbl检查信息.Height - 430
    
ErrHandle:
End Sub



Private Sub picLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim lngLeft As Long, lngTop As Long, lngRight  As Long, lngBottom  As Long
 On Error GoTo ErrHandle
    
    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If Button = 1 Then
        
        '当值达到一定范围就退出函数
        If Me.PicLine.Top + Y < lngTop + 700 Or PicLine.Top + Y > PicList.Height - 450 Then
            Exit Sub
        End If

        '移动控件位置
        ufgStudyList.Height = ufgStudyList.Height + Y
        PicLine.Top = PicLine.Top + Y
        picAppend.Top = picAppend.Top + Y
        picAppend.Height = picAppend.Height - Y
        txtAppend.Height = txtAppend.Height - Y
    End If
    
ErrHandle:
End Sub

Private Sub cbrdock_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim strTemp As String
    
    Select Case control.ID
        Case ID_来源
            control.IconId = IIf(Not (mblncmd门诊 Or mblncmd住院 Or mblncmd外诊 Or mblncmd体检), 90000, 90001)
            
            strTemp = IIf(mblncmd门诊, "门诊", "")
            strTemp = strTemp & IIf(mblncmd住院, IIf(strTemp <> "", ",", "") & "住院", "")
            strTemp = strTemp & IIf(mblncmd外诊, IIf(strTemp <> "", ",", "") & "外诊", "")
            strTemp = strTemp & IIf(mblncmd体检, IIf(strTemp <> "", ",", "") & "体检", "")
            
            If strTemp = "" Then
                strTemp = "来源"
                control.ToolTipText = "根据病人来源进行过滤"
            Else
                control.ToolTipText = "显示病人来源为[" & strTemp & "]的检查"
            End If
        
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
        Case ID_门诊
            control.Checked = mblncmd门诊
            control.IconId = IIf(mblncmd门诊, 90001, 90000)
        Case ID_住院
            control.Checked = mblncmd住院
            control.IconId = IIf(mblncmd住院, 90001, 90000)
        Case ID_外诊
            control.Checked = mblncmd外诊
            control.IconId = IIf(mblncmd外诊, 90001, 90000)
        Case ID_体检
            control.Checked = mblncmd体检
            control.IconId = IIf(mblncmd体检, 90001, 90000)
            
            
            
        Case ID_费用
            control.IconId = IIf(Not (mblncmd已缴 Or mblncmd未缴 Or mblncmd补缴 Or mblncmd无费 Or mblncmd记账), 90000, 90001)
            
            strTemp = strTemp & IIf(mblncmd未缴, IIf(strTemp <> "", ",", "") & "未缴", "")
            strTemp = strTemp & IIf(mblncmd已缴, IIf(strTemp <> "", ",", "") & "已缴", "")
            strTemp = strTemp & IIf(mblncmd记账, IIf(strTemp <> "", ",", "") & "记账", "")
            strTemp = strTemp & IIf(mblncmd补缴, IIf(strTemp <> "", ",", "") & "补缴", "")
            strTemp = strTemp & IIf(mblncmd无费, IIf(strTemp <> "", ",", "") & "无费", "")
            
            If strTemp = "" Then
                strTemp = "费用"
                control.ToolTipText = "根据费用状态进行过滤"
            Else
                control.ToolTipText = "显示费用状态为[" & strTemp & "]的检查"
            End If
            
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
        Case ID_已缴
            control.Checked = mblncmd已缴
            control.IconId = IIf(mblncmd已缴, 90001, 90000)
        Case ID_未缴
            control.Checked = mblncmd未缴
            control.IconId = IIf(mblncmd未缴, 90001, 90000)
        Case ID_记账
            control.Checked = mblncmd记账
            control.IconId = IIf(mblncmd记账, 90001, 90000)
        Case ID_补缴
            control.Checked = mblncmd补缴
            control.IconId = IIf(mblncmd补缴, 90001, 90000)
        Case ID_无费
            control.Checked = mblncmd无费
            control.IconId = IIf(mblncmd无费, 90001, 90000)
            
            
        Case ID_影像类别
            control.IconId = IIf(mintcmd影像类别 = 0, 90000, 90001)
        Case ID_影像类别 + 1 To ID_影像类别 + 40
            control.Checked = mblncmd影像类别(control.ID - ID_影像类别 - 1)
            control.IconId = IIf(control.Checked, 90001, 90000)
       
        If control.ID = ID_影像执行间 Then Stop
        Case ID_影像执行间
            control.IconId = IIf(mintcmd影像执行间 = 0, 90000, 90001)
        Case ID_影像执行间 + 1 To ID_影像执行间 + 40
            control.Checked = mblncmd影像执行间(control.ID - ID_影像执行间 - 1)
            control.IconId = IIf(control.Checked, 90001, 90000)

        Case ID_状态
            control.IconId = IIf(Not (mblncmd登记 Or mblncmd报到 Or mblncmd检查 Or mblncmd报告 Or mblncmd审核 Or mblncmd驳回 Or mblncmd完成), 90000, 90001)
            
            strTemp = IIf(mblncmd登记, "登记", "")
            
            strTemp = strTemp & IIf(mblncmd报到, IIf(strTemp <> "", ",", "") & "报到", "")
            strTemp = strTemp & IIf(mblncmd检查, IIf(strTemp <> "", ",", "") & "检查", "")
            strTemp = strTemp & IIf(mblncmd报告, IIf(strTemp <> "", ",", "") & "报告", "")
            strTemp = strTemp & IIf(mblncmd审核, IIf(strTemp <> "", ",", "") & "审核", "")
            strTemp = strTemp & IIf(mblncmd驳回, IIf(strTemp <> "", ",", "") & "驳回", "")
            strTemp = strTemp & IIf(mblncmd完成, IIf(strTemp <> "", ",", "") & "完成", "")
            
            If strTemp = "" Then
                strTemp = "状态"
                control.ToolTipText = "根据检查状态进行过滤"
            Else
                control.ToolTipText = "显示检查状态为[" & strTemp & "]的检查"
            End If
            
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
        
        Case ID_检查部位
            control.IconId = IIf(mstrcmd部位分组 = "", 90000, 90001)
        Case ID_检查部位 + 1 To 4500
            control.Checked = InStr(mstrcmd部位分组, control.Caption) > 0
            control.IconId = IIf(control.Checked, 90001, 90000)
            
        Case ID_登记
            control.Checked = mblncmd登记
            control.IconId = IIf(mblncmd登记, 90001, 90000)
        Case ID_报到
            control.Checked = mblncmd报到
            control.IconId = IIf(mblncmd报到, 90001, 90000)
        Case ID_检查
            control.Checked = mblncmd检查
            control.IconId = IIf(mblncmd检查, 90001, 90000)
        Case ID_报告
            control.Checked = mblncmd报告
            control.IconId = IIf(mblncmd报告, 90001, 90000)
        Case ID_审核
            control.Checked = mblncmd审核
            control.IconId = IIf(mblncmd审核, 90001, 90000)
        Case ID_驳回
            control.Checked = mblncmd驳回
            control.IconId = IIf(mblncmd驳回, 90001, 90000)
        Case ID_完成
            control.Checked = mblncmd完成
            control.IconId = IIf(mblncmd完成, 90001, 90000)
            
        Case ID_病理类别
            control.IconId = IIf(Not (mblncmd常规 Or mblncmd快速石蜡 Or mblncmd冰冻 Or mblncmd细胞 Or mblncmd尸检 Or mblncmd会诊), 90000, 90001)
            
            strTemp = IIf(mblncmd常规, "常规", "")
            
            strTemp = strTemp & IIf(mblncmd冰冻, IIf(strTemp <> "", ",", "") & "冰冻", "")
            strTemp = strTemp & IIf(mblncmd细胞, IIf(strTemp <> "", ",", "") & "细胞", "")
            strTemp = strTemp & IIf(mblncmd尸检, IIf(strTemp <> "", ",", "") & "尸检", "")
            strTemp = strTemp & IIf(mblncmd会诊, IIf(strTemp <> "", ",", "") & "会诊", "")
            strTemp = strTemp & IIf(mblncmd快速石蜡, IIf(strTemp <> "", ",", "") & "快速石蜡", "")
            
            If strTemp = "" Then
                strTemp = "类别"
                control.ToolTipText = "根据病理类别进行过滤"
            Else
                control.ToolTipText = "显示病理类别为[" & strTemp & "]的检查"
            End If
            
            control.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
            
        Case ID_病理类别_常规
            control.Checked = mblncmd常规
            control.IconId = IIf(mblncmd常规, 90001, 90000)
        Case ID_病理类别_快速石蜡
            control.Checked = mblncmd快速石蜡
            control.IconId = IIf(mblncmd快速石蜡, 90001, 90000)
        Case ID_病理类别_冰冻
            control.Checked = mblncmd冰冻
            control.IconId = IIf(mblncmd冰冻, 90001, 90000)
        Case ID_病理类别_细胞
            control.Checked = mblncmd细胞
            control.IconId = IIf(mblncmd细胞, 90001, 90000)
        Case ID_病理类别_尸检
            control.Checked = mblncmd尸检
            control.IconId = IIf(mblncmd尸检, 90001, 90000)
        Case ID_病理类别_会诊
            control.Checked = mblncmd会诊
            control.IconId = IIf(mblncmd会诊, 90001, 90000)
            
        Case ID_本次住院
            control.IconId = IIf(control.Checked, 90001, 90000)
    End Select
    
ErrHandle:
End Sub

Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub

'费用执行
Private Sub ExecuteStudyMoney()
On Error GoTo ErrHandle
    Dim strSql  As String

    If mListAdviceInf.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    strSql = "Zl_影像费用执行(" & mListAdviceInf.lngAdviceId & "," & mListAdviceInf.lngSendNo & ",2,Null,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
    zlDatabase.ExecuteProcedure strSql, "费用执行"
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub conMenu_WorkModule_Click()
On Error GoTo ErrHandle
    Dim frmWorkModule As New frmWorkModuleCfg
    
    frmWorkModule.blnIsUseQueue = mSysPar.blnUseQueue
    Call frmWorkModule.ShowWorkModuleCfg(mlngModule, Me)
    
    '重新配置工作模块页面
    If frmWorkModule.blnIsOk Then
        
        mblnInitOk = False '防止在子窗体加载过程中对子窗体进行刷新
        
        Call InitSubForm
        
        mblnInitOk = True
        
        Call ShowTab
        
        Call picWindow_Resize
        
        '如果没有检查数据，则不允许操作工作模块，只显示模块背景
        If tcDisable.Visible Then Call tcDisable.Translucence
        
        If Not TabWindow.Selected Is Nothing Then Call TabWindow_SelectedChanged(TabWindow.Selected)
        
    End If
    
    Call Unload(frmWorkModule)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal objControl As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim control As XtremeCommandBars.ICommandBarControl
    Dim i As Long
    
    If mblnMenuDownState Then Exit Sub
    
    '这里需要根据id查找对应的菜单项目，因为通过绑定快捷键执行时，产生的是一个只有id而没有其他任何信息的control菜单项
    Set control = cbrMain.FindControl(, objControl.ID, True, True)
    If control Is Nothing Then
        '如果该菜单为电子病历编辑器的右键菜单，则需要修改右键菜单的id等信息
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.ReplacePopupMenu(objControl)
            
            Set control = cbrMain.FindControl(, objControl.ID, True, True)
        End If
        
        If control Is Nothing Then Exit Sub
    End If
    
    If control.ID = 0 Then Exit Sub
    
    mblnMenuDownState = True
        
    cbrMain.RecalcLayout
    
    '执行影像图像对应的功能
    If Not mfrmWork_PacsImg Is Nothing Then
        If mfrmWork_PacsImg.zlMenu.zlIsModuleMenu(control) Then
            Call mfrmWork_PacsImg.zlMenu.zlExecuteMenu(control.ID)
            
            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    If mblnUseActivexCapture Then
        '使用ActivexExc视频方式的图像采集处理
        If Not mobjWork_ActiveVideo Is Nothing Then
'            If mobjWork_ActiveVideo.zlMenu.zlIsModuleMenu(control) Then
'                '执行ActivexExe视频采集对应菜单功能
'                Call mobjWork_ActiveVideo.zlMenu.zlExecuteMenu(control.ID)
'
'                mblnMenuDownState = False
'                Exit Sub
'            End If
        End If
    End If

    
    '执行病理检查对应功能
    If Not mobjWork_Pathol Is Nothing Then
        If mobjWork_Pathol.zlMenu.zlIsModuleMenu(control) Then
            Call mobjWork_Pathol.zlMenu.zlExecuteMenu(control.ID)
            
            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    '执行HIS模块对应功能
    If Not mobjWork_His Is Nothing Then
        If mobjWork_His.zlMenu.zlIsModuleMenu(control) Then
            If mintChangeUserState = 2 Then  '交换了用户，则不允许操作
                MsgBoxD Me, "请统一用户后再操作。"
            Else
                Call mobjWork_His.zlMenu.zlExecuteMenu(control.ID)
                
'                '----------------------补费时，执行费用------------------
'                If control.ID = conMenu_Edit_Append _
'                Or control.ID = conMenu_Edit_Modify _
'                Or control.ID = conMenu_Edit_NewItem * 10# + 1 _
'                Or control.ID = conMenu_Edit_NewItem * 10# + 2 _
'                Or control.ID = conMenu_Edit_NewItem * 10# + 3 Then
'                    If Val(ufgStudyList.CurText("检查状态")) >= 2 Then
'                        Call ExecuteStudyMoney
'                    End If
'                End If
            End If

            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    If Not mobjWork_Report Is Nothing Then
        If mobjWork_Report.zlMenu.zlIsModuleMenu(control) Then
            '执行报告相关功能时，必须先切换到报告模块，否则不允许执行

            If TabWindow.Selected.tag <> "报告填写" Then
                For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                    If TabWindow(i).tag = "报告填写" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
                Next
            End If
            
            If control.Caption <> "批量打印" Then
                If TabWindow.Selected.tag <> "报告填写" Then
                    mblnMenuDownState = False
                    Exit Sub
                End If
            End If
            
            Call mobjWork_Report.zlMenu.zlExecuteMenu(control.ID)
            
            '如果勾选打开报告同时观片参数，则打开观片站
            If (control.ID = conMenu_PacsReport_Open + 1000000 Or control.ID = conMenu_Edit_Modify + 1000000 _
                Or control.ID = conMenu_Edit_Audit + 1000000 Or control.ID = conMenu_File_Open + 1000000) And mSysPar.blnShowImgAfterReport = True Then
                If Not mfrmWork_PacsImg Is Nothing Then
                    Call mfrmWork_PacsImg.zlMenu.zlExecuteMenu(conMenu_Img_Look)
                End If
            End If
            
            mblnMenuDownState = False
            Exit Sub
        End If
    End If
    
    
    Select Case control.ID

'--------------------------文件------------------
        Case conMenu_File_PrintSet '打印设置
            Call zlPrintSet
            
        Case conMenu_File_Excel '清单打印
            Call Menu_File_Excel_click
            
        Case conMenu_File_Parameter '参数设置
            Call Menu_File_Parmeter_click
            
        Case ConMenu_File_ShortcutSet '快捷键设置
            Call Menu_File_ShortcutSet_click
            
        Case conMenu_Pathol_WorkModule  '站点模式设置
            Call conMenu_WorkModule_Click
            
        Case conMenu_Manage_SetXWParam  '设置新网PACS的参数
            Call Menu_Manage_SetXWParam_click
            
        Case conMenu_File_SendImg '发送图像
            Call conMenu_File_SendImg_click
            
        Case conMenu_Cap_DevSet         '视频设置
            If mblnUseActivexCapture Then
                If Not mobjWork_ActiveVideo Is Nothing Then
                    Call mobjWork_ActiveVideo.zlShowVideoConfig
                    mstrCaptureHot = GetSetting("ZLSOFT", "公共模块", "采集热键", "F8")
                    mstrCaptureAfterHot = GetSetting("ZLSOFT", "公共模块", "后台采集热键", "F7")
                    mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "公共模块", "标记更新热键", "F6")
                End If
            Else
                Exit Sub
            End If
            
        Case conMenu_Manage_ChangeUser
            '交换用户时，需要先判断报告是否需要保存
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
        
            Call ChangeUser
            
            '交换用户后，需要刷新报告编辑器，因为用户交换后，原有报告的编辑用户或者创建用户需要进行更新
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
        Case conMenu_Manage_SwitchUser
            '切换用户时，需要先判断报告是否需要保存
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
            Call SwitchUser
            
            '切换用户后，需要刷新报告编辑器，因为用户切换后，原有报告的编辑用户或者创建用户需要进行更新
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlRefreshFace(True, False, False)
            End If
            
        Case conMenu_Manage_Change_In   '隐藏列表
            If dkpMain.Panes(1).Hidden = False Then
                dkpMain.Panes(1).Hide
            Else
                dkpMain.ShowPane (1)
            End If
            
        Case conMenu_File_Exit '退出
            Unload Me
            
'---------------------------检查-----------------
        Case conMenu_Manage_RequestPrint * 10# + 1 To conMenu_Manage_RequestPrint * 10# + 9 '打印诊疗单据
            Call FuncBillPrint(control)
            
        Case comMenu_Petition_Capture                       '扫描申请单
            Call Menu_Petition_扫描申请单(1)
            
        Case comMenu_Petition_View
            Call Menu_Petition_扫描申请单(0)                '查看申请单
            
        Case conMenu_Manage_Regist                          '登记
            Call Menu_Manage_登记
            
        Case conMenu_Manage_CopyCheck                       '复制登记
            Call Menu_Manage_复制登记
            
        Case conMenu_Manage_Receive                         '报到
            Call Menu_Manage_报到
            
        Case conMenu_Manage_Redo                            '取消登记
            Call Menu_Manage_取消登记
            
        Case conMenu_Manage_ReGet                           '召回取消
            Call Menu_Manage_召回取消
            
        Case conMenu_Manage_ThingModi                       '修改登记
            Call Menu_Manage_修改
        
'        Case conMenu_Manage_ModifBaseInfo               '基本信息调整
'            Call Menu_Manage_ModifBaseInfo
        
        Case conMenu_Manage_Logout                          '取消报到
            Call Menu_Manage_取消报到
            
        Case conMenu_Manage_InQueue                         '排队叫号入队
            Call zlInPacsQueue
            
        Case conMenu_Manage_Transfer                        '关联影像
            Call Menu_Manage_关联影像
            
        Case conMenu_Manage_Cancel                          '取消关联
            Call Menu_Manage_取消关联
            
        Case conMenu_Manage_Review                          '随访
            Call Menu_Manage_随访
            
        Case conMenu_Tool_Analyse
            Call OpenViewer(1, mobjPacsCore, mcurAdviceInf.lngAdviceId, True, Me, "", mblnMoved, mSysPar.blnLocalizerBackward)
        
        Case conMenu_Manage_ReportRelease                   '报告发放
            Call Menu_Manage_报告发放
            
        Case conMenu_Manage_FilmRelease                     '胶片发放
            Call Menu_Manage_胶片发放
            
        Case conMenu_Manage_ReportFilmRelease               '报告胶片同时发放
            Call Menu_Manage_报告胶片同时发放
            
        Case conMenu_Manage_SendArrange                     '发送安排
            Call frmSendArrange.ShowMe(Me, mlngCur科室ID, mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo)
            
        Case conMenu_Manage_ReportExecutor                  '报告执行，即标记报告人
            Call Menu_Manage_ReportExecutor
            
        Case conMenu_Manage_CriticalValues, conMenu_Manage_Normal, conMenu_Manage_Critical        '危机值登记
            Call Menu_Manage_CriticalMark(control.ID)
            
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '结果阴阳性
            Call Menu_Manage_标记阴阳(control.ID)
           
        Case conMenu_Manage_FuHe, conMenu_Manage_JiBenFuHe, conMenu_Manage_BuFuHe   '符合情况
            Call Menu_Manage_符合情况(control.ID)
            
        Case conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel
            Call Menu_Manage_绿色通道(control.ID)
            
        Case conMenu_Manage_ClearUp                           '无报告回退
            Call Menu_Edit_无报告回退
                    
        Case conMenu_Manage_Finish                          '无报告直接完成
            Call Menu_Manage_无报告完成
            
        Case conMenu_Manage_Complete                        '检查完成
            Call Menu_Manage_检查最终完成
                
        Case conMenu_Manage_Undone                          '取消检查完成
            Call Menu_Manage_取消检查完成
            
        Case conMenu_Manage_RelatingPatiet                  '关联病人
            Call Menu_Manage_关联病人
            
        Case conMenu_Manage_Burn                            '图像刻录
            Call Menu_Manage_图像刻录
            
        Case conMenu_Manage_CheckList                       '查看电子申请单
            Call Menu_Manage_CheckList
            
'----------------------------------------收藏---------------------------------------
        Case conMenu_Collection_Manage  '收藏管理
           Call Menu_Manage_收藏管理
        Case conMenu_Collection_To      '收藏到
           Call Menu_Manage_收藏到
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999  '动态收藏类别菜单
           Call Menu_Manage_收藏数据显示(control, 0)
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999   '查看共享
           Call Menu_Manage_收藏数据显示(control, 1)
           
           
'----------------------------------------自定义查询---------------------------------------
        Case conMenu_Manage_ConfigQuery '配置查询
            Call ShowCustomQueryConfig
            
        Case conMenu_Manage_CustomQuery * 1000# To conMenu_Manage_CustomQuery * 1000# + 9999
            Call ExecuteCustomQuery(control.ID - conMenu_Manage_CustomQuery * 1000#)   '执行自定义查询
            
        Case conMenu_Manage_CustomQuery '执行综合查询
            mblnIsIntegratedQuery = True
            Call Menu_View_Filter_click
            
        Case conMenu_View_Filter '过滤
                If mlngDefQuerySchemeId >= 0 Then
                    Call ExecuteCustomQuery(mlngDefQuerySchemeId)
                Else
                    mblnIsIntegratedQuery = True
                    Call Menu_View_Filter_click
                End If

'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button '工具栏
            Call Menu_View_ToolBar_Button_click(control)
            
        Case conMenu_View_FontSize_S    '小字体
            Call SetFontSize(0)
        Case conMenu_View_FontSize_M    '中字体
            Call SetFontSize(1)
        Case conMenu_View_FontSize_L    '大字体
            Call SetFontSize(2)
            
            
        Case conMenu_View_ToolBar_Text '按钮文字
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '大图标
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '状态栏
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_Refresh '刷新
            mblnIsCallModuleRefresh = True    '手动刷新时，需要通知所有模块对其进行更新
            Call RefreshList
            
            '调用排队叫号的刷新操作，如果启用
            Call RefreshPacsQueueData
        Case comMenu_Cap_Process
            Call Menu_Manage_浮动采集
            
'--------------------------帮助-----------------
        Case conMenu_Help_Help
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum
            'Case Menu_Help_Web_Forum_click
        Case conMenu_Help_Web_Home
            Call Menu_Help_Web_Home_click
        Case conMenu_Help_Web_Mail
            Call Menu_Help_Web_Mail_click
        Case conMenu_Help_About
            Call Menu_Help_About_click
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse科室, "|")) + 1
            Call Menu_Dept_Select(control)
        Case conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99
            If control.Parameter <> "" Then '执行发布到当前模块的报表
                With ufgStudyList
                    If Val(.CurKeyValue) <> 0 Then
                        Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, _
                            "NO=" & .CurText("NO"), "性质=" & .CurText("记录性质"), "医嘱id=" & Val(.CurKeyValue), 1)
                    Else
                        Call ReportOpen(gcnOracle, Split(control.Parameter, ",")(0), Split(control.Parameter, ",")(1), Me, "", 1)
                    End If
                End With
            End If
        Case Else
            If Val(ufgStudyList.CurKeyValue) = 0 Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            Select Case TabWindow.Selected.tag
                    
                    
                Case "排队叫号"
                    If Not mobjQueue Is Nothing Then
                        If mintChangeUserState = 2 Then  '交换了用户，则不允许操作
                            MsgBoxD Me, "请统一用户后再操作。"
                        Else
                            mobjQueue.zlExecuteCommandbar control
                        End If
                    End If
                Case "申请费用", "住院医嘱", "门诊医嘱", "住院病历", "门诊病历", "门诊电子病历", "住院电子病历"
                    If Not mobjWork_His Is Nothing Then
                        Call mobjWork_His.zlMenu.zlExecuteMenu(control.ID)
                    End If
                Case "报告填写"
                    If Not mobjWork_Report Is Nothing Then
                        Call mobjWork_Report.zlMenu.zlExecuteMenu(control.ID)
                    End If
            End Select
            
    End Select
    
    mblnMenuDownState = False
Exit Sub
ErrHandle:
    mblnMenuDownState = False
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub RefreshPacsQueueData()
'刷新排队模块数据
    If mSysPar.blnUseQueue And Not mobjQueue Is Nothing And mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Call mobjQueue.zlRefreshQueueData(mstrSelQueueRooms)
    End If
End Sub

Private Sub ShowCustomQueryConfig()
'显示自定义查询配置
    Dim frmCusQuery As New frmCustomQueryCfg
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrHandle
    frmCusQuery.Show 1, Me
    
    If frmCusQuery.mblnIsChange Then
        Call RefreshCustomQueryMenu(cbrMain.FindControl(, conMenu_Manage_Query), mlngCur科室ID)
        Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
        
        mlngDefQuerySchemeId = -1
        mlngSysQuerySchemeId = -1
        
        Set rsTemp = zlDatabase.OpenSQLRecord("select id,是否默认,是否系统查询 from 影像查询方案 where (是否默认=1 or 是否系统查询=1) and( 所属科室=0 or 所属科室 is null or 所属科室=[1]) order by 所属科室 desc,方案序号", "获取默认过滤方案", mlngCur科室ID)
        rsTemp.Filter = "是否默认=1"
        If rsTemp.RecordCount > 0 Then mlngDefQuerySchemeId = Val(Nvl(rsTemp!ID))
        rsTemp.Filter = "是否系统查询=1"
        If rsTemp.RecordCount > 0 Then mlngSysQuerySchemeId = Val(Nvl(rsTemp!ID))
    End If
    
ErrHandle:
    Unload frmCusQuery
End Sub

Private Sub ExecuteCustomQuery(ByVal lngSchemeId As Long)
    Dim strReturn As String
    Dim strPars As Variant
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim strWithCustomQueryTab As String   '自定义子查询
    Dim strWithOrderTab As String   '医嘱子查询
    Dim blnEnabledRules As Boolean  '是否启用了规则
    
    '调用自定义查询窗体时，先判断此方案是否包含了录入项
    strSql = "select id from 影像查询配置 where 方案id=[1] and rownum<=1"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询方案配置", lngSchemeId)
    
    If rsData.RecordCount <= 0 Then
        '此方案不包含了录入项
        Call frmCustomQueryCall.GetQuerySqlAndPars(lngSchemeId, strReturn, strPars)
    Else
        strReturn = frmCustomQueryCall.ShowCustomQuery(lngSchemeId, IIf(mblnAllDepts, 0, mlngCur科室ID), mlngModule, strPars, Me)
    End If
    
    If Trim(strReturn) = "" Then Exit Sub
    
    blnEnabledRules = frmCustomQueryCall.mintEnabledRules = 1
    
    strWithCustomQueryTab = " with tmpCustom  as (" & strReturn & ") "
    
    strWithOrderTab = "tmpOrder as ( select Id as ID from tmpCustom Union All select a.ID from 病人医嘱记录 a, tmpCustom b where a.相关ID=b.ID and a.医嘱状态 <> 4 ) "
    
    '执行自定义查询
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        '不能删除该查询中的“影像检查项目”数据表，因为删除后，会造成数据的查询效率较低（删除后，则需要使用病人医嘱发送的执行部门ID作为条件过滤检查，然而该字段没有索引）
        strSql = "Select  Distinct" & vbNewLine & _
                    "       A.医嘱ID,B.相关ID,A.发送号,A.首次时间 报到时间,A.发送时间 申请时间,A.执行状态,nvl(A.执行过程,0) 检查过程,A.执行间,A.结果阳性 阳性,h.危急状态 危急," & vbNewLine & _
                    "       B.病人ID,B.主页ID,B.挂号单,B.病人科室ID,B.病人来源 来源,B.医嘱内容,B.标本部位," & vbNewLine & _
                    "       Nvl(B.紧急标志, 0) 紧急标志, Nvl(B.婴儿, 0) 婴儿,B.开嘱医生,A.NO,C.当前床号,C.当前病区ID,Decode(B.病人来源,2,C.住院号,C.门诊号) 标识号," & vbNewLine & _
                    "       Nvl(B.姓名,H.姓名) 姓名,G.影像类别,H.检查号,Nvl(B.性别,H.性别) 性别,Nvl(B.年龄,H.年龄) 年龄,H.身高,H.体重,H.影像质量,H.报告质量,H.符合情况," & vbNewLine & _
                    "       Decode(B.病人来源,3,B.开嘱医生,A.发送人) 登记人,H.报到人, H.报告发放,H.发放胶片,H.关联ID,A.记录性质, " & vbNewLine & _
                    "       H.完成人,H.是否电子胶片,H.是否打印,H.报告操作,H.绿色通道,H.报告打印,H.报告人,H.复核人,H.是否技师确认,H.检查技师,H.检查技师二,H.接收日期 采图时间," & vbNewLine & _
                    "       H.随访描述,H.诊断分类,H.检查UID,H.图像位置,A.执行部门ID as 执行科室ID,0 as 转出,F.名称 AS 病人科室, a.采样时间, " & vbNewLine & _
                    "       C.就诊卡号,A.NO as 单据号,C.身份证号,C.结算模式,decode(B.病人来源,2,D.病人类型,C.病人类型) as 病人类型,D.路径状态,A.计费状态,Decode(A.记录性质,2,1,Decode(a.计费状态,3,1,0)) as 收费 ,m.医嘱ID as 申请单医嘱 " & vbNewLine & _
                    " From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C,病案主页 D,影像检查记录 H,影像检查项目 G,部门表 F,影像申请单图像 m , tmpOrder " & vbNewLine & _
                    " Where A.医嘱ID=B.ID And A.医嘱ID=H.医嘱ID(+) And A.发送号=H.发送号(+) " & vbNewLine & _
                    "       And B.诊疗项目ID=G.诊疗项目ID And B.病人ID=C.病人ID And B.病人科室id=F.ID" & vbNewLine & _
                    "       And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) and a.医嘱id = m.医嘱id(+) and a.医嘱Id=tmpOrder.ID and b.医嘱状态 <> 4 "

    Else
        '这里单独对病理的查询进行处理，因为病理需要多连接一些查询的数据表
        strSql = "Select Distinct" & vbNewLine & _
                    "       A.医嘱ID,B.相关ID,A.发送号,A.首次时间 报到时间,A.发送时间 申请时间,A.执行状态,nvl(A.执行过程,0) 检查过程,A.结果阳性 阳性,h.危急状态 危急," & vbNewLine & _
                    "       '' as 病理执行状态, o.取材过程,o.制片过程,o.免疫过程,o.分子过程,o.特染过程, " & vbNewLine & _
                    "       decode(o.检查类型,0,'常规',1,'冰冻',2,'细胞',3,'会诊',4,'尸检',5,'快速石蜡',null) as  检查类别, " & vbNewLine & _
                    "       decode(o.病理号,null,'未核收','已核收') as 核收情况, A.执行部门ID as 执行科室ID, " & vbNewLine & _
                    "       B.病人ID,B.主页ID,B.挂号单,B.病人科室ID, B.病人来源 来源,B.医嘱内容,B.标本部位," & vbNewLine & _
                    "       Nvl(B.紧急标志, 0) 紧急标志, Nvl(B.婴儿, 0) 婴儿,B.开嘱医生,A.NO,C.当前床号,C.当前病区ID,Decode(B.病人来源,2,C.住院号,C.门诊号) 标识号," & vbNewLine & _
                    "       Nvl(B.姓名,H.姓名) 姓名,Nvl(B.性别,H.性别) 性别,Nvl(B.年龄,H.年龄) 年龄,H.身高,H.体重,o.综合质量," & vbNewLine & _
                    "       Decode(B.病人来源,3,B.开嘱医生,A.发送人) 登记人,H.报到人,o.病理号,H.报告发放,H.发放胶片,H.关联ID,A.记录性质, " & vbNewLine & _
                    "       H.完成人,H.是否打印,H.报告操作,H.绿色通道,H.报告打印,H.报告质量,H.报告人,H.复核人,H.是否技师确认,H.检查技师,H.接收日期 采图时间, " & vbNewLine & _
                    "       H.随访描述,H.诊断分类,H.检查UID,H.图像位置,0 as 转出,F.名称 AS 病人科室, a.采样时间, t.当前状态 as 会诊状态, t.会诊医师,t.ID as 会诊ID, " & vbNewLine & _
                    "       C.就诊卡号,A.NO as 单据号,C.身份证号,C.结算模式,decode(B.病人来源,2,D.病人类型,C.病人类型) as 病人类型,D.路径状态,A.计费状态,Decode(A.记录性质,2,1,Decode(a.计费状态,3,1,0)) as 收费 ,m.医嘱ID as 申请单医嘱, " & vbNewLine & _
                    "      (select count(1) from 病理检查信息 V , 病理申请信息 W where V.病理医嘱ID=w.病理医嘱id and v.医嘱id=A.医嘱ID and w.补费状态=1) as 补费 " & vbNewLine & _
                    " From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C,病案主页 D,影像检查记录 H,部门表 F, " & vbNewLine & _
                    "       病理检查信息 o ,影像申请单图像 m,  病理会诊信息 t, tmpOrder " & _
                    " Where A.医嘱ID=B.ID And A.医嘱ID=H.医嘱ID(+) And A.发送号=H.发送号(+) " & vbNewLine & _
                    "       And B.病人ID=C.病人ID And B.病人科室id=F.ID " & vbNewLine & _
                    "       and A.医嘱ID=o.医嘱ID(+) " & vbNewLine & _
                    "       And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) and a.医嘱id = m.医嘱id(+) and o.病理医嘱ID=t.病理医嘱ID(+) and a.医嘱id= tmpOrder.ID and b.医嘱状态 <> 4 "
    End If
    
    strSql = "select " & IIf(blnEnabledRules, "/*+ RULE*/", "") & " * from (" & strWithCustomQueryTab & "," & strWithOrderTab & strSql & ")"
    
    Set ufgStudyList.AdoData = GetDataToLocal(strSql, "自定义查询", strPars(1), strPars(2), strPars(3), strPars(4), strPars(5), strPars(6), strPars(7), strPars(8), strPars(9), strPars(10), _
                                            strPars(11), strPars(12), strPars(13), strPars(14), strPars(15), strPars(16), strPars(17), strPars(18), strPars(19), strPars(20))
    
    ufgStudyList.AdoFilter = GetFilterWhere
    
    '用binddata的方法比使用refreshdata的方法快
    Call ufgStudyList.BindData
    
    '恢复排序
    Call ufgStudyList.ResetSort(mlngSortCol, mintSortOrder)
    
    Call RefreshStatusBarInf
 
    If ufgStudyList.GridRows > 1 Then
        Call ufgStudyList.LocateRow(1)
        Call ufgStudyList_OnSelChange
    End If
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
    '设置字体大小
    mbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, IIf(bytSize = 2, 15, bytSize)))
    
    Call ReMoveCtrl(mbytFontSize)
    Call ReSetFormFontSize
    Call ReSetModuleFontSize(mbytFontSize, IIf(bytSize = 2, 1, bytSize))
    Call SetSelectRowColor
End Sub


Private Sub ReSetModuleFontSize(ByVal bytFontSize As Byte, ByVal bytSize As Byte)
'功能:重新设置各个业务模块窗体的字体大小
    On Error Resume Next
    
    '判断 当前选中的
    Select Case mlngModule
        Case 1290
            If Not mfrmWork_PacsImg Is Nothing Then
                If TabWindow.Selected.tag = "影像图象" Then
                    Call mfrmWork_PacsImg.ReSetFormFontSize(mbytFontSize)
                End If
            End If
            
            If Not mobjWork_His Is Nothing Then
                If Not mobjWork_His.GetExpenseObj Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(bytSize)
                If Not mobjWork_His.GetAdviceObj Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
                If Not mobjWork_His.GetEPRsObj Is Nothing Then Call mobjWork_His.GetEPRsObj.SetFontSize(bytSize)
            End If
            
        Case 1291
            If Not mobjWork_His Is Nothing Then
               If Not mobjWork_His.GetExpenseObj Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(bytSize)
               If Not mobjWork_His.GetAdviceObj Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
               If Not mobjWork_His.GetEPRsObj Is Nothing Then Call mobjWork_His.GetEPRsObj.SetFontSize(bytSize)
            End If
            
        Case 1294
        
            If Not mobjWork_Pathol Is Nothing Then
                Select Case TabWindow.Selected.tag
                    Case "标本核收"
                        Call mobjWork_Pathol.GetModule(mtSpecimen).ReSetFormFontSize(mbytFontSize)
                        
                    Case "病理取材"
                        Call mobjWork_Pathol.GetModule(mtMaterial).ReSetFormFontSize(mbytFontSize)
                        
                    Case "病理制片"
                        Call mobjWork_Pathol.GetModule(mtSlices).ReSetFormFontSize(mbytFontSize)
                        
                        
                    Case "病理特检"
                        Call mobjWork_Pathol.GetModule(mtSpeExam).ReSetFormFontSize(mbytFontSize)
                        
                    Case "过程报告"
                        Call mobjWork_Pathol.GetModule(mtProRep).ReSetFormFontSize(mbytFontSize)
                        
                    Case "申请费用"
                        If Not mobjWork_His Is Nothing Then Call mobjWork_His.GetExpenseObj.SetFontSize(mbytFontSize, bytSize)
                        
                    Case "门诊医嘱", "住院医嘱"
                        If Not mobjWork_His Is Nothing Then Call mobjWork_His.GetAdviceObj.SetFontSize(bytSize)
                    
                End Select
            End If
    End Select
End Sub

Private Sub ReSetFormFontSize()
'功能:重新设置工作站窗体的字体大小
    On Error Resume Next
    
    Dim objCtrl As control
    Dim CtlFont As StdFont
    Dim strFontType As String
    
    Me.FontSize = mbytFontSize
    Set CtlFont = New StdFont
    strFontType = IIf(IsUseClearType = True, "微软雅黑", "宋体")
    CtlFont.Name = strFontType
    
    For Each objCtrl In Me.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("TabStrip") '页面控件
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = mbytFontSize
        Case UCase("Label")
            If objCtrl.Name <> "lblCash" Then
                objCtrl.Font.Name = strFontType
                objCtrl.FontSize = mbytFontSize
                objCtrl.Height = TextHeight("罗") + 60
            End If
        Case UCase("vsFlexGrid")
        
            CtlFont.Name = strFontType
            CtlFont.Size = mbytFontSize
            objCtrl.DataGrid.Font = CtlFont
            
        Case UCase("ucFlexGrid")
            ufgStudyList.HeadFont.Size = mbytFontSize
            objCtrl.DataGrid.Cell(flexcpFontSize, 0, 0, 0, objCtrl.DataGrid.Cols - 1) = mbytFontSize
            objCtrl.DataGrid.FontName = strFontType
            objCtrl.DataGrid.FontSize = mbytFontSize
            objCtrl.DataGrid.RowHeight(0) = TextHeight("罗") + 150
        Case UCase("ComboBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
        Case UCase("OptionButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("CheckBox")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
            objCtrl.Width = TextWidth("罗冠" & objCtrl.Caption)
        Case UCase("DTPicker")
            objCtrl.Font.Name = strFontType
            objCtrl.Font.Size = mbytFontSize
            objCtrl.Width = TextWidth("2012-01-01 23:59:59") * 1.25
            objCtrl.Height = TextHeight("罗") * 1.5
        Case UCase("textBox")
          objCtrl.FontName = strFontType
          objCtrl.FontSize = mbytFontSize
        Case UCase("ReportControl")
            
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            
            CtlFont.Size = mbytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            
            CtlFont.Size = mbytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontName = strFontType
            objCtrl.FontSize = mbytFontSize
        End Select
    Next
    
    Call picAppend_Resize
    
End Sub
Private Sub ReMoveCtrl(ByVal bytFontSize As Byte)
'功能:移动控件位置
    On Error GoTo ErrHandle
    Dim lngMove As Long '控件移动距离
    
    mbytFontSize = bytFontSize
    lngMove = IIf(bytFontSize = 9, 1200, IIf(bytFontSize = 12, 1500, 2000))
    
    
    If glngModul = 1294 Then
        optAccept.Left = optNeed.Left + lngMove
        optFinal.Left = optAccept.Left + lngMove
        optAll.Left = optFinal.Left + lngMove
        tabFilter.Width = IIf(bytFontSize = 15, 4800, 4000)
        tabFilter.Height = IIf(bytFontSize = 9, 375, IIf(bytFontSize = 12, 400, 425))
    End If
    
    '调用病人详细信息 界面重置方法
    Call picAppend_Resize
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub

Private Sub Menu_View_Filter_click()
    On Error GoTo ErrHandle

    If mfrmPACSFilter Is Nothing Then Set mfrmPACSFilter = New frmPACSFilter
    
    With mfrmPACSFilter
        .mlngModul = mlngModule
        .mBeforeDays = mSysPar.lngBeforeDays - 1
        .mDept = mlngCur科室ID '当前科室
        .Show 1, Me
        If Not .mblnOK Then Exit Sub '没有返回条件
        
        '当使用时间条件时，清空固定条件
        PatiIdentify.Text = ""
        SQLCondition.姓名 = ""
        SQLCondition.就诊卡 = ""
        SQLCondition.门诊号 = 0
        SQLCondition.住院号 = 0
        SQLCondition.健康号 = ""
        SQLCondition.单据号 = ""
        SQLCondition.检查号 = 0
        SQLCondition.身份证 = ""
        SQLCondition.IC卡 = ""
        SQLCondition.结果阳性 = -1
        
        
        SQLCondition.开始时间 = Format(.dtpBegin.value, "yyyy-MM-dd HH:mm:00")
        SQLCondition.结束时间 = Format(.dtpEnd.value, "yyyy-MM-dd HH:mm:59")
        
        mblnMoved = MovedByDate(SQLCondition.开始时间)
        
        If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
            gblnXWMoved = mblnMoved
        End If
        
        If .optFindType(1).value = True Then '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）、4=安排时间（病人医嘱记录.开始执行时间）
            SQLCondition.时间类型 = 1
        ElseIf .optFindType(2).value = True Then
            SQLCondition.时间类型 = 2
        ElseIf .optFindType(3).value = True Then
            SQLCondition.时间类型 = 3
        Else
            SQLCondition.时间类型 = 4
        End If
        
        If NeedName(.cboPart.Text) <> "所有部位" Then '检查标本部位
            SQLCondition.标本部位 = NeedName(.cboPart.Text)
        Else
            SQLCondition.标本部位 = ""
        End If
        
        '病人性别
        If NeedName(.CboSex.Text) = "全部" Then
            SQLCondition.性别 = ""
        Else
            SQLCondition.性别 = NeedName(.CboSex.Text)
        End If
        
        '病人年龄
        Select Case NeedName(.cboAgeType.Text)
            Case "岁"
                SQLCondition.开始年龄 = Val(.txtBeginAge.Text) * 365
                SQLCondition.结束年龄 = Val(.txtEndAge.Text) * 365
            Case "月"
                SQLCondition.开始年龄 = Val(.txtBeginAge.Text) * 30
                SQLCondition.结束年龄 = Val(.txtEndAge.Text) * 30
            Case "周"
                SQLCondition.开始年龄 = Val(.txtBeginAge.Text) * 7
                SQLCondition.结束年龄 = Val(.txtEndAge.Text) * 7
            Case "天"
                SQLCondition.开始年龄 = Val(.txtBeginAge.Text) * 1
                SQLCondition.结束年龄 = Val(.txtEndAge.Text) * 1
        End Select
        
        If Trim(.txtBeginAge.Text) = "" Then SQLCondition.开始年龄 = -1
        If Trim(.txtEndAge.Text) = "" Then SQLCondition.结束年龄 = -1
        
        SQLCondition.年龄条件 = Trim(.cboAgeWhere.Text)
        
        If NeedName(.cboDept.Text) <> "所有科室" Then '病人科室
            SQLCondition.病人科室 = .cboDept.ItemData(.cboDept.ListIndex)
        Else
            SQLCondition.病人科室 = 0
        End If

        If NeedName(.cboDiagDOC.Text) <> "所有医生" Then '诊断医生
            SQLCondition.诊断医生 = NeedName(.cboDiagDOC.Text)
        Else
            SQLCondition.诊断医生 = ""
        End If
        
        If NeedName(.cboAuditing.Text) <> "所有医生" Then '审核医生
            SQLCondition.审核医生 = NeedName(.cboAuditing.Text)
        Else
            SQLCondition.审核医生 = ""
        End If
       
      
        If .cboModality.Text <> "所有类别" Then '影像类别
            SQLCondition.影像类别 = Split(.cboModality.Text, "-")(1)
        Else
            SQLCondition.影像类别 = ""
        End If
        
        If Trim(.Txt影像诊断) <> "" Then '影像诊断
            SQLCondition.疾病诊断 = Trim(.Txt影像诊断)
        Else
            SQLCondition.疾病诊断 = ""
        End If
        
        If Trim(.txt报告内容) <> "" Then '报告内容
            SQLCondition.报告内容 = Trim(.txt报告内容)
        Else
            SQLCondition.报告内容 = ""
        End If
        
        If NeedName(.cboYinYangXing.Text) = "阳性" Then
            SQLCondition.结果阳性 = 1
        ElseIf NeedName(.cboYinYangXing.Text) = "阴性" Then
            SQLCondition.结果阳性 = 0
        Else
            SQLCondition.结果阳性 = -1
        End If
        
        If .cbo质量.ListIndex = 0 Then
            SQLCondition.影像质量 = ""
        Else
            SQLCondition.影像质量 = NeedName(.cbo质量.Text)
        End If
        
        If NeedName(.cbo检查技师.Text) = "所有医生" Then
            SQLCondition.检查技师 = ""
        Else
            SQLCondition.检查技师 = NeedName(.cbo检查技师.Text)
        End If
        
        If Trim(.txtPacsRpt(0)) <> "" Then 'PACS报告检索
            SQLCondition.检查所见 = Trim(.txtPacsRpt(0))
        Else
            SQLCondition.检查所见 = ""
        End If
        
        If Trim(.txtPacsRpt(1)) <> "" Then
            SQLCondition.诊断意见 = Trim(.txtPacsRpt(1))
        Else
            SQLCondition.诊断意见 = ""
        End If
        
        If Trim(.txtPacsRpt(2)) <> "" Then
            SQLCondition.建议 = Trim(.txtPacsRpt(2))
        Else
            SQLCondition.建议 = ""
        End If
        
        If Trim(.txt随访.Text) <> "" Then
            SQLCondition.随访 = Trim(.txt随访.Text)
        Else
            SQLCondition.随访 = ""
        End If
        
        Call RefreshList '调用刷新
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
On Error GoTo ErrHandle
    Dim objControl As CommandBarControl, i As Integer
    Dim aryKindInfo() As String
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    If mlngModule = G_LNG_PACSSTATION_MODULE Then
                        '只有医技需要添加“全部科室”的科室选择菜单
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100#, "全部科室")
                    
                        objControl.Category = "Main"
                        objControl.DescriptionText = 0
                        If mblnAllDepts = True Then objControl.Checked = True
                    End If
                    
                    '再添加每一个具体科室
                    For i = 0 To UBound(Split(mstrCanUse科室, "|"))  'mstrCanUse科室=id_编码-名称|id_编码-名称
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i + 1, Split(Split(mstrCanUse科室, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse科室, "|")(i), "_")(0)
                        
                        If mblnAllDepts = False And mlngCur科室ID = objControl.DescriptionText Then
                            objControl.Checked = True
                        End If
                    Next
                End If
            End With
        Case Else
            Select Case Me.TabWindow.Selected.tag
                Case "住院医嘱", "门诊医嘱", "申请费用"
                    Call mobjWork_His.zlMenu.zlRefreshSubMenu(CommandBar)
            End Select
    End Select
ErrHandle:
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error GoTo ErrHandle
    Dim blnNoRecord As Boolean
    Dim intState As Integer
    Dim blnCancel As Boolean
    Dim tt As CommandBarControl
    Dim objControl As XtremeCommandBars.ICommandBarControl
    
    If Not mblnInitOk Then Exit Sub
    
    '如果该菜单为电子病历编辑器的右键菜单，则需要修改菜单id等信息
    Set objControl = cbrMain.FindControl(, control.ID, True, True)
    If objControl Is Nothing Then
        If Not mobjWork_Report Is Nothing Then
            Call mobjWork_Report.ReplacePopupMenu(control)
        End If
    End If
    
    If ufgStudyList.GridCols <= 1 Or ufgStudyList.GridRows <= 1 Or Not ufgStudyList.IsSelectionRow Then
        blnNoRecord = True
    Else
        blnNoRecord = Val(ufgStudyList.CurKeyValue) = 0
    End If
    
    If Not blnNoRecord Then
        intState = Val(ufgStudyList.CurText("检查状态"))
        blnCancel = ufgStudyList.CurText("检查过程") = "已拒绝"
    End If
    
    If TabWindow.ItemCount > 0 Then
        If TabWindow.Selected Is Nothing Then Exit Sub
        
        '更新影像图像菜单
        If Not mfrmWork_PacsImg Is Nothing Then
            If mfrmWork_PacsImg.zlMenu.zlIsModuleMenu(control) Then
                Call mfrmWork_PacsImg.zlMenu.zlUpdateMenu(control)
                Exit Sub
            End If
        End If
        
        '更新病理检查菜单
        If Not mobjWork_Pathol Is Nothing Then
            If mobjWork_Pathol.zlMenu.zlIsModuleMenu(control) Then

                Select Case control.ID
                    Case conMenu_PatholSpecimen
                        control.Visible = IIf(TabWindow.Selected.tag = "标本核收", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholMaterial
                        control.Visible = IIf(TabWindow.Selected.tag = "病理取材", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholSlices
                        control.Visible = IIf(TabWindow.Selected.tag = "病理制片", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholSpeExam
                        control.Visible = IIf(TabWindow.Selected.tag = "病理特检", True, False)
                        
                        Exit Sub
                    Case conMenu_PatholProRep
                        control.Visible = IIf(TabWindow.Selected.tag = "过程报告", True, False)
                        
                        Exit Sub
                End Select
                
                Call mobjWork_Pathol.zlMenu.zlUpdateMenu(control)
                
                Exit Sub
            End If
        End If
        
        '更新HIS模块菜单
        If Not mobjWork_His Is Nothing Then
            
            If InStr("申请费用, 住院医嘱, 门诊医嘱, 住院病历, 门诊病历, 门诊电子病历, 住院电子病历", TabWindow.Selected.tag) > 0 Then
                If mobjWork_His.zlMenu.zlIsModuleMenu(control) Then
                    Call mobjWork_His.zlMenu.zlUpdateMenu(control)
                    
                    '已完成除查阅,以及医嘱中报告查看打印，观片菜单外均不可用
                    If Val(ufgStudyList.CurText("检查状态")) = 6 Then
                        Select Case control.ID
                            Case conMenu_Edit_MarkMap, conMenu_Tool_PlugIn, conMenu_Edit_Compend, conMenu_Manage_ReportLisView, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                                control.Enabled = True
                            Case conMenu_Edit_Copy, conMenu_File_ExportToXML, conMenu_Tool_Search, conMenu_File_Open, conMenu_EditPopup, conMenu_Edit_ChargeDelAudit
                                '这几个菜单不控制
                            Case Else
                                control.Enabled = False
                        End Select
                    End If
                    
                    Exit Sub
                End If
            End If
        End If
        
        If mblnUseActivexCapture Then
            '更新使用ActivexExe方式的视频采集菜单
            If Not mobjWork_ActiveVideo Is Nothing Then
'                If mobjWork_ActiveVideo.zlMenu.zlIsModuleMenu(control) Then
'                    '更新视频采集菜单...
'                    Call mobjWork_ActiveVideo.zlMenu.zlUpdateMenu(control)
'                    Exit Sub
'                End If
            End If
        End If

        
        '更新报告模块菜单
        If Not mobjWork_Report Is Nothing Then
            If mobjWork_Report.zlMenu.zlIsModuleMenu(control) Then
                Call mobjWork_Report.zlMenu.zlUpdateMenu(control)
                
                '当前查看的是历次记录则菜单均不可用
                If cboTimes.ListIndex <> -1 Then
                    If mListAdviceInf.lngAdviceId <> cboTimes.ItemData(cboTimes.ListIndex) Then
                        If control.ID = conMenu_Edit_Copy + 1000000 Or control.ID = conMenu_File_ExportToXML + 1000000 Or control.ID = conMenu_EditPopup + 1000000 _
                            Or control.ID = conMenu_Tool_Search + 1000000 Or control.ID = conMenu_File_Preview + 1000000 Or control.ID = conMenu_File_Print + 1000000 Or control.ID = conMenu_File_NoAskPrint + 1000000 Then
                            '这几个菜单不控制
                        Else
                            control.Enabled = False
                        End If
                    End If
                End If
            
                Exit Sub
            End If
        End If
    End If
    
    
    Select Case control.ID
        Case conMenu_Manage_LocateValue
            control.Enabled = Not blnNoRecord
        Case comMenu_Cap_Process
            control.Enabled = True 'Not blnNoRecord
        Case conMenu_View_Filter * 10#
            control.Caption = "当前科室:" & IIf(mblnAllDepts = True, "全部科室", mstrCur科室)
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse科室, "|")) + 1
            If mblnAllDepts = True Then
                control.Checked = (control.DescriptionText = 0)
            Else
                control.Checked = (control.DescriptionText = mlngCur科室ID)
            End If
        Case conMenu_View_ToolBar_Button '工具栏
            If cbrMain.Count >= 2 Then
                control.Checked = Me.cbrMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbrMain.Count >= 2 Then
                control.Checked = Not (Me.cbrMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            control.Checked = Me.cbrMain.Options.LargeIcons
        Case conMenu_View_StatusBar '状态栏
            control.Checked = Me.stbThis.Visible
        Case conMenu_View_Filter   '过滤
        
        Case conMenu_View_Refresh  '刷新
        
        Case conMenu_Manage_RequestPrint
            control.Enabled = control.CommandBar.Controls.Count > 0 And Not blnNoRecord
            
        Case conMenu_Manage_Regist   '检查登记(&I)
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            End If
        Case conMenu_Manage_CopyCheck '复制登记
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Redo   '取消登记(&R)
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And intState <> -1 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ReGet   '召回取消
            If Not blnNoRecord Then
                control.Enabled = blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ThingModi   '修改信息(&M)
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState < 6 And Not blnCancel
            Else
                control.Enabled = False
            End If
'        Case conMenu_Manage_ModifBaseInfo '基本信息调整
'            If InStr(mstrPrivs, "检查登记") <= 0 Then
'                control.Visible = False
'            ElseIf Not blnNoRecord Then
'                control.Enabled = intState < 6 And Not blnCancel
'            Else
'                control.Enabled = False
'            End If
        Case conMenu_Manage_Receive   '检查报到(&L)
            If InStr(mstrPrivs, "检查报到") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And intState <> -1 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Logout   '取消报到(&D)
            If blnNoRecord Then
                control.Enabled = False
            ElseIf control.Parent Is Nothing Then '当使用热键时，如果不判断parent，将会产生异常
                Exit Sub
            ElseIf control.Parent.type = xtpControlPopup Then
                If InStr(mstrPrivs, "取消报到") <= 0 Then
                    control.Visible = False
                Else
                    control.Visible = True
                    control.ToolTipText = "取消报到"
                    control.Caption = "取消报到(&D)"
                    control.Enabled = (intState = 2 Or intState = 3)
                End If
            Else ' 工具栏中的用取消检查代替取消登记,同一按键完成取消登记和取消检查功能
                control.Visible = IIf(intState <= 1 And intState <> -1, InStr(mstrPrivs, "检查登记") > 0, InStr(mstrPrivs, "取消报到") > 0)
                control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And intState <> -1 And Not blnCancel) '被拒绝的不能被再次拒绝
                control.ToolTipText = IIf(intState <= 1 And intState <> -1, "取消登记", "取消报到")
                control.Caption = "取消"
            End If
        Case conMenu_Manage_InQueue    '排队叫号入队
            control.Visible = mSysPar.blnUseQueue And Not mSysPar.blnAutoInQueue
            control.Enabled = (intState >= 2 And intState <= 5)
            
        Case conMenu_Manage_Transfer   '关联影像(&C)
            If InStr(mstrPrivs, "图像关联") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用
            End If
        Case conMenu_Manage_Cancel   '取消关联(&B)
            If InStr(mstrPrivs, "图像关联") <= 0 Then
                control.Visible = False
            ElseIf (intState >= 2 And intState <= 5) Or intState = -1 Then
                control.Enabled = ufgStudyList.CurText("检查UID") <> ""
            Else
                control.Enabled = False
            End If
            
        Case conMenu_Manage_Review  '随访
            If InStr(mstrPrivs, "随访") <= 0 Then
                control.Visible = False
            ElseIf (Not blnNoRecord And intState > 1 And intState <= 6) Or intState = -1 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Tool_Analyse   '高级图像处理
            If InStr(GetPrivFunc(glngSys, 1289), "基本") <= 0 Then
                control.Visible = False
            ElseIf (Not blnNoRecord And intState > 1 And intState < 6) Or intState = -1 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Release, conMenu_Manage_ReportFilmRelease     '报告发放,报到后，完成后都可以执行

            control.Enabled = IIf(intState >= 2 And intState < 6, True, False)
            
            If mrtReportType = 报告文档编辑器 Then
                If control.ID = conMenu_Manage_ReportFilmRelease Then
                    If mobjWork_Report.GetReportReleaseState(mcurAdviceInf.lngAdviceId) = 3 And Nvl(ufgStudyList.CurText("胶片发放"), "0") = "1" Then
                        control.Caption = "收回所有"
                        control.ToolTipText = "收回已经发放的报告或胶片"
                    Else
                        control.Caption = "发放所有"
                        control.ToolTipText = "报告和胶片同时发放"



                    End If
                End If
            Else
                If Not blnNoRecord Then
                  '修改报告发放按钮的标题
                     If Not blnNoRecord Then
                         If Nvl(ufgStudyList.CurText("报告发放"), "0") = "1" And Nvl(ufgStudyList.CurText("胶片发放"), "0") = "1" Then
                             control.Caption = "收回"
                             control.ToolTipText = "收回已经发放的报告或胶片"
                         Else
                             control.Caption = "发放"
                             control.ToolTipText = IIf(control.ID = conMenu_Manage_Release, "报告或胶片发放", "报告和胶片同时发放")
                         End If
                     End If
                End If
            End If
            
            control.Enabled = Not control.Enabled
            control.Enabled = Not control.Enabled
        Case conMenu_Manage_FilmRelease
            control.Enabled = IIf(intState >= 2 And intState < 6, True, False)
            
            If Not blnNoRecord Then
                 If Nvl(ufgStudyList.CurText("胶片发放"), "0") = "1" Then
                     control.Caption = "胶片收回"
                     control.ToolTipText = "收回已经发放的胶片"
                     
                     If InStr(mstrPrivs, "取消发放") > 0 Then
                         control.Enabled = True
                     Else
                         control.Enabled = False
                     End If
                 Else
                     control.Caption = "胶片发放"
                     control.ToolTipText = "胶片发放"
                 End If
             End If


        Case conMenu_Manage_ReportRelease
            control.Enabled = IIf(intState >= 2 And intState <= 6, True, False)
            
            If mlngModule = G_LNG_PATHOLSYS_NUM Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
                If Not blnNoRecord Then
                    If mrtReportType = 报告文档编辑器 Then
                        If mobjWork_Report.GetReportReleaseState(mcurAdviceInf.lngAdviceId) > 1 Then
                            control.Caption = "报告收回"
                            control.ToolTipText = "收回已经发放的报告"
                        Else
                            control.Caption = "报告发放"
                            control.ToolTipText = "报告发放"
                        End If
                    Else
                        If ufgStudyList.CurText("报告发放") = "1" Then
                            control.Caption = "报告收回"
                            control.ToolTipText = "收回已经发放的报告"
                        Else
                            control.Caption = "报告发放"
                            control.ToolTipText = "报告发放"
                        End If
                    End If
                End If
            End If
            control.Enabled = Not control.Enabled
            control.Enabled = Not control.Enabled
        
        Case conMenu_Manage_SendArrange                     '发送安排
            control.Enabled = IIf(intState >= 2 And intState < 6, True, False)
            
        Case conMenu_Manage_ReportExecutor      '报告执行
            control.Enabled = IIf(intState >= 2 And intState <= 6, True, False)
        
        Case conMenu_Manage_CriticalValues, conMenu_Manage_CriticalSituation, conMenu_Manage_Normal, conMenu_Manage_Critical '危急值
            If mSysPar.lngCriticalValues = 0 Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用
            End If

        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive   '结果阴阳性(&X)
            If mSysPar.blnIgnoreResult = True Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用
                If ufgStudyList.CurText("危急") = " " And control.ID = conMenu_Manage_Result Then control.Enabled = False
            End If
            
        Case conMenu_Manage_FuHe, conMenu_Manage_JiBenFuHe, conMenu_Manage_BuFuHe, conMenu_Manage_FuHeLevel '符合情况
            If mSysPar.lngConformDetermine = 0 Then
                control.Visible = False
            Else
                control.Visible = True
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用
            End If
        
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '绿色通道标记/取消
            If InStr(mstrPrivs, "绿色通道") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState >= 2 And intState <= 5) Or intState = -1 '在2---5之间可用
            End If
        Case conMenu_Manage_Finish   '无报告完成(&F)
            If InStr(mstrPrivs, "无报告完成") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 2 Or intState = 3
            End If
        Case conMenu_Manage_ClearUp   '无报告回退(&U)
            If InStr(mstrPrivs, "无报告完成") <= 0 Then
                control.Visible = False
            ElseIf intState = 5 Then
                control.Enabled = IIf(mrtReportType = 报告文档编辑器, mobjWork_Report.GetCurrDocId(mcurAdviceInf.lngAdviceId) = "", ufgStudyList.CurText("报告人") = "")
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Complete   '检查完成(&E)
            If InStr(mstrPrivs, "检查完成") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState < 6
            End If
        Case conMenu_Manage_Undone   '取消完成(&U)
            If InStr(mstrPrivs, "取消检查完成") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState = 6
            End If
        Case conMenu_File_SendImg  '发送图像
            If InStr(mstrPrivs, "文件发送") <= 0 Then control.Visible = False
        Case conMenu_Img_Contrast, conMenu_Img_Look     '影像对比,影像观片
            If mblnObserve Then
                If blnNoRecord Then control.Enabled = False: Exit Sub

                control.Enabled = mcurAdviceInf.strStudyUID <> ""
            Else
                control.Visible = False
            End If
        Case conMenu_Manage_RelatingPatiet  '关联病人
            If InStr(mstrPrivs, "关联病人") <= 0 Or mSysPar.blnRelatingPatient = False Then
                control.Visible = False
            ElseIf blnNoRecord Or (intState < 2 And intState <> -1) Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
        Case conMenu_Manage_Burn
            control.Visible = IIf(InStr(mstrPrivs, "图像刻录") <= 0, False, True)
        
        Case conMenu_Manage_CheckList   '查看申请单
            control.Visible = True
            If mListAdviceInf.lngAdviceId > 0 And mListAdviceInf.lngPatientFrom <> 3 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_File_SendImg
            If InStr(mstrPrivs, "文件发送") <= 0 Then control.Visible = False
        Case conMenu_File_PrintSet     '打印设置(&S)
        Case conMenu_File_Excel         '清单打印(&L)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_Parameter, conMenu_Cap_DevSet
        
        Case conMenu_Manage_ChangeUser  '用户交换
            If mSysPar.blnChangeUser Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            
        Case conMenu_Manage_SwitchUser  '切换用户
            If mSysPar.blnSwitchUser Then
                control.Visible = True
            Else
                control.Visible = False
            End If
        
        Case conMenu_Manage_SetXWParam      '新网PACS参数设置，如果有此菜单，就显示
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '报表
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_Help_Help, conMenu_Help_About  '帮助
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '帮助WEB
        Case conMenu_File_Exit
        Case ConMenu_File_ShortcutSet
        Case conMenu_Pathol_WorkModule
        Case conMenu_View_ToolBar
        Case conMenu_Manage_Query
        Case conMenu_Manage_CustomQuery * 1000# To conMenu_Manage_CustomQuery * 1000# + 999
        Case conMenu_Manage_CustomQuery
        Case conMenu_Manage_ConfigQuery '查询配置
            control.Visible = CheckPopedom(mstrPrivs, "查询配置")
        Case conMenu_Cap_DevSet     '影像设备设置
        Case conMenu_Manage_Change_In   '隐藏列表
        Case conMenu_Img_3D_MMPR, conMenu_Img_3D_MPR, conMenu_Img_3D_PF, conMenu_Img_3D_SA, conMenu_Img_3D_VA, conMenu_Img_3D_VE '三维重建的几个子菜单不需要设置
        Case conMenu_View_FontSize_S    '小字体
             control.Checked = mbytFontSize = 9
        Case conMenu_View_FontSize_M    '中字体
             control.Checked = mbytFontSize = 12
        Case conMenu_View_FontSize_L    '大字体
             control.Checked = mbytFontSize = 15
        
   '-------------------------------------------------收藏管理部分----------------------------------------------------------
 
        Case conMenu_Collection    '收藏(&C)
            control.Enabled = True
        Case conMenu_Collection_Manage  '收藏管理菜单
            control.Enabled = True
        Case conMenu_Collection_ViewShare      '查看共享
            control.Enabled = True
        Case comMenu_Collection_Type * 10000# To comMenu_Collection_Type * 10000# + 9999  '动态收藏菜单
            control.Enabled = True
        Case conMenu_Collection_ViewShare * 10000# To conMenu_Collection_ViewShare * 10000# + 9999  '动态共享菜单
            control.Enabled = True
         Case conMenu_Collection_To
            
            
    '-------------------------------------------扫描申请单部分-----------------------------------------------

        '扫描申请单
        Case comMenu_Petition_Capture
            If Val(ufgStudyList.CurKeyValue) = 0 Or blnCancel Then
                control.Enabled = False
            Else
                control.Enabled = IIf((intState >= 2 And intState <= 5) Or intState = -1, True, False)
            End If
            
        '查看申请单
        Case comMenu_Petition_View
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Enabled = False
            End If
            
        Case Else
            If blnNoRecord Then
                control.Enabled = False
                Exit Sub
            End If
                    
            
            '已完成除查阅,以及医嘱中报告查看打印，观片菜单外均不可用
            If Val(ufgStudyList.CurText("检查状态")) = 6 Then
                control.Enabled = False
            End If
            
    End Select
ErrHandle:
End Sub

Private Sub InitModuleParameter(Optional blnIsUpdateSearchTime As Boolean = True)
'功能:初始化模块级变量,仅窗体加载时调用一次
    Dim rsTemp As ADODB.Recordset
    
    '获取默认的查询方案id和系统查询方案id
    mlngDefQuerySchemeId = -1
    mlngSysQuerySchemeId = -1
    
    Set rsTemp = zlDatabase.OpenSQLRecord("select id,是否默认,是否系统查询 from 影像查询方案 where (是否默认=1 or 是否系统查询=1) and( 所属科室=0 or 所属科室 is null or 所属科室=[1]) order by 所属科室 desc,方案序号", "获取默认过滤方案", mlngCur科室ID)
    rsTemp.Filter = "是否默认=1"
    If rsTemp.RecordCount > 0 Then mlngDefQuerySchemeId = Val(Nvl(rsTemp!ID))
    rsTemp.Filter = "是否系统查询=1"
    If rsTemp.RecordCount > 0 Then mlngSysQuerySchemeId = Val(Nvl(rsTemp!ID))
    
    mSysPar.lngListColorMark = Nvl(GetDeptPara(mlngCur科室ID, "颜色显示类型", 0))
    mSysPar.blnNameColColorCfg = GetDeptPara(mlngCur科室ID, "姓名颜色区分", 0) = "1"         '姓名颜色区分
    mSysPar.blnOrdinaryNameColColorCfg = GetDeptPara(mlngCur科室ID, "缺省类型病人姓名颜色区分", 0) = "1"       '缺省类型病人姓名颜色区分
    
    If mSysPar.blnNameColColorCfg Then
        gstrSQL = "select 名称 from 病人类型 where 缺省标志=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取缺省病人类型")
        
        If rsTemp.RecordCount > 0 Then mstrDefaultPatientType = Nvl(rsTemp!名称)
    End If
    
    mSysPar.blnAutoPrint = Val(zlDatabase.GetPara("报到后自动打印申请单", glngSys, mlngModule, 0)) '报到后自动打印申请单
    
    mSysPar.blnChangeUser = GetDeptPara(mlngCur科室ID, "允许交换用户", 0) = "1"              '允许交换用户
    mSysPar.blnSwitchUser = GetDeptPara(mlngCur科室ID, "允许切换用户", 0) = "1"              '允许切换用户
    
    mSysPar.blnIsPetitionScan = IIf(Val(GetDeptPara(mlngCur科室ID, "启用申请单扫描", 1)) = 1, True, False)   '读取启用申请单扫描参数
    mSysPar.strImageLevel = Nvl(GetDeptPara(mlngCur科室ID, "影像质量等级", "甲,乙"))
    mSysPar.strReportLevel = Nvl(GetDeptPara(mlngCur科室ID, "报告质量等级", "甲,乙"))
    mSysPar.bln直接检查 = (Val(GetDeptPara(mlngCur科室ID, "登记后直接检查", 0)) = 1)         '登记后直接检查

    mSysPar.lngCriticalValues = Val(GetDeptPara(mlngCur科室ID, "危急情况判断", 0))           '危急情况判断
    mSysPar.blnIgnoreResult = GetDeptPara(mlngCur科室ID, "忽略结果阴阳性", 0) = "1" '        '忽略结果阴阳性
    mSysPar.lngConformDetermine = Val(GetDeptPara(mlngCur科室ID, "符合情况判定", 0))         '符合情况判定
    mSysPar.lngImageLevel = Val(GetDeptPara(mlngCur科室ID, "影像质量判定", 0))               '影像质量判定
    mSysPar.lngReportLevel = Val(GetDeptPara(mlngCur科室ID, "报告质量判定", 0))
    
    mSysPar.lngHintType = Val(GetDeptPara(mlngCur科室ID, "诊断结果提示类型", 0))
    
    mSysPar.blnFinishCommit = GetDeptPara(mlngCur科室ID, "无报告完成后直接完成", 0) = "1" '  '无报告完成后直接完成
    mSysPar.blnReportWithImage = GetDeptPara(mlngCur科室ID, "有图像才能写报告", 0) = "1" '   '有图像才能写报告
    mSysPar.blnReportWithResult = GetDeptPara(mlngCur科室ID, "无影像诊断为阴性", 0) = "1" '  '无影像诊断为阴性
    mSysPar.blnLocalizerBackward = GetDeptPara(mlngCur科室ID, "定位片后置", 0) = "1" '       '定位片后置
    mSysPar.blnCompleteCommit = GetDeptPara(mlngCur科室ID, "审核后直接完成", 0) = "1" '      '审核后直接完成
    
    mSysPar.lngBeforeDays = Val(GetDeptPara(mlngCur科室ID, "默认过滤天数", 2)) '                   '默认过滤天数
    If mSysPar.lngBeforeDays > 15 Or mSysPar.lngBeforeDays <= 0 Then
        mSysPar.lngBeforeDays = 2
    End If
    
    mSysPar.blnWriteCapDoctor = GetDeptPara(mlngCur科室ID, "采集图像者为检查技师", 0) = "1"  '采集图像者为检查技师
    
    mSysPar.blnPrintCommit = GetDeptPara(mlngCur科室ID, "打印后直接完成", 0) = "1" '           '打印后直接完成
    mSysPar.blnCanPrint = GetDeptPara(mlngCur科室ID, "平诊需审核才能打报告") = "1"             '平诊需要审核才能打印 =true
    mSysPar.blnAutoSendWorkList = GetDeptPara(mlngCur科室ID, "报道时自动发送WorkList") = "1"   '报道时自动发送WorkList

                
    '状态提醒
    mSysPar.lngEnregAfterTimeLen = Val(GetDeptPara(mlngCur科室ID, "登记后提醒", 0))
    mSysPar.lngCheckInAfterTimeLen = Val(GetDeptPara(mlngCur科室ID, "报到后提醒", 0))
    mSysPar.lngStudyAfterTimeLen = Val(GetDeptPara(mlngCur科室ID, "检查后提醒", 0))
    mSysPar.lngReportAfterTimeLen = Val(GetDeptPara(mlngCur科室ID, "报告后提醒", 0))
    mSysPar.lngAuditAfterTimeLen = Val(GetDeptPara(mlngCur科室ID, "审核后提醒", 0))
    
    '报告时观片
    mSysPar.blnShowImgAfterReport = (Val(zlDatabase.GetPara("报告时观片", glngSys, mlngModule, 0)) = 1)
    
    If InStr(mstrPrivs, "排队叫号") > 0 And mlngModule <> G_LNG_PATHOLSYS_NUM Then    '有权限使用才根据参数启用
        mSysPar.blnUseQueue = GetDeptPara(mlngCur科室ID, "启动排队叫号", 0) = "1" '          '默认不启用排队叫号
        
        If mSysPar.blnUseQueue Then
            mSysPar.lngQueueWay = GetDeptPara(mlngCur科室ID, "排队叫号方式", 0)             '排队叫号的排队方式
            mSysPar.blnSynStudylist = GetDeptPara(mlngCur科室ID, "同步定位检查列表", 0)
            mSysPar.blnAutoInQueue = GetDeptPara(mlngCur科室ID, "报到后自动排队", 1)
        Else
            mSysPar.lngQueueWay = 0
        End If
    End If
    
    mSysPar.blnRelatingPatient = GetDeptPara(mlngCur科室ID, "启动关联病人", 0) = "1"       '是否使用关
    mSysPar.lngRefreshInterval = Val(GetDeptPara(mlngCur科室ID, "自动刷新间隔", 0))  '     '自动刷新间隔,默认不自动刷新
    
    gblnXWLog = (Val(zlDatabase.GetPara("XW记录接口日志", glngSys, G_LNG_XWPACSVIEW_MODULE, "0")) = 1) '是否记录接口日志
    
    If mSysPar.lngRefreshInterval > 0 Then
        If mSysPar.lngRefreshInterval > 65 Then mSysPar.lngRefreshInterval = 65
        timerRefresh.Interval = mSysPar.lngRefreshInterval * 1000
        timerRefresh.Enabled = True
    Else
        timerRefresh.Enabled = False
    End If

    If blnIsUpdateSearchTime Then
        SQLCondition.开始时间 = CDate(Format(zlDatabase.Currentdate - (mSysPar.lngBeforeDays - 1), "yyyy-mm-dd 00:00"))
        
        mblnMoved = MovedByDate(SQLCondition.开始时间)
        
        If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
            gblnXWMoved = mblnMoved
        End If
    End If
        

    '初始化队列名称列表
    If mSysPar.lngQueueWay = 0 Then
        '初始化队列名称列表
        Dim iCount As Integer
        Dim strSql As String
        
        iCount = 1
        gstrSQL = "Select 执行间,检查设备 From 医技执行房间 where 科室id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取执行间名称", mlngCur科室ID)
        If rsTemp.EOF <> True Then
            ReDim mAstr队列名称(rsTemp.RecordCount) As String
            While rsTemp.EOF = False
                mAstr队列名称(iCount) = mlngCur科室ID & ":" & Nvl(rsTemp!执行间)
                iCount = iCount + 1
                rsTemp.MoveNext
            Wend
    
    '       以下数据用于测试
    '        ReDim mAstr队列名称(8) As String
    '        mAstr队列名称(1) = "42:CT98"
    '        mAstr队列名称(2) = "42:CT99"
    '        mAstr队列名称(3) = "61:CT2"
    '        mAstr队列名称(4) = "61:CT1"
    '        mAstr队列名称(5) = "81:jy1"
    '        mAstr队列名称(6) = "81:jy2"
    '        mAstr队列名称(7) = "82:放射科"
    '        mAstr队列名称(8) = "83:" & Nvl(rsTemp!执行间)
            
        Else
            ReDim mAstr队列名称(0) As String
        End If
    Else
        ReDim mAstr队列名称(1) As String

        mAstr队列名称(1) = mstrCur科室

    End If
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = PicList.hWnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picWindow.hWnd
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
On Error GoTo ErrHandle
    '禁止检查列表 拖动
    Cancel = IIf(((Action = 4 Or Action = 6 Or Action = 5) And Not Pane.Hidden), True, False)
ErrHandle:
End Sub


Private Sub InitStudyList()
    Dim strCols As String
    Dim strDefaultCols As String
    Dim i As Integer
    Dim arrCol() As String
    Dim strTemp As String
    
    
    strCols = zlDatabase.GetPara("检查列表", glngSys, mlngModule, "")
    
    Set ufgStudyList.ImageList = imgList
    
    Select Case mlngModule
         Case G_LNG_PACSSTATION_MODULE   '医技
            strDefaultCols = Replace(IIf(mrtReportType = 报告文档编辑器, M_STR_PUBLIC_COLS_NEW, M_STR_PUBLIC_COLS), "[------]", M_STR_IMAGES_COLS)
                
        Case G_LNG_PATHOLSYS_NUM        '病理
            strDefaultCols = Replace(IIf(mrtReportType = 报告文档编辑器, M_STR_PUBLIC_COLS_NEW, M_STR_PUBLIC_COLS), "[------]", M_STR_PATHOL_COLS)
            
        Case G_LNG_VIDEOSTATION_MODULE  '采集
            strDefaultCols = Replace(IIf(mrtReportType = 报告文档编辑器, M_STR_PUBLIC_COLS_NEW, M_STR_PUBLIC_COLS), "[------]", M_STR_CAPTOR_COLS)
    End Select
    
    
    arrCol() = Split(strCols, "|")
    
    For i = 0 To UBound(arrCol())
        If arrCol(i) <> "" Then
            If InStr(arrCol(i), "申请单") > 0 Then
                strTemp = arrCol(i)
                
                If mSysPar.blnIsPetitionScan Then
                    '当启用申请单扫描时，则申请单列允许进行配置
                    strCols = Replace(strCols, strTemp, Replace(strTemp, ",uncfg", ""))
                Else
                    '当未启用申请单时，不允许对申请单列进行配置
                    strCols = Replace(strCols, strTemp, Replace(Replace(strTemp, ",hide", ""), ",uncfg", "") & ",hide,uncfg")
                    
                    strDefaultCols = Replace(strDefaultCols, "申请单>申请单医嘱,w1100", "申请单>申请单医嘱,w1100,hide,uncfg")
                End If

                Exit For
            End If
        End If
    Next i
    
    
    ufgStudyList.DefaultColNames = strDefaultCols
    ufgStudyList.ColNames = IIf(strCols = "", strDefaultCols, strCols)
    
    ufgStudyList.IsKeepRows = False
    ufgStudyList.IsCopyMode = False
    ufgStudyList.IsAutoRowHeight = False
End Sub


Private Sub InitForm()
    Dim strKinds As String
    Dim blnDo As Boolean
    Dim lngKey As Long
    Dim bytFontSize As Byte
    
    Call WriteLog("InitForm -> Step 1：开始执行...")
    
    '得到个性化风格参数
    blnDo = Val(zlDatabase.GetPara("使用个性化风格")) <> 0
    
    mstrPrivs = gstrPrivs '权限
    mlngModule = glngModul '模块号
    mlngCur科室ID = 0
    mstrCur科室 = ""
    mstrCanUse科室 = ""
    mblnAllDepts = False
    mlngSortCol = 0
    mintSortOrder = 0
    mSysPar.lngQueueWay = 0
    
    '读取字体大小
    bytFontSize = Val(zlDatabase.GetPara("显示字体大小", glngSys, glngModul))
    mbytFontSize = IIf(bytFontSize = 0, 9, IIf(bytFontSize = 1, 12, 15))
    '初始字体状态
    mbyrFontState = 2
    
    
    
    mblnInitOk = False  '初始数据,初始化完成之前不进行数据的提取
    mblnvsRefresh = False
    mblnMenuDownState = False
    mlngFilterTab = 0
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then labHistory.Caption = "病理历史："
    
    
    '判断当前用户是否具有 观片站的基本权限
    mblnObserve = IIf(InStr(GetPrivFunc(glngSys, 1289), "基本") > 0, True, False)
    
    If mlngModule = G_LNG_PATHSTATION_MODULE Then
        mlngFilterTab = Val(zlDatabase.GetPara("过滤页面", glngSys, glngModul))
        
        tabFilter.Visible = True
        picExeState.Visible = True
        
        Call InitFilterPage
    End If
    
    Call WriteLog("InitForm -> Step 2：载入本地注册表参数...")
    
    '判断当前用户是否具有“影像设备目录”的权限，有此权限才可以设置新网的PACS参数
    mblnSetXWParam = IIf(InStr(GetPrivFunc(glngSys, G_LNG_XWPACSVIEW_MODULE), "PACS参数设置") > 0, True, False)
    
    Call InitLocalPars '本地注册表参数
    
    Call WriteLog("InitForm -> Step 3：载入部门相关信息...")
    If Not InitDepts Then Unload Me: Exit Sub '初始化医技科室
    
    mrtReportType = GetDeptPara(mlngCur科室ID, "报告编辑器", 0)                 '报告编辑器
    
    ReDim gConnectedShardDir(0) As String   '初始化共享目录连接串
    
    Call WriteLog("InitForm -> Step 4：初始化部门级参数...")
    Call InitModuleParameter '初始化模块级变量
    
    
    '初始子窗体
    Set mobjEvent = New clsEvent
    Set gobjEvent = mobjEvent
    
    
    '根据参数判断是否启用消息中心
    Set mobjMsgCenter = New clsPacsMsgProcess
    Call mobjMsgCenter.OpenMsgCenter(mlngModule, mlngCur科室ID, mstrPrivs)
    
    Set mobjPacsCore = New zl9PacsCore.clsViewer
    
    
    If mSysPar.blnUseQueue And InStr(GetPrivFunc(glngSys, 1160), "基本") > 0 Then
        Set mobjQueue = New frmWork_Queue
        Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur科室ID, NeedName(mstrCur科室), mstrPrivs)
    Else
        Set mobjQueue = Nothing
    End If

    Call WriteLog("InitForm -> Step 5：读取列表颜色配置...")
    Call ReadStudyListColor(mlngCur科室ID)
    
    Call WriteLog("InitForm -> Step 6：读取快速过滤配置...")
    Call InitFilterCmd
    
    Call WriteLog("InitForm -> Step 7：初始化窗口菜单...")
    Call InitCommandBars
'    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call WriteLog("InitForm -> Step 8：初始化界面布局...")
    Call InitFaceScheme
    
    Call WriteLog("InitForm -> Step 9：初始化检查数据列表...")
    Call InitStudyList
    
     '如果注册表中工具栏相关值为空 并且 已勾选个性化设置，那么向注册表写入工具栏显示模式值
    If mintToolBarWriteReg = 9 Or (mintToolBarWriteReg = 0 And blnDo) Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 3
    End If
    
    '恢复窗体的状态   注：恢复窗体状态 必须放在 向注册表写入工具栏显示模式值 的语句后面，否则会造成工具栏显示模式有误。
    Call RestoreWinState(Me, App.ProductName)
    
    picAppend.Height = Nvl(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "StudyInfHeight", picAppend.Height))
    
     '工具栏--- 文本标签 的设置使用RestoreWinState 恢复不了，还需要单独处理，如未勾选个性化设置，则工具栏默认显示图标和文本
    If blnDo Then
        If Me.cbrMain(2).Controls(1).Style = xtpButtonIconAndCaption Then
            Me.cbrMain(2).ShowTextBelowIcons = True
        Else
            Me.cbrMain(2).ShowTextBelowIcons = False
        End If
    Else
        Me.cbrMain(2).ShowTextBelowIcons = True
    End If
    
    ClearCacheFolder App.Path & "\TmpImage\"    '若临时目录满了，则清空该目录
    
    
    '判断临时目录是否存在
    If Dir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage", vbDirectory) = "" Then
        Call MkDir(IIf(Len(App.Path) > 3, App.Path & "\", App.Path & "") & "TmpImage")
    End If
    
    
    '初始化双用户登陆的参数
    mblnCnOracleIsHIS = True
    mintChangeUserState = 1
    mstrUserNameHIS = UserInfo.姓名
    mstrUserNameNew = UserInfo.姓名
    mstrUserIDHIS = UserInfo.用户名
    mstrUserIDNew = UserInfo.用户名
    
    Set mcnOracleHIS = gcnOracle
    
    Me.stbThis.Panels(4).Text = "报告医生：" & mstrUserNameHIS & "   检查医生：" & mstrUserNameNew
    
    ReDim mobjPacsReportArry(0) As frmReport
    
    gblnUseXinWangView = False
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        gblnUseXinWangView = IsUseXwViewer
    
        '如果是RIS工作站，则连接新网数据库，读取参数
        If gblnUseXinWangView Then
            '挂上截获消息的hook
            plngXWPreWndProc = XWHook(Me.hWnd)
        End If
    End If
    
    mblnFormLoadState = True
    
    Call WriteLog("InitForm -> Step 10：结束执行...")
End Sub


'Private Sub Form_Load()
'On Error GoTo errHandle
'    '初始化相关方法在showstation中调用InitForm进行处理......
'    '这里不能进行相关的初始化处理是因为在clsPacsWork的BHCodeMain方法中，设置显示方式的时候，会触发Load事件，
'    '而Load事件中的某些处理需要相关参数才能正确执行，因此需要将Load中的处理方法单独提取出来，放入ShowStation方法中执行...
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'End Sub


Private Sub RefreshStatusBarInf()
    Dim i As Long
    
    Dim lngDengJi As Long
    Dim lngBaoDao As Long
    Dim lngJianCha As Long
    Dim lngBaoGao As Long
    Dim lngShenHe As Long
    Dim lngBoHui As Long
    Dim lngWanCheng As Long
    Dim lngYiBaoGao As Long
    Dim strTemp As String
    
    
    lngDengJi = 0
    lngBaoDao = 0
    lngJianCha = 0
    lngBaoGao = 0
    lngShenHe = 0
    lngBoHui = 0
    lngWanCheng = 0
    lngYiBaoGao = 0
    
    
    For i = 1 To ufgStudyList.GridRows - 1
        Select Case ufgStudyList.Text(i, "检查过程")
            Case "已登记"
                lngDengJi = lngDengJi + 1
            Case "已报到"
                lngBaoDao = lngBaoDao + 1
            Case "已检查"
                lngJianCha = lngJianCha + 1
            Case "已报告"
                lngYiBaoGao = lngYiBaoGao + 1
            Case "报告中"
                lngBaoGao = lngBaoGao + 1
            Case "已审核"
                lngShenHe = lngShenHe + 1
            Case "已驳回"
                lngBoHui = lngBoHui + 1
            Case "已完成"
                lngWanCheng = lngWanCheng + 1
        End Select
    Next i
    
    strTemp = ""
    If lngDengJi > 0 Then strTemp = "已登记：" & lngDengJi & "    "
    If lngBaoDao > 0 Then strTemp = strTemp & "已报到：" & lngBaoDao & "    "
    If lngJianCha > 0 Then strTemp = strTemp & "已检查：" & lngJianCha & "    "
    If lngBaoGao > 0 Then strTemp = strTemp & "报告中：" & lngBaoGao & "    "
    If lngYiBaoGao > 0 Then strTemp = strTemp & "已报告：" & lngYiBaoGao & "    "
    If lngShenHe > 0 Then strTemp = strTemp & "已审核：" & lngShenHe & "    "
    If lngBoHui > 0 Then strTemp = strTemp & "已驳回：" & lngBoHui & "    "
    If lngWanCheng > 0 Then strTemp = strTemp & "已完成：" & lngWanCheng & "    "
    
    stbThis.Panels(2).Text = "共 " & ufgStudyList.GridRows - 1 & " 条记录": stbThis.Panels(2).Alignment = sbrCenter
    stbThis.Panels(3).Text = strTemp
End Sub


Private Sub InitFilterPage()
    Dim lngHideCount As Long
    
    lngHideCount = 0
    With tabFilter
        .RemoveAll
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .RemoveAll
        



        .InsertItem 0, "取  材", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "取材"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "病理取材")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 1, "制  片", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "制片"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "病理制片")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 2, "免  疫", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "免疫"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "免疫组化")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 3, "分  子", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "分子"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "分子病理")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1


        .InsertItem 4, "特  染", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "特染"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "特殊染色")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 5, "会  诊", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "会诊"
        .Item(tabFilter.ItemCount - 1).Visible = CheckPopedom(mstrPrivs, "会诊反馈")
        If Not .Item(tabFilter.ItemCount - 1).Visible Then lngHideCount = lngHideCount + 1
        
        
        .InsertItem 6, "所  有", picTag.hWnd, 0
        .Item(tabFilter.ItemCount - 1).tag = "所有"
        
    End With

    '当所有功能标签被隐藏时，则直接隐藏tabFilter控件
    tabFilter.Visible = (lngHideCount < tabFilter.ItemCount - 1)
    tabFilter.tag = (lngHideCount < tabFilter.ItemCount - 1)
    
    On Error GoTo errContinue1
    If tabFilter.tag Then
        If Not tabFilter.Item(mlngFilterTab).Visible Then
            tabFilter.Item(tabFilter.ItemCount - 1).Selected = True
        Else
            tabFilter.Item(mlngFilterTab).Selected = True
        End If
    End If
    
    optAccept.Enabled = IIf(tabFilter.Selected.tag = "取材" Or tabFilter.Selected.tag = "会诊" Or tabFilter.Selected.tag = "所有", False, True)
    
    optNeed.Enabled = IIf(tabFilter.Selected.tag = "所有", False, True)
    optFinal.Enabled = IIf(tabFilter.Selected.tag = "所有", False, True)
    optAll.Enabled = IIf(tabFilter.Selected.tag = "所有", False, True)
errContinue1:
End Sub


Private Function GetWindowCaption() As String
    GetWindowCaption = Mid(Me.Caption & " ", 1, InStr(Me.Caption & " ", " "))
End Function


Private Sub DisposeObj()
    If Not mfrmWork_PacsImg Is Nothing Then
        Unload mfrmWork_PacsImg
        Set mfrmWork_PacsImg = Nothing
    End If
    
    If Not mobjQueue Is Nothing Then
        Unload mobjQueue
        Set mobjQueue = Nothing
    End If
    
    If Not mobjPacsCore Is Nothing Then
        mobjPacsCore.Closefrom
        Set mobjPacsCore = Nothing
    End If
    
    If Not mfrmPACSFilter Is Nothing Then
        Unload mfrmPACSFilter
        Set mfrmPACSFilter = Nothing
    End If
    
    If Not mobjWork_Pathol Is Nothing Then
        Call mobjWork_Pathol.Free
        Set mobjWork_Pathol = Nothing
    End If
    
    If Not mobjWork_His Is Nothing Then
        Call mobjWork_His.Free
        Set mobjWork_His = Nothing
    End If
    
    If Not mobjWork_Report Is Nothing Then
        Call mobjWork_Report.Free
        Set mobjWork_Report = Nothing
    End If
    
    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        If Not mobjCaptureHot Is Nothing Then
            Call mobjCaptureHot.FreeHook
            Set mobjCaptureHot = Nothing
        End If
    End If
    
    If mblnUseActivexCapture Then
        '使用Activex的视频采集方式退出
        Set mobjWork_ActiveVideo = Nothing
    End If
    
    If Not gobjMsgCenter Is Nothing Then
        Set gobjMsgCenter = Nothing
    End If

        
    Set mobjEvent = Nothing
    Set mobjSquareCard = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrHandle
    If mblnUseActivexCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlNotifyQuit
        End If
    End If
    
    '关闭消息中心
    If Not gobjMsgCenter Is Nothing Then
        Call gobjMsgCenter.CloseMsgCenter
    End If
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序列", mlngSortCol)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序方向", mintSortOrder)
    
'    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "StudyInfHeight", picAppend.Height)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "StudyListWidth", PicList.Width / Me.ScaleWidth)
        
    '保存字体参数
    zlDatabase.SetPara "显示字体大小", IIf(mbytFontSize = 9, 0, IIf(mbytFontSize = 12, 1, IIf(mbytFontSize = 15, 2, mbytFontSize))), glngSys, glngModul
    
    '恢复窗口名称
    Me.Caption = GetWindowCaption
    
    Call SaveWinState(Me, App.ProductName)
    
    Call DisposeObj
    
    '恢复导航台的数据库联接
    If mblnCnOracleIsHIS = False Then
        Set gcnOracle = mcnOracleHIS
        InitCommon gcnOracle
        RegCheck
        SetDbUser mstrUserIDHIS
        Call GetUserInfo
        Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
    End If
    
    frmTwoUser.intDBState = 1
    
    '如果是RIS工作站，则断开跟新网数据库的连接
    If gblnUseXinWangView Then
        '    卸载hook
        XWUnhook Me.hWnd, plngXWPreWndProc
    End If
    
    mblnFormLoadState = False
    
    Exit Sub
ErrHandle:
    Debug.Print err.Description
End Sub

Private Function InitCardType(ByVal strCardNames As String) As String
'按指定格式初始化卡类型
    Dim i As Integer
    Dim aryKindInfo() As String
    Dim strKinds As String
    
    aryKindInfo = Split(strCardNames, ";")
    
    strKinds = ""
    For i = 0 To UBound(aryKindInfo) - 1
        If strKinds <> "" Then strKinds = strKinds & ";"
        strKinds = strKinds & aryKindInfo(i) & "|" & aryKindInfo(i) & "|-1"
    Next i
    
    InitCardType = strKinds & ";"
End Function

Private Sub InitLocalPars()
    Dim strTemp As String
    Dim strTempArry() As String
    Dim i As Integer
'初始化临时本地参数，以个人设置为主,窗体加载，过滤，本地设置等调用

    mstrCaptureHot = GetSetting("ZLSOFT", "公共模块", "采集热键", "F8")
    mstrCaptureAfterHot = GetSetting("ZLSOFT", "公共模块", "后台采集热键", "F7")
    mstrCaptureAfterTagHot = GetSetting("ZLSOFT", "公共模块", "标记更新热键", "F6")
    
    mblncmd门诊 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "门诊病人", 1))
    mblncmd住院 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "住院病人", 1))
    mblncmd外诊 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "外诊病人", 1))
    mblncmd体检 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "体检病人", 1))
    mblncmd已缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用已缴", 0))
    mblncmd未缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用未缴", 0))
    mblncmd无费 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用无费", 0))
    mblncmd记账 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用记账", 0))
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        mblncmd补缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用补缴", 0))
        
        mblnPopChangGuiWindow = (Val(zlDatabase.GetPara("常规质量窗口", glngSys, mlngModule, 0)) = 1)
        mblnPopKuaiShuWindow = (Val(zlDatabase.GetPara("快速石蜡质量窗口", glngSys, mlngModule, 0)) = 1)
        mblnPopBingDongWindow = (Val(zlDatabase.GetPara("冰冻质量窗口", glngSys, mlngModule, 0)) = 1)
        mblnPopXiBaoWindow = (Val(zlDatabase.GetPara("细胞质量窗口", glngSys, mlngModule, 0)) = 1)
        mblnPopHuiZhenWindow = (Val(zlDatabase.GetPara("会诊质量窗口", glngSys, mlngModule, 0)) = 1)
        mblnPopShiJianWindow = (Val(zlDatabase.GetPara("尸检质量窗口", glngSys, mlngModule, 0)) = 1)
    End If
    
    mblncmd登记 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记病人", 1))
    mblncmd报到 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报到病人", 1))
    mblncmd检查 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查病人", 1))
    mblncmd报告 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报告病人", 1))
    mblncmd审核 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "审核病人", 1))
    mblncmd驳回 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "驳回病人", 1))
    mblncmd完成 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "完成病人", 1))
    
    mlngLocateFindType = zlDatabase.GetPara("定位查找方式", glngSys, mlngModule, 0)
        
    mstrFindWay = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "过滤方式", GetStudyNumberDisplayName)
    mstrLocateWay = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "定位方式", GetStudyNumberDisplayName)
    
    PatiIdentify.IDKindStr = InitCardType(Replace(IIf(mlngLocateFindType = TLocateFindType.lftLocate, CONST_STR_LOCAL_CARD_TYPE, CONST_STR_FIND_CARD_TYPE), "[------]", GetStudyNumberDisplayName))
    PatiIdentify.IDKindIDX = PatiIdentify.GetKindIndex(IIf(mlngLocateFindType = 0, mstrLocateWay, mstrFindWay))
    
    mblncmd本次 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "本次住院", "0"))
    mlngSortCol = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序列", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序方向", 0))
    
    strTemp = zlDatabase.GetPara("影像类别过滤", glngSys, mlngModule, "")
    
    ReDim strTempArry(0)
    ReDim mblncmd影像类别(0)
    
    On Error GoTo errContinue1
    strTempArry = Split(strTemp, ",")
    If UBound(strTempArry) >= 0 Then ReDim mblncmd影像类别(UBound(strTempArry))
    
    For i = 0 To UBound(strTempArry)
        mblncmd影像类别(i) = IIf(UCase(strTempArry(i)) = "TRUE", True, False)
    Next i
        
    ReDim mblncmd影像执行间(0)
errContinue1:
    mSysPar.strFirstTab = zlDatabase.GetPara("工作首页", glngSys, mlngModule, "") '为空表示不使用定制工作首页功能
    mSysPar.blnAutoOpenReport = (Val(zlDatabase.GetPara("开始检查自动打开报告", glngSys, mlngModule, 0)) = 1)
    mSysPar.blnNoShowCancel = (Val(zlDatabase.GetPara("不显示被取消的登记", glngSys, mlngModule, 0)) = 1)
    mSysPar.blnPatTrack = (Val(zlDatabase.GetPara("病人跟踪", glngSys, mlngModule, 0)) = 1)
    mSysPar.strLocalRoom = zlDatabase.GetPara("本机执行间名称", glngSys, mlngModule, "")
    
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        '如果是采集模块，才需要执行该参数
        mSysPar.lngMoneyExeModle = Val(zlDatabase.GetPara("采集费用执行模式", glngSys, mlngModule, 0))
    End If
    
    '得到注册表中关于工具栏显示状态的值，如果为空则等于9
    mintToolBarWriteReg = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\CommandBars", "cbrMainButtonText", 9))
    
    
    With SQLCondition '------------------------ '过滤条件初始
        '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
       .时间类型 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "过滤时间类型", 1))
       .单据号 = ""
       .门诊号 = 0
       .住院号 = 0
       .健康号 = ""
       .就诊卡 = ""
       .姓名 = ""
       .性别 = ""
       .开始年龄 = -1
       .结束年龄 = -1
       .年龄条件 = "="
       .检查号 = 0
       .身份证 = ""
       .IC卡 = ""
       .病人科室 = 0
       .标本部位 = ""
       .诊断医生 = ""
       .审核医生 = ""
       .疾病诊断 = ""
       .报告内容 = ""
       .结果阳性 = -1
       .影像质量 = ""
       .检查技师 = ""
       .检查过程 = ""
       .影像类别 = ""
       .检查所见 = ""
       .诊断意见 = ""
       .建议 = ""
       .随访 = ""
    End With
End Sub

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim str科室IDs As String, str来源 As String
    
    str来源 = "1,2,3"
    If InStr(mstrPrivs, "所有科室") > 0 Then
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')> 0 And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    Else
        strSql = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=" & UserInfo.ID & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')>0  And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    End If
   

    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, GetWindowCaption, CStr("," & str来源 & ","))
    
    If rsTmp.EOF Then
        MsgBoxD Me, "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
        Exit Function
    Else
        str科室IDs = GetUser科室IDs
        Do Until rsTmp.EOF
            mstrCanUse科室 = mstrCanUse科室 & "|" & rsTmp!ID & "_" & rsTmp!编码 & "-" & rsTmp!名称
            mstrCanUse科室IDs = mstrCanUse科室IDs & "," & rsTmp!ID
            
            If rsTmp!ID = UserInfo.部门ID Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '提取默认科室
            If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur科室ID = 0 Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '没有默认科室,取所属检查科室第一个
            rsTmp.MoveNext
        Loop
        
        mstrCanUse科室 = Mid(mstrCanUse科室, 2)
        mstrCanUse科室IDs = Mid(mstrCanUse科室IDs, 2)
        
        If InStr(mstrPrivs, "所有科室") > 0 And mlngCur科室ID = 0 Then
            mlngCur科室ID = Split(Split(mstrCanUse科室, "|")(0), "_")(0)
            mstrCur科室 = Split(Split(mstrCanUse科室, "|")(0), "_")(1)
        End If
        
        If mlngCur科室ID = 0 And InStr(mstrPrivs, "所有科室") <= 0 Then '没有所有科室操作权限,而且操作者科室不属于检查类科室
            MsgBoxD Me, "没有发现你所属科室,不能使用此工作站。", vbInformation, gstrSysName
            Exit Function
        End If
        
        InitDepts = True
    End If
    
    If mlngModule = G_LNG_PACSSTATION_MODULE And gblnUseXinWangView Then
        glngXWDeptID = mlngCur科室ID
    End If
End Function

Private Sub InitFaceScheme()
    Dim lngListWidth As Double
    
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    lngListWidth = Nvl(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "StudyListWidth", 0.35))
    If lngListWidth >= 1 Then lngListWidth = 0.35
    
    '注册表中保存的界面布局Pnae数量不对，则加载默认的Pane设置
    If dkpMain.PanesCount <> 3 Then
        dkpMain.DestroyAll
        
        Set Pane1 = dkpMain.CreatePane(1, lngListWidth * 100, 250, DockLeftOf, Nothing)
        Pane1.Title = "检查列表"
        Pane1.Handle = PicList.hWnd
        Pane1.Options = PaneNoCloseable Or PaneNoFloatable
        
        Set Pane2 = dkpMain.CreatePane(2, (1 - lngListWidth) * 100, 300, DockRightOf, Nothing)
        Pane2.Title = "子窗体"
        Pane2.Handle = picWindow.hWnd
        Pane2.Options = PaneNoCaption Or PaneNoCloseable
    End If
End Sub

'当快速工具栏参数变化时保存
Private Sub SaveFilterCmd()
    Dim strTemp As String
    Dim i As Integer
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "门诊病人", IIf(mblncmd门诊, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "住院病人", IIf(mblncmd住院, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "外诊病人", IIf(mblncmd外诊, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "体检病人", IIf(mblncmd体检, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用已缴", IIf(mblncmd已缴, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用未缴", IIf(mblncmd未缴, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用记账", IIf(mblncmd记账, 1, 0)
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用补缴", IIf(mblncmd补缴, 1, 0)
    End If
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用无费", IIf(mblncmd无费, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记病人", IIf(mblncmd登记, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报到病人", IIf(mblncmd报到, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查病人", IIf(mblncmd检查, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报告病人", IIf(mblncmd报告, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "审核病人", IIf(mblncmd审核, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "驳回病人", IIf(mblncmd驳回, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "完成病人", IIf(mblncmd完成, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "过滤方式", mstrFindWay
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "定位方式", mstrLocateWay
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "本次住院", IIf(mblncmd本次, 1, 0)
    
    If mlngModule = G_LNG_PATHSTATION_MODULE Then
        '病理模块单独处理的部分
        Call zlDatabase.SetPara("常规过滤", IIf(mblncmd常规, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("快速石蜡过滤", IIf(mblncmd快速石蜡, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("冰冻过滤", IIf(mblncmd冰冻, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("细胞过滤", IIf(mblncmd细胞, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("会诊过滤", IIf(mblncmd会诊, 1, 0), glngSys, glngModul)
        Call zlDatabase.SetPara("尸检过滤", IIf(mblncmd尸检, 1, 0), glngSys, glngModul)
        
        Call zlDatabase.SetPara("过滤页面", tabFilter.Selected.Index, glngSys, glngModul)
    End If
    
    If UBound(mblncmd影像类别) >= 0 Then
        strTemp = mblncmd影像类别(0)
    End If
    For i = 1 To UBound(mblncmd影像类别)
        strTemp = strTemp & "," & mblncmd影像类别(i)
    Next i
    Call zlDatabase.SetPara("影像类别过滤", strTemp, glngSys, mlngModule)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        If UBound(mblncmd影像执行间) >= 0 Then
            strTemp = mlngCur科室ID & ":" & IIf(mblncmd影像执行间(0), "T", "F")
        End If
        
        For i = 1 To UBound(mblncmd影像执行间)
            strTemp = strTemp & "," & IIf(mblncmd影像执行间(i), "T", "F")
        Next i
        
        '替换对应科室的执行间配置
        If mstrAllExamineRoomCfg = "" Or InStr(mstrAllExamineRoomCfg, ":") <= 0 Then
            mstrAllExamineRoomCfg = strTemp
        Else
            If InStr(mstrAllExamineRoomCfg, mlngCur科室ID & ":" & mstrCurExamineRoomCfg) > 0 Then
                mstrAllExamineRoomCfg = Replace(mstrAllExamineRoomCfg, mlngCur科室ID & ":" & mstrCurExamineRoomCfg, strTemp)
            Else
                mstrAllExamineRoomCfg = mstrAllExamineRoomCfg & "|" & strTemp
            End If
        End If
        
        '设置成当前配置的值
        If strTemp <> "" Then mstrCurExamineRoomCfg = Split(strTemp, ":")(1)
        
        Call zlDatabase.SetPara("影像执行间过滤", mstrAllExamineRoomCfg, glngSys, mlngModule)
    End If
    
    '保存检查部位
    Call zlDatabase.SetPara("检查部位过滤", mstrcmd部位分组, glngSys, mlngModule)
End Sub

Private Sub InitFilterCmd()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim objPopbar As CommandBarPopup, objCusControl As CommandBarControlCustom
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    Dim i As Integer
    Dim strStudyTypes As String

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrdock.VisualTheme = xtpThemeOfficeXP
    With Me.cbrdock.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    cbrdock.AddImageList img16 '以VB.ImageList的Tag与ID进行关联
    cbrdock.EnableCustomization False
    cbrdock.ActiveMenuBar.Visible = False
    
    Set objBar = cbrdock.Add("来源", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        '来源.........................................................
        Set objControl = .Add(xtpControlButtonPopup, ID_来源, "来源")
        objControl.ToolTipText = "根据病人来源进行过滤"
        
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_门诊, "门诊(&1)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_住院, "住院(&2)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_外诊, "外诊(&3)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_体检, "体检(&4)")
        
        For Each cbrPopControl In objControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
            
            
        '状态.........................................................
        Set objControl = .Add(xtpControlButtonPopup, ID_状态, "状态")
        objControl.ToolTipText = "根据检查状态进行过滤"
        
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_登记, "登记(&1)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_报到, "报到(&2)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_检查, "检查(&3)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_报告, "报告(&4)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_审核, "审核(&5)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_驳回, "驳回(&6)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_完成, "完成(&7)")
    
        For Each cbrPopControl In objControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
        
            
        If mlngModule = G_LNG_PATHSTATION_MODULE Then
            '只有病理系统才有病理类别
            Set objControl = .Add(xtpControlButtonPopup, ID_病理类别, "类别")
            objControl.ToolTipText = "根据病理类别进行过滤"
            
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_病理类别_常规, "常规(&1)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_病理类别_冰冻, "冰冻(&2)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_病理类别_细胞, "细胞(&3)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_病理类别_尸检, "尸检(&4)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_病理类别_会诊, "会诊(&5)")
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_病理类别_快速石蜡, "快速石蜡(&6)")
        Else
            '添加所有影像类别
            Set objControl = .Add(xtpControlButtonPopup, ID_影像类别, "类别   ")
            objControl.ToolTipText = "根据影像类别进行过滤"
            
            strSql = "select 编码,名称 from 影像检查类别"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "影像检查类别")
            
            i = 1
            mintcmd影像类别 = 0
            strTemp = ""
            If rsTemp.RecordCount > 0 Then
                ReDim Preserve mblncmd影像类别(rsTemp.RecordCount - 1)
                
                While rsTemp.EOF = False
                    Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_影像类别 + i, rsTemp("名称") & "(&" & Chr(64 + i) & ")")
                    
                    cbrPopControl.DescriptionText = rsTemp("编码")
                    cbrPopControl.Style = xtpButtonIconAndCaption
                    cbrPopControl.Checked = mblncmd影像类别(i - 1)
                    cbrPopControl.CloseSubMenuOnClick = False
                    
                    If mblncmd影像类别(i - 1) = True Then
                        strStudyTypes = strStudyTypes & "," & rsTemp("名称")
                        
                        mintcmd影像类别 = mintcmd影像类别 + 1
                        strTemp = IIf(strTemp = "", cbrPopControl.Caption, strTemp & "," & cbrPopControl.Caption)
                    End If
                    
                    rsTemp.MoveNext
                    i = i + 1
                Wend
                
                If strStudyTypes <> "" Then strStudyTypes = Mid(strStudyTypes, 2)
                
                If strTemp <> "" Then
                    objControl.ToolTipText = "显示影像类别为[" & strTemp & "]的检查"
                    objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
                End If
            End If
        End If
        
        For Each cbrPopControl In objControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
        
        '费用.........................................................
        Set objControl = .Add(xtpControlButtonPopup, ID_费用, " 费用")
            objControl.ToolTipText = "根据费用状态进行过滤"
            
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_未缴, "未缴(&1)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_已缴, "已缴(&2)")
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_记账, "记账(&3)")
        
        If mlngModule = G_LNG_PATHOLSYS_NUM Then
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_补缴, "补缴(&4)")
        End If
        
        '如果没有补缴菜单，则使用数字3的按键作为快捷键
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_无费, "无费(&" & IIf(mlngModule = G_LNG_PATHOLSYS_NUM, 5, 4) & ")")
        
        For Each cbrPopControl In objControl.CommandBar.Controls
            cbrPopControl.CloseSubMenuOnClick = False
        Next
        
        '添加所有影像执行间
        If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            Set objControl = .Add(xtpControlButtonPopup, ID_影像执行间, "执行间   ")
            objControl.ToolTipText = "根据影像执行间进行过滤"
            
            Call InitExamineRoom(objControl, cbrPopControl, mlngCur科室ID)
        End If
        
        '添加所有检查部位
        If mlngModule = G_LNG_PACSSTATION_MODULE Or mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            Set objControl = .Add(xtpControlButtonPopup, ID_检查部位, "部位   ")
            objControl.ToolTipText = "根据检查部位进行过滤"
            
            Call InitStudyPlace(objControl, cbrPopControl, strStudyTypes)
        End If
    End With
    
    For Each objControl In objBar.Controls
        If objControl.type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("过滤", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_查找方式, "")
        objControl.Style = xtpButtonIcon
        objControl.IconId = IIf(mlngLocateFindType = TLocateFindType.lftLocate, 3, 4)
        
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_查找值, "查找值")
        objCusControl.Handle = PatiIdentify.hWnd
        objCusControl.flags = xtpFlagRightAlign
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_开始查找, IIf(mlngLocateFindType = TLocateFindType.lftLocate, "开始定位", "开始查找"))
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = conMenu_View_Filter
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_本次住院, "本次")
    objControl.ToolTipText = "只显示本次住院检查记录"
    objControl.Style = xtpButtonIconAndCaption
    objControl.IconId = conMenu_View_Filter
    
    With cbrdock.KeyBindings
        .Add FCONTROL, Asc("G"), ID_开始查找
    End With
    cbrdock.RecalcLayout
End Sub

Private Sub InitExamineRoom(objControl As CommandBarControl, cbrPopControl As CommandBarControl, ByVal lngCur科室ID As Long)
'初始化执行间过滤配置
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    Dim strTemp As String
    Dim strTempArry() As String
    
    Dim i As Integer
    Dim strID As String
    
    '读取执行间配置,格式:科室1ID:执行间1选择情况,执行间2选择情况,...|科室2ID:执行间1选择情况,执行间2选择情况,...|...
    '示例：64:T,F,T,F|65:T,F,T,F|...
    mstrAllExamineRoomCfg = zlDatabase.GetPara("影像执行间过滤", glngSys, mlngModule, "")
    
    For i = 0 To UBound(Split(mstrAllExamineRoomCfg, "|"))
        If Val(Split(Split(mstrAllExamineRoomCfg, "|")(i), ":")(0)) = mlngCur科室ID Then
            mstrCurExamineRoomCfg = Split(Split(mstrAllExamineRoomCfg, "|")(i), ":")(1)
            strTemp = mstrCurExamineRoomCfg
            Exit For
        End If
    Next
    
    If mblnAllDepts Then
        If InStr(mstrPrivs, "所有科室") > 0 Then
            strSql = "select 名称,执行间 from 医技执行房间 a, 部门表 b where a.科室Id=b.Id "
        Else
            '查询对应人员所在科室中所包含的执行间
            strSql = "select 名称,执行间 from 医技执行房间 a, 部门人员 b,部门表 c where a.科室id=b.部门id and a.科室Id=c.Id and b.人员id = [1]"
            
            strID = UserInfo.ID
        End If
                
    Else
        strSql = "Select 名称,执行间 From 医技执行房间 a, 部门表 b Where a.科室Id=b.Id and  科室ID=[1]"
        
        strID = lngCur科室ID
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strID)
        
    mintcmd影像执行间 = 0
    mstrSelQueueRooms = ""
    mstrAllQueueRooms = ""
    
    If rsData.RecordCount <= 0 Then
        objControl.Caption = "执行间    "
        objControl.Enabled = False
        
        Exit Sub
    End If
    
    If rsData.RecordCount - 1 > UBound(Split(strTemp, ",")) Then strTemp = strTemp & String(rsData.RecordCount - 1 - UBound(Split(strTemp, ",")), ",")
    strTempArry = Split(strTemp, ",")
    
    i = 1
    strTemp = ""
    
    objControl.Enabled = True
    ReDim Preserve mblncmd影像执行间(rsData.RecordCount - 1)

    While rsData.EOF = False
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_影像执行间 + i, Nvl(rsData("执行间")) & "(&" & Chr(64 + i) & ")")
    
        cbrPopControl.ToolTipText = Nvl(rsData!名称) & "-" & Nvl(rsData!执行间)
        cbrPopControl.DescriptionText = Nvl(rsData!执行间) 'Nvl(rsData!名称) & "-" & Nvl(rsData!执行间)
        
        cbrPopControl.Style = xtpButtonIconAndCaption
        cbrPopControl.Checked = False
        cbrPopControl.CloseSubMenuOnClick = False
    
        '排队叫号队列的名称为“科室名称” + “-” + “执行间名称”
        If mstrAllQueueRooms <> "" Then mstrAllQueueRooms = mstrAllQueueRooms & ","
        mstrAllQueueRooms = mstrAllQueueRooms & Nvl(rsData!名称) & "-" & Nvl(rsData!执行间)
                
        If UCase(strTempArry(i - 1)) = UCase("T") Then
            mintcmd影像执行间 = mintcmd影像执行间 + 1
            mblncmd影像执行间(i - 1) = True
            cbrPopControl.Checked = True
            
            strTemp = IIf(strTemp = "", Mid(cbrPopControl.Caption, 1, InStr(cbrPopControl.Caption, "(") - 1), strTemp & "," & Mid(cbrPopControl.Caption, 1, InStr(cbrPopControl.Caption, "(") - 1))
            
            If mstrSelQueueRooms <> "" Then mstrSelQueueRooms = mstrSelQueueRooms & ","
            mstrSelQueueRooms = mstrSelQueueRooms & Nvl(rsData!名称) & "-" & Nvl(rsData!执行间)
        Else
            mblncmd影像执行间(i - 1) = False
        End If
        
        rsData.MoveNext
        i = i + 1
    Wend
    
    '如果没有选中任何执行间，则默认为选择了所有执行间
    If Trim(mstrSelQueueRooms) = "" Then mstrSelQueueRooms = mstrAllQueueRooms
        
    If strTemp <> "" Then
        objControl.ToolTipText = "显示影像执行间为[" & strTemp & "]的检查"
        objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
    Else
        objControl.Caption = "执行间    "
    End If
End Sub

Private Sub InitStudyPlace(objControl As CommandBarControl, cbrPopControl As CommandBarControl, ByVal strStudyTypes As String)
'初始化检查部位配置
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    Dim rsGroup As ADODB.Recordset
    
    Dim strTemp As String
    Dim objTmpControl As CommandBarControl
    
    Dim i As Integer, j As Integer
    
    objControl.CommandBar.Controls.DeleteAll
    
    mstrcmd部位分组 = zlDatabase.GetPara("检查部位过滤", glngSys, mlngModule, "")
    
    If strStudyTypes = "" Then
        strSql = "Select Distinct 类型, substr(分组,instr(分组,'-')+1) as 分组,名称 From 诊疗检查部位 Order By 类型,分组"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Else
        strSql = "Select Distinct 类型, substr(分组,instr(分组,'-')+1) as 分组,名称 " & _
                 "From 诊疗检查部位 A,Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B " & _
                 "Where A.类型 =B.Column_Value Order By 类型,分组"
        Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strStudyTypes)
    End If
    
    If rsData.RecordCount <= 0 Then
        objControl.Caption = "部位    "
        objControl.Enabled = False
        mstrcmd部位分组 = ""
        mstrcmd部位 = ""
        
        Exit Sub
    End If
    
    i = 1
    objControl.Enabled = True
    
    Dim str部位分组 As String
    Dim str部位 As String
    Dim blnIsExist As Boolean
    
    While rsData.EOF = False
        blnIsExist = False
        
        For j = 1 To objControl.CommandBar.Controls.Count
            Set objTmpControl = objControl.CommandBar.Controls(j)
            
            If Not objTmpControl Is Nothing Then
                If objTmpControl.Caption = Nvl(rsData!分组) Then
                    If InStr(objTmpControl.Category, Nvl(rsData!名称)) <= 0 Then
                        objTmpControl.Category = objTmpControl.Category & "|" & Nvl(rsData!名称)
                    End If
                    
                    blnIsExist = True
                    Exit For
                End If
            End If
        Next
        
        If blnIsExist = False Then
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_检查部位 + i, Nvl(rsData!分组))
        
            cbrPopControl.ToolTipText = Nvl(rsData!分组)
            
            If InStr(cbrPopControl.Category, Nvl(rsData!名称)) <= 0 Then
                cbrPopControl.Category = cbrPopControl.Category & "|" & Nvl(rsData!名称)
            End If
            
            cbrPopControl.DescriptionText = Nvl(rsData!分组)
            
            cbrPopControl.Style = xtpButtonIconAndCaption
            cbrPopControl.Checked = False
            cbrPopControl.CloseSubMenuOnClick = False
                    
            If InStr(mstrcmd部位分组, Nvl(rsData!分组)) > 0 Then
                str部位分组 = str部位分组 & "," & Nvl(rsData!分组)
                
                cbrPopControl.Checked = True
                
                strTemp = IIf(strTemp = "", cbrPopControl.Caption, strTemp & "," & cbrPopControl.Caption)
            End If
            
            i = i + 1
        End If
        
        rsData.MoveNext
    Wend
    
    For i = 1 To objControl.CommandBar.Controls.Count
        Set objTmpControl = objControl.CommandBar.Controls(i)

        If Not objTmpControl Is Nothing Then
            If objTmpControl.Checked = True Then
                 str部位 = str部位 & "," & objTmpControl.Category
            End If
        End If
    Next i
    
    mstrcmd部位分组 = str部位分组
    mstrcmd部位 = str部位
    
    If strTemp <> "" Then
        objControl.ToolTipText = "显示检查部位为[" & strTemp & "]的检查"
        objControl.Caption = Mid(strTemp, 1, 6) & IIf(Len(strTemp) > 6, "...   ", "   ")
    Else
        objControl.Caption = "部位    "
    End If
End Sub

Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim blnShowCaption As Boolean
    
    Dim rsCollection As ADODB.Recordset
    Dim rsViewShare As ADODB.Recordset
    Dim rsShareCount As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    
    Dim i As Integer
    Dim i3DFunc As Integer
    Dim intTxtLen As Integer
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    

'菜单定义
'Begin------------------------文件菜单--------------------------------------默认可见
    Me.cbrMain.ActiveMenuBar.Title = "菜单"
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_FilePopup, "文件", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_PrintSet, "打印设置", "", 181, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Excel, "清单打印", "", 103, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Parameter, "参数设置", "", 181, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, ConMenu_File_ShortcutSet, "快捷键设置", "", 181, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Pathol_WorkModule, "站点模式设置", "", 9004, False)
        
        If mblnSetXWParam = True And mlngModule = G_LNG_PACSSTATION_MODULE Then    '有“影像设备目录”的权限，才允许设置新网PACS的参数
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SetXWParam, "PACS参数设置", "", 9004, False)
        End If
        
        '增加视频采集设置菜单
        If mlngModule <> G_LNG_PACSSTATION_MODULE And mblnUseActivexCapture = True Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Cap_DevSet, "视频设置", "视频设置", 815, False)
        End If
        
        If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
            '增加用户交换菜单
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "用户交换", "交换检查医生和报告医生", 3012, True)
        End If
        
        If mlngModule = G_LNG_PATHSTATION_MODULE Then
            '增加用户切换菜单
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "切换用户", "切换用户", 3012, True)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_SendImg, "发送图像", "", 3061, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Change_In, "隐藏列表", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Exit, "退出", "", 191, True)
    End With


'Begin----------------------检查菜单--------------------------------------默认可见
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ManagePopup, "检查", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Manage_RequestPrint, "打印申请单据", "", 0, False)
        
        '如果启用申请单扫描参数 勾选，则加载“扫描申请单”菜单，未勾选则 不加载
        If mSysPar.blnIsPetitionScan Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, comMenu_Petition_Capture, "扫描申请单", "", 3935, False)
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, comMenu_Petition_View, "查看申请单", "查看已扫描的申请单图像", 3935, True)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Regist, "检查登记", "", 2110, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CopyCheck, "复制登记", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Redo, "取消登记", "", 742, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReGet, "召回取消", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ThingModi, "修改信息", "", 0, False)
'        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ModifBaseInfo, "基本信息调整", "", 4113, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Receive, "检查报到", "", 744, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Logout, "取消报到", "", 743, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_InQueue, "入队", "开始排队", 3534, True)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Transfer, "关联影像", "", 505, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Cancel, "取消关联", "", 506, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Review, "附加信息", "", 232, True)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_CheckList, "查看电子申请", "查看电子申请单", 3564, True)
        If mlngModule = G_LNG_PACSSTATION_MODULE Then
            '根据参数判断报告和胶片的发放方式
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_Release, "发放处理", "报告或胶片发放处理", 3013, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportFilmRelease, "发放所有", "", 3013, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "报告发放", "", 8215, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "胶片发放", "", 8216, False)
        Else
            Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "报告发放", "", 8215, False)
        End If
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "报告执行", "报告人", 3967, True)
        
        If mlngModule = G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_SendArrange, "发送安排", "", 232, False)
        End If
        
        '危急值
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_CriticalSituation, "危急值", "", 8338, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Normal, "普通", "", 8344, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Critical, "危急", "", 8345, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CriticalValues, "处理描述", "", 8338, True)
    
        '检查结果
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_Result, "检查结果", "", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "结果阳性", "", 3506, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "结果阴性", "", 3507, False)

        '符合情况
        If mlngModule <> G_LNG_PATHOLSYS_NUM Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_FuHeLevel, "符合情况", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "符合", "", 3587, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "基本符合", "", 3010, False)
                Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "不符合", "", 3010, False)
        End If
            
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlPopup, conMenu_Manage_GChannel, "绿色通道", "", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelOk, "标记", "", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_GChannelCancel, "取消", "", 0, False)
        
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Finish, "无报告完成", "", 216, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_ClearUp, "无报告回退", "", 3012, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Complete, "检查完成", "", 225, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Undone, "取消完成", "", 219, False)

        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_RelatingPatiet, "关联病人", "", 803, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Manage_Burn, "图像刻录", "", 0, True)
        
        If mlngModule <> G_LNG_PACSSTATION_MODULE Then
            Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Tool_Analyse, "高级处理"): cbrControl.ToolTipText = "高级图像处理"
        End If
        
    End With
    
    
    
'Begin-------------------------------------------------------收藏菜单(默认可见)----------------------------------------------------------

    gstrSQL = "select ID ,上级id,创建人,收藏类别 from 影像收藏类别 where 创建人='" & UserInfo.姓名 & "' Start With 上级id Is Null Connect By Prior ID = 上级id"
    Set rsCollection = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)

    gstrSQL = "select ID ,上级id,创建人,收藏类别,是否共享 from 影像收藏类别 where 创建人<>'" & UserInfo.姓名 & "' Start With 上级id Is Null Connect By Prior ID = 上级id"
    Set rsViewShare = zlDatabase.OpenSQLRecord(gstrSQL, GetWindowCaption)
        
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Collection, "收藏", "", 0, False) ' Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Collection, "收藏", -1, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Collection_Manage, "收藏管理", "", 0, False) '.Add(xtpControlButton, conMenu_Collection_Manage, "收藏管理", -1, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Collection_To, "收藏到...", "", 0, False) '.Add(xtpControlButton, conMenu_Collection_To, "收藏到...")
        
        
        '克隆对象 筛选出共享的数据进行判断
        Set rsShareCount = zlDatabase.CopyNewRec(rsViewShare)
        rsShareCount.Filter = "是否共享=1"
        
        If rsShareCount.RecordCount <> 0 Then
           '递归创建共享菜单
           mlngShareFatherID = 0
           Set rsTemp = zlDatabase.CopyNewRec(rsViewShare)
           rsViewShare.Filter = "上级ID=" & Nvl(rsViewShare!上级ID, 1) & " and 创建人<> '" & UserInfo.姓名 & "'"
           
           Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Collection_ViewShare, "共享查看", "", 0, False)
           Call RecursionCreateShareMenu(rsViewShare, rsTemp, cbrControl)
        End If

        If rsCollection.RecordCount > 0 Then
            '递归创建收藏类别菜单
                 mlngCollectionFatherID = 0
                 Set rsTemp = zlDatabase.CopyNewRec(rsCollection)
                 rsCollection.Filter = "上级ID=" & Nvl(rsCollection!上级ID, 1)
                 Call RecursionCreateCollectionMenu(rsCollection, rsTemp, cbrMenuBar)
        End If
        
    End With
    
    '读取发布到该模块的报表(不含虚拟模块的)
'-----------------------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)")
    cbrMenuBar.ID = conMenu_ReportPopup
    
    Call zlDatabase.ShowReportMenu(cbrMain, glngSys, mlngModule, mstrPrivs, _
                                    "ZL1_INSIDE_1294_01", _
                                    "ZL1_INSIDE_1294_02", _
                                    "ZL1_INSIDE_1294_03", _
                                    "ZL1_INSIDE_1294_04", _
                                    "ZL1_INSIDE_1294_05", _
                                    "ZL1_INSIDE_1294_06", _
                                    "ZL1_INSIDE_1294_07", _
                                    "ZL1_INSIDE_1294_08", _
                                    "ZL1_INSIDE_1294_09", _
                                    "ZL1_INSIDE_1294_10", _
                                    "ZL1_INSIDE_1294_11", _
                                    "ZL1_INSIDE_1294_12", _
                                    "ZL1_INSIDE_1294_13")
    
    If cbrMenuBar.CommandBar.Controls.Count > 0 Then
        cbrMenuBar.Category = M_STR_MODULE_MENU_TAG
        
        For i = 1 To cbrMenuBar.CommandBar.Controls.Count
            cbrMenuBar.CommandBar.Controls(i).Category = M_STR_MODULE_MENU_TAG
        Next i
    Else
        cbrMenuBar.Delete
    End If
    
'Begin----------------------自定义查询菜单--------------------------------------默认可见
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_Manage_Query, "查询", "", 0, False)
    Call RefreshCustomQueryMenu(cbrMenuBar, mlngCur科室ID)
    
    
'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_ViewPopup, "查看", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_ToolBar_Size, "大图标", "", 0, False): cbrPopControl.Checked = True
            End With
            
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_FontSize, "字体大小", "", 0, False)
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar '二级菜单
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_S, "小字体", "", 0, False): cbrPopControl.Checked = True
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_M, "中字体", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_FontSize_L, "大字体", "", 0, False)
            End With
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_StatusBar, "状态栏", "", 0, True): cbrControl.Checked = True
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_View_Filter * 10#, "检查科室", "", 0, False)
'        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Filter, "快速过滤", "", 0, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_View_Refresh, "刷新", "", 0, False)
    End With


'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_HelpPopup, "帮助", "", 0, False)
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Help, "帮助主题", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联", "", 0, False)
            With cbrControl.CommandBar
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Forum, "中联论坛", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Home, "中联主页", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈", "", 0, False)
            End With
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Help_About, "关于…", "", 0, True)
    End With
    

'---------------------设置右上角当前科室----------------------------------
    Set cbrControl = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlPopup, conMenu_View_Filter * 10#, "检查科室", "", 0, False): cbrControl.flags = xtpFlagRightAlign
            
    '最右边显示浮动采集按钮
    If mblnUseActivexCapture Then
        Set cbrControl = CreateModuleMenu(cbrMain.ActiveMenuBar.Controls, xtpControlButton, comMenu_Cap_Process, "浮动采集", "弹出独立采集窗口", 0, False): cbrControl.flags = xtpFlagRightAlign
    End If
        
'---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True

    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Regist, "登记", "检查登记", 211, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Receive, "报到", "检查报到", 744, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Logout, "取消", "取消报到", 743, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_InQueue, "入队", "开始排队", 3534, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, comMenu_Petition_View, "查看申请单", "查看已扫描的申请单图像", 3935, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_CheckList, "查看电子申请", "查看电子申请单", 3564, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Review, "备注", "附加信息", 232, True)
    
    If mlngModule <> G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Tool_Analyse, "高级"): cbrControl.ToolTipText = "高级图像处理"
    End If
    
    '只有影像病理系统才具有用户切换功能
    If mlngModule = G_LNG_PATHSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SwitchUser, "切换", "切换用户", 3012, False, conMenu_Tool_Analyse)
    End If
        
    '根据参数判断报告和胶片的发放方式
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Release, "发放处理", "报告或胶片发放处理", 3013, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportFilmRelease, "发放所有", "发放所有", 3013, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "报告发放", "报告发放", 8215, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FilmRelease, "胶片发放", "胶片发放", 8216, False)
    Else
        Set cbrPopControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportRelease, "报告发放", "报告发放", 8215, False)
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ReportExecutor, "报告执行", "报告人", 3967, False)
    
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_SendArrange, "发送安排", "发送安排", 232, False)
    End If
    
    '危急情况
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_CriticalSituation, "危急值", "危急情况", 8338, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Normal, "普通", "普通", 8344, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Critical, "危急", "危急", 8345, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_CriticalValues, "处理描述", "", 8338, True)
    
    '检查结果阴阳性
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_Result, "结果", "检查结果阴阳性", 3506, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Negative, "阳性", "阳性", 3506, False)
        Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_Positive, "阴性", "阴性", 3507, False)
    
    '如果是病理系统，则没有符合情况按钮
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlSplitButtonPopup, conMenu_Manage_FuHeLevel, "符合情况", "符合情况", 8044, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_FuHe, "符合", "符合", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_JiBenFuHe, "基本符合", "基本符合", 0, False)
            Set cbrPopControl = CreateModuleMenu(cbrControl.CommandBar.Controls, xtpControlButton, conMenu_Manage_BuFuHe, "不符合", "不符合", 0, False)
    End If
        
    '只有影像采集系统才具有用户交换功能
    If mlngModule = G_LNG_VIDEOSTATION_MODULE Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_ChangeUser, "交换", "交换检查医生和报告医生", 3012, False)
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Complete, "完成", "检查最终完成", 225, False)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Filter, "过滤", "过滤", 0, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_View_Refresh, "刷新", "刷新", 0, False)
  
  
     '初始化界面字体 加到这里为了防止在一些特殊操作的时候，会导致字体恢复成初始化
    Call SetFontSize(IIf(mbytFontSize = 12, 1, IIf(mbytFontSize = 15, 2, 0)))
    
'    '创建工作模块所需的菜单
'    Call CreateWorkModuleMenu
End Sub


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'创建该模块内的菜单
    
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If

    CreateModuleMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG
End Function


Private Sub CreateWorkModuleMenu()
'创建工作模块菜单
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
On Error GoTo err
    If Not mobjWork_Pathol Is Nothing And mblnIsLoadPatholModule Then
        Call mobjWork_Pathol.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mobjWork_Pathol.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    '创建影像图像模块相关菜单及工具栏
    If Not mfrmWork_PacsImg Is Nothing And InStr(mstrWorkModule, ";影像图像模块;") > 0 Then
        Call mfrmWork_PacsImg.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mfrmWork_PacsImg.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    If mblnUseActivexCapture Then
        '使用ActivexExe的图像采集菜单处理
        If Not mobjWork_ActiveVideo Is Nothing And InStr(mstrWorkModule, ";影像采集模块;") > 0 Then
            'TODO:创建视频采集模块菜单
'            Call mobjWork_ActiveVideo.zlMenu.zlCreateMenu(Me.cbrMain)
'            Call mobjWork_ActiveVideo.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
        End If
    End If

    
    '必须将报告菜单的创建放在mobjWork_His创建菜单之前，否则在切换到其他模块时，对应的模块菜单不能够被创建
    If Not mobjWork_Report Is Nothing And _
        (InStr(mstrWorkModule, ";影像报告模块;") > 0 Or InStr(mstrWorkModule, ";病理诊断模块;") > 0) Then
        Call mobjWork_Report.zlMenu.zlCreateMenu(Me.cbrMain)
        Call mobjWork_Report.zlMenu.zlCreateToolBar(Me.cbrMain.Item(2))
    End If
    
    If Not mobjWork_His Is Nothing Then
        '因为在PACS系统中 “打印” 菜单项在编辑菜单组下，而病历中在文件菜单下，所以在调用病历的菜单创建过程时，
        '在文件菜单下找不到打印菜单项而报错，而PACS中，清单打印在文件菜单下，所以调用病历的菜单创建过程时将
        '清单打印的id改成打印的id，创建完后，恢复清单打印原来的id
        If TabWindow.Selected.tag = "门诊电子病历" Or TabWindow.Selected.tag = "住院电子病历" Then
            Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
            Set cbrControl = cbrMenuBar.CommandBar.Controls.Find(, conMenu_File_Excel)
            cbrControl.ID = conMenu_File_Print
        End If
        
        Call mobjWork_His.zlMenu.zlCreateMenu(Me.cbrMain)
        
        If TabWindow.Selected.tag = "门诊电子病历" Or TabWindow.Selected.tag = "住院电子病历" Then
            cbrControl.ID = conMenu_File_Excel
        End If
    End If

    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call cbrMain.RecalcLayout
    Exit Sub
err:
    cbrControl.ID = conMenu_File_Excel
End Sub

Private Sub RecursionCreateShareMenu(rsFilterADO As ADODB.Recordset, rsFullADO As ADODB.Recordset, cbrParentControl As CommandBarControl, Optional blnIsShare As Boolean = False)
'递归循环创建共享菜单
    Dim rsFilterTemp As ADODB.Recordset
    Dim i As Long
    Dim cbrControl As CommandBarControl
    Static j As Long
    
    If rsFilterADO.RecordCount = 0 Then Exit Sub
    rsFilterADO.MoveFirst
    
    With cbrParentControl.CommandBar.Controls
        If mlngShareFatherID <> 0 Then
            Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + mlngShareFatherID, "查看当前收藏", -1, False)
            cbrControl.Category = M_STR_MODULE_MENU_TAG
        End If
        
        For i = 1 To rsFilterADO.RecordCount
            rsFullADO.Filter = " 上级ID=" & Nvl(rsFilterADO!ID)

            If rsFullADO.RecordCount > 0 Then
                Set cbrControl = Nothing
  
                If Nvl(rsFilterADO!是否共享) = 1 Or blnIsShare = True Then
                    mlngShareFatherID = Nvl(rsFilterADO!ID)
                    '创建父级菜单 如果上级ID=1 则显示共享人姓名
                    Set cbrControl = .Add(xtpControlButtonPopup, CLng(conMenu_Collection_ViewShare) * 10000# + j, Nvl(rsFilterADO!收藏类别) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & Nvl(rsFilterADO!创建人) & ")", ""), -1, False)
                    cbrControl.DescriptionText = Nvl(rsFilterADO!创建人)
                    cbrControl.Category = M_STR_MODULE_MENU_TAG
                    
                    j = j + 1
                    If i = 1 Then cbrControl.BeginGroup = True
                End If
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '调用自己
                Call RecursionCreateShareMenu(rsFilterTemp, rsFullADO, IIf(cbrControl Is Nothing, cbrParentControl, cbrControl), IIf(Nvl(rsFilterADO!是否共享) = 0, False, True))
            Else
            '创建子级菜单
                If Nvl(rsFilterADO!是否共享) = 1 Or blnIsShare = True Then
                    Set cbrControl = .Add(xtpControlButton, CLng(conMenu_Collection_ViewShare) * 10000# + j, Nvl(rsFilterADO!收藏类别) & Decode(cbrParentControl.ID, conMenu_Collection_ViewShare, "(" & Nvl(rsFilterADO!创建人) & ")", ""), -1, False)
                    cbrControl.DescriptionText = Nvl(rsFilterADO!创建人)
                    cbrControl.Category = M_STR_MODULE_MENU_TAG
                    
                    j = j + 1
                    If i = 1 Then cbrControl.BeginGroup = True
                End If
                mlngShareFatherID = 0
            End If

            If Not rsFilterADO.EOF Then rsFilterADO.MoveNext
        Next
    End With
End Sub



Private Sub RecursionCreateCollectionMenu(rsFilterADO As ADODB.Recordset, rsFullADO As ADODB.Recordset, cbrMenuBar As CommandBarPopup)
'递归循环创建收藏类别菜单
    Dim rsFilterTemp As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim i As Long
    Static j As Long

    If rsFilterADO.RecordCount = 0 Then Exit Sub
    rsFilterADO.MoveFirst

    With cbrMenuBar.CommandBar.Controls
        If mlngCollectionFatherID <> 0 Then
            Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + mlngCollectionFatherID, "查看当前收藏", -1, False)
            cbrControl.Category = M_STR_MODULE_MENU_TAG
        End If

        For i = 1 To rsFilterADO.RecordCount

            rsFullADO.Filter = " 上级ID=" & Nvl(rsFilterADO!ID)
            mlngCollectionFatherID = Nvl(rsFilterADO!ID)
            If rsFullADO.RecordCount > 0 Then
            '创建父级菜单
                Set cbrControl = .Add(xtpControlButtonPopup, CLng(comMenu_Collection_Type) * 10000# + j, Nvl(rsFilterADO!收藏类别), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
                
                Set rsFilterTemp = zlDatabase.CopyNewRec(rsFullADO)
                '调用自己
                Call RecursionCreateCollectionMenu(rsFilterTemp, rsFullADO, cbrControl)
                
            Else
            '创建子级菜单
                Set cbrControl = .Add(xtpControlButton, CLng(comMenu_Collection_Type) * 10000# + j, Nvl(rsFilterADO!收藏类别), -1, False)
                cbrControl.Category = M_STR_MODULE_MENU_TAG
                
                j = j + 1
            End If
            If i = 1 Then cbrControl.BeginGroup = True

            If Not rsFilterADO.EOF Then rsFilterADO.MoveNext
        Next
    End With

End Sub


Private Sub ReadWorkModuleCfg()
    '设置当前需要创建的工作页面
    mstrWorkModule = zlDatabase.GetPara("站点模块", glngSys, mlngModule, "")
    mstrWorkModule = IIf(mstrWorkModule <> "", ";" & mstrWorkModule & ";", "")
    
    '如果模块为空，并且启用了排队叫号，则只显示排队叫号工作模块
    If mstrWorkModule = "" Then 'And Not mblnUseQueue
        Select Case mlngModule
            Case G_LNG_PACSSTATION_MODULE
                mstrWorkModule = ";影像图像模块;影像报告模块;病历记录模块;费用记录模块;医嘱记录模块;"
            
            Case G_LNG_VIDEOSTATION_MODULE
                mstrWorkModule = ";影像采集模块;影像报告模块;病历记录模块;费用记录模块;医嘱记录模块;"
            
            Case G_LNG_PATHOLSYS_NUM
                mstrWorkModule = ";标本核收模块;影像采集模块;病理取材模块;病理制片模块;病理特检模块;过程报告模块;病理诊断模块;病历记录模块;费用记录模块;医嘱记录模块;"
            Case Else
                Exit Sub
        End Select
    End If
    
'    '测试代码
'    mstrWorkModule = ";影像图像模块;影像采集模块;标本核收模块;病理取材模块;病理制片模块;病理特检模块;过程报告模块;影像报告模块;费用记录模块;医嘱记录模块;病历记录模块;"
End Sub


Private Sub InitPatholModuleObj()
    '初始化病理相关模块对象
    If mobjWork_Pathol Is Nothing Then
        Set mobjWork_Pathol = New clsWorkModule_Pathol
        Call mobjWork_Pathol.zlInitModule(mlngModule, mstrPrivs, mlngCur科室ID, Me)
    End If
End Sub

Private Sub InitHisModuleObj()
    '初始化HIS相关模块对象
    If mobjWork_His Is Nothing Then
        Set mobjWork_His = New clsWorkModule_His
        
        If mblnAllDepts Then
            Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, UserInfo.部门ID, Me)
        Else
            Call mobjWork_His.zlInitModule(mlngModule, mstrPrivs, mlngCur科室ID, Me)
        End If
    End If
End Sub

Private Sub InitActiveVideoModuleObj()
'初始化ActivexExe视频采集模块对象
    If mlngModule = G_LNG_PACSSTATION_MODULE Then Exit Sub
    If Not CheckPopedom(mstrPrivs, "视频采集") Then Exit Sub
    If InStr(mstrWorkModule, ";影像采集模块;") < 0 Then Exit Sub
    
    If mobjWork_ActiveVideo Is Nothing Then
        Set mobjWork_ActiveVideo = CreateObject("zl9PacsCapture.clsPacsCapture") ' New zl9PacsCapture.clsPacsCapture
        
        mobjWork_ActiveVideo.ParentWindowKey = Me.Name
        mobjWork_ActiveVideo.AllowEventNotify = True
        
        Call mobjWork_ActiveVideo.RegEventObj(Me)
        
        Call mobjWork_ActiveVideo.zlInitModule(gcnOracle, glngSys, mlngModule, mstrPrivs, mlngCur科室ID, Me.hWnd, Me, True, gblnUseDebugLog)
    End If
End Sub

Private Sub ShowModuleLoadState(Optional ByVal strState As String = "")
'显示载入状态
On Error GoTo ErrHandle
    picLoadState.Left = 0
    picLoadState.Top = 350
    picLoadState.Width = picWindow.Width - 0
    picLoadState.Height = picWindow.Height - 350
    
    
    If strState <> "" Then
        labLoadState.Caption = strState
        Call picLoadState_Resize
    End If
    
    picLoadState.Visible = True
    
ErrHandle:
End Sub

Private Sub HideModuleLoadState()
'隐藏载入状态
    picLoadState.Visible = False
End Sub

Public Sub InitSubForm()
    Dim i As Integer
    Dim blnDoEvents As Boolean

    mblnIsLoadPatholModule = False   '当该变量最后仍然为false时，则根据条件删除病理菜单
    blnDoEvents = True  '当值为true时，则屏蔽工作模块加载过程中的事件处理
    
    Call ShowModuleLoadState
    DoEvents
    
    With TabWindow
        .RemoveAll
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .PaintManager.ColorSet.ButtonSelected = &HFFC0C0
        .PaintManager.ColorSet.ButtonNormal = &HE0E0E0
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ButtonMargin.Top = 4
        .PaintManager.ButtonMargin.Bottom = 4
        .PaintManager.ShowIcons = True
        .RemoveAll
        
        '读取工作模块配置
        Call ReadWorkModuleCfg
    
        If InStr(mstrWorkModule, ";影像图像模块;") > 0 Then
            '创建影像记录模块
            If mfrmWork_PacsImg Is Nothing Then
                Set mfrmWork_PacsImg = New frmWork_Image
                
                Set mfrmWork_PacsImg.PacsCore = mobjPacsCore
                Call mfrmWork_PacsImg.zlInitModule(mlngModule, mstrPrivs, mlngCur科室ID, Me)
            End If
    
            .InsertItem 0, "影像记录", picTemp.hWnd, conMenu_Img_Look
            .Item(TabWindow.ItemCount - 1).tag = "影像图象"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
            
        Else
            '删除对应菜单和工具栏
            If Not mfrmWork_PacsImg Is Nothing Then
                Call mfrmWork_PacsImg.zlMenu.zlClearMenu
                Call mfrmWork_PacsImg.zlMenu.zlClearToolBar
            End If
        End If
                        
        If mlngModule <> G_LNG_PACSSTATION_MODULE And CheckPopedom(mstrPrivs, "视频采集") _
            And InStr(mstrWorkModule, ";影像采集模块;") > 0 Then
            
            If mobjCaptureHot Is Nothing Then
                Set mobjCaptureHot = New zl9PacsControl.clsHookKey
                Call mobjCaptureHot.EnableHook(WM_KEYDOWN, True)
            End If

            If mblnUseActivexCapture Then
                Call InitActiveVideoModuleObj
                
                .InsertItem 1, "影像采集", mobjWork_ActiveVideo.ContainerHwnd, conMenu_Cap_Dynamic
                .Item(TabWindow.ItemCount - 1).tag = "影像采集"
            End If


            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            If mblnUseActivexCapture Then
                'TODO:使用activex视频采集方式后的菜单加载...
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "标本核收") And InStr(mstrWorkModule, ";标本核收模块;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 2, "标本核收", picTemp.hWnd, G_INT_ICONID_SPECIMEN
            .Item(TabWindow.ItemCount - 1).tag = "标本核收"
            
            mblnIsLoadPatholModule = True

            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "病理取材") And InStr(mstrWorkModule, ";病理取材模块;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 3, "病理取材", picTemp.hWnd, G_INT_ICONID_MATERIAL
            .Item(TabWindow.ItemCount - 1).tag = "病理取材"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If CheckPopedom(mstrPrivs, "病理制片") And InStr(mstrWorkModule, ";病理制片模块;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 4, "病理制片", picTemp.hWnd, G_INT_ICONID_SLICES
            .Item(TabWindow.ItemCount - 1).tag = "病理制片"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If (CheckPopedom(mstrPrivs, "免疫组化") Or CheckPopedom(mstrPrivs, "特殊染色") Or CheckPopedom(mstrPrivs, "分子病理")) _
            And InStr(mstrWorkModule, ";病理特检模块;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 5, "病理特检", picTemp.hWnd, G_INT_ICONID_SPEEXAM
            .Item(TabWindow.ItemCount - 1).tag = "病理特检"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If (CheckPopedom(mstrPrivs, "冰冻报告") Or CheckPopedom(mstrPrivs, "特染报告") _
            Or CheckPopedom(mstrPrivs, "分子报告") Or CheckPopedom(mstrPrivs, "免疫报告") _
            Or CheckPopedom(mstrPrivs, "冰冻特检报告查阅")) And InStr(mstrWorkModule, ";过程报告模块;") > 0 Then
            Call InitPatholModuleObj
            
            .InsertItem 6, "冰冻/特检报告", picTemp.hWnd, G_INT_ICONID_PROREPORT
            .Item(TabWindow.ItemCount - 1).tag = "过程报告"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
        
        If GetInsidePrivs(p诊疗报告管理, True) <> "" And _
            (InStr(mstrWorkModule, ";影像报告模块;") > 0 Or InStr(mstrWorkModule, ";病理诊断模块;") > 0) Then
            
            If mobjWork_Report Is Nothing Then
                Set mobjWork_Report = New clsWorkModule_Report
                
                Call mobjWork_Report.zlInitModule(mlngModule, mstrPrivs, mlngCur科室ID, Me)
                
                Set mobjWork_Report.PacsCore = mobjPacsCore
            End If

            .InsertItem 7, "影像报告", picReportContainer.hWnd, 10008 'conMenu_Edit_Compend
            .Item(TabWindow.ItemCount - 1).tag = "报告填写"
            
            mblnIsLoadPatholModule = True
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            '删除对应菜单和工具栏
            If Not mobjWork_Report Is Nothing Then
                Call mobjWork_Report.zlMenu.zlClearMenu
                Call mobjWork_Report.zlMenu.zlClearToolBar
            End If
        End If
        
        
        If Not mblnIsLoadPatholModule And Not mobjWork_Pathol Is Nothing Then
            '没有加载病理模块，且mobjWork_Pathol不为空时，删除病理菜单
            Call mobjWork_Pathol.zlMenu.zlClearMenu
            Call mobjWork_Pathol.zlMenu.zlClearToolBar
        End If
        
        
        If GetInsidePrivs(p医嘱附费管理, True) <> "" And InStr(mstrWorkModule, ";费用记录模块;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 8, "费用记录", picTemp.hWnd, 10007
            .Item(TabWindow.ItemCount - 1).tag = "申请费用"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            '删除对应菜单和工具栏
            If Not mobjWork_His Is Nothing Then
                '暂不处理，his模块的菜单只能在该模块被显示的情况下被创建...
            End If
        End If
        
        If GetInsidePrivs(p住院医嘱下达, True) <> "" And InStr(mstrWorkModule, ";医嘱记录模块;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 9, "医嘱记录", picTemp.hWnd, 10010
            .Item(TabWindow.ItemCount - 1).tag = "住院医嘱"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            '删除对应菜单和工具栏
            If Not mobjWork_His Is Nothing Then
                '暂不处理，his模块的菜单只能在该模块被显示的情况下被创建...
            End If
        End If
        
        If GetInsidePrivs(p门诊医嘱下达, True) <> "" And InStr(mstrWorkModule, ";医嘱记录模块;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 10, "医嘱记录", picTemp.hWnd, 10010  ' conMenu_Edit_NewItem
            .Item(TabWindow.ItemCount - 1).tag = "门诊医嘱": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            '删除对应菜单和工具栏
            If Not mobjWork_His Is Nothing Then
                '暂不处理，his模块的菜单只能在该模块被显示的情况下被创建...
            End If
        End If
        
        If GetInsidePrivs(p住院病历管理, True) <> "" And InStr(mstrWorkModule, ";病历记录模块;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 11, "病历记录", picTemp.hWnd, 10009 ' conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).tag = "住院病历"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            '删除对应菜单和工具栏
            If Not mobjWork_His Is Nothing Then
                '暂不处理，his模块的菜单只能在该模块被显示的情况下被创建...
            End If
        End If
        
        If GetInsidePrivs(p门诊病历管理, True) <> "" And InStr(mstrWorkModule, ";病历记录模块;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 12, "病历记录", picTemp.hWnd, 10009 ' conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).tag = "门诊病历": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            '删除对应菜单和工具栏
            If Not mobjWork_His Is Nothing Then
                '暂不处理，his模块的菜单只能在该模块被显示的情况下被创建...
            End If
        End If
        
        If GetInsidePrivs(p门诊电子病历, True) <> "" And InStr(mstrWorkModule, ";电子病历模块;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 13, "电子病历", picTemp.hWnd, 10009
            .Item(TabWindow.ItemCount - 1).tag = "门诊电子病历": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            '删除对应菜单和工具栏
            If Not mobjWork_His Is Nothing Then
                '暂不处理，his模块的菜单只能在该模块被显示的情况下被创建...
            End If
        End If
        
        If GetInsidePrivs(p住院电子病历, True) <> "" And InStr(mstrWorkModule, ";电子病历模块;") > 0 Then
            Call InitHisModuleObj
            
            .InsertItem 14, "电子病历", picTemp.hWnd, 10009
            .Item(TabWindow.ItemCount - 1).tag = "住院电子病历": .Item(TabWindow.ItemCount - 1).Visible = False
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        Else
            '删除对应菜单和工具栏
            If Not mobjWork_His Is Nothing Then
                '暂不处理，his模块的菜单只能在该模块被显示的情况下被创建...
            End If
        End If
        
        '添加排队叫号页面
        If mSysPar.blnUseQueue = True Then
            mstrWorkModule = mstrWorkModule & ";排队叫号模块;"
            
            If mobjQueue Is Nothing Then
                Set mobjQueue = New frmWork_Queue
                Call mobjQueue.zlInitPacsQueueCfg(mlngModule, mlngCur科室ID, NeedName(mstrCur科室), mstrPrivs)
            End If
            
            .InsertItem 15, "排队叫号", picTemp.hWnd, 10011
            .Item(TabWindow.ItemCount - 1).tag = "排队叫号"
            
            If Not blnDoEvents Then
                DoEvents
                blnDoEvents = True
            End If
        End If
    
'        If Not GetVideoForm Is Nothing Then Call GetVideoForm.ShowVideoWindow(picVideoContainer)
    End With
    
    DoEvents
    
    If GetWorkModuleCount = 1 Then
        TabWindow.PaintManager.ClientMargin.Top = -30
    Else
        TabWindow.PaintManager.ClientMargin.Top = 0
    End If
    
    Call HideModuleLoadState
End Sub

Private Function GetWorkModuleCount() As Long
'获取可见tabwindow的数量
    Dim i As Long
    Dim lngCount As Long
    Dim aryWorkModule() As String
    
    
    aryWorkModule = Split(mstrWorkModule, ";")
    
    For i = LBound(aryWorkModule) To UBound(aryWorkModule)
        If aryWorkModule(i) <> "" Then lngCount = lngCount + 1
    Next i
    
    GetWorkModuleCount = lngCount
End Function


Private Function GetTabWindowIndex() As Long
'获取第一个可见tabwindow的索引
    Dim i As Long
    
    GetTabWindowIndex = -1
    For i = 0 To TabWindow.ItemCount - 1
        If TabWindow.Item(i).Visible Then
            GetTabWindowIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub mobjWork_Report_AfterDeleted(ByVal lngOrderID As Long)
    Call AfterDeleted(lngOrderID)
End Sub

Private Sub mobjWork_Report_AfterDeletedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Call AfterDeletedRich(lngOrderID, strDocId)
End Sub

Private Sub mobjWork_Report_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mobjWork_Report_AfterPrintedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Call AfterPrintedRich(lngOrderID, strDocId)
End Sub

Private Sub mobjWork_Report_AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Object, ByVal lngSaveType As Long)
    Call AfterReportSaved(lngOrderID, frmOwnerForm, lngSaveType)
End Sub

Private Sub mobjWork_Report_AfterSavedRich(ByVal lngOrderID As Long, ByVal strDocId As String, frmOwnerForm As Object, ByVal lngSaveType As Long)
    Call AfterReportSavedRich(lngOrderID, strDocId, frmOwnerForm, lngSaveType)
End Sub

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    On Error GoTo ErrHandle
    
    If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.RefreshReportImage
    
    Exit Sub
    
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjQueue_OnDiagnose(ByVal lngAdviceId As Long, ByVal strExeRoom As String, ByVal strTurnPage As String)
'排队接诊事件
On Error GoTo ErrHandle
    Dim lngIndex As String
    Dim i As Long
    
    lngIndex = ufgStudyList.FindRowIndex(lngAdviceId, "医嘱ID", True)
    
    If lngIndex > 0 Then
        Call ufgStudyList.LocateRow(lngIndex)
        
        If Trim(strTurnPage) <> "" Then
            '跳转到指定的工作模块

            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow(i).tag, strTurnPage) > 0 And TabWindow(i).Visible Then
                    TabWindow(i).Selected = True
                    Exit For
                End If
            Next i
        End If
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub mobjQueue_OnCompleted(ByVal lngAdviceId As Long, ByVal strExeRoom As String)
'排队完成事件
On Error GoTo ErrHandle
    Dim lngIndex As String

    
    lngIndex = ufgStudyList.FindRowIndex(lngAdviceId, "医嘱ID", True)
    
    If lngIndex > 0 Then
        Call ufgStudyList.LocateRow(lngIndex)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjQueue_OnSelChange(ByVal lngAdviceId As Long)
'行选择改变事件
On Error GoTo ErrHandle
    Dim lngIndex As Long
    
    If mSysPar.blnSynStudylist Then
        lngIndex = ufgStudyList.FindRowIndex(lngAdviceId, "医嘱ID", True)
        
        If lngIndex > 0 Then
            Call ufgStudyList.LocateRow(lngIndex)
        End If
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub AfterDeletedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
    Dim intState As Integer
    Dim lngSendNo As Long
    
On Error GoTo ErrHandle
    intState = getStudyStateRich(lngOrderID, strDocId, False, , lngSendNo)
    If intState = 6 Then Exit Sub
    
    gstrSQL = "Zl_影像检查_状态更新(" & lngOrderID & "," & lngSendNo & ",''," & intState & ",0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存检查状态及报告人")
    
    If intState < 4 Then
        gstrSQL = "ZL_影像报告标记_Clear(" & lngOrderID & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "清空标记"
    End If
    
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub AfterDeleted(ByVal lngOrderID As Long)
On Error GoTo ErrHandle
    gstrSQL = "ZL_影像报告标记_Clear(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "清空标记"
    
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub AfterPrintedRich(ByVal lngOrderID As Long, ByVal strDocId As String)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strResultInput As String
    Dim bln保存结果阳性 As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim blnAllReportFinished As Boolean
    Dim intState As Integer, lngSendId As Long
    
    strResultInput = ""
    
    intState = getStudyStateRich(lngOrderID, strDocId, False, blnAllReportFinished, lngSendId, bln保存结果阳性, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
    If intState = 6 Then Exit Sub
    
    BUGEX "AfterPrintedRich lngOrderID=" & lngOrderID & "   strDocId=" & strDocId
    
    strSql = "Select B.危急状态, A.结果阳性, B.影像质量, A.报告质量, B.符合情况,B.医嘱ID " & _
                 "From 影像报告记录 A, 影像检查记录 B " & _
                 "Where A.ID=[1] and A.医嘱id = B.医嘱id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取结果阳性", strDocId)
    
    If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "危急状态|"
    If (Not bln保存结果阳性 And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "结果阳性|"
    If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And InStr(mstrPrivs, "影像质控") > 0 Then strResultInput = strResultInput & "影像质量|"
    If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And InStr(mstrPrivs, "报告质控") > 0 Then strResultInput = strResultInput & "报告质量|"
    If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "符合情况|"
    
    If strResultInput <> "" Then Call PromptResultRich(lngOrderID, strDocId, mlngModule, Me, mlngCur科室ID, strResultInput)
    
    If mSysPar.blnPrintCommit = True Then
        If blnAllReportFinished Then    '所有报告完成后才算完成
            Call Menu_Manage_检查最终完成(lngOrderID, False, strDocId)
        End If
    End If
    
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub AfterPrinted(lngOrderID As Long)
On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strResultInput As String
    
    strResultInput = ""
    gstrSQL = "ZL_影像报告打印_Update(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, "更新打印标记"
    
    strSql = "Select B.危急状态, A.结果阳性, B.影像质量, B.报告质量, B.符合情况 " & _
             "From 病人医嘱发送 A, 影像检查记录 B " & _
             "Where A.医嘱id = B.医嘱id and B.医嘱ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取结果阳性", lngOrderID)
    
    If IsNull(rsTemp!危急状态) And mSysPar.lngCriticalValues <> 0 Then strResultInput = "危急状态|"
    If IsNull(rsTemp!结果阳性) And Not mSysPar.blnIgnoreResult Then strResultInput = strResultInput & "结果阳性|"
    If IsNull(rsTemp!影像质量) And mSysPar.strImageLevel <> "" And mSysPar.lngImageLevel <> 0 And InStr(mstrPrivs, "影像质控") > 0 Then strResultInput = strResultInput & "影像质量|"
    If IsNull(rsTemp!报告质量) And mSysPar.strReportLevel <> "" And mSysPar.lngReportLevel <> 0 And InStr(mstrPrivs, "报告质控") > 0 Then strResultInput = strResultInput & "报告质量|"
    If IsNull(rsTemp!符合情况) And mSysPar.lngConformDetermine <> 0 Then strResultInput = strResultInput & "符合情况|"

    If strResultInput <> "" Then Call PromptResult(lngOrderID, mlngModule, Me, mlngCur科室ID, strResultInput)
    
    If mSysPar.blnPrintCommit = True Then
        Call Menu_Manage_检查最终完成(lngOrderID, False)
    End If
    
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Public Sub AfterReportSavedRich(ByVal lngOrderID As Long, ByVal strDocId As String, frmOwnerForm As Form, ByVal lngSaveType As Long)
'保存报告之后的处理
'执行过程：2-已报到；3-已检查；4-已报告；5-已审核；6-已完成
On Error GoTo ErrHandle
    Dim intState As Integer, lngSendId As Long
    Dim str签名 As String
    Dim str创建人 As String
    Dim str保存人 As String
    Dim bln保存结果阳性 As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    Dim blnAllReportFinished As Boolean
    
    arrSQL = Array()

    Call mobjWork_Report.zlRefreshFace(True)
    
    'intState =1--已登记；2--已报到；3--已检查；4--已报告；5--已审核；6--已完成
    BUGEX "AfterReportSavedRich lngOrderID=" & lngOrderID & "   strDocId=" & strDocId
    
    '获取本次检查的执行过程
    intState = getStudyStateRich(lngOrderID, strDocId, False, blnAllReportFinished, lngSendId, bln保存结果阳性, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
    If intState = 6 Then Exit Sub
    
    If intState = 2 Or intState = 3 Then
        '报告保存时执行费用
        If mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngMoneyExeModle = 2 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            
            gstrSQL = "Zl_影像费用执行(" & lngOrderID & "," & lngSendId & ",4,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    End If
    
    gstrSQL = "Zl_影像检查_状态更新(" & lngOrderID & "," & lngSendId & ",'" & strDocId & "'," & intState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
                    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = gstrSQL
    
    gcnOracle.BeginTrans        '----------保存检查状态及报告人
    
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存检查状态及报告人")
    Next i
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If (intState = 4 Or intState = 5) And IIf(mSysPar.lngHintType = 0, lngSaveType = 1, IIf(mSysPar.lngHintType = 1, lngSaveType = 2, False)) Then
        Dim strResultInput As String
        
        strResultInput = ""
        If mSysPar.blnReportWithResult Then '无影像诊断为阴性  -无提示自动标记
            Call mobjWork_Report.Menu_Manage_标记阴阳(mListAdviceInf.lngAdviceId, "0")
        End If
            
        If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "危急状态|"
        If (Not bln保存结果阳性 And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "结果阳性|"
        If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And InStr(mstrPrivs, "影像质控") > 0 Then strResultInput = strResultInput & "影像质量|"
        If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And InStr(mstrPrivs, "报告质控") > 0 Then strResultInput = strResultInput & "报告质量|"
        If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "符合情况|"
 
        If strResultInput <> "" Then Call PromptResultRich(lngOrderID, strDocId, mlngModule, frmOwnerForm, mlngCur科室ID, strResultInput)
    End If
    
    If intState = 5 And mSysPar.blnCompleteCommit Then   '如果“审核后直接完成”
        If blnAllReportFinished Then '所有报告完成后才算完成
            Call Menu_Manage_检查最终完成(lngOrderID, False, strDocId)
        End If
    End If
    '病人状态跟踪
    Call StateCheck(intState)
    
    '发送状态同步消息
    Call mobjMsgCenter.Send_Msg_StateSync(lngOrderID)
    Exit Sub
ErrHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub AfterReportSaved(lngOrderID As Long, frmOwnerForm As Form, ByVal lngSaveType As Long)
'保存报告之后的处理
'执行过程：2-已报到；3-已检查；4-已报告；5-已审核；6-已完成
On Error GoTo ErrHandle
    Dim intState As Integer, lngSendId As Long
    Dim str签名 As String
    Dim str创建人 As String
    Dim str保存人 As String
    Dim bln保存结果阳性 As Boolean
    Dim blnCriticalValues As Boolean
    Dim blnImageQuality As Boolean
    Dim blnReportQuality As Boolean
    Dim blnConformDetermine As Boolean
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    
    arrSQL = Array()

    Call mobjWork_Report.zlRefreshFace(True)

    '获取本次检查的执行过程
    intState = getStudyState(lngOrderID, False, lngSendId, str创建人, str签名, str保存人, bln保存结果阳性, blnCriticalValues, blnImageQuality, blnReportQuality, blnConformDetermine)
    If intState = 6 Then Exit Sub
    
    'intState =1--已登记；2--已报到；3--已检查；4--已报告；5--已审核；6--已完成（本过程不存在这个返回值）
    If intState = 2 Or intState = 3 Then
        gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & "," & intState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        gstrSQL = "ZL_影像报告保存_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str创建人) & "','')"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        '报告保存时执行费用
        If mlngModule = G_LNG_VIDEOSTATION_MODULE And mSysPar.lngMoneyExeModle = 2 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            
            gstrSQL = "Zl_影像费用执行(" & lngOrderID & "," & lngSendId & ",4,NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    Else
        If intState = 4 Then
            '诊断签名，最后一次签名为医师,执行过程为已报告
            '有可能的情况 1-医师第N次签名 2-主任级别最后一次退签 3-修订模式下保存(签名级别=0)
            gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & "," & intState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            '应该填写创建人才准确，回退的时候，回退的人是保存人，但是不是报告创建人
            '医生诊断签名,无论是第N次，此时，报告人需要保存，复核人需要清空;
            gstrSQL = "ZL_影像报告保存_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str创建人) & "','')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        ElseIf intState = 5 Then
            '审核签名，主任及以上级别签名，签名级别>=2,执行过程为已审核
            gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & "," & intState & ",NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & mlngCur科室ID & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            gstrSQL = "ZL_影像报告保存_Update(" & lngOrderID & ",'" & IIf(mstrRPTExecutor <> "", mstrRPTExecutor, str创建人) & "','" & IIf(str签名 <> "", str签名, str保存人) & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    End If
    
    gcnOracle.BeginTrans        '----------保存检查状态及报告人
    
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存检查状态及报告人")
    Next i
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If (intState = 4 Or intState = 5) And IIf(mSysPar.lngHintType = 0, lngSaveType = 1, IIf(mSysPar.lngHintType = 1, lngSaveType = 2, False)) Then
        Dim strResultInput As String
        
        strResultInput = ""
        If mSysPar.blnReportWithResult Then '无影像诊断为阴性  -无提示自动标记
            gstrSQL = "ZL_影像检查_结果(" & lngOrderID & ",0)"
            zlDatabase.ExecuteProcedure gstrSQL, "标记阴阳性"
        End If
            
        If (Not blnCriticalValues And mSysPar.lngCriticalValues <> 0) Then strResultInput = "危急状态|"
        If (Not bln保存结果阳性 And mSysPar.blnIgnoreResult = False) Then strResultInput = strResultInput & "结果阳性|"
        If (Not blnImageQuality And mSysPar.strImageLevel <> "") And mSysPar.lngImageLevel <> 0 And InStr(mstrPrivs, "影像质控") > 0 Then strResultInput = strResultInput & "影像质量|"
        If (Not blnReportQuality And mSysPar.strReportLevel <> "") And mSysPar.lngReportLevel <> 0 And InStr(mstrPrivs, "报告质控") > 0 Then strResultInput = strResultInput & "报告质量|"
        If (Not blnConformDetermine And mSysPar.lngConformDetermine <> 0) Then strResultInput = strResultInput & "符合情况|"
 
        If strResultInput <> "" Then Call PromptResult(lngOrderID, mlngModule, frmOwnerForm, mlngCur科室ID, strResultInput)
    End If
    
    If intState = 5 And mSysPar.blnCompleteCommit Then   '如果“审核后直接完成”
        Call Menu_Manage_检查最终完成(lngOrderID, False)
    End If
    '病人状态跟踪
    Call StateCheck(intState)
    
    '发送状态同步消息
    Call mobjMsgCenter.Send_Msg_StateSync(lngOrderID)
    Exit Sub
ErrHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub UpdateStudyListState(lngAdviceId As Long, strStudyUID As String, blnAddImage As Boolean, blnStateChanged As Boolean)
    Dim strSql As String
    Dim intRowIndex As Integer
    
    With ufgStudyList
        intRowIndex = .FindRowIndex(CStr(lngAdviceId), "医嘱ID", True)
        
        If blnStateChanged And intRowIndex > 0 Then
            If blnAddImage Then '采图
                .Text(intRowIndex, "检查UID") = Nvl(strStudyUID, "A123456789")
                Call .UpdateSourceData(lngAdviceId, "检查UID", Nvl(strStudyUID, "A123456789"))
                
                Set .DataGrid.Cell(flexcpPicture, intRowIndex, .GetColIndex(GetStudyNumberDisplayName)) = imgList.ListImages(IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "病理", "影像")).Picture '改变图标
                
                If .Text(intRowIndex, "检查过程") = "已报到" Then
                    .Text(intRowIndex, "检查过程") = "已检查"
                    Call .UpdateSourceData(lngAdviceId, "检查过程", 3)
                    
                    .Text(intRowIndex, "检查状态") = 3
                End If
            Else '最后一次册图
                .Text(intRowIndex, "检查UID") = ""
                Call .UpdateSourceData(lngAdviceId, "检查UID", "")
                
                Set .DataGrid.Cell(flexcpPicture, intRowIndex, .GetColIndex(GetStudyNumberDisplayName)) = Nothing '改变图标
                
                If .Text(intRowIndex, "检查过程") = "已检查" Then
                    .Text(intRowIndex, "检查过程") = "已报到"
                    Call .UpdateSourceData(lngAdviceId, "检查过程", 2)
                    
                    .Text(intRowIndex, "检查状态") = 2
                End If
            End If
        End If
        
        '根据设置更新影像检查技师
        If mSysPar.blnWriteCapDoctor = True And blnStateChanged = True Then
            If mblnCnOracleIsHIS Then
                strSql = "Zl_影像检查_检查技师( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrUserNameNew, "") & "')"
                .Text(intRowIndex, "检查技师") = IIf(blnAddImage = True, mstrUserNameNew, "")
            Else
                strSql = "Zl_影像检查_检查技师( " & lngAdviceId & ",'" & IIf(blnAddImage = True, mstrUserNameHIS, "") & "')"
                .Text(intRowIndex, "检查技师") = IIf(blnAddImage = True, mstrUserNameHIS, "")
            End If
            
            zlDatabase.ExecuteProcedure strSql, GetWindowCaption
        End If
    End With
End Sub

Private Sub StateCheck(ByVal intState As Integer, Optional ByVal lngAdviceId As Long)
'----------------------------------------------------------
'功能：在病人列表中定位指定的记录
'参数： intState --病人检查状态   lngAdviceID --病人医嘱ID
'返回：无，直接在病人列表中定位
'----------------------------------------------------------
    If mSysPar.blnPatTrack Then
        If Not mblncmd登记 And Not mblncmd报到 And Not mblncmd检查 And Not mblncmd报告 And Not mblncmd审核 And Not mblncmd驳回 And Not mblncmd完成 Then
            Call RefreshList(lngAdviceId)
            Exit Sub
        End If
        
        Select Case intState '跟据病人新状态确定新状态过滤是否选中
            Case -1
                If Not mblncmd驳回 Then mblncmd驳回 = True
            Case 0, 1
                If Not mblncmd登记 Then mblncmd登记 = True
            Case 2
                If Not mblncmd报到 Then mblncmd报到 = True
            Case 3
                If Not mblncmd检查 Then mblncmd检查 = True
            Case 4
                If Not mblncmd报告 Then mblncmd报告 = True
            Case 5
                If Not mblncmd审核 Then mblncmd审核 = True
            Case 6
                If Not mblncmd完成 Then mblncmd完成 = True
        End Select
        
        Call RefreshList(lngAdviceId)
    Else '不跟踪只刷新列表
        Call RefreshList
    End If
End Sub

Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'功能：显示当前执行医嘱可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
    Dim strSql As String
        
    On Error GoTo errH
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Function
    End If
    
    objPopup.CommandBar.Controls.DeleteAll
    
    strSql = "Select Distinct C.编号,C.名称,C.说明" & _
        " From 病人医嘱记录 A,病历单据应用 B,病历文件列表 C" & _
        " Where A.ID=[1] And A.相关ID IS NULL" & _
        " And A.诊疗项目ID=B.诊疗项目ID" & _
        " And B.应用场合=[2] And B.病历文件ID=C.ID And C.种类=7" & _
        " Order by C.编号"
        
    If mListAdviceInf.intMoved = 1 Then
        strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
        strSql = Replace(strSql, "病人医嘱发送", "H病人医嘱发送")
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mListAdviceInf.lngAdviceId, mListAdviceInf.lngPatientFrom)
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!名称 & "(&0)")
            objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
            objControl.Category = M_STR_MODULE_MENU_TAG
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function


Private Sub FuncBillPrint(objControl As CommandBarControl)
'功能：打印诊疗单据
On Error GoTo ErrHandle
    If objControl.Parameter = "" Then '奇怪，直接按F10时，是一个空的Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    
    If objControl.Parameter = "" Then Exit Sub
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    If ReportPrintSet(gcnOracle, glngSys, objControl.Parameter, Me) Then
        Call ReportOpen(gcnOracle, glngSys, objControl.Parameter, Me, "NO=" & ufgStudyList.CurText("NO"), _
                       "性质=" & Val(ufgStudyList.CurText("记录性质")), "医嘱ID=" & mListAdviceInf.lngAdviceId, 1)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub NotificationAllModuleRefresh()
'通知所有模块刷新
    If Not mobjWork_His Is Nothing Then Call mobjWork_His.NotificationRefresh(hmAll)
    If Not mobjWork_Pathol Is Nothing Then Call mobjWork_Pathol.NotificationRefresh(mtAll)
    If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.NotificationRefresh
    
    If mblnUseActivexCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then Call mobjWork_ActiveVideo.zlNotifyRefresh
    End If

    If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.NotificationRefresh
End Sub


Private Sub DisableWorkModule()
'禁用工作模块
    tcDisable.Visible = True
    tcDisable.Translucence
End Sub


Private Sub EnableWorkModule()
'打开工作模块
    tcDisable.Visible = False
End Sub


Public Sub RefreshList(Optional ByVal lngAdviceId As Long = 0, Optional ByVal blnFromDB As Boolean = True)
'刷新数据列表
    Dim i As Integer
    Dim lngcur医嘱ID As Long
    Dim lngRow As Long
    Dim lngTopRow As Long
    
    If blnIsLoading = True Or ufgStudyList.IsLoading = True Then
        MsgBoxD Me, "数据加载中，请稍后重试...", vbInformation, Me.Caption
        Exit Sub
    End If
    
    blnIsLoading = True

On Error GoTo ErrHandle
    
    mblnAutoRefreshList = True
    
    With ufgStudyList
        If lngAdviceId <> 0 Then
            lngcur医嘱ID = lngAdviceId
        Else
            lngcur医嘱ID = Val(ufgStudyList.CurKeyValue) '当前行医嘱ID
            lngRow = .DataGrid.Row: lngTopRow = .DataGrid.TopRow               '当前行和顶行之间的差距
        End If
    
        
        Call LoadPatiList(blnFromDB)
        
        If ufgStudyList.GridRows <= 1 Then
            '当没有数据时，通知刷新工作模块中相关的数据
            mcurAdviceInf = GetNullAdviceInf
            mListAdviceInf = GetNullAdviceInf
            
            Call RefreshModuleAdviceInf
            Call NotificationAllModuleRefresh
            
            If TabWindow.Selected Is Nothing Then
                '选择第一个工作模块
                For i = 0 To TabWindow.ItemCount - 1
                    If TabWindow.Item(i).Visible Then
                        TabWindow(i).Selected = True
                        
                        mblnAutoRefreshList = False
                        Exit For
                    End If
                Next i
            End If
            
            Call RefreshTabWindow
            
            mblnAutoRefreshList = False
            blnIsLoading = False
            
            Exit Sub
        End If
        
        
        If lngcur医嘱ID = 0 Then
            'Call .LocateRow(1)
            Call ufgStudyList_OnSelChange
            
            mblnAutoRefreshList = False
            blnIsLoading = False
            
            Exit Sub
        End If
        
        '有记录时要重新定位回之前记录\
        lngcur医嘱ID = .FindRowIndex(CStr(lngcur医嘱ID), "医嘱ID", True)
        
        If lngcur医嘱ID <> -1 Then
            lngRow = Abs(lngRow - lngTopRow)
            If .DataGrid.Row = lngcur医嘱ID Then '当行未发生改变时，不会触发OnSelChange事件，因此当行相同时需手动触发CHANGE事件
                Call ufgStudyList_OnSelChange  '强制刷新右边子窗体
            Else
                .DataGrid.Row = lngcur医嘱ID
            End If
            
            .DataGrid.TopRow = IIf((.DataGrid.Row - lngRow) < 1, 1, (.DataGrid.Row - lngRow))
        Else
            If .DataGrid.Row <> 1 Then
                .DataGrid.Row = 1
            Else
                Call ufgStudyList_OnSelChange '强制刷新右边子窗体
            End If
        End If
        
    End With
    
    mblnAutoRefreshList = False
    blnIsLoading = False

    Exit Sub
ErrHandle:
    blnIsLoading = False
End Sub


Private Function GetExecuteState() As Long
'获取病理过程执行状态
    GetExecuteState = -1
    
    Select Case True
        Case optNeed.value And optNeed.Enabled
            GetExecuteState = 0
        Case optAccept.value And optAccept.Enabled
            GetExecuteState = 1
        Case optFinal.value And optFinal.Enabled
            GetExecuteState = 2
        Case optAll.value And optAll.Enabled
            GetExecuteState = 3
    End Select
End Function


Private Function GetFilterData() As ADODB.Recordset
'功能：取得当前过滤的SQL
    Dim strSQLBak As String
    Dim str来源 As String
    
    Dim lngType As Long          '1-使用医嘱相关作为基础检索，2-使用检查相关作为基础检索
    Dim strSql As String
    Dim strSubSql As String
    
    Dim strWithOrderTab As String   '医嘱子查询
    Dim strWithStudyTab As String   '检查子查询
    Dim strWithOrderCols As String  '医嘱相关查询列
    Dim strWithStudyCols As String  '检查相关查询列
    
    Dim strFilterOrder As String            '医嘱信息相关条件
    Dim strFilterStudy As String            '检查信息相关条件
    Dim strFilterDate As String             '查询日期范围条件
    Dim strFilterReportContext As String    '根据报告内容过滤
    Dim strFilterReportAdvice As String     '根据报告建议过滤
    Dim strFilterIllnessDiagnose As String  '根据疾病诊断过滤
    Dim strFilterTemp As String
    
    Dim strPatholCol As String
    Dim strPatholTab As String
    Dim strPatholFilter As String
    
    Set GetFilterData = Nothing
    
    strPatholCol = ""
    strPatholTab = ""
    strPatholFilter = ""
    
    '判断是否连接病理查询的相关表
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        strPatholTab = " 病理检查信息 o, 病理会诊信息 p "
        strPatholCol = " o.取材过程,o.制片过程,o.免疫过程,o.分子过程,o.特染过程,o.检查类型,o.病理号,o.综合质量 "
                        
        strPatholFilter = " h.医嘱ID=o.医嘱ID and o.病理医嘱ID=p.病理医嘱ID(+) "
    End If
    
    lngType = 1 '默认使用医嘱相关检索
    
    With SQLCondition
        '界面查找条件不使用时间索引
        If .门诊号 <> 0 Then
            strFilterOrder = " And C.门诊号=[1]"
        ElseIf .住院号 <> 0 Then
            strFilterOrder = " And C.住院号=[2]"
        ElseIf .健康号 <> "" Then
            strFilterOrder = " And C.健康号=[8]"
        ElseIf .就诊卡 <> "" Then
            strFilterOrder = " And C.就诊卡号=[3]"
        ElseIf .姓名 <> "" And InStr(.姓名, "*") = 0 Then   '姓名特殊处理，带*号表示模糊查询
            .姓名 = .姓名 & "%"
            strFilterOrder = " And C.姓名 like [4]"
        ElseIf .身份证 <> "" Then
            strFilterOrder = " And C.身份证号=[5]"
        ElseIf .IC卡 <> "" Then
            strFilterOrder = " And C.IC卡=[6]"
        ElseIf .单据号 <> "" Then
            strFilterOrder = " And A.NO=[7] "
        ElseIf .检查号 <> 0 Then
            lngType = 2 '使用检查相关信息作为基础检索
            
            If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
            
            If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                strFilterStudy = strFilterStudy & " H.检查号=[9] "
            Else
                '如果是病理系统，则这里需要根据病理号进行查询
                strFilterStudy = strFilterStudy & " o.病理号=upper([9]) "
            End If
        ElseIf .病人ID <> 0 Then
            strFilterOrder = " And C.病人ID=[31]"
        Else
        '其他条件查询，使用时间索引
            
            '填写过滤时间条件
            '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
            If .时间类型 = 1 Then       '按申请时间
                strFilterDate = " A.发送时间 Between [10] and "
            ElseIf .时间类型 = 2 Then   '按报到时间
                strFilterDate = " A.报到时间 Between [10] and "
            ElseIf .时间类型 = 3 Then                        '采图时间或者病理内部检查的申请时间
                lngType = 2
                
                If mlngModule <> G_LNG_PATHOLSYS_NUM Then
                    strFilterDate = strFilterDate & " H.接收日期 Between [10] and  "
                Else
                    strPatholTab = strPatholTab & " , 病理申请信息 q"

                    strFilterDate = strFilterDate & " o.病理医嘱ID = q.病理医嘱ID and q.申请时间 between [10] and "
                End If
            Else                        '按安排时间
                strFilterDate = strFilterDate & " B.开始执行时间 Between [10] and  "
            End If
            
            If .结束时间 <> CDate(0) Then
                strFilterDate = strFilterDate & " [11] "
            Else
                strFilterDate = strFilterDate & " Sysdate+1/(24*3600) "
            End If
            
            '先处理姓名中带*号的，进行带时间索引的模糊查询
            If .姓名 <> "" And InStr(.姓名, "*") <> 0 Then
                .姓名 = Replace(.姓名, "*", "%")
                strFilterOrder = strFilterOrder & " And C.姓名 || '' like [4]"
            End If
            
            If .性别 <> "" Then
                strFilterOrder = strFilterOrder & " And C.性别=[27]"
            End If
        
        
            '病人年龄-开始年龄(只有当条件使用“到”，即在多少年龄之间时，才使用开始年龄)
            If .开始年龄 <> -1 Then
                If .年龄条件 = "~" Then
                    strFilterOrder = strFilterOrder & " And ZL_AgeToDays(C.年龄)>=[28]"
                End If
            End If
            
            '病人年龄-结束年龄
            If .结束年龄 <> -1 Then
                If .年龄条件 = "~" Then
                    strFilterOrder = strFilterOrder & " And ZL_AgeToDays(C.年龄)<=[29]"
                Else
                    strFilterOrder = strFilterOrder & " And ZL_AgeToDays(C.年龄)" & .年龄条件 & "[29]"
                End If
            End If
            
            If .病人科室 <> 0 Then
                strFilterOrder = strFilterOrder & " And B.病人科室ID+0=[12] "
            End If

            If .标本部位 <> "" Then
                strFilterOrder = strFilterOrder & " And instr(B.医嘱内容,[13])>0"
            End If
            
            If .结果阳性 <> -1 Then
                strFilterOrder = strFilterOrder & " And Nvl(A.结果阳性, 0)=[30]"
            End If
            
            If .诊断医生 <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " H.报告人=[14] "
            End If
            
            If .审核医生 <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " H.复核人=[15] "
            End If
            
            If .影像质量 <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " H.影像质量=[16] "
            End If
            
            If .检查技师 <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " H.检查技师=[17] "
            End If
            
            '影像类别有两个地方做过滤条件的选择，过滤窗口和主程序上面，以主程序中的为主
            If mintcmd影像类别 <= 0 Then
                If .影像类别 <> "" Then
                    If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                    strFilterStudy = strFilterStudy & " H.影像类别=[18] "
                End If
            End If
            
            If .随访 <> "" Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & " Instr(H.随访描述,[19])>0 "
            End If
            
            If .疾病诊断 <> "" Then
                strFilterIllnessDiagnose = "( Select t.医嘱id From 病人医嘱报告 t Where t.病历id IN " & _
                                                                        " (Select Distinct A.ID  " & _
                                                                        "From 电子病历记录 A,电子病历内容 B " & _
                                                                        "Where A.创建时间>[10] AND A.Id=B.文件ID  " & _
                                                                            "And B.对象类型=7 And instr(B.对象属性,'52;')>0 And instr(B.内容文本,[20])>0))"
            End If
            
            
            If .检查所见 <> "" Then
                strFilterTemp = " (b.内容文本 ='检查所见' And Instr(c.内容文本, [21]) > 0)"
            End If
            
            If .诊断意见 <> "" Then
                If strFilterTemp = "" Then
                    strFilterTemp = " (b.内容文本 ='诊断意见' And Instr(c.内容文本, [22]) > 0)"
                Else
                    strFilterTemp = strFilterTemp & " or (b.内容文本 ='诊断意见' And Instr(c.内容文本, [22]) > 0)"
                End If
            End If
            
            If .建议 <> "" Then
                If strFilterTemp = "" Then
                    strFilterTemp = " (b.内容文本 ='建议' And Instr(c.内容文本, [23]) > 0)"
                Else
                    strFilterTemp = strFilterTemp & " or (b.内容文本 ='建议' And Instr(c.内容文本, [23]) > 0)"
                End If
            End If
            
            If strFilterTemp <> "" Then
                strFilterTemp = " (" & strFilterTemp & ")"
                strFilterReportAdvice = "( Select t.医嘱id From 病人医嘱报告 t Where t.病历id IN " _
                    & " (Select Distinct a.ID From 电子病历记录 a, 电子病历内容 b,电子病历内容 c " _
                    & " Where a.创建时间 > [10] And a.Id = b.文件id And b.Id = C.父ID And b.对象类型 = 3 And c.对象类型 = 2 And c.终止版 = 0 and " _
                    & strFilterTemp & "))"
            End If
        End If
    
        If mSysPar.blnNoShowCancel Then '不显示取消登记的检查
            strFilterOrder = strFilterOrder & " And A.执行状态<>2 "
        End If
        
        If mblncmd本次 Then        '只显示本次住院记录
            strFilterOrder = strFilterOrder & vbNewLine & " And (B.病人来源=2 And B.主页ID=C.住院次数 Or Nvl(B.病人来源,0)<>2)"
        End If
        
        '是否选择了全部科室
        If mblnAllDepts = True Then
            strFilterOrder = strFilterOrder & " And Instr( [25],A.执行部门ID ) >0"
            
            If lngType = 2 Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & "  Instr([25],H.执行科室ID) >0 "
            End If
        Else
            strFilterOrder = strFilterOrder & " And A.执行部门ID+0=[24]"
            
            If lngType = 2 Then
                If strFilterStudy <> "" Then strFilterStudy = strFilterStudy & " AND "
                strFilterStudy = strFilterStudy & "  H.执行科室ID+0=[24] "
            End If
        End If
        
        '检索报告内容
        If .报告内容 <> "" Then
            strFilterReportContext = "( Select t.医嘱id From 病人医嘱报告 t Where t.病历id In " & _
                                                                    " (Select Distinct A.ID " & _
                                                                    " From 电子病历记录 A,电子病历内容 B " & _
                                                                    " Where A.创建时间>[10] AND A.Id=B.文件ID " & _
                                                                    " And B.对象类型=2 And instr(B.内容文本,[26])>0 And B.终止版 = 0)) "
        End If
        
        
        '构造查询语句
        
        '医嘱相关子查询列
        strWithOrderCols = "A.医嘱ID,A.发送号,A.首次时间,A.发送时间,A.执行状态,A.执行过程,nvl(A.执行间,' ') as 执行间, A.结果阳性 , " & _
                            " A.NO, A.发送人, A.执行部门ID, A.记录性质, A.计费状态, A.采样时间, " & _
                            " B.ID, B.相关ID,B.主页ID,B.病人ID, B.病人科室ID,B.挂号单,B.病人来源,B.医嘱内容,B.标本部位, " & _
                            " B.紧急标志,B.婴儿,B.开嘱医生,B.姓名,B.性别,B.年龄,B.诊疗项目ID,F.影像类别, " & _
                            " C.就诊卡号, C.身份证号, C.结算模式, decode(B.病人来源,2,D.病人类型,C.病人类型) as 病人类型, C.住院号, C.门诊号, C.当前床号, C.当前病区ID, D.路径状态, E.名称 "
    
        '检查相关子查询列
        strWithStudyCols = "H.医嘱ID, H.姓名,H.检查号,H.性别,H.年龄,H.身高,H.体重,H.影像质量,H.报告质量,H.符合情况,H.是否技师确认," & _
                            " H.完成人,H.是否电子胶片,H.是否打印,H.报告操作,H.绿色通道,H.报告打印,H.报告人,H.复核人,H.检查技师,H.检查技师二,H.接收日期,H.图像位置, " & _
                            " H.随访描述,H.诊断分类,H.检查UID,H.发送号,H.关联ID,H.报到人, H.报告发放,H.发放胶片,H.危急状态,J.医嘱ID as 申请单医嘱 " & _
                            IIf(strPatholCol = "", "", "," & strPatholCol & ",p.id as 会诊ID,p.当前状态 as 会诊状态,p.会诊医师" & _
                            ", (select count(1) from 病理检查信息 V , 病理申请信息 W where V.病理医嘱ID=w.病理医嘱id and v.医嘱id=H.医嘱ID and w.补费状态=1) as 补费 ")
    
    
        strSql = ""
        strSubSql = ""
        
        strWithOrderTab = ""
        strWithStudyTab = ""
        
        If lngType = 1 Then
            '以医嘱查询为主
            
            '不能删除该查询中的“影像检查项目”数据表，因为删除后，会造成数据的查询效率较低（删除后，则需要使用病人医嘱发送的执行部门ID作为条件过滤检查，然而该字段没有索引）
            strWithOrderTab = "tmpOrder as(select " & strWithOrderCols & vbNewLine & _
                              " from 病人医嘱发送  A, 病人医嘱记录 B,病人信息 C, 病案主页 D,部门表 E,影像检查项目 F " & vbNewLine & _
                              " Where a.医嘱ID = b.ID And b.病人ID = c.病人ID " & vbNewLine & _
                                      "     And B.病人科室ID=E.ID " & vbNewLine & _
                                      "     And B.病人ID = D.病人ID(+) And B.主页ID+0 = D.主页ID(+) And B.诊疗项目ID+0 =F.诊疗项目ID " & vbNewLine & _
                                      "     " & IIf(strFilterOrder = "", " ", strFilterOrder) & vbNewLine & _
                                                IIf(strFilterDate = "", "", "     And " & strFilterDate) & "  and B.医嘱状态 <> 4" & ")"
                                                      
'            strWithStudyTab = "tmpStudy as(select " & strWithStudyCols & vbNewLine & _
'                                " from 影像检查记录 H " & IIf(strPatholTab = "", "", " ," & strPatholTab) & " , tmpOrder I,影像申请单图像 J " & vbNewLine & _
'                                " where h.医嘱ID=I.医嘱ID  and H.医嘱ID=J.医嘱ID(+) " & vbNewLine & _
'                                IIf(strPatholFilter = "", "", "     And " & strPatholFilter) & _
'                                IIf(strFilterStudy = "", "", "     And " & strFilterStudy) & ")"
                             
                             
            '查询包含主医嘱和部位医嘱的医嘱id数据
            If strFilterStudy <> "" Then
                strSubSql = "(select id from tmpOrder I, tmpStudy J where I.相关Id=J.医嘱Id " & _
                                " Union All " & _
                                " select I.医嘱Id as id from tmpStudy I) K "
            End If
            
            strSql = " with " & strWithOrderTab '& "," & vbNewLine & strWithStudyTab
            
            '避免关联with表查询，oracle10g下存在一个bug，当在另一个with查询中使用另外的with查询数据表时，或造成游标数量提纲以及占用内存解析
            
            strSql = strSql & vbNewLine & _
                        " select distinct L.医嘱ID,L.相关ID,L.发送号,L.首次时间 报到时间,L.发送时间 申请时间,L.执行状态,nvl(L.执行过程,0) 检查过程,L.执行间,L.结果阳性 阳性,M.危急状态 危急," & vbNewLine & _
                        "     L.病人ID,L.主页ID,L.挂号单,L.病人科室ID,L.病人来源 来源,L.医嘱内容,L.标本部位," & vbNewLine & _
                        "     Nvl(L.紧急标志, 0) 紧急标志, Nvl(L.婴儿, 0) 婴儿,L.开嘱医生,L.NO,L.当前床号,L.当前病区ID,Decode(L.病人来源,2,L.住院号,L.门诊号) 标识号," & vbNewLine & _
                        "     Nvl(L.姓名,M.姓名) 姓名,L.影像类别,M.检查号,Nvl(L.性别,M.性别) 性别,Nvl(L.年龄,M.年龄) 年龄,M.身高,M.体重,M.影像质量,M.报告质量,M.符合情况," & vbNewLine & _
                        "     Decode(L.病人来源,3,L.开嘱医生,L.发送人) 登记人,M.报到人, M.报告发放,M.发放胶片,M.关联ID,L.记录性质, " & vbNewLine & _
                        "     M.完成人,M.是否电子胶片,M.是否打印,M.报告操作,M.绿色通道,M.报告打印,M.报告人,M.复核人,M.是否技师确认,M.检查技师,M.检查技师二,M.接收日期 采图时间," & vbNewLine & _
                        "     M.随访描述,M.诊断分类,M.检查UID,M.图像位置,L.执行部门ID as 执行科室ID,0 as 转出,L.名称 AS 病人科室, L.采样时间, " & vbNewLine & _
                        "     L.就诊卡号,L.NO as 单据号,L.身份证号,L.结算模式,L.病人类型,L.路径状态,L.计费状态,Decode(L.记录性质,2,1,Decode(L.计费状态,3,1,0)) as 收费 ,f.医嘱ID As 申请单医嘱 " & vbNewLine & _
                                IIf(strPatholCol = "", "", _
                                        "    ,M.会诊ID,M.会诊状态,M.会诊医师,nvl(M.补费,0) as 补费, '' as 病理执行状态,decode(M.病理号,null,'未核收','已核收') as 核收情况," & _
                                        "    decode(M.检查类型,0,'常规',1,'冰冻',2,'细胞',3,'会诊',4,'尸检',5,'快速石蜡',null) as  检查类别, " & _
                                        Replace(strPatholCol, "o.", "M.") & vbNewLine) & _
                        " From tmpOrder L,影像检查记录 M, 影像申请单图像 f " & vbNewLine & _
                                IIf(strSubSql = "", "", "," & strSubSql) & vbNewLine & _
                        " Where L.Id=M.医嘱ID(+) And m.医嘱ID=f.医嘱ID(+) " & IIf(strSubSql = "", "", " And L.Id=K.Id" & vbNewLine) & _
                                IIf(strPatholFilter = "", "", "     And " & strPatholFilter) & _
                                IIf(strFilterStudy = "", "", "     And " & strFilterStudy) & _
                                IIf(strFilterIllnessDiagnose = "", "", "     And L.Id In" & strFilterIllnessDiagnose & vbNewLine) & _
                                IIf(strFilterReportAdvice = "", "", "     And L.Id In" & strFilterReportAdvice & vbNewLine) & _
                                IIf(strFilterReportContext = "", "", "     And L.Id In" & strFilterReportContext & vbNewLine)
            
            
            
        Else
            '以检查查询为主
            strWithStudyTab = "tmpStudy as(select " & strWithStudyCols & vbNewLine & _
                                " from 影像检查记录 H " & IIf(strPatholTab = "", "", " ," & strPatholTab) & ",影像申请单图像 J " & vbNewLine & _
                                " where H.医嘱ID=J.医嘱ID(+) " & vbNewLine & _
                                IIf(strPatholFilter = "", "", "     And " & strPatholFilter) & vbNewLine & _
                                IIf(strFilterStudy = "", "", "     And " & strFilterStudy) & vbNewLine & _
                                IIf(strFilterDate = "", "", "     And " & strFilterDate) & ")"
                                
            strWithOrderTab = "tmpOrder as(select " & strWithOrderCols & vbNewLine & _
                              " from 病人医嘱发送  A, 病人医嘱记录 B,病人信息 C, tmpStudy G, 病案主页 D,部门表 E,影像检查项目 F " & vbNewLine & _
                              " Where a.医嘱ID = b.ID And b.病人ID = c.病人ID " & vbNewLine & _
                                      "     And B.病人科室ID=E.ID " & vbNewLine & _
                                      "     And (B.ID=G.医嘱ID) " & vbNewLine & _
                                      "     And B.病人ID = D.病人ID(+) And B.主页ID+0 = D.主页ID(+) And B.诊疗项目ID+0 =F.诊疗项目ID " & vbNewLine & _
                                      "     " & strFilterOrder & " and B.医嘱状态 <> 4" & vbNewLine & _
                                      " Union All " & vbNewLine & _
                                " select " & strWithOrderCols & vbNewLine & _
                              " from 病人医嘱发送  A, 病人医嘱记录 B,病人信息 C, tmpStudy G, 病案主页 D,部门表 E,影像检查项目 F " & vbNewLine & _
                              " Where a.医嘱ID = b.ID And b.病人ID = c.病人ID " & vbNewLine & _
                                      "     And B.病人科室ID=E.ID " & vbNewLine & _
                                      "     And (B.相关ID=G.医嘱ID) " & vbNewLine & _
                                      "     And B.病人ID = D.病人ID(+) And B.主页ID+0 = D.主页ID(+) And B.诊疗项目ID+0 =F.诊疗项目ID " & vbNewLine & _
                                      "     " & strFilterOrder & " and B.医嘱状态 <> 4" & vbNewLine & _
                                      ")"
                                                     
            strSql = " with " & strWithStudyTab & "," & vbNewLine & strWithOrderTab
                                             
        
        
            strSql = strSql & vbNewLine & _
                        " select distinct L.医嘱ID,L.相关ID,L.发送号,L.首次时间 报到时间,L.发送时间 申请时间,L.执行状态,nvl(L.执行过程,0) 检查过程,L.执行间,L.结果阳性 阳性,M.危急状态 危急," & vbNewLine & _
                        "     L.病人ID,L.主页ID,L.挂号单,L.病人科室ID,L.病人来源 来源,L.医嘱内容,L.标本部位," & vbNewLine & _
                        "     Nvl(L.紧急标志, 0) 紧急标志, Nvl(L.婴儿, 0) 婴儿,L.开嘱医生,L.NO,L.当前床号,L.当前病区ID,Decode(L.病人来源,2,L.住院号,L.门诊号) 标识号," & vbNewLine & _
                        "     Nvl(L.姓名,M.姓名) 姓名,L.影像类别,M.检查号,Nvl(L.性别,M.性别) 性别,Nvl(L.年龄,M.年龄) 年龄,M.身高,M.体重,M.影像质量,M.报告质量,M.符合情况," & vbNewLine & _
                        "     Decode(L.病人来源,3,L.开嘱医生,L.发送人) 登记人,M.报到人, M.报告发放,M.发放胶片,M.关联ID,L.记录性质, " & vbNewLine & _
                        "     M.完成人,M.是否电子胶片,M.是否打印,M.报告操作,M.绿色通道,M.报告打印,M.报告人,M.复核人,M.是否技师确认,M.检查技师,M.检查技师二,M.接收日期 采图时间," & vbNewLine & _
                        "     M.随访描述,M.诊断分类,M.检查UID,M.图像位置,L.执行部门ID as 执行科室ID,0 as 转出,L.名称 AS 病人科室, L.采样时间, " & vbNewLine & _
                        "     L.就诊卡号,L.NO as 单据号,L.身份证号,L.结算模式,L.病人类型,L.路径状态,L.计费状态,Decode(L.记录性质,2,1,Decode(L.计费状态,3,1,0)) as 收费 ,M.申请单医嘱 " & vbNewLine & _
                                IIf(strPatholCol = "", "", _
                                        "    ,M.会诊ID,M.会诊状态,M.会诊医师,nvl(M.补费,0) as 补费, '' as 病理执行状态,decode(M.病理号,null,'未核收','已核收') as 核收情况," & _
                                        "    decode(M.检查类型,0,'常规',1,'冰冻',2,'细胞',3,'会诊',4,'尸检',5,'快速石蜡',null) as  检查类别, " & _
                                        Replace(strPatholCol, "o.", "M.") & vbNewLine) & _
                        " From tmpOrder L,tmpStudy M " & vbNewLine & _
                                IIf(strSubSql = "", "", "," & strSubSql) & vbNewLine & _
                        " Where L.Id=M.医嘱ID(+) " & IIf(strSubSql = "", "", " And L.Id=K.Id" & vbNewLine) & _
                                IIf(strFilterIllnessDiagnose = "", "", "     And L.Id In" & strFilterIllnessDiagnose & vbNewLine) & _
                                IIf(strFilterReportAdvice = "", "", "     And L.Id In" & strFilterReportAdvice & vbNewLine) & _
                                IIf(strFilterReportContext = "", "", "     And L.Id In" & strFilterReportContext & vbNewLine)

        End If
        
        '如果有数据转出则还要检索后备表
        If mblnMoved Then
            strSql = "Select /*+ RULE*/ * From (" & vbNewLine & strSql & vbNewLine & ")"
            strSQLBak = strSql
            strSQLBak = GetHistoryQuerySql(strSQLBak)
            
            strSQLBak = Replace(strSQLBak, "0 as 转出", "1 as 转出")
            strSql = strSql & " Union ALL " & strSQLBak
            
            strSql = "Select * From (" & vbNewLine & strSql & vbNewLine & ") Order by 检查过程,报到时间,申请时间"
        Else
            strSql = "Select /*+ RULE*/ * From (" & vbNewLine & strSql & vbNewLine & ") Order by 检查过程,报到时间,申请时间"
        End If
        
        '1: 门诊号    2: 住院号    3: 就诊卡号    4: 姓名    5: 身份证号    6: IC卡    7: 单据号    8: 健康号
        '9: 检查号    10: 开始时间    11: 结束时间    12: 病人科室ID    13: 医嘱内容    14: 报告人    15: 复核    16: 影像质量
        '17: 检查技师    18: 影像类别    19: 随访描述    20: 内容文本-疾病诊断    21: 内容文本-检查所见    22: 内容文本-诊断意见    23: 内容文本 -建议
        '24: 执行部门Id    25: 当前所属科室Ids    26: 报告内容    27: 性别    28: 开始年龄    29: 结束年龄    30: 结果阳性    31: 病人ID
        Set GetFilterData = GetDataToLocal(strSql, "提取病人列表", .门诊号, .住院号, .就诊卡, .姓名, .身份证, _
                                            .IC卡, .单据号, .健康号, .检查号, .开始时间, .结束时间, .病人科室, .标本部位, _
                                            .诊断医生, .审核医生, .影像质量, .检查技师, .影像类别, .随访, _
                                            .疾病诊断, .检查所见, .诊断意见, .建议, mlngCur科室ID, _
                                            mstrCanUse科室IDs, .报告内容, .性别, .开始年龄, .结束年龄, .结果阳性, .病人ID)
    End With
End Function


Private Function GetFilterWhere() As String
    Dim objControl As CommandBarControl
    Dim strFilter As String
    Dim strModalitys As String
    Dim lngCurExecuteState As Long
    Dim i As Long
    
    strFilter = ""
        
    '过滤检查类别
    If mlngModule <> G_LNG_PATHOLSYS_NUM And mintcmd影像类别 <> 0 Then
        '影像类别有两个地方做过滤条件的选择，过滤窗口和主程序上面，以主程序中的为主
        Set objControl = cbrdock.FindControl(, ID_影像类别)
        For i = 1 To objControl.CommandBar.Controls.Count
            If objControl.CommandBar.FindControl(, ID_影像类别 + i).Checked = False Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & " 影像类别<>'" & objControl.CommandBar.FindControl(, ID_影像类别 + i).DescriptionText & "'"
            End If
        Next i
    End If

    '过滤检查执行间
    If mlngModule <> G_LNG_PATHOLSYS_NUM And mintcmd影像执行间 <> 0 Then
        Set objControl = cbrdock.FindControl(, ID_影像执行间)
        For i = 1 To objControl.CommandBar.Controls.Count
            If objControl.CommandBar.FindControl(, ID_影像执行间 + i).Checked = False Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & " 执行间<>'" & objControl.CommandBar.FindControl(, ID_影像执行间 + i).DescriptionText & "'"
            End If
        Next i
    End If

    '过滤病人来源
    If (Abs(mblncmd门诊) + Abs(mblncmd住院) + Abs(mblncmd体检) + Abs(mblncmd外诊)) Mod 4 <> 0 Then
        If Not mblncmd门诊 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " 来源<>1"
        End If
        
        If Not mblncmd住院 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " 来源<>2"
        End If
        
        If Not mblncmd体检 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " 来源<>4"
        End If
        
        If Not mblncmd外诊 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " 来源<>3"
        End If
    End If


    '检查过程过滤
    If (Abs(mblncmd登记) + Abs(mblncmd报到) + Abs(mblncmd检查) + Abs(mblncmd报告) + Abs(mblncmd审核) + Abs(mblncmd驳回) + Abs(mblncmd完成)) Mod 7 <> 0 Then
        If Not mblncmd登记 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & " 检查过程<>0 and 检查过程<>1"
        End If
        
        If Not mblncmd报到 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "检查过程<>2"
        End If
        
        If Not mblncmd检查 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "检查过程<>3"
        End If
        
        If Not mblncmd报告 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "检查过程<>4"
        End If
        
        If Not mblncmd审核 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "检查过程<>5 "
        End If
        
        If Not mblncmd驳回 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "检查过程<>-1 "
        End If
        
        If Not mblncmd完成 Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & "检查过程<>6"
        End If
    End If


    '对病理部分，需要单独进行一些过滤处理
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        '如果全选，就相当于全不选
        If (Abs(mblncmd常规) + Abs(mblncmd冰冻) + Abs(mblncmd细胞) + Abs(mblncmd会诊) + Abs(mblncmd尸检) + Abs(mblncmd快速石蜡)) Mod 6 <> 0 Then

            If Not mblncmd常规 Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "检查类别<>'常规'"
            End If

            If Not mblncmd冰冻 Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "检查类别<>'冰冻'"
            End If

            If Not mblncmd细胞 Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "检查类别<>'细胞'"
            End If

            If Not mblncmd会诊 Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "检查类别<>'会诊'"
            End If

            If Not mblncmd尸检 Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "检查类别<>'尸检'"
            End If

            If Not mblncmd快速石蜡 Then
                If strFilter <> "" Then strFilter = strFilter & " and "
                strFilter = strFilter & "检查类别<>'快速石蜡'"
            End If
            
        End If

        '过滤当前页面数据
        If tabFilter.tag Then

            lngCurExecuteState = GetExecuteState
            Select Case tabFilter.Selected.tag
                Case "取材"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '需取材
                        strFilter = strFilter & "取材过程 = 1"
                    ElseIf lngCurExecuteState = 2 Then                      '已取材
                        strFilter = strFilter & "取材过程 = 2"
                    ElseIf lngCurExecuteState = 3 Then                      '所有
                        strFilter = strFilter & "取材过程 > 0"
                    End If

                Case "制片"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '需制片
                        strFilter = strFilter & "制片过程 = 1"
                    ElseIf lngCurExecuteState = 1 Then                      '制片接受
                        strFilter = strFilter & "制片过程 = 2"
                    ElseIf lngCurExecuteState = 2 Then                      '已制片
                        strFilter = strFilter & "制片过程 = 3"
                    ElseIf lngCurExecuteState = 3 Then                      '所有
                        strFilter = strFilter & "制片过程 > 0"
                    End If

                Case "免疫"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '需免疫
                        strFilter = strFilter & "免疫过程 = 1"
                    ElseIf lngCurExecuteState = 1 Then                      '免疫接受
                        strFilter = strFilter & "免疫过程 = 2"
                    ElseIf lngCurExecuteState = 2 Then                      '已免疫
                        strFilter = strFilter & "免疫过程 = 3"
                    ElseIf lngCurExecuteState = 3 Then                      '所有
                        strFilter = strFilter & "免疫过程 > 0"
                    End If

                Case "特染"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '需特染
                        strFilter = strFilter & "特染过程 = 1"
                    ElseIf lngCurExecuteState = 1 Then                      '特染接受
                        strFilter = strFilter & "特染过程 = 2"
                    ElseIf lngCurExecuteState = 2 Then                      '已特染
                        strFilter = strFilter & "特染过程 = 3"
                    ElseIf lngCurExecuteState = 3 Then                      '所有
                        strFilter = strFilter & "特染过程 > 0"
                    End If


                Case "分子"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '需分子
                        strFilter = strFilter & "分子过程 = 1"
                    ElseIf lngCurExecuteState = 1 Then                      '分子接受
                        strFilter = strFilter & "分子过程 = 2"
                    ElseIf lngCurExecuteState = 2 Then                      '已分子
                        strFilter = strFilter & "分子过程 = 3"
                    ElseIf lngCurExecuteState = 3 Then                      '所有
                        strFilter = strFilter & "分子过程 > 0"
                    End If

                Case "会诊"
                    If strFilter <> "" Then strFilter = strFilter & " and "
                    
                    If lngCurExecuteState = 0 Then                          '需会诊
                        strFilter = strFilter & "会诊状态=0 and 会诊医师='" & UserInfo.姓名 & "'"
                    ElseIf lngCurExecuteState = 2 Then                      '已会诊
                        strFilter = strFilter & "会诊状态<>0 and 会诊医师='" & UserInfo.姓名 & "'"
                    ElseIf lngCurExecuteState = 3 Then                      '所有
                        strFilter = strFilter & " 会诊ID > 0 and 会诊医师='" & UserInfo.姓名 & "'"
                    End If

                Case "所有"
            End Select
        End If
    End If
        
    GetFilterWhere = strFilter
End Function


Private Sub LoadPatiList(Optional ByVal blnFromDB As Boolean = True)
'功能：读取当前医技科室的执行医嘱(病人)清单
    Dim rsList As ADODB.Recordset

    If Not mblnInitOk Then Exit Sub      '初始化未完成
    
    mblnvsRefresh = True
    
    If blnFromDB Then
        If Not mblnIsIntegratedQuery Then
            If mlngSysQuerySchemeId > 0 Then    '使用自定义系统查询方案
                Call ExecuteCustomQuery(mlngSysQuerySchemeId)
                mblnvsRefresh = False
                Exit Sub
            End If
        Else
            mblnIsIntegratedQuery = False
        End If
        
        Set rsList = GetFilterData()
        Set ufgStudyList.AdoData = rsList
    End If
    
    ufgStudyList.AdoFilter = GetFilterWhere
    
    '用binddata的方法比使用refreshdata的方法快
    Call ufgStudyList.BindData
    
    '恢复排序
    Call ufgStudyList.ResetSort(mlngSortCol, mintSortOrder)
    
    Call RefreshStatusBarInf
    
    mblnvsRefresh = False
End Sub


Private Sub picLoadState_Resize()
On Error GoTo ErrHandle
    labLoadState.Left = Fix((picLoadState.Width - labLoadState.Width) / 2)
    labLoadState.Top = Fix((picLoadState.Height - labLoadState.Height) / 2)
    
    picSmile.Left = labLoadState.Left - picSmile.Width
    picSmile.Top = labLoadState.Top - 80
    
ErrHandle:
End Sub

Private Sub picReportContainer_Resize()
On Error GoTo ErrHandle
    
    If mobjWork_Report Is Nothing Then Exit Sub
    
    Call mobjWork_Report.UpdateSize
    
ErrHandle:
End Sub



Private Sub picWindow_Resize()
On Error GoTo ErrHandle
    With TabWindow
        If GetWorkModuleCount = 1 Then
            TabWindow.PaintManager.ClientMargin.Top = -30
        Else
            TabWindow.PaintManager.ClientMargin.Top = 0
        End If
        
        .Left = 0
        .Width = picWindow.ScaleWidth
        .Height = picWindow.ScaleHeight + IIf(GetWorkModuleCount = 1, ScaleY(30, vbTwips, vbPixels), 0)
    End With
    
    tcDisable.Left = 0
    tcDisable.Top = IIf(TabWindow.PaintManager.ClientMargin.Top < 0, 0, IIf(mbytFontSize = 9, 440, 470))
    tcDisable.Width = picWindow.ScaleWidth
    tcDisable.Height = picWindow.ScaleHeight - IIf(TabWindow.PaintManager.ClientMargin.Top < 0, 0, IIf(mbytFontSize = 9, 440, 470))
ErrHandle:
End Sub

Private Sub tabFilter_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo ErrHandle
    If Not mblnInitOk Then Exit Sub
    
    If tabFilter.ItemCount < 7 Then Exit Sub
    If Not ufgStudyList.Visible Then Exit Sub
    
    optAccept.Enabled = IIf(Item.tag = "取材" Or Item.tag = "会诊" Or Item.tag = "所有", False, True)
    
    optNeed.Enabled = IIf(Item.tag = "所有", False, True)
    optFinal.Enabled = IIf(Item.tag = "所有", False, True)
    optAll.Enabled = IIf(Item.tag = "所有", False, True)
    
    If (Item.tag = "取材" Or Item.tag = "会诊") And optAccept.value Then
        '当check值被改变时，会触发控件的click事件而执行RefreshList方法
        optNeed.value = True
    Else
        Call RefreshList(, False)
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ConfigSubForm(ByVal Item As XtremeSuiteControls.ITabControlItem)
'配置子窗口界面
On Error GoTo ErrHandle
    Dim lngIndex As Integer
    Dim objItem As XtremeSuiteControls.TabControlItem
    
    If mblnLoadSubFrom Then Exit Sub
    If Item.Handle <> picTemp.hWnd Then Exit Sub
    
    mblnLoadSubFrom = True
    lngIndex = Item.Index
    
    Set objItem = Nothing
    
    Select Case Item.tag
        Case "影像图象"
            Set objItem = TabWindow.InsertItem(lngIndex, "影像记录", mfrmWork_PacsImg.hWnd, Item.Image)
                
        Case "标本核收"
            Set objItem = TabWindow.InsertItem(lngIndex, "标本核收", mobjWork_Pathol.GetModule(mtSpecimen).hWnd, Item.Image)

        Case "病理取材"
            Set objItem = TabWindow.InsertItem(lngIndex, "病理取材", mobjWork_Pathol.GetModule(mtMaterial).hWnd, Item.Image)
            
        Case "病理制片"
            Set objItem = TabWindow.InsertItem(lngIndex, "病理制片", mobjWork_Pathol.GetModule(mtSlices).hWnd, Item.Image)
            
        Case "病理特检"
            Set objItem = TabWindow.InsertItem(lngIndex, "病理特检", mobjWork_Pathol.GetModule(mtSpeExam).hWnd, Item.Image)
        
        Case "过程报告"
            Set objItem = TabWindow.InsertItem(lngIndex, "冰冻/特检报告", mobjWork_Pathol.GetModule(mtProRep).hWnd, Item.Image)
            
        Case "申请费用"
            Set objItem = TabWindow.InsertItem(lngIndex, "费用记录", mobjWork_His.GetModule(hmExpense).hWnd, Item.Image)
            
        Case "住院医嘱"
            Set objItem = TabWindow.InsertItem(lngIndex, "医嘱记录", mobjWork_His.GetModule(hmInAdvice).hWnd, Item.Image)
            
        Case "门诊医嘱"
            Set objItem = TabWindow.InsertItem(lngIndex, "医嘱记录", mobjWork_His.GetModule(hmOutAdvices).hWnd, Item.Image)
            
        Case "住院病历"
            Set objItem = TabWindow.InsertItem(lngIndex, "病历记录", mobjWork_His.GetModule(hmInEPRs).hWnd, Item.Image)
            
        Case "门诊病历"
            Set objItem = TabWindow.InsertItem(lngIndex, "病历记录", mobjWork_His.GetModule(hmOutEPRs).hWnd, Item.Image)
           
        Case "门诊电子病历", "住院电子病历"
            Set objItem = TabWindow.InsertItem(lngIndex, "电子病历", mobjWork_His.GetModule(hmEMR).hWnd, Item.Image)
              
        Case "排队叫号"
            Set objItem = TabWindow.InsertItem(lngIndex, "排队叫号", mobjQueue.hWnd, Item.Image)
            
        Case "影像采集", "报告填写"
            '这里不进行处理
    End Select
    
    Call RefreshModuleAdviceInf
    
    If Not objItem Is Nothing Then
        objItem.tag = Item.tag
        objItem.Selected = True
        
        Call TabWindow.RemoveItem(lngIndex + 1)
    End If
    
    mblnLoadSubFrom = False
Exit Sub
ErrHandle:
    If Not objItem Is Nothing Then
        If objItem.tag = "" Then
            Call TabWindow.RemoveItem(objItem.Index)
        End If
    End If
    
    mblnLoadSubFrom = False
End Sub

Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo ErrHandle
    Dim intStyle As Integer
    Dim blnVisible As Boolean
    Dim blnLargeIcon As Boolean
    Dim cbrControl As CommandBarControl

    
    Call ConfigSubForm(Item)

    If Not mblnInitOk Then Exit Sub
    
    Call ReSetModuleFontSize(mbytFontSize, IIf(mbytFontSize = 9, 0, 1))
    
    Call RefreshTabWindow
    
    '刷新排队叫号模块数据，如果已经启用并且是选择的排队叫号页面
    If Trim(Item.tag) = "排队叫号" Then
        Call RefreshPacsQueueData
    End If

    Call LockWindowUpdate(Me.hWnd)

    '有的菜单，只在工作模块显示的时候， 才显示
    Call CreateWorkModuleMenu
    
    If Val(ufgStudyList.CurKeyValue) <> 0 Then
        '显示可打印的诊疗单据:之所以即时加载,是为了使用F2热键
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    End If
    
    Call LockWindowUpdate(0)
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub GetRGB(ByVal lngColor As Long, lngR As Long, lngG As Long, lngB As Long)
    Dim lngMinVal As Long
    Dim lngMaxVal As Long
    
    lngMinVal = 80
    lngMaxVal = 225
    
    lngR = lngColor Mod 256
    
    If lngR <= lngMinVal Then
        lngR = lngMinVal
    ElseIf lngR > lngMaxVal Then
        lngR = lngMaxVal
    End If
    
    lngG = (Fix(lngColor \ 256)) Mod 256
 
    If lngG <= lngMinVal Then
        lngG = lngMinVal
    ElseIf lngG > lngMaxVal Then
        lngG = lngMaxVal
    End If
    
    lngB = Fix(lngColor \ 256 \ 256)
 
    If lngB <= lngMinVal Then
        lngB = lngMinVal
    ElseIf lngB > lngMaxVal Then
        lngB = lngMaxVal
    End If
End Sub


Private Sub timerCapture_Timer()
On Error GoTo ErrHandle

    timerCapture.Enabled = False
    
    '使用热键进行采集
    If GetKeyAliasEx(mCaptureMsg.lngVirtualKey) = mstrCaptureHot Then
        If mblnUseActivexCapture Then
            If Not mobjWork_ActiveVideo Is Nothing Then
                Call mobjWork_ActiveVideo.zlCaptureImg
            End If
        End If
    '使用热键进行后台采集
    ElseIf GetKeyAliasEx(mCaptureMsg.lngVirtualKey) = mstrCaptureAfterHot Then
        If mblnUseActivexCapture Then
            If Not mobjWork_ActiveVideo Is Nothing Then
                Call mobjWork_ActiveVideo.zlCaptureAfterImg
            End If
        End If
    
    '使用热键进行标记更新
    ElseIf GetKeyAliasEx(mCaptureMsg.lngVirtualKey) = mstrCaptureAfterTagHot Then
        If mblnUseActivexCapture Then
            If Not mobjWork_ActiveVideo Is Nothing Then
                Call mobjWork_ActiveVideo.zlUpdateAfterCaptureInfo
            End If
        End If
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume

End Sub

Private Sub timerOperHint_Timer()
On Error GoTo ErrHandle
    Dim i As Long
    Dim strText As String
    Dim dtOper As Date
    Dim lngColor1 As Long
    Dim lngR As Long, lngG As Long, lngB As Long
    
    If Not (mSysPar.lngEnregAfterTimeLen > 0 Or mSysPar.lngCheckInAfterTimeLen > 0 _
        Or mSysPar.lngStudyAfterTimeLen > 0 Or mSysPar.lngReportAfterTimeLen > 0 Or mSysPar.lngAuditAfterTimeLen > 0) Then
        timerOperHint.Enabled = False
        Exit Sub
    End If
    
    If ufgStudyList.DataGrid.Rows <= 1 Then Exit Sub
    
    '1表示颜色闪烁时显示为设置颜色更深一点的颜色，-1表示显示为设置颜色更浅一点的颜色，0表示显示设置的颜色
    If timerOperHint.tag = "1" Then
        timerOperHint.tag = "-1"
    ElseIf timerOperHint.tag = "-1" Then
        timerOperHint.tag = "0"
    ElseIf timerOperHint.tag = "0" Then
        timerOperHint.tag = "1"
    End If
    
    For i = ufgStudyList.DataGrid.TopRow To ufgStudyList.DataGrid.BottomRow
    
        dtOper = IIf(Nvl(ufgStudyList.Text(i, "采样时间")) = "", Now, ufgStudyList.Text(i, "采样时间"))
        strText = ufgStudyList.Text(i, "检查过程")
        
        Select Case strText
            Case "已登记"
                If mSysPar.lngEnregAfterTimeLen > 0 Then
                    dtOper = Nvl(ufgStudyList.Text(i, "申请时间"))
                    
                    Call SetFlickerColor(i, gdblColor已登记, dtOper, mSysPar.lngEnregAfterTimeLen)
                End If
            Case "已报到"
                If mSysPar.lngCheckInAfterTimeLen > 0 Then
                    Call SetFlickerColor(i, gdblColor已报到, dtOper, mSysPar.lngCheckInAfterTimeLen)
                End If
            Case "已检查"
                If mSysPar.lngStudyAfterTimeLen > 0 Then
                    Call SetFlickerColor(i, gdblColor已检查, dtOper, mSysPar.lngStudyAfterTimeLen)
                End If
            Case "已报告"
                If mSysPar.lngReportAfterTimeLen > 0 Then
                    Call SetFlickerColor(i, gdblColor已报告, dtOper, mSysPar.lngReportAfterTimeLen)
                End If
            Case "已审核"
                If mSysPar.lngAuditAfterTimeLen > 0 Then
                    Call SetFlickerColor(i, gdblColor已审核, dtOper, mSysPar.lngAuditAfterTimeLen)
                End If
        End Select
    Next i
ErrHandle:
End Sub

Private Sub SetFlickerColor(ByVal lngRow As Long, ByVal lngStateColor As Long, ByVal dtOper As Date, ByVal lngAfterTimeLen As Long)
'功能：设置已超时行的闪烁颜色
'参数：lngRow---当前行
'      lngStateColor---将设置的颜色
    Dim lngR As Long, lngG As Long, lngB As Long
    Dim lngPreStateColor As Long
    Dim lngNextStateColor As Long
    
    Call GetRGB(lngStateColor, lngR, lngG, lngB)
    lngNextStateColor = RGB(lngR - 30, lngG - 30, lngB - 30)
    lngPreStateColor = RGB(lngR + 30, lngG + 30, lngB + 30)
    
    If DateDiff("N", dtOper, Now) >= lngAfterTimeLen Then
        If timerOperHint.tag = "1" Then
            ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 0) = lngPreStateColor
        ElseIf timerOperHint.tag = "-1" Then
            ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 0) = lngStateColor
        ElseIf timerOperHint.tag = "0" Then
            ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 0) = lngNextStateColor
        End If
    End If
End Sub

Private Sub timerRefresh_Timer()
On Error GoTo ErrHandle
    '刷新病人列表
    If Not mblnInitOk Then Exit Sub
    If Not Me.Visible Then Exit Sub

    timerRefresh.Enabled = False
    
    Call RefreshList
    
    timerRefresh.Enabled = True
    
ErrHandle:
End Sub


Private Sub ChangeUser()
    Dim strPrivs As String
    Dim strUserID As String
    
    frmTwoUser.intDBState = mintChangeUserState
    frmTwoUser.strUserNameHIS = mstrUserNameHIS
    frmTwoUser.strUserIDHIS = mstrUserIDHIS
    frmTwoUser.Show 1, Me
    
    If frmTwoUser.blnOk = True Then
        If frmTwoUser.intDBState = 1 Then   '统一，则恢复成HIS原来的数据库联接和用户名
            mstrUserNameNew = mstrUserNameHIS
            mstrUserIDNew = mstrUserIDHIS
            mblnCnOracleIsHIS = True
            mintChangeUserState = 1
            Set gcnOracle = mcnOracleHIS
            InitCommon gcnOracle
            SetDbUser mstrUserIDHIS
            RegCheck
            Call GetUserInfo
            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
        ElseIf frmTwoUser.intDBState = 2 Then   '交换，则交换数据库联接
            '如果是使用新数据库联接，先检查权限
            mstrUserNameNew = frmTwoUser.strUserNameNew
            mstrUserIDNew = frmTwoUser.strUserIDNew
            mintChangeUserState = 2
            If frmTwoUser.blnCnOracleIsNew = True Then
                Set gcnOracle = frmTwoUser.cnOracle
                mblnCnOracleIsHIS = False
                
                '初始化zlComLib部件，确保GetPrivFunc提取的是正确的信息
                InitCommon gcnOracle
                RegCheck
                SetDbUser mstrUserIDNew
                
                '查找用户权限
                strPrivs = GetPrivFunc(100, 1291)       '影像采集工作站
                If strPrivs = "" Then
                    MsgBoxD Me, "你不具备使用“影像采集工作站”模块的权限！"
                    
                    '切换回原来的用户
                    Set gcnOracle = mcnOracleHIS
                    
                    InitCommon gcnOracle
                    RegCheck
                    SetDbUser mstrUserIDHIS
                
                    mstrUserNameNew = mstrUserNameHIS
                    mstrUserIDNew = mstrUserIDHIS
                    mblnCnOracleIsHIS = True
                    mintChangeUserState = 1
                End If
                
                strPrivs = GetPrivFunc(100, 1258)       '诊疗报告管理
                If strPrivs = "" Then
                    MsgBoxD Me, "你不具备使用“诊疗报告”模块的权限！"
                    
                    '切换回原来的用户
                    Set gcnOracle = mcnOracleHIS
                    
                    InitCommon gcnOracle
                    RegCheck
                    SetDbUser mstrUserIDHIS
                    
                    mstrUserNameNew = mstrUserNameHIS
                    mstrUserIDNew = mstrUserIDHIS
                    mblnCnOracleIsHIS = True
                    mintChangeUserState = 1
                End If
            Else
                Set gcnOracle = mcnOracleHIS
                
                InitCommon gcnOracle
                RegCheck
                SetDbUser mstrUserIDHIS
                
                mblnCnOracleIsHIS = True
            End If
            
            Call GetUserInfo
            Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
        End If
    End If
    
    If mblnCnOracleIsHIS Then
        Me.stbThis.Panels(4).Text = "报告医生：" & mstrUserNameHIS & "   检查医生：" & mstrUserNameNew
    Else
        Me.stbThis.Panels(4).Text = "报告医生：" & mstrUserNameNew & "   检查医生：" & mstrUserNameHIS
    End If
End Sub

Private Sub SwitchUser()
    Dim strPrivs As String
    
    frmSwitchUser.Show 1, Me
    
    If frmSwitchUser.blnOk Then
        '如果是使用新数据库联接，先检查权限
        mstrUserNameNew = frmSwitchUser.strUserNameNew
        mstrUserIDNew = frmSwitchUser.strUserIDNew

        If frmSwitchUser.blnCnOracleIsNew = True Then
            Set gcnOracle = frmSwitchUser.cnOracle
            mblnCnOracleIsHIS = False
            
            '初始化zlComLib部件，确保GetPrivFunc提取的是正确的信息
            zl9comlib.CloseWindows
            InitCommon gcnOracle
            RegCheck
            SetDbUser mstrUserIDNew
            
            '查找用户权限
            strPrivs = GetPrivFunc(100, 1294)       '影像病理工作站
            If strPrivs = "" Then
                MsgBoxD Me, "你不具备使用“影像病理工作站”模块的权限！"
                
                '切换回原来的用户
                Set gcnOracle = mcnOracleHIS
                
                InitCommon gcnOracle
                RegCheck
                SetDbUser mstrUserIDHIS
            
                mstrUserNameNew = mstrUserNameHIS
                mstrUserIDNew = mstrUserIDHIS
                mblnCnOracleIsHIS = True
            End If
        Else
            Set gcnOracle = mcnOracleHIS
            
            zl9comlib.CloseWindows
            InitCommon gcnOracle
            RegCheck
            SetDbUser mstrUserIDHIS
            
            mblnCnOracleIsHIS = True
        End If
    End If
    
    Call GetUserInfo
    Call gobjRichEPR.InitRichEPR(gcnOracle, gfrmMain, glngSys, False)
    
    If mblnCnOracleIsHIS Then
        Me.stbThis.Panels(4).Text = "报告医生：" & mstrUserNameHIS & "   检查医生：" & mstrUserNameHIS
    Else
        Me.stbThis.Panels(4).Text = "报告医生：" & mstrUserNameNew & "   检查医生：" & mstrUserNameNew
    End If
End Sub

Private Sub SeekNextPati(ByVal blnFirst As Boolean, ByVal strCardName As String, _
    ByVal strFilter As String, Optional blnIsReSeek As Boolean = False)
'------------------------------------------------
'功能：在病人列表中定位指定的记录
'参数： blnFirst -- 是否第一次查找
'返回：无，直接在病人列表中定位
'------------------------------------------------
    Dim i As Long
    Dim intB As Integer
    Dim lngEndRow As Long
    Dim lngSelRow As Long
    Dim strTemp As String
    Dim lngRowIndex As Long

    
    '如果没有记录，则退出
    If ufgStudyList.ShowingRowCount <= 0 Then Exit Sub

    intB = 0
    
    If Not blnFirst Then
        intB = ufgStudyList.DataGrid.Row + 1
        If intB >= ufgStudyList.DataGrid.Rows Then intB = 1
    End If
    
    lngSelRow = ufgStudyList.DataGrid.Row
    lngEndRow = ufgStudyList.DataGrid.Rows - 1

continue1:

    Select Case strCardName
        Case "标识号", "住院号", "门诊号"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("标识号"), False, False)
            
        Case "单据号"
            strTemp = ""
            
            '补全单据号
            If Len(Trim(strFilter)) > 0 Then
                If Len(Trim(strFilter)) < 8 And Not IsNumeric(Trim(strFilter)) Then
                    strTemp = GetFullNO(0, 0)
                    strTemp = Mid(strTemp, 1, Len(strTemp) - Len(strFilter)) & strFilter
                Else
                    strTemp = GetFullNO(Nvl(strFilter, 0), 0)
                End If
            End If
            
            PatiIdentify.Text = strTemp
            
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strTemp, intB, ufgStudyList.GetColIndex("NO"), False, False)
            
        Case GetStudyNumberDisplayName
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex(GetStudyNumberDisplayName), False, False)
            
        Case "姓名", "姓 名", "姓  名", "姓   名"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("姓名"), False, False)
            
            '如果没有找到，则判断输入是否为全字母，如果是，则使用拼音查找
            If lngRowIndex <= 0 And LenB(StrConv(strFilter, vbFromUnicode)) = Len(strFilter) Then
                For i = intB To lngEndRow
                    If zlCommFun.SpellCode(Nvl(ufgStudyList.Text(i, "姓名"), "")) Like UCase(strFilter) & "*" Then
                        lngRowIndex = i
                        Exit For
                    End If
                Next i
            End If
            
        Case "就诊卡", "就诊卡号"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("就诊卡号"), False, False)
            
        Case "身份证号", "身份证"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("身份证号"), False, False)
        
        Case "医嘱ID"
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("医嘱ID"), False, False)
            
        Case Else
            lngRowIndex = ufgStudyList.DataGrid.FindRow(strFilter, intB, ufgStudyList.GetColIndex("病人ID"), False, True)
            
    End Select


    If lngRowIndex > 0 Then
        PatiIdentify.tag = PatiIdentify.Text
        
        On Error GoTo errContinue1

            ufgStudyList.DataGrid.Row = lngRowIndex

            If ufgStudyList.DataGrid.TopRow > ufgStudyList.DataGrid.Row Then ufgStudyList.DataGrid.TopRow = ufgStudyList.DataGrid.Row
            If ufgStudyList.DataGrid.BottomRow - 1 < ufgStudyList.DataGrid.Row Then
                ufgStudyList.DataGrid.TopRow = ufgStudyList.DataGrid.TopRow + (ufgStudyList.DataGrid.Row - ufgStudyList.DataGrid.BottomRow) + 1
            End If

            If lngSelRow = ufgStudyList.DataGrid.Row Then
                '如果该检查为已登记状态，则执行报道操作
                If ufgStudyList.CurText("检查过程") = "已登记" Then
                    Call Menu_Manage_报到
                End If
            End If
        
errContinue1:
        
        Exit Sub
    End If
    
    '如果没有找到数据再执行刷新列表，然后再定位，这样避免每次定位都要刷新列表
    If lngRowIndex <= 0 Then
        If blnIsReSeek Then
        
            Call RefreshList
            blnIsReSeek = False
            
            GoTo continue1
        
        End If
    End If
    
    If intB > 1 Then
        lngEndRow = intB
        intB = 1
        
        GoTo continue1
    End If
    
    ufgStudyList.DataGrid.Row = -1
End Sub

Private Sub Menu_Manage_随访()
On Error GoTo ErrHandle
    Dim strReview As String
    Dim strDeptName As String

    If mListAdviceInf.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    strDeptName = Split(mstrCur科室, "-")(1)
    If frmReview.ShowMe(mListAdviceInf.lngAdviceId, mListAdviceInf.lngSendNo, Me, strDeptName, strReview) = True Then
        ufgStudyList.CurText("随访描述") = strReview
        Call ufgStudyList.UpdateSourceData(mListAdviceInf.lngAdviceId, "随访描述", strReview)
    End If

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_报告发放()
'报告发放
On Error GoTo ErrHandle
    Dim strSql As String

    If mListAdviceInf.lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    With ufgStudyList
        If mrtReportType = 报告文档编辑器 Then
            Call mobjWork_Report.Menu_Manage_报告发放(mListAdviceInf.lngAdviceId, IIf(mobjWork_Report.GetReportReleaseState(mcurAdviceInf.lngAdviceId) > 1, 0, 1))
        Else
            strSql = "Zl_影像报告发放(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(strSql, "报告发放")
            
            .CurText("报告发放") = IIf(Nvl(.CurText("报告发放"), "0") = "0", "1", "0")
            Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "报告发放", IIf(.CurText("报告发放") = "1", "1", "0"))
        End If
    End With
    
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_胶片发放()
'胶片发放
On Error GoTo ErrHandle
    Dim strSql As String

    With ufgStudyList

        If mListAdviceInf.lngAdviceId <= 0 Then
            MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
            Exit Sub
        End If
        
        strSql = "Zl_影像胶片发放(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.姓名 & "')"
        Call zlDatabase.ExecuteProcedure(strSql, "胶片发放")
        
        .CurText("胶片发放") = IIf(Nvl(.CurText("胶片发放"), "0") = "0", "1", "0")
        Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "发放胶片", IIf(Nvl(.CurText("胶片发放"), "0") = "1", "1", "0"))
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_报告胶片同时发放()
'报告胶片同时发放
On Error GoTo ErrHandle
    Dim strSql As String
    
    With ufgStudyList
        
        If mListAdviceInf.lngAdviceId <= 0 Then
            MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
            Exit Sub
        End If
        
        If mrtReportType = 报告文档编辑器 Then
            If mobjWork_Report.GetReportReleaseState(mcurAdviceInf.lngAdviceId) = 3 And Nvl(.CurText("胶片发放"), "0") = "1" Then
                Call mobjWork_Report.Menu_Manage_报告发放(mListAdviceInf.lngAdviceId, 0)
                
                strSql = "Zl_影像胶片发放(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "胶片发放")
                
                .CurText("胶片发放") = "0"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "发放胶片", IIf(Nvl(.CurText("胶片发放"), "0") = "1", "1", "0"))
            Else
                Call mobjWork_Report.Menu_Manage_报告发放(mListAdviceInf.lngAdviceId, 1)
                
                strSql = "Zl_影像胶片发放(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "胶片发放")
                
                .CurText("胶片发放") = "1"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "发放胶片", IIf(Nvl(.CurText("胶片发放"), "0") = "1", "1", "0"))
            End If
        Else
            If Nvl(.CurText("报告发放"), "0") = "1" And Nvl(.CurText("胶片发放"), "0") = "1" Then
                strSql = "Zl_影像报告发放(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "报告发放")
                
                .CurText("报告发放") = "0"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "报告发放", IIf(Nvl(.CurText("报告发放"), "0") = "1", "1", "0"))
            
                strSql = "Zl_影像胶片发放(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "胶片发放")
                
                .CurText("胶片发放") = "0"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "发放胶片", IIf(Nvl(.CurText("胶片发放"), "0") = "1", "1", "0"))
            Else
                strSql = "Zl_影像报告发放(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "报告发放")
                
                .CurText("报告发放") = "1"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "报告发放", IIf(Nvl(.CurText("报告发放"), "0") = "1", "1", "0"))
            
                strSql = "Zl_影像胶片发放(" & mListAdviceInf.lngAdviceId & ",'" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(strSql, "胶片发放")
                
                .CurText("胶片发放") = "1"
                Call .UpdateSourceData(mListAdviceInf.lngAdviceId, "发放胶片", IIf(Nvl(.CurText("胶片发放"), "0") = "1", "1", "0"))
            End If
        End If
        
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Menu_Manage_ReportExecutor()
    Dim strSql As String
    
    Dim strRPTExecutor As String
On Error GoTo ErrHandle
    strRPTExecutor = frmSelectRPTExecutor.GetRPTExecutor(mlngCur科室ID, Me, mstrRPTExecutor)
    
    If strRPTExecutor <> "" Then
        '更新报告人
        strSql = "ZL_影像报告保存_更新报告人(" & mcurAdviceInf.lngAdviceId & ",'" & strRPTExecutor & "')"
        Call zlDatabase.ExecuteProcedure(CStr(strSql), "更新报告人")
        
        '刷新对应检查的报告人
        mstrRPTExecutor = strRPTExecutor
        ufgStudyList.CurText("报告人") = strRPTExecutor
        
        If Not mobjWork_Report Is Nothing And mrtReportType = 报告文档编辑器 Then Call mobjWork_Report.SetDocCreator(mstrRPTExecutor)
        
        stbThis.Panels(4).Text = "报告医生：" & strRPTExecutor & "   检查医生：" & Split(stbThis.Panels(4).Text, "检查医生：")(1)
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub timerVideoEvent_Timer()
On Error GoTo ErrHandle
    timerVideoEvent.Enabled = False
    
    Call DoOnStateChange(mVideoEventInf.vetEventType, mVideoEventInf.lngAdviceId, mVideoEventInf.lngSendNo, mVideoEventInf.strOtherInf)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume

End Sub

Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
On Error GoTo ErrHandle
    PatiIdentify.Text = ""  '切换Item时，要将输入框清空
    If cbrdock.FindControl(, ID_查找方式) Is Nothing Then Exit Sub
    '在快速工具栏点击定位和查找时，更新刷卡控件IDKindStr时，会出发ItemClick事件，
    '导致无法分别记录定位和查找字段信息，所以用此变量标记，为true时不更新mstrLocateWay和mstrFindWay
    If mblnAssignment Then Exit Sub
    
    If cbrdock.FindControl(, ID_查找方式).IconId = 3 Then
        mstrLocateWay = objCard.名称
    Else
        mstrFindWay = objCard.名称
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Public Sub StartReadCard()
'开始读卡
    Dim lngPatientID As Long
    Dim strCurCardName As String
    
    If cbrdock.FindControl(, ID_查找方式).IconId = 3 Then
        strCurCardName = mstrLocateWay
    Else
        strCurCardName = mstrFindWay
    End If
    
    If PatiIdentify.GetCurCard.是否刷卡 Then
        Call OnFilterRead(strCurCardName, PatiIdentify.Text, IIf(lngPatientID > 0, lngPatientID, ""))
    Else
        If PatiIdentify.GetCurCard.接口序号 > 0 Then
            Call mobjSquareCard.zlGetPatiID(PatiIdentify.GetCurCard.接口序号, PatiIdentify.Text, , lngPatientID)
            
            Call OnFilterRead(strCurCardName, PatiIdentify.Text, IIf(lngPatientID > 0, lngPatientID, ""))
        Else
            Call OnFilterRead(strCurCardName, PatiIdentify.Text, "")
        End If
    End If
End Sub

Private Sub OnFilterRead(ByVal strCardName As String, ByVal strFilter As String, ByVal strPatientId As String)
'开始查找数据
On Error GoTo ErrHandle
    If cbrdock.FindControl(, ID_查找方式).IconId = 3 Then
        '定位检查数据
        If strPatientId <> "" Then
            Call SeekNextPati(PatiIdentify.tag <> PatiIdentify.Text, "病人ID", strPatientId, True)
        Else
            Call SeekNextPati(PatiIdentify.tag <> PatiIdentify.Text, strCardName, strFilter, True)
        End If
    Else
        '查找检查数据
        If strPatientId <> "" Then
            Call subRefreshFilterCondition("病人ID", strPatientId)
        Else
            Call subRefreshFilterCondition(strCardName, strFilter)
        End If
        
        Call RefreshList
    End If
    
    Call PatiIdentify.SetFocus
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function GetStudyNumberDisplayName() As String
'获取检查号码显示名称
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "病理号", "检查号")
End Function




Private Sub ufgStudyList_OnBindFilter(strBindFilter As String, strCloneFilter As String)
    strBindFilter = " 相关ID=NULL"
    strCloneFilter = " 相关ID<>NULL"
End Sub

Private Sub ufgStudyList_OnCellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHandle
    frmDegreeCard.ShowMe Val(ufgStudyList.Text(Row, "病人ID")), Val(ufgStudyList.Text(Row, "主页ID")), Me
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnColFormartChange()
On Error GoTo ErrHandle
    Call zlDatabase.SetPara("检查列表", ufgStudyList.GetColsString(ufgStudyList), glngSys, mlngModule)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnColsNameReSet()
On Error GoTo ErrHandle
    '列头恢复默认后重新加载病人列表
    Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnDblClick()
On Error GoTo ErrHandle
    If Val(ufgStudyList.CurKeyValue) <> 0 Then
        '双击病人检查列表时，如果病人检查状态为 已拒绝，目前不做任何处理
        If Nvl(ufgStudyList.CurText("检查过程")) = "已拒绝" Then Exit Sub
        
        Select Case Val(ufgStudyList.CurText("检查状态"))
            Case 1, 0
                Call Menu_Manage_报到
            Case 2, 3               '双击打开书写报告,报告打开时跟据设定是否打开观片站
                Call Menu_RichEPR(conMenu_Edit_Modify)
            Case -1, 4, 5               '双击修订报告,报告打开时跟据设定是否打开观片站
                Call Menu_RichEPR(conMenu_Edit_Audit)
            Case 6                  '查阅
                Call Menu_RichEPR(conMenu_File_Open)
        End Select
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnFilterRowData(rsData As ADODB.Recordset, rsClone As ADODB.Recordset, blnFilterOut As Boolean)
On Error GoTo ErrHandle
    '判断是否已经收费
    '"病人医嘱发送.记录性质"--- 1是收费的，2是记帐的。
    
    '通过"病人医嘱发送.计费状态"直接判断,原有值：-1-无须计费;0-未计费;1-已计费，对于记帐单（包括门诊记帐单），保持原有值不变。
    '对于收费单的发送记录，增加两种状态：2-部分收费，3-全部收费
    
    '没有对应费用的医嘱有两种情况，一种是"-1-无须计费"，即没有设置收费对照，一种是"0-未计费"，即虽然设置了收费对照，但设置为发送后手工计费，即在医技科室去生成。
    '"1-已计费"就是发送时生成了费用的。但生成了费用单据不表示收费了，生成可能是记帐划价单，或收费划价单，其中收费划价单就多两种状态。
    '"2-部分收费"表示部分收费和部分退费的情况，反正没收得完。
    
    '已收费显示状态：已收费；无费用；未收费：
    '未收费----
    '1、主医嘱是收费单的，满足以下条件算未收费
    '   (1)有一条主医嘱和部位医嘱的 计费状态 in (1,2)算未收费 ------“记录性质=1 and 计费状态 in (1,2)”
    '已收费：
    '1、主医嘱是记账的算收费-------“记录性质=2”
    '2、主医嘱是收费单的，满足以下条件算收费
    '   (1)排除未收费后，有一条主医嘱和部位医嘱的 计费状态 =3 算收费-----“记录性质=1 and 计费状态 = 3”
    '无费用
    '1、主医嘱是收费单的，满足以下条件算无费用
    '   (1)所有主医嘱和部位医嘱的 计费状态 in (-1,0)算无费用 ------“记录性质=1 and 计费状态 in (-1,0)”
    
    Dim intNum As Integer
    
    intNum = 0
    If Nvl(rsData!相关ID) = "" Then '说明是主医嘱
        If mstrcmd部位 <> "" Then
            '过滤出子医嘱
            rsClone.Filter = "相关ID = " & Nvl(rsData!医嘱ID)
            
            Do While rsClone.EOF = False
                '当此部位没有勾选时，主医嘱不需要显示到列表中
                If Nvl(rsClone!标本部位) <> "" Then
                    If InStr(mstrcmd部位, Nvl(rsClone!标本部位)) > 0 Then '子医嘱有部位在分组中，主医嘱需要显示
                        intNum = 1
                        Exit Do
                    End If
                End If
                
                rsClone.MoveNext
            Loop
            
            If rsClone.RecordCount > 0 And intNum <= 0 Then '说明所有子医嘱的部位都没有选择，主医嘱不需要显示到列表中
                blnFilterOut = True
                Exit Sub
            End If
        End If
    End If
    
    ' intCharged  '0--未收费；1--已收费；2--无费用
    
    If Nvl(rsData!相关ID) <> "" Then
        '相关id不为空时，说明书部位医嘱，不需要显示到列表中
        blnFilterOut = True
        Exit Sub
    End If

    mlngTempCharged = 2 '无费用
    
    If Nvl(rsData!记录性质, 2) = 2 Then
        '判断病人结算模式，如果病人结算模式为1，则表示先诊疗后结算即记账病人，此时费用状态需要显示为“记”
        If Val(Nvl(rsData!结算模式)) = 1 Then
            mlngTempCharged = 5         '门诊记账病人显示为“记”
        Else
            '住院登记的病人，如果没有计费，则归为无费用
            If Nvl(rsData!计费状态, -1) = 0 Then
    
                    rsClone.Filter = "相关ID = " & Nvl(rsData!医嘱ID)
                    Do While rsClone.EOF = False
                        If Nvl(rsClone!计费状态, -1) = 1 Then
                            '如果是记账医嘱，凡是已计费和全部收费的，表示为已收费
                            mlngTempCharged = 5      '已记账
                            
                        ElseIf Nvl(rsClone!计费状态, -1) = 3 Then
                            mlngTempCharged = 1      '已收费
       
                        ElseIf Nvl(rsClone!计费状态, -1) = 2 Then
                            mlngTempCharged = 0  '未收费
                            
                            Exit Do
                        End If
                        rsClone.MoveNext
                    Loop
                    
            Else
'                mlngTempCharged = 1  '已收费
                mlngTempCharged = 5         '住院记账病人显示为“记”
            End If
            
        End If
    Else
        If Nvl(rsData!计费状态, -1) = 1 Or Nvl(rsData!计费状态, -1) = 2 Then
            mlngTempCharged = 0      '未收费
        Else        '主医嘱的计费状态是 -1,0,3  （3--已收费；-1，0--无费用）
            '查询主医嘱未计费或者已经收费了，还要查部位医嘱的收费情况，所有医嘱都已经收费，才算是收费
            
            '如果主费用是已收费的，先记录成已收费
            If Nvl(rsData!计费状态, -1) = 3 Then
                mlngTempCharged = 1      '已收费
            End If
            
            rsClone.Filter = "相关ID = " & Nvl(rsData!医嘱ID)
            Do While rsClone.EOF = False
                If Nvl(rsClone!计费状态, -1) = 1 Or Nvl(rsClone!计费状态, -1) = 2 Then
                    mlngTempCharged = 0      '未收费

                    Exit Do
                ElseIf Nvl(rsClone!计费状态, -1) = 3 Then
                    mlngTempCharged = 1      '已收费
                End If

                rsClone.MoveNext
            Loop
            
'            '计费状态：-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费，对收费单据多两种状态:2-部分收费，3-全部收费
'            rsClone.Filter = "相关ID = " & Nvl(rsData!医嘱ID) & " and 计费状态=1 and 计费状态=2"
'            If rsClone.RecordCount > 0 Then
'                mlngTempCharged = 0 '未收费
'            Else
'                rsClone.Filter = "相关ID = " & Nvl(rsData!医嘱ID) & " and 计费状态=3"
'                If rsClone.RecordCount > 0 Then mlngTempCharged = 1 '已收费
'            End If
            
        End If
    End If

    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If Nvl(rsData!补费) > 0 Then mlngTempCharged = 4 '需要补费，需补费的检查也是未收费的检查
    End If
    
    If Nvl(rsData!相关ID) = "" And ((mblncmd已缴 = True And mlngTempCharged = 1) Or (mblncmd未缴 = True And (mlngTempCharged = 0 Or mlngTempCharged = 4)) _
        Or (mblncmd无费 = True And mlngTempCharged = 2) Or (mblncmd补缴 = True And mlngTempCharged = 4) Or (mblncmd记账 And mlngTempCharged = 5) _
        Or (mblncmd已缴 = False And mblncmd未缴 = False And mblncmd补缴 = False And mblncmd无费 = False And mblncmd记账 = False)) Then
        blnFilterOut = False
        
        Call RowDataConvert(rsData)
    Else
        blnFilterOut = True
    End If
ErrHandle:
End Sub



Private Sub RowDataConvert(rsData As ADODB.Recordset)
On Error Resume Next
    Dim rsBaby As ADODB.Recordset
    Dim intTxtLen As Long
    
    '如果该数据要显示，则需要转换数据中的部分值
    rsData!申请单 = IIf(Nvl(rsData!申请单) = "", "", "已扫描")
    If mrtReportType = 报告文档编辑器 Then
        rsData!检查过程 = IIf(Val(Nvl(rsData!执行状态)) = 2, "已拒绝", Decode(Val(Nvl(rsData!检查状态, 0)), -1, "已驳回", 0, "已登记", 1, "已登记", _
                                                                                2, "已报到", 3, "已检查", 4, "已报告", 5, "已审核", "已完成"))
    Else
        rsData!检查过程 = IIf(Val(Nvl(rsData!执行状态)) = 2, "已拒绝", Decode(Val(Nvl(rsData!检查状态, 0)), -1, "已驳回", 0, "已登记", 1, "已登记", _
                                                                                2, IIf(Nvl(rsData!报告操作) <> "", "处理中", _
                                                                                        IIf(Nvl(rsData!报告人) = "", "已报到", "报告中")), _
                                                                                3, IIf(Nvl(rsData!报告操作) <> "", "处理中", _
                                                                                        IIf(Nvl(rsData!报告人) = "", "已检查", "报告中")), _
                                                                                4, IIf(Nvl(rsData!报告操作) <> "", "处理中", _
                                                                                        IIf(Nvl(rsData!复核人) <> "", "审核中", "已报告")), _
                                                                                5, "已审核", "已完成"))
    End If
                                                                                
    If Nvl(rsData!婴儿) <> 0 Then
        gstrSQL = "Select Nvl(A.婴儿姓名, B.姓名 || '之子' || Trim(To_Char(A.序号, '9'))) As 婴儿姓名, 婴儿性别, 出生时间" & vbNewLine & _
                    "From 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
                    "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id And A.序号 = [3]"
        
        Set rsBaby = zlDatabase.OpenSQLRecord(gstrSQL, "提取婴儿信息", CLng(rsData!病人ID), CLng(Nvl(rsData!主页ID, 0)), CLng(rsData!婴儿))
        
        If Not rsBaby.EOF Then
            rsData!姓名 = rsBaby!婴儿姓名
            rsData!性别 = Nvl(rsBaby!婴儿性别)
            rsData!年龄 = Nvl(rsBaby!出生时间)
        End If
        
        rsData!婴儿 = "是"
    Else
        rsData!婴儿 = ""
    End If
    
    
    If InStr(Nvl(rsData!医嘱内容), ":") > 0 Then '新的模式保存在医嘱内容中信息是 名称,执行标记:部位(方法,方法),部位---
        rsData!部位方法 = Split(Nvl(rsData!医嘱内容), ":")(1)
        rsData!医嘱内容 = Split(Nvl(rsData!医嘱内容), ":")(0)
    End If
    

    rsData!报告打印 = IIf(Val(Nvl(rsData!报告打印)) = 1, "√", "")
    rsData!报告发放 = IIf(Nvl(rsData!报告发放) = "0", "0", "1")
    
    If mlngModule = G_LNG_PATHSTATION_MODULE Then   '只有医技才具备胶片打印和胶片发放情况
        rsData!胶片打印 = IIf(Val(Nvl(rsData!胶片打印)) = 1, "√", "")
        rsData!胶片发放 = IIf(Val(Nvl(rsData!胶片发放)) = 1, "1", "0")
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then    '获取病理检查执行状态
        rsData!病理执行状态 = GetPatholExecuteState(rsData)
    End If
    
    
    If Val(Nvl(rsData!路径)) = 1 Then
        rsData!路径 = " "
    Else
        rsData!路径 = ""
    End If
    
    
    If Val(Nvl(rsData!紧急)) <> 0 Then
        rsData!紧急 = " "
    Else
        rsData!紧急 = ""
    End If
    
    '病人来源
    If Val(Nvl(rsData!来源)) = 1 Then
        rsData!来源 = "门"
    ElseIf Val(Nvl(rsData!来源)) = 2 Then
        rsData!来源 = "住"
    ElseIf Val(Nvl(rsData!来源)) = 3 Then
        rsData!来源 = "外"
    ElseIf Val(Nvl(rsData!来源)) = 4 Then
        rsData!来源 = "体检"
    Else
        rsData!来源 = ""
    End If
    
    If mlngTempCharged = 0 Then         '未收费
        rsData!收费 = ""
    ElseIf mlngTempCharged = 1 Then     '已收费
        rsData!收费 = " "
    ElseIf mlngTempCharged = 2 Then     '无费用
        rsData!收费 = "  "
    ElseIf mlngTempCharged = 5 Then     '记帐费用
        rsData!收费 = "    "
    Else                                '需补费
        rsData!收费 = "   "
    End If
    
    If Val(Nvl(rsData!阳性)) <> 0 Then
        rsData!阳性 = " "  ' 做排序用
    Else
        rsData!阳性 = ""
    End If
    
    If Val(Nvl(rsData!危急)) <> 0 Then
        rsData!危急 = " "
    Else
        rsData!危急 = ""
    End If
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        intTxtLen = Len(mSysPar.strImageLevel) - Len(Replace(mSysPar.strImageLevel, ",", "")) + 1

        If Trim(Val(Nvl(rsData!影像质量))) <> 0 Then
            If Val(rsData!影像质量) <= intTxtLen Then
                If Trim(Split(mSysPar.strImageLevel, ",")(Val(rsData!影像质量) - 1)) <> "" Then
                    rsData!影像质量 = Trim(Split(mSysPar.strImageLevel, ",")(Val(rsData!影像质量) - 1))
                Else
                    rsData!影像质量 = "未设置"
                End If

            Else
                rsData!影像质量 = "无效等级"
            End If
        End If
    End If


    intTxtLen = Len(mSysPar.strReportLevel) - Len(Replace(mSysPar.strReportLevel, ",", "")) + 1

    If Trim(Val(Nvl(rsData!报告质量))) <> 0 Then
        If Val(rsData!报告质量) <= intTxtLen Then
            If Trim(Split(mSysPar.strReportLevel, ",")(Val(rsData!报告质量) - 1)) <> "" Then
                rsData!报告质量 = Trim(Split(mSysPar.strReportLevel, ",")(Val(rsData!报告质量) - 1))
            Else
                rsData!报告质量 = "未设置"
            End If

        Else
            rsData!报告质量 = "无效等级"
        End If
    End If
    
    If Val(Nvl(rsData!电子胶片)) <> 0 Then
        rsData!电子胶片 = "已生成"
    Else
        rsData!电子胶片 = "未生成"
    End If
err.Clear
End Sub


Private Sub ufgStudyList_OnOrderChange(ByVal lngCol As Long, ByVal lngOrder As Integer, blnCustom As Boolean)
'保存当前的排序信息
On Error GoTo ErrHandle
    mlngSortCol = lngCol
    mintSortOrder = lngOrder
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ufgStudyList_OnRefreshRowData(rsBind As ADODB.Recordset, ByVal lngRow As Long)
On Error GoTo ErrHandle
    Dim strTag As String
    Dim strTemp As String
    Dim i As Long
    
    Dim strPatientType As String
    
    ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 0) = &H8000000F '&HE0E0E0
    
    For i = 0 To ufgStudyList.DataGrid.Cols - 1
        Select Case ufgStudyList.DataGrid.TextMatrix(0, i)
                
            Case "路径"
                If ufgStudyList.Text(lngRow, "路径") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("路径").Picture
                End If
                
            Case "紧急"
                If ufgStudyList.Text(lngRow, "紧急") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("紧急").Picture
                End If
        
            Case "来源"
                strTag = Decode(ufgStudyList.Text(lngRow, "来源"), "门", 1, "住", 2, "外", 3, 4)
                ufgStudyList.DataGrid.Cell(flexcpData, lngRow, i) = strTag
                
                If ufgStudyList.Text(lngRow, "来源") = "住" Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("住院").Picture
                End If
                
            Case "收费" 'TODO:病理还需要考虑补缴费用的情况
                If ufgStudyList.Text(lngRow, "收费") = "" Then          '未收费
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("欠费").Picture
                ElseIf ufgStudyList.Text(lngRow, "收费") = " " Then     '已收费
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("收费").Picture
                ElseIf ufgStudyList.Text(lngRow, "收费") = "   " Then   '补费
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("补费").Picture
                ElseIf ufgStudyList.Text(lngRow, "收费") = "    " Then  '记费
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("记费").Picture
                Else '无费用("  ")
                    '无费用不显示图标
                End If
                
            Case "危急"
                If ufgStudyList.Text(lngRow, "危急") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("危急").Picture
                End If
                
            Case "阳性"
                If ufgStudyList.Text(lngRow, "阳性") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("阳性").Picture
                End If
                
            Case "姓名" '如果为绿色通道，则需要在姓名面前添加图标
                If Val(ufgStudyList.Text(lngRow, "绿色通道")) <> 0 Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("绿色通道").Picture
                End If
                
            Case GetStudyNumberDisplayName  '检查号或者病理号
                If ufgStudyList.Text(lngRow, "检查UID") <> "" Then
                    '病理系统中，检查列表中的检查号显示为病理号
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages(IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "病理", "影像")).Picture
                End If
            
            Case "检查技师"
                If Val(ufgStudyList.Text(lngRow, "是否技师确认")) = 1 Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("检查技师").Picture
                End If
                
            Case "检查过程"
                '根据检查过程，设置不同的颜色
                If mSysPar.lngListColorMark = 0 Then
                    ufgStudyList.DataGrid.Cell(flexcpForeColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = getRowColor(lngRow)
                Else
                    ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = getRowColor(lngRow)
                End If
        End Select
        
    Next i
    
    If mSysPar.blnNameColColorCfg Then
        rsBind.Filter = "医嘱id=" & Val(ufgStudyList.Text(lngRow, "医嘱ID"))
        
        If rsBind.RecordCount > 0 Then strPatientType = Nvl(rsBind!病人类型)
        
        If mstrDefaultPatientType = strPatientType Then
            If mSysPar.blnOrdinaryNameColColorCfg Then
                Call NameColorCfg(lngRow, strPatientType)
            End If
        Else
            Call NameColorCfg(lngRow, strPatientType)
        End If
    End If
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        i = ufgStudyList.GetColIndex("质量")
        
        If ufgStudyList.Text(lngRow, "质量") = "符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbGreen
        If ufgStudyList.Text(lngRow, "质量") = "基本符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbYellow
        If ufgStudyList.Text(lngRow, "质量") = "不符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbRed
    Else
        i = ufgStudyList.GetColIndex("符合情况")
        
        If ufgStudyList.Text(lngRow, "符合情况") = "符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbGreen
        If ufgStudyList.Text(lngRow, "符合情况") = "基本符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbYellow
        If ufgStudyList.Text(lngRow, "符合情况") = "不符合" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, i) = vbRed
    End If
    
ErrHandle:
    Exit Sub
End Sub

Private Sub NameColorCfg(ByVal lngRow As Long, ByVal strPatientType As String)
    Dim lngR1 As Long, lngG1 As Long, lngB1 As Long
    Dim lngR2 As Long, lngG2 As Long, lngB2 As Long
    Dim lngPatiColor As Long, lngForeColor As Long
        
    lngPatiColor = zlDatabase.GetPatiColor(strPatientType)
    lngForeColor = getRowColor(lngRow)
    
    Call GetRGB(lngPatiColor, lngR1, lngG1, lngB1)
    Call GetRGB(lngForeColor, lngR2, lngG2, lngB2)
    
    '当病人类型颜色与列表行的前景色相近时，为了避免字体看不清，需对前景色进行适当处理
    If Abs(lngR1 - lngR2) < 30 Then
        lngR2 = lngR1 - 30
    End If
    
    If Abs(lngG1 - lngG2) < 30 Then
        lngG2 = lngG1 - 30
    End If
    
    If Abs(lngB1 - lngB2) < 30 Then
        lngB2 = lngG1 - 30
    End If
    
    lngForeColor = RGB(lngR2, lngG2, lngB2)
    
    ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, ufgStudyList.GetColIndex("姓名")) = lngPatiColor
    ufgStudyList.DataGrid.Cell(flexcpForeColor, lngRow, ufgStudyList.GetColIndex("姓名")) = lngForeColor
End Sub

Private Sub ufgStudyList_OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'弹出右键菜单
On Error GoTo ErrHandle
    If Button = 2 Then
        Dim control As CommandBarControl, Menucontrol As CommandBarControl
        Dim Popup As CommandBar
        Dim i As Long
        
        Set Popup = cbrMain.Add("右键菜单", xtpBarPopup)
        
        For i = 1 To cbrMain.ActiveMenuBar.Controls.Count
            Set Menucontrol = cbrMain.ActiveMenuBar.Controls(i)
            
'            If Menucontrol.Parent.BarID = conMenu_ManagePopup Then
            If (Menucontrol.ID = conMenu_ManagePopup Or Menucontrol.ID = conMenu_Collection) And Menucontrol.type = xtpControlPopup Then
                For Each control In Menucontrol.CommandBar.Controls
                    '处理右键 "收藏到" 菜单
                    If control.ID <> conMenu_Collection_ViewShare And control.ID <> conMenu_Collection_Manage _
                    And Mid(control.ID, 1, Decode(InStr(control.ID, "0") - 1, -1, 0, InStr(control.ID, "0") - 1)) <> comMenu_Collection_Type _
                    And Mid(control.ID, 1, Decode(InStr(control.ID, "0") - 1, -1, 0, InStr(control.ID, "0") - 1)) <> conMenu_Collection_ViewShare Then
                        '在无报告完成之前，插入模块创建的右键菜单
                        If control.ID = conMenu_Manage_Finish Then
                            If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.zlMenu.zlPopupMenu(Popup)
                            If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlMenu.zlPopupMenu(Popup)
                        End If
                        
                        control.Copy Popup
                    End If
                Next
            End If
        Next i
        
'        If Not mfrmWork_PacsImg Is Nothing Then Call mfrmWork_PacsImg.zlMenu.zlPopupMenu(Popup)
'        If Not mobjWork_Report Is Nothing Then Call mobjWork_Report.zlMenu.zlPopupMenu(Popup)
        
        Popup.ShowPopup
    End If
ErrHandle:
End Sub

Private Function GetNullAdviceInf() As TAdviceInf
    With GetNullAdviceInf
        .lngPatID = 0
        .strPatientName = ""
        .lngPatDept = 0
        .strPatientDepartment = ""
        .lngAdviceId = 0
        .lngUnit = 0
        .lngSendNo = 0
        .strStudyUID = ""
        .blnCanPrint = False
        .blnIsInsidePatient = False
        .intMoved = -1
        .intState = -1
        .intStep = -1
        .strRegNo = ""
        .lngRegId = 0
        .lngExeDepartmentId = 0
        .strExeRoom = ""
        .lngPatientFrom = 0
        .strDoDoctor = ""
        .strStudyNum = ""
        .strBedNum = ""
        .lngMarkNum = 0
        .lngBaby = -1
    End With
End Function

Private Sub FillCurAdviceTxtInfor()
'填充右上方病人基本信息
On Error GoTo ErrHandle
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    If mcurAdviceInf.lngAdviceId <= 0 Then
        labStudyNum.Caption = "[" & GetStudyNumberDisplayName & ":---]"
        lbl个人信息.Caption = "姓名:  性别:  年龄:"
        lbl检查信息.Caption = "病人科室:" & "  标识号:" & "  床  号:"
        Exit Sub
    End If
    
    With ufgStudyList
        lbl个人信息.Caption = "姓名:" & .CurText("姓名") & "  性别:" & .CurText("性别") & "  年龄:" & .CurText("年龄")
        
        If mSysPar.blnNameColColorCfg Then
            If mcurAdviceInf.strPatientType = "" Or (mstrDefaultPatientType = mcurAdviceInf.strPatientType And Not mSysPar.blnOrdinaryNameColColorCfg) Then
                lbl个人信息.ForeColor = &HC00000
                lbl检查信息.ForeColor = &HC00000
            Else
                lbl个人信息.ForeColor = zlDatabase.GetPatiColor(mcurAdviceInf.strPatientType)
                lbl检查信息.ForeColor = zlDatabase.GetPatiColor(mcurAdviceInf.strPatientType)
            End If
        End If
        
        If Not mblnIsHistory Then  '---------------------------非历次检查直接用列表中记录填充
            
            labStudyNum.Caption = "[" & GetStudyNumberDisplayName & ":" & IIf(mcurAdviceInf.strStudyNum <> "", mcurAdviceInf.strStudyNum, "---") & "]"
            lbl检查信息.Caption = "病人科室:" & mcurAdviceInf.strPatientDepartment & _
                                "  标识号:" & mcurAdviceInf.lngMarkNum & _
                                "  床号:" & mcurAdviceInf.strBedNum
                                  
            Select Case .CurText("收费")
                Case ""
                    lblCash.Caption = "欠"
                    lblCash.ForeColor = &H80FF&
                Case " "
                    lblCash.Caption = "收"
                    lblCash.ForeColor = &H8000&
                Case "  "
                    lblCash.Caption = "无"
                    lblCash.ForeColor = &HC00000
                Case "   "
                    lblCash.Caption = "补"
                    lblCash.ForeColor = &HFF&
                Case "    "
                    lblCash.Caption = "记"
                    lblCash.ForeColor = &HFF00FF
            End Select
            
            lblCash.Visible = True

        Else
            If mcurAdviceInf.lngAdviceId > 0 Then
                labStudyNum.Caption = "[" & GetStudyNumberDisplayName & ":" & IIf(mcurAdviceInf.strStudyNum <> "", mcurAdviceInf.strStudyNum, "---") & "]"
                lbl检查信息.Caption = "病人科室:" & mcurAdviceInf.strPatientDepartment & _
                                      "  标 识 号:" & mcurAdviceInf.lngMarkNum & _
                                      "  当前床号:" & mcurAdviceInf.strBedNum
                
                If mcurAdviceInf.lngBaby <> 0 Then
                    
                    strSql = "Select Nvl(A.婴儿姓名, B.姓名 || '之子' || Trim(To_Char(A.序号, '9'))) As 婴儿姓名, 婴儿性别, 出生时间" & vbNewLine & _
                            "From 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
                            "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id And A.序号 = [3]"
                            
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取婴儿信息", mcurAdviceInf.lngPatID, mcurAdviceInf.lngPageID, mcurAdviceInf.lngBaby)
                    
                    If Not rsTemp.EOF Then
                        lbl个人信息.Caption = "姓名:" & Nvl(rsTemp!婴儿姓名) & "  性别:" & Nvl(rsTemp!婴儿性别) & _
                                            "  年龄:" & Nvl(rsTemp!出生时间)
                    End If
                End If
            Else
                labStudyNum.Caption = "[" & GetStudyNumberDisplayName & ":---]"
                lbl检查信息.Caption = "病人科室:" & "  标识号:" & "  床  号:"
            End If
            
            lblCash.Caption = "历"
            lblCash.ForeColor = &HC00000
            lblCash.Visible = True
        End If
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetScanRequestCount(ByVal lngAdviceId As Long) As Long
'获取扫描申请单的数量
On Error GoTo ErrHandle
    Dim lngCount As Long
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    GetScanRequestCount = 0
    
    '如果启用申请单扫描参数 勾选，则在执行查询得到申请单图像数量，未勾选则 不执行
    If mSysPar.blnIsPetitionScan Then
        '根据医嘱ID查询 影像申请单图像表，得到已扫描张数 传入医嘱附项。并处理 VSList
        strSql = "select count(*) as 图像数 from 影像申请单图像 where 医嘱ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "得到图像数量", lngAdviceId)
        
        lngCount = Val(rsTemp!图像数)
    Else
        lngCount = 0
    End If
    
    GetScanRequestCount = lngCount
Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Function



Private Sub FillCurAdviceAppend(Optional ByVal intImgCount As Integer = 0)
'填充左下角医嘱附件
On Error GoTo ErrHandle
    Dim strAppend As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim strTemp As String
    Dim lngCount As Long
    
    With ufgStudyList
    
        If Not mblnIsHistory Then '-------------------------------------------列表选择调用
            If .GridRows <= 1 Then
                txtAppend.Text = ""
                Exit Sub
            End If
    
            txtAppend = "检查项目:" & .CurText("医嘱内容") & vbCrLf
            
            '如果启用申请单扫描参数 勾选，则在医嘱附项显示“申请单状态”，未勾选则 不显示
            If mSysPar.blnIsPetitionScan Then txtAppend = txtAppend & "申请单状态:" & IIf(intImgCount = 0, "未扫描", "已扫描（" & intImgCount & "张）") & vbCrLf
            
            txtAppend = txtAppend & "开嘱医生:" & Rpad(.CurText("开嘱医生"), 8, " ") & vbCrLf
            
            If .CurText("部位方法") <> "" Then
                For i = 0 To UBound(Split(.CurText("部位方法"), "),"))
                    If i = 0 Then
                        txtAppend = txtAppend & "检查部位:" & vbCrLf & Space(2) & "1:" & Split(.CurText("部位方法"), "),")(i) & ")"
                    Else
                        txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(.CurText("部位方法"), "),")(i) & ")"
                    End If
                Next
                If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) '取掉最后的括号
            Else
                txtAppend = txtAppend & "检查部位:" & .CurText("医嘱内容")
            End If
        Else                    '-------------------------------------------历次记录选择调用
            txtAppend = ""
            
            lngCount = GetScanRequestCount(mcurAdviceInf.lngAdviceId)
            
            gstrSQL = "Select 开嘱医生,医嘱内容 From 病人医嘱记录 Where  id =[1]"
            If mcurAdviceInf.intMoved = 1 Then gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医嘱内容", mcurAdviceInf.lngAdviceId)
            
            If rsTemp.EOF = False Then
                strTemp = Nvl(rsTemp!医嘱内容)
                If InStr(strTemp, ":") > 0 Then
                    txtAppend = "检查项目:" & Split(strTemp, ":")(0) & vbCrLf
                Else
                    txtAppend = "检查项目:" & strTemp & vbCrLf
                End If
                
                If mSysPar.blnIsPetitionScan Then txtAppend = txtAppend & "申请单状态:" & IIf(lngCount = 0, "未扫描", "已扫描（" & lngCount & "张）") & vbCrLf
                
                txtAppend = txtAppend & "开嘱医生:" & rsTemp!开嘱医生 & vbCrLf
            End If
            
            If strTemp <> "" Then
                If InStr(strTemp, ":") > 0 Then
                    strTemp = Split(strTemp, ":")(1)
                    For i = 0 To UBound(Split(strTemp, "),"))
                        If i = 0 Then
                            txtAppend = txtAppend & "检查部位:" & vbCrLf & Space(2) & "1:" & Split(strTemp, "),")(i) & ")"
                        Else
                            txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(strTemp, "),")(i) & ")"
                        End If
                    Next
                    If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) '取掉最后的括号
                Else
                    txtAppend = txtAppend & strTemp
                End If
            End If
        End If
        
        gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列" '根据历次记录是否转移判断查历史表
        If mcurAdviceInf.intMoved = 1 Then gstrSQL = Replace(gstrSQL, "病人医嘱附件", "H病人医嘱附件")
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人附件", mcurAdviceInf.lngAdviceId)
        Do Until rsTemp.EOF
            strAppend = strAppend & rsTemp!项目 & ":" & Nvl(rsTemp!内容) & vbCrLf
            rsTemp.MoveNext
        Loop
        
        gstrSQL = "select 信息名,信息值 from 病人信息从表 where 病人ID=[1] and 就诊id=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取外院病人信息", mcurAdviceInf.lngPatID, mcurAdviceInf.lngAdviceId)
        Do Until rsTemp.EOF
            strAppend = strAppend & rsTemp!信息名 & ":" & Nvl(rsTemp!信息值) & vbCrLf
            rsTemp.MoveNext
        Loop
        
        txtAppend = txtAppend & vbCrLf & vbCrLf & strAppend
    End With
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub FillHistoryStudy()
'填充历次检查记录
On Error GoTo ErrHandle
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    Dim strTemp As String
    
    If mListAdviceInf.lngAdviceId = 0 Then
        cboTimes.Clear
        Exit Sub
    End If
    
    cboTimes.tag = "" 'cbotime下拉时用到，用于区别是"增加项目"时触发还是"点击cbotimes"触发
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        'mListAdviceInf.intState = 2表示已拒绝
        strSql = "Select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 C" & _
               " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID " & _
               "" & IIf(mListAdviceInf.intState = 2, "", " And B.执行状态<>2 ") & _
               " AND A.ID=C.医嘱ID"
    Else
        strSql = "Select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,病理检查信息 C" & _
               " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID " & _
               "" & IIf(mListAdviceInf.intState = 2, "", " And B.执行状态<>2 ") & _
               " AND A.ID=C.医嘱ID"
    End If
               
    '是否选择了全部科室
    If mblnAllDepts = True Then
        strSql = strSql & " And Instr( [3],A.执行科室id ) >0 "
    Else
        strSql = strSql & " And A.执行科室id+0 =[2] "
    End If
    
    '启用关联病人，才查询关联ID
    If mSysPar.blnRelatingPatient = True And mListAdviceInf.lngLinkId <> 0 Then
        If mlngModule <> G_LNG_PATHOLSYS_NUM Then
            strSql = strSql & " union select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                " From 病人医嘱记录 A " & _
                " Where A.id in (Select 医嘱ID from 影像检查记录 Where 关联ID =[4]) "
        Else
            strSql = strSql & " union select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                " From 病人医嘱记录 A, 病理检查信息 B " & _
                " Where A.id in (Select 医嘱ID from 影像检查记录 Where 关联ID =[4]) and a.id=b.医嘱ID "
        End If
    End If
    
    strTemp = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
    strTemp = Replace(strTemp, "病人医嘱发送", "H病人医嘱发送")
    strTemp = Replace(strTemp, "影像检查记录", "H影像检查记录")
    strTemp = Replace(strTemp, "病人检查信息", "H病人检查信息")
    strSql = strSql & vbNewLine & " Union ALL " & vbNewLine & strTemp
    strSql = "Select * From (" & vbNewLine & strSql & vbNewLine & ") Order By 开嘱时间 Asc"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "", mListAdviceInf.lngPatID, _
            mlngCur科室ID, mstrCanUse科室IDs, mListAdviceInf.lngLinkId)
    
    cboTimes.Clear
    Do Until rsTemp.EOF
       cboTimes.AddItem "第" & rsTemp.AbsolutePosition & "次/共" & rsTemp.RecordCount & "次(" & Format(rsTemp!开嘱时间, "yyyy-mm-dd") & ")  " & Trim(rsTemp!医嘱内容)
       cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!医嘱ID
       
       If rsTemp!医嘱ID = mListAdviceInf.lngAdviceId Then cboTimes.ListIndex = cboTimes.NewIndex
       
       rsTemp.MoveNext
    Loop
    
    If cboTimes.ListCount > 1 Then
        cboTimes.ForeColor = &HC0&
    Else
        cboTimes.ForeColor = &H80000008
    End If
    
    cboTimes.tag = "完成"

Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub ShowTab()
'根据病人来源控制病历及医嘱选项卡
On Error GoTo ErrHandle
    Dim i As Integer
    Dim strFirstTab As String
    Dim intDefaultIndex As Integer
    Dim blnShowReport As Boolean
    
    If TabWindow.ItemCount <= 0 Then Exit Sub
    
    blnShowReport = False
     
    If Not mblnIsHistory Then '-------------------------------------------列表选择调用
        '判断 无图像不许写报告
        blnShowReport = True
        
        If mSysPar.blnReportWithImage = True Then
            If mcurAdviceInf.strStudyUID = "" Then blnShowReport = False
        End If
    End If
    
    If mcurAdviceInf.lngPatientFrom <> 2 Then '根据病人来源控制病历及医嘱选项卡
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).tag
                Case "门诊病历", "门诊医嘱"
                    TabWindow(i).Visible = True
                    
                Case "住院病历", "住院医嘱"
                    TabWindow(i).Visible = False
                    
                Case "门诊电子病历"
                    TabWindow(i).Visible = True
                
                Case "住院电子病历"
                    TabWindow(i).Visible = False
                    
                Case "影像图象"
                    TabWindow(i).Visible = True
                Case "报告填写"
                    TabWindow(i).Visible = IIf(Not mblnIsHistory, (mcurAdviceInf.intStep > 1 Or mcurAdviceInf.intStep = -1) And blnShowReport Or GetWorkModuleCount = 1, True)
                Case "排队叫号"
                    TabWindow(i).Visible = mSysPar.blnUseQueue 'True '
            End Select
        Next
    Else
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).tag
                Case "门诊病历", "门诊医嘱"
                    TabWindow(i).Visible = False

                Case "住院病历", "住院医嘱"
                    TabWindow(i).Visible = True
                
                Case "门诊电子病历"
                    TabWindow(i).Visible = False
                
                Case "住院电子病历"
                    TabWindow(i).Visible = True

                Case "影像图象"
                    TabWindow(i).Visible = True
                Case "报告填写"
                    TabWindow(i).Visible = IIf(Not mblnIsHistory, (mcurAdviceInf.intStep > 1 Or mcurAdviceInf.intStep = -1) And blnShowReport Or GetWorkModuleCount = 1, True)
                Case "排队叫号"
                    TabWindow(i).Visible = mSysPar.blnUseQueue 'True '
            End Select
        Next
    End If
    
    
    
    intDefaultIndex = GetTabWindowIndex
    
    
    '如果当前被选择的页面不可见，则显示用户的主要工作页面
    If TabWindow.Selected Is Nothing Then
        strFirstTab = mstrFirstTab
        For i = 0 To TabWindow.ItemCount - 1
            If InStr(TabWindow(i).tag, strFirstTab) > 0 And TabWindow(i).Visible Then
                TabWindow(i).Selected = True
                Exit For
            End If
        Next i
    End If
    
    If TabWindow.Selected Is Nothing Then TabWindow(intDefaultIndex).Selected = True

    If TabWindow.Selected.Visible = False Then
        strFirstTab = mstrFirstTab
        For i = 0 To TabWindow.ItemCount - 1
            If InStr(TabWindow(i).tag, strFirstTab) > 0 And TabWindow(i).Visible Then
                TabWindow(i).Selected = True
                Exit For
            End If
        Next i
    End If
    
    If TabWindow.Selected.Visible = False Then
        If intDefaultIndex < 0 Then
            TabWindow.Selected.Visible = True
        Else
            TabWindow(intDefaultIndex).Selected = True
            TabWindow(intDefaultIndex).Visible = True
        End If
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshModuleAdviceInf()
'刷新模块医嘱信息
On Error GoTo ErrHandle
    Dim intStep As Long

    If mcurAdviceInf.intState = 2 Then intStep = -2
    
    '刷新影像医技模块的医嘱信息
    If Not mfrmWork_PacsImg Is Nothing Then
        Call mfrmWork_PacsImg.zlUpdateAdviceInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
        Call mfrmWork_PacsImg.zlUpdateOtherInf(ufgStudyList, cboTimes)
    End If
    
    '刷新视频采集模块的医嘱信息
    If mblnUseActivexCapture Then
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlUpdateStudyInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
        End If
    End If

    
    '刷新病理相关模块的医嘱信息
    If Not mobjWork_Pathol Is Nothing Then
        Call mobjWork_Pathol.zlUpdateAdviceInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
    End If
    
    '刷新HIS相关模块的医嘱信息
    If Not mobjWork_His Is Nothing Then
        Call mobjWork_His.zlUpdateAdviceInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
        Call mobjWork_His.zlUpdateOtherInf(mcurAdviceInf.lngPatID, mcurAdviceInf.lngUnit, mcurAdviceInf.lngPatDept, mcurAdviceInf.lngPageID, _
            mcurAdviceInf.intState, mcurAdviceInf.strRegNo, mblnIsHistory, mcurAdviceInf.blnIsInsidePatient)
    End If
    
    '刷新报告模块的相关医嘱信息
    If Not mobjWork_Report Is Nothing Then
        '未报到前，报告编辑界面不显示
        If mcurAdviceInf.intStep < 2 And mcurAdviceInf.intStep <> -1 Then
            Call mobjWork_Report.zlUpdateAdviceInf(0, 0, 0, 0, mblnMoved)
            Call mobjWork_Report.zlRefreshFace
        Else
            Call mobjWork_Report.zlUpdateAdviceInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngPatID, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved = 1)
        End If
        
        Call mobjWork_Report.zlUpdateOtherInf(picReportContainer, ufgStudyList, mblnIsHistory, mcurAdviceInf.blnCanPrint, mcurAdviceInf.strDoDoctor, mcurAdviceInf.strStudyUID)
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshTabWindow(Optional lngAdviceIDtmp As Long = 0)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：刷新TAB页面
'参数： lngAdviceIDtmp历次记录时传入 , 其它传0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ErrHandle
    
    If TabWindow.Selected Is Nothing Then Exit Sub
    
    If TabWindow.Selected.tag = "" Then Exit Sub
    
    Select Case TabWindow.Selected.tag
        Case "影像图象"
            Call mfrmWork_PacsImg.zlRefreshFace
            
        Case "标本核收"
            Call mobjWork_Pathol.GetModule(mtSpecimen).zlRefreshFace
            
        Case "病理取材"
            Call mobjWork_Pathol.GetModule(mtMaterial).zlRefreshFace
            
        Case "病理制片"
            Call mobjWork_Pathol.GetModule(mtSlices).zlRefreshFace
            
        Case "病理特检"
            Call mobjWork_Pathol.GetModule(mtSpeExam).zlRefreshFace
            
        Case "过程报告"
            Call mobjWork_Pathol.GetModule(mtProRep).zlRefreshFace
            
        Case "报告填写"
            Call mobjWork_Report.zlRefreshFace
            If GetActiveWindow = Me.hWnd Then Call mobjWork_Report.zlShowReportVideo
            
        Case "申请费用", "住院医嘱", "门诊医嘱", "住院病历", "门诊病历", "门诊电子病历", "住院电子病历"
            Call mobjWork_His.zlRefreshFace(, mcurAdviceInf.lngPatientFrom)
            
        Case "影像采集"
            If mblnUseActivexCapture Then
                '使用ActivexExe视频采集的处理方式
                If Not mobjWork_ActiveVideo Is Nothing Then
'                    Call mobjWork_ActiveVideo.zlUpdateStudyInf(mcurAdviceInf.lngAdviceID, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved)
'                    Call mobjWork_ActiveVideo.zlRefreshVideoWindow
'                    Call mobjWork_ActiveVideo.zlRefreshData
                End If
            End If

    End Select

    If mblnUseActivexCapture Then
        '使用ActivexExe视频采集的处理方式
        '处理切换到非采集页面，然后切换检查后，采集不了图象的问题
        If Not mobjWork_ActiveVideo Is Nothing Then
            Call mobjWork_ActiveVideo.zlUpdateStudyInf(mcurAdviceInf.lngAdviceId, mcurAdviceInf.lngSendNo, mcurAdviceInf.intStep, mcurAdviceInf.intMoved)
            Call mobjWork_ActiveVideo.zlRefreshVideoWindow
            Call mobjWork_ActiveVideo.zlRefreshData
'            If mobjWork_ActiveVideo.VideoDockState Then
'                '如果处于浮动窗口状态，则需要对应刷新数据
'                mobjWork_ActiveVideo.zlRefreshData
'            End If
        End If
    End If

    
    If TabWindow.Selected.tag <> "影像采集" And TabWindow.Selected.tag <> "排队叫号" Then
        If mcurAdviceInf.lngAdviceId <= 0 Then
            Call DisableWorkModule
        Else
            Call EnableWorkModule
        End If
    Else
        EnableWorkModule
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_关联病人()
'关联病人
On Error GoTo ErrHandle
    
    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    Call frmReferencePatient.zlShowMe(mListAdviceInf.lngAdviceId, ufgStudyList.CurText("姓名"), Me, True, mlngCur科室ID)
    
    '刷新病人列表
     Call RefreshList
Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Menu_Manage_浮动采集()
On Error GoTo ErrHandle

    If Not GetIsValidOfStorageDevice(mlngCur科室ID) Then
      MsgBoxD Me, "影像存储设备未定义或处于停用，请检查！", vbInformation, gstrSysName
      Exit Sub
    End If
    
    If Not mobjWork_ActiveVideo Is Nothing Then
        Call mobjWork_ActiveVideo.zlShowPopupVideo
        
        If mlngOldAdviceId <> mcurAdviceInf.lngAdviceId And TabWindow.Selected.Caption <> "影像采集" Then
            Call mobjWork_ActiveVideo.zlRefreshData
            mlngOldAdviceId = mcurAdviceInf.lngAdviceId
        End If
    End If
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub Menu_Manage_图像刻录()
'图像刻录
    Dim lngCurAdviceId As Long
    Dim objBurn As Object
    Dim frmBurn As frmImageBurn
    
    If mListAdviceInf.intImageLocation = 1 Then
        Call subXWShowArchiveManager(3)
    Else
        On Error GoTo errExit
            Set objBurn = CreateObject("IMAPI2.MsftDiscMaster2")
            Set objBurn = Nothing
            GoTo continueBurn
errExit:
            Call MsgBoxD(Me, "不能创建刻录对象，请在安装IMAPI2刻录组件后重新进入。", vbOKOnly, Me.Caption)
            Exit Sub
            
continueBurn:
            
            Set frmBurn = New frmImageBurn
        On Error GoTo errFree
            
            lngCurAdviceId = Val(ufgStudyList.CurKeyValue)
            
            Set frmBurn = New frmImageBurn
            Call frmBurn.ShowBurn(mlngModule, mlngCur科室ID, lngCurAdviceId, mblnMoved, Me)
errFree:
            Call Unload(frmBurn)
            Set frmBurn = Nothing
    End If
End Sub

Private Sub Menu_Manage_收藏管理()
'收藏管理
On Error GoTo errFree
    Dim frmCollectionManage As New frmCollectionManage
    Dim lngCount As Long

    Call frmCollectionManage.ShowCollectionManageWind(Me)
    
    '删除现在的工具栏及顶级菜单项
    Call LockWindowUpdate(Me.hWnd)
    For lngCount = cbrMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbrMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbrMain.Count To 2 Step -1
        cbrMain(lngCount).Delete
    Next
    
    Call InitCommandBars
    Call BindMenuShortcut(App.ProductName, mlngModule, cbrMain)
    
    Call CreateWorkModuleMenu
    
    Call LockWindowUpdate(0)
    
errFree:
    Call Unload(frmCollectionManage)
    Set frmCollectionManage = Nothing
End Sub

Private Sub Menu_Manage_收藏到()
'收藏到
    Dim frmToCollection As New frmToCollection
    Dim rsTemp As ADODB.Recordset
    Dim lngAdviceId As Long
    Dim lngSendNo As Long
On Error GoTo errFree

    lngAdviceId = Val(ufgStudyList.CurText("医嘱ID"))
    lngSendNo = Val(ufgStudyList.CurText("发送号"))
    
    If lngAdviceId = 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If
    
    gstrSQL = "select 首次时间 from 病人医嘱发送 where 医嘱ID= " & lngAdviceId & ""
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    '判断选中记录是否报到，如果没有报到则不能进行收藏操作
    Do While Not rsTemp.EOF
        If Nvl(rsTemp!首次时间) = "" Then
            Call MsgBoxD(Me, "该检查未报到，不能收藏！", vbOKOnly, "影像病理工作站")
            Exit Sub
        End If
        
        rsTemp.MoveNext
    Loop
    
    Call frmToCollection.ShowToCollectionWind(Me, lngAdviceId, lngSendNo)
    
errFree:
    Call Unload(frmToCollection)
    Set frmToCollection = Nothing
End Sub


Private Sub Menu_Manage_收藏数据显示(ByVal control As XtremeCommandBars.ICommandBarControl, ByVal bytStyle As Byte)
'收藏数据显示方法
On Error GoTo errHand
    Dim rsList As ADODB.Recordset
    Dim strCollectionType As String
    Dim lngFatherID As Long
    Dim strUser As String
    
    '处理收藏类别字符串
    If InStr(control.Caption, "(") = 0 Then
        strCollectionType = control.Caption
    Else
        strCollectionType = Mid(control.Caption, 1, InStr(control.Caption, "(") - 1)
    End If
    
    '处理创建人数据
    strUser = control.DescriptionText ' Category
    
    '处理父级ID字符串
    If bytStyle = 0 Then
        lngFatherID = CLng(control.ID) - CLng(comMenu_Collection_Type) * 10000#
    ElseIf bytStyle = 1 Then
        lngFatherID = CLng(control.ID) - CLng(conMenu_Collection_ViewShare) * 10000#
    End If
    
    '将参数传入 数据加载方法
    Set rsList = GetCollectionData(strCollectionType, lngFatherID, strUser)
   

    
    Set ufgStudyList.AdoData = rsList
    
    ufgStudyList.AdoFilter = ""
    
    Call ufgStudyList.BindData
   
    If ufgStudyList.AdoData.RecordCount > 0 Then Call ufgStudyList_OnSelChange

    Call RefreshStatusBarInf
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_CheckList()
    If mListAdviceInf.lngAdviceId > 0 Then
        Set mclsCISKernel = New clsCISKernel
        Call mclsCISKernel.ShowPacsApplication(Me, mListAdviceInf.lngAdviceId)
        Set mclsCISKernel = Nothing
    Else
        MsgBox "没有选中病人。", vbInformation + vbOKOnly, gstrSysName
    End If
End Sub
Private Function GetCollectionData(ByVal strCollectionType As String, ByVal lngFatherID As Long, ByVal strUser As String) As ADODB.Recordset
'加载共享数据
    Dim strSql As String
    Dim strWithCollectionTab As String
    Dim strWithOrderTab As String   '医嘱子查询
    
    Set GetCollectionData = Nothing
    
    '根据参数判断连接那一段SQL语句
    If Len(Trim(strCollectionType)) <> 0 And strCollectionType <> "查看当前收藏" Then
        strWithCollectionTab = " with tmpCollection as (select m.医嘱id as id from 影像收藏类别 L,影像收藏内容 M where " & vbNewLine & _
                                        " L.id=M.收藏id" & vbNewLine & _
                                        " and l.创建人='" & Decode(strUser, "", UserInfo.姓名, strUser) & "' and l.收藏类别='" & strCollectionType & "' )"
    ElseIf lngFatherID <> 0 Then
        strWithCollectionTab = " with tmpCollection as (select m.医嘱id as id from 影像收藏类别 L,影像收藏内容 M where " & vbNewLine & _
                                        "L.id=M.收藏id" & vbNewLine & _
                                        " and L.id in (select distinct id from 影像收藏类别 start with id =" & lngFatherID & " connect by prior id=上级id) )"
    End If
    
    strWithOrderTab = "tmpOrder as (select id from tmpCollection Union All select a.ID from 病人医嘱记录 a, tmpCollection b where a.相关ID=b.ID and a.医嘱状态 <> 4 )"
    
    If mlngModule <> G_LNG_PATHOLSYS_NUM Then
        strSql = "Select * From (" & vbNewLine & _
             strWithCollectionTab & "," & vbNewLine & strWithOrderTab & vbNewLine & _
             "Select  Distinct" & vbNewLine & _
                    "       A.医嘱ID,B.相关ID,A.发送号,A.首次时间 报到时间,A.发送时间 申请时间,A.执行状态,nvl(A.执行过程,0) 检查过程,A.执行间,A.结果阳性 阳性,h.危急状态 危急," & vbNewLine & _
                    "       B.病人ID,B.主页ID,B.挂号单,B.病人科室ID,Decode(B.病人来源, 1, '门', 2, '住', 3, '外', 4, '体') 来源,B.医嘱内容,B.标本部位," & vbNewLine & _
                    "       Nvl(B.紧急标志, 0) 紧急标志, Nvl(B.婴儿, 0) 婴儿,B.开嘱医生,A.NO,C.当前床号,C.当前病区ID,Decode(B.病人来源,2,C.住院号,C.门诊号) 标识号,b.开嘱时间,c.门诊号,c.住院号," & vbNewLine & _
                    "       Nvl(B.姓名,H.姓名) 姓名,H.影像类别,H.检查号,Nvl(B.性别,H.性别) 性别,Nvl(B.年龄,H.年龄) 年龄,H.身高,H.体重,H.影像质量,H.符合情况," & vbNewLine & _
                    "       Decode(B.病人来源,3,B.开嘱医生,A.发送人) 登记人,H.报到人,H.报告发放,H.发放胶片,H.关联ID,A.记录性质, " & vbNewLine & _
                    "       H.完成人,H.是否电子胶片,H.是否打印,H.报告操作,H.绿色通道,H.报告打印,H.报告人,H.报告质量,H.复核人,H.是否技师确认,H.检查技师,H.检查技师二,H.接收日期 采图时间," & vbNewLine & _
                    "       H.随访描述,H.诊断分类,H.检查UID,H.图像位置,A.执行部门ID as 执行科室ID,0 as 转出,F.名称 AS 病人科室, a.采样时间, " & vbNewLine & _
                    "       C.就诊卡号,A.NO as 单据号,C.身份证号,C.结算模式,decode(B.病人来源,2,D.病人类型,C.病人类型) as 病人类型,D.路径状态,A.计费状态,Decode(A.记录性质,2,1,Decode(a.计费状态,3,1,0)) as 收费 ,z.医嘱ID as 申请单医嘱" & vbNewLine & _
                    " From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C,病案主页 D,影像检查记录 H,部门表 F, tmpOrder ,影像申请单图像 Z" & vbNewLine & _
                    " Where A.医嘱ID=B.ID And A.医嘱ID=H.医嘱ID(+) And A.发送号=H.发送号(+) " & vbNewLine & _
                    " And B.病人ID=C.病人ID And B.病人科室id=F.ID " & vbNewLine & _
                    " And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) and a.医嘱ID = z.医嘱ID(+) and a.医嘱ID=tmpOrder.id) "
    Else
        strSql = "Select * From (" & vbNewLine & _
             strWithCollectionTab & "," & vbNewLine & strWithOrderTab & vbNewLine & _
             "Select Distinct" & vbNewLine & _
             "       A.医嘱ID,B.相关ID,A.发送号,A.首次时间 报到时间,A.发送时间 申请时间,A.执行状态,nvl(A.执行过程,0) 检查过程,A.结果阳性 阳性,h.危急状态 危急," & vbNewLine & _
             "       '' as 病理执行状态, o.取材过程,o.制片过程,o.免疫过程,o.分子过程,o.特染过程,b.开嘱时间,c.门诊号,c.住院号, " & vbNewLine & _
             "       decode(o.检查类型,0,'常规',1,'冰冻',2,'细胞',3,'会诊',4,'尸检',5,'快速石蜡',null) as  检查类别, " & vbNewLine & _
             "       decode(o.病理号,null,'未核收','已核收') as 核收情况, " & vbNewLine & _
             "       B.病人ID,B.主页ID,B.挂号单,B.病人科室ID,Decode(B.病人来源, 1, '门', 2, '住', 3, '外', 4, '体') 来源,B.医嘱内容,B.标本部位," & vbNewLine & _
             "       Nvl(B.紧急标志, 0) 紧急标志, Nvl(B.婴儿, 0) 婴儿,B.开嘱医生,A.NO,C.当前床号,C.当前病区ID,Decode(B.病人来源,2,C.住院号,C.门诊号) 标识号," & vbNewLine & _
             "       Nvl(B.姓名,H.姓名) 姓名,Nvl(B.性别,H.性别) 性别,Nvl(B.年龄,H.年龄) 年龄,H.身高,H.体重,o.综合质量," & vbNewLine & _
             "       Decode(B.病人来源,3,B.开嘱医生,A.发送人) 登记人,H.报到人,o.病理号,H.报告发放,H.发放胶片,H.关联ID,A.记录性质, " & vbNewLine & _
             "       H.完成人,H.是否打印,H.报告操作,H.绿色通道,H.报告打印,H.报告人,H.报告质量,H.复核人,H.是否技师确认,H.检查技师,H.检查技师二,H.接收日期 采图时间, " & vbNewLine & _
             "       H.随访描述,H.诊断分类,H.检查UID,H.图像位置,0 as 转出,F.名称 AS 病人科室, a.采样时间, Y.当前状态 as 会诊状态, Y.会诊医师, Y.Id as 会诊ID, " & vbNewLine & _
             "       C.就诊卡号,A.NO as 单据号,C.身份证号,C.结算模式,decode(B.病人来源,2,D.病人类型,C.病人类型) as 病人类型,D.路径状态,A.计费状态,Decode(A.记录性质,2,1,Decode(a.计费状态,3,1,0)) as 收费,z.医嘱ID as 申请单医嘱, " & vbNewLine & _
             "      (select count(1) from 病理检查信息 V , 病理申请信息 W where V.病理医嘱ID=w.病理医嘱id and v.医嘱id=A.医嘱ID and w.补费状态=1) as 补费 " & vbNewLine & _
             " From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C,病案主页 D,影像检查记录 H,部门表 F, " & vbNewLine & _
             "       病理检查信息 o, tmpOrder ,影像申请单图像 Z, 病理会诊信息 Y" & vbNewLine & _
             " Where A.医嘱ID=B.ID And A.医嘱ID=H.医嘱ID(+) And A.发送号=H.发送号(+) " & vbNewLine & _
             "       And B.病人ID=C.病人ID And B.病人科室id=F.ID " & vbNewLine & _
             "       and A.医嘱ID=o.医嘱ID(+) and o.病理医嘱ID=Y.病理医嘱ID(+) " & vbNewLine & _
             "       And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) and a.医嘱ID = z.医嘱ID(+) and a.医嘱ID=tmpOrder.id) "
    End If
      
    strSql = strSql & vbNewLine & "Order by 检查过程,报到时间,申请时间"
    
    Set GetCollectionData = GetDataToLocal(strSql, GetWindowCaption)
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Private Sub Menu_Petition_扫描申请单(ByVal intType As Integer)
'intType:0--查看申请单；1--扫描申请单
Dim frmPetitionCap As New frmPetitionCapture

On Error GoTo errFree

    If mListAdviceInf.lngAdviceId <= 0 Then
        MsgBoxD Me, M_STR_HINT_NoSelectData, vbInformation, Me.Caption
        Exit Sub
    End If

    With ufgStudyList
        Call frmPetitionCap.ShowPetitionCaptureWind(mstrPrivs, _
                                                mlngCur科室ID, _
                                                Nvl(.CurText("病人科室")), _
                                                Nvl(.CurText("姓名")), _
                                                Nvl(.CurText("年龄")), _
                                                Nvl(.CurText("性别")), _
                                                Nvl(.CurText("医嘱内容")), _
                                                Nvl(.CurText("部位方法")), _
                                                IIf(InStr(mstrPrivs, "检查登记") <= 0, True, IIf(intType = 0, True, False)), _
                                                False, _
                                                mListAdviceInf.lngAdviceId, _
                                                IIf(.CurText("检查过程") = "已拒绝", 1, IIf(.CurText("检查过程") = "已完成", 2, 0)))
    End With
errFree:
    Call Unload(frmPetitionCap)
    Set frmPetitionCap = Nothing
End Sub

Private Sub ufgStudyList_OnSelChange()
On Error GoTo ErrHandle
    Dim i As Integer
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    '如果是打印清单的操作 则停止行改变事件，不然会造成界面大量刷新
    If mblnIsPrintMode Then Exit Sub
    
    mblnIsHistory = False
    
    If mblnvsRefresh Then Exit Sub
    
    gblnIsStudyChage = True
    
    mcurAdviceInf = GetAdviceDetailInf()
    mListAdviceInf = mcurAdviceInf
    
    Call FillCurAdviceTxtInfor '填充右上方病人基本信息
    Call FillHistoryStudy '填充历次检查记录
    Call SetSelectRowColor
    
    If mListAdviceInf.lngAdviceId <= 0 Then '无记录时处理
        cboTimes.Clear
        txtAppend = ""

        lblCash.Visible = False
        
        If Not TabWindow.Selected Is Nothing Then
            Call ConfigSubForm(TabWindow.Selected)
        End If
    
        Call RefreshModuleAdviceInf
        Call RefreshTabWindow
    Else
        mintImgCount = GetScanRequestCount(mListAdviceInf.lngAdviceId)

        Call RefreshModuleAdviceInf
        
        Call FillCurAdviceAppend(mintImgCount) '填充左下角医嘱附件
        Call ShowTab '根据病人提供不同选项卡
        
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))  '显示可打印的诊疗单据:之所以即时加载,是为了使用F2热键
        
        If Not TabWindow.Selected Is Nothing Then
            Call ConfigSubForm(TabWindow.Selected)
        End If

        
        '判读是否手动刷新的检查列表（如果是手动刷新，则需要通知其他工作模块进行刷新）...
        If mblnIsCallModuleRefresh Then
            mblnIsCallModuleRefresh = False
            
            Call NotificationAllModuleRefresh
        End If
        
        If mstrFirstTab <> "" Then '不为空表示按定制首页显示,由TabWindow调用刷新
            
            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow.Item(i).tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                    Exit For
                End If
            Next
            
            If i = TabWindow.ItemCount Then    '没循环到了触发第1个可视TAB
                For i = 0 To TabWindow.ItemCount - 1
                    If TabWindow.Item(i).Visible Then
                        Exit For
                    End If
                Next i
            End If
            
            '刷新页面，并显示定制首页
            If TabWindow.Item(i).Selected Then
                Call RefreshTabWindow
            Else
                TabWindow.Item(i).Selected = True
            End If
        Else
            Call RefreshTabWindow
        End If
        
    End If
    
    '恢复焦点，因在数据刷新过程中，可能造成列表焦点的失去，失去焦点后，将不能使用鼠标滚轮滚动列表
    If ufgStudyList.Visible And Not mblnAutoRefreshList Then 'GetActiveWindow = Me.hWnd
        Me.dkpMain.Panes(1).Selected = True
        Call ufgStudyList.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetSelectRowColor()
    Dim lngRowSel As Long
    
    lngRowSel = ufgStudyList.DataGrid.RowSel
    
    If lngRowSel < 0 Then Exit Sub
    
    Call SetStateColor(lngRowSel)
    
    If ufgStudyList.DataGrid.Cols > 1 And ufgStudyList.DataGrid.Rows > 1 Then
        ufgStudyList.DataGrid.Cell(flexcpFontBold, 1, 1, ufgStudyList.DataGrid.Rows - 1, ufgStudyList.DataGrid.Cols - 1) = False

        ufgStudyList.DataGrid.Cell(flexcpFontBold, lngRowSel, 1, lngRowSel, ufgStudyList.DataGrid.Cols - 1) = True
        
        ufgStudyList.DataGrid.Cell(flexcpFontSize, 1, 1, ufgStudyList.DataGrid.Rows - 1, ufgStudyList.DataGrid.Cols - 1) = mbytFontSize
    End If
End Sub

Private Sub SetStateColor(ByVal lngRowSel As Long)
    Dim lngForeColor As Long
    Dim lngR As Long, lngG As Long, lngB As Long
    
    If mSysPar.lngListColorMark <> 0 Then Exit Sub
    
    lngForeColor = getRowColor(lngRowSel)
    
    Call GetRGB(lngForeColor, lngR, lngG, lngB)
    
    ufgStudyList.DataGrid.ForeColorSel = RGB(lngR - 30, lngG - 30, lngB - 30)
    ufgStudyList.DataGrid.BackColorSel = &HFEE0E2      '&HFECFD2
End Sub

Private Function getRowColor(ByVal lngRowSel As Long) As Long
    Dim lngRowColor As Long
    
    If ufgStudyList.Text(lngRowSel, "检查过程") = "已拒绝" Then lngRowColor = gdblColor已拒绝
    If ufgStudyList.Text(lngRowSel, "检查过程") = "已完成" Then lngRowColor = gdblColor已完成
    If ufgStudyList.Text(lngRowSel, "检查过程") = "已报到" Then lngRowColor = gdblColor已报到
    If ufgStudyList.Text(lngRowSel, "检查过程") = "已登记" Then lngRowColor = gdblColor已登记
    If ufgStudyList.Text(lngRowSel, "检查过程") = "已检查" Then lngRowColor = gdblColor已检查
    If ufgStudyList.Text(lngRowSel, "检查过程") = "已审核" Then lngRowColor = gdblColor已审核
    If ufgStudyList.Text(lngRowSel, "检查过程") = "处理中" Then lngRowColor = gdblColor处理中
    If ufgStudyList.Text(lngRowSel, "检查过程") = "报告中" Then lngRowColor = gdblColor报告中
    If ufgStudyList.Text(lngRowSel, "检查过程") = "审核中" Then lngRowColor = gdblColor审核中
    If ufgStudyList.Text(lngRowSel, "检查过程") = "已报告" Then lngRowColor = gdblColor已报告
    If ufgStudyList.Text(lngRowSel, "检查过程") = "已驳回" Then lngRowColor = gdblColor已驳回
    
    getRowColor = lngRowColor
End Function

Private Sub Menu_Manage_SetXWParam_click()
'------------------------------------------------
'功能：打开新网PACS的参数设置窗口
'返回：
'------------------------------------------------
    On Error GoTo err
    
    Call frmXWSetParams.zlShowMe(Me)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub conMenu_File_SendImg_click()
'------------------------------------------------
'功能：发送图像
'返回：
'------------------------------------------------
    On Error GoTo err
    If mlngModule = G_LNG_PACSSTATION_MODULE Then
        If mListAdviceInf.lngAdviceId <= 0 Or mListAdviceInf.intImageLocation = 1 Then
            Call subXWShowArchiveManager(2)
        Else
            frmPacsSendImage.ShowMe Me
        End If
    Else
        frmPacsSendImage.ShowMe Me
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


