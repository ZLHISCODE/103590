VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmLabMain 
   Caption         =   "检验技师工作站"
   ClientHeight    =   6750
   ClientLeft      =   1515
   ClientTop       =   675
   ClientWidth     =   10995
   Icon            =   "frmLabMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmLabMain.frx":058A
   ScaleHeight     =   6750
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicWindows 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   7500
      ScaleHeight     =   585
      ScaleWidth      =   795
      TabIndex        =   29
      Top             =   2370
      Width           =   795
   End
   Begin VB.PictureBox picBarCodePrint 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4890
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   26
      Top             =   390
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSWinsockLib.Winsock WinsockC 
      Left            =   750
      Top             =   690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ComboBox cboExesItem 
      Height          =   300
      Left            =   3930
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1590
      Width           =   1875
   End
   Begin VB.ComboBox cboUnionItem 
      Height          =   300
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1590
      Width           =   1875
   End
   Begin VB.TextBox TxtGoto 
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   1530
      Width           =   1755
   End
   Begin VB.ComboBox cboMachine 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1170
      Width           =   1905
   End
   Begin VB.ComboBox cboDept 
      Height          =   300
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   1875
   End
   Begin MSComctlLib.ImageList Imglist 
      Left            =   120
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":6DDC
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":7376
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":7910
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":7EAA
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":8444
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":89DE
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":8D78
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":9112
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":94AC
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":9846
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":100A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":1690A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":1D16C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":239CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":2A230
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":30A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabMain.frx":3102C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   120
      ScaleHeight     =   3945
      ScaleWidth      =   5865
      TabIndex        =   3
      Top             =   1950
      Width           =   5865
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   765
         Left            =   300
         TabIndex        =   4
         Top             =   2460
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptList1 
         Height          =   765
         Left            =   2160
         TabIndex        =   12
         Top             =   2430
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1349
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin MSComCtl2.DTPicker dtpDateEnd 
         Height          =   300
         Left            =   4350
         TabIndex        =   28
         Top             =   3600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   247267329
         CurrentDate     =   40049
      End
      Begin MSComCtl2.DTPicker DTPDate 
         Height          =   300
         Left            =   2970
         TabIndex        =   27
         Top             =   3600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Format          =   247267329
         CurrentDate     =   40049
      End
      Begin VB.ComboBox cbo时间 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3600
         Width           =   1275
      End
      Begin VB.PictureBox PicFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4500
         Picture         =   "frmLabMain.frx":3788E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   24
         Top             =   22
         Width           =   240
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "门诊"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   23
         ToolTipText     =   "门诊和直接登记标本"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "住院"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   825
         TabIndex        =   22
         ToolTipText     =   "住院标本"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "无主"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   1560
         TabIndex        =   21
         ToolTipText     =   "没有病人信息的标本"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "已审"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   2310
         TabIndex        =   20
         ToolTipText     =   "已审核标本"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "未审"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   3045
         TabIndex        =   19
         ToolTipText     =   "未审核标本"
         Top             =   30
         Width           =   735
      End
      Begin VB.CheckBox chkSoure 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         Caption         =   "体检"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   3780
         TabIndex        =   18
         ToolTipText     =   "未审核标本"
         Top             =   30
         Width           =   735
      End
      Begin XtremeSuiteControls.TabControl TabList 
         Height          =   1575
         Left            =   90
         TabIndex        =   11
         Top             =   300
         Width           =   3525
         _Version        =   589884
         _ExtentX        =   6218
         _ExtentY        =   2778
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox PicInfo 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3990
      ScaleHeight     =   315
      ScaleWidth      =   4245
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   4245
      Begin XtremeCommandBars.CommandBars cbrChild 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox PicTab 
      BorderStyle     =   0  'None
      Height          =   1785
      Left            =   6840
      ScaleHeight     =   1785
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   330
      Width           =   3675
      Begin XtremeSuiteControls.TabControl TabCtlWindow 
         Bindings        =   "frmLabMain.frx":3E0E0
         Height          =   1575
         Left            =   90
         TabIndex        =   1
         Top             =   120
         Width           =   3525
         _Version        =   589884
         _ExtentX        =   6218
         _ExtentY        =   2778
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   6390
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabMain.frx":3E0F4
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14314
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
   Begin VSFlex8Ctl.VSFlexGrid vgdList 
      Height          =   900
      Left            =   3000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin RichTextLib.RichTextBox RtfTxt 
      Height          =   885
      Left            =   2190
      TabIndex        =   13
      Top             =   90
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1561
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmLabMain.frx":3E988
   End
   Begin VB.PictureBox PicImage 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   6540
      ScaleHeight     =   2595
      ScaleWidth      =   1935
      TabIndex        =   15
      Top             =   3450
      Width           =   1935
      Begin VB.VScrollBar VScroll 
         Height          =   1245
         Left            =   1620
         Max             =   0
         TabIndex        =   17
         Top             =   150
         Width           =   225
      End
      Begin C1Chart2D8.Chart2D ChartThis 
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   120
         Width           =   885
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   1561
         _ExtentY        =   1296
         _StockProps     =   0
         ControlProperties=   "frmLabMain.frx":3EA17
      End
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   1440
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmLabMain.frx":3EF9A
      Left            =   810
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmLabMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const Dkp_ID_List As Integer = 1                            '样本清单窗格
Private Const Dkp_ID_Locate As Integer = 2                          '定位查找窗格
Private Const Dkp_ID_Request As Integer = 3                         '核对登记窗格
Private Const Dkp_ID_Append As Integer = 4                          '报告、跟踪、费用等附加窗格
Private Const Dkp_ID_Image As Integer = 5                           '显示检验图像
Private Const p医嘱附费管理 As Integer = 1257                       '病人费用模块授权
Private Const p门诊医嘱下达 As Integer = 1252                       '门诊医嘱下达
Private Const p住院医嘱下达 As Integer = 1253                       '住院医嘱下达
Private Const p门诊病历管理 As Integer = 1250                       '门诊病历
Private Const p住院病历管理 As Integer = 1251
Private Const p新版病历管理 As Integer = 2250                       '新版病历

Private Const ID_MENU_MOUSE = 90                                    '右键菜单
Private Const con_主界面筛选_检验中 As String = "门诊病人;住院病人;无主标本;已审标本;未审标本;体检病人;紧急医嘱;紧急标本;质控标本;审核已通过;审核未通过;未做完;已做完;仪器审核通过;仪器审核未通过"
Private Const con_主界面筛选_待核收 As String = "门诊病人;住院病人;体检病人"
'-----------------定义需调入的窗体--------------------
Private WithEvents mfrmRequest As frmLabRequest                     '核收登记窗体
Attribute mfrmRequest.VB_VarHelpID = -1
Private WithEvents mfrmWrite As frmLisStationWrite                  '报告填写窗体
Attribute mfrmWrite.VB_VarHelpID = -1
Private WithEvents mfrmWrite2 As frmLisStationWrite2                '填写微生物
Attribute mfrmWrite2.VB_VarHelpID = -1
Private WithEvents mfrmLabMainSampleUnion  As frmLabMainSampleUnion '标本合并
Attribute mfrmLabMainSampleUnion.VB_VarHelpID = -1
Private WithEvents mclsInAdvices As zlCISKernel.clsDockInAdvices    '住院医嘱
Attribute mclsInAdvices.VB_VarHelpID = -1
Private WithEvents mclsOutAdvices As zlCISKernel.clsDockOutAdvices  '门诊医嘱
Attribute mclsOutAdvices.VB_VarHelpID = -1
Private mclsInEPRs As zlRichEPR.cDockInEPRs               '住院病历
Attribute mclsInEPRs.VB_VarHelpID = -1
Private mclsOutEPRs As zlRichEPR.cDockOutEPRs             '住院病历
Attribute mclsOutEPRs.VB_VarHelpID = -1
Private mfrmTrack As frmLabTrack                                    '历次对比
Private WithEvents mfrmLabMicrobe3Report As frmLabMicrobe3Report    '三级报告
Attribute mfrmLabMicrobe3Report.VB_VarHelpID = -1

'Private mfrmLabMainImage As frmLabMainImage                        '检验图像显示
Private WithEvents mclsExpenses As zlPublicExpense.clsDockExpense        '新的费用\
Attribute mclsExpenses.VB_VarHelpID = -1
Private mclspublicExpenses As zlPublicExpense.clsPublicExpense        '新的费用部件

Private mcolSubForm As Collection                                   '卸载子窗体
Private mblnCompelRefresh As Boolean                                '强制刷新
Private mintUnion As Integer                                        '是否区分仪器进行显示 0=不区分 1=区分
Private mSendReport As Integer                                      '审核后是否自动发送报告 0=发送 1=不发送

'-----------------------------------------------------
'-----------------------------------------------------
Private mlngDeptID As Long                                          '科室ID
Private mlngKey As Long                                             '检验标本ID
Private mintEditState As Integer                                    '当前编辑状态：0-非编辑；1-新增核收；2-新增登记；4-补填病人；3-重新核收；5-报告编辑;6-标本合并;7-三级报告
Private mintHandleState As Integer                                  '当前操作状态:1 = 病人信息 = 2报告单操作 3= 三级报告
Private mintContinue As Integer                                     '目前是否处于连续核收登记状态


Private mstrPrivs As String                                         '权限
'Private objLISComm As Object                                       '通讯接口



'-----------------------------------------------------
'----------------------参数设置变量-------------------
Dim blnChecking As Boolean                                          '是否正在进行操作
Dim blnAutoRefresh As Boolean                                       '是否在收到仪器数据时自动刷新
Dim blnComm As Boolean                                              '是否允许双向通信
Dim blnAutoPrint As Boolean                                         '审核后自动打印
'-----------------------------------------------------
'---------------------审核时判断----------------------
Private mintAuditing As Integer                                     '是否有审核权限 0=没有权限 1=有权限
                                                                    '-1至-24=有效时间计算时取绝对值
Private mDataAuditing As Date                                       '有时间限定后,记录时间
Private mstrAuditingMan As String                                   '审核人,审核时传入过程
Private mstrAuditingManID As String                                 '审核人登陆名(签名用)
Private mblnCancel As Boolean                                       '取消刷新
Private mUserDept As String                                         '用户所属科室字串
Private mblnVerifying(15) As Boolean                                '检验中筛选状态
Private mblnWaitVerify(2) As Boolean                                '等待检验

Private mMakeNoRule As String                                       '标本序号生成的日期规则
Private mstrMachines As String                                      '记录有操作权限的仪器
Private mstrMachineALL As String                                    '记录可以显示的仪器ID字串


Private mbln手工发送杯号 As Boolean                                 '手工发送
Private mbln保存后直接审核 As Boolean                               '记录是否保存后自动审核
Private mstrPrintDepts As String                                    '可以打印的科室
Private mblnAout As Boolean                                         '是否自动找下一个可审核的标本
Private mlngLastShow As Long                                        '最后显示的标本的ID
Private mTodayQCPrivs As String                                     '今日质控权限
Private mHistoryPrivs As String                                     '历史质控权限
Private mTableRefresh As Boolean                                    'Table是否刷新
Private mintLoadShow As Integer                                     '显示过值大于1

Const mcontIntRowHeight As Integer = 230                            '记录行高

Public mblnSendComplete As Boolean                                  '是否完成传送
Public mstrMachineGroup As String                                   '仪器分组
Public mlngMachineID As Long                                        '选择仪器ID -1=手工 0=所有仪器 >0=仪器ID
Private mstrMachineID  As String                                    '保存微生物仪器ID串

Dim mclsEMR As Object                                               '新版电子病历


Dim mblnTabList1 As Boolean                                         '控件分页是否使用过 0=未使用 1=已使用


'==检验中列表
Private Enum mCol
    ID = 0
    紧急
    紧急医嘱
    执行状态
    所属情况
    标本类型
    标本号
    姓名
    性别
    年龄
    检验项目
    标识号
    传送
    结果次数
    医嘱id
    仪器id
    转出
    病人ID
    标本时间
    报告时间
    微生物标本
    收费单
    挂号单
    检验人
    审核人
    样本条码
    婴儿
    病人科室
    发送号
    仪器名
    主页ID
    开嘱科室ID
    报告结果
    年龄数字
    年龄单位
    床号
    申请人
    标本形态
    采样人
    采样时间
    检验标本
    NO
    接收人
    接收时间
    审核时间
    病区id
    病区名称
    定位
    执行科室ID
    标本类别
    医嘱紧急
    标本紧急
    申请科室
    申请类型
    复查
    查阅状态
    报告发送
    病人科室ID
    初审人
    初审时间
    单位
    健康号
    审核未通过
    病人来源
    门诊号
    住院号
    结果为空
    临床路径病人
    仪器审核
End Enum

'==待检验列表
Private Enum mRCol
    病人ID
    紧急
    来源
    姓名
    性别
    年龄
    病人科室
    标识号
    床号
    医嘱内容
    开嘱医生
    开嘱时间
    诊疗项目ID
    医嘱id
    执行状态
    定位
    挂号单
    签收时间
End Enum

'==标本操作
Private Enum mActS
    核收 = 0: 登记: 重新核收: 补填病人
    修改样本号
    批量修改样本号
    删除无主标本
    发送仪器
    批量发送到仪器
    置为质控
    置为对比
    状态回滚
    批量增加
    拒收
    置为无主
    合并标本
    合并标本保存
    修改病人信息
End Enum
'==报告操作
Private Enum mActR
    批量调整报告 = 0
    审核报告
    发送报告
    批量审核报告
    审核取消
    重做结果
    取消重做
    填写报告
    写入病历
    验证签名
    按病人审核
    填写三级报告
End Enum
'保存、拒收、放弃
Private Enum mFileS
    保存
    放弃
End Enum
Private Enum mFilter                        '过滤条件
    姓名 = 0
    性别
    年龄
    年龄单位
    标识号
    标本号
    单据号
    检验类别
    检验人
    检验项目
    检验时间
    送检科室
    送检人
    检验仪器
    细菌
    抗生素
    药敏结果
    是否使用高级
    高级
    病人ID
End Enum
Private Enum mSWork                         '用于键盘快捷操作
    Key_PageUP
    Key_PageDown
    Key_Home
    Key_End
End Enum
Private Const conMenu_IDkind_Change  As Integer = 12345
Private int体检处理方式 As Integer  '1-提示，2-修正，3-不修正
Private int门诊处理方式 As Integer  '1-提示，2-修正，3-不修正
Private int住院处理方式 As Integer  '1-提示，2-修正，3-不修正
Private int院外处理方式 As Integer  '1-提示，2-修正，3-不修正

'--------------------------------------------
'插件相关定义
Implements zl9LisQuery_Def.clsLisQueryHost
Private clsPluginLoader  As PlugInLoader
Private mobjPlugin()   As zl9LisQuery_Def.clsLisQuery
'--------------------------------------------
Private Sub CreateCbs()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrthis.Icons = zlCommFun.GetPubIcons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False
    

    '-----------------------------------------------------
    '菜单定义
    Me.cbrthis.ActiveMenuBar.Title = "菜单"
    Me.cbrthis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&T)…"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "批量打印(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintBedCard, "重打条码(&A)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "放弃(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_RowPrint, "清单打印(&L)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Privacy, "审核人登陆(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&O)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "酶标仪(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "样本(&Y)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Plan, "核收(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "登记(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "批量增加(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "批量病人增加(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "补填病人(&A)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "重新核收(&T)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改样本号(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "批量修改样本号(&D)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_CardBound, "修改病人信息(&P)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "单个发送到仪器(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_BathSend, "批量发送到仪器(&S)"): cbrControl.BeginGroup = True

        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_TOQC, "置为质控(&Q)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "置为比对(&Y)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "查看比对(&B)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_QCRes, "查看本月质控(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, comMenu_LIS_TodayQC, "今日质控(&T)")
        Set cbrControl = .Add(xtpControlButton, comMenu_LIS_History, "历史质控(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_LJAverage, "均值LJ质控(&A)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "状态回滚(&Z)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "删除样本(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_MedRec, "批量删除无主(&L)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "取消核收(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Insert, "标本合并(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "糖耐量合并(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒收(&J)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Statistics_PositiveResults, "阳性结果反馈(&Y)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Statistics_Feedback, "反馈情况查询(&Z)")
    End With
    'conMenu_EditPopup
    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "报告(&E)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Report, "报告填写(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "批量调整(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Dilute, "标本稀释(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "报告审核(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_SendReport, "初审报告(&S)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Audit, "批量审核(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "取消审核(&U)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Seat_Set, "按病人审核(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "样本复查(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消复查(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "报告查询(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_ImportFromXML, "分析数据收集(&G)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_SignVerify, "验证签名(&S)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Import, "自动导入(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ApplyTo, "批量导入(&L)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_SaveSample, "标本保存(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_DropSample, "标本销毁(&H)")
    
    End With
'    '右键菜单
'    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_LIS_RightMenu, "右键菜单", -1, False)
'    cbrMenuBar.ID = conMenu_LIS_RightMenu
'    With cbrMenuBar.CommandBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "报告审核(&A)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "取消审核(&U)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "状态回滚(&Z)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "样本复查(&D)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消复查(&E)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "报告查询(&P)"): cbrControl.BeginGroup = True
'
'        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "发往仪器(&S)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_TOQC, "置为质控(&Q)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "置为比对(&Y)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "查看比对(&B)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "糖耐量合并(&E)")
'
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改样本号(&M)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "删除样本(&D)")
'
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒收(&J)"): cbrControl.BeginGroup = True
'
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "放弃(&C)")
'    End With
'    cbrMenuBar.Visible = False

'    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "费用(&C)", -1, False)
'    cbrMenuBar.ID = conMenu_EditPopup
'    With cbrMenuBar.CommandBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Price, "生成主费(&P)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingAdd, "附加费划价(&I)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "附加费记帐(&S)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "零费记录(&Z)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "修改附加费(&M)"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingDel, "删除附加费(&D)")
'    End With

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
            cbrPopControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "前一条(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "后一条(&L)")
        '----------------------------------------------------------------------------------------------------------------
        '用于快捷方式操作(PageUP,PageDown,Home,End)
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Reference_1, "PageUP"): cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Reference_2, "PageDown"): cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_MeetFinish, "Home"): cbrControl.Visible = False
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_MeetCancel, "End"): cbrControl.Visible = False
        '----------------------------------------------------------------------------------------------------------------
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Notify, "含未收费(&N)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_HideList, "隐藏列表(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, comMenu_LIS_ShowListHead, "选择显示列表"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "显示待核收"): 'cbrControl.BeginGroup = True
        If zlDatabase.GetPara("显示待核收", 100, 1208, "False") = "True" Then
            cbrControl.Checked = True
        End If

        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_LeaveMedi, "隐藏检验图形"): 'cbrControl.BeginGroup = True
        
        If zlDatabase.GetPara("隐藏检验图形", 100, 1208, "True") = "True" Then
            cbrControl.Checked = True
        End If
'        Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "列表选项(&O)")
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_PatientInfo, "病人信息(&P)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Find, "定位(&G)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "快速过滤(&K)")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Search, "组合查询(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_FindNext, "病人历次检验(&H)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportEdit, "标本操作日志(&A)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&F)"): cbrControl.BeginGroup = True
    End With
'-------------------------------------------------------------------------------------------------------------------------------------
    '综合查询插件菜单
    Dim i           As Long
    ReDim mobjPlugin(0) As zl9LisQuery_Def.clsLisQuery
    If Not clsPluginLoader Is Nothing Then
        clsPluginLoader.FindPlugins
        If clsPluginLoader.PluginCount > 0 Then
            ReDim mobjPlugin(clsPluginLoader.PluginCount) As zl9LisQuery_Def.clsLisQuery
            Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_PlugPopup, "外接程序(&A)", -1, False)
            cbrMenuBar.ID = conMenu_PlugPopup
            With cbrMenuBar.CommandBar.Controls
                
                For i = 0 To clsPluginLoader.PluginCount - 1
                    Set mobjPlugin(i) = clsPluginLoader.CreatePlugin(i)
                    If Not mobjPlugin(i) Is Nothing Then
                        mobjPlugin(i).Index = i
                        Set cbrControl = .Add(xtpControlButton, conMenu_PlugPopup * 1000# + 100 + i, mobjPlugin(i).Name)
                    End If
                    If i = 0 Then cbrControl.BeginGroup = True
                Next
            End With
        End If
    End If
'-------------------------------------------------------------------------------------------------------------------------------------

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With

    Set cbrControl = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "检验科室")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Report_DrugQuery, "检验科室")
    cbrCustom.ShortcutText = "科室"
    cbrCustom.Handle = Me.cboDept.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption

    Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Edit_NoPrint, "仪器小组")
    cbrMenuBar.ID = conMenu_Edit_NoPrint
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Owe, "所有小组")
    End With
    cbrMenuBar.Flags = xtpFlagRightAlign

    Set cbrCustom = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Report_Reports, "检验仪器")
    cbrCustom.ShortcutText = "检验仪器"
    cbrCustom.Handle = Me.cboMachine.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption

    Set cbrControl = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "组合项目")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = Me.cbrthis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Report_WorkLog, "检验仪器")
    cbrCustom.ShortcutText = "检验仪器"
    cbrCustom.Handle = Me.cboUnionItem.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    cbrCustom.Style = xtpButtonIconAndCaption

    '快键绑定
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F2, conMenu_Edit_Save
        .Add 0, VK_ESCAPE, conMenu_LIS_Cancel
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F4, conMenu_Manage_Plan
        .Add 0, VK_F8, conMenu_Manage_Regist
        .Add FCONTROL, Asc("T"), conMenu_Tool_Apply
        .Add FCONTROL, Asc("Z"), conMenu_Edit_SendBack
        .Add FCONTROL, VK_DELETE, conMenu_Manage_ClearUp
        .Add 0, VK_F7, conMenu_Manage_Report
        .Add 0, VK_F6, conMenu_Edit_Audit
        .Add FCONTROL, VK_LEFT, conMenu_View_Backward
        .Add FCONTROL, VK_RIGHT, conMenu_View_Forward
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add FCONTROL, Asc("F"), conMenu_Manage_Transfer_Force
        .Add 0, VK_F3, conMenu_View_Filter
        .Add 0, VK_HOME, conMenu_Tool_MeetFinish
        .Add 0, VK_END, conMenu_Tool_MeetCancel
        .Add 0, VK_PAGEUP, conMenu_Tool_Reference_1
        .Add 0, VK_PAGEDOWN, conMenu_Tool_Reference_2
        .Add FCONTROL, Asc("H"), conMenu_View_FindNext
        .Add 0, VK_F9, conMenu_Edit_QCRes
        .Add 0, VK_F11, conMenu_Manage_Logout
        
        .Add 0, VK_F10, conMenu_IDkind_Change
    End With

    '设置不常用菜单
'    With Me.cbrthis.Options
'        .AddHiddenCommand conMenu_File_PrintSet
'        .AddHiddenCommand conMenu_File_Excel
'        .AddHiddenCommand conMenu_View_Jump
'        .AddHiddenCommand conMenu_View_Refresh
'    End With
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "放弃")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Plan, "核收"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "登记")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "批量")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "回滚")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Insert, "合并")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Report, "填报告"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_SendReport, "初审")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审报告")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "补填")
'        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Price, "主费用"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): cbrControl.BeginGroup = True
'        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next


End Sub

Private Sub cboDept_Click()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim intLoop As Integer
    Dim blnSelect As Boolean                    '是否选择
    Dim strCoding As String                     '小组编码
    On Error GoTo errH
    mstrMachines = ""
    mstrMachineALL = ""
    rptList.Tag = ""
    If cboDept.ListCount > 0 Then
        '写入刷新
        mlngDeptID = Val(cboDept.ItemData(cboDept.ListIndex))
        
        gstrSql = "Select Distinct A.编码, A.名称" & vbNewLine & _
                "From 检验小组 A, 检验小组仪器 B, 检验仪器 C" & vbNewLine & _
                "Where A.ID = B.小组id And B.仪器id = C.ID And C.使用小组id = [1] order by a.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDeptID)
        
        Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Edit_NoPrint, True, True)
        With cbrMenuBar.CommandBar.Controls
            .DeleteAll
            .Add xtpControlButton, conMenu_View_Owe, "所有小组"
             If Not cbrMenuBar Is Nothing Then
                 Do Until rsTmp.EOF
                    Set cbrControl = .Add(xtpControlButton, conMenu_View_Owe, Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称")))
                    cbrControl.Checked = (Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称")) = mstrMachineGroup)
                    If cbrControl.Checked = True Then
                        blnSelect = True
                    End If
                    rsTmp.MoveNext
                 Loop
             End If
            If blnSelect = False Then
                cbrMenuBar.CommandBar.Controls(1).Checked = True
                mstrMachineGroup = "所有小组"
            End If
        End With
        
        
'        objLISComm.DeptID = mlngDeptID
        
        cboMachine.Clear
        
        If cboDept.ListCount > 0 Then
        
            cboMachine.AddItem "<所有仪器>": cboMachine.ItemData(cboMachine.NewIndex) = 0
            cboMachine.AddItem "<手工>": cboMachine.ItemData(cboMachine.NewIndex) = -1
            If InStr(mstrMachineGroup, "-") > 0 Then
                strCoding = Mid(mstrMachineGroup, 1, InStr(mstrMachineGroup, "-") - 1)
            End If
            If InStr(mstrPrivs, "所有科室") > 0 Then
                strSQL = "Select Distinct A.名称, A.ID, 1 As 更改,c.编码 ,A.微生物" & vbNewLine & _
                        "From 检验仪器 A, 检验小组仪器 B, 检验小组 C" & vbNewLine & _
                        "Where A.ID = B.仪器id And A.使用小组id = [1] And B.小组id = C.ID "
            Else
                strSQL = "Select Distinct D.ID, D.名称, C.更改,b.编码,D.微生物 " & vbNewLine & _
                        " From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
                        " Where A.小组id = B.ID And B.ID = C.小组id　and 人员id = [2] And C.仪器id = D.ID And D.使用小组id = [1] "
            End If
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptID, UserInfo.ID, strCoding)
            
            If mstrMachineGroup <> "所有小组" Then
                rsTmp.filter = "编码 = '" & strCoding & "'"
            End If
            '清空微生物仪器ID串
            mstrMachineID = ""
            Do Until rsTmp.EOF
                cboMachine.AddItem rsTmp("名称")
                cboMachine.ItemData(cboMachine.NewIndex) = rsTmp("Id")
                If rsTmp("微生物") = 1 Then
                    mstrMachineID = mstrMachineID & rsTmp("id") & ","
                End If
                If rsTmp("id") = mlngMachineID Then
                    cboMachine.ListIndex = cboMachine.NewIndex
                End If
                
                rsTmp.MoveNext
            Loop
            If cboMachine.ListCount > 0 And Trim(cboMachine.Text) = "" Then
                cboMachine.ListIndex = 0
                mlngMachineID = cboMachine.ItemData(cboMachine.ListIndex)
            End If
            
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                If rsTmp.EOF = False Then
                    rsTmp.filter = ""
                    rsTmp.MoveFirst
                    Do Until rsTmp.EOF
                        If Val(Nvl(rsTmp("更改"))) = 1 Then
                            mstrMachines = mstrMachines & ";" & rsTmp("ID")
                        End If
                        mstrMachineALL = mstrMachineALL & "," & rsTmp("ID")
                        rsTmp.MoveNext
                    Loop
                End If
            End If
            If mstrMachines <> "" Then mstrMachines = mstrMachines & ";"
        Else
            mlngMachineID = 0
            mstrMachines = ""
            mstrMachineALL = ""
        End If
    Else
        mlngDeptID = 0
        mstrMachines = ""
        mstrMachineALL = ""
    End If
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
'    RefreshData
    
'    '处理刷新后定位到指定记录
'    On Error Resume Next
'    Me.dkpMain.FindPane(Dkp_ID_List).Select
'    Me.rptList.SetFocus
End Sub

Private Sub cboExesItem_Click()
    Dim lngAdvice As Long
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean                                             '是否转出
    On Error GoTo errH
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            blnCurrMoved = (.Record(mCol.转出).Value = "√")
        End With
    End If
    
    strSQL = "select a.id as 医嘱ID, b.发送号 from 病人医嘱记录 a,病人医嘱发送 b " & vbCrLf & _
        " Where a.ID = b.医嘱id And a.相关id = [1] "
        
    If Me.cboExesItem.ListIndex <> -1 Then lngAdvice = Me.cboExesItem.ItemData(Me.cboExesItem.ListIndex)
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngAdvice)
    If rsTmp.EOF = False Then
        mclsExpenses.zlRefresh mlngDeptID, rsTmp(0) & ":" & rsTmp(1), blnCurrMoved
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboMachine_Click()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errH
    
    If cboMachine.ListCount > 0 Then
        '写入刷新
        mlngMachineID = cboMachine.ItemData(cboMachine.ListIndex)
    Else
        mlngMachineID = 0
    End If
    
'    strsql = "select distinct a.id,a.编码,a.名称 from 诊疗项目目录 a,诊疗执行科室 b,检验报告项目 c , 检验仪器项目 D " & _
             " where a.类别 = 'C' and (a.组合项目 = 1 or a.单独应用 = 1 ) " & _
             " and a.id = b.诊疗项目id and a.id = c.诊疗项目id and c.报告项目id = d.项目ID(+) "
    strSQL = "select distinct a.id,a.编码,a.名称 from 诊疗项目目录 a,诊疗执行科室 b,检验报告项目 c , 检验仪器项目 D " & _
             " where a.类别 = 'C' and (a.组合项目 = 1 or a.单独应用 = 1 ) " & _
             " and a.id = b.诊疗项目id and a.id = c.诊疗项目id(+) and c.报告项目id = d.项目ID(+) " & _
             " and (a.撤档时间 is null or a.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) "
                 
    '处理手工和有仪器的情况
    If cboMachine.ItemData(cboMachine.ListIndex) = -1 Then
        strSQL = strSQL & " And D.仪器ID is null "
    ElseIf cboMachine.ItemData(cboMachine.ListIndex) > 0 Then
        strSQL = strSQL & " And D.仪器ID = [1] "
    Else
        If Me.cboUnionItem.ListCount = 0 Then
            Me.cboUnionItem.Clear
            Me.cboUnionItem.AddItem "<所有组合项目>"
            Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = 0
            Me.cboUnionItem.AddItem "<未知项目>"
            Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = -1
            If Me.cboUnionItem.ListCount > 0 Then Me.cboUnionItem.ListIndex = 0
        End If
        If Not TabCtlWindow(5).Visible Then TabCtlWindow(5).Visible = True
        Exit Sub
    End If
    
    '处理科室
    If cboDept.ItemData(cboDept.ListIndex) > 0 Then
        strSQL = strSQL & " And B.执行科室ID = [2] "
    End If
    
    strSQL = strSQL & " order by a.编码 "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngMachineID, mlngDeptID, CDate("3000/1/1"))
    
    Me.cboUnionItem.Clear
    Me.cboUnionItem.AddItem "<所有组合项目>"
    Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = 0
    Me.cboUnionItem.AddItem "<未知项目>"
    Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = -1
    
    Do Until rsTmp.EOF
        Me.cboUnionItem.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        Me.cboUnionItem.ItemData(Me.cboUnionItem.NewIndex) = rsTmp("ID")
        rsTmp.MoveNext
    Loop
    '处理微生物仪器被选中时标本合并页签隐藏
    If InStr("," & mstrMachineID, "," & mlngMachineID & ",") > 0 Then
        TabCtlWindow(5).Visible = False
        '处理选中了该页签切换到第一个页签
        If TabCtlWindow(5).Selected Then
            TabCtlWindow(1).Selected = True
        End If
    Else
        TabCtlWindow(5).Visible = True
    End If
    If Me.cboUnionItem.ListCount > 0 Then Me.cboUnionItem.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
'    RefreshData
'
'    '处理刷新后定位到指定记录
'    On Error Resume Next
'    Me.dkpMain.FindPane(Dkp_ID_List).Select
'    Me.rptList.SetFocus
End Sub

Private Sub cboUnionItem_Click()
'    If Me.Visible = True Then
        On Error Resume Next
'        DoEvents
'        '等待窗体显示
'        Do While Me.Visible = False
'            DoEvents
'            Me.Show
'        Loop
        If mintLoadShow = 0 Then Exit Sub
        If Me.TabList.Item(1).Selected = True And Me.cboUnionItem.ListCount > 0 Then
            Call RefreshData1
        Else
            Call RefreshData
        End If
'    End If
End Sub

Private Sub cbo时间_Click()
    If Me.Visible = False Then Exit Sub
    If Me.TabList(0).Selected = True Then
        zlDatabase.SetPara "标本范围", cbo时间.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Me.dtpDate.Visible = (Me.cbo时间.Text = "自定义")
        Me.dtpDateEnd.Visible = (Me.cbo时间.Text = "自定义")
        Call RefreshData
    Else
        zlDatabase.SetPara "待核收范围", cbo时间.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Me.dtpDate.Visible = (Me.cbo时间.Text = "自定义")
        Me.dtpDateEnd.Visible = (Me.cbo时间.Text = "自定义")
        Call RefreshData1
    End If
    '刷新
    
End Sub

Private Sub cbrChild_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrCbo As CommandBarComboBox
    
    On Error GoTo errH
    
    Select Case Control.ID
        Case conMenu_File_RoomSet                                                               '定位
            If FindPatient(Control.Text) = True Then
                Control.Text = ""
            End If
            Control.SetFocus
            SendKeys "~"
        Case conMenu_View_Forward                                                               '前一条
            BackOrNextPatient 1
        
        Case conMenu_View_Backward                                                              '后一条
            BackOrNextPatient 2
        
        Case conMenu_Manage_RequestView                                                         '使用条码扫描
            Control.Checked = Not Control.Checked
            zlDatabase.SetPara "使用条码扫描", Control.Checked, 100, 1208
            
        Case conMenu_Manage_RequestPrint                                                        '连续输入
            Control.Checked = Not Control.Checked
            zlDatabase.SetPara "连续输入", Control.Checked, 100, 1208
            '是否连续输入
            mintContinue = IIf(Control.Checked, 1, 0)
            If mintContinue = 1 Then
                Me.cbrthis.FindControl(, conMenu_Manage_Regist, , True).Caption = "连续登记"
                Me.cbrthis.FindControl(, conMenu_Manage_Plan, , True).Caption = "连续核收"
            Else
                Me.cbrthis.FindControl(, conMenu_Manage_Regist, , True).Caption = "登记"
                Me.cbrthis.FindControl(, conMenu_Manage_Plan, , True).Caption = "核收"
            End If
            Me.cbrthis.RecalcLayout
        Case conMenu_Manage_RequestBatPrint                                                         '保存后直接审核
            Control.Checked = Not Control.Checked
            zlDatabase.SetPara "保存后直接审核", Control.Checked, 100, 1208
        Case XTP_ID_WINDOW_LIST '显示备注窗体
            Control.Checked = Not Control.Checked
            zlDatabase.SetPara "显示检验备注", Control.Checked, 100, 1208
            Call mfrmWrite.Resize
            Call mfrmWrite2.Resize
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrChild_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    Left = -120
End Sub


Private Sub cbrChild_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo errH
    
    Select Case Control.ID
        
        Case conMenu_Manage_RequestView                                                         '使用条码扫描
            Control.Checked = Control.Checked
        
        Case conMenu_Manage_RequestPrint                                                        '连续输入
            Control.Checked = Control.Checked
        
        Case conMenu_Manage_RequestBatPrint                                                     '保存后直接审核
            Control.Checked = Control.Checked
        Case XTP_ID_WINDOW_LIST                                                                 '显示检验备注
            Control.Checked = Control.Checked
        
        Case conMenu_View_Forward                                                               '前一条,后一条
            If mintEditState <> 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
                If Me.rptList.Rows.Count <= 1 Then
                    Control.Enabled = False
                Else
                    If Not rptList.FocusedRow Is Nothing Then
                        If Me.rptList.FocusedRow.Index = 0 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                End If
            End If
        Case conMenu_View_Backward
            
            If mintEditState = 4 Or mintEditState = 5 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
                If Me.rptList.Rows.Count <= 1 Then
                    Control.Enabled = False
                Else
                    If Not rptList.FocusedRow Is Nothing Then
                        If Me.rptList.Rows.Count - 1 = Me.rptList.FocusedRow.Index Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                End If
            End If
        Case conMenu_File_RoomSet                                                               '定位
            If mintEditState <> 0 Then
                txtGoto.Enabled = False
            Else
                txtGoto.Enabled = True
            End If
        Case conMenu_Manage_Transfer_Send, conMenu_Edit_UnArchive                               '费用
            Control.Visible = (Me.TabCtlWindow.Selected.Index = 4)
    End Select

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim intLoop As Integer
    
    On Error GoTo errH
           
    Select Case Control.ID
        
        '''''''''''''''''''''''''''''''''''''''文件''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_File_PrintSet                                                      '打印设置
             PrintSetup
            
        Case conMenu_File_Preview                                                       '报告预览
            ReportPrint False
        
        Case conMenu_File_Print                                                         '报告打印

            If InStr(",7,8,", CStr(Me.rptList.FocusedRow.Record(mCol.执行状态).Icon)) = 0 Then
                If MsgBox("您的检验单没有审核!是否确定要打印!", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            ReportPrint True
        
        Case conMenu_File_BatPrint                                                      '批量批印
'            Call frmLisStationPrint.ShowEdit(Me, mlngDeptID)
            Call frmBatchAction.ShowMe(Me, 1, mlngMachineID, mstrPrivs, , , , mlngDeptID, mstrAuditingManID)
        Case conMenu_Edit_Save                                                          '保存
            Call SaveDisposal(mFileS.保存)
            
        Case conMenu_LIS_Cancel                                                         '放弃
            mintHandleState = 0
            Call SaveDisposal(mFileS.放弃)
            
        Case conMenu_File_RowPrint                                                      '清单打印
            'Call zlRptPrint(1)
            Call frmLabReport.Show
        Case conMenu_Edit_Privacy                                                       '审核人登陆
            Call AuditingRegister
            
        Case conMenu_File_Parameter                                                     '参数设置
            SetParameter
            
        Case conMenu_Tool_Monitor                                                       '酶标仪设置
'            If frmMBSetup.ShowMe(Me) Then objLISComm.InitMBPara
            frmLabMB.ShowMe Me, mlngMachineID
            
        Case conMenu_File_Exit                                                          '退出
            Unload Me
        
        ''''''''''''''''''''''''''''''''''''''样本''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Plan                                                        '核收
            Call SampleDisposal(mActS.核收)
        
        Case conMenu_Manage_Regist                                                      '登记
            Call SampleDisposal(mActS.登记)
        
        Case conMenu_Edit_NewParent                                                     '批量增加
            Call SampleDisposal(mActS.批量增加)
        
        Case conMenu_Manage_Logout                                                      '批量病人增加
            frmAddPatient.ShowMe Me, mlngMachineID, mlngDeptID
            
        Case conMenu_Manage_Receive                                                     '补填病人
            mintHandleState = 1
            Call SampleDisposal(mActS.补填病人)
        
        Case conMenu_Manage_Transfer                                                    '重新核收
            Call SampleDisposal(mActS.重新核收)
        
        Case conMenu_Edit_ModifyParent                                                  '修改样本号
            Call SampleDisposal(mActS.修改样本号)
        
        Case conMenu_Manage_Reset                                                       '批量修改样本号
            Call SampleDisposal(mActS.批量修改样本号)
            
'        Case conMenu_Edit_CardBound                                                     '修改病人信息
'            Call SampleDisposal(mActS.修改病人信息)
            
        Case conMenu_Tool_Apply                                                         '发送仪器
            mbln手工发送杯号 = True
            Call SampleDisposal(mActS.发送仪器)
            mbln手工发送杯号 = False
        Case conMenu_Tool_BathSend                                                      '批量发送
            
            Call SampleDisposal(mActS.批量发送到仪器)
        Case conMenu_LIS_TOQC                                                           '置为质控
            Call SampleDisposal(mActS.置为质控)
            
        Case comMenu_LIS_TodayQC                                                        '今日质控权限
            frmQCTodayList.Show vbModal, Me
        
        Case comMenu_LIS_History                                                        '历史质控权限
            frmQCHistory.Show vbModal, Me
            
        Case conMenu_LIS_LJAverage                                                      '均值质控图
            ShowLJAverage
        
        Case conMenu_Edit_QCRes                                                         '查看本月质控
            Call frmLabMainLJ.ShowMe(mlngKey, Me, mlngMachineID)
        
        Case conMenu_Tool_Analyse                                                       '置为比对
            Call SampleDisposal(mActS.置为对比)
            
        Case conMenu_Manage_ReportView                                                  '查看比对
            Call frmQCContrast.ShowMe(Me, mlngMachineID)
        
        Case conMenu_Edit_SendBack                                                      '状态回滚
            Call SampleDisposal(mActS.状态回滚)
        
        Case conMenu_Manage_ClearUp                                                     '删除无主标本
            Call SampleDisposal(mActS.删除无主标本)
        
        Case conMenu_Tool_MedRec                                                        '指量删除无主
'            frmLisStationBatch.ShowEdit Me, mlngDeptID
'            Call RefreshData
            frmBatchAction.ShowMe Me, 3, mlngMachineID, , , , , mlngDeptID, mstrAuditingManID
        Case conMenu_Edit_DeleteParent                                                  '取消核收
            Call SampleDisposal(mActS.置为无主)
            
        Case conMenu_Edit_Insert                                                        '合并标本
            Call SampleDisposal(mActS.合并标本保存)
        
        Case conMenu_Edit_Surplus                                                       '糖耐量合并
            frmLabBloodSugar.ShowMe Me, mlngMachineID, mlngKey
            Call RefreshData
            
        Case conMenu_Manage_Refuse                                                      '拒收
            Call SampleDisposal(mActS.拒收)
        Case conMenu_Statistics_PositiveResults                                         '阳性结果反馈
            ShowPositiveResults (0)
        
        Case conMenu_Statistics_Feedback                                                '反馈情况查询
            ShowPositiveResults (1)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '''''''''''''''''''''''''''''''''''''''报告''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Report                                                      '报告填写
            If Me.TabCtlWindow.Selected.Index = 2 Then
                mintHandleState = 3
                Call ReportDisposal(mActR.填写三级报告)
            Else
                mintHandleState = 2
                Call ReportDisposal(mActR.填写报告)
            End If
            
        Case conMenu_Edit_Adjust                                                        '批量调整
            Call ReportDisposal(mActR.批量调整报告)
        
        Case conMenu_Edit_Dilute                                                        '标本稀释
            frmDiluteSample.ShowMe Me, mlngKey
            mfrmWrite.zlRefresh mlngKey
        
        Case conMenu_Edit_Audit                                                         '报告审核
            Call ReportDisposal(mActR.审核报告)
            
        Case conMenu_LIS_SendReport                                                     '报告发送
            Call ReportDisposal(mActR.发送报告)
        
        Case conMenu_Manage_Audit                                                       '批量审核
            Call ReportDisposal(mActR.批量审核报告)
        
        Case conMenu_Edit_ClearUp                                                       '取消审核
            Call ReportDisposal(mActR.审核取消)
        
        Case conMenu_Edit_Seat_Set                                                      '按病人审核
            Call ReportDisposal(mActR.按病人审核)
        
        Case conMenu_Manage_Redo                                                        '重做结果
            Call ReportDisposal(mActR.重做结果)
        
        Case conMenu_Manage_Undone                                                      '取消重做
            Call ReportDisposal(mActR.取消重做)
       
        Case conMenu_Manage_Transfer_Force                                              '病人报告查询
            If Me.rptList.FocusedRow Is Nothing Then
                frmLabMainFindRePort.ShowMe -1, Me, mstrPrivs
            Else
                frmLabMainFindRePort.ShowMe Val(Me.rptList.FocusedRow.Record(mCol.病人ID).Value), Me, mstrPrivs
            End If
'            Me.SetFocus: Me.TabList.SetFocus: Me.rptList.SetFocus
'            Me.dkpMain.FindPane(Dkp_ID_List).Select
                    
        Case conMenu_File_ImportFromXML                                                 '分析数据采集
            frmLabAnalyseData.ShowMe Me, mlngMachineID
                            
        Case conMenu_LIS_SignVerify                                                     '验证签名
            Call ReportDisposal(mActR.验证签名)
                            
        Case conMenu_Edit_Import                                                        '自动导入
            GetSaveSetup 1
        
        Case conMenu_Edit_ApplyTo                                                       '批量导入
            GetSaveSetup 2
            
        Case conMenu_LIS_SaveSample                                                     '标本保存(存放）
            frmlabONSample.Show vbModal, Me
            
        Case conMenu_LIS_DropSample                                                     '标本销毁
            frmlabDropSample.Show vbModal, Me
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '''''''''''''''''''''''''''''''''''''''费用'''''''''''''''''''''''''''''''''''''''''''''''''
'        Case conMenu_Edit_Price                                                         '生成主费
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_Append
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Manage_ThingAdd                                                    '附加费划价
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_NewItem * 10# + 1
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Edit_ModifyParent                                                  '附加费记帐
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_NewItem * 10# + 2
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Edit_NewItem                                                       '零费记录
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_NewItem * 10# + 3
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Manage_ThingModi                                                   '修改附加费
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_Modify
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
'
'        Case conMenu_Manage_ThingDel                                                    '删除附加费
'            Set cbrControl = Me.cbrChild.FindControl(, conMenu_Edit_Append, False, True)
'            cbrControl.ID = conMenu_Edit_Delete
'            mclsExpenses.zlExecuteCommandBars cbrControl
'            cbrControl.ID = conMenu_Edit_Append
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        '''''''''''''''''''''''''''''''''''''''查看'''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_ToolBar_Button                                                '标准按钮
            Control.Checked = Not Control.Checked
            Me.cbrthis(2).Visible = Control.Checked
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text                                                  '文本标签
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_ToolBar_Size                                                  '大图标
            Control.Checked = Not Control.Checked
            Me.cbrthis.Options.LargeIcons = Not Me.cbrthis.Options.LargeIcons
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_StatusBar                                                     '状态栏
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Control.Checked
            Me.cbrthis.RecalcLayout
        
        Case conMenu_View_Forward                                                       '前一条
            BackOrNextPatient 2
        
        Case conMenu_View_Backward                                                      '后一条
            BackOrNextPatient 1
        
        Case conMenu_Tool_Reference_1                                                   'PAGEUP
            ShortWork mSWork.Key_PageUP
            
        Case conMenu_Tool_Reference_2                                                   'PAGEDOWN
            ShortWork mSWork.Key_PageDown
        
        Case conMenu_Tool_MeetFinish                                                    'HOME
            ShortWork mSWork.Key_Home
        
        Case conMenu_Tool_MeetCancel                                                    'End
            ShortWork mSWork.Key_End
            
        Case conMenu_View_Notify                                                        '含未收费
            Control.Checked = Not Control.Checked
        
        Case conMenu_LIS_HideList                                                       '隐藏列表
            Control.Checked = Not Control.Checked
            ShowOrHideItem Control, Dkp_ID_List
        
        Case conMenu_Manage_LeaveMedi                                                   '隐藏检验图形
            Control.Checked = Not Control.Checked
            ShowOrHideItem Control, Dkp_ID_Image
            
        Case comMenu_LIS_ShowListHead                                                   '选择列表
            If TabList.Selected.Index = 0 Then
                ShowHideListHead Me.rptList.Columns, frmPublicFieldChooser.ShowMe(Me, Me.rptList.Columns)
            Else
                ShowHideListHead Me.rptList1.Columns, frmPublicFieldChooser.ShowMe(Me, Me.rptList1.Columns)
            End If
            
            
        Case conMenu_Manage_ReGet                                                       '显示待核收
            Control.Checked = Not Control.Checked
            Me.TabList.Item(1).Visible = Control.Checked
        
        Case conMenu_LIS_PatientInfo                                                    '病人信息
            If Not Me.rptList.FocusedRow Is Nothing Then
                frmDegreeCard.ShowInfo Me, Val(Me.rptList.FocusedRow.Record(mCol.病人ID).Value)
            End If
        
        Case conMenu_View_Find                                                          '定位
            If Me.txtGoto.Enabled Then Me.txtGoto.SetFocus
            
        Case conMenu_View_Filter                                                        '快速过滤
            Call QUFilter
        
        Case conMenu_View_FindNext                                                      '病人历次检验
            Call QuickFindPatient
        
        Case conMenu_Manage_ReportEdit                                                  '查看审核记录
            frmLabAuditingCourse.ShowMe Me, mlngKey
        
        Case conMenu_View_Refresh                                                       '刷新
            '清空过滤条件
            Me.rptList.Tag = ""
            If Me.TabList.Item(0).Selected = True Then
                Call RefreshData
            Else
                Call RefreshData1
            End If
            
        
        ''''''''''''''''''''''''''''''''''''''帮助''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Help_Help                                                          '帮助主题
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_Web                                                           'WEB上的
            Call zlHomePage(hWnd)
        
        Case conMenu_Help_Web_Home                                                      '主页
            Call zlHomePage(Me.hWnd)
        
        Case conMenu_Help_Web_Mail                                                      '发送反馈
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_Help_About                                                         '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        
        ''''''''''''''''''''''''''''''''''科室列表和仪器列表''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Report_DrugQuery                                                   '科室选择
            
        Case conMenu_Report_Reports                                                     '选择仪器
        
        Case conMenu_View_Owe                                                           '选择分组
            Set cbrMenuBar = Me.cbrthis.ActiveMenuBar.FindControl(, conMenu_Edit_NoPrint, True, True)
            rptList.Tag = ""
            With cbrMenuBar.CommandBar
                For intLoop = 1 To .Controls.Count
                    .Controls(intLoop).Checked = (.Controls(intLoop).Caption = Control.Caption)
                Next
            End With
            mstrMachineGroup = Control.Caption
            Call cboDept_Click
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_IDkind_Change
             mfrmRequest.IdKindChange
        Case conMenu_File_PrintBedCard                                                '重打条码
            '21436 问题
            Call PrintBarcord
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "检验科室=" & mlngDeptID, "检验仪器ID=" & mlngMachineID, "标本ID=" & mlngKey)
            Else
                Select Case Me.TabCtlWindow.Selected.Index
                    Case 4  '费用
                        mclsExpenses.zlExecuteCommandBars Control
                    Case 6  '门诊
                        mclsOutAdvices.zlExecuteCommandBars Control
                    Case 7  '住院
                        mclsInAdvices.zlExecuteCommandBars Control
                    Case Else
                        '------------------------------------------------------
                        '插件相关
                        Dim lngCount As Long
                        If Not clsPluginLoader Is Nothing Then
                            If clsPluginLoader.PluginCount > 0 Then
                                lngCount = Control.ID - conMenu_PlugPopup * 1000# - 100
                                If Mid(Control.ID, 1, 1) = conMenu_PlugPopup And lngCount >= 0 Then
                                    If mobjPlugin(lngCount) Is Nothing Then
                                        MsgBox "不能调用插件!", vbExclamation
                                        Exit Sub
                                    End If
                                    ' 插件能提供它的状态
                                    mobjPlugin(lngCount).InitQuery Me
                                    '执行插件 设为非模态窗体显示
                                    mobjPlugin(lngCount).DoAction Query_ShowModeless
                                End If
                            End If
                        End If
                        '------------------------------------------------------
                End Select
            End If
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowPositiveResults(ByVal intType As Integer) As Boolean
    '功能       调用传染病反馈窗体,和传染病报告窗体
    'intType    0-阳性结果反馈 1-反馈情况查询
    Dim lngKey As Long      '标本ID
    Dim strSQL As String
    Dim rsPatient As Recordset  '病人信息
    Dim rsJYK As Recordset '检验科
    Dim rsBuMen As Recordset '申请科室部门ID
    Dim intSendOk As Integer
    Dim objPublicAdvic As Object
    
    On Error GoTo ErrHand
    
    Set objPublicAdvic = CreateObject("zlPublicAdvice.clsPublicAdvice")
    Call objPublicAdvic.InitDisease(gcnOracle, 100)
    
    If intType = 0 Then
        With Me.rptList
            If .FocusedRow Is Nothing Then
                MsgBox "请选择一个标本", vbInformation
                Exit Function
            End If
            lngKey = Val(rptList.FocusedRow.Record(mCol.ID).Value)
        End With
        If lngKey = 0 Then
            MsgBox "请选择一个标本", vbInformation
        Else
            strSQL = "Select Distinct a.Id, a.病人id, a.主页id, a.标本类型 标本名称,a.申请科室id 送检科室ID, " & _
                     "b.标本送出时间 送检时间, b.送检人 送检医生, a.申请科室id, a.挂号单, " & _
                     " a.检验时间, a.住院号, a.执行科室id  登记科室id,a.医嘱id 医嘱ID From 检验标本记录 A, 病人医嘱发送 B " & _
                     "Where a.医嘱id = b.医嘱id and a.Id = [1] and a.医嘱id is not null"
            Set rsPatient = zlDatabase.OpenSQLRecord(strSQL, "查询HIS病人ID", lngKey)
            
'            strSQL = "select ID 送检科室ID from 部门表 where 名称=[1]"
'            Set rsBuMen = ComOpenSQL(Sel_His_DB, strSQL, "查询HIS病人ID", IIf(IsNull(rsPatient("申请科室")), "", rsPatient("申请科室")))
            
            If rsPatient.RecordCount < 1 Or (IsNull(rsPatient("主页ID")) And IsNull(rsPatient("挂号单"))) Then
                MsgBox "没有查找到病人相关的医嘱信息", vbInformation
            Else
                If IsNull(rsPatient("送检时间")) Then
                    If IsNull(rsPatient("检验时间")) Then
                        intSendOk = objPublicAdvic.ShowDisRegist(Me, 0, , IIf(IsNull(rsPatient("病人ID")), 0, rsPatient("病人ID")), _
                                                       IIf(IsNull(rsPatient("主页ID")), 0, rsPatient("主页ID")), _
                                                       IIf(IsNull(rsPatient("挂号单")), "", rsPatient("挂号单")), _
                                                       IIf(IsNull(rsPatient("医嘱ID")), 0, rsPatient("医嘱ID")), _
                                                       IIf(IsNull(rsPatient("登记科室ID")), 0, rsPatient("登记科室ID")), , _
                                                       IIf(IsNull(rsPatient("送检科室ID")), 0, rsPatient("送检科室ID")), _
                                                       IIf(IsNull(rsPatient("送检医生")), "", rsPatient("送检医生")), _
                                                       IIf(IsNull(rsPatient("标本名称")), "", rsPatient("标本名称")))
                    Else
                        intSendOk = objPublicAdvic.ShowDisRegist(Me, 0, , IIf(IsNull(rsPatient("病人ID")), 0, rsPatient("病人ID")), _
                                                       IIf(IsNull(rsPatient("主页ID")), 0, rsPatient("主页ID")), _
                                                       IIf(IsNull(rsPatient("挂号单")), "", rsPatient("挂号单")), _
                                                       IIf(IsNull(rsPatient("医嘱ID")), 0, rsPatient("医嘱ID")), _
                                                       IIf(IsNull(rsPatient("登记科室ID")), 0, rsPatient("登记科室ID")), , _
                                                       IIf(IsNull(rsPatient("送检科室ID")), 0, rsPatient("送检科室ID")), _
                                                       IIf(IsNull(rsPatient("送检医生")), "", rsPatient("送检医生")), _
                                                       IIf(IsNull(rsPatient("标本名称")), "", rsPatient("标本名称")), , _
                                                       CDate(rsPatient("检验时间")))
                    End If
                Else
                    If IsNull(rsPatient("检验时间")) Then
                        intSendOk = objPublicAdvic.ShowDisRegist(Me, 0, , IIf(IsNull(rsPatient("病人ID")), 0, rsPatient("病人ID")), _
                                                       IIf(IsNull(rsPatient("主页ID")), 0, rsPatient("主页ID")), _
                                                       IIf(IsNull(rsPatient("挂号单")), "", rsPatient("挂号单")), _
                                                       IIf(IsNull(rsPatient("医嘱ID")), 0, rsPatient("医嘱ID")), _
                                                       IIf(IsNull(rsPatient("登记科室ID")), 0, rsPatient("登记科室ID")), _
                                                       CDate(rsPatient("送检时间")), _
                                                       IIf(IsNull(rsPatient("送检科室ID")), 0, rsPatient("送检科室ID")), _
                                                       IIf(IsNull(rsPatient("送检医生")), "", rsPatient("送检医生")), _
                                                       IIf(IsNull(rsPatient("标本名称")), "", rsPatient("标本名称")))
                    Else
                        intSendOk = objPublicAdvic.ShowDisRegist(Me, 0, , IIf(IsNull(rsPatient("病人ID")), 0, rsPatient("病人ID")), _
                                                       IIf(IsNull(rsPatient("主页ID")), 0, rsPatient("主页ID")), _
                                                       IIf(IsNull(rsPatient("挂号单")), "", rsPatient("挂号单")), _
                                                       IIf(IsNull(rsPatient("医嘱ID")), 0, rsPatient("医嘱ID")), _
                                                       IIf(IsNull(rsPatient("登记科室ID")), 0, rsPatient("登记科室ID")), _
                                                       CDate(rsPatient("送检时间")), _
                                                       IIf(IsNull(rsPatient("送检科室ID")), 0, rsPatient("送检科室ID")), _
                                                       IIf(IsNull(rsPatient("送检医生")), "", rsPatient("送检医生")), _
                                                       IIf(IsNull(rsPatient("标本名称")), "", rsPatient("标本名称")), , _
                                                       CDate(rsPatient("检验时间")))
                    End If
                End If
            End If
        End If
    ElseIf intType = 1 Then
        If mlngDeptID <> 0 Then
            Call objPublicAdvic.ShowDisQuery(Val(mlngDeptID))
        End If
    End If
    Set objPublicAdvic = Nothing
    Exit Function
ErrHand:
    MsgBox ("调用阳性结果反馈窗体出错!" & vbCrLf & "错误信息:" & Err.Description & "(" & Err.Number & ")"), vbQuestion, "出错函数:ShowPositiveResults"
End Function

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    On Error Resume Next
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbrthis_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl
    On Error Resume Next
    Select Case Me.TabCtlWindow.Selected.Index
        Case 4
            Select Case CommandBar.Parent.ID
            Case conMenu_Edit_NewItem '补费
                With CommandBar.Controls
                    .DeleteAll
                    '扩1位,为了使用快捷键
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 1, "收费单据(&1)")
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 2, "记帐单据(&2)")
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 3, "零耗费用(&3)")
                    With cbrthis.KeyBindings
                        .Add FCONTROL, vbKeyN, conMenu_Edit_NewItem * 10# + 1
                        .Add FCONTROL, vbKeyB, conMenu_Edit_NewItem * 10# + 2
                    End With
                End With
            End Select
'            Call mclsExpenses.zlPopupCommandBars(CommandBar)
        Case 6
            Select Case CommandBar.Parent.ID
            Case conMenu_Edit_Compend '报告
                With CommandBar.Controls
                    If .Count = 0 Then
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 1, "查阅报告(&W)"
                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 2, "打印报告(&P)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告(&V)"
                    End If
                End With
            End Select
'            Call mclsOutAdvices.zlPopupCommandBars(CommandBar)
        Case 7
            Select Case CommandBar.Parent.ID
            Case conMenu_Edit_Compend '报告
                With CommandBar.Controls
                    If .Count = 0 Then
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 1, "查阅报告(&W)"
                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 2, "打印报告(&P)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告(&V)"
                    End If
                End With
            End Select
'            Call mclsInAdvices.zlPopupCommandBars(CommandBar)
    End Select
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRowCount As Long                                   '当前显示总行数
    Dim intSampleType  As Integer                             '标本类型 = 3 质控   = 4 比对 = -1 普通
    Dim strSource As String                                   '从属来源
    Dim strExeState As String                                 '执行状态 =已检验/检验中
    Dim blMicrobe As Integer                                  '微生物 =true 是微生物
    Dim intReportCount As Integer                             '报告总数
    Dim strPatienName As String                               '姓名
    Dim lngMachineID As Long                                  '仪器ID
    Dim blWaiteDispose As Boolean                             '是否在等待处理列表
    Dim lngExecDept As Long                                   '执行科室ID
    Dim blnIF As Boolean                                      '判断条件
    Dim lngAdvice As Long                                     '医嘱ID
    Dim str初审人 As String                                   '初审人
    Dim blnExec As Boolean                                    '选择所有科室时不能进行补填、填写结果等操作
    Dim lngSendReport As Long                                 '发送报告 <>0 有报告
    Dim bln删除无主标本 As Boolean                            '删除无主标本
    
    On Error GoTo errH
        
    'CSBmk_CS <If Me.Visible = False>
    If Me.Visible = False Then Exit Sub
    
    '读入当前行的信息(用于判断是否Disabled)
    With Me.rptList
        If Not .FocusedRow Is Nothing Then
            lngRowCount = .Rows.Count
            If .Rows.Count = 0 Then Exit Sub
            intSampleType = .FocusedRow.Record(mCol.标本类型).Icon
            lngExecDept = Val(.FocusedRow.Record(mCol.执行科室ID).Value)
            intReportCount = Val(.FocusedRow.Record(mCol.结果次数).Value)
            lngMachineID = Val(.FocusedRow.Record(mCol.仪器id).Value)
            lngAdvice = Val(.FocusedRow.Record(mCol.医嘱id).Value)
            blMicrobe = IIf(Val(.FocusedRow.Record(mCol.微生物标本).Value) = 1, True, False)
            lngSendReport = Val(.FocusedRow.Record(mCol.报告发送).Value)
            str初审人 = .FocusedRow.Record(mCol.初审人).Value
            If .FocusedRow.Record(mCol.执行状态).Value = "已检验" Or .FocusedRow.Record(mCol.执行状态).Value = "已打印" Then
                strExeState = "已检验"
            ElseIf .FocusedRow.Record(mCol.执行状态).Value = "初审" Then
                strExeState = "初审"
            Else
                strExeState = "检验中"
            End If
            strPatienName = .FocusedRow.Record(mCol.姓名).Value
            strSource = .FocusedRow.Record(mCol.所属情况).Value
            If strSource = "" Or strSource = "无主" Then
                If strPatienName = "" Then
                    strSource = "无主"
                Else
                    strSource = "院外"
                End If
            End If
        End If
    End With
    
    '用于判断是不是有操作当前科室报告单权限
    If Me.TabList.Item(0).Selected = True Then
        If InStr(mstrPrivs, "所有科室") = 0 Then
            If mlngMachineID > 0 Or lngMachineID > 0 Then
                blnIF = InStr(";" & Replace(mstrMachineALL, ",", ";") & ";", ";" & IIf(lngMachineID = 0, mlngMachineID, lngMachineID) & ";")
            Else
                blnIF = True
            End If
        Else
            blnIF = True
        End If
    Else
        blnIF = True
    End If
    
    If InStr(";" & mstrPrivs & ";", ";删除无主标本;") > 0 Then bln删除无主标本 = True
    blnExec = InStr(mstrMachines, ";" & IIf(lngMachineID <= 0, mlngMachineID, lngMachineID) & ";")
    '手工项目不进行控制
    If blnExec = False Then
        If mlngMachineID = -1 Or lngMachineID = 0 Then blnExec = True
    End If
'    blnIF = False
    blWaiteDispose = Me.TabList.Item(1).Selected
    
    '处理列表和TAB控件在编辑时不能改变
    If blWaiteDispose = True Then
        If mintEditState = 4 Or mintEditState Then
            Me.PicList.Enabled = False
        Else
            Me.PicList.Enabled = True
        End If
    Else
        '只有连继输入时才限制选择列表
        Me.PicList.Enabled = (Me.rptList.Tag <> "Continue")
    End If
    If mintEditState = 5 Or mintEditState = 1 Or mintEditState = 2 Or mintEditState = 4 Then
        Me.TabCtlWindow.Item(2).Enabled = False
        Me.TabCtlWindow.Item(3).Enabled = False
    Else
        Me.TabCtlWindow.Item(2).Enabled = True
        Me.TabCtlWindow.Item(3).Enabled = True
    End If
    
    Select Case Control.ID
'        '''''''''''''''''''''''''''''''''''''''''''''''''文件'''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_File_PrintSet                                                                      '打印设置
            Control.Enabled = (Me.rptList.Records.Count > 0 And mintEditState = 0)
        Case conMenu_File_RowPrint                                                                      '清单打印
            Control.Enabled = (Me.rptList.Records.Count > 0 And mintEditState = 0)
        Case conMenu_File_BatPrint                                                                      '批量打印
            If InStr(1, mstrPrivs, "批量打印") <= 0 Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (Me.rptList.Records.Count > 0 And mintEditState = 0)
            End If
        Case conMenu_File_Preview, conMenu_File_Print                                                   '报告预览,报告打印
            If InStr(1, mstrPrivs, "报告打印") <= 0 Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If InStr(1, mstrPrivs, "未审核打印") > 0 Or strExeState = "已检验" Or strExeState = "初审" Then
                    Control.Enabled = (Me.rptList.Records.Count > 0 And mintEditState = 0 _
                    And IIf(strSource = "无主", InStr(1, mstrPrivs, "无主打印") > 0, True))
                Else
                    Control.Enabled = False
                End If
            End If
        Case conMenu_Edit_Save, conMenu_LIS_Cancel, conMenu_IDkind_Change                               '保存,放弃,切换
            Control.Enabled = (mintEditState <> 0 And blnIF = True And blnExec = True)
        Case conMenu_Manage_Refuse                                                                      '拒收
'            Control.Enabled = (mintEditState = 1 Or mintEditState = 0 And strExeState = "检验中" _
'            And strSource <> "无主" And intSampleType = -1)
            If Me.rptList1.FocusedRow Is Nothing Or blnIF = False Then
                Control.Enabled = False
            Else
                Control.Enabled = (mintEditState = 0 And Me.rptList1.FocusedRow.Record(mRCol.执行状态).Value <> 2)
            End If
'            Control.Enabled = (Not Me.rptList1.FocusedRow Is Nothing And mintEditState = 0)
        Case conMenu_File_Parameter                                                                     '参数设置
            If InStr(1, mstrPrivs, "参数设置") <= 0 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (mintEditState = 0)
            End If
        Case conMenu_Tool_Monitor                                                                       '酶标仪设置
            Control.Enabled = (mintEditState = 0)
'        ''''''''''''''''''''''''''''''''''''''''''''''''样本''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Plan                                                                        '核收
            If InStr(1, mstrPrivs, "核收标本") <= 0 Or blnIF = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If blWaiteDispose = False Then
                    Control.Enabled = (mintEditState = 0 And mlngDeptID > 0)
                Else
                    Control.Enabled = (Me.rptList1.Rows.Count > 0 And Not Me.rptList1.FocusedRow Is Nothing And (mintEditState = 0) _
                                       And mlngDeptID > 0)
                End If
            End If
        
        Case conMenu_Manage_Regist                                                                          '登记
            If InStr(1, mstrPrivs, "直接申请") <= 0 Or blWaiteDispose = True Or blnIF = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (mintEditState = 0 And mlngDeptID > 0)
            End If
        
        Case conMenu_Manage_Logout                                                                          '批量病人增加登记
            If InStr(1, mstrPrivs, "直接申请") <= 0 Or blWaiteDispose = True Or blnIF = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (mintEditState = 0 And mlngDeptID > 0)
            End If
        Case conMenu_Edit_NewParent                                                                         '批量无主增加
            If InStr(1, mstrPrivs, "直接申请") <= 0 Or blWaiteDispose = True Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (mintEditState = 0 And mlngDeptID > 0)
            End If
                    
        Case conMenu_Manage_Receive                                                                         '补填病人
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or strExeState = "初审" _
               Or intSampleType <> -1 And strSource <> "无主" Or InStr(1, mstrPrivs, "无主处理") <= 0 _
               Or intSampleType = 3 Or blWaiteDispose = True Or blnIF = False Or mlngDeptID = 0 Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
                    
        Case conMenu_LIS_TOQC                                                                               '置为质控
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or strExeState = "初审" _
               Or intSampleType <> -1 Or strSource <> "无主" Or InStr(1, mstrPrivs, "无主处理") <= 0 _
               Or intSampleType = 3 Or blWaiteDispose = True Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Manage_Transfer                                                                        '重新核收
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or strExeState = "初审" _
               Or intSampleType <> -1 Or blWaiteDispose = True Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Tool_Apply                                                                             '发送仪器
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or strExeState = "初审" Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Edit_ModifyParent                                                      '修改样本
            If InStr(1, mstrPrivs, "修改标本号") <= 0 Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or blWaiteDispose = True Or strExeState = "初审" Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
'        Case conMenu_Edit_CardBound                                                         '修改病人信息
'            If InStr(1, mstrPrivs, "修改病人信息") <= 0 Or blnIF = False Or blnExec = False Then
'                Control.Visible = False: Control.Enabled = False
'            Else
'                Control.Visible = True
'                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or blWaiteDispose = True Or strExeState = "初审" Then
'                    Control.Enabled = False
'                Else
'                    Control.Enabled = True
'                End If
'            End If
        Case conMenu_Tool_MedRec                                                            '指量删除无主
            Control.Enabled = (mintEditState = 0 And blnIF = True And blnExec = True And bln删除无主标本 = True)
            
        Case conMenu_Manage_Reset                                                           '批量修改样本号
            Control.Enabled = (mintEditState = 0 And blnIF = True Or blnExec = True)
        Case conMenu_Edit_QCRes                                                             '查看本月质控
            Control.Enabled = (intSampleType = 3 And mintEditState = 0)
            
        Case comMenu_LIS_TodayQC                                                            '今日质控
            Control.Visible = (mTodayQCPrivs <> "")
            
        Case comMenu_LIS_History                                                            '历史质控
            Control.Visible = (mHistoryPrivs <> "")
            
        Case conMenu_Tool_Analyse                                                           '置为比对
            If lngRowCount = 0 Or mintEditState > 0 Or intSampleType <> -1 Or blWaiteDispose = True Or TabList.Item(0).Visible = False _
               Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                '11198 置为比对 菜单在手工标本号下应置为无效
                Control.Enabled = InStr(Me.rptList.FocusedRow.Record(mCol.标本号).Caption, "-") <= 0
            End If
        Case conMenu_Edit_DeleteParent                                                      '置为无主
            If lngRowCount = 0 Or mintEditState > 0 Or intSampleType <> -1 _
               Or strPatienName = "" Or strExeState = "已检验" Or blWaiteDispose = True _
               Or blnIF = False Or blnExec = False Or strExeState = "初审" Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Edit_SendBack                                                          '状态回滚
            If InStr(1, mstrPrivs, "核收撤消") <= 0 Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
'                If lngRowCount = 0 Or mintEditState > 0 Or (strSource = "无主" And intSampleType < 3) Or _
'                    (strExeState = "已检验" And intSampleType <> 4) Then
'                    Control.Enabled = False
'                Else
'                    Control.Enabled = True
'                End If
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or blWaiteDispose = True Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_Manage_ClearUp                                                            '删除样本
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or strExeState = "初审" _
                Or strSource = "住院" Or strSource = "门诊" And strSource <> "无主" _
                Or intSampleType <> -1 Or InStr(1, mstrPrivs, "无主处理") <= 0 Or blWaiteDispose = True _
                Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If

        '''''''''''''''''''''''''''''''''''''''''''''''''报告'''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Manage_Report                                                          '报告填写
            If InStr(1, mstrPrivs, "报告填写") <= 0 Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or blWaiteDispose = True Or strExeState = "初审" Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_Edit_Adjust, conMenu_Edit_Dilute                                          '批量调整,稀释倍数
            If lngRowCount = 0 Or mintEditState > 0 Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        
        Case conMenu_Manage_Audit                                                              '批量审核
            If InStr(1, mstrPrivs, "审核标本") <= 0 Or mintEditState > 0 Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Edit_Audit                                                             '报告审核
            If InStr(1, mstrPrivs, "审核标本") <= 0 Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" _
                    Or strSource = "无主" Or intSampleType = 3 Or blWaiteDispose = True Or _
                    (mSendReport = 1 And str初审人 = "") Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_LIS_SendReport                                                         '初审报告
            If strPatienName = "" Or mSendReport = 0 Or mintEditState > 0 Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                Control.Enabled = (str初审人 = "" And strPatienName <> "" And strExeState = "检验中")
            End If
        Case conMenu_Edit_ClearUp                                                           '取消审核
            If InStr(1, mstrPrivs, "审核取消") <= 0 And InStr(1, mstrPrivs, "24小时审核取消") <= 0 _
                Or blnIF = False Or blnExec = False Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "检验中" Or blWaiteDispose = True Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_Manage_Redo                                                            '重做结果
            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" Or strExeState = "初审" _
                Or blMicrobe = True Or strSource = "无主" Or intSampleType = 3 _
                Or InStr(1, mstrPrivs, "无主处理") <= 0 Or blWaiteDispose = True Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Manage_Undone                                       '取消结果

            If lngRowCount = 0 Or mintEditState > 0 Or intReportCount = 0 Or blWaiteDispose = True _
                        Or blnIF = False Or blnExec = False Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_Edit_Import, conMenu_Edit_ApplyTo                                      '自动导入,批量导入
            Control.Enabled = mintEditState = 0
        Case conMenu_Edit_Insert                                                            '合并
            If Me.TabCtlWindow.Selected.Index = 5 Then
                Control.Visible = True
                Control.Enabled = (mintEditState = 0 And blnIF = True And blnExec = True)
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_Surplus                                                           '糖耐量合并
            Control.Enabled = (mintEditState = 0 And blnIF = True And strExeState <> "已检验" And blnExec = True And strExeState <> "初审")
        Case conMenu_LIS_SignVerify                                                         '验证签名
            If gobjESign Is Nothing Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Enabled = (strExeState = "已检验")
            End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''''费用''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Case conMenu_Edit_Price                                                             '生成主费
'            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" _
'               Or strSource = "无主" Or intSampleType = 3 Or Me.TabCtlWindow.Selected.Index <> 3 _
'                    Or InStr(1, mstrPrivs, "生成主费") <= 0 Or blWaiteDispose = True Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = True
'            End If
'        Case conMenu_Manage_ThingAdd                                                        '附加费划价
'            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" _
'               Or strSource = "无主" Or intSampleType = 3 Or Me.TabCtlWindow.Selected.Index <> 3 _
'                Or InStr(1, mstrPrivs, "附费处理") <= 0 Or blWaiteDispose = True Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = True
'            End If
'        Case conMenu_Edit_ModifyParent, conMenu_Edit_NewItem                                '附加费记帐,零费记录
'            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" _
'               Or strSource = "无主" Or intSampleType = 3 Or Me.TabCtlWindow.Selected.Index <> 3 _
'                    Or InStr(1, mstrPrivs, "附费处理") <= 0 Or blWaiteDispose = True Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = True
'            End If
'        Case conMenu_Manage_ThingModi, conMenu_Manage_ThingDel                              '修改附加费,删除附加费
'            If lngRowCount = 0 Or mintEditState > 0 Or strExeState = "已检验" _
'               Or strSource = "无主" Or intSampleType = 3 Or Me.TabCtlWindow.Selected.Index <> 3 _
'                Or InStr(1, mstrPrivs, "附费处理") <= 0 Or blWaiteDispose = True Then
'                Control.Enabled = False
'            Else
'                Control.Enabled = True
'            End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''查看''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_Backward                                                         '前一条
            If mintEditState = 4 Or mintEditState = 5 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
                If Me.rptList.Rows.Count <= 1 Then
                    Control.Enabled = False
                Else
                    If Not rptList.FocusedRow Is Nothing Then
                        If Me.rptList.FocusedRow.Index = 0 Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                End If
            End If
        Case conMenu_View_Forward                                                          '后一条
            
            If mintEditState = 4 Or mintEditState = 5 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
                If Me.rptList.Rows.Count <= 1 Then
                    Control.Enabled = False
                Else
                    If Not rptList.FocusedRow Is Nothing Then
                        If Me.rptList.Rows.Count - 1 = Me.rptList.FocusedRow.Index Then
                            Control.Enabled = False
                        Else
                            Control.Enabled = True
                        End If
                    End If
                End If
            End If
        Case conMenu_Tool_Reference_1, conMenu_Tool_Reference_2, conMenu_Tool_MeetFinish, conMenu_Tool_MeetCancel
            Control.Visible = False
        Case conMenu_View_Filter                                                           '过滤
            If InStr(1, mstrPrivs, "综合查询") <= 0 Then
                Control.Visible = False: Control.Enabled = False
            Else
                Control.Visible = True
                If mintEditState > 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            End If
        Case conMenu_View_Refresh                                                           '刷新
            If mintEditState > 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_LIS_PatientInfo                                                        '病人信息
            If lngRowCount = 0 Or mintEditState > 0 Or strSource = "无主" _
                Or intSampleType = 3 Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        Case conMenu_View_FindNext                                                          '病人历次检验
            If Not Me.rptList.FocusedRow Is Nothing Then
                Control.Enabled = mintEditState = 0
            Else
                Control.Enabled = False
            End If
        Case conMenu_Manage_Bespeak                                                         '只显示收费
            Control.Checked = Control.Checked
        Case conMenu_View_ToolBar_Button                                                    '显示工具条
            Control.Checked = Me.cbrthis(2).Visible
        Case conMenu_View_ToolBar_Text                                                      '是否显示文字
            Control.Checked = Not (Me.cbrthis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size                                                      '是否显示大图标
            Control.Checked = Me.cbrthis.Options.LargeIcons
        Case conMenu_View_StatusBar                                                         '是否显示状态栏
            Control.Checked = Me.stbThis.Visible
        Case conMenu_Manage_ReGet                                                           '显示待核收
            If mintEditState > 0 Then
                Control.Enabled = False
            Else
                Control.Checked = Control.Checked
                Me.TabList.Item(1).Visible = Control.Checked
                If Control.Checked = False Then Me.TabList.Item(0).Selected = True
            End If
'        case
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Report_DrugQuery, conMenu_Report_Reports, conMenu_Report_WorkLog       '科室,仪器
            If mintEditState <> 0 Then
                Me.cboDept.Enabled = False
                Me.cboMachine.Enabled = False
                Me.cboUnionItem.Enabled = False
            Else
                Me.cboDept.Enabled = True
                Me.cboMachine.Enabled = True
                Me.cboUnionItem.Enabled = True
            End If
        Case Else
            On Error Resume Next
            Select Case Me.TabCtlWindow.Selected.Index
                Case 4 '费用
                    mclsExpenses.zlUpdateCommandBars Control
                Case 6 '门诊医嘱
                    mclsOutAdvices.zlUpdateCommandBars Control
                Case 7 '住院医嘱
                    mclsInAdvices.zlUpdateCommandBars Control
            End Select
    End Select
    
    
    
    
    On Error Resume Next
    '如果当前选择的窗体不是应该回到焦点的窗体就指定
    Select Case gintSelectFocus
        Case 1              '列表
'            Me.dkpMain.FindPane(Dkp_ID_List).Select
'            Me.TabList.SetFocus
'            If Me.TabList.Selected.Index = 0 Then
'                Me.rptList.SetFocus
'                '由于焦点不正确需要下面动作来修正
'                SendKeys "{UP}"
''                SendKeys "{Down}"
'            Else
'                Me.rptList1.SetFocus
'            End If
        Case 2              '病人信息
            Me.dkpMain.FindPane(Dkp_ID_Request).Select
            mfrmRequest.Show
        Case 3              '报告填写
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            TabCtlWindow.SetFocus: mfrmWrite.Vsf.SetFocus
        Case 4
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            TabCtlWindow.SetFocus: mfrmWrite2.Vsf.SetFocus
        Case 5
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            TabCtlWindow.SetFocus: mfrmWrite2.vsfDetail.SetFocus
    End Select
    gintSelectFocus = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub chkSoure_Click(Index As Integer)
    Dim astrItem() As String
    Dim strTypeName As String
    If Me.Visible = False Then Exit Sub
    If Me.TabList.Selected.Index = 0 Then
        astrItem = Split(con_主界面筛选_检验中, ";")
        strTypeName = "检验中"
    Else
        astrItem = Split(con_主界面筛选_待核收, ";")
        strTypeName = "待核收"
    End If
    If strTypeName = "待核收" Then
        If Index = 5 Then
            zlDatabase.SetPara strTypeName & "_" & astrItem(Index - 3), chkSoure(Index).Value, 100, 1208
        Else
            zlDatabase.SetPara strTypeName & "_" & astrItem(Index), chkSoure(Index).Value, 100, 1208
        End If
    Else
        zlDatabase.SetPara strTypeName & "_" & astrItem(Index), chkSoure(Index).Value, 100, 1208
    End If
    If Me.TabList.Selected.Index = 0 Then
        Call GetVerifying
    Else
        Call GetWaitVerify
    End If
    '过滤界面列表
    RptListFilter
End Sub

Private Sub chkSoure_GotFocus(Index As Integer)
    On Error Resume Next
    If Me.TabList.Selected.Index = 0 Then
'        Me.rptList.SetFocus
    Else
'        Me.rptList1.SetFocus
    End If
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Me.Visible = False Then Exit Sub
    Cancel = True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Me.Visible = False Then Exit Sub
    Select Case Item.ID
    Case Dkp_ID_List
        Item.Handle = Me.PicList.hWnd
    Case Dkp_ID_Locate
        Item.Handle = Me.PicInfo.hWnd
    Case Dkp_ID_Request
        Item.Handle = mfrmRequest.hWnd
    Case Dkp_ID_Append
        Item.Handle = Me.picTab.hWnd
    Case Dkp_ID_Image
        Item.Handle = Me.PicImage.hWnd
    End Select
End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    If Me.Visible = False Then Exit Sub
    Me.cbrthis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    Top = lngTop
    Bottom = Me.ScaleHeight - lngBottom
End Sub

Private Sub dkpMain_Resize()
    If Me.Visible = False Then Exit Sub
    Me.cbrthis.RecalcLayout
    ImageTypeSet Me.VScroll.Max
End Sub

Private Sub dtpDate_Change()
    If Me.TabList.Item(1).Selected = True Then
        zlDatabase.SetPara "待核收范围", cbo时间.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Call RefreshData1
    Else
        zlDatabase.SetPara "标本范围", cbo时间.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Call RefreshData
    End If
    
End Sub

Private Sub dtpDateEnd_Change()
    If Me.TabList.Item(1).Selected = True Then
        zlDatabase.SetPara "待核收范围", cbo时间.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Call RefreshData1
    Else
        zlDatabase.SetPara "标本范围", cbo时间.Text & ";" & Me.dtpDate & ";" & Me.dtpDateEnd, 100, 1208
        Call RefreshData
    End If
End Sub

Private Sub Form_Activate()
    '收到检验数据后，触发此事件
    On Error Resume Next
'    If objLISComm.DataReceived And Me.Tag <> "Refresh" And Not blnChecking And blnAutoRefresh And mintEditState = 0 Then
'        Me.Tag = "Refresh"
'        Call RefreshData
'        Me.Tag = ""
'    End If
    
    If mintLoadShow = 0 Then
        '=====================================================
        '由于最后删除图像可能会现问题，先放到打开时来测试
        Call DeleteTmpFile
        '=====================================================
        
        mstrPrivs = gstrPrivs                                       '初使化权限
        
        Call zlDatabase.ShowReportMenu(Me.cbrthis, glngSys, glngModul, mstrPrivs)
        
        LoadAllData
        
        mintLoadShow = mintLoadShow + 1
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If mintEditState > 0 Then Exit Sub
'    If KeyCode = 38 Then
'        BackOrNextPatient 1
'    ElseIf KeyCode = 40 Then
'        BackOrNextPatient 2
'    End If
End Sub

Private Sub Form_Load()
    Dim strPrivs As String
'    Set objLISComm = CreateObject("Zl9LISComm.clsPublic")
'    '启动仪器数据接收初始化
'    If objLISComm Is Nothing Then
'        MsgBox "通讯程序初始化失败!", vbExclamation, gstrSysName
'
'        Unload Me
'    End If
'    objLISComm.InitLISComm gcnOracle, Me
    '--------------------------------------------
    '插件相关
    Set clsPluginLoader = New PlugInLoader
    
    ' the interface the plugins have to implement
    ' 插件必须实现接口
    Set clsPluginLoader.Interface = New zl9LisQuery_Def.clsLisQuery
    '--------------------------------------------
    
    InitinterFace                                               '初始化界面
    
    Call RestoreWinState(Me, App.ProductName)                   '界面恢复
    
    
    '启动仪器数据接收初始化
'    objLISComm.InitLISComm gcnOracle, Me
End Sub

Private Sub DeleteTmpFile()
    '删除部件产生的临时文件
    On Error Resume Next
    If Dir(App.path & "\*.BMP") <> "" Then
        Kill App.path & "\*.BMP"
    End If
    If Dir(App.path & "\*.JPG") <> "" Then
        Kill App.path & "\*.JPG"
    End If
    If Dir(App.path & "\*.GIF") <> "" Then
        Kill App.path & "\*.GIF"
    End If
    If Dir(App.path & "\*.CHT") <> "" Then
        Kill App.path & "\*.CHT"
    End If
    If Dir(App.path & "\*.ZIP") <> "" Then
        Kill App.path & "\*.ZIP"
    End If
    If gobjFSO.FolderExists(App.path & "\ZLLIS_ZIP") Then gobjFSO.DeleteFolder App.path & "\ZLLIS_ZIP", True
End Sub

Private Sub CreateDockPane()
    Dim Pane1 As Pane, Pane2 As Pane, Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Dim lngPane5Width As Long, lngPane2Height As Long, lngPane2Width As Long, lngPane3Height As Long
    
    
    dkpMain.Options.HideClient = True
    
    Set Pane1 = dkpMain.CreatePane(Dkp_ID_List, 200, 150, DockLeftOf, Nothing)
    Pane1.Title = "样本清单"
    Pane1.Handle = Me.PicList.hWnd
'    Pane1.Options = PaneNoCaption

    Set Pane2 = dkpMain.CreatePane(Dkp_ID_Locate, 200, 600, DockRightOf, Nothing)
    Pane2.Title = "病人定位"
    Pane2.Handle = Me.PicInfo.hWnd
'    Pane2.Options = PaneNoCaption
    
    Set Pane3 = dkpMain.CreatePane(Dkp_ID_Request, 400, 600, DockBottomOf, Pane2)
    Pane3.Title = "核收登记"
    Pane3.Handle = mfrmRequest.hWnd
'    Pane3.Options = PaneNoCaption
    
    Set Pane4 = dkpMain.CreatePane(Dkp_ID_Append, 400, 790, DockRightOf, Pane3)
    Pane4.Title = "附加窗体"
    Pane4.Handle = Me.picTab.hWnd
'    Pane4.Options = PaneNoCaption
    
    lngPane5Width = zlDatabase.GetPara("图像宽度", 100, 1208, 200)
    Set Pane5 = dkpMain.CreatePane(Dkp_ID_Image, lngPane5Width, 200, DockRightOf, Pane4)
    Pane5.Title = "图像显示"
    Pane5.Handle = Me.PicImage.hWnd
'    Pane5.Options = PaneNoCaption
    
    Call ShowRequest(False)
    
    Pane1.Select
    
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane
    Dim intLoop As Integer
    On Error Resume Next
    
    If Me.Visible = False Then Exit Sub
    If Me.WindowState = 1 Then Exit Sub

    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Locate)
    Pane1.MinTrackSize.SetSize 6954 / Screen.TwipsPerPixelX, 380 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize Pane1.MaxTrackSize.Width, 380 / Screen.TwipsPerPixelY
    
    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Request)
    Pane1.MinTrackSize.SetSize 3480 / Screen.TwipsPerPixelX, 2295 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize 3480 / Screen.TwipsPerPixelX, 2295 / Screen.TwipsPerPixelY
    
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
    
'    Pane1.MinTrackSize.SetSize 0, 2295 / Screen.TwipsPerPixelY
'    Pane1.MaxTrackSize.SetSize Screen.Width, 2295 / Screen.TwipsPerPixelY
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngloop As Long
    Dim frmThis As Form

    If mintEditState <> 0 Then
        If MsgBox("您正在编辑报告,是否确定要退出？", vbYesNo + vbQuestion + vbDefaultButton2, Me.Caption) = vbYes Then
            Call SaveDisposal(mFileS.放弃)
        Else
            Cancel = True
            Exit Sub
        End If
    End If
    
    Call SaveWinState(Me, App.ProductName)
    Me.Visible = False
    mstrAuditingManID = ""
    
    zlDatabase.SetPara "缺省科室ID", mlngDeptID, 100, 1208
    zlDatabase.SetPara "过滤仪器", mlngMachineID, 100, 1208
    zlDatabase.SetPara "仪器小组", mstrMachineGroup, 100, 1208
    
    '界面过滤列表
    zlDatabase.SetPara "显示待核收", Me.cbrthis.FindControl(, conMenu_Manage_ReGet, True, True).Checked, 100, 1208
    zlDatabase.SetPara "隐藏检验图形", Me.cbrthis.FindControl(, conMenu_Manage_LeaveMedi, True, True).Checked, 100, 1208
    '图像显示的大小
    zlDatabase.SetPara "图像宽度", Me.PicImage.Width / Screen.TwipsPerPixelX, 100, 1208
    '将文件提取时间默认设为当天，供下次使用
    Call zlDatabase.SetPara("文件提取范围", 0, 100, 1208)
    
    
    '保存当前Dkp的风格,保存到数据太长了还是保存在注册表中
'    zlDatabase.SetPara "DKP保存", dkpMain.SaveStateToString, 100, 1208
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
    
    With Me.rptList
        For lngloop = 0 To Me.rptList.SortOrder.Count - 1
            If .SortOrder(lngloop).Caption = "标本号" Then
                zlDatabase.SetPara "标本号排序", .SortOrder(lngloop).SortAscending, 100, 1208
            End If
        Next
    End With
    
    mstrAuditingMan = ""
    mintAuditing = 0
    mintLoadShow = 0
    
    mblnTabList1 = False
    
    '--------------------------------------------------------
    '释放插件
    Dim i As Long
    '释放调用的DLL
    If Not clsPluginLoader Is Nothing Then
        If clsPluginLoader.PluginCount > 0 Then
            For i = 0 To clsPluginLoader.PluginCount - 1
                Call clsPluginLoader.ClosePlugin(i)
            Next
        End If
        Set clsPluginLoader = Nothing
    End If
        
    For i = LBound(mobjPlugin) To UBound(mobjPlugin)
        Set mobjPlugin(i) = Nothing
    Next
    

    '强行Unload,不然不会激活子窗体的事件
'    If mcolSubForm Is Nothing Then
'        For lngLoop = 1 To mcolSubForm.Count
'            Unload mcolSubForm(lngLoop)
'        Next
'    End If
    Set mcolSubForm = Nothing
    
    Me.rptList.Records.DeleteAll
    Me.rptList.Populate
    Me.rptList1.Records.DeleteAll
    Me.rptList1.Populate
    

    Set mclsExpenses = Nothing
    Set mclsInAdvices = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsInEPRs = Nothing
'    Set mclsOutEPRs = Nothing
    Set mclsEMR = Nothing


    Unload mfrmRequest
    Unload mfrmWrite
    Unload mfrmWrite2
    Unload mfrmTrack
    Unload mfrmLabMainSampleUnion
    Unload mfrmLabMicrobe3Report
    
    If Not gobjEmr Is Nothing Then
        Call gobjEmr.CloseForms
    End If

    Set mfrmRequest = Nothing
    Set mfrmWrite = Nothing
    Set mfrmWrite2 = Nothing
    Set mfrmTrack = Nothing
    Set mfrmLabMicrobe3Report = Nothing
    Set mfrmLabMainSampleUnion = Nothing
    
    Me.TabCtlWindow.RemoveAll
    Me.TabList.RemoveAll
    Me.cbrChild.DeleteAll
    Me.cbrthis.DeleteAll
    Me.dkpMain.DestroyAll
    
    '=====================================================
    '最后删除图像文件
    Call DeleteTmpFile
    '=====================================================
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lbl筛选_Click()
    Call picFilter_Click
End Sub

Private Sub mfrmLabMicrobe3Report_StartEdit(Cancel As Boolean)
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    On Error GoTo errH:
    If InStr(",7,8,13,", CStr(Me.rptList.FocusedRow.Record(mCol.执行状态).Icon)) > 0 Then
        '已检验
        Cancel = True
        mintHandleState = 0
    Else
        If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
            Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
            ReportDisposal mActR.填写三级报告
            Cancel = False
        Else
            Cancel = True
        End If
        
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mfrmRequest_ZlAutoSave(ByVal lngSampleID As Long)
    If lngSampleID = 0 Then Exit Sub
    
    On Error GoTo errH:
    If mintContinue = 0 Then
        '不连续操作
        Me.rptList.Tag = ""   '清空连续保存的标记
        mlngKey = lngSampleID
        If mlngMachineID > 0 Or mlngMachineID = -1 Then
            InsertOneRecored mlngKey, True
        Else
            Call RefreshData
        End If
        Call SaveDisposal(mFileS.放弃)
        '核收后是否发送仪器数据
        Call SampleDisposal(mActS.发送仪器)

    Else
        Select Case mintEditState
            Case 4
                RefreshData
                '核收后是否发送仪器数据
                Call SampleDisposal(mActS.发送仪器)
                If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                    If AuditionCheck = True Then
                        Call ReportDisposal(mActR.审核报告)
                    End If
                End If
                
                If MoveStation(1, 1) = False Then                       '向下移动
                    '没有找到记录时退出操作
                    mintHandleState = 0
                    mintEditState = 0
                    Call SaveDisposal(mFileS.放弃)
                Else
                    Call SampleDisposal(mActS.补填病人)
                End If
                
            Case Else
                If Me.rptList.Tag = "" Then
                    '第一次增加时先清除列表
                    Me.rptList.Records.DeleteAll
                    Me.rptList.Tag = "Continue"
                End If
                '添加刚新增的记录到列表中
                InsertOneRecored lngSampleID
                 '核收后是否发送仪器数据
                Call SampleDisposal(mActS.发送仪器)
                If mintEditState = 1 Then
                    Call SampleDisposal(mActS.核收)
                End If
                If mintEditState = 2 Then
                    Call SampleDisposal(mActS.登记)
                End If
        End Select
        
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mfrmWrite2_StartEdit(Cancel As Boolean)
    If Me.rptList.Rows.Count = 0 Then Exit Sub
    On Error GoTo errH
    If InStr(",7,8,13,", CStr(Me.rptList.FocusedRow.Record(mCol.执行状态).Icon)) > 0 Then
        '已检验
        Cancel = True
        mintHandleState = 0
    Else
        '还在进行登记核收补填时自动保存
        If mintEditState >= 1 And mintEditState <= 4 Then
            If Me.cbrthis.FindControl(, conMenu_Edit_Save, , True).Enabled = True And _
               Me.cbrthis.FindControl(, conMenu_Edit_Save, , True).Visible = True Then
                Call SaveDisposal(mFileS.保存)
            End If
        End If
        Cancel = False
'        mintHandleState = 2
        If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
            Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
            ReportDisposal mActR.填写报告
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picFilter_Click()
    Dim vRect As RECT
    If Me.picFilter.Tag = "" Then
        Me.picFilter.Tag = "True"
    Else
        Me.picFilter.Tag = ""
    End If
    If Me.TabList.Item(0).Selected = True Then
        vRect = GetControlRect(Me.picFilter.hWnd)
        frmLabMainSizer.ShowMe Me, "检验中", IIf(Me.picFilter.Tag = "", True, False)
        frmLabMainSizer.Left = vRect.Left - 400
        frmLabMainSizer.Top = vRect.Top + 350
        Call GetVerifying
    Else
        vRect = GetControlRect(Me.picFilter.hWnd)
        frmLabMainSizer.ShowMe Me, "待核收", IIf(Me.picFilter.Tag = "", True, False)
        frmLabMainSizer.Left = vRect.Left - 400
        frmLabMainSizer.Top = vRect.Top + 350

    End If
    Call RptListFilter
End Sub

Private Sub picFilter_LostFocus()
    If Me.TabList.Item(0).Selected = True Then
        frmLabMainSizer.ShowMe Me, "检验中", True
        Call GetVerifying
    Else
        frmLabMainSizer.ShowMe Me, "待核收", True
    End If
    Call RptListFilter
End Sub

Private Sub picList_Click()
    If Me.TabList.Item(0).Selected = True Then
        frmLabMainSizer.ShowMe Me, "检验中", True
        Call GetVerifying
        If Me.picFilter.Tag = "True" Then Call RptListFilter
        Me.picFilter.Tag = ""
    Else
        frmLabMainSizer.ShowMe Me, "待核收", True
        Call GetWaitVerify
        If Me.picFilter.Tag = "True" Then Call RptListFilter
        Me.picFilter.Tag = ""
    End If
    
End Sub

Private Sub picList_GotFocus()
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
'    Me.cbo时间.SetFocus
    If Me.TabList.Tag = "" And Me.TabList.Selected.Index = 0 Then
'        Me.rptList.SetFocus
    Else
'        Me.rptList1.SetFocus
    End If
    Me.TabList.Tag = "Show"
End Sub

Private Sub picList_LostFocus()
'    Me.TabList.Tag = ""
End Sub

Private Sub picList_Resize()
    On Error Resume Next
'    Me.rptList.Top = 0
    Me.TabList.Left = 0
    Me.TabList.Width = PicList.ScaleWidth
    Me.TabList.Height = PicList.ScaleHeight - Me.TabList.Top
    If Me.TabList.Selected.Index = 0 Then
        Me.picFilter.Left = Me.chkSoure(5).Left + Me.chkSoure(5).Width + 30
    Else
        Me.picFilter.Left = Me.chkSoure(2).Left + Me.chkSoure(2).Width + 30
    End If
    Me.cbo时间.Top = Me.TabList.Top + Me.TabList.Height - Me.cbo时间.Height
    Me.cbo时间.Left = 2300
    dtpDate.Top = Me.cbo时间.Top
    dtpDate.Left = Me.cbo时间.Left + Me.cbo时间.Width + 10
    dtpDateEnd.Top = Me.cbo时间.Top
    dtpDateEnd.Left = Me.dtpDate.Left + Me.dtpDate.Width + 10
End Sub

Private Sub picTab_Resize()
    Me.TabCtlWindow.Top = 0
    Me.TabCtlWindow.Left = 0
    Me.TabCtlWindow.Width = Me.picTab.ScaleWidth
    Me.TabCtlWindow.Height = Me.picTab.ScaleHeight
End Sub
Private Sub CreateTableControl()
    Dim Item As TabControlItem
    'Dim ObjchargeWindow As Object
    Dim strPrivs As String
    
    On Error Resume Next
    
    With Me.TabList
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.COLOR = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 0, "检验中", rptList.hWnd, conMenu_Tool_Report
        .InsertItem 1, "待核收", rptList1.hWnd, conMenu_Tool_Report
        .PaintManager.Position = xtpTabPositionBottom
        .PaintManager.LayOut = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        If zlDatabase.GetPara("显示待核收", 100, 1208, "False") = "True" Then
            .Item(1).Visible = True
        Else
            .Item(1).Visible = False
        End If
        .Item(0).Selected = True
    End With
    
    
    With Me.TabCtlWindow
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.COLOR = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem(0, "报告结果", mfrmWrite.hWnd, conMenu_Tool_Report).Tag = "普通报告结果"
        .InsertItem(1, "报告结果", mfrmWrite2.hWnd, conMenu_Tool_Report).Tag = "微生物报告结果"
        .InsertItem(2, "三级报告", mfrmLabMicrobe3Report.hWnd, conMenu_Tool_Report).Tag = "微生物三级报告"
        .InsertItem(3, "历次对比", mfrmTrack.hWnd, conMenu_Edit_Audit).Tag = "历次对比"
        'Set ObjchargeWindow = mclsExpenses.zlGetForm
        strPrivs = GetPrivFunc(glngSys, p医嘱附费管理)  '没有医嘱附费管理权限时不显示
        .InsertItem(4, "费用查询", PicWindows.hWnd, conMenu_Edit_Price).Tag = IIf(strPrivs <> "", "费用查询", "")
        .Item(4).Visible = IIf(strPrivs <> "", True, False)
        .InsertItem(5, "标本合并", mfrmLabMainSampleUnion.hWnd, conMenu_Edit_Archive).Tag = "标本合并"
         strPrivs = GetPrivFunc(glngSys, p门诊医嘱下达)
        .InsertItem(6, "门诊医嘱", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "门诊医嘱", "")
        .Item(6).Visible = IIf(strPrivs <> "", True, False)
        strPrivs = GetPrivFunc(glngSys, p住院医嘱下达)
        .InsertItem(7, "住院医嘱", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "住院医嘱", "")
        .Item(7).Visible = IIf(strPrivs <> "", True, False)
        strPrivs = GetPrivFunc(glngSys, p门诊病历管理)
        .InsertItem(8, "门诊病历", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "门诊病历", "")
        strPrivs = GetPrivFunc(glngSys, p住院病历管理)
        .InsertItem(9, "住院病历", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "住院病历", "")
        strPrivs = GetPrivFunc(glngSys, p新版病历管理)
        .InsertItem(10, "电子病历", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "电子病历", "")
        
        .PaintManager.LayOut = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
        .Item(1).Visible = False
        .Item(6).Visible = False
        .Item(7).Visible = False
        .Item(8).Visible = False
        .Item(9).Visible = False
        .Item(10).Visible = False
    End With
    
'    If Me.TabList.Item(0).Selected = True Then
'        cbo时间.Text = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0)
'        Me.DTPDate.Value = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";" & Format(Now, "yyyy-mm-dd"), ";")(1)
'        Me.dtpDateEnd.Value = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";" & Format(Now, "yyyy-mm-dd") & ";" & Format(Now, "yyyy-mm-dd"), ";")(2)
'    Else
'        cbo时间.Text = zlDatabase.GetPara("待核收范围", 100, 1208, "今  天")
'    End If
    cbo时间.Text = "今  天"
    Me.dtpDate.Visible = (Me.cbo时间.Text = "自定义")
    Me.dtpDateEnd.Visible = (Me.cbo时间.Text = "自定义")
End Sub
Private Function LoadInterFaceCbo() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim lngTmp As Long
    Dim ControlcboDept As CommandBarComboBox
    Dim ControlcboMachine As CommandBarComboBox
    Dim strSQL As String
    
    On Error GoTo errH
    mlngDeptID = zlDatabase.GetPara("缺省科室ID", 100, 1208, mlngDeptID)
    mlngMachineID = zlDatabase.GetPara("过滤仪器", 100, 1208, mlngMachineID)
    
    '2.读取部门数据
    If InStr(mstrPrivs, "所有科室") > 0 Then
        strSQL = "SELECT A.编码||'-'||A.名称 as 名称,A.ID FROM 部门表 A,部门性质说明 B WHERE " & _
                  " (A.撤档时间 IS NULL OR A.撤档时间 =TO_DATE('3000-01-01','YYYY-MM-DD')) AND " & _
                  " A.ID=B.部门ID AND B.工作性质='检验' ORDER BY A.编码||'-'||A.名称"
    Else
        strSQL = "Select A.编码 || '-' || A.名称 As 名称, A.ID" & vbNewLine & _
                "  From 部门表 A, 部门性质说明 B" & vbNewLine & _
                "  Where (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And A.ID = B.部门id And B.工作性质 = '检验' And" & vbNewLine & _
                "        A.ID In (Select Distinct D.使用小组id" & vbNewLine & _
                "                 From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
                "                 Where A.小组id = B.ID And B.ID = C.小组id　and C.仪器id = D.ID And 人员id = [1] and C.查看 = 1)" & vbNewLine & _
                "  Order By A.编码 || '-' || A.名称"
    End If
    
    cboDept.Clear
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    'If InStr(mstrPrivs, "所有科室") > 0 Then
    cboDept.AddItem "所有科室"
    
    Do Until rsTmp.EOF
        cboDept.AddItem rsTmp("名称")
        cboDept.ItemData(cboDept.NewIndex) = rsTmp("ID")
        If rsTmp("id") = IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID) Then
            cboDept.ListIndex = cboDept.NewIndex
            mlngDeptID = IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID)
'            objLISComm.DeptID = mlngDeptID
        End If
        rsTmp.MoveNext
    Loop
    
    If cboDept.ListCount > 0 And Trim(cboDept.Text) = "" Then
        cboDept.ListIndex = 0
        mlngDeptID = cboDept.ItemData(0)
'        objLISComm.DeptID = mlngDeptID
    End If
    
    
    gstrSql = "select 部门ID from 部门人员 where 人员id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, UserInfo.ID)
    Do Until rsTmp.EOF
        mUserDept = mUserDept & ";" & Nvl(rsTmp("部门Id"))
        rsTmp.MoveNext
    Loop
    If mUserDept <> "" Then mUserDept = mUserDept & ";"
    Me.MousePointer = 0
'    If cboDept.ListCount > 0 Then
'
'        cboMachine.Clear
'        cboMachine.AddItem "<所有仪器>": cboMachine.ItemData(cboMachine.NewIndex) = 0
'        cboMachine.AddItem "<手工>": cboMachine.ItemData(cboMachine.NewIndex) = -1
'        strsql = "SELECT a.名称,a.ID FROM 检验仪器 a where 使用小组ID = [1]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlngDeptID)
'        Do Until rsTmp.EOF
'            cboMachine.AddItem rsTmp("名称")
'            cboMachine.ItemData(cboMachine.NewIndex) = rsTmp("Id")
'            If rsTmp("id") = mlngMachineID Then
'                cboMachine.ListIndex = cboMachine.NewIndex
'            End If
'            rsTmp.MoveNext
'        Loop
'        If cboMachine.ListCount > 0 And Trim(cboMachine.Text) = "" Then
'            cboMachine.ListIndex = 0
'            mlngMachineID = cboMachine.ItemData(0)
'        End If
'    End If
    Exit Function
errH:
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitinterFace()
    '界面初始化
    On Error GoTo errH
    
    LoadRegistSetup                     '读入注册表保存设置并初始化变量及窗体
    CreateTableControl                  '创建TAB
    CreateCbs                           '创建工具条
    CreateChildCbs
    CreateDockPane                      '创建浮动窗体
    CreaterptListHead                   '建立列表头
    
    On Error Resume Next
    With Me.WinsockC                    '初始化和接收程序的通讯接口
        .Protocol = sckUDPProtocol
        .RemoteHost = "Localhost"
        .RemotePort = 1000
        .Bind 1001
    End With
    Exit Sub
errH:

    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadAllData()
    
    On Error GoTo errH
    
    '读入数据
    Call GetVerifying                   '读入检验中过滤数据
    Call GetWaitVerify                  '读入待核收过滤数据
    LoadInterFaceCbo                    '读入仪器和科室
    RefreshData                         '刷新
    RptListFilter                       '界面列表刷新
    rptList_SelectionChanged            '刷新状态
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




Private Function GetQuerySQL(ByVal strCondition As String, Optional ByVal bytMode As Byte = 1) As String
    '--------------------------------------------------------------------------------------------------------
    '功能：
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim mlngLoop As Long
     
    On Error Resume Next
    '以下是根据设置条件构成的条件语句
    
    If strCondition = "" Then Exit Function
    
    varTmp = Split(strCondition, "^")
'    If bytMode = 1 Then
'        If Val(varTmp(0)) > 0 Then GetQuerySQL = GetQuerySQL & " AND A.执行科室ID + 0 = " & Val(varTmp(0))
'    Else
'        If Val(varTmp(0)) > 0 Then GetQuerySQL = GetQuerySQL & " AND c.执行科室ID + 0 = " & Val(varTmp(0))
'    End If
    If Val(varTmp(0)) > 0 Then GetQuerySQL = GetQuerySQL & " AND c.执行科室ID + 0 = " & Val(varTmp(0))
    
    If Val(varTmp(1)) > 0 Then GetQuerySQL = GetQuerySQL & " AND c.仪器ID = " & Val(varTmp(1))
    If Trim(varTmp(2)) <> "所  有" Then
        Select Case Trim(varTmp(2))
        Case "指  定"
            GetQuerySQL = GetQuerySQL & " AND c.检验时间 BETWEEN TO_DATE('" & Format(varTmp(3), "yyyy-mm-dd hh:mm") & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(4), "yyyy-mm-dd hh:mm") & "', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            GetQuerySQL = GetQuerySQL & " AND c.检验时间 BETWEEN TO_DATE('" & GetDateTime(varTmp(2), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(2), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
    End If
    varTmp2 = Split(Trim(varTmp(5)), ",")
    strTmp = ""
    For mlngLoop = 0 To UBound(varTmp2)
        If InStr(varTmp2(mlngLoop), "～") = 0 Then
            strTmp = strTmp & "  OR c.标本序号=" & TransSampleNO(varTmp2(mlngLoop))
        Else
            strTmp = strTmp & "  OR c.标本序号 BETWEEN " & TransSampleNO(Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "～") - 1)) & " AND " & TransSampleNO(Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "～") + 1))
        End If
    Next
    If strTmp <> "" Then GetQuerySQL = GetQuerySQL & " AND (1=2 " & strTmp & ")"

    If Trim(varTmp(6)) <> "" Then GetQuerySQL = GetQuerySQL & " AND c.检验人='" & Trim(varTmp(6)) & "'"
    If Trim(varTmp(7)) <> "" Then GetQuerySQL = GetQuerySQL & " AND c.审核人='" & Trim(varTmp(7)) & "'"
    
    If Trim(varTmp(8)) <> "所  有" Then
        
        Select Case Trim(varTmp(8))
        Case "指  定"
            GetQuerySQL = GetQuerySQL & " AND c.审核时间 BETWEEN TO_DATE('" & Format(varTmp(9), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(10), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            GetQuerySQL = GetQuerySQL & " AND c.审核时间 BETWEEN TO_DATE('" & GetDateTime(varTmp(8), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(8), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
        
    End If
    
    If Val(varTmp(11)) > 0 Then
'        If bytMode = 1 Then
'            GetQuerySQL = GetQuerySQL & " AND F.执行状态 = " & IIf(Val(varTmp(11)) = 1, "3", "1")
'        Else
'            GetQuerySQL = GetQuerySQL & " AND c.样本状态 = " & IIf(Val(varTmp(11)) = 1, "1", "2")
'        End If
        GetQuerySQL = GetQuerySQL & " AND c.样本状态 = " & IIf(Val(varTmp(11)) = 1, "1", "2")
    End If
    
    If Val(varTmp(12)) > 0 Then
        GetQuerySQL = GetQuerySQL & " AND c.ID IN (SELECT G.检验标本ID FROM 检验普通结果 G,检验项目 H WHERE H.诊治项目id=G.检验项目id AND G.检验标本ID=c.ID "
        GetQuerySQL = GetQuerySQL & " AND G.检验项目ID=" & Val(varTmp(12))
        
        If Val(varTmp(13)) = 1 Then
            GetQuerySQL = GetQuerySQL & " AND H.结果类型=1 AND DECODE(H.结果类型,1,TO_NUMBER(G.检验结果),0)"
            strTmp = Val(varTmp(16))
        Else
            GetQuerySQL = GetQuerySQL & " AND G.检验结果"
            strTmp = "'" & varTmp(16) & "'"
        End If
        
        Select Case varTmp(15)
        Case "大于"
            GetQuerySQL = GetQuerySQL & ">" & strTmp
        Case "小于"
            GetQuerySQL = GetQuerySQL & "<" & strTmp
        Case "大于等于"
            GetQuerySQL = GetQuerySQL & ">=" & strTmp
        Case "小于等于"
            GetQuerySQL = GetQuerySQL & "<=" & strTmp
        Case "不等于"
            GetQuerySQL = GetQuerySQL & "<>" & strTmp
        Case "包含"
            GetQuerySQL = GetQuerySQL & " LIKE '%" & varTmp(16) & "%'"
        Case "在范围内"
            If Val(varTmp(13)) = 1 Then
                GetQuerySQL = GetQuerySQL & " BETWEEN " & strTmp & " AND " & Val(varTmp(17))
            Else
                GetQuerySQL = GetQuerySQL & " BETWEEN " & strTmp & " AND '" & varTmp(17) & "'"
            End If
        Case Else
            GetQuerySQL = GetQuerySQL & "=" & strTmp
        End Select
        GetQuerySQL = GetQuerySQL & ")"
    End If
    
    If bytMode = 1 Then
        If Trim(varTmp(18)) <> "" Then GetQuerySQL = GetQuerySQL & " AND b.姓名 Like '" & Trim(varTmp(18)) & "%'"
        If Val(varTmp(19)) > 0 Then GetQuerySQL = GetQuerySQL & " AND A.病人科室ID = " & Val(varTmp(19))
        If Val(varTmp(20)) > 0 Then GetQuerySQL = GetQuerySQL & " AND b.住院号=" & varTmp(20)
        If Val(varTmp(21)) > 0 Then GetQuerySQL = GetQuerySQL & " AND b.当前床号=" & Val(varTmp(21))
        If Val(varTmp(22)) > 0 Then GetQuerySQL = GetQuerySQL & " AND b.门诊号=" & varTmp(22)
'        If Trim(varTmp(23)) <> "" Then GetQuerySQL = GetQuerySQL & " AND A.开嘱医生='" & Trim(varTmp(23)) & "'"
'        If Trim(varTmp(UBound(varTmp))) <> "" Then GetQuerySQL = GetQuerySQL & " AND A.开嘱医生='" & Trim(varTmp(UBound(varTmp))) & "'"
        If Val(varTmp(24)) > 0 Then GetQuerySQL = GetQuerySQL & " AND A.开嘱科室ID = " & Val(varTmp(24))
        
        
        If Trim(varTmp(25)) <> "所  有" Then
            Select Case Trim(varTmp(25))
            Case "指  定"
                GetQuerySQL = GetQuerySQL & " AND A.开嘱时间 BETWEEN TO_DATE('" & Format(varTmp(26), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(27), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
            Case Else
                GetQuerySQL = GetQuerySQL & " AND A.开嘱时间 BETWEEN TO_DATE('" & GetDateTime(varTmp(25), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(25), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
            End Select
        End If
    Else
        If Trim(varTmp(18)) <> "" Or Val(varTmp(19)) > 0 Or Val(varTmp(20)) > 0 Or _
                Val(varTmp(21)) > 0 Or _
                Val(varTmp(22)) > 0 Or _
                Trim(varTmp(23)) <> "" Or _
                Val(varTmp(24)) > 0 Or _
                Trim(varTmp(25)) <> "所  有" Then
                
            GetQuerySQL = GetQuerySQL & " AND 1=1 "
        End If
    End If
    
    If Trim(varTmp(28)) <> "" Then GetQuerySQL = GetQuerySQL & " AND c.采样人='" & Trim(varTmp(28)) & "'"
    
    If Trim(varTmp(29)) <> "所  有" Then
        Select Case Trim(varTmp(29))
        Case "指  定"
            GetQuerySQL = GetQuerySQL & " AND c.采样时间 BETWEEN TO_DATE('" & Format(varTmp(30), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(31), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            GetQuerySQL = GetQuerySQL & " AND c.采样时间 BETWEEN TO_DATE('" & GetDateTime(varTmp(29), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(29), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
    End If
    
    If Trim(varTmp(32)) <> "所有类型" And InStr(varTmp(32), "-") > 0 Then GetQuerySQL = GetQuerySQL & " AND c.标本类型='" & zlCommFun.GetNeedName(Trim(varTmp(32))) & "'"
    
    'If GetQuerySQL <> "" Then GetQuerySQL = " AND " & GetQuerySQL
    
End Function



Private Sub rptList_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    frmLabMainSizer.ShowMe Me, "检验中", True
    Call GetVerifying
    If Me.picFilter.Tag = "True" Then Call RptListFilter
    Me.picFilter.Tag = ""
'    Me.cbo时间.SetFocus
'    Me.rptList.SetFocus
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    On Error Resume Next
    If Button = 2 Then
        If rptList.Records.Count <= 0 Then Exit Sub
        If Not rptList.SelectedRows(0).GroupRow Then
            Set objPopup = Me.cbrthis.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "报告审核(&A)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "取消审核(&U)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "状态回滚(&Z)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "样本复查(&D)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消复查(&E)")
                Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "报告查询(&P)"): cbrControl.BeginGroup = True
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "单个发送到仪器(&S)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_LIS_TOQC, "置为质控(&Q)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "置为比对(&Y)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "查看比对(&B)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "糖耐量合并(&E)")
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改样本号和姓名(&M)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "删除样本(&D)")
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒收(&J)"): cbrControl.BeginGroup = True
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "放弃(&C)")
            End With
            objPopup.ShowPopup
        End If
    End If
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    On Error GoTo errH
    If Row.Record(mCol.标本类型).Icon = -1 And InStr(",7,8,13,", Row.Record(mCol.执行状态).Icon) = 0 Then
        If Me.TabCtlWindow.Item(5).Selected = False Then
            mintHandleState = 1
            If Me.cbrthis.FindControl(, conMenu_Manage_Receive, , True).Enabled = True And _
               Me.cbrthis.FindControl(, conMenu_Manage_Receive, , True).Visible = True Then
                Call SampleDisposal(mActS.补填病人)
            End If
        Else
            If Me.cbrthis.FindControl(, conMenu_Edit_Insert, , True).Enabled = True And _
               Me.cbrthis.FindControl(, conMenu_Edit_Insert, , True).Visible = True Then
                Call SampleDisposal(mActS.合并标本)
            End If
        End If
    ElseIf Row.Record(mCol.标本类型).Icon = 3 Then
        Call frmLabMainLJ.ShowMe(mlngKey, Me, mlngMachineID)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptList_SelectionChanged()
    Dim strSampleType As String                     '标本类型-1=普通,3=质控,4=比对
    Dim strEmergen                                  '急诊 -1="普通",1=急
    Dim strState                                    '状态 7,8=已检验
    Dim i As Integer                                '临时变量
    Dim str金额 As String
    Dim rs As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim lngSampleID As Long
    Dim intRow As Integer
    Dim strPricegrade As String                     '价格等级
    
    Dim tmp As Double
    
    On Error GoTo errH
    If Me.Visible = False Then Exit Sub
    
    strSampleType = ""
    strEmergen = ""
    strState = ""
    
    Select Case mintEditState
        Case 1, 2, 4
            If Me.rptList.Tag = "" And mlngKey <> Me.rptList.FocusedRow.Record(mCol.ID).Value Then
                lngSampleID = mfrmRequest.ZlSave()
                mintEditState = 0
                If lngSampleID = 0 Then
                    mfrmRequest.ZlCancel
                Else
                    intRow = Me.rptList.FocusedRow.Index
                    InsertOneRecored lngSampleID, False
                    Me.rptList.FocusedRow = Me.rptList.Rows(intRow)
                End If
                
                
                gintSelectFocus = 1
'                Exit Sub
            Else
'                Me.rptList.SetFocus
                gintSelectFocus = 2
                
            End If
        Case 5
            If TabCtlWindow.Item(0).Selected = True Then
                mfrmWrite.ZlSave
                mfrmWrite.ZlCancel
                mfrmWrite.zlRefresh mlngKey
            Else
                mfrmWrite2.ZlSave
                mfrmWrite2.ZlCancel
                mfrmWrite2.zlRefresh mlngKey
            End If
            mintEditState = 0
    End Select
    
    If Me.rptList.FocusedRow Is Nothing Then
        mlngKey = 0
        strSampleType = ""
    Else
        If Me.rptList.FocusedRow.Record(mCol.ID).Value = mlngKey And mblnCompelRefresh = False Then
            '同一ID时不刷新
            Exit Sub
        End If
        mblnCompelRefresh = False
        mlngKey = Me.rptList.FocusedRow.Record(mCol.ID).Value
        i = Me.rptList.FocusedRow.Record(mCol.标本类型).Icon
        If i = -1 Then
            strSampleType = "普通样本"
        ElseIf i = 3 Then
            strSampleType = "质控样本"
        Else
            strSampleType = "比对样本"
        End If
        
        i = Me.rptList.FocusedRow.Record(mCol.紧急).Icon
        If i = 1 Then
            strEmergen = "紧急"
        End If
        
        i = Me.rptList.FocusedRow.Record(mCol.执行状态).Icon
        If i = 7 Or i = 8 Then
            strState = "已检验"
        Else
            strState = "检验中"
        End If
        
        
    End If
    
    If mintLoadShow = 0 Then Exit Sub
    
    Call mfrmRequest.zlRefresh(Me.rptList.FocusedRow)
    Call mfrmLabMicrobe3Report.zlRefresh(mlngKey)
    
    
'    If Me.rptList.FocusedRow Is Nothing Then
'        Call mfrmWrite.zlRefresh(mlngKey)
'    ElseIf Val(Me.rptList.FocusedRow.Record(mCol.微生物标本).Value) = 1 Then
'        Call mfrmWrite2.zlRefresh(mlngKey)
'    Else
'        Call mfrmWrite.zlRefresh(mlngKey)
'    End If
'
'    Call mfrmRequest.zlRefresh(mlngkey)
    
    RefreshTableWindow Me.TabCtlWindow.Selected.Index
    If mlngKey <> 0 Then
        ReadImageData mlngKey, False
    End If
    
    
    
    Set cbrControl = Me.cbrChild.FindControl(, conMenu_View_FindType, True, True)
    If Not cbrControl Is Nothing Then
        If mlngKey > 0 Then '显示项目数和价格
        
            '获取价格等级
            With Me.rptList.FocusedRow
                strPricegrade = GetAdvicePrice(Val(.Record(mCol.病人ID).Value), Val(.Record(mCol.主页ID).Value))
            End With
            
            If Val(Me.rptList.FocusedRow.Record(mCol.微生物标本).Value) = 1 Then
                gstrSql = "Select /*+ rule */" & vbNewLine & _
                        " Sum(Nvl(收费数量, 0) * Nvl(现价, 0)) As 金额" & vbNewLine & _
                        "From ( --- 根据标本记录中的 条码，医嘱id,相关id，得到对应的诊疗项目id" & vbNewLine & _
                        "       Select C.诊疗项目id" & vbNewLine & _
                        "       From 病人医嘱记录 C, 病人医嘱发送 B, 检验标本记录 A" & vbNewLine & _
                        "       Where B.医嘱id = C.ID And A.样本条码 = B.样本条码 And A.ID = [1]" & vbNewLine & _
                        "       Union" & vbNewLine & _
                        "       Select C.诊疗项目id" & vbNewLine & _
                        "       From 病人医嘱记录 C, 检验标本记录 A" & vbNewLine & _
                        "       Where A.医嘱id = C.ID And A.ID = [1]" & vbNewLine & _
                        "       Union" & vbNewLine & _
                        "       Select C.诊疗项目id From 病人医嘱记录 C, 检验标本记录 A Where A.医嘱id = C.相关id And A.ID = [1]) A," & vbNewLine & _
                        "     (Select E.诊疗项目id, E.收费数量, F.现价, J.编码, J.名称" & vbNewLine & _
                        "       From 诊疗收费关系 E, 收费价目 F, 收费项目目录 J" & vbNewLine & _
                        "       Where F.收费细目id = J.ID And E.收费项目id = F.收费细目id And (F.终止日期 Is Null Or F.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) and f.价格等级" & strPricegrade & ") B" & vbNewLine & _
                        "Where A.诊疗项目id = B.诊疗项目id"

            Else
    
                gstrSql = "Select /*+ rule */" & vbNewLine & _
                            " Sum(Nvl(收费数量, 0) * Nvl(现价, 0)) As 金额" & vbNewLine & _
                            "From (Select Distinct 诊疗项目id" & vbNewLine & _
                            "       From (Select 诊疗项目id" & vbNewLine & _
                            "              From 检验普通结果" & vbNewLine & _
                            "              Where 检验标本id = [1] And 诊疗项目id Is Not Null" & vbNewLine & _
                            "              Union All" & vbNewLine & _
                            "              Select B.诊疗项目id" & vbNewLine & _
                            "              From 检验普通结果 A, 检验报告项目 B, 诊疗项目目录 C" & vbNewLine & _
                            "              Where A.检验项目id = B.报告项目id And B.诊疗项目id = C.ID And C.组合项目 = 0" & vbNewLine & _
                            "              And A.检验标本id = [1] And A.诊疗项目id Is Null)) A," & vbNewLine & _
                            "     (Select E.诊疗项目id, E.收费数量, F.现价, J.编码, J.名称" & vbNewLine & _
                            "       From 诊疗收费关系 E, 收费价目 F, 收费项目目录 J" & vbNewLine & _
                            "       Where F.收费细目id = J.ID And E.收费项目id = F.收费细目id" & vbNewLine & _
                            "             And (F.终止日期 Is Null Or F.终止日期 = To_Date('3000-01-01', 'yyyy-mm-dd')) and f.价格等级" & strPricegrade & ") B" & vbNewLine & _
                            "Where A.诊疗项目id = B.诊疗项目id"
            End If
            Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            str金额 = ""
            If rs.RecordCount > 0 Then
                If Val("" & rs.Fields("金额")) <> 0 Then
                    str金额 = "，项目计价 " & Format("" & rs.Fields("金额"), "0.00")
                End If
            End If

            gstrSql = "Select Count(检验标本id) as 项目数 from 检验普通结果 where 检验结果 is not null And 检验标本id =[1] "
            Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            If rs.RecordCount > 0 Then
                str金额 = "共有" & rs.Fields("项目数") & "项结果" & str金额 & "  "

            End If
        End If '--显示项目数和价格
        
        cbrControl.Caption = str金额 & "   状态;" & strEmergen & " " & strState & " " & strSampleType
        Me.cbrChild.RecalcLayout
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub rptList1_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    frmLabMainSizer.ShowMe Me, "待核收", True
    Call GetWaitVerify
    If Me.picFilter.Tag = "True" Then Call RptListFilter
    Me.picFilter.Tag = ""
End Sub

Private Sub rptList1_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    
    On Error Resume Next
    If Button = 2 Then
        If rptList1.Records.Count <= 0 Then Exit Sub
        If Not rptList1.SelectedRows(0).GroupRow Then
            Set objPopup = Me.cbrthis.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Audit, "报告审核(&A)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "取消审核(&U)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "状态回滚(&Z)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "样本复查(&D)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消复查(&E)")
                Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "报告查询(&P)"): cbrControl.BeginGroup = True
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Apply, "发往仪器(&S)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_LIS_TOQC, "置为质控(&Q)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "置为比对(&Y)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportView, "查看比对(&B)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "糖耐量合并(&E)")
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改样本号(&M)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "删除样本(&D)")
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Refuse, "拒收(&J)"): cbrControl.BeginGroup = True
        
                Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)"): cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_LIS_Cancel, "放弃(&C)")
            End With
            objPopup.ShowPopup
        End If
    End If
End Sub

Private Sub rptList1_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If mintEditState = 0 Then Call SampleDisposal(mActS.核收)
End Sub

Private Sub TabCtlWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible = True And mTableRefresh = False Then
        RefreshTableWindow Item.Index
        Me.TabCtlWindow.Tag = Item.Index
    End If
End Sub


Private Sub RefreshTableWindow(Index As Integer)
    Dim blnCurrMoved As Boolean                                             '是否转出
    Dim lngAdviceID As Long                                                 '医嘱ID
    Dim intReportCount As Integer                                           '报告总数
    Dim blMicrobe As Boolean                                                '是否是微生物
    Dim cbrControl As CommandBarControl                                     '工具条中按钮对象
    Dim strPatientType As String                                            '病人来源
    Dim str挂号单 As String                                                 '挂号单
    Dim lngPatientID As Long                                                '病人ID
    Dim intHomePage As Integer                                              '主页ID
    Dim lngPatientDeptID As Long                                            '病人科室ID
    Dim blnShowButtonText As Boolean                                        '显示按钮文本
    Dim lngCount As Long
    Dim cbrCustom As CommandBarControlCustom
    Dim lng挂号ID As Long
    Dim strPrivs As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            lngAdviceID = Val(.Record(mCol.医嘱id).Value)
            blnCurrMoved = (.Record(mCol.转出).Value = "√")
            intReportCount = Val(.Record(mCol.结果次数).Value)
            blMicrobe = IIf(Val(.Record(mCol.微生物标本).Value) = 1, True, False)
            strPatientType = .Record(mCol.所属情况).Value
            str挂号单 = .Record(mCol.挂号单).Value
            lngPatientID = Val(.Record(mCol.病人ID).Value)
            intHomePage = Val(.Record(mCol.主页ID).Value)
            lngPatientDeptID = Val(.Record(mCol.开嘱科室ID).Value)
        End With
    End If
    
'    Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = False
    
    '如果是微生物时处理比对为不显示
    If blMicrobe = True Then
        Me.TabCtlWindow.Item(3).Visible = False
        Me.TabCtlWindow.Item(2).Visible = True
    Else
        Me.TabCtlWindow.Item(3).Visible = True
        Me.TabCtlWindow.Item(2).Visible = False
    End If
    
    '删除生成的按钮
    DelButton Index
    '刷新子窗口菜单
'    Call LockWindowUpdate(Me.Hwnd)
    
'    If Me.TabCtlWindow.Selected.Index <> Val(Me.TabCtlWindow.Tag) Then
'        '删除现在的工具栏及顶级菜单项
'        For lngCount = cbrthis.ActiveMenuBar.Controls.Count To 1 Step -1
'            cbrthis.ActiveMenuBar.Controls(lngCount).Delete
'        Next
'        For lngCount = cbrthis.Count To 2 Step -1
'            cbrthis(lngCount).Delete
'        Next
'        '重新创建菜单
'        Call CreateCbs
'    End If
    
    If strPatientType = "住院" Then
        
        If Me.TabCtlWindow.Item(7).Tag = "住院医嘱" Then
            Me.TabCtlWindow.Item(6).Visible = False
            Me.TabCtlWindow.Item(7).Visible = True
            If Index = 6 Or Index = 7 Then Index = 7: Me.TabCtlWindow.Item(7).Selected = True
        Else
            Me.TabCtlWindow.Item(6).Visible = False
            Me.TabCtlWindow.Item(7).Visible = False
        End If
        If TabCtlWindow.ItemCount >= 10 Then
            If Me.TabCtlWindow.Item(9).Tag = "住院病历" Then
                Me.TabCtlWindow.Item(9).Visible = True
                Me.TabCtlWindow.Item(8).Visible = False
                If Index = 8 Or Index = 9 Then Index = 9: Me.TabCtlWindow.Item(9).Selected = True
            Else
                Me.TabCtlWindow.Item(8).Visible = False
                Me.TabCtlWindow.Item(9).Visible = False
            End If
        End If
        If TabCtlWindow.ItemCount >= 11 Then
            '电子病历
            strPrivs = GetPrivFunc(glngSys, 2252)
            Me.TabCtlWindow.Item(10).Visible = IIf(strPrivs <> "", True, False)
        End If
    Else

        If Me.TabCtlWindow.Item(6).Tag = "门诊医嘱" Then
            Me.TabCtlWindow.Item(6).Visible = True
            Me.TabCtlWindow.Item(7).Visible = False
            If Index = 6 Or Index = 7 Then Index = 6: Me.TabCtlWindow.Item(6).Selected = True
        Else
            Me.TabCtlWindow.Item(6).Visible = False
            Me.TabCtlWindow.Item(7).Visible = False
        End If
        If Me.TabCtlWindow.Item(8).Tag = "门诊病历" Then
            Me.TabCtlWindow.Item(8).Visible = True
            If TabCtlWindow.ItemCount >= 10 Then
                Me.TabCtlWindow.Item(9).Visible = False
            End If
            If Index = 8 Or Index = 9 Then Index = 8: Me.TabCtlWindow.Item(8).Selected = True
        Else
            Me.TabCtlWindow.Item(8).Visible = False
            If TabCtlWindow.ItemCount >= 10 Then
                Me.TabCtlWindow.Item(9).Visible = False
            End If
        End If
        If TabCtlWindow.ItemCount >= 11 Then
            '电子病历
            strPrivs = GetPrivFunc(glngSys, 2251)
            Me.TabCtlWindow.Item(10).Visible = IIf(strPrivs <> "", True, False)
        End If
    End If
    
    Select Case Index
        Case 0, 1, 2  '普通结果和微生物结果
            If blMicrobe = True Then
                Me.TabCtlWindow.Item(0).Visible = False
                Me.TabCtlWindow.Item(1).Visible = True
                Me.TabCtlWindow.Item(2).Visible = True
                If mintEditState <> 5 Then
                    mfrmWrite2.zlRefresh mlngKey
                End If
                If Index = 0 Then
                    Me.TabCtlWindow.Item(1).Selected = True
                Else
                    Me.TabCtlWindow.Item(Index).Selected = True
                End If
'                Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).HideFlags = xtpHideGeneric
'                Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = False
            Else
                Me.TabCtlWindow.Item(0).Visible = True
                Me.TabCtlWindow.Item(1).Visible = False
                Me.TabCtlWindow.Item(2).Visible = False
                If mintEditState <> 5 Then
                    mfrmWrite.zlRefresh mlngKey
                End If
                Me.TabCtlWindow.Item(0).Selected = True
'                Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = False
'                Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = False
            End If
        Case 3 '历史比对
'            Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = False
'            Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = False
            zlCommFun.ShowFlash " 请稍待正在读出病人历史数据..."
            mfrmTrack.zlRefresh mlngKey
            zlCommFun.StopFlash
        Case 4  '费用
            '第一次打开时再加载
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsExpenses Is Nothing Then
                Set mclspublicExpenses = New zlPublicExpense.clsPublicExpense
                Call mclspublicExpenses.zlInitCommon(glngSys, gcnOracle, gstrDBUser)
                Set mclsExpenses = New zlPublicExpense.clsDockExpense       '费用部件
                mcolSubForm.Add mclsExpenses.zlGetForm, "_费用"             '得到子窗体
            End If
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "费用查询", mcolSubForm("_费用").hWnd, conMenu_Edit_Price).Tag = "费用查询"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            mclsExpenses.zlDefCommandBars Me, Me.cbrthis
            strSQL = "select a.id as 医嘱ID, b.发送号 from 病人医嘱记录 a,病人医嘱发送 b " & vbCrLf & _
                    " Where a.ID = b.医嘱id And a.相关id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngAdviceID)
            If rsTmp.EOF = False Then
                mclsExpenses.zlRefresh mlngDeptID, rsTmp(0) & ":" & rsTmp(1), blnCurrMoved
            End If
'            DelButton Index  '生成前先删除按钮
            
            '是否显示按钮的文字
            blnShowButtonText = Me.cbrthis.FindControl(, conMenu_View_ToolBar_Text, True, True).Checked
            For Each cbrControl In Me.cbrthis(2).Controls
                cbrControl.Style = IIf(blnShowButtonText, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            
            
            strSQL = "select distinct c.名称,b.相关id as 医嘱ID from 检验项目分布 a , 病人医嘱记录 b , 诊疗项目目录 c " & _
                     " where a.医嘱id = b.相关ID and b.诊疗项目ID = c.id and  a.标本ID =[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "费用查询", mlngKey)
            
            If cboExesItem.ListCount > 0 And cboExesItem.ListIndex <> -1 Then lngAdviceID = cboExesItem.ItemData(cboExesItem.ListIndex)
            Me.cboExesItem.Clear
            
            Do Until rsTmp.EOF
                Me.cboExesItem.AddItem rsTmp("名称")
                Me.cboExesItem.ItemData(Me.cboExesItem.NewIndex) = rsTmp("医嘱ID")
                If rsTmp("医嘱ID") = lngAdviceID Then Me.cboExesItem.ListIndex = Me.cboExesItem.NewIndex
                rsTmp.MoveNext
            Loop
            If cboExesItem.ListCount > 0 Then
                If cboExesItem.ListIndex = -1 Then cboExesItem.ListIndex = 0
            End If
'            Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = True
'            Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = True
'            Me.cbrthis.RecalcLayout
'            Me.cbrChild.RecalcLayout
        Case 5  '合并
'            Me.cbrChild.FindControl(, conMenu_Manage_Transfer_Send).Visible = True
'            Me.cbrChild.FindControl(, conMenu_Edit_UnArchive).Visible = True
        Case 6 '门诊医嘱
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsOutAdvices Is Nothing Then
                Set mclsOutAdvices = New zlCISKernel.clsDockOutAdvices      '门诊医嘱
                mcolSubForm.Add mclsOutAdvices.zlGetForm, "_门诊医嘱"
            End If
            '第一次打开时再加载
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "门诊医嘱", mcolSubForm("_门诊医嘱").hWnd, 1).Tag = "门诊医嘱"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(6).Visible = True Then
'                DelButton Index  '生成前先删除按钮
                mclsOutAdvices.zlDefCommandBars Me, Me.cbrthis, 2
                '是否显示按钮的文字
                blnShowButtonText = Me.cbrthis.FindControl(, conMenu_View_ToolBar_Text, True, True).Checked
                For Each cbrControl In Me.cbrthis(2).Controls
                    cbrControl.Style = IIf(blnShowButtonText, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
'                Me.cbrthis.RecalcLayout
    '            MsgBox "病人ID:" & lngPatientID & ";挂号单:" & str挂号单
                mclsOutAdvices.zlRefresh lngPatientID, str挂号单, True, False, lngAdviceID, mlngDeptID
            End If
        Case 7 '住院医嘱
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsInAdvices Is Nothing Then
                Set mclsInAdvices = New zlCISKernel.clsDockInAdvices
                mcolSubForm.Add mclsInAdvices.zlGetForm, "_住院医嘱"
            End If
            '第一次打开时再加载
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "住院医嘱", mcolSubForm("_住院医嘱").hWnd, 1).Tag = "住院医嘱"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(7).Visible = True Then
'                DelButton Index  '生成前先删除按钮
                mclsInAdvices.zlDefCommandBars Me, Me.cbrthis, 2
                '是否显示按钮的文字
                blnShowButtonText = Me.cbrthis.FindControl(, conMenu_View_ToolBar_Text, True, True).Checked
                For Each cbrControl In Me.cbrthis(2).Controls
                    cbrControl.Style = IIf(blnShowButtonText, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
'                Me.cbrthis.RecalcLayout
    '            MsgBox "病人ID:" & lngPatientID & ";主页ID:" & intHomePage & ";病区ID:" & lngPatientDeptID & ";病人科室ID;" & lngPatientDeptID
                mclsInAdvices.zlRefresh lngPatientID, intHomePage, lngPatientDeptID, lngPatientDeptID, 0, False, lngAdviceID, 0, mlngDeptID
            End If
        Case 8 '门诊病历
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsOutEPRs Is Nothing Then
                Set mclsOutEPRs = New zlRichEPR.cDockOutEPRs                '门诊医历
                mcolSubForm.Add mclsOutEPRs.zlGetForm, "_门诊病历"
            End If
            '第一次打开时再加载
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "门诊病历", mcolSubForm("_门诊病历").hWnd, 1).Tag = "门诊病历"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(8).Visible = True Then
                gstrSql = "select ID from 病人挂号记录 where 记录状态=1 and 记录性质=1 and no = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str挂号单)
                If rsTmp.EOF Then
                    lng挂号ID = 0
                Else
                    lng挂号ID = Nvl(rsTmp("ID"))
                End If
                mclsOutEPRs.zlRefresh lngPatientID, lng挂号ID, mlngDeptID, False
            End If
        Case 9 '住院病历
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsInEPRs Is Nothing Then
                Set mclsInEPRs = New zlRichEPR.cDockInEPRs                  '住院病历
                mcolSubForm.Add mclsInEPRs.zlGetForm, "_住院病历"
            End If
            '第一次打开时再加载
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "住院病历", mcolSubForm("_住院病历").hWnd, 1).Tag = "住院病历"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(9).Visible = True Then
                mclsInEPRs.zlRefresh lngPatientID, intHomePage, lngPatientDeptID
            End If
        Case 10 '电子病历
            On Error Resume Next
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
            End If
            If mclsEMR Is Nothing Then
                Set mclsEMR = CreateObject("zlRichEMR.clsDockEMR")
                If Not mclsEMR Is Nothing Then
                    Set gobjEmr = gfrmMain.mobjEMR
                    If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                      Set mclsEMR = Nothing
                    End If
                End If
                mcolSubForm.Add mclsEMR.zlGetForm, "_电子病历"
            End If
            With Me.TabCtlWindow
                If .Item(Index).Handle = PicWindows.hWnd Then
                    mTableRefresh = True
                    .RemoveItem (Index)
                    .InsertItem(Index, "电子病历", mcolSubForm("_电子病历").hWnd, 1).Tag = "电子病历"
                    .Item(Index).Selected = True
                    mTableRefresh = False
                End If
            End With
            If Me.TabCtlWindow.Item(10).Visible = True And Not mclsEMR Is Nothing And lngPatientID <> 0 Then
                If strPatientType = "住院" Then
                    mclsEMR.zlRefresh lngPatientID, intHomePage, lngPatientDeptID, 0, 2
                Else
                    gstrSql = "select ID from 病人挂号记录 where 记录状态=1 and 记录性质=1 and no = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str挂号单)
                    If rsTmp.EOF Then
                        lng挂号ID = 0
                    Else
                        lng挂号ID = Nvl(rsTmp("ID"))
                    End If
                    mclsEMR.zlRefresh lngPatientID, lng挂号ID, lngPatientDeptID, 0, 1
                End If
            End If
    End Select
    
    
'    Me.cbrthis.FindControl(, conMenu_Edit_Insert).Visible = IIf(Index = 4, True, False)
'    cbrThis.ActiveMenuBar.FindControl(, conMenu_LIS_RightMenu).Visible = False
    
'    Me.cbrthis.RecalcLayout
'    Me.cbrChild.RecalcLayout
    
    '如果用了RecalcLayout反而不正常
'    Call LockWindowUpdate(0)
    
    
    mTableRefresh = False
    Exit Sub
errH:
    mTableRefresh = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Property Let AutoRefresh(vData As Boolean)
    '
    '功能:自动刷新
    '
'
End Property
Private Sub QUFilter()
    '功能        快速查询
    Dim strCondition As String
    '定义查询
    AutoRefresh = False
    strCondition = rptList.Tag
    frmLabFilter.ShowMe Me, mlngDeptID, mlngMachineID, mstrMachineALL, strCondition
    If strCondition <> "" Then
        rptList.Tag = strCondition & ";" & 0                            '在最后加上病人ID
        zlCommFun.ShowFlash "正在更新数据请稍候...", Me
        RefreshData True
        zlCommFun.StopFlash
    End If
    AutoRefresh = True
End Sub


Private Sub GetSaveSetup(Mode As Integer)
    '功能提取保存记录
    '参数              =1单一 =2批量
    Dim strFile As String, lngDeviceID As Long, dtStart As Date, dtEnd As Date, strSampleNO As String
    Dim lngMachineID As Long                                '仪器ID
    Dim strSampltDate As String                             '标本时间
    Dim strSampltID As String                                '标本号
    
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            lngMachineID = Val(.Record(mCol.仪器id).Value)
            strSampltDate = .Record(mCol.标本时间).Value
            strSampltID = .Record(mCol.标本号).Value
        End With
    End If
    
    Me.MousePointer = vbHourglass

    If Mode = 1 Then
        strFile = zlDatabase.GetPara("仪器数据文件", 100, 1208, "")
        lngDeviceID = zlDatabase.GetPara("文件提取仪器", 100, 1208, -1)
       '27693  自动导入产生错误日志-提示"类型不匹配"
        '原代码： strSampleNO = strSampltID
        strSampleNO = IIf(Trim(strSampltID) = "", "-1", strSampltID)
        
        dtStart = Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")
        dtEnd = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
        GetResultFromFile WinsockC, WinsockC.LocalIP, strFile, lngDeviceID, strSampleNO, dtStart, dtEnd
    Else
        strFile = zlDatabase.GetPara("仪器数据文件", 100, 1208, "")
        lngDeviceID = zlDatabase.GetPara("文件提取仪器", 100, 1208, -1)
        If Val(zlDatabase.GetPara("文件提取范围", 100, 1208, 0)) = 0 Then '提取当天
            dtStart = Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")
            dtEnd = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
        Else
            dtStart = CDate(Format(zlDatabase.GetPara("文件提取开始日期", 100, 1208, zlDatabase.Currentdate), "yyyy-mm-dd 00:00:00"))
            dtEnd = CDate(Format(zlDatabase.GetPara("文件提取结束日期", 100, 1208, zlDatabase.Currentdate), "yyyy-mm-dd 23:59:59"))
        End If
        GetResultFromFile WinsockC, WinsockC.LocalIP, strFile, lngDeviceID, -1, dtStart, dtEnd
    End If
    Me.MousePointer = vbDefault
    '刷新
    RefreshData
End Sub
Private Sub PrintSetup()
    '打印设置
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long
    Dim strSQL As String
    
    On Error GoTo errH
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    lng医嘱ID = Val(rptList.FocusedRow.Record(mCol.医嘱id).Value)
    lng病人ID = Val(rptList.FocusedRow.Record(mCol.病人ID).Value)
    
    strSQL = "select 发送号 from 病人医嘱发送 a , 病人医嘱记录 b where b.id = a.医嘱id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng医嘱ID)
    If rsTmp.EOF = False Then
        lng发送号 = Nvl(rsTmp(0))
    End If
    
    If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        ReportPrintSet gcnOracle, glngSys, strReportCode, Me
        
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ReportPrint(ByVal blnPrint As Boolean)
    '单个报告打印
    
    Dim strReportCode As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lng医嘱ID As Long, lng发送号 As Long, lng病人ID As Long, lng病人科室ID As Long, str姓名 As String
    Dim strSQL As String

    Dim intLoop As Integer
    On Error GoTo errH
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    lng医嘱ID = Val(rptList.FocusedRow.Record(mCol.医嘱id).Value)
    lng病人ID = Val(rptList.FocusedRow.Record(mCol.病人ID).Value)
    lng发送号 = Val(rptList.FocusedRow.Record(mCol.发送号).Value)
    lng病人科室ID = Val(rptList.FocusedRow.Record(mCol.病人科室ID).Value)
    str姓名 = rptList.FocusedRow.Record(mCol.姓名).Value
    
    '如果没有选择科室就退出
    If InStr("," & mstrPrintDepts & ",", "," & lng病人科室ID & ",") <= 0 And str姓名 <> "" Then
        Exit Sub
    End If
    
    strSQL = "select 发送号 from 病人医嘱发送 a , 病人医嘱记录 b where b.id = a.医嘱id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lng医嘱ID)
    If rsTmp.EOF = False Then
        lng发送号 = Nvl(rsTmp(0))
    End If
    
    If lng医嘱ID = 0 And lng发送号 = 0 Then
        Set rsTmp = zlDatabase.OpenSQLRecord("Select 编号 From ZlReports Where 编号 Like '%-N'", Me.Caption)
        If Not rsTmp.EOF Then
            strReportCode = rsTmp(0)
            Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "ID=" & mlngKey, IIf(blnPrint, 2, 1))
        End If
        Exit Sub
    End If
    
    blnCurrMoved = rptList.SelectedRows(0).Record.Item(mCol.转出).Value = "√"
    Call Open_LIS_Report(Me, lng医嘱ID, lng发送号, lng病人ID, mlngKey, blnCurrMoved, blnPrint)
    
    
    If blnPrint = True And Me.rptList.FocusedRow.Record(mCol.执行状态).Value = "已检验" Then
        If mintUnion = 1 Then
            gstrSql = " select id from 检验标本记录 where 医嘱id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lng医嘱ID)
            Do Until rsTmp.EOF
                strSQL = "ZL_检验标本记录_标本质控(" & rsTmp("ID") & ",'',1)"
                zlDatabase.ExecuteProcedure strSQL, gstrSysName
                rsTmp.MoveNext
            Loop
        Else
            strSQL = "ZL_检验标本记录_标本质控(" & mlngKey & ",'',1)"
            zlDatabase.ExecuteProcedure strSQL, gstrSysName
        End If
        Me.rptList.FocusedRow.Record(mCol.执行状态).Value = "已打印"
        Me.rptList.FocusedRow.Record(mCol.执行状态).Icon = 8
        Me.rptList.Populate
    End If
    

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub SetParameter()
    Dim blnExec As Boolean
    If frmLisStationPara.ShowPara(Me) Then
        AutoRefresh = True
        blnAutoRefresh = Val(zlDatabase.GetPara("自动刷新", 100, 1208, 1))
        blnComm = Val(zlDatabase.GetPara("核收允许双向", 100, 1208, 0))
        blnAutoPrint = zlDatabase.GetPara("审核打印", 100, 1208, 0)
        int体检处理方式 = Val(zlDatabase.GetPara("体检病人信息不一致的处理方式", 100, 1208, True, 1))
        int院外处理方式 = Val(zlDatabase.GetPara("院外病人信息不一致的处理方式", 100, 1208, True, 1))
        int住院处理方式 = Val(zlDatabase.GetPara("住院病人信息不一致的处理方式", 100, 1208, True, 1))
        int门诊处理方式 = Val(zlDatabase.GetPara("门诊病人信息不一致的处理方式", 100, 1208, True, 1))
        
        blnExec = Val(zlDatabase.GetPara("只在核收登记时显示登记窗口", 100, 1208, 0))
        frmLabRequest.mMakeNoRule = zlDatabase.GetPara("标本序号生成规则", 100, 1208, "今  天")
        mMakeNoRule = zlDatabase.GetPara("标本序号生成规则", 100, 1208, "今  天")
        mSendReport = zlDatabase.GetPara("使用二级报告审核", 100, 1208, 0)
        mstrPrintDepts = zlDatabase.GetPara("只打指定科室报告单", 100, 1208, "")
        mblnAout = zlDatabase.GetPara("审核后跳到下一个可审标本", 100, 1208, mblnAout)

        
        Call ShowRequest(Not blnExec)
        cbo时间.Text = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0)
        Me.dtpDate.Value = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";" & Format(Now, "yyyy-mm-dd"), ";")(1)
        Me.dtpDateEnd.Value = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";" & Format(Now, "yyyy-mm-dd") & ";" & Format(Now, "yyyy-mm-dd"), ";")(2)
        mfrmRequest.SetPara
        RefreshData
    End If
End Sub

Private Function RefreshData(Optional blWhere As Boolean) As Boolean
    '功能               '刷新左边mcol(检验中数据)
    '参数               '是否使用条件查询
    Dim strSQL As String
    Dim strSQLbak As String
    Dim rsItem As New ADODB.Recordset
    Dim blnMoved As Boolean                                         '是否移出
    Dim bln住院病人 As Boolean                                      '住院病人
    Dim bln门诊病人 As Boolean                                      '门诊病人
    Dim bln无主验单 As Boolean                                      '有无主验单
    Dim strStart As String                                          '检验开始时间
    Dim strEnd As String                                            '检验结束时间
    Dim Record As ReportRecord                                      '列表记录集
    Dim Item As ReportRecordItem                                    '列表中每一行对象
    Dim Rerow As ReportRow                                          '行对象
    Dim intLoop As Integer                                          '循环临时变量
    Dim lngloop As Long                                             '
    Dim varFilter As Variant                                        '过滤字串数组
    Dim varUnionFilter As Variant                                   '组合查询
    Dim varItem As Variant                                          '组合查询下子项
    Dim intAgeBeging As Integer                                     '年龄开始
    Dim intAgeEnd As Integer                                        '年龄结束
    Dim lngRow As Long                                              '刷新前记录当前行号
    Dim strSample As String                                         '标本序号
    Dim lngAdvice As Long                                           '医嘱号
    Dim lngSampleID As Long                                         '检验标本ID
    Dim strWhere As String                                          '要增加的条件
    Dim strTable As String                                          '要增加的表
    Dim strDeptID As String                                         '科室ID
    Dim strUserMachine  As String                                   '当前用户可以使用的仪器ID
    Dim strTmp As String
    Dim strStartNO As String, strEndNO As String                    '开始和结束NO
    Dim lngRowIndex As Long                                         '行索行
    Dim lngRowID As Long                                            '行ID
    Dim blnPathPatient As Boolean                                   '临床路径病人
    Dim lngUnionItem As Long                                        '组合项目ID

    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    mblnCompelRefresh = True    '刷新时可以强制刷新
    zlCommFun.ShowFlash "正在读取数据数据请等待...", Me
'    Me.stbThis.Panels(2).Text = "正在读取数据数据请等待..."
    Me.MousePointer = 11
            
    On Error GoTo errH
    
    If Not Me.rptList.FocusedRow Is Nothing Then
        lngRow = Me.rptList.FocusedRow.Index
    End If
    
    If cboUnionItem.ListCount > 0 Then
        If (Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex) > 0 And rptList.Tag = "") Or _
        (Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex) = -1 And rptList.Tag = "") Then
            strTable = " ,检验申请项目 C "
            strWhere = " And a.id = c.标本ID "
        End If
    End If
    If rptList.Tag <> "" Then
        varFilter = Split(rptList.Tag, ";")
        If varFilter(mFilter.高级) <> "" And varFilter(mFilter.是否使用高级) = 1 Or _
            InStr(1, varFilter(mFilter.检验项目), ",") > 1 Or varFilter(mFilter.细菌) <> "" Or varFilter(mFilter.抗生素) <> "" _
            Or varFilter(mFilter.药敏结果) <> "" Then
            
            strTable = strTable & " ,检验普通结果 G "
            strWhere = strWhere & " And a.id = g.检验标本ID "
            
            If varFilter(mFilter.抗生素) <> "" Or varFilter(mFilter.药敏结果) <> "" Then
                strTable = strTable & "  , 检验药敏结果 O "
                strWhere = strWhere & " and g.id = O.细菌结果ID "
            End If
        End If
    End If
    
    strSQL = "Select /*+ rule */  distinct      Decode(a.是否传送, 1, '', '传送失败') As 传送," & vbNewLine & _
            "       decode(a.标本类别,1,'紧急',decode(a.紧急,1,'紧急', '')) As 紧急,Decode(a.样本状态, 1, '检验中', 2, '已检验') As 执行状态," & vbNewLine & _
            "       Decode(A.病人来源, 1, '门诊', 2, '住院', 3, '院外', 4, '体检','无主') As 所属情况," & vbNewLine & _
            "       Decode(Sign(Nvl(a.是否质控品, 0)), 0, '普通', 1, '质控', -1, '比对') As 标本类型," & vbNewLine & _
            "       Decode(a.仪器id, Null," & vbNewLine & _
            "                 To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
            "                 a.标本序号) As 标本号显示,a.标本序号, A.挂号单 ," & vbNewLine & _
            "       Decode(A.病人来源, 1, to_char(nvl(a.门诊号,a.标识号)), 2, to_char(nvl(a.住院号,a.标识号)), 3, to_char(nvl(a.NO,a.标识号)), 4, to_char(nvl(a.门诊号,a.标识号)),to_char(a.标识号)) As 标识号,a.姓名,a.性别,a.年龄," & vbNewLine & _
            "       Decode(a.病人来源,2,S.病人类型,b.病人类型) as 病人类型," & vbNewLine & _
            "       a.报告结果 As 结果次数,a.医嘱ID,a.仪器ID,'' As 转出,a.Id,a.核收时间 ,a.打印次数,a.病人id," & vbNewLine & _
            "       a.检验时间,a.微生物标本,a.检验人,a.审核人,To_Char(A.婴儿) As 婴儿,a.样本条码,a.申请科室ID As 开嘱科室id," & vbNewLine & _
            "       a.主页ID,a.报告结果,a.年龄数字,a.年龄单位,a.门诊号,a.住院号,a.出生日期,a.挂号单,a.检验项目,e.名称 as  申请科室,f.名称 as 仪器名称, " & vbNewLine & _
            "       a.申请科室ID as 病人科室ID,a.床号,a.申请人,a.标本形态,a.采样人,a.采样时间,a.标本类型 as 检验标本,a.NO,a.接收人,a.接收时间, " & vbNewLine & _
            "       abs(nvl(a.是否质控品,0)) as 比对次数,a.审核时间,n.名称 as 病区名称,a.执行科室ID,nvl(a.标本类别,0) as 标本类别, " & vbNewLine & _
            "       nvl(a.紧急,0) as 医嘱紧急,nvl(a.标本类别,0) as 标本紧急,decode(a.病人科室,null,M.名称,a.病人科室) as 病人科室, " & vbNewLine & _
            "       a.申请类型,nvl(r.查阅状态,0) as 查阅状态,nvl(r.病历ID,0) as 报告发送,a.初审人,a.初审时间,b.工作单位,p.项目,p.内容,b.健康号, " & vbNewLine & _
            "       a.审核未通过,a.病人来源,a.结果为空,nvl(s.路径状态,0) as 临床路径病人 ,decode(d.仪器是否审核,1,'仪器审核','仪器未审核')  as 仪器审核 " & vbNewLine & _
            " From 检验标本记录 a ,部门表 E , 检验仪器 f , 病人信息 b , 部门表 N , 部门表 M, 病人医嘱报告 R,病人医嘱附件 p,病案主页 S ,检验流水线标本 D " & strTable & vbNewLine & _
            " Where a.申请科室ID = E.id(+)  and a.仪器id=f.id(+) and a.病人id =b.病人id(+) and b.当前科室ID = M.id(+) and " & vbNewLine & _
            " b.当前病区id = n.id(+) And a.医嘱id = R.医嘱ID(+) and a.医嘱ID = P.医嘱ID(+) And p.项目(+)='任务团体' " & vbNewLine & _
            " and a.病人ID = S.病人ID(+) and a.主页ID = s.主页ID(+) and  a.id=d.标本id(+)  " & strWhere
                  
                  
                  
    If mlngDeptID > 0 And rptList.Tag = "" Then
        strSQL = strSQL & " And Instr(To_Char([2]), To_Char(a.执行科室id)) > 0 "
        strDeptID = mlngDeptID
    Else
        If InStr(mstrPrivs, "所有科室") = 0 Or InStr(mstrPrivs, "查看其他科室报告") > 0 Then
            For intLoop = 1 To Me.cboDept.ListCount - 1
                strDeptID = strDeptID & "," & Me.cboDept.ItemData(intLoop)
            Next
            strSQL = strSQL & " And Instr(To_Char([2]), To_Char(a.执行科室id)) > 0  "
        End If
    End If
    
    If cboUnionItem.ListCount > 0 Then
        If Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex) > 0 And rptList.Tag = "" Then
            strSQL = strSQL & " and c.诊疗项目Id = [16] "
        End If
        
        If Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex) = -1 And rptList.Tag = "" Then
            strSQL = strSQL & " and c.诊疗项目id is null "
        End If
    End If
    '使用过滤中的条件进行查询
    If rptList.Tag <> "" Then
        If varFilter(mFilter.检验时间) <> "," Then
            strStart = Mid(varFilter(mFilter.检验时间), 1, InStr(1, varFilter(mFilter.检验时间), ",") - 1)
            strEnd = Mid(varFilter(mFilter.检验时间), InStr(1, varFilter(mFilter.检验时间), ",") + 1)
            strSQL = strSQL & " And a.核收时间 Between [3] And [4] " & vbCrLf
            blnMoved = MovedByDate(CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")))
        Else
            strStart = Now
            strEnd = Now
        End If
        
        If varFilter(mFilter.姓名) <> "" Then
            strSQL = strSQL & " And a.姓名 like [5] "
        End If
        
        If varFilter(mFilter.性别) <> "" Then
            strSQL = strSQL & " And a.性别 = [6] "
        End If
        
        If varFilter(mFilter.年龄) <> "," Then
            If InStr(1, varFilter(mFilter.年龄), ",") = Len(varFilter(mFilter.年龄)) Then
                strSQL = strSQL & " And a.年龄数字 >= [7] And a.年龄单位 = [20] "
                intAgeBeging = Mid(varFilter(mFilter.年龄), 1, InStr(1, varFilter(mFilter.年龄), ",") - 1)
                intAgeEnd = 0
            ElseIf InStr(1, varFilter(mFilter.年龄), ",") = 1 Then
                strSQL = strSQL & " And a.年龄数字 <= [8] and a.年龄单位 = [20] "
                intAgeBeging = 0
                intAgeEnd = Mid(varFilter(mFilter.年龄), 2)
            Else
                strSQL = strSQL & " And a.年龄数字 between  [7] And  [8] And a.年龄单位 = [20] "
                intAgeBeging = Mid(varFilter(mFilter.年龄), 1, InStr(1, varFilter(mFilter.年龄), ",") - 1)
                intAgeEnd = Mid(varFilter(mFilter.年龄), InStr(1, varFilter(mFilter.年龄), ",") + 1)
            End If
        ElseIf varFilter(mFilter.年龄单位) <> "" Then
            strSQL = strSQL & " and a.年龄单位 = [20] "
        End If
        
        If varFilter(mFilter.标本号) <> "" Then
            If varFilter(mFilter.标本号) Like "0*-0*" Then
                varFilter(mFilter.标本号) = TransSampleNO(varFilter(mFilter.标本号))
                strSQL = strSQL & " And a.标本序号 = [9] and a.仪器id is null "
                strStartNO = varFilter(mFilter.标本号)
            Else
                varFilter(mFilter.标本号) = Replace(Replace(varFilter(mFilter.标本号), "～", "~"), "-", "~")
                If InStr(varFilter(mFilter.标本号), "~") > 0 Then
                    strStartNO = Split(varFilter(mFilter.标本号), "~")(0)
                    strEndNO = Split(varFilter(mFilter.标本号), "~")(1)
                    strSQL = strSQL & " And  标本序号  between [9] and [25]  and a.仪器id is not null "
                Else
                    strStartNO = varFilter(mFilter.标本号)
                    strSQL = strSQL & " And  标本序号 = [9] and a.仪器id is not null "
                End If
            End If
        End If

        If varFilter(mFilter.标识号) <> "" Then
            If IsNumeric(varFilter(mFilter.标识号)) Then
                strSQL = strSQL & " and (a.住院号 = [10] or a.门诊号 = [10]) "
            Else
                strSQL = strSQL & " and a.no = [10] "
            End If
        End If
        
        If varFilter(mFilter.检验类别) <> "" Then
            strSQL = strSQL & " And a.操作类型 = [11] "
        End If
        
        If varFilter(mFilter.检验人) <> "" Then
            strSQL = strSQL & " And a.检验人 like [12] "
        End If
        
        If InStr(1, varFilter(mFilter.检验项目), ",") > 1 Then
            If Mid(varFilter(mFilter.检验项目), InStr(1, varFilter(mFilter.检验项目), ",") + 1) = "True" Then
                strSQL = strSQL & " And g.诊疗项目id = [13] "
            Else
                strSQL = strSQL & " And g.检验项目ID = [13] "
            End If
        End If
        
        If varFilter(mFilter.送检科室) <> 0 Then
            strSQL = strSQL & " And a.申请科室ID = [14] "
        End If
        
        If varFilter(mFilter.送检人) <> "" Then
            strSQL = strSQL & " and a.申请人 = [15] "
        End If
        
        If varFilter(mFilter.检验仪器) <> 0 Then
            strSQL = strSQL & " and a.仪器ID = [17] "
        Else
            If InStr(mstrPrivs, "所有科室") = 0 Then
                strSQL = strSQL & " and a.仪器ID in (Select /*+cardinality(a,10)*/ * From Table(Cast(f_Num2list([24]) As zlTools.t_Numlist)) A) "
                strUserMachine = mstrMachineALL
            End If
        End If
        
        '组合查询
        If varFilter(mFilter.高级) <> "" And varFilter(mFilter.是否使用高级) = 1 Then
            varUnionFilter = Split(varFilter(mFilter.高级), ",")
            For intLoop = 0 To UBound(varUnionFilter)
                varItem = Split(varUnionFilter(intLoop), "^")
                
                If intLoop = 0 Then
                    strSQL = strSQL & " And ( g.检验项目Id = " & varItem(0) & _
                                IIf(IsNumeric(varItem(3)), " and zl_to_number(检验结果) " & varItem(2) & varItem(3), _
                                " and g.检验结果 " & varItem(2) & " '" & varItem(3) & "'")
                Else
                    strSQL = strSQL & " OR  g.检验项目Id = " & varItem(0) & _
                        IIf(IsNumeric(varItem(3)), " and zl_to_number(检验结果) " & varItem(2) & varItem(3), _
                            " and g.检验结果 " & varItem(2) & " '" & varItem(3) & "'")
                End If
                
                If varItem(3) <> "" And varItem(4) <> "" Then
                    strSQL = strSQL & " and g.检验项目Id = " & varItem(0) & _
                            IIf(IsNumeric(varItem(5)), " and zl_to_number(检验结果) " & varItem(4) & varItem(5), _
                            " and g.检验结果 " & varItem(4) & " '" & varItem(5) & "'")
                End If
            Next
            strSQL = strSQL & " )"
        End If
        
        If Val(varFilter(mFilter.病人ID)) <> 0 Then
            strSQL = strSQL & " And a.病人id = [18] "
        End If
        
        If Nvl(varFilter(mFilter.单据号)) <> "" Then
            strSQL = strSQL & " And a.no = [19] "
        End If
        
        If varFilter(mFilter.姓名) <> "" Or varFilter(mFilter.性别) <> "" Or varFilter(mFilter.年龄) <> "," Or _
           varFilter(mFilter.标识号) <> "" Or varFilter(mFilter.检验类别) <> "" Or InStr(1, varFilter(mFilter.检验项目), ",") <> 1 _
           Or varFilter(mFilter.送检科室) <> 0 Or varFilter(mFilter.送检人) <> "" Or Val(varFilter(mFilter.病人ID)) <> 0 Or _
           (varFilter(mFilter.高级) <> "" And varFilter(mFilter.是否使用高级) = 1) Or varFilter(mFilter.单据号) <> "" Then
           strSQL = strSQL & " And a.病人ID is not null "
        End If
        
        If Val(varFilter(mFilter.细菌)) <> 0 Then
            strSQL = strSQL & " And g.细菌ID = [21] "
        End If
        
        If Val(varFilter(mFilter.抗生素)) <> 0 Then
            strSQL = strSQL & " And O.抗生素ID = [22] "
        End If
        
        If varFilter(mFilter.药敏结果) <> "" Then
            strSQL = strSQL & " And O.结果类型  = [23] "
        End If
    Else
        '不使用过滤条件时的时间范围
        strStart = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0), 1)
        strEnd = GetDateTime(Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0), 2)
        
        If strStart = "自定义" Then
            strStart = Format(Me.dtpDate.Value, "yyyy-mm-dd 00:00:00")
            strEnd = Format(Me.dtpDateEnd.Value, "yyyy-mm-dd 23:59:59")
        Else
            If strStart = "" Then strStart = GetDateTime("今  天", 1)
            If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
        End If
        
        strSQL = strSQL & " And  a.核收时间 Between [3] And [4] "
        
        blnMoved = MovedByDate(CDate(Format(strStart, "yyyy-MM-dd hh:mm:ss")))
    End If
                  
    If rptList.Tag = "" Then
        strSQL = strSQL & _
              IIf(mlngMachineID <> 0, IIf(mlngMachineID = -1, "And a.仪器ID Is NULL ", "AND a.仪器ID = [1] "), "")
    End If
    
    '处理当前操作员可以操作的仪器
    If rptList.Tag = "" Then
        If mlngMachineID = 0 Then
            If mlngDeptID = 0 Then
                strUserMachine = ""
            Else
                For intLoop = 0 To Me.cboMachine.ListCount - 1
                    strUserMachine = strUserMachine & "," & Me.cboMachine.ItemData(intLoop)
                Next
                strUserMachine = Mid(strUserMachine, 2)
            End If
            If strUserMachine <> "" Then
                strSQL = strSQL & " and f.ID in (Select /*+cardinality(a,10)*/ * From Table(Cast(f_Num2list([24]) As zlTools.t_Numlist)) A) "
            End If
    
        End If
    End If

    
    If blnMoved Then
        strSQLbak = strSQL
        strSQLbak = Replace(strSQLbak, "'' As 转出", "'√' As 转出")
        
        strSQLbak = Replace(strSQLbak, "检验标本记录", "H检验标本记录")
        strSQLbak = Replace(strSQLbak, "检验普通结果", "H检验普通结果")
        strSQLbak = Replace(strSQLbak, "检验申请项目", "H检验申请项目")
        strSQL = strSQL & " Union ALL " & strSQLbak
    End If
    
    strSQL = strSQL & " ORDER BY 标本序号 "
    
    If cboUnionItem.ListCount > 0 Then
        lngUnionItem = Val(Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex))
    End If
    If rptList.Tag <> "" Then
        Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngMachineID, strDeptID, CDate(Format(strStart, "yyyy-MM-dd HH:mm:ss")), _
                     CDate(Format(strEnd, "yyyy-MM-dd HH:mm:ss")), "%" & CStr(varFilter(mFilter.姓名)) & "%", CStr(varFilter(mFilter.性别)), Val(intAgeBeging), Val(intAgeEnd), _
                     CStr(strStartNO), UCase(varFilter(mFilter.标识号)), CStr(varFilter(mFilter.检验类别)), CStr(varFilter(mFilter.检验人)) & "%", _
                     Mid(varFilter(mFilter.检验项目), 1, InStr(1, varFilter(mFilter.检验项目), ",") - 1), CStr(varFilter(mFilter.送检科室)), CStr(varFilter(mFilter.送检人)), _
                     lngUnionItem, CLng(varFilter(mFilter.检验仪器)), CLng(Val(varFilter(mFilter.病人ID))), zlCommFun.GetFullNO(CStr(varFilter(mFilter.单据号))), _
                     CStr(varFilter(mFilter.年龄单位)), Val(varFilter(mFilter.细菌)), Val(varFilter(mFilter.抗生素)), CStr(varFilter(mFilter.药敏结果)), strUserMachine, strEndNO)
    Else
        Set rsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngMachineID, strDeptID, CDate(Format(strStart, "yyyy-MM-dd HH:mm:ss")), _
                     CDate(Format(strEnd, "yyyy-MM-dd HH:mm:ss")), "", "", "", "", "", "", "", "", "", "", "", _
                     lngUnionItem, 0, 0, "", "", 0, 0, "", strUserMachine, strEndNO)
    End If
    
    '刷新前记录一下位置
    If Not Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row Is Nothing Then
        lngRowIndex = Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row.Index - 1
        lngRowID = Me.rptList.Rows(lngRowIndex).Record(mCol.ID).Value
        mlngLastShow = lngRowID
    Else
        If mlngLastShow > 0 Then
            For intLoop = 0 To Me.rptList.Rows.Count - 1
                If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = mlngLastShow Then
                    lngRowIndex = Me.rptList.Rows(intLoop).Record.Index
                    lngRowID = Me.rptList.Rows(intLoop).Record(mCol.ID).Value
                End If
            Next
        End If
    End If

    '按过滤条件只查询一次后清空
'    Me.rptList.Tag = ""
    Me.rptList.Records.DeleteAll
    blnPathPatient = False
    zlCommFun.ShowFlash "正在载入数据..."
    Do Until rsItem.EOF
        
        If lngSampleID <> rsItem("ID") Then
'            Me.stbThis.Panels(2).Text = "正在载入数据(" & lngLoop & ")"
            Set Record = Me.rptList.Records.Add
            
            For intLoop = 0 To Me.rptList.Columns.Count + 1
                Record.AddItem ""
            Next
            
            '前面几列需要处理图标
            Record.Item(mCol.紧急).Value = IIf(Nvl(rsItem("标本紧急")) = 1, "紧急", "")
            If Record.Item(mCol.紧急).Value = "紧急" Then
                Record.Item(mCol.紧急).Icon = 1
            Else
                Record.Item(mCol.紧急).Icon = -1
            End If
            
            Record.Item(mCol.紧急医嘱).Value = IIf(Nvl(rsItem("医嘱紧急")) = 1, "紧急", "")
            If Record.Item(mCol.紧急医嘱).Value = "紧急" Then
                Record.Item(mCol.紧急医嘱).Icon = 14
            Else
                Record.Item(mCol.紧急医嘱).Icon = -1
            End If
            
'            If Nvl(rsItem("执行状态")) = "已检验" Then
'                Record.Item(mCol.执行状态).Value = "已检验"
'                Record.Item(mCol.执行状态).Icon = 7
'            ElseIf CInt(Nvl(rsItem("打印次数"), "0")) > 0 Then
'                Record.Item(mCol.执行状态).Value = "已打印"
'                Record.Item(mCol.执行状态).Icon = 8
'            ElseIf Nvl(rsItem("传送")) = "" Then
'                Record.Item(mCol.执行状态).Value = "已传送"
'                Record.Item(mCol.执行状态).Icon = 6
'            End If
            
'            If Nvl(rsItem("初审人")) <> "" Then
'                Record.Item(mCol.查阅状态).Value = "已初审"
'                Record.Item(mCol.查阅状态).Icon = 13
'            End If

                    
            If Nvl(rsItem("查阅状态")) = 1 Then
                Record.Item(mCol.查阅状态).Value = "已查阅"
                Record.Item(mCol.查阅状态).Icon = 11
            End If
                            
                            
                            
            If CInt(Nvl(rsItem("打印次数"), "0")) > 0 Then
                Record.Item(mCol.执行状态).Value = "已打印"
                Record.Item(mCol.执行状态).Icon = 8
            ElseIf Nvl(rsItem("执行状态")) = "已检验" Then
                Record.Item(mCol.执行状态).Value = "已检验"
                Record.Item(mCol.执行状态).Icon = 7
            ElseIf Nvl(rsItem("初审人")) <> "" Then
                Record.Item(mCol.执行状态).Value = "初审"
                Record.Item(mCol.执行状态).Icon = 13
            ElseIf Nvl(rsItem("传送")) = "" Then
                Record.Item(mCol.执行状态).Value = "已传送"
                Record.Item(mCol.执行状态).Icon = 6
            Else
                Record.Item(mCol.执行状态).Value = ""
                Record.Item(mCol.执行状态).Icon = -1
            End If
            
            If Val(Nvl(rsItem("结果次数"))) > 0 Then
                Record.Item(mCol.复查).Icon = 10
            End If
            If rsItem("仪器审核") & "" = "仪器审核" Then
                Record.Item(mCol.仪器审核).Value = "√"
            Else
                Record.Item(mCol.仪器审核).Value = "×"
            End If
            If Val(Nvl(rsItem("临床路径病人"))) = 1 Then
                Record.Item(mCol.临床路径病人).Icon = 15
                blnPathPatient = True
            Else
                Record.Item(mCol.临床路径病人).Icon = -1
            End If
            
            Record.Item(mCol.姓名).Value = Nvl(rsItem("姓名")) '& IIf(Nvl(rsItem("婴儿"), 0) > 0, "(婴儿)", "")
            If Nvl(rsItem("标本类型")) = "质控" Then
                Record.Item(mCol.标本类型).Value = "质控"
                Record.Item(mCol.标本类型).Icon = 3
                strSQL = "Select A.标本id, B.名称, B.批号, B.水平 From 检验质控记录 A, 检验质控品 B Where A.质控品id = B.ID And A.标本id=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsItem("ID"))))
                Do Until rsTmp.EOF
                    Record.Item(mCol.姓名).Value = "" & rsTmp!名称 & "," & rsTmp!批号 & ",水平" & rsTmp!水平
                    rsTmp.MoveNext
                Loop
            ElseIf Nvl(rsItem("标本类型")) = "比对" Then
                Record.Item(mCol.标本类型).Value = "比对"
                Record.Item(mCol.标本类型).Icon = 4
                Record.Item(mCol.姓名).Value = Record.Item(mCol.姓名).Value & "(" & Nvl(rsItem("比对次数")) & ")"
            End If
            
            Record.Item(mCol.标本号).Value = Val(Nvl(rsItem("标本序号")))
            Record.Item(mCol.标本号).Caption = Trim(rsItem("标本号显示"))

            If Nvl(rsItem("年龄数字")) = "" Then
                
                If Nvl(rsItem("婴儿"), 0) = 0 Then
                    If IsNumeric(Nvl(rsItem("年龄"))) = True Then
                        Record.Item(mCol.年龄).Caption = Nvl(rsItem("年龄")) & "岁"
                    Else
                        If Nvl(rsItem("年龄")) <> "岁" And Nvl(rsItem("年龄")) <> "0岁" Then
                            Record.Item(mCol.年龄).Caption = Nvl(rsItem("年龄"))
                        End If
                    End If
                    If Record.Item(mCol.年龄).Caption <> "" Then
                        Record.Item(mCol.年龄).Value = Val(rsItem("年龄"))
                    End If
                End If
    '            Record.Item(mCol.年龄).Caption = IIf(Nvl(rsItem("婴儿"), 0) > 0, "", _
                                           IIf(Nvl(rsItem("年龄")) = "岁", "", _
                                           IIf(Nvl(rsItem("年龄")) = "0岁", "", IIf(IsNumeric(Nvl(rsItem("年龄"))) = True, rsItem("年龄") & "岁", rsItem("年龄")))))
            Else
                Record.Item(mCol.年龄).Value = Nvl(rsItem("年龄数字"))
                Record.Item(mCol.年龄).Caption = Nvl(rsItem("年龄")) '  Nvl(rsItem("年龄数字")) & Nvl(rsItem("年龄单位"))
            End If
            If Nvl(rsItem("病人类型")) <> "" Then
                Record.Item(mCol.姓名).ForeColor = zlDatabase.GetPatiColor(Nvl(rsItem("病人类型")), False)
            End If
            Record.Item(mCol.性别).Value = Nvl(rsItem("性别"))
            Record.Item(mCol.所属情况).Value = Nvl(rsItem("所属情况"))
            Record.Item(mCol.检验项目).Value = Trim(Nvl(rsItem("检验项目")))
            Record.Item(mCol.标识号).Value = Nvl(rsItem("标识号"))
            
            Record.Item(mCol.结果次数).Value = Nvl(rsItem("结果次数"))
            Record.Item(mCol.医嘱id).Value = Nvl(rsItem("医嘱ID"))
            Record.Item(mCol.仪器id).Value = Nvl(rsItem("仪器ID"))
            Record.Item(mCol.转出).Value = Nvl(rsItem("转出"))
            Record.Item(mCol.病人ID).Value = Nvl(rsItem("病人id"))
            Record.Item(mCol.ID).Value = Nvl(rsItem("ID"))
            Record.Item(mCol.标本时间).Caption = Format(Nvl(rsItem("核收时间")), "MM-dd HH:mm:ss")
            Record.Item(mCol.标本时间).Value = Format(Nvl(rsItem("核收时间")), "YYYY-MM-dd HH:mm:ss")
            Record.Item(mCol.报告时间).Caption = Format(Nvl(rsItem("检验时间")), "MM-dd HH:mm")
            Record.Item(mCol.报告时间).Value = Format(Nvl(rsItem("检验时间")), "YYYY-MM-dd HH:mm")
            Record.Item(mCol.微生物标本).Value = Val(Nvl(rsItem("微生物标本")))
    '        Record.Item(mCol.收费单).Value = Nvl(rsItem("收费单"))
            Record.Item(mCol.挂号单).Value = Nvl(rsItem("挂号单"))
            Record.Item(mCol.检验人).Value = Nvl(rsItem("检验人"))
            Record.Item(mCol.审核人).Value = Nvl(rsItem("审核人"))
            Record.Item(mCol.病人科室).Value = Nvl(rsItem("病人科室"))
            Record.Item(mCol.样本条码).Value = Nvl(rsItem("样本条码"))
            'Record.Item(mCol.发送号).Value = Nvl(rsItem("发送号"))
            Record.Item(mCol.婴儿).Value = Nvl(rsItem("婴儿"))
            Record.Item(mCol.仪器名).Value = Nvl(rsItem("仪器名称"))
            Record.Item(mCol.主页ID).Value = Nvl(rsItem("主页ID"))
            Record.Item(mCol.开嘱科室ID).Value = Nvl(rsItem("开嘱科室Id"))
            Record.Item(mCol.报告结果).Value = Nvl(rsItem("报告结果"))
            Record.Item(mCol.年龄数字).Value = Nvl(rsItem("年龄数字"))
            Record.Item(mCol.年龄单位).Value = Nvl(rsItem("年龄单位"))
            Record.Item(mCol.床号).Value = Nvl(rsItem("床号"))
            Record.Item(mCol.申请人).Value = Nvl(rsItem("申请人"))
            Record.Item(mCol.标本形态).Value = Nvl(rsItem("标本形态"))
            Record.Item(mCol.采样人).Value = Nvl(rsItem("采样人"))
            Record.Item(mCol.采样时间).Value = Nvl(rsItem("采样时间"))
            Record.Item(mCol.检验标本).Value = Nvl(rsItem("检验标本"))
            Record.Item(mCol.NO).Value = Nvl(rsItem("NO"))
            Record.Item(mCol.接收人).Value = Nvl(rsItem("接收人"))
            Record.Item(mCol.接收时间).Value = Nvl(rsItem("接收时间"))
            Record.Item(mCol.审核时间).Value = Nvl(rsItem("审核时间"))
            Record.Item(mCol.病区名称).Value = Nvl(rsItem("病区名称"))
            Record.Item(mCol.执行科室ID).Value = Nvl(rsItem("执行科室ID"))
            Record.Item(mCol.标本类别).Value = Nvl(rsItem("标本类别"))
            Record.Item(mCol.医嘱紧急).Value = Nvl(rsItem("医嘱紧急"))
            Record.Item(mCol.标本紧急).Value = Nvl(rsItem("标本紧急"))
            Record.Item(mCol.申请科室).Value = Nvl(rsItem("申请科室"))
            Record.Item(mCol.申请类型).Value = Nvl(rsItem("申请类型"), 0)
            Record.Item(mCol.报告发送).Value = Nvl(rsItem("报告发送"), 0)
            Record.Item(mCol.病人科室ID).Value = Nvl(rsItem("病人科室ID"), 0)
            Record.Item(mCol.初审人).Value = Nvl(rsItem("初审人"))
            Record.Item(mCol.初审时间).Value = Nvl(rsItem("初审时间"))
            Record.Item(mCol.健康号).Value = Nvl(rsItem("健康号"))
            Record.Item(mCol.审核未通过).Value = Nvl(rsItem("审核未通过"))
            Record.Item(mCol.病人来源).Value = Nvl(rsItem("病人来源"))
            Record.Item(mCol.门诊号).Value = Nvl(rsItem("门诊号"))
            Record.Item(mCol.住院号).Value = Nvl(rsItem("住院号"))
            If Nvl(rsItem("项目")) = "任务团体" Then
                Record.Item(mCol.单位).Value = Nvl(rsItem("内容"))
            End If
            Record.Item(mCol.结果为空).Value = Val(Nvl(rsItem("结果为空")))
            
            
'            Record.Item(mCol.查阅状态).Value = Nvl(rsItem("查阅状态"), 0)

            
            '------晋煤新增
            For i = 0 To rptList.Columns.Count + 1
                If Val("" & rsItem!微生物标本) = 0 Then
                    If Record.Item(mCol.结果为空).Value > 0 Then
                        Record.Item(i).BackColor = vbWhite
                    Else
                        Record.Item(i).BackColor = &HFDD6C6
                    End If
                Else
                    Record.Item(i).BackColor = vbWhite
                End If
            Next
            
            lngloop = lngloop + 1
            If mintLoadShow > 0 Then
                DoEvents
            End If
        End If
        lngSampleID = rsItem("ID")
        rsItem.MoveNext
        If lngloop = 10000 Then
            MsgBox "你选择的条件范围过大，已超过10000条记录了" & vbCrLf & _
                   " 请重新选择条件进行查找!", vbQuestion, Me.Caption
            Call SetControlFocus
            gintSelectFocus = 1
            Exit Do
        End If
    Loop
    
    '没有临床路径病人时不显示列
    Me.rptList.Columns(6).Visible = blnPathPatient
    
    zlCommFun.StopFlash
'    If Me.TabList.Selected.Index = 0 Then
'        Me.rptList.SetFocus
'    Else
'        Me.rptList1.SetFocus
'    End If
    Me.rptList.Populate
    Me.MousePointer = 0
    
    If mintLoadShow = 0 Then
        mintLoadShow = mintLoadShow + 1
        Exit Function
    End If
    
    '过滤界面列表
    RptListFilter
    Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptList.Rows.Count & "个病人。"
    
    
    
    
    '重新定位到以前的位置
    If rptList.Rows.Count > 0 And lngRowIndex > 0 Then
'        Me.rptList.Rows(0).Selected = True
'        Me.rptList.Rows(0).EnsureVisible
        lngloop = 0

        For intLoop = 0 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = lngRowID Then
                lngloop = Me.rptList.Rows(intLoop).Index
                Exit For
            End If
        Next

        If lngRowIndex >= lngloop Then
            lngRowIndex = lngRowIndex - (lngRowIndex - lngloop)
        Else
            lngRowIndex = lngRowIndex + (lngloop - lngRowIndex)
        End If
        Me.rptList.Rows(lngRowIndex).EnsureVisible
    End If
    
        
    For Each Rerow In Me.rptList.Rows
        If Rerow.Record(mCol.ID).Value = mlngKey Then
            Rerow.Selected = True
            mlngKey = Rerow.Record(mCol.ID).Value
            Set Me.rptList.FocusedRow = Rerow
            Me.rptList.Populate
            Exit Function
        End If
    Next
    
    If Me.rptList.Rows.Count > 0 Then
        If lngRow <= Me.rptList.Rows.Count And lngRow > 0 Then
            Set Me.rptList.FocusedRow = rptList.Rows(lngRow - 1)
            mlngKey = rptList.Rows(lngRow - 1).Record(mCol.ID).Value
        Else
            Set Me.rptList.FocusedRow = rptList.Rows(0)
            mlngKey = rptList.Rows(0).Record(mCol.ID).Value
        End If
        Me.rptList.Populate
        Exit Function
    Else
        mlngKey = 0
    End If
    
    
    '刷新列表
    If Not Me.rptList.FocusedRow Is Nothing Then
        Call mfrmRequest.zlRefresh(Me.rptList.FocusedRow)
        Call RefreshTableWindow(Me.TabCtlWindow.Selected.Index)
        If mlngKey <> 0 Then
            ReadImageData mlngKey, False
        End If
    End If
    Exit Function
errH:
'    Me.rptList.SetFocus
    zlCommFun.StopFlash
    Me.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub CreaterptListHead()
    Dim Column As ReportColumn
    Dim i As Integer
    With Me.rptList1.Columns
        
        rptList1.AllowColumnRemove = False
        rptList1.ShowItemsInGroups = False
        
        With rptList1.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptList1.SetImageList imgList
        Set Column = .Add(mCol.紧急, "紧急", 18, False)
        Column.Icon = 0
        Set Column = .Add(mRCol.病人ID, "病人ID", 45, False): Column.Visible = False: Column.ShowInFieldChooser = False
'        column.
        Set Column = .Add(mRCol.来源, "来源", 55, True)
        Set Column = .Add(mRCol.姓名, "姓名", 55, True)
        Set Column = .Add(mRCol.性别, "性别", 55, True)
        Set Column = .Add(mRCol.年龄, "年龄", 55, True)
        Set Column = .Add(mRCol.病人科室, "病人科室", 75, True)
        Set Column = .Add(mRCol.标识号, "标识号", 65, True)
        Set Column = .Add(mRCol.床号, "床号", 65, True)
        Set Column = .Add(mRCol.医嘱内容, "医嘱内容", 75, True)
        Set Column = .Add(mRCol.开嘱医生, "开嘱医生", 75, True)
        Set Column = .Add(mRCol.开嘱时间, "开嘱时间", 75, True)
        Set Column = .Add(mRCol.签收时间, "签收时间", 75, True)
        Set Column = .Add(mRCol.诊疗项目ID, "诊疗项目ID", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mRCol.执行状态, "执行状态", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mRCol.定位, "定位", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mRCol.挂号单, "挂号单", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
    End With
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
        rptList.SetImageList imgList
        
        Set Column = .Add(mCol.紧急, "紧急", 18, False):   Column.Icon = 0
        Set Column = .Add(mCol.紧急医嘱, "紧急医嘱", 18, False): Column.Icon = 14
        Set Column = .Add(mCol.执行状态, "执行状态", 18, False): Column.Icon = 5
        Set Column = .Add(mCol.标本类型, "标本类型", 18, False): Column.Icon = 2
        Set Column = .Add(mCol.复查, "复查", 18, False): Column.Icon = 9
        Set Column = .Add(mCol.查阅状态, "查阅状态", 18, False): Column.Icon = 11
        Set Column = .Add(mCol.临床路径病人, "临床路径病人", 18, False): Column.Icon = 15
        Set Column = .Add(mCol.仪器审核, "仪器审核", 18, False): Column.Icon = 16
        
        Set Column = .Add(mCol.标本号, "标本号", 65, True)
        Column.SortAscending = zlDatabase.GetPara("标本号排序", 100, 1208, 0)
        Column.Sortable = True:  Me.rptList.SortOrder.Add Column
        Set Column = .Add(mCol.姓名, "姓名", 45, True)
        Set Column = .Add(mCol.性别, "性别", 40, True)
        Set Column = .Add(mCol.年龄, "年龄", 40, True)
        Set Column = .Add(mCol.标本时间, "核收时间", 80, True)
        Set Column = .Add(mCol.报告时间, "报告时间", 80, True)
        Set Column = .Add(mCol.检验项目, "检验项目", 90, True)
        Set Column = .Add(mCol.所属情况, "来源", 40, False)
        Set Column = .Add(mCol.标识号, "标识号", 55, True)
        Set Column = .Add(mCol.样本条码, "样本条码", 75, True)
        Set Column = .Add(mCol.健康号, "健康号", 75, True)
        Set Column = .Add(mCol.单位, "单位", 75, True)
        Set Column = .Add(mCol.病人科室, "病人科室", 80, True) ': Column.Visible = False
        Set Column = .Add(mCol.结果次数, "结果次数", 65, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.医嘱id, "医嘱ID", 65, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.仪器id, "仪器ID", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.转出, "转出", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.病人ID, "病人ID", 75, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.ID, "ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.微生物标本, "微生物标本", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.收费单, "收费单", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.挂号单, "挂号单", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.检验人, "检验人", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.审核人, "审核人", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.发送号, "发送号", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.婴儿, "婴儿", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.开嘱科室ID, "开嘱科室ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.主页ID, "主页ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.报告结果, "报告结果", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.年龄数字, "年龄数字", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.年龄单位, "年龄单位", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.床号, "床号", 30, True) ': Column.Visible = False
        Set Column = .Add(mCol.申请人, "申请人", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.标本形态, "标本形态", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.采样人, "采样人", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.采样时间, "采样时间", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.检验标本, "检验标本", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.NO, "NO", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.接收人, "接收人", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.接收时间, "接收时间", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.审核时间, "审核时间", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.病区id, "病区ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.病区名称, "病区", 30, True) ': Column.Visible = False
        Set Column = .Add(mCol.定位, "定位", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.执行科室ID, "执行科室ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.标本类别, "标本类别", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.医嘱紧急, "医嘱紧急", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.标本紧急, "标本紧急", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.申请科室, "申请科室", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.申请类型, "申请类型", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.报告发送, "报告发送", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.病人科室ID, "病人科室ID", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.初审人, "初审人", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.初审时间, "初审时间", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.审核未通过, "审核未通过", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.病人来源, "病人来源", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.门诊号, "门诊号", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.住院号, "住院号", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
        Set Column = .Add(mCol.结果为空, "结果为空", 30, True): Column.Visible = False: Column.ShowInFieldChooser = False
    End With
End Sub

Private Sub SampleDisposal(Disposal As Integer)
    '功能           处理对样本的各种操作
    '               样本号修改、删除无主标本
    Dim strSQL As String                                    '临时SQL语句
    Dim intExeState As Integer                              '执行状态( 7=已检验 8=检验后打印 6=已发送数据给仪器
    Dim strSamptleType As String                            '标本所属情况(院内、院外、外来、无主)
    Dim strSamptleKind  As Integer                          '标本类型(=3质控 =4比对)
    Dim strPatienName As String                             '病人姓名
    Dim lngMachineID As Long                                '仪器ID
    Dim strSampltDate As String                             '标本时间
    Dim strSampltID As String                               '标本号
    Dim blEmergent  As Boolean                              '是否紧急
    Dim lngAdvice As Long                                   '医嘱ID
    Dim lngRetuId As Long                                   '调用模块返回的ID
    Dim rsTmp As New ADODB.Recordset                        '数据集
    Dim rs As New ADODB.Recordset
    Dim intMicrobe As Integer                               '是否是微生物
    Dim strVerifyMan As String                              '检验人
    Dim lngPatientID As Long                                '病人ID
    Dim strDevices As String                                '设备ID集
    Dim strAdviceIDs As String                              '医嘱ID集
    Dim aDevice() As String                                 '设备S
    Dim intLoop As Integer
    Dim astrSQL() As String                                 'SQL数组
    Dim intEmerge As Integer                                '是否使用急诊标志
    Dim lngSampleID As Long                                 '标本ID
    Dim lngBeginDate As Long
    Dim strStartDate As String
    Dim strEndDate As String
    Dim blnRollBak As Boolean                               '回退标志
    Dim strEmergency As String                              '标本类别
    Dim str初审人 As String                                  '初审人
    Dim strAdviceIDall As String                            '医嘱id ,更新新版LIS申请 状态
    
    Dim bln发送杯号 As Boolean                              '是否发送杯号
    Dim intRow As Integer, strIDList() As String              '保存标本ID
    Dim strNoSend As String, lngCount As Long
    On Error GoTo errH
    
    intEmerge = Val(zlDatabase.GetPara("急诊标本", 100, 1208, 0))
    
    If Not rptList.FocusedRow Is Nothing Then                                   '没有焦点行时退出
        With Me.rptList.FocusedRow
            intExeState = .Record(mCol.执行状态).Icon
            strSamptleType = .Record(mCol.所属情况).Value
            strPatienName = .Record(mCol.姓名).Value
            lngMachineID = Val(.Record(mCol.仪器id).Value)
            strSampltDate = .Record(mCol.标本时间).Value
            strSampltID = .Record(mCol.标本号).Value
            blEmergent = IIf(.Record(mCol.标本类别).Value = "1", True, False)
            lngAdvice = Val(.Record(mCol.医嘱id).Value)
            strSamptleKind = .Record(mCol.标本类型).Icon
            intMicrobe = Val(.Record(mCol.微生物标本).Value)
            strVerifyMan = .Record(mCol.检验人).Value
            lngSampleID = .Record(mCol.ID).Value
            strEmergency = .Record(mCol.标本类别).Value
            str初审人 = .Record(mCol.初审人).Value
        End With
    End If
    
    '得到病人Id(等处理列表)
    If Me.TabList.Item(1).Selected = True Then
        If Not Me.rptList1.FocusedRow Is Nothing Then
            With Me.rptList1.FocusedRow
                lngPatientID = Val(.Record(mRCol.病人ID).Value)
            End With
        End If
    End If
    
    Select Case Disposal
        Case mActS.修改样本号                                                           '修改样本号
            
            Dim strNewNo As String, str标本形态 As String, str标本类型 As String, str姓名 As String
            '已检验项目不能进行修改样本号操作
            If intExeState = 7 Or intExeState = 8 Then Exit Sub
            
            frmLisStationModifyNo.ShowEdit Me, mlngKey, strNewNo, str标本形态, str标本类型, strEmergency, str姓名
            If strNewNo = "" Then Exit Sub
            '判断标本是否存在
            strStartDate = GetDateTime(mMakeNoRule, 1, strSampltDate)
            strEndDate = GetDateTime(mMakeNoRule, 2, strSampltDate)
            gstrSql = "Select ID" & vbNewLine & _
                    " From 检验标本记录 A" & vbNewLine & _
                    " Where 核收时间 Between [1] And [2] And 标本序号 = [3] And Nvl(标本类别, 0) = [4] And ID <> [5] " & vbNewLine & _
                    "       And 仪器ID = [6] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strStartDate), CDate(strEndDate), _
                    TransSampleNO(strNewNo), IIf(strEmergency, 1, 0), mlngKey, mlngMachineID)
                    
            If rsTmp.EOF = False Then
                MsgBox "标本号<" & strNewNo & ">已存在！", vbInformation, Me.Caption
                Call SetControlFocus
                Exit Sub        '找到相同时退出
            End If
            If strNewNo <> "" Then
                strSQL = "ZL_检验标本记录_标本序号(" & _
                         mlngKey & ",'" & strNewNo & "','" & str标本形态 & "','" & str标本类型 & "',NULL,NULL,'" & strEmergency & "','" & str姓名 & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
'            InsertOneRecored mlngKey, False
            RefreshData
            gintSelectFocus = 1
'            RefreshData
        Case mActS.批量修改样本号                                                       '批量修改样本号
            
            Call frmBatchAction.ShowMe(Me, 4, mlngMachineID, , , , , mlngDeptID, mstrAuditingManID)
            gintSelectFocus = 1
        Case mActS.删除无主标本                                                         '删除无主标本
            
            If InStr(";" & mstrPrivs & ";", ";删除无主标本;") = 0 Then
                MsgBox "你没有删除无主标本的权限，请与管理系统!", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If strSamptleType = "无主" Then
                If MsgBox("真的要删除无主标本吗？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                End If
                strSQL = "ZL_检验标本记录_标本删除(" & mlngKey & ")"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            If strSamptleType = "院外" Then
                '取消核收
                If MsgBox("真的要删除“" & strPatienName & "”院外标本吗？", _
                            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                End If
                Call SampleRefuse(mlngKey)                               '取消核收
                '删除无主
                strSQL = "ZL_检验标本记录_标本删除(" & mlngKey & ")"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            intLoop = Me.rptList.FocusedRow.Index
            DelItem lngSampleID
            On Error Resume Next
            If Me.rptList.Rows.Count > 0 Then
                If Me.rptList.Rows.Count < intLoop Then
                    Me.rptList.FocusedRow = Me.rptList.Rows(Me.rptList.Rows.Count)
                Else
                    If intLoop = 0 Then intLoop = 1
                    Me.rptList.FocusedRow = Me.rptList.Rows(intLoop - 1)
                End If
            End If
            gintSelectFocus = 1
'            With Me.rptList.FocusedRow
'                .Record.DeleteAll
'            End With
'            Me.rptList.Populate
'            RefreshData
'            Me.rptList.SetFocus
        Case mActS.发送仪器                                                             '发送仪器
'            If mbln手工发送杯号 Then
'                'strSQL = "Select nvl(发送时指定杯号,0) as 发送 From 检验仪器 Where id=[1]"
'                'Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngMachineID)
'                   'Do Until rs.EOF
'                     bln发送杯号 = Val("" & rs!发送) = 1
'                     rs.MoveNext
'                    Loop
'                End If
                If blnComm = False Then Exit Sub        '不需要时直接退出
                
282             SendSample WinsockC, WinsockC.LocalIP, lngMachineID, strSampltDate, strSampltID, "", False, IIf(blEmergent And intEmerge = 1, 1, 0)
                        '成功
284             If blnComm And Not Me.rptList.FocusedRow Is Nothing Then
286                 Me.rptList.FocusedRow.Record(mCol.执行状态).Value = "已传送"
288                 Me.rptList.FocusedRow.Record(mCol.执行状态).Icon = 6
290                 Me.rptList.Populate
                End If
        Case mActS.批量发送到仪器
                
244             'If blnComm = False Then Exit Sub
                Call RefreshData '刷新状态后再发送
                Me.MousePointer = vbHourglass
                    '2013-11-26 支持超过1000个以上的标本发送
246             strNoSend = "": lngCount = 0
248             ReDim strIDList(0) As String
250             intRow = -1
252             If rptList.Rows.Count > 0 Then
254                 For intRow = 0 To rptList.Rows.Count - 1
256                     With rptList
258                         If Not .Rows(intRow).GroupRow Then
                                '7-已检验 8-已打印 13-初审
                            
260                             If InStr(",7,8,13,", CStr(.Rows(intRow).Record(mCol.执行状态).Icon)) <= 0 Then
262                                 strIDList(UBound(strIDList)) = strIDList(UBound(strIDList)) & "," & .Rows(intRow).Record(mCol.ID).Value
264                                 If Len(strIDList(UBound(strIDList))) > 3000 Then ReDim Preserve strIDList(UBound(strIDList) + 1)
                                Else
268                                 strNoSend = strNoSend & vbNewLine & .Rows(intRow).Record(mCol.标本号).Value & " " & .Rows(intRow).Record(mCol.姓名).Value & " 状态=" & CStr(.Rows(intRow).Record(mCol.执行状态).Icon)
270                                 lngCount = lngCount + 1
                                End If
                            End If
                        End With
                    Next
                End If
272             If intRow >= 0 Then Call frmLabMainSendSample.ShowMe(strIDList(), Me)
274             If strNoSend <> "" Then
276                 stbThis.Panels(2).Text = "有" & lngCount & "个标本，因已审核，本次不发往仪器"
'278                 WriteToLog "未传到发送界面的标本有：" & strNoSend
                End If

292             Me.MousePointer = vbDefault
                '连继输入时不刷新
    '            If Me.rptList.Tag <> "Continue" Then
    '                RefreshData
    '            End If
            
        Case mActS.置为质控                                                             '置为质控
            
            '增加
            frmLabMainSetQC.ShowMe Me, mlngKey, strSampltID, lngMachineID, strVerifyMan, 1
            
            InsertOneRecored mlngKey, False
'            If MsgBox("真的要将无主标本转为质控标本吗？", _
'                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'            strSql = "ZL_检验标本记录_标本质控(" & mlngKey & ",1)"
'            zldatabase.ExecuteProcedure strSql, Me.Caption
            
'            RefreshData
            gintSelectFocus = 1
            InsertOneRecored mlngKey, False
            
        Case mActS.置为对比                                                             '置为对比
            
            frmLabToCompare.ShowMe Me, mlngKey
            InsertOneRecored mlngKey, False
'            RefreshData
            gintSelectFocus = 1
        Case mActS.状态回滚                                                             '状态回滚
            '是否只能回滚自已的标本
            If InStr(1, mstrPrivs, "修改他人结果") <= 0 And UserInfo.姓名 <> strVerifyMan And strPatienName <> "" Then
                MsgBox "你不能回滚他人的报告单！", vbInformation, Me.Caption
                Call SetControlFocus
                Exit Sub
            End If
            
            If intExeState = 7 Or intExeState = 8 Or intExeState = 11 Then
                '审核后状态
                If strSamptleKind = 4 Then
                    '取消比对
                    If MsgBox("真的要将“" & strPatienName & "”比对标本转为普通标本吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                    End If
                    strSQL = "ZL_检验标本记录_标本质控(" & mlngKey & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    InsertOneRecored mlngKey, False
                Else
                    '回滚审核
                    If InStr(1, ";" & mstrPrivs & ";", ";审核取消;") > 0 Or InStr(1, ";" & mstrPrivs & ";", ";24小时审核取消;") > 0 Then
                        Call ReportDisposal(mActR.审核取消)
                    Else
                        MsgBox "您没有审核取消的权限!", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            Else
                '审核前
                If strSamptleKind = 4 Then
                    '取消比对
                    If MsgBox("真的要将“" & strPatienName & "”比对标本转为普通标本吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call SetControlFocus
                        gintSelectFocus = 1: Exit Sub
                    End If
                    strSQL = "ZL_检验标本记录_标本质控(" & mlngKey & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    InsertOneRecored mlngKey, False
                ElseIf strSamptleKind = 3 Then
                    '取消质控
                    If MsgBox("真的要将质控标本转为普通标本吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call SetControlFocus
                        gintSelectFocus = 1: Exit Sub
                    End If
'                    frmLabMainSetQC.ShowMe Me, mlngKey, strSampltID, lngMachineID, strVerifyMan, 3
                    strSQL = "ZL_检验质控记录_EDIT(3," & mlngKey & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    Me.rptList.FocusedRow.Record(mCol.姓名).Value = ""
                    Me.rptList.FocusedRow.Record(mCol.标本类型).Value = ""
                    Me.rptList.FocusedRow.Record(mCol.标本类型).Icon = -1
                    Me.rptList.Populate
                ElseIf mSendReport = 1 And str初审人 <> "" Then
                    gstrSql = "Zl_检验标本记录_初审报告(" & mlngKey & ",2,'" & UserInfo.姓名 & "')"
                    zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                    InsertOneRecored mlngKey, False
                Else
                    If strPatienName <> "" Then
                        If MsgBox("是否确定要置为无主标本?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                        End If
                        strSQL = "Select Distinct 医嘱ID From (Select 医嘱ID From 检验项目分布 Where 标本id = [1] " & _
                                "Union All Select 医嘱ID From 检验标本记录 Where ID = [1])"
                        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)

                        intLoop = 0
                        Do While Not rsTmp.EOF
                            If Not IsNull(rsTmp(0)) And Val(Nvl(rsTmp(0))) <> 0 Then
                                '处理双向通信
                                If blnComm Then
                                    strAdviceIDs = strAdviceIDs & "," & rsTmp(0)
                                    gstrSql = "Select Distinct 仪器ID From 检验标本记录 A,检验项目分布 B " & _
                                        " Where B.医嘱ID=[1] And B.标本ID+0=A.ID"
                                    Set rs = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(rsTmp(0)))

                                    Do While Not rs.EOF
                                        If InStr(strDevices, "," & zlCommFun.Nvl(rs(0), 0)) = 0 Then
                                            strDevices = strDevices & "," & zlCommFun.Nvl(rs(0), 0)
                                        End If
                                        rs.MoveNext
                                    Loop
                                End If
                                If intLoop = 0 Then
                                    ReDim Preserve astrSQL(1 To 1)
                                    astrSQL(1) = "ZL_检验标本记录_转为无主(" & rsTmp(0) & ")"
                                Else
                                    ReDim Preserve astrSQL(1 To UBound(astrSQL) + 1)
                                    astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_转为无主(" & rsTmp(0) & ")"
'                                    aStrSQL(ReDimArray(aStrSQL)) = "ZL_检验标本记录_转为无主(" & rsTmp(0) & ")"
                                End If
                                strAdviceIDall = strAdviceIDall & "," & rsTmp(0)
'                                zldatabase.ExecuteProcedure "ZL_检验标本记录_转为无主(" & rsTmp(0) & ")", gstrSysName
                                intLoop = intLoop + 1
                            End If
                            rsTmp.MoveNext
                        Loop
                        
                        If intLoop > 0 Then
                            '处理双向通信
                            If blnComm Then
                                If Len(strDevices) > 0 Then strDevices = Mid(strDevices, 2)
                                If Len(strAdviceIDs) > 0 Then strAdviceIDs = Mid(strAdviceIDs, 2)
                                aDevice = Split(strDevices, ",")
                                mblnSendComplete = False
                                For intLoop = 0 To UBound(aDevice)
                                    SendSample WinsockC, WinsockC.LocalIP, CLng(Val(aDevice(intLoop))), "", 0, strAdviceIDs, True, IIf(blEmergent And intEmerge = 1, 1, 0)
                                Next
                                lngBeginDate = Timer
                                Do
                                    DoEvents
                                Loop Until mblnSendComplete = True Or (CLng(Timer) - lngBeginDate > 2)
                            End If
                            gcnOracle.BeginTrans
                            blnRollBak = True
                            For intLoop = 1 To UBound(astrSQL)
                                If astrSQL(intLoop) <> "" Then Call zlDatabase.ExecuteProcedure(astrSQL(intLoop), Me.Caption)
                            Next
                            gcnOracle.CommitTrans
                        Else
                            gstrSql = "Zl_检验标本记录_置为无主(" & mlngKey & ")"
                            zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                        End If
                        If strAdviceIDall <> "" Then
                            ModifyApplyToLIS strAdviceIDall, 0
                        End If
                        Call RefreshData
'                        InsertOneRecored mlngkey, False
                    Else
                        If InStr(";" & mstrPrivs & ";", ";删除无主标本;") = 0 Then
                            MsgBox "你没有删除无主标本的权限，请与管理系统!", vbInformation, Me.Caption
                            Exit Sub
                        End If
                        
                        If MsgBox("是否确定要删除无主标本?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                            Call SetControlFocus
                            gintSelectFocus = 1: Exit Sub
                        End If
                            strSQL = "ZL_检验标本记录_标本删除(" & mlngKey & ")"
                            zlDatabase.ExecuteProcedure strSQL, Me.Caption
                        intLoop = Me.rptList.FocusedRow.Index
                        DelItem lngSampleID
                        On Error Resume Next
                        If Me.rptList.Rows.Count > 0 Then
                            If Me.rptList.Rows.Count < intLoop Then
                                Me.rptList.FocusedRow = Me.rptList.Rows(Me.rptList.Rows.Count)
                            Else
                                If intLoop = 0 Then intLoop = 1
                                Me.rptList.FocusedRow = Me.rptList.Rows(intLoop - 1)
                            End If
                        End If
'                        Me.rptList.FocusedRow.Record.DeleteAll
'                        Me.rptList.Populate
                    End If
                    '取消核收
'                    Call frmLisStationCheckCancel.ShowEdit(Me, mlngKey, objLISComm)
                End If
            End If
            gintSelectFocus = 1
            RptListFilter
'            RefreshData
        Case mActS.拒收                                                                 '拒收
            
            If mintEditState <> 0 Then
                lngRetuId = mfrmRequest.ZlRefuse()
                If lngRetuId = 0 Then Exit Sub
                If mintContinue = 0 Then
                    '不连续操作
                    mintEditState = 0
                End If
            Else
                If rptList1.FocusedRow Is Nothing Then Exit Sub
                frmLabRefuse.ShowEdit Me, rptList1.FocusedRow.Record(mRCol.医嘱id).Value, Me.WinsockC
                RefreshData1
            End If
            gintSelectFocus = 1
        Case mActS.核收                                                                 '核收
            '在保前如果有保存后立即审核检查是否为同一个人审核
            If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                If AuditionCheck = False Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再核收.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            '由于回车时焦点会跑到"检验备注"处特使用下面方法解决原因不明(待查)
'            If intMicrobe <> 1 Then
'                Me.TabCtlWindow.Item(0).Selected = True
''                mfrmWrite.txtComment.SetFocus
'            Else
'                Me.TabCtlWindow.Item(1).Selected = True
'                SetActiveWindow Me.TabCtlWindow.Item(1).Handle
''                mfrmWrite2.txtComment.SetFocus
'            End If
            Call ShowRequest(True)
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            mintEditState = 1
            If mfrmRequest.ZlEditStart(mActS.核收, mlngDeptID, mlngMachineID, 0, 0, 0, _
                                        IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan), _
                                        Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex), _
                                        lngPatientID) = False Then
                mfrmWrite.ZlClearForm
                mfrmWrite2.ZlClearForm
                mintEditState = 0
            End If
            
            
        
        Case mActS.登记                                                                 '登记
            '在保前如果有保存后立即审核检查是否为同一个人审核
            If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                If AuditionCheck = False Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再登记.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            '由于回车时焦点会跑到"检验备注"处特使用下面方法解决原因不明(待查)
'            If intMicrobe <> 1 Then
'                Me.TabCtlWindow.Item(0).Selected = True
'                mfrmWrite.txtComment.SetFocus
'            Else
'                Me.TabCtlWindow.Item(1).Selected = True
'                SetActiveWindow Me.TabCtlWindow.Item(1).Handle
'                mfrmWrite2.txtComment.SetFocus
'            End If
            Call ShowRequest(True)
            mintEditState = 2
            Me.dkpMain.FindPane(Dkp_ID_Request).Select
            
            If mfrmRequest.ZlEditStart(mActS.登记, mlngDeptID, mlngMachineID, 0, 0, 0, _
                                        IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan), _
                                        Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex)) = False Then
                mfrmWrite.ZlClearForm
                mfrmWrite2.ZlClearForm
                mintEditState = 0
            End If
        Case mActS.批量增加                                                             '批量增加
            
            If frmAddSample.ShowEdit(Me, "", mlngDeptID, mlngMachineID) = True Then
                
                Call RefreshData
                gintSelectFocus = 1
            End If
        
        Case mActS.补填病人                                                             '补填病人
            If intExeState = 7 Or intExeState = 8 Then Exit Sub
            If str初审人 <> "" And mSendReport = 1 Then Exit Sub
            If strPatienName <> "" Then
                '是否只能回滚自已的标本
                If InStr(1, mstrPrivs, "修改他人结果") <= 0 And UserInfo.姓名 <> strVerifyMan Then
                    MsgBox "你不能回滚他人的报告单！", vbInformation, Me.Caption
                    Call SetControlFocus
                    Exit Sub
                End If
            End If
            
            '在保前如果有保存后立即审核检查是否为同一个人审核
            If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                If AuditionCheck = False Then
                    '审核人登陆
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再补填病人.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            '由于回车时焦点会跑到"检验备注"处特使用下面方法解决原因不明(待查)
'            If intMicrobe <> 1 Then
'                Me.TabCtlWindow.Item(0).Selected = True
'                mfrmWrite.txtComment.SetFocus
'            Else
'                Me.TabCtlWindow.Item(1).Selected = True
'                SetActiveWindow Me.TabCtlWindow.Item(1).Handle
'                mfrmWrite2.txtComment.SetFocus
'            End If
            Call ShowRequest(True)
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            mintEditState = 4
            If mfrmRequest.ZlEditStart(mActS.补填病人, mlngDeptID, mlngMachineID, mlngKey, 0, 0, _
                                    IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan), _
                                    Me.cboUnionItem.ItemData(Me.cboUnionItem.ListIndex)) = False Then
                mintEditState = 0
            End If
            
        Case mActS.重新核收                                                             '重新核收
        
            '由于回车时焦点会跑到"检验备注"处特使用下面方法解决原因不明(待查)
'            If intMicrobe <> 1 Then
'                Me.TabCtlWindow.Item(0).Selected = True
'                mfrmWrite.txtComment.SetFocus
'            Else
'                Me.TabCtlWindow.Item(1).Selected = True
'                SetActiveWindow Me.TabCtlWindow.Item(1).Handle
'                mfrmWrite2.txtComment.SetFocus
'            End If
            Me.dkpMain.FindPane(Dkp_ID_Append).Select
            Me.dkpMain.FindPane(Dkp_ID_Request).Select
            mintEditState = 3
            If mfrmRequest.ZlEditStart(mActS.重新核收, mlngDeptID, mlngMachineID, mlngKey, 0, 0, _
                                    IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan)) = False Then
                mintEditState = 0
            End If
        Case mActS.置为无主
'            strsql = "Select Distinct 医嘱ID From (Select 医嘱ID From 检验项目分布 Where 标本id = [1] " & _
'                    "Union All Select 医嘱ID From 检验标本记录 Where ID = [1])"
'            Set rsTmp =zlDatabase.OpenSQLRecord(strsql, Me.Caption, mlngkey)
'            gcnOracle.BeginTrans
'            Do While Not rsTmp.EOF
'                If Not IsNull(rsTmp(0)) Then
'                    zlDatabase.ExecuteProcedure "ZL_检验标本记录_转为无主(" & rsTmp(0) & ")", gstrSysName
'                End If
'                rsTmp.MoveNext
'            Loop
'            gcnOracle.CommitTrans
            '取消核收
            '是否只能回滚自已的标本
            If InStr(1, mstrPrivs, "修改他人结果") <= 0 And UserInfo.姓名 <> strVerifyMan Then
                MsgBox "你不能回滚他人的报告单！", vbInformation, Me.Caption
                Call SetControlFocus
                Exit Sub
            End If
            
            frmLisStationCheckCancel.ShowEdit Me, mlngKey, Me.WinsockC, False, True
'            InsertOneRecored mlngKey, False
            Call RefreshData
            gintSelectFocus = 1
        Case mActS.合并标本
            If Not rptList.FocusedRow Is Nothing Then                                   '没有焦点行时退出
                
                With Me.rptList.FocusedRow
                    Call mfrmLabMainSampleUnion.zlRefresh(mlngKey, Nvl(.Record(mCol.姓名).Value), Nvl(.Record(mCol.性别).Value), Nvl(.Record(mCol.年龄).Value), _
                                                    Nvl(.Record(mCol.检验项目).Value), Nvl(.Record(mCol.标本号).Value), Nvl(.Record(mCol.仪器名).Value), _
                                                    Nvl(.Record(mCol.报告时间).Value), Nvl(.Record(mCol.检验人).Value))
                End With
            End If
        Case mActS.合并标本保存
            mfrmLabMainSampleUnion.ZlSave
            Call RefreshData
        Case mActS.修改病人信息
'            Call ModifyPatientBaseInfo(mlngKey)
            
    End Select
    
    
    '过滤界面列表
'    RptListFilter
'    gintSelectFocus = 1
    If Me.rptList.Rows.Count = 0 And mintEditState = 0 Then
        mfrmRequest.ZlCancel
        mfrmWrite.ZlCancel
        mfrmWrite2.zlRefresh -1
    End If
    Exit Sub
errH:
    AutoRefresh = True                                                      '操作完成(可以进行刷新)
    If blnRollBak = True Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ModifyPatientBaseInfo(lngKey As Long)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                               修改病人基础信息
    '参数    lngKey                     标本ID
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Dim strsql As String
'    Dim rsTmp As New ADODB.Recordset
'    strsql = "Select B.病人id, B.主页id, Decode(b.病人来源, 2, 1, 0) 场合 From 检验标本记录 A, 病人医嘱记录 B Where A.医嘱id = B.Id and a.id = [1]"
'    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lngKey)
'    If rsTmp.EOF = True Then
'        MsgBox "没有找到对应的医嘱不需要修改！", vbInformation, Me.Caption
'        Exit Sub
'    End If
'    Call zldatabase.zlModiPatiBaseInfo(Val(rsTmp("病人id") & ""), Val(rsTmp("主页ID") & ""), "检验报告", rsTmp("场合"))
'    Call InsertOneRecored(lngKey)
End Sub

Private Sub ReportDisposal(Disposal As Integer)
    '功能                   对报告的各种操作
    '                       批量调整报告
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strPatienName As String                             '病人姓名
    Dim strSaveAs As String                                 '是否存出
    Dim intMicrobe As Integer                               '是否是微生物
    Dim strVerifyMan As String                              '检验人
    Dim strAuditingMan As String                            '审核人
    Dim blIf As Boolean                                     '临时记录判断
    Dim strAuditingMain                                     '审核人
    Dim strVerifydate As String                             '检验时间
    Dim strAuditingDate As String                           '审核时间
    Dim lngSampleID As Long                                 '标本ID
    Dim lngPatientType As Integer                           '病人来源
    Dim lngPatientID  As Long                               '病人ID
    Dim intBay As Integer                                   '婴儿
    Dim lngApplyDept As Long                                '开嘱科室ID
    Dim lngAdviceID As Long                                 '医嘱ID
    Dim intRepotrCount As Integer                           '报告结果次数
    Dim lngPatientPage As Integer                           '主页ID
    Dim strErrInfo As String                                '错误提示
    Dim intPrivacy As Integer                               '发送报告单到医生站时是否显示隐私项目
    Dim lngAdvice As Long                                   '医嘱ID
    Dim intUnion As Integer                                 '是否不区分仪器进行显示
    Dim blnClueTo As Boolean                                '是否提示审核对话框
    Dim intLook As Integer                                  '医生站是否查看报告
    Dim strSource As String                                 '取得电子病历签名字串
    Dim lng证书ID As Long                                   '证书ID
    Dim strSign As String                                   '签名后生成的字串
    Dim strTimeStamp As String                              '时间戳
    Dim blnRollBack As Boolean                              '是否回滚
    Dim str初审人 As String                                 '初审人
    Dim intLoop As Integer
    Dim strTmp As String
    Dim lngRow As Long
    Dim astrSQL() As String
    On Error GoTo errH
    ReDim astrSQL(0)
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            strPatienName = .Record(mCol.姓名).Value
            strSaveAs = .Record(mCol.转出).Value
            intMicrobe = Val(.Record(mCol.微生物标本).Value)
            strVerifyMan = .Record(mCol.检验人).Value
            If .Record(mCol.所属情况).Value = "门诊" Then
                lngPatientType = 1
            ElseIf .Record(mCol.所属情况).Value = "住院" Then
                lngPatientType = 2
            ElseIf .Record(mCol.所属情况).Value = "院外" Then
                lngPatientType = 3
            ElseIf .Record(mCol.所属情况).Value = "体检" Then
                lngPatientType = 4
            End If
            lngPatientID = Val(.Record(mCol.病人ID).Value)
            intBay = Val(.Record(mCol.婴儿).Value)
            lngApplyDept = Val(.Record(mCol.开嘱科室ID).Value)
            lngAdviceID = Val(.Record(mCol.医嘱id).Value)
            intRepotrCount = Val(.Record(mCol.报告结果).Value)
            lngPatientPage = Val(.Record(mCol.主页ID).Value)
            lngSampleID = Val(.Record(mCol.ID).Value)
            strAuditingDate = .Record(mCol.审核时间).Value
            lngAdvice = Val(.Record(mCol.医嘱id).Value)
            intLook = IIf(.Record(mCol.查阅状态).Value = "已查阅", 1, 0)
            str初审人 = .Record(mCol.初审人).Value
        End With

    End If
    Select Case Disposal
    
        Case mActR.批量调整报告 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            Call frmLisStationAdjust.ShowEdit(Me, mlngDeptID, mstrPrivs)
            RefreshData
            gintSelectFocus = 1
        Case mActR.审核报告 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If strPatienName = "" Then Exit Sub
            If Me.rptList.FocusedRow.Record(mCol.采样人).Value <> "" Then
                If Me.rptList.FocusedRow.Record(mCol.采样时间).Value <> "" Then
                    If CDate(Me.rptList.FocusedRow.Record(mCol.采样时间).Value) > zlDatabase.Currentdate Then
                        MsgBox "采样时间，大于当前时间，不能进行审核！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
            If InStr(1, mstrPrivs, "审核标本") <= 0 Then
                '没有权限和其他用户登陆时退出
                MsgBox "你没有权限进行审核,请重新登陆具有审核人员进行审核!", vbInformation, gstrSysName
                Call SetControlFocus
                gintSelectFocus = 1
                Exit Sub
            End If
            blIf = False

            If (strVerifyMan = mstrAuditingMan Or (mstrAuditingMan = "" And strVerifyMan = UserInfo.姓名)) And InStr(1, mstrPrivs, "审核限制") > 0 Then
                '没有登陆审核人
                If mintAuditing = 0 Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再登记.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    '同一病人被权限控制不能进行审核
                    
'                    MsgBox "检验人和审核人为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
                End If

                
                '判断登陆时后的审核人是否为同一人.
                If strVerifyMan = mstrAuditingMan Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再登记.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    If strVerifyMan = mstrAuditingMan Then
                        MsgBox "检验人和审核人为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    '登陆进入的审核人和当前用户为同一个人
'                    MsgBox "登陆进入的审核人和当前用户为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
                End If
            End If
            '审核时间是否过期
            If mintAuditing < 0 Then
                If DateDiff("n", mDataAuditing, Now) > Abs(mintAuditing) * 60 Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再登记.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
'                        MsgBox "审核有效时间已过,请重新登陆审核人!", vbInformation, gstrSysName
                    '在有效时间段内可以进行审核
                End If
            End If
            
            blnClueTo = zlDatabase.GetPara("审核时不需要提示", 100, 1208, 0)
            
'            If blnClueTo = False Then
'                If mintHandleState = 0 Then
'                    If MsgBox("真的要审核“" & strPatienName & "”标本的报告吗？", _
'                            vbQuestion + vbYesNo, gstrSysName) = vbNo Then
'                            Call SetControlFocus
'                            Exit Sub
'                    End If
'                End If
'            End If
            
            '11210 权限“未收费审核”，在审核单个病人时，未生效，
            If InStr(mstrPrivs, "未收费审核") <= 0 Then
                If CheckChargeState(mlngKey, False) = False Then
                    MsgBox "单据未收费，不能进行审核！", vbInformation, gstrSysName
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            
            '21137 已归档报告不能审核
            gstrSql = "Select Decode(病案状态, 1, '1-等待审查', 2, '2-拒绝审查', 3, '3-正在审查', 4, '4-审查反馈', 5, '5-审查归档') As 病案状态" & vbNewLine & _
                    "From 检验标本记录 A, 病案主页 B ,病案提交记录 C" & vbNewLine & _
                    "Where A.病人id = B.病人id And A.主页id = B.主页id And A.病人来源 = 2 And Nvl(B.病案状态, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                    " And b.病人id = c.病人Id and B.主页id = C.主页ID "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            If rsTmp.EOF = False Then
                MsgBox "病人本次住院的病案已提交审查，不能进行审核！", vbInformation, Me.Caption
                Exit Sub
            End If
            
            '检查住院病人是否出院后还有划价单
            If CheckExesState(mlngKey) = False Then
                MsgBox "当前住院病人还有划价单未审核，但已出院或预出院！", vbInformation, Me.Caption
                Call SetControlFocus
                gintSelectFocus = 1
                Exit Sub
            End If
            
            '检验病人信息不一致时使用病人信息替换
            Call CheckPatientInfo(mlngKey)
            
            
            '检验审核规则判断
            strErrInfo = ""
            If VerifyAuditingRule(mlngKey, strErrInfo) = 1 Then
                If Mid(strErrInfo, 1, 2) = "1|" And InStr(mstrPrivs, "强制审核规则") <= 0 Then
                    strErrInfo = Mid(strErrInfo, 3)
                    MsgBox "<" & strPatienName & ">的检验单审核未通过!" & vbNewLine & strErrInfo
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
                strErrInfo = Mid(strErrInfo, 3)
                If MsgBox("<" & strPatienName & ">的检验单审核未通过!是否续继?" & vbNewLine & strErrInfo, _
                    vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            
            
            intPrivacy = zlDatabase.GetPara("报告单是否显示隐私项目", 100, 1208, 0)
            If mintUnion = 1 Then
                If mSendReport = 1 And str初审人 = "" Then
                    '初审
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_检验标本记录_初审报告(" & mlngKey & ",1,'" & UserInfo.姓名 & "')"
                Else
                    gstrSql = " select id from 检验标本记录 where 医嘱id = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAdvice)
                    Do Until rsTmp.EOF
                        '签名不成功时退出
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "Signature;" & rsTmp("ID") & ";" & mstrAuditingManID
                        
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_报告审核(" & rsTmp("ID") & ",'" & IIf(mstrAuditingMan = "" _
                            , UserInfo.姓名, mstrAuditingMan) & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
                            

                        
                        ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                        astrSQL(UBound(astrSQL)) = "Zl_检验报告单_Update(" & rsTmp("ID") & "," & intPrivacy & ",'" & gstrUnitName & "')"         '审核后处理病历报告单

                        
                        rsTmp.MoveNext
                    Loop
                End If
            Else
                If mSendReport = 1 And str初审人 = "" Then
                    '初审
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_检验标本记录_初审报告(" & mlngKey & ",1,'" & UserInfo.姓名 & "')"
                Else
                    '签名不成功时退出
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Signature;" & mlngKey & ";" & mstrAuditingManID
                    
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "ZL_检验标本记录_报告审核(" & mlngKey & ",'" & IIf(mstrAuditingMan = "" _
                        , UserInfo.姓名, mstrAuditingMan) & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"

                    
                    ReDim Preserve astrSQL(UBound(astrSQL) + 1)
                    astrSQL(UBound(astrSQL)) = "Zl_检验报告单_Update(" & mlngKey & "," & intPrivacy & ",'" & gstrUnitName & "')"        '审核后处理病历报告单

                End If
            End If
            
            '集中执行SQL
            gcnOracle.BeginTrans
            blnRollBack = True
            For intLoop = 1 To UBound(astrSQL)
                If UCase(Mid(astrSQL(intLoop), 1, 3)) = "ZL_" Then
                    zlDatabase.ExecuteProcedure astrSQL(intLoop), "审核标本"
                Else
                    '签名不成功时退出
                    If Signature(Val(Split(astrSQL(intLoop), ";")(1)), mstrAuditingManID) = False Then
                        gcnOracle.RollbackTrans
                        blnRollBack = False
                        Exit Sub
                    End If
                End If
            Next
            gcnOracle.CommitTrans
            
            Me.rptList.FocusedRow.Record(mCol.执行状态).Value = "已检验"
            Me.rptList.FocusedRow.Record(mCol.执行状态).Icon = 7
            
            If blnAutoPrint Then ReportPrint True                                           '是否完成后直接打印报告
            If mblnAout = False Then
                MoveStation 1, 2
            Else
                MoveStation 1, 0
            End If
            InsertOneRecored mlngKey, False
            gintSelectFocus = 1
        Case mActR.发送报告                                                                         '发送报告
            '检验审核规则判断
            strErrInfo = ""
            If VerifyAuditingRule(mlngKey, strErrInfo) = 1 Then
                If Mid(strErrInfo, 1, 2) = "1|" And InStr(mstrPrivs, "强制审核规则") <= 0 Then
                    strErrInfo = Mid(strErrInfo, 3)
                    MsgBox "<" & strPatienName & ">的检验单审核未通过!" & vbNewLine & strErrInfo
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
                strErrInfo = Mid(strErrInfo, 3)
                If MsgBox("<" & strPatienName & ">的检验单审核未通过!是否续继?" & vbNewLine & strErrInfo, _
                    vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            

            gstrSql = "Zl_检验标本记录_初审报告(" & mlngKey & ",1,'" & UserInfo.姓名 & "')"
            zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            InsertOneRecored mlngKey, False
            MoveStation 1, 2
'            Me.rptList.FocusedRow.Record(mCol.查阅状态).Value = "已查阅"
'            Me.rptList.FocusedRow.Record(mCol.查阅状态).Icon = 12
'            Me.rptList.Populate
        Case mActR.批量审核报告 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If InStr(1, mstrPrivs, "审核标本") <= 0 Then
                '没有权限和其他用户登陆时退出
                MsgBox "你没有权限进行审核,请重新登陆具有审核人员进行审核!", vbInformation, gstrSysName
                Call SetControlFocus
                gintSelectFocus = 1
                Exit Sub
            End If
            blIf = False

            If (strVerifyMan = mstrAuditingMan Or (mstrAuditingMan = "" And strVerifyMan = UserInfo.姓名)) And InStr(1, mstrPrivs, "审核限制") > 0 Then
                '没有登陆审核人
                If mintAuditing = 0 Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再登记.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    '同一病人被权限控制不能进行审核

'                    MsgBox "检验人和审核人为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
                End If

                '判断登陆时后的审核人是否为同一人.
                If strVerifyMan = mstrAuditingMan Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再登记.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    If strVerifyMan = mstrAuditingMan Then
                        MsgBox "检验人和审核人为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                    '登陆进入的审核人和当前用户为同一个人
'                    MsgBox "登陆进入的审核人和当前用户为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
                End If
            End If

'            Call frmLisStationAuditing.ShowEdit(Me, mlngDeptID, mstrPrivs, IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan), _
                                                mintAuditing, _
                                                mDataAuditing)
            '审核时间是否过期
            If mintAuditing < 0 Then
                If DateDiff("n", mDataAuditing, Now) > Abs(mintAuditing) * 60 Then
                    AuditingRegister
                    If mstrAuditingMan = "" Then
                        MsgBox "没有审核人,不能进行审核!请取消保存后审核再登记.", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
'                        MsgBox "审核有效时间已过,请重新登陆审核人!", vbInformation, gstrSysName
                    '在有效时间段内可以进行审核
                End If
            End If
            
            Call frmBatchAction.ShowMe(Me, 2, mlngMachineID, mstrPrivs, IIf(mstrAuditingMan = "", UserInfo.姓名, mstrAuditingMan), _
                                                mintAuditing, _
                                                mDataAuditing, mlngDeptID, mstrAuditingManID)
            gintSelectFocus = 1
        Case mActR.审核取消 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If strSaveAs = "√" Then
                MsgBox "当前检验申请已转入备份，不能取消审核！", vbInformation, gstrSysName
                Call SetControlFocus
                gintSelectFocus = 1
                Exit Sub
            End If
            
            If InStr(";" & mstrPrivs & ";", ";审核取消;") <= 0 Then
                If DateDiff("h", strAuditingDate, zlDatabase.Currentdate) > 24 Then
                    MsgBox "你只能取消24小时内的审核报告单，请联系上级技师取消审核!", vbInformation, Me.Caption
                    Call SetControlFocus
                    Exit Sub
                End If
            End If
            '21434
            If InStr(";" & mstrPrivs & ";", ";已审已打印可回滚;") <= 0 Then
                If Me.rptList.FocusedRow.Record(mCol.执行状态).Icon = 8 Then
                    MsgBox "您只能取消未打印的审核报告单，请联系上级技师取消审核!", vbInformation, Me.Caption
                    Call SetControlFocus
                    Exit Sub
                End If
            End If
            '21137 已归档报告不能取消
            gstrSql = "Select Decode(病案状态, 1, '1-等待审查', 2, '2-拒绝审查', 3, '3-正在审查', 4, '4-审查反馈', 5, '5-审查归档') As 病案状态" & vbNewLine & _
                    "From 检验标本记录 A, 病案主页 B ,病案提交记录 C" & vbNewLine & _
                    "Where A.病人id = B.病人id And A.主页id = B.主页id And A.病人来源 = 2 And Nvl(B.病案状态, 0) >= 1 and A.ID=[1] " & vbNewLine & _
                    " And b.病人id = c.病人Id and B.主页id = C.主页ID "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            If rsTmp.EOF = False Then
                MsgBox "病人本次住院的病案已提交审查，不能取消审核！", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If intLook = 0 Then
                strTmp = "真的要取消“" & strPatienName & "”标本的报告审核吗？"
            Else
                strTmp = "医生已查阅“" & strPatienName & "“的报告，是否确认要取消审核？"
            End If
            
            If MsgBox(strTmp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Call SetControlFocus
                gintSelectFocus = 1: Exit Sub
            End If
            
            gstrSql = "select 检验标本id from 检验签名记录 where 检验标本id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngKey)
            If rsTmp.EOF = False Then
                If gobjESign Is Nothing Then
                    MsgBox "不能取消签名，请在系统参数中设置使用电子签名。", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
            
            
            If mintUnion = 1 Then
                gstrSql = " select id from 检验标本记录 where 医嘱id = [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAdvice)
                Do Until rsTmp.EOF
                strSQL = "ZL_检验标本记录_审核取消(" & rsTmp("ID") & ")"
                    zlDatabase.ExecuteProcedure strSQL, Me.Caption
                    rsTmp.MoveNext
                Loop
            Else
                strSQL = "ZL_检验标本记录_审核取消(" & mlngKey & ")"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            InsertOneRecored mlngKey, False
'            Me.rptList.FocusedRow.Record(mCol.执行状态).Value = ""
'            Me.rptList.FocusedRow.Record(mCol.执行状态).Icon = -1
'            Me.rptList.FocusedRow.Record(mCol.查阅状态).Value = ""
'            Me.rptList.FocusedRow.Record(mCol.查阅状态).Icon = -1
'            Me.rptList.Populate
'            RptListFilter
            gintSelectFocus = 1
'            RefreshData

        Case mActR.按病人审核
            frmPatinetAuditing.ShowMe Me, mstrPrivs, mstrAuditingManID
            Call RefreshData
            
        Case mActR.重做结果 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
            If intMicrobe = 1 Then Exit Sub     '如果是微生物退出
            
            If MsgBox("真的要重做“" & strPatienName & "”标本的检验吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call SetControlFocus
                        gintSelectFocus = 1: Exit Sub
            End If
            strSQL = "ZL_检验标本记录_标本重做(" & mlngKey & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            RefreshData
            gintSelectFocus = 1
        Case mActR.取消重做 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If MsgBox("真的要取消“" & strPatienName & "”的检验结果吗？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Call SetControlFocus
                        gintSelectFocus = 1: Exit Sub
            End If
            strSQL = "ZL_检验标本记录_取消重做(" & mlngKey & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            RefreshData
            gintSelectFocus = 1
        Case mActR.填写报告 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '有初审人时不能再修改结果了
            If str初审人 <> "" And mSendReport = 1 Then Exit Sub
            
            strSQL = "select 检验人,检验时间 from 检验标本记录 where id = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, mlngKey)
            If rsTmp.EOF = False Then
                strVerifyMan = Nvl(rsTmp("检验人"))
                strVerifydate = Nvl(rsTmp("检验时间"))
            End If
            
            '处理是否能修改他人报告
            If UserInfo.姓名 <> strVerifyMan And strVerifyMan <> "" Then
                If InStr(1, mstrPrivs, "修改他人结果") <= 0 Then
                    MsgBox "您没有修改他人报告的权限，请与管理员联系！", vbInformation, gstrSysName
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            
            '处理能够填写或修改非本日检验的报告结果
            If strVerifydate <> "" Then
                If DateDiff("d", CDate(strVerifydate), Now) > 1 Then
                    If InStr(1, mstrPrivs, "修改往日结果") <= 0 Then
                        MsgBox "你没有权限填写或修改非本日检验的报告结果", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            If Val(Me.rptList.FocusedRow.Record(mCol.微生物标本).Value) <> 1 Then
                Me.TabCtlWindow.Item(0).Selected = True
                If mfrmWrite.ZlEditStart(mlngKey) = True Then
                    mintEditState = 5
                End If
            Else
                Me.TabCtlWindow.Item(1).Selected = True
                If mfrmWrite2.ZlEditStart(mlngKey) = True Then
                    mintEditState = 5
                End If
            End If
        Case mActR.填写三级报告
            
            
            '处理是否能修改他人报告
            If UserInfo.姓名 <> strVerifyMan And strVerifyMan <> "" Then
                If InStr(1, mstrPrivs, "修改他人结果") <= 0 Then
                    MsgBox "您没有修改他人报告的权限，请与管理员联系！", vbInformation, gstrSysName
                    Call SetControlFocus
                    gintSelectFocus = 1
                    Exit Sub
                End If
            End If
            
            '处理能够填写或修改非本日检验的报告结果
            If strVerifydate <> "" Then
                If DateDiff("d", CDate(strVerifydate), Now) > 1 Then
                    If InStr(1, mstrPrivs, "修改往日结果") <= 0 Then
                        MsgBox "你没有权限填写或修改非本日检验的报告结果", vbInformation, gstrSysName
                        Call SetControlFocus
                        gintSelectFocus = 1
                        Exit Sub
                    End If
                End If
            End If
            
            
            
            mintEditState = 6
            mfrmLabMicrobe3Report.ZlEditStart
            
            
        Case mActR.写入病历
            
            If intMicrobe = 1 Then
                strSQL = "Zl_检验报告单_Update(" & mlngKey & ",0,'" & gstrUnitName & "')"       '微生物三级报告单发布
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
        Case mActR.验证签名
            Call VerifySignature(mlngKey)
    End Select
    
    '过滤界面列表
'    RptListFilter
    If Me.rptList.Rows.Count = 0 And mintEditState = 0 Then
        mfrmRequest.ZlCancel
        mfrmWrite.ZlCancel
        mfrmWrite2.zlRefresh -1
    End If
    Exit Sub
errH:
    If blnRollBack = True Then
        blnRollBack = False
        gcnOracle.RollbackTrans
    End If
    AutoRefresh = True                                                                      '重新开始自动刷新
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub LoadRegistSetup()
    mintEditState = 0                                           '当前编辑状态：0-非编辑；1-新增核收；2-新增登记；4-补填病人；3-重新核收；5-报告编辑
    
    On Error GoTo errH
    
    '目前是否处于连续核收登记状态
    mintContinue = IIf(zlDatabase.GetPara("连续输入", 100, 1208, False), 1, 0)

    Set mfrmRequest = New frmLabRequest                             '核收登记窗体
    Set mfrmWrite = New frmLisStationWrite                          '报告填写窗体
    Set mfrmWrite2 = New frmLisStationWrite2                        '填写微生物
    Set mfrmTrack = New frmLabTrack                                 '历次对比
    Set mfrmLabMicrobe3Report = New frmLabMicrobe3Report            '三级报告
    Set mfrmLabMainSampleUnion = frmLabMainSampleUnion              '标本合并
'    Set mclsExpenses = New zlCISKernel.clsDockExpense           '费用部件
'    Set mclsOutAdvices = New zlCISKernel.clsDockOutAdvices      '门诊医嘱
'    Set mclsInAdvices = New zlCISKernel.clsDockInAdvices        '住院医嘱
'    Set mclsOutEPRs = New zlRichEPR.cDockOutEPRs                '门诊医历
'    Set mclsInEPRs = New zlRichEPR.cDockInEPRs                  '住院病历
'    Set mcolSubForm = New Collection
    
'    mcolSubForm.Add mclsExpenses.zlGetForm, "_费用"             '得到子窗体
'    mcolSubForm.Add mclsOutAdvices.zlGetForm, "_门诊医嘱"
'    mcolSubForm.Add mclsInAdvices.zlGetForm, "_住院医嘱"
'    mcolSubForm.Add mclsOutEPRs.zlGetForm, "_门诊病历"
'    mcolSubForm.Add mclsInEPRs.zlGetForm, "_住院病历"
    
'    Set mfrmLabMainImage = frmLabMainImage
'    Call mclsExpenses.zlDefCommandBars(Me, Me.cbrthis)
    
    
    '科室和仪器ID
    mlngDeptID = zlDatabase.GetPara("缺省科室ID", 100, 1208, mlngDeptID)
    mlngMachineID = zlDatabase.GetPara("过滤仪器", 100, 1208, mlngMachineID)
    mstrMachineGroup = zlDatabase.GetPara("仪器小组", 100, 1208, mstrMachineGroup)
    mblnAout = zlDatabase.GetPara("审核后跳到下一个可审标本", 100, 1208, mblnAout)
    
    '界面刷新过滤
    Call GetVerifying
'    chkSoure(0).Value = IIf(mblnVerifying(0), 1, 0)
'    chkSoure(1) = IIf(mblnVerifying(1), 1, 0)
'    chkSoure(2) = IIf(mblnVerifying(2), 1, 0)
'    chkSoure(3) = IIf(mblnVerifying(3), 1, 0)
'    chkSoure(4) = IIf(mblnVerifying(4), 1, 0)
'    chkSoure(5) = IIf(mblnVerifying(5), 1, 0)
    '最后刷新Dpk格式，保存数据库太长了
'    dkpMain.LoadStateFromString zlDatabase.GetPara("DKP保存", 100, 1208, "")
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
    
    blnComm = Val(zlDatabase.GetPara("核收允许双向", 100, 1208, 0))
    blnAutoPrint = zlDatabase.GetPara("审核打印", 100, 1208, 0)
    blnAutoRefresh = Val(zlDatabase.GetPara("自动刷新", 100, 1208, 1))
    mintUnion = zlDatabase.GetPara("不区分仪器显示核收项目", 100, 1208, 0)
    mMakeNoRule = zlDatabase.GetPara("标本序号生成规则", 100, 1208, "今  天")
    mSendReport = zlDatabase.GetPara("使用二级报告审核", 100, 1208, 0)
    mstrPrintDepts = zlDatabase.GetPara("只打指定科室报告单", 100, 1208, "")
    
    int体检处理方式 = Val(zlDatabase.GetPara("体检病人信息不一致的处理方式", 100, 1208, True, 1))
    int院外处理方式 = Val(zlDatabase.GetPara("院外病人信息不一致的处理方式", 100, 1208, True, 1))
    int住院处理方式 = Val(zlDatabase.GetPara("住院病人信息不一致的处理方式", 100, 1208, True, 1))
    int门诊处理方式 = Val(zlDatabase.GetPara("门诊病人信息不一致的处理方式", 100, 1208, True, 1))
    
    mTodayQCPrivs = GetPrivFunc(100, 1210)
    mHistoryPrivs = GetPrivFunc(100, 1211)
    
    '检验中和待核收时间选择
    With cbo时间
        .AddItem "今  天"
        .AddItem "昨  天"
        .AddItem "本  周"
        .AddItem "本  月"
        .AddItem "本  季"
        .AddItem "本半年"
        .AddItem "本  年"
        .AddItem "前三天"
        .AddItem "前一周"
        .AddItem "前半月"
        .AddItem "前一月"
        .AddItem "前二月"
        .AddItem "前三月"
        .AddItem "前半年"
        .AddItem "自定义"
    End With
    
    dtpDate.Value = Now
    dtpDateEnd.Value = Now
    
    '电子签名认证中心
    gintCA = Val(zlDatabase.GetPara("电子签名认证中心", glngSys))
    '电子签名控制场合
    gstrESign = zlDatabase.GetPara("电子签名使用场合", glngSys)
    
    If Mid(gstrESign, 6, 1) = "1" Then
        If gintCA <> 0 Then
            'If InStr(GetInsidePrivs(p门诊医嘱下达), "医嘱电子签名") > 0 And gobjESign Is Nothing Then
            If gobjESign Is Nothing Then
                On Error Resume Next
                Set gobjESign = CreateObject("zl9ESign.clsESign")
                Err.Clear: On Error GoTo 0
                If Not gobjESign Is Nothing Then
                    Call gobjESign.Initialize(gcnOracle, glngSys)
                End If
            End If
        Else
            Set gobjESign = Nothing
        End If
    End If
    Exit Sub
errH:
    MsgBox Err.Description
End Sub
Private Sub ShowOrHideItem(Control As CommandBarControl, DkpID As Integer)
    '功能               '显示或隐藏
    Dim Pane As Pane
    Set Pane = Me.dkpMain.FindPane(DkpID)
    If Control.Checked = True Then
        Pane.Close
    Else
        Pane.Select
    End If
    If mlngKey <> 0 Then
        ReadImageData mlngKey, False
    End If
    Me.dkpMain.RecalcLayout
    Me.cbrthis.RecalcLayout
End Sub
Private Sub BackOrNextPatient(Move As Integer)
    '功能                 移动到上一个病人或下一个病人
    '参数                 Move = 1 上一病人 =2 下一病人
    Dim Rerow As ReportRow
    Dim i As Long
    With Me.rptList
        If .Rows.Count <= 0 Then Exit Sub
        i = .SelectedRows(0).Index
        If Move = 1 Then            '向上移动
            If i - 1 >= 0 Then
                i = i - 1
                .FocusedRow = .Rows(i)
            End If
        Else
            If i < .Rows.Count - 1 Then
                i = i + 1
                .FocusedRow = .Rows(i)
            End If
        End If
    End With
End Sub

Private Function FindPatient(strFind As String) As Boolean
    '功能:              查找病人
    '参数               查询字段 标识号 标本号 姓名和拼音缩写
    '规则               "数字为标本号、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
    Dim Rerow As ReportRow
    Dim strPatientID As String                                          '标识号
    Dim strSampleID As String                                           '标本号
    Dim strPatientName As String                                        '病人姓名
    Dim strPatientPY As String                                          '病人拼音简码
    Dim lngPatientID As Long                                            '病人ID
    Dim strSource As String                                             '所属情况
    Dim strRegisterNo As String                                         '挂号单
    Dim strChargeNo As String                                           '收费单
    Dim strBarCode As String                                            '样本条码
    Dim strSQL As String                                                '数据查询语句
    Dim rsTmp As New ADODB.Recordset                                    '数据集
    Dim strWhere As String                                              '过滤条件语句
    On Error GoTo errH
    
    '定位前先刷新一次
'    Call RefreshData
    
    
    
    If Me.TabList(0).Selected = True Then
        If Me.rptList.Rows.Count = 0 Then Exit Function                     '没有记录时退出
        strFind = UCase(strFind)
        For Each Rerow In Me.rptList.Rows
            '先取出所需要的字段信息
            Select Case Mid(strFind, 1, 1)
                Case "-"                                                    '病人ID
                    lngPatientID = Val(Rerow.Record(mCol.病人ID).Value)
                    strWhere = Mid(strFind, 2)
                    If strWhere = lngPatientID Then
                        Me.rptList.FocusedRow = Rerow
                        Me.rptList.Populate
                        mlngKey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "+", "*"                                               '门诊号/住院号
                    strPatientID = Rerow.Record(mCol.标识号).Value
                    strWhere = Mid(strFind, 2)
                    If strWhere = strPatientID Then
                        Me.rptList.FocusedRow = Rerow
                        Me.rptList.Populate
                        mlngKey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "."                                                    '挂号单
                    strRegisterNo = Rerow.Record(mCol.挂号单).Value
                    strWhere = Mid(strFind, 2)
                    If strWhere = strRegisterNo Then
                        Me.rptList.FocusedRow = Rerow
                        Me.rptList.Populate
                        mlngKey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "/"                                                    '收费单
                    strChargeNo = Rerow.Record(mCol.标识号).Value
                    strWhere = zlCommFun.GetFullNO(Mid(strFind, 2))
                    If strWhere = strChargeNo Then
                        Me.rptList.FocusedRow = Rerow
                        Me.rptList.Populate
                        mlngKey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case Else                                                   '标本号、姓名、简码查找、条码
                    strSampleID = Nvl(Rerow.Record(mCol.标本号).Value)
                    strPatientName = Rerow.Record(mCol.姓名).Value
                    strPatientPY = zlCommFun.SpellCode(Rerow.Record(mCol.姓名).Value)
                    strBarCode = Rerow.Record(mCol.样本条码).Value
                    If strSampleID = strFind Or (strPatientName Like UCase(strFind) & "*") Or (strPatientPY Like UCase(strFind) & "*") _
                            Or strBarCode = UCase(strFind) Then
                        If Val(Rerow.Record(mCol.定位).Value) <= 0 Then
                            Rerow.Record(mCol.定位).Value = 1
                            Me.rptList.FocusedRow = Rerow
                            Me.rptList.Populate
                            mlngKey = Rerow.Record(mCol.ID).Value
                            FindPatient = True
                            Exit Function
                        End If
                    End If
                    
            End Select
        Next
        For Each Rerow In Me.rptList.Rows
            Rerow.Record(mRCol.定位).Value = 0
        Next
        '单独处理条码
        If BlnIsNumber(strFind) Then
            strSQL = "select distinct b.病人id  from 病人医嘱发送 a , 病人医嘱记录 b " & _
                     " Where a.医嘱id = b.ID And a.样本条码 = [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strFind)
            If rsTmp.EOF = False Then
                Me.rptList.Tag = ";;,;;;;;;;,True;,;0;;0;;;;1;;" & rsTmp(0)
                RefreshData True
                rptList.Tag = ""
                Exit Function
            End If
        End If
        Me.rptList.Populate
    End If
    
    If Me.TabList(1).Selected = True Then
        If Me.rptList1.Rows.Count = 0 Then Exit Function                     '没有记录时退出
        strFind = UCase(strFind)
        For Each Rerow In Me.rptList1.Rows
            '先取出所需要的字段信息
            Select Case Mid(strFind, 1, 1)
                Case "-"                                                    '病人ID
                    lngPatientID = Val(Rerow.Record(mRCol.病人ID).Value)
                    strWhere = Mid(strFind, 2)
                    If strWhere = lngPatientID Then
                        Me.rptList1.FocusedRow = Rerow
                        Me.rptList1.Populate
'                        mlngkey = Rerow.Record(mRcol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "+", "*"                                               '门诊号/住院号
                    strPatientID = Rerow.Record(mRCol.标识号).Value
                    strWhere = Mid(strFind, 2)
                    If strWhere = strPatientID Then
                        Me.rptList1.FocusedRow = Rerow
                        Me.rptList1.Populate
'                        mlngkey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "."                                                    '挂号单
                    strRegisterNo = Rerow.Record(mRCol.挂号单).Value
                    strWhere = Mid(strFind, 2)
                    If strWhere = strRegisterNo Then
                        Me.rptList1.FocusedRow = Rerow
                        Me.rptList1.Populate
'                        mlngkey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case "/"                                                    '收费单
                    strChargeNo = Rerow.Record(mRCol.标识号).Value
                    strWhere = zlCommFun.GetFullNO(Mid(strFind, 2))
                    If strWhere = strChargeNo Then
                        Me.rptList1.FocusedRow = Rerow
                        Me.rptList1.Populate
'                        mlngkey = Rerow.Record(mCol.ID).Value
                        FindPatient = True
                        Exit Function
                    End If
                Case Else                                                   '标本号、姓名、简码查找、条码
'                    strSampleID = Nvl(Rerow.Record(mCol.标本号).Value)
                    strPatientName = Rerow.Record(mRCol.姓名).Value
                    strPatientPY = zlCommFun.SpellCode(Rerow.Record(mRCol.姓名).Value)
'                    strBarCode = Rerow.Record(mCol.样本条码).Value
                    If strSampleID = strFind Or (strPatientName Like UCase(strFind) & "*") Or (strPatientPY Like UCase(strFind) & "*") _
                            Or strBarCode = UCase(strFind) Then
                        If Val(Rerow.Record(mRCol.定位).Value) <= 0 Then
                            Rerow.Record(mRCol.定位).Value = 1
                            Me.rptList1.FocusedRow = Rerow
                            Me.rptList1.Populate
    '                        mlngkey = Rerow.Record(mCol.ID).Value
                            FindPatient = True
                            Exit Function
                        End If
                    End If
            End Select
        Next
        For Each Rerow In Me.rptList1.Rows
            Rerow.Record(mRCol.定位).Value = 0
        Next
        '单独处理条码
'        If IsNumeric(strFind) = True And Len(strFind) >= 12 Then
'            strsql = "select distinct b.病人id  from 病人医嘱发送 a , 病人医嘱记录 b " & _
'                     " Where a.医嘱id = b.ID And a.样本条码 = [1] "
'            Set rsTmp = zlDatabase.OpenSQLRecord(strsql, gstrSysName, strFind)
'            If rsTmp.EOF = False Then
'                rptList.Tag = ";;,;;;;;,True;,;0;;0;0;;" & rsTmp(0)
'                RefreshData True
'                rptList.Tag = ""
'                Exit Function
'            End If
'        End If
        Me.rptList.Populate
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vgdList, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vgdList
    objPrint.Title.Text = "病人采集清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub InsertOneRecored(lngKey As Long, Optional blnNew As Boolean = True, Optional blnGoto As Boolean = True)
    '功能                                               '通过检验标本ID找到一记录并追加到列表
    '参数   blnNew                                      是否新增标本
    '       blnGoto                                     是否定位
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset, rsZk As ADODB.Recordset
    Dim Record As ReportRecord
    Dim i As Integer
    Dim blnNewRecord As Boolean               '是否新增记录自动判断
    Dim lngRowIndex As Long                                         '行索行
    Dim lngRowID As Long                                            '行ID
    Dim lngloop As Long
    Dim intLoop As Integer
    Dim blnPathPatient As Boolean                                   '临床路径病人
    Dim blnAdviceKey As Long                                        '医嘱ID
    
    mblnCompelRefresh = True    '刷新时可以强制刷新
    blnPathPatient = False
    On Error GoTo errH
    
    strSQL = "select 医嘱id from 检验标本记录 where id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    If rsTmp.EOF = True Then Exit Sub
    blnAdviceKey = Nvl(rsTmp("医嘱ID"), 0)
    
    strSQL = "Select /*+ rule */     Decode(a.是否传送, 1, '', '传送失败') As 传送," & vbNewLine & _
            "       decode(a.标本类别,1,'紧急',decode(a.紧急,1,'紧急', '')) As 紧急,Decode(a.样本状态, 1, '检验中', 2, '已检验') As 执行状态," & vbNewLine & _
            "       Decode(A.病人来源, 1, '门诊', 2, '住院', 3, '院外', 4, '体检','无主') As 所属情况," & vbNewLine & _
            "       Decode(Sign(Nvl(a.是否质控品, 0)), 0, '普通', 1, '质控', -1, '比对') As 标本类型," & vbNewLine & _
            "       Decode(a.仪器id, Null," & vbNewLine & _
            "                 To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000')," & vbNewLine & _
            "                 a.标本序号) As 标本号显示,a.标本序号, A.挂号单 ," & vbNewLine & _
            "       Decode(A.病人来源, 1, to_char(nvl(a.门诊号,a.标识号)), 2, to_char(nvl(a.住院号,a.标识号)), 3, to_char(nvl(a.NO,a.标识号)), 4, to_char(nvl(a.门诊号,a.标识号)),to_char(a.标识号)) As 标识号,a.姓名,a.性别,a.年龄," & vbNewLine & _
            "       Decode(a.病人来源,2,S.病人类型,b.病人类型) as 病人类型," & vbNewLine & _
            "       a.报告结果 As 结果次数,a.医嘱ID,a.仪器ID,'' As 转出,a.Id,a.核收时间 ,a.打印次数,a.病人id," & vbNewLine & _
            "       a.检验时间,a.微生物标本,a.检验人,a.审核人,To_Char(A.婴儿) As 婴儿,a.样本条码,a.申请科室ID As 开嘱科室id," & vbNewLine & _
            "       a.主页ID,a.报告结果,a.年龄数字,a.年龄单位,a.门诊号,a.住院号,a.出生日期,a.挂号单,a.检验项目,e.名称 as 申请科室,f.名称 as 仪器名称, " & vbNewLine & _
            "       a.申请科室ID as 病人科室ID,a.床号,a.申请人,a.标本形态,a.采样人,a.采样时间,a.标本类型 as 检验标本,a.NO,a.接收人,a.接收时间, " & vbNewLine & _
            "       abs(nvl(a.是否质控品,0)) as 比对次数,a.审核时间,nvl(a.标本类别,0) as 标本类别, " & vbNewLine & _
            "       nvl(a.紧急,0) as 医嘱紧急, nvl(a.标本类别,0) as 标本紧急,decode(a.病人科室,null,C.名称,a.病人科室) as 病人科室, " & vbNewLine & _
            "       a.申请类型,nvl(R.查阅状态,0) as 查阅状态,nvl(R.病历ID,0) as 报告发送,a.初审人,a.初审时间,b.工作单位,p.项目,p.内容,b.健康号,  " & vbNewLine & _
            "       a.审核未通过,a.病人来源,a.结果为空,nvl(s.路径状态,0) as 临床路径病人,decode(d.仪器是否审核,1,'仪器审核','仪器未审核')  as 仪器审核 " & vbNewLine & _
            " From 检验标本记录 a ,部门表 E , 检验仪器 f , 病人信息 B , 部门表 C,病人医嘱报告 R,病人医嘱附件 P,病案主页 S,检验流水线标本 d " & vbNewLine & _
            " Where a.申请科室ID = E.id(+) and a.仪器id=f.id(+) and a.病人ID = B.病人ID(+) and B.当前科室ID = C.id(+) " & vbNewLine & _
            " " & IIf(blnAdviceKey = 0, " and  a.ID = [1] ", " and a.医嘱id=[2] ") & vbNewLine & _
            " And a.医嘱ID = R.医嘱ID(+) and A.医嘱Id = P.医嘱ID(+) and a.id=d.标本id(+)" & vbNewLine & vbNewLine & _
            " and a.病人ID = S.病人ID(+) and a.主页ID = s.主页ID(+)  "
            
    If mlngMachineID > 0 Then
        strSQL = strSQL & " And a.仪器id = [3] "
    End If
      
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngKey, blnAdviceKey, mlngMachineID)
                                                             
    If rsTmp.EOF = True Then Exit Sub       '没有时退出

    '刷新前记录一下位置
    If Not Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row Is Nothing Then
        lngRowIndex = Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row.Index - 1
        lngRowID = Me.rptList.Rows(lngRowIndex).Record(mCol.ID).Value
        mlngLastShow = lngRowID
    Else
        If mlngLastShow > 0 Then
            For i = 0 To Me.rptList.Rows.Count - 1
                If Me.rptList.Rows(i).Record(mCol.ID).Value = mlngLastShow Then
                    lngRowIndex = Me.rptList.Rows(i).Record.Index
                    lngRowID = Me.rptList.Rows(i).Record(mCol.ID).Value
                End If
            Next
        End If
    End If

'    If lngRowIndex = 0 And Me.rptList.Rows.Count > 0 Then
'
'    End If
    
    
    blnNewRecord = True
    
    For i = 0 To Me.rptList.Records.Count - 1
        If Me.rptList.Records(i).Item(mCol.ID).Value = lngKey Then
            Set Record = Me.rptList.Records(i)
            blnNewRecord = False
            Exit For
        End If
    Next
    
    If blnNewRecord = True Then
        Set Record = Me.rptList.Records.Add
        For i = 0 To Me.rptList.Columns.Count + 1
            Record.AddItem ""
        Next
    End If
    
    '前面几列需要处理图标
    Record.Item(mCol.紧急).Value = IIf(Nvl(rsTmp("标本紧急")) = 1, "紧急", "")
    If Record.Item(mCol.紧急).Value = "紧急" Then
        Record.Item(mCol.紧急).Icon = 1
    Else
        Record.Item(mCol.紧急).Icon = -1
    End If
    
    Record.Item(mCol.紧急医嘱).Value = IIf(Nvl(rsTmp("医嘱紧急")) = 1, "紧急", "")
    If Record.Item(mCol.紧急医嘱).Value = "紧急" Then
        Record.Item(mCol.紧急医嘱).Icon = 14
    Else
        Record.Item(mCol.紧急医嘱).Icon = -1
    End If
    
'    If Nvl(rsTmp("初审人")) <> "" Then
'        Record.Item(mCol.查阅状态).Value = "已初审"
'        Record.Item(mCol.查阅状态).Icon = 13
'    Else
'        Record.Item(mCol.查阅状态).Value = ""
'        Record.Item(mCol.查阅状态).Icon = -1
'    End If
    
    If Nvl(rsTmp("查阅状态")) = 1 Then
        Record.Item(mCol.查阅状态).Value = "已查阅"
        Record.Item(mCol.查阅状态).Icon = 11
    End If
    If rsTmp("仪器审核") & "" = "仪器审核" Then
        Record.Item(mCol.仪器审核).Value = "√"
    Else
        Record.Item(mCol.仪器审核).Value = "×"
    End If
    If Val(Nvl(rsTmp("临床路径病人"))) = 1 Then
        blnPathPatient = True
        Record.Item(mCol.临床路径病人).Icon = 15
    Else
        Record.Item(mCol.临床路径病人).Icon = -1
    End If
    
    If CInt(Nvl(rsTmp("打印次数"), "0")) > 0 Then
        Record.Item(mCol.执行状态).Value = "已打印"
        Record.Item(mCol.执行状态).Icon = 8
    ElseIf Nvl(rsTmp("执行状态")) = "已检验" Then
        Record.Item(mCol.执行状态).Value = "已检验"
        Record.Item(mCol.执行状态).Icon = 7
    ElseIf Nvl(rsTmp("初审人")) <> "" Then
        Record.Item(mCol.执行状态).Value = "初审"
        Record.Item(mCol.执行状态).Icon = 13
    ElseIf Nvl(rsTmp("传送")) = "" Then
        Record.Item(mCol.执行状态).Value = "已传送"
        Record.Item(mCol.执行状态).Icon = 6
    Else
        Record.Item(mCol.执行状态).Value = ""
        Record.Item(mCol.执行状态).Icon = -1
    End If

    
    Record.Item(mCol.姓名).Value = Nvl(rsTmp("姓名")) '& IIf(Nvl(rsTmp("婴儿"), 0) > 0, "(婴儿)", "")
    If Nvl(rsTmp("标本类型")) = "质控" Then
        Record.Item(mCol.标本类型).Value = "质控"
        Record.Item(mCol.标本类型).Icon = 3
        strSQL = "Select A.标本id, B.名称, B.批号, B.水平 From 检验质控记录 A, 检验质控品 B Where A.质控品id = B.ID And A.标本id=[1]"
        Set rsZk = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(rsTmp("ID"))))
        Do Until rsZk.EOF
            Record.Item(mCol.姓名).Value = "" & rsZk!名称 & "," & rsZk!批号 & ",水平" & rsZk!水平
            rsZk.MoveNext
        Loop
    ElseIf Nvl(rsTmp("标本类型")) = "比对" Then
        Record.Item(mCol.标本类型).Value = "比对"
        Record.Item(mCol.标本类型).Icon = 4
        Record.Item(mCol.姓名).Value = Record.Item(mCol.姓名).Value & "(" & Nvl(rsTmp("比对次数")) & ")"
    Else
        Record.Item(mCol.标本类型).Value = ""
        Record.Item(mCol.标本类型).Icon = -1
    End If
    
    Record.Item(mCol.标本号).Value = Val(Nvl(rsTmp("标本序号")))
    Record.Item(mCol.标本号).Caption = Trim(rsTmp("标本号显示"))

    If Nvl(rsTmp("年龄数字")) = "" Then
        
        If Nvl(rsTmp("婴儿"), 0) = 0 Then
            If IsNumeric(Nvl(rsTmp("年龄"))) = True Then
                Record.Item(mCol.年龄).Caption = Nvl(rsTmp("年龄")) & "岁"
            Else
                If Nvl(rsTmp("年龄")) <> "岁" And Nvl(rsTmp("年龄")) <> "0岁" Then
                    Record.Item(mCol.年龄).Caption = Nvl(rsTmp("年龄"))
                End If
            End If
            If Record.Item(mCol.年龄).Caption <> "" Then
                Record.Item(mCol.年龄).Value = Val(rsTmp("年龄"))
            End If
        End If
    '            Record.Item(mCol.年龄).Caption = IIf(Nvl(rstmp("婴儿"), 0) > 0, "", _
                                   IIf(Nvl(rstmp("年龄")) = "岁", "", _
                                   IIf(Nvl(rstmp("年龄")) = "0岁", "", IIf(IsNumeric(Nvl(rstmp("年龄"))) = True, rstmp("年龄") & "岁", rstmp("年龄")))))
    Else
        Record.Item(mCol.年龄).Value = Nvl(rsTmp("年龄数字"))
        Record.Item(mCol.年龄).Caption = Nvl(rsTmp("年龄")) 'Nvl(rsTmp("年龄数字")) & Nvl(rsTmp("年龄单位"))
    End If
    
    If Nvl(rsTmp("病人类型")) <> "" Then
        Record.Item(mCol.姓名).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp("病人类型")), False)
    End If
    Record.Item(mCol.性别).Value = Nvl(rsTmp("性别"))
    Record.Item(mCol.所属情况).Value = Nvl(rsTmp("所属情况"))
    Record.Item(mCol.检验项目).Value = Trim(Nvl(rsTmp("检验项目")))
    Record.Item(mCol.标识号).Value = Nvl(rsTmp("标识号"))
    
    Record.Item(mCol.结果次数).Value = Nvl(rsTmp("结果次数"))
    Record.Item(mCol.医嘱id).Value = Nvl(rsTmp("医嘱ID"))
    Record.Item(mCol.仪器id).Value = Nvl(rsTmp("仪器ID"))
    Record.Item(mCol.转出).Value = Nvl(rsTmp("转出"))
    Record.Item(mCol.病人ID).Value = Nvl(rsTmp("病人id"))
    Record.Item(mCol.ID).Value = Nvl(rsTmp("ID"))
    Record.Item(mCol.标本时间).Caption = Format(Nvl(rsTmp("核收时间")), "MM-dd HH:mm:ss")
    Record.Item(mCol.标本时间).Value = Format(Nvl(rsTmp("核收时间")), "YYYY-MM-dd HH:mm:ss")
    Record.Item(mCol.报告时间).Caption = Format(Nvl(rsTmp("检验时间")), "MM-dd HH:mm")
    Record.Item(mCol.报告时间).Value = Format(Nvl(rsTmp("检验时间")), "YYYY-MM-dd HH:mm")
    Record.Item(mCol.微生物标本).Value = Val(Nvl(rsTmp("微生物标本")))
    '        Record.Item(mCol.收费单).Value = Nvl(rstmp("收费单"))
    Record.Item(mCol.挂号单).Value = Nvl(rsTmp("挂号单"))
    Record.Item(mCol.检验人).Value = Nvl(rsTmp("检验人"))
    Record.Item(mCol.审核人).Value = Nvl(rsTmp("审核人"))
    Record.Item(mCol.病人科室).Value = Nvl(rsTmp("病人科室"))
    Record.Item(mCol.样本条码).Value = Nvl(rsTmp("样本条码"))
    'Record.Item(mCol.发送号).Value = Nvl(rstmp("发送号"))
    Record.Item(mCol.婴儿).Value = Nvl(rsTmp("婴儿"))
    Record.Item(mCol.仪器名).Value = Nvl(rsTmp("仪器名称"))
    Record.Item(mCol.主页ID).Value = Nvl(rsTmp("主页ID"))
    Record.Item(mCol.开嘱科室ID).Value = Nvl(rsTmp("开嘱科室Id"))
    Record.Item(mCol.报告结果).Value = Nvl(rsTmp("报告结果"))
    Record.Item(mCol.年龄数字).Value = Nvl(rsTmp("年龄数字"))
    Record.Item(mCol.年龄单位).Value = Nvl(rsTmp("年龄单位"))
    Record.Item(mCol.床号).Value = Nvl(rsTmp("床号"))
    Record.Item(mCol.申请人).Value = Nvl(rsTmp("申请人"))
    Record.Item(mCol.标本形态).Value = Nvl(rsTmp("标本形态"))
    Record.Item(mCol.采样人).Value = Nvl(rsTmp("采样人"))
    Record.Item(mCol.采样时间).Value = Nvl(rsTmp("采样时间"))
    Record.Item(mCol.检验标本).Value = Nvl(rsTmp("检验标本"))
    Record.Item(mCol.NO).Value = Nvl(rsTmp("NO"))
    Record.Item(mCol.接收人).Value = Nvl(rsTmp("接收人"))
    Record.Item(mCol.接收时间).Value = Nvl(rsTmp("接收时间"))
    Record.Item(mCol.审核时间).Value = Nvl(rsTmp("审核时间"))
    Record.Item(mCol.标本类别).Value = Nvl(rsTmp("标本类别"))
    Record.Item(mCol.医嘱紧急).Value = Nvl(rsTmp("医嘱紧急"))
    Record.Item(mCol.标本紧急).Value = Nvl(rsTmp("标本紧急"))
    Record.Item(mCol.病人科室).Value = Nvl(rsTmp("病人科室"))
    Record.Item(mCol.申请科室).Value = Nvl(rsTmp("申请科室"))
    Record.Item(mCol.申请类型).Value = Nvl(rsTmp("申请类型"), 0)
    Record.Item(mCol.报告发送).Value = Nvl(rsTmp("报告发送"), 0)
    Record.Item(mCol.病人科室ID).Value = Nvl(rsTmp("病人科室ID"), 0)
    Record.Item(mCol.初审人).Value = Nvl(rsTmp("初审人"))
    Record.Item(mCol.初审时间).Value = Nvl(rsTmp("初审时间"))
    Record.Item(mCol.健康号).Value = Nvl(rsTmp("健康号"))
    Record.Item(mCol.审核未通过).Value = Nvl(rsTmp("审核未通过"))
    Record.Item(mCol.病人来源).Value = Nvl(rsTmp("病人来源"))
    Record.Item(mCol.门诊号).Value = Nvl(rsTmp("门诊号"))
    Record.Item(mCol.住院号).Value = Nvl(rsTmp("住院号"))
    If Nvl(rsTmp("项目")) = "任务团体" Then
        Record.Item(mCol.单位).Value = Nvl(rsTmp("内容"))
    End If
    Record.Item(mCol.结果为空).Value = Val(Nvl(rsTmp("结果为空")))
'    Record.Item(mCol.查阅状态).Value = Nvl(rsTmp("查阅状态"))

    '------晋煤新增
    For i = 0 To rptList.Columns.Count + 1
        If Val("" & rsTmp!微生物标本) = 0 Then
            If Record.Item(mCol.结果为空).Value > 0 Then
                Record.Item(i).BackColor = vbWhite
            Else
                Record.Item(i).BackColor = &HFDD6C6
            End If
        Else
            Record.Item(i).BackColor = vbWhite
        End If
    Next
    

    Me.rptList.Populate
    
    '过滤界面列表
    RptListFilter
    
    '没有临床路径病人时不显示列
    Me.rptList.Columns(6).Visible = blnPathPatient
    
    If blnGoto = True Then
        mfrmRequest.ZlCancel
        If Val(Record.Item(mCol.微生物标本).Value) = 1 Then
            mfrmWrite2.ZlCancel
        Else
            mfrmWrite.ZlCancel
        End If
    End If
    '重新定位到以前的位置
    If rptList.Rows.Count > 0 And lngRowIndex > 0 Then
'        Me.rptList.Rows(0).Selected = True
'        Me.rptList.Rows(0).EnsureVisible
        lngloop = 0

        For intLoop = 0 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = lngRowID Then
                lngloop = Me.rptList.Rows(intLoop).Index
                Exit For
            End If
        Next

        If lngRowIndex >= lngloop Then
            lngRowIndex = lngRowIndex - (lngRowIndex - lngloop)
        Else
            lngRowIndex = lngRowIndex + (lngloop - lngRowIndex)
        End If
        Me.rptList.Rows(lngRowIndex).EnsureVisible
    End If

    
    '不定位时退出
    If blnGoto = False Then
        If Not Me.rptList.FocusedRow Is Nothing Then
            If mlngKey = lngKey Then
                Call mfrmWrite.zlRefresh(mlngKey)
            End If
        End If
        Exit Sub
    End If
    '定位到行
    With Me.rptList
        For i = 0 To .Rows.Count - 1
            If .Rows(i).Record(mCol.医嘱id).Value = blnAdviceKey Or .Rows(i).Record(mCol.ID).Value = lngKey Then
                Set .FocusedRow = .Rows(i)
                Exit For
            End If
        Next
    End With
    
    If Me.TabList.Selected.Index = 0 Then
        Call SetControlFocus
    Else
        Call SetControlFocus
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub AuditingRegister()
    Dim i As Integer
    Dim strVerifyMan As String              '检验人
    Dim blnCancel As Boolean
    Dim strLogID As String
    '功能:          审核人注册,用于在没有审核人权限时登陆审核人后进行审核
    
    '清除以前的(为了保险)
    zlDatabase.SetPara "是否有具有审核权限", 0, 100, 1208
    zlDatabase.SetPara "审核人", 0, 100, 1208
                    
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            strVerifyMan = .Record(mCol.检验人).Value
        End With
    End If
                        
    frmLabAuditingLand.ShowMe Me, strVerifyMan, blnCancel, strLogID
    
    If blnCancel = True Then Exit Sub   '取消时不处理
    
    '得到是否授权
    i = zlDatabase.GetPara("是否有具有审核权限", 100, 1208, 0)
    mstrAuditingMan = zlDatabase.GetPara("审核人", 100, 1208, 0)
    If mstrAuditingMan = "0" Or mstrAuditingMan = "" Then mstrAuditingMan = ""
    If mstrAuditingMan <> "" Then
        mstrAuditingManID = strLogID
        Me.Caption = "检验技师工作站" & "-审核人(" & mstrAuditingMan & ")"
    Else
        mstrAuditingManID = ""
        Me.Caption = "检验技师工作站"
    End If
    '不=0时有权限变化,重新定义时间
'    If i <> 0 Then
        mintAuditing = i
        mDataAuditing = Now
'    End If
    '清除
    zlDatabase.SetPara "是否有具有审核权限", 0, 100, 1208
    zlDatabase.SetPara "审核人", 0, 100, 1208
End Sub
Private Function SaveDisposal(intDisposal As Integer) As Boolean
    '功能                   '对保存 取消进行操作
    Dim lngRetuId As Long                                        '调用其他窗体相应方法的返回值
    Dim Pane1 As Pane                                           '浮动窗体
    Dim rptRow As ReportRow                                     '列表记录集
    Dim intLoop As Integer                                      '当前行位置
    Dim lngLodKey As Long                                       '旧的ID
    
'    Me.SetFocus
    Select Case intDisposal
        Case mFileS.保存
            Select Case mintEditState
            Case 1, 2                                                           '登记,核收保存
                '在保前如果有保存后立即审核检查是否为同一个人审核
                If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                    If AuditionCheck = False Then Exit Function
                End If
                
                lngRetuId = mfrmRequest.ZlSave()
               
                If lngRetuId = 0 Then Exit Function
                
                If mintContinue = 0 Or Me.TabList.Selected.Index = 1 Then
                    '不连续操作
                    Me.rptList.Tag = ""   '清空连续保存的标记
                    mfrmRequest.ZlCancel: mlngKey = lngRetuId
                    

                    mintEditState = 0
                    If mlngMachineID > 0 Or mlngMachineID = -1 Then
                        InsertOneRecored lngRetuId, True
                    Else
                        Call RefreshData
                    End If
                    If Me.TabList.Item(1).Selected = True Then
                        Call RefreshData1
                    End If
                    '用于写入微生物的三级报告
                    Call ReportDisposal(mActR.写入病历)
                    '核收后是否发送仪器数据
                    Call SampleDisposal(mActS.发送仪器)
                    
'                    If Me.rptList.Visible Then
                        '11268 保存后审核参数勾了之后没有审核
                        '在已核收窗体中才支持此功能,待核收窗体中缺少很多信息,不能用这个功能
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                            Call ReportDisposal(mActR.审核报告)
                        End If
'                    End If
                    
                Else
                    If Me.rptList.Tag = "" Then
                        '第一次增加时先清除列表
                        Me.rptList.Records.DeleteAll
                        Me.rptList.Tag = "Continue"
                    End If
                    '添加刚新增的记录到列表中
                    InsertOneRecored lngRetuId, True
                    '核收后是否发送仪器数据
                    Call SampleDisposal(mActS.发送仪器)
                    
                    If Me.rptList.Visible Then
                        '11268 保存后审核参数勾了之后没有审核
                        '在已核收窗体中才支持此功能,待核收窗体中缺少很多信息,不能用这个功能
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                            Call ReportDisposal(mActR.审核报告)
                        End If
                    End If
                    
                    If Me.TabList.Selected.Index = 0 Then
                        '核收或登记
                        SampleDisposal IIf(mintEditState = 1, mActS.核收, mActS.登记)
                    Else
                        mintEditState = 0
                    End If
                End If
            Case 3                                                              '重新核收
                '在保前如果有保存后立即审核检查是否为同一个人审核
                If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                    If AuditionCheck = False Then Exit Function
                End If
                
                lngRetuId = mfrmRequest.ZlSave()
                mfrmRequest.ZlCancel
                If lngRetuId = 0 Then Exit Function
                mintEditState = 0
                Call RefreshData
            Case 4                                                              '补填病人
                If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                    If AuditionCheck = False Then Exit Function
                End If
                lngRetuId = mfrmRequest.ZlSave(mintEditState)
                If lngRetuId = 0 Then Exit Function
                mfrmRequest.ZlCancel
                If Me.TabList.Item(0).Selected = True Then
                    InsertOneRecored lngRetuId, False
                Else
                    Call RefreshData1
                End If
                mintEditState = 0

            Case 5                                                              '报告编辑
                If Val(Me.rptList.FocusedRow.Record(mCol.微生物标本).Value) <> 1 Then
                    If mfrmWrite.ZlSave() = True Then
                        mfrmWrite.ZlCancel
                        mintEditState = 0
                        Call mfrmWrite.zlRefresh(mlngKey)

                        '------晋煤新增
                        Dim strSQL As String, rsTmp As ADODB.Recordset, i As Integer
                        
                        If Not rptList.FocusedRow Is Nothing Then
                            strSQL = "Select Count(A.ID) - Sum(Decode(A.检验结果, Null, 0, 1)) As 无结果记录,Count(A.ID) as 结果数 " & vbNewLine & _
                                    "From 检验普通结果 A" & vbNewLine & _
                                    "Where A.检验标本id = [1]"
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
                            If rsTmp.EOF Then
                                For i = 0 To rptList.Columns.Count - 1
                                   rptList.FocusedRow.Record.Item(i).BackColor = vbWhite
                                Next
                            Else
                                If Val("" & rsTmp.Fields("无结果记录")) = 0 And Val("" & rsTmp.Fields("结果数")) > 0 Then
                                    For i = 0 To rptList.Columns.Count - 1
                                        rptList.FocusedRow.Record.Item(i).BackColor = &HFDD6C6     '&HC0FFFF
                                    Next
                                Else
                                    For i = 0 To rptList.Columns.Count - 1
                                        rptList.FocusedRow.Record.Item(i).BackColor = vbWhite
                                    Next
                                End If
                            End If
                        End If
                        '---------------

                    End If
                Else
                    If mfrmWrite2.ZlSave() = True Then
                        mfrmWrite2.ZlCancel
                        mintEditState = 0
                        Call mfrmWrite2.zlRefresh(mlngKey)
                        '用于写入微生物的三级报告
                        Call ReportDisposal(mActR.写入病历)
                    End If
                End If
            Case 6                      '填写三级报告
                If mfrmLabMicrobe3Report.ZlSave(mlngKey) = True Then
                    mintEditState = 0
                End If
            End Select
        
        Case mFileS.放弃
            Select Case mintEditState
                Case 1, 2, 3, 4
                    If mfrmRequest.ZlCancel() = False Then Exit Function
                    Me.rptList.Tag = ""         '清空连续保存的标记
                    mintEditState = 0
                    If Me.TabList.Selected.Index = 0 Then
                        Call RefreshData
                    Else
                        Call RefreshData1
                    End If
                Case 5
                    If Val(Me.rptList.FocusedRow.Record(mCol.微生物标本).Value) <> 1 Then
                        If mfrmWrite.ZlCancel = False Then Exit Function
                        mintEditState = 0
                        Call mfrmWrite.zlRefresh(mlngKey)
                    Else
                        If mfrmWrite2.ZlCancel = False Then Exit Function
                        mintEditState = 0
                        Call mfrmWrite2.zlRefresh(mlngKey)
                    End If
                Case 7
                    If mfrmLabMicrobe3Report.ZlCancel = True Then
                        mintEditState = 0
                    End If
                Case Else
            End Select
            Me.rptList.Tag = ""    '清空连续保存的标记
            mintEditState = 0
            
            If Not Me.rptList.FocusedRow Is Nothing Then
                InsertOneRecored mlngKey, False
            End If
            
            
                            
'            If Me.TabList.Item(0).Selected = True Then
'                Call RefreshData
'            Else
'                Call RefreshData1
'            End If
    End Select
    
    SaveDisposal = True
    
    On Error Resume Next
    Me.MousePointer = 0
    If Me.rptList.Tag = "" Then
        gintSelectFocus = 1
'        Me.cbo时间.SetFocus
        If Me.TabList.Selected.Index = 0 Then
'            Me.rptList.SetFocus
        Else
'            Me.rptList1.SetFocus
        End If
    End If
    If mintEditState = 0 Then
        Call ShowRequest(False)
    End If
End Function
Private Function SampleRefuse(lngKey As Long) As Boolean
    '标本取消核收(中间需要处理双向通讯)
    '参数                   传入医嘱ID
    Dim blnTran As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset, strQrySQL As String
    Dim strDevices As String, aDevice() As String, strAdviceIDs As String, i As Integer
    Dim intType As Integer                      '标本类别:0=普通，1=急诊
    Dim lngAdviceID As Long                     '医嘱ID
    Dim intEmerge As Integer                    '是否使用急诊标志

    If mlngKey = 0 Then Exit Function
    
    intEmerge = Val(zlDatabase.GetPara("急诊标本", 100, 1208, 0))
    
    On Error GoTo ErrHand

    Me.MousePointer = vbHourglass
    strAdviceIDs = "": strDevices = ""
    
    strSQL = "select distinct nvl(b.标本类别,0) as 标本类别,a.id as 医嘱Id " & _
             " from 病人医嘱记录 a,检验标本记录 b " & _
             " where a.id = b.医嘱ID and b.id = [1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)

    If rs.EOF = True Then Exit Function
    
    If rs.BOF = False Then
        intType = rs("标本类别")
        lngAdviceID = rs("医嘱ID")
            
    End If
    
    
    '处理双向通信
    If blnComm Then
        strAdviceIDs = strAdviceIDs & "," & lngAdviceID
        
        strQrySQL = "Select Distinct 仪器ID From 检验标本记录 A,检验项目分布 B" & _
            " Where B.医嘱ID=[1] And B.标本ID+0=A.ID"
        Set rs = zlDatabase.OpenSQLRecord(strQrySQL, Me.Caption, lngAdviceID)
        Do While Not rs.EOF
            If InStr(strDevices, "," & zlCommFun.Nvl(rs(0), 0)) = 0 Then
                strDevices = strDevices & "," & zlCommFun.Nvl(rs(0), 0)
            End If
            'CSBmk <Type the bookmark name here>
            rs.MoveNext
        Loop
    End If
    
    '处理双向通信
    If blnComm Then
        If Len(strDevices) > 0 Then strDevices = Mid(strDevices, 2)
        If Len(strAdviceIDs) > 0 Then strAdviceIDs = Mid(strAdviceIDs, 2)
        
        
        aDevice = Split(strDevices, ",")
        For i = 0 To UBound(aDevice)
            SendSample WinsockC, WinsockC.LocalIP, CLng(Val(aDevice(i))), "", 0, strAdviceIDs, True, IIf(intEmerge = 1, 0, intType)
        Next
    End If
    Me.MousePointer = vbDefault
    
    strSQL = "ZL_检验标本记录_取消核收(" & lngAdviceID & ")"
    zlDatabase.ExecuteProcedure strSQL, gstrSysName
        
    SampleRefuse = True
   
    Exit Function
    
ErrHand:
    
    Me.MousePointer = vbDefault
    If SampleRefuse = False Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
        
End Function

Private Sub TabList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intLoop As Integer
    On Error Resume Next
    If Me.Visible = False Then Exit Sub
    If Me.TabList.Item(1).Selected = True Then
        Me.TabCtlWindow.Item(0).Selected = True
        Me.TabCtlWindow.Item(2).Visible = False
        Me.TabCtlWindow.Item(3).Visible = False
        Me.TabCtlWindow.Item(4).Visible = False
        Me.TabCtlWindow.Item(5).Visible = False
        Me.TabCtlWindow.Item(6).Visible = False
        For intLoop = 2 To Me.chkSoure.UBound
            Me.chkSoure(intLoop).Visible = False
        Next
        '体检单独处理
        Me.chkSoure(5).Visible = True
        Me.chkSoure(5).Left = Me.chkSoure(2).Left
        Me.picFilter.Left = Me.chkSoure(2).Left + Me.chkSoure(2).Width + 30
        Call mfrmWrite.zlRefresh(-1)
        Call mfrmRequest.ZlCancel
        Call RefreshData1
        
        If mblnTabList1 = False Then
            cbo时间.Text = "今  天"
            mblnTabList1 = True
        Else
            cbo时间.Text = Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";", ";")(0)
            Me.dtpDate.Value = Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";" & Format(Now, "yyyy-mm-dd"), ";")(1)
            Me.dtpDateEnd.Value = Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";" & Format(Now, "yyyy-mm-dd") & ";" & Format(Now, "yyyy-mm-dd"), ";")(2)
        End If
        
        
        Call SetControlFocus
    Else
        Me.rptList.Tag = ""
        Me.TabCtlWindow.Item(0).Visible = True
        Me.TabCtlWindow.Item(0).Selected = True
        
        Me.TabCtlWindow.Item(2).Visible = True
        Me.TabCtlWindow.Item(3).Visible = IIf(Me.TabCtlWindow.Item(3).Tag = "费用查询", True, False)
        Me.TabCtlWindow.Item(4).Visible = True
        Me.TabCtlWindow.Item(5).Visible = True
        If Me.rptList.FocusedRow Is Nothing Then
            Me.TabCtlWindow.Item(6).Visible = False
            Me.TabCtlWindow.Item(7).Visible = False
        Else
            With Me.rptList
                If .Records(mCol.所属情况).Visible = "住院" Then
                    Me.TabCtlWindow.Item(6).Visible = False
                    Me.TabCtlWindow.Item(7).Visible = True
'                    Me.TabCtlWindow.Item(6).Selected = True
                Else
                    Me.TabCtlWindow.Item(6).Visible = True
                    Me.TabCtlWindow.Item(7).Visible = False
'                    Me.TabCtlWindow.Item(5).Selected = True
                End If
            End With
        End If
        Me.chkSoure(5).Left = 3780
        For intLoop = 0 To Me.chkSoure.UBound
            Me.chkSoure(intLoop).Visible = True
        Next
        Me.picFilter.Left = Me.chkSoure(5).Left + Me.chkSoure(5).Width + 30
'        Call RefreshData
        Call mfrmWrite.zlRefresh(mlngKey)
        Call mfrmRequest.zlRefresh(Me.rptList.FocusedRow)
        
        If mblnTabList1 = False Then
            cbo时间.Text = "今  天"
            mblnTabList1 = True
        Else
            cbo时间.Text = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";", ";")(0)
            Me.dtpDate.Value = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";" & Format(Now, "yyyy-mm-dd"), ";")(1)
            Me.dtpDateEnd.Value = Split(zlDatabase.GetPara("标本范围", 100, 1208, "今  天") & ";" & Format(Now, "yyyy-mm-dd") & ";" & Format(Now, "yyyy-mm-dd"), ";")(2)
        End If
        
        Call SetControlFocus
    End If
    
    If mintContinue = 1 Then
        Me.cbrthis.FindControl(, conMenu_Manage_Regist, , True).Caption = "连续登记"
        Me.cbrthis.FindControl(, conMenu_Manage_Plan, , True).Caption = "连续核收"
    Else
        Me.cbrthis.FindControl(, conMenu_Manage_Regist, , True).Caption = "登记"
        Me.cbrthis.FindControl(, conMenu_Manage_Plan, , True).Caption = "核收"
    End If
    Me.cbrthis.RecalcLayout
    Call picList_Click
End Sub

Private Sub txtGoto_GotFocus()
    Me.txtGoto.SelStart = 0
    Me.txtGoto.SelLength = Len(Me.txtGoto.Text)
End Sub

Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If FindPatient(txtGoto.Text) = True Then
'            txtGoto.Text = ""
'        Else
            Call FindPatient(txtGoto.Text)
            txtGoto.SelStart = 0
            txtGoto.SelLength = Len(txtGoto.Text)
'        End If
    End If
End Sub

Private Sub RptListFilter()
    '功能                   列表过滤(门诊病人;住院病人;无主标本;已审标本;未审标本;体检病人;紧急医嘱;紧急标本 进行快速过滤)
    Dim lngloop As Long
    Dim strSource As String             '来源
    Dim strExeState As String           '来行状态
    Dim strPatientName As String        '姓名
    Dim lngItemID As Long               '诊疗项目ID
    Dim lngCboItemID As Long            '选择的组合项目
    Dim int医嘱紧急 As Integer          '医嘱紧急
    Dim int标本紧急 As Integer          '标本紧急
    Dim str审核未通过 As String         '审核未通过
    Dim rsTmp As New ADODB.Recordset    '记录集
    Dim lngRowIndex As Long                                         '行索行
    Dim lngRowID As Long                                            '行ID
    Dim intLoop As Integer
    Dim i As Integer
    Dim str标本类型 As String
    Dim int结果为空 As Integer
    Dim strYiqiShenHe As String
    On Error Resume Next
    
    '刷新前记录一下位置
    If Not Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row Is Nothing Then
        lngRowIndex = Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row.Index - 1
        lngRowID = Me.rptList.Rows(lngRowIndex).Record(mCol.ID).Value
        mlngLastShow = lngRowID
    Else
        If mlngLastShow > 0 Then
            For i = 0 To Me.rptList.Rows.Count - 1
                If Me.rptList.Rows(i).Record(mCol.ID).Value = mlngLastShow Then
                    lngRowIndex = Me.rptList.Rows(i).Record.Index
                    lngRowID = Me.rptList.Rows(i).Record(mCol.ID).Value
                End If
            Next
        End If
    End If
        
'    If Me.rptList.Records.Count <= 0 And Me.rptList1.Records.Count <= 0 Then                           '没有记录时退出
'        Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptList.Rows.Count & "个病人．"
'        Exit Sub
'    End If
    If Me.TabList.Selected.Index = 0 Then
        If Me.rptList.Records.Count <= 0 Then
            Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptList.Rows.Count & "个病人．"
            Exit Sub
        End If
    Else
        If Me.rptList1.Records.Count <= 0 Then
            Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptList1.Rows.Count & "个病人．"
            Exit Sub
        End If
    End If
    With Me.rptList
        For lngloop = 0 To .Records.Count - 1
            .Records(lngloop).Visible = True
            
            strSource = .Records(lngloop).Item(mCol.所属情况).Value
            strExeState = .Records(lngloop).Item(mCol.执行状态).Value
            strPatientName = .Records(lngloop).Item(mCol.姓名).Value
            int医嘱紧急 = Val(.Records(lngloop).Item(mCol.医嘱紧急).Value)
            int标本紧急 = Val(.Records(lngloop).Item(mCol.标本紧急).Value)
            str标本类型 = Nvl(.Records(lngloop).Item(mCol.标本类型).Value)
            str审核未通过 = Nvl(.Records(lngloop).Item(mCol.审核未通过).Value)
            int结果为空 = Val(.Records(lngloop).Item(mCol.结果为空).Value)
            strYiqiShenHe = .Records(lngloop).Item(mCol.仪器审核).Value
            If str审核未通过 = "" Then
                .Records(lngloop).Visible = mblnVerifying(9) And .Records(lngloop).Visible
            End If

            If str审核未通过 <> "" Then
                .Records(lngloop).Visible = mblnVerifying(10) And .Records(lngloop).Visible
            End If

            
            '====门诊
            If strSource = "门诊" Or strSource = "院外" Then
                .Records(lngloop).Visible = mblnVerifying(0) And .Records(lngloop).Visible
            End If
            
            '=====体检
            If strSource = "体检" Then
                .Records(lngloop).Visible = mblnVerifying(5) And .Records(lngloop).Visible
            End If
            
            '====住院
            If strSource = "住院" Then
                .Records(lngloop).Visible = mblnVerifying(1) And .Records(lngloop).Visible
            End If
            
            '====无主
            If strSource = "无主" Then
                .Records(lngloop).Visible = mblnVerifying(2) And .Records(lngloop).Visible
            End If
            
            '==医嘱紧急
            If int医嘱紧急 = 1 Then
                .Records(lngloop).Visible = mblnVerifying(6) And .Records(lngloop).Visible
            End If
            
            If int标本紧急 = 1 Then
                .Records(lngloop).Visible = mblnVerifying(7) And .Records(lngloop).Visible
            End If
            
            If str标本类型 = "质控" Then
                .Records(lngloop).Visible = mblnVerifying(8) And .Records(lngloop).Visible
            End If
            
            If str审核未通过 = "" Then
                .Records(lngloop).Visible = mblnVerifying(9) And .Records(lngloop).Visible
            End If

            If str审核未通过 <> "" Then
                .Records(lngloop).Visible = mblnVerifying(10) And .Records(lngloop).Visible
            End If
            
            
            '已审核
            If strExeState = "已检验" Or strExeState = "已打印" Then
                .Records(lngloop).Visible = (mblnVerifying(3) = True And .Records(lngloop).Visible = True)
                
            End If
            
            If strExeState = "已传送" Or strExeState = "" Then
                .Records(lngloop).Visible = (mblnVerifying(4) = True And .Records(lngloop).Visible = True)
            End If
            
            '未完成的标本是否显示 by cd 2014-01-08
            If int结果为空 > 0 Then
                .Records(lngloop).Visible = (mblnVerifying(11) = True And .Records(lngloop).Visible = True)
            Else
                .Records(lngloop).Visible = (mblnVerifying(12) = True And .Records(lngloop).Visible = True)
            End If
            If strYiqiShenHe = "√" Then
                .Records(lngloop).Visible = mblnVerifying(13) And .Records(lngloop).Visible
            End If
            
            If strYiqiShenHe = "×" Then
                .Records(lngloop).Visible = mblnVerifying(14) And .Records(lngloop).Visible
            End If
        Next
        .Populate
        If Me.rptList.Rows.Count = 0 Then
            mfrmRequest.ZlCancel
            mfrmWrite2.ZlCancel
            mfrmWrite.ZlCancel
        End If
        Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptList.Rows.Count & "个病人．"
    End With
    
    '重新定位到以前的位置
    If rptList.Rows.Count > 0 And lngRowIndex > 0 Then
'        Me.rptList.Rows(0).Selected = True
'        Me.rptList.Rows(0).EnsureVisible
        lngloop = 0

        For intLoop = 0 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = lngRowID Then
                lngloop = Me.rptList.Rows(intLoop).Index
                Exit For
            End If
        Next

        If lngRowIndex >= lngloop Then
            lngRowIndex = lngRowIndex - (lngRowIndex - lngloop)
        Else
            lngRowIndex = lngRowIndex + (lngloop - lngRowIndex)
        End If
        Me.rptList.Rows(lngRowIndex).EnsureVisible
    End If
    
    
    With Me.rptList1
        If Me.TabList.Item(1).Selected = False Then Exit Sub
        If Me.rptList1.Records.Count <= 0 Then
            Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptList1.Rows.Count & "个病人．"
            Exit Sub
        End If
        For lngloop = 0 To .Records.Count - 1
            strSource = .Records(lngloop).Item(mRCol.来源).Value
            .Records(lngloop).Visible = True
            '====门诊
            If strSource = "门诊" Or strSource = "院外" Then
                .Records(lngloop).Visible = mblnWaitVerify(0) And .Records(lngloop).Visible
            End If
            '====住院
            If strSource = "住院" Then
                .Records(lngloop).Visible = mblnWaitVerify(1) And .Records(lngloop).Visible
            End If
            '====体检
            If strSource = "体检" Then
                .Records(lngloop).Visible = mblnWaitVerify(2) And .Records(lngloop).Visible
            End If
        Next
        .Populate

        If mlngMachineID = 0 Then Exit Sub
        If mlngMachineID = -1 Then
            gstrSql = "Select Distinct b.诊疗项目id" & vbNewLine & _
                      " From 检验仪器项目 a, 检验报告项目 b, 诊疗项目目录 c" & vbNewLine & _
                      " Where a.项目id = b.报告项目id And b.诊疗项目id = c.Id"
        Else
            gstrSql = "Select Distinct b.诊疗项目id" & vbNewLine & _
                      " From 检验仪器项目 a, 检验报告项目 b, 诊疗项目目录 c" & vbNewLine & _
                      " Where a.仪器id = [1] And a.项目id = b.报告项目id And b.诊疗项目id = c.Id"
        End If
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, mlngMachineID)

        For lngloop = 0 To .Records.Count - 1
            If .Records(lngloop).Visible = True Then
                .Records(lngloop).Visible = True
                lngItemID = Val(.Records(lngloop).Item(mRCol.诊疗项目ID).Value)
                rsTmp.filter = ""
                rsTmp.filter = "诊疗项目ID = " & lngItemID
                If mlngMachineID = -1 Then
                    If rsTmp.RecordCount > 0 Then .Records(lngloop).Visible = False
                Else
                    If rsTmp.RecordCount <= 0 Then .Records(lngloop).Visible = False
                    lngCboItemID = cboUnionItem.ItemData(cboUnionItem.ListIndex)
                    If lngCboItemID = 0 Then
                    
                    ElseIf lngCboItemID = -1 And .Records(lngloop).Item(mRCol.诊疗项目ID).Value = "" Then
                    
                    ElseIf Val(.Records(lngloop).Item(mRCol.诊疗项目ID).Value) = lngCboItemID Then
                        
                    Else
                        .Records(lngloop).Visible = False
                    End If
                End If
            End If
        Next


        .Populate
        Me.stbThis.Panels(2).Text = "当前共有：" & Me.rptList1.Rows.Count & "个病人．"
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    

End Sub

Private Sub mclsExpenses_StatusTextUpdate(ByVal bytType As Byte, ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    Me.stbThis.Panels(2).Text = Text
End Sub
Private Sub mfrmWrite_StartEdit(Cancel As Boolean)
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    On Error GoTo errH:
    If InStr(",7,8,13,", CStr(Me.rptList.FocusedRow.Record(mCol.执行状态).Icon)) > 0 Then
        '已检验
        Cancel = True
        mintHandleState = 0
    Else
        '还在进行登记核收补填时自动保存
        If mintEditState >= 1 And mintEditState <= 4 Then
            If Me.cbrthis.FindControl(, conMenu_Edit_Save, , True).Enabled = True And _
               Me.cbrthis.FindControl(, conMenu_Edit_Save, , True).Visible = True Then
                Call SaveDisposal(mFileS.保存)
            End If
        End If
        Select Case Me.rptList.FocusedRow.Record(mCol.标本类型).Icon
            Case 3
                If InStr(mstrPrivs, "修改质控结果") = 0 Then
                    Cancel = True
                    mintHandleState = 0
                Else
                    If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
                        Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
                        Cancel = False
                        ReportDisposal mActR.填写报告
                    End If
                End If
            Case 4
                If InStr(mstrPrivs, "修改比对结果") = 0 Then
                    Cancel = True
                    mintHandleState = 0
                Else
                    Cancel = False
                    If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
                        Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
                        ReportDisposal mActR.填写报告
                    End If
                End If
            Case Else
        '        mintHandleState = 2
                If Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Enabled = True And _
                        Me.cbrthis.FindControl(, conMenu_Manage_Report, , True).Visible = True Then
                    ReportDisposal mActR.填写报告
                    Cancel = False
                Else
                    Cancel = True
                End If
        End Select
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function AuditionCheck() As Boolean
    Dim strVerifyMan As String
    
    If Not rptList.FocusedRow Is Nothing Then
        With Me.rptList.FocusedRow
            strVerifyMan = .Record(mCol.检验人).Value
        End With
    End If

    If InStr(1, mstrPrivs, "审核标本") <= 0 Then
        '没有权限和其他用户登陆时退出
        MsgBox "你没有权限进行审核,请重新登陆具有审核人员进行审核!", vbInformation, gstrSysName
        Call SetControlFocus
        gintSelectFocus = 1
        Exit Function
    End If

    '有权限控制时
    If InStr(1, mstrPrivs, "审核限制") > 0 And strVerifyMan = UserInfo.姓名 Then
        '没有登陆审核人
        If mintAuditing = 0 Then
            '同一病人被权限控制不能进行审核
'            MsgBox "检验人和审核人为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
            Exit Function
        End If
        '审核时间是否过期
        If mintAuditing < 0 Then
            If DateDiff("h", mDataAuditing, Now) > Abs(mintAuditing) Then
'                MsgBox "审核有效时间已过,请重新登陆审核人!", vbInformation, gstrSysName
                '在有效时间段内可以进行审核
                Exit Function
            End If
        End If
        
        '判断登陆时后的审核人是否为同一人.
        If strVerifyMan = mstrAuditingMan Then
            '登陆进入的审核人和当前用户为同一个人
'            MsgBox "登陆进入的审核人和当前用户为同一个人,请使用其他用户登陆再试!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    AuditionCheck = True
    
End Function


Private Sub ShortWork(SWork As Integer)
    Dim lngRetuId As Long                       '返回结果

    '功能           快捷操作
    Select Case SWork
        Case mSWork.Key_Home, mSWork.Key_End
            Select Case mintEditState
                Case 0
                    If SWork = mSWork.Key_Home Then
                        BackOrNextPatient 1
                    Else
                        BackOrNextPatient 2
                    End If
                Case 1, 4, 5
                    If SaveDisposal(mFileS.保存) = True Then
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                            If AuditionCheck = False Then
                                Exit Sub
                            End If
                            Call ReportDisposal(mActR.审核报告)
                        Else
                            If SWork = mSWork.Key_End Then
                                If MoveStation(1, 2) = False Then                       '向下移动
                                    '没有找到记录时退出操作
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.放弃)
                                    Exit Sub
                                End If
                            Else
                                If MoveStation(0, 2) = False Then                      '向上移动
                                    '没有找到记录时退出操作
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.放弃)
                                    Exit Sub
                                End If
                            End If
                        End If
                        If mintHandleState = 1 Then
                            Call SampleDisposal(mActS.补填病人)
                        Else
                            Call ReportDisposal(mActR.填写报告)
                        End If
                    End If
            End Select
        Case mSWork.Key_PageDown, mSWork.Key_PageUP
            Select Case mintEditState
                Case 0
                    If SWork = mSWork.Key_PageUP Then
                        BackOrNextPatient 1
                    Else
                        BackOrNextPatient 2
                    End If
                Case 1, 2                    '登记时也保存
                    SaveDisposal (mFileS.保存)
                Case 4                     '补填病人
                    If SaveDisposal(mFileS.保存) = True Then
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = False Then
                            Call ReportDisposal(mActR.填写报告)
                        Else
                            If SWork = mSWork.Key_PageDown Then
                                If MoveStation(1, 2) = False Then                       '向下移动
                                    '没有找到记录时退出操作
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.放弃)
                                    Exit Sub
                                End If
                            Else
                                If MoveStation(0, 2) = False Then                       '向上移动
                                    '没有找到记录时退出操作
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.放弃)
                                    Exit Sub
                                End If
                            End If
                            Call SampleDisposal(mActS.补填病人)
                        End If
                    End If
                    
                Case 5                      '填写报告
                    If SaveDisposal(mFileS.保存) = True Then
                        '保存后审核
                        If Me.cbrChild.ActiveMenuBar.FindControl(, conMenu_Manage_RequestBatPrint, True, True).Checked = True Then
                            If AuditionCheck = False Then
                                Exit Sub
                            End If
                            Call ReportDisposal(mActR.审核报告)
                        Else
                            If SWork = mSWork.Key_PageDown Then
                                If MoveStation(1, 2) = False Then                       '向下移动
                                    '没有找到记录时退出操作
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.放弃)
                                    Exit Sub
                                End If
                            Else
                                If MoveStation(0, 2) = False Then                       '向上移动
                                    '没有找到记录时退出操作
                                    mintHandleState = 0
                                    mintEditState = 0
                                    Call SaveDisposal(mFileS.放弃)
                                    Exit Sub
                                End If
                            End If
                        End If
                        If mintHandleState = 1 Then
                            Call SampleDisposal(mActS.补填病人)
                        Else
                            Call ReportDisposal(mActR.填写报告)
                        End If
                        
                    End If
            End Select
    End Select
    
End Sub
Private Function MoveStation(BackOrNext As Integer, Optional intState As Integer) As Boolean
    '功能               移动到上一个或下一个记录
    '参数               BackOrNext =0 向上 = 1 向下
    '                   intState 查下一条的状态 0 = 下一条未审核记录 = 1 下一条无主 =2 下一条
    
    Dim NowRow As Long
    Dim lngloop As Long

    If Me.rptList.Rows.Count = 0 Then Exit Function
    If Me.rptList.FocusedRow Is Nothing Then Exit Function

    NowRow = Me.rptList.FocusedRow.Index
    

    With Me.rptList

        If BackOrNext = 1 Then
            If NowRow + 1 = .Rows.Count Then Exit Function
            For lngloop = NowRow + 1 To .Rows.Count - 1
                If intState = 0 Then
                    If Val(.Rows(lngloop).Record(mCol.医嘱id).Value) > 0 And .Rows(lngloop).Record(mCol.审核人).Value = "" Then
                        Set .FocusedRow = .Rows(lngloop)
                        .Populate
                        MoveStation = True
                        Exit Function
                    End If
                ElseIf intState = 1 Then
                    If .Rows(lngloop).Record(mCol.姓名).Value = "" And .Rows(lngloop).Record(mCol.标本类型).Value = "" Then
                        Set .FocusedRow = .Rows(lngloop)
                        .Populate
                        MoveStation = True
                        Exit Function
                    End If
                ElseIf intState = 2 Then
                    Set .FocusedRow = .Rows(lngloop)
                    .Populate
                    MoveStation = True
                    Exit Function
                End If
            Next
        Else
            If NowRow - 1 = -1 Then Exit Function
            For lngloop = NowRow - 1 To 0 Step -1
                If intState = 0 Then
                    If Val(.Rows(lngloop).Record(mCol.医嘱id).Value) > 0 And .Rows(lngloop).Record(mCol.审核人).Value = "" Then
                        Set .FocusedRow = .Rows(lngloop)
                        .Populate
                        MoveStation = True
                        Exit Function
                    End If
                ElseIf intState = 1 Then
                    If .Rows(lngloop).Record(mCol.姓名).Value = "" And .Rows(lngloop).Record(mCol.标本类型).Value = "" Then
                        Set .FocusedRow = .Rows(lngloop)
                        .Populate
                        MoveStation = True
                        Exit Function
                    End If
                ElseIf intState = 2 Then
                    Set .FocusedRow = .Rows(lngloop)
                    .Populate
                    MoveStation = True
                    Exit Function
                End If
            Next
        End If
    End With

    
End Function
Public Sub zlRefreshData()
    '刷新数据
    Call RefreshData
End Sub
Private Sub RefreshData1()
    '''''''''''''''''''''''''''''''''''''''''
    '功能           刷待处理列表
    '''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    Dim Record As ReportRecord
    Dim strSQL As String
    Dim lngAdviceID As Long
    Dim lngCorrelation As Long
    Dim intLoop As Integer
    Dim strStart As String
    Dim strEnd As String
    Dim strDeptID As String
    On Error GoTo errH

    
    strStart = GetDateTime(Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";", ";")(0), 1)
    strEnd = GetDateTime(Split(zlDatabase.GetPara("待核收范围", 100, 1208, "今  天") & ";", ";")(0), 2)
    
    If strStart = "自定义" Then
        strStart = Format(Me.dtpDate.Value, "yyyy-mm-dd 00:00:00")
        strEnd = Format(Me.dtpDateEnd.Value, "yyyy-mm-dd 23:59:59")
    Else
        If strStart = "" Then strStart = GetDateTime("今  天", 1)
        If strEnd = "" Then strEnd = GetDateTime("今  天", 2)
    End If
    Me.rptList1.Records.DeleteAll
    
    gstrSql = "Select /*+ Rule */ a.id,a.相关ID,a.病人id,a.紧急标志,decode(a.病人来源,1,'门诊',2,'住院',3,'院外',4,'体检') As 病人来源," & vbNewLine & _
            "       d.姓名,d.性别,d.年龄,e.名称 As 病人科室," & vbNewLine & _
            "       decode(a.病人来源,1,d.门诊号,2,d.住院号,4,d.门诊号) As 标识号," & vbNewLine & _
            "       Decode(a.病人来源,2,S.病人类型,d.病人类型) as 病人类型," & vbNewLine & _
            "       d.当前床号 As 床号,a.医嘱内容, a.开嘱医生 , a.开嘱时间,a.诊疗项目ID,b.执行状态,a.挂号单,b.接收时间 " & vbNewLine & _
            "From 病人医嘱记录 a , 病人医嘱发送 b ,诊疗项目目录 c , 病人信息 d ,部门表 e,病案主页 s" & vbNewLine & _
            "Where a.Id = b.医嘱id And b.执行状态 in (0,2) And a.相关ID Is Not Null" & vbNewLine & _
            "     And c.Id = a.诊疗项目ID And c.类别 = 'C' And a.开嘱时间 Between [2] And [3]" & vbNewLine & _
            "     And a.病人id = d.病人id And a.病人科室id = e.Id [条件]  " & vbNewLine & _
            " and a.病人ID = S.病人ID(+) and a.主页ID = s.主页ID(+)  " & vbNewLine & _
            "Order By a.Id , a.相关ID ,开嘱时间 "
    
    
    If mlngDeptID > 0 And rptList.Tag = "" Then
        strDeptID = mlngDeptID
    Else
        If InStr(mstrPrivs, "所有科室") = 0 Or InStr(mstrPrivs, "查看其他科室报告") > 0 Then
            For intLoop = 1 To Me.cboDept.ListCount - 1
                strDeptID = strDeptID & "," & Me.cboDept.ItemData(intLoop)
            Next
        End If
    End If
    gstrSql = Replace(gstrSql, "[条件]", " And A.执行科室id In (Select * From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) ")

    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strDeptID, CDate(Format(strStart, "yyyy-MM-dd HH:mm:ss")), _
                                        CDate(Format(strEnd, "yyyy-MM-dd HH:mm:ss")))
    
    Do Until rsTmp.EOF
        With Me.rptList1
            If rsTmp("相关ID") <> lngCorrelation Then
                Set Record = .Records.Add
                For intLoop = 0 To .Columns.Count
                    Record.AddItem ""
                Next
                Record(mRCol.病人ID).Value = Nvl(rsTmp("病人ID"))
                Record(mRCol.紧急).Icon = IIf(Val(Nvl(rsTmp("紧急标志"))) = 1, 0, -1)    '1=紧急
                Record(mRCol.来源).Value = Nvl(rsTmp("病人来源"))
                Record(mRCol.姓名).Value = Nvl(rsTmp("姓名")) & IIf(Nvl(rsTmp("执行状态")) = 2, "(拒收)", "")
                Record(mRCol.年龄).Value = Nvl(rsTmp("年龄"))
                Record(mRCol.性别).Value = Nvl(rsTmp("性别"))
                Record(mRCol.病人科室).Value = Nvl(rsTmp("病人科室"))
                Record(mRCol.标识号).Value = Nvl(rsTmp("标识号"))
                Record(mRCol.床号).Value = Nvl(rsTmp("床号"))
                Record(mRCol.医嘱内容).Value = Nvl(rsTmp("医嘱内容"))
                Record(mRCol.开嘱医生).Value = Nvl(rsTmp("开嘱医生"))
                Record(mRCol.开嘱时间).Value = Nvl(rsTmp("开嘱时间"))
                Record(mRCol.诊疗项目ID).Value = Nvl(rsTmp("诊疗项目ID"))
                Record(mRCol.医嘱id).Value = Nvl(rsTmp("相关ID"))
                Record(mRCol.执行状态).Value = Nvl(rsTmp("执行状态"))
                Record(mRCol.挂号单).Value = Nvl(rsTmp("挂号单"))
                Record(mRCol.签收时间).Value = Nvl(rsTmp("接收时间"))
                If Nvl(rsTmp("病人类型")) <> "" Then
                    Record(mRCol.姓名).ForeColor = zlDatabase.GetPatiColor(Nvl(rsTmp("病人类型")), False)
                End If
            Else
                Record(mRCol.医嘱内容).Value = Record(mRCol.医嘱内容).Value & " " & Nvl(rsTmp("医嘱内容"))
            End If
            If Nvl(rsTmp("执行状态")) = 2 Then
                For intLoop = 0 To .Columns.Count
                    Record(intLoop).ForeColor = vbRed
                Next
            End If
            lngCorrelation = Val(Nvl(rsTmp("相关ID")))
        End With
        rsTmp.MoveNext
    Loop
    If Me.TabList.Selected.Index = 0 Then
'        Me.rptList.SetFocus
    Else
'        Me.rptList1.SetFocus
    End If
    Me.rptList1.Populate
    Call RptListFilter
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub QuickFindPatient()
    '快速查找当前病人的历次检验
    Dim strPatient As String                        '病人姓名
    Dim lngPatientID As Long                        '病人ID
    Dim strStart As String                          '开始时间
    Dim strEnd As String                            '结束时间
    Dim i As Integer                                '用于判断是否使用病人ID
    
    On Error Resume Next
    If Me.TabList.Item(0).Selected = False Or Me.rptList.FocusedRow Is Nothing Then Exit Sub
    
    With Me.rptList.FocusedRow
        If .Record(mCol.姓名).Value = "" Then Exit Sub
        strPatient = .Record(mCol.姓名).Value
        lngPatientID = .Record(mCol.病人ID).Value
        i = zlDatabase.GetPara("历史病人识别", 100, 1208, 0)
        strStart = GetDateTime(zlDatabase.GetPara("历次检验范围", 100, 1208, "本  月"), 1)
        strEnd = GetDateTime(zlDatabase.GetPara("历次检验范围", 100, 1208, "本  月"), 2)
        Me.rptList.Tag = strPatient & ";;,;;;;;;;,True;" & strStart & "," & strEnd & ";0;;0;;;;1;;" & IIf(i = 0, lngPatientID, "0")
        Call RefreshData
    End With
    
End Sub






Private Sub DelItem(lngKey As Long)
    '功能           '删除指定的记录
    Dim intLoop As Integer
    Dim lngloop As Integer
    Dim lngRowIndex As Long                                         '行索行
    Dim lngRowID As Long                                            '行ID
    
    
    '刷新前记录一下位置
    If Not Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row Is Nothing Then
        lngRowIndex = Me.rptList.HitTest(5, Me.rptList.Height - mcontIntRowHeight).Row.Index - 1
        lngRowID = Me.rptList.Rows(lngRowIndex).Record(mCol.ID).Value
        mlngLastShow = lngRowID
    Else
        If mlngLastShow > 0 Then
            For intLoop = 0 To Me.rptList.Rows.Count - 1
                If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = mlngLastShow Then
                    lngRowIndex = Me.rptList.Rows(intLoop).Record.Index
                    lngRowID = Me.rptList.Rows(intLoop).Record(mCol.ID).Value
                End If
            Next
        End If
    End If
    
    With Me.rptList
        For intLoop = 0 To .Records.Count - 1
            If .Records(intLoop).Item(mCol.ID).Value = lngKey Then
                .Records.RemoveAt (intLoop)
                .Populate
                Exit For
            End If
        Next
    End With
    
    '重新定位到以前的位置
    If rptList.Rows.Count > 0 And lngRowIndex > 0 Then
'        Me.rptList.Rows(0).Selected = True
'        Me.rptList.Rows(0).EnsureVisible
        lngloop = 0

        For intLoop = 0 To Me.rptList.Rows.Count - 1
            If Me.rptList.Rows(intLoop).Record(mCol.ID).Value = lngRowID Then
                lngloop = Me.rptList.Rows(intLoop).Index
                Exit For
            End If
        Next

        If lngRowIndex >= lngloop Then
            lngRowIndex = lngRowIndex - (lngRowIndex - lngloop)
        Else
            lngRowIndex = lngRowIndex + (lngloop - lngRowIndex)
        End If
        Me.rptList.Rows(lngRowIndex).EnsureVisible
    End If
End Sub
Public Function ReadImageData(lngKeyID As Long, blnSave As Boolean) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim DrawIndex As Integer
    Dim strTime As Date
    Dim strErr As String
    Static objImg As Object
        
    On Error GoTo errH
    strTime = Now
    gstrSql = "select id ,标本ID,图像类型 from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngKeyID)
    '图像排版
    ImageTypeSet rsTmp.RecordCount - 1, True
    '不显示时不更新
    If Me.cbrthis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked = True Then Exit Function
    
    If objImg Is Nothing Then Set objImg = CreateObject("zlLisDev.clsDrawGraph")
    objImg.GetSampleImgInit glngSys, gcnOracle, strErr
    Call objImg.GetSampleImages(lngKeyID, App.path, False, strErr)
    Do Until rsTmp.EOF
        If Dir(App.path & "\" & lngKeyID & "_" & rsTmp("图像类型") & ".cht") = "" Then
            If Dir(App.path & "\" & lngKeyID & "_" & rsTmp("图像类型") & ".cht") <> "" Then
                Me.ChartThis(DrawIndex).Load App.path & "\" & lngKeyID & "_" & rsTmp("图像类型") & ".cht"
                If blnSave Then
                    Kill App.path & "\" & lngKeyID & "_" & rsTmp("图像类型") & ".cht"
                End If
            End If
        Else
            Me.ChartThis(DrawIndex).Load App.path & "\" & lngKeyID & "_" & rsTmp("图像类型") & ".cht"
            If blnSave Then
                Kill App.path & "\" & lngKeyID & "_" & rsTmp("图像类型") & ".cht"
            End If
        End If
        DrawIndex = DrawIndex + 1
        rsTmp.MoveNext
    Loop
    ReadImageData = True
'    Debug.Print "ID=" & lngKeyID & ",用时:" & DateDiff("s", strTime, Now)
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub VScroll1_Change()
'    Me.PicImage.Top = -200
End Sub

Private Sub VScroll_Change()
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub
    For intLoop = 0 To Me.VScroll.Max
        If intLoop < Me.VScroll.Value Then
            Me.ChartThis(intLoop).Visible = False
        Else
            Me.ChartThis(intLoop).Visible = True
            If intLoop = Me.VScroll.Value Then
                Me.ChartThis(intLoop).Top = 0
            Else
                Me.ChartThis(intLoop).Top = Me.ChartThis(intLoop - 1).Top + Me.ChartThis(intLoop - 1).Height + 10
            End If
        End If
    Next
End Sub

Private Sub ImageTypeSet(intCount As Integer, Optional blnReset As Boolean = False)
    '功能           对检验图像进行排版
    '参数           intCount = 图像数
    '               blnReset = 是否需要重新读入
    Dim intLoop As Integer
    Dim Pane5 As Pane
    
'    If blnReset = True Then
'        For intLoop = Me.ChartThis.UBound To 1 Step -1
''            Me.ChartThis(Me.ChartThis.UBound).ChartGroups(1).Data.NumSeries = 0
''            Me.ChartThis(Me.ChartThis.UBound).Header.Text = ""
'            If intLoop <> 0 Then
'                Unload Me.ChartThis(Me.ChartThis.UBound)
'            End If
'        Next
'    End If
    
    On Error Resume Next
    
    For intLoop = 0 To intCount
        If intLoop = 0 Then
            With Me.ChartThis(intLoop)
                .Interior.Image.LayOut = oc2dImageStretched
                .Visible = True
                .Top = 0
                .Left = 0
                .Width = IIf(Me.PicImage.ScaleWidth - Me.VScroll.Width - 20 <= 300, 300, Me.PicImage.ScaleWidth - Me.VScroll.Width - 20)
                .Height = .Width
            End With
        Else
            If blnReset = True And Me.ChartThis.UBound < intLoop Then
                Load Me.ChartThis(intLoop)
            End If
            With Me.ChartThis(intLoop)
'                .ChartGroups(1).Data.NumSeries = intLoop
'                .ChartGroups(1).Data.NumPoints(intLoop) = intLoop
                .Interior.Image.LayOut = oc2dImageStretched
                .Visible = True
                .Top = Me.ChartThis(intLoop - 1).Top + Me.ChartThis(intLoop - 1).Height + 10
                .Left = 0
                .Width = Me.ChartThis(intLoop - 1).Width
                .Height = .Width
                .IsBatched = False
            End With
        End If
    Next
    
    '隐藏多余的Chart控件
    For intLoop = intCount + 1 To Me.ChartThis.UBound
        Me.ChartThis(intLoop).Visible = False
    Next
    
    Set Pane5 = Me.dkpMain.FindPane(Dkp_ID_Image)
    If Not Pane5 Is Nothing Then
        If intCount < 0 Then
            Pane5.Close
        Else
            If Me.cbrthis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked = False Then
                Pane5.Select
            Else
                Pane5.Close
            End If
        End If
    End If
    With Me.VScroll
        .Top = 0
        .Left = Me.PicImage.ScaleWidth - .Width - 10
        .Height = Me.PicImage.ScaleHeight
        .Max = intCount
        .SmallChange = 1
        .LargeChange = 1
    End With
End Sub


Private Function CheckPatientInfo(lngSampleID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim int提示修正 As Integer '1-提示修正，2-不提示修正，3-不修正

    On Error GoTo errH
    
    gstrSql = "Select A.病人来源,A.病人id, A.性别 As 性别1, B.性别 As 性别2, A.年龄 As 年龄1, B.年龄 As 年龄2, A.姓名 As 姓名1, B.姓名 As 姓名2,nvl(a.婴儿,0) as 婴儿 " & vbNewLine & _
                        "From 检验标本记录 A, 病人信息 B" & vbNewLine & _
                        "Where A.病人id = B.病人id And A.ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngSampleID)
    
    '是婴儿时不进行对比
    If rsTmp("婴儿") > 0 Then
        Exit Function
    End If
    
    
    If Nvl(rsTmp("姓名1")) <> Nvl(rsTmp("姓名2")) Or Nvl(rsTmp("性别1")) <> Nvl(rsTmp("性别2")) Or _
        Nvl(rsTmp("年龄1")) <> Nvl(rsTmp("年龄2")) Then
        
        int提示修正 = 1
        
        If rsTmp("病人来源") = 4 Then
            int提示修正 = int体检处理方式
        ElseIf rsTmp("病人来源") = 3 Then
            int提示修正 = int院外处理方式
        ElseIf rsTmp("病人来源") = 2 Then
            int提示修正 = int住院处理方式
        ElseIf rsTmp("病人来源") = 1 Then
            int提示修正 = int门诊处理方式
        End If
        
        If int提示修正 = 1 Then
            If MsgBox("发现检验信息中的病人信息和病人信息中病人信息不一致!" & vbCrLf & "是否需要修正?", _
                vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                gstrSql = "zl_检验标本记录_Update(" & lngSampleID & ",'" & Nvl(rsTmp("姓名2")) & "','" & Nvl(rsTmp("性别2")) & _
                                             "','" & Nvl(rsTmp("年龄2")) & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            End If
        ElseIf int提示修正 = 2 Then
            gstrSql = "zl_检验标本记录_Update(" & lngSampleID & ",'" & Nvl(rsTmp("姓名2")) & "','" & Nvl(rsTmp("性别2")) & _
                                         "','" & Nvl(rsTmp("年龄2")) & "')"
            zlDatabase.ExecuteProcedure gstrSql, Me.Caption
        End If
        CheckPatientInfo = True
        Exit Function
    End If
    CheckPatientInfo = False
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub DelButton(Index As Integer)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    '功能       显示或隐藏按钮
    Dim lngCount As Long
    On Error Resume Next
    '费用查询
    Me.cbrthis.FindControl(, conMenu_EditPopup).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Append).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_NewItem).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Modify).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Delete).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ChargeDelApply).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ChargeDelAudit).Delete
    Me.cbrthis.FindControl(, conMenu_ToolPopup).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_Option).Delete
    Me.cbrthis.FindControl(, conMenu_ToolPopup).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ExtraFeeMove).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ExtraFeeExe).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_ExtraFeeUnExe).Delete
    '医嘱记录
    Me.cbrthis.FindControl(, conMenu_Edit_NewItem).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Modify).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Delete).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Blankoff).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Stop).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Send).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Untread).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_Compend).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_MarkMap).Delete
    Me.cbrthis.FindControl(, conMenu_Edit_MarkKeyMap).Delete
    Me.cbrthis.FindControl(, conMenu_Manage_ReportLisView).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_Sign).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_SignNew).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_SignVerify).Delete
    Me.cbrthis.FindControl(, conMenu_Tool_SignEarse).Delete
    Me.cbrthis.FindControl(, conMenu_View_Append, , True).Delete
    Me.cbrthis.FindControl(, conMenu_View_Hide, , True).Delete
    Me.cbrthis.FindControl(, conMenu_Report_ClinicBill, , True).Delete
    Me.cbrthis.FindControl(, conMenu_View_FontSize, , True).Delete
    Me.cbrthis.FindControl(, conMenu_View_FontSize_S, , True).Delete
    Me.cbrthis.FindControl(, conMenu_View_FontSize_L, , True).Delete
    
    '如果没有 审报告 按钮,则新增一个 审报告 按钮
    Set objBar = Me.cbrthis(2)
    With objBar.Controls
        Set objControl = .Find(, conMenu_Edit_Audit)
        If objControl Is Nothing Then
            Set objControl = .Find(, conMenu_Manage_Report) '从填报告按钮之后开始加入
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "审报告", objControl.Index + 1): objControl.BeginGroup = True   '增加按钮
            objControl.ID = conMenu_Edit_Audit: objPopup.IconId = conMenu_Edit_Audit    '赋值ID和图标
            objControl.Style = xtpButtonIconAndCaption  '同时显示问题和图片
        End If
    End With

'    With Me.cbrthis.KeyBindings
'        .Add FCONTROL, Asc("P"), conMenu_File_Print
'        .Add 0, VK_F2, conMenu_Edit_Save
'        .Add 0, VK_ESCAPE, conMenu_LIS_Cancel
'        .Add 0, VK_F12, conMenu_File_Parameter
'        .Add 0, VK_F4, conMenu_Manage_Plan
'        .Add 0, VK_F8, conMenu_Manage_Regist
'        .Add FCONTROL, Asc("T"), conMenu_Tool_Apply
'        .Add FCONTROL, Asc("Z"), conMenu_Edit_SendBack
'        .Add FCONTROL, VK_DELETE, conMenu_Manage_ClearUp
'        .Add 0, VK_F7, conMenu_Manage_Report
'        .Add 0, VK_F6, conMenu_Edit_Audit
'        .Add FCONTROL, VK_LEFT, conMenu_View_Backward
'        .Add FCONTROL, VK_RIGHT, conMenu_View_Forward
'        .Add 0, VK_F1, conMenu_Help_Help
'        .Add 0, VK_F5, conMenu_View_Refresh
'        .Add FCONTROL, Asc("F"), conMenu_Manage_Transfer_Force
'        .Add 0, VK_F3, conMenu_View_Filter
'        .Add 0, VK_HOME, conMenu_Tool_MeetFinish
'        .Add 0, VK_END, conMenu_Tool_MeetCancel
'        .Add 0, VK_PAGEUP, conMenu_Tool_Reference_1
'        .Add 0, VK_PAGEDOWN, conMenu_Tool_Reference_2
'        .Add FCONTROL, Asc("H"), conMenu_View_FindNext
'        .Add 0, VK_F9, conMenu_Edit_QCRes
'        .Add 0, VK_F11, conMenu_Manage_Logout
'    End With
'
'    If Index = 3 Then
'        '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
'        End If
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "费用(&M)", objMenu.Index + 1, False)
'        objMenu.ID = conMenu_EditPopup
'        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "生成主费用(&N)")
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "补充费加费用(&A)"): objPopup.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改附加费用(&M)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除附加费用(&D)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "销帐申请(&L)"): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelAudit, "销帐审核(&U)")
'        End With
'
'        '工具菜单:主窗体可能没有,放在帮助菜单前面
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
'            objMenu.ID = conMenu_ToolPopup
'        End If
'        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "医嘱附费选项(&O)"): objControl.BeginGroup = True
'            objControl.IconId = conMenu_File_Parameter
'        End With
'
'        '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
'        '-----------------------------------------------------
'        Set objBar = cbrthis(2)
'        For Each objControl In objBar.Controls '先求出前面的最后一个Control
'            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
'                Set objControl = objBar.Controls(objControl.Index - 1): Exit For
'            End If
'        Next
'        With objBar.Controls
'            'Set objControl = .Find(, conMenu_File_Preview) '从预览按钮之后开始加入
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Append, "主费", objControl.Index + 1): objControl.BeginGroup = True
'            Set objPopup = .Add(xtpControlPopup, conMenu_Edit_NewItem, "补费", objControl.Index + 1): objPopup.BeginGroup = True
'                objPopup.ID = conMenu_Edit_NewItem: objPopup.IconId = conMenu_Edit_NewItem
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "改费", objPopup.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删费", objControl.Index + 1)
'        End With
'
'        '命令的快键绑定
'        '-----------------------------------------------------
'        With cbrthis.KeyBindings
'            .Add FCONTROL, vbKeyE, conMenu_Edit_Append '生成主费用
'            .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改附加费用
'            .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除附加费用
'        End With
'
'        '设置不常用命令
'        '-----------------------------------------------------
'        With cbrthis.Options
'        End With
'    End If
'
'    If Index = 5 Then
'        '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
'        End If
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "医嘱(&A)", objMenu.Index + 1, False)
'        objMenu.ID = conMenu_EditPopup
'        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新开医嘱(&A)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改医嘱(&M)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除医嘱(&D)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "医嘱发送(&G)"): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "医嘱作废(&B)")
'
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)")
'            objPopup.BeginGroup = True
'            objPopup.IconId = conMenu_Manage_Report
'
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片处理(&V)")
'        End With
'
'        '查看菜单
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'        With objMenu.CommandBar.Controls
'            Set objControl = .Find(, conMenu_View_StatusBar) '状态栏项后
'            Set objControl = .Add(xtpControlButton, conMenu_View_Append, "附加信息(&A)", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "自动隐藏过滤条件栏(&H)", objControl.Index + 1)
'        End With
'
'        '工具菜单:主窗体可能没有,放在帮助菜单前面
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
'            objMenu.ID = conMenu_ToolPopup
'        End If
'        With objMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "电子签名(&S)", -1, False): objPopup.BeginGroup = True
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "医嘱签名(&I)")
'                objControl.IconId = conMenu_Tool_Sign
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)"): objControl.BeginGroup = True
'            End With
'
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "门诊医嘱选项(&O)"): objControl.BeginGroup = True
'            objControl.IconId = conMenu_File_Parameter
'
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "成套方案定义(&O)"): objControl.BeginGroup = True
'        End With
'
'        '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
'        '-----------------------------------------------------
'        Set objBar = cbrthis(2)
'        For Each objControl In objBar.Controls '先求出前面的最后一个Control
'            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
'                Set objControl = objBar.Controls(objControl.Index - 1): Exit For
'            End If
'        Next
'        With objBar.Controls
'            'Set objControl = .Find(, conMenu_File_Preview) '从预览按钮之后开始加入
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新开", objControl.Index + 1): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送", objControl.Index + 1): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "签名", objControl.Index + 1): objControl.BeginGroup = True
'            objControl.IconId = conMenu_Tool_Sign
'        End With
'
'        '命令的快键绑定
'        '-----------------------------------------------------
'        With cbrthis.KeyBindings
'            .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新开医嘱
'            .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改医嘱
'            .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除医嘱
'            .Add FCONTROL, vbKeyG, conMenu_Edit_Send '医嘱发送
'
'            .Add FCONTROL, vbKeyR, conMenu_Edit_Compend * 10# + 1 '查阅报告
'            .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '观片处理
'
'            .Add FCONTROL, vbKeyH, conMenu_View_Hide '自动隐藏过滤条件栏
'
'            .Add 0, vbKeyF11, conMenu_Tool_Option '医嘱选项
'        End With
'
'        '设置不常用命令
'        '-----------------------------------------------------
'        With cbrthis.Options
'        End With
'    End If
'
'    If Index = 6 Then
'        '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
'        End If
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "医嘱(&A)", objMenu.Index + 1, False)
'        objMenu.ID = conMenu_EditPopup
'        With objMenu.CommandBar.Controls
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新开医嘱(&A)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改医嘱(&M)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除医嘱(&D)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "医嘱作废(&B)"): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "医嘱停止(&S)")
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "临嘱发送(&G)"): objControl.BeginGroup = True
'            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Untread, "医嘱回退(&L)")
'
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)")
'            objPopup.BeginGroup = True
'            objPopup.IconId = conMenu_Manage_Report
'
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片处理(&V)")
'        End With
'
'        '查看菜单
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
'        With objMenu.CommandBar.Controls
'            Set objControl = .Find(, conMenu_View_StatusBar) '状态栏项后
'            Set objControl = .Add(xtpControlButton, conMenu_View_Append, "附加信息(&A)", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "自动隐藏过滤条件栏(&H)", objControl.Index + 1)
'        End With
'
'        '工具菜单:主窗体可能没有,放在帮助菜单前面
'        '-----------------------------------------------------
'        Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
'        If objMenu Is Nothing Then
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
'            Set objMenu = cbrthis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
'            objMenu.ID = conMenu_ToolPopup
'        End If
'        With objMenu.CommandBar.Controls
'            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "电子签名(&S)", -1, False): objPopup.BeginGroup = True
'            With objPopup.CommandBar.Controls
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "医嘱签名(&I)")
'                objControl.IconId = conMenu_Tool_Sign
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
'                Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)"): objControl.BeginGroup = True
'            End With
'
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "住院医嘱选项(&O)"): objControl.BeginGroup = True
'            objControl.IconId = conMenu_File_Parameter
'
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "成套方案定义(&O)"): objControl.BeginGroup = True
'        End With
'
'        '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
'        '-----------------------------------------------------
'        Set objBar = cbrthis(2)
'        For Each objControl In objBar.Controls '先求出前面的最后一个Control
'            If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
'                Set objControl = objBar.Controls(objControl.Index - 1): Exit For
'            End If
'        Next
'        With objBar.Controls
'            'Set objControl = .Find(, conMenu_File_Preview) '从预览按钮之后开始加入
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新嘱", objControl.Index + 1): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除", objControl.Index + 1)
'            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送", objControl.Index + 1): objControl.BeginGroup = True
'            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "签名", objControl.Index + 1): objControl.BeginGroup = True
'            objControl.IconId = conMenu_Tool_Sign
'        End With
'
'        '命令的快键绑定
'        '-----------------------------------------------------
'        With cbrthis.KeyBindings
'            .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新开医嘱
'            .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改医嘱
'            .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除医嘱
'            .Add FCONTROL, vbKeyS, conMenu_Edit_Stop '停止医嘱
'            .Add FCONTROL, vbKeyG, conMenu_Edit_Send '医嘱发送
'            .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread '医嘱回退
'
'            .Add FCONTROL, vbKeyR, conMenu_Edit_Compend * 10# + 1 '查阅报告
'            .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '观片处理
'
'            .Add FCONTROL, vbKeyH, conMenu_View_Hide '自动隐藏过滤条件栏
'
'            .Add 0, vbKeyF11, conMenu_Tool_Option '医嘱选项
'        End With
'
'        '设置不常用命令
'        '-----------------------------------------------------
'        With cbrthis.Options
'        End With
'    End If
'    cbrthis.ActiveMenuBar.FindControl(, conMenu_LIS_RightMenu).Visible = False
'    Me.cbrthis.RecalcLayout
    '删除现在的工具栏及顶级菜单项
'    For lngCount = cbrthis.ActiveMenuBar.Controls.Count To 1 Step -1
'        cbrthis.ActiveMenuBar.Controls(lngCount).Delete
'    Next
'    For lngCount = cbrthis.Count To 2 Step -1
'        cbrthis(lngCount).Delete
'    Next
'    Call CreateCbs
End Sub

Private Sub WinsockC_DataArrival(ByVal bytesTotal As Long)
    '********************返回给技师站的信息*****************************
    'Private Const strSend_Refresh = "Refresh"      '已保存数据可以刷新
    'Private Const strSend_True = "True"            '已操作成功
    'Private Const strSend_False = "False"          '操作失败
    '*******************************************************************
    Dim strData As String
    Dim astrData() As String
    
    On Error Resume Next
    
    With Me.WinsockC
        .GetData strData
        astrData = Split(strData, ";")

        Select Case astrData(1)
            Case "Refresh"
                If Me.Tag <> "Refresh" And blnAutoRefresh And mintEditState = 0 Then
                    Me.Tag = "Refresh"
                    Call InsertOneRecored(Val(astrData(2)), False, False)
                    Me.Tag = ""
                End If
            Case "True"
                mblnSendComplete = True
            Case "False"
                mblnSendComplete = False
            Case Else
                If strData Like "AutoQCCompute|*" Then
                    If Split(strData, "|")(1) <> "" Then frmQCShowInfo.ShowMe "自动计算", Split(strData, "|")(1), Me
                End If
        End Select
    End With
End Sub

Private Sub ShowRequest(blnShow As Boolean)
    '功能       是否显示登记窗体
    '备注       参数选择时才会生效
    Dim Pane1 As Pane
    Dim blnExec As Boolean
    blnExec = Val(zlDatabase.GetPara("只在核收登记时显示登记窗口", 100, 1208, 0))
    If blnExec = False And blnShow = False Then Exit Sub    '没有选择参数时不处理
    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Request)
    If blnShow = True Then
        Pane1.Select
    Else
        Pane1.Close
    End If
    Me.dkpMain.RecalcLayout
End Sub
Private Sub GetVerifying()
    '功能           得到待检验筛选字串
    Dim intLoop As Integer
    Dim astrFilter() As String
    
    astrFilter = Split(con_主界面筛选_检验中, ";")
    For intLoop = 0 To UBound(astrFilter)
        mblnVerifying(intLoop) = zlDatabase.GetPara("检验中_" & astrFilter(intLoop), 100, 1208, True)
        If intLoop <= Me.chkSoure.UBound Then
            Me.chkSoure(intLoop).Value = IIf(mblnVerifying(intLoop), 1, 0)
        End If
    Next
End Sub
Private Sub GetWaitVerify()
    '功能           得到等待检验筛选字串
    Dim intLoop As Integer
    Dim astrFilter() As String
    astrFilter = Split(con_主界面筛选_待核收, ";")
    For intLoop = 0 To UBound(astrFilter)
        mblnWaitVerify(intLoop) = zlDatabase.GetPara("待核收_" & astrFilter(intLoop), 100, 1208, True)
        If intLoop < 2 Then
            Me.chkSoure(intLoop).Value = IIf(mblnWaitVerify(intLoop), 1, 0)
        Else
            Me.chkSoure(5).Value = IIf(mblnWaitVerify(intLoop), 1, 0)
        End If
    Next
End Sub
Private Sub CreateChildCbs()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom

    '子窗体菜单建立
    Me.cbrChild.VisualTheme = xtpThemeOffice2003
    Set Me.cbrChild.Icons = zlCommFun.GetPubIcons
    With Me.cbrChild.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
'        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .LargeIcons = False
    End With
    Me.cbrChild.EnableCustomization False

    Me.cbrChild.ActiveMenuBar.Title = "菜单"
    Me.cbrChild.ActiveMenuBar.Position = xtpBarTop
    Me.cbrChild.ActiveMenuBar.EnableDocking xtpFlagStretched Or xtpFlagHideWrap
    With Me.cbrChild.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "前一条")
        cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "后一条")
        cbrControl.Style = xtpButtonIconAndCaption
        Set cbrControl = .Add(xtpControlLabel, 0, "    定位")
        Set cbrCustom = .Add(xtpControlCustom, conMenu_File_RoomSet, "")
        cbrCustom.Handle = Me.txtGoto.hWnd
        Me.txtGoto.ToolTipText = "数字为标本号和条码、“－”打头为病人ID、“＋”住院号、“*”门诊号、“.”挂号单号、“/”收费单据号"
        
        Set cbrControl = .Add(xtpControlLabel, conMenu_Edit_UnArchive, "    收费项目")
        Set cbrCustom = .Add(xtpControlCustom, conMenu_Manage_Transfer_Send, "")
        cbrCustom.Handle = Me.cboExesItem.hWnd
        
        Set cbrControl = .Add(xtpControlLabel, conMenu_View_FindType, "")
        Set cbrPopControl = .Add(xtpControlButtonPopup, 0, "选项     ")
        cbrPopControl.Flags = xtpFlagRightAlign: cbrPopControl.Style = xtpButtonIconAndCaption
        Set cbrControl = cbrPopControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_RequestView, "使用条码扫描", -1, False)
        cbrControl.Checked = zlDatabase.GetPara("使用条码扫描", 100, 1208, False)
        
        Set cbrControl = cbrPopControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_RequestPrint, "连续输入", -1, False)
        cbrControl.Checked = zlDatabase.GetPara("连续输入", 100, 1208, False)
        
        Set cbrControl = cbrPopControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_RequestBatPrint, "保存后直接审核", -1, False)
        cbrControl.Checked = zlDatabase.GetPara("保存后直接审核", 100, 1208, True, 0)
        
        Set cbrControl = cbrPopControl.CommandBar.Controls.Add(xtpControlButton, XTP_ID_WINDOW_LIST, "显示备注", -1, False)
        cbrControl.Checked = zlDatabase.GetPara("显示检验备注", 100, 1208, False)
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, ""): cbrControl.Visible = False
        
        
        
    End With
    cbrChild.RecalcLayout
'    Call mclsExpenses.zlDefCommandBars(Me, Me.cbrthis)
'    Call mclsInAdvices.zlDefCommandBars(Me, Me.cbrthis, 2)
'    Call mclsOutAdvices.zlDefCommandBars(Me, Me.cbrthis, 2)
'    Call zldatabase.ShowReportMenu(Me.cbrthis, glngSys, glngModul, mstrPrivs)
End Sub

Private Sub SetControlFocus()
    On Error Resume Next
    If Me.Visible = False Or Me.TabList.Enabled = False Then Exit Sub
    If Me.TabList.Selected.Index = 0 Then
        Me.rptList.SetFocus
    Else
        Me.rptList1.SetFocus
    End If
End Sub

Private Sub PrintBarcord()
    Dim intBarCode As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str采集方式 As String, str执行科室 As String, str开嘱时间 As String
    Dim str管码 As String, str采血量 As String, str管名称 As String
    Dim lng病人来源 As Long
    '成生条码到PIC
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    
    With Me.rptList.FocusedRow
        If Trim(.Record(mCol.样本条码).Value) = "" Then Exit Sub
        '采集方式,执行科室,开嘱时间,管码,采血量,试管名称,
        
        strSQL = "Select A.病人来源,D.名称 As 采集方式, F.名称 As 执行科室, To_Char(C.开嘱时间, 'yyyy-MM-dd HH24:mi:ss') As 开嘱时间, E.编码 As 管码, E.采血量," & vbNewLine & _
                "       E.名称 As 管名称, A.病人来源 " & vbNewLine & _
                "From 部门表 F, 采血管类型 E, 诊疗项目目录 D, 病人医嘱记录 C, 病人医嘱发送 B, 检验标本记录 A" & vbNewLine & _
                "Where C.执行科室id = F.ID And D.试管编码 = E.编码(+) And C.诊疗项目id = D.ID And C.ID = B.医嘱id And A.样本条码 = B.样本条码 And A.ID = [1]"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.Record(mCol.ID).Value))
        Do Until rsTmp.EOF
            If Trim("" & rsTmp!管码) = "" Then
                str采集方式 = Trim("" & rsTmp!采集方式)
            Else
                str执行科室 = Trim("" & rsTmp!执行科室)
                str开嘱时间 = Trim("" & rsTmp!开嘱时间)
                str管码 = Trim("" & rsTmp!管码)
                str采血量 = Trim("" & rsTmp!采血量)
                str管名称 = Trim("" & rsTmp!管名称)
                lng病人来源 = Val(Trim("" & rsTmp!病人来源))
            End If
            rsTmp.MoveNext
        Loop
        intBarCode = zlDatabase.GetPara("使用条码", "100", "1211", False, 2)
        If intBarCode = 1 Then
            Bar39 Me.picBarCodePrint, 3, CStr(Trim(.Record(mCol.样本条码).Value)), False, True
        Else
            Bar128 Me.picBarCodePrint, 3, CStr(Trim(.Record(mCol.样本条码).Value)), True
        End If
        SavePicture Me.picBarCodePrint.Image, App.path & "\BarCode.Bmp"
        '开始打印
        
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1211_1", Me, "样本条码=" & Trim(.Record(mCol.样本条码).Value), "项目=" & Trim(.Record(mCol.检验项目).Value), _
        "病人姓名 = " & IIf(Trim(.Record(mCol.姓名).Value) <> "", Trim(.Record(mCol.姓名).Value) & IIf(Val(Trim(.Record(mCol.婴儿).Value)) = 0, "", "(婴儿" & Trim(.Record(mCol.姓名).Value) & ")"), "无"), _
        "性别 = " & IIf(Trim(.Record(mCol.性别).Value) <> "", Trim(.Record(mCol.性别).Value), "无"), _
        "年龄 = " & IIf(Trim(.Record(mCol.年龄).Value) & Trim(.Record(mCol.年龄单位).Value) <> "", Trim(.Record(mCol.年龄).Value) & Trim(.Record(mCol.年龄单位).Value), "无"), _
        "床号 = " & IIf(Trim(.Record(mCol.床号).Value) <> "", Trim(.Record(mCol.床号).Value), "无"), _
        "标识号 = " & IIf(Trim(.Record(mCol.标识号).Value) <> "", Trim(.Record(mCol.标识号).Value), "无"), _
        "所在科室 = " & IIf(Trim(.Record(mCol.病人科室).Value) <> "", Trim(.Record(mCol.病人科室).Value), "无"), _
        "采集方式 = " & IIf(str采集方式 <> "", str采集方式, "无"), _
        "标本 = " & IIf(Trim(.Record(mCol.检验标本).Value) <> "", Trim(.Record(mCol.检验标本).Value), "无"), _
        "执行科室 = " & IIf(str执行科室 <> "", str执行科室, "无"), _
        "开嘱医生 = " & IIf(Trim(.Record(mCol.申请人).Value) <> "", Trim(.Record(mCol.申请人).Value), "无"), _
        "开嘱时间 = " & IIf(str开嘱时间 <> "", str开嘱时间, "无"), _
        "采样人 = " & IIf(Trim(.Record(mCol.采样人).Value) <> "", Trim(.Record(mCol.采样人).Value), "无"), _
        "采样时间 = " & IIf(Trim(.Record(mCol.采样时间).Value) <> "", Format(Trim(.Record(mCol.采样时间).Value), "yyyy-MM-dd HH:mm:ss"), "无"), _
        "管码 = " & IIf(str管码 <> "", str管码, "无"), _
        "采血量 = " & IIf(str采血量 <> "", str采血量, "无"), _
        "试管名称 = " & IIf(str管名称 <> "", str管名称, "无"), _
        "紧急 = " & IIf(Trim(.Record(mCol.紧急).Value) <> "", Trim(.Record(mCol.紧急).Value), "无"), _
        "病人来源 = " & IIf(lng病人来源 <> 0, lng病人来源, "无"), _
        "条码图像1=" & App.path & "\BarCode.Bmp", 2)
        '删除条码图像
        Kill App.path & "\BarCode.Bmp"
    End With
End Sub

''''''''''''''''''''
''' 实现插件的HOST功能
''''''''''''''''''''
Private Property Get clsLisQueryHost_OwnerFormHandle() As Long
    clsLisQueryHost_OwnerFormHandle = Me.hWnd
End Property

Private Function clsLisQueryHost_GetRecordSet(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    '执行查询
    On Error GoTo errH
    Dim lngCount As Long
    Dim var(30) As Variant

    lngCount = UBound(arrInput)
    If lngCount > 30 Then
        MsgBox "不支持超过30个参数的SQL！", vbInformation, Me.Caption
        Exit Function
    End If
    For lngCount = LBound(arrInput) To UBound(arrInput)
        var(lngCount) = arrInput(lngCount)
    Next
    Set clsLisQueryHost_GetRecordSet = zlDatabase.OpenSQLRecord(strSQL, strTitle, var(0), var(1), var(2), var(3), var(4), var(5), var(6), var(7), var(8), var(9), var(10), var(11), var(12), var(13), var(14), var(15), var(16), var(17), var(18), var(19), var(20), var(21), var(22), var(23), var(24), var(25), var(26), var(27), var(28), var(29))
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub clsLisQueryHost_RaiseFinished(objQuery As zl9LisQuery_Def.clsLisQuery)
    '执行完成
    On Error GoTo errH
    If objQuery.Result <> "" Then
        '预留
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function clsLisQueryHost_ClientTrigger(ByVal Index As Long, ByVal strAction As String, strData As String) As String
    '客户端触发的事件
    On Error GoTo errH
    If Not mobjPlugin(Index) Is Nothing Then
        Select Case strAction
        Case "Cmd_Start"
        Case "Cmd_End"
        Case "Cmd_OK"
        Case "Cmd_Cancle"
        End Select
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub ShowHideListHead(Cols As ReportColumns, strFiled As String)
    '显示隐藏列头
    Dim intLoop As Integer
    
    For intLoop = 0 To Cols.Count - 1
        Cols(intLoop).Visible = (InStr(strFiled & ";", ";" & Cols(intLoop).Caption & ";") > 0)
    Next
End Sub

Private Sub ShowLJAverage()
    frmQCLJAverage.ShowMe Me, mstrPrivs, mlngDeptID, mlngMachineID
End Sub

