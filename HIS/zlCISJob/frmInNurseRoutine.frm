VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInNurseRoutine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "病人事务处理"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15270
   Icon            =   "frmInNurseRoutine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleMode       =   0  'User
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   960
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   31
      Top             =   3000
      Width           =   855
   End
   Begin VB.PictureBox picPrompt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1410
      ScaleHeight     =   285
      ScaleWidth      =   11865
      TabIndex        =   29
      Top             =   7530
      Width           =   11865
      Begin VB.Label lblPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   30
         TabIndex        =   30
         Top             =   60
         Width           =   10500
      End
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5955
      Left            =   2115
      ScaleHeight     =   5925
      ScaleWidth      =   5145
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   585
      Visible         =   0   'False
      Width           =   5175
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   5475
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   5160
         _Version        =   589884
         _ExtentX        =   9102
         _ExtentY        =   9657
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   3990
         Picture         =   "frmInNurseRoutine.frx":18F2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "确认"
         Top             =   5550
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   4530
         Picture         =   "frmInNurseRoutine.frx":1E7C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "取消"
         Top             =   5550
         Width           =   450
      End
   End
   Begin VB.PictureBox picCondition 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1320
      ScaleHeight     =   345
      ScaleWidth      =   9990
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Width           =   9990
      Begin VB.CommandButton cmdWarrant 
         Caption         =   "担保"
         Height          =   270
         Left            =   9105
         TabIndex        =   9
         ToolTipText     =   "担保信息查阅"
         Top             =   45
         Width           =   500
      End
      Begin VB.PictureBox pic住院次数 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7665
         ScaleHeight     =   225
         ScaleWidth      =   1335
         TabIndex        =   7
         Top             =   60
         Width           =   1365
         Begin VB.ComboBox cboPages 
            BackColor       =   &H00EAFFFF&
            Height          =   300
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   -45
            Width           =   1425
         End
      End
      Begin VB.PictureBox pic病人 
         Appearance      =   0  'Flat
         BackColor       =   &H00EAFFFF&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   810
         ScaleHeight     =   315
         ScaleWidth      =   1725
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   1755
         Begin VB.TextBox txt病人 
            Appearance      =   0  'Flat
            BackColor       =   &H00EAFFFF&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   0
            MaxLength       =   100
            TabIndex        =   3
            Top             =   70
            Width           =   1335
         End
         Begin VB.Image img病人列表 
            Height          =   360
            Left            =   1350
            Picture         =   "frmInNurseRoutine.frx":2406
            Tag             =   "弹出本病区所有病人列表"
            Top             =   -30
            Width           =   360
         End
      End
      Begin VB.PictureBox pic标识 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00EAFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   3510
         ScaleHeight     =   345
         ScaleWidth      =   3990
         TabIndex        =   4
         Top             =   0
         Width           =   3990
         Begin VB.Label lbl姓名 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "王二麻子王二麻子"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1950
            TabIndex        =   6
            Top             =   60
            Width           =   2040
         End
         Begin VB.Label lbl床号 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "床:内一科-173"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            TabIndex        =   5
            Top             =   60
            Width           =   1755
         End
      End
      Begin VB.Image img下一个 
         Height          =   360
         Left            =   2940
         Picture         =   "frmInNurseRoutine.frx":2B08
         Tag             =   "下一个病人"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image img上一个 
         Height          =   360
         Left            =   2580
         Picture         =   "frmInNurseRoutine.frx":320A
         Tag             =   "上一个病人"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image img详细信息 
         Height          =   360
         Left            =   9630
         Picture         =   "frmInNurseRoutine.frx":390C
         Tag             =   "查看病人详细信息"
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lbl定位 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "定位病人"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   1
         Top             =   90
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   7455
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInNurseRoutine.frx":400E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人颜色"
            TextSave        =   "病人颜色"
            Key             =   "病人颜色"
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
      Height          =   5265
      Left            =   255
      TabIndex        =   27
      Top             =   2085
      Width           =   7770
      _Version        =   589884
      _ExtentX        =   13705
      _ExtentY        =   9287
      _StockProps     =   64
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1215
      ScaleWidth      =   20010
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   750
      Width           =   20010
      Begin VB.Frame fraInfo 
         BackColor       =   &H00EAFFFF&
         Caption         =   "病人详细信息"
         Height          =   975
         Left            =   150
         TabIndex        =   15
         Top             =   90
         Width           =   16965
         Begin VB.ComboBox cbo过敏 
            Height          =   300
            Left            =   9120
            Style           =   2  'Dropdown List
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   570
            Width           =   4845
         End
         Begin VB.Label lbl住院号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "住院号:12345678"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   16
            Top             =   300
            Width           =   1350
         End
         Begin VB.Label lbl病况 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "一般"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2220
            TabIndex        =   17
            Top             =   300
            Width           =   390
         End
         Begin VB.Label lbl性别 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "男"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   180
            Left            =   2910
            TabIndex        =   18
            Top             =   300
            Width           =   195
         End
         Begin VB.Label lbl年龄 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "32岁"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3330
            TabIndex        =   19
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lbl病人类型 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "重庆市医保[YBZH0001]"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   12600
            TabIndex        =   23
            Top             =   300
            Width           =   1800
         End
         Begin VB.Label lbl护理等级 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "一级护理"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   4590
            TabIndex        =   20
            Top             =   300
            Width           =   720
         End
         Begin VB.Label lbl入院时间 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "yyyy-MM-dd HH:mm～yyyy-MM-dd HH:mm"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   6030
            TabIndex        =   21
            Top             =   300
            Width           =   3060
         End
         Begin VB.Label lbl医疗付款方式 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "社会基本医疗保险"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   9945
            TabIndex        =   22
            Top             =   300
            Width           =   1440
         End
         Begin VB.Label lbl诊断 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "诊断:慢性支气管炎、慢性支气管炎、慢性支气管炎、慢性支气"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   24
            Top             =   630
            Width           =   6060
         End
         Begin VB.Label lbl药物过敏 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "药物过敏:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   8100
            TabIndex        =   25
            Top             =   630
            Width           =   810
         End
      End
   End
   Begin MSComctlLib.ImageList imgRPT 
      Left            =   -165
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":48A0
            Key             =   "Pati"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":4E3A
            Key             =   "Notify"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":53D4
            Key             =   "等待审查"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":596E
            Key             =   "拒绝审查"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":5F08
            Key             =   "正在审查"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":64A2
            Key             =   "正在抽查"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":6EB4
            Key             =   "审查反馈"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":78C6
            Key             =   "抽查反馈"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":7E60
            Key             =   "审查整改"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":8872
            Key             =   "抽查整改"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":9284
            Key             =   "审查归档"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":FAE6
            Key             =   "未导入"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":10080
            Key             =   "执行中"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":1061A
            Key             =   "不符合"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":1102C
            Key             =   "正常结束"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":115C6
            Key             =   "变异结束"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":11B60
            Key             =   "Child"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":120FA
            Key             =   "单病种"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":1895C
            Key             =   "Out"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":18EF6
            Key             =   "紧急"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":19490
            Key             =   "男人"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInNurseRoutine.frx":1FCF2
            Key             =   "女人"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   180
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmInNurseRoutine.frx":26554
      Left            =   705
      Top             =   210
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmInNurseRoutine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum PATI_TYPE
    pt入院待入住 = 0
    pt转科待入住 = 1
    pt转病区待入住 = 2
    pt在院 = 3
'    pt家庭病床 = 3.1
'    pt预转科 = 3.2
'    pt转病区 = 3.3
    pt预出 = 4
    pt出院 = 5
    pt死亡 = 6
    pt最近转出 = 7
End Enum
Private Enum PATI_COLUMN
    c_图标 = 0
    C_状态 = 1
    c_床号 = 2
    C_病人ID = 3
    C_主页ID = 4
    c_姓名 = 5
    c_住院号 = 6
    c_入院日期 = 7
    c_出院日期 = 8
    c_病人类型 = 9
End Enum

Private Enum PATI_COLWIDTH
    cw_图标 = 18
    cw_状态 = 0
    cw_床号 = 40
    Cw_病人ID = 0
    cw_主页ID = 0
    cw_姓名 = 60
    cw_住院号 = 60
    cw_入院日期 = 70
    cw_出院日期 = 70
    cw_病人类型 = 100
End Enum

Private mblnShow As Boolean
Private mblnAdd As Boolean
Private mobjBar As CommandBar

'子窗体对象定义
Private mclsEMR As Object  '新版病历zlRichEMR.clsDockEMR
Private WithEvents mclsAdvices As zlPublicAdvice.clsDockInAdvices
Attribute mclsAdvices.VB_VarHelpID = -1
Private WithEvents mclsEPRs As zlRichEPR.cDockInEPRs
Attribute mclsEPRs.VB_VarHelpID = -1
Private WithEvents mclsTends As zl9TendFile.clsTendFile
Attribute mclsTends.VB_VarHelpID = -1
Private mclsTendEPRs As zlRichEPR.cDockInTendEPRs
Attribute mclsTendEPRs.VB_VarHelpID = -1
Private WithEvents mclsFeeQuery As zl9InExse.clsFeeQuery
Attribute mclsFeeQuery.VB_VarHelpID = -1
'Private WithEvents mfrmResponse As frmAuditResponse '审查反馈窗口
Private WithEvents mclsPath As zlPublicPath.clsDockPath
Attribute mclsPath.VB_VarHelpID = -1
Private mclsWardMonitor As clsWardMonitor     '监护仪接口

Private mcolSubForm As Collection
Private mcolSubFormOperation As Collection
Private mfrmActive As Form
Private mobjMipModule As Object
'其它窗体变量
Private mobjParent As Object
Private mstrPrivs As String
Private mstrPage As String
Private mPatiInfo As PatiInfo '历史住院记录中的,不一定为当前的
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng病区ID As Long
Private mstrScope As String
Private mintChange As Integer
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mintPrePage As Integer
Private mblnUnRefresh As Boolean
Private mblnRefreshBar As Boolean
Private mlngRowIndex As Long

Private mblnMonitor As Boolean '监护仪程序是否存在
Private mstrMonitor As String '监护仪程序路径

Private mbytSize As Byte

Private mrsPati As New ADODB.Recordset
Private mblnTabTmp As Boolean
Private mlng婴儿科室ID As Long
Private mlng婴儿病区ID As Long

'整体护理相关变量
Private mstrNurseParentID As String '整体护理中的病人ID
Private mstrRelatedUnitID As String '整体护理病区ID
Private mstrRelatedUserID As String '整体护理人员ID
Private marrTabAttribute '存储每一个tab的属性值(0-普通页面;1-整体护理页面)
Private mColNurseFormUrl As Collection  '整体护理窗体url信息
Private mobjNurseForm As Object '整体护理窗体（多个页面用一个窗体，每次切换进行卸载和创建,主要是为了释放内存）

Public Sub zlInitMip(ByVal objMipModule As Object)
    '消息对象
    Set mobjMipModule = objMipModule
End Sub

Public Sub NurseRoutine(ByVal frmParent As Object, ByVal strPrivs As String, ByVal lng病区ID As Long, ByVal lng病人ID As Long, _
    ByVal dtOutBegin As Date, ByVal dtOutEnd As Date, ByVal intChange As Integer, ByVal strScope As String, tPati As PatiInfo, _
    Optional ByVal strPage As String = "医嘱", Optional ByVal rsThis As ADODB.Recordset, Optional ByVal bytSize As Byte = 0)
    
    If lng病人ID = 0 Then Exit Sub
    
    Set mobjParent = frmParent
    mstrPrivs = strPrivs
    mstrPage = strPage
    mlng病区ID = lng病区ID
    mlng病人ID = lng病人ID
    mdtOutBegin = dtOutBegin
    mdtOutEnd = dtOutEnd
    mintChange = intChange
    mstrScope = strScope
    mPatiInfo = tPati
    mintPrePage = -1            '每次切换病人时清除
    mblnAdd = Not mblnShow
    mbytSize = bytSize
    
    Call RefreshPatiList(rsThis)
    
    If mblnShow Then
        mintPrePage = -1
        Call AddPages
        Exit Sub
    End If
    Call ReSetFontSize
    mblnShow = True
    mobjParent.mblnRoutine = mblnShow
    Me.Show , frmParent
End Sub

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    Dim lngCol As Long
    Dim PATI_COLWIDTH As Variant
    bytFontSize = IIf(mbytSize = 0, 9, IIf(mbytSize = 1, 12, mbytSize))
    
    Me.FontSize = bytFontSize
    Me.FontName = "宋体"
    
    'CommandBars
    Set CtlFont = cbsMain.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsMain.Options.Font = CtlFont
    'DockingPane
    Set CtlFont = DkpMain.PaintManager.CaptionFont
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set DkpMain.PaintManager.CaptionFont = CtlFont
    'TabControl
    Set CtlFont = tbcSub.PaintManager.Font
    If CtlFont Is Nothing Then
        Set CtlFont = Me.Font
    End If
    CtlFont.Size = bytFontSize
    Set tbcSub.PaintManager.Font = CtlFont
            
    lbl定位.FontSize = Me.FontSize
    lbl定位.Top = pic病人.Top + (pic病人.Height - lbl定位.Height) \ 2
    lbl定位.Left = 60
    txt病人.FontSize = Me.FontSize
    txt病人.Top = (pic病人.Height - txt病人.Height)
    pic病人.Left = lbl定位.Left + lbl定位.Width + 20
    img上一个.Left = pic病人.Left + pic病人.Width
    img下一个.Left = img上一个.Left + img上一个.Width
    pic标识.Left = img下一个.Left + img下一个.Width + TextWidth("刘")
    Me.pic标识.Width = lbl姓名.Width + lbl姓名.Left
    pic住院次数.Left = pic标识.Left + pic标识.Width + 50
    cboPages.FontSize = Me.FontSize
    cboPages.Left = -30
    cboPages.Top = -30
    pic住院次数.Height = cboPages.Height - 20
    pic住院次数.Top = (picCondition.Height - pic住院次数.Height) \ 2
    If pic住院次数.Top < 0 Then pic住院次数.Top = 0
    Me.pic住院次数.Width = Me.cboPages.Width - 50
    
    cmdWarrant.FontSize = Me.FontSize
    cmdWarrant.Width = TextWidth(cmdWarrant.Caption & "刘")
    cmdWarrant.Height = TextWidth("刘") + TextWidth("刘") \ 2
    cmdWarrant.Left = pic住院次数.Left + pic住院次数.Width + 50
    cmdWarrant.Top = (picCondition.Height - cmdWarrant.Height) \ 2
    img详细信息.Left = cmdWarrant.Left + cmdWarrant.Width + 100
    picCondition.Width = img详细信息.Left + img详细信息.Width + 60
    
    '病人选择
    Set CtlFont = rptPati.PaintManager.CaptionFont
    CtlFont.Size = bytFontSize
    Set rptPati.PaintManager.CaptionFont = CtlFont
    
    Set CtlFont = rptPati.PaintManager.TextFont
    CtlFont.Size = bytFontSize
    Set rptPati.PaintManager.TextFont = CtlFont
    
    PATI_COLWIDTH = Array(cw_图标, cw_状态, cw_床号, Cw_病人ID, cw_主页ID, cw_姓名, cw_住院号, cw_入院日期, cw_出院日期, cw_病人类型)
    For lngCol = C_状态 To rptPati.Columns.Count - 1
        rptPati.Columns.Column(lngCol).Width = PATI_COLWIDTH(lngCol) + (PATI_COLWIDTH(lngCol) * IIf(mbytSize = 0, 0, 1)) / 3
    Next lngCol
    
    rptPati.Redraw
    
    '病人信息栏
    fraInfo.FontSize = bytFontSize
    lbl住院号.FontSize = bytFontSize
    lbl住院号.Height = TextHeight("刘")
    lbl病况.FontSize = bytFontSize
    lbl病况.Height = TextHeight("刘")
    lbl性别.FontSize = bytFontSize
    lbl性别.Height = TextHeight("刘")
    lbl年龄.FontSize = bytFontSize
    lbl年龄.Height = TextHeight("刘")
    lbl护理等级.FontSize = bytFontSize
    lbl护理等级.Height = TextHeight("刘")
    lbl入院时间.FontSize = bytFontSize
    lbl入院时间.Height = TextHeight("刘")
    lbl医疗付款方式.FontSize = bytFontSize
    lbl医疗付款方式.Height = TextHeight("刘")
    lbl病人类型.FontSize = bytFontSize
    lbl病人类型.Height = TextHeight("刘")
    lbl诊断.FontSize = bytFontSize
    lbl诊断.Height = TextHeight("刘")
    lbl药物过敏.FontSize = bytFontSize
    lbl药物过敏.Height = TextHeight("刘")
    cbo过敏.FontSize = bytFontSize
    cbo过敏.Left = lbl药物过敏.Left + lbl药物过敏.Width + TextHeight("刘")
    
    lblPrompt.FontSize = bytFontSize
    Call Form_Resize
End Sub

'55430:刘鹏飞,2013-02-27,双击作废医嘱定位到病人事物的医嘱页面
Public Sub OrientTabPage(ByVal strTab As String, Optional ByVal strID As String = "")
'-------------------------------------------------------------
'功能:定位到病人事物中指定的页面,以及对应页面指定的文件或医嘱等
'-------------------------------------------------------------
    Dim intIdx As Integer
    Dim blnSeek As Boolean
    
    blnSeek = False
    If strTab = tbcSub.Tag Then blnSeek = True
    If blnSeek = False Then
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            blnSeek = True
            mblnAdd = False
            tbcSub.Item(intIdx).Tag = strTab
            tbcSub.Item(intIdx).Selected = True
        End If
    End If
    '定位页面成功,则定位到具体的位置
    If blnSeek = True Then
        Select Case strTab
            Case "医嘱"
                If Val(strID) = 0 Then Exit Sub
                Call mclsAdvices.zlSeekAndViewEPRReport(Val(strID))
            Case "护理"
        End Select
    End If
End Sub

Public Sub RefreshPatiList(Optional ByVal rsThis As ADODB.Recordset)
    On Error GoTo ErrHand
    
    '刷新病人清单,仍定位到当前操作的病人上
    Call LoadPatient(rsThis)
    mrsPati.Filter = "病人ID=" & mlng病人ID
    '54408:刘鹏飞,2012-10-10,同步处理找不到病人就定位到第一个病人
    '如：在院病人进入病人事物，然后在主界面将此病人出院。如果病人出院时间不在查询出院范围内可能就会出现此情况
    If mrsPati.RecordCount = 0 Then
        mrsPati.Filter = "": mrsPati.MoveFirst
        mlng病人ID = Val(mrsPati!病人ID)
    End If
    mlng主页ID = Val(mrsPati!主页ID)
    mlng婴儿科室ID = Val(mrsPati!婴儿科室ID & "")
    mlng婴儿病区ID = Val(mrsPati!婴儿病区ID & "")
    mrsPati.Filter = ""
    mrsPati.MoveFirst
    '90592:同一病人可能存在多条记录，但状态不同，按照病人ID，主页ID查找
    mrsPati.Find ("Key='" & mlng病人ID & ":" & mlng主页ID & "'")
    rptPati.Records.DeleteAll
    picPati.Visible = False
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LocatePati(ByVal intType As Integer)
    '参数说明:intType:1-上一个病人;2-下一个病人
    '病人范围:在床病人循环,与老版保持一致
    Dim blnExit As Boolean  '强制退出
    On Error Resume Next
    
redo:
    If intType = 1 Then
        mrsPati.MovePrevious
        If mrsPati.BOF Then mrsPati.MoveLast
    Else
        mrsPati.MoveNext
        If mrsPati.EOF Then mrsPati.MoveFirst
    End If
    If mrsPati!病人ID <> 0 Then
        If mrsPati!病人ID <> mlng病人ID Then
            mlng病人ID = mrsPati!病人ID
            mlng主页ID = mrsPati!主页ID
            mintPrePage = -1
            Call AddPages
        Else
            If blnExit Then Exit Sub
            blnExit = True
            GoTo redo
        End If
    Else
        GoTo redo
    End If
    
    picPati.Visible = False
End Sub

Private Sub cmdFilterCancel_Click()
    picPati.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Call rptPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub cmdWarrant_Click()
    Call frmPatiSurety.ShowMe(Me, mlng病人ID, mlng主页ID)
End Sub

Private Sub Form_Activate()
    picPrompt.Visible = Me.stbThis.Visible
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("[']", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Function GetVersion() As String
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    strSQL = " select 版本号 from zlsystems where 编号=100"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取标准版本号")
    GetVersion = rsTemp!版本号
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub LoadPatient(Optional ByVal rsThis As ADODB.Recordset)
    'U10.32开始支持多音字
    Dim strSQL As String
    Dim strBriefCode As String
    Dim blnSupport As Boolean
    Dim strField As String, strValue As String
    Dim rsPati As New ADODB.Recordset
    On Error GoTo ErrHand
    
    blnSupport = (Val(Split(GetVersion, ".")(1)) >= 32)
    If blnSupport Then
        strBriefCode = ",zlpinyincode(NVL(B.姓名,A.姓名),0,0,',',1) AS 简码 "
    Else
        strBriefCode = ",zlspellcode(NVL(B.姓名,A.姓名)) AS 简码"
    End If
    
    '54408:刘鹏飞,2012-10-10,主界面查找出院病人可能在有效时间范围之外
    strField = "Key," & adLongVarChar & ",50|排序," & adDouble & ",2|排序2," & adDouble & ",2|类型," & adLongVarChar & ",50|病人ID," & adDouble & ",18|主页ID," & adDouble & ",18|" & _
           "住院号," & adDouble & ",18|姓名," & adLongVarChar & ",20|简码," & adLongVarChar & ",200|性别," & adLongVarChar & ",10|年龄," & adLongVarChar & ",20|科室," & adLongVarChar & ",50|" & _
           "科室ID," & adDouble & ",18|住院医师," & adLongVarChar & ",20|责任护士," & adLongVarChar & ",20|病案状态," & adLongVarChar & ",20|" & _
           "床号," & adLongVarChar & ",20|护理等级," & adLongVarChar & ",50|费别," & adLongVarChar & ",50|当前病况," & adLongVarChar & ",50|" & _
           "入院日期," & adLongVarChar & ",20|出院日期," & adLongVarChar & ",20|住院天数," & adLongVarChar & ",20|出院方式," & adLongVarChar & ",20|" & _
           "病人类型," & adLongVarChar & ",50|状态," & adLongVarChar & ",10|险类," & adDouble & ",18|就诊卡号," & adLongVarChar & ",20|路径状态," & adLongVarChar & ",20|" & _
           "颜色," & adDouble & ",18|单病种," & adLongVarChar & ",10|婴儿科室ID," & adDouble & ",18|婴儿病区ID," & adDouble & ",18"
    Call Record_Init(mrsPati, strField)
    
'    If rsThis Is Nothing Then
        '入院等入科和转科待入科病人(病人科室所属的病区都可接收)
        'c.科室id + 0,说明：通过H表的索引连接过滤后，记录数量很少，再连接B表则更快
        If Val(Mid(mstrScope, 5, 1)) <> 0 Then
            '84938:刘鹏飞，性能优化(添加条件:A.主页ID=B.主页ID)
            strSQL = _
                "Select /*+ RULE */Distinct" & vbNewLine & _
                " Decode(B.状态,1,0,Decode(c.开始原因,3,1,2)) As 排序, Decode(Nvl(b.病案状态, 0), 0, 999, b.病案状态) As 排序2," & _
                " Decode(B.状态,1,'入院待入住病人',Decode(c.开始原因,3,'转科待入住病人','转病区待入住病人')) As 类型," & _
                " a.病人id, b.主页id, A.门诊号,B.住院号, NVL(b.姓名,a.姓名) 姓名" & strBriefCode & ", NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄," & vbNewLine & _
                " d.名称 As 科室, c.科室id, c.经治医师 As 住院医师,b.责任护士, b.病案状态, lpad(c.床号,10,' ') AS 床号," & _
                " e.名称 As 护理等级, b.费别,b.当前病况, Decode(b.入科时间,NULL,b.入院日期,b.入科时间) AS 入院日期 , b.出院日期,B.出院方式, b.病人类型, b.状态, b.险类, a.就诊卡号," & vbNewLine & _
                " -1 As 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间))+1 as 住院天数,Z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID" & vbNewLine & _
                "From 病人信息 A, 病案主页 B, 病人变动记录 C, 部门表 D, 收费项目目录 E, 病人类型 Z" & vbNewLine & _
                "Where B.病人类型=Z.名称(+) And a.在院 = 1 And a.病人id = b.病人id And A.主页ID=B.主页ID And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id " & vbNewLine & _
                "      And (C.病区ID=[1] or C.病区ID is null) And c.科室id = d.Id" & vbNewLine & _
                "      And (d.站点='" & gstrNodeNo & "' Or d.站点 is Null)" & vbNewLine & _
                "      And b.护理等级id = e.Id(+) And Nvl(c.附加床位, 0) = 0 And c.终止时间 Is Null" & vbNewLine & _
                "      And (c.开始原因 in(1,3) And Exists(Select 1 From 病区科室对应 H Where c.科室id = h.科室id And h.病区id = [1]) or c.开始原因=15 And c.病区id = [1])" & vbNewLine & _
                "      And ((c.开始原因 = 1 And b.状态 = 1) Or (c.开始原因 in (3,15) And c.开始时间 Is Null And b.状态 = 2)) "
        End If
        '在院病人(固定强制显示)
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.状态,3,4,DECODE(B.出院病床, NULL, 3.1,DECODE(B.状态,2,3.2,3))) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.状态,3,'预出院病人',DECODE(B.出院病床, NULL, '家庭病床',DECODE(B.状态,2,'预转科病人', '在院病人'))) as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名" & strBriefCode & ",NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
            " lpad(B.出院病床,10,' ') AS 床号,E.名称 as 护理等级,B.费别,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) AS 入院日期 ,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间))+1 as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z,在院病人 R" & _
            " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And A.主页ID=B.主页ID  And Nvl(B.状态,0)<>1" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And (R.病区ID=[1] Or b.婴儿病区ID=[1]) And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And a.病人ID=R.病人ID And A.当前病区ID=R.病区ID And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
        '出院病人:出院病人可能已有多次住院
        If Val(Mid(mstrScope, 2, 1)) <> 0 Then
            strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
                "Select /*+ RULE */ Decode(B.出院方式,'死亡',6,5) as 排序," & _
                " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
                " Decode(B.出院方式,'死亡','死亡病人','出院病人') as 类型," & _
                " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名" & strBriefCode & ",NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
                " lpad(B.出院病床,10,' ') AS 床号,E.名称 as 护理等级,B.费别,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) AS 入院日期 ,B.出院日期,B.出院方式,B.病人类型," & _
                " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(b.出院日期)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间))+1 as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID" & _
                " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z" & _
                " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.状态=0" & _
                " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And B.当前病区ID+0=[1] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
                " And B.出院日期 Between [2] And [3] And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
        End If
        '转出病人:在院,医生和床号显示本科转出前的
        If Val(Mid(mstrScope, 4, 1)) <> 0 Then
            strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
                "Select /*+ RULE */ Distinct 7 as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,'转出病人' as 类型," & _
                " A.病人ID,B.主页ID,A.门诊号,B.住院号,NVL(B.姓名,A.姓名) 姓名" & strBriefCode & ",NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄,D.名称 as 科室,C.科室ID,C.经治医师 as 住院医师,B.责任护士,B.病案状态," & _
                " lpad(c.床号,10,' ') AS 床号,E.名称 as 护理等级,B.费别,B.当前病况,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) AS 入院日期 ,B.出院日期,B.出院方式,B.病人类型," & _
                " B.状态,B.险类,A.就诊卡号,Nvl(b.路径状态,-1) 路径状态,trunc(sysdate)-trunc(Decode(b.入科时间,NULL,b.入院日期,b.入科时间))+1 as 住院天数,z.颜色,B.单病种,B.婴儿科室ID,B.婴儿病区ID" & _
                " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D,收费项目目录 E,病人类型 Z" & _
                " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.护理等级ID=E.ID(+)" & _
                " And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
                " And B.当前病区ID<>[1] And C.病区ID+0=[1] And C.科室ID=D.ID" & _
                " And Nvl(C.附加床位,0)=0 And C.终止原因 In(3,15) And C.终止时间 Between Sysdate-[7] And Sysdate" & _
                " And Nvl(B.状态,0)<>2 And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
        End If
        strSQL = strSQL & " Order by 排序,床号,主页ID Desc"
        
        Screen.MousePointer = 11
        On Error GoTo ErrHand
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "提取病人列表", mlng病区ID, _
            CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")), _
            Val(Mid(mstrScope, 1, 1)), Val(Mid(mstrScope, 2, 1)), Val(Mid(mstrScope, 5, 1)), mintChange)
        
        '开始装载病人信息
        rsPati.Filter = 0
        Call CopyReocrd(rsPati)
        
        '通过在主界面直接查找的出院病人可能不在出院范围内，此处需要重新加载
        If rsThis Is Nothing Then Exit Sub
        If rsThis.State = adStateClosed Then Exit Sub
        rsThis.Filter = "排序=5 or 排序=6 or 排序=7"
        Call CopyReocrd(rsThis)
        '从新进行病人排序
        mrsPati.Sort = "排序,床号,主页ID Desc"
'    Else
'        Set mrsPati = rsThis.Clone
'    End If
    Screen.MousePointer = 0
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CopyReocrd(ByVal rsPati As ADODB.Recordset)
    '54408:刘鹏飞,2012-10-10
    Dim strField As String, strValue As String
    
    If rsPati.RecordCount <> 0 Then rsPati.MoveFirst
    strField = "Key|排序|排序2|类型|病人ID|主页ID|住院号|姓名|简码|性别|年龄|科室|科室ID|住院医师|责任护士|病案状态|床号|护理等级|费别|当前病况|入院日期|出院日期|住院天数|出院方式|病人类型|状态|险类|就诊卡号|路径状态|颜色|单病种|婴儿科室ID|婴儿病区ID"
    Do While Not rsPati.EOF
        mrsPati.Filter = "Key='" & Val(rsPati!病人ID) & ":" & Val(rsPati!主页ID) & "'"
        If mrsPati.RecordCount = 0 Then
            strValue = Val(rsPati!病人ID) & ":" & Val(rsPati!主页ID) & "|" & rsPati!排序 & "|" & rsPati!排序2 & "|" & rsPati!类型 & "|" & rsPati!病人ID & "|" & rsPati!主页ID & "|" & NVL(rsPati!住院号, 0) & "|" & rsPati!姓名 & "|" & rsPati!简码 & "|" & rsPati!性别 & "|" & _
                      rsPati!年龄 & "|" & NVL(rsPati!科室) & "|" & NVL(rsPati!科室ID, 0) & "|" & NVL(rsPati!住院医师) & "|" & NVL(rsPati!责任护士) & "|" & NVL(rsPati!病案状态, 0) & "|" & NVL(rsPati!床号) & "|" & _
                      NVL(rsPati!护理等级, "三级") & "|" & rsPati!费别 & "|" & NVL(rsPati!当前病况, "一般") & "|" & Format(rsPati!入院日期, "yyyy-MM-dd") & "|" & Format(rsPati!出院日期, "yyyy-MM-dd") & "|" & rsPati!住院天数 & "|" & rsPati!出院方式 & "|" & _
                      NVL(rsPati!病人类型, "普通病人") & "|" & rsPati!状态 & "|" & NVL(rsPati!险类, 0) & "|" & NVL(rsPati!就诊卡号) & "|" & NVL(rsPati!路径状态, 0) & "|" & NVL(rsPati!颜色, 0) & "|" & NVL(rsPati!单病种) & "|" & NVL(rsPati!婴儿科室ID, 0) & "|" & NVL(rsPati!婴儿病区ID, 0)
            
            Call Rec.AddNew(mrsPati, strField, strValue)
        End If
        rsPati.MoveNext
    Loop
End Sub

Private Sub Form_Load()
    Dim objPane As Pane, strTab As String, intIdx As Integer
    Dim blnCol As Boolean, strTmp As String, i As Long, bln路径状态 As Boolean
    Dim intType As Integer, arrTmp As Variant
    
    '整体护理病人业务签入变量
    Dim strTabs As String, strErrMsg As String
    Dim strName As String, strUrl As String, strParam As String
    Dim j As Integer
    Dim objXML As New DOMDocument
    Dim objNodeList As IXMLDOMNodeList
    
    On Error GoTo ErrHand
    
    picPrompt.Visible = False
    
    mblnRefreshBar = False
    marrTabAttribute = Array()
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

    'TabControl
    '-----------------------------------------------------
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, True)
    If GetInsidePrivs(p新版住院病历, True) <> "" Then
        Set mclsEMR = DynamicCreate("zlRichEMR.clsDockEMR", "电子病历")
        If Not mclsEMR Is Nothing Then
            If Not mclsEMR.Init(gobjEmr, gcnOracle, glngSys) Then
                Set mclsEMR = Nothing
            End If
        End If
    End If
    Set mclsAdvices = New zlPublicAdvice.clsDockInAdvices
    Call mclsAdvices.zlInitMip(mobjMipModule)
    Set mclsEPRs = New zlRichEPR.cDockInEPRs
    Set mclsTends = New zl9TendFile.clsTendFile
    Call mclsTends.InitTendFile(gcnOracle, glngSys)
    Set mclsTendEPRs = New zlRichEPR.cDockInTendEPRs
    Set mclsFeeQuery = New zl9InExse.clsFeeQuery
    Call mclsFeeQuery.InitCallByNurse(gfrmMain, gcnOracle, gstrDBUser, glngSys)

    Set mclsPath = New zlPublicPath.clsDockPath
    Call mclsAdvices.zlInitPath(mclsPath)
    Set mclsWardMonitor = New clsWardMonitor
    
    Set mcolSubFormOperation = New Collection
    Set mcolSubForm = New Collection
    If Not mclsEMR Is Nothing Then
        mcolSubForm.Add mclsEMR.zlGetForm, "_新病历"
    End If
    mcolSubForm.Add mclsPath.zlGetForm, "_路径"
    mcolSubForm.Add mclsAdvices.zlGetForm, "_医嘱"
    mcolSubForm.Add mclsFeeQuery.zlGetForm, "_费用"
    mcolSubForm.Add mclsEPRs.zlGetForm, "_病历"
    mcolSubForm.Add mclsTends.zlGetForm, "_护理"
    mcolSubForm.Add mclsTendEPRs.zlGetForm, "_护理病历"
    If mclsWardMonitor.Enabled Then
        mcolSubForm.Add mclsWardMonitor.zlGetForm, "_监护"
    End If
    
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .OneNoteColors = True
            .Position = xtpTabPositionTop
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        If GetInsidePrivs(p临床路径应用, True) <> "" Then
            .InsertItem(intIdx, "临床路径", picTmp.hwnd, 0).Tag = "路径": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p住院医嘱下达, True) <> "" Or GetInsidePrivs(p住院医嘱发送, True) <> "" Then
            .InsertItem(intIdx, "医嘱记录", picTmp.hwnd, 0).Tag = "医嘱": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p费用查询, True) <> "" Then
            .InsertItem(intIdx, "费用记录", picTmp.hwnd, 0).Tag = "费用": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p住院病历管理, True) <> "" Then
            .InsertItem(intIdx, "住院病历", picTmp.hwnd, 0).Tag = "病历": intIdx = intIdx + 1
        End If
        If GetInsidePrivs(p护理记录管理, True) <> "" Then
            .InsertItem(intIdx, "护理记录", picTmp.hwnd, 0).Tag = "护理": intIdx = intIdx + 1
            .InsertItem(intIdx, "护理病历", picTmp.hwnd, 0).Tag = "护理病历": intIdx = intIdx + 1
        End If
        If mclsWardMonitor.Enabled Then
            If InStr(GetInsidePrivs(p住院护士站), "护理监护") > 0 Then
                .InsertItem(intIdx, "护理监护", picTmp.hwnd, 0).Tag = "监护": intIdx = intIdx + 1
            End If
        End If
        If GetInsidePrivs(p新版住院病历, True) <> "" And Not mclsEMR Is Nothing Then
            .InsertItem(intIdx, "电子病历", picTmp.hwnd, 0).Tag = "新病历": intIdx = intIdx + 1
        End If
        
        For i = 0 To tbcSub.ItemCount - 1
            ReDim Preserve marrTabAttribute(UBound(marrTabAttribute) + 1)
            marrTabAttribute(UBound(marrTabAttribute)) = 0
        Next
        '整体护理病人业务嵌入
        If gbln启用整体护理接口 = True Then
            If InitNurseIntegrate = True Then
                If gobjNurseIntegrate.GetPatientMethod(strTabs, strErrMsg) = False Then
                    MsgBox "获取整体护理病人业务标签失败！" & vbCrLf & "详细信息：" & strErrMsg, vbInformation, gstrSysName
                Else
                    If objXML.loadXML(strTabs) = False Then Exit Sub
                    Set mColNurseFormUrl = New Collection
                    Set objNodeList = objXML.selectNodes(".//Tab//Item")
                    For i = 0 To objNodeList.length - 1
                        strName = objNodeList.Item(i).childNodes(0).Text
                        strUrl = objNodeList.Item(i).childNodes(1).Text
                        '读取节点属性值
                        strParam = ""
                        For j = 0 To objNodeList.Item(i).childNodes(1).Attributes.length - 1
                             strParam = strParam & "&" & objNodeList.Item(i).childNodes(1).Attributes(j).nodeName & "=" & objNodeList.Item(i).childNodes(1).Attributes(j).nodeValue
                        Next j
                        If Left(strParam, 1) = "&" Then strParam = Mid(strParam, 2)
                        strUrl = strUrl & IIf(strParam = "", "", "?" & strParam)
                        .InsertItem(intIdx, strName, picTmp.hwnd, 0).Tag = strName: intIdx = intIdx + 1
                        mColNurseFormUrl.Add strUrl, "_" & strName
                        ReDim Preserve marrTabAttribute(UBound(marrTabAttribute) + 1)
                        marrTabAttribute(UBound(marrTabAttribute)) = 1
                    Next i
                End If
            End If
        End If
        
        Call CreatePlugInOK(p住院护士站)
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            strTmp = gobjPlugIn.GetFormCaption(glngSys, p住院护士站)
            Call zlPlugInErrH(err, "GetFormCaption")
            If strTmp <> "" Then
                arrTmp = Split(strTmp, ",")
                For i = 0 To UBound(arrTmp)
                    strTmp = arrTmp(i)
                    
                    mcolSubForm.Add gobjPlugIn.GetForm(glngSys, p住院护士站, strTmp), "_" & strTmp
                    .InsertItem(intIdx, strTmp, mcolSubForm("_" & strTmp).hwnd, 0).Tag = strTmp: intIdx = intIdx + 1
                    ReDim Preserve marrTabAttribute(UBound(marrTabAttribute) + 1)
                    marrTabAttribute(UBound(marrTabAttribute)) = 0
                    Call zlPlugInErrH(err, "GetForm")
                Next
            End If
            err.Clear: On Error GoTo 0
        End If
        
        If .ItemCount = 0 Then
            MsgBox "你没有使用病人事务处理的权限。", vbInformation, gstrSysName
            Unload Me: Exit Sub
        End If
        
        '恢复上次选择的卡片
        strTab = zlDatabase.GetPara("医护功能", glngSys, p住院护士站)
        strTab = mstrPage
        For intIdx = 0 To tbcSub.ItemCount - 1
            If tbcSub(intIdx).Visible And tbcSub(intIdx).Tag = strTab Then Exit For
        Next
        If intIdx <= tbcSub.ItemCount - 1 Then
            strTab = .Item(intIdx).Tag
            .Item(intIdx).Tag = "" '避免激活事件
            .Item(intIdx).Selected = True
            .Item(intIdx).Tag = strTab
        Else
            .Item(0).Selected = True '新建时就自动选中了这个,不会再激活事件
        End If
        '只加载选择的子窗体
        Call tbcSub_SelectedChanged(.Selected)
    End With
    
    '初始化病人选择器
    Dim objCol As ReportColumn
    With rptPati
        .Columns.DeleteAll
        Set objCol = .Columns.Add(c_图标, "", cw_图标, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(C_状态, "状态", cw_状态, True)
        Set objCol = .Columns.Add(c_床号, "床号", cw_床号, True)
        Set objCol = .Columns.Add(C_病人ID, "病人ID", Cw_病人ID, False): objCol.Visible = False
        Set objCol = .Columns.Add(C_主页ID, "主页ID", cw_主页ID, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", cw_姓名, True)
        Set objCol = .Columns.Add(c_住院号, "住院号", cw_住院号, True)
        Set objCol = .Columns.Add(c_入院日期, "入院日期", cw_入院日期, True)
        Set objCol = .Columns.Add(c_出院日期, "出院日期", cw_出院日期, True)
        Set objCol = .Columns.Add(c_病人类型, "病人类型", cw_病人类型, True)
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = (objCol.Index = C_状态)
            objCol.Sortable = True
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有病人..."
        End With
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList Me.imgRPT
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(C_状态)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(c_床号)
    End With

    '读取界面数据
    '-----------------------------------------------------
    mstrMonitor = ""
    mblnMonitor = Dir(App.Path & "\..\gdhs\AC2005.exe") <> ""
    If mblnMonitor Then mstrMonitor = App.Path & "\..\gdhs\AC2005.exe"
    
    '界面恢复:放在最后执行
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    
    '恢复上次病人信息工具栏的状态
    If (zlDatabase.GetPara("病人信息工具栏", glngSys, p住院护士站, 1) = 0) Then
        mobjBar.Visible = False
'        picInfo.Visible = False
    End If
    Me.WindowState = vbMaximized
    
    Call AddPages
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long, strTmp As String
    Dim lng病人ID As Long, lng主页ID As Long
    
    On Error GoTo ErrHand

    If Control.ID <> 0 Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
    Case conMenu_Manage_Monitor '监护仪
        Call ExecuteMonitor
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.picPrompt.Visible = Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Jump '跳转
        If Me.tbcSub.Selected.Index + 1 <= Me.tbcSub.ItemCount - 1 Then
            Me.tbcSub.Item(Me.tbcSub.Selected.Index + 1).Selected = True
        Else
            Me.tbcSub.Item(0).Selected = True
        End If
    Case conMenu_View_Refresh '刷新
        '68116:刘鹏飞,2014-01-06,刷新添加病人列表刷新
        'Call tbcSub_SelectedChanged(tbcSub.Item(tbcSub.Selected.Index))
        lng病人ID = mlng病人ID: lng主页ID = mlng主页ID
        Call RefreshPatiList(mrsPati)
        If lng病人ID = mlng病人ID Then mlng主页ID = lng主页ID
        mintPrePage = -1
        Call AddPages
        Call ReSetFontSize
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        Call gobjKernel.ShowDiagHelp(vbModeless, Me)
    Case conMenu_Tool_Reference_2 '诊疗措施参考
        Call gobjKernel.ShowClincHelp(vbModeless, Me)
    
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '退出
        Unload Me
    Case Else
        mblnUnRefresh = True
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            With mPatiInfo
                strTmp = Split(Control.Parameter, ",")(1)
                If strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1132" Then '住院科室日报
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                             "病区=" & mlng病区ID, "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID)
                ElseIf strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Or strTmp = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then    '病人帐页和催款表
                    Call mclsFeeQuery.zlExecuteCommandBars(Control)
                Else
                    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me, _
                        "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "住院号=" & .住院号, "病人病区=" & .病区ID, _
                        "病人科室=" & .科室ID, "床号=" & .床号)
                End If
            End With
        Else
            Select Case Me.tbcSub.Selected.Tag
            Case "路径"
                If mlng婴儿病区ID <> 0 Then
                    If mlng婴儿病区ID = mlng病区ID Then
                        MsgBox "该病人已经转出本科室了，只有婴儿留在本科室，不允许操作路径。", vbInformation, Me.Caption
                        Exit Sub
                    End If
                End If
                Call mclsPath.zlExecuteCommandBars(Control)
            Case "医嘱"
                Call mclsAdvices.zlExecuteCommandBars(Control)
            Case "费用"
                Call mclsFeeQuery.zlExecuteCommandBars(Control)
            Case "病历"
                Call mclsEPRs.zlExecuteCommandBars(Control)
            Case "护理"
                Call mclsTends.zlExecuteCommandBars(Control)
            Case "护理病历"
                Call mclsTendEPRs.zlExecuteCommandBars(Control)
            Case "新病历"
                Call mclsEMR.zlExecuteCommandBars(Control)
            Case Else
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.ExeButtomClick(glngSys, p住院护士站, mcolSubForm("_" & tbcSub.Selected.Tag), tbcSub.Selected.Tag, Control.Caption, mlng病人ID, mlng主页ID, "")
                    Call zlPlugInErrH(err, "ExeButtomClick")
                    err.Clear: On Error GoTo 0
                End If
            End Select
        End If
        mblnUnRefresh = False
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim rsPatiLog As ADODB.Recordset
    Dim i As Long, j As Long, strPrivs As String
    Dim objControl As CommandBarControl
    On Error GoTo ErrHand

    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case tbcSub.Selected.Tag
    Case "路径"
         Call mclsPath.zlPopupCommandBars(CommandBar)
    Case "医嘱"
        Call mclsAdvices.zlPopupCommandBars(CommandBar)
    Case "费用"
        Call mclsFeeQuery.zlPopupCommandBars(CommandBar)
    Case "病历"

    Case "护理"

    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo ErrHand
    
    Select Case Control.ID
    Case conMenu_Manage_Monitor '监护仪
        Control.Visible = mblnMonitor
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_Tool_Reference_1 '疾病诊断参考
        Control.Visible = GetInsidePrivs(p疾病诊断参考) <> ""
    Case conMenu_Tool_Reference_2 '药品及诊疗参考
        Control.Visible = GetInsidePrivs(p药品诊疗参考) <> ""
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
'    Case conMenu_Tool_MedRecAuditResponse '审查反馈
'        '都可以调用，至少可以查看(当前或历史)
'        Control.Enabled = rptPati.Rows.Count > 0
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_1" Then Control.Visible = tbcSub.Selected.Tag = "费用"  '催款表
            If Split(Control.Parameter, ",")(1) = "ZL" & glngSys \ 100 & "_INSIDE_1139_2" Then Control.Visible = tbcSub.Selected.Tag = "费用"  '病人帐页
        End If
        If Not mblnRefreshBar Then Exit Sub
        Select Case tbcSub.Selected.Tag
        Case "路径"
            Call mclsPath.zlUpdateCommandBars(Control)
        Case "医嘱"
            Call mclsAdvices.zlUpdateCommandBars(Control)
        Case "费用"
            Call mclsFeeQuery.zlUpdateCommandBars(Control)
        Case "病历"
            Call mclsEPRs.zlUpdateCommandBars(Control)
        Case "护理"
            Call mclsTends.zlUpdateCommandBars(Control)
        Case "护理病历"
            Call mclsTendEPRs.zlUpdateCommandBars(Control)
        Case "新病历"
            Call mclsEMR.zlUpdateCommandBars(Control)
        End Select
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub SubWinDefCommandBar(ByVal objItem As TabControlItem)
'功能：刷新子窗体菜单及工具条
    Dim objControl As CommandBarControl
    Dim bytStyle As XTPButtonStyle
    Dim blnShowBar As Boolean
    Dim lngCount As Long, idx As Long
    Dim strName As String
    Dim blnNurseIntegrate As Boolean
    
    On Error GoTo ErrHand
    If gbln启用整体护理接口 = True Then
        blnNurseIntegrate = Val(marrTabAttribute(objItem.Index)) = 1
    End If
    '记录现有菜单样式
    mblnRefreshBar = False
    blnShowBar = True
    bytStyle = xtpButtonIconAndCaption
    If cbsMain.Count >= 2 Then
        idx = GetFirstCommandBar(cbsMain(2).Controls)
        If idx > 0 Then
            blnShowBar = cbsMain(2).Visible
            bytStyle = cbsMain(2).Controls(idx).Style
        End If
    End If

    '刷新子窗口菜单
    Call LockWindowUpdate(Me.hwnd)

    Me.Caption = "病人事务处理 - " & objItem.Caption & "(当前用户：" & UserInfo.姓名 & ")"

    '删除现在的所有菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    '删除工具栏
'    For lngCount = cbsMain.Count To 2 Step -1
'        cbsMain(lngCount).Delete
'    Next

    '主窗口重新加入
    Call MainDefCommandBar
    
    '子窗口重新加入
    Select Case objItem.Tag
    Case "路径"
        Call mclsPath.zlDefCommandBars(Me, Me.cbsMain, 1, True)
    Case "医嘱"
        Call mclsAdvices.zlDefCommandBars(Me, Me.cbsMain, 1, True)
    Case "费用"
        Call mclsFeeQuery.zlDefCommandBars(Me, Me.cbsMain, 1, True)
    Case "病历"
        Call mclsEPRs.zlDefCommandBars(Me.cbsMain, True)
    Case "护理"
        Call mclsTends.zlDefCommandBars(Me.cbsMain, True)
    Case "护理病历"
        Call mclsTendEPRs.zlDefCommandBars(Me.cbsMain, True)
    Case "新病历"
        Call mclsEMR.zlDefCommandBars(Me.cbsMain, True)
    Case Else
        If blnNurseIntegrate = False Then
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                strName = gobjPlugIn.GetButtomName(glngSys, p住院护士站, mcolSubForm("_" & objItem.Tag), objItem.Tag)
                Call zlPlugInErrH(err, "GetButtomName")
                '构建菜单
                If strName <> "" Then Call PlugInInSideBar(cbsMain, strName, 1)
                err.Clear: On Error GoTo 0
            End If
        End If
    End Select
    mblnRefreshBar = True
    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '恢复工具栏按钮属性
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
'        cbsMain(lngCount).EnableDocking xtpFlagHideWrap
'        For Each objControl In cbsMain(lngCount).Controls
'            If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
'                objControl.Style = xtpButtonIcon
'            Else
'                objControl.Style = bytStyle
'            End If
'        Next
'        cbsMain(lngCount).Visible = blnShowBar
    Next

    '如果用了RecalcLayout反而不正常
    Call LockWindowUpdate(0)

    If blnNurseIntegrate = False Then
        Set mfrmActive = mcolSubForm("_" & objItem.Tag)
    Else
        Set mfrmActive = mobjNurseForm
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SubWinRefreshData(ByVal objItem As TabControlItem)
 '功能：刷新子窗体数据及状态
    Dim blnEdit As Boolean, strInPatiNO As String, lng路径状态 As Long
    Dim lngType As PATI_TYPE, lng病区ID As Long, lng科室ID As Long
    Dim lngState As TYPE_PATI_State
    Dim blnNurseIntegrate As Boolean
    
    On Error GoTo ErrHand
    If gbln启用整体护理接口 = True Then
        blnNurseIntegrate = Val(marrTabAttribute(objItem.Index)) = 1
    End If
    
    Call SetOrGetSubFromOperation(objItem.Tag, True)
    If mlng病人ID = 0 Then
        '要求子窗体按无数据处理界面
        Select Case objItem.Tag
        Case "路径"
            Call mclsPath.zlRefresh(0, 0, 0, 0, 0, False)
        Case "医嘱"
            Call mclsAdvices.zlRefresh(0, 0, 0, 0, CDate(0), 0)
        Case "费用"
            Call mclsFeeQuery.zlRefresh(0, 0, 0, 0, 0, False, False, False)
        Case "病历"
            Call mclsEPRs.zlRefresh(0, 0, 0, False, False)
        Case "护理"
            Call mclsTends.zlRefresh(0, 0, 0, False, False)
        Case "护理病历"
            Call mclsTendEPRs.zlRefresh(0, 0, 0, False, False, False)
        Case "监护"
            Call mclsWardMonitor.HideWindow
        Case "新病历"
            Call mclsEMR.zlRefresh(0, 0, 0, 0, 3)
        Case Else
            If blnNurseIntegrate = False Then
                If Not gobjPlugIn Is Nothing Then
                    On Error Resume Next
                    Call gobjPlugIn.RefreshForm(glngSys, p住院护士站, mcolSubForm("_" & objItem.Tag), objItem.Tag, 0, "", 0, False)
                    Call zlPlugInErrH(err, "RefreshForm")
                    err.Clear: On Error GoTo 0
                End If
            Else
                If InitNurseIntegrate = True Then
                    Call gobjNurseIntegrate.RefreshPatientMethod(mobjNurseForm, mobjNurseForm.Tag, mstrNurseParentID, mstrRelatedUnitID, mstrRelatedUserID)
                End If
            End If
        End Select
    Else
        With mPatiInfo
            lngType = .排序
            
            '67485:刘鹏飞,2013-11-13,查看待转出病人应该是转出之前的科室ID
            If pt最近转出 = lngType And mrsPati.RecordCount > 0 Then
                lng科室ID = NVL(mrsPati!科室ID, 0) '最近转出病人为原科室ID
            Else
                lng科室ID = Val("" & .科室ID)
            End If
            
            If InStr("," & pt入院待入住 & "," & pt最近转出 & "," & pt转科待入住 & "," & pt转病区待入住 & ",", "," & lngType & ",") > 0 Then
                '待入住病人，转出病人，传当前界面的病区
                lng病区ID = mlng病区ID
            Else
                lng病区ID = .病区ID
            End If
            If lngType = pt最近转出 Then
                lngState = ps最近转出
            ElseIf lngType = pt转科待入住 Or lngType = pt转病区待入住 Then
                lngState = ps待转入
            Else
                lngState = IIf(.出院日期 = CDate(0), IIf(.状态 = 3, ps预出, ps在院), ps出院)
            End If
            
            Select Case objItem.Tag
            Case "路径"
                Call mclsPath.zlRefresh(mlng病人ID, .主页ID, lng病区ID, lng科室ID, .状态, .数据转出, True, , mlng病区ID)
            Case "医嘱"
                lng路径状态 = .路径状态
                '50906:刘鹏飞,2012-09-18,入院待入住病人，根据参数"允许给待入住病人下达医嘱"决定是否可以下达医嘱
                If lngType = pt入院待入住 And Val(zlDatabase.GetPara("允许给待入住病人下达医嘱", glngSys, p住院医嘱下达, 1)) = 0 Then
                    lngState = ps待转入 'lngState=ps待转入时新开医嘱等功能不可用
                End If
                Call mclsAdvices.zlRefresh(mlng病人ID, .主页ID, lng病区ID, lng科室ID, lngState, .数据转出, , , , lng路径状态, mlng病区ID)
            Case "费用"
                Call mclsFeeQuery.zlRefresh(mlng病人ID, mlng主页ID, Val(.住院号), lng病区ID, .险类, .数据转出, .出院日期 <> CDate("0:00:00"), .结清, False, _
                    lngType = pt最近转出 Or lngType = pt预出 Or lngType = pt出院, lng科室ID)
            Case "病历"
                Call mclsEPRs.zlRefresh(mlng病人ID, .主页ID, mlng病区ID, False, .数据转出, 0, True, lng病区ID, lngState)
            Case "护理"
                blnEdit = True
                If lngType = pt出院 Or lngType = pt死亡 Then
                    If Not (Val(.病案状态) = 0 Or Val(.病案状态) = 2 Or Val(.病案状态) = 999) Then
                        '可能是在院抽查反馈状态，出院后并未提交审查
                        If Val(.病案状态) = 1 Or Val(.病案状态) = 2 Then blnEdit = False
                    End If
                ElseIf lngType = pt转科待入住 Or lngType = pt转病区待入住 Then
                    blnEdit = False
                End If
                blnEdit = blnEdit And (mlng病区ID = .病区ID Or lngType = pt最近转出)
                Call mclsTends.zlRefresh(mlng病人ID, .主页ID, mlng病区ID, blnEdit, False, lng病区ID, lngState)
            Case "护理病历"
                Call mclsTendEPRs.zlRefresh(mlng病人ID, .主页ID, mlng病区ID, True, True, .数据转出)
            Case "监护"
                strInPatiNO = Trim(.住院号)
                If strInPatiNO = "" Then
                    Call mclsWardMonitor.HideWindow
                Else
                    Call mclsWardMonitor.ShowInfor(strInPatiNO)
                End If
            Case "新病历"
                Call mclsEMR.zlRefresh(mlng病人ID, .主页ID, mlng病区ID, lngState, 3)
            Case Else
                If blnNurseIntegrate = False Then
                    If Not gobjPlugIn Is Nothing Then
                        On Error Resume Next
                        Call gobjPlugIn.RefreshForm(glngSys, p住院护士站, mcolSubForm("_" & objItem.Tag), objItem.Tag, mlng病人ID, "", .主页ID, .数据转出, , , _
                                        lng病区ID, lng科室ID, , lngState, , lng路径状态)
                        Call zlPlugInErrH(err, "RefreshForm")
                        err.Clear: On Error GoTo 0
                    End If
                Else
                    If InitNurseIntegrate = True Then
                        Call gobjNurseIntegrate.RefreshPatientMethod(mobjNurseForm, mobjNurseForm.Tag, mstrNurseParentID, mstrRelatedUnitID, mstrRelatedUserID)
                    End If
                End If
            End Select
        End With
    End If
    
    '字体设置
    Select Case objItem.Tag
        Case "路径"
            Call mclsPath.SetFontSize(mbytSize)
        Case "医嘱"
            Call mclsAdvices.SetFontSize(mbytSize)
        Case "费用"
            Call mclsFeeQuery.SetFontSize(mbytSize)
        Case "病历"
            Call mclsEPRs.SetFontSize(mbytSize)
        Case "护理"
            Call mclsTends.SetFontSize(mbytSize)
        Case "护理病历"
            Call mclsTendEPRs.SetFontSize(mbytSize)
        Case "监护"
            'Call mclsWardMonitor.SetFontSize(mbytSize)
        Case "新病历"
            On Error Resume Next
            Call mclsEMR.SetFontSize(mbytSize)
            err.Clear: On Error GoTo 0
        End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
'2.其他命令根据主窗体业务的不同，可能不同
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = conMenu_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)") '固有
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…") '固有
        
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): objControl.BeginGroup = True '固有
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)") '固有
'        With objPopup.CommandBar.Controls
'            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False '固有
'            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False '固有
'            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False '固有
'        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)") '固有

        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Jump, "窗格跳转(&J)")
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    objMenu.ID = conMenu_ToolPopup
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Reference, "资料参考(&R)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Tool_Reference_1, "疾病诊断参考(&D)", -1, False
            .Add xtpControlButton, conMenu_Tool_Reference_2, "诊疗措施参考(&C)", -1, False
        End With
'        Set objControl = .Add(xtpControlButton, conMenu_Tool_MedRecAuditResponse, "审查反馈(&S)")
'            objControl.BeginGroup = True
'            objControl.ToolTipText = "处理或查看病案审查反馈"
    End With

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
    cbsMain(1).EnableDocking xtpFlagHideWrap
    
    If mblnAdd Then
        '工具栏定义
        '-----------------------------------------------------
        Set objBar = cbsMain.Add("条件工具栏", xtpBarTop) '固有
        objBar.EnableDocking xtpFlagStretched
        With objBar.Controls
            Set objCustom = .Add(xtpControlCustom, 1, "")
            objCustom.Handle = picCondition.hwnd
        End With
        Set mobjBar = cbsMain.Add("病人信息工具栏", xtpBarTop) '固有
        mobjBar.EnableDocking xtpFlagStretched
        mobjBar.Closeable = True
        With mobjBar.Controls
            Set objCustom = .Add(xtpControlCustom, 1, "")
            objCustom.Handle = picInfo.hwnd
        End With
        mblnAdd = False
    End If

    '读取发布到该模块的报表(不含虚拟模块的,如:住院科室日报、催款单、催款表都不显示,后面手工加到文件菜单下)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, 1265, mstrPrivs, "ZL1_INSIDE_1261_1", "ZL1_INSIDE_1261_5", "ZL1_INSIDE_1261_4", "ZL1_INSIDE_1261_6", "ZL1_INSIDE_1132", "ZL1_INSIDE_1139_1", "ZL1_INSIDE_1139_3", "ZL1_INSIDE_1261_7", "ZL1_INSIDE_1261_8")

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF6, conMenu_View_Jump '跳转
    End With
    
    Call cbsMain.RecalcLayout
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    
    picInfo.Width = Me.ScaleWidth - 200
    
    Call cbsMain.RecalcLayout
End Sub

Private Sub img病人列表_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngColor As Long, j As Long
    Dim lngloop As Long
    Dim objRow As ReportRow, blnSelect As Boolean
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim lngLeft As Long, lngTop  As Long, lngRight As Long, lngBottom As Long
    If Button <> 1 Then Exit Sub
    On Error GoTo ErrHand
    
    If rptPati.Records.Count = 0 Then
        '显示病人列表供选择
        With mrsPati
            .MoveFirst
            
            Do While Not .EOF
                Set objRecord = Me.rptPati.Records.Add()
                objRecord.Tag = CStr(!病人ID & "," & !主页ID)
                
                Set objItem = objRecord.AddItem("")
                
                '61824:刘鹏飞,2013-05-23,显示单病种标志
                If NVL(!单病种) <> "" Then
                    objItem.Icon = imgRPT.ListImages("单病种").Index - 1
                Else
                    objItem.Icon = Val(IIf(!性别 = "女", imgRPT.ListImages("女人").Index, imgRPT.ListImages("男人").Index)) - 1
                End If
                Set objItem = objRecord.AddItem(CStr(!排序 & !类型))
                objItem.Caption = CStr(!排序 & !类型)
                
                Set objItem = objRecord.AddItem(zlStr.Lpad(NVL(!床号), 10))
                objItem.Caption = Trim(NVL(!床号, " "))
                objRecord.AddItem Val(!病人ID)
                objRecord.AddItem Val(!主页ID)
                objRecord.AddItem CStr(NVL(!姓名))
                Set objItem = objRecord.AddItem(CStr(NVL(!住院号)))
                objItem.Caption = NVL(!住院号, " ")
                
                Set objItem = objRecord.AddItem(Format(!入院日期, "yyyy-MM-dd"))
                objItem.Caption = Format(!入院日期, "yyyy-MM-dd")
                Set objItem = objRecord.AddItem(Format(!出院日期, "yyyy-MM-dd"))
                objItem.Caption = Format(!出院日期, "yyyy-MM-dd")
                
                Set objItem = objRecord.AddItem(NVL(!病人类型))
                objItem.Caption = NVL(!病人类型)
                
                '提取病人类型的颜色
                lngColor = NVL(!颜色, 0)
                If lngColor <> 0 Then
                    For j = 1 To rptPati.Columns.Count - 1
                        objRecord.Item(j).ForeColor = lngColor
                    Next
                End If
                .MoveNext
            Loop
            
            .MoveFirst
            .Find ("Key='" & mlng病人ID & ":" & mlng主页ID & "'")
            If .EOF Then .MoveFirst: .Find "病人ID=" & mlng病人ID
        End With
    End If
    '调整坐标
    Call mobjBar.GetWindowRect(lngLeft, lngTop, lngRight, lngBottom)
    rptPati.Populate '缺省不选中任何行
    picPati.Left = picCondition.Left + Me.pic病人.Left
    picPati.Top = lngTop - Me.Top - 480
    picPati.Visible = True
    mlngRowIndex = -1
    '选中当前病人(先折叠组的话,Rows.Count只有组的个数了,所以先定位,再折叠)
    blnSelect = False
    For lngloop = 0 To rptPati.Rows.Count - 1
        If Not (rptPati.Rows(lngloop).Record Is Nothing) Then
            If Val(rptPati.Rows(lngloop).Record.Item(C_病人ID).Value) = mlng病人ID Then
                Set objRow = rptPati.Rows(lngloop)
            End If
            If Val(rptPati.Rows(lngloop).Record.Item(C_病人ID).Value) = mlng病人ID And Val(rptPati.Rows(lngloop).Record.Item(C_主页ID).Value) = mlng主页ID Then
                Set rptPati.FocusedRow = rptPati.Rows(lngloop)
                blnSelect = True
                Exit For
            End If
        End If
    Next
  
    If blnSelect = False And Not objRow Is Nothing Then
        Set rptPati.FocusedRow = objRow
    End If
    
    '折叠所有组(选中病人那一组不折叠)
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Index <> rptPati.FocusedRow.ParentRow.Index Then
            objRow.Expanded = False
        End If
    Next
    rptPati.FocusedRow.EnsureVisible
    If rptPati.Visible Then rptPati.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub img病人列表_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo Me.pic病人.hwnd, img病人列表.Tag
End Sub

Private Sub img上一个_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LocatePati(1)
End Sub

Private Sub img上一个_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hwnd, img上一个.Tag
End Sub

Private Sub img下一个_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call LocatePati(2)
End Sub

Private Sub img下一个_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hwnd, img下一个.Tag
End Sub

Private Sub img详细信息_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
'    picInfo.Visible = picInfo.Visible Xor True
    mobjBar.Visible = mobjBar.Visible Xor True
End Sub

Private Sub img详细信息_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hwnd, img详细信息.Tag
End Sub


Private Sub lbl姓名_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo pic标识.hwnd, lbl姓名.Caption
End Sub

Private Sub mclsAdvices_ExecLogModi(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, ByVal 执行时间 As String, 完成 As Boolean)
    On Error Resume Next
    mblnUnRefresh = True
    完成 = frmTechnicLog.ShowMe(Me, p住院医嘱发送, 科室ID, 医嘱ID, 发送号, False, 执行时间)
    mblnUnRefresh = False
    On Error GoTo 0
End Sub

Private Sub mclsAdvices_ExecLogNew(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, 完成 As Boolean)
    On Error Resume Next
    mblnUnRefresh = True
    完成 = frmTechnicLog.ShowMe(Me, p住院医嘱发送, 科室ID, 医嘱ID, 发送号, False)
    mblnUnRefresh = False
    On Error GoTo 0
End Sub

Private Sub mclsAdvices_RequestRefresh(ByVal RefreshNotify As Boolean)
'功能：医嘱子窗体要求刷新
    If RefreshNotify = True Then
        '仅刷新医嘱提醒区域(自动刷新时)
        frmNotify.mblnFirst = True
    Else
        '55982:刘鹏飞,2012-11-20,修改发送出院医嘱，死亡医嘱不刷新问题
        '重新刷新医嘱信息
        Call tbcSub_SelectedChanged(tbcSub.Item(tbcSub.Selected.Index))
    End If
End Sub

Private Sub mclspath_RequestRefresh(ByVal lngPathState As Long)
'功能：临床路径中刷新病人信息列表中的状态,-1表示未导入状态
    'todo:需要处理
'    With rptPati.SelectedRows(0)
'        .Record(col_路径状态).Value = lngPathState
'        .Record(col_路径状态).Caption = " "
'        .Record(col_路径状态).Icon = -1 + Choose(lngPathState + 2, imgPati.ListImages("未导入").Index, imgPati.ListImages("不符合").Index, _
'                imgPati.ListImages("执行中").Index, imgPati.ListImages("正常结束").Index, imgPati.ListImages("变异结束").Index)
'    End With
'
'    If rptPati.Columns(col_路径状态).Visible = False Then
'        rptPati.Columns(col_路径状态).Visible = True
'    End If
'    rptPati.Populate
End Sub

Private Sub mclsAdvices_StatusTextUpdate(ByVal Text As String)
'功能：医嘱子窗体要求更新状态栏
    'todo:需要处理
    If Text = "" Then
        If mlng病人ID > 0 And mlng主页ID > 0 Then
            lblPrompt.Caption = IIf(stbThis.Panels(2).Tag = "", "", stbThis.Panels(2).Tag & "，") & _
                mrsPati!姓名 & "：" & GetPati费用信息(mlng病人ID, mlng主页ID)
        Else
            lblPrompt.Caption = stbThis.Panels(2).Tag
        End If
    Else
        lblPrompt.Caption = Text
    End If
    lblPrompt.ForeColor = &H80000008
End Sub

Private Sub cboPages_Click()
'功能：选择某次住院记录时，读取相关的病人信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lng主页ID As Long
    
    If cboPages.ListIndex = -1 Then Exit Sub
    If cboPages.ListIndex = mintPrePage Then Exit Sub
    mintPrePage = cboPages.ListIndex
    mlng主页ID = cboPages.ItemData(cboPages.ListIndex)

    On Error GoTo errH
    '90592:如果列表中相同病人有多条，则选择住院次数后默认定位
    lng主页ID = Val(mrsPati!主页ID)
    If Not Val(mrsPati!主页ID) = mlng主页ID Then
        mrsPati.MoveFirst: mrsPati.Find "Key='" & mlng病人ID & ":" & mlng主页ID & "'"
        If mrsPati.EOF = True Then mrsPati.MoveFirst: mrsPati.Find "Key='" & mlng病人ID & ":" & lng主页ID & "'"
    End If
    strSQL = "Select NVL(b.姓名,a.姓名) 姓名, NVL(b.性别,a.性别) 性别, NVL(b.年龄,a.年龄) 年龄, b.住院号, b.出院病床, b.医疗付款方式, d.信息值 As 医保号, b.险类, b.当前病况, c.名称 As 护理等级, Decode(b.入科时间,NULL,b.入院日期,b.入科时间) AS 入院日期 , b.出院日期, b.编目日期," & vbNewLine & _
            "       b.病人类型, b.状态, b.数据转出, b.出院科室id, b.当前病区id,b.病案状态,B.婴儿科室ID,B.婴儿病区ID, a.住院次数, e.房间号" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 收费项目目录 C, 病案主页从表 D, 床位状况记录 E" & vbNewLine & _
            "Where a.病人id = b.病人id And a.病人id = [1] And b.主页id = [2] And b.护理等级id = c.Id(+) And b.病人id = d.病人id(+) And" & vbNewLine & _
            "      b.主页id = d.主页id(+) And d.信息名(+) = '医保号' And b.出院科室id = e.科室id(+) And b.病人id = e.病人id(+) And b.出院病床 = e.床号(+)"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    With rsTmp
        '保险病人姓名红色显示
        lbl床号.Caption = "床:" & NVL(!出院病床)
        lbl姓名.Caption = NVL(!姓名)
        lbl姓名.ForeColor = NVL(mrsPati!颜色, 0)
        lbl性别.Caption = NVL(!性别)
        lbl年龄.Caption = NVL(!年龄)
        
        lbl住院号.Caption = "住院号:" & NVL(!住院号)
        lbl护理等级.Caption = NVL(!护理等级)
        lbl医疗付款方式.Caption = NVL(!医疗付款方式)

        '危重病人病况红色显示
        lbl病况.Caption = NVL(!当前病况)
        If NVL(!当前病况) = "危" Or NVL(!当前病况) = "重" Or NVL(!当前病况) = "急" Then
            lbl病况.ForeColor = &HC0&
        Else
            lbl病况.ForeColor = lbl住院号.ForeColor
        End If

        lbl入院时间.Caption = Format(!入院日期, "yyyy-MM-dd HH:mm")
        If Not IsNull(!出院日期) Then
            lbl入院时间.Caption = lbl入院时间.Caption & "～" & Format(!出院日期, "yyyy-MM-dd HH:mm")
        End If

        lbl病人类型.Caption = NVL(!病人类型, "普通病人")
        If NVL(!医保号) <> "" Then lbl病人类型.Caption = lbl病人类型.Caption & "[" & NVL(!医保号) & "]"

        '诊断
        lbl诊断.Caption = "诊断:" & GetPatiDiagnose(mlng病人ID, mlng主页ID, 2)

        '病人信息
        mPatiInfo.排序 = mrsPati!排序
        mPatiInfo.状态 = NVL(!状态, 0)
        mPatiInfo.住院号 = NVL(!住院号)
        mPatiInfo.床号 = NVL(!出院病床)
        mPatiInfo.主页ID = mlng主页ID
        mPatiInfo.病区ID = NVL(!当前病区ID, 0)
        mPatiInfo.科室ID = NVL(!出院科室ID, 0)
        mPatiInfo.入院日期 = !入院日期
        If Not IsNull(!出院日期) Then
            mPatiInfo.出院日期 = !出院日期
        Else
            mPatiInfo.出院日期 = CDate(0)
        End If
        mPatiInfo.数据转出 = NVL(!数据转出, 0) <> 0
        mPatiInfo.病案状态 = Val(NVL(!病案状态, 0))
        
        mlng婴儿科室ID = Val(!婴儿科室ID & "")
        mlng婴儿病区ID = Val(!婴儿病区ID & "")
    End With


    '以下信息取当前住院次数的
    strSQL = "Select B.状态,Decode(b.入科时间,NULL,b.入院日期,b.入科时间) AS 入院日期 , b.出院日期,B.住院号,b.出院病床,B.病人性质,B.数据转出,B.险类,b.当前病区id,B.出院科室ID,B.当前病区ID,Decode(Nvl(X.费用余额, 0), 0, '√', '') As 结清" & _
        " From 病案主页 B,病人余额 X" & _
        " Where B.病人ID=[1] And B.主页ID=[2] And B.病人ID = X.病人ID(+) And X.性质(+) = 1 And X.类型(+)=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    With rsTmp
        mPatiInfo.险类 = Val("" & !险类)
        mPatiInfo.结清 = Not IsNull(!结清)
        mPatiInfo.性质 = NVL(!病人性质, 0)
        mPatiInfo.产科 = Sys.DeptHaveProperty(Val(!出院科室ID & ""), "产科")
    End With
    
    '根据姓名长度调整各控件位置及大小
    Me.pic标识.Width = lbl姓名.Width + lbl姓名.Left
    Me.pic住院次数.Width = Me.cboPages.Width - 50
    Me.pic住院次数.Left = pic标识.Left + pic标识.Width + 50
    Me.cmdWarrant.Left = Me.pic住院次数.Left + Me.pic住院次数.Width + 50
    Me.img详细信息.Left = Me.cmdWarrant.Left + Me.cmdWarrant.Width + 100
    picCondition.Width = Me.img详细信息.Left + Me.img详细信息.Width + 100
    
    '获取整体护理病区和病人ID
    Call GeNurseRelatedUnitID
    '提取病人费用信息
    Call mclsAdvices_StatusTextUpdate("")
    
    '刷新子窗体数据
    Call SubWinRefreshData(tbcSub.Selected)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mclsAdvices_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
End Sub

Private Sub mclspath_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
'功能：临床路径中查看电子病历报告
    Call gobjRichEPR.ViewDocument(Me, 报告ID, CanPrint)
End Sub

Private Sub mclsTends_RefreshPrompt(ByVal strInfo As String, ByVal blnImportant As Boolean)
'    lblPrompt.Caption = strInfo
'    lblPrompt.ForeColor = IIf(blnImportant, &HFF&, &H80000008)
End Sub

Private Sub picCondition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picCondition.hwnd, ""
End Sub

Private Sub pic标识_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo pic标识.hwnd, ""
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Call cmdFilterCancel_Click
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If rptPati.Records.Count = 0 Then Exit Sub
    If rptPati.FocusedRow.Record Is Nothing Then Exit Sub
    
    mlng病人ID = Split(rptPati.FocusedRow.Record.Tag, ",")(0)
    mlng主页ID = Split(rptPati.FocusedRow.Record.Tag, ",")(1)
    '如果需要病人定位后按上一个,下一个时按定位前的顺序,可把该语句屏蔽掉
    mrsPati.MoveFirst
    mrsPati.Find "Key='" & mlng病人ID & ":" & mlng主页ID & "'"
    
    picPati.Visible = False
    txt病人.Text = ""
    mintPrePage = -1
    Call AddPages
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptPati_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub rptPati_SelectionChanged()
    '59268:刘鹏飞,2013-04-23,默认排序后会展开所有分组，对于查找不方便。处理方式为点击那组排序后就展开那一组
    Dim objRow As ReportRow
    If rptPati.FocusedRow Is Nothing Then Exit Sub
    If rptPati.FocusedRow.GroupRow = True Then
        mlngRowIndex = rptPati.FocusedRow.Index
        For Each objRow In rptPati.Rows
            If objRow.GroupRow = True Then
                If objRow.Index = rptPati.FocusedRow.Index Then
                    Exit For
                ElseIf objRow.Expanded = True Then
                    mlngRowIndex = mlngRowIndex - objRow.Childs.Count
                End If
            End If
        Next
    End If
End Sub

Private Sub rptPati_SortOrderChanged()
    '59268:刘鹏飞,2013-04-23,默认排序后会展开所有分组，对于查找不方便。处理方式为点击那组排序后就展开那一组
    Dim lngloop As Long
    Dim objRow As ReportRow
    Dim lng病人ID As Long
    If rptPati.FocusedRow Is Nothing Then
        '折叠所有组(选中病人那一组不折叠)
        For Each objRow In rptPati.Rows
            If mlngRowIndex >= 0 And mlngRowIndex <= rptPati.Rows.Count Then
                If objRow.GroupRow And objRow.Index <> mlngRowIndex Then
                    objRow.Expanded = False
                End If
            End If
        Next
    Else
        If rptPati.FocusedRow Is Nothing Then Exit Sub
        lng病人ID = rptPati.FocusedRow.Record.Item(C_病人ID).Value
        '选中当前病人(先折叠组的话,Rows.Count只有组的个数了,所以先定位,再折叠)
        For lngloop = 0 To rptPati.Rows.Count - 1
            If Not (rptPati.Rows(lngloop).Record Is Nothing) Then
                If Val(rptPati.Rows(lngloop).Record.Item(C_病人ID).Value) = lng病人ID Then
                    Set rptPati.FocusedRow = rptPati.Rows(lngloop)
                    Exit For
                End If
            End If
        Next
        
        '折叠所有组(选中病人那一组不折叠)
        For Each objRow In rptPati.Rows
            If objRow.GroupRow And objRow.Index <> rptPati.FocusedRow.ParentRow.Index Then
                objRow.Expanded = False
            End If
        Next
    End If
    If Not rptPati.FocusedRow Is Nothing Then rptPati.FocusedRow.EnsureVisible
    If rptPati.Visible Then rptPati.SetFocus
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "病人颜色" Then
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'功能：刷新子窗体界面及数据
'说明：仅在人为切换界面卡片激活
    Dim Index As Long, objItem As TabControlItem
    If mblnTabTmp Then Exit Sub
    If Item.Tag = "" Then Exit Sub '初始添卡时,还没赋值
    If Item.Tag <> tbcSub.Tag Then Call UnLoadPageForm '加载一个页面GDI就会增长，为了控制GDI增长，切换页面时卸载上一个页面窗体
    If Item.Handle = picTmp.hwnd Then
        Index = Item.Index
        mblnTabTmp = True
        Screen.MousePointer = 11
        On Error GoTo errH
        Select Case Item.Tag
            Case "路径"
                Set objItem = tbcSub.InsertItem(Index, "临床路径", mcolSubForm("_路径").hwnd, 0)
                objItem.Tag = "路径"
            Case "医嘱"
                Set objItem = tbcSub.InsertItem(Index, "医嘱记录", mcolSubForm("_医嘱").hwnd, 0)
                objItem.Tag = "医嘱"
            Case "费用"
                Set objItem = tbcSub.InsertItem(Index, "费用记录", mcolSubForm("_费用").hwnd, 0)
                objItem.Tag = "费用"
            Case "病历"
                Set objItem = tbcSub.InsertItem(Index, "住院病历", mcolSubForm("_病历").hwnd, 0)
                objItem.Tag = "病历"
            Case "护理"
                Set objItem = tbcSub.InsertItem(Index, "护理记录", mcolSubForm("_护理").hwnd, 0)
                objItem.Tag = "护理"
            Case "护理病历"
                Set objItem = tbcSub.InsertItem(Index, "护理病历", mcolSubForm("_护理病历").hwnd, 0)
                objItem.Tag = "护理病历"
            Case "监护"
                Set objItem = tbcSub.InsertItem(Index, "护理监护", mcolSubForm("_监护").hwnd, 0)
                objItem.Tag = "监护"
            Case "新病历"
                Set objItem = tbcSub.InsertItem(Index, "电子病历", mcolSubForm("_新病历").hwnd, 0)
                objItem.Tag = "新病历"
            Case Else '整体护理页面
                Set mobjNurseForm = gobjNurseIntegrate.GetForm(Item.Tag, CStr(mColNurseFormUrl("_" & Item.Tag)))
                Set objItem = tbcSub.InsertItem(Index, Item.Tag, mobjNurseForm.hwnd, 0)
                objItem.Tag = Item.Tag
        End Select
        Call tbcSub.RemoveItem(Index + 1)
        objItem.Selected = True
        Screen.MousePointer = 0
        mblnTabTmp = False
    Else
        Set objItem = Item
    End If
    
    '刷新子窗体对应的CommandBar
    Call SubWinDefCommandBar(objItem)

    '刷新子窗体数据
    If Visible Then Call SubWinRefreshData(objItem)

    If Visible And mfrmActive.Visible And mfrmActive.Enabled Then mfrmActive.SetFocus
    tbcSub.Tag = Item.Tag   '记录上一次选择的卡片
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub UnLoadPageForm()
'加载一个页面GDI就会增长，为了控制GDI增长，切换页面时卸载上一个页面窗体(新版电子病历不处理)
'外挂借口中的窗体是直接绑定的也不用处理
    Dim i As Integer, blnFind As Boolean
    Dim Index As Long, objItem As TabControlItem
    Dim blnNurseIntegrate As Boolean
    '找到上一个选择页面的索引
    For i = 0 To tbcSub.ItemCount - 1
        If tbcSub.Item(i).Tag = tbcSub.Tag Then
            Index = tbcSub.Item(i).Index
            blnFind = True
            Exit For
        End If
    Next i
    If blnFind = False Then Exit Sub
    blnNurseIntegrate = Val(marrTabAttribute(Index)) = 1
    '排开新版病历和外挂接口窗体(Val(marrTabAttribute(Index)) = 1,为整体护理窗口)
    If InStr(1, "'路径'医嘱'费用'病历'护理'护理病历'监护'", "'" & tbcSub.Tag & "'") = 0 And blnNurseIntegrate = False Then Exit Sub
    '128211，1:从医嘱切换到其他页面，在定位病人文本框敲回车窗体就会失去焦点，查了1下午原因不明，医嘱本身占用GDI也不搞，暂时不进行卸载
    '              2:不从上面直接排开一方面是为了tab页面颜色，另一方是为了保证其他页面切换到医嘱点击文本框能全选内容，医嘱界面不知道怎么回事
    If tbcSub.Tag <> "医嘱" Then
        If UnloadSubForm(tbcSub.Tag, blnNurseIntegrate) = False Then Exit Sub
    End If
    
    Screen.MousePointer = 11
    mblnTabTmp = True
    On Error GoTo ErrHand
    Select Case tbcSub.Tag
        Case "路径"
            Set objItem = tbcSub.InsertItem(Index, "临床路径", picTmp.hwnd, 0)
            objItem.Tag = "路径"
        Case "医嘱"
            Set objItem = tbcSub.InsertItem(Index, "医嘱记录", picTmp.hwnd, 0)
            objItem.Tag = "医嘱"
        Case "费用"
            Set objItem = tbcSub.InsertItem(Index, "费用记录", picTmp.hwnd, 0)
            objItem.Tag = "费用"
        Case "病历"
            Set objItem = tbcSub.InsertItem(Index, "住院病历", picTmp.hwnd, 0)
            objItem.Tag = "病历"
        Case "护理"
            Set objItem = tbcSub.InsertItem(Index, "护理记录", picTmp.hwnd, 0)
            objItem.Tag = "护理"
        Case "护理病历"
            Set objItem = tbcSub.InsertItem(Index, "护理病历", picTmp.hwnd, 0)
            objItem.Tag = "护理病历"
        Case "监护"
            Set objItem = tbcSub.InsertItem(Index, "护理监护", picTmp.hwnd, 0)
            objItem.Tag = "监护"
        Case Else '整体护理
            Set objItem = tbcSub.InsertItem(Index, tbcSub.Tag, picTmp.hwnd, 0)
            objItem.Tag = tbcSub.Tag
    End Select
    Call tbcSub.RemoveItem(Index + 1)
    Screen.MousePointer = 0
    mblnTabTmp = False
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function UnloadSubForm(ByVal strTag As String, Optional blnNurseIntegrate As Boolean = False) As Boolean
'功能：卸载相关虚拟窗口
'参数：strTag：非整体护理窗口页签名; blnNurseIntegrate 是否是整体护理页签
    Dim objForm As Object
    On Error Resume Next
    err.Clear
    If blnNurseIntegrate = False Then
        If Not mcolSubForm("_" & strTag) Is Nothing Then
            Call SetOrGetSubFromOperation(strTag, False)  '窗体卸载之前记录窗体数据条件
            Unload mcolSubForm("_" & strTag)
        End If
    Else
        If Not mobjNurseForm Is Nothing Then Unload mobjNurseForm: Set mobjNurseForm = Nothing
    End If
    If err <> 0 Then err.Clear
    UnloadSubForm = True
    On Error GoTo 0
End Function

Private Sub SetOrGetSubFromOperation(ByVal strTag As String, ByVal blnSet As Boolean)
'设置或获取子窗体条件,相关虚拟模块提供统一接口
'       GetFormOperation() as string --获取窗体操作选择，该接口会在窗体卸载前调用
'       RestoreFormOperation(byval strValue as string)-恢复窗体操作选择，该接口会在虚拟窗体刷新前调用
'blnSet =TRUE 恢复子窗体条件设置(刷新前调用),=FALSE 获取子窗体条件设置(窗体卸载前调用)
    Dim strValue As String
    On Error Resume Next
    If blnSet = False Then mcolSubFormOperation.Remove "_" & strTag
    Select Case strTag
        Case "路径"
            If blnSet = False Then
                strValue = mclsPath.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsPath.RestoreFormOperation(strValue)
            End If
        Case "医嘱"
            If blnSet = False Then
                strValue = mclsAdvices.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsAdvices.RestoreFormOperation(strValue)
            End If
        Case "费用"
            If blnSet = False Then
                strValue = mclsFeeQuery.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsFeeQuery.RestoreFormOperation(strValue)
            End If
        Case "病历"   '住院病历
            If blnSet = False Then
                strValue = mclsEPRs.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsEPRs.RestoreFormOperation(strValue)
            End If
        Case "护理"
            If blnSet = False Then
                strValue = mclsTends.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsTends.RestoreFormOperation(strValue)
            End If
        Case "护理病历"
            If blnSet = False Then
                strValue = mclsTendEPRs.GetFormOperation()
                mcolSubFormOperation.Add strValue, "_" & strTag
            Else
                strValue = CStr(mcolSubFormOperation("_" & strTag))
                If strValue <> "" Then Call mclsTendEPRs.RestoreFormOperation(strValue)
            End If
    End Select
    If blnSet = True Then mcolSubFormOperation.Remove "_" & strTag
    
    If err <> 0 Then err.Clear
    On Error GoTo 0
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.tbcSub
        .Left = lngLeft: .Width = lngRight - lngLeft
        .Top = lngTop: .Height = lngBottom - lngTop
    End With
    
    With picPrompt
        .Top = stbThis.Top + 50
        .Height = stbThis.Height - 100
        .Left = stbThis.Panels(2).Left + 50
        .Width = stbThis.Panels(2).Width - 100
    End With
    With lblPrompt
        .Width = picPrompt.Width
        .Height = TextHeight("刘")
        .Top = (picPrompt.Height - .Height) \ 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long, strTmp As String
    Dim blnSetup As Boolean
    
    mlng病人ID = 0
    mlng主页ID = 0
    mlng病区ID = 0
    mblnShow = False
    
    blnSetup = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    If Not tbcSub.Selected Is Nothing Then
        Call zlDatabase.SetPara("医护功能", tbcSub.Selected.Tag, glngSys, p住院护士站, blnSetup)
    End If
    Call zlDatabase.SetPara("病人信息工具栏", IIf(mobjBar.Visible, 1, 0), glngSys, p住院护士站, blnSetup)
    Call SaveWinState(Me, App.ProductName)

    '强行Unload,不然不会激活子窗体的事件
    For i = 1 To mcolSubForm.Count
        Unload mcolSubForm(i)
    Next
    Set mcolSubForm = Nothing
    Set mcolSubFormOperation = Nothing
    If Not mobjNurseForm Is Nothing Then
        Unload mobjNurseForm
        Set mobjNurseForm = Nothing
    End If
    Set mColNurseFormUrl = Nothing
    Set mclsAdvices = Nothing
    Set mclsEMR = Nothing
    Set mclsEPRs = Nothing
    Set mclsTends = Nothing
    Set mclsTendEPRs = Nothing
    Set mclsFeeQuery = Nothing
    Set mclsWardMonitor = Nothing
    Set mclsPath = Nothing
    
    Set mfrmActive = Nothing
    Set mobjMipModule = Nothing
    
    mobjParent.mblnRoutine = mblnShow
    If Not mobjParent Is Nothing Then Set mobjParent = Nothing
    If Not mrsPati Is Nothing Then
        If mrsPati.State = adStateClosed Then mrsPati.Close
        Set mrsPati = Nothing
    End If
End Sub

Private Sub picInfo_GotFocus()
    If cboPages.Enabled And cboPages.Visible Then cboPages.SetFocus
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left * 2

    cbo过敏.Width = fraInfo.Width - cbo过敏.Left - 100
End Sub

Private Sub tbcSub_GotFocus()
    On Error Resume Next
    If Not mfrmActive Is Nothing Then mfrmActive.SetFocus
End Sub

Private Sub ClearPatiInfo()
'功能：清除单个病人相关的显示信息
    mlng病人ID = 0
    mlng主页ID = 0
    mlng婴儿科室ID = 0
    mlng婴儿病区ID = 0
    
    mPatiInfo.状态 = 0
    mPatiInfo.住院号 = ""
    mPatiInfo.床号 = ""
    mPatiInfo.主页ID = 0
    mPatiInfo.病区ID = 0
    mPatiInfo.科室ID = 0
    mPatiInfo.入院日期 = CDate(0)
    mPatiInfo.出院日期 = CDate(0)
    mPatiInfo.数据转出 = False
    mPatiInfo.产科 = False
    mPatiInfo.结清 = False
    mPatiInfo.险类 = 0
    mPatiInfo.性质 = 0

    cboPages.Clear
    cbo过敏.Clear

    lbl床号.Caption = ""
    lbl姓名.Caption = ""
    lbl性别.Caption = ""
    lbl年龄.Caption = ""
    lbl住院号.Caption = ""
    lbl医疗付款方式.Caption = ""
    lbl护理等级.Caption = ""
    lbl病况.Caption = ""
    lbl入院时间.Caption = ""
    lbl诊断.Caption = ""
End Sub

Function ExecuteMonitor() As Boolean
'功能：调用监护仪
    Dim strUser As String, strPass As String, strServer As String
    Dim arrInfo As Variant, i As Long

    'Provider=MSDataShape.1;Extended Properties="Driver={Microsoft ODBC for Oracle};Server=ORCL";Persist Security Info=True;User ID=zlhis;Password=HIS;Data Provider=MSDASQL
    'Provider=OraOLEDB.Oracle.1;Password=HIS;Persist Security Info=True;User ID=ZLHIS;Data Source=ORCL;Extended Properties="PLSQLRSet=1;DistribTx=0"
    arrInfo = Split(gcnOracle.ConnectionString, ";")
    For i = 0 To UBound(arrInfo)
        If UCase(arrInfo(i)) Like UCase("User ID=*") Then
            strUser = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Password=*") Then
            strPass = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Data Source=*") Then
            strServer = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
        ElseIf UCase(arrInfo(i)) Like UCase("Server=*") Then
            strServer = Mid(arrInfo(i), InStr(arrInfo(i), "=") + 1)
            strServer = Replace(strServer, """", "")
        End If
    Next

    On Error GoTo errH

    Shell mstrMonitor & " " & strUser & " " & strPass & " " & strServer, vbNormalFocus

    ExecuteMonitor = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AddPages()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lng病人ID As Long, lng主页ID As Long
    Dim bln留观 As Boolean
    On Error GoTo ErrHand
    
    '清除病人相关信息
    lng病人ID = mlng病人ID: lng主页ID = mlng主页ID
    Call ClearPatiInfo
    mlng病人ID = lng病人ID: mlng主页ID = lng主页ID
    '根据病人ID读取该病人的住院次数
    '52004,刘鹏飞,2012-08-10,住院次数应该默认定位到当前病人当前住院次数
    strSQL = " Select 主页ID,病人性质 From 病案主页 Where 主页ID<>0 And 病人ID=[1] Order by 主页ID Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取住院次数", mlng病人ID)
    
    cboPages.Clear
    Do While Not rsTemp.EOF
        cboPages.AddItem "第 " & rsTemp!主页ID & " 次" & IIf(Val("" & rsTemp!病人性质) = 1, "(门诊留观)", IIf(Val("" & rsTemp!病人性质) = 2, "(住院留观)", ""))
        cboPages.ItemData(cboPages.NewIndex) = rsTemp!主页ID
        If rsTemp!主页ID = mlng主页ID Then
            Call Cbo.SetIndex(cboPages.hwnd, cboPages.NewIndex)
        End If
        If bln留观 = False And Val("" & rsTemp!病人性质) <> 0 Then bln留观 = True
        rsTemp.MoveNext
    Loop
    If cboPages.ListIndex = -1 Then
        Call Cbo.SetIndex(cboPages.hwnd, 0)
    End If
    If bln留观 = True Then
        Call Cbo.SetListWidth(cboPages.hwnd, 2000)
    End If
    Call cboPages_Click
    '52638,刘鹏飞,2012-08-13,加载病人过敏药物信息
    Call LoadPatiAllergy(mlng病人ID, cbo过敏)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txt病人_GotFocus()
    Call zlControl.TxtSelAll(txt病人)
End Sub

Private Sub txt病人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    Dim strOrder As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    strInput = Trim(txt病人.Text)
    If strInput = "" Then Exit Sub
    
    strOrder = mrsPati.Sort
    strInput = " 床号='" & zlStr.Lpad(strInput, 10) & "'"
    mrsPati.Filter = strInput
    If mrsPati.RecordCount = 0 Then
        If Not IsNumeric(Trim(txt病人.Text)) Then
            strInput = " 姓名='" & Trim(txt病人.Text) & "'"
        Else
            strInput = " 住院号=" & Trim(txt病人.Text)
        End If
        mrsPati.Filter = strInput
        
        If mrsPati.RecordCount = 0 Then
            '再按姓名简码过滤一次,不提供弹出选择的功能,要求尽可能输入详细
            mrsPati.Sort = "简码"
            mrsPati.Filter = "简码 LIKE '*" & UCase(Trim(txt病人.Text)) & "*'"
            If mrsPati.RecordCount = 0 Then
                mrsPati.Filter = 0
                mrsPati.Sort = strOrder
                MsgBox "未找到该病人的有效数据，请重新输入！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    mlng病人ID = mrsPati!病人ID
    mlng主页ID = mrsPati!主页ID
    mrsPati.Filter = 0
    mrsPati.Sort = strOrder
    mrsPati.MoveFirst
    mrsPati.Find "Key='" & mlng病人ID & ":" & mlng主页ID & "'"
    
    mintPrePage = -1
    Call AddPages
    
    picPati.Visible = False
End Sub

Private Sub mclsAdvices_DoByAdvice(ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal lngWayID As Long, ByVal strTag As String)
'功能：对医嘱记帐  lngWayID＝conMenu_Edit_AdvicePrice
    Dim lngTmp As Long
    lngTmp = IIf(lng相关ID = 0, lng医嘱ID, lng相关ID)
    Call mclsFeeQuery.zlPatiBilling(Me, mlng病人ID, mlng病区ID, mlng主页ID, Val("" & mPatiInfo.科室ID), False, lngTmp)
End Sub

Private Sub GeNurseRelatedUnitID()
    Dim strErrMsg As String
    '病人切换是获取
    If gbln启用整体护理接口 = True Then
        If InitNurseIntegrate = True Then
            If gobjNurseIntegrate.GetRelatedIDToGUID(mlng病区ID, strErrMsg, mlng病人ID & "|" & mlng主页ID) = False Then
                MsgBox "获取整体护理病区ID接口调用失败！" & vbCrLf & "详细信息：" & strErrMsg, vbInformation, gstrSysName
            Else
                mstrRelatedUnitID = gobjNurseIntegrate.RelatedUnitID
                mstrRelatedUserID = gobjNurseIntegrate.RelatedUserID
                mstrNurseParentID = gobjNurseIntegrate.RelatedPatientID
            End If
        End If
    End If
End Sub

