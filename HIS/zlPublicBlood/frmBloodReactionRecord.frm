VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmBloodReactionRecord 
   Caption         =   "输血反应记录"
   ClientHeight    =   11895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14370
   Icon            =   "frmBloodReactionRecord.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11895
   ScaleWidth      =   14370
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer TimNotify 
      Interval        =   500
      Left            =   1560
      Top             =   0
   End
   Begin VB.PictureBox picTips 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   5400
      ScaleHeight     =   1455
      ScaleWidth      =   1815
      TabIndex        =   42
      Top             =   8520
      Width           =   1815
      Begin XtremeReportControl.ReportControl rptTips 
         Height          =   615
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   735
         _Version        =   589884
         _ExtentX        =   1296
         _ExtentY        =   1085
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox PicPane 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   9465
      ScaleHeight     =   330
      ScaleWidth      =   4185
      TabIndex        =   22
      Top             =   11595
      Width           =   4185
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "输血科待提交"
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
         Index           =   2
         Left            =   2910
         TabIndex        =   28
         Top             =   60
         Width           =   1170
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "医生待提交"
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
         Index           =   1
         Left            =   1455
         TabIndex        =   27
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "已完成"
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
         Index           =   0
         Left            =   390
         TabIndex        =   26
         Top             =   60
         Width           =   585
      End
      Begin VB.Label lbl标签状态 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "■"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   2
         Left            =   2535
         TabIndex        =   25
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lbl标签状态 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000002&
         Caption         =   "■"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lbl标签状态 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "■"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   285
         Index           =   0
         Left            =   30
         TabIndex        =   23
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   4440
      ScaleHeight     =   6495
      ScaleWidth      =   8895
      TabIndex        =   17
      Top             =   540
      Width           =   8895
      Begin zlPublicBlood.usrCardEdit UCE 
         Height          =   8715
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   8805
         _extentx        =   15531
         _extenty        =   15372
      End
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10050
      Left            =   120
      ScaleHeight     =   10050
      ScaleWidth      =   4575
      TabIndex        =   14
      Top             =   465
      Width           =   4575
      Begin VB.ComboBox cbo1 
         Height          =   300
         Left            =   570
         TabIndex        =   0
         Text            =   "所有科室"
         Top             =   300
         Width           =   2700
      End
      Begin VB.PictureBox picTmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   3480
         ScaleHeight     =   165
         ScaleWidth      =   585
         TabIndex        =   34
         Top             =   8160
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.PictureBox picUCP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2460
         Left            =   60
         ScaleHeight     =   2460
         ScaleWidth      =   3855
         TabIndex        =   21
         Top             =   5595
         Width           =   3855
         Begin zlPublicBlood.usrCardPeople UCP 
            Height          =   2055
            Left            =   30
            TabIndex        =   12
            Top             =   0
            Width           =   3135
            _extentx        =   5530
            _extenty        =   3625
         End
      End
      Begin XtremeSuiteControls.TabControl tbcthis 
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   4530
         Width           =   3855
         _Version        =   589884
         _ExtentX        =   6800
         _ExtentY        =   1296
         _StockProps     =   64
      End
      Begin VB.Frame Fra1 
         Caption         =   "过滤条件"
         Height          =   3795
         Left            =   75
         TabIndex        =   15
         Top             =   660
         Width           =   3855
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   75
            TabIndex        =   36
            Top             =   3050
            Width           =   3615
            Begin VB.OptionButton opt 
               Caption         =   "未填"
               Height          =   255
               Index           =   3
               Left            =   2760
               TabIndex        =   44
               Top             =   0
               Width           =   735
            End
            Begin VB.OptionButton opt 
               Caption         =   "所有"
               Height          =   255
               Index           =   2
               Left            =   840
               TabIndex        =   39
               Top             =   0
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.OptionButton opt 
               Caption         =   "无"
               Height          =   255
               Index           =   0
               Left            =   2160
               TabIndex        =   38
               Top             =   0
               Width           =   615
            End
            Begin VB.OptionButton opt 
               Caption         =   "有"
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   37
               Top             =   0
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "输血反应"
               Height          =   180
               Left            =   20
               TabIndex        =   41
               Top             =   10
               Width           =   720
            End
         End
         Begin VB.TextBox TXTDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   930
            MaxLength       =   4
            TabIndex        =   6
            Text            =   "7"
            Top             =   1275
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Frame frmLine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   930
            TabIndex        =   32
            Top             =   1470
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.ComboBox cbotime 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   210
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.CommandButton cmd2 
            Caption         =   "刷新"
            Height          =   300
            Left            =   2955
            TabIndex        =   11
            Top             =   3360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox ChkRection 
            Caption         =   "根据输血反应填写情况过滤"
            Height          =   225
            Left            =   120
            TabIndex        =   7
            Top             =   1620
            Width           =   2475
         End
         Begin MSComCtl2.DTPicker DTP2 
            Height          =   300
            Left            =   945
            TabIndex        =   2
            Top             =   2310
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   283049987
            CurrentDate     =   42593
         End
         Begin MSComCtl2.DTPicker DTP1 
            Height          =   300
            Left            =   945
            TabIndex        =   1
            Top             =   1965
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   283049987
            CurrentDate     =   42593
         End
         Begin VB.CheckBox chk1 
            Caption         =   "未提交"
            Height          =   225
            Index           =   2
            Left            =   2835
            TabIndex        =   10
            Top             =   2730
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chk1 
            Caption         =   "已提交"
            Height          =   225
            Index           =   1
            Left            =   1965
            TabIndex        =   9
            Top             =   2730
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chk1 
            Caption         =   "本人填写"
            Height          =   225
            Index           =   0
            Left            =   915
            TabIndex        =   8
            Top             =   2730
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DTP4 
            Height          =   300
            Left            =   930
            TabIndex        =   5
            Top             =   915
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   283049987
            CurrentDate     =   42711
         End
         Begin MSComCtl2.DTPicker DTP3 
            Height          =   300
            Left            =   930
            TabIndex        =   4
            Top             =   585
            Visible         =   0   'False
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   283049987
            CurrentDate     =   42711
         End
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            Caption         =   "反应时间"
            Height          =   180
            Left            =   120
            TabIndex        =   40
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label lbl6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "显示最近       天转出的病人"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   33
            Top             =   1260
            Visible         =   0   'False
            Width           =   2430
         End
         Begin VB.Label lbl5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "出院日期"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   270
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl8 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "开始时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   630
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl9 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "结束时间"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   120
            TabIndex        =   29
            Top             =   990
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label lbl4 
            Caption         =   "~"
            Height          =   135
            Left            =   720
            TabIndex        =   16
            Top             =   2460
            Width           =   135
         End
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "科室"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   11535
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2461
            MinWidth        =   882
            Picture         =   "frmBloodReactionRecord.frx":07AA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11774
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7585
            MinWidth        =   7585
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
   Begin VSFlex8Ctl.VSFlexGrid VSFBRlist 
      Height          =   855
      Left            =   3000
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   2580
      _cx             =   4551
      _cy             =   1508
      Appearance      =   2
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483638
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpPeoPle 
      Bindings        =   "frmBloodReactionRecord.frx":1290
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmBloodReactionRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngSys As Long   '调用系统号
Private mlngModule As Long           '调用模块号
Private mlng阶段 As Long             '0-门诊处理阶段、1-住院处理阶段  2-输血科处理阶段，对应clspublicblood模块中的场合
Private mstr科室 As String
Private mstr开始时间 As String
Private mstr结束时间 As String
Private mstr填写人 As String
Private mlng提交状态 As Long         '0-全部数据，1-未提交数据，2-已提交数据
Private mArr过滤数据                 '存放科室、时间、是否本人填写、提交状态
Private mstrPrivs As String          '权限串
Private mblnButtonChecked As Boolean '标准按钮
Private mblnTextChecked As Boolean   '文本标签
Private mblnSizeChecked As Boolean   '工具栏大小
Private mblnStatuChecked As Boolean  '状态栏显示
Private mfrmMain As Object           '父窗体
Private mRsBR As ADODB.Recordset     '病人信息记录集
Private mblnHaveBR As Boolean        '判断是否查询到病人
Private mlngtbcIndex As Long         '记录tbcthis选中卡片的index
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mblnStart As Boolean         '判断程序是否开始
Private mArrPosition(0 To 2)         '存放病人在不同状态下的定位数据，比如在院、出院、转出。
Private mArrCheckData(0 To 9)        '存放控件的当前数据0~9对应，cbo1、dtp1~4、chk1(0~2)、cbotime、TXTDay
Private mblnIsSelect As Boolean      '判断是否有病人被选中 true表示有，false表示无
Private mblnIsSubmit As Boolean      '判断是否提交或者回退，用于提交回退后屏蔽保存和取消使能
Public mblnBloodReactionRecordIsOpen As Boolean '非模态状态下，判断窗体是否开启
Private mstrFindKey As String             '输血科新增时，通过查询页面查询到的病人,格式：病人ID-就诊ID-类型(0-住院/2-门诊)
Private mblnADDPeoPle As Boolean     '输血科新增时，判断是新增病人还是定位到病人，新增病人就会重新读取病人查找的sql，定位病人则直接选中查询的病人。
Private mintDeptIndex As Integer  '部门索引
Private mintNotify As Integer '提醒自动刷新间隔(分钟)
Private mblnFirst As Boolean         '第一次启动立即刷新,或切换部门、消息改变时强制刷新

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：初始化CommandBar
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    '------------------------------------------------------------------------------------------------------------------
    '初始化处理
    
    Call CommandBarInit(cbsMain)
    '菜单定义:包括公共部份
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '文件
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.id = conMenu_FilePopup
'    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "预览(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "打印(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "打印设置")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_File_MedRec, "反应记录打印")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_File_MedRecSetup, "打印设置"): objControl.IconId = conMenu_File_PrintSet
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_File_MedRecPreview, "预览反应"): objControl.IconId = conMenu_File_Preview
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_File_MedRecPrint, "打印反应"): objControl.IconId = conMenu_File_Print
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Parameter, "参数设置", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    '编辑
    
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.id = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "新增")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Save, "保存", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "提交")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "回退")
    
    '查看
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.id = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_FindNext, "查找下一个(&N)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    
    mblnButtonChecked = True
    mblnTextChecked = True
    mblnSizeChecked = True
    mblnStatuChecked = True
    
    '帮助
    '------------------------------------------------------------------------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.id = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): objControl.BeginGroup = True
    End With
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched

    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_MedRecPreview, "预览"): objControl.ToolTipText = "预览输血反应单": objControl.IconId = conMenu_File_Preview
        Set objControl = .Add(xtpControlButton, conMenu_File_MedRecPrint, "打印"): objControl.ToolTipText = "打印输血反应单": objControl.IconId = conMenu_File_Print
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "增加"): objControl.BeginGroup = True '
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Modify, IIf(mlng阶段 = 2, "补填", "修改"))
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, IIf(mlng阶段 = 2, "完成", "提交")): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "回退")
        Set objControl = .Add(xtpControlButton, conMenu_View_Detail, "输血执行"): objControl.ToolTipText = "输血执行情况查阅": objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")

    End With
    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlButton Then objControl.Style = xtpButtonIconAndCaption
    Next
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理
    
    With cbsMain.KeyBindings
        .Add 0, vbKeyDelete, conMenu_Edit_Delete            '删除
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify          '修改
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh                 '刷新
        .Add FCONTROL, vbKeyF, conMenu_View_Find            '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext              '继续查找
        .Add FCONTROL, vbKeyS, conMenu_Edit_Save            '保存
        .Add FCONTROL, vbKeyC, conMenu_Edit_Transf_Cancle   '取消
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
        .Add 0, vbKeyF12, conMenu_File_Parameter             '参数设置
        .Add FCONTROL, vbKeyX, conMenu_File_Exit            '退出
    End With
    
    Call gobjDatabase.ShowReportMenu(Me, 2200, p输血反应管理, mstrPrivs)
    InitCommandBar = True
    
    Exit Function
ErrHand:
    
End Function

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long
    With rptTips
        Set objCol = .Columns.Add(0, "病人姓名", 30, True)
        Set objCol = .Columns.Add(1, "消息内容", 60, True)
        Set objCol = .Columns.Add(2, "申请ID", 40, True): objCol.Visible = False
        Set objCol = .Columns.Add(3, "病人ID", 40, True): objCol.Visible = False
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
            objCol.Sortable = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .HideSelection = True
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有提醒内容..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
    End With
End Sub
Private Sub cbo1_Click()
    If mblnStart = False Then Exit Sub
    If mintDeptIndex = cbo1.ListIndex Then Exit Sub
    mintDeptIndex = cbo1.ListIndex
    RefreshBR
    mblnFirst = True
End Sub

Private Sub cbo1_KeyPress(KeyAscii As Integer)
    '功能：将小写转化为大写并查询匹配的部门显示，最后跳转到下一个控件。
    Dim olddata As String
    Dim intIndex As Integer
    olddata = cbo1.Text
    If KeyAscii > 96 And KeyAscii < 123 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = vbKeyReturn Then
        intIndex = findDepart(cbo1.Text)
        If intIndex = -1 Then '找不到则复原
            cbo1.ListIndex = mintDeptIndex
            gobjCommFun.PressKey vbKeyTab
        Else
            cbo1.ListIndex = intIndex
        End If
    End If
End Sub

Private Function findDepart(key As String) As Long
    '功能：查找部门列表中符合条件的部门
    Dim lngi As Long
    Dim blnfind As Boolean
    Dim ArrCbo
    
    For lngi = 0 To cbo1.ListCount - 1
        If cbo1.List(lngi) Like key & "*" Then
            blnfind = True
            findDepart = lngi
            Exit For
        End If
        ArrCbo = Split(cbo1.List(lngi), "-")
        If ArrCbo(0) Like key & "*" Then
            blnfind = True
            findDepart = lngi
            Exit For
        ElseIf UBound(ArrCbo) > 0 Then
            If ArrCbo(1) Like key & "*" Then
                blnfind = True
                findDepart = lngi
                Exit For
            End If
        End If
    Next
    If blnfind = False Then
        findDepart = -1
    End If
End Function

Private Sub Form_Activate()
    Set gobjFScrollBar = UCP.FScrollBar
    glngBooldPepWinProc = GetWindowLong(UCP.objPicBack.hWnd, GWL_WNDPROC)
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
    SetWindowLong UCP.objPicBack.hWnd, GWL_WNDPROC, glngBooldPepWinProc
End Sub

Private Sub rptTips_SelectionChanged()
    Dim lng收发id As Long, lng病人ID As Long, lng就诊id As Long, int病人来源 As Integer
    Dim strSQL As String, rs As Recordset, bytMode As Byte
    Dim strKey As String, lng过滤数据4 As Long
    If Not Visible Then Exit Sub
    lng收发id = Val(rptTips.SelectedRows(0).Record(2).Value)
    lng病人ID = Val(rptTips.SelectedRows(0).Record(3).Value)
    lng就诊id = Val(rptTips.SelectedRows(0).Record(4).Value)
    int病人来源 = IIf(Val(rptTips.SelectedRows(0).Record(5).Value) = 2, 0, 1)
    lng过滤数据4 = mArr过滤数据(4)
    If mArr过滤数据(4) = 0 Then mArr过滤数据(4) = 2
    strKey = lng病人ID & "-" & lng就诊id
    mblnADDPeoPle = Not UCP.findIdPeoPle(strKey, False)
    If mblnADDPeoPle Then
        mstrFindKey = lng病人ID & "-" & lng就诊id & "-" & int病人来源
        Call ExecuteCommand("刷新数据")
        If UCP.findIdPeoPle(strKey, False) Then
            If Not UCE.BloodLocation(lng收发id) Then Call MsgBox("未找到对应血液。", vbInformation, Me.Caption)
        Else
            Call MsgBox("未找到对应病人。", vbInformation, Me.Caption)
        End If
    Else
        If Not UCE.BloodLocation(lng收发id) Then Call MsgBox("未找到对应血液。", vbInformation, Me.Caption)
    End If
    mblnADDPeoPle = False
    mArr过滤数据(4) = lng过滤数据4
End Sub
Private Sub TimNotify_Timer()
    Static strPreTime1 As String
    Dim curTime As Date
    curTime = Now
    '刷新提醒
    If mintNotify > 0 Then
        If strPreTime1 = "" Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
        End If
        If DateDiff("s", CDate(strPreTime1), curTime) > mintNotify * CLng(60) Or mblnFirst Then
            strPreTime1 = Format(curTime, "yyyy-MM-dd HH:mm:ss")
            Call ExecuteCommand("刷新提示")
            mblnFirst = False
        End If
     Else
        If mblnFirst = True Then
            Call ExecuteCommand("刷新提示")
            mblnFirst = False
        End If
    End If
End Sub
Private Sub cbo1_LostFocus()
    If cbo1.Text <> cbo1.List(cbo1.ListIndex) Then cbo1.Text = cbo1.List(cbo1.ListIndex)
End Sub

Private Sub cboTime_Click()
    Dim blnEnable As Boolean, strCurDate As String
    
    blnEnable = Val(cbotime.ItemData(cbotime.ListIndex)) = -1
    strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    DTP3.Value = Format(CDate(strCurDate) - Val(cbotime.ItemData(cbotime.ListIndex)), "YYYY-MM-DD")
    DTP4.Value = Format(strCurDate, "YYYY-MM-DD") & " 23:59:59"
    DTP3.Enabled = blnEnable
    DTP4.Enabled = blnEnable
End Sub

Private Sub cboTime_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strKey  As String
    Dim arrKey
    Select Case Control.id
        Case conMenu_File_Preview '预览病人列表
            Call zlRptPrint(2, VSFBRlist, "病人列表")
        Case conMenu_File_Print '打印病人列表
            Call zlRptPrint(1, VSFBRlist, "病人列表")
        Case conMenu_File_PrintSet '打印设置
            Call zlPrintSet
        Case conMenu_File_MedRecSetup '输血反应报表打印设置
            Call UCE.showPrintSet
        Case conMenu_File_MedRecPreview '输血反应打印预览
            If UCE.lngFYCount = 1 Then
                Call UCE.ShowPrint(2)
            Else
                Call UCE.ShowPrintList(2)
            End If
        Case conMenu_File_MedRecPrint '输血反应打印
            If UCE.lngFYCount = 1 Then
                Call UCE.ShowPrint(1)
            Else
                Call UCE.ShowPrintList(1)
            End If
        Case conMenu_Edit_NewItem: '新增
            If mlng阶段 = 2 Then
                strKey = frmBloodPeoPleSerch.SerchPeople(mfrmMain, mlngModule)
                mstrFindKey = strKey
                If strKey <> "" Then
                    arrKey = Split(strKey, "-")
                    strKey = arrKey(0) & "-" & arrKey(1)
                    mblnADDPeoPle = Not UCP.findIdPeoPle(strKey, False)
                    If mblnADDPeoPle = True Then '如果现有的病人列表里面没有要找到的病人则新增该病人
                        Call ExecuteCommand("刷新数据")
                        If UCP.findIdPeoPle(strKey, False) = True Then
'                            UCE.DataChanged = True
                            UCE.AddPage
                            mblnIsSubmit = False
                        Else
                            mstrFindKey = ""
                        End If
                        mblnADDPeoPle = False
                    Else
                        '如果当前病人列表有查询的数据，则跳到改病人，并新增页面
                        UCE.AddPage
                        mblnIsSubmit = False
                    End If
                    UCP.locked = IIf(mstrFindKey = "", False, True) '在新增状态下锁定ucp控件
                End If
            Else
                UCE.AddPage
                mblnIsSubmit = False
                UCP.locked = True
            End If
        Case conMenu_Edit_Modify: '修改
            UCE.ShowModify
            mblnIsSubmit = False
        Case conMenu_Edit_Delete: '删除
            If IsPrivs(mstrPrivs, "删除他人") = False Then
                If UCE.Doctor <> "" And UCE.Doctor <> UserInfo.姓名 Then
                    MsgBox "您没有权限删除他人记录的数据！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            UCE.ShowDelete
            mblnIsSubmit = True
        Case conMenu_Edit_Save: '保存
            If UCE.ShowSave = False Then Exit Sub
            If mlng阶段 = 2 Then mblnFirst = True
            UCP.locked = False '保存后取消锁定
        Case conMenu_Edit_Transf_Cancle: '取消
            UCE.ShowCancel
            UCP.locked = False '取消后取消锁定
            Call ExecuteCommand("刷新数据")
        Case conMenu_Edit_Audit: '提交
            UCE.SubmitData
            mblnIsSubmit = True
        Case conMenu_Edit_Untread: '回退
            UCE.ShowUntread
            mblnIsSubmit = True
        Case conMenu_View_Detail '执行情况查看
            Call frmBloodExecEdit.ViewExecution(Me, UCE.BloodID)
        Case conMenu_View_Refresh: '刷新
            Call ExecuteCommand("刷新数据")
        Case conMenu_View_ToolBar_Button
            mblnButtonChecked = Not mblnButtonChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_ToolBar_Text
            mblnTextChecked = Not mblnTextChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_ToolBar_Size
            mblnSizeChecked = Not mblnSizeChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_View_StatusBar
            mblnStatuChecked = Not mblnStatuChecked
            Call CommandBarExecutePublic(Control, Me)
        Case conMenu_File_Parameter
            Call ExecuteCommand("本地参数设置")
        Case conMenu_Help_Help              '帮助主题
            Call gobjComlib.ShowHelp(App.ProductName, Me.hWnd, Me.name, Int((2200) / 100))
        Case conMenu_Help_Web_Home 'Web上的中联
            Call gobjComlib.zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Forum         'Web上的论坛
            Call gobjComlib.zlWebForum(Me.hWnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call gobjComlib.zlMailTo(Me.hWnd)
        Case conMenu_Help_About '关于
            Call gobjComlib.ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Exit '退出
            Unload Me
        Case conMenu_View_Find, conMenu_View_FindNext '查找，继续查找
            Call UCP.FindPatiByVbKey(Control.id = conMenu_View_FindNext)
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then
        Bottom = Me.stbThis.Height
        PicPane.Visible = True
        PicPane.Top = Me.stbThis.Top + 60
        PicPane.Left = stbThis.Panels(6).Left + 120
    Else
        PicPane.Visible = False
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case conMenu_File_MedRec, conMenu_File_MedRecSetup, conMenu_File_MedRecPreview, conMenu_File_MedRecPrint
            Control.Visible = IsPrivs(mstrPrivs, "单据打印")
            Control.Enabled = Control.Visible
        Case conMenu_Edit_Modify: '修改
            '无记录反应的权限则修改按钮不可见。
            Control.Visible = IsPrivs(mstrPrivs, "记录反应")
            
            Control.Caption = IIf(UCE.输血科新增 = False And mlng阶段 = 2, "补填", "修改")
            '医生阶段已提交状态 或者 输血科已提交状态 或者 医生阶段输血科新增的数据 或者 新增状态 或者 修改状态 或者 无病人数据 或者 未选中病人的情况下修改按钮不使能，其他情况使能。
            Control.Enabled = Not ((mlng阶段 <> 2 And UCE.lng状态 <> 0) Or (mlng阶段 = 2 And UCE.lng状态 = 2) Or (mlng阶段 <> 2 And UCE.输血科新增 = True) Or UCE.strST = 新增 Or UCE.strST = 修改 Or mblnHaveBR = False Or mblnIsSelect = False)
            '补填时如果病人无输血则不用处理
            If mlng阶段 = 2 And Control.Enabled And Control.Caption = "补填" Then
                Control.Enabled = UCE.有无输血反应
            End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Delete: '删除
            '无删除记录的权限或者是输血科阶段，则删除按钮不可见。
            Control.Visible = IsPrivs(mstrPrivs, "删除记录")   'and not (mlng阶段=2 and not IsPrivs(mstrPrivs, "输血科新增"))，由于需求，输血科在有相关权限的情况下允许删除
            '医生阶段已提交状态 或者 输血科阶段已提交状态 或者 医生阶段输血科新增的数据 或者 新增状态 或者 修改状态 或者 无病人数据 或者 未选中病人的情况下删除按钮不使能，其他情况使能。
            Control.Enabled = Not ((mlng阶段 <> 2 And UCE.lng状态 <> 0) Or (mlng阶段 = 2 And UCE.lng状态 <> 0) Or (mlng阶段 <> 2 And UCE.输血科新增 = True) Or UCE.strST = 新增 Or UCE.strST = 修改 Or mblnHaveBR = False Or mblnIsSelect = False)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_NewItem: '新增
            '输血科没有新增权限时，新增按钮不可见，其他情况新增按钮可见。
            Control.Visible = IsPrivs(mstrPrivs, "记录反应") And Not (mlng阶段 = 2 And Not IsPrivs(mstrPrivs, "输血科新增")) '由于需求，输血科在有相关权限的情况下允许新增
            '如果blnAddPage=true或者没有病人或者没有选中病人时，新增不使能
            Control.Enabled = Not (UCE.blnAddPage = True Or mblnHaveBR = False Or mblnIsSelect = False)
            'blnAddPage=true 或者当前是新增或者修改状态或者当前无病人或者未选中病人时，新增按钮不使能，其他情况可以进行新增。
'            Control.Enabled = Not (UCE.strST = 新增 Or UCE.strST = 修改 Or mblnHaveBR = False Or mblnIsSelect = False) 'UCE.blnAddPage = True Or
            '例外：当输血科阶段，用户有输血科新增权限时，即使病人列表无病人或者未选中病人，也还是可以进行新增操作。
            If mlng阶段 = 2 And (mblnHaveBR = False Or mblnIsSelect = False) Then
                Control.Enabled = True
            End If
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Save: '保存
            Control.Visible = IsPrivs(mstrPrivs, "记录反应")
            '未提交且数据变化时，保存使能
            Control.Enabled = UCE.DataChanged And mblnIsSubmit = False
            
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Transf_Cancle: '取消
            Control.Visible = IsPrivs(mstrPrivs, "记录反应")
            '未提交且数据变化时，取消使能
            Control.Enabled = UCE.DataChanged And mblnIsSubmit = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Audit: '提交
            Control.Visible = IsPrivs(mstrPrivs, "提交回退")
            
            Control.Caption = IIf(mlng阶段 = 2, "完成", "提交")
            '医生阶段已提交数据 或者 输血科阶段已提交数据 或者 医生阶段输血科新增状态 或者 在新增或修改状态 或者 无病人或者未选中病人时提交不使能，其他状态提交使能。
            Control.Enabled = Not ((mlng阶段 <> 2 And UCE.lng状态 <> 0) Or (mlng阶段 = 2 And UCE.lng状态 = 2) Or (mlng阶段 <> 2 And UCE.输血科新增 = True) Or UCE.strST = 新增 Or UCE.strST = 修改 Or mblnHaveBR = False Or mblnIsSelect = False)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Edit_Untread: '回退
            Control.Visible = IsPrivs(mstrPrivs, "提交回退")
            '医生阶段非医生已提交状态 或者 输血科阶段未提交状态 或者 医生阶段输血科新增数据 或者 输血科阶段输血科新增数据未提交状态 或者 在新增或修改状态 或者 无病人或者未选中病人 时回退不使能，其他状态回退使能。
            Control.Enabled = Not ((mlng阶段 <> 2 And UCE.lng状态 <> 1) Or (mlng阶段 = 2 And UCE.lng状态 <> 2) Or (mlng阶段 <> 2 And UCE.输血科新增 = True) Or (mlng阶段 = 2 And UCE.lng状态 = 0 And UCE.输血科新增) Or UCE.strST = 新增 Or UCE.strST = 修改 Or mblnHaveBR = False Or mblnIsSelect = False)
        Case conMenu_View_Detail
            Control.Enabled = UCE.BloodID > 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_File_Parameter     '医生暂时无须用到参数设置
            Control.Visible = mlng阶段 = 2
            Control.Enabled = mlng阶段 = 2
        Case conMenu_View_ToolBar_Button
            Control.Checked = mblnButtonChecked
        Case conMenu_View_ToolBar_Text
            Control.Checked = mblnTextChecked
            
        Case conMenu_View_ToolBar_Size
            Control.Checked = mblnSizeChecked
            
        Case conMenu_View_StatusBar
            stbThis.Visible = mblnStatuChecked
            Control.Checked = mblnStatuChecked
    End Select
End Sub

Private Sub chk1_Click(Index As Integer)
    If chk1(1).Value = Unchecked And chk1(2).Value = Unchecked Then
        If Index = 1 Then
            chk1(2).Value = Checked
        ElseIf Index = 2 Then
            chk1(1).Value = Checked
        End If
    End If
End Sub


Private Sub RefreshBR()
    Dim strCurDate As String
    If cbo1.ListIndex = -1 Then Exit Sub
    
    If mblnStart = True Then
        If Not Me.ActiveControl Is Nothing Then
            Select Case Me.ActiveControl.name
                Case "DTP1", "DTP2", "DTP3", "DTP4"
                    Call gobjControl.ControlSetFocus(picUCP)
            End Select
        End If
    End If
    
    mArr过滤数据(0) = Val(cbo1.ItemData(cbo1.ListIndex)) '选中的科室id，-1表示所有科室
    mArr过滤数据(2) = IIf(chk1(0).Value = Checked, UserInfo.姓名, "") '是否本人填写,如果是则取填写人，否为空,余浪改，以前checked是true，这是导致勾选本人填写无效的主要原因
    If opt(0).Value Then mArr过滤数据(4) = 0
    If opt(1).Value Then mArr过滤数据(4) = 1
    If opt(2).Value Then mArr过滤数据(4) = 2
    If opt(3).Value Then mArr过滤数据(4) = 3
    
    If mlng阶段 = 2 Then
        mstr开始时间 = DTP1.Value
        mstr结束时间 = DTP2.Value
        mArr过滤数据(1) = mstr开始时间 & "'" & mstr结束时间 '过滤时间
    Else
        If (mlngtbcIndex = 2 And mlng阶段 = 1) Or (mlng阶段 = 0 And mlngtbcIndex = 1) Then
            strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
            If cbotime.ItemData(cbotime.ListIndex) = -1 Then
                mstr开始时间 = DTP3.Value
                mstr结束时间 = DTP4.Value
            Else
                mstr开始时间 = Format(CDate(strCurDate) - Val(cbotime.ItemData(cbotime.ListIndex)), "YYYY-MM-DD")
                mstr结束时间 = Format(strCurDate, "YYYY-MM-DD") & " 23:59:59"
            End If
        End If
        mArr过滤数据(1) = mstr开始时间 & "'" & mstr结束时间
    End If
    Call ExecuteCommand("刷新数据")
End Sub

Public Sub BloodReactionRecord(frmMain As Variant, lng阶段 As Long, ByVal lngSys As Long, ByVal lngMoudle As Long, Optional strPrivs As String, Optional lngisModul As Long = 0)
    '功能：输血反应记录的调用函数
    '参数：frmMain-父窗体
    '      lng阶段-0:门诊医生处理阶段1:住院医生处理阶段、2：输血科处理阶段
    '      lngMoudle-模块号
    '      strPrivs-权限串
    '      lngisModul-0-非模态、1-模态
    Dim strSQL As String
    Dim rs部门 As ADODB.Recordset
    Dim lngi As Long
    Dim rs上级部门 As ADODB.Recordset
    Dim objPane As Object
    Dim lngIndex As Long
    Dim strCurDate As String
    
    mblnFirst = True
    ReDim mArr过滤数据(0 To 4)
    lngIndex = 0
    If mblnBloodReactionRecordIsOpen = True Then GoTo TOSHOW
    '初始化全局变量
    Set mfrmMain = frmMain
    mstr科室 = cbo1.Text
    mlngtbcIndex = 0
    mlngSys = lngSys
    mlngModule = lngMoudle
    mstrPrivs = strPrivs
    mlng阶段 = lng阶段
    mblnADDPeoPle = False
    mstrFindKey = ""
    mlng提交状态 = 0 '0-全部数据，1-未提交数据，2-已提交数据
    mblnStart = False
    strCurDate = Format(gobjDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    DTP1.Value = Format(CDate(strCurDate) - 29, "YYYY-MM-DD 00:00:00")
    DTP2.Value = Format(strCurDate, "YYYY-MM-DD 23:59:59")
    TimNotify.Enabled = mlng阶段 = 2
    
    If mlng阶段 = 2 Then
        mstr开始时间 = DTP1.Value
        mstr结束时间 = DTP2.Value
        mintNotify = gobjDatabase.GetPara("消息提醒间隔", 2200, 1938, 0)
    Else
        mstr开始时间 = Format(strCurDate, "YYYY-MM-DD")
        mstr结束时间 = Format(strCurDate, "YYYY-MM-DD") & " 23:59:59"
    End If
    
    '初始化commandbar
    InitCommandBar
    '初始化cboTime控件及相关控件
    initComboTime
    
    '初始化dockingpane
    Me.dkpPeoPle.SetCommandBars Me.cbsMain
    Me.dkpPeoPle.Options.UseSplitterTracker = False '实时拖动
    Me.dkpPeoPle.Options.ThemedFloatingFrames = True
    Me.dkpPeoPle.Options.AlphaDockingContext = True
    Me.dkpPeoPle.Options.HideClient = True
    
    Set objPane = dkpPeoPle.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "病人列表": objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Set objPane = dkpPeoPle.CreatePane(2, 100, 30, DockBottomOf, objPane): objPane.Title = "消息提醒": objPane.Options = PaneNoFloatable Or PaneNoHideable Or PaneNoCloseable
    If mlng阶段 <> 2 Then objPane.Options = PaneActionClosed: dkpPeoPle(2).Close
    Set objPane = dkpPeoPle.CreatePane(3, 700, 100, DockRightOf, Nothing): objPane.Title = "记录": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    '初始化消息提示控件
    InitReportColumn
    '初始化部门科室信息
    If mlng阶段 = 2 Then
        Set rs部门 = GetDeptList("血库", 3, IsPrivs(mstrPrivs, "所有科室"))
    Else
        Set rs部门 = GetDeptList("临床", mlng阶段 + 1, IsPrivs(mstrPrivs, "所有科室"))
    End If
    
    If rs部门.RecordCount <= 0 Then
        MsgBox "你不属于" & IIf(mlng阶段 = 2, "血库", "临床") & "部门！", vbInformation, gstrSysName
        Exit Sub
    End If
    If IsPrivs(mstrPrivs, "所有科室") Then
        cbo1.AddItem "所有科室"
        cbo1.ItemData(cbo1.NewIndex) = -1 '所有科室的id默认为-1
    End If
    
    For lngi = 0 To rs部门.RecordCount - 1
        cbo1.AddItem rs部门.Fields("简码") & "-" & rs部门.Fields("名称").Value
        cbo1.ItemData(cbo1.NewIndex) = rs部门.Fields("id").Value
        If rs部门.Fields("id").Value = UserInfo.部门ID Then lngIndex = IIf(IsPrivs(mstrPrivs, "所有科室") = True, lngi + 1, lngi)
        rs部门.MoveNext
    Next
    cbo1.ListIndex = lngIndex
    mintDeptIndex = lngIndex
    
    mArr过滤数据(0) = -1
    mArr过滤数据(1) = mstr开始时间 & "'" & mstr结束时间
    mArr过滤数据(2) = IIf(chk1(0).Value = Checked, UserInfo.姓名, "")
    mArr过滤数据(3) = mlng提交状态
    If opt(0).Value Then mArr过滤数据(4) = 0
    If opt(1).Value Then mArr过滤数据(4) = 1
    If opt(2).Value Then mArr过滤数据(4) = 2
    If opt(3).Value Then mArr过滤数据(4) = 3
    
    '初始化tabControl
    Call initTabControl(mlng阶段)
    
    '初始化usrCardEdit控件
    UCE.InitEdit
    '初始化usrCardPeople
    UCP.UserInit Me, "颜色|ID|1||||255;住院情况|主页ID;床号;姓名;病历号;性别和年龄;入院日期;填写人", , p输血反应管理
    
    '初始化表格和查询病人信息
    Call ExecuteCommand("初始表格")
    Call RefreshBR
        
    If Not mRsBR Is Nothing Then
        If mRsBR.RecordCount > 0 Then
            mblnHaveBR = True
        End If
    Else
        mblnHaveBR = False
    End If
    mblnStart = True
TOSHOW:
    If mblnBloodReactionRecordIsOpen = True Then
        mArr过滤数据(0) = -1
        mArr过滤数据(1) = mstr开始时间 & "'" & mstr结束时间
        mArr过滤数据(2) = UserInfo.姓名
        mArr过滤数据(3) = mlng提交状态
        If opt(0).Value Then mArr过滤数据(4) = 0
        If opt(1).Value Then mArr过滤数据(4) = 1
        If opt(2).Value Then mArr过滤数据(4) = 2
        If opt(3).Value Then mArr过滤数据(4) = 3
    End If
    mblnBloodReactionRecordIsOpen = True
    If IsObject(mfrmMain) Then
        If frmMain Is Nothing Then
            Me.Show lngisModul
        Else
            Me.Show lngisModul, mfrmMain
        End If
    Else
        gobjComlib.os.ShowChildWindow Me.hWnd, Val(mfrmMain)
    End If
    
End Sub

Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
    '******************************************************************************************************************
    '功能：初始化DockPannel
    '参数：
    '返回：
    '******************************************************************************************************************
    
    
End Function

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    Dim intLoop As Integer
    Dim lngi As Long
    Dim lngj As Long
    Dim rsSAD As New ADODB.Recordset
    Dim Arr部门
    Dim lng病人ID As Long, lng主页id As Long
    On Error GoTo Error
    
    Call SQLRecord(rsSAD)
    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
            Case "初始表格"
                Set mclsVsf = New clsVsf
                With mclsVsf
                    Call .Initialize(Me.Controls, VSFBRlist, True, True)
                    Call .ClearColumn
                    Call .AppendColumn("病人id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("主页id", 0, flexAlignRightCenter, flexDTString, "", , True, , , True)
                    Call .AppendColumn("姓名", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("性别和年龄", 1000, flexAlignLeftCenter, flexDTString, , "", True)
                    Call .AppendColumn("床号", 700, flexAlignLeftCenter, flexDTString, , "", True)
                    Call .AppendColumn("病历号", 1400, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("住院情况", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("入院日期", 1100, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("科室名称", 900, flexAlignLeftCenter, flexDTString, "", , True)
                    Call .AppendColumn("填写人", 1000, flexAlignLeftCenter, flexDTString, "", , True)

                    .AppendRows = False
                    .SysHidden(.ColIndex("病人id")) = True
                    .SysHidden(.ColIndex("主页id")) = True
                    Call .InitializeEdit(True, True, True)
                    Call .InitializeEditColumn(.ColIndex(""), True, vbVsfEditCheck)
                    
                End With
                
            Case "基础病人查询":
                '功能：根据查询语句查询数据库中符合条件的数据并放入指定的集合中，
                Dim strSQL As String
                Dim strSql1 As String
                Dim strSqlRection As String
                Dim lng提交状态 As String
                Dim str本人提交 As String
                Dim lngColor As Long
                Dim rsBR As ADODB.Recordset
                Dim lng科室ID As Long
                
                lng科室ID = Val(cbo1.ItemData(cbo1.ListIndex))
                
                If chk1(1).Value = Checked And chk1(2).Value = Unchecked Then  '已提交
                    mlng提交状态 = 2
                ElseIf chk1(1).Value = Unchecked And chk1(2).Value = Checked Then  '未提交
                    mlng提交状态 = 1
                Else '全部数据
                    mlng提交状态 = 0
                End If

                Select Case mlng阶段
                    Case 0: '门诊病人   去掉语句"and f.执行人=[4]" ,添加了对输血反应记录的记录人的判断，
                        strSQL = " Select f.病人ID || '-' || f.id As id, f.病人id, f.Id As 主页id, f.门诊号 As 病历号, '门' As 住院情况, f.姓名, f.性别 || '/' || f.年龄 As 性别和年龄, f.执行部门id As 科室id, " & vbNewLine & _
                                 "        b.名称 As 科室名称, f.执行时间 As 入院日期, '' As 床号, f.险类, '' As 类型, 255 As 颜色, f.执行人 As 填写人 " & vbNewLine & _
                                 " From 病人挂号记录 f, 部门表 b " & vbNewLine & _
                                 " Where f.执行部门id = b.Id And f.执行状态 = [5] " & vbNewLine & _
                                 " " & IIf(lng科室ID <> -1, " and b.id=[3] ", "") & IIf(mlngtbcIndex = 1, " and f.执行时间 Between [1] And [2] ", "")
                        
                        If ChkRection.Value = 1 Then
                            strSqlRection = " And Exists (Select 1 " & vbNewLine & _
                                 "        From 血液配血记录 e, 血液收发记录 g, 病人医嘱记录 h,输血反应记录 J " & vbNewLine & _
                                 "        Where h.Id = e.申请id And e.Id = g.配发id  and g.id=j.收发id " & IIf(chk1(0).Value = Checked, " and j.记录人=[4] ", "") & " " & IIf(mlng提交状态 = 1, " and J.状态=0 ", IIf(mlng提交状态 = 2, " and J.状态<>0 ", "")) & " And " & vbNewLine & _
                                 "              Mod(g.记录状态, 3) = 1 And g.审核人 Is Not Null And H.诊疗类别='K' And h.挂号单 = f.No And  h.病人id = f.病人id)"
                        Else
                            strSqlRection = " And Exists (Select 1 " & vbNewLine & _
                                 "        From 血液配血记录 e, 血液收发记录 g, 病人医嘱记录 h " & vbNewLine & _
                                 "        Where h.Id = e.申请id And e.Id = g.配发id And h.病人id = f.病人id And " & vbNewLine & _
                                 "              h.挂号单 = f.No And Mod(g.记录状态, 3) = 1 And g.审核人 Is Not Null)"
                        End If
                        
                        strSQL = strSQL & strSqlRection
                        
                        If mlngtbcIndex = 0 Then '正在就诊
                            Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", CDate(mstr开始时间), CDate(mstr结束时间), lng科室ID, UserInfo.姓名, 2)
                        ElseIf mlngtbcIndex = 1 Then '完成就诊
                            Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", CDate(mstr开始时间), CDate(mstr结束时间), lng科室ID, UserInfo.姓名, 1)
                        End If
                        
                    Case 1: '住院病人
                        If mlngtbcIndex = 0 Then '1表示在院
                            strSQL = " Select a.病人ID || '-' || a.主页id As id,a.病人id, a.主页id, a.住院号 As 病历号, '住' As 住院情况, a.姓名, a.性别 || '/' || a.年龄 As 性别和年龄," & vbNewLine & _
                                     " f.科室id As 科室id, b.名称 As 科室名称, a.入院日期, a.入院病床 As 床号, a.险类, a.病人类型 As 类型, 255 As 颜色, a.登记人 As 填写人 " & vbNewLine & _
                                     " From 病案主页 a,病人信息 C, 部门表 b, 在院病人 f " & vbNewLine & _
                                     " Where b.id = f.科室id And a.病人id = c.病人id And a.主页ID = c.主页ID And c.病人id = f.病人id " & IIf(lng科室ID <> -1, " and f.科室id=[3] ", "") & ""
                        ElseIf mlngtbcIndex = 2 Then '2表示出院
                            strSQL = " Select a.病人ID || '-' || a.主页id As id,a.病人id, a.主页id, a.住院号 As 病历号, '住' As 住院情况, a.姓名, a.性别 || '/' || a.年龄 As 性别和年龄," & vbNewLine & _
                                     "        a.出院科室id As 科室id, b.名称 As 科室名称, a.入院日期, a.入院病床 As 床号, a.险类, a.病人类型 As 类型, 255 As 颜色, a.登记人 As 填写人 " & vbNewLine & _
                                     " From 病案主页 a, 病人信息 c, 部门表 b " & vbNewLine & _
                                     " Where a.病人id = c.病人id And a.主页id = c.主页id And a.出院科室id = b.Id " & IIf(lng科室ID <> -1, " and a.出院科室id=[3] ", "") & " And a.出院日期 Between [1] And [2] "
                        Else '1表示转出
                            strSQL = " Select a.病人ID || '-' || a.主页id As id,a.病人id, a.主页id, a.住院号  As 病历号, '住' As 住院情况, a.姓名, a.性别 || '/' || a.年龄 As 性别和年龄, " & vbNewLine & _
                                     "       f.科室id As 科室id, b.名称 As 科室名称, a.入院日期, a.入院病床 As 床号, a.险类, a.病人类型 As 类型, 255 As 颜色, a.登记人 As 填写人 " & vbNewLine & _
                                     " From 病案主页 a, 病人信息 c, 部门表 b, 病人变动记录 f  " & vbNewLine & _
                                     " Where a.病人id = c.病人id And a.主页id = c.主页id And f.病人id = a.病人id And f.主页id = a.主页id And b.Id = f.科室id And f.开始原因 = 3 And " & vbNewLine & _
                                     "       Nvl(f.附加床位, 0) = 0 " & IIf(lng科室ID <> -1, " and f.科室id=[3] ", "") & " And f.开始时间 Between Sysdate - [5] And Sysdate  "
                        End If
                        
                        If ChkRection.Value = 1 Then
                            strSqlRection = " And Exists " & vbNewLine & _
                                     " (Select 1 From 血液配血记录 e, 血液收发记录 g,输血反应记录 h " & vbNewLine & _
                                     "        Where e.Id = g.配发id And e.病人id = a.病人id And e.主页id = a.主页id and g.id=h.收发id  " & IIf(chk1(0).Value = Checked, " and h.记录人=[4] ", "") & " " & IIf(mlng提交状态 = 1, " and h.状态=0 ", IIf(mlng提交状态 = 2, " and h.状态<>0 ", "")) & " And  Mod(g.记录状态, 3) = 1 And g.审核人 Is Not Null)"
                        Else
                            strSqlRection = " And Exists " & vbNewLine & _
                                     " (Select 1 From 血液配血记录 e, 血液收发记录 g " & vbNewLine & _
                                     "        Where e.Id = g.配发id And e.病人id = a.病人id And e.主页id = a.主页id And  Mod(g.记录状态, 3) = 1 And g.审核人 Is Not Null)"
                        End If
                        
                        strSQL = strSQL & strSqlRection
                        
                        Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", CDate(mstr开始时间), CDate(mstr结束时间), lng科室ID, UserInfo.姓名, Val(TXTDay.Text))
                    Case 2: '输血科
                        '确定提交状态的过滤条件
                        If mlng提交状态 = 0 Then '全部数据
                            strSql1 = " and (e.状态<>0 or e.是否输血科新增 =1) "
                        ElseIf mlng提交状态 = 1 Then '未提交数据
                            strSql1 = " and (e.状态<>2 and e.是否输血科新增=1 OR e.状态=1)"
                        Else '已提交数据
                            strSql1 = " and e.状态=2 "
                        End If
                        If mArr过滤数据(4) = 0 Then
                            strSql1 = strSql1 & " and e.有无输血反应 = 2 "
                        ElseIf mArr过滤数据(4) = 1 Then
                            strSql1 = strSql1 & " and e.有无输血反应 = 1 "
                        ElseIf mArr过滤数据(4) = 2 Then
                        ElseIf mArr过滤数据(4) = 3 Then
                            strSql1 = strSql1 & " and e.有无输血反应 = 0 "
                        End If
                        '去掉语句“and a.登记人=[5]” 和  and f.执行人=[5]，新增了对输血科确认人的判断
                        strSQL = " Select a.病人ID || '-' || a.主页id As id,a.病人id, a.主页id, a.住院号 As 病历号, '住' As 住院情况, a.姓名, a.性别 || '/' || a.年龄 As 性别和年龄, " & vbNewLine & _
                                 "        a.出院科室id As 科室id, b.名称 As 科室名称, a.入院日期, a.入院病床 As 床号, a.险类, a.病人类型 As 类型, 255 As 颜色, a.登记人 As 填写人 " & vbNewLine & _
                                 " From 病案主页 a, 部门表 b, " & vbNewLine & _
                                 "      (Select Distinct c.病人id, c.主页id " & vbNewLine & _
                                 "       From 血液配血记录 c, 血液收发记录 d, 输血反应记录 e " & vbNewLine & _
                                 "       Where c.Id = d.配发id And d.Id = e.收发id " & IIf(chk1(0).Value = Checked, " And (e.确认人=[5] or e.是否输血科新增 =1 And e.记录人=[5]) ", "") & " And e.反应时间 Between [1] And [2] " & IIf(lng科室ID = -1, "", " And c.执行部门id = [3] ") & vbNewLine & _
                                 " And c.记录性质 = 1 " & strSql1 & "  And Mod(d.记录状态, 3) = 1 And d.审核人 Is Not Null) K " & vbNewLine & _
                                 "Where K.病人id = a.病人id And K.主页id = a.主页id And a.出院科室id = b.Id(+) "

                        strSQL = strSQL & " Union ALL" & vbNewLine & _
                                " Select f.病人ID || '-' || f.id As id,f.病人id, f.Id As 主页id, f.门诊号 As 病历号, '门' As 住院情况, f.姓名, f.性别 || '/' || f.年龄 As 性别和年龄, f.执行部门id As 科室id, " & vbNewLine & _
                                "        b.名称 As 科室名称, f.执行时间 As 入院日期, '' As 床号, f.险类, '' As 类型, 255 As 颜色, f.执行人 As 填写人 " & vbNewLine & _
                                " From 病人挂号记录 f, 部门表 b, " & vbNewLine & _
                                "      (Select Distinct g.病人id, g.挂号单 " & vbNewLine & _
                                "       From 血液配血记录 c, 血液收发记录 d, 输血反应记录 e, 病人医嘱记录 g " & vbNewLine & _
                                "       Where c.Id = d.配发id And d.Id = e.收发id " & IIf(chk1(0).Value = Checked, " And (e.确认人=[5] or e.是否输血科新增 =1 And e.记录人=[5]) ", "") & " And g.Id = c.申请id And c.记录性质 = 1 " & strSql1 & " And Mod(d.记录状态, 3) = 1 And " & vbNewLine & _
                                "             d.审核人 Is Not Null And e.反应时间 Between [1] And [2] " & IIf(lng科室ID = -1, "", " And c.执行部门id = [3] ") & " And g.诊疗类别 = 'K') h " & vbNewLine & _
                                " Where h.病人id = f.病人id And h.挂号单 = f.No And f.执行部门id = b.Id "
                        
                        '加上查询到的病人
                        If mblnADDPeoPle And mstrFindKey <> "" Then
                            Arr部门 = Split(mstrFindKey, "-")  '病人ID-就诊ID-门诊或住院
                            lng病人ID = Val(Arr部门(0))
                            lng主页id = Val(Arr部门(1))
                            If Val(Arr部门(2)) = 0 Then '住院
                                strSQL = strSQL & " Union ALL" & vbNewLine & _
                                    " Select a.病人id || '-' || a.主页id As Id, a.病人id, a.主页id, a.住院号  As 病历号, '住' As 住院情况, a.姓名," & vbNewLine & _
                                    "       a.性别 || '/' || a.年龄 As 性别和年龄, a.出院科室id As 科室id, d.名称 As 科室名称, a.入院日期, a.入院病床 As 床号, a.险类, a.病人类型 As 类型, 255 As 颜色," & vbNewLine & _
                                    "       a.登记人 As 填写人" & vbNewLine & _
                                    " From 部门表 d, 病案主页 a" & vbNewLine & _
                                    " Where a.出院科室id = d.Id(+) And a.病人id = [6] And a.主页id = [7] And Exists" & vbNewLine & _
                                    "  (Select 1" & vbNewLine & _
                                    "       From 血液收发记录 c, 血液配血记录 b" & vbNewLine & _
                                    "       Where Mod(c.记录状态, 3) = 1 And c.审核人 Is Not Null And b.Id = c.配发id And b.病人id = a.病人id And b.主页id = a.主页id)"

                            Else '门诊
                                strSQL = strSQL & " Union ALL" & vbNewLine & _
                                    " Select a.病人id || '-' || a.Id As Id, a.病人id, a.Id As 主页id, a.门诊号 As 病历号, '门' As 住院情况, a.姓名," & vbNewLine & _
                                    "       a.性别 || '/' || a.年龄 As 性别和年龄, a.执行部门id As 科室id, d.名称 As 科室名称, a.执行时间 As 入院日期, '' As 床号, a.险类, '' As 类型, 255 As 颜色," & vbNewLine & _
                                    "       a.执行人 As 填写人" & vbNewLine & _
                                    " From 部门表 d, 病人挂号记录 a" & vbNewLine & _
                                    " Where a.执行部门id = d.Id And a.病人id = [6] And a.Id = [7] And Exists" & vbNewLine & _
                                    "  (Select 1" & vbNewLine & _
                                    "       From 血液收发记录 c, 血液配血记录 b, 病人医嘱记录 e" & vbNewLine & _
                                    "       Where Mod(c.记录状态, 3) = 1 And c.审核人 Is Not Null And b.Id = c.配发id And b.申请id = e.Id And e.诊疗类别 = 'K' And" & vbNewLine & _
                                    "             e.病人id = a.病人id And e.挂号单 = a.No)"
                            End If
                        End If
                        Set rsBR = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", CDate(mstr开始时间), CDate(mstr结束时间), lng科室ID, mlng提交状态, UserInfo.姓名, lng病人ID, lng主页id)
                End Select
                
                Call RsTitelCopy(rsBR, mRsBR)
                
                With mRsBR
                    If rsBR.RecordCount > 0 Then '以前没有对rsbr的数据做判断会报错
                        For lngi = 0 To rsBR.RecordCount - 1
                            .AddNew
                            For lngj = 0 To rsBR.Fields.Count - 1
                                .Fields(lngj).Value = rsBR.Fields(lngj).Value
                                
                                If .Fields(lngj).name = "入院日期" Then '对日期进行模式化处理，要不然日期显示会有问题
                                    .Fields(lngj).Value = Format(rsBR.Fields("入院日期").Value, "YYYY-MM-DD HH:mm:ss")
                                End If
                                
                                If .Fields(lngj).name = "颜色" Then '重新根据类型和险类分配颜色
                                    If Not IsNull(rsBR!险类) And Len(rsBR!类型) > 0 Then
                                        '病人颜色
                                        lngColor = gobjDatabase.GetPatiColor(Nvl(rsBR!类型))
                                        .Fields("颜色").Value = lngColor
                                    End If
                                End If
                            Next
                            .Update
                            rsBR.MoveNext
                        Next
                        rsBR.MoveFirst
                        If .RecordCount > 0 Then
                            .MoveFirst
                        End If
                    End If
                End With
                
            Case "刷新数据"
                Dim rsTemp As ADODB.Recordset
                Dim StrPosition As String
                mblnFirst = True
                mlng提交状态 = 0
                mblnIsSubmit = True
                
                If InStr(1, cbo1.Text, "-") > 0 Then
                    Arr部门 = Split(cbo1.Text, "-")
                    mstr科室 = Arr部门(1)
                Else
                    mstr科室 = cbo1.Text
                End If

                Call ExecuteCommand("基础病人查询")

                mArr过滤数据(3) = mlng提交状态 '提交状态
'                Set rsTemp = mRsBR
                Call CopyRecord(mRsBR, rsTemp)
                
                If rsTemp.RecordCount > 0 Then
                    mblnHaveBR = True
'                    rsTemp.MoveFirst
                Else
                    mblnHaveBR = False
                End If
                UCP.ShowPeople rsTemp, True
                Call mclsVsf.LoadGrid(rsTemp) '将数据放入隐藏的打印列表中。
                
                Set rsTemp = Nothing

                StrPosition = mArrPosition(mlngtbcIndex)
                UCP.SetCardFocus "病人id'主页id", StrPosition
            Case "刷新提示"
                If mlng阶段 <> 2 Then Exit Function
                With rptTips
                    Set rsSAD = GetReactionTips(Val(cbo1.ItemData(cbo1.ListIndex)))
                    .Records.DeleteAll
                    .Populate
                    If rsSAD.RecordCount <> 0 Then
                        rsSAD.MoveFirst
                        Call LoadRptTips(rsSAD)
                    End If
                End With
            Case "本地参数设置"
                ExecuteCommand = frmBloodReactionRecordSetup.ShowPara(Me)
                mintNotify = gobjDatabase.GetPara("消息提醒间隔", 2200, 1938, 0)
        End Select
    Next
    ExecuteCommand = True
    Exit Function
Error:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    
    ExecuteCommand = False
End Function

Private Sub LoadRptTips(rsData As Recordset)
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strSQL As String, rsTmp As New Recordset
    Dim strTmp As String
    rptTips.Records.DeleteAll
    strTmp = ""
    If rsData Is Nothing Then Exit Sub
    If rsData.RecordCount = 0 Then Exit Sub
    rsData.MoveFirst
    Do While Not rsData.EOF
        If InStr(1, strTmp & ",", "," & Val(Nvl(rsData!病人id, 0))) & "," = 0 Then strTmp = strTmp & "," & Val(Nvl(rsData!病人id, 0))
        rsData.MoveNext
    Loop
    If strTmp = "" Then Exit Sub
    strSQL = "select /*+ CARDINALITY(b,10) */ a.姓名,a.病人id from 病人信息 a,table(f_str2list([1],',')) b where a.病人id = b.column_value"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "病人信息", strTmp)
    strTmp = ""
    rsData.MoveFirst
    Do While Not rsData.EOF
        If InStr(1, strTmp & "|", "|" & Val(Nvl(rsData!病人id, 0)) & "|") = 0 Then
            Set objRecord = Me.rptTips.Records.Add()
            rsTmp.Filter = "病人id = " & Val(Nvl(rsData!病人id, 0))
            Set objItem = objRecord.AddItem(CStr(Nvl(rsTmp!姓名)))
            Set objItem = objRecord.AddItem(CStr(rsData!消息内容))
            Set objItem = objRecord.AddItem(Val(Nvl(rsData!收发ID, 0)))
            Set objItem = objRecord.AddItem(Val(Nvl(rsData!病人id, 0)))
            Set objItem = objRecord.AddItem(Val(Nvl(rsData!就诊id, 0)))
            Set objItem = objRecord.AddItem(Val(Nvl(rsData!病人来源, 0)))
            strTmp = strTmp & "|" & Val(Nvl(rsData!病人id, 0))
        End If
        rsData.MoveNext
    Loop
    rptTips.Populate
End Sub
Private Sub RsTitelCopy(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '功能：新建ToRs记录集，将RsProm的结构复制到ToRs上
    '参数：RsProm-原记录集，ToRs-新建的记录集
    Dim lngi As Long
    Set ToRs = New ADODB.Recordset
    With ToRs '初始化rsReturn
        For lngi = 0 To RsProm.Fields.Count - 1
            .Fields.Append RsProm.Fields(lngi).name, adLongVarChar, 100, adFldIsNullable
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub CopyRecord(ByVal RsProm As ADODB.Recordset, ToRs As ADODB.Recordset)
    '功能：将记录集RsProm的结构还有数据都复制给ToRs
    '参数：RsProm-要赋值的记录集，ToRs-目标记录集
    Dim lngi As Long
    Dim lngj As Long
    Call RsTitelCopy(RsProm, ToRs)
    With ToRs
        If RsProm.RecordCount > 0 Then '以前没有对rsbr的数据做判断会报错
            For lngi = 0 To RsProm.RecordCount - 1
                .AddNew
                For lngj = 0 To RsProm.Fields.Count - 1
                    .Fields(lngj).Value = RsProm.Fields(lngj).Value
                Next
                .Update
                RsProm.MoveNext
            Next
            RsProm.MoveFirst
            If .RecordCount > 0 Then
                .MoveFirst
            End If
        End If
    End With
End Sub

Private Sub chk1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub ChkRection_Click()
    Call pic1_Resize
End Sub

Private Sub ChkRection_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub cmd2_Click()
    RefreshBR
    pic1_Resize
End Sub

Private Sub dkpPeoPle_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
        Case 1
            Item.Handle = pic1.hWnd
        Case 2
            Item.Handle = picTips.hWnd
        Case 3
            Item.Handle = pic2.hWnd
    End Select
End Sub

Private Sub DTP1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub DTP2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub DTP3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub DTP4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call SetPaneRange(dkpPeoPle, 1, 260, 100, 320, Me.ScaleHeight)
    Call SetPaneRange(dkpPeoPle, 2, 260, 40, 320, 100)
    Call SetPaneRange(dkpPeoPle, 3, 100, 100, Me.ScaleWidth, Me.ScaleHeight)
    dkpPeoPle.RecalcLayout
End Sub

Public Function SetPaneRange(dkpM As Object, ByVal intPane As Integer, ByVal lngMinW As Long, lngMinH As Long, lngMaxW As Long, lngMaxH As Long) As Boolean
    '******************************************************************************************************************
    '功能：设置dockingpane的大小范围
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objPan As Pane
    
    On Error Resume Next
    
    Set objPan = dkpM.FindPane(intPane)
    
    If objPan Is Nothing Then Exit Function
    With objPan
        .MaxTrackSize.SetSize lngMaxW, lngMaxH
        .MinTrackSize.SetSize lngMinW, lngMinH
    End With
    
    SetPaneRange = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If UCE.strST = 新增 Or UCE.strST = 修改 Then
        Cancel = (MsgBox("数据未保存，是否不保存就退出？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    mblnBloodReactionRecordIsOpen = False
    mblnStart = False
End Sub

Private Sub initTabControl(Index As Long)
    '功能：初始化tbcthis
    With tbcthis
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .COLOR = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = False
        End With
        
        Select Case Index
            Case 0:
                .InsertItem(0, "正在就诊", picTmp.hWnd, 0).Tag = "正在就诊"
                .InsertItem(1, "完成就诊", picTmp.hWnd, 0).Tag = "完成就诊"
                .Item(0).Selected = True
            Case 1:
                .InsertItem(0, "在院", picTmp.hWnd, 0).Tag = "在院"
                .InsertItem(1, "转出", picTmp.hWnd, 0).Tag = "转出"
                .InsertItem(2, "出院", picTmp.hWnd, 0).Tag = "出院"
                .Item(0).Selected = True
        End Select
    End With
End Sub

Private Sub opt_Click(Index As Integer)
    opt(2).Tag = Index
End Sub
Private Sub picTips_Resize()
    rptTips.Move 0, 0, picTips.Width, picTips.Height
End Sub
Private Sub picUCP_Resize()
    '功能：调整页面的布局
    On Error Resume Next
    UCP.Move 0, 0, picUCP.ScaleWidth, picUCP.ScaleHeight
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '功能：tbcthis控件切换选项卡后刷新数据
    Dim ArrReturn
    ArrReturn = Split(UCP.strReturn, "'")
    If UBound(ArrReturn) >= 0 Then
        mArrPosition(mlngtbcIndex) = Val(ArrReturn(1)) & "'" & Val(ArrReturn(3)) '记录定位信息，以"病人id'主页id"的形式，ArrReturn(1)和ArrReturn(3)分别代表病人id和主页id
    End If
    mlngtbcIndex = Item.Index
    pic1_Resize
    If mblnStart = True Then
        Call RefreshBR
    End If
End Sub

Private Sub initComboTime()
    '功能：对cobtime和相关dtp控件进行初始化
    cbotime.Clear '出院
    With cbotime
        .AddItem "今天内"
        .ItemData(.NewIndex) = 0
        .AddItem "昨天内"
        .ItemData(.NewIndex) = 1
        .AddItem "前天内"
        .ItemData(.NewIndex) = 2
        .AddItem "一周内"
        .ItemData(.NewIndex) = 6
        .AddItem "30天内"
        .ItemData(.NewIndex) = 29
        .AddItem "60天内"
        .ItemData(.NewIndex) = 59
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
    End With
    If cbotime.ListCount > 0 Then cbotime.ListIndex = 0
End Sub

Private Sub TXTDay_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) = True Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then
    Else
        KeyAscii = 0
    End If
    If KeyAscii = vbKeyReturn Then gobjCommFun.PressKey vbKeyTab
End Sub

Private Sub UCP_CardChanged()
    '功能：usrCardPeople控件切换病人后同时刷新该病人的反应记录
    Dim ArrReturn
    Dim lngi As Long
    Dim lng阶段 As Long
    Dim rsData As ADODB.Recordset
    Dim strFilter As String
    Dim blnHaveRection As Boolean
    
    On Error GoTo Errorhand:

    Set rsData = mRsBR
    '将过滤数据整合成字符串，方便传递
    strFilter = mArr过滤数据(0) & "|" & mArr过滤数据(1) & "|" & mArr过滤数据(2) & "|" & mArr过滤数据(3) & "|" & mArr过滤数据(4)
    
    lng阶段 = mlng阶段
    If lng阶段 = 0 Then '把门诊阶段和住院阶段统一为医生阶段，代表数字为1
        lng阶段 = 1
    End If
    
    ArrReturn = Split(UCP.strReturn, "'")
    If UBound(ArrReturn) >= 0 Then
        mArrPosition(mlngtbcIndex) = Val(ArrReturn(1)) & "'" & Val(ArrReturn(3)) '记录定位信息，以"病人id'主页id"的形式，ArrReturn(1)和ArrReturn(3)分别代表病人id和主页id
    End If
    '加载输血反应记录
    If UBound(ArrReturn) = -1 Then
        UCE.ShowClear
        mblnIsSelect = False
    ElseIf rsData.RecordCount > 0 Then
        With rsData
            .MoveFirst
            For lngi = 0 To .RecordCount - 1
                If .Fields("病人ID").Value = Val(ArrReturn(1)) And .Fields("主页id").Value = Val(ArrReturn(3)) Then '住院病人主页id就是主页id，门诊病人的话，主页id代表就诊号
                    If .Fields("住院情况").Value = "住" Then
                        UCE.showInfor Val(.Fields("病人ID").Value), 2, Val(IIf(IsNull(.Fields("主页id").Value) = True, 0, .Fields("主页id").Value)), lng阶段, gcnOracle, Me, p输血反应管理, strFilter, IsPrivs(mstrPrivs, "输血科新增")
                    Else  '门诊病人使用门诊号替代主页id来区分多次挂号的病人
                        UCE.showInfor Val(.Fields("病人ID").Value), 1, Val(IIf(IsNull(.Fields("主页id").Value) = True, 0, .Fields("主页id").Value)), lng阶段, gcnOracle, Me, p输血反应管理, strFilter, IsPrivs(mstrPrivs, "输血科新增")
                    End If
                End If
                .MoveNext
            Next
            .MoveFirst
        End With
        mblnIsSelect = True
    End If
    Set rsData = Nothing
Errorhand:
End Sub

Private Sub pic1_Resize()
    Dim intType As Integer
    On Error Resume Next
    

    lbl1.Left = 120
    cbo1.Left = lbl1.Left + lbl1.Width + 90
    cbo1.Top = 60
    lbl1.Top = cbo1.Top + (cbo1.Height - lbl1.Height) \ 2
    
    Fra1.Left = 60
    Fra1.Top = cbo1.Top + cbo1.Height + 60
    Fra1.Width = pic1.ScaleWidth - 60
    
    cmd2.Visible = True
    If mlng阶段 = 2 Then '输血科
        '显示提示
        picTips.Visible = True
        '不显示ChkRection控件
        ChkRection.Visible = False
        '显示输血反应过滤条件控件，并调整位置
        chk1(0).Visible = True
        chk1(1).Visible = True
        chk1(2).Visible = True
        tbcthis.Visible = False
        
        '输血反应时间
        lbl2.Left = 120
        DTP1.Left = lbl2.Left + lbl2.Width + 90
        DTP1.Top = 210
        lbl2.Top = DTP1.Top + (DTP1.Height - lbl2.Height) \ 2
        DTP2.Left = DTP1.Left
        DTP2.Top = DTP1.Top + DTP1.Height + 60
        lbl4.Left = lbl2.Left + lbl2.Width - lbl4.Width
        lbl4.Top = DTP2.Top + (DTP2.Height - lbl4.Height) \ 2
        
        chk1(0).Left = DTP1.Left
        chk1(0).Top = DTP2.Top + DTP2.Height + 60
        chk1(1).Top = chk1(0).Top
        chk1(2).Top = chk1(0).Top
        
        chk1(1).Value = Checked
        chk1(2).Value = Checked
        
        fra.Top = chk1(0).Top + chk1(0).Height + 60
        cmd2.Top = fra.Top + fra.Height + 60
        
        Fra1.Height = cmd2.Top + cmd2.Height + 120
        
        picUCP.Left = Fra1.Left
        picUCP.Top = Fra1.Top + Fra1.Height + 60
        picUCP.Width = Fra1.Width
        If pic1.ScaleHeight - Fra1.Top - Fra1.Height > 0 Then
            picUCP.Height = pic1.ScaleHeight - Fra1.Top - Fra1.Height
        End If
        '隐藏其他无用的条件
        lbl5.Visible = False
        cbotime.Visible = False
        lbl8.Visible = False
        DTP3.Visible = False
        lbl9.Visible = False
        DTP4.Visible = False
        lbl6.Visible = False
        TXTDay.Visible = False
        frmLine.Visible = False
    Else '住院或者门诊
        '不显示提示
        picTips.Visible = False
        ChkRection.Visible = True
        If (mlng阶段 = 1 Or mlng阶段 = 0) And mlngtbcIndex = 0 Then '在院或者正在就诊
            '控件显示设置
            lbl5.Visible = False
            cbotime.Visible = False
            lbl8.Visible = False
            DTP3.Visible = False
            lbl9.Visible = False
            DTP4.Visible = False
            lbl6.Visible = False
            TXTDay.Visible = False
            frmLine.Visible = False
            ChkRection.Left = 120
            ChkRection.Top = 240
            intType = 0
        ElseIf (mlng阶段 = 1 And mlngtbcIndex = 2) Or (mlng阶段 = 0 And mlngtbcIndex = 1) Then  '出院或者完成就诊
            '控件显示设置
            lbl5.Visible = True
            cbotime.Visible = True
            lbl8.Visible = True
            DTP3.Visible = True
            lbl9.Visible = True
            DTP4.Visible = True
            lbl6.Visible = False
            TXTDay.Visible = False
            frmLine.Visible = False
            
            lbl5.Left = 120
            cbotime.Left = lbl5.Left + lbl5.Width + 90
            cbotime.Top = 210
            
            lbl5.Top = cbotime.Top + (cbotime.Height - lbl5.Height) \ 2
            DTP3.Left = cbotime.Left
            DTP3.Top = cbotime.Top + cbotime.Height + 60
            lbl8.Left = lbl5.Left
            lbl8.Top = DTP3.Top + (DTP3.Height - lbl8.Height) \ 2
            DTP4.Left = DTP3.Left
            DTP4.Top = DTP3.Top + DTP3.Height + 60
            lbl9.Left = lbl5.Left
            lbl9.Top = DTP4.Top + (DTP4.Height - lbl9.Height) \ 2
            ChkRection.Left = 120
            ChkRection.Top = DTP4.Top + DTP4.Height + 60
            
            If mlng阶段 = 1 Then
                lbl5.Caption = "出院日期"
            ElseIf mlng阶段 = 0 Then
                lbl5.Caption = "结诊日期"
            End If
            intType = 1
        ElseIf mlng阶段 = 1 And mlngtbcIndex = 1 Then '转出
            '控件显示设置
            lbl5.Visible = False
            cbotime.Visible = False
            lbl8.Visible = False
            DTP3.Visible = False
            lbl9.Visible = False
            DTP4.Visible = False
            lbl6.Visible = True
            TXTDay.Visible = True
            frmLine.Visible = True
            
            lbl6.Left = 120
            lbl6.Top = 240
            TXTDay.Left = lbl6.Left + 810
            TXTDay.Top = lbl6.Top
            frmLine.Left = TXTDay.Left
            frmLine.Top = TXTDay.Top + TXTDay.Height + 15
            ChkRection.Left = 120
            ChkRection.Top = lbl6.Top + lbl6.Height + 120
            intType = 2
        End If
        If ChkRection.Value = 0 Then
            chk1(0).Visible = False
            chk1(1).Visible = False
            chk1(2).Visible = False
            Select Case intType
                Case 0
                    cmd2.Visible = False
                    Fra1.Height = ChkRection.Top + ChkRection.Height + 120
                Case Else
                    cmd2.Top = ChkRection.Top + ChkRection.Height + 60
                    Fra1.Height = cmd2.Top + cmd2.Height + 120
            End Select
        Else
            chk1(0).Visible = True
            chk1(1).Visible = True
            chk1(2).Visible = True
            
            chk1(0).Left = ChkRection.Left + 180
            chk1(0).Top = ChkRection.Top + ChkRection.Height + 60
            chk1(1).Left = chk1(0).Left + chk1(0).Width
            chk1(1).Top = chk1(0).Top
            chk1(2).Left = chk1(1).Left + chk1(1).Width
            chk1(2).Top = chk1(0).Top
            
            chk1(1).Value = Checked
            chk1(2).Value = Checked
            
            cmd2.Top = chk1(0).Top + chk1(0).Height + 60
            Fra1.Height = cmd2.Top + cmd2.Height + 120
        End If
        
        lbl2.Visible = False
        lbl4.Visible = False
        DTP1.Visible = False
        DTP2.Visible = False
        
        tbcthis.Left = Fra1.Left
        tbcthis.Top = Fra1.Top + Fra1.Height + 60
        tbcthis.Width = Fra1.Width
        tbcthis.Height = 350
        
        picUCP.Left = Fra1.Left
        picUCP.Top = tbcthis.Top + tbcthis.Height
        picUCP.Width = Fra1.Width
        If pic1.ScaleHeight - tbcthis.Top - tbcthis.Height > 0 Then
            picUCP.Height = pic1.ScaleHeight - tbcthis.Top - tbcthis.Height
        End If
    End If
End Sub

Private Sub pic2_Resize()
    On Error Resume Next
    UCE.Move 0, 0, pic2.Width, pic2.ScaleHeight
End Sub
