VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#1.0#0"; "zlIDKind.ocx"
Begin VB.Form frmPACStation 
   Caption         =   "影像医技工作站"
   ClientHeight    =   7305
   ClientLeft      =   10185
   ClientTop       =   345
   ClientWidth     =   11325
   Icon            =   "frmPACStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLocate 
      Height          =   300
      Left            =   4920
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox PicWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   1725
      ScaleHeight     =   3735
      ScaleWidth      =   9510
      TabIndex        =   1
      Top             =   2670
      Width           =   9510
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   625
         Left            =   0
         ScaleHeight     =   630
         ScaleWidth      =   9465
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   15
         Width           =   9465
         Begin VB.Frame fraRegist 
            Height          =   700
            Left            =   0
            TabIndex        =   7
            Top             =   -75
            Width           =   1980
            Begin VB.ComboBox cboTimes 
               Height          =   300
               Left            =   60
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   325
               Width           =   1875
            End
            Begin VB.Label lblRegist 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "检查记录(&G)"
               Height          =   180
               Left            =   105
               TabIndex        =   9
               Top             =   150
               Width           =   990
            End
         End
         Begin VB.Frame fraInfo 
            Height          =   700
            Left            =   1980
            TabIndex        =   4
            Top             =   -75
            Width           =   7410
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
               Height          =   540
               Left            =   6825
               TabIndex        =   10
               Top             =   120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label lbl检查信息 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "检查信息"
               ForeColor       =   &H00C00000&
               Height          =   180
               Left            =   90
               TabIndex        =   6
               Top             =   450
               Width           =   720
            End
            Begin VB.Label lbl个人信息 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "个人信息"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   250
               Left            =   90
               TabIndex        =   5
               Top             =   150
               Width           =   900
            End
         End
      End
      Begin XtremeSuiteControls.TabControl TabWindow 
         Height          =   2415
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   4125
         _Version        =   589884
         _ExtentX        =   7276
         _ExtentY        =   4260
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   30
      ScaleHeight     =   4275
      ScaleWidth      =   4500
      TabIndex        =   12
      Top             =   495
      Width           =   4495
      Begin VB.TextBox txtFilter 
         Appearance      =   0  'Flat
         Height          =   250
         Left            =   855
         TabIndex        =   14
         ToolTipText     =   "*门诊号；+住院号；或手选查找方式；姓名+“*”为模糊查询；输入完成后直接回车开始查找"
         Top             =   45
         Width           =   1485
      End
      Begin VB.TextBox txtAppend 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDD6C6&
         BorderStyle     =   0  'None
         Height          =   2100
         Left            =   630
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1605
         Width           =   2010
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2685
         Left            =   450
         TabIndex        =   15
         Top             =   435
         Width           =   3360
         _cx             =   5927
         _cy             =   4736
         Appearance      =   2
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
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
         Begin VB.CommandButton cmdInfo 
            Caption         =   "…"
            Height          =   240
            Left            =   2730
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "选择项目(*)"
            Top             =   270
            Visible         =   0   'False
            Width           =   270
         End
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
   Begin VB.Timer TimerRefresh 
      Enabled         =   0   'False
      Left            =   7200
      Top             =   600
   End
   Begin zlIDKind.IDKind IDKind 
      Bindings        =   "frmPACStation.frx":1CFA
      Height          =   360
      Left            =   4815
      TabIndex        =   11
      Top             =   225
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   635
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6945
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPACStation.frx":1D0E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7805
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList Imglist 
      Left            =   6570
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":25A2
            Key             =   "紧急"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":2B3C
            Key             =   "住院"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":3416
            Key             =   "阳性"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":3570
            Key             =   "影像"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":3CEA
            Key             =   "收费"
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":4084
            Key             =   "绿色通道"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":41DE
            Key             =   "路径"
            Object.Tag             =   "7"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5940
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":4778
            Key             =   "复选留空"
            Object.Tag             =   "90000"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPACStation.frx":4D12
            Key             =   "复选选中"
            Object.Tag             =   "90001"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPACStation.frx":52AC
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPACStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mintCur业务类型 As Integer = 1 '当前系统操作的业务类型

Private Const ConstrCol = "路径;400|紧急;300|来源;400|收费;300|阳性;300|质量;300|姓名;1200|检查号;1400|检查过程;800|性别;450|年龄;450" & _
                        "|标识号;1400|医嘱内容;2400|部位方法;1400|执行间;600|报到时间;1800|申请时间;1800|开嘱医生;800" & _
                        "|身高;450|体重;450|婴儿;450|登记人;800|报到人;800|完成人;800|打印胶片;800|报告操作;800" & _
                        "|绿色通道;0|报告打印;800|报告人;800|复核人;800|检查技师;800|采图时间;1800|随访描述;2400" & _
                        "|影像类别;0|病人ID;0|主页ID;0|挂号单;0|病人科室ID;0|医嘱ID;1200|发送号;0|检查UID;0" & _
                        "|检查状态;0|NO;0|记录性质;0|转出;0|床号;0|当前病区ID;0|报告发放;800|诊断分类;800" & _
                        "|执行科室ID;0|关联ID;0|病人科室;800|就诊卡号;800|单据号;800|身份证号;800"
Private mstrCol As String   '列表顺序窗体加载时读取注册表，若无值用ConstrCol为默认值

'ID_查找方式+100之后保留7个是作为查找方式选择的
'ID_影像类别之后保留40个号码作为影像类别，从4021-4060
Private Enum FilterID
    ID_门诊 = 4001: ID_住院 = 4002: ID_体检 = 4003: ID_外诊 = 4004
    ID_费用 = 4005: ID_已缴 = 4006: ID_未缴 = 4007: ID_登记 = 4008
    ID_报到 = 4009: ID_检查 = 4010: ID_报告 = 4011: ID_审核 = 4012: ID_完成 = 4013
    ID_查找方式 = 4014: ID_查找值 = 4015: ID_开始查找 = 4016: ID_本次住院 = 4017
    ID_影像类别 = 4020
End Enum

Private mblncmd门诊 As Boolean, mblncmd住院 As Boolean, mblncmd体检 As Boolean, mblncmd外诊 As Boolean, mblncmd已缴 As Boolean, mblncmd未缴 As Boolean
Private mblncmd登记 As Boolean, mblncmd报到 As Boolean, mblncmd检查 As Boolean, mblncmd报告 As Boolean, mblncmd审核 As Boolean, mblncmd完成 As Boolean
Private mblncmd本次 As Boolean
Private mintcmd影像类别 As Integer      '0表示没有选择影像类别，其他数字表示选择的影像类别的数量
Private mblncmd影像类别() As Boolean    '保存当前选择的影像类别是否被选择

Private mstrFirstTab As String '首次显示的页面

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private mobjICCard As Object
Private Enum IDKinds
    C0姓名或就诊卡 = 0
    C1医保号 = 1
    C2身份证号 = 2
    C3IC卡号 = 3
End Enum

'子窗体对像
Private mfrmPACSImg As frmPACSImg       '影像子窗体
Private WithEvents mfrmPacsReport As frmReport     'PACS报告编辑器，嵌入主程序的窗口
Attribute mfrmPacsReport.VB_VarHelpID = -1
Private WithEvents mfrmPacsReportDock As frmReport     'PACS报告编辑器,独立窗口
Attribute mfrmPacsReportDock.VB_VarHelpID = -1
Private WithEvents mobjReport As zlRichEPR.cDockReport  '报告对象
Attribute mobjReport.VB_VarHelpID = -1
Private mobjExpense As zlCISKernel.clsDockExpense       '费用对象
Private WithEvents mobjInAdvice As zlCISKernel.clsDockInAdvices    '住院医嘱对象
Attribute mobjInAdvice.VB_VarHelpID = -1
Private WithEvents mobjOutAdvice As zlCISKernel.clsDockOutAdvices  '门诊医嘱对象
Attribute mobjOutAdvice.VB_VarHelpID = -1
Private mobjInEPRs As zlRichEPR.cDockInEPRs             '住院病历对象
Private mobjOutEPRs As zlRichEPR.cDockOutEPRs           '门诊病历对象
Private WithEvents mobjPacsCore As zl9PacsCore.clsViewer           '观片站对象
Attribute mobjPacsCore.VB_VarHelpID = -1
Private mobjPacsReportArry() As frmReport                   'PACS报告编辑器数组
Private mobjQueue As zlQueueManage.clsQueueManage          '排队叫号
'窗口变量
Private mlngCur科室ID As Long                               '当前科室ID
Private mstrCur科室 As String                               '当前科室 编码-名称
Private mstrCanUse科室 As String                            '当前可用科室  ID_编码-名称
Private mstrCurFindtype As String                           '过滤条件
Private mstrLocateType As String                            '定位条件
Private mblnInitOk As Boolean, mblnvsRefresh As Boolean     '初始化完成,装载表格
Private mstrPrivs As String, mlngModul As Long              '模块号，本模块权限
Private mblnAllDepts As Boolean                             '是否选择全部科室
Private mstrCanUse科室IDs As String                         '当前可用的科室ID串，用“，”分隔，可以直接作为SQL查询条件
Private mlngSortCol As Long                                 '病人列表中，当前进行排序的列
Private mintSortOrder As Integer                            '病人列表中，当前进行排序的方式

'流程控制变量
Private mblnFinishCommit As Boolean                         '无报告完成里,是否无需再次确认
Private mblnCompleteCommit As Boolean                       '审核后无需再次确认
Private mblnIgnoreResult As Boolean                         '忽略阴阳性 '=true 忽略
Private mintResultInput As Integer                          '提示输入阴阳性和影像质量
Private mblnShowImgAtReport As Boolean                      '打开报告时打开观片站
Private mblnReportWithImage As Boolean                      '有图像才能写报告，无图像不可写报告
Private mblnReportWithResult As Boolean                     '无影像诊断为阴性
Private mblnLocalizerBackward As Boolean                    '定位片后置
Private mblnPacsReport As Boolean                           '是否使用PACS报告编辑器，Fasle时使用电子病历编辑器
Private mblnPrintCommit As Boolean                          '打印后直接完成
Private mblnCanPrint As Boolean                             '平诊需要审核才能打印 =true
Private mBeforeDays As Integer                              '默认查询的天数
Private mlngRefreshInterval As Long                         '病人列表自动刷新间隔
Private mblnUseQueue As Boolean                             '是否启用排队叫号
Private mAstr队列名称() As String                           '排队叫号的队列名称，执行间的名称
Private mblnRelatingPatient As Boolean                      '是否启用关联病人
'本机本地参数
Private mstrRoom As String                                  '只处理执行间内的病人
Private mblnPatTrack As Boolean                             '是否对进病人进行跟踪
Private mbln直接检查 As Boolean                             '登记后直接进入检查
Private mblnNoShowCancel As Boolean                         '不显示取消的检查
Private mblnMoved As Boolean                                '当前时间段内是否被转移过
Private mblnOpenReport As Boolean                           '开始检查自动打开报告
Private mblnTechReptSame As Boolean                         '只能填写自己检查的报告
Private mblnUse3D As Boolean                                '是否启用三维重建功能
Private mstr3DExeDir As String                              '三维重建程序路径
Private mstr3DPara As String                                '三维重建参数
Private mstr3DFunctions As String                           '三维重建功能

'过滤条件变量
Private Type Type_SQLCondition
    开始时间 As Date
    结束时间 As Date
    时间类型 As Integer                                 '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
    单据号 As String
    门诊号 As Double
    住院号 As Double
    就诊卡 As String
    姓名 As String
    性别 As String
    开始年龄 As Long
    结束年龄 As Long
    年龄条件 As String
    检查号 As Double
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
End Type
Private SQLCondition As Type_SQLCondition

'历史记录的显示
Private mblnIsHistory As Boolean
Private mlngHOrderID As Long
Private mlngHSendNo As Long
Private mstrHStudyUID As String
Private mblnHMoved As Boolean


Private Sub Menu_File_Excel_click()
Dim bytMode As Byte
    On Error GoTo errHandle
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL

    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vsList
    objPrint.Title.Text = "检查病人清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    bytMode = zlPrintAsk(objPrint)
    If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_File_BatPrint()
    Dim cbrControl As CommandBarControl, strReturn As String, i As Integer
    Dim objReportPrint As New zlRichEPR.cDockReport
    Dim objPacsReport As New frmReport
    Dim strReportString As String
    
    Set cbrControl = Me.cbrMain(2).FindControl(, conMenu_File_Print)
    If Not cbrControl Is Nothing Then
        cbrControl.ID = conMenu_File_BatPrint
    Else
        Exit Sub
    End If

    '选病人
    strReturn = frmDocPrintPatiList.Showfrm(vsList, Me, mblnCanPrint, mblnPacsReport, mlngCur科室ID)
    '循环调用报告打印
    '对于使用PACS报告编辑器打印的，调用PACS报告编辑器窗口来打印
    '返回值由"医嘱ID-是否PACS报告编辑器-执行科室ID|医嘱ID-是否PACS报告编辑器-执行科室ID|..."组成
    For i = 0 To UBound(Split(strReturn, "|"))
        strReportString = Split(strReturn, "|")(i)
        '判断是否使用PACS报告编辑器
        If Split(strReportString, "-")(1) = 1 Then  '使用PACS报告编辑器
            Call objPacsReport.InitReportWindow(CLng(Split(strReportString, "-")(2)), mlngModul, mstrPrivs, True)
            objPacsReport.zlRefresh CLng(Split(strReportString, "-")(0)), Me, False
            Call objPacsReport.zlExecuteCommandBars(cbrControl)
            '需要AfterPrint吗？
        Else    '使用病历编辑器
            If objReportPrint.zlRefresh(CLng(Split(strReportString, "-")(0)), CLng(Split(strReportString, "-")(2)), , , True) > 0 Then
                Call objReportPrint.zlExecuteCommandBars(cbrControl)
                Call AfterPrinted(CLng(Split(strReportString, "-")(0)))
            End If
        End If
    Next
    cbrControl.ID = conMenu_File_Print
    Unload objReportPrint.zlGetForm
End Sub
Private Sub Menu_RichEPR(ByVal cbrID As Long)
'自动打开报告编辑器，同时处理了PACS报告编辑器和电子病历编辑器

    Dim cbrControl As CommandBarControl, i As Integer, blnCanPrint As Boolean
    
    '报告页面不可见时,查找报告页面，并显示出来，如果找不到，退出
    If (mblnPacsReport = True And TabWindow.Selected.Tag <> "PACS报告") _
        Or (mblnPacsReport = False And TabWindow.Selected.Tag <> "病历报告") Then
        If mblnPacsReport = True Then
            For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                If TabWindow(i).Tag = "PACS报告" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
            Next
            If TabWindow.Selected.Tag <> "PACS报告" Then Exit Sub
        Else
            For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                If TabWindow(i).Tag = "病历报告" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
            Next
            If TabWindow.Selected.Tag <> "病历报告" Then Exit Sub
        End If
    Else
        If TabWindow.Selected.Visible = False Then Exit Sub
    End If
    
    '找到报告页面，再打开这个报告页面
    With vsList
        blnCanPrint = IIf(mblnCanPrint, IIf(.Cell(flexcpData, .Row, GetCN("紧急")) = 1, .TextMatrix(.Row, GetCN("报告人")) <> "", .TextMatrix(.Row, GetCN("复核人")) <> ""), True)
        
        '刷新嵌入页面内容
        If mblnPacsReport = True Then
            Call mfrmPacsReport.zlRefresh(Val(.TextMatrix(.Row, GetCN("医嘱ID"))), Me, .TextMatrix(.Row, GetCN("转出")) = 1)
        Else
            Call mobjReport.zlRefresh(Val(.TextMatrix(.Row, GetCN("医嘱ID"))), mlngCur科室ID, True, .TextMatrix(.Row, GetCN("转出")) = 1, blnCanPrint)
        End If
    End With
    
    '判断按键可用性
    Set cbrControl = Me.cbrMain.FindControl(, cbrID, True, True)
    If cbrControl Is Nothing Then Exit Sub
    Call cbrMain_Update(cbrControl)
    If cbrControl.Enabled = False Then Exit Sub
        
    Call cbrMain_Execute(cbrControl)
End Sub
Private Sub Menu_File_Parmeter_click()
    With frmTechnicSetup
        .mlngModul = mlngModul
        .mlng科室ID = mlngCur科室ID
        .mstrPrivs = mstrPrivs
        .Show 1, Me
        If .mblnOK Then
            InitLocalPars
            Call RefreshList
        End If
    End With
End Sub

Private Sub Menu_Help_About_click()
    ShowAbout Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision
End Sub

Private Sub Menu_Help_Help_click()
    '功能：调用帮助主题
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Menu_Help_Web_Forum_click()
    Call zlWebForum(Me.hWnd)
End Sub


Private Sub Menu_Help_Web_Mail_click()
    zlMailTo hWnd
End Sub

Private Sub Menu_Manage_取消关联()
'取消关联的最后结果是，每次取消关联后，图象全部按照序列被拆散成N条临时记录
Dim strFilter As String, rsTmp As ADODB.Recordset, lngAdviceID As Long, lngSendNO As Long
    On Error GoTo errHandle
    '显示序列选择窗口
    With vsList
        lngAdviceID = Nvl(.TextMatrix(.Row, GetCN("医嘱ID")), 0)
        lngSendNO = Nvl(.TextMatrix(.Row, GetCN("发送号")), 0)
    End With
    
    gstrSQL = "select 0 as 选择,B.序列UID as ID ,B.序列号,B.序列描述,SUM(1) AS 图像数 from 影像检查记录 A ," & _
            "影像检查序列 B, 影像检查图象 C Where a.检查UID = B.检查UID And B.序列UID = C.序列UID" & _
            " And a.医嘱ID = [1] and A.发送号= [2] group by B.序列UID,B.序列号,B.序列描述"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceID, lngSendNO)
    
    frmSelectMuli.ShowSelect rsTmp, "ID,3000,0,1;序列号,800,0,1;序列描述,2000,0,1;图像数,800,0,1", 0, 0, 14000, 10000, "取消关联"
    
    If frmSelectMuli.mblnOK = True Then
        strFilter = frmSelectMuli.strFilter
        rsTmp.Filter = strFilter
        '如果有选中序列，则处理每一个序列的取消
        While Not rsTmp.EOF
            subCancelSeriesRelate Me, lngAdviceID, lngSendNO, rsTmp!ID
            rsTmp.MoveNext
        Wend
        
        '设置影像检查状态，如果当前医嘱已经没有图像，而且检查过程为3，则修改为2
        If vsList.TextMatrix(vsList.Row, GetCN("检查状态")) = 3 Then
            gstrSQL = "Select 检查uid From 影像检查记录 Where  医嘱ID=[1] And 发送号=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngAdviceID, lngSendNO)
            If IsNull(rsTmp!检查uid) Then
                gstrSQL = "Zl_影像检查_State(" & lngAdviceID & "," & lngSendNO & ",2)"
                zldatabase.ExecuteProcedure gstrSQL, "取消关联"
            End If
        End If
        
        mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
        Call RefreshList '真正取消关联点确定才刷新
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_无报告完成()
'只有进行中的报告可以操作该菜单,因为此时还没有签名

        On Error GoTo errHandle
        With vsList
            If .TextMatrix(.Row, GetCN("报告人")) <> "" Or .TextMatrix(.Row, GetCN("报告操作")) <> "" Then
                If MsgBoxD(Me, "是否无报告直接完成,直接完成将删除已填写的报告!", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            
            If mblnFinishCommit And InStr(mstrPrivs, "检查完成") > 0 Then '无报告完成后无需再次确认完成,但需要有检查完成的权限
                '此过程,传状态=6,并且报告ID不为空将删除电子病历记录
                
                If bln费用未审核(.TextMatrix(.Row, GetCN("病人ID")), Val(.TextMatrix(.Row, GetCN("主页ID"))), _
                    .TextMatrix(.Row, GetCN("医嘱ID")), CLng(Decode(.TextMatrix(.Row, GetCN("来源")), "门", 1, "住", 2, "外", 3, 4))) Then
                                    
                    '执行后自动审核划价单有效，并且病人已出院，且有未审核的划价单
                    MsgBoxD Me, "该病人已出院，且有未审核的划价单不能完成！", vbExclamation, gstrSysName
                Else
                    gstrSQL = "ZL_影像检查_STATE(" & .TextMatrix(.Row, GetCN("医嘱ID")) & "," & .TextMatrix(.Row, GetCN("发送号")) & ",6,1)"
                End If
            Else
                gstrSQL = "ZL_影像检查_STATE(" & .TextMatrix(.Row, GetCN("医嘱ID")) & "," & .TextMatrix(.Row, GetCN("发送号")) & ",5,1)"
            End If
        End With
        
        Call zldatabase.ExecuteProcedure(gstrSQL, "改变检查过程")
        
        '取消排队信息
        If mblnUseQueue = True And Not mobjQueue Is Nothing Then
            Call mobjQueue.zlQueueExec(mlngCur科室ID & ":" & vsList.TextMatrix(vsList.Row, GetCN("执行间")), _
                                        mintCur业务类型, vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), 4)
        End If
        
        If mblnPatTrack Then
            If mblnFinishCommit Then
                Call StateCheck(6)
            Else
                Call StateCheck(5)
            End If
        Else
            Call RefreshList
        End If
        Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Edit_无报告回退()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If MsgBoxD(Me, "确认要回退该项检查吗？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    With vsList
            '如果有图像，则回退到“已检查”，否则回退到“已报到”
            gstrSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否有图像", CLng(.TextMatrix(.Row, GetCN("医嘱ID"))))
            
            gstrSQL = "ZL_影像检查_STATE(" & .TextMatrix(.Row, GetCN("医嘱ID")) & "," & .TextMatrix(.Row, GetCN("发送号")) & "," & IIf(Nvl(rsTemp!检查uid) = "", 2, 3) & ")"
            zldatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    
    If mblnPatTrack Then
        Call StateCheck(2)
    Else
        Call RefreshList
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_检查最终完成(Optional lng医嘱ID As Long = 0, Optional blnRefresh As Boolean = True)
'可能由其它过程调用，此时传入有医嘱ID，但需要权限判断
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If lng医嘱ID = 0 Then
        lng医嘱ID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))
    End If
    
    If InStr(mstrPrivs, "检查完成") <= 0 Then Exit Sub
    
    gstrSQL = "Select a.发送号,b.病人ID,b.主页ID From 病人医嘱发送 a,病人医嘱记录 b Where a.医嘱id = [1] And a.医嘱ID=b.Id"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "检查最终完成", lng医嘱ID)
    
    If rsTemp.EOF = True Then Exit Sub
    
    If bln费用未审核(rsTemp!病人ID, Nvl(rsTemp!主页ID, 0), _
        Nvl(lng医嘱ID), CLng(Decode(vsList.TextMatrix(vsList.Row, GetCN("来源")), "门", 1, "住", 2, "外", 3, 4))) Then
    
        '执行后自动审核划价单有效，并且病人已出院，且有未审核的划价单
        MsgBoxD Me, "该病人已出院，且有未审核的划价单，不能完成！", vbExclamation, gstrSysName
    Else
        gstrSQL = "ZL_影像检查_STATE(" & lng医嘱ID & "," & rsTemp!发送号 & ",6)"
    End If

    Call zldatabase.ExecuteProcedure(gstrSQL, "改变检查过程")

    '取消排队信息
    If mblnUseQueue = True And Not mobjQueue Is Nothing Then
        Call mobjQueue.zlQueueExec(mlngCur科室ID & ":" & vsList.TextMatrix(vsList.Row, GetCN("执行间")), _
                                    mintCur业务类型, vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), 4)
    End If

    If blnRefresh Then Call StateCheck(6)
    Exit Sub

errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_取消检查完成()
    Dim intState As Integer
    
    On Error GoTo errHandle
    
    With vsList
            If .TextMatrix(.Row, GetCN("转出")) = 1 Then
                MsgBoxD Me, "该病人的本次住院数据已经转出到后备数据库，不允许操作。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            intState = getStudyState(.TextMatrix(.Row, GetCN("医嘱ID")))
            gstrSQL = "ZL_影像检查_STATE(" & .TextMatrix(.Row, GetCN("医嘱ID")) & "," & .TextMatrix(.Row, GetCN("发送号")) & "," & intState & ")"
            zldatabase.ExecuteProcedure gstrSQL, "取消检查完成"
    End With
    
    Call StateCheck(intState)
    Exit Sub

errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_标记阴阳(ByVal lngID As Long)
    Dim iresult As Integer

    On Error GoTo errHandle
    Select Case lngID
        Case conMenu_Manage_Negative
            iresult = 1
        Case conMenu_Manage_Positive
            iresult = 0
    End Select
    With vsList
        gstrSQL = "ZL_影像检查_结果(" & .TextMatrix(.Row, GetCN("医嘱ID")) & "," & iresult & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, "结果阴阳性")
        
        If iresult = 1 Then
            Set .Cell(flexcpPicture, .Row, GetCN("阳性")) = Imglist.ListImages("阳性").Picture
        Else
            Set .Cell(flexcpPicture, .Row, GetCN("阳性")) = Nothing
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_绿色通道(ByVal lngID As Long)
    Dim intResult As Integer

    On Error GoTo errHandle
    Select Case lngID
        Case conMenu_Manage_GChannelOk
            intResult = "1"
        Case conMenu_Manage_GChannelCancel
            intResult = "0"
    End Select
    With vsList
        gstrSQL = "Zl_绿色通道_Update(" & .TextMatrix(.Row, GetCN("医嘱ID")) & ",'" & intResult & "')"
        Call zldatabase.ExecuteProcedure(gstrSQL, "绿色通道")
        .TextMatrix(.Row, GetCN("绿色通道")) = intResult
        If intResult = 1 Then
            Set .Cell(flexcpPicture, .Row, GetCN("姓名")) = Imglist.ListImages("绿色通道").Picture
        Else
            Set .Cell(flexcpPicture, .Row, GetCN("姓名")) = Nothing
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_影像质量(ByVal lngID As Long)
    Dim strResult As String

    On Error GoTo errHandle
    Select Case lngID
        Case conMenu_Manage_First
            strResult = "甲"
        Case conMenu_Manage_Second
            strResult = "乙"
    End Select
    With vsList
        gstrSQL = "Zl_影像质量_Update(" & .TextMatrix(.Row, GetCN("医嘱ID")) & ",'" & strResult & "')"
        Call zldatabase.ExecuteProcedure(gstrSQL, "影像质量")
        .TextMatrix(.Row, GetCN("质量")) = strResult
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_技师执行()
    '技师执行
    Dim strSQL As String
    
    On Error GoTo err
    strSQL = "Zl_影像技师执行('" & UserInfo.姓名 & "'," & vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) & ")"
    Call zldatabase.ExecuteProcedure(strSQL, "技师执行")
    vsList.TextMatrix(vsList.Row, GetCN("检查技师")) = UserInfo.姓名
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_报告发放()
    '报告发放
    Dim strSQL As String
    
    On Error GoTo err
    
    strSQL = "Zl_影像报告发放(" & vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) & ")"
    Call zldatabase.ExecuteProcedure(strSQL, "报告发放")
    vsList.TextMatrix(vsList.Row, GetCN("报告发放")) = IIf(vsList.TextMatrix(vsList.Row, GetCN("报告发放")) = "未发放", "已发放", "未发放")
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_修改()
    
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = vsList.TextMatrix(vsList.Row, GetCN("发送号"))
        .mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))
        .mintEditMode = IIf(vsList.TextMatrix(vsList.Row, GetCN("检查状态")) > 1, 3, 1) '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = mlngCur科室ID
        .InitMvar
        If .RefreshPatiInfor(False) Then   '刷新病人
            .mblnOK = False
            .zlShowMe Me
        End If
        If .mblnOK Then RefreshList '成功返回
    End With
End Sub
Private Sub Menu_Manage_复制登记()
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = mlngCur科室ID
        .mblnOK = False
        .InitMvar
        If .CopyCheck(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), vsList.TextMatrix(vsList.Row, GetCN("发送号"))) = True Then  '刷新病人
            .zlShowMe Me
        End If
        If .mblnOK Then '成功返回
            If mbln直接检查 Then
                Call StateCheck(2, .mlngAdviceID)
            Else
                Call RefreshList
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_登记()
    With frmRISRequest
        .mlngModul = mlngModul
        .mlngSendNo = 0
        .mlngAdviceID = 0
        .mintEditMode = 0 '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = mlngCur科室ID
        .mblnOK = False
        .InitMvar
        .zlShowMe Me
        If .mblnOK Then '成功返回
            If mbln直接检查 Then
                Call StateCheck(2, .mlngAdviceID)
            Else
                Call RefreshList
            End If
            
            If vsList.Rows = 2 Then
              Call vsList.Select(1, 1)
            End If
        End If
    End With
End Sub
Private Sub Menu_Manage_取消登记()
    On Error GoTo errHandle
    
    If MsgBoxD(Me, "确认要取消当前申请吗？" & Chr(10) & Chr(13) & "申请取消后，其对应的医嘱将拒绝执行！", vbExclamation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "ZL_病人医嘱执行_拒绝执行(" & vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) & "," & vsList.TextMatrix(vsList.Row, GetCN("发送号")) & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, "撤消登记")
    Call RefreshList
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_召回取消()
'功能：召回被取消的登记
    On Error GoTo errH
    
    If MsgBoxD(Me, "确实要召回被取消登记的项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "ZL_病人医嘱执行_取消拒绝(" & vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) & "," & vsList.TextMatrix(vsList.Row, GetCN("发送号")) & ")"
    Call zldatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call RefreshList
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub Menu_Manage_报到()
Dim blnFocusFind As Boolean
Dim rsTemp As ADODB.Recordset
    If Me.ActiveControl Is Nothing Then
        blnFocusFind = False
    Else
        blnFocusFind = (Me.ActiveControl.Name = "txtFilter")
    End If
    With frmRISRequest
        .mstrPrivs = mstrPrivs
        .mlngModul = mlngModul
        .mlngSendNo = vsList.TextMatrix(vsList.Row, GetCN("发送号"))
        .mlngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))
        .mintEditMode = 2 '0－登记、1－登记后修改、2－报到、3－报到后修改
        .mlngCurDeptId = mlngCur科室ID
        .InitMvar
        If .RefreshPatiInfor(True) = True Then  '刷新病人
            .mblnOK = False
            .zlShowMe Me
        End If
        If .mblnOK Then  '成功返回
            Call StateCheck(2)
            If mblnOpenReport Then Call Menu_RichEPR(conMenu_Edit_Modify)              '开始检查自动打开报告
        End If
        If blnFocusFind Then txtFilter.SetFocus '自动定位到定位栏
    End With
End Sub
Private Sub Menu_Manage_取消报到()
Dim rsTemp As ADODB.Recordset, lngAdviceID As Long
    
    On Error GoTo errHandle
    With vsList
        If .TextMatrix(.Row, GetCN("检查状态")) <= 1 Then Call Menu_Manage_取消登记: Exit Sub '工具栏调用
        '------------------------------------有签名的需要先回退签名后再撤消
        lngAdviceID = .TextMatrix(.Row, GetCN("医嘱ID"))
        gstrSQL = "Select Distinct B.完成时间 From 病人医嘱报告 A, 电子病历记录 B Where A.病历ID=B.Id And A.医嘱ID=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取是否签名", lngAdviceID)
        If Not rsTemp.EOF Then
            If Nvl(rsTemp!完成时间, "") <> "" Then '签名保存
                MsgBoxD Me, "当前病人的检查报告已经签名,若需取消检查,请先回退签名!", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If

        If MsgBoxD(Me, "取消本次检查将删除相应的检查图像和检查报告，是否继续？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        If .TextMatrix(.Row, GetCN("检查UID")) <> "" And InStr(mstrPrivs, "清除图像") <= 0 Then
            MsgBoxD Me, "您没有清除检查图像权限,不能请除图像,所有不能取消此项检查!", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '取消排队信息
        If mblnUseQueue = True And Not mobjQueue Is Nothing Then
            'Call mobjQueue.zlDelQueue(Split(mstrCur科室, "-")(1) & .TextMatrix(.Row, GetCN("执行间")), lngAdviceID)
            Call mobjQueue.zlDelQueue(mlngCur科室ID & ":" & .TextMatrix(.Row, GetCN("执行间")), lngAdviceID)
        End If
        
        gstrSQL = "ZL_影像检查_CANCEL(" & lngAdviceID & "," & .TextMatrix(.Row, GetCN("发送号")) & ",0)"
        zldatabase.ExecuteProcedure gstrSQL, Me.Caption
        '删除影像文件和目录
        RemoveCheckImages lngAdviceID, .TextMatrix(.Row, GetCN("发送号"))
    End With
    
    Call StateCheck(1)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_观片()
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        If mblnIsHistory = True Then
            Call mfrmPACSImg.zlRefresh(mlngHOrderID, mlngHSendNo, mstrPrivs, mblnHMoved)
        Else
            Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), vsList.TextMatrix(vsList.Row, GetCN("发送号")), mstrPrivs, vsList.TextMatrix(vsList.Row, GetCN("转出")) = 1)
        End If
    End If
    Call mfrmPACSImg.zlMenuClick("影像处理")
End Sub
Private Sub Menu_Manage_对比观片()
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        If mblnIsHistory = True Then
            Call mfrmPACSImg.zlRefresh(mlngHOrderID, mlngHSendNo, mstrPrivs, mblnHMoved)
        Else
            Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), vsList.TextMatrix(vsList.Row, GetCN("发送号")), mstrPrivs, vsList.TextMatrix(vsList.Row, GetCN("转出")) = 1)
        End If
    End If
    Call mfrmPACSImg.zlMenuClick("影像对比")
End Sub
            
Private Sub Menu_Manage_图象删除()
    Dim rsTemp As ADODB.Recordset, lngAdviceID As Long, lngSendNO As Long
    
    On Error GoTo errHandle
    With vsList
        lngAdviceID = .TextMatrix(.Row, GetCN("医嘱ID"))
        lngSendNO = .TextMatrix(.Row, GetCN("发送号"))
    End With
    
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(lngAdviceID, lngSendNO, mstrPrivs, vsList.TextMatrix(vsList.Row, GetCN("转出")) = 1)
    End If
    
    gstrSQL = "select 检查UID from 影像检查记录 where 医嘱ID =[1] and  发送号 = [2]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取检查UID", lngAdviceID, lngSendNO)
    If rsTemp.EOF Then Exit Sub
    
    If MsgBoxD(Me, "是否确认要删除该检查的所有影像？", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    '删除影像文件和目录
    RemoveCheckImages lngAdviceID, lngSendNO
    gstrSQL = "ZL_影像检查_PhotoDelete(" & lngAdviceID & "," & lngSendNO & ")"
    zldatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    '设置影像检查状态，如果检查过程为3，则修改为2
    If vsList.TextMatrix(vsList.Row, GetCN("检查状态")) = 3 Then
        gstrSQL = "Zl_影像检查_State(" & lngAdviceID & "," & lngSendNO & ",2)"
        zldatabase.ExecuteProcedure gstrSQL, "删除图像"
    End If
    
    mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
    Call RefreshList
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
        
Private Sub Menu_Manage_获取图像()
Dim strImageDeviceNumber As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), vsList.TextMatrix(vsList.Row, GetCN("发送号")), mstrPrivs, vsList.TextMatrix(vsList.Row, GetCN("转出")) = 1)
    End If
    
    strImageDeviceNumber = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSImageDeviceSetup", "默认影像设备", "")
    
    '没有默认设备时处理
    If strImageDeviceNumber = "" Then
        If MsgBoxD(Me, "没有设置默认影像检查设备！是否现在设置？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        Else
            frmPACSImageDeviceSetup.Show vbModal, Me
            strImageDeviceNumber = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPACSImageDeviceSetup", "默认影像设备", "")
            If strImageDeviceNumber = "" Then Exit Sub
        End If
    End If
    
    gstrSQL = "select 设备号,设备名, IP地址,端口号,本地AE,设备AE from 影像设备目录 where 设备号 = [1] "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(Mid(strImageDeviceNumber, 2)))
    
    '当默认设备被删除后重新设置
    If rsTemp.EOF = True Then
        MsgBoxD Me, "默认设备已被删除，请重新设置！", vbInformation, gstrSysName
        frmPACSImageDeviceSetup.Show vbModal, Me
        Exit Sub
    End If
        
    '先判断设备的AE，端口是否被正确设置了,未设置好则提示并退出
    If IsNull(rsTemp("端口号")) Or IsNull(rsTemp("设备AE")) Or IsNull(rsTemp("本地AE")) Then
        MsgBoxD Me, "请到“影像设备目录”模块中，设置Q/R查询使用的设备端口号，设备AE和本地AE。"
        Exit Sub
    End If
    
    frmPACSGetDeviceImage.ShowMe Me, rsTemp("IP地址"), rsTemp("端口号"), rsTemp("设备名"), rsTemp("本地AE"), rsTemp("设备AE"), vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))
    Call RefreshList
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub Menu_Manage_关联影像()
Dim strSQL As String
Dim rsTemp As ADODB.Recordset, lngAdviceID As Long, lngSendNO As Long
    
    On Error GoTo errHandle
    With vsList
        lngAdviceID = .TextMatrix(.Row, GetCN("医嘱ID"))
        lngSendNO = .TextMatrix(.Row, GetCN("发送号"))

        Call funRelateSeries(Me, lngAdviceID, lngSendNO)
        '设置影像检查状态，如果原来的状态是已报到，则修改成已检查，
        If .TextMatrix(.Row, GetCN("检查状态")) < 3 Then
            '如果病人已经有图像，则修改成已检查
            strSQL = "Select 检查UID From 影像检查记录 Where 医嘱ID=[1]"
            Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "检查是否有图像", lngAdviceID)
            
            If Not IsNull(rsTemp!检查uid) Then
                gstrSQL = "Zl_影像检查_State(" & lngAdviceID & "," & lngSendNO & ",3)"
                zldatabase.ExecuteProcedure gstrSQL, "关联影像"
            End If
        End If
    End With
    mfrmPACSImg.zlRefresh 0, 0, mstrPrivs
    Call RefreshList
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_View_Locate_Type_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    mstrLocateType = Split(control.Caption, "(")(0)
    cbrMain.RecalcLayout
    If mstrLocateType = "ＩＣ卡" Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
        Else
            txtLocate.Text = mobjICCard.Read_Card(Me)
        End If
    End If
    txtLocate.SetFocus
End Sub

Private Sub Menu_Dept_Select(ByVal control As XtremeCommandBars.ICommandBarControl)
    '更换科室，根据新的条件，重新过滤病人
    '如果选择的是全部科室，则 mlngCur科室ID 不改变
    '如果选择的是某个具体科室，则改变 mlngCur科室ID
    If mlngCur科室ID <> control.DescriptionText Or (control.DescriptionText <> 0 And mblnAllDepts = True) Then
        '选择了具体科室，才改变当前科室的设置
        If control.DescriptionText = 0 Then
            mblnAllDepts = True
        Else
            mblnAllDepts = False
            mlngCur科室ID = control.DescriptionText
            mstrCur科室 = Split(control.Caption, "(")(0)
            
            Call ReadStudyListColor(mlngCur科室ID)
            
            Call mfrmPacsReport.InitReportWindow(mlngCur科室ID, mlngModul, mstrPrivs, True)
        End If
        
        Call cbrMain.RecalcLayout
        Call subChangeDept(0)
    End If
End Sub
Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer
    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
Dim i As Integer, cbrControl As CommandBarControl
    For i = 2 To cbrMain.Count
        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
    Next
    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
End Sub

Private Sub cboTimes_Click()
    If cboTimes.ListCount <= 1 Then Exit Sub
    If cboTimes.Tag = "" Then Exit Sub '此时cbotime项目未增加完成，属listindex赋值触发
    
    On Error GoTo errHandle
    Dim lngAdviceID As Long
    lngAdviceID = cboTimes.ItemData(cboTimes.ListIndex)
    If lngAdviceID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) Then Call vsList_RowColChange: Exit Sub '当次与当前选中医嘱ID相同时不由本函数控制

    mblnIsHistory = True: mlngHOrderID = lngAdviceID '以下三个过程调用有先后顺序，勿调换
    Call FillTxtInfor(mlngHOrderID)  '填充右上方病人基本信息
    Call FillTxtAppend(mlngHOrderID) '填充左下角医嘱附件
    
    '选择了全部科室，而且切换了科室
    If mlngCur科室ID <> Split(fraInfo.Tag, "|")(10) And mblnAllDepts = True Then
        mlngCur科室ID = Split(fraInfo.Tag, "|")(10)
        mstrCur科室 = GetDeptName(mlngCur科室ID, mstrCanUse科室)
        Call subChangeDept(1)
    End If
        
    Call ShowTab(mlngHOrderID)  '根据病人提供不同选项卡
    Call RefreshTabWindow(mlngHOrderID) '刷新子窗体
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboTimes_DropDown()
    Call SendMessage(cboTimes.hWnd, &H160, 500, 0)
End Sub

Private Sub cbrdock_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim strTemp As String
    
    Select Case control.ID
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
            If mblncmd已缴 Then mblncmd未缴 = False
        Case ID_未缴
            mblncmd未缴 = Not control.Checked
            If mblncmd未缴 Then mblncmd已缴 = False
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
                strTemp = "影像类别"
            Else
                strTemp = ""
                For i = 1 To objControl.CommandBar.Controls.Count
                    If objControl.CommandBar.FindControl(, ID_影像类别 + i).Checked = True Then
                        strTemp = IIf(strTemp = "", objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption, strTemp & "," & objControl.CommandBar.FindControl(, ID_影像类别 + i).Caption)
                    End If
                Next i
            End If
            objControl.Caption = strTemp
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
        Case ID_完成
            mblncmd完成 = Not control.Checked
        Case ID_本次住院
            control.Checked = Not control.Checked
            mblncmd本次 = Not mblncmd本次
        Case ID_查找方式 * 100# To ID_查找方式 * 100# + 7
            mstrCurFindtype = Split(control.Caption, "(")(0)
            If InStr(mstrCurFindtype, "ＩＣ卡") > 0 Then
                If mobjICCard Is Nothing Then
                    Set mobjICCard = CreateObject("zlICCard.clsICCard")
                End If
                txtFilter.Text = mobjICCard.Read_Card(Me)
            End If
            
            If txtFilter.PasswordChar = "*" Then '之前是就诊卡号，需要清除并变更掩码
                txtFilter.Text = "": txtFilter.PasswordChar = ""
            End If
            
            txtFilter_GotFocus
            cbrdock.RecalcLayout
            Exit Sub
        Case ID_开始查找
            Call subRefreshFilterCondition(txtFilter.Text)
    End Select
cbrdock.RecalcLayout
Call RefreshList
End Sub

Private Sub subRefreshFilterCondition(strFilter As String)
'------------------------------------------------
'功能：用txtFilter控件的内容更新过滤条件
'参数： strFilter --- 过滤条件
'返回：无
'------------------------------------------------

    On Error GoTo err
    
    With SQLCondition
        .姓名 = ""
        .就诊卡 = ""
        .门诊号 = 0
        .住院号 = 0
        .单据号 = ""
        .检查号 = 0
        .身份证 = ""
        .IC卡 = ""
        Select Case mstrCurFindtype
            Case "姓  名"
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
            Case "单据号"
                .单据号 = Trim(strFilter)
            Case "检查号"
                .检查号 = Val(strFilter)
            Case "身份证"
                .身份证 = Trim(strFilter)
            Case "ＩＣ卡"
                .IC卡 = Trim(strFilter)
        End Select
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbrdock_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    If CommandBar.Parent Is Nothing Then Exit Sub
    If CommandBar.Parent.ID = ID_查找方式 Then
        With CommandBar.Controls
            If .Count = 0 Then '动态子菜单,扩1位
                Set objControl = .Add(xtpControlButton, ID_查找方式 * 100# + 0, "门诊号(&1)")
                Set objControl = .Add(xtpControlButton, ID_查找方式 * 100# + 1, "住院号(&2)")
                Set objControl = .Add(xtpControlButton, ID_查找方式 * 100# + 2, "就诊卡(&3)")
                Set objControl = .Add(xtpControlButton, ID_查找方式 * 100# + 3, "姓  名(&4)")
                Set objControl = .Add(xtpControlButton, ID_查找方式 * 100# + 4, "单据号(&5)")
                Set objControl = .Add(xtpControlButton, ID_查找方式 * 100# + 5, "检查号(&6)")
                Set objControl = .Add(xtpControlButton, ID_查找方式 * 100# + 6, "身份证(&7)")
                Set objControl = .Add(xtpControlButton, ID_查找方式 * 100# + 7, "ＩＣ卡(&8)")
            End If
        End With
    End If
End Sub

Private Sub cbrdock_Resize()
Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    Call Me.cbrdock.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    vsList.Top = lngTop: vsList.Left = lngLeft
    vsList.Width = picList.Width
    vsList.Height = picList.Height - lngTop - txtAppend.Height - 100


    txtAppend.Top = vsList.Top + vsList.Height + 100: txtAppend.Left = lngLeft + 100
    txtAppend.Width = picList.Width - 200
End Sub

Private Sub cbrdock_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
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
            control.Checked = mblncmd已缴 Xor mblncmd未缴
            control.Caption = IIf(mblncmd已缴 Xor mblncmd未缴, IIf(mblncmd已缴, " 已缴费", " 未缴费"), " 费  用")
        Case ID_已缴
            control.Checked = mblncmd已缴
            control.IconId = IIf(mblncmd已缴, 90001, 90000)
        Case ID_未缴
            control.Checked = mblncmd未缴
            control.IconId = IIf(mblncmd未缴, 90001, 90000)
        Case ID_影像类别
            control.IconId = IIf(mintcmd影像类别 = 0, 90000, 90001)
        Case ID_影像类别 + 1 To ID_影像类别 + 40
            control.Checked = mblncmd影像类别(control.ID - ID_影像类别 - 1)
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
        Case ID_完成
            control.Checked = mblncmd完成
            control.IconId = IIf(mblncmd完成, 90001, 90000)
        Case ID_本次住院
            control.IconId = IIf(control.Checked, 90001, 90000)
        Case ID_查找方式
            control.Caption = mstrCurFindtype
        Case ID_查找方式 * 100# To ID_查找方式 * 100# + 7
            control.Checked = (InStr(control.Caption, mstrCurFindtype) > 0)
    End Select
End Sub
Private Sub cbrMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = stbThis.Height
End Sub


'费用执行
Private Sub ExecuteStudyMoney()
    On Error GoTo errHandle
      
    Dim lngAdviceID As Long, lngSendNO As Long
    
    With vsList
        lngAdviceID = Nvl(.TextMatrix(.Row, GetCN("医嘱ID")), 0)
        lngSendNO = Nvl(.TextMatrix(.Row, GetCN("发送号")), 0)
    End With
    
    gstrSQL = "Zl_影像费用执行(" & lngAdviceID & "," & lngSendNO & ",2)"
    zldatabase.ExecuteProcedure gstrSQL, "费用执行"
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    
    If control.ID <> 0 Then
        If cbrMain.FindControl(, control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    cbrMain.RecalcLayout
    Select Case control.ID
    
'--------------------------文件------------------
        Case conMenu_File_PrintSet '打印设置
            Call zlPrintSet
           
        Case conMenu_File_Excel '清单打印
            Call Menu_File_Excel_click
            
        Case conMenu_File_BatPrint '批量打印
            Call Menu_File_BatPrint
            
        Case conMenu_File_Parameter '参数设置
            Call Menu_File_Parmeter_click
            
        Case conMenu_Cap_DevSet '影像设备设置
            frmPACSImageDeviceSetup.Show vbModal, Me
            
        Case conMenu_File_SendImg '发送图像
            frmPacsSendImage.ShowMe Me
        
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
            
        Case conMenu_Manage_Logout                          '取消报到
            Call Menu_Manage_取消报到
            
        Case conMenu_Img_Look                         '观片
            Call Menu_Manage_观片
        
        Case conMenu_Img_Contrast                     '对比观片
            Call Menu_Manage_对比观片
        
        Case conMenu_Img_3D_MMPR                    '三维重建，MMPR
            Call sub三维重建("MMPR")
        Case conMenu_Img_3D_MPR                     '三维重建，MPR
            Call sub三维重建("MPR")
        Case conMenu_Img_3D_PF                     '三维重建,灌注成像
            Call sub三维重建("PF")
        Case conMenu_Img_3D_SA                     '三维重建，表面重建
            Call sub三维重建("SA")
        Case conMenu_Img_3D_VA                     '三维重建，容积重建
            Call sub三维重建("VA")
        Case conMenu_Img_3D_VE                     '三维重建，虚拟内窥镜
            Call sub三维重建("VE")
            
        Case conMenu_Img_Delete                       '图象删除
            Call Menu_Manage_图象删除
        
        Case conMenu_Img_Query                        '从设备获取图象
            Call Menu_Manage_获取图像
        
        Case conMenu_Manage_Transfer                        '关联影像
            Call Menu_Manage_关联影像
            
        Case conMenu_Manage_Cancel                          '取消关联
            Call Menu_Manage_取消关联
        
        Case conMenu_Manage_Review                          '备注
            Call Menu_Manage_随访
        
        Case conMenu_Manage_Negative, conMenu_Manage_Positive                  '结果阴阳性
            Call Menu_Manage_标记阴阳(control.ID)
            
        Case conMenu_Manage_First, conMenu_Manage_Second
            Call Menu_Manage_影像质量(control.ID)
            
        Case conMenu_Manage_TechDoctorExecute               '技师执行
            Call Menu_Manage_技师执行
            
        Case conMenu_Manage_ReportRelease                   '报告发放
            Call Menu_Manage_报告发放
            
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
            
        Case conMenu_Manage_ChangeDevice                    '更换检查设备
            Call Menu_Manage_更换检查设备
        
        Case conMenu_Manage_RelatingPatiet                  '关联病人
            Call Menu_Manage_关联病人
        Case conMenu_File_Preview, conMenu_File_Print       '报告预览和打印
            Dim i As Integer
            '没报告不能打印和预览
            If vsList.TextMatrix(vsList.Row, GetCN("报告人")) = "" Then
                MsgBoxD Me, "当前病人没有检查报告，不能操作，请检查！", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '报告页面不可见时不执行任何操作
            If mblnPacsReport = True Then
                If TabWindow.Selected.Tag <> "PACS报告" Then
                    For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                        If TabWindow(i).Tag = "PACS报告" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
                    Next
                End If
                If TabWindow.Selected.Tag = "PACS报告" Then
                    mfrmPacsReport.zlExecuteCommandBars control
                End If
            Else
                If TabWindow.Selected.Tag <> "病历报告" Then
                    For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                        If TabWindow(i).Tag = "病历报告" And TabWindow(i).Visible = True Then TabWindow(i).Selected = True
                    Next
                End If
                If TabWindow.Selected.Tag = "病历报告" Then
                    mobjReport.zlExecuteCommandBars control
                End If
            End If
        
'---------------------------查看----------------
        Case conMenu_View_ToolBar_Button                        '工具栏
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text                          '按钮文字
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size                          '大图标
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar                             '状态栏
            Call Menu_View_StatusBar_click(control)
        Case conMenu_Manage_LocateType * 10# To conMenu_Manage_LocateType * 10# + 6 '定位
            Call Menu_View_Locate_Type_click(control)
        Case conMenu_View_Filter                                '过滤
            Call Menu_View_Filter_click
        Case conMenu_View_Refresh                               '刷新
            Call RefreshList
        Case conMenu_Help_Help                                  '帮助
            Call Menu_Help_Help_click
        Case conMenu_Help_Web_Forum                             '网上中联
'            Case zlWebForum(Me.Hwnd)
        Case conMenu_Help_Web_Home                              '网上中联
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail                              '电邮中联
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About                                 '关于
            Call Menu_Help_About_click
        Case conMenu_View_Filter * 100# To conMenu_View_Filter * 100# + UBound(Split(mstrCanUse科室, "|")) + 1 '更改当前科室
            Call Menu_Dept_Select(control)
        Case conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99
            If control.parameter <> "" Then '执行发布到当前模块的报表
                With vsList
                    If .TextMatrix(.Row, GetCN("医嘱ID")) <> "" Then
                        Call ReportOpen(gcnOracle, Split(control.parameter, ",")(0), Split(control.parameter, ",")(1), Me, _
                            "NO=" & .TextMatrix(.Row, GetCN("NO")), "性质=" & .TextMatrix(.Row, GetCN("记录性质")), "医嘱id=" & .TextMatrix(.Row, GetCN("医嘱ID")), 1)
                    Else
                        Call ReportOpen(gcnOracle, Split(control.parameter, ",")(0), Split(control.parameter, ",")(1), Me, "", 1)
                    End If
                End With
            End If
        Case Else
            If vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) = "" Then Exit Sub
            
            Select Case TabWindow.Selected.Tag
                Case "PACS报告", "病历报告"
                    '报告被某人打开后再被允许它人编辑或修订
                    If control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Modify Or control.ID = conMenu_PacsReport_Open Or control.ID = conMenu_Edit_Delete Then
                        If CheckConcurrentReport(Me, vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = False Then Exit Sub
                    End If
                    
                    '控制 只能书写自己检查的报告,'不允许书写、修订、删除
                    If mblnTechReptSame = True _
                        And (control.ID = conMenu_Edit_Modify Or control.ID = conMenu_Edit_Audit Or control.ID = conMenu_Edit_Delete) _
                        And Nvl(vsList.TextMatrix(vsList.Row, GetCN("检查技师"))) <> "" _
                        And Nvl(vsList.TextMatrix(vsList.Row, GetCN("检查技师"))) <> UserInfo.姓名 Then
                        MsgBoxD Me, "你不是这个患者的检查技师，无法操作这份报", vbInformation, gstrSysName
                    Else
                        If mblnPacsReport = True Then
                            If control.ID = conMenu_PacsReport_Open Then   '打开报告窗体
                                Call Menu_Manage_PACS报告
                            Else
                                mfrmPacsReport.zlExecuteCommandBars control
                            End If
                        Else
                            mobjReport.zlExecuteCommandBars control
                        End If
                    End If
                Case "申请费用"
                    mobjExpense.zlExecuteCommandBars control
                    
                    '----------------------补费时，执行费用------------------
                    If control.ID = conMenu_Edit_Append _
                    Or control.ID = conMenu_Edit_Modify _
                    Or control.ID = conMenu_Edit_NewItem * 10# + 1 _
                    Or control.ID = conMenu_Edit_NewItem * 10# + 2 _
                    Or control.ID = conMenu_Edit_NewItem * 10# + 3 Then
            
                        If vsList.TextMatrix(vsList.Row, GetCN("检查状态")) >= 2 Then
                            Call ExecuteStudyMoney
                        End If
                    End If
                Case "住院医嘱"
                    mobjInAdvice.zlExecuteCommandBars control
                Case "门诊医嘱"
                    mobjOutAdvice.zlExecuteCommandBars control
                Case "住院病历"
                    mobjInEPRs.zlExecuteCommandBars control
                Case "门诊病历"
                    mobjOutEPRs.zlExecuteCommandBars control
                Case "排队叫号"
                    If Not mobjQueue Is Nothing Then
                        mobjQueue.zlExecuteCommandBars control
                    End If
            End Select
    End Select
End Sub

Private Sub Menu_View_Filter_click()
    On Error GoTo errHandle
    
    With frmPACSFilter
        .mlngModul = mlngModul
        .mBeforeDays = mBeforeDays
        .mDept = mlngCur科室ID '当前科室
        .Show 1, Me
        If Not .mblnOK Then Exit Sub '没有返回条件
        
        '当使用时间条件时，清空固定条件
        txtFilter.Text = ""
        SQLCondition.姓名 = ""
        SQLCondition.就诊卡 = ""
        SQLCondition.门诊号 = 0
        SQLCondition.住院号 = 0
        SQLCondition.单据号 = ""
        SQLCondition.检查号 = 0
        SQLCondition.身份证 = ""
        SQLCondition.IC卡 = ""
        
        SQLCondition.开始时间 = Format(.dtpBegin.value, "yyyy-MM-dd HH:mm:00")
        If Format(.dtpEnd.value, "yyyy-MM-dd HH:mm") = Format(.dtpEnd.Tag, "yyyy-MM-dd HH:mm") Then
            SQLCondition.结束时间 = CDate(0) '表示取当前时间
        Else
            SQLCondition.结束时间 = Format(.dtpEnd.value, "yyyy-MM-dd HH:mm:59")
        End If
        
        mblnMoved = MovedByDate(SQLCondition.开始时间)
        
        If .optFindType(1).value = True Then '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
            SQLCondition.时间类型 = 1
        ElseIf .optFindType(2).value = True Then
            SQLCondition.时间类型 = 2
        Else
            SQLCondition.时间类型 = 3
        End If
        
        If .cboPart.ListIndex <> 0 Then '检查标本部位
            SQLCondition.标本部位 = NeedName(.cboPart.Text)
        Else
            SQLCondition.标本部位 = ""
        End If
        
        '病人性别
        If NeedName(.cboSex.Text) = "全部" Then
            SQLCondition.性别 = ""
        Else
            SQLCondition.性别 = NeedName(.cboSex.Text)
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
        
        If .cboDept.ListIndex <> 0 Then '病人科室
            SQLCondition.病人科室 = .cboDept.ItemData(.cboDept.ListIndex)
        Else
            SQLCondition.病人科室 = 0
        End If

        If .cbodiagdoc.ListIndex <> 0 Then '诊断医生
            SQLCondition.诊断医生 = NeedName(.cbodiagdoc.Text)
        Else
            SQLCondition.诊断医生 = ""
        End If
        
        If .cboAuditing.ListIndex <> 0 Then '审核医生
            SQLCondition.审核医生 = NeedName(.cboAuditing.Text)
        Else
            SQLCondition.审核医生 = ""
        End If
        
        
'        If .cboCheckStep.ListIndex <> 0 Then '检查过程
'            SQLCondition.检查过程 = .cboCheckStep.Text
'        Else
'            SQLCondition.检查过程 = ""
'        End If
        
        
        If .cboModality.ListIndex <> 0 Then '影像类别
            SQLCondition.影像类别 = Split(.cboModality.Text, "--")(1)
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
        
        If .cbo检查技师.ListIndex = 0 Then
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
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl, i As Integer
    
    If CommandBar.Parent Is Nothing Then Exit Sub
    Select Case CommandBar.Parent.ID
        Case conMenu_Manage_LocateType
            With CommandBar.Controls
                If .Count = 0 Then '动态子菜单,扩1位
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10#, "标识号(&1)"): objControl.Category = "Main": objControl.Checked = True
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 1, "就诊卡(&2)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 2, "姓名(&3)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 3, "单据号(&4)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 4, "检查号(&5)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 5, "身份证(&6)"): objControl.Category = "Main"
                    Set objControl = .Add(xtpControlButton, conMenu_Manage_LocateType * 10# + 6, "ＩＣ卡(&7)"): objControl.Category = "Main"
                End If
            End With
        Case conMenu_View_Filter * 10#
            With CommandBar.Controls
                If .Count = 0 Then
                    '先添加全部科室
                    Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100#, "全部科室")
                    objControl.Category = "Main"
                    objControl.DescriptionText = 0
                    If mblnAllDepts = True Then objControl.Checked = True
                    
                    '再添加每一个具体科室
                    For i = 0 To UBound(Split(mstrCanUse科室, "|"))  'mstrCanUse科室=id_编码-名称|id_编码-名称
                        Set objControl = .Add(xtpControlButton, conMenu_View_Filter * 100# + i + 1, Split(Split(mstrCanUse科室, "|")(i), "_")(1) & "(&" & i & ")")
                        objControl.Category = "Main"
                        objControl.DescriptionText = Split(Split(mstrCanUse科室, "|")(i), "_")(0)
                        If mblnAllDepts = False And mlngCur科室ID = objControl.DescriptionText Then objControl.Checked = True
                    Next
                End If
            End With
        Case Else
            Select Case Me.TabWindow.Selected.Tag
                Case "住院医嘱"
                    mobjInAdvice.zlPopupCommandBars CommandBar
                Case "门诊医嘱" '门诊
                    mobjOutAdvice.zlPopupCommandBars CommandBar
                Case "申请费用"
                    mobjExpense.zlPopupCommandBars CommandBar
            End Select
    End Select
End Sub


Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnNoRecord As Boolean, intState As Integer, blnCancel As Boolean
    If Not mblnInitOk Then Exit Sub

    blnNoRecord = Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0
    control.Style = xtpButtonIconAndCaption
    If Not blnNoRecord Then
        intState = Val(vsList.TextMatrix(vsList.Row, GetCN("检查状态")))
        blnCancel = vsList.TextMatrix(vsList.Row, GetCN("检查过程")) = "已拒绝"
    End If

    Select Case control.ID
        Case conMenu_Manage_LocateType
            control.Caption = "按" & mstrLocateType & "定位(&G)"
            control.Enabled = Not blnNoRecord
        Case conMenu_Manage_LocateType * 10# To conMenu_Manage_LocateType * 10# + 6
            control.Checked = (InStr(control.Caption, mstrLocateType) > 0)
        Case conMenu_Manage_LocateValue
            control.Enabled = Not blnNoRecord
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
        Case conMenu_Manage_CopyCheck '再次登记
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
                control.Enabled = intState <= 1 And Not blnCancel
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
                control.Enabled = intState <= 3 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Receive   '检查报到(&L)
            If InStr(mstrPrivs, "检查报到") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState <= 1 And Not blnCancel
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Logout   '取消报到(&D)
            If blnNoRecord Then
                control.Enabled = False
            ElseIf control.Parent.Type = xtpControlPopup Then
                If InStr(mstrPrivs, "取消报到") <= 0 Then
                    control.Visible = False
                Else
                    control.Visible = True
                    control.ToolTipText = "取消报到"
                    control.Caption = "取消报到(&D)"
                    control.Enabled = (intState = 2 Or intState = 3)
                End If
            Else ' 工具栏中的用取消检查代替取消登记,同一按键完成取消登记和取消检查功能
                control.Visible = IIf(intState <= 1, InStr(mstrPrivs, "检查登记") > 0, InStr(mstrPrivs, "取消报到") > 0)
                control.Enabled = (intState = 2 Or intState = 3) Or (intState <= 1 And Not blnCancel) '被拒绝的不能被再次拒绝
                control.ToolTipText = IIf(intState <= 1, "取消登记", "取消报到")
                control.Caption = "取消"
            End If
        Case conMenu_Manage_Transfer   '关联影像(&C)
            If InStr(mstrPrivs, "图像关联") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '在2---5之间可用
            End If
        Case conMenu_Manage_Cancel   '取消关联(&B)
            If InStr(mstrPrivs, "图像关联") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_First, conMenu_Manage_Second, conMenu_Manage_Quality
            If InStr(mstrPrivs, "影像质控") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState <= 5 Then
                control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_TechDoctorExecute   '技师执行,报到后，有图像前可以执行
            If InStr(mstrPrivs, "检查登记") <= 0 Then
                control.Visible = False
            ElseIf intState >= 2 And intState < 3 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ReportRelease       '报告发放,报到后，完成后都可以执行
            If intState >= 2 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If

            '修改报告发放按钮的标题
            If Not blnNoRecord Then
                If vsList.TextMatrix(vsList.Row, GetCN("报告发放")) = "已发放" Then
                    control.Caption = "收回"
                    control.ToolTipText = "收回已经发放的报告"
                Else
                    control.Caption = "发放"
                    control.ToolTipText = "发放报告"
                End If
            End If
        Case conMenu_Manage_Review  '备注
            If InStr(mstrPrivs, "随访") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord And intState > 1 Then
                control.Enabled = True
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Result, conMenu_Manage_Negative, conMenu_Manage_Positive '结果阴阳性(&X)
            If (InStr(GetInsidePrivs(p诊疗报告管理), "报告书写") <= 0 And InStr(GetInsidePrivs(p诊疗报告管理), "报告修订") <= 0) Or _
                mblnIgnoreResult Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '在2---5之间可用
            End If
        Case conMenu_Manage_GChannel, conMenu_Manage_GChannelOk, conMenu_Manage_GChannelCancel '绿色通道标记/取消
            If InStr(mstrPrivs, "绿色通道") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = intState >= 2 And intState <= 5 '在2---5之间可用
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
                control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("报告人")) = ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_Complete   '检查完成(&E)
            If InStr(mstrPrivs, "检查完成") <= 0 Then
                control.Visible = False
            Else
                control.Enabled = (intState = 4 Or intState = 5)
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
            If blnNoRecord Then control.Enabled = False: Exit Sub
            If mblnIsHistory = True Then
                control.Enabled = mstrHStudyUID <> ""
            Else
                control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> ""
            End If
            If control.Parent.Type <> xtpControlPopup Then control.Visible = control.Enabled
        Case conMenu_Img_3D     '三维重建
            If InStr(mstrPrivs, "三维重建操作") <> 0 And mblnUse3D = True Then
                control.Visible = True
            Else
                control.Visible = False
            End If
            If control.Visible = True Then
                If blnNoRecord Then control.Enabled = False: Exit Sub
                If control.Parent.Type <> xtpControlPopup Then
                    control.Visible = vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> ""
                    control.Enabled = control.Visible
                Else
                    control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> ""
                End If
            End If
        Case conMenu_Img_Delete '清除图像
            If InStr(mstrPrivs, "清除图像") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> ""
            Else
                control.Enabled = False
            End If
        Case conMenu_Img_Query ',获取图像
            If InStr(mstrPrivs, "清除图像") <= 0 Then
                control.Visible = False
            ElseIf Not blnNoRecord Then
                control.Enabled = intState > 1
            Else
                control.Enabled = False
            End If
        Case conMenu_Manage_ChangeDevice    '更改影像设备
            If blnNoRecord = True Then
                control.Enabled = False
            Else
                If UCase(vsList.TextMatrix(vsList.Row, GetCN("影像类别"))) = "CR" Or _
                    UCase(vsList.TextMatrix(vsList.Row, GetCN("影像类别"))) = "DR" Or _
                    UCase(vsList.TextMatrix(vsList.Row, GetCN("影像类别"))) = "DX" Or _
                    UCase(vsList.TextMatrix(vsList.Row, GetCN("影像类别"))) = "RF" Then
                    control.Enabled = True
                Else
                    control.Enabled = False
                End If
            End If
        Case conMenu_Manage_RelatingPatiet  '关联病人
            If InStr(mstrPrivs, "关联病人") <= 0 Or mblnRelatingPatient = False Then
                control.Visible = False
            ElseIf blnNoRecord Or intState < 2 Then
                control.Enabled = False
            Else
                control.Enabled = True
            End If
        Case conMenu_File_PrintSet     '打印设置(&S)
        Case conMenu_File_Preview, conMenu_File_Print '报告预览(&V) 报告打印(&P)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_Excel         '清单打印(&L)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_BatPrint    ' 批量打印(&B)
            control.Enabled = Not blnNoRecord
        Case conMenu_File_Parameter     '参数设置(&O)
        Case conMenu_ReportPopup, conMenu_ReportPopup * 100# + 1 To conMenu_ReportPopup * 100# + 99 '报表
        Case conMenu_FilePopup, conMenu_ManagePopup, conMenu_ViewPopup, conMenu_HelpPopup
        Case conMenu_Help_Help, conMenu_Help_About  '帮助
        Case conMenu_Help_Web, conMenu_Help_Web_Forum, conMenu_Help_Web_Home, conMenu_Help_Web_Mail '帮助WEB
        Case conMenu_File_Exit      '退出
        Case conMenu_View_ToolBar   '工具栏
        Case conMenu_Cap_DevSet     '影像设备设置
        Case conMenu_Manage_Change_In   '隐藏列表
        Case Else
            If blnNoRecord Then control.Enabled = False: Exit Sub
            Select Case TabWindow.Selected.Tag
                Case "PACS报告"
                    mfrmPacsReport.zlUpdateCommandBars control
                Case "病历报告"
                    mobjReport.zlUpdateCommandBars control
                Case "申请费用"
                    mobjExpense.zlUpdateCommandBars control
                Case "住院医嘱"
                    mobjInAdvice.zlUpdateCommandBars control
                Case "门诊医嘱"
                    mobjOutAdvice.zlUpdateCommandBars control
                Case "住院病历"
                    mobjInEPRs.zlUpdateCommandBars control
                Case "门诊病历"
                    mobjOutEPRs.zlUpdateCommandBars control
            End Select

            If Not blnNoRecord Then
                '删除只能在已报告和进行中可用
                If control.ID = conMenu_Edit_Delete And Val(vsList.TextMatrix(vsList.Row, GetCN("检查状态"))) >= 4 Then
                    control.Enabled = False
                End If
                '当前查看的是历次记录则菜单均不可用
                If cboTimes.ListIndex <> -1 Then
                    If vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) <> cboTimes.ItemData(cboTimes.ListIndex) Then
                        If control.ID = conMenu_Edit_Copy Or control.ID = conMenu_File_ExportToXML Or control.ID = conMenu_Tool_Search Then
                            '这几个菜单不控制
                        Else
                            control.Enabled = False
                        End If
                    End If
                End If
                '已完成除查阅,以及医嘱中报告查看打印，观片菜单外均不可用
                If Val(vsList.TextMatrix(vsList.Row, GetCN("检查状态"))) = 6 Then
                    Select Case control.ID
                        Case conMenu_Edit_MarkMap, conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3
                            control.Enabled = True
                        Case conMenu_Edit_Copy, conMenu_File_ExportToXML, conMenu_Tool_Search, conMenu_File_Open, conMenu_EditPopup
                            '这几个菜单不控制
                        Case Else
                            control.Enabled = False
                    End Select
                End If
            End If
    End Select
End Sub

Private Sub InitMvar(intType As Integer)
'功能:初始化模块级变量,窗体加载、科室改变时调用
'参数：intType---0从菜单或者FormLoad触发科室改变，刷新病人过滤开始时间；intType---1从病人列表触发科室改变，不用再刷新过滤开始时间

    On Error GoTo err
    
    '读取跟科室相关的流程管理参数
    mblnIgnoreResult = GetDeptPara(mlngCur科室ID, "忽略结果阴阳性", 0) = "1" '        '忽略结果阴阳性
    mblnFinishCommit = GetDeptPara(mlngCur科室ID, "无报告完成后直接完成", 0) = "1" '  '无报告完成后直接完成
    mblnReportWithImage = GetDeptPara(mlngCur科室ID, "有图像才能写报告", 0) = "1" '   '有图像才能写报告
    mblnReportWithResult = GetDeptPara(mlngCur科室ID, "无影像诊断为阴性", 0) = "1" '  '无影像诊断为阴性
    mblnLocalizerBackward = GetDeptPara(mlngCur科室ID, "定位片后置", 0) = "1" '       '定位片后置
    mblnCompleteCommit = GetDeptPara(mlngCur科室ID, "审核后直接完成", 0) = "1" '      '审核后直接完成
    mBeforeDays = Val(GetDeptPara(mlngCur科室ID, "默认过滤天数", 2)) '                   '默认过滤天数
    If mBeforeDays > 15 Or mBeforeDays <= 0 Then
        mBeforeDays = 2
    End If
    mblnTechReptSame = GetDeptPara(mlngCur科室ID, "只能填写自己检查的报告", 0) = "1"  '只能填写自己检查的报告
    mblnPacsReport = GetDeptPara(mlngCur科室ID, "报告编辑器", 0) = "1" '              '报告编辑器
    mintResultInput = Val(GetDeptPara(mlngCur科室ID, "提示阴阳性", 1))    '              '提示阴阳性
    mblnPrintCommit = GetDeptPara(mlngCur科室ID, "打印后直接完成", 0) = "1" '         '打印后直接完成
    mblnCanPrint = GetDeptPara(mlngCur科室ID, "平诊需审核才能打报告") = "1"           '平诊需要审核才能打印 =true
    If InStr(mstrPrivs, "排队叫号") > 0 Then                                        '有权限使用才根据参数启用
        mblnUseQueue = GetDeptPara(mlngCur科室ID, "启动排队叫号", 0) = "1" '          '默认不启用排队叫号
    End If
    mblnRelatingPatient = GetDeptPara(mlngCur科室ID, "启动关联病人", 0) = "1"       '是否使用关联病人
    
    mlngRefreshInterval = Val(GetDeptPara(mlngCur科室ID, "自动刷新间隔", 0)) '      '自动刷新间隔,默认不自动刷新
    If mlngRefreshInterval > 0 Then
        If mlngRefreshInterval > 65 Then mlngRefreshInterval = 65
        TimerRefresh.Interval = mlngRefreshInterval * 1000
        TimerRefresh.Enabled = True
    Else
        TimerRefresh.Enabled = False
    End If

    If intType = 0 Then    '从菜单或者FormLoad触发科室改变，刷新病人过滤开始时间
        SQLCondition.开始时间 = CDate(Format(zldatabase.Currentdate - mBeforeDays, "yyyy-mm-dd 00:00"))
        mblnMoved = MovedByDate(SQLCondition.开始时间)
    End If
    
    '初始化排队叫号的队列名称列表
    Dim iCount As Integer, rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    iCount = 1
    gstrSQL = "Select 执行间,检查设备 From 医技执行房间 where 科室id=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取执行间名称", mlngCur科室ID)
    If rsTemp.EOF <> True Then
        ReDim mAstr队列名称(rsTemp.RecordCount) As String
        While rsTemp.EOF = False
            'mAstr队列名称(iCount) = Split(mstrCur科室, "-")(1) & Nvl(rsTemp!执行间)
            mAstr队列名称(iCount) = mlngCur科室ID & ":" & Nvl(rsTemp!执行间)
            iCount = iCount + 1
            rsTemp.MoveNext
        Wend
    Else
        ReDim mAstr队列名称(0) As String
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
    Call SaveErrLog
End Sub

Private Sub Menu_Manage_PACS报告()
    Dim i As Integer, lngAdviceID As Long, lngSendNO As Long
    
    With vsList
        lngAdviceID = .TextMatrix(.Row, GetCN("医嘱ID"))
        lngSendNO = .TextMatrix(.Row, GetCN("发送号"))
    End With
    
    If Not mfrmPacsReportDock Is Nothing Then
        '先判断当前窗体是否是需要打开的窗体，如果不是，则查找窗口数组
        If lngAdviceID = mfrmPacsReportDock.mlngAdviceID Then
            '当前mfrmPacsReportDock指向的窗体，就是需要打开的窗体
            mfrmPacsReportDock.WindowState = 0  'normal
            mfrmPacsReportDock.ZOrder
            Exit Sub
        End If
    End If
    
    '查找窗口数组,找到需要打开的窗体，则通过Zorder把窗体显示到最前面
    If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
        For i = 1 To UBound(mobjPacsReportArry)
            If lngAdviceID = mobjPacsReportArry(i).mlngAdviceID Then
                Set mfrmPacsReportDock = mobjPacsReportArry(i)
                mfrmPacsReportDock.WindowState = 0  'normal
                mfrmPacsReportDock.ZOrder
                Exit Sub
            End If
        Next i
    End If
    
    '没有找到需要打开的窗体，且打开新窗体,并记录当前窗体
    Set mfrmPacsReportDock = New frmReport
    Set mfrmPacsReportDock.pobjPacsCore = mobjPacsCore
    
    Call mfrmPacsReportDock.InitReportWindow(mlngCur科室ID, mlngModul, mstrPrivs, True)
    
    mfrmPacsReportDock.zlEditReport lngAdviceID, lngSendNO, Me, vsList.TextMatrix(vsList.Row, GetCN("转出")) = 1, vsList.TextMatrix(vsList.Row, GetCN("姓名"))
    
    If SafeArrayGetDim(mobjPacsReportArry) = 0 Then
        ReDim mobjPacsReportArry(1) As frmReport
    Else
        ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) + 1) As frmReport
    End If
    Set mobjPacsReportArry(UBound(mobjPacsReportArry)) = mfrmPacsReportDock
End Sub

Private Sub cmdInfo_Click()
    On Error GoTo errHandle
    frmDegreeCard.ShowMe Val(vsList.TextMatrix(vsList.Row, GetCN("病人ID"))), Val(vsList.TextMatrix(vsList.Row, GetCN("主页ID")))
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picList.hWnd
    ElseIf Item.ID = 2 Then
        Item.Handle = PicWindow.hWnd
    End If
End Sub
Private Sub Form_Load()
    mstrPrivs = gstrPrivs           '权限
    mlngModul = glngModul           '模块号
    mlngCur科室ID = 0
    mstrCur科室 = ""
    mstrCanUse科室 = ""
    mblnAllDepts = False            '默认不选择全部科室
    mlngSortCol = 0
    mintSortOrder = 0
    
    mblnInitOk = False  '初始数据,初始化完成之前不进行数据的提取
    mblnvsRefresh = False
    Call InitLocalPars '本地注册表参数
    If Not InitDepts Then Unload Me: Exit Sub '初始化医技科室
    
    ReDim gConnectedShardDir(0) As String   '初始化共享目录连接串
    
    Call InitMvar(0) '初始化模块级变量
    '初始子窗体
    Set mfrmPACSImg = New frmPACSImg
    
    Set mfrmPacsReport = New frmReport  'PACS报告
    
    
    Set mobjReport = New zlRichEPR.cDockReport
    Set mobjPacsCore = New zl9PacsCore.clsViewer
        mobjReport.PacsCore = mobjPacsCore
    Set mobjExpense = New zlCISKernel.clsDockExpense
    Set mobjInAdvice = New zlCISKernel.clsDockInAdvices
    Set mobjOutAdvice = New zlCISKernel.clsDockOutAdvices
    Set mobjInEPRs = New zlRichEPR.cDockInEPRs
    Set mobjOutEPRs = New zlRichEPR.cDockOutEPRs
    
    If mblnUseQueue And InStr(GetPrivFunc(glngSys, 1160), "基本") > 0 Then
        Set mobjQueue = New zlQueueManage.clsQueueManage      '排队叫号
    
        Call mobjQueue.zlInitVar(gcnOracle, glngSys, mintCur业务类型, 1, "")
    Else
        Set mobjQueue = Nothing
    End If
    
    Call ReadStudyListColor(mlngCur科室ID)
    Call InitFilterCmd
    Call InitCommandBars
    Call InitSubForm
    Call InitFaceScheme
    Call InitList
    
    Set mfrmPacsReport.pobjPacsCore = mobjPacsCore
    Call mfrmPacsReport.InitReportWindow(mlngCur科室ID, mlngModul, mstrPrivs, True)

    Set mfrmPACSImg.pobjPacsCore = mobjPacsCore
    
    '去掉PACS报告窗体的控制框
    FormSetCaption mfrmPacsReport, False, False
    mblnInitOk = True '初始化完成
    
    Call RestoreWinState(Me, App.ProductName)
    '不能被restorewinstate冲掉，所以写在其后
    Call RefreshList
    
    ClearCacheFolder App.Path & "\TmpImage\"    '若临时目录满了，则清空该目录
    Me.stbThis.Panels(3).Text = "报告医生：" & UserInfo.姓名
    ReDim mobjPacsReportArry(0) As frmReport
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strTemp As String
    Dim i As Integer
    
    On Error Resume Next
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "门诊病人", IIf(mblncmd门诊, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "住院病人", IIf(mblncmd住院, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "外诊病人", IIf(mblncmd外诊, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "体检病人", IIf(mblncmd体检, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用已缴", IIf(mblncmd已缴, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用未缴", IIf(mblncmd未缴, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记病人", IIf(mblncmd登记, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报到病人", IIf(mblncmd报到, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查病人", IIf(mblncmd检查, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报告病人", IIf(mblncmd报告, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "审核病人", IIf(mblncmd审核, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "完成病人", IIf(mblncmd完成, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "过滤方式", mstrCurFindtype
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "定位方式", mstrLocateType
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "本次住院", IIf(mblncmd本次, 1, 0)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序列", mlngSortCol
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序方向", mintSortOrder
    
    If UBound(mblncmd影像类别) >= 0 Then
        strTemp = mblncmd影像类别(0)
    End If
    For i = 1 To UBound(mblncmd影像类别)
        strTemp = strTemp & "," & mblncmd影像类别(i)
    Next i
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "影像类别过滤", strTemp
    
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    Call SaveWinState(Me, App.ProductName)
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, mstrCol)
    '判断嵌入式报告编辑器中的报告是否没有保存
    If mblnPacsReport = True Then    '使用PACS报告编辑器
        Call mfrmPacsReport.PromptModify
    End If
    
    Unload mfrmPACSImg
    Unload mfrmPacsReport
    Unload mobjReport.zlGetForm
    Unload mobjExpense.zlGetForm
    Unload mobjInAdvice.zlGetForm
    Unload mobjOutAdvice.zlGetForm
    Unload mobjInEPRs.zlGetForm
    Unload mobjOutEPRs.zlGetForm
    If Not mobjPacsCore Is Nothing Then mobjPacsCore.Closefrom
    
    Set mobjIDCard = Nothing
    Set mfrmPacsReport = Nothing
    Set mobjReport = Nothing
    Set mobjExpense = Nothing
    Set mobjInAdvice = Nothing
    Set mobjOutAdvice = Nothing
    Set mobjInEPRs = Nothing
    Set mobjOutEPRs = Nothing
    Set mobjPacsCore = Nothing
    
     On Error Resume Next
     
    '关闭排队叫号窗体
    If Not mobjQueue Is Nothing Then
        Unload mobjQueue.zlGetForm
        Set mobjQueue = Nothing
    End If
End Sub
Private Function GetCN(ByVal Col As String) As Integer
Dim arrCol As Variant, i As Integer
    If mstrCol = "" Then mstrCol = ConstrCol
    arrCol = Split(mstrCol, "|")
    For i = 0 To UBound(arrCol)
        'If InStr(arrCol(i), Col) > 0 Then GetCN = i: Exit Function
        If Split(arrCol(i), ";")(0) = Col Then GetCN = i: Exit Function
    Next
    GetCN = 0
End Function
Private Function GetCW(ByVal Col As String) As Long
    Dim arrCol As Variant, i As Integer
    arrCol = Split(mstrCol, "|")
    For i = 0 To UBound(arrCol)
        'If InStr(arrCol(i), Col) > 0 Then GetCW = Split(arrCol(i), ";")(1): Exit Function
        If Split(arrCol(i), ";")(0) = Col Then GetCW = Split(arrCol(i), ";")(1): Exit Function
    Next
    GetCW = 0
End Function
Private Sub InitLocalPars()
'初始化临时本地参数，以个人设置，注册表参数为主,窗体加载，本地设置等调用
    Dim strTemp As String
    Dim strTempArry() As String
    Dim i As Integer
    
    On Error GoTo err
    mblncmd门诊 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "门诊病人", 1))
    mblncmd住院 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "住院病人", 1))
    mblncmd外诊 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "外诊病人", 1))
    mblncmd体检 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "体检病人", 1))
    mblncmd已缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用已缴", 0))
    mblncmd未缴 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "费用未缴", 0))
    mblncmd登记 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "登记病人", 1))
    mblncmd报到 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报到病人", 1))
    mblncmd检查 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "检查病人", 1))
    mblncmd报告 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "报告病人", 1))
    mblncmd审核 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "审核病人", 1))
    mblncmd完成 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "完成病人", 1))
    mstrCurFindtype = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "过滤方式", "检查号")
    mstrLocateType = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "定位方式", "检查号")
    mblncmd本次 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "本次住院", "0"))
    mlngSortCol = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序列", 0))
    mintSortOrder = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "排序方向", 0))
    
    strTemp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "影像类别过滤", "")
    ReDim strTempArry(0)
    ReDim mblncmd影像类别(0)
    On Error Resume Next
    strTempArry = Split(strTemp, ",")
    If UBound(strTempArry) >= 0 Then ReDim mblncmd影像类别(UBound(strTempArry))
    For i = 0 To UBound(strTempArry)
        mblncmd影像类别(i) = IIf(UCase(strTempArry(i)) = "TRUE", True, False)
    Next i
    
    On Error GoTo err
    mstrFirstTab = zldatabase.GetPara("工作首页", glngSys, mlngModul, "") '为空表示不使用定制工作首页功能
    mbln直接检查 = (Val(GetDeptPara(mlngCur科室ID, "登记后直接检查", 0)) = 1)
    mblnOpenReport = (Val(zldatabase.GetPara("开始检查自动打开报告", glngSys, mlngModul, 0)) = 1)
    mblnShowImgAtReport = (Val(zldatabase.GetPara("报告时观片", glngSys, mlngModul, 0)) = 1)
    mblnNoShowCancel = (Val(zldatabase.GetPara("不显示被取消的登记", glngSys, mlngModul, 0)) = 1)
    mblnPatTrack = (Val(zldatabase.GetPara("病人跟踪", glngSys, mlngModul, 0)) = 1)
    mstrRoom = zldatabase.GetPara("执行间范围", glngSys, mlngModul, "")
    If mstrRoom <> "" Then mstrRoom = "'," & Replace(mstrRoom, "|", ",") & ",'"
    
    '读取三维重建参数
    mblnUse3D = Val(zldatabase.GetPara("启用三维重建", glngSys, mlngModul, 0))
    mstr3DExeDir = zldatabase.GetPara("3D程序路径", glngSys, mlngModul, "")
    mstr3DPara = zldatabase.GetPara("3D参数", glngSys, mlngModul, "")
    mstr3DFunctions = zldatabase.GetPara("3D功能", glngSys, mlngModul, "")

    With SQLCondition '------------------------ '过滤条件初始
        '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
        .时间类型 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "过滤时间类型", 1))
        .单据号 = ""
        .门诊号 = 0
        .住院号 = 0
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
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitDepts() As Boolean
'功能：初始化住院临床科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str科室IDs As String, str来源 As String
    
    On Error GoTo errH
    
    str来源 = "1,2,3"
    If InStr(mstrPrivs, "所有科室") > 0 Then
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B " & _
            " Where B.部门ID = A.ID " & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')> 0 And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    Else
        strSQL = _
            " Select Distinct A.ID,A.编码,A.名称" & _
            " From 部门表 A,部门性质说明 B,部门人员 C " & _
            " Where B.部门ID = A.ID And A.ID=C.部门ID And C.人员ID=" & UserInfo.ID & _
            " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) " & _
            " and (A.站点='" & gstrNodeNo & "' Or A.站点 is Null ) " & _
            " And instr([1],','||B.服务对象||',')>0  And B.工作性质 IN('检查')" & _
            " Order by A.编码"
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CStr("," & str来源 & ","))
    
    If rsTmp.EOF Then
        MsgBoxD Me, "没有发现医技科室信息,请先到部门管理中设置。", vbInformation, gstrSysName
        Exit Function
    Else
        str科室IDs = GetUser科室IDs
        Do Until rsTmp.EOF
            mstrCanUse科室 = mstrCanUse科室 & "|" & rsTmp!ID & "_" & rsTmp!编码 & "-" & rsTmp!名称
            mstrCanUse科室IDs = mstrCanUse科室IDs & "," & rsTmp!ID
            If rsTmp!ID = UserInfo.部门ID Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '提取默认科室
            If InStr("," & str科室IDs & ",", "," & rsTmp!ID & ",") > 0 And mlngCur科室ID = 0 Then mlngCur科室ID = rsTmp!ID: mstrCur科室 = rsTmp!编码 & "-" & rsTmp!名称 '没有默认科室取第一个所属检查科室
            rsTmp.MoveNext
        Loop
        mstrCanUse科室 = Mid(mstrCanUse科室, 2)
        mstrCanUse科室IDs = Mid(mstrCanUse科室IDs, 2)
        
        If InStr(mstrPrivs, "所有科室") > 0 And mlngCur科室ID = 0 Then
            mlngCur科室ID = Split(Split(mstrCanUse科室, "|")(0), "_")(0)
            mstrCur科室 = Split(Split(mstrCanUse科室, "|")(0), "_")(1)
        End If
        
        If mlngCur科室ID = 0 And InStr(mstrPrivs, "所有科室") <= 0 Then '没有所有科室操作权限,而且操作者科室不属于检查类科室
            MsgBoxD Me, "没有发现你所属科室,不能使用医技工作站。", vbInformation, gstrSysName
            Exit Function
        End If
        InitDepts = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitFaceScheme()
    '初始界面布局
    Dim Pane1 As Pane, Pane2 As Pane
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    Set Pane1 = dkpMain.CreatePane(1, 240, 250, DockLeftOf, Nothing)
    Pane1.Title = "检查列表"
    Pane1.Handle = picList.hWnd
    Pane1.Options = PaneNoCloseable Or PaneNoFloatable
    
    Set Pane2 = dkpMain.CreatePane(2, 700, 250, DockRightOf, Nothing)
    Pane2.Title = "子窗体"
    Pane2.Handle = PicWindow.hWnd
    Pane2.Options = PaneNoCaption Or PaneNoCloseable
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
End Sub
Private Sub InitFilterCmd()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl, cbrPopControl As CommandBarControl
    Dim objPopbar As CommandBarPopup, objCusControl As CommandBarControlCustom
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strTemp As String
    Dim i As Integer

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
        Set objControl = .Add(xtpControlButton, ID_门诊, "门诊")
            objControl.ToolTipText = "显示门诊病人"
        Set objControl = .Add(xtpControlButton, ID_住院, "住院")
            objControl.ToolTipText = "显示住院病人"
        Set objControl = .Add(xtpControlButton, ID_外诊, "外诊")
            objControl.ToolTipText = "显示外诊病人"
        Set objControl = .Add(xtpControlButton, ID_体检, "体检")
            objControl.ToolTipText = "显示体检病人"
        Set objControl = .Add(xtpControlButtonPopup, ID_费用, " 费  用")
            objControl.ToolTipText = "显示费用已缴/未缴病人"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_未缴, "未缴")
            cbrPopControl.ToolTipText = "显示费用未缴病人"
        Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_已缴, "已缴")
            cbrPopControl.ToolTipText = "显示费用已缴病人"
        
        '添加所有影像类别
        Set objControl = .Add(xtpControlButtonPopup, ID_影像类别, "影像类别")
        objControl.ToolTipText = "显示影像类别"
        strSQL = "select 编码,名称 from 影像检查类别"
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "影像检查类别")
        i = 1
        mintcmd影像类别 = 0
        strTemp = ""
        ReDim Preserve mblncmd影像类别(rsTemp.RecordCount - 1)
        While rsTemp.EOF = False
            Set cbrPopControl = objControl.CommandBar.Controls.Add(xtpControlButton, ID_影像类别 + i, rsTemp("名称"))
            cbrPopControl.DescriptionText = rsTemp("编码")
            cbrPopControl.Style = xtpButtonIconAndCaption
            cbrPopControl.Checked = mblncmd影像类别(i - 1)
            cbrPopControl.CloseSubMenuOnClick = False
            If mblncmd影像类别(i - 1) = True Then
                mintcmd影像类别 = mintcmd影像类别 + 1
                strTemp = IIf(strTemp = "", cbrPopControl.Caption, strTemp & "," & cbrPopControl.Caption)
            End If
            rsTemp.MoveNext
            i = i + 1
        Wend
        If strTemp <> "" Then objControl.Caption = strTemp
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("状态", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_登记, "登记")
            objControl.ToolTipText = "显示已登记病人"
        Set objControl = .Add(xtpControlButton, ID_报到, "报到")
            objControl.ToolTipText = "显示已报到病人"
        Set objControl = .Add(xtpControlButton, ID_检查, "检查")
            objControl.ToolTipText = "显示已检查病人"
        Set objControl = .Add(xtpControlButton, ID_报告, "报告")
            objControl.ToolTipText = "显示已报告病人"
        Set objControl = .Add(xtpControlButton, ID_审核, "审核")
            objControl.ToolTipText = "显示已审核病人"
        Set objControl = .Add(xtpControlButton, ID_完成, "完成")
            objControl.ToolTipText = "显示已完成病人"
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set objBar = cbrdock.Add("过滤", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    Set objPopbar = objBar.Controls.Add(xtpControlPopup, ID_查找方式, "查找方式")
        objPopbar.ID = ID_查找方式
        objPopbar.flags = xtpFlagRightAlign
        
    Set objCusControl = objBar.Controls.Add(xtpControlCustom, ID_查找值, "查找值")
        objCusControl.Handle = txtFilter.hWnd
        objCusControl.flags = xtpFlagRightAlign
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_开始查找, "开始查找")
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = conMenu_View_Filter
        
    Set objControl = objBar.Controls.Add(xtpControlButton, ID_本次住院, "本次")
    objControl.ToolTipText = "只显示本次住院检查记录"
    objControl.Style = xtpButtonIconAndCaption
    objControl.IconId = conMenu_View_Filter
    
    With cbrdock.KeyBindings
        .Add FCONTROL, vbKey0, ID_门诊
        .Add FCONTROL, vbKey1, ID_住院
        .Add FCONTROL, vbKey2, ID_外诊
        .Add FCONTROL, vbKey3, ID_体检
        
        .Add FCONTROL, vbKey4, ID_登记
        .Add FCONTROL, vbKey5, ID_报到
        .Add FCONTROL, vbKey6, ID_检查
        .Add FCONTROL, vbKey7, ID_报告
        .Add FCONTROL, vbKey8, ID_审核
        .Add FCONTROL, vbKey9, ID_完成
        .Add FCONTROL, Asc("G"), ID_开始查找
    End With
    cbrdock.RecalcLayout
End Sub
Private Sub InitCommandBars()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim i As Integer
    Dim i3DFunc As Integer

    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
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
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    

'菜单定义
'Begin------------------------文件菜单--------------------------------------默认可见
    Me.cbrMain.ActiveMenuBar.Title = "菜单"
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)"): cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "报告预览(&V)"): cbrControl.IconId = 102
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "报告打印(&P)"): cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "批量打印(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "清单打印(&L)"): cbrControl.BeginGroup = True: cbrControl.IconId = 103
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置(&O)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_Cap_DevSet, "影像设备设置(&D)"):: cbrControl.IconId = 181
        Set cbrControl = .Add(xtpControlButton, conMenu_File_SendImg, "发送图像(&T)"): cbrControl.IconId = 3061
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Change_In, "隐藏列表")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"):: cbrControl.IconId = 191: cbrControl.BeginGroup = True
    End With


'Begin----------------------检查菜单--------------------------------------默认可见
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ManagePopup, "检查(&S)", -1, False)
    cbrMenuBar.ID = conMenu_ManagePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_RequestPrint, "打印申请单据(&J)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "检查登记(&I)"): cbrControl.IconId = 211: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_CopyCheck, "复制登记(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "取消登记(&R)"): cbrControl.IconId = 742
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReGet, "召回取消(&G)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ThingModi, "修改信息(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "检查报到(&L)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "取消报到(&D)"): cbrControl.IconId = 743
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "影像观片(&S)"): cbrControl.IconId = 8111:  cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "观片对比(&E)"): cbrControl.IconId = 8112
        
        '如果启用三维重建功能，则创建对应菜单
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "三维重建"): cbrControl.ID = conMenu_Img_3D
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "容积重建")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "虚拟内窥镜")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "表面重建")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "灌注成像")
                            End Select
                        End If
                    Next i
                End If
        End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Delete, "影像删除(&K)"): cbrControl.IconId = 8113
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Query, "Q/R获取图象(&Q)"): cbrControl.IconId = 8111
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer, "关联影像(&C)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 505: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Cancel, "取消关联(&B)"): cbrControl.IconId = 506
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Review, "备注(&R)"):  cbrControl.BeginGroup = True: cbrControl.IconId = 232
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_TechDoctorExecute, "技师执行"): cbrControl.IconId = 807
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportRelease, "发放"): cbrControl.ToolTipText = "报告发放": cbrControl.IconId = 3013
        
        '忽略结果阴阳性，则不显示结果菜单
        If mblnIgnoreResult = False Then
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "检查结果(&X)"): cbrControl.ID = conMenu_Manage_Result
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "阳性(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "阴性(&Y)"): cbrPopControl.IconId = 3507
        End If
        
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "影像质量(&Y)"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "甲级(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "乙级(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_GChannel, "绿色通道(&G)"): cbrControl.ID = conMenu_Manage_GChannel
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelOk, "标记(&J)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_GChannelCancel, "取消(&Y)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "无报告完成(&F)"): cbrControl.IconId = 216: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ClearUp, "无报告回退(&U)"):  cbrControl.IconId = 3012
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "检查完成(&E)"): cbrControl.IconId = 225
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Undone, "取消完成(&U)"): cbrControl.IconId = 219
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ChangeDevice, "更换设备"): cbrControl.IconId = 3203
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_RelatingPatiet, "关联病人"): cbrControl.IconId = 803
    End With
    
    
'Begin----------------------查看菜单--------------------------------------
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.ID = conMenu_View_ToolBar
            With cbrControl.CommandBar.Controls '二级菜单
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False): cbrPopControl.Checked = True
                Set cbrPopControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False): cbrPopControl.Checked = True
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): cbrControl.Checked = True: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Manage_LocateType, "定位方式(&G)"): cbrControl.ID = conMenu_Manage_LocateType
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Filter * 10#, "检查科室"): cbrControl.ID = conMenu_View_Filter * 10#
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "快速过滤(&K)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&F)")
    End With


'Begin----------------------帮助菜单--------------------------------------默认可见
    Set cbrMenuBar = Me.cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题", -1, False)
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB上的中联(&E)")
            With cbrControl.CommandBar.Controls
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Forum, "中联论坛(&F)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Home, "中联主页(&H)", -1, False)
                Set cbrPopControl = .Add(xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False)
            End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zldatabase.ShowReportMenu(cbrMain, glngSys, mlngModul, mstrPrivs)
    
'----------------------快键绑定------------------------------------------
    With Me.cbrMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print '打印------------------Ctrl+P
        .Add 0, VK_F12, conMenu_File_Parameter      '参数设置--------------F12
        
        .Add 0, VK_F2, conMenu_Manage_Regist       '登记-----------------F2
        .Add 0, VK_F7, conMenu_Manage_CopyCheck    '复制登记-------------F7
        .Add 0, VK_F4, conMenu_Manage_Receive       '报到-----------------F4
        .Add 0, VK_F9, conMenu_Manage_ClearUp       '驳回报告------------F9
        .Add 0, VK_F6, conMenu_Manage_Complete         '审核报告----------F6
        
        
        .Add 0, VK_F1, conMenu_Help_Help              '帮助-------------F1
        .Add 0, VK_F5, conMenu_View_Refresh           '刷新-------------F5
        .Add FCONTROL, Asc("G"), conMenu_Manage_LocateType    '定位方式---------Ctrl+F
        .Add 0, VK_F3, conMenu_View_Filter            '过滤-------------F3
    End With
    
'---------------------设置右上角当前科室----------------------------------
        Set cbrControl = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_Filter * 10#, "检查科室")
            cbrControl.ID = conMenu_View_Filter * 10#: cbrControl.flags = xtpFlagRightAlign: cbrControl.Category = "Main"
    
        Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Manage_LocateType, "标识号(&D)")
            cbrMenuBar.ID = conMenu_Manage_LocateType
            cbrMenuBar.flags = xtpFlagRightAlign
            cbrMenuBar.Category = "Main"
            
        Set cbrCustom = cbrMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_Manage_LocateValue, "定位条件")
            cbrCustom.Handle = txtLocate.hWnd
            cbrCustom.flags = xtpFlagRightAlign
            cbrCustom.Style = xtpButtonIconAndCaption
            cbrCustom.Category = "Main"
    
'---------------------工具栏定义------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
'    cbrToolBar.EnableDocking xtpFlagStretched '+ xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.IconId = 102: cbrControl.ToolTipText = "报告预览"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印"): cbrControl.IconId = 103: cbrControl.ToolTipText = "报告打印"
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Regist, "登记"): cbrControl.BeginGroup = True: cbrControl.IconId = 211
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Receive, "报到"): cbrControl.IconId = 744
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "取消"): cbrControl.IconId = 743: cbrControl.ToolTipText = "取消报到"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Look, "观片"): cbrControl.ToolTipText = "影像观片"
        Set cbrControl = .Add(xtpControlButton, conMenu_Img_Contrast, "对比"): cbrControl.IconId = 8112: cbrControl.ToolTipText = "观片对比"
        '如果启用三维重建功能，则创建对应菜单
        If mblnUse3D = True Then
            Set cbrControl = .Add(xtpControlPopup, conMenu_Img_3D, "三维"): cbrControl.ID = conMenu_Img_3D: cbrControl.ToolTipText = "三维重建"
                If mstr3DFunctions <> "" Then
                    str3DFuncs = Split(mstr3DFunctions, ",")
                    For i = 1 To UBound(str3DFuncs)
                        i3DFunc = Val(str3DFuncs(i))
                        If i3DFunc >= 1 And i3DFunc <= 6 Then
                            Select Case i3DFunc
                                Case 1
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VA, "容积重建")
                                Case 2
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MPR, "MPR")
                                Case 3
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_MMPR, "MMPR")
                                Case 4
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_VE, "虚拟内窥镜")
                                Case 5
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_SA, "表面重建")
                                Case 6
                                    Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Img_3D_PF, "灌注成像")
                            End Select
                        End If
                    Next i
                End If
        End If
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Review, "备注"):  cbrControl.BeginGroup = True: cbrControl.IconId = 232
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_TechDoctorExecute, "技师执行"): cbrControl.IconId = 807
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_ReportRelease, "发放"): cbrControl.ToolTipText = "报告发放": cbrControl.IconId = 3013
        '忽略结果阴阳性，则不显示结果工具栏
        If mblnIgnoreResult = False Then
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Result, "结果"): cbrControl.ID = conMenu_Manage_Result: cbrControl.IconId = 3506: cbrControl.ToolTipText = "检查结果阴阳性"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Negative, "阳性(&X)"): cbrPopControl.IconId = 3506
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Positive, "阴性(&Y)"): cbrPopControl.IconId = 3507
        End If
        Set cbrControl = .Add(xtpControlPopup, conMenu_Manage_Quality, "质量"): cbrControl.ID = conMenu_Manage_Quality: cbrControl.IconId = 3061: cbrControl.ToolTipText = "影像质量"
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_First, "甲级(&J)"): cbrPopControl.IconId = 3587
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_Manage_Second, "乙级(&Y)"): cbrPopControl.IconId = 3010
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Complete, "完成"): cbrControl.IconId = 225: cbrControl.ToolTipText = "检查最终完成"
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "过滤"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
        
    End With

End Sub

Private Sub InitSubForm()
Dim i As Integer
    With TabWindow
        .RemoveAll
        .Icons = frmPubIcons.imgPublic.Icons
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
        
        .InsertItem 0, "影像记录", mfrmPACSImg.hWnd, conMenu_Img_Look
            .Item(TabWindow.ItemCount - 1).Tag = "影像图象"
        
        If GetInsidePrivs(p诊疗报告管理, True) <> "" Then
            .InsertItem 1, "影像报告", mfrmPacsReport.hWnd, conMenu_Edit_Compend
            .Item(TabWindow.ItemCount - 1).Tag = "PACS报告"
        
            .InsertItem 2, "影像报告", mobjReport.zlGetForm.hWnd, conMenu_Edit_Compend
            .Item(TabWindow.ItemCount - 1).Tag = "病历报告"
            
            .Item(1).Visible = False
        End If
        
        If GetInsidePrivs(p医嘱附费管理, True) <> "" Then
            .InsertItem 3, "费用记录", mobjExpense.zlGetForm.hWnd, conMenu_Manage_Request
            .Item(TabWindow.ItemCount - 1).Tag = "申请费用"
        End If
        
        If GetInsidePrivs(p住院医嘱下达, True) <> "" Then
            .InsertItem 4, "医嘱记录", mobjInAdvice.zlGetForm.hWnd, conMenu_Edit_NewItem
            .Item(TabWindow.ItemCount - 1).Tag = "住院医嘱"
        End If
        
        If GetInsidePrivs(p门诊医嘱下达, True) <> "" Then
            .InsertItem 5, "医嘱记录", mobjOutAdvice.zlGetForm.hWnd, conMenu_Edit_NewItem
            .Item(TabWindow.ItemCount - 1).Tag = "门诊医嘱": .Item(TabWindow.ItemCount - 1).Visible = False
        End If
        
        If GetInsidePrivs(p住院病历管理, True) <> "" Then
            .InsertItem 6, "病历记录", mobjInEPRs.zlGetForm.hWnd, conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).Tag = "住院病历"
        End If
        
        If GetInsidePrivs(p门诊病历管理, True) <> "" Then
            .InsertItem 7, "病历记录", mobjOutEPRs.zlGetForm.hWnd, conMenu_Edit_Archive
            .Item(TabWindow.ItemCount - 1).Tag = "门诊病历": .Item(TabWindow.ItemCount - 1).Visible = False
        End If
        
        '添加排队叫号页面
        If mblnUseQueue = True And Not mobjQueue Is Nothing Then
            .InsertItem 8, "排队叫号", mobjQueue.zlGetForm.hWnd, conMenu_File_PrintSingleBill
            .Item(TabWindow.ItemCount - 1).Tag = "排队叫号"
        End If
        
        If mstrFirstTab <> "" Then
            For i = 0 To .ItemCount - 1
                If InStr(.Item(i).Tag, mstrFirstTab) > 0 And .Item(i).Visible Then
                    .Item(i).Selected = True: Exit For
                End If
            Next
        End If
    End With

End Sub


Private Sub InitList()
'初始化表格
Dim C路径 As Long, C紧急 As Long, C来源 As Long, C阳性 As Long, C质量 As Long, C姓名 As Long, C检查号 As Long, C检查过程 As Long, C性别 As Long, C年龄 As Long
Dim C标识号 As Long, C医嘱内容 As Long, C部位方法 As Long, C执行间 As Long, C报到时间 As Long, C申请时间 As Long, C开嘱医生 As Long
Dim C身高 As Long, C体重 As Long, C婴儿 As Long, C登记人 As Long, C报到人 As Long, C完成人 As Long, C打印胶片 As Long, C报告操作 As Long
Dim C绿色通道 As Long, C报告打印 As Long, C报告人 As Long, C复核人 As Long, C检查技师 As Long, C采图时间 As Long, C随访描述 As Long
Dim C影像类别 As Long, C病人ID As Long, C主页ID As Long, C挂号单 As Long, C病人科室ID As Long, C医嘱ID As Long, C发送号 As Long, C检查UID As Long
Dim C检查状态 As Long, CNO As Long, C记录性质 As Long, C转出 As Long, C床号 As Long, C当前病区ID As Long, C报告发放 As Long
Dim C诊断分类 As Long, C执行科室ID As Long, C关联ID As Long, C病人科室 As Long, C就诊卡号 As Long, C单据号 As Long, C身份证号 As Long
Dim C收费 As Long
 
    If mstrCol = "" Then
        mstrCol = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsList), vsList.Name, ConstrCol)
        '判断是否修改过显示的列数，如果修改过，则读取默认值，而不是读取注册表
        If UBound(Split(mstrCol, "|")) <> UBound(Split(ConstrCol, "|")) Then
            mstrCol = ConstrCol
        End If
    End If
    With vsList
        .Clear
        .FixedRows = 1
        .Rows = 2
        .Cols = 55
        '提取列序
        C路径 = GetCN("路径")
        C紧急 = GetCN("紧急"):           C来源 = GetCN("来源"):          C阳性 = GetCN("阳性")
        C质量 = GetCN("质量"):          C姓名 = GetCN("姓名"):          C检查号 = GetCN("检查号")
        C检查过程 = GetCN("检查过程"):  C性别 = GetCN("性别"):          C年龄 = GetCN("年龄")
        C标识号 = GetCN("标识号"):      C医嘱内容 = GetCN("医嘱内容"):  C部位方法 = GetCN("部位方法")
        C执行间 = GetCN("执行间"):      C报到时间 = GetCN("报到时间"):  C申请时间 = GetCN("申请时间")
        C开嘱医生 = GetCN("开嘱医生"):   C身高 = GetCN("身高"):          C体重 = GetCN("体重")
        C婴儿 = GetCN("婴儿"):          C登记人 = GetCN("登记人"):      C报到人 = GetCN("报到人")
        C完成人 = GetCN("完成人"):      C打印胶片 = GetCN("打印胶片"):  C报告操作 = GetCN("报告操作")
        C绿色通道 = GetCN("绿色通道"):  C报告打印 = GetCN("报告打印"):  C报告人 = GetCN("报告人")
        C复核人 = GetCN("复核人"):      C检查技师 = GetCN("检查技师"):  C采图时间 = GetCN("采图时间")
        C随访描述 = GetCN("随访描述"):  C影像类别 = GetCN("影像类别"):  C病人ID = GetCN("病人ID")
        C主页ID = GetCN("主页ID"):      C挂号单 = GetCN("挂号单"):      C医嘱ID = GetCN("医嘱ID")
        C发送号 = GetCN("发送号"):      C病人科室ID = GetCN("病人科室ID"): C检查UID = GetCN("检查UID")
        C检查状态 = GetCN("检查状态"):  CNO = GetCN("NO"):              C记录性质 = GetCN("记录性质")
        C转出 = GetCN("转出"):          C床号 = GetCN("床号"):          C当前病区ID = GetCN("当前病区ID")
        C报告发放 = GetCN("报告发放"):  C诊断分类 = GetCN("诊断分类"):  C执行科室ID = GetCN("执行科室ID")
        C关联ID = GetCN("关联ID"):      C病人科室 = GetCN("病人科室"):  C就诊卡号 = GetCN("就诊卡号")
        C单据号 = GetCN("单据号"):      C身份证号 = GetCN("身份证号"):  C收费 = GetCN("收费")
        '提取并指定列宽
        .ColWidth(C路径) = GetCW("路径")
        .ColWidth(C紧急) = GetCW("紧急"):           .ColWidth(C来源) = GetCW("来源"):           .ColWidth(C阳性) = GetCW("阳性")
        .ColWidth(C质量) = GetCW("质量"):           .ColWidth(C姓名) = GetCW("姓名"):           .ColWidth(C检查号) = GetCW("检查号")
        .ColWidth(C检查过程) = GetCW("检查过程"):   .ColWidth(C性别) = GetCW("性别"):           .ColWidth(C年龄) = GetCW("年龄")
        .ColWidth(C标识号) = GetCW("标识号"):       .ColWidth(C医嘱内容) = GetCW("医嘱内容"):   .ColWidth(C部位方法) = GetCW("部位方法")
        .ColWidth(C执行间) = GetCW("执行间"):       .ColWidth(C报到时间) = GetCW("报到时间"):   .ColWidth(C申请时间) = GetCW("申请时间")
        .ColWidth(C开嘱医生) = GetCW("开嘱医生"):   .ColWidth(C身高) = GetCW("身高"):           .ColWidth(C体重) = GetCW("体重")
        .ColWidth(C婴儿) = GetCW("婴儿"):           .ColWidth(C登记人) = GetCW("登记人"):       .ColWidth(C报到人) = GetCW("报到人")
        .ColWidth(C完成人) = GetCW("完成人"):       .ColWidth(C打印胶片) = GetCW("打印胶片"):   .ColWidth(C报告操作) = GetCW("报告操作")
        .ColWidth(C绿色通道) = GetCW("绿色通道"):   .ColWidth(C报告打印) = GetCW("报告打印"):   .ColWidth(C报告人) = GetCW("报告人")
        .ColWidth(C复核人) = GetCW("复核人"):       .ColWidth(C检查技师) = GetCW("检查技师"):   .ColWidth(C采图时间) = GetCW("采图时间")
        .ColWidth(C随访描述) = GetCW("随访描述"):   .ColWidth(C影像类别) = GetCW("影像类别"):   .ColWidth(C病人ID) = GetCW("病人ID")
        .ColWidth(C主页ID) = GetCW("主页ID"):       .ColWidth(C挂号单) = GetCW("挂号单"):       .ColWidth(C医嘱ID) = GetCW("医嘱ID")
        .ColWidth(C发送号) = GetCW("发送号"):       .ColWidth(C病人科室ID) = GetCW("病人科室ID"): .ColWidth(C检查UID) = GetCW("检查UID")
        .ColWidth(C检查状态) = GetCW("检查状态"):   .ColWidth(CNO) = GetCW("NO"):               .ColWidth(C记录性质) = GetCW("记录性质")
        .ColWidth(C转出) = GetCW("转出"):           .ColWidth(C床号) = GetCW("床号"):           .ColWidth(C当前病区ID) = GetCW("当前病区ID")
        .ColWidth(C报告发放) = GetCW("报告发放"):   .ColWidth(C诊断分类) = GetCW("诊断分类"):   .ColWidth(C执行科室ID) = GetCW("执行科室ID")
        .ColWidth(C关联ID) = GetCW("关联ID"):       .ColWidth(C病人科室) = GetCW("病人科室"):   .ColWidth(C就诊卡号) = GetCW("就诊卡号")
        .ColWidth(C单据号) = GetCW("单据号"):       .ColWidth(C身份证号) = GetCW("身份证号"):   .ColWidth(C收费) = GetCW("收费")
        
        '列名称
        .Cell(flexcpData, 0, C路径) = "路径"
        .Cell(flexcpData, 0, C紧急) = "紧急":               .Cell(flexcpData, 0, C来源) = "来源":               .Cell(flexcpData, 0, C阳性) = "阳性"
        .Cell(flexcpData, 0, C质量) = "质量":               .Cell(flexcpData, 0, C姓名) = "姓名":               .Cell(flexcpData, 0, C检查号) = "检查号"
        .Cell(flexcpData, 0, C检查过程) = "检查过程":       .Cell(flexcpData, 0, C性别) = "性别":               .Cell(flexcpData, 0, C年龄) = "年龄"
        .Cell(flexcpData, 0, C标识号) = "标识号":           .Cell(flexcpData, 0, C医嘱内容) = "医嘱内容":       .Cell(flexcpData, 0, C部位方法) = "部位方法"
        .Cell(flexcpData, 0, C执行间) = "执行间":           .Cell(flexcpData, 0, C报到时间) = "报到时间":       .Cell(flexcpData, 0, C申请时间) = "申请时间"
        .Cell(flexcpData, 0, C开嘱医生) = "开嘱医生":       .Cell(flexcpData, 0, C身高) = "身高":               .Cell(flexcpData, 0, C体重) = "体重"
        .Cell(flexcpData, 0, C婴儿) = "婴儿":               .Cell(flexcpData, 0, C登记人) = "登记人":           .Cell(flexcpData, 0, C报到人) = "报到人"
        .Cell(flexcpData, 0, C完成人) = "完成人":           .Cell(flexcpData, 0, C打印胶片) = "打印胶片":       .Cell(flexcpData, 0, C报告操作) = "报告操作"
        .Cell(flexcpData, 0, C绿色通道) = "绿色通道":       .Cell(flexcpData, 0, C报告打印) = "报告打印":       .Cell(flexcpData, 0, C报告人) = "报告人"
        .Cell(flexcpData, 0, C复核人) = "复核人":           .Cell(flexcpData, 0, C检查技师) = "检查技师":       .Cell(flexcpData, 0, C采图时间) = "采图时间"
        .Cell(flexcpData, 0, C随访描述) = "随访描述":       .Cell(flexcpData, 0, C影像类别) = "影像类别":       .Cell(flexcpData, 0, C病人ID) = "病人ID"
        .Cell(flexcpData, 0, C主页ID) = "主页ID":           .Cell(flexcpData, 0, C挂号单) = "挂号单":           .Cell(flexcpData, 0, C病人科室ID) = "病人科室ID"
        .Cell(flexcpData, 0, C医嘱ID) = "医嘱ID":           .Cell(flexcpData, 0, C发送号) = "发送号":           .Cell(flexcpData, 0, C检查UID) = "检查UID"
        .Cell(flexcpData, 0, C检查状态) = "检查状态":       .Cell(flexcpData, 0, CNO) = "NO":                   .Cell(flexcpData, 0, C记录性质) = "记录性质"
        .Cell(flexcpData, 0, C转出) = "转出":               .Cell(flexcpData, 0, C床号) = "床号":               .Cell(flexcpData, 0, C当前病区ID) = "当前病区ID"
        .Cell(flexcpData, 0, C报告发放) = "报告发放":        .Cell(flexcpData, 0, C诊断分类) = "诊断分类":      .Cell(flexcpData, 0, C执行科室ID) = "执行科室ID"
        .Cell(flexcpData, 0, C关联ID) = "关联ID":           .Cell(flexcpData, 0, C病人科室) = "病人科室":       .Cell(flexcpData, 0, C就诊卡号) = "就诊卡号"
        .Cell(flexcpData, 0, C单据号) = "单据号":           .Cell(flexcpData, 0, C身份证号) = "身份证号":       .Cell(flexcpData, 0, C收费) = "收费"
        
        '显示列名称
        .TextMatrix(0, C路径) = "路径"
        Set .Cell(flexcpPicture, 0, C紧急) = Imglist.ListImages("紧急").Picture
        Set .Cell(flexcpPicture, 0, C来源) = Imglist.ListImages("住院").Picture
        Set .Cell(flexcpPicture, 0, C阳性) = Imglist.ListImages("阳性").Picture
        Set .Cell(flexcpPicture, 0, C收费) = Imglist.ListImages("收费").Picture
        .TextMatrix(0, C质量) = "质":               .TextMatrix(0, C姓名) = "姓名":              .TextMatrix(0, C检查号) = "检查号"
        .TextMatrix(0, C检查过程) = "检查过程":     .TextMatrix(0, C性别) = "性别":             .TextMatrix(0, C年龄) = "年龄"
        .TextMatrix(0, C标识号) = "标识号":         .TextMatrix(0, C医嘱内容) = "医嘱内容":     .TextMatrix(0, C部位方法) = "部位方法"
        .TextMatrix(0, C执行间) = "执行间":         .TextMatrix(0, C报到时间) = "报到时间":     .TextMatrix(0, C申请时间) = "申请时间"
        .TextMatrix(0, C开嘱医生) = "开嘱医生":     .TextMatrix(0, C身高) = "身高":             .TextMatrix(0, C体重) = "体重"
        .TextMatrix(0, C婴儿) = "婴儿":             .TextMatrix(0, C登记人) = "登记人":         .TextMatrix(0, C报到人) = "报到人"
        .TextMatrix(0, C完成人) = "完成人":         .TextMatrix(0, C打印胶片) = "打印胶片":     .TextMatrix(0, C报告操作) = "报告操作"
        .TextMatrix(0, C绿色通道) = "绿色通道":     .TextMatrix(0, C报告打印) = "报告打印":     .TextMatrix(0, C报告人) = "报告人"
        .TextMatrix(0, C复核人) = "复核人":         .TextMatrix(0, C检查技师) = "检查技师":     .TextMatrix(0, C采图时间) = "采图时间"
        .TextMatrix(0, C随访描述) = "随访描述":     .TextMatrix(0, C影像类别) = "影像类别":     .TextMatrix(0, C病人ID) = "病人ID"
        .TextMatrix(0, C主页ID) = "主页ID":         .TextMatrix(0, C挂号单) = "挂号单":         .TextMatrix(0, C病人科室ID) = "病人科室ID"
        .TextMatrix(0, C医嘱ID) = "医嘱ID":         .TextMatrix(0, C发送号) = "发送号":         .TextMatrix(0, C检查UID) = "检查UID"
        .TextMatrix(0, C检查状态) = "检查状态":     .TextMatrix(0, CNO) = "NO":                 .TextMatrix(0, C记录性质) = "记录性质"
        .TextMatrix(0, C转出) = "转出":             .TextMatrix(0, C床号) = "床号":             .TextMatrix(0, C当前病区ID) = "当前病区ID"
        .TextMatrix(0, C报告发放) = "报告发放":     .TextMatrix(0, C诊断分类) = "诊断分类":     .TextMatrix(0, C执行科室ID) = "执行科室ID"
        .TextMatrix(0, C关联ID) = "关联ID":         .TextMatrix(0, C病人科室) = "病人科室":     .TextMatrix(0, C就诊卡号) = "就诊卡号"
        .TextMatrix(0, C单据号) = "单据号":         .TextMatrix(0, C身份证号) = "身份证号"
        
        Dim i As Integer
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignLeftCenter
        Next
        
        '读取和设置病人列表的字体
        .FontName = zldatabase.GetPara("病人列表内容字体", glngSys, mlngModul, "宋体")
        .FontSize = Val(zldatabase.GetPara("病人列表内容字号", glngSys, mlngModul, 9))
        .FontBold = zldatabase.GetPara("病人列表内容粗体", glngSys, mlngModul, 0) = 1
        .FontItalic = zldatabase.GetPara("病人列表内容斜体", glngSys, mlngModul, 0) = 1
        .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = zldatabase.GetPara("病人列表表头字体", glngSys, mlngModul, "宋体")
        .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = Val(zldatabase.GetPara("病人列表表头字号", glngSys, mlngModul, 9))
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = zldatabase.GetPara("病人列表表头粗体", glngSys, mlngModul, 0) = 1
        .Cell(flexcpFontItalic, 0, 0, 0, .Cols - 1) = zldatabase.GetPara("病人列表表头斜体", glngSys, mlngModul, 0) = 1
        .Editable = flexEDNone
    End With
End Sub


Private Sub mfrmPacsReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
    
    '嵌入式编写报告时，保存之后，重新开启自动刷新功能
    Call subTriggleRefreshTimer(True)
End Sub

Private Sub mfrmPacsReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub mfrmPacsReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mfrmPacsReport_AfterSaved(ByVal lngOrderID As Long, frmOwnerForm As Form)
    Call AfterReportSaved(lngOrderID, frmOwnerForm)
End Sub

Private Sub mfrmPacsReport_BeforeEdit()
Dim lngOrderID As Long

    On Error GoTo errHandle
    lngOrderID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))
    If CheckConcurrentReport(Me, lngOrderID) Then '检查是否有人正在操作报告
        Call UpdateReporter(lngOrderID, UserInfo.姓名)
    Else
        Call mfrmPacsReport.PromptModify(True)
    End If
    
    '嵌入式编写报告时，编辑报告之前，先关闭自动刷新功能
    Call subTriggleRefreshTimer(False)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mfrmPacsReportDock_AfterOpen()
    Call AfterReportOpen
End Sub

Private Sub mfrmPacsReportDock_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtFilter.Text = "" And Me.ActiveControl Is txtFilter Then
        IDKind.IDKind = IDKinds.C2身份证号
        mstrCurFindtype = "身份证"
        txtFilter = strID
        Call txtFilter_KeyDown(vbKeyReturn, 0)
    ElseIf txtLocate.Text = "" And Me.ActiveControl Is txtLocate Then
        IDKind.IDKind = IDKinds.C2身份证号
        mstrLocateType = "身份证"
        txtLocate = strID
        Call txtLocate_KeyDown(vbKeyReturn, 0)
    End If
End Sub

Private Sub mobjInAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lng医嘱ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select 医嘱ID FROM 病人医嘱报告 where 病历ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取医嘱ID", CLng(报告ID))
    If rsTemp.EOF Then Exit Sub
    
    lng医嘱ID = Nvl(rsTemp!医嘱ID, 0)
    mobjReport.zlRefresh lng医嘱ID, mlngCur科室ID, False '以不可Edit方式刷新对像
    
    Set cbrControl = cbrMain(2).Controls.Find(, conMenu_Help_Help)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '调用查阅报告
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjInAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
    '超过100张图像的序列，默认每隔5张传一张
    Call OpenViewer(mobjPacsCore, 医嘱ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjOutAdvice_ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean)
Dim cbrControl As CommandBarControl, lng医嘱ID As Long, rsTemp As ADODB.Recordset
    gstrSQL = "select 医嘱ID FROM 病人医嘱报告 where 病历ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取医嘱ID", CLng(报告ID))
    If rsTemp.EOF Then Exit Sub
    
    lng医嘱ID = Nvl(rsTemp!医嘱ID, 0)
    mobjReport.zlRefresh lng医嘱ID, mlngCur科室ID, False '以不可Edit方式刷新对像
    
    Set cbrControl = cbrMain(2).Controls.Find(, conMenu_Help_Help)
    cbrControl.ID = conMenu_File_Open
    mobjReport.zlExecuteCommandBars cbrControl '调用查阅报告
    cbrControl.ID = conMenu_Help_Help
End Sub

Private Sub mobjOutAdvice_ViewPACSImage(ByVal 医嘱ID As Long)
    '超过100张图像的序列，默认每隔5张传一张
    Call OpenViewer(mobjPacsCore, 医嘱ID, False, Me, , , mblnLocalizerBackward, 5)
End Sub

Private Sub mobjPacsCore_AfterSaveOuterImage(strStudyUID As String)
    '保存了外部图像，刷新图像的序列列表
    On Error GoTo err
    
    '没有记录则退出
    If Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0 Then Exit Sub
    
    '是当前的检查，才刷新检查的序列列表
    If vsList.TextMatrix(vsList.Row, GetCN("检查UID")) = strStudyUID Then
        Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), vsList.TextMatrix(vsList.Row, GetCN("发送号")), mstrPrivs, vsList.TextMatrix(vsList.Row, GetCN("转出")) = 1, True)
    End If
    
    Exit Sub
err:
    '不处理
End Sub

Private Sub mobjPacsCore_AfterSaveReportImage(strStudyUID As String)
    If mblnPacsReport = True Then
        mfrmPacsReport.RefPacsPic '刷新图片
        If Not mfrmPacsReportDock Is Nothing Then
            mfrmPacsReportDock.RefPacsPic '刷新图片
        End If
    Else
        mobjReport.RefPacsPic '刷新图片
    End If
End Sub

Private Sub mobjReport_AfterClosed(ByVal lngOrderID As Long)
    Call EditorClosed(lngOrderID)
End Sub
Public Sub EditorClosed(ByVal lngOrderID As Long)
    Dim i As Integer
    Dim j As Integer
    
    Call UpdateReporter(lngOrderID, "")
    '处理PACS报告编辑器的窗口数组
    On Error Resume Next
    If mblnPacsReport = True Then
        '查找窗口数组，找到对应的窗口并删除
        If SafeArrayGetDim(mobjPacsReportArry) <> 0 Then
            For i = 1 To UBound(mobjPacsReportArry)
                If mobjPacsReportArry(i).mlngAdviceID = lngOrderID Then
                    '从数组中删除
                    For j = i To UBound(mobjPacsReportArry)
                        Set mobjPacsReportArry(j) = mobjPacsReportArry(j + 1)
                    Next j
                    ReDim Preserve mobjPacsReportArry(UBound(mobjPacsReportArry) - 1) As frmReport
                    Exit For
                End If
            Next i
        End If
        
        If Not mfrmPacsReportDock Is Nothing Then
            If lngOrderID = mfrmPacsReportDock.mlngAdviceID Then
                '关闭当前报告窗口，将当前窗口设置成空
                Set mfrmPacsReportDock = Nothing
            End If
        End If
    End If
End Sub

Private Sub mobjReport_AfterDeleted(ByVal lngOrderID As Long)
    AfterDeleted lngOrderID
End Sub

Private Sub AfterDeleted(ByVal lngOrderID As Long)
    On Error GoTo errHandle
    gstrSQL = "ZL_影像报告标记_Clear(" & lngOrderID & ")"
    zldatabase.ExecuteProcedure gstrSQL, "清空标记"
    Call RefreshList
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterOpen(ByVal intEditType As zlRichEPR.EditTypeEnum)
    Call AfterReportOpen
End Sub

Private Sub AfterReportOpen()

    On Error GoTo errHandle
    
    Call UpdateReporter(CLng(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))), UserInfo.姓名)
    
    Call OpenViewerWithReport
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub OpenViewerWithReport()
'跟据参数打开报告后同时打开观片站，判断是否打开观片站
    Dim lngOrderID As Long
    
    On Error GoTo err
    
    lngOrderID = vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))
    
    If mblnShowImgAtReport And vsList.TextMatrix(vsList.Row, GetCN("检查UID")) <> "" Then
        Dim intImageInverval As Integer
        
        intImageInverval = Val(mfrmPACSImg.cbrMain.FindControl(, conMenu_Manage_ImageInterval, , True).Text)
        Call OpenViewer(mobjPacsCore, lngOrderID, False, Me, , , mblnLocalizerBackward, intImageInverval)
    End If
    
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mobjReport_AfterPrinted(ByVal lngOrderID As Long)
    Call AfterPrinted(lngOrderID)
End Sub
Public Sub AfterPrinted(lngOrderID As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "ZL_影像报告打印_Update(" & lngOrderID & ")"
    zldatabase.ExecuteProcedure gstrSQL, "更新打印标记"
    
    If Not mblnIgnoreResult And mintResultInput = 2 Then
        strSQL = "Select 结果阳性  From  病人医嘱发送 Where 医嘱id= [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(strSQL, "提取结果阳性", lngOrderID)
        
        If IsNull(rsTemp!结果阳性) Then  '在报告时提示结果阴阳性
            Call PromptResult(lngOrderID, mlngModul, Me)
        End If
    End If
    
    If mblnPrintCommit = True Then
        Call Menu_Manage_检查最终完成(lngOrderID, False)
    End If
    
    Call RefreshList
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub mobjReport_AfterSaved(ByVal lngOrderID As Long)
    Call AfterReportSaved(lngOrderID, Me)
End Sub

Public Sub AfterReportSaved(lngOrderID As Long, frmOwnerForm As Form)
'保存报告之后的处理
'执行过程：2-已报到；3-已检查；4-已报告；5-已审核；6-已完成

    Dim intState As Integer, lngSendId As Long
    Dim str签名 As String
    Dim str创建人 As String
    Dim str保存人 As String
    Dim bln保存结果阳性 As Boolean
    Dim arrSQL() As Variant
    Dim blnInTrans As Boolean
    Dim i As Integer
    
    arrSQL = Array()
    
    On Error GoTo errHandle
    
    If mblnPacsReport = True Then
'        mfrmPacsReport.zlRefresh 0, 0, 0
    Else
        mobjReport.zlRefresh 0, mlngCur科室ID, False
    End If
    
    '获取本次检查的执行过程
    intState = getStudyState(lngOrderID, lngSendId, str创建人, str签名, str保存人, bln保存结果阳性)
    
    'intState =1--已登记；2--已报到；3--已检查；4--已报告；5--已审核；6--已完成（本过程不存在这个返回值）
    If intState = 2 Or intState = 3 Then
        gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & "," & intState & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        gstrSQL = "ZL_影像报告保存_Update(" & lngOrderID & ",'" & str创建人 & "','')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
    Else
        If intState = 4 Then
            '诊断签名，最后一次签名为医师,执行过程为已报告
            '有可能的情况 1-医师第N次签名 2-主任级别最后一次退签 3-修订模式下保存(签名级别=0)
            gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & "," & intState & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            '应该填写创建人才准确，回退的时候，回退的人是保存人，但是不是报告创建人
            '医生诊断签名,无论是第N次，此时，报告人需要保存，复核人需要清空;
            gstrSQL = "ZL_影像报告保存_Update(" & lngOrderID & ",'" & str创建人 & "','')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        ElseIf intState = 5 Then
            '审核签名，主任及以上级别签名，签名级别>=2,执行过程为已审核
            gstrSQL = "Zl_影像检查_State(" & lngOrderID & "," & lngSendId & "," & intState & ")"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            gstrSQL = "ZL_影像报告保存_Update(" & lngOrderID & ",'" & str创建人 & "','" & IIf(str签名 <> "", str签名, str保存人) & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
        End If
    End If
    
    gcnOracle.BeginTrans        '----------保存检查状态及报告人
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存检查状态及报告人")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If intState = 4 Or intState = 5 Then
        If Not mblnIgnoreResult And Not bln保存结果阳性 Then  '在报告时提示结果阴阳性
            If mblnReportWithResult Then '无影像诊断为阴性  -无提示自动标记
                gstrSQL = "ZL_影像检查_结果(" & lngOrderID & ",0)"
                zldatabase.ExecuteProcedure gstrSQL, "标记阴阳性"
            ElseIf mintResultInput = 1 Then
                Call PromptResult(lngOrderID, mlngModul, frmOwnerForm)  ' Me)
            End If
        End If
    End If
    
    If intState = 5 And mblnCompleteCommit Then '如果“审核后直接完成”
        Call Menu_Manage_检查最终完成(lngOrderID, False)
    End If
    
    '病人状态跟踪
    Call StateCheck(intState)
    Exit Sub
errHandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub StateCheck(ByVal intState As Integer, Optional ByVal lngAdviceID As Long)
    
    If mblnPatTrack Then
        Select Case intState '跟据病人新状态确定新状态过滤是否选中
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
        Call RefreshList(lngAdviceID)
    Else '不跟踪只刷新列表
        Call RefreshList
    End If
End Sub
Private Function ShowBillList(objPopup As CommandBarPopup) As Boolean
'功能：显示当前执行医嘱可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim objControl As CommandBarControl
        
    On Error GoTo errH
    
    objPopup.CommandBar.Controls.DeleteAll
    
    With vsList
        gstrSQL = "Select Distinct C.编号,C.名称,C.说明" & _
            " From 病人医嘱记录 A,病历单据应用 B,病历文件列表 C" & _
            " Where A.ID=[1] And A.相关ID IS NULL" & _
            " And A.诊疗项目ID=B.诊疗项目ID" & _
            " And B.应用场合=[2] And B.病历文件ID=C.ID And C.种类=7" & _
            " Order by C.编号"
        If .TextMatrix(.Row, GetCN("转出")) = 1 Then
            gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
            gstrSQL = Replace(gstrSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(.TextMatrix(.Row, GetCN("医嘱ID"))), CLng(Decode(.TextMatrix(.Row, GetCN("来源")), "门", 1, "住", 2, "外", 3, 4)))
    End With
    
    If Not rsTmp.EOF Then
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Manage_RequestPrint * 10# + 1, rsTmp!名称 & "(&0)")
            objControl.parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
        End With
        cbrMain.KeyBindings.Add 0, vbKeyF10, conMenu_Manage_RequestPrint * 10# + 1
    End If
    
    ShowBillList = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub FuncBillPrint(objControl As CommandBarControl)
'功能：打印诊疗单据
    On Error GoTo errH
    If objControl.parameter = "" Then '奇怪，直接按F10时，是一个空的Control
        Set objControl = cbrMain.FindControl(, conMenu_Manage_RequestPrint * 10# + 1, , True)
        If objControl Is Nothing Then Exit Sub
    End If
    If objControl.parameter = "" Then Exit Sub
    

    If ReportPrintSet(gcnOracle, glngSys, objControl.parameter, Me) Then
        Call ReportOpen(gcnOracle, glngSys, objControl.parameter, Me, "NO=" & vsList.TextMatrix(vsList.Row, GetCN("NO")), "性质=" & vsList.TextMatrix(vsList.Row, GetCN("记录性质")), 1)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub RefreshList(Optional ByVal lngAdviceID As Long = 0)
Dim i As Integer, lngcur医嘱ID As Long, lngRow As Long, lngTopRow As Long
    With vsList
        If lngAdviceID <> 0 Then
            lngcur医嘱ID = lngAdviceID
        Else
            lngcur医嘱ID = Val(.TextMatrix(.Row, GetCN("医嘱ID"))) '当前行医嘱ID
            lngRow = .Row: lngTopRow = .TopRow               '当前行和顶行之间的差距
        End If
        
        Call LoadPatiList
        If lngcur医嘱ID = 0 Then
            Call .Select(1, GetCN("姓名"))
            Exit Sub
        End If
        
        '有记录时要重新定位回之前记录
        On Error Resume Next
        lngcur医嘱ID = .FindRow(CStr(lngcur医嘱ID), , GetCN("医嘱ID"))
        If lngcur医嘱ID <> -1 Then
            lngRow = Abs(lngRow - lngTopRow)
            If .Row = lngcur医嘱ID Then '相同时不会触发CHANGE事件
                Call vsList_RowColChange '强制刷新右边子窗体
            Else
                .Row = lngcur医嘱ID
            End If
            .TopRow = .Row - lngRow
        Else
            If .Row <> 1 Then
                .Row = 1
            Else
                Call vsList_RowColChange '强制刷新右边子窗体
            End If
        End If
        err.Clear
    End With
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    fraRegist.Left = 0
    fraRegist.Top = -75
    fraInfo.Top = -75
    fraInfo.Left = fraRegist.Left + fraRegist.Width
    fraInfo.Width = picInfo.ScaleWidth - fraInfo.Left
    
    lblCash.Top = (picInfo.ScaleHeight - lblCash.Height) / 2 - fraInfo.Top
    lblCash.Left = fraInfo.Width - lblCash.Width - 100

    lbl个人信息.Width = lblCash.Left
    lbl检查信息.Width = lblCash.Left
End Sub

Private Sub LoadPatiList()
'功能：读取当前医技科室的执行医嘱(病人)清单
Dim strSQL As String, strSQLBak As String, i As Long, rsList As ADODB.Recordset
Dim str来源 As String
Dim strFilter As String
Dim strModalitys As String
Dim blnUseTime As Boolean       '是否使用时间条件

    If Not mblnInitOk Then Exit Sub      '初始化未完成
    mblnvsRefresh = True
    On Error GoTo errHandle
    With SQLCondition
        blnUseTime = False  '默认不使用时间条件
        '界面查找条件不使用时间索引
        If .门诊号 <> 0 Then
            strFilter = " And C.门诊号=[1]"
        ElseIf .住院号 <> 0 Then
            strFilter = " And C.住院号=[2]"
        ElseIf .就诊卡 <> "" Then
            strFilter = " And C.就诊卡号=[3]"
        ElseIf .姓名 <> "" And InStr(.姓名, "*") = 0 Then   '姓名特殊处理，带*号表示模糊查询
            strFilter = " And C.姓名=[4]"
        ElseIf .身份证 <> "" Then
            strFilter = " And C.身份证号=[5]"
        ElseIf .IC卡 <> "" Then
            strFilter = " And C.IC卡=[6]"
        ElseIf .单据号 <> "" Then
            strFilter = " And A.NO=[7] "
        ElseIf .检查号 <> 0 Then
            strFilter = " And H.检查号=[8] "
        Else
        '其他条件查询，使用时间索引
            blnUseTime = True
            '填写过滤时间条件
            '时间查询方式 1=按申请时间（病人医嘱发送.发送时间）、2=按报到时间（病人医嘱发送.首次时间）、3=采图时间（影像检查记录.接收日期）
            If .时间类型 = 1 Then       '按申请时间
                strFilter = " And A.发送时间 Between [9] and "
            ElseIf .时间类型 = 2 Then   '按报到时间
                strFilter = " And A.首次时间 Between [9] and "
            Else                        '采图时间
                strFilter = " And H.接收日期 Between [9] and "
            End If
            If .结束时间 <> CDate(0) Then
                strFilter = strFilter & " [10] "
            Else
                strFilter = strFilter & " Sysdate+1/(24*3600) "
            End If
            
            '先处理姓名中带*号的，进行带时间索引的模糊查询
            If .姓名 <> "" And InStr(.姓名, "*") <> 0 Then
                .姓名 = Replace(.姓名, "*", "%")
                strFilter = strFilter & " And C.姓名 like [4]"
            End If
            
            If .性别 <> "" Then
                strFilter = strFilter & " And Nvl(H.性别,C.性别)=[29]"
            End If
        
        
            '病人年龄-开始年龄(只有当条件使用“到”，即在多少年龄之间时，才使用开始年龄)
            If .开始年龄 <> -1 Then
                If .年龄条件 = "~" Then
                    strFilter = strFilter & " And ZL_AgeToDays(C.年龄)>=[30]"
                End If
            End If
            
            '病人年龄-结束年龄
            If .结束年龄 <> -1 Then
                If .年龄条件 = "~" Then
                    strFilter = strFilter & " And ZL_AgeToDays(C.年龄)<=[31]"
                Else
                    strFilter = strFilter & " And ZL_AgeToDays(C.年龄)" & .年龄条件 & "[31]"
                End If
            End If
            
            If .病人科室 <> 0 Then
                strFilter = strFilter & " And B.病人科室ID+0=[11] "
            End If
        
            If .标本部位 <> "" Then
                strFilter = strFilter & " And instr(B.医嘱内容,[12])>0"
            End If
            
            If .结果阳性 <> -1 Then
                strFilter = strFilter & " And Nvl(A.结果阳性, 0)=[32]"
            End If
            
            If .诊断医生 <> "" Then
                strFilter = strFilter & " And H.报告人=[13] "
            End If
            
            If .审核医生 <> "" Then
                strFilter = strFilter & " And H.复核人=[14] "
            End If
            
            If .影像质量 <> "" Then
                strFilter = strFilter & " And H.影像质量=[15]"
            End If
            
            If .检查技师 <> "" Then
                strFilter = strFilter & " And H.检查技师=[16]"
            End If
            
            '影像类别有两个地方做过滤条件的选择，过滤窗口和主程序上面，以主程序中的为主
            If mintcmd影像类别 > 0 Then
                Dim objControl As CommandBarControl
                
                Set objControl = cbrdock.FindControl(, ID_影像类别)
                For i = 1 To objControl.CommandBar.Controls.Count
                    If objControl.CommandBar.FindControl(, ID_影像类别 + i).Checked = True Then
                        strModalitys = strModalitys & "," & objControl.CommandBar.FindControl(, ID_影像类别 + i).DescriptionText
                    End If
                Next i
                If strModalitys <> "" Then
                    strFilter = strFilter & " And instr([26],H.影像类别)>0 "
                End If
            Else
                If .影像类别 <> "" Then
                    strFilter = strFilter & " And H.影像类别=[17] "
                End If
            End If
            
            
            
            If .随访 <> "" Then
                strFilter = strFilter & " And  Instr(H.随访描述, [18]) > 0 "
            End If
            
            If .疾病诊断 <> "" Then
                strFilter = strFilter & " And B.ID IN ( Select t.医嘱id From 病人医嘱报告 t Where t.病历id In " & _
                                                                    " (Select Distinct A.ID  " & _
                                                                        "From 电子病历记录 A,电子病历内容 B " & _
                                                                        "Where A.创建时间>[9] AND A.Id=B.文件ID  " & _
                                                                            "And B.对象类型=7 And instr(B.对象属性,'52;')>0 And instr(B.内容文本,[19])>0))"
            End If
            
            Dim strSubFilter As String '增加PACS报告检索条件
            If .检查所见 <> "" Then
                strSubFilter = " (b.内容文本 ='检查所见' And Instr(c.内容文本, [20]) > 0)"
            End If
            
            If .诊断意见 <> "" Then
                If strSubFilter = "" Then
                    strSubFilter = " (b.内容文本 ='诊断意见' And Instr(c.内容文本, [21]) > 0)"
                Else
                    strSubFilter = strSubFilter & " or (b.内容文本 ='诊断意见' And Instr(c.内容文本, [21]) > 0)"
                End If
            End If
            
            If .建议 <> "" Then
                If strSubFilter = "" Then
                    strSubFilter = " (b.内容文本 ='建议' And Instr(c.内容文本, [22]) > 0)"
                Else
                    strSubFilter = strSubFilter & " or (b.内容文本 ='建议' And Instr(c.内容文本, [22]) > 0)"
                End If
            End If
            
            If strSubFilter <> "" Then
                strSubFilter = " (" & strSubFilter & ")"
                
                strFilter = strFilter & " And B.ID IN ( Select t.医嘱id From 病人医嘱报告 t Where t.病历id In  " _
                    & " (Select Distinct a.ID From 电子病历记录 a, 电子病历内容 b,电子病历内容 c " _
                    & " Where a.创建时间 > [9] And a.Id = b.文件id And b.Id = C.父ID And b.对象类型 = 3 And c.对象类型 = 2 And c.终止版 = 0 and " _
                    & strSubFilter & "))"
            End If
           
'            If .检查过程 <> "" Then
'                If .检查过程 = "全部" Then
'
'                ElseIf .检查过程 = "已登记" Then
'                    strFilter = strFilter & " And (A.执行过程 =0 or A.执行过程=1 Or A.执行过程 Is Null) "
'                ElseIf .检查过程 = "已报到" Then
'                    strFilter = strFilter & " And (A.执行过程 = 2 and h.报告人 is null) "
'                ElseIf .检查过程 = "已检查" Then
'                    strFilter = strFilter & " And (A.执行过程 = 3 and h.报告人 is null) "
'                ElseIf .检查过程 = "处理中" Then
'                    strFilter = strFilter & " And (not h.报告操作 is null) "
'                ElseIf .检查过程 = "报告中" Then
'                    strFilter = strFilter & " And ((A.执行过程 =2 or A.执行过程=3) and not h.报告人 is null and h.报告操作 is null) "
'                ElseIf .检查过程 = "已报告" Then
'                    strFilter = strFilter & " And (A.执行过程=4 and h.复核人 is null) "
'                ElseIf .检查过程 = "审核中" Then
'                    strFilter = strFilter & " And (A.执行过程=4 and not h.复核人 is null) "
'                ElseIf .检查过程 = "已审核" Then
'                    strFilter = strFilter & " And A.执行过程=5 "
'                ElseIf .检查过程 = "已完成" Then
'                    strFilter = strFilter & " And A.执行过程=6 "
'                End If
'            End If
        End If
        
        '“过滤窗口”和“界面查找”条件独立，界面查找条件不使用时间索引，以下条件为共用条件
        
        '病人来源 (1-门诊,2-住院,3-外来,4-体检)
        '如果四种来源都选择了，表示查找所有病人，则不添加病人来源的查询条件
        If mblncmd门诊 And mblncmd住院 And mblncmd体检 And mblncmd外诊 Then
        
        Else
            If mblncmd门诊 Then str来源 = "1,"
            If mblncmd住院 Then str来源 = str来源 & "2,"
            If mblncmd外诊 Then str来源 = str来源 & "3,"
            If mblncmd体检 Then str来源 = str来源 & "4,"
            If str来源 <> "" Then   'str来源为空，表示没有选择任何来源，则不添加病人来源的查询条件
                str来源 = Mid(str来源, 1, Len(str来源) - 1)
                strFilter = strFilter & " And Instr([23],B.病人来源)> 0"
            End If
        End If
        
        If mstrRoom <> "" Then  '只显示执行间范围内的
            If Not mblncmd登记 Then
                strFilter = strFilter & " And Instr([24],','|| A.执行间 || ',' )>0"
            Else
                strFilter = strFilter & " And (Instr([24],','|| A.执行间 || ',' )>0 And Nvl(A.执行过程,0)>1 OR Nvl(A.执行过程,0)<2)"
            End If
        End If
    
        If mblnNoShowCancel Then '不显示取消登记的检查
            strFilter = strFilter & " And A.执行状态<>2 "
        End If
        
        If mblncmd本次 Then        '只显示本次住院记录
            strFilter = strFilter & vbNewLine & " And (B.病人来源=2 And B.主页ID=C.住院次数 Or Nvl(B.病人来源,0)<>2)"
        End If

        '是否选择了全部科室
        If mblnAllDepts = True Then
            strFilter = strFilter & " And Instr( [27],A.执行部门ID ) >0"
        Else
            strFilter = strFilter & " And A.执行部门ID+0=[25]"
        End If
         
        '检索报告内容
        If .报告内容 <> "" Then
            strFilter = strFilter & " And B.id IN ( Select t.医嘱id From 病人医嘱报告 t Where t.病历id In " & _
                                                                    " (Select Distinct A.ID " & _
                                                                    " From 电子病历记录 A,电子病历内容 B " & _
                                                                    " Where A.创建时间>[9] AND A.Id=B.文件ID " & _
                                                                    " And B.对象类型=2 And instr(B.内容文本,[28])>0 And B.终止版 = 0)) "
        End If
        
        gstrSQL = "Select /*+ RULE */ Distinct" & vbNewLine & _
                    "       A.医嘱ID,A.发送号,A.首次时间 报到时间,A.发送时间 申请时间,A.执行状态,nvl(A.执行过程,0) 检查过程,A.执行间,A.结果阳性 阳性," & vbNewLine & _
                    "       B.病人ID,B.主页ID,B.挂号单,B.病人科室ID,Decode(B.病人来源, 1, '门', 2, '住', 3, '外', 4, '体') 来源,B.医嘱内容,B.标本部位," & vbNewLine & _
                    "       Nvl(B.紧急标志, 0) 紧急标志, Nvl(B.婴儿, 0) 婴儿,B.开嘱医生,A.NO,C.当前床号,C.当前病区ID,Decode(B.病人来源,2,C.住院号,C.门诊号) 标识号," & vbNewLine & _
                    "       Nvl(H.姓名,C.姓名) 姓名,H.影像类别,H.检查号,Nvl(H.性别,C.性别) 性别,Nvl(H.年龄,C.年龄) 年龄,H.身高,H.体重,H.影像质量," & vbNewLine & _
                    "       Decode(B.病人来源,3,B.开嘱医生,A.发送人) 登记人,H.报到人,H.报告发放,H.关联ID, " & vbNewLine & _
                    "       H.完成人,H.是否打印,H.报告操作,H.绿色通道,H.报告打印,H.报告人,H.复核人,H.检查技师,H.接收日期 采图时间," & vbNewLine & _
                    "       H.随访描述,H.诊断分类,H.检查UID,A.执行部门ID as 执行科室ID,0 as 转出,F.名称 AS 病人科室, " & vbNewLine & _
                    "       C.就诊卡号,A.NO as 单据号,C.身份证号,D.状态 as 路径状态,A.计费状态,Decode(A.记录性质,2,1,Decode(a.计费状态,3,1,0)) as 收费 " & vbNewLine & _
                    " From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C,病人临床路径 D,影像检查记录 H,影像检查项目 G,部门表 F " & vbNewLine & _
                    " Where B.相关ID is NULL And A.医嘱ID=B.ID And A.医嘱ID=H.医嘱ID(+) And A.发送号=H.发送号(+) " & vbNewLine & _
                    " And B.诊疗项目ID=G.诊疗项目ID And B.病人ID=C.病人ID And B.病人科室id=F.ID " & vbNewLine & _
                    " And B.病人ID = D.病人ID(+) And B.主页ID = D.主页ID(+) And D.结束时间(+) is Null "
        gstrSQL = gstrSQL & vbNewLine & strFilter
        
        '通过"病人医嘱发送.计费状态"直接判断,原有值：-1-无须计费;0-未计费;1-已计费，对于记帐单（包括门诊记帐单），保持原有值不变。
        '对于收费单的发送记录，增加两种状态：2-部分收费，3-全部收费
        If mblncmd已缴 = True Then
            gstrSQL = gstrSQL & " and (A.记录性质 <> 1 Or (A.记录性质 = 1 And a.计费状态 = 3)) "
        ElseIf mblncmd未缴 = True Then
            gstrSQL = gstrSQL & " and (A.记录性质 = 1 And A.计费状态 <>3) "
        End If
        
        '当使用检查号查找时一定是报过到的，影像检查记录中有记录，此时取消左连接避免全表扫描
        '使用采集时间过滤，影像检查记录中有记录
        If .检查号 <> 0 Or (blnUseTime = True And SQLCondition.时间类型 = 3) Then
            gstrSQL = Replace(Replace(gstrSQL, "H.医嘱ID(+)", "H.医嘱ID"), "H.发送号(+)", "H.发送号")
        End If
        
        '如果有数据转出则还要检索后备表
        If mblnMoved Then
            strSQLBak = gstrSQL
            strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
            strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
            strSQLBak = Replace(strSQLBak, "影像检查记录", "H影像检查记录")
            strSQLBak = Replace(strSQLBak, "电子病历记录", "H电子病历记录")
            strSQLBak = Replace(strSQLBak, "电子病历内容", "H电子病历内容")
            strSQLBak = Replace(strSQLBak, "0 as 转出", "1 as 转出")
            strSQL = strSQL & " Union ALL " & strSQLBak
        End If
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order by 检查过程,报到时间,申请时间"
    
        Set rsList = zldatabase.OpenSQLRecord(gstrSQL, "提取病人列表", .门诊号, .住院号, .就诊卡, .姓名, .身份证, .IC卡, .单据号, _
                                            .检查号, .开始时间, .结束时间, .病人科室, .标本部位, .诊断医生, .审核医生, .影像质量, _
                                            .检查技师, .影像类别, .随访, .疾病诊断, .检查所见, .诊断意见, .建议, str来源, mstrRoom, _
                                            mlngCur科室ID, strModalitys, mstrCanUse科室IDs, .报告内容, .性别, .开始年龄, .结束年龄, .结果阳性)
    End With
    
    strFilter = ""
    If mblncmd登记 Then strFilter = "检查过程=0 or 检查过程=1 or "
    If mblncmd报到 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=2 or ", "检查过程=2 or ")
    If mblncmd检查 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=3 or ", "检查过程=3 or ")
    If mblncmd报告 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=4 or ", "检查过程=4 or ")
    If mblncmd审核 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=5 or ", "检查过程=5 or ")
    If mblncmd完成 Then strFilter = IIf(strFilter <> "", strFilter & "检查过程=6 or ", "检查过程=6 or ")
    If mblncmd登记 And mblncmd报到 And mblncmd检查 And mblncmd报告 And mblncmd审核 And mblncmd完成 Then
        strFilter = ""
    End If

    If strFilter <> "" Then
        strFilter = Mid(strFilter, 1, Len(strFilter) - 4)
        rsList.Filter = strFilter
    End If
    
    Call FillList(rsList)
    mblnvsRefresh = False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillList(ByVal rsTemp As ADODB.Recordset)
    Dim rsBaby As ADODB.Recordset
    Dim blnShowPath As Boolean      '是否显示路径列
    Dim intPathColNum As Integer
    
    On Error GoTo errHandle
    Call InitList
    If rsTemp.EOF Then stbThis.Panels(2).Text = "没有找到任何匹配的记录": Exit Sub
    
    With vsList
        .Rows = rsTemp.RecordCount + 1
        Do Until rsTemp.EOF
            .Row = rsTemp.AbsolutePosition
            
            If Nvl(rsTemp!路径状态, 0) = 1 Then
               Set .Cell(flexcpPicture, .Row, GetCN("路径")) = Imglist.ListImages("路径").Picture
               .TextMatrix(.Row, GetCN("路径")) = " "
               blnShowPath = True
            End If
            
            .Cell(flexcpData, .Row, GetCN("紧急")) = Val(rsTemp!紧急标志)
            If rsTemp!紧急标志 <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("紧急")) = Imglist.ListImages("紧急").Picture
            End If
            If rsTemp!来源 = "住" Then
                Set .Cell(flexcpPicture, .Row, GetCN("来源")) = Imglist.ListImages("住院").Picture
            End If
            .TextMatrix(.Row, GetCN("来源")) = rsTemp!来源
            .Cell(flexcpData, .Row, GetCN("来源")) = Decode(rsTemp!来源, "门", 1, "住", 2, "外", 3, 4)
            
            If rsTemp!收费 = 1 Then
                Set .Cell(flexcpPicture, .Row, GetCN("收费")) = Imglist.ListImages("收费").Picture
                .TextMatrix(.Row, GetCN("收费")) = " "  ' 做排序用
            End If
            
            If Nvl(rsTemp!阳性, 0) <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("阳性")) = Imglist.ListImages("阳性").Picture
                .TextMatrix(.Row, GetCN("阳性")) = " "  ' 做排序用
            End If
            
            If Nvl(rsTemp!绿色通道, 0) <> 0 Then
                Set .Cell(flexcpPicture, .Row, GetCN("姓名")) = Imglist.ListImages("绿色通道").Picture
            End If
            
            If Nvl(rsTemp!检查uid) <> "" Then
                Set .Cell(flexcpPicture, .Row, GetCN("检查号")) = Imglist.ListImages("影像").Picture
            End If
            
            .TextMatrix(.Row, GetCN("质量")) = Nvl(rsTemp!影像质量)
            .TextMatrix(.Row, GetCN("姓名")) = Nvl(rsTemp!姓名)
            .TextMatrix(.Row, GetCN("检查号")) = Nvl(rsTemp!检查号)
            .TextMatrix(.Row, GetCN("检查过程")) = IIf(rsTemp!执行状态 = 2, "已拒绝", Decode(Nvl(rsTemp!检查过程, 0), 0, "已登记", 1, "已登记", _
                                                                                        2, IIf(Nvl(rsTemp!报告操作) <> "", "处理中", _
                                                                                                IIf(Nvl(rsTemp!报告人) = "", "已报到", "报告中")), _
                                                                                        3, IIf(Nvl(rsTemp!报告操作) <> "", "处理中", _
                                                                                                IIf(Nvl(rsTemp!报告人) = "", "已检查", "报告中")), _
                                                                                        4, IIf(Nvl(rsTemp!报告操作) <> "", "处理中", _
                                                                                                IIf(Nvl(rsTemp!复核人) <> "", "审核中", "已报告")), _
                                                                                        5, "已审核", "已完成"))
            .TextMatrix(.Row, GetCN("性别")) = Nvl(rsTemp!性别)
            .TextMatrix(.Row, GetCN("年龄")) = Nvl(rsTemp!年龄)
            If InStr(Nvl(rsTemp!医嘱内容), ":") > 0 Then '新的模式保存在医嘱内容中信息是 名称,执行标记:部位(方法,方法),部位---
                .TextMatrix(.Row, GetCN("医嘱内容")) = Split(rsTemp!医嘱内容, ":")(0)
                .TextMatrix(.Row, GetCN("部位方法")) = Split(rsTemp!医嘱内容, ":")(1)
            Else
                .TextMatrix(.Row, GetCN("医嘱内容")) = Nvl(rsTemp!医嘱内容)
            End If
            .TextMatrix(.Row, GetCN("执行间")) = Nvl(rsTemp!执行间)
            .TextMatrix(.Row, GetCN("报到时间")) = Nvl(rsTemp!报到时间)
            .TextMatrix(.Row, GetCN("申请时间")) = Nvl(rsTemp!申请时间)
            .TextMatrix(.Row, GetCN("开嘱医生")) = Nvl(rsTemp!开嘱医生)
            .TextMatrix(.Row, GetCN("身高")) = Nvl(rsTemp!身高)
            .TextMatrix(.Row, GetCN("体重")) = Nvl(rsTemp!体重)
            .TextMatrix(.Row, GetCN("婴儿")) = Nvl(rsTemp!婴儿)
            .TextMatrix(.Row, GetCN("登记人")) = Nvl(rsTemp!登记人)
            .TextMatrix(.Row, GetCN("报到人")) = Nvl(rsTemp!报到人)
            .TextMatrix(.Row, GetCN("完成人")) = Nvl(rsTemp!完成人)
            .TextMatrix(.Row, GetCN("打印胶片")) = IIf(Nvl(rsTemp!是否打印) = 1, "已打印", "未打印")
            .TextMatrix(.Row, GetCN("报告操作")) = Nvl(rsTemp!报告操作)
            .TextMatrix(.Row, GetCN("绿色通道")) = Nvl(rsTemp!绿色通道)
            .TextMatrix(.Row, GetCN("报告打印")) = IIf(Nvl(rsTemp!报告打印) = 1, "已打印", "未打印")
            .TextMatrix(.Row, GetCN("报告人")) = Nvl(rsTemp!报告人)
            .TextMatrix(.Row, GetCN("复核人")) = Nvl(rsTemp!复核人)
            .TextMatrix(.Row, GetCN("检查技师")) = Nvl(rsTemp!检查技师)
            .TextMatrix(.Row, GetCN("采图时间")) = Nvl(rsTemp!采图时间)
            .TextMatrix(.Row, GetCN("影像类别")) = Nvl(rsTemp!影像类别)
            .TextMatrix(.Row, GetCN("病人ID")) = Nvl(rsTemp!病人ID, 0)
            .TextMatrix(.Row, GetCN("主页ID")) = Nvl(rsTemp!主页ID, 0)
            .TextMatrix(.Row, GetCN("挂号单")) = Nvl(rsTemp!挂号单)
            .TextMatrix(.Row, GetCN("病人科室ID")) = Nvl(rsTemp!病人科室ID, 0)
            .TextMatrix(.Row, GetCN("医嘱ID")) = Nvl(rsTemp!医嘱ID)
            .TextMatrix(.Row, GetCN("发送号")) = Nvl(rsTemp!发送号)
            .TextMatrix(.Row, GetCN("检查UID")) = Nvl(rsTemp!检查uid)
            .TextMatrix(.Row, GetCN("检查状态")) = Nvl(rsTemp!检查过程)
            .TextMatrix(.Row, GetCN("随访描述")) = Nvl(rsTemp!随访描述)
            .TextMatrix(.Row, GetCN("NO")) = Nvl(rsTemp!NO)
            .TextMatrix(.Row, GetCN("转出")) = Nvl(rsTemp!转出)
            .TextMatrix(.Row, GetCN("床号")) = Nvl(rsTemp!当前床号)
            .TextMatrix(.Row, GetCN("当前病区ID")) = Nvl(rsTemp!当前病区ID, 0)
            .TextMatrix(.Row, GetCN("标识号")) = Nvl(rsTemp!标识号)
            .TextMatrix(.Row, GetCN("报告发放")) = IIf(Nvl(rsTemp!报告发放, 0) = 0, "未发放", "已发放")
            .TextMatrix(.Row, GetCN("诊断分类")) = Nvl(rsTemp!诊断分类)
            .TextMatrix(.Row, GetCN("执行科室ID")) = Nvl(rsTemp!执行科室ID)
            .TextMatrix(.Row, GetCN("关联ID")) = Nvl(rsTemp!关联ID, 0)
            .TextMatrix(.Row, GetCN("病人科室")) = Nvl(rsTemp!病人科室)
            .TextMatrix(.Row, GetCN("就诊卡号")) = Nvl(rsTemp!就诊卡号)
            .TextMatrix(.Row, GetCN("单据号")) = Nvl(rsTemp!单据号)
            .TextMatrix(.Row, GetCN("身份证号")) = Nvl(rsTemp!身份证号)
            
            If Nvl(rsTemp!婴儿) <> 0 Then
                gstrSQL = "Select Nvl(A.婴儿姓名, B.姓名 || '之子' || Trim(To_Char(A.序号, '9'))) As 婴儿姓名, 婴儿性别, 出生时间" & vbNewLine & _
                            "From 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
                            "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id And A.序号 = [3]"

                Set rsBaby = zldatabase.OpenSQLRecord(gstrSQL, "提取婴儿信息", CLng(rsTemp!病人ID), CLng(Nvl(rsTemp!主页ID, 0)), CLng(rsTemp!婴儿))
                If Not rsBaby.EOF Then
                    .TextMatrix(.Row, GetCN("姓名")) = rsBaby!婴儿姓名
                    .TextMatrix(.Row, GetCN("性别")) = Nvl(rsBaby!婴儿性别)
                    .TextMatrix(.Row, GetCN("年龄")) = Nvl(rsBaby!出生时间)
                End If
            End If
            
            If .TextMatrix(.Row, GetCN("检查过程")) = "已拒绝" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已拒绝
            If .TextMatrix(.Row, GetCN("检查过程")) = "已完成" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已完成
            If .TextMatrix(.Row, GetCN("检查过程")) = "已报到" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已报到
            If .TextMatrix(.Row, GetCN("检查过程")) = "已登记" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已登记
            If .TextMatrix(.Row, GetCN("检查过程")) = "已检查" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已检查
            If .TextMatrix(.Row, GetCN("检查过程")) = "已审核" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已审核
            If .TextMatrix(.Row, GetCN("检查过程")) = "处理中" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor处理中
            If .TextMatrix(.Row, GetCN("检查过程")) = "报告中" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor报告中
            If .TextMatrix(.Row, GetCN("检查过程")) = "审核中" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor审核中
            If .TextMatrix(.Row, GetCN("检查过程")) = "已报告" Then .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = gdblColor已报告
            
            rsTemp.MoveNext
        Loop
    End With
    
    '如果没有路径中病人，则不显示路径列
    intPathColNum = GetCN("路径")
    If blnShowPath = False Then
        vsList.ColWidth(intPathColNum) = 0
    Else
        vsList.ColWidth(intPathColNum) = GetCW("路径")
    End If
    
    '恢复排序
    If mlngSortCol <> 0 And mintSortOrder <> 0 Then
        If mlngSortCol < vsList.Cols Then
            vsList.Col = mlngSortCol
            vsList.Sort = mintSortOrder
        End If
    End If
    
    stbThis.Panels(2).Text = "共 " & vsList.Rows - 1 & " 条记录": stbThis.Panels(2).Alignment = sbrCenter
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub PicWindow_Resize()
    On Error Resume Next
    With picInfo
        .Top = 0
        .Left = 0
        .Width = PicWindow.ScaleWidth
    End With
        
    With TabWindow
        .Top = picInfo.ScaleHeight
        .Left = 0
        .Width = PicWindow.ScaleWidth
        .Height = PicWindow.ScaleHeight - picInfo.ScaleHeight
    End With
End Sub
Private Sub TabWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not mblnInitOk Then Exit Sub

    On Error GoTo errHandle
    If mblnIsHistory Then
        RefreshTabWindow mlngHOrderID
    ElseIf Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0 Then
        RefreshTabWindow 0, True
    Else
        RefreshTabWindow 0, False, True
        If vsList.TextMatrix(vsList.Row, GetCN("检查UID")) = "" And mfrmPACSImg.lvwSeq.ListItems.Count > 0 Then '如果经过刷新后发现有图了，则刷新病人列表，目的是为了让观片等按键可用
            vsList.TextMatrix(vsList.Row, GetCN("检查UID")) = mfrmPACSImg.lvwSeq.Tag
        End If
    End If
    
    '删除现在的工具栏及顶级菜单项
    Call LockWindowUpdate(Me.hWnd)
    Dim lngCount As Long
    For lngCount = cbrMain.ActiveMenuBar.Controls.Count To 1 Step -1
        cbrMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbrMain.Count To 2 Step -1
        cbrMain(lngCount).Delete
    Next
    Call InitCommandBars
    
    Select Case Item.Tag
        Case "PACS报告"     '使用PACS报告编辑器
            mfrmPacsReport.zlDefCommandBars Me.cbrMain
        Case "病历报告"     '使用电子病历编辑器写报告
            mobjReport.zlDefCommandBars Me.cbrMain
        Case "申请费用"
            mobjExpense.zlDefCommandBars Me, Me.cbrMain
        Case "住院医嘱"
            mobjInAdvice.zlDefCommandBars Me, Me.cbrMain, 2
        Case "门诊医嘱"
            mobjOutAdvice.zlDefCommandBars Me, Me.cbrMain, 2
        Case "住院病历"
            mobjInEPRs.zlDefCommandBars cbrMain
        Case "门诊病历"
            mobjOutEPRs.zlDefCommandBars cbrMain
        Case "排队叫号"
            If Not Not mobjQueue Is Nothing Then
                mobjQueue.zlDefCommandBars cbrMain
            End If
    End Select
    
    If Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) <> 0 Then
        '显示可打印的诊疗单据:之所以即时加载,是为了使用F2热键
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))
    End If
    
    Call LockWindowUpdate(0)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub TimerRefresh_Timer()
    '刷新病人列表
    Call RefreshList
End Sub

Private Sub txtFilter_Change()
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (txtFilter.Text = "" And Me.ActiveControl Is txtFilter)
    End If
    Call subRefreshFilterCondition(txtFilter.Text)
End Sub

Private Sub txtFilter_GotFocus()
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard         '身份证识别对象
    
    If txtFilter.Text <> "" Then Call zlControl.TxtSelAll(txtFilter)
    If InStr(mstrCurFindtype, "姓  名") > 0 Then
        Call zlCommFun.OpenIme(True)
    End If
    
    If Not mobjIDCard Is Nothing And txtFilter.Text = "" Then '启动身份证读卡设备
        mobjIDCard.SetEnabled (True)
    End If
End Sub
Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtFilter_Validate(False)
        Call zlControl.TxtSelAll(txtFilter)
    End If
End Sub

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        Select Case mstrCurFindtype
            Case "门诊号", "住院号"
                If InStr("*+0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "门诊号", "住院号", "检查号"
                If Len(txtFilter.Text) > 18 Then KeyAscii = 0 '超长
            Case "就诊卡"
                Dim blnCard As Boolean
    
                '去掉磁卡的其他的特殊字符
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = zlCommFun.InputIsCard(Me.txtFilter, KeyAscii, glngSys)
                
                '刷卡完成或确认输入
                If blnCard And Len(Me.txtFilter.Text) = Val(gbytCardLen) - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtFilter.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.txtFilter.Text = Me.txtFilter.Text & Chr(KeyAscii)
                        Me.txtFilter.SelStart = Len(Me.txtFilter.Text)
                    End If
                    KeyAscii = 0
                    Me.txtFilter.Text = UCase(Me.txtFilter)
                    Me.txtFilter.SetFocus
                End If
            Case "单据号"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtFilter.Text = "" Or txtFilter.SelLength = Len(txtFilter.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "姓名"
            
        End Select
    Else
        If Trim(txtFilter.Text) <> "" Then
            If Mid(txtFilter.Text, 1, 1) = "*" And IsNumeric(Mid(txtFilter.Text, 2)) = True Then mstrCurFindtype = "门诊号"
            If Mid(txtFilter.Text, 1, 1) = "+" Then mstrCurFindtype = "住院号"
        End If
        Dim cbrControl As CommandBarControl
        Set cbrControl = cbrdock.FindControl(, ID_开始查找)
        If Not cbrControl Is Nothing Then
            cbrdock_Execute cbrControl
        End If
    End If
End Sub

Private Sub txtFilter_LostFocus()
    Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
    End If
End Sub

Private Sub txtFilter_Validate(Cancel As Boolean)
    If InStr(mstrCurFindtype, "单据号") > 0 Then
        If IsNumeric(txtFilter.Text) Then
            txtFilter.Text = GetFullNO(txtFilter.Text, 0)
        End If
    End If
End Sub

Private Sub SeekNextPati(ByVal blnFirst As Boolean)
'------------------------------------------------
'功能：在病人列表中定位指定的记录
'参数： blnFirst -- 是否第一次查找
'返回：无，直接在病人列表中定位
'------------------------------------------------
    Dim blnOK As Boolean, lngCount As Long, intB As Integer
    Dim lngRow As Long

    '如果没有记录，则退出
    If Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0 Then Exit Sub

    intB = 0
    
    On Error GoTo err
    
    If Not blnFirst Then
        intB = vsList.Row + 1
        If intB >= vsList.Rows Then intB = 1
    End If

    blnOK = False
    For lngCount = intB To vsList.Rows - 1 '在当前状态中查找
        Select Case mstrLocateType
            Case "标识号"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("标识号")), 0) Like txtLocate.Text & "*" Then blnOK = True
            Case "就诊卡", "ＩＣ卡"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("就诊卡号")), 0) Like txtLocate.Text & "*" Then blnOK = True
            Case "单据号"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("NO")), 0) Like txtLocate.Text & "*" Then blnOK = True
            Case "检查号"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("检查号")), 0) Like txtLocate.Text & "*" Then blnOK = True
            Case "姓名"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("姓名")), "") Like txtLocate.Text & "*" Then blnOK = True
                If zlCommFun.SpellCode(Nvl(vsList.TextMatrix(lngCount, GetCN("姓名")), "")) Like UCase(txtLocate.Text) & "*" Then blnOK = True
            Case "身份证"
                If Nvl(vsList.TextMatrix(lngCount, GetCN("身份证号")), 0) Like txtLocate.Text & "*" Then blnOK = True
        End Select
        
        If blnOK Then
            txtLocate.Tag = txtLocate.Text
            On Error Resume Next
            '计算当前行和顶行之间的差距
            lngRow = Abs(vsList.Row - vsList.TopRow)
            
            vsList.Row = lngCount
            vsList.TopRow = vsList.Row - lngRow
            
            Exit Sub
        End If
    Next
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_更换检查设备()
Dim strModality As String
Dim rResult As VbMsgBoxResult
Dim strSQL As String
    
    frmChangeDevice.ShowMe UCase(vsList.TextMatrix(vsList.Row, GetCN("影像类别"))), Me
    strModality = frmChangeDevice.strDeviceType
    
     If strModality <> "" Then
         strSQL = "Zl_影像检查_影像类别(" & vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) & "," & vsList.TextMatrix(vsList.Row, GetCN("发送号")) & ",'" & strModality & "')"
         zldatabase.ExecuteProcedure strSQL, Me.Caption
     End If
     
     '刷新病人列表
     Call RefreshList
End Sub

Private Sub sub3DProcess(strCommand As String, strImageDir As String)
    Dim str3DCommand As String
    
    '组织三维重建语句
    str3DCommand = mstr3DExeDir & " " & mstr3DPara & " " & strCommand & " " & strImageDir
    On Error Resume Next
    Shell str3DCommand
End Sub

Private Sub sub三维重建(strCommand As String)
    Dim strImageDir As String
    
    If TabWindow.Selected.Tag <> "影像图象" Then '起到刷新图像作用
        Call mfrmPACSImg.zlRefresh(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), vsList.TextMatrix(vsList.Row, GetCN("发送号")), mstrPrivs, vsList.TextMatrix(vsList.Row, GetCN("转出")) = 1)
    End If
    
    '组织三维重建需要的图像
    strImageDir = mfrmPACSImg.ZLfun3DImgProcess
    If strImageDir <> "" Then
        Call sub3DProcess(strCommand, strImageDir)
    End If
End Sub


Private Sub Menu_Manage_随访()
Dim strReview As String
Dim strDeptName As String

    On Error GoTo errHandle
    
    strDeptName = Split(mstrCur科室, "-")(1)
    If frmReview.ShowMe(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), vsList.TextMatrix(vsList.Row, GetCN("发送号")), _
        Me, strDeptName, strReview) = True Then
        vsList.TextMatrix(vsList.Row, GetCN("随访描述")) = strReview
    End If

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function zlInQueue(str队列名称 As String, lng业务类型 As Long, lng业务ID As Long, lng科室ID As Long, _
        str患者姓名 As String, lng病人ID As Long, str诊室 As String, str医生姓名 As String, Optional str排队标记 As String = "", Optional str排队号码 As String = "") As Boolean
        
        On Error GoTo err
        
        If mblnUseQueue = True And Not mobjQueue Is Nothing Then
            mobjQueue.zlInQueue str队列名称, lng业务类型, lng业务ID, lng科室ID, str患者姓名, lng病人ID, str诊室, str医生姓名, str排队标记, str排队号码
        End If
        zlInQueue = True
        Exit Function
err:
    If ErrCenter = 1 Then
        Resume
    End If
        
End Function

Private Sub txtLocate_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtLocate.Text = "" And Me.ActiveControl Is txtLocate)
    If txtLocate.Text = "" Then txtLocate.Tag = ""
End Sub

Private Sub txtLocate_GotFocus()
    If mobjIDCard Is Nothing Then Set mobjIDCard = New clsIDCard         '身份证识别对象
    
    If txtLocate.Text <> "" Then Call zlControl.TxtSelAll(txtLocate)
    If mstrLocateType = "姓名" Then
        Call zlCommFun.OpenIme(True)
    End If
    If Not mobjIDCard Is Nothing And txtLocate.Text = "" Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtLocate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call txtLocate_Validate(False)
        Call zlControl.TxtSelAll(txtLocate)
        Call SeekNextPati(txtLocate.Tag <> txtLocate.Text)
    End If
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        Select Case mstrLocateType
            Case "标识号"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "就诊卡"
                Dim blnCard As Boolean
    
                '去掉磁卡的其他的特殊字符
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
                
                blnCard = zlCommFun.InputIsCard(Me.txtLocate, KeyAscii, glngSys)
                
                '刷卡完成或确认输入
                If blnCard And Len(Me.txtLocate.Text) = Val(gbytCardLen) - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtLocate.Text <> "" Then
                    If KeyAscii <> 13 Then
                        Me.txtLocate.Text = Me.txtLocate.Text & Chr(KeyAscii)
                        Me.txtLocate.SelStart = Len(Me.txtLocate.Text)
                    End If
                    KeyAscii = 0
                    Me.txtLocate.Text = UCase(Me.txtLocate)
                    Me.txtLocate.SetFocus
                End If
            Case "单据号"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Not (txtLocate.Text = "" Or txtLocate.SelLength = Len(txtLocate.Text)) _
                    And InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            Case "姓名"
            
        End Select
    End If
End Sub

Private Sub txtLocate_LostFocus()
    Call zlCommFun.OpenIme
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
End Sub

Private Sub txtLocate_Validate(Cancel As Boolean)
    If InStr(mstrLocateType, "单据号") > 0 Then
        If IsNumeric(txtLocate.Text) Then
            txtLocate.Text = GetFullNO(txtLocate.Text, 0)
        End If
    End If
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
Dim i As Integer, strCol As String
    For i = 0 To vsList.Cols - 1
        strCol = strCol & "|" & vsList.Cell(flexcpData, 0, i) & ";" & vsList.ColWidth(i)
    Next
    mstrCol = Mid(strCol, 2)
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能: 显示病人卡片按钮
    If vsList.TextMatrix(NewRow, GetCN("医嘱ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If vsList.LeftCol > GetCN("姓名") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, NewRow, GetCN("姓名")) + vsList.Cell(flexcpWidth, NewRow, GetCN("姓名")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("姓名")) + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub
Private Sub vsList_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'功能:显示病人卡片按钮
    If vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If NewLeftCol > GetCN("姓名") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, GetCN("姓名")) + vsList.Cell(flexcpWidth, vsList.Row, GetCN("姓名")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("姓名")) + 15
            cmdInfo.Visible = True
        End If
    End If
End Sub

Private Sub vsList_AfterSort(ByVal Col As Long, Order As Integer)
    mlngSortCol = Col
    mintSortOrder = Order
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'功能:显示病人卡片按钮
    If vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) = "" Then
        cmdInfo.Visible = False
    Else
        If vsList.LeftCol > GetCN("姓名") Then
            cmdInfo.Visible = False
        Else
            cmdInfo.Left = vsList.Cell(flexcpLeft, vsList.Row, GetCN("姓名")) + vsList.Cell(flexcpWidth, vsList.Row, GetCN("姓名")) - cmdInfo.Width - 15
            cmdInfo.Top = vsList.Cell(flexcpTop, vsList.Row, GetCN("姓名")) + 15
            cmdInfo.Visible = True
        End If
    End If
    
    Dim i As Integer, strCol As String
    For i = 0 To vsList.Cols - 1 '暂存列序列宽，窗体关闭时存于注册表
        strCol = strCol & "|" & vsList.Cell(flexcpData, 0, i) & ";" & vsList.ColWidth(i)
    Next
    mstrCol = Mid(strCol, 2)
End Sub

Private Sub vsList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < GetCN("姓名") Then Cancel = True
End Sub

Private Sub vsList_DblClick()
    If vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")) <> "" Then
        Select Case vsList.TextMatrix(vsList.Row, GetCN("检查状态"))
            Case 1, 0
                Call Menu_Manage_报到
            Case 2, 3               '双击打开书写报告,报告打开时跟据设定是否打开观片站
                Call Menu_RichEPR(IIf(mblnPacsReport, conMenu_PacsReport_Open, conMenu_Edit_Modify))
            Case 4, 5               '双击修订报告,报告打开时跟据设定是否打开观片站
                Call Menu_RichEPR(IIf(mblnPacsReport, conMenu_PacsReport_Open, conMenu_Edit_Audit))
            Case 6                  '查阅,PACS报告编辑器时，直接打开观片站，病历编辑器，打开查阅报告和观片站，
                                    '没有打开观片站，因此需要手工打开观片站
                If mblnPacsReport = False Then Call Menu_RichEPR(conMenu_File_Open)
                Call OpenViewerWithReport
        End Select
    End If
End Sub
Private Sub vsList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim control As CommandBarControl, Menucontrol As CommandBarControl
        Dim Popup As CommandBar
        Set Popup = cbrMain.Add("右键菜单", xtpBarPopup)
        For Each Menucontrol In cbrMain.ActiveMenuBar.Controls
'            If Menucontrol.Parent.BarID = conMenu_ManagePopup Then
            If (Menucontrol.ID <> conMenu_FilePopup And Menucontrol.ID <> conMenu_ToolPopup _
                And Menucontrol.ID <> conMenu_ViewPopup And Menucontrol.ID <> conMenu_HelpPopup) And Menucontrol.Type = xtpControlPopup Then
                For Each control In Menucontrol.CommandBar.Controls
                    control.Copy Popup
                Next
            End If
        Next
        Popup.ShowPopup
    End If
End Sub

Private Sub vsList_RowColChange()
    On Error GoTo errHandle
    mblnIsHistory = False
    If mblnvsRefresh Then Exit Sub
    '判断嵌入式报告编辑器中的报告是否没有保存
    If mblnPacsReport = True Then    '使用PACS报告编辑器
        Call mfrmPacsReport.PromptModify
    End If
    
    If Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0 Then  '无记录时处理
        Call RefreshTabWindow(0, True)
        cboTimes.Clear
        txtAppend = ""
        lbl个人信息.Caption = "姓  名:" & Space(12) & "性  别:" & Space(10) & "年  龄:" & Space(10)
        lbl检查信息.Caption = "检查号:" & Space(17) & "病人科室:" & Space(15) & "标识号:" & Space(12) & "床  号:" & Space(10)
        lblCash.Visible = False
    Else
        '检查当前的执行科室ID，如果科室更换了，则修改科室对应的参数
        If mlngCur科室ID <> vsList.TextMatrix(vsList.Row, GetCN("执行科室ID")) Then
            mlngCur科室ID = vsList.TextMatrix(vsList.Row, GetCN("执行科室ID"))
            mstrCur科室 = GetDeptName(mlngCur科室ID, mstrCanUse科室)
            Call subChangeDept(1)
        End If
        
        Call FillHistory '填充历次检查记录
        Call FillTxtInfor '填充右上方病人基本信息
        Call FillTxtAppend '填充左下角医嘱附件
        Call ShowTab '根据病人提供不同选项卡
        
        Call ShowBillList(cbrMain.FindControl(, conMenu_Manage_RequestPrint, , True))  '显示可打印的诊疗单据:之所以即时加载,是为了使用F2热键
        
        If mstrFirstTab <> "" Then '不为空表示按定制首页显示,由TabWindow调用刷新
            Dim i As Integer
            For i = 0 To TabWindow.ItemCount - 1
                If InStr(TabWindow.Item(i).Tag, mstrFirstTab) > 0 And TabWindow.Item(i).Visible Then
                    If TabWindow.Item(i).Selected Then
                        Call RefreshTabWindow
                    Else
                        TabWindow.Item(i).Selected = True
                    End If
                    Exit Sub
                End If
            Next
            If i = TabWindow.ItemCount Then TabWindow(0).Selected = True '没循环到了触发第0个tab
        Else
            Call RefreshTabWindow
        End If
        
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub FillTxtInfor(Optional lngAdviceID As Long = 0)
'填充右上方病人基本信息
Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    
    With vsList
        lbl个人信息.Caption = "姓  名:" & Rpad(.TextMatrix(.Row, GetCN("姓名")), 12, " ") & "性  别:" & Rpad(.TextMatrix(.Row, GetCN("性别")), 10, " ") & _
                          "年  龄:" & Rpad(.TextMatrix(.Row, GetCN("年龄")), 10, " ")
                          
        If lngAdviceID = 0 Then '---------------------------非历次检查直接用列表中记录填充
            gstrSQL = "Select 名称 From 部门表 Where ID=[1]"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取病人科室", CLng(.TextMatrix(.Row, GetCN("病人科室ID"))))
            lbl检查信息.Caption = "检查号:" & Rpad(.TextMatrix(.Row, GetCN("检查号")), 17, " ") & "病人科室:" & Rpad(rsTemp!名称, 15, " ") & _
                          "标识号:" & Rpad(.TextMatrix(.Row, GetCN("标识号")), 12, " ") & _
                          "床  号:" & Rpad(.TextMatrix(.Row, GetCN("床号")) & "", 10, " ")
                                  
            lblCash.Caption = "收": lblCash.Visible = False
            lblCash.Visible = (.TextMatrix(.Row, GetCN("收费")) = " ")
        Else
            Dim strSQLBak As String
            gstrSQL = "Select A.ID, A.病人科室id, A.开嘱医生,A.病人来源, A.医嘱内容, Nvl(A.婴儿, 0) 婴儿,A.病人id, " & vbNewLine & _
                        " A.主页id, A.挂号单, B.检查号, B.检查uid, C.名称, D.发送号,D.执行状态,0 as 转出,A.执行科室ID " & vbNewLine & _
                        "From 病人医嘱记录 A, 影像检查记录 B, 部门表 C, 病人医嘱发送 D" & vbNewLine & _
                        "Where A.ID = [1] And A.ID = B.医嘱id And A.病人科室id = C.ID And A.ID = D.医嘱id"
            strSQLBak = gstrSQL
            strSQLBak = Replace(strSQLBak, "病人医嘱记录", "H病人医嘱记录")
            strSQLBak = Replace(strSQLBak, "病人医嘱发送", "H病人医嘱发送")
            strSQLBak = Replace(strSQLBak, "影像检查记录", "H影像检查记录")
            strSQLBak = Replace(strSQLBak, "0 as 转出", "1 as 转出")
            gstrSQL = gstrSQL & vbNewLine & " Union ALL " & strSQLBak
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "查历次记录信息", lngAdviceID)
            If Not rsTemp.EOF Then
                mlngHOrderID = lngAdviceID
                mlngHSendNo = Nvl(rsTemp!发送号, 0)
                mstrHStudyUID = Nvl(rsTemp!检查uid)
                mblnHMoved = IIf(rsTemp!转出 = 1, True, False)
                fraInfo.Tag = rsTemp!病人ID & "|" & rsTemp!主页ID & "|" & rsTemp!ID & "|" & rsTemp!发送号 _
                            & "|" & rsTemp!病人科室ID & "|" & rsTemp!挂号单 & "|" & Nvl(rsTemp!病人来源, 3) _
                            & "|" & rsTemp!检查uid & "|" & rsTemp!转出 & "|" & rsTemp!执行状态 & "|" & rsTemp!执行科室ID
                            
                lbl检查信息.Caption = "检查号:" & Rpad(Nvl(rsTemp!检查号), 17, " ") & "病人科室:" & Rpad(rsTemp!名称, 15, " ") & _
                                      "标识号:" & Rpad(.TextMatrix(.Row, GetCN("标识号")), 12, " ") & _
                                      "床  号:" & Rpad(.TextMatrix(.Row, GetCN("床号")) & "", 10, " ")
                
                If rsTemp!婴儿 <> 0 Then
                    Dim lngBaby As Integer, lngPatID As Long, lngPageID As Long
                    lngBaby = rsTemp!婴儿: lngPatID = rsTemp!病人ID: lngPageID = Nvl(rsTemp!主页ID, 0)
                    gstrSQL = "Select Nvl(A.婴儿姓名, B.姓名 || '之子' || Trim(To_Char(A.序号, '9'))) As 婴儿姓名, 婴儿性别, 出生时间" & vbNewLine & _
                            "From 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
                            "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id And A.序号 = [3]"
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取婴儿信息", lngPatID, lngPageID, lngBaby)
                    If Not rsTemp.EOF Then
                        lbl个人信息.Caption = "姓  名:" & Rpad(rsTemp!婴儿姓名, 12, " ") & "性  别:" & Rpad(rsTemp!婴儿性别, 10, " ") & _
                                            "年  龄:" & Rpad(rsTemp!出生时间, 10, " ")
                    End If
                End If
            Else
                lbl检查信息.Caption = "检查号:" & Space(17) & "病人科室:" & Space(15) & "标识号:" & Space(12) & "床  号:" & Space(10)
            End If
            lblCash.Caption = "历": lblCash.Visible = True
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillTxtAppend(Optional lngAdviceIDtmp As Long = 0)
'填充左下角医嘱附件
Dim lngAdviceID As Long, strAppend As String, rsTemp As ADODB.Recordset, i As Integer
    On Error GoTo errHandle
    With vsList
        If lngAdviceIDtmp = 0 Then
            lngAdviceID = Val(.TextMatrix(.Row, GetCN("医嘱ID")))
        Else
            lngAdviceID = lngAdviceIDtmp
        End If
        
        If lngAdviceIDtmp = 0 Then '-------------------------------------------列表选择调用
            txtAppend = "检查项目:" & .TextMatrix(.Row, GetCN("医嘱内容")) & vbCrLf
            txtAppend = txtAppend & "开嘱医生:" & Rpad(.TextMatrix(.Row, GetCN("开嘱医生")), 8, " ") & vbCrLf
            
            If .TextMatrix(.Row, GetCN("部位方法")) <> "" Then
                For i = 0 To UBound(Split(.TextMatrix(.Row, GetCN("部位方法")), "),"))
                    If i = 0 Then
                        txtAppend = txtAppend & "检查部位:" & vbCrLf & Space(2) & "1:" & Split(.TextMatrix(.Row, GetCN("部位方法")), "),")(i) & ")"
                    Else
                        txtAppend = txtAppend & vbCrLf & Space(2) & i + 1 & ":" & Split(.TextMatrix(.Row, GetCN("部位方法")), "),")(i) & ")"
                    End If
                Next
                If Trim(txtAppend) <> "" Then txtAppend = Mid(txtAppend, 1, Len(txtAppend) - 1) '取掉最后的括号
            Else
                txtAppend = txtAppend & "检查部位:" & .TextMatrix(.Row, GetCN("医嘱内容"))
            End If
            gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
            If .TextMatrix(.Row, GetCN("转出")) = 1 Then gstrSQL = Replace(gstrSQL, "病人医嘱附件", "H病人医嘱附件")
        Else                    '-------------------------------------------历次记录选择调用
            Dim strTemp As String
            txtAppend = ""
            
            gstrSQL = "Select 开嘱医生,医嘱内容 From 病人医嘱记录 Where  id =[1]"
            If Split(fraInfo.Tag, "|")(8) = 1 Then gstrSQL = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取医嘱内容", lngAdviceID)
            
            If rsTemp.EOF = False Then
                strTemp = Nvl(rsTemp!医嘱内容)
                If InStr(strTemp, ":") > 0 Then
                    txtAppend = "检查项目:" & Split(strTemp, ":")(0) & vbCrLf
                Else
                    txtAppend = "检查项目:" & strTemp & vbCrLf
                End If
                
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
            gstrSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列" '根据历次记录是否转移判断查历史表
            If Split(fraInfo.Tag, "|")(8) = 1 Then gstrSQL = Replace(gstrSQL, "病人医嘱附件", "H病人医嘱附件")
        End If
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取病人附件", lngAdviceID)
        Do Until rsTemp.EOF
            strAppend = strAppend & rsTemp!项目 & ":" & Nvl(rsTemp!内容) & vbCrLf
            rsTemp.MoveNext
        Loop
        
        txtAppend = txtAppend & vbCrLf & vbCrLf & strAppend
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub FillHistory()
'填充历次检查记录
Dim rsTemp As ADODB.Recordset, strTemp As String
    On Error GoTo errHandle
    With vsList
        cboTimes.Tag = "" 'cbotime下拉时用到，用于区别是"增加项目"时触发还是"点击cbotimes"触发
        gstrSQL = "Select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                   " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 C" & _
                   " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID " & _
                   "" & IIf(.TextMatrix(.Row, GetCN("检查过程")) = "已拒绝", "", " And B.执行状态<>2 ") & _
                   " AND A.ID=C.医嘱ID"
        '是否选择了全部科室
        If mblnAllDepts = True Then
            gstrSQL = gstrSQL & " And Instr( [3],A.执行科室id ) >0 "
        Else
            gstrSQL = gstrSQL & " And A.执行科室id+0 =[2] "
        End If
        
        '启用关联病人，才查询关联ID
        If mblnRelatingPatient = True And .TextMatrix(.Row, GetCN("关联ID")) <> 0 Then
            gstrSQL = gstrSQL & " union select A.ID 医嘱ID,A.开嘱时间  开嘱时间,A.医嘱内容 " & _
                    " From 病人医嘱记录 A " & _
                    " Where A.id in (Select 医嘱ID from 影像检查记录 Where 关联ID =[4]) "
        End If
        
        strTemp = Replace(gstrSQL, "病人医嘱记录", "H病人医嘱记录")
        strTemp = Replace(strTemp, "病人医嘱发送", "H病人医嘱发送")
        strTemp = Replace(strTemp, "影像检查记录", "H影像检查记录")
        gstrSQL = gstrSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
        gstrSQL = "Select * From (" & vbNewLine & gstrSQL & vbNewLine & ") Order By 开嘱时间 Asc"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "", CLng(.TextMatrix(.Row, GetCN("病人ID"))), _
                mlngCur科室ID, mstrCanUse科室IDs, CLng(.TextMatrix(.Row, GetCN("关联ID"))))
        
        cboTimes.Clear
        Do Until rsTemp.EOF
           cboTimes.AddItem "第" & rsTemp.AbsolutePosition & "次(" & Format(rsTemp!开嘱时间, "yyyy-mm-dd") & ")  " & Trim(rsTemp!医嘱内容)
           cboTimes.ItemData(cboTimes.NewIndex) = rsTemp!医嘱ID
           If rsTemp!医嘱ID = .TextMatrix(.Row, GetCN("医嘱ID")) Then cboTimes.ListIndex = cboTimes.NewIndex
           rsTemp.MoveNext
        Loop
        cboTimes.Tag = "完成"
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub ShowTab(Optional lngAdviceID As Long = 0)
'根据病人来源控制病历及医嘱选项卡
Dim int来源 As Integer, i As Integer
Dim strFirstTab As String
Dim intDefaultIndex As Integer

    On Error GoTo errHandle
    
    If lngAdviceID = 0 Then '-------------------------------------------列表选择调用
        int来源 = Val(vsList.Cell(flexcpData, vsList.Row, GetCN("来源")))
        Dim blnShowReport As Boolean
        '判断 无图像不许写报告
        blnShowReport = True
        If mblnReportWithImage = True Then
            If vsList.TextMatrix(vsList.Row, GetCN("检查UID")) = "" Then blnShowReport = False
        End If
    Else                    '-------------------------------------------历次记录选择调用
        '历次记录时fraInfo.Tag = 0病人ID|1主页ID|2医嘱ID|3发送号|4病人科室ID|5挂号单|6病人来源|7检查UID|8转出
        int来源 = Split(fraInfo.Tag, "|")(6)
    End If
    
    If int来源 <> 2 Then '根据病人来源控制病历及医嘱选项卡
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "门诊病历", "门诊医嘱"
                    TabWindow(i).Visible = True
                Case "住院病历", "住院医嘱"
                    TabWindow(i).Visible = False
                Case "影像图象"
                    TabWindow(i).Visible = True
                Case "PACS报告"
                    TabWindow(i).Visible = IIf(lngAdviceID = 0, vsList.TextMatrix(vsList.Row, GetCN("检查状态")) > 1 And blnShowReport, True)
                    If TabWindow(i).Visible = True Then TabWindow(i).Visible = mblnPacsReport
                    '判断是否切换了报告方式
                    If TabWindow.Selected.Tag = "病历报告" And TabWindow(i).Visible = True Then
                        TabWindow(i).Selected = True
                    End If
                Case "病历报告"
                    TabWindow(i).Visible = IIf(lngAdviceID = 0, vsList.TextMatrix(vsList.Row, GetCN("检查状态")) > 1 And blnShowReport, True)
                    If TabWindow(i).Visible = True Then TabWindow(i).Visible = Not mblnPacsReport
                    '判断是否切换了报告方式
                    If TabWindow.Selected.Tag = "PACS报告" And TabWindow(i).Visible = True Then
                        TabWindow(i).Selected = True
                    End If
            End Select
        Next
    Else
        For i = 0 To TabWindow.ItemCount - 1
            Select Case TabWindow(i).Tag
                Case "门诊病历", "门诊医嘱"
                    TabWindow(i).Visible = False
                Case "住院病历", "住院医嘱"
                    TabWindow(i).Visible = True
                Case "影像图象"
                    TabWindow(i).Visible = True
                Case "PACS报告"
                    TabWindow(i).Visible = IIf(lngAdviceID = 0, vsList.TextMatrix(vsList.Row, GetCN("检查状态")) > 1 And blnShowReport, True)
                    If TabWindow(i).Visible = True Then TabWindow(i).Visible = mblnPacsReport
                    '判断是否切换了报告方式
                    If TabWindow.Selected.Tag = "病历报告" And TabWindow(i).Visible = True Then
                        TabWindow(i).Selected = True
                    End If
                Case "病历报告"
                    TabWindow(i).Visible = IIf(lngAdviceID = 0, vsList.TextMatrix(vsList.Row, GetCN("检查状态")) > 1 And blnShowReport, True)
                    If TabWindow(i).Visible = True Then TabWindow(i).Visible = Not mblnPacsReport
                    '判断是否切换了报告方式
                    If TabWindow.Selected.Tag = "PACS报告" And TabWindow(i).Visible = True Then
                        TabWindow(i).Selected = True
                    End If
            End Select
        Next
    End If
    
    '如果当前被选择的页面不可见，则显示用户的主要工作页面
    If TabWindow.Selected.Visible = False Then
        strFirstTab = mstrFirstTab
        If strFirstTab = "" Then strFirstTab = "影像"
        For i = 0 To TabWindow.ItemCount - 1
            If InStr(TabWindow(i).Tag, strFirstTab) > 0 And TabWindow(i).Visible Then
                TabWindow(i).Selected = True
                Exit For
            ElseIf InStr(TabWindow(i).Tag, "影像") > 0 Then
                intDefaultIndex = i
            End If
        Next i
    End If
    If TabWindow.Selected.Visible = False Then
        TabWindow(intDefaultIndex).Visible = True
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub RefreshTabWindow(Optional lngAdviceIDtmp As Long = 0, Optional blnClear As Boolean = False, Optional blnRefresh As Boolean = False)
'lngAdviceIDtmp历次记录时传入 , 其它传0, blnclear清空当前列表, blnRefresh强制刷新
'刷新当前页面,调用：列表选择，历次记录选择，子窗体选择
'历次记录时fraInfo.Tag = 0病人ID|1主页ID|2医嘱ID|3发送号|4病人科室ID|5挂号单|6病人来源|7检查UID|8转出|9执行状态
Dim lngAdviceID As Long, lngSendNO As Long, lngPatID As Long, lngPageID As Long, blnCanPrint As Boolean, blnIsInsidePatient As Boolean
Dim lngUnit As Long, lngPatDept As Long, strRegNo As String, intMoved As Boolean, intState As Integer, i As Integer, intPatientForm As Integer

    On Error GoTo errHandle
    If lngAdviceIDtmp = 0 Then '-----------------------列表选择调用
        If blnClear Then       '无记录时清空所有子窗体
            lngAdviceID = 0: lngSendNO = 0: lngPatID = 0: lngPageID = 0
            lngPatDept = 0: strRegNo = "": intMoved = 0: intState = 0: lngUnit = 0: blnCanPrint = False
        Else
            With vsList
                lngAdviceID = .TextMatrix(.Row, GetCN("医嘱ID")): lngSendNO = .TextMatrix(.Row, GetCN("发送号"))
                lngPatID = .TextMatrix(.Row, GetCN("病人ID")): lngPageID = Val(.TextMatrix(.Row, GetCN("主页ID")))
                lngPatDept = .TextMatrix(.Row, GetCN("病人科室ID")): strRegNo = .TextMatrix(.Row, GetCN("挂号单"))
                intMoved = .TextMatrix(.Row, GetCN("转出"))
                intState = IIf(.TextMatrix(.Row, GetCN("检查过程")) = "已拒绝", 2, IIf(.TextMatrix(.Row, GetCN("检查过程")) = "已完成", 1, 3))
                lngUnit = Val(.TextMatrix(.Row, GetCN("当前病区ID")))
                blnCanPrint = IIf(mblnCanPrint, IIf(.Cell(flexcpData, .Row, GetCN("紧急")) = 1, .TextMatrix(.Row, GetCN("报告人")) <> "", .TextMatrix(.Row, GetCN("复核人")) <> ""), True)
                intPatientForm = Decode(.TextMatrix(.Row, GetCN("来源")), "门", 1, "住", 2, "外", 3, 4)
            End With
        End If
    Else                       '----------------------历次记录选择调用
        lngAdviceID = lngAdviceIDtmp: lngSendNO = Split(fraInfo.Tag, "|")(3)
        lngPatID = Split(fraInfo.Tag, "|")(0): lngPageID = Val(Split(fraInfo.Tag, "|")(1))
        lngPatDept = Split(fraInfo.Tag, "|")(4): strRegNo = Split(fraInfo.Tag, "|")(5)
        intMoved = Split(fraInfo.Tag, "|")(8): intState = Split(fraInfo.Tag, "|")(9)
        lngUnit = lngPatDept
        blnCanPrint = True
        intPatientForm = Split(fraInfo.Tag, "|")(6)
    End If
    
    blnIsInsidePatient = (intPatientForm = 1) Or (intPatientForm = 2)
    
    mfrmPACSImg.zlRefresh lngAdviceID, lngSendNO, mstrPrivs, intMoved = 1, blnRefresh
    
    Select Case TabWindow(TabWindow.Selected.Index).Tag
        Case "申请费用"
            mobjExpense.zlRefresh mlngCur科室ID, lngAdviceID, lngSendNO, intMoved = 1
        Case "PACS报告"
            mfrmPacsReport.zlRefresh lngAdviceID, Me, intMoved = 1
        Case "病历报告"
            mobjReport.zlRefresh lngAdviceID, mlngCur科室ID, Not mblnIsHistory, intMoved = 1, blnCanPrint
        Case "排队叫号"
            If Not mblnIsHistory And Not mobjQueue Is Nothing Then
                mobjQueue.zlRefresh mAstr队列名称, Split(mstrCur科室, "-")(1) & vsList.TextMatrix(vsList.Row, GetCN("执行间")), lngAdviceID
            End If
        Case "住院医嘱"
            If TabWindow.Selected.Visible Then '可能由住院记录转到历次门诊记录,此时可能没有授权门诊医嘱权限
                mobjInAdvice.zlRefresh lngPatID, lngPageID, lngUnit, lngPatDept, 0, intMoved = 1, lngAdviceID, intState, lngPatDept
            Else
                For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                    If TabWindow(i).Tag = "门诊医嘱" Then
                        If strRegNo = "" Then   '自行登记的病人没有挂号单号
                            mobjOutAdvice.zlRefresh lngPatID, "", False
                        Else
                            mobjOutAdvice.zlRefresh lngPatID, strRegNo, Not mblnIsHistory And blnIsInsidePatient, intMoved = 1, lngAdviceID
                        End If
                    End If
                Next
            End If
        Case "门诊医嘱"
            If TabWindow.Selected.Visible Then '可能由门诊记录转到历次住院记录,此时可能没有授权住院医嘱权限
                If strRegNo = "" Then   '自行登记的病人没有挂号单号
                    mobjOutAdvice.zlRefresh lngPatID, "", False
                Else
                    mobjOutAdvice.zlRefresh lngPatID, strRegNo, Not mblnIsHistory And blnIsInsidePatient, intMoved = 1, lngAdviceID
                End If
            Else
                For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                    If TabWindow(i).Tag = "住院医嘱" Then
                      mobjInAdvice.zlRefresh lngPatID, lngPageID, lngUnit, lngPatDept, 0, intMoved = 1, lngAdviceID, intState, lngPatDept
                    End If
                Next
            End If
        Case "住院病历"
            If TabWindow.Selected.Visible Then '可能由住院记录转到历次门诊记录,此时可能没有授权门诊病历权限
                mobjInEPRs.zlRefresh lngPatID, lngPageID, mlngCur科室ID, Not mblnIsHistory, intMoved = 1
            Else
                For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                    If TabWindow(i).Tag = "门诊病历" Then
                       mobjOutEPRs.zlRefresh lngPatID, lngPageID, mlngCur科室ID, Not mblnIsHistory, intMoved = 1
                    End If
                Next
            End If
        Case "门诊病历"
            If TabWindow.Selected.Visible Then '可能由门诊记录转到历次住院记录,此时可能没有授权住院病历权限
                mobjOutEPRs.zlRefresh lngPatID, lngPageID, mlngCur科室ID, Not mblnIsHistory, intMoved = 1
            Else
                For i = 0 To TabWindow.ItemCount - 1 '循环到了才触发
                    If TabWindow(i).Tag = "住院病历" Then
                        mobjInEPRs.zlRefresh lngPatID, lngPageID, mlngCur科室ID, Not mblnIsHistory, intMoved = 1
                    End If
                Next
            End If
    End Select
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub subTriggleRefreshTimer(blnEnable As Boolean)
    '启动或者关闭自动刷新的Timer
    If blnEnable = False Then
        TimerRefresh.Enabled = False
    Else
        TimerRefresh.Enabled = mlngRefreshInterval > 0
    End If
End Sub

Private Function GetDeptName(lngDeptID As Long, strDeptStrings As String) As String
'通过可用的科室串，读取指定科室ID的科室名称
    Dim strDepts() As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    strDepts = Split(strDeptStrings, "|")
    For i = 0 To UBound(strDepts)
        If Split(strDepts(i), "_")(0) = lngDeptID Then
            GetDeptName = Split(strDepts(i), "_")(1)
            Exit For
        End If
    Next i
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub subChangeDept(intType As Integer)
'更换科室，更改跟科室相关的参数
'参数：intType---0从菜单触发科室改变，刷新病人列表；intType---1从病人列表触发科室改变，不用再刷新病人列表
    Dim i As Integer
    
    On Error GoTo errHandle
    Call InitMvar(1)
    mblnInitOk = False '防止在子窗体加载过程中对子窗体进行刷新
    If intType = 0 Then
        Call InitSubForm
    End If
    
    mblnInitOk = True
    If intType = 0 Then Call RefreshList
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Menu_Manage_关联病人()
'关联病人
    
    If Val(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID"))) = 0 Then Exit Sub
    
    On Error GoTo err
    Call frmReferencePatient.zlShowMe(vsList.TextMatrix(vsList.Row, GetCN("医嘱ID")), _
        vsList.TextMatrix(vsList.Row, GetCN("姓名")), Me, True)
    
    '刷新病人列表
     Call RefreshList
    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
